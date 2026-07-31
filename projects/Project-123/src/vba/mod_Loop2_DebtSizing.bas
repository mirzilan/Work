Attribute VB_Name = "mod_Loop2_DebtSizing"
Option Explicit

' Project 123 -- Loop 2: DSCR-sculpted debt sizing
'
' With repayment locked to the target DSCR the balance recursion is linear:
'
'     Balance(t+1) = Balance(t) x (1 + r) - CFADS(t) / TargetDSCR
'
' Setting Balance(T) = 0 and solving gives a closed form: the debt the project can
' support is just the present value of the sculpted debt-service stream at the debt
' rate. That is what SculptedDebtCapacity computes -- so there is no root-find here.
'
' What remains circular is the tax shield: debt size -> interest -> tax -> CFADS ->
' capacity -> debt size. StagedDebtSize breaks that chain, and this macro drives it to
' its fixed point. Convergence is geometric and typically settles in ~7 iterations.
'
' The solve is split in two: SolveAllSilent does the work and returns a result, and
' SolveAllCurrentScenario wraps it with the message box. Goal seek calls the silent one
' dozens of times per run, and a dialog per call would make it unusable.

' Diagnostics from the last SolveAllSilent, for whichever caller wants to report them.
Public LastSolvePasses As Long
Public LastIDCGap As Double
Public LastDebtGap As Double


Public Sub SolveDebtSculpting()
    Dim solved As Boolean
    solved = ConvergeDebtSize(True)
End Sub


' Solves Loop 1 and Loop 2 together. The two are coupled -- debt size sets the
' construction facility, which changes IDC, which changes Total Project Cost and
' therefore depreciation, tax, CFADS and capacity. Alternating the two fixed points
' until both gaps close is the reliable way to land on a mutually consistent answer.
'
' No UI and no calculation-mode juggling: the caller owns both, because goal seek runs
' this inside its own manual-calculation block and must not have it handed back.
Public Function SolveAllSilent() As Boolean
    Dim i As Long
    Dim maxOuter As Long
    Dim tolerance As Double
    Dim sculpting As Boolean

    SolveAllSilent = False
    LastSolvePasses = 0

    tolerance = Range("Cover_CircTolerance").Value
    maxOuter = Range("Cover_MaxIterations").Value
    sculpting = (Range("Cover_DebtSizingMode").Value = "DSCR Sculpted")

    For i = 1 To maxOuter
        ConvergeIDC
        ConvergeDebtSize False

        Application.Calculate
        LastIDCGap = Abs(Range("IDCConvergenceGap").Value)
        LastDebtGap = Abs(Range("DebtSizeConvergenceGap").Value)
        LastSolvePasses = i

        ' In Fixed Gearing mode nothing drives StagedDebtSize toward capacity, so the
        ' debt gap never closes and is not a convergence condition -- only IDC is.
        If LastIDCGap <= tolerance Then
            If Not sculpting Then
                SolveAllSilent = True
                Exit Function
            ElseIf LastDebtGap <= Range("Cover_DebtSizingTolerance").Value Then
                SolveAllSilent = True
                Exit Function
            End If
        End If
    Next i
End Function


Public Sub SolveAllCurrentScenario()
    Dim prevCalcMode As XlCalculation
    Dim converged As Boolean

    prevCalcMode = Application.Calculation
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    On Error GoTo CleanUp

    converged = SolveAllSilent()

    If converged Then
        RecordSolveSnapshot
        MsgBox "Solved in " & LastSolvePasses & " outer pass(es)." & vbCrLf & vbCrLf & _
               "IDC: " & Format(Range("CalculatedIDC").Value, "#,##0") & _
               "  (gap " & Format(LastIDCGap, "#,##0.00") & ")" & vbCrLf & _
               "Debt size: " & Format(Range("StagedDebtSize").Value, "#,##0") & _
               "  (gap " & Format(LastDebtGap, "#,##0.00") & ")" & vbCrLf & _
               "Implied gearing: " & Format(Range("ImpliedGearing").Value, "0.00%"), _
               vbInformation, "Solve All -- Converged"
    Else
        MsgBox "Did NOT fully converge in " & Range("Cover_MaxIterations").Value & _
               " outer passes." & vbCrLf & _
               "IDC gap: " & Format(LastIDCGap, "#,##0.00") & vbCrLf & _
               "Debt size gap: " & Format(LastDebtGap, "#,##0.00"), _
               vbExclamation, "Solve All -- Not Converged"
    End If

CleanUp:
    Application.Calculation = prevCalcMode
    Application.ScreenUpdating = True
    If Err.Number <> 0 Then
        MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical, "Solve All -- Error"
    End If
End Sub


Private Function ConvergeDebtSize(ByVal showResult As Boolean) As Boolean
    Dim capacity As Double
    Dim staged As Double
    Dim tolerance As Double
    Dim maxIterations As Long
    Dim i As Long

    ConvergeDebtSize = False

    If Range("Cover_DebtSizingMode").Value <> "DSCR Sculpted" Then
        If showResult Then
            MsgBox "Debt Sizing Mode is 'Fixed Gearing'. Nothing to sculpt." & vbCrLf & _
                   "Switch the mode on Cover to use DSCR sculpting.", _
                   vbInformation, "Loop 2 -- Skipped"
        End If
        Exit Function
    End If

    tolerance = Range("Cover_DebtSizingTolerance").Value
    maxIterations = Range("Cover_MaxIterations").Value

    For i = 1 To maxIterations
        Application.Calculate

        capacity = Range("SculptedDebtCapacity").Value
        staged = Range("StagedDebtSize").Value

        If Abs(capacity - staged) <= tolerance Then
            Application.Calculate
            ConvergeDebtSize = True
            If showResult Then
                RecordSolveSnapshot
                MsgBox "Debt sizing converged in " & i & " iteration(s)." & vbCrLf & vbCrLf & _
                       "Sculpted debt size: " & Format(Range("StagedDebtSize").Value, "#,##0") & vbCrLf & _
                       "Implied gearing: " & Format(Range("ImpliedGearing").Value, "0.00%") & vbCrLf & _
                       "Residual gap: " & Format(Range("DebtSizeConvergenceGap").Value, "#,##0.00"), _
                       vbInformation, "Loop 2 -- Solved"
            End If
            Exit Function
        End If

        Range("StagedDebtSize").Value = capacity
    Next i

    Application.Calculate
    If showResult Then
        MsgBox "Debt sizing did NOT converge within " & maxIterations & " iterations." & vbCrLf & _
               "Residual gap: " & Format(Range("DebtSizeConvergenceGap").Value, "#,##0.00"), _
               vbExclamation, "Loop 2 -- Not Converged"
    End If
End Function


Private Sub ConvergeIDC()
    Dim newCalc As Double
    Dim tolerance As Double
    Dim maxIterations As Long
    Dim i As Long

    tolerance = Range("Cover_CircTolerance").Value
    maxIterations = Range("Cover_MaxIterations").Value

    For i = 1 To maxIterations
        Application.Calculate
        newCalc = Range("CalculatedIDC").Value
        If Abs(newCalc - Range("StagedIDC").Value) <= tolerance Then Exit Sub
        Range("StagedIDC").Value = newCalc
    Next i
End Sub


' Clears both staged cells so a solve starts from a known state rather than sitting on
' a stale value that happens to look converged.
Public Sub ResetAllStagedValues()
    Range("StagedIDC").Value = 0
    Range("StagedDebtSize").Value = Range("TotalCapex").Value * 0.7
    ' The model is no longer solved, so the freshness snapshot must not keep claiming it is.
    Range("SnapshotStored").ClearContents
    Range("LastSolvedStamp").Value = "(never)"
    Application.Calculate
    MsgBox "Staged IDC and Staged Debt Size reset, solve snapshot cleared." & vbCrLf & _
           "Run Solve All to re-solve.", vbInformation, "Reset"
End Sub
