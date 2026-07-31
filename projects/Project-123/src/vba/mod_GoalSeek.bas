Attribute VB_Name = "mod_GoalSeek"
Option Explicit

' Project 123 -- Goal seek and batch scenario runner
'
' Excel's built-in Goal Seek is unusable in this model. It only recalculates, and this
' workbook is not a pure formula chain: StagedIDC and StagedDebtSize are numbers that a
' macro has to re-converge after any input moves. Built-in Goal Seek would read an IRR
' computed off stale staged values and confidently converge on a wrong revenue.
'
' So the search is written out here, and every trial revenue costs a full SolveAllSilent.
' That is expensive -- roughly 60 solves for a default run -- which is exactly why the
' feasibility bound check below matters.
'
' Bisection rather than secant or Newton. EIRR is monotone in revenue, but it is not
' smooth: the Max Gearing cap and the DSCR floor each put a kink in the curve, and a
' derivative-based step can jump across a kink and diverge. Bisection cannot diverge --
' it just needs the root bracketed, which the bound check establishes up front.

Private Const DRIVER_NAME As String = "ScenarioRevenueRow"


' The macro must write to the active scenario's own column on Assumptions_Constant.
' Writing to the Active column instead would overwrite an INDEX formula with a constant
' and quietly sever every scenario from the selector.
Private Function DriverCell() As Range
    Set DriverCell = Range(DRIVER_NAME).Cells(1, CLng(Range("ActiveScenario").Value))
End Function


Private Function EvaluateAt(ByVal revenue As Double, ByVal metricName As String) As Double
    DriverCell.Value = revenue
    SolveAllSilent
    Application.Calculate
    EvaluateAt = Range(metricName).Value
End Function


Public Sub GoalSeekEIRR()
    RunGoalSeek "Live_EIRR", "GoalSeek_TargetEIRR", "EIRR"
End Sub


Public Sub GoalSeekPIRR()
    RunGoalSeek "Live_PIRR", "GoalSeek_TargetPIRR", "PIRR"
End Sub


' Returns a short status string; also written to Cover so the result survives the dialog
' being dismissed. Silent when quiet is True, which is how the batch runner uses it.
Public Function SeekTarget(ByVal metricName As String, ByVal targetName As String, _
                           ByVal label As String, ByVal quiet As Boolean) As String
    Dim original As Double, lo As Double, hi As Double, mid As Double
    Dim fLo As Double, fHi As Double, fMid As Double
    Dim target As Double, tol As Double
    Dim maxIter As Long, i As Long
    Dim result As String

    original = DriverCell.Value
    target = Range(targetName).Value
    tol = Range("GoalSeek_Tolerance").Value
    maxIter = Range("GoalSeek_MaxIterations").Value
    lo = original * Range("GoalSeek_MinMultiple").Value
    hi = original * Range("GoalSeek_MaxMultiple").Value

    If lo >= hi Or original <= 0 Then
        SeekTarget = "INVALID BOUNDS"
        Range("GoalSeek_Status").Value = SeekTarget
        Exit Function
    End If

    ' Feasibility before search. Two solves settle whether the target is reachable at
    ' all, which is cheaper than discovering it after maxIter bisection steps that were
    ' never going to land -- and it tells you *which* way you are out of range.
    fLo = EvaluateAt(lo, metricName)
    fHi = EvaluateAt(hi, metricName)

    If target > fHi Then
        DriverCell.Value = original
        SolveAllSilent
        result = "INFEASIBLE - " & label & " tops out at " & Format(fHi, "0.00%") & _
                 " at " & Format(Range("GoalSeek_MaxMultiple").Value, "0.00") & "x revenue"
    ElseIf target < fLo Then
        DriverCell.Value = original
        SolveAllSilent
        result = "BELOW RANGE - " & label & " is already " & Format(fLo, "0.00%") & _
                 " at " & Format(Range("GoalSeek_MinMultiple").Value, "0.00") & "x revenue"
    Else
        result = "NOT CONVERGED in " & maxIter & " iterations"
        For i = 1 To maxIter
            mid = (lo + hi) / 2
            fMid = EvaluateAt(mid, metricName)

            If Abs(fMid - target) <= tol Then
                RecordSolveSnapshot
                result = "SOLVED - " & label & " " & Format(fMid, "0.00%") & _
                         " at revenue " & Format(mid, "#,##0") & _
                         " (" & Format(mid / original, "0.000") & "x, " & i & " iterations)"
                Exit For
            End If

            ' Monotone in revenue: below target means we need more of it.
            If fMid < target Then
                lo = mid
            Else
                hi = mid
            End If
        Next i
    End If

    Range("GoalSeek_Status").Value = result
    Application.Calculate
    SeekTarget = result
End Function


Private Sub RunGoalSeek(ByVal metricName As String, ByVal targetName As String, _
                        ByVal label As String)
    Dim prevCalcMode As XlCalculation
    Dim result As String

    If Range("Cover_DebtSizingMode").Value <> "DSCR Sculpted" Then
        If MsgBox("Debt Sizing Mode is 'Fixed Gearing', so debt will not resize as " & _
                  "revenue moves and the " & label & " response will be muted." & vbCrLf & vbCrLf & _
                  "Continue anyway?", vbQuestion + vbYesNo, "Goal Seek") = vbNo Then Exit Sub
    End If

    prevCalcMode = Application.Calculation
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.StatusBar = "Goal seeking " & label & "..."

    On Error GoTo CleanUp

    result = SeekTarget(metricName, targetName, label, False)

CleanUp:
    Application.Calculation = prevCalcMode
    Application.ScreenUpdating = True
    Application.StatusBar = False

    If Err.Number <> 0 Then
        MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical, "Goal Seek -- Error"
    Else
        MsgBox result, vbInformation, "Goal Seek -- " & label
    End If
End Sub


' Walks all 10 scenarios and writes the results to Batch_Results.
'
' This is destructive when a goal-seek mode is selected: it leaves each scenario's
' Annual Revenue at whatever value hit the target. That is the point of the run, but it
' is not undoable, hence the confirmation.
Public Sub RunAllScenarios()
    Dim prevCalcMode As XlCalculation
    Dim originalScenario As Long
    Dim mode As String
    Dim s As Long
    Dim anchor As Range
    Dim row As Range
    Dim converged As Boolean
    Dim seekResult As String
    Dim prompt As String

    mode = CStr(Range("Cover_BatchMode").Value)

    prompt = "Run all 10 scenarios in mode:" & vbCrLf & "    " & mode & vbCrLf & vbCrLf & _
             "This re-solves every scenario and overwrites Batch_Results."
    If mode <> "Solve Only" Then
        prompt = prompt & vbCrLf & vbCrLf & _
                 "It will also OVERWRITE the Annual Revenue of every scenario on " & _
                 "Assumptions_Constant with the goal-sought value. This cannot be undone."
    End If
    prompt = prompt & vbCrLf & vbCrLf & "Continue?"

    If MsgBox(prompt, vbExclamation + vbYesNo + vbDefaultButton2, "Run All 10 Scenarios") = vbNo Then
        Exit Sub
    End If

    originalScenario = CLng(Range("ActiveScenario").Value)
    Set anchor = Range("BatchResults_Anchor")

    prevCalcMode = Application.Calculation
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    On Error GoTo CleanUp

    For s = 1 To 10
        Application.StatusBar = "Batch: scenario " & s & " of 10..."
        Range("ActiveScenario").Value = s
        Application.Calculate

        ' Each scenario must start its own solve from a clean guess, not wherever the
        ' previous scenario's staged values happened to land. Two very different
        ' scenarios back-to-back (e.g. Downside straight after Upside) can otherwise
        ' start Loop 1/Loop 2 so far from the new fixed point that the outer-pass
        ' budget runs out before it reconverges -- and that failure then carries into
        ' every scenario solved after it, including the final restore-and-resolve.
        Range("StagedIDC").Value = 0
        Range("StagedDebtSize").Value = Range("TotalCapex").Value * 0.7
        Application.Calculate

        converged = SolveAllSilent()

        seekResult = "-"
        If mode = "Solve + Goal Seek EIRR" Then
            seekResult = SeekTarget("Live_EIRR", "GoalSeek_TargetEIRR", "EIRR", True)
        ElseIf mode = "Solve + Goal Seek PIRR" Then
            seekResult = SeekTarget("Live_PIRR", "GoalSeek_TargetPIRR", "PIRR", True)
        End If

        Application.Calculate
        Set row = anchor.Offset(s - 1, 0)

        row.Offset(0, 2).Value = Range("Live_TPC").Value
        row.Offset(0, 3).Value = Range("Live_DebtFacility").Value
        row.Offset(0, 4).Value = Range("Live_Gearing").Value
        row.Offset(0, 5).Value = Range("Live_MinDSCR").Value
        row.Offset(0, 6).Value = Range("Live_MinLLCR").Value
        row.Offset(0, 7).Value = Range("Live_EIRR").Value
        row.Offset(0, 8).Value = Range("Live_PIRR").Value
        row.Offset(0, 9).Value = DriverCell.Value
        row.Offset(0, 10).Value = IIf(converged, "OK", "NOT CONV")
        row.Offset(0, 11).Value = seekResult
        row.Offset(0, 12).Value = Range("CheckControl_MasterFlag").Value
    Next s

    Range("ActiveScenario").Value = originalScenario
    Range("StagedIDC").Value = 0
    Range("StagedDebtSize").Value = Range("TotalCapex").Value * 0.7
    Application.Calculate
    converged = SolveAllSilent()
    If converged Then RecordSolveSnapshot

    Range("BatchResults_LastRun").Value = Format(Now, "yyyy-mm-dd hh:nn:ss")
    Range("BatchResults_Mode").Value = mode

CleanUp:
    Application.Calculation = prevCalcMode
    Application.ScreenUpdating = True
    Application.StatusBar = False

    If Err.Number <> 0 Then
        MsgBox "Error " & Err.Number & ": " & Err.Description & vbCrLf & vbCrLf & _
               "Scenario selector restored to " & originalScenario & ".", _
               vbCritical, "Batch -- Error"
        Range("ActiveScenario").Value = originalScenario
        Application.Calculate
    ElseIf Not converged Then
        ThisWorkbook.Worksheets("Batch_Results").Activate
        MsgBox "Batch complete, but the restored scenario " & originalScenario & _
               " did NOT fully reconverge within " & Range("Cover_MaxIterations").Value & _
               " passes." & vbCrLf & vbCrLf & _
               "Run Solve All (Current Scenario) manually before trusting its numbers." & vbCrLf & _
               "Batch_Results for the other scenarios are unaffected -- each one solves and " & _
               "records its own result before moving on.", vbExclamation, "Run All 10 Scenarios"
    Else
        ThisWorkbook.Worksheets("Batch_Results").Activate
        MsgBox "Batch complete. Scenario selector restored to " & originalScenario & "." & vbCrLf & _
               "Results written to Batch_Results.", vbInformation, "Run All 10 Scenarios"
    End If
End Sub
