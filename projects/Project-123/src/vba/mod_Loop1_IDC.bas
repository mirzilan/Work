Attribute VB_Name = "mod_Loop1_IDC"
Option Explicit

' Project 123 -- Loop 1: Construction IDC circularity solver
'
' The circularity: IDC accrues on the drawn debt balance -> IDC is a use of funds ->
' Total Project Cost includes IDC -> Debt Facility = gearing x Total Project Cost ->
' the facility governs how much debt is drawn -> which drives IDC.
'
' StagedIDC breaks the chain: it is a plain value cell that every downstream formula
' reads, so the workbook itself contains no circular reference. This macro drives it to
' the fixed point where StagedIDC equals CalculatedIDC.
'
' Convergence is fast: Pari Passu settles in ~2 iterations (the schedule is sequential and
' does not depend on the facility), Debt First / Equity First in roughly 5-6.

Public Sub SolveConstructionIDC()
    Dim prevCalc As Double
    Dim newCalc As Double
    Dim tolerance As Double
    Dim maxIterations As Long
    Dim i As Long
    Dim prevCalcMode As XlCalculation

    tolerance = Range("Cover_CircTolerance").Value
    maxIterations = Range("Cover_MaxIterations").Value

    prevCalcMode = Application.Calculation
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    On Error GoTo CleanUp

    For i = 1 To maxIterations
        Application.Calculate

        newCalc = Range("CalculatedIDC").Value
        prevCalc = Range("StagedIDC").Value

        If Abs(newCalc - prevCalc) <= tolerance Then
            Application.Calculate
            MsgBox "Construction IDC converged in " & i & " iteration(s)." & vbCrLf & _
                   "IDC: " & Format(Range("CalculatedIDC").Value, "#,##0") & vbCrLf & _
                   "Residual gap: " & Format(Range("IDCConvergenceGap").Value, "#,##0.00"), _
                   vbInformation, "Loop 1 -- Solved"
            GoTo CleanUp
        End If

        Range("StagedIDC").Value = newCalc
    Next i

    Application.Calculate
    MsgBox "Construction IDC did NOT converge within " & maxIterations & " iterations." & vbCrLf & _
           "Residual gap: " & Format(Range("IDCConvergenceGap").Value, "#,##0.00") & vbCrLf & vbCrLf & _
           "Check the gearing and interest rate assumptions, or raise Max Iterations on Cover.", _
           vbExclamation, "Loop 1 -- Not Converged"

CleanUp:
    Application.Calculation = prevCalcMode
    Application.ScreenUpdating = True
    If Err.Number <> 0 Then
        MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical, "Loop 1 -- Error"
    End If
End Sub


' Resets the staged value so a solve can be re-run from a clean start. Useful when
' assumptions have changed a lot and you want to confirm the solve is not just sitting
' on a stale value that happens to look converged.
Public Sub ResetConstructionIDC()
    Range("StagedIDC").Value = 0
    Application.Calculate
    MsgBox "Staged IDC reset to zero. Run Solve Construction IDC to re-solve.", _
           vbInformation, "Loop 1 -- Reset"
End Sub
