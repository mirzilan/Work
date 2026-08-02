Attribute VB_Name = "mod_StressTest"
Option Explicit

' Project 123 -- Phase 2: one-at-a-time sensitivity on the active scenario.
'
' Reads the shock table on Stress_Test (StressTest_SpecAnchor), and for each row:
'   1. save the driver's original value on the active scenario's own column
'   2. apply the shock (relative: value*(1+shock); absolute: value+shock)
'   3. re-solve both circularity loops (SolveAllSilent, same routine Goal Seek uses)
'   4. record EIRR/PIRR/Gearing/Min DSCR to Stress_Test's results table
'   5. restore the original value and re-solve before moving to the next row
'
' Never cumulative -- each row isolates exactly one variable's effect, then the model is
' put back to its base-case solved state. The final restore-and-resolve at the end leaves
' the workbook exactly as it was before the run, same convention as RunAllScenarios.
'
' Column layout on Assumptions_Constant: scenario columns start at column C (column 3).
' This mirrors assumptions_constant.py's FIRST_SCENARIO_COL constant -- if that ever
' moves, this offset must move with it.
Private Const FIRST_SCENARIO_COL As Long = 3

Public Sub RunStressTest()
    Dim prevCalcMode As XlCalculation
    Dim activeScenario As Long
    Dim targetCol As Long
    Dim wsConst As Worksheet
    Dim wsStress As Worksheet
    Dim specAnchor As Range
    Dim resultsAnchor As Range
    Dim specRow As Range
    Dim resultRow As Range
    Dim i As Long
    Dim driverRow As Long
    Dim shockType As String
    Dim shockValue As Double
    Dim originalValue As Double
    Dim baseEIRR As Double
    Dim basePIRR As Double
    Dim converged As Boolean

    If MsgBox("Run one-at-a-time stress test on the active scenario (" & _
              Range("ActiveScenario").Value & ")?" & vbCrLf & vbCrLf & _
              "This temporarily overwrites and restores each shocked driver in turn -- " & _
              "it does not touch Assumptions_Constant permanently.", _
              vbQuestion + vbYesNo + vbDefaultButton2, "Run Stress Test") = vbNo Then
        Exit Sub
    End If

    Set wsConst = ThisWorkbook.Worksheets("Assumptions_Constant")
    Set wsStress = ThisWorkbook.Worksheets("Stress_Test")
    Set specAnchor = Range("StressTest_SpecAnchor")
    Set resultsAnchor = Range("StressTest_ResultsAnchor")

    activeScenario = CLng(Range("ActiveScenario").Value)
    targetCol = FIRST_SCENARIO_COL - 1 + activeScenario

    prevCalcMode = Application.Calculation
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    On Error GoTo CleanUp

    ' Base-case solve first, so every shock is measured against the same starting point.
    Range("StagedIDC").Value = 0
    Range("StagedDebtSize").Value = Range("TotalCapex").Value * 0.7
    Application.Calculate
    converged = SolveAllSilent()
    If converged Then RecordSolveSnapshot
    baseEIRR = Range("Live_EIRR").Value
    basePIRR = Range("Live_PIRR").Value

    i = 0
    Do While Len(Trim$(CStr(specAnchor.Offset(i, 0).Value))) > 0
        Set specRow = specAnchor.Offset(i, 0)
        Set resultRow = resultsAnchor.Offset(i, 0)

        driverRow = CLng(specRow.Offset(0, 1).Value)
        shockType = CStr(specRow.Offset(0, 2).Value)
        shockValue = CDbl(specRow.Offset(0, 3).Value)

        Application.StatusBar = "Stress test: " & specRow.Value & " (" & (i + 1) & ")..."

        originalValue = wsConst.Cells(driverRow, targetCol).Value
        If LCase$(shockType) = "relative" Then
            wsConst.Cells(driverRow, targetCol).Value = originalValue * (1 + shockValue)
        Else
            wsConst.Cells(driverRow, targetCol).Value = originalValue + shockValue
        End If

        Range("StagedIDC").Value = 0
        Range("StagedDebtSize").Value = Range("TotalCapex").Value * 0.7
        Application.Calculate
        converged = SolveAllSilent()
        If converged Then RecordSolveSnapshot
        Application.Calculate

        resultRow.Offset(0, 0).Value = specRow.Value
        resultRow.Offset(0, 1).Value = shockValue
        resultRow.Offset(0, 2).Value = baseEIRR
        resultRow.Offset(0, 3).Value = Range("Live_EIRR").Value
        resultRow.Offset(0, 4).Value = (Range("Live_EIRR").Value - baseEIRR) * 10000
        resultRow.Offset(0, 5).Value = basePIRR
        resultRow.Offset(0, 6).Value = Range("Live_PIRR").Value
        resultRow.Offset(0, 7).Value = (Range("Live_PIRR").Value - basePIRR) * 10000
        resultRow.Offset(0, 8).Value = Range("Live_Gearing").Value
        resultRow.Offset(0, 9).Value = Range("Live_MinDSCR").Value
        resultRow.Offset(0, 10).Value = IIf(converged, "OK", "NOT CONV")

        ' Restore before the next row -- each shock is isolated, never cumulative.
        wsConst.Cells(driverRow, targetCol).Value = originalValue

        i = i + 1
    Loop

    ' Final restore-and-resolve: leaves the workbook exactly as it was before the run.
    Range("StagedIDC").Value = 0
    Range("StagedDebtSize").Value = Range("TotalCapex").Value * 0.7
    Application.Calculate
    converged = SolveAllSilent()
    If converged Then RecordSolveSnapshot

    Range("StressTest_LastRun").Value = Format(Now, "yyyy-mm-dd hh:nn:ss")
    Range("StressTest_RunScenario").Value = Range("ActiveScenario").Value & " - " & _
        Range("ActiveScenarioName").Value

CleanUp:
    Application.Calculation = prevCalcMode
    Application.ScreenUpdating = True
    Application.StatusBar = False

    If Err.Number <> 0 Then
        MsgBox "Error " & Err.Number & ": " & Err.Description & vbCrLf & vbCrLf & _
               "If a driver was mid-shock when this happened, check Assumptions_Constant " & _
               "column " & targetCol & " against its Active column for a value left un-restored.", _
               vbCritical, "Stress Test -- Error"
    Else
        wsStress.Activate
        If converged Then
            MsgBox "Stress test complete." & vbCrLf & _
                   "Results written to Stress_Test. Scenario selector left at " & _
                   activeScenario & ".", vbInformation, "Run Stress Test"
        Else
            MsgBox "Stress test complete, but the final restore did NOT fully reconverge." & _
                   vbCrLf & "Run Solve All (Current Scenario) manually before trusting Live " & _
                   "EIRR/PIRR. Stress_Test's own rows are unaffected -- each one solved and " & _
                   "recorded before this final restore ran.", vbExclamation, "Run Stress Test"
        End If
    End If
End Sub
