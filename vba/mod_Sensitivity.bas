Attribute VB_Name = "mod_Sensitivity"
Option Explicit

' Loops the scenario grid on Returns_Sensitivities (populated by
' scripts/build_model.py: columns B-E hold each scenario's Capex/Opex/WACC/
' Utilization deltas), re-runs RunDebtSizing (mod_DebtSizing) for each one,
' and writes LCOH / Equity IRR / Min DSCR into columns F-H.
'
' Base-case assumption cells are perturbed and restored around each scenario
' so this can be re-run repeatedly without drifting the base case.

Public Sub RunSensitivityGrid()
    Dim baseCapex As Double, baseFixedOpex As Double, baseVarOpex As Double
    Dim baseWACC As Double, baseCapFactor As Double
    Dim r As Long, firstRow As Long, lastRow As Long
    Dim dCapex As Double, dOpex As Double, dWaccBps As Double, dUtil As Double

    baseCapex = Range("TotalCapex").Value
    baseFixedOpex = Range("FixedOpexAnnual").Value
    baseVarOpex = Range("VariableOpexPerKg").Value
    baseWACC = Range("WACC").Value
    baseCapFactor = Range("CapacityFactor").Value

    firstRow = Range("SensitivityGrid_FirstRow").Row
    lastRow = Range("SensitivityGrid_LastRow").Row

    Application.ScreenUpdating = False
    Application.StatusBar = "Running sensitivity grid..."
    On Error GoTo Restore

    For r = firstRow To lastRow
        dCapex = Cells(r, 2).Value
        dOpex = Cells(r, 3).Value
        dWaccBps = Cells(r, 4).Value
        dUtil = Cells(r, 5).Value

        Range("TotalCapex").Value = baseCapex * (1 + dCapex)
        Range("FixedOpexAnnual").Value = baseFixedOpex * (1 + dOpex)
        Range("VariableOpexPerKg").Value = baseVarOpex * (1 + dOpex)
        Range("WACC").Value = baseWACC + dWaccBps / 10000#
        Range("CapacityFactor").Value = baseCapFactor * (1 + dUtil)

        RunDebtSizing ' re-sizes GearingFactor for this scenario, recalcs

        Cells(r, 6).Value = Range("LCOH_USD_per_kg").Value
        Cells(r, 7).Value = Range("EquityIRR").Value
        Cells(r, 8).Value = Application.WorksheetFunction.Min(Range("DSCR_Actual_Row"))
    Next r

Restore:
    Range("TotalCapex").Value = baseCapex
    Range("FixedOpexAnnual").Value = baseFixedOpex
    Range("VariableOpexPerKg").Value = baseVarOpex
    Range("WACC").Value = baseWACC
    Range("CapacityFactor").Value = baseCapFactor
    RunDebtSizing ' restore GearingFactor to the base-case solve
    Application.StatusBar = False
    Application.ScreenUpdating = True

    If Err.Number <> 0 Then
        MsgBox "RunSensitivityGrid failed: " & Err.Description, vbExclamation
    End If
End Sub
