Attribute VB_Name = "mod_Sensitivity"
Option Explicit

' Populates the sensitivity grid on the OUT (Dashboard) sheet. Each scenario row
' (columns B-E: Capex delta%, Power delta%, Price delta%, WACC delta bps) is
' applied to the base-case inputs, the debt is re-sized (RunDebtSizing re-solves
' the sculpting DSCR), and the resulting LCOH (nominal), equity IRR and minimum
' DSCR are written back to columns F-H. Base-case inputs are restored afterwards
' so the grid can be re-run without drift.

Public Sub RunSensitivityGrid()
    Dim baseCapex As Double, basePower As Double, basePrice As Double, baseWACC As Double
    Dim r As Long, firstRow As Long, lastRow As Long
    Dim minDSCR As Double

    baseCapex = Range("TotalCapex").Value
    basePower = Range("PowerPrice").Value
    basePrice = Range("OfftakePrice").Value
    baseWACC = Range("WACC").Value

    firstRow = Range("SensFirstRow").Row
    lastRow = Range("SensLastRow").Row

    Application.ScreenUpdating = False
    Application.StatusBar = "Running sensitivity grid..."
    On Error GoTo Restore

    For r = firstRow To lastRow
        Range("TotalCapex").Value = baseCapex * (1 + Cells(r, 2).Value)
        Range("PowerPrice").Value = basePower * (1 + Cells(r, 3).Value)
        Range("OfftakePrice").Value = basePrice * (1 + Cells(r, 4).Value)
        Range("WACC").Value = baseWACC + Cells(r, 5).Value / 10000#

        RunDebtSizing            ' re-solve sculpting DSCR for this scenario

        minDSCR = Application.WorksheetFunction.Min(Range("DSCR_Row"))
        Cells(r, 6).Value = Range("LCOH_Nominal").Value
        Cells(r, 7).Value = Range("EquityIRR").Value
        Cells(r, 8).Value = minDSCR
    Next r

Restore:
    Range("TotalCapex").Value = baseCapex
    Range("PowerPrice").Value = basePower
    Range("OfftakePrice").Value = basePrice
    Range("WACC").Value = baseWACC
    RunDebtSizing                ' restore base-case solve
    Application.StatusBar = False
    Application.ScreenUpdating = True
    If Err.Number <> 0 Then
        MsgBox "RunSensitivityGrid failed: " & Err.Description, vbExclamation
    End If
End Sub
