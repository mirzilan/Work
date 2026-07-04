Attribute VB_Name = "mod_Utilities"
Option Explicit

' Shared helpers used by mod_DebtSizing and mod_Sensitivity.

Public Sub ForceFullRecalc()
    Application.CalculateFullRebuild
End Sub

Public Function GetNamed(ByVal rangeName As String) As Variant
    GetNamed = ThisWorkbook.Names(rangeName).RefersToRange.Value
End Function

Public Sub SetNamed(ByVal rangeName As String, ByVal value As Variant)
    ThisWorkbook.Names(rangeName).RefersToRange.Value = value
End Sub

Public Sub SetStatusBar(ByVal msg As String)
    Application.StatusBar = msg
End Sub
