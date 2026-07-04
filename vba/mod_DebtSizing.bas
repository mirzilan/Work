Attribute VB_Name = "mod_DebtSizing"
Option Explicit

' Debt sizing for the green-hydrogen LCOH model (FAST/F1F9 build).
'
' Gearing is fixed at the construction-facility level (GearingFactor = 1.0, so
' the term debt equals the construction debt and there is no refinancing gap).
' The single scalar that is SOLVED is the flat sculpting DSCR (TargetDSCR_Input,
' on the Inputs sheet): the value at which the sculpted principal profile
' amortises the debt to exactly zero by the end of the debt tenor.
'
' The period-by-period sculpting recursion itself is live Excel formulas on the
' DB sheet (Interest = rate x opening; Principal = MIN(opening, MAX(0,
' CFADS/TargetDSCR - Interest))), so this module never re-implements the maths.
' It only searches for the one scalar, then leaves the workbook live.
'
' Because the principal is floored at MIN(opening, ...), the ending balance is
' zero for ANY sculpting DSCR below the exact-fit value (debt is simply repaid
' early) and strictly positive above it. So this is a one-sided BOUNDARY search,
' not a sign-change bisection: we look for the largest TargetDSCR that still
' fully amortises (residual <= tolerance). Native Goal Seek is the manual
' fallback: set EndingBalanceResidual to 0 by changing TargetDSCR_Input.

Public Sub RunDebtSizing()
    Dim calcState As XlCalculation
    Dim screenState As Boolean
    Dim lo As Double, hi As Double, mid As Double
    Dim tol As Double, resid As Double
    Dim i As Long, maxIter As Long
    Dim solved As Boolean

    calcState = Application.Calculation
    screenState = Application.ScreenUpdating
    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False
    Application.StatusBar = "Sizing debt (solving sculpting DSCR)..."
    On Error GoTo CleanFail

    tol = Range("Tolerance").Value
    maxIter = CLng(Range("MaxIterations").Value)

    ' Term debt = construction facility (no refinancing gap).
    Range("GearingFactor").Value = 1#

    lo = 1#          ' feasible lower bound (repays very early)
    hi = 6#          ' infeasible upper bound (cannot repay by tenor)

    ' Confirm the bracket: lo must be feasible, hi infeasible.
    If ResidualAt(lo) > tol Then GoTo NoBracket
    If ResidualAt(hi) <= tol Then hi = 12#   ' widen once if needed
    If ResidualAt(hi) <= tol Then GoTo NoBracket

    solved = False
    For i = 1 To maxIter
        mid = (lo + hi) / 2#
        resid = ResidualAt(mid)
        If resid <= tol Then
            lo = mid            ' feasible -> push the boundary up
        Else
            hi = mid            ' infeasible -> pull down
        End If
        If (hi - lo) < 0.000001 Then
            solved = True
            Exit For
        End If
    Next i

    Range("TargetDSCR_Input").Value = lo
    Application.Calculate
    Range("IterationsUsed").Value = i
    Range("ConvergenceStatus").Value = IIf(solved, "Converged", "MaxIterHit")
    GoTo Done

NoBracket:
    Range("ConvergenceStatus").Value = "Infeasible (no bracket)"
    Range("IterationsUsed").Value = 0

Done:
CleanFail:
    Application.Calculation = calcState
    Application.ScreenUpdating = screenState
    Application.StatusBar = False
    If Err.Number <> 0 Then
        MsgBox "RunDebtSizing failed: " & Err.Description, vbExclamation
    End If
End Sub

' Writes a trial sculpting DSCR, recalculates, and returns the ending debt
' balance at tenor end (EndingBalanceResidual). Zero (within tolerance) means
' the debt fully amortises by the tenor.
Private Function ResidualAt(ByVal dscr As Double) As Double
    Range("TargetDSCR_Input").Value = dscr
    Application.Calculate
    ResidualAt = Range("EndingBalanceResidual").Value
End Function

' Alternative sizing mode: hold the sculpting DSCR fixed and instead solve the
' gearing (term/construction debt ratio) that just amortises by tenor end.
' Same boundary logic (ending balance is monotonic increasing in gearing).
Public Sub RunDebtSizing_ByGearing()
    Dim lo As Double, hi As Double, mid As Double, tol As Double
    Dim i As Long, maxIter As Long
    tol = Range("Tolerance").Value
    maxIter = CLng(Range("MaxIterations").Value)
    Application.Calculation = xlCalculationManual
    lo = 0.1: hi = 1.5
    For i = 1 To maxIter
        mid = (lo + hi) / 2#
        Range("GearingFactor").Value = mid
        Application.Calculate
        If Range("EndingBalanceResidual").Value <= tol Then
            lo = mid
        Else
            hi = mid
        End If
        If (hi - lo) < 0.000001 Then Exit For
    Next i
    Range("GearingFactor").Value = lo
    Application.Calculate
    Application.Calculation = xlCalculationAutomatic
    Range("ConvergenceStatus").Value = "Converged (gearing mode)"
    Range("IterationsUsed").Value = i
End Sub
