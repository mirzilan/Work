Attribute VB_Name = "mod_DebtSizing"
Option Explicit

' Sizes the post-COD debt quantum (GearingFactor, on VBA_Control!B3) so the
' flat-target-DSCR sculpting recursion on Debt_Sizing_DSCR fully amortizes to
' zero by tenor end. The recursion itself (Interest_t = r*OpeningBalance_t,
' Principal_t = MAX(0, CFADS_t/TargetDSCR - Interest_t)) is pure Excel formulas
' driven off the single scalar GearingFactor -- this module's only job is to
' bisect that scalar. It never re-implements the recursion in VBA, so there is
' one source of truth for the math.
'
' Default mode (SizingMethod = "GearingQuantum") bisects the debt quantum,
' holding the lender's TargetDSCR covenant fixed -- this is the standard PF
' formulation (a bank sets the DSCR covenant; the debt quantum is solved for).
' It is well-posed because increasing GearingFactor can only weakly raise the
' ending balance (more debt to repay), never lower it -- monotonic, so
' bisection converges.
'
' Alternative mode (SizingMethod = "DSCRMultiple") fixes GearingFactor at 100%
' and instead bisects TargetDSCR_Input on Assumptions.

Public Sub RunDebtSizing()
    Dim calcState As XlCalculation
    Dim screenState As Boolean
    Dim lo As Double, hi As Double, targetDSCR As Double
    Dim tol As Double, maxIter As Long
    Dim result As Double
    Dim converged As Boolean

    calcState = Application.Calculation
    screenState = Application.ScreenUpdating
    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False
    Application.StatusBar = "Sizing debt..."

    tol = Range("Tolerance").Value
    maxIter = CLng(Range("MaxIterations").Value)

    On Error GoTo CleanFail

    Select Case Range("SizingMethod").Value
        Case "DSCRMultiple"
            lo = 1#: hi = 3#
            converged = BisectTargetDSCR(lo, hi, tol, maxIter)
        Case Else ' "GearingQuantum"
            lo = 0#: hi = Range("MaxGearingPct").Value * 2# ' generous upper bound
            targetDSCR = Range("TargetDSCR_Input").Value
            converged = BisectGearing(lo, hi, targetDSCR, tol, maxIter)
    End Select

    Application.Calculate

    If converged Then
        Range("ConvergenceStatus").Value = "Converged"
    Else
        Range("ConvergenceStatus").Value = "MaxIterHit or Infeasible"
    End If

CleanFail:
    Application.Calculation = calcState
    Application.ScreenUpdating = screenState
    Application.StatusBar = False
    If Err.Number <> 0 Then
        MsgBox "RunDebtSizing failed: " & Err.Description, vbExclamation
    End If
End Sub

' Bisects GearingFactor (debt quantum multiplier on ClosingBalanceAtCOD) so the
' ending balance on Debt_Sizing_DSCR (EndingBalanceResidual) nets to zero,
' holding targetDSCR fixed. Returns True if converged within tol/maxIter.
Public Function BisectGearing(ByVal lo As Double, ByVal hi As Double, _
                               ByVal targetDSCR As Double, ByVal tol As Double, _
                               ByVal maxIter As Long) As Boolean
    Dim rLo As Double, rHi As Double, rMid As Double, mid As Double
    Dim i As Long

    Range("TargetDSCR_Input").Value = targetDSCR

    rLo = Residual(lo)
    rHi = Residual(hi)

    If Sgn(rLo) = Sgn(rHi) And rLo <> 0 And rHi <> 0 Then
        Range("IterationsUsed").Value = 0
        BisectGearing = False
        Exit Function
    End If

    For i = 1 To maxIter
        mid = (lo + hi) / 2#
        rMid = Residual(mid)
        If Abs(rMid) < tol Then
            Range("GearingFactor").Value = mid
            Range("IterationsUsed").Value = i
            BisectGearing = True
            Exit Function
        End If
        If Sgn(rMid) = Sgn(rLo) Then
            lo = mid
            rLo = rMid
        Else
            hi = mid
        End If
    Next i

    Range("GearingFactor").Value = mid
    Range("IterationsUsed").Value = maxIter
    BisectGearing = False
End Function

' Mirror-image sizing: fixes GearingFactor at 100% and bisects TargetDSCR_Input
' instead, so the debt (sized 1:1 to the construction debt balance) fully
' amortizes at whatever flat DSCR the cash flows can actually support.
Public Function BisectTargetDSCR(ByVal lo As Double, ByVal hi As Double, _
                                  ByVal tol As Double, ByVal maxIter As Long) As Boolean
    Dim rLo As Double, rHi As Double, rMid As Double, mid As Double
    Dim i As Long

    Range("GearingFactor").Value = 1#

    rLo = ResidualAtDSCR(lo)
    rHi = ResidualAtDSCR(hi)

    If Sgn(rLo) = Sgn(rHi) And rLo <> 0 And rHi <> 0 Then
        Range("IterationsUsed").Value = 0
        BisectTargetDSCR = False
        Exit Function
    End If

    For i = 1 To maxIter
        mid = (lo + hi) / 2#
        rMid = ResidualAtDSCR(mid)
        If Abs(rMid) < tol Then
            Range("TargetDSCR_Input").Value = mid
            Range("IterationsUsed").Value = i
            BisectTargetDSCR = True
            Exit Function
        End If
        ' Ending balance is monotonic DECREASING in TargetDSCR (a looser/lower
        ' target DSCR means more principal repaid each period), so the sign
        ' logic is the mirror of BisectGearing.
        If Sgn(rMid) = Sgn(rLo) Then
            lo = mid
            rLo = rMid
        Else
            hi = mid
        End If
    Next i

    Range("TargetDSCR_Input").Value = mid
    Range("IterationsUsed").Value = maxIter
    BisectTargetDSCR = False
End Function

Private Function Residual(ByVal gearing As Double) As Double
    Range("GearingFactor").Value = gearing
    Application.Calculate
    Residual = Range("EndingBalanceResidual").Value
End Function

Private Function ResidualAtDSCR(ByVal dscr As Double) As Double
    Range("TargetDSCR_Input").Value = dscr
    Application.Calculate
    ResidualAtDSCR = Range("EndingBalanceResidual").Value
End Function
