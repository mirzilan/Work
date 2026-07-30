# VBA Setup — one-time manual step

`openpyxl` cannot author VBA (no Excel/COM in the build environment), so macros are
hand-authored here and pasted in once. After that, regenerating the workbook preserves
them — you only repeat this if the macro code itself changes.

## Step-by-step

1. Open the generated `.xlsx` in Excel.
2. **File → Save As → Excel Macro-Enabled Workbook (`.xlsm`)**. VBA cannot live in `.xlsx`.
3. `Alt + F11` to open the VBA editor.
4. **Insert → Module**. A blank `Module1` appears.
5. Open `mod_Loop1_IDC.bas` from this folder in a text editor, copy everything **below**
   the `Attribute VB_Name` line, and paste it into the module.
   (Alternatively: **File → Import File…** and select the `.bas` directly — this keeps the
   module name and skips the copy/paste.)
6. Close the VBA editor.
7. **Developer tab → Insert → Form Control Button** (the top-left one, *not* ActiveX).
   Draw it on the `Cover` sheet near row 11.
8. Excel prompts **Assign Macro** → select `SolveConstructionIDC` → OK.
9. Right-click the button → **Edit Text** → name it `Solve Construction IDC`.
10. Optionally repeat steps 7–9 for `ResetConstructionIDC`, labelled `Reset IDC`.
11. **Save** (keep `.xlsm`).

If the Developer tab is not visible: **File → Options → Customize Ribbon → tick Developer**.

## Defined names the macro depends on

These are created by the Python build — do not rename them.

| Name | Points to | Purpose |
|---|---|---|
| `StagedIDC` | `Calc_Financing_Cons!$B$6` | The value cell the macro writes; breaks the circular chain |
| `CalculatedIDC` | `Calc_Financing_Cons!$B$10` | Sum of monthly interest accrued — the target value |
| `IDCConvergenceGap` | `Calc_Financing_Cons!$B$11` | `CalculatedIDC − StagedIDC`, reported on completion |
| `Cover_CircTolerance` | `Cover!$B$4` | Convergence tolerance ($) |
| `Cover_MaxIterations` | `Cover!$B$5` | Iteration cap before giving up |
| `Cover_DrawdownMethod` | `Cover!$B$14` | Debt First / Equity First / Pari Passu |

## What "solved" looks like

`Check_Control` should show **IDC Converged (Loop 1)** = OK, and the master flag
`MODEL OK`. Expected IDC for the $100M dummy case:

| Drawdown method | IDC | Iterations |
|---|---|---|
| Debt First | ~$5.46M | ~6 |
| Pari Passu | ~$4.33M | ~2 |
| Equity First | ~$3.18M | ~5 |

Debt First carries the highest IDC because debt is drawn earliest and therefore accrues
interest for longest; Equity First the lowest for the mirror reason. If your ordering
differs, something is wrong.

## Important — stale values

The staged value is a *number*, not a formula. If you change an assumption and do not
re-run the solve, the model will show stale figures that look perfectly valid. Task #13
adds a dirty-flag check that detects this and warns on `Cover`. Until then: **re-run the
solve after every assumption change.**
