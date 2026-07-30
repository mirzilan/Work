# VBA Setup — one-time manual step

`openpyxl` cannot author VBA (no Excel/COM in the build environment), so macros are
hand-authored here and pasted in once. After that, regenerating the workbook preserves
them — you only repeat this if the macro code itself changes.

## Step-by-step

1. Open the generated `.xlsx` in Excel.
2. **File → Save As → Excel Macro-Enabled Workbook (`.xlsm`)**. VBA cannot live in `.xlsx`.
3. `Alt + F11` to open the VBA editor.
4. **Insert → Module**. A blank `Module1` appears.
5. Import **both** `.bas` files from this folder — `mod_Loop1_IDC.bas` and
   `mod_Loop2_DebtSizing.bas`. Easiest route: **File → Import File…** and select each one
   (this keeps the module names). Otherwise create a module per file and paste everything
   **below** the `Attribute VB_Name` line.
6. Close the VBA editor.
7. **Developer tab → Insert → Form Control Button** (the top-left one, *not* ActiveX).
   Draw it on the `Cover` sheet near row 11.
8. Excel prompts **Assign Macro** → pick the macro → OK.
9. Right-click the button → **Edit Text** → give it the label below.
10. Repeat for each button:

| Macro | Button label |
|---|---|
| `SolveAllCurrentScenario` | **Solve All** ← the one you will use most |
| `SolveConstructionIDC` | Solve Construction IDC |
| `SolveDebtSculpting` | Solve Debt Sculpting |
| `ResetAllStagedValues` | Reset |

11. **Save** (keep `.xlsm`).

`Solve All` alternates Loop 1 and Loop 2 until both converge — the two are coupled (debt
size sets the construction facility, which changes IDC, which changes Total Project Cost,
depreciation, tax, CFADS and therefore capacity), so running them together is what lands
on a mutually consistent answer. The individual macros are there for diagnosing one loop
in isolation.

If the Developer tab is not visible: **File → Options → Customize Ribbon → tick Developer**.

## Defined names the macro depends on

These are created by the Python build — do not rename them.

| Name | Points to | Purpose |
|---|---|---|
| `StagedIDC` | `Calc_Financing_Cons!$B$6` | Value cell the macro writes; breaks the Loop 1 chain |
| `CalculatedIDC` | `Calc_Financing_Cons!$B$10` | Sum of monthly interest accrued — Loop 1 target |
| `IDCConvergenceGap` | `Calc_Financing_Cons!$B$11` | `CalculatedIDC − StagedIDC` |
| `StagedDebtSize` | `Calc_Financing_Ops!$B$7` | Value cell the macro writes; breaks the Loop 2 chain |
| `SculptedDebtCapacity` | `Calc_Financing_Ops!$B$8` | PV of the sculpting basis — Loop 2 target |
| `DebtSizeConvergenceGap` | `Calc_Financing_Ops!$B$10` | `Capacity − StagedDebtSize` |
| `ImpliedGearing` | `Calc_Financing_Ops!$B$11` | Solved debt ÷ Total Project Cost |
| `TotalCapex` | `Calc_Capex!$B$4` | Used by the reset macro |
| `Cover_CircTolerance` | `Cover!$B$4` | Loop 1 convergence tolerance ($) |
| `Cover_MaxIterations` | `Cover!$B$5` | Iteration cap before giving up |
| `Cover_DebtSizingTolerance` | `Cover!$B$6` | Loop 2 convergence tolerance ($) |
| `Cover_DrawdownMethod` | `Cover!$B$14` | Debt First / Equity First / Pari Passu |
| `Cover_DebtSizingMode` | `Cover!$B$17` | Fixed Gearing / DSCR Sculpted |

## What "solved" looks like

`Check_Control` should show every check OK and the master flag `MODEL OK`.

Expected results for the $100M dummy case in **DSCR Sculpted** mode (Target DSCR 1.30x),
converging in ~7 outer passes:

| Drawdown method | Sculpted debt | Implied gearing |
|---|---|---|
| Debt First | ~$75.08M | ~71.2% |
| Pari Passu | ~$74.97M | ~71.8% |
| Equity First | ~$74.86M | ~72.4% |

Two signatures tell you the sculpting is working: **min DSCR over the tenor equals the
target exactly** (1.3000 — that is what "locked" means), and the **closing balance at
tenor end is zero**. If DSCR drifts above target across the whole tenor, the debt is
undersized; if the balance does not reach zero, it is oversized.

In **Fixed Gearing** mode the schedule reverts to level PMT amortisation and the Loop 2
convergence check is skipped, but `SculptedDebtCapacity` still reports what the project
*could* support — useful as a sizing sense-check against the fixed assumption.

## Important — stale values

The staged value is a *number*, not a formula. If you change an assumption and do not
re-run the solve, the model will show stale figures that look perfectly valid. Task #13
adds a dirty-flag check that detects this and warns on `Cover`. Until then: **re-run the
solve after every assumption change.**
