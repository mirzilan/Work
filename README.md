# Green Hydrogen LCOH — Project-Finance Model

A bankable, semi-annual project-finance model for a green hydrogen production
project, headlined by **LCOH** (levelised cost of hydrogen, nominal & real),
built to **FAST / F1F9** conventions. Generated deterministically from source by
`scripts/build_model.py` (openpyxl).

## Modelling standard (FAST / F1F9)

- **One row = one calculation = a single formula copied unbroken across the row.**
  Timing is driven by **flag & counter rows** on the Time (`TM`) sheet, never by
  branching a formula per column.
- **Corkscrews for every balance** (`opening = prior closing`), with a **period-0
  column** so even the first period's formula is identical.
- **No live circular references** (iterative calc stays OFF as a tripwire):
  closed-form IDC, all interest on opening balances, **tax paid one period in
  arrears**, DSRA interest on opening balance. The only "solve" is a single scalar
  (the sculpting DSCR), done outside recalc by the macro / Goal Seek.
- Inputs (blue) live only on the `IN` sheet; black = formula; green = cross-sheet
  link. Units column throughout, frozen panes, sheet mnemonics.

## Sheet map

| | | |
|---|---|---|
| `CV` Cover | `IN` Inputs | `TM` Time (flags/indices/DFs) |
| `CX` Construction & capex (closed-form IDC) | `PR` Production & revenue | `OP` Operating costs |
| `WC` Working capital | `TX` Tax (dep., losses, arrears) | `CF` CFADS |
| `DB` Debt (sculpted) | `DS` DSRA | `WF` Cash waterfall |
| `CR` Cover ratios (DSCR/LLCR/PLCR) | `FS` Financial statements | `RT` Returns |
| `LC` LCOH (nominal/real + cost stack) | `CK` Checks | `OUT` Dashboard |

## Two circularity problems, handled the right way

1. **IDC (interest during construction)** — solved in **closed form** on `CX`
   (`I = r·(open + draw/2) / (1 − r·(1−f)/2)`, the linear circularity isolated
   algebraically). No iteration, no VBA. `f` splits interest between cash-funded
   (equity) and capitalised (rolled into the balance).
2. **Post-COD debt sizing** — the debt is **sculpted to a flat DSCR** every period
   (live formulas on `DB`). Gearing is fixed at the construction-facility level
   (no refinancing gap); the flat **sculpting DSCR is solved** so the debt
   amortises to zero exactly at tenor end. That single-scalar root-find is the
   genuine numerical step — done by the `RunDebtSizing` macro, or by native Goal
   Seek. The solved value ships as the default so the workbook opens converged.

## Solving / re-solving the debt

The base case ships pre-solved. After changing any input, re-solve the sculpting
DSCR one of two ways:

- **Macro:** import the `vba/*.bas` files (below) and run `RunDebtSizing`.
- **Goal Seek (no macros):** Data ▸ What-If Analysis ▸ Goal Seek —
  set cell **`EndingBalanceResidual`** to value **0** by changing cell
  **`TargetDSCR_Input`** (labelled *Sculpting DSCR (solved)* on `IN`).

## Files

- `HydrogenLCOH_Model.xlsx` — the model (all sheets, formulas, named ranges).
- `scripts/build_model.py` — regenerates the workbook (`python3 scripts/build_model.py`).
- `vba/mod_DebtSizing.bas` — boundary bisection that solves the sculpting DSCR.
- `vba/mod_Sensitivity.bas` — populates the dashboard sensitivity grid.
- `vba/mod_Utilities.bas` — shared helpers.

## Enabling the macros

The workbook is built headlessly (no Excel in the build environment), so it ships
as `.xlsx` with the VBA as importable text. To wire up the macros:

1. Open `HydrogenLCOH_Model.xlsx` in Excel.
2. Alt+F11 ▸ **File ▸ Import File** ▸ import the three `vba/*.bas` files.
3. On `OUT`, add Form Control buttons and assign `RunDebtSizing` and
   `RunSensitivityGrid`.
4. **Save As ▸ Excel Macro-Enabled Workbook (.xlsm)**.

Without the macros the model is fully functional — use the Goal Seek fallback
above to re-size the debt.

## Base-case headlines (recalculated)

LCOH ≈ **$7.23/kg** nominal (**$6.86/kg** real); project IRR ≈ **13%**, equity
IRR ≈ **23%**; DSCR sculpted flat at ≈ **2.20×** (lender covenant 1.30×);
LLCR ≈ **2.20×**, PLCR ≈ **2.91×**; gearing ≈ 73% of capex; all integrity checks
on `CK` pass (balance sheet balances every period, debt fully amortised, reserves
and cash non-negative). Figures move with the inputs — re-solve after changes.
