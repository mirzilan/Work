# Green Hydrogen LCOH Project-Finance Model

A quarterly-period infrastructure financial model for a green hydrogen production
project, headlined by an **unlevered LCOH** ($/kg H2), with two circularity
problems handled deliberately:

- **IDC (Interest During Construction)** is solved in **closed form** as a plain
  Excel formula (no iterative calculation, no VBA) — see the derivation on the
  `Capex_Construction` tab.
- **Post-COD debt is sculpted to a flat target DSCR.** The sculpting recursion
  itself is live Excel formulas driven by one scalar cell (`GearingFactor`).
  Solving for the debt quantum that fully amortizes by tenor end *is* a genuine
  root-finding problem (CFADS is irregular period to period), so that part is
  done in VBA (bisection).

## Files

- `HydrogenLCOH_Model.xlsx` — the workbook (all sheets, formulas, named ranges).
- `scripts/build_model.py` — regenerates the workbook from scratch (`python3 scripts/build_model.py`, requires `openpyxl`).
- `vba/mod_DebtSizing.bas` — the bisection macro that sizes/sculpts the debt.
- `vba/mod_Utilities.bas` — shared VBA helpers.
- `vba/mod_Sensitivity.bas` — runs the scenario grid on `Returns_Sensitivities`.

## Enabling the macros

This was built headlessly (no Excel available in the build environment), so the
`.xlsx` ships without a compiled VBA project — Excel can't compile VBA from
plain text non-interactively. To wire up the macros:

1. Open `HydrogenLCOH_Model.xlsx` in Excel.
2. Alt+F11 to open the VBA editor, then **File > Import File** and import all
   three `.bas` files from `vba/`.
3. On the `Dashboard` sheet, add two Form Control buttons and assign
   `RunDebtSizing` and `RunSensitivityGrid` to them respectively.
4. **File > Save As** and choose **Excel Macro-Enabled Workbook (.xlsm)**.

**Fallback without macros:** the sculpting recursion is pure formulas keyed off
`GearingFactor`, so you can size debt manually with native Goal Seek:
*Data > What-If Analysis > Goal Seek* — Set cell `EndingBalanceResidual`,
To value `0`, By changing cell `GearingFactor`.

## Model conventions

Quarterly periods across columns (D onward), one line item per row. 2-year
construction + 20-year operations by default (edit `Assumptions`). Iterative
calculation is deliberately left **off** — nothing in this design should ever
need it; a circular-reference warning means a real bug, not a setting to flip.

See `Cover` sheet inside the workbook for the full conventions legend, and
`/root/.claude/plans/plan-first-can-u-playful-beaver.md` (session-local) for
the original design writeup.
