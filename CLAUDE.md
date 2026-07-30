# Work

This repository is a working environment for corporate finance and strategy engineering: financial models, valuation and deal analysis, strategic planning artifacts, and the tools that generate them. Output is typically Excel (`.xlsx`), built programmatically rather than hand-assembled, so it stays auditable, versionable, and reproducible.

## Philosophy

- **Auditable over clever.** Every output number should be traceable to a driver and a formula, not a hardcoded value. If a human reviewer can't follow the logic in the workbook itself, it's not done.
- **Programmatic generation.** Models are built by code (primarily Python + `openpyxl`), not manually in Excel. The code is the source of truth; the `.xlsx` is a generated artifact.
- **Institutional-grade formatting.** Financial workbooks follow standard buy-side/banking conventions: consistent color coding for inputs vs. formulas vs. links, clear section headers, and a layout a deal team could pick up cold.
- **Reproducibility.** Given the same inputs and code, output must be identical. No manual post-processing of generated files.

## Repository structure

- `src/` — Python source for model engines, generators, and shared finance/strategy utilities.
- `data/` — input data and assumptions (never commit real client/deal data — see Confidentiality below).
- `output/` — generated workbooks and reports. Never committed (`*.xlsx` is gitignored); regenerate from `src/` instead.
- Each major initiative (e.g. a specific model build) gets its own subdirectory or package under `src/`, not a flat dump of scripts.

## Conventions

- **Language:** Python, using `openpyxl` for Excel generation unless a task specifically calls for something else.
- **No hardcoded assumptions in code.** Drivers (rates, dates, growth assumptions, capex/opex inputs) belong in a clearly separated inputs layer, not buried in formula-construction logic.
- **Formula-first.** Where Excel can compute a value via formula, prefer writing the formula over precomputing the value in Python and writing a static number — the whole point is a workbook that recalculates correctly if a driver changes.
- **Color coding (standard FAST/Corality-style convention), once formatting work begins:**
  - Blue text — hardcoded inputs/assumptions
  - Black text — formulas within the same sheet
  - Green text — links pulled from another sheet
  - Red text — links pulled from another workbook (avoid where possible)
- **Naming:** sheet names, tab colors, and range names should be descriptive and consistent across models — a reviewer should never have to guess what `Sheet3` is.
- **No comments explaining what code does** — name things well instead. Comments are reserved for non-obvious constraints (e.g. "Excel row limit," "circularity workaround for interest-during-construction").

## Working practices

- Before adding a new model or module, check for existing engines/utilities to extend rather than starting fresh — corporate finance logic (debt sculpting, DSCR checks, 3-statement links, discounting conventions) is meant to be reused across engagements.
- Sanity-check generated workbooks open cleanly and formulas evaluate before calling a task done — a script that "ran successfully" but produced a broken workbook is not a success.
- Large or multi-phase builds should be planned and tracked explicitly (roadmap phases, milestones) rather than built ad hoc.

## Confidentiality

This repo may be used across multiple engagements/clients. Do not commit real client names, deal terms, or non-public financial data into version control. Use anonymized or illustrative figures in anything checked in; treat `data/` as local-only unless a file is explicitly meant to be shared.
