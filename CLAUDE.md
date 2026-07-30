# Work — Jarvis: Personal Finance/Investment Workspace

This is your persistent AI working environment for corporate finance, valuation, investment appraisal, and strategic M&A work. It combines day-to-day deal/project workspaces with a reusable professional toolkit built over time. Everything programmatic (Python/openpyxl), auditable, and reproducible.

## About You

**Background:** 9 years investment & strategy engineering.
- EY (credit risk, derivative valuation) → EY Parthenon (M&A, valuation, appraisal, board-level IC papers; infrastructure, power, utilities, transport, leisure, property sectors) → Gentari Hydrogen (Manager, Ventures & Partnerships; commercial structuring, operating-model design for energy-transition deals, $100M–$2B+ range).
- **Tools:** S&P Capital IQ, Bloomberg, Refinitiv Eikon, Power BI, SAS, VBA, Excel.
- **Education:** CFA Level III candidate.

**Working style:** Direct, dense tables/bullets over prose. Output-focused — deliver models, memos, IC papers, not commentary.

## How to Use This Repo

**Active work** — see `deals/`:
- Each deal gets a timestamped folder: `deals/2026-07-ProjectName/` containing appraisals, memos, IC papers.
- Run workflows from `workflows/` to generate consistent output.

**Reusable toolkit** — see `frameworks/`:
- `templates/` — IC paper, financial model, memo boilerplate.
- `checklists/` — screening, due-diligence, validation checklists.
- `methodologies/` — DCF frameworks, comp-analysis methods, valuation approaches.

**Market & sector knowledge** — see `market_intelligence/`:
- `research/` — sector deep-dives and investment theses.
- `benchmarks/` — competitor analysis, sector multiples, peer comparisons.
- `macro/` — economic outlook, market moves, policy changes.

**Career** — see `career/`:
- Opportunities, target roles, skill-building (CFA), professional network notes.

**Workflows** — see `workflows/`:
- `*.md` files — strategic routines Claude reads and executes (e.g., `build-ic-paper.md`, `weekly-market-review.md`).
- `*.py` files — automation scripts for data transforms, report generation, bulk uploads.

**Work log** — see `daily/`:
- Weekly reviews (e.g., `daily/2026-W31-review.md`) summarizing deal progress, market moves, learning.

**Project initiatives** — see `projects/`:
- Multi-phase builds: `projects/Project-123/` (bankable project finance engine; 4-phase roadmap).

## Philosophy & Conventions

**Auditable over clever.** Every number should trace to a driver and formula. Excel workbooks are readable by non-authors.

**Programmatic generation.** Models built with Python/openpyxl, not Excel macros. Code is the source of truth; `.xlsx` is a generated artifact.

**Institutional-grade formatting.**
- Color coding (FAST/Corality):
  - **Blue text** — hardcoded inputs/assumptions
  - **Black text** — formulas within sheet
  - **Green text** — links from another sheet
  - **Red text** — cross-workbook links (avoid)
- Consistent naming: sheet tabs, range names, headers should be self-documenting.

**Reproducibility.** Same inputs + same code = identical output. No manual post-processing.

**No comments explaining what code does** — name things well instead. Comments only for non-obvious constraints (e.g., Excel row limits, circular-reference workarounds).

**Before building new:** Check `frameworks/` and `projects/` for reusable engines/utilities first (debt sculpting, DSCR checks, 3-statement links, IRR/NPV logic). Don't duplicate.

## Confidentiality & Data

- Do not commit real client names, deal terms, financial data, or terms sheets.
- Use anonymized or illustrative figures in examples.
- `data/` folders in projects are local-only unless explicitly marked for sharing.
- Research and market intelligence can be shared; deal-specific appraisals should not be.

## Getting Started

1. **Read a workflow:** E.g., `workflows/build-ic-paper.md` — it walks you through creating an Investment Committee paper.
2. **Run a weekly review:** Check `daily/2026-W31-review.md` for format, then create your own `daily/2026-W32-review.md`.
3. **Check frameworks:** Browse `frameworks/templates/` for boilerplate, or `frameworks/checklists/` before screening a new deal.
4. **Build a deal folder:** Create `deals/2026-08-NewDeal/` and generate your first IC paper or appraisal.

---

For detailed workflows and templates, see the subfolders and markdown files referenced above.
