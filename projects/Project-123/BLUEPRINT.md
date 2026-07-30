# Project 123 — Blueprint (Phase 1)

Reference document for the bankable project finance model engine. Check every build task against this before starting it — if a task doesn't map back to something here, it's scope creep; if something here has no task, it's a gap.

---

## Vision (4-phase roadmap)

| Phase | Scope |
|---|---|
| **1 (current)** | Core engine skeleton — dual timeline, 3-statements, capex/revenue/opex drivers, debt sculpting, checks |
| 2 | Body work & formatting — FAST/Corality styling, stress testing, scenario highlight (conditional formatting) |
| 3 | Sector plug-ins — H2/NH3, CCGT, Solar+BESS, Commodity Trading (swap `Calc_Revenue_Opex` internals only) |
| 4 | Front-end UI |

---

## Timeline architecture

- **Construction:** Monthly, 1–5+ years (parameterized)
- **Operations:** Quarterly, 20–30 years (parameterized)
- **Output/FS:** Annualized (rolled up from quarterly)
- **Rule:** time always runs **across columns**, consistently, on every sheet. No exceptions.
- Construction and operations resolution differences roll up cleanly at the year boundary (verified — this was a real bug once, fixed).

---

## Full sheet list

| Color | Sheet | Resolution | Status | Purpose |
|---|---|---|---|---|
| 🟦 Input | `Cover` | — | ✅ built (Stage 1b will add VBA buttons) | Circ tolerances, scenario selector (1c), master check light, drawdown/DSRA method selectors |
| 🟦 Input | `Assumptions_Model` | — | ✅ built (single-scenario) | High-level periods, FX, escalation-adjacent settings, all centralized inputs |
| 🟦 Input | `Assumptions_Constant` | — | ⏳ Stage 1c | Scenarios **across columns**, 10 placeholders, Active column via `INDEX`/`MATCH` |
| 🟦 Input | `Assumptions_Periodic_Capex` | Monthly | ⏳ Stage 1c | Scenario **row-blocks**, capex phasing; escalation library embedded at bottom (monthly index) |
| 🟦 Input | `Assumptions_Periodic_Ops` | Quarterly | ⏳ Stage 1c | Scenario row-blocks, volume/price/opex/maintenance-capex drivers; escalation library embedded at bottom (quarterly index) |
| 🟨 Calc | `Calc_Capex` | Monthly | ✅ built (no IDC yet) | Draws by category, cumulative spend, funding split; IDC line pending 1b |
| 🟨 Calc | `Calc_Financing_Cons` | Monthly | ✅ built (fixed ratio) | Debt/equity draws, drawdown method (1b), IDC + Loop 1 (1b) |
| 🟨 Calc | `Calc_Revenue_Opex` | Quarterly | ✅ built (flat dummy) | Revenue, opex, other income; escalation-driven in 1c |
| 🟨 Calc | `Calc_Tax` | Quarterly | ✅ built (single vintage, pre-interest) | Depreciation, loss carryforward, tax; interest deduction (1b), maintenance capex vintage (1c) |
| 🟨 Calc | `Calc_CFADS` | Quarterly | ✅ built (no reserves) | Cash waterfall to FCFE; DSRA/MRA (1c) |
| 🟨 Calc | `Calc_Financing_Ops` | Quarterly | ✅ built (level amortization) | DSCR-**locked** sculpted repayment (1b, replaces level amort), LLCR/PLCR (1c) |
| 🟩 Output | `FS_Quarterly` | Quarterly | ✅ built | 3-statements, FCFF/FCFE built once, PIRR/EIRR via `XIRR` |
| 🟩 Output | `FS_Annual` | Annual | ✅ built (ops-only) | Rolled up from Quarterly; construction-period section pending 1b |
| 🟩 Output | `Valuation_SellDown` | Annual | ⏳ Stage 1d | Standalone, read-only downstream of `FS_Annual`. Per exit-year: implied sale price (`XNPV`), seller's realized EIRR (`XIRR`), buyer's implied PIRR |
| 🟩 Output | `Dashboard` | — | ⏳ Stage 1d | Charts + Sources & Uses table (formula-linked, not chart-derived) |
| 🟥 Check | `Check_Control` | — | ✅ built (8 checks) | Master aggregator, direct-cell-ref pulls (no `INDIRECT`), `MODEL OK`/`ERRORS FOUND` |

---

## Core mechanics

### Circularity — two distinct loops, never conflated

| Loop | What's circular | Breaker |
|---|---|---|
| **Loop 1 — Construction** | IDC ↔ Debt Balance ↔ Total Project Cost | VBA copy-paste convergence (single staged cell) |
| **Loop 2 — Operations** | Debt sculpting ↔ CFADS ↔ Debt Balance ↔ Interest (tax-shield) | VBA copy-paste convergence, nested inside a single-variable root-find |

**Debt sculpting — the corrected method (not bisection-on-two-constraints):**
```
Debt Service_t = CFADS_t / Target DSCR      (locks DSCR exactly, no search needed)
Interest_t     = Opening Balance_t × rate
Principal_t    = Debt Service_t − Interest_t
Closing_t      = Opening_t − Principal_t
```
Only **one** unknown needs solving: starting debt size `D` such that closing balance = 0 exactly at tenor end. Single bisection/secant search, not a joint two-constraint search.

### DSRA / MRA
- Waterfall: `Revenue − Opex − Tax = CFADS − Debt Service − DSRA funding/(release) − MRA funding/(release) = FCFE`
- **DSRA funding method selector:** Cash-funded (trapped from CFADS) OR **LC-backed** (no cash trapped; recurring LC fee charged instead)
- **MRA:** funds routine + lumpy maintenance capex draws (own depreciation vintage in `Calc_Tax`)

### Coverage ratios
- **DSCR:** period-by-period, CFADS/Debt Service
- **LLCR:** forward PV of CFADS to **loan maturity** ÷ outstanding debt (+ DSRA balance)
- **PLCR:** forward PV of CFADS to **end of project life** ÷ outstanding debt (always ≥ LLCR — shows post-debt "tail" cushion)

### Scenarios
- **One selector** (`Cover`), drives `Assumptions_Constant`'s Active column and both Periodic sheets' Active rows via `INDEX`
- `Assumptions_Constant`: scenarios **across columns** (10 placeholders)
- `Assumptions_Periodic_*`: scenarios as **row-blocks**, time still across columns within each block

### Escalations
- Embedded at the bottom of each Periodic sheet (not a separate sheet), resolution matches that sheet (monthly in Capex, quarterly in Ops)
- Annual % input → within-year compounded index; drivers pick which factor applies via a selector cell

### Returns discipline
- `XIRR`/`XNPV` throughout — never `IRR`/`NPV`. Periods aren't uniform (monthly → quarterly → annual).
- PIRR = unlevered (FCFF-based), stays a single whole-of-project number — **never** recalculated per exit year (that was a caught conceptual error)
- EIRR = levered (FCFE-based), the one that varies by holding period / exit point

### Sell-down (`Valuation_SellDown`)
- Annual **scan across exit years** (not a single user-picked date), fully standalone/read-only downstream of `FS_Annual`
- Per exit year: implied sale price, seller's realized EIRR, buyer's implied PIRR (gross-up price + pro-rata debt assumed, discount FCFF pro-rata — shows the underlying project return independent of buyer's own financing choice)

### Checks architecture
- Every calc/output sheet: local Checks block, consistent position, plain pass/fail or count-based
- `Check_Control`: pulls every check via **direct cell reference**, master flag via `COUNTIF`
- **Dirty-flag (Stage 1b, task #13):** checksum of key inputs vs. last-solve snapshot — visible warning if assumptions changed since the last VBA solve, so stale copy-pasted values are never silently trusted

### VBA — build process constraint
- `openpyxl` **cannot author VBA** — no COM/Excel in this environment
- Workflow: I hand you plain-text module code + a defined-names spec + button instructions; you paste into the VBA editor once (`.xlsx` → Save As `.xlsm` → `Alt+F11` → Insert Module → paste → Developer tab → Insert Form Control Button → Assign Macro)
- One-time setup per macro group; future regenerations use `keep_vba=True` to preserve what you've pasted

### Control panel (top of `Assumptions_Constant`, Stage 1c)
- **Per-scenario (fast):** Solve Construction IDC | Solve Debt Sculpting | Solve All (Current Scenario) | Goal Seek → EIRR | Goal Seek → PIRR
- **Batch (explicit, isolated, confirmation prompt):** Run All 10 Scenarios
- **Live display (no button needed):** Current EIRR/PIRR vs. Target, Status
- **Multi-scenario goal-seek stopping logic:** test feasibility at the revenue bound *first* (EIRR should move monotonically) — if even the max plausible revenue can't hit target, flag infeasible immediately rather than burning iterations discovering it

---

## Stage sequencing (with all identified gaps folded in)

| Stage | Tasks | Status |
|---|---|---|
| **1a — Plumbing proof** | Single scenario, flat dummy revenue, fixed-ratio debt, zero circularity | ✅ **complete** (tasks #1–10) |
| **Interim — Input centralization** | `Cover` + `Assumptions_Model`, rewire all `Calc_*` hardcodes to links | ✅ **complete** (task #18) |
| **1b — Circularity** | Loop 1 + drawdown method selector + construction-period `FS_Annual` (#11) → Loop 2 + tax shield/interest deduction (#12) → dirty-flag check (#13) | ⏳ **next**, #11 in progress |
| **1c — Scale out** | 10 scenarios + escalation library (#14) → DSRA/MRA + LC option + LLCR/PLCR + multi-vintage maintenance capex (#15) → control panel + goal-seek (#16) | pending |
| **1d — Sell-down** | `Valuation_SellDown` + `Dashboard` (#17) | pending |

**Dummy test case (Stage 1a validation):** $100M project, 24mo construction, 20yr ops, $15M/yr flat revenue, 70/30 debt/equity, 6% interest, 25% tax. Verified: BS balances all 80 quarters, model winds to exactly $0 at end of life, EIRR (7.16%) > PIRR (6.09%) correctly reflects leverage, EIRR moves monotonically with revenue.

---

## What NOT to do (guardrails, hard-won from earlier missteps)

- Don't put periodic scenario inputs with time-down-rows — breaks the universal time-across-columns rule
- Don't rebuild FCFF/FCFE independently on `FS_Quarterly` and `FS_Annual` — build once, roll up
- Don't recalculate PIRR per exit year in `Valuation_SellDown` — it's a whole-of-project unlevered number, doesn't vary by holding period
- Don't bisect on two joint constraints (DSCR floor + repayment) for debt sculpting — lock DSCR via formula, bisect only on debt size
- Don't build scenario conditional-formatting highlighting now — explicitly Phase 2
- Don't skip the feasibility pre-check in multi-scenario goal-seek — burns iterations discovering what one bound-check would show instantly
- Don't trust VBA-solved values without the dirty-flag check — stale copy-pasted numbers look identical to fresh ones
