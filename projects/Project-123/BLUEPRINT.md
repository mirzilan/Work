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
| 🟦 Input | `Cover` | — | ✅ built | Tolerances, drawdown + debt-sizing selectors, solve-freshness block, master check light; scenario selector (1c) |
| 🟦 Input | `Assumptions_Model` | — | ✅ built (resolved layer) | Live values for the active scenario; every Calc sheet reads here, scenario machinery sits behind it |
| 🟦 Input | `Assumptions_Constant` | — | ✅ built | 16 drivers x 10 scenarios across columns, Active column via `INDEX`; escalation rates live here so they vary by scenario |
| 🟦 Input | `Assumptions_Periodic_Capex` | Monthly | ✅ built | Capex phasing as 10 scenario row-blocks + Active row; per-scenario sum checks. No escalation library — see note below |
| 🟦 Input | `Assumptions_Periodic_Ops` | Quarterly | ✅ built | Revenue + opex volume indices as scenario row-blocks; lumpy maintenance capex row-block; escalation library (4 factors, annual % -> quarterly compounded index) with per-driver factor selectors |
| 🟨 Calc | `Calc_Capex` | Monthly | ✅ built (uses-only, IDC linked) | Draws by category, cumulative spend, IDC + TPC; funding rows link from Financing_Cons |
| 🟨 Calc | `Calc_Financing_Cons` | Monthly | ✅ built (Loop 1 live) | Owns all funding: drawdown method, debt/equity draws, IDC solve |
| 🟨 Calc | `Calc_Revenue_Opex` | Quarterly | ✅ built (index + escalation driven) | Revenue and opex = base x volume index x escalation index; opex escalates off its own base, not off escalated revenue. Also owns maintenance capex (routine % of revenue + lumpy) — it is revenue-driven and `Calc_Tax` must read it |
| 🟨 Calc | `Calc_Tax` | Quarterly | ✅ built (multi-vintage) | Base depreciation on TPC + rolling-window maintenance vintages, tax, interest deduction |
| 🟨 Calc | `Calc_CFADS` | Quarterly | ✅ built (MRA live) | Cash waterfall to FCFE and FCFF; owns the MRA. The DSRA lives on `Calc_Financing_Ops` — see below |
| 🟨 Calc | `Calc_Financing_Ops` | Quarterly | ✅ built (Loop 2, DSRA, LLCR/PLCR live) | DSCR-locked sculpting + closed-form debt sizing; DSRA (cash-funded or LC-backed); LLCR/PLCR |
| 🟩 Output | `FS_Quarterly` | Quarterly | ✅ built | 3-statements, FCFF/FCFE built once, PIRR/EIRR via `XIRR` |
| 🟩 Output | `FS_Annual` | Annual | ✅ built (ops + construction) | Rolled up from Quarterly; separate construction-period annual block |
| 🟩 Output | `Valuation_SellDown` | Annual | ✅ built | Standalone, read-only downstream of `FS_Annual`. Per exit-year: implied sale price (SUMPRODUCT-based XNPV-equivalent), seller's realized EIRR (`XIRR`), buyer's implied PIRR (`XIRR`) |
| 🟩 Output | `Dashboard` | — | ✅ built | Sheet 1 (opens here). Traffic-light Model Status + Solve Freshness, headline EIRR/PIRR vs Target, link-only scenario comparison table from `Batch_Results` |
| 🟩 Output | `Batch_Results` | — | ✅ built | One row per scenario, written as **values** by the batch runner. Formulas would all resolve to whichever scenario was selected when the run ended |
| 🟥 Check | `Check_Control` | — | ✅ built (31 checks) | Master aggregator, direct-cell-ref pulls (no `INDIRECT`), `MODEL OK`/`ERRORS FOUND`. Every cell ref is derived from the source module's row constants, never a literal |

---

## Core mechanics

### Circularity — two distinct loops, never conflated

| Loop | What's circular | Breaker |
|---|---|---|
| **Loop 1 — Construction** | IDC ↔ Debt Balance ↔ Total Project Cost | VBA copy-paste convergence (single staged cell) |
| **Loop 2 — Operations** | Debt size ↔ Interest ↔ Tax ↔ CFADS ↔ Capacity (tax shield) | VBA copy-paste convergence (single staged cell) — no root-find, see below |

**Debt sculpting — closed form, no search at all:**
```
Debt Service_t = CFADS_t / Target DSCR      (locks DSCR exactly)
Interest_t     = Opening Balance_t × rate
Principal_t    = Debt Service_t − Interest_t
Closing_t      = Opening_t − Principal_t
```
Because DSCR is locked, the balance recursion is linear — `Balance(t+1) = Balance(t)×(1+r) − CFADS(t)/DSCR`. Setting `Balance(T) = 0` and solving gives a **closed form**: the supportable debt is just the PV of the sculpted service stream at the debt rate.

$$D = \sum_{t=1}^{T} \frac{CFADS_t / DSCR}{(1+r)^t} = \text{NPV}(r, \text{service})$$

So there is **no bisection** (an earlier draft of this blueprint specified one — superseded). The only iteration left is the tax-shield fixed point, which the staged cell breaks. Converges in ~7 passes.

⚠️ **The PV must discount the *uncapped* basis.** The displayed service row is floored at interest and capped at the outstanding balance; discounting *that* makes the fixed point degenerate — once the balance hits zero the service drops to zero, so the PV simply reproduces whatever balance it was given and any starting debt size looks converged. Keep a separate uncapped "Sculpting Basis" row for the PV.

### DSRA / MRA
- Waterfall: `CFADS − Debt Service − DSRA cash cost − MRA funding/(release) − Maintenance Capex = FCFE`
- CFADS stays **pre-maintenance-capex** — it is the stream the sculpting basis divides by Target DSCR. Maintenance spend is funded through the MRA line below debt service. `FCFF = CFADS − Maintenance Capex`, built once on `Calc_CFADS`
- Maintenance capex and the MRA are separate lines, never netted: the reserve releases cash in the quarter the spend lands, and netting them would hide both
- **Where each reserve lives.** The DSRA sits on `Calc_Financing_Ops` (its requirement is forward debt service, which is owned there); the MRA sits on `Calc_CFADS`. Putting the DSRA on the waterfall sheet was tried and reverted — it forces the LC fee to be sourced from `Calc_CFADS`, which puts Tax and CFADS in a cycle
- Both reserve targets are **forward-looking windows** (DSRA: N months of forward debt service; MRA: N quarters of forward maintenance capex), so each winds itself to zero as its stream runs out — no explicit release event needed, and fundings net to zero over the life
- Window widths are **structural** build-time inputs (`ReservesInputs`), not scenario drivers: they set how many columns a range spans and the ranges are written out at build time. Scenario-varying widths would need `OFFSET`. The *rates* that can vary by scenario (routine maintenance %, LC fee %) live on `Assumptions_Constant`
- **DSRA funding method selector** (`Cover`): Cash-funded (trapped from CFADS; reserve shows as restricted cash on the BS) OR **LC-backed** (no cash trapped, no asset; recurring LC fee charged on the requirement instead)
- ⚠️ **The LC fee is non-deductible**, deliberately. Deducting it runs `LC fee → tax → CFADS → debt service → DSRA requirement → LC fee`, and that *is* a real cycle — the balance recursion ties quarter `t+1` back to `t`, so the forward-looking requirement does not save you. Neither staged cell breaks it. The fee therefore sits in the P&L **below tax** and reaches cash through net income, not through the CFF reserve line (taking it in both would double-count it). Making it deductible needs a third staged cell and a VBA pass — deferred, and understating the LC option's benefit is the safe direction to be wrong in
- **Multi-vintage maintenance depreciation:** every quarter's spend opens its own straight-line vintage. Because all vintages share one life, the charge in quarter `t` is just the spend still inside the window divided by that life — a single rolling-window `SUM`, no NxN vintage matrix. Vintages opened near the end of life are truncated by the horizon, so the charge is never overstated
- Because late vintages outlive the model horizon, **PP&E no longer winds to zero** at end of life. Cash and debt still do. That residual book value is real, not an error — terminal value is not modelled

### Coverage ratios
- **DSCR:** period-by-period, CFADS/Debt Service
- **LLCR:** forward PV of CFADS to **loan maturity** ÷ outstanding debt (+ DSRA balance)
- **PLCR:** forward PV of CFADS to **end of project life** ÷ outstanding debt (always ≥ LLCR — shows post-debt "tail" cushion)

### Scenarios
- **One selector** (`Cover`), drives `Assumptions_Constant`'s Active column and both Periodic sheets' Active rows via `INDEX`
- `Assumptions_Constant`: scenarios **across columns** (10 placeholders)
- `Assumptions_Periodic_*`: scenarios as **row-blocks**, time still across columns within each block

### Escalations
- Embedded at the bottom of `Assumptions_Periodic_Ops`, at that sheet's quarterly resolution
- **Rates** live on `Assumptions_Constant` (so they vary by scenario); only the **index derivation** lives on the periodic sheet
- Annual % → within-year compounded index (`Index_Q = Index_{Q-1} × (1+annual)^(1/4)`), Q1 = 1.00 as the base period; drivers pick which factor applies via a selector cell
- **Not built on `Assumptions_Periodic_Capex`.** Capex is entered as a nominal total with a phasing profile, so a capex escalation index would have nothing to multiply without also changing the "cumulative capex = total capex input" check. Deferred to Phase 3 where sector capex modules land — flagged rather than built as dead weight

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
- **Dirty-flag (✅ built):** `Cover` snapshots all 13 tracked inputs at solve time and compares live vs. stored. Status reads `SOLVED - current` or `RE-RUN SOLVE - assumptions changed`, and the table names *which* input moved. Recorded automatically on every successful solve; cleared by the reset macros. Per-input tracking beats a single hashed checksum here: same protection, but it tells you what changed, and it sidesteps float-precision games in a combined hash

### Verification
- `src/verify.py` drives the built workbook through **LibreOffice headless (UNO)**, running the same two-staged-cell convergence the VBA does, then reads every `Check_Control` row and the headline outputs across all scenarios and both DSRA funding methods
- This evaluates the workbook's **own formulas** rather than re-deriving what they should say. A harness that reimplements the intent agrees with itself — that is exactly how the pari-passu gearing bug survived its first review
- "Solve is current" reads FAIL under the harness, correctly: the snapshot is only written by the VBA's `RecordSolveSnapshot`, which the harness bypasses

### VBA — build process constraint
- `openpyxl` **cannot author VBA** — no COM/Excel in this environment
- All macros address the model through **named ranges only**, never cell literals — which is why Stage 1c's row moves needed no VBA changes at all
- Workflow: I hand you plain-text module code + a defined-names spec + button instructions; you paste into the VBA editor once (`.xlsx` → Save As `.xlsm` → `Alt+F11` → Insert Module → paste → Developer tab → Insert Form Control Button → Assign Macro)
- One-time setup per macro group; future regenerations use `keep_vba=True` to preserve what you've pasted

### Control panel (`Cover`, Stage 1c — ✅ built)
Sited on `Cover`, not `Assumptions_Constant` as originally drafted: every other solve
setting already lives there, and the buttons can only be drawn on one sheet, so splitting
the panel from its own buttons would have been worse than moving it.

- **Per-scenario:** Solve All (Current Scenario) | Goal Seek → EIRR | Goal Seek → PIRR | Solve Construction IDC | Solve Debt Sculpting | Reset | Invalidate Solve
- **Batch (confirmation prompt):** Run All 10 Scenarios, with a `Batch Mode` selector — Solve Only (non-destructive) or Solve + Goal Seek, which **overwrites every scenario's revenue**
- **Live display (no button needed):** Current EIRR/PIRR vs Target, on-target flag, plus the TPC / facility / gearing / min DSCR / min LLCR readings the batch runner records

**Goal seek is hand-rolled, not Excel's.** Built-in Goal Seek only recalculates; this model
needs both staged cells re-converged after any input moves, so built-in would read an IRR
computed off stale staged values and converge confidently on the wrong revenue. Each trial
revenue therefore costs a full silent solve.

**Bisection, not secant or Newton.** EIRR is monotone in revenue but not smooth — the Max
Gearing cap and the DSCR floor each put a kink in the curve, and a derivative-based step
can jump a kink and diverge. Bisection only needs the root bracketed, which the bound check
establishes.

**Feasibility at the bounds first**, before any bisection: two solves settle whether the
target is reachable and which side it is out on, instead of discovering it after the full
iteration budget. Verified to exit at *zero* bisection iterations on both out-of-range
paths, restoring the original revenue.

Measured on the dummy case: monotone across the whole 0.5x–2.0x band, 7–11 bisection
iterations per solve against a budget of 60.

**The driver is the active scenario's own revenue cell** on `Assumptions_Constant` — never
the Active column, which is an `INDEX` formula the macro would overwrite with a constant,
silently severing every scenario from the selector.

---

## Stage sequencing (with all identified gaps folded in)

| Stage | Tasks | Status |
|---|---|---|
| **1a — Plumbing proof** | Single scenario, flat dummy revenue, fixed-ratio debt, zero circularity | ✅ **complete** (tasks #1–10) |
| **Interim — Input centralization** | `Cover` + `Assumptions_Model`, rewire all `Calc_*` hardcodes to links | ✅ **complete** (task #18) |
| **1b — Circularity** | Loop 1 + drawdown selector + construction `FS_Annual` (#11 ✅) → Loop 2 + tax shield (#12 ✅) → dirty-flag check (#13 ✅) | ✅ **complete** |
| **1c — Scale out** | 10 scenarios + escalation library (#14 ✅) → DSRA/MRA + LC option + LLCR/PLCR + multi-vintage maintenance capex (#15 ✅) → control panel + goal-seek (#16 ✅) | ✅ **complete** |
| **1d — Sell-down** | `Valuation_SellDown` + `Dashboard` (#22) | ✅ **complete** |

**Dummy test case (Stage 1a validation):** $100M project, 24mo construction, 20yr ops, $15M/yr flat revenue, 70/30 debt/equity, 6% interest, 25% tax. Verified: BS balances all 80 quarters, model winds to exactly $0 at end of life, EIRR (7.16%) > PIRR (6.09%) correctly reflects leverage, EIRR moves monotonically with revenue.

---

## What NOT to do (guardrails, hard-won from earlier missteps)

- Don't put periodic scenario inputs with time-down-rows — breaks the universal time-across-columns rule
- Don't rebuild FCFF/FCFE independently on `FS_Quarterly` and `FS_Annual` — build once, roll up
- Don't recalculate PIRR per exit year in `Valuation_SellDown` — it's a whole-of-project unlevered number, doesn't vary by holding period
- Don't bisect for debt sculpting at all — lock DSCR via formula and the debt size falls out as a PV in closed form
- Don't discount the capped service row when computing sculpted capacity — use the uncapped basis, or the fixed point goes degenerate and silently "converges" at whatever it started from
- Don't let any drawdown branch reference the raw gearing input instead of the solved facility — the Pari Passu branch did exactly that and pegged implied gearing to the assumption, making Loop 2 look broken while every check still passed. `Check_Control` now guards this ("Debt draws honour the solved facility")
- Don't index another sheet with this module's own row constants — `Calc_Tax` did that and pulled Revenue while labelling it EBITDA, overstating supportable debt by 19%. Import the source module's row constants instead
- Don't let the verification harness re-derive a quantity the workbook takes as an input — the harness computed the pari-passu split as `facility/TPC` while the sheet used the gearing input, so the two disagreed and the bug survived verification. Mirror the formula, don't reimplement the intent
- Don't build scenario conditional-formatting highlighting now — explicitly Phase 2
- Don't skip the feasibility pre-check in multi-scenario goal-seek — burns iterations discovering what one bound-check would show instantly
- Don't trust VBA-solved values without the dirty-flag check — stale copy-pasted numbers look identical to fresh ones
- Don't move a row on an input sheet without checking who indexes it by literal. `Assumptions_Model`'s phasing block shifted two rows for the new resolved drivers and `Calc_Capex` still pointed at the old row — phasing read as zero, so capex, TPC, debt and both IRRs all collapsed to zero while most checks still said OK. Every cross-sheet row reference now comes from the source module's constant
- Don't add a cost that depends on debt service into the tax computation — see the LC fee note above. Anything feeding `Calc_Tax` must be upstream of debt sizing, or it closes a loop with no breaker
- Don't emit a post-2007 Excel function from openpyxl without the `_xlfn.` prefix. `MINIFS` was written bare, resolved to nothing, and `IFERROR` turned that into a clean-looking **0.000 min LLCR** — a wrong number that reads as a real one. Prefer functions old enough not to need the prefix: `SMALL(range, COUNTIF(range,"<=0")+1)` gets the smallest positive entry with no prefix and no array formula
- Don't wrap a formula in `IFERROR` and assume a plausible result means it worked — that is what hid the `MINIFS` failure. Cross-check any new summary cell against the row it summarises
- Don't let goal seek write to an Active/resolved cell. Those hold `INDEX` formulas; a macro writing a constant there detaches the whole scenario machinery from its selector, and nothing checks for it
- Don't write a check that is true by construction. "MRA balance ≥ MRA target" was tautological (the balance *is* the target); it was replaced with "prior quarter's balance ≥ this quarter's capex", which actually fails if the lookforward window is wired backwards
