# VBA Setup — one-time manual step

`openpyxl` cannot author VBA (no Excel/COM in the build environment), so macros are
hand-authored here and pasted in once. After that the workbook becomes the **build
template**, and every future build inherits its VBA project. You only repeat this if the
macro code itself changes.

## What survives a rebuild, and what doesn't

| Thing | Survives? | Why |
|---|---|---|
| Macro code (`.bas` modules and `ThisWorkbook`) | yes | Lives in `xl/vbaProject.bin`, which openpyxl copies across verbatim |
| Named ranges the macros address | yes | The build regenerates them every time |
| **Form Control buttons drawn by hand** | **no** | They live in the sheet XML, which the build regenerates from scratch |

That last row is the reason `ThisWorkbook` redraws the buttons on open, reading the
macro/label table the build writes to `Cover!ButtonSpec` (columns H-I).

**Do not draw buttons by hand** — they disappear on the next build. Adding a button later
is a build-side change, not something you paste.

## One-time setup

1. Open the generated `.xlsx` in Excel.
2. **File -> Save As -> Excel Macro-Enabled Workbook (`.xlsm`)**. VBA cannot live in an
   `.xlsx`. Name it `project123_template.xlsm`.
3. `Alt + F11` to open the VBA editor.
4. **File -> Import File...** and select each of `mod_Loop1_IDC.bas`,
   `mod_Loop2_DebtSizing.bas`, `mod_SolveFreshness.bas` and `mod_GoalSeek.bas`. Importing
   rather than pasting keeps the module names. All four are required — the solve macros
   call into the freshness module, and goal seek calls into the Loop 2 module.
5. In the project tree on the left, expand **Microsoft Excel Objects** and double-click
   **ThisWorkbook**. Paste the whole of `ThisWorkbook.txt` into the code pane that opens.
   This one is a code-behind object, *not* a module — importing it as a module will not
   work, because `Workbook_Open` only fires from `ThisWorkbook`.
6. Close the VBA editor and **save**, keeping `.xlsm`.
7. Close the file and reopen it. The buttons draw themselves on `Cover`. If nothing
   appears, macros are disabled — see Troubleshooting.
8. Send the `.xlsm` back so it can be committed as the build template.

If the Developer tab is not visible: **File -> Options -> Customize Ribbon -> tick
Developer**. You no longer need it for buttons, but it is where the macro list lives.

## Updating a template when the macro code changes

Only needed when a task adds or changes macro logic. To replace a module:

1. `Alt + F11`, right-click the module in the tree, **Remove <name>**, answer **No** to
   the export prompt.
2. **File -> Import File...** and pick the new version.
3. Save, close, reopen, and send the `.xlsm` back.

Buttons need no attention — `Workbook_Open` redraws them from `Cover!ButtonSpec`, so a
new macro appears as a button as soon as the build lists it.

## Rebuilding from then on

    python src/build.py --template

Reads `template/project123_template.xlsm`, drops every sheet, rebuilds them, and writes an
`.xlsm` with the macros already inside. Nothing to paste.

Drop the flag and you get a plain `.xlsx` with no macros — useful for diffing and for the
LibreOffice verification harness.

## The buttons

| Button | Macro | What it does | When |
|---|---|---|---|
| **Solve All (Current Scenario)** | `SolveAllCurrentScenario` | Alternates Loop 1 and Loop 2 until both converge, then stamps the solve snapshot | **The one you press.** After any assumption change |
| Goal Seek -> Target EIRR | `GoalSeekEIRR` | Solves the active scenario's Annual Revenue for the target EIRR on `Cover` | "What revenue do we need to clear our hurdle?" |
| Goal Seek -> Target PIRR | `GoalSeekPIRR` | Same, against the target PIRR | Unlevered version of the same question |
| Run All 10 Scenarios (Batch) | `RunAllScenarios` | Walks every scenario, solves (and optionally goal-seeks) each, writes `Batch_Results` | Producing the scenario comparison table |
| Solve Construction IDC | `SolveConstructionIDC` | Loop 1 only — IDC / debt balance / Total Project Cost | Diagnosing the construction loop alone |
| Solve Debt Sculpting | `SolveDebtSculpting` | Loop 2 only — debt size / interest / tax / CFADS / capacity | Diagnosing the operations loop alone |
| Reset Staged Values | `ResetAllStagedValues` | Zeroes both staged cells, clears the snapshot | The solve has wandered and you want a clean start |
| Invalidate Solve | `InvalidateSolveSnapshot` | Clears the snapshot only, forcing `RE-RUN SOLVE` | Forcing the dirty flag on when you suspect stale numbers |

### Goal seek

The driver is **Annual Revenue on the active scenario's own column** of
`Assumptions_Constant` — never the Active column, which is an `INDEX` formula the macro
must not overwrite.

Excel's built-in Goal Seek cannot be used: it only recalculates, and this model needs the
two staged cells re-converged after every input change. Built-in Goal Seek would read an
IRR computed off stale staged values. So each trial revenue costs a full silent solve, and
the search is bisection — EIRR is monotone in revenue but has kinks where the Max Gearing
cap and the DSCR floor bind, and a derivative step can jump a kink and diverge.

Bounds are **multiples of the scenario's current revenue** (default 0.50x to 2.00x), so one
setting works across all ten. Feasibility is tested at both bounds *first*: if the target is
out of range you get `INFEASIBLE` or `BELOW RANGE` after two solves rather than after sixty,
and the revenue is put back where it was.

Measured on the dummy case: 7-11 bisection iterations per solve, against a budget of 60.

### Batch runner

`Batch Mode` on `Cover` decides what each scenario gets:

- **Solve Only** — non-destructive; just solves and records.
- **Solve + Goal Seek EIRR / PIRR** — **overwrites every scenario's Annual Revenue** with
  the value that hits the target. Not undoable, which is why it asks first.

Results land on `Batch_Results` as values, not formulas — a formula would recalculate to
whichever scenario is selected when the run ends, making all ten rows identical. The
scenario selector is restored when the run finishes, including after an error.

The two loops are coupled — debt size sets the construction facility, which moves IDC,
which moves Total Project Cost, depreciation, tax, CFADS and therefore capacity — so
running them together is what lands on a mutually consistent answer. The single-loop
macros are for diagnosis, not normal use.

## After pressing Solve All

Check two cells before believing anything:

- `Cover!B54` should read **SOLVED - current**. If it says `RE-RUN SOLVE - assumptions
  changed`, what is on screen is stale, and the table below it names the input that moved.
- `Check_Control!B3` should read **MODEL OK**.

## Troubleshooting

**No buttons on open.** Macros are disabled. **File -> Options -> Trust Center -> Trust
Center Settings -> Macro Settings -> Disable all macros with notification**, reopen, click
*Enable Content*. Better: add the project folder as a **Trusted Location** so you are not
asked again.

**Buttons appear but do nothing.** A module failed to import. `Alt + F11` and check all
three `.bas` modules are listed under **Modules**.

**"Cannot run the macro" on click.** The button points at a macro that no longer exists.
Press `Alt + F8`, run `DrawControlPanel`, and it redraws from the current table.

**Buttons sit somewhere awkward.** They anchor at `BUTTON_LEFT` / `BUTTON_TOP` in
`ThisWorkbook`. Change those constants and reopen.

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
| `Cover_DSRAMethod` | `Cover!$B$23` | Cash Funded / LC-Backed |
| `ActiveScenario` | `Cover!$B$20` | Scenario selector, 1-10 |
| `ButtonSpec` | `Cover!$H$3:$I$12` | Macro/label table `Workbook_Open` draws buttons from |
| `ScenarioRevenueRow` | `Assumptions_Constant!$C$11:$L$11` | Goal seek writes into the active scenario's cell in this row |
| `GoalSeek_TargetEIRR` / `_TargetPIRR` | `Cover!$B$27` / `$B$28` | Targets the goal seek aims at |
| `GoalSeek_MinMultiple` / `_MaxMultiple` | `Cover!$B$37` / `$B$38` | Search bounds, as multiples of current revenue |
| `GoalSeek_Tolerance` / `_MaxIterations` | `Cover!$B$39` / `$B$40` | Stopping conditions |
| `GoalSeek_Status` | `Cover!$B$33` | Result text the macro writes |
| `Cover_BatchMode` | `Cover!$B$41` | Solve Only / Solve + Goal Seek EIRR / PIRR |
| `Live_EIRR` / `Live_PIRR` | `Cover!$B$29` / `$B$30` | Live returns the goal seek reads |
| `Live_TPC`, `Live_DebtFacility`, `Live_Gearing`, `Live_MinDSCR`, `Live_MinLLCR` | `Cover!$B$45:$B$49` | Readings the batch runner records per scenario |
| `BatchResults_Anchor` | `Batch_Results!$A$7` | First scenario row the batch runner writes to |
| `SnapshotLive` | `Cover!$B$57:$B$73` | Live value of every tracked input |
| `SnapshotStored` | `Cover!$C$57:$C$73` | Those values as at the last solve |
| `LastSolvedStamp` | `Cover!$B$53` | Timestamp written on each successful solve |
| `SolveStatus` | `Cover!$B$54` | `SOLVED - current` / `RE-RUN SOLVE ...` |
| `SolveFreshnessFlag` | `Cover!$D$75` | 1/0 flag `Check_Control` pulls |

## What "solved" looks like

`Check_Control` should show every check OK and the master flag `MODEL OK`.

Expected results for the $100M dummy case as at Stage 1c (Pari Passu drawdown, **DSCR
Sculpted**, Target DSCR 1.30x, DSRA cash-funded), converging in ~7 outer passes:

| Scenario | Total Project Cost | Sculpted debt | Implied gearing | Min DSCR | LLCR / PLCR at Q1 | PIRR / EIRR |
|---|---|---|---|---|---|---|
| Base | $104.48M | $75.61M | 72.37% | 1.3000 | 1.35 / 1.58 | 6.34% / 5.94% |
| Downside | $113.56M | $51.27M | 45.15% | 1.4000 | 1.46 / 1.68 | 2.36% / -0.52% |
| Base, LC-backed DSRA | $104.48M | $75.61M | 72.37% | 1.3000 | 1.30 / 1.52 | 6.34% / 6.37% |

LC-backing lifts EIRR (no cash trapped) and lowers LLCR (no DSRA balance in the
numerator). Both moves are the right sign — if yours go the other way, something is wired
backwards.

Two signatures tell you the sculpting is working: **min DSCR over the tenor equals the
target exactly** (1.3000 — that is what "locked" means), and the **closing balance at
tenor end is zero**. If DSCR drifts above target across the whole tenor, the debt is
undersized; if the balance does not reach zero, it is oversized.

In **Fixed Gearing** mode the schedule reverts to level PMT amortisation and the Loop 2
convergence check is skipped, but `SculptedDebtCapacity` still reports what the project
*could* support — useful as a sizing sense-check against the fixed assumption.

## Stale values — the freshness guard

The staged cells hold *numbers*, not formulas. Change an assumption without re-solving and
every downstream figure still calculates happily off the old staged values: the model looks
right and is wrong, with nothing on screen to say so.

`Cover` guards against this. It snapshots every input a solve depends on, and the
**Solve Status** cell reads either:

- `SOLVED - current` — nothing has moved since the last solve
- `RE-RUN SOLVE - assumptions changed` — at least one input has drifted

The table beneath it names *which* input moved (Match = 0), so you are not left guessing.
`Check_Control` pulls the same flag, so a stale model also trips the master status.

A freshly generated workbook reports `RE-RUN SOLVE` by design — nothing has been solved
yet. The snapshot is recorded automatically on every successful solve; the reset macros
clear it so a reset model cannot keep claiming to be solved.
