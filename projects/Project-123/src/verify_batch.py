"""Mirrors mod_GoalSeek's RunAllScenarios end to end through LibreOffice.

Neither verify.py nor verify_goalseek.py exercised the batch loop itself -- and both
bugs found in this workbook (staged values carried across scenario switches, snapshot
read before being refreshed) lived exactly there. This walks all N_SCENARIOS the same
way the VBA does: reset staged cells, solve, snapshot, record, move on -- then restores
the original scenario and checks it, same as the macro's own final step."""

import sys

import assumptions_constant as const
import cover
import cover_refs as refs
from verify import (
    OUTPUT,
    _cell,
    _load,
    _start_soffice,
    get,
    get_str,
    reset_staged,
    set_value,
    solve,
)


def record_snapshot(doc):
    """Mirrors mod_SolveFreshness's RecordSolveSnapshot: copy every tracked input's live
    value into the stored column, so the dirty-flag check reads as fresh."""
    live_col = "B"
    stored_col = "C"
    first_row = cover.ROW_FIRST_SNAPSHOT
    last_row = first_row + len(cover.TRACKED_INPUTS) - 1
    for row in range(first_row, last_row + 1):
        live = _cell(doc, "Cover", f"{live_col}{row}")
        stored = _cell(doc, "Cover", f"{stored_col}{row}")
        # UNO's getValue() on a genuinely text cell reads 0 -- none of the tracked numeric
        # inputs (capex, rates, tenor...) are legitimately exactly zero, so this cleanly
        # separates the numeric drivers from the text ones (Drawdown Method, DSRA Method...).
        if live.getValue() == 0 and live.getString():
            stored.setString(live.getString())
        else:
            stored.setValue(live.getValue())
    doc.calculateAll()


def main():
    proc, ctx = _start_soffice()
    try:
        doc = _load(ctx, OUTPUT)
        original_scenario = int(get(doc, ("Cover", refs.CELL_ACTIVE_SCENARIO)))

        rows = []
        print("=" * 100)
        print(f"BATCH — walking all {const.N_SCENARIOS} scenarios (Solve Only, mirrors RunAllScenarios)")
        print("=" * 100)

        for s in range(1, const.N_SCENARIOS + 1):
            set_value(doc, ("Cover", refs.CELL_ACTIVE_SCENARIO), s)
            doc.calculateAll()

            reset_staged(doc)
            passes = solve(doc)
            converged = passes < 40

            record_snapshot(doc)

            # B4 is the Active column's own INDEX formula, already tracking whichever
            # scenario ActiveScenario just pointed it at -- no need to address the
            # scenario's own column directly.
            name = get_str(doc, "Assumptions_Constant", "B4")
            master_flag = get_str(doc, "Check_Control", "B3")
            tpc = get(doc, ("Cover", f"B{cover.ROW_LIVE_TPC}"))
            eirr = get(doc, ("Cover", f"B{cover.ROW_LIVE_EIRR}"))
            pirr = get(doc, ("Cover", f"B{cover.ROW_LIVE_PIRR}"))

            rows.append((s, name, converged, master_flag, tpc, eirr, pirr))
            marker = "  " if master_flag == "MODEL OK" else ">>"
            print(f"{marker} scenario {s:>2} {name:<14} solve={'OK' if converged else 'NOT CONV':<9} "
                  f"TPC {tpc:>14,.0f}  EIRR {eirr:>8.2%}  PIRR {pirr:>8.2%}  {master_flag}")

        set_value(doc, ("Cover", refs.CELL_ACTIVE_SCENARIO), original_scenario)
        doc.calculateAll()
        reset_staged(doc)
        passes = solve(doc)
        final_converged = passes < 40
        if final_converged:
            record_snapshot(doc)
        final_flag = get_str(doc, "Check_Control", "B3")

        print("\n" + "=" * 100)
        print(f"Restored to scenario {original_scenario}: "
              f"{'converged' if final_converged else 'NOT CONVERGED'}, Check_Control = {final_flag}")

        bad_scenarios = [r for r in rows if r[3] != "MODEL OK"]
        bad_final = not final_converged or final_flag != "MODEL OK"

        print("=" * 100)
        if bad_scenarios or bad_final:
            print(f"{len(bad_scenarios)} scenario(s) not MODEL OK after their own solve"
                  + (", and the final restore did not reconverge" if bad_final else "") + ":")
            for r in bad_scenarios:
                print("  ", r[:4])
        else:
            print("Batch runner logic sound: every scenario solved and snapshotted cleanly, "
                  "final restore reconverged.")

        doc.close(False)
        return 1 if (bad_scenarios or bad_final) else 0
    finally:
        try:
            ctx.ServiceManager.createInstanceWithContext(
                "com.sun.star.frame.Desktop", ctx).terminate()
        except Exception:
            pass
        proc.wait(timeout=30)


if __name__ == "__main__":
    sys.exit(main())
