"""Run mod_GoalSeek's algorithm against the real workbook through LibreOffice.

The VBA itself cannot be executed here, so this mirrors it step for step — same driver
cell, same bounds, same feasibility test, same bisection — and checks the properties the
macro depends on: that the driver cell is writable and flows through, that the metric is
monotone in revenue, and that the search lands inside the iteration budget.

Mirroring matters more than covering. A harness that reimplemented "find the revenue that
hits the target" some smarter way would prove nothing about the macro that ships."""

import sys

from verify import _start_soffice, _load, get, set_value, solve, OUTPUT

# Assumptions_Constant!C11:L11 — ScenarioRevenueRow. Column C is scenario 1.
DRIVER_SHEET = "Assumptions_Constant"
DRIVER_ROW = 11
DRIVER_FIRST_COL = 3

LIVE = {"EIRR": ("Cover", "B29"), "PIRR": ("Cover", "B30")}
TARGET = {"EIRR": ("Cover", "B27"), "PIRR": ("Cover", "B28")}
MIN_MULT = ("Cover", "B37")
MAX_MULT = ("Cover", "B38")
TOLERANCE = ("Cover", "B39")
MAX_ITER = ("Cover", "B40")


def driver_cell(doc, scenario):
    from openpyxl.utils import get_column_letter
    col = get_column_letter(DRIVER_FIRST_COL + scenario - 1)
    return (DRIVER_SHEET, f"{col}{DRIVER_ROW}")


def evaluate_at(doc, scenario, revenue, metric):
    """One trial point: write revenue, re-converge both staged cells, read the metric."""
    set_value(doc, driver_cell(doc, scenario), revenue)
    solve(doc)
    return get(doc, LIVE[metric])


def seek(doc, scenario, metric, verbose=True):
    driver = driver_cell(doc, scenario)
    original = get(doc, driver)
    target = get(doc, TARGET[metric])
    tol = get(doc, TOLERANCE)
    max_iter = int(get(doc, MAX_ITER))
    lo = original * get(doc, MIN_MULT)
    hi = original * get(doc, MAX_MULT)

    f_lo = evaluate_at(doc, scenario, lo, metric)
    f_hi = evaluate_at(doc, scenario, hi, metric)

    if verbose:
        print(f"  target {metric} {target:.2%} | bounds {lo:,.0f} -> {hi:,.0f}")
        print(f"  {metric} at bounds: {f_lo:.4%} -> {f_hi:.4%}")

    if f_hi <= f_lo:
        return "NON-MONOTONE", None, f_lo, f_hi, 0

    if target > f_hi:
        set_value(doc, driver, original)
        solve(doc)
        return "INFEASIBLE", None, f_lo, f_hi, 0
    if target < f_lo:
        set_value(doc, driver, original)
        solve(doc)
        return "BELOW RANGE", None, f_lo, f_hi, 0

    for i in range(1, max_iter + 1):
        mid = (lo + hi) / 2
        f_mid = evaluate_at(doc, scenario, mid, metric)
        if abs(f_mid - target) <= tol:
            return "SOLVED", mid, f_lo, f_hi, i
        if f_mid < target:
            lo = mid
        else:
            hi = mid

    return "NOT CONVERGED", None, f_lo, f_hi, max_iter


def monotonicity_probe(doc, scenario, metric, points=7):
    """The bisection is only safe if the metric rises with revenue. Kinks from the Max
    Gearing cap and the DSCR floor are fine; a reversal is not."""
    driver = driver_cell(doc, scenario)
    original = get(doc, driver)
    lo, hi = original * get(doc, MIN_MULT), original * get(doc, MAX_MULT)

    samples = []
    for k in range(points):
        rev = lo + (hi - lo) * k / (points - 1)
        samples.append((rev, evaluate_at(doc, scenario, rev, metric)))

    set_value(doc, driver, original)
    solve(doc)

    reversals = [
        (samples[k][0], samples[k][1], samples[k + 1][1])
        for k in range(len(samples) - 1)
        if samples[k + 1][1] < samples[k][1] - 1e-9
    ]
    return samples, reversals


def main():
    proc, ctx = _start_soffice()
    failures = []
    try:
        doc = _load(ctx, OUTPUT)
        set_value(doc, ("Cover", "B20"), 1)
        solve(doc)

        print("=" * 78)
        print("MONOTONICITY — EIRR vs revenue, Base scenario")
        print("=" * 78)
        samples, reversals = monotonicity_probe(doc, 1, "EIRR")
        for rev, val in samples:
            print(f"  revenue {rev:>13,.0f}   EIRR {val:>9.4%}")
        if reversals:
            print(f"  !! {len(reversals)} reversal(s) — bisection assumption broken")
            failures.append(("monotonicity", reversals))
        else:
            print("  monotone increasing — bisection bracket is valid")

        for scenario, name in ((1, "Base"), (2, "Upside"), (3, "Downside")):
            for metric in ("EIRR", "PIRR"):
                print("\n" + "=" * 78)
                print(f"GOAL SEEK — scenario {scenario} ({name}), target {metric}")
                print("=" * 78)
                set_value(doc, ("Cover", "B20"), scenario)
                solve(doc)
                original = get(doc, driver_cell(doc, scenario))

                status, found, f_lo, f_hi, iters = seek(doc, scenario, metric)
                if status == "SOLVED":
                    achieved = get(doc, LIVE[metric])
                    print(f"  {status}: revenue {found:,.0f} "
                          f"({found / original:.3f}x) -> {metric} {achieved:.4%} in {iters} iters")
                else:
                    print(f"  {status} (bounds gave {f_lo:.4%} .. {f_hi:.4%})")
                    if status in ("NON-MONOTONE", "NOT CONVERGED"):
                        failures.append((f"{name}/{metric}", status))

                # Restore so the next scenario starts from the shipped assumptions.
                set_value(doc, driver_cell(doc, scenario), original)
                solve(doc)

        print("\n" + "=" * 78)
        if failures:
            print("ALGORITHM PROBLEMS:")
            for f in failures:
                print("  ", f)
        else:
            print("Bisection is sound: monotone bracket, every feasible target converged.")
        doc.close(False)
        return 1 if failures else 0
    finally:
        try:
            ctx.ServiceManager.createInstanceWithContext(
                "com.sun.star.frame.Desktop", ctx).terminate()
        except Exception:
            pass
        proc.wait(timeout=30)


if __name__ == "__main__":
    sys.exit(main())
