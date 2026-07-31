"""Cover cell addresses, in one place that imports nothing.

Every sheet reads switches and tolerances off Cover. Those consumers cannot all import
cover.py — Assumptions_Constant is upstream of it — so the addresses lived as literals in
a dozen formulas. That has now caused three silent bugs: a moved phasing row zeroed capex,
and a moved DSRA selector left the model reading a header string and quietly running the
LC-backed branch in every scenario.

Nothing in this module may import a builder. cover.py imports these and lays the sheet out
from them, so the constants stay the single source of truth in both directions."""

# Solve settings
CELL_CIRC_TOLERANCE = "B4"
CELL_MAX_ITERATIONS = "B5"
CELL_DEBT_SIZING_TOLERANCE = "B6"

ABS_CIRC_TOLERANCE = "$B$4"
ABS_DEBT_SIZING_TOLERANCE = "$B$6"

CELL_MASTER_CHECK_LINK = "B9"

# Run controls
CELL_DRAWDOWN_METHOD = "B14"
CELL_DEBT_SIZING_MODE = "B17"
CELL_ACTIVE_SCENARIO = "B20"

ABS_DRAWDOWN_METHOD = "$B$14"
ABS_DEBT_SIZING_MODE = "$B$17"
ABS_ACTIVE_SCENARIO = "$B$20"

# Structuring options
ROW_STRUCTURING_HEADER = 23
ROW_DSRA_METHOD = 24
ROW_DSRA_TIMING = 25
ROW_TLCF_MODE = 26
ROW_LOCKUP_DSCR = 27
ROW_NEGATIVE_CASH = 28

CELL_DSRA_METHOD = f"B{ROW_DSRA_METHOD}"
ABS_DSRA_METHOD = f"$B${ROW_DSRA_METHOD}"
ABS_DSRA_TIMING = f"$B${ROW_DSRA_TIMING}"
ABS_TLCF_MODE = f"$B${ROW_TLCF_MODE}"
ABS_LOCKUP_DSCR = f"$B${ROW_LOCKUP_DSCR}"
ABS_NEGATIVE_CASH = f"$B${ROW_NEGATIVE_CASH}"

# Option values. Consumers compare against these strings, so a renamed option changes
# the dropdown and every formula that tests it together.
DRAWDOWN_METHODS = ["Debt First", "Equity First", "Pari Passu"]
DEBT_SIZING_MODES = ["Fixed Gearing", "DSCR Sculpted"]
DSRA_METHODS = ["Cash Funded", "LC-Backed"]
DSRA_TIMINGS = ["At Financial Close", "From Operating Cash"]
TLCF_MODES = ["Carried Forward", "Forfeited"]
NEGATIVE_CASH_MODES = ["Retain in Buffer", "Inject Equity"]
BATCH_MODES = ["Solve Only", "Solve + Goal Seek EIRR", "Solve + Goal Seek PIRR"]

DSRA_METHOD_CASH = DSRA_METHODS[0]
DSRA_TIMING_AT_CLOSE = DSRA_TIMINGS[0]
TLCF_ENABLED = TLCF_MODES[0]
NEGATIVE_CASH_RETAIN = NEGATIVE_CASH_MODES[0]
SCULPTED = DEBT_SIZING_MODES[1]

N_SCENARIOS = 10
