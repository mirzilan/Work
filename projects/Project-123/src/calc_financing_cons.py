from openpyxl.styles import Font
from openpyxl.worksheet.worksheet import Worksheet
from openpyxl.workbook import Workbook

import assumptions_model as model
import calc_financing_ops as fin_ops
import calc_revenue_opex as rev_opex
import cover_refs as refs
from inputs import ProjectInputs
from timeline import Timeline
from workbook_builder import (
    FIRST_DATA_COL,
    COLOR_INPUT,
    COLOR_FORMULA,
    COLOR_LINK,
    TAB_COLOR_CALC,
    col_letter,
)

CELL_DRAWDOWN_METHOD = "B3"
CELL_INTEREST_RATE = "B4"
CELL_GEARING = "B5"
CELL_STAGED_IDC = "B6"
CELL_TOTAL_PROJECT_COST = "B7"
CELL_DEBT_FACILITY = "B8"
CELL_EQUITY_COMMITMENT = "B9"
CELL_CALCULATED_IDC = "B10"
CELL_CONVERGENCE_GAP = "B11"
CELL_INITIAL_DSRA = "B12"
CELL_INITIAL_BUFFER = "B13"
CELL_TOTAL_FUNDING_REQ = "B14"

# Absolute forms — scalar cells referenced from the periodic columns must not shift on fill
ABS_DRAWDOWN_METHOD = "$B$3"
ABS_INTEREST_RATE = "$B$4"
ABS_GEARING = "$B$5"
ABS_TOTAL_PROJECT_COST = "$B$7"
ABS_DEBT_FACILITY = "$B$8"
ABS_EQUITY_COMMITMENT = "$B$9"
ABS_INITIAL_DSRA = "$B$12"
ABS_INITIAL_BUFFER = "$B$13"

ROW_DATE_HEADER = 13
ROW_MONTH_INDEX = 14

ROW_CAPEX_DRAW = 16
ROW_OPENING_BAL = 17
ROW_INTEREST_ACCRUED = 18
ROW_FUNDING_REQUIREMENT = 19
ROW_CUM_FUNDING_REQUIREMENT = 20
ROW_DEBT_DRAW = 21
ROW_EQUITY_DRAW = 22
ROW_CUM_DEBT_DRAW = 23
ROW_CUM_EQUITY_DRAW = 24
ROW_CLOSING_BAL = 25

ROW_CHECK_HEADER = 28
ROW_CHECK_CLOSING_MATCHES_DRAWS = 29
ROW_CHECK_IDC_CONVERGED = 30
ROW_CHECK_SOURCES_TIE_USES = 31
ROW_CHECK_WITHIN_FACILITY = 32
ROW_CHECK_FACILITY_FULLY_DRAWN = 33


def build_calc_financing_cons(wb: Workbook, timeline: Timeline, inputs: ProjectInputs) -> Worksheet:
    ws = wb.create_sheet("Calc_Financing_Cons")
    ws.sheet_properties.tabColor = TAB_COLOR_CALC

    ws["A1"] = "Calc_Financing_Cons — Monthly Construction Funding & IDC (Loop 1)"
    ws["A1"].font = Font(bold=True, size=12)

    ws["A3"] = "Drawdown Method — linked from Cover"
    ws[CELL_DRAWDOWN_METHOD] = "=Cover!$B$14"
    ws[CELL_DRAWDOWN_METHOD].font = Font(color=COLOR_LINK)

    ws["A4"] = "Interest Rate (Annual) — linked from Assumptions_Model"
    ws[CELL_INTEREST_RATE] = "=Assumptions_Model!$B$5"
    ws[CELL_INTEREST_RATE].font = Font(color=COLOR_LINK)
    ws[CELL_INTEREST_RATE].number_format = "0.00%"

    ws["A5"] = "Gearing (Debt % of Total Project Cost) — linked from Assumptions_Model"
    ws[CELL_GEARING] = "=Assumptions_Model!$B$4"
    ws[CELL_GEARING].font = Font(color=COLOR_LINK)
    ws[CELL_GEARING].number_format = "0.00%"

    ws["A6"] = "Staged IDC ($) — VBA-written, breaks Loop 1 circularity"
    ws[CELL_STAGED_IDC] = 0
    ws[CELL_STAGED_IDC].font = Font(color=COLOR_INPUT, bold=True)
    ws[CELL_STAGED_IDC].number_format = "#,##0"

    ws["A7"] = "Total Project Cost ($) = Total Capex + Staged IDC"
    ws[CELL_TOTAL_PROJECT_COST] = f"=Assumptions_Model!$B$3+{CELL_STAGED_IDC}"
    ws[CELL_TOTAL_PROJECT_COST].font = Font(color=COLOR_FORMULA)
    ws[CELL_TOTAL_PROJECT_COST].number_format = "#,##0"

    # Under DSCR Sculpted the facility is whatever Loop 2 has solved the debt size to be;
    # under Fixed Gearing it is simply the gearing applied to Total Project Cost.
    ws["A8"] = "Debt Facility ($) — per Cover Debt Sizing Mode"
    ws[CELL_DEBT_FACILITY] = (
        f'=IF(Cover!{refs.ABS_DEBT_SIZING_MODE}="{refs.SCULPTED}",'
        f"Calc_Financing_Ops!{fin_ops.ABS_STAGED_DEBT_SIZE},{ABS_GEARING}*{ABS_TOTAL_PROJECT_COST})"
    )
    ws[CELL_DEBT_FACILITY].font = Font(color=COLOR_FORMULA)
    ws[CELL_DEBT_FACILITY].number_format = "#,##0"

    ws["A9"] = "Equity Commitment ($) = Total Funding Requirement - Debt Facility"
    ws[CELL_EQUITY_COMMITMENT] = f"={CELL_TOTAL_FUNDING_REQ}-{CELL_DEBT_FACILITY}"
    ws[CELL_EQUITY_COMMITMENT].font = Font(color=COLOR_FORMULA)
    ws[CELL_EQUITY_COMMITMENT].number_format = "#,##0"

    dsra_quarters = max(1, round(inputs.reserves.dsra_target_months / 3))
    n_months = len(timeline.construction_months)
    first_col = col_letter(0)
    last_col = col_letter(n_months - 1)

    ws["A10"] = "Calculated IDC ($) = SUM of monthly interest accrued"
    ws[CELL_CALCULATED_IDC] = f"=SUM({first_col}{ROW_INTEREST_ACCRUED}:{last_col}{ROW_INTEREST_ACCRUED})"
    ws[CELL_CALCULATED_IDC].font = Font(color=COLOR_FORMULA, bold=True)
    ws[CELL_CALCULATED_IDC].number_format = "#,##0"

    ws["A11"] = "Convergence Gap ($) = Calculated IDC - Staged IDC"
    ws[CELL_CONVERGENCE_GAP] = f"={CELL_CALCULATED_IDC}-{CELL_STAGED_IDC}"
    ws[CELL_CONVERGENCE_GAP].font = Font(color=COLOR_FORMULA, bold=True)
    ws[CELL_CONVERGENCE_GAP].number_format = "#,##0.00"

    # Funding the opening DSRA at close, rather than trapping it out of the first
    # quarter's operating cash.
    #
    # Sized off *scheduled level debt service* on the facility, not off the solved DSRA
    # balance. The solved balance looks like the natural reference and is a circular one:
    # the operations opening debt balance is cumulative construction draws, so making the
    # draws depend on the solved balance closes a loop that no staged cell breaks, and
    # Excel returns Err:522 across the whole model. Level service on the facility depends
    # only on the staged debt size, which is already a breaker.
    #
    # This is also how it is done in practice: the reserve is sized at close off the
    # base-case schedule and funded as a fixed amount. Any difference against the first
    # quarter's actual requirement shows up as a small top-up or release in Q1.
    ws["A12"] = "Initial DSRA Funded at Close ($) — zero when LC-backed or funded from ops"
    ws[CELL_INITIAL_DSRA] = (
        f'=IF(AND(Cover!{refs.ABS_DSRA_TIMING}="{refs.DSRA_TIMING_AT_CLOSE}",'
        f'Cover!{refs.ABS_DSRA_METHOD}="{refs.DSRA_METHOD_CASH}"),'
        f"-PMT({ABS_INTEREST_RATE}/4,Assumptions_Model!$B${model.ROW_DEBT_TENOR_YEARS}*4,"
        f"{CELL_DEBT_FACILITY})*{dsra_quarters},0)"
    )
    ws[CELL_INITIAL_DSRA].font = Font(color=COLOR_LINK)
    ws[CELL_INITIAL_DSRA].number_format = "#,##0"

    # Initial working capital, funded at close like any other cost. Without it the buffer
    # starts empty, so the first quarter has nothing to draw on and any gap between the
    # DSRA sized at close and the first actual requirement becomes an equity call —
    # which is what Upside was still showing.
    ws["A13"] = "Initial Cash Buffer Funded at Close ($)"
    ws[CELL_INITIAL_BUFFER] = (
        f"=Cover!{refs.ABS_CASH_BUFFER_TARGET}*Calc_Revenue_Opex!{col_letter(0)}{rev_opex.ROW_OPEX}"
    )
    ws[CELL_INITIAL_BUFFER].font = Font(color=COLOR_LINK)
    ws[CELL_INITIAL_BUFFER].number_format = "#,##0"

    ws["A14"] = "Total Funding Requirement ($) = TPC + Initial DSRA + Initial Buffer"
    ws[CELL_TOTAL_FUNDING_REQ] = f"={CELL_TOTAL_PROJECT_COST}+{CELL_INITIAL_DSRA}+{CELL_INITIAL_BUFFER}"
    ws[CELL_TOTAL_FUNDING_REQ].font = Font(color=COLOR_FORMULA, bold=True)
    ws[CELL_TOTAL_FUNDING_REQ].number_format = "#,##0"

    _label(ws, ROW_DATE_HEADER, "Period End Date")
    _label(ws, ROW_MONTH_INDEX, "Construction Month #")
    _label(ws, ROW_CAPEX_DRAW, "Capex Draw ($) — linked from Calc_Capex")
    _label(ws, ROW_OPENING_BAL, "Opening Debt Balance ($)")
    _label(ws, ROW_INTEREST_ACCRUED, "Interest Accrued / IDC ($)")
    _label(ws, ROW_FUNDING_REQUIREMENT, "Funding Requirement ($) = Capex Draw + Interest")
    _label(ws, ROW_CUM_FUNDING_REQUIREMENT, "Cumulative Funding Requirement ($)")
    _label(ws, ROW_DEBT_DRAW, "Debt Draw ($) — per Drawdown Method")
    _label(ws, ROW_EQUITY_DRAW, "Equity Draw ($)")
    _label(ws, ROW_CUM_DEBT_DRAW, "Cumulative Debt Draw ($)")
    _label(ws, ROW_CUM_EQUITY_DRAW, "Cumulative Equity Draw ($)")
    _label(ws, ROW_CLOSING_BAL, "Closing Debt Balance ($)")

    ws.cell(row=ROW_CHECK_HEADER, column=1, value="Checks").font = Font(bold=True)
    _label(ws, ROW_CHECK_CLOSING_MATCHES_DRAWS, "Check: Closing Balance = Cumulative Debt Draws")
    _label(ws, ROW_CHECK_IDC_CONVERGED, "Check: IDC Converged (|gap| <= Cover tolerance)")
    _label(ws, ROW_CHECK_SOURCES_TIE_USES, "Check: Cum Debt + Cum Equity = Cum Funding Requirement")
    _label(ws, ROW_CHECK_WITHIN_FACILITY, "Check: Cumulative Debt Draw <= Debt Facility")
    _label(ws, ROW_CHECK_FACILITY_FULLY_DRAWN, "Check: Cumulative Debt Draw = Debt Facility (draws honour the solve)")

    for i, period in enumerate(timeline.construction_months):
        col = col_letter(i)
        prev = col_letter(i - 1) if i > 0 else None

        ws[f"{col}{ROW_DATE_HEADER}"] = period.end
        ws[f"{col}{ROW_DATE_HEADER}"].number_format = "mmm-yy"
        ws[f"{col}{ROW_MONTH_INDEX}"] = i + 1

        _link(ws, col, ROW_CAPEX_DRAW, f"Calc_Capex!{col}6")

        if i == 0:
            _formula(ws, col, ROW_OPENING_BAL, "=0")
        else:
            _formula(ws, col, ROW_OPENING_BAL, f"={prev}{ROW_CLOSING_BAL}")

        _formula(ws, col, ROW_INTEREST_ACCRUED, f"={col}{ROW_OPENING_BAL}*{ABS_INTEREST_RATE}/12")
        funding = f"={col}{ROW_CAPEX_DRAW}+{col}{ROW_INTEREST_ACCRUED}"
        if i == n_months - 1:
            funding += f"+{CELL_INITIAL_DSRA}+{CELL_INITIAL_BUFFER}"
        _formula(ws, col, ROW_FUNDING_REQUIREMENT, funding)

        if i == 0:
            _formula(ws, col, ROW_CUM_FUNDING_REQUIREMENT, f"={col}{ROW_FUNDING_REQUIREMENT}")
        else:
            _formula(
                ws,
                col,
                ROW_CUM_FUNDING_REQUIREMENT,
                f"={prev}{ROW_CUM_FUNDING_REQUIREMENT}+{col}{ROW_FUNDING_REQUIREMENT}",
            )

        _formula(ws, col, ROW_DEBT_DRAW, _debt_draw_formula(col, prev))
        _formula(ws, col, ROW_EQUITY_DRAW, f"={col}{ROW_FUNDING_REQUIREMENT}-{col}{ROW_DEBT_DRAW}")

        if i == 0:
            _formula(ws, col, ROW_CUM_DEBT_DRAW, f"={col}{ROW_DEBT_DRAW}")
            _formula(ws, col, ROW_CUM_EQUITY_DRAW, f"={col}{ROW_EQUITY_DRAW}")
        else:
            _formula(ws, col, ROW_CUM_DEBT_DRAW, f"={prev}{ROW_CUM_DEBT_DRAW}+{col}{ROW_DEBT_DRAW}")
            _formula(ws, col, ROW_CUM_EQUITY_DRAW, f"={prev}{ROW_CUM_EQUITY_DRAW}+{col}{ROW_EQUITY_DRAW}")

        _formula(ws, col, ROW_CLOSING_BAL, f"={col}{ROW_OPENING_BAL}+{col}{ROW_DEBT_DRAW}")

    _check(ws, last_col, ROW_CHECK_CLOSING_MATCHES_DRAWS,
           f"=IF(ROUND({last_col}{ROW_CLOSING_BAL}-{last_col}{ROW_CUM_DEBT_DRAW},2)=0,1,0)")
    _check(ws, last_col, ROW_CHECK_IDC_CONVERGED,
           f"=IF(ABS({CELL_CONVERGENCE_GAP})<=Cover!$B$4,1,0)")
    _check(ws, last_col, ROW_CHECK_SOURCES_TIE_USES,
           f"=IF(ROUND({last_col}{ROW_CUM_DEBT_DRAW}+{last_col}{ROW_CUM_EQUITY_DRAW}"
           f"-{last_col}{ROW_CUM_FUNDING_REQUIREMENT},2)=0,1,0)")
    # Facility tolerance tracks the IDC convergence tolerance: a solve converged to within
    # Cover!$B$4 leaves the derived facility uncertain by the same order, so a tighter
    # hardcoded epsilon here would false-alarm on a correctly converged model.
    _check(ws, last_col, ROW_CHECK_WITHIN_FACILITY,
           f"=IF({last_col}{ROW_CUM_DEBT_DRAW}<={ABS_DEBT_FACILITY}+Cover!$B$4,1,0)")

    # Regression guard: if the draw formulas ever stop referencing the facility (as the
    # Pari Passu branch once did, silently pegging gearing to the raw input and ignoring
    # the Loop 2 solve entirely), cumulative draws diverge from the facility and this trips.
    _check(ws, last_col, ROW_CHECK_FACILITY_FULLY_DRAWN,
           f"=IF(ABS({last_col}{ROW_CUM_DEBT_DRAW}-{ABS_DEBT_FACILITY})<=Cover!$B$4,1,0)")

    _add_named_range(wb, "FinCons_InitialDSRA", "Calc_Financing_Cons", CELL_INITIAL_DSRA)
    _add_named_range(wb, "FinCons_InitialBuffer", "Calc_Financing_Cons", CELL_INITIAL_BUFFER)
    _add_named_range(wb, "FinCons_TotalFundingReq", "Calc_Financing_Cons", CELL_TOTAL_FUNDING_REQ)
    _add_named_range(wb, "StagedIDC", "Calc_Financing_Cons", CELL_STAGED_IDC)
    _add_named_range(wb, "CalculatedIDC", "Calc_Financing_Cons", CELL_CALCULATED_IDC)
    _add_named_range(wb, "IDCConvergenceGap", "Calc_Financing_Cons", CELL_CONVERGENCE_GAP)
    _add_named_range(wb, "FinCons_ClosingBalance_Last", "Calc_Financing_Cons", f"{last_col}{ROW_CLOSING_BAL}")
    _add_named_range(wb, "FinCons_LastCol", "Calc_Financing_Cons", f"{last_col}1")

    ws.freeze_panes = ws.cell(row=ROW_CLOSING_BAL + 1, column=FIRST_DATA_COL)
    ws.column_dimensions["A"].width = 52

    return ws


def _debt_draw_formula(col: str, prev: str | None) -> str:
    """Debt draw depends on the Cover drawdown-method selector.

    Debt First   — draw debt until the facility is exhausted, then equity picks up the rest.
    Equity First — draw equity until the commitment is exhausted, then debt covers the rest.
    Pari Passu   — split every month's funding requirement at the gearing ratio.
    """
    cum = ROW_CUM_FUNDING_REQUIREMENT
    prev_cum_debt = f"MIN({prev}{cum},{ABS_DEBT_FACILITY})" if prev else "0"
    prev_cum_equity = f"MIN({prev}{cum},{ABS_EQUITY_COMMITMENT})" if prev else "0"

    debt_first = f"MIN({col}{cum},{ABS_DEBT_FACILITY})-{prev_cum_debt}"
    equity_first = f"{col}{ROW_FUNDING_REQUIREMENT}-(MIN({col}{cum},{ABS_EQUITY_COMMITMENT})-{prev_cum_equity})"
    # Split at the facility's share of the *total funding requirement*, not the raw
    # gearing input and not Total Project Cost.
    #
    # Not the gearing input: under DSCR Sculpted the facility is the solved debt size, so
    # using the assumption would ignore the solve and peg implied gearing to the input.
    #
    # Not Total Project Cost: once the requirement also covers the opening DSRA, a ratio
    # taken over the smaller TPC scales every draw up, and the draws sum to
    # facility x (1 + DSRA/TPC) — over-drawing the facility. Dividing by the same base the
    # draws are applied to is what makes them sum to the facility exactly.
    pari_passu = (
        f"{col}{ROW_FUNDING_REQUIREMENT}*"
        f"IFERROR({ABS_DEBT_FACILITY}/{CELL_TOTAL_FUNDING_REQ},0)"
    )

    return (
        f'=IF({ABS_DRAWDOWN_METHOD}="Debt First",{debt_first},'
        f'IF({ABS_DRAWDOWN_METHOD}="Equity First",{equity_first},{pari_passu}))'
    )


def _label(ws: Worksheet, row: int, label: str) -> None:
    ws.cell(row=row, column=1, value=label)


def _link(ws: Worksheet, col: str, row: int, formula: str) -> None:
    cell = ws[f"{col}{row}"]
    cell.value = f"={formula}"
    cell.font = Font(color=COLOR_LINK)
    cell.number_format = "#,##0"


def _formula(ws: Worksheet, col: str, row: int, formula: str) -> None:
    cell = ws[f"{col}{row}"]
    cell.value = formula
    cell.font = Font(color=COLOR_FORMULA)
    cell.number_format = "#,##0"


def _check(ws: Worksheet, col: str, row: int, formula: str) -> None:
    cell = ws[f"{col}{row}"]
    cell.value = formula
    cell.font = Font(color=COLOR_FORMULA)


def _add_named_range(wb: Workbook, name: str, sheet: str, cell: str) -> None:
    import re

    from openpyxl.workbook.defined_name import DefinedName

    match = re.match(r"([A-Za-z]+)(\d+)", cell)
    col, row = match.group(1), match.group(2)
    wb.defined_names.add(DefinedName(name, attr_text=f"'{sheet}'!${col}${row}"))
