from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.styles import Font, PatternFill
from openpyxl.formatting.rule import FormulaRule
from openpyxl.worksheet.worksheet import Worksheet
from openpyxl.workbook import Workbook

import assumptions_constant as const
import batch_results as batch
import calc_capex as capex
import calc_financing_cons as fin_cons
import check_control
import fs_annual as fsa
from timeline import Timeline
from workbook_builder import COLOR_LINK, TAB_COLOR_OUTPUT, col_letter

# Sheet 1 — read-only, links only. Nothing here is computed; it just surfaces what
# Check_Control, Cover and Batch_Results already hold, laid out for a one-glance read.

ROW_STATUS_HEADER = 3
ROW_MODEL_STATUS = 4
ROW_SOLVE_FRESHNESS = 5
ROW_ACTIVE_SCENARIO = 6

ROW_RETURNS_HEADER = 9
ROW_TARGET_EIRR = 10
ROW_LIVE_EIRR = 11
ROW_EIRR_VS_TARGET = 12
ROW_TARGET_PIRR = 13
ROW_LIVE_PIRR = 14
ROW_PIRR_VS_TARGET = 15

ROW_SCENARIO_HEADER = 18
ROW_SCENARIO_TABLE_HEADER = 19
ROW_FIRST_SCENARIO = 20

ROW_SOURCES_USES_HEADER = 32
ROW_SU_TABLE_HEADER = 33
ROW_SU_CAPEX = 34
ROW_SU_IDC = 35
ROW_SU_DSRA = 36
ROW_SU_BUFFER = 37
ROW_SU_TOTAL = 38

ROW_CHART_HEADER = 41
ROW_CHART_ANCHOR = 42  # charts float below this row; ~15 rows tall each

FILL_GREEN = PatternFill(start_color="FFC6EFCE", end_color="FFC6EFCE", fill_type="solid")
FILL_RED = PatternFill(start_color="FFFFC7CE", end_color="FFFFC7CE", fill_type="solid")
FILL_AMBER = PatternFill(start_color="FFFFEB9C", end_color="FFFFEB9C", fill_type="solid")


def build_dashboard(wb: Workbook, timeline: Timeline) -> Worksheet:
    ws = wb.create_sheet("Dashboard", 0)  # built at index 0; Legend is inserted ahead of it afterward
    ws.sheet_properties.tabColor = TAB_COLOR_OUTPUT

    ws["A1"] = "Project 123 — Dashboard"
    ws["A1"].font = Font(bold=True, size=14)

    ws.cell(row=ROW_STATUS_HEADER, column=1, value="Status").font = Font(bold=True)

    _linked(ws, ROW_MODEL_STATUS, "Model Status", f"=Check_Control!B{check_control.ROW_MASTER_FLAG}")
    _linked(ws, ROW_SOLVE_FRESHNESS, "Solve Freshness", "=SolveStatus")
    _linked(ws, ROW_ACTIVE_SCENARIO, "Active Scenario",
            f"=ActiveScenario&\" - \"&Assumptions_Constant!$B${const.ROW_SCENARIO_NAME}")

    ws.conditional_formatting.add(
        f"B{ROW_MODEL_STATUS}",
        FormulaRule(formula=[f'B{ROW_MODEL_STATUS}="MODEL OK"'], fill=FILL_GREEN),
    )
    ws.conditional_formatting.add(
        f"B{ROW_MODEL_STATUS}",
        FormulaRule(formula=[f'B{ROW_MODEL_STATUS}="ERRORS FOUND"'], fill=FILL_RED),
    )
    ws.conditional_formatting.add(
        f"B{ROW_SOLVE_FRESHNESS}",
        FormulaRule(formula=[f'B{ROW_SOLVE_FRESHNESS}="SOLVED - current"'], fill=FILL_GREEN),
    )
    ws.conditional_formatting.add(
        f"B{ROW_SOLVE_FRESHNESS}",
        FormulaRule(formula=[f'B{ROW_SOLVE_FRESHNESS}<>"SOLVED - current"'], fill=FILL_AMBER),
    )

    ws.cell(row=ROW_RETURNS_HEADER, column=1, value="Headline Returns vs Target").font = Font(bold=True)

    _linked(ws, ROW_TARGET_EIRR, "Target EIRR", "=GoalSeek_TargetEIRR", "0.00%")
    _linked(ws, ROW_LIVE_EIRR, "Current EIRR", "=Live_EIRR", "0.00%", bold=True)
    _computed(ws, ROW_EIRR_VS_TARGET, "EIRR vs Target",
              f"=B{ROW_LIVE_EIRR}-B{ROW_TARGET_EIRR}", "0.00%")

    _linked(ws, ROW_TARGET_PIRR, "Target PIRR", "=GoalSeek_TargetPIRR", "0.00%")
    _linked(ws, ROW_LIVE_PIRR, "Current PIRR", "=Live_PIRR", "0.00%", bold=True)
    _computed(ws, ROW_PIRR_VS_TARGET, "PIRR vs Target",
              f"=B{ROW_LIVE_PIRR}-B{ROW_TARGET_PIRR}", "0.00%")

    for row in (ROW_EIRR_VS_TARGET, ROW_PIRR_VS_TARGET):
        ws.conditional_formatting.add(
            f"B{row}", FormulaRule(formula=[f"B{row}>=0"], fill=FILL_GREEN),
        )
        ws.conditional_formatting.add(
            f"B{row}", FormulaRule(formula=[f"B{row}<0"], fill=FILL_RED),
        )

    _build_scenario_comparison(ws)
    _build_sources_uses(ws, timeline)
    _build_charts(ws, timeline)

    ws.column_dimensions["A"].width = 24
    ws.column_dimensions["B"].width = 16

    # Status/Returns flags stay visible below this no matter how far the user scrolls
    # into the scenario table, Sources & Uses or charts.
    ws.freeze_panes = "A7"

    return ws


def _build_sources_uses(ws: Worksheet, timeline: Timeline) -> None:
    """Formula-linked, not chart-derived, per the blueprint spec: every cell here reads
    an existing Calc_Capex/Calc_Financing_Cons cell, never recomputes it. Sources and
    Uses are laid out side by side so the tie-out (Check_Financing_Cons's own
    "Cum Debt + Cum Equity = Cum Funding Requirement" check) is visually obvious."""
    last_cons_col = col_letter(len(timeline.construction_months) - 1)

    ws.cell(row=ROW_SOURCES_USES_HEADER, column=1, value="Sources & Uses").font = Font(bold=True)

    for col, header in ((1, "Uses"), (4, "Sources")):
        cell = ws.cell(row=ROW_SU_TABLE_HEADER, column=col, value=header)
        cell.font = Font(bold=True)

    uses = (
        (ROW_SU_CAPEX, "Total Capex", f"=Calc_Capex!${last_cons_col}${capex.ROW_CUM_CAPEX_DRAW}"),
        (ROW_SU_IDC, "IDC Capitalised", f"=Calc_Capex!${last_cons_col}${capex.ROW_CUM_IDC}"),
        (ROW_SU_DSRA, "Initial DSRA Funded at Close", f"=Calc_Financing_Cons!{fin_cons.CELL_INITIAL_DSRA}"),
        (ROW_SU_BUFFER, "Initial Cash Buffer Funded at Close",
         f"=Calc_Financing_Cons!{fin_cons.CELL_INITIAL_BUFFER}"),
    )
    for row, label, formula in uses:
        ws.cell(row=row, column=1, value=label)
        cell = ws.cell(row=row, column=2, value=formula)
        cell.font = Font(color=COLOR_LINK)
        cell.number_format = "#,##0"

    ws.cell(row=ROW_SU_TOTAL, column=1, value="Total Uses").font = Font(bold=True)
    total_uses = ws.cell(row=ROW_SU_TOTAL, column=2,
                         value=f"=Calc_Financing_Cons!{fin_cons.CELL_TOTAL_FUNDING_REQ}")
    total_uses.font = Font(color=COLOR_LINK, bold=True)
    total_uses.number_format = "#,##0"

    sources = (
        (ROW_SU_CAPEX, "Debt Facility", f"=Calc_Financing_Cons!{fin_cons.CELL_DEBT_FACILITY}"),
        (ROW_SU_IDC, "Equity Commitment", f"=Calc_Financing_Cons!{fin_cons.CELL_EQUITY_COMMITMENT}"),
    )
    for row, label, formula in sources:
        ws.cell(row=row, column=4, value=label)
        cell = ws.cell(row=row, column=5, value=formula)
        cell.font = Font(color=COLOR_LINK)
        cell.number_format = "#,##0"

    ws.cell(row=ROW_SU_TOTAL, column=4, value="Total Sources").font = Font(bold=True)
    total_sources = ws.cell(row=ROW_SU_TOTAL, column=5, value=f"=E{ROW_SU_CAPEX}+E{ROW_SU_IDC}")
    total_sources.font = Font(color=COLOR_LINK, bold=True)
    total_sources.number_format = "#,##0"

    ws.column_dimensions["D"].width = 24
    ws.column_dimensions["E"].width = 16


def _build_charts(ws: Worksheet, timeline: Timeline) -> None:
    """Native, formula-driven charts (openpyxl BarChart/LineChart) reading live cell
    ranges -- not images, not chart-derived summary numbers duplicated elsewhere."""
    ws.cell(row=ROW_CHART_HEADER, column=1, value="Charts").font = Font(bold=True)

    returns_chart = BarChart()
    returns_chart.title = "EIRR / PIRR by Scenario"
    returns_chart.y_axis.numFmt = "0%"
    returns_chart.height = 8
    returns_chart.width = 16
    last_scenario_row = ROW_FIRST_SCENARIO + const.N_SCENARIOS - 1
    data = Reference(ws, min_col=3, max_col=4, min_row=ROW_SCENARIO_TABLE_HEADER,
                     max_row=last_scenario_row)
    cats = Reference(ws, min_col=2, min_row=ROW_FIRST_SCENARIO, max_row=last_scenario_row)
    returns_chart.add_data(data, titles_from_data=True)
    returns_chart.set_categories(cats)
    ws.add_chart(returns_chart, f"A{ROW_CHART_ANCHOR}")

    profile_chart = LineChart()
    profile_chart.title = "Annual Revenue / EBITDA / Net Income (Operating Years)"
    profile_chart.y_axis.numFmt = "#,##0"
    profile_chart.height = 8
    profile_chart.width = 16
    fsa_ws = ws.parent["FS_Annual"]
    # FS_Annual reports operating years only (construction is a separate annual block) --
    # same bucket count fs_annual.py itself used to lay the columns out.
    n_annual_years = len(fsa._operations_only_annual_buckets(timeline))
    last_col_idx = 3 + n_annual_years - 1  # FIRST_DATA_COL (col C = 3) through the final year
    # Revenue/EBITDA/Net Income are no longer contiguous rows on FS_Annual now that the
    # full P&L sits between them, so each series is added individually by row rather than
    # relying on a single min_row:max_row block spanning rows that happen to be adjacent.
    for row in (fsa.ROW_REVENUE, fsa.ROW_EBITDA, fsa.ROW_NET_INCOME):
        series_ref = Reference(fsa_ws, min_col=1, max_col=last_col_idx, min_row=row, max_row=row)
        profile_chart.add_data(series_ref, titles_from_data=True, from_rows=True)
    cash_cats = Reference(fsa_ws, min_col=3, max_col=last_col_idx,
                          min_row=fsa.ROW_YEAR_LABEL, max_row=fsa.ROW_YEAR_LABEL)
    profile_chart.set_categories(cash_cats)
    ws.add_chart(profile_chart, f"J{ROW_CHART_ANCHOR}")


def _build_scenario_comparison(ws: Worksheet) -> None:
    """Link-only view of Batch_Results — never re-derives a scenario's numbers, since
    Batch_Results itself only holds values written by the last batch run, not live formulas."""
    ws.cell(row=ROW_SCENARIO_HEADER, column=1,
            value="Scenario Comparison (from last Batch Run)").font = Font(bold=True)

    headers = ["Scenario", "Name", "EIRR", "PIRR", "Min DSCR", "Min LLCR", "Gearing"]
    batch_cols = [1, 2, 8, 9, 6, 7, 5]  # matching Batch_Results.COLUMNS positions
    fmts = [None, None, "0.00%", "0.00%", "0.0000", "0.000", "0.00%"]

    for i, header in enumerate(headers):
        cell = ws.cell(row=ROW_SCENARIO_TABLE_HEADER, column=1 + i, value=header)
        cell.font = Font(bold=True)

    for s in range(const.N_SCENARIOS):
        src_row = batch.ROW_FIRST_SCENARIO + s
        dst_row = ROW_FIRST_SCENARIO + s
        for i, (batch_col, fmt) in enumerate(zip(batch_cols, fmts)):
            src_letter = _col_letter(batch_col)
            cell = ws.cell(row=dst_row, column=1 + i,
                           value=f"=Batch_Results!{src_letter}{src_row}")
            cell.font = Font(color=COLOR_LINK)
            if fmt:
                cell.number_format = fmt

    ws.column_dimensions["C"].width = 12
    ws.column_dimensions["D"].width = 16

    # Scenario highlight (Phase 2): flag the best and worst EIRR across the 10 scenarios
    # so the comparison table reads at a glance, not just as a list of numbers.
    last_scenario_row = ROW_FIRST_SCENARIO + const.N_SCENARIOS - 1
    eirr_range = f"C{ROW_FIRST_SCENARIO}:C{last_scenario_row}"
    eirr_abs_range = f"$C${ROW_FIRST_SCENARIO}:$C${last_scenario_row}"
    ws.conditional_formatting.add(
        eirr_range,
        FormulaRule(formula=[f"C{ROW_FIRST_SCENARIO}=MAX({eirr_abs_range})"], fill=FILL_GREEN),
    )
    ws.conditional_formatting.add(
        eirr_range,
        FormulaRule(formula=[f"C{ROW_FIRST_SCENARIO}=MIN({eirr_abs_range})"], fill=FILL_RED),
    )


def _col_letter(n: int) -> str:
    import openpyxl.utils

    return openpyxl.utils.get_column_letter(n)


def _linked(ws: Worksheet, row: int, label: str, formula: str, fmt: str | None = None,
           bold: bool = False) -> None:
    ws.cell(row=row, column=1, value=label)
    cell = ws.cell(row=row, column=2, value=formula)
    cell.font = Font(color=COLOR_LINK, bold=bold)
    if fmt:
        cell.number_format = fmt


def _computed(ws: Worksheet, row: int, label: str, formula: str, fmt: str | None = None) -> None:
    ws.cell(row=row, column=1, value=label)
    cell = ws.cell(row=row, column=2, value=formula)
    cell.font = Font(bold=True)
    if fmt:
        cell.number_format = fmt
