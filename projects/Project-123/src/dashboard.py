from openpyxl.styles import Font, PatternFill
from openpyxl.formatting.rule import FormulaRule
from openpyxl.worksheet.worksheet import Worksheet
from openpyxl.workbook import Workbook

import assumptions_constant as const
import batch_results as batch
from workbook_builder import COLOR_LINK, TAB_COLOR_OUTPUT

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

FILL_GREEN = PatternFill(start_color="FFC6EFCE", end_color="FFC6EFCE", fill_type="solid")
FILL_RED = PatternFill(start_color="FFFFC7CE", end_color="FFFFC7CE", fill_type="solid")
FILL_AMBER = PatternFill(start_color="FFFFEB9C", end_color="FFFFEB9C", fill_type="solid")


def build_dashboard(wb: Workbook) -> Worksheet:
    ws = wb.create_sheet("Dashboard", 0)  # index 0: opens here
    ws.sheet_properties.tabColor = TAB_COLOR_OUTPUT

    ws["A1"] = "Project 123 — Dashboard"
    ws["A1"].font = Font(bold=True, size=14)

    ws.cell(row=ROW_STATUS_HEADER, column=1, value="Status").font = Font(bold=True)

    _linked(ws, ROW_MODEL_STATUS, "Model Status", "=Check_Control!B3")
    _linked(ws, ROW_SOLVE_FRESHNESS, "Solve Freshness", "=SolveStatus")
    _linked(ws, ROW_ACTIVE_SCENARIO, "Active Scenario",
            "=ActiveScenario&\" - \"&Assumptions_Constant!$B$4")

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

    ws.column_dimensions["A"].width = 24
    ws.column_dimensions["B"].width = 16

    return ws


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
