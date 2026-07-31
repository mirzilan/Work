from openpyxl.styles import Font
from openpyxl.worksheet.worksheet import Worksheet
from openpyxl.workbook import Workbook

import calc_cfads as cfads
import fs_quarterly as fsq
from inputs import ProjectInputs
from timeline import Timeline
from workbook_builder import (
    FIRST_DATA_COL,
    COLOR_FORMULA,
    COLOR_LINK,
    TAB_COLOR_OUTPUT,
)
import openpyxl.utils

ROW_YEAR_LABEL = 2
ROW_REVENUE = 4
ROW_EBITDA = 5
ROW_NET_INCOME = 6
ROW_CASH_CLOSING = 7
ROW_DEBT_CLOSING = 8
ROW_TOTAL_EQUITY_CLOSING = 9

# XIRR helper block: one continuous row of dates + one row of project (unlevered) cash flow,
# one row of equity cash flow, spanning ALL periods (24 construction months + 80 ops quarters).
ROW_XIRR_DATE = 13
ROW_XIRR_PROJECT_CF = 14
ROW_XIRR_EQUITY_CF = 15

ROW_PIRR_LABEL = 18
ROW_PIRR_VALUE = 19
ROW_EIRR_LABEL = 20
ROW_EIRR_VALUE = 21

# Construction-period annual summary (monthly source data rolled to project years)
ROW_CONS_HEADER = 24
ROW_CONS_YEAR_LABEL = 25
ROW_CONS_CAPEX = 26
ROW_CONS_IDC = 27
ROW_CONS_DEBT_DRAWN = 28
ROW_CONS_EQUITY_DRAWN = 29
ROW_CONS_CUM_TPC = 30
ROW_CONS_CLOSING_DEBT = 31


def _annual_col_letter(i: int) -> str:
    return openpyxl.utils.get_column_letter(FIRST_DATA_COL + i)


def _xirr_col_letter(i: int) -> str:
    return openpyxl.utils.get_column_letter(FIRST_DATA_COL + i)


def build_fs_annual(wb: Workbook, timeline: Timeline, inputs: ProjectInputs) -> Worksheet:
    ws = wb.create_sheet("FS_Annual")
    ws.sheet_properties.tabColor = TAB_COLOR_OUTPUT

    ws["A1"] = "FS_Annual — Rolled up from FS_Quarterly; PIRR/EIRR via XIRR (Stage 1a)"
    ws["A1"].font = Font(bold=True, size=12)

    ws.cell(row=ROW_YEAR_LABEL, column=1, value="Project Year")
    ws.cell(row=ROW_REVENUE, column=1, value="Revenue ($) — annual sum")
    ws.cell(row=ROW_EBITDA, column=1, value="EBITDA ($) — annual sum")
    ws.cell(row=ROW_NET_INCOME, column=1, value="Net Income ($) — annual sum")
    ws.cell(row=ROW_CASH_CLOSING, column=1, value="Cash, Closing ($) — year-end")
    ws.cell(row=ROW_DEBT_CLOSING, column=1, value="Debt, Closing ($) — year-end")
    ws.cell(row=ROW_TOTAL_EQUITY_CLOSING, column=1, value="Total Equity, Closing ($) — year-end")

    annual_buckets = _operations_only_annual_buckets(timeline)

    for year_num, (year_index, quarters) in enumerate(annual_buckets.items()):
        col = _annual_col_letter(year_num)
        ws[f"{col}{ROW_YEAR_LABEL}"] = f"Yr {year_num + 1}"

        q_cols = [_quarterly_source_col(timeline, q) for q in quarters]
        first_q_col, last_q_col = q_cols[0], q_cols[-1]

        for row, src_row, is_sum in (
            (ROW_REVENUE, fsq.ROW_REVENUE, True),
            (ROW_EBITDA, fsq.ROW_EBITDA, True),
            (ROW_NET_INCOME, fsq.ROW_NET_INCOME, True),
            (ROW_CASH_CLOSING, fsq.ROW_BS_CASH, False),
            (ROW_DEBT_CLOSING, fsq.ROW_BS_DEBT, False),
            (ROW_TOTAL_EQUITY_CLOSING, fsq.ROW_BS_TOTAL_EQUITY, False),
        ):
            cell = ws[f"{col}{row}"]
            if is_sum:
                cell.value = f"=SUM(FS_Quarterly!{first_q_col}{src_row}:{last_q_col}{src_row})"
            else:
                cell.value = f"=FS_Quarterly!{last_q_col}{src_row}"
            cell.font = Font(color=COLOR_LINK)
            cell.number_format = "#,##0"

    _build_xirr_block(ws, wb, timeline)

    first_xirr_col = _xirr_col_letter(0)

    ws.cell(row=ROW_PIRR_LABEL, column=1, value="Project IRR (PIRR) — unlevered, XIRR")
    pirr_cell = ws.cell(row=ROW_PIRR_VALUE, column=1)
    n_periods = len(timeline.construction_months) + len(timeline.operations_quarters)
    last_xirr_col = _xirr_col_letter(n_periods - 1)
    pirr_cell.value = (
        f"=XIRR({first_xirr_col}{ROW_XIRR_PROJECT_CF}:{last_xirr_col}{ROW_XIRR_PROJECT_CF},"
        f"{first_xirr_col}{ROW_XIRR_DATE}:{last_xirr_col}{ROW_XIRR_DATE})"
    )
    pirr_cell.font = Font(color=COLOR_FORMULA, bold=True)
    pirr_cell.number_format = "0.00%"

    ws.cell(row=ROW_EIRR_LABEL, column=1, value="Equity IRR (EIRR) — levered, XIRR")
    eirr_cell = ws.cell(row=ROW_EIRR_VALUE, column=1)
    eirr_cell.value = (
        f"=XIRR({first_xirr_col}{ROW_XIRR_EQUITY_CF}:{last_xirr_col}{ROW_XIRR_EQUITY_CF},"
        f"{first_xirr_col}{ROW_XIRR_DATE}:{last_xirr_col}{ROW_XIRR_DATE})"
    )
    eirr_cell.font = Font(color=COLOR_FORMULA, bold=True)
    eirr_cell.number_format = "0.00%"

    _build_construction_section(ws, timeline)

    _add_named_range(wb, "FSA_PIRR", "FS_Annual", "A19")
    _add_named_range(wb, "FSA_EIRR", "FS_Annual", "A21")

    return ws


def _build_construction_section(ws: Worksheet, timeline: Timeline) -> None:
    """Construction years reported separately: the source data is monthly (not quarterly),
    and during construction there is no P&L activity — only capitalised spend building the
    balance sheet, which is what lenders look at pre-COD."""
    ws.cell(row=ROW_CONS_HEADER, column=1, value="Construction Period (annual rollup of monthly data)")
    ws.cell(row=ROW_CONS_HEADER, column=1).font = Font(bold=True)

    ws.cell(row=ROW_CONS_YEAR_LABEL, column=1, value="Construction Year")
    ws.cell(row=ROW_CONS_CAPEX, column=1, value="Capex Incurred ($) — annual sum")
    ws.cell(row=ROW_CONS_IDC, column=1, value="IDC Capitalised ($) — annual sum")
    ws.cell(row=ROW_CONS_DEBT_DRAWN, column=1, value="Debt Drawn ($) — annual sum")
    ws.cell(row=ROW_CONS_EQUITY_DRAWN, column=1, value="Equity Drawn ($) — annual sum")
    ws.cell(row=ROW_CONS_CUM_TPC, column=1, value="Cumulative Total Project Cost ($) — year-end")
    ws.cell(row=ROW_CONS_CLOSING_DEBT, column=1, value="Closing Debt Balance ($) — year-end")

    buckets: dict = {}
    for idx, period in enumerate(timeline.construction_months):
        buckets.setdefault(period.year_index, []).append(idx)

    for year_num, (_, month_indices) in enumerate(sorted(buckets.items())):
        col = _annual_col_letter(year_num)
        first_m = openpyxl.utils.get_column_letter(FIRST_DATA_COL + month_indices[0])
        last_m = openpyxl.utils.get_column_letter(FIRST_DATA_COL + month_indices[-1])

        ws[f"{col}{ROW_CONS_YEAR_LABEL}"] = f"C-Yr {year_num + 1}"

        for row, (sheet, src_row, is_sum) in {
            ROW_CONS_CAPEX: ("Calc_Capex", 6, True),
            ROW_CONS_IDC: ("Calc_Financing_Cons", 18, True),
            ROW_CONS_DEBT_DRAWN: ("Calc_Financing_Cons", 21, True),
            ROW_CONS_EQUITY_DRAWN: ("Calc_Financing_Cons", 22, True),
            ROW_CONS_CUM_TPC: ("Calc_Capex", 17, False),
            ROW_CONS_CLOSING_DEBT: ("Calc_Financing_Cons", 25, False),
        }.items():
            cell = ws[f"{col}{row}"]
            if is_sum:
                cell.value = f"=SUM({sheet}!{first_m}{src_row}:{last_m}{src_row})"
            else:
                cell.value = f"={sheet}!{last_m}{src_row}"
            cell.font = Font(color=COLOR_LINK)
            cell.number_format = "#,##0"


def _build_xirr_block(ws: Worksheet, wb: Workbook, timeline: Timeline) -> None:
    ws.cell(row=ROW_XIRR_DATE, column=1, value="XIRR Helper: Period End Date")
    ws.cell(row=ROW_XIRR_PROJECT_CF, column=1, value="XIRR Helper: Project CF (unlevered) — construction outflow, FCFF inflow")
    ws.cell(row=ROW_XIRR_EQUITY_CF, column=1, value="XIRR Helper: Equity CF (levered) — equity draw outflow, FCFE inflow")

    n_cons = len(timeline.construction_months)

    # Construction months: outflows
    for i, period in enumerate(timeline.construction_months):
        col = _xirr_col_letter(i)
        capex_col_in_source = openpyxl.utils.get_column_letter(3 + i)  # matches Calc_Capex's own column layout

        date_cell = ws[f"{col}{ROW_XIRR_DATE}"]
        date_cell.value = f"=Calc_Capex!{capex_col_in_source}2"
        date_cell.font = Font(color=COLOR_LINK)
        date_cell.number_format = "mmm-yy"

        proj_cf_cell = ws[f"{col}{ROW_XIRR_PROJECT_CF}"]
        proj_cf_cell.value = f"=-Calc_Capex!{capex_col_in_source}6"  # ROW_CAPEX_DRAW
        proj_cf_cell.font = Font(color=COLOR_LINK)
        proj_cf_cell.number_format = "#,##0"

        equity_cf_cell = ws[f"{col}{ROW_XIRR_EQUITY_CF}"]
        equity_cf_cell.value = f"=-Calc_Capex!{capex_col_in_source}11"  # ROW_EQUITY_DRAW
        equity_cf_cell.font = Font(color=COLOR_LINK)
        equity_cf_cell.number_format = "#,##0"

    # Operations quarters: inflows
    for i, period in enumerate(timeline.operations_quarters):
        col = _xirr_col_letter(n_cons + i)
        ops_col_in_source = openpyxl.utils.get_column_letter(3 + i)  # matches Calc_CFADS's own column layout

        date_cell = ws[f"{col}{ROW_XIRR_DATE}"]
        date_cell.value = f"=Calc_CFADS!{ops_col_in_source}2"
        date_cell.font = Font(color=COLOR_LINK)
        date_cell.number_format = "mmm-yy"

        # FCFF and FCFE are both built once on Calc_CFADS and only referenced here.
        proj_cf_cell = ws[f"{col}{ROW_XIRR_PROJECT_CF}"]
        proj_cf_cell.value = f"=Calc_CFADS!{ops_col_in_source}{cfads.ROW_FCFF}"
        proj_cf_cell.font = Font(color=COLOR_LINK)
        proj_cf_cell.number_format = "#,##0"

        equity_cf_cell = ws[f"{col}{ROW_XIRR_EQUITY_CF}"]
        equity_cf_cell.value = f"=Calc_CFADS!{ops_col_in_source}{cfads.ROW_FCFE}"
        equity_cf_cell.font = Font(color=COLOR_LINK)
        equity_cf_cell.number_format = "#,##0"


def _operations_only_annual_buckets(timeline: Timeline) -> dict:
    """FS_Annual reports operating years only (Stage 1a); construction-period BS
    will be added when Calc_Financing_Cons circularity (Stage 1b) makes construction
    P&L/BS activity meaningful to show."""
    buckets: dict = {}
    for p in timeline.operations_quarters:
        buckets.setdefault(p.year_index, []).append(p)
    return dict(sorted(buckets.items()))


def _quarterly_source_col(timeline: Timeline, period) -> str:
    idx = timeline.operations_quarters.index(period)
    return openpyxl.utils.get_column_letter(FIRST_DATA_COL + idx)


def _add_named_range(wb: Workbook, name: str, sheet: str, cell: str) -> None:
    import re

    from openpyxl.workbook.defined_name import DefinedName

    match = re.match(r"([A-Za-z]+)(\d+)", cell)
    col, row = match.group(1), match.group(2)
    wb.defined_names.add(DefinedName(name, attr_text=f"'{sheet}'!${col}${row}"))
