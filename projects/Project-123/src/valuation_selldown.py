from openpyxl.styles import Font
from openpyxl.worksheet.worksheet import Worksheet
from openpyxl.workbook import Workbook
import openpyxl.utils

import fs_annual as fsa
from timeline import Timeline
from workbook_builder import (
    FIRST_DATA_COL,
    COLOR_FORMULA,
    COLOR_INPUT,
    COLOR_LINK,
    TAB_COLOR_OUTPUT,
)

# Standalone, read-only downstream of FS_Annual — no FCFF/FCFE is rebuilt here, only
# read via FS_Annual's own whole-of-life XIRR helper row (built once on Calc_CFADS).

ROW_TARGET_RATE = 3
ROW_TABLE_HEADER = 5
ROW_EXIT_YEAR_LABEL = 6
ROW_EXIT_DATE = 7
ROW_SALE_PRICE = 8
ROW_DEBT_OUTSTANDING = 9
ROW_BUYER_EV = 10
ROW_SELLER_EIRR = 11
ROW_BUYER_PIRR = 12

ROW_CHECK_HEADER = 14
ROW_CHECK_FINAL_EXIT_ZERO_PRICE = 15

# Two full-life-width helper rows per exit year: the seller's equity CF with the sale
# price substituted at the exit column, and the buyer's project CF with the price +
# assumed debt substituted at the exit column. XIRR just reads these ranges — the
# substitution is what turns "whole-of-project" cash flow into "this exit year's deal".
ROW_HELPER_HEADER = 18
ROW_FIRST_HELPER = 19


def _annual_col_letter(i: int) -> str:
    return openpyxl.utils.get_column_letter(FIRST_DATA_COL + i)


def build_valuation_selldown(wb: Workbook, timeline: Timeline) -> Worksheet:
    ws = wb.create_sheet("Valuation_SellDown")
    ws.sheet_properties.tabColor = TAB_COLOR_OUTPUT

    ws["A1"] = ("Valuation_SellDown — exit-year IRR scan, read-only downstream of "
                "FS_Annual (Stage 1d)")
    ws["A1"].font = Font(bold=True, size=12)

    # Independent input, not a link to GoalSeek_TargetPIRR: that cell is *your* return
    # target, which goal-seek can move the model's own revenue to hit. Discounting the
    # sell-down at the same cell you're solving toward makes the buyer's implied PIRR
    # come back at ~target by construction, which defeats the point of the scan. Seeded
    # at Cover's own Target PIRR default (8%) so day one it reads sensibly, but this is a
    # deal-specific assumption (buyer's required return) that should be set independently.
    ws.cell(row=ROW_TARGET_RATE, column=1,
            value="Exit Discount Rate (buyer/seller required return)")
    rate_cell = ws.cell(row=ROW_TARGET_RATE, column=2, value=0.08)
    rate_cell.font = Font(color=COLOR_INPUT)
    rate_cell.number_format = "0.00%"

    for row, label in (
        (ROW_EXIT_YEAR_LABEL, "Exit Year"),
        (ROW_EXIT_DATE, "Exit Date"),
        (ROW_SALE_PRICE, "Implied Sale Price ($) — XNPV of remaining FCFF at Target PIRR"),
        (ROW_DEBT_OUTSTANDING, "Debt Outstanding at Exit ($)"),
        (ROW_BUYER_EV, "Buyer Enterprise Value ($) — Price + Debt Assumed"),
        (ROW_SELLER_EIRR, "Seller's Realized EIRR — levered, whole holding period"),
        (ROW_BUYER_PIRR, "Buyer's Implied PIRR — unlevered, remaining project life"),
    ):
        ws.cell(row=row, column=1, value=label)
    ws.cell(row=ROW_TABLE_HEADER, column=1, value="Per Exit Year").font = Font(bold=True)

    n_cons = len(timeline.construction_months)
    n_periods = n_cons + len(timeline.operations_quarters)
    first_life_col = _annual_col_letter(0)
    last_life_col = _annual_col_letter(n_periods - 1)
    fsa_date_row = fsa.ROW_XIRR_DATE
    fsa_project_row = fsa.ROW_XIRR_PROJECT_CF
    fsa_equity_row = fsa.ROW_XIRR_EQUITY_CF

    annual_buckets = fsa._operations_only_annual_buckets(timeline)

    ws.cell(row=ROW_HELPER_HEADER, column=1,
            value="XIRR helper — one modified equity CF row + one modified project CF row "
                  "per exit year, spanning the whole life").font = Font(bold=True, size=9)

    last_sale_price_col = None
    for year_num, (_year_index, quarters) in enumerate(annual_buckets.items()):
        col = _annual_col_letter(year_num)
        last_sale_price_col = col
        exit_q = quarters[-1]
        exit_col_index = n_cons + timeline.operations_quarters.index(exit_q)

        ws[f"{col}{ROW_EXIT_YEAR_LABEL}"] = f"Yr {year_num + 1}"

        exit_date_cell = ws[f"{col}{ROW_EXIT_DATE}"]
        exit_date_cell.value = f"=FS_Annual!{_annual_col_letter(exit_col_index)}{fsa_date_row}"
        exit_date_cell.font = Font(color=COLOR_LINK)
        exit_date_cell.number_format = "mmm-yy"

        price_cell = ws[f"{col}{ROW_SALE_PRICE}"]
        price_cell.value = (
            f"=SUMPRODUCT((FS_Annual!{first_life_col}{fsa_date_row}:{last_life_col}{fsa_date_row}"
            f">{col}${ROW_EXIT_DATE})*FS_Annual!{first_life_col}{fsa_project_row}:{last_life_col}{fsa_project_row}"
            f"/(1+$B${ROW_TARGET_RATE})^((FS_Annual!{first_life_col}{fsa_date_row}:{last_life_col}{fsa_date_row}"
            f"-{col}${ROW_EXIT_DATE})/365))"
        )
        price_cell.font = Font(color=COLOR_FORMULA)
        price_cell.number_format = "#,##0"

        # FS_Annual's Debt-Closing row is populated on year-index columns (0..n_years-1),
        # not the whole-of-life period-index columns the XIRR helper block above uses —
        # reading it at exit_col_index (a period index) was pulling a blank/wrong cell.
        debt_cell = ws[f"{col}{ROW_DEBT_OUTSTANDING}"]
        debt_cell.value = f"=FS_Annual!{_annual_col_letter(year_num)}{fsa.ROW_DEBT_CLOSING}"
        debt_cell.font = Font(color=COLOR_LINK)
        debt_cell.number_format = "#,##0"

        ev_cell = ws[f"{col}{ROW_BUYER_EV}"]
        ev_cell.value = f"={col}{ROW_SALE_PRICE}+{col}{ROW_DEBT_OUTSTANDING}"
        ev_cell.font = Font(color=COLOR_FORMULA)
        ev_cell.number_format = "#,##0"

        row_equity = ROW_FIRST_HELPER + year_num * 2
        row_project = row_equity + 1
        ws.cell(row=row_equity, column=1, value=f"Yr {year_num + 1} exit — Seller Equity CF")
        ws.cell(row=row_project, column=1, value=f"Yr {year_num + 1} exit — Buyer Project CF")

        for i in range(n_periods):
            period_col = _annual_col_letter(i)
            eq_cell = ws[f"{period_col}{row_equity}"]
            proj_cell = ws[f"{period_col}{row_project}"]
            if i < exit_col_index:
                eq_cell.value = f"=FS_Annual!{period_col}{fsa_equity_row}"
                proj_cell.value = 0
            elif i == exit_col_index:
                eq_cell.value = f"=FS_Annual!{period_col}{fsa_equity_row}+{col}{ROW_SALE_PRICE}"
                proj_cell.value = f"=-{col}{ROW_BUYER_EV}"
            else:
                eq_cell.value = 0
                proj_cell.value = f"=FS_Annual!{period_col}{fsa_project_row}"
            eq_cell.font = Font(color=COLOR_FORMULA, size=8)
            proj_cell.font = Font(color=COLOR_FORMULA, size=8)
            eq_cell.number_format = "#,##0"
            proj_cell.number_format = "#,##0"

        eirr_cell = ws[f"{col}{ROW_SELLER_EIRR}"]
        eirr_cell.value = (
            f"=XIRR({first_life_col}{row_equity}:{last_life_col}{row_equity},"
            f"FS_Annual!{first_life_col}{fsa_date_row}:FS_Annual!{last_life_col}{fsa_date_row})"
        )
        eirr_cell.font = Font(color=COLOR_FORMULA, bold=True)
        eirr_cell.number_format = "0.00%"

        pirr_cell = ws[f"{col}{ROW_BUYER_PIRR}"]
        pirr_cell.value = (
            f"=XIRR({first_life_col}{row_project}:{last_life_col}{row_project},"
            f"FS_Annual!{first_life_col}{fsa_date_row}:FS_Annual!{last_life_col}{fsa_date_row})"
        )
        pirr_cell.font = Font(color=COLOR_FORMULA, bold=True)
        pirr_cell.number_format = "0.00%"

    ws.cell(row=ROW_CHECK_HEADER, column=1, value="Checks").font = Font(bold=True)
    ws.cell(row=ROW_CHECK_FINAL_EXIT_ZERO_PRICE, column=1,
            value="Check: Sale price rounds to 0 at the final exit year (no future FCFF left to sell)")
    final_check = ws.cell(row=ROW_CHECK_FINAL_EXIT_ZERO_PRICE, column=2)
    final_check.value = f'=IF(ROUND({last_sale_price_col}{ROW_SALE_PRICE},0)=0,1,0)'
    final_check.font = Font(color=COLOR_FORMULA)

    ws.column_dimensions["A"].width = 60

    _add_named_range(wb, "SellDown_FinalExitZeroCheck", "Valuation_SellDown",
                     f"B{ROW_CHECK_FINAL_EXIT_ZERO_PRICE}")

    return ws


def _add_named_range(wb: Workbook, name: str, sheet: str, cell: str) -> None:
    import re

    from openpyxl.workbook.defined_name import DefinedName

    match = re.match(r"([A-Za-z]+)(\d+)", cell)
    col, row = match.group(1), match.group(2)
    wb.defined_names.add(DefinedName(name, attr_text=f"'{sheet}'!${col}${row}"))
