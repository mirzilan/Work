from openpyxl.styles import Font
from openpyxl.worksheet.worksheet import Worksheet
from openpyxl.workbook import Workbook

from inputs import ProjectInputs
from timeline import Timeline
from workbook_builder import (
    FIRST_DATA_COL,
    COLOR_INPUT,
    COLOR_FORMULA,
    TAB_COLOR_CALC,
    col_letter,
)

ROW_DATE_HEADER = 2
ROW_QUARTER_INDEX = 3

ROW_REVENUE = 5
ROW_OPEX = 6
ROW_EBITDA = 7

ROW_CHECK_HEADER = 11
ROW_CHECK_YEAR1_REVENUE = 12

ANNUAL_REVENUE_CELL = "B4"
OPEX_PCT_CELL = "B9"


def build_calc_revenue_opex(wb: Workbook, timeline: Timeline, inputs: ProjectInputs) -> Worksheet:
    ws = wb.create_sheet("Calc_Revenue_Opex")
    ws.sheet_properties.tabColor = TAB_COLOR_CALC

    ws["A1"] = "Calc_Revenue_Opex — Quarterly Revenue & Opex (flat dummy, Stage 1a)"
    ws["A1"].font = Font(bold=True, size=12)

    ws["A4"] = "Annual Revenue Input ($) — flat dummy placeholder"
    ws[ANNUAL_REVENUE_CELL] = inputs.revenue_opex.annual_revenue
    ws[ANNUAL_REVENUE_CELL].font = Font(color=COLOR_INPUT)
    ws[ANNUAL_REVENUE_CELL].number_format = "#,##0"

    ws["A9"] = "Opex % of Revenue"
    ws[OPEX_PCT_CELL] = inputs.revenue_opex.opex_pct_of_revenue
    ws[OPEX_PCT_CELL].font = Font(color=COLOR_INPUT)
    ws[OPEX_PCT_CELL].number_format = "0.00%"

    _label(ws, ROW_DATE_HEADER, "Period End Date")
    _label(ws, ROW_QUARTER_INDEX, "Operating Quarter #")
    _label(ws, ROW_REVENUE, "Revenue ($)")
    _label(ws, ROW_OPEX, "Opex ($)")
    _label(ws, ROW_EBITDA, "EBITDA ($)")

    ws.cell(row=ROW_CHECK_HEADER, column=1, value="Checks").font = Font(bold=True)
    _label(ws, ROW_CHECK_YEAR1_REVENUE, "Check: Sum of first 4 quarters' revenue = Annual Revenue Input")

    n_quarters = len(timeline.operations_quarters)

    for i, period in enumerate(timeline.operations_quarters):
        col = col_letter(i)

        ws[f"{col}{ROW_DATE_HEADER}"] = period.end
        ws[f"{col}{ROW_DATE_HEADER}"].number_format = "mmm-yy"
        ws[f"{col}{ROW_QUARTER_INDEX}"] = i + 1

        # Flat dummy revenue: annual input / 4, black formula (no escalation yet — Stage 1c)
        rev_cell = ws[f"{col}{ROW_REVENUE}"]
        rev_cell.value = f"=$B$4/4"
        rev_cell.font = Font(color=COLOR_FORMULA)
        rev_cell.number_format = "#,##0"

        opex_cell = ws[f"{col}{ROW_OPEX}"]
        opex_cell.value = f"={col}{ROW_REVENUE}*$B$9"
        opex_cell.font = Font(color=COLOR_FORMULA)
        opex_cell.number_format = "#,##0"

        ebitda_cell = ws[f"{col}{ROW_EBITDA}"]
        ebitda_cell.value = f"={col}{ROW_REVENUE}-{col}{ROW_OPEX}"
        ebitda_cell.font = Font(color=COLOR_FORMULA)
        ebitda_cell.number_format = "#,##0"

    # Check: first 4 quarters' revenue sums to the annual input
    q1, q2, q3, q4 = (col_letter(i) for i in range(4))
    check_cell = ws[f"{q4}{ROW_CHECK_YEAR1_REVENUE}"]
    check_cell.value = (
        f"=IF(ROUND(SUM({q1}{ROW_REVENUE}:{q4}{ROW_REVENUE})-$B$4,2)=0,1,0)"
    )
    check_cell.font = Font(color=COLOR_FORMULA)

    last_col = col_letter(n_quarters - 1)

    _add_named_range(wb, "RevOpex_LastCol", "Calc_Revenue_Opex", f"{last_col}1")
    _add_named_range(wb, "RevOpex_Year1Check", "Calc_Revenue_Opex", f"{q4}{ROW_CHECK_YEAR1_REVENUE}")

    ws.freeze_panes = ws.cell(row=ROW_EBITDA + 1, column=FIRST_DATA_COL)

    return ws


def _label(ws: Worksheet, row: int, label: str) -> None:
    ws.cell(row=row, column=1, value=label)


def _add_named_range(wb: Workbook, name: str, sheet: str, cell: str) -> None:
    import re

    from openpyxl.workbook.defined_name import DefinedName

    match = re.match(r"([A-Za-z]+)(\d+)", cell)
    col, row = match.group(1), match.group(2)
    wb.defined_names.add(DefinedName(name, attr_text=f"'{sheet}'!${col}${row}"))
