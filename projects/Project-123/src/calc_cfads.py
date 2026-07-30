from openpyxl.styles import Font
from openpyxl.worksheet.worksheet import Worksheet
from openpyxl.workbook import Workbook

from inputs import ProjectInputs
from timeline import Timeline
from workbook_builder import (
    FIRST_DATA_COL,
    COLOR_FORMULA,
    COLOR_LINK,
    TAB_COLOR_CALC,
    col_letter,
)

ROW_DATE_HEADER = 2
ROW_QUARTER_INDEX = 3

ROW_CFADS = 5           # EBITDA - Tax (pre-debt-service cash available)
ROW_DEBT_SERVICE = 6    # linked from Calc_Financing_Ops
ROW_FCFE = 7            # CFADS - Debt Service (no DSRA/MRA yet — Stage 1c)

ROW_CHECK_HEADER = 11
ROW_CHECK_FCFE_NOT_BELOW_ZERO_COUNT = 12  # informational, not a hard fail in Stage 1a


def build_calc_cfads(wb: Workbook, timeline: Timeline, inputs: ProjectInputs) -> Worksheet:
    ws = wb.create_sheet("Calc_CFADS")
    ws.sheet_properties.tabColor = TAB_COLOR_CALC

    ws["A1"] = "Calc_CFADS — Quarterly Cash Waterfall (no DSRA/MRA yet, Stage 1a)"
    ws["A1"].font = Font(bold=True, size=12)

    _label(ws, ROW_DATE_HEADER, "Period End Date")
    _label(ws, ROW_QUARTER_INDEX, "Operating Quarter #")
    _label(ws, ROW_CFADS, "CFADS ($) = EBITDA - Tax")
    _label(ws, ROW_DEBT_SERVICE, "Debt Service ($) — linked from Calc_Financing_Ops")
    _label(ws, ROW_FCFE, "FCFE ($) = CFADS - Debt Service")

    ws.cell(row=ROW_CHECK_HEADER, column=1, value="Checks").font = Font(bold=True)
    _label(ws, ROW_CHECK_FCFE_NOT_BELOW_ZERO_COUNT, "Informational: # of quarters with negative FCFE")

    n_quarters = len(timeline.operations_quarters)

    for i, period in enumerate(timeline.operations_quarters):
        col = col_letter(i)

        ws[f"{col}{ROW_DATE_HEADER}"] = period.end
        ws[f"{col}{ROW_DATE_HEADER}"].number_format = "mmm-yy"
        ws[f"{col}{ROW_QUARTER_INDEX}"] = i + 1

        # CFADS = EBITDA - Tax (both linked from Calc_Tax)
        cfads_cell = ws[f"{col}{ROW_CFADS}"]
        cfads_cell.value = f"=Calc_Tax!{col}5-Calc_Tax!{col}9"  # ROW_EBITDA=5, ROW_TAX=9 in calc_tax.py
        cfads_cell.font = Font(color=COLOR_LINK)
        cfads_cell.number_format = "#,##0"

        debt_service_cell = ws[f"{col}{ROW_DEBT_SERVICE}"]
        debt_service_cell.value = f"=Calc_Financing_Ops!{col}19"  # ROW_DEBT_SERVICE in calc_financing_ops.py
        debt_service_cell.font = Font(color=COLOR_LINK)
        debt_service_cell.number_format = "#,##0"

        fcfe_cell = ws[f"{col}{ROW_FCFE}"]
        fcfe_cell.value = f"={col}{ROW_CFADS}-{col}{ROW_DEBT_SERVICE}"
        fcfe_cell.font = Font(color=COLOR_FORMULA)
        fcfe_cell.number_format = "#,##0"

    last_col = col_letter(n_quarters - 1)
    first_col = col_letter(0)

    check_cell = ws[f"{last_col}{ROW_CHECK_FCFE_NOT_BELOW_ZERO_COUNT}"]
    check_cell.value = f"=COUNTIF({first_col}{ROW_FCFE}:{last_col}{ROW_FCFE},\"<0\")"
    check_cell.font = Font(color=COLOR_FORMULA)

    _add_named_range(wb, "CFADS_LastCol", "Calc_CFADS", f"{last_col}1")
    _add_named_range(wb, "CFADS_NegFCFECount", "Calc_CFADS", f"{last_col}{ROW_CHECK_FCFE_NOT_BELOW_ZERO_COUNT}")

    ws.freeze_panes = ws.cell(row=ROW_FCFE + 1, column=FIRST_DATA_COL)

    return ws


def _label(ws: Worksheet, row: int, label: str) -> None:
    ws.cell(row=row, column=1, value=label)


def _add_named_range(wb: Workbook, name: str, sheet: str, cell: str) -> None:
    import re

    from openpyxl.workbook.defined_name import DefinedName

    match = re.match(r"([A-Za-z]+)(\d+)", cell)
    col, row = match.group(1), match.group(2)
    wb.defined_names.add(DefinedName(name, attr_text=f"'{sheet}'!${col}${row}"))
