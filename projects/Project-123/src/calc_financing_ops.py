from openpyxl.styles import Font
from openpyxl.worksheet.worksheet import Worksheet
from openpyxl.workbook import Workbook

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

ROW_DATE_HEADER = 2
ROW_QUARTER_INDEX = 3

ROW_OPENING_BAL = 5
ROW_INTEREST = 6
ROW_PRINCIPAL = 7
ROW_DEBT_SERVICE = 8
ROW_CLOSING_BAL = 9
ROW_DSCR = 10   # display-only in Stage 1a; not yet used to size debt (that's Stage 1b sculpting)

ROW_CHECK_HEADER = 14
ROW_CHECK_FULLY_AMORTIZED = 15

INTEREST_RATE_CELL = "B4"
TENOR_YEARS_CELL = "B9"


def build_calc_financing_ops(wb: Workbook, timeline: Timeline, inputs: ProjectInputs) -> Worksheet:
    ws = wb.create_sheet("Calc_Financing_Ops")
    ws.sheet_properties.tabColor = TAB_COLOR_CALC

    ws["A1"] = "Calc_Financing_Ops — Quarterly Debt Service (fixed amortization, Stage 1a)"
    ws["A1"].font = Font(bold=True, size=12)

    ws["A4"] = "Interest Rate (Annual)"
    ws[INTEREST_RATE_CELL] = inputs.financing.interest_rate_annual
    ws[INTEREST_RATE_CELL].font = Font(color=COLOR_INPUT)
    ws[INTEREST_RATE_CELL].number_format = "0.00%"

    ws["A9"] = "Debt Tenor (Years)"
    ws[TENOR_YEARS_CELL] = inputs.financing.debt_tenor_years
    ws[TENOR_YEARS_CELL].font = Font(color=COLOR_INPUT)

    _label(ws, ROW_DATE_HEADER, "Period End Date")
    _label(ws, ROW_QUARTER_INDEX, "Operating Quarter #")
    _label(ws, ROW_OPENING_BAL, "Opening Debt Balance ($)")
    _label(ws, ROW_INTEREST, "Interest ($)")
    _label(ws, ROW_PRINCIPAL, "Principal ($) — level amortization, not yet DSCR-sculpted")
    _label(ws, ROW_DEBT_SERVICE, "Total Debt Service ($)")
    _label(ws, ROW_CLOSING_BAL, "Closing Debt Balance ($)")
    _label(ws, ROW_DSCR, "DSCR (display only — CFADS / Debt Service)")

    ws.cell(row=ROW_CHECK_HEADER, column=1, value="Checks").font = Font(bold=True)
    _label(ws, ROW_CHECK_FULLY_AMORTIZED, "Check: Closing Balance = 0 at Debt Tenor End")

    n_quarters = len(timeline.operations_quarters)
    tenor_quarters = inputs.financing.debt_tenor_years * 4
    quarterly_rate_expr = f"${INTEREST_RATE_CELL[0]}${INTEREST_RATE_CELL[1:]}/4"

    # Level quarterly payment via PMT, computed once as a named formula reused each period while balance > 0
    payment_formula = (
        f"-PMT({quarterly_rate_expr},$B$9*4,Calc_Financing_Cons!"
        f"{_last_construction_col(timeline)}8)"
    )  # ROW_CLOSING_BAL in calc_financing_cons.py = 8

    for i, period in enumerate(timeline.operations_quarters):
        col = col_letter(i)

        ws[f"{col}{ROW_DATE_HEADER}"] = period.end
        ws[f"{col}{ROW_DATE_HEADER}"].number_format = "mmm-yy"
        ws[f"{col}{ROW_QUARTER_INDEX}"] = i + 1

        opening_cell = ws[f"{col}{ROW_OPENING_BAL}"]
        if i == 0:
            opening_cell.value = f"=Calc_Financing_Cons!{_last_construction_col(timeline)}8"
            opening_cell.font = Font(color=COLOR_LINK)
        else:
            prev_col = col_letter(i - 1)
            opening_cell.value = f"={prev_col}{ROW_CLOSING_BAL}"
            opening_cell.font = Font(color=COLOR_FORMULA)
        opening_cell.number_format = "#,##0"

        interest_cell = ws[f"{col}{ROW_INTEREST}"]
        interest_cell.value = f"={col}{ROW_OPENING_BAL}*{quarterly_rate_expr}"
        interest_cell.font = Font(color=COLOR_FORMULA)
        interest_cell.number_format = "#,##0"

        debt_service_cell = ws[f"{col}{ROW_DEBT_SERVICE}"]
        if i < tenor_quarters:
            debt_service_cell.value = (
                f"=MIN({payment_formula},{col}{ROW_OPENING_BAL}+{col}{ROW_INTEREST})"
            )
        else:
            debt_service_cell.value = 0
        debt_service_cell.font = Font(color=COLOR_FORMULA)
        debt_service_cell.number_format = "#,##0"

        principal_cell = ws[f"{col}{ROW_PRINCIPAL}"]
        principal_cell.value = f"={col}{ROW_DEBT_SERVICE}-{col}{ROW_INTEREST}"
        principal_cell.font = Font(color=COLOR_FORMULA)
        principal_cell.number_format = "#,##0"

        closing_cell = ws[f"{col}{ROW_CLOSING_BAL}"]
        closing_cell.value = f"={col}{ROW_OPENING_BAL}-{col}{ROW_PRINCIPAL}"
        closing_cell.font = Font(color=COLOR_FORMULA)
        closing_cell.number_format = "#,##0"

        dscr_cell = ws[f"{col}{ROW_DSCR}"]
        dscr_cell.value = (
            f"=IF({col}{ROW_DEBT_SERVICE}=0,\"\",Calc_CFADS!{col}5/{col}{ROW_DEBT_SERVICE})"
        )  # Calc_CFADS ROW_CFADS = 5, built next
        dscr_cell.font = Font(color=COLOR_LINK)
        dscr_cell.number_format = "0.00x"

    last_col = col_letter(n_quarters - 1)
    tenor_end_col = col_letter(min(tenor_quarters, n_quarters) - 1)

    check_cell = ws[f"{last_col}{ROW_CHECK_FULLY_AMORTIZED}"]
    check_cell.value = f"=IF(ROUND({tenor_end_col}{ROW_CLOSING_BAL},2)=0,1,0)"
    check_cell.font = Font(color=COLOR_FORMULA)

    _add_named_range(wb, "FinOps_LastCol", "Calc_Financing_Ops", f"{last_col}1")
    _add_named_range(wb, "FinOps_AmortizedCheck", "Calc_Financing_Ops", f"{last_col}{ROW_CHECK_FULLY_AMORTIZED}")
    _add_named_range(wb, "FinOps_TenorEndCol", "Calc_Financing_Ops", f"{tenor_end_col}1")

    ws.freeze_panes = ws.cell(row=ROW_DSCR + 1, column=FIRST_DATA_COL)

    return ws


def _last_construction_col(timeline: Timeline) -> str:
    return col_letter(len(timeline.construction_months) - 1)


def _label(ws: Worksheet, row: int, label: str) -> None:
    ws.cell(row=row, column=1, value=label)


def _add_named_range(wb: Workbook, name: str, sheet: str, cell: str) -> None:
    import re

    from openpyxl.workbook.defined_name import DefinedName

    match = re.match(r"([A-Za-z]+)(\d+)", cell)
    col, row = match.group(1), match.group(2)
    wb.defined_names.add(DefinedName(name, attr_text=f"'{sheet}'!${col}${row}"))
