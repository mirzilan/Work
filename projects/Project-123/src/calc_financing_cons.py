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
ROW_MONTH_INDEX = 3

ROW_DEBT_DRAW = 5       # linked from Calc_Capex
ROW_OPENING_BAL = 6
ROW_INTEREST_ACCRUED = 7   # placeholder: simple accrual on opening balance, not yet capitalized (Stage 1b)
ROW_CLOSING_BAL = 8

ROW_CHECK_HEADER = 12
ROW_CHECK_CLOSING_MATCHES_DRAWS = 13

INTEREST_RATE_CELL = "B4"


def build_calc_financing_cons(wb: Workbook, timeline: Timeline, inputs: ProjectInputs) -> Worksheet:
    ws = wb.create_sheet("Calc_Financing_Cons")
    ws.sheet_properties.tabColor = TAB_COLOR_CALC

    ws["A1"] = "Calc_Financing_Cons — Monthly Construction Debt Draws"
    ws["A1"].font = Font(bold=True, size=12)

    ws["A4"] = "Interest Rate (Annual)"
    ws[INTEREST_RATE_CELL] = inputs.financing.interest_rate_annual
    ws[INTEREST_RATE_CELL].font = Font(color=COLOR_INPUT)
    ws[INTEREST_RATE_CELL].number_format = "0.00%"

    _label(ws, ROW_DATE_HEADER, "Period End Date")
    _label(ws, ROW_MONTH_INDEX, "Construction Month #")
    _label(ws, ROW_DEBT_DRAW, "Debt Draw ($) — linked from Calc_Capex")
    _label(ws, ROW_OPENING_BAL, "Opening Debt Balance ($)")
    _label(ws, ROW_INTEREST_ACCRUED, "Interest Accrued ($) — placeholder, capitalization wired in Stage 1b")
    _label(ws, ROW_CLOSING_BAL, "Closing Debt Balance ($)")

    ws.cell(row=ROW_CHECK_HEADER, column=1, value="Checks").font = Font(bold=True)
    _label(ws, ROW_CHECK_CLOSING_MATCHES_DRAWS, "Check: Closing Balance = Cumulative Debt Draws")

    n_months = len(timeline.construction_months)
    monthly_rate_expr = f"${INTEREST_RATE_CELL[0]}${INTEREST_RATE_CELL[1:]}/12"

    for i, period in enumerate(timeline.construction_months):
        col = col_letter(i)

        ws[f"{col}{ROW_DATE_HEADER}"] = period.end
        ws[f"{col}{ROW_DATE_HEADER}"].number_format = "mmm-yy"
        ws[f"{col}{ROW_MONTH_INDEX}"] = i + 1

        # Debt draw — green link from Calc_Capex (same monthly column position)
        draw_cell = ws[f"{col}{ROW_DEBT_DRAW}"]
        draw_cell.value = f"=Calc_Capex!{col}10"  # ROW_DEBT_DRAW in calc_capex.py
        draw_cell.font = Font(color=COLOR_LINK)
        draw_cell.number_format = "#,##0"

        opening_cell = ws[f"{col}{ROW_OPENING_BAL}"]
        if i == 0:
            opening_cell.value = 0
            opening_cell.font = Font(color=COLOR_INPUT)
        else:
            prev_col = col_letter(i - 1)
            opening_cell.value = f"={prev_col}{ROW_CLOSING_BAL}"
            opening_cell.font = Font(color=COLOR_FORMULA)
        opening_cell.number_format = "#,##0"

        # Interest accrued — simple placeholder, NOT yet capitalized into draws (Stage 1b IDC solve)
        interest_cell = ws[f"{col}{ROW_INTEREST_ACCRUED}"]
        interest_cell.value = f"={col}{ROW_OPENING_BAL}*{monthly_rate_expr}"
        interest_cell.font = Font(color=COLOR_FORMULA)
        interest_cell.number_format = "#,##0"

        closing_cell = ws[f"{col}{ROW_CLOSING_BAL}"]
        closing_cell.value = f"={col}{ROW_OPENING_BAL}+{col}{ROW_DEBT_DRAW}"
        closing_cell.font = Font(color=COLOR_FORMULA)
        closing_cell.number_format = "#,##0"

    last_col = col_letter(n_months - 1)

    check_cell = ws[f"{last_col}{ROW_CHECK_CLOSING_MATCHES_DRAWS}"]
    check_cell.value = (
        f"=IF(ROUND({last_col}{ROW_CLOSING_BAL}-Calc_Capex!{last_col}12,2)=0,1,0)"
    )  # Calc_Capex ROW_CUM_DEBT_DRAW = 12
    check_cell.font = Font(color=COLOR_FORMULA)

    _add_named_range(wb, "FinCons_ClosingBalance_Last", "Calc_Financing_Cons", f"{last_col}{ROW_CLOSING_BAL}")
    _add_named_range(wb, "FinCons_ClosingCheck", "Calc_Financing_Cons", f"{last_col}{ROW_CHECK_CLOSING_MATCHES_DRAWS}")
    _add_named_range(wb, "FinCons_LastCol", "Calc_Financing_Cons", f"{last_col}1")

    ws.freeze_panes = ws.cell(row=ROW_CLOSING_BAL + 1, column=FIRST_DATA_COL)

    return ws


def _label(ws: Worksheet, row: int, label: str) -> None:
    ws.cell(row=row, column=1, value=label)


def _add_named_range(wb: Workbook, name: str, sheet: str, cell: str) -> None:
    import re

    from openpyxl.workbook.defined_name import DefinedName

    match = re.match(r"([A-Za-z]+)(\d+)", cell)
    col, row = match.group(1), match.group(2)
    wb.defined_names.add(DefinedName(name, attr_text=f"'{sheet}'!${col}${row}"))
