from openpyxl.styles import Font
from openpyxl.worksheet.worksheet import Worksheet
from openpyxl.workbook import Workbook

from workbook_builder import TAB_COLOR_CHECK, COLOR_FORMULA

ROW_HEADER = 1
ROW_MASTER_FLAG = 3
ROW_TABLE_HEADER = 5
ROW_FIRST_CHECK = 6

# (Sheet, Description, Cell, check_type)
#   check_type "flag"  — cell is 1/0, 1 = pass -> OK/FAIL
#   check_type "count" — cell is a failure count, 0 = pass -> OK/FAIL
#   check_type "info"  — cell is a count that's informational only -> OK/REVIEW, never FAIL
CHECKS = [
    ("Cover", "Solve is current (assumptions unchanged since last solve)", "D40", "flag"),
    ("Calc_Capex", "Cum Debt + Cum Equity = Cum Capex + Cum IDC", "Z21", "flag"),
    ("Calc_Capex", "Cumulative Capex = Total Capex Input", "Z22", "flag"),
    ("Calc_Financing_Cons", "Closing Balance = Cumulative Debt Draws", "Z29", "flag"),
    ("Calc_Financing_Cons", "IDC Converged (Loop 1)", "Z30", "flag"),
    ("Calc_Financing_Cons", "Cum Debt + Cum Equity = Cum Funding Requirement", "Z31", "flag"),
    ("Calc_Financing_Cons", "Cumulative Debt Draw <= Debt Facility", "Z32", "flag"),
    ("Calc_Revenue_Opex", "Year 1 Revenue Sum = Annual Revenue Input", "F12", "flag"),
    ("Calc_Tax", "Accumulated Depreciation <= Total Project Cost", "CD15", "flag"),
    ("Calc_Financing_Ops", "Closing Balance = 0 at Debt Tenor End", "CD26", "flag"),
    ("Calc_Financing_Ops", "Debt Sizing Converged (Loop 2)", "CD27", "flag"),
    ("Calc_Financing_Ops", "Min DSCR over tenor >= Target DSCR", "CD28", "flag"),
    ("Calc_Financing_Ops", "Quarters where CFADS/DSCR < interest", "CD29", "info"),
    ("Calc_CFADS", "Negative FCFE Quarter Count", "CD12", "info"),
    ("FS_Quarterly", "BS Balance Failures (Assets != Liab+Equity)", "CD38", "count"),
]


def build_check_control(wb: Workbook) -> Worksheet:
    ws = wb.create_sheet("Check_Control")
    ws.sheet_properties.tabColor = TAB_COLOR_CHECK

    ws["A1"] = "Check_Control — Master Audit Aggregator (Stage 1a)"
    ws["A1"].font = Font(bold=True, size=12)

    ws.cell(row=ROW_MASTER_FLAG, column=1, value="MASTER STATUS")
    ws.cell(row=ROW_MASTER_FLAG, column=1).font = Font(bold=True)

    ws.cell(row=ROW_TABLE_HEADER, column=1, value="Sheet")
    ws.cell(row=ROW_TABLE_HEADER, column=2, value="Check Description")
    ws.cell(row=ROW_TABLE_HEADER, column=3, value="Source Cell")
    ws.cell(row=ROW_TABLE_HEADER, column=4, value="Value")
    ws.cell(row=ROW_TABLE_HEADER, column=5, value="Status")
    for c in range(1, 6):
        ws.cell(row=ROW_TABLE_HEADER, column=c).font = Font(bold=True)

    row = ROW_FIRST_CHECK
    status_cells = []
    for sheet, desc, cell, check_type in CHECKS:
        ws.cell(row=row, column=1, value=sheet)
        ws.cell(row=row, column=2, value=desc)
        ws.cell(row=row, column=3, value=f"'{sheet}'!{cell}")

        value_cell = ws.cell(row=row, column=4)
        value_cell.value = f"='{sheet}'!{cell}"
        value_cell.font = Font(color=COLOR_FORMULA)

        status_cell = ws.cell(row=row, column=5)
        if check_type == "flag":
            status_cell.value = f'=IF(D{row}=1,"OK","FAIL")'
        elif check_type == "count":
            status_cell.value = f'=IF(D{row}=0,"OK","FAIL")'
        else:  # info
            status_cell.value = f'=IF(D{row}=0,"OK","REVIEW")'
        status_cell.font = Font(color=COLOR_FORMULA)
        status_cells.append(f"E{row}")

        row += 1

    last_check_row = row - 1

    master_cell = ws.cell(row=ROW_MASTER_FLAG, column=2)
    master_cell.value = (
        f'=IF(COUNTIF(E{ROW_FIRST_CHECK}:E{last_check_row},"FAIL")=0,"MODEL OK","ERRORS FOUND")'
    )
    master_cell.font = Font(bold=True, size=14, color=COLOR_FORMULA)

    _add_named_range(wb, "CheckControl_MasterFlag", "Check_Control", "B3")

    ws.column_dimensions["B"].width = 45
    ws.column_dimensions["C"].width = 22

    return ws


def _add_named_range(wb: Workbook, name: str, sheet: str, cell: str) -> None:
    import re

    from openpyxl.workbook.defined_name import DefinedName

    match = re.match(r"([A-Za-z]+)(\d+)", cell)
    col, row = match.group(1), match.group(2)
    wb.defined_names.add(DefinedName(name, attr_text=f"'{sheet}'!${col}${row}"))
