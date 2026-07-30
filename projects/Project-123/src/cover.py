from openpyxl.styles import Font
from openpyxl.worksheet.worksheet import Worksheet
from openpyxl.workbook import Workbook

from workbook_builder import COLOR_INPUT, TAB_COLOR_INPUT

CELL_CIRC_TOLERANCE = "B4"
CELL_MAX_ITERATIONS = "B5"
CELL_DEBT_SIZING_TOLERANCE = "B6"

CELL_MASTER_CHECK_LINK = "B9"


def build_cover(wb: Workbook) -> Worksheet:
    ws = wb.create_sheet("Cover", 0)  # index 0: first sheet, opens here
    ws.sheet_properties.tabColor = TAB_COLOR_INPUT

    ws["A1"] = "Project 123 — Cover"
    ws["A1"].font = Font(bold=True, size=14)

    ws["A3"] = "VBA Solve Settings (used from Stage 1b onward)"
    ws["A3"].font = Font(bold=True)

    ws["A4"] = "Circularity Tolerance ($)"
    ws[CELL_CIRC_TOLERANCE] = 1.0
    ws[CELL_CIRC_TOLERANCE].font = Font(color=COLOR_INPUT)

    ws["A5"] = "Max Iterations (per convergence loop)"
    ws[CELL_MAX_ITERATIONS] = 100
    ws[CELL_MAX_ITERATIONS].font = Font(color=COLOR_INPUT)

    ws["A6"] = "Debt Sizing Tolerance ($, outer root-find on debt size D)"
    ws[CELL_DEBT_SIZING_TOLERANCE] = 100.0
    ws[CELL_DEBT_SIZING_TOLERANCE].font = Font(color=COLOR_INPUT)

    ws["A8"] = "Model Status"
    ws["A8"].font = Font(bold=True)
    status_cell = ws[CELL_MASTER_CHECK_LINK]
    status_cell.value = "=Check_Control!B3"
    status_cell.font = Font(bold=True, size=14)

    ws["A11"] = "Buttons (Stage 1b+): Solve Construction IDC | Solve Debt Sculpting | Goal Seek -> EIRR | Goal Seek -> PIRR"
    ws["A11"].font = Font(italic=True)
    ws["A12"] = "Placeholder rows only — Form Control buttons + macro assignment are a manual, one-time step (see VBA hand-off docs)."
    ws["A12"].font = Font(italic=True, size=9)

    _add_named_range(wb, "Cover_CircTolerance", "Cover", CELL_CIRC_TOLERANCE)
    _add_named_range(wb, "Cover_MaxIterations", "Cover", CELL_MAX_ITERATIONS)
    _add_named_range(wb, "Cover_DebtSizingTolerance", "Cover", CELL_DEBT_SIZING_TOLERANCE)

    ws.column_dimensions["A"].width = 45

    return ws


def _add_named_range(wb: Workbook, name: str, sheet: str, cell: str) -> None:
    import re

    from openpyxl.workbook.defined_name import DefinedName

    match = re.match(r"([A-Za-z]+)(\d+)", cell)
    col, row = match.group(1), match.group(2)
    wb.defined_names.add(DefinedName(name, attr_text=f"'{sheet}'!${col}${row}"))
