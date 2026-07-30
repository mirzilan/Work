from openpyxl.styles import Font
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.worksheet import Worksheet
from openpyxl.workbook import Workbook

from workbook_builder import COLOR_INPUT, TAB_COLOR_INPUT

CELL_CIRC_TOLERANCE = "B4"
CELL_MAX_ITERATIONS = "B5"
CELL_DEBT_SIZING_TOLERANCE = "B6"

CELL_MASTER_CHECK_LINK = "B9"

CELL_DRAWDOWN_METHOD = "B14"
CELL_DEBT_SIZING_MODE = "B17"

DRAWDOWN_METHODS = ["Debt First", "Equity First", "Pari Passu"]
DEBT_SIZING_MODES = ["Fixed Gearing", "DSCR Sculpted"]


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

    ws["A14"] = "Construction Drawdown Method"
    ws["A14"].font = Font(bold=True)
    method_cell = ws[CELL_DRAWDOWN_METHOD]
    method_cell.value = DRAWDOWN_METHODS[2]
    method_cell.font = Font(color=COLOR_INPUT)

    validation = DataValidation(
        type="list",
        formula1=f'"{",".join(DRAWDOWN_METHODS)}"',
        allow_blank=False,
        showDropDown=False,
    )
    ws.add_data_validation(validation)
    validation.add(method_cell)

    ws["A15"] = (
        "Debt First: draw debt until facility exhausted, then equity. "
        "Equity First: draw equity until commitment exhausted, then debt. "
        "Pari Passu: draw both proportionally each month."
    )
    ws["A15"].font = Font(italic=True, size=9)

    ws["A17"] = "Debt Sizing Mode"
    ws["A17"].font = Font(bold=True)
    sizing_cell = ws[CELL_DEBT_SIZING_MODE]
    sizing_cell.value = DEBT_SIZING_MODES[1]
    sizing_cell.font = Font(color=COLOR_INPUT)

    sizing_validation = DataValidation(
        type="list",
        formula1=f'"{",".join(DEBT_SIZING_MODES)}"',
        allow_blank=False,
        showDropDown=False,
    )
    ws.add_data_validation(sizing_validation)
    sizing_validation.add(sizing_cell)

    ws["A18"] = (
        "Fixed Gearing: debt = gearing x Total Project Cost, level (PMT) amortisation. "
        "DSCR Sculpted: repayment locked to Target DSCR, debt size solved so the balance "
        "amortises to zero exactly at tenor end."
    )
    ws["A18"].font = Font(italic=True, size=9)

    _add_named_range(wb, "Cover_CircTolerance", "Cover", CELL_CIRC_TOLERANCE)
    _add_named_range(wb, "Cover_DebtSizingMode", "Cover", CELL_DEBT_SIZING_MODE)
    _add_named_range(wb, "Cover_MaxIterations", "Cover", CELL_MAX_ITERATIONS)
    _add_named_range(wb, "Cover_DebtSizingTolerance", "Cover", CELL_DEBT_SIZING_TOLERANCE)
    _add_named_range(wb, "Cover_DrawdownMethod", "Cover", CELL_DRAWDOWN_METHOD)

    ws.column_dimensions["A"].width = 45

    return ws


def _add_named_range(wb: Workbook, name: str, sheet: str, cell: str) -> None:
    import re

    from openpyxl.workbook.defined_name import DefinedName

    match = re.match(r"([A-Za-z]+)(\d+)", cell)
    col, row = match.group(1), match.group(2)
    wb.defined_names.add(DefinedName(name, attr_text=f"'{sheet}'!${col}${row}"))
