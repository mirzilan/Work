import openpyxl
from openpyxl.workbook import Workbook

from inputs import ProjectInputs
from timeline import Timeline

# Column layout: column A/B reserved for labels, data starts at column C.
FIRST_DATA_COL = 3  # column C

# FAST/Corality-style font colors (cell-level, per CLAUDE.md convention)
COLOR_INPUT = "FF0000FF"  # blue: hardcoded inputs/assumptions
COLOR_FORMULA = "FF000000"  # black: formulas within sheet
COLOR_LINK = "FF008000"  # green: links from another sheet

# Sheet tab colors (Cover/Assumptions=blue, Calc=grey, Output=green, Check=red)
TAB_COLOR_INPUT = "0000FF"
TAB_COLOR_CALC = "808080"
TAB_COLOR_OUTPUT = "00B050"
TAB_COLOR_CHECK = "FF0000"


def new_workbook() -> Workbook:
    wb = openpyxl.Workbook()
    # Remove the default sheet; each builder module adds its own named sheet.
    wb.remove(wb.active)
    return wb


def col_letter(period_index: int) -> str:
    return openpyxl.utils.get_column_letter(FIRST_DATA_COL + period_index)


def build_workbook(inputs: ProjectInputs, timeline: Timeline, output_path: str) -> Workbook:
    from calc_capex import build_calc_capex
    from calc_financing_cons import build_calc_financing_cons
    from calc_revenue_opex import build_calc_revenue_opex

    wb = new_workbook()
    build_calc_capex(wb, timeline, inputs)
    build_calc_financing_cons(wb, timeline, inputs)
    build_calc_revenue_opex(wb, timeline, inputs)
    wb.save(output_path)
    return wb
