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
    from cover import build_cover
    from assumptions_model import build_assumptions_model
    from calc_capex import build_calc_capex
    from calc_financing_cons import build_calc_financing_cons
    from calc_revenue_opex import build_calc_revenue_opex
    from calc_tax import build_calc_tax
    from calc_financing_ops import build_calc_financing_ops
    from calc_cfads import build_calc_cfads
    from fs_quarterly import build_fs_quarterly
    from fs_annual import build_fs_annual
    from check_control import build_check_control

    wb = new_workbook()
    build_cover(wb, len(timeline.construction_months))
    build_assumptions_model(wb, timeline, inputs)
    build_calc_capex(wb, timeline, inputs)
    build_calc_financing_cons(wb, timeline, inputs)
    build_calc_revenue_opex(wb, timeline, inputs)
    build_calc_tax(wb, timeline, inputs)
    build_calc_financing_ops(wb, timeline, inputs)
    build_calc_cfads(wb, timeline, inputs)
    build_fs_quarterly(wb, timeline, inputs)
    build_fs_annual(wb, timeline, inputs)
    build_check_control(wb)
    wb.save(output_path)
    return wb
