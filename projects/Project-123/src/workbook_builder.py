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


def new_workbook(template_path: str | None = None) -> Workbook:
    """Without a template, a blank workbook. With one, the template's .xlsm is reopened
    with its VBA project intact and every sheet dropped, so the builders repopulate a
    workbook that still carries the macros. openpyxl copies vbaProject.bin verbatim —
    verified to survive a full delete-and-rebuild of all sheets."""
    if template_path:
        wb = openpyxl.load_workbook(template_path, keep_vba=True)
        for name in list(wb.sheetnames):
            del wb[name]
        # Stale names would collide with the ones the builders re-register.
        for name in list(wb.defined_names):
            del wb.defined_names[name]
        return wb

    wb = openpyxl.Workbook()
    # Remove the default sheet; each builder module adds its own named sheet.
    wb.remove(wb.active)
    return wb


def col_letter(period_index: int) -> str:
    return openpyxl.utils.get_column_letter(FIRST_DATA_COL + period_index)


def build_workbook(inputs: ProjectInputs, timeline: Timeline, output_path: str,
                   template_path: str | None = None) -> Workbook:
    from cover import build_cover
    from assumptions_model import build_assumptions_model
    from assumptions_constant import build_assumptions_constant
    from assumptions_periodic_capex import build_assumptions_periodic_capex
    from assumptions_periodic_ops import build_assumptions_periodic_ops
    from calc_capex import build_calc_capex
    from calc_financing_cons import build_calc_financing_cons
    from calc_revenue_opex import build_calc_revenue_opex
    from calc_tax import build_calc_tax
    from calc_financing_ops import build_calc_financing_ops
    from calc_cfads import build_calc_cfads
    from fs_quarterly import build_fs_quarterly
    from fs_annual import build_fs_annual
    from batch_results import build_batch_results
    from check_control import build_check_control


    n_quarters = len(timeline.operations_quarters)
    tenor_quarters = min(inputs.financing.debt_tenor_years * 4, n_quarters)

    wb = new_workbook(template_path)
    build_cover(wb, len(timeline.construction_months),
                tenor_end_col=col_letter(tenor_quarters - 1),
                n_operating_quarters=n_quarters)
    build_assumptions_model(wb, timeline, inputs)
    build_assumptions_constant(wb, inputs)
    build_assumptions_periodic_capex(wb, timeline, inputs)
    build_assumptions_periodic_ops(wb, timeline, inputs)
    build_calc_capex(wb, timeline, inputs)
    build_calc_financing_cons(wb, timeline, inputs)
    build_calc_revenue_opex(wb, timeline, inputs)
    build_calc_tax(wb, timeline, inputs)
    build_calc_financing_ops(wb, timeline, inputs)
    build_calc_cfads(wb, timeline, inputs)
    build_fs_quarterly(wb, timeline, inputs)
    build_fs_annual(wb, timeline, inputs)
    build_batch_results(wb)
    build_check_control(wb, timeline)
    wb.save(output_path)
    return wb
