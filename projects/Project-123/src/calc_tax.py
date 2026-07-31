from openpyxl.styles import Font
from openpyxl.worksheet.worksheet import Worksheet
from openpyxl.workbook import Workbook

from inputs import ProjectInputs
from timeline import Timeline
# Import the source sheet's row constants rather than restating them: indexing another
# sheet with this module's own row numbers is how Calc_Tax once pulled Revenue while
# labelling it EBITDA.
from calc_revenue_opex import ROW_EBITDA as REVOPEX_ROW_EBITDA
from workbook_builder import (
    FIRST_DATA_COL,
    COLOR_FORMULA,
    COLOR_LINK,
    TAB_COLOR_CALC,
    col_letter,
)

ROW_DATE_HEADER = 2
ROW_QUARTER_INDEX = 3

ROW_EBITDA = 5          # linked from Calc_Revenue_Opex
ROW_DEPRECIATION = 6    # straight-line on Total Project Cost over useful life, single vintage
ROW_INTEREST = 7        # linked from Calc_Financing_Ops — the tax shield
ROW_EBT = 8             # EBITDA - Depreciation - Interest
ROW_TAX = 9             # max(EBT,0) * tax rate — no loss carryforward yet
ROW_NET_INCOME = 10

ROW_CHECK_HEADER = 14
ROW_CHECK_ACCUM_DEPR = 15

TAX_RATE_CELL = "B4"
USEFUL_LIFE_YEARS_CELL = "B9"


def build_calc_tax(wb: Workbook, timeline: Timeline, inputs: ProjectInputs) -> Worksheet:
    ws = wb.create_sheet("Calc_Tax")
    ws.sheet_properties.tabColor = TAB_COLOR_CALC

    ws["A1"] = "Calc_Tax — Quarterly Tax (single vintage, Stage 1a)"
    ws["A1"].font = Font(bold=True, size=12)

    ws["A4"] = "Tax Rate — linked from Assumptions_Model"
    ws[TAX_RATE_CELL] = "=Assumptions_Model!$B$10"
    ws[TAX_RATE_CELL].font = Font(color=COLOR_LINK)
    ws[TAX_RATE_CELL].number_format = "0.00%"

    ws["A9"] = "Useful Life (Years) — linked from Assumptions_Model"
    ws[USEFUL_LIFE_YEARS_CELL] = "=Assumptions_Model!$B$11"
    ws[USEFUL_LIFE_YEARS_CELL].font = Font(color=COLOR_LINK)

    _label(ws, ROW_DATE_HEADER, "Period End Date")
    _label(ws, ROW_QUARTER_INDEX, "Operating Quarter #")
    _label(ws, ROW_EBITDA, "EBITDA ($) — linked from Calc_Revenue_Opex")
    _label(ws, ROW_DEPRECIATION, "Depreciation ($) — straight-line, single vintage")
    _label(ws, ROW_INTEREST, "Interest Expense ($) — linked from Calc_Financing_Ops (tax shield)")
    _label(ws, ROW_EBT, "EBT ($) = EBITDA - Depreciation - Interest")
    _label(ws, ROW_TAX, "Tax ($) — no loss carryforward yet")
    _label(ws, ROW_NET_INCOME, "Net Income ($)")

    ws.cell(row=ROW_CHECK_HEADER, column=1, value="Checks").font = Font(bold=True)
    _label(ws, ROW_CHECK_ACCUM_DEPR, "Check: Accumulated Depreciation <= Total Capex")

    n_quarters = len(timeline.operations_quarters)
    last_cons_col = col_letter(len(timeline.construction_months) - 1)
    # Depreciable base is Total Project Cost (capex + capitalised IDC), not capex alone —
    # PP&E on the balance sheet opens at TPC, so depreciating capex only would strand the
    # IDC portion undepreciated for the life of the asset.
    tpc_ref = f"Calc_Capex!${last_cons_col}$17"
    quarterly_depr_expr = f"{tpc_ref}/($B$9*4)"

    for i, period in enumerate(timeline.operations_quarters):
        col = col_letter(i)

        ws[f"{col}{ROW_DATE_HEADER}"] = period.end
        ws[f"{col}{ROW_DATE_HEADER}"].number_format = "mmm-yy"
        ws[f"{col}{ROW_QUARTER_INDEX}"] = i + 1

        ebitda_cell = ws[f"{col}{ROW_EBITDA}"]
        ebitda_cell.value = f"=Calc_Revenue_Opex!{col}{REVOPEX_ROW_EBITDA}"
        ebitda_cell.font = Font(color=COLOR_LINK)
        ebitda_cell.number_format = "#,##0"

        # Depreciation stops once fully depreciated (quarter index < useful_life_years*4)
        useful_life_quarters = inputs.tax.useful_life_years * 4
        depr_cell = ws[f"{col}{ROW_DEPRECIATION}"]
        if i < useful_life_quarters:
            depr_cell.value = f"={quarterly_depr_expr}"
        else:
            depr_cell.value = 0
        depr_cell.font = Font(color=COLOR_FORMULA)
        depr_cell.number_format = "#,##0"

        interest_cell = ws[f"{col}{ROW_INTEREST}"]
        interest_cell.value = f"=Calc_Financing_Ops!{col}17"  # ROW_INTEREST in calc_financing_ops.py
        interest_cell.font = Font(color=COLOR_LINK)
        interest_cell.number_format = "#,##0"

        ebt_cell = ws[f"{col}{ROW_EBT}"]
        ebt_cell.value = f"={col}{ROW_EBITDA}-{col}{ROW_DEPRECIATION}-{col}{ROW_INTEREST}"
        ebt_cell.font = Font(color=COLOR_FORMULA)
        ebt_cell.number_format = "#,##0"

        tax_cell = ws[f"{col}{ROW_TAX}"]
        tax_cell.value = f"=MAX({col}{ROW_EBT},0)*$B$4"
        tax_cell.font = Font(color=COLOR_FORMULA)
        tax_cell.number_format = "#,##0"

        ni_cell = ws[f"{col}{ROW_NET_INCOME}"]
        ni_cell.value = f"={col}{ROW_EBT}-{col}{ROW_TAX}"
        ni_cell.font = Font(color=COLOR_FORMULA)
        ni_cell.number_format = "#,##0"

    last_col = col_letter(n_quarters - 1)
    first_col = col_letter(0)

    check_cell = ws[f"{last_col}{ROW_CHECK_ACCUM_DEPR}"]
    check_cell.value = (
        f"=IF(SUM({first_col}{ROW_DEPRECIATION}:{last_col}{ROW_DEPRECIATION})<={tpc_ref}+0.01,1,0)"
    )
    check_cell.font = Font(color=COLOR_FORMULA)

    _add_named_range(wb, "Tax_LastCol", "Calc_Tax", f"{last_col}1")
    _add_named_range(wb, "Tax_AccumDeprCheck", "Calc_Tax", f"{last_col}{ROW_CHECK_ACCUM_DEPR}")

    ws.freeze_panes = ws.cell(row=ROW_NET_INCOME + 1, column=FIRST_DATA_COL)

    return ws


def _label(ws: Worksheet, row: int, label: str) -> None:
    ws.cell(row=row, column=1, value=label)


def _add_named_range(wb: Workbook, name: str, sheet: str, cell: str) -> None:
    import re

    from openpyxl.workbook.defined_name import DefinedName

    match = re.match(r"([A-Za-z]+)(\d+)", cell)
    col, row = match.group(1), match.group(2)
    wb.defined_names.add(DefinedName(name, attr_text=f"'{sheet}'!${col}${row}"))
