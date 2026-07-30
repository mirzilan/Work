from openpyxl.styles import Font
from openpyxl.worksheet.worksheet import Worksheet
from openpyxl.workbook import Workbook

from inputs import ProjectInputs
from timeline import Timeline
from workbook_builder import FIRST_DATA_COL, COLOR_INPUT, TAB_COLOR_INPUT, col_letter

ROW_TOTAL_CAPEX = 3
ROW_DEBT_PCT = 4
ROW_INTEREST_RATE = 5
ROW_DEBT_TENOR_YEARS = 6
ROW_TARGET_DSCR = 7
ROW_ANNUAL_REVENUE = 8
ROW_OPEX_PCT = 9
ROW_TAX_RATE = 10
ROW_USEFUL_LIFE_YEARS = 11

ROW_MAX_GEARING = 12

ROW_PHASING_LABEL = 14
ROW_PHASING_DATE = 15
ROW_PHASING_PCT = 16

# Cell refs (single-scenario, Stage 1b interim — becomes Assumptions_Constant's Active
# column and Assumptions_Periodic_Capex's Active row once Stage 1c scales to 10 scenarios)
CELL_TOTAL_CAPEX = "B3"
CELL_DEBT_PCT = "B4"
CELL_INTEREST_RATE = "B5"
CELL_DEBT_TENOR_YEARS = "B6"
CELL_TARGET_DSCR = "B7"
CELL_ANNUAL_REVENUE = "B8"
CELL_OPEX_PCT = "B9"
CELL_TAX_RATE = "B10"
CELL_USEFUL_LIFE_YEARS = "B11"


def build_assumptions_model(wb: Workbook, timeline: Timeline, inputs: ProjectInputs) -> Worksheet:
    ws = wb.create_sheet("Assumptions_Model")
    ws.sheet_properties.tabColor = TAB_COLOR_INPUT

    ws["A1"] = "Assumptions_Model — Centralized Inputs (single scenario, Stage 1b interim)"
    ws["A1"].font = Font(bold=True, size=12)

    _input(ws, "A3", "Total Capex ($)", CELL_TOTAL_CAPEX, inputs.capex.total_capex, "#,##0")
    _input(ws, "A4", "Debt % of Capex", CELL_DEBT_PCT, inputs.financing.debt_pct_of_capex, "0.00%")
    _input(ws, "A5", "Interest Rate (Annual)", CELL_INTEREST_RATE, inputs.financing.interest_rate_annual, "0.00%")
    _input(ws, "A6", "Debt Tenor (Years)", CELL_DEBT_TENOR_YEARS, inputs.financing.debt_tenor_years, "0")
    _input(ws, "A7", "Target DSCR (used from Stage 1b Loop 2 onward)", CELL_TARGET_DSCR, inputs.financing.target_dscr, "0.00x")
    _input(ws, "A8", "Annual Revenue ($) — flat dummy placeholder", CELL_ANNUAL_REVENUE, inputs.revenue_opex.annual_revenue, "#,##0")
    _input(ws, "A9", "Opex % of Revenue", CELL_OPEX_PCT, inputs.revenue_opex.opex_pct_of_revenue, "0.00%")
    _input(ws, "A10", "Tax Rate", CELL_TAX_RATE, inputs.tax.tax_rate, "0.00%")
    _input(ws, "A11", "Useful Life (Years)", CELL_USEFUL_LIFE_YEARS, inputs.tax.useful_life_years, "0")
    _input(ws, "A12", "Max Gearing (cap on DSCR-sculpted debt size)", "B12", 0.85, "0.00%")

    ws.cell(row=ROW_PHASING_LABEL, column=1, value="Capex Phasing % by Construction Month").font = Font(bold=True)
    ws.cell(row=ROW_PHASING_DATE, column=1, value="Period End Date")
    ws.cell(row=ROW_PHASING_PCT, column=1, value="Phasing %")

    for i, period in enumerate(timeline.construction_months):
        col = col_letter(i)
        date_cell = ws[f"{col}{ROW_PHASING_DATE}"]
        date_cell.value = period.end
        date_cell.number_format = "mmm-yy"

        pct_cell = ws[f"{col}{ROW_PHASING_PCT}"]
        pct_cell.value = inputs.capex.phasing_pct_by_month[i]
        pct_cell.font = Font(color=COLOR_INPUT)
        pct_cell.number_format = "0.00%"

    ws.freeze_panes = ws.cell(row=ROW_PHASING_PCT + 1, column=FIRST_DATA_COL)

    return ws


def _input(ws: Worksheet, label_cell: str, label: str, value_cell: str, value, number_format: str) -> None:
    ws[label_cell] = label
    cell = ws[value_cell]
    cell.value = value
    cell.font = Font(color=COLOR_INPUT)
    cell.number_format = number_format
