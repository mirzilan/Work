from openpyxl.styles import Font
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.worksheet import Worksheet
from openpyxl.workbook import Workbook

from inputs import ProjectInputs
from workbook_builder import COLOR_INPUT, COLOR_FORMULA, TAB_COLOR_INPUT

N_SCENARIOS = 10
COL_ACTIVE = "B"
FIRST_SCENARIO_COL = 3  # column C

ROW_SCENARIO_HEADER = 3
ROW_SCENARIO_NAME = 4
ROW_FIRST_DRIVER = 6

ROW_CHECK_HEADER = 23
ROW_CHECK_ACTIVE_RESOLVES = 24
ROW_CHECK_ALL_POPULATED = 25

SCENARIO_NAMES = [
    "Base", "Upside", "Downside",
    "Scenario 4", "Scenario 5", "Scenario 6", "Scenario 7",
    "Scenario 8", "Scenario 9", "Scenario 10",
]

# (label, number_format, base, upside, downside)
# Scenarios 4-10 are seeded with the base values as editable placeholders.
DRIVERS = [
    ("Total Capex ($)", "#,##0", 100_000_000.0, 95_000_000.0, 110_000_000.0),
    ("Gearing (Debt % of TPC)", "0.00%", 0.70, 0.75, 0.65),
    ("Interest Rate (Annual)", "0.00%", 0.06, 0.055, 0.07),
    ("Debt Tenor (Years)", "0", 15, 17, 12),
    ("Target DSCR", "0.00x", 1.30, 1.25, 1.40),
    ("Annual Revenue ($)", "#,##0", 15_000_000.0, 18_000_000.0, 12_000_000.0),
    ("Opex % of Revenue", "0.00%", 0.30, 0.27, 0.35),
    ("Tax Rate", "0.00%", 0.25, 0.25, 0.25),
    ("Useful Life (Years)", "0", 20, 20, 20),
    ("Max Gearing (cap)", "0.00%", 0.85, 0.85, 0.85),
    ("Escalation Factor 1 (Annual %)", "0.00%", 0.0, 0.0, 0.0),
    ("Escalation Factor 2 (Annual %)", "0.00%", 0.0, 0.0, 0.0),
    ("Escalation Factor 3 (Annual %)", "0.00%", 0.025, 0.020, 0.040),
    ("Escalation Factor 4 (Annual %)", "0.00%", 0.030, 0.025, 0.045),
    ("Routine Maint Capex (% of Revenue)", "0.00%", 0.015, 0.012, 0.020),
    ("DSRA LC Fee (% p.a. on requirement)", "0.00%", 0.015, 0.015, 0.020),
]

# Row offsets so other modules can reference a driver without hardcoding row numbers.
ROW_TOTAL_CAPEX = ROW_FIRST_DRIVER + 0
ROW_GEARING = ROW_FIRST_DRIVER + 1
ROW_INTEREST_RATE = ROW_FIRST_DRIVER + 2
ROW_DEBT_TENOR = ROW_FIRST_DRIVER + 3
ROW_TARGET_DSCR = ROW_FIRST_DRIVER + 4
ROW_ANNUAL_REVENUE = ROW_FIRST_DRIVER + 5
ROW_OPEX_PCT = ROW_FIRST_DRIVER + 6
ROW_TAX_RATE = ROW_FIRST_DRIVER + 7
ROW_USEFUL_LIFE = ROW_FIRST_DRIVER + 8
ROW_MAX_GEARING = ROW_FIRST_DRIVER + 9
ROW_ESC_FACTOR_1 = ROW_FIRST_DRIVER + 10
ROW_ROUTINE_MAINT_PCT = ROW_FIRST_DRIVER + 14
ROW_DSRA_LC_FEE = ROW_FIRST_DRIVER + 15


def build_assumptions_constant(wb: Workbook, inputs: ProjectInputs) -> Worksheet:
    ws = wb.create_sheet("Assumptions_Constant")
    ws.sheet_properties.tabColor = TAB_COLOR_INPUT

    ws["A1"] = "Assumptions_Constant — Single-Value Drivers, Scenarios Across Columns"
    ws["A1"].font = Font(bold=True, size=12)

    ws.cell(row=ROW_SCENARIO_HEADER, column=1, value="Driver").font = Font(bold=True)
    active_hdr = ws.cell(row=ROW_SCENARIO_HEADER, column=2, value="Active")
    active_hdr.font = Font(bold=True)

    for s in range(N_SCENARIOS):
        col = get_column_letter(FIRST_SCENARIO_COL + s)
        hdr = ws[f"{col}{ROW_SCENARIO_HEADER}"]
        hdr.value = s + 1
        hdr.font = Font(bold=True)

    ws.cell(row=ROW_SCENARIO_NAME, column=1, value="Scenario Name")
    name_active = ws[f"{COL_ACTIVE}{ROW_SCENARIO_NAME}"]
    name_active.value = _index_formula(ROW_SCENARIO_NAME)
    name_active.font = Font(color=COLOR_FORMULA, bold=True)

    for s, name in enumerate(SCENARIO_NAMES):
        col = get_column_letter(FIRST_SCENARIO_COL + s)
        cell = ws[f"{col}{ROW_SCENARIO_NAME}"]
        cell.value = name
        cell.font = Font(color=COLOR_INPUT)

    for i, (label, fmt, base, upside, downside) in enumerate(DRIVERS):
        row = ROW_FIRST_DRIVER + i
        ws.cell(row=row, column=1, value=label)

        active = ws[f"{COL_ACTIVE}{row}"]
        active.value = _index_formula(row)
        active.font = Font(color=COLOR_FORMULA)
        active.number_format = fmt

        values = [base, upside, downside] + [base] * (N_SCENARIOS - 3)
        for s, value in enumerate(values):
            col = get_column_letter(FIRST_SCENARIO_COL + s)
            cell = ws[f"{col}{row}"]
            cell.value = value
            cell.font = Font(color=COLOR_INPUT)
            cell.number_format = fmt

    last_driver_row = ROW_FIRST_DRIVER + len(DRIVERS) - 1
    last_scenario_col = get_column_letter(FIRST_SCENARIO_COL + N_SCENARIOS - 1)
    first_scenario_col = get_column_letter(FIRST_SCENARIO_COL)

    ws.cell(row=ROW_CHECK_HEADER, column=1, value="Checks").font = Font(bold=True)
    ws.cell(row=ROW_CHECK_ACTIVE_RESOLVES, column=1, value="Check: Active column resolves (Total Capex > 0)")
    ws.cell(row=ROW_CHECK_ALL_POPULATED, column=1, value="Check: All scenario cells populated")

    resolves = ws[f"{COL_ACTIVE}{ROW_CHECK_ACTIVE_RESOLVES}"]
    resolves.value = f"=IF(AND(ISNUMBER({COL_ACTIVE}{ROW_TOTAL_CAPEX}),{COL_ACTIVE}{ROW_TOTAL_CAPEX}>0),1,0)"
    resolves.font = Font(color=COLOR_FORMULA)

    populated = ws[f"{COL_ACTIVE}{ROW_CHECK_ALL_POPULATED}"]
    populated.value = (
        f"=IF(COUNT({first_scenario_col}{ROW_FIRST_DRIVER}:{last_scenario_col}{last_driver_row})"
        f"={N_SCENARIOS * len(DRIVERS)},1,0)"
    )
    populated.font = Font(color=COLOR_FORMULA)

    _name(wb, "Const_ActiveResolvesCheck", "Assumptions_Constant", f"{COL_ACTIVE}{ROW_CHECK_ACTIVE_RESOLVES}")
    _name(wb, "Const_AllPopulatedCheck", "Assumptions_Constant", f"{COL_ACTIVE}{ROW_CHECK_ALL_POPULATED}")

    # The goal-seek macro writes revenue into the *scenario* column, never the Active
    # column (which is an INDEX formula). Naming the whole row lets the macro pick the
    # right cell off the scenario selector without hardcoding a row or column.
    from openpyxl.workbook.defined_name import DefinedName

    wb.defined_names.add(DefinedName(
        "ScenarioRevenueRow",
        attr_text=(f"'Assumptions_Constant'!${first_scenario_col}${ROW_ANNUAL_REVENUE}"
                   f":${last_scenario_col}${ROW_ANNUAL_REVENUE}"),
    ))

    ws.column_dimensions["A"].width = 34
    ws.freeze_panes = ws["C6"]

    return ws


def _index_formula(row: int) -> str:
    """Pull the selected scenario's column. Row range is absolute so the formula can be
    copied down the driver list without drifting off the scenario block."""
    first = get_column_letter(FIRST_SCENARIO_COL)
    last = get_column_letter(FIRST_SCENARIO_COL + N_SCENARIOS - 1)
    return f"=INDEX(${first}${row}:${last}${row},1,Cover!$B$20)"


def _name(wb: Workbook, name: str, sheet: str, cell: str) -> None:
    import re

    from openpyxl.workbook.defined_name import DefinedName

    m = re.match(r"([A-Za-z]+)(\d+)", cell)
    wb.defined_names.add(DefinedName(name, attr_text=f"'{sheet}'!${m.group(1)}${m.group(2)}"))
