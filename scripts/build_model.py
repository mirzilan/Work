"""
Builds HydrogenLCOH_Model.xlsx: a quarterly-period project-finance model for a
green hydrogen plant, headlined by an unlevered LCOH calc, with a closed-form
(non-circular) IDC formula and a DSCR-sculpting debt schedule that a VBA
bisection routine (see vba/mod_DebtSizing.bas) sizes by solving for the debt
quantum (GearingFactor) that fully amortizes by tenor end at a flat target
DSCR. See /root/.claude/plans/plan-first-can-u-playful-beaver.md for the design.

Run: python3 scripts/build_model.py
Output: HydrogenLCOH_Model.xlsx (repo root)
"""
import openpyxl
from openpyxl import Workbook
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.utils import get_column_letter
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.worksheet.datavalidation import DataValidation

# ---------------------------------------------------------------------------
# Timeline constants
# ---------------------------------------------------------------------------
CONSTRUCTION_YEARS = 2
OPERATIONS_YEARS = 20
N_CONSTRUCTION = CONSTRUCTION_YEARS * 4          # 8 quarters
N_OPERATIONS = OPERATIONS_YEARS * 4              # 80 quarters
N_PERIODS = N_CONSTRUCTION + N_OPERATIONS        # 88 quarters
COD_PERIOD = N_CONSTRUCTION + 1                  # first operating quarter = 9
DEBT_TENOR_YEARS_DEFAULT = 15
TENOR_QUARTERS_DEFAULT = DEBT_TENOR_YEARS_DEFAULT * 4
TENOR_END_PERIOD_DEFAULT = N_CONSTRUCTION + TENOR_QUARTERS_DEFAULT  # 68

FIRST_PERIOD_COL = 4   # column D
LAST_PERIOD_COL = FIRST_PERIOD_COL + N_PERIODS - 1  # column CM (91)

INPUT_FILL = PatternFill("solid", fgColor="DDEBF7")
LINK_FILL = PatternFill("solid", fgColor="E2EFDA")
HEADER_FONT = Font(bold=True, size=12)
BOLD = Font(bold=True)


def pcol(period):
    """Excel column letter for a given 1-indexed period number."""
    return get_column_letter(FIRST_PERIOD_COL + period - 1)


def add_named_cell(wb, name, sheet, col_letter, row):
    ref = f"'{sheet}'!${col_letter}${row}"
    wb.defined_names[name] = DefinedName(name, attr_text=ref)


def write_input_row(ws, row, label, value, unit="", note=""):
    ws.cell(row=row, column=1, value=label)
    c = ws.cell(row=row, column=2, value=value)
    c.fill = INPUT_FILL
    ws.cell(row=row, column=3, value=unit)
    if note:
        ws.cell(row=row, column=5, value=note)


def label_row(ws, row, label, unit=""):
    ws.cell(row=row, column=1, value=label).font = BOLD
    if unit:
        ws.cell(row=row, column=3, value=unit)


wb = Workbook()
wb.remove(wb.active)

# ===========================================================================
# 1. COVER
# ===========================================================================
cover = wb.create_sheet("Cover")
cover.column_dimensions["A"].width = 100
lines = [
    "Green Hydrogen LCOH Project-Finance Model",
    "",
    "Conventions:",
    "  - Blue fill = hardcoded input. White = formula. Green fill = cross-sheet link.",
    "  - Quarterly periods across columns D onward, one row per line item.",
    f"  - Construction: {CONSTRUCTION_YEARS} yrs ({N_CONSTRUCTION} quarters, periods 1-{N_CONSTRUCTION}).",
    f"  - Operations: {OPERATIONS_YEARS} yrs ({N_OPERATIONS} quarters, periods {COD_PERIOD}-{N_PERIODS}). COD = start of period {COD_PERIOD}.",
    "  - Iterative calculation is deliberately kept OFF (File > Options > Formulas).",
    "    Nothing in this model should ever need it -- if you see a circular-reference",
    "    warning, that flags a real bug, not something to toggle a setting around.",
    "",
    "Circularity handling:",
    "  - IDC (Interest During Construction) is solved in CLOSED FORM as a plain Excel",
    "    formula on Capex_Construction (see the derivation there) -- no VBA, no",
    "    iterative calc needed, even though a naive model would be circular here.",
    "  - Post-COD debt is SCULPTED to a flat target DSCR every period (live formulas",
    "    on Debt_Sizing_DSCR, driven by a single scalar cell 'GearingFactor'). Sizing",
    "    that scalar so the debt fully amortizes by tenor end IS a genuine root-finding",
    "    problem (CFADS is irregular period to period) -- that's what the VBA in",
    "    vba/mod_DebtSizing.bas bisects. Import the .bas files (Alt+F11 > File > Import",
    "    File) and wire 'RunDebtSizing' to the Dashboard button, then Save As .xlsm.",
    "  - FALLBACK if you don't import the macros: the sculpting recursion is pure",
    "    Excel formulas keyed off GearingFactor, so you can size debt manually with",
    "    native Goal Seek: Data > What-If Analysis > Goal Seek, Set cell",
    "    'EndingBalanceResidual' To 0 By changing cell 'GearingFactor'.",
    "",
    "LCOH is UNLEVERED: discounted lifecycle capex+opex at WACC, divided by",
    "discounted lifecycle H2 output. The debt/DSCR schedule is a separate",
    "financeability check and does not feed the LCOH numerator.",
]
for i, line in enumerate(lines, start=1):
    cell = cover.cell(row=i, column=1, value=line)
    if i == 1:
        cell.font = Font(bold=True, size=14)

# ===========================================================================
# 2. ASSUMPTIONS
# ===========================================================================
asm = wb.create_sheet("Assumptions")
asm.column_dimensions["A"].width = 38
asm.column_dimensions["B"].width = 14
asm.column_dimensions["C"].width = 12
asm.column_dimensions["E"].width = 60

row = 1
asm.cell(row=row, column=1, value="ASSUMPTIONS").font = HEADER_FONT
row += 2

label_row(asm, row, "Technical"); row += 1
r = {}
r["capacity_mw"] = row; write_input_row(asm, row, "Electrolyzer Capacity", 100, "MW"); row += 1
r["efficiency"] = row; write_input_row(asm, row, "Electrolyzer Efficiency", 55, "kWh/kg H2"); row += 1
r["cap_factor"] = row; write_input_row(asm, row, "Capacity Factor", 0.90, "%"); row += 1
r["degradation"] = row; write_input_row(asm, row, "Degradation Rate", 0.005, "%/yr"); row += 1
row += 1

label_row(asm, row, "Commercial"); row += 1
r["offtake_price"] = row; write_input_row(asm, row, "H2 Offtake Price", 4.50, "$/kg"); row += 1
r["fixed_opex"] = row; write_input_row(asm, row, "Fixed Opex", 5000000, "$/yr"); row += 1
r["var_opex"] = row; write_input_row(asm, row, "Variable Opex", 0.20, "$/kg"); row += 1
r["power_price"] = row; write_input_row(asm, row, "Power Price", 40, "$/MWh"); row += 1
r["water_cost"] = row; write_input_row(asm, row, "Water Cost", 0.05, "$/kg"); row += 1
r["stack_cost"] = row; write_input_row(asm, row, "Stack Replacement Cost", 20000000, "$"); row += 1
r["stack_interval"] = row; write_input_row(asm, row, "Stack Replacement Interval", 8, "yrs"); row += 1
row += 1

label_row(asm, row, "Capex"); row += 1
r["total_capex"] = row; write_input_row(asm, row, "Total Capex", 400000000, "$"); row += 1
r["constr_gearing"] = row; write_input_row(asm, row, "Construction Debt-Funded %", 0.70, "%",
    "Share of capex draws funded by debt during construction (rest is equity)"); row += 1
row += 1

label_row(asm, row, "Financing"); row += 1
r["wacc"] = row; write_input_row(asm, row, "WACC", 0.08, "%"); row += 1
r["cost_of_debt"] = row; write_input_row(asm, row, "Cost of Debt", 0.06, "%"); row += 1
r["idc_f"] = row; write_input_row(asm, row, "IDC Cash-Funded Fraction (f)", 0.30, "%",
    "0 = fully capitalized, 1 = fully cash-funded by equity/reserve"); row += 1
r["idc_timing"] = row; write_input_row(asm, row, "IDC Timing Convention", "MidPeriod", "",
    "MidPeriod (closed-form algebra) or EndPeriod (simple, already non-circular)"); row += 1
r["target_dscr"] = row; write_input_row(asm, row, "Target DSCR", 1.35, "x"); row += 1
r["max_gearing"] = row; write_input_row(asm, row, "Max Gearing", 0.80, "%"); row += 1
r["debt_tenor"] = row; write_input_row(asm, row, "Debt Tenor", DEBT_TENOR_YEARS_DEFAULT, "yrs"); row += 1
r["constr_years"] = row; write_input_row(asm, row, "Construction Period", CONSTRUCTION_YEARS, "yrs"); row += 1
r["ops_years"] = row; write_input_row(asm, row, "Operations Period", OPERATIONS_YEARS, "yrs"); row += 1

dv = DataValidation(type="list", formula1='"MidPeriod,EndPeriod"', allow_blank=False)
asm.add_data_validation(dv)
dv.add(asm.cell(row=r["idc_timing"], column=2))

# Named ranges for every assumption scalar
NAMES = {
    "ElectrolyzerCapacityMW": r["capacity_mw"],
    "EfficiencykWhPerKg": r["efficiency"],
    "CapacityFactor": r["cap_factor"],
    "DegradationRate": r["degradation"],
    "OfftakePrice": r["offtake_price"],
    "FixedOpexAnnual": r["fixed_opex"],
    "VariableOpexPerKg": r["var_opex"],
    "PowerPrice": r["power_price"],
    "WaterCostPerKg": r["water_cost"],
    "StackReplacementCost": r["stack_cost"],
    "StackReplacementIntervalYrs": r["stack_interval"],
    "TotalCapex": r["total_capex"],
    "ConstructionGearingPct": r["constr_gearing"],
    "WACC": r["wacc"],
    "CostOfDebtRate": r["cost_of_debt"],
    "IDC_CashFundedPct_f": r["idc_f"],
    "IDCTimingConvention": r["idc_timing"],
    "TargetDSCR_Input": r["target_dscr"],
    "MaxGearingPct": r["max_gearing"],
    "DebtTenorYears": r["debt_tenor"],
    "ConstructionYears": r["constr_years"],
    "OperationsYears": r["ops_years"],
}
for name, rr in NAMES.items():
    add_named_cell(wb, name, "Assumptions", "B", rr)

# ===========================================================================
# 3. CONTROL (master timeline)
# ===========================================================================
ctl = wb.create_sheet("Control")
ctl.column_dimensions["A"].width = 30
ctl.cell(row=1, column=1, value="CONTROL / MASTER TIMELINE").font = HEADER_FONT

CTL_ROWS = {
    "period": 3, "year": 4, "qtr_in_year": 5, "constr_flag": 6, "ops_flag": 7,
    "cod_flag": 8, "years_since_cod": 9, "r_debt": 10, "r_wacc": 11, "disc_factor": 12,
}
label_row(ctl, CTL_ROWS["period"], "Period #")
label_row(ctl, CTL_ROWS["year"], "Year #")
label_row(ctl, CTL_ROWS["qtr_in_year"], "Quarter in Year")
label_row(ctl, CTL_ROWS["constr_flag"], "Construction Flag")
label_row(ctl, CTL_ROWS["ops_flag"], "Operations Flag")
label_row(ctl, CTL_ROWS["cod_flag"], "COD Flag")
label_row(ctl, CTL_ROWS["years_since_cod"], "Years Since COD")
label_row(ctl, CTL_ROWS["r_debt"], "Periodic Debt Rate")
label_row(ctl, CTL_ROWS["r_wacc"], "Periodic WACC")
label_row(ctl, CTL_ROWS["disc_factor"], "Cumulative Discount Factor (WACC)")

for p in range(1, N_PERIODS + 1):
    c = pcol(p)
    ctl[f"{c}{CTL_ROWS['period']}"] = p
    ctl[f"{c}{CTL_ROWS['year']}"] = f"=ROUNDUP({c}{CTL_ROWS['period']}/4,0)"
    ctl[f"{c}{CTL_ROWS['qtr_in_year']}"] = f"={c}{CTL_ROWS['period']}-({c}{CTL_ROWS['year']}-1)*4"
    ctl[f"{c}{CTL_ROWS['constr_flag']}"] = f"=IF({c}{CTL_ROWS['period']}<=ConstructionYears*4,1,0)"
    ctl[f"{c}{CTL_ROWS['ops_flag']}"] = f"=IF({c}{CTL_ROWS['period']}>ConstructionYears*4,1,0)"
    ctl[f"{c}{CTL_ROWS['cod_flag']}"] = f"=IF({c}{CTL_ROWS['period']}=ConstructionYears*4+1,1,0)"
    ctl[f"{c}{CTL_ROWS['years_since_cod']}"] = f"={c}{CTL_ROWS['year']}-ConstructionYears"
    ctl[f"{c}{CTL_ROWS['r_debt']}"] = "=(1+CostOfDebtRate)^0.25-1"
    ctl[f"{c}{CTL_ROWS['r_wacc']}"] = "=(1+WACC)^0.25-1"
    ctl[f"{c}{CTL_ROWS['disc_factor']}"] = f"=1/(1+{c}{CTL_ROWS['r_wacc']})^{c}{CTL_ROWS['period']}"

ctl["B1"] = "=ConstructionYears*4+DebtTenorYears*4"
add_named_cell(wb, "TenorEndPeriod", "Control", "B", 1)
ctl["C1"] = "=ConstructionYears*4+1"
add_named_cell(wb, "CODPeriod", "Control", "C", 1)


def row_range_name(wb, name, sheet, row):
    ref = f"'{sheet}'!${pcol(1)}${row}:${pcol(N_PERIODS)}${row}"
    wb.defined_names[name] = DefinedName(name, attr_text=ref)

row_range_name(wb, "Period_Row", "Control", CTL_ROWS["period"])
row_range_name(wb, "ConstructionFlag_Row", "Control", CTL_ROWS["constr_flag"])
row_range_name(wb, "OperationsFlag_Row", "Control", CTL_ROWS["ops_flag"])
row_range_name(wb, "DiscountFactor_Row", "Control", CTL_ROWS["disc_factor"])
row_range_name(wb, "YearsSinceCOD_Row", "Control", CTL_ROWS["years_since_cod"])
row_range_name(wb, "QuarterInYear_Row", "Control", CTL_ROWS["qtr_in_year"])

# ===========================================================================
# 4. CAPEX_CONSTRUCTION (closed-form IDC)
# ===========================================================================
cx = wb.create_sheet("Capex_Construction")
cx.column_dimensions["A"].width = 34
cx.cell(row=1, column=1, value="CAPEX & CONSTRUCTION DRAWDOWN").font = HEADER_FONT

CX_ROWS = {
    "capex_draw": 3, "debt_draw": 4, "equity_draw": 5, "opening_bal": 6,
    "interest": 7, "cash_interest": 8, "cap_interest": 9, "closing_bal": 10,
}
label_row(cx, CX_ROWS["capex_draw"], "Capex Draw (Total)", "$")
label_row(cx, CX_ROWS["debt_draw"], "Debt Draw for Capex (D_t)", "$")
label_row(cx, CX_ROWS["equity_draw"], "Equity Draw for Capex", "$")
label_row(cx, CX_ROWS["opening_bal"], "Opening Debt Balance (B_t-1)", "$")
label_row(cx, CX_ROWS["interest"], "Interest I_t (closed form)", "$")
label_row(cx, CX_ROWS["cash_interest"], "Cash-Funded Interest (f x I_t)", "$")
label_row(cx, CX_ROWS["cap_interest"], "Capitalized Interest ((1-f) x I_t)", "$")
label_row(cx, CX_ROWS["closing_bal"], "Closing Debt Balance (B_t)", "$")

cx.cell(row=13, column=1,
    value="I_t = r*(B_(t-1) + D_t/2) / (1 - r*(1-f)/2)   [mid-period, algebraically isolated -- not circular]")
cx.cell(row=14, column=1,
    value="I_t = r*(B_(t-1) + D_t)                        [end-of-period draw convention, no algebra needed]")

for p in range(1, N_PERIODS + 1):
    c = pcol(p)
    prev = pcol(p - 1) if p > 1 else None
    ctl_c = f"Control!{c}"

    cx[f"{c}{CX_ROWS['capex_draw']}"] = (
        f"=IF({ctl_c}{CTL_ROWS['constr_flag']}=1,TotalCapex/(ConstructionYears*4),0)"
    )
    cx[f"{c}{CX_ROWS['debt_draw']}"] = f"={c}{CX_ROWS['capex_draw']}*ConstructionGearingPct"
    cx[f"{c}{CX_ROWS['equity_draw']}"] = f"={c}{CX_ROWS['capex_draw']}*(1-ConstructionGearingPct)"

    if p == 1:
        cx[f"{c}{CX_ROWS['opening_bal']}"] = 0
    else:
        cx[f"{c}{CX_ROWS['opening_bal']}"] = f"={prev}{CX_ROWS['closing_bal']}"

    r_debt = f"{ctl_c}{CTL_ROWS['r_debt']}"
    b_prev = f"{c}{CX_ROWS['opening_bal']}"
    d_t = f"{c}{CX_ROWS['debt_draw']}"
    mid_formula = f"({r_debt}*({b_prev}+{d_t}/2))/(1-{r_debt}*(1-IDC_CashFundedPct_f)/2)"
    end_formula = f"{r_debt}*({b_prev}+{d_t})"
    cx[f"{c}{CX_ROWS['interest']}"] = (
        f'=IF(IDCTimingConvention="EndPeriod",{end_formula},{mid_formula})'
    )
    cx[f"{c}{CX_ROWS['cash_interest']}"] = f"=IDC_CashFundedPct_f*{c}{CX_ROWS['interest']}"
    cx[f"{c}{CX_ROWS['cap_interest']}"] = f"=(1-IDC_CashFundedPct_f)*{c}{CX_ROWS['interest']}"
    cx[f"{c}{CX_ROWS['closing_bal']}"] = (
        f"={b_prev}+{d_t}+{c}{CX_ROWS['cap_interest']}"
    )

row_range_name(wb, "CapexDraw_Row", "Capex_Construction", CX_ROWS["capex_draw"])
row_range_name(wb, "EquityDrawCapex_Row", "Capex_Construction", CX_ROWS["equity_draw"])
row_range_name(wb, "CashFundedInterest_Row", "Capex_Construction", CX_ROWS["cash_interest"])
add_named_cell(wb, "ClosingBalanceAtCOD", "Capex_Construction", pcol(N_CONSTRUCTION), CX_ROWS["closing_bal"])

# ===========================================================================
# 5. OPS_CFADS
# ===========================================================================
ops = wb.create_sheet("Ops_CFADS")
ops.column_dimensions["A"].width = 30
ops.cell(row=1, column=1, value="OPERATIONS & CFADS").font = HEADER_FONT

OPS_ROWS = {
    "h2_production": 3, "revenue": 4, "fixed_opex": 5, "variable_opex": 6,
    "power_cost": 7, "water_cost": 8, "stack_capex": 9, "ebitda": 10, "cfads": 11,
}
label_row(ops, OPS_ROWS["h2_production"], "H2 Production", "kg")
label_row(ops, OPS_ROWS["revenue"], "Revenue", "$")
label_row(ops, OPS_ROWS["fixed_opex"], "Fixed Opex", "$")
label_row(ops, OPS_ROWS["variable_opex"], "Variable Opex", "$")
label_row(ops, OPS_ROWS["power_cost"], "Power Cost", "$")
label_row(ops, OPS_ROWS["water_cost"], "Water Cost", "$")
label_row(ops, OPS_ROWS["stack_capex"], "Stack Replacement Capex", "$")
label_row(ops, OPS_ROWS["ebitda"], "EBITDA", "$")
label_row(ops, OPS_ROWS["cfads"], "CFADS", "$")

QUARTER_HOURS = 8766 / 4  # average hours per quarter (365.25 days/yr)

for p in range(1, N_PERIODS + 1):
    c = pcol(p)
    ctl_c = f"Control!{c}"
    ops_flag = f"{ctl_c}{CTL_ROWS['ops_flag']}"
    years_since_cod = f"{ctl_c}{CTL_ROWS['years_since_cod']}"
    qtr_in_year = f"{ctl_c}{CTL_ROWS['qtr_in_year']}"

    ops[f"{c}{OPS_ROWS['h2_production']}"] = (
        f"={ops_flag}*ElectrolyzerCapacityMW*1000*{QUARTER_HOURS}*CapacityFactor/EfficiencykWhPerKg"
        f"*(1-DegradationRate)^MAX({years_since_cod}-1,0)"
    )
    h2 = f"{c}{OPS_ROWS['h2_production']}"
    ops[f"{c}{OPS_ROWS['revenue']}"] = f"={h2}*OfftakePrice"
    ops[f"{c}{OPS_ROWS['fixed_opex']}"] = f"={ops_flag}*FixedOpexAnnual/4"
    ops[f"{c}{OPS_ROWS['variable_opex']}"] = f"={h2}*VariableOpexPerKg"
    ops[f"{c}{OPS_ROWS['power_cost']}"] = f"={h2}*EfficiencykWhPerKg/1000*PowerPrice"
    ops[f"{c}{OPS_ROWS['water_cost']}"] = f"={h2}*WaterCostPerKg"
    ops[f"{c}{OPS_ROWS['stack_capex']}"] = (
        f"=IF(AND({ops_flag}=1,{years_since_cod}>0,MOD({years_since_cod},StackReplacementIntervalYrs)=0,"
        f"{qtr_in_year}=4),StackReplacementCost,0)"
    )
    ops[f"{c}{OPS_ROWS['ebitda']}"] = (
        f"={c}{OPS_ROWS['revenue']}-{c}{OPS_ROWS['fixed_opex']}-{c}{OPS_ROWS['variable_opex']}"
        f"-{c}{OPS_ROWS['power_cost']}-{c}{OPS_ROWS['water_cost']}"
    )
    ops[f"{c}{OPS_ROWS['cfads']}"] = f"={c}{OPS_ROWS['ebitda']}-{c}{OPS_ROWS['stack_capex']}"

row_range_name(wb, "H2Production_Row", "Ops_CFADS", OPS_ROWS["h2_production"])
row_range_name(wb, "CFADS_Row", "Ops_CFADS", OPS_ROWS["cfads"])

# ===========================================================================
# 6. DEBT_SIZING_DSCR (sculpting recursion, driven by GearingFactor)
# ===========================================================================
debt = wb.create_sheet("Debt_Sizing_DSCR")
debt.column_dimensions["A"].width = 30
debt.cell(row=1, column=1, value="DEBT SIZING & DSCR SCULPTING").font = HEADER_FONT

DEBT_ROWS = {
    "opening_bal": 4, "interest": 5, "cfads": 6, "principal": 7,
    "dscr_actual": 8, "closing_bal": 9,
}
label_row(debt, DEBT_ROWS["opening_bal"], "Opening Balance", "$")
label_row(debt, DEBT_ROWS["interest"], "Interest_t = r x Opening", "$")
label_row(debt, DEBT_ROWS["cfads"], "CFADS (link)", "$")
label_row(debt, DEBT_ROWS["principal"], "Principal_t (sculpted)", "$")
label_row(debt, DEBT_ROWS["dscr_actual"], "Actual DSCR", "x")
label_row(debt, DEBT_ROWS["closing_bal"], "Closing Balance", "$")

for p in range(1, N_PERIODS + 1):
    c = pcol(p)
    prev = pcol(p - 1) if p > 1 else None
    ctl_c = f"Control!{c}"
    is_cod = (p == COD_PERIOD)
    is_ops_amort = (COD_PERIOD <= p <= TENOR_END_PERIOD_DEFAULT)

    if is_cod:
        debt[f"{c}{DEBT_ROWS['opening_bal']}"] = "=ClosingBalanceAtCOD*GearingFactor"
    elif is_ops_amort:
        debt[f"{c}{DEBT_ROWS['opening_bal']}"] = f"={prev}{DEBT_ROWS['closing_bal']}"
    else:
        debt[f"{c}{DEBT_ROWS['opening_bal']}"] = 0

    ob = f"{c}{DEBT_ROWS['opening_bal']}"
    r_debt = f"{ctl_c}{CTL_ROWS['r_debt']}"
    debt[f"{c}{DEBT_ROWS['interest']}"] = f"={r_debt}*{ob}"
    debt[f"{c}{DEBT_ROWS['cfads']}"] = f"='Ops_CFADS'!{c}{OPS_ROWS['cfads']}"

    if is_ops_amort:
        cfads = f"{c}{DEBT_ROWS['cfads']}"
        interest = f"{c}{DEBT_ROWS['interest']}"
        debt[f"{c}{DEBT_ROWS['principal']}"] = (
            f"=IF({ob}=0,0,MIN({ob},MAX(0,{cfads}/TargetDSCR_Input-{interest})))"
        )
        debt[f"{c}{DEBT_ROWS['dscr_actual']}"] = (
            f"=IF(({interest}+{c}{DEBT_ROWS['principal']})=0,\"\","
            f"{cfads}/({interest}+{c}{DEBT_ROWS['principal']}))"
        )
        debt[f"{c}{DEBT_ROWS['closing_bal']}"] = f"={ob}-{c}{DEBT_ROWS['principal']}"
    else:
        debt[f"{c}{DEBT_ROWS['principal']}"] = 0
        debt[f"{c}{DEBT_ROWS['dscr_actual']}"] = ""
        debt[f"{c}{DEBT_ROWS['closing_bal']}"] = f"={ob}"

row_range_name(wb, "Interest_Row", "Debt_Sizing_DSCR", DEBT_ROWS["interest"])
row_range_name(wb, "Principal_Row", "Debt_Sizing_DSCR", DEBT_ROWS["principal"])
row_range_name(wb, "DSCR_Actual_Row", "Debt_Sizing_DSCR", DEBT_ROWS["dscr_actual"])
row_range_name(wb, "ClosingBalance_Debt_Row", "Debt_Sizing_DSCR", DEBT_ROWS["closing_bal"])
add_named_cell(wb, "EndingBalanceResidual", "Debt_Sizing_DSCR",
               pcol(TENOR_END_PERIOD_DEFAULT), DEBT_ROWS["closing_bal"])

# ===========================================================================
# 7. VBA_CONTROL (scalar cells VBA reads/writes)
# ===========================================================================
vctl = wb.create_sheet("VBA_Control")
vctl.column_dimensions["A"].width = 26
vctl.cell(row=1, column=1, value="VBA CONTROL (bisection scalars)").font = HEADER_FONT
vctl_rows = [
    ("GearingFactor", 1.0, "Debt quantum multiplier on ClosingBalanceAtCOD -- VBA bisects this"),
    ("Tolerance", 1000, "$ tolerance on EndingBalanceResidual for convergence"),
    ("MaxIterations", 100, "Bisection iteration cap"),
    ("SizingMethod", "GearingQuantum", "GearingQuantum (recommended) or DSCRMultiple"),
    ("ConvergenceStatus", "Not run", "Written by RunDebtSizing"),
    ("IterationsUsed", 0, "Written by RunDebtSizing"),
]
row = 3
for name, val, note in vctl_rows:
    vctl.cell(row=row, column=1, value=name)
    c = vctl.cell(row=row, column=2, value=val)
    c.fill = INPUT_FILL
    vctl.cell(row=row, column=4, value=note)
    add_named_cell(wb, name, "VBA_Control", "B", row)
    row += 1

# ===========================================================================
# 8. FINANCIALS_3STATEMENT (summary)
# ===========================================================================
fin = wb.create_sheet("Financials_3Statement")
fin.column_dimensions["A"].width = 30
fin.cell(row=1, column=1, value="FINANCIALS SUMMARY (unaudited, simplified)").font = HEADER_FONT
FIN_ROWS = {"ebitda": 3, "interest_exp": 4, "principal_rep": 5, "net_equity_cf": 6}
label_row(fin, FIN_ROWS["ebitda"], "EBITDA", "$")
label_row(fin, FIN_ROWS["interest_exp"], "Interest Expense", "$")
label_row(fin, FIN_ROWS["principal_rep"], "Principal Repayment", "$")
label_row(fin, FIN_ROWS["net_equity_cf"], "Net Cash Flow to Equity (post-debt)", "$")

for p in range(1, N_PERIODS + 1):
    c = pcol(p)
    fin[f"{c}{FIN_ROWS['ebitda']}"] = f"='Ops_CFADS'!{c}{OPS_ROWS['ebitda']}"
    fin[f"{c}{FIN_ROWS['interest_exp']}"] = f"='Debt_Sizing_DSCR'!{c}{DEBT_ROWS['interest']}"
    fin[f"{c}{FIN_ROWS['principal_rep']}"] = f"='Debt_Sizing_DSCR'!{c}{DEBT_ROWS['principal']}"
    fin[f"{c}{FIN_ROWS['net_equity_cf']}"] = (
        f"={c}{FIN_ROWS['ebitda']}-{c}{FIN_ROWS['interest_exp']}-{c}{FIN_ROWS['principal_rep']}"
    )

# ===========================================================================
# 9. LCOH (unlevered)
# ===========================================================================
lcoh = wb.create_sheet("LCOH")
lcoh.column_dimensions["A"].width = 30
lcoh.cell(row=1, column=1, value="LCOH (UNLEVERED)").font = HEADER_FONT
LCOH_ROWS = {
    "capex": 3, "opex": 4, "total_cost": 5, "disc_factor": 6,
    "disc_cost": 7, "volume": 8, "disc_volume": 9,
}
label_row(lcoh, LCOH_ROWS["capex"], "Capex_t", "$")
label_row(lcoh, LCOH_ROWS["opex"], "Opex_t (fixed+var+power+water+stack)", "$")
label_row(lcoh, LCOH_ROWS["total_cost"], "Total Cost_t", "$")
label_row(lcoh, LCOH_ROWS["disc_factor"], "Discount Factor (WACC)", "")
label_row(lcoh, LCOH_ROWS["disc_cost"], "Discounted Cost_t", "$")
label_row(lcoh, LCOH_ROWS["volume"], "H2 Volume_t", "kg")
label_row(lcoh, LCOH_ROWS["disc_volume"], "Discounted Volume_t", "kg")

for p in range(1, N_PERIODS + 1):
    c = pcol(p)
    lcoh[f"{c}{LCOH_ROWS['capex']}"] = f"='Capex_Construction'!{c}{CX_ROWS['capex_draw']}"
    lcoh[f"{c}{LCOH_ROWS['opex']}"] = (
        f"='Ops_CFADS'!{c}{OPS_ROWS['fixed_opex']}+'Ops_CFADS'!{c}{OPS_ROWS['variable_opex']}"
        f"+'Ops_CFADS'!{c}{OPS_ROWS['power_cost']}+'Ops_CFADS'!{c}{OPS_ROWS['water_cost']}"
        f"+'Ops_CFADS'!{c}{OPS_ROWS['stack_capex']}"
    )
    lcoh[f"{c}{LCOH_ROWS['total_cost']}"] = f"={c}{LCOH_ROWS['capex']}+{c}{LCOH_ROWS['opex']}"
    lcoh[f"{c}{LCOH_ROWS['disc_factor']}"] = f"='Control'!{c}{CTL_ROWS['disc_factor']}"
    lcoh[f"{c}{LCOH_ROWS['disc_cost']}"] = f"={c}{LCOH_ROWS['total_cost']}*{c}{LCOH_ROWS['disc_factor']}"
    lcoh[f"{c}{LCOH_ROWS['volume']}"] = f"='Ops_CFADS'!{c}{OPS_ROWS['h2_production']}"
    lcoh[f"{c}{LCOH_ROWS['disc_volume']}"] = f"={c}{LCOH_ROWS['volume']}*{c}{LCOH_ROWS['disc_factor']}"

total_row = 12
lcoh.cell(row=total_row, column=1, value="TOTALS").font = BOLD
lcoh.cell(row=total_row, column=2, value=f"=SUM({pcol(1)}{LCOH_ROWS['disc_cost']}:{pcol(N_PERIODS)}{LCOH_ROWS['disc_cost']})")
lcoh.cell(row=total_row, column=3, value=f"=SUM({pcol(1)}{LCOH_ROWS['disc_volume']}:{pcol(N_PERIODS)}{LCOH_ROWS['disc_volume']})")
lcoh.cell(row=total_row - 1, column=2, value="Discounted Cost NPV")
lcoh.cell(row=total_row - 1, column=3, value="Discounted Volume NPV")
add_named_cell(wb, "LCOH_CostNPV", "LCOH", "B", total_row)
add_named_cell(wb, "LCOH_VolumeNPV", "LCOH", "C", total_row)

lcoh.cell(row=total_row + 2, column=1, value="LCOH ($/kg)").font = BOLD
lcoh_cell = lcoh.cell(row=total_row + 2, column=2, value="=LCOH_CostNPV/LCOH_VolumeNPV")
lcoh_cell.font = Font(bold=True, size=14)
lcoh_cell.fill = LINK_FILL
add_named_cell(wb, "LCOH_USD_per_kg", "LCOH", "B", total_row + 2)

# ===========================================================================
# 10. RETURNS_SENSITIVITIES
# ===========================================================================
ret = wb.create_sheet("Returns_Sensitivities")
ret.column_dimensions["A"].width = 30
ret.cell(row=1, column=1, value="RETURNS & SENSITIVITIES").font = HEADER_FONT
RET_ROWS = {"equity_cf": 3}
label_row(ret, RET_ROWS["equity_cf"], "Equity Cash Flow", "$")

for p in range(1, N_PERIODS + 1):
    c = pcol(p)
    if p <= N_CONSTRUCTION:
        ret[f"{c}{RET_ROWS['equity_cf']}"] = (
            f"=-('Capex_Construction'!{c}{CX_ROWS['equity_draw']}+'Capex_Construction'!{c}{CX_ROWS['cash_interest']})"
        )
    else:
        ret[f"{c}{RET_ROWS['equity_cf']}"] = f"='Financials_3Statement'!{c}{FIN_ROWS['net_equity_cf']}"

row_range_name(wb, "EquityCashFlow_Row", "Returns_Sensitivities", RET_ROWS["equity_cf"])

irr_row = 6
ret.cell(row=irr_row, column=1, value="Equity IRR (quarterly, annualized)").font = BOLD
ret.cell(row=irr_row, column=2,
    value=f"=(1+IRR({pcol(1)}{RET_ROWS['equity_cf']}:{pcol(N_PERIODS)}{RET_ROWS['equity_cf']}))^4-1")
add_named_cell(wb, "EquityIRR", "Returns_Sensitivities", "B", irr_row)

ret.cell(row=irr_row + 2, column=1, value="Min DSCR (operating period)").font = BOLD
ret.cell(row=irr_row + 2, column=2, value="=MIN(DSCR_Actual_Row)")
ret.cell(row=irr_row + 3, column=1, value="Avg DSCR (operating period)").font = BOLD
ret.cell(row=irr_row + 3, column=2, value="=AVERAGE(DSCR_Actual_Row)")

# Sensitivity grid skeleton for mod_Sensitivity.bas to populate
sens_header_row = irr_row + 6
ret.cell(row=sens_header_row, column=1, value="SENSITIVITY GRID (populated by RunSensitivityGrid macro)").font = BOLD
headers = ["Scenario", "Capex %", "Opex %", "WACC bps", "Utilization %", "LCOH", "Equity IRR", "Min DSCR"]
for i, h in enumerate(headers):
    ret.cell(row=sens_header_row + 1, column=1 + i, value=h).font = BOLD
# (Scenario name, Capex delta %, Opex delta % [applied to Fixed+Variable Opex],
#  WACC delta in bps, Utilization/capacity-factor delta %)
scenarios = [
    ("Base", 0.00, 0.00, 0, 0.00),
    ("Capex +10%", 0.10, 0.00, 0, 0.00),
    ("Capex -10%", -0.10, 0.00, 0, 0.00),
    ("Opex +10%", 0.00, 0.10, 0, 0.00),
    ("Opex -10%", 0.00, -0.10, 0, 0.00),
    ("WACC +100bps", 0.00, 0.00, 100, 0.00),
    ("WACC -100bps", 0.00, 0.00, -100, 0.00),
    ("Utilization +5%", 0.00, 0.00, 0, 0.05),
    ("Utilization -5%", 0.00, 0.00, 0, -0.05),
]
for i, (name, dcapex, dopex, dwacc_bps, dutil) in enumerate(scenarios):
    rr = sens_header_row + 2 + i
    ret.cell(row=rr, column=1, value=name)
    ret.cell(row=rr, column=2, value=dcapex)
    ret.cell(row=rr, column=3, value=dopex)
    ret.cell(row=rr, column=4, value=dwacc_bps)
    ret.cell(row=rr, column=5, value=dutil)

SENS_FIRST_ROW = sens_header_row + 2
SENS_LAST_ROW = sens_header_row + 1 + len(scenarios)
add_named_cell(wb, "SensitivityGrid_FirstRow", "Returns_Sensitivities", "A", SENS_FIRST_ROW)
add_named_cell(wb, "SensitivityGrid_LastRow", "Returns_Sensitivities", "A", SENS_LAST_ROW)

# ===========================================================================
# 11. DASHBOARD
# ===========================================================================
dash = wb.create_sheet("Dashboard")
dash.column_dimensions["A"].width = 30
dash.cell(row=1, column=1, value="DASHBOARD").font = Font(bold=True, size=16)

kpis = [
    ("LCOH ($/kg)", "=LCOH_USD_per_kg"),
    ("Equity IRR", "=EquityIRR"),
    ("Min DSCR", "='Returns_Sensitivities'!B8"),
    ("Avg DSCR", "='Returns_Sensitivities'!B9"),
    ("Gearing Factor (solved)", "=GearingFactor"),
    ("Ending Balance Residual", "=EndingBalanceResidual"),
    ("Convergence Status", "=ConvergenceStatus"),
    ("Iterations Used", "=IterationsUsed"),
]
row = 3
for label, formula in kpis:
    dash.cell(row=row, column=1, value=label).font = BOLD
    c = dash.cell(row=row, column=2, value=formula)
    c.fill = LINK_FILL
    row += 1

row += 2
dash.cell(row=row, column=1,
    value="[Button: Run Debt Sizing]  -- after importing vba/mod_DebtSizing.bas, "
          "assign macro 'RunDebtSizing' to a Form Control button here.")
row += 1
dash.cell(row=row, column=1,
    value="[Button: Run Sensitivity Grid]  -- assign macro 'RunSensitivityGrid' "
          "from vba/mod_Sensitivity.bas here.")

# ===========================================================================
# Save
# ===========================================================================
import os
out_path = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "HydrogenLCOH_Model.xlsx")
wb.calculation.calcMode = "manual"
wb.calculation.iterate = False
wb.save(out_path)
print(f"Wrote {out_path}")
print(f"N_PERIODS={N_PERIODS}, COD_PERIOD={COD_PERIOD}, TENOR_END_PERIOD={TENOR_END_PERIOD_DEFAULT}")
