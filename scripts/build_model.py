"""
Green Hydrogen LCOH — bankable project-finance model generator (FAST/F1F9 style).

Builds HydrogenLCOH_Model.xlsx to professional project-finance conventions:

  * Time runs across columns, one line item per row, a SINGLE formula copied
    unbroken across every period column (FAST's cardinal rule). Timing is
    data-driven via flag & counter rows on the Time sheet, never by
    branching the formula per column.
  * A period-0 column holds opening balances so every corkscrew
    (opening = prior closing) is one identical formula including period 1.
  * No live circular references (iterative calc stays OFF): closed-form IDC,
    all interest on opening balances, tax paid one period in arrears, DSRA
    interest on opening balance. The only "solve" is the single scalar
    GearingFactor (VBA bisection / native Goal Seek), done outside recalc.
  * Semi-annual periods: 2yr construction + 25yr operations = 54 periods.

Sheets (mnemonic): CV Cover, IN Inputs, TM Time, CX Construction/Capex,
PR Production/Revenue, OP Opex, WC Working Capital, TX Tax, CF CFADS,
DB Debt, DS DSRA, WF Waterfall, CR Cover Ratios, FS Financial Statements,
RT Returns, LC LCOH, CK Checks, OUT Dashboard.

Run: python3 scripts/build_model.py    (requires openpyxl)
Output: HydrogenLCOH_Model.xlsx
See /root/.claude/plans/plan-first-can-u-playful-beaver.md (REBUILD v2).
"""
import os
from openpyxl import Workbook
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.utils import get_column_letter
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.worksheet.datavalidation import DataValidation

# --------------------------------------------------------------------------
# Timeline geometry
# --------------------------------------------------------------------------
PPY = 2                      # periods per year (semi-annual)
PERIOD_YEARS = 1.0 / PPY     # 0.5
CONSTR_YEARS = 2
OPS_YEARS = 25
CONSTR_PERIODS = CONSTR_YEARS * PPY      # 4
OPS_PERIODS = OPS_YEARS * PPY            # 50
N = CONSTR_PERIODS + OPS_PERIODS         # 54
COD_PERIOD = CONSTR_PERIODS + 1          # 5 (first operating period)
DEBT_TENOR_YEARS = 15
DEBT_PERIODS = DEBT_TENOR_YEARS * PPY    # 30
TENOR_END_PERIOD = CONSTR_PERIODS + DEBT_PERIODS  # 34
HOURS_PER_PERIOD = 8766 * PERIOD_YEARS   # 4383

# --------------------------------------------------------------------------
# Column layout: A label | C units | E = period 0 | F.. = periods 1..N
# --------------------------------------------------------------------------
LABEL_COL = 1     # A
UNIT_COL = 3      # C
P0_COL = 5        # E  (period 0)


def col(p):
    """Column letter for period p (p=0 -> E)."""
    return get_column_letter(P0_COL + p)


LAST_COL = col(N)

# --------------------------------------------------------------------------
# Styling
# --------------------------------------------------------------------------
INPUT_FILL = PatternFill("solid", fgColor="DDEBF7")   # blue = hardcode
CHECK_OK_FILL = PatternFill("solid", fgColor="C6EFCE")
HEADER_FILL = PatternFill("solid", fgColor="1F3864")
SECTION_FILL = PatternFill("solid", fgColor="D9E1F2")
HEADER_FONT = Font(bold=True, size=12, color="FFFFFF")
TITLE_FONT = Font(bold=True, size=14)
SECTION_FONT = Font(bold=True, size=10)
BOLD = Font(bold=True)
GREY = Font(color="808080", italic=True)
THIN = Side(style="thin", color="BFBFBF")

wb = Workbook()
wb.remove(wb.active)
_defined = {}


import re as _re
def _absolutize(ref):
    """Turn 'E5' / 'E5:F5' into '$E$5' / '$E$5:$F$5' so defined names are
    absolute (a relative reference in a defined name shifts with the
    referencing cell — the classic #DIV/0 / wrong-cell bug)."""
    return _re.sub(r"([A-Z]+)(\d+)", r"$\1$\2", ref)


def add_name(name, sheet, cell):
    _defined[name] = f"'{sheet}'!{_absolutize(cell)}"


def finalize_names():
    for name, ref in _defined.items():
        wb.defined_names[name] = DefinedName(name, attr_text=ref)


def new_sheet(mnemonic, title):
    ws = wb.create_sheet(mnemonic)
    ws.column_dimensions["A"].width = 40
    ws.column_dimensions["B"].width = 2
    ws.column_dimensions["C"].width = 12
    ws.column_dimensions["D"].width = 2
    ws.sheet_view.showGridLines = False
    c = ws.cell(row=1, column=1, value=title)
    c.font = HEADER_FONT
    for cc in range(1, P0_COL + N + 1):
        ws.cell(row=1, column=cc).fill = HEADER_FILL
    ws.freeze_panes = ws.cell(row=3, column=P0_COL)
    return ws


def label(ws, row, text, unit="", bold=False, section=False):
    c = ws.cell(row=row, column=LABEL_COL, value=text)
    if section:
        c.font = SECTION_FONT
        for cc in range(1, P0_COL + N + 1):
            ws.cell(row=row, column=cc).fill = SECTION_FILL
    elif bold:
        c.font = BOLD
    if unit:
        ws.cell(row=row, column=UNIT_COL, value=unit).font = GREY


def series(ws, row, fml, p0=None, num="#,##0", first=1, last=N):
    """Write one formula (fml(p)) across period columns first..last.
    p0 sets the period-0 seed (value or formula)."""
    if p0 is not None:
        cell = ws.cell(row=row, column=P0_COL, value=p0)
        cell.number_format = num
    for p in range(first, last + 1):
        cell = ws.cell(row=row, column=P0_COL + p, value=fml(p))
        cell.number_format = num


def input_cell(ws, row, labeltext, value, unit="", note="", num="#,##0.####"):
    label(ws, row, labeltext, unit)
    c = ws.cell(row=row, column=P0_COL, value=value)
    c.fill = INPUT_FILL
    c.number_format = num
    if note:
        ws.cell(row=row, column=P0_COL + 2, value=note).font = GREY
    return f"'{ws.title}'!{get_column_letter(P0_COL)}{row}"


# ==========================================================================
# IN — Inputs
# ==========================================================================
IN = new_sheet("IN", "INPUTS  —  green hydrogen LCOH model (all blue cells are hardcoded assumptions)")
r = 3
def sec(text):
    global r
    label(IN, r, text, section=True); r += 1
def inp(name, text, val, unit="", note="", num="#,##0.####"):
    global r
    ref = input_cell(IN, r, text, val, unit, note, num)
    add_name(name, IN.title, f"{get_column_letter(P0_COL)}{r}")
    r += 1
    return ref

sec("Model dates & timing")
inp("ModelStartDate", "Financial close date", "=DATE(2026,1,1)", "date", num="yyyy-mm-dd")
inp("PeriodsPerYear", "Periods per year", PPY, "#")
inp("PeriodYears", "Years per period", PERIOD_YEARS, "yrs")
inp("ConstructionPeriods", "Construction periods", CONSTR_PERIODS, "#")
inp("OperationsPeriods", "Operations periods", OPS_PERIODS, "#")
inp("DebtTenorPeriods", "Debt tenor (periods, post-COD)", DEBT_PERIODS, "#")

sec("Technical & production")
inp("CapacityMW", "Electrolyser capacity", 100, "MW")
inp("Efficiency", "Electrolyser efficiency", 50, "kWh/kg")
inp("CapacityFactor", "Capacity factor", 0.90, "%", num="0.0%")
inp("Availability", "Availability", 0.97, "%", num="0.0%")
inp("Degradation", "Annual degradation", 0.010, "%/yr", num="0.0%")

sec("Revenue")
inp("OfftakePrice", "H2 offtake price (base)", 7.00, "$/kg", num="#,##0.00")
inp("PriceEscalation", "Offtake price escalation", 0.02, "%/yr", num="0.0%")

sec("Operating costs (base, real)")
inp("PowerPrice", "Power price", 35, "$/MWh")
inp("WaterCostPerKg", "Water cost", 0.03, "$/kg", num="#,##0.00")
inp("FixedOMAnnual", "Fixed O&M", 6000000, "$/yr")
inp("VarOMPerKg", "Variable O&M", 0.15, "$/kg", num="#,##0.00")
inp("InsuranceAnnual", "Insurance", 2500000, "$/yr")
inp("StackCost", "Stack replacement cost", 30000000, "$")
inp("StackIntervalYears", "Stack replacement interval", 8, "yrs")
inp("OpexEscalation", "Opex escalation (inflation)", 0.02, "%/yr", num="0.0%")

sec("Capex & construction")
inp("TotalCapex", "Total capex (EPC+BoP+owner's)", 400000000, "$")
inp("ConstructionGearing", "Debt-funded share of capex", 0.70, "%", num="0.0%")

sec("Financing")
inp("WACC", "WACC (nominal)", 0.080, "%", num="0.00%")
inp("CostOfEquity", "Cost of equity (nominal)", 0.120, "%", num="0.00%")
inp("CostOfDebt", "Cost of debt (nominal)", 0.060, "%", num="0.00%")
inp("Inflation", "Long-run inflation", 0.020, "%/yr", num="0.00%")
inp("IDC_CashFundedPct_f", "IDC cash-funded fraction (f)", 0.30, "0..1",
    "0 = fully capitalised, 1 = fully cash-funded", num="0.0%")

sec("Debt sizing")
inp("CovenantDSCR", "Lender covenant DSCR (minimum)", 1.30, "x", num="0.00")
inp("MaxGearing", "Max gearing (debt / capex)", 0.75, "%", num="0.0%")

sec("DSRA")
inp("DSRA_Months", "DSRA cover", 6, "months")

sec("Tax")
inp("TaxRate", "Corporate tax rate", 0.25, "%", num="0.0%")
inp("TaxDepLifeYears", "Tax depreciation life (SL)", 20, "yrs")

sec("Working capital")
inp("DebtorDays", "Debtor days", 45, "days")
inp("CreditorDays", "Creditor days", 30, "days")

# Debt sizing (F1F9 style): gearing is fixed at the construction facility level
# (GearingFactor = 1.0, term debt = construction debt, so no refinancing gap),
# and the flat sculpting DSCR is SOLVED so the debt amortises exactly over the
# tenor. The solved sculpting DSCR sits above the lender covenant, which the
# checks confirm. RunDebtSizing / Goal Seek re-solves TargetDSCR_Input after any
# input change (set EndingBalanceResidual to 0 by changing that cell).
sec("Solver control (set by RunDebtSizing macro or Goal Seek)")
c = IN.cell(row=r, column=LABEL_COL, value="Sculpting DSCR (solved)")
IN.cell(row=r, column=UNIT_COL, value="x").font = GREY
tc = IN.cell(row=r, column=P0_COL, value=2.196165); tc.fill = INPUT_FILL; tc.number_format = "0.0000"
add_name("TargetDSCR_Input", IN.title, f"{get_column_letter(P0_COL)}{r}")
r += 1
c = IN.cell(row=r, column=LABEL_COL, value="Gearing factor (term/construction debt)")
IN.cell(row=r, column=UNIT_COL, value="x").font = GREY
gc = IN.cell(row=r, column=P0_COL, value=1.0); gc.fill = INPUT_FILL; gc.number_format = "0.0000"
add_name("GearingFactor", IN.title, f"{get_column_letter(P0_COL)}{r}")
r += 1
c = IN.cell(row=r, column=LABEL_COL, value="Convergence status");
IN.cell(row=r, column=P0_COL, value="Not run")
add_name("ConvergenceStatus", IN.title, f"{get_column_letter(P0_COL)}{r}"); r += 1
c = IN.cell(row=r, column=LABEL_COL, value="Iterations used")
IN.cell(row=r, column=P0_COL, value=0)
add_name("IterationsUsed", IN.title, f"{get_column_letter(P0_COL)}{r}"); r += 1
c = IN.cell(row=r, column=LABEL_COL, value="Sizing tolerance ($)")
IN.cell(row=r, column=P0_COL, value=1000); add_name("Tolerance", IN.title, f"{get_column_letter(P0_COL)}{r}"); r += 1
c = IN.cell(row=r, column=LABEL_COL, value="Max iterations")
IN.cell(row=r, column=P0_COL, value=100); add_name("MaxIterations", IN.title, f"{get_column_letter(P0_COL)}{r}"); r += 1


# ==========================================================================
# TM — Time
# ==========================================================================
TM = new_sheet("TM", "TIME  —  master timeline, flags, counters, indices, discount factors")
TMR = dict(period=3, pstart=4, pend=5, yearfrac=6, constr=7, ops=8, debt=9,
           firstop=10, last=11, opsyears=12, inflidx=13, priceidx=14,
           rdebt=15, rwacc=16, df_nom=17, df_real=18, depflag=19)
label(TM, TMR["period"], "Period number", "#")
label(TM, TMR["pstart"], "Period start", "date")
label(TM, TMR["pend"], "Period end", "date")
label(TM, TMR["yearfrac"], "Year fraction", "yrs")
label(TM, TMR["constr"], "Construction flag", "0/1")
label(TM, TMR["ops"], "Operations flag", "0/1")
label(TM, TMR["debt"], "Debt-period flag", "0/1")
label(TM, TMR["firstop"], "First operating period", "0/1")
label(TM, TMR["last"], "Last period flag", "0/1")
label(TM, TMR["opsyears"], "Operating years elapsed (bop)", "yrs")
label(TM, TMR["inflidx"], "Inflation / opex index", "idx")
label(TM, TMR["priceidx"], "Offtake price index", "idx")
label(TM, TMR["rdebt"], "Periodic debt rate", "%")
label(TM, TMR["rwacc"], "Periodic WACC (nominal)", "%")
label(TM, TMR["df_nom"], "Discount factor (nominal WACC)", "x")
label(TM, TMR["df_real"], "Discount factor (real WACC)", "x")
label(TM, TMR["depflag"], "Depreciation flag", "0/1")

series(TM, TMR["period"], lambda p: p, p0=0, num="#,##0")
series(TM, TMR["pstart"], lambda p: f"=EDATE(ModelStartDate,({col(p)}{TMR['period']}-1)*12/PeriodsPerYear)", p0="=ModelStartDate", num="yyyy-mm")
series(TM, TMR["pend"], lambda p: f"=EDATE(ModelStartDate,{col(p)}{TMR['period']}*12/PeriodsPerYear)", p0="=ModelStartDate", num="yyyy-mm")
series(TM, TMR["yearfrac"], lambda p: f"=({col(p)}{TMR['pend']}-{col(p)}{TMR['pstart']})/365.25", p0=0, num="0.00")
series(TM, TMR["constr"], lambda p: f"=IF(AND({col(p)}{TMR['period']}>=1,{col(p)}{TMR['period']}<=ConstructionPeriods),1,0)", p0=0, num="0")
series(TM, TMR["ops"], lambda p: f"=IF({col(p)}{TMR['period']}>ConstructionPeriods,1,0)", p0=0, num="0")
series(TM, TMR["debt"], lambda p: f"=IF(AND({col(p)}{TMR['period']}>ConstructionPeriods,{col(p)}{TMR['period']}<=ConstructionPeriods+DebtTenorPeriods),1,0)", p0=0, num="0")
series(TM, TMR["firstop"], lambda p: f"=IF({col(p)}{TMR['period']}=ConstructionPeriods+1,1,0)", p0=0, num="0")
series(TM, TMR["last"], lambda p: f"=IF({col(p)}{TMR['period']}=ConstructionPeriods+OperationsPeriods,1,0)", p0=0, num="0")
series(TM, TMR["opsyears"], lambda p: f"=MAX(0,({col(p)}{TMR['period']}-ConstructionPeriods-1))*PeriodYears", p0=0, num="0.0")
series(TM, TMR["inflidx"], lambda p: f"={col(p-1)}{TMR['inflidx']}*(1+OpexEscalation)^{col(p)}{TMR['yearfrac']}", p0=1, num="0.000")
series(TM, TMR["priceidx"], lambda p: f"={col(p-1)}{TMR['priceidx']}*(1+PriceEscalation)^{col(p)}{TMR['yearfrac']}", p0=1, num="0.000")
series(TM, TMR["rdebt"], lambda p: "=(1+CostOfDebt)^(1/PeriodsPerYear)-1", p0=0, num="0.000%")
series(TM, TMR["rwacc"], lambda p: "=(1+WACC)^(1/PeriodsPerYear)-1", p0=0, num="0.000%")
series(TM, TMR["df_nom"], lambda p: f"={col(p-1)}{TMR['df_nom']}/(1+{col(p)}{TMR['rwacc']})", p0=1, num="0.0000")
series(TM, TMR["df_real"], lambda p: f"={col(p-1)}{TMR['df_real']}/(1+((1+WACC)/(1+Inflation))^(1/PeriodsPerYear)-1)", p0=1, num="0.0000")
series(TM, TMR["depflag"], lambda p: f"=IF(AND({col(p)}{TMR['ops']}=1,{col(p)}{TMR['opsyears']}<TaxDepLifeYears+PeriodYears),1,0)", p0=0, num="0")

def TMc(key, p):
    return f"TM!{col(p)}{TMR[key]}"


# ==========================================================================
# CX — Construction & Capex (closed-form IDC)
# ==========================================================================
CX = new_sheet("CX", "CONSTRUCTION & CAPEX  —  drawdown, sources & uses, closed-form IDC")
CXR = dict(capex=3, debtdraw=4, eqdraw=5, dopen=6, interest=7, cashint=8,
           capint=9, dclose=10, cumcapex=11, ppe_add=12)
label(CX, CXR["capex"], "Capex drawn (straight-line)", "$")
label(CX, CXR["debtdraw"], "  Debt drawdown", "$")
label(CX, CXR["eqdraw"], "  Equity drawdown", "$")
label(CX, CXR["dopen"], "Construction debt — opening", "$")
label(CX, CXR["interest"], "  IDC interest (closed form)", "$")
label(CX, CXR["cashint"], "  IDC cash-funded (f·I)", "$")
label(CX, CXR["capint"], "  IDC capitalised ((1-f)·I)", "$")
label(CX, CXR["dclose"], "Construction debt — closing", "$")
label(CX, CXR["cumcapex"], "Cumulative capex + capitalised IDC", "$")
label(CX, CXR["ppe_add"], "PP&E addition (capex + cap. IDC)", "$")
CX.cell(row=13, column=1, value="IDC (mid-period): I = r·(open + draw/2) / (1 − r·(1−f)/2)  — algebraically isolated, non-circular").font = GREY

series(CX, CXR["capex"], lambda p: f"={TMc('constr',p)}*TotalCapex/ConstructionPeriods", p0=0)
series(CX, CXR["debtdraw"], lambda p: f"={col(p)}{CXR['capex']}*ConstructionGearing", p0=0)
series(CX, CXR["eqdraw"], lambda p: f"={col(p)}{CXR['capex']}*(1-ConstructionGearing)", p0=0)
series(CX, CXR["dopen"], lambda p: f"={col(p-1)}{CXR['dclose']}", p0=0)
series(CX, CXR["interest"],
       lambda p: f"=({TMc('rdebt',p)}*({col(p)}{CXR['dopen']}+{col(p)}{CXR['debtdraw']}/2))/(1-{TMc('rdebt',p)}*(1-IDC_CashFundedPct_f)/2)*{TMc('constr',p)}", p0=0)
series(CX, CXR["cashint"], lambda p: f"=IDC_CashFundedPct_f*{col(p)}{CXR['interest']}", p0=0)
series(CX, CXR["capint"], lambda p: f"=(1-IDC_CashFundedPct_f)*{col(p)}{CXR['interest']}", p0=0)
series(CX, CXR["dclose"], lambda p: f"={col(p)}{CXR['dopen']}+{col(p)}{CXR['debtdraw']}+{col(p)}{CXR['capint']}", p0=0)
series(CX, CXR["cumcapex"], lambda p: f"={col(p-1)}{CXR['cumcapex']}+{col(p)}{CXR['capex']}+{col(p)}{CXR['capint']}", p0=0)
series(CX, CXR["ppe_add"], lambda p: f"={col(p)}{CXR['capex']}+{col(p)}{CXR['capint']}", p0=0)

# Debt outstanding at COD (base the sculpting scales via GearingFactor)
add_name("DebtAtCOD", CX.title, f"{col(CONSTR_PERIODS)}{CXR['dclose']}")
add_name("TotalEquityConstruction", CX.title, f"{LAST_COL}{CXR['eqdraw']}")  # not exact; recomputed in checks


# ==========================================================================
# PR — Production & Revenue
# ==========================================================================
PR = new_sheet("PR", "PRODUCTION & REVENUE")
PRR = dict(gross=3, degr=4, h2=5, price=6, rev=7)
label(PR, PRR["gross"], "Gross output (nameplate × CF × avail)", "kg")
label(PR, PRR["degr"], "Degradation factor", "x")
label(PR, PRR["h2"], "H2 produced", "kg")
label(PR, PRR["price"], "Offtake price (escalated)", "$/kg")
label(PR, PRR["rev"], "Revenue", "$")
series(PR, PRR["gross"], lambda p: f"={TMc('ops',p)}*CapacityMW*1000*{HOURS_PER_PERIOD}*CapacityFactor*Availability/Efficiency", p0=0)
series(PR, PRR["degr"], lambda p: f"=(1-Degradation)^{TMc('opsyears',p)}", p0=1, num="0.000")
series(PR, PRR["h2"], lambda p: f"={col(p)}{PRR['gross']}*{col(p)}{PRR['degr']}", p0=0)
series(PR, PRR["price"], lambda p: f"=OfftakePrice*{TMc('priceidx',p)}", p0=0, num="#,##0.00")
series(PR, PRR["rev"], lambda p: f"={col(p)}{PRR['h2']}*{col(p)}{PRR['price']}", p0=0)


# ==========================================================================
# OP — Operating costs
# ==========================================================================
OP = new_sheet("OP", "OPERATING COSTS  (escalated by inflation index)")
OPR = dict(power=3, water=4, fixed=5, var=6, ins=7, stack=8, totopex=9, ebitda=10)
label(OP, OPR["power"], "Power cost", "$")
label(OP, OPR["water"], "Water cost", "$")
label(OP, OPR["fixed"], "Fixed O&M", "$")
label(OP, OPR["var"], "Variable O&M", "$")
label(OP, OPR["ins"], "Insurance", "$")
label(OP, OPR["stack"], "Stack replacement", "$")
label(OP, OPR["totopex"], "Total cash opex (excl. stack)", "$", bold=True)
label(OP, OPR["ebitda"], "EBITDA", "$", bold=True)
series(OP, OPR["power"], lambda p: f"=PR!{col(p)}{PRR['h2']}*Efficiency/1000*PowerPrice*{TMc('inflidx',p)}", p0=0)
series(OP, OPR["water"], lambda p: f"=PR!{col(p)}{PRR['h2']}*WaterCostPerKg*{TMc('inflidx',p)}", p0=0)
series(OP, OPR["fixed"], lambda p: f"={TMc('ops',p)}*FixedOMAnnual*PeriodYears*{TMc('inflidx',p)}", p0=0)
series(OP, OPR["var"], lambda p: f"=PR!{col(p)}{PRR['h2']}*VarOMPerKg*{TMc('inflidx',p)}", p0=0)
series(OP, OPR["ins"], lambda p: f"={TMc('ops',p)}*InsuranceAnnual*PeriodYears*{TMc('inflidx',p)}", p0=0)
# Stack replacement at each StackIntervalYears anniversary within operations.
# opsyears (beginning-of-period) lands on an exact integer once per year, so
# MOD(opsyears, interval)=0 fires in exactly one period per replacement.
def stack_fml(p):
    opsyears = TMc('opsyears', p)
    return (f"=IF(AND({TMc('ops',p)}=1,{opsyears}>0,"
            f"MOD(ROUND({opsyears},4),StackIntervalYears)=0),"
            f"StackCost*{TMc('inflidx',p)},0)")
series(OP, OPR["stack"], stack_fml, p0=0)
series(OP, OPR["totopex"], lambda p: f"={col(p)}{OPR['power']}+{col(p)}{OPR['water']}+{col(p)}{OPR['fixed']}+{col(p)}{OPR['var']}+{col(p)}{OPR['ins']}", p0=0)
series(OP, OPR["ebitda"], lambda p: f"=PR!{col(p)}{PRR['rev']}-{col(p)}{OPR['totopex']}", p0=0)


# ==========================================================================
# WC — Working capital
# ==========================================================================
WC = new_sheet("WC", "WORKING CAPITAL")
WCR = dict(debtors=3, creditors=4, nwc=5, move=6)
label(WC, WCR["debtors"], "Debtors (receivables)", "$")
label(WC, WCR["creditors"], "Creditors (payables)", "$")
label(WC, WCR["nwc"], "Net working capital", "$")
label(WC, WCR["move"], "Increase in NWC (cash outflow)", "$")
series(WC, WCR["debtors"], lambda p: f"=PR!{col(p)}{PRR['rev']}/(365.25*PeriodYears)*DebtorDays", p0=0)
series(WC, WCR["creditors"], lambda p: f"=OP!{col(p)}{OPR['totopex']}/(365.25*PeriodYears)*CreditorDays", p0=0)
series(WC, WCR["nwc"], lambda p: f"={col(p)}{WCR['debtors']}-{col(p)}{WCR['creditors']}", p0=0)
series(WC, WCR["move"], lambda p: f"={col(p)}{WCR['nwc']}-{col(p-1)}{WCR['nwc']}", p0=0)


# ==========================================================================
# TX — Tax (levered tax for FS; unlevered tax for LCOH)
# ==========================================================================
TX = new_sheet("TX", "TAX  —  depreciation, losses, charge (paid one period in arrears)")
TXR = dict(depbase=3, depopen=4, dep=5, depclose=6,
           ebit=7, lev_lossopen=8, lev_ti=9, lev_lossclose=10, lev_charge=11, lev_paid=12,
           unl_lossopen=13, unl_ti=14, unl_lossclose=15, unl_charge=16)
label(TX, TXR["depbase"], "Depreciation base added", "$")
label(TX, TXR["depopen"], "Remaining depreciable — opening", "$")
label(TX, TXR["dep"], "Tax depreciation (SL)", "$")
label(TX, TXR["depclose"], "Remaining depreciable — closing", "$")
label(TX, TXR["ebit"], "EBIT (EBITDA − depreciation)", "$")
label(TX, TXR["lev_lossopen"], "Tax losses b/f (levered)", "$")
label(TX, TXR["lev_ti"], "Taxable income after interest (levered)", "$")
label(TX, TXR["lev_lossclose"], "Tax losses c/f (levered)", "$")
label(TX, TXR["lev_charge"], "Tax charge (levered)", "$", bold=True)
label(TX, TXR["lev_paid"], "Tax paid (1-period arrears)", "$", bold=True)
label(TX, TXR["unl_lossopen"], "Tax losses b/f (unlevered)", "$")
label(TX, TXR["unl_ti"], "Taxable income (unlevered)", "$")
label(TX, TXR["unl_lossclose"], "Tax losses c/f (unlevered)", "$")
label(TX, TXR["unl_charge"], "Tax charge (unlevered, for LCOH)", "$", bold=True)

series(TX, TXR["depbase"], lambda p: f"=CX!{col(p)}{CXR['ppe_add']}", p0=0)
series(TX, TXR["depopen"], lambda p: f"={col(p-1)}{TXR['depclose']}+{col(p)}{TXR['depbase']}", p0=0)
series(TX, TXR["dep"], lambda p: f"=MIN({col(p)}{TXR['depopen']},{TMc('depflag',p)}*CX!{LAST_COL}{CXR['cumcapex']}/TaxDepLifeYears*PeriodYears)", p0=0)
series(TX, TXR["depclose"], lambda p: f"={col(p)}{TXR['depopen']}-{col(p)}{TXR['dep']}", p0=0)
series(TX, TXR["ebit"], lambda p: f"=OP!{col(p)}{OPR['ebitda']}-{col(p)}{TXR['dep']}", p0=0)
# Levered
series(TX, TXR["lev_lossopen"], lambda p: f"={col(p-1)}{TXR['lev_lossclose']}", p0=0)
series(TX, TXR["lev_ti"], lambda p: f"={col(p)}{TXR['ebit']}", p0=0)  # placeholder; overwritten with interest deduction after DB is built
series(TX, TXR["lev_lossclose"], lambda p: f"=MAX(0,{col(p)}{TXR['lev_lossopen']}-MAX(0,{col(p)}{TXR['lev_ti']}))-MIN(0,{col(p)}{TXR['lev_ti']})", p0=0)
series(TX, TXR["lev_charge"], lambda p: f"=TaxRate*MAX(0,{col(p)}{TXR['lev_ti']}-{col(p)}{TXR['lev_lossopen']})", p0=0)
series(TX, TXR["lev_paid"], lambda p: f"={col(p-1)}{TXR['lev_charge']}", p0=0)
# Unlevered
series(TX, TXR["unl_lossopen"], lambda p: f"={col(p-1)}{TXR['unl_lossclose']}", p0=0)
series(TX, TXR["unl_ti"], lambda p: f"={col(p)}{TXR['ebit']}", p0=0)
series(TX, TXR["unl_lossclose"], lambda p: f"=MAX(0,{col(p)}{TXR['unl_lossopen']}-MAX(0,{col(p)}{TXR['unl_ti']}))-MIN(0,{col(p)}{TXR['unl_ti']})", p0=0)
series(TX, TXR["unl_charge"], lambda p: f"=TaxRate*MAX(0,{col(p)}{TXR['unl_ti']}-{col(p)}{TXR['unl_lossopen']})", p0=0)


# ==========================================================================
# DB — Debt (sculpted to flat target DSCR)   [needs CF; CF needs TX lev_paid]
# ==========================================================================
# CFADS is defined on CF; but DB interest is needed by TX lev_ti. To keep the
# graph acyclic: interest depends only on opening balance (period t-1 closing),
# tax is paid in arrears, so CFADS_t depends on tax_paid_t = charge_{t-1}.
# Order of sheets in the file doesn't matter to Excel; we just need refs right.
CF = new_sheet("CF", "CFADS  —  cash available for debt service")
CFR = dict(ebitda=3, taxpaid=4, wcmove=5, stack=6, cfads=7)
DB = new_sheet("DB", "DEBT  —  drawn at COD, sculpted to flat target DSCR")
DBR = dict(dopen=3, interest=4, cfads=5, principal=6, dclose=7, dscr=8, ds=9)

# CF sheet
label(CF, CFR["ebitda"], "EBITDA", "$")
label(CF, CFR["taxpaid"], "Less: tax paid (arrears)", "$")
label(CF, CFR["wcmove"], "Less: increase in NWC", "$")
label(CF, CFR["stack"], "Memo: stack replacement (below CFADS)", "$")
label(CF, CFR["cfads"], "CFADS (before major maintenance)", "$", bold=True)
series(CF, CFR["ebitda"], lambda p: f"=OP!{col(p)}{OPR['ebitda']}", p0=0)
series(CF, CFR["taxpaid"], lambda p: f"=TX!{col(p)}{TXR['lev_paid']}", p0=0)
series(CF, CFR["wcmove"], lambda p: f"=WC!{col(p)}{WCR['move']}", p0=0)
series(CF, CFR["stack"], lambda p: f"=OP!{col(p)}{OPR['stack']}", p0=0)
# Debt is sized on CFADS before lumpy major-maintenance (stack) capex, which is
# sponsor-funded via the cash sweep / equity injection below — this keeps CFADS
# and hence the sculpted DSCR smooth, as major maintenance is reserve/equity
# funded in a bankable structure.
series(CF, CFR["cfads"], lambda p: f"={col(p)}{CFR['ebitda']}-{col(p)}{CFR['taxpaid']}-{col(p)}{CFR['wcmove']}", p0=0)
add_name("CFADS_Row", CF.title, f"{col(1)}{CFR['cfads']}:{LAST_COL}{CFR['cfads']}")

# DB sheet
label(DB, DBR["dopen"], "Debt — opening balance", "$")
label(DB, DBR["interest"], "Interest (rate × opening)", "$")
label(DB, DBR["cfads"], "CFADS (link)", "$")
label(DB, DBR["principal"], "Principal (sculpted to target DSCR)", "$")
label(DB, DBR["dclose"], "Debt — closing balance", "$")
label(DB, DBR["dscr"], "DSCR (actual)", "x")
label(DB, DBR["ds"], "Debt service (int + principal)", "$")
series(DB, DBR["dopen"], lambda p: f"=IF({TMc('firstop',p)}=1,DebtAtCOD*GearingFactor,{col(p-1)}{DBR['dclose']})", p0=0)
series(DB, DBR["interest"], lambda p: f"={TMc('rdebt',p)}*{col(p)}{DBR['dopen']}", p0=0)
series(DB, DBR["cfads"], lambda p: f"=CF!{col(p)}{CFR['cfads']}", p0=0)
series(DB, DBR["principal"], lambda p: f"={TMc('debt',p)}*MIN({col(p)}{DBR['dopen']},MAX(0,{col(p)}{DBR['cfads']}/TargetDSCR_Input-{col(p)}{DBR['interest']}))", p0=0)
series(DB, DBR["dclose"], lambda p: f"={col(p)}{DBR['dopen']}-{col(p)}{DBR['principal']}", p0=0)
series(DB, DBR["ds"], lambda p: f"={col(p)}{DBR['interest']}+{col(p)}{DBR['principal']}", p0=0)
series(DB, DBR["dscr"], lambda p: f"=IF({col(p)}{DBR['ds']}=0,\"\",{col(p)}{DBR['cfads']}/{col(p)}{DBR['ds']})", p0="", num="0.00")
add_name("EndingBalanceResidual", DB.title, f"{col(TENOR_END_PERIOD)}{DBR['dclose']}")
add_name("DSCR_Row", DB.title, f"{col(COD_PERIOD)}{DBR['dscr']}:{col(TENOR_END_PERIOD)}{DBR['dscr']}")

# Fix TX levered taxable income now that DB interest row is known
# (deduct operating interest + construction cash IDC).
for p in range(1, N + 1):
    TX.cell(row=TXR["lev_ti"], column=P0_COL + p,
            value=f"={col(p)}{TXR['ebit']}-DB!{col(p)}{DBR['interest']}-CX!{col(p)}{CXR['cashint']}")


# ==========================================================================
# DS — DSRA
# ==========================================================================
DS = new_sheet("DS", "DSRA  —  debt service reserve account")
DSR = dict(req=3, open=4, move=5, interest=6, close=7)
label(DS, DSR["req"], "Required balance (lookahead cover)", "$")
label(DS, DSR["open"], "DSRA — opening", "$")
label(DS, DSR["move"], "Funding / (release)", "$")
label(DS, DSR["interest"], "Interest income (on opening)", "$")
label(DS, DSR["close"], "DSRA — closing", "$")
# Required = sum of next DSRA cover periods of scheduled debt service.
DSRA_LOOK = max(1, round(6 / 12 * PPY))   # 6 months cover -> 1 semi-annual period
def dsra_req(p):
    terms = []
    for k in range(1, DSRA_LOOK + 1):
        pk = p + k
        cell = f"DB!{col(pk)}{DBR['ds']}" if pk <= N else "0"
        terms.append(cell)
    # Gate by current-period debt flag so the DSRA is only held while debt is
    # live (funded at COD, released as debt amortises) — not pre-funded during
    # construction.
    return f"=({'+'.join(terms)})*{TMc('debt',p)}"
series(DS, DSR["req"], dsra_req, p0=0)
series(DS, DSR["open"], lambda p: f"={col(p-1)}{DSR['close']}", p0=0)
series(DS, DSR["interest"], lambda p: f"={TMc('rdebt',p)}*{col(p)}{DSR['open']}", p0=0)
series(DS, DSR["close"], lambda p: f"={col(p)}{DSR['req']}", p0=0)
series(DS, DSR["move"], lambda p: f"={col(p)}{DSR['close']}-{col(p)}{DSR['open']}", p0=0)


# ==========================================================================
# WF — Cash waterfall
# ==========================================================================
WF = new_sheet("WF", "CASH WATERFALL & DISTRIBUTIONS  (fully swept: cash at bank held at nil)")
WFR = dict(cfads=3, interest=4, principal=5, stack=6, dsra=7, dsraint=8, cafd=9,
           copen=10, inject=11, dist=12, cclose=13)
label(WF, WFR["cfads"], "CFADS (before major maintenance)", "$")
label(WF, WFR["interest"], "Less: interest", "$")
label(WF, WFR["principal"], "Less: principal", "$")
label(WF, WFR["stack"], "Less: stack replacement (major maint.)", "$")
label(WF, WFR["dsra"], "Less: DSRA funding / add release", "$")
label(WF, WFR["dsraint"], "Add: DSRA interest income", "$")
label(WF, WFR["cafd"], "Cash available (pre-distribution)", "$", bold=True)
label(WF, WFR["copen"], "Cash at bank — opening", "$")
label(WF, WFR["inject"], "Equity injection (fund shortfall)", "$")
label(WF, WFR["dist"], "Distributions to equity", "$", bold=True)
label(WF, WFR["cclose"], "Cash at bank — closing", "$")
series(WF, WFR["cfads"], lambda p: f"=CF!{col(p)}{CFR['cfads']}", p0=0)
series(WF, WFR["interest"], lambda p: f"=DB!{col(p)}{DBR['interest']}", p0=0)
series(WF, WFR["principal"], lambda p: f"=DB!{col(p)}{DBR['principal']}", p0=0)
series(WF, WFR["stack"], lambda p: f"=OP!{col(p)}{OPR['stack']}", p0=0)
series(WF, WFR["dsra"], lambda p: f"=DS!{col(p)}{DSR['move']}", p0=0)
series(WF, WFR["dsraint"], lambda p: f"=DS!{col(p)}{DSR['interest']}", p0=0)
series(WF, WFR["cafd"], lambda p: f"={col(p)}{WFR['cfads']}-{col(p)}{WFR['interest']}-{col(p)}{WFR['principal']}-{col(p)}{WFR['stack']}-{col(p)}{WFR['dsra']}+{col(p)}{WFR['dsraint']}", p0=0)
series(WF, WFR["copen"], lambda p: f"={col(p-1)}{WFR['cclose']}", p0=0)
# Fully-swept structure: inject equity to cover any shortfall, distribute any
# surplus in operating periods -> cash at bank held at nil, so no negative cash.
series(WF, WFR["inject"], lambda p: f"=MAX(0,-({col(p)}{WFR['copen']}+{col(p)}{WFR['cafd']}))", p0=0)
series(WF, WFR["dist"], lambda p: f"=MAX(0,{col(p)}{WFR['copen']}+{col(p)}{WFR['cafd']})*{TMc('ops',p)}", p0=0)
series(WF, WFR["cclose"], lambda p: f"={col(p)}{WFR['copen']}+{col(p)}{WFR['cafd']}+{col(p)}{WFR['inject']}-{col(p)}{WFR['dist']}", p0=0)


# ==========================================================================
# CR — Cover ratios
# ==========================================================================
CR = new_sheet("CR", "COVER RATIOS  —  DSCR, LLCR, PLCR")
CRR = dict(dscr=3, npv_ll=4, llcr=5, npv_pl=6, plcr=7)
label(CR, CRR["dscr"], "DSCR (period)", "x")
label(CR, CRR["npv_ll"], "NPV of remaining CFADS to tenor end", "$")
label(CR, CRR["llcr"], "LLCR", "x")
label(CR, CRR["npv_pl"], "NPV of remaining CFADS to project end", "$")
label(CR, CRR["plcr"], "PLCR", "x")
series(CR, CRR["dscr"], lambda p: f"=DB!{col(p)}{DBR['dscr']}", p0="", num="0.00")
# Absolute full-row ranges for the SUMPRODUCT cover-ratio NPVs.
PER_ROW = f"TM!${col(1)}${TMR['period']}:${LAST_COL}${TMR['period']}"
DEBTFLAG_ROW = f"TM!${col(1)}${TMR['debt']}:${LAST_COL}${TMR['debt']}"
OPSFLAG_ROW = f"TM!${col(1)}${TMR['ops']}:${LAST_COL}${TMR['ops']}"
CFADS_FULL = f"CF!${col(1)}${CFR['cfads']}:${LAST_COL}${CFR['cfads']}"
# LLCR: PV (at debt rate) of CFADS from period p to tenor end / debt outstanding.
series(CR, CRR["npv_ll"],
       lambda p: (f"=SUMPRODUCT(({PER_ROW}>=TM!{col(p)}{TMR['period']})*{DEBTFLAG_ROW}*{CFADS_FULL}"
                  f"/(1+{TMc('rdebt',p)})^({PER_ROW}-TM!{col(p)}{TMR['period']}+1))"), p0=0)
series(CR, CRR["llcr"], lambda p: f"=IF(DB!{col(p)}{DBR['dopen']}=0,\"\",{col(p)}{CRR['npv_ll']}/DB!{col(p)}{DBR['dopen']})", p0="", num="0.00")
# PLCR: PV of CFADS from period p to project end / debt outstanding.
series(CR, CRR["npv_pl"],
       lambda p: (f"=SUMPRODUCT(({PER_ROW}>=TM!{col(p)}{TMR['period']})*{OPSFLAG_ROW}*{CFADS_FULL}"
                  f"/(1+{TMc('rdebt',p)})^({PER_ROW}-TM!{col(p)}{TMR['period']}+1))"), p0=0)
series(CR, CRR["plcr"], lambda p: f"=IF(DB!{col(p)}{DBR['dopen']}=0,\"\",{col(p)}{CRR['npv_pl']}/DB!{col(p)}{DBR['dopen']})", p0="", num="0.00")


# ==========================================================================
# FS — Financial statements
# ==========================================================================
FS = new_sheet("FS", "FINANCIAL STATEMENTS  —  P&L and balance sheet")
FSR = dict(
    pl=2, rev=3, opex=4, ebitda=5, stackexp=6, dsrainc=7, dep=8, interest=9,
    pbt=10, tax=11, npat=12,
    bs=14, ppe=15, dsra=16, nwc=17, cash=18, assets=19,
    debt=20, taxpay=21, sharecap=22, retained=23, eql=24, bscheck=25)
label(FS, FSR["pl"], "PROFIT & LOSS", section=True)
label(FS, FSR["rev"], "Revenue", "$")
label(FS, FSR["opex"], "Operating costs", "$")
label(FS, FSR["ebitda"], "EBITDA", "$")
label(FS, FSR["stackexp"], "Stack replacement (expensed)", "$")
label(FS, FSR["dsrainc"], "DSRA interest income", "$")
label(FS, FSR["dep"], "Depreciation", "$")
label(FS, FSR["interest"], "Interest", "$")
label(FS, FSR["pbt"], "Profit before tax", "$")
label(FS, FSR["tax"], "Tax charge", "$")
label(FS, FSR["npat"], "Net profit after tax", "$", bold=True)
label(FS, FSR["bs"], "BALANCE SHEET", section=True)
label(FS, FSR["ppe"], "PP&E (net)", "$")
label(FS, FSR["dsra"], "DSRA", "$")
label(FS, FSR["nwc"], "Net working capital", "$")
label(FS, FSR["cash"], "Cash at bank", "$")
label(FS, FSR["assets"], "Total assets", "$", bold=True)
label(FS, FSR["debt"], "Debt", "$")
label(FS, FSR["taxpay"], "Tax payable", "$")
label(FS, FSR["sharecap"], "Share capital", "$")
label(FS, FSR["retained"], "Retained earnings", "$")
label(FS, FSR["eql"], "Total equity + liabilities", "$", bold=True)
label(FS, FSR["bscheck"], "Balance check (≈0)", "$")
series(FS, FSR["rev"], lambda p: f"=PR!{col(p)}{PRR['rev']}", p0=0)
series(FS, FSR["opex"], lambda p: f"=-OP!{col(p)}{OPR['totopex']}", p0=0)
series(FS, FSR["ebitda"], lambda p: f"=OP!{col(p)}{OPR['ebitda']}", p0=0)
series(FS, FSR["stackexp"], lambda p: f"=-OP!{col(p)}{OPR['stack']}", p0=0)
series(FS, FSR["dsrainc"], lambda p: f"=DS!{col(p)}{DSR['interest']}", p0=0)
series(FS, FSR["dep"], lambda p: f"=-TX!{col(p)}{TXR['dep']}", p0=0)
# Interest expense = operating debt interest + construction cash-funded IDC
# (the capitalised part is on the balance sheet, not the P&L).
series(FS, FSR["interest"], lambda p: f"=-DB!{col(p)}{DBR['interest']}-CX!{col(p)}{CXR['cashint']}", p0=0)
series(FS, FSR["pbt"], lambda p: f"={col(p)}{FSR['ebitda']}+{col(p)}{FSR['stackexp']}+{col(p)}{FSR['dsrainc']}+{col(p)}{FSR['dep']}+{col(p)}{FSR['interest']}", p0=0)
series(FS, FSR["tax"], lambda p: f"=-TX!{col(p)}{TXR['lev_charge']}", p0=0)
series(FS, FSR["npat"], lambda p: f"={col(p)}{FSR['pbt']}+{col(p)}{FSR['tax']}", p0=0)
# Balance sheet
series(FS, FSR["ppe"], lambda p: f"={col(p-1)}{FSR['ppe']}+CX!{col(p)}{CXR['ppe_add']}-TX!{col(p)}{TXR['dep']}", p0=0)
series(FS, FSR["dsra"], lambda p: f"=DS!{col(p)}{DSR['close']}", p0=0)
series(FS, FSR["nwc"], lambda p: f"=WC!{col(p)}{WCR['nwc']}", p0=0)
series(FS, FSR["cash"], lambda p: f"=WF!{col(p)}{WFR['cclose']}", p0=0)
series(FS, FSR["assets"], lambda p: f"={col(p)}{FSR['ppe']}+{col(p)}{FSR['dsra']}+{col(p)}{FSR['nwc']}+{col(p)}{FSR['cash']}", p0=0)
series(FS, FSR["debt"], lambda p: f"=CX!{col(p)}{CXR['dclose']}*{TMc('constr',p)}+DB!{col(p)}{DBR['dclose']}*(1-{TMc('constr',p)})", p0=0)
# Tax payable = accrued tax charge not yet paid (charge in period, cash one period later)
series(FS, FSR["taxpay"], lambda p: f"={col(p-1)}{FSR['taxpay']}+TX!{col(p)}{TXR['lev_charge']}-TX!{col(p)}{TXR['lev_paid']}", p0=0)
# Share capital = cumulative equity: construction equity draws + cash IDC funded
# by equity + operating-period injections + the COD refinancing top-up
# (construction facility is refinanced by the smaller sculpted term debt at COD;
# equity plugs the gap DebtAtCOD*(1-GearingFactor)).
series(FS, FSR["sharecap"], lambda p: f"={col(p-1)}{FSR['sharecap']}+CX!{col(p)}{CXR['eqdraw']}+CX!{col(p)}{CXR['cashint']}+WF!{col(p)}{WFR['inject']}+{TMc('firstop',p)}*DebtAtCOD*(1-GearingFactor)", p0=0)
series(FS, FSR["retained"], lambda p: f"={col(p-1)}{FSR['retained']}+{col(p)}{FSR['npat']}-WF!{col(p)}{WFR['dist']}", p0=0)
series(FS, FSR["eql"], lambda p: f"={col(p)}{FSR['debt']}+{col(p)}{FSR['taxpay']}+{col(p)}{FSR['sharecap']}+{col(p)}{FSR['retained']}", p0=0)
series(FS, FSR["bscheck"], lambda p: f"={col(p)}{FSR['assets']}-{col(p)}{FSR['eql']}", p0=0)


# ==========================================================================
# RT — Returns
# ==========================================================================
RT = new_sheet("RT", "RETURNS")
RTR = dict(eqcf=3, projcf=4)
label(RT, RTR["eqcf"], "Equity cash flow", "$")
label(RT, RTR["projcf"], "Project cash flow (unlevered)", "$")
series(RT, RTR["eqcf"], lambda p: f"=-CX!{col(p)}{CXR['eqdraw']}-CX!{col(p)}{CXR['cashint']}-WF!{col(p)}{WFR['inject']}-{TMc('firstop',p)}*DebtAtCOD*(1-GearingFactor)+WF!{col(p)}{WFR['dist']}", p0=0)
series(RT, RTR["projcf"], lambda p: f"=-CX!{col(p)}{CXR['capex']}+OP!{col(p)}{OPR['ebitda']}-OP!{col(p)}{OPR['stack']}-TX!{col(p)}{TXR['unl_charge']}-WC!{col(p)}{WCR['move']}", p0=0)
RT.cell(row=6, column=1, value="Equity IRR (nominal, annualised)").font = BOLD
RT.cell(row=6, column=P0_COL, value=f"=(1+IRR({col(0)}{RTR['eqcf']}:{LAST_COL}{RTR['eqcf']},0.05))^PeriodsPerYear-1").number_format = "0.0%"
add_name("EquityIRR", RT.title, f"{col(0)}6")
RT.cell(row=7, column=1, value="Project IRR (nominal, annualised)").font = BOLD
RT.cell(row=7, column=P0_COL, value=f"=(1+IRR({col(0)}{RTR['projcf']}:{LAST_COL}{RTR['projcf']},0.05))^PeriodsPerYear-1").number_format = "0.0%"
add_name("ProjectIRR", RT.title, f"{col(0)}7")
RT.cell(row=8, column=1, value="Equity NPV @ cost of equity").font = BOLD
RT.cell(row=8, column=P0_COL, value=f"=NPV(CostOfEquity/PeriodsPerYear,{col(1)}{RTR['eqcf']}:{LAST_COL}{RTR['eqcf']})+{col(0)}{RTR['eqcf']}").number_format = "#,##0"


# ==========================================================================
# LC — LCOH (nominal & real, cost stack)
# ==========================================================================
LC = new_sheet("LC", "LCOH  —  levelised cost of hydrogen (unlevered)")
LCR = dict(capex=3, power=4, water=5, fixed=6, var=7, ins=8, stack=9, tax=10, totcost=11,
           h2=12, dfn=13, dfr=14)
label(LC, LCR["capex"], "Capex", "$")
label(LC, LCR["power"], "Power", "$")
label(LC, LCR["water"], "Water", "$")
label(LC, LCR["fixed"], "Fixed O&M", "$")
label(LC, LCR["var"], "Variable O&M", "$")
label(LC, LCR["ins"], "Insurance", "$")
label(LC, LCR["stack"], "Stack replacement", "$")
label(LC, LCR["tax"], "Tax (unlevered)", "$")
label(LC, LCR["totcost"], "Total cost", "$", bold=True)
label(LC, LCR["h2"], "H2 production", "kg")
label(LC, LCR["dfn"], "Discount factor (nominal)", "x")
label(LC, LCR["dfr"], "Discount factor (real)", "x")
series(LC, LCR["capex"], lambda p: f"=CX!{col(p)}{CXR['capex']}", p0=0)
series(LC, LCR["power"], lambda p: f"=OP!{col(p)}{OPR['power']}", p0=0)
series(LC, LCR["water"], lambda p: f"=OP!{col(p)}{OPR['water']}", p0=0)
series(LC, LCR["fixed"], lambda p: f"=OP!{col(p)}{OPR['fixed']}", p0=0)
series(LC, LCR["var"], lambda p: f"=OP!{col(p)}{OPR['var']}", p0=0)
series(LC, LCR["ins"], lambda p: f"=OP!{col(p)}{OPR['ins']}", p0=0)
series(LC, LCR["stack"], lambda p: f"=OP!{col(p)}{OPR['stack']}", p0=0)
series(LC, LCR["tax"], lambda p: f"=TX!{col(p)}{TXR['unl_charge']}", p0=0)
series(LC, LCR["totcost"], lambda p: "=" + "+".join(f"{col(p)}{LCR[k]}" for k in ["capex","power","water","fixed","var","ins","stack","tax"]), p0=0)
series(LC, LCR["h2"], lambda p: f"=PR!{col(p)}{PRR['h2']}", p0=0)
series(LC, LCR["dfn"], lambda p: f"={TMc('df_nom',p)}", p0=1, num="0.0000")
series(LC, LCR["dfr"], lambda p: f"={TMc('df_real',p)}", p0=1, num="0.0000")

# Component cost stack ($/kg) and headline LCOH
STK = 16
LC.cell(row=STK, column=1, value="LCOH COST STACK ($/kg, nominal)").font = BOLD
comp_rows = {}
comps = [("Capex","capex"),("Power","power"),("Water","water"),("Fixed O&M","fixed"),
         ("Variable O&M","var"),("Insurance","ins"),("Stack","stack"),("Tax","tax")]
rr = STK + 1
denom = f"SUMPRODUCT(LC!{col(1)}{LCR['h2']}:{LAST_COL}{LCR['h2']},LC!{col(1)}{LCR['dfn']}:{LAST_COL}{LCR['dfn']})"
for name, key in comps:
    LC.cell(row=rr, column=1, value=name)
    LC.cell(row=rr, column=P0_COL,
            value=f"=SUMPRODUCT(LC!{col(1)}{LCR[key]}:{LAST_COL}{LCR[key]},LC!{col(1)}{LCR['dfn']}:{LAST_COL}{LCR['dfn']})/{denom}").number_format = "#,##0.000"
    comp_rows[key] = rr
    rr += 1
LC.cell(row=rr, column=1, value="LCOH (nominal)").font = Font(bold=True, size=12)
lcoh_nom = LC.cell(row=rr, column=P0_COL, value="=" + "+".join(f"{col(0)}{comp_rows[k]}" for _, k in comps))
lcoh_nom.number_format = "#,##0.000"; lcoh_nom.fill = CHECK_OK_FILL
add_name("LCOH_Nominal", LC.title, f"{col(0)}{rr}")
rr += 1
denom_r = f"SUMPRODUCT(LC!{col(1)}{LCR['h2']}:{LAST_COL}{LCR['h2']},LC!{col(1)}{LCR['dfr']}:{LAST_COL}{LCR['dfr']})"
LC.cell(row=rr, column=1, value="LCOH (real)").font = Font(bold=True, size=12)
lcoh_real = LC.cell(row=rr, column=P0_COL,
    value=f"=SUMPRODUCT(LC!{col(1)}{LCR['totcost']}:{LAST_COL}{LCR['totcost']},LC!{col(1)}{LCR['dfr']}:{LAST_COL}{LCR['dfr']})/{denom_r}")
lcoh_real.number_format = "#,##0.000"; lcoh_real.fill = CHECK_OK_FILL
add_name("LCOH_Real", LC.title, f"{col(0)}{rr}")


# ==========================================================================
# CK — Checks
# ==========================================================================
CK = new_sheet("CK", "CHECKS  —  model integrity")
CKR = {}
checks = [
    ("Balance sheet balances (all periods)",
     f"=IF(SUMPRODUCT(ABS(FS!{col(1)}{FSR['bscheck']}:{LAST_COL}{FSR['bscheck']}))<1,1,0)"),
    ("Debt fully amortised by tenor end",
     f"=IF(ABS(DB!{col(TENOR_END_PERIOD)}{DBR['dclose']})<Tolerance,1,0)"),
    ("Debt never negative",
     f"=IF(MIN(DB!{col(1)}{DBR['dclose']}:{LAST_COL}{DBR['dclose']})>=-1,1,0)"),
    ("Cash never negative",
     f"=IF(MIN(WF!{col(1)}{WFR['cclose']}:{LAST_COL}{WFR['cclose']})>=-1,1,0)"),
    ("DSRA never negative",
     f"=IF(MIN(DS!{col(1)}{DSR['close']}:{LAST_COL}{DSR['close']})>=-1,1,0)"),
    ("Construction / operations flags partition timeline",
     f"=IF(SUMPRODUCT(TM!{col(1)}{TMR['constr']}:{LAST_COL}{TMR['constr']})+SUMPRODUCT(TM!{col(1)}{TMR['ops']}:{LAST_COL}{TMR['ops']})={N},1,0)"),
    ("Min DSCR ≥ lender covenant (post-COD)",
     f"=IF(MIN(DB!{col(COD_PERIOD)}{DBR['dscr']}:{col(TENOR_END_PERIOD)}{DBR['dscr']})>=CovenantDSCR-0.01,1,0)"),
    ("LCOH computed (headline output present)", f"=IF(LCOH_Nominal>0,1,0)"),
]
rr = 3
for text, fml in checks:
    label(CK, rr, text)
    c = CK.cell(row=rr, column=P0_COL, value=fml); c.number_format = "0"
    CKR[text] = rr
    rr += 1
label(CK, rr + 1, "MASTER CHECK (all pass)", bold=True)
mc = CK.cell(row=rr + 1, column=P0_COL,
             value=f"=IF(SUM({col(0)}3:{col(0)}{rr-1})={len(checks)},1,0)")
mc.number_format = "0"; mc.fill = CHECK_OK_FILL
add_name("MasterCheck", CK.title, f"{col(0)}{rr+1}")


# ==========================================================================
# OUT — Dashboard
# ==========================================================================
OUT = new_sheet("OUT", "DASHBOARD  —  key outputs")
kpis = [
    ("LCOH — nominal ($/kg)", "=LCOH_Nominal", "#,##0.000"),
    ("LCOH — real ($/kg)", "=LCOH_Real", "#,##0.000"),
    ("Project IRR", "=ProjectIRR", "0.0%"),
    ("Equity IRR", "=EquityIRR", "0.0%"),
    ("Min DSCR", f"=MIN(DB!{col(COD_PERIOD)}{DBR['dscr']}:{col(TENOR_END_PERIOD)}{DBR['dscr']})", "0.00"),
    ("Avg DSCR", f"=AVERAGE(DB!{col(COD_PERIOD)}{DBR['dscr']}:{col(TENOR_END_PERIOD)}{DBR['dscr']})", "0.00"),
    ("LLCR at COD", f"=CR!{col(COD_PERIOD)}{CRR['llcr']}", "0.00"),
    ("PLCR at COD", f"=CR!{col(COD_PERIOD)}{CRR['plcr']}", "0.00"),
    ("Debt at COD ($)", "=DebtAtCOD*GearingFactor", "#,##0"),
    ("Gearing factor (solved)", "=GearingFactor", "0.000"),
    ("Convergence status", "=ConvergenceStatus", "General"),
    ("MASTER CHECK (1 = all pass)", "=MasterCheck", "0"),
]
rr = 3
for text, fml, fmt in kpis:
    label(OUT, rr, text, bold=True)
    c = OUT.cell(row=rr, column=P0_COL, value=fml); c.number_format = fmt
    if "MASTER" in text:
        c.fill = CHECK_OK_FILL
    rr += 1
OUT.cell(row=rr + 1, column=1, value="Re-solve after any input change: run RunDebtSizing (macro) or Goal Seek — set EndingBalanceResidual to 0 by changing 'Sculpting DSCR (solved)'.").font = GREY

# Sensitivity grid (populated by the RunSensitivityGrid macro).
sg = rr + 3
OUT.cell(row=sg, column=1, value="SENSITIVITY GRID (run RunSensitivityGrid macro to populate)").font = BOLD
hdrs = ["Scenario", "Capex Δ%", "Power Δ%", "Price Δ%", "WACC Δbps", "→ LCOH nom", "→ Equity IRR", "→ Min DSCR"]
for i, h in enumerate(hdrs):
    OUT.cell(row=sg + 1, column=1 + i, value=h).font = BOLD
scen = [
    ("Base", 0.0, 0.0, 0.0, 0),
    ("Capex +10%", 0.10, 0.0, 0.0, 0),
    ("Capex -10%", -0.10, 0.0, 0.0, 0),
    ("Power +20%", 0.0, 0.20, 0.0, 0),
    ("Power -20%", 0.0, -0.20, 0.0, 0),
    ("Price +10%", 0.0, 0.0, 0.10, 0),
    ("Price -10%", 0.0, 0.0, -0.10, 0),
    ("WACC +100bps", 0.0, 0.0, 0.0, 100),
    ("WACC -100bps", 0.0, 0.0, 0.0, -100),
]
for i, (name, dc, dp, dpr, dw) in enumerate(scen):
    rrow = sg + 2 + i
    OUT.cell(row=rrow, column=1, value=name)
    OUT.cell(row=rrow, column=2, value=dc).number_format = "0%"
    OUT.cell(row=rrow, column=3, value=dp).number_format = "0%"
    OUT.cell(row=rrow, column=4, value=dpr).number_format = "0%"
    OUT.cell(row=rrow, column=5, value=dw)
add_name("SensFirstRow", OUT.title, f"A{sg + 2}")
add_name("SensLastRow", OUT.title, f"A{sg + 1 + len(scen)}")


# ==========================================================================
# CV — Cover (built last, placed first)
# ==========================================================================
CV = new_sheet("CV", "GREEN HYDROGEN — LCOH PROJECT-FINANCE MODEL")
CV.cell(row=1, column=1).value = "GREEN HYDROGEN — LCOH PROJECT-FINANCE MODEL"
cover_lines = [
    ("", ""),
    ("Built to FAST / F1F9 project-finance conventions.", "bold"),
    ("", ""),
    ("Conventions", "sec"),
    ("  Blue cells = hardcoded inputs (Inputs sheet only). Black = formula. Green links = cross-sheet.", ""),
    ("  Time runs across columns; one line item per row; a single formula copies unbroken across each row.", ""),
    ("  Timing is driven by flag & counter rows on the Time (TM) sheet — never by branching a formula per column.", ""),
    ("  A period-0 column holds opening balances, so every corkscrew (opening = prior closing) is one formula.", ""),
    ("  Iterative calculation is OFF and must stay off — nothing here needs it. A circular warning means a real bug.", ""),
    ("", ""),
    ("Circularity handling", "sec"),
    ("  IDC during construction is solved in closed form (CX sheet) — no iteration.", ""),
    ("  All interest accrues on opening balances; tax is paid one period in arrears; DSRA interest on opening balance.", ""),
    ("  Debt is sculpted to a flat target DSCR; the single scalar GearingFactor is solved by the RunDebtSizing macro", ""),
    ("  (or native Goal Seek: set EndingBalanceResidual to 0 by changing GearingFactor).", ""),
    ("", ""),
    ("Sheet index", "sec"),
    ("  CV Cover · IN Inputs · TM Time · CX Construction/Capex · PR Production/Revenue · OP Opex · WC Working capital", ""),
    ("  TX Tax · CF CFADS · DB Debt · DS DSRA · WF Waterfall · CR Cover ratios · FS Financial statements", ""),
    ("  RT Returns · LC LCOH · CK Checks · OUT Dashboard", ""),
    ("", ""),
    ("Key outputs", "sec"),
]
rr = 2
for text, style in cover_lines:
    c = CV.cell(row=rr, column=1, value=text)
    if style == "bold":
        c.font = BOLD
    elif style == "sec":
        c.font = SECTION_FONT
    rr += 1
for text, fml, fmt in [("LCOH — nominal ($/kg)", "=LCOH_Nominal", "#,##0.000"),
                       ("LCOH — real ($/kg)", "=LCOH_Real", "#,##0.000"),
                       ("Equity IRR", "=EquityIRR", "0.0%"),
                       ("Master check (1 = all pass)", "=MasterCheck", "0")]:
    CV.cell(row=rr, column=1, value="  " + text).font = BOLD
    c = CV.cell(row=rr, column=UNIT_COL + 1, value=fml); c.number_format = fmt
    rr += 1

# Order sheets: CV first
wb.move_sheet("CV", -(len(wb.sheetnames) - 1))

# --------------------------------------------------------------------------
finalize_names()
out_path = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "HydrogenLCOH_Model.xlsx")
wb.calculation.calcMode = "auto"
wb.calculation.iterate = False
wb.save(out_path)
print(f"Wrote {out_path}")
print(f"N={N} periods, COD={COD_PERIOD}, tenor end={TENOR_END_PERIOD}, {len(wb.sheetnames)} sheets, {len(_defined)} names")
