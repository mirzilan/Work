from openpyxl.styles import Font
from openpyxl.worksheet.worksheet import Worksheet
from openpyxl.workbook import Workbook

import calc_capex as capex
import calc_cfads as cfads
import calc_financing_cons as fin_cons
import fs_quarterly as fsq
from inputs import ProjectInputs
from timeline import Timeline
from workbook_builder import (
    FIRST_DATA_COL,
    COL_UNITS,
    COL_TOTAL,
    COL_CHECK,
    COL_REMARKS,
    COLOR_FORMULA,
    COLOR_LINK,
    TAB_COLOR_OUTPUT,
    style_total_row,
    style_section_header_row,
)
import openpyxl.utils

# FS_Annual mirrors FS_Quarterly's own layout section-for-section (P&L, then Balance
# Sheet, then Cash Flow) — every row here is either an annual SUM of a FS_Quarterly flow
# row or a year-end (last quarter) pull of a FS_Quarterly stock row. PIRR/EIRR sit below
# the three statements, since they're a whole-of-life output, not part of any one of them.

# Flags block — same three flags, same rows, as FS_Quarterly, so the two read identically
# whichever one the user has open.
ROW_FLAG_ACTIVE_SCENARIO = 2
ROW_FLAG_MODEL_STATUS = 3
ROW_FLAG_SOLVE_FRESHNESS = 4

ROW_COLUMN_HEADER = 6  # Units/Total/Check/Remarks column headers, same row as FS_Quarterly
ROW_YEAR_LABEL = 7

# --- P&L (annual sums) ---
ROW_PNL_HEADER = 9
ROW_REVENUE = 10
ROW_OPEX = 11
ROW_EBITDA = 12
ROW_DEPRECIATION = 13
ROW_EBIT = 14
ROW_INTEREST_EXPENSE = 15
ROW_EBT = 16
ROW_TAX = 17
ROW_LC_FEE = 18
ROW_NET_INCOME = 19

# --- Balance Sheet (year-end, i.e. Q4 of each project year) ---
ROW_BS_HEADER = 21
ROW_BS_ASSETS_HEADER = 22
ROW_BS_CURRENT_ASSETS_HEADER = 23
ROW_BS_CASH = 24
ROW_BS_AR = 25  # trade receivables — annual sum's year-end pull of FS_Quarterly's own row
ROW_BS_TOTAL_CURRENT_ASSETS = 26
ROW_BS_NONCURRENT_ASSETS_HEADER = 27
ROW_BS_DSRA = 28
ROW_BS_MRA = 29
ROW_BS_PPE_NET = 30
ROW_BS_TOTAL_NONCURRENT_ASSETS = 31
ROW_BS_TOTAL_ASSETS = 32

ROW_BS_LIABILITIES_HEADER = 34
ROW_BS_CURRENT_LIAB_HEADER = 35
ROW_BS_DEBT_CURRENT = 36
ROW_BS_AP = 37  # trade payables — year-end pull of FS_Quarterly's own row
ROW_BS_TOTAL_CURRENT_LIAB = 38
ROW_BS_NONCURRENT_LIAB_HEADER = 39
ROW_BS_DEBT_NONCURRENT = 40
ROW_BS_TOTAL_NONCURRENT_LIAB = 41
ROW_BS_TOTAL_LIABILITIES = 42

ROW_BS_EQUITY_HEADER = 44
ROW_BS_PAID_IN_CAPITAL = 45
ROW_BS_RETAINED_EARNINGS = 46
ROW_BS_TOTAL_EQUITY = 47

ROW_BS_TOTAL_LIAB_EQUITY = 49
ROW_BS_CHECK_A_MINUS_L = 50

# --- Cash Flow (annual sums for flows, year-end for the cash walk) ---
ROW_CF_HEADER = 52
ROW_CFO_NI = 53
ROW_CFO_ADDBACK_DEPR = 54
ROW_CFO_WC_CHANGE = 55
ROW_CFO = 56
ROW_CFI = 57
ROW_CFF_PRINCIPAL = 58
ROW_CFF_DIVIDENDS = 59
ROW_CFF_EQUITY_INJECTION = 60
ROW_CFF = 61
ROW_NET_CHANGE_TOTAL_CASH = 62
ROW_OPENING_TOTAL_CASH = 63
ROW_CLOSING_TOTAL_CASH = 64

ROW_CF_RESTRICTED_HEADER = 66
ROW_CF_LESS_DSRA = 67
ROW_CF_LESS_MRA = 68
ROW_CLOSING_CASH = 69

ROW_CFD_HEADER = 71
ROW_CFD_RECEIPTS = 72
ROW_CFD_OPEX_PAID = 73
ROW_CFD_INTEREST_PAID = 74
ROW_CFD_TAX_PAID = 75
ROW_CFD_LC_FEE_PAID = 76
ROW_CFO_DIRECT = 77
ROW_CHECK_DIRECT_TIES_INDIRECT = 78

ROW_CHECK_HEADER = 80
ROW_CHECK_BS_BALANCES_COUNT = 81
ROW_CHECK_CASH_TIES_BUFFER = 82
ROW_CHECK_DIRECT_TIES_INDIRECT_COUNT = 83

# FCFF / FCFE — two views, cross-checked; annual sums of FS_Quarterly's own dual-view rows.
# See FS_Quarterly for the full explanation of why FCFF ties exactly every period while
# FCFE only ties in total (buffer/lock-up timing).
ROW_FCF_HEADER = 85
ROW_FCFF_CF_METHOD = 86
ROW_FCFF_CFADS_METHOD = 87
ROW_CHECK_FCFF_METHODS_TIE_COUNT = 88
ROW_FCFE_CF_METHOD = 89
ROW_FCFE_DIVIDEND_METHOD = 90
ROW_CHECK_FCFE_LIFETIME_TIE = 91

# Backward-compat aliases — the pre-classified BS exposed one row each for Debt and Cash;
# downstream sheets (Valuation_SellDown) still want a single "closing debt"/"closing
# equity" figure, which is unambiguous here since Debt is the only liability category.
ROW_CASH_CLOSING = ROW_CLOSING_CASH
ROW_DEBT_CLOSING = ROW_BS_TOTAL_LIABILITIES
ROW_TOTAL_EQUITY_CLOSING = ROW_BS_TOTAL_EQUITY

# XIRR helper block: one continuous row of dates + one row of project (unlevered) cash flow,
# one row of equity cash flow, spanning ALL periods (24 construction months + 80 ops quarters).
# This block uses per-PERIOD columns (one column per month/quarter of the whole project
# life), which is a different column scheme from the per-YEAR columns the three statements
# above use — the two must never be read from each other's columns.
ROW_XIRR_DATE = 94
ROW_XIRR_PROJECT_CF = 95
ROW_XIRR_EQUITY_CF = 96

ROW_PIRR_LABEL = 99
ROW_PIRR_VALUE = 100
ROW_EIRR_LABEL = 101
ROW_EIRR_VALUE = 102

# Construction-period annual summary (monthly source data rolled to project years)
ROW_CONS_HEADER = 105
ROW_CONS_YEAR_LABEL = 106
ROW_CONS_CAPEX = 107
ROW_CONS_IDC = 108
ROW_CONS_DEBT_DRAWN = 109
ROW_CONS_EQUITY_DRAWN = 110
ROW_CONS_CUM_TPC = 111
ROW_CONS_CLOSING_DEBT = 112

# Construction-period Balance Sheet — year-end, sourced from the same monthly cells as the
# summary above. P&L is legitimately empty pre-COD (no revenue), but the BS still has to
# provably balance every year so the Day-1 operating BS is derived, not asserted.
ROW_CONS_BS_HEADER = 114
ROW_CONS_BS_CASH = 115
ROW_CONS_BS_DSRA = 116
ROW_CONS_BS_PPE = 117
ROW_CONS_BS_TOTAL_ASSETS = 118
ROW_CONS_BS_DEBT = 119
ROW_CONS_BS_PAID_IN_CAPITAL = 120
ROW_CONS_BS_RETAINED_EARNINGS = 121
ROW_CONS_BS_TOTAL_EQUITY = 122
ROW_CONS_BS_TOTAL_LIAB_EQUITY = 123

ROW_CONS_CHECK_HEADER = 125
ROW_CONS_CHECK_BS_BALANCES_COUNT = 126  # count of construction years where the BS does not balance

# Construction Cash Flow Statement — monthly (matches Calc_Capex's own resolution), so the
# XIRR helper's construction-period Equity CF reads from a built statement rather than
# reaching into Calc_Capex directly. Investing includes capitalised interest (IDC is a
# real investment, not a financing cost) so Net Cash Flow ties to zero every month except
# the last, when the DSRA/Buffer are funded — the same cash pattern the construction BS
# above already assumes.
ROW_CONS_CF_HEADER = 128
ROW_CONS_CF_INVESTING = 129     # -(Capex Draw + IDC)
ROW_CONS_CF_DEBT_DRAWN = 130
ROW_CONS_CF_EQUITY_DRAWN = 131
ROW_CONS_CF_FINANCING = 132     # Debt Drawn + Equity Drawn
ROW_CONS_CF_NET = 133

ROW_CONS_CF_CHECK_HEADER = 135
ROW_CONS_CF_CHECK_NET_ZERO_COUNT = 136  # # of months where Net CF != 0 (all but the last)


# Mirrors FS_Quarterly's own _METADATA table — same rows, same meaning, just annual-sum
# totals instead of quarterly. Only covers the three main statements (P&L/BS/CF); the
# XIRR helper and Construction sections use a different column scheme (per-period, not
# per-year) and are deliberately left out of this pass.
_METADATA = [
    (ROW_REVENUE, "$", True, None, None),
    (ROW_OPEX, "$", True, None, None),
    (ROW_EBITDA, "$", True, None, None),
    (ROW_DEPRECIATION, "$", True, None, None),
    (ROW_EBIT, "$", True, None, None),
    (ROW_INTEREST_EXPENSE, "$", True, None, None),
    (ROW_EBT, "$", True, None, None),
    (ROW_TAX, "$", True, None, None),
    (ROW_LC_FEE, "$", True, None, None),
    (ROW_NET_INCOME, "$", True, None, None),

    (ROW_BS_CASH, "$", False, None, None),
    (ROW_BS_AR, "$", False, None, None),
    (ROW_BS_TOTAL_CURRENT_ASSETS, "$", False, None, None),
    (ROW_BS_DSRA, "$", False, None, None),
    (ROW_BS_MRA, "$", False, None, None),
    (ROW_BS_PPE_NET, "$", False, None, None),
    (ROW_BS_TOTAL_NONCURRENT_ASSETS, "$", False, None, None),
    (ROW_BS_TOTAL_ASSETS, "$", False, "ROW_CHECK_BS_BALANCES_COUNT", "A = L + E, see Check column"),
    (ROW_BS_DEBT_CURRENT, "$", False, None, None),
    (ROW_BS_AP, "$", False, None, None),
    (ROW_BS_TOTAL_CURRENT_LIAB, "$", False, None, None),
    (ROW_BS_DEBT_NONCURRENT, "$", False, None, None),
    (ROW_BS_TOTAL_NONCURRENT_LIAB, "$", False, None, None),
    (ROW_BS_TOTAL_LIABILITIES, "$", False, None, None),
    (ROW_BS_PAID_IN_CAPITAL, "$", False, None, None),
    (ROW_BS_RETAINED_EARNINGS, "$", False, None, None),
    (ROW_BS_TOTAL_EQUITY, "$", False, None, None),
    (ROW_BS_TOTAL_LIAB_EQUITY, "$", False, None, None),

    (ROW_CFO_NI, "$", True, None, None),
    (ROW_CFO_ADDBACK_DEPR, "$", True, None, None),
    (ROW_CFO_WC_CHANGE, "$", True, None, None),
    (ROW_CFO, "$", True, None, None),
    (ROW_CFI, "$", True, None, None),
    (ROW_CFF_PRINCIPAL, "$", True, None, None),
    (ROW_CFF_DIVIDENDS, "$", True, None, None),
    (ROW_CFF_EQUITY_INJECTION, "$", True, None, None),
    (ROW_CFF, "$", True, None, None),
    (ROW_NET_CHANGE_TOTAL_CASH, "$", True, None, None),
    (ROW_OPENING_TOTAL_CASH, "$", False, None, None),
    (ROW_CLOSING_TOTAL_CASH, "$", False, None, None),
    (ROW_CF_LESS_DSRA, "$", True, None, None),
    (ROW_CF_LESS_MRA, "$", True, None, None),
    (ROW_CLOSING_CASH, "$", False, "ROW_CHECK_CASH_TIES_BUFFER", None),

    (ROW_CFD_RECEIPTS, "$", True, None, None),
    (ROW_CFD_OPEX_PAID, "$", True, None, None),
    (ROW_CFD_INTEREST_PAID, "$", True, None, None),
    (ROW_CFD_TAX_PAID, "$", True, None, None),
    (ROW_CFD_LC_FEE_PAID, "$", True, None, None),
    (ROW_CFO_DIRECT, "$", True, "ROW_CHECK_DIRECT_TIES_INDIRECT_COUNT", None),

    (ROW_FCFF_CF_METHOD, "$", True, None, None),
    (ROW_FCFF_CFADS_METHOD, "$", True, "ROW_CHECK_FCFF_METHODS_TIE_COUNT", None),
    (ROW_FCFE_CF_METHOD, "$", True, None, None),
    (ROW_FCFE_DIVIDEND_METHOD, "$", True, "ROW_CHECK_FCFE_LIFETIME_TIE", None),
]

_CHECK_ROW_BY_NAME = {
    "ROW_CHECK_BS_BALANCES_COUNT": ROW_CHECK_BS_BALANCES_COUNT,
    "ROW_CHECK_CASH_TIES_BUFFER": ROW_CHECK_CASH_TIES_BUFFER,
    "ROW_CHECK_DIRECT_TIES_INDIRECT_COUNT": ROW_CHECK_DIRECT_TIES_INDIRECT_COUNT,
    "ROW_CHECK_FCFF_METHODS_TIE_COUNT": ROW_CHECK_FCFF_METHODS_TIE_COUNT,
    "ROW_CHECK_FCFE_LIFETIME_TIE": ROW_CHECK_FCFE_LIFETIME_TIE,
}


def _build_metadata_columns(ws: Worksheet, first_col_idx: int, last_col_idx: int) -> None:
    """See FS_Quarterly's own version of this function for the full rationale."""
    first_col = _annual_col_letter(first_col_idx - FIRST_DATA_COL)
    last_col = _annual_col_letter(last_col_idx - FIRST_DATA_COL)

    for row, units, has_total, check_name, remarks in _METADATA:
        units_cell = ws.cell(row=row, column=COL_UNITS, value=units)
        units_cell.font = Font(italic=True, size=9, color="FF808080")

        if has_total:
            total_cell = ws.cell(row=row, column=COL_TOTAL,
                                 value=f"=SUM({first_col}{row}:{last_col}{row})")
            total_cell.font = Font(color=COLOR_FORMULA, bold=True)
            total_cell.number_format = "#,##0"

        if check_name:
            check_row = _CHECK_ROW_BY_NAME[check_name]
            check_cell = ws.cell(row=row, column=COL_CHECK,
                                 value=f"={last_col}{check_row}")
            check_cell.font = Font(color=COLOR_LINK)

        if remarks:
            ws.cell(row=row, column=COL_REMARKS, value=remarks).font = Font(italic=True, size=9)


def _annual_col_letter(i: int) -> str:
    return openpyxl.utils.get_column_letter(FIRST_DATA_COL + i)


def _xirr_col_letter(i: int) -> str:
    return openpyxl.utils.get_column_letter(FIRST_DATA_COL + i)


def build_fs_annual(wb: Workbook, timeline: Timeline, inputs: ProjectInputs) -> Worksheet:
    ws = wb.create_sheet("FS_Annual")
    ws.sheet_properties.tabColor = TAB_COLOR_OUTPUT

    ws["A1"] = ("FS_Annual — same P&L / Balance Sheet / Cash Flow as FS_Quarterly, "
                "annualised; PIRR/EIRR via XIRR (Stage 1a)")
    ws["A1"].font = Font(bold=True, size=12)

    _build_flags_block(ws)

    for col, header in ((COL_UNITS, "Units"), (COL_TOTAL, "Total"), (COL_CHECK, "Check"),
                        (COL_REMARKS, "Remarks")):
        cell = ws.cell(row=ROW_COLUMN_HEADER, column=col, value=header)
        cell.font = Font(bold=True, italic=True, size=9)

    ws.cell(row=ROW_YEAR_LABEL, column=1, value="Project Year")

    _section_header(ws, ROW_PNL_HEADER, "PROFIT & LOSS")
    _label(ws, ROW_REVENUE, "Revenue ($) — annual sum")
    _label(ws, ROW_OPEX, "Opex ($) — annual sum")
    _label(ws, ROW_EBITDA, "EBITDA ($) — annual sum")
    _label(ws, ROW_DEPRECIATION, "Depreciation ($) — annual sum")
    _label(ws, ROW_EBIT, "EBIT ($) — annual sum")
    _label(ws, ROW_INTEREST_EXPENSE, "Interest Expense ($) — annual sum")
    _label(ws, ROW_EBT, "EBT ($) — annual sum")
    _label(ws, ROW_TAX, "Tax ($) — annual sum")
    _label(ws, ROW_LC_FEE, "DSRA LC Fee ($) — annual sum")
    _bold_label(ws, ROW_NET_INCOME, "Net Income ($) — annual sum")

    _section_header(ws, ROW_BS_HEADER, "BALANCE SHEET — year-end")
    _bold_label(ws, ROW_BS_ASSETS_HEADER, "ASSETS")
    _sub_label(ws, ROW_BS_CURRENT_ASSETS_HEADER, "Current Assets")
    _label(ws, ROW_BS_CASH, "Cash & Cash Equivalents ($) — unrestricted")
    _label(ws, ROW_BS_AR, "Trade Receivables ($)")
    _bold_label(ws, ROW_BS_TOTAL_CURRENT_ASSETS, "Total Current Assets ($)")
    _sub_label(ws, ROW_BS_NONCURRENT_ASSETS_HEADER, "Non-Current Assets")
    _label(ws, ROW_BS_DSRA, "DSRA Balance ($) — restricted cash")
    _label(ws, ROW_BS_MRA, "MRA Balance ($) — restricted cash")
    _label(ws, ROW_BS_PPE_NET, "PP&E, Net ($)")
    _bold_label(ws, ROW_BS_TOTAL_NONCURRENT_ASSETS, "Total Non-Current Assets ($)")
    _bold_label(ws, ROW_BS_TOTAL_ASSETS, "TOTAL ASSETS ($)")

    _bold_label(ws, ROW_BS_LIABILITIES_HEADER, "LIABILITIES")
    _sub_label(ws, ROW_BS_CURRENT_LIAB_HEADER, "Current Liabilities")
    _label(ws, ROW_BS_DEBT_CURRENT, "Debt — Current Portion ($)")
    _label(ws, ROW_BS_AP, "Trade Payables ($)")
    _bold_label(ws, ROW_BS_TOTAL_CURRENT_LIAB, "Total Current Liabilities ($)")
    _sub_label(ws, ROW_BS_NONCURRENT_LIAB_HEADER, "Non-Current Liabilities")
    _label(ws, ROW_BS_DEBT_NONCURRENT, "Debt — Non-Current Portion ($)")
    _bold_label(ws, ROW_BS_TOTAL_NONCURRENT_LIAB, "Total Non-Current Liabilities ($)")
    _bold_label(ws, ROW_BS_TOTAL_LIABILITIES, "TOTAL LIABILITIES ($)")

    _bold_label(ws, ROW_BS_EQUITY_HEADER, "EQUITY")
    _label(ws, ROW_BS_PAID_IN_CAPITAL, "Paid-in Capital ($)")
    _label(ws, ROW_BS_RETAINED_EARNINGS, "Retained Earnings ($)")
    _bold_label(ws, ROW_BS_TOTAL_EQUITY, "TOTAL EQUITY ($)")

    _bold_label(ws, ROW_BS_TOTAL_LIAB_EQUITY, "TOTAL LIABILITIES + EQUITY ($)")
    _label(ws, ROW_BS_CHECK_A_MINUS_L, "Check: Assets − Liabilities (should equal Total Equity above)")

    _section_header(ws, ROW_CF_HEADER, "CASH FLOW STATEMENT — INDIRECT METHOD")
    _label(ws, ROW_CFO_NI, "Net Income ($) — annual sum")
    _label(ws, ROW_CFO_ADDBACK_DEPR, "Add back: Depreciation ($) — annual sum")
    _label(ws, ROW_CFO_WC_CHANGE, "Change in Working Capital ($) — annual sum; placeholder, see FS_Quarterly")
    _bold_label(ws, ROW_CFO, "Cash Flow from Operations ($) — annual sum")
    _bold_label(ws, ROW_CFI, "Cash Flow from Investing ($) — annual sum")
    _label(ws, ROW_CFF_PRINCIPAL, "Debt Principal Repayment ($) — annual sum")
    _label(ws, ROW_CFF_DIVIDENDS, "Distributions to Equity ($) — annual sum")
    _label(ws, ROW_CFF_EQUITY_INJECTION, "Equity Injections ($) — annual sum")
    _bold_label(ws, ROW_CFF, "Cash Flow from Financing ($) — annual sum")
    _bold_label(ws, ROW_NET_CHANGE_TOTAL_CASH, "Net Change in Total Cash ($) — annual sum, incl. restricted")
    _label(ws, ROW_OPENING_TOTAL_CASH, "Opening Total Cash ($) — start of year")
    _bold_label(ws, ROW_CLOSING_TOTAL_CASH, "Closing Total Cash ($) — year-end")

    _section_header(ws, ROW_CF_RESTRICTED_HEADER, "Reconciliation — Total Cash to Unrestricted Cash")
    _label(ws, ROW_CF_LESS_DSRA, "Less: DSRA Balance ($) — year-end")
    _label(ws, ROW_CF_LESS_MRA, "Less: MRA Balance ($) — year-end")
    _bold_label(ws, ROW_CLOSING_CASH, "Closing Cash & Cash Equivalents ($) — unrestricted, year-end")

    _section_header(ws, ROW_CFD_HEADER, "CASH FLOW STATEMENT — DIRECT METHOD (cross-check on CFO)")
    _label(ws, ROW_CFD_RECEIPTS, "Cash Received from Customers ($) — annual sum")
    _label(ws, ROW_CFD_OPEX_PAID, "Cash Paid for Opex ($) — annual sum")
    _label(ws, ROW_CFD_INTEREST_PAID, "Cash Paid for Interest ($) — annual sum")
    _label(ws, ROW_CFD_TAX_PAID, "Cash Paid for Tax ($) — annual sum")
    _label(ws, ROW_CFD_LC_FEE_PAID, "Cash Paid — DSRA LC Fee ($) — annual sum")
    _bold_label(ws, ROW_CFO_DIRECT, "Cash Flow from Operations ($) — Direct Method, annual sum")
    _label(ws, ROW_CHECK_DIRECT_TIES_INDIRECT, "Check: Direct CFO = Indirect CFO")

    ws.cell(row=ROW_CHECK_HEADER, column=1, value="Checks").font = Font(bold=True)
    _label(ws, ROW_CHECK_BS_BALANCES_COUNT, "# of years where BS does not balance")
    _label(ws, ROW_CHECK_CASH_TIES_BUFFER, "Check: Closing cash ties to FS_Quarterly at year-end")
    _label(ws, ROW_CHECK_DIRECT_TIES_INDIRECT_COUNT, "# of years where Direct CFO != Indirect CFO")

    _section_header(ws, ROW_FCF_HEADER, "FREE CASH FLOW — TWO VIEWS, CROSS-CHECKED (annual sums)")
    _label(ws, ROW_FCFF_CF_METHOD, "FCFF ($) — Cash Flow Statement method (annual sum)")
    _label(ws, ROW_FCFF_CFADS_METHOD, "FCFF ($) — CFADS method (annual sum); used for reported PIRR")
    _label(ws, ROW_CHECK_FCFF_METHODS_TIE_COUNT,
           "Check: whole-of-life FCFF (CF Method) = FCFF (CFADS Method)")
    _label(ws, ROW_FCFE_CF_METHOD,
           "FCFE ($) — Cash Flow Statement method, pre-distribution-policy (annual sum)")
    _label(ws, ROW_FCFE_DIVIDEND_METHOD,
           "FCFE ($) — Dividend method (annual sum); used for reported EIRR")
    _label(ws, ROW_CHECK_FCFE_LIFETIME_TIE,
           "Check: whole-of-life FCFE (Dividend Method) = whole-of-life FCFE (CF Method) + Initial Buffer")

    annual_buckets = _operations_only_annual_buckets(timeline)
    n_years = len(annual_buckets)

    flow_rows = (
        (ROW_REVENUE, fsq.ROW_REVENUE),
        (ROW_OPEX, fsq.ROW_OPEX),
        (ROW_EBITDA, fsq.ROW_EBITDA),
        (ROW_DEPRECIATION, fsq.ROW_DEPRECIATION),
        (ROW_EBIT, fsq.ROW_EBIT),
        (ROW_INTEREST_EXPENSE, fsq.ROW_INTEREST_EXPENSE),
        (ROW_EBT, fsq.ROW_EBT),
        (ROW_TAX, fsq.ROW_TAX),
        (ROW_LC_FEE, fsq.ROW_LC_FEE),
        (ROW_NET_INCOME, fsq.ROW_NET_INCOME),
        (ROW_CFO_NI, fsq.ROW_CFO_NI),
        (ROW_CFO_ADDBACK_DEPR, fsq.ROW_CFO_ADDBACK_DEPR),
        (ROW_CFO_WC_CHANGE, fsq.ROW_CFO_WC_CHANGE),
        (ROW_CFO, fsq.ROW_CFO),
        (ROW_CFI, fsq.ROW_CFI),
        (ROW_CFF_PRINCIPAL, fsq.ROW_CFF_PRINCIPAL),
        (ROW_CFF_DIVIDENDS, fsq.ROW_CFF_DIVIDENDS),
        (ROW_CFF_EQUITY_INJECTION, fsq.ROW_CFF_EQUITY_INJECTION),
        (ROW_CFF, fsq.ROW_CFF),
        (ROW_NET_CHANGE_TOTAL_CASH, fsq.ROW_NET_CHANGE_TOTAL_CASH),
        (ROW_CFD_RECEIPTS, fsq.ROW_CFD_RECEIPTS),
        (ROW_CFD_OPEX_PAID, fsq.ROW_CFD_OPEX_PAID),
        (ROW_CFD_INTEREST_PAID, fsq.ROW_CFD_INTEREST_PAID),
        (ROW_CFD_TAX_PAID, fsq.ROW_CFD_TAX_PAID),
        (ROW_CFD_LC_FEE_PAID, fsq.ROW_CFD_LC_FEE_PAID),
        (ROW_CFO_DIRECT, fsq.ROW_CFO_DIRECT),
        (ROW_FCFF_CF_METHOD, fsq.ROW_FCFF_CF_METHOD),
        (ROW_FCFF_CFADS_METHOD, fsq.ROW_FCFF_CFADS_METHOD),
        (ROW_FCFE_CF_METHOD, fsq.ROW_FCFE_CF_METHOD),
        (ROW_FCFE_DIVIDEND_METHOD, fsq.ROW_FCFE_DIVIDEND_METHOD),
    )
    stock_rows = (
        (ROW_BS_CASH, fsq.ROW_BS_CASH),
        (ROW_BS_AR, fsq.ROW_BS_AR),
        (ROW_BS_AP, fsq.ROW_BS_AP),
        (ROW_BS_DSRA, fsq.ROW_BS_DSRA),
        (ROW_BS_MRA, fsq.ROW_BS_MRA),
        (ROW_BS_PPE_NET, fsq.ROW_BS_PPE_NET),
        (ROW_BS_TOTAL_ASSETS, fsq.ROW_BS_TOTAL_ASSETS),
        (ROW_BS_DEBT_CURRENT, fsq.ROW_BS_DEBT_CURRENT),
        (ROW_BS_DEBT_NONCURRENT, fsq.ROW_BS_DEBT_NONCURRENT),
        (ROW_BS_TOTAL_LIABILITIES, fsq.ROW_BS_TOTAL_LIABILITIES),
        (ROW_BS_PAID_IN_CAPITAL, fsq.ROW_BS_PAID_IN_CAPITAL),
        (ROW_BS_RETAINED_EARNINGS, fsq.ROW_BS_RETAINED_EARNINGS),
        (ROW_BS_TOTAL_EQUITY, fsq.ROW_BS_TOTAL_EQUITY),
        (ROW_BS_TOTAL_LIAB_EQUITY, fsq.ROW_BS_TOTAL_LIAB_EQUITY),
        (ROW_CLOSING_TOTAL_CASH, fsq.ROW_CLOSING_TOTAL_CASH),
        (ROW_CLOSING_CASH, fsq.ROW_CLOSING_CASH),
        (ROW_BS_TOTAL_CURRENT_ASSETS, fsq.ROW_BS_TOTAL_CURRENT_ASSETS),
        (ROW_BS_TOTAL_NONCURRENT_ASSETS, fsq.ROW_BS_TOTAL_NONCURRENT_ASSETS),
        (ROW_BS_TOTAL_CURRENT_LIAB, fsq.ROW_BS_TOTAL_CURRENT_LIAB),
        (ROW_BS_TOTAL_NONCURRENT_LIAB, fsq.ROW_BS_TOTAL_NONCURRENT_LIAB),
    )

    last_year_last_q_col = None
    for year_num, (year_index, quarters) in enumerate(annual_buckets.items()):
        col = _annual_col_letter(year_num)
        ws[f"{col}{ROW_YEAR_LABEL}"] = f"Yr {year_num + 1}"

        q_cols = [_quarterly_source_col(timeline, q) for q in quarters]
        first_q_col, last_q_col = q_cols[0], q_cols[-1]
        last_year_last_q_col = last_q_col

        for dst_row, src_row in flow_rows:
            cell = ws[f"{col}{dst_row}"]
            cell.value = f"=SUM(FS_Quarterly!{first_q_col}{src_row}:{last_q_col}{src_row})"
            cell.font = Font(color=COLOR_LINK)
            cell.number_format = "#,##0"

        for dst_row, src_row in stock_rows:
            cell = ws[f"{col}{dst_row}"]
            cell.value = f"=FS_Quarterly!{last_q_col}{src_row}"
            cell.font = Font(color=COLOR_LINK)
            cell.number_format = "#,##0"

        # Opening total cash: start-of-year, i.e. FS_Quarterly's own opening figure for
        # this year's first quarter — not last year's closing (same value, but this reads
        # the source directly rather than re-deriving it from the prior annual column).
        opening_cell = ws[f"{col}{ROW_OPENING_TOTAL_CASH}"]
        opening_cell.value = f"=FS_Quarterly!{first_q_col}{fsq.ROW_OPENING_TOTAL_CASH}"
        opening_cell.font = Font(color=COLOR_LINK)
        opening_cell.number_format = "#,##0"

        less_dsra = ws[f"{col}{ROW_CF_LESS_DSRA}"]
        less_dsra.value = f"=-{col}{ROW_BS_DSRA}"
        less_dsra.font = Font(color=COLOR_FORMULA)
        less_dsra.number_format = "#,##0"

        less_mra = ws[f"{col}{ROW_CF_LESS_MRA}"]
        less_mra.value = f"=-{col}{ROW_BS_MRA}"
        less_mra.font = Font(color=COLOR_FORMULA)
        less_mra.number_format = "#,##0"

        check_direct = ws[f"{col}{ROW_CHECK_DIRECT_TIES_INDIRECT}"]
        check_direct.value = f"=IF(ROUND({col}{ROW_CFO_DIRECT}-{col}{ROW_CFO},2)=0,1,0)"
        check_direct.font = Font(color=COLOR_FORMULA)

        check_a_minus_l = ws[f"{col}{ROW_BS_CHECK_A_MINUS_L}"]
        check_a_minus_l.value = f"={col}{ROW_BS_TOTAL_ASSETS}-{col}{ROW_BS_TOTAL_LIABILITIES}"
        check_a_minus_l.font = Font(color=COLOR_FORMULA)
        check_a_minus_l.number_format = "#,##0"

    first_col = _annual_col_letter(0)
    last_col = _annual_col_letter(n_years - 1)

    ann_first_col_idx = FIRST_DATA_COL
    ann_last_col_idx = FIRST_DATA_COL + n_years - 1
    for row in (
        ROW_BS_TOTAL_CURRENT_ASSETS, ROW_BS_TOTAL_NONCURRENT_ASSETS,
        ROW_BS_TOTAL_CURRENT_LIAB, ROW_BS_TOTAL_NONCURRENT_LIAB, ROW_BS_TOTAL_LIABILITIES,
        ROW_BS_TOTAL_EQUITY, ROW_CFO, ROW_CFI, ROW_CFF, ROW_NET_CHANGE_TOTAL_CASH,
        ROW_CLOSING_TOTAL_CASH, ROW_CFO_DIRECT,
    ):
        style_total_row(ws, row, ann_first_col_idx, ann_last_col_idx)
    for row in (ROW_NET_INCOME, ROW_BS_TOTAL_ASSETS, ROW_BS_TOTAL_LIAB_EQUITY, ROW_CLOSING_CASH):
        style_total_row(ws, row, ann_first_col_idx, ann_last_col_idx, grand=True)
    for row in (ROW_PNL_HEADER, ROW_BS_HEADER, ROW_CF_HEADER, ROW_CF_RESTRICTED_HEADER,
               ROW_CFD_HEADER, ROW_FCF_HEADER):
        style_section_header_row(ws, row, ann_first_col_idx, ann_last_col_idx)

    _build_metadata_columns(ws, ann_first_col_idx, ann_last_col_idx)

    bs_check = ws[f"{last_col}{ROW_CHECK_BS_BALANCES_COUNT}"]
    bs_check.value = (
        f"=SUMPRODUCT(--(ROUND({first_col}{ROW_BS_TOTAL_ASSETS}:{last_col}{ROW_BS_TOTAL_ASSETS}"
        f"-{first_col}{ROW_BS_TOTAL_LIAB_EQUITY}:{last_col}{ROW_BS_TOTAL_LIAB_EQUITY},2)<>0))"
    )
    bs_check.font = Font(color=COLOR_FORMULA)

    # Independent tie-out at the final year-end only: FS_Annual's own closing cash against
    # Calc_CFADS' buffer, the waterfall's own derivation — same check FS_Quarterly runs
    # every quarter, confirmed once more here at annual granularity.
    cash_tie = ws[f"{last_col}{ROW_CHECK_CASH_TIES_BUFFER}"]
    cash_tie.value = (
        f"=IF(ROUND({last_col}{ROW_CLOSING_CASH}-Calc_CFADS!{last_year_last_q_col}"
        f"{cfads.ROW_BUFFER_CLOSING},2)=0,1,0)"
    )
    cash_tie.font = Font(color=COLOR_FORMULA)

    direct_tie_count = ws[f"{last_col}{ROW_CHECK_DIRECT_TIES_INDIRECT_COUNT}"]
    direct_tie_count.value = (
        f"=SUMPRODUCT(--({first_col}{ROW_CHECK_DIRECT_TIES_INDIRECT}:"
        f"{last_col}{ROW_CHECK_DIRECT_TIES_INDIRECT}=0))"
    )
    direct_tie_count.font = Font(color=COLOR_FORMULA)

    # Whole-of-life tie, not per-year — see FS_Quarterly for why (Calc_CFADS' FCFF is
    # deliberately ex-working-capital, same convention as PIRR excluding IDC).
    fcff_tie = ws[f"{last_col}{ROW_CHECK_FCFF_METHODS_TIE_COUNT}"]
    fcff_tie.value = (
        f"=IF(ROUND(SUM({first_col}{ROW_FCFF_CF_METHOD}:{last_col}{ROW_FCFF_CF_METHOD})"
        f"-SUM({first_col}{ROW_FCFF_CFADS_METHOD}:{last_col}{ROW_FCFF_CFADS_METHOD}),2)=0,1,0)"
    )
    fcff_tie.font = Font(color=COLOR_FORMULA)

    fcfe_lifetime_tie = ws[f"{last_col}{ROW_CHECK_FCFE_LIFETIME_TIE}"]
    fcfe_lifetime_tie.value = (
        f"=IF(ROUND(SUM({first_col}{ROW_FCFE_DIVIDEND_METHOD}:{last_col}{ROW_FCFE_DIVIDEND_METHOD})"
        f"-SUM({first_col}{ROW_FCFE_CF_METHOD}:{last_col}{ROW_FCFE_CF_METHOD})"
        f"-Calc_Financing_Cons!{fin_cons.ABS_INITIAL_BUFFER},2)=0,1,0)"
    )
    fcfe_lifetime_tie.font = Font(color=COLOR_FORMULA)

    _build_xirr_block(ws, wb, timeline)

    first_xirr_col = _xirr_col_letter(0)

    ws.cell(row=ROW_PIRR_LABEL, column=1, value="Project IRR (PIRR) — unlevered, XIRR")
    pirr_cell = ws.cell(row=ROW_PIRR_VALUE, column=1)
    n_periods = len(timeline.construction_months) + len(timeline.operations_quarters)
    last_xirr_col = _xirr_col_letter(n_periods - 1)
    pirr_cell.value = (
        f"=XIRR({first_xirr_col}{ROW_XIRR_PROJECT_CF}:{last_xirr_col}{ROW_XIRR_PROJECT_CF},"
        f"{first_xirr_col}{ROW_XIRR_DATE}:{last_xirr_col}{ROW_XIRR_DATE})"
    )
    pirr_cell.font = Font(color=COLOR_FORMULA, bold=True)
    pirr_cell.number_format = "0.00%"

    ws.cell(row=ROW_EIRR_LABEL, column=1, value="Equity IRR (EIRR) — levered, XIRR")
    eirr_cell = ws.cell(row=ROW_EIRR_VALUE, column=1)
    eirr_cell.value = (
        f"=XIRR({first_xirr_col}{ROW_XIRR_EQUITY_CF}:{last_xirr_col}{ROW_XIRR_EQUITY_CF},"
        f"{first_xirr_col}{ROW_XIRR_DATE}:{last_xirr_col}{ROW_XIRR_DATE})"
    )
    eirr_cell.font = Font(color=COLOR_FORMULA, bold=True)
    eirr_cell.number_format = "0.00%"

    _build_construction_section(ws, timeline)
    _build_construction_cash_flow(ws, timeline)

    _add_named_range(wb, "FSA_PIRR", "FS_Annual", f"A{ROW_PIRR_VALUE}")
    _add_named_range(wb, "FSA_EIRR", "FS_Annual", f"A{ROW_EIRR_VALUE}")

    ws.freeze_panes = ws.cell(row=ROW_PNL_HEADER, column=FIRST_DATA_COL)
    ws.column_dimensions["A"].width = 62
    ws.column_dimensions["B"].width = 8
    ws.column_dimensions["C"].width = 14
    ws.column_dimensions["D"].width = 10
    ws.column_dimensions["E"].width = 34
    ws.column_dimensions["F"].width = 3

    return ws


def _build_construction_section(ws: Worksheet, timeline: Timeline) -> None:
    """Construction years reported separately: the source data is monthly (not quarterly),
    and during construction there is no P&L activity — only capitalised spend building the
    balance sheet, which is what lenders look at pre-COD."""
    ws.cell(row=ROW_CONS_HEADER, column=1, value="Construction Period (annual rollup of monthly data)")
    ws.cell(row=ROW_CONS_HEADER, column=1).font = Font(bold=True)

    ws.cell(row=ROW_CONS_YEAR_LABEL, column=1, value="Construction Year")
    ws.cell(row=ROW_CONS_CAPEX, column=1, value="Capex Incurred ($) — annual sum")
    ws.cell(row=ROW_CONS_IDC, column=1, value="IDC Capitalised ($) — annual sum")
    ws.cell(row=ROW_CONS_DEBT_DRAWN, column=1, value="Debt Drawn ($) — annual sum")
    ws.cell(row=ROW_CONS_EQUITY_DRAWN, column=1, value="Equity Drawn ($) — annual sum")
    ws.cell(row=ROW_CONS_CUM_TPC, column=1, value="Cumulative Total Project Cost ($) — year-end")
    ws.cell(row=ROW_CONS_CLOSING_DEBT, column=1, value="Closing Debt Balance ($) — year-end")

    ws.cell(row=ROW_CONS_BS_HEADER, column=1,
            value="Balance Sheet (year-end) — P&L is empty pre-COD, no revenue").font = Font(bold=True)
    ws.cell(row=ROW_CONS_BS_CASH, column=1, value="Cash ($) — unrestricted; funded at Financial Close only")
    ws.cell(row=ROW_CONS_BS_DSRA, column=1, value="DSRA Balance ($) — funded at Financial Close only")
    ws.cell(row=ROW_CONS_BS_PPE, column=1, value="PP&E, Net ($) = Cumulative Capex + Cumulative IDC")
    ws.cell(row=ROW_CONS_BS_TOTAL_ASSETS, column=1, value="Total Assets ($)")
    ws.cell(row=ROW_CONS_BS_DEBT, column=1, value="Debt ($)")
    ws.cell(row=ROW_CONS_BS_PAID_IN_CAPITAL, column=1, value="Paid-in Capital ($)")
    ws.cell(row=ROW_CONS_BS_RETAINED_EARNINGS, column=1, value="Retained Earnings ($) — zero pre-COD, no P&L")
    ws.cell(row=ROW_CONS_BS_TOTAL_EQUITY, column=1, value="Total Equity ($)")
    ws.cell(row=ROW_CONS_BS_TOTAL_LIAB_EQUITY, column=1, value="Total Liabilities + Equity ($)")

    ws.cell(row=ROW_CONS_CHECK_HEADER, column=1, value="Checks").font = Font(bold=True)
    ws.cell(row=ROW_CONS_CHECK_BS_BALANCES_COUNT, column=1,
            value="# of construction years where the BS does not balance")

    buckets: dict = {}
    for idx, period in enumerate(timeline.construction_months):
        buckets.setdefault(period.year_index, []).append(idx)
    sorted_buckets = sorted(buckets.items())
    n_cons_years = len(sorted_buckets)

    for year_num, (_, month_indices) in enumerate(sorted_buckets):
        col = _annual_col_letter(year_num)
        first_m = openpyxl.utils.get_column_letter(FIRST_DATA_COL + month_indices[0])
        last_m = openpyxl.utils.get_column_letter(FIRST_DATA_COL + month_indices[-1])
        is_last_cons_year = year_num == n_cons_years - 1

        ws[f"{col}{ROW_CONS_YEAR_LABEL}"] = f"C-Yr {year_num + 1}"

        for row, (sheet, src_row, is_sum) in {
            ROW_CONS_CAPEX: ("Calc_Capex", capex.ROW_CAPEX_DRAW, True),
            ROW_CONS_IDC: ("Calc_Financing_Cons", fin_cons.ROW_INTEREST_ACCRUED, True),
            ROW_CONS_DEBT_DRAWN: ("Calc_Financing_Cons", fin_cons.ROW_DEBT_DRAW, True),
            ROW_CONS_EQUITY_DRAWN: ("Calc_Financing_Cons", fin_cons.ROW_EQUITY_DRAW, True),
            ROW_CONS_CUM_TPC: ("Calc_Capex", capex.ROW_TOTAL_PROJECT_COST, False),
            ROW_CONS_CLOSING_DEBT: ("Calc_Financing_Cons", fin_cons.ROW_CLOSING_BAL, False),
        }.items():
            cell = ws[f"{col}{row}"]
            if is_sum:
                cell.value = f"=SUM({sheet}!{first_m}{src_row}:{last_m}{src_row})"
            else:
                cell.value = f"={sheet}!{last_m}{src_row}"
            cell.font = Font(color=COLOR_LINK)
            cell.number_format = "#,##0"

        # DSRA and unrestricted cash are only ever funded in the final construction month
        # (see Calc_Financing_Cons: Initial DSRA/Buffer are added to that month's funding
        # requirement alone), so every year before COD is genuinely zero here, not omitted.
        dsra_cell = ws[f"{col}{ROW_CONS_BS_DSRA}"]
        if is_last_cons_year:
            dsra_cell.value = f"=Calc_Financing_Cons!{fin_cons.ABS_INITIAL_DSRA}"
            dsra_cell.font = Font(color=COLOR_LINK)
        else:
            dsra_cell.value = 0.0
            dsra_cell.font = Font(color=COLOR_FORMULA)
        dsra_cell.number_format = "#,##0"

        cash_cell = ws[f"{col}{ROW_CONS_BS_CASH}"]
        cash_cell.value = (
            f"=Calc_Financing_Cons!{last_m}{fin_cons.ROW_CUM_DEBT_DRAW}"
            f"+Calc_Financing_Cons!{last_m}{fin_cons.ROW_CUM_EQUITY_DRAW}"
            f"-Calc_Capex!{last_m}{capex.ROW_CUM_CAPEX_DRAW}"
            f"-Calc_Capex!{last_m}{capex.ROW_CUM_IDC}"
            f"-{col}{ROW_CONS_BS_DSRA}"
        )
        cash_cell.font = Font(color=COLOR_FORMULA)
        cash_cell.number_format = "#,##0"

        ppe_cell = ws[f"{col}{ROW_CONS_BS_PPE}"]
        ppe_cell.value = (
            f"=Calc_Capex!{last_m}{capex.ROW_CUM_CAPEX_DRAW}+Calc_Capex!{last_m}{capex.ROW_CUM_IDC}"
        )
        ppe_cell.font = Font(color=COLOR_LINK)
        ppe_cell.number_format = "#,##0"

        assets_cell = ws[f"{col}{ROW_CONS_BS_TOTAL_ASSETS}"]
        assets_cell.value = f"={col}{ROW_CONS_BS_CASH}+{col}{ROW_CONS_BS_DSRA}+{col}{ROW_CONS_BS_PPE}"
        assets_cell.font = Font(color=COLOR_FORMULA)
        assets_cell.number_format = "#,##0"

        debt_bs_cell = ws[f"{col}{ROW_CONS_BS_DEBT}"]
        debt_bs_cell.value = f"={col}{ROW_CONS_CLOSING_DEBT}"
        debt_bs_cell.font = Font(color=COLOR_FORMULA)
        debt_bs_cell.number_format = "#,##0"

        paid_in_cell = ws[f"{col}{ROW_CONS_BS_PAID_IN_CAPITAL}"]
        paid_in_cell.value = f"=Calc_Financing_Cons!{last_m}{fin_cons.ROW_CUM_EQUITY_DRAW}"
        paid_in_cell.font = Font(color=COLOR_LINK)
        paid_in_cell.number_format = "#,##0"

        re_cell = ws[f"{col}{ROW_CONS_BS_RETAINED_EARNINGS}"]
        re_cell.value = 0.0
        re_cell.font = Font(color=COLOR_FORMULA)
        re_cell.number_format = "#,##0"

        equity_cell = ws[f"{col}{ROW_CONS_BS_TOTAL_EQUITY}"]
        equity_cell.value = f"={col}{ROW_CONS_BS_PAID_IN_CAPITAL}+{col}{ROW_CONS_BS_RETAINED_EARNINGS}"
        equity_cell.font = Font(color=COLOR_FORMULA)
        equity_cell.number_format = "#,##0"

        liab_eq_cell = ws[f"{col}{ROW_CONS_BS_TOTAL_LIAB_EQUITY}"]
        liab_eq_cell.value = f"={col}{ROW_CONS_BS_DEBT}+{col}{ROW_CONS_BS_TOTAL_EQUITY}"
        liab_eq_cell.font = Font(color=COLOR_FORMULA)
        liab_eq_cell.number_format = "#,##0"

    cons_first_col_idx = FIRST_DATA_COL
    cons_last_col_idx = FIRST_DATA_COL + n_cons_years - 1
    style_total_row(ws, ROW_CONS_BS_TOTAL_ASSETS, cons_first_col_idx, cons_last_col_idx, grand=True)
    style_total_row(ws, ROW_CONS_BS_TOTAL_LIAB_EQUITY, cons_first_col_idx, cons_last_col_idx, grand=True)
    style_total_row(ws, ROW_CONS_BS_TOTAL_EQUITY, cons_first_col_idx, cons_last_col_idx)
    style_section_header_row(ws, ROW_CONS_BS_HEADER, cons_first_col_idx, cons_last_col_idx)

    first_cons_col = _annual_col_letter(0)
    last_cons_col = _annual_col_letter(n_cons_years - 1)
    cons_bs_check = ws[f"{last_cons_col}{ROW_CONS_CHECK_BS_BALANCES_COUNT}"]
    cons_bs_check.value = (
        f"=SUMPRODUCT(--(ROUND({first_cons_col}{ROW_CONS_BS_TOTAL_ASSETS}:{last_cons_col}{ROW_CONS_BS_TOTAL_ASSETS}"
        f"-{first_cons_col}{ROW_CONS_BS_TOTAL_LIAB_EQUITY}:{last_cons_col}{ROW_CONS_BS_TOTAL_LIAB_EQUITY},2)<>0))"
    )
    cons_bs_check.font = Font(color=COLOR_FORMULA)


def _build_construction_cash_flow(ws: Worksheet, timeline: Timeline) -> None:
    """Monthly Construction Cash Flow Statement, matching Calc_Capex's own column
    resolution. This is what the XIRR helper's construction-period Equity CF reads from —
    see the note there for why Project CF (PIRR) deliberately does not."""
    ws.cell(row=ROW_CONS_CF_HEADER, column=1,
            value="CASH FLOW STATEMENT — construction period (monthly)").font = Font(
        bold=True, size=11, underline="single")
    ws.cell(row=ROW_CONS_CF_INVESTING, column=1,
            value="Cash Flow from Investing ($) = -(Capex Draw + IDC Capitalised)")
    ws.cell(row=ROW_CONS_CF_DEBT_DRAWN, column=1, value="Cash Flow from Financing — Debt Drawn ($)")
    ws.cell(row=ROW_CONS_CF_EQUITY_DRAWN, column=1, value="Cash Flow from Financing — Equity Drawn ($)")
    ws.cell(row=ROW_CONS_CF_FINANCING, column=1, value="Cash Flow from Financing ($)")
    ws.cell(row=ROW_CONS_CF_NET, column=1,
            value="Net Cash Flow ($) — 0 every month except the last (DSRA/Buffer funded then)")

    ws.cell(row=ROW_CONS_CF_CHECK_HEADER, column=1, value="Checks").font = Font(bold=True)
    ws.cell(row=ROW_CONS_CF_CHECK_NET_ZERO_COUNT, column=1,
            value="# of months (excl. the last) where Net Cash Flow != 0")

    n_cons = len(timeline.construction_months)
    for i in range(n_cons):
        col = _xirr_col_letter(i)

        investing_cell = ws[f"{col}{ROW_CONS_CF_INVESTING}"]
        investing_cell.value = (
            f"=-Calc_Capex!{col}{capex.ROW_CAPEX_DRAW}-Calc_Financing_Cons!{col}{fin_cons.ROW_INTEREST_ACCRUED}"
        )
        investing_cell.font = Font(color=COLOR_LINK)
        investing_cell.number_format = "#,##0"

        debt_cell = ws[f"{col}{ROW_CONS_CF_DEBT_DRAWN}"]
        debt_cell.value = f"=Calc_Financing_Cons!{col}{fin_cons.ROW_DEBT_DRAW}"
        debt_cell.font = Font(color=COLOR_LINK)
        debt_cell.number_format = "#,##0"

        equity_cell = ws[f"{col}{ROW_CONS_CF_EQUITY_DRAWN}"]
        equity_cell.value = f"=Calc_Financing_Cons!{col}{fin_cons.ROW_EQUITY_DRAW}"
        equity_cell.font = Font(color=COLOR_LINK)
        equity_cell.number_format = "#,##0"

        financing_cell = ws[f"{col}{ROW_CONS_CF_FINANCING}"]
        financing_cell.value = f"={col}{ROW_CONS_CF_DEBT_DRAWN}+{col}{ROW_CONS_CF_EQUITY_DRAWN}"
        financing_cell.font = Font(color=COLOR_FORMULA, bold=True)
        financing_cell.number_format = "#,##0"

        net_cell = ws[f"{col}{ROW_CONS_CF_NET}"]
        net_cell.value = f"={col}{ROW_CONS_CF_INVESTING}+{col}{ROW_CONS_CF_FINANCING}"
        net_cell.font = Font(color=COLOR_FORMULA, bold=True)
        net_cell.number_format = "#,##0"

    cf_first_col_idx = FIRST_DATA_COL
    cf_last_col_idx = FIRST_DATA_COL + n_cons - 1
    style_total_row(ws, ROW_CONS_CF_FINANCING, cf_first_col_idx, cf_last_col_idx)
    style_total_row(ws, ROW_CONS_CF_NET, cf_first_col_idx, cf_last_col_idx, grand=True)
    style_section_header_row(ws, ROW_CONS_CF_HEADER, cf_first_col_idx, cf_last_col_idx)

    first_col = _xirr_col_letter(0)
    penultimate_col = _xirr_col_letter(n_cons - 2)
    check_cell = ws[f"{_xirr_col_letter(n_cons - 1)}{ROW_CONS_CF_CHECK_NET_ZERO_COUNT}"]
    check_cell.value = (
        f"=SUMPRODUCT(--(ROUND({first_col}{ROW_CONS_CF_NET}:{penultimate_col}{ROW_CONS_CF_NET},2)<>0))"
    )
    check_cell.font = Font(color=COLOR_FORMULA)


def _build_xirr_block(ws: Worksheet, wb: Workbook, timeline: Timeline) -> None:
    ws.cell(row=ROW_XIRR_DATE, column=1, value="XIRR Helper: Period End Date")
    ws.cell(row=ROW_XIRR_PROJECT_CF, column=1, value="XIRR Helper: Project CF (unlevered) — construction outflow, FCFF inflow")
    ws.cell(row=ROW_XIRR_EQUITY_CF, column=1, value="XIRR Helper: Equity CF (levered) — equity draw outflow, FCFE inflow")

    n_cons = len(timeline.construction_months)

    # Construction months: outflows
    for i, period in enumerate(timeline.construction_months):
        col = _xirr_col_letter(i)
        capex_col_in_source = _xirr_col_letter(i)  # matches Calc_Capex's own column layout

        date_cell = ws[f"{col}{ROW_XIRR_DATE}"]
        date_cell.value = f"=Calc_Capex!{capex_col_in_source}{capex.ROW_DATE_HEADER}"
        date_cell.font = Font(color=COLOR_LINK)
        date_cell.number_format = "mmm-yy"

        # Deliberately excludes IDC: PIRR is the return to a hypothetical all-equity,
        # unlevered project, so the outflow is the capex spend itself — IDC only exists
        # because the actual project is debt-financed, and including it would let the
        # financing structure leak into a return that is supposed to be capital-structure-
        # independent. This is why it's sourced from Calc_Capex directly rather than the
        # Construction Cash Flow Statement below, whose Investing line correctly includes
        # IDC for accounting purposes.
        proj_cf_cell = ws[f"{col}{ROW_XIRR_PROJECT_CF}"]
        proj_cf_cell.value = f"=-Calc_Capex!{capex_col_in_source}{capex.ROW_CAPEX_DRAW}"
        proj_cf_cell.font = Font(color=COLOR_LINK)
        proj_cf_cell.number_format = "#,##0"

        # The equity investor's real cash flow, sourced from the Construction Cash Flow
        # Statement below rather than Calc_Capex directly — same figure (Equity Drawn
        # already correctly includes the investor's pro-rata share of IDC), now flowing
        # through a built statement instead of skipping past it.
        equity_cf_cell = ws[f"{col}{ROW_XIRR_EQUITY_CF}"]
        equity_cf_cell.value = f"=-{col}{ROW_CONS_CF_EQUITY_DRAWN}"
        equity_cf_cell.font = Font(color=COLOR_FORMULA)
        equity_cf_cell.number_format = "#,##0"

    # Operations quarters: inflows
    for i, period in enumerate(timeline.operations_quarters):
        col = _xirr_col_letter(n_cons + i)
        ops_col_in_source = _xirr_col_letter(i)  # matches Calc_CFADS's own column layout

        date_cell = ws[f"{col}{ROW_XIRR_DATE}"]
        date_cell.value = f"=Calc_CFADS!{ops_col_in_source}{cfads.ROW_DATE_HEADER}"
        date_cell.font = Font(color=COLOR_LINK)
        date_cell.number_format = "mmm-yy"

        # FCFF and FCFE are both built once on Calc_CFADS and only referenced here.
        proj_cf_cell = ws[f"{col}{ROW_XIRR_PROJECT_CF}"]
        proj_cf_cell.value = f"=Calc_CFADS!{ops_col_in_source}{cfads.ROW_FCFF}"
        proj_cf_cell.font = Font(color=COLOR_LINK)
        proj_cf_cell.number_format = "#,##0"

        equity_cf_cell = ws[f"{col}{ROW_XIRR_EQUITY_CF}"]
        equity_cf_cell.value = f"=Calc_CFADS!{ops_col_in_source}{cfads.ROW_FCFE}"
        equity_cf_cell.font = Font(color=COLOR_LINK)
        equity_cf_cell.number_format = "#,##0"


def n_construction_years(timeline: Timeline) -> int:
    """Exposed so callers (e.g. Check_Control) can locate the last construction-year
    column without re-deriving the year bucketing this module already does."""
    years = {p.year_index for p in timeline.construction_months}
    return len(years)


def _operations_only_annual_buckets(timeline: Timeline) -> dict:
    """FS_Annual reports operating years only (Stage 1a); construction-period BS
    will be added when Calc_Financing_Cons circularity (Stage 1b) makes construction
    P&L/BS activity meaningful to show."""
    buckets: dict = {}
    for p in timeline.operations_quarters:
        buckets.setdefault(p.year_index, []).append(p)
    return dict(sorted(buckets.items()))


def _quarterly_source_col(timeline: Timeline, period) -> str:
    idx = timeline.operations_quarters.index(period)
    return openpyxl.utils.get_column_letter(FIRST_DATA_COL + idx)


def _build_flags_block(ws: Worksheet) -> None:
    """Same three flags, same rows, as FS_Quarterly — see the note there."""
    import assumptions_constant as const
    import check_control

    _label(ws, ROW_FLAG_ACTIVE_SCENARIO, "Active Scenario")
    scenario_cell = ws[f"B{ROW_FLAG_ACTIVE_SCENARIO}"]
    scenario_cell.value = f"=ActiveScenario&\" - \"&Assumptions_Constant!$B${const.ROW_SCENARIO_NAME}"
    scenario_cell.font = Font(color=COLOR_LINK, bold=True)

    _label(ws, ROW_FLAG_MODEL_STATUS, "Model Status")
    status_cell = ws[f"B{ROW_FLAG_MODEL_STATUS}"]
    status_cell.value = f"=Check_Control!B{check_control.ROW_MASTER_FLAG}"
    status_cell.font = Font(color=COLOR_LINK, bold=True)

    _label(ws, ROW_FLAG_SOLVE_FRESHNESS, "Solve Freshness")
    freshness_cell = ws[f"B{ROW_FLAG_SOLVE_FRESHNESS}"]
    freshness_cell.value = "=SolveStatus"
    freshness_cell.font = Font(color=COLOR_LINK, bold=True)


def _label(ws: Worksheet, row: int, label: str) -> None:
    ws.cell(row=row, column=1, value=label)


def _bold_label(ws: Worksheet, row: int, label: str) -> None:
    ws.cell(row=row, column=1, value=label).font = Font(bold=True)


def _sub_label(ws: Worksheet, row: int, label: str) -> None:
    ws.cell(row=row, column=1, value=label).font = Font(italic=True)


def _section_header(ws: Worksheet, row: int, title: str) -> None:
    cell = ws.cell(row=row, column=1, value=title)
    cell.font = Font(bold=True, size=11, underline="single")


def _add_named_range(wb: Workbook, name: str, sheet: str, cell: str) -> None:
    import re

    from openpyxl.workbook.defined_name import DefinedName

    match = re.match(r"([A-Za-z]+)(\d+)", cell)
    col, row = match.group(1), match.group(2)
    wb.defined_names.add(DefinedName(name, attr_text=f"'{sheet}'!${col}${row}"))
