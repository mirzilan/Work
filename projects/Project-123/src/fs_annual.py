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
    COLOR_FORMULA,
    COLOR_LINK,
    TAB_COLOR_OUTPUT,
)
import openpyxl.utils

# FS_Annual mirrors FS_Quarterly's own layout section-for-section (P&L, then Balance
# Sheet, then Cash Flow) — every row here is either an annual SUM of a FS_Quarterly flow
# row or a year-end (last quarter) pull of a FS_Quarterly stock row. PIRR/EIRR sit below
# the three statements, since they're a whole-of-life output, not part of any one of them.

ROW_YEAR_LABEL = 2

# --- P&L (annual sums) ---
ROW_PNL_HEADER = 4
ROW_REVENUE = 5
ROW_OPEX = 6
ROW_EBITDA = 7
ROW_DEPRECIATION = 8
ROW_EBIT = 9
ROW_INTEREST_EXPENSE = 10
ROW_EBT = 11
ROW_TAX = 12
ROW_LC_FEE = 13
ROW_NET_INCOME = 14

# --- Balance Sheet (year-end, i.e. Q4 of each project year) ---
ROW_BS_HEADER = 16
ROW_BS_ASSETS_HEADER = 17
ROW_BS_CURRENT_ASSETS_HEADER = 18
ROW_BS_CASH = 19
ROW_BS_TOTAL_CURRENT_ASSETS = 20
ROW_BS_NONCURRENT_ASSETS_HEADER = 21
ROW_BS_DSRA = 22
ROW_BS_MRA = 23
ROW_BS_PPE_NET = 24
ROW_BS_TOTAL_NONCURRENT_ASSETS = 25
ROW_BS_TOTAL_ASSETS = 26

ROW_BS_LIABILITIES_HEADER = 28
ROW_BS_CURRENT_LIAB_HEADER = 29
ROW_BS_DEBT_CURRENT = 30
ROW_BS_TOTAL_CURRENT_LIAB = 31
ROW_BS_NONCURRENT_LIAB_HEADER = 32
ROW_BS_DEBT_NONCURRENT = 33
ROW_BS_TOTAL_NONCURRENT_LIAB = 34
ROW_BS_TOTAL_LIABILITIES = 35

ROW_BS_EQUITY_HEADER = 37
ROW_BS_PAID_IN_CAPITAL = 38
ROW_BS_RETAINED_EARNINGS = 39
ROW_BS_TOTAL_EQUITY = 40

ROW_BS_TOTAL_LIAB_EQUITY = 42
ROW_BS_CHECK_A_MINUS_L = 43

# --- Cash Flow (annual sums for flows, year-end for the cash walk) ---
ROW_CF_HEADER = 45
ROW_CFO_NI = 46
ROW_CFO_ADDBACK_DEPR = 47
ROW_CFO_WC_CHANGE = 48
ROW_CFO = 49
ROW_CFI = 50
ROW_CFF_PRINCIPAL = 51
ROW_CFF_DIVIDENDS = 52
ROW_CFF_EQUITY_INJECTION = 53
ROW_CFF = 54
ROW_NET_CHANGE_TOTAL_CASH = 55
ROW_OPENING_TOTAL_CASH = 56
ROW_CLOSING_TOTAL_CASH = 57

ROW_CF_RESTRICTED_HEADER = 59
ROW_CF_LESS_DSRA = 60
ROW_CF_LESS_MRA = 61
ROW_CLOSING_CASH = 62

ROW_CFD_HEADER = 64
ROW_CFD_RECEIPTS = 65
ROW_CFD_OPEX_PAID = 66
ROW_CFD_INTEREST_PAID = 67
ROW_CFD_TAX_PAID = 68
ROW_CFD_LC_FEE_PAID = 69
ROW_CFO_DIRECT = 70
ROW_CHECK_DIRECT_TIES_INDIRECT = 71

ROW_CHECK_HEADER = 73
ROW_CHECK_BS_BALANCES_COUNT = 74
ROW_CHECK_CASH_TIES_BUFFER = 75
ROW_CHECK_DIRECT_TIES_INDIRECT_COUNT = 76

# FCFF / FCFE — two views, cross-checked; annual sums of FS_Quarterly's own dual-view rows.
# See FS_Quarterly for the full explanation of why FCFF ties exactly every period while
# FCFE only ties in total (buffer/lock-up timing).
ROW_FCF_HEADER = 78
ROW_FCFF_CF_METHOD = 79
ROW_FCFF_CFADS_METHOD = 80
ROW_CHECK_FCFF_METHODS_TIE_COUNT = 81
ROW_FCFE_CF_METHOD = 82
ROW_FCFE_DIVIDEND_METHOD = 83
ROW_CHECK_FCFE_LIFETIME_TIE = 84

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
ROW_XIRR_DATE = 87
ROW_XIRR_PROJECT_CF = 88
ROW_XIRR_EQUITY_CF = 89

ROW_PIRR_LABEL = 92
ROW_PIRR_VALUE = 93
ROW_EIRR_LABEL = 94
ROW_EIRR_VALUE = 95

# Construction-period annual summary (monthly source data rolled to project years)
ROW_CONS_HEADER = 98
ROW_CONS_YEAR_LABEL = 99
ROW_CONS_CAPEX = 100
ROW_CONS_IDC = 101
ROW_CONS_DEBT_DRAWN = 102
ROW_CONS_EQUITY_DRAWN = 103
ROW_CONS_CUM_TPC = 104
ROW_CONS_CLOSING_DEBT = 105

# Construction-period Balance Sheet — year-end, sourced from the same monthly cells as the
# summary above. P&L is legitimately empty pre-COD (no revenue), but the BS still has to
# provably balance every year so the Day-1 operating BS is derived, not asserted.
ROW_CONS_BS_HEADER = 107
ROW_CONS_BS_CASH = 108
ROW_CONS_BS_DSRA = 109
ROW_CONS_BS_PPE = 110
ROW_CONS_BS_TOTAL_ASSETS = 111
ROW_CONS_BS_DEBT = 112
ROW_CONS_BS_PAID_IN_CAPITAL = 113
ROW_CONS_BS_RETAINED_EARNINGS = 114
ROW_CONS_BS_TOTAL_EQUITY = 115
ROW_CONS_BS_TOTAL_LIAB_EQUITY = 116

ROW_CONS_CHECK_HEADER = 118
ROW_CONS_CHECK_BS_BALANCES_COUNT = 119  # count of construction years where the BS does not balance

# Construction Cash Flow Statement — monthly (matches Calc_Capex's own resolution), so the
# XIRR helper's construction-period Equity CF reads from a built statement rather than
# reaching into Calc_Capex directly. Investing includes capitalised interest (IDC is a
# real investment, not a financing cost) so Net Cash Flow ties to zero every month except
# the last, when the DSRA/Buffer are funded — the same cash pattern the construction BS
# above already assumes.
ROW_CONS_CF_HEADER = 121
ROW_CONS_CF_INVESTING = 122     # -(Capex Draw + IDC)
ROW_CONS_CF_DEBT_DRAWN = 123
ROW_CONS_CF_EQUITY_DRAWN = 124
ROW_CONS_CF_FINANCING = 125     # Debt Drawn + Equity Drawn
ROW_CONS_CF_NET = 126

ROW_CONS_CF_CHECK_HEADER = 128
ROW_CONS_CF_CHECK_NET_ZERO_COUNT = 129  # # of months where Net CF != 0 (all but the last)


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
    _label(ws, ROW_CHECK_FCFF_METHODS_TIE_COUNT, "# of years where the two FCFF views differ (should be 0)")
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

    fcff_tie = ws[f"{last_col}{ROW_CHECK_FCFF_METHODS_TIE_COUNT}"]
    fcff_tie.value = (
        f"=SUMPRODUCT(--(ROUND({first_col}{ROW_FCFF_CF_METHOD}:{last_col}{ROW_FCFF_CF_METHOD}"
        f"-{first_col}{ROW_FCFF_CFADS_METHOD}:{last_col}{ROW_FCFF_CFADS_METHOD},2)<>0))"
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
        capex_col_in_source = openpyxl.utils.get_column_letter(3 + i)  # matches Calc_Capex's own column layout

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
        ops_col_in_source = openpyxl.utils.get_column_letter(3 + i)  # matches Calc_CFADS's own column layout

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
