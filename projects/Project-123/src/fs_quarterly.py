from openpyxl.styles import Font
from openpyxl.worksheet.worksheet import Worksheet
from openpyxl.workbook import Workbook

import calc_capex as capex
import calc_cfads as cfads
import calc_financing_cons as fin_cons
import calc_financing_ops as fin_ops
import calc_tax as tax
from inputs import ProjectInputs
from timeline import Timeline
# Source-sheet row constants, imported rather than restated — see the note in calc_tax.py.
from calc_revenue_opex import (
    ROW_REVENUE as REVOPEX_ROW_REVENUE,
    ROW_OPEX as REVOPEX_ROW_OPEX,
    ROW_EBITDA as REVOPEX_ROW_EBITDA,
)
from workbook_builder import (
    FIRST_DATA_COL,
    COLOR_FORMULA,
    COLOR_LINK,
    TAB_COLOR_OUTPUT,
    col_letter,
)

ROW_DATE_HEADER = 2
ROW_QUARTER_INDEX = 3

# P&L
ROW_REVENUE = 5
ROW_OPEX = 6
ROW_EBITDA = 7
ROW_DEPRECIATION = 8
ROW_EBIT = 9
ROW_INTEREST_EXPENSE = 10
ROW_EBT = 11
ROW_TAX = 12
ROW_LC_FEE = 13  # below tax: non-deductible, see the note in calc_tax.py
ROW_NET_INCOME = 14

# Cash Flow
ROW_CFO_NI = 16
ROW_CFO_ADDBACK_DEPR = 17
ROW_CFO = 18
ROW_CFI = 19
ROW_PRINCIPAL_REPAYMENT = 20
ROW_DSRA_FUNDING = 21
ROW_MRA_FUNDING = 22
ROW_DIVIDENDS_PAID = 23
ROW_EQUITY_INJECTION = 24
ROW_CFF = 25
ROW_NET_CHANGE_CASH = 26
ROW_OPENING_CASH = 27
ROW_CLOSING_CASH = 28

# Balance Sheet
ROW_BS_CASH = 30
ROW_BS_DSRA = 31
ROW_BS_MRA = 32
ROW_BS_PPE_NET = 33
ROW_BS_TOTAL_ASSETS = 34
ROW_BS_DEBT = 35
ROW_BS_PAID_IN_CAPITAL = 36
ROW_BS_RETAINED_EARNINGS = 37
ROW_BS_TOTAL_EQUITY = 38
ROW_BS_TOTAL_LIAB_EQUITY = 39

ROW_CHECK_HEADER = 42
ROW_CHECK_BS_BALANCES_COUNT = 43  # count of quarters where the BS does not balance
ROW_CHECK_CASH_TIES_BUFFER = 44   # closing cash must equal Calc_CFADS' buffer


def build_fs_quarterly(wb: Workbook, timeline: Timeline, inputs: ProjectInputs) -> Worksheet:
    ws = wb.create_sheet("FS_Quarterly")
    ws.sheet_properties.tabColor = TAB_COLOR_OUTPUT

    ws["A1"] = "FS_Quarterly — 3-Statements (reserves as restricted cash, 100% FCFE payout)"
    ws["A1"].font = Font(bold=True, size=12)

    _label(ws, ROW_DATE_HEADER, "Period End Date")
    _label(ws, ROW_QUARTER_INDEX, "Operating Quarter #")

    _label(ws, ROW_REVENUE, "Revenue ($)")
    _label(ws, ROW_OPEX, "Opex ($)")
    _label(ws, ROW_EBITDA, "EBITDA ($)")
    _label(ws, ROW_DEPRECIATION, "Depreciation ($) — base vintage + maintenance vintages")
    _label(ws, ROW_EBIT, "EBIT ($)")
    _label(ws, ROW_INTEREST_EXPENSE, "Interest Expense ($)")
    _label(ws, ROW_EBT, "EBT ($)")
    _label(ws, ROW_TAX, "Tax ($) — linked from Calc_Tax")
    _label(ws, ROW_LC_FEE, "DSRA LC Fee ($) — non-deductible; zero when the DSRA is cash-funded")
    _label(ws, ROW_NET_INCOME, "Net Income ($)")

    _label(ws, ROW_CFO_NI, "Net Income ($)")
    _label(ws, ROW_CFO_ADDBACK_DEPR, "Add back: Depreciation ($)")
    _label(ws, ROW_CFO, "Cash Flow from Operations ($)")
    _label(ws, ROW_CFI, "Cash Flow from Investing ($) — maintenance capex")
    _label(ws, ROW_PRINCIPAL_REPAYMENT, "Debt Principal Repayment ($)")
    _label(ws, ROW_DSRA_FUNDING, "DSRA Funding/(Release) ($) — the LC fee sits in the P&L above")
    _label(ws, ROW_MRA_FUNDING, "MRA Funding/(Release) ($)")
    _label(ws, ROW_DIVIDENDS_PAID, "Distributions to Equity ($)")
    _label(ws, ROW_EQUITY_INJECTION, "Equity Injections ($) — shortfalls the buffer could not cover")
    _label(ws, ROW_CFF, "Cash Flow from Financing ($)")
    _label(ws, ROW_NET_CHANGE_CASH, "Net Change in Cash ($)")
    _label(ws, ROW_OPENING_CASH, "Opening Cash ($)")
    _label(ws, ROW_CLOSING_CASH, "Closing Cash ($)")

    _label(ws, ROW_BS_CASH, "Cash ($) — unrestricted operating buffer")
    _label(ws, ROW_BS_DSRA, "DSRA Balance ($) — restricted cash")
    _label(ws, ROW_BS_MRA, "MRA Balance ($) — restricted cash")
    _label(ws, ROW_BS_PPE_NET, "PP&E, Net ($)")
    _label(ws, ROW_BS_TOTAL_ASSETS, "Total Assets ($)")
    _label(ws, ROW_BS_DEBT, "Debt ($)")
    _label(ws, ROW_BS_PAID_IN_CAPITAL, "Paid-in Capital ($)")
    _label(ws, ROW_BS_RETAINED_EARNINGS, "Retained Earnings ($)")
    _label(ws, ROW_BS_TOTAL_EQUITY, "Total Equity ($)")
    _label(ws, ROW_BS_TOTAL_LIAB_EQUITY, "Total Liabilities + Equity ($)")

    ws.cell(row=ROW_CHECK_HEADER, column=1, value="Checks").font = Font(bold=True)
    _label(ws, ROW_CHECK_BS_BALANCES_COUNT, "# of quarters where BS does not balance")
    _label(ws, ROW_CHECK_CASH_TIES_BUFFER, "Check: Closing cash ties to the Calc_CFADS buffer")

    n_quarters = len(timeline.operations_quarters)
    last_cons_col = col_letter(len(timeline.construction_months) - 1)

    for i, period in enumerate(timeline.operations_quarters):
        col = col_letter(i)
        prev_col = col_letter(i - 1) if i > 0 else None

        ws[f"{col}{ROW_DATE_HEADER}"] = period.end
        ws[f"{col}{ROW_DATE_HEADER}"].number_format = "mmm-yy"
        ws[f"{col}{ROW_QUARTER_INDEX}"] = i + 1

        _link(ws, col, ROW_REVENUE, f"Calc_Revenue_Opex!{col}{REVOPEX_ROW_REVENUE}")
        _link(ws, col, ROW_OPEX, f"Calc_Revenue_Opex!{col}{REVOPEX_ROW_OPEX}")
        _link(ws, col, ROW_EBITDA, f"Calc_Revenue_Opex!{col}{REVOPEX_ROW_EBITDA}")
        _link(ws, col, ROW_DEPRECIATION, f"Calc_Tax!{col}{tax.ROW_TOTAL_DEPRECIATION}")
        _formula(ws, col, ROW_EBIT, f"={col}{ROW_EBITDA}-{col}{ROW_DEPRECIATION}")
        _link(ws, col, ROW_INTEREST_EXPENSE, f"Calc_Financing_Ops!{col}{fin_ops.ROW_INTEREST}")
        _formula(ws, col, ROW_EBT, f"={col}{ROW_EBIT}-{col}{ROW_INTEREST_EXPENSE}")
        _link(ws, col, ROW_TAX, f"Calc_Tax!{col}{tax.ROW_TAX}")
        _link(ws, col, ROW_LC_FEE, f"Calc_Financing_Ops!{col}{fin_ops.ROW_DSRA_LC_FEE}")
        _formula(ws, col, ROW_NET_INCOME, f"={col}{ROW_EBT}-{col}{ROW_TAX}-{col}{ROW_LC_FEE}")

        _formula(ws, col, ROW_CFO_NI, f"={col}{ROW_NET_INCOME}")
        _formula(ws, col, ROW_CFO_ADDBACK_DEPR, f"={col}{ROW_DEPRECIATION}")
        _formula(ws, col, ROW_CFO, f"={col}{ROW_CFO_NI}+{col}{ROW_CFO_ADDBACK_DEPR}")
        _formula(ws, col, ROW_CFI, f"=-Calc_CFADS!{col}{cfads.ROW_MAINT_CAPEX}")
        _link(ws, col, ROW_PRINCIPAL_REPAYMENT, f"Calc_Financing_Ops!{col}{fin_ops.ROW_PRINCIPAL}")
        # Only the reserve's balance movement belongs here — the LC fee already reduced
        # cash through net income, so taking it again would double-count it.
        _link(ws, col, ROW_DSRA_FUNDING, f"Calc_Financing_Ops!{col}{fin_ops.ROW_DSRA_FUNDING}")
        _link(ws, col, ROW_MRA_FUNDING, f"Calc_CFADS!{col}{cfads.ROW_MRA_FUNDING}")
        _link(ws, col, ROW_DIVIDENDS_PAID, f"Calc_CFADS!{col}{cfads.ROW_DISTRIBUTION}")
        _link(ws, col, ROW_EQUITY_INJECTION, f"Calc_CFADS!{col}{cfads.ROW_EQUITY_INJECTION}")
        _formula(ws, col, ROW_CFF,
                 f"=-{col}{ROW_PRINCIPAL_REPAYMENT}-{col}{ROW_DSRA_FUNDING}"
                 f"-{col}{ROW_MRA_FUNDING}-{col}{ROW_DIVIDENDS_PAID}+{col}{ROW_EQUITY_INJECTION}")
        _formula(ws, col, ROW_NET_CHANGE_CASH,
                 f"={col}{ROW_CFO}+{col}{ROW_CFI}+{col}{ROW_CFF}")

        if i == 0:
            _formula(ws, col, ROW_OPENING_CASH, f"=Calc_Financing_Cons!{fin_cons.ABS_INITIAL_BUFFER}")
        else:
            _formula(ws, col, ROW_OPENING_CASH, f"={prev_col}{ROW_CLOSING_CASH}")
        _formula(ws, col, ROW_CLOSING_CASH, f"={col}{ROW_OPENING_CASH}+{col}{ROW_NET_CHANGE_CASH}")

        _formula(ws, col, ROW_BS_CASH, f"={col}{ROW_CLOSING_CASH}")
        # Reserves are cash the project holds but cannot distribute — assets in their own
        # right, so the LC-backed case correctly shows no asset and only a fee.
        _link(ws, col, ROW_BS_DSRA, f"Calc_Financing_Ops!{col}{fin_ops.ROW_DSRA_BALANCE}")
        _link(ws, col, ROW_BS_MRA, f"Calc_CFADS!{col}{cfads.ROW_MRA_BALANCE}")

        # CFI is already negative, so subtracting it capitalises the maintenance spend.
        if i == 0:
            _formula(ws, col, ROW_BS_PPE_NET,
                     f"=Calc_Capex!${last_cons_col}${capex.ROW_TOTAL_PROJECT_COST}"
                     f"-{col}{ROW_CFI}-{col}{ROW_DEPRECIATION}")
        else:
            _formula(ws, col, ROW_BS_PPE_NET,
                     f"={prev_col}{ROW_BS_PPE_NET}-{col}{ROW_CFI}-{col}{ROW_DEPRECIATION}")
        _formula(ws, col, ROW_BS_TOTAL_ASSETS,
                 f"={col}{ROW_BS_CASH}+{col}{ROW_BS_DSRA}+{col}{ROW_BS_MRA}+{col}{ROW_BS_PPE_NET}")

        _link(ws, col, ROW_BS_DEBT, f"Calc_Financing_Ops!{col}{fin_ops.ROW_CLOSING_BAL}")
        if i == 0:
            _formula(ws, col, ROW_BS_PAID_IN_CAPITAL,
                     f"=Calc_Capex!${last_cons_col}${capex.ROW_CUM_EQUITY_DRAW}"
                     f"+{col}{ROW_EQUITY_INJECTION}")
        else:
            _formula(ws, col, ROW_BS_PAID_IN_CAPITAL,
                     f"={prev_col}{ROW_BS_PAID_IN_CAPITAL}+{col}{ROW_EQUITY_INJECTION}")
        if i == 0:
            _formula(ws, col, ROW_BS_RETAINED_EARNINGS,
                     f"={col}{ROW_NET_INCOME}-{col}{ROW_DIVIDENDS_PAID}")
        else:
            _formula(ws, col, ROW_BS_RETAINED_EARNINGS,
                     f"={prev_col}{ROW_BS_RETAINED_EARNINGS}+{col}{ROW_NET_INCOME}-{col}{ROW_DIVIDENDS_PAID}")
        _formula(ws, col, ROW_BS_TOTAL_EQUITY,
                 f"={col}{ROW_BS_PAID_IN_CAPITAL}+{col}{ROW_BS_RETAINED_EARNINGS}")
        _formula(ws, col, ROW_BS_TOTAL_LIAB_EQUITY,
                 f"={col}{ROW_BS_DEBT}+{col}{ROW_BS_TOTAL_EQUITY}")

    first_col = col_letter(0)
    last_col = col_letter(n_quarters - 1)

    check_cell = ws[f"{last_col}{ROW_CHECK_BS_BALANCES_COUNT}"]
    check_cell.value = (
        f"=SUMPRODUCT(--(ROUND({first_col}{ROW_BS_TOTAL_ASSETS}:{last_col}{ROW_BS_TOTAL_ASSETS}"
        f"-{first_col}{ROW_BS_TOTAL_LIAB_EQUITY}:{last_col}{ROW_BS_TOTAL_LIAB_EQUITY},2)<>0))"
    )
    check_cell.font = Font(color=COLOR_FORMULA)

    # An independent tie-out: the statements build cash from CFO/CFI/CFF, while Calc_CFADS
    # builds the same balance from the waterfall. They are separate derivations, so a
    # mismatch means one of them is wrong.
    cash_tie = ws[f"{last_col}{ROW_CHECK_CASH_TIES_BUFFER}"]
    cash_tie.value = (
        f"=IF(SUMPRODUCT(--(ROUND({first_col}{ROW_CLOSING_CASH}:{last_col}{ROW_CLOSING_CASH}"
        f"-Calc_CFADS!{first_col}{cfads.ROW_BUFFER_CLOSING}:"
        f"Calc_CFADS!{last_col}{cfads.ROW_BUFFER_CLOSING},2)<>0))=0,1,0)"
    )
    cash_tie.font = Font(color=COLOR_FORMULA)

    _add_named_range(wb, "FSQ_LastCol", "FS_Quarterly", f"{last_col}1")
    _add_named_range(wb, "FSQ_CashTiesBufferCheck", "FS_Quarterly", f"{last_col}{ROW_CHECK_CASH_TIES_BUFFER}")
    _add_named_range(wb, "FSQ_BSBalancesFailCount", "FS_Quarterly", f"{last_col}{ROW_CHECK_BS_BALANCES_COUNT}")

    ws.freeze_panes = ws.cell(row=ROW_BS_TOTAL_LIAB_EQUITY + 1, column=FIRST_DATA_COL)
    ws.column_dimensions["A"].width = 52

    return ws


def _label(ws: Worksheet, row: int, label: str) -> None:
    ws.cell(row=row, column=1, value=label)


def _link(ws: Worksheet, col: str, row: int, formula: str) -> None:
    cell = ws[f"{col}{row}"]
    cell.value = f"={formula}"
    cell.font = Font(color=COLOR_LINK)
    cell.number_format = "#,##0"


def _formula(ws: Worksheet, col: str, row: int, formula: str) -> None:
    cell = ws[f"{col}{row}"]
    cell.value = formula
    cell.font = Font(color=COLOR_FORMULA)
    cell.number_format = "#,##0"


def _add_named_range(wb: Workbook, name: str, sheet: str, cell: str) -> None:
    import re

    from openpyxl.workbook.defined_name import DefinedName

    match = re.match(r"([A-Za-z]+)(\d+)", cell)
    col, row = match.group(1), match.group(2)
    wb.defined_names.add(DefinedName(name, attr_text=f"'{sheet}'!${col}${row}"))
