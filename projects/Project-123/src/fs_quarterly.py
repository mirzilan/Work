from openpyxl.styles import Font
from openpyxl.worksheet.worksheet import Worksheet
from openpyxl.workbook import Workbook

from inputs import ProjectInputs
from timeline import Timeline
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
ROW_NET_INCOME = 13

# Cash Flow
ROW_CFO_NI = 15
ROW_CFO_ADDBACK_DEPR = 16
ROW_CFO = 17
ROW_CFI = 18
ROW_PRINCIPAL_REPAYMENT = 19
ROW_DIVIDENDS_PAID = 20
ROW_CFF = 21
ROW_NET_CHANGE_CASH = 22
ROW_OPENING_CASH = 23
ROW_CLOSING_CASH = 24

# Balance Sheet
ROW_BS_CASH = 26
ROW_BS_PPE_NET = 27
ROW_BS_TOTAL_ASSETS = 28
ROW_BS_DEBT = 29
ROW_BS_PAID_IN_CAPITAL = 30
ROW_BS_RETAINED_EARNINGS = 31
ROW_BS_TOTAL_EQUITY = 32
ROW_BS_TOTAL_LIAB_EQUITY = 33

ROW_CHECK_HEADER = 37
ROW_CHECK_BS_BALANCES_COUNT = 38  # informational: # of quarters where BS does not balance


def build_fs_quarterly(wb: Workbook, timeline: Timeline, inputs: ProjectInputs) -> Worksheet:
    ws = wb.create_sheet("FS_Quarterly")
    ws.sheet_properties.tabColor = TAB_COLOR_OUTPUT

    ws["A1"] = "FS_Quarterly — 3-Statements (Stage 1a: single vintage, no reserves, 100% FCFE payout)"
    ws["A1"].font = Font(bold=True, size=12)

    _label(ws, ROW_DATE_HEADER, "Period End Date")
    _label(ws, ROW_QUARTER_INDEX, "Operating Quarter #")

    _label(ws, ROW_REVENUE, "Revenue ($)")
    _label(ws, ROW_OPEX, "Opex ($)")
    _label(ws, ROW_EBITDA, "EBITDA ($)")
    _label(ws, ROW_DEPRECIATION, "Depreciation ($)")
    _label(ws, ROW_EBIT, "EBIT ($)")
    _label(ws, ROW_INTEREST_EXPENSE, "Interest Expense ($)")
    _label(ws, ROW_EBT, "EBT ($) — accounting, post-interest")
    _label(ws, ROW_TAX, "Tax ($) — Stage 1a: computed pre-interest-shield (see Calc_Tax)")
    _label(ws, ROW_NET_INCOME, "Net Income ($)")

    _label(ws, ROW_CFO_NI, "Net Income ($)")
    _label(ws, ROW_CFO_ADDBACK_DEPR, "Add back: Depreciation ($)")
    _label(ws, ROW_CFO, "Cash Flow from Operations ($)")
    _label(ws, ROW_CFI, "Cash Flow from Investing ($) — 0, no ops capex in Stage 1a")
    _label(ws, ROW_PRINCIPAL_REPAYMENT, "Debt Principal Repayment ($)")
    _label(ws, ROW_DIVIDENDS_PAID, "Dividends Paid ($) — 100% FCFE payout, Stage 1a")
    _label(ws, ROW_CFF, "Cash Flow from Financing ($)")
    _label(ws, ROW_NET_CHANGE_CASH, "Net Change in Cash ($)")
    _label(ws, ROW_OPENING_CASH, "Opening Cash ($)")
    _label(ws, ROW_CLOSING_CASH, "Closing Cash ($)")

    _label(ws, ROW_BS_CASH, "Cash ($)")
    _label(ws, ROW_BS_PPE_NET, "PP&E, Net ($)")
    _label(ws, ROW_BS_TOTAL_ASSETS, "Total Assets ($)")
    _label(ws, ROW_BS_DEBT, "Debt ($)")
    _label(ws, ROW_BS_PAID_IN_CAPITAL, "Paid-in Capital ($)")
    _label(ws, ROW_BS_RETAINED_EARNINGS, "Retained Earnings ($)")
    _label(ws, ROW_BS_TOTAL_EQUITY, "Total Equity ($)")
    _label(ws, ROW_BS_TOTAL_LIAB_EQUITY, "Total Liabilities + Equity ($)")

    ws.cell(row=ROW_CHECK_HEADER, column=1, value="Checks").font = Font(bold=True)
    _label(ws, ROW_CHECK_BS_BALANCES_COUNT, "Informational: # of quarters where BS does not balance")

    n_quarters = len(timeline.operations_quarters)
    last_cons_col = col_letter(len(timeline.construction_months) - 1)

    for i, period in enumerate(timeline.operations_quarters):
        col = col_letter(i)
        prev_col = col_letter(i - 1) if i > 0 else None

        ws[f"{col}{ROW_DATE_HEADER}"] = period.end
        ws[f"{col}{ROW_DATE_HEADER}"].number_format = "mmm-yy"
        ws[f"{col}{ROW_QUARTER_INDEX}"] = i + 1

        _link(ws, col, ROW_REVENUE, f"Calc_Revenue_Opex!{col}5")
        _link(ws, col, ROW_OPEX, f"Calc_Revenue_Opex!{col}6")
        _link(ws, col, ROW_EBITDA, f"Calc_Revenue_Opex!{col}7")
        _link(ws, col, ROW_DEPRECIATION, f"Calc_Tax!{col}6")
        _formula(ws, col, ROW_EBIT, f"={col}{ROW_EBITDA}-{col}{ROW_DEPRECIATION}")
        _link(ws, col, ROW_INTEREST_EXPENSE, f"Calc_Financing_Ops!{col}6")
        _formula(ws, col, ROW_EBT, f"={col}{ROW_EBIT}-{col}{ROW_INTEREST_EXPENSE}")
        _link(ws, col, ROW_TAX, f"Calc_Tax!{col}8")
        _formula(ws, col, ROW_NET_INCOME, f"={col}{ROW_EBT}-{col}{ROW_TAX}")

        _formula(ws, col, ROW_CFO_NI, f"={col}{ROW_NET_INCOME}")
        _formula(ws, col, ROW_CFO_ADDBACK_DEPR, f"={col}{ROW_DEPRECIATION}")
        _formula(ws, col, ROW_CFO, f"={col}{ROW_CFO_NI}+{col}{ROW_CFO_ADDBACK_DEPR}")
        _formula(ws, col, ROW_CFI, "=0")
        _link(ws, col, ROW_PRINCIPAL_REPAYMENT, f"Calc_Financing_Ops!{col}7")
        _link(ws, col, ROW_DIVIDENDS_PAID, f"Calc_CFADS!{col}7")
        _formula(ws, col, ROW_CFF, f"=-{col}{ROW_PRINCIPAL_REPAYMENT}-{col}{ROW_DIVIDENDS_PAID}")
        _formula(ws, col, ROW_NET_CHANGE_CASH, f"={col}{ROW_CFO}+{col}{ROW_CFI}+{col}{ROW_CFF}")

        if i == 0:
            _formula(ws, col, ROW_OPENING_CASH, "=0")
        else:
            _formula(ws, col, ROW_OPENING_CASH, f"={prev_col}{ROW_CLOSING_CASH}")
        _formula(ws, col, ROW_CLOSING_CASH, f"={col}{ROW_OPENING_CASH}+{col}{ROW_NET_CHANGE_CASH}")

        _formula(ws, col, ROW_BS_CASH, f"={col}{ROW_CLOSING_CASH}")
        if i == 0:
            _formula(ws, col, ROW_BS_PPE_NET, f"=Calc_Capex!${last_cons_col}$17-{col}{ROW_DEPRECIATION}")
        else:
            _formula(ws, col, ROW_BS_PPE_NET, f"={prev_col}{ROW_BS_PPE_NET}-{col}{ROW_DEPRECIATION}")
        _formula(ws, col, ROW_BS_TOTAL_ASSETS, f"={col}{ROW_BS_CASH}+{col}{ROW_BS_PPE_NET}")

        _link(ws, col, ROW_BS_DEBT, f"Calc_Financing_Ops!{col}9")
        _link(ws, col, ROW_BS_PAID_IN_CAPITAL, f"Calc_Capex!${last_cons_col}$13")
        if i == 0:
            _formula(ws, col, ROW_BS_RETAINED_EARNINGS, f"={col}{ROW_NET_INCOME}-{col}{ROW_DIVIDENDS_PAID}")
        else:
            _formula(
                ws,
                col,
                ROW_BS_RETAINED_EARNINGS,
                f"={prev_col}{ROW_BS_RETAINED_EARNINGS}+{col}{ROW_NET_INCOME}-{col}{ROW_DIVIDENDS_PAID}",
            )
        _formula(ws, col, ROW_BS_TOTAL_EQUITY, f"={col}{ROW_BS_PAID_IN_CAPITAL}+{col}{ROW_BS_RETAINED_EARNINGS}")
        _formula(ws, col, ROW_BS_TOTAL_LIAB_EQUITY, f"={col}{ROW_BS_DEBT}+{col}{ROW_BS_TOTAL_EQUITY}")

    first_col = col_letter(0)
    last_col = col_letter(n_quarters - 1)

    check_cell = ws[f"{last_col}{ROW_CHECK_BS_BALANCES_COUNT}"]
    check_cell.value = (
        f"=SUMPRODUCT(--(ROUND({first_col}{ROW_BS_TOTAL_ASSETS}:{last_col}{ROW_BS_TOTAL_ASSETS}"
        f"-{first_col}{ROW_BS_TOTAL_LIAB_EQUITY}:{last_col}{ROW_BS_TOTAL_LIAB_EQUITY},2)<>0))"
    )
    check_cell.font = Font(color=COLOR_FORMULA)

    _add_named_range(wb, "FSQ_LastCol", "FS_Quarterly", f"{last_col}1")
    _add_named_range(wb, "FSQ_BSBalancesFailCount", "FS_Quarterly", f"{last_col}{ROW_CHECK_BS_BALANCES_COUNT}")

    ws.freeze_panes = ws.cell(row=ROW_BS_TOTAL_LIAB_EQUITY + 1, column=FIRST_DATA_COL)

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
