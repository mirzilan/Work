from openpyxl.styles import Font
from openpyxl.worksheet.worksheet import Worksheet
from openpyxl.workbook import Workbook

import assumptions_constant as const
import calc_revenue_opex as rev_opex
from inputs import ProjectInputs
from timeline import Timeline
from workbook_builder import (
    FIRST_DATA_COL,
    COLOR_FORMULA,
    COLOR_LINK,
    TAB_COLOR_CALC,
    col_letter,
)

# Trade receivables/payables, DSO/DPO-style. Receivable/Payable Days are scenario drivers
# on Assumptions_Constant; everything else here is derived. A fixed 91.25-day quarter
# (365/4) is used for the days<->balance conversion — close enough for a DSO/DPO working
# capital estimate, and it keeps the formula simple; a calendar-exact day count would need
# each quarter's actual span, which the timeline doesn't currently expose to formulas.
DAYS_PER_QUARTER = 91.25

ROW_DATE_HEADER = 2
ROW_QUARTER_INDEX = 3

ROW_AR_HEADER = 5
ROW_AR_BALANCE_BF = 6
ROW_AR_ADDITION = 7   # = Revenue for the quarter
ROW_AR_COLLECTIONS = 8  # plug: Balance b/f + Addition - Balance c/f
ROW_AR_BALANCE_CF = 9   # = Revenue x (Receivable Days / 91.25), DSO-style

ROW_AP_HEADER = 11
ROW_AP_BALANCE_BF = 12
ROW_AP_ADDITION = 13   # = Opex for the quarter
ROW_AP_PAYMENTS = 14   # plug: Balance b/f + Addition - Balance c/f
ROW_AP_BALANCE_CF = 15  # = Opex x (Payable Days / 91.25), DPO-style

ROW_CHANGE_IN_NWC = 17  # cash-flow-statement sign: -(ΔAR) + ΔAP
ROW_CUMULATIVE_CHANGE_IN_NWC = 18  # running total — 0 at the final quarter, see below

ROW_CHECK_HEADER = 20
ROW_CHECK_AR_NONNEG = 21
ROW_CHECK_AP_NONNEG = 22
ROW_CHECK_NWC_UNWINDS = 23


def build_calc_working_capital(wb: Workbook, timeline: Timeline, inputs: ProjectInputs) -> Worksheet:
    ws = wb.create_sheet("Calc_Working_Capital")
    ws.sheet_properties.tabColor = TAB_COLOR_CALC

    ws["A1"] = "Calc_Working_Capital — Trade Receivables/Payables (DSO/DPO), Quarterly"
    ws["A1"].font = Font(bold=True, size=12)
    ws["A2"] = (
        "Balances are driven off Receivable/Payable Days (Assumptions_Constant, scenario-varying); "
        "Collections/Payments are the plug that makes the balance roll forward hit that target."
    )
    ws["A2"].font = Font(italic=True, size=9)

    n_quarters = len(timeline.operations_quarters)

    _label(ws, ROW_DATE_HEADER, "Period End Date")
    _label(ws, ROW_QUARTER_INDEX, "Operating Quarter #")

    ws.cell(row=ROW_AR_HEADER, column=1, value="Trade Receivables").font = Font(bold=True)
    _label(ws, ROW_AR_BALANCE_BF, "Balance, Opening ($)")
    _label(ws, ROW_AR_ADDITION, "Addition — Revenue ($)")
    _label(ws, ROW_AR_COLLECTIONS, "Collections from Customers ($) — plug")
    _label(ws, ROW_AR_BALANCE_CF, "Balance, Closing ($) = Revenue x Receivable Days / 91.25")

    ws.cell(row=ROW_AP_HEADER, column=1, value="Trade Payables").font = Font(bold=True)
    _label(ws, ROW_AP_BALANCE_BF, "Balance, Opening ($)")
    _label(ws, ROW_AP_ADDITION, "Addition — Opex ($)")
    _label(ws, ROW_AP_PAYMENTS, "Payments to Suppliers ($) — plug")
    _label(ws, ROW_AP_BALANCE_CF, "Balance, Closing ($) = Opex x Payable Days / 91.25")

    _label(ws, ROW_CHANGE_IN_NWC,
           "Change in Net Working Capital ($) — cash flow sign: -(ΔAR) + ΔAP")
    _label(ws, ROW_CUMULATIVE_CHANGE_IN_NWC,
           "Cumulative Change in NWC ($) — nets to 0 by the final quarter (AR/AP wind down)")

    ws.cell(row=ROW_CHECK_HEADER, column=1, value="Checks").font = Font(bold=True)
    _label(ws, ROW_CHECK_AR_NONNEG, "Check: Receivables balance never negative")
    _label(ws, ROW_CHECK_AP_NONNEG, "Check: Payables balance never negative")
    _label(ws, ROW_CHECK_NWC_UNWINDS, "Check: Cumulative Change in NWC = 0 by the final quarter")

    for i, period in enumerate(timeline.operations_quarters):
        col = col_letter(i)
        prev = col_letter(i - 1) if i > 0 else None

        ws[f"{col}{ROW_DATE_HEADER}"] = period.end
        ws[f"{col}{ROW_DATE_HEADER}"].number_format = "mmm-yy"
        ws[f"{col}{ROW_QUARTER_INDEX}"] = i + 1

        is_last_quarter = i == n_quarters - 1

        _link(ws, col, ROW_AR_ADDITION, f"Calc_Revenue_Opex!{col}{rev_opex.ROW_REVENUE}")
        if is_last_quarter:
            # No going concern past the horizon — same convention Calc_CFADS already
            # uses for the cash buffer target. Without this, AR/AP sit on the books
            # forever uncollected/unpaid, and the model quietly overstates PP&E-adjacent
            # cash generation because working capital would never fully convert to cash.
            _formula(ws, col, ROW_AR_BALANCE_CF, "=0")
        else:
            _formula(ws, col, ROW_AR_BALANCE_CF,
                     f"={col}{ROW_AR_ADDITION}*Assumptions_Constant!$B${const.ROW_RECEIVABLE_DAYS}/{DAYS_PER_QUARTER}")
        if i == 0:
            _formula(ws, col, ROW_AR_BALANCE_BF, "=0")
        else:
            _formula(ws, col, ROW_AR_BALANCE_BF, f"={prev}{ROW_AR_BALANCE_CF}")
        _formula(ws, col, ROW_AR_COLLECTIONS,
                 f"={col}{ROW_AR_BALANCE_BF}+{col}{ROW_AR_ADDITION}-{col}{ROW_AR_BALANCE_CF}")

        _link(ws, col, ROW_AP_ADDITION, f"Calc_Revenue_Opex!{col}{rev_opex.ROW_OPEX}")
        if is_last_quarter:
            _formula(ws, col, ROW_AP_BALANCE_CF, "=0")
        else:
            _formula(ws, col, ROW_AP_BALANCE_CF,
                     f"={col}{ROW_AP_ADDITION}*Assumptions_Constant!$B${const.ROW_PAYABLE_DAYS}/{DAYS_PER_QUARTER}")
        if i == 0:
            _formula(ws, col, ROW_AP_BALANCE_BF, "=0")
        else:
            _formula(ws, col, ROW_AP_BALANCE_BF, f"={prev}{ROW_AP_BALANCE_CF}")
        _formula(ws, col, ROW_AP_PAYMENTS,
                 f"={col}{ROW_AP_BALANCE_BF}+{col}{ROW_AP_ADDITION}-{col}{ROW_AP_BALANCE_CF}")

        _formula(ws, col, ROW_CHANGE_IN_NWC,
                 f"=-({col}{ROW_AR_BALANCE_CF}-{col}{ROW_AR_BALANCE_BF})"
                 f"+({col}{ROW_AP_BALANCE_CF}-{col}{ROW_AP_BALANCE_BF})")
        if i == 0:
            _formula(ws, col, ROW_CUMULATIVE_CHANGE_IN_NWC, f"={col}{ROW_CHANGE_IN_NWC}")
        else:
            _formula(ws, col, ROW_CUMULATIVE_CHANGE_IN_NWC,
                     f"={prev}{ROW_CUMULATIVE_CHANGE_IN_NWC}+{col}{ROW_CHANGE_IN_NWC}")

    first_col = col_letter(0)
    last_col = col_letter(n_quarters - 1)

    ar_check = ws[f"{last_col}{ROW_CHECK_AR_NONNEG}"]
    ar_check.value = f"=IF(MIN({first_col}{ROW_AR_BALANCE_CF}:{last_col}{ROW_AR_BALANCE_CF})>=0,1,0)"
    ar_check.font = Font(color=COLOR_FORMULA)

    ap_check = ws[f"{last_col}{ROW_CHECK_AP_NONNEG}"]
    ap_check.value = f"=IF(MIN({first_col}{ROW_AP_BALANCE_CF}:{last_col}{ROW_AP_BALANCE_CF})>=0,1,0)"
    ap_check.font = Font(color=COLOR_FORMULA)

    unwind_check = ws[f"{last_col}{ROW_CHECK_NWC_UNWINDS}"]
    unwind_check.value = f"=IF(ROUND({last_col}{ROW_CUMULATIVE_CHANGE_IN_NWC},2)=0,1,0)"
    unwind_check.font = Font(color=COLOR_FORMULA)

    _add_named_range(wb, "WC_LastCol", "Calc_Working_Capital", f"{last_col}1")
    _add_named_range(wb, "WC_ARNonNegCheck", "Calc_Working_Capital", f"{last_col}{ROW_CHECK_AR_NONNEG}")
    _add_named_range(wb, "WC_APNonNegCheck", "Calc_Working_Capital", f"{last_col}{ROW_CHECK_AP_NONNEG}")
    _add_named_range(wb, "WC_UnwindsCheck", "Calc_Working_Capital", f"{last_col}{ROW_CHECK_NWC_UNWINDS}")

    ws.freeze_panes = ws.cell(row=ROW_AR_HEADER, column=FIRST_DATA_COL)
    ws.column_dimensions["A"].width = 50

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
