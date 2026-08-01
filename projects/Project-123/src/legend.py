from openpyxl.styles import Font, PatternFill, Border, Side
from openpyxl.worksheet.worksheet import Worksheet
from openpyxl.workbook import Workbook

from workbook_builder import (
    COLOR_INPUT, COLOR_FORMULA, COLOR_LINK,
    TAB_COLOR_INPUT, TAB_COLOR_CALC, TAB_COLOR_OUTPUT, TAB_COLOR_CHECK,
    FILL_SECTION_HEADER,
)

# A read-only "how to read this model" tab, adapted from the FAST/Corality-style legend
# convention (key to colour scheme, tab colour meaning, check semantics) but documenting
# Project 123's own conventions rather than importing a different model's scheme wholesale.

TAB_COLOR_INFO = "808080"

ROW_TITLE = 1
ROW_INTRO = 2

ROW_FONT_HEADER = 4
ROW_FONT_TABLE_HEADER = 5
ROW_FONT_FIRST = 6

ROW_TAB_HEADER = 11
ROW_TAB_TABLE_HEADER = 12
ROW_TAB_FIRST = 13

ROW_CHECK_HEADER = 18
ROW_CHECK_TABLE_HEADER = 19
ROW_CHECK_FIRST = 20

ROW_FLAG_HEADER = 25
ROW_FLAG_TABLE_HEADER = 26
ROW_FLAG_FIRST = 27

ROW_BUTTON_HEADER = 31
ROW_BUTTON_TABLE_HEADER = 32
ROW_BUTTON_FIRST = 33

ROW_NAV_HEADER = 39
ROW_NAV_TEXT = 40


def build_legend(wb: Workbook) -> Worksheet:
    ws = wb.create_sheet("Legend", 0)  # sits before Dashboard — the first thing a new reader sees
    ws.sheet_properties.tabColor = TAB_COLOR_INFO

    ws.cell(row=ROW_TITLE, column=1, value="Project 123 — How to Read This Model").font = Font(bold=True, size=14)
    ws.cell(row=ROW_INTRO, column=1,
            value="Every sheet in this workbook follows the conventions below. Read this once; "
                  "nothing here is scenario- or deal-specific.").font = Font(italic=True, size=9)

    _section(ws, ROW_FONT_HEADER, "Font Colour — What a Cell's Colour Means")
    _table_header(ws, ROW_FONT_TABLE_HEADER, ("Colour", "Example", "Meaning"))
    font_rows = (
        (COLOR_INPUT, "Hardcoded input / assumption — the only cells meant to be edited directly."),
        (COLOR_FORMULA, "Formula that only uses cells on its own sheet."),
        (COLOR_LINK, "Formula that links to a value on a different sheet."),
    )
    for i, (color, meaning) in enumerate(font_rows):
        r = ROW_FONT_FIRST + i
        example = ws.cell(row=r, column=1, value=123456)
        example.font = Font(color=color, bold=True)
        example.number_format = "#,##0"
        ws.cell(row=r, column=2, value=meaning)
    ws.cell(row=ROW_FONT_FIRST + len(font_rows), column=1,
            value="Cross-workbook links (a different file) are deliberately avoided everywhere in this model.").font = Font(italic=True, size=9)

    _section(ws, ROW_TAB_HEADER, "Sheet Tab Colour — What Kind of Sheet You're On")
    _table_header(ws, ROW_TAB_TABLE_HEADER, ("Colour", "", "Meaning"))
    tab_rows = (
        (TAB_COLOR_INPUT, "Input — Cover and Assumptions sheets. Every editable driver lives here."),
        (TAB_COLOR_CALC, "Calculation — the engine. Formulas only; nothing here is meant to be typed into."),
        (TAB_COLOR_OUTPUT, "Output — 3-statements, Dashboard, valuation, batch/stress results."),
        (TAB_COLOR_CHECK, "Check — Check_Control, the master audit aggregator."),
        (TAB_COLOR_INFO, "Info — this Legend tab. Read-only, no model logic."),
    )
    for i, (color, meaning) in enumerate(tab_rows):
        r = ROW_TAB_TABLE_HEADER + 1 + i
        swatch = ws.cell(row=r, column=1, value="   ")
        swatch.fill = PatternFill(start_color=f"FF{color}" if len(color) == 6 else color,
                                  end_color=f"FF{color}" if len(color) == 6 else color, fill_type="solid")
        ws.cell(row=r, column=3, value=meaning)

    _section(ws, ROW_CHECK_HEADER, "Check Semantics — Check_Control and Every Local Checks Block")
    _table_header(ws, ROW_CHECK_TABLE_HEADER, ("Status", "Value", "Meaning"))
    check_rows = (
        ("OK", "1 (flag) or 0 (count)", "Passes. No action needed."),
        ("FAIL", "0 (flag) or >0 (count)", "A hard check failed — investigate before trusting outputs."),
        ("REVIEW", "informational count", "Not an error by itself (e.g. a gearing cap binding as designed) — "
                                          "read the description, don't ignore the row."),
    )
    for i, (status, value, meaning) in enumerate(check_rows):
        r = ROW_CHECK_FIRST + i
        ws.cell(row=r, column=1, value=status).font = Font(bold=True)
        ws.cell(row=r, column=2, value=value)
        ws.cell(row=r, column=3, value=meaning)

    _section(ws, ROW_FLAG_HEADER, "Flags Block — Top of Cover, Dashboard, FS_Quarterly, FS_Annual")
    _table_header(ws, ROW_FLAG_TABLE_HEADER, ("Flag", "", "Meaning"))
    flag_rows = (
        ("Active Scenario", "Which of the 10 scenarios every Calc/FS sheet is currently reading."),
        ("Model Status", "Check_Control's master flag — MODEL OK or ERRORS FOUND, live everywhere."),
        ("Solve Freshness", "Whether the two circularity loops have been re-solved since assumptions "
                            "last changed. Re-run Solve All if this reads stale."),
    )
    for i, (flag, meaning) in enumerate(flag_rows):
        r = ROW_FLAG_FIRST + i
        ws.cell(row=r, column=1, value=flag).font = Font(bold=True)
        ws.cell(row=r, column=3, value=meaning)
    ws.cell(row=ROW_FLAG_FIRST + len(flag_rows) + 1, column=1,
            value="These three rows are frozen on each of those sheets, so they stay visible no matter how "
                  "far you scroll.").font = Font(italic=True, size=9)

    _section(ws, ROW_BUTTON_HEADER, "Control Panel Button Tiers — Cover, Top-Right")
    _table_header(ws, ROW_BUTTON_TABLE_HEADER, ("Tier", "", "When to use it"))
    button_rows = (
        ("[Everyday]", "Solve All, Goal Seek EIRR/PIRR — what you run day to day on the active scenario."),
        ("[Batch]", "Run All 10 Scenarios, Run Stress Test — multi-scenario or multi-shock runs."),
        ("[Debug]", "The individual Loop 1 / Loop 2 solves — for isolating a convergence issue."),
        ("[Recovery]", "Reset Staged Values — when a solve won't converge from where it's sitting."),
    )
    for i, (tier, meaning) in enumerate(button_rows):
        r = ROW_BUTTON_FIRST + i
        ws.cell(row=r, column=1, value=tier).font = Font(bold=True)
        ws.cell(row=r, column=3, value=meaning)

    ws.cell(row=ROW_NAV_HEADER, column=1, value="Where to Start").font = Font(bold=True, size=11, underline="single")
    ws.cell(row=ROW_NAV_TEXT, column=1,
            value="Dashboard (sheet 1) for the headline read. Cover for solve settings and the control "
                  "panel. Assumptions_Constant/Periodic sheets for inputs. Everything else is generated "
                  "and should not be edited by hand.").font = Font(italic=True, size=9)

    ws.column_dimensions["A"].width = 24
    ws.column_dimensions["B"].width = 16
    ws.column_dimensions["C"].width = 70

    return ws


def _section(ws: Worksheet, row: int, title: str) -> None:
    cell = ws.cell(row=row, column=1, value=title)
    cell.font = Font(bold=True, size=12)
    for col in range(1, 4):
        ws.cell(row=row, column=col).fill = FILL_SECTION_HEADER


def _table_header(ws: Worksheet, row: int, headers: tuple) -> None:
    for i, header in enumerate(headers):
        if header:
            ws.cell(row=row, column=1 + i, value=header).font = Font(bold=True, underline="single")
