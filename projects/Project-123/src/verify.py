"""Recalculate the built workbook in LibreOffice and drive the two convergence loops the
VBA would run, then report every Check_Control row plus the headline outputs.

This evaluates the workbook's own formulas rather than re-deriving what they should say —
a harness that reimplements the intent will agree with itself and miss the bug (which is
how the pari-passu gearing error survived its first review)."""

import subprocess
import sys
import time
from pathlib import Path

import uno
from com.sun.star.beans import PropertyValue

import cover_refs as refs

OUTPUT = Path(__file__).resolve().parent.parent / "output" / "project123_stage1c.xlsx"
SOCKET = "socket,host=localhost,port=2002;urp;StarOffice.ComponentContext"

STAGED_IDC = ("Calc_Financing_Cons", "B6")
CALCULATED_IDC = ("Calc_Financing_Cons", "B10")
IDC_GAP = ("Calc_Financing_Cons", "B11")

STAGED_DEBT = ("Calc_Financing_Ops", "B7")
DEBT_CAPACITY = ("Calc_Financing_Ops", "B8")
DEBT_GAP = ("Calc_Financing_Ops", "B10")


def _start_soffice():
    proc = subprocess.Popen(
        ["soffice", "--headless", "--norestore", "--invisible",
         f"--accept={SOCKET}", "--nologo", "--nofirststartwizard"],
        stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
    )
    ctx = None
    resolver = uno.getComponentContext().ServiceManager.createInstanceWithContext(
        "com.sun.star.bridge.UnoUrlResolver", uno.getComponentContext())
    for _ in range(60):
        try:
            ctx = resolver.resolve(f"uno:{SOCKET}")
            break
        except Exception:
            time.sleep(1)
    if ctx is None:
        raise RuntimeError("could not connect to soffice")
    return proc, ctx


def _load(ctx, path):
    desktop = ctx.ServiceManager.createInstanceWithContext(
        "com.sun.star.frame.Desktop", ctx)
    hidden = PropertyValue(); hidden.Name = "Hidden"; hidden.Value = True
    # Explicit filter: type detection on an openpyxl-written file is unreliable headless.
    flt = PropertyValue(); flt.Name = "FilterName"; flt.Value = "Calc MS Excel 2007 XML"
    url = uno.systemPathToFileUrl(str(path))
    return desktop.loadComponentFromURL(url, "_blank", 0, (hidden, flt))


def _cell(doc, sheet, ref):
    return doc.Sheets.getByName(sheet).getCellRangeByName(ref)


def get(doc, addr):
    return _cell(doc, addr[0], addr[1]).getValue()


def get_str(doc, sheet, ref):
    return _cell(doc, sheet, ref).getString()


def set_value(doc, addr, value):
    _cell(doc, addr[0], addr[1]).setValue(value)


def solve(doc, passes=40, tol=1.0):
    """Mirrors mod_Loop2's SolveAllCurrentScenario: alternate the two staged cells until
    both gaps close. Each loop is a plain fixed point — staged := calculated."""
    for i in range(passes):
        for _ in range(30):
            doc.calculateAll()
            gap = get(doc, IDC_GAP)
            if abs(gap) <= tol:
                break
            set_value(doc, STAGED_IDC, get(doc, CALCULATED_IDC))

        doc.calculateAll()
        debt_gap = get(doc, DEBT_GAP)
        if abs(debt_gap) <= 100.0 and abs(get(doc, IDC_GAP)) <= tol:
            return i + 1
        set_value(doc, STAGED_DEBT, get(doc, DEBT_CAPACITY))
    doc.calculateAll()
    return passes


def _headline_refs():
    """Built from the builders' own row constants. These were literals once and silently
    drifted onto the wrong rows the moment a sheet gained a line."""
    import calc_cfads as cfads
    import calc_financing_cons as fin_cons
    import calc_financing_ops as fin_ops
    import cover
    import fs_quarterly as fsq
    from workbook_builder import col_letter

    ops = col_letter(79)
    q1 = col_letter(0)
    return [
        ("Total Project Cost", "Cover", f"B{cover.ROW_LIVE_TPC}", ",.0f"),
        ("Debt Facility", "Calc_Financing_Cons", fin_cons.CELL_DEBT_FACILITY, ",.0f"),
        ("Implied Gearing", "Cover", f"B{cover.ROW_LIVE_GEARING}", ".2%"),
        ("Min DSCR over tenor", "Cover", f"B{cover.ROW_LIVE_MIN_DSCR}", ".4f"),
        ("Min LLCR", "Cover", f"B{cover.ROW_LIVE_MIN_LLCR}", ".3f"),
        ("PLCR at Q1", "Calc_Financing_Ops", f"{q1}{fin_ops.ROW_PLCR}", ".3f"),
        ("DSRA balance Q1", "Calc_Financing_Ops", f"{q1}{fin_ops.ROW_DSRA_BALANCE}", ",.0f"),
        ("MRA balance Q1", "Calc_CFADS", f"{q1}{cfads.ROW_MRA_BALANCE}", ",.0f"),
        ("Cash buffer, final Q", "Calc_CFADS", f"{ops}{cfads.ROW_BUFFER_CLOSING}", ",.0f"),
        ("Equity injections, total", "Calc_CFADS", f"{ops}{cfads.ROW_CHECK_EQUITY_INJECTIONS}", ",.0f"),
        ("PIRR", "FS_Annual", "A19", ".4%"),
        ("EIRR", "FS_Annual", "A21", ".4%"),
        ("Closing cash, final Q", "FS_Quarterly", f"{ops}{fsq.ROW_CLOSING_CASH}", ",.2f"),
        ("Closing debt, final Q", "FS_Quarterly", f"{ops}{fsq.ROW_BS_DEBT}", ",.2f"),
        ("Closing PP&E, final Q", "FS_Quarterly", f"{ops}{fsq.ROW_BS_PPE_NET}", ",.0f"),
    ]


def report(doc, label):
    print(f"\n{'=' * 78}\n{label}\n{'=' * 78}")

    passes = solve(doc)
    print(f"converged in {passes} passes  "
          f"(IDC gap {get(doc, IDC_GAP):,.4f} | debt gap {get(doc, DEBT_GAP):,.4f})")

    print(f"\nMASTER STATUS: {get_str(doc, 'Check_Control', 'B3')}")

    failures = []
    row = 6
    while True:
        sheet = get_str(doc, "Check_Control", f"A{row}")
        if not sheet:
            break
        desc = get_str(doc, "Check_Control", f"B{row}")
        val = get_str(doc, "Check_Control", f"D{row}")
        status = get_str(doc, "Check_Control", f"E{row}")
        marker = "    " if status == "OK" else " >> "
        print(f"{marker}{status:6} {sheet:28} {desc[:46]:46} {val}")
        if status != "OK":
            failures.append((sheet, desc, val, status))
        row += 1

    print("\nHeadline outputs")
    for name, sheet, ref, fmt in _headline_refs():
        try:
            print(f"  {name:26} {format(get(doc, (sheet, ref)), fmt)}")
        except Exception:
            print(f"  {name:26} (unreadable)")

    return failures


def main():
    proc, ctx = _start_soffice()
    try:
        doc = _load(ctx, OUTPUT)
        all_failures = []

        for scenario, name in ((1, "Base"), (2, "Upside"), (3, "Downside")):
            set_value(doc, ("Cover", refs.CELL_ACTIVE_SCENARIO), scenario)
            all_failures += [(name, *f) for f in report(doc, f"SCENARIO {scenario} — {name} (DSRA cash-funded)")]

        set_value(doc, ("Cover", refs.CELL_ACTIVE_SCENARIO), 1)
        _cell(doc, "Cover", refs.CELL_DSRA_METHOD).setString(refs.DSRA_METHODS[1])
        all_failures += [("Base/LC", *f) for f in report(doc, "SCENARIO 1 — Base, DSRA LC-BACKED")]
        _cell(doc, "Cover", refs.CELL_DSRA_METHOD).setString(refs.DSRA_METHOD_CASH)

        print(f"\n{'=' * 78}")
        if all_failures:
            print(f"{len(all_failures)} non-OK rows:")
            for f in all_failures:
                print("  ", f)
        else:
            print("All checks OK across every scenario and both DSRA funding methods.")
        doc.close(False)
        return 1 if any(f[4] == "FAIL" for f in all_failures) else 0
    finally:
        try:
            ctx.ServiceManager.createInstanceWithContext(
                "com.sun.star.frame.Desktop", ctx).terminate()
        except Exception:
            pass
        proc.wait(timeout=30)


if __name__ == "__main__":
    sys.exit(main())
