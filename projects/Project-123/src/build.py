"""Build the workbook.

    python build.py                          -> plain .xlsx, no macros
    python build.py --template <t.xlsm>      -> .xlsm carrying the template's VBA project

The template route reopens a workbook you have already pasted the macros into, drops
every sheet, and rebuilds them. The VBA project rides along, so the macros only ever get
pasted once. Buttons are redrawn by Workbook_Open — see vba/ThisWorkbook.txt for why."""

import argparse
from pathlib import Path

from inputs import dummy_100m_project
from timeline import build_timeline
from workbook_builder import build_workbook

ROOT = Path(__file__).resolve().parent.parent
DEFAULT_TEMPLATE = ROOT / "template" / "project123_template.xlsm"


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--template", nargs="?", const=str(DEFAULT_TEMPLATE), default=None,
                        help="macro-enabled template to inherit the VBA project from; "
                             "bare flag uses template/project123_template.xlsm")
    parser.add_argument("--out", default=None)
    args = parser.parse_args()

    if args.template and not Path(args.template).exists():
        raise SystemExit(
            f"template not found: {args.template}\n"
            "Paste the macros into a generated workbook, save it there as .xlsm, "
            "then re-run. See src/vba/README_VBA_SETUP.md."
        )

    suffix = "xlsm" if args.template else "xlsx"
    output = args.out or str(ROOT / "output" / f"project123_stage1c.{suffix}")

    inputs = dummy_100m_project()
    timeline = build_timeline(inputs.dates)
    build_workbook(inputs, timeline, output, template_path=args.template)

    carried = " (VBA carried over from template)" if args.template else ""
    print(f"Wrote {output}{carried}")


if __name__ == "__main__":
    main()
