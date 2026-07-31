"""Build the workbook.

    python build.py                          -> plain .xlsx, no macros
    python build.py --template <t.xlsm>      -> .xlsm carrying the template's VBA project

The template route reopens a workbook you have already pasted the macros into, drops
every sheet, and rebuilds them. The VBA project rides along, so the macros only ever get
pasted once. Buttons are redrawn by Workbook_Open — see vba/ThisWorkbook.txt for why."""

import argparse
import re
from datetime import date
from pathlib import Path

from inputs import dummy_100m_project
from timeline import build_timeline
from workbook_builder import build_workbook

ROOT = Path(__file__).resolve().parent.parent
DEFAULT_TEMPLATE = ROOT / "template" / "project123_template.xlsm"
OUTPUT_DIR = ROOT / "output"

# Every build gets a date + version tag (DDMMYYYY_v0.NN) so successive builds on the same
# day don't silently overwrite each other and the filename alone tells you when a given
# workbook was generated. Version resets to v0.01 each new day and bumps for every further
# build that day, scanning existing files rather than tracking state anywhere else.
FILENAME_RE = re.compile(r"^project123_stage1c_(\d{8})_v(\d+)\.(\d+)\.(xlsm|xlsx)$")


def _next_version_tag(build_date: date, suffix: str) -> str:
    date_tag = build_date.strftime("%d%m%Y")
    highest = 0
    if OUTPUT_DIR.exists():
        for f in OUTPUT_DIR.iterdir():
            m = FILENAME_RE.match(f.name)
            if m and m.group(1) == date_tag and m.group(4) == suffix:
                highest = max(highest, int(m.group(3)))
    return f"{date_tag}_v0.{highest + 1:02d}"


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
    if args.out:
        output = args.out
    else:
        version_tag = _next_version_tag(date.today(), suffix)
        output = str(OUTPUT_DIR / f"project123_stage1c_{version_tag}.{suffix}")

    inputs = dummy_100m_project()
    timeline = build_timeline(inputs.dates)
    build_workbook(inputs, timeline, output, template_path=args.template)

    carried = " (VBA carried over from template)" if args.template else ""
    print(f"Wrote {output}{carried}")


if __name__ == "__main__":
    main()
