import sys
from pathlib import Path

from inputs import dummy_100m_project
from timeline import build_timeline
from workbook_builder import build_workbook

DEFAULT_OUTPUT = Path(__file__).resolve().parent.parent / "output" / "project123_stage1c.xlsx"


def main(output_path: str = str(DEFAULT_OUTPUT)) -> None:
    inputs = dummy_100m_project()
    timeline = build_timeline(inputs.dates)
    build_workbook(inputs, timeline, output_path)
    print(f"Wrote {output_path}")


if __name__ == "__main__":
    main(*sys.argv[1:])
