from dataclasses import dataclass
from datetime import date

from dateutil.relativedelta import relativedelta

from inputs import ProjectDates


@dataclass
class Period:
    index: int  # 0-based, continuous across construction + operations
    start: date
    end: date
    is_construction: bool
    quarter_index: int | None  # None during construction; 0-based within operations
    year_index: int  # 0-based calendar year of period, used for annual rollup


@dataclass
class Timeline:
    construction_months: list[Period]
    operations_quarters: list[Period]

    @property
    def all_periods(self) -> list[Period]:
        return self.construction_months + self.operations_quarters

    def annual_buckets(self) -> dict[int, list[Period]]:
        """Group all periods (construction + operations) by project year for FS_Annual rollup."""
        buckets: dict[int, list[Period]] = {}
        for p in self.all_periods:
            buckets.setdefault(p.year_index, []).append(p)
        return buckets


def build_timeline(dates: ProjectDates) -> Timeline:
    construction_months = _build_construction_months(dates)
    if construction_months:
        ops_start = construction_months[-1].end + relativedelta(days=1)
    else:
        ops_start = dates.construction_start
    operations_quarters = _build_operations_quarters(ops_start, dates.operations_years, dates.construction_start)
    return Timeline(construction_months=construction_months, operations_quarters=operations_quarters)


def _build_construction_months(dates: ProjectDates) -> list[Period]:
    periods = []
    cursor = dates.construction_start
    for i in range(dates.construction_months):
        month_end = cursor + relativedelta(months=1) - relativedelta(days=1)
        year_index = _year_index(cursor, dates.construction_start)
        periods.append(
            Period(
                index=i,
                start=cursor,
                end=month_end,
                is_construction=True,
                quarter_index=None,
                year_index=year_index,
            )
        )
        cursor = cursor + relativedelta(months=1)
    return periods


def _build_operations_quarters(ops_start: date, operations_years: int, project_start: date) -> list[Period]:
    periods = []
    cursor = ops_start
    n_quarters = operations_years * 4
    for i in range(n_quarters):
        quarter_end = cursor + relativedelta(months=3) - relativedelta(days=1)
        year_index = _year_index(cursor, project_start)
        periods.append(
            Period(
                index=i,
                start=cursor,
                end=quarter_end,
                is_construction=False,
                quarter_index=i,
                year_index=year_index,
            )
        )
        cursor = cursor + relativedelta(months=3)
    return periods


def _year_index(current: date, project_start: date) -> int:
    return current.year - project_start.year
