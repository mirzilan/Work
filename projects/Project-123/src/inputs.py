from dataclasses import dataclass
from datetime import date


@dataclass
class ProjectDates:
    construction_start: date
    construction_months: int
    operations_years: int


@dataclass
class CapexInputs:
    total_capex: float
    phasing_pct_by_month: list[float]  # length == construction_months, sums to 1.0


@dataclass
class FinancingInputs:
    debt_pct_of_capex: float
    interest_rate_annual: float
    debt_tenor_years: int
    target_dscr: float


@dataclass
class RevenueOpexInputs:
    annual_revenue: float
    opex_pct_of_revenue: float


@dataclass
class TaxInputs:
    tax_rate: float
    useful_life_years: int


@dataclass
class ReservesInputs:
    """Window widths, not scenario drivers. Each one sets how many columns a range spans,
    and the ranges are written out explicitly at build time — a scenario-varying width
    would need OFFSET, which the model avoids. Rates that *can* vary by scenario (routine
    maintenance %, LC fee) live on Assumptions_Constant instead."""

    dsra_target_months: int
    mra_lookforward_quarters: int
    maint_capex_useful_life_years: int


@dataclass
class MaintenanceInputs:
    lumpy_overhaul_amount: float
    lumpy_overhaul_every_n_quarters: int


@dataclass
class ProjectInputs:
    dates: ProjectDates
    capex: CapexInputs
    financing: FinancingInputs
    revenue_opex: RevenueOpexInputs
    tax: TaxInputs
    reserves: ReservesInputs
    maintenance: MaintenanceInputs


def dummy_100m_project() -> ProjectInputs:
    construction_months = 24
    dates = ProjectDates(
        construction_start=date(2027, 1, 1),
        construction_months=construction_months,
        operations_years=20,
    )

    # Simple S-curve placeholder: ramps up, peaks mid-construction, tapers off.
    phasing = _s_curve_phasing(construction_months)

    capex = CapexInputs(total_capex=100_000_000.0, phasing_pct_by_month=phasing)

    financing = FinancingInputs(
        debt_pct_of_capex=0.70,
        interest_rate_annual=0.06,
        debt_tenor_years=15,
        target_dscr=1.30,
    )

    revenue_opex = RevenueOpexInputs(
        annual_revenue=15_000_000.0,  # flat dummy placeholder, refined in later stages
        opex_pct_of_revenue=0.30,
    )

    tax = TaxInputs(tax_rate=0.25, useful_life_years=20)

    reserves = ReservesInputs(
        dsra_target_months=6,
        mra_lookforward_quarters=4,
        maint_capex_useful_life_years=10,
    )

    maintenance = MaintenanceInputs(
        lumpy_overhaul_amount=2_000_000.0,
        lumpy_overhaul_every_n_quarters=20,
    )

    return ProjectInputs(
        dates=dates,
        capex=capex,
        financing=financing,
        revenue_opex=revenue_opex,
        tax=tax,
        reserves=reserves,
        maintenance=maintenance,
    )


def _s_curve_phasing(n_months: int) -> list[float]:
    import math

    weights = []
    midpoint = n_months / 2
    steepness = 6.0 / n_months
    prev_cdf = 0.0
    for m in range(1, n_months + 1):
        cdf = 1 / (1 + math.exp(-steepness * (m - midpoint)))
        weights.append(cdf - prev_cdf)
        prev_cdf = cdf

    total = sum(weights)
    return [w / total for w in weights]
