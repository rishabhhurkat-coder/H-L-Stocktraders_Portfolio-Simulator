from __future__ import annotations

from datetime import date

from portfolio_simulator.models import Scenario, ScheduleRow, SimulationResult


def add_months(value: date, months: int) -> date:
    year = value.year + ((value.month - 1 + months) // 12)
    month = ((value.month - 1 + months) % 12) + 1
    return date(year, month, 1)


def months_between(start: date, end: date) -> int:
    return ((end.year - start.year) * 12) + (end.month - start.month)


def total_investment_months(scenario: Scenario) -> int:
    return (scenario.investment_years * 12) + scenario.investment_months


def total_swp_months(scenario: Scenario) -> int:
    return (scenario.swp_years * 12) + scenario.swp_months


def sip_end_date(scenario: Scenario) -> date | None:
    if scenario.sip_start_date is None:
        return None
    sip_months = total_investment_months(scenario)
    if sip_months <= 0:
        return scenario.sip_start_date
    return add_months(scenario.sip_start_date, sip_months - 1)


def swp_start_month(scenario: Scenario) -> int | None:
    if not scenario.swp_enabled:
        return None
    if scenario.swp_start_mode == "after_sip":
        return total_investment_months(scenario)
    if scenario.swp_start_mode == "after_start_years":
        return scenario.swp_start_year * 12
    if scenario.sip_start_date is None or scenario.swp_start_date_override is None:
        return None
    return max(0, months_between(scenario.sip_start_date, scenario.swp_start_date_override))


def swp_start_date(scenario: Scenario) -> date | None:
    if scenario.sip_start_date is None or not scenario.swp_enabled:
        return None
    if scenario.swp_start_mode == "specific_date" and scenario.swp_start_date_override is not None:
        return scenario.swp_start_date_override
    start_month = swp_start_month(scenario)
    if start_month is None:
        return None
    return add_months(scenario.sip_start_date, start_month)


def event_offset(scenario: Scenario, event_date: date) -> int:
    if scenario.sip_start_date is None:
        return 0
    return max(0, months_between(scenario.sip_start_date, event_date))


def simulate_timeline(
    scenario: Scenario,
    swp_amount_override: float | None = None,
    include_schedule: bool = False,
) -> tuple[float, float, float, int, list[ScheduleRow], bool]:
    monthly_return = (scenario.annual_roi / 100) / 12
    portfolio_value = 0.0
    total_invested = 0.0
    current_sip = scenario.monthly_sip
    sip_months = total_investment_months(scenario)
    swp_months = total_swp_months(scenario) if scenario.swp_enabled else 0
    swp_start = swp_start_month(scenario) if scenario.swp_enabled else None
    actual_swp = swp_amount_override if swp_amount_override is not None else scenario.monthly_swp_amount
    total_months = max(sip_months, (swp_start + swp_months) if swp_start is not None else 0)
    if total_months <= 0:
        total_months = sip_months

    cash_flow_map: dict[int, float] = {}
    for event in scenario.cash_flow_events:
        offset = event_offset(scenario, event.event_date)
        if 0 <= offset < max(sip_months, total_months):
            signed_amount = event.amount if event.flow_type == "add" else -event.amount
            cash_flow_map[offset] = cash_flow_map.get(offset, 0.0) + signed_amount

    schedule_rows: list[ScheduleRow] = []
    sip_end_value = 0.0
    total_months_simulated = 0
    depleted = False

    for month_index in range(total_months):
        if scenario.step_up_enabled and month_index > 0 and month_index % 12 == 0:
            current_sip *= 1 + (scenario.step_up_rate / 100)

        opening_balance = portfolio_value
        growth = opening_balance * monthly_return
        portfolio_value = opening_balance + growth

        sip_amount = current_sip if month_index < sip_months else 0.0
        if sip_amount:
            portfolio_value += sip_amount
            total_invested += sip_amount

        lumpsum_amount = cash_flow_map.get(month_index, 0.0)
        if lumpsum_amount:
            portfolio_value += lumpsum_amount
            if lumpsum_amount > 0:
                total_invested += lumpsum_amount
            if portfolio_value < 0:
                depleted = True
                portfolio_value = 0.0

        swp_amount = 0.0
        if swp_start is not None and swp_start <= month_index < (swp_start + swp_months):
            swp_amount = actual_swp
            portfolio_value -= swp_amount
            if portfolio_value < 0:
                depleted = True
                portfolio_value = 0.0

        total_months_simulated += 1

        if include_schedule:
            schedule_rows.append(
                ScheduleRow(
                    period_number=month_index + 1,
                    phase=resolve_phase(month_index, sip_months, swp_start, swp_months),
                    period_date=add_months(scenario.sip_start_date, month_index) if scenario.sip_start_date else None,
                    opening_balance=opening_balance,
                    sip_amount=sip_amount,
                    swp_amount=swp_amount,
                    lumpsum_amount=lumpsum_amount,
                    contribution=sip_amount,
                    withdrawal=swp_amount + abs(lumpsum_amount) if lumpsum_amount < 0 else swp_amount,
                    growth=growth,
                    closing_balance=portfolio_value,
                )
            )

        if month_index + 1 == sip_months:
            sip_end_value = portfolio_value

    if sip_months == 0:
        sip_end_value = portfolio_value

    return portfolio_value, total_invested, sip_end_value, total_months_simulated, schedule_rows, depleted


def resolve_phase(month_index: int, sip_months: int, swp_start: int | None, swp_months: int) -> str:
    in_sip = month_index < sip_months
    in_swp = swp_start is not None and swp_start <= month_index < (swp_start + swp_months)
    if in_sip and in_swp:
        return "SIP+SWP"
    if in_sip:
        return "SIP"
    if in_swp:
        return "SWP"
    return "HOLD"


def projected_value_at_month(scenario: Scenario, month_index: int) -> float:
    if month_index < 0:
        return 0.0
    adjusted = Scenario(
        sip_start_date=scenario.sip_start_date,
        monthly_sip=scenario.monthly_sip,
        investment_years=scenario.investment_years,
        investment_months=scenario.investment_months,
        annual_roi=scenario.annual_roi,
        inflation_rate=scenario.inflation_rate,
        step_up_enabled=scenario.step_up_enabled,
        step_up_rate=scenario.step_up_rate,
        cash_flow_events=list(scenario.cash_flow_events),
        swp_enabled=False,
    )
    final_value, _, _, _, _, _ = simulate_timeline(adjusted, include_schedule=False)
    if month_index >= total_investment_months(adjusted):
        return final_value

    month_count = month_index + 1
    partial = Scenario(
        sip_start_date=scenario.sip_start_date,
        monthly_sip=scenario.monthly_sip,
        investment_years=month_count // 12,
        investment_months=month_count % 12,
        annual_roi=scenario.annual_roi,
        inflation_rate=scenario.inflation_rate,
        step_up_enabled=scenario.step_up_enabled,
        step_up_rate=scenario.step_up_rate,
        cash_flow_events=[event for event in scenario.cash_flow_events if event_offset(scenario, event.event_date) <= month_index],
        swp_enabled=False,
    )
    value, _, _, _, _, _ = simulate_timeline(partial, include_schedule=False)
    return value


def projected_value_before_year(scenario: Scenario, year: int) -> float:
    return projected_value_at_month(scenario, (year * 12) - 1)


def projected_value_at_sip_end(scenario: Scenario) -> float:
    value, _, _, _, _, _ = simulate_timeline(
        Scenario(
            sip_start_date=scenario.sip_start_date,
            monthly_sip=scenario.monthly_sip,
            investment_years=scenario.investment_years,
            investment_months=scenario.investment_months,
            annual_roi=scenario.annual_roi,
            inflation_rate=scenario.inflation_rate,
            step_up_enabled=scenario.step_up_enabled,
            step_up_rate=scenario.step_up_rate,
            cash_flow_events=list(scenario.cash_flow_events),
            swp_enabled=False,
        ),
        include_schedule=False,
    )
    return value


def maximum_monthly_swp(scenario: Scenario) -> float:
    if not scenario.swp_enabled or total_swp_months(scenario) <= 0:
        return 0.0

    low = 0.0
    start_month = swp_start_month(scenario) or 0
    swp_start_value = projected_value_at_month(scenario, start_month - 1)
    high = max(1.0, swp_start_value)

    while swp_is_sustainable(scenario, high):
        high *= 2
        if high > 1_000_000_000:
            break

    for _ in range(50):
        mid = (low + high) / 2
        if swp_is_sustainable(scenario, mid):
            low = mid
        else:
            high = mid

    return low


def swp_is_sustainable(scenario: Scenario, monthly_swp_amount: float) -> bool:
    final_value, _, _, _, _, depleted = simulate_timeline(
        scenario,
        swp_amount_override=monthly_swp_amount,
        include_schedule=False,
    )
    return not depleted and final_value >= 0


def run_simulation(scenario: Scenario) -> SimulationResult:
    actual_monthly_swp = scenario.monthly_swp_amount if scenario.swp_enabled else 0.0
    final_value, total_invested, sip_end_value, total_months_simulated, schedule_rows, _ = simulate_timeline(
        scenario,
        swp_amount_override=actual_monthly_swp,
        include_schedule=True,
    )

    inflation_rate = scenario.inflation_rate / 100
    total_years = total_months_simulated / 12 if total_months_simulated else 0
    total_profit = final_value - total_invested

    if total_invested > 0 and final_value > 0 and total_years > 0:
        cagr = (final_value / total_invested) ** (1 / total_years) - 1
    else:
        cagr = 0.0

    if total_years > 0:
        inflation_adjusted_value = final_value / ((1 + inflation_rate) ** total_years)
    else:
        inflation_adjusted_value = final_value

    return SimulationResult(
        total_invested=total_invested,
        final_portfolio_value=final_value,
        total_profit=total_profit,
        cagr=cagr,
        inflation_adjusted_value=inflation_adjusted_value,
        total_months_simulated=total_months_simulated,
        sip_end_value=sip_end_value,
        swp_start_value=projected_value_at_month(scenario, (swp_start_month(scenario) or 0) - 1) if scenario.swp_enabled else sip_end_value,
        actual_monthly_swp=actual_monthly_swp,
        schedule_rows=schedule_rows,
    )
