from __future__ import annotations

from dataclasses import dataclass, field
from datetime import date
from typing import Literal


SWPMode = Literal["none", "fixed", "deplete", "target_balance"]
CashFlowType = Literal["add", "withdraw"]


@dataclass
class CashFlowEvent:
    flow_type: CashFlowType
    event_date: date
    amount: float


@dataclass
class Scenario:
    sip_start_date: date | None = None
    monthly_sip: float = 0.0
    investment_years: int = 0
    investment_months: int = 0
    annual_roi: float = 0.0
    inflation_rate: float = 0.0
    step_up_enabled: bool = False
    step_up_rate: float = 0.0
    cash_flow_events: list[CashFlowEvent] = field(default_factory=list)
    swp_enabled: bool = False
    swp_start_mode: Literal["after_sip", "after_start_years", "specific_date"] = "after_sip"
    swp_start_year: int = 0
    swp_start_date_override: date | None = None
    swp_years: int = 0
    swp_months: int = 0
    swp_mode: SWPMode = "none"
    monthly_swp_amount: float = 0.0
    swp_target_balance: float = 0.0


@dataclass
class ScheduleRow:
    period_number: int
    phase: str
    period_date: date | None
    opening_balance: float
    sip_amount: float
    swp_amount: float
    lumpsum_amount: float
    contribution: float
    withdrawal: float
    growth: float
    closing_balance: float


@dataclass
class SimulationResult:
    total_invested: float
    final_portfolio_value: float
    total_profit: float
    cagr: float
    inflation_adjusted_value: float
    total_months_simulated: int
    sip_end_value: float
    swp_start_value: float
    actual_monthly_swp: float
    schedule_rows: list[ScheduleRow]
