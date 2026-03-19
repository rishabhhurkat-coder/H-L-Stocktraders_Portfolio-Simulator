from __future__ import annotations

from datetime import date


def format_currency(amount: float) -> str:
    if int(round(abs(amount))) == 0:
        return "-"
    sign = "-" if amount < 0 else ""
    value = int(round(abs(amount)))
    digits = str(value)

    if len(digits) <= 3:
        return f"{sign}\u20b9{digits}"

    last_three = digits[-3:]
    remaining = digits[:-3]
    parts: list[str] = []

    while len(remaining) > 2:
        parts.append(remaining[-2:])
        remaining = remaining[:-2]

    if remaining:
        parts.append(remaining)

    grouped = ",".join(reversed(parts)) + f",{last_three}"
    return f"{sign}\u20b9{grouped}"


def format_percentage(value: float) -> str:
    return f"{value:.2f} %"


def format_years(years: int) -> str:
    return f"{years}"


def format_month_year(value: date | None) -> str:
    if value is None:
        return "--"
    return value.strftime("%b %Y")


def format_tenure(years: int, months: int) -> str:
    year_label = "Year" if years == 1 else "Years"
    month_label = "Month" if months == 1 else "Months"
    return f"{years} {year_label}, {months} {month_label}"
