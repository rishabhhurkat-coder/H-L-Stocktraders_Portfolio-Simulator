from __future__ import annotations

import os
import re
from pathlib import Path
from datetime import date
from time import sleep

from rich import box
from rich.console import Console, Group
from rich.panel import Panel
from rich.table import Table
from rich.text import Text

from portfolio_simulator.formatting import (
    format_currency,
    format_month_year,
    format_percentage,
    format_tenure,
)
from portfolio_simulator.models import CashFlowEvent, Scenario, SimulationResult
from portfolio_simulator.reporting import SheetPreview, build_export_previews, export_reports
from portfolio_simulator.simulation import (
    add_months,
    maximum_monthly_swp,
    months_between,
    projected_value_at_month,
    run_simulation,
    sip_end_date,
    swp_start_date,
    swp_start_month,
    total_investment_months,
    total_swp_months,
)


class PortfolioSimulatorApp:
    RUPEE = "\u20b9"
    PROMPT_WIDTH = 24
    MAX_SCREEN_WIDTH = 76

    def __init__(self) -> None:
        self.console = Console()

    def clear_console(self) -> None:
        if os.name == "nt":
            os.system("cls")
        else:
            self.console.clear()

    @property
    def screen_width(self) -> int:
        return min(self.MAX_SCREEN_WIDTH, max(60, self.console.size.width - 4))

    def run(self) -> None:
        while True:
            choice = self.launch_screen()
            if choice == 2:
                return

            scenario = Scenario()
            self.investment_setup(scenario)
            self.return_assumptions(scenario)
            self.stepup_screen(scenario)
            self.cash_flow_screen(scenario)
            self.swp_screen(scenario)

            action = self.summary_loop(scenario)
            if action == "exit":
                return

    def launch_screen(self) -> int:
        self.clear_console()
        self.render_header("MUTUAL FUND PORTFOLIO SIMULATOR")
        self.render_lines(["1  New Simulation", "2  Exit"])
        self.render_divider()
        return self.prompt_menu_choice("Choice", {"1": 1, "2": 2})

    def investment_setup(self, scenario: Scenario) -> None:
        self.render_prompt_screen("INVESTMENT SETUP", scenario)
        scenario.sip_start_date = self.prompt_month_year("SIP Start Date")

        self.render_prompt_screen("INVESTMENT SETUP", scenario)
        scenario.monthly_sip = self.prompt_positive_float("Monthly SIP Amount")

        self.render_prompt_screen("INVESTMENT SETUP", scenario)
        scenario.investment_years = self.prompt_non_negative_int("Tenure Years")
        scenario.investment_months = self.prompt_bounded_int("Tenure Months", minimum=0, maximum=11)

        if total_investment_months(scenario) == 0:
            scenario.investment_months = 1

    def return_assumptions(self, scenario: Scenario) -> None:
        self.render_prompt_screen("RETURN ASSUMPTIONS", scenario)
        scenario.annual_roi = self.prompt_non_negative_float("Expected ROI (Annual)")

        self.render_prompt_screen("RETURN ASSUMPTIONS", scenario)
        scenario.inflation_rate = self.prompt_non_negative_float("Inflation Rate")

    def stepup_screen(self, scenario: Scenario) -> None:
        self.render_screen(
            "STEP-UP SIP",
            scenario,
            lines=["Increase SIP every year?"],
            menu_lines=["1  Yes", "2  No"],
        )
        scenario.step_up_enabled = self.prompt_menu_choice("Choice", {"1": True, "2": False})
        scenario.step_up_rate = 0.0

        if scenario.step_up_enabled:
            self.render_prompt_screen("STEP-UP SIP", scenario)
            scenario.step_up_rate = self.prompt_non_negative_float("Annual Step-Up Rate")

    def cash_flow_screen(self, scenario: Scenario) -> None:
        cash_flow_question = Text.assemble(
            "Do you want to ",
            ("ADD", "bold green"),
            " or ",
            ("WITHDRAW", "bold red"),
            " some amount during investment?",
        )
        self.render_screen(
            "INVESTMENT CASH FLOWS",
            scenario,
            lines=[cash_flow_question],
            menu_lines=["1  Yes", "2  No"],
        )
        wants_cash_flows = self.prompt_menu_choice("Choice", {"1": True, "2": False})
        if not wants_cash_flows:
            scenario.cash_flow_events.clear()
            return

        if scenario.sip_start_date is None:
            return

        while True:
            self.render_cash_flow_type_screen(scenario)
            flow_type = self.prompt_menu_choice("Choice", {"1": "add", "2": "withdraw"})
            event_date = self.prompt_cash_flow_event_date(scenario, flow_type)
            amount = self.prompt_positive_float("Amount", label_style="bold cyan")
            scenario.cash_flow_events.append(CashFlowEvent(flow_type=flow_type, event_date=event_date, amount=amount))
            scenario.cash_flow_events.sort(key=lambda event: event.event_date)

            self.render_cash_flow_review(scenario)
            choice = self.prompt_menu_choice("Choice", {"1": 1, "2": 2, "3": 3})
            if choice == 1:
                continue
            if choice == 2:
                return
            scenario.cash_flow_events.clear()

    def swp_screen(self, scenario: Scenario) -> None:
        self.render_screen(
            "SYSTEMATIC WITHDRAWAL",
            scenario,
            lines=["Do you want SWP?"],
            menu_lines=["1  Yes", "2  No"],
        )
        scenario.swp_enabled = self.prompt_menu_choice("Choice", {"1": True, "2": False})
        scenario.swp_years = 0
        scenario.swp_months = 0
        scenario.monthly_swp_amount = 0.0
        scenario.swp_start_mode = "after_sip"
        scenario.swp_start_year = 0

        if not scenario.swp_enabled:
            return

        self.render_screen(
            "SWP START",
            scenario,
            lines=["From when should SWP start?"],
            menu_lines=["1  After End of SIP", "2  Specific Years After Start of SIP"],
        )
        start_mode = self.prompt_menu_choice("Choice", {"1": "after_sip", "2": "after_start_years"})
        scenario.swp_start_mode = start_mode

        if scenario.swp_start_mode == "after_start_years":
            self.render_screen("SWP START", scenario, lines=["Enter years after SIP start for SWP to begin."])
            scenario.swp_start_year = self.prompt_non_negative_int("Start After Years")

        self.render_prompt_screen("SWP DURATION", scenario)
        scenario.swp_years = self.prompt_non_negative_int("SWP Duration Years")
        scenario.swp_months = self.prompt_bounded_int("SWP Duration Months", minimum=0, maximum=11)

        if total_swp_months(scenario) == 0:
            scenario.swp_months = 1

        start_month = swp_start_month(scenario) or 0
        swp_value_at_start = projected_value_at_month(scenario, start_month - 1)
        max_swp = maximum_monthly_swp(scenario)

        self.render_screen(
            "SWP AMOUNT",
            scenario,
            rows=[
                ("SWP Start Date", format_month_year(swp_start_date(scenario))),
                ("Expected Fund Value", format_currency(swp_value_at_start)),
                ("SWP Duration", format_tenure(scenario.swp_years, scenario.swp_months)),
                ("Max Monthly SWP", format_currency(max_swp)),
            ],
        )

        if max_swp <= 0:
            self.console.print("No SWP capacity available for the selected start and duration.")
            self.pause()
            scenario.swp_enabled = False
            return

        scenario.monthly_swp_amount = self.prompt_bounded_float(
            "Monthly SWP Amount",
            minimum=0.01,
            maximum=max_swp,
        )

    def summary_loop(self, scenario: Scenario) -> str:
        while True:
            choice = self.scenario_summary(scenario)
            if choice == 1:
                self.running_simulation()
                result = run_simulation(scenario)
                results_action = self.results_loop(scenario, result)
                if results_action == "new":
                    return "new"
                if results_action == "edit":
                    self.edit_menu(scenario)
                    continue
                return "exit"
            if choice == 2:
                self.edit_menu(scenario)
                continue
            return "new"

    def scenario_summary(self, scenario: Scenario) -> int:
        self.render_screen(
            "REVIEW SCENARIO",
            scenario,
            lines=["Use the live summary above to review the scenario."],
            menu_lines=["1  Run Simulation", "2  Edit Inputs", "3  Cancel"],
        )
        return self.prompt_menu_choice("Choice", {"1": 1, "2": 2, "3": 3})

    def running_simulation(self) -> None:
        self.render_screen("RUNNING SIMULATION", None, lines=["Calculating portfolio growth..."])
        sleep(0.8)

    def results_loop(self, scenario: Scenario, result: SimulationResult) -> str:
        while True:
            results_choice = self.display_results(scenario, result)
            if results_choice == 1:
                return "new"
            if results_choice == 2:
                return "edit"
            if results_choice == 3:
                self.export_preview_menu(scenario, result)
                continue
            return "exit"

    def display_results(self, scenario: Scenario, result: SimulationResult) -> int:
        rows = [
            ("→ Total Invested", self.results_currency(result.total_invested), True),
            ("→ SIP Phase Value", self.results_currency(result.sip_end_value), True),
            ("→ SWP Start Value", self.results_currency(result.swp_start_value), True),
            ("→ Monthly SWP Used", self.results_currency(result.actual_monthly_swp), True),
            ("→ Final Portfolio Value", self.results_currency(result.final_portfolio_value), True),
            ("→ Total Profit", self.results_currency(result.total_profit), True),
            ("→ Portfolio CAGR", format_percentage(result.cagr * 100), True),
            ("→ Inflation Adjusted Value", self.results_currency(result.inflation_adjusted_value), True),
        ]
        self.clear_console()
        self.render_header("SIMULATION RESULTS")
        self.console.print(self.summary_panel(scenario))
        self.render_divider()
        self.console.print(self.key_value_rows(rows, label_width=28))
        self.render_divider()
        self.render_lines(
            [
                "1  New Simulation",
                "2  Edit Scenario",
                "3  Export Preview",
                "4  Exit",
            ]
        )
        self.render_divider()
        return self.prompt_menu_choice("Choice", {"1": 1, "2": 2, "3": 3, "4": 4})

    def export_preview_menu(self, scenario: Scenario, result: SimulationResult) -> None:
        previews = build_export_previews(scenario, result)

        while True:
            menu_lines = [f"{index}  {sheet.name}" for index, sheet in enumerate(previews, start=1)]
            export_option = len(previews) + 1
            back_option = len(previews) + 2
            menu_lines.append(f"{export_option}  Export to Excel + CSV")
            menu_lines.append(f"{back_option}  Back")

            self.render_screen(
                "EXPORT PREVIEW",
                scenario,
                lines=[
                    "These are the workbook sheets that will go to Excel.",
                    "No Excel file is written yet. Review sheet by sheet first.",
                ],
                menu_lines=menu_lines,
            )

            options = {str(index): index for index in range(1, back_option + 1)}
            choice = self.prompt_menu_choice("Choice", options)
            if choice == export_option:
                workbook_path, csv_path = export_reports(Path.cwd(), scenario, result)
                self.render_screen(
                    "EXPORT COMPLETE",
                    scenario,
                    lines=[
                        "Excel and CSV files were generated.",
                        f"Excel : {workbook_path}",
                        f"CSV   : {csv_path}",
                    ],
                )
                self.pause()
                continue
            if choice == back_option:
                return
            self.display_export_sheet(previews[choice - 1])

    def display_export_sheet(self, sheet: SheetPreview) -> None:
        self.clear_console()
        self.render_header(f"EXPORT SHEET - {sheet.name}")

        if sheet.rows:
            rows = [(label, value, False) for label, value in sheet.rows]
            self.console.print(self.key_value_rows(rows))

        if sheet.table_rows:
            if sheet.rows:
                self.render_divider()
            self.console.print(self.preview_table(sheet))
            if sheet.footer_rows:
                self.render_divider()
                self.console.print(
                    Panel(
                        self.totals_rows(sheet.footer_rows),
                        title=Text(sheet.footer_title or "SUMMARY", style="bold cyan"),
                        title_align="center",
                        box=box.SQUARE,
                        padding=(0, 1),
                        width=self.screen_width,
                    )
                )

        self.render_divider()
        self.pause()

    def edit_menu(self, scenario: Scenario) -> None:
        self.clear_console()
        self.render_header("EDIT PARAMETERS")
        self.console.print(self.summary_panel(scenario))
        self.render_divider()
        self.console.print(self.edit_menu_panel())
        self.render_divider()
        choice = self.prompt_menu_choice(
            "Choice",
            {"1": 1, "2": 2, "3": 3, "4": 4, "5": 5, "6": 6, "7": 7, "8": 8, "9": 9},
        )

        if choice == 1:
            self.render_prompt_screen("EDIT SIP START DATE", scenario)
            scenario.sip_start_date = self.prompt_month_year("SIP Start Date")
        elif choice == 2:
            self.render_prompt_screen("EDIT SIP AMOUNT", scenario)
            scenario.monthly_sip = self.prompt_positive_float("Monthly SIP Amount")
        elif choice == 3:
            self.render_prompt_screen("EDIT TENURE", scenario)
            scenario.investment_years = self.prompt_non_negative_int("Tenure Years")
            scenario.investment_months = self.prompt_bounded_int("Tenure Months", minimum=0, maximum=11)
            if total_investment_months(scenario) == 0:
                scenario.investment_months = 1
        elif choice == 4:
            self.render_prompt_screen("EDIT ROI", scenario)
            scenario.annual_roi = self.prompt_non_negative_float("Expected ROI (Annual)")
        elif choice == 5:
            self.render_prompt_screen("EDIT INFLATION", scenario)
            scenario.inflation_rate = self.prompt_non_negative_float("Inflation Rate")
        elif choice == 6:
            self.stepup_screen(scenario)
        elif choice == 7:
            self.cash_flow_screen(scenario)
        elif choice == 8:
            self.swp_screen(scenario)

    def render_screen(
        self,
        section_title: str,
        scenario: Scenario | None,
        rows: list[tuple[str, str]] | None = None,
        lines: list[object] | None = None,
        menu_lines: list[str] | None = None,
    ) -> None:
        self.clear_console()
        self.render_header(section_title)
        if scenario is not None:
            self.console.print(self.summary_panel(scenario))
            self.render_divider()
        if rows:
            display_rows = [(label, value, False) for label, value in rows]
            self.console.print(self.key_value_rows(display_rows))
        if lines:
            self.render_lines(lines)
        if menu_lines:
            if rows or lines:
                self.console.print("")
            self.render_lines(menu_lines)
        self.render_divider()

    def render_prompt_screen(self, section_title: str, scenario: Scenario | None) -> None:
        self.clear_console()
        self.render_header(section_title)
        if scenario is not None:
            self.console.print(self.summary_panel(scenario))
        self.render_divider()

    def render_cash_flow_type_screen(self, scenario: Scenario) -> None:
        rows = [("Investment Window", self.cash_flow_window(scenario))]
        self.render_screen(
            "INVESTMENT CASH FLOWS",
            scenario,
            rows=rows,
            lines=["Select cash flow type."],
            menu_lines=[
                Text.assemble("1  ", ("Add", "green")),
                Text.assemble("2  ", ("Withdraw", "red")),
            ],
        )
        if scenario.cash_flow_events:
            self.console.print(self.cash_flow_table(scenario))
            self.render_divider()

    def render_cash_flow_review(self, scenario: Scenario) -> None:
        self.clear_console()
        self.render_divider()
        self.console.print(self.cash_flow_table(scenario))
        self.render_divider()
        self.render_lines(["1  Add Another Entry", "2  Confirm Cash Flows", "3  Clear All"])
        self.render_divider()

    def render_cash_flow_preview(
        self,
        scenario: Scenario,
        flow_type: str,
        years_after_start: int | None,
        event_date: date,
        fund_value: float,
    ) -> None:
        action = "Add" if flow_type == "add" else "Withdraw"
        rows = [("Type", action)]
        if years_after_start is not None:
            rows.append(("After Years", str(years_after_start)))
        rows.extend(
            [
                ("Event Date", format_month_year(event_date)),
                ("Fund Value", format_currency(fund_value)),
            ]
        )
        change_label = "Change Years" if years_after_start is not None else "Change Date"
        self.render_screen(
            "CASH FLOW PREVIEW",
            scenario,
            rows=rows,
            menu_lines=["1  Continue", f"2  {change_label}"],
        )

    def render_cash_flow_action_screen(self, scenario: Scenario, flow_type: str) -> None:
        action_title = "ADD AMOUNT" if flow_type == "add" else "WITHDRAW AMOUNT"
        action_verb = "add" if flow_type == "add" else "withdraw"
        self.clear_console()
        self.render_header("INVESTMENT CASH FLOWS")
        self.console.print(self.summary_panel(scenario))
        self.render_divider()
        self.render_subheader(action_title)
        self.render_divider()
        self.console.print(self.key_value_rows([("Investment Window", self.cash_flow_window(scenario), False)]))
        self.console.print("")
        self.render_lines(
            [
                f"1  After how many years do you want to {action_verb}?",
                "2  Choose Specific Date",
            ]
        )
        self.render_divider()

    def render_subheader(self, title: str) -> None:
        self.render_divider()
        self.console.print(Text(title.center(self.screen_width), style="bold cyan"))
        self.render_divider()

    def render_header(self, title: str) -> None:
        self.render_divider()
        self.console.print(Text(title.center(self.screen_width), style="bold cyan"), end="")
        self.console.print("")
        self.render_divider()

    def render_divider(self) -> None:
        self.console.print("\u2500" * self.screen_width)

    def summary_panel(self, scenario: Scenario) -> Panel:
        return Panel(
            self.summary_body(scenario),
            title=Text("SCENARIO SUMMARY", style="bold cyan"),
            title_align="center",
            box=box.SQUARE,
            padding=(0, 1),
            width=self.screen_width,
        )

    def edit_menu_panel(self) -> Panel:
        return Panel(
            self.two_column_menu(
                [
                    "1  SIP Start Date",
                    "2  SIP Amount",
                    "3  Tenure",
                    "4  ROI",
                    "5  Inflation",
                    "6  Step-Up",
                    "7  Cash Flows",
                    "8  SWP",
                    "9  Back",
                ]
            ),
            title=Text("EDIT MENU", style="bold cyan"),
            title_align="center",
            box=box.SQUARE,
            padding=(0, 1),
            width=self.screen_width,
        )

    def key_value_rows(self, rows: list[tuple[str, object, bool]], label_width: int = 22) -> Table:
        value_width = max(16, self.screen_width - (label_width + 12))
        table = Table.grid(expand=False)
        table.add_column(justify="left", width=label_width, no_wrap=True)
        table.add_column(justify="left", width=1)
        table.add_column(justify="left", width=2)
        table.add_column(justify="left", width=value_width)

        for label, value, important in rows:
            if isinstance(value, Text):
                renderable = value
            else:
                style = "green" if important else ""
                renderable = Text(str(value), style=style)
            colon = ":" if label else ""
            gap = "  " if label else ""
            table.add_row(label, colon, gap, renderable)

        return table

    def totals_rows(self, rows: list[tuple[str, str]]) -> Table:
        value_width = max(16, self.screen_width - 34)
        table = Table.grid(expand=False)
        table.add_column(justify="left", width=22, no_wrap=True)
        table.add_column(justify="left", width=1)
        table.add_column(justify="left", width=2)
        table.add_column(justify="left", width=value_width)

        for label, value in rows:
            table.add_row(
                Text(label, style="bold cyan"),
                ":",
                "",
                Text(value, style=self.total_value_style(value)),
            )

        return table

    def total_value_style(self, value: str) -> str:
        stripped = value.strip()
        if stripped.startswith("-"):
            return "bold red"
        if stripped and stripped != "-":
            return "bold green"
        return ""

    def two_column_menu(self, items: list[str]) -> Table:
        table = Table.grid(expand=False)
        column_width = max(20, (self.screen_width - 8) // 2)
        table.add_column(justify="left", width=column_width, no_wrap=True)
        table.add_column(justify="left", width=column_width, no_wrap=True)
        for index in range(0, len(items), 2):
            left = items[index]
            right = items[index + 1] if index + 1 < len(items) else ""
            table.add_row(left, right)
        return table

    def preview_table(self, sheet: SheetPreview) -> Table:
        table = Table(box=box.SQUARE, expand=False)
        for header in sheet.headers or []:
            table.add_column(header)
        section_before = set(sheet.section_before_rows or [])
        section_after = set(sheet.section_after_rows or [])
        for index, row in enumerate(sheet.table_rows or []):
            if index in section_before and index > 0:
                table.add_section()
            table.add_row(*row)
            if index in section_after:
                table.add_section()
        return table

    def cash_flow_table(self, scenario: Scenario) -> Table:
        table = Table(box=box.SQUARE, expand=False)
        table.add_column("#")
        table.add_column("Type")
        table.add_column("Date")
        table.add_column("Amount")
        for index, event in enumerate(scenario.cash_flow_events, start=1):
            style = "green" if event.flow_type == "add" else "red"
            table.add_row(
                str(index),
                Text(event.flow_type.title(), style=style),
                Text(format_month_year(event.event_date), style=style),
                Text(format_currency(event.amount), style=style),
                style=style,
            )
        return table

    def render_lines(self, lines: list[object]) -> None:
        for line in lines:
            self.console.print(line)

    def prompt_menu_choice(self, label: str, options: dict[str, object]) -> object:
        while True:
            choice = self.prompt_text(label).strip()
            if choice in options:
                return options[choice]
            self.console.print("Enter a valid option.")

    def prompt_month_year(self, label: str) -> date:
        while True:
            raw = self.prompt_text(label).strip()
            parsed = self.parse_month_year(raw)
            if parsed is not None:
                return parsed
            self.console.print("Use Jan-25, Jan-2025, 1-25, 01-2025, or 01-08.")

    def prompt_cash_flow_event_date(self, scenario: Scenario, flow_type: str) -> date:
        while True:
            self.render_cash_flow_action_screen(scenario, flow_type)
            choice = self.prompt_menu_choice("Choice", {"1": "years", "2": "date"})
            if choice == "years":
                return self.prompt_cash_flow_timing_by_years(scenario, flow_type)
            return self.prompt_cash_flow_timing_by_date(scenario, flow_type)

    def prompt_cash_flow_timing_by_years(self, scenario: Scenario, flow_type: str) -> date:
        assert scenario.sip_start_date is not None
        max_year = max(1, total_investment_months(scenario) // 12)
        action = "add" if flow_type == "add" else "withdraw"
        while True:
            self.render_prompt_screen("INVESTMENT CASH FLOWS", scenario)
            self.console.print(self.key_value_rows([("Investment Window", self.cash_flow_window(scenario), False)]))
            self.console.print("")
            self.render_lines([f"After how many years do you want to {action}?"])
            self.render_divider()
            years_after_start = self.prompt_bounded_int("Years", minimum=1, maximum=max_year)
            event_date = add_months(scenario.sip_start_date, years_after_start * 12)
            fund_value = projected_value_at_month(scenario, (years_after_start * 12) - 1)
            self.render_cash_flow_preview(scenario, flow_type, years_after_start, event_date, fund_value)
            choice = self.prompt_menu_choice("Choice", {"1": 1, "2": 2})
            if choice == 1:
                return event_date

    def prompt_cash_flow_timing_by_date(self, scenario: Scenario, flow_type: str) -> date:
        assert scenario.sip_start_date is not None
        end_date = sip_end_date(scenario)
        assert end_date is not None
        while True:
            self.render_prompt_screen("INVESTMENT CASH FLOWS", scenario)
            self.console.print(self.key_value_rows([("Investment Window", self.cash_flow_window(scenario), False)]))
            self.render_divider()
            event_date = self.prompt_month_year("Specific Date")
            if not scenario.sip_start_date <= event_date <= end_date:
                self.console.print(
                    f"Enter a date between {format_month_year(scenario.sip_start_date)} and {format_month_year(end_date)}."
                )
                continue
            month_offset = months_between(scenario.sip_start_date, event_date)
            fund_value = projected_value_at_month(scenario, month_offset - 1)
            self.render_cash_flow_preview(scenario, flow_type, None, event_date, fund_value)
            choice = self.prompt_menu_choice("Choice", {"1": 1, "2": 2})
            if choice == 1:
                return event_date

    def prompt_positive_float(self, label: str, label_style: str = "") -> float:
        while True:
            value = self.prompt_float(label, label_style=label_style)
            if value > 0:
                return value
            self.console.print("Enter a value greater than 0.")

    def prompt_bounded_float(self, label: str, minimum: float, maximum: float) -> float:
        while True:
            value = self.prompt_float(label)
            if minimum <= value <= maximum:
                return value
            self.console.print(f"Enter a value between {minimum:.2f} and {maximum:.2f}.")

    def prompt_non_negative_float(self, label: str) -> float:
        while True:
            value = self.prompt_float(label)
            if value >= 0:
                return value
            self.console.print("Enter 0 or more.")

    def prompt_float(self, label: str, label_style: str = "") -> float:
        while True:
            raw = self.prompt_text(label, label_style=label_style).replace(",", "").replace(self.RUPEE, "").strip()
            try:
                return float(raw)
            except ValueError:
                self.console.print("Enter a valid number.")

    def prompt_non_negative_int(self, label: str) -> int:
        while True:
            value = self.prompt_int(label)
            if value >= 0:
                return value
            self.console.print("Enter 0 or more.")

    def prompt_bounded_int(self, label: str, minimum: int, maximum: int) -> int:
        while True:
            value = self.prompt_int(label)
            if minimum <= value <= maximum:
                return value
            self.console.print(f"Enter a value between {minimum} and {maximum}.")

    def prompt_int(self, label: str) -> int:
        while True:
            raw = self.prompt_text(label).strip()
            try:
                return int(raw)
            except ValueError:
                self.console.print("Enter a whole number.")

    def prompt_text(self, label: str, label_style: str = "") -> str:
        prompt = Text()
        prompt.append(f"{label:<{self.PROMPT_WIDTH}}", style=label_style)
        prompt.append(": ")
        value = self.console.input(prompt)
        self.console.print("")
        return value

    def parse_month_year(self, raw: str) -> date | None:
        parts = [part for part in re.split(r"[\s/-]+", raw.strip()) if part]
        if len(parts) != 2:
            return None

        month = self.parse_month(parts[0])
        year = self.parse_year(parts[1])
        if month is None or year is None:
            return None

        try:
            return date(year, month, 1)
        except ValueError:
            return None

    def parse_month(self, raw: str) -> int | None:
        month_map = {
            "jan": 1,
            "feb": 2,
            "mar": 3,
            "apr": 4,
            "may": 5,
            "jun": 6,
            "jul": 7,
            "aug": 8,
            "sep": 9,
            "oct": 10,
            "nov": 11,
            "dec": 12,
        }
        token = raw.strip().lower()
        if token.isdigit():
            month = int(token)
            return month if 1 <= month <= 12 else None
        key = token[:3]
        return month_map.get(key)

    def parse_year(self, raw: str) -> int | None:
        token = raw.strip()
        if not token.isdigit():
            return None
        if len(token) == 2:
            value = int(token)
            return 2000 + value if value <= 49 else 1900 + value
        if len(token) == 4:
            value = int(token)
            return value if 1900 <= value <= 2100 else None
        return None

    def pause(self) -> None:
        self.console.input(f"{'Press Enter':<{self.PROMPT_WIDTH}}: ")
        self.console.print("")

    def summary_currency(self, value: float) -> str:
        return format_currency(value) if value > 0 else ""

    def results_currency(self, value: float) -> str:
        formatted = format_currency(value)
        return formatted.replace("₹", "₹ ", 1) if formatted != "-" else formatted

    def summary_month_year(self, value: date | None) -> str:
        return format_month_year(value) if value is not None else ""

    def summary_percentage(self, value: float) -> str:
        return format_percentage(value) if value > 0 else ""

    def summary_tenure(self, scenario: Scenario) -> str:
        months = total_investment_months(scenario)
        return self.long_tenure_text(scenario.investment_years, scenario.investment_months) if months > 0 else ""

    def step_up_summary(self, scenario: Scenario) -> str:
        if not scenario.step_up_enabled:
            return "No"
        if scenario.step_up_rate <= 0:
            return "Yes / rate pending"
        return f"{scenario.step_up_rate:.2f} % yearly"

    def cash_flow_window(self, scenario: Scenario) -> str:
        if scenario.sip_start_date is None:
            return "--"
        return f"{format_month_year(scenario.sip_start_date)} to {format_month_year(sip_end_date(scenario))}"

    def cash_flow_summary(self, scenario: Scenario) -> str:
        if not scenario.cash_flow_events:
            return "No"
        adds = sum(event.amount for event in scenario.cash_flow_events if event.flow_type == "add")
        withdrawals = sum(event.amount for event in scenario.cash_flow_events if event.flow_type == "withdraw")
        net = adds - withdrawals
        return f"Net {format_currency(net)}"

    def swp_summary(self, scenario: Scenario) -> str:
        if not scenario.swp_enabled:
            return "No"
        duration = self.long_tenure_text(scenario.swp_years, scenario.swp_months)
        if total_swp_months(scenario) == 0:
            return "Yes / duration pending"
        start_label = format_month_year(swp_start_date(scenario))
        amount = format_currency(scenario.monthly_swp_amount) if scenario.monthly_swp_amount > 0 else "amount pending"
        return f"{start_label}\n{duration}\n{amount}"

    def summary_body(self, scenario: Scenario) -> Group:
        lines: list[Text] = []
        lines.extend(
            self.summary_item_lines("SIP Start Date", Text(self.summary_month_year(scenario.sip_start_date)))
        )
        lines.extend(
            self.summary_item_lines("Monthly SIP", Text(self.summary_currency(scenario.monthly_sip), style="bold green"))
        )
        lines.extend(
            self.summary_item_lines("Investment Tenure", Text(self.summary_tenure(scenario), style="bold yellow"))
        )
        lines.extend(
            self.summary_item_lines("Expected Return", Text(self.summary_percentage(scenario.annual_roi), style="bold cyan"))
        )
        lines.extend(
            self.summary_item_lines("Inflation", Text(self.summary_percentage(scenario.inflation_rate), style="bold red"))
        )
        step_up_style = "bold green" if scenario.step_up_enabled else ""
        lines.extend(self.summary_item_lines("Step-Up SIP", Text(self.step_up_summary(scenario), style=step_up_style)))
        lines.extend(self.cash_flow_block_lines(scenario))
        lines.extend(self.swp_block_lines(scenario))
        return Group(*lines)

    def summary_item_lines(self, label: str, value: Text) -> list[Text]:
        return [self.summary_value_line(label, value), self.summary_separator_line()]

    def summary_value_line(self, label: str, value: Text, show_colon: bool = True) -> Text:
        content_width = self.summary_content_width()
        line = Text()
        line.append(f"{label:<{self.summary_label_width()}}")
        line.append(" :  " if show_colon else "    ")
        line.append_text(value)
        plain_len = len(line.plain)
        if plain_len < content_width:
            line.append(" " * (content_width - plain_len))
        return line

    def summary_separator_line(self) -> Text:
        return Text("─" * self.summary_content_width())

    def summary_content_width(self) -> int:
        return self.screen_width - 4

    def summary_label_width(self) -> int:
        return 21

    def summary_prefix_width(self) -> int:
        return self.summary_label_width() + 4

    def cash_flow_block_lines(self, scenario: Scenario) -> list[Text]:
        if not scenario.cash_flow_events:
            return [self.summary_value_line("Cash Flows", Text("No")), self.summary_separator_line()]

        withdrawals = [event for event in scenario.cash_flow_events if event.flow_type == "withdraw"]
        adds = [event for event in scenario.cash_flow_events if event.flow_type == "add"]
        lines: list[Text] = []
        box_lines = self.cash_flow_box_lines(withdrawals, adds)
        lines.append(self.summary_value_line("Cash Flows", box_lines[0]))
        for box_line in box_lines[1:]:
            lines.append(self.summary_value_line("", box_line))

        net = sum(event.amount for event in adds) - sum(event.amount for event in withdrawals)
        net_style = "bold green" if net > 0 else "bold red" if net < 0 else ""
        lines.append(self.summary_value_line("", Text(f"Net {format_currency(net)}", style=net_style)))
        lines.append(self.summary_separator_line())
        return lines

    def cash_flow_box_lines(self, withdrawals: list[CashFlowEvent], adds: list[CashFlowEvent]) -> list[Text]:
        total_width = self.summary_content_width() - self.summary_prefix_width()
        inner_width = max(31, total_width)
        col_width = (inner_width - 3) // 2
        right_width = inner_width - 3 - col_width

        def pad(value: str, width: int) -> str:
            return value[:width].ljust(width)

        lines = [
            Text(f"┌{'─' * col_width}┬{'─' * right_width}┐"),
            Text.assemble(
                "│",
                (pad("Withdraw", col_width), "bold red"),
                "│",
                (pad("Add", right_width), "bold green"),
                "│",
            ),
            Text(f"├{'─' * col_width}┼{'─' * right_width}┤"),
        ]

        row_count = max(len(withdrawals), len(adds), 1)
        for index in range(row_count):
            left = self.cash_flow_box_entry(withdrawals[index]) if index < len(withdrawals) else ""
            right = self.cash_flow_box_entry(adds[index]) if index < len(adds) else ""
            lines.append(
                Text.assemble(
                    "│",
                    (pad(left, col_width), "bold red" if left else ""),
                    "│",
                    (pad(right, right_width), "bold green" if right else ""),
                    "│",
                )
            )

        lines.append(Text(f"└{'─' * col_width}┴{'─' * right_width}┘"))
        return lines

    def cash_flow_box_entry(self, event: CashFlowEvent) -> str:
        amount = -event.amount if event.flow_type == "withdraw" else event.amount
        return f"{format_month_year(event.event_date)}  {format_currency(amount)}"

    def swp_block_lines(self, scenario: Scenario) -> list[Text]:
        if not scenario.swp_enabled:
            return [self.summary_value_line("SWP Plan", Text("No"))]
        if total_swp_months(scenario) == 0:
            return [self.summary_value_line("SWP Plan", Text("Yes / duration pending"))]

        amount = format_currency(scenario.monthly_swp_amount) if scenario.monthly_swp_amount > 0 else "amount pending"
        return [
            self.summary_value_line("SWP Plan", Text(format_month_year(swp_start_date(scenario)))),
            self.summary_value_line("", Text(self.long_tenure_text(scenario.swp_years, scenario.swp_months))),
            self.summary_value_line("", Text(amount, style="bold green" if scenario.monthly_swp_amount > 0 else "")),
        ]

    def long_tenure_text(self, years: int, months: int) -> str:
        parts: list[str] = []
        if years:
            parts.append(f"{years} Year" if years == 1 else f"{years} Years")
        if months:
            parts.append(f"{months} Month" if months == 1 else f"{months} Months")
        return " ".join(parts)
