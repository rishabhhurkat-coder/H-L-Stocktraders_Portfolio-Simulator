from __future__ import annotations

import base64
from dataclasses import replace
import logging
import os
import subprocess
import sys
from datetime import date
from io import BytesIO
from pathlib import Path

import pandas as pd
import streamlit as st
from PIL import Image

from portfolio_simulator.formatting import format_currency, format_month_year, format_percentage, format_tenure
from portfolio_simulator.models import CashFlowEvent, Scenario, SimulationResult
from portfolio_simulator.reporting import export_excel_report, export_pdf_report_bytes
from portfolio_simulator.simulation import (
    add_months,
    maximum_monthly_swp,
    projected_value_at_month,
    run_simulation,
    sip_end_date,
    swp_start_date,
    swp_start_month,
    total_investment_months,
)

try:
    from streamlit.runtime.scriptrunner import get_script_run_ctx
except Exception:  # pragma: no cover
    def get_script_run_ctx() -> object | None:
        return None


logging.getLogger("streamlit.runtime.scriptrunner_utils.script_run_context").setLevel(logging.ERROR)

APP_ROOT = Path(__file__).resolve().parent
ASSETS_DIR = APP_ROOT / "assets"
NAV_DATA_FILE = APP_ROOT / "Funds Historical NAV" / "hdfc midcap fund NAV since 2007.xlsx"

LOGO_IMAGE_CANDIDATES = [
    ASSETS_DIR / "hl_logo.png",
    ASSETS_DIR / "Sp Graphics Logo.jpg",
]
LOGO_CROP_TOP_RATIO = 0.0
LOGO_DISPLAY_WIDTH = 175
OCCUPATION_OPTIONS = [
    "Business Owner",
    "Salaried Employee",
    "Student",
    "Homemaker",
    "Farmer",
    "Doctor",
    "Engineer",
    "Teacher",
    "Lawyer",
    "Chartered Accountant",
    "Banking Professional",
    "Government Employee",
    "Sales Professional",
    "Marketing Professional",
    "IT Professional",
    "Consultant",
    "Trader",
    "Retired",
    "Freelancer",
    "Other",
]
DEVELOPER_NAME = "Rishabh Hurkat"
DEVELOPER_PHONE = "88830488312"
DEVELOPER_EMAIL = "hlstocktraders@gmail.com"
SIP_ONLY_MODE = "SIP Only Mode"
SIP_CF_MODE = "SIP + Investment CF Mode"
SIP_SWP_MODE = "SIP + SWP Mode"
ALL_COMBINATION_MODE = "All Combination Mode"
ANALYSIS_VARIANT_OPTIONS = [
    SIP_ONLY_MODE,
    SIP_CF_MODE,
    SIP_SWP_MODE,
    ALL_COMBINATION_MODE,
]


def clear_console() -> None:
    os.system("cls" if os.name == "nt" else "clear")


def ensure_streamlit_runtime() -> None:
    if __name__ != "__main__":
        return
    if get_script_run_ctx() is not None:
        return

    script_path = Path(__file__).resolve()
    clear_console()
    print("Streamlit Connected!!", flush=True)
    subprocess.run(
        [
            sys.executable,
            "-m",
            "streamlit",
            "run",
            str(script_path),
            "--browser.gatherUsageStats",
            "false",
            "--browser.serverAddress",
            "localhost",
            "--server.address",
            "127.0.0.1",
            "--server.headless",
            "false",
        ],
        check=False,
        stdout=subprocess.DEVNULL,
        stderr=subprocess.DEVNULL,
    )
    raise SystemExit


ensure_streamlit_runtime()

st.set_page_config(page_title="Mutual Fund Portfolio Simulator", page_icon="chart_with_upwards_trend", layout="wide")


def apply_theme() -> None:
    st.markdown(
        """
        <style>
        :root {
            color-scheme: light;
            --app-bg: #ffffff;
            --surface-bg: #f6f9fc;
            --border-color: #dbe5ee;
            --heading-color: #103b52;
            --text-color: #16202a;
            --muted-text: #4c5d6b;
            --button-primary: #93C5FD;
            --button-primary-hover: #60A5FA;
            --button-primary-border: #60A5FA;
            --button-primary-text: #103b52;
        }
        html, body, [data-testid="stAppViewContainer"], [data-testid="stHeader"] {
            background: var(--app-bg);
            color: var(--text-color);
        }
        .stApp {
            background: var(--app-bg);
            color: var(--text-color);
        }
        .block-container {
            padding-top: 1.25rem;
            padding-bottom: 2rem;
            max-width: 1400px;
        }
        h1, h2, h3 {
            color: var(--heading-color);
        }
        [data-testid="stMetricValue"] {
            color: #107c41;
        }
        [data-testid="stToolbar"] {
            color-scheme: light;
        }
        .stButton > button,
        .stDownloadButton > button {
            border-radius: 8px !important;
            font-weight: 600 !important;
        }
        .stButton > button[kind="primary"],
        .stDownloadButton > button[kind="primary"],
        [data-testid="stBaseButton-primary"] {
            background: var(--button-primary) !important;
            color: var(--button-primary-text) !important;
            border: 1px solid var(--button-primary-border) !important;
        }
        .stButton > button[kind="primary"]:hover,
        .stDownloadButton > button[kind="primary"]:hover,
        [data-testid="stBaseButton-primary"]:hover {
            background: var(--button-primary-hover) !important;
            border-color: var(--button-primary-border) !important;
            color: var(--button-primary-text) !important;
        }
        .stButton > button[kind="secondary"],
        .stDownloadButton > button[kind="secondary"],
        [data-testid="stBaseButton-secondary"] {
            background: linear-gradient(180deg, #ffffff 0%, #f3f8ff 100%) !important;
            color: var(--button-primary) !important;
            border: 1px solid rgba(37, 99, 235, 0.24) !important;
        }
        .stButton > button[kind="secondary"]:hover,
        .stDownloadButton > button[kind="secondary"]:hover,
        [data-testid="stBaseButton-secondary"]:hover {
            background: #eff6ff !important;
            color: var(--button-primary-hover) !important;
            border-color: rgba(59, 130, 246, 0.32) !important;
        }
        .stButton > button:focus,
        .stDownloadButton > button:focus,
        [data-testid="stBaseButton-primary"]:focus,
        [data-testid="stBaseButton-secondary"]:focus {
            box-shadow: 0 0 0 0.2rem rgba(37, 99, 235, 0.18) !important;
        }
        [data-testid="stDataFrame"] {
            margin-top: 0 !important;
        }
        [data-testid="stDataFrame"] [role="columnheader"],
        [data-testid="stDataFrame"] [role="gridcell"] {
            padding-top: 12px !important;
            padding-bottom: 12px !important;
        }
        [data-baseweb="input"] input,
        [data-baseweb="base-input"] input,
        div[data-baseweb="select"] > div,
        textarea,
        .stDateInput input,
        .stNumberInput input,
        .stTextInput input {
            background: #ffffff !important;
            color: var(--text-color) !important;
        }
        div[data-baseweb="select"] *,
        .stRadio label,
        .stCheckbox label,
        .stMarkdown,
        .stText,
        p,
        label {
            color: var(--text-color);
        }
        .dashboard-card {
            background: var(--surface-bg);
            border: 1px solid var(--border-color);
            border-radius: 14px;
            padding: 18px 20px;
            margin-bottom: 14px;
        }
        .section-title {
            font-size: 0.9rem;
            font-weight: 700;
            letter-spacing: 0.04em;
            color: #0b7285;
            text-transform: uppercase;
            margin-bottom: 0.35rem;
        }
        .field-group-title {
            font-size: 0.85rem;
            font-weight: 800;
            letter-spacing: 0.03em;
            color: #103b52;
            text-transform: uppercase;
            margin: 0.25rem 0 0.5rem 0;
        }
        .summary-grid {
            display: grid;
            grid-template-columns: 190px 1fr;
            gap: 8px 16px;
            font-size: 0.97rem;
        }
        .summary-label {
            color: var(--muted-text);
            font-weight: 600;
        }
        .summary-value {
            color: var(--text-color);
            font-weight: 600;
        }
        .small-note {
            color: #61707f;
            font-size: 0.9rem;
        }
        [data-testid="stWidgetLabel"] p {
            color: #103b52 !important;
            font-weight: 700 !important;
            opacity: 1 !important;
        }
        .brand-logo-card {
            border: 1px solid var(--border-color);
            border-radius: 14px;
            padding: 10px;
            background: var(--surface-bg);
        }
        .profile-card {
            background: var(--surface-bg);
            border: 1px solid var(--border-color);
            border-radius: 14px;
            padding: 14px 16px;
            min-height: 110px;
        }
        .profile-title {
            color: #103b52;
            font-weight: 700;
            letter-spacing: 0.03em;
            text-transform: uppercase;
            font-size: 0.85rem;
            margin-bottom: 0.35rem;
        }
        .profile-row {
            color: #000000;
            font-weight: 700;
            line-height: 1.45;
        }
        .top-brand-bar {
            display: inline-flex;
            align-items: center;
            justify-content: flex-start;
            padding: 10px 16px;
            border: 1px solid var(--border-color);
            border-radius: 12px;
            background: var(--app-bg);
            margin-bottom: 0.9rem;
            width: auto;
            max-width: 100%;
            box-sizing: border-box;
            margin-bottom: 0;
        }
        .top-brand-row {
            display: flex;
            align-items: flex-start;
            gap: 10px;
            margin-bottom: 0.8rem;
        }
        .top-brand-logo img {
            display: block;
            height: auto;
        }
        .top-brand-wrap {
            display: flex;
            align-items: flex-start;
        }
        .top-brand-contact {
            display: flex;
            flex-direction: column;
            justify-content: center;
            gap: 6px;
            color: #4f5d75;
            font-size: 0.98rem;
        }
        .top-brand-name {
            color: #103b52;
            font-weight: 700;
            line-height: 1.2;
        }
        .top-brand-line {
            display: flex;
            align-items: center;
            gap: 8px;
            white-space: nowrap;
            line-height: 1.2;
        }
        .top-brand-icon {
            color: #103b52;
            font-weight: 700;
        }
        .result-grid {
            display: grid;
            grid-template-columns: repeat(3, minmax(180px, 1fr));
            gap: 18px;
            margin-bottom: 0.35rem;
        }
        .result-item {
            background: var(--app-bg);
            border: 1px solid var(--border-color);
            border-radius: 12px;
            padding: 12px 14px;
        }
        .result-label {
            color: var(--muted-text);
            font-weight: 500;
            font-size: 0.95rem;
            margin-bottom: 6px;
        }
        .result-value {
            font-size: 2.05rem;
            line-height: 1.1;
            font-weight: 500;
        }
        .result-positive {
            color: #c92a2a;
        }
        .result-negative {
            color: #107c41;
        }
        .result-neutral {
            color: #495057;
        }
        </style>
        """,
        unsafe_allow_html=True,
    )


def default_cash_flows() -> pd.DataFrame:
    return pd.DataFrame(
        {
            "Type": pd.Series(dtype="string"),
            "Date": pd.Series(dtype="datetime64[ns]"),
            "Amount": pd.Series(dtype="float64"),
        }
    )


@st.cache_data(show_spinner=False)
def load_nav_data(path: Path) -> pd.DataFrame:
    return pd.read_excel(path)


def initialize_state() -> None:
    if "scenario_inputs" not in st.session_state:
        st.session_state.scenario_inputs = {
            "sip_start_date": date.today().replace(day=1),
            "monthly_sip": 10000.0,
            "investment_years": 10,
            "investment_months": 0,
            "annual_roi": 12.0,
            "inflation_rate": 6.0,
            "step_up_enabled": False,
            "step_up_rate": 10.0,
            "cash_flows_enabled": False,
            "swp_enabled": False,
            "swp_start_mode": "after_sip",
            "swp_start_year": 10,
            "swp_start_date_override": date.today().replace(day=1),
            "swp_years": 10,
            "swp_months": 0,
            "monthly_swp_amount": 10000.0,
        }
    if "cash_flows_df" not in st.session_state:
        st.session_state.cash_flows_df = default_cash_flows()
    else:
        st.session_state.cash_flows_df = normalize_cash_flows_df(st.session_state.cash_flows_df)
    if "cash_flows_enabled" not in st.session_state.scenario_inputs:
        st.session_state.scenario_inputs["cash_flows_enabled"] = not st.session_state.cash_flows_df.empty
    if "swp_start_date_override" not in st.session_state.scenario_inputs:
        st.session_state.scenario_inputs["swp_start_date_override"] = st.session_state.scenario_inputs["sip_start_date"]
    st.session_state.scenario_inputs["swp_months"] = 0
    if "last_result" not in st.session_state:
        st.session_state.last_result = None
    if "last_scenario" not in st.session_state:
        st.session_state.last_scenario = None
    if "customer_profile" not in st.session_state:
        st.session_state.customer_profile = None
    if "analysis_mode" not in st.session_state:
        st.session_state.analysis_mode = "Assumed Returns Basis"
    if "selected_analysis_variant" not in st.session_state:
        st.session_state.selected_analysis_variant = ALL_COMBINATION_MODE
    if "last_analysis_results" not in st.session_state:
        st.session_state.last_analysis_results = None
    if "show_scenario_comparison" not in st.session_state:
        st.session_state.show_scenario_comparison = True
    if "show_portfolio_growth" not in st.session_state:
        st.session_state.show_portfolio_growth = True


def get_selected_analysis_variant() -> str:
    selected = str(st.session_state.get("selected_analysis_variant", ALL_COMBINATION_MODE))
    if selected not in ANALYSIS_VARIANT_OPTIONS:
        selected = ALL_COMBINATION_MODE
        st.session_state.selected_analysis_variant = selected
    return selected


def render_analysis_variant_buttons(section_key: str) -> str:
    current = get_selected_analysis_variant()
    st.caption("Select Analysis Variant")
    option_cols = st.columns(len(ANALYSIS_VARIANT_OPTIONS))
    for idx, option in enumerate(ANALYSIS_VARIANT_OPTIONS):
        with option_cols[idx]:
            if st.button(
                option,
                key=f"analysis_variant_{section_key}_{idx}",
                type="primary" if option == current else "secondary",
                use_container_width=True,
            ):
                st.session_state.selected_analysis_variant = option
                st.rerun()
    return get_selected_analysis_variant()


def render_collapsible_header(title: str, state_key: str, button_key: str) -> bool:
    current_state = bool(st.session_state.get(state_key, True))
    title_col, toggle_col = st.columns([0.94, 0.06], gap="small")
    with title_col:
        st.markdown(f'<div class="section-title">{title}</div>', unsafe_allow_html=True)
    with toggle_col:
        toggle_label = "-" if current_state else "+"
        if st.button(toggle_label, key=button_key, use_container_width=True):
            st.session_state[state_key] = not current_state
            st.rerun()
    return bool(st.session_state.get(state_key, True))


@st.dialog("Customer Profile", width="large")
def render_customer_profile_dialog() -> None:
    current = st.session_state.get("customer_profile") or {}

    name = st.text_input("Customer Name *", value=str(current.get("name", "")))
    birth_date = st.date_input("Birth Date", value=current.get("birth_date", date.today()), format="DD-MM-YYYY")
    selected_occupation = str(current.get("occupation", OCCUPATION_OPTIONS[0]))
    if selected_occupation not in OCCUPATION_OPTIONS:
        selected_occupation = OCCUPATION_OPTIONS[0]
    occupation = st.selectbox(
        "Occupation",
        options=OCCUPATION_OPTIONS,
        index=OCCUPATION_OPTIONS.index(selected_occupation),
    )
    address = st.text_area("Address", value=str(current.get("address", "")), height=90)
    city = st.text_input("City *", value=str(current.get("city", "")))
    contact = st.text_input("Contact Details", value=str(current.get("contact_details", "")))

    save_col, cancel_col = st.columns(2)
    with save_col:
        save_clicked = st.button("Save Profile", type="primary", use_container_width=True)
    with cancel_col:
        cancel_clicked = st.button("Cancel", use_container_width=True)

    if save_clicked:
        if not name.strip() or not city.strip():
            st.error("Please fill mandatory fields: Customer Name and City.")
            return
        st.session_state.customer_profile = {
            "name": name.strip(),
            "birth_date": birth_date,
            "occupation": occupation.strip(),
            "address": address.strip(),
            "city": city.strip(),
            "contact_details": contact.strip(),
        }
        st.rerun()

    if cancel_clicked:
        st.rerun()


@st.dialog("Client Details For PDF", width="large")
def render_pdf_export_dialog(scenario: Scenario, analysis_results: dict[str, SimulationResult]) -> None:
    current = st.session_state.get("customer_profile") or {}
    st.caption("Review client details and confirm PDF export.")

    name = st.text_input("Customer Name *", value=str(current.get("name", "")), key="pdf_customer_name")
    birth_date = st.date_input(
        "Birth Date",
        value=current.get("birth_date", date.today()),
        format="DD-MM-YYYY",
        key="pdf_customer_birth_date",
    )
    selected_occupation = str(current.get("occupation", OCCUPATION_OPTIONS[0]))
    if selected_occupation not in OCCUPATION_OPTIONS:
        selected_occupation = OCCUPATION_OPTIONS[0]
    occupation = st.selectbox(
        "Occupation",
        options=OCCUPATION_OPTIONS,
        index=OCCUPATION_OPTIONS.index(selected_occupation),
        key="pdf_customer_occupation",
    )
    address = st.text_area("Address", value=str(current.get("address", "")), height=90, key="pdf_customer_address")
    city = st.text_input("City *", value=str(current.get("city", "")), key="pdf_customer_city")
    contact = st.text_input("Contact Details", value=str(current.get("contact_details", "")), key="pdf_customer_contact")

    export_col, cancel_col = st.columns(2)
    with export_col:
        export_clicked = st.button("Confirm & Prepare PDF", type="primary", use_container_width=True, key="pdf_export_confirm")
    with cancel_col:
        cancel_clicked = st.button("Cancel", use_container_width=True, key="pdf_export_cancel")

    if export_clicked:
        if not name.strip() or not city.strip():
            st.error("Please fill mandatory fields: Customer Name and City.")
            return
        profile = {
            "name": name.strip(),
            "birth_date": birth_date,
            "occupation": occupation.strip(),
            "address": address.strip(),
            "city": city.strip(),
            "contact_details": contact.strip(),
        }
        st.session_state.customer_profile = profile
        try:
            logo_path = next((path for path in LOGO_IMAGE_CANDIDATES if path.exists()), None)
            pdf_name, pdf_data = export_pdf_report_bytes(
                scenario,
                analysis_results,
                get_selected_analysis_variant(),
                customer_profile=profile,
                logo_path=logo_path,
                developer_name=DEVELOPER_NAME,
                developer_phone=DEVELOPER_PHONE,
                developer_email=DEVELOPER_EMAIL,
            )
            st.session_state.pdf_export_filename = pdf_name
            st.session_state.pdf_export_data = pdf_data
            st.session_state.pdf_export_success = f"PDF prepared: {pdf_name}"
            st.rerun()
        except RuntimeError as exc:
            st.error(str(exc))
        except Exception as exc:  # pragma: no cover
            st.error(f"PDF export failed: {exc}")

    if cancel_clicked:
        st.rerun()


@st.dialog("Client Details For Excel", width="large")
def render_excel_export_dialog(scenario: Scenario, analysis_results: dict[str, SimulationResult]) -> None:
    current = st.session_state.get("customer_profile") or {}
    st.caption("Review client details and confirm Excel export.")

    name = st.text_input("Customer Name *", value=str(current.get("name", "")), key="excel_customer_name")
    birth_date = st.date_input(
        "Birth Date",
        value=current.get("birth_date", date.today()),
        format="DD-MM-YYYY",
        key="excel_customer_birth_date",
    )
    selected_occupation = str(current.get("occupation", OCCUPATION_OPTIONS[0]))
    if selected_occupation not in OCCUPATION_OPTIONS:
        selected_occupation = OCCUPATION_OPTIONS[0]
    occupation = st.selectbox(
        "Occupation",
        options=OCCUPATION_OPTIONS,
        index=OCCUPATION_OPTIONS.index(selected_occupation),
        key="excel_customer_occupation",
    )
    address = st.text_area("Address", value=str(current.get("address", "")), height=90, key="excel_customer_address")
    city = st.text_input("City *", value=str(current.get("city", "")), key="excel_customer_city")
    contact = st.text_input("Contact Details", value=str(current.get("contact_details", "")), key="excel_customer_contact")

    export_col, cancel_col = st.columns(2)
    with export_col:
        export_clicked = st.button("Confirm & Prepare Excel", type="primary", use_container_width=True, key="excel_export_confirm")
    with cancel_col:
        cancel_clicked = st.button("Cancel", use_container_width=True, key="excel_export_cancel")

    if export_clicked:
        if not name.strip() or not city.strip():
            st.error("Please fill mandatory fields: Customer Name and City.")
            return
        profile = {
            "name": name.strip(),
            "birth_date": birth_date,
            "occupation": occupation.strip(),
            "address": address.strip(),
            "city": city.strip(),
            "contact_details": contact.strip(),
        }
        st.session_state.customer_profile = profile
        workbook_name, workbook_bytes = export_excel_report(
            Path.cwd(),
            scenario,
            analysis_results,
            get_selected_analysis_variant(),
            customer_profile=profile,
            developer_name=DEVELOPER_NAME,
            developer_phone=DEVELOPER_PHONE,
            developer_email=DEVELOPER_EMAIL,
        )
        st.session_state.excel_export_filename = workbook_name
        st.session_state.excel_export_data = workbook_bytes
        st.session_state.excel_export_success = f"Excel prepared: {workbook_name}"
        st.rerun()

    if cancel_clicked:
        st.rerun()


def render_developer_details() -> None:
    logo_col, details_col = st.columns([0.3, 0.7], gap="medium")
    logo_path = next((path for path in LOGO_IMAGE_CANDIDATES if path.exists()), None)
    with logo_col:
        st.markdown('<div class="brand-logo-card">', unsafe_allow_html=True)
        if logo_path is not None:
            st.image(str(logo_path), use_container_width=True)
        else:
            st.warning("Logo not found at the configured path.")
        st.markdown("</div>", unsafe_allow_html=True)
    with details_col:
        st.markdown(
            """
            <div class="dashboard-card">
                <div class="section-title">Developers Details</div>
                <div class="summary-grid">
                    <div class="summary-label">Developers</div><div class="summary-value">H&amp;L Stocktrades</div>
                    <div class="summary-label">Contact Details</div><div class="summary-value">Rishabh Hurkat, 8830488312</div>
                    <div class="summary-label">Email</div><div class="summary-value">hlstocktraders@gmail.com</div>
                </div>
            </div>
            """,
            unsafe_allow_html=True,
        )


def render_top_brand_bar() -> None:
    logo_path = next((path for path in LOGO_IMAGE_CANDIDATES if path.exists()), None)
    logo_data: bytes | None = None
    logo_display_height = int(LOGO_DISPLAY_WIDTH * (1 - LOGO_CROP_TOP_RATIO))
    info_box_height = int(logo_display_height * 0.70)
    info_box_offset = max(0, (logo_display_height - info_box_height) // 2)
    if logo_path is not None:
        try:
            with Image.open(logo_path) as logo:
                width, height = logo.size
                crop_top = int(height * LOGO_CROP_TOP_RATIO)
                cropped = logo.crop((0, crop_top, width, height))
                if width > 0:
                    logo_display_height = int(LOGO_DISPLAY_WIDTH * (cropped.height / width))
                    info_box_height = int(logo_display_height * 0.70)
                    info_box_offset = max(0, (logo_display_height - info_box_height) // 2)
                buffer = BytesIO()
                cropped.save(buffer, format="PNG")
                logo_data = buffer.getvalue()
        except Exception:
            logo_data = logo_path.read_bytes()

    if logo_data is None:
        st.markdown("**H&L Stocktrades**")
        return

    logo_b64 = base64.b64encode(logo_data).decode("ascii")
    st.markdown(
        f"""
        <div class="top-brand-row">
            <div class="top-brand-logo">
                <img src="data:image/png;base64,{logo_b64}" style="width:{LOGO_DISPLAY_WIDTH}px;" />
            </div>
            <div class="top-brand-wrap" style="margin-top:{info_box_offset}px;">
                <div class="top-brand-bar" style="height:{info_box_height}px;">
                    <div class="top-brand-contact">
                        <div class="top-brand-name">{DEVELOPER_NAME}</div>
                        <div class="top-brand-line"><span class="top-brand-icon">&#128222;</span><span>{DEVELOPER_PHONE}</span></div>
                        <div class="top-brand-line"><span class="top-brand-icon">&#9993;</span><span>{DEVELOPER_EMAIL}</span></div>
                    </div>
                </div>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def normalize_cash_flows_df(df: pd.DataFrame) -> pd.DataFrame:
    normalized = df.copy()
    if "Type" not in normalized.columns:
        normalized["Type"] = pd.Series(dtype="string")
    if "Date" not in normalized.columns:
        normalized["Date"] = pd.Series(dtype="datetime64[ns]")
    if "Amount" not in normalized.columns:
        normalized["Amount"] = pd.Series(dtype="float64")

    normalized = normalized[["Type", "Date", "Amount"]]
    normalized["Type"] = normalized["Type"].astype("string")
    normalized["Date"] = pd.to_datetime(normalized["Date"], errors="coerce")
    normalized["Amount"] = pd.to_numeric(normalized["Amount"], errors="coerce")
    return normalized


def format_inr_amount(amount: float) -> str:
    sign = "-" if amount < 0 else ""
    integer_text, decimal_text = f"{abs(amount):.2f}".split(".")
    if len(integer_text) <= 3:
        grouped = integer_text
    else:
        last_three = integer_text[-3:]
        remaining = integer_text[:-3]
        chunks: list[str] = []
        while len(remaining) > 2:
            chunks.append(remaining[-2:])
            remaining = remaining[:-2]
        if remaining:
            chunks.append(remaining)
        grouped = ",".join(reversed(chunks)) + f",{last_three}"
    return f"{sign}\u20b9 {grouped}.{decimal_text}"


def parse_currency_amount(raw_value: str) -> float | None:
    cleaned = raw_value.strip().replace("\u20b9", "").replace("INR", "").replace(",", "").strip()
    if not cleaned:
        return None
    try:
        return float(cleaned)
    except ValueError:
        return None


def normalize_currency_state(key: str, min_value: float) -> None:
    raw_value = str(st.session_state.get(key, "")).strip()
    parsed = parse_currency_amount(raw_value)
    if parsed is not None and parsed >= min_value:
        st.session_state[key] = format_inr_amount(parsed)


def currency_input(label: str, *, key: str, value: float, min_value: float = 0.0) -> float:
    if key not in st.session_state:
        st.session_state[key] = format_inr_amount(value)
    st.text_input(label, key=key, on_change=normalize_currency_state, args=(key, min_value))
    parsed = parse_currency_amount(str(st.session_state.get(key, "")))
    if parsed is None or parsed < min_value:
        st.caption(f"Enter amount like {format_inr_amount(max(min_value, 25000.0))}")
        return float(value)
    return float(parsed)


def cash_flow_display_df(df: pd.DataFrame) -> tuple[pd.DataFrame, list[int]]:
    normalized = normalize_cash_flows_df(df).dropna(subset=["Type", "Date", "Amount"]).copy()
    if normalized.empty:
        return pd.DataFrame(columns=["Type", "Date", "Amount"]), []
    normalized = normalized[normalized["Amount"] > 0]
    if normalized.empty:
        return pd.DataFrame(columns=["Type", "Date", "Amount"]), []
    row_index_map = normalized.index.tolist()
    normalized["Type"] = normalized["Type"].astype(str).str.title()
    normalized["Date"] = pd.to_datetime(normalized["Date"], errors="coerce").dt.strftime("%d-%m-%Y")
    normalized["Amount"] = normalized["Amount"].astype(float).apply(format_inr_amount)
    return normalized[["Type", "Date", "Amount"]], row_index_map


def month_offset_from_sip(sip_start: date, value: date) -> int:
    return max(0, ((value.year - sip_start.year) * 12) + (value.month - sip_start.month))


def portfolio_value_on_date(value: date) -> float:
    scenario, _ = build_scenario_from_state()
    if scenario.sip_start_date is None:
        return 0.0
    return projected_value_at_month(scenario, month_offset_from_sip(scenario.sip_start_date, value.replace(day=1)))


def sip_value_on_date(scenario: Scenario, month_index: int) -> float:
    sip_months = total_investment_months(scenario)
    if month_index < 0 or month_index >= sip_months:
        return 0.0
    sip_amount = scenario.monthly_sip
    if scenario.step_up_enabled and scenario.step_up_rate > 0:
        sip_amount *= (1 + (scenario.step_up_rate / 100)) ** (month_index // 12)
    return sip_amount


def swp_end_date(scenario: Scenario) -> date | None:
    if not scenario.swp_enabled:
        return sip_end_date(scenario)
    start = swp_start_date(scenario)
    if start is None:
        return sip_end_date(scenario)
    swp_months = max(1, scenario.swp_years * 12)
    return add_months(start, swp_months - 1)


def cash_flow_window_bounds(scenario: Scenario) -> tuple[date, date]:
    start = scenario.sip_start_date or date.today().replace(day=1)
    end = swp_end_date(scenario) or sip_end_date(scenario) or start
    if end < start:
        end = start
    return start, end


@st.dialog("Cash Flow Entry", width="small", dismissible=False)
def render_cash_flow_entry_dialog(cash_flow_start: date, cash_flow_end: date) -> None:
    entry_type = str(st.session_state.get("cf_editor_type", "Add")).title()
    st.markdown(f'<div class="section-title">Type: {entry_type}</div>', unsafe_allow_html=True)

    max_month_gap = max(0, ((cash_flow_end.year - cash_flow_start.year) * 12) + (cash_flow_end.month - cash_flow_start.month))
    max_years = max_month_gap // 12

    if "cf_editor_timing" not in st.session_state:
        st.session_state["cf_editor_timing"] = "specific_date"
    if "cf_editor_years" not in st.session_state:
        st.session_state["cf_editor_years"] = 0
    if "cf_editor_date" not in st.session_state:
        st.session_state["cf_editor_date"] = cash_flow_start
    if "cf_editor_amount" not in st.session_state:
        st.session_state["cf_editor_amount"] = format_inr_amount(10000.0)

    timing_mode = st.selectbox(
        "Entry Mode",
        options=["after_start_years", "specific_date"],
        format_func=lambda value: "Specific Year After Start" if value == "after_start_years" else "Specific Date",
        key="cf_editor_timing",
    )

    if timing_mode == "after_start_years":
        years_after = st.number_input(
            "Start After Years",
            min_value=0,
            max_value=int(max_years),
            step=1,
            key="cf_editor_years",
        )
        entry_date = add_months(cash_flow_start, int(years_after) * 12)
        st.text_input("Date", value=entry_date.strftime("%d-%m-%Y"), disabled=True)
    else:
        if st.session_state["cf_editor_date"] < cash_flow_start or st.session_state["cf_editor_date"] > cash_flow_end:
            st.session_state["cf_editor_date"] = cash_flow_start
        entry_date = st.date_input(
            "Date",
            min_value=cash_flow_start,
            max_value=cash_flow_end,
            format="DD-MM-YYYY",
            key="cf_editor_date",
        )

    st.caption(f"Portfolio Value on {entry_date.strftime('%d-%m-%Y')}: {format_currency(portfolio_value_on_date(entry_date))}")

    amount_default = parse_currency_amount(str(st.session_state.get("cf_editor_amount", format_inr_amount(10000.0))))
    amount_value = currency_input(
        "Amount",
        key="cf_editor_amount",
        value=float(amount_default if amount_default is not None else 10000.0),
        min_value=1.0,
    )

    confirm_col, cancel_col = st.columns(2)
    with confirm_col:
        confirm_clicked = st.button("Confirm", use_container_width=True, type="primary", key="cf_editor_confirm")
    with cancel_col:
        cancel_clicked = st.button("Cancel", use_container_width=True, key="cf_editor_cancel")

    if confirm_clicked:
        final_date = entry_date.replace(day=1)
        editor_mode = st.session_state.get("cf_editor_mode")
        if editor_mode == "edit" and st.session_state.get("cf_editor_row") is not None:
            target_row = int(st.session_state["cf_editor_row"])
            st.session_state.cash_flows_df.at[target_row, "Type"] = entry_type
            st.session_state.cash_flows_df.at[target_row, "Date"] = pd.Timestamp(final_date)
            st.session_state.cash_flows_df.at[target_row, "Amount"] = float(amount_value)
        else:
            new_row = pd.DataFrame([{"Type": entry_type, "Date": pd.Timestamp(final_date), "Amount": float(amount_value)}])
            st.session_state.cash_flows_df = pd.concat([st.session_state.cash_flows_df, new_row], ignore_index=True)
        st.session_state.cash_flows_df = normalize_cash_flows_df(st.session_state.cash_flows_df)
        st.session_state["cf_editor_open"] = False
        st.rerun()

    if cancel_clicked:
        st.session_state["cf_editor_open"] = False
        st.rerun()


def build_scenario_from_state() -> tuple[Scenario, list[str]]:
    raw = st.session_state.scenario_inputs
    errors: list[str] = []

    sip_start = raw["sip_start_date"].replace(day=1)
    years = int(raw["investment_years"])
    months = int(raw["investment_months"])
    if years == 0 and months == 0:
        months = 1
    swp_start_override_raw = raw.get("swp_start_date_override")
    swp_start_override = (
        pd.Timestamp(swp_start_override_raw).date() if swp_start_override_raw is not None else None
    )

    cash_flow_events: list[CashFlowEvent] = []
    if bool(raw.get("cash_flows_enabled", False)):
        df = st.session_state.cash_flows_df.copy()
        for _, row in df.iterrows():
            if pd.isna(row.get("Type")) or pd.isna(row.get("Date")) or pd.isna(row.get("Amount")):
                continue
            amount = float(row["Amount"])
            if amount <= 0:
                continue
            flow_date = pd.Timestamp(row["Date"]).date().replace(day=1)
            flow_type = "add" if str(row["Type"]).lower() == "add" else "withdraw"
            cash_flow_events.append(CashFlowEvent(flow_type=flow_type, event_date=flow_date, amount=amount))

    scenario = Scenario(
        sip_start_date=sip_start,
        monthly_sip=float(raw["monthly_sip"]),
        investment_years=years,
        investment_months=months,
        annual_roi=float(raw["annual_roi"]),
        inflation_rate=float(raw["inflation_rate"]),
        step_up_enabled=bool(raw["step_up_enabled"]),
        step_up_rate=float(raw["step_up_rate"]) if raw["step_up_enabled"] else 0.0,
        cash_flow_events=sorted(cash_flow_events, key=lambda event: event.event_date),
        swp_enabled=bool(raw["swp_enabled"]),
        swp_start_mode=str(raw["swp_start_mode"]),
        swp_start_year=int(raw["swp_start_year"]),
        swp_start_date_override=swp_start_override,
        swp_years=int(raw["swp_years"]),
        swp_months=0,
        monthly_swp_amount=float(raw["monthly_swp_amount"]) if raw["swp_enabled"] else 0.0,
    )

    if scenario.monthly_sip <= 0:
        errors.append("Monthly SIP must be greater than 0.")
    if scenario.annual_roi < 0:
        errors.append("Expected return cannot be negative.")
    if scenario.inflation_rate < 0:
        errors.append("Inflation cannot be negative.")
    if scenario.step_up_enabled and scenario.step_up_rate < 0:
        errors.append("Step-up rate cannot be negative.")

    cf_start, cf_end = cash_flow_window_bounds(scenario)
    for event in scenario.cash_flow_events:
        if not (cf_start <= event.event_date <= cf_end):
            errors.append(
                f"Cash flow on {event.event_date.strftime('%d-%m-%Y')} must be between "
                f"{cf_start.strftime('%d-%m-%Y')} and {cf_end.strftime('%d-%m-%Y')}."
            )

    if scenario.swp_enabled:
        if scenario.swp_start_mode == "after_start_years" and scenario.swp_start_year < 0:
            errors.append("SWP start year cannot be negative.")
        if scenario.swp_start_mode == "specific_date" and scenario.swp_start_date_override is None:
            errors.append("Please select a specific SWP start date.")
        if scenario.swp_years <= 0:
            errors.append("SWP tenure must be at least 1 year.")

    return scenario, errors


def render_summary_card(scenario: Scenario) -> None:
    cash_flow_summary = "No"
    if scenario.cash_flow_events:
        adds = sum(event.amount for event in scenario.cash_flow_events if event.flow_type == "add")
        withdrawals = sum(event.amount for event in scenario.cash_flow_events if event.flow_type == "withdraw")
        net = adds - withdrawals
        cash_flow_summary = f"{len(scenario.cash_flow_events)} entries | Net {format_currency(net)}"

    rows = [
        ("SIP Start", format_month_year(scenario.sip_start_date)),
        ("SIP End", format_month_year(sip_end_date(scenario))),
        ("Monthly SIP", format_currency(scenario.monthly_sip)),
        ("Tenure", format_tenure(scenario.investment_years, scenario.investment_months)),
        ("Return", format_percentage(scenario.annual_roi)),
        ("Inflation", format_percentage(scenario.inflation_rate)),
        ("Step-Up", f"{scenario.step_up_rate:.2f} %" if scenario.step_up_enabled else "No"),
        ("Cash Flows", cash_flow_summary),
    ]
    html = ['<div class="dashboard-card"><div class="section-title">Live Scenario</div><div class="summary-grid">']
    for label, value in rows:
        html.append(f'<div class="summary-label">{label}</div><div class="summary-value">{value}</div>')
    html.append("</div></div>")
    st.markdown("".join(html), unsafe_allow_html=True)


def render_swp_snapshot_card(scenario: Scenario, preview_result: SimulationResult | None = None) -> None:
    if scenario.swp_enabled:
        monthly_swp = preview_result.actual_monthly_swp if preview_result is not None else scenario.monthly_swp_amount
        swp_months = max(1, scenario.swp_years * 12)
        total_withdrawal = monthly_swp * swp_months
        remaining_fund = preview_result.final_portfolio_value if preview_result is not None else 0.0
        rows = [
            ("SWP Start", format_month_year(swp_start_date(scenario))),
            ("SWP End", format_month_year(swp_end_date(scenario))),
            ("SWP Tenure", format_tenure(scenario.swp_years, scenario.swp_months)),
            ("Monthly SWP", format_currency(monthly_swp)),
            ("Total Withdrawal", format_currency(total_withdrawal)),
            ("Remaining Fund", format_currency(remaining_fund)),
        ]
    else:
        rows = [
            ("SWP Status", "Not Enabled"),
            ("Default End (SIP)", format_month_year(sip_end_date(scenario))),
        ]

    html = ['<div class="dashboard-card"><div class="section-title">SWP Snapshot</div><div class="summary-grid">']
    for label, value in rows:
        html.append(f'<div class="summary-label">{label}</div><div class="summary-value">{value}</div>')
    html.append("</div></div>")
    st.markdown("".join(html), unsafe_allow_html=True)


def goal_based_amount_invested(result: SimulationResult) -> float:
    # Goal-based invested amount uses net cash flow across the full schedule.
    return float(sum((row.sip_amount - row.swp_amount + row.lumpsum_amount) for row in result.schedule_rows))


def render_invested_profit_donut(analysis_results: dict[str, SimulationResult] | None) -> None:
    st.markdown('<div class="section-title">Goal-Based Strategy Share</div>', unsafe_allow_html=True)
    selected_variant = get_selected_analysis_variant()
    if not analysis_results:
        st.caption("Run simulation to view share chart.")
        return
    result = analysis_results.get(selected_variant)
    if result is None:
        st.caption("Selected analysis variant is not available.")
        return

    total_invested = float(goal_based_amount_invested(result))
    total_profit = float(result.total_profit)
    final_portfolio = float(result.final_portfolio_value)

    if final_portfolio <= 0:
        st.caption("Final portfolio value is not positive, so share chart is unavailable.")
        return

    if total_invested < 0:
        invested_impact = 0.0
        profit_impact = abs(total_profit) if abs(total_profit) > 0 else abs(final_portfolio)
    else:
        invested_impact = abs(total_invested)
        profit_impact = abs(total_profit)
    total_impact = invested_impact + profit_impact
    if total_impact <= 0:
        st.caption("No impact values available to render donut chart.")
        return

    share_df = pd.DataFrame(
        {
            "Part": ["Invested", "Profit"],
            "Value": [invested_impact, profit_impact],
            "DisplayValue": [format_currency(invested_impact), format_currency(profit_impact)],
        }
    )

    st.vega_lite_chart(
        share_df,
        {
            "mark": {"type": "arc", "innerRadius": 62},
            "encoding": {
                "theta": {"field": "Value", "type": "quantitative"},
                "color": {
                    "field": "Part",
                    "type": "nominal",
                    "scale": {"range": ["#3a86ff", "#2ec4b6"]},
                    "legend": {"title": None, "orient": "bottom"},
                },
                "tooltip": [
                    {"field": "Part", "type": "nominal"},
                    {"field": "DisplayValue", "type": "nominal", "title": "Value"},
                ],
            },
            "view": {"stroke": None},
            "height": 250,
        },
        key=(
            f"goal_donut_{ANALYSIS_VARIANT_OPTIONS.index(selected_variant)}_"
            f"{round(total_invested, 2)}_{round(total_profit, 2)}_{round(final_portfolio, 2)}"
        ),
        use_container_width=True,
    )

    invested_pct = (invested_impact / total_impact) * 100
    profit_pct = (profit_impact / total_impact) * 100
    st.caption(
        f"Net Invested: {format_currency(total_invested)} ({invested_pct:.1f}% impact) | "
        f"Total Profit: {format_currency(total_profit)} ({profit_pct:.1f}% impact)"
    )


def result_cards(result: SimulationResult) -> None:
    goal_invested = goal_based_amount_invested(result)
    metrics = [
        ("Total Invested", goal_invested),
        ("Final Portfolio", float(result.final_portfolio_value)),
        ("Total Profit", float(result.total_profit)),
    ]

    def value_class(label: str, value: float) -> str:
        # For Final Portfolio and Total Profit: positive green, negative red.
        if label in ("Final Portfolio", "Total Profit"):
            if value > 0:
                return "result-negative"
            if value < 0:
                return "result-positive"
            return "result-neutral"
        # Keep existing visual rule for Total Invested.
        if value < 0:
            return "result-negative"
        if value > 0:
            return "result-positive"
        return "result-neutral"

    html = ['<div class="result-grid">']
    for label, numeric_value in metrics:
        html.append(
            f'<div class="result-item">'
            f'<div class="result-label">{label}</div>'
            f'<div class="result-value {value_class(label, numeric_value)}">{format_currency(numeric_value)}</div>'
            f"</div>"
        )
    html.append("</div>")
    st.markdown("".join(html), unsafe_allow_html=True)


def schedule_columns_for_mode(mode_label: str | None) -> list[str]:
    if mode_label == SIP_ONLY_MODE:
        return ["#", "Date", "Type", "SIP", "Net CF", "Growth", "Value (Close)", "Close Value Raw"]
    if mode_label == SIP_CF_MODE:
        return ["#", "Date", "Type", "SIP", "Investment CashFlows", "Net CF", "Growth", "Value (Close)", "Close Value Raw"]
    if mode_label == SIP_SWP_MODE:
        return ["#", "Date", "Type", "SIP", "SWP", "Net CF", "Growth", "Value (Close)", "Close Value Raw"]
    return [
        "#",
        "Date",
        "Type",
        "SIP",
        "SWP",
        "Investment CashFlows",
        "Net CF",
        "Growth",
        "Value (Close)",
        "Close Value Raw",
    ]


def schedule_dataframe(result: SimulationResult, mode_label: str | None = None) -> pd.DataFrame:
    rows = []
    for row in result.schedule_rows:
        net_cf = row.sip_amount - row.swp_amount + row.lumpsum_amount
        rows.append(
            {
                "#": row.period_number,
                "Date": format_month_year(row.period_date),
                "Type": row.phase,
                "SIP": format_currency(row.sip_amount),
                "SWP": format_currency(row.swp_amount),
                "Investment CashFlows": format_currency(row.lumpsum_amount),
                "Net CF": format_currency(net_cf),
                "Growth": format_currency(row.growth),
                "Value (Close)": format_currency(row.closing_balance),
                "Close Value Raw": row.closing_balance,
            }
        )
    df = pd.DataFrame(rows)
    if df.empty:
        return df
    selected_columns = [column for column in schedule_columns_for_mode(mode_label) if column in df.columns]
    return df[selected_columns]


def chart_dataframe(result: SimulationResult) -> pd.DataFrame:
    return pd.DataFrame(
        {
            "Date": [row.period_date for row in result.schedule_rows if row.period_date is not None],
            "Portfolio Value": [row.closing_balance for row in result.schedule_rows if row.period_date is not None],
        }
    ).set_index("Date")


def scenario_variant(base: Scenario, **overrides: object) -> Scenario:
    patched = dict(overrides)
    if "cash_flow_events" not in patched:
        patched["cash_flow_events"] = list(base.cash_flow_events)
    return replace(base, **patched)


def build_analysis_scenarios(base: Scenario) -> dict[str, Scenario]:
    return {
        SIP_ONLY_MODE: scenario_variant(
            base,
            cash_flow_events=[],
            swp_enabled=False,
            swp_months=0,
            monthly_swp_amount=0.0,
        ),
        SIP_CF_MODE: scenario_variant(
            base,
            swp_enabled=False,
            swp_months=0,
            monthly_swp_amount=0.0,
        ),
        SIP_SWP_MODE: scenario_variant(
            base,
            cash_flow_events=[],
        ),
        ALL_COMBINATION_MODE: scenario_variant(base),
    }


def compute_analysis_results(base: Scenario) -> dict[str, SimulationResult]:
    scenarios = build_analysis_scenarios(base)
    return {label: run_simulation(scenario_obj) for label, scenario_obj in scenarios.items()}


def timeline_index_for_chart(scenario: Scenario) -> pd.DatetimeIndex:
    start = scenario.sip_start_date
    if start is None:
        return pd.DatetimeIndex([])
    end = swp_end_date(scenario) if scenario.swp_enabled else sip_end_date(scenario)
    if end is None:
        end = start
    if end < start:
        end = start
    return pd.date_range(start=start, end=end, freq="MS")


def result_series_on_timeline(result: SimulationResult, timeline: pd.DatetimeIndex, label: str) -> pd.Series:
    points = {
        pd.Timestamp(row.period_date): row.closing_balance
        for row in result.schedule_rows
        if row.period_date is not None
    }
    if not points:
        return pd.Series(index=timeline, data=[0.0] * len(timeline), name=label)
    series = pd.Series(points, name=label).sort_index()
    if timeline.empty:
        return series
    aligned = series.reindex(timeline).ffill().fillna(0.0)
    aligned.name = label
    return aligned


def comparison_metrics_card(analysis_results: dict[str, SimulationResult]) -> None:
    selected_variant = render_analysis_variant_buttons("scenario_comparison")

    def totals_from_result(result: SimulationResult) -> dict[str, float]:
        total_sip = sum(row.sip_amount for row in result.schedule_rows)
        total_swp = sum(abs(row.swp_amount) for row in result.schedule_rows)
        total_investment_cashflows = sum(row.lumpsum_amount for row in result.schedule_rows)
        net_cf = sum((row.sip_amount - row.swp_amount + row.lumpsum_amount) for row in result.schedule_rows)
        return {
            "total_sip": float(total_sip),
            "total_swp": float(total_swp),
            "total_investment_cashflows": float(total_investment_cashflows),
            "net_cf": float(net_cf),
        }

    sip_only_result = analysis_results[SIP_ONLY_MODE]
    sip_cf_result = analysis_results[SIP_CF_MODE]
    sip_swp_result = analysis_results[SIP_SWP_MODE]
    all_combo_result = analysis_results[ALL_COMBINATION_MODE]

    sip_totals = totals_from_result(sip_only_result)
    cf_totals = totals_from_result(sip_cf_result)
    swp_totals = totals_from_result(sip_swp_result)
    combo_totals = totals_from_result(all_combo_result)

    scenario_values: dict[str, dict[str, float]] = {
        SIP_ONLY_MODE: {
            "Final Portfolio": sip_only_result.final_portfolio_value,
            "Amount Invested": sip_totals["total_sip"],
            "Total Profit": sip_only_result.total_profit,
        },
        SIP_CF_MODE: {
            "Final Portfolio": sip_cf_result.final_portfolio_value,
            "Amount Invested": cf_totals["total_sip"] + cf_totals["total_investment_cashflows"],
            "Total Profit": sip_cf_result.total_profit,
        },
        SIP_SWP_MODE: {
            "Final Portfolio": sip_swp_result.final_portfolio_value,
            "Amount Invested": swp_totals["total_sip"] - swp_totals["total_swp"],
            "Total Profit": sip_swp_result.total_profit,
        },
        ALL_COMBINATION_MODE: {
            "Final Portfolio": all_combo_result.final_portfolio_value,
            "Amount Invested": combo_totals["net_cf"],
            "Total Profit": all_combo_result.total_profit,
        },
    }

    compare_against = SIP_ONLY_MODE
    if selected_variant not in scenario_values:
        selected_variant = ALL_COMBINATION_MODE
    scenario_order = [compare_against] if selected_variant == compare_against else [compare_against, selected_variant]

    change_map: dict[str, float] = {}
    if selected_variant != compare_against:
        change_map[selected_variant] = (
            scenario_values[selected_variant]["Final Portfolio"] - scenario_values[compare_against]["Final Portfolio"]
        )

    comparison_rows: list[dict[str, str]] = []
    for metric_name in ("Final Portfolio", "Amount Invested", "Total Profit"):
        row: dict[str, str] = {"Metric": metric_name}
        for scenario_name in scenario_order:
            row[scenario_name] = format_currency(scenario_values[scenario_name][metric_name])
        comparison_rows.append(row)

    change_row: dict[str, str] = {"Metric": "Change (Compared with SIP Only)"}
    for scenario_name in scenario_order:
        if scenario_name == compare_against:
            change_row[scenario_name] = "-"
        else:
            change_row[scenario_name] = format_currency(change_map.get(scenario_name, 0.0))
    comparison_rows.append(change_row)

    comparison_df = pd.DataFrame(comparison_rows)

    def style_change_row(row: pd.Series) -> list[str]:
        if str(row.get("Metric", "")) != "Change (Compared with SIP Only)":
            return [""] * len(row)
        styles = ["font-weight: 800; color: #103b52;"]
        for scenario_name in scenario_order:
            if scenario_name == compare_against:
                styles.append("font-weight: 700; color: #495057;")
            else:
                delta = change_map.get(scenario_name, 0.0)
                color = "#107c41" if delta > 0 else "#c92a2a" if delta < 0 else "#495057"
                styles.append(f"font-weight: 800; color: {color};")
        return styles

    styled_comparison = comparison_df.style.apply(style_change_row, axis=1)

    chart_rows: list[dict[str, object]] = []
    for metric_name in ("Final Portfolio", "Amount Invested", "Total Profit"):
        for scenario_name in scenario_order:
            chart_rows.append(
                {
                    "Metric": metric_name,
                    "Scenario": scenario_name,
                    "Value": float(scenario_values[scenario_name][metric_name]),
                }
            )
    chart_df = pd.DataFrame(chart_rows)

    if len(scenario_order) == 1:
        color_domain = scenario_order
        color_range = ["#3a86ff"]
    else:
        color_domain = scenario_order
        color_range = ["#3a86ff", "#ff9f1c"]

    table_col, chart_col = st.columns([3, 2], gap="large")

    with table_col:
        st.dataframe(
            styled_comparison,
            use_container_width=True,
            hide_index=True,
            row_height=80,
            height=420,
        )

    with chart_col:
        st.vega_lite_chart(
            chart_df,
            {
                "mark": {
                    "type": "bar",
                    "size": 18,
                    "cornerRadiusTopLeft": 6,
                    "cornerRadiusTopRight": 6,
                },
                "encoding": {
                    "x": {
                        "field": "Metric",
                        "type": "nominal",
                        "axis": {"title": None, "labelAngle": 0},
                        "scale": {"paddingInner": 0.4, "paddingOuter": 0.2},
                    },
                    "xOffset": {"field": "Scenario"},
                    "y": {"field": "Value", "type": "quantitative", "axis": {"title": None}},
                    "color": {
                        "field": "Scenario",
                        "type": "nominal",
                        "scale": {"domain": color_domain, "range": color_range},
                        "legend": {"title": None, "orient": "bottom"},
                    },
                    "tooltip": [
                        {"field": "Metric", "type": "nominal"},
                        {"field": "Scenario", "type": "nominal"},
                        {"field": "Value", "type": "quantitative", "format": ",.2f"},
                    ],
                },
                "height": 420,
                "view": {"stroke": None},
            },
            use_container_width=True,
        )


def render_builder() -> tuple[Scenario, list[str]]:
    st.subheader("Scenario Builder")

    inputs = st.session_state.scenario_inputs

    col1, col2 = st.columns(2)
    with col1:
        st.markdown('<div class="field-group-title">SIP & Tenure</div>', unsafe_allow_html=True)
        inputs["sip_start_date"] = st.date_input(
            "SIP Start Month",
            value=inputs["sip_start_date"],
            min_value=date(1900, 1, 1),
            format="DD-MM-YYYY",
        )
        inputs["monthly_sip"] = currency_input(
            "Monthly SIP",
            key="monthly_sip_input",
            value=float(inputs["monthly_sip"]),
            min_value=1.0,
        )
        tenure_left, tenure_right = st.columns(2)
        with tenure_left:
            inputs["investment_years"] = st.number_input(
                "Investment Years", min_value=0, value=int(inputs["investment_years"]), step=1
            )
        with tenure_right:
            inputs["investment_months"] = st.number_input(
                "Investment Months", min_value=0, max_value=11, value=int(inputs["investment_months"]), step=1
            )
        inputs["step_up_enabled"] = st.toggle("Annual Step-Up", value=bool(inputs["step_up_enabled"]))
        if inputs["step_up_enabled"]:
            inputs["step_up_rate"] = st.number_input("Step-Up Rate %", min_value=0.0, value=float(inputs["step_up_rate"]), step=0.5)
    with col2:
        st.markdown('<div class="field-group-title">Return</div>', unsafe_allow_html=True)
        inputs["annual_roi"] = st.number_input("Expected Return %", min_value=0.0, value=float(inputs["annual_roi"]), step=0.5)
        inputs["inflation_rate"] = st.number_input("Inflation %", min_value=0.0, value=float(inputs["inflation_rate"]), step=0.5)

    st.markdown('<div class="section-title">Investment Cash Flows</div>', unsafe_allow_html=True)
    inputs["cash_flows_enabled"] = st.toggle(
        "Enable",
        value=bool(inputs.get("cash_flows_enabled", False)),
    )
    if inputs["cash_flows_enabled"]:
        preview_for_bounds, _ = build_scenario_from_state()
        cash_flow_start, cash_flow_end = cash_flow_window_bounds(preview_for_bounds)
        st.caption("Use actions below. Date can be chosen by years-after-start or specific date.")
        st.caption(f"Allowed range: {cash_flow_start.strftime('%d-%m-%Y')} to {cash_flow_end.strftime('%d-%m-%Y')}")

        display_df, row_index_map = cash_flow_display_df(st.session_state.cash_flows_df)
        table_state = st.dataframe(
            display_df,
            use_container_width=True,
            hide_index=True,
            selection_mode="single-row",
            on_select="rerun",
            key="cash_flow_table",
        )
        st.caption("Select one row from the table, then use Edit or Delete.")
        selected_ui_row = table_state.selection.rows[0] if table_state.selection.rows else None
        selected_df_row = row_index_map[selected_ui_row] if selected_ui_row is not None and selected_ui_row < len(row_index_map) else None

        cash_flow_actions = st.columns(5)
        with cash_flow_actions[0]:
            if st.button("Add", use_container_width=True, key="cf_btn_add"):
                st.session_state["cf_editor_open"] = True
                st.session_state["cf_editor_mode"] = "add"
                st.session_state["cf_editor_row"] = None
                st.session_state["cf_editor_type"] = "Add"
                st.session_state["cf_editor_timing"] = "specific_date"
                st.session_state["cf_editor_years"] = 0
                st.session_state["cf_editor_date"] = cash_flow_start
                st.session_state["cf_editor_amount"] = format_inr_amount(10000.0)
                st.rerun()
        with cash_flow_actions[1]:
            if st.button("Withdraw", use_container_width=True, key="cf_btn_withdraw"):
                st.session_state["cf_editor_open"] = True
                st.session_state["cf_editor_mode"] = "add"
                st.session_state["cf_editor_row"] = None
                st.session_state["cf_editor_type"] = "Withdraw"
                st.session_state["cf_editor_timing"] = "specific_date"
                st.session_state["cf_editor_years"] = 0
                st.session_state["cf_editor_date"] = cash_flow_start
                st.session_state["cf_editor_amount"] = format_inr_amount(10000.0)
                st.rerun()
        with cash_flow_actions[2]:
            edit_disabled = selected_df_row is None
            if st.button("Edit", use_container_width=True, disabled=edit_disabled, key="cf_btn_edit"):
                selected_row = st.session_state.cash_flows_df.loc[selected_df_row]
                selected_date = pd.Timestamp(selected_row["Date"]).date().replace(day=1)
                months_since_start = month_offset_from_sip(cash_flow_start, selected_date)
                if months_since_start % 12 == 0:
                    timing_mode = "after_start_years"
                    years_after = months_since_start // 12
                else:
                    timing_mode = "specific_date"
                    years_after = 0
                st.session_state["cf_editor_open"] = True
                st.session_state["cf_editor_mode"] = "edit"
                st.session_state["cf_editor_row"] = int(selected_df_row)
                st.session_state["cf_editor_type"] = str(selected_row["Type"]).title()
                st.session_state["cf_editor_timing"] = timing_mode
                st.session_state["cf_editor_years"] = int(years_after)
                st.session_state["cf_editor_date"] = selected_date
                st.session_state["cf_editor_amount"] = format_inr_amount(float(selected_row["Amount"]))
                st.rerun()
        with cash_flow_actions[3]:
            delete_disabled = selected_df_row is None
            if st.button("Delete", use_container_width=True, disabled=delete_disabled, key="cf_btn_delete"):
                if selected_df_row is not None:
                    st.session_state.cash_flows_df = normalize_cash_flows_df(
                        st.session_state.cash_flows_df.drop(index=selected_df_row).reset_index(drop=True)
                    )
                    st.rerun()
        with cash_flow_actions[4]:
            if st.button("Clear All", use_container_width=True, key="cf_btn_clear"):
                st.session_state.cash_flows_df = default_cash_flows()
                st.rerun()

        if st.session_state.get("cf_editor_open", False):
            render_cash_flow_entry_dialog(cash_flow_start, cash_flow_end)

    st.markdown('<div class="section-title">Systematic Withdrawal Plan</div>', unsafe_allow_html=True)
    inputs["swp_enabled"] = st.toggle("Enable SWP", value=bool(inputs["swp_enabled"]))
    if inputs["swp_enabled"]:
        swp_type_col, swp_start_col = st.columns(2)
        with swp_type_col:
            mode_options = ["after_sip", "after_start_years", "specific_date"]
            if inputs.get("swp_start_mode") not in mode_options:
                inputs["swp_start_mode"] = "after_sip"
            inputs["swp_start_mode"] = st.selectbox(
                "SWP Type",
                options=mode_options,
                format_func=lambda value: {
                    "after_sip": "After SIP End",
                    "after_start_years": "Specific Year After Start",
                    "specific_date": "Specific Date",
                }[value],
                index=mode_options.index(inputs["swp_start_mode"]),
            )
        with swp_start_col:
            if inputs["swp_start_mode"] == "after_start_years":
                inputs["swp_start_year"] = st.number_input(
                    "Start After Years",
                    min_value=0,
                    value=int(inputs["swp_start_year"]),
                    step=1,
                )
            elif inputs["swp_start_mode"] == "after_sip":
                sip_months = (int(inputs["investment_years"]) * 12) + int(inputs["investment_months"])
                sip_months = max(1, sip_months)
                sip_end = add_months(
                    inputs["sip_start_date"].replace(day=1),
                    sip_months,
                )
                st.text_input("SWP Start Date", value=sip_end.strftime("%d-%m-%Y"), disabled=True)
            else:
                specific_default = inputs.get("swp_start_date_override") or inputs["sip_start_date"]
                inputs["swp_start_date_override"] = st.date_input(
                    "SWP Start Date",
                    value=specific_default,
                    min_value=date(1900, 1, 1),
                    format="DD-MM-YYYY",
                )

        preview_scenario, preview_errors = build_scenario_from_state()
        max_swp = maximum_monthly_swp(preview_scenario) if not preview_errors else 0.0
        start_month = swp_start_month(preview_scenario) or 0
        swp_value = projected_value_at_month(preview_scenario, start_month - 1)

        swp_bottom_left, swp_bottom_right = st.columns(2)
        with swp_bottom_left:
            inputs["swp_years"] = st.number_input(
                "SWP Tenure (Years)",
                min_value=1,
                value=max(1, int(inputs["swp_years"])),
                step=1,
            )
            default_swp = min(float(inputs["monthly_swp_amount"]), max_swp) if max_swp > 0 else float(inputs["monthly_swp_amount"])
            inputs["monthly_swp_amount"] = currency_input(
                "Monthly SWP",
                key="monthly_swp_input",
                value=float(default_swp),
                min_value=0.0,
            )
            if max_swp > 0 and float(inputs["monthly_swp_amount"]) > max_swp:
                inputs["monthly_swp_amount"] = float(max_swp)
                st.caption(f"Monthly SWP capped at {format_currency(max_swp)} based on current scenario.")
        with swp_bottom_right:
            st.caption(f"Expected Fund Value: {format_currency(swp_value)}")
            st.caption(f"Max Monthly SWP: {format_currency(max_swp)}")
            st.caption(f"Current SIP Value: {format_inr_amount(sip_value_on_date(preview_scenario, start_month))}")

    scenario, errors = build_scenario_from_state()
    return scenario, errors


def render_results(scenario: Scenario, analysis_results: dict[str, SimulationResult]) -> None:
    selected_variant = get_selected_analysis_variant()
    selected_result = analysis_results.get(selected_variant, analysis_results[ALL_COMBINATION_MODE])

    st.subheader(f"Simulation Results ({selected_variant})")
    result_cards(selected_result)

    st.markdown('<div class="dashboard-card">', unsafe_allow_html=True)
    show_scenario_comparison = render_collapsible_header(
        "Scenario Comparison",
        "show_scenario_comparison",
        "toggle_scenario_comparison",
    )
    if show_scenario_comparison:
        comparison_metrics_card(analysis_results)
    else:
        st.caption("Section hidden. Click + to view.")
    st.markdown("</div>", unsafe_allow_html=True)

    st.markdown('<div class="dashboard-card">', unsafe_allow_html=True)
    show_portfolio_growth = render_collapsible_header(
        "Portfolio Growth",
        "show_portfolio_growth",
        "toggle_portfolio_growth",
    )
    if show_portfolio_growth:
        selected_variant = render_analysis_variant_buttons("portfolio_growth")
        growth_result = analysis_results.get(selected_variant, analysis_results[ALL_COMBINATION_MODE])
        growth_rows = [row for row in growth_result.schedule_rows if row.period_date is not None]
        if not growth_rows:
            st.caption("No timeline data available for selected analysis variant.")
        else:
            chart_df = pd.DataFrame(
                {
                    "Date": [pd.Timestamp(row.period_date) for row in growth_rows],
                    "Portfolio Value": [float(row.closing_balance) for row in growth_rows],
                    "Cumulative Net Cash Flow": pd.Series(
                        [float(row.sip_amount - row.swp_amount + row.lumpsum_amount) for row in growth_rows]
                    ).cumsum(),
                }
            )
            chart_long = chart_df.melt(
                id_vars=["Date"],
                value_vars=["Cumulative Net Cash Flow", "Portfolio Value"],
                var_name="Line",
                value_name="Value",
            )
            st.vega_lite_chart(
                chart_long,
                {
                    "mark": {"type": "line", "strokeWidth": 3},
                    "encoding": {
                        "x": {"field": "Date", "type": "temporal", "axis": {"title": None}},
                        "y": {"field": "Value", "type": "quantitative", "axis": {"title": None}},
                        "color": {
                            "field": "Line",
                            "type": "nominal",
                            "scale": {
                                "domain": ["Cumulative Net Cash Flow", "Portfolio Value"],
                                "range": ["#ffd60a", "#1d4ed8"],
                            },
                            "legend": {"title": None},
                        },
                        "tooltip": [
                            {"field": "Date", "type": "temporal", "title": "Date"},
                            {"field": "Line", "type": "nominal"},
                            {"field": "Value", "type": "quantitative", "format": ",.2f"},
                        ],
                    },
                    "height": 320,
                    "view": {"stroke": None},
                },
                key=(
                    f"portfolio_growth_{selected_variant}_"
                    f"{len(growth_rows)}_"
                    f"{round(float(growth_result.final_portfolio_value), 2)}_"
                    f"{round(float(goal_based_amount_invested(growth_result)), 2)}"
                ),
                use_container_width=True,
            )
    else:
        st.caption("Section hidden. Click + to view.")
    st.markdown("</div>", unsafe_allow_html=True)

    selected_variant = get_selected_analysis_variant()
    selected_result = analysis_results.get(selected_variant, analysis_results[ALL_COMBINATION_MODE])
    schedule_df = schedule_dataframe(selected_result, selected_variant)
    st.markdown('<div class="dashboard-card">', unsafe_allow_html=True)
    st.markdown('<div class="section-title">Cash Flow Schedule</div>', unsafe_allow_html=True)
    st.dataframe(schedule_df.drop(columns=["Close Value Raw"]), use_container_width=True, hide_index=True)
    st.markdown("</div>", unsafe_allow_html=True)

    profile_for_export = st.session_state.get("customer_profile")
    export_col1, export_col2 = st.columns(2)
    with export_col1:
        if st.button("Generate Excel Report", type="primary", use_container_width=True, key="open_excel_export_dialog_btn"):
            render_excel_export_dialog(scenario, analysis_results)
    with export_col2:
        if st.button("Generate PDF Report", type="primary", use_container_width=True, key="open_pdf_export_dialog_btn"):
            render_pdf_export_dialog(scenario, analysis_results)

    download_col1, download_col2 = st.columns(2)
    with download_col1:
        excel_ready = bool(st.session_state.get("excel_export_data"))
        if excel_ready:
            st.download_button(
                "Download Prepared Excel",
                data=st.session_state["excel_export_data"],
                file_name=st.session_state.get("excel_export_filename", "MutualFundPortfolioSimulator_FullReport.xlsx"),
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="secondary",
                use_container_width=True,
                key="download_excel_report_btn_prepared",
            )
            if st.session_state.get("excel_export_success"):
                st.caption(str(st.session_state["excel_export_success"]))
    with download_col2:
        pdf_ready = bool(st.session_state.get("pdf_export_data"))
        if pdf_ready:
            st.download_button(
                "Download Prepared PDF",
                data=st.session_state["pdf_export_data"],
                file_name=st.session_state.get("pdf_export_filename", "MutualFundPortfolioSimulator_FullReport.pdf"),
                mime="application/pdf",
                type="secondary",
                use_container_width=True,
                key="download_pdf_report_btn_prepared",
            )
            if st.session_state.get("pdf_export_success"):
                st.caption(str(st.session_state["pdf_export_success"]))


def main() -> None:
    apply_theme()
    initialize_state()

    render_top_brand_bar()
    st.title("Mutual Fund Portfolio Simulator")
    mode_col, reset_col = st.columns([0.82, 0.18], gap="small")
    with mode_col:
        selected_mode = st.radio(
            "Select Analysis Mode",
            options=["Assumed Returns Basis", "Actual NAV Basis"],
            horizontal=True,
            key="analysis_mode",
        )
    with reset_col:
        st.markdown("<div style='height: 1.9rem;'></div>", unsafe_allow_html=True)
        reset_clicked = st.button("Reset", use_container_width=True, key="top_reset_btn")

    if reset_clicked:
        for key in (
            "scenario_inputs",
            "cash_flows_df",
            "last_result",
            "last_scenario",
            "last_analysis_results",
            "monthly_sip_input",
            "monthly_swp_input",
            "cf_editor_open",
            "cf_editor_mode",
            "cf_editor_row",
            "cf_editor_type",
            "cf_editor_timing",
            "cf_editor_years",
            "cf_editor_date",
            "cf_editor_amount",
            "cash_flow_table",
            "customer_profile",
            "pdf_export_success",
            "pdf_export_filename",
            "pdf_export_data",
            "pdf_customer_name",
            "pdf_customer_birth_date",
            "pdf_customer_occupation",
            "pdf_customer_address",
            "pdf_customer_city",
            "pdf_customer_contact",
            "excel_export_success",
            "excel_export_filename",
            "excel_export_data",
            "excel_customer_name",
            "excel_customer_birth_date",
            "excel_customer_occupation",
            "excel_customer_address",
            "excel_customer_city",
            "excel_customer_contact",
            "selected_analysis_variant",
            "show_scenario_comparison",
            "show_portfolio_growth",
        ):
            st.session_state.pop(key, None)
        st.rerun()

    if selected_mode == "Actual NAV Basis":
        if not NAV_DATA_FILE.exists():
            st.error(
                "NAV data file not found. Expected relative path: "
                f"{NAV_DATA_FILE.relative_to(APP_ROOT)}"
            )
            return
        try:
            nav_df = load_nav_data(NAV_DATA_FILE)
        except Exception as exc:
            st.error(f"Failed to load NAV data: {exc}")
            return
        st.info("Actual NAV Basis is ready for deployment pathing. Strategy logic implementation is pending.")
        st.caption(f"NAV file: {NAV_DATA_FILE.relative_to(APP_ROOT)}")
        st.caption(f"NAV rows loaded: {len(nav_df)}")
        return

    left, right = st.columns([1.25, 0.75], gap="large")

    with left:
        scenario, errors = render_builder()

        if errors:
            for error in errors:
                st.error(error)
            st.session_state.last_result = None
            st.session_state.last_scenario = None
            st.session_state.last_analysis_results = None
        else:
            analysis_results = compute_analysis_results(scenario)
            selected_variant = get_selected_analysis_variant()
            result = analysis_results.get(selected_variant, analysis_results[ALL_COMBINATION_MODE])
            st.session_state.last_scenario = scenario
            st.session_state.last_result = result
            st.session_state.last_analysis_results = analysis_results

    preview_analysis_results: dict[str, SimulationResult] | None = st.session_state.last_analysis_results
    preview_result_for_cards: SimulationResult | None = None
    if preview_analysis_results:
        selected_variant = get_selected_analysis_variant()
        preview_result_for_cards = preview_analysis_results.get(
            selected_variant,
            preview_analysis_results[ALL_COMBINATION_MODE],
        )

    with right:
        render_summary_card(scenario)
        render_swp_snapshot_card(scenario, preview_result_for_cards)
        render_invested_profit_donut(preview_analysis_results)

    saved_scenario = st.session_state.last_scenario
    saved_analysis_results = st.session_state.last_analysis_results
    if saved_scenario is not None and saved_analysis_results is not None:
        st.divider()
        render_results(saved_scenario, saved_analysis_results)


main()
    
