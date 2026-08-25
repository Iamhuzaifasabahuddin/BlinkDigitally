import calendar
import io
import logging
from datetime import datetime
from io import BytesIO
from itertools import zip_longest

import gspread
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import pytz
import streamlit as st
from google.oauth2.service_account import Credentials
from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak
from reportlab.platypus.flowables import HRFlowable

st.set_page_config(page_title="Blink Digitally", page_icon="📊", layout="centered")

logging.basicConfig(level=logging.INFO,
                    format='%(funcName)s --> %(message)s : %(asctime)s - %(levelname)s',
                    datefmt="%d-%m-%Y %I:%M:%S %p")

# ---------------------------------------------------------------------------
# Google Sheets setup
# ---------------------------------------------------------------------------
creds_dict = {
    "type": st.secrets["connections"]["gsheets"]["type"],
    "project_id": st.secrets["connections"]["gsheets"]["project_id"],
    "private_key_id": st.secrets["connections"]["gsheets"]["private_key_id"],
    "private_key": st.secrets["connections"]["gsheets"]["private_key"].replace("\\n", "\n"),
    "client_email": st.secrets["connections"]["gsheets"]["client_email"],
    "client_id": st.secrets["connections"]["gsheets"]["client_id"],
    "auth_uri": st.secrets["connections"]["gsheets"]["auth_uri"],
    "token_uri": st.secrets["connections"]["gsheets"]["token_uri"],
    "auth_provider_x509_cert_url": st.secrets["connections"]["gsheets"]["auth_provider_x509_cert_url"],
    "client_x509_cert_url": st.secrets["connections"]["gsheets"]["client_x509_cert_url"]
}
SCOPES = ["https://www.googleapis.com/auth/spreadsheets.readonly"]
SPREADSHEET_ID = st.secrets["connections"]["gsheets"]["SPREADSHEET_ID"]

# Sheet names
sheet_usa = "USA"
sheet_uk = "UK"
sheet_audio = "AudioBook"
sheet_printing = "Printing"
sheet_copyright = "Copyright"
sheet_a_plus = "A_plus"
sheet_sales = "Sales"
sheet_nielsen = "Nielsen ISBN"
sheet_chargeback = "Chargeback"

# ---------------------------------------------------------------------------
# Constants
# ---------------------------------------------------------------------------
DATE_FORMAT = "%d-%B-%Y"
DATE_COLUMNS = ["Publishing Date", "Last Edit (Revision)", "Trustpilot Review Date"]
PRINTING_DATE_COLUMNS = ["Order Date", "Shipping Date", "Fulfilled"]
PLATFORMS = ["Amazon", "Barnes & Noble", "Ingram Spark", "Draft2Digital", "Kobo", "LULU", "FAV", "ACX"]
USA_BRANDS = ["BookMarketeers", "Writers Clique", "KDP", "Aurora Writers"]
USA_PRINTING_BRANDS = ["BookMarketeers", "Writers Clique", "Aurora Writers", "KDP"]
UK_BRANDS = ["Authors Solution", "Book Publication", "Books Publisher"]
REVIEW_BRANDS = ["BookMarketeers", "Writers Clique", "Aurora Writers",
                 "Authors Solution", "Book Publication", "Books Publisher"]
ALLOWED_BRANDS = ["BookMarketeers", "Writers Clique", "Aurora Writers",
                  "Authors Solution", "Book Publication"]
COPYRIGHT_RATES = {"USA": 65, "Canada": 46, "UK": 42}
REVIEW_DETAIL_COLUMNS = ["Project Manager", "Name", "Brand", "Trustpilot Review Date",
                         "Trustpilot Review Links", "Status"]

PKST_DATE = pytz.timezone("Asia/Karachi")
now_pk = datetime.now(PKST_DATE)

month_list = list(calendar.month_name)[1:]
current_month = now_pk.month
current_month_name = calendar.month_name[current_month]
current_year = now_pk.year

st.markdown("""
 <style>
    #MainMenu {visibility: hidden;}
    header {visibility: hidden;}
</style>
""", unsafe_allow_html=True)


def get_min_year() -> int:
    """Minimum year available in the data."""
    return 2025


def normalize_name(name):
    """Normalize a name to consistent format (Title Case, stripped whitespace)"""
    if pd.isna(name) or name == "":
        return ""
    return str(name).strip().title()


@st.cache_resource
def get_gsheets_client(creds_dict: dict, spreadsheet_id: str):
    """Create and cache Google Sheets client + spreadsheet"""
    creds = Credentials.from_service_account_info(creds_dict, scopes=SCOPES)
    client = gspread.authorize(creds)
    return client.open_by_key(spreadsheet_id)


spreadsheet = get_gsheets_client(creds_dict, SPREADSHEET_ID)


@st.cache_data(ttl=1800)
def get_sheet_data(sheet_name: str) -> pd.DataFrame:
    """Get data from Google Sheets using gspread"""
    try:
        worksheet = spreadsheet.worksheet(sheet_name)
        raw_data = worksheet.get_all_values()
        if not raw_data:
            return pd.DataFrame()
        headers = raw_data[0]
        rows = raw_data[1:]
        data = pd.DataFrame(rows, columns=headers)
        if "Project Manager" in data.columns:
            data["Project Manager"] = data["Project Manager"].apply(normalize_name)
        return data
    except Exception as e:
        logging.error(f"Error getting data from sheet {sheet_name}: {e}")
        return pd.DataFrame()


# ---------------------------------------------------------------------------
# Cleaning / formatting helpers
# ---------------------------------------------------------------------------
def clean_data(data: pd.DataFrame, truncate_at: str = None) -> pd.DataFrame:
    """Truncate at a column boundary and normalize date/text columns."""
    if data is None or data.empty:
        return pd.DataFrame()
    cols = list(data.columns)
    if truncate_at and truncate_at in cols:
        data = data.iloc[:, :cols.index(truncate_at) + 1]
    for col in DATE_COLUMNS:
        if col in data.columns:
            data[col] = pd.to_datetime(data[col], format=DATE_FORMAT, errors="coerce")
    for col in ("Copyright", "Issues", "Last Edit (Revision)"):
        if col in data.columns:
            data[col] = data[col].astype(str).fillna("N/A")
    return data


def format_dates(data: pd.DataFrame, cols: list) -> pd.DataFrame:
    """Format datetime columns to '%d-%B-%Y' strings for display."""
    for col in cols:
        if col in data.columns:
            data[col] = pd.to_datetime(data[col], errors="coerce").dt.strftime(DATE_FORMAT)
            data[col] = data[col].replace("NaT", "N/A")
    return data


def safe_concat(dfs):
    dfs = [df for df in dfs if df is not None and not df.empty]
    return pd.concat(dfs, ignore_index=True) if dfs else pd.DataFrame()


def reindex1(data: pd.DataFrame) -> pd.DataFrame:
    data.index = range(1, len(data) + 1)
    return data


# ---------------------------------------------------------------------------
# Publishing data loaders (month / year / range / date-range share one core)
# ---------------------------------------------------------------------------
def load_publishing_data(sheet_name: str, *, month: int = None, year: int = None,
                         start_year: int = None, end_year: int = None,
                         start_date=None, end_date=None,
                         remove_duplicates: bool = False) -> pd.DataFrame:
    """Filter publishing data by Publishing Date over month / year / range / dates."""
    data = clean_data(get_sheet_data(sheet_name), truncate_at="Issues")
    if data.empty or "Publishing Date" not in data.columns:
        return pd.DataFrame()

    pub = data["Publishing Date"]
    if month and year:
        mask = (pub.dt.month == month) & (pub.dt.year == year)
    elif start_year and end_year:
        mask = (pub.dt.year >= start_year) & (pub.dt.year <= end_year)
    elif start_year:
        mask = pub.dt.year >= start_year
    elif year:
        mask = pub.dt.year == year
    elif start_date and end_date:
        mask = (pub.dt.date >= start_date) & (pub.dt.date <= end_date)
    else:
        return pd.DataFrame()

    data = data[mask].sort_values(by="Publishing Date", ascending=True)
    if remove_duplicates and "Name" in data.columns:
        data = data.drop_duplicates(subset=["Name"], keep="first")
    if data.empty:
        return pd.DataFrame()

    data = format_dates(data, DATE_COLUMNS)
    return reindex1(data)


def load_data(sheet_name: str, month_number: int, year: int) -> pd.DataFrame:
    return load_publishing_data(sheet_name, month=month_number, year=year)


def load_data_year(sheet_name: str, year: int) -> pd.DataFrame:
    return load_publishing_data(sheet_name, year=year)


def load_data_search(sheet_name: str, end_year: int, start_year: int = get_min_year()) -> pd.DataFrame:
    return load_publishing_data(sheet_name, start_year=start_year, end_year=end_year)


def load_data_filter(sheet_name: str, start_date: datetime, end_date: datetime,
                     remove_duplicates: bool = False) -> pd.DataFrame:
    return load_publishing_data(sheet_name, start_date=start_date, end_date=end_date,
                                remove_duplicates=remove_duplicates)


def clean_data_reviews(sheet_name: str) -> pd.DataFrame:
    """Clean a publishing sheet for summary computations."""
    data = clean_data(get_sheet_data(sheet_name), truncate_at="Issues")
    if data.empty:
        return data
    data = data.sort_values(by="Publishing Date", ascending=True)
    return reindex1(data)


# ---------------------------------------------------------------------------
# Review data loaders (month / year / range / date-range share one core)
# ---------------------------------------------------------------------------
def _load_reviews_core(sheet_name: str, *, month: int = None, year: int = None,
                       start_year: int = None, end_year: int = None,
                       start_date=None, end_date=None,
                       name: str = None, review_type: str = None,
                       restrict_brands: bool = False) -> pd.DataFrame:
    """Core review loader: filters on Trustpilot Review Date, optional per-PM filter."""
    data = clean_data(get_sheet_data(sheet_name), truncate_at="Issues")
    if data.empty or "Trustpilot Review Date" not in data.columns:
        return pd.DataFrame()

    rd = data["Trustpilot Review Date"]
    if start_date and end_date:
        mask = (rd.dt.date >= start_date) & (rd.dt.date <= end_date)
    elif start_year and end_year:
        mask = (rd.dt.year >= start_year) & (rd.dt.year <= end_year)
    elif start_year:
        mask = rd.dt.year >= start_year
    elif month and year:
        mask = (rd.dt.month == month) & (rd.dt.year == year)
    elif year:
        mask = rd.dt.year == year
    else:
        return pd.DataFrame()

    data = data[mask].sort_values(by="Trustpilot Review Date", ascending=True)
    if data.empty:
        return pd.DataFrame()

    if name and review_type:
        if "Project Manager" not in data.columns or "Trustpilot Review" not in data.columns:
            return pd.DataFrame()
        data = data[(data["Project Manager"] == name) & (data["Trustpilot Review"] == review_type)]
        if restrict_brands and "Brand" in data.columns:
            data = data[data["Brand"].isin(REVIEW_BRANDS)]
        if not data.empty:
            data = data.drop_duplicates(subset=["Name"])
    elif "Name" in data.columns:
        if "Trustpilot Review" in data.columns:
            # Keep latest per client, preferring 'Attained' rows.
            data = (
                data.sort_values(
                    by=["Trustpilot Review", "Trustpilot Review Date"],
                    key=lambda col: col.eq("Attained") if col.name == "Trustpilot Review" else col,
                    ascending=[False, False],
                )
                .drop_duplicates(subset=["Name"], keep="last")
            )
        else:
            data = data.drop_duplicates(subset=["Name"])
    if data.empty:
        return pd.DataFrame()
    return reindex1(data)


def load_reviews(sheet_name: str, year: int, month_number: int = None) -> pd.DataFrame:
    """Reviews for a month (or year) with per-client dedupe."""
    if month_number:
        return _load_reviews_core(sheet_name, month=month_number, year=year)
    return _load_reviews_core(sheet_name, year=year)


def load_reviews_year(sheet_name: str, year: int, name: str, type_: str = "Attained") -> pd.DataFrame:
    return _load_reviews_core(sheet_name, start_year=year, end_year=year,
                              name=name, review_type=type_, restrict_brands=True)


def load_reviews_year_to_date(sheet_name: str, year: int, name: str, type_: str = "Attained") -> pd.DataFrame:
    return _load_reviews_core(sheet_name, start_year=get_min_year(), end_year=year,
                              name=name, review_type=type_, restrict_brands=True)


def load_reviews_filter(sheet_name: str, start_date: datetime, end_date: datetime, name: str,
                        type_: str = "Attained") -> pd.DataFrame:
    return _load_reviews_core(sheet_name, start_date=start_date, end_date=end_date,
                              name=name, review_type=type_, restrict_brands=True)


def load_reviews_year_multiple(sheet_name: str, start_year: int, end_year: int, name: str,
                               type_: str = "Attained") -> pd.DataFrame:
    return _load_reviews_core(sheet_name, start_year=start_year, end_year=end_year,
                              name=name, review_type=type_, restrict_brands=True)


# ---------------------------------------------------------------------------
# Printing data
# ---------------------------------------------------------------------------
def printing_data(month: int = None, year: int = None,
                  start_year: int = None, end_year: int = None) -> tuple[pd.DataFrame, pd.DataFrame]:
    """Printing orders + per-month totals, filtered on Order Date."""
    empty = pd.DataFrame()
    data = get_sheet_data(sheet_printing)
    if data.empty or "Order Date" not in data.columns:
        return empty, empty
    data = clean_data(data, truncate_at="Accepted").astype(str)
    for col in PRINTING_DATE_COLUMNS:
        if col in data.columns:
            data[col] = pd.to_datetime(data[col], format=DATE_FORMAT, errors="coerce")

    od = data["Order Date"]
    if month and year:
        mask = (od.dt.month == month) & (od.dt.year == year)
    elif start_year and end_year:
        mask = (od.dt.year >= start_year) & (od.dt.year <= end_year)
    elif start_year:
        mask = od.dt.year >= start_year
    else:
        mask = od.dt.year == year
    data = data[mask].sort_values(by="Order Date", ascending=True)
    if data.empty:
        return empty, empty

    if "Order Cost" in data.columns:
        data["Order Cost"] = pd.to_numeric(
            data["Order Cost"].str.replace("$", "", regex=False).str.replace(",", "", regex=False),
            errors="coerce").fillna(0)
    if "No of Copies" in data.columns:
        data["No of Copies"] = pd.to_numeric(data["No of Copies"], errors="coerce").fillna(0)

    month_totals = pd.DataFrame(columns=["Month", "Total Copies", "Total Cost ($)"])
    if "Order Cost" in data.columns and "No of Copies" in data.columns:
        data["Month"] = data["Order Date"].dt.to_period("M")
        month_totals = data.groupby("Month").agg(
            Total_Copies=("No of Copies", "sum"),
            Total_Cost=("Order Cost", "sum"),
        ).reset_index()
        month_totals["Month"] = month_totals["Month"].dt.strftime("%B %Y")
        month_totals.columns = ["Month", "Total Copies", "Total Cost ($)"]
        month_totals = month_totals.sort_values(by="Total Cost ($)", ascending=False)
        month_totals = reindex1(month_totals)
        month_totals["Total Cost ($)"] = month_totals["Total Cost ($)"].map("${:,.2f}".format)

    data = format_dates(data, PRINTING_DATE_COLUMNS)
    return reindex1(data), month_totals


def get_printing_data_month(month: int, year: int) -> pd.DataFrame:
    data, _ = printing_data(month=month, year=year)
    if "Month" in data.columns:
        data = data.drop(columns="Month")
    return data


def printing_data_year(year: int) -> tuple[pd.DataFrame, pd.DataFrame]:
    return printing_data(year=year)


def printing_data_search(year: int) -> tuple[pd.DataFrame, pd.DataFrame]:
    return printing_data(start_year=get_min_year(), end_year=year)


def printing_data_year_multiple(start_year: int, end_year: int) -> tuple[pd.DataFrame, pd.DataFrame]:
    return printing_data(start_year=start_year, end_year=end_year)


# ---------------------------------------------------------------------------
# Copyright data
# ---------------------------------------------------------------------------
def copyright_data(month: int = None, year: int = None,
                   start_year: int = None, end_year: int = None) -> tuple[pd.DataFrame, int, int]:
    """Copyright submissions filtered on Submission Date. Returns (df, approved, rejected)."""
    data = get_sheet_data(sheet_copyright)
    if data.empty or "Submission Date" not in data.columns:
        return pd.DataFrame(), 0, 0
    data = clean_data(data, truncate_at="Country").astype(str)
    data["Submission Date"] = pd.to_datetime(data["Submission Date"], format=DATE_FORMAT, errors="coerce")

    sd = data["Submission Date"]
    if month and year:
        mask = (sd.dt.month == month) & (sd.dt.year == year)
    elif start_year and end_year:
        mask = (sd.dt.year >= start_year) & (sd.dt.year <= end_year)
    elif start_year:
        mask = sd.dt.year >= start_year
    else:
        mask = sd.dt.year == year
    data = data[mask].sort_values(by="Submission Date", ascending=True)

    approved = len(data[data["Result"] == "Yes"]) if "Result" in data.columns else 0
    rejected = len(data[data["Result"] == "No"]) if "Result" in data.columns else 0
    data["Submission Date"] = data["Submission Date"].dt.strftime(DATE_FORMAT)
    data = data.fillna("N/A")
    return reindex1(data), approved, rejected


def get_copyright_month(month: int, year: int) -> tuple[pd.DataFrame, int, int]:
    return copyright_data(month=month, year=year)


def copyright_year(year: int) -> tuple[pd.DataFrame, int, int]:
    return copyright_data(year=year)


def copyright_search(year: int) -> tuple[pd.DataFrame, int, int]:
    return copyright_data(start_year=get_min_year(), end_year=year)


def copyright_year_multiple(start_year: int, end_year: int) -> tuple[pd.DataFrame, int, int]:
    return copyright_data(start_year=start_year, end_year=end_year)


# ---------------------------------------------------------------------------
# A+ content data
# ---------------------------------------------------------------------------
def a_plus_data(month: int = None, year: int = None,
                start_year: int = None, end_year: int = None) -> tuple[pd.DataFrame, int]:
    """A+ content filtered on A+ Content Date. Returns (df, published_count)."""
    data = get_sheet_data(sheet_a_plus)
    if data.empty or "A+ Content Date" not in data.columns:
        return pd.DataFrame(), 0
    data = clean_data(data, truncate_at="Issues").astype(str)
    data["A+ Content Date"] = pd.to_datetime(data["A+ Content Date"], format=DATE_FORMAT, errors="coerce")

    d = data["A+ Content Date"]
    if month and year:
        mask = (d.dt.month == month) & (d.dt.year == year)
    elif start_year and end_year:
        mask = (d.dt.year >= start_year) & (d.dt.year <= end_year)
    elif start_year:
        mask = d.dt.year >= start_year
    else:
        mask = d.dt.year == year
    data = data[mask].sort_values(by="A+ Content Date", ascending=True)

    published = len(data[data["Status"] == "Published"]) if "Status" in data.columns else 0
    data["A+ Content Date"] = data["A+ Content Date"].dt.strftime(DATE_FORMAT)
    data = data.fillna("N/A")
    return reindex1(data), published


def get_A_plus_month(month: int, year: int) -> tuple[pd.DataFrame, int]:
    return a_plus_data(month=month, year=year)


def get_A_plus_year(year: int) -> tuple[pd.DataFrame, int]:
    return a_plus_data(year=year)


def get_A_plus_year_multiple(start_year: int, end_year: int) -> tuple[pd.DataFrame, int]:
    return a_plus_data(start_year=start_year, end_year=end_year)


# ---------------------------------------------------------------------------
# Similarity queries
# ---------------------------------------------------------------------------
def get_names_in_both_months(sheet_name: str, month_1: str, year1: int, month_2: str, year2: int) -> tuple:
    """
    Identifies names that appear in both months from a Google Sheet.
    Returns:
        - A set of matching names
        - A dictionary with individual counts for both months
    """
    df = get_sheet_data(sheet_name)
    if df.empty or "Name" not in df.columns or "Publishing Date" not in df.columns:
        logging.warning("Missing 'Name' or 'Date' columns or data is empty.")
        return set(), {}, 0

    df['Publishing Date'] = pd.to_datetime(df['Publishing Date'], format=DATE_FORMAT, errors='coerce')
    df = df.dropna(subset=['Publishing Date', 'Name'])
    df['Month'] = df['Publishing Date'].dt.month_name()
    df['Year'] = df['Publishing Date'].dt.year

    month_1_names = set(df[(df['Month'] == month_1) & (df['Year'] == year1)]['Name'].str.strip())
    month_2_names = set(df[(df['Month'] == month_2) & (df['Year'] == year2)]['Name'].str.strip())

    if month_1_names & month_2_names:
        names_in_both = month_1_names.intersection(month_2_names)
        counts = {}
        for name in names_in_both:
            month_1_count = df[
                (df['Month'] == month_1) & (df['Year'] == year1) & (df['Name'].str.strip() == name)
            ].shape[0]
            month_2_count = df[
                (df['Month'] == month_2) & (df['Year'] == year2) & (df['Name'].str.strip() == name)
            ].shape[0]
            counts[name] = {f"{month_1}-{year1}": month_1_count, f"{month_2}-{year2}": month_2_count}
        return names_in_both, counts, len(names_in_both)
    return set(), {}, 0


def get_names_in_both_years(sheet_name: str, year1: int, year2: int) -> tuple:
    """Identifies names that appear in both years from a Google Sheet."""
    df = get_sheet_data(sheet_name)
    if df.empty or "Name" not in df.columns or "Publishing Date" not in df.columns:
        logging.warning("Missing 'Name' or 'Publishing Date' columns or data is empty.")
        return set(), {}, 0

    df['Publishing Date'] = pd.to_datetime(df['Publishing Date'], format=DATE_FORMAT, errors='coerce')
    df = df.dropna(subset=['Publishing Date', 'Name'])
    df['Year'] = df['Publishing Date'].dt.year
    df['Name'] = df['Name'].str.strip()

    year_1_names = set(df[df['Year'] == year1]['Name'])
    year_2_names = set(df[df['Year'] == year2]['Name'])
    names_in_both = year_1_names & year_2_names

    counts = {}
    for name in names_in_both:
        year1_df = df[(df['Year'] == year1) & (df['Name'] == name)]
        year2_df = df[(df['Year'] == year2) & (df['Name'] == name)]
        counts[name] = {
            str(year1): {
                "count": year1_df.shape[0],
                "publishing_dates": year1_df['Publishing Date'].dt.strftime(DATE_FORMAT).tolist()
            },
            str(year2): {
                "count": year2_df.shape[0],
                "publishing_dates": year2_df['Publishing Date'].dt.strftime(DATE_FORMAT).tolist()
            }
        }
    return names_in_both, counts, len(names_in_both)


def get_clients_returning_in_month(sheet_name: str, start_year: int, target_month: str,
                                   target_year: int) -> tuple:
    """
    Identifies clients published starting from `start_year` that also appear
    in the specified `target_month` and `target_year`.
    """
    df = get_sheet_data(sheet_name)
    if df.empty or "Name" not in df.columns or "Publishing Date" not in df.columns:
        logging.warning("Missing 'Name' or 'Publishing Date' columns or data is empty.")
        return set(), {}, 0

    df['Publishing Date'] = pd.to_datetime(df['Publishing Date'], format=DATE_FORMAT, errors='coerce')
    df = df.dropna(subset=['Publishing Date', 'Name'])
    df['Year'] = df['Publishing Date'].dt.year
    df['Month'] = df['Publishing Date'].dt.month_name()
    df['Name'] = df['Name'].str.strip()

    baseline_df = df[df['Year'] == start_year]
    target_df = df[(df['Year'] == target_year) & (df['Month'] == target_month)]
    returning_clients = set(baseline_df['Name']) & set(target_df['Name'])

    counts = {}
    for name in returning_clients:
        client_baseline = baseline_df[baseline_df['Name'] == name]
        client_target = target_df[target_df['Name'] == name]
        counts[name] = {
            f"from_{start_year}_baseline": {
                "count": client_baseline.shape[0],
                "publishing_dates": client_baseline['Publishing Date'].dt.strftime(DATE_FORMAT).tolist()
            },
            f"{target_year}_{target_month}": {
                "count": client_target.shape[0],
                "publishing_dates": client_target['Publishing Date'].dt.strftime(DATE_FORMAT).tolist()
            }
        }
    return returning_clients, counts, len(returning_clients)


def get_names_in_year(sheet_name: str, year: int):
    """
    Finds names that appear in multiple months within the same year.
    Returns:
        - A DataFrame of names with counts per month
        - A dictionary summary of names with their total appearances and months active
        - Total count of such names
    """
    df = get_sheet_data(sheet_name)
    if df.empty or "Name" not in df.columns or "Publishing Date" not in df.columns:
        logging.warning("Missing 'Name' or 'Publishing Date' columns, or data is empty.")
        return pd.DataFrame(), {}, 0

    df['Publishing Date'] = pd.to_datetime(df['Publishing Date'], format=DATE_FORMAT, errors='coerce')
    df = df.dropna(subset=['Publishing Date', 'Name'])
    df['Month'] = df['Publishing Date'].dt.month_name()
    df['Year'] = df['Publishing Date'].dt.year
    df = df[df['Year'] == year]

    if df.empty:
        logging.warning(f"No records found for year {year}.")
        return pd.DataFrame(), {}, 0

    monthly_counts = (
        df.groupby(['Name', 'Month'])
        .size()
        .unstack(fill_value=0)
        .reindex(columns=[
            'January', 'February', 'March', 'April', 'May', 'June',
            'July', 'August', 'September', 'October', 'November', 'December'
        ], fill_value=0)
    )
    monthly_counts['Active Months'] = (monthly_counts > 0).sum(axis=1)
    multi_month_names = monthly_counts[monthly_counts['Active Months'] > 1].copy()

    month_cols = multi_month_names.columns[:-2]
    summary = {}
    for name in multi_month_names.index:
        active_months = [month for month in month_cols if multi_month_names.at[name, month] > 0]
        indexed_months = {i + 1: month for i, month in enumerate(active_months)}
        summary[name] = {
            "Months Active": indexed_months,
            "Month Count": int(multi_month_names.at[name, "Active Months"]),
        }
    return multi_month_names, summary, len(multi_month_names)


# ---------------------------------------------------------------------------
# Charts
# ---------------------------------------------------------------------------
def create_review_pie_chart(review_data: dict, title: str):
    """Create pie chart for review distribution"""
    if not review_data or sum(review_data.values()) == 0:
        return None
    custom_colors = {
        "Attained": "#7dff8d",
        "Pending": "#ffc444",
        "Negative": "#ff4b4b",
        "Sent": "#77e5f7"
    }
    fig = px.pie(
        values=list(review_data.values()),
        names=list(review_data.keys()),
        title=title,
        color=list(review_data.keys()),
        color_discrete_map=custom_colors
    )
    fig.update_traces(textposition="inside", textinfo="percent+label")
    return fig


def create_platform_comparison_chart(usa_data: dict, uk_data: dict):
    """Create comparison chart for platforms"""
    platforms = PLATFORMS
    fig = go.Figure(data=[
        go.Bar(name='USA', x=platforms, y=list(usa_data.values()), marker_color="#23A0F8"),
        go.Bar(name='UK', x=platforms, y=list(uk_data.values()), marker_color="#ff7f0e")
    ])
    fig.update_layout(
        title='Platform Distribution: USA vs UK',
        barmode='group',
        xaxis_title='Platforms',
        yaxis_title='Number of Reviews'
    )
    return fig


def create_brand_chart(usa_brands: dict, uk_brands: dict):
    """Create brand distribution chart"""
    all_brands = list(usa_brands.keys()) + list(uk_brands.keys())
    all_values = list(usa_brands.values()) + list(uk_brands.values())
    regions = ['USA'] * len(usa_brands) + ['UK'] * len(uk_brands)

    fig = px.bar(
        x=all_brands,
        y=all_values,
        color=regions,
        title='Brand Distribution by Region',
        color_discrete_map={'USA': '#23A0F8', 'UK': '#ff7f0e'},
        labels={"x": "Brands", "y": "Number of Clients"}
    )
    return fig


# ---------------------------------------------------------------------------
# Stats helpers (shared by all summary generators)
# ---------------------------------------------------------------------------
def _pm_counts(df: pd.DataFrame, col_name: str, review_type: str) -> tuple[pd.DataFrame, int]:
    """Per-PM count table + total for one review type."""
    if df.empty:
        return pd.DataFrame(columns=["Project Manager", col_name]), 0
    out = df[df["Trustpilot Review"] == review_type].groupby("Project Manager")[
        "Trustpilot Review"].count().reset_index()
    out.columns = ["Project Manager", col_name]
    out.index = range(1, len(out) + 1)
    return out, int(out[col_name].sum())


def _brand_platform_stats(usa_clean: pd.DataFrame, uk_clean: pd.DataFrame, allowed_brands: list) -> dict:
    """Brand/platform counts, review status counts and pending/sent details."""

    def _region_review_counts(df):
        if "Trustpilot Review" in df.columns and "Brand" in df.columns:
            counts = df[df["Brand"].isin(allowed_brands)]["Trustpilot Review"].value_counts()
            return counts.get("Sent", 0), counts.get("Pending", 0), counts.get("Negative", 0)
        return 0, 0, 0

    usa_sent, usa_pending, usa_negative = _region_review_counts(usa_clean)
    uk_sent, uk_pending, uk_negative = _region_review_counts(uk_clean)

    usa_brand_counts = usa_clean["Brand"].value_counts()
    uk_brand_counts = uk_clean["Brand"].value_counts()
    usa_platform_counts = usa_clean["Platform"].value_counts()
    uk_platform_counts = uk_clean["Platform"].value_counts()

    usa_brands = {b: int(usa_brand_counts.get(b, 0)) for b in USA_BRANDS}
    uk_brands = {b: int(uk_brand_counts.get(b, 0)) for b in UK_BRANDS}
    usa_platforms = {p: int(usa_platform_counts.get(p, 0)) for p in PLATFORMS}
    uk_platforms = {p: int(uk_platform_counts.get(p, 0)) for p in PLATFORMS}

    combined_pending_sent = pd.concat([usa_clean, uk_clean], ignore_index=True)
    pending_sent_details = combined_pending_sent[
        combined_pending_sent["Trustpilot Review"].isin(["Sent", "Pending"]) &
        combined_pending_sent["Brand"].isin(allowed_brands)
        ]
    pending_sent_details = pending_sent_details[
        ["Name", "Brand", "Project Manager", "Trustpilot Review", "Status"]]
    pending_sent_details.index = range(1, len(pending_sent_details) + 1)

    return {
        "usa_brands": usa_brands,
        "uk_brands": uk_brands,
        "usa_platforms": usa_platforms,
        "uk_platforms": uk_platforms,
        "pending_sent_details": pending_sent_details,
        "usa_review": {"Sent": usa_sent, "Pending": usa_pending, "Negative": usa_negative},
        "uk_review": {"Sent": uk_sent, "Pending": uk_pending, "Negative": uk_negative},
    }


def _printing_stats(printing_df: pd.DataFrame) -> dict:
    def _safe_max(s: pd.Series):
        return s.max() if len(s) else 0

    copies = printing_df["No of Copies"] if "No of Copies" in printing_df.columns else pd.Series(dtype=float)
    cost = printing_df["Order Cost"] if "Order Cost" in printing_df.columns else pd.Series(dtype=float)
    total_copies = copies.sum()
    total_cost = cost.sum()
    average = total_cost / total_copies if total_copies > 0 else 0
    return {
        "Total_copies": total_copies,
        "Total_cost": total_cost,
        "Highest_cost": _safe_max(cost),
        "Lowest_cost": _safe_max(cost) if len(cost) == 0 else cost.min(),
        "Highest_copies": _safe_max(copies),
        "Lowest_copies": _safe_max(copies) if len(copies) == 0 else copies.min(),
        "Average": average,
    }


def _copyright_stats(copyright_df: pd.DataFrame, result_count: int, result_count_no: int) -> dict:
    country = copyright_df["Country"].value_counts() if "Country" in copyright_df.columns else pd.Series(dtype=int)
    usa = country.get("USA", 0)
    canada = country.get("Canada", 0)
    uk = country.get("UK", 0)
    return {
        "Total_copyrights": len(copyright_df),
        "Total_cost_copyright": (usa * COPYRIGHT_RATES["USA"]) + (canada * COPYRIGHT_RATES["Canada"]) + (uk * COPYRIGHT_RATES["UK"]),
        "result_count": result_count,
        "result_count_no": result_count_no,
        "usa_copyrights": usa,
        "canada_copyrights": canada,
        "uk": uk,
    }


def _publishing_monthly(df: pd.DataFrame, col_name: str) -> pd.DataFrame:
    if df.empty:
        return pd.DataFrame(columns=["Month", "Month_Sort", col_name])
    out = df.groupby(df["Publishing Date"].dt.to_period("M")).size().reset_index(name=col_name)
    out["Month"] = out["Publishing Date"].dt.strftime("%B %Y")
    out["Month_Sort"] = out["Publishing Date"].dt.to_timestamp()
    return out[["Month", "Month_Sort", col_name]]


def _reviews_monthly(df: pd.DataFrame, col_name: str) -> pd.DataFrame:
    if df.empty:
        return pd.DataFrame(columns=["Month", col_name])
    out = df.groupby(df["Trustpilot Review Date"].dt.to_period("M")).size().reset_index(name=col_name)
    out["Month"] = out["Trustpilot Review Date"].dt.strftime("%B %Y")
    return out[["Month", col_name]]


def _region_reviews_per_month(usa_df: pd.DataFrame, uk_df: pd.DataFrame, kind: str) -> pd.DataFrame:
    """Merge USA/UK monthly review counts into a single table."""
    usa_col = f"USA {kind} Reviews"
    uk_col = f"UK {kind} Reviews"
    total_col = f"Total {kind} Reviews"
    out = pd.merge(
        _reviews_monthly(usa_df, usa_col),
        _reviews_monthly(uk_df, uk_col),
        on="Month", how="outer"
    ).fillna(0)
    out[total_col] = out[usa_col] + out[uk_col]
    out["Month_Num"] = pd.to_datetime(out["Month"], format="%B %Y")
    out = out.sort_values(by=total_col, ascending=False)
    out.index = range(1, len(out) + 1)
    return out.drop(columns="Month_Num")


def _merged_attained(attained_details: pd.DataFrame) -> pd.DataFrame:
    if attained_details.empty:
        return pd.DataFrame(columns=["Project Manager", "Count", "Clients"])
    attained_count = attained_details.groupby("Project Manager").size().reset_index(name="Count")
    attained_clients = attained_details.groupby("Project Manager")["Name"].apply(list).reset_index(name="Clients")
    merged = attained_count.merge(attained_clients, on="Project Manager", how="left")
    merged = merged.sort_values(by="Count", ascending=False)
    merged.index = range(1, len(merged) + 1)
    return merged


def _monthly_review_table(details_df: pd.DataFrame, col_name: str) -> pd.DataFrame:
    """Monthly counts of a review type from a details dataframe."""
    if details_df.empty or "Trustpilot Review Date" not in details_df.columns:
        return pd.DataFrame(columns=["Month", col_name])
    df = details_df.copy()
    df["Trustpilot Review Date"] = pd.to_datetime(df["Trustpilot Review Date"], errors="coerce")
    out = df.groupby(df["Trustpilot Review Date"].dt.to_period("M")).size().reset_index(name=col_name)
    out["Month"] = out["Trustpilot Review Date"].dt.strftime("%B %Y")
    out = out.sort_values(by=col_name, ascending=False)
    out.index = range(1, len(out) + 1)
    return out[["Month", col_name]]


# ---------------------------------------------------------------------------
# Summary generators
# ---------------------------------------------------------------------------
def summary(month: int, year: int) -> dict:
    """Monthly summary for the 'Summary' report. Returns {} when no data."""
    usa_clean = clean_data_reviews(sheet_usa)
    uk_clean = clean_data_reviews(sheet_uk)
    usa_clean = usa_clean[(usa_clean["Publishing Date"].dt.month == month) &
                          (usa_clean["Publishing Date"].dt.year == year)]
    uk_clean = uk_clean[(uk_clean["Publishing Date"].dt.month == month) &
                        (uk_clean["Publishing Date"].dt.year == year)]
    if usa_clean.empty or uk_clean.empty:
        return {}

    usa_clean = usa_clean.drop_duplicates(subset=["Name"], keep="last")
    uk_clean = uk_clean.drop_duplicates(subset=["Name"], keep="last")
    Issues_usa = usa_clean["Issues"].value_counts()
    Issues_uk = uk_clean["Issues"].value_counts()
    total_unique_clients = usa_clean["Name"].nunique() + uk_clean["Name"].nunique()

    combined = pd.concat([usa_clean[["Name", "Brand", "Project Manager", "Email"]],
                          uk_clean[["Name", "Brand", "Project Manager", "Email"]]])
    combined.index = range(1, len(combined) + 1)

    bps = _brand_platform_stats(usa_clean, uk_clean, ALLOWED_BRANDS)

    usa_reviews_df = load_reviews(sheet_usa, year, month)
    uk_reviews_df = load_reviews(sheet_uk, year, month)
    combined_data = safe_concat([usa_reviews_df, uk_reviews_df])

    _, usa_total_attained = _pm_counts(usa_reviews_df, "Attained Reviews", "Attained")
    _, usa_total_negative = _pm_counts(usa_reviews_df, "Negative Reviews", "Negative")
    _, uk_total_attained = _pm_counts(uk_reviews_df, "Attained Reviews", "Attained")
    _, uk_total_negative = _pm_counts(uk_reviews_df, "Negative Reviews", "Negative")

    if not combined_data.empty:
        attained_reviews_per_pm = (
            combined_data[combined_data["Trustpilot Review"] == "Attained"]
            .groupby("Project Manager")["Trustpilot Review"].count()
            .reset_index(name="Attained Reviews")
            .sort_values(by="Attained Reviews", ascending=False)
        )
        attained_reviews_per_pm.index = range(1, len(attained_reviews_per_pm) + 1)
        negative_reviews_per_pm = (
            combined_data[combined_data["Trustpilot Review"] == "Negative"]
            .groupby("Project Manager")["Trustpilot Review"].count()
            .reset_index(name="Negative Reviews")
            .sort_values(by="Negative Reviews", ascending=False)
        )
        negative_reviews_per_pm.index = range(1, len(negative_reviews_per_pm) + 1)
        review_details_df = combined_data.sort_values(by="Project Manager", ascending=True)
        review_details_df["Trustpilot Review Date"] = pd.to_datetime(
            review_details_df["Trustpilot Review Date"], errors="coerce"
        ).dt.strftime(DATE_FORMAT)
        attained_details = review_details_df[
            review_details_df["Trustpilot Review"] == "Attained"][REVIEW_DETAIL_COLUMNS]
        attained_details.index = range(1, len(attained_details) + 1)
        negative_details = review_details_df[
            review_details_df["Trustpilot Review"] == "Negative"][REVIEW_DETAIL_COLUMNS]
        negative_details.index = range(1, len(negative_details) + 1)
    else:
        attained_reviews_per_pm = pd.DataFrame(columns=["Project Manager", "Attained Reviews"])
        negative_reviews_per_pm = pd.DataFrame(columns=["Project Manager", "Negative Reviews"])
        attained_details = pd.DataFrame(columns=REVIEW_DETAIL_COLUMNS)
        negative_details = attained_details.copy()

    usa_review = {
        "Attained": usa_total_attained,
        "Sent": bps["usa_review"]["Sent"],
        "Pending": bps["usa_review"]["Pending"],
        "Negative": bps["usa_review"]["Negative"] + usa_total_negative,
    }
    uk_review = {
        "Attained": uk_total_attained,
        "Sent": bps["uk_review"]["Sent"],
        "Pending": bps["uk_review"]["Pending"],
        "Negative": bps["uk_review"]["Negative"] + uk_total_negative,
    }

    printing_stats = _printing_stats(get_printing_data_month(month, year))
    copyright_stats = _copyright_stats(*get_copyright_month(month, year))
    _, a_plus_count = get_A_plus_month(month, year)

    return {
        "usa_review": usa_review, "uk_review": uk_review,
        "usa_brands": bps["usa_brands"], "uk_brands": bps["uk_brands"],
        "usa_platforms": bps["usa_platforms"], "uk_platforms": bps["uk_platforms"],
        "printing_stats": printing_stats, "copyright_stats": copyright_stats,
        "a_plus_count": a_plus_count, "total_unique_clients": total_unique_clients,
        "chargeback": _chargeback_for_period(month, year),
        "combined": combined, "attained_reviews_per_pm": attained_reviews_per_pm,
        "attained_details": attained_details, "pending_sent_details": bps["pending_sent_details"],
        "negative_reviews_per_pm": negative_reviews_per_pm, "negative_details": negative_details,
        "Issues_usa": Issues_usa, "Issues_uk": Issues_uk,
    }


def generate_year_summary(start_year: int, end_year: int = None) -> dict:
    """Year (or multi-year) summary for 'Year Summary' and 'Custom Summary'. Returns {} when no data."""
    if end_year is None:
        end_year = start_year
    # Single-year reports historically used a 5-brand list; multi-year used 6.
    review_brands = REVIEW_BRANDS if end_year != start_year else ALLOWED_BRANDS

    usa_clean = clean_data_reviews(sheet_usa)
    uk_clean = clean_data_reviews(sheet_uk)
    usa_clean = usa_clean[(usa_clean["Publishing Date"].dt.year >= start_year) &
                          (usa_clean["Publishing Date"].dt.year <= end_year)]
    uk_clean = uk_clean[(uk_clean["Publishing Date"].dt.year >= start_year) &
                        (uk_clean["Publishing Date"].dt.year <= end_year)]
    if usa_clean.empty and uk_clean.empty:
        return {}

    usa_clean = usa_clean.drop_duplicates(subset=["Name"], keep="first")
    uk_clean = uk_clean.drop_duplicates(subset=["Name"], keep="first")

    bps = _brand_platform_stats(usa_clean, uk_clean, review_brands)
    Issues_usa = usa_clean["Issues"].value_counts()
    Issues_uk = uk_clean["Issues"].value_counts()
    total_unique_clients = usa_clean["Name"].nunique() + uk_clean["Name"].nunique()

    combined = pd.concat([usa_clean[["Name", "Brand", "Project Manager", "Email"]],
                          uk_clean[["Name", "Brand", "Project Manager", "Email"]]])
    combined.index = range(1, len(combined) + 1)

    combined_monthly = pd.merge(
        _publishing_monthly(usa_clean, "USA Published"),
        _publishing_monthly(uk_clean, "UK Published"),
        on=["Month", "Month_Sort"], how="outer"
    ).fillna(0)
    combined_monthly["Total Published"] = combined_monthly["USA Published"] + combined_monthly["UK Published"]
    combined_monthly = combined_monthly.sort_values("Month_Sort").drop(columns="Month_Sort")
    combined_monthly.index = range(1, len(combined_monthly) + 1)

    # Per-PM review loads (same PM list logic as the original reports)
    pms_uk = load_data_search(sheet_uk, current_year)
    pms_usa = load_data_search(sheet_usa, current_year)
    pm_list_usa = list(set(pms_usa["Project Manager"].dropna().unique().tolist() + ["Unknown"]))
    pm_list_uk = list(set(pms_uk["Project Manager"].dropna().unique().tolist() + ["Unknown"]))

    usa_reviews_per_pm = safe_concat([
        _load_reviews_core(sheet_usa, start_year=start_year, end_year=end_year,
                           name=pm, review_type="Attained", restrict_brands=True)
        for pm in pm_list_usa])
    uk_reviews_per_pm = safe_concat([
        _load_reviews_core(sheet_uk, start_year=start_year, end_year=end_year,
                           name=pm, review_type="Attained", restrict_brands=True)
        for pm in pm_list_uk])
    combined_data = safe_concat([usa_reviews_per_pm, uk_reviews_per_pm])

    _, usa_total_attained = _pm_counts(usa_reviews_per_pm, "Attained Reviews", "Attained")
    _, uk_total_attained = _pm_counts(uk_reviews_per_pm, "Attained Reviews", "Attained")

    if not combined_data.empty:
        attained_reviews_per_pm = (
            combined_data.groupby("Project Manager")["Trustpilot Review"]
            .count().reset_index()
        )
        attained_reviews_per_pm.columns = ["Project Manager", "Attained Reviews"]
        attained_reviews_per_pm = attained_reviews_per_pm.sort_values(by="Attained Reviews", ascending=False)
        attained_reviews_per_pm.index = range(1, len(attained_reviews_per_pm) + 1)

        review_details_df = combined_data.sort_values(by="Project Manager", ascending=True)
        review_details_df["Trustpilot Review Date"] = pd.to_datetime(
            review_details_df["Trustpilot Review Date"], errors="coerce"
        ).dt.strftime(DATE_FORMAT)
        attained_details = review_details_df[REVIEW_DETAIL_COLUMNS]
        attained_details.index = range(1, len(attained_details) + 1)
        attained_details["Trustpilot Review Date"] = pd.to_datetime(
            attained_details["Trustpilot Review Date"], errors="coerce")
        merged_attained = _merged_attained(attained_details)
        attained_details["Trustpilot Review Date"] = pd.to_datetime(
            attained_details["Trustpilot Review Date"], errors="coerce").dt.strftime(DATE_FORMAT)
        attained_reviews_per_month = _region_reviews_per_month(usa_reviews_per_pm, uk_reviews_per_pm, "Attained")
    else:
        attained_reviews_per_pm = pd.DataFrame(columns=["Project Manager", "Attained Reviews"])
        attained_details = pd.DataFrame(columns=REVIEW_DETAIL_COLUMNS)
        merged_attained = pd.DataFrame(columns=["Project Manager", "Count", "Clients"])
        attained_reviews_per_month = pd.DataFrame(columns=["Month", "Total Attained Reviews"])

    usa_negative_per_pm = safe_concat([
        _load_reviews_core(sheet_usa, start_year=start_year, end_year=end_year,
                           name=pm, review_type="Negative", restrict_brands=True)
        for pm in pm_list_usa])
    uk_negative_per_pm = safe_concat([
        _load_reviews_core(sheet_uk, start_year=start_year, end_year=end_year,
                           name=pm, review_type="Negative", restrict_brands=True)
        for pm in pm_list_uk])
    combined_negative_data = safe_concat([usa_negative_per_pm, uk_negative_per_pm])

    _, usa_total_negative = _pm_counts(usa_negative_per_pm, "Negative Reviews", "Negative")
    _, uk_total_negative = _pm_counts(uk_negative_per_pm, "Negative Reviews", "Negative")

    if not combined_negative_data.empty:
        negative_reviews_per_pm = (
            combined_negative_data.groupby("Project Manager")["Trustpilot Review"]
            .count().reset_index()
        )
        negative_reviews_per_pm.columns = ["Project Manager", "Negative Reviews"]
        negative_reviews_per_pm = negative_reviews_per_pm.sort_values(by="Negative Reviews", ascending=False)
        negative_reviews_per_pm.index = range(1, len(negative_reviews_per_pm) + 1)

        negative_details_df = combined_negative_data.sort_values(by="Project Manager", ascending=True)
        negative_details_df["Trustpilot Review Date"] = pd.to_datetime(
            negative_details_df["Trustpilot Review Date"], errors="coerce"
        ).dt.strftime(DATE_FORMAT)
        negative_details = negative_details_df[REVIEW_DETAIL_COLUMNS]
        negative_details.index = range(1, len(negative_details) + 1)
        negative_details["Trustpilot Review Date"] = pd.to_datetime(
            negative_details["Trustpilot Review Date"], errors="coerce")
        negative_reviews_per_month = _region_reviews_per_month(usa_negative_per_pm, uk_negative_per_pm, "Negative")
        negative_details["Trustpilot Review Date"] = pd.to_datetime(
            negative_details["Trustpilot Review Date"], errors="coerce").dt.strftime(DATE_FORMAT)
    else:
        negative_reviews_per_pm = pd.DataFrame(columns=["Project Manager", "Negative Reviews"])
        negative_details = pd.DataFrame(columns=REVIEW_DETAIL_COLUMNS)
        negative_reviews_per_month = pd.DataFrame(columns=["Month", "Total Negative Reviews"])

    usa_review = {
        "Attained": usa_total_attained,
        "Sent": bps["usa_review"]["Sent"],
        "Pending": bps["usa_review"]["Pending"],
        "Negative": usa_total_negative,
    }
    uk_review = {
        "Attained": uk_total_attained,
        "Sent": bps["uk_review"]["Sent"],
        "Pending": bps["uk_review"]["Pending"],
        "Negative": uk_total_negative,
    }

    printing_data_out, monthly_printing = printing_data_year_multiple(start_year, end_year)
    printing_stats = _printing_stats(printing_data_out)
    copyright_stats = _copyright_stats(*copyright_year_multiple(start_year, end_year))
    _, a_plus_count = get_A_plus_year_multiple(start_year, end_year)

    return {
        "usa_review": usa_review, "uk_review": uk_review,
        "usa_brands": bps["usa_brands"], "uk_brands": bps["uk_brands"],
        "usa_platforms": bps["usa_platforms"], "uk_platforms": bps["uk_platforms"],
        "printing_stats": printing_stats, "monthly_printing": monthly_printing,
        "copyright_stats": copyright_stats, "a_plus_count": a_plus_count,
        "chargeback": _chargeback_for_period(start_year=start_year, end_year=end_year),
        "total_unique_clients": total_unique_clients, "combined": combined,
        "attained_reviews_per_pm": attained_reviews_per_pm, "attained_details": attained_details,
        "merged_attained": merged_attained, "attained_reviews_per_month": attained_reviews_per_month,
        "pending_sent_details": bps["pending_sent_details"],
        "negative_reviews_per_pm": negative_reviews_per_pm, "negative_details": negative_details,
        "negative_reviews_per_month": negative_reviews_per_month,
        "combined_monthly": combined_monthly,
        "Issues_usa": Issues_usa, "Issues_uk": Issues_uk,
    }


# ---------------------------------------------------------------------------
# PDF report
# ---------------------------------------------------------------------------
def generate_summary_report_pdf(
        usa_review_data, uk_review_data, usa_brands, uk_brands, usa_platforms, uk_platforms,
        printing_stats, copyright_stats, a_plus,
        selected_month=None, start_year=None, end_year=None, filename=None
):
    """
    Generate a PDF summary report with proper year / range handling
    """
    if selected_month and start_year and end_year:
        title_text = f"{selected_month} ({start_year}–{end_year}) Summary Report"
        filename = f"{selected_month}_{start_year}_{end_year}_Summary_Report.pdf"
    elif selected_month and start_year:
        title_text = f"{selected_month} {start_year} Summary Report"
        filename = f"{selected_month}_{start_year}_Summary_Report.pdf"
    elif start_year and end_year:
        title_text = f"{start_year}–{end_year} Summary Report"
        filename = f"{start_year}_{end_year}_Summary_Report.pdf"
    elif start_year:
        title_text = f"{start_year} Summary Report"
        filename = f"{start_year}_Summary_Report.pdf"
    else:
        title_text = "Summary Report"
        filename = "Summary_Report.pdf"

    filename = filename.replace(" ", "_")

    buffer = BytesIO()
    doc = SimpleDocTemplate(
        buffer, pagesize=A4, rightMargin=72, leftMargin=72, topMargin=72, bottomMargin=18
    )

    styles = getSampleStyleSheet()
    title_style = ParagraphStyle(
        'Title', parent=styles['Heading1'], fontSize=24, spaceAfter=30,
        alignment=TA_CENTER, textColor=colors.darkblue
    )
    section_style = ParagraphStyle(
        'Section', parent=styles['Heading2'], fontSize=16, spaceBefore=20,
        spaceAfter=12, textColor=colors.darkgreen
    )
    subsection_style = ParagraphStyle(
        'SubSection', parent=styles['Heading3'], fontSize=12, spaceBefore=12,
        spaceAfter=8, textColor=colors.darkblue
    )

    story = []
    story.append(Paragraph(title_text, title_style))
    story.append(Spacer(1, 20))

    usa_total = sum(usa_review_data.values())
    uk_total = sum(uk_review_data.values())
    usa_attained = usa_review_data.get("Attained", 0)
    uk_attained = uk_review_data.get("Attained", 0)
    combined_total = usa_total + uk_total
    combined_attained = usa_attained + uk_attained
    usa_pct = (usa_attained / usa_total * 100) if usa_total else 0
    uk_pct = (uk_attained / uk_total * 100) if uk_total else 0
    combined_pct = (combined_attained / combined_total * 100) if combined_total else 0

    story.append(Paragraph("📝 Review Analytics", section_style))
    story.append(HRFlowable(width="100%", thickness=1, color=colors.lightgrey))
    story.append(Spacer(1, 12))

    review_table = Table([
        ["Region", "Total Reviews", "Attained", "Success Rate"],
        ["USA", f"{usa_total:,}", f"{usa_attained:,}", f"{usa_pct:.1f}%"],
        ["UK", f"{uk_total:,}", f"{uk_attained:,}", f"{uk_pct:.1f}%"],
        ["Combined", f"{combined_total:,}", f"{combined_attained:,}", f"{combined_pct:.1f}%"],
    ])
    review_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
        ('GRID', (0, 0), (-1, -1), 1, colors.black),
        ('ALIGN', (1, 1), (-1, -1), 'CENTER'),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold')
    ]))
    story.append(review_table)

    story.append(Spacer(1, 20))
    story.append(Paragraph("📱 Platform Distribution", subsection_style))
    for label, platforms in [("USA", usa_platforms), ("UK", uk_platforms)]:
        story.append(Paragraph(f"{label} Platforms", styles['Normal']))
        table = Table([["Platform", "Count"]] + [[k, v] for k, v in platforms.items()])
        table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.lightblue),
            ('GRID', (0, 0), (-1, -1), 1, colors.black),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold')
        ]))
        story.append(table)
        story.append(Spacer(1, 12))

    story.append(Paragraph("🏷️ Brand Performance", subsection_style))
    brand_table = Table(
        [["USA Brand", "Count", "UK Brand", "Count"]] +
        list(zip_longest(
            list(usa_brands.keys()) + ["Total"],
            list(usa_brands.values()) + [sum(usa_brands.values())],
            list(uk_brands.keys()) + ["Total"],
            list(uk_brands.values()) + [sum(uk_brands.values())]
        ))
    )
    brand_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.lightgreen),
        ('GRID', (0, 0), (-1, -1), 1, colors.black),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold')
    ]))
    story.append(brand_table)

    story.append(PageBreak())
    story.append(Paragraph("🖨️ Printing Analytics", section_style))
    story.append(HRFlowable(width="100%", thickness=1, color=colors.lightgrey))

    printing_table = Table([
        ["Metric", "Value"],
        ["Total Copies", f"{printing_stats['Total_copies']:,}"],
        ["Total Cost", f"${printing_stats['Total_cost']:,.2f}"],
        ["Avg Cost/Copy", f"${printing_stats['Average']:.2f}"]
    ])
    printing_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.orange),
        ('GRID', (0, 0), (-1, -1), 1, colors.black)
    ]))
    story.append(printing_table)

    story.append(Spacer(1, 20))
    story.append(Paragraph("©️ Copyright Analytics", section_style))
    story.append(HRFlowable(width="100%", thickness=1, color=colors.lightgrey))

    success_rate = (
        copyright_stats['result_count'] / copyright_stats['Total_copyrights'] * 100
        if copyright_stats['Total_copyrights'] else 0
    )
    copyright_table = Table([
        ["Metric", "Value"],
        ["Total Copyrights", copyright_stats['Total_copyrights']],
        ["Success Rate", f"{success_rate:.1f}%"],
        ["Total Cost", f"${copyright_stats['Total_cost_copyright']:,}"]
    ])
    copyright_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.purple),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
        ('GRID', (0, 0), (-1, -1), 1, colors.black)
    ]))
    story.append(copyright_table)

    story.append(Spacer(1, 20))
    story.append(Paragraph("A+ Content", section_style))
    story.append(Table([["Total A+", a_plus]]))

    story.append(Spacer(1, 30))
    story.append(HRFlowable(width="100%", thickness=1, color=colors.lightgrey))
    story.append(Paragraph(
        f"Generated on {datetime.now().strftime('%B %d, %Y %I:%M %p')}", styles['Normal']
    ))

    doc.build(story)
    pdf_data = buffer.getvalue()
    buffer.close()
    return pdf_data, filename


# ---------------------------------------------------------------------------
# Sales / ISBN
# ---------------------------------------------------------------------------
def sales(month: int, year: int) -> pd.DataFrame:
    data = get_sheet_data(sheet_sales)
    if data.empty or "Payment Date" not in data.columns:
        return pd.DataFrame()
    data = clean_data(data, truncate_at="Payment")
    data["Payment Date"] = pd.to_datetime(data["Payment Date"], errors="coerce")
    data = data[(data["Payment Date"].dt.month == month) & (data["Payment Date"].dt.year == year)]
    if "Payment" in data.columns:
        data["Payment"] = pd.to_numeric(
            data["Payment"].astype(str).str.replace("$", "", regex=False).str.replace(",", "", regex=False),
            errors="coerce")
    data["Payment Date"] = data["Payment Date"].dt.strftime(DATE_FORMAT)
    return reindex1(data)


def sales_year(year: int) -> pd.DataFrame:
    data = get_sheet_data(sheet_sales)
    if data.empty or "Payment Date" not in data.columns:
        return pd.DataFrame()
    data = clean_data(data, truncate_at="Payment")
    data["Payment Date"] = pd.to_datetime(data["Payment Date"], errors="coerce")
    data = data[data["Payment Date"].dt.year == year]
    if "Payment" in data.columns:
        data["Payment"] = pd.to_numeric(
            data["Payment"].astype(str).str.replace("$", "", regex=False).str.replace(",", "", regex=False),
            errors="coerce")
    data["Payment Date"] = data["Payment Date"].dt.strftime(DATE_FORMAT)
    return reindex1(data)


def nielsen_isbn() -> pd.DataFrame:
    data = get_sheet_data(sheet_nielsen)
    if data.empty:
        return pd.DataFrame()
    return clean_data(data, truncate_at="Author")


CHARGEBACK_DATE_COLUMNS = ["Payment Date", "Chargeback Date", "Claim Submission Date",
                           "Rebuttal Date", "Result Date"]


def _parse_money(value) -> float:
    if pd.isna(value):
        return 0.0
    cleaned = str(value).replace("$", "").replace(",", "").strip()
    try:
        return float(cleaned)
    except ValueError:
        return 0.0


def load_chargeback() -> pd.DataFrame:
    data = get_sheet_data(sheet_chargeback)
    if data.empty:
        return data
    if "Payment" in data.columns:
        data["Payment"] = data["Payment"].apply(_parse_money)
    if "Result" in data.columns:
        result = data["Result"].astype(str).str.strip()
        data["Favourable"] = result.str.contains("Favourable", case=False, na=False)
        data["Lost"] = result.str.contains("Rejected|Lost", case=False, na=False)
        data["Awaiting"] = ~(data["Favourable"] | data["Lost"])
    if "Chargeback Date" in data.columns:
        data["_Chargeback_dt"] = pd.to_datetime(
            data["Chargeback Date"], format=DATE_FORMAT, errors="coerce")
    return format_dates(data, CHARGEBACK_DATE_COLUMNS)


def render_chargeback_stats(df: pd.DataFrame, heading: str = "### 📊 Chargeback Statistics") -> None:
    total = len(df)
    total_payment = df["Payment"].sum() if "Payment" in df else 0.0
    fav = df[df["Favourable"]] if "Favourable" in df else df.iloc[0:0]
    lost = df[df["Lost"]] if "Lost" in df else df.iloc[0:0]
    awaiting = df[df["Awaiting"]] if "Awaiting" in df else df.iloc[0:0]
    fav_count = len(fav)
    payment_saved = fav["Payment"].sum() if "Payment" in df else 0.0
    lost_count = len(lost)
    lost_payment = lost["Payment"].sum() if "Payment" in df else 0.0
    awaiting_count = len(awaiting)
    awaiting_payment = awaiting["Payment"].sum() if "Payment" in df else 0.0
    decided = fav_count + lost_count
    win_rate = (fav_count / decided * 100) if decided else 0.0
    st.markdown(heading)
    c1, c2, c3 = st.columns(3)
    c1.metric("💳 Total Chargebacks", total)
    c2.metric("💰 Total Payment at Risk", f"${total_payment:,.2f}")
    c3.metric("✅ Payment Saved", f"${payment_saved:,.2f}")
    c1, c2, c3 = st.columns(3)
    c1.metric("🏆 Won (Favourable)", f"{fav_count} ({win_rate:.1f}%)")
    c2.metric("❌ Lost (Rejected)", f"{lost_count}")
    c3.metric("💸 Lost Payment", f"${lost_payment:,.2f}")
    c1, c2 = st.columns(2)
    c1.metric("⏳ Awaiting Decision", f"{awaiting_count}")
    c2.metric("⏳ Payment Pending Decision", f"${awaiting_payment:,.2f}")


def _chargeback_for_period(month=None, year=None, start_year=None, end_year=None) -> pd.DataFrame:
    data = load_chargeback()
    if data.empty or "_Chargeback_dt" not in data.columns:
        return data
    dt = data["_Chargeback_dt"]
    if start_year and end_year:
        mask = (dt.dt.year >= start_year) & (dt.dt.year <= end_year)
    elif year:
        mask = dt.dt.year == year
        if month:
            mask = mask & (dt.dt.month == month)
    elif month:
        mask = dt.dt.month == month
    else:
        return data
    return data[mask]


def render_chargeback_section(cb: pd.DataFrame, download_name: str) -> None:
    if cb is None or cb.empty:
        st.info("No chargeback data available for this period.")
        return
    render_chargeback_stats(cb, "### 📊 Chargeback Statistics")
    disp = cb.drop(columns=[c for c in ["_Chargeback_dt"] if c in cb.columns])
    disp.index = range(1, len(disp) + 1)
    st.markdown("### 📄 Chargeback Records")
    st.dataframe(disp)
    download_excel_button(disp, download_name)


# ---------------------------------------------------------------------------
# Shared UI helpers
# ---------------------------------------------------------------------------
def download_excel_button(data: pd.DataFrame, filename: str, key=None, label="📥 Download Excel"):
    buffer = io.BytesIO()
    data.to_excel(buffer, index=False)
    buffer.seek(0)
    st.download_button(
        label=label,
        data=buffer,
        file_name=filename,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        help="Click to download the Excel report",
        key=key,
    )


def render_multiple_platforms(data: pd.DataFrame):
    with st.expander("🧮 Clients with multiple platform publishing"):
        df = data.copy()
        if "Issues" in df.columns:
            df = df[~df["Issues"].isin(["Printing Only"])]
        platform_counts = df.groupby(["Name", "Book Name & Link"])["Platform"].nunique().reset_index(
            name="Platform_Count")
        platforms_per_client = df.groupby(["Name", "Book Name & Link"])["Platform"].unique().reset_index(
            name="Platforms")
        platform_stats = platform_counts.merge(platforms_per_client, how="left", on=["Name", "Book Name & Link"])
        platform_stats.index = range(1, len(platform_stats) + 1)
        st.dataframe(platform_stats)


def render_printing_country_stats(df: pd.DataFrame, country_name: str):
    if df.empty:
        st.warning(f"⚠️ No data found for {country_name} brands.")
        return
    total_orders = len(df)
    total_copies = df["No of Copies"].sum()
    total_cost = df["Order Cost"].sum()
    highest_cost = df["Order Cost"].max()
    lowest_cost = df["Order Cost"].min()
    highest_copies = df["No of Copies"].max()
    lowest_copies = df["No of Copies"].min()
    avg_cost_per_copy = round(total_cost / total_copies, 2) if total_copies else 0

    st.markdown(f"### 🌍 {country_name} Printing Summary")
    st.markdown(f"""
    - 🧾 **Total Orders:** {total_orders}
    - 📦 **Total Copies Printed:** `{total_copies}`
    - 💰 **Total Cost:** `${total_cost:,.2f}`
    - 📈 **Highest Order Cost:** `${highest_cost:,.2f}`
    - 📉 **Lowest Order Cost:** `${lowest_cost:,.2f}`
    - 🔢 **Highest Copies in One Order:** `{highest_copies}`
    - 🧮 **Lowest Copies in One Order:** `{lowest_copies}`
    - 💵 **Average Cost per Copy:** `${avg_cost_per_copy:,.2f}`
    """)
    brand_spending = df.groupby("Brand")["Order Cost"].sum().reset_index().sort_values(
        by="Order Cost", ascending=False)
    brand_spending["Order Cost"] = brand_spending["Order Cost"].map("${:,.2f}".format)
    brand_spending.index = range(1, len(brand_spending) + 1)
    brand_orders = df.groupby("Brand")["No of Copies"].sum().reset_index().sort_values(
        by="No of Copies", ascending=False)
    brand_orders.index = range(1, len(brand_orders) + 1)

    st.markdown(f"#### 💼 Brand-wise Spending for {country_name}")
    st.dataframe(brand_spending, width="stretch")
    st.markdown(f"#### 💼 Brand-wise Orders for {country_name}")
    st.dataframe(brand_orders, width="stretch")
    st.markdown("---")


def render_printing_period(data: pd.DataFrame, heading: str, download_filename: str,
                           stats_heading: str, download_key=None, monthly=None):
    st.markdown(heading)
    show_data = data.copy()
    show_data["Order Cost"] = show_data["Order Cost"].map("${:,.2f}".format)
    st.dataframe(show_data)

    total_copies = data["No of Copies"].sum()
    total_cost = data["Order Cost"].sum()
    highest_cost = data["Order Cost"].max()
    highest_copies = data["No of Copies"].max()
    lowest_cost = data["Order Cost"].min()
    lowest_copies = data["No of Copies"].min()
    average = round(total_cost / total_copies, 2) if total_copies else 0

    download_excel_button(data, download_filename, key=download_key)
    if monthly is not None and not monthly.empty:
        with st.expander("🖨 Monthly Printing Data"):
            st.dataframe(monthly)

    st.markdown("---")
    st.markdown(f"### {stats_heading}")
    st.markdown(f"""
    - 🧾 **Total Orders:** {len(data)}
    - 📦 **Total Copies Printed:** `{total_copies}`
    - 💰 **Total Cost:** `${total_cost:,.2f}`
    - 📈 **Highest Order Cost:** `${highest_cost:,.2f}`
    - 📉 **Lowest Order Cost:** `${lowest_cost:,.2f}`
    - 🔢 **Highest Copies in One Order:** `{highest_copies}`
    - 🧮 **Lowest Copies in One Order:** `{lowest_copies}`
    - 💵 **Average Cost per Copy:** `${average:,.2f}`
    """)

    usa_data = data[data["Brand"].isin(USA_PRINTING_BRANDS)]
    uk_data = data[data["Brand"].isin(UK_BRANDS)]
    usa_col, uk_col = st.columns(2)
    with usa_col:
        render_printing_country_stats(usa_data, "USA 🦅")
    with uk_col:
        render_printing_country_stats(uk_data, "UK ☕")


def render_copyright_period(data: pd.DataFrame, approved: int, rejected: int,
                            heading: str, download_filename: str = None,
                            download_key=None, no_data_msg: str = ""):
    if data.empty:
        st.warning(no_data_msg)
        return
    st.markdown(heading)
    st.dataframe(data)
    if download_filename:
        download_excel_button(data, download_filename, key=download_key)

    total_titles = len(data)
    country_counts = data["Country"].value_counts()
    country_usa = country_counts.get("USA", 0)
    country_uk = country_counts.get("UK", 0)
    country_canada = country_counts.get("Canada", 0)
    total_cost = (country_usa * COPYRIGHT_RATES["USA"]) + (country_canada * COPYRIGHT_RATES["Canada"]) + \
                 (country_uk * COPYRIGHT_RATES["UK"])

    st.markdown("---")
    st.markdown("### 📊 Summary Statistics (All Data)")
    st.markdown(f"""
    - 🧾 **Total Titles:** `{total_titles}`
    - 💵 **Total Cost:** `${total_cost}`
    - ✅ **Approved:** `{approved}` ({approved / total_titles:.1%})
    - ❌ **Rejected:** `{rejected}` ({rejected / total_titles:.1%})
    - 🦅 **USA:** `{country_usa}`
    - 🍁 **Canada:** `{country_canada}`
    - ☕ **UK:** `{country_uk}`
    """)


def render_review_details_view(*, choice: str, data: pd.DataFrame, pm_list: list,
                               attained_loader, negative_loader, data_title: str,
                               period_label: str, summary_title: str, excel_filename: str,
                               download_key=None, show_yearly=False, show_merged_yearly=False,
                               show_attained_yearly=False, show_negative_yearly=False):
    """Shared renderer for the Yearly / Start-to-Year / Filtered tabs in View Data."""
    data_rm_dupes = data.copy()
    if "Name" in data_rm_dupes.columns:
        data_rm_dupes = data_rm_dupes.drop_duplicates(subset=["Name"], keep="first")

    reviews_per_pm = safe_concat([df for df in (attained_loader(pm) for pm in pm_list) if not df.empty])
    reviews_n_pm = safe_concat([df for df in (negative_loader(pm) for pm in pm_list) if not df.empty])

    attained_pm = pd.DataFrame(columns=["Project Manager", "Attained Reviews"])
    negative_pm = pd.DataFrame(columns=["Project Manager", "Negative Reviews"])
    if not reviews_per_pm.empty:
        attained_pm = reviews_per_pm.groupby("Project Manager")["Trustpilot Review"].count().reset_index()
        attained_pm.columns = ["Project Manager", "Attained Reviews"]
        attained_pm = attained_pm.sort_values(by="Attained Reviews", ascending=False)
        attained_pm.index = range(1, len(attained_pm) + 1)
    if not reviews_n_pm.empty:
        negative_pm = reviews_n_pm.groupby("Project Manager")["Trustpilot Review"].count().reset_index()
        negative_pm.columns = ["Project Manager", "Negative Reviews"]
        negative_pm = negative_pm.sort_values(by="Negative Reviews", ascending=False)
        negative_pm.index = range(1, len(negative_pm) + 1)

    total_attained = attained_pm["Attained Reviews"].sum() if not attained_pm.empty else 0
    total_negative = negative_pm["Negative Reviews"].sum() if not negative_pm.empty else 0

    attained_details_total = pd.DataFrame(columns=REVIEW_DETAIL_COLUMNS)
    negative_details_total = pd.DataFrame(columns=REVIEW_DETAIL_COLUMNS)
    if not reviews_per_pm.empty:
        review_details_total = reviews_per_pm.sort_values(by="Project Manager", ascending=True)
        review_details_total["Trustpilot Review Date"] = pd.to_datetime(
            review_details_total["Trustpilot Review Date"], errors="coerce"
        ).dt.strftime(DATE_FORMAT)
        attained_details_total = review_details_total[
            review_details_total["Trustpilot Review"] == "Attained"][REVIEW_DETAIL_COLUMNS].copy()
        attained_details_total.index = range(1, len(attained_details_total) + 1)
    if not reviews_n_pm.empty:
        review_details_negative = reviews_n_pm.sort_values(by="Project Manager", ascending=True)
        review_details_negative["Trustpilot Review Date"] = pd.to_datetime(
            review_details_negative["Trustpilot Review Date"], errors="coerce"
        ).dt.strftime(DATE_FORMAT)
        negative_details_total = review_details_negative[
            review_details_negative["Trustpilot Review"] == "Negative"][REVIEW_DETAIL_COLUMNS].copy()
        negative_details_total.index = range(1, len(negative_details_total) + 1)

    attained_reviews_per_month = _monthly_review_table(attained_details_total, "Total Attained Reviews")
    negative_reviews_per_month = _monthly_review_table(negative_details_total, "Total Negative Reviews")

    st.markdown(data_title)
    st.dataframe(data)
    render_multiple_platforms(data)
    download_excel_button(data, excel_filename, key=download_key)

    brands = data_rm_dupes["Brand"].value_counts()
    platforms = data["Platform"].value_counts()
    publishing = data_rm_dupes["Status"].value_counts()

    filtered_data = data_rm_dupes[data_rm_dupes["Brand"].isin(REVIEW_BRANDS)]
    pending_sent_details = filtered_data[(filtered_data["Trustpilot Review"] == "Sent") | (
            filtered_data["Trustpilot Review"] == "Pending")]
    review_counts = filtered_data["Trustpilot Review"].value_counts()
    sent = review_counts.get("Sent", 0)
    pending = review_counts.get("Pending", 0)
    attained = total_attained
    negative = total_negative
    total_reviews = sent + pending + attained + negative
    percentage = round((attained / total_reviews * 100), 1) if total_reviews > 0 else 0

    unique_clients_count_per_pm = data_rm_dupes.groupby('Project Manager')['Name'].nunique().reset_index()
    unique_clients_count_per_pm.columns = ['Project Manager', 'Unique Clients']
    unique_clients_count_per_pm.index = range(1, len(unique_clients_count_per_pm) + 1)
    clients_list = data_rm_dupes.groupby('Project Manager')["Name"].apply(list).reset_index(name="Clients")
    merged_df = unique_clients_count_per_pm.merge(clients_list, on='Project Manager', how='left')
    merged_df.index = range(1, len(merged_df) + 1)
    total_unique_clients = data['Name'].nunique()
    Issues = data_rm_dupes["Issues"].value_counts()

    data_rm_dupes2 = data.copy().drop_duplicates(["Name"], keep="first")

    col1, col2 = st.columns(2)
    with col1:
        st.markdown("---")
        st.markdown(summary_title)
        st.markdown(f"""
        - 🧾 **Total Entries:** `{len(data)}`
        - 👥 **Total Unique Clients:** `{total_unique_clients}`
        - 🗳️ **Total Trustpilot Reviews:** `{total_reviews}`
        - 🟢 **'Attained' Reviews:** `{attained}`
        - 🔴 **'Negative' Reviews:** `{negative}`
        - 📈 **Attainment Rate:** `{percentage}%`
        - 📉 **Negative Rate:** `{round((negative / total_reviews) * 100, 1) if total_reviews > 0 else 0}%`
        - 💫 **Self Publishing:** `{Issues.get("Self Publishing", 0)}`
        - 🖨 **Printing Only:** `{Issues.get("Printing Only", 0)}`

        **Brands**
        - 📘 **BookMarketeers:** `{brands.get("BookMarketeers", "N/A")}`
        - 📘 **Aurora Writers:** `{brands.get("Aurora Writers", "N/A")}`
        - 📙 **Writers Clique:** `{brands.get("Writers Clique", "N/A")}`
        - 📕 **KDP:** `{brands.get("KDP", "N/A")}`
        - 📔 **Authors Solution:** `{brands.get("Authors Solution", "N/A")}`
        - 📘 **Book Publication:** `{brands.get("Book Publication", "N/A")}`
        - 📘 **Books Publisher:** `{brands.get("Books Publisher", "N/A")}`

        **Platforms**
        - 🅰 **Amazon:** `{platforms.get("Amazon", "N/A")}`
        - 📔 **Barnes & Noble:** `{platforms.get("Barnes & Noble", "N/A")}`
        - ⚡ **Ingram Spark:** `{platforms.get("Ingram Spark", "N/A")}`
        - 📚 **Kobo:** `{platforms.get("Kobo", "N/A")}`
        - 📚 **Draft2Digital:** `{platforms.get("Draft2Digital", "N/A")}`
        - 📚 **LULU:** `{platforms.get("LULU", "N/A")}`
        - 🔉 **Findaway Voices:** `{platforms.get("FAV", "N/A")}`
        - 🔉 **ACX:** `{platforms.get("ACX", "N/A")}`
        """)
        data_rm_dupes.index = range(1, len(data_rm_dupes) + 1)

        with st.expander(f"🤵🏻 Clients List {period_label}"):
            st.dataframe(data_rm_dupes)
        with st.expander("🤵🏻🤵🏻 Publishing Per Month"):
            data_month = data_rm_dupes.copy()
            data_month["Publishing Date"] = pd.to_datetime(data_month["Publishing Date"], errors="coerce")
            data_month["Month"] = data_month["Publishing Date"].dt.strftime("%B %Y")
            data_month["Month_Sort"] = data_month["Publishing Date"].dt.to_period("M").dt.to_timestamp()

            unique_clients_count_per_month = (
                data_month.groupby(["Month", "Month_Sort"])["Name"].nunique().reset_index(name="Total Published")
            )
            clients_list_per_month = (
                data_month.groupby(["Month", "Month_Sort"])["Name"].apply(list).reset_index(name="Clients")
            )
            publishing_per_month = unique_clients_count_per_month.merge(
                clients_list_per_month, on=["Month", "Month_Sort"], how="left"
            )
            publishing_per_month = publishing_per_month.sort_values("Month_Sort").drop(columns="Month_Sort")
            publishing_per_month.index = range(1, len(publishing_per_month) + 1)
            st.dataframe(publishing_per_month)

            pm_unique_clients_per_month = (
                data_month.groupby(["Month", "Project Manager"])["Name"].nunique().reset_index(name="Total Published")
            )
            pm_unique_clients_per_month_distribution = (
                data_month.groupby(["Month", "Project Manager"])["Name"].apply(list).reset_index(name="Clients")
            )
            merged_pm_client_distribution = pm_unique_clients_per_month.merge(
                pm_unique_clients_per_month_distribution, on=["Month", "Project Manager"], how="left")
            merged_pm_client_distribution.index = range(1, len(merged_pm_client_distribution) + 1)
            st.dataframe(merged_pm_client_distribution)

            if show_yearly:
                yearly_data = data_rm_dupes.copy()
                yearly_data["Publishing Date"] = pd.to_datetime(yearly_data["Publishing Date"], errors="coerce")
                yearly_data["Year"] = yearly_data["Publishing Date"].dt.to_period("Y").dt.strftime("%Y")
                unique_clients_count_per_year = (
                    yearly_data.groupby("Year")["Name"].nunique().reset_index()
                )
                unique_clients_count_per_year.columns = ["Year", "Total Published"]
                clients_list_per_year = (
                    yearly_data.groupby("Year")["Name"].apply(list).reset_index(name="Clients")
                )
                publishing_per_year = unique_clients_count_per_year.merge(
                    clients_list_per_year, on="Year", how="left"
                )
                publishing_per_year = publishing_per_year.sort_values(by="Total Published", ascending=False)
                publishing_per_year.index = range(1, len(publishing_per_year) + 1)
                st.dataframe(publishing_per_year)

        with st.expander(f"📈 Publishing Stats {period_label}"):
            publishing_stats = data_rm_dupes2.groupby('Publishing Date')["Name"].apply(list).reset_index(
                name="Clients")
            publishing_counts = data_rm_dupes2.groupby('Publishing Date')["Name"].count().reset_index(name="Counts")
            publishing_merged = publishing_counts.merge(publishing_stats, on='Publishing Date', how='left')
            publishing_merged.index = range(1, len(publishing_merged) + 1)
            st.dataframe(publishing_merged)
        with st.expander(f"💫 Self Publishing List {period_label}"):
            self_publishing_df = data_rm_dupes2[data_rm_dupes2["Issues"] == "Self Publishing"]
            self_publishing_df.index = range(1, len(self_publishing_df) + 1)
            st.dataframe(self_publishing_df)
        with st.expander(f"🖨 Printing Only List {period_label}"):
            printing_only_df = data_rm_dupes2[data_rm_dupes2["Issues"] == "Printing Only"]
            printing_only_df.index = range(1, len(printing_only_df) + 1)
            st.dataframe(printing_only_df)
        with st.expander("🟢 Attained Reviews Per Month"):
            st.dataframe(attained_reviews_per_month)
            if show_attained_yearly:
                df = attained_reviews_per_month.copy()
                df["Month"] = pd.to_datetime(df["Month"])
                df["Year"] = df["Month"].dt.year
                yearly_total = df.groupby("Year")["Total Attained Reviews"].sum()
                st.dataframe(yearly_total)
            if show_merged_yearly:
                df2 = attained_details_total.copy()
                df2["Trustpilot Review Date"] = pd.to_datetime(df2["Trustpilot Review Date"], errors="coerce")
                df2["Year"] = df2["Trustpilot Review Date"].dt.year
                yearly_names = df2.groupby("Year")["Name"].apply(list).reset_index(name="Clients")
                yearly_total_df = df.groupby("Year")["Total Attained Reviews"].sum().reset_index(name="Total Reviews")
                merged_yearly = yearly_total_df.merge(yearly_names, how="left", on="Year")
                merged_yearly.index = range(1, len(merged_yearly) + 1)
                st.dataframe(merged_yearly)
        with st.expander("🔴 Negative Reviews Per Month"):
            st.dataframe(negative_reviews_per_month)
            if show_negative_yearly:
                df = negative_reviews_per_month.copy()
                df["Month"] = pd.to_datetime(df["Month"])
                df["Year"] = df["Month"].dt.year
                yearly_total = df.groupby("Year")["Total Negative Reviews"].sum()
                st.dataframe(yearly_total)

    with col2:
        st.markdown("---")
        st.markdown("#### 🔍 Review & Publishing Status")
        st.markdown(f"""
        - 📝 **Sent**: `{sent}`
        - 📝 **Pending**: `{pending}`
        - 📝 **Attained**: `{attained}`
        """)
        st.markdown("**Publishing Status**")
        for status_type, count_s in publishing.items():
            st.markdown(f"- 📘 **{status_type}**: `{count_s}`")
        with st.expander("📊 View Clients Per PM Data"):
            st.dataframe(merged_df)
        with st.expander("❓ Pending & Sent Reviews"):
            pending_sent_details = pending_sent_details[
                ["Name", "Brand", "Project Manager", "Trustpilot Review", "Status"]]
            pending_sent_details.index = range(1, len(pending_sent_details) + 1)
            st.dataframe(pending_sent_details)
            breakdown_pending_sent = pending_sent_details["Trustpilot Review"].value_counts()
            st.dataframe(breakdown_pending_sent)
        with st.expander("👏 Attained Reviews Per PM"):
            st.dataframe(attained_pm)
            st.dataframe(attained_details_total)
            attained_count = attained_details_total.groupby("Project Manager").size().reset_index(name="Count")
            attained_clients = attained_details_total.groupby("Project Manager")["Name"].apply(
                list).reset_index(name="Clients")
            merged_attained = attained_count.merge(attained_clients, on="Project Manager", how="left")
            merged_attained = merged_attained.sort_values(by="Count", ascending=False)
            merged_attained.index = range(1, len(merged_attained) + 1)
            st.dataframe(merged_attained)
            st.dataframe(attained_details_total["Status"].value_counts())
        with st.expander("🏷️ Reviews Per Brand"):
            attained_brands = attained_details_total["Brand"].value_counts()
            st.dataframe(attained_brands)
        with st.expander("❌ Negative Reviews Per PM"):
            st.dataframe(negative_pm)
            st.dataframe(negative_details_total)
            st.dataframe(negative_details_total["Status"].value_counts())

    st.markdown("---")


# ---------------------------------------------------------------------------
# Summary report renderers
# ---------------------------------------------------------------------------
def render_month_summary_report(data: dict, title: str, excel_filename: str, pdf_data, pdf_filename):
    """Renders the monthly 'Summary' report screen."""
    usa_review_data = data["usa_review"]
    uk_review_data = data["uk_review"]
    usa_brands = data["usa_brands"]
    uk_brands = data["uk_brands"]
    usa_platforms = data["usa_platforms"]
    uk_platforms = data["uk_platforms"]
    printing_stats = data["printing_stats"]
    copyright_stats = data["copyright_stats"]
    a_plus = data["a_plus_count"]
    total_unique_clients = data["total_unique_clients"]
    combined = data["combined"]
    attained_reviews_per_pm = data["attained_reviews_per_pm"]
    attained_df = data["attained_details"]
    pending_sent_details = data["pending_sent_details"]
    negative_reviews_per_pm = data["negative_reviews_per_pm"]
    negative_details = data["negative_details"]
    Issues_usa = data["Issues_usa"]
    Issues_uk = data["Issues_uk"]

    usa_total = sum(usa_review_data.values())
    usa_attained = usa_review_data["Attained"] if "Attained" in usa_review_data else 0
    usa_attained_pct = (usa_attained / usa_total * 100) if usa_total > 0 else 0

    uk_total = sum(uk_review_data.values())
    uk_attained = uk_review_data["Attained"] if "Attained" in uk_review_data else 0
    uk_attained_pct = (uk_attained / uk_total * 100) if uk_total > 0 else 0

    combined_total = usa_total + uk_total
    combined_attained = usa_attained + uk_attained
    combined_attained_pct = (combined_attained / combined_total * 100) if combined_total > 0 else 0

    st.header(title)
    st.divider()
    st.markdown('<h2 class="section-header">📝 Review Analytics</h2>', unsafe_allow_html=True)

    col1, col2 = st.columns(2)
    with col1:
        usa_pie = create_review_pie_chart(usa_review_data, "USA Trustpilot Reviews")
        if usa_pie:
            st.plotly_chart(usa_pie, width="stretch", key="usa_pie")
        st.subheader("🇺🇸 USA Reviews")
        st.metric("📊 Total Reviews", usa_total)
        st.metric("🟢 Total Attained", usa_attained)
        st.metric("🔴 Total Negative", usa_review_data.get("Negative", 0))
        st.metric("🎯 Attained Percentage", f"{usa_attained_pct:.1f}%")
        st.metric("💫 Self Published", Issues_usa.get("Self Publishing", 0))
        st.metric("🖨 Printing Only", Issues_usa.get("Printing Only", 0))
        st.metric("👥 Total Unique", total_unique_clients)
        unique_clients_count_per_pm = combined.groupby('Project Manager')['Name'].nunique().reset_index()
        unique_clients_count_per_pm.columns = ['Project Manager', 'Unique Clients']
        unique_clients_count_per_pm.index = range(1, len(unique_clients_count_per_pm) + 1)
        clients_list = combined.groupby('Project Manager')["Name"].apply(list).reset_index(name="Clients")
        merged_df = unique_clients_count_per_pm.merge(clients_list, on='Project Manager', how='left')
        merged_df.index = range(1, len(merged_df) + 1)
        with st.expander("🤵🏻 Total Clients"):
            st.dataframe(combined)
        download_excel_button(combined, excel_filename)

    with col2:
        uk_pie = create_review_pie_chart(uk_review_data, "UK Trustpilot Reviews")
        if uk_pie:
            st.plotly_chart(uk_pie, width="stretch", key="uk_pie")
        st.subheader("🇬🇧 UK Reviews")
        st.metric("📊 Total Reviews", uk_total)
        st.metric("🟢 Total Attained", uk_attained)
        st.metric("🔴 Total Negative", uk_review_data.get("Negative", 0))
        st.metric("🎯 Attained Percentage", f"{uk_attained_pct:.1f}%")
        st.metric("💫 Self Published", Issues_uk.get("Self Publishing", 0))
        st.metric("🖨 Printing Only", Issues_uk.get("Printing Only", 0))
        with st.expander("📊 View Clients Per PM Data"):
            st.dataframe(merged_df)
        with st.expander("❓ Pending & Sent Reviews"):
            st.dataframe(pending_sent_details)
            breakdown_pending_sent = pending_sent_details["Trustpilot Review"].value_counts()
            st.dataframe(breakdown_pending_sent)
        with st.expander("👏 Reviews Per PM"):
            st.dataframe(attained_reviews_per_pm)
            st.dataframe(attained_df)
            st.dataframe(attained_df["Status"].value_counts())
        with st.expander("🏷️ Reviews Per Brand"):
            attained_brands = attained_df["Brand"].value_counts()
            st.dataframe(attained_brands)
        with st.expander("❌ Negative Reviews Per PM"):
            st.dataframe(negative_reviews_per_pm)
            st.dataframe(negative_details)
            st.dataframe(negative_details["Status"].value_counts())

    st.subheader("📱 Platform Distribution")
    platform_chart = create_platform_comparison_chart(usa_platforms, uk_platforms)
    st.plotly_chart(platform_chart, width="stretch", key="platform_chart")

    st.subheader("🏷️ Brand Performance")
    brand_chart = create_brand_chart(usa_brands, uk_brands)
    st.plotly_chart(brand_chart, width="stretch", key="brand_chart")

    col1, col2 = st.columns(2)
    with col1:
        st.subheader("USA Brand Breakdown")
        usa_df = pd.DataFrame(list(usa_brands.items()), columns=['Brand', 'Count'])
        st.dataframe(usa_df, hide_index=True)
        st.markdown(f"""
        - 📊 **Total Count Across Brands:** `{usa_df["Count"].sum()}`
        """)
        st.subheader("USA Platform Breakdown")
        usa_platform_df = pd.DataFrame(list(usa_platforms.items()), columns=['Platform', 'Count'])
        st.dataframe(usa_platform_df, hide_index=True)
        st.markdown(f"""
        - 📊 **Total Count Across Platforms:** `{usa_platform_df["Count"].sum()}`
        """)
    with col2:
        st.subheader("UK Brand Breakdown")
        uk_df = pd.DataFrame(list(uk_brands.items()), columns=['Brand', 'Count'])
        st.dataframe(uk_df, hide_index=True)
        st.markdown(f"""
        - 📊 **Total Count Across Brands:** `{uk_df["Count"].sum()}`
        """)
        st.subheader("UK Platform Breakdown")
        uk_platform_df = pd.DataFrame(list(uk_platforms.items()), columns=['Platform', 'Count'])
        st.dataframe(uk_platform_df, hide_index=True)
        st.markdown(f"""
        - 📊 **Total Count Across Platforms:** `{uk_platform_df["Count"].sum()}`
        """)

    st.divider()
    st.markdown('<h2 class="section-header">🖨️ Printing Analytics</h2>', unsafe_allow_html=True)

    col1, col2, col3 = st.columns(3)
    with col1:
        st.subheader("📊 Volume Metrics")
        st.metric("Total Copies", f"{printing_stats['Total_copies']:,}")
        st.metric("Highest Copies", printing_stats['Highest_copies'])
        st.metric("Lowest Copies", printing_stats['Lowest_copies'])
    with col2:
        st.subheader("💰 Cost Metrics")
        st.metric("Total Cost", f"${printing_stats['Total_cost']:,.2f}")
        st.metric("Highest Cost", f"${printing_stats['Highest_cost']:.2f}")
        st.metric("Lowest Cost", f"${printing_stats['Lowest_cost']:.2f}")
    with col3:
        st.subheader("📈 Efficiency")
        st.metric("Average Cost per Copy", f"${printing_stats['Average']:.2f}")
        fig_gauge = go.Figure(go.Indicator(
            mode="gauge+number",
            value=printing_stats['Average'],
            domain={'x': [0, 1], 'y': [0, 1]},
            title={'text': "Avg Cost/Copy"},
            gauge={
                'axis': {'range': [None, 15]},
                'bar': {'color': "darkblue"},
                'steps': [
                    {'range': [0, 5], 'color': "lightgray"},
                    {'range': [5, 10], 'color': "gray"}],
                'threshold': {
                    'line': {'color': "red", 'width': 4},
                    'thickness': 0.75,
                    'value': 10}}))
        fig_gauge.update_layout(height=200)
        st.plotly_chart(fig_gauge, width="stretch")

    st.divider()
    st.markdown('<h2 class="section-header">©️ Copyright Analytics</h2>', unsafe_allow_html=True)

    col1, col2 = st.columns(2)
    with col1:
        st.subheader("📋 Copyright Summary")
        st.metric("Total Copyrights", copyright_stats['Total_copyrights'])
        st.metric("Total Cost", f"${copyright_stats['Total_cost_copyright']:,}")
        st.metric("Success Rate", f"{copyright_stats['result_count']}/{copyright_stats['Total_copyrights']}")
        success_rate = (
                copyright_stats['result_count'] / copyright_stats['Total_copyrights'] * 100) if \
            copyright_stats['Total_copyrights'] > 0 else 0
        st.metric("Success Percentage", f"{success_rate:.1f}%")
        st.metric("Rejection Rate", f"{copyright_stats['result_count_no']}/{copyright_stats['Total_copyrights']}")
        rejection_rate = (
                copyright_stats['result_count_no'] / copyright_stats['Total_copyrights'] * 100) if \
            copyright_stats['Total_copyrights'] > 0 else 0
        st.metric("Rejection Percentage", f"{rejection_rate:.1f}%")
    with col2:
        st.subheader("🌍 Country Distribution")
        copyright_countries = {
            'USA': copyright_stats['usa_copyrights'],
            'Canada': copyright_stats['canada_copyrights'],
            'UK': copyright_stats['uk']
        }
        fig_copyright = px.pie(
            values=list(copyright_countries.values()),
            names=list(copyright_countries.keys()),
            title="Copyright Applications by Country",
            color_discrete_sequence=["#23A0F8", "#d62728", "#F7E319"]
        )
        st.plotly_chart(fig_copyright, width="stretch", key="copyright_chart")
        cp1, cp2, cp3 = st.columns(3)
        with cp1:
            st.metric('Usa', copyright_stats['usa_copyrights'])
        with cp2:
            st.metric('Canada', copyright_stats['canada_copyrights'])
        with cp3:
            st.metric('UK', copyright_stats['uk'])

    st.divider()
    cola = st.columns(1)
    with cola[0]:
        st.subheader("🅰➕ Content")
        st.metric("A+ Count", f"{a_plus} Published")

    st.divider()
    st.markdown('<h2 class="section-header">💳 Chargeback Analytics</h2>', unsafe_allow_html=True)
    render_chargeback_section(data.get("chargeback"), f"Chargeback_{excel_filename}")

    st.divider()
    st.markdown('<h2 class="section-header">📈 Executive Summary</h2>', unsafe_allow_html=True)

    summary_col1, summary_col2, summary_col3 = st.columns(3)
    with summary_col1:
        st.markdown("### 📝 Reviews")
        st.write(f"• **Combined Reviews**: {combined_total}")
        st.write(f"• **Success Rate**: {combined_attained_pct:.1f}%")
        st.write(f"• **USA Attained**: {usa_attained}")
        st.write(f"• **UK Attained**: {uk_attained}")
    with summary_col2:
        st.markdown("### 🖨️ Printing")
        st.write(f"• **Total Copies**: {printing_stats['Total_copies']:,}")
        st.write(f"• **Total Cost**: ${printing_stats['Total_cost']:,.2f}")
        st.write(f"• **Cost Efficiency**: ${printing_stats['Average']:.2f}/copy")
    with summary_col3:
        st.markdown("### ©️ Copyright")
        st.write(f"• **Applications**: {copyright_stats['Total_copyrights']}")
        st.write(f"• **Success Rate**: {success_rate:.1f}%")
        st.write(f"• **Rejection Rate**: {rejection_rate:.1f}%")
        st.write(f"• **Total Cost**: ${copyright_stats['Total_cost_copyright']:,}")

    st.success(f"Summary report for {title} generated!")
    st.download_button(
        label="📥 Download PDF Report",
        data=pdf_data,
        file_name=pdf_filename,
        mime="application/pdf",
        help="Click to download the PDF report"
    )


def render_year_summary_report(data: dict, title: str, excel_filename: str, pdf_data, pdf_filename):
    """Renders the 'Year Summary' and 'Custom Summary' report screens."""
    usa_review_data = data["usa_review"]
    uk_review_data = data["uk_review"]
    usa_brands = data["usa_brands"]
    uk_brands = data["uk_brands"]
    usa_platforms = data["usa_platforms"]
    uk_platforms = data["uk_platforms"]
    printing_stats = data["printing_stats"]
    monthly_printing = data["monthly_printing"]
    copyright_stats = data["copyright_stats"]
    a_plus = data["a_plus_count"]
    total_unique_clients = data["total_unique_clients"]
    combined = data["combined"]
    attained_reviews_per_pm = data["attained_reviews_per_pm"]
    attained_df = data["attained_details"]
    merged_attained = data["merged_attained"]
    attained_reviews_per_month = data["attained_reviews_per_month"]
    pending_sent_details = data["pending_sent_details"]
    negative_reviews_per_pm = data["negative_reviews_per_pm"]
    negative_details = data["negative_details"]
    negative_reviews_per_month = data["negative_reviews_per_month"]
    publishing_per_month = data["combined_monthly"]
    Issues_usa = data["Issues_usa"]
    Issues_uk = data["Issues_uk"]

    usa_total = sum(usa_review_data.values())
    usa_attained = usa_review_data["Attained"] if "Attained" in usa_review_data else 0
    usa_attained_pct = (usa_attained / usa_total * 100) if usa_total > 0 else 0

    uk_total = sum(uk_review_data.values())
    uk_attained = uk_review_data["Attained"] if "Attained" in uk_review_data else 0
    uk_attained_pct = (uk_attained / uk_total * 100) if uk_total > 0 else 0

    combined_total = usa_total + uk_total
    combined_attained = usa_attained + uk_attained
    combined_attained_pct = (combined_attained / combined_total * 100) if combined_total > 0 else 0

    st.header(title)
    st.divider()
    st.markdown('<h2 class="section-header">📝 Review Analytics</h2>', unsafe_allow_html=True)

    col1, col2 = st.columns(2)
    with col1:
        usa_pie = create_review_pie_chart(usa_review_data, "USA Trustpilot Reviews")
        if usa_pie:
            st.plotly_chart(usa_pie, width="stretch", key="usa_pie")
        st.subheader("🇺🇸 USA Reviews")
        st.metric("🤵🏻 Total Clients", sum(usa_brands.values()))
        st.metric("📊 Total Reviews", usa_total)
        st.metric("🟢 Total Attained", usa_attained)
        st.metric("🔴 Total Negative", usa_review_data.get("Negative", 0))
        st.metric("🎯 Attained Percentage", f"{usa_attained_pct:.1f}%")
        st.metric("👥 Total Unique", total_unique_clients)
        st.metric("💫 Self Published", Issues_usa.get("Self Publishing", 0))
        st.metric("🖨 Printing Only", Issues_usa.get("Printing Only", 0))
        unique_clients_count_per_pm = combined.groupby('Project Manager')['Name'].nunique().reset_index()
        unique_clients_count_per_pm.columns = ['Project Manager', 'Unique Clients']
        unique_clients_count_per_pm.index = range(1, len(unique_clients_count_per_pm) + 1)
        clients_list = combined.groupby('Project Manager')["Name"].apply(list).reset_index(name="Clients")
        merged_df = unique_clients_count_per_pm.merge(clients_list, on='Project Manager', how='left')
        merged_df.index = range(1, len(merged_df) + 1)
        with st.expander("🤵🏻 Total Clients"):
            st.dataframe(combined)
        with st.expander("🤵🏻🤵🏻 Publishing Per Month"):
            st.dataframe(publishing_per_month)
        download_excel_button(combined, excel_filename)
        with st.expander("🟢 Attained Reviews Per Month"):
            st.dataframe(attained_reviews_per_month)
            df = attained_reviews_per_month.copy()
            df["Month"] = pd.to_datetime(df["Month"])
            df["Year"] = df["Month"].dt.year
            yearly_total = df.groupby("Year")["Total Attained Reviews"].sum()
            st.dataframe(yearly_total)
            usa_yearly = df.groupby("Year")["USA Attained Reviews"].sum()
            uk_yearly = df.groupby("Year")["UK Attained Reviews"].sum()
            st.dataframe(usa_yearly)
            st.dataframe(uk_yearly)
        with st.expander("🔴 Negative Reviews Per Month"):
            st.dataframe(negative_reviews_per_month)
            df = negative_reviews_per_month.copy()
            df["Month"] = pd.to_datetime(df["Month"])
            df["Year"] = df["Month"].dt.year
            yearly_total = df.groupby("Year")["Total Negative Reviews"].sum()
            st.dataframe(yearly_total)
            usa_yearly = df.groupby("Year")["USA Negative Reviews"].sum()
            uk_yearly = df.groupby("Year")["UK Negative Reviews"].sum()
            st.dataframe(usa_yearly)
            st.dataframe(uk_yearly)

    with col2:
        uk_pie = create_review_pie_chart(uk_review_data, "UK Trustpilot Reviews")
        if uk_pie:
            st.plotly_chart(uk_pie, width="stretch", key="uk_pie")
        st.subheader("🇬🇧 UK Reviews")
        st.metric("🤵🏻 Total Clients", sum(uk_brands.values()))
        st.metric("📊 Total Reviews", uk_total)
        st.metric("🟢 Total Attained", uk_attained)
        st.metric("🔴 Total Negative", uk_review_data.get("Negative", 0))
        st.metric("🎯 Attained Percentage", f"{uk_attained_pct:.1f}%")
        st.metric("💫 Self Published", Issues_uk.get("Self Publishing", 0))
        st.metric("🖨 Printing Only", Issues_uk.get("Printing Only", 0))
        with st.expander("📊 View Clients Per PM Data"):
            st.dataframe(merged_df)
        with st.expander("❓ Pending & Sent Reviews"):
            st.dataframe(pending_sent_details)
            breakdown_pending_sent = pending_sent_details["Trustpilot Review"].value_counts()
            st.dataframe(breakdown_pending_sent)
        with st.expander("👏 Reviews Per PM"):
            st.dataframe(attained_reviews_per_pm)
            st.dataframe(attained_df)
            st.dataframe(merged_attained)
            st.dataframe(attained_df["Status"].value_counts())
        with st.expander("🏷️ Reviews Per Brand"):
            attained_brands = attained_df["Brand"].value_counts()
            st.dataframe(attained_brands)
        with st.expander("❌ Negative Reviews Per PM"):
            st.dataframe(negative_reviews_per_pm)
            st.dataframe(negative_details)
            st.dataframe(negative_details["Status"].value_counts())

    st.subheader("📱 Platform Distribution")
    platform_chart = create_platform_comparison_chart(usa_platforms, uk_platforms)
    st.plotly_chart(platform_chart, width="stretch")

    st.subheader("🏷️ Brand Performance")
    brand_chart = create_brand_chart(usa_brands, uk_brands)
    st.plotly_chart(brand_chart, width="stretch", key="brand_chart")

    col1, col2 = st.columns(2)
    with col1:
        st.subheader("USA Brand Breakdown")
        usa_df = pd.DataFrame(list(usa_brands.items()), columns=['Brand', 'Count'])
        st.dataframe(usa_df, hide_index=True)
        st.markdown(f"""
        - 📊 **Total Count Across Brands:** `{usa_df["Count"].sum()}`
        """)
        st.subheader("USA Platform Breakdown")
        usa_platform_df = pd.DataFrame(list(usa_platforms.items()), columns=['Platform', 'Count'])
        st.dataframe(usa_platform_df, hide_index=True)
        st.markdown(f"""
        - 📊 **Total Count Across Platforms:** `{usa_platform_df["Count"].sum()}`
        """)
    with col2:
        st.subheader("UK Brand Breakdown")
        uk_df = pd.DataFrame(list(uk_brands.items()), columns=['Brand', 'Count'])
        st.dataframe(uk_df, hide_index=True)
        st.markdown(f"""
        - 📊 **Total Count Across Brands:** `{uk_df["Count"].sum()}`
        """)
        st.subheader("UK Platform Breakdown")
        uk_platform_df = pd.DataFrame(list(uk_platforms.items()), columns=['Platform', 'Count'])
        st.dataframe(uk_platform_df, hide_index=True)
        st.markdown(f"""
        - 📊 **Total Count Across Platforms:** `{uk_platform_df["Count"].sum()}`
        """)

    st.divider()
    st.markdown('<h2 class="section-header">🖨️ Printing Analytics</h2>', unsafe_allow_html=True)

    col1, col2, col3 = st.columns(3)
    with col1:
        st.subheader("📊 Volume Metrics")
        st.metric("Total Copies", f"{printing_stats['Total_copies']:,}")
        st.metric("Highest Copies", printing_stats['Highest_copies'])
        st.metric("Lowest Copies", printing_stats['Lowest_copies'])
    with col2:
        st.subheader("💰 Cost Metrics")
        st.metric("Total Cost", f"${printing_stats['Total_cost']:,.2f}")
        st.metric("Highest Cost", f"${printing_stats['Highest_cost']:.2f}")
        st.metric("Lowest Cost", f"${printing_stats['Lowest_cost']:.2f}")
    with col3:
        st.subheader("📈 Efficiency")
        st.metric("Average Cost per Copy", f"${printing_stats['Average']:.2f}")
        fig_gauge = go.Figure(go.Indicator(
            mode="gauge+number",
            value=printing_stats['Average'],
            domain={'x': [0, 1], 'y': [0, 1]},
            title={'text': "Avg Cost/Copy"},
            gauge={
                'axis': {'range': [None, 15]},
                'bar': {'color': "darkblue"},
                'steps': [
                    {'range': [0, 5], 'color': "lightgray"},
                    {'range': [5, 10], 'color': "gray"}],
                'threshold': {
                    'line': {'color': "red", 'width': 4},
                    'thickness': 0.75,
                    'value': 10}}))
        fig_gauge.update_layout(height=200)
        st.plotly_chart(fig_gauge, width="stretch")
    with st.expander("🖨 Monthly Printing Data"):
        st.dataframe(monthly_printing)

    st.divider()
    st.markdown('<h2 class="section-header">©️ Copyright Analytics</h2>', unsafe_allow_html=True)

    col1, col2 = st.columns(2)
    with col1:
        st.subheader("📋 Copyright Summary")
        st.metric("Total Copyrights", copyright_stats['Total_copyrights'])
        st.metric("Total Cost", f"${copyright_stats['Total_cost_copyright']:,}")
        st.metric("Success Rate", f"{copyright_stats['result_count']}/{copyright_stats['Total_copyrights']}")
        success_rate = (
                copyright_stats['result_count'] / copyright_stats['Total_copyrights'] * 100) if \
            copyright_stats['Total_copyrights'] > 0 else 0
        st.metric("Success Percentage", f"{success_rate:.1f}%")
        st.metric("Rejection Rate", f"{copyright_stats['result_count_no']}/{copyright_stats['Total_copyrights']}")
        rejection_rate = (
                copyright_stats['result_count_no'] / copyright_stats['Total_copyrights'] * 100) if \
            copyright_stats['Total_copyrights'] > 0 else 0
        st.metric("Rejection Percentage", f"{rejection_rate:.1f}%")
    with col2:
        st.subheader("🌍 Country Distribution")
        copyright_countries = {
            'USA': copyright_stats['usa_copyrights'],
            'Canada': copyright_stats['canada_copyrights'],
            'UK': copyright_stats['uk']
        }
        fig_copyright = px.pie(
            values=list(copyright_countries.values()),
            names=list(copyright_countries.keys()),
            title="Copyright Applications by Country",
            color_discrete_sequence=["#23A0F8", "#d62728", "#F7E319"]
        )
        st.plotly_chart(fig_copyright, width="stretch", key="copyright_chart")
        cp1, cp2, cp3 = st.columns(3)
        with cp1:
            st.metric('Usa', copyright_stats['usa_copyrights'])
        with cp2:
            st.metric('Canada', copyright_stats['canada_copyrights'])
        with cp3:
            st.metric('UK', copyright_stats['uk'])

    st.divider()
    cola = st.columns(1)
    with cola[0]:
        st.subheader("🅰➕ Content")
        st.metric("A+ Count", f"{a_plus} Published")

    st.divider()
    st.markdown('<h2 class="section-header">💳 Chargeback Analytics</h2>', unsafe_allow_html=True)
    render_chargeback_section(data.get("chargeback"), f"Chargeback_{excel_filename}")

    st.divider()
    st.markdown('<h2 class="section-header">📈 Executive Summary</h2>', unsafe_allow_html=True)

    summary_col1, summary_col2, summary_col3 = st.columns(3)
    with summary_col1:
        st.markdown("### 📝 Reviews")
        st.write(f"• **Combined Reviews**: {combined_total}")
        st.write(f"• **Success Rate**: {combined_attained_pct:.1f}%")
        st.write(f"• **USA Attained**: {usa_attained}")
        st.write(f"• **UK Attained**: {uk_attained}")
    with summary_col2:
        st.markdown("### 🖨️ Printing")
        st.write(f"• **Total Copies**: {printing_stats['Total_copies']:,}")
        st.write(f"• **Total Cost**: ${printing_stats['Total_cost']:,.2f}")
        st.write(f"• **Cost Efficiency**: ${printing_stats['Average']:.2f}/copy")
    with summary_col3:
        st.markdown("### ©️ Copyright")
        st.write(f"• **Applications**: {copyright_stats['Total_copyrights']}")
        st.write(f"• **Success Rate**: {success_rate:.1f}%")
        st.write(f"• **Rejection Rate**: {rejection_rate:.1f}%")
        st.write(f"• **Total Cost**: ${copyright_stats['Total_cost_copyright']:,}")

    st.success(f"Summary report for {title} generated!")
    st.download_button(
        label="📥 Download PDF Report",
        data=pdf_data,
        file_name=pdf_filename,
        mime="application/pdf",
        help="Click to download the PDF report"
    )


# ---------------------------------------------------------------------------
# Main app
# ---------------------------------------------------------------------------
def main() -> None:
    with st.container():
        st.title("📊 Blink Digitally Publishing Dashboard")
        if st.button("🔃 Fetch Latest"):
            st.cache_data.clear()
            st.success("Fetched new data")
        action = st.selectbox("What would you like to do?",
                              ["View Data", "Printing", "Copyright", "Generate Similarity",
                               "Summary", "Year Summary", "Custom Summary", "Reviews", "Sales", "ISBN", "Chargeback"],
                              index=None,
                              placeholder="Select Action")

        selected_month = None
        selected_month_number = None
        number = None
        choice = None

        if action in ["View Data"]:
            choice = st.selectbox("Select Data To View", ["USA", "UK"], index=None,
                                  placeholder="Select Data to View")
        if action in ["View Data"]:
            selected_month = st.selectbox(
                "Select Month", month_list, index=current_month - 1, placeholder="Select Month"
            )
            selected_month_number = month_list.index(selected_month) + 1 if selected_month else None
        if action in ["Year Summary", "View Data", "Reviews"]:
            number = st.number_input("Enter Year", min_value=int(get_min_year()), max_value=current_year,
                                     value=current_year, step=1)

        if action == "View Data" and choice and selected_month and number:
            tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs(
                ["Monthly", "Yearly", "Start to Year", "Filter", "Search", "By Brand"])

            sheet_name = {"UK": sheet_uk, "USA": sheet_usa}.get(choice)

            with tab1:
                st.subheader(f"📂 Viewing Data for {choice} - {selected_month} {number}")
                if sheet_name:
                    data = load_data(sheet_name, selected_month_number, number)
                    if not data.empty:
                        data_rm_dupes = data.copy()
                        if "Name" in data_rm_dupes.columns:
                            data_rm_dupes = data_rm_dupes.drop_duplicates(subset=["Name"], keep="first")
                        review_data = load_reviews(sheet_name, number, selected_month_number)

                        attained_reviews_per_pm = pd.DataFrame(columns=["Project Manager", "Attained Reviews"])
                        attained_details = pd.DataFrame(columns=REVIEW_DETAIL_COLUMNS)
                        review_details_df = pd.DataFrame()
                        if not review_data.empty:
                            attained_reviews_per_pm = review_data[
                                review_data["Trustpilot Review"] == "Attained"
                            ].groupby("Project Manager")["Trustpilot Review"].count().reset_index()
                            review_details_df = review_data.sort_values(by="Project Manager", ascending=True)
                            review_details_df["Trustpilot Review Date"] = pd.to_datetime(
                                review_details_df["Trustpilot Review Date"], errors="coerce"
                            ).dt.strftime(DATE_FORMAT)
                            attained_reviews_per_pm.columns = ["Project Manager", "Attained Reviews"]
                            attained_reviews_per_pm = attained_reviews_per_pm.sort_values(
                                by="Attained Reviews", ascending=False)
                            attained_reviews_per_pm.index = range(1, len(attained_reviews_per_pm) + 1)
                            attained_details = review_details_df[
                                review_details_df["Trustpilot Review"] == "Attained"
                            ][REVIEW_DETAIL_COLUMNS].copy()
                            attained_details.index = range(1, len(attained_details) + 1)

                        negative_reviews_per_pm = pd.DataFrame(columns=["Project Manager", "Negative Reviews"])
                        negative_details = pd.DataFrame(columns=REVIEW_DETAIL_COLUMNS)
                        if not review_data.empty:
                            negative_reviews_per_pm = review_data[
                                review_data["Trustpilot Review"] == "Negative"
                            ].groupby("Project Manager")["Trustpilot Review"].count().reset_index()
                            negative_reviews_per_pm.columns = ["Project Manager", "Negative Reviews"]
                            negative_reviews_per_pm = negative_reviews_per_pm.sort_values(
                                by="Negative Reviews", ascending=False)
                            negative_reviews_per_pm.index = range(1, len(negative_reviews_per_pm) + 1)
                            negative_details = review_details_df[
                                review_details_df["Trustpilot Review"] == "Negative"
                            ][REVIEW_DETAIL_COLUMNS].copy()
                            negative_details.index = range(1, len(negative_details) + 1)

                        st.markdown("### 📄 Detailed Entry Data")
                        st.dataframe(data)
                        render_multiple_platforms(data)
                        download_excel_button(data, f"{choice}_{selected_month}_{number}.xlsx")

                        brands = data_rm_dupes["Brand"].value_counts()
                        writers_clique = brands.get("Writers Clique", "N/A")
                        bookmarketeers = brands.get("BookMarketeers", "N/A")
                        aurora_writers = brands.get("Aurora Writers", "N/A")
                        kdp = brands.get("KDP", "N/A")
                        authors_solution = brands.get("Authors Solution", "N/A")
                        book_publication = brands.get("Book Publication", "N/A")
                        books_publisher = brands.get("Books Publisher", "N/A")

                        platforms = data["Platform"].value_counts()
                        amazon = platforms.get("Amazon", "N/A")
                        bn = platforms.get("Barnes & Noble", "N/A")
                        ingram = platforms.get("Ingram Spark", "N/A")
                        fav = platforms.get("FAV", "N/A")
                        acx = platforms.get("ACX", "N/A")
                        kobo = platforms.get("Kobo", "N/A")
                        d2d = platforms.get("Draft2Digital", "N/A")
                        lulu = platforms.get("LULU", "N/A")

                        filtered_data = data_rm_dupes[data_rm_dupes["Brand"].isin(REVIEW_BRANDS)]
                        sent = filtered_data["Trustpilot Review"].value_counts().get("Sent", 0)
                        pending = filtered_data["Trustpilot Review"].value_counts().get("Pending", 0)
                        pending_sent_details = filtered_data[(filtered_data["Trustpilot Review"] == "Sent") | (
                                filtered_data["Trustpilot Review"] == "Pending")]
                        review = {
                            "Sent": sent,
                            "Pending": pending,
                            "Attained": attained_reviews_per_pm["Attained Reviews"].sum(),
                            "Negative": negative_reviews_per_pm["Negative Reviews"].sum()
                        }
                        publishing = data_rm_dupes["Status"].value_counts()
                        total_reviews = sum(review.values())
                        attained = attained_reviews_per_pm["Attained Reviews"].sum()
                        negative = negative_reviews_per_pm["Negative Reviews"].sum()
                        percentage = round((attained / total_reviews * 100), 1) if total_reviews > 0 else 0

                        unique_clients_count_per_pm = data_rm_dupes.groupby('Project Manager')[
                            'Name'].nunique().reset_index()
                        unique_clients_count_per_pm.columns = ['Project Manager', 'Unique Clients']
                        unique_clients_count_per_pm.index = range(1, len(unique_clients_count_per_pm) + 1)
                        total_unique_clients = data['Name'].nunique()
                        clients_list = data_rm_dupes.groupby('Project Manager')["Name"].apply(list).reset_index(
                            name="Clients")
                        merged_df = unique_clients_count_per_pm.merge(clients_list, on='Project Manager',
                                                                      how='left')
                        merged_df.index = range(1, len(merged_df) + 1)
                        Issues = data_rm_dupes["Issues"].value_counts()

                        data_rm_dupes2 = data.copy().drop_duplicates(["Name"], keep="first")

                        col1, col2 = st.columns(2)
                        with col1:
                            st.markdown("---")
                            st.markdown("### ⭐ Trustpilot Review Summary")
                            st.markdown(f"""
                                        - 🧾 **Total Entries:** `{len(data)}`
                                        - 👥 **Total Unique Clients:** `{total_unique_clients}`
                                        - 🗳️ **Total Trustpilot Reviews:** `{total_reviews}`
                                        - 🟢 **'Attained' Reviews:** `{attained}`
                                        - 🔴 **'Negative' Reviews:** `{negative}`
                                        - 📈 **Attainment Rate:** `{percentage}%`
                                        - 📉 **Negative Rate:** `{round((negative / total_reviews) * 100, 1) if total_reviews > 0 else 0}%`
                                        - 💫 **Self Publishing:** `{Issues.get("Self Publishing", 0)}`
                                        - 🖨 **Printing Only:** `{Issues.get("Printing Only", 0)}`

                                        **Brands**
                                        - 📘 **BookMarketeers:** `{bookmarketeers}`
                                        - 📘 **Aurora Writers:** `{aurora_writers}`
                                        - 📙 **Writers Clique:** `{writers_clique}`
                                        - 📕 **KDP:** `{kdp}`
                                        - 📔 **Authors Solution:** `{authors_solution}`
                                        - 📘 **Book Publication:** `{book_publication}`
                                        - 📘 **Books Publisher:** `{books_publisher}`

                                        **Platforms**
                                        - 🅰 **Amazon:** `{amazon}`
                                        - 📔 **Barnes & Noble:** `{bn}`
                                        - ⚡ **Ingram Spark:** `{ingram}`
                                        - 📚 **Kobo:** `{kobo}`
                                        - 📚 **Draft2Digital:** `{d2d}`
                                        - 📚 **LULU:** `{lulu}`
                                        - 🔉 **Findaway Voices:** `{fav}`
                                        - 🔉 **ACX:** `{acx}`
                                        """)
                            data_rm_dupes.index = range(1, len(data_rm_dupes) + 1)
                            with st.expander(f"🤵🏻 Clients List {choice} {selected_month} {number}"):
                                st.dataframe(data_rm_dupes)
                            with st.expander(f"📈 Publishing Stats {choice} {selected_month} {number}"):
                                publishing_stats = data_rm_dupes2.groupby('Publishing Date')["Name"].apply(
                                    list).reset_index(name="Clients")
                                publishing_counts = data_rm_dupes2.groupby('Publishing Date')[
                                    "Name"].count().reset_index(name="Counts")
                                publishing_merged = publishing_counts.merge(publishing_stats, on='Publishing Date',
                                                                            how='left')
                                publishing_merged.index = range(1, len(publishing_merged) + 1)
                                st.dataframe(publishing_merged)
                            with st.expander(f"💫 Self Publishing List {choice} {selected_month} {number}"):
                                self_publishing_df = data_rm_dupes2[data_rm_dupes2["Issues"] == "Self Publishing"]
                                self_publishing_df.index = range(1, len(self_publishing_df) + 1)
                                st.dataframe(self_publishing_df)
                            with st.expander(f"🖨 Printing Only List {choice} {selected_month} {number}"):
                                printing_only_df = data_rm_dupes2[data_rm_dupes2["Issues"] == "Printing Only"]
                                printing_only_df.index = range(1, len(printing_only_df) + 1)
                                st.dataframe(printing_only_df)
                        with col2:
                            st.markdown("---")
                            st.markdown("#### 🔍 Review & Publishing Status Breakdown")
                            for review_type, count in review.items():
                                st.markdown(f"- 📝 **{review_type}**: `{count}`")
                            for status_type, count_s in publishing.items():
                                st.markdown(f"- 📘 **{status_type}**: `{count_s}`")
                            with st.expander("📊 View Clients Per PM Data"):
                                st.dataframe(merged_df)
                            with st.expander("❓ Pending & Sent Reviews"):
                                pending_sent_details = pending_sent_details[
                                    ["Name", "Brand", "Project Manager", "Trustpilot Review", "Status"]]
                                pending_sent_details.index = range(1, len(pending_sent_details) + 1)
                                st.dataframe(pending_sent_details)
                                breakdown_pending_sent = pending_sent_details["Trustpilot Review"].value_counts()
                                st.dataframe(breakdown_pending_sent)
                            with st.expander("👏 Attained Reviews Per PM"):
                                st.dataframe(attained_reviews_per_pm)
                                st.dataframe(attained_details)
                                attained_count = attained_details.groupby("Project Manager").size().reset_index(
                                    name="Count")
                                attained_clients = attained_details.groupby("Project Manager")["Name"].apply(
                                    list).reset_index(name="Clients")
                                merged_attained = attained_count.merge(attained_clients, on="Project Manager",
                                                                       how="left")
                                merged_attained = merged_attained.sort_values(by="Count", ascending=False)
                                merged_attained.index = range(1, len(merged_attained) + 1)
                                st.dataframe(merged_attained)
                                st.dataframe(attained_details["Status"].value_counts())
                            with st.expander("🏷️ Reviews Per Brand"):
                                attained_brands = attained_details["Brand"].value_counts()
                                st.dataframe(attained_brands)
                            with st.expander("❌ Negative Reviews Per PM"):
                                st.dataframe(negative_reviews_per_pm)
                                st.dataframe(negative_details)
                                st.dataframe(negative_details["Status"].value_counts())
                    else:
                        st.info(f"No Data found for {choice} {selected_month} {number}")
                    st.markdown("---")

            with tab2:
                st.subheader(f"📂 Yearly Data for {choice}")
                number2 = st.number_input("Enter Year", min_value=int(get_min_year()), max_value=current_year,
                                          value=current_year, step=1, key="year_total")
                if number2 and sheet_name:
                    data = load_data_year(sheet_name, number2)
                    if not data.empty:
                        pms = load_data_search(sheet_name, current_year)
                        pm_list = list(set((pms["Project Manager"].dropna().unique().tolist() + ["Unknown"])))
                        render_review_details_view(
                            choice=choice, data=data, pm_list=pm_list,
                            attained_loader=lambda pm: load_reviews_year(sheet_name, number2, pm, "Attained"),
                            negative_loader=lambda pm: load_reviews_year(sheet_name, number2, pm, "Negative"),
                            data_title=f"### 📄 Total Data for {choice} - {number2}",
                            period_label=f"{choice} {number2}",
                            summary_title="### ⭐ Annual Summary",
                            excel_filename=f"{choice}_Total_{number2}.xlsx",
                        )
                    else:
                        st.info(f"No Data Found for {choice} {number2}")

            with tab3:
                st.subheader(f"📂 Start to Year Data for {choice}")
                number5 = st.number_input("Enter Year", min_value=int(get_min_year()), max_value=current_year,
                                          value=get_min_year(), step=1, key="year_total_to_date_start")
                number4 = st.number_input("Enter Year", min_value=int(get_min_year()), max_value=current_year,
                                          value=current_year, step=1, key="year_total_to_date")
                if number4 and number5 and sheet_name:
                    data = load_data_search(sheet_name, number4, number5)
                    if not data.empty:
                        pm_list = list(set((data["Project Manager"].dropna().unique().tolist() + ["Unknown"])))
                        render_review_details_view(
                            choice=choice, data=data, pm_list=pm_list,
                            attained_loader=lambda pm: load_reviews_year_to_date(sheet_name, number4, pm, "Attained"),
                            negative_loader=lambda pm: load_reviews_year_to_date(sheet_name, number4, pm, "Negative"),
                            data_title=f"### 📄 Year to Year Data for {choice} - {number5} to {number4}",
                            period_label=f"{choice} - {number5} to {number4}",
                            summary_title="### ⭐ Start to Year Summary",
                            excel_filename=f"{choice}_Total_{number5}-{number4}.xlsx",
                            download_key="Year_to_date",
                            show_yearly=True, show_merged_yearly=True,
                            show_attained_yearly=True, show_negative_yearly=True,
                        )
                    else:
                        st.info(f"No Data Found for {choice} - {get_min_year()} to {number4}")

            with tab4:
                st.subheader(f"📂 Filtered Data for {choice}")
                col_start, col_end = st.columns(2)
                with col_start:
                    start_date = st.date_input(
                        "📅 Start Date",
                        value=datetime(current_year, 1, 1).date(),
                        min_value=datetime(int(get_min_year()), 1, 1).date(),
                        max_value=datetime.now().date(),
                        key="start_date_filter"
                    )
                with col_end:
                    end_date = st.date_input(
                        "📅 End Date",
                        value=datetime.now().date(),
                        min_value=start_date,
                        max_value=datetime.now().date(),
                        key="end_date_filter"
                    )
                remove_duplicates = st.checkbox("Remove Duplicates")
                if start_date and end_date and sheet_name:
                    data = load_data_filter(sheet_name, start_date, end_date, remove_duplicates)
                    if not data.empty:
                        pm_list = list(set((data["Project Manager"].dropna().unique().tolist() + ["Unknown"])))
                        period_label = f"{choice} - {start_date.strftime('%B %Y')} to {end_date.strftime('%B %Y')}"
                        render_review_details_view(
                            choice=choice, data=data, pm_list=pm_list,
                            attained_loader=lambda pm: load_reviews_filter(sheet_name, start_date, end_date, pm, "Attained"),
                            negative_loader=lambda pm: load_reviews_filter(sheet_name, start_date, end_date, pm, "Negative"),
                            data_title=f"### 📄 Year to Year Data for {period_label}",
                            period_label=period_label,
                            summary_title="### ⭐ Filtered Summary",
                            excel_filename=f"{choice}_Total_{start_date.strftime('%B %Y')} to {end_date.strftime('%B %Y')}.xlsx",
                            download_key="Filtered_data",
                            show_attained_yearly=True, show_negative_yearly=True,
                        )
                    else:
                        st.info(f"No Data Found for {period_label}")

            with tab5:
                st.subheader(f"🔍 Search Data for {choice}")
                number3 = st.number_input("Enter Year for Search", min_value=int(get_min_year()),
                                          max_value=current_year, value=current_year, step=1, key="year_search")
                if number3 and sheet_name:
                    data = load_data_search(sheet_name, number3)
                    if data.empty:
                        st.warning(f"⚠️ No Data Available for {choice} in {number3}")
                    else:
                        search_term = st.text_input("Search by Name / Book / Email",
                                                    placeholder="Enter client name, email or book to search",
                                                    key="search_term")
                        if search_term and search_term.strip():
                            search_df = data[
                                data["Book Name & Link"].str.contains(search_term, case=False, na=False)
                                | data["Name"].str.contains(search_term, case=False, na=False)
                                | data["Email"].str.contains(search_term, case=False, na=False)
                                ]
                            if search_df.empty:
                                st.warning(f"⚠️ No results found for '{search_term}'")
                            else:
                                st.success(f"✅ Found {len(search_df)} result(s) for '{search_term}'")
                                search_df.index = range(1, len(search_df) + 1)
                                st.dataframe(search_df)
                                download_excel_button(
                                    search_df, f"{choice}_Search_{search_term}_{number3}.xlsx",
                                    label="📥 Download Search Results")
                                st.markdown("---")
                                st.markdown("### 📊 Search Results Summary")
                                col1, col2 = st.columns(2)
                                with col1:
                                    st.markdown(f"- 🧾 **Total Matches:** `{len(search_df)}`")
                                    if "Brand" in search_df.columns:
                                        brands = search_df["Brand"].value_counts()
                                        st.markdown("**Brands in Results:**")
                                        for brand, count in brands.items():
                                            st.markdown(f"  - {brand}: `{count}`")
                                with col2:
                                    if "Platform" in search_df.columns:
                                        platforms = search_df["Platform"].value_counts()
                                        st.markdown("**Platforms in Results:**")
                                        for platform, count in platforms.items():
                                            st.markdown(f"  - {platform}: `{count}`")
                                st.markdown("### 📊 Search Results with unique titles only")
                                drop_search_dupes = search_df.drop_duplicates(subset=["Name", "Book Name & Link"])
                                drop_search_dupes.index = range(1, len(drop_search_dupes) + 1)
                                st.dataframe(drop_search_dupes[["Name", "Book Name & Link"]])
                                st.markdown("---")
                                st.markdown("### 📊 Search Results with unique titles only")
                                col1, col2 = st.columns(2)
                                with col1:
                                    st.markdown(f"- 🧾 **Total Matches:** `{len(drop_search_dupes)}`")
                                    if "Brand" in drop_search_dupes.columns:
                                        brands = drop_search_dupes["Brand"].value_counts()
                                        st.markdown("**Brands in Results:**")
                                        for brand, count in brands.items():
                                            st.markdown(f"  - {brand}: `{count}`")
                                with col2:
                                    if "Platform" in drop_search_dupes.columns:
                                        platforms = drop_search_dupes["Platform"].value_counts()
                                        st.markdown("**Platforms in Results:**")
                                        for platform, count in platforms.items():
                                            st.markdown(f"  - {platform}: `{count}`")
                        else:
                            st.info("👆 Enter name/book/email above to search")

            with tab6:
                st.subheader(f"📊 Filter Data by Brand for {choice}")
                number4 = st.number_input("Select Year", min_value=int(get_min_year()),
                                          max_value=current_year, value=current_year, step=1, key="year_filter")
                if number4 and sheet_name:
                    selected_brand = USA_BRANDS if sheet_name == "USA" else UK_BRANDS
                    brand_selection = st.selectbox("Select Brand", selected_brand, key="brand_selection")
                    data = load_data_year(sheet_name, number4)
                    if data.empty:
                        st.warning(f"⚠️ No Data Available for {choice} in {number4}")
                    else:
                        filtered_df = data[data["Brand"] == brand_selection]
                        if filtered_df.empty:
                            st.warning(f"⚠️ No records for brand '{brand_selection}' in {number4}")
                        else:
                            filtered_df = filtered_df.drop_duplicates(["Name"], keep="first")
                            filtered_df.index = range(1, len(filtered_df) + 1)
                            st.dataframe(filtered_df)
                            download_excel_button(
                                filtered_df, f"{choice}_Brand_{brand_selection}_{number4}.xlsx",
                                label="📥 Download Filtered Data")

        elif action == "Printing":
            tab1, tab2, tab3, tab4, tab5 = st.tabs(["Monthly", "Yearly", "Start to Year", "Search", "Stats"])

            with tab1:
                selected_month = st.selectbox("Select Month", month_list, index=current_month - 1,
                                              placeholder="Select Month")
                number = st.number_input("Enter Year", min_value=int(get_min_year()), max_value=current_year,
                                         value=current_year, step=1)
                selected_month_number = month_list.index(selected_month) + 1 if selected_month else None
                if selected_month and number:
                    st.subheader(f"🖨️ Printing Summary for {selected_month} {number}")
                    data = get_printing_data_month(selected_month_number, number)
                    if not data.empty:
                        render_printing_period(
                            data,
                            heading="### 📄 Detailed Printing Data",
                            download_filename=f"Printing_{selected_month}_{number}.xlsx",
                            stats_heading="📊 Summary Statistics",
                        )
                    else:
                        st.warning(f"⚠️ No Data Available for Printing in {selected_month} {number}")

            with tab2:
                number2 = st.number_input("Enter Year", min_value=int(get_min_year()), max_value=current_year,
                                          value=current_year, step=1, key="year_key")
                data, monthly = printing_data_year(number2)
                if not data.empty:
                    render_printing_period(
                        data,
                        heading=f"### 📄 Yearly Printing Data for {number2}",
                        download_filename=f"Printing_{number2}.xlsx",
                        stats_heading="📊 Summary Statistics (All Data)",
                        monthly=monthly,
                    )
                else:
                    st.warning(f"⚠️ No Data Available for Printing in {number2}")

            with tab3:
                number2 = st.number_input("Enter Year", min_value=int(get_min_year()), max_value=current_year,
                                          value=current_year, step=1, key="printing_year_to_year")
                data, monthly = printing_data_search(number2)
                if not data.empty:
                    render_printing_period(
                        data,
                        heading=f"### 📄 Start to Year Printing Data for 2025 to {number2}",
                        download_filename=f"Printing_Start to Year_{number2}.xlsx",
                        stats_heading="📊 Summary Statistics (Start to Year)",
                        download_key="Start_to_Year",
                        monthly=monthly,
                    )
                else:
                    st.warning(f"⚠️ No Data Available for Printing in Start to Year {number2}")

            with tab4:
                number3 = st.number_input("Enter Year", min_value=int(get_min_year()), max_value=current_year,
                                          value=current_year, step=1, key="search_key")
                data, _ = printing_data_search(number3)
                search_term = st.text_input("Search by Name / Book", placeholder="Enter Search Term",
                                            key="search_term")
                if search_term and search_term.strip():
                    search_df = data[
                        data["Book"].str.contains(search_term, case=False, na=False)
                        | data["Name"].str.contains(search_term, case=False, na=False)
                        ]
                    if search_df.empty:
                        st.warning("No such orders found!")
                    else:
                        df = search_df.copy()
                        total_orders = len(df)
                        total_copies = df["No of Copies"].sum()
                        total_cost = df["Order Cost"].sum()
                        highest_cost = df["Order Cost"].max()
                        lowest_cost = df["Order Cost"].min()
                        highest_copies = df["No of Copies"].max()
                        lowest_copies = df["No of Copies"].min()
                        avg_cost_per_copy = round(total_cost / total_copies, 2) if total_copies else 0
                        st.markdown("### 🌍 Printing Summary")
                        st.markdown(f"""
                        - 🧾 **Total Orders:** {total_orders}
                        - 📦 **Total Copies Printed:** `{total_copies}`
                        - 💰 **Total Cost:** `${total_cost:,.2f}`
                        - 📈 **Highest Order Cost:** `${highest_cost:,.2f}`
                        - 📉 **Lowest Order Cost:** `${lowest_cost:,.2f}`
                        - 🔢 **Highest Copies in One Order:** `{highest_copies}`
                        - 🧮 **Lowest Copies in One Order:** `{lowest_copies}`
                        - 💵 **Average Cost per Copy:** `${avg_cost_per_copy:,.2f}`
                        """)
                        search_df["Order Cost"] = search_df["Order Cost"].map("${:,.2f}".format)
                        search_df.index = range(1, len(search_df) + 1)
                        st.dataframe(search_df)
                else:
                    st.info("👆 Enter name/book above to search")

            with tab5:
                st.subheader("📊 Year-over-Year Printing Stats")
                year1 = st.number_input("Enter Previous Year", min_value=int(get_min_year()),
                                        max_value=current_year, value=current_year - 1, step=1)
                year2 = st.number_input("Enter Current Year", min_value=int(get_min_year()),
                                        max_value=current_year, value=current_year, step=1)
                data1, _ = printing_data_year(year1)
                data2, _ = printing_data_year(year2)

                def pct_change(new, old):
                    return round(((new - old) / old) * 100, 2) if old else 0

                if not data1.empty and not data2.empty:
                    total_orders1, total_orders2 = len(data1), len(data2)
                    total_copies1, total_copies2 = data1["No of Copies"].sum(), data2["No of Copies"].sum()
                    total_cost1, total_cost2 = data1["Order Cost"].sum(), data2["Order Cost"].sum()

                    st.subheader("🌍 Overall Printing Comparison")
                    col1, col2, col3 = st.columns(3)
                    col1.metric(f"🧾 Orders ({year1} → {year2})", total_orders2,
                                f"{pct_change(total_orders2, total_orders1)}%")
                    col2.metric(f"📦 Copies ({year1} → {year2})", total_copies2,
                                f"{pct_change(total_copies2, total_copies1)}%")
                    col3.metric(f"💰 Cost ({year1} → {year2})", f"${total_cost2:,.2f}",
                                f"{pct_change(total_cost2, total_cost1)}%")
                    col1, col2, col3 = st.columns(3)
                    col1.metric(f"🧾 Orders ({year2} <- {year1})", total_orders1,
                                f"{pct_change(total_orders1, total_orders2)}%")
                    col2.metric(f"📦 Copies ({year2} <- {year1})", total_copies1,
                                f"{pct_change(total_copies1, total_copies2)}%")
                    col3.metric(f"💰 Cost ({year2} <- {year1})", f"${total_cost1:,.2f}",
                                f"{pct_change(total_cost1, total_cost2)}%")
                    st.markdown("---")

                    def country_stats(df, brands):
                        df = df[df["Brand"].isin(brands)]
                        return len(df), df["No of Copies"].sum(), df["Order Cost"].sum()

                    usa_orders1, usa_copies1, usa_cost1 = country_stats(data1, USA_PRINTING_BRANDS)
                    usa_orders2, usa_copies2, usa_cost2 = country_stats(data2, USA_PRINTING_BRANDS)
                    st.subheader("🇺🇸 USA Printing Comparison")
                    col1, col2, col3 = st.columns(3)
                    col1.metric("🧾 Orders", usa_orders2, f"{pct_change(usa_orders2, usa_orders1)}%")
                    col2.metric("📦 Copies", usa_copies2, f"{pct_change(usa_copies2, usa_copies1)}%")
                    col3.metric("💰 Cost", f"${usa_cost2:,.2f}", f"{pct_change(usa_cost2, usa_cost1)}%")
                    col1, col2, col3 = st.columns(3)
                    col1.metric("🧾 Orders (Reverse)", usa_orders1, f"{pct_change(usa_orders1, usa_orders2)}%")
                    col2.metric("📦 Copies (Reverse)", usa_copies1, f"{pct_change(usa_copies1, usa_copies2)}%")
                    col3.metric("💰 Cost (Reverse)", f"${usa_cost1:,.2f}", f"{pct_change(usa_cost1, usa_cost2)}%")
                    st.markdown("---")

                    uk_orders1, uk_copies1, uk_cost1 = country_stats(data1, UK_BRANDS)
                    uk_orders2, uk_copies2, uk_cost2 = country_stats(data2, UK_BRANDS)
                    st.subheader("🇬🇧 UK Printing Comparison")
                    col1, col2, col3 = st.columns(3)
                    col1.metric("🧾 Orders", uk_orders2, f"{pct_change(uk_orders2, uk_orders1)}%")
                    col2.metric("📦 Copies", uk_copies2, f"{pct_change(uk_copies2, uk_copies1)}%")
                    col3.metric("💰 Cost", f"${uk_cost2:,.2f}", f"{pct_change(uk_cost2, uk_cost1)}%")
                    col1, col2, col3 = st.columns(3)
                    col1.metric("🧾 Orders (Reverse)", uk_orders1, f"{pct_change(uk_orders1, uk_orders2)}%")
                    col2.metric("📦 Copies (Reverse)", uk_copies1, f"{pct_change(uk_copies1, uk_copies2)}%")
                    col3.metric("💰 Cost (Reverse)", f"${uk_cost1:,.2f}", f"{pct_change(uk_cost1, uk_cost2)}%")
                else:
                    st.warning("⚠️ No data available for one or both years.")

        elif action == "Copyright":
            tab1, tab2, tab3 = st.tabs(["Monthly", "Yearly", "Search"])

            with tab1:
                selected_month = st.selectbox("Select Month", month_list, index=current_month - 1,
                                              placeholder="Select Month")
                number = st.number_input("Enter Year", min_value=int(get_min_year()), max_value=current_year,
                                         value=current_year, step=1, key="number_copyright")
                selected_month_number = month_list.index(selected_month) + 1 if selected_month else None
                if selected_month and number:
                    st.subheader(f"© Copyright Summary for {selected_month} {number}")
                    data, approved, rejected = get_copyright_month(selected_month_number, number)
                    render_copyright_period(
                        data, approved, rejected,
                        heading="### 📊 Summary Statistics (All Data)",
                        download_filename=f"Copyright_{selected_month}_{number}.xlsx",
                        no_data_msg=f"⚠️ No Data Available for {selected_month} {number}")

            with tab2:
                number2 = st.number_input("Enter Year", min_value=int(get_min_year()), max_value=current_year,
                                          value=current_year, step=1, key="copyright_year_total")
                data, approved, rejected = copyright_year(number2)
                st.subheader(f"© Yearly Copyright Data for {number2}")
                render_copyright_period(
                    data, approved, rejected,
                    heading="### 📊 Summary Statistics (All Data)",
                    download_filename=f"Copyright_{number2}.xlsx",
                    no_data_msg=f"⚠️ No Data Available for {number2}")

            with tab3:
                number3 = st.number_input("Enter Year", min_value=int(get_min_year()), max_value=current_year,
                                          value=current_year, step=1, key="copyright_search")
                data, _, _ = copyright_search(number3)
                search_term = st.text_input("Search by Title / Name", placeholder="Enter Search Term")
                if search_term and not data.empty:
                    search_df = data[
                        data["Book Name & Link"].str.contains(search_term, case=False, na=False)
                        | data["Name"].str.contains(search_term, case=False, na=False)
                        ]
                    if search_df.empty:
                        st.warning("No matching records found.")
                    else:
                        approved = len(search_df[search_df["Result"] == "Yes"]) if "Result" in search_df.columns else 0
                        rejected = len(search_df[search_df["Result"] == "No"]) if "Result" in search_df.columns else 0
                        st.markdown("---")
                        st.markdown("### 📊 Summary Statistics (All Data)")
                        total_titles = len(search_df)
                        country_counts = search_df["Country"].value_counts()
                        country_usa = country_counts.get("USA", 0)
                        country_uk = country_counts.get("UK", 0)
                        country_canada = country_counts.get("Canada", 0)
                        total_cost = (country_usa * COPYRIGHT_RATES["USA"]) + \
                                     (country_canada * COPYRIGHT_RATES["Canada"]) + \
                                     (country_uk * COPYRIGHT_RATES["UK"])
                        st.markdown(f"""
                        - 🧾 **Total Titles:** `{total_titles}`
                        - 💵 **Total Cost:** `${total_cost}`
                        - ✅ **Approved:** `{approved}` ({approved / total_titles:.1%})
                        - ❌ **Rejected:** `{rejected}` ({rejected / total_titles:.1%})
                        - 🦅 **USA:** `{country_usa}`
                        - 🍁 **Canada:** `{country_canada}`
                        - ☕ **UK:** `{country_uk}`
                        """)
                        search_df.index = range(1, len(search_df) + 1)
                        st.dataframe(search_df)
                else:
                    st.info("👆 Enter name/book above to search")

        elif action == "Generate Similarity":
            tab1, tab2, tab3, tab4 = st.tabs(["Queries", "Yearly Queries", "Compare Years", "Custom"])

            def safe_month_index(month_offset: int, month_list_len: int) -> int:
                """Ensure selectbox index is within valid range."""
                return max(0, min(month_offset, month_list_len - 1))

            with tab1:
                st.header("Compare clients with months")
                choice = st.selectbox("Select Data To View", ["USA", "UK"], index=None, key="choice_tab1")
                sheet_name = {"UK": sheet_uk, "USA": sheet_usa}.get(choice)
                index_month1 = safe_month_index(current_month - 2, len(month_list))
                index_month2 = safe_month_index(current_month - 1, len(month_list))
                selected_month_1 = st.selectbox("Select Month 1", month_list, index=index_month1, key="month1_tab1")
                number1 = st.number_input("Enter Year 1", min_value=int(get_min_year()), max_value=current_year,
                                          value=current_year, step=1, key="year1_tab1")
                selected_month_2 = st.selectbox("Select Month 2", month_list, index=index_month2, key="month2_tab1")
                number2 = st.number_input("Enter Year 2", min_value=int(get_min_year()), max_value=current_year,
                                          value=current_year, step=1, key="year2_tab1")
                if sheet_name:
                    if st.button("Generate Similar Clients", key="btn_generate_tab1"):
                        with st.spinner(
                                f"Generating Similarity Report for {selected_month_1} & {selected_month_2} for {choice}..."):
                            data1, data2, data3 = get_names_in_both_months(
                                sheet_name, selected_month_1, number1, selected_month_2, number2)
                            if not data1:
                                st.info("No similarities found")
                            else:
                                st.metric(label="Total Number of Same Clients", value=data3)
                                st.write("Names:")
                                st.json(data1, expanded=True)
                                st.write("Detailed Names:")
                                st.json(data2, expanded=False)

            with tab2:
                choice = st.selectbox("Select Data To View", ["USA", "UK"], index=None, key="choice_tab2")
                sheet_name = {"UK": sheet_uk, "USA": sheet_usa}.get(choice)
                number3 = st.number_input("Enter Year", min_value=int(get_min_year()), max_value=current_year,
                                          value=current_year, step=1, key="year_tab2")
                if sheet_name and number3:
                    df_year, Total_year, year_count = get_names_in_year(sheet_name, number3)
                    if not df_year.empty:
                        st.metric(label="Total Number of Same Clients", value=year_count)
                        st.write("Yearly Data:")
                        st.write(df_year)
                        st.write("Total Year:")
                        st.json(Total_year, expanded=False)
                    else:
                        st.warning(f"No Similarities found for {number3}-{choice}")

            with tab3:
                st.header("Compare clients with Years")
                choice = st.selectbox("Select Data To View", ["USA", "UK"], index=None, key="choice_tab3")
                sheet_name = {"UK": sheet_uk, "USA": sheet_usa}.get(choice)
                number1 = st.number_input("Enter Year 1", min_value=int(get_min_year()), max_value=current_year,
                                          value=current_year - 1, step=1, key="year1_tab3")
                number2 = st.number_input("Enter Year 2", min_value=int(get_min_year()), max_value=current_year,
                                          value=current_year, step=1, key="year2_tab3")
                if sheet_name:
                    if st.button("Generate Similar Clients", key="btn_generate_tab3"):
                        with st.spinner(f"Generating Similarity Report for {number1} & {number2} for {choice}..."):
                            data1, data2, data3 = get_names_in_both_years(sheet_name, number1, number2)
                            if not data1:
                                st.info("No similarities found")
                            else:
                                st.metric(label="Total Number of Same Clients", value=data3)
                                st.write("Names:")
                                st.json(data1, expanded=True)
                                st.write("Detailed Names:")
                                st.json(data2, expanded=False)
                                for name, years in data2.items():
                                    with st.expander(name):
                                        for year, data in years.items():
                                            st.markdown(f"### {year}")
                                            st.write(f"**Count:** {data['count']}")
                                            st.write("**Publishing Dates:**")
                                            st.markdown("\n".join([f"- {d}" for d in data["publishing_dates"]]))

            with tab4:
                st.header("Compare clients custom")
                choice = st.selectbox("Select Data To View", ["USA", "UK"], index=None, key="choice_tab4")
                sheet_name = {"UK": sheet_uk, "USA": sheet_usa}.get(choice)
                selected_month_1 = st.selectbox("Select Month 1", month_list, index=current_month - 1,
                                                key="month1_tab4")
                number1 = st.number_input("Enter Year 1", min_value=int(get_min_year()), max_value=current_year,
                                          value=current_year - 1, step=1, key="year1_tab4")
                number2 = st.number_input("Enter Year 2", min_value=int(get_min_year()), max_value=current_year,
                                          value=current_year, step=1, key="year2_tab4")
                if sheet_name:
                    if st.button(f"Search Similar Clients for {selected_month_1}", key="btn_generate_tab4"):
                        with st.spinner("Searching repeating clients for targeted month"):
                            data1, data2, data3 = get_clients_returning_in_month(
                                sheet_name, number1, selected_month_1, number2)
                            if not data1:
                                st.info("No similarities found")
                            else:
                                st.metric(label="Total Number of Same Clients", value=data3)
                                st.write("Names:")
                                st.json(data1, expanded=True)
                                st.write("Detailed Names:")
                                st.json(data2, expanded=False)
                                for name, years in data2.items():
                                    with st.expander(name):
                                        for year, data in years.items():
                                            st.markdown(f"### {year}")
                                            st.write(f"**Count:** {data['count']}")
                                            st.write("**Publishing Dates:**")
                                            st.markdown("\n".join([f"- {d}" for d in data["publishing_dates"]]))

        elif action == "Summary":
            st.header("📄 Generate Summary Report")
            selected_month = st.selectbox("Select Month", month_list, index=current_month - 1,
                                          placeholder="Select Month")
            number = st.number_input("Enter Year", min_value=int(get_min_year()), max_value=current_year,
                                     value=current_year, step=1)
            selected_month_number = month_list.index(selected_month) + 1 if selected_month else None

            uk_clean = clean_data_reviews(sheet_uk)
            usa_clean = clean_data_reviews(sheet_usa)
            usa_clean = usa_clean[(usa_clean["Publishing Date"].dt.month == selected_month_number) &
                                  (usa_clean["Publishing Date"].dt.year == number)]
            uk_clean = uk_clean[(uk_clean["Publishing Date"].dt.month == selected_month_number) &
                                (uk_clean["Publishing Date"].dt.year == number)]

            if usa_clean.empty or uk_clean.empty:
                st.error(f"Cannot generate summary — no data available for the month {selected_month} {number}.")
            else:
                if st.button("Generate Summary"):
                    with st.spinner(f"Generating Summary Report for {selected_month} {number}..."):
                        data = summary(selected_month_number, number)
                        if not data:
                            st.error(f"Cannot generate summary — no data available for the month {selected_month} {number}.")
                        else:
                            pdf_data, pdf_filename = generate_summary_report_pdf(
                                data["usa_review"], data["uk_review"], data["usa_brands"], data["uk_brands"],
                                data["usa_platforms"], data["uk_platforms"], data["printing_stats"],
                                data["copyright_stats"], data["a_plus_count"],
                                selected_month=selected_month, start_year=number)
                            render_month_summary_report(
                                data,
                                f"{selected_month} {number} Summary Report",
                                f"USA+UK_{selected_month}_{number}.xlsx",
                                pdf_data, pdf_filename)

        elif action == "Year Summary" and number:
            st.header("📄 Generate Year Summary Report")
            uk_clean = clean_data_reviews(sheet_uk)
            usa_clean = clean_data_reviews(sheet_usa)
            usa_clean = usa_clean[usa_clean["Publishing Date"].dt.year == number]
            uk_clean = uk_clean[uk_clean["Publishing Date"].dt.year == number]

            if usa_clean.empty or uk_clean.empty:
                st.error(f"Cannot generate summary — no data available for the Year {number}.")
            else:
                if st.button("Generate Year Summary Report"):
                    with st.spinner("Generating Year Summary Report"):
                        data = generate_year_summary(number)
                        if not data:
                            st.error(f"Cannot generate summary — no data available for the Year {number}.")
                        else:
                            pdf_data, pdf_filename = generate_summary_report_pdf(
                                data["usa_review"], data["uk_review"], data["usa_brands"], data["uk_brands"],
                                data["usa_platforms"], data["uk_platforms"], data["printing_stats"],
                                data["copyright_stats"], data["a_plus_count"], start_year=number)
                            render_year_summary_report(
                                data, f"{number} Summary Report", f"USA+UK_{number}.xlsx", pdf_data, pdf_filename)

        elif action == "Custom Summary":
            start_year = st.number_input("Enter Year", min_value=int(get_min_year()), max_value=current_year,
                                         value=current_year - 1, step=1, key="start_year")
            end_year = st.number_input("Enter Year", min_value=int(get_min_year()), max_value=current_year,
                                       value=current_year, step=1, key="end_year")
            st.header("📄 Generate Multi Year Summary Report")

            uk_clean = clean_data_reviews(sheet_uk)
            usa_clean = clean_data_reviews(sheet_usa)
            usa_clean = usa_clean[(usa_clean["Publishing Date"].dt.year >= start_year) &
                                  (usa_clean["Publishing Date"].dt.year <= end_year)]
            uk_clean = uk_clean[(uk_clean["Publishing Date"].dt.year >= start_year) &
                                (uk_clean["Publishing Date"].dt.year <= end_year)]

            if usa_clean.empty or uk_clean.empty:
                st.error(f"Cannot generate summary — no data available for the Years {start_year}-{end_year}.")
            else:
                if st.button("Generate Year Summary Report"):
                    with st.spinner("Generating Year Summary Report"):
                        data = generate_year_summary(start_year, end_year)
                        if not data:
                            st.error(f"Cannot generate summary — no data available for the Years {start_year}-{end_year}.")
                        else:
                            pdf_data, pdf_filename = generate_summary_report_pdf(
                                data["usa_review"], data["uk_review"], data["usa_brands"], data["uk_brands"],
                                data["usa_platforms"], data["uk_platforms"], data["printing_stats"],
                                data["copyright_stats"], data["a_plus_count"],
                                start_year=start_year, end_year=end_year)
                            render_year_summary_report(
                                data,
                                f"{start_year}-{end_year} Summary Report",
                                f"USA+UK_{start_year}-{end_year}.xlsx",
                                pdf_data, pdf_filename)

        elif action == "Sales":
            tab1, tab2 = st.tabs(["Monthly", "Yearly"])

            with tab1:
                selected_month = st.selectbox("Select Month", month_list, index=current_month - 1,
                                              placeholder="Select Month")
                number = st.number_input("Enter Year", min_value=int(get_min_year()), max_value=current_year,
                                         value=current_year, step=1)
                selected_month_number = month_list.index(selected_month) + 1 if selected_month else None
                if selected_month and number:
                    data = sales(selected_month_number, number)
                    if not data.empty:
                        total_sales = data["Payment"].sum()
                        show_data = data.copy()
                        show_data["Payment"] = show_data["Payment"].map("${:,.2f}".format)
                        st.markdown("### 📄 Detailed Monthly Sales Data")
                        st.dataframe(show_data)
                        st.markdown("---")
                        st.markdown("### 📊 Monthly Summary")
                        st.markdown(f"""
                        - 🧾 **Total Clients:** `{len(data)}`
                        - 💰 **Total Sales:** `${total_sales:,.2f}`
                        """)
                    else:
                        st.warning(f"⚠️ No Data Available for Sales in {selected_month} {number}")

            with tab2:
                year = st.number_input("Enter Year", min_value=int(get_min_year()), max_value=current_year,
                                       value=current_year, step=1, key="sales_year")
                data = sales_year(year)
                if not data.empty:
                    total_sales = data["Payment"].sum()
                    show_data = data.copy()
                    show_data["Payment"] = show_data["Payment"].map("${:,.2f}".format)
                    st.markdown(f"### 📄 Total Sales Data for {year}")
                    st.dataframe(show_data)
                    st.markdown("---")
                    st.markdown("### 📊 Yearly Summary")
                    st.markdown(f"""
                    - 🧾 **Total Clients:** `{len(data)}`
                    - 💰 **Total Sales:** `${total_sales:,.2f}`
                    """)
                else:
                    st.warning(f"⚠️ No Data Available for Sales in {year}")

        elif action == "Reviews" and number:
            uk_clean = clean_data_reviews(sheet_uk)
            usa_clean = clean_data_reviews(sheet_usa)
            usa_clean = usa_clean[usa_clean["Publishing Date"].dt.year == number]
            uk_clean = uk_clean[uk_clean["Publishing Date"].dt.year == number]

            if usa_clean.empty or uk_clean.empty:
                st.warning("No combined review data found.")
            else:
                usa_clean = usa_clean.drop_duplicates(subset=["Name"], keep="last")
                uk_clean = uk_clean.drop_duplicates(subset=["Name"], keep="last")
                pm_list_usa = list(set((usa_clean["Project Manager"].dropna().unique().tolist() + ["Unknown"])))
                pm_list_uk = list(set((uk_clean["Project Manager"].dropna().unique().tolist() + ["Unknown"])))
                usa_reviews_per_pm = safe_concat(
                    [load_reviews_year(sheet_usa, number, pm, "Attained") for pm in pm_list_usa])
                uk_reviews_per_pm = safe_concat(
                    [load_reviews_year(sheet_uk, number, pm, "Attained") for pm in pm_list_uk])
                combined_data = safe_concat([usa_reviews_per_pm, uk_reviews_per_pm])

                if not combined_data.empty:
                    combined_data["Trustpilot Review Date"] = pd.to_datetime(
                        combined_data["Trustpilot Review Date"], format=DATE_FORMAT, errors="coerce")
                    combined_data["Month-Year"] = combined_data["Trustpilot Review Date"].dt.to_period("M").astype(str)

                    monthly_counts = (
                        combined_data.groupby(["Project Manager", "Month-Year"])
                        .size().reset_index(name="Review Count")
                    )
                    monthly_clients = (
                        combined_data.groupby(["Project Manager", "Month-Year"])["Name"]
                        .apply(list).reset_index(name="Clients")
                    )
                    monthly_summary = pd.merge(monthly_counts, monthly_clients, on=["Project Manager", "Month-Year"],
                                               how="left")
                    monthly_summary["Month-Year"] = pd.to_datetime(monthly_summary["Month-Year"])
                    monthly_summary = monthly_summary.sort_values(["Project Manager", "Month-Year"])
                    monthly_summary["Month-Year"] = monthly_summary["Month-Year"].dt.strftime("%B %Y")
                    monthly_summary.index = range(1, len(monthly_summary) + 1)

                    st.subheader("📅 Monthly Review Counts per PM")
                    with st.expander("🟢 Monthly Attained Counts per PM (with Clients)"):
                        st.dataframe(monthly_summary, width="stretch")

                    monthly_pivot = monthly_summary.pivot_table(
                        index="Project Manager", columns="Month-Year", values="Review Count", fill_value=0)
                    monthly_pivot = monthly_pivot.reindex(
                        sorted(monthly_pivot.columns, key=lambda x: pd.to_datetime(x)), axis=1)
                    monthly_pivot.columns = [pd.to_datetime(col).strftime("%B %Y") for col in monthly_pivot.columns]
                    with st.expander("📊 Monthly Review Count Pivot Table"):
                        st.dataframe(monthly_pivot, width="stretch")
                else:
                    st.warning("No combined review data found.")

        elif action == "ISBN":
            st.title("Nielsen ISBN")
            t1, t2, t3 = st.tabs(["All", "Search", "Filter By Brand"])
            data = nielsen_isbn()
            with t1:
                st.subheader("All Records")
                df_all = data.copy()
                df_all.index = range(1, len(df_all) + 1)
                st.dataframe(df_all)
            with t2:
                st.subheader("Search")
                search_term = st.text_input("Search by ISBN / Book / Author",
                                            placeholder="Enter Search Term", key="search_term")
                if search_term and search_term.strip():
                    search_df = data[
                        data["ISBN"].str.contains(search_term, case=False, na=False)
                        | data["Title"].str.contains(search_term, case=False, na=False)
                        | data["Author"].str.contains(search_term, case=False, na=False)
                        | data["Subtitle"].str.contains(search_term, case=False, na=False)
                        ]
                    if search_df.empty:
                        st.warning("No such ISBNs found!")
                    else:
                        search_df.index = range(1, len(search_df) + 1)
                        st.dataframe(search_df)
                else:
                    st.info("👆 Enter ISBN/book/author above to search")
            with t3:
                st.subheader("Filter By Brand")
                if "Brand" in data.columns:
                    brands = sorted(data["Brand"].dropna().unique())
                    selected_brand = st.selectbox("Select Brand", brands)
                    filtered_df = data[data["Brand"] == selected_brand]
                    if filtered_df.empty:
                        st.warning("No records found for this brand.")
                    else:
                        filtered_df.index = range(1, len(filtered_df) + 1)
                        st.dataframe(filtered_df)
                else:
                    st.error("Brand column not found in dataset.")

        elif action == "Chargeback":
            st.title("💳 Chargeback")
            tab_all, tab_month, tab_year, tab_brand = st.tabs(
                ["All", "Monthly", "Yearly", "Brand-wise"])

            with tab_all:
                data = load_chargeback()
                if data.empty:
                    st.warning("⚠️ No Chargeback data available.")
                else:
                    render_chargeback_stats(data, "### 📊 Overall Chargeback Statistics")
                    st.markdown("### 📄 All Chargeback Records")
                    disp = data.drop(columns=[c for c in ["_Chargeback_dt"] if c in data.columns])
                    disp.index = range(1, len(disp) + 1)
                    st.dataframe(disp)
                    download_excel_button(disp, "Chargeback.xlsx")

            with tab_month:
                selected_month = st.selectbox("Select Month", month_list, index=current_month - 1,
                                              placeholder="Select Month", key="chargeback_month")
                number = st.number_input("Enter Year", min_value=int(get_min_year()), max_value=current_year,
                                         value=current_year, step=1, key="chargeback_month_year")
                selected_month_number = month_list.index(selected_month) + 1 if selected_month else None
                if selected_month and number:
                    data = load_chargeback()
                    if "_Chargeback_dt" in data.columns:
                        data = data[(data["_Chargeback_dt"].dt.month == selected_month_number) &
                                    (data["_Chargeback_dt"].dt.year == number)]
                    if data.empty:
                        st.warning(f"⚠️ No Chargeback data for {selected_month} {number}.")
                    else:
                        render_chargeback_stats(
                            data, f"### 📊 Chargeback Statistics — {selected_month} {number}")
                        st.markdown("### 📄 Records")
                        disp = data.drop(columns=[c for c in ["_Chargeback_dt"] if c in data.columns])
                        disp.index = range(1, len(disp) + 1)
                        st.dataframe(disp)

            with tab_year:
                number2 = st.number_input("Enter Year", min_value=int(get_min_year()), max_value=current_year,
                                          value=current_year, step=1, key="chargeback_year")
                data = load_chargeback()
                if "_Chargeback_dt" in data.columns:
                    data = data[data["_Chargeback_dt"].dt.year == number2]
                if data.empty:
                    st.warning(f"⚠️ No Chargeback data for {number2}.")
                else:
                    render_chargeback_stats(data, f"### 📊 Chargeback Statistics — {number2}")
                    st.markdown("### 📄 Records")
                    disp = data.drop(columns=[c for c in ["_Chargeback_dt"] if c in data.columns])
                    disp.index = range(1, len(disp) + 1)
                    st.dataframe(disp)

            with tab_brand:
                data = load_chargeback()
                if data.empty or "Brand" not in data.columns:
                    st.warning("⚠️ No Chargeback data available.")
                else:
                    rows = []
                    for brand, grp in data.groupby("Brand"):
                        total = len(grp)
                        total_payment = grp["Payment"].sum()
                        fav = grp[grp["Favourable"]]
                        lost = grp[grp["Lost"]]
                        payment_saved = fav["Payment"].sum()
                        lost_payment = lost["Payment"].sum()
                        rows.append({
                            "Brand": brand,
                            "Chargebacks": total,
                            "Payment at Risk": f"${total_payment:,.2f}",
                            "Won": len(fav),
                            "Payment Saved": f"${payment_saved:,.2f}",
                            "Lost": len(lost),
                            "Lost Payment": f"${lost_payment:,.2f}",
                        })
                    summary = pd.DataFrame(rows).sort_values(by="Payment Saved", ascending=False)
                    summary.index = range(1, len(summary) + 1)
                    summary.index.name = "#"
                    st.markdown("### 🏷️ Brand-wise Breakdown")
                    st.dataframe(summary)
                    download_excel_button(summary, "Chargeback_Brandwise.xlsx")
                    st.markdown("### 📄 All Records")
                    disp = data.drop(columns=[c for c in ["_Chargeback_dt"] if c in data.columns])
                    disp.index = range(1, len(disp) + 1)
                    st.dataframe(disp)


if __name__ == '__main__':
    main()
