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
# Google Sheets setup
SCOPES = ["https://www.googleapis.com/auth/spreadsheets.readonly"]
# creds = Credentials.from_service_account_file("credentials.json", scopes=SCOPES)
SPREADSHEET_ID = st.secrets["connections"]["gsheets"]["SPREADSHEET_ID"]

@st.cache_resource
def get_gsheets_client(creds_dict: dict, spreadsheet_id: str):
    """Create and cache Google Sheets client + spreadsheet"""
    creds = Credentials.from_service_account_info(
        creds_dict,
        scopes=SCOPES
    )
    client = gspread.authorize(creds)
    spreadsheet = client.open_by_key(spreadsheet_id)
    return spreadsheet

spreadsheet = get_gsheets_client(creds_dict, SPREADSHEET_ID)

# Sheet names
sheet_usa = "USA"
sheet_uk = "UK"
sheet_audio = "AudioBook"
sheet_printing = "Printing"
sheet_copyright = "Copyright"
sheet_a_plus = "A_plus"
sheet_sales = "Sales"

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
    """Gets Minimum year from the data"""
    # uk_clean = clean_data_reviews(sheet_uk)
    # audio = clean_data_reviews(sheet_audio)
    # usa_clean = clean_data_reviews(sheet_usa)
    # combined = pd.concat([uk_clean, usa_clean, audio])
    #
    # combined["Publishing Date"] = pd.to_datetime(combined["Publishing Date"], errors="coerce")
    #
    # min_year = combined["Publishing Date"].dt.year.min()

    return 2025

def normalize_name(name):
    """Normalize a name to consistent format (Title Case, stripped whitespace)"""
    if pd.isna(name) or name == "":
        return ""
    return str(name).strip().title()


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
        print(f"Error getting data from sheet {sheet_name}: {e}")
        logging.error(f"Error getting data from sheet {sheet_name}: {e}")
        return pd.DataFrame()


def clean_data(data: pd.DataFrame) -> pd.DataFrame:
    """Clean and prepare the dataframe"""
    if data.empty:
        return pd.DataFrame()

    columns = list(data.columns)
    if "Issues" in columns:
        end_col_index = columns.index("Issues")
        data = data.iloc[:, :end_col_index + 1]
    date_columns = ["Publishing Date", "Last Edit (Revision)", "Trustpilot Review Date"]
    for col in date_columns:
        if col in data.columns:
            data[col] = pd.to_datetime(data[col], format="%d-%B-%Y", errors="coerce")

    data[["Copyright", "Issues", "Last Edit (Revision)", "Trustpilot Review Date"]] = data[
        ["Copyright", "Issues", "Last Edit (Revision)", "Trustpilot Review Date"]].astype(str)

    data[["Copyright", "Issues", "Last Edit (Revision)", "Trustpilot Review Date"]] = data[
        ["Copyright", "Issues", "Last Edit (Revision)", "Trustpilot Review Date"]].fillna("N/A")

    return data


def load_data(sheet_name: str, month_number: int, year: int) -> pd.DataFrame:
    """Load data from Google Sheets with optional month filtering"""
    try:
        data = get_sheet_data(sheet_name)
        data = clean_data(data)

        if "Publishing Date" in data.columns:
            data = data[(data["Publishing Date"].dt.month == month_number) & (data["Publishing Date"].dt.year == year)]

        if data.empty:
            return pd.DataFrame()

        data = data.sort_values(by="Publishing Date", ascending=True)

        for col in ["Publishing Date", "Last Edit (Revision)", "Trustpilot Review Date"]:
            if col in data.columns:
                data[col] = pd.to_datetime(data[col], errors="coerce").dt.strftime("%d-%B-%Y")
        # if "Name" in data.columns:
        #     data = data.drop_duplicates(subset=["Name"])
        data.index = range(1, len(data) + 1)
        return data

    except Exception as e:
        st.error(f"Error loading data: {e}")
        logging.error(f"An Error Occurred: {e}")
        return pd.DataFrame()


def load_data_year(sheet_name: str, year: int) -> pd.DataFrame:
    """Load data from Google Sheets with optional month filtering"""
    try:
        data = get_sheet_data(sheet_name)
        data = clean_data(data)

        if "Publishing Date" in data.columns:
            data = data[data["Publishing Date"].dt.year == year]

        if data.empty:
            return pd.DataFrame()

        data = data.sort_values(by="Publishing Date", ascending=True)

        for col in ["Publishing Date", "Last Edit (Revision)", "Trustpilot Review Date"]:
            if col in data.columns:
                data[col] = pd.to_datetime(data[col], errors="coerce").dt.strftime("%d-%B-%Y")
        data.index = range(1, len(data) + 1)
        return data

    except Exception as e:
        st.error(f"Error loading data: {e}")
        logging.error(f"An Error Occurred: {e}")
        return pd.DataFrame()

def load_data_search(sheet_name: str, end_year: int, start_year: int = get_min_year()) -> pd.DataFrame:
    """Load data from Google Sheets with optional month filtering"""
    try:
        data = get_sheet_data(sheet_name)
        data = clean_data(data)

        if "Publishing Date" in data.columns:
            data = data[
                (data["Publishing Date"].dt.year >= start_year) &
                (data["Publishing Date"].dt.year <= end_year)
            ]

        if data.empty:
            return pd.DataFrame()

        data = data.sort_values(by="Publishing Date", ascending=True)

        for col in ["Publishing Date", "Last Edit (Revision)", "Trustpilot Review Date"]:
            if col in data.columns:
                data[col] = pd.to_datetime(data[col], errors="coerce").dt.strftime("%d-%B-%Y")
        data.index = range(1, len(data) + 1)
        return data

    except Exception as e:
        st.error(f"Error loading data: {e}")
        logging.error(f"An Error Occurred: {e}")
        return pd.DataFrame()

def load_data_filter(sheet_name: str, start_date: datetime, end_date: datetime, remove_duplicates: bool = False) -> pd.DataFrame:
    """Load data from Google Sheets with optional month filtering"""
    try:
        data = get_sheet_data(sheet_name)
        data = clean_data(data)
        if remove_duplicates:
            data = load_data_search(sheet_name, end_date.year, start_date.year)
            data = data.drop_duplicates(subset=["Name"], keep="first")
            for col in ["Publishing Date", "Last Edit (Revision)", "Trustpilot Review Date"]:
                if col in data.columns:
                    data[col] = pd.to_datetime(data[col], errors="coerce")
        if "Publishing Date" in data.columns:
            data = data[
                (data["Publishing Date"].dt.date >= start_date) &
                (data["Publishing Date"].dt.date <= end_date)
            ]

        if data.empty:
            return pd.DataFrame()

        data = data.sort_values(by="Publishing Date", ascending=True)

        for col in ["Publishing Date", "Last Edit (Revision)", "Trustpilot Review Date"]:
            if col in data.columns:
                data[col] = pd.to_datetime(data[col], errors="coerce").dt.strftime("%d-%B-%Y")
        data.index = range(1, len(data) + 1)
        return data

    except Exception as e:
        st.error(f"Error loading data: {e}")
        logging.error(f"An Error Occurred: {e}")
        return pd.DataFrame()

def load_reviews(sheet_name: str, year: int, month_number=None) -> pd.DataFrame:
    data = get_sheet_data(sheet_name)
    if data.empty:
        return pd.DataFrame()

    columns = list(data.columns)
    if "Issues" in columns:
        end_col_index = columns.index("Issues")
        data = data.iloc[:, :end_col_index + 1]
    date_columns = ["Publishing Date", "Last Edit (Revision)", "Trustpilot Review Date"]
    for col in date_columns:
        if col in data.columns:
            data[col] = pd.to_datetime(data[col], format="%d-%B-%Y", errors="coerce")

    data[["Copyright", "Issues", "Last Edit (Revision)"]] = data[
        ["Copyright", "Issues", "Last Edit (Revision)"]].astype(str)

    data[["Copyright", "Issues", "Last Edit (Revision)"]] = data[
        ["Copyright", "Issues", "Last Edit (Revision)"]].fillna("N/A")
    try:
        if "Trustpilot Review Date" in data.columns and month_number:
            data = data[(data["Trustpilot Review Date"].dt.month == month_number) & (
                    data["Trustpilot Review Date"].dt.year == year)]
        else:
            data = data[(data["Trustpilot Review Date"].dt.year == year)]

        if data.empty:
            return pd.DataFrame()

        data = data.sort_values(by="Trustpilot Review Date", ascending=True)

        if "Name" in data.columns:
            if "Trustpilot Review" in data.columns and "Trustpilot Review Date" in data.columns:
                data = (
                    data.sort_values(
                        by=["Trustpilot Review", "Trustpilot Review Date"],
                        key=lambda col: col.eq("Attained") if col.name == "Trustpilot Review" else col,
                        ascending=[False, False]
                    )
                    .drop_duplicates(subset=["Name"], keep="last")
                )
            else:
                data = data.drop_duplicates(subset=["Name"])
        data.index = range(1, len(data) + 1)
        return data
    except Exception as e:
        st.error(f"Error loading data: {e}")
        logging.error(f"An Error Occurred: {e}")
        return pd.DataFrame()


def load_reviews_year(sheet_name: str, year: int, name: str, type_: str = "Attained") -> pd.DataFrame:
    data = get_sheet_data(sheet_name)
    if data.empty:
        return pd.DataFrame()

    columns = list(data.columns)
    if "Issues" in columns:
        end_col_index = columns.index("Issues")
        data = data.iloc[:, :end_col_index + 1]
    date_columns = ["Publishing Date", "Last Edit (Revision)", "Trustpilot Review Date"]
    for col in date_columns:
        if col in data.columns:
            data[col] = pd.to_datetime(data[col], format="%d-%B-%Y", errors="coerce")

    data[["Copyright", "Issues", "Last Edit (Revision)"]] = data[
        ["Copyright", "Issues", "Last Edit (Revision)"]].astype(str)

    data[["Copyright", "Issues", "Last Edit (Revision)"]] = data[
        ["Copyright", "Issues", "Last Edit (Revision)"]].fillna("N/A")
    try:
        if "Trustpilot Review Date" in data.columns:
            data = data[(data["Trustpilot Review Date"].dt.year == year)]

        else:
            return pd.DataFrame()

        data = data.sort_values(by="Trustpilot Review Date", ascending=True)

        data_original = data.copy()
        data = data_original[
            (data_original["Project Manager"] == name) &
            (data_original["Trustpilot Review"] == type_) &
            (data_original["Brand"].isin(
                ["BookMarketeers", "Writers Clique", "Authors Solution", "Book Publication", "Aurora Writers"]))
            ]

        data = data.sort_values(by="Trustpilot Review Date", ascending=True)
        data = data.drop_duplicates(subset=["Name"])
        data.index = range(1, len(data) + 1)
        return data
    except Exception as e:
        st.error(f"Error loading data: {e}")
        logging.error(f"An Error Occurred: {e}")
        return pd.DataFrame()

def load_reviews_year_to_date(sheet_name: str, year: int, name: str, type_: str = "Attained") -> pd.DataFrame:
    data = get_sheet_data(sheet_name)
    if data.empty:
        return pd.DataFrame()

    columns = list(data.columns)
    if "Issues" in columns:
        end_col_index = columns.index("Issues")
        data = data.iloc[:, :end_col_index + 1]
    date_columns = ["Publishing Date", "Last Edit (Revision)", "Trustpilot Review Date"]
    for col in date_columns:
        if col in data.columns:
            data[col] = pd.to_datetime(data[col], format="%d-%B-%Y", errors="coerce")

    data[["Copyright", "Issues", "Last Edit (Revision)"]] = data[
        ["Copyright", "Issues", "Last Edit (Revision)"]].astype(str)

    data[["Copyright", "Issues", "Last Edit (Revision)"]] = data[
        ["Copyright", "Issues", "Last Edit (Revision)"]].fillna("N/A")
    try:
        if "Trustpilot Review Date" in data.columns:
            data = data[
                (data["Trustpilot Review Date"].dt.year >= get_min_year()) &
                (data["Trustpilot Review Date"].dt.year <= year)
            ]

        else:
            return pd.DataFrame()

        data = data.sort_values(by="Trustpilot Review Date", ascending=True)

        data_original = data.copy()
        data = data_original[
            (data_original["Project Manager"] == name) &
            (data_original["Trustpilot Review"] == type_) &
            (data_original["Brand"].isin(
                ["BookMarketeers", "Writers Clique", "Authors Solution", "Book Publication", "Aurora Writers"]))
            ]

        data = data.sort_values(by="Trustpilot Review Date", ascending=True)
        data = data.drop_duplicates(subset=["Name"])
        data.index = range(1, len(data) + 1)
        return data
    except Exception as e:
        st.error(f"Error loading data: {e}")
        logging.error(f"An Error Occurred: {e}")
        return pd.DataFrame()

def load_reviews_filter(sheet_name: str, start_date: datetime, end_date: datetime, name: str, type_: str = "Attained") -> pd.DataFrame:
    data = get_sheet_data(sheet_name)
    if data.empty:
        return pd.DataFrame()

    columns = list(data.columns)
    if "Issues" in columns:
        end_col_index = columns.index("Issues")
        data = data.iloc[:, :end_col_index + 1]
    date_columns = ["Publishing Date", "Last Edit (Revision)", "Trustpilot Review Date"]
    for col in date_columns:
        if col in data.columns:
            data[col] = pd.to_datetime(data[col], format="%d-%B-%Y", errors="coerce")

    data[["Copyright", "Issues", "Last Edit (Revision)"]] = data[
        ["Copyright", "Issues", "Last Edit (Revision)"]].astype(str)

    data[["Copyright", "Issues", "Last Edit (Revision)"]] = data[
        ["Copyright", "Issues", "Last Edit (Revision)"]].fillna("N/A")
    try:
        if "Trustpilot Review Date" in data.columns:
            data = data[
                (data["Trustpilot Review Date"].dt.date >= start_date) &
                (data["Trustpilot Review Date"].dt.date <= end_date)
            ]

        else:
            return pd.DataFrame()

        data = data.sort_values(by="Trustpilot Review Date", ascending=True)

        data_original = data.copy()
        data = data_original[
            (data_original["Project Manager"] == name) &
            (data_original["Trustpilot Review"] == type_) &
            (data_original["Brand"].isin(
                ["BookMarketeers", "Writers Clique", "Authors Solution", "Book Publication", "Aurora Writers"]))
            ]

        data = data.sort_values(by="Trustpilot Review Date", ascending=True)
        data = data.drop_duplicates(subset=["Name"])
        data.index = range(1, len(data) + 1)
        return data
    except Exception as e:
        st.error(f"Error loading data: {e}")
        logging.error(f"An Error Occurred: {e}")
        return pd.DataFrame()

def load_reviews_year_multiple(sheet_name: str, start_year: int, end_year: int, name: str, type_: str = "Attained") -> pd.DataFrame:
    data = get_sheet_data(sheet_name)
    if data.empty:
        return pd.DataFrame()

    columns = list(data.columns)
    if "Issues" in columns:
        end_col_index = columns.index("Issues")
        data = data.iloc[:, :end_col_index + 1]
    date_columns = ["Publishing Date", "Last Edit (Revision)", "Trustpilot Review Date"]
    for col in date_columns:
        if col in data.columns:
            data[col] = pd.to_datetime(data[col], format="%d-%B-%Y", errors="coerce")

    data[["Copyright", "Issues", "Last Edit (Revision)"]] = data[
        ["Copyright", "Issues", "Last Edit (Revision)"]].astype(str)

    data[["Copyright", "Issues", "Last Edit (Revision)"]] = data[
        ["Copyright", "Issues", "Last Edit (Revision)"]].fillna("N/A")
    try:
        if "Trustpilot Review Date" in data.columns:
            data = data[
                (data["Trustpilot Review Date"].dt.year >= start_year) &
                (data["Trustpilot Review Date"].dt.year <= end_year)

            ]

        else:
            return pd.DataFrame()

        data = data.sort_values(by="Trustpilot Review Date", ascending=True)

        data_original = data.copy()
        data = data_original[
            (data_original["Project Manager"] == name) &
            (data_original["Trustpilot Review"] == type_) &
            (data_original["Brand"].isin(
                ["BookMarketeers", "Writers Clique", "Authors Solution", "Book Publication", "Aurora Writers"]))
            ]

        data = data.sort_values(by="Trustpilot Review Date", ascending=True)
        data = data.drop_duplicates(subset=["Name"])
        data.index = range(1, len(data) + 1)
        return data
    except Exception as e:
        st.error(f"Error loading data: {e}")
        logging.error(f"An Error Occurred: {e}")
        return pd.DataFrame()

def clean_data_reviews(sheet_name: str) -> pd.DataFrame:
    """Clean the data from Google Sheets"""
    data = get_sheet_data(sheet_name)

    if data.empty:
        return data

    columns = list(data.columns)
    if "Issues" in columns:
        end_col_index = columns.index("Issues")
        data = data.iloc[:, :end_col_index + 1]

    for col in ["Publishing Date", "Last Edit (Revision)", "Trustpilot Review Date"]:
        if col in data.columns:
            data[col] = pd.to_datetime(data[col], format="%d-%B-%Y", errors="coerce")

    data = data.sort_values(by="Publishing Date", ascending=True)
    data.index = range(1, len(data) + 1)

    return data


def get_printing_data_month(month: int, year: int) -> pd.DataFrame:
    """Get printing data for the current month"""
    data = get_sheet_data(sheet_printing)

    if data.empty:
        return pd.DataFrame()

    columns = list(data.columns)
    if "Accepted" in columns:
        end_col_index = columns.index("Accepted")
        data = data.iloc[:, :end_col_index + 1]
        data = data.astype(str)

    for col in ["Order Date", "Shipping Date", "Fulfilled"]:
        if col in data.columns:
            data[col] = pd.to_datetime(data[col], format="%d-%B-%Y", errors="coerce")

    data = data[(data["Order Date"].dt.month == month) & (data["Order Date"].dt.year == year)]

    data = data.sort_values(by="Order Date", ascending=True)
    if "Order Cost" in data.columns:
        data["Order Cost"] = data["Order Cost"].fillna(0)
        data["Order Cost"] = data["Order Cost"].astype(str)
        data["Order Cost"] = pd.to_numeric(
            data["Order Cost"].str.replace("$", "", regex=False).str.replace(",", "", regex=False),
            errors="coerce").fillna(0)

    if "No of Copies" in data.columns:
        data["No of Copies"] = pd.to_numeric(data["No of Copies"], errors='coerce').fillna(0)

    data = data.sort_values(by="Order Date", ascending=True)

    for col in ["Order Date", "Shipping Date", "Fulfilled"]:
        if col in data.columns:
            data[col] = pd.to_datetime(data[col], errors="coerce").dt.strftime("%d-%B-%Y")

    data.index = range(1, len(data) + 1)

    return data


def printing_data_year(year: int) -> tuple[pd.DataFrame, pd.DataFrame]:
    data = get_sheet_data(sheet_printing)

    if data.empty:
        return pd.DataFrame(), pd.DataFrame()

    columns = list(data.columns)
    if "Accepted" in columns:
        end_col_index = columns.index("Accepted")
        data = data.iloc[:, :end_col_index + 1]

    data = data.astype(str)

    for col in ["Order Date", "Shipping Date", "Fulfilled"]:
        if col in data.columns:
            data[col] = pd.to_datetime(data[col], format="%d-%B-%Y", errors="coerce")

    data = data[data["Order Date"].dt.year == year]

    data = data.sort_values(by="Order Date", ascending=True)
    if data.empty:
        return pd.DataFrame(), pd.DataFrame()

    if "Order Cost" in data.columns:
        data["Order Cost"] = pd.to_numeric(
            data["Order Cost"].str.replace("$", "", regex=False).str.replace(",", "", regex=False),
            errors="coerce"
        ).fillna(0)

    if "No of Copies" in data.columns:
        data["No of Copies"] = pd.to_numeric(data["No of Copies"], errors='coerce').fillna(0)

    data['Month'] = data['Order Date'].dt.to_period('M')

    month_totals = data.groupby('Month').agg(
        Total_Copies=('No of Copies', 'sum'),
        Total_Cost=('Order Cost', 'sum')
    ).reset_index()

    month_totals['Month'] = month_totals['Month'].dt.strftime('%B %Y')
    month_totals.columns = ["Month", "Total Copies", "Total Cost ($)"]
    month_totals = month_totals.sort_values(by="Total Cost ($)", ascending=False)
    month_totals.index = range(1, len(month_totals) + 1)
    month_totals["Total Cost ($)"] = month_totals["Total Cost ($)"].map("${:,.2f}".format)
    for col in ["Order Date", "Shipping Date", "Fulfilled"]:
        if col in data.columns:
            data[col] = data[col].dt.strftime("%d-%B-%Y")

    data.index = range(1, len(data) + 1)

    return data, month_totals

def printing_data_search(year: int) -> tuple[pd.DataFrame, pd.DataFrame]:
    data = get_sheet_data(sheet_printing)

    if data.empty:
        return pd.DataFrame(), pd.DataFrame()

    columns = list(data.columns)
    if "Accepted" in columns:
        end_col_index = columns.index("Accepted")
        data = data.iloc[:, :end_col_index + 1]

    data = data.astype(str)

    for col in ["Order Date", "Shipping Date", "Fulfilled"]:
        if col in data.columns:
            data[col] = pd.to_datetime(data[col], format="%d-%B-%Y", errors="coerce")

    data = data[
        (data["Order Date"].dt.year >= get_min_year()) &
        (data["Order Date"].dt.year <= year)

    ]

    data = data.sort_values(by="Order Date", ascending=True)
    if data.empty:
        return pd.DataFrame(), pd.DataFrame()

    if "Order Cost" in data.columns:
        data["Order Cost"] = pd.to_numeric(
            data["Order Cost"].str.replace("$", "", regex=False).str.replace(",", "", regex=False),
            errors="coerce"
        ).fillna(0)

    if "No of Copies" in data.columns:
        data["No of Copies"] = pd.to_numeric(data["No of Copies"], errors='coerce').fillna(0)

    data['Month'] = data['Order Date'].dt.to_period('M')

    month_totals = data.groupby('Month').agg(
        Total_Copies=('No of Copies', 'sum'),
        Total_Cost=('Order Cost', 'sum')
    ).reset_index()

    month_totals['Month'] = month_totals['Month'].dt.strftime('%B %Y')
    month_totals.columns = ["Month", "Total Copies", "Total Cost ($)"]
    month_totals = month_totals.sort_values(by="Total Cost ($)", ascending=False)
    month_totals.index = range(1, len(month_totals) + 1)
    month_totals["Total Cost ($)"] = month_totals["Total Cost ($)"].map("${:,.2f}".format)
    for col in ["Order Date", "Shipping Date", "Fulfilled"]:
        if col in data.columns:
            data[col] = data[col].dt.strftime("%d-%B-%Y")

    data.index = range(1, len(data) + 1)

    return data, month_totals

def printing_data_year_multiple(start_year: int, end_year: int) -> tuple[pd.DataFrame, pd.DataFrame]:
    data = get_sheet_data(sheet_printing)

    if data.empty:
        return pd.DataFrame(), pd.DataFrame()

    columns = list(data.columns)
    if "Accepted" in columns:
        end_col_index = columns.index("Accepted")
        data = data.iloc[:, :end_col_index + 1]

    data = data.astype(str)

    for col in ["Order Date", "Shipping Date", "Fulfilled"]:
        if col in data.columns:
            data[col] = pd.to_datetime(data[col], format="%d-%B-%Y", errors="coerce")

    data = data[
        (data["Order Date"].dt.year >= start_year) &
        (data["Order Date"].dt.year <= end_year)
    ]

    data = data.sort_values(by="Order Date", ascending=True)
    if data.empty:
        return pd.DataFrame(), pd.DataFrame()

    if "Order Cost" in data.columns:
        data["Order Cost"] = pd.to_numeric(
            data["Order Cost"].str.replace("$", "", regex=False).str.replace(",", "", regex=False),
            errors="coerce"
        ).fillna(0)

    if "No of Copies" in data.columns:
        data["No of Copies"] = pd.to_numeric(data["No of Copies"], errors='coerce').fillna(0)

    data['Month'] = data['Order Date'].dt.to_period('M')

    month_totals = data.groupby('Month').agg(
        Total_Copies=('No of Copies', 'sum'),
        Total_Cost=('Order Cost', 'sum')
    ).reset_index()

    month_totals['Month'] = month_totals['Month'].dt.strftime('%B %Y')
    month_totals.columns = ["Month", "Total Copies", "Total Cost ($)"]
    month_totals = month_totals.sort_values(by="Total Cost ($)", ascending=False)
    month_totals.index = range(1, len(month_totals) + 1)
    month_totals["Total Cost ($)"] = month_totals["Total Cost ($)"].map("${:,.2f}".format)
    for col in ["Order Date", "Shipping Date", "Fulfilled"]:
        if col in data.columns:
            data[col] = data[col].dt.strftime("%d-%B-%Y")

    data.index = range(1, len(data) + 1)

    return data, month_totals

def get_copyright_month(month: int, year: int) -> tuple[pd.DataFrame, int, int]:
    """Get copyright data for the current month"""
    data = get_sheet_data(sheet_copyright)

    if data.empty:
        return pd.DataFrame(), 0, 0

    columns = list(data.columns)
    if "Country" in columns:
        end_col_index = columns.index("Country")
        data = data.iloc[:, :end_col_index + 1]
    data = data.astype(str)

    if "Submission Date" in data.columns:
        data["Submission Date"] = pd.to_datetime(data["Submission Date"], format="%d-%B-%Y", errors='coerce')
        data = data[
            (data["Submission Date"].dt.month == month) & (data["Submission Date"].dt.year == year)]

    data = data.sort_values(by=["Submission Date"], ascending=True)
    result_count = len(data[data["Result"] == "Yes"]) if "Result" in data.columns else 0
    result_count_no = len(data[data["Result"] == "No"]) if "Result" in data.columns else 0
    if "Submission Date" in data.columns:
        data["Submission Date"] = data["Submission Date"].dt.strftime("%d-%B-%Y")

    data = data.fillna("N/A")

    data.index = range(1, len(data) + 1)

    return data, result_count, result_count_no


def copyright_year(year: int) -> tuple[pd.DataFrame, int, int]:
    data = get_sheet_data(sheet_copyright)

    if data.empty:
        return pd.DataFrame(), 0, 0

    columns = list(data.columns)
    if "Country" in columns:
        end_col_index = columns.index("Country")
        data = data.iloc[:, :end_col_index + 1]
    data = data.astype(str)

    if "Submission Date" in data.columns:
        data["Submission Date"] = pd.to_datetime(data["Submission Date"], format="%d-%B-%Y", errors='coerce')
        data = data[
            (data["Submission Date"].dt.year == year)]
    data = data.sort_values(by=["Submission Date"], ascending=True)

    result_count = len(data[data["Result"] == "Yes"]) if "Result" in data.columns else 0
    result_count_no = len(data[data["Result"] == "No"]) if "Result" in data.columns else 0
    if "Submission Date" in data.columns:
        data["Submission Date"] = data["Submission Date"].dt.strftime("%d-%B-%Y")

    data = data.fillna("N/A")

    data.index = range(1, len(data) + 1)

    return data, result_count, result_count_no

def copyright_search(year: int) -> tuple[pd.DataFrame, int, int]:
    data = get_sheet_data(sheet_copyright)

    if data.empty:
        return pd.DataFrame(), 0, 0

    columns = list(data.columns)
    if "Country" in columns:
        end_col_index = columns.index("Country")
        data = data.iloc[:, :end_col_index + 1]
    data = data.astype(str)

    if "Submission Date" in data.columns:
        data["Submission Date"] = pd.to_datetime(data["Submission Date"], format="%d-%B-%Y", errors='coerce')
        data = data[
            (data["Submission Date"].dt.year >= get_min_year()) &
            (data["Submission Date"].dt.year <= year)

        ]
    data = data.sort_values(by=["Submission Date"], ascending=True)

    result_count = len(data[data["Result"] == "Yes"]) if "Result" in data.columns else 0
    result_count_no = len(data[data["Result"] == "No"]) if "Result" in data.columns else 0
    if "Submission Date" in data.columns:
        data["Submission Date"] = data["Submission Date"].dt.strftime("%d-%B-%Y")

    data = data.fillna("N/A")

    data.index = range(1, len(data) + 1)

    return data, result_count, result_count_no

def copyright_year_multiple(start_year: int, end_year: int) -> tuple[pd.DataFrame, int, int]:
    data = get_sheet_data(sheet_copyright)

    if data.empty:
        return pd.DataFrame(), 0, 0

    columns = list(data.columns)
    if "Country" in columns:
        end_col_index = columns.index("Country")
        data = data.iloc[:, :end_col_index + 1]
    data = data.astype(str)

    if "Submission Date" in data.columns:
        data["Submission Date"] = pd.to_datetime(data["Submission Date"], format="%d-%B-%Y", errors='coerce')
        data = data[
            (data["Submission Date"].dt.year >= start_year) &
            (data["Submission Date"].dt.year <= end_year)

        ]
    data = data.sort_values(by=["Submission Date"], ascending=True)

    result_count = len(data[data["Result"] == "Yes"]) if "Result" in data.columns else 0
    result_count_no = len(data[data["Result"] == "No"]) if "Result" in data.columns else 0
    if "Submission Date" in data.columns:
        data["Submission Date"] = data["Submission Date"].dt.strftime("%d-%B-%Y")

    data = data.fillna("N/A")

    data.index = range(1, len(data) + 1)

    return data, result_count, result_count_no

def get_A_plus_month(month: int, year: int) -> tuple[pd.DataFrame, int]:
    data = get_sheet_data(sheet_a_plus)
    if data.empty:
        return pd.DataFrame(), 0

    columns = list(data.columns)
    if "Issues" in columns:
        end_col_index = columns.index("Issues")
        data = data.iloc[:, :end_col_index + 1]
    data = data.astype(str)

    if "A+ Content Date" in data.columns:
        data["A+ Content Date"] = pd.to_datetime(data["A+ Content Date"], format="%d-%B-%Y", errors='coerce')
        data = data[
            (data["A+ Content Date"].dt.month == month) & (data["A+ Content Date"].dt.year == year)]
    data = data.sort_values(by=["A+ Content Date"], ascending=True)

    result_count = len(data[data["Status"] == "Published"]) if "Status" in data.columns else 0

    if "A+ Content Date" in data.columns:
        data["A+ Content Date"] = data["A+ Content Date"].dt.strftime("%d-%B-%Y")

    data = data.fillna("N/A")

    data.index = range(1, len(data) + 1)

    return data, result_count


def get_A_plus_year(year: int) -> tuple[pd.DataFrame, int]:
    data = get_sheet_data(sheet_a_plus)
    if data.empty:
        return pd.DataFrame(), 0

    columns = list(data.columns)
    if "Issues" in columns:
        end_col_index = columns.index("Issues")
        data = data.iloc[:, :end_col_index + 1]
    data = data.astype(str)

    if "A+ Content Date" in data.columns:
        data["A+ Content Date"] = pd.to_datetime(data["A+ Content Date"], format="%d-%B-%Y", errors='coerce')
        data = data[
            (data["A+ Content Date"].dt.year == year)]
    data = data.sort_values(by=["A+ Content Date"], ascending=True)

    result_count = len(data[data["Status"] == "Published"]) if "Status" in data.columns else 0

    if "A+ Content Date" in data.columns:
        data["A+ Content Date"] = data["A+ Content Date"].dt.strftime("%d-%B-%Y")

    data = data.fillna("N/A")

    data.index = range(1, len(data) + 1)

    return data, result_count

def get_A_plus_year_multiple(start_year: int, end_year: int) -> tuple[pd.DataFrame, int]:
    data = get_sheet_data(sheet_a_plus)
    if data.empty:
        return pd.DataFrame(), 0

    columns = list(data.columns)
    if "Issues" in columns:
        end_col_index = columns.index("Issues")
        data = data.iloc[:, :end_col_index + 1]
    data = data.astype(str)

    if "A+ Content Date" in data.columns:
        data["A+ Content Date"] = pd.to_datetime(data["A+ Content Date"] ,errors='coerce')
        data = data[
            (data["A+ Content Date"].dt.year >= start_year) &
            (data["A+ Content Date"].dt.year <= end_year)
        ]
    data = data.sort_values(by=["A+ Content Date"], ascending=True)

    result_count = len(data[data["Status"] == "Published"]) if "Status" in data.columns else 0

    if "A+ Content Date" in data.columns:
        data["A+ Content Date"] = data["A+ Content Date"].dt.strftime("%d-%B-%Y")

    data = data.fillna("N/A")

    data.index = range(1, len(data) + 1)

    return data, result_count

def get_names_in_both_months(sheet_name: str, month_1: str, year1: int, month_2: str, year2: int) -> tuple:
    """
    Identifies names that appear in both June and July from a Google Sheet.
    Returns:
        - A set of matching names
        - A dictionary with individual counts for June and July
    """
    df = get_sheet_data(sheet_name)

    if df.empty or "Name" not in df.columns or "Publishing Date" not in df.columns:
        logging.warning("Missing 'Name' or 'Date' columns or data is empty.")
        return set(), {}, 0

    df['Publishing Date'] = pd.to_datetime(df['Publishing Date'], format="%d-%B-%Y", errors='coerce')
    df = df.dropna(subset=['Publishing Date', 'Name'])

    df['Month'] = df['Publishing Date'].dt.month_name()
    df['Year'] = df['Publishing Date'].dt.year

    month_1_names = set(
        df[(df['Month'] == month_1) & (df['Year'] == year1)]['Name'].str.strip()
    )

    month_2_names = set(
        df[(df['Month'] == month_2) & (df['Year'] == year2)]['Name'].str.strip()
    )

    if month_1_names & month_2_names:
        names_in_both = month_1_names.intersection(month_2_names)

        counts = {}
        for name in names_in_both:
            month_1_count = df[
                (df['Month'] == month_1) &
                (df['Year'] == year1) &
                (df['Name'].str.strip() == name)
                ].shape[0]

            month_2_count = df[
                (df['Month'] == month_2) &
                (df['Year'] == year2) &
                (df['Name'].str.strip() == name)
                ].shape[0]

            counts[name] = {
                f"{month_1}-{year1}": month_1_count,
                f"{month_2}-{year2}": month_2_count,
            }

        return names_in_both, counts, len(names_in_both)
    else:
        return set(), {}, 0

def get_names_in_both_years(sheet_name: str, year1: int, year2: int) -> tuple:
    """
    Identifies names that appear in both years from a Google Sheet.
    """
    df = get_sheet_data(sheet_name)

    if df.empty or "Name" not in df.columns or "Publishing Date" not in df.columns:
        logging.warning("Missing 'Name' or 'Publishing Date' columns or data is empty.")
        return set(), {}, 0

    df['Publishing Date'] = pd.to_datetime(
        df['Publishing Date'], format="%d-%B-%Y", errors='coerce'
    )
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
                "publishing_dates": year1_df['Publishing Date']
                .dt.strftime("%d-%B-%Y")
                .tolist()
            },
            str(year2): {
                "count": year2_df.shape[0],
                "publishing_dates": year2_df['Publishing Date']
                .dt.strftime("%d-%B-%Y")
                .tolist()
            }
        }

    return names_in_both, counts, len(names_in_both)


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

    df['Publishing Date'] = pd.to_datetime(df['Publishing Date'], format="%d-%B-%Y", errors='coerce')
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


def create_review_pie_chart(review_data: dict[str, int], title: str):
    """Create pie chart for review distribution"""
    global labels, values
    if isinstance(review_data, dict):
        if not review_data or sum(review_data.values()) == 0:
            return None
        values = list(review_data.values())
        labels = list(review_data.keys())

    custom_colors = {
        "Attained": "#7dff8d",
        "Pending": "#ffc444",
        "Negative": "#ff4b4b",
        "Sent": "#77e5f7"
    }

    fig = px.pie(
        values=values,
        names=labels,
        title=title,
        color=labels,
        color_discrete_map=custom_colors
    )
    fig.update_traces(textposition="inside", textinfo="percent+label")
    return fig


def create_platform_comparison_chart(usa_data: dict[str, int], uk_data: dict[str, int]):
    """Create comparison chart for platforms"""
    platforms = ['Amazon', 'Barnes & Noble', 'Ingram Spark', 'FAV']

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


def create_brand_chart(usa_brands: dict[str, int], uk_brands: dict[str, int]):
    """Create brand distribution chart"""
    all_brands = list(usa_brands.keys()) + list(uk_brands.keys())
    all_values = list(usa_brands.values()) + list(uk_brands.values())
    regions = ['USA'] * len(usa_brands) + ['UK'] * len(uk_brands)

    fig = px.bar(
        x=all_brands,
        y=all_values,
        color=regions,
        title='Brand Distribution by Region',
        color_discrete_map={'USA': '#23A0F8', 'UK': '#ff7f0e'}
    )
    return fig


def safe_concat(dfs):
    dfs = [df for df in dfs if not df.empty]
    return pd.concat(dfs, ignore_index=True) if dfs else pd.DataFrame()


def summary(month: int, year: int):
    uk_clean = clean_data_reviews(sheet_uk)
    usa_clean = clean_data_reviews(sheet_usa)

    usa_clean = usa_clean[
        (usa_clean["Publishing Date"].dt.month == month) &
        (usa_clean["Publishing Date"].dt.year == year)
        ]
    uk_clean = uk_clean[
        (uk_clean["Publishing Date"].dt.month == month) &
        (uk_clean["Publishing Date"].dt.year == year)
        ]

    usa_clean_platforms = usa_clean[
        (usa_clean["Publishing Date"].dt.month == month) &
        (usa_clean["Publishing Date"].dt.year == year)
        ]
    uk_clean_platforms = uk_clean[
        (uk_clean["Publishing Date"].dt.month == month) &
        (uk_clean["Publishing Date"].dt.year == year)
        ]

    if usa_clean.empty:
        print("No values found in USA sheet.")
        return
    if uk_clean.empty:
        print("No values found in UK sheet.")
        return
    if usa_clean.empty and uk_clean.empty:
        return
    usa_clean = usa_clean.drop_duplicates(subset=["Name"], keep="last")
    uk_clean = uk_clean.drop_duplicates(subset=["Name"], keep="last")
    Issues_usa = usa_clean["Issues"].value_counts()
    Issues_uk = uk_clean["Issues"].value_counts()
    total_usa = usa_clean["Name"].nunique()
    total_uk = uk_clean["Name"].nunique()
    total_unique_clients = total_usa + total_uk

    combined = pd.concat([usa_clean[["Name", "Brand", "Project Manager", "Email"]],
                          uk_clean[["Name", "Brand", "Project Manager", "Email"]]])
    combined.index = range(1, len(combined) + 1)

    brands = usa_clean["Brand"].value_counts()
    writers_clique = brands.get("Writers Clique", 0)
    bookmarketeers = brands.get("BookMarketeers", 0)
    aurora_writers = brands.get("Aurora Writers", 0)
    kdp = brands.get("KDP", 0)

    uk_brand = uk_clean["Brand"].value_counts()
    authors_solution = uk_brand.get("Authors Solution", 0)
    book_publication = uk_brand.get("Book Publication", 0)

    usa_platforms = usa_clean_platforms["Platform"].value_counts()
    usa_amazon = usa_platforms.get("Amazon", 0)
    usa_bn = usa_platforms.get("Barnes & Noble", 0)
    usa_ingram = usa_platforms.get("Ingram Spark", 0)
    usa_d2d = usa_platforms.get("Draft2Digital", 0)
    usa_fav = usa_platforms.get("FAV", 0)
    usa_acx = usa_platforms.get("ACX", 0)

    uk_platforms = uk_clean_platforms["Platform"].value_counts()
    uk_amazon = uk_platforms.get("Amazon", 0)
    uk_bn = uk_platforms.get("Barnes & Noble", 0)
    uk_ingram = uk_platforms.get("Ingram Spark", 0)
    uk_d2d = uk_platforms.get("Draft2Digital", 0)
    uk_fav = uk_platforms.get("FAV", 0)
    uk_kobo = uk_platforms.get("Kobo", 0)
    uk_acx = uk_platforms.get("ACX", 0)

    allowed_brands = ["BookMarketeers", "Writers Clique", "Aurora Writers", "Authors Solution", "Book Publication"]

    if "Trustpilot Review" in usa_clean.columns and "Brand" in usa_clean.columns:
        usa_filtered = usa_clean[usa_clean["Brand"].isin(allowed_brands)]
        usa_review_sent = usa_filtered["Trustpilot Review"].value_counts().get("Sent", 0)
        usa_review_pending = usa_filtered["Trustpilot Review"].value_counts().get("Pending", 0)
        usa_review_na = usa_filtered["Trustpilot Review"].value_counts().get("Negative", 0)
    else:
        usa_review_sent = usa_review_pending = usa_review_na = 0

    if "Trustpilot Review" in uk_clean.columns and "Brand" in uk_clean.columns:
        uk_filtered = uk_clean[uk_clean["Brand"].isin(allowed_brands)]
        uk_review_sent = uk_filtered["Trustpilot Review"].value_counts().get("Sent", 0)
        uk_review_pending = uk_filtered["Trustpilot Review"].value_counts().get("Pending", 0)
        uk_review_na = uk_filtered["Trustpilot Review"].value_counts().get("Negative", 0)
    else:
        uk_review_sent = uk_review_pending = uk_review_na = 0
    combined_pending_sent = pd.concat([usa_clean, uk_clean], ignore_index=True)
    pending_sent_details = combined_pending_sent[
        ((combined_pending_sent["Trustpilot Review"] == "Sent") |
         (combined_pending_sent["Trustpilot Review"] == "Pending")) &
        (combined_pending_sent["Brand"].isin(allowed_brands))
        ]
    pending_sent_details = pending_sent_details[["Name", "Brand", "Project Manager", "Trustpilot Review", "Status"]]
    pending_sent_details.index = range(1, len(pending_sent_details) + 1)

    usa_reviews_df = load_reviews(sheet_usa, year, month)
    uk_reviews_df = load_reviews(sheet_uk, year, month)
    combined_data = safe_concat([usa_reviews_df, uk_reviews_df])

    if not usa_reviews_df.empty:
        usa_attained_pm = (
            usa_reviews_df[usa_reviews_df["Trustpilot Review"] == "Attained"]
            .groupby("Project Manager")["Trustpilot Review"]
            .count()
            .reset_index()
        )
        usa_attained_pm.columns = ["Project Manager", "Attained Reviews"]
        usa_attained_pm.index = range(1, len(usa_attained_pm) + 1)
        usa_total_attained = usa_attained_pm["Attained Reviews"].sum()

        usa_negative_pm = (
            usa_reviews_df[usa_reviews_df["Trustpilot Review"] == "Negative"]
            .groupby("Project Manager")["Trustpilot Review"]
            .count()
            .reset_index()
        )
        usa_negative_pm.columns = ["Project Manager", "Negative Reviews"]
        usa_negative_pm.index = range(1, len(usa_negative_pm) + 1)
        usa_total_negative = usa_negative_pm["Negative Reviews"].sum()
    else:
        usa_attained_pm = pd.DataFrame(columns=["Project Manager", "Attained Reviews"])
        usa_negative_pm = pd.DataFrame(columns=["Project Manager", "Negative Reviews"])
        usa_total_attained = 0
        usa_total_negative = 0

    if not uk_reviews_df.empty:
        uk_attained_pm = (
            uk_reviews_df[uk_reviews_df["Trustpilot Review"] == "Attained"]
            .groupby("Project Manager")["Trustpilot Review"]
            .count()
            .reset_index()
        )
        uk_attained_pm.columns = ["Project Manager", "Attained Reviews"]
        uk_attained_pm.index = range(1, len(uk_attained_pm) + 1)
        uk_total_attained = uk_attained_pm["Attained Reviews"].sum()

        uk_negative_pm = (
            uk_reviews_df[uk_reviews_df["Trustpilot Review"] == "Negative"]
            .groupby("Project Manager")["Trustpilot Review"]
            .count()
            .reset_index()
        )
        uk_negative_pm.columns = ["Project Manager", "Negative Reviews"]
        uk_negative_pm.index = range(1, len(uk_negative_pm) + 1)
        uk_total_negative = uk_negative_pm["Negative Reviews"].sum()
    else:
        uk_attained_pm = pd.DataFrame(columns=["Project Manager", "Attained Reviews"])
        uk_negative_pm = pd.DataFrame(columns=["Project Manager", "Negative Reviews"])
        uk_total_attained = 0
        uk_total_negative = 0

    if not combined_data.empty:
        attained_reviews_per_pm = (
            combined_data[combined_data["Trustpilot Review"] == "Attained"]
            .groupby("Project Manager")["Trustpilot Review"]
            .count()
            .reset_index(name="Attained Reviews")
        )
        attained_reviews_per_pm = attained_reviews_per_pm.sort_values(by="Attained Reviews", ascending=False)
        attained_reviews_per_pm.index = range(1, len(attained_reviews_per_pm) + 1)

        negative_reviews_per_pm = (
            combined_data[combined_data["Trustpilot Review"] == "Negative"]
            .groupby("Project Manager")["Trustpilot Review"]
            .count()
            .reset_index(name="Negative Reviews")
        )
        negative_reviews_per_pm = negative_reviews_per_pm.sort_values(by="Negative Reviews", ascending=False)
        negative_reviews_per_pm.index = range(1, len(negative_reviews_per_pm) + 1)

        review_details_df = combined_data.sort_values(by="Project Manager", ascending=True)
        review_details_df["Trustpilot Review Date"] = pd.to_datetime(
            review_details_df["Trustpilot Review Date"], errors="coerce"
        ).dt.strftime("%d-%B-%Y")

        attained_details = review_details_df[
            review_details_df["Trustpilot Review"] == "Attained"
            ][["Project Manager", "Name", "Brand", "Trustpilot Review Date", "Trustpilot Review Links", "Status"]]
        attained_details.index = range(1, len(attained_details) + 1)

        negative_details = review_details_df[
            review_details_df["Trustpilot Review"] == "Negative"
            ][["Project Manager", "Name", "Brand", "Trustpilot Review Date", "Trustpilot Review Links", "Status"]]
        negative_details.index = range(1, len(negative_details) + 1)

    else:
        attained_reviews_per_pm = pd.DataFrame(columns=["Project Manager", "Attained Reviews"])
        negative_reviews_per_pm = pd.DataFrame(columns=["Project Manager", "Negative Reviews"])
        attained_details = pd.DataFrame(
            columns=["Project Manager", "Name", "Brand", "Trustpilot Review Date", "Trustpilot Review Links", "Status"])
        negative_details = attained_details.copy()

    usa_review = {
        "Attained": usa_total_attained,
        "Sent": usa_review_sent,
        "Pending": usa_review_pending,
        "Negative": usa_review_na + usa_total_negative
    }

    uk_review = {
        "Attained": uk_total_attained,
        "Sent": uk_review_sent,
        "Pending": uk_review_pending,
        "Negative": uk_review_na + uk_total_negative
    }

    printing_data = get_printing_data_month(month, year)
    Total_copies = printing_data["No of Copies"].sum() if "No of Copies" in printing_data.columns else 0
    Total_cost = printing_data["Order Cost"].sum() if "Order Cost" in printing_data.columns else 0
    Highest_cost = printing_data["Order Cost"].max() if "Order Cost" in printing_data.columns else 0
    Highest_copies = printing_data["No of Copies"].max() if "No of Copies" in printing_data.columns else 0
    Lowest_cost = printing_data["Order Cost"].min() if "Order Cost" in printing_data.columns else 0
    Lowest_copies = printing_data["No of Copies"].min() if "No of Copies" in printing_data.columns else 0

    Average = Total_cost / Total_copies if Total_copies > 0 else 0
    if all(col in printing_data.columns for col in ["Order Cost", "No of Copies"]):
        printing_data['Cost_Per_Copy'] = printing_data['Order Cost'] / printing_data['No of Copies']

    copyright_data, result_count, result_count_no = get_copyright_month(month, year)
    Total_copyrights = len(copyright_data)

    country = copyright_data["Country"].value_counts()
    usa = country.get("USA", 0)
    canada = country.get("Canada", 0)
    uk = country.get("UK", 0)
    Total_cost_copyright = (usa * 65) + (canada * 63) + (uk * 35)
    a_plus, a_plus_count = get_A_plus_month(month, year)

    usa_brands = {'BookMarketeers': bookmarketeers, 'Writers Clique': writers_clique, 'KDP': kdp,
                  'Aurora Writers': aurora_writers}

    uk_brands = {'Authors Solution': authors_solution, 'Book Publication': book_publication}

    usa_platforms = {'Amazon': usa_amazon, 'Barnes & Noble': usa_bn, 'Ingram Spark': usa_ingram,"Draft2Digital": usa_d2d,  "FAV": usa_fav, "ACX":usa_acx}
    uk_platforms = {'Amazon': uk_amazon, 'Barnes & Noble': uk_bn, 'Ingram Spark': uk_ingram, "FAV": uk_fav,
                    "Kobo": uk_kobo,"Draft2Digital": uk_d2d, "ACX":uk_acx}

    printing_stats = {
        'Total_copies': Total_copies,
        'Total_cost': Total_cost,
        'Highest_cost': Highest_cost,
        'Lowest_cost': Lowest_cost,
        'Highest_copies': Highest_copies,
        'Lowest_copies': Lowest_copies,
        'Average': Average
    }

    copyright_stats = {
        'Total_copyrights': Total_copyrights,
        'Total_cost_copyright': Total_cost_copyright,
        'result_count': result_count,
        'result_count_no': result_count_no,
        'usa_copyrights': usa,
        'canada_copyrights': canada,
        'uk': uk
    }

    return usa_review, uk_review, usa_brands, uk_brands, usa_platforms, uk_platforms, printing_stats, copyright_stats, a_plus_count, total_unique_clients, combined, attained_reviews_per_pm, attained_details, pending_sent_details, negative_reviews_per_pm, negative_details, Issues_usa, Issues_uk


def generate_year_summary(year: int):
    uk_clean = clean_data_reviews(sheet_uk)
    usa_clean = clean_data_reviews(sheet_usa)

    usa_clean = usa_clean[
        (usa_clean["Publishing Date"].dt.year == year)
    ]
    uk_clean = uk_clean[
        (uk_clean["Publishing Date"].dt.year == year)
    ]

    usa_clean_platforms = usa_clean[
        (usa_clean["Publishing Date"].dt.year == year)
    ]
    uk_clean_platforms = uk_clean[
        (uk_clean["Publishing Date"].dt.year == year)
    ]

    if usa_clean.empty:
        print("No values found in USA sheet.")
    if uk_clean.empty:
        print("No values found in UK sheet.")
        return
    if usa_clean.empty and uk_clean.empty:
        return

    usa_clean = usa_clean.drop_duplicates(subset=["Name"], keep="first")
    uk_clean = uk_clean.drop_duplicates(subset=["Name"], keep="first")
    Issues_usa = usa_clean["Issues"].value_counts()
    Issues_uk = uk_clean["Issues"].value_counts()
    total_usa = usa_clean["Name"].nunique()
    total_uk = uk_clean["Name"].nunique()
    total_unique_clients = total_usa + total_uk

    combined = pd.concat([usa_clean[["Name", "Brand", "Project Manager", "Email"]],
                          uk_clean[["Name", "Brand", "Project Manager", "Email"]]])
    combined.index = range(1, len(combined) + 1)

    brands = usa_clean["Brand"].value_counts()
    writers_clique = brands.get("Writers Clique", 0)
    bookmarketeers = brands.get("BookMarketeers", 0)
    aurora_writers = brands.get("Aurora Writers", 0)
    kdp = brands.get("KDP", 0)

    uk_brand = uk_clean["Brand"].value_counts()
    authors_solution = uk_brand.get("Authors Solution", 0)
    book_publication = uk_brand.get("Book Publication", 0)

    usa_platforms = usa_clean_platforms["Platform"].value_counts()
    usa_amazon = usa_platforms.get("Amazon", 0)
    usa_bn = usa_platforms.get("Barnes & Noble", 0)
    usa_ingram = usa_platforms.get("Ingram Spark", 0)
    usa_d2d = usa_platforms.get("Draft2Digital", 0)
    usa_fav = usa_platforms.get("FAV", 0)
    usa_acx = usa_platforms.get("ACX", 0)

    uk_platforms = uk_clean_platforms["Platform"].value_counts()
    uk_amazon = uk_platforms.get("Amazon", 0)
    uk_bn = uk_platforms.get("Barnes & Noble", 0)
    uk_ingram = uk_platforms.get("Ingram Spark", 0)
    uk_d2d = uk_platforms.get("Draft2Digital", 0)
    uk_fav = uk_platforms.get("FAV", 0)
    uk_kobo = uk_platforms.get("Kobo", 0)
    uk_acx = uk_platforms.get("ACX", 0)

    allowed_brands = ["BookMarketeers", "Writers Clique", "Aurora Writers", "Authors Solution", "Book Publication"]

    if "Trustpilot Review" in usa_clean.columns and "Brand" in usa_clean.columns:
        usa_filtered = usa_clean[usa_clean["Brand"].isin(allowed_brands)]
        usa_review_sent = usa_filtered["Trustpilot Review"].value_counts().get("Sent", 0)
        usa_review_pending = usa_filtered["Trustpilot Review"].value_counts().get("Pending", 0)
        usa_review_na = usa_filtered["Trustpilot Review"].value_counts().get("Negative", 0)
    else:
        usa_review_sent = usa_review_pending = usa_review_na = 0

    if "Trustpilot Review" in uk_clean.columns and "Brand" in uk_clean.columns:
        uk_filtered = uk_clean[uk_clean["Brand"].isin(allowed_brands)]
        uk_review_sent = uk_filtered["Trustpilot Review"].value_counts().get("Sent", 0)
        uk_review_pending = uk_filtered["Trustpilot Review"].value_counts().get("Pending", 0)
        uk_review_na = uk_filtered["Trustpilot Review"].value_counts().get("Negative", 0)
    else:
        uk_review_sent = uk_review_pending = uk_review_na = 0

    combined_pending_sent = pd.concat([usa_clean, uk_clean], ignore_index=True)
    pending_sent_details = combined_pending_sent[
        ((combined_pending_sent["Trustpilot Review"] == "Sent") |
         (combined_pending_sent["Trustpilot Review"] == "Pending")) &
        (combined_pending_sent["Brand"].isin(allowed_brands))
        ]
    pending_sent_details = pending_sent_details[["Name", "Brand", "Project Manager", "Trustpilot Review", "Status"]]
    pending_sent_details.index = range(1, len(pending_sent_details) + 1)

    pm_list_usa = list(set((usa_clean["Project Manager"].dropna().unique().tolist() + ["Unknown"])))
    pm_list_uk = list(set((uk_clean["Project Manager"].dropna().unique().tolist() + ["Unknown"])))

    usa_reviews_per_pm = safe_concat([load_reviews_year(sheet_usa, year, pm, "Attained") for pm in pm_list_usa])
    uk_reviews_per_pm = safe_concat([load_reviews_year(sheet_uk, year, pm, "Attained") for pm in pm_list_uk])
    combined_data = safe_concat([usa_reviews_per_pm, uk_reviews_per_pm])

    usa_monthly = (
        usa_clean.groupby(usa_clean["Publishing Date"].dt.to_period("M"))
        .size()
        .reset_index(name="USA Published")
    )
    usa_monthly["Month"] = usa_monthly["Publishing Date"].dt.strftime("%B %Y")
    usa_monthly = usa_monthly[["Month", "USA Published"]]

    uk_monthly = (
        uk_clean.groupby(uk_clean["Publishing Date"].dt.to_period("M"))
        .size()
        .reset_index(name="UK Published")
    )
    uk_monthly["Month"] = uk_monthly["Publishing Date"].dt.strftime("%B %Y")
    uk_monthly = uk_monthly[["Month", "UK Published"]]

    combined_monthly = pd.merge(
        usa_monthly,
        uk_monthly,
        on="Month",
        how="outer"
    ).fillna(0)

    combined_monthly["Total Published"] = combined_monthly["USA Published"] + combined_monthly["UK Published"]

    combined_monthly["Month_Num"] = pd.to_datetime(combined_monthly["Month"], format="%B %Y")
    combined_monthly = combined_monthly.sort_values("Total Published", ascending=False).drop(columns="Month_Num")

    combined_monthly.index = range(1, len(combined_monthly) + 1)

    if not usa_reviews_per_pm.empty:
        usa_attained_pm = (
            usa_reviews_per_pm
            .groupby("Project Manager")["Trustpilot Review"]
            .count()
            .reset_index()
        )
        usa_attained_pm.columns = ["Project Manager", "Attained Reviews"]
        usa_attained_pm.index = range(1, len(usa_attained_pm) + 1)
        usa_total_attained = usa_attained_pm["Attained Reviews"].sum()
    else:
        usa_attained_pm = pd.DataFrame(columns=["Project Manager", "Attained Reviews"])
        usa_total_attained = 0

    if not uk_reviews_per_pm.empty:
        uk_attained_pm = (
            uk_reviews_per_pm
            .groupby("Project Manager")["Trustpilot Review"]
            .count()
            .reset_index()
        )
        uk_attained_pm.columns = ["Project Manager", "Attained Reviews"]
        uk_attained_pm.index = range(1, len(uk_attained_pm) + 1)
        uk_total_attained = uk_attained_pm["Attained Reviews"].sum()
    else:
        uk_attained_pm = pd.DataFrame(columns=["Project Manager", "Attained Reviews"])
        uk_total_attained = 0

    if not combined_data.empty:
        attained_reviews_per_pm = (
            combined_data
            .groupby("Project Manager")["Trustpilot Review"]
            .count()
            .reset_index()
        )
        attained_reviews_per_pm.columns = ["Project Manager", "Attained Reviews"]
        attained_reviews_per_pm = attained_reviews_per_pm.sort_values(by="Attained Reviews", ascending=False)
        attained_reviews_per_pm.index = range(1, len(attained_reviews_per_pm) + 1)

        review_details_df = combined_data.sort_values(by="Project Manager", ascending=True)
        review_details_df["Trustpilot Review Date"] = pd.to_datetime(
            review_details_df["Trustpilot Review Date"], errors="coerce"
        ).dt.strftime("%d-%B-%Y")

        attained_details = review_details_df[
            ["Project Manager", "Name", "Brand", "Trustpilot Review Date", "Trustpilot Review Links", "Status"]
        ]
        attained_details.index = range(1, len(attained_details) + 1)

        attained_details["Trustpilot Review Date"] = pd.to_datetime(
            attained_details["Trustpilot Review Date"], errors="coerce"
        )

        if not usa_reviews_per_pm.empty:
            usa_attained_monthly = (
                usa_reviews_per_pm.groupby(usa_reviews_per_pm["Trustpilot Review Date"].dt.to_period("M"))
                .size()
                .reset_index(name="USA Attained Reviews")
            )
            usa_attained_monthly["Month"] = usa_attained_monthly["Trustpilot Review Date"].dt.strftime("%B %Y")
            usa_attained_monthly = usa_attained_monthly[["Month", "USA Attained Reviews"]]
        else:
            usa_attained_monthly = pd.DataFrame(columns=["Month", "USA Attained Reviews"])

        if not uk_reviews_per_pm.empty:
            uk_attained_monthly = (
                uk_reviews_per_pm.groupby(uk_reviews_per_pm["Trustpilot Review Date"].dt.to_period("M"))
                .size()
                .reset_index(name="UK Attained Reviews")
            )
            uk_attained_monthly["Month"] = uk_attained_monthly["Trustpilot Review Date"].dt.strftime("%B %Y")
            uk_attained_monthly = uk_attained_monthly[["Month", "UK Attained Reviews"]]
        else:
            uk_attained_monthly = pd.DataFrame(columns=["Month", "UK Attained Reviews"])
        attained_reviews_per_month = pd.merge(
            usa_attained_monthly,
            uk_attained_monthly,
            on="Month",
            how="outer"
        ).fillna(0)

        attained_reviews_per_month["Total Attained Reviews"] = (
                attained_reviews_per_month["USA Attained Reviews"] + attained_reviews_per_month["UK Attained Reviews"]
        )

        attained_reviews_per_month["Month_Num"] = pd.to_datetime(attained_reviews_per_month["Month"], format="%B %Y")
        attained_reviews_per_month = attained_reviews_per_month.sort_values(by="Total Attained Reviews",
                                                                            ascending=False)
        attained_reviews_per_month.index = range(1, len(attained_reviews_per_month) + 1)
        attained_reviews_per_month = attained_reviews_per_month.drop(columns="Month_Num")

        attained_details["Trustpilot Review Date"] = pd.to_datetime(
            attained_details["Trustpilot Review Date"], errors="coerce"
        ).dt.strftime("%d-%B-%Y")

    else:
        attained_reviews_per_pm = pd.DataFrame(columns=["Project Manager", "Attained Reviews"])
        attained_details = pd.DataFrame(
            columns=["Project Manager", "Name", "Brand", "Trustpilot Review Date", "Trustpilot Review Links", "Status"])
        attained_reviews_per_month = pd.DataFrame(columns=["Month", "Total Attained Reviews"])

    usa_negative_per_pm = [load_reviews_year(sheet_usa, year, pm, "Negative") for pm in pm_list_usa]
    usa_negative_per_pm = safe_concat([df for df in usa_negative_per_pm if not df.empty])

    uk_negative_per_pm = [load_reviews_year(sheet_uk, year, pm, "Negative") for pm in pm_list_uk]
    uk_negative_per_pm = safe_concat([df for df in uk_negative_per_pm if not df.empty])

    combined_negative_data = safe_concat([usa_negative_per_pm, uk_negative_per_pm])

    if not usa_negative_per_pm.empty:
        usa_negative_pm = (
            usa_negative_per_pm
            .groupby("Project Manager")["Trustpilot Review"]
            .count()
            .reset_index()
        )
        usa_negative_pm.columns = ["Project Manager", "Negative Reviews"]
        usa_negative_pm.index = range(1, len(usa_negative_pm) + 1)
        usa_total_negative = usa_negative_pm["Negative Reviews"].sum()
    else:
        usa_negative_pm = pd.DataFrame(columns=["Project Manager", "Negative Reviews"])
        usa_total_negative = 0

    if not uk_negative_per_pm.empty:
        uk_negative_pm = (
            uk_negative_per_pm
            .groupby("Project Manager")["Trustpilot Review"]
            .count()
            .reset_index()
        )
        uk_negative_pm.columns = ["Project Manager", "Negative Reviews"]
        uk_negative_pm.index = range(1, len(uk_negative_pm) + 1)
        uk_total_negative = uk_negative_pm["Negative Reviews"].sum()
    else:
        uk_negative_pm = pd.DataFrame(columns=["Project Manager", "Negative Reviews"])
        uk_total_negative = 0

    if not combined_negative_data.empty:

        negative_reviews_per_pm = (
            combined_negative_data
            .groupby("Project Manager")["Trustpilot Review"]
            .count()
            .reset_index()
        )
        negative_reviews_per_pm.columns = ["Project Manager", "Negative Reviews"]
        negative_reviews_per_pm = negative_reviews_per_pm.sort_values(by="Negative Reviews", ascending=False)
        negative_reviews_per_pm.index = range(1, len(negative_reviews_per_pm) + 1)

        negative_details_df = combined_negative_data.sort_values(by="Project Manager", ascending=True)
        negative_details_df["Trustpilot Review Date"] = pd.to_datetime(
            negative_details_df["Trustpilot Review Date"], errors="coerce"
        ).dt.strftime("%d-%B-%Y")

        negative_details = negative_details_df[
            ["Project Manager", "Name", "Brand", "Trustpilot Review Date", "Trustpilot Review Links", "Status"]
        ]
        negative_details.index = range(1, len(negative_details) + 1)

        negative_details["Trustpilot Review Date"] = pd.to_datetime(
            negative_details["Trustpilot Review Date"], errors="coerce"
        )

        if not usa_negative_per_pm.empty:
            usa_negative_monthly = (
                usa_negative_per_pm.groupby(usa_negative_per_pm["Trustpilot Review Date"].dt.to_period("M"))
                .size()
                .reset_index(name="USA Negative Reviews")
            )
            usa_negative_monthly["Month"] = usa_negative_monthly["Trustpilot Review Date"].dt.strftime("%B %Y")
            usa_negative_monthly = usa_negative_monthly[["Month", "USA Negative Reviews"]]
        else:
            usa_negative_monthly = pd.DataFrame(columns=["Month", "USA Negative Reviews"])

        # UK monthly negative reviews
        if not uk_negative_per_pm.empty:
            uk_negative_monthly = (
                uk_negative_per_pm.groupby(uk_negative_per_pm["Trustpilot Review Date"].dt.to_period("M"))
                .size()
                .reset_index(name="UK Negative Reviews")
            )
            uk_negative_monthly["Month"] = uk_negative_monthly["Trustpilot Review Date"].dt.strftime("%B %Y")
            uk_negative_monthly = uk_negative_monthly[["Month", "UK Negative Reviews"]]
        else:
            uk_negative_monthly = pd.DataFrame(columns=["Month", "UK Negative Reviews"])

        # Merge USA and UK negative trends
        negative_reviews_per_month = pd.merge(
            usa_negative_monthly,
            uk_negative_monthly,
            on="Month",
            how="outer"
        ).fillna(0)

        negative_reviews_per_month["Total Negative Reviews"] = (
                negative_reviews_per_month["USA Negative Reviews"] + negative_reviews_per_month["UK Negative Reviews"]
        )

        # Sort by month
        negative_reviews_per_month["Month_Num"] = pd.to_datetime(negative_reviews_per_month["Month"], format="%B %Y")
        negative_reviews_per_month = negative_reviews_per_month.sort_values(by="Total Negative Reviews",
                                                                            ascending=False)
        negative_reviews_per_month.index = range(1, len(negative_reviews_per_month) + 1)
        negative_reviews_per_month = negative_reviews_per_month.drop(columns="Month_Num")
        negative_details["Trustpilot Review Date"] = pd.to_datetime(
            negative_details["Trustpilot Review Date"], errors="coerce"
        ).dt.strftime("%d-%B-%Y")

    else:
        negative_reviews_per_pm = pd.DataFrame(columns=["Project Manager", "Negative Reviews"])
        negative_details = pd.DataFrame(
            columns=["Project Manager", "Name", "Brand", "Trustpilot Review Date", "Trustpilot Review Links", "Status"])
        negative_reviews_per_month = pd.DataFrame(columns=["Month", "Total Negative Reviews"])

    usa_review = {
        "Attained": usa_total_attained,
        "Sent": usa_review_sent,
        "Pending": usa_review_pending,
        "Negative": usa_total_negative
    }

    uk_review = {
        "Attained": uk_total_attained,
        "Sent": uk_review_sent,
        "Pending": uk_review_pending,
        "Negative": uk_total_negative
    }

    printing_data, monthly_printing = printing_data_year(year)
    Total_copies = printing_data["No of Copies"].sum() if "No of Copies" in printing_data.columns else 0
    Total_cost = printing_data["Order Cost"].sum() if "Order Cost" in printing_data.columns else 0
    Highest_cost = printing_data["Order Cost"].max() if "Order Cost" in printing_data.columns else 0
    Highest_copies = printing_data["No of Copies"].max() if "No of Copies" in printing_data.columns else 0
    Lowest_cost = printing_data["Order Cost"].min() if "Order Cost" in printing_data.columns else 0
    Lowest_copies = printing_data["No of Copies"].min() if "No of Copies" in printing_data.columns else 0

    Average = Total_cost / Total_copies if Total_copies > 0 else 0
    if all(col in printing_data.columns for col in ["Order Cost", "No of Copies"]):
        printing_data['Cost_Per_Copy'] = printing_data['Order Cost'] / printing_data['No of Copies']

    copyright_data, result_count, result_count_no = copyright_year(year)
    Total_copyrights = len(copyright_data)
    country = copyright_data["Country"].value_counts()
    usa = country.get("USA", 0)
    canada = country.get("Canada", 0)
    uk = country.get("UK", 0)
    Total_cost_copyright = (usa * 65) + (canada * 63) + (uk * 35)

    a_plus, a_plus_count = get_A_plus_year(year)

    usa_brands = {'BookMarketeers': bookmarketeers, 'Writers Clique': writers_clique, 'KDP': kdp,
                  'Aurora Writers': aurora_writers}
    uk_brands = {'Authors Solution': authors_solution, 'Book Publication': book_publication}

    usa_platforms = {'Amazon': usa_amazon, 'Barnes & Noble': usa_bn, 'Ingram Spark': usa_ingram,"Draft2Digital":usa_d2d, "FAV": usa_fav, "ACX": usa_acx}
    uk_platforms = {'Amazon': uk_amazon, 'Barnes & Noble': uk_bn, 'Ingram Spark': uk_ingram,"Draft2Digital":uk_d2d, "FAV": uk_fav,
                    "Kobo": uk_kobo, "ACX": uk_acx}

    printing_stats = {
        'Total_copies': Total_copies,
        'Total_cost': Total_cost,
        'Highest_cost': Highest_cost,
        'Lowest_cost': Lowest_cost,
        'Highest_copies': Highest_copies,
        'Lowest_copies': Lowest_copies,
        'Average': Average
    }

    copyright_stats = {
        'Total_copyrights': Total_copyrights,
        'Total_cost_copyright': Total_cost_copyright,
        'result_count': result_count,
        'result_count_no': result_count_no,
        'usa_copyrights': usa,
        'canada_copyrights': canada,
        'uk': uk
    }

    return usa_review, uk_review, usa_brands, uk_brands, usa_platforms, uk_platforms, printing_stats, monthly_printing, copyright_stats, a_plus_count, total_unique_clients, combined, attained_reviews_per_pm, attained_details, attained_reviews_per_month, pending_sent_details, negative_reviews_per_pm, negative_details, negative_reviews_per_month, combined_monthly, Issues_usa, Issues_uk


def generate_year_summary_multiple(start_year: int, end_year: int):
    uk_clean = clean_data_reviews(sheet_uk)
    usa_clean = clean_data_reviews(sheet_usa)

    usa_clean = usa_clean[
        (usa_clean["Publishing Date"].dt.year >= start_year) &
        (usa_clean["Publishing Date"].dt.year <= end_year)

    ]
    uk_clean = uk_clean[
        (uk_clean["Publishing Date"].dt.year >= start_year) &
        (uk_clean["Publishing Date"].dt.year <= end_year)
    ]

    usa_clean_platforms = usa_clean[
        (usa_clean["Publishing Date"].dt.year >= start_year) &
        (usa_clean["Publishing Date"].dt.year <= end_year)
        ]
    uk_clean_platforms = uk_clean[
        (uk_clean["Publishing Date"].dt.year >= start_year) &
        (uk_clean["Publishing Date"].dt.year <= end_year)
    ]

    if usa_clean.empty:
        print("No values found in USA sheet.")
    if uk_clean.empty:
        print("No values found in UK sheet.")
        return
    if usa_clean.empty and uk_clean.empty:
        return

    usa_clean = usa_clean.drop_duplicates(subset=["Name"], keep="first")
    uk_clean = uk_clean.drop_duplicates(subset=["Name"], keep="first")
    Issues_usa = usa_clean["Issues"].value_counts()
    Issues_uk = uk_clean["Issues"].value_counts()
    total_usa = usa_clean["Name"].nunique()
    total_uk = uk_clean["Name"].nunique()
    total_unique_clients = total_usa + total_uk

    combined = pd.concat([usa_clean[["Name", "Brand", "Project Manager", "Email"]],
                          uk_clean[["Name", "Brand", "Project Manager", "Email"]]])
    combined.index = range(1, len(combined) + 1)

    brands = usa_clean["Brand"].value_counts()
    writers_clique = brands.get("Writers Clique", 0)
    bookmarketeers = brands.get("BookMarketeers", 0)
    aurora_writers = brands.get("Aurora Writers", 0)
    kdp = brands.get("KDP", 0)

    uk_brand = uk_clean["Brand"].value_counts()
    authors_solution = uk_brand.get("Authors Solution", 0)
    book_publication = uk_brand.get("Book Publication", 0)

    usa_platforms = usa_clean_platforms["Platform"].value_counts()
    usa_amazon = usa_platforms.get("Amazon", 0)
    usa_bn = usa_platforms.get("Barnes & Noble", 0)
    usa_ingram = usa_platforms.get("Ingram Spark", 0)
    usa_d2d = usa_platforms.get("Draft2Digital", 0)
    usa_fav = usa_platforms.get("FAV", 0)
    usa_acx = usa_platforms.get("ACX", 0)

    uk_platforms = uk_clean_platforms["Platform"].value_counts()
    uk_amazon = uk_platforms.get("Amazon", 0)
    uk_bn = uk_platforms.get("Barnes & Noble", 0)
    uk_ingram = uk_platforms.get("Ingram Spark", 0)
    uk_d2d = uk_platforms.get("Draft2Digital", 0)
    uk_fav = uk_platforms.get("FAV", 0)
    uk_kobo = uk_platforms.get("Kobo", 0)
    uk_acx = uk_platforms.get("ACX", 0)

    allowed_brands = ["BookMarketeers", "Writers Clique", "Aurora Writers", "Authors Solution", "Book Publication"]

    if "Trustpilot Review" in usa_clean.columns and "Brand" in usa_clean.columns:
        usa_filtered = usa_clean[usa_clean["Brand"].isin(allowed_brands)]
        usa_review_sent = usa_filtered["Trustpilot Review"].value_counts().get("Sent", 0)
        usa_review_pending = usa_filtered["Trustpilot Review"].value_counts().get("Pending", 0)
        usa_review_na = usa_filtered["Trustpilot Review"].value_counts().get("Negative", 0)
    else:
        usa_review_sent = usa_review_pending = usa_review_na = 0

    if "Trustpilot Review" in uk_clean.columns and "Brand" in uk_clean.columns:
        uk_filtered = uk_clean[uk_clean["Brand"].isin(allowed_brands)]
        uk_review_sent = uk_filtered["Trustpilot Review"].value_counts().get("Sent", 0)
        uk_review_pending = uk_filtered["Trustpilot Review"].value_counts().get("Pending", 0)
        uk_review_na = uk_filtered["Trustpilot Review"].value_counts().get("Negative", 0)
    else:
        uk_review_sent = uk_review_pending = uk_review_na = 0

    combined_pending_sent = pd.concat([usa_clean, uk_clean], ignore_index=True)
    pending_sent_details = combined_pending_sent[
        ((combined_pending_sent["Trustpilot Review"] == "Sent") |
         (combined_pending_sent["Trustpilot Review"] == "Pending")) &
        (combined_pending_sent["Brand"].isin(allowed_brands))
        ]
    pending_sent_details = pending_sent_details[["Name", "Brand", "Project Manager", "Trustpilot Review", "Status"]]
    pending_sent_details.index = range(1, len(pending_sent_details) + 1)

    pm_list_usa = list(set((usa_clean["Project Manager"].dropna().unique().tolist() + ["Unknown"])))
    pm_list_uk = list(set((uk_clean["Project Manager"].dropna().unique().tolist() + ["Unknown"])))

    usa_reviews_per_pm = safe_concat([load_reviews_year_multiple(sheet_usa, start_year, end_year, pm, "Attained") for pm in pm_list_usa])
    uk_reviews_per_pm = safe_concat([load_reviews_year_multiple(sheet_uk, start_year, end_year, pm, "Attained") for pm in pm_list_uk])
    combined_data = safe_concat([usa_reviews_per_pm, uk_reviews_per_pm])

    usa_monthly = (
        usa_clean.groupby(usa_clean["Publishing Date"].dt.to_period("M"))
        .size()
        .reset_index(name="USA Published")
    )
    usa_monthly["Month"] = usa_monthly["Publishing Date"].dt.strftime("%B %Y")
    usa_monthly = usa_monthly[["Month", "USA Published"]]

    uk_monthly = (
        uk_clean.groupby(uk_clean["Publishing Date"].dt.to_period("M"))
        .size()
        .reset_index(name="UK Published")
    )
    uk_monthly["Month"] = uk_monthly["Publishing Date"].dt.strftime("%B %Y")
    uk_monthly = uk_monthly[["Month", "UK Published"]]

    combined_monthly = pd.merge(
        usa_monthly,
        uk_monthly,
        on="Month",
        how="outer"
    ).fillna(0)

    combined_monthly["Total Published"] = combined_monthly["USA Published"] + combined_monthly["UK Published"]

    combined_monthly["Month_Num"] = pd.to_datetime(combined_monthly["Month"], format="%B %Y")
    combined_monthly = combined_monthly.sort_values("Total Published", ascending=False).drop(columns="Month_Num")

    combined_monthly.index = range(1, len(combined_monthly) + 1)

    if not usa_reviews_per_pm.empty:
        usa_attained_pm = (
            usa_reviews_per_pm
            .groupby("Project Manager")["Trustpilot Review"]
            .count()
            .reset_index()
        )
        usa_attained_pm.columns = ["Project Manager", "Attained Reviews"]
        usa_attained_pm.index = range(1, len(usa_attained_pm) + 1)
        usa_total_attained = usa_attained_pm["Attained Reviews"].sum()
    else:
        usa_attained_pm = pd.DataFrame(columns=["Project Manager", "Attained Reviews"])
        usa_total_attained = 0

    if not uk_reviews_per_pm.empty:
        uk_attained_pm = (
            uk_reviews_per_pm
            .groupby("Project Manager")["Trustpilot Review"]
            .count()
            .reset_index()
        )
        uk_attained_pm.columns = ["Project Manager", "Attained Reviews"]
        uk_attained_pm.index = range(1, len(uk_attained_pm) + 1)
        uk_total_attained = uk_attained_pm["Attained Reviews"].sum()
    else:
        uk_attained_pm = pd.DataFrame(columns=["Project Manager", "Attained Reviews"])
        uk_total_attained = 0

    if not combined_data.empty:
        attained_reviews_per_pm = (
            combined_data
            .groupby("Project Manager")["Trustpilot Review"]
            .count()
            .reset_index()
        )
        attained_reviews_per_pm.columns = ["Project Manager", "Attained Reviews"]
        attained_reviews_per_pm = attained_reviews_per_pm.sort_values(by="Attained Reviews", ascending=False)
        attained_reviews_per_pm.index = range(1, len(attained_reviews_per_pm) + 1)

        review_details_df = combined_data.sort_values(by="Project Manager", ascending=True)
        review_details_df["Trustpilot Review Date"] = pd.to_datetime(
            review_details_df["Trustpilot Review Date"], errors="coerce"
        ).dt.strftime("%d-%B-%Y")

        attained_details = review_details_df[
            ["Project Manager", "Name", "Brand", "Trustpilot Review Date", "Trustpilot Review Links", "Status"]
        ]
        attained_details.index = range(1, len(attained_details) + 1)

        attained_details["Trustpilot Review Date"] = pd.to_datetime(
            attained_details["Trustpilot Review Date"], errors="coerce"
        )

        if not usa_reviews_per_pm.empty:
            usa_attained_monthly = (
                usa_reviews_per_pm.groupby(usa_reviews_per_pm["Trustpilot Review Date"].dt.to_period("M"))
                .size()
                .reset_index(name="USA Attained Reviews")
            )
            usa_attained_monthly["Month"] = usa_attained_monthly["Trustpilot Review Date"].dt.strftime("%B %Y")
            usa_attained_monthly = usa_attained_monthly[["Month", "USA Attained Reviews"]]
        else:
            usa_attained_monthly = pd.DataFrame(columns=["Month", "USA Attained Reviews"])

        if not uk_reviews_per_pm.empty:
            uk_attained_monthly = (
                uk_reviews_per_pm.groupby(uk_reviews_per_pm["Trustpilot Review Date"].dt.to_period("M"))
                .size()
                .reset_index(name="UK Attained Reviews")
            )
            uk_attained_monthly["Month"] = uk_attained_monthly["Trustpilot Review Date"].dt.strftime("%B %Y")
            uk_attained_monthly = uk_attained_monthly[["Month", "UK Attained Reviews"]]
        else:
            uk_attained_monthly = pd.DataFrame(columns=["Month", "UK Attained Reviews"])
        attained_reviews_per_month = pd.merge(
            usa_attained_monthly,
            uk_attained_monthly,
            on="Month",
            how="outer"
        ).fillna(0)

        attained_reviews_per_month["Total Attained Reviews"] = (
                attained_reviews_per_month["USA Attained Reviews"] + attained_reviews_per_month["UK Attained Reviews"]
        )

        attained_reviews_per_month["Month_Num"] = pd.to_datetime(attained_reviews_per_month["Month"], format="%B %Y")
        attained_reviews_per_month = attained_reviews_per_month.sort_values(by="Total Attained Reviews",
                                                                            ascending=False)
        attained_reviews_per_month.index = range(1, len(attained_reviews_per_month) + 1)
        attained_reviews_per_month = attained_reviews_per_month.drop(columns="Month_Num")

        attained_details["Trustpilot Review Date"] = pd.to_datetime(
            attained_details["Trustpilot Review Date"], errors="coerce"
        ).dt.strftime("%d-%B-%Y")

    else:
        attained_reviews_per_pm = pd.DataFrame(columns=["Project Manager", "Attained Reviews"])
        attained_details = pd.DataFrame(
            columns=["Project Manager", "Name", "Brand", "Trustpilot Review Date", "Trustpilot Review Links", "Status"])
        attained_reviews_per_month = pd.DataFrame(columns=["Month", "Total Attained Reviews"])

    usa_negative_per_pm = [load_reviews_year_multiple(sheet_usa, start_year, end_year,  pm, "Negative") for pm in pm_list_usa]
    usa_negative_per_pm = safe_concat([df for df in usa_negative_per_pm if not df.empty])

    uk_negative_per_pm = [load_reviews_year_multiple(sheet_uk, start_year, end_year, pm, "Negative") for pm in pm_list_uk]
    uk_negative_per_pm = safe_concat([df for df in uk_negative_per_pm if not df.empty])

    combined_negative_data = safe_concat([usa_negative_per_pm, uk_negative_per_pm])

    if not usa_negative_per_pm.empty:
        usa_negative_pm = (
            usa_negative_per_pm
            .groupby("Project Manager")["Trustpilot Review"]
            .count()
            .reset_index()
        )
        usa_negative_pm.columns = ["Project Manager", "Negative Reviews"]
        usa_negative_pm.index = range(1, len(usa_negative_pm) + 1)
        usa_total_negative = usa_negative_pm["Negative Reviews"].sum()
    else:
        usa_negative_pm = pd.DataFrame(columns=["Project Manager", "Negative Reviews"])
        usa_total_negative = 0

    if not uk_negative_per_pm.empty:
        uk_negative_pm = (
            uk_negative_per_pm
            .groupby("Project Manager")["Trustpilot Review"]
            .count()
            .reset_index()
        )
        uk_negative_pm.columns = ["Project Manager", "Negative Reviews"]
        uk_negative_pm.index = range(1, len(uk_negative_pm) + 1)
        uk_total_negative = uk_negative_pm["Negative Reviews"].sum()
    else:
        uk_negative_pm = pd.DataFrame(columns=["Project Manager", "Negative Reviews"])
        uk_total_negative = 0

    if not combined_negative_data.empty:

        negative_reviews_per_pm = (
            combined_negative_data
            .groupby("Project Manager")["Trustpilot Review"]
            .count()
            .reset_index()
        )
        negative_reviews_per_pm.columns = ["Project Manager", "Negative Reviews"]
        negative_reviews_per_pm = negative_reviews_per_pm.sort_values(by="Negative Reviews", ascending=False)
        negative_reviews_per_pm.index = range(1, len(negative_reviews_per_pm) + 1)

        negative_details_df = combined_negative_data.sort_values(by="Project Manager", ascending=True)
        negative_details_df["Trustpilot Review Date"] = pd.to_datetime(
            negative_details_df["Trustpilot Review Date"], errors="coerce"
        ).dt.strftime("%d-%B-%Y")

        negative_details = negative_details_df[
            ["Project Manager", "Name", "Brand", "Trustpilot Review Date", "Trustpilot Review Links", "Status"]
        ]
        negative_details.index = range(1, len(negative_details) + 1)

        negative_details["Trustpilot Review Date"] = pd.to_datetime(
            negative_details["Trustpilot Review Date"], errors="coerce"
        )

        if not usa_negative_per_pm.empty:
            usa_negative_monthly = (
                usa_negative_per_pm.groupby(usa_negative_per_pm["Trustpilot Review Date"].dt.to_period("M"))
                .size()
                .reset_index(name="USA Negative Reviews")
            )
            usa_negative_monthly["Month"] = usa_negative_monthly["Trustpilot Review Date"].dt.strftime("%B %Y")
            usa_negative_monthly = usa_negative_monthly[["Month", "USA Negative Reviews"]]
        else:
            usa_negative_monthly = pd.DataFrame(columns=["Month", "USA Negative Reviews"])

        # UK monthly negative reviews
        if not uk_negative_per_pm.empty:
            uk_negative_monthly = (
                uk_negative_per_pm.groupby(uk_negative_per_pm["Trustpilot Review Date"].dt.to_period("M"))
                .size()
                .reset_index(name="UK Negative Reviews")
            )
            uk_negative_monthly["Month"] = uk_negative_monthly["Trustpilot Review Date"].dt.strftime("%B %Y")
            uk_negative_monthly = uk_negative_monthly[["Month", "UK Negative Reviews"]]
        else:
            uk_negative_monthly = pd.DataFrame(columns=["Month", "UK Negative Reviews"])

        # Merge USA and UK negative trends
        negative_reviews_per_month = pd.merge(
            usa_negative_monthly,
            uk_negative_monthly,
            on="Month",
            how="outer"
        ).fillna(0)

        negative_reviews_per_month["Total Negative Reviews"] = (
                negative_reviews_per_month["USA Negative Reviews"] + negative_reviews_per_month["UK Negative Reviews"]
        )

        # Sort by month
        negative_reviews_per_month["Month_Num"] = pd.to_datetime(negative_reviews_per_month["Month"], format="%B %Y")
        negative_reviews_per_month = negative_reviews_per_month.sort_values(by="Total Negative Reviews",
                                                                            ascending=False)
        negative_reviews_per_month.index = range(1, len(negative_reviews_per_month) + 1)
        negative_reviews_per_month = negative_reviews_per_month.drop(columns="Month_Num")
        negative_details["Trustpilot Review Date"] = pd.to_datetime(
            negative_details["Trustpilot Review Date"], errors="coerce"
        ).dt.strftime("%d-%B-%Y")

    else:
        negative_reviews_per_pm = pd.DataFrame(columns=["Project Manager", "Negative Reviews"])
        negative_details = pd.DataFrame(
            columns=["Project Manager", "Name", "Brand", "Trustpilot Review Date", "Trustpilot Review Links", "Status"])
        negative_reviews_per_month = pd.DataFrame(columns=["Month", "Total Negative Reviews"])

    usa_review = {
        "Attained": usa_total_attained,
        "Sent": usa_review_sent,
        "Pending": usa_review_pending,
        "Negative": usa_total_negative
    }

    uk_review = {
        "Attained": uk_total_attained,
        "Sent": uk_review_sent,
        "Pending": uk_review_pending,
        "Negative": uk_total_negative
    }

    printing_data, monthly_printing = printing_data_year_multiple(start_year, end_year)
    Total_copies = printing_data["No of Copies"].sum() if "No of Copies" in printing_data.columns else 0
    Total_cost = printing_data["Order Cost"].sum() if "Order Cost" in printing_data.columns else 0
    Highest_cost = printing_data["Order Cost"].max() if "Order Cost" in printing_data.columns else 0
    Highest_copies = printing_data["No of Copies"].max() if "No of Copies" in printing_data.columns else 0
    Lowest_cost = printing_data["Order Cost"].min() if "Order Cost" in printing_data.columns else 0
    Lowest_copies = printing_data["No of Copies"].min() if "No of Copies" in printing_data.columns else 0

    Average = Total_cost / Total_copies if Total_copies > 0 else 0
    if all(col in printing_data.columns for col in ["Order Cost", "No of Copies"]):
        printing_data['Cost_Per_Copy'] = printing_data['Order Cost'] / printing_data['No of Copies']

    copyright_data, result_count, result_count_no = copyright_year_multiple(start_year, end_year)
    Total_copyrights = len(copyright_data)
    country = copyright_data["Country"].value_counts()
    usa = country.get("USA", 0)
    canada = country.get("Canada", 0)
    uk = country.get("UK", 0)
    Total_cost_copyright = (usa * 65) + (canada * 63) + (uk * 35)

    a_plus, a_plus_count = get_A_plus_year_multiple(start_year, end_year)

    usa_brands = {'BookMarketeers': bookmarketeers, 'Writers Clique': writers_clique, 'KDP': kdp,
                  'Aurora Writers': aurora_writers}
    uk_brands = {'Authors Solution': authors_solution, 'Book Publication': book_publication}

    usa_platforms = {'Amazon': usa_amazon, 'Barnes & Noble': usa_bn, 'Ingram Spark': usa_ingram,"Draft2Digital":usa_d2d, "FAV": usa_fav, "ACX": usa_acx}
    uk_platforms = {'Amazon': uk_amazon, 'Barnes & Noble': uk_bn, 'Ingram Spark': uk_ingram,"Draft2Digital":uk_d2d, "FAV": uk_fav,
                    "Kobo": uk_kobo, "ACX": uk_acx}

    printing_stats = {
        'Total_copies': Total_copies,
        'Total_cost': Total_cost,
        'Highest_cost': Highest_cost,
        'Lowest_cost': Lowest_cost,
        'Highest_copies': Highest_copies,
        'Lowest_copies': Lowest_copies,
        'Average': Average
    }

    copyright_stats = {
        'Total_copyrights': Total_copyrights,
        'Total_cost_copyright': Total_cost_copyright,
        'result_count': result_count,
        'result_count_no': result_count_no,
        'usa_copyrights': usa,
        'canada_copyrights': canada,
        'uk': uk
    }

    return usa_review, uk_review, usa_brands, uk_brands, usa_platforms, uk_platforms, printing_stats, monthly_printing, copyright_stats, a_plus_count, total_unique_clients, combined, attained_reviews_per_pm, attained_details, attained_reviews_per_month, pending_sent_details, negative_reviews_per_pm, negative_details, negative_reviews_per_month, combined_monthly, Issues_usa, Issues_uk


def logging_function() -> None:
    """Creates a console and file logging handler that logs messages
        Returns:
              None: Returns nothing but calls the logging function
    """
    logging.basicConfig(level=logging.INFO, format='%(funcName)s --> %(message)s : %(asctime)s - %(levelname)s',
                        datefmt="%d-%m-%Y %I:%M:%S %p")

    file_handler = logging.FileHandler('Reviews.log')
    file_handler.setLevel(level=logging.WARNING)
    formatter = logging.Formatter('%(funcName)s --> %(message)s : %(asctime)s - %(levelname)s',
                                  "%d-%m-%Y %I:%M:%S %p")
    file_handler.setFormatter(formatter)

    logger = logging.getLogger('')
    logger.addHandler(file_handler)


def generate_summary_report_pdf(
    usa_review_data,
    uk_review_data,
    usa_brands,
    uk_brands,
    usa_platforms,
    uk_platforms,
    printing_stats,
    copyright_stats,
    a_plus,
    selected_month=None,
    start_year=None,
    end_year=None,
    filename=None
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
        buffer,
        pagesize=A4,
        rightMargin=72,
        leftMargin=72,
        topMargin=72,
        bottomMargin=18
    )

    styles = getSampleStyleSheet()

    title_style = ParagraphStyle(
        'Title',
        parent=styles['Heading1'],
        fontSize=24,
        spaceAfter=30,
        alignment=TA_CENTER,
        textColor=colors.darkblue
    )

    section_style = ParagraphStyle(
        'Section',
        parent=styles['Heading2'],
        fontSize=16,
        spaceBefore=20,
        spaceAfter=12,
        textColor=colors.darkgreen
    )

    subsection_style = ParagraphStyle(
        'SubSection',
        parent=styles['Heading3'],
        fontSize=12,
        spaceBefore=12,
        spaceAfter=8,
        textColor=colors.darkblue
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
        copyright_stats['result_count'] /
        copyright_stats['Total_copyrights'] * 100
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
    story.append(
        Paragraph(
            f"Generated on {datetime.now().strftime('%B %d, %Y %I:%M %p')}",
            styles['Normal']
        )
    )

    doc.build(story)

    pdf_data = buffer.getvalue()
    buffer.close()

    return pdf_data, filename





def sales(month: int, year: int) -> pd.DataFrame:
    data = get_sheet_data(sheet_sales)

    if data.empty:
        return pd.DataFrame()

    columns = list(data.columns)
    if "Payment" in columns:
        end_col_index = columns.index("Payment")
        data = data.iloc[:, :end_col_index + 1]
        data["Payment Date"] = pd.to_datetime(data["Payment Date"], errors="coerce")

    if "Payment Date" in data.columns:
        data = data[(data["Payment Date"].dt.month == month) & (data["Payment Date"].dt.year == year)]

    if "Payment" in data.columns:
        data["Payment"] = data["Payment"].astype(str)
        data["Payment"] = pd.to_numeric(
            data["Payment"].str.replace("$", "", regex=False).str.replace(",", "", regex=False), errors="coerce")
    data["Payment Date"] = data["Payment Date"].dt.strftime("%d-%B-%Y")

    data.index = range(1, len(data) + 1)

    return data


def sales_year(year: int) -> pd.DataFrame:
    data = get_sheet_data(sheet_sales)

    if data.empty:
        return pd.DataFrame()

    columns = list(data.columns)
    if "Payment" in columns:
        end_col_index = columns.index("Payment")
        data = data.iloc[:, :end_col_index + 1]
        data["Payment Date"] = pd.to_datetime(data["Payment Date"], errors="coerce")

    if "Payment Date" in data.columns:
        data = data[data["Payment Date"].dt.year == year]

    if "Payment" in data.columns:
        data["Payment"] = data["Payment"].astype(str)
        data["Payment"] = pd.to_numeric(
            data["Payment"].str.replace("$", "", regex=False).str.replace(",", "", regex=False), errors="coerce")

    data["Payment Date"] = data["Payment Date"].dt.strftime("%d-%B-%Y")
    data.index = range(1, len(data) + 1)

    return data


def main() -> None:
    with st.container():
        st.title("📊 Blink Digitally Publishing Dashboard")
        if st.button("🔃 Fetch Latest"):
            st.cache_data.clear()
            st.success("Fetched new data")
        action = st.selectbox("What would you like to do?",
                              ["View Data", "Printing", "Copyright", "Generate Similarity",
                               "Summary",
                               "Year Summary", "Custom Summary", "Reviews", "Sales"],
                              index=None,
                              placeholder="Select Action")

        country = None
        selected_month = None
        selected_month_number = None
        status = None
        choice = None
        number = None
        if action in ["View Data"]:
            choice = st.selectbox("Select Data To View", ["USA", "UK"], index=None,
                                  placeholder="Select Data to View")

        if action in ["View Data"]:
            selected_month = st.selectbox(
                "Select Month",
                month_list,
                index=current_month - 1,
                placeholder="Select Month"
            )
            selected_month_number = month_list.index(selected_month) + 1 if selected_month else None
        if action in ["Year Summary", "View Data", "Reviews"]:
            number = st.number_input("Enter Year", min_value=int(get_min_year()), max_value=current_year,
                                     value=current_year, step=1)

        if action == "View Data" and choice and selected_month and number:
            tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs(["Monthly", "Yearly","Start to Year", "Filter", "Search", "By Brand"])

            sheet_name = {
                "UK": sheet_uk,
                "USA": sheet_usa,
            }.get(choice)

            with tab1:
                st.subheader(f"📂 Viewing Data for {choice} - {selected_month} {number}")

                if sheet_name:
                    data = load_data(sheet_name, selected_month_number, number)
                    if not data.empty:
                        data_rm_dupes = data.copy()
                        if "Name" in data_rm_dupes.columns:
                            data_rm_dupes = data_rm_dupes.drop_duplicates(subset=["Name"], keep="first")
                        review_data = load_reviews(sheet_name, number, selected_month_number)

                        if not review_data.empty:
                            attained_reviews_per_pm = review_data[
                                review_data["Trustpilot Review"] == "Attained"
                                ].groupby("Project Manager")["Trustpilot Review"].count().reset_index()

                            review_details_df = review_data.sort_values(by="Project Manager", ascending=True)
                            review_details_df["Trustpilot Review Date"] = pd.to_datetime(
                                review_details_df["Trustpilot Review Date"], errors="coerce"
                            ).dt.strftime("%d-%B-%Y")
                        else:
                            attained_reviews_per_pm = pd.DataFrame()
                        if not attained_reviews_per_pm.empty:
                            attained_reviews_per_pm.columns = ["Project Manager", "Attained Reviews"]
                            attained_reviews_per_pm = attained_reviews_per_pm.sort_values(
                                by="Attained Reviews", ascending=False
                            )
                            attained_reviews_per_pm.index = range(1, len(attained_reviews_per_pm) + 1)

                            attained_details = review_details_df[
                                review_details_df["Trustpilot Review"] == "Attained"
                                ][["Project Manager", "Name", "Brand", "Trustpilot Review Date",
                                   "Trustpilot Review Links",
                                   "Status"]].copy()
                            attained_details.index = range(1, len(attained_details) + 1)
                        else:
                            attained_reviews_per_pm = pd.DataFrame(columns=["Project Manager", "Attained Reviews"])
                            attained_details = pd.DataFrame(columns=[
                                "Project Manager", "Name", "Brand", "Trustpilot Review Date", "Trustpilot Review Links",
                                "Status"
                            ])

                        if not review_data.empty:
                            negative_reviews_per_pm = review_data[
                                review_data["Trustpilot Review"] == "Negative"
                                ].groupby("Project Manager")["Trustpilot Review"].count().reset_index()
                        else:
                            negative_reviews_per_pm = pd.DataFrame()
                        if not negative_reviews_per_pm.empty:
                            negative_reviews_per_pm.columns = ["Project Manager", "Negative Reviews"]
                            negative_reviews_per_pm = negative_reviews_per_pm.sort_values(
                                by="Negative Reviews", ascending=False
                            )
                            negative_reviews_per_pm.index = range(1, len(negative_reviews_per_pm) + 1)

                            negative_details = review_details_df[
                                review_details_df["Trustpilot Review"] == "Negative"
                                ][["Project Manager", "Name", "Brand", "Trustpilot Review Date",
                                   "Trustpilot Review Links",
                                   "Status"]].copy()
                            negative_details.index = range(1, len(negative_details) + 1)
                        else:
                            negative_reviews_per_pm = pd.DataFrame(columns=["Project Manager", "Negative Reviews"])
                            negative_details = pd.DataFrame(columns=[
                                "Project Manager", "Name", "Brand", "Trustpilot Review Date", "Trustpilot Review Links",
                                "Status"
                            ])

                        if data.empty:
                            st.info(f"No data available for {selected_month} {number} for {choice}")
                        else:
                            st.markdown("### 📄 Detailed Entry Data")
                            st.dataframe(data)
                            with st.expander("🧮 Clients with multiple platform publishing"):
                                data_multiple_platforms = data.copy()

                                data_multiple_platforms = data_multiple_platforms[
                                    ~data_multiple_platforms["Issues"].isin(["Printing Only"])]
                                platform_counts = data_multiple_platforms.groupby(["Name", "Book Name & Link"])[
                                    "Platform"].nunique().reset_index(name="Platform_Count")

                                platforms_per_client = data_multiple_platforms.groupby(["Name", "Book Name & Link"])[
                                    "Platform"].apply(
                                    list).reset_index(name="Platforms")
                                platform_stats = platform_counts.merge(platforms_per_client, how="left",
                                                                       on=["Name", "Book Name & Link"])
                                platform_stats.index = range(1, len(platform_stats) + 1)
                                st.dataframe(platform_stats)
                            buffer = io.BytesIO()
                            data.to_excel(buffer, index=False)
                            buffer.seek(0)

                            st.download_button(
                                label="📥 Download Excel",
                                data=buffer,
                                file_name=f"{choice}_{selected_month}_{number}.xlsx",
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                help="Click to download the Excel report"
                            )

                            brands = data_rm_dupes["Brand"].value_counts()
                            writers_clique = brands.get("Writers Clique", "N/A")
                            bookmarketeers = brands.get("BookMarketeers", "N/A")
                            aurora_writers = brands.get("Aurora Writers", "N/A")
                            kdp = brands.get("KDP", "N/A")
                            authors_solution = brands.get("Authors Solution", "N/A")
                            book_publication = brands.get("Book Publication", "N/A")

                            platforms = data["Platform"].value_counts()
                            amazon = platforms.get("Amazon", "N/A")
                            bn = platforms.get("Barnes & Noble", "N/A")
                            ingram = platforms.get("Ingram Spark", "N/A")
                            fav = platforms.get("FAV", "N/A")
                            acx = platforms.get("ACX", "N/A")
                            kobo = platforms.get("Kobo", "N/A")
                            d2d = platforms.get("Draft2Digital", "N/A")

                            filtered_data = data_rm_dupes[data_rm_dupes["Brand"].isin(
                                ["BookMarketeers", "Writers Clique", "Aurora Writers", "Authors Solution",
                                 "Book Publication"])]
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
    
                                            **Platforms**
                                            - 🅰 **Amazon:** `{amazon}`
                                            - 📔 **Barnes & Noble:** `{bn}`
                                            - ⚡ **Ingram Spark:** `{ingram}`
                                            - 📚 **Kobo:** `{kobo}`
                                            - 📚 **Draft2Digital:** `{d2d}`
                                            - 🔉 **Findaway Voices:** `{fav}`
                                            - 🔉 **ACX:** `{acx}`
                                            """)
                                data_rm_dupes.index = range(1, len(data_rm_dupes) + 1)

                                with st.expander(f"🤵🏻 Clients List {choice} {selected_month} {number}"):
                                    st.dataframe(data_rm_dupes)

                                with st.expander(f"📈 Publishing Stats {choice} {selected_month} {number}"):
                                    data_rm_dupes2 = data.copy()
                                    data_rm_dupes2 = data_rm_dupes2.drop_duplicates(["Name"], keep="first")
                                    publishing_stats = data_rm_dupes2.groupby('Publishing Date')["Name"].apply(
                                        list).reset_index(name="Clients")
                                    publishing_counts = data_rm_dupes2.groupby('Publishing Date')[
                                        "Name"].count().reset_index(
                                        name="Counts")
                                    publishing_merged = publishing_counts.merge(publishing_stats, on='Publishing Date',
                                                                                how='left'
                                                                                )
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
                                    attained_count = (
                                        attained_details
                                        .groupby("Project Manager")
                                        .size()
                                        .reset_index(name="Count")
                                    )
                                    attained_clients = (
                                        attained_details
                                        .groupby("Project Manager")
                                        ["Name"].apply(list)
                                        .reset_index(name="Clients")
                                    )
                                    merged_attained = attained_count.merge(attained_clients, on="Project Manager", how="left")
                                    merged_attained = merged_attained.sort_values(by="Count", ascending=False)
                                    merged_attained.index = range(1, len(merged_attained)+1)
                                    st.dataframe(merged_attained)
                                    st.dataframe(attained_details["Status"].value_counts())
                                with st.expander("🏷️ Reviews Per Brand"):
                                    attained_brands = attained_details["Brand"].value_counts()
                                    st.dataframe(
                                        attained_brands)
                                with st.expander("❌ Negative Reviews Per PM"):
                                    st.dataframe(negative_reviews_per_pm)
                                    st.dataframe(negative_details)
                                    st.dataframe(negative_details["Status"].value_counts())

                        st.markdown("---")
                    else:
                        st.info(f"No Data found for {choice} {selected_month} {number}")
            with tab2:
                st.subheader(f"📂 Yearly Data for {choice}")
                number2 = st.number_input("Enter Year", min_value=int(get_min_year()), max_value=current_year,
                                          value=current_year, step=1,
                                          key="year_total")

                if number2 and sheet_name:

                    data = load_data_year(sheet_name, number2)
                    if not data.empty:
                        data_rm_dupes = data.copy()
                        if "Name" in data_rm_dupes.columns:
                            data_rm_dupes = data_rm_dupes.drop_duplicates(subset=["Name"], keep="first")

                        pm_list = list(set((data["Project Manager"].dropna().unique().tolist() + ["Unknown"])))
                        reviews_per_pm = [load_reviews_year(choice, number2, pm, "Attained") for pm in pm_list]
                        reviews_per_pm = safe_concat([df for df in reviews_per_pm if not df.empty])

                        reviews_n_pm = [load_reviews_year(choice, number2, pm, "Negative") for pm in pm_list]
                        reviews_n_pm = safe_concat([df for df in reviews_n_pm if not df.empty])

                        if not reviews_n_pm.empty:
                            negative_pm = (
                                reviews_n_pm.groupby("Project Manager")["Trustpilot Review"]
                                .count()
                                .reset_index()
                            )
                        else:
                            negative_pm = pd.DataFrame(columns=["Project Manager", "Negative Reviews"])

                        if not reviews_per_pm.empty:
                            attained_pm = (
                                reviews_per_pm
                                .groupby("Project Manager")["Trustpilot Review"]
                                .count()
                                .reset_index()
                            )
                        else:
                            attained_pm = pd.DataFrame(columns=["Project Manager", "Attained Reviews"])

                        if not attained_pm.empty:
                            attained_pm.columns = ["Project Manager", "Attained Reviews"]
                            attained_pm = attained_pm.sort_values(by="Attained Reviews", ascending=False)
                            attained_pm.index = range(1, len(attained_pm) + 1)
                            total_attained = attained_pm["Attained Reviews"].sum()
                        else:
                            attained_pm = pd.DataFrame(columns=["Project Manager", "Attained Reviews"])
                            total_attained = 0

                        if not reviews_per_pm.empty:
                            review_details_total = reviews_per_pm.sort_values(by="Project Manager", ascending=True)
                            review_details_total["Trustpilot Review Date"] = pd.to_datetime(
                                review_details_total["Trustpilot Review Date"], errors="coerce"
                            ).dt.strftime("%d-%B-%Y")
                            attained_details_total = review_details_total[
                                review_details_total["Trustpilot Review"] == "Attained"
                                ][["Project Manager", "Name", "Brand", "Trustpilot Review Date",
                                   "Trustpilot Review Links",
                                   "Status"]].copy()
                            attained_details_total.index = range(1, len(attained_details_total) + 1)
                        else:
                            attained_details_total = pd.DataFrame(columns=[
                                "Project Manager", "Name", "Brand", "Trustpilot Review Date", "Trustpilot Review Links",
                                "Status"
                            ])

                        if not negative_pm.empty:
                            negative_pm.columns = ["Project Manager", "Negative Reviews"]
                            negative_pm = negative_pm.sort_values(by="Negative Reviews", ascending=False)
                            negative_pm.index = range(1, len(negative_pm) + 1)
                            total_negative = negative_pm["Negative Reviews"].sum()
                        else:
                            negative_pm = pd.DataFrame(columns=["Project Manager", "Negative Reviews"])
                            total_negative = 0

                        if not reviews_n_pm.empty:
                            review_details_negative = reviews_n_pm.sort_values(by="Project Manager", ascending=True)
                            review_details_negative["Trustpilot Review Date"] = pd.to_datetime(
                                review_details_negative["Trustpilot Review Date"], errors="coerce"
                            ).dt.strftime("%d-%B-%Y")

                            negative_details_total = review_details_negative[
                                review_details_negative["Trustpilot Review"] == "Negative"
                                ][["Project Manager", "Name", "Brand", "Trustpilot Review Date",
                                   "Trustpilot Review Links",
                                   "Status"]].copy()
                            negative_details_total.index = range(1, len(negative_details_total) + 1)
                        else:
                            negative_details_total = pd.DataFrame(columns=[
                                "Project Manager", "Name", "Brand", "Trustpilot Review Date", "Trustpilot Review Links",
                                "Status"
                            ])

                        if not attained_details_total.empty:
                            attained_months_copy = attained_details_total.copy()
                            attained_months_copy["Trustpilot Review Date"] = pd.to_datetime(
                                attained_months_copy["Trustpilot Review Date"], errors="coerce"
                            )

                            attained_reviews_per_month = (
                                attained_months_copy.groupby(
                                    attained_months_copy["Trustpilot Review Date"].dt.to_period("M"))
                                .size()
                                .reset_index(name="Total Attained Reviews")
                            )

                            attained_reviews_per_month["Month"] = attained_reviews_per_month[
                                "Trustpilot Review Date"].dt.strftime("%B %Y")
                            attained_reviews_per_month = attained_reviews_per_month.sort_values(
                                by="Total Attained Reviews", ascending=False
                            )
                            attained_reviews_per_month.index = range(1, len(attained_reviews_per_month) + 1)
                            attained_reviews_per_month = attained_reviews_per_month.drop("Trustpilot Review Date",
                                                                                         axis=1)
                        else:
                            attained_reviews_per_month = pd.DataFrame(columns=["Month", "Total Attained Reviews"])

                        if not negative_details_total.empty:
                            negative_months_copy = negative_details_total.copy()
                            negative_months_copy["Trustpilot Review Date"] = pd.to_datetime(
                                negative_months_copy["Trustpilot Review Date"], errors="coerce"
                            )

                            negative_reviews_per_month = (
                                negative_months_copy.groupby(
                                    negative_months_copy["Trustpilot Review Date"].dt.to_period("M"))
                                .size()
                                .reset_index(name="Total Negative Reviews")
                            )

                            negative_reviews_per_month["Month"] = negative_reviews_per_month[
                                "Trustpilot Review Date"].dt.strftime("%B %Y")
                            negative_reviews_per_month = negative_reviews_per_month.sort_values(
                                by="Total Negative Reviews", ascending=False
                            )
                            negative_reviews_per_month.index = range(1, len(negative_reviews_per_month) + 1)
                            negative_reviews_per_month = negative_reviews_per_month.drop("Trustpilot Review Date",
                                                                                         axis=1)
                        else:
                            negative_reviews_per_month = pd.DataFrame(columns=["Month", "Total Negative Reviews"])

                        if data.empty:
                            st.warning(f"⚠️ No Data Available for {choice} in {number2}")
                        else:
                            st.markdown(f"### 📄 Total Data for {choice} - {number2}")
                            st.dataframe(data)

                            with st.expander("🧮 Clients with multiple platform publishing"):

                                data_multiple_platforms = data.copy()

                                data_multiple_platforms = data_multiple_platforms[
                                    ~data_multiple_platforms["Issues"].isin(["Printing Only"])]
                                platform_counts = data_multiple_platforms.groupby(["Name", "Book Name & Link"])[
                                    "Platform"].nunique().reset_index(name="Platform_Count")

                                platforms_per_client = data_multiple_platforms.groupby(["Name", "Book Name & Link"])[
                                    "Platform"].apply(
                                    list).reset_index(name="Platforms")
                                platform_stats = platform_counts.merge(platforms_per_client, how="left",
                                                                       on=["Name", "Book Name & Link"])
                                platform_stats.index = range(1, len(platform_stats) + 1)
                                st.dataframe(platform_stats)
                            buffer = io.BytesIO()
                            data.to_excel(buffer, index=False)
                            buffer.seek(0)

                            st.download_button(
                                label="📥 Download Excel",
                                data=buffer,
                                file_name=f"{choice}_Total_{number2}.xlsx",
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                help="Click to download the Excel report"
                            )

                            brands = data_rm_dupes["Brand"].value_counts()
                            platforms = data["Platform"].value_counts()
                            publishing = data_rm_dupes["Status"].value_counts()

                            filtered_data = data_rm_dupes[data_rm_dupes["Brand"].isin(
                                ["BookMarketeers", "Writers Clique", "Aurora Writers", "Authors Solution",
                                 "Book Publication"])]
                            pending_sent_details = filtered_data[(filtered_data["Trustpilot Review"] == "Sent") | (
                                    filtered_data["Trustpilot Review"] == "Pending")]
                            review_counts = filtered_data["Trustpilot Review"].value_counts()
                            sent = review_counts.get("Sent", 0)
                            pending = review_counts.get("Pending", 0)
                            attained = total_attained
                            negative = negative_pm["Negative Reviews"].sum()
                            total_reviews = sent + pending + attained + negative
                            percentage = round((attained / total_reviews * 100), 1) if total_reviews > 0 else 0

                            unique_clients_count_per_pm = data_rm_dupes.groupby('Project Manager')[
                                'Name'].nunique().reset_index()
                            unique_clients_count_per_pm.columns = ['Project Manager', 'Unique Clients']
                            unique_clients_count_per_pm.index = range(1, len(unique_clients_count_per_pm) + 1)
                            clients_list = data_rm_dupes.groupby('Project Manager')["Name"].apply(list).reset_index(
                                name="Clients")
                            merged_df = unique_clients_count_per_pm.merge(clients_list, on='Project Manager',
                                                                          how='left')
                            merged_df.index = range(1, len(merged_df) + 1)
                            total_unique_clients = data['Name'].nunique()

                            Issues = data_rm_dupes["Issues"].value_counts()

                            col1, col2 = st.columns(2)
                            with col1:
                                st.markdown("---")
                                st.markdown("### ⭐ Annual Summary")
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
    
                                **Platforms**
                                - 🅰 **Amazon:** `{platforms.get("Amazon", "N/A")}`
                                - 📔 **Barnes & Noble:** `{platforms.get("Barnes & Noble", "N/A")}`
                                - ⚡ **Ingram Spark:** `{platforms.get("Ingram Spark", "N/A")}`
                                - 📚 **Kobo:** `{platforms.get("Kobo", "N/A")}`
                                - 📚 **Draft2Digital:** `{platforms.get("Draft2Digital", "N/A")}`
                                - 🔉 **Findaway Voices:** `{platforms.get("FAV", "N/A")}`
                                - 🔉 **ACX:** `{platforms.get("ACX", "N/A")}`
                                """)
                                data_rm_dupes.index = range(1, len(data_rm_dupes) + 1)

                                with st.expander(f"🤵🏻 Clients List {choice} {number2}"):
                                    st.dataframe(data_rm_dupes)
                                with st.expander("🤵🏻🤵🏻 Publishing Per Month"):
                                    data_month = data_rm_dupes.copy()
                                    data_month["Publishing Date"] = pd.to_datetime(data_month["Publishing Date"],
                                                                                   errors="coerce")

                                    data_month["Month"] = data_month["Publishing Date"].dt.to_period("M").dt.strftime(
                                        "%B %Y")

                                    unique_clients_count_per_month = (
                                        data_month.groupby("Month")["Name"].nunique()
                                        .reset_index()
                                    )
                                    unique_clients_count_per_month.columns = ["Month", "Total Published"]
                                    clients_list_per_month = (
                                        data_month.groupby("Month")["Name"]
                                        .apply(list)
                                        .reset_index(name="Clients")
                                    )

                                    publishing_per_month = unique_clients_count_per_month.merge(
                                        clients_list_per_month, on="Month", how="left"
                                    )

                                    publishing_per_month = publishing_per_month.sort_values(
                                        by="Total Published", ascending=False
                                    )
                                    publishing_per_month.index = range(1, len(publishing_per_month) + 1)
                                    st.dataframe(publishing_per_month)

                                    pm_unique_clients_per_month = (
                                        data_month
                                        .groupby(["Month", "Project Manager"])["Name"]
                                        .nunique()
                                        .reset_index(name="Total Published")
                                    )
                                    pm_unique_clients_per_month_distribution = (
                                        data_month
                                        .groupby(["Month", "Project Manager"])["Name"]
                                        .apply(list)
                                        .reset_index(name="Clients")
                                    )
                                    merged_pm_client_distribution = pm_unique_clients_per_month.merge(
                                        pm_unique_clients_per_month_distribution, on=["Month", "Project Manager"],
                                        how="left")
                                    merged_pm_client_distribution.index = range(1,
                                                                                len(merged_pm_client_distribution) + 1)
                                    st.dataframe(merged_pm_client_distribution)
                                with st.expander(f"📈 Publishing Stats {choice} {number2}"):
                                    data_rm_dupes2 = data.copy()
                                    data_rm_dupes2 = data_rm_dupes2.drop_duplicates(["Name"], keep="first")
                                    publishing_stats = data_rm_dupes2.groupby('Publishing Date')["Name"].apply(
                                        list).reset_index(name="Clients")
                                    publishing_counts = data_rm_dupes2.groupby('Publishing Date')[
                                        "Name"].count().reset_index(
                                        name="Counts")
                                    publishing_merged = publishing_counts.merge(publishing_stats, on='Publishing Date',
                                                                                how='left'
                                                                                )
                                    publishing_merged.index = range(1, len(publishing_merged) + 1)
                                    st.dataframe(publishing_merged)
                                with st.expander(f"💫 Self Publishing List {choice} {number2}"):
                                    self_publishing_df = data_rm_dupes2[data_rm_dupes2["Issues"] == "Self Publishing"]
                                    self_publishing_df.index = range(1, len(self_publishing_df) + 1)
                                    st.dataframe(self_publishing_df)
                                with st.expander(f"🖨 Printing Only List {choice} {number2}"):
                                    printing_only_df = data_rm_dupes2[data_rm_dupes2["Issues"] == "Printing Only"]
                                    printing_only_df.index = range(1, len(printing_only_df) + 1)
                                    st.dataframe(printing_only_df)
                                with st.expander("🟢 Attained Reviews Per Month"):
                                    st.dataframe(attained_reviews_per_month)
                                with st.expander("🔴 Negative Reviews Per Month"):
                                    st.dataframe(negative_reviews_per_month)
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
                                    attained_count = (
                                        attained_details_total
                                        .groupby("Project Manager")
                                        .size()
                                        .reset_index(name="Count")
                                    )
                                    attained_clients = (
                                        attained_details_total
                                        .groupby("Project Manager")
                                        ["Name"].apply(list)
                                        .reset_index(name="Clients")
                                    )
                                    merged_attained = attained_count.merge(attained_clients, on="Project Manager", how="left")
                                    merged_attained = merged_attained.sort_values(by="Count", ascending=False)
                                    merged_attained.index = range(1, len(merged_attained)+1)
                                    st.dataframe(merged_attained)
                                    st.dataframe(attained_details_total["Status"].value_counts())
                                with st.expander("🏷️ Reviews Per Brand"):
                                    attained_brands = attained_details_total["Brand"].value_counts()
                                    st.dataframe(
                                        attained_brands)
                                with st.expander("❌ Negative Reviews Per PM"):
                                    st.dataframe(negative_pm)
                                    st.dataframe(negative_details_total)
                                    st.dataframe(negative_details_total["Status"].value_counts())

                            st.markdown("---")
                    else:
                        st.info(f"No Data Found for {choice} {number2}")
            with tab3:
                st.subheader(f"📂 Start to Year Data for {choice}")
                number5 = st.number_input("Enter Year", min_value=int(get_min_year()), max_value=current_year,
                                          value=get_min_year(), step=1,
                                          key="year_total_to_date_start")
                number4 = st.number_input("Enter Year", min_value=int(get_min_year()), max_value=current_year,
                                          value=current_year, step=1,
                                          key="year_total_to_date")

                if number4 and number5 and sheet_name:

                    data = load_data_search(sheet_name, number4, number5)
                    if not data.empty:
                        data_rm_dupes = data.copy()
                        if "Name" in data_rm_dupes.columns:
                            data_rm_dupes = data_rm_dupes.drop_duplicates(subset=["Name"], keep="first")

                        pm_list = list(set((data["Project Manager"].dropna().unique().tolist() + ["Unknown"])))
                        reviews_per_pm = [load_reviews_year_to_date(choice, number4, pm, "Attained") for pm in pm_list]
                        reviews_per_pm = safe_concat([df for df in reviews_per_pm if not df.empty])

                        reviews_n_pm = [load_reviews_year_to_date(choice, number4, pm, "Negative") for pm in pm_list]
                        reviews_n_pm = safe_concat([df for df in reviews_n_pm if not df.empty])

                        if not reviews_n_pm.empty:
                            negative_pm = (
                                reviews_n_pm.groupby("Project Manager")["Trustpilot Review"]
                                .count()
                                .reset_index()
                            )
                        else:
                            negative_pm = pd.DataFrame(columns=["Project Manager", "Negative Reviews"])

                        if not reviews_per_pm.empty:
                            attained_pm = (
                                reviews_per_pm
                                .groupby("Project Manager")["Trustpilot Review"]
                                .count()
                                .reset_index()
                            )
                        else:
                            attained_pm = pd.DataFrame(columns=["Project Manager", "Attained Reviews"])

                        if not attained_pm.empty:
                            attained_pm.columns = ["Project Manager", "Attained Reviews"]
                            attained_pm = attained_pm.sort_values(by="Attained Reviews", ascending=False)
                            attained_pm.index = range(1, len(attained_pm) + 1)
                            total_attained = attained_pm["Attained Reviews"].sum()
                        else:
                            attained_pm = pd.DataFrame(columns=["Project Manager", "Attained Reviews"])
                            total_attained = 0

                        if not reviews_per_pm.empty:
                            review_details_total = reviews_per_pm.sort_values(by="Project Manager", ascending=True)
                            review_details_total["Trustpilot Review Date"] = pd.to_datetime(
                                review_details_total["Trustpilot Review Date"], errors="coerce"
                            ).dt.strftime("%d-%B-%Y")
                            attained_details_total = review_details_total[
                                review_details_total["Trustpilot Review"] == "Attained"
                                ][["Project Manager", "Name", "Brand", "Trustpilot Review Date",
                                   "Trustpilot Review Links",
                                   "Status"]].copy()
                            attained_details_total.index = range(1, len(attained_details_total) + 1)
                        else:
                            attained_details_total = pd.DataFrame(columns=[
                                "Project Manager", "Name", "Brand", "Trustpilot Review Date", "Trustpilot Review Links",
                                "Status"
                            ])

                        if not negative_pm.empty:
                            negative_pm.columns = ["Project Manager", "Negative Reviews"]
                            negative_pm = negative_pm.sort_values(by="Negative Reviews", ascending=False)
                            negative_pm.index = range(1, len(negative_pm) + 1)
                            total_negative = negative_pm["Negative Reviews"].sum()
                        else:
                            negative_pm = pd.DataFrame(columns=["Project Manager", "Negative Reviews"])
                            total_negative = 0

                        if not reviews_n_pm.empty:
                            review_details_negative = reviews_n_pm.sort_values(by="Project Manager", ascending=True)
                            review_details_negative["Trustpilot Review Date"] = pd.to_datetime(
                                review_details_negative["Trustpilot Review Date"], errors="coerce"
                            ).dt.strftime("%d-%B-%Y")

                            negative_details_total = review_details_negative[
                                review_details_negative["Trustpilot Review"] == "Negative"
                                ][["Project Manager", "Name", "Brand", "Trustpilot Review Date",
                                   "Trustpilot Review Links",
                                   "Status"]].copy()
                            negative_details_total.index = range(1, len(negative_details_total) + 1)
                        else:
                            negative_details_total = pd.DataFrame(columns=[
                                "Project Manager", "Name", "Brand", "Trustpilot Review Date", "Trustpilot Review Links",
                                "Status"
                            ])

                        if not attained_details_total.empty:
                            attained_months_copy = attained_details_total.copy()
                            attained_months_copy["Trustpilot Review Date"] = pd.to_datetime(
                                attained_months_copy["Trustpilot Review Date"], errors="coerce"
                            )

                            attained_reviews_per_month = (
                                attained_months_copy.groupby(
                                    attained_months_copy["Trustpilot Review Date"].dt.to_period("M"))
                                .size()
                                .reset_index(name="Total Attained Reviews")
                            )

                            attained_reviews_per_month["Month"] = attained_reviews_per_month[
                                "Trustpilot Review Date"].dt.strftime("%B %Y")
                            attained_reviews_per_month = attained_reviews_per_month.sort_values(
                                by="Total Attained Reviews", ascending=False
                            )
                            attained_reviews_per_month.index = range(1, len(attained_reviews_per_month) + 1)
                            attained_reviews_per_month = attained_reviews_per_month.drop("Trustpilot Review Date",
                                                                                         axis=1)
                        else:
                            attained_reviews_per_month = pd.DataFrame(columns=["Month", "Total Attained Reviews"])

                        if not negative_details_total.empty:
                            negative_months_copy = negative_details_total.copy()
                            negative_months_copy["Trustpilot Review Date"] = pd.to_datetime(
                                negative_months_copy["Trustpilot Review Date"], errors="coerce"
                            )

                            negative_reviews_per_month = (
                                negative_months_copy.groupby(
                                    negative_months_copy["Trustpilot Review Date"].dt.to_period("M"))
                                .size()
                                .reset_index(name="Total Negative Reviews")
                            )

                            negative_reviews_per_month["Month"] = negative_reviews_per_month[
                                "Trustpilot Review Date"].dt.strftime("%B %Y")
                            negative_reviews_per_month = negative_reviews_per_month.sort_values(
                                by="Total Negative Reviews", ascending=False
                            )
                            negative_reviews_per_month.index = range(1, len(negative_reviews_per_month) + 1)
                            negative_reviews_per_month = negative_reviews_per_month.drop("Trustpilot Review Date",
                                                                                         axis=1)
                        else:
                            negative_reviews_per_month = pd.DataFrame(columns=["Month", "Total Negative Reviews"])

                        if data.empty:
                            st.warning(f"⚠️ No Data Available for {choice} in {number5} to {number4}")
                        else:
                            st.markdown(f"### 📄 Year to Year Data for {choice} - {number5} to {number4}")
                            st.dataframe(data)

                            with st.expander("🧮 Clients with multiple platform publishing"):

                                data_multiple_platforms = data.copy()

                                data_multiple_platforms = data_multiple_platforms[
                                    ~data_multiple_platforms["Issues"].isin(["Printing Only"])]
                                platform_counts = data_multiple_platforms.groupby(["Name", "Book Name & Link"])[
                                    "Platform"].nunique().reset_index(name="Platform_Count")

                                platforms_per_client = data_multiple_platforms.groupby(["Name", "Book Name & Link"])[
                                    "Platform"].apply(
                                    list).reset_index(name="Platforms")
                                platform_stats = platform_counts.merge(platforms_per_client, how="left",
                                                                       on=["Name", "Book Name & Link"])
                                platform_stats.index = range(1, len(platform_stats) + 1)
                                st.dataframe(platform_stats)
                            buffer = io.BytesIO()
                            data.to_excel(buffer, index=False)
                            buffer.seek(0)

                            st.download_button(
                                label="📥 Download Excel",
                                data=buffer,
                                file_name=f"{choice}_Total_{number5}-{number4}.xlsx",
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                help="Click to download the Excel report",
                                key="Year_to_date"
                            )

                            brands = data_rm_dupes["Brand"].value_counts()
                            platforms = data["Platform"].value_counts()
                            publishing = data_rm_dupes["Status"].value_counts()

                            filtered_data = data_rm_dupes[data_rm_dupes["Brand"].isin(
                                ["BookMarketeers", "Writers Clique", "Aurora Writers", "Authors Solution",
                                 "Book Publication"])]
                            pending_sent_details = filtered_data[(filtered_data["Trustpilot Review"] == "Sent") | (
                                    filtered_data["Trustpilot Review"] == "Pending")]
                            review_counts = filtered_data["Trustpilot Review"].value_counts()
                            sent = review_counts.get("Sent", 0)
                            pending = review_counts.get("Pending", 0)
                            attained = total_attained
                            negative = negative_pm["Negative Reviews"].sum()
                            total_reviews = sent + pending + attained + negative
                            percentage = round((attained / total_reviews * 100), 1) if total_reviews > 0 else 0

                            unique_clients_count_per_pm = data_rm_dupes.groupby('Project Manager')[
                                'Name'].nunique().reset_index()
                            unique_clients_count_per_pm.columns = ['Project Manager', 'Unique Clients']
                            unique_clients_count_per_pm.index = range(1, len(unique_clients_count_per_pm) + 1)
                            clients_list = data_rm_dupes.groupby('Project Manager')["Name"].apply(list).reset_index(
                                name="Clients")
                            merged_df = unique_clients_count_per_pm.merge(clients_list, on='Project Manager',
                                                                          how='left')
                            merged_df.index = range(1, len(merged_df) + 1)
                            total_unique_clients = data['Name'].nunique()

                            Issues = data_rm_dupes["Issues"].value_counts()

                            col1, col2 = st.columns(2)
                            with col1:
                                st.markdown("---")
                                st.markdown("### ⭐ Start to Year Summary")
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
    
                                                **Platforms**
                                                - 🅰 **Amazon:** `{platforms.get("Amazon", "N/A")}`
                                                - 📔 **Barnes & Noble:** `{platforms.get("Barnes & Noble", "N/A")}`
                                                - ⚡ **Ingram Spark:** `{platforms.get("Ingram Spark", "N/A")}`
                                                - 📚 **Kobo:** `{platforms.get("Kobo", "N/A")}`
                                                - 📚 **Draft2Digital:** `{platforms.get("Draft2Digital", "N/A")}`
                                                - 🔉 **Findaway Voices:** `{platforms.get("FAV", "N/A")}`
                                                - 🔉 **ACX:** `{platforms.get("ACX", "N/A")}`
                                                """)
                                data_rm_dupes.index = range(1, len(data_rm_dupes) + 1)

                                with st.expander(f"🤵🏻 Clients List {choice} - {number5} to {number4}"):
                                    st.dataframe(data_rm_dupes)
                                with st.expander("🤵🏻🤵🏻 Publishing Per Month"):
                                    data_month = data_rm_dupes.copy()
                                    data_month["Publishing Date"] = pd.to_datetime(data_month["Publishing Date"],
                                                                                   errors="coerce")

                                    data_month["Month"] = data_month["Publishing Date"].dt.to_period("M").dt.strftime(
                                        "%B %Y")

                                    unique_clients_count_per_month = (
                                        data_month.groupby("Month")["Name"].nunique()
                                        .reset_index()
                                    )
                                    unique_clients_count_per_month.columns = ["Month", "Total Published"]
                                    clients_list_per_month = (
                                        data_month.groupby("Month")["Name"]
                                        .apply(list)
                                        .reset_index(name="Clients")
                                    )

                                    publishing_per_month = unique_clients_count_per_month.merge(
                                        clients_list_per_month, on="Month", how="left"
                                    )

                                    publishing_per_month = publishing_per_month.sort_values(
                                        by="Total Published", ascending=False
                                    )
                                    publishing_per_month.index = range(1, len(publishing_per_month) + 1)
                                    st.dataframe(publishing_per_month)

                                    pm_unique_clients_per_month = (
                                        data_month
                                        .groupby(["Month", "Project Manager"])["Name"]
                                        .nunique()
                                        .reset_index(name="Total Published")
                                    )
                                    pm_unique_clients_per_month_distribution = (
                                        data_month
                                        .groupby(["Month", "Project Manager"])["Name"]
                                        .apply(list)
                                        .reset_index(name="Clients")
                                    )
                                    merged_pm_client_distribution = pm_unique_clients_per_month.merge(pm_unique_clients_per_month_distribution, on=["Month", "Project Manager"], how="left")
                                    merged_pm_client_distribution.index = range(1, len(merged_pm_client_distribution)+1)
                                    st.dataframe(merged_pm_client_distribution)

                                    yearly_data = data_rm_dupes.copy()
                                    yearly_data["Publishing Date"] = pd.to_datetime(yearly_data["Publishing Date"],
                                                                                   errors="coerce")
                                    yearly_data["Year"] = yearly_data["Publishing Date"].dt.to_period("Y").dt.strftime(
                                        "%Y")

                                    unique_clients_count_per_year = (
                                        yearly_data.groupby("Year")["Name"].nunique()
                                        .reset_index()
                                    )
                                    unique_clients_count_per_year.columns = ["Year", "Total Published"]
                                    clients_list_per_year = (
                                        yearly_data.groupby("Year")["Name"]
                                        .apply(list)
                                        .reset_index(name="Clients")
                                    )

                                    publishing_per_year = unique_clients_count_per_year.merge(
                                        clients_list_per_year, on="Year", how="left"
                                    )

                                    publishing_per_year = publishing_per_year.sort_values(
                                        by="Total Published", ascending=False
                                    )
                                    publishing_per_year.index = range(1, len(publishing_per_year) + 1)
                                    st.dataframe(publishing_per_year)

                                with st.expander(f"📈 Publishing Stats {choice} - {number5} to {number4}"):
                                    data_rm_dupes2 = data.copy()
                                    data_rm_dupes2 = data_rm_dupes2.drop_duplicates(["Name"], keep="first")
                                    publishing_stats = data_rm_dupes2.groupby('Publishing Date')["Name"].apply(
                                        list).reset_index(name="Clients")
                                    publishing_counts = data_rm_dupes2.groupby('Publishing Date')[
                                        "Name"].count().reset_index(
                                        name="Counts")
                                    publishing_merged = publishing_counts.merge(publishing_stats, on='Publishing Date',
                                                                                how='left'
                                                                                )
                                    publishing_merged.index = range(1, len(publishing_merged) + 1)
                                    st.dataframe(publishing_merged)
                                with st.expander(f"💫 Self Publishing List {choice} - {number5} to {number4}"):
                                    self_publishing_df = data_rm_dupes2[data_rm_dupes2["Issues"] == "Self Publishing"]
                                    self_publishing_df.index = range(1, len(self_publishing_df) + 1)
                                    st.dataframe(self_publishing_df)
                                with st.expander(f"🖨 Printing Only List {choice} - {number5} to {number4}"):
                                    printing_only_df = data_rm_dupes2[data_rm_dupes2["Issues"] == "Printing Only"]
                                    printing_only_df.index = range(1, len(printing_only_df) + 1)
                                    st.dataframe(printing_only_df)
                                with st.expander("🟢 Attained Reviews Per Month"):
                                    st.dataframe(attained_reviews_per_month)
                                with st.expander("🔴 Negative Reviews Per Month"):
                                    st.dataframe(negative_reviews_per_month)
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
                                    attained_count = (
                                        attained_details_total
                                        .groupby("Project Manager")
                                        .size()
                                        .reset_index(name="Count")
                                    )
                                    attained_clients = (
                                        attained_details_total
                                        .groupby("Project Manager")
                                        ["Name"].apply(list)
                                        .reset_index(name="Clients")
                                    )
                                    merged_attained = attained_count.merge(attained_clients, on="Project Manager", how="left")
                                    merged_attained = merged_attained.sort_values(by="Count", ascending=False)
                                    merged_attained.index = range(1, len(merged_attained)+1)
                                    st.dataframe(merged_attained)

                                    st.dataframe(attained_details_total["Status"].value_counts())
                                with st.expander("🏷️ Reviews Per Brand"):
                                    attained_brands = attained_details_total["Brand"].value_counts()
                                    st.dataframe(
                                        attained_brands)
                                with st.expander("❌ Negative Reviews Per PM"):
                                    st.dataframe(negative_pm)
                                    st.dataframe(negative_details_total)
                                    st.dataframe(negative_details_total["Status"].value_counts())

                            st.markdown("---")
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

                    if remove_duplicates:
                        data = load_data_filter(sheet_name, start_date, end_date, True)
                    else:
                        data = load_data_filter(sheet_name, start_date, end_date)


                    if not data.empty:
                        data_rm_dupes = data.copy()
                        if "Name" in data_rm_dupes.columns:
                            data_rm_dupes = data_rm_dupes.drop_duplicates(subset=["Name"], keep="first")

                        pm_list = list(set((data["Project Manager"].dropna().unique().tolist() + ["Unknown"])))
                        reviews_per_pm = [load_reviews_filter(choice, start_date, end_date, pm, "Attained") for pm in pm_list]
                        reviews_per_pm = safe_concat([df for df in reviews_per_pm if not df.empty])

                        reviews_n_pm = [load_reviews_filter(choice, start_date, end_date, pm, "Negative") for pm in pm_list]
                        reviews_n_pm = safe_concat([df for df in reviews_n_pm if not df.empty])

                        if not reviews_n_pm.empty:
                            negative_pm = (
                                reviews_n_pm.groupby("Project Manager")["Trustpilot Review"]
                                .count()
                                .reset_index()
                            )
                        else:
                            negative_pm = pd.DataFrame(columns=["Project Manager", "Negative Reviews"])

                        if not reviews_per_pm.empty:
                            attained_pm = (
                                reviews_per_pm
                                .groupby("Project Manager")["Trustpilot Review"]
                                .count()
                                .reset_index()
                            )
                        else:
                            attained_pm = pd.DataFrame(columns=["Project Manager", "Attained Reviews"])

                        if not attained_pm.empty:
                            attained_pm.columns = ["Project Manager", "Attained Reviews"]
                            attained_pm = attained_pm.sort_values(by="Attained Reviews", ascending=False)
                            attained_pm.index = range(1, len(attained_pm) + 1)
                            total_attained = attained_pm["Attained Reviews"].sum()
                        else:
                            attained_pm = pd.DataFrame(columns=["Project Manager", "Attained Reviews"])
                            total_attained = 0

                        if not reviews_per_pm.empty:
                            review_details_total = reviews_per_pm.sort_values(by="Project Manager", ascending=True)
                            review_details_total["Trustpilot Review Date"] = pd.to_datetime(
                                review_details_total["Trustpilot Review Date"], errors="coerce"
                            ).dt.strftime("%d-%B-%Y")
                            attained_details_total = review_details_total[
                                review_details_total["Trustpilot Review"] == "Attained"
                                ][["Project Manager", "Name", "Brand", "Trustpilot Review Date",
                                   "Trustpilot Review Links",
                                   "Status"]].copy()
                            attained_details_total.index = range(1, len(attained_details_total) + 1)
                        else:
                            attained_details_total = pd.DataFrame(columns=[
                                "Project Manager", "Name", "Brand", "Trustpilot Review Date", "Trustpilot Review Links",
                                "Status"
                            ])

                        if not negative_pm.empty:
                            negative_pm.columns = ["Project Manager", "Negative Reviews"]
                            negative_pm = negative_pm.sort_values(by="Negative Reviews", ascending=False)
                            negative_pm.index = range(1, len(negative_pm) + 1)
                            total_negative = negative_pm["Negative Reviews"].sum()
                        else:
                            negative_pm = pd.DataFrame(columns=["Project Manager", "Negative Reviews"])
                            total_negative = 0

                        if not reviews_n_pm.empty:
                            review_details_negative = reviews_n_pm.sort_values(by="Project Manager", ascending=True)
                            review_details_negative["Trustpilot Review Date"] = pd.to_datetime(
                                review_details_negative["Trustpilot Review Date"], errors="coerce"
                            ).dt.strftime("%d-%B-%Y")

                            negative_details_total = review_details_negative[
                                review_details_negative["Trustpilot Review"] == "Negative"
                                ][["Project Manager", "Name", "Brand", "Trustpilot Review Date",
                                   "Trustpilot Review Links",
                                   "Status"]].copy()
                            negative_details_total.index = range(1, len(negative_details_total) + 1)
                        else:
                            negative_details_total = pd.DataFrame(columns=[
                                "Project Manager", "Name", "Brand", "Trustpilot Review Date", "Trustpilot Review Links",
                                "Status"
                            ])

                        if not attained_details_total.empty:
                            attained_months_copy = attained_details_total.copy()
                            attained_months_copy["Trustpilot Review Date"] = pd.to_datetime(
                                attained_months_copy["Trustpilot Review Date"], errors="coerce"
                            )

                            attained_reviews_per_month = (
                                attained_months_copy.groupby(
                                    attained_months_copy["Trustpilot Review Date"].dt.to_period("M"))
                                .size()
                                .reset_index(name="Total Attained Reviews")
                            )

                            attained_reviews_per_month["Month"] = attained_reviews_per_month[
                                "Trustpilot Review Date"].dt.strftime("%B %Y")
                            attained_reviews_per_month = attained_reviews_per_month.sort_values(
                                by="Total Attained Reviews", ascending=False
                            )
                            attained_reviews_per_month.index = range(1, len(attained_reviews_per_month) + 1)
                            attained_reviews_per_month = attained_reviews_per_month.drop("Trustpilot Review Date",
                                                                                         axis=1)
                        else:
                            attained_reviews_per_month = pd.DataFrame(columns=["Month", "Total Attained Reviews"])

                        if not negative_details_total.empty:
                            negative_months_copy = negative_details_total.copy()
                            negative_months_copy["Trustpilot Review Date"] = pd.to_datetime(
                                negative_months_copy["Trustpilot Review Date"], errors="coerce"
                            )

                            negative_reviews_per_month = (
                                negative_months_copy.groupby(
                                    negative_months_copy["Trustpilot Review Date"].dt.to_period("M"))
                                .size()
                                .reset_index(name="Total Negative Reviews")
                            )

                            negative_reviews_per_month["Month"] = negative_reviews_per_month[
                                "Trustpilot Review Date"].dt.strftime("%B %Y")
                            negative_reviews_per_month = negative_reviews_per_month.sort_values(
                                by="Total Negative Reviews", ascending=False
                            )
                            negative_reviews_per_month.index = range(1, len(negative_reviews_per_month) + 1)
                            negative_reviews_per_month = negative_reviews_per_month.drop("Trustpilot Review Date",
                                                                                         axis=1)
                        else:
                            negative_reviews_per_month = pd.DataFrame(columns=["Month", "Total Negative Reviews"])

                        if data.empty:
                            st.warning(f"⚠️ No Data Available for {choice} in {start_date.strftime("%B %Y")} to {end_date.strftime("%B %Y")}")
                        else:
                            st.markdown(f"### 📄 Year to Year Data for {choice} - {start_date.strftime("%B %Y")} to {end_date.strftime("%B %Y")}")
                            st.dataframe(data)

                            with st.expander("🧮 Clients with multiple platform publishing"):

                                data_multiple_platforms = data.copy()

                                data_multiple_platforms = data_multiple_platforms[
                                    ~data_multiple_platforms["Issues"].isin(["Printing Only"])]
                                platform_counts = data_multiple_platforms.groupby(["Name", "Book Name & Link"])[
                                    "Platform"].nunique().reset_index(name="Platform_Count")

                                platforms_per_client = data_multiple_platforms.groupby(["Name", "Book Name & Link"])[
                                    "Platform"].apply(
                                    list).reset_index(name="Platforms")
                                platform_stats = platform_counts.merge(platforms_per_client, how="left",
                                                                       on=["Name", "Book Name & Link"])
                                platform_stats.index = range(1, len(platform_stats) + 1)
                                st.dataframe(platform_stats)
                            buffer = io.BytesIO()
                            data.to_excel(buffer, index=False)
                            buffer.seek(0)

                            st.download_button(
                                label="📥 Download Excel",
                                data=buffer,
                                file_name=f"{choice}_Total_{start_date.strftime("%B %Y")} to {end_date.strftime("%B %Y")}.xlsx",
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                help="Click to download the Excel report",
                                key="Filtered_data"
                            )

                            brands = data_rm_dupes["Brand"].value_counts()
                            platforms = data["Platform"].value_counts()
                            publishing = data_rm_dupes["Status"].value_counts()

                            filtered_data = data_rm_dupes[data_rm_dupes["Brand"].isin(
                                ["BookMarketeers", "Writers Clique", "Aurora Writers", "Authors Solution",
                                 "Book Publication"])]
                            pending_sent_details = filtered_data[(filtered_data["Trustpilot Review"] == "Sent") | (
                                    filtered_data["Trustpilot Review"] == "Pending")]
                            review_counts = filtered_data["Trustpilot Review"].value_counts()
                            sent = review_counts.get("Sent", 0)
                            pending = review_counts.get("Pending", 0)
                            attained = total_attained
                            negative = negative_pm["Negative Reviews"].sum()
                            total_reviews = sent + pending + attained + negative
                            percentage = round((attained / total_reviews * 100), 1) if total_reviews > 0 else 0

                            unique_clients_count_per_pm = data_rm_dupes.groupby('Project Manager')[
                                'Name'].nunique().reset_index()
                            unique_clients_count_per_pm.columns = ['Project Manager', 'Unique Clients']
                            unique_clients_count_per_pm.index = range(1, len(unique_clients_count_per_pm) + 1)
                            clients_list = data_rm_dupes.groupby('Project Manager')["Name"].apply(list).reset_index(
                                name="Clients")
                            merged_df = unique_clients_count_per_pm.merge(clients_list, on='Project Manager',
                                                                          how='left')
                            merged_df.index = range(1, len(merged_df) + 1)
                            total_unique_clients = data['Name'].nunique()

                            Issues = data_rm_dupes["Issues"].value_counts()

                            col1, col2 = st.columns(2)
                            with col1:
                                st.markdown("---")
                                st.markdown("### ⭐ Filtered Summary")
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

                                                **Platforms**
                                                - 🅰 **Amazon:** `{platforms.get("Amazon", "N/A")}`
                                                - 📔 **Barnes & Noble:** `{platforms.get("Barnes & Noble", "N/A")}`
                                                - ⚡ **Ingram Spark:** `{platforms.get("Ingram Spark", "N/A")}`
                                                - 📚 **Kobo:** `{platforms.get("Kobo", "N/A")}` 
                                                - 📚 **Draft2Digital:** `{platforms.get("Draft2Digital", "N/A")}`
                                                - 🔉 **Findaway Voices:** `{platforms.get("FAV", "N/A")}`
                                                - 🔉 **ACX:** `{platforms.get("ACX", "N/A")}`
                                                """)
                                data_rm_dupes.index = range(1, len(data_rm_dupes) + 1)

                                with st.expander(f"🤵🏻 Clients List {choice} - {start_date.strftime("%B %Y")} to {end_date.strftime("%B %Y")}"):
                                    st.dataframe(data_rm_dupes)
                                with st.expander("🤵🏻🤵🏻 Publishing Per Month"):
                                    data_month = data_rm_dupes.copy()
                                    data_month["Publishing Date"] = pd.to_datetime(data_month["Publishing Date"],
                                                                                   errors="coerce")

                                    data_month["Month"] = data_month["Publishing Date"].dt.to_period("M").dt.strftime(
                                        "%B %Y")

                                    unique_clients_count_per_month = (
                                        data_month.groupby("Month")["Name"].nunique()
                                        .reset_index()
                                    )
                                    unique_clients_count_per_month.columns = ["Month", "Total Published"]
                                    clients_list_per_month = (
                                        data_month.groupby("Month")["Name"]
                                        .apply(list)
                                        .reset_index(name="Clients")
                                    )

                                    publishing_per_month = unique_clients_count_per_month.merge(
                                        clients_list_per_month, on="Month", how="left"
                                    )

                                    publishing_per_month = publishing_per_month.sort_values(
                                        by="Total Published", ascending=False
                                    )
                                    publishing_per_month.index = range(1, len(publishing_per_month) + 1)
                                    st.dataframe(publishing_per_month)

                                    pm_unique_clients_per_month = (
                                        data_month
                                        .groupby(["Month", "Project Manager"])["Name"]
                                        .nunique()
                                        .reset_index(name="Total Published")
                                    )
                                    pm_unique_clients_per_month_distribution = (
                                        data_month
                                        .groupby(["Month", "Project Manager"])["Name"]
                                        .apply(list)
                                        .reset_index(name="Clients")
                                    )
                                    merged_pm_client_distribution = pm_unique_clients_per_month.merge(
                                        pm_unique_clients_per_month_distribution, on=["Month", "Project Manager"],
                                        how="left")
                                    merged_pm_client_distribution.index = range(1,
                                                                                len(merged_pm_client_distribution) + 1)
                                    st.dataframe(merged_pm_client_distribution)
                                with st.expander(f"📈 Publishing Stats {choice} - {start_date.strftime("%B %Y")} to {end_date.strftime("%B %Y")}"):
                                    data_rm_dupes2 = data.copy()
                                    data_rm_dupes2 = data_rm_dupes2.drop_duplicates(["Name"], keep="first")
                                    publishing_stats = data_rm_dupes2.groupby('Publishing Date')["Name"].apply(
                                        list).reset_index(name="Clients")
                                    publishing_counts = data_rm_dupes2.groupby('Publishing Date')[
                                        "Name"].count().reset_index(
                                        name="Counts")
                                    publishing_merged = publishing_counts.merge(publishing_stats, on='Publishing Date',
                                                                                how='left'
                                                                                )
                                    publishing_merged.index = range(1, len(publishing_merged) + 1)
                                    st.dataframe(publishing_merged)
                                with st.expander(f"💫 Self Publishing List {choice} - {start_date.strftime("%B %Y")} to {end_date.strftime("%B %Y")}"):
                                    self_publishing_df = data_rm_dupes2[data_rm_dupes2["Issues"] == "Self Publishing"]
                                    self_publishing_df.index = range(1, len(self_publishing_df) + 1)
                                    st.dataframe(self_publishing_df)
                                with st.expander(f"🖨 Printing Only List {choice} - {start_date.strftime("%B %Y")} to {end_date.strftime("%B %Y")}"):
                                    printing_only_df = data_rm_dupes2[data_rm_dupes2["Issues"] == "Printing Only"]
                                    printing_only_df.index = range(1, len(printing_only_df) + 1)
                                    st.dataframe(printing_only_df)
                                with st.expander("🟢 Attained Reviews Per Month"):
                                    st.dataframe(attained_reviews_per_month)
                                with st.expander("🔴 Negative Reviews Per Month"):
                                    st.dataframe(negative_reviews_per_month)
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
                                    attained_count = (
                                        attained_details_total
                                        .groupby("Project Manager")
                                        .size()
                                        .reset_index(name="Count")
                                    )
                                    attained_clients = (
                                        attained_details_total
                                        .groupby("Project Manager")
                                        ["Name"].apply(list)
                                        .reset_index(name="Clients")
                                    )
                                    merged_attained = attained_count.merge(attained_clients, on="Project Manager", how="left")
                                    merged_attained = merged_attained.sort_values(by="Count", ascending=False)
                                    merged_attained.index = range(1, len(merged_attained)+1)
                                    st.dataframe(merged_attained)
                                    st.dataframe(attained_details_total["Status"].value_counts())
                                with st.expander("🏷️ Reviews Per Brand"):
                                    attained_brands = attained_details_total["Brand"].value_counts()
                                    st.dataframe(
                                        attained_brands)
                                with st.expander("❌ Negative Reviews Per PM"):
                                    st.dataframe(negative_pm)
                                    st.dataframe(negative_details_total)
                                    st.dataframe(negative_details_total["Status"].value_counts())

                            st.markdown("---")
                    else:
                        st.info(f"No Data Found for {choice} - {start_date.strftime("%B %Y")} to {end_date.strftime("%B %Y")}")
            with tab5:
                st.subheader(f"🔍 Search Data for {choice}")

                number3 = st.number_input("Enter Year for Search", min_value=int(get_min_year()),
                                          max_value=current_year, value=current_year, step=1,
                                          key="year_search")

                if number3 and sheet_name:
                    data = load_data_search(sheet_name, number3)

                    if data.empty:
                        st.warning(f"⚠️ No Data Available for {choice} in {number3}")
                    else:
                        search_term = st.text_input("Search by Name / Book / Email",
                                                    placeholder="Enter client name, email or book to search",
                                                    key="search_term")

                        if search_term and search_term.strip():
                            search_term_clean = search_term.strip()
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

                                buffer = io.BytesIO()
                                search_df.to_excel(buffer, index=False)
                                buffer.seek(0)

                                st.download_button(
                                    label="📥 Download Search Results",
                                    data=buffer,
                                    file_name=f"{choice}_Search_{search_term}_{number3}.xlsx",
                                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                    help="Click to download search results"
                                )

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
                        else:
                            st.info("👆 Enter name/book/email above to search")
            with tab6:
                st.subheader(f"📊 Filter Data by Brand for {choice}")

                number4 = st.number_input(
                    "Select Year",
                    min_value=int(get_min_year()),
                    max_value=current_year,
                    value=current_year,
                    step=1,
                    key="year_filter"
                )

                usa_brands = ["BookMarketeers", "Writers Clique", "KDP", "Aurora Writers"]
                uk_brands = ["Authors Solution", "Book Publication"]

                if number4 and sheet_name:
                    selected_brand = usa_brands if sheet_name == "USA" else uk_brands
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

                            buffer = io.BytesIO()
                            filtered_df.to_excel(buffer, index=False)
                            buffer.seek(0)
                            st.download_button(
                                label="📥 Download Filtered Data",
                                data=buffer,
                                file_name=f"{choice}_Brand_{brand_selection}_{number4}.xlsx",
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                help="Click to download filtered data"
                            )
        elif action == "Printing":
            tab1, tab2, tab3, tab4, tab5 = st.tabs(["Monthly", "Yearly", "Start to Year", "Search", "Stats"])

            with tab1:
                usa_brands = ["BookMarketeers", "Writers Clique", "Aurora Writers", "KDP"]
                uk_brands = ["Authors Solution", "Book Publication"]
                selected_month = st.selectbox(
                    "Select Month",
                    month_list,
                    index=current_month - 1,
                    placeholder="Select Month"
                )
                number = st.number_input("Enter Year", min_value=int(get_min_year()), max_value=current_year,
                                         value=current_year, step=1)
                selected_month_number = month_list.index(selected_month) + 1 if selected_month else None

                if selected_month and number:
                    st.subheader(f"🖨️ Printing Summary for {selected_month} {number}")

                    data = get_printing_data_month(selected_month_number, number)

                    if not data.empty:
                        show_data = data.copy()
                        show_data["Order Cost"] = show_data["Order Cost"].map("${:,.2f}".format)

                        Total_copies = data["No of Copies"].sum()

                        Total_cost = data["Order Cost"].sum()

                        Highest_cost = data["Order Cost"].max()

                        Highest_copies = data["No of Copies"].max()

                        Lowest_cost = data["Order Cost"].min()

                        Lowest_copies = data["No of Copies"].min()

                        Average = round(Total_cost / Total_copies, 2) if Total_copies else 0

                        st.markdown("### 📄 Detailed Printing Data")

                        st.dataframe(show_data)
                        buffer = io.BytesIO()
                        data.to_excel(buffer, index=False)
                        buffer.seek(0)

                        st.download_button(
                            label="📥 Download Excel",
                            data=buffer,
                            file_name=f"Printing_{selected_month}_{number}.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            help="Click to download the Excel report"
                        )
                        st.markdown("---")

                        st.markdown("### 📊 Summary Statistics")

                        st.markdown(f"""
        
                           - 🧾 **Total Orders:** {len(data)}
        
                           - 📦 **Total Copies Printed:** `{Total_copies}`
        
                           - 💰 **Total Cost:** `${Total_cost:,.2f}`
        
                           - 📈 **Highest Order Cost:** `${Highest_cost:,.2f}`
        
                           - 📉 **Lowest Order Cost:** `${Lowest_cost:,.2f}`
        
                           - 🔢 **Highest Copies in One Order:** `{Highest_copies}`
        
                           - 🧮 **Lowest Copies in One Order:** `{Lowest_copies}`
        
                           - 🧾 **Average Cost per Copy:** `${Average:,.2f}`
        
                           """)
                        st.markdown("---")

                        usa_data = data[data["Brand"].isin(usa_brands)]
                        uk_data = data[data["Brand"].isin(uk_brands)]

                        def show_country_stats(df, country_name):
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

                            brand_spending = (
                                df.groupby("Brand")["Order Cost"]
                                .sum()
                                .reset_index()
                                .sort_values(by="Order Cost", ascending=False)
                            )
                            brand_spending["Order Cost"] = brand_spending["Order Cost"].map("${:,.2f}".format)
                            brand_spending.index = range(1, len(brand_spending) + 1)

                            brand_orders = (
                                df.groupby("Brand")["No of Copies"]
                                .sum()
                                .reset_index()
                                .sort_values(by="No of Copies", ascending=False)

                            )
                            brand_orders.index = range(1, len(brand_orders) + 1)

                            st.markdown(f"#### 💼 Brand-wise Spending for {country_name}")
                            st.dataframe(brand_spending, width="stretch")
                            st.markdown(f"#### 💼 Brand-wise Orders for {country_name}")
                            st.dataframe(brand_orders, width="stretch")

                            st.markdown("---")

                        usa_col, uk_col = st.columns(2)

                        with usa_col:
                            show_country_stats(usa_data, "USA 🦅")
                        with uk_col:
                            show_country_stats(uk_data, "UK ☕")

                    else:
                        st.warning(f"⚠️ No Data Available for Printing in {selected_month} {number}")
            with tab2:
                number2 = st.number_input("Enter Year2", min_value=int(get_min_year()), max_value=current_year,
                                          value=current_year, step=1)
                usa_brands = ["BookMarketeers", "Writers Clique", "Aurora Writers", "KDP"]
                uk_brands = ["Authors Solution", "Book Publication"]
                data, monthly = printing_data_year(number2)

                if not data.empty:
                    st.markdown(f"### 📄 Yearly Printing Data for {number2}")
                    show_data = data.copy()
                    show_data["Order Cost"] = show_data["Order Cost"].map("${:,.2f}".format)
                    st.dataframe(show_data)

                    Total_copies = data["No of Copies"].sum()
                    Total_cost = data["Order Cost"].sum()
                    Highest_cost = data["Order Cost"].max()
                    Highest_copies = data["No of Copies"].max()
                    Lowest_cost = data["Order Cost"].min()
                    Lowest_copies = data["No of Copies"].min()
                    Average = round(Total_cost / Total_copies, 2) if Total_copies else 0

                    buffer = io.BytesIO()
                    data.to_excel(buffer, index=False)
                    buffer.seek(0)

                    st.download_button(
                        label="📥 Download Excel",
                        data=buffer,
                        file_name=f"Printing_{number2}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        help="Click to download the Excel report"
                    )

                    st.markdown("---")

                    st.markdown("### 📊 Summary Statistics (All Data)")
                    st.markdown(f"""
                    - 🧾 **Total Orders:** {len(data)}
                    - 📦 **Total Copies Printed:** `{Total_copies}`
                    - 💰 **Total Cost:** `${Total_cost:,.2f}`
                    - 📈 **Highest Order Cost:** `${Highest_cost:,.2f}`
                    - 📉 **Lowest Order Cost:** `${Lowest_cost:,.2f}`
                    - 🔢 **Highest Copies in One Order:** `{Highest_copies}`
                    - 🧮 **Lowest Copies in One Order:** `{Lowest_copies}`
                    - 💵 **Average Cost per Copy:** `${Average:,.2f}`
                    """)

                    usa_data = data[data["Brand"].isin(usa_brands)]
                    uk_data = data[data["Brand"].isin(uk_brands)]

                    def show_country_stats(df, country_name):
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

                        brand_spending = (
                            df.groupby("Brand")["Order Cost"]
                            .sum()
                            .reset_index()
                            .sort_values(by="Order Cost", ascending=False)
                        )
                        brand_spending["Order Cost"] = brand_spending["Order Cost"].map("${:,.2f}".format)
                        brand_spending.index = range(1, len(brand_spending) + 1)

                        brand_orders = (
                            df.groupby("Brand")["No of Copies"]
                            .sum()
                            .reset_index()
                            .sort_values(by="No of Copies", ascending=False)

                        )
                        brand_orders.index = range(1, len(brand_orders) + 1)

                        st.markdown(f"#### 💼 Brand-wise Spending in {country_name}")
                        st.dataframe(brand_spending, width="stretch")
                        st.markdown(f"#### 💼 Brand-wise Orders for {country_name}")
                        st.dataframe(brand_orders, width="stretch")

                        st.markdown("---")

                    usa_col, uk_col = st.columns(2)

                    with usa_col:
                        show_country_stats(usa_data, "USA 🦅")
                    with uk_col:
                        show_country_stats(uk_data, "UK ☕")

                else:
                    st.warning(f"⚠️ No Data Available for Printing in {number2}")
            with tab3:
                number2 = st.number_input("Enter Year2", min_value=int(get_min_year()), max_value=current_year,
                                          value=current_year, step=1, key="printing_year_to_year")
                usa_brands = ["BookMarketeers", "Writers Clique", "Aurora Writers", "KDP"]
                uk_brands = ["Authors Solution", "Book Publication"]
                data, monthly = printing_data_search(number2)

                if not data.empty:
                    st.markdown(f"### 📄 Start to Year Printing Data for 2025 to {number2}")
                    show_data = data.copy()
                    show_data["Order Cost"] = show_data["Order Cost"].map("${:,.2f}".format)
                    st.dataframe(show_data)

                    Total_copies = data["No of Copies"].sum()
                    Total_cost = data["Order Cost"].sum()
                    Highest_cost = data["Order Cost"].max()
                    Highest_copies = data["No of Copies"].max()
                    Lowest_cost = data["Order Cost"].min()
                    Lowest_copies = data["No of Copies"].min()
                    Average = round(Total_cost / Total_copies, 2) if Total_copies else 0

                    buffer = io.BytesIO()
                    data.to_excel(buffer, index=False)
                    buffer.seek(0)

                    st.download_button(
                        label="📥 Download Excel",
                        data=buffer,
                        file_name=f"Printing_Start to Year_{number2}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        help="Click to download the Excel report",
                        key="Start_to_Year"
                    )

                    st.markdown("---")

                    st.markdown("### 📊 Summary Statistics (Start to Year)")
                    st.markdown(f"""
                    - 🧾 **Total Orders:** {len(data)}
                    - 📦 **Total Copies Printed:** `{Total_copies}`
                    - 💰 **Total Cost:** `${Total_cost:,.2f}`
                    - 📈 **Highest Order Cost:** `${Highest_cost:,.2f}`
                    - 📉 **Lowest Order Cost:** `${Lowest_cost:,.2f}`
                    - 🔢 **Highest Copies in One Order:** `{Highest_copies}`
                    - 🧮 **Lowest Copies in One Order:** `{Lowest_copies}`
                    - 💵 **Average Cost per Copy:** `${Average:,.2f}`
                    """)

                    usa_data = data[data["Brand"].isin(usa_brands)]
                    uk_data = data[data["Brand"].isin(uk_brands)]

                    def show_country_stats(df, country_name):
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

                        brand_spending = (
                            df.groupby("Brand")["Order Cost"]
                            .sum()
                            .reset_index()
                            .sort_values(by="Order Cost", ascending=False)
                        )
                        brand_spending["Order Cost"] = brand_spending["Order Cost"].map("${:,.2f}".format)
                        brand_spending.index = range(1, len(brand_spending) + 1)

                        brand_orders = (
                            df.groupby("Brand")["No of Copies"]
                            .sum()
                            .reset_index()
                            .sort_values(by="No of Copies", ascending=False)

                        )
                        brand_orders.index = range(1, len(brand_orders) + 1)

                        st.markdown(f"#### 💼 Brand-wise Spending in {country_name}")
                        st.dataframe(brand_spending, width="stretch")
                        st.markdown(f"#### 💼 Brand-wise Orders for {country_name}")
                        st.dataframe(brand_orders, width="stretch")

                        st.markdown("---")

                    usa_col, uk_col = st.columns(2)

                    with usa_col:
                        show_country_stats(usa_data, "USA 🦅")
                    with uk_col:
                        show_country_stats(uk_data, "UK ☕")

                else:
                    st.warning(f"⚠️ No Data Available for Printing in Start to Year {number2}")
            with tab4:
                number3 = st.number_input("Enter Year3", min_value=int(get_min_year()), max_value=current_year,
                                          value=current_year, step=1)
                data, _ = printing_data_search(number3)
                search_term = st.text_input("Search by Name / Book", placeholder="Enter Search Term", key="search_term")

                if search_term and search_term.strip():
                    search_term_clean = search_term.strip()
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

                        st.markdown(f"### 🌍 Printing Summary")
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

                year1 = st.number_input(
                    "Enter Previous Year",
                    min_value=int(get_min_year()),
                    max_value=current_year,
                    value=current_year - 1,
                    step=1
                )

                year2 = st.number_input(
                    "Enter Current Year",
                    min_value=int(get_min_year()),
                    max_value=current_year,
                    value=current_year,
                    step=1
                )

                data1, _ = printing_data_year(year1)
                data2, _ = printing_data_year(year2)

                def pct_change(new, old):
                    return round(((new - old) / old) * 100, 2) if old else 0

                if not data1.empty and not data2.empty:

                    # ---------------- OVERALL TOTALS ----------------
                    total_orders1 = len(data1)
                    total_orders2 = len(data2)

                    total_copies1 = data1["No of Copies"].sum()
                    total_copies2 = data2["No of Copies"].sum()

                    total_cost1 = data1["Order Cost"].sum()
                    total_cost2 = data2["Order Cost"].sum()

                    st.subheader("🌍 Overall Printing Comparison")

                    # Forward (Year1 → Year2)
                    col1, col2, col3 = st.columns(3)
                    col1.metric(
                        f"🧾 Orders ({year1} → {year2})",
                        total_orders2,
                        f"{pct_change(total_orders2, total_orders1)}%"
                    )
                    col2.metric(
                        f"📦 Copies ({year1} → {year2})",
                        total_copies2,
                        f"{pct_change(total_copies2, total_copies1)}%"
                    )
                    col3.metric(
                        f"💰 Cost ({year1} → {year2})",
                        f"${total_cost2:,.2f}",
                        f"{pct_change(total_cost2, total_cost1)}%"
                    )


                    col1, col2, col3 = st.columns(3)
                    col1.metric(
                        f"🧾 Orders ({year2} <- {year1})",
                        total_orders1,
                        f"{pct_change(total_orders1, total_orders2)}%"
                    )
                    col2.metric(
                        f"📦 Copies ({year2} <- {year1})",
                        total_copies1,
                        f"{pct_change(total_copies1, total_copies2)}%"
                    )
                    col3.metric(
                        f"💰 Cost ({year2} <- {year1})",
                        f"${total_cost1:,.2f}",
                        f"{pct_change(total_cost1, total_cost2)}%"
                    )

                    st.markdown("---")

                    usa_brands = ["BookMarketeers", "Writers Clique", "Aurora Writers", "KDP"]
                    uk_brands = ["Authors Solution", "Book Publication"]

                    def country_stats(df, brands):
                        df = df[df["Brand"].isin(brands)]
                        return len(df), df["No of Copies"].sum(), df["Order Cost"].sum()

                    usa_orders1, usa_copies1, usa_cost1 = country_stats(data1, usa_brands)
                    usa_orders2, usa_copies2, usa_cost2 = country_stats(data2, usa_brands)

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


                    uk_orders1, uk_copies1, uk_cost1 = country_stats(data1, uk_brands)
                    uk_orders2, uk_copies2, uk_cost2 = country_stats(data2, uk_brands)

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

                selected_month = st.selectbox(

                    "Select Month",

                    month_list,

                    index=current_month - 1,

                    placeholder="Select Month"

                )

                number = st.number_input(

                    "Enter Year",

                    min_value=int(get_min_year()),

                    max_value=current_year,

                    value=current_year,

                    step=1,
                    key="number_copyright"

                )

                selected_month_number = month_list.index(selected_month) + 1 if selected_month else None

                if selected_month and number:

                    st.subheader(f"© Copyright Summary for {selected_month} {number}")

                    data, approved, rejected = get_copyright_month(selected_month_number, number)

                    if not data.empty:

                        st.dataframe(data)

                        buffer = io.BytesIO()

                        data.to_excel(buffer, index=False)

                        buffer.seek(0)

                        st.download_button(

                            label="📥 Download Excel",

                            data=buffer,

                            file_name=f"Copyright_{selected_month}_{number}.xlsx",

                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

                        )
                        total_titles = len(data)
                        country_counts = data["Country"].value_counts()

                        country_usa = country_counts.get("USA", 0)
                        country_uk = country_counts.get("UK", 0)
                        country_canada = country_counts.get("Canada", 0)
                        total_titles = len(data)

                        total_cost = (country_usa * 65) + (country_canada * 63) + (country_uk * 35)


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


                    else:

                        st.warning(f"⚠️ No Data Available for {selected_month} {number}")


            with tab2:

                number2 = st.number_input(

                    "Enter Year",

                    min_value=int(get_min_year()),

                    max_value=current_year,

                    value=current_year,

                    step=1,

                    key="copyright_year_total"

                )

                data, approved, rejected = copyright_year(number2)

                if not data.empty:

                    st.subheader(f"© Yearly Copyright Data for {number2}")

                    st.dataframe(data)

                    buffer = io.BytesIO()

                    data.to_excel(buffer, index=False)

                    buffer.seek(0)

                    st.download_button(

                        label="📥 Download Excel",

                        data=buffer,

                        file_name=f"Copyright_{number2}.xlsx",

                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

                    )

                    total_titles = len(data)

                    country_counts = data["Country"].value_counts()

                    country_usa = country_counts.get("USA", 0)
                    country_uk = country_counts.get("UK", 0)
                    country_canada = country_counts.get("Canada", 0)
                    total_titles = len(data)

                    total_cost = (country_usa * 65) + (country_canada * 63) + (country_uk * 35)


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


                else:

                    st.warning(f"⚠️ No Data Available for {number2}")

            with tab3:

                number3 = st.number_input(

                    "Enter Year",

                    min_value=int(get_min_year()),

                    max_value=current_year,

                    value=current_year,

                    step=1,

                    key="copyright_search"

                )

                data, _, _ = copyright_search(number3)

                search_term = st.text_input(

                    "Search by Title / Name",

                    placeholder="Enter Search Term"

                )

                if search_term and not data.empty:

                    search_df = data[

                        data["Book Name & Link"].str.contains(search_term, case=False, na=False)

                        | data["Name"].str.contains(search_term, case=False, na=False)

                        ]

                    if search_df.empty:

                        st.warning("No matching records found.")

                    else:

                        total_titles = len(search_df)

                        country_counts = search_df["Country"].value_counts()


                        country_usa = country_counts.get("USA", 0)
                        country_uk = country_counts.get("UK", 0)
                        country_canada = country_counts.get("Canada", 0)

                        total_cost = (country_usa * 65) + (country_canada * 63) + (country_uk * 35)

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

                        search_df.index = range(1, len(search_df) + 1)

                        st.dataframe(search_df)
                else:
                    st.info("👆 Enter name/book above to search")
        elif action == "Generate Similarity":

            tab1, tab2, tab3 = st.tabs(["Queries", "Yearly Queries", "Compare Years"])

            def safe_month_index(month_offset: int, month_list_len: int) -> int:
                """Ensure selectbox index is within valid range."""
                return max(0, min(month_offset, month_list_len - 1))

            with tab1:
                st.header("Compare clients with months")
                choice = st.selectbox(
                    "Select Data To View",
                    ["USA", "UK"],
                    index=None,
                    key="choice_tab1"
                )
                sheet_name = {"UK": sheet_uk, "USA": sheet_usa}.get(choice)


                index_month1 = safe_month_index(current_month - 2, len(month_list))
                index_month2 = safe_month_index(current_month - 1, len(month_list))

                selected_month_1 = st.selectbox(
                    "Select Month 1",
                    month_list,
                    index=index_month1,
                    key="month1_tab1"
                )

                number1 = st.number_input(
                    "Enter Year 1",
                    min_value=int(get_min_year()),
                    max_value=current_year,
                    value=current_year,
                    step=1,
                    key="year1_tab1"
                )

                selected_month_2 = st.selectbox(
                    "Select Month 2",
                    month_list,
                    index=index_month2,
                    key="month2_tab1"
                )

                number2 = st.number_input(
                    "Enter Year 2",
                    min_value=int(get_min_year()),
                    max_value=current_year,
                    value=current_year,
                    step=1,
                    key="year2_tab1"
                )

                if sheet_name:
                    if st.button("Generate Similar Clients", key="btn_generate_tab1"):
                        with st.spinner(
                                f"Generating Similarity Report for {selected_month_1} & {selected_month_2} for {choice}..."):
                            data1, data2, data3 = get_names_in_both_months(
                                sheet_name, selected_month_1, number1,
                                selected_month_2, number2
                            )

                            if not data1:
                                st.info("No similarities found")
                            else:
                                st.metric(label="Total Number of Same Clients", value=data3)
                                st.write("Names:")
                                st.json(data1, expanded=True)
                                st.write("Detailed Names:")
                                st.json(data2, expanded=False)

            with tab2:
                choice = st.selectbox(
                    "Select Data To View",
                    ["USA", "UK"],
                    index=None,
                    key="choice_tab2"
                )

                sheet_name = {"UK": sheet_uk, "USA": sheet_usa}.get(choice)

                number3 = st.number_input(
                    "Enter Year",
                    min_value=int(get_min_year()),
                    max_value=current_year,
                    value=current_year,
                    step=1,
                    key="year_tab2"
                )

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
                choice = st.selectbox(
                    "Select Data To View",
                    ["USA", "UK"],
                    index=None,
                    key="choice_tab3"
                )
                sheet_name = {"UK": sheet_uk, "USA": sheet_usa}.get(choice)

                number1 = st.number_input(
                    "Enter Year 1",
                    min_value=int(get_min_year()),
                    max_value=current_year,
                    value=current_year-1,
                    step=1,
                    key="year1_tab3"
                )

                number2 = st.number_input(
                    "Enter Year 2",
                    min_value=int(get_min_year()),
                    max_value=current_year,
                    value=current_year,
                    step=1,
                    key="year2_tab3"
                )

                if sheet_name:
                    if st.button("Generate Similar Clients", key="btn_generate_tab3"):
                        with st.spinner(
                                f"Generating Similarity Report for {number1} & {number2} for {choice}..."):
                            data1, data2, data3 = get_names_in_both_years(
                                sheet_name, number1,
                                number2
                            )

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
                                            st.markdown(
                                                "\n".join([f"- {d}" for d in data["publishing_dates"]])
                                            )

        elif action == "Summary":
            st.header("📄 Generate Summary Report")
            selected_month = st.selectbox(
                "Select Month",
                month_list,
                index=current_month - 1,
                placeholder="Select Month"
            )
            number = st.number_input("Enter Year", min_value=int(get_min_year()), max_value=current_year,
                                     value=current_year, step=1)
            selected_month_number = month_list.index(selected_month) + 1 if selected_month else None
            uk_clean = clean_data_reviews(sheet_uk)
            usa_clean = clean_data_reviews(sheet_usa)

            usa_clean = usa_clean[
                (usa_clean["Publishing Date"].dt.month == selected_month_number) &
                (usa_clean["Publishing Date"].dt.year == number)
                ]
            uk_clean = uk_clean[
                (uk_clean["Publishing Date"].dt.month == selected_month_number) &
                (uk_clean["Publishing Date"].dt.year == number)
                ]
            no_data = False

            if usa_clean.empty:
                no_data = True

            if uk_clean.empty:
                no_data = True
            if usa_clean.empty and uk_clean.empty:
                no_data = True

            if no_data:
                st.error(f"Cannot generate summary — no data available for the month {selected_month} {number}.")
            else:
                if st.button("Generate Summary"):
                    with st.spinner(f"Generating Summary Report for {selected_month} {number}..."):
                        usa_review_data, uk_review_data, usa_brands, uk_brands, usa_platforms, uk_platforms, printing_stats, copyright_stats, a_plus, total_unique_clients, combined, attained_reviews_per_pm, attained_df, pending_sent_details, negative_reviews_per_pm, negative_details, Issues_usa, Issues_uk = summary(
                            selected_month_number, number)
                        pdf_data, pdf_filename = generate_summary_report_pdf(usa_review_data, uk_review_data,
                                                                             usa_brands, uk_brands,
                                                                             usa_platforms, uk_platforms,
                                                                             printing_stats, copyright_stats,
                                                                             a_plus,
                                                                             selected_month=selected_month, start_year=number)

                        usa_total = sum(usa_review_data.values())
                        usa_attained = usa_review_data["Attained"] if "Attained" in usa_review_data else 0

                        usa_attained_pct = (usa_attained / usa_total * 100) if usa_total > 0 else 0

                        uk_total = sum(uk_review_data.values())
                        uk_attained = uk_review_data["Attained"] if "Attained" in uk_review_data else 0

                        uk_attained_pct = (uk_attained / uk_total * 100) if uk_total > 0 else 0

                        combined_total = usa_total + uk_total
                        combined_attained = usa_attained + uk_attained
                        combined_attained_pct = (
                                combined_attained / combined_total * 100) if combined_total > 0 else 0

                        st.header(f"{selected_month} {number} Summary Report")

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
                            unique_clients_count_per_pm = combined.groupby('Project Manager')[
                                'Name'].nunique().reset_index()
                            unique_clients_count_per_pm.columns = ['Project Manager', 'Unique Clients']
                            unique_clients_count_per_pm.index = range(1, len(unique_clients_count_per_pm) + 1)
                            clients_list = combined.groupby('Project Manager')["Name"].apply(list).reset_index(
                                name="Clients")
                            merged_df = unique_clients_count_per_pm.merge(clients_list, on='Project Manager',
                                                                          how='left')
                            merged_df.index = range(1, len(merged_df) + 1)
                            with st.expander("🤵🏻 Total Clients"):
                                st.dataframe(combined)
                            buffer = io.BytesIO()
                            combined.to_excel(buffer, index=False)
                            buffer.seek(0)

                            st.download_button(
                                label="📥 Download Excel",
                                data=buffer,
                                file_name=f"USA+UK_{selected_month}_{number}.xlsx",
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                help="Click to download the Excel report"
                            )
                        with col2:
                            uk_pie = create_review_pie_chart(uk_review_data, "UK Trustpilot Reviews")
                            if uk_pie:
                                st.plotly_chart(uk_pie, width="stretch", key="uk_pie")
                            st.subheader("🇬🇧 UK Reviews")
                            st.metric("📊 Total Reviews", uk_total)
                            st.metric("🟢 Total Attained", uk_attained)
                            st.metric("🔴 Total Negative", uk_review_data.get("Negative", 0))
                            st.metric("🎯Attained Percentage", f"{uk_attained_pct:.1f}%")
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
                            total_count_usa = usa_df["Count"].sum()
                            st.markdown(f"""
                                            - 📊 **Total Count Across Brands:** `{total_count_usa}`
                                            """)

                            st.subheader("USA Platform Breakdown")
                            usa_platform_df = pd.DataFrame(list(usa_platforms.items()),
                                                           columns=['Platform', 'Count'])
                            st.dataframe(usa_platform_df, hide_index=True)
                            total_count_usa_platforms = usa_platform_df["Count"].sum()
                            st.markdown(f"""
                                            - 📊 **Total Count Across Platforms:** `{total_count_usa_platforms}`
                                            """)
                        with col2:
                            st.subheader("UK Brand Breakdown")
                            uk_df = pd.DataFrame(list(uk_brands.items()), columns=['Brand', 'Count'])
                            st.dataframe(uk_df, hide_index=True)
                            total_count_uk = uk_df["Count"].sum()
                            st.markdown(f"""
                                            - 📊 **Total Count Across Brands:** `{total_count_uk}`
                                            """)
                            st.subheader("UK Platform Breakdown")
                            uk_platform_df = pd.DataFrame(list(uk_platforms.items()), columns=['Platform', 'Count'])
                            st.dataframe(uk_platform_df, hide_index=True)
                            total_count_uk_platforms = uk_platform_df["Count"].sum()
                            st.markdown(f"""
                                            - 📊 **Total Count Across Platforms:** `{total_count_uk_platforms}`
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

                        st.markdown('<h2 class="section-header">©️ Copyright Analytics</h2>',
                                    unsafe_allow_html=True)

                        col1, col2 = st.columns(2)

                        with col1:
                            st.subheader("📋 Copyright Summary")
                            st.metric("Total Copyrights", copyright_stats['Total_copyrights'])
                            st.metric("Total Cost", f"${copyright_stats['Total_cost_copyright']:,}")
                            st.metric("Success Rate",
                                      f"{copyright_stats['result_count']}/{copyright_stats['Total_copyrights']}")

                            success_rate = (
                                    copyright_stats['result_count'] / copyright_stats['Total_copyrights'] * 100) if \
                                copyright_stats['Total_copyrights'] > 0 else 0
                            st.metric("Success Percentage", f"{success_rate:.1f}%")
                            st.metric("Rejection Rate",
                                      f"{copyright_stats['result_count_no']}/{copyright_stats['Total_copyrights']}")

                            rejection_rate = (
                                    copyright_stats['result_count_no'] / copyright_stats[
                                'Total_copyrights'] * 100) if copyright_stats['Total_copyrights'] > 0 else 0
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

                            with cp2:
                                st.metric('UK', copyright_stats['uk'])

                        st.divider()

                        cola = st.columns(1)

                        with cola[0]:
                            st.subheader("🅰➕ Content")
                            st.metric("A+ Count", f"{a_plus} Published")

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
                    st.success(f"Summary report for {selected_month} {number} generated!")
                    st.download_button(
                        label="📥 Download PDF Report",
                        data=pdf_data,
                        file_name=pdf_filename,
                        mime="application/pdf",
                        help="Click to download the PDF report"
                    )
        elif action == "Year Summary" and number:

            st.header("📄 Generate Year Summary Report")

            uk_clean = clean_data_reviews(sheet_uk)
            usa_clean = clean_data_reviews(sheet_usa)

            usa_clean = usa_clean[
                (usa_clean["Publishing Date"].dt.year == number)
            ]
            uk_clean = uk_clean[
                (uk_clean["Publishing Date"].dt.year == number)
            ]
            no_data = False

            if usa_clean.empty:
                no_data = True

            if uk_clean.empty:
                no_data = True

            if usa_clean.empty and uk_clean.empty:
                no_data = True

            if no_data:
                st.error(f"Cannot generate summary — no data available for the Year {number}.")
            else:
                if st.button("Generate Year Summary Report"):
                    with st.spinner("Generating Year Summary Report"):
                        usa_review_data, uk_review_data, usa_brands, uk_brands, usa_platforms, uk_platforms, printing_stats, monthly_printing, copyright_stats, a_plus, total_unique_clients, combined, attained_reviews_per_pm, attained_df, attained_reviews_per_month, pending_sent_details, negative_reviews_per_pm, negative_details, negative_per_month, publishing_per_month, Issues_usa, Issues_uk = generate_year_summary(
                            number)
                        pdf_data, pdf_filename = generate_summary_report_pdf(usa_review_data, uk_review_data,
                                                                             usa_brands, uk_brands,
                                                                             usa_platforms, uk_platforms,
                                                                             printing_stats, copyright_stats, a_plus,
                                                                             start_year=number)

                        usa_total = sum(usa_review_data.values())
                        usa_attained = usa_review_data["Attained"] if "Attained" in usa_review_data else 0

                        usa_attained_pct = (usa_attained / usa_total * 100) if usa_total > 0 else 0

                        uk_total = sum(uk_review_data.values())
                        uk_attained = uk_review_data["Attained"] if "Attained" in uk_review_data else 0

                        uk_attained_pct = (uk_attained / uk_total * 100) if uk_total > 0 else 0

                        combined_total = usa_total + uk_total
                        combined_attained = usa_attained + uk_attained
                        combined_attained_pct = (combined_attained / combined_total * 100) if combined_total > 0 else 0

                        st.header(f"{number} Summary Report")
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
                            unique_clients_count_per_pm = combined.groupby('Project Manager')[
                                'Name'].nunique().reset_index()
                            unique_clients_count_per_pm.columns = ['Project Manager', 'Unique Clients']
                            unique_clients_count_per_pm.index = range(1, len(unique_clients_count_per_pm) + 1)
                            clients_list = combined.groupby('Project Manager')["Name"].apply(list).reset_index(
                                name="Clients")
                            merged_df = unique_clients_count_per_pm.merge(clients_list, on='Project Manager',
                                                                          how='left')
                            merged_df.index = range(1, len(merged_df) + 1)

                            with st.expander("🤵🏻 Total Clients"):
                                st.dataframe(combined)
                            with st.expander("🤵🏻🤵🏻 Publishing Per Month"):
                                st.dataframe(publishing_per_month)
                            buffer = io.BytesIO()
                            combined.to_excel(buffer, index=False)
                            buffer.seek(0)

                            st.download_button(
                                label="📥 Download Excel",
                                data=buffer,
                                file_name=f"USA+UK_{number}.xlsx",
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                help="Click to download the Excel report"
                            )
                            with st.expander("🟢 Attained Reviews Per Month"):
                                st.dataframe(attained_reviews_per_month)

                            with st.expander("🔴 Negative Reviews Per Month"):
                                st.dataframe(negative_per_month)
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
                            total_count_usa = usa_df["Count"].sum()
                            st.markdown(f"""
                                                        - 📊 **Total Count Across Brands:** `{total_count_usa}`
                                                        """)

                            st.subheader("USA Platform Breakdown")
                            usa_platform_df = pd.DataFrame(list(usa_platforms.items()),
                                                           columns=['Platform', 'Count'])
                            st.dataframe(usa_platform_df, hide_index=True)
                            total_count_usa_platforms = usa_platform_df["Count"].sum()
                            st.markdown(f"""
                                                        - 📊 **Total Count Across Platforms:** `{total_count_usa_platforms}`
                                                        """)
                        with col2:
                            st.subheader("UK Brand Breakdown")
                            uk_df = pd.DataFrame(list(uk_brands.items()), columns=['Brand', 'Count'])
                            st.dataframe(uk_df, hide_index=True)
                            total_count_uk = uk_df["Count"].sum()
                            st.markdown(f"""
                                                        - 📊 **Total Count Across Brands:** `{total_count_uk}`
                                                        """)
                            st.subheader("UK Platform Breakdown")
                            uk_platform_df = pd.DataFrame(list(uk_platforms.items()), columns=['Platform', 'Count'])
                            st.dataframe(uk_platform_df, hide_index=True)
                            total_count_uk_platforms = uk_platform_df["Count"].sum()
                            st.markdown(f"""
                                                        - 📊 **Total Count Across Platforms:** `{total_count_uk_platforms}`
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
                            st.metric("Success Rate",
                                      f"{copyright_stats['result_count']}/{copyright_stats['Total_copyrights']}")

                            success_rate = (
                                    copyright_stats['result_count'] / copyright_stats['Total_copyrights'] * 100) if \
                                copyright_stats['Total_copyrights'] > 0 else 0
                            st.metric("Success Percentage", f"{success_rate:.1f}%")
                            st.metric("Rejection Rate",
                                      f"{copyright_stats['result_count_no']}/{copyright_stats['Total_copyrights']}")

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

                            with cp2:
                                st.metric('UK', copyright_stats['uk'])

                        st.divider()

                        cola = st.columns(1)

                        with cola[0]:
                            st.subheader("🅰➕ Content")
                            st.metric("A+ Count", f"{a_plus} Published")

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

                    st.success(f"Summary report for {number} generated!")

                    st.download_button(
                        label="📥 Download PDF Report",
                        data=pdf_data,
                        file_name=pdf_filename,
                        mime="application/pdf",
                        help="Click to download the PDF report"
                    )

        elif action == "Sales":
            tab1, tab2 = st.tabs(["Monthly", "Yearly"])

            with tab1:
                selected_month = st.selectbox(
                    "Select Month",
                    month_list,
                    index=current_month - 1,
                    placeholder="Select Month"
                )

                number = st.number_input(
                    "Enter Year",
                    min_value=int(get_min_year()),
                    max_value=current_year,
                    value=current_year,
                    step=1
                )

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
                year = st.number_input(
                    "Enter Year",
                    min_value=int(get_min_year()),
                    max_value=current_year,
                    value=current_year,
                    step=1,
                    key="sales_year"
                )

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

            usa_clean = usa_clean[
                (usa_clean["Publishing Date"].dt.year == number)
            ]
            uk_clean = uk_clean[
                (uk_clean["Publishing Date"].dt.year == number)
            ]

            usa_clean_platforms = usa_clean[
                (usa_clean["Publishing Date"].dt.year == number)
            ]
            uk_clean_platforms = uk_clean[
                (uk_clean["Publishing Date"].dt.year == number)
            ]

            if usa_clean.empty:
                print("No values found in USA sheet.")
            if uk_clean.empty:
                print("No values found in UK sheet.")
                return
            if usa_clean.empty and uk_clean.empty:
                return

            usa_clean = usa_clean.drop_duplicates(subset=["Name"], keep="last")
            uk_clean = uk_clean.drop_duplicates(subset=["Name"], keep="last")
            total_usa = usa_clean["Name"].nunique()
            total_uk = uk_clean["Name"].nunique()
            pm_list_usa = list(set((usa_clean["Project Manager"].dropna().unique().tolist() + ["Unknown"])))
            pm_list_uk = list(set((uk_clean["Project Manager"].dropna().unique().tolist() + ["Unknown"])))
            usa_reviews_per_pm = safe_concat(
                [load_reviews_year(sheet_usa, number, pm, "Attained") for pm in pm_list_usa])
            uk_reviews_per_pm = safe_concat([load_reviews_year(sheet_uk, number, pm, "Attained") for pm in pm_list_uk])
            combined_data = safe_concat([usa_reviews_per_pm, uk_reviews_per_pm])

            if not combined_data.empty:
                combined_data["Trustpilot Review Date"] = pd.to_datetime(
                    combined_data["Trustpilot Review Date"], format="%d-%B-%Y", errors="coerce"
                )

                combined_data["Month-Year"] = combined_data["Trustpilot Review Date"].dt.to_period("M").astype(str)

                monthly_counts = (
                    combined_data.groupby(["Project Manager", "Month-Year"])
                    .size()
                    .reset_index(name="Review Count")
                )

                monthly_clients = (
                    combined_data.groupby(["Project Manager", "Month-Year"])["Name"]
                    .apply(list)
                    .reset_index(name="Clients")
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
                    index="Project Manager",
                    columns="Month-Year",
                    values="Review Count",
                    fill_value=0
                )

                monthly_pivot = monthly_pivot.reindex(
                    sorted(monthly_pivot.columns, key=lambda x: pd.to_datetime(x)),
                    axis=1
                )
                monthly_pivot.columns = [
                    pd.to_datetime(col).strftime("%B %Y") for col in monthly_pivot.columns
                ]

                with st.expander("📊 Monthly Review Count Pivot Table"):
                    st.dataframe(monthly_pivot, width="stretch")

            else:
                st.warning("No combined review data found.")

        elif action == "Custom Summary":

            start_year = st.number_input("Enter Year", min_value=int(get_min_year()), max_value=current_year,
                                     value=current_year-1, step=1, key="start_year")
            end_year = st.number_input("Enter Year", min_value=int(get_min_year()), max_value=current_year,
                                     value=current_year, step=1, key="end_year")
            st.header("📄 Generate Multi Year Summary Report")

            uk_clean = clean_data_reviews(sheet_uk)
            usa_clean = clean_data_reviews(sheet_usa)

            usa_clean = usa_clean[
                (usa_clean["Publishing Date"].dt.year >= start_year) &
                 (usa_clean["Publishing Date"].dt.year<=  end_year)
            ]
            uk_clean = uk_clean[
                (uk_clean["Publishing Date"].dt.year >= start_year) &
                (uk_clean["Publishing Date"].dt.year <= end_year)
            ]
            no_data = False

            if usa_clean.empty:
                no_data = True

            if uk_clean.empty:
                no_data = True

            if usa_clean.empty and uk_clean.empty:
                no_data = True

            if no_data:
                st.error(f"Cannot generate summary — no data available for the Years {start_year}-{end_year}.")
            else:
                if st.button("Generate Year Summary Report"):
                    with st.spinner("Generating Year Summary Report"):
                        usa_review_data, uk_review_data, usa_brands, uk_brands, usa_platforms, uk_platforms, printing_stats, monthly_printing, copyright_stats, a_plus, total_unique_clients, combined, attained_reviews_per_pm, attained_df, attained_reviews_per_month, pending_sent_details, negative_reviews_per_pm, negative_details, negative_per_month, publishing_per_month, Issues_usa, Issues_uk = generate_year_summary_multiple(
                            start_year, end_year)
                        pdf_data, pdf_filename = generate_summary_report_pdf(usa_review_data, uk_review_data,
                                                                             usa_brands, uk_brands,
                                                                             usa_platforms, uk_platforms,
                                                                             printing_stats, copyright_stats, a_plus,
                                                                             selected_month, start_year=start_year, end_year=end_year)

                        usa_total = sum(usa_review_data.values())
                        usa_attained = usa_review_data["Attained"] if "Attained" in usa_review_data else 0

                        usa_attained_pct = (usa_attained / usa_total * 100) if usa_total > 0 else 0

                        uk_total = sum(uk_review_data.values())
                        uk_attained = uk_review_data["Attained"] if "Attained" in uk_review_data else 0

                        uk_attained_pct = (uk_attained / uk_total * 100) if uk_total > 0 else 0

                        combined_total = usa_total + uk_total
                        combined_attained = usa_attained + uk_attained
                        combined_attained_pct = (combined_attained / combined_total * 100) if combined_total > 0 else 0

                        st.header(f"{start_year}-{end_year} Summary Report")
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
                            unique_clients_count_per_pm = combined.groupby('Project Manager')[
                                'Name'].nunique().reset_index()
                            unique_clients_count_per_pm.columns = ['Project Manager', 'Unique Clients']
                            unique_clients_count_per_pm.index = range(1, len(unique_clients_count_per_pm) + 1)
                            clients_list = combined.groupby('Project Manager')["Name"].apply(list).reset_index(
                                name="Clients")
                            merged_df = unique_clients_count_per_pm.merge(clients_list, on='Project Manager',
                                                                          how='left')
                            merged_df.index = range(1, len(merged_df) + 1)

                            with st.expander("🤵🏻 Total Clients"):
                                st.dataframe(combined)
                            with st.expander("🤵🏻🤵🏻 Publishing Per Month"):
                                st.dataframe(publishing_per_month)
                            buffer = io.BytesIO()
                            combined.to_excel(buffer, index=False)
                            buffer.seek(0)

                            st.download_button(
                                label="📥 Download Excel",
                                data=buffer,
                                file_name=f"USA+UK_{number}.xlsx",
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                help="Click to download the Excel report"
                            )
                            with st.expander("🟢 Attained Reviews Per Month"):
                                st.dataframe(attained_reviews_per_month)

                            with st.expander("🔴 Negative Reviews Per Month"):
                                st.dataframe(negative_per_month)
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
                            total_count_usa = usa_df["Count"].sum()
                            st.markdown(f"""
                                                        - 📊 **Total Count Across Brands:** `{total_count_usa}`
                                                        """)

                            st.subheader("USA Platform Breakdown")
                            usa_platform_df = pd.DataFrame(list(usa_platforms.items()),
                                                           columns=['Platform', 'Count'])
                            st.dataframe(usa_platform_df, hide_index=True)
                            total_count_usa_platforms = usa_platform_df["Count"].sum()
                            st.markdown(f"""
                                                        - 📊 **Total Count Across Platforms:** `{total_count_usa_platforms}`
                                                        """)
                        with col2:
                            st.subheader("UK Brand Breakdown")
                            uk_df = pd.DataFrame(list(uk_brands.items()), columns=['Brand', 'Count'])
                            st.dataframe(uk_df, hide_index=True)
                            total_count_uk = uk_df["Count"].sum()
                            st.markdown(f"""
                                                        - 📊 **Total Count Across Brands:** `{total_count_uk}`
                                                        """)
                            st.subheader("UK Platform Breakdown")
                            uk_platform_df = pd.DataFrame(list(uk_platforms.items()), columns=['Platform', 'Count'])
                            st.dataframe(uk_platform_df, hide_index=True)
                            total_count_uk_platforms = uk_platform_df["Count"].sum()
                            st.markdown(f"""
                                                        - 📊 **Total Count Across Platforms:** `{total_count_uk_platforms}`
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
                            st.metric("Success Rate",
                                      f"{copyright_stats['result_count']}/{copyright_stats['Total_copyrights']}")

                            success_rate = (
                                    copyright_stats['result_count'] / copyright_stats['Total_copyrights'] * 100) if \
                                copyright_stats['Total_copyrights'] > 0 else 0
                            st.metric("Success Percentage", f"{success_rate:.1f}%")
                            st.metric("Rejection Rate",
                                      f"{copyright_stats['result_count_no']}/{copyright_stats['Total_copyrights']}")

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

                            with cp2:
                                st.metric('UK', copyright_stats['uk'])

                        st.divider()

                        cola = st.columns(1)

                        with cola[0]:
                            st.subheader("🅰➕ Content")
                            st.metric("A+ Count", f"{a_plus} Published")

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

                    st.success(f"Summary report for {number} generated!")

                    st.download_button(
                        label="📥 Download PDF Report",
                        data=pdf_data,
                        file_name=pdf_filename,
                        mime="application/pdf",
                        help="Click to download the PDF report"
                    )


if __name__ == '__main__':
    main()
