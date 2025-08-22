import logging
import os

import gspread
import pandas as pd
from dotenv import load_dotenv
from google.oauth2.service_account import Credentials
from slack_sdk import WebClient

load_dotenv('Info.env')
SLACK_BOT_TOKEN = os.getenv("SLACK")
client = WebClient(token=SLACK_BOT_TOKEN)

SPREADSHEET_ID = os.getenv("SPREADSHEET_ID")

SCOPES = ["https://www.googleapis.com/auth/spreadsheets.readonly"]
# Initialize gspread client
def get_gspread_client():
    """Initialize and return gspread client with service account credentials"""
    try:
        credentials = Credentials.from_service_account_file("credentials.json", scopes=SCOPES)
        gc = gspread.authorize(credentials)
        return gc
    except Exception as e:
        logging.error(f"Failed to initialize gspread client: {e}")
        return None


def get_sheet_data(sheet_name):
    """Get data from a specific sheet using gspread"""
    try:
        gc = get_gspread_client()
        if not gc:
            return pd.DataFrame()

        spreadsheet = gc.open_by_key(SPREADSHEET_ID)
        worksheet = spreadsheet.worksheet(sheet_name)

        records = worksheet.get_all_values()
        headers = records[0]
        rows = records[1:]
        df = pd.DataFrame(rows, columns=headers)

        return df
    except Exception as e:
        logging.error(f"Failed to get data from sheet {sheet_name}: {e}")
        return pd.DataFrame()


# Sheet names
sheet_usa = "USA"
sheet_uk = "UK"
sheet_audio = "AudioBook"
sheet_printing = "Printing"
sheet_copyright = "Copyright"

def clean_data_reviews(sheet_name: str) -> pd.DataFrame:
    """Clean the data from Google Sheets using gspread"""
    data = get_sheet_data(sheet_name)

    if data.empty:
        logging.warning(f"No data found in sheet: {sheet_name}")
        return data

    columns = list(data.columns)
    if "Issues" in columns:
        end_col_index = columns.index("Issues")
        data = data.iloc[:, :end_col_index + 1]

    # Handle date columns
    for col in ["Publishing Date", "Last Edit (Revision)", "Trustpilot Review Date"]:
        if col in data.columns:
            data[col] = pd.to_datetime(data[col], errors="coerce")

    data = data.sort_values(by="Publishing Date", ascending=True)
    data.index = range(1, len(data) + 1)

    return data


def load_data_reviews(sheet_name, name) -> tuple:
    """Load and filter data for a specific project manager"""
    data = clean_data_reviews(sheet_name)

    if "Name" in data.columns:
        data = data.drop_duplicates(subset=["Name"])

    data_original = data
    if data.empty:
        return pd.DataFrame(), 0, pd.NaT, pd.NaT, 0, 0

    # Filter data based on criteria
    data_count = data_original[
        (data_original["Project Manager"] == name) &
        ((data_original["Trustpilot Review"] == "Pending") | (data_original["Trustpilot Review"] == "Sent")) &
        (data_original["Brand"].isin(["BookMarketeers", "Writers Clique", "Authors Solution"])) &
        (data_original["Status"] == "Published")
        ]



    data = data_original[
        (data_original["Project Manager"] == name) &
        # ((data_original["Trustpilot Review"] == "Pending") | (data_original["Trustpilot Review"] == "Sent")) &
        ((data_original["Trustpilot Review"] == "Pending")) &
        (data_original["Brand"].isin(["BookMarketeers", "Writers Clique", "Authors Solution"])) &
        (data_original["Status"] == "Published")
        ]

    data = data.sort_values(by="Publishing Date", ascending=True)

    # Clean strings and drop missing
    data_original["Trustpilot Review"] = data_original["Trustpilot Review"].astype(str).str.strip().str.lower()
    data_original["Project Manager"] = data_original["Project Manager"].astype(str).str.strip()
    name = name.strip()

    # Drop rows where essential fields are missing
    data_original = data_original.dropna(subset=["Trustpilot Review", "Project Manager"])

    # Now calculate attained count
    attained = len(
        data_original[
            (data_original["Trustpilot Review"] == "attained") &
            (data_original["Project Manager"] == name)
            ]
    )

    # Calculate statistics
    total_percentage = 0
    # attained = len(
    #     data_original[(data_original["Trustpilot Review"] == "Attained") & (data_original["Project Manager"] == name)]
    # )

    total_reviews = len(data_count) + attained

    min_date = data["Publishing Date"].min() if not data.empty else pd.NaT
    max_date = data["Publishing Date"].max() if not data.empty else pd.NaT

    if total_reviews > 0:
        total_percentage = (attained / total_reviews)

    data.index = range(1, len(data) + 1)
    return data

def get_printing_data_reviews(month, year) -> pd.DataFrame:
    """Get printing data for the current month using gspread"""
    data = get_sheet_data(sheet_printing)

    if data.empty:
        return data

    columns = list(data.columns)
    if "Fulfilled" in columns:
        end_col_index = columns.index("Fulfilled")
        data = data.iloc[:, :end_col_index + 1]
        data = data.astype(str)

    for col in ["Order Date", "Shipping Date", "Fulfilled"]:
        if col in data.columns:
            data[col] = pd.to_datetime(data[col], errors="coerce")

    data = data[(data["Order Date"].dt.month == month) & (data["Order Date"].dt.year == year)]

    if "Order Cost" in data.columns:
        data["Order Cost"] = data["Order Cost"].astype(str)
        data['Order Cost'] = pd.to_numeric(data['Order Cost'].str.replace('$', '', regex=False), errors='coerce')

    if "No of Copies" in data.columns:
        data["No of Copies"] = pd.to_numeric(data["No of Copies"], errors='coerce')

    data = data.sort_values(by="Order Date", ascending=True)

    for col in ["Order Date", "Shipping Date", "Fulfilled"]:
        if col in data.columns:
            data[col] = pd.to_datetime(data[col], errors="coerce").dt.strftime("%d-%B-%Y")

    data.index = range(1, len(data) + 1)
    data = data.fillna("N/A")

    return data


def get_names_in_both_months(sheet_name: str) -> tuple:
    """
    Identifies names that appear in both June and July from a Google Sheet.
    Returns:
        - A set of matching names
        - A dictionary with individual counts for June and July
    """
    df = get_sheet_data(sheet_name)

    if df.empty or "Name" not in df.columns or "Publishing Date" not in df.columns:
        logging.warning("Missing 'Name' or 'Date' columns or data is empty.")
        return set(), {}

    # Parse date column
    df['Publishing Date'] = pd.to_datetime(df['Publishing Date'], errors='coerce')
    df = df.dropna(subset=['Publishing Date', 'Name'])

    # Extract month name
    df['Month'] = df['Publishing Date'].dt.month_name()

    # Get June and July name sets
    june_names = set(df[df['Month'] == 'June']['Name'].str.strip())
    july_names = set(df[df['Month'] == 'July']['Name'].str.strip())

    # Find intersection
    names_in_both = june_names.intersection(july_names)

    # Optional: Count how many times each name appears in June and July
    counts = {}
    for name in names_in_both:
        june_count = df[(df['Month'] == 'June') & (df['Name'].str.strip() == name)].shape[0]
        july_count = df[(df['Month'] == 'July') & (df['Name'].str.strip() == name)].shape[0]
        counts[name] = {
            "June": june_count,
            "July": july_count,
        }

    return names_in_both, counts, len(names_in_both)


names, counts, total = get_names_in_both_months(sheet_usa)
print(names)
print(counts)
print(total)