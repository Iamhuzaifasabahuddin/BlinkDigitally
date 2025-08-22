import calendar
import logging
import os
import shutil
from datetime import datetime

import gspread
import pandas as pd
from dotenv import load_dotenv
from google.oauth2.service_account import Credentials

load_dotenv('Info.env')
base_path_uk = r"C:\Users\Huzaifa Sabah Uddin\Monthly\UK"
base_path_usa = r"C:\Users\Huzaifa Sabah Uddin\Monthly\USA"
base_path_printing = r"C:\Users\Huzaifa Sabah Uddin\Monthly\Printing"

# current_month = datetime.today().month
current_month = 4
current_month_name = calendar.month_name[current_month]
current_year = datetime.today().year

sheet_usa = "USA"
sheet_uk = "UK"
sheet_audio = "AudioBook"
sheet_printing = "Printing"
sheet_copyright = "Copyright"
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

        # Get all records as a list of dictionaries
        records = worksheet.get_all_values()
        headers = records[0]
        rows = records[1:]
        # Convert to DataFrame
        df = pd.DataFrame(rows, columns=headers)

        return df
    except Exception as e:
        logging.error(f"Failed to get data from sheet {sheet_name}: {e}")
        return pd.DataFrame()


def clean_data(url: str) -> pd.DataFrame:
    data = get_sheet_data(url)
    columns = list(data.columns)
    end_col_index = columns.index("Issues")
    data = data.iloc[:, :end_col_index + 1]
    return data


def load_data(url):
    data = clean_data(url)
    data["Publishing Date"] = pd.to_datetime(data["Publishing Date"], errors='coerce')
    data = data[data["Publishing Date"].dt.month == current_month]

    data = data.sort_values(by=["Publishing Date"], ascending=True)
    data.index = range(1, len(data) + 1)

    data["Publishing Date"] = data["Publishing Date"].dt.strftime("%m-%B-%Y")

    return data


def printing():
    data = get_sheet_data(sheet_printing)
    columns = list(data.columns)
    end_col_index = columns.index("Fulfilled")
    data = data.iloc[:, :end_col_index + 1]

    data["Order Date"] = pd.to_datetime(data["Order Date"], errors='coerce')

    data = data[data["Order Date"].dt.month == current_month]
    data['Order Cost'] = pd.to_numeric(data['Order Cost'].str.replace('$', '', regex=False))

    data = data.sort_values(by=["Order Date"], ascending=True)
    data.index = range(1, len(data) + 1)

    data["Order Date"] = data["Order Date"].dt.strftime("%m-%B-%Y")

    return data


def move_file_safely(src_path, dest_dir):
    base_name = os.path.basename(src_path)
    name, ext = os.path.splitext(base_name)
    dest_path = os.path.join(dest_dir, base_name)
    counter = 1

    while os.path.exists(dest_path):
        dest_path = os.path.join(dest_dir, f"{name} ({counter}){ext}")
        counter += 1

    shutil.move(src_path, dest_path)
    print(f"Moved to: {dest_path}")


def generate_monthly():
    data_uk = load_data(sheet_uk)
    data_usa = load_data(sheet_usa)

    uk_filename = f"{current_month_name}-{current_year} UK.xlsx"
    usa_filename = f"{current_month_name}-{current_year} USA.xlsx"

    data_uk.to_excel(uk_filename, index=False)
    data_usa.to_excel(usa_filename, index=False)

    os.makedirs(base_path_uk, exist_ok=True)
    os.makedirs(base_path_usa, exist_ok=True)

    move_file_safely(uk_filename, base_path_uk)
    move_file_safely(usa_filename, base_path_usa)

    data_printing = printing()
    printing_filename = f"{current_month_name}-{current_year} PRINTING.xlsx"

    data_printing.to_excel(printing_filename, index=False)
    os.makedirs(base_path_printing, exist_ok=True)
    move_file_safely(printing_filename, base_path_printing)


if __name__ == '__main__':
    generate_monthly()
