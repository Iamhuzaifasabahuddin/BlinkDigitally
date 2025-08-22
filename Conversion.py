import logging
import os
from pprint import pprint

import gspread
import pandas as pd
from dotenv import load_dotenv
from google.oauth2.service_account import Credentials



load_dotenv('Info.env')
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

def clean_data_to_dict(sheet_name: str) -> dict[str, str]:
    """Clean the data from Google Sheets using gspread"""
    data = get_sheet_data(sheet_name)

    if data.empty:
        logging.warning(f"No data found in sheet: {sheet_name}")
        return data

    columns = list(data.columns)
    if "Email" in columns:
        end_col_index = columns.index("Email")
        data = data.iloc[:, :end_col_index + 1]

    data.index = range(1, len(data) + 1)

    name_email_dict = dict(zip(data['Name'], data['Email']))
    return name_email_dict


pprint(clean_data_to_dict("QA"))

