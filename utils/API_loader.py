import streamlit as st
import gspread
from google.oauth2.service_account import Credentials
import logging
import pandas as pd

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