import streamlit as st
import pandas as pd
from API_loader import get_sheet_data


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

def safe_concat(dfs):
    dfs = [df for df in dfs if not df.empty]
    return pd.concat(dfs, ignore_index=True) if dfs else pd.DataFrame()