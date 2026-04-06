from data_cleaner import get_sheet_data, clean_data, get_min_year
import pandas as pd
import streamlit as st
import logging
from datetime import datetime

sheet_usa = "USA"
sheet_uk = "UK"
sheet_audio = "AudioBook"
sheet_printing = "Printing"
sheet_copyright = "Copyright"
sheet_a_plus = "A_plus"
sheet_sales = "Sales"

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