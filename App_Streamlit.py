import calendar
import logging
from datetime import datetime
from io import BytesIO

import gspread
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st
from google.oauth2.service_account import Credentials
from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak
from reportlab.platypus.flowables import HRFlowable
from slack_sdk import WebClient
from slack_sdk.errors import SlackApiError

st.set_page_config(page_title="Blink Digitally", page_icon="📊", layout="centered")
client = WebClient(token=st.secrets["Slack"]["Slack"])

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
creds = Credentials.from_service_account_info(creds_dict, scopes=SCOPES)
gs_client = gspread.authorize(creds)

# Spreadsheet configuration
SPREADSHEET_ID = st.secrets["connections"]["gsheets"]["SPREADSHEET_ID"]
spreadsheet = gs_client.open_by_key(SPREADSHEET_ID)

# Sheet names
sheet_usa = "USA"
sheet_uk = "UK"
sheet_audio = "AudioBook"
sheet_printing = "Printing"
sheet_copyright = "Copyright"
sheet_a_plus = "A_plus"
sheet_sales = "Sales"

month_list = list(calendar.month_name)[1:]
current_month = datetime.today().month
# current_month = 4  # For testing specific month
current_month_name = calendar.month_name[current_month]
current_year = datetime.today().year

st.markdown("""
 <style>
    #MainMenu {visibility: hidden;}
    header {visibility: hidden;}
</style>
""", unsafe_allow_html=True)

name_usa = {
    "Aiza Ali": "aiza.ali@topsoftdigitals.pk",
    "Ahmed Asif": "ahmed.asif@topsoftdigitals.pk",
    "Ahsan Javed": "ahsan.javed@topsoftdigitals.pk",
    "Shozab Hasan": "shozab.hasan@topsoftdigitals.pk",
    "Asad Waqas": "asad.waqas@topsoftdigitals.pk",
    "Shaikh Arsalan": "shaikh.arsalan@topsoftdigitals.pk",
    "Maheen Sami": "maheen.sami@topsoftdigitals.pk",
    "Mubashir Khan": "Mubashir.khan@topsoftdigitals.pk",
    "Muhammad Ali": "muhammad.ali@topsoftdigitals.pk",
    "Valencia Angelo": "valencia.angelo@topsoftdigitals.pk"
}

names_uk = {
    "Hadia Ghazanfar": "hadia.ghazanfar@topsoftdigitals.pk",
    "Youha": "youha.khan@topsoftdigitals.pk",
    "Syed Ahsan Shahzad": "ahsan.shahzad@topsoftdigitals.pk"
}

general_message = """Hiya
:bangbang: Please ask the following Clients for their feedback about their respective projects for the ones marked as _*Pending*_ & for those marked as _*Sent*_ please remind the clients once again that their feedback truly matters and helps us grow and make essential changes to make the process even more fluid!
BM: https://bookmarketeers.com/
WC: https://writersclique.com/
AS: https://authorssolution.co.uk/"""


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
            data[col] = pd.to_datetime(data[col], errors="coerce")

    data[["Copyright", "Issues", "Last Edit (Revision)", "Trustpilot Review Date"]] = data[
        ["Copyright", "Issues", "Last Edit (Revision)", "Trustpilot Review Date"]].astype(str)

    data[["Copyright", "Issues", "Last Edit (Revision)", "Trustpilot Review Date"]] = data[
        ["Copyright", "Issues", "Last Edit (Revision)", "Trustpilot Review Date"]].fillna("N/A")

    return data


def load_data(sheet_name, month_number, year) -> pd.DataFrame:
    """Load data from Google Sheets with optional month filtering"""
    try:
        data = get_sheet_data(sheet_name)
        data = clean_data(data)

        if "Publishing Date" in data.columns:
            data = data[(data["Publishing Date"].dt.month == month_number) & (data["Publishing Date"].dt.year == year)]

        data = data.sort_values(by="Publishing Date", ascending=True)

        for col in ["Publishing Date", "Last Edit (Revision)", "Trustpilot Review Date"]:
            if col in data.columns:
                data[col] = pd.to_datetime(data[col], errors="coerce").dt.strftime("%d-%B-%Y")
        if "Name" in data.columns:
            data = data.drop_duplicates(subset=["Name"])
        data.index = range(1, len(data) + 1)
        return data

    except Exception as e:
        st.error(f"Error loading data: {e}")
        return pd.DataFrame()


def load_reviews(sheet_name, year, month_number=None) -> pd.DataFrame:
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
            data[col] = pd.to_datetime(data[col], errors="coerce")

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

        data = data.sort_values(by="Trustpilot Review Date", ascending=True)

        # for col in ["Publishing Date", "Last Edit (Revision)", "Trustpilot Review Date"]:
        #     if col in data.columns:
        #         data[col] = pd.to_datetime(data[col], errors="coerce").dt.strftime("%d-%B-%Y")
        if "Name" in data.columns:
            data = data.drop_duplicates(subset=["Name"])
        data.index = range(1, len(data) + 1)
        return data
    except Exception as e:
        st.error(f"Error loading data: {e}")
        return pd.DataFrame()


def review_data(sheet_name, month, year, status) -> pd.DataFrame:
    """Filter data by month and review status"""
    data = load_data(sheet_name, month, year)
    if not data.empty and month and status:
        if "Publishing Date" in data.columns and "Trustpilot Review" in data.columns:
            data = data[(data["Publishing Date"].dt.month == month) & (data["Publishing Date"].dt.year == year)]
            data = data[data["Trustpilot Review"] == status]
        data = data.sort_values(by="Publishing Date", ascending=True)
    data.index = range(1, len(data) + 1)
    return data


def get_printing_data(month, year) -> pd.DataFrame:
    """Get printing data filtered by month"""
    try:
        data = get_sheet_data(sheet_printing)

        columns = list(data.columns)
        if "Fulfilled" in columns:
            end_col_index = columns.index("Fulfilled")
            data = data.iloc[:, :end_col_index + 1]

        data = data.astype(str)
        for col in ["Order Date", "Shipping Date", "Fulfilled"]:
            if col in data.columns:
                data[col] = pd.to_datetime(data[col], errors="coerce")

        if month and "Order Date" in data.columns:
            data = data[(data["Order Date"].dt.month == month) & (data["Order Date"].dt.year == year)]

        if "Order Cost" in data.columns:
            data["Order Cost"] = data["Order Cost"].astype(str)
            data["Order Cost"] = pd.to_numeric(
                data["Order Cost"].str.replace("$", "", regex=False).str.replace(",", "", regex=False), errors="coerce")

        data = data.sort_values(by="Order Date", ascending=True)

        data["No of Copies"] = data["No of Copies"].astype(float)
        for col in ["Order Date", "Shipping Date", "Fulfilled"]:
            if col in data.columns:
                data[col] = pd.to_datetime(data[col], errors="coerce").dt.strftime("%d-%B-%Y")

        data.index = range(1, len(data) + 1)
        data = data.fillna("N/A")

        return data
    except Exception as e:
        st.error(f"Error loading printing data: {e}")
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
            data[col] = pd.to_datetime(data[col], errors="coerce")

    data = data.sort_values(by="Publishing Date", ascending=True)
    data.index = range(1, len(data) + 1)

    return data


def load_data_reviews(sheet_name, name) -> tuple[pd.DataFrame, float, pd.Timestamp, pd.Timestamp, int, int]:
    """Load and filter data for a specific project manager"""
    data = clean_data_reviews(sheet_name)
    data_original = data

    # Filter data based on criteria
    data = data_original[
        (data_original["Project Manager"] == name) &
        ((data_original["Trustpilot Review"] == "Pending") | (data_original["Trustpilot Review"] == "Sent")) &
        (data_original["Brand"].isin(["BookMarketeers", "Writers Clique", "Authors Solution"])) &
        (data_original["Status"] == "Published")
        ]

    data = data.sort_values(by=["Publishing Date"], ascending=True)

    # Calculate statistics
    total_percentage = 0
    attained = len(
        data_original[(data_original["Trustpilot Review"] == "Attained") & (data_original["Project Manager"] == name)]
    )
    total_reviews = len(data) + attained

    min_date = data["Publishing Date"].min() if not data.empty else pd.NaT
    max_date = data["Publishing Date"].max() if not data.empty else pd.NaT

    if total_reviews > 0:
        total_percentage = (attained / total_reviews)

    data.index = range(1, len(data) + 1)
    return data, total_percentage, min_date, max_date, attained, total_reviews


def load_data_audio(name) -> pd.DataFrame:
    """Load and filter audio book data for a specific project manager"""
    return load_data_reviews(sheet_audio, name)


def get_user_id_by_email(email):
    try:
        response = client.users_lookupByEmail(email=email)
        return response['user']['id']
    except SlackApiError as e:
        print(f"Error finding user: {e.response['error']}")
        logging.error(e)
        return None


def send_dm(user_id, message):
    try:
        response = client.chat_postMessage(
            channel=user_id,
            text=message
        )
    except SlackApiError as e:
        print(f"❌ Error sending message: {e.response['error']}")
        logging.error(e)


def send_df_as_text(name, sheet_name, email) -> None:
    """Send DataFrame as text to a user"""
    user_id = get_user_id_by_email(email)
    # user_id = get_user_id_by_email("huzaifa.sabah@topsoftdigitals.pk")

    if not user_id:
        print(f"❌ Could not find user ID for {name}")
        return

    df, percentage, min_date, max_date, attained, total_reviews = load_data_reviews(sheet_name, name)
    df_audio, percentage_audio, min_date_audio, max_date_audio, attained_audio, total_reviews_audio = load_data_audio(
        name)

    if df.empty and df_audio.empty:
        print(f"⚠️ No data for {name}")
        return
    min_month_name = min(min_date, min_date_audio).strftime("%B")
    max_month_name = max(max_date, max_date_audio).strftime("%B")

    def truncate_title(x):
        """Truncate long titles"""
        return x[:40] + "..." if isinstance(x, str) and len(x) > 40 else x

    for dframe in [df, df_audio]:
        if "Book Name & Link" in dframe.columns and not dframe.empty:
            dframe["Book Name & Link"] = dframe["Book Name & Link"].apply(truncate_title)

    display_columns = ["Name", "Brand", "Book Name & Link", "Publishing Date", "Trustpilot Review"]
    display_df = df[display_columns] if not df.empty and all(col in df.columns for col in display_columns) else df
    display_df_audio = df_audio[display_columns] if not df_audio.empty and all(
        col in df_audio.columns for col in display_columns) else df_audio

    for dframe in [display_df, display_df_audio]:
        if "Publishing Date" in dframe.columns and not dframe.empty:
            dframe["Publishing Date"] = pd.to_datetime(dframe["Publishing Date"], errors='coerce').dt.strftime(
                "%d-%B-%Y")

    merged_df = pd.concat([display_df, display_df_audio], ignore_index=True)
    display_columns = ["Name", "Brand", "Book Name & Link", "Publishing Date", "Trustpilot Review"]
    merged_df = merged_df[display_columns]
    if not merged_df.empty:
        markdown_table = merged_df.to_markdown(index=False)

        if len({min_month_name, max_month_name}) > 1:
            message = (
                f"{general_message}\n\n"
                f"Hi *{name.split()[0]}*! Here's your Trustpilot update from {min_month_name} to {max_month_name} {current_year} 📄\n\n"
                f"*Summary:* {len(merged_df)} pending reviews\n\n"
                f"*Review Retention:* {attained + attained_audio} out of {total_reviews + total_reviews_audio} "
                f"({((attained + attained_audio) / (total_reviews + total_reviews_audio)):.1%})\n\n"
                f"```\n{markdown_table}\n```"
            )
        else:
            message = (
                f"{general_message}\n\n"
                f"Hi *{name.split()[0]}*! Here's your Trustpilot update for {min_month_name} {current_year} 📄\n\n"
                f"*Summary:* {len(merged_df)} pending reviews\n\n"
                f"*Review Retention:* {attained + attained_audio} out of {total_reviews + total_reviews_audio} "
                f"({((attained + attained_audio) / (total_reviews + total_reviews_audio)):.1%})\n\n"
                f"```\n{markdown_table}\n```"
            )

        try:
            conversation = client.conversations_open(users=user_id)
            channel_id = conversation['channel']['id']

            response = client.chat_postMessage(
                channel=channel_id,
                text=message,
                mrkdwn=True
            )
            send_dm(get_user_id_by_email("huzaifa.sabah@topsoftdigitals.pk"), f"✅ Review sent to {name}")
        except SlackApiError as e:
            print(f"❌ Error sending message to {name}: {e.response['error']}")
            print(f"Detailed error: {str(e)}")
            logging.error(e)


def get_printing_data_reviews(month, year) -> pd.DataFrame:
    """Get printing data for the current month"""
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
        data["Order Cost"] = pd.to_numeric(
            data["Order Cost"].str.replace("$", "", regex=False).str.replace(",", "", regex=False), errors="coerce")

    if "No of Copies" in data.columns:
        data["No of Copies"] = pd.to_numeric(data["No of Copies"], errors='coerce')

    data = data.sort_values(by="Order Date", ascending=True)

    for col in ["Order Date", "Shipping Date", "Fulfilled"]:
        if col in data.columns:
            data[col] = pd.to_datetime(data[col], errors="coerce").dt.strftime("%d-%B-%Y")

    data.index = range(1, len(data) + 1)
    data = data.fillna("N/A")

    return data


def printing_data_all(year) -> pd.DataFrame:
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

    data = data[(data["Order Date"].dt.year == year)]

    if "Order Cost" in data.columns:
        data["Order Cost"] = data["Order Cost"].astype(str)
        data["Order Cost"] = pd.to_numeric(
            data["Order Cost"].str.replace("$", "", regex=False).str.replace(",", "", regex=False), errors="coerce")

    if "No of Copies" in data.columns:
        data["No of Copies"] = pd.to_numeric(data["No of Copies"], errors='coerce')

    data = data.sort_values(by="Order Date", ascending=True)

    for col in ["Order Date", "Shipping Date", "Fulfilled"]:
        if col in data.columns:
            data[col] = pd.to_datetime(data[col], errors="coerce").dt.strftime("%d-%B-%Y")

    data.index = range(1, len(data) + 1)
    data = data.fillna("N/A")
    return data


def get_copyright_data(month, year) -> tuple[pd.DataFrame, int, int]:
    """Get copyright data for the current month"""
    data = get_sheet_data(sheet_copyright)

    if data.empty:
        return data, 0

    columns = list(data.columns)
    if "Country" in columns:
        end_col_index = columns.index("Country")
        data = data.iloc[:, :end_col_index + 1]
    data = data.astype(str)

    if "Submission Date" in data.columns:
        data["Submission Date"] = pd.to_datetime(data["Submission Date"], errors='coerce')
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


def copyright_all(year) -> tuple[pd.DataFrame, int, int]:
    data = get_sheet_data(sheet_copyright)

    if data.empty:
        return data, 0

    columns = list(data.columns)
    if "Country" in columns:
        end_col_index = columns.index("Country")
        data = data.iloc[:, :end_col_index + 1]
    data = data.astype(str)

    if "Submission Date" in data.columns:
        data["Submission Date"] = pd.to_datetime(data["Submission Date"], errors='coerce')
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


def get_A_plus_data(month, year) -> tuple[pd.DataFrame, int]:
    data = get_sheet_data(sheet_a_plus)
    if data.empty:
        return data

    columns = list(data.columns)
    if "Issues" in columns:
        end_col_index = columns.index("Issues")
        data = data.iloc[:, :end_col_index + 1]
    data = data.astype(str)

    if "A+ Content Date" in data.columns:
        data["A+ Content Date"] = pd.to_datetime(data["A+ Content Date"], errors='coerce')
        data = data[
            (data["A+ Content Date"].dt.month == month) & (data["A+ Content Date"].dt.year == year)]
    data = data.sort_values(by=["A+ Content Date"], ascending=True)

    result_count = len(data[data["Status"] == "Published"]) if "Status" in data.columns else 0

    if "A+ Content Date" in data.columns:
        data["A+ Content Date"] = data["A+ Content Date"].dt.strftime("%d-%B-%Y")

    data = data.fillna("N/A")

    data.index = range(1, len(data) + 1)

    return data, result_count


def get_A_plus_all(year) -> tuple[pd.DataFrame, int]:
    data = get_sheet_data(sheet_a_plus)
    if data.empty:
        return data

    columns = list(data.columns)
    if "Issues" in columns:
        end_col_index = columns.index("Issues")
        data = data.iloc[:, :end_col_index + 1]
    data = data.astype(str)

    if "A+ Content Date" in data.columns:
        data["A+ Content Date"] = pd.to_datetime(data["A+ Content Date"], errors='coerce')
        data = data[
            (data["A+ Content Date"].dt.year == year)]
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

    df['Publishing Date'] = pd.to_datetime(df['Publishing Date'], errors='coerce')
    df = df.dropna(subset=['Publishing Date', 'Name'])

    df['Month'] = df['Publishing Date'].dt.month_name()
    df['Year'] = df['Publishing Date'].dt.year

    month_1_names = set(
        df[(df['Month'] == month_1) & (df['Year'] == year1)]['Name'].str.strip()
    )

    month_2_names = set(
        df[(df['Month'] == month_2) & (df['Year'] == year2)]['Name'].str.strip()
    )

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


def format_review_counts_reviews(review_counts):
    """Format review counts as a string"""
    if review_counts.empty:
        return "No data"
    return ", ".join([f"{status}: {count}" for status, count in review_counts.items()])


def create_review_pie_chart(review_data, title):
    """Create pie chart for review distribution"""
    if isinstance(review_data, dict):
        if not review_data or sum(review_data.values()) == 0:
            return None
        values = list(review_data.values())
        labels = list(review_data.keys())
    else:  # Case 2: If review_data is a Pandas Series
        if review_data.empty or review_data.sum() == 0:
            return None
        values = review_data.values
        labels = review_data.index


    custom_colors = {
        "Attained": "#7dff8d",
        "Pending": "#ffc444",
        "-ve Attained": "#ff4b4b",
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


def create_platform_comparison_chart(usa_data, uk_data):
    """Create comparison chart for platforms"""
    platforms = ['Amazon', 'Barnes & Noble', 'Ingram Spark']

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


def create_brand_chart(usa_brands, uk_brands):
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


def summary(month, year):
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
    usa_clean = usa_clean.drop_duplicates(subset=["Name"])
    uk_clean = uk_clean.drop_duplicates(subset=["Name"])
    total_usa = usa_clean["Name"].nunique()
    total_uk = uk_clean["Name"].nunique()
    total_unique_clients = total_usa + total_uk

    combined = pd.concat([usa_clean[["Name", "Brand"]], uk_clean[["Name", "Brand"]]])
    combined.index = range(1, len(combined) + 1)

    brands = usa_clean["Brand"].value_counts()
    writers_clique = brands.get("Writers Clique", 0)
    bookmarketeers = brands.get("BookMarketeers", 0)
    aurora_writers = brands.get("Aurora Writers", 0)
    kdp = brands.get("KDP", "N/A")

    uk_brand = uk_clean["Brand"].value_counts()
    authors_solution = uk_brand.get("Authors Solution", 0)
    book_publication = uk_brand.get("Book Publication", 0)

    usa_platforms = usa_clean_platforms["Platform"].value_counts()
    usa_amazon = usa_platforms.get("Amazon", 0)
    usa_bn = usa_platforms.get("Barnes & Noble", 0)
    usa_ingram = usa_platforms.get("Ingram Spark", 0)

    uk_platforms = uk_clean_platforms["Platform"].value_counts()
    uk_amazon = uk_platforms.get("Amazon", 0)
    uk_bn = uk_platforms.get("Barnes & Noble", 0)
    uk_ingram = uk_platforms.get("Ingram Spark", 0)

    usa_review_sent = usa_clean[
        "Trustpilot Review"].value_counts().get("Sent", 0) if "Trustpilot Review" in usa_clean.columns else pd.Series()
    usa_review_pending = usa_clean[
        "Trustpilot Review"].value_counts().get("Pending",
                                                0) if "Trustpilot Review" in usa_clean.columns else pd.Series()
    usa_review_na = usa_clean[
        "Trustpilot Review"].value_counts().get("-ve Attained",
                                                0) if "Trustpilot Review" in usa_clean.columns else pd.Series()

    uk_review_sent = uk_clean["Trustpilot Review"].value_counts().get("Sent",
                                                                      0) if "Trustpilot Review" in uk_clean.columns else pd.Series()
    uk_review_pending = uk_clean["Trustpilot Review"].value_counts().get("Pending",
                                                                         0) if "Trustpilot Review" in uk_clean.columns else pd.Series()
    uk_review_na = uk_clean["Trustpilot Review"].value_counts().get("-ve Attained",
                                                                    0) if "Trustpilot Review" in uk_clean.columns else pd.Series()


    usa_reviews_df = load_reviews(sheet_usa, year, month)
    uk_reviews_df = load_reviews(sheet_uk, year, month)
    combined_data = pd.concat([usa_reviews_df, uk_reviews_df], ignore_index=True)

    usa_attained_pm = (
        usa_reviews_df[usa_reviews_df["Trustpilot Review"] == "Attained"]
        .groupby("Project Manager")["Trustpilot Review"]
        .count()
        .reset_index()
    )
    uk_attained_pm = (
        uk_reviews_df[uk_reviews_df["Trustpilot Review"] == "Attained"]
        .groupby("Project Manager")["Trustpilot Review"]
        .count()
        .reset_index()
    )

    attained_reviews_per_pm = (
        combined_data[combined_data["Trustpilot Review"] == "Attained"]
        .groupby("Project Manager")["Trustpilot Review"]
        .count()
        .reset_index()
    )

    usa_attained_pm.columns = ["Project Manager", "Attained Reviews"]
    usa_attained_pm.index = range(1, len(usa_attained_pm) + 1)
    usa_total_attained = usa_attained_pm["Attained Reviews"].sum()

    uk_attained_pm.columns = ["Project Manager", "Attained Reviews"]
    uk_attained_pm.index = range(1, len(uk_attained_pm) + 1)
    uk_total_attained = uk_attained_pm["Attained Reviews"].sum()

    attained_reviews_per_pm.columns = ["Project Manager", "Attained Reviews"]
    attained_reviews_per_pm.index = range(1, len(attained_reviews_per_pm) + 1)

    usa_review = {
        "Attained": usa_total_attained,
        "Sent": usa_review_sent,
        "Pending": usa_review_pending,
        "-ve Attained": usa_review_na
    }

    uk_review = {
        "Attained": uk_total_attained,
        "Sent": uk_review_sent,
        "Pending": uk_review_pending,
        "-ve Attained": uk_review_na
    }

    printing_data = get_printing_data_reviews(month, year)
    Total_copies = printing_data["No of Copies"].sum() if "No of Copies" in printing_data.columns else 0
    Total_cost = printing_data["Order Cost"].sum() if "Order Cost" in printing_data.columns else 0
    Highest_cost = printing_data["Order Cost"].max() if "Order Cost" in printing_data.columns else 0
    Highest_copies = printing_data["No of Copies"].max() if "No of Copies" in printing_data.columns else 0
    Lowest_cost = printing_data["Order Cost"].min() if "Order Cost" in printing_data.columns else 0
    Lowest_copies = printing_data["No of Copies"].min() if "No of Copies" in printing_data.columns else 0

    Average = Total_cost / Total_copies if Total_copies > 0 else 0
    if all(col in printing_data.columns for col in ["Order Cost", "No of Copies"]):
        printing_data['Cost_Per_Copy'] = printing_data['Order Cost'] / printing_data['No of Copies']

    copyright_data, result_count, result_count_no = get_copyright_data(month, year)
    Total_copyrights = len(copyright_data)
    Total_cost_copyright = Total_copyrights * 65
    country = copyright_data["Country"].value_counts()
    usa = country.get("USA", "N/A")
    canada = country.get("Canada", "N/A")
    uk = country.get("UK", "N/A")

    a_plus, a_plus_count = get_A_plus_data(month, year)

    usa_brands = {'BookMarketeers': bookmarketeers, 'Writers Clique': writers_clique, 'KDP': kdp,
                  'Aurora Writers': aurora_writers}

    uk_brands = {'Authors Solution': authors_solution, 'Book Publication': book_publication}

    usa_platforms = {'Amazon': usa_amazon, 'Barnes & Noble': usa_bn, 'Ingram Spark': usa_ingram}
    uk_platforms = {'Amazon': uk_amazon, 'Barnes & Noble': uk_bn, 'Ingram Spark': uk_ingram}

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

    return usa_review, uk_review, usa_brands, uk_brands, usa_platforms, uk_platforms, printing_stats, copyright_stats, a_plus_count, total_unique_clients, combined, attained_reviews_per_pm


def generate_year_summary(year):
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
        return
    if uk_clean.empty:
        print("No values found in UK sheet.")
        return
    if usa_clean.empty and uk_clean.empty:
        return

    usa_clean = usa_clean.drop_duplicates(subset=["Name"])
    uk_clean = uk_clean.drop_duplicates(subset=["Name"])
    total_usa = usa_clean["Name"].nunique()
    total_uk = uk_clean["Name"].nunique()
    total_unique_clients = total_usa + total_uk

    combined = pd.concat([usa_clean[["Name", "Brand"]], uk_clean[["Name", "Brand"]]])
    combined.index = range(1, len(combined) + 1)

    brands = usa_clean["Brand"].value_counts()
    writers_clique = brands.get("Writers Clique", 0)
    bookmarketeers = brands.get("BookMarketeers", 0)
    aurora_writers = brands.get("Aurora Writers", 0)
    kdp = brands.get("KDP", "N/A")

    uk_brand = uk_clean["Brand"].value_counts()
    authors_solution = uk_brand.get("Authors Solution", 0)
    book_publication = uk_brand.get("Book Publication", 0)

    usa_platforms = usa_clean_platforms["Platform"].value_counts()
    usa_amazon = usa_platforms.get("Amazon", 0)
    usa_bn = usa_platforms.get("Barnes & Noble", 0)
    usa_ingram = usa_platforms.get("Ingram Spark", 0)

    uk_platforms = uk_clean_platforms["Platform"].value_counts()
    uk_amazon = uk_platforms.get("Amazon", 0)
    uk_bn = uk_platforms.get("Barnes & Noble", 0)
    uk_ingram = uk_platforms.get("Ingram Spark", 0)

    usa_review_sent = usa_clean[
        "Trustpilot Review"].value_counts().get("Sent", 0) if "Trustpilot Review" in usa_clean.columns else pd.Series()
    usa_review_pending = usa_clean[
        "Trustpilot Review"].value_counts().get("Pending",
                                                0) if "Trustpilot Review" in usa_clean.columns else pd.Series()
    usa_review_na = usa_clean[
        "Trustpilot Review"].value_counts().get("-ve Attained",
                                                0) if "Trustpilot Review" in usa_clean.columns else pd.Series()

    uk_review_sent = uk_clean["Trustpilot Review"].value_counts().get("Sent",
                                                                      0) if "Trustpilot Review" in uk_clean.columns else pd.Series()
    uk_review_pending = uk_clean["Trustpilot Review"].value_counts().get("Pending",
                                                                         0) if "Trustpilot Review" in uk_clean.columns else pd.Series()
    uk_review_na = uk_clean["Trustpilot Review"].value_counts().get("-ve Attained",
                                                                         0) if "Trustpilot Review" in uk_clean.columns else pd.Series()

    usa_reviews_df = load_reviews(sheet_usa, year)

    uk_reviews_df = load_reviews(sheet_uk, year)
    combined_data = pd.concat([usa_reviews_df, uk_reviews_df], ignore_index=True)

    usa_attained_pm = (
        usa_reviews_df[usa_reviews_df["Trustpilot Review"] == "Attained"]
        .groupby("Project Manager")["Trustpilot Review"]
        .count()
        .reset_index()
    )
    uk_attained_pm = (
        uk_reviews_df[uk_reviews_df["Trustpilot Review"] == "Attained"]
        .groupby("Project Manager")["Trustpilot Review"]
        .count()
        .reset_index()
    )

    attained_reviews_per_pm = (
        combined_data[combined_data["Trustpilot Review"] == "Attained"]
        .groupby("Project Manager")["Trustpilot Review"]
        .count()
        .reset_index()
    )

    usa_attained_pm.columns = ["Project Manager", "Attained Reviews"]
    usa_attained_pm.index = range(1, len(usa_attained_pm) + 1)
    usa_total_attained = usa_attained_pm["Attained Reviews"].sum()

    uk_attained_pm.columns = ["Project Manager", "Attained Reviews"]
    uk_attained_pm.index = range(1, len(uk_attained_pm) + 1)
    uk_total_attained = uk_attained_pm["Attained Reviews"].sum()

    attained_reviews_per_pm.columns = ["Project Manager", "Attained Reviews"]
    attained_reviews_per_pm.index = range(1, len(attained_reviews_per_pm) + 1)

    usa_review = {
        "Attained": usa_total_attained,
        "Sent": usa_review_sent,
        "Pending": usa_review_pending,
        "-ve Attained": usa_review_na
    }

    uk_review = {
        "Attained": uk_total_attained,
        "Sent": uk_review_sent,
        "Pending": uk_review_pending,
        "-ve Attained": uk_review_na
    }

    printing_data = printing_data_all(year)
    Total_copies = printing_data["No of Copies"].sum() if "No of Copies" in printing_data.columns else 0
    Total_cost = printing_data["Order Cost"].sum() if "Order Cost" in printing_data.columns else 0
    Highest_cost = printing_data["Order Cost"].max() if "Order Cost" in printing_data.columns else 0
    Highest_copies = printing_data["No of Copies"].max() if "No of Copies" in printing_data.columns else 0
    Lowest_cost = printing_data["Order Cost"].min() if "Order Cost" in printing_data.columns else 0
    Lowest_copies = printing_data["No of Copies"].min() if "No of Copies" in printing_data.columns else 0

    Average = Total_cost / Total_copies if Total_copies > 0 else 0
    if all(col in printing_data.columns for col in ["Order Cost", "No of Copies"]):
        printing_data['Cost_Per_Copy'] = printing_data['Order Cost'] / printing_data['No of Copies']

    copyright_data, result_count, result_count_no = copyright_all(year)
    Total_copyrights = len(copyright_data)
    Total_cost_copyright = Total_copyrights * 65
    country = copyright_data["Country"].value_counts()
    usa = country.get("USA", "N/A")
    canada = country.get("Canada", "N/A")
    uk = country.get("UK", "N/A")

    a_plus, a_plus_count = get_A_plus_all(year)

    usa_brands = {'BookMarketeers': bookmarketeers, 'Writers Clique': writers_clique, 'KDP': kdp,
                  'Aurora Writers': aurora_writers}
    uk_brands = {'Authors Solution': authors_solution, 'Book Publication': book_publication}

    usa_platforms = {'Amazon': usa_amazon, 'Barnes & Noble': usa_bn, 'Ingram Spark': usa_ingram}
    uk_platforms = {'Amazon': uk_amazon, 'Barnes & Noble': uk_bn, 'Ingram Spark': uk_ingram}

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

    return usa_review, uk_review, usa_brands, uk_brands, usa_platforms, uk_platforms, printing_stats, copyright_stats, a_plus_count, total_unique_clients, combined, attained_reviews_per_pm


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


def generate_summary_report_pdf(usa_review_data, uk_review_data, usa_brands, uk_brands,
                                usa_platforms, uk_platforms, printing_stats, copyright_stats, a_plus,
                                selected_month, number, filename="summary_report.pdf"):
    """
    Generate a PDF version of the Streamlit summary report
    """

    # Create PDF document in memory
    if selected_month:
        filename = f"{selected_month}-{number} Summary Report.pdf"
    else:
        filename = f"{number} Summary Report.pdf"

    # Create BytesIO object to store PDF in memory
    buffer = BytesIO()

    doc = SimpleDocTemplate(buffer, pagesize=A4,
                            rightMargin=72, leftMargin=72,
                            topMargin=72, bottomMargin=18)

    # Get styles
    styles = getSampleStyleSheet()

    # Custom styles
    title_style = ParagraphStyle(
        'CustomTitle',
        parent=styles['Heading1'],
        fontSize=24,
        spaceAfter=30,
        alignment=TA_CENTER,
        textColor=colors.darkblue
    )

    section_style = ParagraphStyle(
        'SectionHeader',
        parent=styles['Heading2'],
        fontSize=16,
        spaceAfter=12,
        spaceBefore=20,
        textColor=colors.darkgreen
    )

    subsection_style = ParagraphStyle(
        'SubsectionHeader',
        parent=styles['Heading3'],
        fontSize=12,
        spaceAfter=8,
        spaceBefore=12,
        textColor=colors.darkblue
    )

    story = []

    if selected_month:
        story.append(Paragraph(f"{selected_month} {number} Summary Report", title_style))
    else:
        story.append(Paragraph(f"{number} Summary Report", title_style))

    story.append(Spacer(1, 20))

    # Calculate metrics
    # USA reviews
    usa_total = usa_review_data.sum() if hasattr(usa_review_data, "sum") else sum(usa_review_data.values())
    if isinstance(usa_review_data, dict):
        usa_attained = usa_review_data.get("Attained", 0)
    else:  # assume it's a Pandas object
        usa_attained = usa_review_data["Attained"] if "Attained" in usa_review_data else 0
    usa_attained_pct = (usa_attained / usa_total * 100) if usa_total > 0 else 0

    # UK reviews
    uk_total = uk_review_data.sum() if hasattr(uk_review_data, "sum") else sum(uk_review_data.values())
    if isinstance(uk_review_data, dict):
        uk_attained = uk_review_data.get("Attained", 0)
    else:
        uk_attained = uk_review_data["Attained"] if "Attained" in uk_review_data else 0
    uk_attained_pct = (uk_attained / uk_total * 100) if uk_total > 0 else 0

    combined_total = usa_total + uk_total
    combined_attained = usa_attained + uk_attained
    combined_attained_pct = (combined_attained / combined_total * 100) if combined_total > 0 else 0

    # Review Analytics Section
    story.append(Paragraph("📝 Review Analytics", section_style))
    story.append(HRFlowable(width="100%", thickness=1, lineCap='round', color=colors.lightgrey))
    story.append(Spacer(1, 12))

    # Review summary table
    review_data = [
        ['Region', 'Total Reviews', 'Attained', 'Success Rate'],
        ['🇺🇸 USA', f"{usa_total:,}", f"{usa_attained:,}", f"{usa_attained_pct:.1f}%"],
        ['🇬🇧 UK', f"{uk_total:,}", f"{uk_attained:,}", f"{uk_attained_pct:.1f}%"],
        ['Combined', f"{combined_total:,}", f"{combined_attained:,}", f"{combined_attained_pct:.1f}%"]
    ]

    review_table = Table(review_data)
    review_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, 0), 12),
        ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
        ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
        ('GRID', (0, 0), (-1, -1), 1, colors.black)
    ]))

    story.append(review_table)
    story.append(Spacer(1, 20))

    # Platform Distribution
    story.append(Paragraph("📱 Platform Distribution", subsection_style))

    total_client_usa = 0
    total_client_uk = 0
    # USA Platforms
    if usa_platforms:
        story.append(Paragraph("USA Platform Breakdown:", styles['Normal']))
        usa_platform_data = [['Platform', 'Count']]
        for platform, count in usa_platforms.items():
            usa_platform_data.append([platform, str(count)])

        usa_platform_table = Table(usa_platform_data)
        usa_platform_table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.lightblue),
            ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('GRID', (0, 0), (-1, -1), 1, colors.black)
        ]))
        story.append(usa_platform_table)
        story.append(Spacer(1, 10))

    # UK Platforms
    if uk_platforms:
        story.append(Paragraph("UK Platform Breakdown:", styles['Normal']))
        uk_platform_data = [['Platform', 'Count']]
        for platform, count in uk_platforms.items():
            uk_platform_data.append([platform, str(count)])

        uk_platform_table = Table(uk_platform_data)
        uk_platform_table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.lightblue),
            ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('GRID', (0, 0), (-1, -1), 1, colors.black)
        ]))
        story.append(uk_platform_table)
        story.append(Spacer(1, 20))

    # Brand Performance
    story.append(Paragraph("🏷️ Brand Performance", subsection_style))

    # Brand tables side by side
    if usa_brands and uk_brands:
        brand_data = [['USA Brands', 'Count', 'UK Brands', 'Count']]
        usa_items = list(usa_brands.items())
        uk_items = list(uk_brands.items())
        max_len = max(len(usa_items), len(uk_items))

        for i in range(max_len):
            usa_brand = usa_items[i] if i < len(usa_items) else ('', '')
            uk_brand = uk_items[i] if i < len(uk_items) else ('', '')
            brand_data.append([usa_brand[0], str(usa_brand[1]), uk_brand[0], str(uk_brand[1])])

        # Add totals row
        total_usa = sum(usa_brands.values())
        total_uk = sum(uk_brands.values())
        brand_data.append(['Total', str(total_usa), 'Total', str(total_uk)])

        brand_table = Table(brand_data)
        brand_table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.lightgreen),
            ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('GRID', (0, 0), (-1, -1), 1, colors.black),
            ('FONTNAME', (0, -1), (-1, -1), 'Helvetica-Bold'),
            ('BACKGROUND', (0, -1), (-1, -1), colors.lightgrey)
        ]))
        story.append(brand_table)
        story.append(Spacer(1, 20))

    story.append(PageBreak())
    story.append(Paragraph("🖨️ Printing Analytics", section_style))
    story.append(HRFlowable(width="100%", thickness=1, lineCap='round', color=colors.lightgrey))
    story.append(Spacer(1, 12))

    printing_data = [
        ['Metric', 'Value'],
        ['Total Copies', f"{printing_stats['Total_copies']:,}"],
        ['Highest Copies', str(printing_stats['Highest_copies'])],
        ['Lowest Copies', str(printing_stats['Lowest_copies'])],
        ['Total Cost', f"${printing_stats['Total_cost']:,.2f}"],
        ['Highest Cost', f"${printing_stats['Highest_cost']:.2f}"],
        ['Lowest Cost', f"${printing_stats['Lowest_cost']:.2f}"],
        ['Average Cost per Copy', f"${printing_stats['Average']:.2f}"]
    ]

    printing_table = Table(printing_data)
    printing_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.orange),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
        ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('GRID', (0, 0), (-1, -1), 1, colors.black),
        ('BACKGROUND', (0, 1), (-1, -1), colors.lightyellow)
    ]))

    story.append(printing_table)
    story.append(Spacer(1, 20))

    # Copyright Analytics Section
    story.append(Paragraph("©️ Copyright Analytics", section_style))
    story.append(HRFlowable(width="100%", thickness=1, lineCap='round', color=colors.lightgrey))
    story.append(Spacer(1, 12))
    success = copyright_stats['result_count']
    success_rate = (copyright_stats['result_count'] / copyright_stats['Total_copyrights'] * 100) if copyright_stats[
                                                                                                        'Total_copyrights'] > 0 else 0
    rejection_rate = (copyright_stats['result_count_no'] / copyright_stats['Total_copyrights'] * 100) if \
        copyright_stats[
            'Total_copyrights'] > 0 else 0
    copyright_data = [
        ['Metric', 'Value'],
        ['Total Copyrights', str(copyright_stats['Total_copyrights'])],
        ['Total Cost', f"${copyright_stats['Total_cost_copyright']:,}"],
        ['Success Count', f"{copyright_stats['result_count']}/{copyright_stats['Total_copyrights']}"],
        ['Success Percentage', f"{success_rate:.1f}%"],
        ['Rejected', f"{copyright_stats["result_count_no"]}"],
        ['Rejection Percentage', f"{rejection_rate:.1f}%"],
        ['USA Copyrights', str(copyright_stats['usa_copyrights'])],
        ['Canada Copyrights', str(copyright_stats['canada_copyrights'])]
    ]

    copyright_table = Table(copyright_data)
    copyright_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.purple),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
        ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('GRID', (0, 0), (-1, -1), 1, colors.black),
        ('BACKGROUND', (0, 1), (-1, -1), colors.lavender)
    ]))

    story.append(copyright_table)
    story.append(Spacer(1, 20))

    story.append(Paragraph("A+ Content", section_style))
    story.append(HRFlowable(width="100%", thickness=1, lineCap='round', color=colors.lightgrey))
    story.append(Spacer(1, 12))
    a_plus_data = [
        ['Metric', 'Count'],
        ['Total A+', str(a_plus)]
    ]

    a_plus_table = Table(a_plus_data)
    a_plus_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.blue),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
        ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('GRID', (0, 0), (-1, -1), 1, colors.black),
        ('BACKGROUND', (0, 1), (-1, -1), colors.beige)
    ]))

    story.append(a_plus_table)
    story.append(Spacer(1, 20))

    # Executive Summary Section
    story.append(Paragraph("📈 Executive Summary", section_style))
    story.append(HRFlowable(width="100%", thickness=1, lineCap='round', color=colors.lightgrey))
    story.append(Spacer(1, 12))

    exec_summary = f"""
    <b>Reviews:</b><br/>
    • Combined Reviews: {combined_total:,}<br/>
    • Success Rate: {combined_attained_pct:.1f}%<br/>
    • USA Attained: {usa_attained:,}<br/>
    • UK Attained: {uk_attained:,}<br/><br/>

    <b>Printing:</b><br/>
    • Total Copies: {printing_stats['Total_copies']:,}<br/>
    • Total Cost: ${printing_stats['Total_cost']:,.2f}<br/>
    • Cost Efficiency: ${printing_stats['Average']:.2f}/copy<br/><br/>

    <b>Copyright:</b><br/>
    • Applications: {copyright_stats['Total_copyrights']}<br/>
    • Success Rate: {success_rate:.1f}%<br/>
    • Rejection Rate: {rejection_rate:.1f}%<br/>
    • Total Cost: ${copyright_stats['Total_cost_copyright']:,}<br/>

    <b>A+ Content:</b><br/>
    • Total A+: {a_plus}<br/>

    """

    story.append(Paragraph(exec_summary, styles['Normal']))
    story.append(Spacer(1, 20))

    # Footer
    story.append(HRFlowable(width="100%", thickness=1, lineCap='round', color=colors.lightgrey))
    story.append(Spacer(1, 10))
    footer_text = f"Report generated on {datetime.now().strftime('%B %d, %Y at %I:%M %p')}"
    story.append(Paragraph(footer_text, styles['Normal']))

    # Build the PDF
    doc.build(story)

    # Get PDF data from buffer
    pdf_data = buffer.getvalue()
    buffer.close()

    return pdf_data, filename


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


def sales(month, year):
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

    data.index = range(1, len(data) + 1)

    return data


def main():
    with st.container():
        st.title("📊 Blink Digitally Publishing Dashboard")
        action = st.selectbox("What would you like to do?",
                              ["View Data", "Sales", "Printing", "Copyright", "Generate Similarity & Summary",
                               "Year Summary"],
                              index=None,
                              placeholder="Select Action")

        country = None
        selected_month = None
        selected_month_number = None
        status = None
        choice = None
        number = None
        if action in ["View Data", "Reviews"]:
            choice = st.selectbox("Select Data To View", ["USA", "UK", "AudioBook"], index=None,
                                  placeholder="Select Data to View")

        if action in ["View Data", "Reviews", "Copyright", "Sales"]:
            selected_month = st.selectbox(
                "Select Month",
                month_list,
                index=current_month - 1,
                placeholder="Select Month"
            )
            selected_month_number = month_list.index(selected_month) + 1 if selected_month else None
        if action in ["Year Summary", "Copyright", "View Data", "Reviews", "Sales"]:
            number = st.number_input("Enter Year", min_value=int(get_min_year()), step=1)
        if action == "Reviews":
            status = st.selectbox("Status", ["Pending", "Sent", "Attained"], index=None, placeholder="Select Status")

        if action == "View Data" and choice and selected_month and number:
            st.subheader(f"📂 Viewing Data for {choice} - {selected_month}")

            sheet_name = {
                "UK": sheet_uk,
                "USA": sheet_usa,
                "AudioBook": sheet_audio
            }.get(choice)
            if sheet_name:
                data = load_data(sheet_name, selected_month_number, number)

                review_data = load_reviews(sheet_name, number, selected_month_number)

                attained_reviews_per_pm = (
                    review_data[review_data["Trustpilot Review"] == "Attained"]
                    .groupby("Project Manager")["Trustpilot Review"]
                    .count()
                    .reset_index()
                )

                attained_reviews_per_pm.columns = ["Project Manager", "Attained Reviews"]
                attained_reviews_per_pm.index = range(1, len(attained_reviews_per_pm) + 1)

                if data.empty:
                    st.info(f"No data available for {selected_month} {number} for {choice}")
                else:
                    st.markdown("### 📄 Detailed Entry Data")
                    st.dataframe(data)

                    brands = data["Brand"].value_counts()
                    writers_clique = brands.get("Writers Clique", "N/A")
                    bookmarketeers = brands.get("BookMarketeers", "N/A")
                    aurora_writers = brands.get("Aurora Writers", "N/A")
                    kdp = brands.get("KDP", "N/A")
                    authors_solution = brands.get("Authors Solution", "N/A")

                    platforms = data["Platform"].value_counts()
                    amazon = platforms.get("Amazon", "N/A")
                    bn = platforms.get("Barnes & Noble", "N/A")
                    ingram = platforms.get("Ingram Spark", "N/A")
                    fav = platforms.get("FAV", "N/A")

                    sent = data["Trustpilot Review"].value_counts().get("Sent", 0)
                    pending = data["Trustpilot Review"].value_counts().get("Pending", 0)

                    review = {
                        "Sent": sent,
                        "Pending": pending,
                        "Attained": attained_reviews_per_pm["Attained Reviews"].sum()
                    }
                    publishing = data["Status"].value_counts()
                    total_reviews = sum(review.values())
                    # attained = reviews.get("Attained", 0)
                    attained = attained_reviews_per_pm["Attained Reviews"].sum()
                    percentage = round((attained / total_reviews * 100), 1) if total_reviews > 0 else 0

                    unique_clients_count_per_pm = data.groupby('Project Manager')['Name'].nunique().reset_index()
                    unique_clients_count_per_pm.columns = ['Project Manager', 'Unique Clients']
                    unique_clients_count_per_pm.index = range(1, len(unique_clients_count_per_pm) + 1)
                    total_unique_clients = data['Name'].nunique()

                    col1, col2 = st.columns(2)
                    with col1:

                        st.markdown("---")
                        st.markdown("### ⭐ Trustpilot Review Summary")
                        st.markdown(f"""
                                    - 🧾 **Total Entries:** `{len(data)}`
                                    - 🗳️ **Total Trustpilot Reviews:** `{total_reviews}`
                                    - 🟢 **'Attained' Reviews:** `{attained}`
                                    - 📊 **Attainment Rate:** `{percentage}%`

                                    **Brands**
                                    - 📘 **BookMarketeers:** `{bookmarketeers}`
                                    - 📘 **Aurora Writers:** `{aurora_writers}`
                                    - 📙 **Writers Clique:** `{writers_clique}`
                                    - 📕 **KDP:** `{kdp}`
                                    - 📘 **Authors Solution:** `{authors_solution}`

                                    **Platforms**
                                    - 🅰 **Amazon:** `{amazon}`
                                    - 📔 **Barnes & Noble:** `{bn}`
                                    - ⚡ **Ingram Spark:** `{ingram}`
                                    - 🔉 **Findaway Voices:** `{fav}`
                                    """)
                    with col2:
                        st.markdown("---")

                        st.markdown("#### 🔍 Review & Publishing Status Breakdown")
                        for review_type, count in review.items():
                            st.markdown(f"- 📝 **{review_type}**: `{count}`")

                        for status_type, count_s in publishing.items():
                            st.markdown(f"- 📘 **{status_type}**: `{count_s}`")
                        st.write("🤼 **Clients Per PM**")
                        st.dataframe(unique_clients_count_per_pm)
                        st.markdown(f"""

                            - ✅ **Total Unique:** `{total_unique_clients}` """
                                    )
                        st.write("👏 **Reviews Per PM**")
                        st.dataframe(attained_reviews_per_pm)
                st.markdown("---")
        # elif action == "Reviews" and choice and selected_month and status and number:
        #     sheet_name = {
        #         "UK": sheet_uk,
        #         "USA": sheet_usa,
        #         "AudioBook": sheet_audio
        #     }.get(choice)
        #     if sheet_name:
        #         data = review_data(sheet_name, selected_month_number, number, status)
        #
        #         st.subheader(f"🔍 Review Data - {status} in {selected_month} ({country})")
        #         if not data.empty:
        #             st.dataframe(data)
        #         else:
        #             st.info(f"No matching reviews found for {selected_month_number} {number}")
        elif action == "Printing":
            tab1, tab2 = st.tabs(["Monthly", "Total"])

            with tab1:
                selected_month = st.selectbox(
                    "Select Month",
                    month_list,
                    index=current_month - 1,
                    placeholder="Select Month"
                )
                number = st.number_input("Enter Year", min_value=int(get_min_year()), step=1)
                selected_month_number = month_list.index(selected_month) + 1 if selected_month else None

                if selected_month and number:
                    st.subheader(f"🖨️ Printing Summary for {selected_month}")

                    data = get_printing_data(selected_month_number, number)

                    if not data.empty:

                        Total_copies = data["No of Copies"].sum()

                        Total_cost = data["Order Cost"].sum()

                        Highest_cost = data["Order Cost"].max()

                        Highest_copies = data["No of Copies"].max()

                        Lowest_cost = data["Order Cost"].min()

                        Lowest_copies = data["No of Copies"].min()

                        # data['Cost_Per_Copy'] = data['Order Cost'] / data['No of Copies']
                        #
                        # da

                        Average = round(Total_cost / Total_copies, 2) if Total_copies else 0

                        st.markdown("### 📄 Detailed Printing Data")

                        st.dataframe(data)
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

                    else:
                        st.warning(f"⚠️ No Data Available for Printing in {selected_month} {number}")
            with tab2:
                number2 = st.number_input("Enter Year2", min_value=int(get_min_year()), step=1)
                data = printing_data_all(number2)
                if not data.empty:
                    st.markdown(f"### 📄 Total Printing Data for {number2}")
                    st.dataframe(data)
                else:
                    st.warning(f"⚠️ No Data Available for Printing in {number}")


        elif action == "Copyright" and selected_month and number:
            st.subheader(f"© Copyright Summary for {selected_month}")

            data, result, result_no = get_copyright_data(selected_month_number, number)

            if not data.empty:
                st.dataframe(data)

                total_copyrights = len(data)
                total_cost_copyright = total_copyrights * 65
                country = data["Country"].value_counts()
                usa = country.get("USA", "N/A")
                canada = country.get("Canada", "N/A")
                uk = country.get("UK", "N/A")
                st.markdown("---")
                st.markdown(f"""
                ### 📊 Summary Stats

                - 🧾 **Total Copyrighted Titles:** `{total_copyrights}`
                - 💵 **Copyright Total Cost:** `${total_cost_copyright}`
                - ✅ **Total Approved:** `{result} / {total_copyrights}`
                - 📈 **Total Approved rate:** `{result / total_copyrights:.1%}`
                - ❌ **Total Rejected:** `{result_no} / {total_copyrights}`
                - 📈 **Total Rejected rate:** `{result_no / total_copyrights:.1%}`
                - 🦅 **USA:** `{usa}`
                - 🍁 **Canada:** `{canada}`
                - ☕ **UK:** `{uk}`
                """)
                st.markdown("---")
            else:
                st.warning(f"⚠️ No Data Available for Copyright in {selected_month} {number}")
        elif action == "Generate Similarity & Summary":

            tab1, tab2 = st.tabs(["Queries", "Summary"])

            with tab1:
                st.header("Compare clients with months")
                choice = st.selectbox("Select Data To View", ["USA", "UK"], index=None,
                                      placeholder="Select Data to View")
                sheet_name = {
                    "UK": sheet_uk,
                    "USA": sheet_usa,
                }.get(choice)
                selected_month_1 = st.selectbox(
                    "Select Month 1",
                    month_list,
                    index=current_month - 2,
                    placeholder="Select Month 1"
                )
                number1 = st.number_input("Enter Year 1", min_value=int(get_min_year()), step=1)
                selected_month_2 = st.selectbox(
                    "Select Month 2",
                    month_list,
                    index=current_month - 1,
                    placeholder="Select Month 2"
                )
                number2 = st.number_input("Enter Year 2", min_value=int(get_min_year()), step=1)
                if sheet_name:
                    if st.button("Generate Similar Clients"):
                        with st.spinner(
                                f"Generating Similarity Report for {selected_month_1} N {selected_month_2} for {choice}..."):

                            data1, data2, data3 = get_names_in_both_months(sheet_name, selected_month_1, number1,
                                                                           selected_month_2, number2)

                            if not data1:
                                print(
                                    f"No Similar Clients found for {selected_month_1} N {selected_month_2} for {choice}")
                            st.write(data1)
                            st.write(data2)
                            st.write(data3)

                # st.subheader("🦅 USA Team")
                # usa_selected = st.multiselect("Select USA team members:", list(name_usa.keys()))
                #
                # st.subheader("☕ UK Team")
                # uk_selected = st.multiselect("Select UK team members:", list(names_uk.keys()))
                #
                # if st.button("Send Review Updates"):
                #     progress_bar = st.progress(0)
                #     total_members = len(usa_selected) + len(uk_selected)
                #     count = 0
                #
                #     for name in usa_selected:
                #         if name in name_usa:
                #             send_df_as_text(name, sheet_usa, name_usa[name])
                #             time.sleep(2)
                #             count += 1
                #             progress_bar.progress(count / total_members)
                #
                #     for name in uk_selected:
                #         if name in names_uk:
                #             send_df_as_text(name, sheet_uk, names_uk[name])
                #             time.sleep(2)
                #             count += 1
                #             progress_bar.progress(count / total_members)
                #
                #     st.success(f"Sent review updates to {count} team members!")

            with tab2:
                st.header("📄 Generate Summary Report")
                selected_month = st.selectbox(
                    "Select Month",
                    month_list,
                    index=current_month - 1,
                    placeholder="Select Month"
                )
                number = st.number_input("Enter Year", min_value=int(get_min_year()), step=1)
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
                            usa_review_data, uk_review_data, usa_brands, uk_brands, usa_platforms, uk_platforms, printing_stats, copyright_stats, a_plus, total_unique_clients, combined, attained_reviews_per_pm = summary(
                                selected_month_number, number)
                            pdf_data, pdf_filename = generate_summary_report_pdf(usa_review_data, uk_review_data,
                                                                                 usa_brands, uk_brands,
                                                                                 usa_platforms, uk_platforms,
                                                                                 printing_stats, copyright_stats,
                                                                                 a_plus,
                                                                                 selected_month, number)

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
                                    st.plotly_chart(usa_pie, use_container_width=True)

                                st.subheader("🇺🇸 USA Reviews")
                                st.metric("Total Reviews", usa_total)
                                st.metric("Total Attained", usa_attained)
                                st.metric("Attained Percentage", f"{usa_attained_pct:.1f}%")
                                st.metric("Total Unique", total_unique_clients)
                                st.dataframe(combined)
                            with col2:
                                uk_pie = create_review_pie_chart(uk_review_data, "UK Trustpilot Reviews")
                                if uk_pie:
                                    st.plotly_chart(uk_pie, use_container_width=True)
                                st.subheader("🇬🇧 UK Reviews")
                                st.metric("Total Reviews", uk_total)
                                st.metric("Total Attained", uk_attained)
                                st.metric("Attained Percentage", f"{uk_attained_pct:.1f}%")
                                st.dataframe(attained_reviews_per_pm)
                            st.subheader("📱 Platform Distribution")
                            platform_chart = create_platform_comparison_chart(usa_platforms, uk_platforms)
                            st.plotly_chart(platform_chart, use_container_width=True)

                            st.subheader("🏷️ Brand Performance")
                            brand_chart = create_brand_chart(usa_brands, uk_brands)
                            st.plotly_chart(brand_chart, use_container_width=True)

                            col1, col2 = st.columns(2)

                            with col1:
                                st.subheader("USA Brand Breakdown")
                                usa_df = pd.DataFrame(list(usa_brands.items()), columns=['Brand', 'Count'])
                                st.dataframe(usa_df, hide_index=True)

                                st.subheader("USA Platform Breakdown")
                                usa_platform_df = pd.DataFrame(list(usa_platforms.items()),
                                                               columns=['Platform', 'Count'])
                                st.dataframe(usa_platform_df, hide_index=True)

                            with col2:
                                st.subheader("UK Brand Breakdown")
                                uk_df = pd.DataFrame(list(uk_brands.items()), columns=['Brand', 'Count'])
                                st.dataframe(uk_df, hide_index=True)

                                st.subheader("UK Platform Breakdown")
                                uk_platform_df = pd.DataFrame(list(uk_platforms.items()), columns=['Platform', 'Count'])
                                st.dataframe(uk_platform_df, hide_index=True)

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
                                st.plotly_chart(fig_gauge, use_container_width=True)

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
                                st.plotly_chart(fig_copyright, use_container_width=True)

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
                        usa_review_data, uk_review_data, usa_brands, uk_brands, usa_platforms, uk_platforms, printing_stats, copyright_stats, a_plus, total_unique_clients, combined, attained_reviews_per_pm = generate_year_summary(
                            number)
                        pdf_data, pdf_filename = generate_summary_report_pdf(usa_review_data, uk_review_data,
                                                                             usa_brands, uk_brands,
                                                                             usa_platforms, uk_platforms,
                                                                             printing_stats, copyright_stats, a_plus,
                                                                             selected_month, number)

                        # Calculate metrics for display
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
                                st.plotly_chart(usa_pie, use_container_width=True)

                            st.subheader("🇺🇸 USA Reviews")
                            st.metric("Total Reviews", usa_total)
                            st.metric("Total Attained", usa_attained)
                            st.metric("Attained Percentage", f"{usa_attained_pct:.1f}%")
                            st.metric("Total Unique", total_unique_clients)
                            st.dataframe(combined)

                        with col2:
                            uk_pie = create_review_pie_chart(uk_review_data, "UK Trustpilot Reviews")
                            if uk_pie:
                                st.plotly_chart(uk_pie, use_container_width=True)
                            st.subheader("🇬🇧 UK Reviews")
                            st.metric("Total Reviews", uk_total)
                            st.metric("Total Attained", uk_attained)
                            st.metric("Attained Percentage", f"{uk_attained_pct:.1f}%")
                            st.dataframe(attained_reviews_per_pm)
                        st.subheader("📱 Platform Distribution")
                        platform_chart = create_platform_comparison_chart(usa_platforms, uk_platforms)
                        st.plotly_chart(platform_chart, use_container_width=True)

                        st.subheader("🏷️ Brand Performance")
                        brand_chart = create_brand_chart(usa_brands, uk_brands)
                        st.plotly_chart(brand_chart, use_container_width=True)

                        col1, col2 = st.columns(2)

                        with col1:
                            st.subheader("USA Brand Breakdown")
                            usa_df = pd.DataFrame(list(usa_brands.items()), columns=['Brand', 'Count'])
                            st.dataframe(usa_df, hide_index=True)

                            st.subheader("USA Platform Breakdown")
                            usa_platform_df = pd.DataFrame(list(usa_platforms.items()), columns=['Platform', 'Count'])
                            st.dataframe(usa_platform_df, hide_index=True)

                        with col2:
                            st.subheader("UK Brand Breakdown")
                            uk_df = pd.DataFrame(list(uk_brands.items()), columns=['Brand', 'Count'])
                            st.dataframe(uk_df, hide_index=True)

                            st.subheader("UK Platform Breakdown")
                            uk_platform_df = pd.DataFrame(list(uk_platforms.items()), columns=['Platform', 'Count'])
                            st.dataframe(uk_platform_df, hide_index=True)

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
                            st.plotly_chart(fig_gauge, use_container_width=True)

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
                            st.plotly_chart(fig_copyright, use_container_width=True)

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

        elif action == "Sales" and selected_month and number:
            data = sales(selected_month_number, number)
            if not data.empty:
                Total_sales = data["Payment"].sum()

                st.markdown("### 📄 Detailed Sales Data")

                st.dataframe(data)
                st.markdown("---")

                st.markdown("### 📊 Summary Statistics")

                st.markdown(f"""

                              - 🧾 **Total Clients:** {len(data)}

                              - 💰 **Total Sales:** `{Total_sales}`

                              """)
                st.markdown("---")

            else:

                st.warning(f"⚠️ No Data Available for Sales in {selected_month} {number}")


if __name__ == '__main__':
    main()
