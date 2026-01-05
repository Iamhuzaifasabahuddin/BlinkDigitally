import calendar
import logging
import os
import time
from datetime import datetime

import gspread
import pandas as pd
import pytz
import streamlit as st
import streamlit_authenticator as stauth
from google.oauth2.service_account import Credentials
from slack_sdk import WebClient
from slack_sdk.errors import SlackApiError

st.set_page_config(
    page_title="Trustpilot Review Manager",
    page_icon="⭐",
    layout="centered",
    initial_sidebar_state="collapsed"
)
APP_PASSWORD_NORMAL = st.secrets["APP"]["app_password_normal"]
APP_PASSWORD_ADMIN = st.secrets["APP"]["app_password_admin"]


# Initialize logging
@st.cache_resource
def setup_logging():
    logging.basicConfig(
        level=logging.INFO,
        format='%(funcName)s --> %(message)s : %(asctime)s - %(levelname)s',
        datefmt="%d-%m-%Y %I:%M:%S %p"
    )
    file_handler = logging.FileHandler('Reviews.log')
    file_handler.setLevel(level=logging.WARNING)
    formatter = logging.Formatter(
        '%(funcName)s --> %(message)s : %(asctime)s - %(levelname)s',
        "%d-%m-%Y %I:%M:%S %p"
    )
    file_handler.setFormatter(formatter)
    logger = logging.getLogger('')
    logger.addHandler(file_handler)


setup_logging()

SLACK_BOT_TOKEN = os.getenv("SLACK") or st.secrets["Slack"]["Slack"]
client = WebClient(token=SLACK_BOT_TOKEN)
channel_usa = os.getenv("CHANNEL_USA") or st.secrets["Channels"]["usa"]
channel_uk = os.getenv("CHANNEL_UK") or st.secrets["Channels"]["uk"]
SPREADSHEET_ID = os.getenv("SPREADSHEET_ID") or st.secrets['connections']['gsheets']['SPREADSHEET_ID']

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

# Constants
sheet_usa = "USA"
sheet_uk = "UK"
PKST_DATE = pytz.timezone("Asia/Karachi")

now_pk = datetime.now(PKST_DATE)
current_year = now_pk.year

name_usa = {
    "Aiza Ali": "aiza.ali@topsoftdigitals.pk",
    "Ahmed Asif": "ahmed.asif@topsoftdigitals.pk",
    "Asad Waqas": "asad.waqas@topsoftdigitals.pk",
    "Kamal Muhammad Issa": "kamal.muhammed.issa@topsoftdigitals.pk",
    "Maheen Sami": "maheen.sami@topsoftdigitals.pk",
    "Mubashir Khan": "Mubashir.khan@topsoftdigitals.pk",
    "Muhammad Ali": "muhammad.ali@topsoftdigitals.pk",
    "Valencia Angelo": "valencia.angelo@topsoftdigitals.pk",
    "Ukasha Asadullah": "ukasha.asadullah@topsoftdigitals.pk",
    "Ahsan Javed": "ahsan.javed@topsoftdigitals.pk",
    "Muhammad Saad Sethi": "saad.sethi@topsoftdigitals.pk",
    "Tooba Shoaib": "tooba.shoaib@topsoftdigitals.pk"
}

names_uk = {
    "Youha": "youha.khan@topsoftdigitals.pk",
    "Hassan Siddiqui": "hassan.siddiqui@topsoftdigitals.pk",
    "Emaan Zaidi": "emaan.zaidi@topsoftdigitals.pk",
    "Elishba": "elishba@topsoftdigitals.pk",
    "Shahrukh Yousuf": "shahrukh.yousuf@topsoftdigitals.pk"
}

general_message = """Hiya
:bangbang: Please ask the following Clients for their feedback about their respective projects for the ones marked as _*Pending*_ & for those marked as _*Sent*_ please remind the clients once again that their feedback truly matters and helps us grow and make essential changes to make the process even more fluid!
BM: https://bookmarketeers.com/
WC: https://writersclique.com/
AW: https://aurorawriters.com/
AS: https://authorssolution.co.uk/
BP: https://bookpublication.co.uk/
"""


def get_gspread_client():
    try:
        credentials = Credentials.from_service_account_info(creds_dict, scopes=SCOPES)
        gc = gspread.authorize(credentials)
        return gc
    except Exception as e:
        logging.error(f"Failed to initialize gspread client: {e}")
        st.error(f"Failed to connect to Google Sheets: {e}")
        return None


@st.cache_data(ttl=300)
def get_sheet_data(sheet_name):
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
        st.error(f"Failed to get data from sheet {sheet_name}: {e}")
        return pd.DataFrame()


def normalize_name(name):
    if pd.isna(name) or name == "":
        return ""
    return str(name).strip().title()


def clean_data_reviews(sheet_name: str) -> pd.DataFrame:
    data = get_sheet_data(sheet_name)
    if data.empty:
        logging.warning(f"No data found in sheet: {sheet_name}")
        return data

    columns = list(data.columns)
    if "Issues" in columns:
        end_col_index = columns.index("Issues")
        data = data.iloc[:, :end_col_index + 1]

    if "Project Manager" in data.columns:
        data["Project Manager"] = data["Project Manager"].apply(normalize_name)

    for col in ["Publishing Date", "Last Edit (Revision)", "Trustpilot Review Date"]:
        if col in data.columns:
            data[col] = pd.to_datetime(data[col], format="%d-%B-%Y", errors="coerce")

    data = data.sort_values(by="Publishing Date", ascending=True)
    data.index = range(1, len(data) + 1)
    return data

def load_sent_reviews(sheet_name: str, name: str):
    data = clean_data_reviews(sheet_name)
    name = normalize_name(name)
    data_original = data.copy()

    if data.empty:
        return pd.DataFrame(), pd.NaT, pd.NaT

    data = data_original[
        (data_original["Project Manager"] == name) &
        ((data_original["Trustpilot Review"] == "Pending") | (data_original["Trustpilot Review"] == "Sent")) &
        (data_original["Brand"].isin(
            ["BookMarketeers", "Writers Clique", "Authors Solution", "Book Publication", "Aurora Writers"])) &
        (data_original["Status"] == "Published")
        ]

    data = data.sort_values(by="Publishing Date", ascending=True)
    if "Name" in data.columns:
        data = data.drop_duplicates(subset=["Name"], keep="last")

    min_date = data["Publishing Date"].min() if not data.empty else pd.NaT
    max_date = data["Publishing Date"].max() if not data.empty else pd.NaT
    data.index = range(1, len(data) + 1)

    return data, min_date, max_date

def load_pending_reviews(sheet_name: str, name: str) -> tuple:
    data = clean_data_reviews(sheet_name)
    name = normalize_name(name)
    data_original = data.copy()

    if data.empty:
        return pd.DataFrame(), pd.NaT, pd.NaT

    data = data_original[
        (data_original["Project Manager"] == name) &
        (data_original["Trustpilot Review"] == "Pending") &
        (data_original["Brand"].isin(
            ["BookMarketeers", "Writers Clique", "Authors Solution", "Book Publication", "Aurora Writers"])) &
        (data_original["Status"] == "Published")
        ]

    data = data.sort_values(by="Publishing Date", ascending=True)
    if "Name" in data.columns:
        data = data.drop_duplicates(subset=["Name"], keep="last")

    min_date = data["Publishing Date"].min() if not data.empty else pd.NaT
    max_date = data["Publishing Date"].max() if not data.empty else pd.NaT
    data.index = range(1, len(data) + 1)

    return data, min_date, max_date


def load_attained_reviews(sheet_name: str, name: str, year: int, month_number=None) -> pd.DataFrame:
    data = get_sheet_data(sheet_name)
    if data.empty:
        return pd.DataFrame()

    columns = list(data.columns)
    if "Issues" in columns:
        end_col_index = columns.index("Issues")
        data = data.iloc[:, :end_col_index + 1]

    if "Project Manager" in data.columns:
        data["Project Manager"] = data["Project Manager"].apply(normalize_name)

    name = normalize_name(name)

    date_columns = ["Publishing Date", "Last Edit (Revision)", "Trustpilot Review Date"]
    for col in date_columns:
        if col in data.columns:
            data[col] = pd.to_datetime(data[col], format="%d-%B-%Y", errors="coerce")

    try:
        if "Trustpilot Review Date" in data.columns and month_number:
            data = data[(data["Trustpilot Review Date"].dt.month == month_number) &
                        (data["Trustpilot Review Date"].dt.year == year)]
        else:
            data = data[(data["Trustpilot Review Date"].dt.year == year)]

        data_original = data.copy()
        data = data_original[
            (data_original["Project Manager"] == name) &
            (data_original["Trustpilot Review"] == "Attained") &
            (data_original["Brand"].isin(
                ["BookMarketeers", "Writers Clique", "Authors Solution", "Book Publication", "Aurora Writers"]))
            ]

        data = data.sort_values(by="Trustpilot Review Date", ascending=True)
        data = data.drop_duplicates(subset=["Name"])
        data["Trustpilot Review Date"] = pd.to_datetime(
            data["Trustpilot Review Date"], errors="coerce"
        ).dt.strftime("%d-%B-%Y")
        data.index = range(1, len(data) + 1)
        return data
    except Exception as e:
        logging.error(f"Error loading attained reviews: {e}")
        return pd.DataFrame()


def load_total_reviews(sheet_name: str, name: str, year: int, month_number=None):
    data = clean_data_reviews(sheet_name)
    name = normalize_name(name)

    if "Name" in data.columns:
        data = data.drop_duplicates(subset=["Name"], keep="last")

    if data.empty:
        return 0

    date_columns = ["Publishing Date", "Last Edit (Revision)", "Trustpilot Review Date"]
    for col in date_columns:
        if col in data.columns:
            data[col] = pd.to_datetime(data[col], format="%d-%B-%Y", errors="coerce")

    try:
        if "Publishing Date" in data.columns and month_number:
            data = data[(data["Publishing Date"].dt.month == month_number) &
                        (data["Publishing Date"].dt.year == year)]
        else:
            data = data[(data["Publishing Date"].dt.year == year)]

        data_original = data.copy()
        data_count = data_original[
            (data_original["Project Manager"] == name) &
            ((data_original["Trustpilot Review"] == "Pending") |
             (data_original["Trustpilot Review"] == "Sent")) &
            (data_original["Brand"].isin(
                ["BookMarketeers", "Writers Clique", "Authors Solution", "Book Publication", "Aurora Writers"])) &
            (data_original["Status"] == "Published")
            ]
        data_count["Publishing Date"] = pd.to_datetime(
            data_count["Publishing Date"], errors="coerce"
        ).dt.strftime("%d-%B-%Y")

        return data_count
    except Exception as e:
        logging.error(f"Error calculating total reviews: {e}")
        return 0


def get_user_id_by_email(email: str):
    try:
        response = client.users_lookupByEmail(email=email)
        return response['user']['id']
    except SlackApiError as e:
        logging.error(f"Error finding user: {e.response['error']}")
        return None


def send_dm(user_id: str, message: str):
    try:
        client.chat_postMessage(channel=user_id, text=message)
    except SlackApiError as e:
        logging.error(f"Error sending DM: {e.response['error']}")


def send_pending_reviews_per_pm(name: str, sheet_name: str, email: str, channel: str) -> bool:
    user_id = get_user_id_by_email(email)
    # user_id = get_user_id_by_email("huzaifa.sabah@topsoftdigitals.pk")

    if not user_id:
        st.error(f"Could not find user ID for {name}")
        return False

    df, min_date, max_date = load_pending_reviews(sheet_name, name)

    if df.empty:
        st.warning(f"No pending reviews found for {name}")
        return False

    min_dates = [d for d in [min_date] if pd.notna(d)]
    max_dates = [d for d in [max_date] if pd.notna(d)]

    if not min_dates or not max_dates:
        st.warning(f"No valid dates for {name}")
        return False

    min_month_name = min(min_dates).strftime("%B")
    max_month_name = max(max_dates).strftime("%B")

    def truncate_title(x):
        return x[:20] + "..." if isinstance(x, str) and len(x) > 20 else x

    if "Book Name & Link" in df.columns and not df.empty:
        df["Book Name & Link"] = df["Book Name & Link"].apply(truncate_title)

    if "Publishing Date" in df.columns and not df.empty:
        df["Publishing Date"] = pd.to_datetime(df["Publishing Date"], errors='coerce').dt.strftime("%d-%B-%Y")

    merged_df = df[["Name", "Brand", "Book Name & Link", "Publishing Date", "Trustpilot Review"]]

    if not merged_df.empty:
        markdown_table = merged_df.to_markdown(index=True)

        if len({min_month_name, max_month_name}) > 1:
            message = (
                f"{general_message}\n\n"
                f"Hi 👋🏻 <@{user_id}>! Here's your Trustpilot update from {min_month_name} to {max_month_name} {current_year} 📄\n\n"
                f"*Summary:* ❓ {len(merged_df)} pending reviews\n\n"
                f"```\n{markdown_table}\n```"
            )
        else:
            message = (
                f"{general_message}\n\n"
                f"Hi 👋🏻 <@{user_id}>! Here's your Trustpilot update for {min_month_name} {current_year} 📄\n\n"
                f"*Summary:* ❓ {len(merged_df)} pending reviews\n\n"
                f"```\n{markdown_table}\n```"
            )

        try:
            # conversation = client.conversations_open(users=user_id)
            # channel_id = conversation['channel']['id']
            client.chat_postMessage(channel=channel, text=message, mrkdwn=True)
            send_dm(get_user_id_by_email("huzaifa.sabah@topsoftdigitals.pk"),
                    f"✅ Review sent to {name}")
            return True
        except SlackApiError as e:
            st.error(f"Error sending message to {name}: {e.response['error']}")
            logging.error(e)
            return False
    return False


def send_attained_reviews_per_pm(pm_name: str, email: str, sheet_name: str, year: int,
                                 channel: str, month=None) -> bool:
    user_id = get_user_id_by_email(email)
    # user_id = get_user_id_by_email("huzaifa.sabah@topsoftdigitals.pk")

    if not user_id:
        st.error(f"Could not find user ID for {pm_name}")
        return False

    total_reviews = len(load_total_reviews(sheet_name, pm_name, year, month))
    review_data = load_attained_reviews(sheet_name, pm_name, year, month)

    review_details_df = review_data.copy()
    attained_details = review_details_df[["Name", "Brand", "Trustpilot Review Date"]]
    attained_details.index = range(1, len(attained_details) + 1)

    if attained_details.empty:
        st.warning(f"No attained reviews found for {pm_name}")
        return False

    markdown_table = attained_details.to_markdown(index=True)
    Total = total_reviews + len(attained_details)
    percentage = len(attained_details) / Total if Total > 0 else 0

    message = (
        f"Hi 👋🏻 <@{user_id}>! Here's your Trustpilot Summary from {current_year} 🧮\n\n"
        f"*Summary:* ❓ {total_reviews} pending reviews\n\n"
        f"*Review Retention:* 🎯 {len(attained_details)} out of {Total} "
        f"(📊 {percentage:.1%})\n\n"
        f"```\n{markdown_table}\n```"
    )

    try:
        # conversation = client.conversations_open(users=user_id)
        # channel_id = conversation['channel']['id']
        client.chat_postMessage(channel=channel, text=message, mrkdwn=True)
        send_dm(get_user_id_by_email("huzaifa.sabah@topsoftdigitals.pk"),
                f"✅ Attained review details sent for {pm_name}")
        return True
    except SlackApiError as e:
        st.error(f"Error sending message for {pm_name}: {e.response['error']}")
        logging.error(e)
        return False

def printing_data_month(month: int, year: int, choice: str) -> pd.DataFrame:
    """Get printing data filtered by month"""
    try:
        usa_brands = ["BookMarketeers", "Writers Clique", "Aurora Writers", "KDP"]
        uk_brands = ["Authors Solution", "Book Publication"]

        if choice == "USA":
            selected_brands = usa_brands
        else:
            selected_brands = uk_brands
        data = get_sheet_data("Printing")

        columns = list(data.columns)
        if "Accepted" in columns:
            end_col_index = columns.index("Accepted")
            data = data.iloc[:, :end_col_index + 1]

        data = data.astype(str)
        for col in ["Order Date", "Shipping Date", "Fulfilled"]:
            if col in data.columns:
                data[col] = pd.to_datetime(data[col], format="%d-%B-%Y", errors="coerce")

        if month and "Order Date" in data.columns:
            data = data[(data["Order Date"].dt.month == month) & (data["Order Date"].dt.year == year) & (data["Brand"].isin(selected_brands))]
        if data.empty:
            return pd.DataFrame()
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

        return data
    except Exception as e:
        st.error(f"Error loading printing data: {e}")
        return pd.DataFrame()

def printing_data_year(year: int, choice: str) -> pd.DataFrame:
    data = get_sheet_data("Printing")
    usa_brands = ["BookMarketeers", "Writers Clique", "Aurora Writers", "KDP"]
    uk_brands = ["Authors Solution", "Book Publication"]

    if choice == "USA":
        selected_brands = usa_brands
    else:
        selected_brands = uk_brands
    data = get_sheet_data("Printing")
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

    data = data[(data["Order Date"].dt.year == year) & (data["Brand"].isin(selected_brands))]

    if data.empty:
        return pd.DataFrame()

    if "Order Cost" in data.columns:
        data["Order Cost"] = pd.to_numeric(
            data["Order Cost"].str.replace("$", "", regex=False).str.replace(",", "", regex=False),
            errors="coerce"
        )

    if "No of Copies" in data.columns:
        data["No of Copies"] = pd.to_numeric(data["No of Copies"], errors='coerce')

    for col in ["Order Date", "Shipping Date", "Fulfilled"]:
        if col in data.columns:
            data[col] = data[col].dt.strftime("%d-%B-%Y")

    data.index = range(1, len(data) + 1)

    return data

def get_printing_upcoming(choice: str):
    data = get_sheet_data("Printing")
    usa_brands = ["BookMarketeers", "Writers Clique", "Aurora Writers", "KDP"]
    uk_brands = ["Authors Solution", "Book Publication"]

    if choice == "USA":
        selected_brands = usa_brands
    else:
        selected_brands = uk_brands
    data = get_sheet_data("Printing")
    if data.empty:
        return pd.DataFrame(), pd.DataFrame()

    columns = list(data.columns)
    if "Accepted" in columns:
        end_col_index = columns.index("Accepted")
        data = data.iloc[:, :end_col_index + 1]

    data = data.astype(str)
    if data.empty:
        return pd.DataFrame(), pd.DataFrame()


    data = data[(data["Type"] == "Upcoming") & (data["Brand"].isin(selected_brands))]
    data.index = range(1, len(data) + 1)
    return data

def fetch(region: str):
    st.cache_data.clear()
    st.info(f"Fetching latest reviews for {region} ...")
    st.session_state.fetched = True
    pkt = pytz.timezone("Asia/Karachi")
    now_pkt = datetime.now(pkt)
    st.session_state.last_fetch_time = now_pkt.strftime("%d-%B-%Y @ %I:%M %p")
    st.rerun()


def main():
    config = {
        'credentials': {
            'usernames': {}
        },
        'cookie': {
            'name': st.secrets.get("cookie_name", "Review_manager_cookie"),
            'key': st.secrets.get("cookie_key", "some_signature_key"),
            'expiry_days': st.secrets.get("cookie_expiry_days", 30)
        }
    }

    for key in st.secrets:
        if key.startswith("auth_username_"):
            username = st.secrets[key]
            suffix = key.replace("auth_username_", "")
            config['credentials']['usernames'][username] = {
                'name': st.secrets.get(f"auth_name_{suffix}", username),
                'email': st.secrets.get(f"auth_email_{suffix}", f"{username}@example.com"),
                'password': st.secrets.get(f"auth_password_{suffix}", "")
            }

    authenticator = stauth.Authenticate(
        config['credentials'],
        config['cookie']['name'],
        config['cookie']['key'],
        config['cookie']['expiry_days']
    )
    if st.session_state.get('authentication_status') is None:
        st.title("🔑 Review Manager Login")

    authenticator.login()

    if st.session_state.get('authentication_status') is True:

        st.session_state.authenticated_admin = False
        st.session_state.authenticated_normal = False

        if st.session_state.get('name') == "Admin":
            st.session_state.authenticated_admin = True
        elif st.session_state.get('name') == "PM":
            st.session_state.authenticated_normal = True

        st.title("⭐ Printing Data & Trustpilot Review Manager")
        st.markdown("---")

        with st.sidebar:
            st.header("📋 Settings")
            region = st.selectbox("Select Region", ["USA", "UK"])

            if "fetched" not in st.session_state:
                st.session_state.fetched = False
            if "last_fetch_time" not in st.session_state:
                st.session_state.last_fetch_time = None

            if st.session_state.authenticated_admin:
                action = st.radio(
                    "Select Action",
                    ["View Reviews", "Send Pending Reviews", "Send Attained Reviews", "Bulk Send", "Printing Data"]
                )
            elif st.session_state.authenticated_normal:
                action = st.radio(
                    "Select Action",
                    ["View Reviews", "Printing Data"]
                )

            if st.button("🔃 Fetch Latest"):
                fetch(region)

            if st.session_state.fetched:
                st.success(
                    f"✅ Latest reviews fetched for {region} at {st.session_state.last_fetch_time} PKST"
                )
                st.session_state.fetched = False

            st.markdown("---")
            st.info("💡 Use this app to manage and send Trustpilot reviews to project managers.")
            authenticator.logout()

        sheet_name = sheet_usa if region == "USA" else sheet_uk
        pm_names = name_usa if region == "USA" else names_uk
        channel = channel_usa if region == "USA" else channel_uk

        if action == "View Reviews":
            st.header(f"📊 View Reviews - {region}")
            col1, col2 = st.columns(2)

            with col1:
                selected_pm = st.selectbox("Select Project Manager", list(pm_names.keys()))

            with col2:
                pass

            review_type_P, review_type_A = st.tabs(["Pending", "Attained"])

            with review_type_P:
                df, min_date, max_date = load_pending_reviews(sheet_name, selected_pm)
                df2, _, _ = load_sent_reviews(sheet_name, selected_pm)

                if not df.empty:
                    st.success(f"Found {len(df)} pending reviews for {selected_pm}")
                    if pd.notna(min_date) and pd.notna(max_date):
                        st.info(f"Date Range: {min_date.strftime('%B %Y')} - {max_date.strftime('%B %Y')}")

                    df["Publishing Date"] = pd.to_datetime(df["Publishing Date"], errors='coerce').dt.strftime("%d-%B-%Y")
                    df = df[[
                        "Name", "Brand", "Publishing Date", "Status",
                        "Trustpilot Review", "Trustpilot Review Date", "Trustpilot Review Links"
                    ]].rename(columns={"Status": "Publishing Status"})
                    with st.expander("❓ Pending Reviews"):
                        st.dataframe(df, width="stretch")
                else:
                    st.warning(f"No pending reviews found for {selected_pm}")

                if not df2.empty:
                    df2["Publishing Date"] = pd.to_datetime(df2["Publishing Date"], errors='coerce').dt.strftime("%d-%B-%Y")
                    df2 = df2[[
                        "Name", "Brand", "Publishing Date", "Status", "Trustpilot Review"
                    ]].rename(columns={"Status": "Publishing Status"})
                    with st.expander("📦 Pending Reviews Total"):
                        st.dataframe(df2, width="stretch")
                else:
                    st.warning(f"No total pending reviews found for {selected_pm}")

            with review_type_A:
                year = st.number_input("Select Year", min_value=2025, max_value=current_year, value=current_year)
                month = st.selectbox("Select Month (Optional)", ["All"] + list(calendar.month_name)[1:])
                month_number = None if month == "All" else list(calendar.month_name).index(month)

                df = load_attained_reviews(sheet_name, selected_pm, year, month_number)
                total_reviews = load_total_reviews(sheet_name, selected_pm, year, month_number)

                if not df.empty:
                    total = len(total_reviews) + len(df)
                    percent = len(df) / total if total > 0 else 0

                    col1, col2, col3, col4 = st.columns(4)
                    col1.metric("✅ Attained Reviews", len(df))
                    col2.metric("❓ Pending Reviews", len(total_reviews))
                    col3.metric("🤵🏻 Total Reviews", f"{total}")
                    col4.metric("🎯 Retention Rate", f"{percent:.1%}")

                    df["Publishing Date"] = pd.to_datetime(df["Publishing Date"], errors='coerce').dt.strftime("%d-%B-%Y")
                    df = df[[
                        "Name", "Brand", "Publishing Date", "Status",
                        "Trustpilot Review", "Trustpilot Review Date", "Trustpilot Review Links"
                    ]].rename(columns={"Status": "Publishing Status"})
                    with st.expander(f"🟢 Attained Reviews {month} {year}"):
                        st.dataframe(df, width="stretch")
                    with st.expander(f"❓ Pending Reviews {month} {year}"):
                        total_reviews = total_reviews[[
                        "Name", "Brand", "Publishing Date", "Status",
                        "Trustpilot Review", "Trustpilot Review Date", "Trustpilot Review Links"
                    ]].rename(columns={"Status": "Publishing Status"})
                        total_reviews.index = range(1, len(total_reviews) + 1)
                        st.dataframe(total_reviews, width="stretch")
                else:
                    st.warning(f"No attained reviews found for {selected_pm}")

        elif action == "Printing Data":
            month_list = list(calendar.month_name)[1:]
            current_month = datetime.today().strftime("%B")

            tab_m, tab_y, tab_u, tab_s = st.tabs(["Monthly", "Yearly", "Upcoming", "Search"])

            with tab_m:

                month = st.selectbox(
                    "Select Month:",
                    month_list,
                    index=month_list.index(current_month)
                )
                month_number = month_list.index(month) + 1
                year = st.number_input(
                    "Select Year:",
                    min_value=2025,
                    max_value=current_year,
                    value=current_year,
                    key="year1"
                )
                df = printing_data_month(month_number, year, region)

                if not df.empty:
                    st.subheader(f"🖨️ Printing Data for {month} {year} - {region}")
                    df = df[["Name", "Brand", "Project Manager", "Address", "Phone #", "Book", "Format", "Ink Type", "No of Copies", "Order Date", "Delivery Method", "Status", "Courier", "Tracking Number", "Shipping Date", "Fulfilled", "Type", "Accepted"]]
                    st.dataframe(df, width="stretch")

                else:
                    st.warning(f"No printing data found for {month} {year}")
            with tab_y:
                year2 = st.number_input(
                    "Select Year:",
                    min_value=2025,
                    max_value=current_year,
                    value=current_year,
                    key="year2"
                )
                df2 = printing_data_year(year2, region)

                if not df2.empty:
                    st.subheader(f"🖨️ Printing Data for {year2} - {region}")
                    df2 = df2[["Name", "Brand", "Project Manager", "Address", "Book", "Format", "Ink Type", "No of Copies", "Order Date",
                             "Delivery Method", "Status", "Courier", "Tracking Number", "Shipping Date", "Fulfilled",
                             "Type", "Accepted"]]
                    st.dataframe(df2, width="stretch")

                else:
                    st.warning(f"No printing data found for {year2}")

            with tab_u:
                df3 = get_printing_upcoming(region)

                if not df3.empty:
                    df3 = df3[["Name", "Brand", "Project Manager", "Address", "Book", "Format", "Ink Type",
                               "No of Copies",
                               "Delivery Method", "Status",
                               "Type"]]
                    st.dataframe(df3, width="stretch")
                else:
                    st.info("No upcoming printings ahead!")

            with tab_s:
                st.subheader(f"🔍 Search Data for {region}")

                number3 = st.number_input("Enter Year for Search", min_value=2025, step=1,
                                          value=current_year, key="year_search")

                if number3 and sheet_name:
                    data = printing_data_year(number3, region)

                    if data.empty:
                        st.warning(f"⚠️ No Data Available for {region} in {number3}")
                    else:
                        search_term = st.text_input("Search by Name", placeholder="Enter client name to search",
                                                    key="search_term").strip()

                        if search_term:
                            search_df = data[data['Name'].str.contains(search_term, case=False, na=False)]

                            if search_df.empty:
                                st.warning(f"⚠️ No results found for '{search_term}'")
                            else:
                                st.success(f"✅ Found {len(search_df)} result(s) for '{search_term}'")
                                search_df.index = range(1, len(search_df)+1)
                                st.dataframe(search_df[["Name", "Brand", "Project Manager", "Address", "Book", "Format", "Ink Type",
                               "No of Copies",
                               "Delivery Method", "Status",
                               "Type"]])
                        else:
                            st.info("👆 Enter a name above to search")

        elif action == "Send Pending Reviews":
            if not st.session_state.authenticated_admin:
                st.error("🚫 You do not have permission to access this section.")
                st.stop()

            st.header(f"📤 Send Pending Reviews - {region}")
            selected_pm = st.selectbox("Select Project Manager", list(pm_names.keys()))
            email = pm_names[selected_pm]
            df, _, _ = load_pending_reviews(sheet_name, selected_pm)

            if not df.empty:
                st.info(f"Found {len(df)} pending reviews for {selected_pm}")
                st.dataframe(df, width="stretch")
                if st.button("📨 Send to Slack", type="primary"):
                    with st.spinner("Sending..."):
                        success = send_pending_reviews_per_pm(selected_pm, sheet_name, email, channel)
                        if success:
                            st.success(f"✅ Successfully sent pending reviews to {selected_pm}")
                        else:
                            st.error("❌ Failed to send reviews")
            else:
                st.warning(f"No pending reviews found for {selected_pm}")

        elif action == "Send Attained Reviews":
            if not st.session_state.authenticated_admin:
                st.error("🚫 You do not have permission to access this section.")
                st.stop()

            st.header(f"📊 Send Attained Reviews - {region}")
            col1, col2 = st.columns(2)
            with col1:
                selected_pm = st.selectbox("Select Project Manager", list(pm_names.keys()))
            with col2:
                year = st.number_input("Select Year", min_value=2025, max_value=current_year, value=current_year)

            email = pm_names[selected_pm]
            df = load_attained_reviews(sheet_name, selected_pm, year)
            total_reviews = len(load_total_reviews(sheet_name, selected_pm, year))

            if not df.empty:
                total = total_reviews + len(df)
                percent = len(df) / total if total > 0 else 0
                col1, col2, col3, col4 = st.columns(4)
                col1.metric("✅ Attained", len(df))
                col2.metric("❓ Pending", total_reviews)
                col3.metric("🤵🏻 Total", total)
                col4.metric("🎯 Retention", f"{percent:.1%}")
                df["Publishing Date"] = pd.to_datetime(df["Publishing Date"], errors='coerce').dt.strftime("%d-%B-%Y")
                df = df[["Name", "Brand", "Publishing Date", "Status", "Trustpilot Review", "Trustpilot Review Date",
                         "Trustpilot Review Links"]]
                st.dataframe(df, width="stretch")
                if st.button("📨 Send to Slack", type="primary"):
                    with st.spinner("Sending..."):
                        success = send_attained_reviews_per_pm(selected_pm, email, sheet_name, year, channel)
                        if success:
                            st.success(f"✅ Successfully sent attained reviews to {selected_pm}")
                        else:
                            st.error("❌ Failed to send reviews")
            else:
                st.warning(f"No attained reviews found for {selected_pm}")


        elif action == "Bulk Send":
            if not st.session_state.authenticated_admin:
                st.error("🚫 You do not have permission to access this section.")
                st.stop()

            st.header(f"🚀 Bulk Send - {region}")
            send_type = st.radio("Select Type", ["Pending Reviews", "Attained Reviews"])

            if send_type == "Pending Reviews":
                if st.button("📤 Send to All PMs", type="primary"):
                    progress_bar = st.progress(0)
                    status_text = st.empty()
                    total_pms = len(pm_names)
                    success_count = 0
                    for idx, (name, email) in enumerate(pm_names.items()):
                        status_text.text(f"Processing {name}...")
                        success = send_pending_reviews_per_pm(name, sheet_name, email, channel)
                        if success:
                            success_count += 1
                        progress_bar.progress((idx + 1) / total_pms)
                    status_text.empty()
                    progress_bar.empty()
                    st.success(f"✅ Sent to {success_count}/{total_pms} project managers")

            else:
                year = st.number_input("Select Year", min_value=2025, max_value=current_year, value=current_year)
                if st.button("📤 Send to All PMs", type="primary"):
                    progress_bar = st.progress(0)
                    status_text = st.empty()
                    total_pms = len(pm_names)
                    success_count = 0
                    for idx, (name, email) in enumerate(pm_names.items()):
                        status_text.text(f"Processing {name}...")
                        success = send_attained_reviews_per_pm(name, email, sheet_name, year, channel)
                        if success:
                            success_count += 1
                        progress_bar.progress((idx + 1) / total_pms)
                    status_text.empty()
                    progress_bar.empty()
                    st.success(f"✅ Sent to {success_count}/{total_pms} project managers")
    elif st.session_state.get('authentication_status') is False:
        st.error('Username/password is incorrect')
        st.stop()
    elif st.session_state.get('authentication_status') is None:
        st.warning('Please enter your username and password')
        st.stop()
if __name__ == "__main__":
    main()
