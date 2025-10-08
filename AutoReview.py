import calendar
import logging
import os
from datetime import datetime

import gspread
import pandas as pd
from dotenv import load_dotenv
from google.oauth2.service_account import Credentials
from slack_sdk import WebClient
from slack_sdk.errors import SlackApiError

load_dotenv('.testing/Info.env')
SLACK_BOT_TOKEN = os.getenv("SLACK")
client = WebClient(token=SLACK_BOT_TOKEN)
channel_usa = os.getenv("CHANNEL_USA")
channel_uk = os.getenv("CHANNEL_UK")
# Google Sheets setup with gspread
SPREADSHEET_ID = os.getenv("SPREADSHEET_ID")

SCOPES = ["https://www.googleapis.com/auth/spreadsheets.readonly"]


def get_gspread_client():
    """Initialize and return gspread client with service account credentials"""
    try:
        credentials = Credentials.from_service_account_file(".testing/credentials.json", scopes=SCOPES)
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

name_usa = {
    "Aiza Ali": "aiza.ali@topsoftdigitals.pk",
    "Ahmed Asif": "ahmed.asif@topsoftdigitals.pk",
    "Asad Waqas": "asad.waqas@topsoftdigitals.pk",
    "Maheen Sami": "maheen.sami@topsoftdigitals.pk",
    "Mubashir Khan": "Mubashir.khan@topsoftdigitals.pk",
    "Muhammad Ali": "muhammad.ali@topsoftdigitals.pk",
    "Valencia Angelo": "valencia.angelo@topsoftdigitals.pk",
    "Ukasha Asadullah": "ukasha.asadullah@topsoftdigitals.pk"
}

names_uk = {
    "Youha": "youha.khan@topsoftdigitals.pk",
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

current_month = datetime.today().month
# current_month = 4
current_month_name = calendar.month_name[current_month]
current_year = datetime.today().year


def normalize_name(name):
    """Normalize a name to consistent format (Title Case, stripped whitespace)"""
    if pd.isna(name) or name == "":
        return ""
    return str(name).strip().title()


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

    if "Project Manager" in data.columns:
        data["Project Manager"] = data["Project Manager"].apply(normalize_name)

    # Handle date columns
    for col in ["Publishing Date", "Last Edit (Revision)", "Trustpilot Review Date"]:
        if col in data.columns:
            data[col] = pd.to_datetime(data[col], format="%d-%B-%Y", errors="coerce")

    data = data.sort_values(by="Publishing Date", ascending=True)
    data.index = range(1, len(data) + 1)

    return data


def load_pending_reviews(sheet_name: str, name: str) -> tuple:
    """Load and filter data for a specific project manager"""
    data = clean_data_reviews(sheet_name)

    # Normalize the input name for consistent comparison
    name = normalize_name(name)

    data_original = data.copy()
    if data.empty:
        return pd.DataFrame(), 0, pd.NaT, pd.NaT, 0, 0

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

        data_original = data.copy()
        data = data_original[
            (data_original["Project Manager"] == name) &
            ((data_original["Trustpilot Review"] == "Attained")) &
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
        return pd.DataFrame()


def load_total_reviews(sheet_name: str, name: str, year: int, month_number=None):
    data = clean_data_reviews(sheet_name)

    name = normalize_name(name)

    if "Name" in data.columns:
        data = data.drop_duplicates(subset=["Name"], keep="last")

    if data.empty:
        return pd.DataFrame(), 0, pd.NaT, pd.NaT, 0, 0
    date_columns = ["Publishing Date", "Last Edit (Revision)", "Trustpilot Review Date"]
    for col in date_columns:
        if col in data.columns:
            data[col] = pd.to_datetime(data[col], format="%d-%B-%Y", errors="coerce")

    try:
        if "Trustpilot Review Date" in data.columns and month_number:
            data = data[(data["Trustpilot Review Date"].dt.month == month_number) & (
                    data["Trustpilot Review Date"].dt.year == year)]
        else:
            data = data[(data["Publishing Date"].dt.year == year)]
        data_original = data.copy()
        data_count = data_original[
            (data_original["Project Manager"] == name) &
            ((data_original["Trustpilot Review"] == "Pending") | (data_original["Trustpilot Review"] == "Sent")) &
            (data_original["Brand"].isin(
                ["BookMarketeers", "Writers Clique", "Authors Solution", "Book Publication", "Aurora Writers"])) &
            (data_original["Status"] == "Published")
            ]
        total_reviews = len(data_count)
        return total_reviews
    except Exception as e:
        print(f"Error {e}")
        return pd.DataFrame()


def get_user_id_by_email(email: str):
    try:
        response = client.users_lookupByEmail(email=email)
        return response['user']['id']
    except SlackApiError as e:
        print(f"Error finding user: {e.response['error']}")
        logging.error(e)
        return None


def send_dm(user_id: str, message: str):
    try:
        response = client.chat_postMessage(
            channel=user_id,
            text=message
        )
    except SlackApiError as e:
        print(f"❌ Error sending message: {e.response['error']}")
        logging.error(e)


def send_pending_reviews_per_pm(name: str, sheet_name: str, email: str, channel: str) -> None:
    """Send DataFrame as text to a user"""
    user_id = get_user_id_by_email(email)
    # user_id = get_user_id_by_email("huzaifa.sabah@topsoftdigitals.pk")

    if not user_id:
        print(f"❌ Could not find user ID for {name}")
        return

    df, min_date, max_date = load_pending_reviews(sheet_name, name)

    if df.empty:
        print(f"⚠️ No data for {name}")
        return

    min_dates = [d for d in [min_date] if pd.notna(d)]
    max_dates = [d for d in [max_date] if pd.notna(d)]

    if not min_dates or not max_dates:
        print(f"⚠️ No valid dates for {name}")
        return

    min_month_name = min(min_dates).strftime("%B")
    max_month_name = max(max_dates).strftime("%B")

    def truncate_title(x):
        """Truncate long titles"""
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
            response = client.chat_postMessage(
                channel=channel,
                text=message,
                mrkdwn=True
            )
            send_dm(get_user_id_by_email("huzaifa.sabah@topsoftdigitals.pk"), f"✅ Review sent to {name}")
        except SlackApiError as e:
            print(f"❌ Error sending message to {name}: {e.response['error']}")
            print(f"Detailed error: {str(e)}")
            logging.error(e)


def send_attained_reviews_per_pm(pm_name: str, email: str, sheet_name: str, year: int, channel: str,
                                 month=None) -> None:
    """Send attained Trustpilot reviews for a specific Project Manager"""
    user_id = get_user_id_by_email(email)
    # user_id = get_user_id_by_email("huzaifa.sabah@topsoftdigitals.pk")
    if not user_id:
        print(f"❌ Could not find user ID for {pm_name}")
        return
    total_reviews = load_total_reviews(sheet_name, pm_name, year, month)
    review_data = load_attained_reviews(sheet_name, pm_name, year, month)

    review_details_df = review_data.copy()
    attained_details = review_details_df[["Name", "Brand", "Trustpilot Review Date"]]

    attained_details.index = range(1, len(attained_details) + 1)

    if attained_details.empty:
        print(f"⚠️ No attained reviews found for {pm_name}")
        return

    markdown_table = attained_details.to_markdown(index=True)

    Total = total_reviews + len(attained_details)
    percentage = len(attained_details) / Total
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
        client.chat_postMessage(
            channel=channel,
            text=message,
            mrkdwn=True
        )
        send_dm(get_user_id_by_email("huzaifa.sabah@topsoftdigitals.pk"),
                f"✅ Attained review details sent for {pm_name}")
        print(f"✅ Attained review details sent for {pm_name}")

    except SlackApiError as e:
        print(f"❌ Error sending message for {pm_name}: {e.response['error']}")
        logging.error(e)


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


if __name__ == '__main__':
    for name, email in name_usa.items():
        send_pending_reviews_per_pm(name, sheet_usa, email, channel_usa)
        # send_attained_reviews_per_pm(name, email, sheet_usa, 2025, channel_usa)

    # for name, email in names_uk.items():
    #     send_pending_reviews_per_pm(name, sheet_uk, email, channel_uk)
    #     # send_attained_reviews_per_pm(name, email, sheet_uk, 2025, channel_uk)

    # summary(5, 2025)
    # generate_year_summary(2025)
