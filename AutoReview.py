import calendar
import logging
import os
import tempfile
from datetime import datetime

import gspread
import matplotlib.pyplot as plt
import pandas as pd
from dotenv import load_dotenv
from google.oauth2.service_account import Credentials
from slack_sdk import WebClient
from slack_sdk.errors import SlackApiError

load_dotenv('Info.env')
SLACK_BOT_TOKEN = os.getenv("SLACK")
client = WebClient(token=SLACK_BOT_TOKEN)
channel_usa = os.getenv("CHANNEL_USA")
channel_uk = os.getenv("CHANNEL_UK")
# Google Sheets setup with gspread
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
        (data_original["Brand"].isin(
            ["BookMarketeers", "Writers Clique", "Authors Solution", "Book Publication", "Aurora Writers"])) &
        (data_original["Status"] == "Published")
        ]

    data = data_original[
        (data_original["Project Manager"] == name) &
        # ((data_original["Trustpilot Review"] == "Pending") | (data_original["Trustpilot Review"] == "Sent")) &
        ((data_original["Trustpilot Review"] == "Pending")) &
        (data_original["Brand"].isin(
            ["BookMarketeers", "Writers Clique", "Authors Solution", "Book Publication", "Aurora Writers"])) &
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
    return data, total_percentage, min_date, max_date, attained, total_reviews


def load_data_audio(name) -> tuple:
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


def send_df_as_text(name, sheet_name, email, channel) -> None:
    """Send DataFrame as text to a user"""
    user_id = get_user_id_by_email(email)
    # user_id = get_user_id_by_email("huzaifa.sabah@topsoftdigitals.pk")

    if not user_id:
        print(f"❌ Could not find user ID for {name}")
        return

    df, percentage, min_date, max_date, attained, total_reviews = load_data_reviews(sheet_name, name)

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

    display_columns = ["Name", "Brand", "Book Name & Link", "Publishing Date", "Trustpilot Review"]
    merged_df = df[[col for col in display_columns if col in df.columns]]

    if not merged_df.empty:
        markdown_table = merged_df.to_markdown(index=False)

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


def load_total_reviews(sheet_name: str, name):
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
        (data_original["Brand"].isin(
            ["BookMarketeers", "Writers Clique", "Authors Solution", "Book Publication", "Aurora Writers"])) &
        (data_original["Status"] == "Published")
        ]

    data = data_original[
        (data_original["Project Manager"] == name) &
        # ((data_original["Trustpilot Review"] == "Pending") | (data_original["Trustpilot Review"] == "Sent")) &
        ((data_original["Trustpilot Review"] == "Pending")) &
        (data_original["Brand"].isin(
            ["BookMarketeers", "Writers Clique", "Authors Solution", "Book Publication", "Aurora Writers"])) &
        (data_original["Status"] == "Published")
        ]

    data = data.sort_values(by="Publishing Date", ascending=True)

    # Clean strings and drop missing
    data_original["Trustpilot Review"] = data_original["Trustpilot Review"].astype(str).str.strip().str.lower()
    data_original["Project Manager"] = data_original["Project Manager"].astype(str).str.strip()
    name = name.strip()

    data_original = data_original.dropna(subset=["Trustpilot Review", "Project Manager"])
    total_reviews = len(data_count)

    min_date = data["Publishing Date"].min() if not data.empty else pd.NaT
    max_date = data["Publishing Date"].max() if not data.empty else pd.NaT

    data.index = range(1, len(data) + 1)
    return data, total_reviews


def load_reviews(sheet_name, name, year, month_number=None) -> pd.DataFrame:
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

        data_original = data
        data = data_original[
            (data_original["Project Manager"] == name) &
            ((data_original["Trustpilot Review"] == "Attained")) &
            (data_original["Brand"].isin(
                ["BookMarketeers", "Writers Clique", "Authors Solution", "Book Publication", "Aurora Writers"]))
            ]

        data = data.sort_values(by="Trustpilot Review Date", ascending=True)
        data = data.drop_duplicates(subset=["Name"])
        data.index = range(1, len(data) + 1)
        return data
    except Exception as e:
        return pd.DataFrame()

def send_pm_attained_reviews(pm_name, email, sheet_name, year, channel, month=None) -> None:
    """Send attained Trustpilot reviews for a specific Project Manager"""
    user_id = get_user_id_by_email(email)
    # user_id = get_user_id_by_email("huzaifa.sabah@topsoftdigitals.pk")
    if not user_id:
        print(f"❌ Could not find user ID for {pm_name}")
        return
    df, total_reviews = load_total_reviews(sheet_name, pm_name)
    review_data = load_reviews(sheet_name, name,  year, month)

    review_details_df = review_data
    review_details_df["Trustpilot Review Date"] = pd.to_datetime(
        review_details_df["Trustpilot Review Date"], errors="coerce"
    ).dt.strftime("%d-%B-%Y")

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
        send_df_as_text(name, sheet_usa, email, channel_usa)
        # send_pm_attained_reviews(name, email, sheet_usa, 2025, channel_usa)

    # for name, email in names_uk.items():
    #     send_df_as_text(name, sheet_uk, email, channel_uk)
    #     # send_pm_attained_reviews(name, email, sheet_uk, 2025, channel_uk)

    # summary(5, 2025)
    # generate_year_summary(2025)
