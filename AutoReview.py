import calendar
import logging
import os
from datetime import datetime

import pandas as pd
from dotenv import load_dotenv
from slack_sdk import WebClient
from slack_sdk.errors import SlackApiError

from App import current_month

load_dotenv('Info.env')
SLACK_BOT_TOKEN = os.getenv("SLACK")
client = WebClient(token=SLACK_BOT_TOKEN)

# Sheet URLs
SPREADSHEET_ID = os.getenv("SPREADSHEET_ID")
sheet_usa = "USA"
sheet_uk = "UK"
url_usa = f"https://docs.google.com/spreadsheets/d/{SPREADSHEET_ID}/gviz/tq?tqx=out:csv&sheet={sheet_usa}"
url_uk = f"https://docs.google.com/spreadsheets/d/{SPREADSHEET_ID}/gviz/tq?tqx=out:csv&sheet={sheet_uk}"

current_month = datetime.today().month
current_month_name = calendar.month_name[current_month]


def clean_data(url: str) -> pd.DataFrame:
    data = pd.read_csv(url)
    columns = list(data.columns)
    end_col_index = columns.index("Issues")
    data = data.iloc[:, :end_col_index + 1]

    for col in ["Publishing Date", "Last Edit (Revision)", "Trustpilot Review Date"]:
        data[col] = pd.to_datetime(data[col], errors="coerce")

    return data


def load_data(url, name):
    data = clean_data(url)
    data = data[data["Publishing Date"].dt.month == current_month]
    data = data[
        (data["Project Manager"] == name) &
        ((data["Trustpilot Review"] == "Pending") | (data["Trustpilot Review"] == "Sent")) &
        (data["Brand"].isin(["BookMarketeers", "Writers Clique", "Authors Solution"])) &
        (data["Status"] == "Published")
        ]

    # Sort by Trustpilot Review (Sent first), then by Publishing Date (oldest first)
    data["Review Sort"] = data["Trustpilot Review"].map({"Sent": 0, "Pending": 1})
    data = data.sort_values(by=["Review Sort", "Publishing Date"], ascending=[True, True])
    data = data.drop(columns="Review Sort")
    data.index = range(1, len(data) + 1)
    return data


name_usa = {
    "Aiza Ali": "aiza.ali@topsoftdigitals.pk",
    "Ahmed Asif": "ahmed.asif@topsoftdigitals.pk",
    "Shozab Hasan": "shozab.hasan@topsoftdigitals.pk",
    "Asad Waqas": "asad.waqas@topsoftdigitals.pk",
    "Shaikh Arsalan": "shaikh.arsalan@topsoftdigitals.pk",
    "Maheen Sami": "maheen.sami@topsoftdigitals.pk"
}

names_uk = {
    "Hadia Ghazanfar": "hadia.ghazanfar@topsoftdigitals.pk",
    "Youha": "youha.khan@topsoftdigitals.pk",
    "Syed Ahsan Shahzad": "ahsan.shahzad@topsoftdigitals.pk"
}

general_message = """Hiya
:bangbang: Please ask the following Clients for their feedback about their respective projects for the ones marked as pending & for those marked as Sent please remind the clients once again that their feedback truly matters and helps us grow and make essential changes to make the process even more fluid!
BM: https://bookmarketeers.com/
WC: https://writersclique.com/
AS: https://authorssolution.co.uk/"""


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


def send_df_as_text(name, url, email):
    user_id = get_user_id_by_email(email)

    if not user_id:
        print(f"❌ Could not find user ID for {name}")
        return

    df = load_data(url, name)

    if df.empty:
        print(f"⚠️ No data for {name}")
        return

    display_columns = ["Name", "Brand", "Book Name & Link", "Publishing Date", "Trustpilot Review"]
    if all(col in df.columns for col in display_columns):
        display_df = df[display_columns]
    else:
        display_df = df

    # Create a message with summary
    message = (
        f"{general_message}\n\n"
        f"Hi {name.split()[0]}! Here's your Trustpilot update for {current_month_name} 📄\n\n"
        f"*Summary:* {len(df)} pending reviews\n\n"
        f"```{display_df.to_markdown()}```"
    )

    try:
        # Open a DM channel
        conversation = client.conversations_open(users=user_id)
        channel_id = conversation['channel']['id']

        # Send the formatted text message
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
        send_df_as_text(name, url_usa, email)

    for name, email in names_uk.items():
        send_df_as_text(name, url_uk, email)
