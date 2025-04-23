import calendar
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
        ((data["Trustpilot Review"] == "Pending") | (data["Trustpilot Review"] == "Sent"))
        & ((data["Brand"] == "BookMarketeers") | (data["Brand"] == "Writers Clique") | (data["Brand"] == "Authors Solution"))
    ]
    print(data.head())
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
    "Youha Khan": "youha.khan@topsoftdigitals.pk",
    "Syed Ahsan Shahzad": "ahsan.shahzad@topsoftdigitals.pk"
}


def get_user_id_by_email(email):
    try:
        response = client.users_lookupByEmail(email=email)
        return response['user']['id']
    except SlackApiError as e:
        print(f"Error finding user: {e.response['error']}")
        return None


def send_dm(user_id, message):
    response = client.auth_test()
    print(response)
    try:
        response = client.chat_postMessage(
            channel=user_id,
            text=message
        )
        print("✅ Message sent!")
    except SlackApiError as e:
        print(f"❌ Error sending message: {e.response['error']}")

def send_df_as_text(name, url, email=None):
    user_email = "huzaifa.sabah@topsoftdigitals.pk"
    user_id = get_user_id_by_email(user_email)

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


    for col in display_df.columns:
        if pd.api.types.is_datetime64_any_dtype(display_df[col]):
            display_df[col] = display_df[col].dt.strftime('%d-%B-%Y')

    header = "| " + " | ".join(display_df.columns) + " |"
    separator = "| " + " | ".join(["---"] * len(display_df.columns)) + " |"

    rows = []
    for _, row in display_df.iterrows():
        formatted_row = [str(val) if pd.notna(val) else "" for val in row]
        rows.append("| " + " | ".join(formatted_row) + " |")

    markdown_table = "\n".join([header, separator] + rows)

    # Create a message with summary
    message = (
        f"Hi {name.split()[0]}! Here's your Trustpilot update for {current_month_name} 📄\n\n"
        f"*Summary:* {len(df)} pending reviews\n\n"
        f"```{markdown_table}```"
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
        print(f"✅ Data sent to {name}")
    except SlackApiError as e:
        print(f"❌ Error sending message to {name}: {e.response['error']}")
        print(f"Detailed error: {str(e)}")
for name, email in name_usa.items():
    send_df_as_text(name, url_usa)
