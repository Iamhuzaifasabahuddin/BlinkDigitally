import calendar
import logging
import os
import tempfile
from datetime import datetime

import matplotlib.pyplot as plt
import pandas as pd
from dotenv import load_dotenv
from slack_sdk import WebClient
from slack_sdk.errors import SlackApiError

load_dotenv('Info.env')
SLACK_BOT_TOKEN = os.getenv("SLACK")
client = WebClient(token=SLACK_BOT_TOKEN)

# Sheet URLs
SPREADSHEET_ID = os.getenv("SPREADSHEET_ID")
sheet_usa = "USA"
sheet_uk = "UK"
url_usa = f"https://docs.google.com/spreadsheets/d/{SPREADSHEET_ID}/gviz/tq?tqx=out:csv&sheet={sheet_usa}"
url_uk = f"https://docs.google.com/spreadsheets/d/{SPREADSHEET_ID}/gviz/tq?tqx=out:csv&sheet={sheet_uk}"
url_printing = f"https://docs.google.com/spreadsheets/d/{SPREADSHEET_ID}/gviz/tq?tqx=out:csv&sheet=Printing"

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
    data_original = data[data["Publishing Date"].dt.month == current_month]
    data = data_original[
        (data_original["Project Manager"] == name) &
        ((data_original["Trustpilot Review"] == "Pending") | (data_original["Trustpilot Review"] == "Sent")) &
        (data_original["Brand"].isin(["BookMarketeers", "Writers Clique", "Authors Solution"])) &
        (data_original["Status"] == "Published")
        ]
    data = data.sort_values(by=["Publishing Date"], ascending=True)
    data.index = range(1, len(data) + 1)

    attained = len(
        data_original[(data_original["Trustpilot Review"] == "Attained") & (data_original["Project Manager"] == name)])
    total_reviews = len(data) + attained
    total_percentage = (attained / total_reviews)

    return data, total_percentage


name_usa = {
    "Aiza Ali": "aiza.ali@topsoftdigitals.pk",
    "Ahmed Asif": "ahmed.asif@topsoftdigitals.pk",
    "Shozab Hasan": "shozab.hasan@topsoftdigitals.pk",
    "Asad Waqas": "asad.waqas@topsoftdigitals.pk",
    "Shaikh Arsalan": "shaikh.arsalan@topsoftdigitals.pk",
    "Maheen Sami": "maheen.sami@topsoftdigitals.pk",
    "Mubashir Khan": "Mubashir.khan@topsoftdigitals.pk"
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

    df, percentage = load_data(url, name)

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
        f"Hi *{name.split()[0]}*! Here's your Trustpilot update for {current_month_name} 📄\n\n"
        f"*Summary:* {len(df)} pending reviews\n\n"
        f"*Review Retention:* {percentage:.1%}\n\n"
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


def printing():
    data = pd.read_csv(url_printing)
    columns = list(data.columns)
    end_col_index = columns.index("Fulfilled")
    data = data.iloc[:, :end_col_index + 1]

    for col in ["Order Date", "Shipping Date", "Fulfilled"]:
        data[col] = pd.to_datetime(data[col], errors="coerce")

    data = data[data["Order Date"].dt.month == current_month]
    data['Order Cost'] = pd.to_numeric(data['Order Cost'].str.replace('$', '', regex=False))

    return data


def summary():
    uk_clean = clean_data(url_uk)
    usa_clean = clean_data(url_usa)

    user_id = get_user_id_by_email("farmanali@topsoftdigitals.pk")

    usa_clean = usa_clean[usa_clean["Publishing Date"].dt.month == current_month]
    uk_clean = uk_clean[uk_clean["Publishing Date"].dt.month == current_month]

    usa_review = usa_clean["Trustpilot Review"].value_counts()
    uk_review = uk_clean["Trustpilot Review"].value_counts()

    usa_chart_path = generate_review_pie_chart(usa_review, "USA Trustpilot Reviews")
    uk_chart_path = generate_review_pie_chart(uk_review, "UK Trustpilot Reviews")

    usa_total = usa_review.sum()
    uk_total = uk_review.sum()

    usa_attained = usa_review.get('Attained', 0)
    uk_attained = uk_review.get('Attained', 0)

    usa_attained_pct = (usa_attained / usa_total * 100).round(1) if usa_total > 0 else 0
    uk_attained_pct = (uk_attained / uk_total * 100).round(1) if uk_total > 0 else 0

    combined_total = usa_total + uk_total
    combined_attained = usa_attained + uk_attained
    combined_attained_pct = (combined_attained / combined_total * 100).round(1) if combined_total > 0 else 0

    printing_ = printing()
    printing_data = printing_
    Total_copies = printing_data["No of Copies"].sum()
    Total_cost = printing_data["Order Cost"].sum()
    Highest_cost = printing_data["Order Cost"].max()
    Highest_copies = printing_data["No of Copies"].max()
    Lowest_cost = printing_data["Order Cost"].min()
    Lowest_copies = printing_data["No of Copies"].min()
    printing_data['Cost_Per_Copy'] = printing_data['Order Cost'] / printing_data['No of Copies']
    Average = Total_cost / Total_copies
    message = f"""
*{current_month_name} Trustpilot Reviews & Printing Summary*

*USA Reviews:*
• Total Reviews: {usa_total}
• Status Breakdown: {format_review_counts(usa_review)}
• Attained Percentage: {usa_attained_pct}%

*UK Reviews:*
• Total Reviews: {uk_total}
• Status Breakdown: {format_review_counts(uk_review)}
• Attained Percentage: {uk_attained_pct}%

*Combined Stats:*
• Total Reviews: {combined_total}
• Attained Reviews: {combined_attained} ({combined_attained_pct}%)

*Printing Stats:*
• Total Copies: {Total_copies}
• Total Cost: ${Total_cost}
• Highest Copies: {Highest_copies}
• Highest Cost: ${Highest_cost}
• Lowest Copies: {Lowest_copies}
• Lowest Cost: ${Lowest_cost}
• Average Cost: ${Average} per copy
    """

    try:
        conversation = client.conversations_open(users=user_id)
        channel_id = conversation['channel']['id']

        response = client.chat_postMessage(
            channel=channel_id,
            text=message,
            mrkdwn=True
        )

        client.files_upload_v2(
            channel=channel_id,
            file=usa_chart_path,
            title="USA Trustpilot Reviews"
        )

        client.files_upload_v2(
            channel=channel_id,
            file=uk_chart_path,
            title="UK Trustpilot Reviews"
        )

        send_dm(get_user_id_by_email("huzaifa.sabah@topsoftdigitals.pk"), f"✅ Review summary sent with charts")
    except SlackApiError as e:
        print(f"❌ Error sending message: {e.response['error']}")
        print(f"Detailed error: {str(e)}")
        logging.error(e)


def generate_review_pie_chart(review_counts, title):
    """Generate a pie chart for review counts and save it to a temporary file"""

    # Create temporary file for the chart
    temp_file = tempfile.NamedTemporaryFile(suffix='.png', delete=False)
    file_path = temp_file.name
    temp_file.close()

    # Calculate percentages
    total = review_counts.sum()
    percentages = (review_counts / total * 100).round(1)

    colors = {'Pending': 'red', 'Sent': 'purple', 'Attained': 'green'}
    chart_colors = [colors.get(status, '#cccccc') for status in review_counts.index]

    labels = [f"{status}\n({percent}%)" for status, percent in zip(review_counts.index, percentages)]

    # Create the pie chart
    plt.figure(figsize=(8, 6))
    plt.pie(review_counts, labels=labels, colors=chart_colors, autopct='', startangle=90, shadow=True)
    plt.title(title)

    # Add a legend with raw counts
    legend_labels = [f"{status}: {count}" for status, count in zip(review_counts.index, review_counts)]
    plt.legend(legend_labels, loc='best')

    plt.axis('equal')
    plt.tight_layout()

    plt.savefig(file_path)
    plt.close()

    return file_path


def format_review_counts(review_counts):
    """Format review counts as a string"""
    return ", ".join([f"{status}: {count}" for status, count in review_counts.items()])


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
    # for name, email in name_usa.items():
    #     send_df_as_text(name, url_usa, email)
    #
    # for name, email in names_uk.items():
    #     send_df_as_text(name, url_uk, email)
    summary()
