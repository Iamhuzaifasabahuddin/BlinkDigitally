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
url_Audio = f"https://docs.google.com/spreadsheets/d/{SPREADSHEET_ID}/gviz/tq?tqx=out:csv&sheet=AudioBook"
url_printing = f"https://docs.google.com/spreadsheets/d/{SPREADSHEET_ID}/gviz/tq?tqx=out:csv&sheet=Printing"
url_copyright = f"https://docs.google.com/spreadsheets/d/{SPREADSHEET_ID}/gviz/tq?tqx=out:csv&sheet=Copyright"

current_month = datetime.today().month
# current_month = 4
current_month_name = calendar.month_name[current_month]
current_year = datetime.today().year


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
    # data_original = data[(data["Publishing Date"].dt.month == current_month) & (data["Publishing Date"].dt.year == current_year)]
    data_original = data
    data = data_original[
        (data_original["Project Manager"] == name) &
        ((data_original["Trustpilot Review"] == "Pending") | (data_original["Trustpilot Review"] == "Sent")) &
        (data_original["Brand"].isin(["BookMarketeers", "Writers Clique", "Authors Solution"])) &
        (data_original["Status"] == "Published")
        ]
    data = data.sort_values(by=["Publishing Date"], ascending=True)
    data.index = range(1, len(data) + 1)
    total_percentage = 0
    attained = len(
        data_original[(data_original["Trustpilot Review"] == "Attained") & (data_original["Project Manager"] == name)])
    total_reviews = len(data) + attained
    min_date = data["Publishing Date"].min()
    max_date = data["Publishing Date"].max()
    if total_reviews > 0:
        total_percentage = (attained / total_reviews)

    return data, total_percentage, min_date, max_date, attained, total_reviews


def load_data_audio(name):
    data = clean_data(url_Audio)
    data_original = data
    # data_original = data[
    #     (data["Publishing Date"].dt.month == current_month) & (data["Publishing Date"].dt.year == current_year)]
    data = data_original[
        (data_original["Project Manager"] == name) &
        ((data_original["Trustpilot Review"] == "Pending") | (data_original["Trustpilot Review"] == "Sent")) &
        (data_original["Brand"].isin(["BookMarketeers", "Writers Clique", "Authors Solution"])) &
        (data_original["Status"] == "Published")
        ]
    data = data.sort_values(by=["Publishing Date"], ascending=True)
    data.index = range(1, len(data) + 1)
    total_percentage = 0
    attained = len(
        data_original[(data_original["Trustpilot Review"] == "Attained") & (data_original["Project Manager"] == name)])
    total_reviews = len(data) + attained
    min_date = data["Publishing Date"].min()
    max_date = data["Publishing Date"].max()
    if total_reviews > 0:
        total_percentage = (attained / total_reviews)

    return data, total_percentage, min_date, max_date, attained, total_reviews


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
:bangbang: Please ask the following Clients for their feedback about their respective projects for the ones marked as _*Pending*_ & for those marked as _*Sent*_ please remind the clients once again that their feedback truly matters and helps us grow and make essential changes to make the process even more fluid!
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

    df, percentage, min_date, max_date, attained, total_reviews = load_data(url, name)
    df_audio, percentage_audio, min_date_audio, max_date_audio, attained_audio, total_reviews_audio = load_data_audio(
        name)

    if df.empty and df_audio.empty:
        print(f"⚠️ No data for {name}")
        return

    min_month_name = min(min_date, min_date_audio).strftime("%B")
    max_month_name = max(max_date, max_date_audio).strftime("%B")

    def truncate_title(x):
        return x[:60] + "..." if isinstance(x, str) and len(x) > 60 else x

    for dframe in [df, df_audio]:
        if "Book Name & Link" in dframe.columns:
            dframe["Book Name & Link"] = dframe["Book Name & Link"].apply(truncate_title)

    display_columns = ["Name", "Brand", "Book Name & Link", "Publishing Date", "Trustpilot Review"]
    display_df = df[display_columns] if all(col in df.columns for col in display_columns) else df
    display_df_audio = df_audio[display_columns] if all(
        col in df_audio.columns for col in display_columns) else df_audio

    for dframe in [display_df, display_df_audio]:
        if "Publishing Date" in dframe.columns:
            dframe["Publishing Date"] = pd.to_datetime(dframe["Publishing Date"], errors='coerce').dt.strftime(
                "%d-%B-%Y")

    merged_df = pd.concat([display_df, display_df_audio], ignore_index=True)
    if not merged_df.empty:
        markdown_table = merged_df.to_markdown(index=False)

        if len(set([min_month_name, max_month_name])) > 1:
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


def printing():
    data = pd.read_csv(url_printing)
    columns = list(data.columns)
    end_col_index = columns.index("Fulfilled")
    data = data.iloc[:, :end_col_index + 1]

    for col in ["Order Date", "Shipping Date", "Fulfilled"]:
        data[col] = pd.to_datetime(data[col], errors="coerce")

    data = data[(data["Order Date"].dt.month == current_month) & (data["Order Date"].dt.year == current_year)]
    data['Order Cost'] = pd.to_numeric(data['Order Cost'].str.replace('$', '', regex=False))

    return data


def Copyright():
    data = pd.read_csv(url_copyright)
    columns = list(data.columns)
    end_col_index = columns.index("Type")
    data = data.iloc[:, :end_col_index + 1]

    data["Submission Date"] = pd.to_datetime(data["Submission Date"], errors='coerce')
    data = data[(data["Submission Date"].dt.month == current_month) & (data["Submission Date"].dt.year == current_year)]

    data = data.sort_values(by=["Submission Date"], ascending=True)
    data.index = range(1, len(data) + 1)

    result_count = len(data[data["Result"] == "Yes"])
    data["Submission Date"] = data["Submission Date"].dt.strftime("%d-%B-%Y")

    return data, result_count


def summary():
    uk_clean = clean_data(url_uk)
    usa_clean = clean_data(url_usa)

    user_id = get_user_id_by_email("farmanali@topsoftdigitals.pk")

    usa_clean = usa_clean[
        (usa_clean["Publishing Date"].dt.month == current_month) & (
                usa_clean["Publishing Date"].dt.year == current_year)]
    uk_clean = uk_clean[
        (uk_clean["Publishing Date"].dt.month == current_month) & (uk_clean["Publishing Date"].dt.year == current_year)]

    if usa_clean.empty:
        print("No values found in USA sheet.")
    if uk_clean.empty:
        print("No values found in UK sheet.")
    if usa_clean.empty or uk_clean.empty:
        return

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

    copyright_, result_count = Copyright()
    copyright_data = copyright_
    Total_copyrights = len(copyright_data)
    Total_cost_copyright = Total_copyrights * 65

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
• Total Cost: ${Total_cost:.2f}
• Highest Copies: {Highest_copies}
• Highest Cost: ${Highest_cost:.2f}
• Lowest Copies: {Lowest_copies}
• Lowest Cost: ${Lowest_cost:.2f}
• Average Cost: ${Average:.2f} per copy

*Copyright Stats:*
• Total Copyrights: {Total_copyrights}
• Total Cost:${Total_cost_copyright}
• Total Successful:{result_count} / {Total_copyrights}
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
    for name, email in name_usa.items():
        send_df_as_text(name, url_usa, email)
    #
    # for name, email in names_uk.items():
    #     send_df_as_text(name, url_uk, email)
    # summary()
