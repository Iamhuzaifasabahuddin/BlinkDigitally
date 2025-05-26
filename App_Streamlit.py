import calendar
import logging
import tempfile
from datetime import datetime

import matplotlib.pyplot as plt
import pandas as pd
import streamlit as st
from slack_sdk import WebClient
from slack_sdk.errors import SlackApiError
from streamlit_gsheets import GSheetsConnection

st.set_page_config(page_title="Blink Digitally", page_icon="📊", layout="wide")
client = WebClient(token=st.secrets["Slack"]["Slack"])
conn = st.connection("gsheets", type=GSheetsConnection)

sheet_usa = "USA"
sheet_uk = "UK"
sheet_audio = "AudioBook"
sheet_printing = "Printing"
sheet_copyright = "Copyright"

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


def clean_data(data: pd.DataFrame) -> pd.DataFrame:
    """Clean and prepare the dataframe"""
    if data.empty:
        return pd.DataFrame()

    columns = list(data.columns)
    if "Issues" in columns:
        end_col_index = columns.index("Issues")
        data = data.iloc[:, :end_col_index + 1]
    data = data.astype(str)
    date_columns = ["Publishing Date", "Last Edit (Revision)", "Trustpilot Review Date"]
    for col in date_columns:
        if col in data.columns:
            data[col] = pd.to_datetime(data[col], errors="coerce")

    data = data.fillna("N/A")
    return data


def load_data(sheet_name, month_number, year) -> pd.DataFrame:
    """Load data from Google Sheets with optional month filtering"""
    try:
        data = conn.read(worksheet=sheet_name, ttl=0)
        data = clean_data(data)

        if "Publishing Date" in data.columns:
            data = data[(data["Publishing Date"].dt.month == month_number) & (data["Publishing Date"].dt.year == year)]

        data = data.sort_values(by="Publishing Date", ascending=True)

        for col in ["Publishing Date", "Last Edit (Revision)", "Trustpilot Review Date"]:
            if col in data.columns:
                data[col] = pd.to_datetime(data[col], errors="coerce").dt.strftime("%d-%B-%Y")

        data.index = range(1, len(data) + 1)
        data = data.fillna("N/A")
        return data
    except Exception as e:
        st.error(f"Error loading data: {e}")
        return pd.DataFrame()


# def add_data(sheet_name, new_row):
#     """Add a new row to the specified sheet"""
#     try:
#
#         current_data = conn.read(worksheet=sheet_name)
#
#         updated_data = pd.concat([current_data, pd.DataFrame([new_row], columns=current_data.columns)],
#                                  ignore_index=True)
#         conn.update(worksheet=sheet_name, data=updated_data, headers=False)
#         st.success("Data successfully added to the sheet.")
#         st.dataframe(load_data(sheet_name))
#     except Exception as e:
#         st.error(f"An error occurred: {e}")


def review_data(sheet_name, month, year, status) -> pd.DataFrame:
    """Filter data by month and review status"""
    data = load_data(sheet_name, month, year)
    if not data.empty and month and status:
        if "Publishing Date" in data.columns and "Trustpilot Review" in data.columns:
            data = data[(data["Publishing Date"].dt.month == month) & (data["Publishing Date"].dt.year == year)]
            data = data[data["Trustpilot Review"] == status]
        data = data.sort_values(by="Publishing Date", ascending=True)
    data.index = range(1, len(data) + 1)
    data = data.fillna("N/A")
    return data


def get_printing_data(month, year) -> pd.DataFrame:
    """Get printing data filtered by month"""
    try:
        data = conn.read(worksheet=sheet_printing, ttl=0)

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
            data["Order Cost"] = pd.to_numeric(data["Order Cost"].str.replace("$", "", regex=False), errors="coerce")

        data = data.sort_values(by="Order Date", ascending=True)

        for col in ["Order Date", "Shipping Date", "Fulfilled"]:
            if col in data.columns:
                data[col] = pd.to_datetime(data[col], errors="coerce").dt.strftime("%d-%B-%Y")

        data.index = range(1, len(data) + 1)
        data = data.fillna("N/A")

        return data
    except Exception as e:
        st.error(f"Error loading printing data: {e}")
        return pd.DataFrame()


def get_printing_data_reviews(month, year) -> pd.DataFrame:
    """Get printing data for the current month"""
    data = conn.read(worksheet=sheet_printing, ttl=0)

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
        data['Order Cost'] = pd.to_numeric(data['Order Cost'].str.replace('$', '', regex=False), errors='coerce')

    data = data.sort_values(by="Order Date", ascending=True)

    for col in ["Order Date", "Shipping Date", "Fulfilled"]:

        if col in data.columns:
            data[col] = pd.to_datetime(data[col], errors="coerce").dt.strftime("%d-%B-%Y")

    data.index = range(1, len(data) + 1)
    data = data.fillna("N/A")
    return data


def printing_data_all(year) -> pd.DataFrame:
    data = conn.read(worksheet=sheet_printing, ttl=0)

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
        data['Order Cost'] = pd.to_numeric(data['Order Cost'].str.replace('$', '', regex=False), errors='coerce')

    data = data.sort_values(by="Order Date", ascending=True)

    for col in ["Order Date", "Shipping Date", "Fulfilled"]:

        if col in data.columns:
            data[col] = pd.to_datetime(data[col], errors="coerce").dt.strftime("%d-%B-%Y")

    data.index = range(1, len(data) + 1)
    data = data.fillna("N/A")
    return data


def get_copyright_data(month, year) -> (pd.DataFrame, int):
    """Get copyright data for the current month"""
    data = conn.read(worksheet=sheet_copyright, ttl=0)

    columns = list(data.columns)
    if "Type" in columns:
        end_col_index = columns.index("Country")
        data = data.iloc[:, :end_col_index + 1]
    data = data.astype(str)

    if "Submission Date" in data.columns:
        data["Submission Date"] = pd.to_datetime(data["Submission Date"], errors='coerce')
        data = data[
            (data["Submission Date"].dt.month == month) & (data["Submission Date"].dt.year == year)]

    data = data.sort_values(by=["Submission Date"], ascending=True)
    result_count = len(data[data["Result"] == "Yes"]) if "Result" in data.columns else 0

    if "Submission Date" in data.columns:
        data["Submission Date"] = data["Submission Date"].dt.strftime("%d-%B-%Y")

    data = data.fillna("N/A")
    data.index = range(1, len(data) + 1)

    return data, result_count


def copyright_all(year) -> (pd.DataFrame, int):
    data = conn.read(worksheet=sheet_copyright, ttl=0)

    columns = list(data.columns)
    if "Type" in columns:
        end_col_index = columns.index("Country")
        data = data.iloc[:, :end_col_index + 1]
    data = data.astype(str)

    if "Submission Date" in data.columns:
        data["Submission Date"] = pd.to_datetime(data["Submission Date"], errors='coerce')
        data = data[
            (data["Submission Date"].dt.year == year)]
    data = data.sort_values(by=["Submission Date"], ascending=True)

    result_count = len(data[data["Result"] == "Yes"]) if "Result" in data.columns else 0

    if "Submission Date" in data.columns:
        data["Submission Date"] = data["Submission Date"].dt.strftime("%d-%B-%Y")

    data = data.fillna("N/A")
    data.index = range(1, len(data) + 1)

    return data, result_count


def clean_data_reviews(sheet_name: str) -> pd.DataFrame:
    """Clean the data from Google Sheets"""
    data = conn.read(worksheet=sheet_name, ttl=0)

    # Find the index of the "Issues" column if it exists
    columns = list(data.columns)
    if "Issues" in columns:
        end_col_index = columns.index("Issues")
        data = data.iloc[:, :end_col_index + 1]

    # Convert date columns to datetime
    for col in ["Publishing Date", "Last Edit (Revision)", "Trustpilot Review Date"]:
        if col in data.columns:
            data[col] = pd.to_datetime(data[col], errors="coerce")

    data = data.sort_values(by="Publishing Date", ascending=True)

    for col in ["Publishing Date", "Last Edit (Revision)", "Trustpilot Review Date"]:
        if col in data.columns:
            data[col] = pd.to_datetime(data[col], errors="coerce").dt.strftime("%d-%B-%Y")
    data.index = range(1, len(data) + 1)
    data = data.fillna("N/A")

    return data


def load_data_reviews(sheet_name, name) -> (pd.DataFrame, float, datetime, datetime, int):
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
    data = data.fillna("N/A")
    return data, total_percentage, min_date, max_date, attained, total_reviews


def load_data_audio(name) -> pd.DataFrame:
    """Load and filter audio book data for a specific project manager"""
    return load_data_reviews(sheet_audio, name)


def get_user_id_by_email(email):
    """Get Slack user ID by email"""
    try:
        response = client.users_lookupByEmail(email=email)
        return response['user']['id']
    except SlackApiError as e:
        print(f"Error finding user: {e.response['error']}")
        logging.error(e)
        return None


def send_dm(user_id, message) -> None:
    """Send a direct message to a user"""
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


def generate_year_summary(year) -> None:
    user_id = get_user_id_by_email("huzaifa.sabah@topsoftdigitals.pk")

    uk_clean = clean_data_reviews(sheet_uk)
    usa_clean = clean_data_reviews(sheet_usa)
    usa_clean = usa_clean[
        (usa_clean["Publishing Date"].dt.year == year)
    ]
    uk_clean = uk_clean[
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

    usa_review = usa_clean[
        "Trustpilot Review"].value_counts() if "Trustpilot Review" in usa_clean.columns else pd.Series()
    uk_review = uk_clean["Trustpilot Review"].value_counts() if "Trustpilot Review" in uk_clean.columns else pd.Series()

    usa_chart_path = generate_review_pie_chart(usa_review, "USA Trustpilot Reviews")
    uk_chart_path = generate_review_pie_chart(uk_review, "UK Trustpilot Reviews")

    usa_total = usa_review.sum() if not usa_review.empty else 0
    uk_total = uk_review.sum() if not uk_review.empty else 0

    usa_attained = usa_review.get('Attained', 0)
    uk_attained = uk_review.get('Attained', 0)

    usa_attained_pct = (usa_attained / usa_total * 100).round(1) if usa_total > 0 else 0
    uk_attained_pct = (uk_attained / uk_total * 100).round(1) if uk_total > 0 else 0

    combined_total = usa_total + uk_total
    combined_attained = usa_attained + uk_attained
    combined_attained_pct = (combined_attained / combined_total * 100).round(1) if combined_total > 0 else 0

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

    copyright_data, result_count = copyright_all(year)
    Total_copyrights = len(copyright_data)
    Total_cost_copyright = Total_copyrights * 65

    message = f"""
    *{current_year} Trustpilot Reviews & Printing Summary*

    *USA Reviews:*
    • Total Reviews: {usa_total}
    • Status Breakdown: {format_review_counts_reviews(usa_review)}
    • Attained Percentage: {usa_attained_pct}%

    *UK Reviews:*
    • Total Reviews: {uk_total}
    • Status Breakdown: {format_review_counts_reviews(uk_review)}
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
    • Total Cost: ${Total_cost_copyright}
    • Total Successful: {result_count} / {Total_copyrights}
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


def summary(month, year) -> None:
    """Generate and send summary report to management"""
    # Get the data
    uk_clean = clean_data_reviews(sheet_uk)
    usa_clean = clean_data_reviews(sheet_usa)

    user_id = get_user_id_by_email("farmanali@topsoftdigitals.pk")

    usa_clean = usa_clean[
        (usa_clean["Publishing Date"].dt.month == month) &
        (usa_clean["Publishing Date"].dt.year == year)
        ]
    uk_clean = uk_clean[
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

    usa_review = usa_clean[
        "Trustpilot Review"].value_counts() if "Trustpilot Review" in usa_clean.columns else pd.Series()
    uk_review = uk_clean["Trustpilot Review"].value_counts() if "Trustpilot Review" in uk_clean.columns else pd.Series()

    usa_chart_path = generate_review_pie_chart(usa_review, "USA Trustpilot Reviews")
    uk_chart_path = generate_review_pie_chart(uk_review, "UK Trustpilot Reviews")

    usa_total = usa_review.sum() if not usa_review.empty else 0
    uk_total = uk_review.sum() if not uk_review.empty else 0

    usa_attained = usa_review.get('Attained', 0)
    uk_attained = uk_review.get('Attained', 0)

    usa_attained_pct = (usa_attained / usa_total * 100).round(1) if usa_total > 0 else 0
    uk_attained_pct = (uk_attained / uk_total * 100).round(1) if uk_total > 0 else 0

    combined_total = usa_total + uk_total
    combined_attained = usa_attained + uk_attained
    combined_attained_pct = (combined_attained / combined_total * 100).round(1) if combined_total > 0 else 0

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

    copyright_data, result_count = get_copyright_data(month, year)
    Total_copyrights = len(copyright_data)
    Total_cost_copyright = Total_copyrights * 65

    message = f"""
*{current_month_name} Trustpilot Reviews & Printing Summary*

*USA Reviews:*
• Total Reviews: {usa_total}
• Status Breakdown: {format_review_counts_reviews(usa_review)}
• Attained Percentage: {usa_attained_pct}%

*UK Reviews:*
• Total Reviews: {uk_total}
• Status Breakdown: {format_review_counts_reviews(uk_review)}
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
• Total Cost: ${Total_cost_copyright}
• Total Successful: {result_count} / {Total_copyrights}
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

    # Handle empty data
    if review_counts.empty:
        plt.figure(figsize=(8, 6))
        plt.text(0.5, 0.5, "No data available", horizontalalignment='center', verticalalignment='center')
        plt.title(title)
        plt.axis('off')
        plt.savefig(file_path)
        plt.close()
        return file_path

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


def format_review_counts_reviews(review_counts):
    """Format review counts as a string"""
    if review_counts.empty:
        return "No data"
    return ", ".join([f"{status}: {count}" for status, count in review_counts.items()])


def get_min_year() -> int:
    """Gets Minimum year from the data"""
    uk_clean = clean_data_reviews(sheet_uk)
    audio = clean_data_reviews(sheet_audio)
    usa_clean = clean_data_reviews(sheet_usa)
    combined = pd.concat([uk_clean, usa_clean, audio])

    combined["Publishing Date"] = pd.to_datetime(combined["Publishing Date"], errors="coerce")

    min_year = combined["Publishing Date"].dt.year.min()

    return min_year


with st.container():
    st.title("📊 Data Management Portal")
    action = st.selectbox("What would you like to do?",
                          ["View Data", "Reviews", "Printing", "Copyright", "Generate Review & Summary",
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

    if action in ["View Data", "Reviews", "Printing", "Copyright"]:
        selected_month = st.selectbox(
            "Select Month",
            month_list,
            index=current_month - 1,
            placeholder="Select Month"
        )
        selected_month_number = month_list.index(selected_month) + 1 if selected_month else None
    if action in ["Year Summary", "Copyright", "Printing", "View Data", "Reviews"]:
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

            if data.empty:
                st.info(f"No data available for {selected_month} {number} for {choice}")
            else:
                st.markdown("### 📄 Detailed Entry Data")
                st.dataframe(data)

                brands = data["Brand"].value_counts()
                writers_clique = brands.get("Writers Clique", "N/A")
                bookmarketeers = brands.get("BookMarketeers", "N/A")
                kdp = brands.get("KDP", "N/A")
                authors_solution = brands.get("Authors Solution", "N/A")

                platforms = data["Platform"].value_counts()
                amazon = platforms.get("Amazon", "N/A")
                bn = platforms.get("Barnes & Noble", "N/A")
                ingram = platforms.get("Ingram Spark", "N/A")
                fav = platforms.get("FAV", "N/A")

                reviews = data["Trustpilot Review"].value_counts()
                total_reviews = reviews.sum()
                attained = reviews.get("Attained", 0)
                percentage = round((attained / total_reviews * 100), 1) if total_reviews > 0 else 0
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
                                - 📙 **Writers Clique:** `{writers_clique}`
                                - 📕 **KDP:** `{kdp}`
                                - 📘 **Authors Solution:** `{authors_solution}`
                                
                                **Platofrms**
                                - 🅰 **Amazon:** `{amazon}`
                                - 📔 **Barnes & Noble:** `{bn}`
                                - ⚡ **Ingram Spark:** `{ingram}`
                                - 🔉 **Findaway Voices:** `{fav}`
                                """)
                with col2:
                    st.markdown("---")

                    st.markdown("#### 🔍 Review Type Breakdown")
                    for review_type, count in reviews.items():
                        st.markdown(f"- 📝 **{review_type}**: `{count}`")
            st.markdown("---")
    # elif action == "Print Data" and country and selected_month:
    #     st.subheader(f"🖰 Print Data for {country} - {selected_month}")
    #     sheet_name = sheet_uk if country == "UK" else sheet_usa
    #
    #     data = load_data(sheet_name, selected_month_number, number)
    #
    #     for col in ["Publishing Date", "Last Edit (Revision)", "Trustpilot Review Date"]:
    #         if col in data.columns:
    #             data[col] = pd.to_datetime(data[col], errors="coerce").dt.strftime("%d-%B-%Y")
    #
    #     if not data.empty:
    #         st.dataframe(data)
    #     else:
    #         st.warning(f"No data available for {selected_month} for {country}")
    elif action == "Reviews" and choice and selected_month and status and number:
        sheet_name = {
            "UK": sheet_uk,
            "USA": sheet_usa,
            "AudioBook": sheet_audio
        }.get(choice)
        if sheet_name:
            data = review_data(sheet_name, selected_month_number, number, status)

            st.subheader(f"🔍 Review Data - {status} in {selected_month} ({country})")
            if not data.empty:
                st.dataframe(data)
            else:
                st.info(f"No matching reviews found for {selected_month_number} {number}")
    elif action == "Printing" and selected_month and number:

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
    elif action == "Copyright" and selected_month and number:
        st.subheader(f"© Copyright Summary for {selected_month}")

        data, result = get_copyright_data(selected_month_number, number)

        if not data.empty:
            st.dataframe(data)

            total_copyrights = len(data)
            total_cost_copyright = total_copyrights * 65
            country = data["Country"].value_counts()
            usa = country.get("USA", "N/A")
            canada = country.get("Canada", "N/A")
            st.markdown("---")
            st.markdown(f"""
            ### 📊 Summary Stats

            - 🧾 **Total Copyrighted Titles:** `{total_copyrights}`
            - 💵 **Copyright Total Cost:** `${total_cost_copyright}`
            - ✅ **Total Approved:** `{result} / {total_copyrights}`
            - 🦅 **USA:** `{usa}`
            - 🍁 **Canada:** `{canada}`
            """)
            st.markdown("---")
        else:
            st.warning(f"⚠️ No Data Available for Copyright in {selected_month} {number}")
    elif action == "Generate Review & Summary":

        tab1, tab2 = st.tabs(["Send Reviews", "Summary"])

        with tab1:
            st.header("Send Review Updates")

            st.subheader("🦅 USA Team")
            usa_selected = st.multiselect("Select USA team members:", list(name_usa.keys()))

            st.subheader("☕ UK Team")
            uk_selected = st.multiselect("Select UK team members:", list(names_uk.keys()))

            if st.button("Send Review Updates"):
                progress_bar = st.progress(0)
                total_members = len(usa_selected) + len(uk_selected)
                count = 0

                for name in usa_selected:
                    if name in name_usa:
                        send_df_as_text(name, sheet_usa, name_usa[name])
                        count += 1
                        progress_bar.progress(count / total_members)

                for name in uk_selected:
                    if name in names_uk:
                        send_df_as_text(name, sheet_uk, names_uk[name])
                        count += 1
                        progress_bar.progress(count / total_members)

                st.success(f"Sent review updates to {count} team members!")

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
                if st.button("Generate and Send Summary") and selected_month and number:
                    if no_data:
                        st.error(
                            f"Cannot generate summary — no data available for the month {selected_month} {number}.")
                    else:
                        with st.spinner(f"Generating summary report for {selected_month} {number}..."):
                            summary(selected_month_number, number)
                        st.success(f"Summary report for {selected_month} {number} generated and sent!")
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

            if st.button("Year Summary"):
                if no_data:
                    st.error(f"Cannot generate summary — no data available for the Year {number}.")
                else:
                    with st.spinner(f"Generating summary report for Year {number}..."):
                        generate_year_summary(number)
                    st.success(f"Summary report for Year {number} generated and sent!")
