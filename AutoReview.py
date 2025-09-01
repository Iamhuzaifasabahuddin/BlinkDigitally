import calendar
import logging
import os
import tempfile
import time
from datetime import datetime
import matplotlib.pyplot as plt
import pandas as pd
import gspread
from google.oauth2.service_account import Credentials
from dotenv import load_dotenv
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
    "Hadia Ghazanfar": "hadia.ghazanfar@topsoftdigitals.pk",
    "Youha": "youha.khan@topsoftdigitals.pk",
    "Emaan Zaidi": "emaan.zaidi@topsoftdigitals.pk",
    "Elishba": "elishba@topsoftdigitals.pk",
    "Shahrukh Yousuf": "shahrukh.yousuf@topsoftdigitals.pk"
}

general_message = """Hiya
:bangbang: Please ask the following Clients for their feedback about their respective projects for the ones marked as _*Pending*_ & for those marked as _*Sent*_ please remind the clients once again that their feedback truly matters and helps us grow and make essential changes to make the process even more fluid!
BM: https://bookmarketeers.com/
WC: https://writersclique.com/
AS: https://authorssolution.co.uk/"""

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
        (data_original["Brand"].isin(["BookMarketeers", "Writers Clique", "Authors Solution", "Book Publication"])) &
        (data_original["Status"] == "Published")
        ]



    data = data_original[
        (data_original["Project Manager"] == name) &
        # ((data_original["Trustpilot Review"] == "Pending") | (data_original["Trustpilot Review"] == "Sent")) &
        ((data_original["Trustpilot Review"] == "Pending")) &
        (data_original["Brand"].isin(["BookMarketeers", "Writers Clique", "Authors Solution", "Book Publication"])) &
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
    df_audio, percentage_audio, min_date_audio, max_date_audio, attained_audio, total_reviews_audio = load_data_audio(
        name)

    if df.empty and df_audio.empty:
        print(f"⚠️ No data for {name}")
        return

    min_dates = [d for d in [min_date, min_date_audio] if pd.notna(d)]
    max_dates = [d for d in [max_date, max_date_audio] if pd.notna(d)]

    if not min_dates or not max_dates:
        print(f"⚠️ No valid dates for {name}")
        return

    min_month_name = min(min_dates).strftime("%B")
    max_month_name = max(max_dates).strftime("%B")

    def truncate_title(x):
        """Truncate long titles"""
        return x[:20] + "..." if isinstance(x, str) and len(x) > 20 else x

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
                f"Hi <@{user_id}>!  Here's your Trustpilot update from {min_month_name} to {max_month_name} {current_year} 📄\n\n"
                f"*Summary:* {len(merged_df)} pending reviews\n\n"
                f"*Review Retention:* {attained + attained_audio} out of {total_reviews + total_reviews_audio} "
                f"({((attained + attained_audio) / (total_reviews + total_reviews_audio)):.1%})\n\n"
                f"```\n{markdown_table}\n```"
            )
        else:
            message = (
                f"{general_message}\n\n"
                f"Hi <@{user_id}>! Here's your Trustpilot update for {min_month_name} {current_year} 📄\n\n"
                f"*Summary:* {len(merged_df)} pending reviews\n\n"
                f"*Review Retention:* {attained + attained_audio} out of {total_reviews + total_reviews_audio} "
                f"({((attained + attained_audio) / (total_reviews + total_reviews_audio)):.1%})\n\n"
                f"```\n{markdown_table}\n```"
            )

        try:


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


def get_printing_data_reviews(month, year) -> pd.DataFrame:
    """Get printing data for the current month using gspread"""
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
        data['Order Cost'] = pd.to_numeric(data['Order Cost'].str.replace('$', '', regex=False), errors='coerce')

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
    """Get all printing data for a specific year using gspread"""
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
        data['Order Cost'] = pd.to_numeric(data['Order Cost'].str.replace('$', '', regex=False), errors='coerce')

    if "No of Copies" in data.columns:
        data["No of Copies"] = pd.to_numeric(data["No of Copies"], errors='coerce')

    data = data.sort_values(by="Order Date", ascending=True)

    for col in ["Order Date", "Shipping Date", "Fulfilled"]:
        if col in data.columns:
            data[col] = pd.to_datetime(data[col], errors="coerce").dt.strftime("%d-%B-%Y")

    data.index = range(1, len(data) + 1)
    data = data.fillna("N/A")
    return data



def get_copyright_data(month, year) -> tuple:
    """Get copyright data for the current month using gspread"""
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

    if "Submission Date" in data.columns:
        data["Submission Date"] = data["Submission Date"].dt.strftime("%d-%B-%Y")

    data = data.fillna("N/A")
    data.index = range(1, len(data) + 1)

    return data, result_count


def copyright_all(year) -> tuple:
    """Get all copyright data for a specific year using gspread"""
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

    if "Submission Date" in data.columns:
        data["Submission Date"] = data["Submission Date"].dt.strftime("%d-%B-%Y")

    data = data.fillna("N/A")
    data.index = range(1, len(data) + 1)

    return data, result_count


def summary(month, year) -> None:
    """Generate and send summary report to management"""
    uk_clean = clean_data_reviews(sheet_uk)
    usa_clean = clean_data_reviews(sheet_usa)

    # user_id = get_user_id_by_email("farmanali@topsoftdigitals.pk")
    user_id = get_user_id_by_email("huzaifa.sabah@topsoftdigitals.pk")

    if usa_clean.empty and uk_clean.empty:
        print("No data found in either USA or UK sheets.")
        return

    if not usa_clean.empty:
        usa_clean = usa_clean[
            (usa_clean["Publishing Date"].dt.month == month) &
            (usa_clean["Publishing Date"].dt.year == year)
            ]

    if not uk_clean.empty:
        uk_clean = uk_clean[
            (uk_clean["Publishing Date"].dt.month == month) &
            (uk_clean["Publishing Date"].dt.year == year)
            ]

    if usa_clean.empty:
        print("No values found in USA sheet for the specified period.")
    if uk_clean.empty:
        print("No values found in UK sheet for the specified period.")

    if usa_clean.empty and uk_clean.empty:
        return

    # Process USA data
    brands = usa_clean["Brand"].value_counts() if not usa_clean.empty else pd.Series()
    writers_clique = brands.get("Writers Clique", "N/A")
    bookmarketeers = brands.get("BookMarketeers", "N/A")
    kdp = brands.get("KDP", "N/A")

    usa_platforms = usa_clean["Platform"].value_counts() if not usa_clean.empty else pd.Series()
    usa_amazon = usa_platforms.get("Amazon", 0)
    usa_bn = usa_platforms.get("Barnes & Noble", 0)
    usa_ingram = usa_platforms.get("Ingram Spark", 0)

    # Process UK data
    uk_brand = uk_clean["Brand"].value_counts() if not uk_clean.empty else pd.Series()
    authors_solution = uk_brand.get("Authors Solution", "N/A")

    uk_platforms = uk_clean["Platform"].value_counts() if not uk_clean.empty else pd.Series()
    uk_amazon = uk_platforms.get("Amazon", 0)
    uk_bn = uk_platforms.get("Barnes & Noble", 0)
    uk_ingram = uk_platforms.get("Ingram Spark", 0)

    combined_amazon = int(usa_amazon) + int(uk_amazon)
    combined_bn = int(usa_bn) + int(uk_bn)
    combined_ingram = int(usa_ingram) + int(uk_ingram)

    usa_review = usa_clean[
        "Trustpilot Review"].value_counts() if "Trustpilot Review" in usa_clean.columns and not usa_clean.empty else pd.Series()
    uk_review = uk_clean[
        "Trustpilot Review"].value_counts() if "Trustpilot Review" in uk_clean.columns and not uk_clean.empty else pd.Series()

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
    country = copyright_data["Country"].value_counts() if not copyright_data.empty else pd.Series()
    usa = country.get("USA", "N/A")
    canada = country.get("Canada", "N/A")

    message = f"""
*{current_month_name} {year} Trustpilot Reviews & Printing Summary*

*USA Reviews:*
• Total Reviews: {usa_total}
• Status Breakdown: {format_review_counts_reviews(usa_review)}
• Attained Percentage: {usa_attained_pct}%
    *Brands*
    - 📘 *BookMarketeers:* `{bookmarketeers}`
    - 📙 *Writers Clique:* `{writers_clique}`
    - 📕 *KDP:* `{kdp}`

    *Platforms*
    - 🅰 *Amazon:* `{usa_amazon}`
    - 📔 *Barnes & Noble:* `{usa_bn}`
    - ⚡ *Ingram Spark:* `{usa_ingram}`

*UK Reviews:*
    • Total Reviews: {uk_total}
    • Status Breakdown: {format_review_counts_reviews(uk_review)}
    • Attained Percentage: {uk_attained_pct}%

    *Brand*
    - 📘 *Authors Solution:* `{authors_solution}`

    *Platforms*
    - 🅰 *Amazon:* `{uk_amazon}`
    - 📔 *Barnes & Noble:* `{uk_bn}`
    - ⚡ *Ingram Spark:* `{uk_ingram}`

*Combined Stats:*
    • Total Reviews: {combined_total}
    • Attained Reviews: {combined_attained} ({combined_attained_pct}%)
    • Platform Totals:
      - 🅰 *Amazon:* `{combined_amazon}`
      - 📔 *Barnes & Noble:* `{combined_bn}`
      - ⚡ *Ingram Spark:* `{combined_ingram}`

*Printing Stats:*
•🧾 Total Copies: {Total_copies}
•💰 Total Cost: ${Total_cost:.2f}
•📈 Highest Cost: ${Highest_cost:.2f}
•📉 Lowest Cost: ${Lowest_cost:.2f}
•🔢 Highest Copies: {Highest_copies}
•🧮 Lowest Copies: {Lowest_copies}
•🧾 Average Cost: ${Average:.2f} per copy

*Copyright Stats:*
    •🧾 Total Copyrights: {Total_copyrights}
    •💵 Total Cost: ${Total_cost_copyright}
    •✅ Total Successful: {result_count} / {Total_copyrights}
    • 🦅 *USA:* `{usa}`
    • 🍁 *Canada:* `{canada}`
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


def generate_year_summary(year) -> None:
    # user_id = get_user_id_by_email("farmanali@topsoftdigitals.pk")
    user_id = get_user_id_by_email("huzaifa.sabah@topsoftdigitals.pk")

    uk_clean = clean_data_reviews(sheet_uk)
    usa_clean = clean_data_reviews(sheet_usa)

    if not usa_clean.empty:
        usa_clean = usa_clean[
            (usa_clean["Publishing Date"].dt.year == year)
        ]
    if not uk_clean.empty:
        uk_clean = uk_clean[
            (uk_clean["Publishing Date"].dt.year == year)
        ]

    if usa_clean.empty:
        print("No values found in USA sheet.")
    if uk_clean.empty:
        print("No values found in UK sheet.")
    if usa_clean.empty and uk_clean.empty:
        return

    usa_review = usa_clean[
        "Trustpilot Review"].value_counts() if "Trustpilot Review" in usa_clean.columns and not usa_clean.empty else pd.Series()
    uk_review = uk_clean[
        "Trustpilot Review"].value_counts() if "Trustpilot Review" in uk_clean.columns and not uk_clean.empty else pd.Series()

    brands = usa_clean["Brand"].value_counts() if not usa_clean.empty else pd.Series()
    writers_clique = brands.get("Writers Clique", 0)
    bookmarketeers = brands.get("BookMarketeers", 0)
    kdp = brands.get("KDP", 0)

    uk_brand = uk_clean["Brand"].value_counts() if not uk_clean.empty else pd.Series()
    authors_solution = uk_brand.get("Authors Solution", 0)

    usa_platforms = usa_clean["Platform"].value_counts() if not usa_clean.empty else pd.Series()
    usa_amazon = usa_platforms.get("Amazon", 0)
    usa_bn = usa_platforms.get("Barnes & Noble", 0)
    usa_ingram = usa_platforms.get("Ingram Spark", 0)

    uk_platforms = uk_clean["Platform"].value_counts() if not uk_clean.empty else pd.Series()
    uk_amazon = uk_platforms.get("Amazon", 0)
    uk_bn = uk_platforms.get("Barnes & Noble", 0)
    uk_ingram = uk_platforms.get("Ingram Spark", 0)

    combined_amazon = int(usa_amazon) + int(uk_amazon)
    combined_bn = int(usa_bn) + int(uk_bn)
    combined_ingram = int(usa_ingram) + int(uk_ingram)

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
    country = copyright_data["Country"].value_counts()
    usa = country.get("USA", "N/A")
    canada = country.get("Canada", "N/A")

    message = f"""
        *{year} Trustpilot Reviews & Printing Summary*

        *USA Reviews:*
        • Total Reviews: {usa_total}
        • Status Breakdown: {format_review_counts_reviews(usa_review)}
        • Attained Percentage: {usa_attained_pct}%

        *Brands*
        - 📘 *BookMarketeers:* `{bookmarketeers}`
        - 📙 *Writers Clique:* `{writers_clique}`
        - 📕 *KDP:* `{kdp}`

        *Platforms*
        - 🅰 *Amazon:* `{usa_amazon}`
        - 📔 *Barnes & Noble:* `{usa_bn}`
        - ⚡ *Ingram Spark:* `{usa_ingram}`

        *UK Reviews:*
        • Total Reviews: {uk_total}
        • Status Breakdown: {format_review_counts_reviews(uk_review)}
        • Attained Percentage: {uk_attained_pct}%

        *Brand*
        - 📘 *Authors Solution:* `{authors_solution}`

        *Platforms*
        - 🅰 *Amazon:* `{uk_amazon}`
        - 📔 **Barnes & Noble:* `{uk_bn}`
        - ⚡ *Ingram Spark:* `{uk_ingram}`

        *Combined Stats:*
        • Total Reviews: {combined_total}
        • Attained Reviews: {combined_attained} ({combined_attained_pct}%)
        • Platform Totals:
          - 🅰 *Amazon:* `{combined_amazon}`
          - 📔 *Barnes & Noble:* `{combined_bn}`
          - ⚡ *Ingram Spark:* `{combined_ingram}`

        *Printing Stats:*
        •🧾 Total Copies: {Total_copies}
        •💰 Total Cost: ${Total_cost:.2f}
        •📈 Highest Cost: ${Highest_cost:.2f}
        •📉 Lowest Cost: ${Lowest_cost:.2f}
        •🔢 Highest Copies: {Highest_copies}
        •🧮 Lowest Copies: {Lowest_copies}
        •🧾 Average Cost: ${Average:.2f} per copy

        *Copyright Stats:*
        •🧾 Total Copyrights: {Total_copyrights}
        •💵 Total Cost: ${Total_cost_copyright}
        •✅ Total Successful: {result_count} / {Total_copyrights}
        • 🦅 *USA:* `{usa}`
        • 🍁 *Canada:* `{canada}`
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


def format_review_counts_reviews(review_counts):
    """Format review counts as a string"""
    if review_counts.empty:
        return "No data"
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
        time.sleep(2)
        send_df_as_text(name, sheet_usa, email, channel_usa)
    #
    # for name, email in names_uk.items():
    #     # time.sleep(5)
    #     send_df_as_text(name, sheet_uk, email, channel_uk)
    # summary(5, 2025)
    # generate_year_summary(2025)
