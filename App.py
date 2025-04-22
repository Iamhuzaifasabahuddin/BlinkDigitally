import calendar
import os
from datetime import datetime

import pandas as pd
import streamlit as st
from dotenv import load_dotenv
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

# Load env variables
load_dotenv("Info.env")

# Google Sheet Info
SPREADSHEET_ID = os.getenv("SPREADSHEET_ID")
sheet_usa = "USA"
sheet_uk = "UK"
SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]

# Sheet URLs
url_usa = f"https://docs.google.com/spreadsheets/d/{SPREADSHEET_ID}/gviz/tq?tqx=out:csv&sheet={sheet_usa}"
url_uk = f"https://docs.google.com/spreadsheets/d/{SPREADSHEET_ID}/gviz/tq?tqx=out:csv&sheet={sheet_uk}"
url_printing = f"https://docs.google.com/spreadsheets/d/{SPREADSHEET_ID}/gviz/tq?tqx=out:csv&sheet=Printing"

month_list = list(calendar.month_name)[1:]
current_month = datetime.today().month

st.title("📊 Data Management Portal")

action = st.selectbox("What would you like to do?", ["View Data", "Add Data", "Print Data"], index=None, placeholder="Select Action")

country = None
selected_month = None
selected_month_number = None

if action in ["View Data", "Add Data", "Print Data"]:
    country = st.selectbox("Select Country", ["UK", "USA"], index=None, placeholder="Select Country")

if action in ["View Data", "Print Data"]:
    selected_month = st.selectbox(
        "Select Month",
        month_list,
        index=current_month - 1,
        placeholder="Select Month"
    )
    selected_month_number = month_list.index(selected_month) + 1 if selected_month else None


def load_data(url, month_number=None):
    data = pd.read_csv(url)
    columns = list(data.columns)
    end_col_index = columns.index("Issues")
    data = data.iloc[:, :end_col_index + 1]

    for col in ["Publishing Date", "Last Edit (Revision)", "Trustpilot Review Date"]:
        data[col] = pd.to_datetime(data[col], errors="coerce")

    if month_number:
        data = data[data["Publishing Date"].dt.month == month_number]

    return data


def count_rows(url) -> int:
    data = pd.read_csv(url)
    columns = list(data.columns)
    end_col_index = columns.index("Issues")
    data = data.iloc[:, :end_col_index + 1]

    for col in ["Publishing Date", "Last Edit (Revision)", "Trustpilot Review Date"]:
        data[col] = pd.to_datetime(data[col], errors="coerce")

    return len(data)


def Add_data(row: int, country: str, data: list, url):
    creds = None
    token_path = r'C:\Users\Huzaifa Sabah Uddin\PycharmProjects\BlinkDigitally\token.json'
    credentials_path = r"C:\Users\Huzaifa Sabah Uddin\PycharmProjects\BlinkDigitally\Hexz.json"

    if os.path.exists(token_path):
        creds = Credentials.from_authorized_user_file(token_path, SCOPES)

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(credentials_path, SCOPES)
            creds = flow.run_local_server(port=0)
        with open(token_path, 'w') as token:
            token.write(creds.to_json())

    try:
        service = build('sheets', 'v4', credentials=creds)
        sheet = service.spreadsheets()

        update_range = f"{country}!A{row}:P{row}"
        response = sheet.values().update(
            spreadsheetId=SPREADSHEET_ID,
            range=update_range,
            valueInputOption="USER_ENTERED",
            body={"values": [data]}
        ).execute()
        st.success("Data successfully added to the sheet.")
        st.dataframe(load_data(url))


    except HttpError as err:
        st.error(f"An error occurred: {err}")


if action == "View Data" and country and selected_month:
    st.subheader(f"📂 Viewing Data for {country} - {selected_month}")
    url = url_uk if country == "UK" else url_usa
    data = load_data(url, selected_month_number)
    st.dataframe(data if not data.empty else "No data available.")

elif action == "Add Data" and country:
    st.subheader(f"➕ Add Data for {country}")
    url = url_uk if country == "UK" else url_usa
    name = st.text_input("Name")
    if country == "UK":
        brand = st.selectbox("Select Brand", ["Authors Solution", "KDP"], index=None, placeholder="Select Country")
    else:
        brand = st.selectbox("Select Brand", ["BookMarketeers", "Writers Clique", "KDP"], index=None, placeholder="Select Country")
    book_link = st.text_input("Book Name & Link")
    format_ = st.selectbox("Select Format", ["eBook", "Paperback", "Hardcover", "eBook & Paperback", "eBook, Paperback & Hardcover"], index=None, placeholder="Select Format")
    copyright_ = st.text_input("Copyright") or "N/A"
    isbn = st.text_input("ISBN") or "Free"

    if country == "UK":
        manager = st.selectbox(
            "Project Manager",
            ["Syed Ahsan Shahzad", "Youha", "Hadia Ghazanfar"],
            index=None,
            placeholder="Select Project Manager"
        )

    else:
        manager = st.selectbox(
            "Project Manager",
            ["shaikh arsalan", "ahmed asif", "maheen sami", "aiza ali", "shozab hasan", "asad waqas"],
            index=None,
            placeholder="Select Project Manager"
        )

    email = st.text_input("Email")
    password = st.text_input("Password")
    platform = st.selectbox("Select Platform", ["Amazon", "Ingram Spark"], index=None, placeholder="Select Platform")
    status = st.selectbox("Select Status", ["Pending", "In-Progress", "Published"], index=None, placeholder="Select Status")
    publish_date = st.date_input("Publishing Date", )
    last_edit = st.date_input("Last Edit (Revision)", value=None)
    review = st.selectbox("Select Review Status", ["Pending", "Sent", "Attained"], index=None, placeholder="Select Review Status")
    review_date = st.date_input("Trustpilot Review Date", value=None)
    issues = st.text_input("Issues") or "N/A"

    if st.button("Submit"):
        row_number = count_rows(url_uk if country == "UK" else url_usa) + 1
        print(row_number)
        Add_data(row_number, country, [
            name, brand, book_link, format_, copyright_, isbn,
            manager, email, password, platform, status,
            publish_date.strftime('%d-%B-%Y'),
            last_edit.strftime('%d-%B-%Y') if last_edit else "N/A",
            review, review_date.strftime('%d-%B-%Y') if review_date else "N/A",
            issues
        ], url)

elif action == "Print Data" and country and selected_month:
    st.subheader(f"🖰 Print Data for {country} - {selected_month}")
    url = url_uk if country == "UK" else url_usa
    data = load_data(url, selected_month_number)

    if not data.empty:
        csv = data.to_csv(index=False).encode("utf-8")
        st.download_button("📥 Download CSV", data=csv, file_name=f"{country}_{selected_month}.csv", mime="text/csv")
        st.dataframe(data)
    else:
        st.warning("No data available to print.")
