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
url_Audio = f"https://docs.google.com/spreadsheets/d/{SPREADSHEET_ID}/gviz/tq?tqx=out:csv&sheet=AudioBook"
url_printing = f"https://docs.google.com/spreadsheets/d/{SPREADSHEET_ID}/gviz/tq?tqx=out:csv&sheet=Printing"

month_list = list(calendar.month_name)[1:]
current_month = datetime.today().month
st.title("📊 Data Management Portal")
action = st.selectbox("What would you like to do?", ["View Data", "Add Data", "Print Data", "Reviews", "Printing"],
                      index=None,
                      placeholder="Select Action")

country = None
selected_month = None
selected_month_number = None
status = None
choice = None

if action in ["Add Data", "Print Data", "Reviews"]:
    country = st.selectbox("Select Country", ["UK", "USA"], index=None, placeholder="Select Country")

if action == "View Data":
    choice = st.selectbox("Select Data To View", ["UK", "USA", "AudioBook"], index=None, placeholder="Select Data to View")

if action in ["View Data", "Print Data", "Reviews", "Printing"]:
    selected_month = st.selectbox(
        "Select Month",
        month_list,
        index=current_month - 1,
        placeholder="Select Month"
    )
    selected_month_number = month_list.index(selected_month) + 1 if selected_month else None

if action == "Reviews":
    status = st.selectbox("Status", ["Pending", "Sent", "Attained"], index=None, placeholder="Select Status")


def clean_data(url: str) -> pd.DataFrame:
    data = pd.read_csv(url)
    columns = list(data.columns)
    end_col_index = columns.index("Issues")
    data = data.iloc[:, :end_col_index + 1]

    for col in ["Publishing Date", "Last Edit (Revision)", "Trustpilot Review Date"]:
        data[col] = pd.to_datetime(data[col], errors="coerce")

    return data


def load_data(url, month_number=None):
    data = clean_data(url)
    if month_number:
        data = data[data["Publishing Date"].dt.month == month_number]
    data.index = range(1, len(data) + 1)
    return data


def count_rows(url) -> int:
    data = clean_data(url)
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


def Review_data(url: str, month: int, status: str):
    data = clean_data(url)
    if month and status:
        data = data[data["Publishing Date"].dt.month == month]
        data = data[data["Trustpilot Review"] == status]
    data.index = range(1, len(data) + 1)
    return data

def Printing(url: str, month: int):
    data = pd.read_csv(url)
    columns = list(data.columns)
    end_col_index = columns.index("Fulfilled")
    data = data.iloc[:, :end_col_index + 1]

    for col in ["Order Date", "Shipping Date", "Fulfilled"]:
        data[col] = pd.to_datetime(data[col], errors="coerce")
    if month:
        data = data[data["Order Date"].dt.month == month]
    data['Order Cost'] = pd.to_numeric(data['Order Cost'].str.replace('$', '', regex=False))
    data.index = range(1, len(data) + 1)

    return data



if action == "View Data" and choice and selected_month:
    st.subheader(f"📂 Viewing Data for {choice} - {selected_month}")
    url = None
    if choice == "UK":
        url = url_uk
    elif choice == "AudioBook":
        url = url_Audio
    else:
        url = url_usa

    data = load_data(url, selected_month_number)

    for col in ["Publishing Date", "Last Edit (Revision)", "Trustpilot Review Date"]:
        data[col] = pd.to_datetime(data[col], errors="coerce").dt.strftime(
                "%d-%B-%Y")
    if data.empty:
        st.info(f"No data available for {selected_month} for {choice}")
    else:
        st.dataframe(data)

elif action == "Add Data" and country:
    st.subheader(f"➕ Add Data for {country}")
    url = url_uk if country == "UK" else url_usa
    name = st.text_input("Name")
    if country == "UK":
        brand = st.selectbox("Brand", ["Authors Solution", "KDP"], index=None, placeholder="Select Brand")
    else:
        brand = st.selectbox("Brand", ["BookMarketeers", "Writers Clique", "KDP"], index=None,
                             placeholder="Select Brand")
    book_link = st.text_input("Book Name & Link")
    format_ = st.selectbox("Format",
                           ["eBook", "Paperback", "Hardcover", "eBook & Paperback", "eBook, Paperback & Hardcover"],
                           index=None, placeholder="Select Format")
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
    platform = st.selectbox("Platform", ["Amazon", "Ingram Spark"], index=None, placeholder="Select Platform")
    status = st.selectbox("Status", ["Pending", "In-Progress", "Published"], index=None, placeholder="Select Status")
    publish_date = st.date_input("Publishing Date", )
    last_edit = st.date_input("Last Edit (Revision)", value=None)
    review = st.selectbox("Review Status", ["Pending", "Sent", "Attained"], index=None,
                          placeholder="Select Review Status")
    review_date = st.date_input("Trustpilot Review Date", value=None)
    issues = st.text_input("Issues") or "N/A"

    if st.button("Submit"):
        row_number = count_rows(url_uk if country == "UK" else url_usa) + 2
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
    for col in ["Publishing Date", "Last Edit (Revision)", "Trustpilot Review Date"]:
        data[col] = pd.to_datetime(data[col], errors="coerce").dt.strftime(
            "%d-%B-%Y")
    if not data.empty:
        # csv = data.to_csv(index=False).encode("utf-8")
        # st.download_button("📥 Download CSV", data=csv, file_name=f"{country}_{selected_month}.csv", mime="text/csv")
        st.dataframe(data)
    else:
        st.warning(f"No data available for {selected_month} for {country}")

elif action == "Reviews" and country and selected_month and status:
    url = url_uk if country == "UK" else url_usa
    data = Review_data(url, selected_month_number, status)
    for col in ["Publishing Date", "Last Edit (Revision)", "Trustpilot Review Date"]:
        data[col] = pd.to_datetime(data[col], errors="coerce").dt.strftime(
                "%d-%B-%Y")
    st.subheader(f"🔍 Review Data - {status} in {selected_month} ({country})")
    st.dataframe(data if not data.empty else "No matching reviews found.")

elif action == "Printing" and selected_month:
    st.subheader(f"🔍 Printing Data for {selected_month})")
    data = Printing(url_printing, selected_month_number)
    for col in ["Order Date", "Shipping Date", "Fulfilled"]:
        data[col] = pd.to_datetime(data[col], errors="coerce").dt.strftime(
            "%d-%B-%Y")
    if not data.empty:
        st.dataframe(data)
    else:
        st.warning(f"No Data Available for Printing for {selected_month}")