import os
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from dotenv import load_dotenv

# Load environment variables
load_dotenv('Info.env')

SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]
SPREADSHEET_ID = os.getenv("SPREADSHEET_ID")


def main():
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

        # Get Sheet Name
        Sheet_Name = input("Enter Sheet Name: ")
        while not Sheet_Name.strip():
            Sheet_Name = input("Sheet name can't be empty. Please enter again: ")

        # Get Row Number
        Row_no = input("Enter Row No (must be a positive number): ")
        while not Row_no.isdigit() or int(Row_no) <= 0:
            Row_no = input("Invalid input. Please enter a positive row number: ")
        Row_no = int(Row_no)

        # Ask user for A to P columns (16 fields)
        headers = [
            "Name", "Brand", "Book Name & Link", "Format", "Copyright", "ISBN", "Project Manager",
            "Email", "Password", "Platform", "Status", "Publishing Date",
            "Last Edit (Revision)", "Trustpilot Review", "Trustpilot Review Date", "Issues"
        ]

        row_data = []
        print("\nPlease enter the following details:")
        for header in headers:
            value = input(f"{header}: ")
            row_data.append(value)

        # Write the data to the specified row
        update_range = f"{Sheet_Name}!A{Row_no}:P{Row_no}"
        response = sheet.values().update(
            spreadsheetId=SPREADSHEET_ID,
            range=update_range,
            valueInputOption="USER_ENTERED",
            body={"values": [row_data]}
        ).execute()

        print(f"\n✅ Data successfully written to row {Row_no} in sheet '{Sheet_Name}'.")
        print(f"⏎ Updated range: {response.get('updatedRange')}")

    except HttpError as err:
        print(f"An error occurred: {err}")

if __name__ == '__main__':
    main()
