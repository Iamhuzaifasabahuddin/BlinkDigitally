import calendar
import os
import shutil
from datetime import datetime

import pandas as pd
from dotenv import load_dotenv

load_dotenv('Info.env')

# Sheet URLs
SPREADSHEET_ID = os.getenv("SPREADSHEET_ID")
sheet_usa = "USA"
sheet_uk = "UK"
url_usa = f"https://docs.google.com/spreadsheets/d/{SPREADSHEET_ID}/gviz/tq?tqx=out:csv&sheet={sheet_usa}"
url_uk = f"https://docs.google.com/spreadsheets/d/{SPREADSHEET_ID}/gviz/tq?tqx=out:csv&sheet={sheet_uk}"
base_path_uk = r"C:\Users\Huzaifa Sabah Uddin\Monthly\UK"
base_path_usa = r"C:\Users\Huzaifa Sabah Uddin\Monthly\USA"

current_month = datetime.today().month
# current_month = 4
current_month_name = calendar.month_name[current_month]


def clean_data(url: str) -> pd.DataFrame:
    data = pd.read_csv(url)
    columns = list(data.columns)
    end_col_index = columns.index("Issues")
    data = data.iloc[:, :end_col_index + 1]
    return data


def load_data(url):
    data = clean_data(url)
    data["Publishing Date"] = pd.to_datetime(data["Publishing Date"], errors='coerce')
    data = data[data["Publishing Date"].dt.month == current_month]

    data = data.sort_values(by=["Publishing Date"], ascending=True)
    data.index = range(1, len(data) + 1)

    data["Publishing Date"] = data["Publishing Date"].dt.strftime("%m-%B-%Y")

    return data


def move_file_safely(src_path, dest_dir):
    base_name = os.path.basename(src_path)
    name, ext = os.path.splitext(base_name)
    dest_path = os.path.join(dest_dir, base_name)
    counter = 1

    while os.path.exists(dest_path):
        dest_path = os.path.join(dest_dir, f"{name} ({counter}){ext}")
        counter += 1

    shutil.move(src_path, dest_path)
    print(f"Moved to: {dest_path}")


def generate_monthly():
    data_uk = load_data(url_uk)
    data_usa = load_data(url_usa)

    uk_filename = f"{current_month_name} UK.xlsx"
    usa_filename = f"{current_month_name} USA.xlsx"

    data_uk.to_excel(uk_filename, index=False)
    data_usa.to_excel(usa_filename, index=False)

    os.makedirs(base_path_uk, exist_ok=True)
    os.makedirs(base_path_usa, exist_ok=True)

    move_file_safely(uk_filename, base_path_uk)
    move_file_safely(usa_filename, base_path_usa)


if __name__ == '__main__':
    generate_monthly()
