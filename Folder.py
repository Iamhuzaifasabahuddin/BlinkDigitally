import os

import pandas as pd
from dotenv import load_dotenv
from rapidfuzz import fuzz

load_dotenv("Info.env")

SPREADSHEET_ID = os.getenv("SPREADSHEET_ID")
sheet_usa = "USA"
sheet_uk = "UK"

url_usa = f"https://docs.google.com/spreadsheets/d/{SPREADSHEET_ID}/gviz/tq?tqx=out:csv&sheet={sheet_usa}"
url_uk = f"https://docs.google.com/spreadsheets/d/{SPREADSHEET_ID}/gviz/tq?tqx=out:csv&sheet={sheet_uk}"

data_uk = pd.read_csv(url_uk)
data_usa = pd.read_csv(url_usa)

names_usa = list(set(data_usa["Name"].dropna().tolist()))
names_uk = list(set(data_uk["Name"].dropna().tolist()))
base_path_usa = r"C:\Users\Huzaifa Sabah Uddin\Publishing\USA"
base_path_uk = r"C:\Users\Huzaifa Sabah Uddin\Publishing\UK"
main_folders_usa = [name for name in os.listdir(base_path_usa)
                    if os.path.isdir(os.path.join(base_path_usa, name))]
main_folders_uk = [name for name in os.listdir(base_path_uk)
                   if os.path.isdir(os.path.join(base_path_uk, name))]
true_list_usa = []
true_list_uk = []
created_list_usa = []
created_list_uk = []


def make_dir_usa() -> None:

    try:

        for name in names_usa:
            match_found = False
            for folder in main_folders_usa:
                similarity = fuzz.partial_ratio(name.lower(), folder.lower())
                if similarity >= 80:
                    true_list_usa.append(folder)
                    match_found = True
                    break

            if not match_found:
                new_folder_path = os.path.join(base_path_usa, name)
                os.makedirs(new_folder_path, exist_ok=True)
                created_list_usa.append(name)

        print("✅ Matched folders USA:", true_list_usa)
        print("📁 Created folders USA:", created_list_usa)
    except Exception as error:
        print(error)

def make_dir_uk() -> None:
    try:

        for name in names_uk:
            match_found = False
            for folder in main_folders_uk:
                similarity = fuzz.partial_ratio(name.lower(), folder.lower())
                if similarity >= 80:
                    true_list_uk.append(folder)
                    match_found = True
                    break

            if not match_found:
                new_folder_path = os.path.join(base_path_uk, name)
                os.makedirs(new_folder_path, exist_ok=True)
                created_list_uk.append(name)

        print("✅ Matched folders UK:", true_list_uk)
        print("📁 Created folders UK:", created_list_uk)
    except Exception as error:
        print(error)



if __name__ == '__main__':
    make_dir_uk()
    make_dir_usa()