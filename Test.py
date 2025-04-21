import pandas as pd
from dotenv import load_dotenv
import os
load_dotenv('Info.env')

# Load environment variables
SPREADSHEET_ID = os.getenv("SPREADSHEET_ID")
sheet_name = "USA"

url = f"https://docs.google.com/spreadsheets/d/{SPREADSHEET_ID}/gviz/tq?tqx=out:csv&sheet={sheet_name}"
data = pd.read_csv(url)
# print(data.head(5))
# print(data.columns)
columns = list(data.columns)

# Find the index of the "Issues" column
end_col_index = columns.index("Issues")

# Slice the DataFrame to include columns up to and including "Issues"
data_subset = data.iloc[:, :end_col_index + 1]

# # Print the first 5 rows of the subset
# # print(data_subset.head(5))
# data["Publishing Date"] = pd.to_datetime(data["Publishing Date"], errors="coerce")
#
# april_data = data[data["Publishing Date"].dt.month == 5]
#
# # Show the result
# print(april_data.head())
