import calendar
import os
from datetime import datetime

import pandas as pd
import streamlit as st
from dotenv import load_dotenv

load_dotenv("Info.env")

SPREADSHEET_ID = os.getenv("SPREADSHEET_ID")
sheet_usa = "USA"
sheet_uk = "UK"

url_usa = f"https://docs.google.com/spreadsheets/d/{SPREADSHEET_ID}/gviz/tq?tqx=out:csv&sheet={sheet_usa}"
url_uk = f"https://docs.google.com/spreadsheets/d/{SPREADSHEET_ID}/gviz/tq?tqx=out:csv&sheet={sheet_uk}"

current_month = datetime.today().month

dropdown_country = st.selectbox(
    "Which Country Data to View?",
    ("UK", "USA"),
    index=None,
    placeholder="Select the country you wish to view",
)

# Month selection dropdown
month_list = list(calendar.month_name)[1:]
dropdown_month = st.selectbox(
    "Which Month Data to View?",
    month_list,
    index=current_month - 1,
    placeholder="Select Month"
)

selected_month_number = list(calendar.month_name).index(dropdown_month)


def view_data():
    if dropdown_country == "UK":
        data_uk = pd.read_csv(url_uk)
        columns = list(data_uk.columns)
        end_col_index = columns.index("Issues")
        data_uk = data_uk.iloc[:, :end_col_index + 1]
        data_uk["Publishing Date"] = pd.to_datetime(data_uk["Publishing Date"], errors="coerce")
        data_uk["Last Edit (Revision)"] = pd.to_datetime(data_uk["Last Edit (Revision)"], errors="coerce")
        data_uk["Trustpilot Review Date"] = pd.to_datetime(data_uk["Trustpilot Review Date"], errors="coerce")
        data_uk = data_uk[data_uk["Publishing Date"].dt.month == selected_month_number]

        # if d:
        #     start_date, end_date = d
        #     data_uk = data_uk[
        #         (data_uk["Publishing Date"] >= pd.to_datetime(start_date)) &
        #         (data_uk["Publishing Date"] <= pd.to_datetime(end_date))
        #         ]
        # else:

        st.title("UK")
        if len(data_uk) >0:
            st.dataframe(data_uk)
        else:
            st.write("No data available")
    elif dropdown_country == "USA":
        data_usa = pd.read_csv(url_usa)
        columns = list(data_usa.columns)
        end_col_index = columns.index("Issues")
        data_usa = data_usa.iloc[:, :end_col_index + 1]
        data_usa["Publishing Date"] = pd.to_datetime(data_usa["Publishing Date"], errors="coerce")
        data_usa["Last Edit (Revision)"] = pd.to_datetime(data_usa["Last Edit (Revision)"], errors="coerce")
        data_usa["Trustpilot Review Date"] = pd.to_datetime(data_usa["Trustpilot Review Date"], errors="coerce")
        data_usa = data_usa[data_usa["Publishing Date"].dt.month == selected_month_number]

        st.title("USA")

        if len(data_usa) >0:
            st.dataframe(data_usa)
        else:
            st.write("No data available")


if __name__ == '__main__':
    view_data()
