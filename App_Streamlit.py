import calendar
from datetime import datetime

import pandas as pd
import streamlit as st
from streamlit_gsheets import GSheetsConnection

conn = st.connection("gsheets", type=GSheetsConnection)

sheet_usa = "USA"
sheet_uk = "UK"
sheet_audio = "AudioBook"
sheet_printing = "Printing"

month_list = list(calendar.month_name)[1:]
current_month = datetime.today().month

st.markdown("""
 <style>
    #MainMenu {visibility: hidden;}
    header {visibility: hidden;}
</style>
""", unsafe_allow_html=True)


def clean_data(data: pd.DataFrame) -> pd.DataFrame:
    """Clean and prepare the dataframe"""
    if data.empty:
        return pd.DataFrame()

    columns = list(data.columns)
    if "Issues" in columns:
        end_col_index = columns.index("Issues")
        data = data.iloc[:, :end_col_index + 1]

    date_columns = ["Publishing Date", "Last Edit (Revision)", "Trustpilot Review Date"]
    for col in date_columns:
        if col in data.columns:
            data[col] = pd.to_datetime(data[col], errors="coerce")

    return data


def load_data(sheet_name, month_number=None):
    """Load data from Google Sheets with optional month filtering"""
    try:
        data = conn.read(worksheet=sheet_name)
        data = clean_data(data)

        if month_number and "Publishing Date" in data.columns:
            data = data[data["Publishing Date"].dt.month == month_number]

        data.index = range(1, len(data) + 1)

        return data
    except Exception as e:
        st.error(f"Error loading data: {e}")
        return pd.DataFrame()


def add_data(sheet_name, new_row):
    """Add a new row to the specified sheet"""
    try:

        current_data = conn.read(worksheet=sheet_name)

        updated_data = pd.concat([current_data, pd.DataFrame([new_row], columns=current_data.columns)],
                                 ignore_index=True)
        conn.update(worksheet=sheet_name, data=updated_data, headers=False)
        st.success("Data successfully added to the sheet.")
        st.dataframe(load_data(sheet_name))
    except Exception as e:
        st.error(f"An error occurred: {e}")


def review_data(sheet_name, month, status):
    """Filter data by month and review status"""
    data = load_data(sheet_name)
    if not data.empty and month and status:
        if "Publishing Date" in data.columns and "Trustpilot Review" in data.columns:
            data = data[data["Publishing Date"].dt.month == month]
            data = data[data["Trustpilot Review"] == status]
    data.index = range(1, len(data) + 1)
    return data


def get_printing_data(month):
    """Get printing data filtered by month"""
    try:
        data = conn.read(worksheet=sheet_printing)

        for col in ["Order Date", "Shipping Date", "Fulfilled"]:
            if col in data.columns:
                data[col] = pd.to_datetime(data[col], errors="coerce")

        if month and "Order Date" in data.columns:
            data = data[data["Order Date"].dt.month == month]

        if "Order Cost" in data.columns:
            data["Order Cost"] = data["Order Cost"].astype(str)
            data["Order Cost"] = pd.to_numeric(data["Order Cost"].str.replace("$", "", regex=False), errors="coerce")
        data.index = range(1, len(data) + 1)
        return data
    except Exception as e:
        st.error(f"Error loading printing data: {e}")
        return pd.DataFrame()

def format_review_counts(review_counts):
    """Format review counts as a string"""
    return ", ".join([f"{status}: {count}" for status, count in review_counts.items()])


with st.container():
    st.title("📊 Data Management Portal")
    action = st.selectbox("What would you like to do?",
                          ["View Data", "Add Data", "Print Data", "Reviews", "Printing"],
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
        choice = st.selectbox("Select Data To View", ["UK", "USA", "AudioBook"], index=None,
                              placeholder="Select Data to View")

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
    if action == "View Data" and choice and selected_month:
        st.subheader(f"📂 Viewing Data for {choice} - {selected_month}")

        sheet_name = {
            "UK": sheet_uk,
            "USA": sheet_usa,
            "AudioBook": sheet_audio
        }.get(choice)

        if sheet_name:
            data = load_data(sheet_name, selected_month_number)

            for col in ["Publishing Date", "Last Edit (Revision)", "Trustpilot Review Date"]:
                if col in data.columns:
                    data[col] = pd.to_datetime(data[col], errors="coerce").dt.strftime("%d-%B-%Y")

            if data.empty:
                st.info(f"No data available for {selected_month} for {choice}")
            else:
                st.markdown("### 📄 Detailed Entry Data")
                st.dataframe(data)

                st.markdown("### ⭐ Trustpilot Review Summary")
                reviews = data["Trustpilot Review"].value_counts()
                total_reviews = reviews.sum()
                attained = reviews.get("Attained", 0)
                percentage = round((attained / total_reviews * 100), 1) if total_reviews > 0 else 0

                st.markdown(f"""
                            - 🧾 **Total Entries:** `{len(data)}`
                            - 🗳️ **Total Trustpilot Reviews:** `{total_reviews}`
                            - 🟢 **'Attained' Reviews:** `{attained}`
                            - 📊 **Attainment Rate:** `{percentage}%`
                            """)

                st.markdown("#### 🔍 Review Type Breakdown")
                for review_type, count in reviews.items():
                    st.markdown(f"- 📝 **{review_type}**: `{count}`")
    elif action == "Add Data" and country:
        st.subheader(f"➕ Add Data for {country}")
        sheet_name = sheet_uk if country == "UK" else sheet_usa

        name = st.text_input("Name")

        if country == "UK":
            brand = st.selectbox("Brand", ["Authors Solution", "KDP"], index=None, placeholder="Select Brand")
        else:
            brand = st.selectbox("Brand", ["BookMarketeers", "Writers Clique", "KDP"], index=None,
                                 placeholder="Select Brand")

        book_link = st.text_input("Book Name & Link")
        format_ = st.selectbox("Format",
                               ["eBook", "Paperback", "Hardcover", "eBook & Paperback",
                                "eBook, Paperback & Hardcover"],
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
        status = st.selectbox("Status", ["Pending", "In-Progress", "Published"], index=None,
                              placeholder="Select Status")
        publish_date = st.date_input("Publishing Date")
        last_edit = st.date_input("Last Edit (Revision)", value=None)
        review = st.selectbox("Review Status", ["Pending", "Sent", "Attained"], index=None,
                              placeholder="Select Review Status")
        review_date = st.date_input("Trustpilot Review Date", value=None)
        issues = st.text_input("Issues") or "N/A"

        if st.button("Submit"):
            new_row = {
                "Name": name,
                "Brand": brand,
                "Book Name & Link": book_link,
                "Format": format_,
                "Copyright": copyright_,
                "ISBN": isbn,
                "Project Manager": manager,
                "Email": email,
                "Password": password,
                "Platform": platform,
                "Status": status,
                "Publishing Date": publish_date.strftime('%d-%B-%Y'),
                "Last Edit (Revision)": last_edit.strftime('%d-%B-%Y') if last_edit else "N/A",
                "Trustpilot Review": review,
                "Trustpilot Review Date": review_date.strftime('%d-%B-%Y') if review_date else "N/A",
                "Issues": issues
            }
            add_data(sheet_name, new_row)

    elif action == "Print Data" and country and selected_month:
        st.subheader(f"🖰 Print Data for {country} - {selected_month}")
        sheet_name = sheet_uk if country == "UK" else sheet_usa

        data = load_data(sheet_name, selected_month_number)

        for col in ["Publishing Date", "Last Edit (Revision)", "Trustpilot Review Date"]:
            if col in data.columns:
                data[col] = pd.to_datetime(data[col], errors="coerce").dt.strftime("%d-%B-%Y")

        if not data.empty:
            st.dataframe(data)

            # Uncomment to enable download
            # csv = data.to_csv(index=False).encode("utf-8")
            # st.download_button("📥 Download CSV", data=csv, file_name=f"{country}_{selected_month}.csv", mime="text/csv")
        else:
            st.warning(f"No data available for {selected_month} for {country}")

    elif action == "Reviews" and country and selected_month and status:
        sheet_name = sheet_uk if country == "UK" else sheet_usa
        data = review_data(sheet_name, selected_month_number, status)

        for col in ["Publishing Date", "Last Edit (Revision)", "Trustpilot Review Date"]:
            if col in data.columns:
                data[col] = pd.to_datetime(data[col], errors="coerce").dt.strftime("%d-%B-%Y")

        st.subheader(f"🔍 Review Data - {status} in {selected_month} ({country})")

        if not data.empty:
            st.dataframe(data)
        else:
            st.info("No matching reviews found.")
    elif action == "Printing" and selected_month:

        st.subheader(f"🖨️ Printing Summary for {selected_month}")

        data = get_printing_data(selected_month_number)

        if not data.empty:

            # Calculations

            Total_copies = data["No of Copies"].sum()

            Total_cost = data["Order Cost"].sum()

            Highest_cost = data["Order Cost"].max()

            Highest_copies = data["No of Copies"].max()

            Lowest_cost = data["Order Cost"].min()

            Lowest_copies = data["No of Copies"].min()

            data['Cost_Per_Copy'] = data['Order Cost'] / data['No of Copies']

            Average = round(Total_cost / Total_copies, 2) if Total_copies else 0


            for col in ["Order Date", "Shipping Date", "Fulfilled"]:

                if col in data.columns:
                    data[col] = pd.to_datetime(data[col], errors="coerce").dt.strftime("%d-%B-%Y")

            st.markdown("### 📄 Detailed Printing Data")

            st.dataframe(data)

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

        else:

            st.warning(f"⚠️ No Data Available for Printing in {selected_month}")
