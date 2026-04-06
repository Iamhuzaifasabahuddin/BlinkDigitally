import pandas as pd
from data_loader import get_sheet_data, sheet_printing, sheet_copyright, sheet_a_plus
from data_cleaner import get_min_year

def get_printing_data_month(month: int, year: int) -> pd.DataFrame:
    """Get printing data for the current month"""
    data = get_sheet_data(sheet_printing)

    if data.empty:
        return pd.DataFrame()

    columns = list(data.columns)
    if "Accepted" in columns:
        end_col_index = columns.index("Accepted")
        data = data.iloc[:, :end_col_index + 1]
        data = data.astype(str)

    for col in ["Order Date", "Shipping Date", "Fulfilled"]:
        if col in data.columns:
            data[col] = pd.to_datetime(data[col], format="%d-%B-%Y", errors="coerce")

    data = data[(data["Order Date"].dt.month == month) & (data["Order Date"].dt.year == year)]

    data = data.sort_values(by="Order Date", ascending=True)
    if "Order Cost" in data.columns:
        data["Order Cost"] = data["Order Cost"].fillna(0)
        data["Order Cost"] = data["Order Cost"].astype(str)
        data["Order Cost"] = pd.to_numeric(
            data["Order Cost"].str.replace("$", "", regex=False).str.replace(",", "", regex=False),
            errors="coerce").fillna(0)

    if "No of Copies" in data.columns:
        data["No of Copies"] = pd.to_numeric(data["No of Copies"], errors='coerce').fillna(0)

    data = data.sort_values(by="Order Date", ascending=True)

    for col in ["Order Date", "Shipping Date", "Fulfilled"]:
        if col in data.columns:
            data[col] = pd.to_datetime(data[col], errors="coerce").dt.strftime("%d-%B-%Y")

    data.index = range(1, len(data) + 1)

    return data


def printing_data_year(year: int) -> tuple[pd.DataFrame, pd.DataFrame]:
    data = get_sheet_data(sheet_printing)

    if data.empty:
        return pd.DataFrame(), pd.DataFrame()

    columns = list(data.columns)
    if "Accepted" in columns:
        end_col_index = columns.index("Accepted")
        data = data.iloc[:, :end_col_index + 1]

    data = data.astype(str)

    for col in ["Order Date", "Shipping Date", "Fulfilled"]:
        if col in data.columns:
            data[col] = pd.to_datetime(data[col], format="%d-%B-%Y", errors="coerce")

    data = data[data["Order Date"].dt.year == year]

    data = data.sort_values(by="Order Date", ascending=True)
    if data.empty:
        return pd.DataFrame(), pd.DataFrame()

    if "Order Cost" in data.columns:
        data["Order Cost"] = pd.to_numeric(
            data["Order Cost"].str.replace("$", "", regex=False).str.replace(",", "", regex=False),
            errors="coerce"
        ).fillna(0)

    if "No of Copies" in data.columns:
        data["No of Copies"] = pd.to_numeric(data["No of Copies"], errors='coerce').fillna(0)

    data['Month'] = data['Order Date'].dt.to_period('M')

    month_totals = data.groupby('Month').agg(
        Total_Copies=('No of Copies', 'sum'),
        Total_Cost=('Order Cost', 'sum')
    ).reset_index()

    month_totals['Month'] = month_totals['Month'].dt.strftime('%B %Y')
    month_totals.columns = ["Month", "Total Copies", "Total Cost ($)"]
    month_totals = month_totals.sort_values(by="Total Cost ($)", ascending=False)
    month_totals.index = range(1, len(month_totals) + 1)
    month_totals["Total Cost ($)"] = month_totals["Total Cost ($)"].map("${:,.2f}".format)
    for col in ["Order Date", "Shipping Date", "Fulfilled"]:
        if col in data.columns:
            data[col] = data[col].dt.strftime("%d-%B-%Y")

    data.index = range(1, len(data) + 1)

    return data, month_totals

def printing_data_search(year: int) -> tuple[pd.DataFrame, pd.DataFrame]:
    data = get_sheet_data(sheet_printing)

    if data.empty:
        return pd.DataFrame(), pd.DataFrame()

    columns = list(data.columns)
    if "Accepted" in columns:
        end_col_index = columns.index("Accepted")
        data = data.iloc[:, :end_col_index + 1]

    data = data.astype(str)

    for col in ["Order Date", "Shipping Date", "Fulfilled"]:
        if col in data.columns:
            data[col] = pd.to_datetime(data[col], format="%d-%B-%Y", errors="coerce")

    data = data[
        (data["Order Date"].dt.year >= get_min_year()) &
        (data["Order Date"].dt.year <= year)

    ]

    data = data.sort_values(by="Order Date", ascending=True)
    if data.empty:
        return pd.DataFrame(), pd.DataFrame()

    if "Order Cost" in data.columns:
        data["Order Cost"] = pd.to_numeric(
            data["Order Cost"].str.replace("$", "", regex=False).str.replace(",", "", regex=False),
            errors="coerce"
        ).fillna(0)

    if "No of Copies" in data.columns:
        data["No of Copies"] = pd.to_numeric(data["No of Copies"], errors='coerce').fillna(0)

    data['Month'] = data['Order Date'].dt.to_period('M')

    month_totals = data.groupby('Month').agg(
        Total_Copies=('No of Copies', 'sum'),
        Total_Cost=('Order Cost', 'sum')
    ).reset_index()

    month_totals['Month'] = month_totals['Month'].dt.strftime('%B %Y')
    month_totals.columns = ["Month", "Total Copies", "Total Cost ($)"]
    month_totals = month_totals.sort_values(by="Total Cost ($)", ascending=False)
    month_totals.index = range(1, len(month_totals) + 1)
    month_totals["Total Cost ($)"] = month_totals["Total Cost ($)"].map("${:,.2f}".format)
    for col in ["Order Date", "Shipping Date", "Fulfilled"]:
        if col in data.columns:
            data[col] = data[col].dt.strftime("%d-%B-%Y")

    data.index = range(1, len(data) + 1)

    return data, month_totals

def printing_data_year_multiple(start_year: int, end_year: int) -> tuple[pd.DataFrame, pd.DataFrame]:
    data = get_sheet_data(sheet_printing)

    if data.empty:
        return pd.DataFrame(), pd.DataFrame()

    columns = list(data.columns)
    if "Accepted" in columns:
        end_col_index = columns.index("Accepted")
        data = data.iloc[:, :end_col_index + 1]

    data = data.astype(str)

    for col in ["Order Date", "Shipping Date", "Fulfilled"]:
        if col in data.columns:
            data[col] = pd.to_datetime(data[col], format="%d-%B-%Y", errors="coerce")

    data = data[
        (data["Order Date"].dt.year >= start_year) &
        (data["Order Date"].dt.year <= end_year)
    ]

    data = data.sort_values(by="Order Date", ascending=True)
    if data.empty:
        return pd.DataFrame(), pd.DataFrame()

    if "Order Cost" in data.columns:
        data["Order Cost"] = pd.to_numeric(
            data["Order Cost"].str.replace("$", "", regex=False).str.replace(",", "", regex=False),
            errors="coerce"
        ).fillna(0)

    if "No of Copies" in data.columns:
        data["No of Copies"] = pd.to_numeric(data["No of Copies"], errors='coerce').fillna(0)

    data['Month'] = data['Order Date'].dt.to_period('M')

    month_totals = data.groupby('Month').agg(
        Total_Copies=('No of Copies', 'sum'),
        Total_Cost=('Order Cost', 'sum')
    ).reset_index()

    month_totals['Month'] = month_totals['Month'].dt.strftime('%B %Y')
    month_totals.columns = ["Month", "Total Copies", "Total Cost ($)"]
    month_totals = month_totals.sort_values(by="Total Cost ($)", ascending=False)
    month_totals.index = range(1, len(month_totals) + 1)
    month_totals["Total Cost ($)"] = month_totals["Total Cost ($)"].map("${:,.2f}".format)
    for col in ["Order Date", "Shipping Date", "Fulfilled"]:
        if col in data.columns:
            data[col] = data[col].dt.strftime("%d-%B-%Y")

    data.index = range(1, len(data) + 1)

    return data, month_totals

def get_copyright_month(month: int, year: int) -> tuple[pd.DataFrame, int, int]:
    """Get copyright data for the current month"""
    data = get_sheet_data(sheet_copyright)

    if data.empty:
        return pd.DataFrame(), 0, 0

    columns = list(data.columns)
    if "Country" in columns:
        end_col_index = columns.index("Country")
        data = data.iloc[:, :end_col_index + 1]
    data = data.astype(str)

    if "Submission Date" in data.columns:
        data["Submission Date"] = pd.to_datetime(data["Submission Date"], format="%d-%B-%Y", errors='coerce')
        data = data[
            (data["Submission Date"].dt.month == month) & (data["Submission Date"].dt.year == year)]

    data = data.sort_values(by=["Submission Date"], ascending=True)
    result_count = len(data[data["Result"] == "Yes"]) if "Result" in data.columns else 0
    result_count_no = len(data[data["Result"] == "No"]) if "Result" in data.columns else 0
    if "Submission Date" in data.columns:
        data["Submission Date"] = data["Submission Date"].dt.strftime("%d-%B-%Y")

    data = data.fillna("N/A")

    data.index = range(1, len(data) + 1)

    return data, result_count, result_count_no


def copyright_year(year: int) -> tuple[pd.DataFrame, int, int]:
    data = get_sheet_data(sheet_copyright)

    if data.empty:
        return pd.DataFrame(), 0, 0

    columns = list(data.columns)
    if "Country" in columns:
        end_col_index = columns.index("Country")
        data = data.iloc[:, :end_col_index + 1]
    data = data.astype(str)

    if "Submission Date" in data.columns:
        data["Submission Date"] = pd.to_datetime(data["Submission Date"], format="%d-%B-%Y", errors='coerce')
        data = data[
            (data["Submission Date"].dt.year == year)]
    data = data.sort_values(by=["Submission Date"], ascending=True)

    result_count = len(data[data["Result"] == "Yes"]) if "Result" in data.columns else 0
    result_count_no = len(data[data["Result"] == "No"]) if "Result" in data.columns else 0
    if "Submission Date" in data.columns:
        data["Submission Date"] = data["Submission Date"].dt.strftime("%d-%B-%Y")

    data = data.fillna("N/A")

    data.index = range(1, len(data) + 1)

    return data, result_count, result_count_no

def copyright_search(year: int) -> tuple[pd.DataFrame, int, int]:
    data = get_sheet_data(sheet_copyright)

    if data.empty:
        return pd.DataFrame(), 0, 0

    columns = list(data.columns)
    if "Country" in columns:
        end_col_index = columns.index("Country")
        data = data.iloc[:, :end_col_index + 1]
    data = data.astype(str)

    if "Submission Date" in data.columns:
        data["Submission Date"] = pd.to_datetime(data["Submission Date"], format="%d-%B-%Y", errors='coerce')
        data = data[
            (data["Submission Date"].dt.year >= get_min_year()) &
            (data["Submission Date"].dt.year <= year)

        ]
    data = data.sort_values(by=["Submission Date"], ascending=True)

    result_count = len(data[data["Result"] == "Yes"]) if "Result" in data.columns else 0
    result_count_no = len(data[data["Result"] == "No"]) if "Result" in data.columns else 0
    if "Submission Date" in data.columns:
        data["Submission Date"] = data["Submission Date"].dt.strftime("%d-%B-%Y")

    data = data.fillna("N/A")

    data.index = range(1, len(data) + 1)

    return data, result_count, result_count_no

def copyright_year_multiple(start_year: int, end_year: int) -> tuple[pd.DataFrame, int, int]:
    data = get_sheet_data(sheet_copyright)

    if data.empty:
        return pd.DataFrame(), 0, 0

    columns = list(data.columns)
    if "Country" in columns:
        end_col_index = columns.index("Country")
        data = data.iloc[:, :end_col_index + 1]
    data = data.astype(str)

    if "Submission Date" in data.columns:
        data["Submission Date"] = pd.to_datetime(data["Submission Date"], format="%d-%B-%Y", errors='coerce')
        data = data[
            (data["Submission Date"].dt.year >= start_year) &
            (data["Submission Date"].dt.year <= end_year)

        ]
    data = data.sort_values(by=["Submission Date"], ascending=True)

    result_count = len(data[data["Result"] == "Yes"]) if "Result" in data.columns else 0
    result_count_no = len(data[data["Result"] == "No"]) if "Result" in data.columns else 0
    if "Submission Date" in data.columns:
        data["Submission Date"] = data["Submission Date"].dt.strftime("%d-%B-%Y")

    data = data.fillna("N/A")

    data.index = range(1, len(data) + 1)

    return data, result_count, result_count_no

def get_A_plus_month(month: int, year: int) -> tuple[pd.DataFrame, int]:
    data = get_sheet_data(sheet_a_plus)
    if data.empty:
        return pd.DataFrame(), 0

    columns = list(data.columns)
    if "Issues" in columns:
        end_col_index = columns.index("Issues")
        data = data.iloc[:, :end_col_index + 1]
    data = data.astype(str)

    if "A+ Content Date" in data.columns:
        data["A+ Content Date"] = pd.to_datetime(data["A+ Content Date"], format="%d-%B-%Y", errors='coerce')
        data = data[
            (data["A+ Content Date"].dt.month == month) & (data["A+ Content Date"].dt.year == year)]
    data = data.sort_values(by=["A+ Content Date"], ascending=True)

    result_count = len(data[data["Status"] == "Published"]) if "Status" in data.columns else 0

    if "A+ Content Date" in data.columns:
        data["A+ Content Date"] = data["A+ Content Date"].dt.strftime("%d-%B-%Y")

    data = data.fillna("N/A")

    data.index = range(1, len(data) + 1)

    return data, result_count


def get_A_plus_year(year: int) -> tuple[pd.DataFrame, int]:
    data = get_sheet_data(sheet_a_plus)
    if data.empty:
        return pd.DataFrame(), 0

    columns = list(data.columns)
    if "Issues" in columns:
        end_col_index = columns.index("Issues")
        data = data.iloc[:, :end_col_index + 1]
    data = data.astype(str)

    if "A+ Content Date" in data.columns:
        data["A+ Content Date"] = pd.to_datetime(data["A+ Content Date"], format="%d-%B-%Y", errors='coerce')
        data = data[
            (data["A+ Content Date"].dt.year == year)]
    data = data.sort_values(by=["A+ Content Date"], ascending=True)

    result_count = len(data[data["Status"] == "Published"]) if "Status" in data.columns else 0

    if "A+ Content Date" in data.columns:
        data["A+ Content Date"] = data["A+ Content Date"].dt.strftime("%d-%B-%Y")

    data = data.fillna("N/A")

    data.index = range(1, len(data) + 1)

    return data, result_count

def get_A_plus_year_multiple(start_year: int, end_year: int) -> tuple[pd.DataFrame, int]:
    data = get_sheet_data(sheet_a_plus)
    if data.empty:
        return pd.DataFrame(), 0

    columns = list(data.columns)
    if "Issues" in columns:
        end_col_index = columns.index("Issues")
        data = data.iloc[:, :end_col_index + 1]
    data = data.astype(str)

    if "A+ Content Date" in data.columns:
        data["A+ Content Date"] = pd.to_datetime(data["A+ Content Date"] ,errors='coerce')
        data = data[
            (data["A+ Content Date"].dt.year >= start_year) &
            (data["A+ Content Date"].dt.year <= end_year)
        ]
    data = data.sort_values(by=["A+ Content Date"], ascending=True)

    result_count = len(data[data["Status"] == "Published"]) if "Status" in data.columns else 0

    if "A+ Content Date" in data.columns:
        data["A+ Content Date"] = data["A+ Content Date"].dt.strftime("%d-%B-%Y")

    data = data.fillna("N/A")

    data.index = range(1, len(data) + 1)

    return data, result_count