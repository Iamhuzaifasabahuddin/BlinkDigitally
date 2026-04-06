import logging

import pandas as pd

from API_loader import get_sheet_data


def get_names_in_both_months(sheet_name: str, month_1: str, year1: int, month_2: str, year2: int) -> tuple:
    """
    Identifies names that appear in both June and July from a Google Sheet.
    Returns:
        - A set of matching names
        - A dictionary with individual counts for June and July
    """
    df = get_sheet_data(sheet_name)

    if df.empty or "Name" not in df.columns or "Publishing Date" not in df.columns:
        logging.warning("Missing 'Name' or 'Date' columns or data is empty.")
        return set(), {}, 0

    df['Publishing Date'] = pd.to_datetime(df['Publishing Date'], format="%d-%B-%Y", errors='coerce')
    df = df.dropna(subset=['Publishing Date', 'Name'])

    df['Month'] = df['Publishing Date'].dt.month_name()
    df['Year'] = df['Publishing Date'].dt.year

    month_1_names = set(
        df[(df['Month'] == month_1) & (df['Year'] == year1)]['Name'].str.strip()
    )

    month_2_names = set(
        df[(df['Month'] == month_2) & (df['Year'] == year2)]['Name'].str.strip()
    )

    if month_1_names & month_2_names:
        names_in_both = month_1_names.intersection(month_2_names)

        counts = {}
        for name in names_in_both:
            month_1_count = df[
                (df['Month'] == month_1) &
                (df['Year'] == year1) &
                (df['Name'].str.strip() == name)
                ].shape[0]

            month_2_count = df[
                (df['Month'] == month_2) &
                (df['Year'] == year2) &
                (df['Name'].str.strip() == name)
                ].shape[0]

            counts[name] = {
                f"{month_1}-{year1}": month_1_count,
                f"{month_2}-{year2}": month_2_count,
            }

        return names_in_both, counts, len(names_in_both)
    else:
        return set(), {}, 0

def get_names_in_both_years(sheet_name: str, year1: int, year2: int) -> tuple:
    """
    Identifies names that appear in both years from a Google Sheet.
    """
    df = get_sheet_data(sheet_name)

    if df.empty or "Name" not in df.columns or "Publishing Date" not in df.columns:
        logging.warning("Missing 'Name' or 'Publishing Date' columns or data is empty.")
        return set(), {}, 0

    df['Publishing Date'] = pd.to_datetime(
        df['Publishing Date'], format="%d-%B-%Y", errors='coerce'
    )
    df = df.dropna(subset=['Publishing Date', 'Name'])

    df['Year'] = df['Publishing Date'].dt.year
    df['Name'] = df['Name'].str.strip()

    year_1_names = set(df[df['Year'] == year1]['Name'])
    year_2_names = set(df[df['Year'] == year2]['Name'])

    names_in_both = year_1_names & year_2_names

    counts = {}

    for name in names_in_both:
        year1_df = df[(df['Year'] == year1) & (df['Name'] == name)]
        year2_df = df[(df['Year'] == year2) & (df['Name'] == name)]

        counts[name] = {
            str(year1): {
                "count": year1_df.shape[0],
                "publishing_dates": year1_df['Publishing Date']
                .dt.strftime("%d-%B-%Y")
                .tolist()
            },
            str(year2): {
                "count": year2_df.shape[0],
                "publishing_dates": year2_df['Publishing Date']
                .dt.strftime("%d-%B-%Y")
                .tolist()
            }
        }

    return names_in_both, counts, len(names_in_both)

def get_clients_returning_in_month(
    sheet_name: str,
    start_year: int,
    target_month: str,
    target_year: int
) -> tuple:
    """
    Identifies clients who were published starting from `start_year`
    and also appear in the specified `target_month` and `target_year`.

    Returns clients that are "duplicates" — same client, different books.
    """
    df = get_sheet_data(sheet_name)

    if df.empty or "Name" not in df.columns or "Publishing Date" not in df.columns:
        logging.warning("Missing 'Name' or 'Publishing Date' columns or data is empty.")
        return set(), {}, 0

    df['Publishing Date'] = pd.to_datetime(
        df['Publishing Date'], format="%d-%B-%Y", errors='coerce'
    )
    df = df.dropna(subset=['Publishing Date', 'Name'])

    df['Year'] = df['Publishing Date'].dt.year
    df['Month'] = df['Publishing Date'].dt.month_name()
    df['Name'] = df['Name'].str.strip()

    baseline_df = df[
        (df['Year'] == start_year)
    ]

    target_df = df[
        (df['Year'] == target_year) &
        (df['Month'] == target_month)
    ]

    baseline_clients = set(baseline_df['Name'])
    target_clients = set(target_df['Name'])

    returning_clients = baseline_clients & target_clients

    counts = {}
    for name in returning_clients:
        client_baseline = baseline_df[baseline_df['Name'] == name]
        client_target = target_df[target_df['Name'] == name]

        counts[name] = {
            f"from_{start_year}_baseline": {
                "count": client_baseline.shape[0],
                "publishing_dates": client_baseline['Publishing Date']
                    .dt.strftime("%d-%B-%Y")
                    .tolist()
            },
            f"{target_year}_{target_month}": {
                "count": client_target.shape[0],
                "publishing_dates": client_target['Publishing Date']
                    .dt.strftime("%d-%B-%Y")
                    .tolist()
            }
        }

    return returning_clients, counts, len(returning_clients)

def get_names_in_year(sheet_name: str, year: int):
    """
    Finds names that appear in multiple months within the same year.

    Returns:
        - A DataFrame of names with counts per month
        - A dictionary summary of names with their total appearances and months active
        - Total count of such names
    """
    df = get_sheet_data(sheet_name)

    if df.empty or "Name" not in df.columns or "Publishing Date" not in df.columns:
        logging.warning("Missing 'Name' or 'Publishing Date' columns, or data is empty.")
        return pd.DataFrame(), {}, 0

    df['Publishing Date'] = pd.to_datetime(df['Publishing Date'], format="%d-%B-%Y", errors='coerce')
    df = df.dropna(subset=['Publishing Date', 'Name'])
    df['Month'] = df['Publishing Date'].dt.month_name()
    df['Year'] = df['Publishing Date'].dt.year

    df = df[df['Year'] == year]

    if df.empty:
        logging.warning(f"No records found for year {year}.")
        return pd.DataFrame(), {}, 0

    monthly_counts = (
        df.groupby(['Name', 'Month'])
        .size()
        .unstack(fill_value=0)
        .reindex(columns=[
            'January', 'February', 'March', 'April', 'May', 'June',
            'July', 'August', 'September', 'October', 'November', 'December'
        ], fill_value=0)
    )

    monthly_counts['Active Months'] = (monthly_counts > 0).sum(axis=1)
    multi_month_names = monthly_counts[monthly_counts['Active Months'] > 1].copy()

    month_cols = multi_month_names.columns[:-2]
    summary = {}
    for name in multi_month_names.index:
        active_months = [month for month in month_cols if multi_month_names.at[name, month] > 0]

        indexed_months = {i + 1: month for i, month in enumerate(active_months)}

        summary[name] = {
            "Months Active": indexed_months,
            "Month Count": int(multi_month_names.at[name, "Active Months"]),
        }

    return multi_month_names, summary, len(multi_month_names)