import pandas as pd
import streamlit as st
from data_cleaner import clean_data_reviews, safe_concat
from data_loader import load_reviews, sheet_uk, sheet_usa, load_reviews_year_multiple, load_reviews_year

from diff_sheets_loader import get_A_plus_year, get_copyright_month, get_printing_data_month, get_A_plus_month, \
     printing_data_year, copyright_year, get_A_plus_year_multiple, printing_data_year_multiple, \
    copyright_year_multiple


def summary(month: int, year: int):
    uk_clean = clean_data_reviews(sheet_uk)
    usa_clean = clean_data_reviews(sheet_usa)

    usa_clean = usa_clean[
        (usa_clean["Publishing Date"].dt.month == month) &
        (usa_clean["Publishing Date"].dt.year == year)
        ]
    uk_clean = uk_clean[
        (uk_clean["Publishing Date"].dt.month == month) &
        (uk_clean["Publishing Date"].dt.year == year)
        ]

    usa_clean_platforms = usa_clean[
        (usa_clean["Publishing Date"].dt.month == month) &
        (usa_clean["Publishing Date"].dt.year == year)
        ]
    uk_clean_platforms = uk_clean[
        (uk_clean["Publishing Date"].dt.month == month) &
        (uk_clean["Publishing Date"].dt.year == year)
        ]

    if usa_clean.empty:
        print("No values found in USA sheet.")
        return
    if uk_clean.empty:
        print("No values found in UK sheet.")
        return
    if usa_clean.empty and uk_clean.empty:
        return
    usa_clean = usa_clean.drop_duplicates(subset=["Name"], keep="last")
    uk_clean = uk_clean.drop_duplicates(subset=["Name"], keep="last")
    Issues_usa = usa_clean["Issues"].value_counts()
    Issues_uk = uk_clean["Issues"].value_counts()
    total_usa = usa_clean["Name"].nunique()
    total_uk = uk_clean["Name"].nunique()
    total_unique_clients = total_usa + total_uk

    combined = pd.concat([usa_clean[["Name", "Brand", "Project Manager", "Email"]],
                          uk_clean[["Name", "Brand", "Project Manager", "Email"]]])
    combined.index = range(1, len(combined) + 1)

    brands = usa_clean["Brand"].value_counts()
    writers_clique = brands.get("Writers Clique", 0)
    bookmarketeers = brands.get("BookMarketeers", 0)
    aurora_writers = brands.get("Aurora Writers", 0)
    kdp = brands.get("KDP", 0)

    uk_brand = uk_clean["Brand"].value_counts()
    authors_solution = uk_brand.get("Authors Solution", 0)
    book_publication = uk_brand.get("Book Publication", 0)

    usa_platforms = usa_clean_platforms["Platform"].value_counts()
    usa_amazon = usa_platforms.get("Amazon", 0)
    usa_bn = usa_platforms.get("Barnes & Noble", 0)
    usa_ingram = usa_platforms.get("Ingram Spark", 0)
    usa_d2d = usa_platforms.get("Draft2Digital", 0)
    usa_lulu = usa_platforms.get("LULU", 0)
    usa_kobo = usa_platforms.get("Kobo", 0)
    usa_fav = usa_platforms.get("FAV", 0)
    usa_acx = usa_platforms.get("ACX", 0)

    uk_platforms = uk_clean_platforms["Platform"].value_counts()
    uk_amazon = uk_platforms.get("Amazon", 0)
    uk_bn = uk_platforms.get("Barnes & Noble", 0)
    uk_ingram = uk_platforms.get("Ingram Spark", 0)
    uk_d2d = uk_platforms.get("Draft2Digital", 0)
    uk_lulu = uk_platforms.get("LULU", 0)
    uk_fav = uk_platforms.get("FAV", 0)
    uk_kobo = uk_platforms.get("Kobo", 0)
    uk_acx = uk_platforms.get("ACX", 0)

    allowed_brands = ["BookMarketeers", "Writers Clique", "Aurora Writers", "Authors Solution", "Book Publication"]

    if "Trustpilot Review" in usa_clean.columns and "Brand" in usa_clean.columns:
        usa_filtered = usa_clean[usa_clean["Brand"].isin(allowed_brands)]
        usa_review_sent = usa_filtered["Trustpilot Review"].value_counts().get("Sent", 0)
        usa_review_pending = usa_filtered["Trustpilot Review"].value_counts().get("Pending", 0)
        usa_review_na = usa_filtered["Trustpilot Review"].value_counts().get("Negative", 0)
    else:
        usa_review_sent = usa_review_pending = usa_review_na = 0

    if "Trustpilot Review" in uk_clean.columns and "Brand" in uk_clean.columns:
        uk_filtered = uk_clean[uk_clean["Brand"].isin(allowed_brands)]
        uk_review_sent = uk_filtered["Trustpilot Review"].value_counts().get("Sent", 0)
        uk_review_pending = uk_filtered["Trustpilot Review"].value_counts().get("Pending", 0)
        uk_review_na = uk_filtered["Trustpilot Review"].value_counts().get("Negative", 0)
    else:
        uk_review_sent = uk_review_pending = uk_review_na = 0
    combined_pending_sent = pd.concat([usa_clean, uk_clean], ignore_index=True)
    pending_sent_details = combined_pending_sent[
        ((combined_pending_sent["Trustpilot Review"] == "Sent") |
         (combined_pending_sent["Trustpilot Review"] == "Pending")) &
        (combined_pending_sent["Brand"].isin(allowed_brands))
        ]
    pending_sent_details = pending_sent_details[["Name", "Brand", "Project Manager", "Trustpilot Review", "Status"]]
    pending_sent_details.index = range(1, len(pending_sent_details) + 1)

    usa_reviews_df = load_reviews(sheet_usa, year, month)
    uk_reviews_df = load_reviews(sheet_uk, year, month)
    combined_data = safe_concat([usa_reviews_df, uk_reviews_df])

    if not usa_reviews_df.empty:
        usa_attained_pm = (
            usa_reviews_df[usa_reviews_df["Trustpilot Review"] == "Attained"]
            .groupby("Project Manager")["Trustpilot Review"]
            .count()
            .reset_index()
        )
        usa_attained_pm.columns = ["Project Manager", "Attained Reviews"]
        usa_attained_pm.index = range(1, len(usa_attained_pm) + 1)
        usa_total_attained = usa_attained_pm["Attained Reviews"].sum()

        usa_negative_pm = (
            usa_reviews_df[usa_reviews_df["Trustpilot Review"] == "Negative"]
            .groupby("Project Manager")["Trustpilot Review"]
            .count()
            .reset_index()
        )
        usa_negative_pm.columns = ["Project Manager", "Negative Reviews"]
        usa_negative_pm.index = range(1, len(usa_negative_pm) + 1)
        usa_total_negative = usa_negative_pm["Negative Reviews"].sum()
    else:
        usa_attained_pm = pd.DataFrame(columns=["Project Manager", "Attained Reviews"])
        usa_negative_pm = pd.DataFrame(columns=["Project Manager", "Negative Reviews"])
        usa_total_attained = 0
        usa_total_negative = 0

    if not uk_reviews_df.empty:
        uk_attained_pm = (
            uk_reviews_df[uk_reviews_df["Trustpilot Review"] == "Attained"]
            .groupby("Project Manager")["Trustpilot Review"]
            .count()
            .reset_index()
        )
        uk_attained_pm.columns = ["Project Manager", "Attained Reviews"]
        uk_attained_pm.index = range(1, len(uk_attained_pm) + 1)
        uk_total_attained = uk_attained_pm["Attained Reviews"].sum()

        uk_negative_pm = (
            uk_reviews_df[uk_reviews_df["Trustpilot Review"] == "Negative"]
            .groupby("Project Manager")["Trustpilot Review"]
            .count()
            .reset_index()
        )
        uk_negative_pm.columns = ["Project Manager", "Negative Reviews"]
        uk_negative_pm.index = range(1, len(uk_negative_pm) + 1)
        uk_total_negative = uk_negative_pm["Negative Reviews"].sum()
    else:
        uk_attained_pm = pd.DataFrame(columns=["Project Manager", "Attained Reviews"])
        uk_negative_pm = pd.DataFrame(columns=["Project Manager", "Negative Reviews"])
        uk_total_attained = 0
        uk_total_negative = 0

    if not combined_data.empty:
        attained_reviews_per_pm = (
            combined_data[combined_data["Trustpilot Review"] == "Attained"]
            .groupby("Project Manager")["Trustpilot Review"]
            .count()
            .reset_index(name="Attained Reviews")
        )
        attained_reviews_per_pm = attained_reviews_per_pm.sort_values(by="Attained Reviews", ascending=False)
        attained_reviews_per_pm.index = range(1, len(attained_reviews_per_pm) + 1)

        negative_reviews_per_pm = (
            combined_data[combined_data["Trustpilot Review"] == "Negative"]
            .groupby("Project Manager")["Trustpilot Review"]
            .count()
            .reset_index(name="Negative Reviews")
        )
        negative_reviews_per_pm = negative_reviews_per_pm.sort_values(by="Negative Reviews", ascending=False)
        negative_reviews_per_pm.index = range(1, len(negative_reviews_per_pm) + 1)

        review_details_df = combined_data.sort_values(by="Project Manager", ascending=True)
        review_details_df["Trustpilot Review Date"] = pd.to_datetime(
            review_details_df["Trustpilot Review Date"], errors="coerce"
        ).dt.strftime("%d-%B-%Y")

        attained_details = review_details_df[
            review_details_df["Trustpilot Review"] == "Attained"
            ][["Project Manager", "Name", "Brand", "Trustpilot Review Date", "Trustpilot Review Links", "Status"]]
        attained_details.index = range(1, len(attained_details) + 1)

        negative_details = review_details_df[
            review_details_df["Trustpilot Review"] == "Negative"
            ][["Project Manager", "Name", "Brand", "Trustpilot Review Date", "Trustpilot Review Links", "Status"]]
        negative_details.index = range(1, len(negative_details) + 1)

    else:
        attained_reviews_per_pm = pd.DataFrame(columns=["Project Manager", "Attained Reviews"])
        negative_reviews_per_pm = pd.DataFrame(columns=["Project Manager", "Negative Reviews"])
        attained_details = pd.DataFrame(
            columns=["Project Manager", "Name", "Brand", "Trustpilot Review Date", "Trustpilot Review Links", "Status"])
        negative_details = attained_details.copy()

    usa_review = {
        "Attained": usa_total_attained,
        "Sent": usa_review_sent,
        "Pending": usa_review_pending,
        "Negative": usa_review_na + usa_total_negative
    }

    uk_review = {
        "Attained": uk_total_attained,
        "Sent": uk_review_sent,
        "Pending": uk_review_pending,
        "Negative": uk_review_na + uk_total_negative
    }

    printing_data = get_printing_data_month(month, year)
    Total_copies = printing_data["No of Copies"].sum() if "No of Copies" in printing_data.columns else 0
    Total_cost = printing_data["Order Cost"].sum() if "Order Cost" in printing_data.columns else 0
    Highest_cost = printing_data["Order Cost"].max() if "Order Cost" in printing_data.columns else 0
    Highest_copies = printing_data["No of Copies"].max() if "No of Copies" in printing_data.columns else 0
    Lowest_cost = printing_data["Order Cost"].min() if "Order Cost" in printing_data.columns else 0
    Lowest_copies = printing_data["No of Copies"].min() if "No of Copies" in printing_data.columns else 0

    Average = Total_cost / Total_copies if Total_copies > 0 else 0
    if all(col in printing_data.columns for col in ["Order Cost", "No of Copies"]):
        printing_data['Cost_Per_Copy'] = printing_data['Order Cost'] / printing_data['No of Copies']

    copyright_data, result_count, result_count_no = get_copyright_month(month, year)
    Total_copyrights = len(copyright_data)

    country = copyright_data["Country"].value_counts()
    usa = country.get("USA", 0)
    canada = country.get("Canada", 0)
    uk = country.get("UK", 0)
    Total_cost_copyright = (usa * 65) + (canada * 46) + (uk * 42)
    a_plus, a_plus_count = get_A_plus_month(month, year)

    usa_brands = {'BookMarketeers': bookmarketeers, 'Writers Clique': writers_clique, 'KDP': kdp,
                  'Aurora Writers': aurora_writers}

    uk_brands = {'Authors Solution': authors_solution, 'Book Publication': book_publication}

    usa_platforms = {'Amazon': usa_amazon, 'Barnes & Noble': usa_bn, 'Ingram Spark': usa_ingram,"Draft2Digital":usa_d2d,"Kobo": usa_kobo, "LULU":usa_lulu, "FAV": usa_fav, "ACX": usa_acx}
    uk_platforms = {'Amazon': uk_amazon, 'Barnes & Noble': uk_bn, 'Ingram Spark': uk_ingram, "Draft2Digital":uk_d2d,"Kobo": uk_kobo,"LULU":uk_lulu, "FAV": uk_fav,
                     "ACX": uk_acx}

    printing_stats = {
        'Total_copies': Total_copies,
        'Total_cost': Total_cost,
        'Highest_cost': Highest_cost,
        'Lowest_cost': Lowest_cost,
        'Highest_copies': Highest_copies,
        'Lowest_copies': Lowest_copies,
        'Average': Average
    }

    copyright_stats = {
        'Total_copyrights': Total_copyrights,
        'Total_cost_copyright': Total_cost_copyright,
        'result_count': result_count,
        'result_count_no': result_count_no,
        'usa_copyrights': usa,
        'canada_copyrights': canada,
        'uk': uk
    }

    return usa_review, uk_review, usa_brands, uk_brands, usa_platforms, uk_platforms, printing_stats, copyright_stats, a_plus_count, total_unique_clients, combined, attained_reviews_per_pm, attained_details, pending_sent_details, negative_reviews_per_pm, negative_details, Issues_usa, Issues_uk


def generate_year_summary(year: int):
    uk_clean = clean_data_reviews(sheet_uk)
    usa_clean = clean_data_reviews(sheet_usa)

    usa_clean = usa_clean[
        (usa_clean["Publishing Date"].dt.year == year)
    ]
    uk_clean = uk_clean[
        (uk_clean["Publishing Date"].dt.year == year)
    ]

    usa_clean_platforms = usa_clean[
        (usa_clean["Publishing Date"].dt.year == year)
    ]
    uk_clean_platforms = uk_clean[
        (uk_clean["Publishing Date"].dt.year == year)
    ]

    if usa_clean.empty:
        print("No values found in USA sheet.")
    if uk_clean.empty:
        print("No values found in UK sheet.")
        return
    if usa_clean.empty and uk_clean.empty:
        return

    usa_clean = usa_clean.drop_duplicates(subset=["Name"], keep="first")
    uk_clean = uk_clean.drop_duplicates(subset=["Name"], keep="first")
    Issues_usa = usa_clean["Issues"].value_counts()
    Issues_uk = uk_clean["Issues"].value_counts()
    total_usa = usa_clean["Name"].nunique()
    total_uk = uk_clean["Name"].nunique()
    total_unique_clients = total_usa + total_uk

    combined = pd.concat([usa_clean[["Name", "Brand", "Project Manager", "Email"]],
                          uk_clean[["Name", "Brand", "Project Manager", "Email"]]])
    combined.index = range(1, len(combined) + 1)

    brands = usa_clean["Brand"].value_counts()
    writers_clique = brands.get("Writers Clique", 0)
    bookmarketeers = brands.get("BookMarketeers", 0)
    aurora_writers = brands.get("Aurora Writers", 0)
    kdp = brands.get("KDP", 0)

    uk_brand = uk_clean["Brand"].value_counts()
    authors_solution = uk_brand.get("Authors Solution", 0)
    book_publication = uk_brand.get("Book Publication", 0)

    usa_platforms = usa_clean_platforms["Platform"].value_counts()
    usa_amazon = usa_platforms.get("Amazon", 0)
    usa_bn = usa_platforms.get("Barnes & Noble", 0)
    usa_ingram = usa_platforms.get("Ingram Spark", 0)
    usa_d2d = usa_platforms.get("Draft2Digital", 0)
    usa_lulu = usa_platforms.get("LULU", 0)
    usa_kobo = usa_platforms.get("Kobo", 0)
    usa_fav = usa_platforms.get("FAV", 0)
    usa_acx = usa_platforms.get("ACX", 0)

    uk_platforms = uk_clean_platforms["Platform"].value_counts()
    uk_amazon = uk_platforms.get("Amazon", 0)
    uk_bn = uk_platforms.get("Barnes & Noble", 0)
    uk_ingram = uk_platforms.get("Ingram Spark", 0)
    uk_d2d = uk_platforms.get("Draft2Digital", 0)
    uk_lulu = uk_platforms.get("LULU", 0)
    uk_fav = uk_platforms.get("FAV", 0)
    uk_kobo = uk_platforms.get("Kobo", 0)
    uk_acx = uk_platforms.get("ACX", 0)

    allowed_brands = ["BookMarketeers", "Writers Clique", "Aurora Writers", "Authors Solution", "Book Publication"]

    if "Trustpilot Review" in usa_clean.columns and "Brand" in usa_clean.columns:
        usa_filtered = usa_clean[usa_clean["Brand"].isin(allowed_brands)]
        usa_review_sent = usa_filtered["Trustpilot Review"].value_counts().get("Sent", 0)
        usa_review_pending = usa_filtered["Trustpilot Review"].value_counts().get("Pending", 0)
        usa_review_na = usa_filtered["Trustpilot Review"].value_counts().get("Negative", 0)
    else:
        usa_review_sent = usa_review_pending = usa_review_na = 0

    if "Trustpilot Review" in uk_clean.columns and "Brand" in uk_clean.columns:
        uk_filtered = uk_clean[uk_clean["Brand"].isin(allowed_brands)]
        uk_review_sent = uk_filtered["Trustpilot Review"].value_counts().get("Sent", 0)
        uk_review_pending = uk_filtered["Trustpilot Review"].value_counts().get("Pending", 0)
        uk_review_na = uk_filtered["Trustpilot Review"].value_counts().get("Negative", 0)
    else:
        uk_review_sent = uk_review_pending = uk_review_na = 0

    combined_pending_sent = pd.concat([usa_clean, uk_clean], ignore_index=True)
    pending_sent_details = combined_pending_sent[
        ((combined_pending_sent["Trustpilot Review"] == "Sent") |
         (combined_pending_sent["Trustpilot Review"] == "Pending")) &
        (combined_pending_sent["Brand"].isin(allowed_brands))
        ]
    pending_sent_details = pending_sent_details[["Name", "Brand", "Project Manager", "Trustpilot Review", "Status"]]
    pending_sent_details.index = range(1, len(pending_sent_details) + 1)

    pm_list_usa = list(set((usa_clean["Project Manager"].dropna().unique().tolist() + ["Unknown"])))
    pm_list_uk = list(set((uk_clean["Project Manager"].dropna().unique().tolist() + ["Unknown"])))

    usa_reviews_per_pm = safe_concat([load_reviews_year(sheet_usa, year, pm, "Attained") for pm in pm_list_usa])
    uk_reviews_per_pm = safe_concat([load_reviews_year(sheet_uk, year, pm, "Attained") for pm in pm_list_uk])
    combined_data = safe_concat([usa_reviews_per_pm, uk_reviews_per_pm])

    usa_monthly = (
        usa_clean.groupby(usa_clean["Publishing Date"].dt.to_period("M"))
        .size()
        .reset_index(name="USA Published")
    )
    usa_monthly["Month"] = usa_monthly["Publishing Date"].dt.strftime("%B %Y")
    usa_monthly = usa_monthly[["Month", "USA Published"]]

    uk_monthly = (
        uk_clean.groupby(uk_clean["Publishing Date"].dt.to_period("M"))
        .size()
        .reset_index(name="UK Published")
    )
    uk_monthly["Month"] = uk_monthly["Publishing Date"].dt.strftime("%B %Y")
    uk_monthly = uk_monthly[["Month", "UK Published"]]

    combined_monthly = pd.merge(
        usa_monthly,
        uk_monthly,
        on="Month",
        how="outer"
    ).fillna(0)

    combined_monthly["Total Published"] = combined_monthly["USA Published"] + combined_monthly["UK Published"]

    combined_monthly["Month_Num"] = pd.to_datetime(combined_monthly["Month"], format="%B %Y")
    combined_monthly = combined_monthly.sort_values("Total Published", ascending=False).drop(columns="Month_Num")

    combined_monthly.index = range(1, len(combined_monthly) + 1)

    if not usa_reviews_per_pm.empty:
        usa_attained_pm = (
            usa_reviews_per_pm
            .groupby("Project Manager")["Trustpilot Review"]
            .count()
            .reset_index()
        )
        usa_attained_pm.columns = ["Project Manager", "Attained Reviews"]
        usa_attained_pm.index = range(1, len(usa_attained_pm) + 1)
        usa_total_attained = usa_attained_pm["Attained Reviews"].sum()
    else:
        usa_attained_pm = pd.DataFrame(columns=["Project Manager", "Attained Reviews"])
        usa_total_attained = 0

    if not uk_reviews_per_pm.empty:
        uk_attained_pm = (
            uk_reviews_per_pm
            .groupby("Project Manager")["Trustpilot Review"]
            .count()
            .reset_index()
        )
        uk_attained_pm.columns = ["Project Manager", "Attained Reviews"]
        uk_attained_pm.index = range(1, len(uk_attained_pm) + 1)
        uk_total_attained = uk_attained_pm["Attained Reviews"].sum()
    else:
        uk_attained_pm = pd.DataFrame(columns=["Project Manager", "Attained Reviews"])
        uk_total_attained = 0

    if not combined_data.empty:
        attained_reviews_per_pm = (
            combined_data
            .groupby("Project Manager")["Trustpilot Review"]
            .count()
            .reset_index()
        )
        attained_reviews_per_pm.columns = ["Project Manager", "Attained Reviews"]
        attained_reviews_per_pm = attained_reviews_per_pm.sort_values(by="Attained Reviews", ascending=False)
        attained_reviews_per_pm.index = range(1, len(attained_reviews_per_pm) + 1)

        review_details_df = combined_data.sort_values(by="Project Manager", ascending=True)
        review_details_df["Trustpilot Review Date"] = pd.to_datetime(
            review_details_df["Trustpilot Review Date"], errors="coerce"
        ).dt.strftime("%d-%B-%Y")

        attained_details = review_details_df[
            ["Project Manager", "Name", "Brand", "Trustpilot Review Date", "Trustpilot Review Links", "Status"]
        ]
        attained_details.index = range(1, len(attained_details) + 1)

        attained_details["Trustpilot Review Date"] = pd.to_datetime(
            attained_details["Trustpilot Review Date"], errors="coerce"
        )
        attained_count = (
            attained_details
            .groupby("Project Manager")
            .size()
            .reset_index(name="Count")
        )
        attained_clients = (
            attained_details
            .groupby("Project Manager")
            ["Name"].apply(list)
            .reset_index(name="Clients")
        )
        merged_attained = attained_count.merge(attained_clients, on="Project Manager", how="left")
        merged_attained = merged_attained.sort_values(by="Count", ascending=False)
        merged_attained.index = range(1, len(merged_attained) + 1)
        if not usa_reviews_per_pm.empty:
            usa_attained_monthly = (
                usa_reviews_per_pm.groupby(usa_reviews_per_pm["Trustpilot Review Date"].dt.to_period("M"))
                .size()
                .reset_index(name="USA Attained Reviews")
            )
            usa_attained_monthly["Month"] = usa_attained_monthly["Trustpilot Review Date"].dt.strftime("%B %Y")
            usa_attained_monthly = usa_attained_monthly[["Month", "USA Attained Reviews"]]
        else:
            usa_attained_monthly = pd.DataFrame(columns=["Month", "USA Attained Reviews"])

        if not uk_reviews_per_pm.empty:
            uk_attained_monthly = (
                uk_reviews_per_pm.groupby(uk_reviews_per_pm["Trustpilot Review Date"].dt.to_period("M"))
                .size()
                .reset_index(name="UK Attained Reviews")
            )
            uk_attained_monthly["Month"] = uk_attained_monthly["Trustpilot Review Date"].dt.strftime("%B %Y")
            uk_attained_monthly = uk_attained_monthly[["Month", "UK Attained Reviews"]]
        else:
            uk_attained_monthly = pd.DataFrame(columns=["Month", "UK Attained Reviews"])
        attained_reviews_per_month = pd.merge(
            usa_attained_monthly,
            uk_attained_monthly,
            on="Month",
            how="outer"
        ).fillna(0)

        attained_reviews_per_month["Total Attained Reviews"] = (
                attained_reviews_per_month["USA Attained Reviews"] + attained_reviews_per_month["UK Attained Reviews"]
        )

        attained_reviews_per_month["Month_Num"] = pd.to_datetime(attained_reviews_per_month["Month"], format="%B %Y")
        attained_reviews_per_month = attained_reviews_per_month.sort_values(by="Total Attained Reviews",
                                                                            ascending=False)
        attained_reviews_per_month.index = range(1, len(attained_reviews_per_month) + 1)
        attained_reviews_per_month = attained_reviews_per_month.drop(columns="Month_Num")

        attained_details["Trustpilot Review Date"] = pd.to_datetime(
            attained_details["Trustpilot Review Date"], errors="coerce"
        ).dt.strftime("%d-%B-%Y")

    else:
        attained_reviews_per_pm = pd.DataFrame(columns=["Project Manager", "Attained Reviews"])
        attained_details = pd.DataFrame(
            columns=["Project Manager", "Name", "Brand", "Trustpilot Review Date", "Trustpilot Review Links", "Status"])
        attained_reviews_per_month = pd.DataFrame(columns=["Month", "Total Attained Reviews"])

    usa_negative_per_pm = [load_reviews_year(sheet_usa, year, pm, "Negative") for pm in pm_list_usa]
    usa_negative_per_pm = safe_concat([df for df in usa_negative_per_pm if not df.empty])

    uk_negative_per_pm = [load_reviews_year(sheet_uk, year, pm, "Negative") for pm in pm_list_uk]
    uk_negative_per_pm = safe_concat([df for df in uk_negative_per_pm if not df.empty])

    combined_negative_data = safe_concat([usa_negative_per_pm, uk_negative_per_pm])

    if not usa_negative_per_pm.empty:
        usa_negative_pm = (
            usa_negative_per_pm
            .groupby("Project Manager")["Trustpilot Review"]
            .count()
            .reset_index()
        )
        usa_negative_pm.columns = ["Project Manager", "Negative Reviews"]
        usa_negative_pm.index = range(1, len(usa_negative_pm) + 1)
        usa_total_negative = usa_negative_pm["Negative Reviews"].sum()
    else:
        usa_negative_pm = pd.DataFrame(columns=["Project Manager", "Negative Reviews"])
        usa_total_negative = 0

    if not uk_negative_per_pm.empty:
        uk_negative_pm = (
            uk_negative_per_pm
            .groupby("Project Manager")["Trustpilot Review"]
            .count()
            .reset_index()
        )
        uk_negative_pm.columns = ["Project Manager", "Negative Reviews"]
        uk_negative_pm.index = range(1, len(uk_negative_pm) + 1)
        uk_total_negative = uk_negative_pm["Negative Reviews"].sum()
    else:
        uk_negative_pm = pd.DataFrame(columns=["Project Manager", "Negative Reviews"])
        uk_total_negative = 0

    if not combined_negative_data.empty:

        negative_reviews_per_pm = (
            combined_negative_data
            .groupby("Project Manager")["Trustpilot Review"]
            .count()
            .reset_index()
        )
        negative_reviews_per_pm.columns = ["Project Manager", "Negative Reviews"]
        negative_reviews_per_pm = negative_reviews_per_pm.sort_values(by="Negative Reviews", ascending=False)
        negative_reviews_per_pm.index = range(1, len(negative_reviews_per_pm) + 1)

        negative_details_df = combined_negative_data.sort_values(by="Project Manager", ascending=True)
        negative_details_df["Trustpilot Review Date"] = pd.to_datetime(
            negative_details_df["Trustpilot Review Date"], errors="coerce"
        ).dt.strftime("%d-%B-%Y")

        negative_details = negative_details_df[
            ["Project Manager", "Name", "Brand", "Trustpilot Review Date", "Trustpilot Review Links", "Status"]
        ]
        negative_details.index = range(1, len(negative_details) + 1)

        negative_details["Trustpilot Review Date"] = pd.to_datetime(
            negative_details["Trustpilot Review Date"], errors="coerce"
        )

        if not usa_negative_per_pm.empty:
            usa_negative_monthly = (
                usa_negative_per_pm.groupby(usa_negative_per_pm["Trustpilot Review Date"].dt.to_period("M"))
                .size()
                .reset_index(name="USA Negative Reviews")
            )
            usa_negative_monthly["Month"] = usa_negative_monthly["Trustpilot Review Date"].dt.strftime("%B %Y")
            usa_negative_monthly = usa_negative_monthly[["Month", "USA Negative Reviews"]]
        else:
            usa_negative_monthly = pd.DataFrame(columns=["Month", "USA Negative Reviews"])

        # UK monthly negative reviews
        if not uk_negative_per_pm.empty:
            uk_negative_monthly = (
                uk_negative_per_pm.groupby(uk_negative_per_pm["Trustpilot Review Date"].dt.to_period("M"))
                .size()
                .reset_index(name="UK Negative Reviews")
            )
            uk_negative_monthly["Month"] = uk_negative_monthly["Trustpilot Review Date"].dt.strftime("%B %Y")
            uk_negative_monthly = uk_negative_monthly[["Month", "UK Negative Reviews"]]
        else:
            uk_negative_monthly = pd.DataFrame(columns=["Month", "UK Negative Reviews"])

        # Merge USA and UK negative trends
        negative_reviews_per_month = pd.merge(
            usa_negative_monthly,
            uk_negative_monthly,
            on="Month",
            how="outer"
        ).fillna(0)

        negative_reviews_per_month["Total Negative Reviews"] = (
                negative_reviews_per_month["USA Negative Reviews"] + negative_reviews_per_month["UK Negative Reviews"]
        )

        # Sort by month
        negative_reviews_per_month["Month_Num"] = pd.to_datetime(negative_reviews_per_month["Month"], format="%B %Y")
        negative_reviews_per_month = negative_reviews_per_month.sort_values(by="Total Negative Reviews",
                                                                            ascending=False)
        negative_reviews_per_month.index = range(1, len(negative_reviews_per_month) + 1)
        negative_reviews_per_month = negative_reviews_per_month.drop(columns="Month_Num")
        negative_details["Trustpilot Review Date"] = pd.to_datetime(
            negative_details["Trustpilot Review Date"], errors="coerce"
        ).dt.strftime("%d-%B-%Y")

    else:
        negative_reviews_per_pm = pd.DataFrame(columns=["Project Manager", "Negative Reviews"])
        negative_details = pd.DataFrame(
            columns=["Project Manager", "Name", "Brand", "Trustpilot Review Date", "Trustpilot Review Links", "Status"])
        negative_reviews_per_month = pd.DataFrame(columns=["Month", "Total Negative Reviews"])

    usa_review = {
        "Attained": usa_total_attained,
        "Sent": usa_review_sent,
        "Pending": usa_review_pending,
        "Negative": usa_total_negative
    }

    uk_review = {
        "Attained": uk_total_attained,
        "Sent": uk_review_sent,
        "Pending": uk_review_pending,
        "Negative": uk_total_negative
    }

    printing_data, monthly_printing = printing_data_year(year)
    Total_copies = printing_data["No of Copies"].sum() if "No of Copies" in printing_data.columns else 0
    Total_cost = printing_data["Order Cost"].sum() if "Order Cost" in printing_data.columns else 0
    Highest_cost = printing_data["Order Cost"].max() if "Order Cost" in printing_data.columns else 0
    Highest_copies = printing_data["No of Copies"].max() if "No of Copies" in printing_data.columns else 0
    Lowest_cost = printing_data["Order Cost"].min() if "Order Cost" in printing_data.columns else 0
    Lowest_copies = printing_data["No of Copies"].min() if "No of Copies" in printing_data.columns else 0

    Average = Total_cost / Total_copies if Total_copies > 0 else 0
    if all(col in printing_data.columns for col in ["Order Cost", "No of Copies"]):
        printing_data['Cost_Per_Copy'] = printing_data['Order Cost'] / printing_data['No of Copies']

    copyright_data, result_count, result_count_no = copyright_year(year)
    Total_copyrights = len(copyright_data)
    country = copyright_data["Country"].value_counts()
    usa = country.get("USA", 0)
    canada = country.get("Canada", 0)
    uk = country.get("UK", 0)
    Total_cost_copyright = (usa * 65) + (canada * 46) + (uk * 42)

    a_plus, a_plus_count = get_A_plus_year(year)

    usa_brands = {'BookMarketeers': bookmarketeers, 'Writers Clique': writers_clique, 'KDP': kdp,
                  'Aurora Writers': aurora_writers}
    uk_brands = {'Authors Solution': authors_solution, 'Book Publication': book_publication}

    usa_platforms = {'Amazon': usa_amazon, 'Barnes & Noble': usa_bn, 'Ingram Spark': usa_ingram,"Draft2Digital":usa_d2d,"Kobo": usa_kobo, "LULU":usa_lulu, "FAV": usa_fav, "ACX": usa_acx}
    uk_platforms = {'Amazon': uk_amazon, 'Barnes & Noble': uk_bn, 'Ingram Spark': uk_ingram, "Draft2Digital":uk_d2d,"Kobo": uk_kobo,"LULU":uk_lulu, "FAV": uk_fav,
                     "ACX": uk_acx}

    printing_stats = {
        'Total_copies': Total_copies,
        'Total_cost': Total_cost,
        'Highest_cost': Highest_cost,
        'Lowest_cost': Lowest_cost,
        'Highest_copies': Highest_copies,
        'Lowest_copies': Lowest_copies,
        'Average': Average
    }

    copyright_stats = {
        'Total_copyrights': Total_copyrights,
        'Total_cost_copyright': Total_cost_copyright,
        'result_count': result_count,
        'result_count_no': result_count_no,
        'usa_copyrights': usa,
        'canada_copyrights': canada,
        'uk': uk
    }

    return usa_review, uk_review, usa_brands, uk_brands, usa_platforms, uk_platforms, printing_stats, monthly_printing, copyright_stats, a_plus_count, total_unique_clients, combined, attained_reviews_per_pm, attained_details, merged_attained,  attained_reviews_per_month, pending_sent_details, negative_reviews_per_pm, negative_details, negative_reviews_per_month, combined_monthly, Issues_usa, Issues_uk


def generate_year_summary_multiple(start_year: int, end_year: int):
    uk_clean = clean_data_reviews(sheet_uk)
    usa_clean = clean_data_reviews(sheet_usa)

    usa_clean = usa_clean[
        (usa_clean["Publishing Date"].dt.year >= start_year) &
        (usa_clean["Publishing Date"].dt.year <= end_year)

    ]
    uk_clean = uk_clean[
        (uk_clean["Publishing Date"].dt.year >= start_year) &
        (uk_clean["Publishing Date"].dt.year <= end_year)
    ]

    usa_clean_platforms = usa_clean[
        (usa_clean["Publishing Date"].dt.year >= start_year) &
        (usa_clean["Publishing Date"].dt.year <= end_year)
        ]
    uk_clean_platforms = uk_clean[
        (uk_clean["Publishing Date"].dt.year >= start_year) &
        (uk_clean["Publishing Date"].dt.year <= end_year)
    ]

    if usa_clean.empty:
        print("No values found in USA sheet.")
    if uk_clean.empty:
        print("No values found in UK sheet.")
        return
    if usa_clean.empty and uk_clean.empty:
        return

    usa_clean = usa_clean.drop_duplicates(subset=["Name"], keep="first")
    uk_clean = uk_clean.drop_duplicates(subset=["Name"], keep="first")
    Issues_usa = usa_clean["Issues"].value_counts()
    Issues_uk = uk_clean["Issues"].value_counts()
    total_usa = usa_clean["Name"].nunique()
    total_uk = uk_clean["Name"].nunique()
    total_unique_clients = total_usa + total_uk

    combined = pd.concat([usa_clean[["Name", "Brand", "Project Manager", "Email"]],
                          uk_clean[["Name", "Brand", "Project Manager", "Email"]]])
    combined.index = range(1, len(combined) + 1)

    brands = usa_clean["Brand"].value_counts()
    writers_clique = brands.get("Writers Clique", 0)
    bookmarketeers = brands.get("BookMarketeers", 0)
    aurora_writers = brands.get("Aurora Writers", 0)
    kdp = brands.get("KDP", 0)

    uk_brand = uk_clean["Brand"].value_counts()
    authors_solution = uk_brand.get("Authors Solution", 0)
    book_publication = uk_brand.get("Book Publication", 0)

    usa_platforms = usa_clean_platforms["Platform"].value_counts()
    usa_amazon = usa_platforms.get("Amazon", 0)
    usa_bn = usa_platforms.get("Barnes & Noble", 0)
    usa_ingram = usa_platforms.get("Ingram Spark", 0)
    usa_d2d = usa_platforms.get("Draft2Digital", 0)
    usa_lulu = usa_platforms.get("LULU", 0)
    usa_fav = usa_platforms.get("FAV", 0)
    usa_kobo = usa_platforms.get("Kobo", 0)
    usa_acx = usa_platforms.get("ACX", 0)

    uk_platforms = uk_clean_platforms["Platform"].value_counts()
    uk_amazon = uk_platforms.get("Amazon", 0)
    uk_bn = uk_platforms.get("Barnes & Noble", 0)
    uk_ingram = uk_platforms.get("Ingram Spark", 0)
    uk_d2d = uk_platforms.get("Draft2Digital", 0)
    uk_lulu = uk_platforms.get("LULU", 0)
    uk_fav = uk_platforms.get("FAV", 0)
    uk_kobo = uk_platforms.get("Kobo", 0)
    uk_acx = uk_platforms.get("ACX", 0)

    allowed_brands = ["BookMarketeers", "Writers Clique", "Aurora Writers", "Authors Solution", "Book Publication"]

    if "Trustpilot Review" in usa_clean.columns and "Brand" in usa_clean.columns:
        usa_filtered = usa_clean[usa_clean["Brand"].isin(allowed_brands)]
        usa_review_sent = usa_filtered["Trustpilot Review"].value_counts().get("Sent", 0)
        usa_review_pending = usa_filtered["Trustpilot Review"].value_counts().get("Pending", 0)
        usa_review_na = usa_filtered["Trustpilot Review"].value_counts().get("Negative", 0)
    else:
        usa_review_sent = usa_review_pending = usa_review_na = 0

    if "Trustpilot Review" in uk_clean.columns and "Brand" in uk_clean.columns:
        uk_filtered = uk_clean[uk_clean["Brand"].isin(allowed_brands)]
        uk_review_sent = uk_filtered["Trustpilot Review"].value_counts().get("Sent", 0)
        uk_review_pending = uk_filtered["Trustpilot Review"].value_counts().get("Pending", 0)
        uk_review_na = uk_filtered["Trustpilot Review"].value_counts().get("Negative", 0)
    else:
        uk_review_sent = uk_review_pending = uk_review_na = 0

    combined_pending_sent = pd.concat([usa_clean, uk_clean], ignore_index=True)
    pending_sent_details = combined_pending_sent[
        ((combined_pending_sent["Trustpilot Review"] == "Sent") |
         (combined_pending_sent["Trustpilot Review"] == "Pending")) &
        (combined_pending_sent["Brand"].isin(allowed_brands))
        ]
    pending_sent_details = pending_sent_details[["Name", "Brand", "Project Manager", "Trustpilot Review", "Status"]]
    pending_sent_details.index = range(1, len(pending_sent_details) + 1)

    pm_list_usa = list(set((usa_clean["Project Manager"].dropna().unique().tolist() + ["Unknown"])))
    pm_list_uk = list(set((uk_clean["Project Manager"].dropna().unique().tolist() + ["Unknown"])))

    usa_reviews_per_pm = safe_concat([load_reviews_year_multiple(sheet_usa, start_year, end_year, pm, "Attained") for pm in pm_list_usa])
    uk_reviews_per_pm = safe_concat([load_reviews_year_multiple(sheet_uk, start_year, end_year, pm, "Attained") for pm in pm_list_uk])
    combined_data = safe_concat([usa_reviews_per_pm, uk_reviews_per_pm])

    usa_monthly = (
        usa_clean.groupby(usa_clean["Publishing Date"].dt.to_period("M"))
        .size()
        .reset_index(name="USA Published")
    )
    usa_monthly["Month"] = usa_monthly["Publishing Date"].dt.strftime("%B %Y")
    usa_monthly = usa_monthly[["Month", "USA Published"]]

    uk_monthly = (
        uk_clean.groupby(uk_clean["Publishing Date"].dt.to_period("M"))
        .size()
        .reset_index(name="UK Published")
    )
    uk_monthly["Month"] = uk_monthly["Publishing Date"].dt.strftime("%B %Y")
    uk_monthly = uk_monthly[["Month", "UK Published"]]

    combined_monthly = pd.merge(
        usa_monthly,
        uk_monthly,
        on="Month",
        how="outer"
    ).fillna(0)

    combined_monthly["Total Published"] = combined_monthly["USA Published"] + combined_monthly["UK Published"]

    combined_monthly["Month_Num"] = pd.to_datetime(combined_monthly["Month"], format="%B %Y")
    combined_monthly = combined_monthly.sort_values("Total Published", ascending=False).drop(columns="Month_Num")

    combined_monthly.index = range(1, len(combined_monthly) + 1)

    if not usa_reviews_per_pm.empty:
        usa_attained_pm = (
            usa_reviews_per_pm
            .groupby("Project Manager")["Trustpilot Review"]
            .count()
            .reset_index()
        )
        usa_attained_pm.columns = ["Project Manager", "Attained Reviews"]
        usa_attained_pm.index = range(1, len(usa_attained_pm) + 1)
        usa_total_attained = usa_attained_pm["Attained Reviews"].sum()
    else:
        usa_attained_pm = pd.DataFrame(columns=["Project Manager", "Attained Reviews"])
        usa_total_attained = 0

    if not uk_reviews_per_pm.empty:
        uk_attained_pm = (
            uk_reviews_per_pm
            .groupby("Project Manager")["Trustpilot Review"]
            .count()
            .reset_index()
        )
        uk_attained_pm.columns = ["Project Manager", "Attained Reviews"]
        uk_attained_pm.index = range(1, len(uk_attained_pm) + 1)
        uk_total_attained = uk_attained_pm["Attained Reviews"].sum()
    else:
        uk_attained_pm = pd.DataFrame(columns=["Project Manager", "Attained Reviews"])
        uk_total_attained = 0

    if not combined_data.empty:
        attained_reviews_per_pm = (
            combined_data
            .groupby("Project Manager")["Trustpilot Review"]
            .count()
            .reset_index()
        )
        attained_reviews_per_pm.columns = ["Project Manager", "Attained Reviews"]
        attained_reviews_per_pm = attained_reviews_per_pm.sort_values(by="Attained Reviews", ascending=False)
        attained_reviews_per_pm.index = range(1, len(attained_reviews_per_pm) + 1)

        review_details_df = combined_data.sort_values(by="Project Manager", ascending=True)
        review_details_df["Trustpilot Review Date"] = pd.to_datetime(
            review_details_df["Trustpilot Review Date"], errors="coerce"
        ).dt.strftime("%d-%B-%Y")

        attained_details = review_details_df[
            ["Project Manager", "Name", "Brand", "Trustpilot Review Date", "Trustpilot Review Links", "Status"]
        ]
        attained_details.index = range(1, len(attained_details) + 1)

        attained_details["Trustpilot Review Date"] = pd.to_datetime(
            attained_details["Trustpilot Review Date"], errors="coerce"
        )
        attained_count = (
            attained_details
            .groupby("Project Manager")
            .size()
            .reset_index(name="Count")
        )
        attained_clients = (
            attained_details
            .groupby("Project Manager")
            ["Name"].apply(list)
            .reset_index(name="Clients")
        )
        merged_attained = attained_count.merge(attained_clients, on="Project Manager", how="left")
        merged_attained = merged_attained.sort_values(by="Count", ascending=False)
        merged_attained.index = range(1, len(merged_attained) + 1)

        if not usa_reviews_per_pm.empty:
            usa_attained_monthly = (
                usa_reviews_per_pm.groupby(usa_reviews_per_pm["Trustpilot Review Date"].dt.to_period("M"))
                .size()
                .reset_index(name="USA Attained Reviews")
            )
            usa_attained_monthly["Month"] = usa_attained_monthly["Trustpilot Review Date"].dt.strftime("%B %Y")
            usa_attained_monthly = usa_attained_monthly[["Month", "USA Attained Reviews"]]
        else:
            usa_attained_monthly = pd.DataFrame(columns=["Month", "USA Attained Reviews"])

        if not uk_reviews_per_pm.empty:
            uk_attained_monthly = (
                uk_reviews_per_pm.groupby(uk_reviews_per_pm["Trustpilot Review Date"].dt.to_period("M"))
                .size()
                .reset_index(name="UK Attained Reviews")
            )
            uk_attained_monthly["Month"] = uk_attained_monthly["Trustpilot Review Date"].dt.strftime("%B %Y")
            uk_attained_monthly = uk_attained_monthly[["Month", "UK Attained Reviews"]]
        else:
            uk_attained_monthly = pd.DataFrame(columns=["Month", "UK Attained Reviews"])
        attained_reviews_per_month = pd.merge(
            usa_attained_monthly,
            uk_attained_monthly,
            on="Month",
            how="outer"
        ).fillna(0)

        attained_reviews_per_month["Total Attained Reviews"] = (
                attained_reviews_per_month["USA Attained Reviews"] + attained_reviews_per_month["UK Attained Reviews"]
        )

        attained_reviews_per_month["Month_Num"] = pd.to_datetime(attained_reviews_per_month["Month"], format="%B %Y")
        attained_reviews_per_month = attained_reviews_per_month.sort_values(by="Total Attained Reviews",
                                                                            ascending=False)
        attained_reviews_per_month.index = range(1, len(attained_reviews_per_month) + 1)
        attained_reviews_per_month = attained_reviews_per_month.drop(columns="Month_Num")

        attained_details["Trustpilot Review Date"] = pd.to_datetime(
            attained_details["Trustpilot Review Date"], errors="coerce"
        ).dt.strftime("%d-%B-%Y")

    else:
        attained_reviews_per_pm = pd.DataFrame(columns=["Project Manager", "Attained Reviews"])
        attained_details = pd.DataFrame(
            columns=["Project Manager", "Name", "Brand", "Trustpilot Review Date", "Trustpilot Review Links", "Status"])
        attained_reviews_per_month = pd.DataFrame(columns=["Month", "Total Attained Reviews"])

    usa_negative_per_pm = [load_reviews_year_multiple(sheet_usa, start_year, end_year,  pm, "Negative") for pm in pm_list_usa]
    usa_negative_per_pm = safe_concat([df for df in usa_negative_per_pm if not df.empty])

    uk_negative_per_pm = [load_reviews_year_multiple(sheet_uk, start_year, end_year, pm, "Negative") for pm in pm_list_uk]
    uk_negative_per_pm = safe_concat([df for df in uk_negative_per_pm if not df.empty])

    combined_negative_data = safe_concat([usa_negative_per_pm, uk_negative_per_pm])

    if not usa_negative_per_pm.empty:
        usa_negative_pm = (
            usa_negative_per_pm
            .groupby("Project Manager")["Trustpilot Review"]
            .count()
            .reset_index()
        )
        usa_negative_pm.columns = ["Project Manager", "Negative Reviews"]
        usa_negative_pm.index = range(1, len(usa_negative_pm) + 1)
        usa_total_negative = usa_negative_pm["Negative Reviews"].sum()
    else:
        usa_negative_pm = pd.DataFrame(columns=["Project Manager", "Negative Reviews"])
        usa_total_negative = 0

    if not uk_negative_per_pm.empty:
        uk_negative_pm = (
            uk_negative_per_pm
            .groupby("Project Manager")["Trustpilot Review"]
            .count()
            .reset_index()
        )
        uk_negative_pm.columns = ["Project Manager", "Negative Reviews"]
        uk_negative_pm.index = range(1, len(uk_negative_pm) + 1)
        uk_total_negative = uk_negative_pm["Negative Reviews"].sum()
    else:
        uk_negative_pm = pd.DataFrame(columns=["Project Manager", "Negative Reviews"])
        uk_total_negative = 0

    if not combined_negative_data.empty:

        negative_reviews_per_pm = (
            combined_negative_data
            .groupby("Project Manager")["Trustpilot Review"]
            .count()
            .reset_index()
        )
        negative_reviews_per_pm.columns = ["Project Manager", "Negative Reviews"]
        negative_reviews_per_pm = negative_reviews_per_pm.sort_values(by="Negative Reviews", ascending=False)
        negative_reviews_per_pm.index = range(1, len(negative_reviews_per_pm) + 1)

        negative_details_df = combined_negative_data.sort_values(by="Project Manager", ascending=True)
        negative_details_df["Trustpilot Review Date"] = pd.to_datetime(
            negative_details_df["Trustpilot Review Date"], errors="coerce"
        ).dt.strftime("%d-%B-%Y")

        negative_details = negative_details_df[
            ["Project Manager", "Name", "Brand", "Trustpilot Review Date", "Trustpilot Review Links", "Status"]
        ]
        negative_details.index = range(1, len(negative_details) + 1)

        negative_details["Trustpilot Review Date"] = pd.to_datetime(
            negative_details["Trustpilot Review Date"], errors="coerce"
        )

        if not usa_negative_per_pm.empty:
            usa_negative_monthly = (
                usa_negative_per_pm.groupby(usa_negative_per_pm["Trustpilot Review Date"].dt.to_period("M"))
                .size()
                .reset_index(name="USA Negative Reviews")
            )
            usa_negative_monthly["Month"] = usa_negative_monthly["Trustpilot Review Date"].dt.strftime("%B %Y")
            usa_negative_monthly = usa_negative_monthly[["Month", "USA Negative Reviews"]]
        else:
            usa_negative_monthly = pd.DataFrame(columns=["Month", "USA Negative Reviews"])

        # UK monthly negative reviews
        if not uk_negative_per_pm.empty:
            uk_negative_monthly = (
                uk_negative_per_pm.groupby(uk_negative_per_pm["Trustpilot Review Date"].dt.to_period("M"))
                .size()
                .reset_index(name="UK Negative Reviews")
            )
            uk_negative_monthly["Month"] = uk_negative_monthly["Trustpilot Review Date"].dt.strftime("%B %Y")
            uk_negative_monthly = uk_negative_monthly[["Month", "UK Negative Reviews"]]
        else:
            uk_negative_monthly = pd.DataFrame(columns=["Month", "UK Negative Reviews"])

        # Merge USA and UK negative trends
        negative_reviews_per_month = pd.merge(
            usa_negative_monthly,
            uk_negative_monthly,
            on="Month",
            how="outer"
        ).fillna(0)

        negative_reviews_per_month["Total Negative Reviews"] = (
                negative_reviews_per_month["USA Negative Reviews"] + negative_reviews_per_month["UK Negative Reviews"]
        )

        # Sort by month
        negative_reviews_per_month["Month_Num"] = pd.to_datetime(negative_reviews_per_month["Month"], format="%B %Y")
        negative_reviews_per_month = negative_reviews_per_month.sort_values(by="Total Negative Reviews",
                                                                            ascending=False)
        negative_reviews_per_month.index = range(1, len(negative_reviews_per_month) + 1)
        negative_reviews_per_month = negative_reviews_per_month.drop(columns="Month_Num")
        negative_details["Trustpilot Review Date"] = pd.to_datetime(
            negative_details["Trustpilot Review Date"], errors="coerce"
        ).dt.strftime("%d-%B-%Y")

    else:
        negative_reviews_per_pm = pd.DataFrame(columns=["Project Manager", "Negative Reviews"])
        negative_details = pd.DataFrame(
            columns=["Project Manager", "Name", "Brand", "Trustpilot Review Date", "Trustpilot Review Links", "Status"])
        negative_reviews_per_month = pd.DataFrame(columns=["Month", "Total Negative Reviews"])

    usa_review = {
        "Attained": usa_total_attained,
        "Sent": usa_review_sent,
        "Pending": usa_review_pending,
        "Negative": usa_total_negative
    }

    uk_review = {
        "Attained": uk_total_attained,
        "Sent": uk_review_sent,
        "Pending": uk_review_pending,
        "Negative": uk_total_negative
    }

    printing_data, monthly_printing = printing_data_year_multiple(start_year, end_year)
    Total_copies = printing_data["No of Copies"].sum() if "No of Copies" in printing_data.columns else 0
    Total_cost = printing_data["Order Cost"].sum() if "Order Cost" in printing_data.columns else 0
    Highest_cost = printing_data["Order Cost"].max() if "Order Cost" in printing_data.columns else 0
    Highest_copies = printing_data["No of Copies"].max() if "No of Copies" in printing_data.columns else 0
    Lowest_cost = printing_data["Order Cost"].min() if "Order Cost" in printing_data.columns else 0
    Lowest_copies = printing_data["No of Copies"].min() if "No of Copies" in printing_data.columns else 0

    Average = Total_cost / Total_copies if Total_copies > 0 else 0
    if all(col in printing_data.columns for col in ["Order Cost", "No of Copies"]):
        printing_data['Cost_Per_Copy'] = printing_data['Order Cost'] / printing_data['No of Copies']

    copyright_data, result_count, result_count_no = copyright_year_multiple(start_year, end_year)
    Total_copyrights = len(copyright_data)
    country = copyright_data["Country"].value_counts()
    usa = country.get("USA", 0)
    canada = country.get("Canada", 0)
    uk = country.get("UK", 0)
    Total_cost_copyright = (usa * 65) + (canada * 46) + (uk * 42)

    a_plus, a_plus_count = get_A_plus_year_multiple(start_year, end_year)

    usa_brands = {'BookMarketeers': bookmarketeers, 'Writers Clique': writers_clique, 'KDP': kdp,
                  'Aurora Writers': aurora_writers}
    uk_brands = {'Authors Solution': authors_solution, 'Book Publication': book_publication}

    usa_platforms = {'Amazon': usa_amazon, 'Barnes & Noble': usa_bn, 'Ingram Spark': usa_ingram,"Draft2Digital":usa_d2d,"Kobo": usa_kobo, "LULU":usa_lulu, "FAV": usa_fav, "ACX": usa_acx}
    uk_platforms = {'Amazon': uk_amazon, 'Barnes & Noble': uk_bn, 'Ingram Spark': uk_ingram, "Draft2Digital":uk_d2d,"Kobo": uk_kobo,"LULU":uk_lulu, "FAV": uk_fav,
                     "ACX": uk_acx}

    printing_stats = {
        'Total_copies': Total_copies,
        'Total_cost': Total_cost,
        'Highest_cost': Highest_cost,
        'Lowest_cost': Lowest_cost,
        'Highest_copies': Highest_copies,
        'Lowest_copies': Lowest_copies,
        'Average': Average
    }

    copyright_stats = {
        'Total_copyrights': Total_copyrights,
        'Total_cost_copyright': Total_cost_copyright,
        'result_count': result_count,
        'result_count_no': result_count_no,
        'usa_copyrights': usa,
        'canada_copyrights': canada,
        'uk': uk
    }

    return usa_review, uk_review, usa_brands, uk_brands, usa_platforms, uk_platforms, printing_stats, monthly_printing, copyright_stats, a_plus_count, total_unique_clients, combined, attained_reviews_per_pm, attained_details, merged_attained, attained_reviews_per_month, pending_sent_details, negative_reviews_per_pm, negative_details, negative_reviews_per_month, combined_monthly, Issues_usa, Issues_uk
