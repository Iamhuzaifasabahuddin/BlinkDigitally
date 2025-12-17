import calendar
import datetime
import json
import time
from http import HTTPStatus

import requests
import streamlit as st
from bs4 import BeautifulSoup
from requests import RequestException
from selenium import webdriver
from selenium.webdriver import EdgeOptions

def is_valid_url(url: str) -> bool:
    try:
        response = requests.get(url, timeout=5)
        return response.status_code == HTTPStatus.OK
    except RequestException:
        return False


def fetch_page_source(url):
    options = EdgeOptions()
    options.add_argument("--headless=new")
    driver = webdriver.Edge(options=options)
    driver.get(url)
    time.sleep(3)
    html = driver.page_source
    driver.quit()
    return html


def parse_reviews_from_html(html, max_reviews=5):
    soup = BeautifulSoup(html, "html.parser")
    script_tag = soup.find("script", type="application/json")
    if not script_tag:
        return []

    data_json = json.loads(script_tag.string)
    reviews_data = data_json.get("props", {}).get("pageProps", {}).get("reviews", [])

    reviews = []
    for r in reviews_data[:max_reviews]:
        reviews.append({
            "author": r.get("consumer", {}).get("displayName"),
            "title": r.get("title"),
            "rating": r.get("rating"),
            "publishedDate": r.get("dates", {}).get("publishedDate"),
            "content": r.get("text")
        })
    return reviews


@st.cache_data(ttl=3600)
def get_latest_trustpilot_reviews(url, max_reviews=5):
    if is_valid_url(url):
        html = fetch_page_source(url)
        reviews = parse_reviews_from_html(html, max_reviews=max_reviews)
        return reviews

st.title("Trustpilot Latest Reviews Scraper")

sites = {
    "Bookmarketeers": "https://www.trustpilot.com/review/bookmarketeers.com",
    "Writers Clique": "https://www.trustpilot.com/review/writersclique.com",
    "Aurora Writers": "https://www.trustpilot.com/review/aurorawriters.com",
    "Authors Solution": "https://www.trustpilot.com/review/authorssolution.co.uk",
    "Book Publication": "https://www.trustpilot.com/review/bookpublication.co.uk",
}

selected_site = st.selectbox("Select website to fetch reviews from:", list(sites.keys()))
num_reviews = st.slider("Number of latest reviews to fetch:", min_value=1, max_value=20, value=5)

months = ["All"] + list(calendar.month_name[1:])
current_month_index = datetime.datetime.now().month
selected_month_name = st.selectbox("Filter reviews by month:", months, index=current_month_index)

if st.button("Fetch Reviews"):
    url = sites[selected_site]
    with st.spinner(f"Fetching {num_reviews} latest reviews from {selected_site}..."):
        try:
            reviews = get_latest_trustpilot_reviews(url, max_reviews=100)
            if reviews:
                filtered_reviews = []
                for r in reviews:
                    iso_str = r['publishedDate']
                    dt = datetime.datetime.fromisoformat(iso_str.replace("Z", "+00:00"))

                    if selected_month_name == "All" or dt.strftime('%B') == selected_month_name:
                        r['date_obj'] = dt
                        filtered_reviews.append(r)

                # Limit after filtering
                filtered_reviews = filtered_reviews[:num_reviews]

                if filtered_reviews:
                    for i, r in enumerate(filtered_reviews, 1):
                        st.subheader(f"Review #{i}")
                        st.write(f"**Author:** {r['author']}")
                        st.write(f"**Rating:** {r['rating']} ⭐")
                        st.write(f"**Published:** {r['date_obj'].strftime('%d-%B-%Y')}")
                        st.write(f"**Title:** {r['title']}")
                        st.write(f"**Content:** {r['content']}")
                        st.markdown("---")
                else:
                    st.warning("No reviews found for the selected month.")
            else:
                st.warning("No reviews found or page JSON could not be parsed.")
        except Exception as e:
            st.error(f"Error: {e}")
