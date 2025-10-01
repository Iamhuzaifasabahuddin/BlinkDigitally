# Blink Digitally Publishing Dashboard v1.0.5

**Short description:**  
Blink Digitally publishing dashboard is a Streamlit-based dashboard that fetches, caches, cleans and visualizes book / review data from Google Sheets, generates monthly dynamic stats (pie charts, bar charts), produces downloadable PDF summaries, and pushes automated, prettified weekly client-review reminders and in-depth review tables to Slack channels (USA / UK). It includes data-management tools, an enhanced reviews view, and deployment-ready Streamlit configuration.

---

## Key Features
- Fetch Google Sheets data via Google Sheets API, using `gspread` + service account.
  - View Monthly USA / UK Data
  - View Printing Data & Search in Printing Data
  - View Copyright Data
  - Compare Client list in two different months
  - Get Monthly & Yearly Summary Data
- Local/function caching to limit API usage (default cache TTL: **2 minutes**).  
- Data cleaning and typing with `pandas`.  
- Dynamic monthly statistics: bar charts, pie charts, summary KPIs.  
- Enhanced reviews UI: full review text, rating, review date, origin (Trustpilot, other), brand metadata.  
- Status tracking: **Published**, **Pending**, **In Progress**, etc., with counts and trends.  
- PDF export: download full summary / reports as a prettified PDF.  
- Slack integration: weekly automated reminders to project managers with review summaries; sends prettified review tables to `#usa` and `#uk` channels with client, brand, origin, and link.  
- Deployment: app deployable on Streamlit Sharing / Streamlit Cloud (or any standard container).  
- Extensible: easy to add other data sources or notification channels.  

---

## Architecture Overview
- **Front-end UI**: Streamlit (interactive dashboard, charts, filters, export buttons).  
- **Data ingestion**: `gspread` to fetch Google Sheets rows, parsed into `pandas.DataFrame`.  
- **Caching layer**: Streamlit caching (`@st.cache_data`) with TTL = 120 seconds (2 minutes) to limit repeated Google Sheets API calls.  
- **Business logic**: cleaning, status slicing, metrics calculation, monthly aggregation.  
- **Analytics & charts**: `matplotlib` / `plotly` (or `altair`) to render bar charts, pie charts, time-series KPI visuals.  
- **Notifications**: Slack API (Slack SDK) for channels and direct messages; scheduled weekly jobs to send review reminders and tables.  
- **Export**: PDF generation (e.g., `reportlab`, `weasyprint`, or `pdfkit` + HTML templates) for download.  
- **Hosting**: Streamlit Cloud; optionally Docker container for other providers.  
