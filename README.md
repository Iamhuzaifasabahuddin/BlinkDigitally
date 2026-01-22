# 📄 Blink Digitally — Publishing Dashboard

Version: v1.2.0

Short description  
Blink Digitally publishing dashboard is a Streamlit-based dashboard that fetches, caches, cleans and visualizes book / review data from Google Sheets, generates monthly and yearly dynamic stats (pie charts, bar charts, KPIs), produces downloadable PDF summaries, and sends automated, prettified weekly client-review reminders and in-depth review tables to Slack channels (USA / UK). It includes data-management tools, an enhanced reviews view, and deployment-ready Streamlit configuration.

Repository & reference
- Repo: https://github.com/Iamhuzaifasabahuddin/BlinkDigitally
- Primary app file: App_Streamlit.py (source used for this README)

---

Table of contents
- Key Features
- Architecture Overview
- Requirements
- Configuration (st.secrets)
- PDF export
- Caching & performance
- Contributing
- License
- Changelog (v1.2.0)

---

Key Features
- Fetch Google Sheets data via Google Sheets API using `gspread` + service account
  - View Monthly USA / UK Data
  - View Printing Data & Search in Printing Data
  - View Copyright Data
  - Compare client lists across two months or years (similarity reports)
  - Monthly, Yearly, and multi-year summaries and custom-range reports
- Local / function caching to limit API usage (Streamlit cache usage; default TTL: 2 minutes)
- Data cleaning & typing with `pandas`
- Dynamic visualizations with `plotly` (pie charts, bar charts, KPIs, gauges)
- Enhanced reviews UI: full review text, rating, review date, origin (Trustpilot, other), brand metadata
- Status tracking: Published, Pending, In Progress, Printing Only, Self Publishing, etc., with counts and trends
- PDF export: prettified summary PDF using ReportLab
- Deployment-ready: deploy on Streamlit Cloud / Streamlit Sharing or containerize with Docker
- Extensible: designed so you can add additional data sources or notification targets

---

Architecture Overview
- Front-end UI: Streamlit (interactive dashboard, filters, export UI)
- Data ingestion: `gspread` reads Google Sheets → parsed into `pandas.DataFrame`
- Caching layer: Streamlit caching (`@st.cache_resource` for the gspread client and `@st.cache_data` for sheet reads) to reduce repeated Google Sheets calls
- Business logic: data cleaning, slicing, status filtering, aggregations and KPI calculations
- Analytics & charts: `plotly` used for interactive charts; `reportlab` used for PDF generation
- Export: PDF generation with ReportLab (prettified tables and KPIs)
- Hosting: Streamlit Cloud, or any container platform (Docker)

---

Requirements
Minimum tested Python packages (install in a virtualenv):
- Python 3.13+
- streamlit
- gspread
- google-auth (google-auth and google-auth-oauthlib if needed)
- pandas
- plotly
- reportlab
- pytz
- openpyxl (for Excel export)
- (optional) python-dotenv, schedule / APScheduler for local scheduled tasks

Install using pip:
```bash
pip install -r requirements.txt
```

---

Configuration (st.secrets)
The app uses Streamlit secrets to store the Google service account credentials and spreadsheet ID.

Example structure for Streamlit `st.secrets` (Streamlit Cloud / Sharing or `.streamlit/secrets.toml`):
```toml
[connections.gsheets]
type = "<service_account_type>"
project_id = "<project_id>"
private_key_id = "<private_key_id>"
private_key = "-----BEGIN PRIVATE KEY-----\n...REPLACED...\n-----END PRIVATE KEY-----\n"
client_email = "<account-email@project.iam.gserviceaccount.com>"
client_id = "<client_id>"
auth_uri = "https://accounts.google.com/o/oauth2/auth"
token_uri = "https://oauth2.googleapis.com/token"
auth_provider_x509_cert_url = "https://www.googleapis.com/oauth2/v1/certs"
client_x509_cert_url = "<cert_url>"
SPREADSHEET_ID = "<your_google_spreadsheet_key>"
```

Notes:
- The private key line breaks are handled in App_Streamlit.py by replacing `\\n` with `\n`. If you paste the key into Streamlit secrets, escape properly or use triple-quoted string in the UI.
- On Streamlit Cloud, add these values via the Secrets UI (Settings → Secrets).

---

Usage summary
- Use the top-left UI to choose actions:
  - View Data: View monthly/yearly reviews & detailed rows for USA/UK.
  - Printing: View & export printing orders; year-over-year comparisons.
  - Copyright: View application status and export lists.
  - Generate Similarity: find clients appearing in two months or two years.
  - Summary / Year Summary / Custom Summary: create dashboards and PDF exports.
  - Reviews: explore attained/pending/negative reviews per PM, monthly pivot tables.
  - Sales: monthly / yearly sales exports.
- Download: Excel and PDF download buttons are provided in the UI for each view.

---

PDF export
- Prettified PDF summaries are built using ReportLab in `generate_summary_report_pdf`.
- The PDF generator accepts a monthly or yearly (or multi-year) context and outputs:
  - Review analytics table & KPIs
  - Platform and brand breakdowns
  - Printing analytics and cost KPIs
  - Copyright analytics
  - A+ Content counts
- The app produces the PDF content in memory and returns a downloadable file.

---

Caching & performance
- The app caches the gspread client (`@st.cache_resource`) and sheet reads (`@st.cache_data`) with a TTL of 300 seconds in the code — you can adjust TTL to 120 seconds (2 minutes) if you prefer the documented default.
- Avoid lowering TTL too much because it increases Google Sheets API usage and may hit rate limits.
- When the "Fetch Latest" button is used, the Streamlit cache is cleared to force reload from Google Sheets.

---

Contributing
- Bug reports, feature requests and PRs welcome. Please open issues describing:
  - The steps to reproduce
  - Expected vs actual behavior
  - Any relevant logs or traceback
- Suggested improvements:
  - Add tests around data parsing & cleaning functions
  - Add unit tests for summary aggregation logic
  - Add Dockerfile for easier deployment portability

---

License
- Add your preferred license to the repository (e.g., MIT, Apache 2.0). This README does not include a license by default.

---

Changelog — v1.2.0
- Major improvements:
  - Improved PDF summary layout and additional KPIs
  - Enhanced reviews UI and better per-PM summaries
  - Added multi-year / custom-range summary generation
  - Improved printing analytics and year-over-year comparisons
  - Added similarity (same client) detection across months and years
  - Slack notification formatting and weekly automation hooks
  - Caching and error handling improvements
  - Minor bug fixes and more robust date parsing

---

Contact / questions
- Repo owner: Iamhuzaifasabahuddin
- For help with credentials or deployment, create an issue in the repo describing the environment and logs.
