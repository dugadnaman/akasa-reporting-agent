# ✈️ Akasa Air — Campaign Reports Cleaner & Consolidator

A specialized analytics pipeline and interactive web application designed for Akasa Air's digital marketing and retention teams. It automates the extraction, cleaning, normalization, and consolidation of multi-channel campaign reports (Email, WhatsApp, App Push, Web Push, and Omnichannel Flows) from **MoEngage** into a single executive attribution report (`final_campaign_report.xlsx`).

---

## 🎯 Objectives & Business Value

1. **Multi-Channel Attribution**: Unifies disjointed campaign exports from Email, WhatsApp, App Push (Android + iOS), Web Push, and MoEngage Journeys/Flows into a single unified report.
2. **Flight Booking Conversion Focus**: Normalizes Primary Goal 1 (`FB_Successful` — Flight Booking Successful) tracking metrics across channels, isolating direct and view-through ticket sales.
3. **Automated Data Hygiene**: Eliminates manual spreadsheet wrangling by programmatically resolving tag permutations, standardizing column casing, and pruning zero-activity draft campaigns.

---

## 🛠️ Issues in Original Code & Applied Fixes

| Issue Identified | Impact in Original Code | Applied Fix |
| :--- | :--- | :--- |
| **Tag Order Permutations** | String tags like `"['AMD', 'OneWayAbandoned_V3']"` and `"['OneWayAbandoned_V3', 'AMD']"` appeared as two duplicate campaigns. | Added `normalize_campaign_tag()` which parses, strips, and alphabetically sorts list tags into a single canonical campaign key. |
| **Click Column Casing Desync** | Email/WhatsApp used `'Total clicks'` while Push used `'Total Clicks'`, causing two separate click columns in output. | Standardized all channels onto unified `'Total Clicks'` column. |
| **Dead Filter / Shift Bug** | `df_cleaned` was computed to drop all-zero rows, but `df.loc[]` was used on un-filtered DataFrame. In `overall.xls`, positional slicing (`iloc[:, start_idx:]`) caused an empty DataFrame slice. | Rewrote filter to select non-zero rows explicitly on metric columns without positional index shift bugs. |
| **Regex Whitespace Error** | Dropping secondary goals used `goal 2 \| goal 3` with spaces, missing `Goal 3` columns without whitespace. | Standardized regex to strip all secondary/tertiary goals cleanly. |
| **Push Metric Aggregation** | Android and iOS metrics required manual calculation. | Automated summation of Android + iOS sent, impressions, clicks, conversions, and revenue into unified App Push metrics. |

---

## 🚀 Running the Web Tool

### Option 1: Run Locally (Fastest)

1. Ensure dependencies are installed:
   ```bash
   pip install -r requirements.txt
   ```

2. Start the Streamlit application:
   ```bash
   streamlit run app.py
   ```

3. Open your browser at `http://localhost:8501`.

### Option 2: Run with Docker / Container

1. Build and run with Docker Compose:
   ```bash
   docker compose up -d --build
   ```

2. Or build with standard Docker:
   ```bash
   docker build -t akasa-reports-cleaner .
   docker run -p 8501:8501 akasa-reports-cleaner
   ```

3. Access the web interface at `http://localhost:8501`.

---

## 📊 How to Use the UI Web Tool

1. **Pre-loaded Sample Data**: Select this mode to instantly analyze the bundled reports in `AkasaReportsCleaner/csvFiles/`.
2. **Upload Custom Reports**: Upload your weekly or monthly export files for:
   - Email Report (`.xls` / `.xlsx`)
   - WhatsApp Report (`.xls` / `.xlsx`)
   - Push Notifications Report (`.xls` / `.xlsx`)
   - Overall Flows Report (`.xls` / `.xlsx`)
3. **KPI Metrics & Visualizations**: Review aggregate volumes, converted user counts, direct revenue, and user journey flow revenue.
4. **Interactive Explorer**: Search and filter by campaign tag or channel.
5. **Download Final Report**: Click **"Download Excel (.xlsx)"** to export the styled `final_campaign_report.xlsx`.

---

## 📁 Repository Structure

```
├── app.py                      # Interactive Streamlit UI web application
├── pipeline.py                 # Core normalization and data aggregation engine
├── requirements.txt            # Python dependencies
├── Dockerfile                  # Container definition
├── docker-compose.yml          # Docker Compose orchestration
├── README.md                   # Documentation & operations guide
└── AkasaReportsCleaner/
    ├── csvFiles/               # Sample input MoEngage reports
    │   ├── email.xls
    │   ├── whatsapp.xls
    │   ├── push.xls
    │   └── overall.xls
    └── ...                     # Legacy prototype notebooks
```
