# BikeShare Analytics — Website

A Flask website for managing and analyzing the Bike Sharing dataset (Washington D.C.,
2011–2012), built against `DetailToBuildWeb.pdf` and `Bike_Sharing_Analysis_Final.xlsx`.

## 1. What's included

| Requirement (spec) | Implementation |
|---|---|
| Database | SQLite (`bikeshare.db`), created by `init_db.py` |
| Raw_Data import | All 17,379 original records imported on first run, protected from edit/delete |
| Web_Entries | `web_entries` table — add / edit / delete via the website |
| Excel integration | `analytics.py` reads/writes Excel with `openpyxl`; exports never touch the original workbook |
| 5 pages | Dashboard, Data Management, Add/Edit Data, Analysis Results, Data Update Status |
| Auto update | Every add/edit/delete recalculates descriptive/correlation/regression stats and logs a sync event |
| Manual update | "Update Analysis Now" button on the Data Update Status page |
| Export buttons | `Bike_Sharing_Export.xlsx` (raw DB dump) and `Bike_Sharing_Analysis_Updated.xlsx` (template + recalculated results) |
| Tableau | Embedded Tableau Public dashboard; a saved URL box and **Update Charts** button on the Dashboard page; CSV / Excel export buttons on the Data Management page |
| Unit conversion | `analytics.py` — temp×41=°C, atemp×50=°C, hum×100=%, windspeed×67 |

Authentication/roles were explicitly **out of scope** for this build (single-user).

## 2. Setup

```bash
cd bikeshare_app
python -m venv venv
source venv/bin/activate        # venv\Scripts\activate on Windows
pip install -r requirements.txt
python init_db.py               # creates bikeshare.db and imports Raw_Data (run once)
python app.py                   # starts the dev server on http://127.0.0.1:5000
```

Open http://127.0.0.1:5000 in a browser.

To start over with a fresh database, delete `bikeshare.db` and re-run `python init_db.py`.

## 3. Project structure

```
bikeshare_app/
├── app.py                 Flask routes + REST API
├── analytics.py           Calculations, unit conversion, Excel export
├── init_db.py             One-time DB setup / Raw_Data import
├── requirements.txt
├── bikeshare.db            (created after init_db.py)
├── data/
│   ├── Bike_Sharing_Analysis_Final.xlsx   original workbook (template; never modified)
│   └── raw_data.csv                       exported Raw_Data used by init_db.py
├── exports/                (created at runtime by the export buttons)
├── templates/              Jinja2 page templates (dashboard, data_management, add_edit,
│                            analysis_results, update_status, base)
└── static/
    ├── css/style.css       Light theme, blue/green/navy (spec section 10)
    └── js/main.js          Shared fetch/formatting helpers
```

## 4. User manual

### Dashboard (`/`)
Shows the embedded Tableau Public dashboard. To update its charts after the data changes:
1. On the **Data Management** page, click **Export Data for Tableau (CSV)** or **(Excel)**. The file contains every record
   currently in the database (original data plus website additions, edits and deletions),
   with the same 17 columns as the original dataset.
2. Open your Tableau workbook, update its data source with that file, update the charts and
   republish to Tableau Public.
3. Copy the updated Dashboard URL and paste it into the URL box on the Dashboard page.
4. Click **Update Charts**. The URL is saved in the database (so it survives reloads and
   restarts) and the charts are loaded from it.

### Data Management (`/data`)
Browse all 17,379+ records with pagination, search by date, and filter by data source
(Original vs. Website), season, or weather. Records from `Raw_Data` show **Protected**
in the Actions column (no edit/delete); records added via the website show **Edit** and
**Delete** buttons. Deleting asks for confirmation first.

### Add / Edit Data (`/data/add`)
One shared form for both actions (`/data/add` to create, `/data/add?id=<id>` to edit).
Enter temperature/humidity/wind in everyday units — the backend converts them to the
normalized Excel format automatically. Total Rentals is calculated live from
Casual + Registered and is read-only.

### Analysis Results (`/analysis`)
Descriptive statistics, correlation table, and regression results for rentals vs.
temperature, plus hourly/seasonal/weather summary tables — recomputed over
Raw_Data + Web_Entries combined every time this page loads. The summaries by hour,
season and weather are shown as tables. Two buttons let you download the latest raw data or
generate a refreshed analysis workbook.

### Data Update Status (`/status`)
Shows the most recent synchronization time, records processed, and current status,
plus a full history table. The **Update Analysis Now** button manually triggers
recalculation (useful after bulk changes or to retry a failed sync).

## 5. Notes on "live" Excel/Tableau sync

The spec's "database connected to Excel" / "Tableau refresh" steps are implemented as:
- **DB → Excel**: on-demand generation of `.xlsx` files via the export buttons (closest
  practical equivalent to a live ODBC/Power Query link, which depends on the
  institution's actual Excel/Tableau environment — see spec section 6.4/9).
- **Database → Tableau Public**: Tableau Public cannot connect to this website's database, so
  the Data Management page exports the current database to CSV / Excel; you update and republish
  the workbook in Tableau, then paste the new URL and click **Update Charts**.
