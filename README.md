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
| Tableau | Excel → `.hyper` → `.twb` pipeline (`tableau_sync.py`) + embedded dashboard — see section 5 |
| Unit conversion | `analytics.py` — temp×41=°C, atemp×50=°C, hum×100=%, windspeed×67 |

Authentication/roles were explicitly **out of scope** for this build (single-user).

## 2. Setup and running the website (terminal)

Python 3.10+ is required (the project was built with 3.12). Open a terminal **inside the project folder**
(the folder that contains `app.py`).

**Windows (PowerShell or Command Prompt)**
```powershell
cd "path\to\Website_BikeShare Analytics_3"
python -m venv venv
venv\Scripts\activate            # PowerShell may need:  Set-ExecutionPolicy -Scope Process RemoteSigned
pip install -r requirements.txt
python init_db.py                 # only if bikeshare.db does not exist yet
python app.py
```

**macOS / Linux**
```bash
cd "path/to/Website_BikeShare Analytics_3"
python3 -m venv venv
source venv/bin/activate
pip install -r requirements.txt
python init_db.py                 # only if bikeshare.db does not exist yet
python app.py
```

Then open http://127.0.0.1:5000 in a browser. Stop the server with `Ctrl + C`; leave the virtual
environment with `deactivate`. Next time you only need: activate the venv, then `python app.py`.

To start over with a fresh database, delete `bikeshare.db` and re-run `python init_db.py`.

## 3. Project structure

```
bikeshare_app/
├── app.py                 Flask routes + REST API
├── analytics.py           Calculations, unit conversion, Excel export
├── init_db.py             One-time DB setup / Raw_Data import
├── tableau_sync.py        Dashboard pipeline: Excel → .hyper → .twb/.twbx
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
    ├── js/main.js          Shared fetch/formatting helpers
    └── tableau/            template/ (original .twb), generated .hyper / .twb / .twbx
```

## 4. User manual

### Dashboard (`/`)
Shows total rentals, average rentals/hour, casual vs. registered users, and four charts
(monthly trend, hourly pattern, seasonal comparison, weather comparison). The Tableau
Dashboard area is a placeholder until the real workbook/URL is supplied — swap in the
Tableau embed code where the placeholder `<div class="placeholder-box">` is in
`templates/dashboard.html`.

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
temperature, plus hourly/seasonal/weather summary charts — recomputed over
Raw_Data + Web_Entries combined every time this page loads. Two buttons let you
download the latest raw data or generate a refreshed analysis workbook.

### Data Update Status (`/status`)
Shows the most recent synchronization time, records processed, and current status,
plus a full history table. The **Update Analysis Now** button manually triggers
recalculation (useful after bulk changes or to retry a failed sync).

## 5. Dashboard update pipeline (Excel → .hyper → .twb → Dashboard)

After every add / edit / delete (and on the Dashboard's **Refresh data now** button) the site runs
`tableau_sync.py`:

1. **Excel → .hyper** — all records (original + website) are written to
   `exports/Bike_Sharing_Tableau_Data.xlsx`, which is loaded into `static/tableau/Bike_Sharing.hyper`
   (Tableau Hyper API; table `Extract.Extract`, the 17 dataset columns).
2. **.hyper → .twb** — `static/tableau/template/Project_data_Analysis.twb` (the original workbook, never
   modified) is copied to `static/tableau/Project_data_Analysis.twb` with its data source switched from the
   Excel union (Raw_Data + Web_Entries) to the `.hyper`. Worksheets, calculated fields, parameters and the
   dashboard layout are unchanged. The workbook and extract are also zipped as `Project_data_Analysis.twbx`.
3. **Dashboard** — the Dashboard page embeds the published workbook. A browser cannot open a local
   `.twb/.twbx`, so the new data appears there after the `.twbx` is published (Tableau: *Save to Tableau
   Public*, or publish to Tableau Cloud/Server) and its URL is saved with **Update Charts**.

Run it manually: `python tableau_sync.py [--excel file.xlsx]`. Result and errors are logged on the
*Data Update Status* page (event `tableau_sync`).

### Live charts (auto-updating, no Tableau needed)

A browser cannot render a `.twb`/`.twbx` file, and Tableau Public has no API for uploading a new extract,
so the Tableau charts only change after you re-publish the workbook by hand. To get charts that really update
by themselves, the top of the Dashboard page draws four Chart.js charts (monthly, hourly, season, weather) plus
KPI cards straight from `/api/dashboard` (Raw_Data + Web_Entries) and re-checks every 5 seconds
(`LIVE_REFRESH_MS` in `templates/dashboard.html`). Add / edit / delete a record and every open Dashboard tab
updates within a few seconds. The Excel → .hyper → .twb pipeline keeps running for the Tableau workbook.
Runs of the pipeline are serialized with a lock and the Excel file is written atomically, which fixes the
occasional `Bad CRC-32` sync error.

## 6. Notes

- **DB → Excel**: `.xlsx` files are generated on demand/after every change (no live ODBC link).
- **Holiday / Working Day**: working day = Monday–Friday and not a public holiday. The form fills it in
  automatically and the server rejects contradictory combinations.
