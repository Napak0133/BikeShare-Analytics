# BikeShare Analytics — Website

A Flask website for managing and analysing the Bike Sharing dataset (Washington D.C., 2011–2012).

## How the data flows

```
Add / Edit / Delete on the website
        │
        ▼
data/Bike_Sharing_Analysis_Final.xlsx      ← the ONE data file (Web_Entries sheet, updated in place)
        │
        ▼
data/Bike_Sharing.hyper                    ← rebuilt from that Excel file (17 columns, Extract.Extract)
        │
        ├──────────────────────────────────► Dashboard charts queried from .hyper with SQL
        │
        └── when Tableau Cloud credentials are configured:
            packaged Tableau workbook (.twbx) → overwrite published Cloud workbook → /tableau embed
        │
        ▼
The website charts redraw after a save; the embedded Tableau view loads the published Cloud workbook
```

* No extra data files are created. There is no database of records and no export copies.
* `Raw_Data` (the original 17,379 rows) is never changed and cannot be edited or deleted from the site.
* Only the `Web_Entries` sheet is rewritten. The rest of the workbook (pivot tables, Excel tables,
  formatting, the other sheets) is copied byte-for-byte — the site does **not** save through openpyxl,
  which would damage the pivot tables.
* If you edit the Excel file by hand, the site notices (file timestamp) and rebuilds the .hyper
  before the next chart request. **Update Status → Re-sync from Excel now** does the same on demand.
* Website add/edit/delete and manual re-sync attempt to publish the latest extract to Tableau Cloud
  when the Tableau PAT environment variables below are configured. A Cloud publish failure does
  not undo the saved Excel record; it is reported in the sync log and returned as a warning.
* If you edit the Excel workbook directly, the website rebuilds its local Hyper extract when it next
  detects the change. Use **Update Status → Re-sync from Excel now** to also publish that change to
  Tableau Cloud.

## Setup and running

Python 3.10+ is required. Open a terminal in the project folder (the one containing `app.py`).

**Windows**
```powershell
python -m venv venv
venv\Scripts\activate
pip install -r requirements.txt
Copy-Item .env.local.example .env
# Open .env in a text editor and fill in the Tableau PAT and view URL.
python app.py
```

Open http://127.0.0.1:5000. Keep the terminal running while you use the site. This runs locally
on this computer; it does not make the website public on the internet.

### Local Tableau Cloud setup

The local app saves website records to `Web_Entries` in the Excel workbook, rebuilds
`data/Bike_Sharing.hyper`, and then publishes a packaged workbook to Tableau Cloud. Publish is
attempted after each successful website add/edit/delete and after **Update Status → Re-sync from
Excel now**. If Tableau publishing fails, the Excel save remains in place and the sync log shows an
error; retry with the manual re-sync after resolving the error.

1. Copy `.env.local.example` to `.env` once. Keep `.env` on this computer; it is excluded from Git.
2. Fill in `TABLEAU_PAT_NAME` and `TABLEAU_PAT_SECRET` with the name and secret of the PAT that
   successfully published the workbook. Do not put the secret in a command, chat, screenshot, or
   committed file.
3. Set `TABLEAU_VIEW_URL` to the full dashboard view URL from Tableau Cloud's **Share → Copy Link**.
   The URL must point to a view inside the published workbook, not just the project or workbook list.
4. In the Connected App details, copy its **Client ID** and the generated **Secret ID** and
   **Secret Value** into `TABLEAU_EMBED_CLIENT_ID`, `TABLEAU_EMBED_SECRET_ID`, and
   `TABLEAU_EMBED_SECRET`. Set `TABLEAU_EMBED_USERNAME` to the email/username of a Tableau Cloud user
   who has permission to view the workbook. Do not use the PAT for these fields.
5. The Connected App must be enabled, scoped to the project containing this workbook, and have a
   domain allowlist entry matching the local host used in the browser (`localhost` or `127.0.0.1`).
   If you switch between those hostnames, allow both. Keep the browser configured to allow
   third-party cookies for Tableau Cloud.
6. Start (or restart) Flask with `python app.py` from this project folder. The app reads `.env` at
   startup, so changing it while Flask is running requires a restart.
7. Use the website at `http://127.0.0.1:5000`. `/` is the existing Chart.js dashboard;
   **Tableau Cloud** in the top navigation opens the authenticated embedded view at `/tableau`.
8. Add one test record and check **Update Status** for a `tableau_publish` success entry. Then open
   the Tableau Cloud workbook and confirm the same record appears.

The PAT and Connected App credentials are separate: the PAT publishes the workbook; the Connected
App creates a short-lived JWT for the embedded view. The JWT is sent to the browser, as required by
Tableau embedding, but the signing secret remains on the Flask server. This local app embeds as the
single `TABLEAU_EMBED_USERNAME` identity, so anyone who can access the local app sees the Tableau
view with that account's permissions. The [.env.local.example](.env.local.example) template contains
the local setting names. This setup is for local use on this computer, not for public multi-user
hosting.

**macOS / Linux**
```bash
python3 -m venv venv
source venv/bin/activate
pip install -r requirements.txt
cp .env.local.example .env
# Edit .env and fill in the Tableau PAT and view URL.
python app.py
```

Open http://127.0.0.1:5000. Next time: activate the venv, then `python app.py`.

> **Close `Bike_Sharing_Analysis_Final.xlsx` in Excel while you use the site.** Windows locks an open
> file, so a save would be refused (the site then shows "close it and try again" and changes nothing).

## Project structure

```
app.py            Flask routes + API
datastore.py      Reads Raw_Data / Web_Entries and writes Web_Entries in the Excel file
hyperstore.py     Builds data/Bike_Sharing.hyper and queries it for the Dashboard views
models.py         Forecast (trend x seasonality) and What-if (log-linear demand model), no extra packages
pipeline.py       Excel → .hyper after every change; optional Tableau Cloud publish; sync log
tableau_sync.py   Packages the app .hyper with the Tableau template and publishes the workbook
analytics.py      Statistics (Analysis page) + unit conversions
templates/        Chart.js Dashboard, Tableau embed, Data Management, Add/Edit, Analysis, Update Status
static/            CSS and JS (EN/TH translations in i18n.js)
data/             Bike_Sharing_Analysis_Final.xlsx and Bike_Sharing.hyper
logs/             sync_log.db — history shown on the Update Status page (created automatically)
```

## Pages

* **Dashboard** — five views switched with the buttons on the left (Year / User Type / Day Type filters apply as noted).
  * *Overview* — KPI cards, user-type ratio, day-type averages, monthly trend.
  * *Hour* — average rentals for each hour of the day (weekday vs holiday/weekend, casual vs registered).
  * *Weather* — average rentals by weather, temperature, humidity and wind speed.
  * *Forecast* — monthly actuals plus a 3 / 6 / 12-month forecast with a 95% range (uses all years; the Year and Day Type filters are greyed out).
  * *What-if* — choose hour, day type, weather, temperature, humidity, wind and year to see predicted rentals against
    typical conditions. The model is refitted from the .hyper whenever the data changes; the sliders answer instantly.
* **Data Management** — browse and filter all records. Website records can be edited or deleted.
* **Add / Edit** — values are entered in everyday units (°C, %, wind speed); the site converts them to the
  normalized values stored in Excel (temp×41, atemp×50, hum×100, windspeed×67). Total = casual + registered.
* **Analysis Results** — descriptive statistics, correlation, regression and summaries over Raw_Data + Web_Entries.
  *Download Excel workbook* gives you the live workbook.
* **Update Status** — every save to Excel and every .hyper rebuild.

## Rules the site enforces

* **IDs** continue from the original data (first website record = 17380) and are the same on the site,
  in Excel and in the .hyper.
* **No duplicates** — one record per date + hour across Raw_Data and Web_Entries.
* `yr`, `mnth` and `weekday` are calculated from the date; working day = Monday–Friday and not a holiday.

## Notes

* The `Pivot_*`, `Tableau_Data` and `Descriptive_Result` sheets inside the workbook are not recalculated by the
  site (the pivot tables are built on `tblRawHourly` only). Refresh them in Excel if you use them.
* The Chart.js library is loaded from a CDN, so the browser needs internet access to draw charts.
