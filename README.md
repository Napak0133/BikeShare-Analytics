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
        ▼
Dashboard charts                           ← queried from the .hyper with SQL
        │
        ▼
Open Dashboard pages redraw (about 1 second)
```

* No extra data files are created. There is no database of records and no export copies.
* `Raw_Data` (the original 17,379 rows) is never changed and cannot be edited or deleted from the site.
* Only the `Web_Entries` sheet is rewritten. The rest of the workbook (pivot tables, Excel tables,
  formatting, the other sheets) is copied byte-for-byte — the site does **not** save through openpyxl,
  which would damage the pivot tables.
* If you edit the Excel file by hand, the site notices (file timestamp) and rebuilds the .hyper
  before the next chart request. **Update Status → Re-sync from Excel now** does the same on demand.

## Setup and running

Python 3.10+ is required. Open a terminal in the project folder (the one containing `app.py`).

**Windows**
```powershell
python -m venv venv
venv\Scripts\activate
pip install -r requirements.txt
python app.py
```

**macOS / Linux**
```bash
python3 -m venv venv
source venv/bin/activate
pip install -r requirements.txt
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
pipeline.py       Excel → .hyper after every change, change detection, sync log
analytics.py      Statistics (Analysis page) + unit conversions
templates/        Dashboard, Data Management, Add/Edit, Analysis Results, Update Status
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
