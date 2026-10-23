"""
init_db.py
----------
Creates the SQLite database (bikeshare.db) for the BikeShare Analytics website
and imports the original Raw_Data (17,379 records) from the project's Excel
workbook (data/raw_data.csv, exported from Bike_Sharing_Analysis_Final.xlsx).

Run this once before starting the Flask app:
    python init_db.py

Tables created:
  raw_data     -> original, read-only historical dataset (protected)
  web_entries  -> records added/edited/deleted through the website
  sync_log     -> history of Excel/Tableau synchronization events (Data Update Status page)

Both raw_data and web_entries share the same column structure (A:Q from the
original workbook) so they can be combined for analysis.
"""

import sqlite3
import csv
import os

DB_PATH = os.path.join(os.path.dirname(__file__), "bikeshare.db")
CSV_PATH = os.path.join(os.path.dirname(__file__), "data", "raw_data.csv")

SCHEMA = """
CREATE TABLE IF NOT EXISTS raw_data (
    instant      INTEGER PRIMARY KEY,
    dteday       TEXT NOT NULL,
    season       INTEGER NOT NULL,
    yr           INTEGER NOT NULL,
    mnth         INTEGER NOT NULL,
    hr           INTEGER NOT NULL,
    holiday      INTEGER NOT NULL,
    weekday      INTEGER NOT NULL,
    workingday   INTEGER NOT NULL,
    weathersit   INTEGER NOT NULL,
    temp         REAL NOT NULL,
    atemp        REAL NOT NULL,
    hum          REAL NOT NULL,
    windspeed    REAL NOT NULL,
    casual       INTEGER NOT NULL,
    registered   INTEGER NOT NULL,
    cnt          INTEGER NOT NULL
);

CREATE TABLE IF NOT EXISTS web_entries (
    instant      INTEGER PRIMARY KEY AUTOINCREMENT,
    dteday       TEXT NOT NULL,
    season       INTEGER NOT NULL,
    yr           INTEGER NOT NULL,
    mnth         INTEGER NOT NULL,
    hr           INTEGER NOT NULL,
    holiday      INTEGER NOT NULL,
    weekday      INTEGER NOT NULL,
    workingday   INTEGER NOT NULL,
    weathersit   INTEGER NOT NULL,
    temp         REAL NOT NULL,
    atemp        REAL NOT NULL,
    hum          REAL NOT NULL,
    windspeed    REAL NOT NULL,
    casual       INTEGER NOT NULL,
    registered   INTEGER NOT NULL,
    cnt          INTEGER NOT NULL,
    created_at   TEXT DEFAULT (datetime('now')),
    updated_at   TEXT DEFAULT (datetime('now'))
);

CREATE TABLE IF NOT EXISTS sync_log (
    id               INTEGER PRIMARY KEY AUTOINCREMENT,
    event_time       TEXT DEFAULT (datetime('now')),
    event_type       TEXT NOT NULL,      -- 'auto_sync' | 'manual_sync' | 'export' | 'analysis_update'
    status           TEXT NOT NULL,      -- 'success' | 'error'
    records_synced   INTEGER DEFAULT 0,
    message          TEXT
);
"""


def main():
    if os.path.exists(DB_PATH):
        print(f"Database already exists at {DB_PATH}. Delete it first if you want to re-import.")
        return

    conn = sqlite3.connect(DB_PATH)
    cur = conn.cursor()
    cur.executescript(SCHEMA)

    # Import Raw_Data from CSV (original 17,379-row dataset)
    with open(CSV_PATH, newline="", encoding="utf-8") as f:
        reader = csv.DictReader(f)
        rows = []
        for r in reader:
            dteday = r["dteday"].split(" ")[0]  # keep date part only
            rows.append((
                int(r["instant"]), dteday, int(r["season"]), int(r["yr"]), int(r["mnth"]),
                int(r["hr"]), int(r["holiday"]), int(r["weekday"]), int(r["workingday"]),
                int(r["weathersit"]), float(r["temp"]), float(r["atemp"]), float(r["hum"]),
                float(r["windspeed"]), int(r["casual"]), int(r["registered"]), int(r["cnt"]),
            ))

    cur.executemany(
        """INSERT INTO raw_data
           (instant, dteday, season, yr, mnth, hr, holiday, weekday, workingday,
            weathersit, temp, atemp, hum, windspeed, casual, registered, cnt)
           VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",
        rows,
    )

    cur.execute(
        "INSERT INTO sync_log (event_type, status, records_synced, message) VALUES (?,?,?,?)",
        ("auto_sync", "success", len(rows), "Initial import of Raw_Data from Bike_Sharing_Analysis_Final.xlsx"),
    )

    conn.commit()
    print(f"Imported {len(rows)} rows into raw_data.")
    print(f"Database created at {DB_PATH}")
    conn.close()


if __name__ == "__main__":
    main()
