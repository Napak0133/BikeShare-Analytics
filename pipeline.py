"""
pipeline.py
-----------
Keeps the three layers in step.  The flow after EVERY add / edit / delete is:

    1. the record is written into  data/Bike_Sharing_Analysis_Final.xlsx  (Web_Entries sheet)
    2. the .hyper  data/Bike_Sharing.hyper  is rebuilt from that Excel file
    3. the website's charts query the .hyper   (see hyperstore.dashboard_data)
    4. open Dashboard pages notice the new `version` and redraw

If somebody edits the Excel file by hand, ensure_fresh() notices (file timestamp changed) and
rebuilds the .hyper before the next chart request, so Excel and the charts never drift apart.

Every step is recorded in logs/sync_log.db (shown on the Data Update Status page).
"""

import datetime
import os
import sqlite3

import datastore
import hyperstore

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
LOG_DB = os.path.join(BASE_DIR, "logs", "sync_log.db")
lock = datastore.write_lock                   # the same lock the Excel writer uses

state = {"sig": None, "version": 0, "built_at": None, "rows": 0}


# ------------------------------------------------------------------ sync log

def _log_conn():
    os.makedirs(os.path.dirname(LOG_DB), exist_ok=True)
    conn = sqlite3.connect(LOG_DB)
    conn.row_factory = sqlite3.Row
    conn.execute("""CREATE TABLE IF NOT EXISTS sync_log (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        event_time TEXT DEFAULT (datetime('now', 'localtime')),
        event_type TEXT NOT NULL, status TEXT NOT NULL,
        records_synced INTEGER DEFAULT 0, message TEXT)""")
    return conn


def log(event_type, status, records_synced=0, message=""):
    conn = _log_conn()
    conn.execute("INSERT INTO sync_log (event_type, status, records_synced, message) VALUES (?,?,?,?)",
                 (event_type, status, records_synced, message))
    conn.commit()
    conn.close()


def read_log(limit=100):
    conn = _log_conn()
    rows = [dict(r) for r in conn.execute("SELECT * FROM sync_log ORDER BY id DESC LIMIT ?", (limit,))]
    conn.close()
    return rows


# ------------------------------------------------------------------ Excel -> .hyper

def _excel_sig():
    st = os.stat(datastore.EXCEL_PATH)
    return (st.st_mtime_ns, st.st_size)


def rebuild_hyper():
    """Excel -> .hyper. Returns the number of rows now in the .hyper."""
    with lock:
        sig = _excel_sig()
        n = hyperstore.rebuild(datastore.all_rows())
        state.update(sig=sig, rows=n, version=state["version"] + 1,
                     built_at=datetime.datetime.now().isoformat(timespec="seconds"))
        return n


def ensure_fresh(reason="Excel file changed"):
    """Cheap check (one file stat). Rebuilds the .hyper if it is missing or Excel changed since."""
    with lock:
        if hyperstore.exists() and state["sig"] == _excel_sig():
            return False
        try:
            n = rebuild_hyper()
            log("excel_sync", "success", n, f"{reason}: Bike_Sharing.hyper rebuilt from the Excel file.")
            return True
        except Exception as e:
            log("excel_sync", "error", 0, f"{reason}: {e}")
            raise


def commit(action, instant, mutate):
    """Run `mutate()` (it writes the Excel file), then rebuild the .hyper.
    Returns (result, warning). `warning` is set when Excel was saved but the .hyper refresh failed."""
    with lock:
        try:
            result = mutate()
        except Exception as e:
            log(action, "error", 0, f"id={instant}: {e}")
            raise
        ident = result if action == "add" else instant      # a new record's id is only known now
        try:
            n = rebuild_hyper()
            log(action, "success", 1,
                f"id={ident}: saved to Web_Entries in Bike_Sharing_Analysis_Final.xlsx; "
                f"Bike_Sharing.hyper rebuilt ({n:,} rows); charts updated.")
            return result, None
        except Exception as e:
            log(action, "error", 1, f"id={ident}: saved to Excel, but the .hyper refresh failed: {e}")
            return result, f"Saved to Excel, but the charts could not be refreshed: {e}"
