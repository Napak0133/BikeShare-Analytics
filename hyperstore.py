"""
hyperstore.py
-------------
data/Bike_Sharing.hyper  - the Tableau extract built from the Excel workbook.

  * rebuild(rows)        rewrites the .hyper from the rows read out of Excel (all 17 columns, table
                         Extract.Extract - the same names Tableau itself uses, so the file still opens
                         in Tableau Desktop)
  * dashboard_data()     the website's Dashboard charts are queried from this .hyper with SQL

One Hyper process is started on first use and kept for the life of the website, so queries are fast.
"""

import atexit
import datetime
import os
import threading

from tableauhyperapi import (
    Connection, CreateMode, Date, HyperProcess, Inserter, SqlType, TableDefinition,
    TableName, Telemetry,
)

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
HYPER_PATH = os.path.join(BASE_DIR, "data", "Bike_Sharing.hyper")
TABLE = TableName("Extract", "Extract")

COLUMNS = [
    ("instant", SqlType.big_int()), ("dteday", SqlType.date()), ("season", SqlType.big_int()),
    ("yr", SqlType.big_int()), ("mnth", SqlType.big_int()), ("hr", SqlType.big_int()),
    ("holiday", SqlType.big_int()), ("weekday", SqlType.big_int()), ("workingday", SqlType.big_int()),
    ("weathersit", SqlType.big_int()), ("temp", SqlType.double()), ("atemp", SqlType.double()),
    ("hum", SqlType.double()), ("windspeed", SqlType.double()), ("casual", SqlType.big_int()),
    ("registered", SqlType.big_int()), ("cnt", SqlType.big_int()),
]

_lock = threading.RLock()
_process = None


def _hyper():
    global _process
    if _process is None:
        _process = HyperProcess(Telemetry.DO_NOT_SEND_USAGE_DATA_TO_TABLEAU, parameters={"log_config": ""})
        atexit.register(close)
    return _process


def close():
    global _process
    with _lock:
        if _process is not None:
            _process.close()
            _process = None


def exists():
    return os.path.isfile(HYPER_PATH)


def rebuild(rows):
    """Write all rows (Raw_Data + Web_Entries from Excel) into a fresh .hyper. Returns the row count."""
    table = TableDefinition(TABLE, [TableDefinition.Column(n, t) for n, t in COLUMNS])
    names = [n for n, _ in COLUMNS]
    data = []
    for r in rows:
        d = datetime.date.fromisoformat(r["dteday"])
        data.append([Date(d.year, d.month, d.day) if n == "dteday" else r[n] for n in names])
    with _lock:
        os.makedirs(os.path.dirname(HYPER_PATH), exist_ok=True)
        with Connection(_hyper().endpoint, HYPER_PATH, CreateMode.CREATE_AND_REPLACE) as conn:
            conn.catalog.create_schema_if_not_exists(TABLE.schema_name)
            conn.catalog.create_table(table)
            with Inserter(conn, table) as ins:
                ins.add_rows(data)
                ins.execute()
            return int(conn.execute_scalar_query(f"SELECT COUNT(*) FROM {TABLE}"))


def _grouped(conn, field):
    rows = conn.execute_list_query(f"SELECT {field}, SUM(cnt) FROM {TABLE} GROUP BY {field} ORDER BY {field}")
    return {int(k): int(v) for k, v in rows}


def dashboard_data():
    """Everything the Dashboard needs, aggregated by SQL inside the .hyper."""
    with _lock:
        with Connection(_hyper().endpoint, HYPER_PATH) as conn:
            n, total, casual, registered = conn.execute_list_query(
                f"SELECT COUNT(*), COALESCE(SUM(cnt),0), COALESCE(SUM(casual),0), COALESCE(SUM(registered),0) "
                f"FROM {TABLE}")[0]
            monthly = []
            for yr, mnth, tot, cas, reg, cnt in conn.execute_list_query(
                    f"SELECT yr, mnth, SUM(cnt), SUM(casual), SUM(registered), COUNT(*) FROM {TABLE} "
                    f"GROUP BY yr, mnth ORDER BY yr, mnth"):
                monthly.append({"year": 2011 + int(yr), "month": int(mnth), "total": int(tot),
                                "casual": int(cas), "registered": int(reg), "records": int(cnt),
                                "avg_per_record": round(int(tot) / int(cnt), 3)})
            n, total, casual, registered = int(n), int(total), int(casual), int(registered)
            return {
                "record_count": n, "total_rentals": total,
                "total_casual": casual, "total_registered": registered,
                "avg_per_hour": round(total / n, 2) if n else 0,
                "monthly": monthly,
                "hourly": _grouped(conn, "hr"),
                "season": _grouped(conn, "season"),
                "weather": _grouped(conn, "weathersit"),
            }


_USER_COLS = {"cnt": "cnt", "casual": "casual", "registered": "registered"}


def overview_data(year=0, user="cnt", daytype="all"):
    """Everything the Overview dashboard needs, for the chosen Year / User Type / Day Type.

    year     0 = all years, otherwise a calendar year (2011, 2012 ...)
    user     "cnt" (all users), "casual" or "registered"  -> the column the KPIs, trend and day-type chart add up
    daytype  "all", "weekday" (working day) or "off" (holiday / weekend) -> applies to the Day Type chart only,
             exactly like the Tableau dashboard where that filter sits on the Day Type AVG Rental sheet.
    """
    col = _USER_COLS.get(user, "cnt")
    year = int(year or 0)
    yr = "EXTRACT(YEAR FROM dteday)"
    where = f"WHERE {yr} = {year}" if year else ""
    with _lock:
        with Connection(_hyper().endpoint, HYPER_PATH) as conn:
            years = [int(r[0]) for r in conn.execute_list_query(
                f"SELECT DISTINCT {yr} FROM {TABLE} ORDER BY 1")]
            n, sel, casual, registered, total, peak = conn.execute_list_query(
                f"SELECT COUNT(*), COALESCE(SUM({col}),0), COALESCE(SUM(casual),0), "
                f"COALESCE(SUM(registered),0), COALESCE(SUM(cnt),0), COALESCE(MAX(cnt),0) "
                f"FROM {TABLE} {where}")[0]
            prior = None
            if year:
                p = conn.execute_list_query(
                    f"SELECT SUM({col}) FROM {TABLE} WHERE {yr} = {year - 1}")[0][0]
                prior = int(p) if p else None
            monthly = [{"year": int(y), "month": int(m), "value": int(v)} for y, m, v in conn.execute_list_query(
                f"SELECT {yr}, EXTRACT(MONTH FROM dteday), SUM({col}) FROM {TABLE} {where} "
                f"GROUP BY {yr}, EXTRACT(MONTH FROM dteday) ORDER BY 1, 2")]
            day_rows = conn.execute_list_query(
                f"SELECT wd, AVG(daily) FROM (SELECT dteday, MAX(workingday) AS wd, SUM({col}) AS daily "
                f"FROM {TABLE} {where} GROUP BY dteday) AS d GROUP BY wd")
    day = {"weekday": None, "off": None}
    for wd, avg in day_rows:
        day["weekday" if int(wd) == 1 else "off"] = round(float(avg), 1)
    if daytype in day:
        day = {k: (v if k == daytype else None) for k, v in day.items()}
    n, sel, casual, registered, total, peak = (int(x) for x in (n, sel, casual, registered, total, peak))
    return {
        "years": years, "year": year, "user": col, "daytype": daytype,
        "total": sel, "prior": prior,
        "yoy": ((sel - prior) / prior) if prior else None,
        "record_count": n, "peak_per_hour": peak,
        "casual": casual, "registered": registered,
        "casual_pct": (casual / total) if total else 0,
        "registered_pct": (registered / total) if total else 0,
        "monthly": monthly, "daytype_avg": day,
    }


# ====================================================================== extra views (Hour / Weather / Forecast / What-if)

def _where(year=0, daytype="all"):
    """WHERE clause shared by the Hour and Weather views. `yr` is the year minus 2011 (same as the dataset)."""
    parts = []
    if year:
        parts.append(f"yr = {int(year) - 2011}")
    if daytype == "weekday":
        parts.append("workingday = 1")
    elif daytype == "off":
        parts.append("workingday = 0")
    return ("WHERE " + " AND ".join(parts)) if parts else ""


def hour_data(year=0, user="cnt", daytype="all"):
    """Average rentals per hour-of-day, split by day type and by user type."""
    col = _USER_COLS.get(user, "cnt")
    where = _where(year, daytype)
    with _lock:
        with Connection(_hyper().endpoint, HYPER_PATH) as conn:
            rows = conn.execute_list_query(
                f"SELECT hr, workingday, AVG({col}), AVG(casual), AVG(registered), SUM({col}), COUNT(*) "
                f"FROM {TABLE} {where} GROUP BY hr, workingday ORDER BY hr, workingday")
    by_day = {"weekday": [None] * 24, "off": [None] * 24}
    tot = [0] * 24                       # rentals per hour (all day types) for the selected user type
    cas, reg, cnt = [0.0] * 24, [0.0] * 24, [0] * 24
    for hr, wd, avg, ac, ar, s, n in rows:
        hr = int(hr)
        by_day["weekday" if int(wd) == 1 else "off"][hr] = round(float(avg), 1)
        tot[hr] += int(s)
        cas[hr] += float(ac) * int(n)
        reg[hr] += float(ar) * int(n)
        cnt[hr] += int(n)
    avg_all = [round(tot[h] / cnt[h], 1) if cnt[h] else None for h in range(24)]
    return {
        "year": int(year or 0), "user": col, "daytype": daytype,
        "by_day": by_day, "avg_all": avg_all,
        "casual": [round(cas[h] / cnt[h], 1) if cnt[h] else None for h in range(24)],
        "registered": [round(reg[h] / cnt[h], 1) if cnt[h] else None for h in range(24)],
        "total": sum(tot), "records": sum(cnt),
    }


def _bin(rows, idx, step, scale, nmax):
    """rows = (value, avg, n) grouped by the raw stored value; re-bin into `step`-wide bins of value*scale."""
    acc = {}
    for r in rows:
        v, avg, n = r[idx], float(r[idx + 1]), int(r[idx + 2])
        b = min(int(v * scale // step) * step, nmax)
        s = acc.setdefault(b, [0.0, 0])
        s[0] += avg * n
        s[1] += n
    return [{"from": b, "to": b + step, "avg": round(s[0] / s[1], 1), "n": s[1]} for b, s in sorted(acc.items()) if s[1] >= 20]


def weather_data(year=0, user="cnt", daytype="all"):
    """Average rentals per hour by weather situation and by temperature / humidity / wind speed bins."""
    col = _USER_COLS.get(user, "cnt")
    where = _where(year, daytype)
    with _lock:
        with Connection(_hyper().endpoint, HYPER_PATH) as conn:
            sit = conn.execute_list_query(
                f"SELECT weathersit, AVG({col}), COUNT(*) FROM {TABLE} {where} GROUP BY weathersit ORDER BY weathersit")
            temp = conn.execute_list_query(
                f"SELECT \"temp\", AVG({col}), COUNT(*) FROM {TABLE} {where} GROUP BY \"temp\"")
            hum = conn.execute_list_query(
                f"SELECT \"hum\", AVG({col}), COUNT(*) FROM {TABLE} {where} GROUP BY \"hum\"")
            wind = conn.execute_list_query(
                f"SELECT \"windspeed\", AVG({col}), COUNT(*) FROM {TABLE} {where} GROUP BY \"windspeed\"")
    return {
        "year": int(year or 0), "user": col, "daytype": daytype,
        "weather": [{"id": int(s), "avg": round(float(a), 1), "n": int(n)} for s, a, n in sit],
        "temp": _bin(temp, 0, 5, 41, 40),            # °C  (stored as temp = °C / 41)
        "hum": _bin(hum, 0, 10, 100, 90),            # %
        "wind": _bin(wind, 0, 10, 67, 50),           # wind speed (stored as value / 67)
    }


def monthly_series(user="cnt"):
    """Monthly totals over ALL years (the Forecast view ignores the Year filter) + records per month."""
    col = _USER_COLS.get(user, "cnt")
    with _lock:
        with Connection(_hyper().endpoint, HYPER_PATH) as conn:
            rows = conn.execute_list_query(
                f"SELECT yr, mnth, SUM({col}), COUNT(*) FROM {TABLE} GROUP BY yr, mnth ORDER BY yr, mnth")
    return [{"year": 2011 + int(y), "month": int(m), "value": int(v), "records": int(n)} for y, m, v, n in rows]


def model_rows():
    """The columns the What-if model is trained on (one tuple per hourly record)."""
    with _lock:
        with Connection(_hyper().endpoint, HYPER_PATH) as conn:
            return conn.execute_list_query(
                f"SELECT hr, workingday, weathersit, \"temp\", \"hum\", \"windspeed\", yr, casual, registered, cnt FROM {TABLE}")


def years():
    """Calendar years present in the .hyper (so every view can keep the Year drop-down up to date)."""
    with _lock:
        with Connection(_hyper().endpoint, HYPER_PATH) as conn:
            return [2011 + int(r[0]) for r in conn.execute_list_query(f"SELECT DISTINCT yr FROM {TABLE} ORDER BY 1")]
