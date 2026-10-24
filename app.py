"""
app.py — BikeShare Analytics website (Flask)

Pages (per spec section 5):
  5.1 /                     Dashboard
  5.2 /data                 Data Management
  5.3 /data/add             Add and Edit Data (shared form)
  5.4 /analysis             Analysis Results
  5.5 /status                Data Update Status

API endpoints used by the pages' JS (static/js/main.js):
  GET    /api/records               paginated, searchable, filterable table rows
  GET    /api/records/<id>          single record (for Edit form pre-fill)
  POST   /api/records               create a web_entries record
  PUT    /api/records/<id>          update a web_entries record
  DELETE /api/records/<id>          delete a web_entries record (web_entries only)
  GET    /api/dashboard             dashboard summary + chart data
  GET    /api/analysis              analysis results (descriptive/correlation/regression/pivots)
  POST   /api/analysis/update       manually trigger recalculation (5.5 "Update Analysis" button)
  GET    /api/sync-log              history for the Data Update Status page
  GET    /export/latest             downloads Bike_Sharing_Export.xlsx (7.1)
  GET    /export/updated-workbook   downloads Bike_Sharing_Analysis_Updated.xlsx (7.2)
  GET    /export/tableau-csv        downloads Bike_Sharing_Tableau_Data.csv (Tableau update file)
  GET    /export/tableau-xlsx       downloads Bike_Sharing_Tableau_Data.xlsx (Tableau update file)
  GET    /api/tableau-url           saved Tableau Public dashboard URL
  POST   /api/tableau-url           save the Tableau Public dashboard URL
  POST   /api/tableau/refresh       run Excel -> .hyper -> .twb/.twbx now (Dashboard "Refresh" button)
  GET    /api/tableau/status        state of the last pipeline run + generated files
  GET    /export/tableau-workbook   downloads Project_data_Analysis.twbx (twb + hyper) for publishing
  GET    /export/tableau-hyper      downloads Bike_Sharing.hyper
"""

from flask import Flask, render_template, request, jsonify, send_file
import os
import threading
import sqlite3
import datetime
from urllib.parse import urlparse
import analytics

app = Flask(__name__)


def get_conn():
    conn = sqlite3.connect(analytics.DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn


def init_settings():
    """Small key/value table that keeps the saved Tableau Public URL."""
    conn = get_conn()
    conn.execute("CREATE TABLE IF NOT EXISTS settings (key TEXT PRIMARY KEY, value TEXT NOT NULL)")
    conn.commit()
    conn.close()


init_settings()

DEFAULT_TABLEAU_URL = "https://public.tableau.com/views/Project_data_Analysis/Overview"


def log_sync(event_type, status, records_synced=0, message=""):
    conn = get_conn()
    conn.execute(
        "INSERT INTO sync_log (event_type, status, records_synced, message) VALUES (?,?,?,?)",
        (event_type, status, records_synced, message),
    )
    conn.commit()
    conn.close()


def auto_sync_after_change(action, instant=None):
    """Spec 6.2/6.3/8.5: after any add/edit/delete, the system must
    automatically update Excel analysis + log the sync event."""
    try:
        analytics.run_full_analysis()
        log_sync("auto_sync", "success", records_synced=1,
                  message=f"Record {action} (id={instant}); analysis recalculated.")
    except Exception as e:
        log_sync("auto_sync", "error", records_synced=0, message=str(e))
    refresh_tableau_files(f"Record {action} (id={instant})")


_pipeline_lock = threading.Lock()   # only one Excel -> .hyper -> .twb run at a time


def refresh_tableau_files(reason):
    """Dashboard pipeline: Excel file -> .hyper -> .twb/.twbx (see tableau_sync.py).
    Never raises: a failure is logged on the Data Update Status page and the save itself still succeeds.
    Runs are serialized: two overlapping runs used to overwrite the Excel file while the other was
    reading it ("Bad CRC-32" errors in the sync log)."""
    try:
        import tableau_sync                       # needs the 'tableauhyperapi' package
        with _pipeline_lock:
            excel_path, _ = analytics.export_tableau_xlsx()      # the Excel file
            result = tableau_sync.run_pipeline(excel_path)       # Excel -> .hyper -> .twb -> .twbx
        log_sync("tableau_sync", "success", records_synced=result["rows"],
                  message=f"{reason}: Excel -> Bike_Sharing.hyper -> Project_data_Analysis.twb/.twbx updated.")
        return result
    except Exception as e:
        log_sync("tableau_sync", "error", message=f"{reason}: {e}")
        return {"error": str(e)}


# ---------------- Page routes ----------------

@app.route("/")
def dashboard():
    return render_template("dashboard.html")


@app.route("/data")
def data_management():
    return render_template("data_management.html")


@app.route("/data/add")
def add_edit():
    edit_id = request.args.get("id")
    return render_template("add_edit.html", edit_id=edit_id)


@app.route("/analysis")
def analysis_results():
    return render_template("analysis_results.html")


@app.route("/status")
def update_status():
    return render_template("update_status.html")


# ---------------- API: records (Data Management + Add/Edit) ----------------

SEASON_LABELS = {1: "Spring", 2: "Summer", 3: "Fall", 4: "Winter"}
WEATHER_LABELS = {1: "Clear", 2: "Mist/Cloudy", 3: "Light Rain/Snow", 4: "Severe Weather"}


@app.route("/api/years")
def api_years():
    conn = get_conn()
    years = set()
    for table in ("raw_data", "web_entries"):
        rows = conn.execute(f"SELECT DISTINCT strftime('%Y', dteday) AS y FROM {table}").fetchall()
        years.update(int(r["y"]) for r in rows if r["y"])
    conn.close()
    return jsonify(sorted(years))


@app.route("/api/records")
def api_records():
    page = int(request.args.get("page", 1))
    per_page = int(request.args.get("per_page", 25))
    month = request.args.get("month", "").strip()   # 1-12
    year = request.args.get("year", "").strip()      # e.g. 2011
    source = request.args.get("source", "all")  # all | raw | web
    season = request.args.get("season", "")
    weathersit = request.args.get("weathersit", "")

    conn = get_conn()
    results = []

    def fetch(table, tag):
        q = f"SELECT *, '{tag}' AS source FROM {table} WHERE 1=1"
        params = []
        if month:
            q += " AND CAST(strftime('%m', dteday) AS INTEGER) = ?"
            params.append(int(month))
        if year:
            q += " AND CAST(strftime('%Y', dteday) AS INTEGER) = ?"
            params.append(int(year))
        if season:
            q += " AND season = ?"
            params.append(season)
        if weathersit:
            q += " AND weathersit = ?"
            params.append(weathersit)
        return conn.execute(q, params).fetchall()

    if source in ("all", "raw"):
        results += [dict(r) for r in fetch("raw_data", "raw")]
    if source in ("all", "web"):
        results += [dict(r) for r in fetch("web_entries", "web")]

    results.sort(key=lambda r: (r["dteday"], r["hr"]))
    total = len(results)
    start = (page - 1) * per_page
    page_rows = results[start:start + per_page]

    for r in page_rows:
        r["season_label"] = SEASON_LABELS.get(r["season"], r["season"])
        r["weather_label"] = WEATHER_LABELS.get(r["weathersit"], r["weathersit"])
        r["temp_c"] = analytics.temp_to_celsius(r["temp"])
        r["hum_pct"] = analytics.hum_to_percent(r["hum"])

    conn.close()
    return jsonify({
        "total": total, "page": page, "per_page": per_page,
        "total_pages": max(1, (total + per_page - 1) // per_page),
        "records": page_rows,
    })


@app.route("/api/records/<int:instant>")
def api_record_detail(instant):
    source = request.args.get("source", "web")
    table = "raw_data" if source == "raw" else "web_entries"
    conn = get_conn()
    row = conn.execute(f"SELECT * FROM {table} WHERE instant = ?", (instant,)).fetchone()
    conn.close()
    if not row:
        return jsonify({"error": "Record not found"}), 404
    return jsonify(dict(row))


def validate_payload(data):
    errors = []
    required_numeric = ["season", "yr", "mnth", "hr", "holiday", "weekday", "workingday",
                         "weathersit", "casual", "registered"]
    for f in required_numeric:
        if f not in data:
            errors.append(f"Missing field: {f}")
    if not data.get("dteday"):
        errors.append("Date (dteday) is required")
    try:
        if int(data.get("casual", -1)) < 0:
            errors.append("Casual users must be a non-negative integer")
        if int(data.get("registered", -1)) < 0:
            errors.append("Registered users must be a non-negative integer")
    except (TypeError, ValueError):
        errors.append("Casual/Registered must be integers")
    try:
        hr = int(data.get("hr", -1))
        if not (0 <= hr <= 23):
            errors.append("Hour must be between 0 and 23")
    except (TypeError, ValueError):
        errors.append("Hour must be an integer 0-23")

    # holiday / workingday consistency (same rule as the original dataset):
    #   workingday = 1  <=>  Monday-Friday and not a public holiday
    try:
        holiday, workingday, weekday = int(data.get("holiday")), int(data.get("workingday")), int(data.get("weekday"))
        expected = 1 if (1 <= weekday <= 5 and holiday == 0) else 0
        if workingday != expected:
            if holiday == 1:
                errors.append("A public holiday cannot be a working day.")
            elif weekday in (0, 6):
                errors.append("Saturdays and Sundays cannot be working days.")
            else:
                errors.append("A normal Monday-Friday that is not a public holiday must be a working day.")
    except (TypeError, ValueError):
        pass  # missing/invalid values are reported by the checks above

    bounds = {
        "temp_c": (0, 41), "atemp_c": (0, 50), "hum_pct": (0, 100),
        "windspeed_val": (0, 67), "casual": (0, 1000), "registered": (0, 1000),
    }
    for field, (lo, hi) in bounds.items():
        if field in data:
            try:
                v = float(data[field])
                if not (lo <= v <= hi):
                    errors.append(f"{field} must be between {lo} and {hi}")
            except (TypeError, ValueError):
                errors.append(f"{field} must be a number")
    return errors


@app.route("/api/records", methods=["POST"])
def api_create_record():
    data = request.get_json(force=True)
    errors = validate_payload(data)
    if errors:
        return jsonify({"errors": errors}), 400

    casual = int(data["casual"])
    registered = int(data["registered"])
    cnt = casual + registered  # auto-calculated total (spec 5.3)

    # Convert user-facing units -> normalized Excel units (spec 8.1)
    temp_norm = analytics.celsius_to_norm_temp(float(data.get("temp_c", 0)))
    atemp_norm = analytics.celsius_to_norm_atemp(float(data.get("atemp_c", data.get("temp_c", 0))))
    hum_norm = analytics.percent_to_norm_hum(float(data.get("hum_pct", 0)))
    windspeed_norm = analytics.original_to_norm_windspeed(float(data.get("windspeed_val", 0)))

    conn = get_conn()
    cur = conn.execute(
        """INSERT INTO web_entries
           (dteday, season, yr, mnth, hr, holiday, weekday, workingday, weathersit,
            temp, atemp, hum, windspeed, casual, registered, cnt)
           VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",
        (data["dteday"], data["season"], data["yr"], data["mnth"], data["hr"],
         data["holiday"], data["weekday"], data["workingday"], data["weathersit"],
         temp_norm, atemp_norm, hum_norm, windspeed_norm, casual, registered, cnt),
    )
    conn.commit()
    new_id = cur.lastrowid
    conn.close()

    auto_sync_after_change("added", new_id)
    return jsonify({"success": True, "id": new_id, "cnt": cnt}), 201


@app.route("/api/records/<int:instant>", methods=["PUT"])
def api_update_record(instant):
    conn = get_conn()
    existing = conn.execute("SELECT * FROM web_entries WHERE instant = ?", (instant,)).fetchone()
    if not existing:
        conn.close()
        return jsonify({"error": "Record not found, or it belongs to the protected original dataset"}), 404

    data = request.get_json(force=True)
    errors = validate_payload(data)
    if errors:
        conn.close()
        return jsonify({"errors": errors}), 400

    casual = int(data["casual"])
    registered = int(data["registered"])
    cnt = casual + registered

    temp_norm = analytics.celsius_to_norm_temp(float(data.get("temp_c", 0)))
    atemp_norm = analytics.celsius_to_norm_atemp(float(data.get("atemp_c", data.get("temp_c", 0))))
    hum_norm = analytics.percent_to_norm_hum(float(data.get("hum_pct", 0)))
    windspeed_norm = analytics.original_to_norm_windspeed(float(data.get("windspeed_val", 0)))

    conn.execute(
        """UPDATE web_entries SET dteday=?, season=?, yr=?, mnth=?, hr=?, holiday=?, weekday=?,
           workingday=?, weathersit=?, temp=?, atemp=?, hum=?, windspeed=?, casual=?, registered=?,
           cnt=?, updated_at=datetime('now') WHERE instant=?""",
        (data["dteday"], data["season"], data["yr"], data["mnth"], data["hr"],
         data["holiday"], data["weekday"], data["workingday"], data["weathersit"],
         temp_norm, atemp_norm, hum_norm, windspeed_norm, casual, registered, cnt, instant),
    )
    conn.commit()
    conn.close()

    auto_sync_after_change("edited", instant)
    return jsonify({"success": True, "cnt": cnt})


@app.route("/api/records/<int:instant>", methods=["DELETE"])
def api_delete_record(instant):
    """Spec 1.1 / 8.3: only website-added (web_entries) records can be deleted.
    The original Raw_Data must be protected."""
    conn = get_conn()
    existing = conn.execute("SELECT * FROM web_entries WHERE instant = ?", (instant,)).fetchone()
    if not existing:
        conn.close()
        return jsonify({"error": "Only records added through the website can be deleted."}), 403
    conn.execute("DELETE FROM web_entries WHERE instant = ?", (instant,))
    conn.commit()
    conn.close()
    auto_sync_after_change("deleted", instant)
    return jsonify({"success": True})


# ---------------- API: dashboard / analysis / sync status ----------------

@app.route("/api/dashboard")
def api_dashboard():
    conn = get_conn()
    rows = analytics.combined_rows(conn)
    conn.close()
    total_rentals = sum(r["cnt"] for r in rows)
    total_casual = sum(r["casual"] for r in rows)
    total_registered = sum(r["registered"] for r in rows)
    avg_per_hour = round(total_rentals / len(rows), 2) if rows else 0

    conn = get_conn()
    monthly = analytics.compute_monthly_summary(conn)
    hourly = analytics.compute_pivot(conn, "hr")
    season = analytics.compute_pivot(conn, "season")
    weather = analytics.compute_pivot(conn, "weathersit")
    conn.close()

    return jsonify({
        "total_rentals": total_rentals,
        "total_casual": total_casual,
        "total_registered": total_registered,
        "avg_per_hour": avg_per_hour,
        "record_count": len(rows),
        "monthly": monthly,
        "hourly": hourly,
        "season": {SEASON_LABELS.get(int(k), k): v for k, v in season.items()},
        "weather": {WEATHER_LABELS.get(int(k), k): v for k, v in weather.items()},
    })


@app.route("/api/analysis")
def api_analysis():
    result = analytics.run_full_analysis()
    result["season_labels"] = {str(k): v for k, v in
                                {k: SEASON_LABELS[k] for k in SEASON_LABELS}.items()}
    result["weather_labels"] = {str(k): v for k, v in
                                 {k: WEATHER_LABELS[k] for k in WEATHER_LABELS}.items()}
    return jsonify(result)


@app.route("/api/analysis/update", methods=["POST"])
def api_analysis_update():
    try:
        result = analytics.run_full_analysis()
        log_sync("manual_sync", "success", records_synced=result["descriptive"].get("count", 0),
                  message="Manual 'Update Analysis' triggered by user.")
        return jsonify({"success": True, "result": result})
    except Exception as e:
        log_sync("manual_sync", "error", message=str(e))
        return jsonify({"success": False, "error": str(e)}), 500


@app.route("/api/sync-log")
def api_sync_log():
    conn = get_conn()
    rows = conn.execute(
        "SELECT * FROM sync_log ORDER BY id DESC LIMIT 100"
    ).fetchall()
    conn.close()
    return jsonify([dict(r) for r in rows])


# ---------------- Excel export endpoints (spec section 7 & 9) ----------------

@app.route("/export/latest")
def export_latest():
    try:
        path, count = analytics.export_latest_data()
        log_sync("export", "success", records_synced=count,
                  message="Bike_Sharing_Export.xlsx generated from current database.")
        return send_file(path, as_attachment=True, download_name="Bike_Sharing_Export.xlsx")
    except Exception as e:
        log_sync("export", "error", message=str(e))
        return jsonify({"error": str(e)}), 500


@app.route("/export/updated-workbook")
def export_updated_workbook():
    try:
        path, analysis = analytics.generate_updated_analysis_workbook()
        log_sync("analysis_update", "success",
                  records_synced=analysis["descriptive"].get("count", 0),
                  message="Bike_Sharing_Analysis_Updated.xlsx generated (template preserved).")
        return send_file(path, as_attachment=True, download_name="Bike_Sharing_Analysis_Updated.xlsx")
    except Exception as e:
        log_sync("analysis_update", "error", message=str(e))
        return jsonify({"error": str(e)}), 500


# ---------- Tableau Public: data files + saved dashboard URL ----------

def _tableau_export(builder, filename, label):
    try:
        path, count = builder()
        log_sync("export", "success", records_synced=count,
                  message=f"{filename} generated for Tableau Public ({label}).")
        return send_file(path, as_attachment=True, download_name=filename)
    except Exception as e:
        log_sync("export", "error", message=str(e))
        return jsonify({"error": str(e)}), 500


@app.route("/export/tableau-csv")
def export_tableau_csv():
    return _tableau_export(analytics.export_tableau_csv, "Bike_Sharing_Tableau_Data.csv", "CSV")


@app.route("/export/tableau-xlsx")
def export_tableau_xlsx():
    return _tableau_export(analytics.export_tableau_xlsx, "Bike_Sharing_Tableau_Data.xlsx", "Excel")


@app.route("/api/tableau-url")
def api_get_tableau_url():
    conn = get_conn()
    row = conn.execute("SELECT value FROM settings WHERE key='tableau_url'").fetchone()
    conn.close()
    if row:
        return jsonify({"url": row["value"], "is_default": False})
    return jsonify({"url": DEFAULT_TABLEAU_URL, "is_default": True})


@app.route("/api/tableau-url", methods=["POST"])
def api_set_tableau_url():
    data = request.get_json(force=True) or {}
    url = str(data.get("url", "")).strip()
    parsed = urlparse(url)
    if parsed.scheme not in ("http", "https") or not parsed.netloc:
        return jsonify({"error": "Please enter a valid URL starting with https://"}), 400
    conn = get_conn()
    conn.execute("INSERT INTO settings (key, value) VALUES ('tableau_url', ?) "
                 "ON CONFLICT(key) DO UPDATE SET value=excluded.value", (url,))
    conn.commit()
    conn.close()
    return jsonify({"success": True, "url": url})


# ---------- Dashboard pipeline: Excel -> .hyper -> .twb ----------

TABLEAU_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "static", "tableau")


@app.route("/api/tableau/refresh", methods=["POST"])
def api_tableau_refresh():
    result = refresh_tableau_files("Manual refresh")
    if "error" in result:
        return jsonify({"success": False, "error": result["error"]}), 500
    return jsonify({"success": True, **result})


@app.route("/api/tableau/status")
def api_tableau_status():
    conn = get_conn()
    last = conn.execute(
        "SELECT event_time, status, records_synced, message FROM sync_log "
        "WHERE event_type='tableau_sync' ORDER BY id DESC LIMIT 1").fetchone()
    conn.close()
    files = {}
    for key, name in (("hyper", "Bike_Sharing.hyper"), ("twb", "Project_data_Analysis.twb"),
                      ("twbx", "Project_data_Analysis.twbx")):
        path = os.path.join(TABLEAU_DIR, name)
        files[key] = {"name": name, "exists": os.path.exists(path),
                      "modified": datetime.datetime.fromtimestamp(os.path.getmtime(path)).isoformat(timespec="seconds")
                      if os.path.exists(path) else None}
    return jsonify({"last_run": dict(last) if last else None, "files": files})


@app.route("/export/tableau-workbook")
def export_tableau_workbook():
    path = os.path.join(TABLEAU_DIR, "Project_data_Analysis.twbx")
    if not os.path.exists(path):
        refresh_tableau_files("Download requested")
    return send_file(path, as_attachment=True, download_name="Project_data_Analysis.twbx")


@app.route("/export/tableau-hyper")
def export_tableau_hyper():
    path = os.path.join(TABLEAU_DIR, "Bike_Sharing.hyper")
    if not os.path.exists(path):
        refresh_tableau_files("Download requested")
    return send_file(path, as_attachment=True, download_name="Bike_Sharing.hyper")


if __name__ == "__main__":
    app.run(debug=False, port=5000)
