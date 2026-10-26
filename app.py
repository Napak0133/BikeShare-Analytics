"""
app.py — BikeShare Analytics website (Flask)

Data flow (see pipeline.py):
    website form  ->  Excel workbook (Web_Entries sheet, same file, no copies)
                  ->  Bike_Sharing.hyper rebuilt from the workbook
                  ->  Dashboard charts queried from the .hyper
                  ->  open Dashboard pages redraw straight away

Pages
  /            Dashboard
  /data        Data Management (browse, edit, delete)
  /data/add    Add / Edit record (shared form)
  /analysis    Analysis Results
  /status      Data Update Status

API (used by static/js/main.js and the page scripts)
  GET    /api/records, /api/records/<id>, /api/years
  POST   /api/records               add a record   -> Excel -> .hyper
  PUT    /api/records/<id>          edit a record  -> Excel -> .hyper
  DELETE /api/records/<id>          delete a record-> Excel -> .hyper
  GET    /api/dashboard             chart data, queried from the .hyper
  GET    /api/version               tiny "has the data changed?" check the Dashboard polls
  GET    /api/analysis              statistics over Raw_Data + Web_Entries
  POST   /api/analysis/update       re-read the Excel file and rebuild the .hyper now
  GET    /api/sync-log              history for the Data Update Status page
  GET    /export/workbook           download the live Excel workbook
"""

import datetime
import os

from dotenv import load_dotenv
from flask import Flask, jsonify, render_template, request, send_file

load_dotenv()

import analytics
import datastore
import hyperstore
import models
import pipeline
from tableau_embed import TableauEmbedConfigError, create_embed_token

app = Flask(__name__)

SEASON_LABELS = {1: "Spring", 2: "Summer", 3: "Fall", 4: "Winter"}
WEATHER_LABELS = {1: "Clear", 2: "Mist/Cloudy", 3: "Light Rain/Snow", 4: "Severe Weather"}

# Make sure the .hyper exists and matches the Excel file before the first request.
try:
    pipeline.ensure_fresh("Startup")
except Exception as exc:                       # the error is already in the sync log
    print(f"[startup] could not build the .hyper: {exc}")


@app.errorhandler(datastore.ExcelLockedError)
def excel_locked(e):
    return jsonify({"error": str(e)}), 423


# ---------------- Pages ----------------

@app.route("/")
def dashboard():
    return render_template("dashboard.html", active="dashboard")


@app.route("/tableau")
def tableau_dashboard():
    server_url = os.environ.get(
        "TABLEAU_SERVER_URL", "https://prod-apsoutheast-b.online.tableau.com"
    ).rstrip("/")
    view_url = os.environ.get("TABLEAU_VIEW_URL", "").strip()
    embed_token = None
    embed_error = None
    if view_url:
        try:
            embed_token = create_embed_token()
        except TableauEmbedConfigError as exc:
            embed_error = str(exc)
    return render_template(
        "tableau_dashboard.html",
        active="tableau",
        tableau_view_url=view_url,
        tableau_embed_token=embed_token,
        tableau_embed_error=embed_error,
        tableau_embed_script_url=(
            f"{server_url}/javascripts/api/tableau.embedding.3.latest.min.js"
        ),
    )


@app.route("/data")
def data_management():
    return render_template("data_management.html", active="data")


@app.route("/data/add")
def add_edit():
    return render_template("add_edit.html", edit_id=request.args.get("id"), active="data")


@app.route("/analysis")
def analysis_results():
    return render_template("analysis_results.html", active="analysis")


@app.route("/status")
def update_status():
    return render_template("update_status.html", active="status")


# ---------------- Records (Data Management + Add/Edit) ----------------

@app.route("/api/years")
def api_years():
    return jsonify(sorted({int(r["dteday"][:4]) for r in datastore.all_rows()}))


@app.route("/api/records")
def api_records():
    page = max(1, int(request.args.get("page", 1)))
    per_page = max(1, min(200, int(request.args.get("per_page", 25))))
    month, year = request.args.get("month", "").strip(), request.args.get("year", "").strip()
    source = request.args.get("source", "all")                  # all | raw | web
    season, weathersit = request.args.get("season", ""), request.args.get("weathersit", "")

    def keep(r):
        return ((source == "all" or r["source"] == source)
                and (not month or int(r["dteday"][5:7]) == int(month))
                and (not year or int(r["dteday"][:4]) == int(year))
                and (not season or r["season"] == int(season))
                and (not weathersit or r["weathersit"] == int(weathersit)))

    results = [r for r in datastore.all_rows() if keep(r)]
    total = len(results)
    start = (page - 1) * per_page
    page_rows = [dict(r) for r in results[start:start + per_page]]
    for r in page_rows:
        r["season_label"] = SEASON_LABELS.get(r["season"], r["season"])
        r["weather_label"] = WEATHER_LABELS.get(r["weathersit"], r["weathersit"])
        r["temp_c"] = analytics.temp_to_celsius(r["temp"])
        r["hum_pct"] = analytics.hum_to_percent(r["hum"])
    return jsonify({"total": total, "page": page, "per_page": per_page,
                    "total_pages": max(1, (total + per_page - 1) // per_page), "records": page_rows})


@app.route("/api/records/<int:instant>")
def api_record_detail(instant):
    rows = datastore.load_raw() if request.args.get("source") == "raw" else datastore.load_web()
    for r in rows:
        if r["instant"] == instant:
            return jsonify(r)
    return jsonify({"error": "Record not found"}), 404


def build_record(data, ignore_instant=None):
    """Validate the form payload and convert it to the 17 Excel columns.
    Returns (record, errors, http_status)."""
    errors = []

    def num(name, lo, hi, integer=False):
        try:
            v = float(data[name])
            if integer and v != int(v):
                raise ValueError
        except (KeyError, TypeError, ValueError):
            errors.append(f"{name} is required and must be {'an integer' if integer else 'a number'}")
            return None
        if not (lo <= v <= hi):
            errors.append(f"{name} must be between {lo} and {hi}")
            return None
        return int(v) if integer else v

    try:
        day = datetime.date.fromisoformat(str(data.get("dteday", ""))[:10])
        if day.year < 2011:
            errors.append("Date must be in 2011 or later")
    except ValueError:
        day = None
        errors.append("Date (dteday) is required, format YYYY-MM-DD")

    hr = num("hr", 0, 23, True)
    season = num("season", 1, 4, True)
    holiday = num("holiday", 0, 1, True)
    workingday = num("workingday", 0, 1, True)
    weathersit = num("weathersit", 1, 4, True)
    temp_c = num("temp_c", 0, 41)
    atemp_c = num("atemp_c", 0, 50)
    hum_pct = num("hum_pct", 0, 100)
    wind = num("windspeed_val", 0, 67)
    casual = num("casual", 0, 1000, True)
    registered = num("registered", 0, 1000, True)
    if errors:
        return None, errors, 400

    # yr / mnth / weekday always come from the date itself (never trusted from the browser)
    weekday = day.isoweekday() % 7                   # Sunday = 0 ... Saturday = 6, like the dataset
    expected = 1 if (1 <= weekday <= 5 and holiday == 0) else 0
    if workingday != expected:
        if holiday == 1:
            errors.append("A public holiday cannot be a working day.")
        elif weekday in (0, 6):
            errors.append("Saturdays and Sundays cannot be working days.")
        else:
            errors.append("A normal Monday-Friday that is not a public holiday must be a working day.")
        return None, errors, 400

    dup = datastore.find_duplicate(day.isoformat(), hr, ignore_instant)
    if dup:
        where = "the original data" if dup["source"] == "raw" else "the website entries"
        return None, [f"A record for {day.isoformat()} at hour {hr} already exists in {where} "
                      f"(id {dup['instant']}). Edit that record instead."], 409

    rec = {
        "dteday": day.isoformat(), "season": season, "yr": day.year - 2011, "mnth": day.month, "hr": hr,
        "holiday": holiday, "weekday": weekday, "workingday": workingday, "weathersit": weathersit,
        "temp": analytics.celsius_to_norm_temp(temp_c), "atemp": analytics.celsius_to_norm_atemp(atemp_c),
        "hum": analytics.percent_to_norm_hum(hum_pct),
        "windspeed": analytics.original_to_norm_windspeed(wind),
        "casual": casual, "registered": registered, "cnt": casual + registered,
    }
    return rec, [], 200


def _is_raw(instant):
    return any(r["instant"] == instant for r in datastore.load_raw())


@app.route("/api/records", methods=["POST"])
def api_create_record():
    with pipeline.lock:                              # check + write as one step
        rec, errors, status = build_record(request.get_json(force=True) or {})
        if errors:
            return jsonify({"errors": errors}), status
        new_id, warning = pipeline.commit("add", "new", lambda: datastore.add_entry(rec))
    return jsonify({"success": True, "id": new_id, "cnt": rec["cnt"], "warning": warning}), 201


@app.route("/api/records/<int:instant>", methods=["PUT"])
def api_update_record(instant):
    with pipeline.lock:
        if not any(r["instant"] == instant for r in datastore.load_web()):
            if _is_raw(instant):
                return jsonify({"error": "This record belongs to the protected original dataset."}), 403
            return jsonify({"error": "Record not found."}), 404
        rec, errors, status = build_record(request.get_json(force=True) or {}, ignore_instant=instant)
        if errors:
            return jsonify({"errors": errors}), status
        _, warning = pipeline.commit("edit", instant, lambda: datastore.update_entry(instant, rec))
    return jsonify({"success": True, "cnt": rec["cnt"], "warning": warning})


@app.route("/api/records/<int:instant>", methods=["DELETE"])
def api_delete_record(instant):
    """Only website-added records can be deleted; the original Raw_Data stays protected."""
    with pipeline.lock:
        if not any(r["instant"] == instant for r in datastore.load_web()):
            return jsonify({"error": "Only records added through the website can be deleted."}), 403
        _, warning = pipeline.commit("delete", instant, lambda: datastore.delete_entry(instant))
    return jsonify({"success": True, "warning": warning})


# ---------------- Dashboard (reads the .hyper) ----------------

@app.route("/api/version")
def api_version():
    pipeline.ensure_fresh()                          # picks up hand-edits made in Excel
    return jsonify({"version": pipeline.state["version"], "built_at": pipeline.state["built_at"]})


@app.route("/api/dashboard")
def api_dashboard():
    pipeline.ensure_fresh()
    d = hyperstore.dashboard_data()
    d["season"] = {SEASON_LABELS.get(k, k): v for k, v in d["season"].items()}
    d["weather"] = {WEATHER_LABELS.get(k, k): v for k, v in d["weather"].items()}
    d["version"] = pipeline.state["version"]
    d["built_at"] = pipeline.state["built_at"]
    return jsonify(d)


@app.route("/api/overview")
def api_overview():
    """Overview dashboard data: ?year=0|2011|2012&user=cnt|casual|registered&daytype=all|weekday|off"""
    pipeline.ensure_fresh()
    try:
        year = int(request.args.get("year", 0))
    except ValueError:
        year = 0
    user = request.args.get("user", "cnt")
    daytype = request.args.get("daytype", "all")
    if daytype not in ("all", "weekday", "off"):
        daytype = "all"
    d = hyperstore.overview_data(year, user, daytype)
    d["version"] = pipeline.state["version"]
    d["built_at"] = pipeline.state["built_at"]
    return jsonify(d)


def _filters():
    """Year / User Type / Day Type from the query string (same meaning as on the Overview page)."""
    try:
        year = int(request.args.get("year", 0))
    except ValueError:
        year = 0
    user = request.args.get("user", "cnt")
    if user not in ("cnt", "casual", "registered"):
        user = "cnt"
    daytype = request.args.get("daytype", "all")
    if daytype not in ("all", "weekday", "off"):
        daytype = "all"
    return year, user, daytype


def _stamp(d):
    d["version"] = pipeline.state["version"]
    d["built_at"] = pipeline.state["built_at"]
    d["years"] = hyperstore.years()
    d["record_count"] = pipeline.state["rows"]
    return d


@app.route("/api/hour")
def api_hour():
    """Hour view: ?year=&user=&daytype=  -> average rentals for each hour of the day."""
    pipeline.ensure_fresh()
    return jsonify(_stamp(hyperstore.hour_data(*_filters())))


@app.route("/api/weather")
def api_weather():
    """Weather view: ?year=&user=&daytype=  -> average rentals by weather, temperature, humidity, wind."""
    pipeline.ensure_fresh()
    return jsonify(_stamp(hyperstore.weather_data(*_filters())))


@app.route("/api/forecast")
def api_forecast():
    """Forecast view: ?user=cnt|casual|registered&months=3|6|12 (all years are used for the trend)."""
    pipeline.ensure_fresh()
    _, user, _ = _filters()
    try:
        months = int(request.args.get("months", 12))
    except ValueError:
        months = 12
    if months not in models.MONTHS_AHEAD:
        months = 12
    d = models.forecast_monthly(hyperstore.monthly_series(user), months)
    return jsonify(_stamp({"user": user, "result": d}))


_whatif_cache = {"version": None, "model": None}


@app.route("/api/whatif")
def api_whatif():
    """What-if view: the fitted model (cached until the data changes). The browser does the arithmetic."""
    pipeline.ensure_fresh()
    v = pipeline.state["version"]
    if _whatif_cache["version"] != v:
        _whatif_cache.update(version=v, model=models.fit_whatif(hyperstore.model_rows()))
    return jsonify(_stamp({"model": _whatif_cache["model"]}))


# ---------------- Analysis + status ----------------

@app.route("/api/analysis")
def api_analysis():
    pipeline.ensure_fresh()
    return jsonify(analytics.run_full_analysis(datastore.all_rows()))


@app.route("/api/analysis/update", methods=["POST"])
def api_analysis_update():
    """'Re-sync now': rebuild the .hyper from Excel and publish to Tableau Cloud when configured."""
    try:
        with pipeline.lock:
            n = pipeline.rebuild_hyper()
            warning = pipeline.publish_tableau(n)
            message = "Manual re-sync: .hyper rebuilt from the Excel file."
            if warning:
                message += f" {warning}"
            else:
                message += " Tableau Cloud workbook published."
            pipeline.log("manual_sync", "success", n, message)
        return jsonify({"success": True, "records": n, "warning": warning})
    except Exception as e:
        pipeline.log("manual_sync", "error", 0, str(e))
        return jsonify({"success": False, "error": str(e)}), 500


@app.route("/api/sync-log")
def api_sync_log():
    return jsonify(pipeline.read_log())


@app.route("/export/workbook")
def export_workbook():
    return send_file(datastore.EXCEL_PATH, as_attachment=True,
                     download_name="Bike_Sharing_Analysis_Final.xlsx")


if __name__ == "__main__":
    app.run(debug=False, port=5000)
