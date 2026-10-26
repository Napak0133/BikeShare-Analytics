"""
analytics.py
------------
Statistics for the Analysis Results page, calculated over the rows read from the Excel workbook
(Raw_Data + Web_Entries), plus the unit conversions between everyday units and the normalized
values stored in Excel (Data_Dictionary):

    temp_celsius       = temp * 41          atemp_celsius = atemp * 50
    hum_percent        = hum * 100          windspeed_original = windspeed * 67
"""

import datetime
import statistics

# ---------- unit conversion helpers ----------

def temp_to_celsius(v): return round(v * 41, 2)
def celsius_to_norm_temp(c): return round(c / 41, 4)
def atemp_to_celsius(v): return round(v * 50, 2)
def celsius_to_norm_atemp(c): return round(c / 50, 4)
def hum_to_percent(v): return round(v * 100, 1)
def percent_to_norm_hum(p): return round(p / 100, 4)
def windspeed_to_original(v): return round(v * 67, 2)
def original_to_norm_windspeed(w): return round(w / 67, 4)


# ---------- descriptive statistics ----------

def compute_descriptive_stats(rows):
    cnts = [r["cnt"] for r in rows]
    if not cnts:
        return {}
    try:
        mode = statistics.mode(cnts)
    except statistics.StatisticsError:
        mode = None
    n = len(cnts)
    stdev = statistics.stdev(cnts) if n > 1 else 0
    return {
        "count": n, "mean": round(statistics.mean(cnts), 4), "median": statistics.median(cnts),
        "mode": mode, "std_dev": round(stdev, 4), "variance": round(stdev ** 2, 4),
        "min": min(cnts), "max": max(cnts), "total_rentals": sum(cnts),
    }


# ---------- correlation / regression ----------

def _pearson(x, y):
    n = len(x)
    if n < 2:
        return 0.0
    mx, my = sum(x) / n, sum(y) / n
    num = sum((a - mx) * (b - my) for a, b in zip(x, y))
    dx = sum((a - mx) ** 2 for a in x) ** 0.5
    dy = sum((b - my) ** 2 for b in y) ** 0.5
    return 0.0 if dx == 0 or dy == 0 else num / (dx * dy)


def compute_correlation(rows):
    temp, hum = [r["temp"] for r in rows], [r["hum"] for r in rows]
    wind, cnt = [r["windspeed"] for r in rows], [r["cnt"] for r in rows]
    return {
        "temp_cnt": round(_pearson(temp, cnt), 4), "hum_cnt": round(_pearson(hum, cnt), 4),
        "windspeed_cnt": round(_pearson(wind, cnt), 4), "temp_hum": round(_pearson(temp, hum), 4),
    }


def compute_regression(rows):
    x, y = [r["temp"] for r in rows], [r["cnt"] for r in rows]
    n = len(x)
    if n < 2:
        return {}
    mx, my = sum(x) / n, sum(y) / n
    sxx = sum((a - mx) ** 2 for a in x)
    sxy = sum((a - mx) * (b - my) for a, b in zip(x, y))
    slope = sxy / sxx if sxx else 0
    intercept = my - slope * mx
    r = _pearson(x, y)
    return {
        "slope": round(slope, 4), "intercept": round(intercept, 4), "r_value": round(r, 4),
        "r_squared": round(r ** 2, 4),
        "equation": f"cnt = {round(intercept, 2)} + {round(slope, 2)} * temp", "observations": n,
    }


# ---------- pivot-style summaries ----------

def compute_pivot(rows, field):
    totals = {}
    for r in rows:
        totals[r[field]] = totals.get(r[field], 0) + r["cnt"]
    return dict(sorted(totals.items()))


def run_full_analysis(rows):
    return {
        "descriptive": compute_descriptive_stats(rows),
        "correlation": compute_correlation(rows),
        "regression": compute_regression(rows),
        "pivot_hourly": compute_pivot(rows, "hr"),
        "pivot_season": compute_pivot(rows, "season"),
        "pivot_weather": compute_pivot(rows, "weathersit"),
        "computed_at": datetime.datetime.now().isoformat(timespec="seconds"),
    }
