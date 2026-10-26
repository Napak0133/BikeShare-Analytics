"""
models.py
---------
Small, dependency-free models behind the Forecast and What-if views of the Dashboard.

* forecast_monthly()  trend x seasonality: every calendar month gets a seasonal index, the de-seasonalised
                      series gets a straight-line trend, and the forecast is trend x index.
* fit_whatif()        log-linear demand model:  ln(rentals + 1) = base[hour, day type] + b . features
                      features = weather (3 dummies), temperature, temperature^2, humidity, wind speed, year.
                      The coefficients go to the browser, so the What-if sliders answer instantly.
"""

import math

MONTHS_AHEAD = (3, 6, 12)


# ------------------------------------------------------------------ tiny linear algebra

def _solve(a, b):
    """Solve a.x = b (Gaussian elimination with partial pivoting). `a` is small (<= ~10 x 10)."""
    n = len(b)
    m = [row[:] + [b[i]] for i, row in enumerate(a)]
    for c in range(n):
        p = max(range(c, n), key=lambda r: abs(m[r][c]))
        if abs(m[p][c]) < 1e-12:
            m[p][c] = 1e-12                       # ridge-like guard for a degenerate column
        m[c], m[p] = m[p], m[c]
        for r in range(c + 1, n):
            f = m[r][c] / m[c][c]
            for k in range(c, n + 1):
                m[r][k] -= f * m[c][k]
    x = [0.0] * n
    for i in range(n - 1, -1, -1):
        x[i] = (m[i][n] - sum(m[i][k] * x[k] for k in range(i + 1, n))) / m[i][i]
    return x


def _ols(xs, ys):
    """Least squares with an intercept in column 0 (xs rows already contain the leading 1)."""
    k = len(xs[0])
    xtx = [[0.0] * k for _ in range(k)]
    xty = [0.0] * k
    for row, y in zip(xs, ys):
        for i in range(k):
            xi = row[i]
            xty[i] += xi * y
            for j in range(i, k):
                xtx[i][j] += xi * row[j]
    for i in range(k):
        for j in range(i):
            xtx[i][j] = xtx[j][i]
    return _solve(xtx, xty)


# ------------------------------------------------------------------ forecast

def forecast_monthly(series, horizon=12):
    """series: [{year, month, value, records}] -> history + forecast.  Returns None if there is too little data."""
    if not series:
        return None
    # A month that only holds a few website entries would drag the trend down: keep months with enough records.
    full = sorted(r["records"] for r in series)[len(series) // 2]
    hist = [r for r in series if r["records"] >= 0.6 * full]
    if len(hist) < 6:
        return None
    t0 = hist[0]["year"] * 12 + hist[0]["month"]
    tt = [r["year"] * 12 + r["month"] - t0 for r in hist]
    vv = [float(r["value"]) for r in hist]

    seas = {m: 1.0 for m in range(1, 13)}
    a = b = 0.0
    for _ in range(25):                                         # alternate: trend <-> seasonal index
        des = [v / seas[r["month"]] for v, r in zip(vv, hist)]
        n = len(tt)
        mt, mv = sum(tt) / n, sum(des) / n
        sxx = sum((x - mt) ** 2 for x in tt)
        b = sum((x - mt) * (y - mv) for x, y in zip(tt, des)) / sxx if sxx else 0.0
        a = mv - b * mt
        ratios = {}
        for x, v, r in zip(tt, vv, hist):
            tr = a + b * x
            if tr > 0:
                ratios.setdefault(r["month"], []).append(v / tr)
        raw = {m: sum(rs) / len(rs) for m, rs in ratios.items()}
        mean = sum(raw.values()) / len(raw) if raw else 1.0
        seas = {m: (raw[m] / mean if m in raw else 1.0) for m in range(1, 13)}

    fit = [max(0.0, (a + b * x) * seas[r["month"]]) for x, r in zip(tt, hist)]
    rmse = math.sqrt(sum((v - f) ** 2 for v, f in zip(vv, fit)) / len(vv))
    last = hist[-1]
    out = []
    y, m = last["year"], last["month"]
    for i in range(1, int(horizon) + 1):
        m += 1
        if m > 12:
            m, y = 1, y + 1
        x = y * 12 + m - t0
        f = max(0.0, (a + b * x) * seas[m])
        out.append({"year": y, "month": m, "value": round(f), "low": round(max(0.0, f - 1.96 * rmse)),
                    "high": round(f + 1.96 * rmse)})
    last12 = sum(r["value"] for r in hist[-12:])
    return {
        "history": [{"year": r["year"], "month": r["month"], "value": r["value"], "fit": round(f)}
                    for r, f in zip(hist, fit)],
        "forecast": out, "horizon": int(horizon),
        "forecast_total": sum(o["value"] for o in out),
        "last12_total": last12, "last12_months": min(12, len(hist)),
        "trend_per_month": round(b), "rmse": round(rmse),
        "months_used": len(hist), "months_skipped": len(series) - len(hist),
    }


# ------------------------------------------------------------------ what-if

FEATURES = ["w2", "w3", "temp", "temp2", "hum", "wind", "yr"]
TARGETS = {"cnt": 9, "casual": 7, "registered": 8}            # column index inside model_rows()


def _features(wsit, temp, hum, wind, yr, stats):
    """Feature vector (without the intercept). temp/hum/wind are in everyday units (°C, %, wind speed)."""
    zt = (temp - stats["temp"][0]) / stats["temp"][1]
    return [1.0 if wsit == 2 else 0.0, 1.0 if wsit >= 3 else 0.0, zt, zt * zt,
            (hum - stats["hum"][0]) / stats["hum"][1], (wind - stats["wind"][0]) / stats["wind"][1],
            float(yr)]


def _mean_std(vals):
    n = len(vals)
    m = sum(vals) / n
    s = math.sqrt(sum((v - m) ** 2 for v in vals) / n) or 1.0
    return [m, s]


def fit_whatif(rows):
    """rows from hyperstore.model_rows(). Returns the model the page uses, or None if there is too little data."""
    # Use only years with a real amount of data: a lone test record from another year would otherwise
    # dominate the "year" coefficient.
    per_year = {}
    for r in rows:
        per_year[r[6]] = per_year.get(r[6], 0) + 1
    keep = {y for y, n in per_year.items() if n >= 500}
    rows = [r for r in rows if r[6] in keep]
    if len(rows) < 500:
        return None
    temps = [r[3] * 41 for r in rows]
    hums = [r[4] * 100 for r in rows]
    winds = [r[5] * 67 for r in rows]
    stats = {"temp": _mean_std(temps), "hum": _mean_std(hums), "wind": _mean_std(winds)}
    xs = [_features(r[2], t, h, w, r[6], stats) for r, t, h, w in zip(rows, temps, hums, winds)]
    cells = [(r[0], int(r[1])) for r in rows]

    models = {}
    for name, idx in TARGETS.items():
        y = [math.log(r[idx] + 1) for r in rows]
        gm = sum(y) / len(y)
        base, beta = {}, [0.0] * len(FEATURES)
        for _ in range(4):                                       # backfitting: base <-> coefficients
            sums = {}
            for c, yy, x in zip(cells, y, xs):
                s = sums.setdefault(c, [0.0, 0])
                s[0] += yy - sum(bj * xj for bj, xj in zip(beta, x))
                s[1] += 1
            base = {c: s[0] / s[1] for c, s in sums.items()}
            sol = _ols([[1.0] + x for x in xs], [yy - base[c] for c, yy in zip(cells, y)])
            beta = sol[1:]
            for c in base:
                base[c] += sol[0]
        for hr in range(24):                                     # hour/day-type cell with no data: borrow a neighbour
            for wd in (0, 1):
                if (hr, wd) not in base:
                    base[(hr, wd)] = base.get((hr, 1 - wd), gm)
        resid = [yy - base[c] - sum(bj * xj for bj, xj in zip(beta, x)) for c, yy, x in zip(cells, y, xs)]
        smear = sum(math.exp(e) for e in resid) / len(resid)
        actual = [r[idx] for r in rows]
        pred = [max(0.0, math.exp(yy - e) * smear - 1) for yy, e in zip(y, resid)]
        ma = sum(actual) / len(actual)
        ss_res = sum((a - p) ** 2 for a, p in zip(actual, pred))
        ss_tot = sum((a - ma) ** 2 for a in actual) or 1.0
        models[name] = {
            "base": [[round(base[(h, wd)], 5) for wd in (0, 1)] for h in range(24)],   # base[hour][workingday]
            "beta": [round(b, 5) for b in beta], "smear": round(smear, 5),
            "r2": round(1 - ss_res / ss_tot, 3),
        }

    years = sorted({2011 + int(r[6]) for r in rows})
    wcount = {}
    for r in rows:
        wcount[r[2]] = wcount.get(r[2], 0) + 1
    return {
        "features": FEATURES, "stats": stats, "models": models, "years": years,
        "typical": {"temp": round(stats["temp"][0], 1), "hum": round(stats["hum"][0], 1),
                    "wind": round(stats["wind"][0], 1), "weather": max(wcount, key=wcount.get),
                    "year": years[-1]},
        "ranges": {"temp": [round(min(temps)), round(max(temps))], "hum": [0, 100],
                   "wind": [0, round(max(winds))]},
        "records": len(rows),
    }
