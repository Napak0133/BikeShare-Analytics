"""
datastore.py
------------
The Excel workbook  data/Bike_Sharing_Analysis_Final.xlsx  is the ONE place where the data lives.

    Raw_Data     original 17,379 records  (read only - the website never changes it)
    Web_Entries  records added on the website (add / edit / delete happen here, in place)

Why this file does not use openpyxl to save
-------------------------------------------
openpyxl rewrites the whole workbook when it saves and is known to damage pivot tables, so the
workbook's 3 pivot tables / 3 Excel tables could be lost. Instead, this module edits ONLY the
Web_Entries sheet (and the size of its Excel table) inside the .xlsx package and copies every other
part of the file byte-for-byte. Raw_Data, the pivot tables, formatting and the other sheets are
therefore exactly what you sent.

Everything is written atomically (temp file, then replace), so a crash can never leave a half-written
workbook behind.
"""

import datetime
import os
import re
import threading
import zipfile
from posixpath import normpath, dirname, join as pjoin

from lxml import etree

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
EXCEL_PATH = os.path.join(BASE_DIR, "data", "Bike_Sharing_Analysis_Final.xlsx")

RAW_SHEET = "Raw_Data"
WEB_SHEET = "Web_Entries"

COLUMNS = ["instant", "dteday", "season", "yr", "mnth", "hr", "holiday", "weekday", "workingday",
           "weathersit", "temp", "atemp", "hum", "windspeed", "casual", "registered", "cnt"]
INT_COLS = {"instant", "season", "yr", "mnth", "hr", "holiday", "weekday", "workingday",
            "weathersit", "casual", "registered", "cnt"}
COL_LETTERS = "ABCDEFGHIJKLMNOPQ"            # A:Q, same as the original workbook

NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
RNS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PKG_RNS = "http://schemas.openxmlformats.org/package/2006/relationships"
Q = lambda tag: f"{{{NS}}}{tag}"

EXCEL_EPOCH = datetime.date(1899, 12, 30)
_CELL_RE = re.compile(r"^([A-Z]+)(\d+)$")

write_lock = threading.RLock()               # one writer at a time
_raw_cache = {"key": None, "rows": None}


class ExcelLockedError(Exception):
    """The workbook could not be replaced - usually because it is open in Excel."""


# ------------------------------------------------------------------ package helpers

def _rel_target(base_part, target):
    """Resolve a relationship Target relative to the part that owns the .rels file."""
    if target.startswith("/"):
        return target.lstrip("/")
    return normpath(pjoin(dirname(base_part), target))


def _sheet_parts(zf):
    """{sheet name: part path} read from workbook.xml (works even if Excel renumbers the sheets)."""
    wb = etree.fromstring(zf.read("xl/workbook.xml"))
    rels = etree.fromstring(zf.read("xl/_rels/workbook.xml.rels"))
    target = {r.get("Id"): r.get("Target") for r in rels}
    out = {}
    for s in wb.find(Q("sheets")):
        out[s.get("name")] = _rel_target("xl/workbook.xml", target[s.get(f"{{{RNS}}}id")])
    return out


def _table_part(zf, sheet_part):
    rels_name = pjoin(dirname(sheet_part), "_rels", os.path.basename(sheet_part) + ".rels")
    if rels_name not in zf.namelist():
        return None
    for r in etree.fromstring(zf.read(rels_name)):
        if r.get("Type", "").endswith("/table"):
            return _rel_target(sheet_part, r.get("Target"))
    return None


def _shared_strings(zf):
    if "xl/sharedStrings.xml" not in zf.namelist():
        return []
    root = etree.fromstring(zf.read("xl/sharedStrings.xml"))
    return ["".join(t.text or "" for t in si.iter(Q("t"))) for si in root.findall(Q("si"))]


# ------------------------------------------------------------------ reading

def _to_date(v):
    if isinstance(v, (int, float)):
        return EXCEL_EPOCH + datetime.timedelta(days=int(v))
    return datetime.date.fromisoformat(str(v)[:10])


def _parse_sheet(xml_stream, zf, source):
    """Yield one dict per data row of a worksheet (rows without a full A:Q record are skipped)."""
    shared = None
    rows = []
    for _, row in etree.iterparse(xml_stream, tag=Q("row")):
        if row.get("r") == "1":              # header
            row.clear()
            continue
        vals = {}
        for c in row.findall(Q("c")):
            m = _CELL_RE.match(c.get("r", ""))
            if not m or m.group(1) not in COL_LETTERS:
                continue
            v = c.find(Q("v"))
            t = c.get("t")
            if t == "inlineStr":
                txt = "".join(x.text or "" for x in c.iter(Q("t")))
            elif v is None or v.text is None:
                continue
            elif t == "s":
                if shared is None:
                    shared = _shared_strings(zf)
                txt = shared[int(v.text)]
            else:
                txt = v.text
            vals[COLUMNS[COL_LETTERS.index(m.group(1))]] = txt
        row.clear()
        if len(vals) < len(COLUMNS):
            continue
        try:
            rec = {}
            for k in COLUMNS:
                if k == "dteday":
                    rec[k] = _to_date(float(vals[k]) if re.fullmatch(r"-?\d+(\.\d+)?", vals[k]) else vals[k]).isoformat()
                elif k in INT_COLS:
                    rec[k] = int(float(vals[k]))
                else:
                    rec[k] = float(vals[k])
        except (ValueError, TypeError):
            continue
        rec["source"] = source
        rows.append(rec)
    return rows


def load_raw():
    """Raw_Data rows. Cached in memory until the Raw_Data sheet itself changes (checked by its CRC)."""
    with zipfile.ZipFile(EXCEL_PATH) as zf:
        part = _sheet_parts(zf)[RAW_SHEET]
        info = zf.getinfo(part)
        key = (info.CRC, info.file_size)
        if _raw_cache["key"] != key:
            with zf.open(part) as f:
                _raw_cache["rows"] = _parse_sheet(f, zf, "raw")
            _raw_cache["key"] = key
        return _raw_cache["rows"]


def load_web():
    """Web_Entries rows, read fresh from the workbook every time."""
    with zipfile.ZipFile(EXCEL_PATH) as zf:
        part = _sheet_parts(zf)[WEB_SHEET]
        with zf.open(part) as f:
            return _parse_sheet(f, zf, "web")


def all_rows():
    """Raw_Data + Web_Entries, sorted by date and hour."""
    rows = list(load_raw()) + load_web()
    rows.sort(key=lambda r: (r["dteday"], r["hr"], r["instant"]))
    return rows


def next_instant():
    """Website records continue the numbering of the original data (17380, 17381, ...), so an id
    means the same thing on the website, in Excel and in the .hyper."""
    ids = [r["instant"] for r in load_raw()] + [r["instant"] for r in load_web()]
    return (max(ids) if ids else 0) + 1


# ------------------------------------------------------------------ writing (Web_Entries only)

def _col_index(letters):
    n = 0
    for ch in letters:
        n = n * 26 + ord(ch) - 64
    return n


def _set_cell(row_el, col, rownum, value, default_style):
    ref = f"{col}{rownum}"
    cell = None
    for c in row_el.findall(Q("c")):
        if c.get("r") == ref:
            cell = c
            break
    if cell is None:                          # create it in the right column position
        cell = etree.Element(Q("c"), r=ref)
        if default_style is not None:
            cell.set("s", default_style)
        idx = _col_index(col)
        placed = False
        for i, c in enumerate(row_el.findall(Q("c"))):
            if _col_index(_CELL_RE.match(c.get("r")).group(1)) > idx:
                c.addprevious(cell)
                placed = True
                break
        if not placed:
            row_el.append(cell)
    for child in list(cell):
        cell.remove(child)
    cell.attrib.pop("t", None)
    if value is not None:
        v = etree.SubElement(cell, Q("v"))
        v.text = value


def _fmt(key, rec):
    if key == "dteday":
        return str((datetime.date.fromisoformat(rec[key][:10]) - EXCEL_EPOCH).days)
    if key in INT_COLS:
        return str(int(rec[key]))
    return repr(float(rec[key]))


def _rewrite_web_sheet(xml_bytes, records):
    """Return the Web_Entries sheet XML with data rows 2.. replaced by `records` (columns A:Q only;
    the developer notes in S:W and all formatting are kept)."""
    root = etree.fromstring(xml_bytes)
    sheet_data = root.find(Q("sheetData"))
    rows = {int(r.get("r")): r for r in sheet_data.findall(Q("row"))}

    def has_data(r):
        for c in r.findall(Q("c")):
            col = _CELL_RE.match(c.get("r")).group(1)
            if col in COL_LETTERS and c.find(Q("v")) is not None:
                return True
        return False

    old_last = max([n for n, r in rows.items() if n > 1 and has_data(r)], default=1)
    new_last = len(records) + 1
    styles = {}                                # style of each column in the first data row
    first = rows.get(2)
    if first is not None:
        for c in first.findall(Q("c")):
            styles[_CELL_RE.match(c.get("r")).group(1)] = c.get("s")

    for rownum in range(2, max(old_last, new_last) + 1):
        row_el = rows.get(rownum)
        if row_el is None:
            row_el = etree.Element(Q("row"), r=str(rownum))
            after = [n for n in rows if n < rownum]
            if after:
                rows[max(after)].addnext(row_el)
            else:
                sheet_data.append(row_el)
            rows[rownum] = row_el
        rec = records[rownum - 2] if rownum - 2 < len(records) else None
        for letter, key in zip(COL_LETTERS, COLUMNS):
            _set_cell(row_el, letter, rownum, _fmt(key, rec) if rec else None, styles.get(letter))

    dim = root.find(Q("dimension"))
    if dim is not None:
        m = re.match(r"^([A-Z]+\d+):([A-Z]+)(\d+)$", dim.get("ref", ""))
        if m and int(m.group(3)) < new_last:
            dim.set("ref", f"{m.group(1)}:{m.group(2)}{new_last}")
    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)


def _rewrite_table(xml_bytes, n_records):
    root = etree.fromstring(xml_bytes)
    root.set("ref", f"A1:Q{max(2, n_records + 1)}")      # an Excel table needs at least one data row
    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)


def write_web_entries(records):
    """Replace the content of Web_Entries with `records` (list of dicts with the 17 columns).
    Every other part of the workbook is copied unchanged."""
    records = sorted(records, key=lambda r: r["instant"])
    with write_lock:
        tmp = EXCEL_PATH + ".tmp"
        try:
            with zipfile.ZipFile(EXCEL_PATH) as zin:
                sheet_part = _sheet_parts(zin)[WEB_SHEET]
                table_part = _table_part(zin, sheet_part)
                with zipfile.ZipFile(tmp, "w") as zout:
                    for item in zin.infolist():
                        data = zin.read(item.filename)
                        if item.filename == sheet_part:
                            data = _rewrite_web_sheet(data, records)
                        elif item.filename == table_part:
                            data = _rewrite_table(data, len(records))
                        zout.writestr(item, data, compress_type=item.compress_type)
            try:
                os.replace(tmp, EXCEL_PATH)
            except PermissionError as e:
                raise ExcelLockedError(
                    "Bike_Sharing_Analysis_Final.xlsx is open in Excel. Close it and try again.") from e
        finally:
            if os.path.exists(tmp):
                try:
                    os.remove(tmp)
                except OSError:
                    pass


# ------------------------------------------------------------------ record operations

def find_duplicate(dteday, hr, ignore_instant=None):
    """The record already stored for this date + hour (in Raw_Data or Web_Entries), if any."""
    for r in list(load_raw()) + load_web():
        if r["dteday"] == dteday and r["hr"] == hr and r["instant"] != ignore_instant:
            return r
    return None


def add_entry(rec):
    with write_lock:
        web = load_web()
        rec = dict(rec, instant=next_instant())
        web.append(rec)
        write_web_entries(web)
        return rec["instant"]


def update_entry(instant, rec):
    with write_lock:
        web = load_web()
        for i, r in enumerate(web):
            if r["instant"] == instant:
                web[i] = dict(rec, instant=instant)
                write_web_entries(web)
                return True
        return False


def delete_entry(instant):
    with write_lock:
        web = load_web()
        kept = [r for r in web if r["instant"] != instant]
        if len(kept) == len(web):
            return False
        write_web_entries(kept)
        return True
