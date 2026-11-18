#!/usr/bin/env python3
"""
Omni Metric dashboard builder  (KPN Fresh)

Reads  base_1.csv   (daily rows per outlet and order source)
       Store_Mas.csv (store master: Store ID, Outlet_Name, Cluster_Manager, Finance_Code)
Keeps ONLY the stores listed in Store_Mas, injects the data into
dashboard_template.html and writes a self-contained, offline Metric_2026.html.

Usage (run inside the folder):
    python build_dashboard.py
    python build_dashboard.py base_1.csv Store_Mas.csv Metric_2026.html
    python build_dashboard.py --template dashboard_template.html

Defaults: base_1.csv, the first *store_mas*.csv found, Metric_2026.html.
Only the Python standard library is used.
"""
import argparse
import calendar
import csv
import datetime as dt
import json
import math
import re
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
PLACEHOLDER = "/*__DATA__*/null"
SHEET_URL = "https://docs.google.com/spreadsheets/d/1QaGKl1RRLgB4rHjbBz7Rfn7IKtXdfNBznpfjeMOSV44/edit?gid=0#gid=0"
LOGO_TOKEN = "__LOGO_SRC__"
LOGO_URL = ("https://play-lh.googleusercontent.com/tt3fN6DFHea7-902eGXuDO8O7abRWnMbHWUOzdMGe6sR58E1VXdJSD4XJjuZNjuisD5dlUGXgyURv-QDlt6ZHw=w480-h960-rw")
LOGO_MIME = {".png": "image/png", ".jpg": "image/jpeg", ".jpeg": "image/jpeg", ".webp": "image/webp",
             ".gif": "image/gif", ".svg": "image/svg+xml"}

# (canonical name, accepted header spellings). Order = column order inside each embedded row.
METRICS = [
    ("orders", ["orders"]),
    ("GMV", ["gmv"]),
    ("fnv_sale", ["fnv_sale"]),
    ("fruits_sale", ["fruits_sale"]),
    ("vegetables_sale", ["vegetables_sale"]),
    ("Milk_Sale", ["milk_sale"]),
    ("Fruits_Second_Sale", ["fruits_second_sale"]),
    ("Vegi_Second_Sale", ["vegi_second_sale"]),
    ("Salad_Sale", ["salad_sale"]),
    ("Loose_Provition_Sale", ["loose_provition_sale", "loose_provision_sale"]),
    ("OwnBrand_Sale", ["ownbrand_sale"]),
    ("Fmcg_Sale", ["fmcg_sale"]),
    ("Staples_Sale", ["staples_sale"]),
    ("icecream_Sale", ["icecream_sale"]),
    ("Rice Bag Qty", ["rice bag qty", "rice_bag_qty"]),
    ("Salad_QTY", ["salad_qty"]),
    ("Mob_Pent", ["mob_pent"]),
    ("EC_Used", ["ec_used"]),
    ("Fruit_Bills", ["fruit_bills", "fruit bills"]),
    ("Veg_Bills", ["veg_bills", "veg bills"]),
    ("PL_Bills", ["pl_bills", "pl bills"]),
]
SCAN_NAMES = ["scan_percentage", "scan_pct", "scan %"]
DATE_FORMATS = ["%d-%b-%y", "%d-%b-%Y", "%d/%m/%Y", "%d-%m-%Y", "%Y-%m-%d", "%d/%m/%y", "%d-%m-%y"]


def norm(s):
    return re.sub(r"[^a-z0-9]", "", str(s).lower())


def read_csv(path):
    """Read a CSV into a list of dicts keyed by normalised header names."""
    last_err = None
    for enc in ("utf-8-sig", "cp1252", "latin-1"):
        try:
            with open(path, newline="", encoding=enc) as f:
                reader = csv.reader(f)
                header = next(reader)
                keys = [norm(h) for h in header]
                rows = [dict(zip(keys, r)) for r in reader if any(c.strip() for c in r)]
            return header, rows
        except UnicodeDecodeError as e:
            last_err = e
    raise last_err


def pick(row_keys, names):
    for n in names:
        k = norm(n)
        if k in row_keys:
            return k
    return None


def num(v):
    """Parse numbers like '2,47,839.33', '85%', '' -> float (0.0 on junk)."""
    if v is None:
        return 0.0
    s = str(v).strip().replace(",", "").replace("%", "").replace("\u20b9", "")
    if s in ("", "-", "nan", "NaN", "None", "NULL", "null"):
        return 0.0
    try:
        f = float(s)
    except ValueError:
        return 0.0
    return f if math.isfinite(f) else 0.0


def sid(v):
    s = str(v).strip()
    if re.fullmatch(r"\d+(\.0+)?", s):
        s = str(int(float(s)))
    return s


_date_cache = {}


def parse_date(s):
    s = str(s).strip()
    if s in _date_cache:
        return _date_cache[s]
    out = None
    for fmt in DATE_FORMATS:
        try:
            out = dt.datetime.strptime(s, fmt).date()
            break
        except ValueError:
            continue
    _date_cache[s] = out
    return out


def compact(x):
    x = round(x, 2)
    return int(x) if x == int(x) else x


def download_sheet(url, tab, dest):
    """Download one tab of a Google Sheet as CSV (sheet must be shared 'Anyone with the link: Viewer')."""
    import urllib.error
    import urllib.parse
    import urllib.request
    m = re.search(r"/d/([A-Za-z0-9_-]+)", url)
    if not m:
        sys.exit("ERROR: could not find the sheet id in the URL.")
    q = urllib.parse.urlencode({"tqx": "out:csv", "sheet": tab})
    endpoint = f"https://docs.google.com/spreadsheets/d/{m.group(1)}/gviz/tq?{q}"
    req = urllib.request.Request(endpoint, headers={"User-Agent": "Mozilla/5.0"})
    try:
        with urllib.request.urlopen(req, timeout=180) as r:
            data = r.read()
    except (urllib.error.URLError, OSError) as e:
        sys.exit(f"ERROR: could not download tab '{tab}': {e}\n"
                 "Check your internet connection, or File > Download > CSV in Google Sheets and save it as base_1.csv.")
    head = data[:300].lower()
    if b"<html" in head or b"<!doctype" in head:
        sys.exit("ERROR: Google returned a sign-in page. Set the sheet to 'Anyone with the link: Viewer', "
                 "or download the tab as CSV and save it as base_1.csv.")
    dest.write_bytes(data)
    return len(data)


def logo_source(arg):
    """Local logo file -> embedded data URI (works offline). Otherwise the online URL."""
    import base64
    cands = []
    if arg and not arg.lower().startswith("http"):
        cands.append(Path(arg))
    elif not arg:
        for folder in (Path.cwd(), HERE):
            cands += sorted(p for p in folder.glob("logo.*") if p.suffix.lower() in LOGO_MIME)
    for c in cands:
        if c.exists():
            data = base64.b64encode(c.read_bytes()).decode("ascii")
            return f"data:{LOGO_MIME[c.suffix.lower()]};base64,{data}", f"embedded {c.name}"
    return (arg if arg else LOGO_URL), "online URL (save the logo as logo.png next to this script to embed it)"


def find_store_master(folder):
    hits = sorted(p for p in folder.glob("*.csv") if "store_mas" in p.name.lower())
    return hits[0] if hits else None


def main():
    ap = argparse.ArgumentParser(description="Build the Metric_2026 dashboard")
    ap.add_argument("base", nargs="?", default=None, help="daily data CSV (default base_1.csv)")
    ap.add_argument("stores", nargs="?", default=None, help="store master CSV (default first *store_mas*.csv)")
    ap.add_argument("out", nargs="?", default="Metric_2026.html", help="output HTML")
    ap.add_argument("--template", default=None, help="HTML template (default dashboard_template.html)")
    ap.add_argument("--sheet", action="store_true", help="download the base_1 tab from Google Sheets first (saves base_1.csv)")
    ap.add_argument("--sheet-url", default=SHEET_URL, help="Google Sheet link used with --sheet")
    ap.add_argument("--tab", default="base_1", help="tab name used with --sheet (default base_1)")
    ap.add_argument("--logo", default=None, help="logo image file or URL (default: logo.* in the folder, else the KPN URL)")
    a = ap.parse_args()

    base_p = Path(a.base) if a.base else (Path("base_1.csv") if Path("base_1.csv").exists() else HERE / "base_1.csv")
    if a.stores:
        sm_p = Path(a.stores)
    else:
        sm_p = find_store_master(Path.cwd()) or find_store_master(HERE)
    tpl_p = Path(a.template) if a.template else (
        Path("dashboard_template.html") if Path("dashboard_template.html").exists() else HERE / "dashboard_template.html")
    out_p = Path(a.out)

    if a.sheet:
        base_p = Path(a.base) if a.base else Path("base_1.csv")
        n = download_sheet(a.sheet_url, a.tab, base_p)
        print(f"Downloaded tab '{a.tab}' from Google Sheets -> {base_p} ({n/1024:.0f} KB)")

    for label, p in (("data file", base_p), ("store master", sm_p), ("template", tpl_p)):
        if p is None or not Path(p).exists():
            sys.exit(f"ERROR: {label} not found ({p}). Run this script inside the dashboard folder.")

    # ---- store master -------------------------------------------------------
    _, sm_rows = read_csv(sm_p)
    if not sm_rows:
        sys.exit("ERROR: Store master is empty.")
    keys = set(sm_rows[0].keys())
    k_id = pick(keys, ["Store ID", "Store_ID", "StoreID", "outlet_id"])
    k_name = pick(keys, ["Outlet_Name", "Outlet Name", "Store Name"])
    k_cl = pick(keys, ["Cluster_Manager", "Cluster Manager", "Cluster"])
    k_fin = pick(keys, ["Finance_Code", "Finance Code", "Fin Code"])
    missing = [n for n, k in (("Store ID", k_id), ("Outlet_Name", k_name), ("Cluster_Manager", k_cl)) if k is None]
    if missing:
        sys.exit(f"ERROR: Store master is missing column(s): {', '.join(missing)}")

    stores, store_ix = [], {}
    for r in sm_rows:
        i = sid(r[k_id])
        if not i or i in store_ix:
            continue
        store_ix[i] = len(stores)
        stores.append([i, r[k_name].strip(), r[k_cl].strip() or "Unassigned", (r.get(k_fin) or "").strip() if k_fin else ""])

    # ---- daily data ---------------------------------------------------------
    _, rows_in = read_csv(base_p)
    if not rows_in:
        sys.exit("ERROR: base data is empty.")
    keys = set(rows_in[0].keys())
    k_out = pick(keys, ["outlet_id", "Store ID"])
    k_src = pick(keys, ["Order_source", "Order source"])
    k_dt = pick(keys, ["dt", "date"])
    if not (k_out and k_src and k_dt):
        sys.exit("ERROR: base data needs outlet_id, Order_source and dt columns.")
    mcols = []
    for name, alts in METRICS:
        k = pick(keys, alts)
        if k is None:
            print(f"WARN: column '{name}' not found - filled with 0")
        mcols.append(k)
    k_scan = pick(keys, SCAN_NAMES)
    if k_scan is None:
        print("WARN: Scan_Percentage not found - Scan % will be 0")

    out_rows, bad_dt, bad_src, skipped_store = [], 0, 0, 0
    unknown_ids = set()
    raw_tot = [0.0] * len(METRICS)
    week_days, month_last = {}, {}
    last_date = None
    seen_ids = set()

    for r in rows_in:
        d = parse_date(r.get(k_dt, ""))
        if d is None:
            bad_dt += 1
            continue
        src_t = r.get(k_src, "").strip().upper()
        if src_t not in ("ECOM", "POS"):
            bad_src += 1
            continue
        o = sid(r.get(k_out, ""))
        if o not in store_ix:
            skipped_store += 1
            unknown_ids.add(o)
            continue
        seen_ids.add(o)
        iy, iw, _ = d.isocalendar()
        ym, wk = d.year * 100 + d.month, iy * 100 + iw
        vals = [num(r.get(k)) if k else 0.0 for k in mcols]
        scan = num(r.get(k_scan)) if k_scan else 0.0
        for i, v in enumerate(vals):
            raw_tot[i] += v
        out_rows.append([store_ix[o], ym, wk, d.day, 0 if src_t == "ECOM" else 1] + vals + [scan])
        week_days.setdefault(wk, set()).add(d)
        month_last[ym] = max(month_last.get(ym, 0), d.day)
        if last_date is None or d > last_date:
            last_date = d

    if not out_rows:
        sys.exit("ERROR: no usable rows after filtering. Check Store ID matching and the dt format.")

    # Scan % may arrive as a fraction (0.87) or a percent (87); store as percent.
    scan_max = max(r[-1] for r in out_rows)
    if 0 < scan_max <= 1.5:
        for r in out_rows:
            r[-1] *= 100

    for r in out_rows:
        for j in range(5, len(r)):
            r[j] = compact(r[j])

    months = sorted(month_last)
    data = {
        "stores": stores,
        "months": months,
        "monthLast": month_last,
        "monthDays": {ym: calendar.monthrange(ym // 100, ym % 100)[1] for ym in months},
        "weekDays": {wk: len(v) for wk, v in week_days.items()},
        "lastDate": last_date.isoformat(),
        "rows": out_rows,
    }
    blob = json.dumps(data, separators=(",", ":")).replace("</", "<\\/")

    tpl = tpl_p.read_text(encoding="utf-8")
    if PLACEHOLDER not in tpl:
        sys.exit(f"ERROR: placeholder {PLACEHOLDER} not found in the template.")
    logo_src, logo_note = logo_source(a.logo)
    html = tpl.replace(PLACEHOLDER, blob, 1).replace(LOGO_TOKEN, logo_src)
    out_p.write_text(html, encoding="utf-8")

    # ---- reconciliation -----------------------------------------------------
    emb_tot = [sum(r[5 + i] for r in out_rows) for i in range(len(METRICS))]
    worst = max((abs(e - t) for e, t in zip(emb_tot, raw_tot)), default=0)
    print(f"Wrote {out_p}  ({out_p.stat().st_size/1024:.0f} KB)")
    print(f"  logo: {logo_note}")
    print(f"  stores in master: {len(stores)}   with data: {len(seen_ids)}   rows kept: {len(out_rows):,}")
    print(f"  data through {last_date:%d-%b-%Y}   months: {len(months)}   weeks: {len(week_days)}")
    print(f"  reconcile  orders={emb_tot[0]:,.0f}  GMV={emb_tot[1]:,.2f}  (max drift vs CSV {worst:.2f})")
    if bad_dt:
        print(f"WARN: skipped {bad_dt} row(s) with an unreadable dt")
    if bad_src:
        print(f"WARN: skipped {bad_src} row(s) with Order_source other than ECOM/POS")
    if skipped_store:
        print(f"INFO: dropped {skipped_store} row(s) for {len(unknown_ids)} outlet(s) not in Store_Mas")
    idle = [s[0] for s in stores if s[0] not in seen_ids]
    if idle:
        print(f"WARN: {len(idle)} master store(s) have no data: {', '.join(idle[:10])}")


if __name__ == "__main__":
    main()
