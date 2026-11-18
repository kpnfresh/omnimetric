#!/usr/bin/env python3
"""
Omni Metric - one-file builder + publisher  (KPN Fresh)

What it does, in order:
  1. Downloads the base_1 tab from the Google Sheet  (A:AB, all columns, read by header name)
  2. Reads the store master (local Store_Mas.csv, else the "Store_Mas" tab of the same sheet)
  3. Builds the self-contained dashboard (template is embedded below) -> Omni_Metric.html
  4. Commits index.html and pushes it to GitHub  (kpnfresh/omnimetric, branch main)

Run:
    python omni_metric.py                 # download + build + push
    python omni_metric.py --no-push       # download + build only
    python omni_metric.py --csv base_1.csv --stores Store_Mas.csv --no-push    # use local files
    python omni_metric.py --force         # push even if checks (bills columns, stale data) fail

Needs: Python 3.8+ and git (already signed in to GitHub). No other packages.
The Google Sheet must be shared as "Anyone with the link: Viewer".
"""
import argparse
import base64
import calendar
import csv
import datetime as dt
import io
import json
import math
import re
import shutil
import subprocess
import sys
import urllib.error
import urllib.parse
import urllib.request
from pathlib import Path

# ----------------------------------------------------------------------------
# SETTINGS - change here only
# ----------------------------------------------------------------------------
SHEET_URL = "https://docs.google.com/spreadsheets/d/1QaGKl1RRLgB4rHjbBz7Rfn7IKtXdfNBznpfjeMOSV44/edit?gid=0#gid=0"
BASE_TAB = "base_1"
STORE_TAB = "Store_Mas"
REPO_URL = "https://github.com/kpnfresh/omnimetric.git"
BRANCH = "main"
PUBLISH_NAME = "index.html"
LOCAL_OUT = "Omni_Metric.html"
LOGO_URL = ("https://play-lh.googleusercontent.com/tt3fN6DFHea7-902eGXuDO8O7abRWnMbHWUOzdMGe6sR58E1VXdJSD4XJjuZNjuisD5dlUGXgyURv-QDlt6ZHw=w480-h960-rw")
MAX_STALE_DAYS = 3          # warn / block push if the newest data row is older than this many days

HERE = Path(__file__).resolve().parent
PLACEHOLDER = "/*__DATA__*/null"
LOGO_TOKEN = "__LOGO_SRC__"
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
BILLS = ("Fruit_Bills", "Veg_Bills", "PL_Bills")
SCAN_NAMES = ["scan_percentage", "scan_pct", "scan %"]
DATE_FORMATS = ["%d-%b-%y", "%d-%b-%Y", "%d/%m/%Y", "%d-%m-%Y", "%Y-%m-%d", "%d/%m/%y", "%d-%m-%y"]


# ----------------------------------------------------------------------------
# small helpers
# ----------------------------------------------------------------------------
def say(msg=""):
    print(msg, flush=True)


def die(msg):
    print("\nERROR: " + msg, file=sys.stderr, flush=True)
    sys.exit(1)


def norm(s):
    return re.sub(r"[^a-z0-9]", "", str(s).lower())


def pick(keys, names):
    for n in names:
        k = norm(n)
        if k in keys:
            return k
    return None


def num(v):
    """'2,47,839.33', '85%', '' -> float (0.0 on junk)."""
    if v is None:
        return 0.0
    s = str(v).strip().replace(",", "").replace("%", "").replace("₹", "")
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


def parse_csv_text(text):
    """CSV text -> (header, [dict keyed by normalised header])."""
    reader = csv.reader(io.StringIO(text))
    try:
        header = next(reader)
    except StopIteration:
        return [], []
    keys = [norm(h) for h in header]
    rows = [dict(zip(keys, r)) for r in reader if any(c.strip() for c in r)]
    return header, rows


def read_local_csv(path):
    raw = Path(path).read_bytes()
    for enc in ("utf-8-sig", "cp1252", "latin-1"):
        try:
            return raw.decode(enc)
        except UnicodeDecodeError:
            continue
    return raw.decode("latin-1", errors="replace")


# ----------------------------------------------------------------------------
# Google Sheets download
# ----------------------------------------------------------------------------
def http_get(url, timeout=180):
    req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
    with urllib.request.urlopen(req, timeout=timeout) as r:
        return r.read(), r.headers.get("Content-Type", "")


def fetch_tab(sheet_url, tab, need, gid=None):
    """Download one tab as CSV text. `need` = normalised header names that must be present.

    Tries the raw export (by gid) first - it keeps every cell exactly as shown - then the
    tab-name endpoint. Returns the first response whose header has all the `need` columns.
    """
    m = re.search(r"/d/([A-Za-z0-9_-]+)", sheet_url)
    if not m:
        die("Could not find the sheet id in SHEET_URL.")
    sheet_id = m.group(1)
    if gid is None:
        g = re.search(r"[#&?]gid=(\d+)", sheet_url)
        gid = g.group(1) if g else None
    tries = []
    if gid is not None:
        tries.append((f"gid {gid}", f"https://docs.google.com/spreadsheets/d/{sheet_id}/export?format=csv&gid={gid}"))
    tries.append((f"tab name '{tab}'", f"https://docs.google.com/spreadsheets/d/{sheet_id}/gviz/tq?"
                  f"tqx=out:csv&sheet={urllib.parse.quote(tab)}"))
    notes = []
    for label, url in tries:
        try:
            data, _ = http_get(url)
        except (urllib.error.URLError, OSError) as e:
            notes.append(f"{label}: {e}")
            continue
        head = data[:400].lower()
        if b"<html" in head or b"<!doctype" in head:
            notes.append(f"{label}: Google asked for sign-in (share the sheet as 'Anyone with the link: Viewer')")
            continue
        text = data.decode("utf-8-sig", errors="replace")
        header, _rows = parse_csv_text(text.split("\n", 1)[0] + "\n")
        have = {norm(h) for h in header}
        if all(any(norm(n) in have for n in group) for group in need):
            say(f"  downloaded '{tab}' via {label} ({len(data)/1024:,.0f} KB)")
            return text
        notes.append(f"{label}: this is not the '{tab}' tab (columns {', '.join(header[:6])}...)")
    die(f"Could not download tab '{tab}' from Google Sheets.\n  - " + "\n  - ".join(notes) +
        "\n  Fix the sharing, or save the tab as a CSV and run:  python omni_metric.py --csv <file> --no-push")


# ----------------------------------------------------------------------------
# Logo
# ----------------------------------------------------------------------------
def logo_source(arg):
    """Local logo file -> data URI; else download LOGO_URL and embed it; else link to the URL."""
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
    url = arg if arg else LOGO_URL
    try:
        data, ctype = http_get(url, timeout=25)
        ctype = (ctype or "image/png").split(";")[0].strip()
        if ctype.startswith("image/") and 200 < len(data) < 2_000_000:
            return f"data:{ctype};base64,{base64.b64encode(data).decode('ascii')}", "downloaded and embedded"
    except (urllib.error.URLError, OSError):
        pass
    return url, "linked online (could not embed; save a logo.png next to the script to embed it)"


# ----------------------------------------------------------------------------
# Build
# ----------------------------------------------------------------------------
def load_stores(stores_arg):
    path = Path(stores_arg) if stores_arg else None
    if path is None:
        for folder in (Path.cwd(), HERE):
            hits = sorted(p for p in folder.glob("*.csv") if "store_mas" in p.name.lower())
            if hits:
                path = hits[0]
                break
    if path is not None:
        if not path.exists():
            die(f"Store master not found: {path}")
        say(f"  store master: {path.name} (local file)")
        text = read_local_csv(path)
    else:
        say(f"  no local Store_Mas.csv - using the '{STORE_TAB}' tab of the sheet")
        text = fetch_tab(SHEET_URL, STORE_TAB, need=[["store id", "store_id", "storeid", "outlet_id"],
                                                     ["outlet_name", "outlet name", "store name"]])
    _, rows = parse_csv_text(text)
    if not rows:
        die("Store master is empty.")
    keys = set(rows[0].keys())
    k_id = pick(keys, ["Store ID", "Store_ID", "StoreID", "outlet_id"])
    k_name = pick(keys, ["Outlet_Name", "Outlet Name", "Store Name"])
    k_cl = pick(keys, ["Cluster_Manager", "Cluster Manager", "Cluster"])
    k_fin = pick(keys, ["Finance_Code", "Finance Code", "Fin Code"])
    missing = [n for n, k in (("Store ID", k_id), ("Outlet_Name", k_name), ("Cluster_Manager", k_cl)) if k is None]
    if missing:
        die("Store master is missing column(s): " + ", ".join(missing))
    stores, ix = [], {}
    for r in rows:
        i = sid(r[k_id])
        if not i or i in ix:
            continue
        ix[i] = len(stores)
        stores.append([i, r[k_name].strip(), r[k_cl].strip() or "Unassigned",
                       (r.get(k_fin) or "").strip() if k_fin else ""])
    return stores, ix


def build_data(base_text, stores, store_ix):
    _, rows_in = parse_csv_text(base_text)
    if not rows_in:
        die("base_1 is empty.")
    keys = set(rows_in[0].keys())
    k_out = pick(keys, ["outlet_id", "Store ID"])
    k_src = pick(keys, ["Order_source", "Order source"])
    k_dt = pick(keys, ["dt", "date"])
    if not (k_out and k_src and k_dt):
        die("base_1 needs outlet_id, Order_source and dt columns.")
    mcols, missing_cols = [], []
    for name, alts in METRICS:
        k = pick(keys, alts)
        if k is None:
            missing_cols.append(name)
        mcols.append(k)
    k_scan = pick(keys, SCAN_NAMES)
    if k_scan is None:
        missing_cols.append("Scan_Percentage")

    out_rows, bad_dt, bad_src, skipped = [], 0, 0, 0
    unknown, seen = set(), set()
    raw_tot = [0.0] * len(METRICS)
    week_days, month_last = {}, {}
    last_date = None
    for r in rows_in:
        d = parse_date(r.get(k_dt, ""))
        if d is None:
            bad_dt += 1
            continue
        src = r.get(k_src, "").strip().upper()
        if src not in ("ECOM", "POS"):
            bad_src += 1
            continue
        o = sid(r.get(k_out, ""))
        if o not in store_ix:
            skipped += 1
            unknown.add(o)
            continue
        seen.add(o)
        iy, iw, _ = d.isocalendar()
        ym, wk = d.year * 100 + d.month, iy * 100 + iw
        vals = [num(r.get(k)) if k else 0.0 for k in mcols]
        scan = num(r.get(k_scan)) if k_scan else 0.0
        for i, v in enumerate(vals):
            raw_tot[i] += v
        out_rows.append([store_ix[o], ym, wk, d.day, 0 if src == "ECOM" else 1] + vals + [scan])
        week_days.setdefault(wk, set()).add(d)
        month_last[ym] = max(month_last.get(ym, 0), d.day)
        if last_date is None or d > last_date:
            last_date = d
    if not out_rows:
        die("No usable rows after filtering. Check Store ID matching and the dt format.")

    # Scan % may arrive as a fraction (0.87) or a percent (87); store as percent.
    if 0 < max(r[-1] for r in out_rows) <= 1.5:
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
    emb = [sum(r[5 + i] for r in out_rows) for i in range(len(METRICS))]
    info = {
        "rows": len(out_rows), "last_date": last_date, "months": len(months), "weeks": len(week_days),
        "orders": emb[0], "gmv": emb[1], "drift": max(abs(e - t) for e, t in zip(emb, raw_tot)),
        "bad_dt": bad_dt, "bad_src": bad_src, "skipped": skipped, "unknown": unknown,
        "idle": [s[0] for s in stores if s[0] not in seen], "missing_cols": missing_cols,
        "bill_totals": {b: emb[[m[0] for m in METRICS].index(b)] for b in BILLS},
    }
    return data, info


def render(data, logo_src):
    blob = json.dumps(data, separators=(",", ":")).replace("</", "<\\/")
    if PLACEHOLDER not in TEMPLATE:
        die("Template placeholder missing (script was edited incorrectly).")
    return TEMPLATE.replace(PLACEHOLDER, blob, 1).replace(LOGO_TOKEN, logo_src)


# ----------------------------------------------------------------------------
# GitHub publish (uses your own git login)
# ----------------------------------------------------------------------------
def git(args, cwd, check=True):
    p = subprocess.run(["git"] + args, cwd=str(cwd), capture_output=True, text=True)
    if check and p.returncode != 0:
        die(f"git {' '.join(args)} failed:\n{(p.stderr or p.stdout).strip()}")
    return p


def publish(html, repo_url, branch, message):
    if shutil.which("git") is None:
        die("git is not installed or not on PATH.")
    repo = HERE if (HERE / ".git").exists() else HERE / "omnimetric_repo"
    if not (repo / ".git").exists():
        say(f"  cloning {repo_url} -> {repo.name}/")
        git(["clone", repo_url, str(repo)], HERE)
    git(["checkout", branch], repo, check=False)
    pull = git(["pull", "--rebase", "--autostash", "origin", branch], repo, check=False)
    if pull.returncode != 0 and "couldn't find remote ref" not in pull.stderr.lower():
        die("git pull failed:\n" + (pull.stderr or pull.stdout).strip())
    (repo / PUBLISH_NAME).write_text(html, encoding="utf-8")
    git(["add", PUBLISH_NAME], repo)
    if not git(["status", "--porcelain", "--", PUBLISH_NAME], repo).stdout.strip():
        say("  index.html is already up to date on GitHub - nothing to push.")
        return
    git(["commit", "-m", message], repo)
    push = git(["push", "origin", f"HEAD:{branch}"], repo, check=False)
    if push.returncode != 0:
        git(["pull", "--rebase", "--autostash", "origin", branch], repo)
        git(["push", "origin", f"HEAD:{branch}"], repo)
    say(f"  pushed {PUBLISH_NAME} to {repo_url.replace('.git', '')} ({branch})")


# ----------------------------------------------------------------------------
def main():
    try:
        sys.stdout.reconfigure(errors="replace")
        sys.stderr.reconfigure(errors="replace")
    except Exception:
        pass
    ap = argparse.ArgumentParser(description="Download base_1 from Google Sheets, build Omni Metric, push to GitHub")
    ap.add_argument("--csv", help="use this local base_1 CSV instead of downloading the sheet")
    ap.add_argument("--stores", help="local store master CSV (default: Store_Mas.csv next to the script, else the sheet tab)")
    ap.add_argument("--no-push", action="store_true", help="build only, do not touch GitHub")
    ap.add_argument("--force", action="store_true", help="push even if the safety checks fail")
    ap.add_argument("--logo", help="logo file or URL (default: logo.* next to the script, else the KPN logo URL)")
    ap.add_argument("--repo", default=REPO_URL, help="GitHub repo URL to push to")
    ap.add_argument("--branch", default=BRANCH)
    a = ap.parse_args()

    say("Omni Metric")
    say("1/4  Data")
    if a.csv:
        if not Path(a.csv).exists():
            die(f"CSV not found: {a.csv}")
        say(f"  base data: {a.csv} (local file)")
        base_text = read_local_csv(a.csv)
    else:
        base_text = fetch_tab(SHEET_URL, BASE_TAB, need=[["outlet_id"], ["dt", "date"], ["orders"], ["order_source"]])
        (HERE / "base_1.csv").write_text(base_text, encoding="utf-8")
    stores, ix = load_stores(a.stores)

    say("2/4  Build")
    data, info = build_data(base_text, stores, ix)
    logo_src, logo_note = logo_source(a.logo)
    html = render(data, logo_src)
    out = HERE / LOCAL_OUT
    out.write_text(html, encoding="utf-8")
    say(f"  wrote {out.name} ({out.stat().st_size/1024:,.0f} KB)   logo: {logo_note}")
    say(f"  stores in master: {len(stores)}   with data: {len(stores) - len(info['idle'])}   rows: {info['rows']:,}")
    say(f"  data through {info['last_date']:%d-%b-%Y}   months: {info['months']}   weeks: {info['weeks']}")
    say(f"  reconcile  orders={info['orders']:,.0f}  GMV={info['gmv']:,.2f}  (max drift vs CSV {info['drift']:.2f})")
    say("  bills totals  " + "   ".join(f"{k}={v:,.0f}" for k, v in info["bill_totals"].items()))

    problems = []
    if info["missing_cols"]:
        problems.append("columns not found in the sheet: " + ", ".join(info["missing_cols"]))
    elif all(v == 0 for v in info["bill_totals"].values()):
        problems.append("Fruit/Veg/PL bills are all zero")
    age = (dt.date.today() - info["last_date"]).days
    if age > MAX_STALE_DAYS:
        problems.append(f"newest data is {age} days old ({info['last_date']:%d-%b-%Y})")
    if info["bad_dt"]:
        problems.append(f"{info['bad_dt']} row(s) skipped - unreadable dt")
    if info["bad_src"]:
        say(f"  note: {info['bad_src']} row(s) skipped - Order_source not ECOM/POS")
    if info["skipped"]:
        say(f"  note: {info['skipped']} row(s) for {len(info['unknown'])} outlet(s) not in the store master were dropped")
    if info["idle"]:
        say(f"  note: {len(info['idle'])} master store(s) have no data: {', '.join(info['idle'][:8])}")
    for p in problems:
        say(f"  WARNING: {p}")

    say("3/4  Checks")
    blocking = [p for p in problems if not p.endswith("unreadable dt")]
    if a.no_push:
        say("  --no-push: skipping GitHub.")
        say(f"4/4  Done. Open {out}")
        return
    if blocking and not a.force:
        die("Not pushing because of the warnings above. Fix the data, or re-run with --force to push anyway.")
    say("  ok" if not problems else "  continuing (--force or minor warnings)")

    say("4/4  Push to GitHub")
    publish(html, a.repo, a.branch, f"Update Omni Metric data through {info['last_date']:%d-%b-%Y}")
    say("Done.")


# ----------------------------------------------------------------------------
# Dashboard template (HTML/CSS/JS) - edit the dashboard look and logic here
# ----------------------------------------------------------------------------
TEMPLATE = r'''<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Omni Metric | KPN Fresh</title>
<link rel="icon" type="image/png" href="https://play-lh.googleusercontent.com/tt3fN6DFHea7-902eGXuDO8O7abRWnMbHWUOzdMGe6sR58E1VXdJSD4XJjuZNjuisD5dlUGXgyURv-QDlt6ZHw=w64-h64">
<link rel="apple-touch-icon" href="https://play-lh.googleusercontent.com/tt3fN6DFHea7-902eGXuDO8O7abRWnMbHWUOzdMGe6sR58E1VXdJSD4XJjuZNjuisD5dlUGXgyURv-QDlt6ZHw=w180-h180">
<style>
:root{
  --brand:#08A977; --brand-d:#067A56; --brand-deep:#0A6B3F; --brand-soft:#E8F7F1; --brand-line:#BFE6D5;
  --ink:#1E2226; --ink-2:#3A4047; --muted:#6B7280; --faint:#9AA1AB;
  --line:#E4E7EB; --line-2:#EEF0F3; --panel:#F6F7F8; --page:#EFF0F2; --white:#fff;
  --up:#0A8F5F; --up-bg:#E1F5EC; --down:#D23B36; --down-bg:#FCE8E7;
  --flat:#6B7280; --flat-bg:#EEF0F3; --amber:#8A5A00; --amber-bg:#FFF2CC;
  --radius:10px; --shadow:0 1px 2px rgba(20,28,36,.05),0 4px 14px -8px rgba(20,28,36,.12);
  --tb:76px; --off:76px;
}
*{box-sizing:border-box}
html{-webkit-text-size-adjust:100%}
body{margin:0;background:var(--page);color:var(--ink);
  font:14px/1.45 "Segoe UI Variable Text","Segoe UI",Inter,system-ui,-apple-system,Roboto,"Helvetica Neue",Arial,sans-serif;
  font-variant-numeric:tabular-nums;-webkit-font-smoothing:antialiased}
button,select,input{font:inherit;color:inherit}
:focus-visible{outline:2px solid var(--brand);outline-offset:2px;border-radius:6px}
[hidden]{display:none!important}

/* Top bar */
.topbar{position:sticky;top:0;z-index:60;background:#fff;border-top:4px solid var(--brand-deep);border-bottom:1px solid var(--line);
  min-height:var(--tb);padding:10px clamp(14px,2vw,28px);display:flex;align-items:center;justify-content:space-between;gap:14px;flex-wrap:wrap}
.tb-l{display:flex;align-items:center;gap:16px;min-width:0}
.logo-img{height:48px;width:auto;max-width:132px;object-fit:contain;border-radius:6px;flex:none;display:block}
.logo{width:50px;height:50px;border-radius:6px;background:#FFE500;display:flex;flex-direction:column;align-items:center;justify-content:center;
  line-height:1;flex:none;user-select:none}
.logo b{font-weight:900;font-size:15px;letter-spacing:-.03em;color:#1D1D1B}
.logo i{font-weight:800;font-size:13px;color:#E8471C;letter-spacing:-.04em;margin-top:1px}
.topbar h1{margin:0;font-size:clamp(22px,2.4vw,28px);font-weight:600;letter-spacing:-.01em;display:flex;align-items:center;gap:10px}
.pin{border:0;background:none;padding:5px;border-radius:6px;cursor:pointer;color:var(--faint);display:inline-grid;place-items:center;transition:color .15s,background .15s}
.pin svg{transform:rotate(35deg);transition:transform .2s,fill .2s}
.pin:hover{color:var(--brand);background:var(--brand-soft)}
.pin[aria-pressed="true"]{color:var(--brand)}
.pin[aria-pressed="true"] svg{transform:rotate(0);fill:rgba(8,169,119,.28)}
.hot{display:none}
.hot-tab{position:absolute;left:50%;top:0;transform:translateX(-50%);height:18px;padding:0 14px;background:var(--brand-deep);color:#fff;
  border-radius:0 0 9px 9px;display:flex;align-items:center;gap:6px;font-size:11px;font-weight:600;cursor:pointer}
@media (min-width:901px){
  .topbar{position:fixed;top:0;left:0;right:0;transform:translateY(-101%);transition:transform .28s cubic-bezier(.2,.7,.2,1)}
  body.tb-pinned .topbar{position:sticky;transform:none}
  body.tb-peek .topbar{transform:none;box-shadow:0 12px 30px -12px rgba(20,28,36,.35)}
  body:not(.tb-pinned){--off:0px}
  .hot{display:block;position:fixed;top:0;left:0;right:0;height:18px;z-index:59}
  .hot::before{content:"";position:absolute;left:0;right:0;top:0;height:4px;background:var(--brand-deep)}
  body.tb-pinned .hot,body.tb-peek .hot{display:none}
}
@media (max-width:900px){.pin{display:none}}
.chips{display:flex;gap:8px;flex-wrap:wrap}
.chip{padding:6px 12px;border-radius:999px;background:var(--panel);border:1px solid var(--line);font-size:13px;color:var(--muted)}
.chip b{color:var(--ink);font-weight:650}
.chip.on{background:var(--brand-soft);border-color:var(--brand-line);color:var(--brand-d)}
.chip.on b{color:var(--brand-d)}

/* Layout */
.wrap{max-width:1500px;margin:0 auto;padding:18px clamp(14px,2vw,28px) 40px}
.card{background:#fff;border:1px solid var(--line);border-radius:var(--radius);box-shadow:var(--shadow)}
.card-h,.panel-h{display:flex;align-items:center;justify-content:space-between;gap:12px;padding:14px 20px;background:var(--panel);
  border-bottom:1px solid var(--line);border-radius:var(--radius) var(--radius) 0 0}
.card-h h2,.panel-h h2{margin:0;font-size:18px;font-weight:600;letter-spacing:-.005em}
.panel-h span{color:var(--muted);font-size:13px}
.linkbtn{border:0;background:none;color:var(--brand);font-weight:600;cursor:pointer;padding:4px 6px;border-radius:6px}
.linkbtn:hover{color:var(--brand-d);text-decoration:underline}

/* Filters */
.filters{padding:20px;display:flex;flex-wrap:wrap;gap:18px 26px;align-items:flex-end}
.fg{display:flex;flex-direction:column;gap:8px;min-width:0}
.fg>label,.fg>span.lbl{font-size:14px;font-weight:600;color:var(--ink)}
.seg{display:inline-flex;background:#F1F2F4;border:1px solid var(--line);border-radius:8px;padding:3px;gap:2px;height:40px}
.seg button{border:0;background:transparent;padding:0 14px;border-radius:6px;cursor:pointer;font-weight:600;color:var(--ink-2);
  transition:background .15s,color .15s}
.seg button:hover{background:#E6E8EB}
.seg button[aria-pressed="true"]{background:var(--brand);color:#fff;box-shadow:0 1px 3px rgba(6,122,86,.45)}
.sel,.ms-btn{appearance:none;-webkit-appearance:none;border:1px solid #D7DBE0;background:#fff;border-radius:6px;
  padding:8px 34px 8px 12px;cursor:pointer;font-weight:500;min-height:40px;
  background-image:url("data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' width='12' height='12' viewBox='0 0 12 12'%3E%3Cpath d='M2.5 4.5 6 8l3.5-3.5' fill='none' stroke='%236B7280' stroke-width='1.6' stroke-linecap='round' stroke-linejoin='round'/%3E%3C/svg%3E");
  background-repeat:no-repeat;background-position:right 11px center}
.sel{text-align:center;text-align-last:center}
.sel:hover,.ms-btn:hover{border-color:#B6BDC6}
.sel.small{min-width:76px}
.btn{display:inline-flex;align-items:center;gap:8px;border:1px solid #D7DBE0;background:#fff;border-radius:6px;padding:0 16px;font-weight:600;cursor:pointer;min-height:40px;
  transition:background .15s,border-color .15s}
.btn:hover{background:var(--panel)}
.btn.brand{background:var(--brand);border-color:var(--brand);color:#fff;box-shadow:0 1px 2px rgba(6,122,86,.35)}
.btn.brand:hover{background:var(--brand-d);border-color:var(--brand-d)}

/* Multi-select */
.ms{position:relative}
.ms-btn{min-width:190px;max-width:260px;text-align:left;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;display:block;width:100%}
.ms-panel{position:absolute;z-index:30;top:calc(100% + 6px);left:0;width:290px;background:#fff;border:1px solid var(--line);
  border-radius:10px;box-shadow:0 18px 36px -12px rgba(20,28,36,.3);padding:10px}
.ms-search{width:100%;border:1px solid #D7DBE0;border-radius:6px;padding:8px 10px;background:var(--panel)}
.ms-search:focus{background:#fff}
.ms-actions{display:flex;justify-content:space-between;margin:8px 2px 4px}
.ms-actions button{border:0;background:none;color:var(--brand);font-weight:600;cursor:pointer;padding:2px 4px;border-radius:6px}
.ms-list{list-style:none;margin:0;padding:0;max-height:260px;overflow:auto}
.ms-list li label{display:flex;gap:10px;align-items:center;padding:7px 8px;border-radius:6px;cursor:pointer}
.ms-list li label:hover{background:var(--panel)}
.ms-list input{accent-color:var(--brand);width:16px;height:16px;margin:0}
.ms-list small{color:var(--faint);margin-left:auto;padding-left:8px;white-space:nowrap}
.ms-empty{padding:12px 8px;color:var(--muted)}

/* KPI */
.kpis{display:grid;grid-template-columns:repeat(4,minmax(0,1fr));gap:14px;margin:16px 0}
.kpi{position:relative;overflow:hidden;padding:18px 20px;display:flex;flex-direction:column;gap:6px;min-height:150px}
.kpi .k-label{font-size:15px;font-weight:700;color:var(--ink-2)}
.kpi .k-cap{font-size:12px;color:var(--faint);margin-top:-4px}
.kpi .k-val{font-size:clamp(24px,2.4vw,32px);font-weight:700;letter-spacing:-.02em;margin-top:6px}
.kpi .k-pills{display:flex;flex-wrap:wrap;gap:6px 12px;margin-top:auto;min-height:26px}
.kpi .k-pill{display:inline-flex;align-items:center;gap:6px;font-size:12px;color:var(--muted);font-weight:600}
.kpi .k-bar{height:4px;border-radius:99px;background:var(--line-2);margin-top:2px}
.kpi .k-bar i{display:block;height:100%;border-radius:99px;background:#C3CAD3}
.kpi .badge{margin-left:8px;font-size:11px;font-weight:700;color:var(--brand-d);background:#fff;border:1px solid var(--brand-line);
  padding:1px 8px;border-radius:999px;vertical-align:2px}
.kpi.lead{background:linear-gradient(165deg,#F7FCFA 0%,var(--brand-soft) 100%);border-color:var(--brand-line)}
.kpi.lead::before{content:"";position:absolute;left:0;right:0;top:0;height:3px;background:linear-gradient(90deg,var(--brand),#4FCFA2)}
.kpi.lead .k-label{color:var(--brand-d)}
.kpi.lead .k-bar i{background:linear-gradient(90deg,var(--brand),#4FCFA2)}
.hint{margin:14px 0 0;padding:10px 14px;border-radius:8px;background:var(--amber-bg);color:var(--amber);font-weight:600;font-size:13px}

/* Pills */
.pill{display:inline-flex;align-items:center;gap:4px;padding:3px 9px;border-radius:999px;font-weight:700;font-size:12.5px;line-height:1.3;white-space:nowrap}
.pill i{font-style:normal;font-size:9px}
.pill.up{background:var(--up-bg);color:var(--up)}
.pill.down{background:var(--down-bg);color:var(--down)}
.pill.flat{background:var(--flat-bg);color:var(--flat)}

/* Tables */
.panel{margin-bottom:18px;overflow:hidden;overflow:clip}
.tbl-wrap{overflow:visible}
table{border-collapse:separate;border-spacing:0;width:100%}
th,td{padding:12px 8px;border-bottom:1px solid var(--line-2);text-align:center}
td{white-space:nowrap}
th+th,td+td{border-left:1px solid var(--line-2)}
thead th{position:sticky;top:var(--off);transition:top .28s cubic-bezier(.2,.7,.2,1);z-index:3;white-space:normal;line-height:1.25;background:var(--panel);border-bottom:1px solid var(--line);
  font-weight:700;color:var(--ink);vertical-align:bottom;padding-top:14px;padding-bottom:14px}
thead .cap{font-size:11.5px;font-weight:500;color:var(--muted);margin-top:5px}
thead .cap.warn{color:var(--amber);background:var(--amber-bg);display:inline-block;padding:1px 8px;border-radius:999px}
thead .sel{width:100%;min-width:92px;padding-left:8px;padding-right:26px;font-weight:600;min-height:36px}
thead th.cur{background:#EAF7F1;box-shadow:inset 0 -2px 0 var(--brand)}
tbody td.n{font-weight:600}
tbody td.cur{background:#F5FBF8}
tbody tr:hover td{background:#F4F6F8}
tbody tr:hover td.cur{background:#EDF8F3}
tbody td.l{font-weight:500}
tbody td.id{color:var(--muted)}
tr.total td{position:sticky;bottom:0;z-index:2;background:#F1F4F3;font-weight:800;border-top:2px solid var(--line);border-bottom:0}
.sortbtn{border:0;background:transparent;cursor:pointer;display:inline-flex;flex-direction:column;align-items:center;gap:2px;
  font-weight:700;padding:4px 8px;border-radius:6px;color:var(--ink)}
.sortbtn:hover{background:#E9ECEF}
.sortbtn small{font-weight:500;color:var(--muted);font-size:11.5px;white-space:normal;max-width:84px;line-height:1.25;text-align:center}
.sortbtn .ar{font-size:10px;color:var(--brand);height:10px}
.empty{padding:44px 20px;text-align:center;color:var(--muted)}
.empty b{display:block;color:var(--ink);font-size:15px;margin-bottom:4px}
footer{text-align:center;color:var(--faint);font-weight:600;padding:8px 0 4px;letter-spacing:.02em}

/* Shell + metrics sidebar */
.shell{display:grid;grid-template-columns:216px minmax(0,1fr);transition:grid-template-columns .28s cubic-bezier(.2,.7,.2,1);min-height:calc(100vh - var(--tb))}
.shell.collapsed{grid-template-columns:0 minmax(0,1fr)}
.main{min-width:0}
.side{position:sticky;top:var(--off);height:calc(100vh - var(--off));transition:top .28s cubic-bezier(.2,.7,.2,1),height .28s cubic-bezier(.2,.7,.2,1);overflow:hidden;background:#fff;border-right:1px solid var(--line);z-index:40}
.shell.collapsed .side{visibility:hidden;border-right-color:transparent;transition:visibility 0s .28s}
.side-in{width:216px;height:100%;display:flex;flex-direction:column}
.side-h{display:flex;align-items:center;justify-content:space-between;gap:8px;padding:16px 12px 10px 16px}
.side-h h2{margin:0;font-size:17px;font-weight:600}
.icon-btn{border:1px solid var(--line);background:#fff;border-radius:6px;width:36px;height:36px;display:inline-grid;place-items:center;
  cursor:pointer;color:var(--ink-2);flex:none;transition:background .15s,border-color .15s}
.icon-btn:hover{background:var(--panel);border-color:#C5CBD3}
.side-s{padding:0 12px 6px}
.side-s input{width:100%;border:1px solid #D7DBE0;border-radius:6px;padding:8px 12px;background:#fff;min-height:38px}
.side-list{flex:1;overflow:auto;padding:2px 8px 16px}
.mg{display:flex;align-items:center;gap:8px;margin:14px 0 4px;padding:0 10px;font-size:12.5px;font-weight:700;color:var(--brand-d)}
.mg svg{flex:none}
.mi{position:relative;display:flex;align-items:center;gap:10px;width:100%;text-align:left;border:0;background:transparent;cursor:pointer;
  padding:8px 10px 8px 14px;border-radius:6px;font-weight:500;font-size:13.5px;color:var(--ink-2);transition:background .15s,color .15s}
.mi:hover{background:#F1F3F5}
.mi .mi-ic{display:inline-grid;place-items:center;flex:none;color:var(--faint)}
.mi:hover .mi-ic{color:var(--ink-2)}
.mi[aria-current="true"] .mi-ic{color:var(--brand)}
.mi::before{content:"";position:absolute;left:0;top:8px;bottom:8px;width:3px;border-radius:3px;background:transparent}
.mi[aria-current="true"]{background:var(--brand-soft);color:var(--brand-d);font-weight:650}
.mi[aria-current="true"]::before{background:var(--brand)}
.mi .ch{margin-left:auto;display:inline-grid;place-items:center;color:var(--muted);flex:none}
.scrim{display:none}

@media (max-width:900px){
  .shell,.shell.collapsed{grid-template-columns:minmax(0,1fr)}
  .side{position:fixed;left:0;top:0;bottom:0;height:auto;width:290px;z-index:80;transition:transform .28s cubic-bezier(.2,.7,.2,1);box-shadow:0 0 50px rgba(20,28,36,.3)}
  .shell.collapsed .side{transform:translateX(-105%);visibility:hidden;transition:transform .28s,visibility 0s .28s}
  .scrim{display:block;position:fixed;inset:0;background:rgba(20,28,36,.35);z-index:70}
  .shell.collapsed .scrim{display:none}
  .topbar{position:static}
  :root{--off:0px}
  .tbl-wrap{overflow-x:auto}
}
@media (max-width:1100px){.kpis{grid-template-columns:repeat(2,minmax(0,1fr))}}
@media (max-width:640px){
  .kpis{grid-template-columns:1fr}
  .ms-btn{min-width:160px}
  .filters{padding:16px}
}
@media (prefers-reduced-motion:reduce){*{transition:none!important}}
</style>
</head>
<body class="tb-peek">

<div class="hot" id="tbHot" aria-hidden="true"><span class="hot-tab"><svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.4" stroke-linecap="round" stroke-linejoin="round"><path d="m6 9 6 6 6-6"/></svg>Menu</span></div>
<header class="topbar" id="topbar">
  <div class="tb-l">
    <button type="button" class="icon-btn" id="btnShow" aria-label="Show or hide metrics panel" aria-controls="side" title="Metrics panel">
      <svg width="18" height="18" viewBox="0 0 18 18" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round"><path d="M3 5h12M3 9h12M3 13h8"/></svg>
    </button>
    <img class="logo-img" src="__LOGO_SRC__" alt="KPN fresh" onerror="this.hidden=true;this.nextElementSibling.hidden=false">
    <div class="logo" role="img" aria-label="KPN fresh" hidden><b>KPN</b><i>fresh</i></div>
    <h1>Omni Metric
      <button type="button" class="pin" id="btnPin" aria-pressed="false" aria-label="Pin header" title="Pin header so it stays visible">
        <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.75" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M12 17v5"/><path d="M9 10.76a2 2 0 0 1-1.11 1.79l-1.78.9A2 2 0 0 0 5 15.24V16a1 1 0 0 0 1 1h12a1 1 0 0 0 1-1v-.76a2 2 0 0 0-1.11-1.79l-1.78-.9A2 2 0 0 1 15 10.76V7a1 1 0 0 1 1-1 2 2 0 0 0 0-4H8a2 2 0 0 0 0 4 1 1 0 0 1 1 1z"/></svg>
      </button>
    </h1>
  </div>
  <div class="chips">
    <span class="chip on">Metric <b id="chipMetric">-</b></span>
    <span class="chip">Data through <b id="chipDate">-</b></span>
    <span class="chip"><b id="chipStores">-</b> stores</span>
  </div>
</header>

<div class="shell" id="shell">
<aside class="side" id="side" aria-label="Metrics">
  <div class="side-in">
    <div class="side-h">
      <h2>Metrics</h2>
      <button type="button" class="icon-btn" id="btnHide" aria-label="Hide metrics panel" title="Hide panel">
        <svg width="16" height="16" viewBox="0 0 16 16" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><path d="M10 3.5 5.5 8l4.5 4.5"/></svg>
      </button>
    </div>
    <div class="side-s"><input type="search" id="metricSearch" placeholder="Find a metric" aria-label="Find a metric"></div>
    <nav class="side-list" id="metricList" aria-label="Metric list"></nav>
  </div>
</aside>
<div class="scrim" id="scrim"></div>

<div class="main">
<main class="wrap">
  <section class="card" aria-label="Filters">
    <div class="card-h"><h2>Filters</h2><button type="button" class="linkbtn" id="btnReset">Clear filters</button></div>
    <div class="filters">
      <div class="fg"><span class="lbl" id="l-gran">Time granularity</span>
        <div class="seg" role="group" aria-labelledby="l-gran" id="segGran">
          <button type="button" data-v="month" aria-pressed="true">Month</button>
          <button type="button" data-v="week" aria-pressed="false">Week</button>
        </div></div>

      <div class="fg" id="gCmp"><span class="lbl" id="l-cmp">Month comparison</span>
        <div class="seg" role="group" aria-labelledby="l-cmp" id="segCmp">
          <button type="button" data-v="full" aria-pressed="true">Full month</button>
          <button type="button" data-v="mtd" aria-pressed="false">MTD</button>
        </div></div>

      <div class="fg" id="gUpto" hidden><label for="selUpto">Up to day</label>
        <select class="sel small" id="selUpto"></select></div>

      <div class="fg"><span class="lbl" id="l-ch">Channel</span>
        <div class="seg" role="group" aria-labelledby="l-ch" id="segCh">
          <button type="button" data-v="ECOM" aria-pressed="false">ECOM</button>
          <button type="button" data-v="POS" aria-pressed="false">POS</button>
          <button type="button" data-v="OVERALL" aria-pressed="true">Overall</button>
        </div></div>

      <div class="fg"><span class="lbl">Cluster</span><div class="ms" id="msCluster"></div></div>
      <div class="fg"><span class="lbl">Store name</span><div class="ms" id="msStore"></div></div>
    </div>
  </section>

  <div class="hint" id="hint" hidden></div>

  <section class="kpis" id="kpis" aria-label="Key figures"></section>

  <section class="card panel" aria-label="Store-wise summary">
    <div class="panel-h"><h2>Store-wise summary</h2><span id="storeNote"></span></div>
    <div class="tbl-wrap" id="storeWrap"></div>
  </section>

  <section class="card panel" aria-label="Cluster-wise summary">
    <div class="panel-h"><h2>Cluster-wise summary</h2><span id="clusterNote"></span></div>
    <div class="tbl-wrap" id="clusterWrap"></div>
  </section>

  <footer>KPN Fresh</footer>
</main>
</div>
</div>

<script>
(function(){
'use strict';
const DATA = /*__DATA__*/null;

/* ---------- constants ---------- */
const NM=21, W=24, POSORD=21, SSUM=22, SN=23;   // raw metric columns, then POS orders / scan sum / scan days
const METRICS=['Orders','GMV','FnV sale','Fruits sale','Vegetables sale','Milk sale','Fruits second sale','Vegi second sale',
  'Salad sale','Loose provision sale','Own brand sale','FMCG sale','Staples sale','Ice cream sale','Rice bag qty','Salad qty',
  'Mob Pent %','EC used','Scan %','Fruit Bills %','Veg Bills %','PL Bills %'];
const PCT_METRICS=new Set([16,18,19,20,21]);
const MN=['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
const $=(s,r=document)=>r.querySelector(s);
const esc=s=>String(s).replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));

/* ---------- date helpers ---------- */
const mLabel=ym=>MN[ym%100-1]+'-'+String(Math.floor(ym/100)).slice(2);
function ymShift(ym,d){let y=Math.floor(ym/100),m=ym%100-1+d;y+=Math.floor(m/12);m=((m%12)+12)%12;return y*100+m+1}
function isoWeekNum(d){
  const t=new Date(Date.UTC(d.getUTCFullYear(),d.getUTCMonth(),d.getUTCDate()));
  const dn=(t.getUTCDay()+6)%7; t.setUTCDate(t.getUTCDate()-dn+3);
  const fy=new Date(Date.UTC(t.getUTCFullYear(),0,4));
  return 1+Math.round(((t-fy)/864e5-3+((fy.getUTCDay()+6)%7))/7);
}
const weeksInYear=y=>isoWeekNum(new Date(Date.UTC(y,11,28)));
function weekPrev(w){const y=Math.floor(w/100),n=w%100;return n>1?w-1:(y-1)*100+weeksInYear(y-1)}
function weekShiftY(w,d){const y=Math.floor(w/100)+d;let n=w%100;if(n>weeksInYear(y))n=weeksInYear(y);return y*100+n}
const isPartial=w=>DATA.weekDays[w]!==undefined&&DATA.weekDays[w]<7;
const wLabel=w=>String(w)+(isPartial(w)?' (partial)':'');

/* ---------- state ---------- */
const S={gran:'month',channel:'OVERALL',metric:1,mtd:false,upto:31,m:[0,0,0,0],w:[0,0,0,0],
  clusters:new Set(),stores:new Set(),sort:{cluster:null,store:null}};
const months=DATA.months.slice().sort((a,b)=>a-b);
const weeks=Object.keys(DATA.weekDays).map(Number).sort((a,b)=>a-b);
const latestYm=months[months.length-1];

function setMonthCols(c4){
  S.m=[ymShift(c4,-13),ymShift(c4,-12),ymShift(c4,-1),c4];
  const run=(c4===latestYm&&DATA.monthLast[c4]<DATA.monthDays[c4]);
  S.mtd=run; S.upto=DATA.monthLast[c4]||DATA.monthDays[c4]||31;
}
function setWeekCols(c4){
  const c3=weekPrev(c4);
  S.w=[weekShiftY(c3,-1),weekShiftY(c4,-1),c3,c4];
}
(function init(){
  setMonthCols(latestYm);
  let w4=weeks[weeks.length-1];
  for(let i=weeks.length-1;i>=0;i--){if(DATA.weekDays[weeks[i]]>=7){w4=weeks[i];break}}
  setWeekCols(w4);
})();

/* ---------- store selection ---------- */
const stores=DATA.stores; // [id,name,cluster,fin]
const sel=new Uint8Array(stores.length);
const cmpStore=(a,b)=>stores[a][2].localeCompare(stores[b][2])||stores[a][1].localeCompare(stores[b][1]);
function recomputeSel(){
  for(let s=0;s<stores.length;s++){
    sel[s]=((S.clusters.size===0||S.clusters.has(stores[s][2]))&&(S.stores.size===0||S.stores.has(s)))?1:0;
  }
}

/* ---------- aggregation ---------- */
const colKeys=()=>S.gran==='month'?S.m:S.w;
function compute(){
  const map=new Map();
  colKeys().forEach((k,i)=>{if(!map.has(k))map.set(k,[]);map.get(k).push(i)});
  const acc=new Float64Array(stores.length*4*W);
  const mon=S.gran==='month',ki=mon?1:2,ch=S.channel,mtd=mon&&S.mtd,upto=S.upto;
  const rows=DATA.rows;
  for(let n=0;n<rows.length;n++){
    const r=rows[n],s=r[0];
    if(!sel[s])continue;
    const src=r[4];
    if(ch==='ECOM'&&src!==0)continue;
    if(ch==='POS'&&src!==1)continue;
    const list=map.get(r[ki]); if(!list)continue;
    if(mtd&&r[3]>upto)continue;
    for(let q=0;q<list.length;q++){
      const o=(s*4+list[q])*W;
      for(let m=0;m<NM;m++)acc[o+m]+=r[5+m];
      if(src===1)acc[o+POSORD]+=r[5];
      if(src===0&&r[5+NM]>0){acc[o+SSUM]+=r[5+NM];acc[o+SN]++}
    }
  }
  return acc;
}
function valueAt(a,o){
  const m=S.metric;
  if(m===16)return a[o+POSORD]>0?a[o+16]/a[o+POSORD]*100:0;
  if(m===18)return a[o+SN]>0?a[o+SSUM]/a[o+SN]:0;
  if(m>=19)return a[o]>0?a[o+m-1]/a[o]*100:0;   // Fruit/Veg/PL bills % = bills / orders
  return a[o+m];
}
const vals4=(a,base)=>[0,1,2,3].map(p=>valueAt(a,base+p*W));
function addInto(t,a,base){for(let i=0;i<4*W;i++)t[i]+=a[base+i]}
const growth=(r,b)=>{const g=b?r/b-1:0;return isFinite(g)?g:0};
const pctDefs=()=>S.gran==='month'
  ?[{r:1,b:0,t:'MoM'},{r:3,b:2,t:'MoM'}]
  :[{r:1,b:0,t:'WoW'},{r:3,b:2,t:'WoW'},{r:2,b:0,t:'YoY'},{r:3,b:1,t:'YoY'}];

/* ---------- formatting ---------- */
const nf0=new Intl.NumberFormat('en-IN',{maximumFractionDigits:0});
const fmt=v=>PCT_METRICS.has(S.metric)?v.toFixed(1)+'%':nf0.format(Math.round(v));
function fmtCompact(v){
  if(PCT_METRICS.has(S.metric))return v.toFixed(1)+'%';
  const a=Math.abs(v);
  if(a>=1e7)return (v/1e7).toFixed(2)+' Cr';
  if(a>=1e5)return (v/1e5).toFixed(2)+' L';
  return nf0.format(Math.round(v));
}
function pill(g){
  if(!g)return '<span class="pill flat">0.00%</span>';
  const up=g>0;
  return `<span class="pill ${up?'up':'down'}"><i>${up?'&#9650;':'&#9660;'}</i>${Math.abs(g*100).toFixed(2)}%</span>`;
}
const colLabel=i=>S.gran==='month'?mLabel(S.m[i]):String(S.w[i]);
function caption(i){
  if(S.gran==='month'){
    const ym=S.m[i];
    if(DATA.monthLast[ym]===undefined)return {t:'No data',w:false};
    return {t:S.mtd?`Day 1\u2013${S.upto}`:'Full month',w:false};
  }
  const w=S.w[i],n=DATA.weekDays[w];
  if(n===undefined)return {t:'No data',w:false};
  return n<7?{t:`Partial: ${n} of 7 days`,w:true}:{t:`Week ${w%100}`,w:false};
}

/* ---------- table rendering ---------- */
function periodOpts(i){
  const cur=colKeys()[i];
  const base=S.gran==='month'?months:weeks;
  const set=new Set(base);set.add(cur);
  return [...set].sort((a,b)=>b-a).map(v=>
    `<option value="${v}"${v===cur?' selected':''}>${S.gran==='month'?mLabel(v):wLabel(v)}</option>`).join('');
}
function periodTh(i){
  const c=caption(i);
  return `<th class="${i===3?'cur':''}"><select class="sel period-sel" data-i="${i}" aria-label="Period ${i+1}">${periodOpts(i)}</select>
    <div class="cap${c.w?' warn':''}">${esc(c.t)}</div></th>`;
}
function pctTh(kind,di,d){
  const st=S.sort[kind];
  const on=st&&st.i===di;
  const arrow=on?(st.dir===1?'&#9650;':'&#9660;'):'';
  return `<th ${on?`aria-sort="${st.dir===1?'ascending':'descending'}"`:''}>
    <button type="button" class="sortbtn" data-kind="${kind}" data-di="${di}" title="Sort by ${d.t}">
      <span>${d.t}</span><small>${esc(colLabel(d.r))} vs ${esc(colLabel(d.b))}</small><span class="ar">${arrow}</span></button></th>`;
}
function buildTable(kind,lead,rows,total,defs){
  const head=`<thead><tr>${lead.map((h,i)=>`<th class="l">${h}</th>`).join('')}${[0,1,2,3].map(periodTh).join('')}${defs.map((d,i)=>pctTh(kind,i,d)).join('')}</tr></thead>`;
  const st=S.sort[kind];
  if(st)rows=rows.slice().sort((a,b)=>st.dir*(a.g[st.i]-b.g[st.i]));
  const cells=(v,g)=>v.map((x,p)=>`<td class="n${p===3?' cur':''}">${fmt(x)}</td>`).join('')+g.map(x=>`<td class="p">${pill(x)}</td>`).join('');
  const body=rows.map(r=>`<tr>${r.lead}${cells(r.v,r.g)}</tr>`).join('');
  const tot=`<tr class="total"><td class="l" colspan="${lead.length}">Grand total</td>${cells(total.v,total.g)}</tr>`;
  return `<table aria-label="${kind} table">${head}<tbody>${body}${tot}</tbody></table>`;
}

let last=null;
function render(){
  const active=document.activeElement;
  const refocus=active&&active.classList&&active.classList.contains('period-sel')?active.dataset.i:null;
  const defs=pctDefs();
  const acc=compute();
  const idx=[...stores.keys()].filter(s=>sel[s]).sort(cmpStore);
  const grand=new Float64Array(4*W);
  const clMap=new Map();
  const storeRows=idx.map(s=>{
    const base=s*4*W,v=vals4(acc,base),st=stores[s];
    addInto(grand,acc,base);
    const c=st[2];
    if(!clMap.has(c))clMap.set(c,{a:new Float64Array(4*W),n:0});
    const g=clMap.get(c);addInto(g.a,acc,base);g.n++;
    return {s,v,g:defs.map(d=>growth(v[d.r],v[d.b])),
      lead:`<td class="l id">${esc(st[3])}</td><td class="l id">${esc(st[0])}</td><td class="l">${esc(st[1])}</td><td class="l">${esc(c)}</td>`};
  });
  const gv=vals4(grand,0);
  const total={v:gv,g:defs.map(d=>growth(gv[d.r],gv[d.b]))};
  const clusterRows=[...clMap.keys()].sort((a,b)=>a.localeCompare(b)).map(c=>{
    const o=clMap.get(c),v=vals4(o.a,0);
    return {v,g:defs.map(d=>growth(v[d.r],v[d.b])),lead:`<td class="l">${esc(c)}</td><td class="l id">${o.n}</td>`};
  });

  $('#kpis').innerHTML=kpiHTML(total,defs);
  const none='<div class="empty"><b>No stores match these filters</b>Clear a cluster or store filter to see data.</div>';
  $('#clusterWrap').innerHTML=idx.length?buildTable('cluster',['Cluster','Stores'],clusterRows,total,defs):none;
  $('#storeWrap').innerHTML=idx.length?buildTable('store',['Fin code','Store ID','Store name','Cluster'],storeRows,total,defs):none;
  $('#clusterNote').textContent=`${clusterRows.length} cluster${clusterRows.length===1?'':'s'}`;
  $('#storeNote').textContent=`${METRICS[S.metric]}, ${idx.length} store${idx.length===1?'':'s'}`;
  $('#chipMetric').textContent=METRICS[S.metric];
  const bad=((S.metric===16||S.metric===17)&&S.channel==='ECOM')||(S.metric===18&&S.channel==='POS');
  const hint=$('#hint');hint.hidden=!bad;
  if(bad)hint.textContent=S.metric===18?'Scan % is recorded on ECOM only. Switch Channel to ECOM or Overall to see values.'
    :`${METRICS[S.metric]} is recorded on POS only. Switch Channel to POS or Overall to see values.`;
  last={storeRows,total,defs};
  if(refocus!==null){const el=document.querySelector(`#storeWrap .period-sel[data-i="${refocus}"]`);if(el)el.focus()}
}
function kpiHTML(total,defs){
  return [0,1,2,3].map(p=>{
    const c=caption(p);
    const pills=defs.filter(d=>d.r===p).map(d=>
      `<span class="k-pill">${d.t} ${pill(growth(total.v[d.r],total.v[d.b]))}</span>`).join('');
    const mx=Math.max(...total.v,0);
    const wpc=mx>0?Math.max(2,total.v[p]/mx*100):0;
    return `<article class="card kpi${p===3?' lead':''}">
      <div class="k-label">${esc(colLabel(p))}${p===3?'<span class="badge">Latest</span>':''}</div><div class="k-cap">${esc(c.t)}</div>
      <div class="k-val" title="${esc(fmt(total.v[p]))}">${fmtCompact(total.v[p])}</div>
      <div class="k-bar" aria-hidden="true"><i style="width:${wpc.toFixed(1)}%"></i></div>
      <div class="k-pills">${pills}</div></article>`;
  }).join('');
}

/* ---------- multi-select ---------- */
function Multi(root,cfg){
  const set=cfg.set;
  root.innerHTML=`<button type="button" class="ms-btn" aria-haspopup="true" aria-expanded="false"></button>
    <div class="ms-panel" hidden><input class="ms-search" type="search" placeholder="Search" aria-label="Search ${cfg.noun}">
    <div class="ms-actions"><button type="button" data-a="all">Select all</button><button type="button" data-a="clear">Clear</button></div>
    <ul class="ms-list"></ul></div>`;
  const btn=$('.ms-btn',root),panel=$('.ms-panel',root),search=$('.ms-search',root),list=$('.ms-list',root);
  function items(){return cfg.items()}
  function summary(){
    if(set.size===0)btn.textContent='All '+cfg.noun;
    else if(set.size===1){const it=items().find(i=>set.has(i.v));btn.textContent=it?it.t:'1 selected'}
    else btn.textContent=set.size+' selected';
  }
  function draw(){
    const q=search.value.trim().toLowerCase();
    const its=items().filter(i=>!q||(i.t+' '+(i.sub||'')).toLowerCase().includes(q));
    list.innerHTML=its.length?its.map(i=>`<li><label><input type="checkbox" value="${esc(i.v)}"${set.has(i.v)?' checked':''}>
      <span>${esc(i.t)}</span>${i.sub?`<small>${esc(i.sub)}</small>`:''}</label></li>`).join(''):'<li class="ms-empty">Nothing matches your search.</li>';
    summary();
  }
  function open(o){panel.hidden=!o;btn.setAttribute('aria-expanded',o);if(o){draw();search.focus()}}
  btn.addEventListener('click',()=>open(panel.hidden));
  search.addEventListener('input',draw);
  list.addEventListener('change',e=>{
    const v=cfg.num?Number(e.target.value):e.target.value;
    e.target.checked?set.add(v):set.delete(v);summary();cfg.onChange();
  });
  panel.addEventListener('click',e=>{
    const a=e.target.dataset&&e.target.dataset.a;if(!a)return;
    const q=search.value.trim().toLowerCase();
    items().filter(i=>!q||(i.t+' '+(i.sub||'')).toLowerCase().includes(q)).forEach(i=>a==='all'?set.add(i.v):set.delete(i.v));
    if(a==='clear'&&!q)set.clear();
    draw();cfg.onChange();
  });
  document.addEventListener('click',e=>{if(!root.contains(e.target))open(false)});
  root.addEventListener('keydown',e=>{if(e.key==='Escape'){open(false);btn.focus()}});
  return {draw,summary};
}
const clusterNames=[...new Set(stores.map(s=>s[2]))].sort((a,b)=>a.localeCompare(b));
const msCluster=Multi($('#msCluster'),{noun:'clusters',set:S.clusters,
  items:()=>clusterNames.map(c=>({v:c,t:c})),
  onChange:()=>{
    // keep only stores that belong to the chosen clusters
    if(S.clusters.size)[...S.stores].forEach(s=>{if(!S.clusters.has(stores[s][2]))S.stores.delete(s)});
    msStore.draw();recomputeSel();render();
  }});
const msStore=Multi($('#msStore'),{noun:'stores',set:S.stores,num:true,
  items:()=>[...stores.keys()].filter(s=>S.clusters.size===0||S.clusters.has(stores[s][2])).sort(cmpStore)
    .map(s=>({v:s,t:stores[s][1],sub:stores[s][2]})),
  onChange:()=>{recomputeSel();render()}});

/* ---------- controls ---------- */
function syncSeg(id,val){document.querySelectorAll(`#${id} button`).forEach(b=>b.setAttribute('aria-pressed',b.dataset.v===val))}
function syncControls(){
  const mon=S.gran==='month';
  $('#gCmp').hidden=!mon;$('#gUpto').hidden=!(mon&&S.mtd);
  syncSeg('segGran',S.gran);syncSeg('segCmp',S.mtd?'mtd':'full');syncSeg('segCh',S.channel);
  $('#selUpto').value=S.upto;
}
$('#selUpto').innerHTML=Array.from({length:31},(_,i)=>`<option value="${i+1}">${i+1}</option>`).join('');
$('#segGran').addEventListener('click',e=>{const v=e.target.dataset.v;if(v){S.gran=v;S.sort={cluster:null,store:null};syncControls();render()}});
$('#segCmp').addEventListener('click',e=>{const v=e.target.dataset.v;if(v){S.mtd=v==='mtd';syncControls();render()}});
$('#segCh').addEventListener('click',e=>{const v=e.target.dataset.v;if(v){S.channel=v;syncControls();render()}});
$('#selUpto').addEventListener('change',e=>{S.upto=+e.target.value;render()});

/* ---------- metrics sidebar ---------- */
const isMobile=()=>!!(window.matchMedia&&window.matchMedia('(max-width:900px)').matches);
const IC=d=>`<svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.75" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">${d}</svg>`;
const METRIC_ICONS=[
  /* 0 Orders: receipt */ IC('<path d="M4 2v20l2-1 2 1 2-1 2 1 2-1 2 1 2-1 2 1V2l-2 1-2-1-2 1-2-1-2 1-2-1-2 1Z"/><path d="M14 8H8"/><path d="M16 12H8"/><path d="M13 16H8"/>'),
  /* 1 GMV: rupee */ IC('<path d="M6 3h12"/><path d="M6 8h12"/><path d="m6 13 8.5 8"/><path d="M6 13h3"/><path d="M9 13c6.667 0 6.667-10 0-10"/>'),
  /* 2 FnV: leaf */ IC('<path d="M11 20A7 7 0 0 1 9.8 6.1C15.5 5 17 4.48 19 2c1 2 2 4.18 2 8 0 5.5-4.78 10-10 10Z"/><path d="M2 21c0-3 1.85-5.36 5.08-6C9.5 14.52 12 13 13 12"/>'),
  /* 3 Fruits: apple */ IC('<path d="M12 20.94c1.5 0 2.75 1.06 4 1.06 3 0 6-8 6-12.22A4.91 4.91 0 0 0 17 5c-2.22 0-4 1.44-5 2-1-.56-2.78-2-5-2a4.9 4.9 0 0 0-5 4.78C2 14 5 22 8 22c1.25 0 2.5-1.06 4-1.06Z"/><path d="M10 2c1 .5 2 2 2 5"/>'),
  /* 4 Vegetables: carrot */ IC('<path d="M2.27 21.7s9.87-3.5 12.73-6.36a4.5 4.5 0 0 0-6.36-6.37C5.77 11.84 2.27 21.7 2.27 21.7zM8.64 14l-2.05-2.04M15.34 15l-2.46-2.46"/><path d="M22 9s-1.33-2-3.5-2C16.86 7 15 9 15 9s1.33 2 3.5 2S22 9 22 9z"/><path d="M15 2s-2 1.33-2 3.5S15 9 15 9s2-1.84 2-3.5C17 3.33 15 2 15 2z"/>'),
  /* 5 Milk */ IC('<path d="M8 2h8"/><path d="M9 2v2.789a4 4 0 0 1-.672 2.219l-.656.984A4 4 0 0 0 7 10.212V20a2 2 0 0 0 2 2h6a2 2 0 0 0 2-2v-9.789a4 4 0 0 0-.672-2.219l-.656-.984A4 4 0 0 1 15 4.788V2"/><path d="M7 15a6.472 6.472 0 0 1 5 0 6.47 6.47 0 0 0 5 0"/>'),
  /* 6 Fruits second: discount badge */ IC('<path d="M3.85 8.62a4 4 0 0 1 4.78-4.77 4 4 0 0 1 6.74 0 4 4 0 0 1 4.78 4.78 4 4 0 0 1 0 6.74 4 4 0 0 1-4.77 4.78 4 4 0 0 1-6.75 0 4 4 0 0 1-4.78-4.77 4 4 0 0 1 0-6.76Z"/><path d="m15 9-6 6"/><path d="M9 9h.01"/><path d="M15 15h.01"/>'),
  /* 7 Vegi second */ IC('<path d="M3.85 8.62a4 4 0 0 1 4.78-4.77 4 4 0 0 1 6.74 0 4 4 0 0 1 4.78 4.78 4 4 0 0 1 0 6.74 4 4 0 0 1-4.77 4.78 4 4 0 0 1-6.75 0 4 4 0 0 1-4.78-4.77 4 4 0 0 1 0-6.76Z"/><path d="m15 9-6 6"/><path d="M9 9h.01"/><path d="M15 15h.01"/>'),
  /* 8 Salad: bowl */ IC('<path d="M7 21h10"/><path d="M12 21a9 9 0 0 0 9-9H3a9 9 0 0 0 9 9Z"/><path d="M11.38 12a2.4 2.4 0 0 1-.4-4.77 2.4 2.4 0 0 1 3.2-2.77 2.4 2.4 0 0 1 3.47-.63 2.4 2.4 0 0 1 3.37 3.37 2.4 2.4 0 0 1-1.1 3.7 2.51 2.51 0 0 1 .03 1.1"/><path d="m13 12 4-4"/><path d="M10.9 7.25A3.99 3.99 0 0 0 4 10c0 .73.2 1.41.54 2"/>'),
  /* 9 Loose provision: scale */ IC('<path d="m16 16 3-8 3 8c-.87.65-1.92 1-3 1s-2.13-.35-3-1Z"/><path d="m2 16 3-8 3 8c-.87.65-1.92 1-3 1s-2.13-.35-3-1Z"/><path d="M7 21h10"/><path d="M12 3v18"/><path d="M3 7h2c2 0 5-1 7-2 2 1 5 2 7 2h2"/>'),
  /* 10 Own brand: award */ IC('<path d="m15.477 12.89 1.515 8.526a.5.5 0 0 1-.81.47l-3.58-2.687a1 1 0 0 0-1.197 0l-3.586 2.686a.5.5 0 0 1-.81-.469l1.514-8.526"/><circle cx="12" cy="8" r="6"/>'),
  /* 11 FMCG: cart */ IC('<circle cx="8" cy="21" r="1"/><circle cx="19" cy="21" r="1"/><path d="M2.05 2.05h2l2.66 12.42a2 2 0 0 0 2 1.58h9.78a2 2 0 0 0 1.95-1.57l1.65-7.43H5.12"/>'),
  /* 12 Staples: wheat */ IC('<path d="M2 22 16 8"/><path d="M3.47 12.53 5 11l1.53 1.53a3.5 3.5 0 0 1 0 4.94L5 19l-1.53-1.53a3.5 3.5 0 0 1 0-4.94Z"/><path d="M7.47 8.53 9 7l1.53 1.53a3.5 3.5 0 0 1 0 4.94L9 15l-1.53-1.53a3.5 3.5 0 0 1 0-4.94Z"/><path d="M11.47 4.53 13 3l1.53 1.53a3.5 3.5 0 0 1 0 4.94L13 11l-1.53-1.53a3.5 3.5 0 0 1 0-4.94Z"/><path d="M20 2h2v2a4 4 0 0 1-4 4h-2V6a4 4 0 0 1 4-4Z"/><path d="M11.47 17.47 13 19l-1.53 1.53a3.5 3.5 0 0 1-4.94 0L5 19l1.53-1.53a3.5 3.5 0 0 1 4.94 0Z"/><path d="M15.47 13.47 17 15l-1.53 1.53a3.5 3.5 0 0 1-4.94 0L9 15l1.53-1.53a3.5 3.5 0 0 1 4.94 0Z"/><path d="M19.47 9.47 21 11l-1.53 1.53a3.5 3.5 0 0 1-4.94 0L13 11l1.53-1.53a3.5 3.5 0 0 1 4.94 0Z"/>'),
  /* 13 Ice cream: cone */ IC('<path d="m7 11 4.08 10.35a1 1 0 0 0 1.84 0L17 11"/><path d="M17 7A5 5 0 0 0 7 7"/><path d="M17 7a2 2 0 0 1 0 4H7a2 2 0 0 1 0-4"/>'),
  /* 14 Rice bag qty: bag */ IC('<path d="M6 2 3 6v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2V6l-3-4Z"/><path d="M3 6h18"/><path d="M16 10a4 4 0 0 1-8 0"/>'),
  /* 15 Salad qty: package */ IC('<path d="m7.5 4.27 9 5.15"/><path d="M21 8a2 2 0 0 0-1-1.73l-7-4a2 2 0 0 0-2 0l-7 4A2 2 0 0 0 3 8v8a2 2 0 0 0 1 1.73l7 4a2 2 0 0 0 2 0l7-4A2 2 0 0 0 21 16Z"/><path d="m3.3 7 8.7 5 8.7-5"/><path d="M12 22V12"/>'),
  /* 16 Mob Pent: phone */ IC('<rect width="14" height="20" x="5" y="2" rx="2" ry="2"/><path d="M12 18h.01"/>'),
  /* 17 EC used: ticket */ IC('<path d="M2 9a3 3 0 0 1 0 6v2a2 2 0 0 0 2 2h16a2 2 0 0 0 2-2v-2a3 3 0 0 1 0-6V7a2 2 0 0 0-2-2H4a2 2 0 0 0-2 2Z"/><path d="M13 5v2"/><path d="M13 17v2"/><path d="M13 11v2"/>'),
  /* 18 Scan %: scan line */ IC('<path d="M3 7V5a2 2 0 0 1 2-2h2"/><path d="M17 3h2a2 2 0 0 1 2 2v2"/><path d="M21 17v2a2 2 0 0 1-2 2h-2"/><path d="M7 21H5a2 2 0 0 1-2-2v-2"/><path d="M7 12h10"/>'),
  /* 19 Fruit Bills %: cherry */ IC('<path d="M2 17a5 5 0 0 0 10 0c0-2.76-2.5-5-5-3-2.5-2-5 .24-5 3Z"/><path d="M12 17a5 5 0 0 0 10 0c0-2.76-2.5-5-5-3-2.5-2-5 .24-5 3Z"/><path d="M7 14c3.22-2.91 4.29-8.75 5-12 1.66 2.38 4.94 9 5 12"/><path d="M22 9c-4.29 0-7.14-2.33-10-7 5.71 0 10 4.67 10 7Z"/>'),
  /* 20 Veg Bills %: sprout */ IC('<path d="M7 20h10"/><path d="M10 20c5.5-2.5.8-6.4 3-10"/><path d="M9.5 9.4c1.1.8 1.8 2.2 2.3 3.7-2 .4-3.5.4-4.8-.3-1.2-.6-2.3-1.9-3-4.2 2.8-.5 4.4 0 5.5.8z"/><path d="M14.1 6a7 7 0 0 0-1.1 4c1.9-.1 3.3-.6 4.3-1.4 1-1 1.6-2.3 1.7-4.6-2.7.1-4 1-4.9 2z"/>'),
  /* 21 PL Bills %: basket */ IC('<path d="m15 11-1 9"/><path d="m19 11-4-7"/><path d="M2 11h20"/><path d="m3.5 11 1.6 7.4a2 2 0 0 0 2 1.6h9.8a2 2 0 0 0 2-1.6l1.7-7.4"/><path d="M4.5 15.5h15"/><path d="m5 11 4-7"/><path d="m9 11 1 9"/>')
];
const CHI=d=>`<svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.75" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">${d}</svg>`;
const CH_ICONS={
  pos:CHI('<path d="m2 7 4.41-4.41A2 2 0 0 1 7.83 2h8.34a2 2 0 0 1 1.42.59L22 7"/><path d="M4 12v8a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2v-8"/><path d="M15 22v-4a2 2 0 0 0-2-2h-2a2 2 0 0 0-2 2v4"/><path d="M2 7h20"/><path d="M22 7v3a2 2 0 0 1-2 2a2.7 2.7 0 0 1-1.59-.63.7.7 0 0 0-.82 0A2.7 2.7 0 0 1 16 12a2.7 2.7 0 0 1-1.59-.63.7.7 0 0 0-.82 0A2.7 2.7 0 0 1 12 12a2.7 2.7 0 0 1-1.59-.63.7.7 0 0 0-.82 0A2.7 2.7 0 0 1 8 12a2.7 2.7 0 0 1-1.59-.63.7.7 0 0 0-.82 0A2.7 2.7 0 0 1 4 12a2 2 0 0 1-2-2V7"/>'),
  ecom:CHI('<circle cx="12" cy="12" r="10"/><path d="M12 2a14.5 14.5 0 0 0 0 20 14.5 14.5 0 0 0 0-20"/><path d="M2 12h20"/>')};
const CHAN={16:'pos',17:'pos',18:'ecom'};
const CHAN_TXT={pos:'Recorded on POS only',ecom:'Recorded on ECOM only'};
function drawMetrics(){
  const q=$('#metricSearch').value.trim().toLowerCase();
  const its=METRICS.map((_,i)=>i).filter(m=>!q||METRICS[m].toLowerCase().includes(q));
  $('#metricList').innerHTML=its.map(m=>{
    const c=CHAN[m];
    return `<button type="button" class="mi" data-m="${m}"${m===S.metric?' aria-current="true"':''}>
      <span class="mi-ic">${METRIC_ICONS[m]}</span><span>${METRICS[m]}</span>${c?`<span class="ch" title="${CHAN_TXT[c]}" role="img" aria-label="${CHAN_TXT[c]}">${CH_ICONS[c]}</span>`:''}</button>`;
  }).join('')||'<div class="mg"><span>No metric matches.</span></div>';
}
$('#metricSearch').addEventListener('input',drawMetrics);
$('#metricList').addEventListener('click',e=>{
  const b=e.target.closest('.mi');if(!b)return;
  S.metric=+b.dataset.m;S.sort={cluster:null,store:null};drawMetrics();render();
  if(isMobile())setSide(false);
});
function setSide(open){
  $('#shell').classList.toggle('collapsed',!open);
  $('#btnShow').setAttribute('aria-expanded',open);
  try{localStorage.setItem('m26_side',open?'1':'0')}catch(_){}
}
$('#btnShow').addEventListener('click',()=>setSide($('#shell').classList.contains('collapsed')));
$('#btnHide').addEventListener('click',()=>setSide(false));
$('#scrim').addEventListener('click',()=>setSide(false));
document.addEventListener('keydown',e=>{if(e.key==='Escape'&&isMobile())setSide(false)});
$('#btnReset').addEventListener('click',()=>{S.clusters.clear();S.stores.clear();msCluster.draw();msStore.draw();recomputeSel();render()});

document.addEventListener('change',e=>{
  const t=e.target;if(!t.classList||!t.classList.contains('period-sel'))return;
  const i=+t.dataset.i,v=+t.value;
  if(S.gran==='month'){
    if(i===3)setMonthCols(v);else S.m[i]=v;
    syncControls();
  }else{
    S.w[i]=v;
    if(i===0)S.w[2]=weekShiftY(v,1);
    else if(i===2)S.w[0]=weekShiftY(v,-1);
    else if(i===1)S.w[3]=weekShiftY(v,1);
    else S.w[1]=weekShiftY(v,-1);
  }
  render();
});
document.addEventListener('click',e=>{
  const b=e.target.closest&&e.target.closest('.sortbtn');if(!b)return;
  const kind=b.dataset.kind,di=+b.dataset.di,st=S.sort[kind];
  if(!st||st.i!==di)S.sort[kind]={i:di,dir:1};        // 1st click: negative to positive
  else if(st.dir===1)S.sort[kind]={i:di,dir:-1};      // 2nd click: positive to negative
  else S.sort[kind]=null;                             // 3rd click: reset
  render();
});

/* ---------- header pin / auto-hide ---------- */
(function(){
  const tb=$('#topbar'),hot=$('#tbHot'),btn=$('#btnPin'),body=document.body;
  let pinned=false,timer=null,intro=null;
  const within=()=>tb.contains(document.activeElement);
  function paint(){
    body.classList.toggle('tb-pinned',pinned);
    btn.setAttribute('aria-pressed',pinned);
    btn.title=pinned?'Unpin header so it hides automatically':'Pin header so it stays visible';
    btn.setAttribute('aria-label',pinned?'Unpin header':'Pin header');
  }
  function peek(v){if(v)body.classList.add('tb-peek');else body.classList.remove('tb-peek')}
  function setPin(v,save){
    pinned=v;clearTimeout(timer);clearTimeout(intro);paint();
    if(v)peek(false);                       // docked: always visible, no overlay needed
    else if(!tb.matches(':hover')&&!within())peek(false);   // unpinned: hide right away
    if(save){try{localStorage.setItem('m26_pin',v?'1':'0')}catch(_){}}
  }
  const open=()=>{clearTimeout(timer);clearTimeout(intro);if(!pinned)peek(true)};
  const closeSoon=()=>{clearTimeout(timer);timer=setTimeout(()=>{if(!pinned&&!within())peek(false)},400)};
  hot.addEventListener('mouseenter',open);
  tb.addEventListener('mouseenter',open);
  tb.addEventListener('mouseleave',closeSoon);
  tb.addEventListener('focusin',open);
  tb.addEventListener('focusout',()=>setTimeout(()=>{if(!within()&&!tb.matches(':hover'))closeSoon()},0));
  document.addEventListener('keydown',e=>{if(e.key==='Escape'&&!pinned&&!isMobile()){peek(false);if(within())document.activeElement.blur()}});
  btn.addEventListener('click',()=>setPin(!pinned,true));
  let saved=false;try{saved=localStorage.getItem('m26_pin')==='1'}catch(_){}
  pinned=saved;paint();
  if(pinned)peek(false);
  else intro=setTimeout(()=>{if(!tb.matches(':hover')&&!within())peek(false)},2600);  // brief reveal on load, then auto-hide
})();

/* ---------- boot ---------- */
(function(){
  const d=new Date(DATA.lastDate+'T00:00:00Z');
  $('#chipDate').textContent=d.getUTCDate()+' '+MN[d.getUTCMonth()]+' '+d.getUTCFullYear();
  $('#chipStores').textContent=stores.length;
  let open=!isMobile();
  try{const v=localStorage.getItem('m26_side');if(v!==null&&open)open=v==='1'}catch(_){}
  setSide(open);
  recomputeSel();syncControls();msCluster.draw();msStore.draw();drawMetrics();render();
})();
})();
</script>
</body>
</html>
'''

if __name__ == "__main__":
    main()
