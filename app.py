# -*- coding: utf-8 -*-
"""
Mainspring Report Automation  ·  Flask edition
──────────────────────────────────────────────
Run :  python app.py
Open:  http://localhost:5000
"""
import io, base64, os, re as _re
from datetime import datetime
from flask import Flask, request, jsonify
import pandas as pd
from openpyxl import Workbook, load_workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.chart import PieChart, Reference
from openpyxl.cell.cell import MergedCell

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 100 * 1024 * 1024

# ══════════════════════════════════════════════════════════════════════════════
# PROCESSING HELPERS
# ══════════════════════════════════════════════════════════════════════════════
def find_col(df, *keywords):
    if df is None: return None
    low = [c.lower() for c in df.columns]
    for kw in keywords:
        kw = kw.lower()
        for i, c in enumerate(low):
            if kw in c: return df.columns[i]
    return None

def norm(s): return s.astype(str).str.lower().str.strip()

BLANKS = {"", "nan", "none", "nat", "n/a", "na", "null"}

# ── Upload storage ─────────────────────────────────────────────────────────────
UPLOAD_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "uploads")
os.makedirs(UPLOAD_DIR, exist_ok=True)

# ── Month helpers ──────────────────────────────────────────────────────────────
_MONTH_MAP = {
    'january':1,'february':2,'march':3,'april':4,'may':5,'june':6,
    'july':7,'august':8,'september':9,'october':10,'november':11,'december':12,
    'jan':1,'feb':2,'mar':3,'apr':4,'jun':6,'jul':7,'aug':8,
    'sep':9,'sept':9,'oct':10,'nov':11,'dec':12
}

def parse_month_year(label):
    """'May 2026' -> (5, 2026). Returns (None, None) on failure."""
    if not label: return None, None
    mn = yr = None
    for p in label.lower().split():
        if p in _MONTH_MAP:  mn = _MONTH_MAP[p]
        elif p.isdigit() and len(p) == 4: yr = int(p)
    return mn, yr

def _parse_dates(series):
    """
    Parse a date Series robustly for ServiceNow exports.
    Tries DD/MM/YYYY (dayfirst=True) and MM/DD/YYYY (dayfirst=False).
    Uses whichever produces fewer NaT values.
    On a tie (all days ≤12, ambiguous), prefers DD/MM/YYYY — the standard
    format for ServiceNow exports in India/UK deployments.
    Suppresses pandas format-inference warnings.
    """
    import warnings
    with warnings.catch_warnings():
        warnings.simplefilter("ignore")
        d_dmy = pd.Series(pd.to_datetime(series, errors="coerce", dayfirst=True))
        d_mdy = pd.Series(pd.to_datetime(series, errors="coerce", dayfirst=False))
    # Prefer DD/MM/YYYY on tie — covers Indian/UK ServiceNow deployments
    return d_dmy if d_dmy.isna().sum() <= d_mdy.isna().sum() else d_mdy

def filter_by_month(df, month_num, year, label=""):
    """
    Filter dataframe to the selected reporting month.
    Priority 1: CLOSED/RESOLVED date (tickets resolved this month).
    Priority 2: OPENED/CREATED date (fallback for new tickets).
    Returns (filtered_df, debug_dict).
    """
    before = len(df) if df is not None else 0
    dbg = {"sheet": label, "col": None, "all_cols": [],
           "before": before, "after": before, "applied": False, "note": ""}

    if df is None or not month_num or not year:
        dbg["note"] = "no df or no month/year"
        return df, dbg

    dbg["all_cols"] = list(df.columns)

    def _try(col):
        """
        Parse col as dates and filter to month_num/year.
        Returns (filtered_df, note_str):
          - filtered_df is None ONLY when the column is not a real date column (>80% NaT)
            or on parse exception.
          - An empty filtered_df (0 rows) is a VALID result — don't fall through.
        """
        try:
            dates = _parse_dates(df[col])
            nat_p = int(pd.isna(dates).mean() * 100)
            if nat_p > 80:
                return None, f"col='{col}' skipped — {nat_p}% NaT (not a date column)"
            dti  = pd.DatetimeIndex(dates)
            mask = (dti.month == month_num) & (dti.year == year)
            res  = df[mask].reset_index(drop=True)
            return res, f"col='{col}' matched {len(res)}/{len(df)} rows (NaT={nat_p}%)"
        except Exception as ex:
            return None, f"col='{col}' err={ex}"

    # Priority 1: closed/resolved date columns
    # Try these in order — first parseable date column wins, even if 0 rows.
    cc = find_col(df,
                  "resolved at", "closed at", "close time", "close date",
                  "resolution date", "resolved", "closed", "end date")
    if cc:
        res, note = _try(cc)
        if res is not None:          # None = not a date column, skip
            dbg.update(col=str(cc), after=len(res), applied=True,
                       note=f"CLOSED-date | {note}")
            print(f"  [{label}] filter -> CLOSED col='{cc}' {len(res)}/{before} rows")
            return res, dbg

    # Priority 2: opened/created date (fallback — only reached if no closed-date col found)
    oc = find_col(df,
                  "opened at", "created at", "open time", "start date",
                  "date opened", "sys_created", "opened", "created", "incident date")
    if oc:
        res, note = _try(oc)
        if res is not None:
            dbg.update(col=str(oc), after=len(res), applied=True,
                       note=f"OPENED-date fallback | {note}")
            print(f"  [{label}] filter -> OPENED col='{oc}' {len(res)}/{before} rows")
            return res, dbg

    dbg["note"] = f"no date col matched {month_num}/{year} — full set kept"
    return df, dbg

def filter_by_daterange(df, start_dt, end_dt, label=""):
    """
    Filter dataframe to a date range [start_dt, end_dt] inclusive.
    Same column priority as filter_by_month: closed/resolved first, then opened.
    Returns (filtered_df, debug_dict).
    """
    before = len(df) if df is not None else 0
    dbg = {"sheet": label, "col": None, "before": before, "after": before,
           "applied": False, "note": ""}
    if df is None or df.empty:
        dbg["note"] = "empty df"
        return df, dbg
    try:
        if pd.isna(start_dt) or pd.isna(end_dt):
            dbg["note"] = "invalid date range"
            return df, dbg
    except (TypeError, ValueError):
        dbg["note"] = "invalid date range"
        return df, dbg
    end_inc = end_dt + pd.Timedelta(days=1) - pd.Timedelta(seconds=1)

    def _try(col):
        try:
            dates = _parse_dates(df[col])
            nat_p = int(pd.isna(dates).mean() * 100)
            if nat_p > 80:
                return None, f"col='{col}' skipped ({nat_p}% NaT)"
            dti  = pd.DatetimeIndex(dates)
            mask = (dti >= start_dt) & (dti <= end_inc)
            res  = df[mask].reset_index(drop=True)
            return res, f"col='{col}' matched {len(res)}/{len(df)} rows (NaT={nat_p}%)"
        except Exception as ex:
            return None, f"col='{col}' err={ex}"

    cc = find_col(df,
                  "resolved at", "closed at", "close time", "close date",
                  "resolution date", "resolved", "closed", "end date")
    if cc:
        res, note = _try(cc)
        if res is not None:
            dbg.update(col=str(cc), after=len(res), applied=True, note=f"CLOSED | {note}")
            return res, dbg

    oc = find_col(df,
                  "opened at", "created at", "open time", "start date",
                  "date opened", "sys_created", "opened", "created", "incident date")
    if oc:
        res, note = _try(oc)
        if res is not None:
            dbg.update(col=str(oc), after=len(res), applied=True, note=f"OPENED | {note}")
            return res, dbg

    dbg["note"] = "no date col found"
    return df, dbg

# ── ServiceNow ticket-number detection ────────────────────────────────────────
_SN_PAT = _re.compile(r'^(INC|RITM|CHG|PRB|REQ|TASK|SCTASK|SC)\d{5,}$', _re.IGNORECASE)

def _find_ticket_id_col(df):
    """
    Find the ticket-number column by scanning cell values for SN patterns
    (INC0012345, RITM0001234, CHG…, PRB…).
    Falls back to column-name keywords if pattern scan fails.
    """
    if df is None or df.empty: return None
    best_col, best_ratio = None, 0.0
    for col in df.columns:
        try:
            sample = df[col].dropna().astype(str).head(50)
            if sample.empty: continue
            ratio = sample.str.match(_SN_PAT).mean()
            if ratio > best_ratio:
                best_ratio, best_col = ratio, col
        except Exception:
            pass
    if best_ratio >= 0.60:
        return best_col
    return find_col(df, "number","ticket number","ticket no","ticket id",
                    "incident number","request item","change number",
                    "problem number","ref no","reference","sys_id")

def dedup_tickets(df, label=""):
    """
    Deduplicate a ServiceNow export so every unique ticket is counted exactly once.
    Keeps the row with the most-recent updated/resolved date (= final state).
    Returns (deduped_df, debug_dict).
    """
    if df is None or len(df) == 0:
        return df, {"sheet":label,"id_col":None,"before":0,"after":0,"dupes":0,"deduped":False,"note":"empty"}
    before = len(df)
    dbg = {"sheet":label,"id_col":None,"before":before,"after":before,"dupes":0,"deduped":False,"note":""}

    id_col = _find_ticket_id_col(df)
    if id_col is None:
        dbg["note"] = "ticket ID col not found — cols: " + ", ".join(str(c) for c in df.columns[:15])
        return df, dbg
    dbg["id_col"] = str(id_col)

    try:
        work = df.copy()
        upd_col = find_col(work, "updated","sys_updated","modified","last modified","update time")
        if not upd_col:
            upd_col = find_col(work, "resolved","closed","close time")
        if upd_col:
            work["__upd"] = pd.to_datetime(work[upd_col], errors="coerce")
            work = work.sort_values("__upd", na_position="first")
            work.drop(columns=["__upd"], inplace=True)
        deduped = work.drop_duplicates(subset=[id_col], keep="last").reset_index(drop=True)
        after   = len(deduped)
        dupes   = before - after
        note    = (f"{dupes} duplicate rows removed" if dupes > 0 else "no duplicates")
        dbg.update(after=after, dupes=dupes, deduped=dupes>0, note=note)
        return deduped, dbg
    except Exception as ex:
        dbg["note"] = f"exception: {ex}"
        return df, dbg

def _bytes_to_df(data, filename):
    """Parse raw bytes into a DataFrame."""
    if not data or not filename: return None
    try:
        if filename.lower().endswith(".csv"):
            return pd.read_csv(io.BytesIO(data))
        return pd.read_excel(io.BytesIO(data))
    except Exception:
        return None

def save_upload_cycle(raw_files, month_label=""):
    """
    Save uploaded file bytes to a timestamped folder.
    raw_files: dict { "RITM": (filename, bytes), ... }
    Returns folder path.
    """
    ts    = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    safe  = month_label.replace(" ","_").replace("/","-") if month_label else ""
    fname = f"{ts}_{safe}" if safe else ts
    folder = os.path.join(UPLOAD_DIR, fname)
    os.makedirs(folder, exist_ok=True)
    for key, (orig, data) in raw_files.items():
        if data and orig:
            with open(os.path.join(folder, f"{key}_{orig}"), "wb") as fh:
                fh.write(data)
    with open(os.path.join(folder, "_info.txt"), "w", encoding="utf-8") as fh:
        fh.write(f"Cycle : {ts}\nMonth : {month_label}\nFiles : {', '.join(raw_files)}\n")
    return folder

def _load_sheet(folder_path, prefix):
    """Load first file in folder whose name starts with prefix_."""
    try:
        for fn in sorted(os.listdir(folder_path)):
            if fn.startswith(prefix + "_") and not fn.startswith("_"):
                path = os.path.join(folder_path, fn)
                if fn.lower().endswith(".csv"):
                    return pd.read_csv(path)
                return pd.read_excel(path)
    except Exception:
        pass
    return None

RITM_SUCCESS  = {"closed complete","fulfilled","closed - complete","closed-complete","close complete"}
RITM_TERMINAL = RITM_SUCCESS | {"closed","cancelled","rejected","closed incomplete",
                                 "resolved","complete","withdrawn","closed - incomplete","void","voided"}

def _read_df(fs):
    if fs is None or not fs.filename: return None
    try:
        data = fs.read()
        return _bytes_to_df(data, fs.filename)
    except Exception:
        return None

# ══════════════════════════════════════════════════════════════════════════════
# TEMPLATE ROWS
# ══════════════════════════════════════════════════════════════════════════════
TEMPLATE_ROWS = [
    ("data",    "Number of reopened tickets",                                           "reopened_tickets",         "Mostit Always 0"),
    ("data",    "Total tickets Resolved",                                               "total_tickets_resolved",   "Total no of inc+PRBLM+MACM"),
    ("empty",   "", None, ""),
    ("section", "Mesaure Entry for Process Specific Metrics", None, ""),
    ("data",    "Total Number of Problem records resolved",                             "prob_records_resolved",    "Total number of problem ticket, stated should be closed complete"),
    ("data",    "Total Number of Incidents resolved",                                   "inc_resolved",             "Total number of closed and resolved in state cloum"),
    ("data",    "Total Number of Incident carry forwarded",                             "inc_carry_fwd",            "Filter other than closed and resolved in state cloum"),
    ("data",    "Total Number of Incidents resolved",                                   "inc_resolved_2",           "Total number of Incident ticket, stated should be closed complete (same as row 7)"),
    ("data",    "Total Number of New incident received",                                "new_inc_received",         "Total count of inc irrespective of state"),
    ("sumit_b", "Number of resources onboarded within the agreed time frame",           "resources_onboarded_sla",  ""),
    ("sumit_b", "Total number of resources onboarded",                                  "total_resources",          ""),
    ("sumit_b", "Actual effort of closed problem tickets",                              "effort_prob_closed",       "Sumit will fill"),
    ("data",    "Total number of unknown incident closed",                              "unknown_inc_closed",       "Incidents with State closed & Resolved AND unselect Known Error from Resolution code."),
    ("data",    "Total number of unknown incident reference with KEDB articles",        "unknown_inc_kedb",         "Closed & Resolved, unselect Known Error, unselect blanks from Problem"),
    ("data",    "Total change requests to be delivered",                                "change_deliver",           "total number of macms"),
    ("data",    "Total changes implemented sucessfully  within agreed timelines",       "changes_success",          "same as 15"),
    ("data",    "Number of changes that had to be rolled back",                         "changes_rolled_back",      "Never rollback"),
    ("data",    "total number of changes  closed",                                      "changes_closed",           "same as 15"),
    ("data",    "Total number of known and unknown incident closed",                    "ku_inc_closed",            "Select all in PROBLEM cloumn, In resolution code select all remove blanks"),
    ("data",    "Total number of known and unknown incident with KEDB articles Referenced", "ku_inc_kedb",          "Keep same as 19 row and in PROBLEM remove blanks"),
    ("data",    "Number of proactive problem tickets raised",                           "proactive_raised",         "Total number of problem ticket, stated should be closed complete"),
    ("data",    "Number of proactive problem tickets resolved",                         "proactive_resolved",       "Total number of problem ticket, stated should be closed complete"),
    ("data",    "Total number of Known error incidents closed",                         "known_err_inc_closed",     "Incidents with State closed & Resolved AND select Known Error from Resolution code."),
    ("data",    "Number of New known incident Received",                                "new_known_inc",            "everytime 0"),
    ("data",    "Number of Known Incident carry forwarded",                             "known_inc_carry_fwd",      "everytime 0"),
    ("data",    "Number of Standard Service Requests reopened",                         "ssr_reopened",             "everytime 0"),
    ("data",    "Total Number of Unknown Incident carry forwarded",                     "unk_inc_carry_fwd",        "everytime 0"),
    ("data",    "Total Number of New Unknown Incident received",                        "new_unk_inc_received",     "unselect Known Error from Resolution code."),
    ("data",    "Total unresolved Standard Service Requests open beyond their agreed time", "ssr_open_beyond_sla",  "in RITM sheet, Remove closed and fullfilled, Consider rest"),
    ("data",    "Total Number of Standard Service Requests",                            "total_ssr",                "Total number of RITM's"),
    ("data",    "Number of Standard Service Requests completed on time",                "ssr_completed_on_time",    "in RITM sheet, select closed and fullfilled"),
    ("data",    "Number of Ad hoc Service Requests reopened",                           "adhoc_reopened",           "everytime 0"),
    ("data",    "Total Number of Ad hoc Service Requests resolved",                     "adhoc_resolved",           "In RITM search for keyword ad hoc in short description"),
    ("data",    "Total Number of Standard Service Requests resolved",                   "ssr_resolved",             "Same as 31"),
    ("data",    "Total Number of Ad hoc Service Requests",                              "total_adhoc",              "Total number of adhoc request"),
    ("data",    "Total unresolved Ad hoc Service Requests open beyond their agreed timeline", "adhoc_open_sla",     "Total number of adhoc request"),
    ("data",    "Number of Ad hoc Service Requests completed on time",                  "adhoc_on_time",            "Total number of adhoc request"),
    ("sumit_ab","Total Effort required to resolve Standard Service Requests",           "effort_ssr",               ""),
    ("sumit_ab","Total Effort required to resolve Ad hoc Service Requests",             "effort_adhoc",             ""),
    ("sumit_ab","Actual Effort of unknown incidents closed",                            "effort_unknown_inc",       ""),
    ("sumit_ab","Actual Effort of known incidents closed",                              "effort_known_inc",         ""),
]

SUMIT_KEYS = {r[2] for r in TEMPLATE_ROWS if r[0] in ("sumit_b","sumit_ab") and r[2]}

# ══════════════════════════════════════════════════════════════════════════════
# CALCULATE METRICS
# ══════════════════════════════════════════════════════════════════════════════
def calculate_metrics(ritm_df, inc_df, macm_df, prob_df, prob_skipped, inc_received_count=None):
    # dfs are pre-filtered (month + dedup) by the caller

    # ── RITM ──────────────────────────────────────────────────────────────────
    total_ritm = len(ritm_df) if ritm_df is not None else 0
    ritm_completed = ritm_open_count = total_adhoc = 0
    if ritm_df is not None:
        rs = find_col(ritm_df, "state", "status")
        if rs:
            rn = norm(ritm_df[rs])
            ritm_success    = rn.isin(RITM_SUCCESS)
            ritm_open       = ~rn.isin(RITM_TERMINAL)
            ritm_completed  = int(ritm_success.sum())
            ritm_open_count = int(ritm_open.sum())
        else:
            ritm_completed  = 0
            ritm_open_count = total_ritm
        # Ad hoc = RITM rows with "ad hoc" in short description
        ds = find_col(ritm_df, "short description", "description", "short desc")
        if ds:
            total_adhoc = int(ritm_df[ds].astype(str).str.lower()
                              .str.contains(r"ad.?hoc", regex=True, na=False).sum())
        else:
            total_adhoc = 0

    # ── INCIDENTS ─────────────────────────────────────────────────────────────
    total_inc = len(inc_df) if inc_df is not None else 0
    # new_inc_received = tickets OPENED this month (passed from caller if available)
    _new_inc = inc_received_count if inc_received_count is not None else total_inc
    inc_resolved = inc_carry_fwd = unknown_closed = unknown_kedb = 0
    ku_closed = ku_kedb = known_err_closed = new_unk_received = 0
    if inc_df is not None:
        i_st = find_col(inc_df, "state", "status")
        i_rc = find_col(inc_df, "resolution code", "resolution_code", "resolution")
        i_pr = find_col(inc_df, "problem")

        if i_st:
            ins        = norm(inc_df[i_st])
            inc_closed = ins.isin({"closed", "resolved"})
        else:
            inc_closed = pd.Series([False] * total_inc)

        inc_open = ~inc_closed

        if i_rc:
            rc           = norm(inc_df[i_rc])
            known_err    = rc.str.contains("known error", na=False)
            rc_not_blank = ~rc.isin(BLANKS)
        else:
            known_err    = pd.Series([False] * total_inc)
            rc_not_blank = pd.Series([False] * total_inc)

        if i_pr:
            pr       = norm(inc_df[i_pr])
            has_prob = ~pr.isin(BLANKS)
        else:
            has_prob = pd.Series([False] * total_inc)

        inc_resolved     = int(inc_closed.sum())
        inc_carry_fwd    = int(inc_open.sum())
        unknown_closed   = int((inc_closed & ~known_err).sum())
        unknown_kedb     = int((inc_closed & ~known_err & has_prob).sum())
        ku_closed        = int((inc_closed & rc_not_blank).sum())
        ku_kedb          = int((inc_closed & rc_not_blank & has_prob).sum())
        known_err_closed = int((inc_closed & known_err).sum())

    # ── PROBLEM TICKETS ───────────────────────────────────────────────────────
    total_prob = 0
    if not prob_skipped and prob_df is not None:
        ps = find_col(prob_df, "state", "status")
        if ps:
            pn = norm(prob_df[ps])
            total_prob = int(pn.isin({"closed complete","closed - complete",
                                      "close complete","closed","resolved"}).sum())
        else:
            total_prob = 0

    # ── MACM / CHANGES ────────────────────────────────────────────────────────
    total_macm = macm_success = macm_rolled_back = 0
    if macm_df is not None:
        total_macm = len(macm_df)
        ms = find_col(macm_df, "state", "status")
        if ms:
            mn = norm(macm_df[ms])
            MACM_OK  = {"closed complete","closed - complete","closed","implemented",
                        "successful","complete","close complete","closed-complete","resolved"}
            MACM_BAD = {"cancelled","failed","rolled back","roll back","rejected","withdrawn"}
            macm_success     = int(mn.isin(MACM_OK).sum())
            macm_rolled_back = int(mn.isin(MACM_BAD).sum())
        else:
            macm_success     = 0
            macm_rolled_back = 0

    # ── TOTALS ────────────────────────────────────────────────────────────────
    total_resolved = inc_resolved + ritm_completed + total_prob + macm_success

    return dict(
        reopened_tickets        = 0,
        total_tickets_resolved  = total_resolved,
        prob_records_resolved   = total_prob,
        inc_resolved            = inc_resolved,
        inc_carry_fwd           = inc_carry_fwd,
        inc_resolved_2          = inc_resolved,
        new_inc_received        = _new_inc,
        resources_onboarded_sla = None,
        total_resources         = None,
        effort_prob_closed      = None,
        unknown_inc_closed      = unknown_closed,
        unknown_inc_kedb        = unknown_kedb,
        change_deliver          = total_macm,
        changes_success         = macm_success,     # MACMs successfully completed
        changes_rolled_back     = macm_rolled_back,
        changes_closed          = macm_success,     # same as changes_success
        ku_inc_closed           = ku_closed,
        ku_inc_kedb             = ku_kedb,
        proactive_raised        = total_prob,
        proactive_resolved      = total_prob,
        known_err_inc_closed    = known_err_closed,
        new_known_inc           = 0,                # everytime 0
        known_inc_carry_fwd     = 0,                # everytime 0
        ssr_reopened            = 0,                # everytime 0
        unk_inc_carry_fwd       = 0,                # everytime 0
        new_unk_inc_received    = 0,                # everytime 0
        ssr_open_beyond_sla     = ritm_open_count,
        total_ssr               = total_ritm,
        ssr_completed_on_time   = ritm_completed,
        adhoc_reopened          = 0,
        adhoc_resolved          = total_adhoc,
        ssr_resolved            = ritm_completed,
        total_adhoc             = total_adhoc,
        adhoc_open_sla          = total_adhoc,
        adhoc_on_time           = total_adhoc,
        effort_ssr              = None,
        effort_adhoc            = None,
        effort_unknown_inc      = None,
        effort_known_inc        = None,
    )

# ══════════════════════════════════════════════════════════════════════════════
# FILL / GENERATE EXCEL
# ══════════════════════════════════════════════════════════════════════════════
def fill_excel(ref_bytes, metrics, manual_inputs):
    wb = load_workbook(io.BytesIO(ref_bytes))
    ws = wb.worksheets[0]
    name_map = {}
    for rtype, name, key, _ in TEMPLATE_ROWS:
        if key and rtype not in ("empty", "section"):
            is_sumit = rtype in ("sumit_b", "sumit_ab")
            name_map[name.lower().strip()] = (key, is_sumit)
    val_col = 2
    for hdr_cell in ws[1]:
        col_idx = hdr_cell.column
        if col_idx is not None and hdr_cell.value and "value" in str(hdr_cell.value).lower():
            val_col = int(col_idx); break
    filled = 0
    for row_cells in ws.iter_rows():
        a_cell = row_cells[0]
        if a_cell.value is None: continue
        raw = str(a_cell.value).lower().strip()
        if raw not in name_map: continue
        mkey, is_sumit = name_map[raw]
        v = (manual_inputs.get(mkey) or "") if is_sumit else (metrics.get(mkey, 0) or 0)
        target = ws.cell(row=int(a_cell.row), column=int(val_col))
        if not isinstance(target, MergedCell):
            target.value = v; filled += 1
    buf = io.BytesIO(); wb.save(buf); buf.seek(0)
    return buf, filled

def generate_excel(metrics, manual_inputs, month_label):
    wb = Workbook(); ws = wb.worksheets[0]
    setattr(ws, "title", (month_label or "Report")[:31])
    HDR="375623"; CRIT="E2EFDA"; YEL="FFFF00"; WHT="FFFFFF"
    ws.column_dimensions["A"].width=52; ws.column_dimensions["B"].width=10; ws.column_dimensions["C"].width=90
    thin=Side(style="thin",color="999999"); bdr=Border(left=thin,right=thin,top=thin,bottom=thin)
    def cel(row,col,val,bold=False,clr=WHT,fc="000000",ha="left",wrap=True,sz=10,ital=False):
        c=ws[f"{col}{row}"]; c.value=val
        c.font=Font(name="Calibri",bold=bold,italic=ital,size=sz,color=fc)
        c.fill=PatternFill("solid",fgColor=clr)
        c.alignment=Alignment(horizontal=ha,vertical="center",wrap_text=wrap); c.border=bdr
    def gv(key):
        if key in SUMIT_KEYS:
            v=manual_inputs.get(key); return v if (v is not None and v!=0) else ""
        return metrics.get(key,0) or 0
    cel(1,"A","Measure Name",bold=True,clr=HDR,fc=WHT,sz=11)
    cel(1,"B","Value",bold=True,clr=HDR,fc=WHT,ha="center",sz=11)
    cel(1,"C","Measurement Criteria",bold=True,clr=HDR,fc=WHT,sz=11)
    ws.row_dimensions[1].height=20
    row=2
    for rtype,name,key,criteria in TEMPLATE_ROWS:
        if rtype=="empty": ws.row_dimensions[row].height=8; row+=1; continue
        if rtype=="section":
            ws.merge_cells(f"A{row}:C{row}"); c=ws[f"A{row}"]
            c.value=name; c.font=Font(name="Calibri",bold=True,size=10,color=WHT)
            c.fill=PatternFill("solid",fgColor=HDR)
            c.alignment=Alignment(horizontal="left",vertical="center")
            c.border=bdr; ws.row_dimensions[row].height=18; row+=1; continue
        val=gv(key)
        if rtype=="data":
            cel(row,"A",name); cel(row,"B",val,bold=True,ha="center"); cel(row,"C",criteria,clr=CRIT,ital=True)
        elif rtype=="sumit_b":
            cel(row,"A",name); cel(row,"B",val,bold=True,clr=YEL,ha="center"); cel(row,"C",criteria,clr=CRIT)
        elif rtype=="sumit_ab":
            cel(row,"A",name,clr=YEL); cel(row,"B",val,bold=True,clr=YEL,ha="center"); cel(row,"C",criteria,clr=CRIT)
        ws.row_dimensions[row].height=18; row+=1
    ws.row_dimensions[row].height=5; row+=1
    ws.merge_cells(f"A{row}:C{row}"); n=ws[f"A{row}"]
    n.value="* Yellow cells = manual entry (Sumit)"
    n.font=Font(name="Calibri",size=9,italic=True,color="856404")
    n.fill=PatternFill("solid",fgColor="FFFBEA")
    n.alignment=Alignment(horizontal="left",vertical="center"); ws.row_dimensions[row].height=16
    ws2=wb.create_sheet("Summary")
    for i,h in enumerate(["Category","Count"],1):
        c=ws2.cell(1,i,h); c.font=Font(bold=True,color=WHT)
        c.fill=PatternFill("solid",fgColor="375623"); c.alignment=Alignment(horizontal="center")
    ws2.column_dimensions["A"].width=30; ws2.column_dimensions["B"].width=12
    summary=[("Total Incidents",metrics.get("new_inc_received",0)),
             ("Incidents Resolved",metrics.get("inc_resolved",0)),
             ("Carry Forward",metrics.get("inc_carry_fwd",0)),
             ("Known Error Closed",metrics.get("known_err_inc_closed",0)),
             ("Unknown Closed",metrics.get("unknown_inc_closed",0)),
             ("Total RITMs",metrics.get("total_ssr",0)),
             ("RITMs Completed",metrics.get("ssr_completed_on_time",0)),
             ("Total MACMs",metrics.get("change_deliver",0)),
             ("Problem Records",metrics.get("prob_records_resolved",0))]
    for r,(cat,v) in enumerate(summary,2): ws2.cell(r,1,cat); ws2.cell(r,2,v)
    pie=PieChart(); pie.title="Ticket Distribution"; pie.style=10
    pie.add_data(Reference(ws2,min_col=2,min_row=1,max_row=6),titles_from_data=True)
    pie.set_categories(Reference(ws2,min_col=1,min_row=2,max_row=6))
    pie.width=14; pie.height=10; ws2.add_chart(pie,"D2")
    # ── Pivot Summary sheet ───────────────────────────────────────────────────
    wp=wb.create_sheet("Pivot Summary")
    pc=PatternFill("solid",fgColor="1e3a5f")
    pg=PatternFill("solid",fgColor="e8f5e9")
    py=PatternFill("solid",fgColor="fff9c4")
    po=PatternFill("solid",fgColor="fff3e0")
    pb=PatternFill("solid",fgColor="e3f2fd")
    pp=PatternFill("solid",fgColor="f3e5f5")
    pthin=Side(style="thin",color="cccccc")
    pbdr=Border(left=pthin,right=pthin,top=pthin,bottom=pthin)
    def ph(r,c,v,fill=None,bold=False,ha="center",fc="000000",sz=10):
        cell=wp.cell(r,c,v)
        cell.font=Font(name="Calibri",bold=bold,size=sz,color=fc)
        cell.alignment=Alignment(horizontal=ha,vertical="center",wrap_text=True)
        cell.border=pbdr
        if fill: cell.fill=fill
        return cell
    wp.column_dimensions["A"].width=32
    for col in ["B","C","D","E","F"]: wp.column_dimensions[col].width=16
    # Title
    wp.merge_cells("A1:F1")
    t=wp["A1"]; t.value=f"Mainspring Monthly Report — {month_label}"
    t.font=Font(name="Calibri",bold=True,size=13,color="FFFFFF")
    t.fill=pc; t.alignment=Alignment(horizontal="center",vertical="center")
    wp.row_dimensions[1].height=28
    # Header row
    for ci,hdr in enumerate(["Category","Total Tickets","Resolved/Closed","Carry Forward / Open","Success Rate","Notes"],1):
        ph(2,ci,hdr,fill=PatternFill("solid",fgColor="2563eb"),bold=True,fc="FFFFFF",sz=10)
    wp.row_dimensions[2].height=20
    inc_total   = metrics.get("new_inc_received",0)
    inc_res     = metrics.get("inc_resolved",0)
    inc_cf      = metrics.get("inc_carry_fwd",0)
    inc_rate    = f"{round(inc_res/max(inc_total,1)*100)}%" if inc_total else "—"
    ritm_total  = metrics.get("total_ssr",0)
    ritm_done   = metrics.get("ssr_completed_on_time",0)
    ritm_open   = metrics.get("ssr_open_beyond_sla",0)
    ritm_rate   = f"{round(ritm_done/max(ritm_total,1)*100)}%" if ritm_total else "—"
    macm_total  = metrics.get("change_deliver",0)
    macm_ok     = metrics.get("changes_success",0)
    macm_rb     = metrics.get("changes_rolled_back",0)
    macm_rate   = f"{round(macm_ok/max(macm_total,1)*100)}%" if macm_total else "—"
    prob_total  = metrics.get("prob_records_resolved",0)
    total_res   = metrics.get("total_tickets_resolved",0)
    rows_data=[
        ("Incidents (INC)",         inc_total,  inc_res,  inc_cf,   inc_rate,  "Closed+Resolved vs Carry Forward", pg),
        ("RITM (Svc Requests)",     ritm_total, ritm_done,ritm_open,ritm_rate, "Closed Complete + Fulfilled",      pb),
        ("Changes (MACM)",          macm_total, macm_ok,  macm_rb,  macm_rate, "Successfully Implemented",         py),
        ("Problem Tickets (PRB)",   prob_total, prob_total,0,       "—",       "Closed Complete",                  po),
    ]
    for ri,( cat,tot,res,cf,rate,note,fill) in enumerate(rows_data,3):
        ph(ri,1,cat, fill=fill,bold=True,ha="left")
        ph(ri,2,tot, fill=fill,bold=True)
        ph(ri,3,res, fill=fill,bold=True)
        ph(ri,4,cf,  fill=fill)
        ph(ri,5,rate,fill=fill)
        ph(ri,6,note,fill=fill,ha="left")
        wp.row_dimensions[ri].height=18
    # Total row
    tr=3+len(rows_data)
    wp.merge_cells(f"A{tr}:D{tr}")
    ph(tr,1,"TOTAL TICKETS RESOLVED",fill=PatternFill("solid",fgColor="1e3a5f"),bold=True,fc="FFFFFF",ha="left",sz=11)
    ph(tr,5,total_res,fill=PatternFill("solid",fgColor="1e3a5f"),bold=True,fc="FFFFFF",sz=12)
    ph(tr,6,"INC+RITM+PROB+MACM",fill=PatternFill("solid",fgColor="1e3a5f"),fc="FFFFFF",ha="left")
    wp.row_dimensions[tr].height=24
    # Known/Unknown breakdown
    kr=tr+2
    ph(kr,1,"Known Error Incidents",fill=pp,bold=True,ha="left")
    ph(kr,2,metrics.get("known_err_inc_closed",0),fill=pp)
    ph(kr,3,"—",fill=pp); ph(kr,4,"—",fill=pp); ph(kr,5,"—",fill=pp)
    ph(kr,6,"State Closed + Res.Code=Known Error",fill=pp,ha="left")
    ph(kr+1,1,"Unknown Incidents Closed",fill=pg,bold=True,ha="left")
    ph(kr+1,2,metrics.get("unknown_inc_closed",0),fill=pg)
    ph(kr+1,3,"—",fill=pg); ph(kr+1,4,"—",fill=pg); ph(kr+1,5,"—",fill=pg)
    ph(kr+1,6,"Closed, Res.Code ≠ Known Error",fill=pg,ha="left")
    ph(kr+2,1,"Ad Hoc Requests",fill=py,bold=True,ha="left")
    ph(kr+2,2,metrics.get("total_adhoc",0),fill=py)
    ph(kr+2,3,metrics.get("adhoc_resolved",0),fill=py)
    ph(kr+2,4,"—",fill=py); ph(kr+2,5,"—",fill=py)
    ph(kr+2,6,"RITM with 'ad hoc' in Short Desc",fill=py,ha="left")
    for rr in [kr,kr+1,kr+2]: wp.row_dimensions[rr].height=17

    buf=io.BytesIO(); wb.save(buf); buf.seek(0); return buf

def build_table_html(M, mi):
    rows = ""
    for rtype, name, key, criteria in TEMPLATE_ROWS:
        if rtype == "empty":
            rows += '<tr><td colspan="3" class="empty-row"></td></tr>'; continue
        if rtype == "section":
            rows += f'<tr class="sec-row"><td colspan="3">{name}</td></tr>'; continue
        if rtype == "sumit_b":
            v = mi.get(key) or ""
            rows += f'<tr><td class="name-cell">{name}</td><td class="val-cell sumit">{v}</td><td class="crit-cell">{criteria}</td></tr>'
        elif rtype == "sumit_ab":
            v = mi.get(key) or ""
            rows += f'<tr><td class="name-cell sumit">{name}</td><td class="val-cell sumit">{v}</td><td class="crit-cell">{criteria}</td></tr>'
        else:
            v = M.get(key, 0) or 0
            rows += f'<tr><td class="name-cell">{name}</td><td class="val-cell">{v}</td><td class="crit-cell">{criteria}</td></tr>'
    return rows

# ══════════════════════════════════════════════════════════════════════════════
# PER-STEP PREVIEW ROUTES
# ══════════════════════════════════════════════════════════════════════════════
@app.route("/preview/ritm", methods=["POST"])
def preview_ritm():
    try:
        df = _read_df(request.files.get("file"))
        if df is None: return jsonify({"ok": False, "error": "Could not read file"})
        total = len(df); completed = open_count = adhoc = 0
        rs = find_col(df, "state", "status")
        if rs:
            rn = norm(df[rs])
            completed  = int(rn.isin(RITM_SUCCESS).sum())
            open_count = int((~rn.isin(RITM_TERMINAL)).sum())
        else:
            completed  = 0
            open_count = total
        ds = find_col(df, "short description", "description")
        if ds:
            adhoc = int(df[ds].astype(str).str.lower()
                        .str.contains(r"ad.?hoc", regex=True, na=False).sum())
        return jsonify({"ok": True, "total": total, "completed": completed,
                        "open": open_count, "adhoc": adhoc, "cols": list(df.columns[:10])})
    except Exception as e:
        return jsonify({"ok": False, "error": str(e)})

@app.route("/preview/incidents", methods=["POST"])
def preview_incidents():
    try:
        df = _read_df(request.files.get("file"))
        if df is None: return jsonify({"ok": False, "error": "Could not read file"})
        total = len(df)
        resolved = carry_fwd = unknown_closed = known_err_closed = unknown_kedb = 0
        i_st = find_col(df, "state", "status")
        i_rc = find_col(df, "resolution code", "resolution_code")
        i_pr = find_col(df, "problem")
        inc_closed = norm(df[i_st]).isin(["closed","resolved"]) if i_st else pd.Series([False]*total)
        known_err  = norm(df[i_rc]).str.contains("known error", na=False) if i_rc else pd.Series([False]*total)
        has_prob   = ~norm(df[i_pr]).isin(BLANKS) if i_pr else pd.Series([False]*total)
        resolved         = int(inc_closed.sum())
        carry_fwd        = int((~inc_closed).sum())
        unknown_closed   = int((inc_closed & ~known_err).sum())
        known_err_closed = int((inc_closed & known_err).sum())
        unknown_kedb     = int((inc_closed & ~known_err & has_prob).sum())
        return jsonify({"ok": True, "total": total, "resolved": resolved,
                        "carry_fwd": carry_fwd, "unknown_closed": unknown_closed,
                        "known_err_closed": known_err_closed, "unknown_kedb": unknown_kedb,
                        "cols": list(df.columns[:10])})
    except Exception as e:
        return jsonify({"ok": False, "error": str(e)})

@app.route("/preview/problems", methods=["POST"])
def preview_problems():
    try:
        df = _read_df(request.files.get("file"))
        if df is None: return jsonify({"ok": False, "error": "Could not read file"})
        total = len(df); resolved = 0
        ps = find_col(df, "state", "status")
        if ps:
            resolved = int(norm(df[ps]).isin({"closed complete","closed - complete","close complete","closed"}).sum())
        else:
            resolved = 0
        return jsonify({"ok": True, "total": total, "resolved": resolved, "cols": list(df.columns[:10])})
    except Exception as e:
        return jsonify({"ok": False, "error": str(e)})

@app.route("/preview/macm", methods=["POST"])
def preview_macm():
    try:
        df = _read_df(request.files.get("file"))
        if df is None: return jsonify({"ok": False, "error": "Could not read file"})
        return jsonify({"ok": True, "total": len(df), "cols": list(df.columns[:10])})
    except Exception as e:
        return jsonify({"ok": False, "error": str(e)})

# ══════════════════════════════════════════════════════════════════════════════
# FINAL PROCESS ROUTE
# ══════════════════════════════════════════════════════════════════════════════
@app.route("/")
def index(): return HTML_PAGE

@app.route("/process", methods=["POST"])
def process():
    try:
        f            = request.files
        prob_skipped = request.form.get("prob_skipped") == "true"
        month_label  = request.form.get("month_label", "Monthly Report").strip() or "Monthly Report"
        dedup_debug  = {}
        filter_debug = {}

        # ── Step 1: Read raw bytes ────────────────────────────────────────────
        def _raw(key):
            fs = f.get(key)
            if fs and fs.filename:
                return fs.filename, fs.read()
            return None, None

        ritm_fname, ritm_raw = _raw("ritm")
        inc_fname,  inc_raw  = _raw("incidents")
        macm_fname, macm_raw = _raw("macm")
        prob_fname, prob_raw = _raw("problems")
        ref_fs = f.get("reference")

        # ── Step 2: Save cycle to disk ────────────────────────────────────────
        raw_files = {}
        if ritm_raw:  raw_files["RITM"]      = (ritm_fname, ritm_raw)
        if inc_raw:   raw_files["Incidents"] = (inc_fname,  inc_raw)
        if macm_raw:  raw_files["MACM"]      = (macm_fname, macm_raw)
        if prob_raw:  raw_files["Problems"]  = (prob_fname, prob_raw)
        cycle_folder = save_upload_cycle(raw_files, month_label) if raw_files else ""

        # ── Step 3: Parse bytes → DataFrames ─────────────────────────────────
        ritm_df = _bytes_to_df(ritm_raw, ritm_fname)
        inc_df  = _bytes_to_df(inc_raw,  inc_fname)
        macm_df = _bytes_to_df(macm_raw, macm_fname)
        prob_df = _bytes_to_df(prob_raw, prob_fname)

        if ritm_df is None: return jsonify({"ok": False, "error": "Could not read RITM file"})
        if inc_df  is None: return jsonify({"ok": False, "error": "Could not read Incidents file"})
        if macm_df is None: return jsonify({"ok": False, "error": "Could not read MACM file"})

        # ── Step 4: Deduplicate — count each unique ticket once ───────────────
        ritm_df, dd_r = dedup_tickets(ritm_df, "RITM")
        inc_df,  dd_i = dedup_tickets(inc_df,  "INC")
        macm_df, dd_m = dedup_tickets(macm_df, "MACM")
        dedup_debug   = {"RITM": dd_r, "INC": dd_i, "MACM": dd_m}
        if prob_df is not None:
            prob_df, dd_p = dedup_tickets(prob_df, "PROB")
            dedup_debug["PROB"] = dd_p

        print("\n[DEDUP]")
        for sh, d in dedup_debug.items():
            print(f"  {sh:8s}  id_col={str(d.get('id_col','?')):30s}  "
                  f"before={d.get('before',0):4d}  dupes={d.get('dupes',0):4d}  "
                  f"after={d.get('after',0):4d}")

        # ── Step 5: Filter to selected month ─────────────────────────────────
        mn, yr = parse_month_year(month_label)
        # "New incidents received" = total unique incidents in the uploaded file
        # (irrespective of state — the user exports only the reporting month)
        inc_received_count = len(inc_df) if inc_df is not None else 0
        filter_debug = {}
        if mn and yr:
            print(f"\n[FILTER] Month={mn}/{yr}")
            ritm_df, fd_r = filter_by_month(ritm_df, mn, yr, "RITM")
            inc_df,  fd_i = filter_by_month(inc_df,  mn, yr, "INC")
            macm_df, fd_m = filter_by_month(macm_df, mn, yr, "MACM")
            filter_debug  = {"RITM": fd_r, "INC": fd_i, "MACM": fd_m}
            if prob_df is not None:
                prob_df, fd_p = filter_by_month(prob_df, mn, yr, "PROB")
                filter_debug["PROB"] = fd_p

            print("\n[FILTER RESULTS]")
            for sh, d in filter_debug.items():
                print(f"  {sh:8s}  col={str(d.get('col','?')):30s}  "
                      f"before={d.get('before',0):4d}  after={d.get('after',0):4d}  "
                      f"note={d.get('note','')}")

        # ── Step 6: Calculate metrics ─────────────────────────────────────────
        mi = {}
        for k in SUMIT_KEYS:
            try: mi[k] = int(float(request.form.get(k) or 0))
            except Exception: mi[k] = 0

        M = calculate_metrics(ritm_df, inc_df, macm_df, prob_df, prob_skipped,
                              inc_received_count=inc_received_count)
        for k, v in mi.items(): M[k] = v if v else None

        if ref_fs and ref_fs.filename:
            ref_bytes = ref_fs.read()
            xl_buf, filled = fill_excel(ref_bytes, M, mi)
            dl_name = ref_fs.filename
            source  = f"Template filled — {filled} rows updated · {dl_name}"
        else:
            xl_buf  = generate_excel(M, mi, month_label)
            dl_name = f"Mainspring_{month_label.replace(' ','_')}.xlsx"
            filled  = 0; source = "New Excel generated"

        excel_b64  = base64.b64encode(xl_buf.read()).decode()
        M_safe     = {k: (v if v is not None else 0) for k, v in M.items()}
        mi_safe    = {k: (v if v is not None else 0) for k, v in mi.items()}
        table_rows = build_table_html(M_safe, mi_safe)
        counts     = {"ritm": len(ritm_df), "inc": len(inc_df),
                      "macm": len(macm_df), "prob": len(prob_df) if prob_df is not None else 0,
                      "month_filtered": bool(mn and yr)}

        return jsonify({"ok": True, "metrics": M_safe, "manual": mi_safe,
                        "excel_b64": excel_b64, "filename": dl_name, "filled": filled,
                        "source": source, "month": month_label,
                        "table_rows": table_rows, "counts": counts,
                        "dedup_debug": dedup_debug, "filter_debug": filter_debug,
                        "saved_folder": os.path.basename(cycle_folder) if cycle_folder else ""})
    except Exception as e:
        import traceback
        return jsonify({"ok": False, "error": str(e), "trace": traceback.format_exc()})

# ══════════════════════════════════════════════════════════════════════════════
# SEND EMAIL ROUTE
# ══════════════════════════════════════════════════════════════════════════════
@app.route("/send_email", methods=["POST"])
def send_email():
    import smtplib, base64 as b64lib
    from email.mime.multipart import MIMEMultipart
    from email.mime.base import MIMEBase
    from email.mime.text import MIMEText
    from email import encoders
    try:
        from_email = request.form.get("from_email", "").strip()
        password   = request.form.get("password",   "").strip()
        to_email   = request.form.get("to_email",   "").strip()
        cc_email   = request.form.get("cc_email",   "").strip()
        month      = request.form.get("month",      "Monthly Report")
        excel_b64  = request.form.get("excel_b64",  "")
        filename   = request.form.get("filename",   "Mainspring_Report.xlsx")

        if not from_email: return jsonify({"ok": False, "error": "Sender email is required"})
        if not password:   return jsonify({"ok": False, "error": "Gmail App Password is required"})
        if not to_email:   return jsonify({"ok": False, "error": "Recipient email is required"})
        if not excel_b64:  return jsonify({"ok": False, "error": "No report data to send"})

        excel_bytes = b64lib.b64decode(excel_b64)

        msg = MIMEMultipart()
        msg["From"]    = from_email
        msg["To"]      = to_email
        msg["Subject"] = f"Mainspring Monthly Report — {month}"
        if cc_email:
            msg["Cc"] = cc_email

        body = (
            f"Hi Sumit,\n\n"
            f"Please find attached the Mainspring Monthly Report for {month}.\n\n"
            f"Report includes:\n"
            f"  • RITM (Standard & Ad hoc Service Requests)\n"
            f"  • Incident Analysis (Known / Unknown / KEDB)\n"
            f"  • MACM Change Requests\n"
            f"  • Problem Ticket Records\n\n"
            f"All yellow cells still require your manual effort values.\n\n"
            f"Regards,\nMainspring Report Automation"
        )
        msg.attach(MIMEText(body, "plain"))

        part = MIMEBase("application", "octet-stream")
        part.set_payload(excel_bytes)
        encoders.encode_base64(part)
        part.add_header("Content-Disposition", f'attachment; filename="{filename}"')
        msg.attach(part)

        recipients = [to_email] + ([cc_email] if cc_email else [])
        with smtplib.SMTP_SSL("smtp.gmail.com", 465, timeout=20) as server:
            server.login(from_email, password)
            server.sendmail(from_email, recipients, msg.as_string())

        return jsonify({"ok": True, "message": f"Email sent to {to_email}"})
    except smtplib.SMTPAuthenticationError:
        return jsonify({"ok": False, "error": "Authentication failed — check your Gmail App Password"})
    except smtplib.SMTPException as e:
        return jsonify({"ok": False, "error": f"SMTP error: {e}"})
    except Exception as e:
        return jsonify({"ok": False, "error": str(e)})

# ══════════════════════════════════════════════════════════════════════════════
# REFILTER ROUTE
# ══════════════════════════════════════════════════════════════════════════════
@app.route("/refilter", methods=["POST"])
def refilter():
    """Re-filter saved session files by a custom date range and return updated metrics."""
    try:
        folder       = request.form.get("folder",       "").strip()
        start_date   = request.form.get("start_date",   "").strip()
        end_date     = request.form.get("end_date",     "").strip()
        prob_skipped = request.form.get("prob_skipped") == "true"
        month_label  = (request.form.get("month_label", "Monthly Report").strip()
                        or "Monthly Report")

        if not folder:
            return jsonify({"ok": False, "error": "No session folder — please re-upload files"})

        folder_path = os.path.join(UPLOAD_DIR, folder)
        if not os.path.isdir(folder_path):
            return jsonify({"ok": False,
                            "error": "Session folder not found — please re-upload files"})

        import warnings
        with warnings.catch_warnings():
            warnings.simplefilter("ignore")
            start_dt = pd.to_datetime(start_date, errors="coerce")
            end_dt   = pd.to_datetime(end_date,   errors="coerce")

        if pd.isna(start_dt) or pd.isna(end_dt):
            return jsonify({"ok": False, "error": "Invalid date range — check start/end dates"})

        if start_dt > end_dt:
            return jsonify({"ok": False, "error": "Start date must be before end date"})

        ritm_df = _load_sheet(folder_path, "RITM")
        inc_df  = _load_sheet(folder_path, "Incidents")
        macm_df = _load_sheet(folder_path, "MACM")
        prob_df = _load_sheet(folder_path, "Problems") if not prob_skipped else None

        if ritm_df is None:
            return jsonify({"ok": False, "error": "RITM file not found — please re-upload"})
        if inc_df is None:
            return jsonify({"ok": False, "error": "Incidents file not found — please re-upload"})
        if macm_df is None:
            return jsonify({"ok": False, "error": "MACM file not found — please re-upload"})

        ritm_df, _ = dedup_tickets(ritm_df, "RITM")
        inc_df,  _ = dedup_tickets(inc_df,  "INC")
        macm_df, _ = dedup_tickets(macm_df, "MACM")
        if prob_df is not None:
            prob_df, _ = dedup_tickets(prob_df, "PROB")

        inc_received_count = len(inc_df) if inc_df is not None else 0

        ritm_df, fd_r = filter_by_daterange(ritm_df, start_dt, end_dt, "RITM")
        inc_df,  fd_i = filter_by_daterange(inc_df,  start_dt, end_dt, "INC")
        macm_df, fd_m = filter_by_daterange(macm_df, start_dt, end_dt, "MACM")

        sheet_counts = {
            "INC":  {"before": fd_i.get("before",0), "after": fd_i.get("after",0),
                     "col": fd_i.get("col") or "?",  "applied": fd_i.get("applied",False)},
            "RITM": {"before": fd_r.get("before",0), "after": fd_r.get("after",0),
                     "col": fd_r.get("col") or "?",  "applied": fd_r.get("applied",False)},
            "MACM": {"before": fd_m.get("before",0), "after": fd_m.get("after",0),
                     "col": fd_m.get("col") or "?",  "applied": fd_m.get("applied",False)},
        }
        if prob_df is not None:
            prob_df, fd_p = filter_by_daterange(prob_df, start_dt, end_dt, "PROB")
            sheet_counts["PROB"] = {
                "before": fd_p.get("before",0), "after": fd_p.get("after",0),
                "col": fd_p.get("col") or "?",  "applied": fd_p.get("applied",False)}

        M = calculate_metrics(ritm_df, inc_df, macm_df, prob_df, prob_skipped,
                              inc_received_count=inc_received_count)
        M_safe     = {k: (v if v is not None else 0) for k, v in M.items()}
        table_rows = build_table_html(M_safe, {})

        return jsonify({"ok": True, "metrics": M_safe, "manual": {},
                        "table_rows": table_rows, "sheet_counts": sheet_counts,
                        "start": start_date, "end": end_date})
    except Exception as e:
        import traceback
        return jsonify({"ok": False, "error": str(e), "trace": traceback.format_exc()})

# ══════════════════════════════════════════════════════════════════════════════
# FRONTEND
# ══════════════════════════════════════════════════════════════════════════════
HTML_PAGE = r"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>Mainspring Report Automation</title>
<link rel="preconnect" href="https://fonts.googleapis.com">
<link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800;900&display=swap" rel="stylesheet">
<script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.0/dist/chart.umd.min.js"></script>
<style>
/* ── Design System ─────────────────────────────────────────────────────────── */
:root{
  --bg:#f3f2f1;--surface:#fff;--border:#e1dfdd;--border-light:#f0eeed;
  --primary:#0078d4;--primary-dark:#005a9e;--primary-light:#eff6ff;
  --dark:#201f1e;--dark2:#323130;--muted:#605e5c;--faint:#a19f9d;
  --green:#107c10;--green-bg:#f1faf1;--green-border:#c8e6c9;
  --red:#a4262c;--red-bg:#fdf3f4;--red-border:#f9c8ca;
  --orange:#d83b01;--orange-bg:#fff4ce;--orange-border:#fce100;
  --purple:#8764b8;--purple-bg:#f4f0fa;--purple-dark:#6b4f9e;
  --shadow-sm:0 1px 4px rgba(0,0,0,.08);
  --shadow:0 2px 12px rgba(0,0,0,.09);
  --shadow-lg:0 6px 28px rgba(0,0,0,.13);
  --radius:8px;--radius-lg:12px;--radius-xl:16px;
}
*,*::before,*::after{box-sizing:border-box;margin:0;padding:0}
html{scroll-behavior:smooth}
body{font-family:'Inter',system-ui,-apple-system,sans-serif;background:var(--bg);
  color:var(--dark2);min-height:100vh;font-size:14px;line-height:1.5;
  -webkit-font-smoothing:antialiased}
.wrap{max-width:920px;margin:0 auto;padding:0 1.25rem 2rem;transition:max-width .4s cubic-bezier(.4,0,.2,1)}
.wrap.dash-mode{max-width:1520px}

/* ── topbar ── */
.topbar{background:#1e3a5f;border-radius:12px;padding:.75rem 1.5rem;
  display:flex;align-items:center;justify-content:space-between;margin-bottom:1.2rem}
.tb-title{color:#fff;font-size:1rem;font-weight:700;display:flex;align-items:center;gap:.5rem}
.tb-badge{background:#2563eb;color:#fff;font-size:.66rem;font-weight:800;
  border-radius:6px;padding:.18rem .48rem;letter-spacing:.5px}
.btn-reset{background:rgba(255,255,255,.13);color:#fff;border:1px solid rgba(255,255,255,.28);
  border-radius:7px;padding:.34rem .9rem;font-size:.78rem;font-weight:600;cursor:pointer;
  transition:all .15s;display:none}
.btn-reset:hover{background:rgba(255,255,255,.22)}
.btn-reset.show{display:block}

/* ── progress bar ── */
.prog-wrap{background:#fff;border-radius:12px;box-shadow:0 1px 8px rgba(0,0,0,.08);
  padding:.95rem 1.3rem 1rem;margin-bottom:1.2rem}
.prog-label{font-size:.65rem;font-weight:800;text-transform:uppercase;letter-spacing:1px;
  color:#94a3b8;margin-bottom:.75rem}
.prog-row{display:flex;align-items:flex-start}
.pstep{display:flex;flex-direction:column;align-items:center;gap:.18rem;flex:0 0 auto;min-width:72px}
.pline{flex:1;height:2px;background:#e2e8f0;margin-top:17px;transition:background .3s}
.pline.done{background:#16a34a}
.pc{width:36px;height:36px;border-radius:50%;border:2.5px solid #e2e8f0;background:#f8fafc;
  color:#94a3b8;font-size:.8rem;font-weight:800;display:flex;align-items:center;justify-content:center;
  transition:all .3s;flex-shrink:0}
.pn{font-size:.67rem;font-weight:700;color:#94a3b8;text-align:center;transition:color .3s;white-space:nowrap}
.ps{font-size:.6rem;color:#cbd5e1;text-align:center;min-height:13px;transition:color .3s}
.pstep.done .pc{background:#16a34a;border-color:#16a34a;color:#fff;box-shadow:0 0 0 3px rgba(22,163,74,.13)}
.pstep.done .pn{color:#16a34a}.pstep.done .ps{color:#16a34a}
.pstep.active .pc{background:#2563eb;border-color:#2563eb;color:#fff;box-shadow:0 0 0 4px rgba(37,99,235,.16)}
.pstep.active .pn{color:#2563eb;font-weight:800}.pstep.active .ps{color:#2563eb}

/* ── step panel ── */
.step-panel{display:none}.step-panel.active{display:block}

/* ── step card ── */
.scard{background:#fff;border-radius:14px;border-left:4px solid #2563eb;
  box-shadow:0 2px 14px rgba(0,0,0,.08)}
.scard-hdr{padding:1rem 1.5rem .8rem;display:flex;align-items:center;gap:.75rem;
  border-bottom:1px solid #f1f5f9}
.scard-ico{width:42px;height:42px;border-radius:11px;background:#eff6ff;
  display:flex;align-items:center;justify-content:center;font-size:1.25rem;flex-shrink:0}
.scard-title{font-size:.97rem;font-weight:700;color:#1e293b}
.scard-sub{font-size:.73rem;color:#64748b;margin-top:2px}
.scard-body{padding:1.3rem 1.5rem 1.5rem}

/* ── upload zone ── */
.uzone{border:2px dashed #93c5fd;border-radius:11px;background:#f8fbff;
  min-height:180px;display:flex;flex-direction:column;align-items:center;
  justify-content:center;cursor:pointer;transition:all .22s;
  padding:1.6rem 1.5rem;text-align:center;position:relative;gap:.35rem}
.uzone:hover,.uzone.over{border-color:#2563eb;background:#eff6ff}
.uzone.done{border-color:#16a34a;background:#f0fdf4}
.uzone input{position:absolute;inset:0;opacity:0;cursor:pointer;width:100%;height:100%}
.uz-ico{font-size:3rem;color:#93c5fd;line-height:1;transition:color .2s}
.uzone.done .uz-ico{color:#16a34a}
.uz-title{font-size:.93rem;font-weight:600;color:#334155;margin-top:.25rem}
.uz-hint{font-size:.74rem;color:#94a3b8}

/* ── file badge row ── */
.file-row{display:none;align-items:center;gap:.6rem;padding:.6rem .85rem;
  background:#f8fafc;border:1px solid #e2e8f0;border-radius:8px;margin-top:.8rem}
.file-row.show{display:flex}
.fr-ico{font-size:1rem;color:#2563eb}
.fr-name{flex:1;font-size:.82rem;font-weight:600;color:#334155;
  white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
.fr-sz{background:#dcfce7;color:#166534;font-size:.67rem;font-weight:700;
  border-radius:5px;padding:.13rem .42rem;flex-shrink:0}

/* ── status bar ── */
.sbar{display:none;align-items:center;gap:.5rem;padding:.6rem .95rem;
  background:#f0fdf4;border:1.5px solid #86efac;border-radius:8px;
  margin-top:.7rem;font-size:.83rem;font-weight:700;color:#166534}
.sbar.show{display:flex}
.sbar.err{background:#fef2f2;border-color:#fca5a5;color:#991b1b}
.skip-bar{display:none;align-items:center;gap:.5rem;padding:.6rem .95rem;
  background:#fef9c3;border:1.5px solid #fde047;border-radius:8px;
  margin-top:.7rem;font-size:.83rem;font-weight:700;color:#854d0e}
.skip-bar.show{display:flex}
.cols-note{font-size:.69rem;color:#94a3b8;margin-top:.5rem}
.cols-note b{color:#64748b}

/* ── action row ── */
.act-row{display:flex;align-items:center;justify-content:space-between;
  margin-top:1.1rem;gap:.6rem;flex-wrap:wrap}
.act-left{display:flex;gap:.5rem;flex-wrap:wrap}

/* ── buttons ── */
.btn-p{background:#2563eb;color:#fff;border:none;border-radius:9px;
  padding:.62rem 1.5rem;font-size:.88rem;font-weight:700;cursor:pointer;
  box-shadow:0 2px 8px rgba(37,99,235,.28);transition:all .18s;
  display:inline-flex;align-items:center;gap:.4rem}
.btn-p:hover:not(:disabled){background:#1d4ed8;transform:translateY(-1px)}
.btn-p:disabled{background:#93c5fd;cursor:not-allowed;transform:none;box-shadow:none}
.btn-next{background:#16a34a;color:#fff;border:none;border-radius:9px;
  padding:.6rem 1.4rem;font-size:.87rem;font-weight:700;cursor:pointer;
  box-shadow:0 2px 8px rgba(22,163,74,.22);transition:all .18s;
  display:none;align-items:center;gap:.4rem}
.btn-next:hover{background:#15803d;transform:translateY(-1px)}
.btn-next.show{display:inline-flex}
.btn-back{background:#fff;color:#475569;border:1.5px solid #e2e8f0;border-radius:9px;
  padding:.58rem 1.1rem;font-size:.85rem;font-weight:600;cursor:pointer;transition:all .15s}
.btn-back:hover{background:#f8fafc}
.btn-skip{background:#fefce8;color:#92400e;border:1.5px solid #fde047;border-radius:9px;
  padding:.55rem 1.05rem;font-size:.82rem;font-weight:600;cursor:pointer;transition:all .15s;
  display:inline-flex;align-items:center;gap:.35rem}
.btn-skip:hover{background:#fef9c3}

/* ── manual inputs ── */
.manual-box{background:#fff;border-radius:12px;border:1px solid #e2e8f0;
  box-shadow:0 1px 6px rgba(0,0,0,.05);overflow:hidden;margin-bottom:1rem}
.mtoggle{width:100%;background:none;border:none;padding:.82rem 1.4rem;
  display:flex;align-items:center;justify-content:space-between;cursor:pointer;
  font-size:.86rem;font-weight:700;color:#334155}
.mtoggle:hover{background:#f8fafc}
.mbody{padding:1rem 1.4rem 1.3rem;display:none}
.mbody.open{display:block}
.fgrid{display:grid;grid-template-columns:1fr 1fr;gap:.72rem}
.fg{display:flex;flex-direction:column;gap:.2rem}
.fg label{font-size:.73rem;font-weight:600;color:#475569}
.fg input,.fg select{border:1px solid #e2e8f0;border-radius:7px;
  padding:.4rem .62rem;font-size:.83rem;color:#1e293b;outline:none;width:100%}
.fg input:focus,.fg select:focus{border-color:#2563eb;box-shadow:0 0 0 3px rgba(37,99,235,.1)}
.mrow{display:grid;grid-template-columns:2fr 1fr;gap:.72rem;margin-bottom:.72rem}

/* ── error ── */
.err-box{background:#fef2f2;border:1px solid #fca5a5;border-left:4px solid #dc2626;
  border-radius:9px;padding:.62rem 1rem;color:#7f1d1d;font-size:.82rem;margin:.5rem 0;display:none}
.err-box.show{display:block}

/* ── spinner ── */
.spin-ov{display:none;position:fixed;inset:0;background:rgba(240,244,248,.88);
  z-index:999;flex-direction:column;align-items:center;justify-content:center;gap:.85rem}
.spin-ov.show{display:flex}
.spinner{width:42px;height:42px;border:4px solid #bfdbfe;border-top-color:#2563eb;
  border-radius:50%;animation:spin .7s linear infinite}
@keyframes spin{to{transform:rotate(360deg)}}
.spin-txt{font-size:.88rem;font-weight:600;color:#1e40af}

/* ══ ALL FILES PROCESSED card ══ */
#done-card{display:none;background:#fff;border-radius:16px;
  box-shadow:0 4px 24px rgba(0,0,0,.1);padding:2.5rem 2rem;text-align:center;margin-bottom:1.2rem}
#done-card.show{display:block}
.done-circle{width:80px;height:80px;border-radius:50%;background:#16a34a;
  display:flex;align-items:center;justify-content:center;margin:0 auto 1.2rem;
  font-size:2.2rem;color:#fff;box-shadow:0 4px 20px rgba(22,163,74,.35)}
.done-title{font-size:1.5rem;font-weight:800;color:#1e293b;margin-bottom:.4rem}
.done-sub{font-size:.85rem;color:#64748b;margin-bottom:1.4rem}
.done-badges{display:flex;justify-content:center;gap:.75rem;flex-wrap:wrap;margin-bottom:1.6rem}
.dbadge{display:inline-flex;align-items:center;gap:.4rem;border-radius:20px;
  padding:.4rem .95rem;font-size:.88rem;font-weight:700;border:2px solid}
.dbadge.ritm{background:#eff6ff;color:#1d4ed8;border-color:#bfdbfe}
.dbadge.inc{background:#fef9c3;color:#854d0e;border-color:#fde047}
.dbadge.macm{background:#f0fdf4;color:#166534;border-color:#86efac}
.dbadge.prob{background:#fff7ed;color:#9a3412;border-color:#fed7aa}
.done-actions{display:flex;justify-content:center;gap:.75rem;flex-wrap:wrap}
.btn-dashboard{background:#2563eb;color:#fff;border:none;border-radius:10px;
  padding:.72rem 1.6rem;font-size:.92rem;font-weight:700;cursor:pointer;
  box-shadow:0 2px 10px rgba(37,99,235,.3);transition:all .18s;
  display:inline-flex;align-items:center;gap:.45rem}
.btn-dashboard:hover{background:#1d4ed8;transform:translateY(-1px)}
.btn-dl-done{background:#fff;color:#166534;border:2px solid #86efac;border-radius:10px;
  padding:.7rem 1.5rem;font-size:.9rem;font-weight:700;cursor:pointer;transition:all .18s;
  display:inline-flex;align-items:center;gap:.4rem}
.btn-dl-done:hover{background:#f0fdf4}
.btn-start-over{background:#fff;color:#64748b;border:1.5px solid #e2e8f0;border-radius:10px;
  padding:.68rem 1.3rem;font-size:.88rem;font-weight:600;cursor:pointer;transition:all .15s}
.btn-start-over:hover{background:#f8fafc}

/* ══ DASHBOARD ══ */
#dashboard{display:none}
#dashboard.show{display:block}

/* toolbar */
.dash-toolbar{display:flex;align-items:center;justify-content:space-between;
  background:#fff;border-radius:14px;padding:.85rem 1.4rem;
  box-shadow:0 2px 12px rgba(0,0,0,.08);margin-bottom:1rem;gap:1rem;flex-wrap:wrap;
  border-left:5px solid #2563eb}
.dt-left{display:flex;flex-direction:column;gap:.12rem}
.dt-title{font-size:1rem;font-weight:800;color:#1e293b;display:flex;align-items:center;gap:.45rem}
.dt-sub{font-size:.73rem;color:#64748b}
.dt-right{display:flex;gap:.5rem;flex-wrap:wrap;align-items:center}
.btn-dl{background:#16a34a;color:#fff;border:none;border-radius:9px;
  padding:.5rem 1rem;font-size:.83rem;font-weight:700;cursor:pointer;
  display:inline-flex;align-items:center;gap:.35rem;transition:all .14s;
  box-shadow:0 2px 6px rgba(22,163,74,.25)}
.btn-dl:hover{background:#15803d;transform:translateY(-1px)}
.btn-csv{background:#fff;color:#334155;border:1.5px solid #e2e8f0;border-radius:9px;
  padding:.48rem .9rem;font-size:.81rem;font-weight:600;cursor:pointer;transition:all .14s}
.btn-csv:hover{background:#f8fafc}
.btn-email-dash{background:#7c3aed;color:#fff;border:none;border-radius:9px;
  padding:.5rem 1rem;font-size:.83rem;font-weight:700;cursor:pointer;
  display:inline-flex;align-items:center;gap:.35rem;transition:all .14s;
  box-shadow:0 2px 6px rgba(124,58,237,.25)}
.btn-email-dash:hover{background:#6d28d9;transform:translateY(-1px)}
.btn-email-done{background:#7c3aed;color:#fff;border:none;border-radius:10px;
  padding:.7rem 1.4rem;font-size:.9rem;font-weight:700;cursor:pointer;transition:all .18s;
  display:inline-flex;align-items:center;gap:.4rem;box-shadow:0 2px 8px rgba(124,58,237,.25)}
.btn-email-done:hover{background:#6d28d9;transform:translateY(-1px)}

/* dash hero banner */
.dash-banner{background:linear-gradient(135deg,#0f2744 0%,#1e3a5f 50%,#15302a 100%);
  border-radius:16px;padding:1.4rem 1.8rem;display:flex;justify-content:space-between;
  align-items:center;margin-bottom:1rem;position:relative;overflow:hidden}
.dash-banner::after{content:'';position:absolute;right:-40px;top:-40px;
  width:200px;height:200px;border-radius:50%;
  background:rgba(255,255,255,.04);pointer-events:none}
.dash-banner h1{color:#fff;font-size:1.25rem;font-weight:800;margin:0;
  display:flex;align-items:center;gap:.5rem}
.dash-banner p{color:#93c5fd;font-size:.76rem;margin:.22rem 0 0}
.dash-month-block{text-align:right}
.dash-month{color:#fff;font-size:1.1rem;font-weight:800;letter-spacing:.3px}
.dash-msub{color:#6ee7b7;font-size:.7rem;margin-top:.18rem;letter-spacing:.3px}

/* section headers */
.sec-head{display:flex;align-items:center;gap:.5rem;font-size:.78rem;font-weight:800;
  color:#1e293b;text-transform:uppercase;letter-spacing:.8px;
  padding:.5rem 0 .4rem;margin:1.1rem 0 .6rem;
  border-bottom:2px solid #e2e8f0;position:relative}
.sec-head::before{content:'';position:absolute;bottom:-2px;left:0;
  width:3rem;height:2px;background:#2563eb;border-radius:2px}

/* ── Date Range Filter ──────────────────────────────────────────────────────── */
.df-card{background:#fff;border-radius:14px;border:1.5px solid #bae6fd;padding:1rem 1.2rem;margin-bottom:1rem;box-shadow:0 2px 10px rgba(3,105,161,.08);}
.df-hdr{display:flex;align-items:center;justify-content:space-between;margin-bottom:.75rem;gap:.5rem;flex-wrap:wrap;}
.df-title{font-size:.88rem;font-weight:800;color:#0369a1;display:flex;align-items:center;gap:.45rem;}
.df-badge{background:#0369a1;color:#fff;font-size:.65rem;font-weight:700;padding:.2rem .55rem;border-radius:20px;letter-spacing:.5px;white-space:nowrap;}
.df-badge.active{background:#16a34a;}
.df-controls{display:flex;gap:.65rem;align-items:flex-end;flex-wrap:wrap;}
.df-grp{display:flex;flex-direction:column;gap:.22rem;}
.df-lbl{font-size:.68rem;font-weight:700;color:#64748b;text-transform:uppercase;letter-spacing:.5px;}
.df-input{padding:.38rem .65rem;border-radius:8px;border:1.5px solid #bae6fd;font-size:.83rem;font-weight:600;color:#0369a1;background:#f0f9ff;outline:none;cursor:pointer;min-width:135px;}
.df-input:focus{border-color:#0369a1;box-shadow:0 0 0 2px rgba(3,105,161,.12);}
.df-apply{background:#0369a1;color:#fff;border:none;padding:.42rem 1.1rem;border-radius:8px;font-size:.82rem;font-weight:700;cursor:pointer;transition:.15s;white-space:nowrap;}
.df-apply:hover:not(:disabled){background:#0c4a6e;}
.df-apply:disabled{background:#94a3b8;cursor:not-allowed;}
.df-reset-btn{background:#e2e8f0;color:#475569;border:none;padding:.42rem .9rem;border-radius:8px;font-size:.82rem;font-weight:600;cursor:pointer;transition:.15s;white-space:nowrap;}
.df-reset-btn:hover{background:#cbd5e1;}
.df-breakdown{margin-top:.85rem;padding-top:.75rem;border-top:1px dashed #bae6fd;}
.df-range-lbl{font-size:.74rem;font-weight:700;color:#0369a1;margin-bottom:.5rem;}
.df-sheets{display:flex;gap:.55rem;flex-wrap:wrap;}
.df-sheet{background:#f0f9ff;border:1px solid #bae6fd;border-radius:10px;padding:.5rem .9rem;min-width:130px;}
.df-sheet-name{font-size:.78rem;font-weight:800;color:#0369a1;}
.df-sheet-rows{font-size:.8rem;color:#1e293b;margin-top:.15rem;}
.df-sheet-col{font-size:.68rem;color:#64748b;margin-top:.1rem;}
.df-sheet.warn .df-sheet-name{color:#b45309;}
.df-sheet.warn{background:#fff7ed;border-color:#fcd34d;}
.df-spin{font-size:.77rem;color:#0369a1;font-style:italic;margin-top:.5rem;}
.df-ferr{font-size:.77rem;color:#dc2626;background:#fef2f2;border:1px solid #fecaca;border-radius:7px;padding:.38rem .65rem;margin-top:.4rem;}

/* KPI grid */
.kpi-row{display:grid;grid-template-columns:repeat(7,1fr);gap:.8rem;margin-bottom:1rem}
.kpi{background:#fff;border-radius:14px;padding:1rem .6rem;text-align:center;
  box-shadow:0 2px 12px rgba(0,0,0,.07);border-top:4px solid transparent;
  transition:transform .18s,box-shadow .18s;cursor:default;position:relative;overflow:hidden}
.kpi:hover{transform:translateY(-3px);box-shadow:0 6px 20px rgba(0,0,0,.11)}
.kpi::after{content:'';position:absolute;bottom:0;left:0;right:0;height:3px;
  background:inherit;opacity:.12}
.kpi-ico{font-size:1.4rem;line-height:1;margin-bottom:.3rem}
.kpi-val{font-size:2.1rem;font-weight:900;line-height:1;letter-spacing:-1px}
.kpi-lbl{font-size:.59rem;color:#64748b;margin-top:.3rem;text-transform:uppercase;
  letter-spacing:.5px;font-weight:600;line-height:1.3}

/* charts */
.chart-row{display:grid;gap:1rem;margin-bottom:1rem}
.chart-row.c3{grid-template-columns:repeat(3,1fr)}
.chart-row.c3b{grid-template-columns:repeat(3,1fr)}
.chart-card{background:#fff;border-radius:14px;padding:1.1rem 1.1rem .9rem;
  box-shadow:0 2px 12px rgba(0,0,0,.07);height:320px;position:relative;overflow:hidden;
  transition:box-shadow .18s}
.chart-card:hover{box-shadow:0 4px 20px rgba(0,0,0,.11)}
.chart-card canvas{display:block!important}

/* metrics table */
.tbl-wrap{overflow-x:auto;border-radius:12px;border:1px solid #e2e8f0;
  margin-top:.4rem;box-shadow:0 1px 6px rgba(0,0,0,.05)}
table.mtbl{width:100%;border-collapse:collapse;font-size:.78rem}
table.mtbl th{background:#1e3a5f;color:#fff;font-weight:700;padding:9px 12px;text-align:left;
  font-size:.76rem;letter-spacing:.3px}
table.mtbl th.thv{text-align:center;width:72px}
table.mtbl td{padding:6px 12px;border:1px solid #e2e8f0;vertical-align:middle}
table.mtbl .vc{text-align:center;font-weight:800;font-size:.88rem}
table.mtbl .cc{background:#f0fdf4;color:#374151;font-style:italic;font-size:.72rem;color:#15803d}
table.mtbl .su{background:#fef9c3;font-weight:700;color:#854d0e}
table.mtbl .er{height:4px;background:#f8fafc;border:none}
table.mtbl .sr td{background:#1e3a5f;color:#fff;font-weight:700;padding:6px 12px;font-size:.76rem}
table.mtbl tbody tr:hover td{background:#f8fafc}
table.mtbl .nc{background:#fff}

/* ── email modal ── */
.modal-ov{display:none;position:fixed;inset:0;background:rgba(15,23,42,.6);
  z-index:2000;align-items:center;justify-content:center;backdrop-filter:blur(4px)}
.modal-ov.show{display:flex}
.email-modal{background:#fff;border-radius:18px;width:100%;max-width:480px;
  box-shadow:0 20px 60px rgba(0,0,0,.25);overflow:hidden;margin:1rem}
.em-hdr{background:linear-gradient(135deg,#7c3aed,#4f46e5);padding:1.2rem 1.5rem;
  display:flex;align-items:center;justify-content:space-between}
.em-title{color:#fff;font-size:1rem;font-weight:800;display:flex;align-items:center;gap:.5rem}
.em-close{background:rgba(255,255,255,.2);border:none;color:#fff;border-radius:50%;
  width:28px;height:28px;cursor:pointer;font-size:.9rem;font-weight:700;
  display:flex;align-items:center;justify-content:center;transition:background .15s}
.em-close:hover{background:rgba(255,255,255,.35)}
.em-body{padding:1.3rem 1.5rem;display:flex;flex-direction:column;gap:.85rem}
.em-footer{padding:.9rem 1.5rem 1.3rem;border-top:1px solid #f1f5f9;
  display:flex;flex-direction:column;gap:.75rem}
.em-note{font-size:.71rem;color:#94a3b8;line-height:1.5;
  background:#f8fafc;border-radius:8px;padding:.6rem .85rem;border:1px solid #e2e8f0}
.em-btns{display:flex;gap:.6rem;justify-content:flex-end}
.btn-email-send{background:linear-gradient(135deg,#7c3aed,#4f46e5);color:#fff;border:none;
  border-radius:9px;padding:.6rem 1.4rem;font-size:.88rem;font-weight:700;cursor:pointer;
  display:inline-flex;align-items:center;gap:.4rem;transition:all .15s;
  box-shadow:0 2px 8px rgba(124,58,237,.3)}
.btn-email-send:hover:not(:disabled){opacity:.9;transform:translateY(-1px)}
.btn-email-send:disabled{opacity:.55;cursor:not-allowed;transform:none}
.em-status{padding:0 1.5rem .9rem;font-size:.82rem;font-weight:600;min-height:0;
  transition:all .2s}
.em-status.ok{color:#166534;background:#f0fdf4;padding:.55rem 1.5rem;border-top:1px solid #bbf7d0}
.em-status.err{color:#991b1b;background:#fef2f2;padding:.55rem 1.5rem;border-top:1px solid #fecaca}

@media(max-width:900px){
  .kpi-row{grid-template-columns:repeat(4,1fr)}
  .chart-row.c3,.chart-row.c3b{grid-template-columns:1fr 1fr}
}
@media(max-width:640px){
  .fgrid{grid-template-columns:1fr}
  .kpi-row{grid-template-columns:repeat(2,1fr)}
  .chart-row.c3,.chart-row.c3b{grid-template-columns:1fr}
  .pstep{min-width:52px}.pn{font-size:.58rem}
  .dash-banner{flex-direction:column;gap:.5rem;text-align:center}
  .dash-month-block{text-align:center}
  .dash-toolbar{flex-direction:column;align-items:flex-start}
}
</style>
</head>
<body>
<div class="wrap">

<!-- topbar -->
<div class="topbar">
  <div class="tb-title">
    <span class="tb-badge">MS</span>
    Mainspring Report Automation
  </div>
  <button class="btn-reset" id="btn-reset" onclick="resetAll()">&#8635; Reset</button>
</div>

<!-- ══ WIZARD ══════════════════════════════════════════════════════════════ -->
<div id="wizard" style="max-width:860px;margin:0 auto">

  <!-- progress -->
  <div class="prog-wrap">
    <div class="prog-label">Upload Progress</div>
    <div class="prog-row">
      <div class="pstep active" id="ps1"><div class="pc" id="pc1">1</div><div class="pn">RITM</div><div class="ps" id="psub1">Waiting</div></div>
      <div class="pline" id="pl1"></div>
      <div class="pstep" id="ps2"><div class="pc" id="pc2">2</div><div class="pn">Incident</div><div class="ps" id="psub2">Waiting</div></div>
      <div class="pline" id="pl2"></div>
      <div class="pstep" id="ps3"><div class="pc" id="pc3">3</div><div class="pn">MACM</div><div class="ps" id="psub3">Waiting</div></div>
      <div class="pline" id="pl3"></div>
      <div class="pstep" id="ps4"><div class="pc" id="pc4">4</div><div class="pn">Problem Tickets</div><div class="ps" id="psub4">Waiting</div></div>
      <div class="pline" id="pl4"></div>
      <div class="pstep" id="ps5"><div class="pc" id="pc5">5</div><div class="pn">Reference</div><div class="ps" id="psub5">Waiting</div></div>
      <div class="pline" id="pl5"></div>
      <div class="pstep" id="ps6"><div class="pc" id="pc6">6</div><div class="pn">Results</div><div class="ps" id="psub6">Waiting</div></div>
    </div>
  </div>

  <!-- STEP 1: RITM -->
  <div class="step-panel active" id="panel-1">
    <div class="scard">
      <div class="scard-hdr">
        <div class="scard-ico">&#128203;</div>
        <div>
          <div class="scard-title">Step 1 &#8212; Upload RITM File</div>
          <div class="scard-sub">ServiceNow RITM (Request Item) export &nbsp;&#183;&nbsp; State &amp; Short Description columns needed</div>
        </div>
      </div>
      <div class="scard-body">
        <div style="background:#f0f9ff;border:1.5px solid #bae6fd;border-radius:10px;padding:.85rem 1.1rem;margin-bottom:1rem;display:flex;align-items:center;gap:1rem;flex-wrap:wrap">
          <div style="font-size:.72rem;font-weight:800;color:#0369a1;text-transform:uppercase;letter-spacing:.8px;white-space:nowrap">&#128197; Report Period</div>
          <div style="display:flex;gap:.65rem;align-items:center;flex-wrap:wrap">
            <select id="month-select" style="padding:.38rem .7rem;border-radius:7px;border:1.5px solid #bae6fd;font-size:.85rem;font-weight:600;background:#fff;color:#0369a1">
              <option>January</option><option>February</option><option>March</option>
              <option>April</option><option>May</option><option>June</option>
              <option>July</option><option>August</option><option>September</option>
              <option>October</option><option>November</option><option>December</option>
            </select>
            <input type="number" id="year-input" value="2026" min="2020" max="2040"
              style="padding:.38rem .7rem;border-radius:7px;border:1.5px solid #bae6fd;font-size:.85rem;font-weight:600;background:#fff;color:#0369a1;width:90px">
          </div>
          <div style="font-size:.7rem;color:#0369a1;opacity:.7">Select the reporting month before uploading files</div>
        </div>
        <label class="uzone" id="dz1" ondragover="dzDrag(event,this)" ondragleave="dzLeave(this)" ondrop="dzDrop(event,this,'f1','dz1')">
          <input type="file" id="f1" accept=".xlsx,.xls,.csv" onchange="fileChosen(this,'dz1','f1','ritm')">
          <div class="uz-ico">&#9729;</div>
          <div class="uz-title">Drag &amp; drop your RITM file here</div>
          <div class="uz-hint">or click to browse &nbsp;&#183;&nbsp; Excel &nbsp;&#183;&nbsp; CSV</div>
        </label>
        <div class="file-row" id="fr-ritm"><span class="fr-ico">&#9745;</span><span class="fr-name" id="fn-ritm"></span><span class="fr-sz" id="fs-ritm"></span></div>
        <div class="sbar" id="sb-ritm"></div>
        <div class="cols-note" id="cn-ritm"></div>
        <div class="act-row">
          <div class="act-left">
            <button class="btn-p" id="pb-ritm" onclick="processStep('ritm')" disabled>&#9654; Process RITM</button>
          </div>
          <button class="btn-next" id="nx1" onclick="goStep(2)">Next: Incidents &#8594;</button>
        </div>
      </div>
    </div>
  </div>

  <!-- STEP 2: INCIDENTS -->
  <div class="step-panel" id="panel-2">
    <div class="scard">
      <div class="scard-hdr">
        <div class="scard-ico">&#127903;</div>
        <div>
          <div class="scard-title">Step 2 &#8212; Upload Incident File</div>
          <div class="scard-sub">ServiceNow Incident export &nbsp;&#183;&nbsp; State, Resolution Code, Problem columns needed</div>
        </div>
      </div>
      <div class="scard-body">
        <label class="uzone" id="dz2" ondragover="dzDrag(event,this)" ondragleave="dzLeave(this)" ondrop="dzDrop(event,this,'f2','dz2')">
          <input type="file" id="f2" accept=".xlsx,.xls,.csv" onchange="fileChosen(this,'dz2','f2','inc')">
          <div class="uz-ico">&#9729;</div>
          <div class="uz-title">Drag &amp; drop your Incident file here</div>
          <div class="uz-hint">or click to browse &nbsp;&#183;&nbsp; Excel &nbsp;&#183;&nbsp; CSV</div>
        </label>
        <div class="file-row" id="fr-inc"><span class="fr-ico">&#9745;</span><span class="fr-name" id="fn-inc"></span><span class="fr-sz" id="fs-inc"></span></div>
        <div class="sbar" id="sb-inc"></div>
        <div class="cols-note" id="cn-inc"></div>
        <div class="act-row">
          <div class="act-left">
            <button class="btn-back" onclick="goStep(1)">&#8592; Back</button>
            <button class="btn-p" id="pb-inc" onclick="processStep('inc')" disabled>&#9654; Process Incidents</button>
          </div>
          <button class="btn-next" id="nx2" onclick="goStep(3)">Next: MACM &#8594;</button>
        </div>
      </div>
    </div>
  </div>

  <!-- STEP 3: MACM -->
  <div class="step-panel" id="panel-3">
    <div class="scard">
      <div class="scard-hdr">
        <div class="scard-ico">&#128295;</div>
        <div>
          <div class="scard-title">Step 3 &#8212; Upload MACM File</div>
          <div class="scard-sub">ServiceNow Change (MACM) export &nbsp;&#183;&nbsp; Number, State columns needed</div>
        </div>
      </div>
      <div class="scard-body">
        <label class="uzone" id="dz3" ondragover="dzDrag(event,this)" ondragleave="dzLeave(this)" ondrop="dzDrop(event,this,'f3','dz3')">
          <input type="file" id="f3" accept=".xlsx,.xls,.csv" onchange="fileChosen(this,'dz3','f3','macm')">
          <div class="uz-ico">&#9729;</div>
          <div class="uz-title">Drag &amp; drop your MACM file here</div>
          <div class="uz-hint">or click to browse &nbsp;&#183;&nbsp; Excel &nbsp;&#183;&nbsp; CSV</div>
        </label>
        <div class="file-row" id="fr-macm"><span class="fr-ico">&#9745;</span><span class="fr-name" id="fn-macm"></span><span class="fr-sz" id="fs-macm"></span></div>
        <div class="sbar" id="sb-macm"></div>
        <div class="cols-note" id="cn-macm"></div>
        <div class="act-row">
          <div class="act-left">
            <button class="btn-back" onclick="goStep(2)">&#8592; Back</button>
            <button class="btn-p" id="pb-macm" onclick="processStep('macm')" disabled>&#9654; Process MACM</button>
          </div>
          <button class="btn-next" id="nx3" onclick="goStep(4)">Next: Problem Tickets &#8594;</button>
        </div>
      </div>
    </div>
  </div>

  <!-- STEP 4: PROBLEM TICKETS -->
  <div class="step-panel" id="panel-4">
    <div class="scard">
      <div class="scard-hdr">
        <div class="scard-ico">&#128196;</div>
        <div>
          <div class="scard-title">Step 4 &#8212; Upload Problem Tickets File</div>
          <div class="scard-sub">ServiceNow Problem export &nbsp;&#183;&nbsp; State column needed &nbsp;&#183;&nbsp; or skip if none this month</div>
        </div>
      </div>
      <div class="scard-body">
        <label class="uzone" id="dz4" ondragover="dzDrag(event,this)" ondragleave="dzLeave(this)" ondrop="dzDrop(event,this,'f4','dz4')">
          <input type="file" id="f4" accept=".xlsx,.xls,.csv" onchange="fileChosen(this,'dz4','f4','prob')">
          <div class="uz-ico">&#9729;</div>
          <div class="uz-title">Drag &amp; drop your Problem Tickets file here</div>
          <div class="uz-hint">or click to browse &nbsp;&#183;&nbsp; Excel &nbsp;&#183;&nbsp; CSV</div>
        </label>
        <div class="file-row" id="fr-prob"><span class="fr-ico">&#9745;</span><span class="fr-name" id="fn-prob"></span><span class="fr-sz" id="fs-prob"></span></div>
        <div class="sbar" id="sb-prob"></div>
        <div class="skip-bar" id="skip-bar">&#9940; Skipped &#8212; No problem tickets this month</div>
        <div class="cols-note" id="cn-prob"></div>
        <div class="act-row">
          <div class="act-left">
            <button class="btn-back" onclick="goStep(3)">&#8592; Back</button>
            <button class="btn-p" id="pb-prob" onclick="processStep('prob')" disabled>&#9654; Process Problems</button>
            <button class="btn-skip" onclick="skipProblems()">&#9940; Skip</button>
          </div>
          <button class="btn-next" id="nx4" onclick="goStep(5)">Next: Reference &#8594;</button>
        </div>
      </div>
    </div>
  </div>

  <!-- STEP 5: REFERENCE + MANUAL -->
  <div class="step-panel" id="panel-5">
    <div class="scard">
      <div class="scard-hdr">
        <div class="scard-ico">&#128202;</div>
        <div>
          <div class="scard-title">Step 5 &#8212; Reference Excel &amp; Finalise</div>
          <div class="scard-sub">Upload your monthly template, set report month, add effort values, then generate</div>
        </div>
      </div>
      <div class="scard-body">


        <div style="margin-bottom:1rem">
          <div style="font-size:.72rem;font-weight:800;color:#475569;text-transform:uppercase;letter-spacing:.8px;margin-bottom:.5rem">
            Reference Excel Template <span style="font-weight:400;color:#94a3b8;text-transform:none;letter-spacing:0;font-size:.71rem">(optional)</span>
          </div>
          <label class="uzone" id="dz5" style="min-height:120px;border-color:#86efac;background:#f0fdf4"
            ondragover="dzDrag(event,this)" ondragleave="dzLeave(this)" ondrop="dzDrop(event,this,'f5','dz5')">
            <input type="file" id="f5" accept=".xlsx,.xls" onchange="fileChosen(this,'dz5','f5',null)">
            <div class="uz-ico" style="color:#86efac;font-size:2.2rem">&#128202;</div>
            <div class="uz-title">Drag &amp; drop your Excel template here</div>
            <div class="uz-hint">Excel only (.xlsx / .xls) &nbsp;&#183;&nbsp; All formatting preserved &nbsp;&#183;&nbsp; Leave empty to auto-generate</div>
          </label>
          <div class="file-row" id="fr-ref"><span class="fr-ico">&#9745;</span><span class="fr-name" id="fn-ref"></span><span class="fr-sz" id="fs-ref"></span></div>
        </div>

        <div class="manual-box" style="border-color:#e9d5ff">
          <button class="mtoggle" onclick="toggleEmail()" style="color:#7c3aed">
            <span>&#9993;&nbsp; Email Configuration &nbsp;<span style="font-weight:400;color:#94a3b8;font-size:.76rem">(pre-fill to send report to Sumit after processing)</span></span>
            <span id="earrow">&#9660;</span>
          </button>
          <div class="mbody" id="ebody">
            <div class="fgrid">
              <div class="fg"><label>Your Gmail</label><input type="email" id="pre-from" placeholder="you@gmail.com" oninput="syncEmail('from',this.value)"></div>
              <div class="fg"><label>Gmail App Password</label><input type="password" id="pre-pass" placeholder="App password" oninput="syncEmail('pass',this.value)"></div>
              <div class="fg"><label>Send To (Sumit)</label><input type="email" id="pre-to" placeholder="sumit@company.com" oninput="syncEmail('to',this.value)"></div>
              <div class="fg"><label>CC (optional)</label><input type="email" id="pre-cc" placeholder="optional" oninput="syncEmail('cc',this.value)"></div>
            </div>
          </div>
        </div>

        <div class="manual-box">
          <button class="mtoggle" onclick="toggleManual()">
            <span>&#9881;&nbsp; Manual / Effort Values &nbsp;<span style="font-weight:400;color:#94a3b8;font-size:.76rem">(yellow cells — Sumit fills these)</span></span>
            <span id="marrow">&#9660;</span>
          </button>
          <div class="mbody" id="mbody">
            <div class="fgrid">
              <div class="fg"><label>Resources onboarded within SLA</label><input type="number" id="mi-resources_onboarded_sla" value="0" min="0"></div>
              <div class="fg"><label>Total resources onboarded</label><input type="number" id="mi-total_resources" value="0" min="0"></div>
              <div class="fg"><label>Actual effort &#8211; Problem tickets closed</label><input type="number" id="mi-effort_prob_closed" value="0" min="0"></div>
              <div class="fg"><label>Total Effort &#8211; Standard Service Requests</label><input type="number" id="mi-effort_ssr" value="0" min="0"></div>
              <div class="fg"><label>Total Effort &#8211; Ad hoc Service Requests</label><input type="number" id="mi-effort_adhoc" value="0" min="0"></div>
              <div class="fg"><label>Actual Effort &#8211; Unknown incidents closed</label><input type="number" id="mi-effort_unknown_inc" value="0" min="0"></div>
              <div class="fg"><label>Actual Effort &#8211; Known incidents closed</label><input type="number" id="mi-effort_known_inc" value="0" min="0"></div>
            </div>
          </div>
        </div>

        <div class="err-box" id="err-box"></div>
        <div class="act-row">
          <button class="btn-back" onclick="goStep(4)">&#8592; Back</button>
          <button class="btn-p" id="finbtn" onclick="processReport()">&#128640;&nbsp; Process &amp; Generate Excel</button>
        </div>
      </div>
    </div>
  </div>

</div><!-- /wizard -->

<!-- ══ ALL FILES PROCESSED card ═════════════════════════════════════════════ -->
<div id="done-card" style="max-width:700px;margin:0 auto 1.2rem">
  <div class="done-circle">&#10003;</div>
  <div class="done-title">All Files Processed!</div>
  <div class="done-sub">Your ticket data is ready to view and export.</div>
  <div class="done-badges">
    <span class="dbadge ritm">&#128203; RITMs &nbsp;<strong id="done-ritm">0</strong></span>
    <span class="dbadge inc">&#127903; Incidents &nbsp;<strong id="done-inc">0</strong></span>
    <span class="dbadge macm">&#128295; MACM &nbsp;<strong id="done-macm">0</strong></span>
    <span class="dbadge prob" id="done-prob-badge">&#128196; Problems &nbsp;<strong id="done-prob">0</strong></span>
  </div>
  <div class="done-actions">
    <button class="btn-dashboard" onclick="showDashboard()">&#128202; View Dashboard</button>
    <button class="btn-dl-done" id="done-dl-btn">&#11123; Download Excel</button>
    <button class="btn-email-done" onclick="openEmailModal()">&#9993; Send to Sumit</button>
    <button class="btn-start-over" onclick="resetAll()">&#8635; Start Over</button>
  </div>
</div>

<!-- ══ DASHBOARD ═════════════════════════════════════════════════════════════-->
<div id="dashboard">

  <!-- toolbar -->
  <div class="dash-toolbar">
    <div class="dt-left">
      <div class="dt-title">&#128202; Mainspring Report Dashboard</div>
      <div class="dt-sub" id="dt-sub">Ready</div>
    </div>
    <div class="dt-right">
      <button class="btn-dl" id="dl-top">&#11123; Excel</button>
      <button class="btn-csv" id="dl-csv">&#11123; CSV</button>
      <button class="btn-email-dash" onclick="openEmailModal()">&#9993; Send Email</button>
      <button class="btn-back" onclick="resetAll()">&#8635; New Report</button>
    </div>
  </div>

  <!-- hero banner -->
  <div class="dash-banner">
    <div>
      <h1>&#128202; Mainspring Monthly Report</h1>
      <p id="dash-sub">All metrics calculated &amp; verified</p>
    </div>
    <div class="dash-month-block">
      <div class="dash-month" id="dash-month"></div>
      <div class="dash-msub">Monthly Metrics Report</div>
    </div>
  </div>

  <!-- Date Range Filter -->
  <div class="df-card" id="df-card">
    <div class="df-hdr">
      <div class="df-title">&#128197; Date Range Filter <span style="font-size:.7rem;font-weight:500;color:#64748b;margin-left:.3rem">&#8212; each sheet filtered independently by its date column</span></div>
      <span class="df-badge" id="df-badge">Full Month</span>
    </div>
    <div class="df-controls">
      <div class="df-grp">
        <div class="df-lbl">Start Date</div>
        <input type="date" id="dr-start" class="df-input" oninput="syncDFbadge()">
      </div>
      <div class="df-grp">
        <div class="df-lbl">End Date</div>
        <input type="date" id="dr-end" class="df-input" oninput="syncDFbadge()">
      </div>
      <button class="df-apply" id="df-apply-btn" onclick="applyDateFilter()">&#128197; Apply Filter</button>
      <button class="df-reset-btn" onclick="resetDateFilter()">&#8635; Reset</button>
    </div>
    <div id="df-breakdown" style="display:none">
      <div class="df-range-lbl" id="df-range-lbl"></div>
      <div class="df-sheets" id="df-sheets"></div>
    </div>
    <div class="df-spin" id="df-spin" style="display:none">Applying filter&#8230;</div>
    <div class="df-ferr" id="df-ferr" style="display:none"></div>
  </div>

  <!-- KPIs -->
  <div class="sec-head">&#127775; Key Performance Indicators</div>
  <div class="kpi-row">
    <div class="kpi" style="border-top-color:#0f3460">
      <div class="kpi-ico">&#128203;</div>
      <div class="kpi-val" id="kpi-total" style="color:#0f3460">0</div>
      <div class="kpi-lbl">Total Resolved</div>
    </div>
    <div class="kpi" style="border-top-color:#1565c0">
      <div class="kpi-ico">&#127903;</div>
      <div class="kpi-val" id="kpi-inc-r" style="color:#1565c0">0</div>
      <div class="kpi-lbl">Inc. Received</div>
    </div>
    <div class="kpi" style="border-top-color:#198754">
      <div class="kpi-ico">&#10003;</div>
      <div class="kpi-val" id="kpi-inc-s" style="color:#198754">0</div>
      <div class="kpi-lbl">Inc. Resolved</div>
    </div>
    <div class="kpi" style="border-top-color:#dc3545">
      <div class="kpi-ico">&#8594;</div>
      <div class="kpi-val" id="kpi-carry" style="color:#dc3545">0</div>
      <div class="kpi-lbl">Carry Forward</div>
    </div>
    <div class="kpi" style="border-top-color:#0077b6">
      <div class="kpi-ico">&#128295;</div>
      <div class="kpi-val" id="kpi-ritm" style="color:#0077b6">0</div>
      <div class="kpi-lbl">RITMs Done</div>
    </div>
    <div class="kpi" style="border-top-color:#6f42c1">
      <div class="kpi-ico">&#128260;</div>
      <div class="kpi-val" id="kpi-macm" style="color:#6f42c1">0</div>
      <div class="kpi-lbl">MACMs Closed</div>
    </div>
    <div class="kpi" style="border-top-color:#fd7e14">
      <div class="kpi-ico">&#128196;</div>
      <div class="kpi-val" id="kpi-prob" style="color:#fd7e14">0</div>
      <div class="kpi-lbl">Problem Records</div>
    </div>
  </div>

  <!-- Pivot Summary Table -->
  <div class="sec-head">&#128202; Monthly Pivot Summary</div>
  <div id="pivot-table-wrap" style="overflow-x:auto;border-radius:12px;border:1px solid #e2e8f0;margin-bottom:1rem;box-shadow:0 2px 8px rgba(0,0,0,.07)">
    <table style="width:100%;border-collapse:collapse;font-size:.82rem">
      <thead>
        <tr style="background:#1e3a5f">
          <th style="color:#fff;padding:10px 14px;text-align:left;font-weight:700;width:28%">Category</th>
          <th style="color:#fff;padding:10px 10px;text-align:center;font-weight:700">Total</th>
          <th style="color:#fff;padding:10px 10px;text-align:center;font-weight:700">Resolved / Closed</th>
          <th style="color:#fff;padding:10px 10px;text-align:center;font-weight:700">Carry Fwd / Open</th>
          <th style="color:#fff;padding:10px 10px;text-align:center;font-weight:700">Success Rate</th>
        </tr>
      </thead>
      <tbody id="pivot-body"></tbody>
    </table>
  </div>

  <!-- Incident charts -->
  <div class="sec-head" id="chart-section">&#128200; Incident Analysis</div>
  <div class="chart-row c3">
    <div class="chart-card"><canvas id="chart-res"></canvas></div>
    <div class="chart-card"><canvas id="chart-knu"></canvas></div>
    <div class="chart-card"><canvas id="chart-kedb"></canvas></div>
  </div>

  <!-- RITM / Change / Mix charts -->
  <div class="sec-head">&#128218; RITM &nbsp;&#183;&nbsp; Change &nbsp;&#183;&nbsp; Ticket Mix</div>
  <div class="chart-row c3b">
    <div class="chart-card"><canvas id="chart-ritm"></canvas></div>
    <div class="chart-card"><canvas id="chart-chg"></canvas></div>
    <div class="chart-card"><canvas id="chart-mix"></canvas></div>
  </div>

  <!-- Metrics table -->
  <div class="sec-head">&#128203; Complete Metrics Table</div>
  <div class="tbl-wrap">
    <table class="mtbl">
      <thead><tr>
        <th style="width:44%">Measure Name</th>
        <th class="thv">Value</th>
        <th>Measurement Criteria</th>
      </tr></thead>
      <tbody id="tbl-body"></tbody>
    </table>
  </div>
  <div style="margin:1.4rem 0 .5rem;display:flex;gap:.65rem;flex-wrap:wrap;align-items:center">
    <button class="btn-dl" id="dl-bot">&#11123; Download Excel</button>
    <button class="btn-email-dash" onclick="openEmailModal()">&#9993; Send to Sumit</button>
    <button class="btn-back" onclick="resetAll()">&#8635; New Report</button>
  </div>
</div>

</div><!-- /wrap -->

<!-- ══ EMAIL MODAL ═══════════════════════════════════════════════════════════ -->
<div class="modal-ov" id="email-modal">
  <div class="email-modal">
    <div class="em-hdr">
      <div class="em-title">&#9993; Send Report by Email</div>
      <button class="em-close" onclick="closeEmailModal()">&#10005;</button>
    </div>
    <div class="em-body">
      <div class="fg"><label>Your Gmail Address</label>
        <input type="email" id="em-from" placeholder="you@gmail.com" autocomplete="email"></div>
      <div class="fg"><label>Gmail App Password &nbsp;<span style="font-weight:400;color:#94a3b8;font-size:.7rem">(not your normal password)</span></label>
        <input type="password" id="em-pass" placeholder="xxxx xxxx xxxx xxxx" autocomplete="new-password"></div>
      <div class="fg"><label>Send To &nbsp;<span style="font-weight:400;color:#94a3b8;font-size:.7rem">(Sumit&#39;s email)</span></label>
        <input type="email" id="em-to" placeholder="sumit@company.com"></div>
      <div class="fg"><label>CC &nbsp;<span style="font-weight:400;color:#94a3b8;font-size:.7rem">(optional)</span></label>
        <input type="email" id="em-cc" placeholder="manager@company.com (optional)"></div>
    </div>
    <div class="em-footer">
      <div class="em-note">
        &#128274; Enable Gmail 2-Step Verification &#8594; Google Account &#8594; Security &#8594; App Passwords &#8594; Generate a password for &ldquo;Mail&rdquo;. Paste that 16-char password above.
      </div>
      <div class="em-btns">
        <button class="btn-back" onclick="closeEmailModal()">Cancel</button>
        <button class="btn-email-send" id="em-send-btn" onclick="sendEmail()">&#9993; Send Now</button>
      </div>
    </div>
    <div class="em-status" id="em-status"></div>
  </div>
</div>

<div class="spin-ov" id="spinner">
  <div class="spinner"></div>
  <div class="spin-txt" id="spin-txt">Processing&#8230;</div>
</div>

<script>
const FILES = {ritm:null,inc:null,macm:null,prob:null};
let PROB_SKIPPED = false;
const CHARTS = {};
let STORED_DATA = null;  // holds last /process response
// date filter session state
let G_savedFolder  = '';
let G_probSkipped  = false;
let G_monthLabel   = '';
let G_origMetrics  = null;
let G_origManual   = {};
let G_origTableRows= '';

// ── step nav ──────────────────────────────────────────────────────────────────
function goStep(n) {
  for (let i=1;i<=5;i++) document.getElementById('panel-'+i).classList.toggle('active',i===n);
  for (let i=1;i<=6;i++) {
    const ps=document.getElementById('ps'+i);
    ps.classList.remove('active','done');
    if(i<n) ps.classList.add('done');
    else if(i===n) ps.classList.add('active');
  }
  for (let i=1;i<=5;i++) document.getElementById('pl'+i).classList.toggle('done',i<n);
  window.scrollTo({top:0,behavior:'smooth'});
}

// ── drag / drop ───────────────────────────────────────────────────────────────
function dzDrag(e,el){e.preventDefault();el.classList.add('over');}
function dzLeave(el){el.classList.remove('over');}
function dzDrop(e,el,fid,dzid){
  e.preventDefault();dzLeave(el);
  const dt=e.dataTransfer;
  if(dt&&dt.files&&dt.files.length){
    const inp=document.getElementById(fid);
    try{const t=new DataTransfer();t.items.add(dt.files[0]);inp.files=t.files;inp.dispatchEvent(new Event('change'));}catch(_){}
  }
}

function fmt(b){if(b<1024)return b+' B';if(b<1048576)return(b/1024).toFixed(1)+' KB';return(b/1048576).toFixed(1)+' MB';}

function fileChosen(inp,dzid,fid,type){
  const dz=document.getElementById(dzid);
  if(inp.files&&inp.files.length){
    const f=inp.files[0]; dz.classList.add('done');
    const key=type||'ref';
    const frId='fr-'+(type||'ref'), fnId='fn-'+(type||'ref'), fsId='fs-'+(type||'ref');
    document.getElementById(fnId).textContent=f.name;
    document.getElementById(fsId).textContent=fmt(f.size);
    document.getElementById(frId).classList.add('show');
    if(type){const pb=document.getElementById('pb-'+type);if(pb)pb.disabled=false;}
  } else {
    dz.classList.remove('done');
  }
}

// ── per-step preview ──────────────────────────────────────────────────────────
const EP={ritm:'/preview/ritm',inc:'/preview/incidents',macm:'/preview/macm',prob:'/preview/problems'};
const STEPOF={ritm:1,inc:2,macm:3,prob:4};
const TNAME={ritm:'RITM',inc:'Incidents',macm:'MACM',prob:'Problems'};

async function processStep(type){
  const f=document.getElementById('f'+(STEPOF[type])).files[0]; if(!f) return;
  const pb=document.getElementById('pb-'+type);
  pb.disabled=true; pb.innerHTML='<span style="display:inline-block;animation:spin .6s linear infinite">&#9696;</span>&nbsp;Processing&#8230;';
  setSpinner(true,'Reading '+f.name+'&#8230;');
  const fd=new FormData(); fd.append('file',f);
  try{
    const r=await fetch(EP[type],{method:'POST',body:fd});
    const d=await r.json();
    setSpinner(false);
    const sb=document.getElementById('sb-'+type); sb.classList.remove('err');
    if(!d.ok){
      sb.innerHTML='&#9888; '+d.error; sb.classList.add('show','err');
      pb.disabled=false; pb.innerHTML='&#9654; Process '+TNAME[type]; return;
    }
    FILES[type]=f;
    pb.innerHTML='&#10003; Processed'; pb.style.background='#16a34a';
    sb.innerHTML='&#10003; '+buildMsg(type,d); sb.classList.add('show');
    document.getElementById('psub'+STEPOF[type]).textContent=ticketLbl(type,d);
    if(d.cols&&d.cols.length){
      const cn=document.getElementById('cn-'+type);
      if(cn) cn.innerHTML='<b>Columns detected:</b> '+d.cols.join(', ')+(d.cols.length===10?'&#8230;':'');
    }
    document.getElementById('nx'+STEPOF[type]).classList.add('show');
  }catch(e){
    setSpinner(false); const sb=document.getElementById('sb-'+type);
    sb.innerHTML='&#9888; '+e.message; sb.classList.add('show','err');
    pb.disabled=false; pb.innerHTML='&#9654; Process '+TNAME[type];
  }
}

function buildMsg(t,d){
  if(t==='ritm') return 'RITM &#8212; '+d.total+' tickets &nbsp;&#183;&nbsp; Completed: '+d.completed+' &nbsp;&#183;&nbsp; Open/SLA: '+d.open+' &nbsp;&#183;&nbsp; Ad Hoc: '+d.adhoc;
  if(t==='inc')  return 'Incidents &#8212; '+d.total+' total &nbsp;&#183;&nbsp; Resolved: '+d.resolved+' &nbsp;&#183;&nbsp; Carry Fwd: '+d.carry_fwd+' &nbsp;&#183;&nbsp; Known Err: '+d.known_err_closed;
  if(t==='macm') return 'MACM &#8212; '+d.total+' change requests found';
  if(t==='prob') return 'Problems &#8212; '+d.total+' rows &nbsp;&#183;&nbsp; Resolved: '+d.resolved;
  return '';
}
function ticketLbl(t,d){
  if(t==='ritm') return d.total+' tickets';
  if(t==='inc')  return d.total+' tickets';
  if(t==='macm') return d.total+' changes';
  if(t==='prob') return d.resolved+' resolved';
  return '';
}

// ── skip problems ─────────────────────────────────────────────────────────────
function skipProblems(){
  PROB_SKIPPED=true; FILES.prob=null;
  document.getElementById('skip-bar').classList.add('show');
  document.getElementById('pb-prob').style.display='none';
  document.getElementById('psub4').textContent='Skipped';
  document.getElementById('nx4').classList.add('show');
}

// ── email config sync (pre-fill → modal) ─────────────────────────────────────
function syncEmail(field,val){
  const map={from:'em-from',pass:'em-pass',to:'em-to',cc:'em-cc'};
  const el=document.getElementById(map[field]); if(el) el.value=val;
}
function toggleEmail(){
  const b=document.getElementById('ebody'),a=document.getElementById('earrow');
  b.classList.toggle('open'); a.textContent=b.classList.contains('open')?'&#9650;':'&#9660;';
}

// ── manual toggle ─────────────────────────────────────────────────────────────
function toggleManual(){
  const b=document.getElementById('mbody'),a=document.getElementById('marrow');
  b.classList.toggle('open'); a.textContent=b.classList.contains('open')?'&#9650;':'&#9660;';
}

// ── email modal ───────────────────────────────────────────────────────────────
function openEmailModal(){
  document.getElementById('em-status').textContent='';
  document.getElementById('em-status').className='em-status';
  document.getElementById('em-send-btn').disabled=false;
  document.getElementById('em-send-btn').innerHTML='&#9993; Send Now';
  const sf=localStorage.getItem('ms_from');
  const st=localStorage.getItem('ms_to');
  const sc=localStorage.getItem('ms_cc');
  const ef=document.getElementById('em-from');
  const et=document.getElementById('em-to');
  const ec=document.getElementById('em-cc');
  if(sf&&!ef.value) ef.value=sf;
  if(st&&!et.value) et.value=st;
  if(sc&&!ec.value) ec.value=sc;
  document.getElementById('email-modal').classList.add('show');
}
function closeEmailModal(){document.getElementById('email-modal').classList.remove('show');}

async function sendEmail(){
  if(!STORED_DATA){alert('No report data. Please process files first.');return;}
  const from=document.getElementById('em-from').value.trim();
  const pass=document.getElementById('em-pass').value.trim();
  const to=document.getElementById('em-to').value.trim();
  const cc=document.getElementById('em-cc').value.trim();
  if(!from||!pass||!to){
    setEmStatus('err','Please fill in Gmail address, App Password, and recipient email.');return;
  }
  const btn=document.getElementById('em-send-btn');
  btn.disabled=true; btn.innerHTML='<span style="display:inline-block;animation:spin .6s linear infinite">&#9696;</span>&nbsp;Sending&hellip;';
  setEmStatus('','');
  const fd=new FormData();
  fd.append('from_email',from); fd.append('password',pass);
  fd.append('to_email',to);     fd.append('cc_email',cc);
  fd.append('month',STORED_DATA.month||'Monthly Report');
  fd.append('excel_b64',STORED_DATA.excel_b64);
  fd.append('filename',STORED_DATA.filename);
  try{
    const r=await fetch('/send_email',{method:'POST',body:fd});
    const d=await r.json();
    if(d.ok){
      setEmStatus('ok','&#10003; '+d.message);
      btn.innerHTML='&#10003; Sent!'; btn.style.background='#16a34a';
      localStorage.setItem('ms_from',from);
      localStorage.setItem('ms_to',to);
      if(cc) localStorage.setItem('ms_cc',cc);
    } else {
      setEmStatus('err','&#9888; '+d.error);
      btn.disabled=false; btn.innerHTML='&#9993; Send Now';
    }
  }catch(e){
    setEmStatus('err','&#9888; Network error: '+e.message);
    btn.disabled=false; btn.innerHTML='&#9993; Send Now';
  }
}
function setEmStatus(type,msg){
  const el=document.getElementById('em-status');
  el.innerHTML=msg; el.className='em-status'+(type?' '+type:'');
}

// ── final process ─────────────────────────────────────────────────────────────
async function processReport(){
  hideErr();
  if(!FILES.ritm){showErr('RITM file missing &#8212; go back to Step 1.');return;}
  if(!FILES.inc) {showErr('Incidents file missing &#8212; go back to Step 2.');return;}
  if(!FILES.macm){showErr('MACM file missing &#8212; go back to Step 3.');return;}
  const fd=new FormData();
  fd.append('ritm',FILES.ritm); fd.append('incidents',FILES.inc); fd.append('macm',FILES.macm);
  if(FILES.prob&&!PROB_SKIPPED) fd.append('problems',FILES.prob);
  fd.append('prob_skipped',PROB_SKIPPED?'true':'false');
  const refF=document.getElementById('f5').files[0]; if(refF) fd.append('reference',refF);
  ['resources_onboarded_sla','total_resources','effort_prob_closed',
   'effort_ssr','effort_adhoc','effort_unknown_inc','effort_known_inc'].forEach(k=>{
    fd.append(k,document.getElementById('mi-'+k).value||'0');
  });
  fd.append('month_label',document.getElementById('month-select').value+' '+document.getElementById('year-input').value);
  setSpinner(true,'Calculating metrics and building Excel&#8230;');
  document.getElementById('finbtn').disabled=true;
  try{
    const resp=await fetch('/process',{method:'POST',body:fd});
    const data=await resp.json();
    setSpinner(false);
    if(!data.ok){showErr(data.error||'Processing failed');document.getElementById('finbtn').disabled=false;return;}
    STORED_DATA=data;
    showDashboard();
  }catch(e){
    setSpinner(false); showErr('Network error: '+e.message);
    document.getElementById('finbtn').disabled=false;
  }
}

// ── "All Files Processed!" card ───────────────────────────────────────────────
function showDoneCard(data){
  const c=data.counts;
  document.getElementById('done-ritm').textContent=c.ritm||0;
  document.getElementById('done-inc').textContent=c.inc||0;
  document.getElementById('done-macm').textContent=c.macm||0;
  document.getElementById('done-prob').textContent=c.prob||0;
  if(!c.prob) document.getElementById('done-prob-badge').style.display='none';

  const xlH=()=>triggerDL(data.excel_b64,data.filename,'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
  document.getElementById('done-dl-btn').onclick=xlH;

  // mark all progress done
  for(let i=1;i<=5;i++){document.getElementById('ps'+i).classList.remove('active');document.getElementById('ps'+i).classList.add('done');}
  for(let i=1;i<=5;i++) document.getElementById('pl'+i).classList.add('done');
  document.getElementById('ps6').classList.add('active');
  document.getElementById('psub5').textContent='Done';
  document.getElementById('psub6').textContent='Ready';

  document.getElementById('wizard').style.display='none';
  document.getElementById('done-card').classList.add('show');
  document.getElementById('btn-reset').classList.add('show');
  window.scrollTo({top:0,behavior:'smooth'});
}

// ── animated KPI counter ──────────────────────────────────────────────────────
function animCount(id,target,dur=900){
  const el=document.getElementById(id); if(!el) return;
  const start=Date.now(),from=0;
  (function tick(){
    const pct=Math.min(1,(Date.now()-start)/dur);
    const ease=1-Math.pow(1-pct,3);
    el.textContent=Math.round(from+(target-from)*ease);
    if(pct<1) requestAnimationFrame(tick);
  })();
}

// ── applyMetrics — update KPIs, pivot, charts, metrics table ─────────────────
function applyMetrics(M, mi, tableRows){
  if(tableRows) document.getElementById('tbl-body').innerHTML=tableRows;
  const pv=document.getElementById('pivot-body');
  if(pv&&M){
    const pct=(n,d)=>d>0?Math.round(n/d*100)+'%':'—';
    const rows=[
      {cat:'Incidents (INC)',       bg:'#e8f5e9',total:M.new_inc_received||0,   res:M.inc_resolved||0,        cf:M.inc_carry_fwd||0,       rate:pct(M.inc_resolved,M.new_inc_received)},
      {cat:'RITM (Svc Requests)',   bg:'#e3f2fd',total:M.total_ssr||0,          res:M.ssr_completed_on_time||0,cf:M.ssr_open_beyond_sla||0, rate:pct(M.ssr_completed_on_time,M.total_ssr)},
      {cat:'Changes (MACM)',        bg:'#fff9c4',total:M.change_deliver||0,     res:M.changes_success||0,     cf:M.changes_rolled_back||0,  rate:pct(M.changes_success,M.change_deliver)},
      {cat:'Problem Tickets (PRB)', bg:'#fff3e0',total:M.prob_records_resolved||0,res:M.prob_records_resolved||0,cf:0,                     rate:'—'},
    ];
    const tot=M.total_tickets_resolved||0;
    pv.innerHTML=rows.map(r=>`<tr style="background:${r.bg}">
      <td style="padding:8px 14px;font-weight:700;border:1px solid #e2e8f0">${r.cat}</td>
      <td style="padding:8px 10px;text-align:center;font-weight:800;border:1px solid #e2e8f0">${r.total}</td>
      <td style="padding:8px 10px;text-align:center;font-weight:800;color:#16a34a;border:1px solid #e2e8f0">${r.res}</td>
      <td style="padding:8px 10px;text-align:center;color:#dc2626;border:1px solid #e2e8f0">${r.cf}</td>
      <td style="padding:8px 10px;text-align:center;font-weight:700;border:1px solid #e2e8f0">${r.rate}</td>
    </tr>`).join('')+
    `<tr style="background:#1e3a5f">
      <td colspan="2" style="padding:9px 14px;color:#fff;font-weight:800;font-size:.9rem;border:1px solid #334">TOTAL TICKETS RESOLVED</td>
      <td style="padding:9px 10px;text-align:center;color:#6ee7b7;font-weight:900;font-size:1.1rem;border:1px solid #334">${tot}</td>
      <td colspan="2" style="padding:9px 10px;color:#93c5fd;font-size:.78rem;border:1px solid #334">INC + RITM + PROB + MACM</td>
    </tr>`;
  }
  animCount('kpi-total', M.total_tickets_resolved||0);
  animCount('kpi-inc-r', M.new_inc_received||0);
  animCount('kpi-inc-s', M.inc_resolved||0);
  animCount('kpi-carry', M.inc_carry_fwd||0,700);
  animCount('kpi-ritm',  M.ssr_completed_on_time||0);
  animCount('kpi-macm',  M.changes_closed||0);
  animCount('kpi-prob',  M.prob_records_resolved||0);
  requestAnimationFrame(()=>{
    renderCharts(M);
    setTimeout(()=>{Object.values(CHARTS).forEach(c=>{try{c.resize();}catch(_){}});},400);
  });
}

// ── View Dashboard ────────────────────────────────────────────────────────────
function showDashboard(){
  if(!STORED_DATA) return;
  const data=STORED_DATA, M=data.metrics;
  const month=data.month||'';

  G_savedFolder   = data.saved_folder || '';
  G_probSkipped   = PROB_SKIPPED;
  G_monthLabel    = month;
  G_origMetrics   = M;
  G_origManual    = data.manual || {};
  G_origTableRows = data.table_rows || '';

  document.getElementById('dt-sub').textContent=data.source;
  document.getElementById('dash-month').textContent=month;
  document.getElementById('dash-sub').textContent='All metrics calculated \u00b7 '+month;

  const xlH=()=>triggerDL(data.excel_b64,data.filename,'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
  document.getElementById('dl-top').onclick=xlH;
  document.getElementById('dl-bot').onclick=xlH;
  document.getElementById('dl-csv').onclick=()=>triggerDLText(buildCsv(),'Mainspring_'+month.replace(/ /g,'_')+'.csv','text/csv');

  const dr=monthLabelToDates(month);
  if(dr){
    document.getElementById('dr-start').value=dr.start;
    document.getElementById('dr-end').value=dr.end;
  }
  syncDFbadge();

  document.getElementById('done-card').classList.remove('show');
  document.getElementById('wizard').style.display='none';
  document.getElementById('btn-reset').classList.add('show');
  for(let i=1;i<=5;i++){
    document.getElementById('ps'+i).classList.remove('active');
    document.getElementById('ps'+i).classList.add('done');
    document.getElementById('pl'+i).classList.add('done');
  }
  document.getElementById('dashboard').classList.add('show');
  document.querySelector('.wrap').classList.add('dash-mode');
  document.getElementById('ps6').classList.remove('active');
  document.getElementById('ps6').classList.add('done');
  document.getElementById('psub6').textContent='Done';

  applyMetrics(M, data.manual||{}, data.table_rows||'');

  window.scrollTo({top:0,behavior:'smooth'});
  setTimeout(()=>{
    const cs=document.getElementById('chart-section');
    if(cs) cs.scrollIntoView({behavior:'smooth',block:'start'});
  },1100);
}

// ── Date Range Filter helpers ─────────────────────────────────────────────────
function monthLabelToDates(label){
  const months=['January','February','March','April','May','June','July','August','September','October','November','December'];
  const parts=label.split(' ');
  let m=-1,y=0;
  for(const p of parts){
    const mi=months.findIndex(x=>x.toLowerCase()===p.toLowerCase());
    if(mi>=0) m=mi;
    if(/^\d{4}$/.test(p)) y=parseInt(p);
  }
  if(m<0||!y) return null;
  const start=`${y}-${String(m+1).padStart(2,'0')}-01`;
  const lastDay=new Date(y,m+1,0).getDate();
  const end=`${y}-${String(m+1).padStart(2,'0')}-${String(lastDay).padStart(2,'0')}`;
  return {start,end};
}

function syncDFbadge(){
  const s=document.getElementById('dr-start').value;
  const e=document.getElementById('dr-end').value;
  const b=document.getElementById('df-badge');
  if(!b) return;
  if(!s||!e){b.textContent='Full Month';b.classList.remove('active');return;}
  const dr=monthLabelToDates(G_monthLabel||'');
  if(dr&&s===dr.start&&e===dr.end){b.textContent='Full Month';b.classList.remove('active');}
  else{b.textContent='Custom Range';b.classList.add('active');}
}

async function applyDateFilter(){
  const s=document.getElementById('dr-start').value;
  const e=document.getElementById('dr-end').value;
  if(!s||!e){alert('Please select both start and end dates.');return;}
  if(s>e){alert('Start date must be before end date.');return;}
  if(!G_savedFolder){alert('Session expired — please re-upload files.');return;}
  const btn=document.getElementById('df-apply-btn');
  const spin=document.getElementById('df-spin');
  const ferr=document.getElementById('df-ferr');
  btn.disabled=true; spin.style.display=''; ferr.style.display='none';
  document.getElementById('df-breakdown').style.display='none';
  try{
    const fd=new FormData();
    fd.append('folder',G_savedFolder);
    fd.append('start_date',s);
    fd.append('end_date',e);
    fd.append('prob_skipped',G_probSkipped?'true':'false');
    fd.append('month_label',G_monthLabel);
    const resp=await fetch('/refilter',{method:'POST',body:fd});
    const data=await resp.json();
    if(!data.ok){ferr.textContent='Error: '+(data.error||'Unknown error');ferr.style.display='';spin.style.display='none';btn.disabled=false;return;}
    applyMetrics(data.metrics,data.manual||{},data.table_rows||'');
    showSheetBreakdown(data.sheet_counts,s,e);
    syncDFbadge();
  }catch(err){ferr.textContent='Network error: '+err.message;ferr.style.display='';}
  spin.style.display='none'; btn.disabled=false;
}

function resetDateFilter(){
  const dr=monthLabelToDates(G_monthLabel||'');
  if(dr){document.getElementById('dr-start').value=dr.start;document.getElementById('dr-end').value=dr.end;}
  syncDFbadge();
  document.getElementById('df-breakdown').style.display='none';
  document.getElementById('df-ferr').style.display='none';
  if(G_origMetrics) applyMetrics(G_origMetrics,G_origManual,G_origTableRows);
}

function showSheetBreakdown(counts,start,end){
  const wrap=document.getElementById('df-breakdown');
  const lbl=document.getElementById('df-range-lbl');
  const sheets=document.getElementById('df-sheets');
  if(!wrap||!lbl||!sheets) return;
  const fmt=d=>d.split('-').reverse().join('/');
  lbl.textContent='Filtered: '+fmt(start)+' to '+fmt(end);
  const ORDER=['INC','RITM','MACM','PROB'];
  const COLORS={'INC':'#16a34a','RITM':'#0077b6','MACM':'#7c3aed','PROB':'#ea580c'};
  sheets.innerHTML=ORDER.filter(k=>counts[k]).map(k=>{
    const c=counts[k];
    const pct=c.before>0?Math.round(c.after/c.before*100):0;
    const warn=c.after===0&&c.before>0;
    return '<div class="df-sheet'+(warn?' warn':'')+'">'+
      '<div class="df-sheet-name" style="color:'+(COLORS[k]||'#0369a1')+'">'+k+'</div>'+
      '<div class="df-sheet-rows">'+c.after+' <span style="color:#94a3b8">/ '+c.before+'</span> rows <span style="font-weight:700">('+pct+'%)</span></div>'+
      '<div class="df-sheet-col">'+(c.applied?'col: '+c.col:'no date col')+'</div>'+
      '</div>';
  }).join('');
  wrap.style.display='';
}

// ── charts ────────────────────────────────────────────────────────────────────
function renderCharts(M){
  destroyCharts();
  CHARTS.res=new Chart(document.getElementById('chart-res'),{type:'doughnut',
    data:{labels:['Resolved','Carry Forward'],datasets:[{data:[M.inc_resolved||0,M.inc_carry_fwd||0],backgroundColor:['#198754','#dc3545'],borderWidth:2}]},
    options:cOpts('Resolution Status',{cutout:'55%'})});
  const yInt={y:{beginAtZero:true,ticks:{precision:0,font:{size:11}},grid:{color:'rgba(0,0,0,.06)'}}};
  CHARTS.knu=new Chart(document.getElementById('chart-knu'),{type:'bar',
    data:{labels:['Known Error','Unknown'],datasets:[
      {label:'Closed',data:[M.known_err_inc_closed||0,M.unknown_inc_closed||0],backgroundColor:['#fd7e14','#0f3460'],borderRadius:4},
      {label:'KEDB',data:[(M.ku_inc_kedb||0)-(M.unknown_inc_kedb||0),M.unknown_inc_kedb||0],backgroundColor:['#ffc107','#198754'],borderRadius:4}]},
    options:cOpts('Known vs Unknown',{scales:yInt})});
  const uk=M.unknown_inc_kedb||0,un=Math.max(0,(M.unknown_inc_closed||0)-uk);
  const kk=Math.max(0,(M.ku_inc_kedb||0)-uk),kn=Math.max(0,(M.known_err_inc_closed||0)-kk);
  CHARTS.kedb=new Chart(document.getElementById('chart-kedb'),{type:'bar',
    data:{labels:['Unknown','Known Error'],datasets:[{label:'With KEDB',data:[uk,kk],backgroundColor:'#198754',borderRadius:4},{label:'No KEDB',data:[un,kn],backgroundColor:'#dc3545',borderRadius:4}]},
    options:cOpts('KEDB Coverage',{scales:yInt})});
  CHARTS.ritm=new Chart(document.getElementById('chart-ritm'),{type:'bar',
    data:{labels:['Total','Completed','Open/SLA','Ad Hoc'],datasets:[{data:[M.total_ssr||0,M.ssr_completed_on_time||0,M.ssr_open_beyond_sla||0,M.total_adhoc||0],backgroundColor:['#0077b6','#198754','#dc3545','#6f42c1'],borderRadius:4}]},
    options:cOpts('RITM Breakdown',{plugins:{legend:{display:false}},scales:yInt})});
  CHARTS.chg=new Chart(document.getElementById('chart-chg'),{type:'doughnut',
    data:{labels:['Implemented','Rolled Back'],datasets:[{data:[M.changes_success||0,Math.max(0,(M.change_deliver||0)-(M.changes_success||0))],backgroundColor:['#375623','#dc3545'],borderWidth:2}]},
    options:cOpts('Change Requests',{cutout:'55%'})});
  CHARTS.mix=new Chart(document.getElementById('chart-mix'),{type:'doughnut',
    data:{labels:['Incidents','RITMs','MACMs','Problems'],datasets:[{data:[M.new_inc_received||0,M.total_ssr||0,M.changes_closed||0,M.prob_records_resolved||0],backgroundColor:['#0f3460','#0077b6','#375623','#fd7e14'],borderWidth:2}]},
    options:cOpts('Ticket Mix',{})});
}
function cOpts(t,x){
  const base={responsive:true,maintainAspectRatio:false,
    plugins:{
      title:{display:true,text:t,font:{size:13,weight:'bold'},padding:{bottom:6}},
      legend:{position:'bottom',labels:{font:{size:11},padding:10,boxWidth:13}}
    },
    layout:{padding:{top:4,bottom:2}}
  };
  if(x.cutout!==undefined) base.cutout=x.cutout;
  if(x.scales) base.scales=x.scales;
  if(x.plugins){
    if(x.plugins.legend) Object.assign(base.plugins.legend,x.plugins.legend);
    if(x.plugins.title) Object.assign(base.plugins.title,x.plugins.title);
  }
  return base;
}
function destroyCharts(){Object.values(CHARTS).forEach(c=>{try{c.destroy();}catch(_){}});Object.keys(CHARTS).forEach(k=>delete CHARTS[k]);}

// ── downloads ─────────────────────────────────────────────────────────────────
function triggerDL(b64,fn,mime){const a=document.createElement('a');a.href='data:'+mime+';base64,'+b64;a.download=fn;document.body.appendChild(a);a.click();document.body.removeChild(a);}
function triggerDLText(txt,fn,mime){triggerDL(btoa(unescape(encodeURIComponent(txt))),fn,mime);}
function buildCsv(){const rows=[['Measure Name','Value']];document.querySelectorAll('#tbl-body tr').forEach(tr=>{const c=tr.querySelectorAll('td');if(c.length>=2){const n=c[0].textContent.trim(),v=c[1].textContent.trim();if(n)rows.push(['"'+n.replace(/"/g,'\\"')+'\"',v]);}});return rows.map(r=>r.join(',')).join('\n');}

// ── spinner ───────────────────────────────────────────────────────────────────
function setSpinner(on,msg){document.getElementById('spinner').classList.toggle('show',on);if(msg)document.getElementById('spin-txt').innerHTML=msg;}

// ── error ─────────────────────────────────────────────────────────────────────
function showErr(msg){const e=document.getElementById('err-box');e.innerHTML='&#9888; '+msg;e.classList.add('show');}
function hideErr(){document.getElementById('err-box').classList.remove('show');}

// ── init: auto-set PREVIOUS month (report is for month-1) ─────────────────────
(function(){
  const now=new Date();
  const months=['January','February','March','April','May','June','July','August','September','October','November','December'];
  let rm=now.getMonth()-1, ry=now.getFullYear();
  if(rm<0){rm=11;ry--;}
  const ms=document.getElementById('month-select');
  const yi=document.getElementById('year-input');
  if(ms) ms.value=months[rm];
  if(yi) yi.value=ry;
})();

// ── reset ─────────────────────────────────────────────────────────────────────
function resetAll(){
  destroyCharts(); STORED_DATA=null;
  FILES.ritm=FILES.inc=FILES.macm=FILES.prob=null; PROB_SKIPPED=false;
  ['f1','f2','f3','f4','f5'].forEach(id=>{const e=document.getElementById(id);if(e)e.value='';});
  ['dz1','dz2','dz3','dz4','dz5'].forEach(id=>{const e=document.getElementById(id);if(e)e.classList.remove('done','over');});
  ['fr-ritm','fr-inc','fr-macm','fr-prob','fr-ref'].forEach(id=>{const e=document.getElementById(id);if(e)e.classList.remove('show');});
  ['sb-ritm','sb-inc','sb-macm','sb-prob'].forEach(id=>{const e=document.getElementById(id);if(e){e.innerHTML='';e.classList.remove('show','err');}});
  ['cn-ritm','cn-inc','cn-macm','cn-prob'].forEach(id=>{const e=document.getElementById(id);if(e)e.innerHTML='';});
  ['nx1','nx2','nx3','nx4'].forEach(id=>{const e=document.getElementById(id);if(e)e.classList.remove('show');});
  document.getElementById('skip-bar').classList.remove('show');
  [['pb-ritm','RITM'],['pb-inc','Incidents'],['pb-macm','MACM'],['pb-prob','Problems']].forEach(([id,lbl])=>{
    const e=document.getElementById(id);if(e){e.disabled=true;e.style.background='';e.style.display='';e.innerHTML='&#9654; Process '+lbl;}
  });
  document.getElementById('finbtn').disabled=false;
  document.getElementById('done-card').classList.remove('show');
  document.getElementById('done-prob-badge').style.display='';
  document.getElementById('dashboard').classList.remove('show');
  document.querySelector('.wrap').classList.remove('dash-mode');
  document.getElementById('wizard').style.display='';
  document.getElementById('btn-reset').classList.remove('show');
  document.getElementById('mbody').classList.remove('open');
  document.getElementById('marrow').textContent='&#9660;';
  document.getElementById('ebody').classList.remove('open');
  document.getElementById('earrow').textContent='&#9660;';
  closeEmailModal();
  for(let i=1;i<=6;i++) document.getElementById('psub'+i).textContent='Waiting';
  hideErr(); goStep(1);
}
</script>
</body>
</html>"""

# ══════════════════════════════════════════════════════════════════════════════
if __name__ == "__main__":
    import threading, webbrowser
    print("\n  Mainspring Report Automation")
    print("  -----------------------------")
    print("  Open: http://localhost:5001\n")
    threading.Timer(1.2, lambda: webbrowser.open("http://localhost:5001")).start()
    app.run(debug=False, port=5001, host="0.0.0.0")