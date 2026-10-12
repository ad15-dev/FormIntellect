import io
import json
import os
import random
import re
import sys
import time
import urllib.parse
from collections import Counter
from datetime import datetime, timedelta

import pandas as pd
import pytz
import requests
import streamlit as st

try:
    from openpyxl import load_workbook
    from openpyxl.chart import BarChart, Reference
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    from openpyxl.utils import get_column_letter
    OPENPYXL_AVAILABLE = True
except ImportError:
    OPENPYXL_AVAILABLE = False

# ============================================================
# CONFIG
# ============================================================
IST = pytz.timezone("Asia/Kolkata")
MIN_SECONDS_PER_ROW = 2

st.set_page_config(
    page_title="formIntellect",
    page_icon="⚡",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# ============================================================
# MINIMAL MODERN UI
# ============================================================
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&display=swap');

    /* Base */
    .stApp {
        background: #0a0a0b;
        color: #e4e4e7;
        font-family: 'Inter', -apple-system, sans-serif;
    }
    .main .block-container {
        padding: 2rem 1.5rem 4rem;
        max-width: 1400px;
        margin: 0 auto;
    }
    @media (min-width: 768px) {
        .main .block-container { padding: 3rem 2.5rem 5rem; }
    }

    /* Typography */
    h1 {
        font-size: clamp(1.8rem, 4vw, 2.5rem);
        font-weight: 600;
        letter-spacing: -0.03em;
        color: #fafafa !important;
        margin-bottom: 0.3rem;
    }
    h2 {
        font-size: clamp(1.2rem, 2.5vw, 1.4rem);
        font-weight: 600;
        letter-spacing: -0.02em;
        color: #fafafa !important;
        margin: 2rem 0 1rem;
    }
    h3 {
        font-size: 1rem;
        font-weight: 500;
        color: #a1a1aa !important;
        margin-bottom: 0.75rem;
    }
    p, label, .stMarkdown { color: #a1a1aa !important; font-size: 0.9rem; }

    /* Hero */
    .hero {
        padding: 2rem 0 1.5rem;
        border-bottom: 1px solid #1f1f23;
        margin-bottom: 2rem;
    }
    .hero-sub {
        color: #71717a;
        font-size: 0.95rem;
        font-weight: 400;
        margin-top: 0.5rem;
    }
    .hero-badge {
        display: inline-block;
        padding: 0.25rem 0.75rem;
        background: rgba(99, 102, 241, 0.1);
        color: #818cf8;
        border: 1px solid rgba(99, 102, 241, 0.2);
        border-radius: 100px;
        font-size: 0.75rem;
        font-weight: 500;
        margin-bottom: 1rem;
    }

    /* Cards */
    .card {
        background: #111113;
        border: 1px solid #1f1f23;
        border-radius: 12px;
        padding: 1.5rem;
        margin-bottom: 1rem;
        transition: border-color 0.2s;
    }
    .card:hover { border-color: #27272a; }
    .card-title {
        font-size: 0.75rem;
        text-transform: uppercase;
        letter-spacing: 0.08em;
        color: #71717a;
        font-weight: 500;
        margin-bottom: 1rem;
    }

    /* Inputs */
    .stTextInput > div > div > input,
    .stNumberInput > div > div > input,
    .stSelectbox > div > div > div,
    .stTextArea textarea {
        background: #0a0a0b !important;
        border: 1px solid #27272a !important;
        color: #fafafa !important;
        border-radius: 8px !important;
        padding: 0.65rem 0.9rem !important;
        font-size: 0.875rem !important;
        font-family: 'Inter', sans-serif !important;
        transition: border-color 0.15s !important;
    }
    .stTextInput > div > div > input:focus,
    .stNumberInput > div > div > input:focus {
        border-color: #6366f1 !important;
        outline: none !important;
        box-shadow: 0 0 0 3px rgba(99, 102, 241, 0.1) !important;
    }
    .stTextInput label, .stNumberInput label, .stSelectbox label {
        font-size: 0.8rem !important;
        color: #a1a1aa !important;
        font-weight: 500 !important;
        margin-bottom: 0.4rem !important;
    }

    /* Buttons */
    .stButton > button {
        background: #18181b !important;
        border: 1px solid #27272a !important;
        color: #fafafa !important;
        border-radius: 8px !important;
        padding: 0.6rem 1.2rem !important;
        font-weight: 500 !important;
        font-size: 0.875rem !important;
        transition: all 0.15s !important;
        width: 100%;
    }
    .stButton > button:hover {
        background: #1f1f23 !important;
        border-color: #3f3f46 !important;
    }
    .stButton > button[kind="primary"] {
        background: #6366f1 !important;
        border: 1px solid #6366f1 !important;
        color: #fff !important;
        font-weight: 600 !important;
        box-shadow: 0 1px 2px rgba(0,0,0,0.3);
    }
    .stButton > button[kind="primary"]:hover {
        background: #4f46e5 !important;
        border-color: #4f46e5 !important;
    }

    /* File Uploader */
    .stFileUploader > div {
        background: #0a0a0b !important;
        border: 1px dashed #27272a !important;
        border-radius: 10px !important;
        padding: 2rem !important;
    }
    .stFileUploader > div:hover { border-color: #6366f1 !important; }

    /* Progress */
    .stProgress > div > div > div > div {
        background: linear-gradient(90deg, #6366f1, #8b5cf6) !important;
        border-radius: 4px !important;
    }
    .stProgress > div > div {
        background: #1f1f23 !important;
        border-radius: 4px !important;
    }

    /* Alerts */
    .stAlert {
        background: #111113 !important;
        border: 1px solid #1f1f23 !important;
        border-radius: 8px !important;
        padding: 0.9rem 1rem !important;
    }
    .stAlert-success { border-left: 3px solid #10b981 !important; }
    .stAlert-error { border-left: 3px solid #ef4444 !important; }
    .stAlert-warning { border-left: 3px solid #f59e0b !important; }
    .stAlert-info { border-left: 3px solid #6366f1 !important; }

    /* Metrics */
    .metric {
        background: #111113;
        border: 1px solid #1f1f23;
        border-radius: 10px;
        padding: 1.1rem 1.25rem;
    }
    .metric-value {
        font-size: 1.75rem;
        font-weight: 600;
        color: #fafafa;
        letter-spacing: -0.02em;
        line-height: 1;
    }
    .metric-label {
        font-size: 0.72rem;
        color: #71717a;
        text-transform: uppercase;
        letter-spacing: 0.08em;
        margin-top: 0.5rem;
        font-weight: 500;
    }
    .metric.accent .metric-value { color: #818cf8; }
    .metric.success .metric-value { color: #34d399; }
    .metric.warn .metric-value { color: #fbbf24; }
    .metric.danger .metric-value { color: #f87171; }

    /* Section divider */
    .divider {
        height: 1px;
        background: #1f1f23;
        margin: 1.5rem 0;
    }

    /* Hide clutter */
    #MainMenu, footer, header { visibility: hidden; }
    .stDeployButton { display: none !important; }

    /* Dataframe */
    .stDataFrame { border-radius: 8px; overflow: hidden; }

    /* Checkbox */
    .stCheckbox > label { color: #a1a1aa !important; }

    /* Expander */
    .streamlit-expanderHeader {
        background: #111113 !important;
        border: 1px solid #1f1f23 !important;
        border-radius: 8px !important;
    }
</style>
""", unsafe_allow_html=True)

# ============================================================
# UTILITIES (from CSE2.ipynb - preserved)
# ============================================================
def format_mm_ss(seconds):
    m, s = divmod(int(seconds), 60)
    return f"{m:02d}:{s:02d}"

def format_human(seconds):
    m, s = divmod(int(seconds), 60)
    return f"{m}m {s}s" if m else f"{s}s"

def parse_time_input(time_str):
    time_str = time_str.strip()
    for fmt in ("%I:%M %p", "%I:%M%p", "%I %p", "%I%p"):
        try: return datetime.strptime(time_str.upper(), fmt).time()
        except ValueError: pass
    for fmt in ("%H:%M", "%H:%M:%S"):
        try: return datetime.strptime(time_str, fmt).time()
        except ValueError: pass
    raise ValueError(f"Unrecognized time format: '{time_str}'")

def extract_entry_ids(prefilled_link):
    parsed = urllib.parse.urlparse(prefilled_link)
    pairs = urllib.parse.parse_qsl(parsed.query, keep_blank_values=True)
    seen, ids = set(), []
    for key, _ in pairs:
        if key.startswith("entry.") and key not in seen:
            ids.append(key); seen.add(key)
    return ids

def _parse_fb_blob(html):
    token = "FB_PUBLIC_LOAD_DATA"
    idx = html.find(token)
    if idx == -1: return None
    start = html.find("[", idx)
    if start == -1: return None
    depth, in_str, escape = 0, False, False
    for i in range(start, len(html)):
        ch = html[i]
        if in_str:
            if escape: escape = False
            elif ch == "\\": escape = True
            elif ch == '"': in_str = False
            continue
        if ch == '"': in_str = True
        elif ch == "[": depth += 1
        elif ch == "]":
            depth -= 1
            if depth == 0:
                try: return json.loads(html[start:i + 1])
                except: return None
    return None

def _options_from_blob(blob):
    result = {}
    try: items = blob[1][1]
    except: return result
    for item in items:
        if not isinstance(item, list) or len(item) < 5: continue
        question = item[4]
        if not isinstance(question, list) or not question: continue
        q0 = question[0]
        if not isinstance(q0, list) or len(q0) < 2: continue
        eid = q0[0]
        if eid is None: continue
        opts = []
        if isinstance(q0[1], list):
            for opt in q0[1]:
                if isinstance(opt, list) and opt and isinstance(opt[0], str):
                    opts.append(opt[0])
        if opts: result[f"entry.{eid}"] = opts
    return result

def _grid_from_blob(blob):
    result = {}
    try: items = blob[1][1]
    except: return result
    for item in items:
        if not isinstance(item, list) or len(item) < 5: continue
        q_label = item[1] if len(item) > 1 and isinstance(item[1], str) else ""
        question = item[4]
        if not isinstance(question, list) or len(question) < 2: continue
        sub_entries, col_options, row_labels = [], [], []
        for q_item in question:
            if not isinstance(q_item, list) or len(q_item) < 2: continue
            eid = q_item[0]
            if eid is None: continue
            opts = []
            if isinstance(q_item[1], list):
                for opt in q_item[1]:
                    if isinstance(opt, list) and opt and isinstance(opt[0], str):
                        opts.append(opt[0])
            if opts:
                sub_entries.append(f"entry.{eid}")
                if not col_options: col_options = opts
            lbl = q_item[3] if len(q_item) > 3 and isinstance(q_item[3], str) else f"entry.{eid}"
            row_labels.append(lbl)
        if len(sub_entries) >= 2:
            if len(row_labels) != len(sub_entries): row_labels = list(sub_entries)
            for i, eid in enumerate(sub_entries):
                result[eid] = {
                    "question_label": q_label, "sub_label": row_labels[i],
                    "all_sub_entries": sub_entries, "options": col_options,
                }
    return result

def fetch_form_metadata(prefilled_link):
    try: resp = requests.get(prefilled_link, timeout=30)
    except: return {}, {}, {}, []
    html = resp.text
    field_types = {}
    for m in re.finditer(r'name="(entry\.\d+)"[^>]*type="([^"]+)"', html):
        eid, itype = m.group(1), m.group(2).lower()
        if eid not in field_types: field_types[eid] = itype
        elif "checkbox" in (field_types[eid], itype): field_types[eid] = "checkbox"
    for m in re.finditer(r'<textarea[^>]*name="(entry\.\d+)"', html):
        field_types.setdefault(m.group(1), "textarea")
    for m in re.finditer(r'<select[^>]*name="(entry\.\d+)"', html):
        field_types.setdefault(m.group(1), "select")
    options_by_entry, grid_info, questions_order = {}, {}, []
    blob = _parse_fb_blob(html)
    if blob:
        options_by_entry = _options_from_blob(blob)
        grid_info = _grid_from_blob(blob)
        for eid in grid_info: field_types[eid] = "grid"
        try:
            for item in blob[1][1]:
                if not isinstance(item, list) or len(item) < 5: continue
                label = item[1] if len(item) > 1 and isinstance(item[1], str) else ""
                question = item[4]
                if not isinstance(question, list): continue
                for q_item in question:
                    if not isinstance(q_item, list) or not q_item: continue
                    eid_raw = q_item[0]
                    if eid_raw is None: continue
                    eid = f"entry.{eid_raw}"
                    ftype = field_types.get(eid, "text")
                    questions_order.append((eid, label, ftype))
        except: pass
    if not questions_order:
        for eid in extract_entry_ids(prefilled_link):
            ftype = field_types.get(eid, "text")
            questions_order.append((eid, eid, ftype))
    return field_types, options_by_entry, grid_info, questions_order

def normalize_value(raw, entry_type="text"):
    if pd.isna(raw): return "NA"
    if isinstance(raw, (pd.Timestamp, datetime)):
        if entry_type == "date": return raw.strftime("%Y-%m-%d")
        if entry_type == "time": return raw.strftime("%H:%M")
        if entry_type == "datetime-local": return raw.strftime("%Y-%m-%dT%H:%M")
    text = str(raw).strip()
    if not text: return "NA"
    if text.lower() in {"na", "n/a", "none", "null"}: return text
    if entry_type == "datetime-local" and "  " in text and "T" not in text:
        text = text.replace("  ", "T", 1)
    return text

def _case_match(value, options):
    lower_map = {o.lower(): o for o in options if isinstance(o, str)}
    return lower_map.get(value.lower())

def best_option(value, options):
    if not options: return value
    if value in options: return value
    matched = _case_match(value, options)
    if matched: return matched
    try:
        num = float(value)
        candidates = []
        for o in options:
            try: candidates.append((abs(float(o) - num), o))
            except: pass
        if candidates: return sorted(candidates)[0][1]
    except: pass
    return options[0]

def _try_parse(value, fmt):
    try:
        datetime.strptime(value, fmt); return True
    except ValueError: return False

def auto_correct(df, pending_indices, required_columns, field_types, options_by_entry):
    corrections, errors = [], []
    for idx in pending_indices:
        for col, entry in required_columns.items():
            etype = field_types.get(entry, "text")
            value = normalize_value(df.at[idx, col], etype)
            new_val, reason = value, None
            opts = options_by_entry.get(entry)
            if opts:
                if etype == "checkbox" or ", " in value:
                    parts = [v.strip() for v in value.split(", ") if v.strip()]
                    fixed = [best_option(v, opts) for v in parts]
                    candidate = ", ".join(dict.fromkeys(fixed))
                    if candidate != value: new_val, reason = candidate, "normalized checkbox options"
                else:
                    candidate = best_option(value, opts)
                    if candidate != value: new_val, reason = candidate, "normalized to allowed option"
            if etype == "date":
                try:
                    new_val = pd.to_datetime(value).strftime("%Y-%m-%d")
                    if new_val != value: reason = "normalized date"
                except: errors.append((idx, col, entry, value, "invalid date")); continue
            elif etype == "time":
                try:
                    new_val = pd.to_datetime(value).strftime("%H:%M")
                    if new_val != value: reason = "normalized time"
                except: errors.append((idx, col, entry, value, "invalid time")); continue
            elif etype == "datetime-local":
                try:
                    new_val = pd.to_datetime(value).strftime("%Y-%m-%dT%H:%M")
                    if new_val != value: reason = "normalized datetime"
                except: errors.append((idx, col, entry, value, "invalid datetime")); continue
            if new_val != value:
                df.at[idx, col] = new_val
                corrections.append((idx, col, entry, value, new_val, reason))
    return corrections, errors

def validate_data(df, pending_indices, required_columns, field_types, options_by_entry):
    errors = []
    for idx in pending_indices:
        for col, entry in required_columns.items():
            etype = field_types.get(entry, "text")
            value = normalize_value(df.at[idx, col], etype)
            if etype in {"radio", "select"} and ", " in value:
                errors.append((idx, col, entry, value, "multiple values for single-choice")); continue
            opts = options_by_entry.get(entry)
            if opts:
                parts = [v.strip() for v in value.split(", ") if v.strip()] if (etype == "checkbox" or ", " in value) else [value]
                bad = [v for v in parts if v not in opts]
                if bad:
                    errors.append((idx, col, entry, value, f"not in options: {', '.join(bad)}")); continue
            if etype == "date":
                if not _try_parse(value, "%Y-%m-%d"):
                    errors.append((idx, col, entry, value, "invalid date (YYYY-MM-DD)"))
            elif etype == "time":
                if not any(_try_parse(value, f) for f in ("%H:%M", "%H:%M:%S")):
                    errors.append((idx, col, entry, value, "invalid time (HH:MM or HH:MM:SS)"))
            elif etype == "datetime-local":
                if not any(_try_parse(value, f) for f in ("%Y-%m-%dT%H:%M", "%Y-%m-%dT%H:%M:%S")):
                    errors.append((idx, col, entry, value, "invalid datetime (YYYY-MM-DDTHH:MM)"))
    return errors

def map_columns(df, prefilled_link, status_col):
    entry_cols = [c for c in df.columns if c.startswith("entry.")]
    prefilled_ids = extract_entry_ids(prefilled_link)
    non_status_cols = [c for c in df.columns if c != status_col]
    if entry_cols: return {c: c for c in entry_cols}
    if prefilled_ids:
        matched = [c for c in prefilled_ids if c in df.columns]
        if matched: return {c: c for c in matched}
        if len(prefilled_ids) == len(non_status_cols):
            return dict(zip(non_status_cols, prefilled_ids))
    return {}

def wait_until(target_dt):
    while True:
        remaining = (target_dt - datetime.now(IST)).total_seconds()
        if remaining <= 0: break
        slp = max(0.1, remaining - random.uniform(1.5, 4.5))
        time.sleep(min(slp, remaining))

def submit_row(row, required_columns, field_types, form_url, prefilled_link, max_retries=3):
    payload = {}
    for col, entry in required_columns.items():
        etype = field_types.get(entry, "text")
        value = normalize_value(row[col], etype)
        if ", " in value: payload[entry] = [v.strip() for v in value.split(", ")]
        else: payload[entry] = value
    session = requests.Session()
    headers = {
        "User-Agent": "Mozilla/5.0",
        "Referer": prefilled_link,
        "Origin": "https://docs.google.com",
    }
    for _ in range(max_retries):
        try:
            r = session.post(form_url, data=payload, headers=headers, timeout=30)
            if r.status_code == 200 and ("Your response has been recorded" in r.text or "formResponse" in r.url):
                return 200, "Success"
            return r.status_code, f"HTTP {r.status_code}"
        except: pass
        time.sleep(2)
    return 0, "Failed after retries"

def generate_analysis_sheet(df, file_path, submitted_count):
    if not OPENPYXL_AVAILABLE: return
    lower = file_path.lower()
    if not (lower.endswith(".xlsx") or lower.endswith(".xls")): return
    wb = load_workbook(file_path)
    if "Analysis" in wb.sheetnames: del wb["Analysis"]
    ws = wb.create_sheet("Analysis")
    NAVY, BLUE, LIGHT_BLUE, WHITE = "1F3864", "2E75B6", "DEEAF1", "FFFFFF"
    solid = lambda c: PatternFill("solid", fgColor=c)
    thin = Side(style="thin", color="CCCCCC")
    bdr = Border(left=thin, right=thin, top=thin, bottom=thin)
    ctr = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.merge_cells("A1:F1")
    ws["A1"] = "Response Analysis"
    ws["A1"].fill, ws["A1"].font = solid(NAVY), Font(bold=True, color=WHITE, size=14)
    ws["A1"].alignment = ctr
    ws.row_dimensions[1].height = 28
    ws["A2"] = f"Generated: {datetime.now().strftime('%d-%m-%Y %I:%M %p')}"
    ws["B2"] = f"Total rows: {len(df)}"
    ws["C2"] = f"Submitted: {submitted_count}"
    for cell in (ws["A2"], ws["B2"], ws["C2"]):
        cell.font = Font(color="444444", size=10, italic=True)
    ws.column_dimensions["A"].width = 5
    ws.column_dimensions["B"].width = 32
    ws.column_dimensions["C"].width = 10
    ws.column_dimensions["D"].width = 12
    status_names = {"status", "submit", "submitted"}
    analysis_cols = [c for c in df.columns if c.lower() not in status_names]
    current_row = 4
    for col_name in analysis_cols:
        series = df[col_name].dropna().astype(str).str.strip()
        series = series[series != ""]
        if series.empty: continue
        is_checkbox = series.str.contains(", ").any()
        ws.merge_cells(start_row=current_row, start_column=1, end_row=current_row, end_column=5)
        cell = ws.cell(row=current_row, column=1, value=col_name)
        cell.fill, cell.font, cell.alignment = solid(BLUE), Font(bold=True, color=WHITE, size=11), ctr
        ws.row_dimensions[current_row].height = 20
        current_row += 1
        hdr_row = current_row
        for ci, h in enumerate(["#", "Answer", "Count", "% of Total"], 1):
            c = ws.cell(row=current_row, column=ci, value=h)
            c.fill, c.font, c.alignment, c.border = solid(BLUE), Font(bold=True, color=WHITE, size=10), ctr, bdr
        current_row += 1
        if is_checkbox:
            all_vals = []
            for v in series: all_vals.extend(x.strip() for x in v.split(", ") if x.strip())
            freq, total_responders = Counter(all_vals), len(series)
        else:
            freq, total_responders = Counter(series.tolist()), len(series)
        items_sorted = sorted(freq.items(), key=lambda x: -x[1])
        total_count = sum(freq.values())
        data_start = current_row
        for ri, (answer, count) in enumerate(items_sorted, 1):
            pct = (count / total_responders * 100) if total_responders else 0
            row_vals = [ri, answer, count, f"{pct:.1f}%"]
            for ci, val in enumerate(row_vals, 1):
                c = ws.cell(row=current_row, column=ci, value=val)
                c.fill = solid(LIGHT_BLUE if ri % 2 == 0 else WHITE)
                c.font, c.alignment, c.border = Font(color="1F1F1F", size=10), ctr, bdr
            current_row += 1
        total_row_idx = current_row
        total_pct = "100%" if not is_checkbox else "---"
        for ci, val in enumerate(["", "TOTAL", total_count, total_pct], 1):
            c = ws.cell(row=current_row, column=ci, value=val)
            c.fill, c.font, c.alignment, c.border = solid(NAVY), Font(bold=True, color=WHITE, size=10), ctr, bdr
        current_row += 1
        if items_sorted:
            chart = BarChart()
            chart.type, chart.grouping = "col", "clustered"
            chart.title = str(col_name)[:30]
            chart.y_axis.title, chart.x_axis.title = "Count", "Answer"
            chart.style, chart.width, chart.height = 10, 18, 12
            data_ref = Reference(ws, min_col=3, max_col=3, min_row=hdr_row, max_row=total_row_idx - 1)
            cats_ref = Reference(ws, min_col=2, max_col=2, min_row=data_start, max_row=total_row_idx - 1)
            chart.add_data(data_ref, titles_from_data=True)
            chart.set_categories(cats_ref)
            ws.add_chart(chart, f"{get_column_letter(7)}{data_start}")
        current_row += 2
    wb.save(file_path)

# ============================================================
# MAIN APP
# ============================================================
def main():
    # Hero
    st.markdown("""
    <div class="hero">
        <div class="hero-badge">⚡ Automated Form Submitter</div>
        <h1>formIntellect</h1>
        <p class="hero-sub">Submit Google Form responses with precision timing and smart validation.</p>
    </div>
    """, unsafe_allow_html=True)

    # ---- Section 1: Inputs ----
    st.markdown("### 01  ·  Configuration")
    st.markdown('<div class="card">', unsafe_allow_html=True)

    c1, c2 = st.columns([2, 1])
    with c1:
        prefilled_link = st.text_input(
            "Google Form Pre-filled Link",
            placeholder="https://docs.google.com/forms/d/e/.../viewform?usp=pp_url&entry..."
        )
    with c2:
        source_type = st.selectbox("Data Source", ["Upload File", "Google Sheets URL"])

    df, output_path = None, "responses_output.xlsx"

    if source_type == "Upload File":
        uploaded_file = st.file_uploader("Drop file here", type=["xlsx", "xls", "csv", "tsv", "ods", "json"])
        if uploaded_file:
            try:
                name = uploaded_file.name.lower()
                if name.endswith('.csv'): df = pd.read_csv(uploaded_file)
                elif name.endswith('.tsv'): df = pd.read_csv(uploaded_file, sep='\t')
                elif name.endswith('.json'): df = pd.read_json(uploaded_file)
                elif name.endswith('.ods'): df = pd.read_excel(uploaded_file, engine='odf')
                else: df = pd.read_excel(uploaded_file)
                output_path = uploaded_file.name
                st.success(f"Loaded **{len(df)}** rows · **{len(df.columns)}** columns")
            except Exception as e:
                st.error(f"Error: {e}")
    else:
        sheets_url = st.text_input("Google Sheets CSV Export URL", placeholder="https://docs.google.com/spreadsheets/.../export?format=csv")
        if sheets_url and st.button("Load Sheet"):
            try:
                resp = requests.get(sheets_url, timeout=30)
                df = pd.read_csv(io.StringIO(resp.text))
                st.success(f"Loaded **{len(df)}** rows")
            except Exception as e:
                st.error(f"Error: {e}")

    st.markdown('</div>', unsafe_allow_html=True)

    if df is None:
        st.markdown('<div class="divider"></div>', unsafe_allow_html=True)
        st.info("Upload a file to begin.")
        return

    df.columns = df.columns.str.strip()

    # Status column
    status_col = None
    for col in df.columns:
        if col.strip().lower() in {"status", "submit", "submitted"}:
            status_col = col; break
    if status_col is None:
        status_col = "Status"
        df[status_col] = ""

    required_columns = map_columns(df, prefilled_link, status_col)

    # ---- Section 2: Validation ----
    st.markdown("### 02  ·  Validation & Mapping")
    st.markdown('<div class="card">', unsafe_allow_html=True)

    status_series = df[status_col].astype(str).str.lower().fillna("")
    pending_indices = df[status_series != "submitted"].index.tolist()

    col_a, col_b, col_c = st.columns(3)
    with col_a:
        st.markdown(f'<div class="metric"><div class="metric-value">{len(df)}</div><div class="metric-label">Total Rows</div></div>', unsafe_allow_html=True)
    with col_b:
        st.markdown(f'<div class="metric accent"><div class="metric-value">{len(pending_indices)}</div><div class="metric-label">Pending</div></div>', unsafe_allow_html=True)
    with col_c:
        st.markdown(f'<div class="metric success"><div class="metric-value">{len(required_columns)}</div><div class="metric-label">Mapped Fields</div></div>', unsafe_allow_html=True)

    if st.button("🔍  Validate & Auto-Correct", use_container_width=True):
        if not prefilled_link:
            st.error("Please enter a pre-filled link first.")
        elif not required_columns:
            st.error("Could not auto-map columns. Ensure headers match entry.XXXXXX or column count matches form fields.")
        else:
            with st.spinner("Fetching form metadata..."):
                field_types, options_by_entry, grid_info, questions_order = fetch_form_metadata(prefilled_link)

            st.session_state['field_types'] = field_types
            st.session_state['options_by_entry'] = options_by_entry

            corrections, corr_errors = auto_correct(df, pending_indices, required_columns, field_types, options_by_entry)
            if corrections:
                st.success(f"Auto-corrected **{len(corrections)}** values")
            if corr_errors:
                st.error(f"Unable to auto-correct **{len(corr_errors)}** values")
                for idx, col, entry, value, reason in corr_errors[:5]:
                    st.write(f"• Row {idx+1} · `{col}` · {value} · {reason}")

            errors = validate_data(df, pending_indices, required_columns, field_types, options_by_entry)
            if errors:
                st.error(f"Validation failed · **{len(errors)}** errors")
                for idx, col, entry, value, reason in errors[:5]:
                    st.write(f"• Row {idx+1} · `{col}` · {value} · {reason}")
            else:
                st.success("All data validated successfully ✓")
                st.session_state['validated'] = True

    st.markdown('</div>', unsafe_allow_html=True)

    # ---- Section 3: Schedule ----
    st.markdown("### 03  ·  Schedule")
    st.markdown('<div class="card">', unsafe_allow_html=True)

    now_ist = datetime.now(IST)
    default_start = now_ist.strftime("%H:%M")
    default_end = (now_ist + timedelta(minutes=3)).strftime("%H:%M")

    c1, c2 = st.columns(2)
    with c1:
        start_time_str = st.text_input("Start Time", value=default_start, help="12-hour (10:40 AM) or 24-hour (10:40)")
    with c2:
        end_time_str = st.text_input("End Time", value=default_end, help="12-hour (10:43 AM) or 24-hour (10:43)")

    c3, c4 = st.columns(2)
    with c3:
        use_random = st.checkbox("Randomize row order", value=True)
        subset_n = st.number_input("Max rows (0 = all)", min_value=0, value=0)
    with c4:
        delay_min = st.number_input("Min delay (sec)", min_value=1, value=2)
        delay_max = st.number_input("Max delay (sec)", min_value=1, value=5)

    st.markdown('</div>', unsafe_allow_html=True)
    st.markdown('<div class="divider"></div>', unsafe_allow_html=True)

    # ---- Launch ----
    if st.button("⚡  Launch Campaign", type="primary", use_container_width=True):
        if not prefilled_link or "/viewform" not in prefilled_link:
            st.error("Valid pre-filled link is required.")
            return
        if 'validated' not in st.session_state:
            st.error("Please validate data first.")
            return
        if not required_columns:
            st.error("No column mapping found.")
            return

        form_url = prefilled_link.split("/viewform")[0] + "/formResponse"
        field_types = st.session_state.get('field_types', {})
        options_by_entry = st.session_state.get('options_by_entry', {})

        if use_random: random.shuffle(pending_indices)
        if subset_n > 0: pending_indices = pending_indices[:subset_n]

        total_rows = len(pending_indices)
        if total_rows == 0:
            st.warning("No pending rows to submit.")
            return

        # Time window config (mirrors CSE2.ipynb configure_time_window logic)
        min_required_sec = total_rows * MIN_SECONDS_PER_ROW
        try:
            start_time = parse_time_input(start_time_str)
            end_time = parse_time_input(end_time_str)
            today = now_ist.date()
            start_dt = IST.localize(datetime.combine(today, start_time))
            end_dt = IST.localize(datetime.combine(today, end_time))
        except ValueError as e:
            st.error(f"Time error: {e}")
            return

        if end_dt <= start_dt: end_dt += timedelta(days=1)
        original_window_sec = int((end_dt - start_dt).total_seconds())

        # Auto-swap if min > max for delay ranges
        if delay_min > delay_max:
            delay_min, delay_max = delay_max, delay_min

        # Try up to 1000 combinations satisfying constraints
        rs = re_val = rand_start = rand_end = window_sec = 0
        found = False
        for _ in range(1000):
            rs = random.randint(delay_min, delay_max)
            re_val = random.randint(delay_min * 2, delay_max * 2)
            rand_start = start_dt + timedelta(seconds=rs)
            rand_end = end_dt + timedelta(seconds=re_val)
            if rand_end <= rand_start: rand_end = rand_start + timedelta(seconds=1)
            window_sec = int((rand_end - rand_start).total_seconds())
            if window_sec > original_window_sec and window_sec >= min_required_sec:
                found = True; break

        if not found:
            # Fallback: expand end window
            window_sec = max(total_rows + 5, original_window_sec + 10)
            rand_start = start_dt + timedelta(seconds=delay_min)
            rand_end = rand_start + timedelta(seconds=window_sec)

        # Summary metrics
        st.markdown('<div class="divider"></div>', unsafe_allow_html=True)
        st.markdown("### 04  ·  Execution")

        m1, m2, m3, m4 = st.columns(4)
        with m1: st.markdown(f'<div class="metric accent"><div class="metric-value">{total_rows}</div><div class="metric-label">Rows</div></div>', unsafe_allow_html=True)
        with m2: st.markdown(f'<div class="metric"><div class="metric-value">{format_mm_ss(window_sec)}</div><div class="metric-label">Window</div></div>', unsafe_allow_html=True)
        with m3: st.markdown(f'<div class="metric success"><div class="metric-value">{rand_start.strftime("%H:%M:%S")}</div><div class="metric-label">Start</div></div>', unsafe_allow_html=True)
        with m4: st.markdown(f'<div class="metric warn"><div class="metric-value">{rand_end.strftime("%H:%M:%S")}</div><div class="metric-label">End</div></div>', unsafe_allow_html=True)

        # Wait until start
        now = datetime.now(IST)
        if now < rand_start:
            wait_sec = int((rand_start - now).total_seconds())
            st.info(f"⏳ Waiting **{wait_sec}s** until start time...")
            time.sleep(wait_sec)

        # Execution
        offsets = sorted(random.sample(range(1, max(window_sec, total_rows) + 1), total_rows))
        start_ref = datetime.now(IST)
        submitted_count = 0

        progress_bar = st.progress(0)
        status_text = st.empty()
        log_container = st.container()

        for i, (offset, idx) in enumerate(zip(offsets, pending_indices), 1):
            target = start_ref + timedelta(seconds=offset)
            wait_until(target)

            status_code, reason = submit_row(df.loc[idx], required_columns, field_types, form_url, prefilled_link)
            ist_now = datetime.now(IST).strftime("%H:%M:%S")

            if status_code == 200:
                df.at[idx, status_col] = "submitted"
                submitted_count += 1
                status_text.success(f"✅ Row {idx+1} submitted at {ist_now}")
                if i % 5 == 0:
                    try:
                        if output_path.lower().endswith(('.xlsx', '.xls')): df.to_excel(output_path, index=False)
                        else: df.to_csv(output_path, index=False)
                    except: pass
            else:
                status_text.error(f"❌ Row {idx+1} failed · {reason}")

            progress_bar.progress(i / total_rows)

        # Final save
        try:
            if output_path.lower().endswith(('.xlsx', '.xls')): df.to_excel(output_path, index=False)
            else: df.to_csv(output_path, index=False)
        except: pass

        st.balloons()
        st.success(f"🎉 Complete · **{submitted_count}/{total_rows}** submitted")

        if output_path.lower().endswith(('.xlsx', '.xls')):
            with st.spinner("Generating analysis sheet..."):
                generate_analysis_sheet(df, output_path, submitted_count)
            try:
                st.download_button(
                    "📥  Download Updated File",
                    data=open(output_path, 'rb').read(),
                    file_name=output_path,
                    mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                    use_container_width=True
                )
            except: pass

if __name__ == "__main__":
    main()
