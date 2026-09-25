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
    import openpyxl
    from openpyxl import load_workbook
    from openpyxl.chart import BarChart, PieChart, Reference
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    from openpyxl.utils import get_column_letter
    OPENPYXL_AVAILABLE = True
except ImportError:
    OPENPYXL_AVAILABLE = False

IST = pytz.timezone("Asia/Kolkata")
MIN_SECONDS_PER_ROW = 2

st.set_page_config(page_title="Form Auto-Submitter", layout="wide", page_icon="📝")

# ============================================================
#  EXACT NOTEBOOK FUNCTIONS (Unaltered Logic)
# ============================================================

def format_mm_ss(seconds):
    m, s = divmod(int(seconds), 60)
    return f"{m:02d}:{s:02d}"

def format_human(seconds):
    m, s = divmod(int(seconds), 60)
    if m: return f"{m}m {s}s"
    return f"{s}s"

def _try_parse(value, fmt):
    try:
        datetime.strptime(value, fmt)
        return True
    except ValueError:
        return False

def parse_time_input(time_str):
    time_str = time_str.strip()
    for fmt in ("%I:%M %p", "%I:%M%p", "%I %p", "%I%p"):
        try: return datetime.strptime(time_str.upper(), fmt).time()
        except ValueError: pass
    for fmt in ("%H:%M", "%H:%M:%S"):
        try: return datetime.strptime(time_str, fmt).time()
        except ValueError: pass
    raise ValueError(f"Unrecognized time format: '{time_str}'. Use 12-hour (e.g. 09:30 AM) or 24-hour (e.g. 13:30).")

def extract_entry_ids(prefilled_link):
    parsed = urllib.parse.urlparse(prefilled_link)
    pairs  = urllib.parse.parse_qsl(parsed.query, keep_blank_values=True)
    seen, ids = set(), []
    for key, _ in pairs:
        if key.startswith("entry.") and key not in seen:
            ids.append(key)
            seen.add(key)
    return ids

def _parse_fb_blob(html):
    token = "FB_PUBLIC_LOAD_DATA"
    idx   = html.find(token)
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
                except json.JSONDecodeError: return None
    return None

def _options_from_blob(blob):
    result = {}
    try: items = blob[1][1]
    except Exception: return result
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
    except Exception: return result
    for item in items:
        if not isinstance(item, list) or len(item) < 5: continue
        q_label  = item[1] if len(item) > 1 and isinstance(item[1], str) else ""
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
                result[eid] = {"question_label": q_label, "sub_label": row_labels[i], "all_sub_entries": sub_entries, "options": col_options}
    return result

def fetch_form_metadata(prefilled_link):
    try: resp = requests.get(prefilled_link, timeout=30)
    except requests.exceptions.RequestException: return {}, {}, {}, []
    html, field_types = resp.text, {}
    for m in re.finditer(r'name="(entry\.\d+)"[^>]*type="([^"]+)"', html):
        eid, itype = m.group(1), m.group(2).lower()
        if eid not in field_types: field_types[eid] = itype
        elif "checkbox" in (field_types[eid], itype): field_types[eid] = "checkbox"
    for m in re.finditer(r'<textarea[^>]*name="(entry\.\d+)"', html): field_types.setdefault(m.group(1), "textarea")
    for m in re.finditer(r'<select[^>]*name="(entry\.\d+)"', html): field_types.setdefault(m.group(1), "select")
    
    options_by_entry, grid_info, questions_order = {}, {}, []
    blob = _parse_fb_blob(html)
    if blob:
        options_by_entry = _options_from_blob(blob)
        grid_info        = _grid_from_blob(blob)
        for eid in grid_info: field_types[eid] = "grid"
        try:
            for item in blob[1][1]:
                if not isinstance(item, list) or len(item) < 5: continue
                label    = item[1] if len(item) > 1 and isinstance(item[1], str) else ""
                question = item[4]
                if not isinstance(question, list): continue
                for q_item in question:
                    if not isinstance(q_item, list) or not q_item: continue
                    eid_raw = q_item[0]
                    if eid_raw is None: continue
                    eid   = f"entry.{eid_raw}"
                    ftype = field_types.get(eid, "text")
                    questions_order.append((eid, label, ftype))
        except Exception: pass
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
    if entry_type == "datetime-local" and "  " in text and "T" not in text: text = text.replace("  ", "T", 1)
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
            except (ValueError, TypeError): pass
        if candidates: return sorted(candidates)[0][1]
    except ValueError: pass
    return options[0]

def auto_correct(df, pending_indices, required_columns, field_types, options_by_entry):
    corrections, errors = [], []
    for idx in pending_indices:
        for col, entry in required_columns.items():
            etype  = field_types.get(entry, "text")
            value  = normalize_value(df.at[idx, col], etype)
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
                except Exception: errors.append((idx, col, entry, value, "invalid date")); continue
            elif etype == "time":
                try:
                    new_val = pd.to_datetime(value).strftime("%H:%M")
                    if new_val != value: reason = "normalized time"
                except Exception: errors.append((idx, col, entry, value, "invalid time")); continue
            elif etype == "datetime-local":
                try:
                    new_val = pd.to_datetime(value).strftime("%Y-%m-%dT%H:%M")
                    if new_val != value: reason = "normalized datetime"
                except Exception: errors.append((idx, col, entry, value, "invalid datetime")); continue
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
                parts = ([v.strip() for v in value.split(", ") if v.strip()] if (etype == "checkbox" or ", " in value) else [value])
                bad = [v for v in parts if v not in opts]
                if bad: errors.append((idx, col, entry, value, f"not in options: {', '.join(bad)}")); continue
            if etype == "date":
                if not _try_parse(value, "%Y-%m-%d"): errors.append((idx, col, entry, value, "invalid date (YYYY-MM-DD)"))
            elif etype == "time":
                if not any(_try_parse(value, f) for f in ("%H:%M", "%H:%M:%S")): errors.append((idx, col, entry, value, "invalid time (HH:MM or HH:MM:SS)"))
            elif etype == "datetime-local":
                if not any(_try_parse(value, f) for f in ("%Y-%m-%dT%H:%M", "%Y-%m-%dT%H:%M:%S")): errors.append((idx, col, entry, value, "invalid datetime (YYYY-MM-DDTHH:MM)"))
    return errors

def configure_time_window(total_rows, start_input, end_input, start_rand_min, start_rand_max, end_rand_min, end_rand_max):
    min_required_sec = total_rows * MIN_SECONDS_PER_ROW
    today = datetime.now(IST).date()
    try:
        start_time = parse_time_input(start_input)
        end_time   = parse_time_input(end_input)
        start_dt   = IST.localize(datetime.combine(today, start_time))
        end_dt     = IST.localize(datetime.combine(today, end_time))
    except ValueError as e:
        raise ValueError(str(e))

    if end_dt <= start_dt: end_dt += timedelta(days=1)
    original_window_sec = int((end_dt - start_dt).total_seconds())

    if start_rand_min > start_rand_max: start_rand_min, start_rand_max = start_rand_max, start_rand_min
    if end_rand_min > end_rand_max: end_rand_min, end_rand_max = end_rand_max, end_rand_min

    rs = re_val = rand_start = rand_end = window_sec = 0
    found = False
    for _ in range(1000):
        rs  = random.randint(start_rand_min, start_rand_max)
        re_val  = random.randint(end_rand_min,   end_rand_max)
        rand_start = start_dt + timedelta(seconds=rs)
        rand_end   = end_dt   + timedelta(seconds=re_val)
        if rand_end <= rand_start: rand_end = rand_start + timedelta(seconds=1)
        window_sec = int((rand_end - rand_start).total_seconds())
        if window_sec > original_window_sec and window_sec >= min_required_sec:
            found = True; break

    if not found:
        raise ValueError("Unable to find a valid randomization. Ensure end_rand_max > start_rand_max and widen your base time range.")

    return rand_start, rand_end, window_sec

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
    headers = {"User-Agent": "Mozilla/5.0", "Referer": prefilled_link, "Origin": "https://docs.google.com"}
    for _ in range(max_retries):
        try:
            r = session.post(form_url, data=payload, headers=headers, timeout=30)
            if r.status_code == 200 and ("Your response has been recorded" in r.text or "formResponse" in r.url):
                return 200, "Success"
            return r.status_code, f"HTTP {r.status_code}: {r.text[:200]}"
        except requests.exceptions.RequestException: pass
        time.sleep(2)
    return 0, "Failed after max retries"

def map_columns(df, prefilled_link, status_col):
    entry_cols      = [c for c in df.columns if c.startswith("entry.")]
    prefilled_ids   = extract_entry_ids(prefilled_link)
    non_status_cols = [c for c in df.columns if c != status_col]
    if entry_cols: return {c: c for c in entry_cols}
    if prefilled_ids:
        matched = [c for c in prefilled_ids if c in df.columns]
        if matched: return {c: c for c in matched}
        if len(prefilled_ids) == len(non_status_cols): return dict(zip(non_status_cols, prefilled_ids))
    raise ValueError("Unable to map columns to form entry IDs.")

def generate_analysis_sheet_buffer(df, submitted_count):
    buffer = io.BytesIO()
    df.to_excel(buffer, index=False, engine='openpyxl')
    buffer.seek(0)
    wb = load_workbook(buffer)
    if "Analysis" in wb.sheetnames: del wb["Analysis"]
    ws = wb.create_sheet("Analysis")
    NAVY, BLUE, LIGHT_BLUE, WHITE = "1F3864", "2E75B6", "DEEAF1", "FFFFFF"
    def solid(color): return PatternFill("solid", fgColor=color)
    thin = Side(style="thin", color="CCCCCC")
    bdr  = Border(left=thin, right=thin, top=thin, bottom=thin)
    ctr  = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.merge_cells("A1:F1")
    ws["A1"] = "Response Analysis"
    ws["A1"].fill, ws["A1"].font, ws["A1"].alignment = solid(NAVY), Font(bold=True, color=WHITE, size=14), ctr
    ws.row_dimensions[1].height = 28
    ws["A2"] = f"Generated: {datetime.now().strftime('%d-%m-%Y %I:%M %p')}"
    ws["B2"] = f"Total rows: {len(df)}"
    ws["C2"] = f"Submitted: {submitted_count}"
    for cell in (ws["A2"], ws["B2"], ws["C2"]): cell.font = Font(color="444444", size=10, italic=True)
    ws.column_dimensions["A"].width, ws.column_dimensions["B"].width = 5, 32
    ws.column_dimensions["C"].width, ws.column_dimensions["D"].width = 10, 12
    status_names  = {"status", "submit", "submitted"}
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
        total_count  = sum(freq.values())
        data_start   = current_row
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
            chart.type, chart.grouping, chart.title = "col", "clustered", str(col_name)[:30]
            chart.y_axis.title, chart.x_axis.title, chart.style = "Count", "Answer", 10
            chart.width, chart.height = 18, 12
            data_ref = Reference(ws, min_col=3, max_col=3, min_row=hdr_row, max_row=total_row_idx - 1)
            cats_ref = Reference(ws, min_col=2, max_col=2, min_row=data_start, max_row=total_row_idx - 1)
            chart.add_data(data_ref, titles_from_data=True)
            chart.set_categories(cats_ref)
            ws.add_chart(chart, f"{get_column_letter(7)}{data_start}")
        current_row += 2
    out_buffer = io.BytesIO()
    wb.save(out_buffer)
    out_buffer.seek(0)
    return out_buffer

# ============================================================
#  STREAMLIT UI & EXECUTION
# ============================================================

st.title("📝 Google Form Auto-Submitter")
st.caption("Exact replica of the original notebook logic.")

with st.form("submission_form"):
    st.subheader("1. File & Form Setup")
    col1, col2 = st.columns(2)
    with col1:
        uploaded_file = st.file_uploader("Upload Data File (.xlsx/.xls/.csv/.tsv/.ods/.json)", type=["xlsx", "xls", "csv", "tsv", "ods", "json"])
    with col2:
        sheets_url = st.text_input("Or enter Google Sheets URL:", placeholder="https://docs.google.com/spreadsheets/...")
    
    prefilled_link = st.text_input("Enter Google Form pre-filled link:")
    
    st.subheader("2. Row Selection")
    col3, col4 = st.columns(2)
    with col3:
        use_random = st.checkbox("Use random rows? (y/n)", value=False)
    with col4:
        subset_n = st.number_input("How many random rows to submit? (0 = all)", min_value=0, value=0)
        
    st.subheader("3. Time Configuration")
    st.caption("Time formats accepted: 12-hour (e.g. 09:30 AM, 11:00 PM) or 24-hour (e.g. 21:30)")
    col5, col6 = st.columns(2)
    with col5:
        start_input = st.text_input("Start time:", value="10:40")
        start_rand_min = st.number_input("Random delay AFTER start (min seconds):", min_value=0, value=1)
        start_rand_max = st.number_input("Random delay AFTER start (max seconds):", min_value=0, value=1)
    with col6:
        end_input = st.text_input("End time:", value="10:43")
        end_rand_min = st.number_input("Random delay AFTER end (min seconds):", min_value=0, value=2)
        end_rand_max = st.number_input("Random delay AFTER end (max seconds):", min_value=0, value=3)

    submitted = st.form_submit_button("🚀 Start Submission Process", type="primary", use_container_width=True)

if submitted:
    # 1. Load Data
    try:
        if sheets_url:
            resp = requests.get(sheets_url, timeout=30)
            df = pd.read_csv(io.StringIO(resp.text))
        elif uploaded_file is not None:
            if uploaded_file.name.endswith('.csv'): df = pd.read_csv(uploaded_file)
            elif uploaded_file.name.endswith('.tsv'): df = pd.read_csv(uploaded_file, sep='\t')
            elif uploaded_file.name.endswith('.ods'): df = pd.read_excel(uploaded_file, engine='odf')
            elif uploaded_file.name.endswith('.json'): df = pd.read_json(uploaded_file)
            else: df = pd.read_excel(uploaded_file)
        else:
            st.error("Please upload a file or provide a Google Sheets URL.")
            st.stop()
    except Exception as e:
        st.error(f"Error loading file: {e}")
        st.stop()

    df.columns = df.columns.str.strip()
    form_url = prefilled_link.split("/viewform")[0] + "/formResponse"

    # 2. Map Columns
    status_col = None
    for col in df.columns:
        if col.strip().lower() in {"status", "submit", "submitted"}:
            status_col = col; break
    if status_col is None:
        status_col = "Status"
        df[status_col] = ""

    try:
        required_columns = map_columns(df, prefilled_link, status_col)
    except ValueError as e:
        st.error(str(e))
        st.stop()

    status_series   = df[status_col].astype(str).str.lower().fillna("")
    pending_indices = df[status_series != "submitted"].index.tolist()

    if use_random:
        random.shuffle(pending_indices)
        if subset_n > 0:
            pending_indices = random.sample(pending_indices, min(subset_n, len(pending_indices)))

    total_rows = len(pending_indices)
    if total_rows == 0:
        st.error("No pending rows to submit.")
        st.stop()

    # 3. Fetch Metadata & Validate
    with st.spinner("Fetching form metadata..."):
        field_types, options_by_entry, grid_info, _ = fetch_form_metadata(prefilled_link)
    
    corrections, corr_errors = auto_correct(df, pending_indices, required_columns, field_types, options_by_entry)
    if corr_errors:
        st.error("Data Validation Failed. Unable to auto-correct.")
        for idx, col, entry, value, reason in corr_errors:
            st.write(f"Row {idx + 1} | Column: {col} | Field: {entry} | Value: {value} | Reason: {reason}")
        st.stop()

    # 4. Configure Time Window
    try:
        rand_start, rand_end, window_sec = configure_time_window(
            total_rows, start_input, end_input, 
            start_rand_min, start_rand_max, 
            end_rand_min, end_rand_max
        )
    except ValueError as e:
        st.error(str(e))
        st.stop()

    if window_sec <= total_rows:
        window_sec = max(total_rows + 5, window_sec)

    st.success(f"Time Window Configured: {rand_start.strftime('%I:%M:%S %p')} to {rand_end.strftime('%I:%M:%S %p')} IST")

    # 5. Wait for Start Time
    now = datetime.now(IST)
    if now < rand_start:
        wait_sec = int((rand_start - now).total_seconds())
        with st.status(f"Waiting until START time ({rand_start.strftime('%I:%M:%S %p IST')})...", expanded=True) as status:
            st.write(f"Sleeping for {format_human(wait_sec)}...")
            time.sleep(wait_sec)
            status.update(label="Start time reached! Beginning submissions.", state="complete")

    # 6. Submission Loop with Tabular Log
    offsets = sorted(random.sample(range(1, window_sec + 1), total_rows))
    start_ref = datetime.now(IST)
    submitted_count = 0

    log_df = pd.DataFrame(columns=["Row", "Status", "Submitted At", "HTTP Code", "Next Entry", "Gap"])
    log_placeholder = st.empty()
    st.caption("Live Submission Log:")

    for i, (offset, idx) in enumerate(zip(offsets, pending_indices), 1):
        target = start_ref + timedelta(seconds=offset)
        wait_until(target)

        status_code, reason = submit_row(df.loc[idx], required_columns, field_types, form_url, prefilled_link)
        ist_now = datetime.now(IST).strftime("%d-%m-%Y %I:%M:%S %p")

        if status_code == 200:
            df.at[idx, status_col] = "submitted"
            submitted_count += 1
            status_text = "Submitted"
            
            if i < total_rows:
                gap_sec = offsets[i] - offsets[i - 1]
                next_time = (start_ref + timedelta(seconds=offsets[i])).strftime("%d-%m-%Y %I:%M:%S %p IST")
                gap_text = format_human(gap_sec)
            else:
                next_time, gap_text = "FINAL ENTRY", "-"
        else:
            status_text = f"FAILED: {reason}"
            next_time, gap_text = "-", "-"

        new_row = {
            "Row": idx + 1, 
            "Status": status_text, 
            "Submitted At": ist_now, 
            "HTTP Code": status_code, 
            "Next Entry": next_time, 
            "Gap": gap_text
        }
        log_df = pd.concat([log_df, pd.DataFrame([new_row])], ignore_index=True)
        log_placeholder.dataframe(log_df, use_container_width=True, hide_index=True)

    st.balloons()
    st.success(f"PROCESS COMPLETED. Successfully submitted {submitted_count} rows.")

    # 7. Generate Analysis Sheet
    if OPENPYXL_AVAILABLE and submitted_count > 0:
        with st.spinner("Generating Analysis sheet..."):
            analysis_buffer = generate_analysis_sheet_buffer(df, submitted_count)
            st.download_button(
                label="📥 Download Updated Excel with Analysis Sheet",
                data=analysis_buffer,
                file_name="submitted_data_analysis.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary",
                use_container_width=True
            )
