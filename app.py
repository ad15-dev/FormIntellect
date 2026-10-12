"""
FormIntellect — Google Form submission scheduler.

Streamlit front-end for the Google Form Auto-Submitter script.
The engine below is a faithful port of the original script (same inputs,
same validation, same scheduling rules, same output format). The only
changes are mechanical: print() -> log(), input() -> widgets,
sys.exit() -> AppError, and submission runs in a background thread so the
page stays live.

Run:  streamlit run app.py
"""
from __future__ import annotations

import html
import io
import json
import os
import random
import re
import tempfile
import threading
import time
import urllib.parse
from collections import Counter
from dataclasses import dataclass
from datetime import datetime, timedelta

import pandas as pd
import pytz
import requests
import streamlit as st

try:
    from openpyxl import load_workbook
    from openpyxl.chart import BarChart
    from openpyxl.chart import Reference
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    from openpyxl.utils import get_column_letter
    OPENPYXL_AVAILABLE = True
except ImportError:
    OPENPYXL_AVAILABLE = False

IST = pytz.timezone("Asia/Kolkata")
MIN_SECONDS_PER_ROW = 2


class AppError(Exception):
    """Raised where the original script printed an error and called sys.exit()."""


# ============================================================
#  UTILITY
# ============================================================

def format_mm_ss(seconds):
    m, s = divmod(int(seconds), 60)
    return f"{m:02d}:{s:02d}"


def format_human(seconds):
    m, s = divmod(int(seconds), 60)
    if m:
        return f"{m}m {s}s"
    return f"{s}s"


def _try_parse(value, fmt):
    try:
        datetime.strptime(value, fmt)
        return True
    except ValueError:
        return False


def parse_time_input(time_str):
    """12-hour (09:30 AM, 11:00PM) or 24-hour (21:30) -> datetime.time"""
    time_str = time_str.strip()
    for fmt in ("%I:%M %p", "%I:%M%p", "%I %p", "%I%p"):
        try:
            return datetime.strptime(time_str.upper(), fmt).time()
        except ValueError:
            pass
    for fmt in ("%H:%M", "%H:%M:%S"):
        try:
            return datetime.strptime(time_str, fmt).time()
        except ValueError:
            pass
    raise ValueError(
        f"Unrecognized time format: '{time_str}'. "
        "Use 12-hour (e.g. 09:30 AM / 11:00 PM) or 24-hour (e.g. 13:30)."
    )


# ============================================================
#  FILE LOADING / SAVING
# ============================================================

def load_dataframe(file_path, log):
    """Supports .xlsx .xls .csv .tsv .ods .json and Google Sheets URL."""
    lp = file_path.strip()

    if lp.startswith("https://"):
        try:
            resp = requests.get(lp, timeout=30)
            resp.raise_for_status()
        except requests.exceptions.RequestException as exc:
            raise AppError(f"Could not fetch Google Sheet: {exc}")
        df = pd.read_csv(io.StringIO(resp.text))
        log(f"Google Sheet loaded: {len(df)} rows, {len(df.columns)} columns")
        output_path = "responses_output.xlsx"
        log("WARNING: Google Sheets is read-only. Updated file will be saved locally as:")
        log(f"         {output_path}")
        return df, True, output_path

    lower = lp.lower()
    try:
        if lower.endswith(".xlsx") or lower.endswith(".xls"):
            df = pd.read_excel(lp)
        elif lower.endswith(".ods"):
            df = pd.read_excel(lp, engine="odf")
        elif lower.endswith(".tsv"):
            df = pd.read_csv(lp, sep="\t")
        elif lower.endswith(".csv"):
            df = pd.read_csv(lp)
        elif lower.endswith(".json"):
            df = pd.read_json(lp)
        else:
            raise AppError(
                f"Unsupported file format: {lp}\n"
                "Supported: .xlsx  .xls  .csv  .tsv  .ods  .json  Google Sheets URL"
            )
    except FileNotFoundError:
        raise AppError(f"File not found: {lp}")

    return df, False, lp


def save_dataframe(df, file_path):
    lower = file_path.lower()
    if lower.endswith(".xlsx") or lower.endswith(".xls"):
        df.to_excel(file_path, index=False)
    elif lower.endswith(".tsv"):
        df.to_csv(file_path, sep="\t", index=False)
    elif lower.endswith(".json"):
        df.to_json(file_path, orient="records", indent=2)
    elif lower.endswith(".ods"):
        df.to_excel(file_path, index=False, engine="odf")
    else:
        df.to_csv(file_path, index=False)


# ============================================================
#  FORM METADATA
# ============================================================

def extract_entry_ids(prefilled_link):
    parsed = urllib.parse.urlparse(prefilled_link)
    pairs = urllib.parse.parse_qsl(parsed.query, keep_blank_values=True)
    seen, ids = set(), []
    for key, _ in pairs:
        if key.startswith("entry.") and key not in seen:
            ids.append(key)
            seen.add(key)
    return ids


def _parse_fb_blob(html_text):
    token = "FB_PUBLIC_LOAD_DATA_"
    idx = html_text.find(token)
    if idx == -1:
        return None
    start = html_text.find("[", idx)
    if start == -1:
        return None
    depth, in_str, escape = 0, False, False
    for i in range(start, len(html_text)):
        ch = html_text[i]
        if in_str:
            if escape:
                escape = False
            elif ch == "\\":
                escape = True
            elif ch == '"':
                in_str = False
            continue
        if ch == '"':
            in_str = True
        elif ch == "[":
            depth += 1
        elif ch == "]":
            depth -= 1
            if depth == 0:
                try:
                    return json.loads(html_text[start:i + 1])
                except json.JSONDecodeError:
                    return None
    return None


def _options_from_blob(blob):
    result = {}
    try:
        items = blob[1][1]
    except Exception:
        return result
    for item in items:
        if not isinstance(item, list) or len(item) < 5:
            continue
        question = item[4]
        if not isinstance(question, list) or not question:
            continue
        q0 = question[0]
        if not isinstance(q0, list) or len(q0) < 2:
            continue
        eid = q0[0]
        if eid is None:
            continue
        opts = []
        if isinstance(q0[1], list):
            for opt in q0[1]:
                if isinstance(opt, list) and opt and isinstance(opt[0], str):
                    opts.append(opt[0])
        if opts:
            result[f"entry.{eid}"] = opts
    return result


def _grid_from_blob(blob):
    result = {}
    try:
        items = blob[1][1]
    except Exception:
        return result
    for item in items:
        if not isinstance(item, list) or len(item) < 5:
            continue
        q_label = item[1] if len(item) > 1 and isinstance(item[1], str) else ""
        question = item[4]
        if not isinstance(question, list) or len(question) < 2:
            continue
        sub_entries = []
        col_options = []
        row_labels = []
        for q_item in question:
            if not isinstance(q_item, list) or len(q_item) < 2:
                continue
            eid = q_item[0]
            if eid is None:
                continue
            opts = []
            if isinstance(q_item[1], list):
                for opt in q_item[1]:
                    if isinstance(opt, list) and opt and isinstance(opt[0], str):
                        opts.append(opt[0])
            if opts:
                sub_entries.append(f"entry.{eid}")
                if not col_options:
                    col_options = opts
            lbl = q_item[3] if len(q_item) > 3 and isinstance(q_item[3], str) else f"entry.{eid}"
            row_labels.append(lbl)
        if len(sub_entries) >= 2:
            if len(row_labels) != len(sub_entries):
                row_labels = list(sub_entries)
            for i, eid in enumerate(sub_entries):
                result[eid] = {
                    "question_label": q_label,
                    "sub_label": row_labels[i],
                    "all_sub_entries": sub_entries,
                    "options": col_options,
                }
    return result


def fetch_form_metadata(prefilled_link):
    try:
        resp = requests.get(prefilled_link, timeout=30)
    except requests.exceptions.RequestException:
        return {}, {}, {}, []

    page = resp.text
    field_types = {}

    for m in re.finditer(r'name="(entry\.\d+)"[^>]*type="([^"]+)"', page):
        eid, itype = m.group(1), m.group(2).lower()
        if eid not in field_types:
            field_types[eid] = itype
        elif "checkbox" in (field_types[eid], itype):
            field_types[eid] = "checkbox"

    for m in re.finditer(r'<textarea[^>]*name="(entry\.\d+)"', page):
        field_types.setdefault(m.group(1), "textarea")
    for m in re.finditer(r'<select[^>]*name="(entry\.\d+)"', page):
        field_types.setdefault(m.group(1), "select")

    options_by_entry, grid_info, questions_order = {}, {}, []

    blob = _parse_fb_blob(page)
    if blob:
        options_by_entry = _options_from_blob(blob)
        grid_info = _grid_from_blob(blob)
        for eid in grid_info:
            field_types[eid] = "grid"
        try:
            for item in blob[1][1]:
                if not isinstance(item, list) or len(item) < 5:
                    continue
                label = item[1] if len(item) > 1 and isinstance(item[1], str) else ""
                question = item[4]
                if not isinstance(question, list):
                    continue
                for q_item in question:
                    if not isinstance(q_item, list) or not q_item:
                        continue
                    eid_raw = q_item[0]
                    if eid_raw is None:
                        continue
                    eid = f"entry.{eid_raw}"
                    ftype = field_types.get(eid, "text")
                    questions_order.append((eid, label, ftype))
        except Exception:
            pass

    if not questions_order:
        for eid in extract_entry_ids(prefilled_link):
            ftype = field_types.get(eid, "text")
            questions_order.append((eid, eid, ftype))

    return field_types, options_by_entry, grid_info, questions_order


# ============================================================
#  VALUE NORMALIZATION
# ============================================================

def normalize_value(raw, entry_type="text"):
    if pd.isna(raw):
        return "NA"
    if isinstance(raw, (pd.Timestamp, datetime)):
        if entry_type == "date":
            return raw.strftime("%Y-%m-%d")
        if entry_type == "time":
            return raw.strftime("%H:%M")
        if entry_type == "datetime-local":
            return raw.strftime("%Y-%m-%dT%H:%M")
    text = str(raw).strip()
    if not text:
        return "NA"
    if text.lower() in {"na", "n/a", "none", "null"}:
        return text
    if entry_type == "datetime-local" and " " in text and "T" not in text:
        text = text.replace(" ", "T", 1)
    return text


def _case_match(value, options):
    lower_map = {o.lower(): o for o in options if isinstance(o, str)}
    return lower_map.get(value.lower())


def best_option(value, options):
    if not options:
        return value
    if value in options:
        return value
    matched = _case_match(value, options)
    if matched:
        return matched
    try:
        num = float(value)
        candidates = []
        for o in options:
            try:
                candidates.append((abs(float(o) - num), o))
            except (ValueError, TypeError):
                pass
        if candidates:
            return sorted(candidates)[0][1]
    except ValueError:
        pass
    return options[0]


# ============================================================
#  AUTO-CORRECT & VALIDATE
# ============================================================

def auto_correct(df, pending_indices, required_columns, field_types, options_by_entry):
    corrections, errors = [], []
    for idx in pending_indices:
        for col, entry in required_columns.items():
            etype = field_types.get(entry, "text")
            value = normalize_value(df.at[idx, col], etype)
            new_val = value
            reason = None

            opts = options_by_entry.get(entry)
            if opts:
                if etype == "checkbox" or "," in value:
                    parts = [v.strip() for v in value.split(",") if v.strip()]
                    fixed = [best_option(v, opts) for v in parts]
                    candidate = ", ".join(dict.fromkeys(fixed))
                    if candidate != value:
                        new_val = candidate
                        reason = "normalized checkbox options"
                else:
                    candidate = best_option(value, opts)
                    if candidate != value:
                        new_val = candidate
                        reason = "normalized to allowed option"

            if etype == "date":
                try:
                    new_val = pd.to_datetime(value).strftime("%Y-%m-%d")
                    if new_val != value:
                        reason = "normalized date"
                except Exception:
                    errors.append((idx, col, entry, value, "invalid date"))
                    continue
            elif etype == "time":
                try:
                    new_val = pd.to_datetime(value).strftime("%H:%M")
                    if new_val != value:
                        reason = "normalized time"
                except Exception:
                    errors.append((idx, col, entry, value, "invalid time"))
                    continue
            elif etype == "datetime-local":
                try:
                    new_val = pd.to_datetime(value).strftime("%Y-%m-%dT%H:%M")
                    if new_val != value:
                        reason = "normalized datetime"
                except Exception:
                    errors.append((idx, col, entry, value, "invalid datetime"))
                    continue

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

            if etype in {"radio", "select"} and "," in value:
                errors.append((idx, col, entry, value, "multiple values for single-choice"))
                continue

            opts = options_by_entry.get(entry)
            if opts:
                parts = (
                    [v.strip() for v in value.split(",") if v.strip()]
                    if (etype == "checkbox" or "," in value)
                    else [value]
                )
                bad = [v for v in parts if v not in opts]
                if bad:
                    errors.append((idx, col, entry, value, f"not in options: {', '.join(bad)}"))
                    continue

            if etype == "date":
                if not _try_parse(value, "%Y-%m-%d"):
                    errors.append((idx, col, entry, value, "invalid date (YYYY-MM-DD)"))
            elif etype == "time":
                if not any(_try_parse(value, f) for f in ("%H:%M", "%H:%M:%S")):
                    errors.append((idx, col, entry, value, "invalid time (HH:MM or HH:MM:SS)"))
            elif etype == "datetime-local":
                if not any(_try_parse(value, f) for f in ("%Y-%m-%dT%H:%M", "%Y-%m-%dT%H:%M:%S")):
                    errors.append((idx, col, entry, value, "invalid datetime (YYYY-MM-DDTHH:MM)"))

    if errors:
        out = ["DATA VALIDATION FAILED"]
        for idx, col, entry, value, reason in errors:
            out.append(
                f"Row {idx + 1} | Column: {col} | Field: {entry} | "
                f"Value: {value} | Reason: {reason}"
            )
            opts = options_by_entry.get(entry)
            if opts:
                out.append(f"  Options: {opts}")
        out.append("")
        out.append("Fix the above cells and re-run.")
        raise AppError("\n".join(out))


# ============================================================
#  TIME WINDOW SCHEDULER
# ============================================================

def compute_time_window(start_input, end_input, start_rand_min, start_rand_max,
                        end_rand_min, end_rand_max, total_rows, log):
    """
    Same rules as the original configure_time_window():
      1. Final window  >  original window (strictly greater)
      2. Final window  >= total_rows x 2 seconds
    min/max pairs are auto-swapped with a warning. Up to 1000 attempts.
    (The original re-prompted on failure; here the user edits the inputs.)
    """
    min_required_sec = total_rows * MIN_SECONDS_PER_ROW
    log(f"\nMinimum required window: {format_human(min_required_sec)}")
    log("Time formats accepted: 12-hour (e.g. 09:30 AM, 11:00 PM) or 24-hour (e.g. 21:30)")

    today = datetime.now(IST).date()
    try:
        start_time = parse_time_input(start_input)
        end_time = parse_time_input(end_input)
        start_dt = IST.localize(datetime.combine(today, start_time))
        end_dt = IST.localize(datetime.combine(today, end_time))
    except ValueError as e:
        raise AppError(str(e))

    if end_dt <= start_dt:
        end_dt += timedelta(days=1)

    original_window_sec = int((end_dt - start_dt).total_seconds())

    if start_rand_min > start_rand_max:
        log("WARNING: Random START seconds: min is greater than max. "
            "Swapping the values and continuing.")
        start_rand_min, start_rand_max = start_rand_max, start_rand_min

    if end_rand_min > end_rand_max:
        log("WARNING: Random END seconds: min is greater than max. "
            "Swapping the values and continuing.")
        end_rand_min, end_rand_max = end_rand_max, end_rand_min

    rs = re_ = window_sec = 0
    rand_start = rand_end = None
    found = False

    for _ in range(1000):
        rs = random.randint(start_rand_min, start_rand_max)
        re_ = random.randint(end_rand_min, end_rand_max)
        rand_start = start_dt + timedelta(seconds=rs)
        rand_end = end_dt + timedelta(seconds=re_)
        if rand_end <= rand_start:
            rand_end = rand_start + timedelta(seconds=1)
        window_sec = int((rand_end - rand_start).total_seconds())
        if window_sec > original_window_sec and window_sec >= min_required_sec:
            found = True
            break

    if not found:
        raise AppError(
            "Unable to find a valid randomization after 1000 attempts where:\n"
            "   - Final Time Window > Original Time Window (strictly greater), and\n"
            f"  - Final Time Window >= {format_human(min_required_sec)} "
            f"(minimum for {total_rows} rows).\n"
            "\nRule explanation:\n"
            "   Random offsets are added to your start and end times independently.\n"
            "   The resulting window MUST be strictly longer than the base window\n"
            "   you entered. Equal length is not accepted. This guarantees that\n"
            "   randomization only ever expands the submission window, never shrinks it.\n"
            "\nHow to fix:\n"
            "   - Increase end_rand_max so the end time is pushed further out.\n"
            "   - Reduce start_rand_max so the start time is not pushed too far in.\n"
            "   - Ensure end_rand_max > start_rand_max for reliable window expansion.\n"
            "   - Widen your base start/end time range (e.g. 09:00 AM to 11:00 AM)."
        )

    log("\n========== TIME CONFIGURATION ==========")
    log(f"Entered START Time     : {start_dt.strftime('%d-%m-%Y %I:%M:%S %p IST')}")
    log(f"Entered END Time       : {end_dt.strftime('%d-%m-%Y %I:%M:%S %p IST')}")
    log(f"\nRandom START Seconds   : {rs}")
    log(f"Random END Seconds     : {re_}")
    log(f"\nRandomized START Time  : {rand_start.strftime('%d-%m-%Y %I:%M:%S %p IST')}")
    log(f"Randomized END Time    : {rand_end.strftime('%d-%m-%Y %I:%M:%S %p IST')}")
    log(f"\nOriginal Time Window   : {format_mm_ss(original_window_sec)} (MM:SS)")
    log(f"Final Time Window      : {format_mm_ss(window_sec)} (MM:SS)")
    log(f"\nTotal Rows to Submit   : {total_rows}")
    log("========================================")

    return rand_start, rand_end, window_sec


def wait_until(target_dt, stop):
    """Sleep until target; returns False if stopped early."""
    while True:
        remaining = (target_dt - datetime.now(IST)).total_seconds()
        if remaining <= 0:
            return True
        if stop.is_set():
            return False
        slp = max(0.1, remaining - random.uniform(1.5, 4.5))
        stop.wait(min(slp, remaining, 1.0))


# ============================================================
#  SUBMISSION
# ============================================================

def submit_row(row, required_columns, field_types, form_url, prefilled_link, max_retries=3):
    payload = {}
    for col, entry in required_columns.items():
        etype = field_types.get(entry, "text")
        value = normalize_value(row[col], etype)
        if "," in value:
            payload[entry] = [v.strip() for v in value.split(",")]
        else:
            payload[entry] = value

    session = requests.Session()
    headers = {
        "User-Agent": "Mozilla/5.0",
        "Referer": prefilled_link,
        "Origin": "https://docs.google.com",
    }
    for _ in range(max_retries):
        try:
            r = session.post(form_url, data=payload, headers=headers, timeout=30)
            if r.status_code == 200 and (
                "Your response has been recorded" in r.text
                or "formResponse" in r.url
            ):
                return 200, "Success"
            return r.status_code, f"HTTP {r.status_code}: {r.text[:200]}"
        except requests.exceptions.RequestException:
            pass
        time.sleep(2)
    return 0, "Failed after max retries"


# ============================================================
#  ANALYSIS SHEET
# ============================================================

def generate_analysis_sheet(df, file_path, submitted_count, log):
    if not OPENPYXL_AVAILABLE:
        log("WARNING: openpyxl not installed. Analysis sheet skipped.")
        return
    lower = file_path.lower()
    if not (lower.endswith(".xlsx") or lower.endswith(".xls")):
        log("WARNING: Analysis sheet is only supported for .xlsx / .xls files.")
        return

    wb = load_workbook(file_path)
    if "Analysis" in wb.sheetnames:
        del wb["Analysis"]
    ws = wb.create_sheet("Analysis")

    NAVY, BLUE, LIGHT_BLUE, WHITE = "1F3864", "2E75B6", "DEEAF1", "FFFFFF"

    def solid(color):
        return PatternFill("solid", fgColor=color)

    thin = Side(style="thin", color="CCCCCC")
    bdr = Border(left=thin, right=thin, top=thin, bottom=thin)
    ctr = Alignment(horizontal="center", vertical="center", wrap_text=True)

    ws.merge_cells("A1:F1")
    ws["A1"] = "Response Analysis"
    ws["A1"].fill = solid(NAVY)
    ws["A1"].font = Font(bold=True, color=WHITE, size=14)
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
        if series.empty:
            continue

        is_checkbox = series.str.contains(",").any()

        ws.merge_cells(start_row=current_row, start_column=1, end_row=current_row, end_column=5)
        cell = ws.cell(row=current_row, column=1, value=col_name)
        cell.fill = solid(BLUE)
        cell.font = Font(bold=True, color=WHITE, size=11)
        cell.alignment = ctr
        ws.row_dimensions[current_row].height = 20
        current_row += 1

        hdr_row = current_row
        for ci, h in enumerate(["#", "Answer", "Count", "% of Total"], 1):
            c = ws.cell(row=current_row, column=ci, value=h)
            c.fill = solid(BLUE)
            c.font = Font(bold=True, color=WHITE, size=10)
            c.alignment = ctr
            c.border = bdr
        current_row += 1

        if is_checkbox:
            all_vals = []
            for v in series:
                all_vals.extend(x.strip() for x in v.split(",") if x.strip())
            freq = Counter(all_vals)
            total_responders = len(series)
        else:
            freq = Counter(series.tolist())
            total_responders = len(series)

        items_sorted = sorted(freq.items(), key=lambda x: -x[1])
        total_count = sum(freq.values())
        data_start = current_row

        for ri, (answer, count) in enumerate(items_sorted, 1):
            pct = (count / total_responders * 100) if total_responders else 0
            for ci, val in enumerate([ri, answer, count, f"{pct:.1f}%"], 1):
                c = ws.cell(row=current_row, column=ci, value=val)
                c.fill = solid(LIGHT_BLUE if ri % 2 == 0 else WHITE)
                c.font = Font(color="1F1F1F", size=10)
                c.alignment = ctr
                c.border = bdr
            current_row += 1

        total_row_idx = current_row
        total_pct = "100%" if not is_checkbox else "---"
        for ci, val in enumerate(["", "TOTAL", total_count, total_pct], 1):
            c = ws.cell(row=current_row, column=ci, value=val)
            c.fill = solid(NAVY)
            c.font = Font(bold=True, color=WHITE, size=10)
            c.alignment = ctr
            c.border = bdr
        current_row += 1

        if items_sorted:
            chart = BarChart()
            chart.type = "col"
            chart.grouping = "clustered"
            chart.title = str(col_name)[:30]
            chart.y_axis.title = "Count"
            chart.x_axis.title = "Answer"
            chart.style = 10
            chart.width = 18
            chart.height = 12
            data_ref = Reference(ws, min_col=3, max_col=3, min_row=hdr_row, max_row=total_row_idx - 1)
            cats_ref = Reference(ws, min_col=2, max_col=2, min_row=data_start, max_row=total_row_idx - 1)
            chart.add_data(data_ref, titles_from_data=True)
            chart.set_categories(cats_ref)
            ws.add_chart(chart, f"{get_column_letter(7)}{data_start}")

        current_row += 2

    wb.save(file_path)
    log(f"Analysis sheet added/updated in: {file_path}")


# ============================================================
#  COLUMN MAPPING
# ============================================================

def map_columns(df, prefilled_link, status_col):
    entry_cols = [c for c in df.columns if c.startswith("entry.")]
    prefilled_ids = extract_entry_ids(prefilled_link)
    non_status_cols = [c for c in df.columns if c != status_col]

    if entry_cols:
        return {c: c for c in entry_cols}

    if prefilled_ids:
        matched = [c for c in prefilled_ids if c in df.columns]
        if matched:
            return {c: c for c in matched}
        if len(prefilled_ids) == len(non_status_cols):
            return dict(zip(non_status_cols, prefilled_ids))

    raise AppError(
        "Unable to map columns to form entry IDs.\n"
        "Tip:   Rename columns to entry.<id>, or ensure column count matches the prefilled link."
    )


# ============================================================
#  PREPARE  (everything the script did before the submit loop)
# ============================================================

@dataclass
class Prepared:
    df: pd.DataFrame
    output_path: str
    status_col: str
    required_columns: dict
    field_types: dict
    pending_indices: list
    form_url: str
    prefilled_link: str
    rand_start: datetime
    rand_end: datetime
    window_sec: int
    pre_lines: list


def prepare(inp: dict, workdir: str) -> Prepared:
    lines: list = []
    log = lines.append

    def section(title):
        log(f"\n{title}")

    if inp["upload"]:
        name, data = inp["upload"]
        file_path = os.path.join(workdir, os.path.basename(name))
        with open(file_path, "wb") as fh:
            fh.write(data)
    else:
        file_path = inp["file_path"].strip().strip('"').strip("'")
        if not file_path:
            raise AppError("Enter a file path / Google Sheets URL, or upload a file.")

    prefilled_link = inp["link"].strip()
    if not prefilled_link or "/viewform" not in prefilled_link:
        raise AppError("Valid pre-filled link is required.")

    form_url = prefilled_link.split("/viewform")[0] + "/formResponse"

    df, is_google_sheets, output_path = load_dataframe(file_path, log)
    if is_google_sheets:
        output_path = os.path.join(workdir, "responses_output.xlsx")
    df.columns = df.columns.str.strip()

    status_col = None
    for col in df.columns:
        if col.strip().lower() in {"status", "submit", "submitted"}:
            status_col = col
            break
    if status_col is None:
        status_col = "Status"
        df[status_col] = ""

    required_columns = map_columns(df, prefilled_link, status_col)

    status_series = df[status_col].astype(str).str.lower().fillna("")
    pending_indices = df[status_series != "submitted"].index.tolist()

    if inp["use_random"].lower() in {"y", "yes"}:
        random.shuffle(pending_indices)
        subset_input = inp["subset"].strip()
        if subset_input:
            try:
                n = max(1, min(int(subset_input), len(pending_indices)))
                pending_indices = random.sample(pending_indices, n)
            except ValueError:
                raise AppError("Invalid number.")

    total_rows = len(pending_indices)
    if total_rows == 0:
        raise AppError("No pending rows to submit.")

    section("PREFILLED LINK CHECK:")
    log(prefilled_link)

    section("FIELD MAPPING:")
    for col, entry in required_columns.items():
        log(f"- {col} -> {entry}")

    field_types, options_by_entry, _grid_info, _ = fetch_form_metadata(prefilled_link)
    if field_types:
        section("FIELD TYPES (detected):")
        for entry in required_columns.values():
            log(f"- {entry}: {field_types.get(entry, 'unknown')}")
    else:
        section("WARNING: Could not detect form field types (continuing with data checks).")

    corrections, corr_errors = auto_correct(
        df, pending_indices, required_columns, field_types, options_by_entry
    )
    if corrections:
        section("AUTO-CORRECTIONS APPLIED:")
        for idx, col, entry, old_val, new_val, reason in corrections:
            log(f"Row {idx + 1} | Column: {col} | Field: {entry} | "
                f"{old_val} -> {new_val} | Reason: {reason}")
        save_dataframe(df, output_path)

    if corr_errors:
        out = ["UNABLE TO AUTO-CORRECT:"]
        for idx, col, entry, value, reason in corr_errors:
            out.append(f"Row {idx + 1} | Column: {col} | Field: {entry} | "
                       f"Value: {value} | Reason: {reason}")
        out.append("")
        out.append("Fix the above cells and re-run.")
        raise AppError("\n".join(out))

    validate_data(df, pending_indices, required_columns, field_types, options_by_entry)

    rand_start, rand_end, window_sec = compute_time_window(
        inp["start"], inp["end"],
        inp["srmin"], inp["srmax"], inp["ermin"], inp["ermax"],
        total_rows, log,
    )

    return Prepared(
        df=df, output_path=output_path, status_col=status_col,
        required_columns=required_columns, field_types=field_types,
        pending_indices=pending_indices, form_url=form_url,
        prefilled_link=prefilled_link, rand_start=rand_start,
        rand_end=rand_end, window_sec=window_sec, pre_lines=lines,
    )


# ============================================================
#  BACKGROUND JOB
# ============================================================

class Job:
    """Runs the submit loop in a thread; the UI only reads snapshots."""

    def __init__(self, prep: Prepared):
        self.prep = prep
        self.total = len(prep.pending_indices)
        self.lock = threading.Lock()
        self.stop = threading.Event()
        self.status = "waiting"          # waiting | running | done | stopped | error
        self.rows: list = []             # execution table rows
        self.submitted = 0
        self.failed = 0
        self.next_target = None
        self.error = None
        self.notified = False
        self.thread = threading.Thread(target=self._run, daemon=True)

    # -- helpers -------------------------------------------------
    def start(self):
        self.thread.start()

    @property
    def active(self):
        return self.status in ("waiting", "running")

    def _add(self, **row):
        with self.lock:
            self.rows.append(row)

    def _note(self, text):
        self._add(kind="note", text=text)

    def snapshot(self):
        with self.lock:
            return {
                "status": self.status, "rows": list(self.rows),
                "submitted": self.submitted, "failed": self.failed,
                "total": self.total, "next_target": self.next_target,
                "error": self.error,
            }

    # -- main loop (mirrors the original run_mode1 submit loop) ---
    def _run(self):
        p = self.prep
        try:
            now = datetime.now(IST)
            if now < p.rand_start:
                self.next_target = p.rand_start
                self._note(f"Waiting until START time ({p.rand_start.strftime('%I:%M:%S %p IST')})")
                if not wait_until(p.rand_start, self.stop):
                    self._note("PROCESS STOPPED")
                    self.status = "stopped"
                    return

            self.status = "running"
            total_rows = self.total
            window_sec = p.window_sec
            if window_sec <= total_rows:
                window_sec = max(total_rows + 5, window_sec)

            offsets = sorted(random.sample(range(1, window_sec + 1), total_rows))
            start_ref = datetime.now(IST)

            for i, (offset, idx) in enumerate(zip(offsets, p.pending_indices), 1):
                if self.stop.is_set():
                    break
                target = start_ref + timedelta(seconds=offset)
                self.next_target = target
                if not wait_until(target, self.stop):
                    break

                status_code, reason = submit_row(
                    p.df.loc[idx], p.required_columns, p.field_types,
                    p.form_url, p.prefilled_link,
                )
                ist_now = datetime.now(IST).strftime("%d-%m-%Y %I:%M:%S %p")

                if status_code == 200:
                    p.df.at[idx, p.status_col] = "submitted"
                    save_dataframe(p.df, p.output_path)
                    with self.lock:
                        self.submitted += 1

                    if i < total_rows:
                        gap_sec = offsets[i] - offsets[i - 1]
                        next_time = (start_ref + timedelta(seconds=offsets[i])).strftime(
                            "%d-%m-%Y %I:%M:%S %p IST")
                        self._add(
                            kind="ok", row=idx + 1, time=ist_now, http=status_code,
                            nxt=next_time, gap=format_human(gap_sec),
                            text=(f"Row {idx + 1} submitted at {ist_now} IST | "
                                  f"HTTP {status_code} | NEXT ENTRY: {next_time} | "
                                  f"GAP: {format_human(gap_sec)}"),
                        )
                    else:
                        self._add(
                            kind="final", row=idx + 1, time=ist_now, http=status_code,
                            text=(f"Row {idx + 1} submitted at {ist_now} IST | "
                                  f"HTTP {status_code} | FINAL ENTRY"),
                        )
                else:
                    with self.lock:
                        self.failed += 1
                    self._add(kind="fail", row=idx + 1, time=ist_now, reason=reason,
                              text=f"Row {idx + 1} FAILED at {ist_now} IST | {reason}")
                    self._add(kind="detail", text=f"Row data: {p.df.loc[idx].to_dict()}")

            self.next_target = None
            if self.stop.is_set():
                self._note("PROCESS STOPPED")
                self.status = "stopped"
                return

            self._note("PROCESS COMPLETED")
            self._note("Generating Analysis sheet...")
            generate_analysis_sheet(p.df, p.output_path, self.submitted, self._note)
            self.status = "done"
        except Exception as exc:  # noqa: BLE001 — surface anything to the UI
            self.error = f"{type(exc).__name__}: {exc}"
            self.status = "error"
            self._note(f"ERROR: {self.error}")

    def log_text(self):
        lines = list(self.prep.pre_lines) + [""]
        with self.lock:
            lines += [r["text"] for r in self.rows]
        return "\n".join(lines)


# ============================================================
#  UI
# ============================================================

CSS = """
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&family=JetBrains+Mono:wght@400;500&display=swap');
:root{
  --bg:#0F0F0F; --surface:#1C1C1C; --surface2:#161616; --border:#2E2E2E;
  --text:#EDEDED; --muted:#A0A0A0; --green:#3ECF8E; --green-h:#2EB87A;
  --red:#F04438; --amber:#F5A524;
}
html, body, .stApp, [class*="st-"]{font-family:'Inter',system-ui,sans-serif;}
.stApp{background:var(--bg); color:var(--text);}
header[data-testid="stHeader"]{background:transparent;}
#MainMenu, footer{visibility:hidden;}
.block-container{max-width:940px; padding-top:2.4rem; padding-bottom:4rem;}

/* hero */
.fi-hero{display:flex; align-items:center; gap:14px; margin-bottom:6px;}
.fi-logo{width:38px; height:38px; border-radius:10px; background:var(--green);
  display:flex; align-items:center; justify-content:center;}
.fi-name{font-size:1.7rem; font-weight:700; letter-spacing:-0.02em; line-height:1.1;}
.fi-name span{color:var(--green);}
.fi-tag{color:var(--muted); font-size:.92rem; margin:2px 0 26px 52px;}

/* cards */
div[data-testid="stVerticalBlockBorderWrapper"]{
  background:var(--surface); border:1px solid var(--border)!important; border-radius:14px;}
.fi-step{display:flex; align-items:center; gap:10px; font-weight:600; font-size:.95rem; margin-bottom:6px;}
.fi-step b{font-family:'JetBrains Mono',monospace; font-size:.72rem; font-weight:500;
  color:var(--green); border:1px solid rgba(62,207,142,.35); border-radius:6px; padding:2px 7px;}

/* inputs */
[data-testid="stWidgetLabel"] p{color:var(--muted)!important; font-size:.8rem!important; font-weight:500;}
div[data-baseweb="input"], div[data-baseweb="base-input"], div[data-baseweb="select"] > div{
  background:var(--surface2)!important; border-color:var(--border)!important; border-radius:10px!important;}
div[data-baseweb="input"]:focus-within{border-color:var(--green)!important;}
input, textarea{color:var(--text)!important;}
[data-testid="stFileUploaderDropzone"]{background:var(--surface2); border:1px dashed var(--border); border-radius:10px;}

/* buttons */
.stButton > button, .stDownloadButton > button{
  border-radius:10px; border:1px solid var(--border); background:var(--surface2);
  color:var(--text); font-weight:600; padding:.55rem 1.1rem; transition:all .15s ease;}
.stButton > button:hover, .stDownloadButton > button:hover{border-color:var(--green); color:var(--green);}
.stButton > button[kind="primary"], button[data-testid="stBaseButton-primary"]{
  background:var(--green)!important; border-color:var(--green)!important; color:#0F0F0F!important;}
.stButton > button[kind="primary"]:hover, button[data-testid="stBaseButton-primary"]:hover{
  background:var(--green-h)!important; border-color:var(--green-h)!important; color:#0F0F0F!important;}

/* stats */
.fi-stats{display:grid; grid-template-columns:repeat(4,1fr); gap:12px; margin:6px 0 14px;}
.fi-stat{background:var(--surface); border:1px solid var(--border); border-radius:12px; padding:14px 16px;}
.fi-stat-l{color:var(--muted); font-size:.72rem; text-transform:uppercase; letter-spacing:.08em;}
.fi-stat-v{font-size:1.45rem; font-weight:600; margin-top:4px; font-family:'JetBrains Mono',monospace;}
.fi-stat-s{color:var(--muted); font-size:.74rem; margin-top:2px; min-height:1em;}
.fi-stat-v.g{color:var(--green);} .fi-stat-v.r{color:var(--red);} .fi-stat-v.a{color:var(--amber);}
.fi-bar{height:6px; background:var(--border); border-radius:99px; overflow:hidden; margin-bottom:18px;}
.fi-bar > div{height:100%; background:var(--green); border-radius:99px; transition:width .4s ease;}

/* execution table */
.fi-tbl-h{display:flex; justify-content:space-between; align-items:center; margin:4px 0 8px;
  font-size:.78rem; text-transform:uppercase; letter-spacing:.08em; color:var(--muted);}
.fi-tbl{display:flex; flex-direction:column-reverse; max-height:440px; overflow-y:auto;
  background:var(--surface2); border:1px solid var(--border); border-radius:12px;}
.fi-row{font-family:'JetBrains Mono',monospace; font-size:.78rem; line-height:1.5;
  padding:9px 14px 9px 30px; border-top:1px solid var(--border); position:relative;
  color:var(--text); word-break:break-word;}
.fi-row:last-child{border-top:none;}
.fi-row::before{content:''; position:absolute; left:13px; top:15px; width:7px; height:7px;
  border-radius:50%; background:var(--green);}
.fi-row.final::before{box-shadow:0 0 0 3px rgba(62,207,142,.2);}
.fi-row.fail::before{background:var(--red);}
.fi-row.note::before, .fi-row.detail::before{background:var(--muted);}
.fi-row.note, .fi-row.detail, .fi-row.empty{color:var(--muted);}
.fi-row.fail{color:#ffb4ae;}
.fi-row .r{font-weight:500;}
.fi-row .k{color:var(--muted);}
.fi-row .sep{color:#555;}
.fi-row .pill{color:var(--green); background:rgba(62,207,142,.12); border-radius:5px; padding:1px 6px;}

@media (max-width:720px){ .fi-stats{grid-template-columns:repeat(2,1fr);} }
</style>
"""

STATUS_LABEL = {
    "waiting": ("Waiting", "a", "for start time"),
    "running": ("Running", "g", "submitting"),
    "done": ("Completed", "g", "all rows processed"),
    "stopped": ("Stopped", "a", "stopped by user"),
    "error": ("Error", "r", "see table"),
}


def esc(v):
    return html.escape(str(v))


def stat(label, value, sub="", tone=""):
    return (f'<div class="fi-stat"><div class="fi-stat-l">{label}</div>'
            f'<div class="fi-stat-v {tone}">{value}</div><div class="fi-stat-s">{sub}</div></div>')


def row_html(r):
    k = r["kind"]
    if k in ("ok", "final"):
        s = (f'<span class="r">Row {r["row"]}</span> submitted at {esc(r["time"])} IST '
             f'<span class="sep">|</span> <span class="pill">HTTP {r["http"]}</span> '
             f'<span class="sep">|</span> ')
        if k == "ok":
            s += (f'<span class="k">NEXT ENTRY:</span> {esc(r["nxt"])} '
                  f'<span class="sep">|</span> <span class="k">GAP:</span> {esc(r["gap"])}')
        else:
            s += '<span class="k">FINAL ENTRY</span>'
        return f'<div class="fi-row {k}">{s}</div>'
    if k == "fail":
        return (f'<div class="fi-row fail"><span class="r">Row {r["row"]}</span> FAILED at '
                f'{esc(r["time"])} IST <span class="sep">|</span> {esc(r["reason"])}</div>')
    return f'<div class="fi-row {k}">{esc(r["text"])}</div>'


def render_table(rows):
    if rows:
        body = "".join(row_html(r) for r in reversed(rows))   # column-reverse keeps newest in view
    else:
        body = '<div class="fi-row empty">No submissions yet…</div>'
    return (f'<div class="fi-tbl-h"><span>Execution table</span><span>{len(rows)} entries</span></div>'
            f'<div class="fi-tbl">{body}</div>')


def render_monitor(job: Job):
    s = job.snapshot()
    label, tone, sub = STATUS_LABEL[s["status"]]
    done_n = s["submitted"] + s["failed"]
    pct = int(done_n / s["total"] * 100) if s["total"] else 0

    if s["next_target"] is not None and s["status"] in ("waiting", "running"):
        secs = max(0, int((s["next_target"] - datetime.now(IST)).total_seconds()))
        nxt_v, nxt_s = format_mm_ss(secs), s["next_target"].strftime("%I:%M:%S %p IST")
    else:
        nxt_v, nxt_s = "—", ""

    st.markdown(
        '<div class="fi-stats">'
        + stat("Status", label, sub, tone)
        + stat("Submitted", f'{s["submitted"]}<span style="color:#666"> / {s["total"]}</span>', f"{pct}% processed", "g")
        + stat("Failed", s["failed"], "", "r" if s["failed"] else "")
        + stat("Next entry in", nxt_v, nxt_s)
        + f'</div><div class="fi-bar"><div style="width:{pct}%"></div></div>',
        unsafe_allow_html=True,
    )
    st.markdown(render_table(s["rows"]), unsafe_allow_html=True)

    if job.active:
        if st.button("Stop", key="stop_btn"):
            job.stop.set()
    elif not job.notified:
        job.notified = True
        st.rerun()                      # one full rerun to swap in the final layout


def main():
    st.set_page_config(page_title="FormIntellect", page_icon="✅", layout="centered")
    st.markdown(CSS, unsafe_allow_html=True)

    if "workdir" not in st.session_state:
        st.session_state.workdir = tempfile.mkdtemp(prefix="formintellect_")
    job: Job | None = st.session_state.get("job")
    locked = bool(job and job.active)

    st.markdown(
        '<div class="fi-hero"><div class="fi-logo"><svg viewBox="0 0 24 24" width="20" height="20" '
        'fill="none" stroke="#0F0F0F" stroke-width="2.8" stroke-linecap="round" stroke-linejoin="round">'
        '<path d="M20 6 9 17l-5-5"/></svg></div>'
        '<div class="fi-name">Form<span>Intellect</span></div></div>'
        '<div class="fi-tag">Validated, paced Google Form submissions from your spreadsheet.</div>',
        unsafe_allow_html=True,
    )

    # ---------------- inputs (same prompts as the original script) ----------------
    with st.container(border=True):
        st.markdown('<div class="fi-step"><b>01</b> Source</div>', unsafe_allow_html=True)
        file_path = st.text_input(
            "Enter file path (.xlsx/.xls/.csv/.tsv/.ods/.json or Google Sheets URL)",
            key="file_path", disabled=locked)
        upload = st.file_uploader(
            "…or upload the file instead", type=["xlsx", "xls", "csv", "tsv", "ods", "json"],
            key="upload", disabled=locked)
        link = st.text_input("Enter Google Form pre-filled link", key="link", disabled=locked)

    with st.container(border=True):
        st.markdown('<div class="fi-step"><b>02</b> Rows</div>', unsafe_allow_html=True)
        use_random = st.radio("Use random rows? (y/n)", ["y", "n"], index=1, horizontal=True,
                              key="use_random", disabled=locked)
        subset = ""
        if use_random == "y":
            subset = st.text_input("How many random rows to submit? (blank = all)",
                                   key="subset", disabled=locked)

    with st.container(border=True):
        st.markdown('<div class="fi-step"><b>03</b> Schedule</div>', unsafe_allow_html=True)
        c1, c2 = st.columns(2)
        start = c1.text_input("Start time", placeholder="09:30 AM  or  21:30", key="start", disabled=locked)
        end = c2.text_input("End time", placeholder="11:00 AM  or  23:00", key="end", disabled=locked)
        c3, c4, c5, c6 = st.columns(4)
        srmin = c3.number_input("Random delay AFTER start (min seconds)", value=0, step=1, key="srmin", disabled=locked)
        srmax = c4.number_input("Random delay AFTER start (max seconds)", value=0, step=1, key="srmax", disabled=locked)
        ermin = c5.number_input("Random delay AFTER end (min seconds)", value=0, step=1, key="ermin", disabled=locked)
        ermax = c6.number_input("Random delay AFTER end (max seconds)", value=0, step=1, key="ermax", disabled=locked)

    if st.button("Start submission", type="primary", disabled=locked, use_container_width=True):
        if not start.strip() or not end.strip():
            st.error("Start time and End time are required.")
        else:
            inp = dict(
                file_path=file_path, upload=(upload.name, upload.getvalue()) if upload else None,
                link=link, use_random=use_random, subset=subset, start=start, end=end,
                srmin=int(srmin), srmax=int(srmax), ermin=int(ermin), ermax=int(ermax),
            )
            try:
                with st.spinner("Running preflight checks…"):
                    prep = prepare(inp, st.session_state.workdir)
                job = Job(prep)
                job.start()
                st.session_state.job = job
                st.rerun()
            except AppError as exc:
                head, _, rest = str(exc).partition("\n")
                st.error(head)
                if rest.strip():
                    st.code(rest.strip(), language=None)

    # ---------------- monitor ----------------
    if job:
        st.markdown("<div style='height:18px'></div>", unsafe_allow_html=True)
        with st.expander("Preflight checks", expanded=False):
            st.code("\n".join(job.prep.pre_lines).strip(), language=None)

        monitor = st.fragment(run_every=1 if job.active else None)(render_monitor)
        monitor(job)

        if not job.active:
            d1, d2, d3 = st.columns([1, 1, 1])
            out = job.prep.output_path
            if os.path.exists(out):
                with open(out, "rb") as fh:
                    d1.download_button("Download updated file", fh.read(),
                                       file_name=os.path.basename(out), use_container_width=True)
            d2.download_button("Download execution log", job.log_text(),
                               file_name="formintellect_log.txt", use_container_width=True)
            if d3.button("New run", use_container_width=True):
                del st.session_state["job"]
                st.rerun()


main()
