import io
import json
import os
import random
import re
import sqlite3
import time
import urllib.parse
from collections import Counter
from datetime import datetime, timedelta
from typing import List, Dict, Optional

import pandas as pd
import pytz
import requests
import streamlit as st
from openpyxl import load_workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

# ============================================================
# CONFIGURATION
# ============================================================
IST = pytz.timezone("Asia/Kolkata")
MIN_SECONDS_PER_ROW = 2
DB_PATH = "campaigns.db"

st.set_page_config(
    page_title="formIntellect",
    page_icon="⚡",
    layout="centered",
    initial_sidebar_state="collapsed"
)

# ============================================================
# SUPABASE THEME CSS
# ============================================================
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap');

    .stApp {
        background-color: #09090b !important;
        color: #fafafa !important;
        font-family: 'Inter', -apple-system, BlinkMacSystemFont, sans-serif !important;
    }
    main .block-container {
        padding-top: 2.5rem;
        padding-bottom: 3rem;
        max-width: 1000px;
    }

    h1, h2, h3, h4, h5 {
        color: #fafafa !important;
        font-weight: 600 !important;
        letter-spacing: -0.02em;
    }
    h1 { font-size: 2rem; margin-bottom: 0.2rem; }
    h2 { font-size: 1.4rem; margin-top: 2rem; margin-bottom: 1rem; }
    h3 { font-size: 1.1rem; color: #a1a1aa !important; font-weight: 500 !important; }
    p, span, label, .stMarkdown, .stCaption {
        color: #a1a1aa !important;
        font-size: 0.9rem !important;
    }

    .stTextInput > div > div > input,
    .stNumberInput > div > div > input,
    .stSelectbox > div > div > div {
        background-color: #18181b !important;
        border: 1px solid #27272a !important;
        color: #fafafa !important;
        border-radius: 6px !important;
        padding: 10px 14px !important;
        font-size: 14px !important;
        font-family: 'Inter', sans-serif !important;
    }
    .stTextInput > div > div > input:focus,
    .stNumberInput > div > div > input:focus {
        border-color: #3ecf8e !important;
        box-shadow: 0 0 0 2px rgba(62, 207, 142, 0.15) !important;
    }

    .stFileUploader > div {
        background-color: #18181b !important;
        border: 1px dashed #3f3f46 !important;
        border-radius: 6px !important;
        padding: 24px !important;
    }
    .stFileUploader > div:hover {
        border-color: #3ecf8e !important;
    }

    .stButton > button {
        background-color: #18181b !important;
        border: 1px solid #27272a !important;
        color: #fafafa !important;
        border-radius: 6px !important;
        padding: 10px 20px !important;
        font-weight: 500 !important;
        transition: all 0.2s ease;
        width: 100%;
    }
    .stButton > button:hover {
        background-color: #27272a !important;
        border-color: #3ecf8e !important;
        color: #3ecf8e !important;
    }
    
    .stButton > button[kind="primary"] {
        background-color: #3ecf8e !important;
        border: 1px solid #3ecf8e !important;
        color: #09090b !important;
        font-weight: 600 !important;
        box-shadow: 0 4px 12px rgba(62, 207, 142, 0.2);
    }
    .stButton > button[kind="primary"]:hover {
        background-color: #32b67a !important;
        border-color: #32b67a !important;
    }

    .stProgress > div > div > div > div {
        background-color: #3ecf8e !important;
        border-radius: 4px !important;
    }
    .stProgress > div > div {
        background-color: #27272a !important;
        border-radius: 4px !important;
    }

    .stAlert {
        background-color: #18181b !important;
        border: 1px solid #27272a !important;
        border-radius: 6px !important;
    }
    .stAlert-success { border-left: 4px solid #3ecf8e !important; }
    .stAlert-error { border-left: 4px solid #f43f5e !important; }
    .stAlert-warning { border-left: 4px solid #f59e0b !important; }

    hr {
        border-color: #27272a !important;
        margin: 2rem 0 !important;
    }

    #MainMenu, footer, header {
        visibility: hidden;
    }

    .card {
        background-color: #18181b;
        border: 1px solid #27272a;
        border-radius: 8px;
        padding: 24px;
        margin-bottom: 24px;
    }

    .metric-card {
        background-color: #18181b;
        border: 1px solid #27272a;
        border-radius: 8px;
        padding: 20px;
        text-align: center;
        transition: all 0.3s ease;
    }
    .metric-card:hover {
        border-color: #3ecf8e;
        transform: translateY(-2px);
    }
    .metric-value {
        font-size: 2.5rem;
        font-weight: 700;
        color: #3ecf8e;
        line-height: 1;
    }
    .metric-value.total { color: #3b82f6; }
    .metric-value.success { color: #3ecf8e; }
    .metric-value.failed { color: #f43f5e; }
    .metric-value.pending { color: #f59e0b; }
    .metric-label {
        font-size: 0.85rem;
        color: #a1a1aa;
        margin-top: 8px;
        text-transform: uppercase;
        letter-spacing: 0.05em;
        font-weight: 500;
    }

    .progress-container {
        background-color: #18181b;
        border: 1px solid #27272a;
        border-radius: 8px;
        padding: 20px;
        margin-bottom: 24px;
    }
    .progress-header {
        display: flex;
        justify-content: space-between;
        align-items: center;
        margin-bottom: 12px;
    }
    .progress-title {
        font-size: 1rem;
        font-weight: 600;
        color: #fafafa;
    }
    .progress-percentage {
        font-size: 1.2rem;
        font-weight: 700;
        color: #3ecf8e;
    }
</style>
""", unsafe_allow_html=True)

# ============================================================
# DATABASE SETUP
# ============================================================
def init_db():
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    c.execute('''
        CREATE TABLE IF NOT EXISTS campaigns (
            id TEXT PRIMARY KEY,
            name TEXT NOT NULL,
            form_url TEXT NOT NULL,
            source_type TEXT NOT NULL,
            source_file TEXT,
            total_rows INTEGER DEFAULT 0,
            success_count INTEGER DEFAULT 0,
            failed_count INTEGER DEFAULT 0,
            pending_count INTEGER DEFAULT 0,
            status TEXT DEFAULT 'draft',
            start_time TEXT,
            end_time TEXT,
            config_json TEXT,
            created_at TEXT,
            updated_at TEXT
        )
    ''')
    conn.commit()
    conn.close()

def save_campaign(campaign_data):
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    c.execute('''
        INSERT OR REPLACE INTO campaigns 
        (id, name, form_url, source_type, source_file, total_rows, success_count, 
         failed_count, pending_count, status, start_time, end_time, config_json, 
         created_at, updated_at)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    ''', (
        campaign_data['id'],
        campaign_data['name'],
        campaign_data['form_url'],
        campaign_data['source_type'],
        campaign_data.get('source_file'),
        campaign_data.get('total_rows', 0),
        campaign_data.get('success_count', 0),
        campaign_data.get('failed_count', 0),
        campaign_data.get('pending_count', 0),
        campaign_data.get('status', 'draft'),
        campaign_data.get('start_time'),
        campaign_data.get('end_time'),
        json.dumps(campaign_data.get('config', {})),
        campaign_data.get('created_at', datetime.now(IST).isoformat()),
        datetime.now(IST).isoformat()
    ))
    conn.commit()
    conn.close()

def get_all_campaigns():
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    c.execute('SELECT * FROM campaigns ORDER BY created_at DESC')
    rows = c.fetchall()
    conn.close()
    
    campaigns = []
    for row in rows:
        campaigns.append({
            'id': row[0],
            'name': row[1],
            'form_url': row[2],
            'source_type': row[3],
            'source_file': row[4],
            'total_rows': row[5],
            'success_count': row[6],
            'failed_count': row[7],
            'pending_count': row[8],
            'status': row[9],
            'start_time': row[10],
            'end_time': row[11],
            'config': json.loads(row[12]) if row[12] else {},
            'created_at': row[13],
            'updated_at': row[14]
        })
    return campaigns

# ============================================================
# UTILITY FUNCTIONS (From CSE2.ipynb)
# ============================================================
def format_mm_ss(seconds):
    m, s = divmod(int(seconds), 60)
    return f"{m:02d}:{s:02d}"

def format_human(seconds):
    m, s = divmod(int(seconds), 60)
    if m: return f"{m}m {s}s"
    return f"{s}s"

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
            ids.append(key)
            seen.add(key)
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

def fetch_form_metadata(prefilled_link):
    try:
        resp = requests.get(prefilled_link, timeout=30)
        html = resp.text
    except:
        return {}, {}, [], []

    field_types = {}
    for m in re.finditer(r'name="(entry\.\d+)"[^>]*type="([^"]+)"', html):
        eid, itype = m.group(1), m.group(2).lower()
        field_types[eid] = itype

    options_by_entry, grid_info, questions_order = {}, {}, []
    blob = _parse_fb_blob(html)
    if blob:
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
                    
                    if isinstance(q_item[1], list):
                        opts = [opt[0] for opt in q_item[1] if isinstance(opt, list) and opt and isinstance(opt[0], str)]
                        if opts: options_by_entry[eid] = opts
        except: pass

    return field_types, options_by_entry, grid_info, questions_order

def normalize_value(raw, entry_type="text"):
    if pd.isna(raw): return "NA"
    text = str(raw).strip()
    return text if text else "NA"

def best_option(value, options):
    if not options or value in options: return value
    lower_map = {o.lower(): o for o in options if isinstance(o, str)}
    return lower_map.get(value.lower(), options[0])

def validate_and_correct(df, pending_indices, required_columns, field_types, options_by_entry):
    errors = []
    for idx in pending_indices:
        for col, entry in required_columns.items():
            etype = field_types.get(entry, "text")
            value = normalize_value(df.at[idx, col], etype)
            opts = options_by_entry.get(entry)
            
            if opts:
                candidate = best_option(value, opts)
                if candidate != value:
                    df.at[idx, col] = candidate
                    
            if opts and df.at[idx, col] not in opts:
                errors.append(f"Row {idx+1}, Col '{col}': '{df.at[idx, col]}' not in options")
    return errors

def generate_analysis_buffer(df, submitted_count):
    buffer = io.BytesIO()
    df.to_excel(buffer, index=False, engine='openpyxl')
    buffer.seek(0)
    
    wb = load_workbook(buffer)
    if "Analysis" in wb.sheetnames: del wb["Analysis"]
    ws = wb.create_sheet("Analysis")
    
    NAVY, BLUE, LIGHT_BLUE, WHITE = "1F3864", "2E75B6", "DEEAF1", "FFFFFF"
    solid = lambda c: PatternFill("solid", fgColor=c)
    thin = Side(style="thin", color="CCCCCC")
    bdr = Border(left=thin, right=thin, top=thin, bottom=thin)
    ctr = Alignment(horizontal="center", vertical="center", wrap_text=True)
    
    ws["A1"] = "Response Analysis"
    ws["A1"].fill, ws["A1"].font = solid(NAVY), Font(bold=True, color=WHITE, size=14)
    
    current_row = 3
    for col_name in df.columns:
        if col_name.lower() in {"status", "submit"}: continue
        series = df[col_name].dropna().astype(str).str.strip()
        if series.empty: continue
        
        freq = Counter(series.tolist())
        ws.cell(row=current_row, column=1, value=col_name).fill = solid(BLUE)
        current_row += 1
        
        for ri, (ans, cnt) in enumerate(sorted(freq.items(), key=lambda x: -x[1]), 1):
            ws.cell(row=current_row, column=1, value=ans).border = bdr
            ws.cell(row=current_row, column=2, value=cnt).border = bdr
            current_row += 1
        current_row += 1
        
    out_buffer = io.BytesIO()
    wb.save(out_buffer)
    out_buffer.seek(0)
    return out_buffer

# ============================================================
# OFFLINE SCRIPT GENERATOR
# ============================================================
def generate_offline_script(campaign_config, df_data):
    script = f'''
import pandas as pd
import requests
import time
import random
import io
from datetime import datetime, timedelta
import pytz

# --- CAMPAIGN CONFIGURATION ---
IST = pytz.timezone("Asia/Kolkata")
FORM_URL = "{campaign_config['form_url']}"
COLUMN_MAPPING = {json.dumps(campaign_config['column_mapping'])}
FIELD_TYPES = {json.dumps(campaign_config['field_types'])}
OPTIONS_BY_ENTRY = {json.dumps(campaign_config['options_by_entry'])}
START_TIME_STR = "{campaign_config['start_time']}"
END_TIME_STR = "{campaign_config['end_time']}"
DELAY_MIN = {campaign_config['delay_min']}
DELAY_MAX = {campaign_config['delay_max']}
USE_RANDOM = {campaign_config['use_random']}
SUBSET_N = {campaign_config['subset_n']}

# --- EMBEDDED DATA ---
CSV_DATA = """{df_data}"""

def normalize_value(raw, entry_type="text"):
    if pd.isna(raw): return "NA"
    text = str(raw).strip()
    return text if text else "NA"

def best_option(value, options):
    if not options or value in options: return value
    lower_map = {{o.lower(): o for o in options if isinstance(o, str)}}
    return lower_map.get(value.lower(), options[0])

def main():
    print("🚀 Starting Offline Auto-Submitter...")
    print("Campaign: {campaign_config['name']}")
    df = pd.read_csv(io.StringIO(CSV_DATA))
    
    if "Status" not in df.columns:
        df["Status"] = ""
        
    pending_indices = df[df["Status"].str.lower() != "submitted"].index.tolist()
    
    if USE_RANDOM:
        random.shuffle(pending_indices)
        if SUBSET_N > 0:
            pending_indices = pending_indices[:SUBSET_N]
    
    # Time Logic
    now = datetime.now(IST)
    today = now.date()
    
    try:
        sh, sm = map(int, START_TIME_STR.split(":")[:2])
        eh, em = map(int, END_TIME_STR.split(":")[:2])
    except:
        print("❌ Invalid time format. Using immediate start.")
        sh, sm = now.hour, now.minute
        eh, em = (now + timedelta(minutes=5)).hour, (now + timedelta(minutes=5)).minute
    
    start_dt = IST.localize(datetime(today.year, today.month, today.day, sh, sm))
    end_dt = IST.localize(datetime(today.year, today.month, today.day, eh, em))
    
    if end_dt <= start_dt:
        end_dt += timedelta(days=1)
        
    if now < start_dt:
        wait_sec = (start_dt - now).total_seconds()
        print(f"⏳ Waiting until {{start_dt.strftime('%H:%M')}}... ({{int(wait_sec)}} seconds)")
        print("💡 TIP: Configure your OS power settings to prevent sleep during long runs.")
        time.sleep(wait_sec)
        
    print(f"✅ Starting submission window...")
    
    submitted_count = 0
    
    for idx in pending_indices:
        row = df.loc[idx]
        payload = {{}}
        for col, entry in COLUMN_MAPPING.items():
            etype = FIELD_TYPES.get(entry, "text")
            val = normalize_value(row[col], etype)
            opts = OPTIONS_BY_ENTRY.get(entry)
            if opts:
                val = best_option(val, opts)
            payload[entry] = val
            
        try:
            r = requests.post(FORM_URL, data=payload, timeout=30)
            if r.status_code == 200:
                df.at[idx, "Status"] = "submitted"
                submitted_count += 1
                print(f"✅ Row {{idx+1}} Submitted")
            else:
                print(f"❌ Row {{idx+1}} Failed: HTTP {{r.status_code}}")
        except Exception as e:
            print(f"❌ Row {{idx+1}} Error: {{e}}")
            
        time.sleep(random.randint(DELAY_MIN, DELAY_MAX))
        
    print(f"\\n🎉 Process Complete. {{submitted_count}} rows submitted.")
    print("Saving updated CSV...")
    df.to_csv("submitted_results.csv", index=False)
    print("Saved to submitted_results.csv")
    print("\\n💡 You can now close this window. The script has finished.")

if __name__ == "__main__":
    main()
'''
    return script

# ============================================================
# MAIN APPLICATION
# ============================================================
def main():
    init_db()
    
    # Navigation
    if 'page' not in st.session_state:
        st.session_state.page = 'dashboard'
    if 'campaign_config' not in st.session_state:
        st.session_state.campaign_config = {}
    if 'df' not in st.session_state:
        st.session_state.df = None
    
    # Header
    st.markdown("<h1>⚡ formIntellect</h1>", unsafe_allow_html=True)
    st.markdown("<p style='color:#a1a1aa; margin-top:-10px;'>A jugad that passes the authenticity of real interaction.</p>", unsafe_allow_html=True)
    st.divider()
    
    # Navigation Buttons
    col1, col2, col3 = st.columns(3)
    with col1:
        if st.button("📊 Dashboard", use_container_width=True):
            st.session_state.page = 'dashboard'
            st.rerun()
    with col2:
        if st.button("➕ New Campaign", use_container_width=True):
            st.session_state.page = 'new_campaign'
            st.session_state.campaign_config = {}
            st.session_state.df = None
            st.rerun()
    with col3:
        if st.button("📜 History", use_container_width=True):
            st.session_state.page = 'history'
            st.rerun()
    
    st.divider()
    
    # Page Router
    if st.session_state.page == 'dashboard':
        show_dashboard()
    elif st.session_state.page == 'new_campaign':
        show_new_campaign()
    elif st.session_state.page == 'history':
        show_history()

def show_dashboard():
    st.markdown("<h2>Campaign Dashboard</h2>", unsafe_allow_html=True)
    
    campaigns = get_all_campaigns()
    
    # Statistics
    total_campaigns = len(campaigns)
    running = sum(1 for c in campaigns if c['status'] == 'running')
    total_success = sum(c['success_count'] for c in campaigns)
    total_failed = sum(c['failed_count'] for c in campaigns)
    
    col1, col2, col3, col4 = st.columns(4)
    with col1:
        st.markdown(f"""
        <div class="metric-card">
            <div class="metric-value total">{total_campaigns}</div>
            <div class="metric-label">Total Campaigns</div>
        </div>
        """, unsafe_allow_html=True)
    with col2:
        st.markdown(f"""
        <div class="metric-card">
            <div class="metric-value success">{running}</div>
            <div class="metric-label">Running</div>
        </div>
        """, unsafe_allow_html=True)
    with col3:
        st.markdown(f"""
        <div class="metric-card">
            <div class="metric-value success">{total_success}</div>
            <div class="metric-label">Successful</div>
        </div>
        """, unsafe_allow_html=True)
    with col4:
        st.markdown(f"""
        <div class="metric-card">
            <div class="metric-value failed">{total_failed}</div>
            <div class="metric-label">Failed</div>
        </div>
        """, unsafe_allow_html=True)
    
    st.divider()
    
    # Recent Campaigns
    st.markdown("<h3>Recent Campaigns</h3>", unsafe_allow_html=True)
    
    if not campaigns:
        st.info("No campaigns yet. Create your first campaign!")
        return
    
    for campaign in campaigns[:5]:
        with st.container():
            col1, col2, col3, col4 = st.columns([3, 2, 2, 1])
            with col1:
                st.markdown(f"**{campaign['name']}**")
                st.caption(campaign['form_url'][:50] + "...")
            with col2:
                st.caption(f"Created: {campaign['created_at'][:10]}")
            with col3:
                st.caption(f"Status: {campaign['status'].upper()}")
            with col4:
                st.caption(f"{campaign['success_count']}/{campaign['total_rows']}")

def show_new_campaign():
    st.markdown("<h2>Create New Campaign</h2>", unsafe_allow_html=True)
    
    # Step 1: Basic Info
    st.markdown("### 1. Campaign Information")
    campaign_name = st.text_input("Campaign Name", placeholder="Student Feedback Campaign")
    form_url = st.text_input("Google Form Pre-filled Link", placeholder="https://docs.google.com/forms/d/e/.../viewform?...")
    
    st.divider()
    
    # Step 2: Data Source
    st.markdown("### 2. Data Source")
    source_type = st.radio("Source Type", ["Upload File", "Google Sheets URL"], horizontal=True)
    
    df = None
    if source_type == "Upload File":
        uploaded_file = st.file_uploader("Upload Data File", type=["xlsx", "csv", "tsv", "ods", "json"])
        if uploaded_file:
            try:
                if uploaded_file.name.endswith('.csv'):
                    df = pd.read_csv(uploaded_file)
                elif uploaded_file.name.endswith('.tsv'):
                    df = pd.read_csv(uploaded_file, sep='\t')
                else:
                    df = pd.read_excel(uploaded_file)
                st.success(f"✅ Loaded {len(df)} rows, {len(df.columns)} columns")
            except Exception as e:
                st.error(f"Error loading file: {e}")
    else:
        sheets_url = st.text_input("Google Sheets URL", placeholder="https://docs.google.com/spreadsheets/...")
        if sheets_url and st.button("Connect"):
            try:
                resp = requests.get(sheets_url, timeout=30)
                df = pd.read_csv(io.StringIO(resp.text))
                st.success(f"✅ Loaded {len(df)} rows from Google Sheets")
            except Exception as e:
                st.error(f"Error loading sheet: {e}")
    
    if df is not None:
        st.session_state.df = df
        df.columns = df.columns.str.strip()
        
        st.divider()
        
        # Step 3: Form Detection
        st.markdown("### 3. Form Detection & Validation")
        if st.button("🔍 Detect Form Fields"):
            if not form_url:
                st.error("Please enter the Google Form pre-filled link")
            else:
                with st.spinner("Fetching form metadata..."):
                    field_types, options_by_entry, grid_info, questions_order = fetch_form_metadata(form_url)
                
                st.session_state.campaign_config['field_types'] = field_types
                st.session_state.campaign_config['options_by_entry'] = options_by_entry
                st.session_state.campaign_config['form_url'] = form_url.split("/viewform")[0] + "/formResponse"
                
                st.success(f"✅ Form detected: {len(questions_order)} fields")
                
                # Auto-mapping
                prefilled_ids = extract_entry_ids(form_url)
                non_status_cols = [c for c in df.columns if c.lower() not in {"status", "submit"}]
                
                if len(prefilled_ids) == len(non_status_cols):
                    mapping = dict(zip(non_status_cols, prefilled_ids))
                    st.session_state.campaign_config['column_mapping'] = mapping
                    st.success(f"✅ Auto-mapped {len(mapping)} columns")
                else:
                    st.warning("Column count mismatch. Please map manually below.")
                    mapping = {}
                    for col in non_status_cols:
                        selected = st.selectbox(f"Map `{col}` to:", ["-- Select --"] + prefilled_ids)
                        if selected != "-- Select --":
                            mapping[col] = selected
                    st.session_state.campaign_config['column_mapping'] = mapping
        
        if 'column_mapping' in st.session_state.campaign_config:
            st.divider()
            
            # Step 4: Validation
            st.markdown("### 4. Validation & Auto-Correction")
            if st.button("✓ Validate Data"):
                pending_indices = df.index.tolist()
                errors = validate_and_correct(
                    df, pending_indices, 
                    st.session_state.campaign_config['column_mapping'],
                    st.session_state.campaign_config['field_types'],
                    st.session_state.campaign_config['options_by_entry']
                )
                
                if errors:
                    st.error("Validation errors found:")
                    for e in errors[:10]:
                        st.write(f"• {e}")
                else:
                    st.success("✅ All data validated successfully")
                    st.session_state.df = df
            
            st.divider()
            
            # Step 5: Scheduling
            st.markdown("### 5. Schedule Submission")
            col1, col2 = st.columns(2)
            with col1:
                use_random = st.checkbox("Randomize row order", value=True)
                subset_n = st.number_input("Max rows (0 = all)", min_value=0, value=0)
            with col2:
                start_time = st.text_input("Start time", value="10:40")
                end_time = st.text_input("End time", value="10:43")
            
            col3, col4 = st.columns(2)
            with col3:
                delay_min = st.number_input("Min delay (sec)", min_value=1, value=2)
            with col4:
                delay_max = st.number_input("Max delay (sec)", min_value=1, value=5)
            
            st.session_state.campaign_config.update({
                'name': campaign_name,
                'use_random': use_random,
                'subset_n': subset_n,
                'start_time': start_time,
                'end_time': end_time,
                'delay_min': delay_min,
                'delay_max': delay_max
            })
            
            st.divider()
            
            # Step 6: Launch
            st.markdown("### 6. Launch Campaign")
            st.warning("""
            **⚠️ Important for Long Runs:**
            
            To run this campaign even when your screen is off:
            1. Click "Download Offline Script" below
            2. Save the `.py` file to your computer
            3. Open terminal/command prompt
            4. Run: `python offline_runner.py`
            5. Configure your OS power settings to prevent sleep
            
            The script will continue running independently!
            """)
            
            col1, col2 = st.columns(2)
            with col1:
                if st.button("📥 Download Offline Script", type="primary", use_container_width=True):
                    csv_data = df.to_csv(index=False)
                    script = generate_offline_script(st.session_state.campaign_config, csv_data)
                    
                    st.download_button(
                        label="💾 Save Script",
                        data=script,
                        file_name="offline_runner.py",
                        mime="text/plain",
                        use_container_width=True
                    )
                    
                    # Save campaign to database
                    campaign_data = {
                        'id': str(hash(campaign_name + str(datetime.now()))),
                        'name': campaign_name,
                        'form_url': form_url,
                        'source_type': source_type,
                        'total_rows': len(df),
                        'pending_count': len(df),
                        'status': 'ready',
                        'start_time': start_time,
                        'end_time': end_time,
                        'config': st.session_state.campaign_config
                    }
                    save_campaign(campaign_data)
                    st.success("✅ Campaign saved! Download the script above to run offline.")
            
            with col2:
                if st.button("🚀 Run in Browser", use_container_width=True):
                    st.info("Running in browser... Keep this tab open!")
                    run_browser_submission()

def run_browser_submission():
    config = st.session_state.campaign_config
    df = st.session_state.df
    
    pending_indices = df.index.tolist()
    if config['use_random']:
        random.shuffle(pending_indices)
    if config['subset_n'] > 0:
        pending_indices = pending_indices[:config['subset_n']]
    
    total = len(pending_indices)
    
    # Initialize counters
    success_count = 0
    failed_count = 0
    pending_count = total
    
    # Progress and Stats Container
    progress_container = st.container()
    
    with progress_container:
        # Progress Bar Section
        st.markdown("""
        <div class="progress-container">
            <div class="progress-header">
                <div class="progress-title">Submission Progress</div>
                <div class="progress-percentage" id="progress-pct">0%</div>
            </div>
        </div>
        """, unsafe_allow_html=True)
        
        progress_bar = st.progress(0)
        
        # Statistics Cards
        col1, col2, col3, col4 = st.columns(4)
        with col1:
            total_placeholder = st.empty()
            total_placeholder.markdown(f"""
            <div class="metric-card">
                <div class="metric-value total">{total}</div>
                <div class="metric-label">Total Rows</div>
            </div>
            """, unsafe_allow_html=True)
        with col2:
            success_placeholder = st.empty()
            success_placeholder.markdown(f"""
            <div class="metric-card">
                <div class="metric-value success">0</div>
                <div class="metric-label">Successful</div>
            </div>
            """, unsafe_allow_html=True)
        with col3:
            failed_placeholder = st.empty()
            failed_placeholder.markdown(f"""
            <div class="metric-card">
                <div class="metric-value failed">0</div>
                <div class="metric-label">Failed</div>
            </div>
            """, unsafe_allow_html=True)
        with col4:
            pending_placeholder = st.empty()
            pending_placeholder.markdown(f"""
            <div class="metric-card">
                <div class="metric-value pending">{total}</div>
                <div class="metric-label">Pending</div>
            </div>
            """, unsafe_allow_html=True)
    
    st.divider()
    
    # Submission Log
    st.markdown("<h3>Live Submission Log</h3>", unsafe_allow_html=True)
    log_df = pd.DataFrame(columns=["Row", "Status", "Submitted At", "HTTP Code", "Next Entry", "Gap"])
    log_placeholder = st.empty()
    
    for i, idx in enumerate(pending_indices):
        row = df.loc[idx]
        payload = {}
        for col, entry in config['column_mapping'].items():
            val = normalize_value(row[col], config['field_types'].get(entry, "text"))
            payload[entry] = val
        
        # Update progress
        progress_pct = (i + 1) / total
        progress_bar.progress(progress_pct)
        
        # Submit
        try:
            r = requests.post(config['form_url'], data=payload, timeout=30)
            ist_now = datetime.now(IST).strftime("%d-%m-%Y %I:%M:%S %p")
            
            if r.status_code == 200:
                success_count += 1
                status_text = "Submitted"
                
                if i < total - 1:
                    gap_sec = random.randint(config['delay_min'], config['delay_max'])
                    next_time = (datetime.now(IST) + timedelta(seconds=gap_sec)).strftime("%d-%m-%Y %I:%M:%S %p IST")
                    gap_text = format_human(gap_sec)
                else:
                    next_time, gap_text = "FINAL ENTRY", "-"
            else:
                failed_count += 1
                status_text = f"FAILED: HTTP {r.status_code}"
                next_time, gap_text = "-", "-"
        except Exception as e:
            failed_count += 1
            ist_now = datetime.now(IST).strftime("%d-%m-%Y %I:%M:%S %p")
            status_text = f"ERROR: {str(e)}"
            next_time, gap_text = "-", "-"
        
        pending_count -= 1
        
        # Update statistics
        success_placeholder.markdown(f"""
        <div class="metric-card">
            <div class="metric-value success">{success_count}</div>
            <div class="metric-label">Successful</div>
        </div>
        """, unsafe_allow_html=True)
        
        failed_placeholder.markdown(f"""
        <div class="metric-card">
            <div class="metric-value failed">{failed_count}</div>
            <div class="metric-label">Failed</div>
        </div>
        """, unsafe_allow_html=True)
        
        pending_placeholder.markdown(f"""
        <div class="metric-card">
            <div class="metric-value pending">{pending_count}</div>
            <div class="metric-label">Pending</div>
        </div>
        """, unsafe_allow_html=True)
        
        # Update log
        new_row = {
            "Row": idx + 1,
            "Status": status_text,
            "Submitted At": ist_now,
            "HTTP Code": r.status_code if r.status_code == 200 else "Error",
            "Next Entry": next_time,
            "Gap": gap_text
        }
        log_df = pd.concat([log_df, pd.DataFrame([new_row])], ignore_index=True)
        log_placeholder.dataframe(log_df, use_container_width=True, hide_index=True)
        
        # Delay
        if i < total - 1:
            time.sleep(random.randint(config['delay_min'], config['delay_max']))
    
    # Completion
    st.balloons()
    st.success(f"✅ Campaign completed! {success_count} rows submitted successfully.")

def show_history():
    st.markdown("<h2>Campaign History</h2>", unsafe_allow_html=True)
    
    campaigns = get_all_campaigns()
    
    if not campaigns:
        st.info("No campaign history yet.")
        return
    
    for campaign in campaigns:
        with st.expander(f"{campaign['name']} - {campaign['status'].upper()}"):
            col1, col2 = st.columns(2)
            with col1:
                st.write(f"**Form:** {campaign['form_url'][:60]}...")
                st.write(f"**Created:** {campaign['created_at']}")
                st.write(f"**Total Rows:** {campaign['total_rows']}")
            with col2:
                st.write(f"**Success:** {campaign['success_count']}")
                st.write(f"**Failed:** {campaign['failed_count']}")
                st.write(f"**Pending:** {campaign['pending_count']}")

if __name__ == "__main__":
    main()
