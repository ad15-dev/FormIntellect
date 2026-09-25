import io
import json
import random
import re
import time
import urllib.parse
from datetime import datetime, timedelta

import pandas as pd
import pytz
import requests
import streamlit as st

# ============================================================
# CONFIG & THEME
# ============================================================
st.set_page_config(
    page_title="Form Auto-Submitter",
    page_icon="🚀",
    layout="centered",
    initial_sidebar_state="collapsed"
)

# Custom CSS for Minimal Modern Look
st.markdown("""
<style>
    /* Hide default Streamlit footer */
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    
    /* Minimalist Inputs */
    .stTextInput > div > div > input {
        border-radius: 10px;
        padding: 12px;
        background-color: #f8f9fa;
    }
    .stFileUploader > label {
        border-radius: 10px;
        border: 2px dashed #ccc;
    }
    
    /* Progress Bar Styling */
    .stProgress > div > div > div > div {
        background-color: #4CAF50;
    }
    
    /* Card-like containers */
    .css-1r6slb0 {
        border-radius: 15px;
        box-shadow: 0 4px 6px rgba(0,0,0,0.1);
        padding: 20px;
        background-color: white;
    }
</style>
""", unsafe_allow_html=True)

IST = pytz.timezone("Asia/Kolkata")

# Initialize Session State
if 'step' not in st.session_state:
    st.session_state.step = 1
if 'df' not in st.session_state:
    st.session_state.df = None
if 'metadata' not in st.session_state:
    st.session_state.metadata = {}
if 'config' not in st.session_state:
    st.session_state.config = {}

# ============================================================
# UTILITY FUNCTIONS (From Notebook)
# ============================================================

def extract_entry_ids(prefilled_link):
    parsed = urllib.parse.urlparse(prefilled_link)
    pairs = urllib.parse.parse_qsl(parsed.query, keep_blank_values=True)
    return list({k for k, _ in pairs if k.startswith("entry.")})

def _parse_fb_blob(html):
    token = "FB_PUBLIC_LOAD_DATA_"
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
        return {}, {}, []

    field_types = {}
    for m in re.finditer(r'name="(entry\.\d+)"[^>]*type="([^"]+)"', html):
        eid, itype = m.group(1), m.group(2).lower()
        field_types[eid] = itype
        
    options_by_entry, questions_order = {}, []
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
        except Exception:
            pass
            
    return field_types, options_by_entry, questions_order

def normalize_value(raw, entry_type="text"):
    if pd.isna(raw): return "NA"
    text = str(raw).strip()
    return text if text else "NA"

def best_option(value, options):
    if not options or value in options: return value
    lower_map = {o.lower(): o for o in options if isinstance(o, str)}
    return lower_map.get(value.lower(), options[0])

# ============================================================
# STEP 1: UPLOAD & LINK
# ============================================================
def step_1_upload():
    st.title("📝 Form Auto-Submitter")
    st.caption("Minimalist • Fast • Reliable")
    
    st.markdown("### 1. Setup")
    uploaded_file = st.file_uploader("Upload Data File", type=["xlsx", "csv"], label_visibility="collapsed")
    form_url_input = st.text_input("Google Form Pre-filled Link", placeholder="https://docs.google.com/forms/...")
    
    col1, col2 = st.columns(2)
    with col1:
        use_random = st.checkbox("Randomize Order", value=True)
    with col2:
        subset_n = st.number_input("Max Rows (0=All)", min_value=0, value=0)

    if st.button("Next Step →", type="primary", use_container_width=True):
        if not uploaded_file or not form_url_input:
            st.error("Please upload a file and enter the form link.")
            return

        # Load Data
        try:
            if uploaded_file.name.endswith('.csv'):
                df = pd.read_csv(uploaded_file)
            else:
                df = pd.read_excel(uploaded_file)
            df.columns = df.columns.str.strip()
        except Exception as e:
            st.error(f"Error reading file: {e}")
            return

        # Fetch Metadata
        with st.spinner("Analyzing Form Structure..."):
            field_types, options_by_entry, questions_order = fetch_form_metadata(form_url_input)
        
        if not questions_order:
            st.warning("Could not detect form fields automatically. Ensure link is valid.")
            # Allow proceeding for manual mapping if needed, but ideally stop
            # return 

        # Store in Session
        st.session_state.df = df
        st.session_state.metadata = {
            "field_types": field_types,
            "options_by_entry": options_by_entry,
            "questions_order": questions_order,
            "form_url": form_url_input.split("/viewform")[0] + "/formResponse",
            "prefilled_link": form_url_input
        }
        st.session_state.config = {
            "use_random": use_random,
            "subset_n": subset_n
        }
        st.session_state.step = 2
        st.rerun()

# ============================================================
# STEP 2: COLUMN MAPPING
# ============================================================
def step_2_mapping():
    st.title("🔗 Map Columns")
    df = st.session_state.df
    meta = st.session_state.metadata
    
    st.info("Verify that your Excel columns match the Form fields below.")
    
    prefilled_ids = extract_entry_ids(meta['prefilled_link'])
    non_status_cols = [c for c in df.columns if c.lower() not in {"status", "submit"}]
    
    # Auto-map logic
    if len(prefilled_ids) == len(non_status_cols):
        initial_map = dict(zip(non_status_cols, prefilled_ids))
    else:
        initial_map = {col: "" for col in non_status_cols}

    mapping = {}
    cols = st.columns(2)
    for i, col_name in enumerate(non_status_cols):
        with cols[i % 2]:
            options = ["-- Select Entry ID --"] + prefilled_ids
            default_idx = 0
            if col_name in initial_map and initial_map[col_name] in options:
                default_idx = options.index(initial_map[col_name])
            
            selected_id = st.selectbox(f"`{col_name}`", options, index=default_idx, key=f"map_{col_name}")
            if selected_id != "-- Select Entry ID --":
                mapping[col_name] = selected_id

    if st.button("Next Step →", type="primary", use_container_width=True):
        if len(mapping) != len(prefilled_ids):
            st.error("Please map all form fields.")
            return
        
        st.session_state.config['column_mapping'] = mapping
        st.session_state.step = 3
        st.rerun()

# ============================================================
# STEP 3: TIME CONFIGURATION
# ============================================================
def step_3_time():
    st.title("⏰ Schedule Submission")
    
    now = datetime.now(IST)
    
    col1, col2 = st.columns(2)
    with col1:
        start_time = st.time_input("Start Time", value=now.time())
    with col2:
        end_time = st.time_input("End Time", value=(now + timedelta(minutes=5)).time())
        
    st.divider()
    st.subheader("Human-like Delays")
    col3, col4 = st.columns(2)
    with col3:
        delay_min = st.number_input("Min Delay (sec)", min_value=1, value=2)
    with col4:
        delay_max = st.number_input("Max Delay (sec)", min_value=1, value=5)

    if st.button("Review & Launch →", type="primary", use_container_width=True):
        st.session_state.config['start_time'] = start_time
        st.session_state.config['end_time'] = end_time
        st.session_state.config['delay_min'] = delay_min
        st.session_state.config['delay_max'] = delay_max
        st.session_state.step = 4
        st.rerun()

# ============================================================
# STEP 4: REVIEW & DEPLOY
# ============================================================
def step_4_deploy():
    st.title("🚀 Ready to Launch")
    
    config = st.session_state.config
    df = st.session_state.df
    
    st.success("Configuration Complete!")
    
    st.metric(label="Rows to Process", value=len(df))
    st.metric(label="Mapped Fields", value=len(config['column_mapping']))
    
    st.divider()
    
    st.subheader("Choose Execution Mode")
    
    tab1, tab2 = st.tabs(["🌐 Run in Browser", "💻 Run Offline (Screen Off)"])
    
    with tab1:
        st.warning("Keep this tab open. If you close it, submission stops.")
        if st.button("Start Submission Now", type="primary", use_container_width=True):
            run_submission_directly()
            
    with tab2:
        st.info("Best for long lists. Download the script and run it on your computer. You can turn off your screen.")
        
        script_content = generate_offline_script()
        
        st.download_button(
            label="📥 Download Offline Runner (.py)",
            data=script_content,
            file_name="offline_runner.py",
            mime="text/plain",
            type="secondary",
            use_container_width=True
        )

# ============================================================
# HELPER: GENERATE OFFLINE SCRIPT
# ============================================================
def generate_offline_script():
    config = st.session_state.config
    meta = st.session_state.metadata
    df = st.session_state.df
    
    csv_data = df.to_csv(index=False)
    
    script = f'''
import pandas as pd
import requests
import time
import random
import io
from datetime import datetime, timedelta
import pytz

# --- CONFIG ---
IST = pytz.timezone("Asia/Kolkata")
FORM_URL = "{meta['form_url']}"
COLUMN_MAPPING = {json.dumps(config['column_mapping'])}
FIELD_TYPES = {json.dumps(meta['field_types'])}
OPTIONS_BY_ENTRY = {json.dumps(meta['options_by_entry'])}
START_TIME_STR = "{config['start_time']}"
END_TIME_STR = "{config['end_time']}"
DELAY_MIN = {config['delay_min']}
DELAY_MAX = {config['delay_max']}

# --- DATA ---
CSV_DATA = """{csv_data}"""

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
    df = pd.read_csv(io.StringIO(CSV_DATA))
    
    if "Status" not in df.columns:
        df["Status"] = ""
        
    pending_indices = df[df["Status"].str.lower() != "submitted"].index.tolist()
    
    # Time Logic
    now = datetime.now(IST)
    today = now.date()
    
    sh, sm = map(int, START_TIME_STR.split(":")[:2])
    eh, em = map(int, END_TIME_STR.split(":")[:2])
    
    start_dt = IST.localize(datetime(today.year, today.month, today.day, sh, sm))
    end_dt = IST.localize(datetime(today.year, today.month, today.day, eh, em))
    
    if end_dt <= start_dt:
        end_dt += timedelta(days=1)
        
    if now < start_dt:
        wait_sec = (start_dt - now).total_seconds()
        print(f"⏳ Waiting until {{start_dt.strftime('%H:%M')}}...")
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
    df.to_csv("submitted_results.csv", index=False)
    print("Saved to submitted_results.csv")

if __name__ == "__main__":
    main()
'''
    return script

# ============================================================
# HELPER: RUN DIRECTLY (With Progress Bar)
# ============================================================
def run_submission_directly():
    config = st.session_state.config
    meta = st.session_state.metadata
    df = st.session_state.df
    
    # Prepare Data
    pending_indices = df.index.tolist()
    if config['use_random']:
        random.shuffle(pending_indices)
    if config['subset_n'] > 0:
        pending_indices = pending_indices[:config['subset_n']]
        
    total = len(pending_indices)
    
    # UI Layout for Progress
    progress_bar = st.progress(0)
    status_text = st.empty()
    log_container = st.container(height=200)
    
    submitted_count = 0
    
    for i, idx in enumerate(pending_indices):
        row = df.loc[idx]
        payload = {}
        for col, entry in config['column_mapping'].items():
            val = normalize_value(row[col], meta['field_types'].get(entry, "text"))
            payload[entry] = val
            
        # Update UI
        progress_pct = (i + 1) / total
        progress_bar.progress(progress_pct)
        status_text.text(f"Submitting Row {i+1}/{total}...")
        
        # Submit
        try:
            r = requests.post(meta['form_url'], data=payload, timeout=30)
            if r.status_code == 200:
                submitted_count += 1
                log_container.write(f"✅ Row {idx+1} Success")
            else:
                log_container.write(f"❌ Row {idx+1} Failed (HTTP {r.status_code})")
        except Exception as e:
            log_container.write(f"❌ Row {idx+1} Error: {e}")
            
        # Delay
        if i < total - 1:
            delay = random.randint(config['delay_min'], config['delay_max'])
            time.sleep(delay)
            
    status_text.text("Submission Complete!")
    st.balloons()
    st.success(f"Successfully submitted {submitted_count} rows.")

# ============================================================
# MAIN ROUTER
# ============================================================
def main():
    if st.session_state.step == 1:
        step_1_upload()
    elif st.session_state.step == 2:
        step_2_mapping()
    elif st.session_state.step == 3:
        step_3_time()
    elif st.session_state.step == 4:
        step_4_deploy()

if __name__ == "__main__":
    main()