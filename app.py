import sys
import os

# Add current directory to path for local module imports
current_dir = os.path.dirname(os.path.abspath(__file__))
if current_dir not in sys.path:
    sys.path.insert(0, current_dir)

import streamlit as st
import tempfile

# Import everything at the top level with error handling
try:
    from utils import CSS, md
    from ui import inject_css, view_configure, view_preflight, view_execute, view_results, stepper
except ImportError as e:
    st.error(f"❌ Failed to import modules: {e}")
    st.error(f"Current directory: {os.getcwd()}")
    st.error(f"Script directory: {current_dir}")
    st.error(f"Python path: {sys.path}")
    
    # List files in current directory
    try:
        files = os.listdir(current_dir)
        st.error(f"Files in directory: {files}")
    except Exception as list_err:
        st.error(f"Cannot list directory: {list_err}")
    
    st.stop()

def current_stage():
    job = st.session_state.get("job")
    if job: 
        return "execute" if job.active else "results"
    return "preflight" if st.session_state.get("prep") else "configure"

def main():
    st.set_page_config(page_title="FormIntellect", page_icon="✅", layout="centered")
    inject_css(CSS)
    
    if "workdir" not in st.session_state:
        st.session_state.workdir = tempfile.mkdtemp(prefix="formintellect_")
    
    md('<div class="fi-hero"><div class="fi-logo">✓</div>'
       '<div class="fi-name">Form<span>Intellect</span></div></div>'
       '<div class="fi-tag">Checked, paced Google Form submissions from your spreadsheet.</div>')

    stage = current_stage()
    stepper(stage)

    if stage == "configure": 
        view_configure()
    elif stage == "preflight": 
        view_preflight(st.session_state.prep)
    elif stage == "execute": 
        view_execute(st.session_state.job)
    else: 
        view_results(st.session_state.job)

if __name__ == "__main__":
    main()
