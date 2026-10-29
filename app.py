import sys
import os

# Ensure the current directory is in the Python path so local modules can be found
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

import streamlit as st
import tempfile

# Now import from local files
from utils import CSS
from ui import inject_css, view_configure, view_preflight, view_execute, view_results, stepper

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
    
    # Import md here to avoid circular dependency issues if any
    from utils import md
    
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
