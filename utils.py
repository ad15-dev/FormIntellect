import os
import re
import html
from datetime import datetime, time as dtime, timedelta
import pytz

IST = pytz.timezone("Asia/Kolkata")
MIN_SECONDS_PER_ROW = 2
STATUS_NAMES = {"status", "submit", "submitted"}
BLANK = "NA"
SUPPORTED_EXT = (".xlsx", ".xls", ".csv", ".tsv", ".ods", ".json")
ALLOW_LOCAL_PATHS = os.environ.get("FORMINTELLECT_ALLOW_LOCAL_PATHS") == "1"

CSS = """
:root{
  --bg:#0B0D10; --surface:#14171C; --surface2:#0F1216; --border:#252A31; --border2:#323843;
  --text:#ECEEF1; --muted:#9AA3AE; --faint:#6B7480;
  --green:#3ECF8E; --green-h:#2EB87A; --green-bg:rgba(62,207,142,.10);
  --red:#F97066; --red-bg:rgba(249,112,102,.08);
  --amber:#F5B544; --amber-bg:rgba(245,181,68,.08);
  --blue:#6CB6FF; --blue-bg:rgba(108,182,255,.08);
}
html, body, .stApp, [class*="st-"]{font-family:'Inter',system-ui,-apple-system,'Segoe UI',Roboto,sans-serif;}
.stApp{background:radial-gradient(1200px 500px at 50% -200px, rgba(62,207,142,.07), transparent 70%), var(--bg); color:var(--text);}
.stApp p, .stApp li, .stApp label, .stMarkdown{color:var(--text);}
header[data-testid="stHeader"]{background:transparent;}
#MainMenu, footer{visibility:hidden;}
.block-container{max-width:960px; padding-top:2rem; padding-bottom:4rem;}
[data-testid="stCaptionContainer"], .stCaption{color:var(--muted)!important;}
.fi-hero{display:flex; align-items:center; gap:14px; margin-bottom:4px;}
.fi-logo{width:40px; height:40px; border-radius:11px; background:linear-gradient(135deg,#3ECF8E,#22A06B);
  display:flex; align-items:center; justify-content:center; box-shadow:0 6px 24px rgba(62,207,142,.25);
  color:#0B0D10; font-size:1.3rem; font-weight:800;}
.fi-name{font-size:1.75rem; font-weight:700; letter-spacing:-0.02em; line-height:1.1;}
.fi-name span{color:var(--green);}
.fi-tag{color:var(--muted); font-size:.93rem; margin:4px 0 26px 54px;}
.fi-stepper{display:flex; align-items:center; gap:12px; margin:0 0 30px; padding:14px 18px;
  background:var(--surface); border:1px solid var(--border); border-radius:14px;}
.fi-st{display:flex; align-items:center; gap:9px; color:var(--faint); font-size:.86rem; font-weight:500; white-space:nowrap;}
.fi-st .n{width:28px; height:28px; border-radius:50%; border:1px solid var(--border2);
  display:flex; align-items:center; justify-content:center; font-family:'JetBrains Mono',ui-monospace,SFMono-Regular,Menlo,Consolas,monospace; font-size:.74rem;}
.fi-st .l small{display:block; color:var(--faint); font-size:.7rem; font-weight:400;}
.fi-st.active{color:var(--text);}
.fi-st.active .n{border-color:var(--green); color:var(--green); background:var(--green-bg); box-shadow:0 0 0 4px rgba(62,207,142,.08);}
.fi-st.done{color:var(--text);}
.fi-st.done .n{background:var(--green); border-color:var(--green); color:#0B0D10; font-weight:700;}
.fi-st-line{flex:1; height:2px; background:var(--border); min-width:14px; border-radius:2px;}
.fi-st-line.done{background:var(--green);}
.fi-sec h2{font-size:1.35rem; font-weight:650; margin:0; letter-spacing:-.01em; color:var(--text);}
.fi-sec p{color:var(--muted); font-size:.92rem; margin:6px 0 20px;}
div[data-testid="stVerticalBlockBorderWrapper"]{
  background:var(--surface); border:1px solid var(--border)!important; border-radius:16px;}
.fi-step{display:flex; align-items:center; gap:12px; margin-bottom:2px;}
.fi-step b{font-family:'JetBrains Mono',ui-monospace,SFMono-Regular,Menlo,Consolas,monospace; font-size:.72rem; font-weight:500;
  color:var(--green); background:var(--green-bg); border:1px solid rgba(62,207,142,.3); border-radius:7px; padding:3px 8px;}
.fi-step .t{font-weight:600; font-size:1rem;}
.fi-step-d{color:var(--muted); font-size:.84rem; margin:0 0 10px 46px;}
[data-testid="stWidgetLabel"] p{color:var(--muted)!important; font-size:.82rem!important; font-weight:500;}
div[data-baseweb="input"], div[data-baseweb="base-input"], div[data-baseweb="select"] > div, div[data-baseweb="textarea"]{
  background:var(--surface2)!important; border-color:var(--border2)!important; border-radius:10px!important;}
div[data-baseweb="input"]:focus-within{border-color:var(--green)!important;}
input, textarea{color:var(--text)!important;}
[data-testid="stFileUploaderDropzone"]{background:var(--surface2); border:1px dashed var(--border2); border-radius:12px;}
[data-testid="stFileUploaderDropzone"] *{color:var(--muted);}
div[role="radiogroup"] label p{color:var(--text)!important; font-size:.9rem!important;}
[data-testid="stSlider"] [role="slider"]{background:var(--green)!important; border-color:var(--green)!important;}
[data-testid="stTickBarMin"], [data-testid="stTickBarMax"], [data-testid="stSliderThumbValue"]{color:var(--muted)!important;}
button[data-baseweb="tab"]{color:var(--muted); font-weight:500;}
button[data-baseweb="tab"][aria-selected="true"]{color:var(--text);}
div[data-baseweb="tab-highlight"]{background:var(--green)!important;}
div[data-baseweb="tab-border"]{background:var(--border)!important;}
[data-testid="stExpander"] details{background:var(--surface); border:1px solid var(--border)!important; border-radius:12px;}
[data-testid="stExpander"] summary p{color:var(--muted)!important; font-size:.86rem;}
[data-testid="stCode"] pre, .stCode pre{background:var(--surface2)!important; border:1px solid var(--border); border-radius:10px;}
[data-testid="stCode"] code{color:#C9D1D9!important;}
.stButton > button, .stDownloadButton > button{
  border-radius:10px; border:1px solid var(--border2); background:var(--surface);
  color:var(--text); font-weight:600; padding:.6rem 1.1rem; transition:all .15s ease;}
.stButton > button:hover, .stDownloadButton > button:hover{border-color:var(--green); color:var(--green);}
.stButton > button[kind="primary"], 
button[data-testid="stBaseButton-primary"],
div[data-baseweb="button"] button[kind="primary"] {
  background:var(--green)!important; 
  border-color:var(--green)!important; 
  color:#0B0D10!important;
}
.stButton > button[kind="primary"]:hover, 
button[data-testid="stBaseButton-primary"]:hover,
div[data-baseweb="button"] button[kind="primary"]:hover {
  background:var(--green-h)!important; 
  border-color:var(--green-h)!important; 
  color:#0B0D10!important;
}
.fi-call{display:flex; gap:12px; padding:14px 16px; border-radius:12px; border:1px solid; margin:6px 0 16px; font-size:.9rem; line-height:1.5;}
.fi-call .ic{width:22px; height:22px; border-radius:50%; flex:none; display:flex; align-items:center; justify-content:center; font-size:.72rem; font-weight:700; margin-top:1px;}
.fi-call b{font-weight:600;}
.fi-call.info{background:var(--blue-bg); border-color:rgba(108,182,255,.3);} .fi-call.info .ic{background:rgba(108,182,255,.18); color:var(--blue);}
.fi-call.ok{background:var(--green-bg); border-color:rgba(62,207,142,.3);} .fi-call.ok .ic{background:rgba(62,207,142,.18); color:var(--green);}
.fi-call.warn{background:var(--amber-bg); border-color:rgba(245,181,68,.3);} .fi-call.warn .ic{background:rgba(245,181,68,.18); color:var(--amber);}
.fi-call.err{background:var(--red-bg); border-color:rgba(249,112,102,.3);} .fi-call.err .ic{background:rgba(249,112,102,.18); color:var(--red);}
.fi-err{background:var(--red-bg); border:1px solid rgba(249,112,102,.35); border-radius:16px; padding:18px 20px; margin:16px 0 12px;}
.fi-err-h{display:flex; gap:12px; align-items:flex-start;}
.fi-err-h .ic{width:30px; height:30px; border-radius:50%; background:rgba(249,112,102,.2); color:var(--red);
  font-weight:700; display:flex; align-items:center; justify-content:center; flex:none;}
.fi-err .t{font-weight:650; font-size:1.05rem; color:var(--text);}
.fi-err .w{color:var(--muted); font-size:.82rem; margin-top:3px;}
.fi-err .w span{color:var(--red); font-weight:500;}
.fi-err-grid{display:grid; grid-template-columns:1fr 1fr; gap:12px; margin-top:16px;}
.fi-err-box{background:rgba(0,0,0,.18); border:1px solid var(--border); border-radius:12px; padding:12px 14px;}
.fi-err-box .k{font-size:.7rem; text-transform:uppercase; letter-spacing:.09em; color:var(--muted); margin-bottom:5px;}
.fi-err-box .v{font-size:.9rem; line-height:1.55; color:var(--text);}
.fi-err-box.fix{border-color:rgba(62,207,142,.3);} .fi-err-box.fix .k{color:var(--green);}
.fi-checks{background:var(--surface); border:1px solid var(--border); border-radius:14px; margin-bottom:18px;}
.fi-check{display:flex; gap:12px; padding:13px 16px; border-top:1px solid var(--border);}
.fi-check:first-child{border-top:none;}
.fi-check .ic{width:22px; height:22px; border-radius:50%; flex:none; margin-top:1px;
  display:flex; align-items:center; justify-content:center; font-size:.7rem; font-weight:700;}
.fi-check.ok .ic{background:rgba(62,207,142,.15); color:var(--green);}
.fi-check.warn .ic{background:rgba(245,181,68,.15); color:var(--amber);}
.fi-check.err .ic{background:rgba(249,112,102,.15); color:var(--red);}
.fi-check .t{font-weight:600; font-size:.9rem;}
.fi-check .d{color:var(--muted); font-size:.82rem; margin-top:2px; word-break:break-word;}
.fi-tw{border:1px solid var(--border); border-radius:12px; overflow:auto; max-height:440px; background:var(--surface); margin:6px 0 12px;}
.fi-table{width:100%; border-collapse:collapse; font-size:.85rem;}
.fi-table th{position:sticky; top:0; text-align:left; color:var(--muted); font-weight:500; font-size:.72rem;
  text-transform:uppercase; letter-spacing:.06em; padding:10px 12px; border-bottom:1px solid var(--border); background:var(--surface2);}
.fi-table td{padding:10px 12px; border-bottom:1px solid var(--border); vertical-align:top; color:var(--text); line-height:1.45;}
.fi-table tr:last-child td{border-bottom:none;}
.fi-table td.row{font-family:'JetBrains Mono',ui-monospace,SFMono-Regular,Menlo,Consolas,monospace; color:var(--red); white-space:nowrap;}
.fi-table td.val{font-family:'JetBrains Mono',ui-monospace,SFMono-Regular,Menlo,Consolas,monospace; color:#FFC2BC;}
.fi-table td.fix{color:#B7F0D4;}
.fi-table td.mono{font-family:'JetBrains Mono',ui-monospace,SFMono-Regular,Menlo,Consolas,monospace; font-size:.8rem; color:var(--muted);}
.fi-stats{display:grid; grid-template-columns:repeat(4,1fr); gap:12px; margin:6px 0 14px;}
.fi-stat{background:var(--surface); border:1px solid var(--border); border-radius:14px; padding:14px 16px;}
.fi-stat-l{color:var(--muted); font-size:.72rem; text-transform:uppercase; letter-spacing:.08em;}
.fi-stat-v{font-size:1.5rem; font-weight:600; margin-top:4px; font-family:'JetBrains Mono',ui-monospace,SFMono-Regular,Menlo,Consolas,monospace;}
.fi-stat-v.sm{font-size:1.04rem;}
.fi-stat-s{color:var(--muted); font-size:.75rem; margin-top:2px; min-height:1em;}
.fi-stat-v.g{color:var(--green);} .fi-stat-v.r{color:var(--red);} .fi-stat-v.a{color:var(--amber);}
.fi-bar{height:8px; background:var(--border); border-radius:99px; overflow:hidden; margin-bottom:18px; display:flex;}
.fi-bar > .ok{height:100%; background:var(--green); transition:width .4s ease;}
.fi-bar > .bad{height:100%; background:var(--red); transition:width .4s ease;}
.fi-tbl-h{display:flex; justify-content:space-between; align-items:center; margin:4px 0 8px;
  font-size:.76rem; text-transform:uppercase; letter-spacing:.08em; color:var(--muted);}
.fi-tbl{display:flex; flex-direction:column-reverse; max-height:440px; overflow-y:auto;
  background:var(--surface2); border:1px solid var(--border); border-radius:12px;}
.fi-row{font-size:.84rem; line-height:1.5; padding:10px 14px 10px 32px; border-top:1px solid var(--border);
  position:relative; color:var(--text); word-break:break-word;}
.fi-row:last-child{border-top:none;}
.fi-row::before{content:''; position:absolute; left:14px; top:17px; width:8px; height:8px; border-radius:50%; background:var(--green);}
.fi-row.final::before{box-shadow:0 0 0 3px rgba(62,207,142,.25);}
.fi-row.fail{background:var(--red-bg);}
.fi-row.fail::before{background:var(--red);}
.fi-row.note::before{background:var(--faint);}
.fi-row.note, .fi-row.empty{color:var(--muted);}
.fi-row .r{font-weight:600;}
.fi-row .k{color:var(--muted);}
.fi-row .m{font-family:'JetBrains Mono',ui-monospace,SFMono-Regular,Menlo,Consolas,monospace; font-size:.8rem;}
.fi-row .why{display:block; color:#FFC2BC; margin-top:3px;}
.fi-row .todo{display:block; color:#B7F0D4; margin-top:2px;}
.fi-row .pill{font-family:'JetBrains Mono',ui-monospace,SFMono-Regular,Menlo,Consolas,monospace; font-size:.74rem; color:var(--green); background:var(--green-bg); border-radius:6px; padding:1px 7px; margin-left:6px;}
.fi-row.fail .pill{color:var(--red); background:rgba(249,112,102,.12);}
.fi-prev{display:grid; grid-template-columns:1fr auto 1fr; gap:10px; align-items:center; background:var(--surface2);
  border:1px solid var(--border); border-radius:12px; padding:12px 16px; margin-top:8px;}
.fi-prev .k{font-size:.7rem; color:var(--muted); text-transform:uppercase; letter-spacing:.08em;}
.fi-prev .v{font-family:'JetBrains Mono',ui-monospace,SFMono-Regular,Menlo,Consolas,monospace; font-size:.95rem; margin-top:2px;}
.fi-prev .arrow{color:var(--faint); font-size:1.2rem;}
@media (max-width:720px){
  .fi-stats{grid-template-columns:repeat(2,1fr);}
  .fi-err-grid{grid-template-columns:1fr;}
  .fi-st .l{display:none;} .fi-st.active .l{display:inline;}
  .fi-tag{margin-left:0;}
}
"""

def esc(v):
    return html.escape(str(v))

def fmt_clock(dt):
    return dt.strftime("%I:%M:%S %p").lstrip("0")

def format_human(seconds):
    m, s = divmod(int(seconds), 60)
    h, m = divmod(m, 60)
    if h:
        return f"{h}h {m}m"
    if m:
        return f"{m}m {s}s"
    return f"{s}s"

def format_mm_ss(seconds):
    m, s = divmod(int(seconds), 60)
    h, m = divmod(m, 60)
    if h:
        return f"{h}:{m:02d}:{s:02d}"
    return f"{m:02d}:{s:02d}"