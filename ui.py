import streamlit as st
import os
from utils import CSS, esc, fmt_clock, format_human, format_mm_ss, IST
from core import Prepared, Job, AppError, explain_exception, console_text, reroll_schedule, sheet_row

STAGES = [("configure", "Set up", "data, form, timing"), ("preflight", "Check", "review before sending"),
          ("execute", "Send", "live progress"), ("results", "Results", "summary & files")]

STATUS_LABEL = {
    "waiting": ("Waiting", "a", "for the start time"),
    "running": ("Sending", "g", "in progress"),
    "done": ("Finished", "g", "all rows processed"),
    "stopped": ("Stopped", "a", "stopped by you"),
    "error": ("Error", "r", "see the explanation"),
}

def md(html_text): st.markdown(html_text, unsafe_allow_html=True)
def inject_css(css): md(f"<style>{css}</style>")

def stat(label, value, sub="", tone="", small=False, raw_value=False):
    cls = f"fi-stat-v {tone} {'sm' if small else ''}"
    v = value if raw_value else esc(value)
    return (f'<div class="fi-stat"><div class="fi-stat-l">{esc(label)}</div>'
            f'<div class="{cls}">{v}</div><div class="fi-stat-s">{esc(sub)}</div></div>')

def stats_grid(*cards): md('<div class="fi-stats">' + "".join(cards) + "</div>")
def callout(kind, text_html):
    icon = {"info": "i", "ok": "✓", "warn": "!", "err": "✕"}[kind]
    md(f'<div class="fi-call {kind}"><span class="ic">{icon}</span><div>{text_html}</div></div>')

def html_table(columns, rows, classes=None):
    classes = classes or [""] * len(columns)
    head = "".join(f"<th>{esc(c)}</th>" for c in columns)
    body = "".join("<tr>" + "".join(f'<td class="{classes[i]}">{esc(v)}</td>' for i, v in enumerate(r)) + "</tr>" for r in rows)
    md(f'<div class="fi-tw"><table class="fi-table"><thead><tr>{head}</tr></thead><tbody>{body}</tbody></table></div>')

def error_card(title, where="", why="", fix="", problems=None, technical=""):
    where_html = f'<div class="w">Where: <span>{esc(where)}</span></div>' if where else ""
    boxes = ""
    if why or fix:
        boxes = ('<div class="fi-err-grid">'
                 f'<div class="fi-err-box"><div class="k">Why it happened</div><div class="v">{esc(why)}</div></div>'
                 f'<div class="fi-err-box fix"><div class="k">How to fix it</div><div class="v">{esc(fix)}</div></div>'
                 '</div>')
    md(f'<div class="fi-err"><div class="fi-err-h"><span class="ic">!</span>'
       f'<div><div class="t">{esc(title)}</div>{where_html}</div></div>{boxes}</div>')
    if problems:
        html_table(["Sheet row", "Column", "Current value", "What's wrong", "How to fix"],
                   [[p.row, p.column, p.value, p.problem, p.fix] for p in problems[:300]],
                   ["row", "", "val", "", "fix"])
        if len(problems) > 300: st.caption(f"Showing first 300 of {len(problems)} problems.")
    if technical:
        with st.expander("Technical details"): st.code(technical, language=None)

def show_error(exc: AppError): error_card(exc.title, exc.where, exc.why, exc.fix, exc.problems, exc.technical)

def stepper(stage):
    idx = [k for k, _, _ in STAGES].index(stage)
    parts = []
    for i, (_, label, sub) in enumerate(STAGES):
        cls = "done" if i < idx else "active" if i == idx else ""
        mark = "✓" if i < idx else str(i + 1)
        parts.append(f'<div class="fi-st {cls}"><span class="n">{mark}</span>'
                     f'<span class="l">{label}<small>{sub}</small></span></div>')
        if i < len(STAGES) - 1: parts.append(f'<div class="fi-st-line {"done" if i < idx else ""}"></div>')
    md('<div class="fi-stepper">' + "".join(parts) + "</div>")

def section_head(title, desc): md(f'<div class="fi-sec"><h2>{esc(title)}</h2><p>{esc(desc)}</p></div>')
def step_head(num, title, desc): md(f'<div class="fi-step"><b>{num}</b><span class="t">{esc(title)}</span></div><div class="fi-step-d">{esc(desc)}</div>')

def schedule_preview(start_t, end_t, srange, erange):
    from core import base_window
    s, e = base_window(start_t, end_t)
    s_lo, s_hi = s + __import__('datetime').timedelta(seconds=min(srange)), s + __import__('datetime').timedelta(seconds=max(srange))
    e_lo, e_hi = e + __import__('datetime').timedelta(seconds=min(erange)), e + __import__('datetime').timedelta(seconds=max(erange))
    rng = lambda a, b: fmt_clock(a) if a == b else f"{fmt_clock(a)} – {fmt_clock(b)}"
    md('<div class="fi-prev">'
       f'<div><div class="k">First entry around</div><div class="v">{rng(s_lo, s_hi)}</div></div>'
       '<div class="arrow">→</div>'
       f'<div><div class="k">Last entry by</div><div class="v">{rng(e_lo, e_hi)}</div></div>'
       '</div>')
    if __import__('datetime').datetime.now(IST) > s_hi:
        st.caption("Start time passed, sending begins immediately.")

def view_configure():
    saved = st.session_state.get("inp", {})
    section_head("Set up your run", "Four quick steps. Nothing is sent until you've reviewed everything.")
    with st.container(border=True):
        step_head("01", "Your data", "Spreadsheet with one row per response.")
        kinds = ["Upload a file", "Google Sheet link", "Enter data manually"] + (["File on this computer"] if __import__('os').environ.get("FORMINTELLECT_ALLOW_LOCAL_PATHS") == "1" else [])
        kind_keys = {"Upload a file": "upload", "Google Sheet link": "sheet", "Enter data manually": "manual", "File on this computer": "local"}
        prev = next((k for k, v in kind_keys.items() if v == saved.get("source_kind")), kinds[0])
        kind_label = st.radio("Where is your data?", kinds, horizontal=True, key="source_kind", index=kinds.index(prev) if prev in kinds else 0, label_visibility="collapsed")
        kind = kind_keys[kind_label]
        upload, sheet_url, local_path, manual_data = None, "", "", ""
        if kind == "upload":
            upload = st.file_uploader("Spreadsheet file", type=[e.lstrip(".") for e in __import__('utils').SUPPORTED_EXT], key="upload", help="Excel, CSV, TSV, etc.")
            if not upload and saved.get("upload"): st.caption(f"Using earlier file: {saved['upload'][0]}")
        elif kind == "sheet":
            sheet_url = st.text_input("Google Sheet link", value=saved.get("sheet_url", ""), key="sheet_url", placeholder="https://docs.google.com/spreadsheets/d/…", help="Share as 'Anyone with link'.")
        elif kind == "manual":
            manual_data = st.text_area("Paste CSV data here", value=saved.get("manual_data", ""), key="manual_data", height=200, help="Include headers. Example:\nName,Email\nJohn,john@example.com")
            st.caption("Tip: First line must be headers.")
        else:
            local_path = st.text_input("Full path to file", value=saved.get("local_path", ""), key="local_path", placeholder="C:\\data\\responses.xlsx")
        st.caption("Tip: A column named **Status** tracks sent rows.")

    with st.container(border=True):
        step_head("02", "Your Google Form", "Paste the pre-filled link.")
        link = st.text_input("Pre-filled form link", value=saved.get("link", ""), key="link", placeholder="https://docs.google.com/forms/d/e/…/viewform?entry.123=…")
        with st.expander("How do I get a pre-filled link?"):
            md("1. Open form in Google Forms (edit mode).\n"
               "2. Click **⋮** menu → **Get pre-filled link**.\n"
               "3. Type any answer in every question.\n"
               "4. Click **Get link** → **Copy link**, and paste above.\n\n"
               "Columns should match question order or be named `entry.123456789`.")

    with st.container(border=True):
        step_head("03", "Which rows", "Rows marked 'submitted' are skipped.")
        use_random = st.toggle("Pick rows at random", value=bool(saved.get("random", False)), key="random", help="Off: top to bottom. On: random order.")
        sample_n = 0
        if use_random: sample_n = st.number_input("How many rows (0 = all)", min_value=0, step=1, value=int(saved.get("sample_n", 0)), key="sample_n")

    with st.container(border=True):
        step_head("04", "When to submit", "Entries spread randomly between these times (IST).")
        c1, c2 = st.columns(2)
        start = c1.time_input("Start time", value=saved.get("start", __import__('datetime').time(9, 0)), step=60, key="start")
        end = c2.time_input("End time", value=saved.get("end", __import__('datetime').time(11, 0)), step=60, key="end", help="If end < start, runs past midnight.")
        c3, c4 = st.columns(2)
        s_min = saved.get("srange", (0, 0))[0] if isinstance(saved.get("srange"), tuple) else 0
        s_max = saved.get("srange", (0, 0))[1] if isinstance(saved.get("srange"), tuple) else 0
        e_min = saved.get("erange", (0, 0))[0] if isinstance(saved.get("erange"), tuple) else 0
        e_max = saved.get("erange", (0, 0))[1] if isinstance(saved.get("erange"), tuple) else 0
        srange_min = c3.number_input("Min extra wait before first entry (seconds)", min_value=0, max_value=3600, value=int(s_min), step=5, key="srange_min")
        srange_max = c3.number_input("Max extra wait before first entry (seconds)", min_value=0, max_value=3600, value=int(s_max), step=5, key="srange_max")
        erange_min = c4.number_input("Min extra time after end (seconds)", min_value=0, max_value=3600, value=int(e_min), step=5, key="erange_min")
        erange_max = c4.number_input("Max extra time after end (seconds)", min_value=0, max_value=3600, value=int(e_max), step=5, key="erange_max")
        srange = (srange_min, srange_max)
        erange = (erange_min, erange_max)
        schedule_preview(start, end, srange, erange)

    if st.button("Check everything →", type="primary", width="stretch"):
        inp = dict(source_kind=kind, upload=(upload.name, upload.getvalue()) if upload else saved.get("upload"), sheet_url=sheet_url, local_path=local_path, manual_data=manual_data, link=link, random=bool(use_random), sample_n=int(sample_n), start=start, end=end, srange=tuple(srange), erange=tuple(erange))
        st.session_state.inp = inp
        try:
            with st.spinner("Reading data and checking form…"):
                from core import prepare
                st.session_state.prep = prepare(inp, st.session_state.workdir)
            st.rerun()
        except AppError as exc: show_error(exc)
        except Exception as exc: error_card(**explain_exception(exc), where="While checking setup")

def build_checks(prep: Prepared):
    s, cfg = prep.summary, prep.time_cfg
    checks = []
    src = f'{s["source"]} · {s["rows"]} rows · {s["cols"]} columns · {s["pending"]} not yet submitted'
    checks.append(("ok", "Data loaded", src))
    checks.append(("ok", "Rows to send", f'{s["selected"]} of {s["pending"]} pending rows, picked at random' if s["random"] else f'All {s["selected"]} pending rows, top to bottom'))
    checks.append(("ok", "Columns matched", f'{len(prep.mapping)} column(s) matched to {s["entry_fields"]} question(s)'))
    if prep.form.field_types: checks.append(("ok", "Form read", "Question types and choices read."))
    elif s["form_reachable"]: checks.append(("warn", "Form partly read", "Questions couldn't be fully read."))
    else: checks.append(("warn", "Couldn't open form", "Answers couldn't be checked against choices."))
    if prep.corrections: checks.append(("warn", "Small fixes made", f"{len(prep.corrections)} value(s) adjusted."))
    else: checks.append(("ok", "Small fixes", "None needed."))
    checks.append(("ok", "Answers checked", f'{s["selected"]} rows × {s["validated_fields"]} questions look good.'))
    sched = f'{fmt_clock(prep.rand_start)} → {fmt_clock(prep.rand_end)} IST · {format_human(cfg["window"] / max(cfg["total_rows"], 1))}'
    if __import__('datetime').datetime.now(IST) >= prep.rand_start: checks.append(("warn", "Schedule", sched + " · start time passed, sending begins now."))
    else: checks.append(("ok", "Schedule", sched))
    return checks

def view_preflight(prep: Prepared):
    section_head("Check before sending", "Everything passed. Nothing sent yet.")
    cfg = prep.time_cfg
    callout("ok", f"<b>Ready to send {cfg['total_rows']} row(s)</b> between <b>{fmt_clock(prep.rand_start)}</b> and <b>{fmt_clock(prep.rand_end)}</b> IST.")
    icons = {"ok": "✓", "warn": "!", "err": "✕"}
    rows = "".join(f'<div class="fi-check {k}"><span class="ic">{icons[k]}</span><div><div class="t">{esc(t)}</div><div class="d">{esc(d)}</div></div></div>' for k, t, d in build_checks(prep))
    md(f'<div class="fi-checks">{rows}</div>')
    t_map, t_fix, t_sched, t_console = st.tabs(["Column matching", "Auto-fixes", "Schedule", "Full log"])
    with t_map:
        html_table(["Your column", "Form question", "Question ID", "Type", "Required"], [[m["Column"], m["Question"], m["ID"], m["Type"], m["Required"]] for m in prep.mapping], ["", "", "mono", "mono", ""])
    with t_fix:
        if prep.corrections:
            st.caption("Values changed to match form exactly.")
            html_table(["Sheet row", "Column", "Was", "Now", "Why"], [[c["Row"], c["Column"], c["From"], c["To"], c["Why"]] for c in prep.corrections], ["mono", "", "val", "fix", ""])
        else: st.caption("No values needed changing.")
    with t_sched:
        es, ee = cfg["entered_start"], cfg["entered_end"]
        stats_grid(stat("You entered", fmt_clock(es), es.strftime("%d %b %Y"), small=True), stat("Until", fmt_clock(ee), ee.strftime("%d %b %Y"), small=True), stat("Actual start", fmt_clock(cfg["rand_start"]), f'+{cfg["rs"]}s random wait', "g", small=True), stat("Actual end", fmt_clock(cfg["rand_end"]), f'+{cfg["re"]}s random extra', "g", small=True))
        avg = cfg["window"] / max(cfg["total_rows"], 1)
        stats_grid(stat("Window entered", format_human(cfg["original_window"]), "", small=True), stat("Window used", format_human(cfg["window"]), "", "g", small=True), stat("Rows to send", cfg["total_rows"], f'needs at least {format_human(cfg["min_required"])}', small=True), stat("Average gap", format_human(avg), "between entries", small=True))
        if st.button("Pick new random times", help="Draws new random offsets"):
            try:
                reroll_schedule(prep)
                st.rerun()
            except AppError as exc: show_error(exc)
    with t_console: st.code(console_text(prep), language=None)
    md("<div style='height:6px'></div>")
    b1, b2 = st.columns([1, 2])
    if b1.button("← Back to set up", width="stretch"):
        st.session_state.prep = None
        st.rerun()
    if b2.button("Start sending →", type="primary", width="stretch"):
        job = Job(prep)
        job.start()
        st.session_state.job = job
        st.rerun()

def render_feed(rows):
    def row_html(r):
        k = r["kind"]
        if k in ("ok", "final"):
            s = f'<span class="r">Sheet row {r["row"]}</span> sent at <span class="m">{esc(r["time"])}</span><span class="pill">✓ sent</span>'
            if k == "ok": s += f' <span class="k">· next at</span> <span class="m">{esc(r["nxt"])}</span> <span class="k">(in {esc(r["gap"])})</span>'
            else: s += ' <span class="k">· last entry</span>'
            return f'<div class="fi-row {k}">{s}</div>'
        if k == "fail":
            return f'<div class="fi-row fail"><span class="r">Sheet row {r["row"]}</span> not sent at <span class="m">{esc(r["time"])}</span><span class="pill">✕ failed</span><span class="why">Why: {esc(r["problem"])}</span><span class="todo">What to do: {esc(r["fix"])}</span></div>'
        return f'<div class="fi-row {k}">{esc(r["text"])}</div>'
    
    if rows: body = "".join(row_html(r) for r in reversed(rows))
    else: body = '<div class="fi-row empty">No entries yet…</div>'
    md(f'<div class="fi-tbl-h"><span>Activity</span><span>{len(rows)} events</span></div><div class="fi-tbl">{body}</div>')

def render_failures(failures):
    html_table(["Sheet row", "Time (IST)", "What went wrong", "What to do"], [[f["row"], f["time"], f["problem"], f["fix"]] for f in failures], ["row", "mono", "", "fix"])

def render_monitor(job: Job):
    s = job.snapshot()
    label, tone, sub = STATUS_LABEL[s["status"]]
    total = max(s["total"], 1)
    ok_pct = s["submitted"] / total * 100
    bad_pct = s["failed"] / total * 100
    if s["next_target"] is not None and s["status"] in ("waiting", "running"):
        secs = max(0, int((s["next_target"] - __import__('datetime').datetime.now(IST)).total_seconds()))
        nxt_v, nxt_s = format_mm_ss(secs), f'at {fmt_clock(s["next_target"])} IST'
    else: nxt_v, nxt_s = "—", ""
    md('<div class="fi-stats">' + stat("Status", label, sub, tone) + stat("Sent", f'{s["submitted"]}<span style="color:#6B7480"> / {s["total"]}</span>', f"{int(ok_pct)}% done", "g", raw_value=True) + stat("Failed", s["failed"], "explained below" if s["failed"] else "none so far", "r" if s["failed"] else "") + stat("Next entry in", nxt_v, nxt_s) + f'</div><div class="fi-bar"><div class="ok" style="width:{ok_pct:.1f}%"></div><div class="bad" style="width:{bad_pct:.1f}%"></div></div>')
    if s["failures"]:
        callout("err", f"<b>{len(s['failures'])} row(s) couldn't be sent.</b> Failed rows stay pending.")
        render_failures(s["failures"])
    render_feed(s["rows"])
    if job.active:
        if st.button("Stop sending", key="stop_btn", help="Finishes current entry, then stops."): job.stop.set()
    elif not job.notified:
        job.notified = True
        st.rerun()

def view_execute(job: Job):
    section_head("Sending", "Entries being sent in background. Keep tab open.")
    monitor = st.fragment(run_every=1)(render_monitor)
    monitor(job)
    with st.expander("Check log", expanded=False): st.code(console_text(job.prep), language=None)

def view_results(job: Job):
    s = job.snapshot()
    section_head("Results", "What was sent, what wasn't, and why.")
    if s["status"] == "error":
        err = s["error"] or {}
        error_card(err.get("title", "Something went wrong"), "During sending", err.get("why", ""), err.get("fix", ""), technical=err.get("technical", ""))
    elif s["status"] == "stopped": callout("warn", "<b>You stopped the run.</b> Sent rows marked 'submitted'. Upload updated file to continue.")
    elif s["failed"]: callout("warn", f"<b>{s['submitted']} sent, {s['failed']} failed.</b> Failed rows stay pending.")
    else: callout("ok", f"<b>All {s['submitted']} row(s) sent successfully.</b>")
    stats_grid(stat("Sent", f'{s["submitted"]}<span style="color:#6B7480"> / {s["total"]}</span>', "", "g", raw_value=True), stat("Failed", s["failed"], "", "r" if s["failed"] else ""), stat("Status", STATUS_LABEL[s["status"]][0], STATUS_LABEL[s["status"]][2], STATUS_LABEL[s["status"]][1]), stat("Run time", format_human(job.elapsed()), "from first entry"))
    if s["failures"]:
        md('<div class="fi-tbl-h"><span>Rows not sent</span></div>')
        render_failures(s["failures"])
        with st.expander("Technical details for failed rows"): st.code("\n\n".join(f"Sheet row {f['row']}:\n{f['technical']}" for f in s["failures"]), language=None)
    render_feed(s["rows"])
    md("<div style='height:14px'></div>")
    cols = st.columns(4 if s["failures"] else 3)
    out = job.prep.output_path
    if os.path.exists(out):
        with open(out, "rb") as fh: cols[0].download_button("Download updated file", fh.read(), file_name=os.path.basename(out), width="stretch", help="Data with Status column filled.")
    else: cols[0].button("No updated file", disabled=True, width="stretch")
    cols[1].download_button("Download run log", job.log_text(), file_name="formintellect_log.txt", width="stretch")
    if s["failures"]: cols[2].download_button("Download failed rows", job.failures_csv(), file_name="formintellect_failed_rows.csv", width="stretch")
    if cols[-1].button("Start a new run", type="primary", width="stretch"):
        st.session_state.pop("job", None)
        st.session_state.pop("prep", None)
        st.rerun()
    if s["status"] == "done" and out.lower().endswith(".xlsx"): st.caption("Updated workbook includes 'Analysis' sheet.")
    with st.expander("Check log", expanded=False): st.code(console_text(job.prep), language=None)