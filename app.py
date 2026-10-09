"""Naphtali PM Group - Schedule Converter (Streamlit web app).

Upload any Excel schedule -> policy-compliant, verified MS Project / Primavera P6 files.
"""
from __future__ import annotations

import datetime as dt
import hashlib
import html
import re

import pandas as pd
import plotly.graph_objects as go
import streamlit as st

from scheduler import exporters
from scheduler.engine import Options, apply_edited_predecessors, calendar_name_for, convert
from scheduler.model import load_policy

APP_NAME = "Schedule Converter"
POLICY = load_policy()

st.set_page_config(page_title=f"{APP_NAME} | Naphtali PM Group", page_icon=":material/view_timeline:",
                   layout="wide", initial_sidebar_state="expanded")

# dark-blue palette
BG, SURF, SURF2, LINE = "#0B1626", "#11213A", "#16304F", "#23405F"
TEXT, MUTED, BLUE, GOLD, GREEN, RED = "#E6EDF5", "#8FA3BA", "#4C9AFF", "#D4AF37", "#3FB97A", "#F0645A"
NAVY, INK = SURF, TEXT          # used by the Gantt chart colours

st.markdown(
    f"""
    <style>
      @import url('https://fonts.googleapis.com/css2?family=Archivo:wght@600;700&display=swap');
      header[data-testid="stHeader"] {{background: transparent;}}
      .block-container {{padding-top: 1.1rem; padding-bottom: 2rem; max-width: 1240px;}}
      h1, h2, h3, h4 {{letter-spacing: -0.01em;}}

      .sc-band {{background: linear-gradient(0deg, {SURF}, {SURF}); border: 1px solid {LINE}; border-radius: 14px;
                 padding: 1.4rem 1.8rem; display: grid; grid-template-columns: minmax(0, 1fr) 360px; gap: 1.5rem;
                 align-items: center; margin-bottom: 1.2rem;}}
      .sc-band h1 {{font-family: Archivo, sans-serif; font-weight: 700; font-size: 2.3rem; line-height: 1.1;
                    color: #FFFFFF; margin: 0 0 .35rem; padding: 0;}}
      .sc-band p {{color: {MUTED}; font-size: 1.02rem; margin: 0;}}
      .sc-gantt {{width: 100%; height: auto; display: block;}}
      .sc-gantt .bar {{transform-box: fill-box; transform-origin: left center; animation: grow .7s ease-out both;}}
      .sc-gantt .ms {{transform-box: fill-box; transform-origin: center; animation: pop .45s ease-out both;
                     animation-delay: 1.15s;}}
      @keyframes grow {{from {{transform: scaleX(0);}} to {{transform: scaleX(1);}}}}
      @keyframes pop {{from {{transform: scale(0) rotate(45deg); opacity: 0;}} to {{transform: scale(1) rotate(45deg); opacity: 1;}}}}
      @media (prefers-reduced-motion: reduce) {{ .sc-gantt .bar, .sc-gantt .ms {{animation: none;}} }}
      @media (max-width: 760px) {{ .sc-band {{grid-template-columns: 1fr;}} .sc-band h1 {{font-size: 1.8rem;}} }}

      [data-testid="stFileUploaderDropzone"] {{background: {SURF}; border: 1.5px dashed #3A5F8A;
                                               padding: 1.8rem 1.4rem; border-radius: 14px;}}
      [data-testid="stFileUploaderDropzone"]:hover {{border-color: {BLUE};}}

      .sc-steps {{display: grid; grid-template-columns: repeat(3, minmax(0, 1fr)); gap: 1rem; margin-top: 1.2rem;}}
      .sc-step {{display: flex; align-items: center; gap: .8rem; padding: .9rem 1rem; border-radius: 12px;
                 background: {SURF}; border: 1px solid {LINE}; color: {TEXT}; font-weight: 600;}}
      .sc-step .n {{flex: none; width: 30px; height: 30px; border-radius: 50%; display: grid; place-items: center;
                    background: {SURF2}; color: {BLUE}; font-family: Archivo, sans-serif; font-weight: 700;}}
      @media (max-width: 760px) {{ .sc-steps {{grid-template-columns: 1fr;}} }}

      .sc-proj {{display: flex; flex-wrap: wrap; align-items: baseline; gap: .4rem 1rem; margin: .3rem 0 .9rem;}}
      .sc-proj h2 {{font-family: Archivo, sans-serif; font-size: 1.5rem; color: #FFFFFF; margin: 0; padding: 0;}}
      .sc-chip {{display: inline-block; background: {SURF2}; color: {MUTED}; border-radius: 999px;
                 padding: .12rem .6rem; font-size: .8rem; margin-right: .3rem;}}

      .sc-verdict {{display: grid; grid-template-columns: auto 1fr auto; gap: 1.1rem; align-items: center;
                    background: {SURF}; border: 1px solid {LINE}; border-radius: 14px; padding: 1rem 1.3rem;
                    margin-bottom: 1rem;}}
      .sc-verdict .seal {{width: 48px; height: 48px; border-radius: 50%; display: grid; place-items: center;
                          background: rgba(63,185,122,.14); color: {GREEN};}}
      .sc-verdict.bad .seal {{background: rgba(240,100,90,.14); color: {RED};}}
      .sc-verdict h3 {{font-family: Archivo, sans-serif; margin: 0; font-size: 1.2rem; color: #FFFFFF; padding: 0;}}
      .sc-verdict .sub {{color: {MUTED}; font-size: .9rem; margin-top: .15rem;}}
      .sc-verdict .finish {{text-align: right;}}
      .sc-verdict .finish b {{display: block; font-family: Archivo, sans-serif; font-size: 1.5rem; color: {BLUE};}}
      .sc-verdict .finish span {{color: {MUTED}; font-size: .82rem;}}
      @media (max-width: 760px) {{ .sc-verdict {{grid-template-columns: auto 1fr;}}
                                   .sc-verdict .finish {{grid-column: 1 / -1; text-align: left;}} }}

      .sc-figs {{display: grid; grid-template-columns: repeat(5, minmax(0, 1fr)); gap: .7rem; margin-bottom: 1.4rem;}}
      .sc-fig {{background: {SURF}; border: 1px solid {LINE}; border-radius: 12px; padding: .7rem 1rem;}}
      .sc-fig b {{display: block; font-family: Archivo, sans-serif; font-size: 1.55rem; color: #FFFFFF; line-height: 1.2;}}
      .sc-fig span {{color: {MUTED}; font-size: .82rem;}}
      @media (max-width: 760px) {{ .sc-figs {{grid-template-columns: repeat(2, minmax(0, 1fr));}} }}

      [data-testid="stDownloadButton"] button {{font-weight: 600; padding: .6rem .9rem; border-radius: 10px;}}
      [data-testid="stTabs"] [role="tab"] {{font-weight: 600;}}

      [data-testid="stSidebar"] h2 {{font-family: Archivo, sans-serif; font-size: 1.1rem; color: #FFFFFF;}}
      .sc-rules {{display: flex; flex-wrap: wrap; gap: .4rem; margin: .1rem 0 1rem;}}
      .sc-rules span {{background: {SURF}; border: 1px solid #1C3452; color: {TEXT}; border-radius: 999px;
                       padding: .2rem .65rem; font-size: .8rem;}}

      .sc-footer {{margin-top: 2.5rem; padding: .9rem 0 .2rem; border-top: 1px solid {LINE}; text-align: center;
                   color: {MUTED}; font-size: .86rem;}}
      .sc-footer b {{color: {TEXT};}}
      .sc-footer i {{display: inline-block; width: 7px; height: 7px; background: {GOLD}; transform: rotate(45deg);
                     margin: 0 .55rem .1rem;}}
    </style>
    """,
    unsafe_allow_html=True,
)


def footer():
    st.markdown('<div class="sc-footer"><i></i>A <b>Naphtali PM Group</b> product. &copy; 2026<i></i></div>',
                unsafe_allow_html=True)


GANTT_SVG = f"""
<svg class="sc-gantt" viewBox="0 0 420 150" role="img" aria-label="Gantt bars ending in a milestone">
  <g stroke="{LINE}" stroke-width="1">
    <line x1="20" y1="6" x2="20" y2="146"/><line x1="120" y1="6" x2="120" y2="146"/>
    <line x1="220" y1="6" x2="220" y2="146"/><line x1="320" y1="6" x2="320" y2="146"/>
  </g>
  <g fill="none" stroke="#3A5F8A" stroke-width="1.4">
    <path d="M118 22 h10 v22 h4"/><path d="M196 50 h8 v22 h4"/><path d="M262 78 h8 v22 h4"/>
    <path d="M340 106 h8 v14 h8"/>
  </g>
  <rect class="bar" x="24" y="14" width="94" height="16" rx="4" fill="{BLUE}" style="animation-delay:.05s"/>
  <rect class="bar" x="132" y="42" width="64" height="16" rx="4" fill="{BLUE}" style="animation-delay:.25s"/>
  <rect class="bar" x="208" y="70" width="54" height="16" rx="4" fill="#2E5E99" style="animation-delay:.45s"/>
  <rect class="bar" x="274" y="98" width="66" height="16" rx="4" fill="#2E5E99" style="animation-delay:.65s"/>
  <rect class="ms" x="356" y="112" width="15" height="15" fill="{GOLD}" transform="rotate(45 363 120)"/>
</svg>"""

st.markdown(
    f"""<div class="sc-band">
      <div><h1>{APP_NAME}</h1><p>Excel schedule in. Verified MS Project &amp; P6 files out.</p></div>
      <div>{GANTT_SVG}</div>
    </div>""",
    unsafe_allow_html=True,
)

# --------------------------------------------------------------------------- #
# Sidebar: firm policy
# --------------------------------------------------------------------------- #
with st.sidebar:
    st.header("Firm policy")
    st.markdown('<div class="sc-rules"><span>Title at level 1</span><span>"… completed" milestones</span>'
                '<span>Day-one SS links</span><span>Named calendar</span></div>', unsafe_allow_html=True)
    with st.expander("Logic", icon=":material/account_tree:"):
        max_lead = st.number_input("Max lead (weeks)", 0.0, 10.0, float(POLICY.get("max_lead_weeks", 2)), 0.5,
                                   help="Longer overlaps use SS + lag instead of FS with a lead.")
        close_ends = st.toggle("Close open ends", value=bool(POLICY.get("close_open_ends", False)),
                               help="Tie activities with no successor to their summary's completed milestone.")
        proj_ms = st.toggle("Project completed milestone",
                            value=bool(POLICY.get("completion_milestone_for_project_title", False)),
                            help="Add '<project> completed' under the project title.")
        use_src = st.toggle("Use workbook predecessors", value=True,
                            help="If the workbook already has a Predecessors column, keep its logic.")
    with st.expander("Calendar", icon=":material/calendar_month:"):
        dayfirst = st.radio("Text dates", ["dd/mm/yyyy", "mm/dd/yyyy"], index=0, horizontal=True) == "dd/mm/yyyy"
        hol_text = st.text_area("Holidays (yyyy-mm-dd, one per line)", "", height=110)
    st.caption("Files are never stored.")

holidays = []
for line in hol_text.splitlines():
    line = line.strip()
    if line:
        try:
            holidays.append(dt.date.fromisoformat(line))
        except ValueError:
            st.sidebar.warning(f"Ignored holiday '{line}' (use yyyy-mm-dd)")

# --------------------------------------------------------------------------- #
# Upload
# --------------------------------------------------------------------------- #
up = st.file_uploader("Schedule workbook", type=["xlsx", "xlsm", "xls", "csv"], label_visibility="collapsed",
                      help="Any workbook with activities and dates, start weeks + durations, or a drawn Gantt.")

if up is None:
    st.markdown('<div class="sc-steps"><div class="sc-step"><span class="n">1</span>Upload the workbook</div>'
                '<div class="sc-step"><span class="n">2</span>Check the dates</div>'
                '<div class="sc-step"><span class="n">3</span>Download your files</div></div>',
                unsafe_allow_html=True)
    footer()
    st.stop()

data = up.getvalue()
file_sig = hashlib.md5(data).hexdigest()

base_opts = Options(dayfirst=dayfirst, holidays=holidays, use_source_logic=use_src,
                    policy={"max_lead_weeks": max_lead, "close_open_ends": close_ends,
                            "completion_milestone_for_project_title": proj_ms})
try:
    with st.spinner("Reading the schedule…"):
        probe = convert(data, up.name, base_opts)
except Exception as exc:  # noqa: BLE001
    st.error(f"No schedule could be read from {up.name}. {exc}")
    footer()
    st.stop()

modes = {"wbs": "WBS codes", "indent": "Indentation", "summary-rows": "Heading rows",
         "group": "Group column", "flat": "Flat list"}

with st.expander("Project settings", icon=":material/tune:"):
    c1, c2, c3 = st.columns([2, 1.4, 1.2])
    sheet = c1.selectbox("Sheet", probe.sheets_found, index=probe.sheets_found.index(probe.sheet_name))
    mode = c2.selectbox("Outline from", list(modes), format_func=modes.get,
                        index=list(modes).index(probe.info["mode"]))
    start = c3.date_input("Start", probe.stats["excel_start"], format="DD/MM/YYYY",
                          help="Change to move the whole schedule to a new start date.")
    if sheet != probe.sheet_name:
        probe = convert(data, up.name, Options(sheet_name=sheet, dayfirst=dayfirst, holidays=holidays,
                                               use_source_logic=use_src, policy=base_opts.policy))
    name = st.text_input("Project name", probe.project.name)
    auto_cal = calendar_name_for(name, {**POLICY, **base_opts.policy})
    cal_name = st.text_input("Calendar name", auto_cal, max_chars=51, key=f"cal-{auto_cal}",
                             help="MS Project allows at most 51 characters. Long project names are shortened.")

opts = Options(sheet_name=sheet, project_name=name, calendar_name=cal_name, dayfirst=dayfirst, holidays=holidays,
               use_source_logic=use_src, hierarchy_mode=mode,
               project_start=start if start != probe.stats["excel_start"] else None,
               policy=base_opts.policy)
res = convert(data, up.name, opts)

edit_key = f"edits-{file_sig}-{sheet}-{mode}"
if st.session_state.get(edit_key):
    res = apply_edited_predecessors(res, st.session_state[edit_key])

stt = res.stats

# --------------------------------------------------------------------------- #
# Result
# --------------------------------------------------------------------------- #
esc = html.escape
chips = "".join(f'<span class="sc-chip">{esc(c)}</span>' for c in [res.sheet_name, modes.get(res.info["mode"], "")])
st.markdown(f'<div class="sc-proj"><h2>{esc(res.project.name)}</h2><div>{chips}</div></div>',
            unsafe_allow_html=True)

CHECK = ('<svg width="26" height="26" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.6" '
         'stroke-linecap="round" stroke-linejoin="round"><path d="M4 12.5l5 5L20 6.5"/></svg>')
WARN = ('<svg width="26" height="26" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.6" '
        'stroke-linecap="round"><path d="M12 6v8"/><path d="M12 18h.01"/></svg>')
n_items = stt["activities"] + stt["milestones"]
ok = stt["date_mismatches"] == 0 and stt["finish_matches"]
if ok:
    head, sub = "All dates match the Excel", f"{n_items} of {n_items} activities verified"
else:
    head, sub = (f"{stt['date_mismatches']} date{'s' if stt['date_mismatches'] != 1 else ''} differ from the Excel",
                 "See Quality checks")
st.markdown(
    f"""<div class="sc-verdict {'ok' if ok else 'bad'}">
      <div class="seal">{CHECK if ok else WARN}</div>
      <div><h3>{head}</h3><div class="sub">{sub}</div></div>
      <div class="finish"><span>Finish</span><b>{stt['output_finish']:%d %b %Y}</b>
        <span>{stt['duration_days']} working days</span></div>
    </div>""", unsafe_allow_html=True)

figs = [(stt["activities"], "Activities"), (stt["milestones"] + stt["completion_milestones"], "Milestones"),
        (stt["summaries"], "Summaries"), (stt["links"], "Links"), (stt["critical"], "Critical")]
st.markdown('<div class="sc-figs">' + "".join(f'<div class="sc-fig"><b>{v}</b><span>{l}</span></div>'
                                             for v, l in figs) + "</div>", unsafe_allow_html=True)

# --------------------------------------------------------------------------- #
# Downloads
# --------------------------------------------------------------------------- #
safe = re.sub(r"[^A-Za-z0-9]+", "_", res.project.name).strip("_")[:80] or "Schedule"


@st.cache_resource(show_spinner=False)
def _xer_ok() -> bool:
    return exporters.xer_available()


d = st.columns(4, gap="small")
d[0].download_button("MS Project XML", exporters.to_mspdi(res), f"{safe}.xml", "application/xml",
                     type="primary", icon=":material/download:", width="stretch",
                     help="Recommended. Opens in MS Project with calendar, start date and logic already set.")
d[1].download_button("Import CSV", exporters.to_csv(res), f"{safe}_Import_Ready.csv", "text/csv",
                     icon=":material/download:", width="stretch",
                     help="For the MS Project import wizard (usual field mapping).")
if _xer_ok():
    try:
        d[2].download_button("Primavera P6", exporters.to_xer(res), f"{safe}.xer", "application/octet-stream",
                             icon=":material/download:", width="stretch", help="Native .xer file.")
    except Exception:  # noqa: BLE001
        d[2].caption("P6: import the XML")
else:
    d[2].caption("P6: import the XML")
d[3].download_button("QA report", exporters.to_qa_xlsx(res), f"{safe}_QA_Report.xlsx",
                     "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                     icon=":material/download:", width="stretch",
                     help="Excel: date check, logic health, trace to Excel rows.")
st.write("")

# --------------------------------------------------------------------------- #
# Tabs
# --------------------------------------------------------------------------- #
t_sched, t_gantt, t_qa, t_how = st.tabs([":material/table_rows: Schedule & logic", ":material/view_timeline: Gantt",
                                         ":material/fact_check: Quality checks", ":material/help: How to open"])

table = exporters.schedule_table(res)

with t_sched:
    st.caption("Edit **Predecessors**, then **Re-check dates**.")
    show = table[["ID", "WBS", "Task Name", "Duration", "Start Date", "Finish Date", "Predecessors",
                  "Outline Level", "_float_days"]].rename(columns={"_float_days": "Float (days)"})
    show["Task Name"] = [("    " * (lvl - 1)) + n for n, lvl in zip(show["Task Name"], show["Outline Level"])]
    edited = st.data_editor(
        show, hide_index=True, width="stretch", height=560, key=f"ed-{edit_key}",
        disabled=[c for c in show.columns if c != "Predecessors"],
        column_config={"Start Date": st.column_config.DateColumn(format="DD-MMM-YY"),
                       "Finish Date": st.column_config.DateColumn(format="DD-MMM-YY"),
                       "Task Name": st.column_config.TextColumn(width="large")},
    )
    b1, b2, _ = st.columns([1, 1.2, 3])
    if b1.button("Re-check dates", type="primary", icon=":material/rule:"):
        changes = {int(i): p for i, p, o in zip(edited["ID"], edited["Predecessors"], show["Predecessors"])
                   if (p or "") != (o or "")}
        prev = dict(st.session_state.get(edit_key, {}))
        prev.update(changes)
        st.session_state[edit_key] = prev
        st.rerun()
    if b2.button("Restore automatic logic", icon=":material/restart_alt:"):
        st.session_state.pop(edit_key, None)
        st.rerun()

with t_gantt:
    g = table[table["Start Date"].notna()].copy()
    fig = go.Figure()
    colors = {"project": "#C9D6E6", "summary": MUTED, "task": BLUE, "milestone": GOLD, "completion": GREEN}
    y = [f"{r.ID}  {r['Task Name'][:60]}" for _, r in g.iterrows()]
    for (_, r), label in zip(g.iterrows(), y):
        crit = r["_float_days"] is not None and r["_float_days"] == r["_float_days"] and r["_float_days"] <= 0
        if r["_kind"] in ("milestone", "completion"):
            fig.add_trace(go.Scatter(x=[pd.Timestamp(r["Finish Date"]) + pd.Timedelta(hours=17)], y=[label],
                                     mode="markers", marker_symbol="diamond", marker_size=11,
                                     marker_color=colors[r["_kind"]], showlegend=False,
                                     hovertext=f"{r['Task Name']}<br>{r['Finish Date']:%d-%b-%y}", hoverinfo="text"))
        else:
            s = pd.Timestamp(r["Start Date"])
            f = pd.Timestamp(r["Finish Date"]) + pd.Timedelta(days=1)
            thick = 0.35 if r["_kind"] in ("summary", "project") else 0.6
            col = RED if (crit and r["_kind"] == "task") else colors[r["_kind"]]
            fig.add_trace(go.Bar(x=[(f - s).total_seconds() * 1000], base=[s], y=[label], orientation="h",
                                 width=thick, marker_color=col, showlegend=False,
                                 hovertext=f"{r['Task Name']}<br>{r['Start Date']:%d-%b-%y} → "
                                           f"{r['Finish Date']:%d-%b-%y}<br>{r['Duration']} · {r['Predecessors']}",
                                 hoverinfo="text"))
    fig.update_layout(height=max(420, 24 * len(g) + 80), barmode="overlay", margin=dict(l=10, r=10, t=30, b=10),
                      paper_bgcolor=SURF, plot_bgcolor=SURF,
                      font=dict(family="Public Sans, sans-serif", color=TEXT),
                      hoverlabel=dict(bgcolor=SURF2, font_size=12),
                      xaxis=dict(type="date", side="top", gridcolor=LINE, linecolor=LINE),
                      yaxis=dict(autorange="reversed", tickfont=dict(size=11), gridcolor="#152A44"))
    st.markdown(f'<div style="color:{MUTED};font-size:.88rem;margin:.4rem 0">'
                f'<span style="color:{RED}">■</span> critical &nbsp;&nbsp; '
                f'<span style="color:{BLUE}">■</span> activity &nbsp;&nbsp; '
                f'<span style="color:{GOLD}">◆</span> Excel milestone &nbsp;&nbsp; '
                f'<span style="color:{GREEN}">◆</span> completed milestone</div>', unsafe_allow_html=True)
    st.plotly_chart(fig, width="stretch")

with t_qa:
    q = pd.DataFrame(res.issues, columns=["Severity", "ID", "Task", "Message"])
    q["ID"] = q["ID"].astype(str)
    rel = ", ".join(f"{v} {k}" for k, v in sorted(stt["rel_mix"].items()))
    qa_figs = [(stt["date_mismatches"], "Date differences"),
               (f"{stt['open_starts']} / {stt['open_ends']}", "Open starts / ends"),
               (f"{stt['leads']} / {stt['lags']}", "Leads / lags"), (rel, "Link types")]
    st.markdown('<div class="sc-figs" style="grid-template-columns:repeat(4,minmax(0,1fr));margin-top:.6rem">'
                + "".join(f'<div class="sc-fig"><b>{v}</b><span>{l}</span></div>' for v, l in qa_figs)
                + "</div>", unsafe_allow_html=True)
    main = q[q["Severity"] != "info"]
    notes = q[q["Severity"] == "info"]
    if main.empty:
        st.success("No errors or warnings.", icon=":material/check_circle:")
    else:
        main = main.sort_values("Severity", key=lambda x: x.map({"error": 0, "warning": 1}))
        st.dataframe(main, hide_index=True, width="stretch")
    if not notes.empty:
        with st.expander(f"Notes ({len(notes)})"):
            st.dataframe(notes[["ID", "Task", "Message"]], hide_index=True, width="stretch")
    if res.notes:
        st.info("\n".join(res.notes))

with t_how:
    h1, h2 = st.columns(2, gap="large")
    h1.markdown(f"""
#### MS Project XML
1. **File › Open** › type **XML Format** › pick the file
2. **As a new project** › Save as `.mpp`

Calendar *{res.calendar_name}* and start date are already set.
""")
    h2.markdown(f"""
#### Import CSV
1. **File › Open** › type **CSV** › **New map** › **Tasks**, headers on
2. Map the 7 columns to the same-named MS Project fields
3. Set start **{stt['output_start']:%d %b %Y}** and the calendar, then **Auto Schedule** all
""")
    st.caption("Primavera P6: File › Import › XER (or Microsoft Project XML), then F9.")

# --------------------------------------------------------------------------- #
# Optional AI review (only if an Anthropic API key is configured in Streamlit secrets)
# --------------------------------------------------------------------------- #
api_key = None
try:
    api_key = st.secrets.get("ANTHROPIC_API_KEY")
except Exception:  # noqa: BLE001 - no secrets file
    api_key = None

if api_key:
    with st.expander("AI planning review", icon=":material/psychology:"):
        st.caption("Sends the converted task list (names, dates, logic) to Claude for a planner's critique.")
        if st.button("Run AI review"):
            try:
                import anthropic

                client = anthropic.Anthropic(api_key=api_key)
                csv_text = exporters.to_csv(res).decode("utf-8", "ignore")
                prompt = (
                    "You are a senior planning engineer (PMI-SP) reviewing a Level 2 schedule converted from "
                    "Excel for MS Project. Dates are fixed by the client; logic was inferred to reproduce them. "
                    "In under 300 words, flag logic that looks unrealistic, missing hand-overs, excessive "
                    "leads/lags, and DCMA-14 concerns, then list the top 5 improvements. Be concise.\n\n"
                    f"QA stats: {stt}\n\nSchedule CSV:\n{csv_text}")
                msg = client.messages.create(model=st.secrets.get("ANTHROPIC_MODEL", "claude-sonnet-5-5"),
                                             max_tokens=1200, messages=[{"role": "user", "content": prompt}])
                st.markdown("".join(b.text for b in msg.content if getattr(b, "type", "") == "text"))
            except Exception as exc:  # noqa: BLE001
                st.error(f"AI review failed: {exc}")

footer()
