"""Naphtali PM Group - Schedule Converter (Streamlit web app).

Upload any Excel schedule -> policy-compliant, verified MS Project / Primavera P6 files.
"""
from __future__ import annotations

import datetime as dt
import hashlib
import re

import pandas as pd
import plotly.graph_objects as go
import streamlit as st

from scheduler import exporters
from scheduler.engine import Options, apply_edited_predecessors, convert
from scheduler.model import load_policy

APP_NAME = "Schedule Converter"
POLICY = load_policy()

st.set_page_config(page_title=f"Naphtali PM Group | {APP_NAME}", page_icon="📅", layout="wide")

st.markdown(
    """
    <style>
      .block-container {padding-top: 1.6rem; padding-bottom: 5rem;}
      .npm-hero h1 {font-size: 2.0rem; margin-bottom: .1rem; color: #0B2545;}
      .npm-hero p {color: #5b6472; margin-top: 0;}
      .npm-badge {display:inline-block; padding:.15rem .55rem; border-radius:999px; font-size:.78rem;
                  background:#EEF2F8; color:#0B2545; margin-right:.35rem; border:1px solid #d9e1ee;}
      .npm-ok {border-left: 5px solid #1f8a4c; background:#eef8f2; padding:.8rem 1rem; border-radius:6px;}
      .npm-bad {border-left: 5px solid #b42318; background:#fdf0ef; padding:.8rem 1rem; border-radius:6px;}
      .npm-footer {position: fixed; left: 0; right: 0; bottom: 0; text-align: center; padding: .55rem;
                   font-size: .85rem; color: #E8ECF3; background: #0B2545; z-index: 999;}
      .npm-footer b {color: #C9A227;}
      @media (prefers-color-scheme: dark) {
        .npm-hero h1 {color: #E8ECF3;}
        .npm-ok {background:#12301f;} .npm-bad {background:#3a1513;}
        .npm-badge {background:#1b2a44; color:#E8ECF3; border-color:#2b3d5e;}
      }
    </style>
    """,
    unsafe_allow_html=True,
)


def footer():
    st.markdown('<div class="npm-footer">A <b>Naphtali PM Group</b> product. &copy; 2026</div>',
                unsafe_allow_html=True)


# --------------------------------------------------------------------------- #
# Header
# --------------------------------------------------------------------------- #
st.markdown(
    f"""<div class="npm-hero"><h1>📅 {APP_NAME}</h1>
    <p>Upload any Excel schedule and get a <b>policy-compliant, verified</b> file for
    <b>Microsoft Project</b> and <b>Primavera P6</b>, with zero date discrepancies.</p></div>""",
    unsafe_allow_html=True,
)

# --------------------------------------------------------------------------- #
# Sidebar: firm policy
# --------------------------------------------------------------------------- #
with st.sidebar:
    st.header("Firm scheduling policy")
    st.caption("Defaults come from firm_policy.json. Changes here apply to this session only.")
    st.markdown(
        "- Project title is **Outline Level 1**\n"
        "- Every summary ends with **'<summary> completed'** milestone, linked to its latest-finishing activity\n"
        "- Day-one activities are **SS-linked to the first activity**\n"
        "- Calendar named **'<project name> CALENDAR'**"
    )
    max_lead = st.number_input("Maximum lead (negative lag) allowed, weeks", 0.0, 10.0,
                               float(POLICY.get("max_lead_weeks", 2)), 0.5,
                               help="Beyond this, an SS + positive lag is used instead of FS with a lead.")
    close_ends = st.toggle("Close open ends (tie activities with no successor to their summary's completion "
                           "milestone)", value=bool(POLICY.get("close_open_ends", False)))
    proj_ms = st.toggle("Add '<project> completed' milestone under the project title",
                        value=bool(POLICY.get("completion_milestone_for_project_title", False)))
    use_src = st.toggle("Use predecessors found in the spreadsheet (if any)", value=True)
    dayfirst = st.radio("Text dates in the file are written", ["Day first (dd/mm/yyyy)", "Month first (mm/dd/yyyy)"],
                        index=0).startswith("Day")
    hol_text = st.text_area("Public holidays / non-working days (one date per line, yyyy-mm-dd)", "",
                            help="Added to the project calendar and respected by the verification.")
    st.divider()
    st.caption("Files are processed in memory and are not stored.")

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
up = st.file_uploader("Upload an Excel schedule", type=["xlsx", "xlsm", "xls", "csv"],
                      help="Any workbook with an activity list and dates, start weeks + durations, or a Gantt matrix.")

if up is None:
    c1, c2, c3 = st.columns(3)
    c1.markdown("**1 · Upload**  \nAny schedule workbook: WBS-coded, system-grouped, week-based or Gantt-only.")
    c2.markdown("**2 · Review**  \nThe app builds the hierarchy, applies firm policy, writes the logic and "
                "verifies every date against Excel.")
    c3.markdown("**3 · Download**  \nImport-ready CSV, MS Project XML (calendar included), P6 XER and a QA report.")
    footer()
    st.stop()

data = up.getvalue()
file_sig = hashlib.md5(data).hexdigest()

# first pass (auto-detect) to fill the controls
base_opts = Options(dayfirst=dayfirst, holidays=holidays, use_source_logic=use_src,
                    policy={"max_lead_weeks": max_lead, "close_open_ends": close_ends,
                            "completion_milestone_for_project_title": proj_ms})
try:
    probe = convert(data, up.name, base_opts)
except Exception as exc:  # noqa: BLE001
    st.error(f"Could not read a schedule from this file: {exc}")
    footer()
    st.stop()

with st.expander("Detected structure & project settings", expanded=True):
    c1, c2, c3 = st.columns([2, 2, 1.3])
    sheet = c1.selectbox("Schedule sheet", probe.sheets_found, index=probe.sheets_found.index(probe.sheet_name))
    modes = {"wbs": "WBS codes", "indent": "Indentation", "summary-rows": "Summary / heading rows",
             "group": "Group column (System / Phase / Area)", "flat": "Flat list"}
    mode = c2.selectbox("Hierarchy taken from", list(modes), format_func=modes.get,
                        index=list(modes).index(probe.info["mode"]))
    start = c3.date_input("Project start", probe.stats["excel_start"],
                          help="Change to re-baseline the whole schedule to a new start (e.g. revised NTP date).")
    if sheet != probe.sheet_name:
        probe = convert(data, up.name, Options(sheet_name=sheet, dayfirst=dayfirst, holidays=holidays,
                                               use_source_logic=use_src, policy=base_opts.policy))
    name = st.text_input("Project name (Outline Level 1 title)", probe.project.name)
    st.caption(f"Calendar: **{name.strip()}{POLICY.get('calendar_name_suffix', ' CALENDAR')}**  ·  "
               f"Columns used: " + ", ".join(f"{k} → col {chr(65 + v) if v < 26 else v + 1}"
                                                for k, v in probe.columns.items()))

opts = Options(sheet_name=sheet, project_name=name, dayfirst=dayfirst, holidays=holidays,
               use_source_logic=use_src, hierarchy_mode=mode,
               project_start=start if start != probe.stats["excel_start"] else None,
               policy=base_opts.policy)
res = convert(data, up.name, opts)

# planner edits (kept per file + settings)
edit_key = f"edits-{file_sig}-{sheet}-{mode}"
if st.session_state.get(edit_key):
    res = apply_edited_predecessors(res, st.session_state[edit_key])

stt = res.stats

# --------------------------------------------------------------------------- #
# Verification banner + metrics
# --------------------------------------------------------------------------- #
badges = "".join(f'<span class="npm-badge">{b}</span>' for b in [
    f"Sheet: {res.sheet_name}", f"Hierarchy: {modes.get(res.info['mode'], res.info['mode'])}",
    f"Logic: {res.info.get('logic_source')}", f"Header row {res.header_row}"])
st.markdown(badges, unsafe_allow_html=True)

if stt["date_mismatches"] == 0 and stt["finish_matches"]:
    st.markdown(
        f'<div class="npm-ok">✅ <b>Verified.</b> All {stt["activities"] + stt["milestones"]} activities and '
        f'milestones land exactly on the Excel dates when MS Project auto-schedules. Finish '
        f'<b>{stt["calc_finish"]:%a %d %b %Y}</b> ({stt["duration_days"]} working days).</div>',
        unsafe_allow_html=True)
else:
    st.markdown(
        f'<div class="npm-bad">⚠️ <b>{stt["date_mismatches"]} date difference(s)</b> between the logic and the '
        f'Excel dates. Calculated finish {stt["calc_finish"]:%d %b %Y} vs Excel '
        f'{stt["excel_finish"]:%d %b %Y}. See <b>Quality checks</b>.</div>', unsafe_allow_html=True)

st.write("")
m = st.columns(6)
m[0].metric("Activities", stt["activities"])
m[1].metric("Milestones", stt["milestones"] + stt["completion_milestones"])
m[2].metric("Summary tasks", stt["summaries"])
m[3].metric("Relationships", stt["links"])
m[4].metric("Critical activities", stt["critical"])
m[5].metric("Finish", f"{stt['output_finish']:%d-%b-%y}")

# --------------------------------------------------------------------------- #
# Downloads
# --------------------------------------------------------------------------- #
safe = re.sub(r"[^A-Za-z0-9]+", "_", res.project.name).strip("_")[:80] or "Schedule"
st.subheader("Download")
d = st.columns(4)
d[0].download_button("⬇️ Import-ready CSV", exporters.to_csv(res), f"{safe}_Import_Ready.csv", "text/csv",
                     type="primary", width="stretch")
d[1].download_button("⬇️ MS Project XML (with calendar)", exporters.to_mspdi(res), f"{safe}.xml",
                     "application/xml", width="stretch",
                     help="Opens directly in MS Project (File > Open > XML). Calendar, start date, auto-scheduling "
                          "and logic are already set. Also imports into Primavera P6.")


@st.cache_resource(show_spinner=False)
def _xer_ok() -> bool:
    return exporters.xer_available()


if _xer_ok():
    try:
        d[2].download_button("⬇️ Primavera P6 XER", exporters.to_xer(res), f"{safe}.xer",
                             "application/octet-stream", width="stretch")
    except Exception as exc:  # noqa: BLE001
        d[2].caption(f"P6 XER unavailable: {exc}")
else:
    d[2].caption("P6: import the MS Project XML (File > Import > Microsoft Project).")
d[3].download_button("⬇️ QA report (Excel)", exporters.to_qa_xlsx(res), f"{safe}_QA_Report.xlsx",
                     "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", width="stretch")

# --------------------------------------------------------------------------- #
# Tabs
# --------------------------------------------------------------------------- #
t_sched, t_gantt, t_qa, t_how = st.tabs(["Schedule & logic", "Gantt preview", "Quality checks", "How to import"])

table = exporters.schedule_table(res)

with t_sched:
    st.caption("Edit the **Predecessors** column if you want different logic (e.g. `12SS+2 wks, 8FS`), "
               "then press **Re-verify**. Every change is re-checked against the Excel dates.")
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
    b1, b2, _ = st.columns([1, 1, 4])
    if b1.button("Re-verify with my edits", type="primary"):
        changes = {int(i): p for i, p, o in zip(edited["ID"], edited["Predecessors"], show["Predecessors"])
                   if (p or "") != (o or "")}
        prev = dict(st.session_state.get(edit_key, {}))
        prev.update(changes)
        st.session_state[edit_key] = prev
        st.rerun()
    if b2.button("Reset to automatic logic"):
        st.session_state.pop(edit_key, None)
        st.rerun()

with t_gantt:
    g = table[table["Start Date"].notna()].copy()
    fig = go.Figure()
    colors = {"project": "#0B2545", "summary": "#13315C", "task": "#3E7CB1", "milestone": "#C9A227",
              "completion": "#1f8a4c"}
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
            col = "#b42318" if (crit and r["_kind"] == "task") else colors[r["_kind"]]
            fig.add_trace(go.Bar(x=[(f - s).total_seconds() * 1000], base=[s], y=[label], orientation="h",
                                 width=thick, marker_color=col, showlegend=False,
                                 hovertext=f"{r['Task Name']}<br>{r['Start Date']:%d-%b-%y} → "
                                           f"{r['Finish Date']:%d-%b-%y}<br>{r['Duration']} · {r['Predecessors']}",
                                 hoverinfo="text"))
    fig.update_layout(height=max(420, 22 * len(g) + 80), barmode="overlay", margin=dict(l=10, r=10, t=30, b=10),
                      xaxis=dict(type="date", side="top", gridcolor="rgba(128,128,128,.2)"),
                      yaxis=dict(autorange="reversed", tickfont=dict(size=11)))
    st.caption("Red = critical (zero float) · Gold ◆ = Excel milestone · Green ◆ = policy completion milestone")
    st.plotly_chart(fig, width="stretch")

with t_qa:
    q = pd.DataFrame(res.issues, columns=["Severity", "ID", "Task", "Message"])
    q["ID"] = q["ID"].astype(str)
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Date differences vs Excel", stt["date_mismatches"])
    c2.metric("Open starts / open ends", f"{stt['open_starts']} / {stt['open_ends']}")
    c3.metric("Leads / lags", f"{stt['leads']} / {stt['lags']}")
    c4.metric("Relationships", stt["links"])
    c4.caption(" · ".join(f"{k} {v}" for k, v in sorted(stt["rel_mix"].items())))
    if q.empty:
        st.success("No issues found.")
    else:
        order = {"error": 0, "warning": 1, "info": 2}
        q = q.sort_values("Severity", key=lambda s: s.map(order))
        st.dataframe(q, hide_index=True, width="stretch")
    if res.notes:
        st.info("\n".join(res.notes))

with t_how:
    st.markdown(f"""
**Option A – MS Project XML (recommended, fewest steps)**
1. In MS Project: **File › Open**, set the file type to **XML Format (*.xml)** and pick the downloaded `.xml`.
2. Choose **As a new project**. Everything is already set: project start, the calendar
   **{res.calendar_name}**, auto-scheduling, outline, WBS and logic.
3. Save as `.mpp`.

**Option B – Import-ready CSV (as before)**
1. **File › Open** › file type **Text (*.csv)** › select the CSV › **New map** › **As a new project** ›
   tick **Tasks** and **Import includes headers**.
2. Map: Task Name → Name, Duration → Duration, Start Date → Start, Finish Date → Finish,
   Predecessors → Predecessors, Outline Level → Outline Level, WBS → WBS. Save the map for next time.
3. **Project › Project Information**: Start date **{stt['output_start']:%d %b %Y}**.
4. **Project › Change Working Time › Create New Calendar**: **{res.calendar_name}**, then select it in
   Project Information.
5. Select all (Ctrl + A) › **Task › Auto Schedule**. Finish should read **{stt['output_finish']:%a %d %b %Y}**.

**Primavera P6**: import the `.xer` (File › Import › Primavera PM – (XER)) or the MS Project `.xml`
(File › Import › Microsoft Project). Press F9 to schedule.
""")

# --------------------------------------------------------------------------- #
# Optional AI review (only if an Anthropic API key is configured in Streamlit secrets)
# --------------------------------------------------------------------------- #
api_key = None
try:
    api_key = st.secrets.get("ANTHROPIC_API_KEY")
except Exception:  # noqa: BLE001 - no secrets file
    api_key = None

if api_key:
    with st.expander("🤖 AI planning review (optional)"):
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
