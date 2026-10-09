# Schedule Converter — Naphtali PM Group

Upload **any Excel schedule** and get files that open in **Microsoft Project** and **Primavera P6**
with the firm's scheduling policy applied and **every date verified against the Excel**.

Live app: `https://schedule-converter.streamlit.app` (Streamlit Community Cloud)

---

## What it does

| Step | What happens |
|---|---|
| Read | Finds the schedule sheet and header row automatically. Works with flexible column names (Activity / Task Name / Description, Start / Planned Start, Dur / Duration / Dur (wks), Start Wk, WBS, Type, System / Phase / Area, Predecessors). Supports `.xlsx`, `.xlsm`, `.xls`, `.csv`. |
| Structure | Builds the outline from WBS codes, indentation, summary/heading rows, or a grouping column (System / Phase / Area). Week-number schedules and Gantt-only matrices (bars drawn as filled cells) are converted to dates. Legend and note rows are ignored and listed. |
| Firm policy | 1. Project title is **Outline Level 1**. 2. Every summary ends with **"<summary> completed"** (0 d milestone) linked FS to its **latest-finishing** activity. 3. Day-one activities are **SS-linked to the first activity** (anchor). 4. Calendar named **"<project name> CALENDAR"**. WBS codes are regenerated (1, 1.1, 1.1.1 …). |
| Logic | Uses predecessors found in the sheet; otherwise writes logic from the Excel dates: zero-lag FS first, then SS/FF; leads only up to a set limit (default 2 wks); links stay inside their system/section unless it is a true hand-over; decision gates (e.g. DG1, DG2) are linked to every matching activity finishing on the gate date. No open starts. |
| Verify | A CPM engine that schedules the way MS Project does (forward/backward pass, standard calendar plus holidays) recalculates the schedule and compares every activity with the Excel. Any date difference is reported, so the file is checked before it reaches MS Project. |
| Output | Import-ready **CSV** (same 8 columns as before) · **MS Project XML** (opens directly, with calendar, start date, auto-scheduling) · **Primavera P6 XER** · **QA report** (Excel: summary, issues, traceability to Excel rows, float). |

Extras in the app: a sidebar for policy settings (max lead, close open ends, holidays, date format), a
re-baseline option (move the whole schedule to a new start/NTP date), predecessors you can edit and
re-verify, a Gantt preview with the critical path, and an optional AI planning review.

## Repository layout

```
app.py                  Streamlit user interface
firm_policy.json        Firm scheduling policy (edit to change defaults)
scheduler/
  reader.py             Sheet / header / column detection, dates, week & Gantt parsing
  builder.py            Outline, completion milestones, WBS
  logic.py              Predecessor inference + parsing of existing predecessors
  cpm.py                MS Project-equivalent forward/backward pass
  engine.py             Orchestration + quality checks
  exporters.py          CSV, MS Project XML, P6 XER, QA workbook
tests/test_engine.py    Regression tests on synthetic schedules (pytest)
requirements.txt        Python packages (pinned)
packages.txt            Java runtime for the P6 XER export (Streamlit Cloud installs it)
.streamlit/config.toml  Theme
```

## Update the GitHub repository (Temmydesign/schedule-converter)

The simplest way, in the browser:

1. Open the repository on GitHub, then **Add file › Upload files**.
2. Drag in **everything inside this folder**, including the `scheduler`, `tests` and `.streamlit`
   folders (drag the folders themselves so the structure is kept). Hidden folders such as `.streamlit` may not
   appear in Windows Explorer; turn on **View › Hidden items** first.
3. Commit to `main`. The existing `app.py` and `requirements.txt` are replaced.
4. Streamlit Cloud rebuilds automatically (first build takes 3–5 minutes because it installs Java).
   If it does not, open share.streamlit.io › your app › **⋮ › Reboot**.

> Do **not** upload client workbooks. `.gitignore` excludes a `samples/` folder for local testing.
> If client confidentiality matters, set the repository to **Private** (Settings › General ›
> Danger Zone › Change visibility). Streamlit Community Cloud can deploy from private repos, and
> you can restrict who can open the app under **App settings › Sharing**.

### Optional: AI planning review
In share.streamlit.io › app › **Settings › Secrets**, add:
```toml
ANTHROPIC_API_KEY = "sk-ant-..."
# ANTHROPIC_MODEL = "claude-sonnet-5-5"
```
The "AI planning review" panel only appears when a key is set.

### If the Java install ever fails on Streamlit Cloud
Delete `packages.txt` and the `mpxj` / `JPype1` lines in `requirements.txt`. The app keeps working.
For P6, import the MS Project XML (File › Import › Microsoft Project).

## Run locally / test

```bash
pip install -r requirements.txt
streamlit run app.py
pytest -q
```

## Changing the policy
Edit `firm_policy.json`, e.g. `"max_lead_weeks": 0` to forbid leads, `"close_open_ends": true`,
`"completion_milestone_label_aliases"` to rename specific completion milestones (e.g. `MS` → `Programme Milestones`),
or `working_days` / `hours_per_day` for a 6-day site calendar.

---
A Naphtali PM Group product. © 2026
