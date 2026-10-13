# Schedule Converter — Naphtali PM Group

Upload **any Excel schedule** and get files that open in **Microsoft Project** and **Primavera P6**
with the firm's scheduling policy applied and **every date verified against the Excel**.

**New in Phase 1 (AutoSchedule Engine): Document scope ingestion.** Upload a Statement of Work, charter or scope
narrative (`.docx`, text-based `.pdf`, `.txt`, `.md`). Google Gemini proposes a scope register with a citation for
every line; a planner reviews and approves it in the **Review Gate**; the same proven CPM engine then builds,
verifies (zero-day variance) and exports the baseline, with a **DCMA-14** report. The AI never sets a date.

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

Work week: detected automatically (Mon–Fri, Mon–Sat or Mon–Sun) from weekend dates, day-based
durations and wording such as "6-day week" or "24/7"; firm default Mon–Fri. Override it, and the hours per
day, in the sidebar. On 6- and 7-day calendars durations and lags are written in days so MS Project
imports them exactly.

Any layout:
- Dates, week/month/day numbers ("Start (Month)", "Start Wk", "Duration (Months)"), or a drawn Gantt.
- Headers with units in brackets or line breaks; codes explained in a legend (e.g. "A - Substructure").
- No dates at all: period 1 is assumed to start next Monday (shown as a warning) until you set the real start.
- Workbook predecessors are kept; self-links and loops are removed; links that contradict the dates are
  fitted (type/lag) so nothing moves. Every change is listed in **File check**.
- Layout still not recognised: the **column mapper** appears instead of an error.
- **Standardised workbook** (File check) gives the schedule back in the standard layout to edit and re-upload;
  a blank template is on the start page.

Extras in the app: a sidebar for policy settings (max lead, close open ends, holidays, date format), a
re-baseline option (move the whole schedule to a new start/NTP date), predecessors you can edit and
re-verify, a Gantt preview with the critical path, and an optional AI planning review.

## Repository layout

```
app.py                  Streamlit user interface (Excel converter tab + Document scope ingestion tab)
firm_policy.json        Firm scheduling policy (edit to change defaults)
scheduler/
  reader.py             Sheet / header / column detection, dates, week & Gantt parsing
  builder.py            Outline, completion milestones, WBS
  logic.py              Predecessor inference + parsing of existing predecessors
  cpm.py                MS Project-equivalent forward/backward pass
  engine.py             Orchestration + quality checks
  exporters.py          CSV, MS Project XML, P6 XER, QA workbook
  doc_reader.py         NEW  .docx / .pdf / .txt text extraction into cited segments (P12, T2.R5, PG3.4)
  extractor.py          NEW  Gemini structured extraction (response_schema=ProjectScope) + citation checks
  scope_builder.py      NEW  Review Gate validation, deterministic dates, bridge into engine.convert
  dcma.py               NEW  DCMA 14-point assessment (both tabs) + Excel report
tests/test_engine.py    Regression tests on synthetic schedules (pytest)
tests/test_scope_pipeline.py  NEW  Document pipeline tests (mocked Gemini, no network)
requirements.txt        Python packages (pinned)
packages.txt            Java runtime for the P6 XER export (Streamlit Cloud installs it)
.streamlit/config.toml  Theme
.streamlit/secrets.toml.example  NEW  Template for local secrets (copy to secrets.toml; never commit it)
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

### Required for Document scope ingestion: Gemini API key
In share.streamlit.io › app › **Settings › Secrets**, add:
```toml
GEMINI_API_KEY = "your-google-ai-studio-key"
# GEMINI_MODEL = "auto"   # default; or a model id such as "gemini-3.8-flash"
```
Without a key the Excel converter works as before and the document tab shows extraction as disabled.
With `auto` the app asks Google which models the key can use and picks the newest Flash model, falling back
automatically on 404 (model not available to the key), quota/rate limits (429) or unsupported settings (400).
Free-tier keys work. Use **Check models for this key** in the sidebar to see the order it will try.
For local testing you can paste a key into the sidebar field instead (kept only for that browser session).

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

Optional DCMA-14 thresholds (defaults shown; add only to change them):
`"dcma_threshold_pct": 5`, `"dcma_high_float_days": 44`, `"dcma_high_duration_days": 44`,
`"dcma_fs_min_pct": 90`, `"dcma_cpt_delay_days": 600`.

---
A Naphtali PM Group product. © 2026
