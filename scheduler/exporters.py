"""Outputs: MS Project import CSV, MS Project XML (with the policy calendar), QA workbook."""
from __future__ import annotations

import datetime as dt
import io
import unicodedata
import xml.etree.ElementTree as ET

import pandas as pd

from . import cpm
from .logic import link_text

_ASCII_MAP = {"–": "-", "—": "-", "‘": "'", "’": "'", "“": '"', "”": '"',
              "…": "...", " ": " ", "•": "-", "◆": "", "♦": ""}


def clean_text(s: str) -> str:
    for k, v in _ASCII_MAP.items():
        s = s.replace(k, v)
    return unicodedata.normalize("NFC", s).strip()


def shown_unit(n, dpw: int = 5) -> str:
    # Weeks are only shown for a 5-day week: MS Project's CSV import converts "wks" with its
    # default 40 h/week, so 6- and 7-day calendars are written in days to stay exact.
    return n.dur_unit if dpw == 5 else "d"


def duration_text(n, dpw: int = 5) -> str:
    if n.is_summary:
        return ""
    d = n.dur
    if shown_unit(n, dpw) == "w" and d % 5 == 0:
        w = d // 5
        return "1 wk" if w == 1 else f"{w} wks"
    return "1 day" if d == 1 else f"{d} days"


def schedule_table(res) -> pd.DataFrame:
    """The import table (one row per MS Project task) using output dates."""
    dpw = res.out_cal.days_per_week
    rows = []
    for n in res.nodes:
        s, f = cpm.display_dates(n, res.out_cal)
        preds = ", ".join(link_text(l, shown_unit(n, dpw), dpw) for l in n.preds)
        rows.append({
            "ID": n.id,
            "Task Name": clean_text(n.name),
            "Duration": duration_text(n, dpw),
            "Start Date": s,
            "Finish Date": f,
            "Predecessors": preds,
            "Outline Level": n.level,
            "WBS": n.wbs,
            "_kind": n.kind,
            "_excel_row": n.src.excel_row if n.src else None,
            "_excel_start": n.src.start if n.src else None,
            "_excel_finish": n.src.finish if n.src else None,
            "_float_days": res.floats.get(n) if n in res.floats else None,
        })
    return pd.DataFrame(rows)


def to_csv(res) -> bytes:
    df = schedule_table(res)
    fmt = res.policy.get("csv_date_format", "%Y-%m-%d")
    cols = res.policy.get("csv_columns") or ["ID", "Task Name", "Duration", "Start Date", "Finish Date",
                                             "Predecessors", "Outline Level", "WBS"]
    out = df[cols].copy()
    for c in ("Start Date", "Finish Date"):
        out[c] = out[c].map(lambda d: d.strftime(fmt) if d else "")
    text = out.to_csv(index=False, lineterminator="\r\n")
    try:
        return text.encode("ascii")
    except UnicodeEncodeError:
        return text.encode("utf-8-sig")


# --------------------------------------------------------------------------- #
# MS Project XML (MSPDI) - opens directly in MS Project and imports into P6
# --------------------------------------------------------------------------- #
NS = "http://schemas.microsoft.com/project"


def _hm(t: str) -> tuple:
    h, m = t.split(":")
    return int(h), int(m)


def _dur_iso(days: float, hpd: float) -> str:
    total_min = int(round(days * hpd * 60))
    return f"PT{total_min // 60}H{total_min % 60}M0S"


def to_mspdi(res, author: str = "Naphtali PM Group") -> bytes:
    pol = res.policy
    cal = res.out_cal
    dpw = cal.days_per_week
    hpd = float(pol.get("hours_per_day", 8))
    ds, df_ = _hm(pol.get("day_start", "08:00")), _hm(pol.get("day_finish", "17:00"))
    lunch = pol.get("lunch_break") or []

    def at(d: dt.date, hm) -> str:
        return dt.datetime(d.year, d.month, d.day, hm[0], hm[1]).strftime("%Y-%m-%dT%H:%M:%S")

    def start_dt(b):
        return at(cal.date(b), ds)

    def finish_dt(b):
        return at(cal.date(b - 1), df_)

    ET.register_namespace("", NS)
    P = ET.Element(f"{{{NS}}}Project")

    def sub(parent, tag, text=None):
        e = ET.SubElement(parent, f"{{{NS}}}{tag}")
        if text is not None:
            e.text = str(text)
        return e

    proj_items = [n for n in res.nodes if n.schedulable]
    proj_ef = max(n.ef for n in proj_items)
    now = dt.datetime.now().strftime("%Y-%m-%dT%H:%M:%S")

    sub(P, "SaveVersion", 14)
    sub(P, "Name", clean_text(res.project.name)[:200] + ".xml")
    sub(P, "Title", clean_text(res.project.name))
    sub(P, "Company", "Naphtali PM Group")
    sub(P, "Author", author)
    sub(P, "CreationDate", now)
    sub(P, "ScheduleFromStart", 1)
    sub(P, "StartDate", start_dt(0))
    sub(P, "FinishDate", finish_dt(proj_ef))
    sub(P, "CalendarUID", 1)
    sub(P, "DefaultStartTime", f"{ds[0]:02d}:{ds[1]:02d}:00")
    sub(P, "DefaultFinishTime", f"{df_[0]:02d}:{df_[1]:02d}:00")
    sub(P, "MinutesPerDay", int(hpd * 60))
    sub(P, "MinutesPerWeek", int(hpd * 60 * dpw))
    sub(P, "DaysPerMonth", int(round(dpw * 52 / 12)))
    sub(P, "DefaultTaskType", 1)
    sub(P, "DurationFormat", 9 if (res.info.get("duration_unit") == "w" and dpw == 5) else 7)
    sub(P, "WorkFormat", 2)
    sub(P, "HonorConstraints", 1)
    sub(P, "NewTasksEffortDriven", 0)
    sub(P, "NewTasksEstimated", 0)
    sub(P, "WeekStartDay", 1)
    sub(P, "CurrentDate", now)
    sub(P, "NewTaskStartDate", 0)
    sub(P, "NewTasksAreManual", 0)

    # ---- calendar (firm policy: '<Project name> CALENDAR') ---- #
    cals = sub(P, "Calendars")
    c = sub(cals, "Calendar")
    sub(c, "UID", 1)
    sub(c, "Name", clean_text(res.calendar_name)[:51])
    sub(c, "IsBaseCalendar", 1)
    sub(c, "IsBaselineCalendar", 0)
    sub(c, "BaseCalendarUID", -1)
    wds = sub(c, "WeekDays")
    names = ["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"]
    working = set(cal.working_days)
    for i, nm in enumerate(names, start=1):
        wd = sub(wds, "WeekDay")
        sub(wd, "DayType", i)
        sub(wd, "DayWorking", 1 if nm in working else 0)
        if nm in working:
            wts = sub(wd, "WorkingTimes")
            periods = [(pol.get("day_start", "08:00"), pol.get("day_finish", "17:00"))]
            if len(lunch) == 2:
                periods = [(pol.get("day_start", "08:00"), lunch[0]), (lunch[1], pol.get("day_finish", "17:00"))]
            for a, b in periods:
                wt = sub(wts, "WorkingTime")
                sub(wt, "FromTime", f"{a}:00")
                sub(wt, "ToTime", f"{b}:00")
    for h in cal.holidays:
        wd = sub(wds, "WeekDay")
        sub(wd, "DayType", 0)
        sub(wd, "DayWorking", 0)
        tp = sub(wd, "TimePeriod")
        sub(tp, "FromDate", at(h, (0, 0)))
        sub(tp, "ToDate", at(h, (23, 59)))

    # ---- tasks ---- #
    tasks = sub(P, "Tasks")
    t0 = sub(tasks, "Task")
    for tag, val in (("UID", 0), ("ID", 0), ("Name", clean_text(res.project.name)), ("Type", 1),
                     ("IsNull", 0), ("WBS", 0), ("OutlineNumber", 0), ("OutlineLevel", 0)):
        sub(t0, tag, val)
    sub(t0, "Start", start_dt(0))
    sub(t0, "Finish", finish_dt(proj_ef))
    sub(t0, "Duration", _dur_iso(proj_ef, hpd))
    sub(t0, "DurationFormat", 7)
    sub(t0, "Milestone", 0)
    sub(t0, "Summary", 1)
    sub(t0, "RemainingDuration", _dur_iso(proj_ef, hpd))

    for n in res.nodes:
        t = sub(tasks, "Task")
        sub(t, "UID", n.uid)
        sub(t, "ID", n.id)
        sub(t, "Name", clean_text(n.name))
        sub(t, "Active", 1)
        sub(t, "Manual", 0)
        sub(t, "Type", 1)
        sub(t, "IsNull", 0)
        sub(t, "CreateDate", now)
        sub(t, "WBS", n.wbs)
        sub(t, "OutlineNumber", _outline_number(n))
        sub(t, "OutlineLevel", n.level)
        sub(t, "Priority", 500)
        if n.is_milestone:
            stamp = finish_dt(n.es) if (n.ms_end_of_day and n.es > 0) else start_dt(n.es)
            sub(t, "Start", stamp)
            sub(t, "Finish", stamp)
        else:
            sub(t, "Start", start_dt(n.es))
            sub(t, "Finish", finish_dt(n.ef))
        sub(t, "Duration", _dur_iso(n.dur, hpd))
        sub(t, "DurationFormat", 9 if (shown_unit(n, dpw) == "w" and n.dur % 5 == 0) else 7)
        sub(t, "ResumeValid", 0)
        sub(t, "EffortDriven", 0)
        sub(t, "Recurring", 0)
        sub(t, "OverAllocated", 0)
        sub(t, "Estimated", 0)
        sub(t, "Milestone", 1 if n.is_milestone else 0)
        sub(t, "Summary", 1 if n.is_summary else 0)
        sub(t, "Critical", 0)
        sub(t, "IsSubproject", 0)
        sub(t, "IsSubprojectReadOnly", 0)
        sub(t, "ExternalTask", 0)
        sub(t, "FixedCostAccrual", 3)
        # MS Project imports Duration as ActualDuration + RemainingDuration:
        # without RemainingDuration every task would come in as 0 days.
        sub(t, "RemainingDuration", _dur_iso(n.dur, hpd))
        sub(t, "ConstraintType", 0)
        sub(t, "CalendarUID", -1)
        if n.src is not None:
            sub(t, "Notes", f"Source: {res.sheet_name} row {n.src.excel_row}")
        for l in n.preds:
            pl = sub(t, "PredecessorLink")
            sub(pl, "PredecessorUID", l.pred.uid)
            sub(pl, "Type", {"FF": 0, "FS": 1, "SF": 2, "SS": 3}[l.type])
            sub(pl, "CrossProject", 0)
            sub(pl, "LinkLag", int(round(l.lag * hpd * 60 * 10)))
            sub(pl, "LagFormat", 9 if (shown_unit(n, dpw) == "w" and l.lag % 5 == 0) else 7)

    ET.indent(P, space="  ")
    body = ET.tostring(P, encoding="utf-8", xml_declaration=False)
    return b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' + body


def _outline_number(n) -> str:
    return n.wbs


# --------------------------------------------------------------------------- #
# Primavera P6 XER (optional: needs Java + the 'mpxj' package)
# --------------------------------------------------------------------------- #
def xer_available() -> bool:
    try:
        _start_jvm()
        return True
    except Exception:
        return False


def _start_jvm():
    import jpype
    import mpxj  # noqa: F401  (adds the MPXJ jars to the class path)

    if not jpype.isJVMStarted():
        jpype.startJVM("-Xmx384m", convertStrings=True)


def to_xer(res) -> bytes:
    import os
    import tempfile

    _start_jvm()
    from org.mpxj.reader import UniversalProjectReader
    from org.mpxj.writer import FileFormat, UniversalProjectWriter

    with tempfile.TemporaryDirectory() as tmp:
        src = os.path.join(tmp, "schedule.xml")
        dst = os.path.join(tmp, "schedule.xer")
        with open(src, "wb") as fh:
            fh.write(to_mspdi(res))
        project = UniversalProjectReader().read(src)
        UniversalProjectWriter(FileFormat.XER).write(project, dst)
        with open(dst, "rb") as fh:
            raw = fh.read()
    return _patch_xer(raw, res, str(res.policy.get("p6_xer_version", XER_VERSION)))


XER_VERSION = "19.12"   # P6 refuses XER files from a newer version; 19.12 opens in P6 19 and later


def _patch_xer(raw: bytes, res, version: str = XER_VERSION) -> bytes:
    """Make the XER import cleanly in P6 19+: version header, activity types, planned dates, data date."""
    text = raw.decode("cp1252", errors="replace")
    lines = text.split("\r\n") if "\r\n" in text else text.split("\n")
    cal = res.out_cal
    by_uid = {n.uid: n for n in res.nodes}

    def p6(d, hm):
        return f"{d:%Y-%m-%d} {hm}"

    out, table, fields = [], None, []
    for line in lines:
        cols = line.split("\t")
        if cols[0] == "ERMHDR" and len(cols) > 1:
            cols[1] = version
        elif cols[0] == "%T":
            table = cols[1] if len(cols) > 1 else None
        elif cols[0] == "%F":
            fields = cols[1:]
        elif cols[0] == "%R" and table in ("TASK", "PROJECT"):
            rec = dict(zip(fields, cols[1:] + [""] * (len(fields) - len(cols) + 1)))
            if table == "PROJECT":
                start = p6(cal.date(0), "08:00")
                rec["last_recalc_date"] = rec.get("last_recalc_date") or start   # data date
                rec["plan_start_date"] = rec.get("plan_start_date") or start
                rec["def_task_type"] = "TT_Task"
            else:
                n = by_uid.get(int(rec.get("task_id") or -1))
                if n is not None:
                    s, f = cpm.display_dates(n, cal)
                    if n.is_milestone:
                        start_ms = not n.ms_end_of_day
                        rec["task_type"] = "TT_Mile" if start_ms else "TT_FinMile"
                        stamp = p6(s, "08:00") if start_ms else p6(f, "17:00")
                        rec["target_start_date"] = rec["target_end_date"] = stamp
                        rec["early_start_date"] = rec["early_end_date"] = stamp
                    else:
                        rec["task_type"] = "TT_Task"
                        rec["target_start_date"] = rec["early_start_date"] = p6(s, "08:00")
                        rec["target_end_date"] = rec["early_end_date"] = p6(f, "17:00")
            cols = ["%R"] + [rec.get(k, "") for k in fields]
        out.append("\t".join(cols))
    return "\r\n".join(out).encode("cp1252", errors="replace")


# --------------------------------------------------------------------------- #
# QA workbook
# --------------------------------------------------------------------------- #
def to_qa_xlsx(res) -> bytes:
    df = schedule_table(res)
    st = res.stats
    summary = pd.DataFrame([
        ("Project", res.project.name),
        ("Calendar", res.calendar_name),
        ("Source sheet", res.sheet_name),
        ("Hierarchy detected from", res.info.get("mode")),
        ("Logic", res.info.get("logic_source")),
        ("Excel start / finish", f"{st['excel_start']:%d-%b-%Y} / {st['excel_finish']:%d-%b-%Y}"),
        ("Calculated finish (MS Project)", f"{st['calc_finish']:%d-%b-%Y}"),
        ("Output start / finish", f"{st['output_start']:%d-%b-%Y} / {st['output_finish']:%d-%b-%Y}"),
        ("Duration (working days)", st["duration_days"]),
        ("Activities / milestones / completion milestones",
         f"{st['activities']} / {st['milestones']} / {st['completion_milestones']}"),
        ("Summary tasks (incl. project title)", st["summaries"]),
        ("Relationships", f"{st['links']}  {st['rel_mix']}"),
        ("Leads (negative lag) / positive lags", f"{st['leads']} / {st['lags']}"),
        ("Open starts / open ends", f"{st['open_starts']} / {st['open_ends']}"),
        ("Dates differing from Excel", st["date_mismatches"]),
        ("Critical activities (zero float)", st["critical"]),
    ], columns=["Check", "Result"])
    issues = pd.DataFrame(res.issues, columns=["Severity", "ID", "Task", "Message"])
    trace = df.rename(columns={"_excel_row": "Excel row", "_excel_start": "Excel start",
                               "_excel_finish": "Excel finish", "_float_days": "Total float (days)",
                               "_kind": "Type"})
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as xw:
        summary.to_excel(xw, sheet_name="Summary", index=False)
        issues.to_excel(xw, sheet_name="Issues", index=False)
        trace.to_excel(xw, sheet_name="Schedule & Traceability", index=False)
        for ws in xw.book.worksheets:
            for col in ws.columns:
                width = max(len(str(c.value or "")) for c in col[:200])
                ws.column_dimensions[col[0].column_letter].width = min(max(10, width + 2), 80)
    return buf.getvalue()
