"""End-to-end conversion: spreadsheet bytes -> verified, policy-compliant schedule."""
from __future__ import annotations

import datetime as dt
import re
from collections import Counter
from dataclasses import dataclass, field
from typing import Optional

from . import builder, calendar_detect, cpm, logic, reader
from .model import Link, WorkCalendar, load_policy


class NoScheduleFound(ValueError):
    """Raised when no sheet/header could be recognised; the app then offers manual column mapping."""


@dataclass
class Options:
    sheet_name: Optional[str] = None
    project_name: Optional[str] = None
    project_start: Optional[dt.date] = None      # re-baseline to a new start date
    dayfirst: bool = True
    holidays: list = field(default_factory=list)
    use_source_logic: bool = True
    hierarchy_mode: Optional[str] = None         # wbs | indent | summary-rows | group | flat
    policy: dict = field(default_factory=dict)   # overrides on firm_policy.json
    calendar_name: Optional[str] = None          # explicit calendar name (max 51 chars in MS Project)
    fit_source_logic: bool = True                # adjust workbook links so the Excel dates are kept
    work_week: Optional[int] = None              # 5, 6 or 7 working days; None = auto-detect
    mapping: Optional[dict] = None               # manual layout: {sheet, header_row (1-based), cols{key: idx}, period_unit}
    hours_per_day: Optional[float] = None        # None = firm policy


@dataclass
class Result:
    project: object
    nodes: list
    cal: WorkCalendar                 # calendar on the original Excel dates
    out_cal: WorkCalendar             # calendar used for output (may be re-baselined)
    policy: dict
    calendar_name: str
    sheet_name: str
    sheets_found: list
    header_row: int
    columns: dict
    rows: list
    ignored: list
    notes: list
    info: dict
    anchor: object
    issues: list
    stats: dict
    floats: dict


def convert(data: bytes, filename: str, opts: Optional[Options] = None) -> Result:
    opts = opts or Options()
    policy = load_policy()
    policy.update(opts.policy or {})

    sheets = reader.read_sheets(data, filename)
    if opts.mapping:
        sheet, header = reader.header_from_mapping(sheets, opts.mapping)
        ranked = [(sheet, header, 0)]
    else:
        ranked = reader.rank_sheets(sheets)
        if not ranked:
            raise NoScheduleFound(
                "The schedule layout wasn't recognised automatically. Map the columns below: the app needs an "
                "activity column and either start/finish dates, start period numbers with durations, or a "
                "Gantt timeline.")
        choice = ranked[0]
        if opts.sheet_name:
            match = [x for x in ranked if x[0].name == opts.sheet_name]
            if match:
                choice = match[0]
        sheet, header, _ = choice

    # ---------------- work week ---------------- #
    default_days = policy.get("working_days") or calendar_detect.week_days(5)
    rows, ignored, notes = reader.extract_rows(sheet, header, opts.project_start, opts.dayfirst,
                                               len(default_days))
    if opts.work_week in (5, 6, 7):
        n_days, why, how = opts.work_week, "set by the planner", "manual"
    else:
        n_days, why, how = calendar_detect.detect_work_week(sheets, rows, default_days)
    working_days = default_days if (n_days == len(default_days) and how != "manual") \
        else calendar_detect.week_days(n_days)
    if len(working_days) != len(default_days):
        # week-number / Gantt schedules: finish dates depend on the length of the working week
        rows, ignored, notes = reader.extract_rows(sheet, header, opts.project_start, opts.dayfirst,
                                                   len(working_days))
    if opts.hours_per_day:
        hpd = float(opts.hours_per_day)
        if abs(hpd - float(policy.get("hours_per_day", 8))) > 1e-9:
            h0, m0 = (int(x) for x in str(policy.get("day_start", "08:00")).split(":"))
            lunch = 1 if policy.get("lunch_break") else 0
            end = h0 * 60 + m0 + int(round((hpd + lunch) * 60))
            if end > 24 * 60:                     # long shifts: no lunch split, start earlier if needed
                policy["lunch_break"] = []
                end = min(h0 * 60 + m0 + int(round(hpd * 60)), 24 * 60)
                h0 = max(0, (end - int(round(hpd * 60))) // 60)
                policy["day_start"] = f"{h0:02d}:00"
            policy["day_finish"] = f"{end // 60:02d}:{end % 60:02d}" if end < 24 * 60 else "23:59"
        policy["hours_per_day"] = hpd
    policy["working_days"] = working_days

    dated = [r for r in rows if r.start or r.finish]
    if not dated:
        raise ValueError(f"Sheet '{sheet.name}' has activities but no usable dates.")
    excel_start = min((r.start or r.finish) for r in dated)

    cal = WorkCalendar(excel_start, working_days, opts.holidays, policy.get("hours_per_day", 8))
    out_start = opts.project_start or excel_start
    out_cal = WorkCalendar(out_start, working_days, opts.holidays, policy.get("hours_per_day", 8))

    project_name = (opts.project_name or reader.suggest_project_name(sheet, header.row, filename)).strip()
    group_header = ""
    if "group" in header.cols:
        group_header = str(sheet.cell(header.row, header.cols["group"]) or "").strip()

    project, nodes, info = builder.build_outline(rows, project_name, cal, policy,
                                                 group_header or "Group", opts.hierarchy_mode,
                                                 header.group_legend)
    info["assumed_start"] = header.assumed_start
    info["mapping"] = {"sheet": sheet.name, "header_row": header.row + 1, "cols": dict(header.cols),
                       "period_unit": header.period_unit,
                       "headers": {k: str(sheet.cell(header.row, c) or "").replace("\n", " ").strip()
                                   for k, c in header.cols.items()},
                       "manual": bool(opts.mapping)}
    info["period_unit"] = header.period_unit if "start_wk" in header.cols else None

    # ---------------- logic ---------------- #
    sched = [n for n in nodes if n.kind in ("task", "milestone")]
    src_linked, src_problems = set(), []
    if opts.use_source_logic and "preds" in header.cols:
        src_linked, src_problems = logic.apply_source_logic(nodes, info.get("duration_unit", "d"),
                                                            cal.days_per_week, opts.fit_source_logic)
    to_infer = set(n for n in sched if n not in src_linked)
    anchor = logic.infer_logic(sched, policy, only=to_infer, dpw=cal.days_per_week)
    logic.link_completion_milestones(project)
    if policy.get("close_open_ends"):
        info["open_ends_closed"] = logic.close_open_ends(nodes)
    info["logic_source"] = ("spreadsheet + inferred" if src_linked else "inferred from dates")
    info["work_week"] = {"days": cal.days_per_week, "label": calendar_detect.LABEL[cal.days_per_week],
                         "reason": why, "source": how, "hours_per_day": policy.get("hours_per_day", 8)}

    # ---------------- verify (CPM) ---------------- #
    cpm.forward_pass(nodes, 0)
    floats = cpm.backward_pass(nodes)

    issues, stats = quality_checks(nodes, cal, out_cal, rows, header, anchor, floats, src_problems, policy)
    for r_no, name, why in ignored + info.get("dropped", []):
        issues.append(("info", "", name[:80], f"Excel row {r_no} ignored: {why}"))

    return Result(project, nodes, cal, out_cal, policy,
                  calendar_name_for(project_name, policy, opts.calendar_name),
                  sheet.name, [s[0].name for s in ranked], header.row + 1,
                  {k: v for k, v in header.cols.items()}, rows, ignored, notes, info,
                  anchor, issues, stats, floats)


# --------------------------------------------------------------------------- #
# Calendar name (MS Project rejects calendar names longer than 51 characters)
# --------------------------------------------------------------------------- #
MSP_CALENDAR_MAX = 51


def calendar_name_for(project_name: str, policy: dict, override: Optional[str] = None) -> str:
    limit = int(policy.get("calendar_name_max_length", MSP_CALENDAR_MAX))
    if override and override.strip():
        return override.strip()[:limit]
    suffix = policy.get("calendar_name_suffix", " CALENDAR")
    base = re.sub(r"\s+", " ", project_name).strip()

    def fits(b):
        return len(b) + len(suffix) <= limit

    steps = [
        lambda b: re.sub(r"\s*\([^)]*\)", "", b),                                  # drop (notes)
        lambda b: re.sub(r"\bLevel\s*(\d)\b", r"L\1", b, flags=re.I),             # Level 2 -> L2
        lambda b: _acronym_first_segment(b),                                        # "NMHIRP Conceptual ..." -> "NMHIRP"
        lambda b: re.sub(r"\s+-\s+", " ", b),                                       # drop " - "
        lambda b: re.sub(r"\bSchedule\b", "Sch", b, flags=re.I),
    ]
    for step in steps:
        if fits(base):
            break
        base = step(base).strip()
    if not fits(base):
        room = limit - len(suffix)
        cut = base[:room]
        base = cut.rsplit(" ", 1)[0] if " " in cut and not base[room:room + 1] in ("", " ") else cut
        base = base.rstrip(" -")
        words = base.split()
        while len(words) > 1 and words[-1].lower() in {"and", "of", "the", "for", "&", "at", "in", "with", "to", "-"}:
            words.pop()
        base = " ".join(words)
    return base + suffix


def _acronym_first_segment(b: str) -> str:
    parts = re.split(r"\s+-\s+", b)
    words = parts[0].split()
    if len(parts) > 1 and len(words) > 1 and re.fullmatch(r"[A-Z0-9&/]{2,}", words[0]):
        parts[0] = words[0]
    return " - ".join(parts)


# --------------------------------------------------------------------------- #
# Quality checks (DCMA-style + Excel reconciliation)
# --------------------------------------------------------------------------- #
def quality_checks(nodes, cal, out_cal, rows, header, anchor, floats, src_problems, policy):
    issues = []
    sched = [n for n in nodes if n.schedulable]
    src_items = [n for n in sched if n.src is not None]

    # 1. Excel date reconciliation
    mismatches = 0
    for n in src_items:
        r = n.src
        if n.kind == "milestone":
            lo = cal.start_b(r.start) if r.start else n.t_start
            hi = cal.finish_b(r.finish) if r.finish else n.t_finish
            ok = lo <= n.es <= hi
        else:
            ok = (n.es == n.t_start and n.ef == n.t_finish)
        if not ok:
            mismatches += 1
            s, f = cpm.display_dates(n, cal)
            issues.append(("error", n.id, n.name,
                           f"Calculated {s:%d-%b-%y} -> {f:%d-%b-%y} differs from Excel "
                           f"{_d(r.start)} -> {_d(r.finish)}"))

    # 2. Source summary rows vs rolled-up dates
    for n in nodes:
        if n.kind == "summary" and n.src is not None and n.src.start and n.src.finish and n.es is not None:
            s, f = cpm.display_dates(n, cal)
            if s != n.src.start or f != n.src.finish:
                issues.append(("warning", n.id, n.name,
                               f"Summary in Excel shows {_d(n.src.start)} -> {_d(n.src.finish)}; "
                               f"its activities roll up to {s:%d-%b-%y} -> {f:%d-%b-%y}"))

    # 3. Duration column vs dates
    for n in src_items:
        r = n.src
        if r.duration is None or not (r.start and r.finish) or n.kind != "task":
            continue
        wd = cal.working_days_between(r.start, r.finish)
        if (r.dur_unit or "d") == "mo":
            continue                              # calendar months vary in working days
        exp = r.duration * (cal.days_per_week if (r.dur_unit or "d") == "w" else 1)
        if abs(wd - exp) > 0.01:
            issues.append(("warning", n.id, n.name,
                           f"Excel duration {r.duration:g} {r.dur_unit or 'd'} does not match its dates "
                           f"({wd} working days). Dates were used."))

    # 4. Non-working dates in Excel
    for n in src_items:
        for lab, d in (("start", n.src.start), ("finish", n.src.finish)):
            if d and not cal.is_working(d) and not (n.kind == "milestone"):
                issues.append(("warning", n.id, n.name, f"Excel {lab} {d:%a %d-%b-%y} is a non-working day"))

    # 5. Logic health
    succ = Counter()
    for n in sched:
        for l in n.preds:
            succ[l.pred] += 1
    open_starts = [n for n in sched if not n.preds and n is not anchor]
    open_ends = [n for n in sched if succ[n] == 0 and n.kind != "completion"]
    links = [l for n in sched for l in n.preds]
    leads = [l for l in links if l.lag < 0]
    lags = [l for l in links if l.lag > 0]
    for n in open_starts:
        issues.append(("warning", n.id, n.name, "No predecessor (open start)"))
    for n in open_ends:
        issues.append(("info", n.id, n.name, "No successor (open end) - covered by the summary's "
                                             "completion milestone date"))

    proj_ef = max(n.ef for n in sched)
    excel_finish = max((r.finish for r in rows if r.finish), default=None)
    calc_finish = cal.date(proj_ef - 1)
    out_s = out_cal.date(0)
    out_f = out_cal.date(proj_ef - 1)
    crit = [n for n in sched if floats.get(n, 1) <= 0]

    stats = {
        "activities": sum(1 for n in sched if n.kind == "task"),
        "milestones": sum(1 for n in sched if n.kind == "milestone"),
        "completion_milestones": sum(1 for n in sched if n.kind == "completion"),
        "summaries": sum(1 for n in nodes if n.is_summary),
        "rows_total": len(nodes),
        "links": len(links),
        "rel_mix": dict(Counter(l.type for l in links)),
        "leads": len(leads),
        "lags": len(lags),
        "open_starts": len(open_starts),
        "open_ends": len(open_ends),
        "date_mismatches": mismatches,
        "excel_start": cal.date(0),
        "excel_finish": excel_finish,
        "calc_finish": calc_finish,
        "finish_matches": excel_finish == calc_finish,
        "output_start": out_s,
        "output_finish": out_f,
        "duration_days": proj_ef,
        "critical": len(crit),
        "anchor": anchor.name if anchor else "",
        "src_logic_problems": sum(1 for p in src_problems if p[0] == "warning"),
        "src_links_fitted": sum(1 for p in src_problems if p[0] == "info"),
    }
    for sev, pid, name, msg in src_problems:
        issues.append((sev, pid, name, msg))
    if not stats["finish_matches"]:
        issues.insert(0, ("error", "", "Project", f"Calculated finish {calc_finish:%d-%b-%y} differs from "
                                                  f"Excel finish {_d(excel_finish)}"))
    return issues, stats


def _d(d):
    return d.strftime("%d-%b-%y") if d else "-"


# --------------------------------------------------------------------------- #
# Manual edits from the app (planner overrides a predecessor string)
# --------------------------------------------------------------------------- #
def apply_edited_predecessors(res: Result, edits: dict):
    """edits: {node_id: 'text like 8FS-1 wk, 10SS'}; re-runs verification."""
    by_id = {n.id: n for n in res.nodes}
    unit = res.info.get("duration_unit", "d")
    for nid, text in edits.items():
        n = by_id.get(int(nid))
        if n is None or n.is_summary:
            continue
        links = []
        for tok in [t.strip() for t in str(text or "").replace(";", ",").split(",") if t.strip()]:
            m = logic._TOKEN.match(tok)
            if not m:
                continue
            try:
                p = by_id[int(m.group("ref"))]
            except (ValueError, KeyError):
                continue
            if p.is_summary or p is n:
                continue
            links.append(Link(p, (m.group("type") or "FS").upper(),
                              logic._lag_days(m.group("lag"), unit, res.cal.days_per_week), "edited by planner"))
        n.preds = links
    for n in res.nodes:
        n.es = n.ef = None
    cpm.forward_pass(res.nodes, 0)
    res.floats = cpm.backward_pass(res.nodes)
    res.issues, res.stats = quality_checks(res.nodes, res.cal, res.out_cal, res.rows, None, res.anchor,
                                           res.floats, [], res.policy)
    return res


# --------------------------------------------------------------------------- #
# Layout inspection (for manual column mapping in the app)
# --------------------------------------------------------------------------- #
def inspect_layout(data: bytes, filename: str) -> list:
    """[{sheet, header_row (1-based guess), columns: [(idx, label)], guess: {key: idx}}] for every sheet."""
    out = []
    for sh in reader.read_sheets(data, filename):
        hr = reader.guess_header_row(sh)
        cols, guess = [], {}
        if hr is not None:
            for c, v in enumerate(sh.values[hr]):
                label = str(v).replace("\n", " ").strip() if v not in (None, "") else ""
                letter = reader.col_letter(c)
                cols.append((c, f"{letter}: {label}" if label else letter))
                k = reader.classify_header(v) if isinstance(v, str) else None
                if k and k not in guess:
                    guess[k] = c
        out.append({"sheet": sh.name, "header_row": (hr + 1) if hr is not None else 1, "columns": cols,
                    "guess": guess, "rows": sh.nrows})
    return out
