"""End-to-end conversion: spreadsheet bytes -> verified, policy-compliant schedule."""
from __future__ import annotations

import datetime as dt
from collections import Counter
from dataclasses import dataclass, field
from typing import Optional

from . import builder, cpm, logic, reader
from .model import Link, WorkCalendar, load_policy


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
    ranked = reader.rank_sheets(sheets)
    if not ranked:
        raise ValueError(
            "No schedule found. The workbook needs a sheet with an activity/task column and "
            "either Start/Finish dates, Start week + Duration, or a Gantt timeline.")
    choice = ranked[0]
    if opts.sheet_name:
        match = [x for x in ranked if x[0].name == opts.sheet_name]
        if match:
            choice = match[0]
    sheet, header, _ = choice

    rows, ignored, notes = reader.extract_rows(sheet, header, opts.project_start, opts.dayfirst)
    dated = [r for r in rows if r.start or r.finish]
    if not dated:
        raise ValueError(f"Sheet '{sheet.name}' has activities but no usable dates.")
    excel_start = min((r.start or r.finish) for r in dated)

    working_days = policy.get("working_days")
    cal = WorkCalendar(excel_start, working_days, opts.holidays, policy.get("hours_per_day", 8))
    out_start = opts.project_start or excel_start
    out_cal = WorkCalendar(out_start, working_days, opts.holidays, policy.get("hours_per_day", 8))

    project_name = (opts.project_name or reader.suggest_project_name(sheet, header.row, filename)).strip()
    group_header = ""
    if "group" in header.cols:
        group_header = str(sheet.cell(header.row, header.cols["group"]) or "").strip()

    project, nodes, info = builder.build_outline(rows, project_name, cal, policy,
                                                 group_header or "Group", opts.hierarchy_mode)

    # ---------------- logic ---------------- #
    sched = [n for n in nodes if n.kind in ("task", "milestone")]
    src_linked, src_problems = set(), []
    if opts.use_source_logic and "preds" in header.cols:
        src_linked, src_problems = logic.apply_source_logic(nodes, info.get("duration_unit", "d"))
    to_infer = set(n for n in sched if n not in src_linked)
    anchor = logic.infer_logic(sched, policy, only=to_infer)
    logic.link_completion_milestones(project)
    if policy.get("close_open_ends"):
        info["open_ends_closed"] = logic.close_open_ends(nodes)
    info["logic_source"] = ("spreadsheet + inferred" if src_linked else "inferred from dates")

    # ---------------- verify (CPM) ---------------- #
    cpm.forward_pass(nodes, 0)
    floats = cpm.backward_pass(nodes)

    issues, stats = quality_checks(nodes, cal, out_cal, rows, header, anchor, floats, src_problems, policy)
    for r_no, name, why in ignored + info.get("dropped", []):
        issues.append(("info", "", name[:80], f"Excel row {r_no} ignored: {why}"))

    return Result(project, nodes, cal, out_cal, policy,
                  project_name + policy.get("calendar_name_suffix", " CALENDAR"),
                  sheet.name, [s[0].name for s in ranked], header.row + 1,
                  {k: v for k, v in header.cols.items()}, rows, ignored, notes, info,
                  anchor, issues, stats, floats)


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
        exp = r.duration * (5 if (r.dur_unit or "d") == "w" else 1)
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
        "src_logic_problems": len(src_problems),
    }
    for pid, name, msg in src_problems:
        issues.append(("warning", pid, name, msg))
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
                              logic._lag_days(m.group("lag"), unit), "edited by planner"))
        n.preds = links
    for n in res.nodes:
        n.es = n.ef = None
    cpm.forward_pass(res.nodes, 0)
    res.floats = cpm.backward_pass(res.nodes)
    res.issues, res.stats = quality_checks(res.nodes, res.cal, res.out_cal, res.rows, None, res.anchor,
                                           res.floats, [], res.policy)
    return res
