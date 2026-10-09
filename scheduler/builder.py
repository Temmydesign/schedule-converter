"""Build the policy-compliant outline (hierarchy + completion milestones).

Firm policy implemented here:
1. Row 1 is the project title (Outline Level 1); everything sits beneath it.
2. Every summary task ends with a 0-duration milestone "<summary> completed".
3. WBS codes are regenerated to follow the outline (1, 1.1, 1.1.1 ...).
"""
from __future__ import annotations

import re
from collections import Counter
from typing import Optional

from .model import Node, SourceRow, WorkCalendar

_CODE_RE = re.compile(r"^[A-Za-z0-9]+(\.[A-Za-z0-9]+)*\.?$")


def _norm_code(code: str) -> Optional[str]:
    if not code:
        return None
    c = str(code).strip().rstrip(".")
    if not _CODE_RE.match(c):
        return None
    parts = c.split(".")
    while len(parts) > 1 and parts[-1] in ("0", "00"):
        parts.pop()
    return ".".join(parts)


def infer_duration_unit(rows: list, cal: WorkCalendar) -> str:
    """Decide if the source 'duration' numbers are weeks or days."""
    votes = Counter()
    for r in rows:
        if r.is_summary or r.duration in (None, 0) or not (r.start and r.finish):
            continue
        wd = cal.working_days_between(r.start, r.finish)
        if abs(wd - r.duration * cal.days_per_week) < 0.01:
            votes["w"] += 1
        elif abs(wd - r.duration) < 0.01:
            votes["d"] += 1
    if votes:
        return votes.most_common(1)[0][0]
    return "d"


def choose_hierarchy_mode(rows: list) -> str:
    codes = [_norm_code(r.wbs) for r in rows if r.wbs]
    if rows and len(codes) >= 0.6 * len(rows):
        depths = {c.count(".") for c in codes if c}
        if len(depths) > 1:
            return "wbs"
    indents = {r.indent for r in rows}
    if len(indents) > 1:
        return "indent"
    if any(r.is_summary for r in rows):
        return "summary-rows"
    if any(r.group for r in rows):
        return "group"
    return "flat"


def _levels(rows: list, mode: str) -> list:
    """Relative outline level (1 = directly under the project title) per row."""
    levels = []
    if mode == "wbs":
        seen = {}
        for r in rows:
            code = _norm_code(r.wbs)
            if not code:
                lvl = (levels[-1] + 1) if levels and rows[len(levels) - 1].is_summary else (levels[-1] if levels else 1)
            else:
                parts = code.split(".")
                lvl = 1
                for k in range(len(parts) - 1, 0, -1):
                    parent = ".".join(parts[:k])
                    if parent in seen:
                        lvl = seen[parent] + 1
                        break
                seen[code] = lvl
            levels.append(lvl)
    elif mode == "indent":
        ranks = {v: i + 1 for i, v in enumerate(sorted({r.indent for r in rows}))}
        levels = [ranks[r.indent] for r in rows]
    elif mode == "summary-rows":
        cur = 1
        for r in rows:
            if r.is_summary:
                levels.append(1)
                cur = 2
            else:
                levels.append(cur)
    else:
        levels = [1] * len(rows)
    # MS Project cannot jump more than one outline level at a time
    fixed = []
    for i, lvl in enumerate(levels):
        prev = fixed[-1] if fixed else 0
        fixed.append(max(1, min(lvl, prev + 1)))
    return fixed


def build_outline(rows: list, project_name: str, cal: WorkCalendar, policy: dict,
                  group_header: str = "System", mode: Optional[str] = None, group_legend: Optional[dict] = None):
    """Return (project_node, all_nodes_in_order, info dict)."""
    info = {"dropped": [], "mode": None}
    unit = infer_duration_unit(rows, cal)
    for r in rows:
        if not r.dur_unit:
            r.dur_unit = unit
    mode = mode or choose_hierarchy_mode(rows)
    info["mode"] = mode
    info["duration_unit"] = unit

    project = Node(name=project_name, kind="project", level=1, label=project_name, dur_unit=unit)

    # ------------------------------------------------------------------ #
    if mode == "group":
        order = []
        groups = {}
        for r in rows:
            g = r.group or "Ungrouped"
            if g not in groups:
                groups[g] = []
                order.append(g)
            groups[g].append(r)
        prefix = (group_header.strip().title() + ": ") if policy.get("group_summary_prefix_from_column_header", True) and group_header else ""
        for g in order:
            desc = (group_legend or {}).get(g)
            if desc:
                word = group_header.strip().title() if group_header else ""
                gname = f"{word} {g} - {desc}".strip()
                glabel = f"{word} {g}".strip()
            else:
                gname, glabel = f"{prefix}{g}", g
            gnode = Node(name=gname, kind="summary", level=2, parent=project, label=glabel, dur_unit=unit)
            project.children.append(gnode)
            for r in groups[g]:
                n = _row_node(r, 3, gnode, unit)
                gnode.children.append(n)
    else:
        levels = _levels(rows, mode)
        stack = [project]
        nodes = []
        for r, lvl in zip(rows, levels):
            while len(stack) > lvl:
                stack.pop()
            parent = stack[-1]
            n = _row_node(r, lvl + 1, parent, unit)
            parent.children.append(n)
            nodes.append((n, lvl))
            stack.append(n)
        # summary = has children; dateless rows without children are dropped
        for n, _ in nodes:
            if n.children:
                n.kind = "summary"
            elif n.src.is_summary and not (n.src.start or n.src.finish):
                n.parent.children.remove(n)
                info["dropped"].append((n.src.excel_row, n.name, "heading without activities"))
            elif n.src.is_summary:
                n.kind = "milestone" if n.src.is_milestone else "task"
    # ------------------------------------------------------------------ #
    # completion milestones (firm policy)
    aliases = {k.upper(): v for k, v in policy.get("completion_milestone_label_aliases", {}).items()}
    suffix = policy.get("completion_milestone_suffix", "completed")

    def add_completion(summary: Node):
        for ch in list(summary.children):
            if ch.is_summary:
                add_completion(ch)
        if summary.kind == "project" and not policy.get("completion_milestone_for_project_title", False):
            return
        label = summary.label or summary.name
        if label == summary.name:                 # e.g. "System: PRG" -> "PRG", "Phase A - Substructure" -> "Phase A"
            m = re.match(r"^[A-Za-z][A-Za-z ]{1,20}:\s*(\S.*)$", label) or \
                re.match(r"^([A-Za-z]+ [A-Za-z0-9]{1,4}) - \S.*$", label)
            if m:
                label = m.group(1).strip()
        label = aliases.get(label.strip().upper(), label)
        cm = Node(name=f"{label} {suffix}", kind="completion", level=summary.level + 1,
                  parent=summary, label=label, dur_unit=summary.dur_unit)
        summary.children.append(cm)

    add_completion(project)

    ordered = list(walk(project))
    for i, n in enumerate(ordered, start=1):
        n.id = i
        n.uid = i
        n.order = i
    assign_wbs(project)
    set_targets(ordered, cal)
    return project, ordered, info


def _row_node(r: SourceRow, level: int, parent: Node, unit: str) -> Node:
    kind = "milestone" if r.is_milestone else "task"
    return Node(name=r.name, kind=kind, level=level, src=r, parent=parent,
                label=r.name, dur_unit=r.dur_unit or unit)


def walk(node: Node):
    yield node
    for ch in node.children:
        yield from walk(ch)


def assign_wbs(project: Node):
    project.wbs = "1"

    def rec(node: Node):
        for i, ch in enumerate(node.children, start=1):
            ch.wbs = f"{node.wbs}.{i}"
            ch.level = node.level + 1
            rec(ch)

    rec(project)


def set_targets(ordered: list, cal: WorkCalendar):
    """Translate Excel dates into working-day boundaries (targets)."""
    for n in ordered:
        r = n.src
        if r is None or n.is_summary:
            continue
        s = cal.start_b(r.start) if r.start else None
        f = cal.finish_b(r.finish) if r.finish else None
        if n.kind == "milestone":
            # milestone: sits at the end of its Excel window, except a project
            # start milestone (e.g. NTP) which sits at the very start
            if s is not None and s <= 0:
                n.t_start = n.t_finish = s
                n.ms_end_of_day = False
            else:
                b = f if f is not None else s
                n.t_start = n.t_finish = b
                n.ms_end_of_day = f is not None
            n.dur = 0
        else:
            if s is None and f is not None:
                s = f - 1
            if f is None and s is not None:
                f = s + 1
            n.t_start, n.t_finish = s, max(f, s + 1)
            n.dur = n.t_finish - n.t_start
