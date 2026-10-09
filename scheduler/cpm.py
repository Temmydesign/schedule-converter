"""Forward/backward pass that mirrors MS Project auto-scheduling (ASAP, no constraints).

Used to *prove* the generated logic reproduces the Excel dates before the file
ever reaches MS Project, and to compute float / critical path for the QA report.
"""
from __future__ import annotations

from collections import defaultdict, deque

from .model import Node, WorkCalendar


class CycleError(Exception):
    pass


def topo_order(items: list) -> list:
    succ = defaultdict(list)
    indeg = {n: 0 for n in items}
    for n in items:
        for l in n.preds:
            if l.pred in indeg:
                succ[l.pred].append(n)
                indeg[n] += 1
    q = deque(sorted([n for n in items if indeg[n] == 0], key=lambda x: x.order))
    out = []
    while q:
        n = q.popleft()
        out.append(n)
        for s in succ[n]:
            indeg[s] -= 1
            if indeg[s] == 0:
                q.append(s)
    if len(out) != len(items):
        stuck = [n for n in items if indeg[n] > 0]
        raise CycleError("Circular logic between: " + ", ".join(f"{n.id} {n.name}" for n in stuck[:8]))
    return out


def forward_pass(ordered: list, project_start_b: int = 0):
    items = [n for n in ordered if n.schedulable]
    for n in topo_order(items):
        contrib = []
        for l in n.preds:
            p = l.pred
            if p.ef is None:
                continue
            if l.type == "FS":
                v = p.ef + l.lag
                flag = (p.kind not in ("milestone", "completion")) or p.ms_end_of_day
            elif l.type == "SS":
                v = p.es + l.lag
                flag = p.is_milestone and p.ms_end_of_day
            elif l.type == "FF":
                v = p.ef + l.lag - n.dur
                flag = True
            else:  # SF
                v = p.es + l.lag - n.dur
                flag = False
            contrib.append((v, flag))
        es = max([project_start_b] + [v for v, _ in contrib])
        end_flag = any(f for v, f in contrib if v == es)
        n.es = es
        n.ef = es + n.dur
        if n.is_milestone:
            n.ms_end_of_day = end_flag if n.preds else False
    rollup(ordered)


def rollup(ordered: list):
    for n in reversed(ordered):
        if n.is_summary:
            kids = [c for c in n.children if c.es is not None]
            if kids:
                n.es = min(c.es for c in kids)
                n.ef = max(c.ef for c in kids)
                n.dur = n.ef - n.es


def backward_pass(ordered: list):
    """Total float (working days) for every schedulable item."""
    items = [n for n in ordered if n.schedulable]
    order = topo_order(items)
    proj_ef = max(n.ef for n in items)
    succ = defaultdict(list)
    for n in items:
        for l in n.preds:
            succ[l.pred].append((n, l))
    ls, lf = {}, {}
    for n in reversed(order):
        f = proj_ef
        for s, l in succ[n]:
            if l.type == "FS":
                f = min(f, ls[s] - l.lag)
            elif l.type == "SS":
                f = min(f, ls[s] - l.lag + n.dur)
            elif l.type == "FF":
                f = min(f, lf[s] - l.lag)
            else:
                f = min(f, lf[s] - l.lag + n.dur)
        lf[n] = f
        ls[n] = f - n.dur
    return {n: ls[n] - n.es for n in items}


def display_dates(n: Node, cal: WorkCalendar):
    """(start_date, finish_date) as MS Project will display them."""
    if n.es is None:
        return None, None
    if n.is_summary:
        spans = [display_dates(c, cal) for c in n.children]
        spans = [s for s in spans if s[0] is not None]
        if not spans:
            return None, None
        return min(s for s, _ in spans), max(f for _, f in spans)
    if n.is_milestone:
        d = cal.date(n.es - 1) if (n.ms_end_of_day and n.es > 0) else cal.date(n.es)
        return d, d
    return cal.date(n.es), cal.date(n.ef - 1)
