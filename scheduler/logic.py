"""Relationship (predecessor) logic.

Two sources of logic:
1. Predecessors already present in the spreadsheet (parsed and re-mapped to the
   new IDs).
2. Inference from the Excel dates when no logic exists. Every relationship is
   chosen so that MS Project's forward pass lands each activity on exactly the
   dates in the spreadsheet (zero discrepancy), while following good planning
   practice:
     * prefer links inside the same summary (WBS section / system),
     * prefer zero-lag FS, then SS, then FF,
     * leads (negative lags) only up to the policy limit (default 2 wks),
       otherwise a positive SS lag is used,
     * the earliest activity is the single anchor; every other activity has at
       least one predecessor (no open starts),
     * decision-gate milestones converge on every matching activity that
       finishes on the gate date (e.g. DG1 <- all six "DG1" reviews).
"""
from __future__ import annotations

import re
from typing import Optional

from .model import Link, Node

STOPWORDS = {
    "and", "of", "the", "to", "for", "with", "in", "on", "a", "an", "by", "at", "or", "all",
    "systems", "system", "wk", "wks", "week", "weeks", "day", "days", "business", "s", "first",
    "then", "issue", "phase", "stage", "out", "up", "per", "via", "into", "from", "is",
}


def _tokens(n: Node) -> set:
    toks = set(re.findall(r"[a-z0-9]+", n.name.lower()))
    p = n.parent
    while p is not None:                          # drop the group code (e.g. "ssb")
        toks -= set(re.findall(r"[a-z0-9]+", (p.label or "").lower()))
        p = p.parent
    return {t for t in toks if t not in STOPWORDS and not t.isdigit()}


def _sim(a: set, b: set) -> float:
    if not a or not b:
        return 0.0
    return len(a & b) / len(a | b)


def _key(n: Node):
    return (n.t_start, n.order)


def _fmt_lag(days: int, unit: str, dpw: int = 5) -> str:
    if days == 0:
        return ""
    sign = "+" if days > 0 else "-"
    d = abs(days)
    if unit == "w" and d % dpw == 0:
        w = d // dpw
        return f"{sign}{w} wk" if w == 1 else f"{sign}{w} wks"
    return f"{sign}{d} day" if d == 1 else f"{sign}{d} days"


def link_text(link: Link, unit: str = "w", dpw: int = 5) -> str:
    return f"{link.pred.id}{link.type}{_fmt_lag(link.lag, unit, dpw)}"


# --------------------------------------------------------------------------- #
# Inference
# --------------------------------------------------------------------------- #
def infer_logic(items: list, policy: dict, only: Optional[set] = None, dpw: int = 5):
    """Attach predecessors to every schedulable item (except the anchor)."""
    max_lead = int(round(float(policy.get("max_lead_weeks", 2)) * dpw))
    items = [n for n in items if n.t_start is not None]
    if not items:
        return None
    anchor = min(items, key=_key)
    toks = {n: _tokens(n) for n in items}

    for T in items:
        if T is anchor or (only is not None and T not in only):
            continue
        cands = [P for P in items if P is not T and _key(P) < _key(T)]
        if not cands:
            continue
        # planners read top-down: use rows listed above when they can drive T
        above = [P for P in cands if P.order < T.order]
        if above:
            cands = above
        day_one = T.t_start == anchor.t_start

        # 1) decision gates: converge on every matching activity finishing on the gate date
        if T.kind == "milestone":
            gate = [(P, _sim(toks[T], toks[P])) for P in cands
                    if P.t_finish == T.t_start and P.kind in ("task", "milestone")]
            gate = [(P, s) for P, s in gate if s > 0]
            if gate:
                top = max(s for _, s in gate)
                T.preds = [Link(P, "FS", 0, "gate convergence" if len(gate) > 1 else "inferred from Excel dates")
                           for P, s in gate if s >= 0.8 * top]
                T.preds.sort(key=lambda l: l.pred.order)
                continue

        nearest_same = None
        for P in cands:
            if P.parent is T.parent and (nearest_same is None or P.order > nearest_same.order):
                nearest_same = P

        best = None
        for P in cands:
            same = P.parent is T.parent
            # inside a summary: proximity decides; across summaries: meaning decides
            base = 50 if same else 10 * _sim(toks[T], toks[P])
            gap = abs(T.order - P.order)
            base -= 0.5 * gap if same else 0.05 * gap
            if P is nearest_same:
                base += 15
            options = []
            fs = T.t_start - P.t_finish
            ss = T.t_start - P.t_start
            if same:
                if fs >= -max_lead:
                    sc = base + 10 + (65 if fs == 0 else -3 * abs(fs) / 5 - (5 if fs < 0 else 0))
                    options.append((sc, "FS", fs))
                sc = base + 5 + (45 if ss == 0 else -3 * abs(ss) / 5)
                if P is anchor and ss == 0:
                    sc += 8
                options.append((sc, "SS", ss))
                if T.kind == "task" and P.kind == "task" and T.t_finish == P.t_finish:
                    options.append((base + 3 + 30, "FF", 0))
            else:
                # separate work fronts are only tied by true hand-overs (FS, no lag)
                # or by the project anchor (policy: kick-off activities SS to the anchor)
                if fs == 0:
                    options.append((base + 10 + 65, "FS", 0))
                if P is anchor:
                    sc = base + 5 + (45 + 8 + (40 if day_one else 0) if ss == 0 else -3 * abs(ss) / 5)
                    options.append((sc, "SS", ss))
                else:
                    # last resort (e.g. first activity of a section with no hand-over):
                    # heavily penalised so it never beats a sensible in-section link
                    if fs != 0 and fs >= -max_lead:
                        options.append((base + 10 - 3 * abs(fs) / 5 - 30, "FS", fs))
                    options.append((base + 5 + (45 if ss == 0 else -3 * abs(ss) / 5) - 30, "SS", ss))
            for sc, typ, lag in options:
                cand = (sc, -abs(lag), P.order, typ, lag, P)
                if best is None or cand[:3] > best[:3]:
                    best = cand
        sc, _, _, typ, lag, P = best
        T.preds = [Link(P, typ, lag, "inferred from Excel dates")]
    return anchor


# --------------------------------------------------------------------------- #
# Source predecessors
# --------------------------------------------------------------------------- #
_TOKEN = re.compile(
    r"^(?P<ref>.+?)\s*(?P<type>FS|SS|FF|SF)?\s*(?P<lag>[+-]\s*\d+(?:\.\d+)?\s*"
    r"(?:ed|d|day|days|w|wk|wks|week|weeks|h|hr|hrs|hour|hours|mo|mon|mons|month|months)?)?$",
    re.I,
)


def _lag_days(text: str, default_unit: str, dpw: int = 5) -> int:
    if not text:
        return 0
    m = re.match(r"([+-])\s*(\d+(?:\.\d+)?)\s*([a-z]*)", text.strip().lower())
    if not m:
        return 0
    v = float(m.group(2)) * (1 if m.group(1) == "+" else -1)
    u = m.group(3) or default_unit
    if u.startswith("w"):
        v *= dpw
    elif u.startswith("mo"):
        v *= round(dpw * 52 / 12)
    elif u.startswith("h"):
        v /= 8
    return int(round(v))


def apply_source_logic(nodes: list, default_unit: str = "d", dpw: int = 5):
    """Parse predecessor text found in the spreadsheet. Returns (linked set, problems)."""
    keys = {}
    srcs = [n for n in nodes if n.src is not None]
    # priority: activity ID column, then WBS code, then Excel row number
    for getter in (lambda s: s.src_id, lambda s: s.wbs, lambda s: str(s.excel_row)):
        for n in srcs:
            k = getter(n.src)
            if k:
                keys.setdefault(str(k).strip().upper(), n)
                if re.match(r"^\d+\.0$", str(k)):
                    keys.setdefault(str(k)[:-2], n)
    linked, problems = set(), []
    for n in nodes:
        if n.src is None or not n.src.preds or n.is_summary:
            continue
        links = []
        for tok in re.split(r"[,;\n]+", n.src.preds):
            tok = tok.strip()
            if not tok:
                continue
            m = _TOKEN.match(tok)
            ref = m.group("ref").strip().upper() if m else tok.upper()
            P = keys.get(ref)
            if P is None:
                problems.append((n.id, n.name, f"predecessor '{tok}' not found"))
                continue
            if P.is_summary:
                cm = next((c for c in P.children if c.kind == "completion"), None)
                if cm is None:
                    problems.append((n.id, n.name, f"predecessor '{tok}' is a summary - skipped"))
                    continue
                P = cm
            typ = (m.group("type") or "FS").upper() if m else "FS"
            lag = _lag_days(m.group("lag"), default_unit, dpw) if m else 0
            links.append(Link(P, typ, lag, "from spreadsheet"))
        if links:
            n.preds = links
            linked.add(n)
    return linked, problems


# --------------------------------------------------------------------------- #
# Policy links
# --------------------------------------------------------------------------- #
def link_completion_milestones(project: Node):
    """'<summary> completed' <- FS from the latest-finishing activity in the summary."""

    def finish_of(n: Node):
        if n.is_summary:
            cm = next((c for c in n.children if c.kind == "completion"), None)
            if cm is not None:
                return cm.t_finish, cm
            sub = [finish_of(c) for c in n.children]
            sub = [x for x in sub if x[0] is not None]
            return max(sub, key=lambda x: (x[0], x[1].order)) if sub else (None, None)
        return n.t_finish, n

    def rec(s: Node):
        for ch in s.children:
            if ch.is_summary:
                rec(ch)
        cm = next((c for c in s.children if c.kind == "completion"), None)
        if cm is None:
            return
        best = None
        for ch in s.children:
            if ch is cm:
                continue
            f, drv = finish_of(ch)
            if f is None or drv is None:
                continue
            if best is None or (f, ch.order) >= (best[0], best[2]):
                best = (f, drv, ch.order)
        if best:
            cm.t_start = cm.t_finish = best[0]
            cm.preds = [Link(best[1], "FS", 0, "policy: completion milestone")]
            cm.ms_end_of_day = not (best[1].kind in ("milestone", "completion") and not best[1].ms_end_of_day)

    rec(project)


def close_open_ends(ordered: list):
    """Optional: tie activities without successors to their summary's completion milestone."""
    succ = {n: 0 for n in ordered}
    for n in ordered:
        for l in n.preds:
            succ[l.pred] = succ.get(l.pred, 0) + 1
    added = 0
    for n in ordered:
        if not n.schedulable or n.kind == "completion" or succ.get(n, 0):
            continue
        p = n.parent
        while p is not None:
            cm = next((c for c in p.children if c.kind == "completion"), None)
            if cm is not None:
                cm.preds.append(Link(n, "FS", 0, "policy: close open end"))
                added += 1
                break
            p = p.parent
    return added
