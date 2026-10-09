"""Work-week detection (Mon-Fri / Mon-Sat / Mon-Sun) from the workbook itself.

Evidence, strongest first:
1. Activity dates that fall on Saturdays / Sundays.
2. Day-based durations that only fit a 5-, 6- or 7-day week.
3. Wording anywhere in the workbook ("6-day week", "Mon-Sat", "24/7", "business days" ...).
Falls back to the firm default when nothing points elsewhere.
"""
from __future__ import annotations

import datetime as dt
import re
from collections import Counter

import numpy as np

WEEKS = {
    5: ["Mon", "Tue", "Wed", "Thu", "Fri"],
    6: ["Mon", "Tue", "Wed", "Thu", "Fri", "Sat"],
    7: ["Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"],
}
LABEL = {5: "Mon–Fri", 6: "Mon–Sat", 7: "Mon–Sun"}

_TEXT = {
    7: re.compile(r"(\b7[\s-]?day\s*(work)?\s*week|\bseven[\s-]?day|\b24\s*/\s*7\b|\b7\s*days?\s*(a|per|/)\s*w(ee)?k|"
                  r"\bmon(day)?\s*(-|–|to)\s*sun(day)?\b|\bcalendar\s+days\b|\b7d\s*/\s*w)", re.I),
    6: re.compile(r"(\b6[\s-]?day\s*(work)?\s*week|\bsix[\s-]?day|\b6\s*days?\s*(a|per|/)\s*w(ee)?k|"
                  r"\bmon(day)?\s*(-|–|to)\s*sat(urday)?\b|\b6d\s*/\s*w)", re.I),
    5: re.compile(r"(\b5[\s-]?day\s*(work)?\s*week|\bfive[\s-]?day|\b5\s*days?\s*(a|per|/)\s*w(ee)?k|"
                  r"\bmon(day)?\s*(-|–|to)\s*fri(day)?\b|\bbusiness\s+days?\b|\b5d\s*/\s*w)", re.I),
}


def week_days(n: int) -> list:
    return list(WEEKS[n])


def detect_work_week(sheets: list, rows: list, default_days: list):
    """Return (n_days, reason, evidence_level). evidence_level: dates | durations | text | default."""
    default_n = len(default_days) if len(default_days) in WEEKS else 5

    # 1) weekend dates on activities
    dates = []
    for r in rows:
        if r.is_summary:
            continue
        for d in (r.start, r.finish):
            if d:
                dates.append(d)
    wd = Counter(d.weekday() for d in dates)
    n = max(len(dates), 1)
    sat, sun = wd.get(5, 0), wd.get(6, 0)
    # a single stray weekend date is treated as a typo (flagged in Quality checks), not a calendar
    if sat + sun >= 2 and (sat + sun) / n >= 0.03:
        if sun:
            return 7, f"{sun} activity date{'s' if sun > 1 else ''} fall on a Sunday", "dates"
        return 6, f"{sat} activity dates fall on a Saturday, none on a Sunday", "dates"

    # 2) day-based durations
    fits = Counter()
    tested = 0
    for r in rows:
        if r.is_summary or r.is_milestone or not (r.start and r.finish) or r.duration in (None, 0):
            continue
        if (r.dur_unit or "d") != "d":
            continue
        tested += 1
        for k in (5, 6, 7):
            mask = [1] * k + [0] * (7 - k)
            cnt = int(np.busday_count(np.datetime64(r.start), np.datetime64(r.finish + dt.timedelta(days=1)), weekmask=mask))
            if abs(cnt - r.duration) < 0.01:
                fits[k] += 1
    if tested >= 3 and fits:
        best, hits = fits.most_common(1)[0]
        if hits >= 0.6 * tested and list(fits.values()).count(hits) == 1 and best != default_n:
            return best, f"{hits} of {tested} day-based durations only fit a {LABEL[best]} week", "durations"

    # 3) wording in the workbook
    hits = Counter()
    for sh in sheets:
        for row in sh.values[:400]:
            for v in row:
                if isinstance(v, str) and 4 <= len(v) <= 400:
                    for k, pat in _TEXT.items():
                        if pat.search(v):
                            hits[k] += 1
    for k in (7, 6, 5):
        if hits.get(k) and hits[k] >= max(hits.get(j, 0) for j in (5, 6, 7)):
            if k == default_n:
                break
            return k, f"workbook text refers to a {LABEL[k]} working week", "text"
    if hits.get(default_n):
        return default_n, f"workbook text refers to a {LABEL[default_n]} working week", "text"
    return default_n, f"no weekend work found; firm default {LABEL[default_n]}", "default"
