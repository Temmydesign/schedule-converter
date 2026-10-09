"""Core data structures and the working-time calendar used by the engine.

Time model
----------
All scheduling maths is done on *working-day boundaries* (integers).
Boundary ``b`` is the start of working day ``b`` (08:00) which is the same
instant, in working time, as the end of working day ``b-1`` (17:00).

* A task occupying working days ``s .. s+d-1`` has start boundary ``s`` and
  finish boundary ``s+d``.
* A milestone sits on a single boundary.

This is exactly how MS Project's CPM engine treats a standard calendar, so a
schedule that reproduces the boundaries here reproduces the dates in MS Project.
"""
from __future__ import annotations

import datetime as dt
import json
import os
from dataclasses import dataclass, field
from typing import Optional

import numpy as np

DAY_NAMES = ["Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"]


# --------------------------------------------------------------------------- #
# Policy
# --------------------------------------------------------------------------- #
_POLICY_PATH = os.path.join(os.path.dirname(os.path.dirname(__file__)), "firm_policy.json")


def load_policy(path: str = _POLICY_PATH) -> dict:
    with open(path, "r", encoding="utf-8") as fh:
        data = json.load(fh)
    return {k: v for k, v in data.items() if not k.startswith("_")}


# --------------------------------------------------------------------------- #
# Calendar
# --------------------------------------------------------------------------- #
class WorkCalendar:
    """Working-day calendar (default Mon-Fri, optional holidays)."""

    def __init__(self, origin: dt.date, working_days=None, holidays=None, hours_per_day: float = 8.0):
        working_days = working_days or DAY_NAMES[:5]
        self.weekmask = [1 if d in working_days else 0 for d in DAY_NAMES]
        self.holidays = sorted({h for h in (holidays or [])})
        self.hours_per_day = float(hours_per_day)
        self._bdc = np.busdaycalendar(weekmask=self.weekmask,
                                      holidays=[np.datetime64(h) for h in self.holidays])
        self.origin = self.roll_forward(origin)

    # -- helpers ------------------------------------------------------------ #
    def is_working(self, d: dt.date) -> bool:
        return bool(np.is_busday(np.datetime64(d), busdaycal=self._bdc))

    def roll_forward(self, d: dt.date) -> dt.date:
        return np.busday_offset(np.datetime64(d), 0, roll="forward", busdaycal=self._bdc).astype(dt.date)

    def idx(self, d: dt.date) -> int:
        """Working-day index of date *d* (non-working dates roll forward)."""
        d = self.roll_forward(d)
        return int(np.busday_count(np.datetime64(self.origin), np.datetime64(d), busdaycal=self._bdc))

    def start_b(self, d: dt.date) -> int:
        """Boundary at the start of date d (= working days before d)."""
        return int(np.busday_count(np.datetime64(self.origin), np.datetime64(d), busdaycal=self._bdc))

    def finish_b(self, d: dt.date) -> int:
        """Boundary at the end of date d (= working days up to and including d)."""
        return int(np.busday_count(np.datetime64(self.origin), np.datetime64(d + dt.timedelta(days=1)),
                                   busdaycal=self._bdc))

    def date(self, i: int) -> dt.date:
        return np.busday_offset(np.datetime64(self.origin), int(i), roll="forward",
                                busdaycal=self._bdc).astype(dt.date)

    def working_days_between(self, start: dt.date, finish: dt.date) -> int:
        """Inclusive count of working days from start to finish."""
        return self.idx(finish) + (1 if self.is_working(finish) else 0) - self.idx(start)


# --------------------------------------------------------------------------- #
# Source rows (what was read from Excel)
# --------------------------------------------------------------------------- #
@dataclass
class SourceRow:
    excel_row: int
    name: str
    start: Optional[dt.date] = None
    finish: Optional[dt.date] = None
    duration: Optional[float] = None
    dur_unit: str = "w"              # "w" weeks or "d" days (unit of the source duration)
    wbs: Optional[str] = None
    type: Optional[str] = None
    group: Optional[str] = None
    src_id: Optional[str] = None
    preds: Optional[str] = None
    indent: int = 0
    bold: bool = False
    is_milestone: bool = False
    is_summary: bool = False         # summary per the source (Type=Sum, WBS parent, section header)


# --------------------------------------------------------------------------- #
# Output schedule
# --------------------------------------------------------------------------- #
@dataclass
class Link:
    pred: "Node"
    type: str = "FS"                 # FS | SS | FF | SF
    lag: int = 0                     # working days (negative = lead)
    reason: str = ""


@dataclass
class Node:
    name: str
    kind: str                        # project | summary | task | milestone | completion
    level: int = 1                   # MS Project outline level
    src: Optional[SourceRow] = None
    parent: Optional["Node"] = None
    children: list = field(default_factory=list)
    preds: list = field(default_factory=list)   # list[Link]
    label: str = ""                  # short label used for "<label> completed"
    dur_unit: str = "w"

    # targets from Excel (boundaries) ------------------------------------- #
    t_start: Optional[int] = None
    t_finish: Optional[int] = None
    dur: int = 0                     # working days

    # results ------------------------------------------------------------- #
    id: int = 0
    uid: int = 0
    wbs: str = ""
    es: Optional[int] = None
    ef: Optional[int] = None
    ms_end_of_day: bool = True       # milestone shown at end of the previous day
    order: int = 0

    @property
    def is_summary(self) -> bool:
        return self.kind in ("project", "summary")

    @property
    def is_milestone(self) -> bool:
        return self.kind in ("milestone", "completion")

    @property
    def schedulable(self) -> bool:
        return not self.is_summary

    def __hash__(self):
        return id(self)

    def __eq__(self, other):
        return self is other

    def __repr__(self):
        return f"<{self.kind} {self.id}: {self.name[:40]}>"
