"""Faithful Python ports of the Node backend helpers used by the Excel
processing engine (utils/normalizeStatus.js + utils/timeUtils.js), plus
JavaScript-compatible number helpers so results match the original app
(Math.round rounds half UP, parseFloat does a leading-number parse).

Pure functions, no side effects."""

from __future__ import annotations

import datetime as _dt
import math
import re

MIN_PER_DAY = 24 * 60


# ── JavaScript number semantics ───────────────────────────────────────────
def js_round(x) -> int:
    """JavaScript Math.round: rounds .5 halves UP (Python round() would
    round half-to-even, giving different minutes on exact halves)."""
    return math.floor(x + 0.5)


def js_parse_float(v) -> float:
    """JavaScript parseFloat(): numbers pass through, strings are parsed for
    a leading float ("2.5 hrs" -> 2.5), anything else is NaN."""
    if isinstance(v, bool):
        return float("nan")
    if isinstance(v, (int, float)):
        return float(v)
    if v is None:
        return float("nan")
    m = re.match(r"^\s*[+-]?(?:\d+\.?\d*|\.\d+)(?:[eE][+-]?\d+)?", str(v))
    return float(m.group(0)) if m else float("nan")


# ── utils/normalizeStatus.js ──────────────────────────────────────────────
def normalize_status(raw) -> str:
    """Normalize SPST status values.

        WOP        -> WO
        WOP/WO     -> WO
        WO/WOP     -> WO
        PHP        -> PH
        PHP/PH     -> PH
        PH/PHP     -> PH

    Other slashed values (DP/ABS, ABS/DP, etc.) are preserved. Duplicated
    halves (X/X) collapse to X.
    """
    if raw is None:
        return ""
    s = re.sub(r"\s+", "", str(raw).strip().upper())
    if not s:
        return ""

    def _part(part: str) -> str:
        if part == "WOP":
            return "WO"
        if part == "PHP":
            return "PH"
        return part

    s = "/".join(_part(p) for p in s.split("/"))

    if "/" in s:
        parts = s.split("/")
        if len(parts) == 2 and parts[0] == parts[1]:
            return parts[0]

    return s


# ── utils/timeUtils.js ────────────────────────────────────────────────────
def parse_time_to_minutes(value):
    """Return minutes-since-midnight for a cell value, or None.

    Mirrors parseTimeToMinutes(): Excel decimal day-fractions in [0, 1) and
    "HH:MM[:SS] [AM/PM]" strings. openpyxl additionally hands back
    datetime.time / datetime.datetime objects for time-formatted cells
    (SheetJS keeps the raw serial number), so those are accepted too.
    """
    if value is None or value == "" or value == "--":
        return None

    # openpyxl converts time-formatted cells to time/datetime objects
    if isinstance(value, _dt.datetime):
        return value.hour * 60 + value.minute
    if isinstance(value, _dt.time):
        return value.hour * 60 + value.minute

    if isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        if 0 <= value < 1:
            return js_round(value * MIN_PER_DAY)
        return None

    s = str(value).strip()
    m = re.match(r"^(\d{1,2}):(\d{1,2})(?::\d{1,2})?\s*(AM|PM)?$", s, re.IGNORECASE)
    if not m:
        return None

    hours = int(m.group(1))
    mins = int(m.group(2))
    ampm = m.group(3).upper() if m.group(3) else None

    if ampm == "PM" and hours != 12:
        hours += 12
    if ampm == "AM" and hours == 12:
        hours = 0

    if hours < 0 or hours > 23 or mins < 0 or mins > 59:
        return None
    return hours * 60 + mins


def format_minutes_to_time(minutes) -> str:
    """"hh:mm AM/PM" from minutes-since-midnight (wraps across midnight)."""
    total = ((int(round(minutes)) % MIN_PER_DAY) + MIN_PER_DAY) % MIN_PER_DAY
    h = total // 60
    m = total % 60
    ampm = "PM" if h >= 12 else "AM"
    h = h % 12
    if h == 0:
        h = 12
    return f"{h:02d}:{m:02d} {ampm}"
