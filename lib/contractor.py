"""Contractor attendance processing engine (Alqaswa-style paysheet books).

The OT, check-in, check-out and rest rules are the ones the HR engine uses —
OT_CONFIG and distribute_ot are imported from lib/process.py rather than
re-stated, so the two tabs can never drift apart. What differs is only the
workbook *shape*:

1.  The attendance grid is not the first sheet. Every sheet is scanned for a
    header row carrying SPST + ARRV, so "Attendance Report" is found wherever
    it sits and whatever it is called.
2.  The monthly OT total does not live in the attendance grid — it comes from
    the "Wages Register" sheet's "OT Hrs" column, matched per employee on
    (Employee Code, Employee Name). The pair is the key because a contractor
    code is not unique: RCC000 covers two different people. An empty "OT Hrs"
    simply means no OT that month.
2a. Header spellings vary between contractors ("Employee Code" vs "Emp Code",
    "WORK" vs "Work"), so columns are matched on a normalized name.
3.  SHIFT IN / SHIFT OUT are dotted text ("9.00 AM", "7.00 PM "), so times are
    read with parse_time_loose().
4.  Punches are read as text in either notation and written back as text:
    ARRV / DEPT as 12-hour clock readings ("09:02 AM", "06:31 PM"), WORK as a
    plain HH:MM duration ("09:29") since a span has no AM/PM.
5.  A qualifying day's OT is written back into the attendance "OT Hours"
    column as the per-day figure, so the column sums to the Wages Register
    total for that employee.

Rules applied per row, matching lib/process.py:

*   SPST is normalized in place; anything mentioning WO collapses to "WO".
*   WO / PH / ABS / PL: the person did not work, so ARRV, DEPT, OT Hours and
    WORK are all left genuinely empty — no 00:00, no 0, nothing invented for a
    day nobody was there.
*   ABS/DP and DP/ABS half-days: punches kept, a span over 9.5 h trimmed to
    9.5 h minus 1..30 min by moving DEPT (ABS/DP) or ARRV (DP/ABS).
*   DP days with a full shift: ARRV = shift-in + 0..20 min; DEPT keeps a
    recorded punch within +/-60 min of shift-out, else snaps to shift-out; a
    base span over 9.5 h is trimmed to 9.5 h minus 0..30 min.
*   The employee's monthly OT is split over the qualifying days in 1..2 h
    chunks (distribute_ot), a fractional .5 riding on one day as 1 h 30 m.
*   On an OT day the trimmed departure is pushed later by exactly the OT
    granted, so WORK = capped base + OT. The 9.5 h rest cap therefore still
    holds for the ordinary part of the day even on a 10 h shift.

Returns the processed workbook bytes plus a stats dict.
"""

from __future__ import annotations

import io
import math
import random
import re

from openpyxl import load_workbook

from .process import OT_CONFIG, _normalize_key, _rand_int, distribute_ot
from .utils import (
    MIN_PER_DAY,
    format_minutes_to_time,
    js_parse_float,
    js_round,
    normalize_status,
    parse_time_loose,
)

# How far down a sheet to look for its header row before giving up.
_HEADER_SCAN_ROWS = 30

# Statuses that mean the person did not work that day. ARRV, DEPT, OT Hours
# and WORK are all left genuinely empty for these rows — nothing is invented
# for a day nobody was there, not even a 00:00.
NON_WORKING = ("WO", "PH", "ABS", "PL")

# Header spellings seen across contractor books for the same two columns.
_CODE_KEYS = ("employeecode", "empcode", "code")
_NAME_KEYS = ("employeename", "empname", "name")


def _pick(hm, keys):
    """First of `keys` present in a header map."""
    for k in keys:
        if k in hm:
            return hm[k]
    return None


def _match_key(value) -> str:
    """Collapse whitespace and case so 'Rohit Karbele ' and 'Rohit Karbele'
    are the same person across the two sheets."""
    return re.sub(r"\s+", " ", str(value if value is not None else "")).strip().upper()


def _header_map(ws, row) -> dict:
    """{normalized header name -> column index} for one row."""
    out = {}
    for c in range(1, (ws.max_column or 0) + 1):
        v = ws.cell(row=row, column=c).value
        if v is not None and str(v).strip() != "":
            out.setdefault(_normalize_key(v), c)
    return out


def _scan_for_header(ws, predicate):
    """First row within the scan window whose header map satisfies predicate."""
    if not ws.max_row:
        return None, None
    first = ws.min_row or 1
    last = min(first + _HEADER_SCAN_ROWS, ws.max_row)
    for r in range(first, last + 1):
        hm = _header_map(ws, r)
        if hm and predicate(hm):
            return r, hm
    return None, None


def _sheet_order(wb, hints, skip_title=None):
    """Sheet titles, the ones named like what we want tried first, so a
    workbook with two plausible sheets resolves the obvious way."""
    names = [n for n in wb.sheetnames if n != skip_title]
    preferred = [n for n in names if any(h in n.lower() for h in hints)]
    return preferred + [n for n in names if n not in preferred]


def _find_attendance(wb):
    """The sheet holding the per-day attendance grid."""

    def looks_like_attendance(hm):
        has_status = "spst" in hm or "status" in hm
        has_punch = any(k in hm for k in ("arrv", "arrival", "entry", "dept"))
        return has_status and has_punch

    for name in _sheet_order(wb, ("atten", "muster")):
        ws = wb[name]
        row, hm = _scan_for_header(ws, looks_like_attendance)
        if row is not None:
            return ws, row, hm
    return None, None, None


def _find_wages(wb, skip_title):
    """The sheet holding the monthly OT totals."""

    def looks_like_wages(hm):
        has_code = any(k in hm for k in _CODE_KEYS)
        has_ot = any(k in hm for k in ("othrs", "othours", "othour", "ot"))
        return has_code and has_ot

    for name in _sheet_order(wb, ("wage", "pay", "salary"), skip_title):
        ws = wb[name]
        row, hm = _scan_for_header(ws, looks_like_wages)
        if row is not None:
            return ws, row, hm
    return None, None, None


def _read_wages_ot(ws, hrow, hm, ws_vals):
    """Monthly OT hours per employee from the wages sheet.

    Returns (by_code_and_name, by_code). Values are read from the cached
    results when the cell holds a formula, so a computed "OT Hrs" column still
    works.
    """
    c_code = _pick(hm, _CODE_KEYS)
    c_name = _pick(hm, _NAME_KEYS)
    c_ot = hm.get("othrs") or hm.get("othours") or hm.get("othour") or hm.get("ot")

    by_pair: dict[tuple[str, str], float] = {}
    by_code: dict[str, float] = {}
    if c_code is None or c_ot is None:
        return by_pair, by_code

    for r in range(hrow + 1, (ws.max_row or hrow) + 1):
        code_raw = ws.cell(row=r, column=c_code).value
        if code_raw is None or str(code_raw).strip() == "":
            continue  # totals row / spacer

        raw = ws.cell(row=r, column=c_ot).value
        if isinstance(raw, str) and raw.startswith("=") and ws_vals is not None:
            raw = ws_vals.cell(row=r, column=c_ot).value
        ot = js_parse_float(raw)
        if math.isnan(ot) or ot < 0:
            ot = 0.0

        ck = _match_key(code_raw)
        nk = _match_key(ws.cell(row=r, column=c_name).value if c_name else "")
        by_pair[(ck, nk)] = by_pair.get((ck, nk), 0.0) + ot
        by_code[ck] = by_code.get(ck, 0.0) + ot

    return by_pair, by_code


def _fmt_clock(minutes) -> str:
    """A punch time as 12-hour text, e.g. "06:31 PM".

    Only ARRV and DEPT go through here. WORK is a duration, not a clock
    reading, so it keeps the plain HH:MM form via _fmt_span().
    """
    return format_minutes_to_time(js_round(minutes))


def _fmt_span(minutes) -> str:
    """A worked span as "HH:MM" text, e.g. "10:07"."""
    m = js_round(minutes)
    if m < 0:
        m += MIN_PER_DAY
    if m < 0:
        m = 0
    return f"{m // 60:02d}:{m % 60:02d}"


def process_contractor_workbook(file_bytes: bytes, filename: str = ""):
    """Process a contractor paysheet workbook. Returns (xlsx_bytes, stats)."""
    name = (filename or "").lower()
    if name.endswith(".xls") or name.endswith(".csv"):
        raise ValueError(
            "Contractor paysheets must be .xlsx or .xlsm — the monthly OT is "
            "read from the Wages Register sheet, which a .xls/.csv conversion "
            "would drop. Re-save the file as .xlsx and upload it again."
        )

    wb = load_workbook(io.BytesIO(file_bytes))  # keeps formulas intact on save
    try:
        wb_vals = load_workbook(io.BytesIO(file_bytes), data_only=True)
    except Exception:  # noqa: BLE001 — cached values are a nicety, not required
        wb_vals = None

    ws, header_row, hm = _find_attendance(wb)
    if ws is None:
        raise ValueError(
            "No attendance sheet found. Expected a sheet with SPST and "
            "ARRV / DEPT columns (for example 'Attendance Report')."
        )

    col = {
        "code": _pick(hm, _CODE_KEYS),
        "name": _pick(hm, _NAME_KEYS),
        "spst": hm.get("spst") or hm.get("status"),
        "shiftIn": hm.get("shiftin"),
        "shiftOut": hm.get("shiftout"),
        "arrv": hm.get("arrv") or hm.get("arrival") or hm.get("entry"),
        "dept": hm.get("dept"),
        "work": hm.get("work"),
        "ot": hm.get("othours") or hm.get("othrs"),
    }
    if col["spst"] is None:
        raise ValueError("No SPST / Status column found in the attendance sheet.")

    wages_ws, wages_row, wages_hm = _find_wages(wb, ws.title)
    if wages_ws is None:
        raise ValueError(
            "No wages sheet found. Expected a sheet with 'Employee Code' and "
            "'OT Hrs' columns (for example 'Wages Register') holding the "
            "monthly OT total per employee."
        )
    ot_by_pair, ot_by_code = _read_wages_ot(
        wages_ws,
        wages_row,
        wages_hm,
        wb_vals[wages_ws.title] if wb_vals is not None and wages_ws.title in wb_vals.sheetnames else None,
    )

    # ── cell helpers ──────────────────────────────────────────────────────
    def cell_at(R, C):
        return None if C is None else ws.cell(row=R, column=C)

    def read_min(R, C):
        c = cell_at(R, C)
        return parse_time_loose(c.value) if c is not None else None

    def write_clock(R, C, minutes):
        if C is None:
            return
        c = ws.cell(row=R, column=C)
        c.value = _fmt_clock(minutes)
        c.number_format = "@"

    def write_work(R, minutes):
        if col["work"] is None:
            return
        c = ws.cell(row=R, column=col["work"])
        c.value = _fmt_span(minutes)
        c.number_format = "@"

    def write_ot(R, hours):
        if col["ot"] is None:
            return
        c = ws.cell(row=R, column=col["ot"])
        c.value = round(hours, 2)
        c.number_format = "0.0"

    def clear_cell(R, C):
        """Leave the cell genuinely empty — no 0, no "00:00", no blank string."""
        if C is None:
            return
        c = ws.cell(row=R, column=C)
        c.value = None

    arrv_fixed = 0
    dept_fixed = 0
    spst_normalized = 0
    work_updated = 0
    blanked_rows = 0

    employees: dict[tuple[str, str], dict] = {}

    # ── First pass: normalize each row, collect the OT-managed days ───────
    for R in range(header_row + 1, (ws.max_row or header_row) + 1):
        spst_cell = cell_at(R, col["spst"])
        if spst_cell is None or spst_cell.value is None:
            continue

        status = str(spst_cell.value or "").upper().strip()
        ns = normalize_status(status)
        if "WO" in ns:
            ns = "WO"

        if ns and ns != status:
            spst_cell.value = ns
            spst_normalized += 1

        # Weekly off / holiday / absent / paid leave — the person did not work,
        # so all four generated columns stay empty rather than being filled
        # with zeroes.
        if ns in NON_WORKING:
            for C in (col["arrv"], col["dept"], col["ot"], col["work"]):
                clear_cell(R, C)
            blanked_rows += 1
            continue

        si_min = read_min(R, col["shiftIn"])
        so_min = read_min(R, col["shiftOut"])

        # Half-days: keep the recorded punches, trim anything over 9.5 h
        if ns in ("ABS/DP", "DP/ABS"):
            a = read_min(R, col["arrv"])
            d = read_min(R, col["dept"])
            write_ot(R, 0)
            if a is not None and d is not None:
                w = d - a
                if w < 0:
                    w += MIN_PER_DAY
                if w > OT_CONFIG["trimTargetMin"]:
                    target = OT_CONFIG["trimTargetMin"] - _rand_int(
                        1, OT_CONFIG["noOtTrimRandomMin"]
                    )
                    if ns == "ABS/DP":
                        write_clock(R, col["dept"], a + target)
                    else:
                        write_clock(R, col["arrv"], d - target)
                    w = target
                write_work(R, w)
            else:
                write_work(R, 0)
            continue

        eligible = ("DP" in ns) and si_min is not None and so_min is not None

        if eligible:
            ck = _match_key(cell_at(R, col["code"]).value if col["code"] else "")
            nk = _match_key(cell_at(R, col["name"]).value if col["name"] else "")
            key = (ck, nk)
            e = employees.get(key)
            if e is None:
                e = {"days": [], "code": ck, "name": nk}
                employees[key] = e

            W = OT_CONFIG["naturalWindowMin"]
            orig_dep = read_min(R, col["dept"])
            shift_len = ((so_min - si_min) % MIN_PER_DAY + MIN_PER_DAY) % MIN_PER_DAY

            # check-in: shift-in + 0..20 random minutes
            arr = si_min + _rand_int(0, OT_CONFIG["arrvMaxLateMin"])
            write_clock(R, col["arrv"], arr)
            arrv_fixed += 1

            # check-out: keep a natural punch near shift-out, else snap to it
            dep = orig_dep if (orig_dep is not None and abs(orig_dep - so_min) <= W) else so_min

            base_worked = dep - arr
            if base_worked < 0:
                base_worked += MIN_PER_DAY

            # OT only on a pure DP day whose recorded work is not already long
            ot_ok = ns == "DP" and base_worked <= shift_len + OT_CONFIG["otWorkedSlackMin"]

            # rest cap: the ordinary part of the day never exceeds 9.5 h
            if base_worked > OT_CONFIG["trimTargetMin"]:
                base_worked = OT_CONFIG["trimTargetMin"] - _rand_int(
                    0, OT_CONFIG["noOtTrimRandomMin"]
                )

            e["days"].append({"R": R, "arr": arr, "base": base_worked, "otOk": ot_ok})
        else:
            # PL, OD, LWP, unknown… — same treatment as the HR engine
            orig_arr = read_min(R, col["arrv"])
            orig_dep = read_min(R, col["dept"])
            write_ot(R, 0)

            a = orig_arr
            if si_min is not None and (orig_arr is not None or orig_dep is not None):
                a = si_min + _rand_int(0, OT_CONFIG["arrvMaxLateMin"])
                write_clock(R, col["arrv"], a)
                arrv_fixed += 1

            d = orig_dep
            if a is not None and d is not None:
                w = d - a
                if w < 0:
                    w += MIN_PER_DAY
                if w > OT_CONFIG["trimTargetMin"]:
                    target = OT_CONFIG["trimTargetMin"] - _rand_int(
                        0, OT_CONFIG["noOtTrimRandomMin"]
                    )
                    write_clock(R, col["dept"], a + target)
                    w = target
                write_work(R, w)
            else:
                write_work(R, 0)

    # ── Second pass: spread each employee's monthly OT over their days ────
    ot_assigned = 0.0
    matched = 0
    shortfall = 0.0
    unmatched: list[str] = []
    widened: list[str] = []

    for key, e in employees.items():
        ck, nk = key
        if key in ot_by_pair:
            ot_raw = ot_by_pair[key]
            matched += 1
        elif ck in ot_by_code:
            ot_raw = ot_by_code[ck]
            matched += 1
        else:
            ot_raw = 0.0
            unmatched.append(f"{ck} / {nk}".strip(" /"))

        int_hours = math.floor(ot_raw)
        frac_min = js_round((ot_raw - int_hours) * 60)

        ot_days = [d for d in e["days"] if d["otOk"]]

        # The strict pool is the HR rule: a pure DP day whose recorded span is
        # not already longer than the shift. Re-processing a sheet that already
        # carries OT shrinks that pool, because the recorded departures include
        # the OT. Widen to every DP day rather than quietly drop paid hours.
        if ot_raw > len(ot_days) * OT_CONFIG["maxOtHours"]:
            ot_days = list(e["days"])
            widened.append(f"{ck} / {nk}".strip(" /"))

        capacity = len(ot_days) * OT_CONFIG["maxOtHours"]
        if ot_raw > capacity:
            shortfall += ot_raw - capacity

        day_ot_min = [0] * len(ot_days)
        pool = list(range(len(ot_days)))

        half_placed = False
        if frac_min > 0 and int_hours >= 1 and len(pool) >= 1:
            half_idx = pool[random.randrange(len(pool))]
            day_ot_min[half_idx] = 60 + frac_min
            int_hours -= 1
            pool = [i for i in pool if i != half_idx]
            half_placed = True

        alloc = distribute_ot(
            int_hours,
            len(pool),
            OT_CONFIG["minOtHours"],
            OT_CONFIG["maxOtHours"],
            OT_CONFIG["jitter"],
        )
        for j in range(len(pool)):
            day_ot_min[pool[j]] += alloc[j] * 60

        if frac_min > 0 and not half_placed and len(ot_days) >= 1:
            day_ot_min[random.randrange(len(ot_days))] += frac_min

        for i, d in enumerate(ot_days):
            d["otMin"] = day_ot_min[i]

        for d in e["days"]:
            ot_min = d.get("otMin", 0) or 0
            worked = d["base"] + ot_min
            dep = d["arr"] + worked

            write_clock(d["R"], col["dept"], dep)
            dept_fixed += 1
            write_work(d["R"], worked)
            write_ot(d["R"], ot_min / 60.0)
            work_updated += 1
            ot_assigned += ot_min / 60.0

    out = io.BytesIO()
    wb.save(out)
    out.seek(0)

    stats = {
        "employees": len(employees),
        "arrvFixed": arrv_fixed,
        "deptFixed": dept_fixed,
        "spstNormalized": spst_normalized,
        "workUpdated": work_updated,
        "blankedRows": blanked_rows,
        "otHours": round(ot_assigned, 2),
        "matched": matched,
        "unmatched": unmatched,
        "widened": widened,
        "shortfall": round(shortfall, 2),
        "attendanceSheet": ws.title,
        "wagesSheet": wages_ws.title,
    }
    return out.getvalue(), stats
