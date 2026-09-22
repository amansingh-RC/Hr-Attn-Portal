# Royal Chain HR Processing Tool — Streamlit

A **database-free** tool with two tabs: upload an Excel → it is processed →
download the updated file. That's it.

| Tab | Input | Engine |
|-----|-------|--------|
| **🏢 HR Dashboard** | the monthly attendance export (`.xlsx`, `.xls`, `.csv`) | [`lib/process.py`](lib/process.py) |
| **👷 Contractor** | a contractor paysheet workbook (`.xlsx`, `.xlsm`) | [`lib/contractor.py`](lib/contractor.py) |

The whole data-manipulation engine of the original backend
(`routes/process.js`) is ported 1:1 to Python in [`lib/process.py`](lib/process.py).
The Contractor tab **imports `OT_CONFIG` and `distribute_ot` from that module**
rather than restating them, so the OT, check-in, check-out and rest rules can
never drift apart between the two tabs.

## What processing does (identical to the original)

| Rule | Behaviour |
|------|-----------|
| **Column detection** | Header names matched ignoring case/spaces/`_`/`/`. `DEPT` is the *departure time* column only when a separate `Department` column exists. Works with `SPST`/`Status`, `ARRV`/`Arrival`/`Entry`, `Employee Code`/`Code`/`Employee Name`/`Name`, `Shift In`, `Shift Out`, `Work`, `OT Hours`. |
| **SPST normalization** | `WOP → WO`, `PHP → PH`, `WOP/WO → WO`, `PH/PHP → PH`, `X/X → X` — rewritten in place. |
| **WO / PH / ABS rows** | ARRV & DEPT blanked, WORK → 0, OT Hours → 0. |
| **ABS/DP & DP/ABS half-days** | Recorded punches kept; a span over 9.5 h is trimmed to 9.5 h − 1..30 min by moving DEPT (ABS/DP) or ARRV (DP/ABS); WORK recomputed. |
| **DP rows** (status contains `DP`, shift defined) | Arrival rewritten to *shift-in + 0..20 random min*; departure keeps a natural punch within ±60 min of shift-out, else snaps to shift-out; spans over 9.5 h trimmed by 0..30 min. |
| **OT distribution** | Per employee, the monthly `OT Hours` total is split across pure-`DP` days (whose base span ≤ shift + 10 min) in 1–2 h chunks with jittered weights; a fractional `.5` rides on one day as 1 h 30 m. On OT days `DEPT = ARRV + shift length + OT`. |
| **Other rows** (PL, unknown…) | Arrival randomized the same way when a shift-in and any punch exist; WORK recomputed, spans over 9.5 h trimmed. |
| **Cell formats** | Times stored as Excel decimal day-fractions (`h:mm AM/PM`), WORK as `[h]:mm` — same as the SheetJS output. |
| **Stats** | `employees`, `arrvFixed`, `deptFixed`, `spstNormalized`, `workUpdated` — the same numbers the backend returned in its `X-Process-Stats` header. |

Accepted inputs: `.xlsx`, `.xls` (legacy, via `xlrd`), `.csv`.
Output: `RC_HR_Processed_<original name>.xlsx`.

## Contractor tab

Same rules, different workbook shape. A contractor paysheet is a multi-sheet
book (invoice + wages register + attendance), so the engine has to find its way
around rather than assume sheet 1:

| Concern | HR tab | Contractor tab |
|---------|--------|----------------|
| **Attendance grid** | first sheet | any sheet with `SPST` + `ARRV`/`DEPT` headers; sheets named `*atten*`/`*muster*` are tried first |
| **Monthly OT total** | the `OT Hours` column, repeated on every row of the employee | the wages sheet's `OT Hrs` column (sheets named `*wage*`/`*pay*`/`*salary*` first) |
| **Employee key** | `Employee Code` | `Employee Code` **+** `Employee Name` — a contractor code is not unique (`RCC000` covers two people) |
| **Shift times** | real time cells (`9:30 AM`) | dotted text (`9.00 AM`), read with `parse_time_loose()` |
| **Times written as** | Excel day-fractions | 24-hour text (`09:05`, `18:31`), matching the sheet's own notation |
| **`OT Hours` column** | zeroed | receives the **per-day** OT, so the column sums to the wages-register total |

Every other sheet in the workbook — invoice, wages register — is left exactly
as it was, formulas included.

**The 9.5 h rest cap and OT.** A contractor shift can be longer than the cap
(9:00 AM–7:00 PM is 10 h). The ordinary part of the day is still capped at
9.5 h minus 0–30 min, and the OT granted is added on top of *that*, so
`WORK = capped base + OT`. On a 1 h OT day the result is ≈ `10:07`, not
`11:00`.

**When the paid OT does not fit.** OT normally goes only to a pure `DP` day
whose recorded span is not already longer than the shift. If those days cannot
absorb the wages-register total — which happens when the uploaded sheet
*already* carries OT in its departure times, i.e. you re-process an output
file — the engine falls back to every `DP` day rather than silently paying out
less, and the UI says which employees this affected. Anything still unplaceable
is reported as a shortfall instead of being dropped.

Accepted inputs: `.xlsx`, `.xlsm`. `.xls`/`.csv` are rejected with a clear
message, because flattening to one sheet would discard the wages register.
Output: `RC_Contractor_Processed_<original name>.xlsx`.

## Where the original logic lives

| Original (Node) | Here (Python) |
|-----------------|---------------|
| `routes/process.js` (the whole engine) | `lib/process.py` |
| `utils/normalizeStatus.js`, `utils/timeUtils.js` | `lib/utils.py` |
| `HRProcessing.jsx` (upload → process → download page) | `app.py` |
| — (new, no Node counterpart) | `lib/contractor.py` |

The DB-backed pages of the old dashboard (Login, Dashboard, Employees,
Reports, Sync To DB, Profile) were removed on purpose — this build is a pure
file tool with **no database**. The previous full app is archived in
`Streamlit-HR-Dashboard.zip`.

## Setup

Requires **Python 3.10+**.

```powershell
cd "Streamlit-HR-Dashboard"
python -m pip install -r requirements.txt
python -m streamlit run app.py
```

Or just double-click **`run.bat`**. The app opens at <http://localhost:8501>.
