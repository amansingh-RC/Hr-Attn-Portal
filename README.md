# Royal Chain HR Processing Tool — Streamlit

A single-page, **database-free** tool: upload the attendance Excel → it is
processed with the **exact same logic** as the original React + Express app →
download the updated file. That's it.

The whole data-manipulation engine of the original backend
(`routes/process.js`) is ported 1:1 to Python in [`lib/process.py`](lib/process.py).

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

## Where the original logic lives

| Original (Node) | Here (Python) |
|-----------------|---------------|
| `routes/process.js` (the whole engine) | `lib/process.py` |
| `utils/normalizeStatus.js`, `utils/timeUtils.js` | `lib/utils.py` |
| `HRProcessing.jsx` (upload → process → download page) | `app.py` |

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
