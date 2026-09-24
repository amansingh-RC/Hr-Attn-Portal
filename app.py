"""Royal Chain HR Processing Tool — Streamlit, no database.

Two tabs, one engine family:

* **HR Dashboard** — the original attendance tool. Upload an attendance
  Excel/CSV → SPST values are normalized, monthly OT is distributed across
  working days, arrival/departure times are rewritten and worked hours
  recomputed → download the processed .xlsx (lib/process.py).
* **Contractor** — the same OT, check-in, check-out and rest rules applied to
  a contractor paysheet workbook, where the attendance grid is one sheet among
  several and the monthly OT total comes from the Wages Register
  (lib/contractor.py).

Nothing is stored anywhere.

Run:  streamlit run app.py
"""

import streamlit as st

from lib.contractor import process_contractor_workbook
from lib.process import process_workbook
from lib.ui import inject_css, top_bar

st.set_page_config(
    page_title="Royal Chain HR Processing",
    page_icon="🏢",
    layout="centered",
    initial_sidebar_state="collapsed",
)

XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"


def _run_panel(*, key, title, caption, types, prefix, engine, metrics, extra=None):
    """One upload → process → download card.

    key      unique session-state prefix for this tab
    engine   callable(bytes, filename) -> (xlsx_bytes, stats)
    metrics  [(label, stats_key), ...] shown after a successful run
    extra    optional callable(stats) rendered under the metrics
    """
    with st.container(border=True):
        st.markdown(f"#### {title}")
        st.caption(caption)

        st.markdown("**Excel File**")
        upload = st.file_uploader(
            "Excel File",
            type=types,
            label_visibility="collapsed",
            key=f"{key}_upload",
        )
        if upload is not None:
            st.caption(f"{upload.name} — {upload.size / 1024:.1f} KB")

        # Choosing a different file clears the previous result
        sig = (upload.name, upload.size) if upload is not None else None
        if st.session_state.get(f"{key}_sig") != sig:
            st.session_state[f"{key}_sig"] = sig
            st.session_state.pop(f"{key}_result", None)
            st.session_state.pop(f"{key}_error", None)

        if st.button("📄 Process File", type="primary", key=f"{key}_btn"):
            st.session_state.pop(f"{key}_result", None)
            st.session_state.pop(f"{key}_error", None)
            if upload is None:
                st.session_state[f"{key}_error"] = "Please select an Excel file first."
            else:
                with st.spinner("Processing..."):
                    try:
                        data, stats = engine(upload.getvalue(), upload.name)
                        base = upload.name.rsplit(".", 1)[0]
                        st.session_state[f"{key}_result"] = {
                            "data": data,
                            "name": f"{prefix}{base}.xlsx",
                            "stats": stats,
                        }
                    except Exception as e:  # noqa: BLE001
                        st.session_state[f"{key}_error"] = str(e)

        error = st.session_state.get(f"{key}_error")
        if error:
            st.error(error)

        result = st.session_state.get(f"{key}_result")
        if result:
            st.download_button(
                "⬇ Download File",
                data=result["data"],
                file_name=result["name"],
                mime=XLSX_MIME,
                type="primary",
                key=f"{key}_dl",
            )
            st.success(
                "**File processed successfully.** Click **Download File** to save "
                f"`{result['name']}`."
            )

            stats = result["stats"]
            cols = st.columns(len(metrics))
            for c, (label, sk) in zip(cols, metrics):
                c.metric(label, stats[sk])
            if extra is not None:
                extra(stats)


def hr_tab():
    st.markdown('<p class="rc-h1">Manage HR Data</p>', unsafe_allow_html=True)
    st.markdown(
        '<p class="rc-sub">Upload the attendance Excel, process it, and download '
        "the updated file. No database — nothing leaves this page.</p>",
        unsafe_allow_html=True,
    )
    st.write("")

    _run_panel(
        key="proc",
        title="📊 Royal Chain HR Data",
        caption=(
            "Select an Excel file (.xlsx, .xls, .csv). The file will be processed "
            "and the updated copy offered for download."
        ),
        types=["xlsx", "xls", "csv"],
        prefix="RC_HR_Processed_",
        engine=process_workbook,
        metrics=[
            ("Employees", "employees"),
            ("Arrivals fixed", "arrvFixed"),
            ("Departures fixed", "deptFixed"),
            ("SPST normalized", "spstNormalized"),
            ("Work updated", "workUpdated"),
        ],
    )


def contractor_tab():
    st.markdown('<p class="rc-h1">Manage Contractor Data</p>', unsafe_allow_html=True)
    st.markdown(
        '<p class="rc-sub">Upload the file.</p>',
        unsafe_allow_html=True,
    )
    st.write("")

    def sources(stats):
        st.caption(
            f"Attendance sheet: **{stats['attendanceSheet']}** · "
            f"OT source: **{stats['wagesSheet']}** · "
            f"{stats['otHours']:g} OT hours distributed · "
            f"{stats['blankedRows']} off / absent rows left blank"
        )
        if not stats["otHours"] and stats["matched"]:
            st.info(
                "The wages sheet has no OT hours for anyone this month, so every "
                "OT Hours cell is 0. Fill in the **OT Hrs** column of the wages "
                "sheet and upload again if that is not right."
            )
        if stats["unmatched"]:
            st.warning(
                "No Wages Register row matched these employees, so they were "
                "given 0 OT hours: " + ", ".join(stats["unmatched"])
            )
        if stats["widened"]:
            st.info(
                "The strictly eligible days could not hold the month's paid OT "
                "for these employees, so every DP day was used instead — this "
                "happens when the uploaded sheet already carries OT in its "
                "departure times: " + ", ".join(stats["widened"])
            )
        if stats["shortfall"]:
            st.warning(
                f"{stats['shortfall']:g} OT hours could not be placed — even at "
                f"{2} h per day there are not enough working days. Check the "
                "Wages Register totals against the attendance sheet."
            )

    _run_panel(
        key="con",
        title="👷 Contractor Paysheet",
        caption=(
            "Select the contractor paysheet workbook (.xlsx, .xlsm). It must "
            "contain required Columns — ARRV, DEPT, OT Hours and WORK may come "
            "in empty, they are generated here. WO / PH / ABS / PL days stay blank."
        ),
        types=["xlsx", "xlsm"],
        prefix="RC_Contractor_Processed_",
        engine=process_contractor_workbook,
        metrics=[
            ("Employees", "employees"),
            ("Arrivals fixed", "arrvFixed"),
            ("Departures fixed", "deptFixed"),
            ("SPST normalized", "spstNormalized"),
            ("Work updated", "workUpdated"),
        ],
        extra=sources,
    )


def main():
    inject_css()
    top_bar()

    tab_hr, tab_con = st.tabs(["🏢  HR Dashboard", "👷  Contractor"])
    with tab_hr:
        hr_tab()
    with tab_con:
        contractor_tab()


if __name__ == "__main__":
    main()
