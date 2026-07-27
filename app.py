"""Royal Chain HR Processing Tool — Streamlit, no database.

A single-page tool with the exact data-manipulation engine of the original
React + Express app (routes/process.js): upload an attendance Excel/CSV →
SPST values are normalized, monthly OT is distributed across working days,
arrival/departure times are rewritten and worked hours recomputed → download
the processed .xlsx. Nothing is stored anywhere.

Run:  streamlit run app.py
"""

import streamlit as st

from lib.process import process_workbook
from lib.ui import inject_css, top_bar

st.set_page_config(
    page_title="Royal Chain HR Processing",
    page_icon="🏢",
    layout="centered",
    initial_sidebar_state="collapsed",
)

XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"


def main():
    inject_css()
    top_bar()

    st.markdown('<p class="rc-h1">Manage HR Data</p>', unsafe_allow_html=True)
    st.markdown(
        '<p class="rc-sub">Upload the attendance Excel, process it, and download '
        "the updated file. No database — nothing leaves this page.</p>",
        unsafe_allow_html=True,
    )
    st.write("")

    with st.container(border=True):
        st.markdown("#### 📊 Royal Chains HR Data")
        st.caption(
            "Select an Excel file (.xlsx, .xls, .csv). The file will be processed "
            "and the updated copy offered for download."
        )

        st.markdown("**Excel File**")
        upload = st.file_uploader(
            "Excel File",
            type=["xlsx", "xls", "csv"],
            label_visibility="collapsed",
            key="proc_upload",
        )
        if upload is not None:
            st.caption(f"{upload.name} — {upload.size / 1024:.1f} KB")

        # Choosing a different file clears the previous result (mirrors
        # handleFileChange in HRProcessing.jsx)
        sig = (upload.name, upload.size) if upload is not None else None
        if st.session_state.get("proc_sig") != sig:
            st.session_state["proc_sig"] = sig
            st.session_state.pop("proc_result", None)
            st.session_state.pop("proc_error", None)

        if st.button("📄 Process File", type="primary"):
            st.session_state.pop("proc_result", None)
            st.session_state.pop("proc_error", None)
            if upload is None:
                st.session_state["proc_error"] = "Please select an Excel file first."
            else:
                with st.spinner("Processing..."):
                    try:
                        data, stats = process_workbook(upload.getvalue(), upload.name)
                        base = upload.name.rsplit(".", 1)[0]
                        st.session_state["proc_result"] = {
                            "data": data,
                            "name": f"RC_HR_Processed_{base}.xlsx",
                            "stats": stats,
                        }
                    except Exception as e:  # noqa: BLE001
                        st.session_state["proc_error"] = str(e)

        error = st.session_state.get("proc_error")
        if error:
            st.error(error)

        result = st.session_state.get("proc_result")
        if result:
            st.download_button(
                "⬇ Download File",
                data=result["data"],
                file_name=result["name"],
                mime=XLSX_MIME,
                type="primary",
            )
            st.success(
                "**File processed successfully.** Click **Download File** to save "
                f"`{result['name']}`."
            )

            stats = result["stats"]
            cols = st.columns(5)
            for col, (label, val) in zip(
                cols,
                [
                    ("Employees", stats["employees"]),
                    ("Arrivals fixed", stats["arrvFixed"]),
                    ("Departures fixed", stats["deptFixed"]),
                    ("SPST normalized", stats["spstNormalized"]),
                    ("Work updated", stats["workUpdated"]),
                ],
            ):
                col.metric(label, val)


if __name__ == "__main__":
    main()
