"""
Peer Eval Processor — Streamlit UI. Logic lives in `peer_eval_core.py`.
"""
from __future__ import annotations

import tempfile
from pathlib import Path

import streamlit as st

from peer_eval_core import process_to_bytes


def main():
    st.set_page_config(page_title="U-Check", page_icon="📊", layout="centered")
    st.title("U-Check")
    st.markdown(
        "Upload your **Qualtrics peer evaluation** Excel export. The app builds one sheet per "
        "reviewed student plus a **Summary** sheet."
    )

    uploaded = st.file_uploader("Qualtrics Excel file (.xlsx)", type=["xlsx"])

    total_points = st.number_input(
        "Total points for the peer evaluation (used in grade formulas)",
        min_value=0.0,
        max_value=1000.0,
        value=25.0,
        step=0.5,
    )

    if uploaded is not None and st.button("Generate Excel report", type="primary"):
        suffix = Path(uploaded.name).suffix or ".xlsx"
        with tempfile.NamedTemporaryFile(delete=False, suffix=suffix) as tmp:
            tmp.write(uploaded.getvalue())
            tmp_path = tmp.name

        try:
            with st.spinner("Processing…"):
                out_bytes, out_name = process_to_bytes(tmp_path, total_points=float(total_points))
            st.success("Done.")
            st.download_button(
                label="Download processed workbook",
                data=out_bytes,
                file_name=out_name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
        except Exception as e:
            st.error(str(e))
        finally:
            Path(tmp_path).unlink(missing_ok=True)

    with st.expander("What file do I need?"):
        st.markdown(
            "- Export from Qualtrics as **Excel**.\n"
            "- The survey should match the structure this tool expects (8 Likert items + comment per teammate block).\n"
            "- Reviewer names are read from column **Q4**."
        )


if __name__ == "__main__":
    main()
