import io
import os
import sys
import tempfile
from pathlib import Path
import importlib.util

import streamlit as st


st.set_page_config(page_title="Trinity Customer Report Generator", layout="wide")
st.title("Trinity Customer Report Generator")
st.write("Upload CSV → Run your Python logic → Download Excel report")

# ---- Settings ----
DEFAULT_SCRIPT_NAME = "TRINITY CUSTOMER ANALYSIS CODE.py"  # your file in repo root


def load_logic_module(script_path: Path):
    """
    Load a Python file as a module dynamically.
    Works even if filename has spaces.
    """
    spec = importlib.util.spec_from_file_location("trinity_logic", str(script_path))
    if spec is None or spec.loader is None:
        raise ImportError(f"Cannot load python file: {script_path}")

    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


# ---- Uploads ----
uploaded_csv = st.file_uploader("Upload CSV file", type=["csv"])

# Optional: allow uploading the code file too (useful on Streamlit Cloud)
uploaded_code = st.file_uploader(
    "Upload Python logic file (optional, if not included in repo)",
    type=["py"],
    help="If you already placed the .py in the repo, you can skip this."
)

run_btn = st.button("Generate Report")

if run_btn:
    if not uploaded_csv:
        st.error("Please upload a CSV file first.")
        st.stop()

    with st.spinner("Running report generation..."):
        with tempfile.TemporaryDirectory() as tmpdir:
            tmpdir = Path(tmpdir)

            # Save CSV to temp
            csv_path = tmpdir / "input.csv"
            csv_path.write_bytes(uploaded_csv.getvalue())

            # Determine the code file path
            if uploaded_code:
                code_path = tmpdir / "uploaded_logic.py"
                code_path.write_bytes(uploaded_code.getvalue())
            else:
                # Use local script bundled in repo
                code_path = Path(__file__).resolve().parent / DEFAULT_SCRIPT_NAME

            if not code_path.exists():
                st.error(
                    f"Logic file not found: {code_path}\n\n"
                    f"Either upload the .py file above or place `{DEFAULT_SCRIPT_NAME}` next to app.py."
                )
                st.stop()

            # Load and execute
            try:
                logic_module = load_logic_module(code_path)
            except Exception as e:
                st.exception(e)
                st.stop()

            # Validate expected function
            if not hasattr(logic_module, "build_output"):
                st.error("Your logic file must contain a function named: build_output(input_csv_path, output_xlsx_path)")
                st.stop()

            # Output file path
            out_xlsx = tmpdir / "TRINITY_CUSTOMER_REPORT.xlsx"

            # Run the function
            try:
                logic_module.build_output(csv_path, out_xlsx)
            except Exception as e:
                st.error("Error while running build_output(). See details below.")
                st.exception(e)
                st.stop()

            if not out_xlsx.exists():
                st.error("Report generation finished but output file was not created.")
                st.stop()

            # Download
            st.success("Report generated successfully!")
            st.download_button(
                "Download Excel Report",
                data=out_xlsx.read_bytes(),
                file_name="TRINITY_CUSTOMER_REPORT.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

st.divider()
st.caption("Tip: On Streamlit Cloud, upload the CSV and (if needed) upload the .py logic file too.")
