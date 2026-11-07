from __future__ import annotations

from pathlib import Path
import tempfile

import pandas as pd
import streamlit as st

from barcode_report import build_pivot_table_sheet, load_excel_data


st.set_page_config(page_title="Barcode Report", page_icon="📊", layout="wide")
st.title("Barcode Report")
st.write("Upload a CSV or Excel report, build a pivot-style summary, and download the results as an Excel workbook.")

uploaded_file = st.file_uploader("Choose a source report", type=["csv", "xlsx", "xls", "xlsm"])

if uploaded_file is not None:
    source_suffix = Path(uploaded_file.name).suffix.lower() or ".xlsx"
    try:
        with tempfile.TemporaryDirectory() as temp_dir:
            source_path = Path(temp_dir) / f"source{source_suffix}"
            source_path.write_bytes(uploaded_file.getvalue())
            df = load_excel_data(source_path)

            if df.empty:
                st.warning("The uploaded file contains no data rows.")
                st.stop()

            st.subheader("Preview")
            st.dataframe(df.head(100), use_container_width=True)
            st.caption(f"{len(df):,} rows · {len(df.columns):,} columns")

            columns = list(df.columns)
            default_index = columns.index("Ticket Status") if "Ticket Status" in columns else 0
            default_column = columns.index("Time Frame") if "Time Frame" in columns else min(1, len(columns) - 1)
            default_value = columns.index("Ticket Number") if "Ticket Number" in columns else 0

            st.subheader("Summary settings")
            left, middle, right = st.columns(3)
            with left:
                index_column = st.selectbox("Row field", columns, index=default_index)
            with middle:
                pivot_column = st.selectbox("Column field", columns, index=default_column)
            with right:
                value_column = st.selectbox("Value field", columns, index=default_value)

            aggfunc = st.selectbox("Aggregation", ["count", "sum", "mean", "median", "min", "max"], index=0)

            actioned_by_filter: list[str] | None = None
            if "Actioned By" in df.columns:
                available_values = sorted(df["Actioned By"].dropna().astype(str).str.strip().unique().tolist())
                preferred_values = [value for value in ("nr3", "mhussain3") if value in available_values]
                selected_values = st.multiselect(
                    "Filter Actioned By",
                    available_values,
                    default=preferred_values,
                    help="Leave empty to include all Actioned By values.",
                )
                actioned_by_filter = selected_values or None

            sheet_name = st.text_input("Summary sheet name", value="Pivot", max_chars=31)
            if not sheet_name.strip():
                st.error("Enter a worksheet name.")
                st.stop()

            filtered_df = df
            if actioned_by_filter and "Actioned By" in df.columns:
                selected_lower = {value.strip().lower() for value in actioned_by_filter}
                filtered_df = df[
                    df["Actioned By"].fillna("").astype(str).str.strip().str.lower().isin(selected_lower)
                ].copy()

            if filtered_df.empty:
                st.warning("No rows match the selected Actioned By values.")
            elif st.button("Generate report", type="primary"):
                output_path = Path(temp_dir) / "barcode_report.xlsx"
                try:
                    build_pivot_table_sheet(
                        source_path=source_path,
                        output_path=output_path,
                        index=index_column,
                        columns=pivot_column,
                        values=value_column,
                        aggfunc=aggfunc,
                        sheet_name=sheet_name.strip(),
                        actioned_by_filter=actioned_by_filter,
                    )
                    output_bytes = output_path.read_bytes()
                    st.success("Report generated.")
                    st.download_button(
                        "Download Excel report",
                        data=output_bytes,
                        file_name="barcode_report.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    )
                except (ValueError, KeyError, TypeError) as exc:
                    st.error(f"Could not generate the report: {exc}")
    except Exception as exc:
        st.error(f"Could not read the uploaded file: {exc}")
else:
    st.info("Upload a CSV or Excel file to get started.")
