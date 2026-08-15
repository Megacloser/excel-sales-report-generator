"""Streamlit web app for the sales report generator."""

from io import BytesIO

import pandas as pd
import streamlit as st
from openpyxl import Workbook

from generate_sales_report import (
    DATE_FORMAT_EXAMPLES,
    build_summary,
    load_data,
    write_data_sheet,
    write_summary_sheet,
)

st.set_page_config(page_title="Sales Report Generator", page_icon="📊", layout="centered")

st.title("📊 Sales Report Generator")
st.write(
    "Upload a sales CSV with columns **Customer, Product, Amount, Date** and get a "
    "polished, presentation-ready Excel report — grouped and totaled by customer."
)
st.caption(
    "Column headers are matched case-insensitively and accept common synonyms — "
    "e.g. Customer/customer/Client/Клиент, Product/Item/Товар, Amount/Sum/Сумма, Date/Дата."
)
st.caption(
    f"Date accepts several formats: {DATE_FORMAT_EXAMPLES}. "
    "Example row: `John Smith, Laptop, 899.99, 2026-01-05`"
)

if "report_bytes" not in st.session_state:
    st.session_state.report_bytes = None
    st.session_state.summary = None
    st.session_state.df = None

uploaded_file = st.file_uploader("Upload CSV file", type=["csv"])

if uploaded_file is not None:
    if st.button("Process", type="primary"):
        try:
            df = load_data(uploaded_file)

            if df.empty:
                st.error("The uploaded CSV is empty. Please upload a file with at least one row of data.")
            elif df["Amount"].isna().all():
                st.error("The 'Amount' column contains no valid numbers. Check the CSV formatting.")
            elif df["Date"].isna().all():
                st.error("The 'Date' column contains no valid dates. Expected format: MM/DD/YYYY.")
            else:
                summary = build_summary(df)

                wb = Workbook()
                write_summary_sheet(wb, summary)
                write_data_sheet(wb, df)

                buffer = BytesIO()
                wb.save(buffer)

                st.session_state.report_bytes = buffer.getvalue()
                st.session_state.summary = summary
                st.session_state.df = df
                st.success(f"Report generated: {len(summary)} customers, {len(df)} orders.")

        except ValueError as e:
            st.session_state.report_bytes = None
            st.error(f"Invalid file format: {e}")
        except pd.errors.EmptyDataError:
            st.session_state.report_bytes = None
            st.error("The uploaded file is empty or not a valid CSV.")
        except pd.errors.ParserError:
            st.session_state.report_bytes = None
            st.error("Could not parse the file. Make sure it's a valid, comma-separated CSV.")
        except UnicodeDecodeError:
            st.session_state.report_bytes = None
            st.error("Could not read the file's text encoding. Please save the CSV as UTF-8.")
        except Exception as e:
            st.session_state.report_bytes = None
            st.error(f"Unexpected error while processing the file: {e}")
else:
    st.session_state.report_bytes = None

if st.session_state.report_bytes is not None:
    summary = st.session_state.summary
    df = st.session_state.df

    col1, col2, col3 = st.columns(3)
    col1.metric("Customers", len(summary))
    col2.metric("Orders", len(df))
    col3.metric("Total Revenue", f"${summary['Total Spent'].sum():,.2f}")

    st.subheader("Customer Summary")
    st.dataframe(
        summary.style.format({"Total Spent": "${:,.2f}", "% of Total": "{:.1%}"}),
        use_container_width=True,
        hide_index=True,
    )

    st.subheader("Raw Data Preview")
    preview = df.head(20).copy()
    preview["Date"] = preview["Date"].dt.strftime("%m/%d/%Y")
    st.dataframe(
        preview.style.format({"Amount": "${:,.2f}"}),
        use_container_width=True,
        hide_index=True,
    )

    st.download_button(
        label="Download Excel Report",
        data=st.session_state.report_bytes,
        file_name="sales_report.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
else:
    st.info("Upload a CSV file and click **Process** to generate the report.")
