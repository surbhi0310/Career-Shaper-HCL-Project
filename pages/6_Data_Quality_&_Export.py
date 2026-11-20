"""
pages/6_Data_Quality_&_Export.py - Data Profiling, Quality & Export Center
Project: PJ-IVD-2026 | Surbhi Sharma (2500320140152) | ABES Engineering College
"""

import os
import streamlit as st
import pandas as pd
from src.database import load_table, get_connection
from src.ui import render_header

st.set_page_config(page_title="Data Quality & Export | Invoice Analytics", layout="wide")

render_header("Data Quality Profiling & Export Center", "Inspect Raw SQLite Tables, Profile Missing Values & Export Cleaned Data")

tabs = st.tabs(["Raw Tables Inspector", "Data Quality & Missing Values", "Data Export Hub"])

with tabs[0]:
    st.subheader("Database Tables Browser")
    selected_table = st.selectbox("Select SQLite Table to Browse:", ["vendors", "purchase_orders", "invoices", "line_items", "discrepancies"])
    df = load_table(selected_table)
    st.write(f"Showing **{len(df):,}** records from `{selected_table}` table:")
    st.dataframe(df, use_container_width=True, height=350)

with tabs[1]:
    st.subheader("Missing Values & Nulls Profiling")
    tables = ["vendors", "purchase_orders", "invoices", "line_items", "discrepancies"]
    dq_records = []
    for t in tables:
        t_df = load_table(t)
        null_counts = t_df.isnull().sum()
        for col, nulls in null_counts.items():
            dq_records.append({
                "table_name": t,
                "column_name": col,
                "null_count": nulls,
                "null_percentage": round((nulls / len(t_df)) * 100, 2),
                "data_type": str(t_df[col].dtype),
                "quality_status": "Clean" if nulls == 0 else "Intentional Injected Missing Value"
            })
    dq_df = pd.DataFrame(dq_records)
    st.dataframe(dq_df, use_container_width=True)

with tabs[2]:
    st.subheader("Export Cleaned Data & Analytical Artefacts")
    
    excel_path = os.path.join(os.path.dirname(os.path.dirname(__file__)), "data", "invoice_data.xlsx")
    if os.path.exists(excel_path):
        with open(excel_path, "rb") as f:
            bytes_data = f.read()
        st.download_button(
            label="📊 Download Full Multi-Tab Project Excel Workbook (data/invoice_data.xlsx)",
            data=bytes_data,
            file_name="Invoice_Data_Validation_Analysis_All_Tables.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    
    st.markdown("---")
    st.caption("All datasets conform to synthetic generation standards under academic project PJ-IVD-2026.")
