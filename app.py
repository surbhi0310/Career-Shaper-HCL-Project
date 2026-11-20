"""
app.py - Invoice Data Validation & Discrepancy Analytics Platform
Root entry point for Streamlit Community Cloud deployment.
Project: PJ-IVD-2026 | Surbhi Sharma (2500320140152) | ABES Engineering College
Mentor: Kalpana Dwivedi
"""

import streamlit as st
import pandas as pd
import plotly.express as px
from src.database import get_summary_metrics, load_table
from src.ui import render_header, render_kpi, format_inr

st.set_page_config(
    page_title="Invoice Data Validation Platform | ABES",
    page_icon="🔑",
    layout="wide",
    initial_sidebar_state="expanded"
)

def main():
    render_header(
        title="Invoice Data Validation & Discrepancy Analytics Platform",
        subtitle="End-to-End Enterprise Accounts Payable (AP) Anomaly Detection & Reconciliation System"
    )

    # Sidebar project info
    st.sidebar.markdown(
        """
        <div style="display: flex; align-items: center; gap: 10px; margin-bottom: 12px; padding: 8px 12px; background: #F1F5F9; border-radius: 8px;">
            <span style="font-size: 26px;">🔑</span>
            <span style="font-size: 15px; font-weight: 700; color: #1E3A8A; letter-spacing: 0.02em;">AP AUDIT PORTAL</span>
        </div>
        """,
        unsafe_allow_html=True
    )
    st.sidebar.title("Navigation")
    st.sidebar.markdown("""
    **Project ID:** `PJ-IVD-2026`  
    **Author:** Surbhi Sharma  
    **Roll No:** `2500320140152`  
    **Course:** MCA 2nd Year (DA)  
    **College:** ABES Engineering College  
    **Mentor:** Kalpana Dwivedi  
    """)
    st.sidebar.divider()
    st.sidebar.info("Select a module from the sidebar above or use the quick links below.")

    # KPI Summary Row
    metrics = get_summary_metrics()
    c1, c2, c3, c4 = st.columns(4)
    with c1:
        render_kpi("Total Invoices Processed", f"{metrics['total_invoices']:,}", "100% Population Coverage", "val-primary")
    with c2:
        render_kpi("Total Procurement Spend", format_inr(metrics['total_spend']), "Across 30 Master Vendors", "val-primary")
    with c3:
        render_kpi("Flagged Discrepancies", f"{metrics['total_discrepancies']:,}", f"Error Rate: {metrics['error_rate']}%", "val-danger")
    with c4:
        render_kpi("Prevented Capital Leakage", format_inr(metrics['prevented_leakage']), f"{metrics['duplicate_invoices']} Duplicate Collisions", "val-success")

    st.markdown("---")

    # High-level overview columns
    col_left, col_right = st.columns([3, 2])
    
    with col_left:
        st.subheader("Executive Architecture & System Scope")
        st.markdown("""
        The **Invoice Data Validation & Discrepancy Analytics Platform** provides automated, zero-tolerance reconciliation
        across corporate billing pipelines. It eliminates manual sampling vulnerabilities by executing algorithmic checks on:
        - **Line-Item Arithmetic:** Instant calculation verification (`Quantity × Unit Price + Tax = Line Total`).
        - **Duplicate Bill Detection:** Rolling 14-day fuzzy matching on vendor, amounts, and dates.
        - **Tax Rate Integrity:** Cross-referencing billed GST slabs against statutory tax categories.
        - **PO-to-Invoice 2-Way & 3-Way Matching:** Real-time variance tracking against approved PO budgets.
        - **Vendor Risk Index (VRI):** Quantitative multi-factor scoring of supplier compliance.
        """)

    with col_right:
        st.subheader("Discrepancy Breakdown")
        df_disc = load_table("discrepancies")
        if not df_disc.empty:
            fig = px.pie(
                df_disc,
                names="disc_type",
                values="variance_amount",
                title="Financial Exposure by Discrepancy Type",
                hole=0.45,
                color_discrete_sequence=px.colors.qualitative.Bold
            )
            fig.update_layout(margin=dict(t=30, b=10, l=10, r=10), height=280)
            st.plotly_chart(fig, use_container_width=True)

    st.markdown("---")
    st.subheader("System Modules")
    m1, m2, m3 = st.columns(3)
    with m1:
        st.markdown("#### 📊 1. Overview & KPIs")
        st.caption("Comprehensive dashboard with time-series trend analysis, category distributions, and spend metrics.")
    with m2:
        st.markdown("#### 🔍 2. Invoice Validation")
        st.caption("Interactive invoice inspector, line-item error highlighter, and audit exception export.")
    with m3:
        st.markdown("#### ⚖️ 3. PO Reconciliation")
        st.caption("2-Way matching between PO budgets and billed amounts with price/quantity drift profiling.")

    m4, m5, m6 = st.columns(3)
    with m4:
        st.markdown("#### 🏢 4. Vendor Risk Profiling")
        st.caption("Composite Vendor Risk Index (VRI) rankings, rogue vendor identification, and compliance tiers.")
    with m5:
        st.markdown("#### 📈 5. Statistical EDA")
        st.caption("Benford's Law forensic analysis, IQR & Z-score outlier detection, and variance distributions.")
    with m6:
        st.markdown("#### 💾 6. Data Quality & Export")
        st.caption("Raw database table browser, missing value profiling, and Excel/CSV download hub.")

if __name__ == '__main__':
    main()
