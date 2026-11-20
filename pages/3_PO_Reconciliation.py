"""
pages/3_PO_Reconciliation.py - 2-Way & 3-Way PO Reconciliation
Project: PJ-IVD-2026 | Surbhi Sharma (2500320140152) | ABES Engineering College
"""

import streamlit as st
import pandas as pd
import plotly.express as px
from src.database import load_table
from src.validation import reconcile_po_variance
from src.ui import render_header

st.set_page_config(page_title="PO Reconciliation | Invoice Analytics", layout="wide")

render_header("Purchase Order (PO) 2-Way Reconciliation", "Matching Invoiced Values with Approved Procurement Contracts")

df_invoices = load_table("invoices")
df_pos = load_table("purchase_orders")
df_vendors = load_table("vendors")

recon_df = reconcile_po_variance(df_invoices, df_pos)
if "vendor_id" not in recon_df.columns:
    if "vendor_id_inv" in recon_df.columns:
        recon_df["vendor_id"] = recon_df["vendor_id_inv"]
    elif "vendor_id_po" in recon_df.columns:
        recon_df["vendor_id"] = recon_df["vendor_id_po"]

recon_df = recon_df.merge(df_vendors[["vendor_id", "vendor_name"]], on="vendor_id", how="left")

# Tolerance Filter
st.sidebar.header("Tolerance Parameters")
tolerance = st.sidebar.slider("Acceptable Variance Tolerance (%)", min_value=0.0, max_value=10.0, value=2.0, step=0.5)

recon_df["is_over_budget"] = recon_df["variance_pct"] > tolerance

# Metrics
c1, c2, c3 = st.columns(3)
over_budget_count = recon_df["is_over_budget"].sum()
total_recon = len(recon_df)
overrun_total = recon_df[recon_df["is_over_budget"]]["po_variance"].sum()

c1.metric("Reconciled Invoices", total_recon)
c2.metric("Over-Budget Invoices", over_budget_count, f"{round(over_budget_count/total_recon*100, 1)}% of total", delta_color="inverse")
c3.metric("Total Budget Overrun Exposure", f"₹{overrun_total:,.2f}", delta_color="inverse")

st.markdown("---")

col_chart1, col_chart2 = st.columns(2)

with col_chart1:
    st.subheader("PO Amount vs. Invoiced Amount")
    fig_scatter = px.scatter(
        recon_df,
        x="po_amount",
        y="total_amount",
        color="is_over_budget",
        hover_data=["invoice_id", "po_id", "vendor_name", "variance_pct"],
        labels={"po_amount": "Approved PO Budget (INR)", "total_amount": "Invoiced Amount (INR)"},
        title="PO vs. Invoiced Amount (Variance Outliers)",
        color_discrete_map={True: "#DC2626", False: "#16A34A"}
    )
    # Add 45-degree parity reference line
    fig_scatter.add_shape(
        type="line", line=dict(dash="dash", color="gray"),
        x0=0, x1=recon_df["po_amount"].max(),
        y0=0, y1=recon_df["po_amount"].max()
    )
    st.plotly_chart(fig_scatter, use_container_width=True)

with col_chart2:
    st.subheader("Distribution of Variance Percentage (%)")
    fig_hist = px.histogram(
        recon_df,
        x="variance_pct",
        nbins=25,
        title="PO Variance Distribution (%)",
        labels={"variance_pct": "Variance (%) from PO Budget"},
        color_discrete_sequence=["#1E3A8A"]
    )
    st.plotly_chart(fig_hist, use_container_width=True)

st.markdown("---")
st.subheader("Over-Budget Reconciliations Table")
over_df = recon_df[recon_df["is_over_budget"]].sort_values(by="po_variance", ascending=False)
st.dataframe(
    over_df[["invoice_id", "po_id", "vendor_name", "department", "po_amount", "total_amount", "po_variance", "variance_pct"]],
    use_container_width=True
)
