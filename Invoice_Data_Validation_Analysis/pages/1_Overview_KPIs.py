"""
pages/1_Overview_KPIs.py - Executive KPIs & Trend Analytics
Project: PJ-IVD-2026 | Surbhi Sharma (2500320140152) | ABES Engineering College
"""

import streamlit as st
import pandas as pd
import plotly.express as px
from src.database import load_table, get_summary_metrics
from src.analytics import get_monthly_spend_trend
from src.ui import render_header, render_kpi

st.set_page_config(page_title="Executive KPIs | Invoice Analytics", layout="wide")

render_header("Executive Overview & Macro KPIs", "Macro-level Financial Exposure & Trend Tracking")

metrics = get_summary_metrics()
c1, c2, c3, c4 = st.columns(4)
with c1:
    render_kpi("Total Invoices", f"{metrics['total_invoices']:,}", "Analyzed in SQLite", "val-primary")
with c2:
    render_kpi("Gross Spend", f"₹{metrics['total_spend']:,.2f}", "Total Billed Value", "val-primary")
with c3:
    render_kpi("Flagged Leakage", f"₹{metrics['prevented_leakage']:,.2f}", f"{metrics['total_discrepancies']} Invoices Flagged", "val-danger")
with c4:
    render_kpi("Discrepancy Rate", f"{metrics['error_rate']}%", "Industry Avg: 2.0% - 3.5%", "val-warning")

st.markdown("---")

df_inv = load_table("invoices")
df_vendors = load_table("vendors")
df_merged = df_inv.merge(df_vendors, on="vendor_id", how="left")

col1, col2 = st.columns(2)

with col1:
    st.subheader("Monthly Spend & Discrepancy Trends")
    trend_df = get_monthly_spend_trend(df_inv)
    fig_trend = px.bar(
        trend_df,
        x="month",
        y="total_spend",
        title="Gross Spend (INR) by Month",
        labels={"total_spend": "Total Spend (INR)", "month": "Billing Month"},
        color_discrete_sequence=["#1E3A8A"]
    )
    st.plotly_chart(fig_trend, use_container_width=True)

with col2:
    st.subheader("Discrepancy Rate (%) by Month")
    fig_rate = px.line(
        trend_df,
        x="month",
        y="discrepancy_rate_pct",
        markers=True,
        title="Discrepancy Rate Trend (%)",
        labels={"discrepancy_rate_pct": "Discrepancy Rate (%)", "month": "Billing Month"},
        color_discrete_sequence=["#DC2626"]
    )
    st.plotly_chart(fig_rate, use_container_width=True)

st.markdown("---")
st.subheader("Procurement Spend & Leakage by Vendor Category")

cat_summary = df_merged.groupby("category").agg(
    invoices=("invoice_id", "count"),
    spend=("total_amount", "sum")
).reset_index().sort_values(by="spend", ascending=False)

fig_cat = px.bar(
    cat_summary,
    x="category",
    y="spend",
    color="category",
    title="Spend Distribution across Categories (INR)",
    labels={"spend": "Total Spend (INR)", "category": "Product / Service Category"}
)
fig_cat.update_layout(xaxis_tickangle=-45, showlegend=False)
st.plotly_chart(fig_cat, use_container_width=True)
