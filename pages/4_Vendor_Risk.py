"""
pages/4_Vendor_Risk.py - Vendor Risk Index (VRI) & Compliance League
Project: P053 | Surbhi Sharma (2500320140152) | ABES Engineering College
"""

import streamlit as st
import pandas as pd
import plotly.express as px
from src.database import load_table
from src.validation import calculate_vendor_risk_index
from src.ui import render_header

st.set_page_config(page_title="Vendor Risk Index | Invoice Analytics", layout="wide")

render_header("Vendor Risk Index (VRI) & Supplier Compliance", "Algorithmic Risk Profiling to Prioritize Audits & Renegotiate Contracts")

df_vendors = load_table("vendors")
df_invoices = load_table("invoices")
df_discrepancies = load_table("discrepancies")

vri_df = calculate_vendor_risk_index(df_vendors, df_discrepancies, df_invoices)

# KPI Summary
high_risk = (vri_df["risk_tier"] == "High Risk (Critical Audit)").sum()
med_risk = (vri_df["risk_tier"] == "Medium Risk (Sample Audit)").sum()
low_risk = (vri_df["risk_tier"] == "Low Risk (Auto-Approve)").sum()

k1, k2, k3 = st.columns(3)
k1.metric("🔴 High Risk Vendors", high_risk, "Requires 100% Pre-Payment Audit", delta_color="inverse")
k2.metric("🟡 Medium Risk Vendors", med_risk, "Requires 25% Sample Review", delta_color="off")
k3.metric("🟢 Low Risk Vendors", low_risk, "Eligible for Straight-Through Processing", delta_color="normal")

st.markdown("---")

c1, c2 = st.columns([3, 2])

with c1:
    st.subheader("Vendor Risk Index (VRI) Rankings")
    fig_bar = px.bar(
        vri_df.head(15),
        x="vri_score",
        y="vendor_name",
        orientation="h",
        color="risk_tier",
        title="Top 15 Vendors by Risk Score (0 - 100)",
        labels={"vri_score": "Vendor Risk Index (VRI)", "vendor_name": "Vendor Name"},
        color_discrete_map={
            "High Risk (Critical Audit)": "#DC2626",
            "Medium Risk (Sample Audit)": "#D97706",
            "Low Risk (Auto-Approve)": "#16A34A"
        }
    )
    fig_bar.update_layout(yaxis=dict(autorange="reversed"))
    st.plotly_chart(fig_bar, use_container_width=True)

with c2:
    st.subheader("Billed Spend vs. Discrepancy Leakage")
    fig_scatter = px.scatter(
        vri_df,
        x="total_billed",
        y="total_variance",
        size="discrepancy_count",
        color="risk_tier",
        hover_data=["vendor_name", "error_rate_pct", "vri_score"],
        labels={"total_billed": "Total Spend (INR)", "total_variance": "Total Discrepancy Amount (INR)"},
        title="Spend Exposure vs. Variance Leakage",
        color_discrete_map={
            "High Risk (Critical Audit)": "#DC2626",
            "Medium Risk (Sample Audit)": "#D97706",
            "Low Risk (Auto-Approve)": "#16A34A"
        }
    )
    st.plotly_chart(fig_scatter, use_container_width=True)

st.markdown("---")
st.subheader("Complete Vendor Risk Scorecard")
st.dataframe(
    vri_df[["vendor_id", "vendor_name", "category", "city", "total_invoices", "total_billed", "discrepancy_count", "error_rate_pct", "total_variance", "vri_score", "risk_tier"]],
    use_container_width=True
)
