"""
pages/5_Statistical_EDA.py - Forensic & Statistical Anomaly Profiling
Project: PJ-IVD-2026 | Surbhi Sharma (2500320140152) | ABES Engineering College
"""

import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from src.database import load_table
from src.analytics import compute_benford_distribution, detect_outliers_iqr, detect_outliers_zscore
from src.ui import render_header

st.set_page_config(page_title="Statistical EDA | Invoice Analytics", layout="wide")

render_header("Statistical EDA & Forensic Analysis", "Exploratory Profiling, Benford's Law & Outlier Clustering")

df_invoices = load_table("invoices")
df_lines = load_table("line_items")

tab1, tab2, tab3 = st.tabs(["Benford's Law Forensic Audit", "Line-Item Price Distributions & Outliers", "Invoice Amount Variance & Boxplots"])

with tab1:
    st.subheader("First-Digit Analysis against Theoretical Benford's Law")
    st.markdown("""
    **Benford's Law** asserts that in naturally occurring financial datasets, the number **1** will be the leading first digit
    approximately 30.1% of the time, while **9** will be the leading digit only 4.6% of the time. Significant deviations indicate
    artificial number invention, split purchase orders, or fraudulent invoicing.
    """)
    
    benford_df, chi2, p_val = compute_benford_distribution(df_invoices["total_amount"])
    
    if not benford_df.empty:
        fig_benford = go.Figure()
        fig_benford.add_trace(go.Bar(
            x=benford_df["digit"],
            y=benford_df["actual_pct"],
            name="Actual First Digit (%)",
            marker_color="#1E3A8A"
        ))
        fig_benford.add_trace(go.Scatter(
            x=benford_df["digit"],
            y=benford_df["theoretical_pct"],
            name="Benford Expected (%)",
            line=dict(color="#DC2626", width=3, dash="dash")
        ))
        fig_benford.update_layout(
            title="Leading Digit Distribution vs. Theoretical Benford's Curve",
            xaxis=dict(title="Leading First Digit (1 - 9)", tickmode="linear"),
            yaxis=dict(title="Frequency Percentage (%)"),
            barmode="group"
        )
        st.plotly_chart(fig_benford, use_container_width=True)
        
        c1, c2 = st.columns(2)
        c1.metric("Chi-Square Goodness of Fit Statistic", f"{chi2:.2f}")
        c2.metric("P-Value", f"{p_val:.4f}", "Conforms to Normal Distribution" if p_val > 0.05 else "Investigate Deviations")

with tab2:
    st.subheader("Line-Item Unit Price Outlier Detection")
    
    selected_cat = st.selectbox("Filter Category:", ["All Categories"] + sorted(df_lines["category"].unique().tolist()))
    subset_lines = df_lines if selected_cat == "All Categories" else df_lines[df_lines["category"] == selected_cat]
    
    is_outlier_iqr, lb, ub = detect_outliers_iqr(subset_lines["unit_price"])
    subset_lines = subset_lines.copy()
    subset_lines["is_outlier"] = is_outlier_iqr
    
    fig_box = px.box(
        subset_lines,
        x="category",
        y="unit_price",
        color="category",
        points="all",
        title="Unit Price Spread & Outlier Points by Category",
        labels={"unit_price": "Unit Price (INR)", "category": "Product Category"}
    )
    fig_box.update_layout(xaxis_tickangle=-45, showlegend=False)
    st.plotly_chart(fig_box, use_container_width=True)
    
    st.info(f"IQR Detection: Lower Bound = ₹{lb:,.2f} | Upper Bound = ₹{ub:,.2f} | Detected Outliers = {is_outlier_iqr.sum()} records")

with tab3:
    st.subheader("Invoice Total Amount Distribution & Skewness")
    fig_hist = px.histogram(
        df_invoices,
        x="total_amount",
        nbins=30,
        marginal="box",
        title="Distribution of Total Invoice Values (INR)",
        labels={"total_amount": "Invoice Total Amount (INR)"},
        color_discrete_sequence=["#2563EB"]
    )
    st.plotly_chart(fig_hist, use_container_width=True)
