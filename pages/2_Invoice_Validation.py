"""
pages/2_Invoice_Validation.py - Invoice Inspection & Rule Validation
Project: PJ-IVD-2026 | Surbhi Sharma (2500320140152) | ABES Engineering College
"""

import streamlit as st
import pandas as pd
from src.database import load_table, get_connection
from src.validation import validate_arithmetic, detect_duplicates
from src.ui import render_header

st.set_page_config(page_title="Invoice Validation | Invoice Analytics", layout="wide")

render_header("Automated Invoice Validation Engine", "Zero-Tolerance Multi-Dimensional Discrepancy Detection")

df_invoices = load_table("invoices")
df_vendors = load_table("vendors")
df_lines = load_table("line_items")
df_disc = load_table("discrepancies")

# Filter Sidebar
st.sidebar.header("Filter Invoices")
status_filter = st.sidebar.multiselect(
    "Validation Status",
    options=df_invoices["validation_status"].unique(),
    default=df_invoices["validation_status"].unique()
)

vendor_list = ["All Vendors"] + sorted(df_vendors["vendor_name"].tolist())
selected_vendor = st.sidebar.selectbox("Vendor Filter", vendor_list)

filtered_inv = df_invoices[df_invoices["validation_status"].isin(status_filter)]
if selected_vendor != "All Vendors":
    v_id = df_vendors[df_vendors["vendor_name"] == selected_vendor]["vendor_id"].iloc[0]
    filtered_inv = filtered_inv[filtered_inv["vendor_id"] == v_id]

# Display Statistics
s1, s2, s3 = st.columns(3)
s1.metric("Invoices Matching Filters", len(filtered_inv))
s2.metric("Flagged Discrepancies", (filtered_inv["validation_status"] == "DISCREPANCY_DETECTED").sum())
s3.metric("Clean / Validated Invoices", (filtered_inv["validation_status"] == "VALIDATED").sum())

st.markdown("### Invoices Registry")
st.dataframe(
    filtered_inv[["invoice_id", "invoice_number", "po_id", "vendor_id", "invoice_date", "total_amount", "validation_status", "payment_status"]],
    use_container_width=True,
    height=250
)

st.markdown("---")
st.subheader("Invoice Deep-Dive & Line-Item Auditor")

selected_inv_id = st.selectbox("Select Invoice to Inspect:", options=filtered_inv["invoice_id"].tolist())

if selected_inv_id:
    inv_meta = filtered_inv[filtered_inv["invoice_id"] == selected_inv_id].iloc[0]
    inv_lines = df_lines[df_lines["invoice_id"] == selected_inv_id]
    inv_discs = df_disc[df_disc["invoice_id"] == selected_inv_id]
    
    col_a, col_b = st.columns([2, 1])
    with col_a:
        st.markdown(f"**Invoice Number:** `{inv_meta['invoice_number']}` | **PO Ref:** `{inv_meta['po_id']}` | **Date:** `{inv_meta['invoice_date']}`")
        st.markdown(f"**Total Amount Billed:** ₹{inv_meta['total_amount']:,.2f} (Subtotal: ₹{inv_meta['subtotal']:,.2f}, Tax: ₹{inv_meta['tax_amount']:,.2f})")
    with col_b:
        if not inv_discs.empty:
            st.error(f"⚠️ {len(inv_discs)} Discrepancy Flag(s) Detected!")
        else:
            st.success("✅ Zero Discrepancies Detected. Clean Bill.")

    st.markdown("#### Line Items")
    st.dataframe(
        inv_lines[["item_id", "item_description", "category", "quantity", "unit_price", "tax_rate", "tax_amount", "line_total", "has_discrepancy"]],
        use_container_width=True
    )

    if not inv_discs.empty:
        st.markdown("#### Discrepancy Log for this Invoice")
        st.dataframe(
            inv_discs[["disc_id", "disc_type", "severity", "variance_amount", "root_cause", "status"]],
            use_container_width=True
        )

# Download exception manifest
st.markdown("---")
st.subheader("Download Audit Exception Manifest")
csv_exceptions = df_disc.to_csv(index=False).encode('utf-8')
st.download_button(
    label="📥 Download Full Discrepancy Exception Manifest (CSV)",
    data=csv_exceptions,
    file_name="Invoice_Audit_Discrepancies_Exception_Report.csv",
    mime="text/csv"
)
