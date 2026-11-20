"""
validation.py - Discrepancy & Validation Rules Engine
Project: PJ-IVD-2026 | Surbhi Sharma (2500320140152) | ABES Engineering College
"""

import pandas as pd
import numpy as np

def validate_arithmetic(line_items_df, tolerance=0.01):
    """
    Verifies that Quantity * Unit Price + Tax Amount == Line Total
    Returns DataFrame with calculation discrepancies.
    """
    df = line_items_df.copy()
    expected_tax = np.round(df["quantity"] * df["unit_price"] * df["tax_rate"], 2)
    expected_total = np.round((df["quantity"] * df["unit_price"]) + expected_tax, 2)
    
    diff = np.abs(df["line_total"] - expected_total)
    df["expected_total"] = expected_total
    df["arithmetic_error"] = diff > tolerance
    df["error_variance"] = np.round(df["line_total"] - expected_total, 2)
    return df

def detect_duplicates(invoices_df, days_window=14):
    """
    Detects potential duplicate invoices: Same vendor, identical total amount,
    submitted within a rolling window of N days.
    """
    df = invoices_df.copy()
    df["invoice_date_dt"] = pd.to_datetime(df["invoice_date"])
    df = df.sort_values(by=["vendor_id", "total_amount", "invoice_date_dt"])
    
    duplicates = []
    # Group by vendor and total amount
    grouped = df.groupby(["vendor_id", "total_amount"])
    for (v_id, amt), group in grouped:
        if len(group) > 1:
            dates = group["invoice_date_dt"].tolist()
            inv_ids = group["invoice_id"].tolist()
            for i in range(len(dates) - 1):
                delta = (dates[i+1] - dates[i]).days
                if delta <= days_window:
                    duplicates.append({
                        "vendor_id": v_id,
                        "primary_invoice": inv_ids[i],
                        "duplicate_invoice": inv_ids[i+1],
                        "amount": amt,
                        "date_diff_days": delta
                    })
    return pd.DataFrame(duplicates)

def validate_tax_rates(line_items_df, valid_slabs=(0.0, 0.05, 0.12, 0.18, 0.28)):
    """
    Flags line items where tax rate violates statutory tax slabs.
    """
    df = line_items_df.copy()
    df["invalid_tax_rate"] = ~df["tax_rate"].round(2).isin(valid_slabs)
    return df

def reconcile_po_variance(invoices_df, pos_df, tolerance_pct=0.02):
    """
    Performs 2-Way matching between Invoices and POs.
    Flags invoices exceeding approved PO budget beyond tolerance.
    """
    merged = invoices_df.merge(pos_df, on="po_id", suffixes=("_inv", "_po"), how="inner")
    if "vendor_id_inv" in merged.columns:
        merged["vendor_id"] = merged["vendor_id_inv"]
    elif "vendor_id_po" in merged.columns:
        merged["vendor_id"] = merged["vendor_id_po"]
    merged["po_variance"] = merged["total_amount"] - merged["po_amount"]
    merged["variance_pct"] = np.round((merged["po_variance"] / merged["po_amount"]) * 100, 2)
    merged["is_over_budget"] = merged["variance_pct"] > (tolerance_pct * 100)
    return merged

def calculate_vendor_risk_index(vendors_df, discrepancies_df, invoices_df):
    """
    Computes a composite Vendor Risk Index (VRI) from 0 to 100 based on:
    - Error rate (Discrepancy count / Invoices count)
    - Severity weighting (Critical=3x, High=2x, Medium=1x)
    - Total variance leakage amount
    """
    v_stats = invoices_df.groupby("vendor_id").agg(
        total_invoices=("invoice_id", "count"),
        total_billed=("total_amount", "sum")
    ).reset_index()
    
    # Severity points
    disc_weights = {"Critical": 3.0, "High": 2.0, "Medium": 1.0, "Low": 0.5}
    d_df = discrepancies_df.copy()
    d_df["weight"] = d_df["severity"].map(disc_weights).fillna(1.0)
    
    d_stats = d_df.groupby("vendor_id").agg(
        discrepancy_count=("disc_id", "count"),
        weighted_points=("weight", "sum"),
        total_variance=("variance_amount", "sum")
    ).reset_index()
    
    merged = vendors_df.merge(v_stats, on="vendor_id", how="left").fillna(0)
    merged = merged.merge(d_stats, on="vendor_id", how="left").fillna(0)
    
    # Error frequency
    merged["error_rate_pct"] = np.where(
        merged["total_invoices"] > 0,
        np.round((merged["discrepancy_count"] / merged["total_invoices"]) * 100, 2),
        0.0
    )
    
    # Risk Score Formula (0 to 100)
    # 40% error rate + 40% weighted points normalized + 20% baseline reliability penalty
    max_pts = max(merged["weighted_points"].max(), 1.0)
    score = (
        (merged["error_rate_pct"] * 0.45) +
        ((merged["weighted_points"] / max_pts) * 45) +
        ((1.0 - merged["baseline_reliability"]) * 10)
    )
    merged["vri_score"] = np.round(np.clip(score, 0, 100), 1)
    
    def assign_tier(vri):
        if vri >= 65:
            return "High Risk (Critical Audit)"
        elif vri >= 35:
            return "Medium Risk (Sample Audit)"
        else:
            return "Low Risk (Auto-Approve)"
            
    merged["risk_tier"] = merged["vri_score"].apply(assign_tier)
    return merged.sort_values(by="vri_score", ascending=False)
