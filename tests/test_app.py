"""
test_app.py - Integration, Statistical & App Tests
Project: P053 | Surbhi Sharma (2500320140152) | ABES Engineering College
"""

import pytest
import pandas as pd
import numpy as np
from src.database import get_summary_metrics, load_analysis_view
from src.analytics import compute_benford_distribution, detect_outliers_iqr, get_monthly_spend_trend

def test_summary_metrics():
    metrics = get_summary_metrics()
    assert metrics["total_invoices"] >= 300
    assert metrics["total_spend"] > 0
    assert metrics["total_discrepancies"] > 0
    assert metrics["prevented_leakage"] > 0
    assert metrics["total_vendors"] == 30
    assert metrics["duplicate_invoices"] >= 4

def test_benford_distribution():
    # Synthetic invoice amounts
    np.random.seed(42)
    fake_amounts = np.random.lognormal(mean=10, sigma=1.5, size=500)
    df_benford, chi2, p_val = compute_benford_distribution(pd.Series(fake_amounts))
    assert len(df_benford) == 9
    assert df_benford["actual_count"].sum() == 500
    # First digit 1 should be most frequent per Benford's law
    assert df_benford.loc[df_benford["digit"] == 1, "actual_count"].values[0] > df_benford.loc[df_benford["digit"] == 9, "actual_count"].values[0]

def test_outliers_iqr():
    series = pd.Series([10, 12, 11, 14, 13, 12, 15, 100, 11, 12]) # 100 is outlier
    is_outlier, lb, ub = detect_outliers_iqr(series)
    assert is_outlier.sum() == 1
    assert is_outlier.iloc[7] == True

def test_analysis_view_loaded():
    view_df = load_analysis_view()
    assert not view_df.empty
    assert "invoice_id" in view_df.columns
    assert "vendor_name" in view_df.columns
    assert "po_amount" in view_df.columns
    assert "disc_type" in view_df.columns
