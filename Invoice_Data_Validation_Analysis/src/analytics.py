"""
analytics.py - Statistical Analytics & EDA Operations
Project: PJ-IVD-2026 | Surbhi Sharma (2500320140152) | ABES Engineering College
"""

import pandas as pd
import numpy as np
from scipy import stats

def compute_benford_distribution(numbers_series):
    """
    Computes first-digit distribution of numeric values against theoretical Benford's Law.
    Used by auditors to detect artificial/fraudulent amounts.
    """
    digits = []
    for val in numbers_series:
        try:
            val_clean = str(abs(float(val))).lstrip("0").replace(".", "")
            if val_clean:
                first_d = int(val_clean[0])
                if 1 <= first_d <= 9:
                    digits.append(first_d)
        except (ValueError, TypeError):
            continue
            
    if not digits:
        return pd.DataFrame()
        
    actual_counts = pd.Series(digits).value_counts().sort_index()
    total_obs = len(digits)
    
    benford_df = pd.DataFrame({"digit": range(1, 10)})
    benford_df["theoretical_pct"] = np.round(np.log10(1 + 1 / benford_df["digit"]) * 100, 2)
    benford_df["actual_count"] = benford_df["digit"].map(actual_counts).fillna(0).astype(int)
    benford_df["actual_pct"] = np.round((benford_df["actual_count"] / total_obs) * 100, 2)
    benford_df["deviation_pct"] = np.round(benford_df["actual_pct"] - benford_df["theoretical_pct"], 2)
    
    # Chi-square goodness-of-fit test
    expected_counts = (benford_df["theoretical_pct"] / 100) * total_obs
    chi2_stat, p_val = stats.chisquare(f_obs=benford_df["actual_count"], f_exp=expected_counts)
    
    return benford_df, chi2_stat, p_val

def detect_outliers_iqr(series, multiplier=1.5):
    """
    Identifies outliers using Interquartile Range (IQR) method.
    """
    q25 = np.percentile(series.dropna(), 25)
    q75 = np.percentile(series.dropna(), 75)
    iqr = q75 - q25
    lower_bound = q25 - (multiplier * iqr)
    upper_bound = q75 + (multiplier * iqr)
    is_outlier = (series < lower_bound) | (series > upper_bound)
    return is_outlier, lower_bound, upper_bound

def detect_outliers_zscore(series, threshold=3.0):
    """
    Identifies outliers using standard Z-score method.
    """
    clean_s = series.dropna()
    mean = np.mean(clean_s)
    std = np.std(clean_s)
    if std == 0:
        return pd.Series(False, index=series.index), 0, 0
    z_scores = np.abs((series - mean) / std)
    return z_scores > threshold, mean, std

def get_monthly_spend_trend(invoices_df):
    """
    Aggregates invoice count, total spend, and discrepancy trends by month.
    """
    df = invoices_df.copy()
    df["month"] = pd.to_datetime(df["invoice_date"]).dt.to_period("M").astype(str)
    
    trend = df.groupby("month").agg(
        invoice_count=("invoice_id", "count"),
        total_spend=("total_amount", "sum"),
        discrepancies=("validation_status", lambda x: (x == "DISCREPANCY_DETECTED").sum())
    ).reset_index()
    
    trend["discrepancy_rate_pct"] = np.round((trend["discrepancies"] / trend["invoice_count"]) * 100, 2)
    return trend
