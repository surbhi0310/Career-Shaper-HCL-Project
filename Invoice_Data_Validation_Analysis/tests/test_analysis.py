"""
test_analysis.py - Validation & Analytical Rules Tests
Project: PJ-IVD-2026 | Surbhi Sharma (2500320140152) | ABES Engineering College
"""

import pytest
import pandas as pd
import numpy as np
from src.validation import (
    validate_arithmetic,
    detect_duplicates,
    validate_tax_rates,
    reconcile_po_variance,
    calculate_vendor_risk_index
)

def test_validate_arithmetic_detection():
    # Synthetic test data: 1 valid line, 1 invalid line
    data = pd.DataFrame([
        {"item_id": "L1", "quantity": 10, "unit_price": 100.0, "tax_rate": 0.18, "line_total": 1180.0},
        {"item_id": "L2", "quantity": 5, "unit_price": 200.0, "tax_rate": 0.18, "line_total": 1300.0}  # Expected 1180, diff +120
    ])
    result = validate_arithmetic(data)
    assert not result.loc[0, "arithmetic_error"]
    assert result.loc[1, "arithmetic_error"]
    assert result.loc[1, "error_variance"] == 120.0

def test_detect_duplicates_logic():
    data = pd.DataFrame([
        {"invoice_id": "INV-01", "vendor_id": "VEN-01", "total_amount": 50000.0, "invoice_date": "2026-03-01"},
        {"invoice_id": "INV-02", "vendor_id": "VEN-01", "total_amount": 50000.0, "invoice_date": "2026-03-05"}, # within 4 days (duplicate)
        {"invoice_id": "INV-03", "vendor_id": "VEN-01", "total_amount": 50000.0, "invoice_date": "2026-04-20"}, # 46 days later (legitimate)
        {"invoice_id": "INV-04", "vendor_id": "VEN-02", "total_amount": 50000.0, "invoice_date": "2026-03-02"}  # different vendor
    ])
    dups = detect_duplicates(data, days_window=14)
    assert len(dups) == 1
    assert dups.iloc[0]["primary_invoice"] == "INV-01"
    assert dups.iloc[0]["duplicate_invoice"] == "INV-02"

def test_validate_tax_slabs():
    data = pd.DataFrame([
        {"item_id": "L1", "tax_rate": 0.18},
        {"item_id": "L2", "tax_rate": 0.05},
        {"item_id": "L3", "tax_rate": 0.15} # Non-standard slab!
    ])
    result = validate_tax_rates(data)
    assert not result.loc[0, "invalid_tax_rate"]
    assert not result.loc[1, "invalid_tax_rate"]
    assert result.loc[2, "invalid_tax_rate"]

def test_po_reconciliation():
    invs = pd.DataFrame([
        {"invoice_id": "I1", "po_id": "P1", "total_amount": 100000.0},
        {"invoice_id": "I2", "po_id": "P2", "total_amount": 125000.0} # +25% over PO budget
    ])
    pos = pd.DataFrame([
        {"po_id": "P1", "po_amount": 100000.0},
        {"po_id": "P2", "po_amount": 100000.0}
    ])
    recon = reconcile_po_variance(invs, pos, tolerance_pct=0.02)
    assert len(recon) == 2
    assert not recon.loc[0, "is_over_budget"]
    assert recon.loc[1, "is_over_budget"]

def test_vendor_risk_score_bounds():
    vendors = pd.DataFrame([
        {"vendor_id": "V1", "vendor_name": "Reliable Corp", "baseline_reliability": 0.98},
        {"vendor_id": "V2", "vendor_name": "Rogue Corp", "baseline_reliability": 0.60}
    ])
    invoices = pd.DataFrame([
        {"invoice_id": f"I{i}", "vendor_id": "V1", "total_amount": 10000.0} for i in range(10)
    ] + [
        {"invoice_id": f"I{i+10}", "vendor_id": "V2", "total_amount": 10000.0} for i in range(10)
    ])
    discrepancies = pd.DataFrame([
        {"disc_id": "D1", "vendor_id": "V2", "severity": "Critical", "variance_amount": 5000.0},
        {"disc_id": "D2", "vendor_id": "V2", "severity": "High", "variance_amount": 2500.0},
        {"disc_id": "D3", "vendor_id": "V2", "severity": "High", "variance_amount": 3500.0}
    ])
    vri = calculate_vendor_risk_index(vendors, discrepancies, invoices)
    assert len(vri) == 2
    rogue_row = vri[vri["vendor_id"] == "V2"].iloc[0]
    good_row = vri[vri["vendor_id"] == "V1"].iloc[0]
    assert rogue_row["vri_score"] > good_row["vri_score"]
    assert rogue_row["risk_tier"] in ["High Risk (Critical Audit)", "Medium Risk (Sample Audit)"]
