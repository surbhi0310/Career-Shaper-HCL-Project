"""
test_data.py - Unit & Database Integrity Tests
Project: PJ-IVD-2026 | Surbhi Sharma (2500320140152) | ABES Engineering College
"""

import os
import sqlite3
import pytest
import pandas as pd
from src.database import get_connection, load_table, DB_PATH

def test_database_file_exists():
    assert os.path.exists(DB_PATH), f"SQLite database file {DB_PATH} must exist."

def test_tables_row_counts():
    conn = get_connection()
    cur = conn.cursor()
    
    tables = {
        "vendors": 30,
        "purchase_orders": 150,
        "invoices": 304,
        "line_items": 1056,
        "discrepancies": 69
    }
    
    for tbl, min_count in tables.items():
        cur.execute(f"SELECT COUNT(*) FROM {tbl}")
        count = cur.fetchone()[0]
        assert count >= min_count, f"Table {tbl} has {count} rows, expected at least {min_count}."
    conn.close()

def test_primary_keys_uniqueness():
    conn = get_connection()
    cur = conn.cursor()
    
    pks = [
        ("vendors", "vendor_id"),
        ("purchase_orders", "po_id"),
        ("invoices", "invoice_id"),
        ("line_items", "item_id"),
        ("discrepancies", "disc_id")
    ]
    
    for tbl, pk in pks:
        cur.execute(f"SELECT COUNT({pk}), COUNT(DISTINCT {pk}) FROM {tbl}")
        total, distinct = cur.fetchone()
        assert total == distinct, f"Primary key {pk} in table {tbl} is not unique!"
    conn.close()

def test_foreign_key_integrity():
    conn = get_connection()
    cur = conn.cursor()
    
    # Check purchase_orders -> vendors
    cur.execute("SELECT COUNT(*) FROM purchase_orders WHERE vendor_id NOT IN (SELECT vendor_id FROM vendors)")
    orphan_pos = cur.fetchone()[0]
    assert orphan_pos == 0, f"Found {orphan_pos} orphan purchase orders without valid vendor."
    
    # Check invoices -> vendors
    cur.execute("SELECT COUNT(*) FROM invoices WHERE vendor_id NOT IN (SELECT vendor_id FROM vendors)")
    orphan_invs = cur.fetchone()[0]
    assert orphan_invs == 0, f"Found {orphan_invs} orphan invoices without valid vendor."
    
    # Check line_items -> invoices
    cur.execute("SELECT COUNT(*) FROM line_items WHERE invoice_id NOT IN (SELECT invoice_id FROM invoices)")
    orphan_lines = cur.fetchone()[0]
    assert orphan_lines == 0, f"Found {orphan_lines} orphan line items without valid invoice."
    conn.close()
