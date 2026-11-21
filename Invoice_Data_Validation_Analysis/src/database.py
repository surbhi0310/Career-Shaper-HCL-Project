"""
database.py - Database Access Layer
Project: P053 | Surbhi Sharma (2500320140152) | ABES Engineering College
"""

import os
import sqlite3
import pandas as pd

DB_PATH = os.path.join(os.path.dirname(os.path.dirname(__file__)), "data", "app.db")

def get_connection(db_path=DB_PATH):
    conn = sqlite3.connect(db_path)
    conn.row_factory = sqlite3.Row
    return conn

def load_table(table_name, db_path=DB_PATH):
    conn = get_connection(db_path)
    df = pd.read_sql_query(f"SELECT * FROM {table_name}", conn)
    conn.close()
    return df

def load_analysis_view(db_path=DB_PATH):
    conn = get_connection(db_path)
    df = pd.read_sql_query("SELECT * FROM v_invoice_analysis", conn)
    conn.close()
    return df

def get_summary_metrics(db_path=DB_PATH):
    conn = get_connection(db_path)
    cur = conn.cursor()
    
    cur.execute("SELECT COUNT(*), SUM(total_amount) FROM invoices")
    total_invoices, total_spend = cur.fetchone()
    
    cur.execute("SELECT COUNT(*), SUM(variance_amount) FROM discrepancies")
    total_discrepancies, total_leakage = cur.fetchone()
    
    cur.execute("SELECT COUNT(DISTINCT vendor_id) FROM vendors")
    total_vendors = cur.fetchone()[0]

    cur.execute("SELECT COUNT(*) FROM discrepancies WHERE disc_type = 'Duplicate Billing Collision'")
    duplicate_count = cur.fetchone()[0]

    conn.close()
    return {
        "total_invoices": total_invoices or 0,
        "total_spend": total_spend or 0.0,
        "total_discrepancies": total_discrepancies or 0,
        "prevented_leakage": total_leakage or 0.0,
        "total_vendors": total_vendors or 0,
        "duplicate_invoices": duplicate_count or 0,
        "error_rate": round((total_discrepancies / total_invoices * 100), 2) if total_invoices else 0.0
    }
