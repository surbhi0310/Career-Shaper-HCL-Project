"""
generate_data.py - Synthetic Dataset Generator for Invoice Data Validation Platform
Project: P053 | Surbhi Sharma (2500320140152) | ABES Engineering College
"""

import os
import random
import sqlite3
import datetime
import pandas as pd
import numpy as np

def generate_synthetic_data(data_dir="data", seed=42):
    random.seed(seed)
    np.random.seed(seed)
    os.makedirs(data_dir, exist_ok=True)
    
    # 1. Generate Vendors (30 Master Records)
    indian_cities = ["Noida, UP", "Gurugram, HR", "Bengaluru, KA", "Mumbai, MH", "Hyderabad, TS", "Chennai, TN", "Pune, MH", "Ahmedabad, GJ", "Kolkata, WB", "Jaipur, RJ"]
    categories = ["IT Hardware & Servers", "Cloud Infrastructure", "Office Supplies & Furniture", "Logistics & Freight", "Industrial Machinery", "Facility Management", "Consulting & Legal", "Security Services"]
    
    vendor_names = [
        "Tata Consultancy Infotech Ltd", "Infosys Enterprise Hardware Ltd", "Larsen & Toubro Electricals Ltd",
        "Reliance Digital Supplies Pvt Ltd", "Mahindra Industrial Logistics", "Godrej Corporate Furnishings",
        "Wipro Cloud Solutions India", "Adani Power & Utilities Ltd", "Bharat Heavy Industrial Equipment",
        "Havells Commercial Electricals", "Blue Dart Express Logistics", "HCL Services & Peripherals",
        "TVS Supply Chain Solutions", "Voltas Facility Aircon Ltd", "Tech Mahindra Digital Networks",
        "Bajaj Corporate Solutions", "Kirloskar Power Systems", "Asian Paints Industrial Coating",
        "Pidilite Construction Chemicals", "Schneider Electric India", "Siemens Healthineers India",
        "Bosch Automotive Services", "Sterling & Wilson Energy", "Thermax Boiler Systems",
        "Crompton Greaves Commercial", "Titan Corporate Gifting", "ITC Paperboards & Packaging",
        "DCM Shriram Enterprises", "Polycab Wires & Cables Ltd", "Apollo Tyres Fleet Supplies"
    ]
    
    vendors = []
    state_codes = ["07", "09", "27", "29", "36", "33", "24", "19", "08", "06"]
    for i, name in enumerate(vendor_names, start=1):
        v_id = f"VEN-{i:03d}"
        state_code = random.choice(state_codes)
        pan = f"AAAC{random.randint(1000, 9999)}{chr(65 + (i % 26))}"
        gstin = f"{state_code}{pan}1Z{random.randint(1, 9)}"
        category = random.choice(categories)
        city = random.choice(indian_cities)
        payment_terms = random.choice([15, 30, 45, 60])
        # Mark 3 vendors as high-risk rogue vendors for pattern testing
        is_rogue = 1 if i in [4, 12, 23] else 0
        baseline_reliability = round(random.uniform(0.55, 0.75) if is_rogue else random.uniform(0.92, 0.99), 2)
        
        vendors.append({
            "vendor_id": v_id,
            "vendor_name": name,
            "gstin": gstin,
            "category": category,
            "city": city,
            "payment_terms_days": payment_terms,
            "is_high_risk": is_rogue,
            "baseline_reliability": baseline_reliability
        })
    df_vendors = pd.DataFrame(vendors)

    # 2. Generate Purchase Orders (150 POs)
    pos = []
    base_date = datetime.date(2026, 1, 1)
    for po_idx in range(1, 151):
        po_id = f"PO-2026-{po_idx:04d}"
        v = random.choice(vendors)
        v_id = v["vendor_id"]
        po_days = random.randint(1, 280)
        po_date = base_date + datetime.timedelta(days=po_days)
        budget = round(random.uniform(25000, 850000), 2)
        status = random.choices(["Approved", "Fulfilled", "Partially Fulfilled", "Closed"], weights=[0.2, 0.5, 0.25, 0.05])[0]
        
        pos.append({
            "po_id": po_id,
            "vendor_id": v_id,
            "po_date": po_date.strftime("%Y-%m-%d"),
            "po_amount": budget,
            "status": status,
            "department": random.choice(["Operations", "Information Technology", "Procurement", "Facilities", "Supply Chain", "Finance"])
        })
    df_pos = pd.DataFrame(pos)

    # 3. Generate Invoices (300 Invoices) & Line Items (approx 1,200 records)
    invoices = []
    line_items = []
    discrepancy_logs = []
    
    line_item_counter = 1
    disc_counter = 1
    
    standard_tax_slabs = [0.05, 0.12, 0.18, 0.28]
    item_catalog = [
        ("Server Rack Chassis 42U", "IT Hardware & Servers", 75000.0, 0.18),
        ("Cisco Core Switch 24-Port", "IT Hardware & Servers", 45000.0, 0.18),
        ("Dell PowerEdge R750 Server", "IT Hardware & Servers", 185000.0, 0.18),
        ("Ergonomic Mesh Office Chair", "Office Supplies & Furniture", 8500.0, 0.18),
        ("Modular Workstation Desk", "Office Supplies & Furniture", 16000.0, 0.18),
        ("Industrial Diesel Generator 25kVA", "Industrial Machinery", 320000.0, 0.28),
        ("Heavy Pallet Hydraulic Jack", "Industrial Machinery", 28000.0, 0.18),
        ("Logistics Freight Container Shipment", "Logistics & Freight", 42000.0, 0.05),
        ("Interstate Fleet Haulage Service", "Logistics & Freight", 35000.0, 0.05),
        ("Facility HVAC Annual Maintenance", "Facility Management", 95000.0, 0.18),
        ("Corporate Legal Due Diligence", "Consulting & Legal", 120000.0, 0.18),
        ("Security Guard Force Monthly Retainer", "Security Services", 65000.0, 0.18),
        ("Industrial Coating Resin 50L Drum", "Industrial Machinery", 14500.0, 0.18),
        ("Copper Armoured Cable 100m Roll", "Industrial Machinery", 19500.0, 0.18),
        ("Packaging Corrugated Carton 500pk", "Office Supplies & Furniture", 6200.0, 0.12)
    ]
    
    for inv_idx in range(1, 301):
        inv_id = f"INV-2026-{inv_idx:04d}"
        matched_po = random.choice(pos)
        v_id = matched_po["vendor_id"]
        v_meta = next(v for v in vendors if v["vendor_id"] == v_id)
        
        po_dt = datetime.datetime.strptime(matched_po["po_date"], "%Y-%m-%d").date()
        inv_lag = random.randint(3, 45)
        inv_date = po_dt + datetime.timedelta(days=inv_lag)
        due_date = inv_date + datetime.timedelta(days=v_meta["payment_terms_days"])
        
        # Decide if this invoice will have injected discrepancies
        # Higher probability for rogue vendors
        inject_error = False
        disc_type = None
        
        rand_val = random.random()
        if v_meta["is_high_risk"] == 1:
            if rand_val < 0.65:
                inject_error = True
        else:
            if rand_val < 0.12:
                inject_error = True
        
        num_items = random.randint(1, 6)
        inv_subtotal = 0.0
        inv_tax_total = 0.0
        
        current_inv_items = []
        for _ in range(num_items):
            item_desc, cat, u_price, t_rate = random.choice(item_catalog)
            # Add minor price variance to unit price
            u_price = round(u_price * random.uniform(0.95, 1.05), 2)
            qty = random.randint(1, 15)
            
            # Intentionally inject arithmetic discrepancy on 1st item if targeted
            calc_tax = round(qty * u_price * t_rate, 2)
            calc_line_total = round((qty * u_price) + calc_tax, 2)
            
            line_has_error = False
            if inject_error and disc_type is None and random.random() < 0.5:
                # Discrepancy Type 1: Arithmetic error (overcharging total)
                disc_type = "Arithmetic Mismatch"
                error_amount = round(random.uniform(500, 7500), 2)
                calc_line_total += error_amount
                line_has_error = True
                discrepancy_logs.append({
                    "disc_id": f"DISC-{disc_counter:04d}",
                    "invoice_id": inv_id,
                    "po_id": matched_po["po_id"],
                    "vendor_id": v_id,
                    "disc_type": "Arithmetic Mismatch",
                    "severity": "High",
                    "variance_amount": error_amount,
                    "root_cause": f"Quantity x Price + Tax != Total. Discrepancy of INR {error_amount:,.2f}",
                    "status": "Under Review"
                })
                disc_counter += 1
            elif inject_error and disc_type is None and random.random() < 0.35:
                # Discrepancy Type 3: Tax Slab Mismatch
                disc_type = "Tax Rate Mismatch"
                wrong_tax_rate = 0.05 if t_rate == 0.18 else 0.28
                tax_variance = round(abs(wrong_tax_rate - t_rate) * (qty * u_price), 2)
                t_rate = wrong_tax_rate
                calc_tax = round(qty * u_price * t_rate, 2)
                calc_line_total = round((qty * u_price) + calc_tax, 2)
                line_has_error = True
                discrepancy_logs.append({
                    "disc_id": f"DISC-{disc_counter:04d}",
                    "invoice_id": inv_id,
                    "po_id": matched_po["po_id"],
                    "vendor_id": v_id,
                    "disc_type": "Tax Rate Mismatch",
                    "severity": "Medium",
                    "variance_amount": tax_variance,
                    "root_cause": f"Statutory tax rate violated. Applied {t_rate*100}% instead of standard bracket.",
                    "status": "Escalated to Vendor"
                })
                disc_counter += 1

            inv_subtotal += (qty * u_price)
            inv_tax_total += calc_tax
            
            line_items.append({
                "item_id": f"LIT-{line_item_counter:05d}",
                "invoice_id": inv_id,
                "item_description": item_desc,
                "category": cat,
                "quantity": qty,
                "unit_price": u_price,
                "tax_rate": t_rate,
                "tax_amount": calc_tax,
                "line_total": calc_line_total,
                "has_discrepancy": 1 if line_has_error else 0
            })
            line_item_counter += 1

        inv_subtotal = round(inv_subtotal, 2)
        inv_tax_total = round(inv_tax_total, 2)
        inv_final_total = round(inv_subtotal + inv_tax_total, 2)

        # Check for PO Price/Quantity drift (Discrepancy Type 4)
        if inject_error and disc_type is None and random.random() < 0.4:
            disc_type = "PO Budget Overrun"
            po_overrun = round(inv_final_total - matched_po["po_amount"], 2)
            if po_overrun > 0:
                discrepancy_logs.append({
                    "disc_id": f"DISC-{disc_counter:04d}",
                    "invoice_id": inv_id,
                    "po_id": matched_po["po_id"],
                    "vendor_id": v_id,
                    "disc_type": "PO Budget Overrun",
                    "severity": "High",
                    "variance_amount": po_overrun,
                    "root_cause": f"Billed amount exceeds authorized PO budget by INR {po_overrun:,.2f}",
                    "status": "Pending Management Approval"
                })
                disc_counter += 1

        # Data Quality Issue injection: 2.5% Missing PO Reference
        assigned_po = matched_po["po_id"]
        if random.random() < 0.025:
            assigned_po = None
            discrepancy_logs.append({
                "disc_id": f"DISC-{disc_counter:04d}",
                "invoice_id": inv_id,
                "po_id": "UNMAPPED",
                "vendor_id": v_id,
                "disc_type": "Missing PO Reference",
                "severity": "Medium",
                "variance_amount": inv_final_total,
                "root_cause": "Invoice received without an approved PO reference number.",
                "status": "Blocked"
            })
            disc_counter += 1

        inv_status = "Flagged" if disc_type is not None else random.choices(["Approved", "Paid", "Pending Approval"], weights=[0.4, 0.45, 0.15])[0]

        invoices.append({
            "invoice_id": inv_id,
            "invoice_number": f"INV/2026/{random.randint(10000, 99999)}",
            "po_id": assigned_po,
            "vendor_id": v_id,
            "invoice_date": inv_date.strftime("%Y-%m-%d"),
            "due_date": due_date.strftime("%Y-%m-%d"),
            "subtotal": inv_subtotal,
            "tax_amount": inv_tax_total,
            "total_amount": inv_final_total,
            "currency": "INR",
            "validation_status": "DISCREPANCY_DETECTED" if disc_type else "VALIDATED",
            "payment_status": inv_status
        })

    # Injected Discrepancy Type 2: Intentional Duplicates (4 duplicate pairs)
    for dup_idx in range(4):
        target_inv = random.choice(invoices[:100])
        dup_inv_id = f"INV-2026-{301 + dup_idx:04d}"
        dup_date = datetime.datetime.strptime(target_inv["invoice_date"], "%Y-%m-%d").date() + datetime.timedelta(days=random.randint(2, 10))
        dup_due = dup_date + datetime.timedelta(days=30)
        
        invoices.append({
            "invoice_id": dup_inv_id,
            "invoice_number": target_inv["invoice_number"] + "-DUP",
            "po_id": target_inv["po_id"],
            "vendor_id": target_inv["vendor_id"],
            "invoice_date": dup_date.strftime("%Y-%m-%d"),
            "due_date": dup_due.strftime("%Y-%m-%d"),
            "subtotal": target_inv["subtotal"],
            "tax_amount": target_inv["tax_amount"],
            "total_amount": target_inv["total_amount"],
            "currency": "INR",
            "validation_status": "DISCREPANCY_DETECTED",
            "payment_status": "Flagged"
        })
        
        discrepancy_logs.append({
            "disc_id": f"DISC-{disc_counter:04d}",
            "invoice_id": dup_inv_id,
            "po_id": target_inv["po_id"] if target_inv["po_id"] else "UNMAPPED",
            "vendor_id": target_inv["vendor_id"],
            "disc_type": "Duplicate Billing Collision",
            "severity": "Critical",
            "variance_amount": target_inv["total_amount"],
            "root_cause": f"Duplicate invoice detected for vendor {target_inv['vendor_id']} matching {target_inv['invoice_id']} (Amount: INR {target_inv['total_amount']:,.2f})",
            "status": "Rejected"
        })
        disc_counter += 1

    df_invoices = pd.DataFrame(invoices)
    df_line_items = pd.DataFrame(line_items)
    df_discrepancies = pd.DataFrame(discrepancy_logs)

    # 4. Save to CSVs
    df_vendors.to_csv(os.path.join(data_dir, "vendors.csv"), index=False)
    df_pos.to_csv(os.path.join(data_dir, "purchase_orders.csv"), index=False)
    df_invoices.to_csv(os.path.join(data_dir, "invoices.csv"), index=False)
    df_line_items.to_csv(os.path.join(data_dir, "line_items.csv"), index=False)
    df_discrepancies.to_csv(os.path.join(data_dir, "discrepancies.csv"), index=False)
    
    # 5. Save to Multi-tab Excel file
    excel_path = os.path.join(data_dir, "invoice_data.xlsx")
    with pd.ExcelWriter(excel_path, engine="openpyxl") as writer:
        df_vendors.to_excel(writer, sheet_name="Vendors", index=False)
        df_pos.to_excel(writer, sheet_name="Purchase_Orders", index=False)
        df_invoices.to_excel(writer, sheet_name="Invoices", index=False)
        df_line_items.to_excel(writer, sheet_name="Line_Items", index=False)
        df_discrepancies.to_excel(writer, sheet_name="Discrepancies", index=False)

    # 6. Build SQLite Database with Foreign Keys and Constraints
    db_path = os.path.join(data_dir, "app.db")
    if os.path.exists(db_path):
        os.remove(db_path)
        
    conn = sqlite3.connect(db_path)
    cursor = conn.cursor()
    cursor.execute("PRAGMA foreign_keys = ON;")
    
    # SQLite DDL
    cursor.execute("""
    CREATE TABLE vendors (
        vendor_id TEXT PRIMARY KEY,
        vendor_name TEXT NOT NULL,
        gstin TEXT NOT NULL UNIQUE,
        category TEXT NOT NULL,
        city TEXT NOT NULL,
        payment_terms_days INTEGER NOT NULL,
        is_high_risk INTEGER NOT NULL DEFAULT 0,
        baseline_reliability REAL NOT NULL
    );
    """)
    
    cursor.execute("""
    CREATE TABLE purchase_orders (
        po_id TEXT PRIMARY KEY,
        vendor_id TEXT NOT NULL,
        po_date TEXT NOT NULL,
        po_amount REAL NOT NULL,
        status TEXT NOT NULL,
        department TEXT NOT NULL,
        FOREIGN KEY (vendor_id) REFERENCES vendors(vendor_id)
    );
    """)

    cursor.execute("""
    CREATE TABLE invoices (
        invoice_id TEXT PRIMARY KEY,
        invoice_number TEXT NOT NULL,
        po_id TEXT,
        vendor_id TEXT NOT NULL,
        invoice_date TEXT NOT NULL,
        due_date TEXT NOT NULL,
        subtotal REAL NOT NULL,
        tax_amount REAL NOT NULL,
        total_amount REAL NOT NULL,
        currency TEXT NOT NULL DEFAULT 'INR',
        validation_status TEXT NOT NULL,
        payment_status TEXT NOT NULL,
        FOREIGN KEY (vendor_id) REFERENCES vendors(vendor_id),
        FOREIGN KEY (po_id) REFERENCES purchase_orders(po_id)
    );
    """)

    cursor.execute("""
    CREATE TABLE line_items (
        item_id TEXT PRIMARY KEY,
        invoice_id TEXT NOT NULL,
        item_description TEXT NOT NULL,
        category TEXT NOT NULL,
        quantity INTEGER NOT NULL,
        unit_price REAL NOT NULL,
        tax_rate REAL NOT NULL,
        tax_amount REAL NOT NULL,
        line_total REAL NOT NULL,
        has_discrepancy INTEGER NOT NULL DEFAULT 0,
        FOREIGN KEY (invoice_id) REFERENCES invoices(invoice_id)
    );
    """)

    cursor.execute("""
    CREATE TABLE discrepancies (
        disc_id TEXT PRIMARY KEY,
        invoice_id TEXT NOT NULL,
        po_id TEXT,
        vendor_id TEXT NOT NULL,
        disc_type TEXT NOT NULL,
        severity TEXT NOT NULL,
        variance_amount REAL NOT NULL,
        root_cause TEXT NOT NULL,
        status TEXT NOT NULL,
        FOREIGN KEY (invoice_id) REFERENCES invoices(invoice_id),
        FOREIGN KEY (vendor_id) REFERENCES vendors(vendor_id)
    );
    """)

    # Create analytical view (joined 32 columns)
    cursor.execute("""
    CREATE VIEW v_invoice_analysis AS
    SELECT 
        i.invoice_id,
        i.invoice_number,
        i.invoice_date,
        i.due_date,
        i.subtotal AS inv_subtotal,
        i.tax_amount AS inv_tax,
        i.total_amount AS inv_total,
        i.validation_status,
        i.payment_status,
        v.vendor_id,
        v.vendor_name,
        v.gstin,
        v.category AS vendor_category,
        v.city AS vendor_city,
        v.payment_terms_days,
        v.is_high_risk,
        v.baseline_reliability,
        p.po_id,
        p.po_date,
        p.po_amount,
        p.department,
        p.status AS po_status,
        (i.total_amount - p.po_amount) AS po_variance,
        d.disc_id,
        d.disc_type,
        d.severity AS disc_severity,
        d.variance_amount AS disc_variance,
        d.root_cause,
        d.status AS disc_status
    FROM invoices i
    LEFT JOIN vendors v ON i.vendor_id = v.vendor_id
    LEFT JOIN purchase_orders p ON i.po_id = p.po_id
    LEFT JOIN discrepancies d ON i.invoice_id = d.invoice_id;
    """)

    conn.commit()

    # Insert Data into SQLite
    df_vendors.to_sql("vendors", conn, if_exists="append", index=False)
    df_pos.to_sql("purchase_orders", conn, if_exists="append", index=False)
    df_invoices.to_sql("invoices", conn, if_exists="append", index=False)
    df_line_items.to_sql("line_items", conn, if_exists="append", index=False)
    df_discrepancies.to_sql("discrepancies", conn, if_exists="append", index=False)
    
    conn.close()

    print(f"Data Generation Completed Successfully:")
    print(f" - Vendors: {len(df_vendors)} records")
    print(f" - Purchase Orders: {len(df_pos)} records")
    print(f" - Invoices: {len(df_invoices)} records")
    print(f" - Line Items: {len(df_line_items)} records")
    print(f" - Discrepancies Logged: {len(df_discrepancies)} anomalies")
    print(f" - SQLite DB: {db_path}")
    print(f" - Excel Workbook: {excel_path}")

if __name__ == "__main__":
    generate_synthetic_data()
