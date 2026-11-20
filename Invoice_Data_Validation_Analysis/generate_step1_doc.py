import docx
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import parse_xml, OxmlElement
from docx.oxml.ns import nsdecls, qn

def set_cell_background(cell, fill_hex):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{fill_hex}"/>')
    tcPr.append(shd)

def set_cell_margins(cell, top=100, bottom=100, left=150, right=150):
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = parse_xml(f'<w:tcMar {nsdecls("w")}><w:top w:w="{top}" w:type="dxa"/><w:bottom w:w="{bottom}" w:type="dxa"/><w:left w:w="{left}" w:type="dxa"/><w:right w:w="{right}" w:type="dxa"/></w:tcMar>')
    tcPr.append(tcMar)

def style_table_header(row, col_widths, bg_hex="1E3A8A", font_color="FFFFFF"):
    for idx, cell in enumerate(row.cells):
        set_cell_background(cell, bg_hex)
        set_cell_margins(cell, top=140, bottom=140, left=150, right=150)
        cell.width = col_widths[idx]
        for p in cell.paragraphs:
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT
            for r in p.runs:
                r.bold = True
                r.font.name = "Calibri"
                r.font.size = Pt(10)
                r.font.color.rgb = RGBColor.from_string(font_color)

def style_table_row(row, col_widths, is_even=False):
    bg_hex = "F8FAFC" if is_even else "FFFFFF"
    for idx, cell in enumerate(row.cells):
        set_cell_background(cell, bg_hex)
        set_cell_margins(cell, top=100, bottom=100, left=150, right=150)
        cell.width = col_widths[idx]
        for p in cell.paragraphs:
            for r in p.runs:
                r.font.name = "Calibri"
                r.font.size = Pt(9.5)
                r.font.color.rgb = RGBColor.from_string("333333")

def create_document():
    doc = Document()
    
    # Page Margins: 1 inch all around
    sections = doc.sections
    for section in sections:
        section.top_margin = Inches(1)
        section.bottom_margin = Inches(1)
        section.left_margin = Inches(1)
        section.right_margin = Inches(1)
    
    # Set default styles
    normal_style = doc.styles['Normal']
    normal_style.font.name = 'Calibri'
    normal_style.font.size = Pt(10.5)
    normal_style.font.color.rgb = RGBColor.from_string('2D3748')
    normal_style.paragraph_format.line_spacing = 1.15
    normal_style.paragraph_format.space_after = Pt(6)

    # Document Header / Pre-title
    p_meta = doc.add_paragraph()
    r_meta = p_meta.add_run("ABES ENGINEERING COLLEGE | DEPARTMENT OF MASTER OF COMPUTER APPLICATIONS\nMCA 2nd Year — Data Analytics Specialization | Academic Project Deliverable")
    r_meta.font.size = Pt(8.5)
    r_meta.font.bold = True
    r_meta.font.color.rgb = RGBColor.from_string("64748B")
    p_meta.alignment = WD_ALIGN_PARAGRAPH.RIGHT

    # Main Title
    p_title = doc.add_paragraph()
    r_title = p_title.add_run("STEP 1: PROJECT UNDERSTANDING & CHARTER")
    r_title.font.name = "Calibri"
    r_title.font.size = Pt(22)
    r_title.font.bold = True
    r_title.font.color.rgb = RGBColor.from_string("1E3A8A")
    p_title.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p_title.paragraph_format.space_after = Pt(2)

    p_sub = doc.add_paragraph()
    r_sub = p_sub.add_run("Project ID: PJ-IVD-2026 | Invoice Data Validation & Discrepancy Analytics Platform")
    r_sub.font.size = Pt(13)
    r_sub.font.bold = True
    r_sub.font.color.rgb = RGBColor.from_string("475569")
    p_sub.paragraph_format.space_after = Pt(14)

    # Metadata Table
    table_meta = doc.add_table(rows=6, cols=2)
    table_meta.alignment = WD_TABLE_ALIGNMENT.CENTER
    meta_widths = [Inches(2.2), Inches(4.3)]
    
    metadata_entries = [
        ("Student Name / Candidate", "Surbhi Sharma"),
        ("University Roll Number", "2500320140152"),
        ("Program & Specialization", "Master of Computer Applications (MCA 2nd Year) — Data Analytics"),
        ("Institution", "ABES Engineering College, Ghaziabad"),
        ("Project Mentor / Guide", "Kalpana Dwivedi"),
        ("Domain & Technology Stack", "Financial Analytics / AP Automation (Python, Streamlit, Plotly, SQLite, Pytest)")
    ]

    for i, (k, v) in enumerate(metadata_entries):
        row = table_meta.rows[i]
        c0, c1 = row.cells[0], row.cells[1]
        c0.text = k
        c1.text = v
        c0.paragraphs[0].runs[0].bold = True
        c0.paragraphs[0].runs[0].font.color.rgb = RGBColor.from_string("1E3A8A")
        set_cell_background(c0, "F1F5F9")
        set_cell_background(c1, "FFFFFF")
        set_cell_margins(c0, 80, 80, 120, 120)
        set_cell_margins(c1, 80, 80, 120, 120)
        c0.width, c1.width = meta_widths[0], meta_widths[1]

    doc.add_paragraph().paragraph_format.space_after = Pt(12)

    # Section 1: Executive Summary & Business Context
    h1 = doc.add_heading(level=1)
    r_h1 = h1.add_run("1. Executive Summary & Business Context")
    r_h1.font.color.rgb = RGBColor.from_string("1E3A8A")
    r_h1.font.bold = True

    doc.add_paragraph(
        "Accounts Payable (AP) and procurement operations form the operational backbone of any enterprise. "
        "Every fiscal year, corporations process thousands of vendor invoices across decentralized purchase orders, "
        "multiple delivery notes, and varied contract frameworks. In manual or legacy enterprise workflows, invoice processing "
        "suffers from severe vulnerabilities: mathematical errors (line-item miscalculations and rounding faults), "
        "tax calculation inconsistencies (GST/VAT discrepancies), duplicate or near-duplicate invoice submissions, and "
        "unreconciled purchase order variances (price creep and quantity mismatches)."
    )
    doc.add_paragraph(
        "Industry benchmarks indicate that 1.5% to 3% of processed B2B invoices contain discrepancies, and undetected "
        "duplicate payments bleed up to 0.5% of total annual corporate disbursements. For an enterprise handling "
        "₹100 Crore in annual procurements, this leakage represents ₹50 Lakhs in avoidable financial losses, alongside "
        "strained vendor relationships, delayed tax input credits, and heightened regulatory audit scrutiny."
    )
    doc.add_paragraph(
        "The Invoice Data Validation & Discrepancy Analytics Platform (PJ-IVD-2026) addresses this critical corporate "
        "challenge by engineering an automated, high-precision data validation and statistical analytics engine. "
        "Built on Python, Streamlit, Plotly, and an ACID-compliant SQLite backend, the application ingests multi-vendor invoice "
        "streams, executes real-time multi-dimensional validation algorithms, scores vendor risk profiles, isolates anomalies "
        "through interactive Exploratory Data Analysis (EDA), and empowers auditors with instantaneous root-cause drill-downs."
    )

    # Section 2: Problem Statement & Opportunity Questions (OQ)
    h2 = doc.add_heading(level=1)
    r_h2 = h2.add_run("2. Problem Statement & Opportunity Questions (OQ)")
    r_h2.font.color.rgb = RGBColor.from_string("1E3A8A")
    r_h2.font.bold = True

    doc.add_paragraph(
        "Enterprise invoice reconciliation currently operates in silos. Human audit teams sample less than 10-15% "
        "of invoices manually, focusing almost exclusively on high-value transactions while letting thousands of small-to-mid "
        "dollar discrepancies bypass inspection. To systematically solve this problem, five primary Opportunity Questions (OQ) "
        "are formulated below to govern system discovery and architecture:"
    )

    table_oq = doc.add_table(rows=6, cols=3)
    table_oq.alignment = WD_TABLE_ALIGNMENT.CENTER
    oq_widths = [Inches(1.0), Inches(2.3), Inches(3.2)]

    oq_headers = ["ID", "Opportunity Question (OQ)", "Architectural Focus & Business Impact"]
    for idx, name in enumerate(oq_headers):
        table_oq.rows[0].cells[idx].text = name
    style_table_header(table_oq.rows[0], oq_widths)

    oq_data = [
        ("OQ-01", "Automated Line-Item Arithmetic & Tax Integrity", 
         "How can statistical validation algorithms verify quantity x unit price + tax = line total across 100% of line items with zero-tolerance precision before ledger posting?"),
        ("OQ-02", "Algorithmic Duplicate & Collision Detection", 
         "What multi-parameter fuzzy matching logic can catch duplicate billings (identical vendor, invoice number variations, matching amount within rolling time windows) without false flagging legitimate repeat orders?"),
        ("OQ-03", "PO-to-Invoice 2-Way & 3-Way Variance Profiling", 
         "How can visual Exploratory Data Analysis (EDA) profile price variance and quantity drift between approved Purchase Orders and received Invoices to reveal vendor overbilling patterns?"),
        ("OQ-04", "Quantitative Vendor Risk Scoring & Profiling", 
         "How can historical error frequencies, tax mismatches, and compliance violations be aggregated into a composite Vendor Risk Index (VRI) to optimize procurement renegotiations and audit sampling?"),
        ("OQ-05", "Interactive Auditor Triaging & Actionable EDA", 
         "How can an intuitive, responsive Streamlit dashboard provide sub-second latency drill-down into discrepant transactions, generating downloadable audit exception manifests?")
    ]

    for i, (oq_id, oq_title, oq_desc) in enumerate(oq_data, start=1):
        row = table_oq.rows[i]
        row.cells[0].text = oq_id
        row.cells[1].text = oq_title
        row.cells[2].text = oq_desc
        row.cells[0].paragraphs[0].runs[0].bold = True
        style_table_row(row, oq_widths, is_even=(i % 2 == 0))

    doc.add_paragraph().paragraph_format.space_after = Pt(10)

    # Section 3: Stakeholder Analysis Matrix (SH)
    h3 = doc.add_heading(level=1)
    r_h3 = h3.add_run("3. Stakeholder Analysis Matrix (SH)")
    r_h3.font.color.rgb = RGBColor.from_string("1E3A8A")
    r_h3.font.bold = True

    doc.add_paragraph(
        "To guarantee that the solution fulfills operational, regulatory, and executive needs, five key stakeholder personas (SH) "
        "have been identified and mapped across their responsibilities, pain points, and core system requirements:"
    )

    table_sh = doc.add_table(rows=6, cols=4)
    table_sh.alignment = WD_TABLE_ALIGNMENT.CENTER
    sh_widths = [Inches(0.9), Inches(1.6), Inches(2.0), Inches(2.0)]

    sh_headers = ["ID", "Stakeholder Persona", "Operational Pain Points", "System Value & Required Capabilities"]
    for idx, name in enumerate(sh_headers):
        table_sh.rows[0].cells[idx].text = name
    style_table_header(table_sh.rows[0], sh_widths)

    sh_data = [
        ("SH-01", "Accounts Payable (AP) Manager", 
         "High cycle times, unmonitored backlogs, delayed early-payment discounts, lack of operational KPI visibility.",
         "Real-time operational dashboard with overall health metrics, discrepancy breakdown, cycle time analytics, and auditor workload distribution."),
        ("SH-02", "Internal Audit & Compliance Officer", 
         "Inability to sample 100% of transactions, risk of GST/tax non-compliance, lack of immutable audit logs for statutory reviews.",
         "Automated 100% population audit, comprehensive discrepancy rules engine, traceable anomaly logging, and exportable audit trail."),
        ("SH-03", "Chief Financial Officer (CFO)", 
         "Unintended cash leakage from duplicate payments, inaccurate working capital forecasting, vendor overcharges.",
         "High-level financial exposure metrics, total prevented cash loss KPIs, working capital risk distribution, and macro-level cost recovery trends."),
        ("SH-04", "Procurement & Vendor Relationship Lead", 
         "Unreliable vendor delivery metrics, lack of empirical data during annual contract renegotiations regarding vendor billing accuracy.",
         "Vendor Risk Index (VRI), vendor error league tables, price drift analysis by item category, and supplier compliance scorecards."),
        ("SH-05", "AP Processing Auditor / Specialist", 
         "Tedious manual cross-referencing between POs and invoices, repetitive calculator checks, fatigue-induced oversight.",
         "Instant automated discrepancy tagging, clear visual highlighting of error root causes, single-click exception report generation.")
    ]

    for i, (sh_id, persona, pain, val) in enumerate(sh_data, start=1):
        row = table_sh.rows[i]
        row.cells[0].text = sh_id
        row.cells[1].text = persona
        row.cells[2].text = pain
        row.cells[3].text = val
        row.cells[0].paragraphs[0].runs[0].bold = True
        style_table_row(row, sh_widths, is_even=(i % 2 == 0))

    doc.add_paragraph().paragraph_format.space_after = Pt(10)

    # Section 4: Project Scope (In-Scope vs. Out-of-Scope)
    h4 = doc.add_heading(level=1)
    r_h4 = h4.add_run("4. Project Scope Definition")
    r_h4.font.color.rgb = RGBColor.from_string("1E3A8A")
    r_h4.font.bold = True

    table_scope = doc.add_table(rows=6, cols=2)
    table_scope.alignment = WD_TABLE_ALIGNMENT.CENTER
    scope_widths = [Inches(3.25), Inches(3.25)]

    scope_headers = ["In-Scope (Delivered Functionality)", "Out-of-Scope (Deferred / Excluded)"]
    for idx, name in enumerate(scope_headers):
        table_scope.rows[0].cells[idx].text = name
    style_table_header(table_scope.rows[0], scope_widths)

    scope_data = [
        ("Multi-table relational schema (Vendors, Purchase Orders, Invoices, Line Items, Discrepancies) hosted on SQLite.", 
         "Direct real-time production ERP database connection (SAP, Oracle, NetSuite) via proprietary enterprise connectors."),
        ("Automated synthetic dataset generation (5,000+ realistic records with injected mathematical, tax, and duplicate anomalies).", 
         "Raw scanned OCR document extraction from mobile camera snaps or unparsed handwritten receipts."),
        ("Algorithmic multi-level validation: arithmetic checks, tax rate cross-verification, duplicate invoice detection, PO-matching.", 
         "Direct automated bank payment release via ACH, NEFT/RTGS, or SWIFT banking APIs."),
        ("Exploratory Data Analysis (EDA) module with interactive Plotly visual distributions, statistical metrics, and boxplots.", 
         "Enterprise Single Sign-On (SSO) with Active Directory or multi-tenant cloud organization partitioning."),
        ("Responsive Streamlit application styled with modern CSS/Tailwind UI cards, sidebar filters, and exportable exception logs.", 
         "Real-time streaming event bus architecture (e.g., Apache Kafka / Flink streaming clusters).")
    ]

    for i, (insc, outsc) in enumerate(scope_data, start=1):
        row = table_scope.rows[i]
        row.cells[0].text = insc
        row.cells[1].text = outsc
        style_table_row(row, scope_widths, is_even=(i % 2 == 0))

    doc.add_paragraph().paragraph_format.space_after = Pt(10)

    # Section 5: Assumptions & Constraints
    h5 = doc.add_heading(level=1)
    r_h5 = h5.add_run("5. Project Assumptions & Technical Constraints")
    r_h5.font.color.rgb = RGBColor.from_string("1E3A8A")
    r_h5.font.bold = True

    doc.add_paragraph(
        "To establish a realistic operating boundary for academic execution and cloud deployment, "
        "the following assumptions and constraints have been established:"
    )

    doc.add_heading(level=2).add_run("5.1 Business & Domain Assumptions")
    p_ass = doc.add_paragraph()
    assumptions = [
        ("A-01: Synthetic Data Integrity — ", "All experimental and operational data will be purely synthetic, mathematically consistent, and free of actual confidential enterprise or proprietary PII."),
        ("A-02: Currency & Precision — ", "Financial amounts are standardized in INR (₹) / USD ($) using two decimal places of monetary precision to eliminate floating-point calculation drift."),
        ("A-03: Two-Way & Three-Way Reconciliation Logic — ", "Validation relies on standard 2-Way (Invoice vs. PO) and simulated 3-Way (Invoice vs. PO vs. Goods Receipt Quantity) matching rules."),
        ("A-04: Standardized Tax Framework — ", "Tax computations follow structured percentage slabs (e.g., 5%, 12%, 18%, 28% GST brackets) allowing deterministic algorithmic verification.")
    ]
    for prefix, body in assumptions:
        p = doc.add_paragraph()
        r1 = p.add_run(prefix)
        r1.bold = True
        r1.font.color.rgb = RGBColor.from_string("1E3A8A")
        p.add_run(body)

    doc.add_heading(level=2).add_run("5.2 Technical Constraints")
    constraints = [
        ("C-01: Deployment Environment — ", "The final application must operate smoothly within Streamlit Community Cloud resource bounds (single-container instance, 1 GB RAM limit)."),
        ("C-02: Embedded Database Engine — ", "The relational backend is strictly constrained to SQLite, requiring zero host configuration while supporting full ACID transactions and relational foreign keys."),
        ("C-03: Testing Standards — ", "All core validation logic, statistical algorithms, and data access layers must be covered by automated Pytest unit and integration suites."),
        ("C-04: Architecture Modularity — ", "The codebase will enforce clean separation of concerns: data generation, database management, validation engine, statistical analytics, and Streamlit presentation UI.")
    ]
    for prefix, body in constraints:
        p = doc.add_paragraph()
        r1 = p.add_run(prefix)
        r1.bold = True
        r1.font.color.rgb = RGBColor.from_string("1E3A8A")
        p.add_run(body)

    # Section 6: Success Criteria & Quantitative Impact
    h6 = doc.add_heading(level=1)
    r_h6 = h6.add_run("6. Quantitative Success Criteria & Evaluation Metrics")
    r_h6.font.color.rgb = RGBColor.from_string("1E3A8A")
    r_h6.font.bold = True

    table_crit = doc.add_table(rows=5, cols=3)
    table_crit.alignment = WD_TABLE_ALIGNMENT.CENTER
    crit_widths = [Inches(1.8), Inches(2.2), Inches(2.5)]

    crit_headers = ["Performance Dimension", "Target Quantitative Metric", "Evaluation Methodology"]
    for idx, name in enumerate(crit_headers):
        table_crit.rows[0].cells[idx].text = name
    style_table_header(table_crit.rows[0], crit_widths)

    crit_data = [
        ("Anomaly Detection Recall", "100% of injected synthetic anomalies detected", "Automated validation engine benchmark against labeled synthetic discrepancy ground truth."),
        ("False Positive Rate", "< 2.0% on valid baseline transactions", "Controlled evaluation run over 5,000 pristine synthetic invoice records."),
        ("System Query Latency", "< 800 ms response time for all dashboard filters", "Performance benchmarking using SQLite indexed queries across 20,000+ line-item records."),
        ("Test Suite Coverage", "> 85% statement coverage across core modules", "Execution of Pytest coverage reports targeting validation rules, database models, and utility functions.")
    ]

    for i, (dim, tgt, meth) in enumerate(crit_data, start=1):
        row = table_crit.rows[i]
        row.cells[0].text = dim
        row.cells[1].text = tgt
        row.cells[2].text = meth
        row.cells[0].paragraphs[0].runs[0].bold = True
        style_table_row(row, crit_widths, is_even=(i % 2 == 0))

    doc.add_paragraph().paragraph_format.space_after = Pt(14)

    # Document Sign-off Block
    h_sign = doc.add_heading(level=2)
    h_sign.add_run("7. Academic Review & Sign-Off Authorization").font.color.rgb = RGBColor.from_string("1E3A8A")

    table_sign = doc.add_table(rows=2, cols=2)
    table_sign.alignment = WD_TABLE_ALIGNMENT.CENTER
    sign_widths = [Inches(3.25), Inches(3.25)]
    
    r0 = table_sign.rows[0]
    r0.cells[0].text = "Candidate Signature:\n\n___________________________\nSurbhi Sharma\nRoll No: 2500320140152\nMCA 2nd Year, ABES Engineering College"
    r0.cells[1].text = "Project Mentor Signature:\n\n___________________________\nKalpana Dwivedi\nAssistant Professor / Project Guide\nDepartment of Computer Applications, ABES"
    
    set_cell_background(r0.cells[0], "F8FAFC")
    set_cell_background(r0.cells[1], "F8FAFC")
    set_cell_margins(r0.cells[0], 120, 120, 150, 150)
    set_cell_margins(r0.cells[1], 120, 120, 150, 150)
    r0.cells[0].width = sign_widths[0]
    r0.cells[1].width = sign_widths[1]

    doc.save("STEP_01_Project_Understanding.docx")
    print("STEP_01_Project_Understanding.docx successfully generated.")

if __name__ == "__main__":
    create_document()
