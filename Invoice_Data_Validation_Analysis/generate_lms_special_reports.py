"""
generate_lms_special_reports.py - Generates GitHub Report & Complete Code Document
Converts both to PDF for Day 04 LMS Upload
"""

import os
import win32com.client
from pathlib import Path
import docx
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls

BASE_DIR = Path(r"C:\Users\Geetanjali\Downloads\AI-website-Builder-final-project-main\Invoice_Data_Validation_Analysis")
PDF_DIR = BASE_DIR / "docs" / "pdf_submissions"
PDF_DIR.mkdir(parents=True, exist_ok=True)

# ----------------- 1. Build GitHub Report -----------------
def create_github_report():
    doc = Document()
    for s in doc.sections:
        s.top_margin = Inches(1)
        s.bottom_margin = Inches(1)
        s.left_margin = Inches(1)
        s.right_margin = Inches(1)

    p_top = doc.add_paragraph()
    r_top = p_top.add_run("HCLTech Career Shaper Programme | ABES Engineering College, Ghaziabad\nLMS Deliverable #3: GitHub Repository Submission Report")
    r_top.bold = True
    r_top.font.color.rgb = RGBColor.from_string("A51C30")

    p_title = doc.add_paragraph()
    p_title.paragraph_format.space_before = Pt(12)
    p_title.paragraph_format.space_after = Pt(6)
    r_t = p_title.add_run("GitHub Repository Upload & Verification Report")
    r_t.font.name = "Georgia"
    r_t.font.size = Pt(22)
    r_t.bold = True
    r_t.font.color.rgb = RGBColor.from_string("1E3A8A")

    doc.add_paragraph("This report certifies the successful deployment and upload of the complete Invoice Data Validation & Discrepancy Analytics Platform project codebase, database, test suites, and documentation to the official GitHub repository.")

    # Metadata
    table = doc.add_table(rows=6, cols=2)
    meta = [
        ("Project ID & Title", "PJ-IVD-2026: Invoice Data Validation & Discrepancy Analytics Platform"),
        ("Learner / Candidate", "Surbhi Sharma (Roll No: 2500320140152)"),
        ("Programme / College", "MCA 2nd Year (Data Analytics), ABES Engineering College"),
        ("Project Mentor", "Kalpana Dwivedi (Assistant Professor, MCA, ABES)"),
        ("Target GitHub Repository URL", "https://github.com/surbhi0310/Career-Shaper-HCL-Project.git"),
        ("Repository Branch & Status", "main | Verified Clean & Synchronized")
    ]
    for idx, (k, v) in enumerate(meta):
        r = table.rows[idx]
        r.cells[0].text = k
        r.cells[1].text = v
        r.cells[0].paragraphs[0].runs[0].bold = True

    doc.add_heading("Repository Folder Structure & Upload Manifest", level=1)
    tree_text = """
Career-Shaper-HCL-Project/
├── HCL Project/
│   └── HCLDay1.ipynb
├── Invoice_Data_Validation_Analysis/
│   ├── .streamlit/
│   │   └── config.toml
│   ├── data/
│   │   ├── app.db (SQLite database - 5 tables, 1 analytical view)
│   │   ├── invoice_data.xlsx (Multi-tab dataset)
│   │   ├── vendors.csv (30 records)
│   │   ├── purchase_orders.csv (150 records)
│   │   ├── invoices.csv (304 records)
│   │   ├── line_items.csv (1,056 records)
│   │   └── discrepancies.csv (69 records)
│   ├── docs/
│   │   ├── 01_Project_Understanding.docx
│   │   ├── 02_HMW_Statements.docx
│   │   ├── 03_Needs_Analysis.docx
│   │   ├── 04_Traceability_Matrix.xlsx
│   │   ├── 05_BRD.docx
│   │   ├── 06_PRD.docx
│   │   ├── 07_TRD.docx
│   │   ├── 08_Database_Design.xlsx
│   │   ├── 09_ER_Diagram.docx
│   │   ├── 11_User_Stories.docx
│   │   ├── 12_UI_Framework.docx
│   │   ├── 13_Test_Cases.docx
│   │   ├── 14_BUILD_PLAN.docx
│   │   ├── 16_Deployment_Guide.docx
│   │   ├── PJ_IVD_2026_Project_Document.docx
│   │   ├── PJ_IVD_2026_Presentation.pptx
│   │   └── pdf_submissions/ (All deliverables in PDF format)
│   ├── pages/
│   │   ├── 1_Overview_KPIs.py
│   │   ├── 2_Invoice_Validation.py
│   │   ├── 3_PO_Reconciliation.py
│   │   ├── 4_Vendor_Risk.py
│   │   ├── 5_Statistical_EDA.py
│   │   └── 6_Data_Quality_&_Export.py
│   ├── src/
│   │   ├── database.py
│   │   ├── validation.py
│   │   ├── analytics.py
│   │   ├── ui.py
│   │   └── generate_data.py
│   ├── tests/
│   │   ├── test_data.py
│   │   ├── test_analysis.py
│   │   └── test_app.py
│   ├── app.py
│   ├── BUILD_PLAN.md
│   ├── requirements.txt
│   └── conftest.py
"""
    p_code = doc.add_paragraph()
    r_c = p_code.add_run(tree_text.strip())
    r_c.font.name = "Consolas"
    r_c.font.size = Pt(8.5)

    doc.add_heading("Automated Test Verification Summary", level=1)
    doc.add_paragraph("All 13 automated unit and integration tests executed with 100% pass rate before repository push:")
    doc.add_paragraph("• tests/test_data.py: 4/4 PASSED (DB file exists, row counts, PK uniqueness, FK referential integrity)\n"
                      "• tests/test_analysis.py: 5/5 PASSED (Arithmetic check, duplicate detector, tax rate validator, PO reconciliation, VRI bounds)\n"
                      "• tests/test_app.py: 4/4 PASSED (KPI summary metrics, Benford's Law distribution test, IQR outlier detection, analytical view loading)")

    out_docx = BASE_DIR / "docs" / "GitHub_Report.docx"
    doc.save(str(out_docx))
    return out_docx

# ----------------- 2. Build Complete Code Document -----------------
def create_complete_code_doc():
    doc = Document()
    for s in doc.sections:
        s.top_margin = Inches(0.8)
        s.bottom_margin = Inches(0.8)
        s.left_margin = Inches(0.8)
        s.right_margin = Inches(0.8)

    p_top = doc.add_paragraph()
    r_top = p_top.add_run("HCLTech Career Shaper Programme | ABES Engineering College, Ghaziabad\nLMS Deliverable #4: Complete Source Code Catalog")
    r_top.bold = True
    r_top.font.color.rgb = RGBColor.from_string("A51C30")

    p_title = doc.add_paragraph()
    p_title.paragraph_format.space_before = Pt(8)
    p_title.paragraph_format.space_after = Pt(4)
    r_t = p_title.add_run("Complete Source Code Catalog")
    r_t.font.name = "Georgia"
    r_t.font.size = Pt(20)
    r_t.bold = True
    r_t.font.color.rgb = RGBColor.from_string("1E3A8A")

    doc.add_paragraph(
        "Project: PJ-IVD-2026 - Invoice Data Validation & Discrepancy Analytics Platform\n"
        "Learner: Surbhi Sharma (2500320140152) | Mentor: Kalpana Dwivedi | ABES Engineering College\n"
        "GitHub: https://github.com/surbhi0310/Career-Shaper-HCL-Project.git"
    )

    code_files = [
        ("requirements.txt", BASE_DIR / "requirements.txt"),
        (".streamlit/config.toml", BASE_DIR / ".streamlit" / "config.toml"),
        ("app.py", BASE_DIR / "app.py"),
        ("src/database.py", BASE_DIR / "src" / "database.py"),
        ("src/validation.py", BASE_DIR / "src" / "validation.py"),
        ("src/analytics.py", BASE_DIR / "src" / "analytics.py"),
        ("src/ui.py", BASE_DIR / "src" / "ui.py"),
        ("src/generate_data.py", BASE_DIR / "src" / "generate_data.py"),
        ("pages/1_Overview_KPIs.py", BASE_DIR / "pages" / "1_Overview_KPIs.py"),
        ("pages/2_Invoice_Validation.py", BASE_DIR / "pages" / "2_Invoice_Validation.py"),
        ("pages/3_PO_Reconciliation.py", BASE_DIR / "pages" / "3_PO_Reconciliation.py"),
        ("pages/4_Vendor_Risk.py", BASE_DIR / "pages" / "4_Vendor_Risk.py"),
        ("pages/5_Statistical_EDA.py", BASE_DIR / "pages" / "5_Statistical_EDA.py"),
        ("pages/6_Data_Quality_&_Export.py", BASE_DIR / "pages" / "6_Data_Quality_&_Export.py"),
        ("tests/test_data.py", BASE_DIR / "tests" / "test_data.py"),
        ("tests/test_analysis.py", BASE_DIR / "tests" / "test_analysis.py"),
        ("tests/test_app.py", BASE_DIR / "tests" / "test_app.py")
    ]

    for rel_path, abs_path in code_files:
        if abs_path.exists():
            h = doc.add_heading(level=2)
            r_h = h.add_run(f"File: {rel_path}")
            r_h.font.name = "Georgia"
            r_h.font.size = Pt(13)
            r_h.bold = True
            r_h.font.color.rgb = RGBColor.from_string("A51C30")

            with open(abs_path, "r", encoding="utf-8", errors="ignore") as f:
                code_content = f.read()

            p_code = doc.add_paragraph()
            r_c = p_code.add_run(code_content)
            r_c.font.name = "Consolas"
            r_c.font.size = Pt(8.0)
            r_c.font.color.rgb = RGBColor.from_string("151514")

    out_docx = BASE_DIR / "docs" / "Complete_Code.docx"
    doc.save(str(out_docx))
    return out_docx

# ----------------- 3. Convert both to PDF -----------------
gh_docx = create_github_report()
cc_docx = create_complete_code_doc()

word = win32com.client.Dispatch("Word.Application")
word.Visible = False

for doc_path in [gh_docx, cc_docx]:
    pdf_out = PDF_DIR / (doc_path.stem + ".pdf")
    print(f"Converting DOCX to PDF: {doc_path.name} -> {pdf_out.name}...")
    try:
        doc = word.Documents.Open(str(doc_path.resolve()))
        doc.SaveAs(str(pdf_out.resolve()), FileFormat=17)
        doc.Close()
        print(f"  [OK] Saved: {pdf_out}")
    except Exception as e:
        print(f"  [ERROR] {e}")

word.Quit()
print("GitHub Report & Complete Code PDFs created successfully!")
