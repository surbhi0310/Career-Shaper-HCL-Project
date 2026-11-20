"""
convert_all_to_pdf.py - Automated PDF Generation via Word & PowerPoint COM
Generates exact LMS deliverables matching Day 04 Tutorial Plan instructions
"""

import os
import sys
import win32com.client
from pathlib import Path

BASE_DIR = Path(r"C:\Users\Geetanjali\Downloads\AI-website-Builder-final-project-main\Invoice_Data_Validation_Analysis")
DOCS_DIR = BASE_DIR / "docs"
PDF_DIR = BASE_DIR / "docs" / "pdf_submissions"
PDF_DIR.mkdir(parents=True, exist_ok=True)

# 1. Convert Word documents to PDF
word_files = list(DOCS_DIR.glob("*.docx"))
print(f"Found {len(word_files)} Word files to convert.")

word = win32com.client.Dispatch("Word.Application")
word.Visible = False

for doc_path in word_files:
    pdf_name = doc_path.stem + ".pdf"
    pdf_out = PDF_DIR / pdf_name
    print(f"Converting DOCX: {doc_path.name} -> {pdf_name}...")
    try:
        doc = word.Documents.Open(str(doc_path.resolve()))
        # wdFormatPDF = 17
        doc.SaveAs(str(pdf_out.resolve()), FileFormat=17)
        doc.Close()
        print(f"  [OK] Saved: {pdf_out}")
    except Exception as e:
        print(f"  [ERROR] {e}")

word.Quit()

# 2. Convert PowerPoint presentation to PDF
ppt_files = list(DOCS_DIR.glob("*.pptx"))
print(f"\nFound {len(ppt_files)} PPT files to convert.")

ppt_app = win32com.client.Dispatch("PowerPoint.Application")
# ppt_app.Visible = True

for ppt_path in ppt_files:
    pdf_name = ppt_path.stem + ".pdf"
    pdf_out = PDF_DIR / pdf_name
    print(f"Converting PPTX: {ppt_path.name} -> {pdf_name}...")
    try:
        pres = ppt_app.Presentations.Open(str(ppt_path.resolve()), WithWindow=False)
        # ppSaveAsPDF = 32
        pres.SaveAs(str(pdf_out.resolve()), 32)
        pres.Close()
        print(f"  [OK] Saved: {pdf_out}")
    except Exception as e:
        print(f"  [ERROR] {e}")

ppt_app.Quit()

print("\nAll DOCX and PPTX conversions to PDF complete!")
