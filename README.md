# HCLTech Career Shaper Programme - Project Repository

**Institution:** ABES Engineering College, Ghaziabad  
**Student:** Surbhi Sharma | **Roll Number:** 2500320140152  
**Specialization:** MCA 2nd Year (Data Analytics)  
**Project Mentor:** Dinesh Sereno  

---

## 📂 Repository Contents

### 1. [`Invoice_Data_Validation_Analysis/`](Invoice_Data_Validation_Analysis/)
The primary end-to-end Data Analytics project: **Invoice Data Validation & Discrepancy Analytics Platform** (`PJ-IVD-2026`).

- **Executive Multi-Page App:** Streamlit application (`app.py` & `pages/1-6`)
- **Core Engine:** Rules engine & statistical analytics (`src/`)
- **Database & Data:** SQLite database (`data/app.db`) and Excel sheets (`data/invoice_data.xlsx`)
- **Automated Tests:** 13/13 passing `pytest` test suite (`tests/`)
- **Academic Documentation:** Complete 18-step deliverables (`docs/`), ABES Project Document, Presentation (`.pptx` & `.pdf`), and LMS submission files (`docs/pdf_submissions/`).

---

### Quick Start
```bash
cd Invoice_Data_Validation_Analysis
pip install -r requirements.txt
pytest -v
streamlit run app.py
```
