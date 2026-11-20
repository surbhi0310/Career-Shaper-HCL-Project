# BUILD_PLAN.md: Invoice Data Validation & Discrepancy Analytics Platform

**Project ID:** `PJ-IVD-2026`  
**Candidate:** Surbhi Sharma | **Roll No:** 2500320140152 | **MCA 2nd Year DA**, ABES Engineering College  
**Mentor:** Kalpana Dwivedi  
**Stack:** Python 3.12, pandas, NumPy, SciPy, Plotly, Streamlit, SQLite, pytest  

---

## 1. Project Summary & Deliverables
The **Invoice Data Validation & Discrepancy Analytics Platform** automates anomaly detection, tax validation, duplicate tracking, PO reconciliation, and vendor risk profiling for enterprise Accounts Payable (AP) departments. It eliminates manual sampling vulnerabilities by testing 100% of invoice transactions.

### Key Deliverable Links:
- **Project Charter:** [docs/01_Project_Understanding.docx](docs/01_Project_Understanding.docx)
- **HMW Statements:** [docs/02_HMW_Statements.docx](docs/02_HMW_Statements.docx)
- **Needs Analysis:** [docs/03_Needs_Analysis.docx](docs/03_Needs_Analysis.docx)
- **Traceability Matrix:** [docs/04_Traceability_Matrix.xlsx](docs/04_Traceability_Matrix.xlsx)
- **BRD:** [docs/05_BRD.docx](docs/05_BRD.docx)
- **TRD:** [docs/07_TRD.docx](docs/07_TRD.docx)
- **Database Design:** [docs/08_Database_Design.xlsx](docs/08_Database_Design.xlsx)
- **ER Diagram:** [docs/09_ER_Diagram.docx](docs/09_ER_Diagram.docx)
- **User Stories:** [docs/11_User_Stories.docx](docs/11_User_Stories.docx)
- **UI Framework:** [docs/12_UI_Framework.docx](docs/12_UI_Framework.docx)
- **Test Cases:** [docs/13_Test_Cases.docx](docs/13_Test_Cases.docx)
- **Build Plan Document:** [docs/14_BUILD_PLAN.docx](docs/14_BUILD_PLAN.docx)
- **Final Project Document:** [docs/PJ_IVD_2026_Project_Document.docx](docs/PJ_IVD_2026_Project_Document.docx)
- **Presentation Deck:** [docs/PJ_IVD_2026_Presentation.pptx](docs/PJ_IVD_2026_Presentation.pptx)

---

## 2. Final Folder Architecture
```text
AI-website-Builder-final-project-main/
│
├── .streamlit/
│   └── config.toml                  # UI theme colors & port config
├── data/
│   ├── app.db                       # SQLite database (ACID compliant)
│   ├── invoice_data.xlsx            # Multi-sheet Excel dataset
│   ├── vendors.csv                  # 30 Master vendors
│   ├── purchase_orders.csv          # 150 Purchase Orders
│   ├── invoices.csv                 # 304 Invoices
│   ├── line_items.csv               # 1,056 Line Items
│   └── discrepancies.csv            # 69 Injected Discrepancies
├── docs/                            # Downloadable Step deliverables
├── pages/
│   ├── 1_Overview_KPIs.py           # Executive dashboard & trends
│   ├── 2_Invoice_Validation.py      # Detailed invoice line-item auditor
│   ├── 3_PO_Reconciliation.py       # 2-Way PO budget matching
│   ├── 4_Vendor_Risk.py             # Vendor Risk Index (VRI)
│   ├── 5_Statistical_EDA.py         # Benford's Law & IQR Outliers
│   └── 6_Data_Quality_&_Export.py   # Raw table viewer & downloads
├── src/
│   ├── __init__.py
│   ├── database.py                  # SQLite connection & cached queries
│   ├── validation.py                # Mathematical & duplicate rules engine
│   ├── analytics.py                 # Statistical algorithms & Benford test
│   ├── ui.py                        # Tailwind CSS injection & KPI cards
│   └── generate_data.py             # Seed-42 synthetic dataset generator
├── tests/
│   ├── __init__.py
│   ├── test_data.py                 # DB schema, PK/FK integrity tests
│   ├── test_analysis.py             # Business rule & validation logic tests
│   └── test_app.py                  # Integration & stats tests
├── app.py                           # Multi-page Streamlit portal
├── conftest.py                      # Pytest sys.path bootstrap
├── requirements.txt                 # Pinned dependencies
└── README.md                        # Project documentation & run guide
```

---

## 3. Build Tasks Execution Matrix
- **`BUILD-01`**: Setup environment, folder architecture, and requirements.txt.
- **`BUILD-02`**: Implement `src/generate_data.py` (seed 42, 5 injected anomaly patterns).
- **`BUILD-03`**: Build database layer `src/database.py` and analytical view `v_invoice_analysis`.
- **`BUILD-04`**: Develop validation engine `src/validation.py` (arithmetic, duplicates, tax, PO matching, VRI).
- **`BUILD-05`**: Implement statistical EDA suite `src/analytics.py` (Benford's Law, IQR, Z-scores).
- **`BUILD-06`**: Design Tailwind UI components `src/ui.py` and theme `.streamlit/config.toml`.
- **`BUILD-07`**: Build main landing portal `app.py` and 6 interactive multi-page dashboards.
- **`BUILD-08`**: Author and pass automated test suite `tests/` (`13/13` passed).
- **`BUILD-09`**: Compile all Word documents, Excel workbooks, and PowerPoint presentations.
- **`BUILD-10`**: Finalize Streamlit Community Cloud deployment readiness.

---

## 4. Definition of Done (DoD)
1. All 12 Business Requirements (`BR-01` to `BR-12`) implemented and traced.
2. 100% of injected synthetic anomalies detected by validation rules.
3. 13 unit and integration tests passing cleanly via `pytest`.
4. All 18 step deliverables generated and placed in `docs/`.
5. ABES Project Document and Presentation generated strictly conforming to template specs.

---

## 5. Execution Commands
```bash
# 1. Install dependencies
pip install -r requirements.txt

# 2. Run unit & integration test suite
pytest -v

# 3. Launch interactive Streamlit application
streamlit run app.py
```

---

## 6. Viva Preparation: 10 Critical Technical Questions
1. **Q: Why use SQLite instead of MySQL or PostgreSQL for this application?**  
   *A:* SQLite is a zero-configuration, serverless, self-contained ACID database that embeds directly into the repository. It deploys seamlessly on Streamlit Community Cloud without requiring external containerized database hosts, while providing indexed sub-second query performance for our dataset size.
2. **Q: How does the duplicate detection algorithm handle invoice numbering variations?**  
   *A:* In addition to exact invoice number matching, the algorithm groups transactions by Vendor ID and Total Amount within a rolling 14-day window. This flags duplicate submissions even if the vendor altered the invoice number suffix or prefix.
3. **Q: What is the mathematical significance of Benford's Law in forensic auditing?**  
   *A:* Benford's Law states that the probability of leading digit $d$ is $P(d) = \log_{10}(1 + 1/d)$. Natural procurement amounts conform to this logarithmic decline ($30.1\%$ for 1 vs $4.6\%$ for 9). Deviations detected via Chi-Square goodness-of-fit test expose invented numbers or split invoices designed to bypass threshold approvals.
4. **Q: How is the Vendor Risk Index (VRI) mathematically computed?**  
   *A:* $\text{VRI} = (\text{Error Rate } \% \times 0.45) + ((\text{Weighted Discrepancy Points} / \text{Max Points}) \times 45) + ((1 - \text{Reliability}) \times 10)$, scaled from 0 to 100 and mapped to High, Medium, and Low risk tiers.
5. **Q: How do you verify line-item arithmetic integrity?**  
   *A:* For each line item, the engine calculates $\text{Expected Total} = (\text{Quantity} \times \text{Unit Price}) + (\text{Quantity} \times \text{Unit Price} \times \text{Tax Rate})$. If $|\text{Line Total} - \text{Expected Total}| > 0.01$, an arithmetic discrepancy is flagged.
6. **Q: What is the purpose of the analytical view `v_invoice_analysis`?**  
   *A:* It pre-joins invoices, vendors, purchase orders, and discrepancy logs into a denormalized 32-column view. This simplifies Streamlit queries, eliminates repetitive multi-table joins, and provides sub-second render times.
7. **Q: How does the application handle missing PO references?**  
   *A:* Invoices with null PO references are classified as "Missing PO Reference" discrepancies with Medium severity and their payment status is automatically marked as "Blocked".
8. **Q: What is the advantage of using Tailwind styling inside Streamlit?**  
   *A:* Standard Streamlit components have rigid default aesthetics. Injecting scoped Tailwind CSS classes allows enterprise-grade card design, custom badge styling, hover effects, and responsive visual hierarchy.
9. **Q: How does the Requirements Traceability Matrix (RTM) guarantee project quality?**  
   *A:* RTM maps each Opportunity Question (`OQ`) to a How Might We (`HMW`), Business Requirement (`BR`), User Story (`US`), UI Screen (`UI`), and Test Case (`TC`). It verifies 100% test coverage and zero orphaned requirements.
10. **Q: How is 2-Way PO matching implemented?**  
    *A:* Invoices are matched with their corresponding PO record by `po_id`. If `(Total Invoiced Amount - PO Budget) / PO Budget > Tolerance %` (default 2%), the invoice is flagged for budget overrun and escalated to management.
