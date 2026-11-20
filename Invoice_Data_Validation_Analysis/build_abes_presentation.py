"""
build_abes_presentation.py - Generates the ABES Project Presentation (docs/PJ_IVD_2026_Presentation.pptx)
Strictly conforming to ABES 11-Slide Deck Specifications
Project: PJ-IVD-2026 | Surbhi Sharma (2500320140152) | Mentor: Kalpana Dwivedi
"""

import os
import pptx
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.dml.color import RGBColor

DOCS_DIR = "docs"
os.makedirs(DOCS_DIR, exist_ok=True)
PPT_PATH = os.path.join(DOCS_DIR, "PJ_IVD_2026_Presentation.pptx")

# Color Palette
RED = RGBColor(0xA5, 0x1C, 0x30)
BLACK = RGBColor(0x15, 0x15, 0x14)
DARK = RGBColor(0x3A, 0x3A, 0x38)
GREY = RGBColor(0x5E, 0x5E, 0x5B)
CARD_BG = RGBColor(0xF4, 0xF1, 0xF1)
PINK = RGBColor(0xD9, 0xB5, 0xBC)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
BORDER_COLOR = RGBColor(0xD9, 0xD9, 0xD9)

def add_header_footer_content_slide(slide, title_text, slide_num):
    # Title (.5, .3, 9, .75) Cambria 28 bold black, centred, middle
    txBox = slide.shapes.add_textbox(Inches(0.5), Inches(0.3), Inches(9.0), Inches(0.75))
    tf = txBox.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    p = tf.paragraphs[0]
    p.text = title_text
    p.alignment = PP_ALIGN.CENTER
    p.font.name = "Cambria"
    p.font.size = Pt(28)
    p.font.bold = True
    p.font.color.rgb = BLACK

    # Footer (.5, 5.2, 7.5, .3) 9 grey, middle
    ftBox = slide.shapes.add_textbox(Inches(0.5), Inches(5.2), Inches(7.5), Inches(0.3))
    tf_ft = ftBox.text_frame
    tf_ft.word_wrap = True
    tf_ft.vertical_anchor = MSO_ANCHOR.MIDDLE
    p_ft = tf_ft.paragraphs[0]
    p_ft.text = "HCLTech Career Shaper ABES Engineering College | PJ-IVD-2026 - Invoice Validation Platform"
    p_ft.font.name = "Calibri"
    p_ft.font.size = Pt(9)
    p_ft.font.color.rgb = GREY

    # Slide number (9, 5.2, .5, .3) 9, color 8A8A86, right
    numBox = slide.shapes.add_textbox(Inches(9.0), Inches(5.2), Inches(0.5), Inches(0.3))
    tf_num = numBox.text_frame
    tf_num.vertical_anchor = MSO_ANCHOR.MIDDLE
    p_num = tf_num.paragraphs[0]
    p_num.text = str(slide_num)
    p_num.alignment = PP_ALIGN.RIGHT
    p_num.font.name = "Calibri"
    p_num.font.size = Pt(9)
    p_num.font.color.rgb = RGBColor(0x8A, 0x8A, 0x86)

def add_badge(slide, x, y, text):
    # BADGE (x, y, t): circle .42 red; t centred, 12 bold white
    shape = slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(x), Inches(y), Inches(0.42), Inches(0.42))
    shape.fill.solid()
    shape.fill.fore_color.rgb = RED
    shape.line.fill.background()
    tf = shape.text_frame
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    p = tf.paragraphs[0]
    p.text = text
    p.alignment = PP_ALIGN.CENTER
    p.font.name = "Calibri"
    p.font.size = Pt(12)
    p.font.bold = True
    p.font.color.rgb = WHITE

def add_card(slide, x, y, w, h):
    shape = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(x), Inches(y), Inches(w), Inches(h))
    shape.fill.solid()
    shape.fill.fore_color.rgb = CARD_BG
    shape.line.color.rgb = CARD_BG
    return shape

def add_badge_card(slide, x, y, w, h, badge_char, head_text, body_text):
    add_card(slide, x, y, w, h)
    add_badge(slide, x + 0.2, y + 0.2, badge_char)
    
    # HEAD (x+.75, y+.17, w-.9, .45) 15 bold black, middle
    h_box = slide.shapes.add_textbox(Inches(x + 0.75), Inches(y + 0.17), Inches(w - 0.9), Inches(0.45))
    h_tf = h_box.text_frame
    h_tf.word_wrap = True
    h_tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    p_h = h_tf.paragraphs[0]
    p_h.text = head_text
    p_h.font.name = "Calibri"
    p_h.font.size = Pt(14)
    p_h.font.bold = True
    p_h.font.color.rgb = BLACK

    # BODY (x+.2, y+.72, w-.4, h-.85) 13 dark, top
    b_box = slide.shapes.add_textbox(Inches(x + 0.2), Inches(y + 0.72), Inches(w - 0.4), Inches(h - 0.85))
    b_tf = b_box.text_frame
    b_tf.word_wrap = True
    b_tf.vertical_anchor = MSO_ANCHOR.TOP
    p_b = b_tf.paragraphs[0]
    p_b.text = body_text
    p_b.font.name = "Calibri"
    p_b.font.size = Pt(12)
    p_b.font.color.rgb = DARK

def add_plain_card(slide, x, y, w, h, head_text, body_text):
    add_card(slide, x, y, w, h)
    
    # HEAD (x+.2, y+.17, w-.35, .45)
    h_box = slide.shapes.add_textbox(Inches(x + 0.2), Inches(y + 0.17), Inches(w - 0.35), Inches(0.45))
    h_tf = h_box.text_frame
    h_tf.word_wrap = True
    h_tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    p_h = h_tf.paragraphs[0]
    p_h.text = head_text
    p_h.font.name = "Calibri"
    p_h.font.size = Pt(14)
    p_h.font.bold = True
    p_h.font.color.rgb = BLACK

    # BODY (x+.2, y+.72, w-.4, h-.85)
    b_box = slide.shapes.add_textbox(Inches(x + 0.2), Inches(y + 0.72), Inches(w - 0.4), Inches(h - 0.85))
    b_tf = b_box.text_frame
    b_tf.word_wrap = True
    b_tf.vertical_anchor = MSO_ANCHOR.TOP
    p_b = b_tf.paragraphs[0]
    p_b.text = body_text
    p_b.font.name = "Calibri"
    p_b.font.size = Pt(12)
    p_b.font.color.rgb = DARK

def set_speaker_note(slide, note_text):
    notes_slide = slide.notes_slide
    text_frame = notes_slide.notes_text_frame
    text_frame.text = note_text

def build_presentation():
    prs = Presentation()
    prs.slide_width = Inches(10.0)
    prs.slide_height = Inches(5.625)
    blank_layout = prs.slide_layouts[6]

    # ==================== SLIDE 1: DARK (Title) ====================
    s1 = prs.slides.add_slide(blank_layout)
    bg1 = s1.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, Inches(10), Inches(5.625))
    bg1.fill.solid()
    bg1.fill.fore_color.rgb = BLACK
    bg1.line.fill.background()

    # [ABES logo] (.5, .35, 1.6, .5) italic 10 pink, middle
    b_logo = s1.shapes.add_textbox(Inches(0.5), Inches(0.35), Inches(1.6), Inches(0.5))
    tf_l = b_logo.text_frame
    tf_l.vertical_anchor = MSO_ANCHOR.MIDDLE
    p_l = tf_l.paragraphs[0]
    p_l.text = "[ABES logo]"
    p_l.font.name = "Calibri"
    p_l.font.size = Pt(10)
    p_l.font.italic = True
    p_l.font.color.rgb = PINK

    # Program header (2.2, .35, 7.3, .5) 11 pink, right, middle
    b_prog = s1.shapes.add_textbox(Inches(2.2), Inches(0.35), Inches(7.3), Inches(0.5))
    tf_p = b_prog.text_frame
    tf_p.vertical_anchor = MSO_ANCHOR.MIDDLE
    p_p = tf_p.paragraphs[0]
    p_p.text = "HCLTech Career Shaper Programme | ABES Engineering College, Ghaziabad"
    p_p.alignment = PP_ALIGN.RIGHT
    p_p.font.name = "Calibri"
    p_p.font.size = Pt(11)
    p_p.font.color.rgb = PINK

    # Title (.5, 1.55, 9, 1.5) Cambria 36 bold white, left, bottom
    b_title = s1.shapes.add_textbox(Inches(0.5), Inches(1.55), Inches(9.0), Inches(1.5))
    tf_t = b_title.text_frame
    tf_t.word_wrap = True
    tf_t.vertical_anchor = MSO_ANCHOR.BOTTOM
    p_t = tf_t.paragraphs[0]
    p_t.text = "PJ-IVD-2026 - Invoice Data Validation & Discrepancy Analytics Platform"
    p_t.font.name = "Cambria"
    p_t.font.size = Pt(32)
    p_t.font.bold = True
    p_t.font.color.rgb = WHITE

    # Lines (.5, 3.2, 9, 1.7) 15 F4F1F1, left, top
    b_lines = s1.shapes.add_textbox(Inches(0.5), Inches(3.2), Inches(9.0), Inches(1.7))
    tf_ln = b_lines.text_frame
    tf_ln.word_wrap = True
    tf_ln.vertical_anchor = MSO_ANCHOR.TOP
    p_ln1 = tf_ln.paragraphs[0]
    p_ln1.text = "Surbhi Sharma | Roll No 2500320140152 / MCA 2nd Year Data Analyst"
    p_ln1.font.name = "Calibri"
    p_ln1.font.size = Pt(15)
    p_ln1.font.color.rgb = CARD_BG
    
    p_ln2 = tf_ln.add_paragraph()
    p_ln2.space_before = Pt(6)
    p_ln2.text = "Mentor: Kalpana Dwivedi / 20 November 2026"
    p_ln2.font.name = "Calibri"
    p_ln2.font.size = Pt(14)
    p_ln2.font.color.rgb = PINK

    set_speaker_note(s1, "Good morning respected evaluators and mentor. I am Surbhi Sharma, presenting project PJ-IVD-2026, an enterprise Invoice Data Validation and Discrepancy Analytics platform designed to eliminate Accounts Payable financial leakage.")

    # ==================== SLIDE 2: Problem and How Might We ====================
    s2 = prs.slides.add_slide(blank_layout)
    add_header_footer_content_slide(s2, "Problem and How Might We", 2)
    
    # BADGE CARD (.5, 1.25, 4.3, 3.7) P, Problem statement
    prob_text = ("Manual sample audits inspect less than 15% of vendor invoices, allowing subtle arithmetic errors, duplicate submissions, and PO price creep to escape.\n\n"
                 "This undetected capital leakage bleeds 0.5%–2% of total corporate procurement spend and causes severe GST tax penalty risks.")
    add_badge_card(s2, 0.5, 1.25, 4.3, 3.7, "P", "Problem Statement", prob_text)
    
    # BADGE CARD (5.2, 1.25, 4.3, 1.75) 1, HMW-01
    hmw1_text = "How might we automatically verify mathematical line-item calculations for AP Specialists so that billing errors are caught before ledger posting?"
    add_badge_card(s2, 5.2, 1.25, 4.3, 1.75, "1", "HMW-01: Arithmetic Verification", hmw1_text)

    # BADGE CARD (5.2, 3.2, 4.3, 1.75) 2, HMW-02
    hmw2_text = "How might we detect near-duplicate invoice submissions across rolling date windows for the CFO so that double disbursements are prevented?"
    add_badge_card(s2, 5.2, 3.2, 4.3, 1.75, "2", "HMW-02: Duplicate Detection", hmw2_text)

    set_speaker_note(s2, "Our problem stems from manual sampling in Accounts Payable. We framed this into key How Might We statements: automating 100% calculation checks and detecting rolling duplicate billing collisions.")

    # ==================== SLIDE 3: Stakeholders and users ====================
    s3 = prs.slides.add_slide(blank_layout)
    add_header_footer_content_slide(s3, "Stakeholders and users", 3)

    sh_cards = [
        (0.5, "1", "AP Specialist (SH-05)", "Primary User: Needs sub-second line-item inspection, discrepancy flags, and one-click dispute export.\n\nInfluence: Medium"),
        (3.55, "2", "AP Manager & CFO (SH-01/03)", "Secondary User: Needs executive KPI tracking, duplicate prevention, and operational backlog reduction.\n\nInfluence: High"),
        (6.6, "3", "Compliance & Mentor (SH-02)", "Evaluator / Mentor: Demands 100% population audit, statutory GST validation, and an immutable audit log.\n\nInfluence: High")
    ]
    for x_c, b_c, h_c, body_c in sh_cards:
        add_badge_card(s3, x_c, 1.35, 2.85, 3.5, b_c, h_c, body_c)

    set_speaker_note(s3, "We mapped three core user groups: frontline AP Specialists who process bills, executive CFOs focused on capital leakage, and audit evaluators demanding full compliance traceability.")

    # ==================== SLIDE 4: Requirements and traceability ====================
    s4 = prs.slides.add_slide(blank_layout)
    add_header_footer_content_slide(s4, "Requirements and traceability", 4)

    # TABLE at (.5, 1.3), column widths 1, 3.4, 1, 1.2, 1.2, 1.2
    t_shape = s4.shapes.add_table(5, 6, Inches(0.5), Inches(1.3), Inches(9.0), Inches(2.5))
    tbl4 = t_shape.table
    c_w = [Inches(1.0), Inches(3.4), Inches(1.0), Inches(1.2), Inches(1.2), Inches(1.2)]
    for idx, w in enumerate(c_w):
        tbl4.columns[idx].width = w

    t_headers = ["HMW", "Requirement", "Priority", "User story", "Test", "Screen"]
    for idx, h in enumerate(t_headers):
        cell = tbl4.cell(0, idx)
        cell.fill.solid()
        cell.fill.fore_color.rgb = RED
        cell.text_frame.text = h
        p = cell.text_frame.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER
        p.font.name = "Calibri"
        p.font.size = Pt(11)
        p.font.bold = True
        p.font.color.rgb = WHITE

    t_rows = [
        ("HMW-01", "Line-item arithmetic verification", "Must", "US-01", "TC-F-01", "UI-02"),
        ("HMW-02", "Duplicate invoice detection", "Must", "US-03", "TC-F-02", "UI-02"),
        ("HMW-03", "Tax slab statutory validation", "Must", "US-05", "TC-F-03", "UI-02"),
        ("HMW-04", "PO-to-invoice budget reconciliation", "Must", "US-06", "TC-F-04", "UI-03")
    ]
    for r_idx, r_data in enumerate(t_rows, start=1):
        for c_idx, val in enumerate(r_data):
            cell = tbl4.cell(r_idx, c_idx)
            cell.text_frame.text = val
            p = cell.text_frame.paragraphs[0]
            p.alignment = PP_ALIGN.CENTER if c_idx != 1 else PP_ALIGN.LEFT
            p.font.name = "Calibri"
            p.font.size = Pt(10)
            p.font.color.rgb = BLACK

    # Reference note (.5, 4.3, 9, .35)
    r_box = s4.shapes.add_textbox(Inches(0.5), Inches(4.3), Inches(9.0), Inches(0.35))
    p_r = r_box.text_frame.paragraphs[0]
    p_r.text = "Full matrix: Project Document, section 4"
    p_r.font.name = "Calibri"
    p_r.font.size = Pt(12)
    p_r.font.italic = True
    p_r.font.color.rgb = GREY

    set_speaker_note(s4, "Every requirement strictly traces back to a How Might We statement, and forward to user stories, test cases, and UI screens. Zero orphaned requirements exist.")

    # ==================== SLIDE 5: Data and preparation ====================
    s5 = prs.slides.add_slide(blank_layout)
    add_header_footer_content_slide(s5, "Data and preparation", 5)

    # Four stats x .5 / 2.8 / 5.1 / 7.4, w 2.1
    stats = [
        (0.5, "1,056", "Line-Item Rows"),
        (2.8, "32", "Analytical Columns"),
        (5.1, "12 Mo", "Time Horizon (2026)"),
        (7.4, "69", "Anomalies Injected/Fixed")
    ]
    for x_s, val_s, lbl_s in stats:
        # value y 1.3 h .8 Cambria 30 bold red, middle
        vb = s5.shapes.add_textbox(Inches(x_s), Inches(1.3), Inches(2.1), Inches(0.8))
        tf_v = vb.text_frame
        tf_v.vertical_anchor = MSO_ANCHOR.MIDDLE
        p_v = tf_v.paragraphs[0]
        p_v.text = val_s
        p_v.alignment = PP_ALIGN.CENTER
        p_v.font.name = "Cambria"
        p_v.font.size = Pt(30)
        p_v.font.bold = True
        p_v.font.color.rgb = RED

        # label y 2.1 h .35 12 grey, middle
        lb = s5.shapes.add_textbox(Inches(x_s), Inches(2.1), Inches(2.1), Inches(0.35))
        tf_l = lb.text_frame
        tf_l.vertical_anchor = MSO_ANCHOR.MIDDLE
        p_l = tf_l.paragraphs[0]
        p_l.text = lbl_s
        p_l.alignment = PP_ALIGN.CENTER
        p_l.font.name = "Calibri"
        p_l.font.size = Pt(12)
        p_l.font.color.rgb = GREY

    # PLAIN CARD (.5, 2.7, 4.3, 2.25) Source
    src_body = ("Source: Synthetic B2B procurement dataset generated via seed-42 script (src/generate_data.py).\n\n"
                "Grain: One row per invoice line item with foreign key links to vendor master, PO allocations, and invoices.")
    add_plain_card(s5, 0.5, 2.7, 4.3, 2.25, "Dataset Source & Grain", src_body)

    # PLAIN CARD (5.2, 2.7, 4.3, 2.25) Cleaning
    cln_body = ("1. Missing PO Mappings: 8 unmapped bills isolated & routed to review.\n"
                "2. Duplicate Bills: 4 duplicate billing collisions trapped.\n"
                "3. Extreme Price Outliers: 3 pricing spikes tagged via IQR.")
    add_plain_card(s5, 5.2, 2.7, 4.3, 2.25, "Cleaning & Data Hygiene", cln_body)

    set_speaker_note(s5, "Our dataset represents over 1,000 line items across 30 vendors. We modeled realistic data quality challenges including missing PO references and duplicate submissions.")

    # ==================== SLIDE 6: Solution and technology ====================
    s6 = prs.slides.add_slide(blank_layout)
    add_header_footer_content_slide(s6, "Solution and technology", 6)

    # Five steps x .5 / 2.35 / 4.2 / 6.05 / 7.9, y 1.5, w 1.6, h 1.2
    steps = [
        (0.5, "Raw data\n(CSV / xlsx)", False),
        (2.35, "Clean & KPIs\n(Python, pandas)", False),
        (4.2, "Store\n(SQLite data/app.db)", False),
        (6.05, "Dashboard\n(Streamlit)", True),
        (7.9, "Deploy\n(GitHub -> Cloud)", False)
    ]
    for x_st, text_st, is_dash in steps:
        st_shape = s6.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(x_st), Inches(1.5), Inches(1.6), Inches(1.2))
        st_shape.fill.solid()
        st_shape.fill.fore_color.rgb = RED if is_dash else CARD_BG
        st_shape.line.fill.background()
        tf = st_shape.text_frame
        tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        p = tf.paragraphs[0]
        p.text = text_st
        p.alignment = PP_ALIGN.CENTER
        p.font.name = "Calibri"
        p.font.size = Pt(11)
        p.font.bold = True
        p.font.color.rgb = WHITE if is_dash else BLACK

    # Arrows between steps
    arrows = [2.1, 3.95, 5.8, 7.65]
    for x_ar in arrows:
        ar_box = s6.shapes.add_textbox(Inches(x_ar), Inches(1.5), Inches(0.25), Inches(1.2))
        tf_ar = ar_box.text_frame
        tf_ar.vertical_anchor = MSO_ANCHOR.MIDDLE
        p_ar = tf_ar.paragraphs[0]
        p_ar.text = ">"
        p_ar.alignment = PP_ALIGN.CENTER
        p_ar.font.name = "Calibri"
        p_ar.font.size = Pt(18)
        p_ar.font.bold = True
        p_ar.font.color.rgb = RED

    # Tools line (.5, 3.15, 9, .5) 13 dark, middle
    t_box = s6.shapes.add_textbox(Inches(0.5), Inches(3.15), Inches(9.0), Inches(0.5))
    tf_t = t_box.text_frame
    tf_t.vertical_anchor = MSO_ANCHOR.MIDDLE
    p_t = tf_t.paragraphs[0]
    p_t.text = "Tools: Python 3.12 · pandas · NumPy · SciPy · Plotly · Streamlit · SQLite · pytest"
    p_t.alignment = PP_ALIGN.CENTER
    p_t.font.name = "Calibri"
    p_t.font.size = Pt(13)
    p_t.font.bold = True
    p_t.font.color.rgb = DARK

    # Local deploy line (.5, 3.75, 9, .5) 13 italic grey, middle
    d_box = s6.shapes.add_textbox(Inches(0.5), Inches(3.75), Inches(9.0), Inches(0.5))
    tf_d = d_box.text_frame
    tf_d.vertical_anchor = MSO_ANCHOR.MIDDLE
    p_d = tf_d.paragraphs[0]
    p_d.text = "Run locally: streamlit run app.py -> push to GitHub -> deploy on Streamlit Community Cloud"
    p_d.alignment = PP_ALIGN.CENTER
    p_d.font.name = "Calibri"
    p_d.font.size = Pt(13)
    p_d.font.italic = True
    p_d.font.color.rgb = GREY

    set_speaker_note(s6, "Our architecture ingests raw procurement data into an indexed SQLite store. Streamlit powers the interactive multi-page frontend, styled with Tailwind cards.")

    # ==================== SLIDE 7: Dashboard ====================
    s7 = prs.slides.add_slide(blank_layout)
    add_header_footer_content_slide(s7, "Dashboard", 7)

    # Frame (.5, 1.25, 5.5, 3.7) fill F4F1F1, dashed 1 pt BFBFBF
    frame = s7.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.5), Inches(1.25), Inches(5.5), Inches(3.7))
    frame.fill.solid()
    frame.fill.fore_color.rgb = CARD_BG
    frame.line.color.rgb = RGBColor(0xBF, 0xBF, 0xBF)
    tf_fr = frame.text_frame
    tf_fr.vertical_anchor = MSO_ANCHOR.MIDDLE
    p_fr = tf_fr.paragraphs[0]
    p_fr.text = "[Paste a dashboard screenshot here]"
    p_fr.alignment = PP_ALIGN.CENTER
    p_fr.font.name = "Calibri"
    p_fr.font.size = Pt(14)
    p_fr.font.italic = True
    p_fr.font.color.rgb = GREY

    # Badges & Screen purposes
    screen_list = [
        (1.3, "1", "UI-01", "Executive Spend & Error Rate KPIs"),
        (2.2, "2", "UI-02", "Automated Line-Item Audit Inspector"),
        (3.1, "3", "UI-03", "PO-to-Invoice Budget Reconciliation"),
        (4.0, "4", "UI-04", "Vendor Risk Index (VRI) Scorecards")
    ]
    for y_pos, b_num, ui_id, ui_desc in screen_list:
        add_badge(s7, 6.7, y_pos, b_num)
        desc_box = s7.shapes.add_textbox(Inches(7.25), Inches(y_pos), Inches(2.25), Inches(0.42))
        tf_desc = desc_box.text_frame
        tf_desc.vertical_anchor = MSO_ANCHOR.MIDDLE
        p_desc = tf_desc.paragraphs[0]
        r_id = p_desc.add_run()
        r_id.text = f"{ui_id}: "
        r_id.font.bold = True
        r_id.font.color.rgb = RED
        r_tx = p_desc.add_run()
        r_tx.text = ui_desc
        r_tx.font.color.rgb = BLACK
        p_desc.font.name = "Calibri"
        p_desc.font.size = Pt(11)

    set_speaker_note(s7, "Here is our dashboard layout. The top row showcases real-time KPIs, and the four tabs allow auditors to instantly filter discrepancies and download exception manifests.")

    # ==================== SLIDE 8: Key findings ====================
    s8 = prs.slides.add_slide(blank_layout)
    add_header_footer_content_slide(s8, "Key findings", 8)

    # Three CARDS y 1.3, w 2.85, h 3.6, x .5 / 3.55 / 6.6
    findings = [
        (0.5, "100%", "Calculation Error Detection", "Evidence: 69/69 synthetic arithmetic discrepancies flagged with zero false negatives."),
        (3.55, "4", "Duplicate Invoices Trapped", "Evidence: ₹3.8 Lakhs in duplicate payments halted via rolling 14-day fuzzy matching."),
        (6.6, "3", "High-Risk Rogue Vendors", "Evidence: VRI scored 3 vendors > 65 due to chronic tax and pricing variances.")
    ]
    for x_f, num_f, mean_f, evid_f in findings:
        add_card(s8, x_f, 1.3, 2.85, 3.6)
        
        # number y 1.5 h 1 Cambria 40 bold red, middle
        nb = s8.shapes.add_textbox(Inches(x_f + 0.2), Inches(1.5), Inches(2.45), Inches(1.0))
        tf_n = nb.text_frame
        tf_n.vertical_anchor = MSO_ANCHOR.MIDDLE
        p_n = tf_n.paragraphs[0]
        p_n.text = num_f
        p_n.alignment = PP_ALIGN.CENTER
        p_n.font.name = "Cambria"
        p_n.font.size = Pt(38)
        p_n.font.bold = True
        p_n.font.color.rgb = RED

        # meaning y 2.6 h .9 14 bold black, top
        mb = s8.shapes.add_textbox(Inches(x_f + 0.2), Inches(2.6), Inches(2.45), Inches(0.9))
        tf_m = mb.text_frame
        tf_m.word_wrap = True
        tf_m.vertical_anchor = MSO_ANCHOR.TOP
        p_m = tf_m.paragraphs[0]
        p_m.text = mean_f
        p_m.alignment = PP_ALIGN.CENTER
        p_m.font.name = "Calibri"
        p_m.font.size = Pt(13)
        p_m.font.bold = True
        p_m.font.color.rgb = BLACK

        # Evidence y 3.6 h 1.1 12 grey, top
        eb = s8.shapes.add_textbox(Inches(x_f + 0.2), Inches(3.6), Inches(2.45), Inches(1.1))
        tf_e = eb.text_frame
        tf_e.word_wrap = True
        tf_e.vertical_anchor = MSO_ANCHOR.TOP
        p_e = tf_e.paragraphs[0]
        p_e.text = evid_f
        p_e.font.name = "Calibri"
        p_e.font.size = Pt(11)
        p_e.font.color.rgb = GREY

    set_speaker_note(s8, "Our key findings show 100% anomaly detection recall, prevention of multiple duplicate payments, and automated isolation of 3 high-risk rogue vendors.")

    # ==================== SLIDE 9: Testing and validation ====================
    s9 = prs.slides.add_slide(blank_layout)
    add_header_footer_content_slide(s9, "Testing and validation", 9)

    # Passed / total (.5, 1.35, 2.6, 1) Cambria 44 bold red, middle
    pb = s9.shapes.add_textbox(Inches(0.5), Inches(1.35), Inches(2.6), Inches(1.0))
    tf_p = pb.text_frame
    tf_p.vertical_anchor = MSO_ANCHOR.MIDDLE
    p_p = tf_p.paragraphs[0]
    p_p.text = "13 / 13"
    p_p.alignment = PP_ALIGN.CENTER
    p_p.font.name = "Cambria"
    p_p.font.size = Pt(44)
    p_p.font.bold = True
    p_p.font.color.rgb = RED

    # Subtext (.5, 2.4, 2.6, .9) 13 grey, middle
    sb = s9.shapes.add_textbox(Inches(0.5), Inches(2.4), Inches(2.6), Inches(0.9))
    tf_s = sb.text_frame
    tf_s.vertical_anchor = MSO_ANCHOR.MIDDLE
    p_s = tf_s.paragraphs[0]
    p_s.text = "Automated Pytest Tests Passed\nManual Scenarios Verified: 12\nPass Rate: 100%"
    p_s.alignment = PP_ALIGN.CENTER
    p_s.font.name = "Calibri"
    p_s.font.size = Pt(12)
    p_s.font.color.rgb = GREY

    # TABLE at (3.4, 1.35), column widths 1.1, 3.9, 1.1
    t9_shape = s9.shapes.add_table(4, 3, Inches(3.4), Inches(1.35), Inches(6.1), Inches(2.4))
    tbl9 = t9_shape.table
    tbl9.columns[0].width = Inches(1.1)
    tbl9.columns[1].width = Inches(3.9)
    tbl9.columns[2].width = Inches(1.1)

    t9_headers = ["Test", "Checks", "Result"]
    for idx, h in enumerate(t9_headers):
        cell = tbl9.cell(0, idx)
        cell.fill.solid()
        cell.fill.fore_color.rgb = RED
        cell.text_frame.text = h
        p = cell.text_frame.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER
        p.font.name = "Calibri"
        p.font.size = Pt(11)
        p.font.bold = True
        p.font.color.rgb = WHITE

    t9_rows = [
        ("TC-T-05", "Arithmetic Validation (Qty x Price + Tax != Total)", "PASSED"),
        ("TC-T-06", "14-Day Rolling Window Duplicate Collision Engine", "PASSED"),
        ("TC-T-09", "Vendor Risk Index (VRI) Scoring & Tier Bounds", "PASSED")
    ]
    for r_idx, (t_id, chk, res) in enumerate(t9_rows, start=1):
        cell0 = tbl9.cell(r_idx, 0)
        cell1 = tbl9.cell(r_idx, 1)
        cell2 = tbl9.cell(r_idx, 2)
        cell0.text_frame.text = t_id
        cell1.text_frame.text = chk
        cell2.text_frame.text = res
        for c in [cell0, cell1, cell2]:
            c.text_frame.paragraphs[0].font.name = "Calibri"
            c.text_frame.paragraphs[0].font.size = Pt(10)
        cell2.text_frame.paragraphs[0].font.bold = True
        cell2.text_frame.paragraphs[0].font.color.rgb = RGBColor(0x16, 0x65, 0x34)

    set_speaker_note(s9, "All 13 automated technical tests in pytest passed cleanly, confirming mathematical precision, duplicate detection reliability, and relational database integrity.")

    # ==================== SLIDE 10: Recommendation, limitations and next steps ====================
    s10 = prs.slides.add_slide(blank_layout)
    add_header_footer_content_slide(s10, "Recommendation, limitations and next steps", 10)

    # BADGE CARDS y 1.25, w 2.85, h 3.7, x .5 / 3.58 / 6.65
    rec_body = ("Enforce automated pre-payment validation for all vendor invoices above ₹25,000.\n\n"
                "Incorporate Vendor Risk Index (VRI) into annual supplier contract renegotiations.")
    add_badge_card(s10, 0.5, 1.25, 2.85, 3.7, "R", "Recommendation", rec_body)

    lim_body = ("Currently relies on structured digital invoice data; paper receipts require manual pre-digitization.\n\n"
                "Single-instance SQLite backend limits high-concurrency multi-tenant enterprise deployment.")
    add_badge_card(s10, 3.58, 1.25, 2.85, 3.7, "L", "Limitations", lim_body)

    nxt_body = ("Implement OCR pipeline for automated PDF/image invoice parsing.\n\n"
                "Migrate SQLite store to PostgreSQL and build automated Webhook ERP connectors.")
    add_badge_card(s10, 6.65, 1.25, 2.85, 3.7, "N", "Next Steps", nxt_body)

    set_speaker_note(s10, "We recommend adopting automated pre-payment validation for all invoices over ₹25,000. Next steps include adding OCR scanning and PostgreSQL connectors.")

    # ==================== SLIDE 11: DARK (Closing) ====================
    s11 = prs.slides.add_slide(blank_layout)
    bg11 = s11.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, Inches(10), Inches(5.625))
    bg11.fill.solid()
    bg11.fill.fore_color.rgb = BLACK
    bg11.line.fill.background()

    # Title (.5, 1.55, 9, 1.5)
    b_t11 = s11.shapes.add_textbox(Inches(0.5), Inches(1.55), Inches(9.0), Inches(1.5))
    tf_t11 = b_t11.text_frame
    tf_t11.vertical_anchor = MSO_ANCHOR.BOTTOM
    p_t11 = tf_t11.paragraphs[0]
    p_t11.text = "Live Demo & Q&A"
    p_t11.font.name = "Cambria"
    p_t11.font.size = Pt(36)
    p_t11.font.bold = True
    p_t11.font.color.rgb = WHITE

    # Lines (.5, 3.2, 9, 1.7)
    b_l11 = s11.shapes.add_textbox(Inches(0.5), Inches(3.2), Inches(9.0), Inches(1.7))
    tf_l11 = b_l11.text_frame
    tf_l11.vertical_anchor = MSO_ANCHOR.TOP
    p1 = tf_l11.paragraphs[0]
    p1.text = "App: https://invoicehclproject.streamlit.app"
    p1.font.name = "Calibri"
    p1.font.size = Pt(14)
    p1.font.color.rgb = PINK

    p2 = tf_l11.add_paragraph()
    p2.space_before = Pt(6)
    p2.text = "Code: https://github.com/2500320140152/invoice-validation-analytics"
    p2.font.name = "Calibri"
    p2.font.size = Pt(14)
    p2.font.color.rgb = CARD_BG

    p3 = tf_l11.add_paragraph()
    p3.space_before = Pt(8)
    p3.text = "Thank you! Open for questions and viva discussion."
    p3.font.name = "Calibri"
    p3.font.size = Pt(15)
    p3.font.bold = True
    p3.font.color.rgb = WHITE

    set_speaker_note(s11, "Thank you for your time. I invite you to test the live application at our URL and welcome any questions regarding our architecture and findings.")

    prs.save(PPT_PATH)
    print(f"ABES Presentation successfully generated at {PPT_PATH}")

if __name__ == "__main__":
    build_presentation()
