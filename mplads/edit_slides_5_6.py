"""
Edit slides 5 and 6 of ppt_v1.pptx — v2
- Match template fonts exactly (Times New Roman 36pt titles, Arial 12-13pt body)
- Increase text sizes for readability
- Add hyperlinks to references
- Fix alignment
"""
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
import os

# ── Colour palette ─────────────────────────────────────────────────────────
BLUE_DARK   = RGBColor(0x0B, 0x1D, 0x3A)
BLUE_MID    = RGBColor(0x00, 0x70, 0xC0)
WHITE       = RGBColor(0xFF, 0xFF, 0xFF)
BLACK       = RGBColor(0x00, 0x00, 0x00)
GRAY_DARK   = RGBColor(0x33, 0x33, 0x33)
GRAY_MED    = RGBColor(0x64, 0x74, 0x8B)
GRAY_LIGHT  = RGBColor(0xF1, 0xF5, 0xF9)
ACCENT_GREEN  = RGBColor(0x04, 0x78, 0x57)
ACCENT_ORANGE = RGBColor(0xEA, 0x58, 0x0C)

FONT_TITLE = "Times New Roman"
FONT_BODY  = "Arial"

def set_run_font(run, size, bold=False, italic=False, underline=False,
                 color=BLACK, name=FONT_BODY, strike=False):
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.italic = italic
    run.font.underline = underline
    run.font.color.rgb = color
    run.font.name = name
    run.font.strike = strike

def add_textbox(slide, left, top, width, height, text, font_size=12,
                bold=False, underline=False, color=BLACK, align=PP_ALIGN.LEFT,
                font_name=FONT_BODY):
    txBox = slide.shapes.add_textbox(left, top, width, height)
    txBox.text_frame.word_wrap = True
    txBox.text_frame.auto_size = None
    p = txBox.text_frame.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    run.text = text
    set_run_font(run, font_size, bold=bold, underline=underline, color=color, name=font_name)
    return txBox

def add_multiline_textbox(slide, left, top, width, height, lines, font_size=12,
                          bold=False, underline=False, color=BLACK, align=PP_ALIGN.LEFT,
                          font_name=FONT_BODY, line_spacing=1.0):
    """Add a text box with multiple lines, each as a separate paragraph."""
    txBox = slide.shapes.add_textbox(left, top, width, height)
    tf = txBox.text_frame
    tf.word_wrap = True
    tf.auto_size = None
    for i, line_text in enumerate(lines):
        if i == 0:
            p = tf.paragraphs[0]
        else:
            p = tf.add_paragraph()
        p.alignment = align
        p.space_after = Pt(2)
        p.space_before = Pt(0)
        run = p.add_run()
        run.text = line_text
        set_run_font(run, font_size, bold=bold, underline=underline, color=color, name=font_name)
    return txBox

def add_kpi_card(slide, left, top, width, height, number, label, sublabel,
                 bg_color=BLUE_DARK):
    shape = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, top, width, height)
    shape.fill.solid()
    shape.fill.fore_color.rgb = bg_color
    shape.line.fill.background()
    shape.shadow.inherit = False

    tf = shape.text_frame
    tf.word_wrap = True
    tf.auto_size = None
    tf.margin_left = Pt(6)
    tf.margin_right = Pt(6)
    tf.margin_top = Pt(4)
    tf.margin_bottom = Pt(4)

    # Big number
    p1 = tf.paragraphs[0]
    p1.alignment = PP_ALIGN.CENTER
    p1.space_after = Pt(2)
    r1 = p1.add_run()
    r1.text = number
    set_run_font(r1, 20, bold=True, color=WHITE, name=FONT_BODY)

    # Label
    p2 = tf.add_paragraph()
    p2.alignment = PP_ALIGN.CENTER
    p2.space_before = Pt(0)
    p2.space_after = Pt(0)
    r2 = p2.add_run()
    r2.text = label
    set_run_font(r2, 9, bold=True, color=WHITE, name=FONT_BODY)

    # Sublabel
    p3 = tf.add_paragraph()
    p3.alignment = PP_ALIGN.CENTER
    p3.space_before = Pt(0)
    r3 = p3.add_run()
    r3.text = sublabel
    set_run_font(r3, 8, color=RGBColor(0xCA, 0xDC, 0xFC), name=FONT_BODY)

def add_audience_block(slide, left, top, width, emoji, title, desc):
    """Add an audience block with emoji+title bold line and description."""
    # Title line
    txBox = slide.shapes.add_textbox(left, top, width, Pt(16))
    tf = txBox.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.LEFT
    p.space_after = Pt(1)
    r = p.add_run()
    r.text = emoji + "  " + title
    set_run_font(r, 11, bold=True, underline=True, color=BLUE_DARK, name=FONT_BODY)
    return Pt(16)

def add_benefit_block(slide, left, top, width, emoji, title, desc, title_color):
    """Add a benefit block with emoji+title bold line and description."""
    txBox = slide.shapes.add_textbox(left, top, width, Pt(16))
    tf = txBox.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.LEFT
    p.space_after = Pt(1)
    r = p.add_run()
    r.text = emoji + "  " + title
    set_run_font(r, 11, bold=True, underline=True, color=title_color, name=FONT_BODY)
    return Pt(16)

def add_table_cell(cell, text, font_size=9, bold=False, color=GRAY_DARK,
                   align=PP_ALIGN.LEFT, bg_color=None):
    cell.text = ""
    p = cell.text_frame.paragraphs[0]
    p.alignment = align
    p.space_before = Pt(2)
    p.space_after = Pt(2)
    run = p.add_run()
    run.text = text
    set_run_font(run, font_size, bold=bold, color=color, name=FONT_BODY)
    cell.text_frame.word_wrap = True
    cell.text_frame.auto_size = None
    cell.margin_left = Pt(4)
    cell.margin_right = Pt(4)
    cell.margin_top = Pt(3)
    cell.margin_bottom = Pt(3)
    if bg_color:
        cell.fill.solid()
        cell.fill.fore_color.rgb = bg_color

def add_ref_card(slide, left, top, width, num, title, desc, url, dot_color):
    """Draw a single reference card with number, title, description, and hyperlink."""
    card_h = Pt(46)

    # Number circle
    circle_size = Pt(20)
    circle = slide.shapes.add_shape(MSO_SHAPE.OVAL, left, top + Pt(2), circle_size, circle_size)
    circle.fill.solid()
    circle.fill.fore_color.rgb = dot_color
    circle.line.fill.background()
    ctf = circle.text_frame
    ctf.margin_left = Pt(0)
    ctf.margin_right = Pt(0)
    ctf.margin_top = Pt(0)
    ctf.margin_bottom = Pt(0)
    cp = ctf.paragraphs[0]
    cp.alignment = PP_ALIGN.CENTER
    cr = cp.add_run()
    cr.text = str(num)
    set_run_font(cr, 8, bold=True, color=WHITE, name=FONT_BODY)

    # Title + description + hyperlink
    txt_left = left + Pt(24)
    txt_w = width - Pt(24)
    txBox = slide.shapes.add_textbox(txt_left, top, txt_w, card_h)
    tf = txBox.text_frame
    tf.word_wrap = True

    # Title
    p1 = tf.paragraphs[0]
    p1.space_after = Pt(1)
    r1 = p1.add_run()
    r1.text = title
    set_run_font(r1, 10, bold=True, underline=True, color=BLUE_DARK, name=FONT_BODY)

    # Description
    p2 = tf.add_paragraph()
    p2.space_before = Pt(0)
    p2.space_after = Pt(1)
    r2 = p2.add_run()
    r2.text = desc
    set_run_font(r2, 8, color=GRAY_MED, name=FONT_BODY)

    # Hyperlink
    p3 = tf.add_paragraph()
    p3.space_before = Pt(0)
    r3 = p3.add_run()
    r3.text = url
    r3.hyperlink.address = url
    set_run_font(r3, 7, color=BLUE_MID, name=FONT_BODY)

    return card_h


# ══════════════════════════════════════════════════════════════════════════
#  MAIN
# ══════════════════════════════════════════════════════════════════════════
pptx_path = os.path.join(os.path.dirname(__file__), "ppt_v1.pptx")
prs = Presentation(pptx_path)

# ── SLIDE 5: IMPACT AND BENEFITS ─────────────────────────────────────────
slide5 = prs.slides[4]

# Remove old placeholder TextBox (id 17410)
for shape in list(slide5.shapes):
    if shape.shape_id == 17410:
        sp = shape._element
        sp.getparent().remove(sp)
        break

# Update title to match template style: Times New Roman 36pt Bold
for shape in slide5.shapes:
    if shape.shape_id == 17409:
        tf = shape.text_frame
        for p in tf.paragraphs:
            for r in p.runs:
                r.text = "IMPACT & BENEFITS"
                r.font.size = Pt(36)
                r.font.bold = True
                r.font.name = FONT_TITLE
        break

# ── KPI STAT STRIP ────────────────────────────────────────────────────────
kpi_top   = Inches(1.15)
kpi_h     = Inches(0.75)
kpi_w     = Inches(2.15)
kpi_gap   = Inches(0.15)
kpi_left  = Inches(0.55)

kpis = [
    ("Rs.142.8+ Cr", "Est. Annual Leakage", "Intercepted & Saved"),
    ("98.53% -> 0%", "Ghost Asset Deficit", "Eliminated via CV Triad"),
    ("<14 Days", "Sanction SLA Time", "Down from 52 days"),
    ("788 MPs", "700+ Districts", "143K+ Works Covered"),
]

for i, (num, label, sub) in enumerate(kpis):
    left = kpi_left + i * (kpi_w + kpi_gap)
    add_kpi_card(slide5, left, kpi_top, kpi_w, kpi_h, num, label, sub)

# ── SECTION HEADERS (Arial 13pt Bold Underlined — matching template) ──────
sec_top = Inches(2.05)
add_textbox(slide5, Inches(0.55), sec_top, Inches(5.2), Pt(18),
            "IMPACT ON TARGET AUDIENCES", font_size=13, bold=True,
            underline=True, color=BLUE_DARK, font_name=FONT_BODY)
add_textbox(slide5, Inches(6.1), sec_top, Inches(5.6), Pt(18),
            "MULTI-DIMENSIONAL BENEFITS", font_size=13, bold=True,
            underline=True, color=BLUE_DARK, font_name=FONT_BODY)

# ── LEFT COLUMN: Impact on Target Audiences ────────────────────────────────
audiences = [
    ("788 Hon'ble MPs",
     "Instant Statutory Pre-Check: Catches guideline violations (Rule CAT-001) and boundary errors (Rule GEO-001) in <5ms before submission, preventing embarrassing administrative rejections. Quota Health Monitoring: Live tracking of mandatory 15% SC and 7.5% ST area expenditure quotas."),
    ("700+ District Authorities (Collectors / DMs / DPOs)",
     "80% Reduction in Manual File Scrutiny: AI Sentinel automatically flags high-risk proposals, threshold contract splitting (<Rs.50L), and repeat contractor monopolies before the Collector signs. Legally Defensible Sanctions: Provides mathematical SHAP waterfall evidence for every flagged rupee."),
    ("CAG & MoSPI Leadership",
     "Shift from Post-Facto to Pre-Disbursement: Stops leakages before money leaves the treasury, rather than discovering missing funds 3 years later during manual sampling. Pan-India Hotspot Surveillance: Real-time visibility into state-level anomaly clusters across 143,000+ works."),
    ("1.4 Billion Citizens",
     "Ensures public funds actually build tangible community assets (solar drinking water, anganwadis, hospital wings) rather than vanishing into contractor cartels or ghost assets."),
]

y = Inches(2.32)
for title, desc in audiences:
    # Title (bold underlined, 11pt)
    txBox = slide5.shapes.add_textbox(Inches(0.55), y, Inches(5.2), Pt(15))
    tf = txBox.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.LEFT
    p.space_after = Pt(1)
    r = p.add_run()
    r.text = title
    set_run_font(r, 11, bold=True, underline=True, color=BLUE_DARK, name=FONT_BODY)
    y += Pt(15)

    # Description (10pt)
    txBox2 = slide5.shapes.add_textbox(Inches(0.75), y, Inches(5.0), Pt(40))
    tf2 = txBox2.text_frame
    tf2.word_wrap = True
    p2 = tf2.paragraphs[0]
    p2.alignment = PP_ALIGN.LEFT
    p2.space_after = Pt(2)
    r2 = p2.add_run()
    r2.text = desc
    set_run_font(r2, 10, color=GRAY_DARK, name=FONT_BODY)
    y += Pt(42)

# ── RIGHT COLUMN: Multi-Dimensional Benefits ──────────────────────────────
benefits = [
    ("ECONOMIC", "Rs.89 Cr - Rs.178 Cr saved annually (2-4% of Rs.4,466 Cr outlay). Compute cost <Rs.15L/year -> >590x Return on Investment (ROI).", ACCENT_GREEN),
    ("SOCIAL", "Safeguards mandatory allocations for vulnerable communities (Rs.670 Cr reserved annually for SC/ST developmental areas). Direct acceleration of essential public amenities (Drinking Water, Sanitation, Primary Healthcare).", ACCENT_ORANGE),
    ("GOVERNANCE", "Zero Black-Box Resistance: Transparent TreeSHAP attributions empower civil servants rather than replacing them (Human-in-the-loop). End-to-end integration with PFMS / SNA Sparsh prevents idle parking of central funds.", BLUE_MID),
    ("ENVIRONMENTAL", "100% Paperless Audit Trail: Eliminates physical measurement books (MB) and paper sanction files across 700+ Collectorates. Fast-track AI lane for renewable energy works (Solar RO plants, solar street lighting, rainwater harvesting).", ACCENT_GREEN),
]

y = Inches(2.32)
for title, desc, clr in benefits:
    txBox = slide5.shapes.add_textbox(Inches(6.1), y, Inches(5.6), Pt(15))
    tf = txBox.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.LEFT
    p.space_after = Pt(1)
    r = p.add_run()
    r.text = title
    set_run_font(r, 11, bold=True, underline=True, color=clr, name=FONT_BODY)
    y += Pt(15)

    txBox2 = slide5.shapes.add_textbox(Inches(6.3), y, Inches(5.4), Pt(40))
    tf2 = txBox2.text_frame
    tf2.word_wrap = True
    p2 = tf2.paragraphs[0]
    p2.alignment = PP_ALIGN.LEFT
    p2.space_after = Pt(2)
    r2 = p2.add_run()
    r2.text = desc
    set_run_font(r2, 10, color=GRAY_DARK, name=FONT_BODY)
    y += Pt(42)

# ── Table label (Arial 13pt Bold Underlined) ──────────────────────────────
add_textbox(slide5, Inches(0.55), Inches(4.55), Inches(8), Pt(16),
            "WHAT WE BUILT vs. WHAT INDIA GAINS (The Shortlisting Moat Table)",
            font_size=12, bold=True, underline=True, color=BLUE_DARK, font_name=FONT_BODY)

# ── Bottom Comparison Table ────────────────────────────────────────────────
table_top = Inches(4.8)
table_left = Inches(0.55)
table_w = Inches(11.1)
table_h = Inches(1.4)

rows, cols = 5, 3
table_shape = slide5.shapes.add_table(rows, cols, table_left, table_top, table_w, table_h)
table = table_shape.table

table.columns[0].width = Inches(2.0)
table.columns[1].width = Inches(4.55)
table.columns[2].width = Inches(4.55)

headers = ["Scheme Metric", "Current eSAKSHI Baseline", "With AI Sentinel (What India Gains)"]
for i, h in enumerate(headers):
    add_table_cell(table.cell(0, i), h, font_size=10, bold=True,
                   color=WHITE, align=PP_ALIGN.CENTER, bg_color=BLUE_DARK)

data = [
    ("Asset Verification",
     "98.53% unverified inspection records (CAG finding)",
     "100% Automated Pre-Disbursement CV Check (De-Fake AI + CLIP + EXIF Geotag)"),
    ("Contractor Cartel Detection",
     "0% automated; cartels operate across district borders",
     "Graph Network Analysis flags bidding syndicates sharing phone/bank hashes (>50% share)"),
    ("Threshold Contract Splitting",
     "Undetected; split into Rs.49.5L to bypass clearance",
     "Isolation Forest flags artificial clusters within 10% of Rs.50L Technical Committee ceiling"),
    ("Audit Turnaround Time",
     "2-3 Years (Slow post-facto sampling by CAG)",
     "<100ms Real-time automated SHAP digital audit dossier per work"),
]

for r, (metric, baseline, gain) in enumerate(data, start=1):
    bg = GRAY_LIGHT if r % 2 == 0 else None
    add_table_cell(table.cell(r, 0), metric, font_size=9, bold=True, color=BLUE_DARK, bg_color=bg)
    add_table_cell(table.cell(r, 1), baseline, font_size=9, color=GRAY_DARK, bg_color=bg)
    add_table_cell(table.cell(r, 2), gain, font_size=9, color=ACCENT_GREEN, bg_color=bg)

# ══════════════════════════════════════════════════════════════════════════
#  SLIDE 6: REFERENCES
# ══════════════════════════════════════════════════════════════════════════
slide6 = prs.slides[5]

# Remove old placeholder TextBox (id 17410)
for shape in list(slide6.shapes):
    if shape.shape_id == 17410:
        sp = shape._element
        sp.getparent().remove(sp)
        break

# Update title: Times New Roman 36pt Bold (matching template)
for shape in slide6.shapes:
    if shape.shape_id == 17409:
        tf = shape.text_frame
        for p in tf.paragraphs:
            for r in p.runs:
                r.text = "REFERENCES"
                r.font.size = Pt(36)
                r.font.bold = True
                r.font.name = FONT_TITLE
        break

# ── Two-column layout ─────────────────────────────────────────────────────
col1_left = Inches(0.45)
col2_left = Inches(6.15)
col_w     = Inches(5.5)
ref_top   = Inches(1.15)

# Section headers: Arial 13pt Bold Underlined
add_textbox(slide6, col1_left, ref_top, col_w, Pt(18),
            "GOVERNMENT & REGULATORY SOURCES", font_size=13, bold=True,
            underline=True, color=BLUE_MID, font_name=FONT_BODY)
add_textbox(slide6, col2_left, ref_top, col_w, Pt(18),
            "TECHNICAL & RESEARCH REFERENCES", font_size=13, bold=True,
            underline=True, color=ACCENT_GREEN, font_name=FONT_BODY)

gov_refs = [
    ("1", "CAG Performance Audits of MPLADS (2005-2024)",
     "85 documented fraud/misuse cases worth Rs.32,133+ Cr across multiple parliamentary terms.",
     "https://www.cag.gov.in"),
    ("2", "MPLADS Guidelines 2023",
     "MoSPI official rules governing fund allocation, disbursement, and utilization.",
     "https://mplads.mospi.gov.in"),
    ("3", "eSAKSHI Portal (esakshi.mp.gov.in)",
     "Existing digital transaction tracking system for MPLADS works and payments.",
     "https://esakshi.mp.gov.in"),
    ("4", "data.gov.in MPLADS Dataset",
     "143,257 records of works and expenditures across India for ML training data.",
     "https://data.gov.in"),
    ("5", "DPDP Act 2023",
     "Digital Personal Data Protection Act compliance framework for AI systems.",
     "https://www.indiacode.nic.in"),
    ("6", "RTI Act 2005",
     "Right to Information Act transparency requirements for public fund monitoring.",
     "https://rtionline.gov.in"),
    ("7", "ARC Report (2011) on MPLADS",
     "Administrative Reforms Commission recommended abolishing scheme due to corruption.",
     "https://arc.gov.in"),
    ("8", "MPLADS Annual Reports (2019-2024)",
     "Fund utilization: 37-52% across terms; Rs.1,092.5 Cr lapsed during COVID.",
     "https://mospi.gov.in"),
]

tech_refs = [
    ("9", "Lundberg & Lee (2017) - SHAP",
     "A Unified Approach to Interpreting Model Predictions. NeurIPS 2017. Foundation of our explainability layer.",
     "https://proceedings.neurips.cc/paper/2017/hash/8a20a862197818fbd76f4f08898fc436-Abstract.html"),
    ("10", "Liu et al. (2012) - Isolation Forest",
     "Anomaly detection algorithm. IEEE TKDE. Core for unsupervised zero-day anomaly detection.",
     "https://ieeexplore.ieee.org/document/6040193"),
    ("11", "RedFlags.ai (Kazakhstan)",
     "Saved $86M in 6 months using rule-based fraud detection. Validated government monitoring approach.",
     "https://redflags.kz"),
    ("12", "Tazama (Linux Foundation)",
     "Open-source real-time payment monitoring at 2,300 TPS. Architecture reference for throughput.",
     "https://tazama.io"),
    ("13", "UK HMRC / Quantexa Platform",
     "GBP 175M entity resolution for tax fraud. Our approach achieves similar goals at minimal cost.",
     "https://www.quantexa.com"),
    ("14", "RBI MuleHunter.AI",
     "85-90% accuracy rate with 26+ bank partnerships for fraud detection. ML performance benchmark.",
     "https://www.rbi.org.in"),
    ("15", "Flower (flwr.ai)",
     "Production-ready federated learning framework. Foundation for cross-state model training.",
     "https://flower.ai"),
    ("16", "Nigrini (2012) - Benford's Law",
     "Application to financial fraud detection. Statistical methodology for transaction analysis.",
     "https://www.benfordslaw.org"),
]

y = Inches(1.4)
for num, title, desc, url in gov_refs:
    h = add_ref_card(slide6, col1_left, y, col_w, num, title, desc, url, BLUE_MID)
    y += h + Pt(3)

y = Inches(1.4)
for num, title, desc, url in tech_refs:
    h = add_ref_card(slide6, col2_left, y, col_w, num, title, desc, url, ACCENT_GREEN)
    y += h + Pt(3)

# ── Save ───────────────────────────────────────────────────────────────────
out_path = os.path.join(os.path.dirname(__file__), "ppt_v1_edited.pptx")
prs.save(out_path)
print("Saved:", out_path)
