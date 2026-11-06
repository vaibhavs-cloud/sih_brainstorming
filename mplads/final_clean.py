"""
Final clean rebuild of slides 5 and 6.
Matches the clean Image 1 layout exactly, adds hyperlinks.
Fixes Image 3 overflow by using proper text sizing.
"""
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE
import os

FONT_TITLE = "Times New Roman"
FONT_BODY  = "Arial"

BLUE_DARK   = RGBColor(0x0B, 0x1D, 0x3A)
BLUE_MID    = RGBColor(0x00, 0x70, 0xC0)
WHITE       = RGBColor(0xFF, 0xFF, 0xFF)
BLACK       = RGBColor(0x00, 0x00, 0x00)
GRAY_DARK   = RGBColor(0x33, 0x33, 0x33)
GRAY_MED    = RGBColor(0x64, 0x74, 0x8B)
GRAY_LIGHT  = RGBColor(0xF1, 0xF5, 0xF9)
ACCENT_GREEN  = RGBColor(0x04, 0x78, 0x57)
ACCENT_ORANGE = RGBColor(0xEA, 0x58, 0x0C)

def sr(run, size, bold=False, ul=False, color=BLACK, name=FONT_BODY):
    """Set run font properties."""
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.underline = ul
    run.font.color.rgb = color
    run.font.name = name

def tb(slide, left, top, w, h, text, sz=12, bold=False, ul=False,
       color=BLACK, align=PP_ALIGN.LEFT, name=FONT_BODY):
    """Add a single-line textbox."""
    box = slide.shapes.add_textbox(left, top, w, h)
    box.text_frame.word_wrap = True
    box.text_frame.auto_size = None
    p = box.text_frame.paragraphs[0]
    p.alignment = align
    r = p.add_run()
    r.text = text
    sr(r, sz, bold, ul, color, name)
    return box

def kpi_card(slide, left, top, w, h, num, label, sub, bg=BLUE_DARK):
    s = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, top, w, h)
    s.fill.solid()
    s.fill.fore_color.rgb = bg
    s.line.fill.background()
    s.shadow.inherit = False
    tf = s.text_frame
    tf.word_wrap = True
    tf.auto_size = None
    tf.margin_left = Pt(6)
    tf.margin_right = Pt(6)
    tf.margin_top = Pt(3)
    tf.margin_bottom = Pt(3)

    p1 = tf.paragraphs[0]
    p1.alignment = PP_ALIGN.CENTER
    p1.space_after = Pt(1)
    r1 = p1.add_run()
    r1.text = num
    sr(r1, 18, bold=True, color=WHITE)

    p2 = tf.add_paragraph()
    p2.alignment = PP_ALIGN.CENTER
    p2.space_before = Pt(0)
    p2.space_after = Pt(0)
    r2 = p2.add_run()
    r2.text = label
    sr(r2, 8, bold=True, color=WHITE)

    p3 = tf.add_paragraph()
    p3.alignment = PP_ALIGN.CENTER
    p3.space_before = Pt(0)
    r3 = p3.add_run()
    r3.text = sub
    sr(r3, 7, color=RGBColor(0xCA, 0xDC, 0xFC))

def tcell(cell, text, sz=9, bold=False, color=GRAY_DARK,
          align=PP_ALIGN.LEFT, bg=None):
    cell.text = ""
    p = cell.text_frame.paragraphs[0]
    p.alignment = align
    p.space_before = Pt(2)
    p.space_after = Pt(2)
    r = p.add_run()
    r.text = text
    sr(r, sz, bold, color=color)
    cell.text_frame.word_wrap = True
    cell.text_frame.auto_size = None
    cell.margin_left = Pt(4)
    cell.margin_right = Pt(4)
    cell.margin_top = Pt(2)
    cell.margin_bottom = Pt(2)
    if bg:
        cell.fill.solid()
        cell.fill.fore_color.rgb = bg

def ref_card(slide, left, top, w, num, title, desc, url, dot_clr):
    """Single reference card: circle + title + desc + hyperlink."""
    # Number circle
    cs = Pt(18)
    c = slide.shapes.add_shape(MSO_SHAPE.OVAL, left, top + Pt(1), cs, cs)
    c.fill.solid()
    c.fill.fore_color.rgb = dot_clr
    c.line.fill.background()
    ctf = c.text_frame
    ctf.margin_left = Pt(0)
    ctf.margin_right = Pt(0)
    ctf.margin_top = Pt(0)
    ctf.margin_bottom = Pt(0)
    cp = ctf.paragraphs[0]
    cp.alignment = PP_ALIGN.CENTER
    cr = cp.add_run()
    cr.text = str(num)
    sr(cr, 7, bold=True, color=WHITE)

    # Text area
    tl = left + Pt(22)
    tw = w - Pt(22)
    box = slide.shapes.add_textbox(tl, top, tw, Pt(48))
    tf = box.text_frame
    tf.word_wrap = True

    # Title (bold underlined)
    p1 = tf.paragraphs[0]
    p1.space_after = Pt(0)
    p1.space_before = Pt(0)
    r1 = p1.add_run()
    r1.text = title
    sr(r1, 9, bold=True, ul=True, color=BLUE_DARK)

    # Description
    p2 = tf.add_paragraph()
    p2.space_before = Pt(1)
    p2.space_after = Pt(1)
    r2 = p2.add_run()
    r2.text = desc
    sr(r2, 8, color=GRAY_MED)

    # Hyperlink URL
    p3 = tf.add_paragraph()
    p3.space_before = Pt(0)
    r3 = p3.add_run()
    r3.text = url
    r3.hyperlink.address = url
    sr(r3, 7, color=BLUE_MID)

    return Pt(48)


# ══════════════════════════════════════════════════════════════════════════
#  MAIN
# ══════════════════════════════════════════════════════════════════════════
pptx_path = os.path.join(os.path.dirname(__file__), "ppt_v1_final.pptx")
prs = Presentation(pptx_path)

# ══════════════════════════════════════════════════════════════════════════
#  SLIDE 5: IMPACT & BENEFITS
# ══════════════════════════════════════════════════════════════════════════
slide5 = prs.slides[4]

# Remove old placeholder text boxes (the "Potential impact..." and "Benefits..." boxes)
for shape in list(slide5.shapes):
    if shape.name == 'TextBox 8':
        sp = shape._element
        sp.getparent().remove(sp)

# Update title
for shape in slide5.shapes:
    if shape.name == 'Title 1':
        for p in shape.text_frame.paragraphs:
            for r in p.runs:
                r.text = "IMPACT & BENEFITS"
                r.font.size = Pt(36)
                r.font.bold = True
                r.font.name = FONT_TITLE
        break

# ── KPI Cards ──────────────────────────────────────────────────────────────
kpi_y = Inches(1.12)
kpi_h = Inches(0.7)
kpi_w = Inches(2.15)
kpi_gap = Inches(0.15)
kpi_x = Inches(0.55)

kpis = [
    ("Rs.142.8+ Cr", "Est. Annual Leakage", "Intercepted & Saved"),
    ("98.53% -> 0%", "Ghost Asset Deficit", "Eliminated via CV Triad"),
    ("<14 Days", "Sanction SLA Time", "Down from 52 days"),
    ("788 MPs", "700+ Districts", "143K+ Works Covered"),
]
for i, (n, l, s) in enumerate(kpis):
    kpi_card(slide5, kpi_x + i * (kpi_w + kpi_gap), kpi_y, kpi_w, kpi_h, n, l, s)

# ── Section Headers ────────────────────────────────────────────────────────
tb(slide5, Inches(0.55), Inches(1.95), Inches(5.2), Pt(16),
   "IMPACT ON TARGET AUDIENCES", 12, bold=True, ul=True, color=BLUE_DARK)
tb(slide5, Inches(6.1), Inches(1.95), Inches(5.6), Pt(16),
   "MULTI-DIMENSIONAL BENEFITS", 12, bold=True, ul=True, color=BLUE_DARK)

# ── Left Column: Audiences ─────────────────────────────────────────────────
# Each block: 1-line bold title + 2-3 line description
# Total available height: ~2.6 inches (from y=2.2 to y=4.8)
# 4 blocks, each ~0.6 inches

audiences = [
    ("788 Hon'ble MPs",
     "Instant Statutory Pre-Check: Catches guideline violations in <5ms before submission. Quota Health Monitoring: Live tracking of 15% SC and 7.5% ST expenditure quotas."),
    ("700+ District Authorities",
     "80% Reduction in Manual Scrutiny: AI Sentinel flags high-risk proposals, threshold splitting (<Rs.50L), and repeat contractor monopolies. Legally Defensible Sanctions via SHAP evidence."),
    ("CAG & MoSPI Leadership",
     "Shift from Post-Facto to Pre-Disbursement: Stops leakages before money leaves the treasury. Pan-India Hotspot Surveillance across 143,000+ works."),
    ("1.4 Billion Citizens",
     "Ensures public funds build tangible community assets (solar water, anganwadis, hospitals) rather than vanishing into contractor cartels or ghost assets."),
]

y = Inches(2.18)
for title, desc in audiences:
    # Bold underlined title (10pt)
    box = slide5.shapes.add_textbox(Inches(0.55), y, Inches(5.2), Pt(13))
    tf = box.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.space_after = Pt(1)
    r = p.add_run()
    r.text = title
    sr(r, 10, bold=True, ul=True, color=BLUE_DARK)
    y += Pt(14)

    # Description (9pt)
    box2 = slide5.shapes.add_textbox(Inches(0.75), y, Inches(5.0), Pt(36))
    tf2 = box2.text_frame
    tf2.word_wrap = True
    p2 = tf2.paragraphs[0]
    p2.space_after = Pt(2)
    r2 = p2.add_run()
    r2.text = desc
    sr(r2, 9, color=GRAY_DARK)
    y += Pt(38)

# ── Right Column: Benefits ─────────────────────────────────────────────────
benefits = [
    ("ECONOMIC", "Rs.89 Cr - Rs.178 Cr saved annually (2-4% of Rs.4,466 Cr outlay). Compute cost <Rs.15L/year -> >590x ROI.", ACCENT_GREEN),
    ("SOCIAL", "Safeguards mandatory allocations for vulnerable communities (Rs.670 Cr reserved for SC/ST). Direct acceleration of drinking water, sanitation, healthcare.", ACCENT_ORANGE),
    ("GOVERNANCE", "Zero Black-Box Resistance: Transparent TreeSHAP attributions empower civil servants (Human-in-the-loop). End-to-end integration with PFMS / SNA Sparsh.", BLUE_MID),
    ("ENVIRONMENTAL", "100% Paperless Audit Trail across 700+ Collectorates. Fast-track AI lane for renewable energy works (Solar RO, street lighting, rainwater harvesting).", ACCENT_GREEN),
]

y = Inches(2.18)
for title, desc, clr in benefits:
    box = slide5.shapes.add_textbox(Inches(6.1), y, Inches(5.6), Pt(13))
    tf = box.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.space_after = Pt(1)
    r = p.add_run()
    r.text = title
    sr(r, 10, bold=True, ul=True, color=clr)
    y += Pt(14)

    box2 = slide5.shapes.add_textbox(Inches(6.3), y, Inches(5.4), Pt(36))
    tf2 = box2.text_frame
    tf2.word_wrap = True
    p2 = tf2.paragraphs[0]
    p2.space_after = Pt(2)
    r2 = p2.add_run()
    r2.text = desc
    sr(r2, 9, color=GRAY_DARK)
    y += Pt(38)

# ── Table Label ────────────────────────────────────────────────────────────
tb(slide5, Inches(0.55), Inches(4.35), Inches(8), Pt(14),
   "WHAT WE BUILT vs. WHAT INDIA GAINS", 11, bold=True, ul=True, color=BLUE_DARK)

# ── Comparison Table ───────────────────────────────────────────────────────
tbl_s = slide5.shapes.add_table(5, 3, Inches(0.55), Inches(4.55), Inches(11.1), Inches(1.5))
tbl = tbl_s.table
tbl.columns[0].width = Inches(2.0)
tbl.columns[1].width = Inches(4.55)
tbl.columns[2].width = Inches(4.55)

for i, h in enumerate(["Scheme Metric", "Current eSAKSHI Baseline", "With AI Sentinel"]):
    tcell(tbl.cell(0, i), h, 10, bold=True, color=WHITE, align=PP_ALIGN.CENTER, bg=BLUE_DARK)

data = [
    ("Asset Verification", "98.53% unverified (CAG finding)", "100% Automated CV Check (De-Fake + CLIP + EXIF)"),
    ("Cartel Detection", "0% automated detection", "Graph Network flags bidding syndicates (>50% share)"),
    ("Contract Splitting", "Undetected; split at Rs.49.5L", "Isolation Forest flags clusters within 10% of Rs.50L"),
    ("Audit Turnaround", "2-3 Years (post-facto sampling)", "<100ms Real-time SHAP audit dossier"),
]
for r, (m, b, g) in enumerate(data, 1):
    bg = GRAY_LIGHT if r % 2 == 0 else None
    tcell(tbl.cell(r, 0), m, 9, bold=True, color=BLUE_DARK, bg=bg)
    tcell(tbl.cell(r, 1), b, 9, color=GRAY_DARK, bg=bg)
    tcell(tbl.cell(r, 2), g, 9, color=ACCENT_GREEN, bg=bg)


# ══════════════════════════════════════════════════════════════════════════
#  SLIDE 6: REFERENCES (clean layout matching Image 1)
# ══════════════════════════════════════════════════════════════════════════
slide6 = prs.slides[5]

# Remove old placeholder
for shape in list(slide6.shapes):
    if shape.name == 'TextBox 8':
        sp = shape._element
        sp.getparent().remove(sp)

# Update title
for shape in slide6.shapes:
    if shape.name == 'Title 1':
        for p in shape.text_frame.paragraphs:
            for r in p.runs:
                r.text = "REFERENCES"
                r.font.size = Pt(36)
                r.font.bold = True
                r.font.name = FONT_TITLE
        break

# ── Section Headers ────────────────────────────────────────────────────────
tb(slide6, Inches(0.45), Inches(1.15), Inches(5.5), Pt(16),
   "GOVERNMENT & REGULATORY SOURCES", 12, bold=True, ul=True, color=BLUE_MID)
tb(slide6, Inches(6.15), Inches(1.15), Inches(5.5), Pt(16),
   "TECHNICAL & RESEARCH REFERENCES", 12, bold=True, ul=True, color=ACCENT_GREEN)

# ── References (matching Image 1 layout exactly) ──────────────────────────
# Image 1 spacing: each card ~0.65 inches apart vertically
# 8 cards per column, from y=1.45 to y=~6.8 (fits in slide)

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

# Each card is Pt(48) tall + Pt(4) gap = Pt(52) per card
# 8 cards = 8 * 52 = Pt(416) = ~5.78 inches
# Starting at y=1.45, ending at ~7.23 (fits in 7.5" slide)

card_h = Pt(48)
card_gap = Pt(4)
y_start = Inches(1.45)

y = y_start
for num, title, desc, url in gov_refs:
    h = ref_card(slide6, Inches(0.45), y, Inches(5.5), num, title, desc, url, BLUE_MID)
    y += h + card_gap

y = y_start
for num, title, desc, url in tech_refs:
    h = ref_card(slide6, Inches(6.15), y, Inches(5.5), num, title, desc, url, ACCENT_GREEN)
    y += h + card_gap

# ── Save ───────────────────────────────────────────────────────────────────
out = os.path.join(os.path.dirname(__file__), "ppt_v1_v3.pptx")
prs.save(out)
print("Saved:", out)
