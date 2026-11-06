"""
Clean rebuild of slides 5 and 6 using python-pptx.
Removes all non-template shapes and rebuilds with proper formatting.
"""
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
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

# Template shape IDs to preserve (from original template)
TEMPLATE_IDS_SLIDE5 = {10, 17409, 6, 7, 11, 2}  # Rectangle 9, Title 1, SlideNum, Footer, Picture, Oval
TEMPLATE_IDS_SLIDE6 = {10, 17409, 6, 7, 12, 11}  # Rectangle 9, Title 1, SlideNum, Footer, Picture, Oval

TEMPLATE_NAMES = {'Rectangle 9', 'Title 1', 'Slide Number Placeholder 5',
                  'Footer Placeholder 6', 'Nan-Chalants'}

def set_run(run, size, bold=False, underline=False, color=BLACK, name=FONT_BODY):
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.underline = underline
    run.font.color.rgb = color
    run.font.name = name

def add_textbox(slide, left, top, width, height, text, size=12, bold=False,
                underline=False, color=BLACK, align=PP_ALIGN.LEFT, name=FONT_BODY):
    tb = slide.shapes.add_textbox(left, top, width, height)
    tb.text_frame.word_wrap = True
    tb.text_frame.auto_size = None
    p = tb.text_frame.paragraphs[0]
    p.alignment = align
    r = p.add_run()
    r.text = text
    set_run(r, size, bold, underline, color, name)
    return tb

def add_kpi(slide, left, top, w, h, num, label, sub, bg=BLUE_DARK):
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
    tf.margin_top = Pt(4)
    tf.margin_bottom = Pt(4)

    p1 = tf.paragraphs[0]
    p1.alignment = PP_ALIGN.CENTER
    p1.space_after = Pt(2)
    r1 = p1.add_run()
    r1.text = num
    set_run(r1, 20, bold=True, color=WHITE)

    p2 = tf.add_paragraph()
    p2.alignment = PP_ALIGN.CENTER
    p2.space_before = Pt(0)
    p2.space_after = Pt(0)
    r2 = p2.add_run()
    r2.text = label
    set_run(r2, 9, bold=True, color=WHITE)

    p3 = tf.add_paragraph()
    p3.alignment = PP_ALIGN.CENTER
    p3.space_before = Pt(0)
    r3 = p3.add_run()
    r3.text = sub
    set_run(r3, 8, color=RGBColor(0xCA, 0xDC, 0xFC))

def add_table_cell(cell, text, size=9, bold=False, color=GRAY_DARK,
                   align=PP_ALIGN.LEFT, bg=None):
    cell.text = ""
    p = cell.text_frame.paragraphs[0]
    p.alignment = align
    p.space_before = Pt(2)
    p.space_after = Pt(2)
    r = p.add_run()
    r.text = text
    set_run(r, size, bold, color=color)
    cell.text_frame.word_wrap = True
    cell.text_frame.auto_size = None
    cell.margin_left = Pt(4)
    cell.margin_right = Pt(4)
    cell.margin_top = Pt(3)
    cell.margin_bottom = Pt(3)
    if bg:
        cell.fill.solid()
        cell.fill.fore_color.rgb = bg

def add_ref_card(slide, left, top, w, num, title, desc, url, dot_clr):
    # Number circle
    cs = Pt(20)
    c = slide.shapes.add_shape(MSO_SHAPE.OVAL, left, top + Pt(2), cs, cs)
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
    set_run(cr, 8, bold=True, color=WHITE)

    # Text box
    tl = left + Pt(24)
    tw = w - Pt(24)
    tb = slide.shapes.add_textbox(tl, top, tw, Pt(50))
    tf = tb.text_frame
    tf.word_wrap = True

    p1 = tf.paragraphs[0]
    p1.space_after = Pt(1)
    r1 = p1.add_run()
    r1.text = title
    set_run(r1, 10, bold=True, underline=True, color=BLUE_DARK)

    p2 = tf.add_paragraph()
    p2.space_before = Pt(0)
    p2.space_after = Pt(1)
    r2 = p2.add_run()
    r2.text = desc
    set_run(r2, 8, color=GRAY_MED)

    p3 = tf.add_paragraph()
    p3.space_before = Pt(0)
    r3 = p3.add_run()
    r3.text = url
    r3.hyperlink.address = url
    set_run(r3, 7, color=BLUE_MID)

    return Pt(50)


# ══════════════════════════════════════════════════════════════════════════
#  MAIN
# ══════════════════════════════════════════════════════════════════════════
pptx_path = os.path.join(os.path.dirname(__file__), "ppt_v1.pptx")
prs = Presentation(pptx_path)

# ── SLIDE 5 ────────────────────────────────────────────────────────────────
slide5 = prs.slides[4]

# Remove all shapes that are NOT template elements
# Keep: Rectangle 9 (footer), Title 1, Slide Number, Footer, Picture 10, Oval 1
# Remove: everything else (old KPI cards, textboxes, table, etc.)
keep_ids = set()
remove_ids = []
for shape in slide5.shapes:
    name = shape.name
    sid = shape.shape_id
    is_template = (name in TEMPLATE_NAMES or
                   name.startswith('Rectangle') or
                   name.startswith('Title') or
                   name.startswith('Slide Number') or
                   name.startswith('Footer') or
                   name.startswith('Picture') or
                   name.startswith('Oval'))
    if is_template:
        keep_ids.add(sid)
    else:
        remove_ids.append(shape)

print('Slide 5: keeping', len(keep_ids), 'template shapes, removing', len(remove_ids), 'content shapes')
for shape in remove_ids:
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

# ── Add KPI cards ──────────────────────────────────────────────────────────
kpi_top = Inches(1.15)
kpi_h = Inches(0.75)
kpi_w = Inches(2.15)
kpi_gap = Inches(0.15)
kpi_left = Inches(0.55)

kpis = [
    ("Rs.142.8+ Cr", "Est. Annual Leakage", "Intercepted & Saved"),
    ("98.53% -> 0%", "Ghost Asset Deficit", "Eliminated via CV Triad"),
    ("<14 Days", "Sanction SLA Time", "Down from 52 days"),
    ("788 MPs", "700+ Districts", "143K+ Works Covered"),
]
for i, (num, label, sub) in enumerate(kpis):
    add_kpi(slide5, kpi_left + i * (kpi_w + kpi_gap), kpi_top, kpi_w, kpi_h, num, label, sub)

# ── Section headers ────────────────────────────────────────────────────────
add_textbox(slide5, Inches(0.55), Inches(2.05), Inches(5.2), Pt(18),
            "IMPACT ON TARGET AUDIENCES", 13, bold=True, underline=True, color=BLUE_DARK)
add_textbox(slide5, Inches(6.1), Inches(2.05), Inches(5.6), Pt(18),
            "MULTI-DIMENSIONAL BENEFITS", 13, bold=True, underline=True, color=BLUE_DARK)

# ── Left: Audiences ────────────────────────────────────────────────────────
audiences = [
    ("788 Hon'ble MPs",
     "Instant Statutory Pre-Check: Catches guideline violations (Rule CAT-001) and boundary errors (Rule GEO-001) in <5ms before submission. Quota Health Monitoring: Live tracking of mandatory 15% SC and 7.5% ST area expenditure quotas."),
    ("700+ District Authorities (Collectors / DMs / DPOs)",
     "80% Reduction in Manual File Scrutiny: AI Sentinel automatically flags high-risk proposals, threshold contract splitting (<Rs.50L), and repeat contractor monopolies. Legally Defensible Sanctions: Provides mathematical SHAP waterfall evidence for every flagged rupee."),
    ("CAG & MoSPI Leadership",
     "Shift from Post-Facto to Pre-Disbursement: Stops leakages before money leaves the treasury. Pan-India Hotspot Surveillance: Real-time visibility into state-level anomaly clusters across 143,000+ works."),
    ("1.4 Billion Citizens",
     "Ensures public funds actually build tangible community assets (solar drinking water, anganwadis, hospital wings) rather than vanishing into contractor cartels or ghost assets."),
]

y = Inches(2.32)
for title, desc in audiences:
    tb = slide5.shapes.add_textbox(Inches(0.55), y, Inches(5.2), Pt(15))
    tf = tb.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.LEFT
    p.space_after = Pt(1)
    r = p.add_run()
    r.text = title
    set_run(r, 11, bold=True, underline=True, color=BLUE_DARK)
    y += Pt(15)

    tb2 = slide5.shapes.add_textbox(Inches(0.75), y, Inches(5.0), Pt(40))
    tf2 = tb2.text_frame
    tf2.word_wrap = True
    p2 = tf2.paragraphs[0]
    p2.alignment = PP_ALIGN.LEFT
    p2.space_after = Pt(2)
    r2 = p2.add_run()
    r2.text = desc
    set_run(r2, 10, color=GRAY_DARK)
    y += Pt(42)

# ── Right: Benefits ────────────────────────────────────────────────────────
benefits = [
    ("ECONOMIC", "Rs.89 Cr - Rs.178 Cr saved annually (2-4% of Rs.4,466 Cr outlay). Compute cost <Rs.15L/year -> >590x ROI.", ACCENT_GREEN),
    ("SOCIAL", "Safeguards mandatory allocations for vulnerable communities (Rs.670 Cr reserved annually for SC/ST areas). Direct acceleration of drinking water, sanitation, healthcare.", ACCENT_ORANGE),
    ("GOVERNANCE", "Zero Black-Box Resistance: Transparent TreeSHAP attributions empower civil servants (Human-in-the-loop). End-to-end integration with PFMS / SNA Sparsh.", BLUE_MID),
    ("ENVIRONMENTAL", "100% Paperless Audit Trail across 700+ Collectorates. Fast-track AI lane for renewable energy works (Solar RO, solar street lighting, rainwater harvesting).", ACCENT_GREEN),
]

y = Inches(2.32)
for title, desc, clr in benefits:
    tb = slide5.shapes.add_textbox(Inches(6.1), y, Inches(5.6), Pt(15))
    tf = tb.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.LEFT
    p.space_after = Pt(1)
    r = p.add_run()
    r.text = title
    set_run(r, 11, bold=True, underline=True, color=clr)
    y += Pt(15)

    tb2 = slide5.shapes.add_textbox(Inches(6.3), y, Inches(5.4), Pt(40))
    tf2 = tb2.text_frame
    tf2.word_wrap = True
    p2 = tf2.paragraphs[0]
    p2.alignment = PP_ALIGN.LEFT
    p2.space_after = Pt(2)
    r2 = p2.add_run()
    r2.text = desc
    set_run(r2, 10, color=GRAY_DARK)
    y += Pt(42)

# ── Table label ────────────────────────────────────────────────────────────
add_textbox(slide5, Inches(0.55), Inches(4.55), Inches(8), Pt(16),
            "WHAT WE BUILT vs. WHAT INDIA GAINS", 12, bold=True, underline=True, color=BLUE_DARK)

# ── Comparison Table ───────────────────────────────────────────────────────
tbl_shape = slide5.shapes.add_table(5, 3, Inches(0.55), Inches(4.8), Inches(11.1), Inches(1.4))
tbl = tbl_shape.table
tbl.columns[0].width = Inches(2.0)
tbl.columns[1].width = Inches(4.55)
tbl.columns[2].width = Inches(4.55)

for i, h in enumerate(["Scheme Metric", "Current eSAKSHI Baseline", "With AI Sentinel"]):
    add_table_cell(tbl.cell(0, i), h, 10, bold=True, color=WHITE, align=PP_ALIGN.CENTER, bg=BLUE_DARK)

data = [
    ("Asset Verification", "98.53% unverified (CAG finding)", "100% Automated CV Check (De-Fake + CLIP + EXIF)"),
    ("Cartel Detection", "0% automated detection", "Graph Network flags bidding syndicates (>50% share)"),
    ("Contract Splitting", "Undetected; split at Rs.49.5L", "Isolation Forest flags clusters within 10% of Rs.50L"),
    ("Audit Turnaround", "2-3 Years (post-facto sampling)", "<100ms Real-time SHAP audit dossier"),
]
for r, (m, b, g) in enumerate(data, 1):
    bg = GRAY_LIGHT if r % 2 == 0 else None
    add_table_cell(tbl.cell(r, 0), m, 9, bold=True, color=BLUE_DARK, bg=bg)
    add_table_cell(tbl.cell(r, 1), b, 9, color=GRAY_DARK, bg=bg)
    add_table_cell(tbl.cell(r, 2), g, 9, color=ACCENT_GREEN, bg=bg)


# ══════════════════════════════════════════════════════════════════════════
#  SLIDE 6: REFERENCES
# ══════════════════════════════════════════════════════════════════════════
slide6 = prs.slides[5]

# Remove non-template shapes
remove_ids = []
for shape in slide6.shapes:
    name = shape.name
    is_template = (name in TEMPLATE_NAMES or
                   name.startswith('Rectangle') or
                   name.startswith('Title') or
                   name.startswith('Slide Number') or
                   name.startswith('Footer') or
                   name.startswith('Picture') or
                   name.startswith('Oval'))
    if not is_template:
        remove_ids.append(shape)

print('Slide 6: removing', len(remove_ids), 'content shapes')
for shape in remove_ids:
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

# ── Section headers ────────────────────────────────────────────────────────
add_textbox(slide6, Inches(0.45), Inches(1.15), Inches(5.5), Pt(18),
            "GOVERNMENT & REGULATORY SOURCES", 13, bold=True, underline=True, color=BLUE_MID)
add_textbox(slide6, Inches(6.15), Inches(1.15), Inches(5.5), Pt(18),
            "TECHNICAL & RESEARCH REFERENCES", 13, bold=True, underline=True, color=ACCENT_GREEN)

# ── References ─────────────────────────────────────────────────────────────
gov_refs = [
    ("1", "CAG Performance Audits of MPLADS (2005-2024)",
     "85 documented fraud/misuse cases worth Rs.32,133+ Cr across multiple parliamentary terms.",
     "https://www.cag.gov.in"),
    ("2", "MPLADS Guidelines 2023",
     "MoSPI official rules governing fund allocation, disbursement, and utilization.",
     "https://mplads.mospi.gov.in"),
    ("3", "eSAKSHI Portal",
     "Existing digital transaction tracking system for MPLADS works and payments.",
     "https://esakshi.mp.gov.in"),
    ("4", "data.gov.in MPLADS Dataset",
     "143,257 records of works and expenditures across India for ML training.",
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
     "A Unified Approach to Interpreting Model Predictions. NeurIPS 2017.",
     "https://proceedings.neurips.cc/paper/2017/hash/8a20a862197818fbd76f4f08898fc436-Abstract.html"),
    ("10", "Liu et al. (2012) - Isolation Forest",
     "Anomaly detection algorithm. IEEE TKDE.",
     "https://ieeexplore.ieee.org/document/6040193"),
    ("11", "RedFlags.ai (Kazakhstan)",
     "Saved $86M in 6 months using rule-based fraud detection.",
     "https://redflags.kz"),
    ("12", "Tazama (Linux Foundation)",
     "Open-source real-time payment monitoring at 2,300 TPS.",
     "https://tazama.io"),
    ("13", "UK HMRC / Quantexa Platform",
     "GBP 175M entity resolution for tax fraud detection.",
     "https://www.quantexa.com"),
    ("14", "RBI MuleHunter.AI",
     "85-90% accuracy rate with 26+ bank partnerships.",
     "https://www.rbi.org.in"),
    ("15", "Flower (flwr.ai)",
     "Production-ready federated learning framework.",
     "https://flower.ai"),
    ("16", "Nigrini (2012) - Benford's Law",
     "Application to financial fraud detection methodology.",
     "https://www.benfordslaw.org"),
]

y = Inches(1.4)
for num, title, desc, url in gov_refs:
    h = add_ref_card(slide6, Inches(0.45), y, Inches(5.5), num, title, desc, url, BLUE_MID)
    y += h + Pt(3)

y = Inches(1.4)
for num, title, desc, url in tech_refs:
    h = add_ref_card(slide6, Inches(6.15), y, Inches(5.5), num, title, desc, url, ACCENT_GREEN)
    y += h + Pt(3)

# ── Save ───────────────────────────────────────────────────────────────────
out = os.path.join(os.path.dirname(__file__), "ppt_v1_v2.pptx")
prs.save(out)
print("Saved:", out)
