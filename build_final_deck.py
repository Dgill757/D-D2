#!/usr/bin/env python3
"""
Build TheFuture_Final.pptx from TheFuture_Demov2.pptx
Comprehensive modification script.
"""

from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE

# ── Helper Functions ──────────────────────────────────────────────

def set_dark_bg(slide):
    """Set dark background #0F172A on a slide."""
    bg = slide.background
    fill = bg.fill
    fill.solid()
    fill.fore_color.rgb = RGBColor(0x0F, 0x17, 0x2A)


def add_text_box(slide, left, top, width, height, text,
                 font_size, bold=False, color='FFFFFF',
                 alignment=PP_ALIGN.CENTER, font_name='Calibri',
                 anchor=None):
    """Add a text box with consistent formatting. Returns the shape."""
    txBox = slide.shapes.add_textbox(
        Inches(left), Inches(top), Inches(width), Inches(height)
    )
    tf = txBox.text_frame
    tf.word_wrap = True
    if anchor is not None:
        tf.paragraphs[0].alignment = alignment  # set before anchor
        # Set anchor via XML attribute
        from lxml import etree
        nsmap = {'a': 'http://schemas.openxmlformats.org/drawingml/2006/main'}
        bodyPr = tf._txBody.find('.//a:bodyPr', nsmap)
        if bodyPr is not None:
            bodyPr.set('anchor', anchor)
    p = tf.paragraphs[0]
    p.alignment = alignment
    run = p.add_run()
    run.text = text
    run.font.size = Pt(font_size)
    run.font.bold = bold
    run.font.color.rgb = RGBColor.from_string(color)
    run.font.name = font_name
    return txBox


def add_card(slide, left, top, width, height):
    """Add a rounded rectangle card with consistent styling."""
    shape = slide.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE,
        Inches(left), Inches(top), Inches(width), Inches(height)
    )
    shape.fill.solid()
    shape.fill.fore_color.rgb = RGBColor(0x16, 0x28, 0x40)
    shape.line.fill.solid()
    shape.line.fill.fore_color.rgb = RGBColor(0xD9, 0x77, 0x06)
    shape.line.width = Emu(19050)
    # Set corner radius adjustment
    try:
        shape.adjustments[0] = 0.04  # ~adj=10000 out of 50000 scale
    except Exception:
        pass
    return shape


def add_multi_para_textbox(slide, left, top, width, height, paragraphs_data,
                           alignment=PP_ALIGN.CENTER, anchor_val=None):
    """
    Add a text box with multiple paragraphs.
    paragraphs_data: list of dicts with keys: text, size, bold, color, font_name (optional), alignment (optional)
    Returns the text box shape.
    """
    txBox = slide.shapes.add_textbox(
        Inches(left), Inches(top), Inches(width), Inches(height)
    )
    tf = txBox.text_frame
    tf.word_wrap = True

    if anchor_val is not None:
        from lxml import etree
        nsmap = {'a': 'http://schemas.openxmlformats.org/drawingml/2006/main'}
        bodyPr = tf._txBody.find('.//a:bodyPr', nsmap)
        if bodyPr is not None:
            bodyPr.set('anchor', anchor_val)

    for i, pdata in enumerate(paragraphs_data):
        if i == 0:
            p = tf.paragraphs[0]
        else:
            p = tf.add_paragraph()

        p.alignment = pdata.get('alignment', alignment)
        # Set space_before/space_after if provided
        if 'space_before' in pdata:
            p.space_before = pdata['space_before']
        if 'space_after' in pdata:
            p.space_after = pdata['space_after']

        run = p.add_run()
        run.text = pdata['text']
        run.font.size = Pt(pdata['size'])
        run.font.bold = pdata.get('bold', False)
        run.font.color.rgb = RGBColor.from_string(pdata.get('color', 'FFFFFF'))
        run.font.name = pdata.get('font_name', 'Calibri')

    return txBox


# ── Main Script ───────────────────────────────────────────────────

prs = Presentation('/home/user/D-D2/TheFuture_Demov2.pptx')

print(f"Starting with {len(prs.slides)} slides")

# ── TASK 1: Set dark backgrounds on ALL existing slides ───────────
for i, slide in enumerate(prs.slides):
    set_dark_bg(slide)
    print(f"  Set background on slide {i}")

# ── TASK 2: Enhance Slide 5 (index 4) ────────────────────────────
slide4 = prs.slides[4]

# Add stat banner between title and orbital layout
add_text_box(
    slide4,
    left=1.50, top=1.00, width=10.30, height=0.50,
    text="$262 Billion in claims denied annually — and it's only getting worse",
    font_size=20, bold=True, color='D97706',
    alignment=PP_ALIGN.CENTER
)

# Find and modify the "$50M annual costs" card
for shape in slide4.shapes:
    if shape.has_text_frame:
        full_text = shape.text_frame.text
        if "$50M" in full_text:
            print(f"  Found '$50M' card: '{full_text}'")
            # Clear existing text and set new text
            tf = shape.text_frame
            for p in tf.paragraphs:
                for run in p.runs:
                    if "$50M" in run.text or "annual" in run.text.lower():
                        pass  # We'll replace below
            # Replace the text - clear all paragraphs
            # Keep formatting from existing runs
            p = tf.paragraphs[0]
            # Clear existing runs by setting text
            for run in p.runs:
                run.text = ""
            # Set text on first run or add a new run
            if p.runs:
                p.runs[0].text = "💰  $5M+ lost per provider yearly"
            else:
                run = p.add_run()
                run.text = "💰  $5M+ lost per provider yearly"
                run.font.size = Pt(16)
                run.font.bold = True
                run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
                run.font.name = 'Calibri'
            print(f"  Updated card text to: '{tf.text}'")

print("Task 2 complete: Slide 5 enhanced")

# ── TASK 3: Create "$262 Billion Opportunity" slide ───────────────
blank_layout = prs.slide_layouts[6]  # Blank layout

slide_262b = prs.slides.add_slide(blank_layout)
set_dark_bg(slide_262b)

# Title
add_text_box(slide_262b, 0.50, 0.25, 12.33, 0.80,
             "The $262 Billion Opportunity", 42, bold=True, color='FFFFFF')

# Subtitle
add_text_box(slide_262b, 0.50, 0.95, 12.33, 0.50,
             "The Healthcare Administration Crisis", 20, bold=False, color='94A3B8')

# Hero stat
add_text_box(slide_262b, 2.00, 1.80, 9.33, 1.20,
             "$262B", 96, bold=True, color='D97706')

# Hero sub
add_text_box(slide_262b, 2.00, 3.10, 9.33, 0.50,
             "in claims denied annually from $3 trillion submitted",
             18, bold=False, color='FFFFFF')

# 4 stat cards in 2x2 grid
stat_cards = [
    # (left, top, stat, stat_color, label)
    (0.60, 3.80, "68M", "10B981", "Medicare Beneficiaries (2026)"),
    (7.10, 3.80, "85M", "10B981", "Medicaid Enrollees"),
    (0.60, 5.50, "$1.2T", "10B981", "Total Medicare Spending"),
    (7.10, 5.50, "10,000", "10B981", "Americans Turn 65 Every Day"),
]

for left, top, stat, stat_color, label in stat_cards:
    # Card background
    add_card(slide_262b, left, top, 5.50, 1.40)
    # Text box with two paragraphs overlaying the card
    add_multi_para_textbox(
        slide_262b, left + 0.10, top + 0.10, 5.30, 1.20,
        [
            {'text': stat, 'size': 36, 'bold': True, 'color': stat_color},
            {'text': label, 'size': 16, 'bold': False, 'color': 'FFFFFF'},
        ],
        alignment=PP_ALIGN.CENTER,
        anchor_val='ctr'
    )

# Bottom banner
add_text_box(slide_262b, 0.30, 6.80, 12.70, 0.50,
             "51% denial increase since 2021  •  CMS now mandating AI prior authorization pilots  •  10,000 aging into Medicare daily",
             14, bold=False, color='94A3B8')

print("Task 3 complete: $262B Opportunity slide created")

# ── TASK 4: Create "Why We Win" slide ─────────────────────────────
slide_win = prs.slides.add_slide(blank_layout)
set_dark_bg(slide_win)

# Title
add_text_box(slide_win, 0.50, 0.25, 12.33, 0.80,
             "Why We Win", 42, bold=True, color='FFFFFF')

# Subtitle
add_text_box(slide_win, 0.50, 0.95, 12.33, 0.50,
             "Built Different From Day One", 20, bold=False, color='94A3B8')

# 4 advantage cards in 2x2 grid
advantage_cards = [
    (0.50, 1.60, "95%+", "Automation Rate",
     "Competitors achieve 60-70%. Our agentic AI handles routine claims with zero human touch."),
    (6.80, 1.60, "48 Hours", "Implementation Time",
     "vs. 6-12 months with legacy systems. 500 charts to train, not 10,000+."),
    (0.50, 4.10, "65%", "Denial Reduction",
     "Plus 40% admin cost savings and 5+ day reduction in A/R — within 90 days."),
    (6.80, 4.10, "1st", "Unified Platform",
     "First to combine insurance operations + provider tools + member portal in a single ecosystem."),
]

for left, top, stat, heading, body in advantage_cards:
    # Card background
    add_card(slide_win, left, top, 5.80, 2.20)
    # Text box with 3 paragraphs overlaying the card
    add_multi_para_textbox(
        slide_win, left + 0.20, top + 0.15, 5.40, 1.90,
        [
            {'text': stat, 'size': 36, 'bold': True, 'color': '10B981'},
            {'text': heading, 'size': 20, 'bold': True, 'color': 'FFFFFF',
             'space_before': Pt(4)},
            {'text': body, 'size': 14, 'bold': False, 'color': '94A3B8',
             'space_before': Pt(4)},
        ],
        alignment=PP_ALIGN.LEFT,
        anchor_val='ctr'
    )

# Bottom tagline
add_text_box(slide_win, 0.30, 6.60, 12.70, 0.50,
             "Real-time CMS policy updates  •  Instant eligibility verification  •  Predictive denial prevention",
             14, bold=False, color='94A3B8')

print("Task 4 complete: Why We Win slide created")

# ── TASK 5: Create "The Revenue Engine" slide ─────────────────────
slide_rev = prs.slides.add_slide(blank_layout)
set_dark_bg(slide_rev)

# Title
add_text_box(slide_rev, 0.50, 0.25, 12.33, 0.80,
             "The Revenue Engine", 42, bold=True, color='FFFFFF')

# Subtitle
add_text_box(slide_rev, 0.50, 0.95, 12.33, 0.50,
             "Triple-Layer Revenue Model", 20, bold=False, color='94A3B8')

# 3 tall column cards
revenue_columns = [
    {
        'left': 0.40, 'top': 1.60,
        'icon': '💰', 'heading': 'Insurance Premiums',
        'price': '$200-300 PMPM',
        'bullets': [
            '▸  CMS capitation payments',
            '▸  State Medicaid contracts',
            '▸  Member premium collections',
            '▸  Risk adjustment optimization',
        ]
    },
    {
        'left': 4.70, 'top': 1.60,
        'icon': '⚡', 'heading': 'SaaS Platform',
        'price': '$50-150 /provider/mo',
        'bullets': [
            '▸  Claims automation platform',
            '▸  Prior authorization AI',
            '▸  Analytics & reporting suite',
            '▸  API access & integrations',
        ]
    },
    {
        'left': 9.00, 'top': 1.60,
        'icon': '📈', 'heading': 'Performance Revenue',
        'price': '15-25% of Savings',
        'bullets': [
            '▸  Denial reduction guarantees',
            '▸  Shared savings contracts',
            '▸  Value-based care bonuses',
            '▸  Quality metrics improvements',
        ]
    },
]

for col in revenue_columns:
    left = col['left']
    top = col['top']

    # Card background
    add_card(slide_rev, left, top, 3.80, 4.20)

    # Build paragraphs data
    paras = [
        {'text': col['icon'], 'size': 24, 'bold': False, 'color': 'FFFFFF',
         'alignment': PP_ALIGN.CENTER},
        {'text': col['heading'], 'size': 22, 'bold': True, 'color': 'FFFFFF',
         'alignment': PP_ALIGN.CENTER, 'space_before': Pt(6)},
        {'text': col['price'], 'size': 18, 'bold': True, 'color': 'D97706',
         'alignment': PP_ALIGN.CENTER, 'space_before': Pt(4)},
    ]
    for bullet in col['bullets']:
        paras.append({
            'text': bullet, 'size': 14, 'bold': False, 'color': 'FFFFFF',
            'alignment': PP_ALIGN.LEFT, 'space_before': Pt(4),
        })

    # Text box overlaying the card
    add_multi_para_textbox(
        slide_rev, left + 0.15, top + 0.20, 3.50, 3.80,
        paras,
        alignment=PP_ALIGN.CENTER,
        anchor_val='t'
    )

# Bottom banner
add_text_box(slide_rev, 0.30, 6.30, 12.70, 0.50,
             "70% Gross Margin  •  8-Month CAC Payback  •  Scalable Across All Verticals",
             16, bold=True, color='10B981')

print("Task 5 complete: Revenue Engine slide created")

# ── TASK 6: Reorder slides ────────────────────────────────────────
# Current indices after adding 3 slides (now 17 total):
#   0-13: original slides
#   14: $262B Opportunity (new)
#   15: Why We Win (new)
#   16: Revenue Engine (new)

new_order = [0, 1, 2, 3, 4, 14, 5, 6, 7, 8, 15, 9, 10, 11, 16, 12, 13]

print(f"\nReordering {len(prs.slides)} slides...")
print(f"New order: {new_order}")

xml_slides = prs.slides._sldIdLst
slides_list = list(xml_slides)

assert len(slides_list) == 17, f"Expected 17 slides, got {len(slides_list)}"
assert len(new_order) == 17, f"Expected 17 in new_order, got {len(new_order)}"

# Remove all
for s in slides_list:
    xml_slides.remove(s)

# Re-add in new order
for idx in new_order:
    xml_slides.append(slides_list[idx])

print("Task 6 complete: Slides reordered")

# ── Save ──────────────────────────────────────────────────────────
output_path = '/home/user/D-D2/TheFuture_Final.pptx'
prs.save(output_path)
print(f"\nSaved to {output_path}")
print(f"Final slide count: {len(prs.slides)}")

# Verify
prs2 = Presentation(output_path)
print(f"\nVerification - slide count: {len(prs2.slides)}")
for i, slide in enumerate(prs2.slides):
    texts = []
    for shape in slide.shapes:
        if shape.has_text_frame:
            for p in shape.text_frame.paragraphs:
                if p.text.strip():
                    texts.append(p.text.strip()[:60])
                    break
            if texts:
                break
    title = texts[0] if texts else "(no text)"
    print(f"  Slide {i}: {title}")

print("\nDone!")
