#!/usr/bin/env python3
"""
Lighthouse - Futuristic PowerPoint Presentation Builder
Summit Voice AI × QualifyOS Demo Deck
"""

import math
from datetime import date, timedelta
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE, MSO_SHAPE_TYPE
from pptx.enum.action import PP_ACTION
from pptx.oxml.ns import qn, nsmap
from lxml import etree
import copy

# ─── COLOR PALETTE ───────────────────────────────────────────────────────────
DEEP_NAVY    = RGBColor(0x0A, 0x16, 0x28)
NAVY_MID     = RGBColor(0x10, 0x20, 0x3C)
NAVY_LIGHT   = RGBColor(0x16, 0x2A, 0x50)
ELEC_TEAL    = RGBColor(0x06, 0xB6, 0xD4)
TEAL_DIM     = RGBColor(0x04, 0x7A, 0x8E)
TEAL_GLOW    = RGBColor(0x22, 0xD3, 0xEE)
NEON_ORANGE  = RGBColor(0xFF, 0x6B, 0x35)
ORANGE_DIM   = RGBColor(0xCC, 0x55, 0x2A)
WHITE        = RGBColor(0xFF, 0xFF, 0xFF)
OFF_WHITE    = RGBColor(0xE2, 0xE8, 0xF0)
GRAY_300     = RGBColor(0xCB, 0xD5, 0xE1)
GRAY_400     = RGBColor(0x94, 0xA3, 0xB8)
GRAY_500     = RGBColor(0x64, 0x74, 0x8B)
GRAY_700     = RGBColor(0x33, 0x41, 0x55)
DARK_BG      = RGBColor(0x05, 0x0B, 0x18)
GREEN_BRIGHT = RGBColor(0x10, 0xB9, 0x81)
RED_BRIGHT   = RGBColor(0xEF, 0x44, 0x44)
YELLOW_BRIGHT= RGBColor(0xF5, 0x9E, 0x0B)
PURPLE_BRIGHT= RGBColor(0x8B, 0x5C, 0xF6)

# Presentation dimensions (widescreen 16:9)
SLIDE_W = Inches(13.333)
SLIDE_H = Inches(7.5)

# Tomorrow's date
TOMORROW = date.today() + timedelta(days=1)
DATE_STR = TOMORROW.strftime("%B %d, %Y")

prs = Presentation()
prs.slide_width = SLIDE_W
prs.slide_height = SLIDE_H


# ═══════════════════════════════════════════════════════════════════════════════
# UTILITY FUNCTIONS
# ═══════════════════════════════════════════════════════════════════════════════

def set_slide_bg_gradient(slide, color1_hex, color2_hex):
    """Set a two-stop linear gradient background on a slide."""
    bg = slide.background
    fill = bg.fill
    fill.gradient()
    fill.gradient_stops[0].color.rgb = RGBColor.from_string(color1_hex)
    fill.gradient_stops[0].position = 0.0
    fill.gradient_stops[1].color.rgb = RGBColor.from_string(color2_hex)
    fill.gradient_stops[1].position = 1.0
    # Set angle to vertical (top to bottom)
    gf = fill._fill._element
    gf.set('rotWithShape', '0')


def set_slide_bg_solid(slide, color):
    """Set a solid color background on a slide."""
    bg = slide.background
    fill = bg.fill
    fill.solid()
    fill.fore_color.rgb = color


def add_textbox(slide, left, top, width, height, text, font_size=18,
                color=WHITE, bold=False, alignment=PP_ALIGN.LEFT,
                font_name="Inter", line_spacing=1.2):
    """Add a text box with styled text."""
    txBox = slide.shapes.add_textbox(left, top, width, height)
    tf = txBox.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = text
    p.font.size = Pt(font_size)
    p.font.color.rgb = color
    p.font.bold = bold
    p.font.name = font_name
    p.alignment = alignment
    p.space_after = Pt(0)
    p.space_before = Pt(0)
    # line spacing
    pPr = p._pPr
    if pPr is None:
        pPr = p._p.get_or_add_pPr()
    lnSpc = etree.SubElement(pPr, qn('a:lnSpc'))
    spc = etree.SubElement(lnSpc, qn('a:spcPct'))
    spc.set('val', str(int(line_spacing * 100000)))
    return txBox


def add_multiline_textbox(slide, left, top, width, height, lines,
                          font_name="Inter", line_spacing=1.3):
    """Add a text box with multiple styled lines.
    lines: list of dicts with keys: text, size, color, bold, alignment
    """
    txBox = slide.shapes.add_textbox(left, top, width, height)
    tf = txBox.text_frame
    tf.word_wrap = True
    for i, line_def in enumerate(lines):
        if i == 0:
            p = tf.paragraphs[0]
        else:
            p = tf.add_paragraph()
        p.text = line_def.get('text', '')
        p.font.size = Pt(line_def.get('size', 18))
        p.font.color.rgb = line_def.get('color', WHITE)
        p.font.bold = line_def.get('bold', False)
        p.font.name = line_def.get('font', font_name)
        p.alignment = line_def.get('alignment', PP_ALIGN.LEFT)
        sp_after = line_def.get('space_after', 4)
        p.space_after = Pt(sp_after)
        p.space_before = Pt(line_def.get('space_before', 0))
        # line spacing
        pPr = p._pPr
        if pPr is None:
            pPr = p._p.get_or_add_pPr()
        lnSpc = etree.SubElement(pPr, qn('a:lnSpc'))
        spc = etree.SubElement(lnSpc, qn('a:spcPct'))
        spc.set('val', str(int(line_spacing * 100000)))
    return txBox


def add_rounded_rect(slide, left, top, width, height, fill_color,
                     border_color=None, border_width=Pt(1), corner_radius=None):
    """Add a rounded rectangle shape."""
    shape = slide.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE, left, top, width, height
    )
    shape.fill.solid()
    shape.fill.fore_color.rgb = fill_color
    if border_color:
        shape.line.color.rgb = border_color
        shape.line.width = border_width
    else:
        shape.line.fill.background()
    return shape


def add_circle(slide, left, top, size, fill_color, border_color=None, border_width=Pt(1)):
    """Add a circle / oval."""
    shape = slide.shapes.add_shape(
        MSO_SHAPE.OVAL, left, top, size, size
    )
    shape.fill.solid()
    shape.fill.fore_color.rgb = fill_color
    if border_color:
        shape.line.color.rgb = border_color
        shape.line.width = border_width
    else:
        shape.line.fill.background()
    return shape


def set_shape_glow(shape, color, radius=Pt(8), alpha=40000):
    """Add a glow effect to a shape via XML manipulation."""
    spPr = shape._element.spPr
    # Create effectLst if not present
    effectLst = spPr.find(qn('a:effectLst'))
    if effectLst is None:
        effectLst = etree.SubElement(spPr, qn('a:effectLst'))
    glow = etree.SubElement(effectLst, qn('a:glow'))
    glow.set('rad', str(int(radius)))
    srgbClr = etree.SubElement(glow, qn('a:srgbClr'))
    srgbClr.set('val', '{:02X}{:02X}{:02X}'.format(color[0], color[1], color[2]))
    alphaEl = etree.SubElement(srgbClr, qn('a:alpha'))
    alphaEl.set('val', str(alpha))


def add_shape_text(shape, text, font_size=14, color=WHITE, bold=False,
                   alignment=PP_ALIGN.CENTER, font_name="Inter"):
    """Add text to a shape."""
    tf = shape.text_frame
    tf.word_wrap = True
    tf.paragraphs[0].alignment = alignment
    p = tf.paragraphs[0]
    p.text = text
    p.font.size = Pt(font_size)
    p.font.color.rgb = color
    p.font.bold = bold
    p.font.name = font_name
    return tf


def add_multitext_shape(shape, lines, font_name="Inter"):
    """Add multiple styled text lines to a shape.
    lines: list of (text, size, color, bold)
    """
    tf = shape.text_frame
    tf.word_wrap = True
    for i, (text, size, color, bold) in enumerate(lines):
        if i == 0:
            p = tf.paragraphs[0]
        else:
            p = tf.add_paragraph()
        p.text = text
        p.font.size = Pt(size)
        p.font.color.rgb = color
        p.font.bold = bold
        p.font.name = font_name
        p.alignment = PP_ALIGN.CENTER
        p.space_after = Pt(2)


def add_progress_dots(slide, current, total=10):
    """Add progress indicator dots at bottom of slide."""
    dot_size = Inches(0.12)
    gap = Inches(0.28)
    total_w = total * gap
    start_x = (SLIDE_W - total_w) / 2
    y = SLIDE_H - Inches(0.35)
    for i in range(total):
        x = start_x + i * gap
        if i == current:
            c = ELEC_TEAL
            sz = Inches(0.14)
        else:
            c = GRAY_700
            sz = dot_size
        dot = slide.shapes.add_shape(
            MSO_SHAPE.OVAL, x, y, sz, sz
        )
        dot.fill.solid()
        dot.fill.fore_color.rgb = c
        dot.line.fill.background()


def add_home_button(slide, target_slide_idx=0):
    """Add a subtle home navigation icon in the top-right corner."""
    btn = slide.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE,
        SLIDE_W - Inches(0.7), Inches(0.15), Inches(0.5), Inches(0.35)
    )
    btn.fill.solid()
    btn.fill.fore_color.rgb = NAVY_LIGHT
    btn.line.fill.background()
    tf = btn.text_frame
    tf.word_wrap = False
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    run = p.add_run()
    run.text = "⌂"
    run.font.size = Pt(16)
    run.font.color.rgb = GRAY_400
    run.font.name = "Inter"
    # Add hyperlink to slide 1
    try:
        click = run._r.get_or_add_hlinkClick()
        rId = slide.part.relate_to(
            prs.slides[target_slide_idx].part,
            'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide'
        )
        click.set(qn('r:id'), rId)
        click.set('action', 'ppaction://hlinksldjump')
    except:
        pass


def add_particle_dots(slide, count=35):
    """Add subtle floating particle dots for ambiance."""
    import random
    random.seed(42)  # reproducible
    for _ in range(count):
        x = random.randint(0, int(SLIDE_W) - 50000)
        y = random.randint(0, int(SLIDE_H) - 50000)
        size = random.randint(15000, 55000)  # EMU: small dots
        opacity_val = random.randint(8, 25)
        dot = slide.shapes.add_shape(MSO_SHAPE.OVAL, x, y, size, size)
        dot.fill.solid()
        # Random choice between teal and white particles
        if random.random() < 0.6:
            dot.fill.fore_color.rgb = ELEC_TEAL
        else:
            dot.fill.fore_color.rgb = WHITE
        dot.line.fill.background()
        # Set opacity via XML
        solidFill = dot._element.spPr.find(qn('a:solidFill'))
        if solidFill is not None:
            srgb = solidFill.find(qn('a:srgbClr'))
            if srgb is not None:
                alpha = etree.SubElement(srgb, qn('a:alpha'))
                alpha.set('val', str(opacity_val * 1000))


def add_subtle_line(slide, x1, y1, x2, y2, color=GRAY_700, width=Pt(1)):
    """Add a subtle connecting line."""
    connector = slide.shapes.add_connector(
        1,  # straight connector
        x1, y1, x2, y2
    )
    connector.line.color.rgb = color
    connector.line.width = width
    # Set line opacity
    ln = connector._element.find(qn('a:ln'), connector._element.nsmap) or \
         connector._element.spPr.find(qn('a:ln'))
    if ln is not None:
        solidFill = ln.find(qn('a:solidFill'))
        if solidFill is not None:
            srgb = solidFill.find(qn('a:srgbClr'))
            if srgb is not None:
                alpha = etree.SubElement(srgb, qn('a:alpha'))
                alpha.set('val', '50000')


def add_chevron_arrow(slide, left, top, width, height, color):
    """Add a right-pointing chevron/arrow."""
    shape = slide.shapes.add_shape(
        MSO_SHAPE.CHEVRON, left, top, width, height
    )
    shape.fill.solid()
    shape.fill.fore_color.rgb = color
    shape.line.fill.background()
    return shape


def add_notes(slide, text):
    """Add presenter notes to a slide."""
    notes_slide = slide.notes_slide
    notes_slide.notes_text_frame.text = text


def set_shape_opacity(shape, opacity_pct):
    """Set a shape's fill opacity (0-100)."""
    solidFill = shape._element.spPr.find(qn('a:solidFill'))
    if solidFill is None:
        return
    color_elem = solidFill[0] if len(solidFill) > 0 else None
    if color_elem is not None:
        alpha = etree.SubElement(color_elem, qn('a:alpha'))
        alpha.set('val', str(int(opacity_pct * 1000)))


# ═══════════════════════════════════════════════════════════════════════════════
# SLIDE 1: SIMPLE INTRO
# ═══════════════════════════════════════════════════════════════════════════════

def build_slide_1():
    slide = prs.slides.add_slide(prs.slide_layouts[6])  # blank layout
    set_slide_bg_gradient(slide, '050B18', '0A1628')

    # Floating particles
    add_particle_dots(slide, 40)

    # Subtle top accent line
    line_shape = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE,
        Inches(4.5), Inches(1.8), Inches(4.3), Pt(2)
    )
    line_shape.fill.solid()
    line_shape.fill.fore_color.rgb = ELEC_TEAL
    line_shape.line.fill.background()
    set_shape_opacity(line_shape, 60)

    # Main title
    add_textbox(slide,
        Inches(1.5), Inches(2.1), Inches(10.3), Inches(1.2),
        "What We've Been Building",
        font_size=58, color=WHITE, bold=True, alignment=PP_ALIGN.CENTER,
        font_name="Inter"
    )

    # Subtitle with teal accent
    add_textbox(slide,
        Inches(2.5), Inches(3.4), Inches(8.3), Inches(0.7),
        "Summit Voice AI  ×  Lighthouse",
        font_size=34, color=ELEC_TEAL, bold=False, alignment=PP_ALIGN.CENTER,
        font_name="Inter"
    )

    # Names
    add_textbox(slide,
        Inches(3.5), Inches(4.5), Inches(6.3), Inches(0.5),
        "Dan Gill  &  David Newsom",
        font_size=22, color=GRAY_400, bold=False, alignment=PP_ALIGN.CENTER,
        font_name="Inter"
    )

    # Date
    add_textbox(slide,
        Inches(4.5), Inches(5.1), Inches(4.3), Inches(0.4),
        DATE_STR,
        font_size=16, color=GRAY_500, bold=False, alignment=PP_ALIGN.CENTER,
        font_name="Inter"
    )

    # Bottom accent line
    line_shape2 = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE,
        Inches(4.5), Inches(5.7), Inches(4.3), Pt(2)
    )
    line_shape2.fill.solid()
    line_shape2.fill.fore_color.rgb = ELEC_TEAL
    line_shape2.line.fill.background()
    set_shape_opacity(line_shape2, 40)

    # Lighthouse branding - small subtle text at very bottom
    add_textbox(slide,
        Inches(5), Inches(6.5), Inches(3.3), Inches(0.3),
        "LIGHTHOUSE",
        font_size=11, color=GRAY_700, bold=True, alignment=PP_ALIGN.CENTER,
        font_name="Inter"
    )

    add_progress_dots(slide, 0)

    add_notes(slide, """SLIDE 1 - INTRO
[TIMING: ~30 seconds]

Keep this simple. Don't oversell.

SAY: "Thanks for making time. We want to show you what we've been building together. This isn't a pitch deck - this is more of a 'hey, look at this cool thing' kind of conversation."

CLICK: Advance to next slide when ready.
""")

    return slide


# ═══════════════════════════════════════════════════════════════════════════════
# SLIDE 2: THE REAL STORY (TIMELINE)
# ═══════════════════════════════════════════════════════════════════════════════

def build_slide_2():
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_slide_bg_gradient(slide, '050B18', '0A1628')
    add_particle_dots(slide, 15)

    # Title
    add_textbox(slide,
        Inches(0.6), Inches(0.3), Inches(5), Inches(0.7),
        "The Real Story",
        font_size=42, color=WHITE, bold=True, alignment=PP_ALIGN.LEFT
    )
    # Subtitle
    add_textbox(slide,
        Inches(0.6), Inches(0.85), Inches(6), Inches(0.4),
        "How we got here",
        font_size=18, color=GRAY_400, bold=False, alignment=PP_ALIGN.LEFT
    )

    # Timeline data
    timeline = [
        ("2020", "🚀", "Started Summit Marketing Group", "Agency roots, learned the game", ELEC_TEAL),
        ("2022", "🧠", "Built first Voice AI systems", "The hard way — trial by fire", ELEC_TEAL),
        ("2023", "🏥", "Pitched Aetna on Voice AI", "Healthcare was always the vision", TEAL_GLOW),
        ("2024", "🏠", "Roofing opportunity → QualifyOS", "Proved the tech in the real world", NEON_ORANGE),
        ("2025", "🤝", "Connected with David", "M/W/F 9:30 PM grind sessions", NEON_ORANGE),
        ("Now", "⚡", "Building Lighthouse", "What insurance should've been all along", NEON_ORANGE),
    ]

    # Vertical timeline
    line_x = Inches(1.6)
    start_y = Inches(1.7)
    step_y = Inches(0.88)

    # Draw vertical timeline line
    tl_line = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE,
        line_x + Inches(0.18), start_y, Pt(2), step_y * (len(timeline) - 1) + Inches(0.3)
    )
    tl_line.fill.solid()
    tl_line.fill.fore_color.rgb = GRAY_700
    tl_line.line.fill.background()

    for i, (year, icon, title, desc, color) in enumerate(timeline):
        y = start_y + i * step_y

        # Timeline node (circle)
        node = add_circle(slide, line_x, y, Inches(0.38), NAVY_LIGHT, color, Pt(2))
        set_shape_glow(node, (color[0], color[1], color[2]) if isinstance(color, tuple) else
                       (int(str(color)[:2], 16) if isinstance(color, str) else color[0],
                        color[1] if hasattr(color, '__getitem__') else 0,
                        color[2] if hasattr(color, '__getitem__') else 0),
                       radius=Pt(6), alpha=30000)

        # Icon on node
        add_shape_text(node, icon, font_size=14, color=WHITE)

        # Year label (left of timeline)
        add_textbox(slide,
            Inches(0.2), y + Inches(0.02), Inches(1.2), Inches(0.35),
            year, font_size=16, color=color, bold=True, alignment=PP_ALIGN.RIGHT
        )

        # Title (right of timeline)
        add_textbox(slide,
            Inches(2.2), y - Inches(0.02), Inches(4.2), Inches(0.3),
            title, font_size=17, color=WHITE, bold=True, alignment=PP_ALIGN.LEFT
        )

        # Description
        add_textbox(slide,
            Inches(2.2), y + Inches(0.26), Inches(4.2), Inches(0.25),
            desc, font_size=13, color=GRAY_400, bold=False, alignment=PP_ALIGN.LEFT
        )

    # Right side: abstract decorative element - "lighthouse beam" effect
    # Large teal glow circle (abstract)
    glow_circle = add_circle(slide, Inches(8.5), Inches(1.5), Inches(4.5),
                             DEEP_NAVY, ELEC_TEAL, Pt(1))
    set_shape_opacity(glow_circle, 8)
    set_shape_glow(glow_circle, (0x06, 0xB6, 0xD4), radius=Pt(20), alpha=15000)

    # Decorative text on right
    add_textbox(slide,
        Inches(8.2), Inches(3.0), Inches(4.5), Inches(1.5),
        "From agency hustle\nto AI infrastructure\nfor healthcare",
        font_size=22, color=GRAY_500, bold=False, alignment=PP_ALIGN.CENTER,
        line_spacing=1.5
    )

    # Small accent: "5 years in the making"
    accent_box = add_rounded_rect(slide, Inches(9.0), Inches(5.0), Inches(2.8), Inches(0.5),
                                   NAVY_LIGHT, ELEC_TEAL, Pt(1))
    add_shape_text(accent_box, "5 years in the making", font_size=13, color=ELEC_TEAL, bold=True)

    add_progress_dots(slide, 1)
    add_home_button(slide)

    add_notes(slide, """SLIDE 2 - THE REAL STORY (TIMELINE)
[TIMING: 2-3 minutes]

Walk through the timeline naturally, like telling a story to a friend.

SAY: "So here's how we got here. It wasn't planned - it evolved."

FOR EACH NODE:
- 2020: "Started Summit as a marketing agency. Learned how businesses actually work."
- 2022: "Got obsessed with Voice AI. Built systems the hard way - no shortcuts."
- 2023: "Pitched Aetna. They loved the idea but timing wasn't right. Healthcare stuck with us though."
- 2024: "Roofing opportunity came along. Built QualifyOS to prove the stack works."
- 2025: "Met David. Started our M/W/F 9:30 PM sessions. That's when everything clicked."
- Now: "Building Lighthouse - what insurance should have been all along."

KEY POINT: Emphasize this wasn't a pivot - it was an evolution. Each step built on the last.
""")

    return slide


# ═══════════════════════════════════════════════════════════════════════════════
# SLIDE 3: WHERE WE STARTED - ROOFING
# ═══════════════════════════════════════════════════════════════════════════════

def build_slide_3():
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_slide_bg_gradient(slide, '050B18', '0A1628')
    add_particle_dots(slide, 10)

    # ─── LEFT SIDE (40%) ───
    add_textbox(slide,
        Inches(0.6), Inches(0.3), Inches(5), Inches(0.7),
        "Where We Started",
        font_size=42, color=WHITE, bold=True
    )
    add_textbox(slide,
        Inches(0.6), Inches(0.85), Inches(5), Inches(0.4),
        "Why roofing first?",
        font_size=18, color=NEON_ORANGE, bold=False
    )

    bullets = [
        ("Biggest ROI in home services", "Every storm = massive demand"),
        ("Clear pain points", "Missed calls = lost revenue, period"),
        ("24/7 Voice AI receptionist", "Never miss another call"),
        ("Backend automation stack", "From lead to close, fully automated"),
    ]

    for i, (title, desc) in enumerate(bullets):
        y = Inches(1.7) + i * Inches(1.05)

        # Bullet accent bar
        bar = slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE,
            Inches(0.6), y, Pt(3), Inches(0.6)
        )
        bar.fill.solid()
        bar.fill.fore_color.rgb = ELEC_TEAL
        bar.line.fill.background()

        add_textbox(slide, Inches(0.9), y, Inches(4.5), Inches(0.35),
                    title, font_size=18, color=WHITE, bold=True)
        add_textbox(slide, Inches(0.9), y + Inches(0.3), Inches(4.5), Inches(0.3),
                    desc, font_size=14, color=GRAY_400)

    # Proof label
    add_textbox(slide,
        Inches(0.6), Inches(5.9), Inches(4), Inches(0.3),
        "This proved the tech stack works.",
        font_size=14, color=GRAY_500, bold=False
    )

    # ─── RIGHT SIDE (60%) - Metrics ───
    # Dark card background
    metrics_bg = add_rounded_rect(slide, Inches(5.8), Inches(0.8), Inches(7), Inches(6.0),
                                   NAVY_MID, GRAY_700, Pt(1))
    set_shape_opacity(metrics_bg, 60)

    # Metrics header
    add_textbox(slide,
        Inches(6.2), Inches(1.0), Inches(6), Inches(0.5),
        "BEFORE → AFTER",
        font_size=14, color=GRAY_400, bold=True, alignment=PP_ALIGN.CENTER
    )

    metrics = [
        ("Missed Calls", "40%", "0%", RED_BRIGHT, GREEN_BRIGHT),
        ("Response Time", "4 hours", "Instant", YELLOW_BRIGHT, GREEN_BRIGHT),
        ("Conversion Rate", "12%", "34%", YELLOW_BRIGHT, GREEN_BRIGHT),
    ]

    for i, (label, before, after, color_before, color_after) in enumerate(metrics):
        y = Inches(1.8) + i * Inches(1.5)
        card_x = Inches(6.2)

        # Metric label
        add_textbox(slide, card_x, y, Inches(6), Inches(0.35),
                    label, font_size=16, color=GRAY_300, bold=True, alignment=PP_ALIGN.CENTER)

        # Before card
        before_card = add_rounded_rect(slide, card_x + Inches(0.3), y + Inches(0.4),
                                        Inches(2.3), Inches(0.8), NAVY_LIGHT, color_before, Pt(1))
        add_multitext_shape(before_card, [
            ("BEFORE", 10, GRAY_500, False),
            (before, 28, color_before, True),
        ])

        # Arrow
        add_textbox(slide,
            card_x + Inches(2.7), y + Inches(0.5), Inches(0.8), Inches(0.5),
            "→", font_size=28, color=GRAY_500, bold=True, alignment=PP_ALIGN.CENTER
        )

        # After card
        after_card = add_rounded_rect(slide, card_x + Inches(3.5), y + Inches(0.4),
                                       Inches(2.3), Inches(0.8), NAVY_LIGHT, color_after, Pt(1))
        set_shape_glow(after_card, (color_after[0], color_after[1], color_after[2]),
                       radius=Pt(6), alpha=25000)
        add_multitext_shape(after_card, [
            ("AFTER", 10, GRAY_500, False),
            (after, 28, color_after, True),
        ])

    # QualifyOS badge
    badge = add_rounded_rect(slide, Inches(7.8), Inches(6.0), Inches(3.2), Inches(0.45),
                              NAVY_LIGHT, NEON_ORANGE, Pt(1))
    add_shape_text(badge, "Powered by QualifyOS", font_size=13, color=NEON_ORANGE, bold=True)

    add_progress_dots(slide, 2)
    add_home_button(slide)

    add_notes(slide, """SLIDE 3 - WHERE WE STARTED (ROOFING)
[TIMING: 1-2 minutes]

SAY: "Before healthcare, we proved everything in roofing. Why? Because it's the highest-ROI home service vertical and the pain points are crystal clear."

WALK THROUGH BULLETS:
- "Every storm creates a surge. Roofers who answer first, win."
- "We built a 24/7 Voice AI receptionist. Zero missed calls."
- "The whole backend is automated - from lead capture to follow-up."

POINT TO METRICS:
- "Look at the numbers. 40% missed calls to zero. 4-hour response time to instant. Conversion nearly tripled."
- "This is what gave us confidence the stack works."

TRANSITION: "So we'd proven the tech. But healthcare was always where we wanted to be..."
""")

    return slide


# ═══════════════════════════════════════════════════════════════════════════════
# SLIDE 4: WHAT WE SAW IN HEALTHCARE
# ═══════════════════════════════════════════════════════════════════════════════

def build_slide_4():
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_slide_bg_gradient(slide, '050B18', '0A1628')

    # Title
    add_textbox(slide,
        Inches(0.6), Inches(0.3), Inches(6), Inches(0.7),
        "What We Saw in Healthcare",
        font_size=42, color=WHITE, bold=True
    )
    add_textbox(slide,
        Inches(0.6), Inches(0.85), Inches(8), Inches(0.4),
        "Real problems. Real people. Nobody listening.",
        font_size=18, color=GRAY_400
    )

    # Center frustrated person icon (large circle)
    center_x = Inches(5.5)
    center_y = Inches(3.2)
    center_circle = add_circle(slide, center_x, center_y, Inches(1.5),
                                NAVY_LIGHT, NEON_ORANGE, Pt(2))
    set_shape_glow(center_circle, (0xFF, 0x6B, 0x35), radius=Pt(12), alpha=20000)
    add_shape_text(center_circle, "😤\nMembers", font_size=18, color=WHITE, bold=True)

    # Problem bubbles - 4 corners around center
    problems = [
        (Inches(3.2), Inches(1.5), "⏱️  45+ min wait times", "Average hold time for\nmember support calls", ELEC_TEAL),
        (Inches(8.0), Inches(2.2), "💰  $50M+ annual costs", "Contact center spend\nfor large insurers", NEON_ORANGE),
        (Inches(8.0), Inches(4.8), "📱  Reddit & Facebook\n      complaints piling up", "Members venting online\nbecause no one listens", RED_BRIGHT),
        (Inches(3.2), Inches(4.8), "🔄  IVR mazes\n      making it worse", "Press 1, press 2, press\nhang up in frustration", YELLOW_BRIGHT),
    ]

    for x, y, title, desc, accent_color in problems:
        bubble = add_rounded_rect(slide, x, y, Inches(3.2), Inches(1.15),
                                   NAVY_MID, accent_color, Pt(1))
        tf = bubble.text_frame
        tf.word_wrap = True
        tf.margin_left = Inches(0.15)
        tf.margin_right = Inches(0.15)
        tf.margin_top = Inches(0.1)

        p1 = tf.paragraphs[0]
        p1.text = title
        p1.font.size = Pt(15)
        p1.font.color.rgb = WHITE
        p1.font.bold = True
        p1.font.name = "Inter"
        p1.space_after = Pt(4)

        p2 = tf.add_paragraph()
        p2.text = desc
        p2.font.size = Pt(11)
        p2.font.color.rgb = GRAY_400
        p2.font.name = "Inter"

    # Bottom callout
    callout_bg = add_rounded_rect(slide, Inches(3.5), Inches(6.3), Inches(6.3), Inches(0.6),
                                   NAVY_MID, ELEC_TEAL, Pt(1))
    set_shape_glow(callout_bg, (0x06, 0xB6, 0xD4), radius=Pt(8), alpha=20000)
    add_shape_text(callout_bg, "What if someone actually listened?",
                   font_size=20, color=ELEC_TEAL, bold=True)

    add_progress_dots(slide, 3)
    add_home_button(slide)

    add_notes(slide, """SLIDE 4 - WHAT WE SAW IN HEALTHCARE
[TIMING: 2 minutes]

SAY: "Here's what we kept seeing in healthcare."

WALK THROUGH EACH BUBBLE (they appear one by one):
1. "45-minute average wait times. That's just... broken."
2. "$50 million a year in contact center costs for large insurers."
3. "Meanwhile, members are on Reddit and Facebook complaining because nobody's listening."
4. "And the IVR systems? They make it worse. Press 1, press 2, press give up."

PAUSE, then:
SAY: "So we asked a simple question..." [point to bottom callout]
"What if someone actually listened?"

TRANSITION: "That's when we had the idea that changed everything..."
""")

    return slide


# ═══════════════════════════════════════════════════════════════════════════════
# SLIDE 5: THE IDEA - FIND THEM FIRST
# ═══════════════════════════════════════════════════════════════════════════════

def build_slide_5():
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_slide_bg_gradient(slide, '050B18', '0A1628')

    # Title
    add_textbox(slide,
        Inches(0.6), Inches(0.3), Inches(8), Inches(0.7),
        "The Idea: Find Them First",
        font_size=42, color=WHITE, bold=True
    )
    add_textbox(slide,
        Inches(0.6), Inches(0.85), Inches(10), Inches(0.4),
        "Don't wait for members to call. Find them where they're already struggling.",
        font_size=18, color=GRAY_400
    )

    # ─── FLOW DIAGRAM: 4 main nodes ───
    nodes = [
        ("🔍", "Discovery", "Find struggling\nmembers online", ELEC_TEAL),
        ("🧠", "Intelligence", "Analyze & enrich\ncontact data", PURPLE_BRIGHT),
        ("⚡", "Response", "Reach out via\nright channel", NEON_ORANGE),
        ("✅", "Resolution", "Solve their\nproblem", GREEN_BRIGHT),
    ]

    node_w = Inches(2.4)
    node_h = Inches(1.6)
    start_x = Inches(0.8)
    gap = Inches(0.7)
    y = Inches(2.0)

    for i, (icon, title, desc, color) in enumerate(nodes):
        x = start_x + i * (node_w + gap)

        # Node card
        card = add_rounded_rect(slide, x, y, node_w, node_h, NAVY_MID, color, Pt(2))
        set_shape_glow(card, (color[0], color[1], color[2]),
                       radius=Pt(8), alpha=25000)

        tf = card.text_frame
        tf.word_wrap = True
        tf.margin_left = Inches(0.15)
        tf.margin_top = Inches(0.12)

        p = tf.paragraphs[0]
        p.text = icon
        p.font.size = Pt(28)
        p.alignment = PP_ALIGN.CENTER
        p.space_after = Pt(4)

        p2 = tf.add_paragraph()
        p2.text = title
        p2.font.size = Pt(20)
        p2.font.color.rgb = color
        p2.font.bold = True
        p2.font.name = "Inter"
        p2.alignment = PP_ALIGN.CENTER
        p2.space_after = Pt(6)

        p3 = tf.add_paragraph()
        p3.text = desc
        p3.font.size = Pt(12)
        p3.font.color.rgb = GRAY_400
        p3.font.name = "Inter"
        p3.alignment = PP_ALIGN.CENTER

        # Arrow between nodes (except last)
        if i < len(nodes) - 1:
            arrow_x = x + node_w + Inches(0.15)
            add_textbox(slide, arrow_x, y + Inches(0.5), Inches(0.4), Inches(0.5),
                        "→", font_size=28, color=GRAY_500, bold=True, alignment=PP_ALIGN.CENTER)

    # ─── DETAIL PANELS below each node ───
    details = [
        [  # Discovery
            ("Reddit", "r/insurance complaints"),
            ("Facebook", "Medicare help groups"),
            ("LinkedIn", "Healthcare discussions"),
            ("Blogs", "Comment sections"),
        ],
        [  # Intelligence
            ("Extract", "Pain point analysis"),
            ("Enrich", "Contact information"),
            ("Score", "Urgency level (1-10)"),
            ("Route", "Smart channel routing"),
        ],
        [  # Response
            ("📞 Call", "High priority (8-10)"),
            ("💬 SMS", "Medium priority (5-7)"),
            ("📧 Email", "Lower priority (1-4)"),
        ],
        [  # Resolution
            ("Dialogue", "Natural conversation"),
            ("Options", "All insurance plans"),
            ("Match", "Best fit selection"),
            ("Connect", "Provider handoff"),
        ],
    ]

    for i, detail_items in enumerate(details):
        x = start_x + i * (node_w + gap)
        detail_y = Inches(4.0)

        for j, (label, desc) in enumerate(detail_items):
            item_y = detail_y + j * Inches(0.65)
            color = nodes[i][3]

            # Small dot
            dot = add_circle(slide, x + Inches(0.15), item_y + Inches(0.08), Inches(0.15),
                            color, border_color=None)

            # Detail text
            add_textbox(slide, x + Inches(0.4), item_y - Inches(0.02), Inches(2.0), Inches(0.22),
                        label, font_size=13, color=WHITE, bold=True)
            add_textbox(slide, x + Inches(0.4), item_y + Inches(0.18), Inches(2.0), Inches(0.22),
                        desc, font_size=11, color=GRAY_500)

    # Bottom tagline
    add_textbox(slide,
        Inches(2), Inches(6.7), Inches(9.3), Inches(0.4),
        "Click any node above for drill-down detail  ·  All running autonomously by 6 AM every day",
        font_size=13, color=GRAY_500, alignment=PP_ALIGN.CENTER
    )

    add_progress_dots(slide, 4)
    add_home_button(slide)

    add_notes(slide, """SLIDE 5 - THE IDEA: FIND THEM FIRST
[TIMING: 2-3 minutes]

This is the BIG IDEA slide. Take your time here.

SAY: "Here's the core insight: don't wait for members to call. Find them where they're already struggling."

WALK THROUGH THE FLOW:
1. DISCOVERY: "Every morning at 6 AM, our system scans Reddit, Facebook, LinkedIn, blog comments - anywhere people are talking about insurance problems."

2. INTELLIGENCE: "AI analyzes each post, extracts the pain point, finds contact info, and scores urgency on a 1-10 scale."

3. RESPONSE: "High urgency? Phone call. Medium? SMS. Lower? Email with resources. Right channel, right time."

4. RESOLUTION: "When we connect, it's a real conversation. Not a sales pitch. We help them understand ALL their options and find the best fit."

CLICK: You can click each node to show the detail panel below it.

KEY POINT: "The member never had to call anyone. We found them and helped before they even knew to ask."
""")

    return slide


# ═══════════════════════════════════════════════════════════════════════════════
# SLIDE 6: HOW IT ACTUALLY WORKS
# ═══════════════════════════════════════════════════════════════════════════════

def build_slide_6():
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_slide_bg_gradient(slide, '050B18', '0A1628')

    # Title
    add_textbox(slide,
        Inches(0.6), Inches(0.2), Inches(8), Inches(0.6),
        "How It Actually Works",
        font_size=42, color=WHITE, bold=True
    )
    add_textbox(slide,
        Inches(0.6), Inches(0.75), Inches(10), Inches(0.3),
        "A day in the life of Lighthouse",
        font_size=16, color=GRAY_400
    )

    phases = [
        ("6 AM", "Discovery Phase", "🔍", ELEC_TEAL,
         ["Reddit: insurance complaint posts", "Facebook: Medicare group discussions",
          "LinkedIn: healthcare pain threads", "Blog comments: frustrated members"]),
        ("7 AM", "Intelligence Phase", "🧠", PURPLE_BRIGHT,
         ["Extract core struggles & sentiment", "Find + verify contact information",
          "Score urgency: 🟢 Low  🟡 Med  🔴 High", "Route to optimal response channel"]),
        ("8 AM", "Response Phase", "⚡", NEON_ORANGE,
         ["🔴 High priority → Voice AI call", "🟡 Medium → Personalized SMS",
          "🟢 Lower → Email with resources"]),
        ("All Day", "Resolution Phase", "✅", GREEN_BRIGHT,
         ["Natural Voice AI conversation", "Access to ALL insurance options",
          "Best-fit matching & recommendation", "Provider connection + follow-up"]),
    ]

    card_w = Inches(5.5)
    card_h = Inches(1.3)
    start_x = Inches(0.6)
    start_y = Inches(1.35)
    gap_y = Inches(1.42)

    for i, (time, title, icon, color, items) in enumerate(phases):
        y = start_y + i * gap_y

        # Time badge
        time_badge = add_rounded_rect(slide, start_x, y + Inches(0.1),
                                       Inches(1.1), Inches(0.45), color)
        add_shape_text(time_badge, time, font_size=14, color=DEEP_NAVY, bold=True)

        # Phase card
        card = add_rounded_rect(slide, start_x + Inches(1.3), y,
                                 card_w, card_h, NAVY_MID, color, Pt(1))

        tf = card.text_frame
        tf.word_wrap = True
        tf.margin_left = Inches(0.15)
        tf.margin_top = Inches(0.08)

        p = tf.paragraphs[0]
        p.text = f"{icon}  {title}"
        p.font.size = Pt(18)
        p.font.color.rgb = color
        p.font.bold = True
        p.font.name = "Inter"
        p.space_after = Pt(6)

        for item in items:
            pi = tf.add_paragraph()
            pi.text = f"    {item}"
            pi.font.size = Pt(12)
            pi.font.color.rgb = GRAY_300
            pi.font.name = "Inter"
            pi.space_after = Pt(2)

        # Connecting line to next phase
        if i < len(phases) - 1:
            conn_line = slide.shapes.add_shape(
                MSO_SHAPE.RECTANGLE,
                start_x + Inches(0.5), y + card_h,
                Pt(2), gap_y - card_h
            )
            conn_line.fill.solid()
            conn_line.fill.fore_color.rgb = GRAY_700
            conn_line.line.fill.background()

    # ─── RIGHT SIDE: Visual summary / stats ───
    right_x = Inches(7.5)

    # Stats card
    stats_bg = add_rounded_rect(slide, right_x, Inches(1.35), Inches(5.3), Inches(5.55),
                                 NAVY_MID, GRAY_700, Pt(1))
    set_shape_opacity(stats_bg, 50)

    add_textbox(slide, right_x + Inches(0.3), Inches(1.5), Inches(4.7), Inches(0.4),
                "DAILY PERFORMANCE", font_size=12, color=GRAY_500, bold=True,
                alignment=PP_ALIGN.CENTER)

    stats = [
        ("47", "Discoveries", ELEC_TEAL),
        ("94%", "Resolution Rate", GREEN_BRIGHT),
        ("2.3 min", "Avg Response", NEON_ORANGE),
        ("4.8★", "Member Satisfaction", PURPLE_BRIGHT),
    ]

    for j, (val, label, color) in enumerate(stats):
        sy = Inches(2.1) + j * Inches(1.15)
        stat_card = add_rounded_rect(slide, right_x + Inches(0.3), sy,
                                      Inches(4.7), Inches(0.9), NAVY_LIGHT, color, Pt(1))
        tf = stat_card.text_frame
        tf.word_wrap = True
        tf.margin_left = Inches(0.2)
        tf.margin_top = Inches(0.08)

        p1 = tf.paragraphs[0]
        p1.text = val
        p1.font.size = Pt(30)
        p1.font.color.rgb = color
        p1.font.bold = True
        p1.font.name = "Inter"
        p1.space_after = Pt(0)

        p2 = tf.add_paragraph()
        p2.text = label
        p2.font.size = Pt(13)
        p2.font.color.rgb = GRAY_400
        p2.font.name = "Inter"

    add_progress_dots(slide, 5)
    add_home_button(slide)

    add_notes(slide, """SLIDE 6 - HOW IT ACTUALLY WORKS
[TIMING: 2-3 minutes]

SAY: "Let me walk you through what a typical day looks like."

WALK THROUGH EACH PHASE:
- 6 AM: "System wakes up, starts scanning. Reddit, Facebook, LinkedIn, blog comments. Looking for real people with real insurance problems."
- 7 AM: "AI processes everything. Extracts the actual struggle, finds contact info, scores urgency."
- 8 AM: "Outreach begins. High priority gets a phone call. Medium gets a text. Others get a helpful email."
- All Day: "Conversations happen naturally. Voice AI talks to members, understands ALL options, finds best fit."

POINT TO RIGHT PANEL:
- "These are real-ish numbers from our prototype. 47 discoveries, 94% resolution, under 3 minutes average response."

KEY POINT: "This all runs autonomously. No human has to trigger anything."
""")

    return slide


# ═══════════════════════════════════════════════════════════════════════════════
# SLIDE 7: LIVE DEMO - THE DASHBOARD
# ═══════════════════════════════════════════════════════════════════════════════

def build_slide_7():
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_slide_bg_gradient(slide, '0A1628', '10203C')

    # Title bar
    title_bar = add_rounded_rect(slide, Inches(0), Inches(0), SLIDE_W, Inches(0.7),
                                  DEEP_NAVY)
    add_textbox(slide, Inches(0.4), Inches(0.1), Inches(3), Inches(0.5),
                "🏠 Lighthouse Dashboard", font_size=18, color=ELEC_TEAL, bold=True)
    add_textbox(slide, Inches(9.5), Inches(0.15), Inches(3.5), Inches(0.4),
                "Live  ·  Last updated: 2 min ago", font_size=12, color=GREEN_BRIGHT)

    # ─── KPI CARDS ROW ───
    kpis = [
        ("47", "Today's Discoveries", "↑ 12%", ELEC_TEAL),
        ("12", "Active Conversations", "↑ 3", PURPLE_BRIGHT),
        ("94%", "Resolution Rate", "↑ 2.1%", GREEN_BRIGHT),
        ("2.3 min", "Avg Response Time", "↓ 18%", NEON_ORANGE),
    ]

    kpi_w = Inches(2.9)
    kpi_h = Inches(1.3)
    kpi_gap = Inches(0.3)
    kpi_start_x = Inches(0.4)
    kpi_y = Inches(0.9)

    for i, (val, label, change, color) in enumerate(kpis):
        x = kpi_start_x + i * (kpi_w + kpi_gap)
        card = add_rounded_rect(slide, x, kpi_y, kpi_w, kpi_h, NAVY_MID, color, Pt(1))

        tf = card.text_frame
        tf.word_wrap = True
        tf.margin_left = Inches(0.15)
        tf.margin_top = Inches(0.1)

        p = tf.paragraphs[0]
        p.text = val
        p.font.size = Pt(34)
        p.font.color.rgb = color
        p.font.bold = True
        p.font.name = "Inter"
        p.space_after = Pt(2)

        p2 = tf.add_paragraph()
        p2.text = label
        p2.font.size = Pt(12)
        p2.font.color.rgb = GRAY_400
        p2.font.name = "Inter"
        p2.space_after = Pt(2)

        p3 = tf.add_paragraph()
        p3.text = change
        p3.font.size = Pt(11)
        p3.font.color.rgb = GREEN_BRIGHT
        p3.font.name = "JetBrains Mono"

    # ─── CHARTS ROW ───
    chart_y = Inches(2.45)
    chart_h = Inches(2.3)

    # Pain Point Distribution (left chart mockup)
    chart1_bg = add_rounded_rect(slide, Inches(0.4), chart_y, Inches(6.2), chart_h,
                                  NAVY_MID, GRAY_700, Pt(1))
    add_textbox(slide, Inches(0.7), chart_y + Inches(0.1), Inches(3), Inches(0.3),
                "Pain Point Distribution", font_size=13, color=GRAY_300, bold=True)

    # Simulated bar chart
    bars_data = [
        ("Claim Denied", 0.85, NEON_ORANGE),
        ("Coverage Gap", 0.70, ELEC_TEAL),
        ("Wait Times", 0.60, PURPLE_BRIGHT),
        ("Billing Error", 0.45, YELLOW_BRIGHT),
        ("Provider Issue", 0.35, GREEN_BRIGHT),
    ]

    for j, (lbl, pct, clr) in enumerate(bars_data):
        by = chart_y + Inches(0.5) + j * Inches(0.33)
        # Label
        add_textbox(slide, Inches(0.7), by, Inches(1.5), Inches(0.25),
                    lbl, font_size=10, color=GRAY_400, font_name="Inter")
        # Bar background
        bar_bg = slide.shapes.add_shape(
            MSO_SHAPE.ROUNDED_RECTANGLE,
            Inches(2.3), by + Inches(0.02), Inches(4.0), Inches(0.22),
        )
        bar_bg.fill.solid()
        bar_bg.fill.fore_color.rgb = NAVY_LIGHT
        bar_bg.line.fill.background()
        # Bar fill
        bar_fill = slide.shapes.add_shape(
            MSO_SHAPE.ROUNDED_RECTANGLE,
            Inches(2.3), by + Inches(0.02), Inches(4.0 * pct), Inches(0.22),
        )
        bar_fill.fill.solid()
        bar_fill.fill.fore_color.rgb = clr
        bar_fill.line.fill.background()

    # Sentiment Over Time (right chart mockup)
    chart2_bg = add_rounded_rect(slide, Inches(6.8), chart_y, Inches(6.1), chart_h,
                                  NAVY_MID, GRAY_700, Pt(1))
    add_textbox(slide, Inches(7.1), chart_y + Inches(0.1), Inches(3), Inches(0.3),
                "Sentiment Analysis", font_size=13, color=GRAY_300, bold=True)

    # Simulated sentiment indicators
    sentiments = [
        ("Positive", "62%", GREEN_BRIGHT),
        ("Neutral", "24%", GRAY_400),
        ("Frustrated", "14%", RED_BRIGHT),
    ]
    for j, (lbl, pct, clr) in enumerate(sentiments):
        sx = Inches(7.2) + j * Inches(1.8)
        sy = chart_y + Inches(0.6)

        # Big number
        add_textbox(slide, sx, sy, Inches(1.5), Inches(0.5),
                    pct, font_size=32, color=clr, bold=True, alignment=PP_ALIGN.CENTER)
        add_textbox(slide, sx, sy + Inches(0.5), Inches(1.5), Inches(0.3),
                    lbl, font_size=12, color=GRAY_400, alignment=PP_ALIGN.CENTER)

    # Trend line mockup (simple dots)
    trend_y_base = chart_y + Inches(1.5)
    trend_points = [0.6, 0.55, 0.62, 0.58, 0.65, 0.70, 0.72, 0.68, 0.75, 0.78]
    for k, tp in enumerate(trend_points):
        tx = Inches(7.3) + k * Inches(0.48)
        ty = trend_y_base + Inches(0.7) * (1 - tp)
        dot = add_circle(slide, tx, ty, Inches(0.1), ELEC_TEAL)

    # ─── MEMBER TABLE ───
    table_y = Inches(4.95)
    table_bg = add_rounded_rect(slide, Inches(0.4), table_y, Inches(12.5), Inches(2.2),
                                 NAVY_MID, GRAY_700, Pt(1))

    add_textbox(slide, Inches(0.7), table_y + Inches(0.1), Inches(3), Inches(0.3),
                "Recent Member Interactions", font_size=13, color=GRAY_300, bold=True)

    # Table headers
    headers = ["Name", "Source", "Pain Point", "Priority", "Status"]
    col_widths = [Inches(1.8), Inches(1.4), Inches(3.0), Inches(1.2), Inches(1.8)]
    col_starts = [Inches(0.7)]
    for w in col_widths[:-1]:
        col_starts.append(col_starts[-1] + w + Inches(0.3))

    header_y = table_y + Inches(0.45)
    for ci, (hdr, cs) in enumerate(zip(headers, col_starts)):
        add_textbox(slide, cs, header_y, col_widths[ci], Inches(0.25),
                    hdr, font_size=11, color=GRAY_500, bold=True, font_name="Inter")

    # Separator line
    sep = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE,
                                  Inches(0.7), header_y + Inches(0.3), Inches(12.0), Pt(1))
    sep.fill.solid()
    sep.fill.fore_color.rgb = GRAY_700
    sep.line.fill.background()

    # Table rows
    rows = [
        ("Sarah M.", "Reddit", "Claim denied after surgery", "🔴 High", "✅ Called - Resolved"),
        ("John D.", "Facebook", "Coverage confusion", "🟡 Medium", "💬 SMS sent"),
        ("Maria L.", "LinkedIn", "Premium increase question", "🟡 Medium", "📧 Email sent"),
        ("Robert K.", "Reddit", "Prior auth nightmare", "🔴 High", "📞 In conversation"),
        ("Lisa T.", "Blog", "Plan comparison help", "🟢 Low", "📧 Resources sent"),
    ]

    for ri, row in enumerate(rows):
        ry = header_y + Inches(0.4) + ri * Inches(0.3)
        for ci, (val, cs) in enumerate(zip(row, col_starts)):
            color = WHITE
            if ci == 3:  # Priority
                if "High" in val: color = RED_BRIGHT
                elif "Medium" in val: color = YELLOW_BRIGHT
                else: color = GREEN_BRIGHT
            elif ci == 4:  # Status
                if "Resolved" in val: color = GREEN_BRIGHT
                elif "conversation" in val: color = ELEC_TEAL
                else: color = GRAY_300

            add_textbox(slide, cs, ry, col_widths[ci], Inches(0.22),
                        val, font_size=11, color=color, font_name="JetBrains Mono")

    # "Running right now" indicator
    add_textbox(slide, Inches(9.5), Inches(7.0), Inches(3.5), Inches(0.25),
                "● This is all running right now",
                font_size=10, color=GREEN_BRIGHT, alignment=PP_ALIGN.RIGHT)

    add_progress_dots(slide, 6)
    add_home_button(slide)

    add_notes(slide, """SLIDE 7 - LIVE DEMO: THE DASHBOARD
[TIMING: 3-4 minutes - THIS IS THE MONEY SLIDE]

SAY: "Now let me show you what this actually looks like when it's running."

WALK THROUGH THE DASHBOARD:

TOP ROW (KPIs):
- "47 discoveries today. 12 active conversations right now. 94% resolution rate."
- "Average response time: 2.3 minutes. Not hours. Minutes."

CHARTS:
- "On the left, pain point distribution. Claim denials are #1 - that tells us a lot."
- "On the right, sentiment analysis. 62% positive because we're actually helping."

TABLE:
- "Here's the real magic. Sarah posted on Reddit about a denied claim. We found her at 6 AM, called by 8 AM, resolved by 10."
- "John was confused about coverage on Facebook. Got a text with a clear explanation."
- Point to each row briefly.

BOTTOM RIGHT:
- "And yeah - this is running right now."

TIP: If you have the actual Replit dashboard running, now is the time to switch to a live demo. If not, this slide IS the demo.

TRANSITION: "Let me show you the Voice AI that powers those conversations..."
""")

    return slide


# ═══════════════════════════════════════════════════════════════════════════════
# SLIDE 8: THE VOICE AI DEMO
# ═══════════════════════════════════════════════════════════════════════════════

def build_slide_8():
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_slide_bg_gradient(slide, '050B18', '0A1628')

    # Title section
    add_textbox(slide,
        Inches(0.6), Inches(0.3), Inches(8), Inches(0.7),
        "The Voice AI",
        font_size=42, color=WHITE, bold=True
    )
    add_textbox(slide,
        Inches(0.6), Inches(0.85), Inches(10), Inches(0.4),
        "QualifyOS + Summit Voice AI  ·  The voice people actually want to talk to",
        font_size=18, color=ELEC_TEAL
    )

    # ─── LEFT: Capabilities ───
    capabilities = [
        ("🌐", "24/7 Multilingual Support", "90+ languages, real-time translation"),
        ("💚", "Empathetic, Not Robotic", "Trained on real member conversations"),
        ("🏥", "CMS-Trained for Healthcare", "Knows regulations, plans, terminology"),
        ("🔗", "Integrates with Any System", "CRM, EHR, claims, billing — all of it"),
        ("🧠", "Context-Aware Memory", "Remembers the full conversation arc"),
    ]

    cap_x = Inches(0.6)
    cap_y_start = Inches(1.6)

    for i, (icon, title, desc) in enumerate(capabilities):
        y = cap_y_start + i * Inches(0.85)

        # Icon circle
        icon_circle = add_circle(slide, cap_x, y, Inches(0.5), NAVY_LIGHT, ELEC_TEAL, Pt(1))
        add_shape_text(icon_circle, icon, font_size=16)

        add_textbox(slide, cap_x + Inches(0.65), y + Inches(0.02), Inches(4.5), Inches(0.25),
                    title, font_size=16, color=WHITE, bold=True)
        add_textbox(slide, cap_x + Inches(0.65), y + Inches(0.28), Inches(4.5), Inches(0.22),
                    desc, font_size=12, color=GRAY_400)

    # ─── RIGHT: Voice visualization ───
    right_x = Inches(6.5)

    # Waveform visualization (abstract)
    wave_bg = add_rounded_rect(slide, right_x, Inches(1.5), Inches(6.3), Inches(2.8),
                                NAVY_MID, ELEC_TEAL, Pt(1))
    set_shape_glow(wave_bg, (0x06, 0xB6, 0xD4), radius=Pt(10), alpha=15000)

    # Voice icon in center
    voice_icon = add_circle(slide, right_x + Inches(2.4), Inches(2.1), Inches(1.5),
                             DEEP_NAVY, ELEC_TEAL, Pt(2))
    set_shape_glow(voice_icon, (0x06, 0xB6, 0xD4), radius=Pt(15), alpha=30000)
    add_shape_text(voice_icon, "🎙️", font_size=36)

    # Sound wave bars (simulated)
    wave_heights = [0.4, 0.7, 1.0, 0.8, 1.2, 0.9, 1.1, 0.6, 0.9, 1.0, 0.7, 0.5, 0.8, 1.1, 0.6]
    bar_w = Inches(0.12)
    bar_gap = Inches(0.08)
    wave_start_x = right_x + Inches(0.3)
    wave_center_y = Inches(3.0)

    for k, h in enumerate(wave_heights):
        bx = wave_start_x + k * (bar_w + bar_gap)
        bar_h = Inches(h * 0.5)
        by = wave_center_y - bar_h / 2

        bar = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, bx, by, bar_w, bar_h)
        bar.fill.solid()
        bar.fill.fore_color.rgb = ELEC_TEAL
        bar.line.fill.background()
        set_shape_opacity(bar, 50 + k * 2)

    # ─── CONVERSATION PREVIEW ───
    conv_y = Inches(4.7)
    conv_bg = add_rounded_rect(slide, right_x, conv_y, Inches(6.3), Inches(2.4),
                                NAVY_MID, GRAY_700, Pt(1))

    add_textbox(slide, right_x + Inches(0.2), conv_y + Inches(0.1), Inches(3), Inches(0.25),
                "SAMPLE CONVERSATION", font_size=10, color=GRAY_500, bold=True)

    # Member message
    member_bubble = add_rounded_rect(slide,
        right_x + Inches(0.3), conv_y + Inches(0.4), Inches(4.8), Inches(0.7),
        NAVY_LIGHT, GRAY_500, Pt(1))
    tf = member_bubble.text_frame
    tf.word_wrap = True
    tf.margin_left = Inches(0.1)
    tf.margin_top = Inches(0.05)
    p = tf.paragraphs[0]
    p.text = "👤 Member"
    p.font.size = Pt(10)
    p.font.color.rgb = GRAY_400
    p.font.bold = True
    p.font.name = "Inter"
    p2 = tf.add_paragraph()
    p2.text = '"I\'m confused about my prescription coverage... my pharmacy said my meds aren\'t covered anymore?"'
    p2.font.size = Pt(12)
    p2.font.color.rgb = WHITE
    p2.font.name = "Inter"

    # AI response
    ai_bubble = add_rounded_rect(slide,
        right_x + Inches(1.2), conv_y + Inches(1.25), Inches(4.8), Inches(0.9),
        NAVY_LIGHT, ELEC_TEAL, Pt(1))
    tf2 = ai_bubble.text_frame
    tf2.word_wrap = True
    tf2.margin_left = Inches(0.1)
    tf2.margin_top = Inches(0.05)
    p3 = tf2.paragraphs[0]
    p3.text = "🤖 Lighthouse AI"
    p3.font.size = Pt(10)
    p3.font.color.rgb = ELEC_TEAL
    p3.font.bold = True
    p3.font.name = "Inter"
    p4 = tf2.add_paragraph()
    p4.text = '"I understand how frustrating that must be. Let me pull up your specific plan and walk through your options. There may be alternatives we can find that are fully covered."'
    p4.font.size = Pt(12)
    p4.font.color.rgb = WHITE
    p4.font.name = "Inter"

    add_progress_dots(slide, 7)
    add_home_button(slide)

    add_notes(slide, """SLIDE 8 - VOICE AI DEMO
[TIMING: 2-3 minutes]

SAY: "Now the secret weapon - the Voice AI itself."

WALK THROUGH CAPABILITIES:
- "It speaks 90+ languages. Real-time translation."
- "It's empathetic. Trained on actual member conversations, not scripts."
- "It knows CMS regulations, plan details, the terminology."
- "It plugs into everything - CRM, claims, billing."

POINT TO CONVERSATION:
- Read both messages aloud.
- "Notice the AI doesn't say 'let me transfer you.' It says 'let me help you right now.'"
- "That's the difference. No hold time. No transfers. Actual help."

IF YOU HAVE AN AUDIO DEMO: Play it here. Even 30 seconds of the Voice AI in action is powerful.

KEY POINT: "This isn't a chatbot. It's a knowledgeable, empathetic voice that actually solves problems."
""")

    return slide


# ═══════════════════════════════════════════════════════════════════════════════
# SLIDE 9: BEYOND HEALTHCARE
# ═══════════════════════════════════════════════════════════════════════════════

def build_slide_9():
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_slide_bg_gradient(slide, '050B18', '0A1628')

    add_textbox(slide,
        Inches(0.6), Inches(0.3), Inches(8), Inches(0.7),
        "Beyond Healthcare",
        font_size=42, color=WHITE, bold=True
    )
    add_textbox(slide,
        Inches(0.6), Inches(0.85), Inches(10), Inches(0.4),
        "Same tech stack, different problems. This pattern works everywhere.",
        font_size=18, color=GRAY_400
    )

    # 2x3 Grid of cards
    cards = [
        ("🏠", "Roofing", "Where it started", "0% missed calls\n34% conversion", NEON_ORANGE),
        ("🔧", "Home Services", "HVAC, plumbing, electric", "24/7 booking\n3x lead volume", ELEC_TEAL),
        ("🛡️", "Insurance Brokers", "Independent agents", "Automated quoting\n60% time saved", PURPLE_BRIGHT),
        ("📞", "Call Centers", "Enterprise support", "87% automation\n$2M+ saved/yr", RED_BRIGHT),
        ("🎧", "Customer Service", "Any industry", "< 30s response\n95% satisfaction", GREEN_BRIGHT),
        ("🏗️", "Complex Products", "Finance, legal, tech", "Guided journeys\n2x conversion", YELLOW_BRIGHT),
    ]

    card_w = Inches(3.8)
    card_h = Inches(2.2)
    gap_x = Inches(0.45)
    gap_y = Inches(0.4)
    start_x = Inches(0.6)
    start_y = Inches(1.5)

    for i, (icon, title, subtitle, stat, color) in enumerate(cards):
        row = i // 3
        col = i % 3
        x = start_x + col * (card_w + gap_x)
        y = start_y + row * (card_h + gap_y)

        card = add_rounded_rect(slide, x, y, card_w, card_h, NAVY_MID, color, Pt(1))

        tf = card.text_frame
        tf.word_wrap = True
        tf.margin_left = Inches(0.2)
        tf.margin_top = Inches(0.15)
        tf.margin_right = Inches(0.15)

        # Icon
        p = tf.paragraphs[0]
        p.text = icon
        p.font.size = Pt(30)
        p.alignment = PP_ALIGN.LEFT
        p.space_after = Pt(4)

        # Title
        p2 = tf.add_paragraph()
        p2.text = title
        p2.font.size = Pt(20)
        p2.font.color.rgb = color
        p2.font.bold = True
        p2.font.name = "Inter"
        p2.space_after = Pt(2)

        # Subtitle
        p3 = tf.add_paragraph()
        p3.text = subtitle
        p3.font.size = Pt(13)
        p3.font.color.rgb = GRAY_400
        p3.font.name = "Inter"
        p3.space_after = Pt(8)

        # Stats
        p4 = tf.add_paragraph()
        p4.text = stat
        p4.font.size = Pt(14)
        p4.font.color.rgb = WHITE
        p4.font.bold = True
        p4.font.name = "JetBrains Mono"

    add_progress_dots(slide, 8)
    add_home_button(slide)

    add_notes(slide, """SLIDE 9 - BEYOND HEALTHCARE
[TIMING: 1-2 minutes]

SAY: "Healthcare is our focus, but the tech works everywhere."

QUICKLY TOUCH EACH CARD:
- "Roofing - that's where we proved it."
- "Home services, insurance brokers, call centers - same pattern."
- "Anywhere there's a complex product and frustrated customers, Lighthouse fits."

DON'T linger here too long. This slide is about showing scale, not deep-diving every vertical.

KEY POINT: "The stack is the product. The vertical is just the application."

TRANSITION: "Let me show you what's under the hood..."
""")

    return slide


# ═══════════════════════════════════════════════════════════════════════════════
# SLIDE 10: THE n8n BUILD
# ═══════════════════════════════════════════════════════════════════════════════

def build_slide_10():
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_slide_bg_gradient(slide, '050B18', '0A1628')

    add_textbox(slide,
        Inches(0.6), Inches(0.3), Inches(8), Inches(0.7),
        "Under the Hood",
        font_size=42, color=WHITE, bold=True
    )
    add_textbox(slide,
        Inches(0.6), Inches(0.85), Inches(10), Inches(0.4),
        "The n8n automation workflow that powers everything",
        font_size=18, color=GRAY_400
    )

    # ─── WORKFLOW DIAGRAM ───
    # Nodes as rounded rectangles connected by lines
    workflow_nodes = [
        ("⏰", "Trigger", "6 AM Daily", Inches(0.5), Inches(2.2), GRAY_500),
        ("📡", "Reddit API", "Scan posts", Inches(2.5), Inches(1.5), NEON_ORANGE),
        ("📘", "Facebook API", "Scan groups", Inches(2.5), Inches(2.7), ELEC_TEAL),
        ("💼", "LinkedIn API", "Scan threads", Inches(2.5), Inches(3.9), PURPLE_BRIGHT),
        ("🧠", "Claude AI", "Analyze &\nScore", Inches(5.0), Inches(2.5), ELEC_TEAL),
        ("📊", "Google Sheets", "Data store", Inches(7.3), Inches(1.5), GREEN_BRIGHT),
        ("🗄️", "CRM", "Member DB", Inches(7.3), Inches(3.5), PURPLE_BRIGHT),
        ("🔀", "Router", "Priority\nSplit", Inches(9.3), Inches(2.5), YELLOW_BRIGHT),
        ("📞", "Voice AI", "High priority\ncalls", Inches(11.3), Inches(1.5), RED_BRIGHT),
        ("💬", "SMS", "Medium\npriority", Inches(11.3), Inches(2.7), NEON_ORANGE),
        ("📧", "Email", "Lower\npriority", Inches(11.3), Inches(3.9), GREEN_BRIGHT),
    ]

    node_w = Inches(1.6)
    node_h = Inches(0.95)

    # Draw connections first (so they're behind nodes)
    connections = [
        (0, 1), (0, 2), (0, 3),  # Trigger to APIs
        (1, 4), (2, 4), (3, 4),  # APIs to Claude
        (4, 5), (4, 6),          # Claude to storage
        (4, 7),                   # Claude to Router
        (7, 8), (7, 9), (7, 10), # Router to outputs
    ]

    for start_idx, end_idx in connections:
        sx = workflow_nodes[start_idx][3] + node_w
        sy = workflow_nodes[start_idx][4] + node_h / 2
        ex = workflow_nodes[end_idx][3]
        ey = workflow_nodes[end_idx][4] + node_h / 2
        try:
            connector = slide.shapes.add_connector(
                1, sx, sy, ex, ey
            )
            connector.line.color.rgb = GRAY_700
            connector.line.width = Pt(1.5)
        except:
            pass

    # Draw nodes
    for icon, name, desc, x, y, color in workflow_nodes:
        node = add_rounded_rect(slide, x, y, node_w, node_h, NAVY_MID, color, Pt(1.5))
        set_shape_glow(node, (color[0], color[1], color[2]),
                       radius=Pt(5), alpha=20000)

        tf = node.text_frame
        tf.word_wrap = True
        tf.margin_left = Inches(0.08)
        tf.margin_top = Inches(0.05)

        p = tf.paragraphs[0]
        p.text = f"{icon} {name}"
        p.font.size = Pt(11)
        p.font.color.rgb = color
        p.font.bold = True
        p.font.name = "Inter"
        p.alignment = PP_ALIGN.CENTER
        p.space_after = Pt(1)

        p2 = tf.add_paragraph()
        p2.text = desc
        p2.font.size = Pt(9)
        p2.font.color.rgb = GRAY_400
        p2.font.name = "Inter"
        p2.alignment = PP_ALIGN.CENTER

    # ─── CALLOUT ANNOTATIONS ───
    callouts = [
        (Inches(0.3), Inches(3.5), "Runs every morning\nat 6 AM", ELEC_TEAL),
        (Inches(4.5), Inches(4.6), "Claude analyzes sentiment\nand extracts pain points", ELEC_TEAL),
        (Inches(8.8), Inches(4.6), "Routes high-priority\nto instant call", NEON_ORANGE),
    ]

    for cx, cy, text, color in callouts:
        callout = add_rounded_rect(slide, cx, cy, Inches(2.6), Inches(0.65),
                                    NAVY_LIGHT, color, Pt(1))
        tf = callout.text_frame
        tf.word_wrap = True
        tf.margin_left = Inches(0.1)
        tf.margin_top = Inches(0.05)
        p = tf.paragraphs[0]
        p.text = text
        p.font.size = Pt(11)
        p.font.color.rgb = color
        p.font.name = "Inter"
        p.alignment = PP_ALIGN.CENTER

    # Bottom quote
    quote_bg = add_rounded_rect(slide, Inches(1.5), Inches(5.8), Inches(10.3), Inches(0.6),
                                 NAVY_MID, GRAY_700, Pt(1))
    add_shape_text(quote_bg,
                   "With tools like Claude Code, we prototype workflows like this in hours, not weeks.",
                   font_size=16, color=GRAY_300, bold=False)

    # Tech stack badges
    badges = ["n8n", "Claude AI", "Voice AI", "Google Sheets", "Replit", "Python"]
    badge_start_x = Inches(2.5)
    badge_y = Inches(6.6)
    for bi, badge_text in enumerate(badges):
        bx = badge_start_x + bi * Inches(1.5)
        b = add_rounded_rect(slide, bx, badge_y, Inches(1.3), Inches(0.35),
                              NAVY_LIGHT, ELEC_TEAL, Pt(1))
        add_shape_text(b, badge_text, font_size=10, color=ELEC_TEAL, bold=True)

    add_progress_dots(slide, 9)
    add_home_button(slide)

    add_notes(slide, """SLIDE 10 - THE n8n BUILD
[TIMING: 2-3 minutes]

SAY: "Here's what's actually running under the hood."

WALK THROUGH THE WORKFLOW:
- "Trigger fires at 6 AM every day."
- "Hits Reddit, Facebook, LinkedIn APIs simultaneously."
- "Everything flows into Claude for analysis - sentiment, pain points, urgency scoring."
- "Data goes to Google Sheets and CRM."
- "Router splits by priority: high gets a call, medium gets a text, lower gets an email."

POINT TO CALLOUTS:
- Highlight the key decision points.

BOTTOM QUOTE:
- "And here's the thing - with Claude Code, we can prototype a workflow like this in hours, not weeks. The speed of iteration is insane."

TECH BADGES:
- "Our stack: n8n for automation, Claude for AI, custom Voice AI, all hosted on Replit."

CLOSING THOUGHT:
- "This is a demo. But it's a working demo. And we're adding to it every day."

TRANSITION: "So... that's what we've been building. Questions?"
""")

    return slide


# ═══════════════════════════════════════════════════════════════════════════════
# BUILD ALL SLIDES
# ═══════════════════════════════════════════════════════════════════════════════

print("Building Lighthouse presentation...")
print("  → Slide 1: Simple Intro")
build_slide_1()
print("  → Slide 2: The Real Story")
build_slide_2()
print("  → Slide 3: Where We Started - Roofing")
build_slide_3()
print("  → Slide 4: What We Saw in Healthcare")
build_slide_4()
print("  → Slide 5: The Idea - Find Them First")
build_slide_5()
print("  → Slide 6: How It Actually Works")
build_slide_6()
print("  → Slide 7: Live Demo - Dashboard")
build_slide_7()
print("  → Slide 8: Voice AI Demo")
build_slide_8()
print("  → Slide 9: Beyond Healthcare")
build_slide_9()
print("  → Slide 10: The n8n Build")
build_slide_10()

# Save
output_path = "/home/user/D-D2/Lighthouse_Demo.pptx"
prs.save(output_path)
print(f"\n✓ Presentation saved to: {output_path}")
print(f"  Total slides: {len(prs.slides)}")
print(f"  Date on deck: {DATE_STR}")
