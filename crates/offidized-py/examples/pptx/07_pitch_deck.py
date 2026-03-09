"""Build a complete pitch deck from scratch.

Demonstrates building a real multi-slide presentation programmatically
using shape effects, charts, tables, transitions, and speaker notes --
most of which require direct XML manipulation in python-pptx.
"""

from offidized import Presentation, SlideTransition

EMU = 914400
SLIDE_W = 12192000
SLIDE_H = 6858000

pres = Presentation()
pres.set_slide_size(SLIDE_W, SLIDE_H)

# Theme colors
NAVY = "0B1D3A"
BLUE = "1B65A6"
ORANGE = "FF6B35"
WHITE = "FFFFFF"
GRAY = "B0B0B0"
LIGHT_BG = "F0F2F5"


def make_text(shape, text, size, color, bold=False, align="ctr"):
    para = shape.add_paragraph_with_text(text)
    run = para.get_run(0)
    run.set_font_size(size)
    run.set_font_color(color)
    run.set_bold(bold)
    run.set_font_name("Calibri")
    para.set_alignment(align)
    return para


# ═══════════════════════════════════════════════════════════════════════
# Slide 1: Title
# ═══════════════════════════════════════════════════════════════════════

idx = pres.add_slide()
slide = pres.get_slide(idx)
slide.set_background_solid(NAVY)
slide.set_transition(SlideTransition("fade"))
slide.set_notes("Introduce the company. Pause for effect.")

# Logo / brand shape
s = slide.add_shape("Brand")
shape = slide.get_shape(s)
shape.set_geometry(int(4.5 * EMU), int(1.5 * EMU), int(3.5 * EMU), int(1.5 * EMU))
shape.set_preset_geometry("roundRect")
shape.set_gradient_fill([(0, BLUE), (100, ORANGE)], angle=90)
shape.set_shadow(40000, 40000, 100000, "000000")
make_text(shape, "ACME", 5000, WHITE, bold=True)

# Tagline
s = slide.add_shape("Tagline")
shape = slide.get_shape(s)
shape.set_geometry(int(1.5 * EMU), int(3.5 * EMU), int(9.5 * EMU), EMU)
make_text(shape, "Reinventing Widgets for the AI Era", 2800, GRAY)

# Date
s = slide.add_shape("Date")
shape = slide.get_shape(s)
shape.set_geometry(int(1.5 * EMU), int(5 * EMU), int(9.5 * EMU), int(0.5 * EMU))
make_text(shape, "Series A  |  Q1 2026  |  Confidential", 1200, "666666")

# ═══════════════════════════════════════════════════════════════════════
# Slide 2: Problem
# ═══════════════════════════════════════════════════════════════════════

idx = pres.add_slide()
slide = pres.get_slide(idx)
slide.set_background_solid(LIGHT_BG)
slide.set_transition(SlideTransition("push"))
slide.set_notes("Explain the pain point. Use customer quotes.")

s = slide.add_shape("Section")
shape = slide.get_shape(s)
shape.set_geometry(0, 0, SLIDE_W, int(1.2 * EMU))
shape.set_solid_fill_srgb(NAVY)
make_text(shape, "The Problem", 3200, WHITE, bold=True)

problems = [
    ("\u26a0", "Manual processes cost enterprises $2.3B/year"),
    ("\u23f1", "Average widget deployment takes 6+ months"),
    ("\u274c", "87% of widget projects fail in first year"),
]

for i, (icon, text) in enumerate(problems):
    y = int((2 + i * 1.5) * EMU)

    # Icon circle
    s = slide.add_shape(f"Icon{i}")
    shape = slide.get_shape(s)
    shape.set_geometry(int(1.5 * EMU), y, int(0.8 * EMU), int(0.8 * EMU))
    shape.set_preset_geometry("ellipse")
    shape.set_solid_fill_srgb(ORANGE)
    make_text(shape, icon, 2000, WHITE, bold=True)

    # Text
    s = slide.add_shape(f"Text{i}")
    shape = slide.get_shape(s)
    shape.set_geometry(int(2.8 * EMU), y, int(8 * EMU), int(0.8 * EMU))
    make_text(shape, text, 2000, NAVY, align="l")

# ═══════════════════════════════════════════════════════════════════════
# Slide 3: Solution
# ═══════════════════════════════════════════════════════════════════════

idx = pres.add_slide()
slide = pres.get_slide(idx)
slide.set_background_solid(LIGHT_BG)
slide.set_transition(SlideTransition("wipe"))
slide.set_notes("Demo the product if possible.")

s = slide.add_shape("Section")
shape = slide.get_shape(s)
shape.set_geometry(0, 0, SLIDE_W, int(1.2 * EMU))
shape.set_solid_fill_srgb(BLUE)
make_text(shape, "Our Solution", 3200, WHITE, bold=True)

# Feature cards
features = [
    ("AI-Powered", "Automated widget\nconfiguration", BLUE),
    ("10x Faster", "Deploy in days,\nnot months", ORANGE),
    ("99.9% SLA", "Enterprise-grade\nreliability", "548235"),
]

for i, (title, desc, color) in enumerate(features):
    x = int((1 + i * 3.5) * EMU)
    s = slide.add_shape(title)
    shape = slide.get_shape(s)
    shape.set_geometry(x, int(2 * EMU), int(3 * EMU), int(3.5 * EMU))
    shape.set_preset_geometry("roundRect")
    shape.set_solid_fill_srgb(WHITE)
    shape.set_outline(color, width_pt=2.0)
    shape.set_shadow(30000, 30000, 60000, "CCCCCC")

    # Title
    para = shape.add_paragraph_with_text(title)
    run = para.get_run(0)
    run.set_font_size(2200)
    run.set_font_color(color)
    run.set_bold(True)
    para.set_alignment("ctr")

    # Spacer + description
    para2 = shape.add_paragraph()
    para2.set_space_before_pts(800)
    r = para2.add_run(desc)
    r.set_font_size(1400)
    r.set_font_color("555555")
    para2.set_alignment("ctr")

# ═══════════════════════════════════════════════════════════════════════
# Slide 4: Traction (chart + table)
# ═══════════════════════════════════════════════════════════════════════

idx = pres.add_slide()
slide = pres.get_slide(idx)
slide.set_background_solid(LIGHT_BG)
slide.set_transition(SlideTransition("fade"))
slide.set_notes("Emphasize MoM growth. Highlight key logos.")

s = slide.add_shape("Section")
shape = slide.get_shape(s)
shape.set_geometry(0, 0, SLIDE_W, int(1.2 * EMU))
shape.set_solid_fill_srgb(NAVY)
make_text(shape, "Traction", 3200, WHITE, bold=True)

# Chart
chart_idx = slide.add_chart("bar")
chart = slide.get_chart(chart_idx)
chart.set_title("MRR Growth ($K)")
for month, val in [
    ("Jan", 45),
    ("Feb", 62),
    ("Mar", 89),
    ("Apr", 121),
    ("May", 158),
    ("Jun", 203),
]:
    chart.add_data_point(month, val)
chart.set_value_axis_title("$K MRR")

# KPI table
table_idx = slide.add_table(3, 3)
table = slide.get_table(table_idx)
table.set_geometry(int(6.5 * EMU), int(2 * EMU), int(5 * EMU), int(2.5 * EMU))

kpis = [
    ("Metric", "Current", "Target"),
    ("ARR", "$2.4M", "$10M"),
    ("Customers", "47", "200"),
]
for row, vals in enumerate(kpis):
    for col, val in enumerate(vals):
        table.set_cell_text(row, col, val)
        table.set_cell_font_size(row, col, 1200)
        if row == 0:
            table.set_cell_fill(row, col, NAVY)
            table.set_cell_font_color(row, col, WHITE)
            table.set_cell_bold(row, col, True)

# ═══════════════════════════════════════════════════════════════════════
# Slide 5: The Ask
# ═══════════════════════════════════════════════════════════════════════

idx = pres.add_slide()
slide = pres.get_slide(idx)
slide.set_background_solid(NAVY)
slide.set_transition(SlideTransition("fade"))
slide.set_notes("State the ask clearly. Open for questions.")

s = slide.add_shape("Ask")
shape = slide.get_shape(s)
shape.set_geometry(int(2 * EMU), int(1.5 * EMU), int(8.5 * EMU), int(1.5 * EMU))
make_text(shape, "Raising $5M Series A", 4000, WHITE, bold=True)

s = slide.add_shape("Use")
shape = slide.get_shape(s)
shape.set_geometry(int(2 * EMU), int(3.2 * EMU), int(8.5 * EMU), EMU)
make_text(shape, "to scale engineering, sales, and expand to EMEA", 2000, GRAY)

# CTA button
s = slide.add_shape("CTA")
shape = slide.get_shape(s)
shape.set_geometry(int(4 * EMU), int(5 * EMU), int(4.5 * EMU), int(0.8 * EMU))
shape.set_preset_geometry("roundRect")
shape.set_solid_fill_srgb(ORANGE)
shape.set_shadow(30000, 30000, 60000, "000000")
make_text(shape, "investor@acme.ai", 1800, WHITE, bold=True)

print(f"Built pitch deck: {pres.slide_count()} slides")
pres.save("07_pitch_deck.pptx")
print("Created 07_pitch_deck.pptx")
