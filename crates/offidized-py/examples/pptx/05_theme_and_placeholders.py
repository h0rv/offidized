"""Theme manipulation and slide placeholders.

python-pptx provides read-only access to theme colors and has limited
placeholder manipulation. offidized lets you read and write theme colors,
fonts, and placeholder metadata directly.
"""

from offidized import Presentation

EMU = 914400

pres = Presentation()
pres.set_slide_size(12192000, 6858000)

# ─── Read and modify theme colors ────────────────────────────────────

# Read existing theme (may be None on a blank presentation)
print("Theme colors:")
for name in [
    "dk1",
    "lt1",
    "dk2",
    "lt2",
    "accent1",
    "accent2",
    "accent3",
    "accent4",
    "accent5",
    "accent6",
    "hlink",
    "folHlink",
]:
    color = pres.theme_color(name)
    print(f"  {name}: {color}")

# Set a custom color scheme
pres.set_theme_color("accent1", "FF6B35")  # orange
pres.set_theme_color("accent2", "2EC4B6")  # teal
pres.set_theme_color("accent3", "E71D36")  # red
pres.set_theme_color("accent4", "011627")  # dark navy

# Check theme fonts
fonts = pres.theme_fonts()
if fonts:
    print(f"Theme fonts: major={fonts[0]}, minor={fonts[1]}")

# ─── Slide with placeholder-like shapes ──────────────────────────────

slide_idx = pres.add_slide()
slide = pres.get_slide(slide_idx)
slide.set_background_solid("011627")

# Title placeholder-style shape
s1 = slide.add_shape("Title")
shape1 = slide.get_shape(s1)
shape1.set_geometry(EMU, int(0.5 * EMU), 10 * EMU, int(1.5 * EMU))
shape1.set_placeholder_kind("title")
shape1.set_placeholder_idx(0)
para = shape1.add_paragraph_with_text("Custom Theme Demo")
run = para.get_run(0)
run.set_font_color("FFFFFF")
run.set_font_size(4000)
run.set_bold(True)
para.set_alignment("ctr")

# Subtitle placeholder-style shape
s2 = slide.add_shape("Subtitle")
shape2 = slide.get_shape(s2)
shape2.set_geometry(EMU, int(2.2 * EMU), 10 * EMU, EMU)
shape2.set_placeholder_kind("body")
shape2.set_placeholder_idx(1)
para = shape2.add_paragraph_with_text(
    "Theme colors and placeholders, fully programmable"
)
run = para.get_run(0)
run.set_font_color("AAAAAA")
run.set_font_size(2000)
para.set_alignment("ctr")

# Color swatches showing the theme colors
colors = [
    ("accent1", "FF6B35", "Accent 1"),
    ("accent2", "2EC4B6", "Accent 2"),
    ("accent3", "E71D36", "Accent 3"),
    ("accent4", "011627", "Accent 4"),
]

for i, (name, hex_color, label) in enumerate(colors):
    x = int((1.5 + i * 2.5) * EMU)
    s = slide.add_shape(label)
    shape = slide.get_shape(s)
    shape.set_geometry(x, int(3.8 * EMU), int(2 * EMU), int(2 * EMU))
    shape.set_preset_geometry("roundRect")
    shape.set_solid_fill_srgb(hex_color)
    shape.set_outline("FFFFFF", width_pt=1.0)

    para = shape.add_paragraph_with_text(label)
    run = para.get_run(0)
    run.set_font_color("FFFFFF")
    run.set_font_size(1400)
    run.set_bold(True)
    para.set_alignment("ctr")

    para2 = shape.add_paragraph_with_text(f"#{hex_color}")
    run2 = para2.get_run(0)
    run2.set_font_color("FFFFFF")
    run2.set_font_size(1000)
    run2.set_font_name("Consolas")
    para2.set_alignment("ctr")

# ─── Slide masters info ──────────────────────────────────────────────

print(f"\nSlide masters: {pres.slide_master_count()}")
print(f"Total layouts: {pres.layout_count()}")
print(f"Slides: {pres.slide_count()}")

pres.save("05_theme_and_placeholders.pptx")
print("Created 05_theme_and_placeholders.pptx")
