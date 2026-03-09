"""Shape effects: shadow, glow, reflection, pattern fills, and gradients.

python-pptx has no API for shape effects (shadow, glow, reflection) or
pattern fills. offidized exposes the full OOXML effect pipeline.
"""

from offidized import Presentation

EMU = 914400  # 1 inch in EMUs

pres = Presentation()
pres.set_slide_size(12192000, 6858000)  # 16:9 widescreen
slide_idx = pres.add_slide()
slide = pres.get_slide(slide_idx)

# --- Shape 1: Gradient fill with drop shadow ---
s1 = slide.add_shape("Gradient + Shadow")
shape1 = slide.get_shape(s1)
shape1.set_geometry(EMU, EMU, 3 * EMU, 2 * EMU)
shape1.set_preset_geometry("roundRect")
shape1.set_gradient_fill(
    stops=[(0, "4472C4"), (50, "2F5597"), (100, "1B3A6B")],
    angle=270,  # top to bottom
)
shape1.set_shadow(
    offset_x=50000,  # ~0.5pt right
    offset_y=50000,  # ~0.5pt down
    blur_radius=100000,
    color="000000",
)
# Add text
para = shape1.add_paragraph_with_text("Drop Shadow")
run = para.get_run(0)
run.set_font_color("FFFFFF")
run.set_font_size(2400)  # 24pt
run.set_bold(True)
para.set_alignment("ctr")

# --- Shape 2: Solid fill with glow effect ---
s2 = slide.add_shape("Glow Effect")
shape2 = slide.get_shape(s2)
shape2.set_geometry(5 * EMU, EMU, 3 * EMU, 2 * EMU)
shape2.set_preset_geometry("ellipse")
shape2.set_solid_fill_srgb("FF6B35")
shape2.set_glow(radius=200000, color="FF6B35")  # warm orange glow
para = shape2.add_paragraph_with_text("Glow")
run = para.get_run(0)
run.set_font_color("FFFFFF")
run.set_font_size(2800)
run.set_bold(True)
para.set_alignment("ctr")

# --- Shape 3: Pattern fill with outline ---
s3 = slide.add_shape("Pattern Fill")
shape3 = slide.get_shape(s3)
shape3.set_geometry(9 * EMU, EMU, 3 * EMU, 2 * EMU)
shape3.set_preset_geometry("diamond")
shape3.set_pattern_fill("dkDnDiag", foreground="2F5597", background="D6E4F0")
shape3.set_outline("2F5597", width_pt=2.0)
para = shape3.add_paragraph_with_text("Pattern")
run = para.get_run(0)
run.set_font_color("1B3A6B")
run.set_font_size(2000)
run.set_bold(True)
para.set_alignment("ctr")

# --- Shape 4: Reflection effect ---
s4 = slide.add_shape("Reflection")
shape4 = slide.get_shape(s4)
shape4.set_geometry(3 * EMU, int(3.5 * EMU), 3 * EMU, 2 * EMU)
shape4.set_preset_geometry("rect")
shape4.set_solid_fill_srgb("548235")
shape4.set_reflection(blur_radius=50000, distance=30000)
para = shape4.add_paragraph_with_text("Reflection")
run = para.get_run(0)
run.set_font_color("FFFFFF")
run.set_font_size(2400)
run.set_bold(True)
para.set_alignment("ctr")

# --- Shape 5: All effects combined ---
s5 = slide.add_shape("All Effects")
shape5 = slide.get_shape(s5)
shape5.set_geometry(7 * EMU, int(3.5 * EMU), 3 * EMU, 2 * EMU)
shape5.set_preset_geometry("roundRect")
shape5.set_gradient_fill(
    stops=[(0, "C0504D"), (100, "8B2F2F")],
    angle=315,
)
shape5.set_shadow(offset_x=40000, offset_y=40000, blur_radius=80000, color="000000")
shape5.set_glow(radius=100000, color="C0504D")
shape5.set_outline("FFFFFF", width_pt=1.5, dash="solid")
para = shape5.add_paragraph_with_text("Combined")
run = para.get_run(0)
run.set_font_color("FFFFFF")
run.set_font_size(2400)
run.set_bold(True)
para.set_alignment("ctr")

pres.save("01_styled_shapes.pptx")
print("Created 01_styled_shapes.pptx")
