"""In-memory roundtrip: create, serialize to bytes, reload, and modify.

offidized supports full bytes roundtrip (from_bytes / to_bytes) for all
three formats. This enables serverless workflows, streaming pipelines,
and testing without touching the filesystem.
"""

from offidized import Presentation

EMU = 914400

# ─── Step 1: Create a presentation in memory ────────────────────────

pres = Presentation()
pres.set_slide_size(12192000, 6858000)

slide_idx = pres.add_slide()
slide = pres.get_slide(slide_idx)
slide.set_background_solid("2F5597")

s = slide.add_shape("Original")
shape = slide.get_shape(s)
shape.set_geometry(2 * EMU, 2 * EMU, 8 * EMU, 3 * EMU)
shape.set_preset_geometry("roundRect")
shape.set_solid_fill_srgb("FFFFFF")

para = shape.add_paragraph_with_text("Version 1")
run = para.get_run(0)
run.set_font_size(3600)
run.set_font_color("2F5597")
run.set_bold(True)
para.set_alignment("ctr")

# ─── Step 2: Serialize to bytes (no filesystem) ─────────────────────

pptx_bytes = pres.to_bytes()
print(f"Serialized to {len(pptx_bytes):,} bytes")

# ─── Step 3: Reload from bytes and modify ────────────────────────────

pres2 = Presentation.from_bytes(pptx_bytes)
print(f"Reloaded: {pres2.slide_count()} slide(s)")

# Modify the reloaded presentation
slide = pres2.get_slide(0)
shape = slide.get_shape(0)
para = shape.get_paragraph(0)
run = para.get_run(0)
run.set_text("Version 2 (Modified)")
run.set_font_color("C0504D")

# Add a second slide
slide2_idx = pres2.add_slide()
slide2 = pres2.get_slide(slide2_idx)
slide2.set_background_solid("548235")
s = slide2.add_shape("Added")
shape2 = slide2.get_shape(s)
shape2.set_geometry(2 * EMU, 2 * EMU, 8 * EMU, 3 * EMU)
shape2.set_preset_geometry("roundRect")
shape2.set_solid_fill_srgb("FFFFFF")
para = shape2.add_paragraph_with_text("Added in Round 2")
run = para.get_run(0)
run.set_font_size(3600)
run.set_font_color("548235")
run.set_bold(True)
para.set_alignment("ctr")

# ─── Step 4: Serialize again ─────────────────────────────────────────

pptx_bytes2 = pres2.to_bytes()
print(f"Re-serialized to {len(pptx_bytes2):,} bytes ({pres2.slide_count()} slides)")

# ─── Step 5: Verify the full roundtrip ───────────────────────────────

pres3 = Presentation.from_bytes(pptx_bytes2)
assert pres3.slide_count() == 2

slide = pres3.get_slide(0)
shape = slide.get_shape(0)
text = shape.get_paragraph(0).get_run(0).text()
assert "Version 2" in text
print(f"Verified: slide 1 text = '{text}'")

slide2 = pres3.get_slide(1)
shape2 = slide2.get_shape(0)
text2 = shape2.get_paragraph(0).get_run(0).text()
assert "Round 2" in text2
print(f"Verified: slide 2 text = '{text2}'")

# Save final version
pres3.save("06_in_memory_roundtrip.pptx")
print("Created 06_in_memory_roundtrip.pptx")
