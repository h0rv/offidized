"""Rich text formatting: multi-run paragraphs with mixed styles, bullets,
character spacing, baseline shifts, and hyperlinks.

Shows the full text formatting pipeline that offidized exposes, including
character spacing and baseline (superscript/subscript) which python-pptx
handles awkwardly or not at all.
"""

from offidized import Presentation

EMU = 914400

pres = Presentation()
pres.set_slide_size(12192000, 6858000)
slide_idx = pres.add_slide()
slide = pres.get_slide(slide_idx)

# ─── Title with mixed formatting ────────────────────────────────────

s = slide.add_shape("Title")
shape = slide.get_shape(s)
shape.set_geometry(EMU, int(0.3 * EMU), 10 * EMU, int(1.2 * EMU))

para = shape.add_paragraph()
para.set_alignment("ctr")

r1 = para.add_run("offi")
r1.set_font_size(4400)
r1.set_font_color("2F5597")
r1.set_bold(True)
r1.set_font_name("Calibri")

r2 = para.add_run("dized")
r2.set_font_size(4400)
r2.set_font_color("C0504D")
r2.set_bold(True)
r2.set_font_name("Calibri")

r3 = para.add_run("  Text Formatting")
r3.set_font_size(3200)
r3.set_font_color("404040")
r3.set_font_name("Calibri")

# ─── Bullet list with icons ─────────────────────────────────────────

s2 = slide.add_shape("Features")
shape2 = slide.get_shape(s2)
shape2.set_geometry(EMU, int(1.8 * EMU), 5 * EMU, int(4.5 * EMU))

features = [
    ("Bold", True, False, False),
    ("Italic", False, True, False),
    ("Underline", False, False, True),
    ("Mixed styles in one paragraph", False, False, False),
]

for text, bold, italic, underline in features:
    p = shape2.add_paragraph()
    p.set_bullet_char("\u2022")
    p.set_bullet_color("2F5597")
    p.set_margin_left(457200)  # 0.5 inch
    p.set_indent(-228600)  # hanging indent
    r = p.add_run(text)
    r.set_font_size(1800)
    r.set_font_color("333333")
    if bold:
        r.set_bold(True)
    if italic:
        r.set_italic(True)
    if underline:
        r.set_underline("sng")

# Add a paragraph with superscript/subscript
p = shape2.add_paragraph()
p.set_bullet_char("\u2022")
p.set_bullet_color("2F5597")
p.set_margin_left(457200)
p.set_indent(-228600)

r = p.add_run("H")
r.set_font_size(1800)
r.set_font_color("333333")

r_sub = p.add_run("2")
r_sub.set_font_size(1200)
r_sub.set_font_color("C0504D")
r_sub.set_baseline(-25000)  # subscript

r = p.add_run("O and E=mc")
r.set_font_size(1800)
r.set_font_color("333333")

r_sup = p.add_run("2")
r_sup.set_font_size(1200)
r_sup.set_font_color("C0504D")
r_sup.set_baseline(30000)  # superscript

# ─── Code-style text box ────────────────────────────────────────────

s3 = slide.add_shape("Code")
shape3 = slide.get_shape(s3)
shape3.set_geometry(int(6.5 * EMU), int(1.8 * EMU), int(5.5 * EMU), int(4.5 * EMU))
shape3.set_solid_fill_srgb("1E1E1E")
shape3.set_outline("333333", width_pt=1.0)
shape3.set_word_wrap(True)

lines = [
    ("from ", "CC7832"),
    ("offidized ", "A9B7C6"),
    ("import ", "CC7832"),
    ("Presentation\n", "A9B7C6"),
    ("\n", "A9B7C6"),
    ("pres ", "A9B7C6"),
    ("= ", "CC7832"),
    ("Presentation", "FFC66D"),
    ("()\n", "A9B7C6"),
    ("slide ", "A9B7C6"),
    ("= ", "CC7832"),
    ("pres", "A9B7C6"),
    (".", "CC7832"),
    ("get_slide", "FFC66D"),
    ("(", "A9B7C6"),
    ("0", "6897BB"),
    (")\n", "A9B7C6"),
    ("\n", "A9B7C6"),
    ("# ", "808080"),
    ("Effects!", "808080"),
    ("\n", "A9B7C6"),
    ("shape", "A9B7C6"),
    (".", "CC7832"),
    ("set_shadow", "FFC66D"),
    ("(...)", "A9B7C6"),
]

para = shape3.add_paragraph()
para.set_alignment("l")
for text, color in lines:
    r = para.add_run(text)
    r.set_font_size(1400)
    r.set_font_color(color)
    r.set_font_name("Consolas")

# ─── Hyperlink demo ─────────────────────────────────────────────────

s4 = slide.add_shape("Link")
shape4 = slide.get_shape(s4)
shape4.set_geometry(EMU, int(6.5 * EMU), 10 * EMU, int(0.5 * EMU))

para = shape4.add_paragraph()
para.set_alignment("ctr")
r = para.add_run("Built with Rust ")
r.set_font_size(1200)
r.set_font_color("808080")

r_link = para.add_run("\u2022 github.com/h0rv/offidized")
r_link.set_font_size(1200)
r_link.set_font_color("2F5597")
r_link.set_underline("sng")
r_link.set_hyperlink_url("https://github.com/h0rv/offidized")

pres.save("04_rich_text_formatting.pptx")
print("Created 04_rich_text_formatting.pptx")
