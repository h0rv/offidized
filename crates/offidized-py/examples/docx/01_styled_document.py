# /// script
# requires-python = ">=3.12"
# dependencies = []
# ///
"""Build a styled document with headings, paragraphs, and rich formatting.

Run from the offidized-py crate directory:
    cd crates/offidized-py && uv run python examples/docx/01_styled_document.py

Demonstrates paragraph styling, run-level formatting, numbered/bulleted
lists, spacing control, and indentation — all in one call per property
rather than python-docx's XML-level manipulation for many of these.
"""

from offidized import Document

doc = Document()

# ─── Document metadata ─────────────────────────────────────────────────
props = doc.document_properties()
props.set_title("offidized Document Demo")
props.set_creator("offidized Python bindings")
props.set_subject("Styled document creation")
props.set_description("Demonstrates rich document formatting with offidized")

# ─── Title ──────────────────────────────────────────────────────────────
title = doc.add_heading("offidized: Document Formatting", 1)

# ─── Intro paragraph with mixed formatting ──────────────────────────────
para = doc.add_paragraph("")
para.clear_style_id()
r1 = para.add_run("offidized ")
r1.set_bold(True)
r1.set_color("2F5597")
r1.set_font_size_half_points(28)
r1.set_font_family("Calibri")

r2 = para.add_run("gives you full control over Word documents from Python. ")
r2.set_font_size_half_points(24)
r2.set_font_family("Calibri")

r3 = para.add_run("Every property is one method call.")
r3.set_italic(True)
r3.set_font_size_half_points(24)
r3.set_font_family("Calibri")

# ─── Heading + paragraph spacing ────────────────────────────────────────
doc.add_heading("Paragraph Spacing Control", 2)

# python-docx requires XML hacking for precise spacing.
# offidized exposes it directly.
p1 = doc.add_paragraph("This paragraph has 240 twips (12pt) space before.")
p1.set_spacing_before_twips(240)

p2 = doc.add_paragraph(
    "This paragraph has 480 twips (24pt) space before and double line spacing."
)
p2.set_spacing_before_twips(480)
p2.set_line_spacing_twips(480)

p3 = doc.add_paragraph(
    "This paragraph has a first-line indent of 720 twips (0.5 inch)."
)
p3.set_indent_first_line_twips(720)

# ─── Bullet list ────────────────────────────────────────────────────────
doc.add_heading("Bullet Lists", 2)
items = [
    "Shadow, glow, and reflection effects",
    "Full theme color manipulation",
    "In-memory roundtrip (from_bytes / to_bytes)",
    "Slide transitions and speaker notes",
]
for item in items:
    doc.add_bulleted_paragraph(item)

# ─── Numbered list ──────────────────────────────────────────────────────
doc.add_heading("Numbered Steps", 2)
steps = [
    "Install: pip install offidized",
    "Create a Document()",
    "Add content with simple method calls",
    "Save or serialize to bytes",
]
for step in steps:
    doc.add_numbered_paragraph(step)

# ─── Run-level formatting showcase ──────────────────────────────────────
doc.add_heading("Text Formatting", 2)

# Bold, italic, underline, strikethrough
para = doc.add_paragraph("")
for label, setter in [
    ("Bold ", "bold"),
    ("Italic ", "italic"),
    ("Underline ", "underline"),
    ("Strikethrough ", "strikethrough"),
    ("Small Caps ", "small_caps"),
    ("All Caps ", "all_caps"),
]:
    r = para.add_run(label)
    r.set_font_size_half_points(24)
    getattr(r, f"set_{setter}")(True)

# Superscript and subscript
para2 = doc.add_paragraph("")
r = para2.add_run("H")
r.set_font_size_half_points(24)
r_sub = para2.add_run("2")
r_sub.set_font_size_half_points(20)
r_sub.set_subscript(True)
r_sub.set_color("C0504D")
r = para2.add_run("O is water, E=mc")
r.set_font_size_half_points(24)
r_sup = para2.add_run("2")
r_sup.set_font_size_half_points(20)
r_sup.set_superscript(True)
r_sup.set_color("C0504D")

# Highlight colors
para3 = doc.add_paragraph("")
for color in ["yellow", "green", "cyan", "magenta"]:
    r = para3.add_run(f" {color} ")
    r.set_highlight_color(color)
    r.set_font_size_half_points(24)

# ─── Page break control ────────────────────────────────────────────────
doc.add_heading("Page Flow Control", 2)

p = doc.add_paragraph("This paragraph forces a page break before it.")
p.set_page_break_before(True)

p = doc.add_paragraph("This paragraph keeps with the next (no orphan heading).")
p.set_keep_next(True)

p = doc.add_paragraph(
    "This paragraph keeps its lines together (no mid-paragraph break)."
)
p.set_keep_lines(True)

# ─── Save ───────────────────────────────────────────────────────────────
doc.save("01_styled_document.docx")
print("Created 01_styled_document.docx")
