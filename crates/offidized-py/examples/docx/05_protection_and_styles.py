# /// script
# requires-python = ">=3.12"
# dependencies = []
# ///
"""Document protection and custom style registry.

Run from the offidized-py crate directory:
    cd crates/offidized-py && uv run python examples/docx/05_protection_and_styles.py

python-docx has limited protection support (read-only, no granular control)
and minimal style management. offidized lets you set protection mode,
register custom styles, and query the full style registry.
"""

from offidized import Document

doc = Document()

doc.add_heading("Protection & Styles", 1)

# ─── Document protection ───────────────────────────────────────────────
doc.add_heading("Document Protection", 2)

doc.add_paragraph(
    "offidized supports four protection modes: readOnly, comments, "
    "trackedChanges, and forms. python-docx can only read protection "
    "state, not set it."
)

# Set protection — only allow comments
doc.set_protection("comments", True)

prot = doc.protection()
if prot:
    edit_type, enforced = prot
    doc.add_bulleted_paragraph(f"Protection mode: {edit_type}")
    doc.add_bulleted_paragraph(f"Enforced: {enforced}")

# Clear and set a different mode
doc.clear_protection()
doc.set_protection("readOnly", True)

doc.add_paragraph("Protection changed to readOnly mode.")

# Clear for the rest of the demo
doc.clear_protection()

# ─── Style registry ────────────────────────────────────────────────────
doc.add_heading("Style Registry", 2)

doc.add_paragraph(
    "offidized exposes the full style registry. You can list all styles, "
    "add paragraph/character/table styles, and apply them by ID."
)

# List built-in styles
doc.add_paragraph(f"Built-in styles: {doc.style_count()}")

# Add custom styles
doc.add_paragraph_style("CustomHighlight")
doc.add_character_style("InlineCode")
doc.add_table_style("CustomTable")

doc.add_paragraph(f"After adding 3 custom styles: {doc.style_count()}")

# Show all styles
doc.add_heading("All Registered Styles", 3)
for kind, style_id, name in doc.styles():
    display = f"{style_id}"
    if name:
        display += f" ({name})"
    doc.add_bulleted_paragraph(f"[{kind}] {display}")

# ─── Apply styles ──────────────────────────────────────────────────────
doc.add_heading("Applying Styles", 2)

p = doc.add_paragraph_with_style("This paragraph uses Heading2 style.", "Heading2")
doc.add_paragraph(f"Style ID: {p.style_id()}")

p2 = doc.add_paragraph("This paragraph has a custom style applied.")
p2.set_style_id("CustomHighlight")
doc.add_paragraph(f"Style ID: {p2.style_id()}")

# Run-level style
p3 = doc.add_paragraph("")
r = p3.add_run("This run has a character style applied.")
r.set_style_id("InlineCode")
doc.add_paragraph(f"Run style ID: {r.style_id()}")

doc.save("05_protection_and_styles.docx")
print("Created 05_protection_and_styles.docx")
