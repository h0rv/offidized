# /// script
# requires-python = ">=3.12"
# dependencies = []
# ///
"""In-memory roundtrip: create, serialize to bytes, reload, and modify.

Run from the offidized-py crate directory:
    cd crates/offidized-py && uv run python examples/docx/06_in_memory_roundtrip.py

offidized supports full bytes roundtrip (from_bytes / to_bytes) for docx.
This enables serverless workflows, streaming pipelines, and testing
without touching the filesystem.
"""

from offidized import Document

# ─── Step 1: Create a document in memory ────────────────────────────────

doc = Document()

doc.add_heading("Version 1", 1)
doc.add_paragraph("This document was created entirely in memory.")
doc.add_paragraph("It has never touched the filesystem.")
doc.add_bulleted_paragraph("Created with offidized")
doc.add_bulleted_paragraph("Serialized to bytes")
doc.add_bulleted_paragraph("Reloaded and modified")

# ─── Step 2: Serialize to bytes (no filesystem) ─────────────────────────

docx_bytes = doc.to_bytes()
print(f"Serialized to {len(docx_bytes):,} bytes")

# ─── Step 3: Reload from bytes and modify ────────────────────────────────

doc2 = Document.from_bytes(docx_bytes)
print(f"Reloaded: {doc2.paragraph_count()} paragraphs")

# Modify the heading
heading = doc2.paragraph(0)
run = heading.run(0)
run.set_text("Version 2 (Modified)")

# Add new content
doc2.add_heading("Added in Round 2", 2)
doc2.add_paragraph("This section was added after the in-memory roundtrip.")
doc2.add_numbered_paragraph("Roundtrip fidelity preserved")
doc2.add_numbered_paragraph("New content appended cleanly")

# ─── Step 4: Serialize again ─────────────────────────────────────────────

docx_bytes2 = doc2.to_bytes()
print(
    f"Re-serialized to {len(docx_bytes2):,} bytes ({doc2.paragraph_count()} paragraphs)"
)

# ─── Step 5: Verify the full roundtrip ────────────────────────────────────

doc3 = Document.from_bytes(docx_bytes2)

heading_text = doc3.paragraph(0).text()
assert "Version 2" in heading_text
print(f"Verified: heading = '{heading_text}'")

# Check the added content survived
found = False
for i in range(doc3.paragraph_count()):
    if "Round 2" in doc3.paragraph(i).text():
        found = True
        break
assert found, "Added content should survive roundtrip"
print("Verified: 'Round 2' content survived roundtrip")
print(f"Total paragraphs: {doc3.paragraph_count()}")

# Save final version
doc3.save("06_in_memory_roundtrip.docx")
print("Created 06_in_memory_roundtrip.docx")
