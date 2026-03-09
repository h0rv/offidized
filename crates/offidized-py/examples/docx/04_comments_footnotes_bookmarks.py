# /// script
# requires-python = ">=3.12"
# dependencies = []
# ///
"""Comments, footnotes, endnotes, and bookmarks.

Run from the offidized-py crate directory:
    cd crates/offidized-py && uv run python examples/docx/04_comments_footnotes_bookmarks.py

python-docx has limited or no support for comments, footnotes, endnotes,
and bookmarks. offidized provides full create/read APIs for all of these.
"""

from offidized import Document

doc = Document()

doc.add_heading("Annotations & References", 1)

# ─── Comments ───────────────────────────────────────────────────────────
doc.add_heading("Comments", 2)

doc.add_paragraph(
    "This document demonstrates comment, footnote, endnote, and bookmark "
    "support. python-docx cannot create comments at all — offidized can."
)

# Add comments (id, author, text)
doc.add_comment(1, "Reviewer A", "Consider rephrasing this section.")
doc.add_comment(2, "Reviewer B", "Data needs to be verified against Q4 actuals.")
doc.add_comment(3, "Legal", "Approved for external distribution.")

p = doc.add_paragraph(f"This document has {doc.comment_count()} comments attached.")

# Read them back
for cid, author, text in doc.comments():
    doc.add_bulleted_paragraph(f"[{author}] {text}")

# ─── Footnotes ──────────────────────────────────────────────────────────
doc.add_heading("Footnotes", 2)

doc.add_paragraph(
    "Footnotes appear at the bottom of the page. "
    "python-docx can read footnotes but has no clean API to create them."
)

doc.add_footnote(1, "Source: Annual Report 2025, page 42.")
doc.add_footnote(2, "All figures are in USD unless otherwise noted.")
doc.add_footnote(3, "Adjusted for inflation using CPI-U index.")

p = doc.add_paragraph(f"Added {doc.footnote_count()} footnotes to this document.")

for fid, text in doc.footnotes():
    doc.add_bulleted_paragraph(f"Footnote {fid}: {text}")

# ─── Endnotes ───────────────────────────────────────────────────────────
doc.add_heading("Endnotes", 2)

doc.add_paragraph(
    "Endnotes appear at the end of the document. "
    "python-docx has no endnote creation support at all."
)

doc.add_endnote(1, "See Appendix A for full methodology.")
doc.add_endnote(2, "Data collected between Jan 2025 and Dec 2025.")

p = doc.add_paragraph(f"Added {doc.endnote_count()} endnotes.")

for eid, text in doc.endnotes():
    doc.add_bulleted_paragraph(f"Endnote {eid}: {text}")

# ─── Bookmarks ──────────────────────────────────────────────────────────
doc.add_heading("Bookmarks", 2)

doc.add_paragraph(
    "Bookmarks mark ranges of content for cross-references and navigation. "
    "python-docx has no bookmark creation API."
)

# Bookmarks reference paragraph indices
doc.add_bookmark(1, "introduction", 0, 0)
doc.add_bookmark(2, "comments_section", 2, 5)
doc.add_bookmark(3, "conclusion", 0, 0)

p = doc.add_paragraph(f"Added {doc.bookmark_count()} bookmarks:")
for bid, name in doc.bookmarks():
    doc.add_bulleted_paragraph(f"Bookmark '{name}' (id={bid})")

# ─── Summary ────────────────────────────────────────────────────────────
doc.add_heading("Summary", 2)

doc.add_paragraph(
    f"This document contains {doc.comment_count()} comments, "
    f"{doc.footnote_count()} footnotes, {doc.endnote_count()} endnotes, "
    f"and {doc.bookmark_count()} bookmarks — all created programmatically "
    f"with simple method calls."
)

doc.save("04_comments_footnotes_bookmarks.docx")
print("Created 04_comments_footnotes_bookmarks.docx")
