# /// script
# requires-python = ">=3.12"
# dependencies = []
# ///
"""Build a complete contract/agreement document from scratch.

Run from the offidized-py crate directory:
    cd crates/offidized-py && uv run python examples/docx/07_contract.py

Demonstrates building a real multi-section document programmatically:
headers/footers, page setup, mixed formatting, tables, numbered clauses,
tab stops, paragraph shading, and document protection — features that
require extensive XML manipulation in python-docx.
"""

from offidized import Document

doc = Document()

# ─── Page setup ─────────────────────────────────────────────────────────
section = doc.section()
section.set_page_size_twips(12240, 15840)  # Letter
section.set_margin_top_twips(1440)
section.set_margin_bottom_twips(1440)
section.set_margin_left_twips(1800)
section.set_margin_right_twips(1440)

section.set_header_text("ACME Corp — Master Services Agreement")
section.set_footer_text("Confidential | Page 1")

# ─── Document properties ───────────────────────────────────────────────
props = doc.document_properties()
props.set_title("Master Services Agreement")
props.set_creator("ACME Corp Legal Department")
props.set_subject("Professional Services Contract")

# ─── Title block ────────────────────────────────────────────────────────
title = doc.add_heading("Master Services Agreement", 1)

# Parties paragraph with mixed formatting
para = doc.add_paragraph("")
r = para.add_run("This Master Services Agreement")
r.set_bold(True)
r.set_font_size_half_points(24)
r = para.add_run(' (the "Agreement") is entered into as of ')
r.set_font_size_half_points(24)
r = para.add_run("February 27, 2026")
r.set_bold(True)
r.set_font_size_half_points(24)
r = para.add_run(", by and between:")
r.set_font_size_half_points(24)

# ─── Parties table ──────────────────────────────────────────────────────
parties = doc.add_table(2, 2)
parties.set_width_twips(8640)
parties.set_column_widths_twips([4320, 4320])
parties.set_layout("fixed")

parties.set_cell_text(0, 0, "ACME Corporation")
parties.cell(0, 0).set_shading_color("1F3864")
parties.set_cell_text(0, 1, "Client Company, Inc.")
parties.cell(0, 1).set_shading_color("1F3864")

parties.set_cell_text(1, 0, "123 Innovation Drive\nSan Francisco, CA 94105")
parties.set_cell_text(1, 1, "456 Business Avenue\nNew York, NY 10001")

# ─── Numbered clauses ──────────────────────────────────────────────────
clauses = [
    (
        "Scope of Services",
        [
            "ACME shall provide professional consulting services as described in "
            "each Statement of Work (SOW) executed under this Agreement.",
            "Services shall be performed by qualified personnel with relevant "
            "expertise and industry certifications.",
            "ACME reserves the right to assign subcontractors subject to Client "
            "written approval, which shall not be unreasonably withheld.",
        ],
    ),
    (
        "Term and Termination",
        [
            "This Agreement shall commence on the Effective Date and continue "
            "for a period of twelve (12) months unless terminated earlier.",
            "Either party may terminate this Agreement for cause upon thirty (30) "
            "days written notice if the other party materially breaches any term.",
            "Upon termination, all outstanding invoices become immediately due "
            "and payable within fifteen (15) business days.",
        ],
    ),
    (
        "Compensation",
        [
            "Client shall pay ACME in accordance with the rates specified in each SOW.",
            "Invoices shall be submitted monthly and are due within thirty (30) days.",
            "Late payments shall accrue interest at 1.5% per month or the maximum "
            "rate permitted by law, whichever is less.",
        ],
    ),
    (
        "Confidentiality",
        [
            "Each party agrees to maintain the confidentiality of all proprietary "
            "information received from the other party.",
            "Confidential information shall not be disclosed to third parties "
            "without prior written consent.",
            "This obligation survives termination for a period of three (3) years.",
        ],
    ),
    (
        "Limitation of Liability",
        [
            "Neither party shall be liable for indirect, incidental, or consequential "
            "damages arising from this Agreement.",
            "ACME's total liability shall not exceed the fees paid by Client in the "
            "twelve (12) months preceding the claim.",
        ],
    ),
]

for clause_num, (title, paragraphs) in enumerate(clauses, start=1):
    heading = doc.add_heading(f"{clause_num}. {title}", 2)

    for i, text in enumerate(paragraphs, start=1):
        p = doc.add_paragraph(f"{clause_num}.{i}  {text}")
        p.set_indent_left_twips(720)  # 0.5 inch indent
        p.set_indent_hanging_twips(360)  # hanging indent for the number
        p.set_spacing_after_twips(120)

# ─── Signature block ───────────────────────────────────────────────────
doc.add_heading("Signatures", 2)

sig_table = doc.add_table(4, 2)
sig_table.set_width_twips(8640)
sig_table.set_column_widths_twips([4320, 4320])

sig_table.set_cell_text(0, 0, "ACME Corporation")
sig_table.cell(0, 0).set_shading_color("F2F2F2")
sig_table.set_cell_text(0, 1, "Client Company, Inc.")
sig_table.cell(0, 1).set_shading_color("F2F2F2")

sig_table.set_cell_text(1, 0, "Signature: ________________________")
sig_table.set_cell_text(1, 1, "Signature: ________________________")

sig_table.set_cell_text(2, 0, "Name: Jane Smith")
sig_table.set_cell_text(2, 1, "Name: ________________________")

sig_table.set_cell_text(3, 0, "Date: ________________________")
sig_table.set_cell_text(3, 1, "Date: ________________________")

# ─── Highlight box (shaded paragraph) ──────────────────────────────────
notice = doc.add_paragraph(
    "NOTICE: This Agreement constitutes the entire understanding between the "
    "parties. No modifications shall be effective unless in writing and signed "
    "by authorized representatives of both parties."
)
notice.set_shading_color("FFF2CC")
notice.set_spacing_before_twips(480)

# ─── Tab stops ──────────────────────────────────────────────────────────
# python-docx has no tab stop API. offidized exposes them directly.
ref = doc.add_paragraph("Document Reference:")
ref.set_spacing_before_twips(240)

ref_line = doc.add_paragraph("Agreement No.\tMSA-2026-001\tVersion\t3.0")
ref_line.add_tab_stop(2880, "left")  # 2 inches
ref_line.add_tab_stop(5760, "left")  # 4 inches
ref_line.add_tab_stop(7920, "left")  # 5.5 inches

# ─── Protect the final document ────────────────────────────────────────
doc.set_protection("readOnly", True)

print(f"Built contract: {doc.paragraph_count()} paragraphs, {doc.table_count()} tables")
doc.save("07_contract.docx")
print("Created 07_contract.docx")
