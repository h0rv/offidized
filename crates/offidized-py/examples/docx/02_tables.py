# /// script
# requires-python = ">=3.12"
# dependencies = []
# ///
"""Tables with cell merging, styling, column widths, and layout control.

Run from the offidized-py crate directory:
    cd crates/offidized-py && uv run python examples/docx/02_tables.py

python-docx can create tables but lacks cell merge (vertical), table width
control, layout mode switching, and per-cell shading in a clean API.
offidized provides all of these as single method calls.
"""

from offidized import Document

doc = Document()

doc.add_heading("Table Features", 1)

# ─── Basic styled table ────────────────────────────────────────────────
doc.add_heading("Sales by Region", 2)

table = doc.add_table(5, 4)
table.set_alignment("center")
table.set_width_twips(9000)
table.set_column_widths_twips([3000, 2000, 2000, 2000])
table.set_layout("fixed")

# Headers
headers = ["Region", "Q1", "Q2", "Q3"]
for col, header in enumerate(headers):
    table.set_cell_text(0, col, header)
    cell = table.cell(0, col)
    cell.set_shading_color("2F5597")
    cell.set_vertical_alignment("center")

# Data
data = [
    ("North America", "$12.5M", "$15.3M", "$18.7M"),
    ("EMEA", "$8.2M", "$9.5M", "$11.3M"),
    ("APAC", "$5.1M", "$7.8M", "$9.2M"),
    ("Total", "$25.8M", "$32.6M", "$39.2M"),
]
for row, (region, *values) in enumerate(data, start=1):
    table.set_cell_text(row, 0, region)
    for col, val in enumerate(values, start=1):
        table.set_cell_text(row, col, val)

# Style the total row
for col in range(4):
    cell = table.cell(4, col)
    cell.set_shading_color("D6E4F0")

# ─── Table with vertical merge ─────────────────────────────────────────
doc.add_heading("Vertical Cell Merge", 2)
p = doc.add_paragraph(
    "python-docx has no API for vertical cell merging. "
    "offidized provides set_vertical_merge() directly."
)
p.set_spacing_after_twips(240)

table2 = doc.add_table(4, 3)
table2.set_width_twips(8000)
table2.set_column_widths_twips([2500, 2500, 3000])

table2.set_cell_text(0, 0, "Category")
table2.set_cell_text(0, 1, "Sub-item")
table2.set_cell_text(0, 2, "Value")

# "Hardware" spans rows 1-2
table2.set_cell_text(1, 0, "Hardware")
table2.cell(1, 0).set_vertical_merge("restart")
table2.set_cell_text(2, 0, "")
table2.cell(2, 0).set_vertical_merge("continue")

table2.set_cell_text(1, 1, "Servers")
table2.set_cell_text(1, 2, "$45,000")
table2.set_cell_text(2, 1, "Networking")
table2.set_cell_text(2, 2, "$12,000")

# "Software" in row 3
table2.set_cell_text(3, 0, "Software")
table2.set_cell_text(3, 1, "Licenses")
table2.set_cell_text(3, 2, "$28,000")

# Header shading
for col in range(3):
    table2.cell(0, col).set_shading_color("1F3864")

# ─── Table with horizontal merge ───────────────────────────────────────
doc.add_heading("Horizontal Cell Merge", 2)

table3 = doc.add_table(3, 4)
table3.set_width_twips(8000)

# Merge top row across all 4 columns for a banner
table3.merge_cells_horizontally(0, 0, 3)
table3.set_cell_text(0, 0, "Quarterly Budget Summary")
table3.cell(0, 0).set_shading_color("548235")

table3.set_cell_text(1, 0, "Department")
table3.set_cell_text(1, 1, "Budget")
table3.set_cell_text(1, 2, "Spent")
table3.set_cell_text(1, 3, "Remaining")

table3.set_cell_text(2, 0, "Engineering")
table3.set_cell_text(2, 1, "$500K")
table3.set_cell_text(2, 2, "$320K")
table3.set_cell_text(2, 3, "$180K")

# ─── Dynamic table: add/insert/remove rows ─────────────────────────────
doc.add_heading("Dynamic Row Manipulation", 2)
p = doc.add_paragraph(
    "offidized supports add_row(), insert_row(), and remove_row() "
    "for dynamic table construction."
)
p.set_spacing_after_twips(240)

table4 = doc.add_table(2, 2)
table4.set_cell_text(0, 0, "Name")
table4.set_cell_text(0, 1, "Role")
table4.set_cell_text(1, 0, "Alice")
table4.set_cell_text(1, 1, "Engineer")

# Add more rows dynamically
table4.add_row()
table4.set_cell_text(2, 0, "Bob")
table4.set_cell_text(2, 1, "Designer")

table4.add_row()
table4.set_cell_text(3, 0, "Charlie")
table4.set_cell_text(3, 1, "PM")

# Insert a row at position 1 (after header)
table4.insert_row(1)
table4.set_cell_text(1, 0, "** TEAM LEAD **")
table4.set_cell_text(1, 1, "Zara")

doc.save("02_tables.docx")
print("Created 02_tables.docx")
