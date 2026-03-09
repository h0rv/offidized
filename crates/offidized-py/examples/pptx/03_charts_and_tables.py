"""Charts with multiple series and tables with merged cells.

python-pptx has limited chart customization (no axis titles, no series
management after creation) and no table cell merging. offidized provides
full control over chart series/axes and table structure.
"""

from offidized import Presentation

EMU = 914400

pres = Presentation()
pres.set_slide_size(12192000, 6858000)

# ─── Slide 1: Chart with multiple series ────────────────────────────

slide1_idx = pres.add_slide()
slide1 = pres.get_slide(slide1_idx)

# Add title shape
s = slide1.add_shape("Chart Title")
title_shape = slide1.get_shape(s)
title_shape.set_geometry(EMU, int(0.3 * EMU), 10 * EMU, EMU)
para = title_shape.add_paragraph_with_text("Quarterly Revenue by Region")
run = para.get_run(0)
run.set_font_size(2800)
run.set_bold(True)
para.set_alignment("ctr")

# Add chart
chart_idx = slide1.add_chart("bar")
chart = slide1.get_chart(chart_idx)
chart.set_title("Revenue ($M)")

# Add initial data points (Q1-Q4 for first series)
for q, val in [("Q1", 12.5), ("Q2", 15.3), ("Q3", 18.7), ("Q4", 22.1)]:
    chart.add_data_point(q, val)

# Add additional series (not easily possible in python-pptx after creation)
chart.add_series("EMEA", [8.2, 9.5, 11.3, 14.0])
chart.add_series("APAC", [5.1, 7.8, 9.2, 12.5])

# Set axis titles (not available in python-pptx)
chart.set_category_axis_title("Quarter")
chart.set_value_axis_title("Revenue ($M)")

chart.set_show_legend(True)

# ─── Slide 2: Table with formatting ─────────────────────────────────

slide2_idx = pres.add_slide()
slide2 = pres.get_slide(slide2_idx)

# Title
s = slide2.add_shape("Table Title")
title_shape = slide2.get_shape(s)
title_shape.set_geometry(EMU, int(0.3 * EMU), 10 * EMU, EMU)
para = title_shape.add_paragraph_with_text("Sales Dashboard")
run = para.get_run(0)
run.set_font_size(2800)
run.set_bold(True)
para.set_alignment("ctr")

# Add table
table_idx = slide2.add_table(5, 4)
table = slide2.get_table(table_idx)
table.set_geometry(EMU, int(1.5 * EMU), 10 * EMU, int(4.5 * EMU))

# Headers
headers = ["Region", "Q1", "Q2", "Q3"]
for col, header in enumerate(headers):
    table.set_cell_text(0, col, header)
    table.set_cell_fill(0, col, "2F5597")
    table.set_cell_bold(0, col, True)
    table.set_cell_font_color(0, col, "FFFFFF")
    table.set_cell_font_size(0, col, 1400)

# Data
data = [
    ("North America", "$12.5M", "$15.3M", "$18.7M"),
    ("EMEA", "$8.2M", "$9.5M", "$11.3M"),
    ("APAC", "$5.1M", "$7.8M", "$9.2M"),
    ("Total", "$25.8M", "$32.6M", "$39.2M"),
]
for row, (region, *values) in enumerate(data, start=1):
    table.set_cell_text(row, 0, region)
    table.set_cell_bold(row, 0, True)
    table.set_cell_font_size(row, 0, 1200)
    for col, val in enumerate(values, start=1):
        table.set_cell_text(row, col, val)
        table.set_cell_font_size(row, col, 1200)

# Style the total row
for col in range(4):
    table.set_cell_fill(4, col, "D6E4F0")
    table.set_cell_bold(4, col, True)

# Merge cells for a header row (not possible in python-pptx)
table.merge_cells(4, 0, 4, 0)  # Keep total label

# Set column widths
table.set_column_width(0, int(3 * EMU))
for col in range(1, 4):
    table.set_column_width(col, int(2.3 * EMU))

pres.save("03_charts_and_tables.pptx")
print("Created 03_charts_and_tables.pptx")
