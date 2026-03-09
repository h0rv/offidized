# /// script
# requires-python = ">=3.12"
# dependencies = []
# ///
"""
Excel features that go beyond what openpyxl can do.

Run from the offidized-py crate directory:
    cd crates/offidized-py && uv run python examples/xlsx/beyond_openpyxl.py

Each example demonstrates a capability that either doesn't exist in openpyxl
or requires significantly more code / breaks roundtrip fidelity.
"""

from offidized import (
    Workbook,
    XlsxStyle,
    XlsxFont,
    XlsxFill,
    XlsxBorder,
    XlsxAlignment,
    XlsxWorksheetTable,
    XlsxTableColumn,
    XlsxDataValidation,
    XlsxComment,
    XlsxSparklineGroup,
    XlsxSparkline,
    XlsxConditionalFormatting,
    XlsxChart,
    XlsxChartSeries,
    XlsxChartDataRef,
    XlsxPageSetup,
    XlsxPageMargins,
    XlsxSheetViewOptions,
    XlsxSheetProtection,
    XlsxWorkbookProtection,
    XlsxPrintHeaderFooter,
    ir_derive,
    ir_apply,
)
import tempfile
import os

OUT = tempfile.mkdtemp(prefix="offidized_examples_")
print(f"Output directory: {OUT}\n")


# ─────────────────────────────────────────────────────────────────────────────
# 1. ROUNDTRIP FIDELITY — the killer feature
# ─────────────────────────────────────────────────────────────────────────────
# openpyxl destroys charts, sparklines, and other features on save.
# offidized preserves everything it doesn't explicitly modify.

print("1. Roundtrip fidelity")
print("   openpyxl: Opens file with chart, saves → chart is gone")
print("   offidized: Opens file with chart, saves → chart survives")
print()


# ─────────────────────────────────────────────────────────────────────────────
# 2. STRUCTURED TABLES with totals rows
# ─────────────────────────────────────────────────────────────────────────────
# openpyxl can create basic tables but totals row support is limited.

print("2. Structured tables with totals rows")

wb = Workbook()
ws = wb.add_sheet("Sales")

# Header
for i, h in enumerate(["Product", "Q1", "Q2", "Q3", "Q4"]):
    ws.set_cell_value(f"{chr(65 + i)}1", h)

# Data
data = [
    ("Widgets", 1200, 1500, 1800, 2100),
    ("Gadgets", 800, 950, 1100, 1300),
    ("Gizmos", 500, 600, 750, 900),
]
for r, (product, *quarters) in enumerate(data, start=2):
    ws.set_cell_value(f"A{r}", product)
    for c, val in enumerate(quarters):
        ws.set_cell_value(f"{chr(66 + c)}{r}", val)

# Create a structured table with totals row
tbl = XlsxWorksheetTable("SalesTable", "A1:E4")
tbl.show_totals_row = True
tbl.style_name = "TableStyleMedium9"

# Define columns with totals functions
cols = [
    ("Product", None, "Total"),
    ("Q1", "sum", None),
    ("Q2", "sum", None),
    ("Q3", "sum", None),
    ("Q4", "sum", None),
]
for name, formula, label in cols:
    col = XlsxTableColumn(name, 0)
    if formula:
        col.totals_row_formula = formula
    if label:
        col.totals_row_label = label
    tbl.add_column(col)

ws.add_table(tbl)

path = os.path.join(OUT, "02_tables.xlsx")
wb.save(path)
print(f"   Saved: {path}\n")


# ─────────────────────────────────────────────────────────────────────────────
# 3. DATA VALIDATION — dropdown lists, ranges, custom formulas
# ─────────────────────────────────────────────────────────────────────────────

print("3. Data validation (dropdown, numeric range, custom formula)")

wb = Workbook()
ws = wb.add_sheet("Form")

ws.set_cell_value("A1", "Status")
ws.set_cell_value("A2", "Score")
ws.set_cell_value("A3", "Start Date")

# Dropdown list
dv_list = XlsxDataValidation.list(["B1"], '"Open,In Progress,Closed,Deferred"')
dv_list.prompt_title = "Pick status"
dv_list.prompt_message = "Choose from the dropdown"
dv_list.show_input_message = True
ws.add_data_validation(dv_list)

# Numeric range (1-100)
dv_range = XlsxDataValidation.whole(["B2"], "1")
dv_range.formula2 = "100"
dv_range.error_title = "Invalid score"
dv_range.error_message = "Score must be between 1 and 100"
dv_range.show_error_message = True
ws.add_data_validation(dv_range)

# Date validation
dv_date = XlsxDataValidation.date(["B3"], "2024-01-01")
dv_date.formula2 = "2025-12-31"
ws.add_data_validation(dv_date)

# Custom formula validation
dv_custom = XlsxDataValidation.custom(["C1:C100"], "LEN(C1)<=50")
dv_custom.error_message = "Max 50 characters"
ws.add_data_validation(dv_custom)

path = os.path.join(OUT, "03_data_validation.xlsx")
wb.save(path)
print(f"   Saved: {path}\n")


# ─────────────────────────────────────────────────────────────────────────────
# 4. CONDITIONAL FORMATTING — cell rules, expression rules
# ─────────────────────────────────────────────────────────────────────────────

print("4. Conditional formatting")

wb = Workbook()
ws = wb.add_sheet("Dashboard")

# Sample data
ws.set_cell_value("A1", "Metric")
ws.set_cell_value("B1", "Value")
for i, (metric, val) in enumerate(
    [
        ("Revenue", 95000),
        ("Costs", 62000),
        ("Profit", 33000),
        ("Growth", 12),
        ("Churn", 3),
        ("NPS", 72),
    ],
    start=2,
):
    ws.set_cell_value(f"A{i}", metric)
    ws.set_cell_value(f"B{i}", val)

# Highlight cells above 30000 — "cellIs" with greaterThan
cf = XlsxConditionalFormatting("cellIs", ["B2:B7"], ["30000"])
ws.add_conditional_formatting(cf)

# Expression-based rule — highlight rows where value > 50000
cf2 = XlsxConditionalFormatting("expression", ["A2:B7"], ["$B2>50000"])
ws.add_conditional_formatting(cf2)

path = os.path.join(OUT, "04_conditional_formatting.xlsx")
wb.save(path)
print(f"   Saved: {path}\n")


# ─────────────────────────────────────────────────────────────────────────────
# 5. SPARKLINES — inline charts in cells
# ─────────────────────────────────────────────────────────────────────────────
# openpyxl has zero sparkline support. offidized creates them natively.

print("5. Sparklines (openpyxl cannot create these at all)")

wb = Workbook()
ws = wb.add_sheet("Trends")

ws.set_cell_value("A1", "Region")
ws.set_cell_value("B1", "Jan")
ws.set_cell_value("C1", "Feb")
ws.set_cell_value("D1", "Mar")
ws.set_cell_value("E1", "Apr")
ws.set_cell_value("F1", "May")
ws.set_cell_value("G1", "Jun")
ws.set_cell_value("H1", "Trend")

regions = [
    ("North", [100, 120, 115, 130, 145, 160]),
    ("South", [80, 75, 90, 85, 95, 110]),
    ("East", [60, 70, 65, 80, 75, 90]),
    ("West", [110, 100, 95, 105, 115, 125]),
]

for r, (region, values) in enumerate(regions, start=2):
    ws.set_cell_value(f"A{r}", region)
    for c, v in enumerate(values):
        ws.set_cell_value(f"{chr(66 + c)}{r}", v)

# Create sparkline group — one sparkline per region in column H
group = XlsxSparklineGroup()
group.sparkline_type = "line"
group.high_point = True
group.low_point = True
group.markers = True
group.line_weight = 1.5

for r in range(2, 6):
    group.add_sparkline(XlsxSparkline(f"H{r}", f"Trends!$B${r}:$G${r}"))

ws.add_sparkline_group(group)

path = os.path.join(OUT, "05_sparklines.xlsx")
wb.save(path)
print(f"   Saved: {path}\n")


# ─────────────────────────────────────────────────────────────────────────────
# 6. CHARTS — embedded in worksheets
# ─────────────────────────────────────────────────────────────────────────────

print("6. Charts")

wb = Workbook()
ws = wb.add_sheet("Revenue")

ws.set_cell_value("A1", "Month")
ws.set_cell_value("B1", "Revenue")
months = ["Jan", "Feb", "Mar", "Apr", "May", "Jun"]
values = [42000, 48000, 45000, 52000, 58000, 63000]
for i, (m, v) in enumerate(zip(months, values), start=2):
    ws.set_cell_value(f"A{i}", m)
    ws.set_cell_value(f"B{i}", v)

chart = XlsxChart("bar")
chart.title = "Monthly Revenue"
series = XlsxChartSeries(0, 0)
series.set_categories(XlsxChartDataRef.from_formula("Revenue!$A$2:$A$7"))
series.set_values(XlsxChartDataRef.from_formula("Revenue!$B$2:$B$7"))
chart.add_series(series)
ws.add_chart(chart)

path = os.path.join(OUT, "06_charts.xlsx")
wb.save(path)
print(f"   Saved: {path}\n")


# ─────────────────────────────────────────────────────────────────────────────
# 7. FULL STYLE SYSTEM — fonts, fills, borders, alignment, number formats
# ─────────────────────────────────────────────────────────────────────────────

print("7. Full style system")

wb = Workbook()
ws = wb.add_sheet("Styled")

# Create styles
header_font = XlsxFont()
header_font.bold = True
header_font.size = "14"
header_font.color = "FFFFFF"

header_fill = XlsxFill()
header_fill.pattern = "solid"
header_fill.foreground_color = "2F5496"

header_border = XlsxBorder()
header_border.bottom_style = "thick"
header_border.bottom_color = "1F3864"

header_align = XlsxAlignment()
header_align.horizontal = "center"
header_align.vertical = "center"

header_style = XlsxStyle()
header_style.set_font(header_font)
header_style.set_fill(header_fill)
header_style.set_border(header_border)
header_style.set_alignment(header_align)
header_id = wb.add_style(header_style)

# Currency style
currency_style = XlsxStyle()
currency_style.set_number_format("$#,##0.00")
currency_font = XlsxFont()
currency_font.name = "Consolas"
currency_style.set_font(currency_font)
currency_id = wb.add_style(currency_style)

# Percent style
pct_style = XlsxStyle()
pct_style.set_number_format("0.0%")
pct_id = wb.add_style(pct_style)

# Apply
headers = ["Item", "Amount", "Change"]
for i, h in enumerate(headers):
    ref = f"{chr(65 + i)}1"
    ws.set_cell_value(ref, h)
    cell = ws.cell(ref)
    cell.set_style_id(header_id)

items = [
    ("Revenue", 125000.50, 0.12),
    ("Costs", 87500.75, -0.03),
    ("Profit", 37499.75, 0.28),
]
for r, (item, amount, change) in enumerate(items, start=2):
    ws.set_cell_value(f"A{r}", item)
    ws.set_cell_value(f"B{r}", amount)
    ws.cell(f"B{r}").set_style_id(currency_id)
    ws.set_cell_value(f"C{r}", change)
    ws.cell(f"C{r}").set_style_id(pct_id)

path = os.path.join(OUT, "07_styles.xlsx")
wb.save(path)
print(f"   Saved: {path}\n")


# ─────────────────────────────────────────────────────────────────────────────
# 8. COMMENTS, HYPERLINKS, MERGED CELLS, FREEZE PANES
# ─────────────────────────────────────────────────────────────────────────────

print("8. Comments, hyperlinks, merged cells, freeze panes")

wb = Workbook()
ws = wb.add_sheet("Mixed")

# Merged header
ws.add_merged_range("A1:D1")
ws.set_cell_value("A1", "Quarterly Report — 2025")

# Freeze below header
ws.freeze_panes(0, 2)

# Hyperlink
ws.set_cell_value("A3", "See our website")
ws.add_hyperlink("A3", "https://sweetspot.so")

# Comment
cmt = XlsxComment("B3", "Analyst", "This value is estimated", False)
ws.add_comment(cmt)
ws.set_cell_value("B3", 42000)

# Auto-filter
ws.set_cell_value("A5", "Name")
ws.set_cell_value("B5", "Score")
ws.set_cell_value("C5", "Grade")
ws.set_auto_filter("A5:C5")

path = os.path.join(OUT, "08_mixed_features.xlsx")
wb.save(path)
print(f"   Saved: {path}\n")


# ─────────────────────────────────────────────────────────────────────────────
# 9. WORKBOOK LINT — catch errors before sending to Excel
# ─────────────────────────────────────────────────────────────────────────────
# No equivalent in openpyxl. offidized can lint workbooks for common issues.

print("9. Workbook lint (unique to offidized)")

wb = Workbook()
ws = wb.add_sheet("Test")
ws.set_cell_value("A1", "Hello")
ws.set_cell_formula("B1", "SUM(A1:A10)")

report = wb.lint([])
print(
    f"   Lint report: {report.error_count()} errors, {report.warning_count()} warnings"
)
print(f"   Is clean: {report.is_clean()}")
for f in report.findings():
    print(f"   [{f.severity()}] {f.code()}: {f.message()}")
print()


# ─────────────────────────────────────────────────────────────────────────────
# 10. PAGE SETUP + PRINT SETTINGS — production-ready output
# ─────────────────────────────────────────────────────────────────────────────

print("10. Page setup and print settings")

wb = Workbook()
ws = wb.add_sheet("Print Ready")

# Build a small report so the print settings have something to show
ws.add_merged_range("A1:F1")
ws.set_cell_value("A1", "Sweetspot — Quarterly Revenue Report")

col_headers = ["Region", "Q1", "Q2", "Q3", "Q4", "Total"]
for i, h in enumerate(col_headers):
    ws.set_cell_value(f"{chr(65 + i)}3", h)

rows = [
    ("Northeast", 120000, 135000, 142000, 158000),
    ("Southeast", 98000, 102000, 110000, 118000),
    ("Midwest", 87000, 91000, 95000, 103000),
    ("West", 145000, 152000, 161000, 175000),
    ("International", 63000, 72000, 81000, 94000),
]
for r, (region, *quarters) in enumerate(rows, start=4):
    ws.set_cell_value(f"A{r}", region)
    for c, val in enumerate(quarters):
        ws.set_cell_value(f"{chr(66 + c)}{r}", val)
    ws.set_cell_formula(f"F{r}", f"SUM(B{r}:E{r})")

# Totals row
ws.set_cell_value("A9", "Total")
for c in range(5):
    col = chr(66 + c)
    ws.set_cell_formula(f"{col}9", f"SUM({col}4:{col}8)")

# Page setup (properties)
setup = XlsxPageSetup()
setup.orientation = "landscape"
setup.paper_size = 1  # Letter
setup.fit_to_width = 1
setup.fit_to_height = 0  # as many pages tall as needed
ws.set_page_setup(setup)

# Margins (properties)
margins = XlsxPageMargins()
margins.top = 0.75
margins.bottom = 0.75
margins.left = 0.5
margins.right = 0.5
ws.set_page_margins(margins)

# Header/footer (properties)
hf = XlsxPrintHeaderFooter()
hf.odd_header = "&C&B Sweetspot Quarterly Report"
hf.odd_footer = "&L&D &T&RPage &P of &N"
ws.set_header_footer(hf)

# Print area
ws.set_print_area("A1:F50")

# Sheet view (properties)
view = XlsxSheetViewOptions()
view.show_gridlines = False
view.zoom_scale = 85
ws.set_sheet_view_options(view)

path = os.path.join(OUT, "10_page_setup.xlsx")
wb.save(path)
print(f"   Saved: {path}\n")


# ─────────────────────────────────────────────────────────────────────────────
# 11. SHEET + WORKBOOK PROTECTION
# ─────────────────────────────────────────────────────────────────────────────

print("11. Sheet and workbook protection")

wb = Workbook()
ws = wb.add_sheet("Protected")
ws.set_cell_value("A1", "This sheet is protected")

# Detailed sheet protection — lock everything except specific actions
prot = XlsxSheetProtection()
prot.sheet = True
prot.format_cells = False  # allow formatting
prot.sort = False  # allow sorting
prot.auto_filter = False  # allow filtering
prot.insert_rows = True  # block row insertion
prot.delete_rows = True  # block row deletion
ws.set_protection_detail(prot)

# Workbook-level protection
wb_prot = XlsxWorkbookProtection()
wb_prot.lock_structure = True  # can't add/remove/rename sheets
wb.set_workbook_protection(wb_prot)

path = os.path.join(OUT, "11_protection.xlsx")
wb.save(path)
print(f"   Saved: {path}\n")


# ─────────────────────────────────────────────────────────────────────────────
# 12. IR DERIVE/APPLY — text-based editing workflow
# ─────────────────────────────────────────────────────────────────────────────
# Unique to offidized. Export a workbook to a text IR, edit it as plain text,
# then apply changes back. This is what powers the AI agent workflow.

print("12. IR derive/apply (unique to offidized — powers the AI agent)")

wb = Workbook()
ws = wb.add_sheet("Budget")
ws.set_cell_value("A1", "Item")
ws.set_cell_value("B1", "Amount")
ws.set_cell_value("A2", "Rent")
ws.set_cell_value("B2", 2500)
ws.set_cell_value("A3", "Utilities")
ws.set_cell_value("B3", 200)

path = os.path.join(OUT, "12_ir_source.xlsx")
wb.save(path)

# Derive text IR from the file
ir_text = ir_derive(path)
print("   IR preview (first 300 chars):")
print(f"   {ir_text[:300]}...")
print()

# Modify the IR text and apply back
modified_ir = ir_text.replace("Utilities", "Utilities + Internet")
modified_ir = modified_ir.replace("200", "350")
ir_apply(modified_ir, path)

# Verify
wb2 = Workbook.open(path)
ws2 = wb2.sheet("Budget")
print("   After IR apply:")
print(f"   A3 = {ws2.cell_value('A3')}")
print(f"   B3 = {ws2.cell_value('B3')}")
print()


print("=" * 60)
print(f"All examples saved to: {OUT}")
print("Open them in Excel to verify.")
