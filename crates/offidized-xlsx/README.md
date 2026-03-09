# offidized-xlsx

High-level Excel API for reading, writing, and manipulating `.xlsx` files with full roundtrip fidelity.

Part of [offidized](../../README.md).

## Usage

```rust
use offidized_xlsx::Workbook;

let mut wb = Workbook::new();
let ws = wb.add_sheet("Sales");
ws.cell_mut("A1")?.set_value("Product");
ws.cell_mut("B1")?.set_value("Revenue");
ws.cell_mut("A2")?.set_value("Widget");
ws.cell_mut("B2")?.set_value(42_000);
wb.save("output.xlsx")?;
```

Supports cells, formulas, styles, shared strings, auto-filters, tables, conditional formatting, data validation, merge cells, freeze panes, images, charts, sparklines, pivot tables, and finance-specific features (models, chart templates, lint framework).
