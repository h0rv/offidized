# offidized

Umbrella crate that re-exports all offidized format crates.

Part of [offidized](../../README.md).

## Usage

```rust
use offidized::xlsx::Workbook;

let mut wb = Workbook::new();
let ws = wb.add_sheet("Sheet1");
ws.cell_mut("A1")?.set_value("Hello");
wb.save("output.xlsx")?;
```

## Features

- `xlsx` (default) — re-exports `offidized-xlsx`
- `docx` (default) — re-exports `offidized-docx`
- `pptx` (default) — re-exports `offidized-pptx`
