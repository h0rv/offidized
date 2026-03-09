# Unified Agent API (`offidized-ir`)

## Purpose

One API surface for AI agents to discover editable content and apply edits across `.xlsx`, `.docx`, and `.pptx`.

## Core Types

- `UnifiedDocument`
- `UnifiedNode`
- `UnifiedNodeId`
- `UnifiedEdit`
- `UnifiedEditReport`
- `UnifiedCapabilities`

## ID Grammar

- Spreadsheet cell: `sheet:<sheet>/cell:<A1>`
- Spreadsheet table cell: `sheet:<sheet>/table:<table>/cell:<A1>`
- Spreadsheet chart title: `sheet:<sheet>/chart:<chart>/title`
- Spreadsheet chart series name: `sheet:<sheet>/chart:<chart>/series:<series>/name`
- Spreadsheet style node: `sheet:<sheet>/cell:<A1>/style`
- Doc paragraph: `paragraph:<index>`
- Doc paragraph style: `paragraph:<index>/style`
- Doc table cell: `docx_table:<table>/cell:<row>,<col>`
- Slide title: `slide:<slide>/title`
- Slide subtitle: `slide:<slide>/subtitle`
- Slide notes: `slide:<slide>/notes`
- Slide shape text: `slide:<slide>/shape:<anchor>`
- Slide shape style: `slide:<slide>/shape:<anchor>/style`
- Slide table cell: `slide:<slide>/table:<table>/cell:<row>,<col>`
- Slide chart title: `slide:<slide>/chart:<chart>/title`
- Slide chart series name: `slide:<slide>/chart:<chart>/series:<series>/name`

Indexes are 1-based for table/slide/paragraph identifiers as emitted by derive.

## Usage

```rust
use offidized_ir::{UnifiedDocument, UnifiedDeriveOptions, UnifiedEdit};

let mut doc = UnifiedDocument::derive("deck.pptx".as_ref(), UnifiedDeriveOptions::default())?;
let nodes = doc.nodes();

let report = doc.apply_edits(&[
    UnifiedEdit::new("slide:1/title", "Q2 Earnings"),
    UnifiedEdit::new("slide:1/notes", "Speaker notes updated"),
    UnifiedEdit::new("sheet:Data/cell:A1/style", "")
        .with_payload(UnifiedEditPayload::XlsxCellStyle({
            let mut p = CellStylePatch::new();
            p.set_bold(true);
            p.set_number_format("$#,##0");
            p
        }))
        .with_group("txn-1"),
])?;

assert!(report.applied >= 1);
doc.save_as("deck.updated.pptx".as_ref(), &Default::default())?;
# Ok::<(), offidized_ir::IrError>(())
```

## CLI

- `ofx nodes <file>`
- `ofx capabilities <file>`
- `ofx edit <file> --edit '<id>=<text>' ... --edits-json edits.json --strict -o <out>|-i`
- `ofx edit <file> --edit ... --lint` adds pre-apply diagnostics:
  - `missing_target`
  - `ambiguous_anchor`
  - `invalid_table_coordinates`

`--edits-json` supports transaction groups and typed payloads:

```json
[
  {
    "file": "book.xlsx",
    "id": "sheet:Data/cell:A1/style",
    "group": "txn-1",
    "payload": {
      "kind": "xlsx_cell_style",
      "bold": true,
      "number_format": "$#,##0"
    }
  }
]
```

Cross-file groups are applied atomically when using `--edits-json` with `--in-place`.
`--strict` evaluates apply diagnostics/skips; lint diagnostics are reported separately in `lint_diagnostics`.

## Agent Flows

Agent-oriented integrations should mirror the same unified edit model:

- load the source file into a workspace
- derive canonical IDs/capabilities
- lint proposed edits
- preview changes before write
- apply changes and persist the output artifact
