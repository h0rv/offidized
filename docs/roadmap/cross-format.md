# Cross-Format Consistency Roadmap

Status: `active`

## Goal

Provide one agent-facing API across `xlsx`, `docx`, and `pptx` where concepts overlap, while preserving format-specific power behind explicit opt-in APIs.

## P0 (Foundation)

### P0.1 Unified Content Node Surface

Status: `done`

Implemented in `offidized-ir`:

- `list_nodes_from_ir(ir)`
- `derive_content_nodes(path)`
- `UnifiedNode { id: UnifiedNodeId, kind, text }`
- `UnifiedNodeId` parser/formatter contract
- `UnifiedNodeKind` variants for shared editable content concepts:
  - spreadsheet cell
  - paragraph
  - docx table cells
  - slide title/subtitle/notes/shape text
  - pptx table cells
  - xlsx table cells
  - xlsx chart titles/series names
  - xlsx style nodes
  - docx paragraph style nodes
  - pptx shape text style nodes
  - pptx chart titles/series names

Node id contract:

- xlsx: `sheet:<name>/cell:<A1>`
- xlsx table: `sheet:<name>/table:<table>/cell:<A1>`
- xlsx chart title: `sheet:<name>/chart:<index>/title`
- xlsx chart series name: `sheet:<name>/chart:<index>/series:<index>/name`
- xlsx style: `sheet:<name>/cell:<A1>/style`
- docx: `paragraph:<index>`
- docx paragraph style: `paragraph:<index>/style`
- docx table: `docx_table:<table>/cell:<row>,<col>`
- pptx: `slide:<index>/title|subtitle|notes|shape:<anchor>`
- pptx shape style: `slide:<index>/shape:<anchor>/style`
- pptx table: `slide:<index>/table:<table>/cell:<row>,<col>`
- pptx chart title: `slide:<index>/chart:<index>/title`
- pptx chart series name: `slide:<index>/chart:<index>/series:<index>/name`

### P0.2 Unified Edit Surface

Status: `done`

Implemented in `offidized-ir`:

- `UnifiedEdit::new(id, text)`
- `apply_edits_to_ir(ir, edits)`
- `edit_file_content(source, output, edits, apply_options)`
- `UnifiedDocument::{derive, from_ir, apply_edits, save_as}`
- `UnifiedEditReport` + structured diagnostics
- `UnifiedCapabilities`

Behavior:

- Uses existing derive/apply pipeline; no adapter bypass.
- Supports content-mode edits across all 3 formats.
- Preserves style/full-mode style section when editing full-mode IR content.
- Supports direct mutation queue for xlsx/pptx chart/style metadata updates while still applying through the canonical writer pipeline.

### P0.3 Test Coverage for Node/Edit Contract

Status: `done`

Added unit tests in `crates/offidized-ir/src/unified_api.rs` for:

- xlsx node listing + cell edit
- docx node listing + paragraph/table-cell edits
- pptx node listing + title/shape/notes/table-cell edits
- `UnifiedNodeId` parse/format roundtrip
- missing-target diagnostics in `UnifiedEditReport`

### P0.4 CLI Facade

Status: `done`

Added CLI commands:

- `ofx nodes <file>`
- `ofx capabilities <file>`
- `ofx edit <file> --edit '<id>=<text>' ... --edits-json edits.json --strict -o <out>|-i`

## P1 (Next)

- Add optional validation/lint profile for unified ids (`missing-target`, `ambiguous-anchor`, `invalid-table-coordinates`). `done`
- Add edit transaction groups (all-or-nothing for grouped edits across files). `done`
- Add structured style edit grammar (JSON schema) beyond `key=value;...`. `done`
- Expand style parity beyond current docx paragraph + pptx shape run level (e.g., table style ids, multi-run selection).
