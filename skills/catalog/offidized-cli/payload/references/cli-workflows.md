# CLI Workflows

Prefer `cargo run -p offidized-cli -- <subcommand>` while working in this repo. Use installed `ofx` only if the user explicitly wants the globally installed binary.

## Install Or Run

- Repo-local: `cargo run -p offidized-cli -- --help`
- Global install: `cargo install --path crates/offidized-cli`
- Installed binary: `ofx --help`

## Inspect Before Editing

- File metadata: `cargo run -p offidized-cli -- info report.xlsx`
- Read spreadsheet cells: `cargo run -p offidized-cli -- read report.xlsx 'Sales!A1:D10'`
- Read DOCX paragraphs: `cargo run -p offidized-cli -- read contract.docx --paragraphs 0-4`
- List OPC parts: `cargo run -p offidized-cli -- part report.xlsx --list`

Use `part <file> <uri>` to extract a specific XML part when raw package debugging is required.

## Point Edits

- Set one value: `cargo run -p offidized-cli -- set report.xlsx 'Sales!B2' 42000 -o out.xlsx`
- Replace text: `cargo run -p offidized-cli -- replace proposal.docx Draft Final -o proposal-final.docx`
- Apply JSON patch from stdin: `cargo run -p offidized-cli -- patch report.xlsx -o out.xlsx`

Prefer `-o <path>` to preserve the input file unless the user explicitly wants in-place edits.

## Unified Cross-Format Edits

Use this path for content edits across `xlsx`, `docx`, and `pptx`:

1. Discover IDs: `cargo run -p offidized-cli -- nodes file.pptx`
2. Check supported edit surface: `cargo run -p offidized-cli -- capabilities file.pptx`
3. Apply validated edits: `cargo run -p offidized-cli -- edit file.pptx --edit '<id-from-nodes>=New text' --lint --strict -o out.pptx`

Treat IDs returned by `nodes` as canonical. Do not guess IDs.

### JSON Edit Payloads

Use `--edits-json` for multiple edits or typed payloads:

```json
[
  {
    "id": "<id-from-ofx-nodes>",
    "text": "Quarterly Update"
  },
  {
    "id": "<id-from-ofx-nodes>",
    "group": "title-block",
    "payload": {
      "kind": "pptx_text_style",
      "bold": true,
      "font_size": 24,
      "font_color": "FF1F4E79",
      "font_name": "Aptos"
    }
  },
  {
    "id": "<id-from-ofx-nodes>",
    "payload": {
      "kind": "xlsx_cell_style",
      "bold": true,
      "number_format": "$#,##0.00"
    }
  }
]
```

Notes:

- `file` is optional per edit and only matters for cross-file batches.
- Cross-file batches require `--in-place` and cannot be combined with `-o`.
- `--force` skips checksum validation only; it does not disable `--lint` or `--strict`.

## IR Workflow

Use IR when the user wants a reviewable text representation:

```bash
cargo run -p offidized-cli -- derive report.xlsx -o report.ir
$EDITOR report.ir
cargo run -p offidized-cli -- apply report.ir -o report-updated.xlsx
```

Useful options:

- `derive --mode content`
- `derive --mode style`
- `derive --mode full`
- `derive --sheet Sales`
- `apply --source moved-report.xlsx`
- `apply --dry-run`

`apply --dry-run` is only a light sanity check today. It does not emit a full diff.

## Spreadsheet-Specific Commands

- Create workbook: `cargo run -p offidized-cli -- create report.xlsx Dashboard Data`
- Copy range: `cargo run -p offidized-cli -- copy-range report.xlsx 'Sheet1!A1:B10' 'Sheet1!D1' -o out.xlsx`
- Move range: `cargo run -p offidized-cli -- move-range report.xlsx 'Sheet1!A1:B10' 'Sheet1!D1' -o out.xlsx`
- List pivots: `cargo run -p offidized-cli -- pivots report.xlsx`
- List charts: `cargo run -p offidized-cli -- charts report.xlsx`
- Add chart: `cargo run -p offidized-cli -- add-chart report.xlsx --sheet Sales --type bar --series 'Revenue | Sales!A2:A5 | Sales!B2:B5' -o out.xlsx`
- Evaluate formula: `cargo run -p offidized-cli -- eval report.xlsx '=SUM(B2:B10)' --sheet Sales`
- Lint workbook: `cargo run -p offidized-cli -- lint report.xlsx`

## When To Escalate

Switch to `references/python-bindings.md` when the user needs:

- Loops or generated content
- Repeated transformations across many files
- Programmatic document construction
- Richer object APIs than the CLI exposes ergonomically
