# offidized-ir: Bidirectional Lossless Text IR for Office Files

## Context

AI agents cannot reliably read or edit Office files (xlsx/docx/pptx). Current approaches — OCR, PDF conversion,
lossy Python libraries — are expensive, one-directional, and destroy formatting. offidized already has
roundtrip-fidelity parsing. This adds a text IR that agents read and edit as source code, with
lossless bidirectional sync to the binary file.

Core loop: derive (binary → text) → agent edits text → apply (text → binary). Everything untouched survives via
offidized's roundtrip layer.

## Design Principles

1.  **Stdout-first** — `ofx derive` writes to stdout by default (pipe-friendly for agents). `-o` flag for file output.
2.  **Modal** — Three modes (content, style, full) so the agent only sees what it needs. Content mode is the 80% case.
3.  **Cell-per-line** — Each editable unit is a unique, addressable line. Works perfectly with the Edit tool's string replacement.
4.  **Additive apply** — Cells/paragraphs in the IR are updated. Cells NOT in the IR are left unchanged. Explicit `<empty>` to delete. Safe by default.
5.  **Format-native** — xlsx uses spreadsheet conventions, docx uses markdown, pptx uses slide/shape structure. Each format gets the most natural representation.

## Architecture

```
offidized-ir
├── src/
│   ├── lib.rs          — Public API: derive(), apply(), Mode enum
│   ├── header.rs       — IR header parsing/writing (TOML front matter)
│   ├── xlsx/
│   │   ├── content.rs  — xlsx content mode derive + apply
│   │   ├── style.rs    — xlsx style mode
│   │   └── full.rs     — xlsx full mode
│   ├── docx/
│   │   └── content.rs  — docx content mode derive + apply
│   └── pptx/
│       └── content.rs  — pptx content mode derive + apply
```

offidized-ir depends on offidized-xlsx, offidized-docx, offidized-pptx.
offidized-cli depends on offidized-ir.

## IR Header

Every IR starts with a TOML front matter block (using `+++` delimiters). TOML over YAML to
avoid type coercion footguns (the Norway problem, bare true/false, etc.):

```
+++
source = "report.xlsx"
format = "xlsx"
mode = "content"
version = 1
checksum = "sha256:a1b2c3d4..."
+++
```

- `source` — Path to the original file (for apply to find it)
- `format` — xlsx | docx | pptx
- `mode` — content | style | full
- `version` — IR format version (for forward compat)
- `checksum` — SHA-256 of source file at derive time (staleness detection)

File extensions: `.xlsx.ir`, `.docx.ir`, `.pptx.ir`.

## xlsx Content Mode

### Derive Output

```
+++
source = "quarterly-report.xlsx"
format = "xlsx"
mode = "content"
version = 1
checksum = "sha256:e3b0c44298..."
+++

=== Sheet: Revenue ===
A1: Category
B1: Q3 2025
C1: Q4 2025
A2: Hardware
B2: =SUM(B3:B5)
C2: =SUM(C3:C5)
```

### Value Encoding

| IR Text            | Interpreted As         | CellValue          |
| ------------------ | ---------------------- | ------------------ |
| `42000`            | Bare number            | Number(42000.0)    |
| `3.14`             | Bare float             | Number(3.14)       |
| `=SUM(B3:B5)`      | Starts with `=`        | Formula            |
| `true` / `false`   | Boolean literal        | Bool               |
| `#REF!` / `#NAME?` | Error value            | Error              |
| `"42"`             | Quoted → forced string | String("42")       |
| `Category`         | Everything else        | String("Category") |
| `<empty>`          | Explicit empty         | Clear cell         |
| (line omitted)     | Not in IR              | Left unchanged     |

Derive quoting invariant: The derive→parse roundtrip MUST be lossless. Strings that look like numbers, booleans, formulas, errors, or `<empty>` are always quoted on derive.

Cell ordering: Row-major (A1, B1, C1, A2, B2, C2, ...). Empty cells between populated cells are omitted (sparse representation).

### Key Guarantees

- Only cell values/formulas are touched. Styles are never modified.
- Charts, conditional formatting, tables, images, pivot tables, merged cells: all untouched.
- Cells not in the IR are not modified or deleted.
- New cells (refs that don't exist in source) are created.

## docx Content Mode

Markdown with paragraph anchors `[pN]` for addressability:

```
+++
source = "proposal.docx"
format = "docx"
mode = "content"
version = 1
checksum = "sha256:..."
+++

[p1] # Q4 Financial Summary

[p2] The quarterly results show **strong growth** across all segments.

[p4] - Launched enterprise tier with 12 new contracts
[p5] - Reduced churn to 2.1% (from 3.4%)

[t1]
| Category | Q3 | Q4 |
|----------|---:|---:|
| Hardware | $42K | $51K |
```

Markdown mapping: `#` → Heading 1, `**bold**` → Bold run, `- item` → Bulleted list, `[text](url)` → Hyperlink, `> quote` → Block quote, markdown tables → docx tables.

Structures that can't be represented losslessly (nested lists, content controls, tracked changes) are emitted with `<!-- complex: ... -->` placeholders and left unchanged on apply.

## pptx Content Mode

```
+++
source = "deck.pptx"
format = "pptx"
mode = "content"
version = 1
checksum = "sha256:..."
+++

--- slide 1 [Title Slide] ---

[title] Q4 Business Review
[subtitle] December 2025 | Confidential

--- slide 2 [Title and Content] ---

[title] Revenue Overview
[shape "Key Info"]

- 15% YoY growth
- New enterprise contracts: 12

[notes] Remember to mention the APAC expansion timeline.
```

Mapping: `--- slide N [Layout] ---` = slide delimiter, `[title]`/`[subtitle]` = placeholder shapes, `[shape "Name"]` = shape by name, `[notes]` = slide notes. Non-text content (`(chart)`, `(image)`, `(table)`) is read-only in content mode.

## Style Mode

Separate IR for style-only information (column widths, row formatting, cell overrides, conditional formatting, sheet properties for xlsx; paragraph styles, section layout for docx; slide backgrounds, shape positions, fonts for pptx).

## Full Mode

Combines content and style. Cell-per-line with style annotations in `{...}`:

```
A1: Category {bold, fill=#4472C4, font-color=#FFFFFF}
B2: =SUM(B3:B5) {format="#,##0"}
```

## CLI Interface

```bash
# Derive
ofx derive report.xlsx                      # → stdout
ofx derive report.xlsx -o report.xlsx.ir    # → file
ofx derive report.xlsx --mode style         # style mode
ofx derive report.xlsx --sheet Revenue      # single sheet

# Apply
ofx apply report.xlsx.ir -o updated.xlsx    # output to new file
ofx apply report.xlsx.ir -i                 # in-place modify

# Diff via standard tools
diff <(ofx derive a.xlsx) <(ofx derive b.xlsx)
```

## API

```rust
pub fn derive(path: &Path, options: DeriveOptions) -> Result<String>;
pub fn apply(ir: &str, output: &Path) -> Result<ApplyResult>;
```

For agent integration, the agent calls `derive()` → gets a string → edits it → calls `apply()`. No temp files required.
