# offidized

**Office, oxidized.** A Rust-native OOXML library for reading, understanding, and editing Excel (.xlsx), Word (.docx), and PowerPoint (.pptx) files with full roundtrip fidelity.

offidized is designed for real-world files, not just clean-room generation. Open an existing workbook, document, or deck, change the parts you care about, and save it without requiring Microsoft Office, COM automation, or a heavyweight runtime. That makes it useful for backends, CLIs, browser/WASM tools, and AI agents that need to inspect and update documents safely.

## Goals

- **100% roundtrip fidelity** — open a file, modify what you need, save it. Everything you didn't touch stays byte-for-byte identical.
- **High-level API** — `workbook.sheet("Sales").cell("A1").set_value(42)`, not XML tree surgery.
- **Full OOXML spec coverage** — generated types from XSD schemas cover every element. High-level API for common operations; drop to typed XML for everything else.
- **Multi-target** — Rust core with PyO3 (Python), wasm-bindgen (browser/Node), and C FFI bindings.
- **No runtime dependencies** — no .NET, no Java, no COM. Just a Rust binary.

## Architecture

```
User code
    ↓
High-level API        (offidized-xlsx, offidized-docx, offidized-pptx)
    ↓                           ↓
    ↓                    Formula Engine (offidized-formula)
    ↓
Generated types       (offidized-schema — every OOXML element as a Rust struct)
    ↓
OPC layer             (offidized-opc — ZIP, relationships, content types)
    ↓
.xlsx / .docx / .pptx
```

## Crates

| Crate               | Purpose                                                           |
| ------------------- | ----------------------------------------------------------------- |
| `offidized`         | Umbrella — re-exports everything                                  |
| `offidized-opc`     | Open Packaging Convention — ZIP I/O, relationships, content types |
| `offidized-schema`  | Generated Rust types from OOXML XSD schemas                       |
| `offidized-codegen` | Build-time code generator (XSD → Rust structs)                    |
| `offidized-xlsx`    | High-level Excel API                                              |
| `offidized-docx`    | High-level Word API                                               |
| `offidized-pptx`    | High-level PowerPoint API                                         |
| `offidized-formula` | Excel formula parser & evaluator                                  |
| `offidized-ir`      | Bidirectional text IR for agent workflows (derive/apply)          |
| `offidized-cli`     | CLI tool (`ofx`) for file operations                              |
| `offidized-py`      | PyO3 Python bindings                                              |
| `offidized-wasm`    | wasm-bindgen bindings for browser/Node                            |
| `offidized-ffi`     | C ABI bindings                                                    |
| `offidized-docview` | Document viewer component (Rust + TypeScript)                     |
| `offidized-pptview` | Presentation viewer component (Rust + TypeScript)                 |

**Non-Rust projects** (in `crates/` but not part of the Cargo workspace):

| Project           | Purpose                              |
| ----------------- | ------------------------------------ |
| `offidized-agent` | Experimental Bun/TypeScript demo app |

## Roundtrip Strategy

The key insight from Microsoft's Open XML SDK: any XML element your code doesn't explicitly model gets preserved as a raw XML fragment on read and written back verbatim on save. This means:

1. You can implement 20% of the spec and still have 100% roundtrip fidelity
2. Users can always drop to the `offidized-schema` typed XML layer for niche features
3. New features are additive — they never break existing roundtrip behavior

## Status

All three format crates (xlsx, docx, pptx) are at 90%+ parity with their C# reference implementations. The OPC foundation layer is effectively complete. Run `cargo test --workspace` for current test counts.

See `docs/` for active roadmaps and design docs.

## Development Phases

### Phase 1: Foundation (OPC + Schema Codegen) ✓

- [x] OPC: ZIP read/write, relationships, content types, roundtrip preservation
- [x] Codegen: XSD → Rust typed API generator
- [x] Schema: Generated SpreadsheetML/WordprocessingML/PresentationML types
- [x] Raw XML preservation for unknown elements

### Phase 2: Excel (offidized-xlsx) ✓

- [x] Workbook, Worksheet, Cell, Row, Column, Range
- [x] Cell values, formulas, shared strings, styles
- [x] Formula parsing and evaluation engine
- [x] Auto-filters, tables, conditional formatting, data validation
- [x] Merge cells, freeze panes, images, charts, sparklines, pivot tables
- [x] Range copy/move, insert/delete rows/columns with ancillary data shifting
- [x] Finance model API, chart templates, workbook lint framework

### Phase 3: Word (offidized-docx) ✓

- [x] Document, Paragraph, Run, Text, character/paragraph formatting
- [x] Styles, tables, sections, headers/footers, images, hyperlinks, lists
- [x] Bullet/numbering API, document protection, content controls, theme colors
- [ ] Tracked changes (accept/reject)

### Phase 4: PowerPoint (offidized-pptx)

- [x] Presentation, Slide, SlideLayout, SlideMaster, shapes, text, tables, images
- [x] Shape ops, find/replace, table cell merge, charts, transitions, animations
- [x] Notes, comments, image manipulation
- [ ] Layout/master write API

### Phase 5: Bindings ✓

- [x] PyO3, wasm-bindgen, C FFI

## Build & Test

```bash
cargo build                    # Build all crates
cargo test --workspace         # Run all tests
cargo test -p offidized-xlsx   # Test specific crate
cargo clippy --workspace -- -D warnings  # Lint
```

Quality gates:

```bash
./scripts/quality.sh           # Full pre-push suite
cargo qfmt                     # Format check
cargo qclippy                  # Clippy with deny warnings
cargo qcheck                   # Fast compile check
just all                       # fmt + clippy + check
```

Reference regression:

```bash
./scripts/bootstrap_references.sh              # Clone reference repos
./scripts/reference_regression.sh              # Curated deterministic tests
./scripts/differential_corpus.sh               # Rust vs C# roundtrip comparison
./scripts/differential_corpus.sh --rust-only   # Rust-only roundtrip
```

## Acknowledgments

offidized's API design and correctness draws heavily from studying the C# OOXML ecosystem:

- [**ClosedXML**](https://github.com/ClosedXML/ClosedXML) — The gold standard for high-level Excel APIs. Primary reference for xlsx API design, style system, and range handling.
- [**ShapeCrawler**](https://github.com/ShapeCrawler/ShapeCrawler) — Modern PowerPoint library. Reference for pptx shape model, text frames, and chart support.
- [**OfficeIMO**](https://github.com/EvotecIT/OfficeIMO) — Clean Word library. Reference for docx paragraph model, sections, and headers/footers.
- [**Open XML SDK**](https://github.com/dotnet/Open-XML-SDK) — Microsoft's foundation layer. We use its JSON schema data as input to our codegen, and its OPC packaging layer as the reference for `offidized-opc`.
- [**NPOI**](https://github.com/nissl-lab/npoi) — .NET port of Apache POI. Fallback reference for edge cases across all formats.

See [`docs/references.md`](docs/references.md) for a detailed study guide.

## License

MIT

Schema data derived from .NET Open XML SDK is licensed under the MIT license.
