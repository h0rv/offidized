# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

**offidized** ("Office, oxidized") is a Rust-native OOXML library for reading, writing,
and manipulating Excel (.xlsx), Word (.docx), and PowerPoint (.pptx) files with full
roundtrip fidelity.

## Architecture

```
offidized (umbrella re-export)
├── offidized-opc         — ZIP/OPC layer (shared by all formats)
├── offidized-schema      — Generated typed structs for all OOXML XML elements
├── offidized-codegen     — Code generator: XSD/JSON → Rust structs
├── offidized-formula     — Excel formula parser and evaluator
├── offidized-xlsx        — High-level Excel API (ClosedXML-inspired)
├── offidized-docx        — High-level Word API
├── offidized-pptx        — High-level PowerPoint API
├── offidized-ir          — Bidirectional text IR for agent workflows (derive/apply)
├── offidized-cli         — CLI tool (ofx) for file operations
├── offidized-docview     — Document viewer component (Rust + TypeScript)
├── offidized-pptview     — Presentation viewer component (Rust + TypeScript)
├── offidized-py          — PyO3 Python bindings
├── offidized-wasm        — WebAssembly bindings
└── offidized-ffi         — C FFI bindings
```

Non-Rust projects (in `crates/` but outside the Cargo workspace):

- `offidized-agent` — Bun/TypeScript AI agent server
- `offidized-mcp` — Cloudflare Worker MCP server

## Key Design Principles

### 1. Roundtrip Fidelity Above All

Every parsed XML element must preserve unknown children and attributes as `RawXmlNode`.
Open a complex Excel file, change one cell, save — everything else must survive unchanged.
Test this with real-world files from Excel, Word, and PowerPoint.

### 2. Layered Access

Users get a high-level API (`cell("A1").set_value(42)`), but can always drop to:

- Generated schema types (typed XML elements)
- Raw `offidized-opc::Package` (direct part access)
  This means features don't need to be in the high-level API to be usable.

### 3. Code Generation for Schema Types

The OOXML spec has hundreds of element types. We generate Rust structs from the same
JSON schema data that .NET Open XML SDK uses (in `references/Open-XML-SDK/data`).
The codegen lives in `offidized-codegen`. Never hand-write schema types — generate them.

### 4. Shared OPC Layer

`offidized-opc` knows nothing about spreadsheets, documents, or presentations. It only
understands ZIP packages, relationships, content types, and part URIs. All format-specific
logic lives in the format crates.

## Reference Implementations

**Primary references (C# only):** Clone these into `references/` for study:

- `references/Open-XML-SDK/` — C# low-level OOXML (~36k hand-written + codegen).
  Study for JSON schema data (`data/` directory), packaging layer, validation. This is
  the authoritative source for schema definitions.
- `references/ClosedXML/` — C# Excel library (~92k lines). Study for xlsx API design,
  style system, range handling, IO serialization patterns.
- `references/ShapeCrawler/` — C# PowerPoint library. Study for pptx API design,
  shape trees, chart subsystem, layout/master handling.
- `references/ooxmlsdk/` — Rust port of Open XML SDK schema layer.
  Study for Rust codegen patterns, how they handle the JSON schema data.

**Note:** Prefer C# reference implementations over Python libraries. The C# ecosystem
has more mature OOXML handling and better roundtrip fidelity patterns.

## Build & Test

### Basic Commands

```bash
cargo build                    # Build all crates (excludes offidized-py which needs maturin)
cargo test --workspace         # Run all tests
cargo test -p offidized-opc    # Test specific crate (opc/xlsx/docx/pptx)
cargo clippy --workspace -- -D warnings  # Lint all code
```

### Quality Gates

Run the full pre-push quality suite:

```bash
./scripts/quality.sh
```

Or use cargo aliases (if configured):

```bash
cargo qfmt      # Format check
cargo qclippy   # Clippy with deny warnings
cargo qcheck    # Fast compile check
```

### Reference Regression Testing

Bootstrap/update reference repositories:

```bash
./scripts/bootstrap_references.sh
```

Run curated deterministic tests against reference implementations:

```bash
./scripts/reference_regression.sh          # Full run
./scripts/reference_regression.sh --skip-full  # Skip final workspace test
./scripts/reference_regression.sh --offline    # Offline mode
```

Logs and reports are written to `artifacts/reference-regression/<timestamp>/`

### Differential Corpus Testing

Run Rust vs C# (Open XML SDK) differential roundtrip on all `*.docx`, `*.xlsx`, and `*.pptx`
files under `references/`:

```bash
./scripts/differential_corpus.sh                           # Full corpus
./scripts/differential_corpus.sh --rust-only               # Skip C# comparison
./scripts/differential_corpus.sh --max-files 10            # Smoke test first 10 files
./scripts/differential_corpus.sh --rust-engine opc --compare-input  # OPC layer only
```

Reports are written to `artifacts/differential-regression/<timestamp>/`

## Current Status

- **OPC layer:** Complete. Handles roundtrip fidelity, relationships, content types.
- **xlsx:** 90%+ parity. See `README.md` for feature checklist.
- **docx:** 90%+ parity. Remaining: tracked changes accept/reject.
- **pptx:** 90%+ parity. Remaining: layout/master write API.
- **Bindings:** All three (py/wasm/ffi) scaffolded and functional.

Run `cargo test --workspace` for current test counts.

## Common Bugs & Patterns

### quick_xml Temporary Lifetime (E0716)

**Problem:** In quick_xml parsers, `e.name().as_ref()` creates a temporary `QName` that gets
dropped before its borrow is used.

```rust
// ❌ WRONG: creates temporary that's dropped too early
let local = local_name(e.name().as_ref());

// ✅ RIGHT: bind the name first
let name_bytes = e.name();
let local = local_name(name_bytes.as_ref());
```

This pattern appears frequently when parsing XML events. Always bind `e.name()` to a variable
before calling `.as_ref()` on it.

## Code Style

### Error Handling

- Use `thiserror` for error types in library crates
- Use `anyhow` only in codegen/tools and CLI applications
- Never use `.unwrap()`, `.expect()`, `panic!()`, or `todo!()` in library code (enforced by clippy)

### API Design

- Prefer `&str` over `String` in public APIs where possible
- All public types, functions, and modules need doc comments
- Return `Result<T, Error>` for fallible operations

### Testing

- Test with real OOXML files from Excel/Word/PowerPoint, not just synthetic test files
- Roundtrip tests are critical: open file, modify minimally, save, verify everything else unchanged
- Use `pretty_assertions` for better test failure output

### Performance

- Keep CI fast: schema generation is a separate build step, not a proc macro
- Use spatial indexing (rstar) for cell/range lookups in xlsx

### Lints

The workspace enforces strict lints (see `Cargo.toml`):

**Rust lints:**

- `unsafe_code = "forbid"` — no unsafe code allowed
- `unused_lifetimes = "deny"`
- `unused_qualifications = "deny"`

**Clippy lints:**

- `unwrap_used = "deny"` — no `.unwrap()` or `.expect()`
- `panic = "deny"` — no `panic!()`, `todo!()`, `unimplemented!()`
- `dbg_macro = "deny"` — no `dbg!()` in committed code
- `clone_on_ref_ptr = "deny"` — avoid unnecessary Arc/Rc clones
- `redundant_clone = "deny"` — catch performance issues

All code must pass `cargo clippy --workspace -- -D warnings` with zero warnings.

## OOXML Resources

- ECMA-376 spec: https://ecma-international.org/publications-and-standards/standards/ecma-376/
- Open XML explained: http://officeopenxml.com/
- SpreadsheetML reference: http://officeopenxml.com/SScontentOverview.php
- WordprocessingML reference: http://officeopenxml.com/WPcontentOverview.php
- PresentationML reference: http://officeopenxml.com/PRcontentOverview.php
