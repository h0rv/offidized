# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this crate.

## What This Is

offidized-xlview is a WASM-based Excel (XLSX) viewer with Canvas 2D rendering. It uses offidized-xlsx for parsing and compiles to WebAssembly for browser rendering. View-only (no editing). Part of the offidized monorepo.

## Build & Development Commands

```bash
# Build (within workspace)
cargo build -p offidized-xlview

# Build WASM (dev)
wasm-pack build crates/offidized-xlview --target web --dev

# Build WASM (release)
wasm-pack build crates/offidized-xlview --target web --release

# Run Rust unit tests
cargo test -p offidized-xlview

# Clippy (matches CI)
cargo clippy -p offidized-xlview -- -D warnings

# Format check
cargo fmt -p offidized-xlview -- --check

# Full workspace quality check
./scripts/quality.sh
```

## Strict Lint Rules

Inherits the workspace-wide lint configuration from the root `Cargo.toml`:

- **Forbidden**: `unsafe` code
- **Denied**: `.unwrap()`, `.expect()`, `panic!()`, `todo!()`, `unimplemented!()`, `dbg!()`, lossy casts
- **Required**: Use `.get()` for indexing, `?` for error propagation, proper `Result`/`Option` handling

CI runs with `RUSTFLAGS=-Dwarnings` -- all warnings are errors.

## Architecture

### Parsing Pipeline

XLSX parsing is handled entirely by `offidized-xlsx`. The adapter layer (`src/adapter.rs`) converts offidized-xlsx types into the view model types used by the rendering pipeline.

The parsing flow:

1. `offidized-xlsx` opens the XLSX package via `offidized-opc`
2. Shared strings, themes, styles, worksheets are parsed by offidized-xlsx
3. `adapter.rs` converts parsed data into offidized-xlview's internal types (`src/types/`)

### Rendering Pipeline

- `src/layout/` -- Pre-computes cell positions (`sheet_layout.rs`) and manages viewport/scroll state (`viewport.rs`)
- `src/render/` -- Backend-agnostic rendering trait (`backend.rs`), with Canvas 2D implementation in `render/canvas/`
  - `canvas/renderer.rs` -- Main cell/text/border drawing
  - `canvas/headers.rs` -- Row/column headers
  - `canvas/frozen.rs` -- Frozen pane support
  - `canvas/indicators.rs` -- Comment markers, etc.
- `src/render/blit.rs` -- Tile-based off-screen caching (512px tiles)
- `src/render/selection.rs` -- Selection overlay

### Viewer (`src/viewer/`)

The main WASM-exported viewer struct. Owns the parsed workbook, layout engine, and renderer. Handles scroll, click, keyboard events, sheet switching, and coordinates the full render cycle.

### WASM Exports (`src/lib.rs`)

- `XlView` -- Full interactive viewer (canvas-based)
- `parse_xlsx()` / `parse_xlsx_to_js()` -- Parse-only APIs returning JSON or JsValue
- `version()` -- Library version

## Test Structure

- **Rust unit tests**: Cover parsing adapter, layout, and rendering modules. Run with `cargo test -p offidized-xlview`.
- **Test data**: Shared test fixtures in `test/` directory and workspace-level test data.

## Key Differences from Standalone xlview

- Parsing is delegated to `offidized-xlsx` instead of a built-in parser
- Types come from offidized-xlsx/offidized-schema, adapted via `src/adapter.rs`
- Build uses workspace-level configuration and quality gates
- Shares the OPC layer, styles, and theme handling with the rest of offidized
