# offidized-wasm

WebAssembly bindings for offidized, via `wasm-bindgen`.

Part of [offidized](../../README.md).

## Build

```bash
wasm-pack build crates/offidized-wasm --target web
```

## Overview

Exposes xlsx/docx/pptx read/write APIs and the unified agent API (derive targets, lint edits, preview edits) for use in browser and Node.js environments.
