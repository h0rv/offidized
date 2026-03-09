# offidized-opc

Open Packaging Convention (OPC) implementation for OOXML files. Handles ZIP I/O, relationships, content types, and roundtrip preservation.

Part of [offidized](../../README.md).

## Usage

```rust
use offidized_opc::Package;

let pkg = Package::open("spreadsheet.xlsx")?;
for part in pkg.parts() {
    println!("{}", part.uri());
}
```

This crate knows nothing about spreadsheets, documents, or presentations. All format-specific logic lives in the format crates (`offidized-xlsx`, `offidized-docx`, `offidized-pptx`).
