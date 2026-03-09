# offidized-cli

CLI tool (`ofx`) for reading, writing, and manipulating OOXML files.

Part of [offidized](../../README.md).

## Install

```bash
cargo install --path crates/offidized-cli
```

## Usage

```bash
ofx info report.xlsx           # Show file metadata
ofx derive report.xlsx         # Derive text IR to stdout
ofx apply report.xlsx.ir -o out.xlsx  # Apply IR edits
ofx nodes report.xlsx          # List unified edit targets
ofx edit report.xlsx --edit 'sheet:Sheet1/cell:A1=Hello' -o out.xlsx
ofx lint report.xlsx           # Run workbook lint checks
```
