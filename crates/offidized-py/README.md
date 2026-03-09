# offidized

**Office, oxidized.** Rust-native OOXML library for reading, writing, and manipulating Excel (.xlsx), Word (.docx), and PowerPoint (.pptx) files with full roundtrip fidelity.

## Install

```bash
pip install offidized
```

Supports Python 3.9+ on Linux, macOS, and Windows.

## Python Binding Coverage

The Rust crates have 90%+ feature parity (see root README for details). Below tracks what's exposed to Python.

### Excel (.xlsx)

- [x] Workbook open/save/bytes roundtrip
- [x] Sheet CRUD (add, remove, list)
- [x] Cell read/write (string, number, bool, date)
- [x] Cell formulas
- [x] Cell styles (font, fill, border, alignment, number format)
- [x] Merged ranges
- [x] Freeze panes
- [x] Auto-filter
- [x] Hyperlinks
- [x] Defined names
- [x] Find cells
- [x] Sheet protection (basic + detailed)
- [x] Workbook protection
- [x] Tab color
- [x] Row height / column width
- [x] Images
- [x] Charts
- [x] Conditional formatting
- [x] Data validation
- [x] Cell comments
- [x] Page setup / print settings
- [x] Sheet view options (gridlines, zoom)
- [x] Sparklines
- [x] Pivot tables
- [x] Named/structured tables (full)
- [x] Workbook lint
- [x] Sheet visibility (hidden/very hidden)
- [x] Rich text cells

### Word (.docx)

- [x] Document open/save/bytes roundtrip
- [x] Paragraphs CRUD
- [x] Headings
- [x] Bulleted / numbered paragraphs
- [x] Paragraph styling (alignment, spacing, indentation)
- [x] Runs with formatting (bold, italic, underline, strikethrough, color, font)
- [x] Hyperlinks on runs
- [x] Tables CRUD + cell text read/write
- [x] Inline images (add, count)
- [x] Sections (page size, orientation, margins)
- [x] Document properties
- [x] Comments (add, list, count)
- [x] Footnotes (add, list, count)
- [x] Bookmarks (add, list, count)
- [x] Content controls (count)
- [x] Document protection (basic)
- [x] Body items iteration
- [x] Table structural mutation (add/remove rows/columns)
- [x] Table cell merge
- [x] Table borders / widths / layout
- [x] Section headers and footers
- [x] Endnotes
- [x] Style registry (inspect/create named styles)
- [x] Paragraph borders / shading
- [x] Tab stops
- [ ] Floating images
- [ ] Field codes (TOC, date, etc.)

### PowerPoint (.pptx)

- [x] Presentation open/save/from_bytes
- [x] Slide CRUD (add, remove, clone, move)
- [x] Slide dimensions
- [x] Find/replace text (presentation-wide)
- [x] Presentation properties
- [x] Shapes (add, list, solid fill, alt text, preset geometry, word wrap)
- [x] Shape position / size (`set_geometry(x, y, width, height)` in EMUs)
- [x] Shape rotation (`set_rotation(degrees)`)
- [x] Shape outline (`set_outline(color, width_pt, dash_style)`)
- [x] Shape gradient fill (`set_gradient_fill(stops, angle)`)
- [x] Shape paragraphs (add, list, alignment, spacing, indent, bullets)
- [x] TextRun formatting (bold, italic, underline, strikethrough, font, color, size, hyperlinks, spacing)
- [x] Tables (cell text, column width, row height, cell fill/bold/italic)
- [x] Table cell font size (`set_cell_font_size`) and color (`set_cell_font_color`)
- [x] Charts (title, legend, data points)
- [x] Images (name, content type)
- [x] Slide notes
- [x] Slide transitions
- [x] Slide show settings / custom shows
- [x] Presentation to_bytes
- [x] Shape flip (horizontal/vertical)
- [x] Shape pattern fill / picture fill
- [x] Shape effects (shadow, glow, reflection)
- [x] Slide background
- [x] Slide masters / layouts
- [x] Slide placeholders
- [x] Table cell merge / borders / alignment
- [x] Table add/remove rows/columns
- [x] Chart series / axes / type switching
- [x] Image position / size / replace / export
- [ ] Animation / timing
- [x] Theme colors / fonts

### IR (Intermediate Representation)

- [x] `ir_derive` / `ir_derive_from_bytes`
- [x] `ir_apply` / `ir_apply_to_bytes`
- [x] `UnifiedDocument` (derive, from_ir, to_ir, nodes, capabilities, apply_edits, lint_edits, save_as)

## About

Python binding for [offidized](https://github.com/h0rv/offidized), built with [PyO3](https://pyo3.rs) and [maturin](https://www.maturin.rs). The heavy lifting happens in Rust — Python gets a thin, ergonomic wrapper.

## License

MIT OR Apache-2.0
