# offidized-docx

High-level Word API for reading, writing, and manipulating `.docx` files with full roundtrip fidelity.

Part of [offidized](../../README.md).

## Usage

```rust
use offidized_docx::Document;

let mut doc = Document::new();
doc.add_paragraph("Hello, World!")
    .style("Heading1");
doc.add_paragraph("This is a paragraph.");
doc.save("output.docx")?;
```

Supports paragraphs, runs, character/paragraph formatting, styles, tables, sections, headers/footers, images, hyperlinks, lists, numbering, document protection, content controls, and theme colors.
