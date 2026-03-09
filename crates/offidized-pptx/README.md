# offidized-pptx

High-level PowerPoint API for reading, writing, and manipulating `.pptx` files with full roundtrip fidelity.

Part of [offidized](../../README.md).

## Usage

```rust
use offidized_pptx::Presentation;

let mut pres = Presentation::open("deck.pptx")?;
if let Some(slide) = pres.slide_mut(1) {
    slide.title_mut()?.set_text("Updated Title");
}
pres.save("output.pptx")?;
```

Supports slides, layouts, masters, shapes, text, tables, images, charts, transitions, animations, notes, comments, and shape operations (find/replace, table cell merge).
