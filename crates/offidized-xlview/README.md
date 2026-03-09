# offidized-xlview

XLSX viewer for the web. Renders Excel files in the browser using WebAssembly and Canvas 2D. View-only. Part of the [offidized](../../README.md) monorepo.

Parsing is backed by `offidized-xlsx` -- the same roundtrip-fidelity OOXML engine used across the offidized ecosystem. The rendering pipeline converts parsed workbook data into Canvas 2D draw calls via WASM.

## Features

- **Styling** -- fonts, colors, borders, fills, gradients, conditional formatting
- **Charts** -- bar, line, pie, scatter, area, radar, and more
- **Rich content** -- images, shapes, comments, hyperlinks, sparklines
- **Layout** -- frozen panes, merged cells, hidden rows/columns, grouping
- **Performance** -- 100k+ cells at 120fps, tile-cached Canvas 2D rendering
- **Interaction** -- native scroll, cell selection, keyboard navigation, copy-to-clipboard

## Install

```bash
npm install offidized-xlview
```

Or via CDN:

```html
<script type="module" src="https://unpkg.com/offidized-xlview"></script>
```

## Quick Start

### Drop-in (1 line)

```html
<script type="module" src="https://unpkg.com/offidized-xlview"></script>
<xl-view src="spreadsheet.xlsx" style="width:100%;height:600px"></xl-view>
```

The `<xl-view>` custom element handles canvas setup, resize, DPR, and rendering automatically. Scroll, click, keyboard, sheet tabs, and selection all work out of the box.

### Programmatic

```js
import { mount } from "offidized-xlview";

const viewer = await mount(document.getElementById("container"));
const res = await fetch("spreadsheet.xlsx");
viewer.load(new Uint8Array(await res.arrayBuffer()));
```

`mount()` creates canvases, wires up resize handling, and returns a controller with `load()`, `destroy()`, and the underlying `viewer` instance.

### React

```tsx
import { useEffect, useRef } from "react";
import { mount, type MountedViewer } from "offidized-xlview";

export function ExcelViewer({ url }: { url: string }) {
  const containerRef = useRef<HTMLDivElement>(null);

  useEffect(() => {
    let mounted: MountedViewer | null = null;
    (async () => {
      mounted = await mount(containerRef.current!);
      const res = await fetch(url);
      mounted.load(new Uint8Array(await res.arrayBuffer()));
    })();
    return () => {
      mounted?.destroy();
    };
  }, [url]);

  return <div ref={containerRef} style={{ width: "100%", height: 600 }} />;
}
```

### Full Control

For custom canvas pipelines, use the WASM API directly:

```js
import init, { XlView } from "offidized-xlview/core";
await init();

const viewer = XlView.newWithOverlay(
  baseCanvas,
  overlayCanvas,
  devicePixelRatio,
);
viewer.load(data);
viewer.render();
```

## Parse Only (No Rendering)

For parse-only access without rendering, use `offidized-xlsx` directly:

```rust
use offidized_xlsx::Workbook;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let data = std::fs::read("spreadsheet.xlsx")?;
    let workbook = Workbook::open_from_bytes(&data)?;
    // ... access cells, styles, etc.
    Ok(())
}
```

## API

### XlView (WASM)

| Method                                 | Description                                        |
| -------------------------------------- | -------------------------------------------------- |
| `new(canvas, dpr)`                     | Create viewer (single canvas, no overlay)          |
| `new_with_overlay(base, overlay, dpr)` | Create viewer with selection overlay (recommended) |
| `load(data)`                           | Load XLSX from `Uint8Array`                        |
| `render()`                             | Render current view                                |
| `resize(w, h, dpr)`                    | Handle container resize                            |
| `set_active_sheet(index)`              | Switch to sheet by index                           |
| `sheet_count()`                        | Number of sheets                                   |
| `sheet_name(index)`                    | Get sheet name                                     |
| `active_sheet()`                       | Current sheet index                                |
| `get_selection()`                      | Get selected cell range `[r1, c1, r2, c2]`         |
| `set_headers_visible(bool)`            | Toggle row/column headers                          |
| `free()`                               | Release WASM memory                                |

### Standalone Functions

| Function                 | Description                    |
| ------------------------ | ------------------------------ |
| `parse_xlsx(data)`       | Parse XLSX, return JSON string |
| `parse_xlsx_to_js(data)` | Parse XLSX, return JS object   |
| `version()`              | Library version                |

## Browser Support

Chrome 57+, Firefox 52+, Safari 11+, Edge 79+

## Relationship to offidized

offidized-xlview is the viewer component of the offidized OOXML toolkit. It depends on:

- **offidized-xlsx** -- XLSX parsing, cell data, styles, themes, relationships
- **offidized-opc** -- ZIP/OPC package handling (transitive via offidized-xlsx)
- **offidized-wasm** -- Shared WASM binding utilities

The parsing layer is shared with the rest of offidized, so files opened in the viewer use the same roundtrip-fidelity parser that powers the read/write API.

## Build from Source

```bash
# From the monorepo root
cargo build -p offidized-xlview

# WASM build
cargo install wasm-pack
wasm-pack build crates/offidized-xlview --target web --release
```

## License

[MIT](../../LICENSE)
