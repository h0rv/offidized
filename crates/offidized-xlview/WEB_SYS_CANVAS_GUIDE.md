# Web-sys Canvas 2D API Guide

## Required web-sys Features

```toml
[dependencies.web-sys]
version = "0.3"
features = [
    "Window",
    "Document",
    "Element",
    "HtmlCanvasElement",
    "CanvasRenderingContext2d",
    "TextMetrics",
    "MouseEvent",
    "WheelEvent",
    "KeyboardEvent",
    "Performance",
    "Navigator",
    "console",
]
```

## Getting a 2D Context

```rust
use wasm_bindgen::JsCast;
use web_sys::{CanvasRenderingContext2d, HtmlCanvasElement};

fn get_context(canvas: &HtmlCanvasElement) -> Result<CanvasRenderingContext2d, String> {
    canvas
        .get_context("2d")
        .map_err(|_| "Failed to get context")?
        .ok_or("No 2d context")?
        .dyn_into::<CanvasRenderingContext2d>()
        .map_err(|_| "Failed to cast to CanvasRenderingContext2d")
}
```

## Core Drawing Methods

### Rectangles

```rust
// Clear area
ctx.clear_rect(x, y, width, height);

// Fill rectangle
ctx.set_fill_style_str("#ff0000");
ctx.fill_rect(x, y, width, height);

// Stroke rectangle
ctx.set_stroke_style_str("#000000");
ctx.set_line_width(1.0);
ctx.stroke_rect(x, y, width, height);
```

### Lines (Grid Lines)

```rust
// Draw multiple lines efficiently with a single path
ctx.begin_path();
for x in column_positions {
    ctx.move_to(x, 0.0);
    ctx.line_to(x, height);
}
for y in row_positions {
    ctx.move_to(0.0, y);
    ctx.line_to(width, y);
}
ctx.set_stroke_style_str("#e0e0e0");
ctx.set_line_width(1.0);
ctx.stroke();
```

### Text

```rust
// Set font (CSS font string format)
ctx.set_font("11px Roboto, sans-serif");

// Measure text
let metrics = ctx.measure_text("Hello").unwrap();
let text_width = metrics.width();

// Draw text
ctx.set_fill_style_str("#000000");
ctx.fill_text("Hello", x, y).unwrap();

// Text alignment
ctx.set_text_align("left");   // left, right, center, start, end
ctx.set_text_baseline("top"); // top, hanging, middle, alphabetic, ideographic, bottom
```

### Clipping (for cell text overflow)

```rust
ctx.save();
ctx.begin_path();
ctx.rect(cell_x, cell_y, cell_width, cell_height);
ctx.clip();

// Draw text (will be clipped to cell bounds)
ctx.fill_text(text, x, y).unwrap();

ctx.restore(); // Removes clip
```

## Crisp Pixel Rendering

### The 0.5 Pixel Offset for 1px Lines

Canvas coordinates are at pixel boundaries. A 1px line centered on a boundary spans 0.5px on each side, causing anti-aliasing blur.

**Solution:** Offset by 0.5 pixels for crisp lines.

```rust
// BLURRY - line at x=10 spans pixels 9.5 to 10.5
ctx.move_to(10.0, 0.0);
ctx.line_to(10.0, height);

// CRISP - line at x=10.5 spans pixels 10 to 11
ctx.move_to(10.5, 0.0);
ctx.line_to(10.5, height);
```

**Helper function:**

```rust
fn crisp(x: f64) -> f64 {
    x.floor() + 0.5
}
```

### Device Pixel Ratio (Retina/HiDPI)

Canvas must be sized in physical pixels for sharp rendering:

```rust
fn setup_canvas_scaling(
    canvas: &HtmlCanvasElement,
    ctx: &CanvasRenderingContext2d,
    logical_width: u32,
    logical_height: u32,
) -> f64 {
    let window = web_sys::window().unwrap();
    let dpr = window.device_pixel_ratio();

    // Set canvas buffer size to physical pixels
    canvas.set_width((logical_width as f64 * dpr) as u32);
    canvas.set_height((logical_height as f64 * dpr) as u32);

    // Set CSS size to logical pixels
    let style = canvas.style();
    style.set_property("width", &format!("{}px", logical_width)).unwrap();
    style.set_property("height", &format!("{}px", logical_height)).unwrap();

    // Scale context so drawing uses logical coordinates
    ctx.scale(dpr, dpr).unwrap();

    dpr
}
```

After scaling, all drawing operations use logical pixels - the browser handles the DPR multiplication.

## Performance Tips

### 1. Batch by Fill Style

State changes are expensive. Group draws by color:

```rust
// SLOW - alternating colors
for cell in cells {
    ctx.set_fill_style_str(&cell.color);
    ctx.fill_rect(cell.x, cell.y, cell.w, cell.h);
}

// FAST - batch by color
let mut by_color: HashMap<String, Vec<&Cell>> = HashMap::new();
for cell in cells {
    by_color.entry(cell.color.clone()).or_default().push(cell);
}
for (color, cells) in by_color {
    ctx.set_fill_style_str(&color);
    for cell in cells {
        ctx.fill_rect(cell.x, cell.y, cell.w, cell.h);
    }
}
```

### 2. Single Path for Grid Lines

Don't stroke individual lines:

```rust
// SLOW - individual strokes
for x in cols {
    ctx.begin_path();
    ctx.move_to(x, 0.0);
    ctx.line_to(x, h);
    ctx.stroke();
}

// FAST - single path
ctx.begin_path();
for x in cols {
    ctx.move_to(x, 0.0);
    ctx.line_to(x, h);
}
ctx.stroke();
```

### 3. Cache Font Strings

Don't rebuild font strings every cell:

```rust
// SLOW
for cell in cells {
    ctx.set_font(&format!("{}px {}", cell.font_size, cell.font_family));
}

// FAST - cache the string
let font_string = format!("{}px {}", font_size, font_family);
ctx.set_font(&font_string);
```

### 4. Text Measurement Caching

`measure_text()` is expensive. Cache results:

```rust
struct TextCache {
    cache: HashMap<(String, String), f64>, // (text, font) -> width
}

impl TextCache {
    fn measure(&mut self, ctx: &CanvasRenderingContext2d, text: &str, font: &str) -> f64 {
        let key = (text.to_string(), font.to_string());
        if let Some(&width) = self.cache.get(&key) {
            return width;
        }
        let width = ctx.measure_text(text).unwrap().width();
        self.cache.insert(key, width);
        width
    }
}
```

## Common Gotchas

### 1. Context Reset on Resize

Canvas resize clears all context state (fill style, font, transforms, etc.). Always re-apply after resize:

```rust
fn resize(&mut self, width: u32, height: u32) {
    self.canvas.set_width(width);
    self.canvas.set_height(height);

    // Context state is now reset!
    self.apply_default_styles();
}
```

### 2. JsCast for DOM Types

web-sys requires explicit casting:

```rust
use wasm_bindgen::JsCast;

let element = document.get_element_by_id("canvas").unwrap();
let canvas: HtmlCanvasElement = element.dyn_into().unwrap();
```

### 3. Feature Flags

Every web-sys type needs its Cargo feature enabled. Missing features cause compile errors.

### 4. Coordinate System

Canvas origin (0,0) is top-left. Y increases downward.

### 5. fill_text Baseline

Default baseline is "alphabetic" (text hangs below y). For top-aligned text:

```rust
ctx.set_text_baseline("top");
ctx.fill_text("Hello", x, y).unwrap(); // Top of text at y
```

## Text Truncation with Ellipsis

```rust
fn truncate_text(
    ctx: &CanvasRenderingContext2d,
    text: &str,
    max_width: f64,
) -> String {
    let full_width = ctx.measure_text(text).unwrap().width();
    if full_width <= max_width {
        return text.to_string();
    }

    let ellipsis = "...";
    let ellipsis_width = ctx.measure_text(ellipsis).unwrap().width();
    let available = max_width - ellipsis_width;

    if available <= 0.0 {
        return ellipsis.to_string();
    }

    // Binary search for fit
    let chars: Vec<char> = text.chars().collect();
    let mut low = 0;
    let mut high = chars.len();

    while low < high {
        let mid = (low + high + 1) / 2;
        let truncated: String = chars[..mid].iter().collect();
        let width = ctx.measure_text(&truncated).unwrap().width();
        if width <= available {
            low = mid;
        } else {
            high = mid - 1;
        }
    }

    let truncated: String = chars[..low].iter().collect();
    format!("{}{}", truncated, ellipsis)
}
```

## Color Format

Canvas accepts CSS color strings:

- Hex: `"#ff0000"`, `"#f00"`
- RGB: `"rgb(255, 0, 0)"`
- RGBA: `"rgba(255, 0, 0, 0.5)"`
- Named: `"red"`, `"transparent"`

For Excel colors (often ARGB), convert:

```rust
fn argb_to_css(argb: &str) -> String {
    // Excel: "FFFF0000" (AARRGGBB)
    // CSS: "#FF0000" or "rgba(255, 0, 0, 1.0)"
    if argb.len() == 8 {
        let a = u8::from_str_radix(&argb[0..2], 16).unwrap_or(255);
        let r = u8::from_str_radix(&argb[2..4], 16).unwrap_or(0);
        let g = u8::from_str_radix(&argb[4..6], 16).unwrap_or(0);
        let b = u8::from_str_radix(&argb[6..8], 16).unwrap_or(0);
        format!("rgba({}, {}, {}, {:.2})", r, g, b, a as f64 / 255.0)
    } else {
        format!("#{}", argb)
    }
}
```

## References

- [web-sys on crates.io](https://crates.io/crates/web-sys)
- [web-sys docs](https://docs.rs/web-sys/latest/web_sys/)
- [MDN Canvas Tutorial](https://developer.mozilla.org/en-US/docs/Web/API/Canvas_API/Tutorial)
- [MDN devicePixelRatio](https://developer.mozilla.org/en-US/docs/Web/API/Window/devicePixelRatio)
- [wasm-bindgen Guide](https://rustwasm.github.io/wasm-bindgen/)
