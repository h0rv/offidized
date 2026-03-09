//! Canvas 2D rendering backend.
//!
//! Implements the RenderBackend trait using HTML Canvas 2D API via web-sys.

use std::borrow::Cow;
use std::collections::{HashMap, VecDeque};
use std::fmt::Write as _;
use std::rc::Rc;
use wasm_bindgen::JsCast;
use web_sys::{
    CanvasPattern, CanvasRenderingContext2d, Document, HtmlCanvasElement, HtmlImageElement,
};

use crate::error::Result;
use crate::types::drawing::{Drawing, EmbeddedImage};

use crate::layout::{SheetLayout, Viewport};
use crate::render::backend::{
    BorderStyleData, CellRenderData, CellStyleData, RenderBackend, RenderParams,
};
use crate::render::blit::scrollable_region;
use crate::render::colors::{palette, parse_color, Rgb};
use crate::render::selection::selection_rects;

use super::frozen::render_frozen_dividers;
use super::headers::{render_column_headers, render_header_corner, render_row_headers};
use super::indicators::{
    render_comment_indicators, render_filter_buttons, render_validation_indicators,
};

/// UI Constants
const TAB_BAR_HEIGHT: f64 = 28.0;
// Native scrollbars are provided by the browser scroll container,
// so the canvas should not reserve space for them.
pub(super) const SCROLLBAR_SIZE: f64 = 0.0;
const CELL_PADDING: f64 = 4.0;
const TILE_SIZE: f64 = 512.0;
const TILE_CACHE_CAP: usize = 256;

/// EMU conversion constants
/// 1 inch = 914400 EMUs, 96 pixels = 1 inch, so 1 pixel = 9525 EMUs
const EMU_PER_PIXEL: f64 = 9525.0;
const TEXT_MEASURE_CACHE_CAP: usize = 4096;
const TEXT_WRAP_CACHE_CAP: usize = 1024;

struct TextMeasureCache {
    entries: HashMap<String, f64>,
    order: VecDeque<String>,
    max_entries: usize,
    scratch: String,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash)]
struct TileKey {
    x: i32,
    y: i32,
}

struct TileCacheEntry {
    canvas: CacheCanvas,
    last_used: u64,
}

struct TileCache {
    tiles: HashMap<TileKey, TileCacheEntry>,
    scale: f32,
    dpr: f32,
}

#[derive(Clone, Copy, Debug)]
struct TileRenderResult {
    rendered_visible: bool,
    deferred_prefetch: bool,
}

impl TileCache {
    fn new() -> Self {
        Self {
            tiles: HashMap::new(),
            scale: 1.0,
            dpr: 1.0,
        }
    }

    fn clear(&mut self) {
        self.tiles.clear();
    }
}

impl TextMeasureCache {
    fn new(max_entries: usize) -> Self {
        Self {
            entries: HashMap::new(),
            order: VecDeque::new(),
            max_entries,
            scratch: String::new(),
        }
    }

    fn clear(&mut self) {
        self.entries.clear();
        self.order.clear();
    }

    fn get(&mut self, font: &str, text: &str) -> Option<f64> {
        if self.max_entries == 0 {
            return None;
        }
        let key = Self::build_key(&mut self.scratch, font, text);
        self.entries.get(key).copied()
    }

    fn insert(&mut self, font: &str, text: &str, width: f64) {
        if self.max_entries == 0 {
            return;
        }
        let key = Self::build_key(&mut self.scratch, font, text);
        if self.entries.contains_key(key) {
            return;
        }
        let key_owned = key.to_string();
        self.entries.insert(key_owned.clone(), width);
        self.order.push_back(key_owned);
        self.enforce_cap();
    }

    #[cfg(test)]
    fn len(&self) -> usize {
        self.entries.len()
    }

    fn build_key<'a>(scratch: &'a mut String, font: &str, text: &str) -> &'a str {
        scratch.clear();
        scratch.reserve(font.len() + 1 + text.len());
        scratch.push_str(font);
        scratch.push('\n');
        scratch.push_str(text);
        scratch.as_str()
    }

    fn enforce_cap(&mut self) {
        while self.entries.len() > self.max_entries {
            if let Some(oldest) = self.order.pop_front() {
                self.entries.remove(&oldest);
            } else {
                break;
            }
        }
    }
}

struct TextWrapCache {
    entries: HashMap<String, Rc<Vec<String>>>,
    order: VecDeque<String>,
    max_entries: usize,
    scratch: String,
}

impl TextWrapCache {
    fn new(max_entries: usize) -> Self {
        Self {
            entries: HashMap::new(),
            order: VecDeque::new(),
            max_entries,
            scratch: String::new(),
        }
    }

    fn clear(&mut self) {
        self.entries.clear();
        self.order.clear();
    }

    fn get(&mut self, font: &str, max_width: f64, text: &str) -> Option<Rc<Vec<String>>> {
        if self.max_entries == 0 {
            return None;
        }
        let key = Self::build_key(&mut self.scratch, font, max_width, text);
        self.entries.get(key).cloned()
    }

    fn insert(
        &mut self,
        font: &str,
        max_width: f64,
        text: &str,
        lines: Vec<String>,
    ) -> Rc<Vec<String>> {
        if self.max_entries == 0 {
            return Rc::new(lines);
        }
        let key = Self::build_key(&mut self.scratch, font, max_width, text);
        if let Some(existing) = self.entries.get(key) {
            return Rc::clone(existing);
        }
        let lines = Rc::new(lines);
        let key_owned = key.to_string();
        self.entries.insert(key_owned.clone(), Rc::clone(&lines));
        self.order.push_back(key_owned);
        self.enforce_cap();
        lines
    }

    #[cfg(test)]
    fn len(&self) -> usize {
        self.entries.len()
    }

    fn build_key<'a>(scratch: &'a mut String, font: &str, max_width: f64, text: &str) -> &'a str {
        scratch.clear();
        scratch.reserve(font.len() + text.len() + 32);
        scratch.push_str(font);
        scratch.push('\n');
        let _ = write!(scratch, "{:x}", max_width.to_bits());
        scratch.push('\n');
        scratch.push_str(text);
        scratch.as_str()
    }

    fn enforce_cap(&mut self) {
        while self.entries.len() > self.max_entries {
            if let Some(oldest) = self.order.pop_front() {
                self.entries.remove(&oldest);
            } else {
                break;
            }
        }
    }
}

struct CacheCanvas {
    canvas: HtmlCanvasElement,
    ctx: CanvasRenderingContext2d,
    width: u32,
    height: u32,
}

/// Canvas 2D renderer implementing the RenderBackend trait
pub struct CanvasRenderer {
    canvas: HtmlCanvasElement,
    pub(crate) ctx: CanvasRenderingContext2d,
    width: u32,
    height: u32,
    dpr: f32,
    /// Cache for pattern fills (key: "pattern_type:fg_color:bg_color")
    pattern_cache: HashMap<String, CanvasPattern>,
    /// Cache for loaded images (key: image_id from EmbeddedImage)
    image_cache: HashMap<String, HtmlImageElement>,
    /// Cache for text measurements (key: "font\\ntext")
    text_measure_cache: TextMeasureCache,
    /// Cache for wrapped text layout (key: "font\\nwidth\\ntext")
    text_wrap_cache: TextWrapCache,
    /// Tile cache for scrollable area rendering
    tile_cache: TileCache,
    tile_cache_sheet: usize,
    tile_cache_layout_ptr: usize,
    frame_id: u64,
    deferred_prefetch_tiles: bool,
}

impl CanvasRenderer {
    /// Create a new Canvas renderer from an HtmlCanvasElement
    pub fn new(canvas: HtmlCanvasElement) -> Result<Self> {
        let ctx = canvas
            .get_context("2d")
            .map_err(|_| "Failed to get 2d context")?
            .ok_or("No 2d context available")?
            .dyn_into::<CanvasRenderingContext2d>()
            .map_err(|_| "Failed to cast to CanvasRenderingContext2d")?;

        let width = canvas.width();
        let height = canvas.height();

        Ok(Self {
            canvas,
            ctx,
            width,
            height,
            dpr: 1.0,
            pattern_cache: HashMap::new(),
            image_cache: HashMap::new(),
            text_measure_cache: TextMeasureCache::new(TEXT_MEASURE_CACHE_CAP),
            text_wrap_cache: TextWrapCache::new(TEXT_WRAP_CACHE_CAP),
            tile_cache: TileCache::new(),
            tile_cache_sheet: usize::MAX,
            tile_cache_layout_ptr: 0,
            frame_id: 0,
            deferred_prefetch_tiles: false,
        })
    }

    pub fn has_deferred_prefetch_tiles(&self) -> bool {
        self.deferred_prefetch_tiles
    }

    /// Set the CSS dimensions of the canvas element (logical pixels).
    #[allow(clippy::cast_possible_truncation)]
    pub fn set_canvas_css_size(&self, css_w: f32, css_h: f32) {
        let style = self.canvas.style();
        let _ = style.set_property("width", &format!("{}px", css_w));
        let _ = style.set_property("height", &format!("{}px", css_h));
    }

    /// Helper to get crisp pixel position for 1px lines
    pub(super) fn crisp(x: f64) -> f64 {
        x.floor() + 0.5
    }

    /// Draw a filled rectangle
    pub(super) fn fill_rect(&self, x: f64, y: f64, w: f64, h: f64, color: &str) {
        self.ctx.set_fill_style_str(color);
        self.ctx.fill_rect(x, y, w, h);
    }

    /// Draw a stroked line
    pub(super) fn stroke_line(&self, x1: f64, y1: f64, x2: f64, y2: f64, width: f64, color: &str) {
        self.ctx.begin_path();
        self.ctx.set_stroke_style_str(color);
        self.ctx.set_line_width(width);
        self.ctx.move_to(Self::crisp(x1), Self::crisp(y1));
        self.ctx.line_to(Self::crisp(x2), Self::crisp(y2));
        self.ctx.stroke();
    }

    /// Get the document for creating offscreen canvases
    fn get_document(&self) -> Option<Document> {
        web_sys::window()?.document()
    }

    /// Create a pattern canvas and draw the pattern on it
    fn create_pattern_canvas(
        &self,
        pattern_type: &str,
        fg_color: &str,
        bg_color: &str,
    ) -> Option<CanvasPattern> {
        let document = self.get_document()?;

        // Create offscreen canvas
        let pattern_canvas = document
            .create_element("canvas")
            .ok()?
            .dyn_into::<HtmlCanvasElement>()
            .ok()?;

        // Pattern size depends on the pattern type
        let (width, height) = Self::pattern_size(pattern_type);
        pattern_canvas.set_width(width);
        pattern_canvas.set_height(height);

        let pattern_ctx = pattern_canvas
            .get_context("2d")
            .ok()??
            .dyn_into::<CanvasRenderingContext2d>()
            .ok()?;

        // Draw the pattern
        Self::draw_pattern(
            &pattern_ctx,
            pattern_type,
            fg_color,
            bg_color,
            width,
            height,
        );

        // Create the repeating pattern
        self.ctx
            .create_pattern_with_html_canvas_element(&pattern_canvas, "repeat")
            .ok()?
    }

    /// Get pattern tile size for different pattern types
    fn pattern_size(pattern_type: &str) -> (u32, u32) {
        match pattern_type {
            "gray125" => (8, 8),
            "gray0625" => (16, 16),
            "darkgray" | "mediumgray" | "lightgray" => (4, 4),
            "darkhorizontal" | "lighthorizontal" => (1, 4),
            "darkvertical" | "lightvertical" => (4, 1),
            "darkdown" | "lightdown" | "darkup" | "lightup" => (8, 8),
            "darkgrid" | "lightgrid" => (4, 4),
            "darktrellis" | "lighttrellis" => (8, 8),
            _ => (8, 8),
        }
    }

    fn ensure_cache_canvas(
        document: &Document,
        dpr: f32,
        target: &mut Option<CacheCanvas>,
        width: f64,
        height: f64,
        changed: &mut bool,
    ) {
        if width <= 0.0 || height <= 0.0 {
            if target.is_some() {
                *target = None;
                *changed = true;
            }
            return;
        }
        #[allow(clippy::cast_possible_truncation, clippy::cast_sign_loss)]
        let w_px = (width * f64::from(dpr)).round().max(1.0) as u32;
        #[allow(clippy::cast_possible_truncation, clippy::cast_sign_loss)]
        let h_px = (height * f64::from(dpr)).round().max(1.0) as u32;
        let recreate = match target {
            Some(cache) => cache.width != w_px || cache.height != h_px,
            None => true,
        };
        if !recreate {
            return;
        }

        let Ok(element) = document.create_element("canvas") else {
            return;
        };
        let Ok(canvas) = element.dyn_into::<HtmlCanvasElement>() else {
            return;
        };
        canvas.set_width(w_px);
        canvas.set_height(h_px);
        let Ok(ctx) = canvas.get_context("2d") else {
            return;
        };
        let Some(ctx) = ctx else {
            return;
        };
        let Ok(ctx) = ctx.dyn_into::<CanvasRenderingContext2d>() else {
            return;
        };
        *target = Some(CacheCanvas {
            canvas,
            ctx,
            width: w_px,
            height: h_px,
        });
        *changed = true;
    }

    fn create_cache_canvas(&self, width: f64, height: f64) -> Option<CacheCanvas> {
        let document = self.get_document()?;
        let mut canvas = None;
        let mut changed = false;
        Self::ensure_cache_canvas(
            &document,
            self.dpr,
            &mut canvas,
            width,
            height,
            &mut changed,
        );
        canvas
    }

    fn reset_tile_cache_if_needed(&mut self, scale: f32) {
        if (self.tile_cache.scale - scale).abs() > f32::EPSILON
            || (self.tile_cache.dpr - self.dpr).abs() > f32::EPSILON
        {
            self.tile_cache.clear();
            self.tile_cache.scale = scale;
            self.tile_cache.dpr = self.dpr;
        }
    }

    fn evict_tile_cache_if_needed(&mut self) {
        while self.tile_cache.tiles.len() > TILE_CACHE_CAP {
            if let Some((oldest_key, _)) = self
                .tile_cache
                .tiles
                .iter()
                .min_by_key(|(_, entry)| entry.last_used)
                .map(|(k, v)| (*k, v.last_used))
            {
                self.tile_cache.tiles.remove(&oldest_key);
            } else {
                break;
            }
        }
    }

    /// Draw the actual pattern on the pattern canvas
    fn draw_pattern(
        ctx: &CanvasRenderingContext2d,
        pattern_type: &str,
        fg_color: &str,
        bg_color: &str,
        width: u32,
        height: u32,
    ) {
        let w = width as f64;
        let h = height as f64;

        // Fill background
        ctx.set_fill_style_str(bg_color);
        ctx.fill_rect(0.0, 0.0, w, h);

        // Draw foreground pattern
        ctx.set_fill_style_str(fg_color);
        ctx.set_stroke_style_str(fg_color);
        ctx.set_line_width(1.0);

        match pattern_type {
            // Gray patterns - dots at intervals
            "gray125" => {
                // 12.5% coverage - 1 pixel every 8 pixels
                ctx.fill_rect(0.0, 0.0, 1.0, 1.0);
            }
            "gray0625" => {
                // 6.25% coverage - 1 pixel every 16 pixels
                ctx.fill_rect(0.0, 0.0, 1.0, 1.0);
            }
            "darkgray" => {
                // ~75% coverage - checkerboard with more fg
                ctx.fill_rect(0.0, 0.0, 1.0, 1.0);
                ctx.fill_rect(2.0, 0.0, 1.0, 1.0);
                ctx.fill_rect(0.0, 2.0, 1.0, 1.0);
                ctx.fill_rect(2.0, 2.0, 1.0, 1.0);
                ctx.fill_rect(1.0, 1.0, 1.0, 1.0);
                ctx.fill_rect(3.0, 1.0, 1.0, 1.0);
                ctx.fill_rect(1.0, 3.0, 1.0, 1.0);
                ctx.fill_rect(3.0, 3.0, 1.0, 1.0);
                // Add extra coverage
                ctx.fill_rect(0.0, 1.0, 1.0, 1.0);
                ctx.fill_rect(2.0, 1.0, 1.0, 1.0);
                ctx.fill_rect(0.0, 3.0, 1.0, 1.0);
                ctx.fill_rect(2.0, 3.0, 1.0, 1.0);
            }
            "mediumgray" => {
                // ~50% coverage - checkerboard
                ctx.fill_rect(0.0, 0.0, 1.0, 1.0);
                ctx.fill_rect(2.0, 0.0, 1.0, 1.0);
                ctx.fill_rect(0.0, 2.0, 1.0, 1.0);
                ctx.fill_rect(2.0, 2.0, 1.0, 1.0);
                ctx.fill_rect(1.0, 1.0, 1.0, 1.0);
                ctx.fill_rect(3.0, 1.0, 1.0, 1.0);
                ctx.fill_rect(1.0, 3.0, 1.0, 1.0);
                ctx.fill_rect(3.0, 3.0, 1.0, 1.0);
            }
            "lightgray" => {
                // ~25% coverage - sparse dots
                ctx.fill_rect(0.0, 0.0, 1.0, 1.0);
                ctx.fill_rect(2.0, 2.0, 1.0, 1.0);
            }

            // Horizontal lines
            "darkhorizontal" => {
                ctx.fill_rect(0.0, 0.0, 1.0, 2.0);
            }
            "lighthorizontal" => {
                ctx.fill_rect(0.0, 0.0, 1.0, 1.0);
            }

            // Vertical lines
            "darkvertical" => {
                ctx.fill_rect(0.0, 0.0, 2.0, 1.0);
            }
            "lightvertical" => {
                ctx.fill_rect(0.0, 0.0, 1.0, 1.0);
            }

            // Diagonal down (top-left to bottom-right)
            "darkdown" => {
                for i in 0..8 {
                    ctx.fill_rect(i as f64, i as f64, 2.0, 2.0);
                }
            }
            "lightdown" => {
                for i in 0..8 {
                    ctx.fill_rect(i as f64, i as f64, 1.0, 1.0);
                }
            }

            // Diagonal up (bottom-left to top-right)
            "darkup" => {
                for i in 0..8 {
                    ctx.fill_rect(i as f64, (7 - i) as f64, 2.0, 2.0);
                }
            }
            "lightup" => {
                for i in 0..8 {
                    ctx.fill_rect(i as f64, (7 - i) as f64, 1.0, 1.0);
                }
            }

            // Grid (horizontal + vertical)
            "darkgrid" => {
                // Horizontal line
                ctx.fill_rect(0.0, 0.0, w, 2.0);
                // Vertical line
                ctx.fill_rect(0.0, 0.0, 2.0, h);
            }
            "lightgrid" => {
                // Horizontal line
                ctx.fill_rect(0.0, 0.0, w, 1.0);
                // Vertical line
                ctx.fill_rect(0.0, 0.0, 1.0, h);
            }

            // Trellis (diagonal crosshatch)
            "darktrellis" => {
                // Diagonal lines both ways
                for i in 0..8 {
                    ctx.fill_rect(i as f64, i as f64, 2.0, 2.0);
                    ctx.fill_rect(i as f64, (7 - i) as f64, 2.0, 2.0);
                }
            }
            "lighttrellis" => {
                for i in 0..8 {
                    ctx.fill_rect(i as f64, i as f64, 1.0, 1.0);
                    ctx.fill_rect(i as f64, (7 - i) as f64, 1.0, 1.0);
                }
            }

            _ => {
                // Unknown pattern - just fill solid with fg
                ctx.fill_rect(0.0, 0.0, w, h);
            }
        }
    }

    /// Get or create a cached pattern
    fn get_or_create_pattern(
        &mut self,
        pattern_type: &str,
        fg_color: &str,
        bg_color: &str,
    ) -> Option<CanvasPattern> {
        let cache_key = format!("{}:{}:{}", pattern_type, fg_color, bg_color);

        if let Some(pattern) = self.pattern_cache.get(&cache_key) {
            return Some(pattern.clone());
        }

        // Create new pattern
        let pattern = self.create_pattern_canvas(pattern_type, fg_color, bg_color)?;
        self.pattern_cache.insert(cache_key, pattern.clone());
        Some(pattern)
    }

    /// Check if a pattern type needs pattern fill (not solid or none)
    fn needs_pattern_fill(pattern_type: Option<&str>) -> bool {
        !matches!(pattern_type, None | Some("none") | Some("solid"))
    }

    /// Fill a cell with a pattern
    fn fill_pattern(&mut self, x: f64, y: f64, w: f64, h: f64, style: &CellStyleData) {
        let pattern_type = style.pattern_type.as_deref().unwrap_or("solid");

        // For solid or none, use simple fill
        if !Self::needs_pattern_fill(Some(pattern_type)) {
            if let Some(ref bg_color_str) = style.bg_color {
                if let Some(bg_color) = parse_color(bg_color_str) {
                    self.fill_rect(x, y, w, h, &bg_color);
                }
            }
            return;
        }

        // Get colors for pattern
        // fg_color is the pattern foreground (the lines/dots)
        // bg_color is the background behind the pattern
        let fg_color = style
            .pattern_fg_color
            .as_ref()
            .and_then(|c| parse_color(c))
            .unwrap_or_else(|| palette::BLACK.to_string());

        let bg_color = style
            .pattern_bg_color
            .as_ref()
            .and_then(|c| parse_color(c))
            .or_else(|| style.bg_color.as_ref().and_then(|c| parse_color(c)))
            .unwrap_or_else(|| palette::WHITE.to_string());

        // Try to create/get the pattern
        if let Some(pattern) = self.get_or_create_pattern(pattern_type, &fg_color, &bg_color) {
            // Translate to cell position so the pattern tiles from the cell's top-left corner,
            // not from the canvas origin. This prevents pattern movement when scrolling.
            self.ctx.save();
            let _ = self.ctx.translate(x, y);
            self.ctx.set_fill_style_canvas_pattern(&pattern);
            self.ctx.fill_rect(0.0, 0.0, w, h);
            self.ctx.restore();
        } else {
            // Fallback to bg_color if pattern creation fails
            if let Some(ref bg_color_str) = style.bg_color {
                if let Some(bg_color) = parse_color(bg_color_str) {
                    self.fill_rect(x, y, w, h, &bg_color);
                }
            }
        }
    }

    /// Render cell backgrounds (including pattern fills)
    fn render_cell_backgrounds(
        &mut self,
        cells: &[&CellRenderData],
        layout: &SheetLayout,
        viewport: &Viewport,
        style_cache: &[Option<CellStyleData>],
        default_style: &Option<CellStyleData>,
    ) {
        for &cell in cells {
            let Some(style) = Self::resolve_cell_style(cell, style_cache, default_style) else {
                continue;
            };
            let rect = layout.cell_rect(cell.row, cell.col);
            if rect.skip || rect.width <= 0.0 || rect.height <= 0.0 {
                continue;
            }

            let (sx, sy) = viewport.to_screen_frozen(rect.x, rect.y, cell.row, cell.col, layout);
            let x = f64::from(sx);
            let y = f64::from(sy);
            let w = f64::from(rect.width);
            let h = f64::from(rect.height);

            // Check if this cell has a pattern fill
            let has_pattern = Self::needs_pattern_fill(style.pattern_type.as_deref());
            let has_bg = style.bg_color.is_some();

            if has_pattern {
                // Use pattern fill
                self.fill_pattern(x, y, w, h, style);
            } else if has_bg {
                // Use solid color fill
                if let Some(bg_color) = style.bg_color.as_ref().and_then(|c| parse_color(c)) {
                    self.fill_rect(x, y, w, h, &bg_color);
                }
            }
        }
    }

    /// Render grid lines using a single path for performance
    /// Skips drawing line segments inside merged cell regions
    fn render_grid_lines(
        &self,
        layout: &SheetLayout,
        viewport: &Viewport,
        start_row: u32,
        end_row: u32,
        start_col: u32,
        end_col: u32,
    ) {
        let content_height = f64::from(viewport.height) - SCROLLBAR_SIZE;
        let content_width = f64::from(viewport.width) - SCROLLBAR_SIZE;

        self.ctx.begin_path();
        self.ctx.set_stroke_style_str(palette::GRID_LINE);
        self.ctx.set_line_width(1.0);

        // Vertical lines (column separators at column boundary col)
        // A vertical line at column col is inside a merge if: origin_col < col < origin_col + col_span
        for col in start_col..=end_col + 1 {
            if let Some(&x) = layout.col_positions.get(col as usize) {
                let sx = f64::from(viewport.screen_x_for_grid(x, col, layout));
                if sx >= 0.0 && sx <= content_width {
                    let skip_ranges = layout
                        .merge_vline_skips
                        .get(col as usize)
                        .map(|r| r.as_slice())
                        .unwrap_or(&[]);

                    // Draw vertical line segments, skipping merged areas
                    let mut current_row = start_row;
                    for &(skip_start, skip_end) in skip_ranges {
                        // Draw segment from current_row to skip_start
                        if current_row < skip_start && skip_start <= end_row + 1 {
                            let y1 = layout
                                .row_positions
                                .get(current_row as usize)
                                .copied()
                                .unwrap_or(0.0);
                            let y2 = layout
                                .row_positions
                                .get(skip_start as usize)
                                .copied()
                                .unwrap_or(y1);
                            let sy1 =
                                f64::from(viewport.screen_y_for_grid(y1, current_row, layout))
                                    .max(0.0);
                            let sy2 = f64::from(viewport.screen_y_for_grid(y2, skip_start, layout))
                                .min(content_height);
                            if sy1 < sy2 {
                                self.ctx.move_to(Self::crisp(sx), sy1);
                                self.ctx.line_to(Self::crisp(sx), sy2);
                            }
                        }
                        // Move past the skip range
                        if skip_end > current_row {
                            current_row = skip_end;
                        }
                    }

                    // Draw remaining segment from current_row to end
                    if current_row <= end_row + 1 {
                        let y1 = layout
                            .row_positions
                            .get(current_row as usize)
                            .copied()
                            .unwrap_or(0.0);
                        let y2 = layout
                            .row_positions
                            .get((end_row + 1) as usize)
                            .copied()
                            .unwrap_or(y1);
                        let sy1 =
                            f64::from(viewport.screen_y_for_grid(y1, current_row, layout)).max(0.0);
                        let sy2 = f64::from(viewport.screen_y_for_grid(y2, end_row + 1, layout))
                            .min(content_height);
                        if sy1 < sy2 {
                            self.ctx.move_to(Self::crisp(sx), sy1);
                            self.ctx.line_to(Self::crisp(sx), sy2);
                        }
                    }
                }
            }
        }

        // Horizontal lines (row separators at row boundary row)
        // A horizontal line at row row is inside a merge if: origin_row < row < origin_row + row_span
        for row in start_row..=end_row + 1 {
            if let Some(&y) = layout.row_positions.get(row as usize) {
                let sy = f64::from(viewport.screen_y_for_grid(y, row, layout));
                if sy >= 0.0 && sy <= content_height {
                    let skip_ranges = layout
                        .merge_hline_skips
                        .get(row as usize)
                        .map(|r| r.as_slice())
                        .unwrap_or(&[]);

                    // Draw horizontal line segments, skipping merged areas
                    let mut current_col = start_col;
                    for &(skip_start, skip_end) in skip_ranges {
                        // Draw segment from current_col to skip_start
                        if current_col < skip_start && skip_start <= end_col + 1 {
                            let x1 = layout
                                .col_positions
                                .get(current_col as usize)
                                .copied()
                                .unwrap_or(0.0);
                            let x2 = layout
                                .col_positions
                                .get(skip_start as usize)
                                .copied()
                                .unwrap_or(x1);
                            let sx1 =
                                f64::from(viewport.screen_x_for_grid(x1, current_col, layout))
                                    .max(0.0);
                            let sx2 = f64::from(viewport.screen_x_for_grid(x2, skip_start, layout))
                                .min(content_width);
                            if sx1 < sx2 {
                                self.ctx.move_to(sx1, Self::crisp(sy));
                                self.ctx.line_to(sx2, Self::crisp(sy));
                            }
                        }
                        // Move past the skip range
                        if skip_end > current_col {
                            current_col = skip_end;
                        }
                    }

                    // Draw remaining segment from current_col to end
                    if current_col <= end_col + 1 {
                        let x1 = layout
                            .col_positions
                            .get(current_col as usize)
                            .copied()
                            .unwrap_or(0.0);
                        let x2 = layout
                            .col_positions
                            .get((end_col + 1) as usize)
                            .copied()
                            .unwrap_or(x1);
                        let sx1 =
                            f64::from(viewport.screen_x_for_grid(x1, current_col, layout)).max(0.0);
                        let sx2 = f64::from(viewport.screen_x_for_grid(x2, end_col + 1, layout))
                            .min(content_width);
                        if sx1 < sx2 {
                            self.ctx.move_to(sx1, Self::crisp(sy));
                            self.ctx.line_to(sx2, Self::crisp(sy));
                        }
                    }
                }
            }
        }

        self.ctx.stroke();
    }

    /// Render grid lines for frozen rows region (top row strip, excluding corner)
    fn render_grid_lines_for_frozen_rows(
        &self,
        layout: &SheetLayout,
        viewport: &Viewport,
        start_col: u32,
        end_col: u32,
    ) {
        let frozen_rows = layout.frozen_rows;
        if frozen_rows == 0 {
            return;
        }

        let frozen_height = f64::from(layout.frozen_rows_height());
        let content_width = f64::from(viewport.width) - SCROLLBAR_SIZE;
        // Limit grid lines to actual data width, not full viewport
        let data_width = f64::from(layout.total_width());
        let max_x = content_width.min(data_width);

        self.ctx.begin_path();
        self.ctx.set_stroke_style_str(palette::GRID_LINE);
        self.ctx.set_line_width(1.0);

        // Vertical lines in frozen rows area
        for col in start_col..=end_col + 1 {
            if col < layout.frozen_cols {
                continue; // Skip corner area
            }
            if let Some(&x) = layout.col_positions.get(col as usize) {
                let sx = f64::from(viewport.screen_x_for_grid(x, col, layout));
                if sx >= 0.0 && sx <= max_x {
                    self.ctx.move_to(Self::crisp(sx), 0.0);
                    self.ctx.line_to(Self::crisp(sx), frozen_height);
                }
            }
        }

        // Horizontal lines in frozen rows area
        for row in 0..=frozen_rows {
            if let Some(&y) = layout.row_positions.get(row as usize) {
                let sy = f64::from(y); // Frozen rows don't scroll
                if sy >= 0.0 && sy <= frozen_height {
                    let start_x = f64::from(layout.frozen_cols_width());
                    self.ctx.move_to(start_x, Self::crisp(sy));
                    self.ctx.line_to(max_x, Self::crisp(sy));
                }
            }
        }

        self.ctx.stroke();
    }

    /// Render grid lines for frozen columns region (left column strip, excluding corner)
    fn render_grid_lines_for_frozen_cols(
        &self,
        layout: &SheetLayout,
        viewport: &Viewport,
        start_row: u32,
        end_row: u32,
    ) {
        let frozen_cols = layout.frozen_cols;
        if frozen_cols == 0 {
            return;
        }

        let frozen_width = f64::from(layout.frozen_cols_width());
        let content_height = f64::from(viewport.height) - SCROLLBAR_SIZE;

        self.ctx.begin_path();
        self.ctx.set_stroke_style_str(palette::GRID_LINE);
        self.ctx.set_line_width(1.0);

        // Horizontal lines in frozen cols area
        for row in start_row..=end_row + 1 {
            if row < layout.frozen_rows {
                continue; // Skip corner area
            }
            if let Some(&y) = layout.row_positions.get(row as usize) {
                let sy = f64::from(viewport.screen_y_for_grid(y, row, layout));
                if sy >= 0.0 && sy <= content_height {
                    self.ctx.move_to(0.0, Self::crisp(sy));
                    self.ctx.line_to(frozen_width, Self::crisp(sy));
                }
            }
        }

        // Vertical lines in frozen cols area
        for col in 0..=frozen_cols {
            if let Some(&x) = layout.col_positions.get(col as usize) {
                let sx = f64::from(x); // Frozen cols don't scroll
                if sx >= 0.0 && sx <= frozen_width {
                    let start_y = f64::from(layout.frozen_rows_height());
                    self.ctx.move_to(Self::crisp(sx), start_y);
                    self.ctx.line_to(Self::crisp(sx), content_height);
                }
            }
        }

        self.ctx.stroke();
    }

    /// Render grid lines for the corner region (frozen rows AND cols intersection)
    fn render_grid_lines_for_corner(&self, layout: &SheetLayout, viewport: &Viewport) {
        let frozen_rows = layout.frozen_rows;
        let frozen_cols = layout.frozen_cols;
        if frozen_rows == 0 || frozen_cols == 0 {
            return;
        }

        let frozen_width = f64::from(layout.frozen_cols_width());
        let frozen_height = f64::from(layout.frozen_rows_height());
        let _ = viewport; // Unused but kept for API consistency

        self.ctx.begin_path();
        self.ctx.set_stroke_style_str(palette::GRID_LINE);
        self.ctx.set_line_width(1.0);

        // Vertical lines in corner
        for col in 0..=frozen_cols {
            if let Some(&x) = layout.col_positions.get(col as usize) {
                let sx = f64::from(x);
                if sx >= 0.0 && sx <= frozen_width {
                    self.ctx.move_to(Self::crisp(sx), 0.0);
                    self.ctx.line_to(Self::crisp(sx), frozen_height);
                }
            }
        }

        // Horizontal lines in corner
        for row in 0..=frozen_rows {
            if let Some(&y) = layout.row_positions.get(row as usize) {
                let sy = f64::from(y);
                if sy >= 0.0 && sy <= frozen_height {
                    self.ctx.move_to(0.0, Self::crisp(sy));
                    self.ctx.line_to(frozen_width, Self::crisp(sy));
                }
            }
        }

        self.ctx.stroke();
    }

    /// Render cell borders
    fn render_cell_borders(
        &self,
        cells: &[&CellRenderData],
        layout: &SheetLayout,
        viewport: &Viewport,
        style_cache: &[Option<CellStyleData>],
        default_style: &Option<CellStyleData>,
    ) {
        for &cell in cells {
            let Some(style) = Self::resolve_cell_style(cell, style_cache, default_style) else {
                continue;
            };

            let rect = layout.cell_rect(cell.row, cell.col);
            if rect.skip {
                continue;
            }

            let (sx, sy) = viewport.to_screen_frozen(rect.x, rect.y, cell.row, cell.col, layout);
            let x1 = f64::from(sx);
            let y1 = f64::from(sy);
            let x2 = f64::from(sx + rect.width * viewport.scale);
            let y2 = f64::from(sy + rect.height * viewport.scale);

            // Top border
            if let Some(ref border) = style.border_top {
                self.draw_border_line(x1, y1, x2, y1, border);
            }

            // Right border
            if let Some(ref border) = style.border_right {
                self.draw_border_line(x2, y1, x2, y2, border);
            }

            // Bottom border
            if let Some(ref border) = style.border_bottom {
                self.draw_border_line(x1, y2, x2, y2, border);
            }

            // Left border
            if let Some(ref border) = style.border_left {
                self.draw_border_line(x1, y1, x1, y2, border);
            }

            // Diagonal down (top-left to bottom-right)
            if let Some(ref border) = style.border_diagonal_down {
                self.draw_border_line(x1, y1, x2, y2, border);
            }

            // Diagonal up (bottom-left to top-right)
            if let Some(ref border) = style.border_diagonal_up {
                self.draw_border_line(x1, y2, x2, y1, border);
            }
        }
    }

    fn draw_border_line(&self, x1: f64, y1: f64, x2: f64, y2: f64, border: &BorderStyleData) {
        let color = border
            .color
            .as_ref()
            .and_then(|c| parse_color(c))
            .unwrap_or_else(|| palette::BLACK.to_string());
        let width = border.width();
        self.stroke_line(x1, y1, x2, y2, width, &color);
    }

    fn resolve_cell_style<'a>(
        cell: &'a CellRenderData,
        style_cache: &'a [Option<CellStyleData>],
        default_style: &'a Option<CellStyleData>,
    ) -> Option<&'a CellStyleData> {
        if let Some(ref style) = cell.style_override {
            return Some(style);
        }
        if let Some(idx) = cell.style_idx {
            return style_cache.get(idx).and_then(|s| s.as_ref());
        }
        default_style.as_ref()
    }

    fn measure_text_cached(&mut self, text: &str, font: &str) -> f64 {
        if let Some(width) = self.text_measure_cache.get(font, text) {
            return width;
        }
        let width = self
            .ctx
            .measure_text(text)
            .map(|m| m.width())
            .unwrap_or(0.0);
        self.text_measure_cache.insert(font, text, width);
        width
    }

    fn content_bounds(
        &self,
        viewport: &Viewport,
        header_offset_x: f64,
        header_offset_y: f64,
    ) -> (f64, f64, f64, f64) {
        let content_width = (f64::from(viewport.width) - SCROLLBAR_SIZE - header_offset_x).max(0.0);
        let content_height =
            (f64::from(viewport.height) - SCROLLBAR_SIZE - header_offset_y).max(0.0);
        (
            header_offset_x,
            header_offset_y,
            content_width,
            content_height,
        )
    }

    #[allow(clippy::too_many_arguments, clippy::cast_possible_truncation)]
    fn render_scrollable_tiles(
        &mut self,
        params: &RenderParams<'_>,
        cells: &[&CellRenderData],
        layout: &SheetLayout,
        viewport: &Viewport,
        content_width: f64,
        content_height: f64,
        default_font: &str,
    ) -> TileRenderResult {
        let frozen_width = f64::from(layout.frozen_cols_width());
        let frozen_height = f64::from(layout.frozen_rows_height());
        let scroll_w = (content_width - frozen_width).max(0.0);
        let scroll_h = (content_height - frozen_height).max(0.0);
        if scroll_w <= 0.0 || scroll_h <= 0.0 {
            return TileRenderResult {
                rendered_visible: true,
                deferred_prefetch: false,
            };
        }

        self.reset_tile_cache_if_needed(viewport.scale);

        let scale = f64::from(viewport.scale);
        let scroll_x = f64::from(viewport.scroll_x);
        let scroll_y = f64::from(viewport.scroll_y);
        let origin_x = frozen_width;
        let origin_y = frozen_height;

        let visible_tile_x_start = ((scroll_x - origin_x) / TILE_SIZE).floor() as i32;
        let visible_tile_x_end = ((scroll_x + scroll_w - origin_x) / TILE_SIZE).floor() as i32;
        let visible_tile_y_start = ((scroll_y - origin_y) / TILE_SIZE).floor() as i32;
        let visible_tile_y_end = ((scroll_y + scroll_h - origin_y) / TILE_SIZE).floor() as i32;
        let prefetch_tiles = i32::try_from(params.tile_prefetch).unwrap_or(i32::MAX);
        let tile_x_start = visible_tile_x_start.saturating_sub(prefetch_tiles);
        let tile_x_end = visible_tile_x_end.saturating_add(prefetch_tiles);
        let tile_y_start = visible_tile_y_start.saturating_sub(prefetch_tiles);
        let tile_y_end = visible_tile_y_end.saturating_add(prefetch_tiles);

        let tile_logical = TILE_SIZE * scale;
        if tile_logical <= 0.0 {
            return TileRenderResult {
                rendered_visible: true,
                deferred_prefetch: false,
            };
        }
        #[allow(clippy::cast_possible_truncation, clippy::cast_sign_loss)]
        let tile_px = (tile_logical * f64::from(self.dpr)).round().max(1.0) as u32;

        let smoothing = self.ctx.image_smoothing_enabled();
        self.ctx.set_image_smoothing_enabled(false);
        self.ctx.save();
        self.ctx.begin_path();
        self.ctx
            .rect(frozen_width, frozen_height, scroll_w, scroll_h);
        self.ctx.clip();

        let mut rendered_visible = true;
        let mut deferred_prefetch = false;
        let mut remaining_prefetch_budget = params.max_prefetch_tile_renders;
        for ty in tile_y_start..=tile_y_end {
            for tx in tile_x_start..=tile_x_end {
                let tile_origin_x = origin_x + f64::from(tx) * TILE_SIZE;
                let tile_origin_y = origin_y + f64::from(ty) * TILE_SIZE;
                let screen_x = frozen_width + (tile_origin_x - scroll_x) * scale;
                let screen_y = frozen_height + (tile_origin_y - scroll_y) * scale;
                let key = TileKey { x: tx, y: ty };
                let in_visible_x = (visible_tile_x_start..=visible_tile_x_end).contains(&tx);
                let in_visible_y = (visible_tile_y_start..=visible_tile_y_end).contains(&ty);
                let is_prefetch_only = !(in_visible_x && in_visible_y);

                let mut needs_render = false;
                let mut needs_canvas = false;
                let mut entry_missing = false;

                if let Some(entry) = self.tile_cache.tiles.get(&key) {
                    if entry.canvas.width != tile_px || entry.canvas.height != tile_px {
                        needs_canvas = true;
                    }
                } else {
                    needs_canvas = true;
                    entry_missing = true;
                }

                if needs_canvas {
                    if is_prefetch_only && matches!(remaining_prefetch_budget, Some(0)) {
                        deferred_prefetch = true;
                        continue;
                    }
                    if let Some(canvas) = self.create_cache_canvas(tile_logical, tile_logical) {
                        self.tile_cache.tiles.insert(
                            key,
                            TileCacheEntry {
                                canvas,
                                last_used: self.frame_id,
                            },
                        );
                        needs_render = true;
                    } else if entry_missing {
                        if !is_prefetch_only {
                            rendered_visible = false;
                        }
                        continue;
                    }
                }

                if needs_render {
                    if is_prefetch_only {
                        if let Some(budget) = remaining_prefetch_budget.as_mut() {
                            *budget = budget.saturating_sub(1);
                        }
                    }
                    if let Some(mut entry) = self.tile_cache.tiles.remove(&key) {
                        self.render_tile_content(
                            &mut entry.canvas,
                            params,
                            cells,
                            layout,
                            viewport,
                            tile_origin_x,
                            tile_origin_y,
                            tile_logical,
                            default_font,
                        );
                        entry.last_used = self.frame_id;
                        self.tile_cache.tiles.insert(key, entry);
                    }
                } else if let Some(entry) = self.tile_cache.tiles.get_mut(&key) {
                    entry.last_used = self.frame_id;
                }

                if let Some(entry) = self.tile_cache.tiles.get(&key) {
                    let _ = self.ctx.draw_image_with_html_canvas_element_and_dw_and_dh(
                        &entry.canvas.canvas,
                        screen_x,
                        screen_y,
                        tile_logical,
                        tile_logical,
                    );
                } else if !is_prefetch_only {
                    rendered_visible = false;
                }
            }
        }

        self.ctx.restore();
        self.ctx.set_image_smoothing_enabled(smoothing);
        self.evict_tile_cache_if_needed();
        TileRenderResult {
            rendered_visible,
            deferred_prefetch,
        }
    }

    fn with_canvas_context<F>(
        &mut self,
        canvas: &HtmlCanvasElement,
        ctx: &CanvasRenderingContext2d,
        f: F,
    ) where
        F: FnOnce(&mut Self),
    {
        let prev_canvas = self.canvas.clone();
        let prev_ctx = self.ctx.clone();
        self.canvas = canvas.clone();
        self.ctx = ctx.clone();
        f(self);
        self.canvas = prev_canvas;
        self.ctx = prev_ctx;
    }

    #[allow(clippy::too_many_arguments, clippy::cast_possible_truncation)]
    fn render_tile_content(
        &mut self,
        tile_canvas: &mut CacheCanvas,
        params: &RenderParams<'_>,
        cells: &[&CellRenderData],
        layout: &SheetLayout,
        viewport: &Viewport,
        tile_origin_x: f64,
        tile_origin_y: f64,
        tile_logical: f64,
        default_font: &str,
    ) {
        let tile_x1 = tile_origin_x + TILE_SIZE;
        let tile_y1 = tile_origin_y + TILE_SIZE;
        let start_row = layout
            .row_at_y(tile_origin_y as f32)
            .unwrap_or(layout.max_row)
            .max(layout.frozen_rows);
        let end_row = layout
            .row_at_y(tile_y1 as f32)
            .unwrap_or(layout.max_row)
            .min(layout.max_row);
        let start_col = layout
            .col_at_x(tile_origin_x as f32)
            .unwrap_or(layout.max_col)
            .max(layout.frozen_cols);
        let end_col = layout
            .col_at_x(tile_x1 as f32)
            .unwrap_or(layout.max_col)
            .min(layout.max_col);

        if start_row > end_row || start_col > end_col {
            return;
        }

        let tile_cells: Vec<&CellRenderData> = cells
            .iter()
            .copied()
            .filter(|c| {
                c.row >= start_row && c.row <= end_row && c.col >= start_col && c.col <= end_col
            })
            .collect();

        let frozen_width = f64::from(layout.frozen_cols_width());
        let frozen_height = f64::from(layout.frozen_rows_height());
        let mut tile_viewport = viewport.clone();
        tile_viewport.scroll_x = (tile_origin_x + frozen_width) as f32;
        tile_viewport.scroll_y = (tile_origin_y + frozen_height) as f32;
        tile_viewport.width = (tile_logical + SCROLLBAR_SIZE) as f32;
        tile_viewport.height = (tile_logical + SCROLLBAR_SIZE) as f32;

        self.with_canvas_context(&tile_canvas.canvas, &tile_canvas.ctx, |renderer| {
            let _ = renderer.ctx.reset_transform();
            let _ = renderer
                .ctx
                .scale(f64::from(renderer.dpr), f64::from(renderer.dpr));
            renderer
                .ctx
                .clear_rect(0.0, 0.0, tile_logical, tile_logical);
            renderer.ctx.save();
            renderer.ctx.begin_path();
            renderer.ctx.rect(0.0, 0.0, tile_logical, tile_logical);
            renderer.ctx.clip();
            renderer.fill_rect(0.0, 0.0, tile_logical, tile_logical, palette::WHITE);

            renderer.render_conditional_formatting(
                &tile_cells,
                params.conditional_formatting,
                params.dxf_styles,
                layout,
                &tile_viewport,
                params.conditional_formatting_cache,
            );
            let dxf_overrides = Self::collect_cell_is_dxf_overrides(
                &tile_cells,
                params.conditional_formatting,
                params.dxf_styles,
                params.conditional_formatting_cache,
            );
            renderer.render_cell_backgrounds(
                &tile_cells,
                layout,
                &tile_viewport,
                params.style_cache,
                params.default_style,
            );
            renderer.render_grid_lines(
                layout,
                &tile_viewport,
                start_row,
                end_row,
                start_col,
                end_col,
            );
            renderer.render_cell_borders(
                &tile_cells,
                layout,
                &tile_viewport,
                params.style_cache,
                params.default_style,
            );
            renderer.render_cell_text(
                &tile_cells,
                layout,
                &tile_viewport,
                default_font,
                &dxf_overrides,
                params.style_cache,
                params.default_style,
            );
            renderer.ctx.restore();
        });
    }

    /// Render cell text
    #[allow(clippy::too_many_arguments)]
    fn render_cell_text(
        &mut self,
        cells: &[&CellRenderData],
        layout: &SheetLayout,
        viewport: &Viewport,
        default_font: &str,
        dxf_overrides: &HashMap<(u32, u32), &crate::types::DxfStyle>,
        style_cache: &[Option<CellStyleData>],
        default_style: &Option<CellStyleData>,
    ) {
        let mut last_font = String::new();
        let mut last_color = String::new();

        for &cell in cells {
            // Skip cells with no value (unless they have rich_text)
            let has_value = cell.value.as_ref().is_some_and(|v| !v.is_empty());
            let has_rich_text = cell.rich_text.as_ref().is_some_and(|runs| !runs.is_empty());

            if !has_value && !has_rich_text {
                continue;
            }

            let rect = layout.cell_rect(cell.row, cell.col);
            if rect.skip || f64::from(rect.width) <= CELL_PADDING * 2.0 {
                continue;
            }

            let style = Self::resolve_cell_style(cell, style_cache, default_style);
            let (sx, sy) = viewport.to_screen_frozen(rect.x, rect.y, cell.row, cell.col, layout);

            // Check for DXF override from cellIs conditional formatting
            let dxf_override = dxf_overrides.get(&(cell.row, cell.col));

            // Get style properties
            let font_size = style.and_then(|s| s.font_size).unwrap_or(11.0);
            // Check if cell has hyperlink - use blue color unless overridden
            // DXF font_color takes precedence over cell style
            let font_color = if let Some(dxf) = dxf_override {
                if let Some(ref dxf_color) = dxf.font_color {
                    parse_color(dxf_color).unwrap_or_else(|| palette::BLACK.to_string())
                } else if cell.has_hyperlink == Some(true) {
                    style
                        .and_then(|s| s.font_color.as_ref())
                        .and_then(|c| parse_color(c))
                        .unwrap_or_else(|| "#0563C1".to_string())
                } else {
                    style
                        .and_then(|s| s.font_color.as_ref())
                        .and_then(|c| parse_color(c))
                        .unwrap_or_else(|| palette::BLACK.to_string())
                }
            } else if cell.has_hyperlink == Some(true) {
                style
                    .and_then(|s| s.font_color.as_ref())
                    .and_then(|c| parse_color(c))
                    .unwrap_or_else(|| "#0563C1".to_string()) // Excel hyperlink blue
            } else {
                style
                    .and_then(|s| s.font_color.as_ref())
                    .and_then(|c| parse_color(c))
                    .unwrap_or_else(|| palette::BLACK.to_string())
            };
            // Also ensure underline is set for hyperlinks or DXF override
            let has_underline = dxf_override.and_then(|d| d.underline).unwrap_or(false)
                || style.and_then(|s| s.underline).unwrap_or(false)
                || cell.has_hyperlink == Some(true);
            // DXF bold/italic override cell style
            let bold = dxf_override
                .and_then(|d| d.bold)
                .unwrap_or_else(|| style.and_then(|s| s.bold).unwrap_or(false));
            let italic = dxf_override
                .and_then(|d| d.italic)
                .unwrap_or_else(|| style.and_then(|s| s.italic).unwrap_or(false));
            let font_family = style
                .and_then(|s| s.font_family.as_deref())
                .unwrap_or(default_font);
            let align_h = style.and_then(|s| s.align_h.as_deref()).unwrap_or("left");
            let align_v = style.and_then(|s| s.align_v.as_deref()).unwrap_or("center");
            let wrap_text = style.and_then(|s| s.wrap_text).unwrap_or(false);

            // Build font string
            let font_style = if italic { "italic " } else { "" };
            let font_weight = if bold { "bold " } else { "" };
            let font_string = format!(
                "{}{}{}px {}",
                font_style, font_weight, font_size, font_family
            );
            if font_string != last_font {
                self.ctx.set_font(&font_string);
                last_font = font_string.clone();
            }
            if font_color != last_color {
                self.ctx.set_fill_style_str(&font_color);
                last_color = font_color.clone();
            }

            // Calculate indent pixels (each indent level adds 10 pixels)
            let indent_pixels = style.and_then(|s| s.indent).unwrap_or(0) as f64 * 10.0;

            // Calculate available width
            let max_width = f64::from(rect.width) - CELL_PADDING * 2.0 - indent_pixels;

            // Check if we have rich text to render
            if let Some(ref runs) = cell.rich_text {
                if !runs.is_empty() {
                    // Rich text needs clipping (multiple runs with varying fonts)
                    self.ctx.save();
                    self.ctx.begin_path();
                    self.ctx.rect(
                        f64::from(sx),
                        f64::from(sy),
                        f64::from(rect.width),
                        f64::from(rect.height),
                    );
                    self.ctx.clip();
                    self.render_rich_text(runs.as_ref(), style, sx, sy, &rect, default_font);
                    last_font.clear();
                    last_color.clear();
                    self.ctx.restore();
                    continue;
                }
            }

            // Fall through to plain text rendering
            let Some(ref value) = cell.value else {
                continue;
            };

            // Get rotation if present
            let rotation = style.and_then(|s| s.rotation).unwrap_or(0);

            // Only clip for wrapped or rotated text — plain text is
            // pre-truncated by truncate_text() so clipping is redundant.
            let needs_clip = wrap_text || rotation != 0;
            if needs_clip {
                self.ctx.save();
                self.ctx.begin_path();
                self.ctx.rect(
                    f64::from(sx),
                    f64::from(sy),
                    f64::from(rect.width),
                    f64::from(rect.height),
                );
                self.ctx.clip();
            }

            if wrap_text && rotation == 0 {
                // Text wrapping mode
                let line_height = f64::from(font_size) * 1.2;
                let cell_height = f64::from(rect.height);
                let lines = self.wrap_text_cached(value, max_width, &font_string);
                let total_text_height = lines.len() as f64 * line_height;

                // Calculate starting Y position based on vertical alignment
                let base_y = match align_v {
                    "top" => f64::from(sy) + CELL_PADDING + f64::from(font_size),
                    "bottom" => {
                        f64::from(sy) + cell_height - CELL_PADDING - total_text_height
                            + f64::from(font_size)
                    }
                    _ => {
                        // center
                        let content_start = f64::from(sy) + (cell_height - total_text_height) / 2.0;
                        content_start + f64::from(font_size)
                    }
                };

                let cell_bottom = f64::from(sy) + cell_height - CELL_PADDING;

                for (i, line) in lines.iter().enumerate() {
                    let line_y = base_y + (i as f64) * line_height;

                    // Stop rendering if we've exceeded the cell height
                    if line_y > cell_bottom {
                        break;
                    }

                    let line_width = self.measure_text_cached(line, &font_string);

                    let line_x = match align_h {
                        "center" => f64::from(sx) + (f64::from(rect.width) - line_width) / 2.0,
                        "right" => {
                            f64::from(sx) + f64::from(rect.width) - CELL_PADDING - line_width
                        }
                        _ => f64::from(sx) + CELL_PADDING + indent_pixels, // left is default
                    };

                    let _ = self.ctx.fill_text(line, line_x, line_y);

                    // Draw strikethrough for this line if enabled (DXF override takes precedence)
                    let strikethrough = dxf_override
                        .and_then(|d| d.strikethrough)
                        .unwrap_or_else(|| style.and_then(|s| s.strikethrough).unwrap_or(false));
                    if strikethrough {
                        let strike_y = line_y - f64::from(font_size) / 3.0;
                        self.ctx.begin_path();
                        self.ctx.set_stroke_style_str(&font_color);
                        self.ctx.set_line_width(1.0);
                        self.ctx.move_to(line_x, strike_y);
                        self.ctx.line_to(line_x + line_width, strike_y);
                        self.ctx.stroke();
                    }

                    // Draw underline for this line if enabled
                    if has_underline {
                        let underline_y = line_y + 2.0;
                        self.ctx.begin_path();
                        self.ctx.set_stroke_style_str(&font_color);
                        self.ctx.set_line_width(1.0);
                        self.ctx.move_to(line_x, underline_y);
                        self.ctx.line_to(line_x + line_width, underline_y);
                        self.ctx.stroke();
                    }
                }
            } else {
                // Non-wrapped text (truncate with ellipsis)
                let display_text = self.truncate_text(value, max_width, &font_string);

                // Calculate text position based on alignment
                let text_width = self.measure_text_cached(display_text.as_ref(), &font_string);

                let text_x = match align_h {
                    "center" => f64::from(sx) + (f64::from(rect.width) - text_width) / 2.0,
                    "right" => f64::from(sx) + f64::from(rect.width) - CELL_PADDING - text_width,
                    _ => f64::from(sx) + CELL_PADDING + indent_pixels, // left is default
                };

                let text_y = match align_v {
                    "top" => f64::from(sy) + f64::from(font_size) + CELL_PADDING,
                    "bottom" => f64::from(sy) + f64::from(rect.height) - CELL_PADDING,
                    _ => f64::from(sy) + f64::from(rect.height) / 2.0 + f64::from(font_size) / 3.0, // center is default
                };

                if rotation != 0 {
                    self.ctx.save();

                    // Move origin to text position
                    self.ctx.translate(text_x, text_y).ok();

                    // Calculate angle in radians
                    let angle_rad = if rotation == 255 {
                        std::f64::consts::FRAC_PI_2 // 90 degrees for vertical
                    } else if rotation <= 90 {
                        -(rotation as f64) * std::f64::consts::PI / 180.0 // counterclockwise
                    } else {
                        (rotation as f64 - 90.0) * std::f64::consts::PI / 180.0 // clockwise
                    };

                    self.ctx.rotate(angle_rad).ok();

                    // Draw at origin (since we translated)
                    self.ctx.fill_text(display_text.as_ref(), 0.0, 0.0).ok();

                    self.ctx.restore();
                } else {
                    // Normal text drawing
                    let _ = self.ctx.fill_text(display_text.as_ref(), text_x, text_y);
                }

                // Draw strikethrough if enabled (DXF override takes precedence)
                let strikethrough = dxf_override
                    .and_then(|d| d.strikethrough)
                    .unwrap_or_else(|| style.and_then(|s| s.strikethrough).unwrap_or(false));
                if strikethrough && rotation == 0 {
                    let line_y = text_y - f64::from(font_size) / 3.0; // Middle of text
                    self.ctx.begin_path();
                    self.ctx.set_stroke_style_str(&font_color);
                    self.ctx.set_line_width(1.0);
                    self.ctx.move_to(text_x, line_y);
                    self.ctx.line_to(text_x + text_width, line_y);
                    self.ctx.stroke();
                }

                // Draw underline if enabled (style underline or hyperlink)
                if has_underline && rotation == 0 {
                    let line_y = text_y + 2.0; // Just below the baseline
                    self.ctx.begin_path();
                    self.ctx.set_stroke_style_str(&font_color);
                    self.ctx.set_line_width(1.0);
                    self.ctx.move_to(text_x, line_y);
                    self.ctx.line_to(text_x + text_width, line_y);
                    self.ctx.stroke();
                }
            }

            if needs_clip {
                self.ctx.restore();
            }
        }
    }

    fn wrap_text_cached(&mut self, text: &str, max_width: f64, font: &str) -> Rc<Vec<String>> {
        if let Some(lines) = self.text_wrap_cache.get(font, max_width, text) {
            return lines;
        }
        let lines = self.wrap_text(text, max_width, font);
        self.text_wrap_cache.insert(font, max_width, text, lines)
    }

    /// Wrap text into lines that fit within max_width
    fn wrap_text(&mut self, text: &str, max_width: f64, font: &str) -> Vec<String> {
        let mut lines: Vec<String> = Vec::new();
        let mut current_line = String::new();

        for word in text.split_whitespace() {
            if current_line.is_empty() {
                // First word on the line
                let word_width = self.measure_text_cached(word, font);

                if word_width > max_width {
                    // Word is too long, need to break it
                    let broken = self.break_word(word, max_width, font);
                    let broken_len = broken.len();
                    for (i, part) in broken.into_iter().enumerate() {
                        if i == broken_len - 1 {
                            // Last part becomes the current line
                            current_line = part;
                        } else {
                            lines.push(part);
                        }
                    }
                } else {
                    current_line = word.to_string();
                }
            } else {
                // Try adding word to current line
                let test_line = format!("{} {}", current_line, word);
                let test_width = self.measure_text_cached(&test_line, font);

                if test_width <= max_width {
                    current_line = test_line;
                } else {
                    // Current line is full, start new line
                    lines.push(std::mem::take(&mut current_line));

                    let word_width = self.measure_text_cached(word, font);

                    if word_width > max_width {
                        // Word is too long, need to break it
                        let broken = self.break_word(word, max_width, font);
                        let broken_len = broken.len();
                        for (i, part) in broken.into_iter().enumerate() {
                            if i == broken_len - 1 {
                                current_line = part;
                            } else {
                                lines.push(part);
                            }
                        }
                    } else {
                        current_line = word.to_string();
                    }
                }
            }
        }

        // Don't forget the last line
        if !current_line.is_empty() {
            lines.push(current_line);
        }

        lines
    }

    /// Break a single word that's too long to fit on one line
    fn break_word(&mut self, word: &str, max_width: f64, font: &str) -> Vec<String> {
        let mut parts: Vec<String> = Vec::new();
        let chars: Vec<char> = word.chars().collect();
        let mut start = 0;

        while start < chars.len() {
            let mut end = chars.len();

            // Binary search for how many chars fit
            while end > start + 1 {
                let test: String = chars
                    .get(start..end)
                    .map(|slice| slice.iter().collect())
                    .unwrap_or_default();
                let width = self.measure_text_cached(&test, font);

                if width <= max_width {
                    break;
                }
                end -= 1;
            }

            // Ensure we take at least one character to avoid infinite loop
            if end == start {
                end = start + 1;
            }

            let part: String = chars
                .get(start..end)
                .map(|slice| slice.iter().collect())
                .unwrap_or_default();
            parts.push(part);
            start = end;
        }

        parts
    }

    /// Truncate text with ellipsis if it exceeds max width
    fn truncate_text<'a>(&mut self, text: &'a str, max_width: f64, font: &str) -> Cow<'a, str> {
        if self.measure_text_cached(text, font) <= max_width {
            return Cow::Borrowed(text);
        }

        let ellipsis = "\u{2026}"; // Unicode ellipsis
        let ellipsis_width = self.measure_text_cached(ellipsis, font);
        let available = max_width - ellipsis_width;

        if available <= 0.0 {
            return Cow::Borrowed(ellipsis);
        }

        // Binary search for the maximum text that fits
        let chars: Vec<char> = text.chars().collect();
        let mut low = 0;
        let mut high = chars.len();

        while low < high {
            let mid = (low + high).div_ceil(2);
            let truncated: String = chars.iter().take(mid).collect();
            let width = self.measure_text_cached(&truncated, font);
            if width <= available {
                low = mid;
            } else {
                high = mid - 1;
            }
        }

        let mut truncated: String = chars.iter().take(low).collect();
        truncated.push_str(ellipsis);
        Cow::Owned(truncated)
    }

    /// Render rich text with per-run formatting
    fn render_rich_text(
        &mut self,
        runs: &[crate::types::TextRunData],
        style: Option<&CellStyleData>,
        sx: f32,
        sy: f32,
        rect: &crate::layout::CellRect,
        default_font: &str,
    ) {
        // Get cell-level defaults from style
        let default_font_size = style.and_then(|s| s.font_size).unwrap_or(11.0);
        let default_font_family = style
            .and_then(|s| s.font_family.as_deref())
            .unwrap_or(default_font);
        let default_bold = style.and_then(|s| s.bold).unwrap_or(false);
        let default_italic = style.and_then(|s| s.italic).unwrap_or(false);
        let default_font_color = style
            .and_then(|s| s.font_color.as_ref())
            .and_then(|c| parse_color(c))
            .unwrap_or_else(|| palette::BLACK.to_string());
        let align_h = style.and_then(|s| s.align_h.as_deref()).unwrap_or("left");
        let align_v = style.and_then(|s| s.align_v.as_deref()).unwrap_or("center");

        // Calculate total width of all runs for alignment
        let mut total_width = 0.0;
        for run in runs {
            let font_size = run.font_size.unwrap_or(default_font_size);
            let bold = run.bold.unwrap_or(default_bold);
            let italic = run.italic.unwrap_or(default_italic);
            let font_family = run.font_family.as_deref().unwrap_or(default_font_family);

            let font_style = if italic { "italic " } else { "" };
            let font_weight = if bold { "bold " } else { "" };
            let font_string = format!(
                "{}{}{}px {}",
                font_style, font_weight, font_size, font_family
            );
            self.ctx.set_font(&font_string);
            total_width += self.measure_text_cached(&run.text, &font_string);
        }

        // Calculate starting X position based on alignment
        let text_x_start = match align_h {
            "center" => f64::from(sx) + (f64::from(rect.width) - total_width) / 2.0,
            "right" => f64::from(sx) + f64::from(rect.width) - CELL_PADDING - total_width,
            _ => f64::from(sx) + CELL_PADDING, // left is default
        };

        // Calculate Y position based on alignment (use default font size for baseline)
        let text_y = match align_v {
            "top" => f64::from(sy) + f64::from(default_font_size) + CELL_PADDING,
            "bottom" => f64::from(sy) + f64::from(rect.height) - CELL_PADDING,
            _ => f64::from(sy) + f64::from(rect.height) / 2.0 + f64::from(default_font_size) / 3.0, // center is default
        };

        // Render each run
        let mut x_offset = text_x_start;
        let cell_right = f64::from(sx) + f64::from(rect.width) - CELL_PADDING;

        for run in runs {
            // Stop if we've exceeded the cell width
            if x_offset >= cell_right {
                break;
            }

            // Get run-specific formatting, falling back to cell defaults
            let font_size = run.font_size.unwrap_or(default_font_size);
            let bold = run.bold.unwrap_or(default_bold);
            let italic = run.italic.unwrap_or(default_italic);
            let font_family = run.font_family.as_deref().unwrap_or(default_font_family);
            let font_color = run
                .font_color
                .as_ref()
                .and_then(|c| parse_color(c))
                .unwrap_or_else(|| default_font_color.clone());
            let underline = run.underline.unwrap_or(false);
            let strikethrough = run.strikethrough.unwrap_or(false);

            // Build font string for this run
            let font_style = if italic { "italic " } else { "" };
            let font_weight = if bold { "bold " } else { "" };
            let font_string = format!(
                "{}{}{}px {}",
                font_style, font_weight, font_size, font_family
            );
            self.ctx.set_font(&font_string);
            self.ctx.set_fill_style_str(&font_color);

            // Draw text
            self.ctx.fill_text(&run.text, x_offset, text_y).ok();

            // Measure text width for this run
            let run_width = self.measure_text_cached(&run.text, &font_string);

            // Draw underline if enabled
            if underline {
                let line_y = text_y + 2.0;
                self.ctx.begin_path();
                self.ctx.set_stroke_style_str(&font_color);
                self.ctx.set_line_width(1.0);
                self.ctx.move_to(x_offset, line_y);
                self.ctx.line_to(x_offset + run_width, line_y);
                self.ctx.stroke();
            }

            // Draw strikethrough if enabled
            if strikethrough {
                let line_y = text_y - f64::from(font_size) / 3.0;
                self.ctx.begin_path();
                self.ctx.set_stroke_style_str(&font_color);
                self.ctx.set_line_width(1.0);
                self.ctx.move_to(x_offset, line_y);
                self.ctx.line_to(x_offset + run_width, line_y);
                self.ctx.stroke();
            }

            // Advance X position
            x_offset += run_width;
        }
    }

    /// Render tab bar with horizontal scrolling support (Google Sheets style)
    fn render_tabs(
        &self,
        sheet_names: &[String],
        tab_colors: &[Option<String>],
        viewport: &Viewport,
        active_sheet: usize,
    ) {
        let tab_y = f64::from(viewport.height);
        let total_height = f64::from(viewport.height) + TAB_BAR_HEIGHT;
        let viewport_width = f64::from(viewport.width);
        let tab_scroll = f64::from(viewport.tab_scroll_x);

        // Tab styling constants
        const TAB_PADDING: f64 = 16.0; // Horizontal padding inside each tab
        const TAB_MIN_WIDTH: f64 = 60.0; // Minimum tab width
        const TAB_MAX_WIDTH: f64 = 200.0; // Maximum tab width
        const TAB_GAP: f64 = 1.0; // Gap between tabs
        const TAB_HEIGHT: f64 = 24.0; // Tab height
        const TAB_TOP_MARGIN: f64 = 2.0; // Space above tabs
        const SCROLL_BUTTON_WIDTH: f64 = 28.0;

        // Calculate tab widths using proper text measurement estimation
        // Each character is approximately 7px wide for the font we use
        let char_width = 7.0;
        let tab_widths: Vec<f64> = sheet_names
            .iter()
            .map(|name| {
                let text_width = name.len() as f64 * char_width;
                (text_width + TAB_PADDING * 2.0).clamp(TAB_MIN_WIDTH, TAB_MAX_WIDTH)
            })
            .collect();

        // Calculate total width of all tabs
        let total_tab_width: f64 = tab_widths.iter().sum::<f64>()
            + (tab_widths.len().saturating_sub(1)) as f64 * TAB_GAP
            + 12.0; // Extra padding at start

        // Check if we need scroll buttons
        let needs_scroll = total_tab_width > viewport_width;
        let button_width = if needs_scroll {
            SCROLL_BUTTON_WIDTH
        } else {
            0.0
        };
        let tab_area_start = button_width;
        let tab_area_end = viewport_width - button_width;

        // Tab bar background
        self.fill_rect(0.0, tab_y, viewport_width, TAB_BAR_HEIGHT, palette::TAB_BG);

        // Tab bar top border (subtle)
        self.stroke_line(
            0.0,
            tab_y + 0.5,
            viewport_width,
            tab_y + 0.5,
            1.0,
            palette::TAB_BORDER,
        );

        // Draw scroll buttons if needed
        if needs_scroll {
            let max_scroll = (total_tab_width - viewport_width + button_width * 2.0).max(0.0);

            // Left scroll button
            let left_enabled = tab_scroll > 0.0;
            self.render_tab_scroll_button(
                0.0,
                tab_y,
                button_width,
                TAB_BAR_HEIGHT,
                true, // is_left
                left_enabled,
            );

            // Right scroll button
            let right_enabled = tab_scroll < max_scroll;
            self.render_tab_scroll_button(
                viewport_width - button_width,
                tab_y,
                button_width,
                TAB_BAR_HEIGHT,
                false, // is_left
                right_enabled,
            );
        }

        // Clip tab drawing to the tab area (between scroll buttons)
        self.ctx.save();
        self.ctx.begin_path();
        self.ctx.rect(
            tab_area_start,
            tab_y,
            tab_area_end - tab_area_start,
            TAB_BAR_HEIGHT,
        );
        self.ctx.clip();

        // Draw individual tabs with scroll offset applied
        let mut x = 8.0f64 + tab_area_start - tab_scroll;
        for (i, (name, &tab_width)) in sheet_names.iter().zip(tab_widths.iter()).enumerate() {
            let is_active = i == active_sheet;
            let tab_top = tab_y + TAB_TOP_MARGIN;

            // Skip tabs that are completely outside the visible area
            if x + tab_width < tab_area_start || x > tab_area_end {
                x += tab_width + TAB_GAP;
                continue;
            }

            // Get custom tab color if present
            let custom_color = tab_colors.get(i).and_then(|c| c.as_ref());

            // Tab background color
            let bg_color = if let Some(color) = custom_color {
                if is_active {
                    Self::lighten_color(color, 0.4)
                } else {
                    Self::lighten_color(color, 0.1)
                }
            } else if is_active {
                palette::TAB_ACTIVE.to_string()
            } else {
                palette::TAB_INACTIVE.to_string()
            };

            // Draw tab background with rounded top corners
            self.fill_rounded_rect_top(x, tab_top, tab_width, TAB_HEIGHT, 4.0, &bg_color);

            // Active tab gets a colored bottom border indicator
            if is_active {
                const DEFAULT_INDICATOR: &str = "#1A73E8";
                let indicator_color = custom_color
                    .map(|s| s.as_str())
                    .unwrap_or(DEFAULT_INDICATOR);
                self.ctx.set_fill_style_str(indicator_color);
                self.ctx
                    .fill_rect(x + 1.0, total_height - 3.0, tab_width - 2.0, 3.0);
            }

            // Tab side borders (subtle)
            if !is_active {
                self.ctx.set_stroke_style_str(palette::TAB_BORDER);
                self.ctx.set_line_width(1.0);
                self.ctx.begin_path();
                self.ctx.move_to(x + tab_width - 0.5, tab_top + 4.0);
                self.ctx
                    .line_to(x + tab_width - 0.5, tab_top + TAB_HEIGHT - 4.0);
                self.ctx.stroke();
            }

            // Tab text - use system font stack for better rendering
            let font_weight = if is_active { "500" } else { "400" };
            self.ctx.set_font(&format!(
                "{} 12px -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif",
                font_weight
            ));
            self.ctx.set_text_align("center");
            self.ctx.set_text_baseline("middle");

            let text_color = if let Some(color) = custom_color {
                if Self::is_light_color(color) {
                    palette::TAB_TEXT_ACTIVE
                } else {
                    palette::WHITE
                }
            } else if is_active {
                palette::TAB_TEXT_ACTIVE
            } else {
                palette::TAB_TEXT
            };
            self.ctx.set_fill_style_str(text_color);

            // Truncate text if needed
            #[allow(clippy::cast_possible_truncation, clippy::cast_sign_loss)]
            let max_chars = ((tab_width - TAB_PADDING) / char_width).max(0.0) as usize;
            let display_name = if name.len() > max_chars && max_chars > 3 {
                format!("{}...", &name[..max_chars - 3])
            } else {
                name.clone()
            };
            let _ = self.ctx.fill_text(
                &display_name,
                x + tab_width / 2.0,
                tab_top + TAB_HEIGHT / 2.0,
            );

            x += tab_width + TAB_GAP;
        }

        self.ctx.restore();
    }

    /// Render a scroll button for the tab bar
    fn render_tab_scroll_button(
        &self,
        x: f64,
        y: f64,
        width: f64,
        height: f64,
        is_left: bool,
        enabled: bool,
    ) {
        // Button background
        let bg = if enabled {
            palette::TAB_SCROLL_BG
        } else {
            palette::TAB_BG
        };
        self.fill_rect(x, y, width, height, bg);

        // Draw arrow icon
        let icon_color = if enabled {
            palette::TAB_SCROLL_ICON
        } else {
            "#BDC1C6" // Disabled color
        };
        self.ctx.set_fill_style_str(icon_color);
        self.ctx.set_font("12px sans-serif");
        self.ctx.set_text_align("center");
        self.ctx.set_text_baseline("middle");

        let arrow = if is_left { "◂" } else { "▸" };
        let _ = self.ctx.fill_text(arrow, x + width / 2.0, y + height / 2.0);

        // Right border for left button, left border for right button
        self.ctx.set_stroke_style_str(palette::TAB_BORDER);
        self.ctx.set_line_width(1.0);
        self.ctx.begin_path();
        if is_left {
            self.ctx.move_to(x + width - 0.5, y);
            self.ctx.line_to(x + width - 0.5, y + height);
        } else {
            self.ctx.move_to(x + 0.5, y);
            self.ctx.line_to(x + 0.5, y + height);
        }
        self.ctx.stroke();
    }

    /// Fill a rectangle with only the top corners rounded
    fn fill_rounded_rect_top(&self, x: f64, y: f64, w: f64, h: f64, r: f64, color: &str) {
        self.ctx.set_fill_style_str(color);
        self.ctx.begin_path();
        self.ctx.move_to(x + r, y);
        self.ctx.line_to(x + w - r, y);
        self.ctx.quadratic_curve_to(x + w, y, x + w, y + r);
        self.ctx.line_to(x + w, y + h);
        self.ctx.line_to(x, y + h);
        self.ctx.line_to(x, y + r);
        self.ctx.quadratic_curve_to(x, y, x + r, y);
        self.ctx.close_path();
        self.ctx.fill();
    }

    /// Lighten a hex color by a factor (0.0 to 1.0)
    fn lighten_color(hex: &str, factor: f64) -> String {
        Rgb::from_hex(hex)
            .map(|c| c.lighten(factor).to_hex())
            .unwrap_or_else(|| format!("#{}", hex.trim_start_matches('#')))
    }

    /// Check if a color is light (high luminance)
    fn is_light_color(hex: &str) -> bool {
        Rgb::from_hex(hex).is_none_or(|c| c.is_light())
    }

    /// Render selection highlight over selected cells
    ///
    /// Selection is specified as (start_row, start_col, end_row, end_col).
    /// Draws a semi-transparent blue fill over the selection and a solid border around it.
    fn render_selection(
        &self,
        selection: (u32, u32, u32, u32),
        layout: &SheetLayout,
        viewport: &Viewport,
    ) {
        let rects = selection_rects(selection, layout, viewport);
        if rects.is_empty() {
            return;
        }

        let content_width = f64::from(viewport.width) - SCROLLBAR_SIZE;
        let content_height = f64::from(viewport.height) - SCROLLBAR_SIZE;

        // Draw semi-transparent blue fill: rgba(0, 120, 215, 0.2)
        self.ctx.set_fill_style_str("rgba(0, 120, 215, 0.2)");

        for rect in &rects {
            let x = rect.x;
            let y = rect.y;
            let w = rect.w;
            let h = rect.h;

            if x >= content_width || y >= content_height || x + w <= 0.0 || y + h <= 0.0 {
                continue;
            }

            let clip_x = x.max(0.0);
            let clip_y = y.max(0.0);
            let clip_w = (x + w).min(content_width) - clip_x;
            let clip_h = (y + h).min(content_height) - clip_y;
            if clip_w > 0.0 && clip_h > 0.0 {
                self.ctx.fill_rect(clip_x, clip_y, clip_w, clip_h);
            }
        }

        self.ctx.set_stroke_style_str("#0078D7");
        self.ctx.set_line_width(2.0);
        self.ctx.begin_path();

        for rect in &rects {
            let x = rect.x;
            let y = rect.y;
            let w = rect.w;
            let h = rect.h;

            if x >= content_width || y >= content_height || x + w <= 0.0 || y + h <= 0.0 {
                continue;
            }

            let left_x = x.max(0.0);
            let right_x = (x + w).min(content_width);
            let top_y = y.max(0.0);
            let bottom_y = (y + h).min(content_height);

            if rect.draw_top && y >= 0.0 && y <= content_height && right_x > left_x {
                self.ctx.move_to(left_x, y);
                self.ctx.line_to(right_x, y);
            }
            if rect.draw_bottom && y + h >= 0.0 && y + h <= content_height && right_x > left_x {
                self.ctx.move_to(left_x, y + h);
                self.ctx.line_to(right_x, y + h);
            }
            if rect.draw_left && x >= 0.0 && x <= content_width && bottom_y > top_y {
                self.ctx.move_to(x, top_y);
                self.ctx.line_to(x, bottom_y);
            }
            if rect.draw_right && x + w >= 0.0 && x + w <= content_width && bottom_y > top_y {
                self.ctx.move_to(x + w, top_y);
                self.ctx.line_to(x + w, bottom_y);
            }
        }

        self.ctx.stroke();

        // Draw resize handle (small square) in bottom-right corner
        let handle_rect = rects.iter().find(|r| r.draw_bottom && r.draw_right);
        if let Some(rect) = handle_rect {
            let handle_size = 8.0;
            let handle_x = rect.x + rect.w - handle_size / 2.0;
            let handle_y = rect.y + rect.h - handle_size / 2.0;

            if handle_x + handle_size > 0.0
                && handle_x < content_width
                && handle_y + handle_size > 0.0
                && handle_y < content_height
            {
                let hx = handle_x.max(0.0).min(content_width - handle_size);
                let hy = handle_y.max(0.0).min(content_height - handle_size);

                if (handle_x - hx).abs() < 1.0 && (handle_y - hy).abs() < 1.0 {
                    self.ctx.set_fill_style_str("#0078D7");
                    self.ctx.fill_rect(hx, hy, handle_size, handle_size);

                    self.ctx.set_fill_style_str("#FFFFFF");
                    self.ctx
                        .fill_rect(hx + 1.5, hy + 1.5, handle_size - 3.0, handle_size - 3.0);
                }
            }
        }
    }

    /// Render embedded images (drawings with type "picture")
    fn render_images(
        &mut self,
        drawings: &[Drawing],
        images: &[EmbeddedImage],
        layout: &SheetLayout,
        viewport: &Viewport,
    ) {
        // Build a lookup map for images by their ID
        let image_map: HashMap<&str, &EmbeddedImage> =
            images.iter().map(|img| (img.id.as_str(), img)).collect();

        for drawing in drawings {
            // Only render pictures (images), skip charts and shapes for now
            if drawing.drawing_type != "picture" {
                continue;
            }

            // Get the image_id to look up the embedded image data
            let Some(ref image_id) = drawing.image_id else {
                continue;
            };

            // Calculate the image position and size based on anchor type
            let (x, y, width, height) = match drawing.anchor_type.as_str() {
                "twoCellAnchor" => {
                    // Position determined by from/to cell anchors
                    self.calculate_two_cell_anchor_bounds(drawing, layout)
                }
                "oneCellAnchor" => {
                    // Position determined by from cell + extent size
                    self.calculate_one_cell_anchor_bounds(drawing, layout)
                }
                "absoluteAnchor" => {
                    // Absolute position in EMUs
                    self.calculate_absolute_anchor_bounds(drawing)
                }
                _ => continue,
            };

            // Skip images with zero or negative size
            if width <= 0.0 || height <= 0.0 {
                continue;
            }

            // Convert to screen coordinates
            // For images, we use the from_col/from_row to determine frozen behavior
            let from_col = drawing.from_col.unwrap_or(0);
            let from_row = drawing.from_row.unwrap_or(0);
            let (screen_x, screen_y) = viewport.to_screen_frozen(x, y, from_row, from_col, layout);
            let screen_width = width * viewport.scale;
            let screen_height = height * viewport.scale;

            // Get or create the HtmlImageElement for this image
            let img_element = self.get_or_create_image(image_id, &image_map);

            // Draw the image if it's loaded, otherwise show a loading spinner
            if let Some(ref img) = img_element {
                if img.natural_width() > 0 {
                    let _ = self.ctx.draw_image_with_html_image_element_and_dw_and_dh(
                        img,
                        f64::from(screen_x),
                        f64::from(screen_y),
                        f64::from(screen_width),
                        f64::from(screen_height),
                    );
                } else {
                    self.draw_loading_spinner(screen_x, screen_y, screen_width, screen_height);
                }
            }
        }
    }

    /// Draw a small loading spinner centered in the given rectangle.
    #[allow(clippy::cast_possible_truncation)]
    fn draw_loading_spinner(&self, x: f32, y: f32, w: f32, h: f32) {
        let cx = f64::from(x + w * 0.5);
        let cy = f64::from(y + h * 0.5);
        let radius = f64::from(w.min(h) * 0.15).clamp(4.0, 16.0);
        let angle = (self.frame_id % 60) as f64 * std::f64::consts::TAU / 60.0;

        self.ctx.save();
        self.ctx.begin_path();
        let _ = self
            .ctx
            .arc(cx, cy, radius, angle, angle + std::f64::consts::TAU * 0.75);
        self.ctx.set_line_width(2.0);
        self.ctx.set_stroke_style_str("#999");
        self.ctx.stroke();
        self.ctx.restore();
    }

    /// Calculate bounds for twoCellAnchor (from cell to cell)
    #[allow(clippy::cast_possible_truncation)]
    pub(super) fn calculate_two_cell_anchor_bounds(
        &self,
        drawing: &Drawing,
        layout: &SheetLayout,
    ) -> (f32, f32, f32, f32) {
        let from_col = drawing.from_col.unwrap_or(0);
        let from_row = drawing.from_row.unwrap_or(0);
        let to_col = drawing.to_col.unwrap_or(from_col);
        let to_row = drawing.to_row.unwrap_or(from_row);

        // Get cell offsets (in EMUs -> pixels)
        let from_col_off = drawing.from_col_off.unwrap_or(0) as f64 / EMU_PER_PIXEL;
        let from_row_off = drawing.from_row_off.unwrap_or(0) as f64 / EMU_PER_PIXEL;
        let to_col_off = drawing.to_col_off.unwrap_or(0) as f64 / EMU_PER_PIXEL;
        let to_row_off = drawing.to_row_off.unwrap_or(0) as f64 / EMU_PER_PIXEL;

        // Calculate full anchor bounds (the area the shape should fit within)
        let anchor_x1 = layout
            .col_positions
            .get(from_col as usize)
            .copied()
            .unwrap_or(0.0) as f64
            + from_col_off;
        let anchor_y1 = layout
            .row_positions
            .get(from_row as usize)
            .copied()
            .unwrap_or(0.0) as f64
            + from_row_off;
        let anchor_x2 = layout
            .col_positions
            .get(to_col as usize)
            .copied()
            .unwrap_or(anchor_x1 as f32) as f64
            + to_col_off;
        let anchor_y2 = layout
            .row_positions
            .get(to_row as usize)
            .copied()
            .unwrap_or(anchor_y1 as f32) as f64
            + to_row_off;

        let anchor_width = anchor_x2 - anchor_x1;
        let anchor_height = anchor_y2 - anchor_y1;

        // When xfrm values are available, use them for precise sizing and center
        // the shape within the anchor bounds. Otherwise, use anchor bounds directly.
        if let (Some(cx), Some(cy)) = (drawing.xfrm_cx, drawing.xfrm_cy) {
            let width = cx as f64 / EMU_PER_PIXEL;
            let height = cy as f64 / EMU_PER_PIXEL;

            // Center the xfrm-sized shape within the anchor bounds
            let x = anchor_x1 + (anchor_width - width) / 2.0;
            let y = anchor_y1 + (anchor_height - height) / 2.0;

            (x as f32, y as f32, width as f32, height as f32)
        } else {
            // No xfrm - use anchor bounds directly
            (
                anchor_x1 as f32,
                anchor_y1 as f32,
                anchor_width as f32,
                anchor_height as f32,
            )
        }
    }

    /// Calculate bounds for oneCellAnchor (from cell + extent size)
    #[allow(clippy::cast_possible_truncation)]
    pub(super) fn calculate_one_cell_anchor_bounds(
        &self,
        drawing: &Drawing,
        layout: &SheetLayout,
    ) -> (f32, f32, f32, f32) {
        // Note: xfrm values contain Excel's pre-calculated absolute positions, but they
        // require our layout to match Excel's pixel-perfectly. Since column widths may
        // differ, we use the cell-anchor based approach which is relative to cell positions.
        let from_col = drawing.from_col.unwrap_or(0);
        let from_row = drawing.from_row.unwrap_or(0);

        // Get column/row offsets (in EMUs)
        let from_col_off = drawing.from_col_off.unwrap_or(0) as f64 / EMU_PER_PIXEL;
        let from_row_off = drawing.from_row_off.unwrap_or(0) as f64 / EMU_PER_PIXEL;

        // Get extent size (in EMUs)
        let width = drawing.extent_cx.unwrap_or(0) as f64 / EMU_PER_PIXEL;
        let height = drawing.extent_cy.unwrap_or(0) as f64 / EMU_PER_PIXEL;

        // Get cell position from layout
        let x = layout
            .col_positions
            .get(from_col as usize)
            .copied()
            .unwrap_or(0.0) as f64
            + from_col_off;
        let y = layout
            .row_positions
            .get(from_row as usize)
            .copied()
            .unwrap_or(0.0) as f64
            + from_row_off;

        (x as f32, y as f32, width as f32, height as f32)
    }

    /// Calculate bounds for absoluteAnchor (absolute position in EMUs)
    #[allow(clippy::cast_possible_truncation)]
    pub(super) fn calculate_absolute_anchor_bounds(
        &self,
        drawing: &Drawing,
    ) -> (f32, f32, f32, f32) {
        let x = drawing.pos_x.unwrap_or(0) as f64 / EMU_PER_PIXEL;
        let y = drawing.pos_y.unwrap_or(0) as f64 / EMU_PER_PIXEL;
        let width = drawing.extent_cx.unwrap_or(0) as f64 / EMU_PER_PIXEL;
        let height = drawing.extent_cy.unwrap_or(0) as f64 / EMU_PER_PIXEL;

        (x as f32, y as f32, width as f32, height as f32)
    }

    /// Get or create an HtmlImageElement for the given image ID
    fn get_or_create_image(
        &mut self,
        image_id: &str,
        image_map: &HashMap<&str, &EmbeddedImage>,
    ) -> Option<HtmlImageElement> {
        // Check cache first
        if let Some(img) = self.image_cache.get(image_id) {
            return Some(img.clone());
        }

        // Look up the embedded image data
        // The image_id in Drawing has been resolved during parsing to the actual
        // image path like "xl/media/image1.png", which matches EmbeddedImage.id
        let embedded = image_map.get(image_id).copied().or_else(|| {
            // Fallback: try to find by matching the filename
            image_map
                .values()
                .find(|img| {
                    img.id.ends_with(image_id)
                        || img.filename.as_ref().is_some_and(|f| f == image_id)
                })
                .copied()
        })?;

        // Create HtmlImageElement and set src
        let document = self.get_document()?;
        let img = document
            .create_element("img")
            .ok()?
            .dyn_into::<HtmlImageElement>()
            .ok()?;

        // Set the src as a data URL
        let data_url = format!("data:{};base64,{}", embedded.mime_type, embedded.data);
        img.set_src(&data_url);

        // Cache the image element
        self.image_cache.insert(image_id.to_string(), img.clone());

        Some(img)
    }
}

impl CanvasRenderer {
    pub fn render_base(&mut self, params: &RenderParams<'_>) -> Result<()> {
        self.render_internal(params, false)
    }

    pub fn render_overlay(&mut self, params: &RenderParams<'_>) -> Result<()> {
        self.ctx
            .reset_transform()
            .map_err(|_| "Failed to reset transform")?;
        self.ctx
            .clear_rect(0.0, 0.0, f64::from(self.width), f64::from(self.height));

        // Save clean state so any clips set during rendering are fully
        // cleaned up when we restore at the end.  Canvas 2D clips survive
        // reset_transform(), so without this wrapper a clip leak in any
        // sub-call would corrupt subsequent frames.
        self.ctx.save();

        let _ = self.ctx.scale(f64::from(self.dpr), f64::from(self.dpr));

        // Calculate header offset (same as render_internal)
        let header_offset_x = if params.show_headers {
            f64::from(params.header_config.row_header_width)
        } else {
            0.0
        };
        let header_offset_y = if params.show_headers {
            f64::from(params.header_config.col_header_height)
        } else {
            0.0
        };

        // Apply header offset for selection/comments (which use cell coordinates)
        if params.show_headers {
            let _ = self.ctx.translate(header_offset_x, header_offset_y);
        }

        self.render_overlay_layer(params, header_offset_x, header_offset_y);

        // Restore to clean state (removes any clips from this frame).
        self.ctx.restore();
        Ok(())
    }

    #[allow(clippy::cast_possible_truncation)]
    fn render_overlay_layer(
        &mut self,
        params: &RenderParams<'_>,
        header_offset_x: f64,
        header_offset_y: f64,
    ) {
        let layout = params.layout;
        let viewport = params.viewport;
        let (_, _, content_width, content_height) =
            self.content_bounds(viewport, header_offset_x, header_offset_y);

        // Overlay region caching is disabled — always render fresh.
        // The bitmap cache approach had coordinate/invalidation bugs causing
        // frozen-row column desync.  Rendering fresh each frame is correct
        // and fast enough (~200 canvas calls, <2ms).

        // --- Frozen panes ---
        let has_frozen = layout.frozen_rows > 0 || layout.frozen_cols > 0;
        if has_frozen {
            let frozen_width = f64::from(layout.frozen_cols_width());
            let frozen_height = f64::from(layout.frozen_rows_height());
            let data_width = f64::from(layout.total_width()).min(content_width);
            let data_height = f64::from(layout.total_height()).min(content_height);

            let (start_row, end_row) =
                viewport.visible_rows_in_height(layout, content_height as f32);
            let (start_col, end_col) = viewport.visible_cols_in_width(layout, content_width as f32);

            let default_font = if let Some(minor) = params.minor_font {
                format!("{}, Arial, sans-serif", minor)
            } else {
                "Calibri, Arial, sans-serif".to_string()
            };

            let frozen_cells: Vec<&CellRenderData> = params
                .cells
                .iter()
                .filter(|c| c.row < layout.frozen_rows || c.col < layout.frozen_cols)
                .collect();

            // === PASS 2: Frozen rows ===
            if layout.frozen_rows > 0 {
                self.ctx.save();
                self.ctx.begin_path();
                let frozen_row_width = (data_width - frozen_width + 1.0).max(0.0);
                self.ctx
                    .rect(frozen_width, 0.0, frozen_row_width, frozen_height);
                self.ctx.clip();

                self.fill_rect(
                    frozen_width,
                    0.0,
                    (data_width - frozen_width).max(0.0),
                    frozen_height,
                    palette::WHITE,
                );

                let frozen_row_cells: Vec<&CellRenderData> = frozen_cells
                    .iter()
                    .filter(|c| c.row < layout.frozen_rows && c.col >= layout.frozen_cols)
                    .copied()
                    .collect();

                self.render_conditional_formatting(
                    &frozen_row_cells,
                    params.conditional_formatting,
                    params.dxf_styles,
                    layout,
                    viewport,
                    params.conditional_formatting_cache,
                );
                let frozen_row_dxf_overrides = Self::collect_cell_is_dxf_overrides(
                    &frozen_row_cells,
                    params.conditional_formatting,
                    params.dxf_styles,
                    params.conditional_formatting_cache,
                );
                self.render_cell_backgrounds(
                    &frozen_row_cells,
                    layout,
                    viewport,
                    params.style_cache,
                    params.default_style,
                );
                self.render_grid_lines_for_frozen_rows(layout, viewport, start_col, end_col);
                self.render_cell_borders(
                    &frozen_row_cells,
                    layout,
                    viewport,
                    params.style_cache,
                    params.default_style,
                );
                self.render_cell_text(
                    &frozen_row_cells,
                    layout,
                    viewport,
                    &default_font,
                    &frozen_row_dxf_overrides,
                    params.style_cache,
                    params.default_style,
                );

                self.ctx.restore();
            }

            // === PASS 3: Frozen columns ===
            if layout.frozen_cols > 0 {
                self.ctx.save();
                self.ctx.begin_path();
                let frozen_col_height = (data_height - frozen_height + 1.0).max(0.0);
                self.ctx
                    .rect(0.0, frozen_height, frozen_width, frozen_col_height);
                self.ctx.clip();

                self.fill_rect(
                    0.0,
                    frozen_height,
                    frozen_width,
                    (data_height - frozen_height).max(0.0),
                    palette::WHITE,
                );

                let frozen_col_cells: Vec<&CellRenderData> = frozen_cells
                    .iter()
                    .filter(|c| c.col < layout.frozen_cols && c.row >= layout.frozen_rows)
                    .copied()
                    .collect();

                self.render_conditional_formatting(
                    &frozen_col_cells,
                    params.conditional_formatting,
                    params.dxf_styles,
                    layout,
                    viewport,
                    params.conditional_formatting_cache,
                );
                let frozen_col_dxf_overrides = Self::collect_cell_is_dxf_overrides(
                    &frozen_col_cells,
                    params.conditional_formatting,
                    params.dxf_styles,
                    params.conditional_formatting_cache,
                );
                self.render_cell_backgrounds(
                    &frozen_col_cells,
                    layout,
                    viewport,
                    params.style_cache,
                    params.default_style,
                );
                self.render_grid_lines_for_frozen_cols(layout, viewport, start_row, end_row);
                self.render_cell_borders(
                    &frozen_col_cells,
                    layout,
                    viewport,
                    params.style_cache,
                    params.default_style,
                );
                self.render_cell_text(
                    &frozen_col_cells,
                    layout,
                    viewport,
                    &default_font,
                    &frozen_col_dxf_overrides,
                    params.style_cache,
                    params.default_style,
                );

                self.ctx.restore();
            }

            // === PASS 4: Corner ===
            if layout.frozen_rows > 0 && layout.frozen_cols > 0 {
                self.ctx.save();
                self.ctx.begin_path();
                self.ctx.rect(0.0, 0.0, frozen_width, frozen_height);
                self.ctx.clip();

                self.fill_rect(0.0, 0.0, frozen_width, frozen_height, palette::WHITE);

                let corner_cells: Vec<&CellRenderData> = frozen_cells
                    .iter()
                    .filter(|c| c.row < layout.frozen_rows && c.col < layout.frozen_cols)
                    .copied()
                    .collect();

                self.render_conditional_formatting(
                    &corner_cells,
                    params.conditional_formatting,
                    params.dxf_styles,
                    layout,
                    viewport,
                    params.conditional_formatting_cache,
                );
                let corner_dxf_overrides = Self::collect_cell_is_dxf_overrides(
                    &corner_cells,
                    params.conditional_formatting,
                    params.dxf_styles,
                    params.conditional_formatting_cache,
                );
                self.render_cell_backgrounds(
                    &corner_cells,
                    layout,
                    viewport,
                    params.style_cache,
                    params.default_style,
                );
                self.render_grid_lines_for_corner(layout, viewport);
                self.render_cell_borders(
                    &corner_cells,
                    layout,
                    viewport,
                    params.style_cache,
                    params.default_style,
                );
                self.render_cell_text(
                    &corner_cells,
                    layout,
                    viewport,
                    &default_font,
                    &corner_dxf_overrides,
                    params.style_cache,
                    params.default_style,
                );

                self.ctx.restore();
            }

            // === Dividers — always re-render (just 2 thin lines, negligible cost) ===
            // Not cached because capture_overlay_region would need the entire content
            // area, and blitting that back would overwrite freshly rendered frozen panes.
            render_frozen_dividers(&self.ctx, layout, viewport, self.dpr);
        }

        // Selection — always re-render (changes frequently)
        if let Some(selection) = params.selection {
            self.render_selection(selection, layout, viewport);
        }

        // Comment indicators — always re-render (cheap)
        render_comment_indicators(&self.ctx, params.cells, layout, viewport, self.dpr);

        // Validation indicators — always re-render (cheap)
        #[allow(clippy::cast_possible_truncation)]
        render_validation_indicators(
            &self.ctx,
            params.data_validations,
            layout,
            viewport,
            content_width as f32,
            content_height as f32,
            self.dpr,
        );

        // Filter buttons — always re-render (cheap)
        #[allow(clippy::cast_possible_truncation)]
        render_filter_buttons(
            &self.ctx,
            params.auto_filter,
            layout,
            viewport,
            content_width as f32,
            self.dpr,
        );

        // Render dividers for non-frozen case is handled above; for no frozen panes
        // render_frozen_dividers is a no-op (checks frozen_rows/cols == 0).

        // Reset header offset translation before rendering fixed-position elements
        // Headers, scrollbars, and tabs should render at fixed canvas positions
        if params.show_headers {
            let _ = self.ctx.translate(-header_offset_x, -header_offset_y);
        }

        // Row/column headers
        if params.show_headers {
            render_column_headers(
                &self.ctx,
                layout,
                viewport,
                params.header_config,
                params.header_selection,
                content_width,
            );

            render_row_headers(
                &self.ctx,
                layout,
                viewport,
                params.header_config,
                params.header_selection,
                content_height,
            );

            render_header_corner(
                &self.ctx,
                params.header_config,
                params
                    .header_selection
                    .map(|s| s.selection_type == crate::types::SelectionType::All)
                    .unwrap_or(false),
            );
        }

        // Tab bar — always re-render (only if using canvas-based tabs, not DOM tabs)
        if params.show_tab_bar {
            self.render_tabs(
                params.sheet_names,
                params.tab_colors,
                viewport,
                params.active_sheet,
            );
        }
    }

    fn render_internal(&mut self, params: &RenderParams<'_>, draw_overlay: bool) -> Result<()> {
        let viewport = params.viewport;
        let layout = params.layout;
        let layout_ptr = layout as *const SheetLayout as usize;
        if self.tile_cache_sheet != params.active_sheet || self.tile_cache_layout_ptr != layout_ptr
        {
            self.tile_cache.clear();
            self.tile_cache_sheet = params.active_sheet;
            self.tile_cache_layout_ptr = layout_ptr;
        }

        // Build default font string from theme fonts
        // Use theme minor_font (body text) as the primary default, with web-safe fallbacks
        let default_font = if let Some(minor) = params.minor_font {
            format!("{}, Arial, sans-serif", minor)
        } else {
            "Calibri, Arial, sans-serif".to_string()
        };

        // Header offset for cell content when headers are visible
        let header_offset_x = if params.show_headers {
            f64::from(params.header_config.row_header_width)
        } else {
            0.0
        };
        let header_offset_y = if params.show_headers {
            f64::from(params.header_config.col_header_height)
        } else {
            0.0
        };

        // Canvas dimensions (in logical pixels)
        // Tab bar is now DOM-based, so viewport.height is the full canvas height
        let total_height = f64::from(viewport.height);
        let total_width = f64::from(viewport.width);

        let scroll_region = scrollable_region(
            layout,
            viewport,
            header_offset_x,
            header_offset_y,
            SCROLLBAR_SIZE,
        );
        let content_width = scroll_region.content_width;
        let content_height = scroll_region.content_height;
        self.frame_id = self.frame_id.wrapping_add(1);
        self.deferred_prefetch_tiles = false;

        // Reset transform and save clean state.  Canvas 2D clips survive
        // reset_transform(), so save/restore ensures clip leaks from prior
        // frames are cleaned up at the end.
        self.ctx
            .reset_transform()
            .map_err(|_| "Failed to reset transform")?;
        self.ctx.save();

        let _ = self.ctx.scale(f64::from(self.dpr), f64::from(self.dpr));

        // 1. Clear canvas with white background
        self.fill_rect(0.0, 0.0, total_width, total_height, palette::WHITE);

        // Apply header offset for cell rendering
        if params.show_headers {
            let _ = self.ctx.translate(header_offset_x, header_offset_y);
        }

        // 2. Get visible range (use content area to avoid header/scrollbar mismatch)
        #[allow(clippy::cast_possible_truncation)]
        let (start_row, end_row) = viewport.visible_rows_in_height(layout, content_height as f32);
        #[allow(clippy::cast_possible_truncation)]
        let (start_col, end_col) = viewport.visible_cols_in_width(layout, content_width as f32);

        // Get frozen dimensions
        let frozen_width = f64::from(layout.frozen_cols_width());
        let frozen_height = f64::from(layout.frozen_rows_height());
        let has_frozen = layout.frozen_rows > 0 || layout.frozen_cols > 0;
        let mut all_cells: Vec<&CellRenderData> = Vec::with_capacity(params.cells.len());
        all_cells.extend(params.cells.iter());

        if has_frozen {
            // Only render scrollable cells in PASS 1 — frozen panes (PASS 2-4) are
            // always handled by render_overlay_layer so they stay fixed on screen
            // when the buffer canvas scrolls via CSS left/top.
            let scrollable_cells: Vec<&CellRenderData> = params
                .cells
                .iter()
                .filter(|c| c.row >= layout.frozen_rows && c.col >= layout.frozen_cols)
                .collect();

            // === PASS 1: Render scrollable content with clipping ===
            let scrollable_result = self.render_scrollable_tiles(
                params,
                &scrollable_cells,
                layout,
                viewport,
                content_width,
                content_height,
                &default_font,
            );
            self.deferred_prefetch_tiles = scrollable_result.deferred_prefetch;
            if !scrollable_result.rendered_visible {
                self.ctx.save();
                self.ctx.begin_path();
                self.ctx.rect(
                    frozen_width,
                    frozen_height,
                    content_width - frozen_width,
                    content_height - frozen_height,
                );
                self.ctx.clip();

                self.render_conditional_formatting(
                    &scrollable_cells,
                    params.conditional_formatting,
                    params.dxf_styles,
                    layout,
                    viewport,
                    params.conditional_formatting_cache,
                );
                let scrollable_dxf_overrides = Self::collect_cell_is_dxf_overrides(
                    &scrollable_cells,
                    params.conditional_formatting,
                    params.dxf_styles,
                    params.conditional_formatting_cache,
                );
                self.render_cell_backgrounds(
                    &scrollable_cells,
                    layout,
                    viewport,
                    params.style_cache,
                    params.default_style,
                );
                self.render_grid_lines(layout, viewport, start_row, end_row, start_col, end_col);
                self.render_cell_borders(
                    &scrollable_cells,
                    layout,
                    viewport,
                    params.style_cache,
                    params.default_style,
                );
                self.render_cell_text(
                    &scrollable_cells,
                    layout,
                    viewport,
                    &default_font,
                    &scrollable_dxf_overrides,
                    params.style_cache,
                    params.default_style,
                );

                self.ctx.restore();
            }
        } else {
            // No frozen panes - render everything normally
            let scrollable_result = self.render_scrollable_tiles(
                params,
                &all_cells,
                layout,
                viewport,
                content_width,
                content_height,
                &default_font,
            );
            self.deferred_prefetch_tiles = scrollable_result.deferred_prefetch;
            if !scrollable_result.rendered_visible {
                self.render_conditional_formatting(
                    &all_cells,
                    params.conditional_formatting,
                    params.dxf_styles,
                    layout,
                    viewport,
                    params.conditional_formatting_cache,
                );
                let dxf_overrides = Self::collect_cell_is_dxf_overrides(
                    &all_cells,
                    params.conditional_formatting,
                    params.dxf_styles,
                    params.conditional_formatting_cache,
                );
                self.render_cell_backgrounds(
                    &all_cells,
                    layout,
                    viewport,
                    params.style_cache,
                    params.default_style,
                );
                self.render_grid_lines(layout, viewport, start_row, end_row, start_col, end_col);
                self.render_cell_borders(
                    &all_cells,
                    layout,
                    viewport,
                    params.style_cache,
                    params.default_style,
                );
                self.render_cell_text(
                    &all_cells,
                    layout,
                    viewport,
                    &default_font,
                    &dxf_overrides,
                    params.style_cache,
                    params.default_style,
                );
            }
        }

        // 7. Render embedded images
        self.render_images(params.drawings, params.images, layout, viewport);

        // 8. Render shapes
        self.render_shapes(params.drawings, layout, viewport);

        // 9. Render charts
        self.render_charts(params.charts, layout, viewport);

        // 10. Render sparklines
        self.render_sparklines(params.sparkline_groups, &all_cells, layout, viewport);

        if draw_overlay {
            self.render_overlay_layer(params, header_offset_x, header_offset_y);
        }

        // Restore clean state (removes any clip leaks from this frame).
        self.ctx.restore();
        Ok(())
    }
}

impl RenderBackend for CanvasRenderer {
    fn init(&mut self) -> Result<()> {
        // Canvas 2D doesn't need explicit initialization
        Ok(())
    }

    fn resize(&mut self, width: u32, height: u32, dpr: f32) {
        self.width = width;
        self.height = height;
        self.dpr = dpr;
        self.text_measure_cache.clear();
        self.text_wrap_cache.clear();
        self.tile_cache.clear();
        self.tile_cache.dpr = dpr;

        // Set canvas buffer size to physical pixels
        self.canvas.set_width(width);
        self.canvas.set_height(height);

        // Scale context for DPR (all drawing uses logical coordinates after this)
        let _ = self.ctx.scale(f64::from(dpr), f64::from(dpr));
    }

    fn render(&mut self, params: &RenderParams<'_>) -> Result<()> {
        self.render_internal(params, true)
    }

    fn width(&self) -> u32 {
        self.width
    }

    fn height(&self) -> u32 {
        self.height
    }
}

#[cfg(all(feature = "editing", target_arch = "wasm32"))]
impl CanvasRenderer {
    /// Clear the tile cache (call after data mutations so stale tiles are discarded).
    pub(crate) fn clear_tile_cache(&mut self) {
        self.tile_cache.clear();
    }
}

#[cfg(test)]
#[allow(
    clippy::unwrap_used,
    clippy::expect_used,
    clippy::indexing_slicing,
    clippy::float_cmp,
    clippy::panic
)]
mod tests {
    use super::{TextMeasureCache, TextWrapCache};
    use std::rc::Rc;

    #[test]
    fn text_measure_cache_reuses_entries() {
        let mut cache = TextMeasureCache::new(2);
        assert_eq!(cache.get("11px Arial", "hello"), None);
        cache.insert("11px Arial", "hello", 12.0);
        assert_eq!(cache.get("11px Arial", "hello"), Some(12.0));
        cache.insert("11px Arial", "hello", 22.0);
        assert_eq!(cache.get("11px Arial", "hello"), Some(12.0));
        assert_eq!(cache.len(), 1);
    }

    #[test]
    fn text_measure_cache_enforces_cap() {
        let mut cache = TextMeasureCache::new(2);
        cache.insert("11px Arial", "a", 1.0);
        cache.insert("11px Arial", "b", 2.0);
        cache.insert("11px Arial", "c", 3.0);
        assert_eq!(cache.len(), 2);
        assert_eq!(cache.get("11px Arial", "a"), None);
        assert_eq!(cache.get("11px Arial", "b"), Some(2.0));
        assert_eq!(cache.get("11px Arial", "c"), Some(3.0));
    }

    #[test]
    fn text_wrap_cache_reuses_entries() {
        let mut cache = TextWrapCache::new(2);
        assert!(cache.get("11px Arial", 120.0, "hello").is_none());
        let first = cache.insert("11px Arial", 120.0, "hello", vec!["hello".to_string()]);
        let second = cache
            .get("11px Arial", 120.0, "hello")
            .expect("cache should contain entry");
        assert!(Rc::ptr_eq(&first, &second));
        assert_eq!(cache.len(), 1);
    }

    #[test]
    fn text_wrap_cache_enforces_cap() {
        let mut cache = TextWrapCache::new(2);
        cache.insert("11px Arial", 120.0, "a", vec!["a".to_string()]);
        cache.insert("11px Arial", 120.0, "b", vec!["b".to_string()]);
        cache.insert("11px Arial", 120.0, "c", vec!["c".to_string()]);
        assert_eq!(cache.len(), 2);
        assert!(cache.get("11px Arial", 120.0, "a").is_none());
    }
}
