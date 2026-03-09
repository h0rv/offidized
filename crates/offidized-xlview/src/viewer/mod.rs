//! Main XlView struct - the primary entry point for the Canvas 2D viewer.
//!
//! This module provides the WASM-exported `XlView` struct that handles:
//! - Loading and parsing XLSX files
//! - Managing viewport state (scroll, zoom, active sheet)
//! - Coordinating between layout computation and Canvas 2D rendering
//! - Handling user interactions (scroll, click, keyboard)
//!
//! Event handlers for selection and copy are automatically registered when the
//! viewer is created - no manual JavaScript wiring required.

mod clipboard;
mod events;
mod scroll;

use std::collections::{HashMap, HashSet};
use std::sync::Arc;
use wasm_bindgen::prelude::*;

#[cfg(target_arch = "wasm32")]
use js_sys::Function;
#[cfg(target_arch = "wasm32")]
use js_sys::Reflect;
#[cfg(target_arch = "wasm32")]
use serde::Serialize;
#[cfg(target_arch = "wasm32")]
use std::cell::RefCell;
#[cfg(target_arch = "wasm32")]
use std::rc::Rc;
#[cfg(target_arch = "wasm32")]
use wasm_bindgen::closure::Closure;
#[cfg(target_arch = "wasm32")]
use web_sys::{
    HtmlCanvasElement, HtmlDivElement, HtmlElement, KeyboardEvent, MouseEvent, WheelEvent,
};

use crate::cell_ref::parse_sqref;
use crate::layout::{SheetLayout, Viewport};
#[cfg(target_arch = "wasm32")]
use crate::render::TextRunData;
use crate::render::{BorderStyleData, CellRenderData, CellStyleData};

#[cfg(target_arch = "wasm32")]
use crate::render::{CanvasRenderer, RenderBackend, RenderParams};
#[cfg(target_arch = "wasm32")]
use crate::types::selection::{HeaderConfig, Selection};
use crate::types::style::StyleRef;
#[cfg(target_arch = "wasm32")]
use crate::types::CompiledFormat;
use crate::types::Workbook;

/// Size of the resize handle in logical pixels
#[cfg(target_arch = "wasm32")]
pub(crate) const RESIZE_HANDLE_SIZE: f32 = 8.0;

/// Extra rows/cols to include around the viewport to reduce popping during scroll.
#[cfg(target_arch = "wasm32")]
const VISIBLE_PADDING: u32 = 1;

/// Tile size used by the renderer cache (logical pixels).
#[cfg(target_arch = "wasm32")]
const TILE_SIZE: f32 = 512.0;
/// Prefetch rings while idle (in tiles).
#[cfg(target_arch = "wasm32")]
const TILE_PREFETCH_IDLE: u32 = 1;
/// Prefetch rings during explicit idle warmup after load/sheet switch.
#[cfg(target_arch = "wasm32")]
const TILE_PREFETCH_PREWARM: u32 = 2;
/// Prefetch rings while actively scrolling (in tiles).
#[cfg(target_arch = "wasm32")]
const TILE_PREFETCH_SCROLLING: u32 = 0;
/// Number of idle base-render frames to run wider prefetch after load/sheet switch.
#[cfg(target_arch = "wasm32")]
const PREFETCH_WARMUP_FRAMES: u8 = 6;
/// Limit of new prefetch-only tile renders per frame when idle.
#[cfg(target_arch = "wasm32")]
const PREFETCH_RENDER_BUDGET_IDLE: u32 = 3;
/// Minimum adaptive prefetch-only tile render budget per frame.
#[cfg(target_arch = "wasm32")]
const PREFETCH_RENDER_BUDGET_MIN: u32 = 1;
/// Maximum adaptive prefetch-only tile render budget per frame.
#[cfg(target_arch = "wasm32")]
const PREFETCH_RENDER_BUDGET_MAX: u32 = 4;
/// Time window to consider scroll "active" after the latest scroll event.
#[cfg(target_arch = "wasm32")]
const SCROLL_ACTIVE_WINDOW_MS: f64 = 80.0;

#[cfg(target_arch = "wasm32")]
fn scroll_left_f64(element: &HtmlDivElement) -> f64 {
    Reflect::get(element.as_ref(), &JsValue::from_str("scrollLeft"))
        .ok()
        .and_then(|value| value.as_f64())
        .unwrap_or(element.scroll_left() as f64)
}

#[cfg(target_arch = "wasm32")]
fn scroll_top_f64(element: &HtmlDivElement) -> f64 {
    Reflect::get(element.as_ref(), &JsValue::from_str("scrollTop"))
        .ok()
        .and_then(|value| value.as_f64())
        .unwrap_or(element.scroll_top() as f64)
}

/// Convert a 0-based column index to Excel column letters (A, B, ..., Z, AA, AB, ...)
#[cfg(target_arch = "wasm32")]
pub(crate) fn col_to_letter(col: u32) -> String {
    let mut result = String::new();
    let mut n = col + 1; // Convert to 1-based
    while n > 0 {
        n -= 1;
        let c = char::from(b'A' + (n % 26) as u8);
        result.insert(0, c);
        n /= 26;
    }
    result
}

/// Shared state that can be accessed by event handlers (wasm32 only)
#[cfg(target_arch = "wasm32")]
pub(crate) struct SharedState {
    pub(crate) workbook: Option<Workbook>,
    pub(crate) layouts: Vec<Arc<SheetLayout>>,
    pub(crate) viewport: Viewport,
    pub(crate) active_sheet: usize,
    pub(crate) width: u32,
    pub(crate) height: u32,
    pub(crate) dpr: f32,
    pub(crate) needs_render: bool,
    pub(crate) needs_overlay_render: bool,
    pub(crate) sheet_names: Vec<String>,
    pub(crate) tab_colors: Vec<Option<String>>,
    pub(crate) render_styles: Vec<Option<CellStyleData>>,
    pub(crate) default_render_style: Option<CellStyleData>,
    pub(crate) visible_cells: Vec<CellRenderData>,
    pub(crate) last_visible_row_ranges: Vec<(u32, u32)>,
    pub(crate) last_visible_col_ranges: Vec<(u32, u32)>,
    pub(crate) last_visible_sheet: Option<usize>,
    pub(crate) render_callback: Option<Function>,
    pub(crate) scroll_settle_timer: Option<i32>,
    pub(crate) scroll_settle_closure: Option<Closure<dyn FnMut()>>,
    pub(crate) last_scroll_ms: f64,
    pub(crate) prefetch_warmup_frames: u8,
    pub(crate) last_base_draw_ms: f64,
    pub(crate) selection_start: Option<(u32, u32)>,
    pub(crate) selection_end: Option<(u32, u32)>,
    pub(crate) is_selecting: bool,
    pub(crate) is_resizing: bool, // True when dragging the resize handle
    // Header-related state
    pub(crate) show_headers: bool,
    pub(crate) header_config: HeaderConfig,
    pub(crate) selection: Option<Selection>,
    pub(crate) header_drag_mode: Option<HeaderDragMode>,
}

/// Mode for header dragging (selecting rows/columns)
#[cfg(target_arch = "wasm32")]
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) enum HeaderDragMode {
    /// Dragging across row headers
    Row,
    /// Dragging across column headers
    Column,
}

/// Target of a hit test (what was clicked)
#[cfg(target_arch = "wasm32")]
#[derive(Debug, Clone, Copy, PartialEq)]
pub(crate) enum HitTarget {
    /// A regular cell at (row, col)
    Cell(u32, u32),
    /// A row header at the given row index
    RowHeader(u32),
    /// A column header at the given column index
    ColumnHeader(u32),
    /// The corner header (select all)
    CornerHeader,
    /// Nothing (outside any interactive region)
    None,
}

// Timing helper for WASM metrics.
#[cfg(target_arch = "wasm32")]
pub(crate) fn now_ms() -> f64 {
    if let Some(window) = web_sys::window() {
        if let Some(perf) = window.performance() {
            return perf.now();
        }
    }
    js_sys::Date::now()
}

#[cfg(target_arch = "wasm32")]
#[derive(Serialize)]
struct LoadMetrics {
    parse_ms: f64,
    layout_ms: f64,
    total_ms: f64,
}

#[cfg(target_arch = "wasm32")]
#[derive(Serialize)]
struct RenderMetrics {
    prep_ms: f64,
    draw_ms: f64,
    total_ms: f64,
    visible_cells: u32,
    skipped: bool,
}

#[cfg(target_arch = "wasm32")]
#[derive(Serialize)]
struct ScrollMetrics {
    requested_dx: f32,
    requested_dy: f32,
    applied_dx: f32,
    applied_dy: f32,
    scroll_x: f32,
    scroll_y: f32,
    max_x: f32,
    max_y: f32,
    total_width: f32,
    total_height: f32,
    viewport_width: f32,
    viewport_height: f32,
}

/// The main viewer struct exported to JavaScript
#[wasm_bindgen]
pub struct XlView {
    #[cfg(target_arch = "wasm32")]
    state: Rc<RefCell<SharedState>>,
    #[cfg(target_arch = "wasm32")]
    renderer: CanvasRenderer,
    #[cfg(target_arch = "wasm32")]
    overlay_renderer: Option<CanvasRenderer>,
    #[cfg(target_arch = "wasm32")]
    #[allow(dead_code)]
    closures: Vec<Closure<dyn FnMut(MouseEvent)>>,
    #[cfg(target_arch = "wasm32")]
    #[allow(dead_code)]
    wheel_closure: Option<Closure<dyn FnMut(WheelEvent)>>,
    #[cfg(target_arch = "wasm32")]
    #[allow(dead_code)]
    key_closure: Option<Closure<dyn FnMut(KeyboardEvent)>>,
    #[cfg(target_arch = "wasm32")]
    #[allow(dead_code)]
    scroll_closure: Option<Closure<dyn FnMut(web_sys::Event)>>,
    #[cfg(target_arch = "wasm32")]
    #[allow(dead_code)] // Kept to maintain DOM reference
    flex_wrapper: Option<HtmlDivElement>,
    #[cfg(target_arch = "wasm32")]
    scroll_container: Option<HtmlDivElement>,
    #[cfg(target_arch = "wasm32")]
    scroll_spacer: Option<HtmlDivElement>,
    #[cfg(target_arch = "wasm32")]
    tab_bar: Option<HtmlDivElement>,

    // Non-wasm32 fields
    #[cfg(not(target_arch = "wasm32"))]
    workbook: Option<Workbook>,
    #[cfg(not(target_arch = "wasm32"))]
    layouts: Vec<Arc<SheetLayout>>,
    #[cfg(not(target_arch = "wasm32"))]
    viewport: Viewport,
    #[cfg(not(target_arch = "wasm32"))]
    active_sheet: usize,
    #[cfg(not(target_arch = "wasm32"))]
    #[allow(dead_code)]
    width: u32,
    #[cfg(not(target_arch = "wasm32"))]
    #[allow(dead_code)]
    height: u32,
    #[cfg(not(target_arch = "wasm32"))]
    #[allow(dead_code)]
    dpr: f32,
    #[cfg(not(target_arch = "wasm32"))]
    needs_render: bool,
    #[cfg(not(target_arch = "wasm32"))]
    #[allow(dead_code)]
    needs_overlay_render: bool,
    #[cfg(not(target_arch = "wasm32"))]
    #[allow(dead_code)]
    #[cfg(not(target_arch = "wasm32"))]
    sheet_names: Vec<String>,
    #[cfg(not(target_arch = "wasm32"))]
    #[allow(dead_code)]
    tab_colors: Vec<Option<String>>,
    #[cfg(not(target_arch = "wasm32"))]
    #[allow(dead_code)]
    render_styles: Vec<Option<CellStyleData>>,
    #[cfg(not(target_arch = "wasm32"))]
    #[allow(dead_code)]
    default_render_style: Option<CellStyleData>,
    #[cfg(not(target_arch = "wasm32"))]
    #[allow(dead_code)]
    visible_cells: Vec<CellRenderData>,
    #[cfg(not(target_arch = "wasm32"))]
    selection_start: Option<(u32, u32)>,
    #[cfg(not(target_arch = "wasm32"))]
    selection_end: Option<(u32, u32)>,
    #[cfg(not(target_arch = "wasm32"))]
    #[allow(dead_code)]
    is_selecting: bool,
    #[cfg(not(target_arch = "wasm32"))]
    #[allow(dead_code)]
    is_resizing: bool,
}

// ============================================================================
// WASM32 Implementation
// ============================================================================

#[cfg(target_arch = "wasm32")]
#[wasm_bindgen]
impl XlView {
    fn invalidate_visible_cell_cache(s: &mut SharedState, clear_visible_cells: bool) {
        if clear_visible_cells {
            s.visible_cells.clear();
        }
        s.last_visible_row_ranges.clear();
        s.last_visible_col_ranges.clear();
        s.last_visible_sheet = None;
    }

    fn can_reuse_visible_cells(
        s: &SharedState,
        active_sheet: usize,
        row_ranges: &[(u32, u32)],
        col_ranges: &[(u32, u32)],
    ) -> bool {
        s.last_visible_sheet == Some(active_sheet)
            && !s.visible_cells.is_empty()
            && s.last_visible_row_ranges.as_slice() == row_ranges
            && s.last_visible_col_ranges.as_slice() == col_ranges
    }

    fn merge_ranges(mut ranges: Vec<(u32, u32)>) -> Vec<(u32, u32)> {
        if ranges.len() <= 1 {
            return ranges;
        }
        ranges.sort_by_key(|r| r.0);
        let mut merged: Vec<(u32, u32)> = Vec::with_capacity(ranges.len());
        for (start, end) in ranges {
            if let Some(last) = merged.last_mut() {
                if start <= last.1.saturating_add(1) {
                    last.1 = last.1.max(end);
                    continue;
                }
            }
            merged.push((start, end));
        }
        merged
    }

    fn begin_idle_prewarm(s: &mut SharedState) {
        s.prefetch_warmup_frames = PREFETCH_WARMUP_FRAMES;
        s.last_scroll_ms = 0.0;
        s.last_base_draw_ms = 0.0;
    }

    fn is_scroll_active(last_scroll_ms: f64, now: f64) -> bool {
        (now - last_scroll_ms).max(0.0) < SCROLL_ACTIVE_WINDOW_MS
    }

    fn adaptive_prefetch_budget(last_base_draw_ms: f64) -> u32 {
        if !last_base_draw_ms.is_finite() || last_base_draw_ms <= 0.0 {
            return PREFETCH_RENDER_BUDGET_IDLE;
        }
        if last_base_draw_ms > 12.0 {
            PREFETCH_RENDER_BUDGET_MIN
        } else if last_base_draw_ms > 8.0 {
            2
        } else if last_base_draw_ms > 4.0 {
            PREFETCH_RENDER_BUDGET_IDLE
        } else {
            PREFETCH_RENDER_BUDGET_MAX
        }
    }

    fn tile_render_policy(
        scrolling_active: bool,
        prefetch_warmup_frames: u8,
        last_base_draw_ms: f64,
    ) -> (u32, Option<u32>) {
        if scrolling_active {
            (TILE_PREFETCH_SCROLLING, None)
        } else {
            let tile_prefetch = if prefetch_warmup_frames > 0 {
                TILE_PREFETCH_PREWARM
            } else {
                TILE_PREFETCH_IDLE
            };
            (
                tile_prefetch,
                Some(Self::adaptive_prefetch_budget(last_base_draw_ms)),
            )
        }
    }

    fn visible_cell_ranges(
        viewport: &Viewport,
        layout: &SheetLayout,
        tile_prefetch: u32,
    ) -> (Vec<(u32, u32)>, Vec<(u32, u32)>) {
        let frozen_width = layout.frozen_cols_width();
        let frozen_height = layout.frozen_rows_height();
        let scrollable_width = (viewport.width - frozen_width).max(0.0);
        let scrollable_height = (viewport.height - frozen_height).max(0.0);

        let (mut scroll_row_start, mut scroll_row_end, mut scroll_col_start, mut scroll_col_end) =
            if scrollable_width > 0.0 && scrollable_height > 0.0 {
                let origin_x = frozen_width;
                let origin_y = frozen_height;
                let tile_prefetch = match u16::try_from(tile_prefetch) {
                    Ok(value) => f32::from(value),
                    Err(_) => f32::from(u16::MAX),
                };
                let tile_x_start =
                    ((viewport.scroll_x - origin_x) / TILE_SIZE).floor() - tile_prefetch;
                let tile_x_end = ((viewport.scroll_x + scrollable_width - origin_x) / TILE_SIZE)
                    .floor()
                    + tile_prefetch;
                let tile_y_start =
                    ((viewport.scroll_y - origin_y) / TILE_SIZE).floor() - tile_prefetch;
                let tile_y_end = ((viewport.scroll_y + scrollable_height - origin_y) / TILE_SIZE)
                    .floor()
                    + tile_prefetch;
                let tile_start_x = origin_x + tile_x_start * TILE_SIZE;
                let tile_end_x = origin_x + (tile_x_end + 1.0) * TILE_SIZE;
                let tile_start_y = origin_y + tile_y_start * TILE_SIZE;
                let tile_end_y = origin_y + (tile_y_end + 1.0) * TILE_SIZE;

                let row_start = layout.row_at_y(tile_start_y).unwrap_or(layout.max_row);
                let row_end = layout.row_at_y(tile_end_y).unwrap_or(layout.max_row);
                let col_start = layout.col_at_x(tile_start_x).unwrap_or(layout.max_col);
                let col_end = layout.col_at_x(tile_end_x).unwrap_or(layout.max_col);
                (row_start, row_end, col_start, col_end)
            } else {
                let (row_start, row_end) = viewport.visible_rows(layout);
                let (col_start, col_end) = viewport.visible_cols(layout);
                (row_start, row_end, col_start, col_end)
            };

        scroll_row_start = scroll_row_start.max(layout.frozen_rows);
        scroll_col_start = scroll_col_start.max(layout.frozen_cols);
        scroll_row_start = scroll_row_start.saturating_sub(VISIBLE_PADDING);
        scroll_row_end = (scroll_row_end + VISIBLE_PADDING).min(layout.max_row);
        scroll_col_start = scroll_col_start.saturating_sub(VISIBLE_PADDING);
        scroll_col_end = (scroll_col_end + VISIBLE_PADDING).min(layout.max_col);

        let mut row_ranges = Vec::new();
        if layout.frozen_rows > 0 {
            row_ranges.push((0, layout.frozen_rows.saturating_sub(1)));
        }
        if scroll_row_start <= scroll_row_end {
            row_ranges.push((scroll_row_start, scroll_row_end));
        }

        let mut col_ranges = Vec::new();
        if layout.frozen_cols > 0 {
            col_ranges.push((0, layout.frozen_cols.saturating_sub(1)));
        }
        if scroll_col_start <= scroll_col_end {
            col_ranges.push((scroll_col_start, scroll_col_end));
        }

        (
            Self::merge_ranges(row_ranges),
            Self::merge_ranges(col_ranges),
        )
    }
    /// Create a new viewer instance
    ///
    /// This initializes Canvas 2D rendering and sets up event handlers automatically.
    /// Selection (mousedown/mousemove/mouseup) and copy (Ctrl+C) work out of the box.
    #[wasm_bindgen(constructor)]
    pub fn new(canvas: HtmlCanvasElement, dpr: f32) -> Result<XlView, JsValue> {
        console_error_panic_hook::set_once();

        let physical_width = canvas.width().max(1);
        let physical_height = canvas.height().max(1);

        let mut renderer =
            CanvasRenderer::new(canvas.clone()).map_err(|e| JsValue::from_str(&e.to_string()))?;
        renderer
            .init()
            .map_err(|e| JsValue::from_str(&e.to_string()))?;
        renderer.resize(physical_width, physical_height, dpr);

        let logical_width = physical_width as f32 / dpr;
        let logical_height = physical_height as f32 / dpr;

        let viewport = Viewport {
            scroll_x: 0.0,
            scroll_y: 0.0,
            width: logical_width,
            height: logical_height.max(100.0),
            scale: 1.0,
            tab_scroll_x: 0.0,
        };

        let state = Rc::new(RefCell::new(SharedState {
            workbook: None,
            layouts: Vec::new(),
            viewport,
            active_sheet: 0,
            width: physical_width,
            height: physical_height,
            dpr,
            needs_render: true,
            needs_overlay_render: false,
            sheet_names: Vec::new(),
            tab_colors: Vec::new(),
            render_styles: Vec::new(),
            default_render_style: None,
            visible_cells: Vec::new(),
            last_visible_row_ranges: Vec::new(),
            last_visible_col_ranges: Vec::new(),
            last_visible_sheet: None,
            render_callback: None,
            scroll_settle_timer: None,
            scroll_settle_closure: None,
            last_scroll_ms: 0.0,
            prefetch_warmup_frames: 0,
            last_base_draw_ms: 0.0,
            selection_start: None,
            selection_end: None,
            is_selecting: false,
            is_resizing: false,
            show_headers: true,
            header_config: HeaderConfig::default(),
            selection: None,
            header_drag_mode: None,
        }));

        // Set up native scrollbars with flexbox layout BEFORE wiring mouse events,
        // so the scroll_container is available as the event target.
        let (flex_wrapper, scroll_container, scroll_spacer, tab_bar, scroll_closure) =
            Self::setup_native_scroll(&canvas, None, &state, logical_width, logical_height);

        // Shared tooltip element for hover feedback (optional)
        let tooltip: Option<HtmlElement> = web_sys::window()
            .and_then(|window| window.document())
            .and_then(|document| document.get_element_by_id("tooltip"))
            .and_then(|element| element.dyn_into::<HtmlElement>().ok());

        // Mouse events go on the scroll container (z-index 1, on top of canvas).
        // Use the container's own bounding rect for coordinate extraction (not
        // event.target(), which could be the spacer child).
        let event_target: &HtmlElement = scroll_container
            .as_ref()
            .map(|c| c.as_ref() as &HtmlElement)
            .unwrap_or(&canvas);
        let mut closures: Vec<Closure<dyn FnMut(MouseEvent)>> = Vec::new();

        // Mouse down
        {
            let state = Rc::clone(&state);
            let container_ref = event_target.clone();
            let closure = Closure::wrap(Box::new(move |event: MouseEvent| {
                let rect = container_ref.get_bounding_client_rect();
                let x = event.client_x() as f32 - rect.left() as f32;
                let y = event.client_y() as f32 - rect.top() as f32;
                Self::internal_mouse_down(&state, x, y);
            }) as Box<dyn FnMut(MouseEvent)>);
            event_target
                .add_event_listener_with_callback("mousedown", closure.as_ref().unchecked_ref())
                .ok();
            closures.push(closure);
        }

        // Mouse move (selection + hover feedback)
        {
            let state = Rc::clone(&state);
            let container_for_move = event_target.clone();
            let tooltip = tooltip.clone();
            let closure = Closure::wrap(Box::new(move |event: MouseEvent| {
                let rect = container_for_move.get_bounding_client_rect();
                let x = event.client_x() as f32 - rect.left() as f32;
                let y = event.client_y() as f32 - rect.top() as f32;
                Self::internal_mouse_move(&state, x, y);
                Self::update_hover_ui(
                    &state,
                    &container_for_move,
                    tooltip.as_ref(),
                    x,
                    y,
                    event.client_x(),
                    event.client_y(),
                );
            }) as Box<dyn FnMut(MouseEvent)>);
            event_target
                .add_event_listener_with_callback("mousemove", closure.as_ref().unchecked_ref())
                .ok();
            closures.push(closure);
        }

        // Mouse up
        {
            let state = Rc::clone(&state);
            let closure = Closure::wrap(Box::new(move |_event: MouseEvent| {
                Self::internal_mouse_up(&state);
            }) as Box<dyn FnMut(MouseEvent)>);
            event_target
                .add_event_listener_with_callback("mouseup", closure.as_ref().unchecked_ref())
                .ok();
            closures.push(closure);
        }

        // Mouse leave (hide tooltip + reset cursor)
        {
            let container_for_leave = event_target.clone();
            let tooltip = tooltip.clone();
            let closure = Closure::wrap(Box::new(move |_event: MouseEvent| {
                let _ = container_for_leave
                    .style()
                    .set_property("cursor", "default");
                if let Some(tooltip) = tooltip.as_ref() {
                    let _ = tooltip.style().set_property("display", "none");
                }
            }) as Box<dyn FnMut(MouseEvent)>);
            event_target
                .add_event_listener_with_callback("mouseleave", closure.as_ref().unchecked_ref())
                .ok();
            closures.push(closure);
        }

        // Click (tabs + hyperlinks)
        {
            let state = Rc::clone(&state);
            let container_ref = event_target.clone();
            let closure = Closure::wrap(Box::new(move |event: MouseEvent| {
                let rect = container_ref.get_bounding_client_rect();
                let x = event.client_x() as f32 - rect.left() as f32;
                let y = event.client_y() as f32 - rect.top() as f32;
                Self::internal_click(&state, x, y);
            }) as Box<dyn FnMut(MouseEvent)>);
            event_target
                .add_event_listener_with_callback("click", closure.as_ref().unchecked_ref())
                .ok();
            closures.push(closure);
        }

        // Wheel scroll is handled by native scroll container (no manual handler needed)
        let wheel_closure: Option<Closure<dyn FnMut(WheelEvent)>> = None;

        // Keyboard handler on document for Ctrl+C
        let key_closure = {
            let state = Rc::clone(&state);
            let closure = Closure::wrap(Box::new(move |event: KeyboardEvent| {
                let key = event.key();
                let ctrl = event.ctrl_key() || event.meta_key();
                if Self::internal_key_down(&state, &key, ctrl) {
                    event.prevent_default();
                }
            }) as Box<dyn FnMut(KeyboardEvent)>);

            if let Some(window) = web_sys::window() {
                if let Some(document) = window.document() {
                    document
                        .add_event_listener_with_callback(
                            "keydown",
                            closure.as_ref().unchecked_ref(),
                        )
                        .ok();
                }
            }
            Some(closure)
        };

        Ok(XlView {
            state,
            renderer,
            overlay_renderer: None,
            closures,
            wheel_closure,
            key_closure,
            scroll_closure,
            flex_wrapper,
            scroll_container,
            scroll_spacer,
            tab_bar,
        })
    }

    #[wasm_bindgen(js_name = "newWithOverlay")]
    pub fn new_with_overlay(
        base_canvas: HtmlCanvasElement,
        overlay_canvas: HtmlCanvasElement,
        dpr: f32,
    ) -> Result<XlView, JsValue> {
        console_error_panic_hook::set_once();

        let physical_width = base_canvas.width().max(1);
        let physical_height = base_canvas.height().max(1);

        let mut renderer = CanvasRenderer::new(base_canvas.clone())
            .map_err(|e| JsValue::from_str(&e.to_string()))?;
        renderer
            .init()
            .map_err(|e| JsValue::from_str(&e.to_string()))?;
        renderer.resize(physical_width, physical_height, dpr);

        let overlay_dom = overlay_canvas.clone();
        let mut overlay_renderer =
            CanvasRenderer::new(overlay_canvas).map_err(|e| JsValue::from_str(&e.to_string()))?;
        overlay_renderer
            .init()
            .map_err(|e| JsValue::from_str(&e.to_string()))?;
        overlay_renderer.resize(physical_width, physical_height, dpr);

        let logical_width = physical_width as f32 / dpr;
        let logical_height = physical_height as f32 / dpr;

        let viewport = Viewport {
            scroll_x: 0.0,
            scroll_y: 0.0,
            width: logical_width,
            height: logical_height.max(100.0),
            scale: 1.0,
            tab_scroll_x: 0.0,
        };

        let state = Rc::new(RefCell::new(SharedState {
            workbook: None,
            layouts: Vec::new(),
            viewport,
            active_sheet: 0,
            width: physical_width,
            height: physical_height,
            dpr,
            needs_render: true,
            needs_overlay_render: false,
            sheet_names: Vec::new(),
            tab_colors: Vec::new(),
            render_styles: Vec::new(),
            default_render_style: None,
            visible_cells: Vec::new(),
            last_visible_row_ranges: Vec::new(),
            last_visible_col_ranges: Vec::new(),
            last_visible_sheet: None,
            render_callback: None,
            scroll_settle_timer: None,
            scroll_settle_closure: None,
            last_scroll_ms: 0.0,
            prefetch_warmup_frames: 0,
            last_base_draw_ms: 0.0,
            selection_start: None,
            selection_end: None,
            is_selecting: false,
            is_resizing: false,
            show_headers: true,
            header_config: HeaderConfig::default(),
            selection: None,
            header_drag_mode: None,
        }));

        // Set up native scrollbars with flexbox layout BEFORE wiring mouse events,
        // so the scroll_container is available as the event target.
        let (flex_wrapper, scroll_container, scroll_spacer, tab_bar, scroll_closure) =
            Self::setup_native_scroll(
                &base_canvas,
                Some(&overlay_dom),
                &state,
                logical_width,
                logical_height,
            );

        let tooltip: Option<HtmlElement> = web_sys::window()
            .and_then(|window| window.document())
            .and_then(|document| document.get_element_by_id("tooltip"))
            .and_then(|element| element.dyn_into::<HtmlElement>().ok());

        // Mouse events go on the scroll container (z-index 1, on top of canvas).
        let event_target: &HtmlElement = scroll_container
            .as_ref()
            .map(|c| c.as_ref() as &HtmlElement)
            .unwrap_or(&base_canvas);
        let mut closures: Vec<Closure<dyn FnMut(MouseEvent)>> = Vec::new();

        {
            let state = Rc::clone(&state);
            let container_ref = event_target.clone();
            let closure = Closure::wrap(Box::new(move |event: MouseEvent| {
                let rect = container_ref.get_bounding_client_rect();
                let x = event.client_x() as f32 - rect.left() as f32;
                let y = event.client_y() as f32 - rect.top() as f32;
                Self::internal_mouse_down(&state, x, y);
            }) as Box<dyn FnMut(MouseEvent)>);
            event_target
                .add_event_listener_with_callback("mousedown", closure.as_ref().unchecked_ref())
                .ok();
            closures.push(closure);
        }

        {
            let state = Rc::clone(&state);
            let container_for_move = event_target.clone();
            let tooltip = tooltip.clone();
            let closure = Closure::wrap(Box::new(move |event: MouseEvent| {
                let rect = container_for_move.get_bounding_client_rect();
                let x = event.client_x() as f32 - rect.left() as f32;
                let y = event.client_y() as f32 - rect.top() as f32;
                Self::internal_mouse_move(&state, x, y);
                Self::update_hover_ui(
                    &state,
                    &container_for_move,
                    tooltip.as_ref(),
                    x,
                    y,
                    event.client_x(),
                    event.client_y(),
                );
            }) as Box<dyn FnMut(MouseEvent)>);
            event_target
                .add_event_listener_with_callback("mousemove", closure.as_ref().unchecked_ref())
                .ok();
            closures.push(closure);
        }

        {
            let state = Rc::clone(&state);
            let closure = Closure::wrap(Box::new(move |_event: MouseEvent| {
                Self::internal_mouse_up(&state);
            }) as Box<dyn FnMut(MouseEvent)>);
            event_target
                .add_event_listener_with_callback("mouseup", closure.as_ref().unchecked_ref())
                .ok();
            closures.push(closure);
        }

        {
            let container_for_leave = event_target.clone();
            let tooltip = tooltip.clone();
            let closure = Closure::wrap(Box::new(move |_event: MouseEvent| {
                let _ = container_for_leave
                    .style()
                    .set_property("cursor", "default");
                if let Some(tooltip) = tooltip.as_ref() {
                    let _ = tooltip.style().set_property("display", "none");
                }
            }) as Box<dyn FnMut(MouseEvent)>);
            event_target
                .add_event_listener_with_callback("mouseleave", closure.as_ref().unchecked_ref())
                .ok();
            closures.push(closure);
        }

        {
            let state = Rc::clone(&state);
            let container_ref = event_target.clone();
            let closure = Closure::wrap(Box::new(move |event: MouseEvent| {
                let rect = container_ref.get_bounding_client_rect();
                let x = event.client_x() as f32 - rect.left() as f32;
                let y = event.client_y() as f32 - rect.top() as f32;
                Self::internal_click(&state, x, y);
            }) as Box<dyn FnMut(MouseEvent)>);
            event_target
                .add_event_listener_with_callback("click", closure.as_ref().unchecked_ref())
                .ok();
            closures.push(closure);
        }

        // Wheel scroll is handled by native scroll container (no manual handler needed)
        let wheel_closure: Option<Closure<dyn FnMut(WheelEvent)>> = None;

        let key_closure = {
            let state = Rc::clone(&state);
            let closure = Closure::wrap(Box::new(move |event: KeyboardEvent| {
                let key = event.key();
                let ctrl = event.ctrl_key() || event.meta_key();
                if Self::internal_key_down(&state, &key, ctrl) {
                    event.prevent_default();
                }
            }) as Box<dyn FnMut(KeyboardEvent)>);

            if let Some(window) = web_sys::window() {
                if let Some(document) = window.document() {
                    document
                        .add_event_listener_with_callback(
                            "keydown",
                            closure.as_ref().unchecked_ref(),
                        )
                        .ok();
                }
            }
            Some(closure)
        };

        Ok(XlView {
            state,
            renderer,
            overlay_renderer: Some(overlay_renderer),
            closures,
            wheel_closure,
            key_closure,
            scroll_closure,
            flex_wrapper,
            scroll_container,
            scroll_spacer,
            tab_bar,
        })
    }

    /// Set up native browser scrollbars with flexbox layout
    ///
    /// Creates this DOM structure:
    /// ```text
    /// flex_wrapper (display: flex, flex-direction: column, 100% height)
    /// ├── canvas (position: absolute, pointer-events: none, z-index: 0, 1x viewport)
    /// ├── scroll_container (flex: 1, overflow: auto, z-index: 1, background: transparent)
    /// │   └── spacer (sized to scrollable content)
    /// ├── overlay (position: absolute, pointer-events: none, z-index: 2)
    /// └── tab_bar (auto height, DOM-based tabs)
    /// ```
    ///
    /// The canvas is a direct child of flex_wrapper (behind the scroll container).
    /// On each scroll frame, the renderer blits cached tiles directly to the
    /// viewport-sized canvas — no buffer underrun is possible.
    fn setup_native_scroll(
        canvas: &HtmlCanvasElement,
        overlay_canvas: Option<&HtmlCanvasElement>,
        state: &Rc<RefCell<SharedState>>,
        width: f32,
        height: f32,
    ) -> (
        Option<HtmlDivElement>,
        Option<HtmlDivElement>,
        Option<HtmlDivElement>,
        Option<HtmlDivElement>,
        Option<Closure<dyn FnMut(web_sys::Event)>>,
    ) {
        let Some(window) = web_sys::window() else {
            return (None, None, None, None, None);
        };
        let Some(document) = window.document() else {
            return (None, None, None, None, None);
        };
        let Some(parent) = canvas.parent_element() else {
            return (None, None, None, None, None);
        };

        // Helper to create a div element
        let create_div = || -> Option<HtmlDivElement> {
            document
                .create_element("div")
                .ok()
                .and_then(|el| el.dyn_into::<HtmlDivElement>().ok())
        };

        let Some(flex_wrapper) = create_div() else {
            return (None, None, None, None, None);
        };
        let Some(scroll_container) = create_div() else {
            return (None, None, None, None, None);
        };
        let Some(spacer) = create_div() else {
            return (None, None, None, None, None);
        };
        let Some(tab_bar) = create_div() else {
            return (None, None, None, None, None);
        };

        // Ensure parent has position for absolute children
        if let Some(parent_el) = parent.dyn_ref::<HtmlElement>() {
            let parent_style = parent_el.style();
            // Only set position if not already set
            if parent_style
                .get_property_value("position")
                .unwrap_or_default()
                .is_empty()
            {
                let _ = parent_style.set_property("position", "relative");
            }
        }

        // Flex wrapper: column layout, fills parent completely
        let wrapper_style = flex_wrapper.style();
        let _ = wrapper_style.set_property("display", "flex");
        let _ = wrapper_style.set_property("flex-direction", "column");
        let _ = wrapper_style.set_property("width", "100%");
        let _ = wrapper_style.set_property("height", "100%");
        let _ = wrapper_style.set_property("position", "absolute");
        let _ = wrapper_style.set_property("top", "0");
        let _ = wrapper_style.set_property("left", "0");

        // Scroll container: fills remaining space, handles scrolling.
        // Sits on top of canvas (z-index 1) so scrollbars are visible and
        // mouse events (wheel/click) hit the container. Background is transparent
        // so the canvas underneath shows through.
        let container_style = scroll_container.style();
        let _ = container_style.set_property("flex", "1");
        let _ = container_style.set_property("overflow", "auto");
        let _ = container_style.set_property("position", "relative");
        let _ = container_style.set_property("z-index", "1");
        let _ = container_style.set_property("background", "transparent");
        let _ = container_style.set_property("min-height", "0"); // Important for flex children
                                                                 // Mark so JS can find the scroll container for viewport sizing
        let _ = scroll_container.set_attribute("data-xlview-scroll", "");

        // Spacer: sized to content to create scroll area.
        let spacer_style = spacer.style();
        let _ = spacer_style.set_property("position", "absolute");
        let _ = spacer_style.set_property("top", "0");
        let _ = spacer_style.set_property("left", "0");
        let _ = spacer_style.set_property("width", &format!("{}px", width));
        let _ = spacer_style.set_property("height", &format!("{}px", height));

        // Main canvas: viewport-sized, sits behind scroll container (z-index 0).
        // Tiles are blitted directly to this canvas on each scroll frame.
        let canvas_style = canvas.style();
        let _ = canvas_style.set_property("inset", "auto");
        let _ = canvas_style.set_property("position", "absolute");
        let _ = canvas_style.set_property("top", "0");
        let _ = canvas_style.set_property("left", "0");
        let _ = canvas_style.set_property("right", "auto");
        let _ = canvas_style.set_property("bottom", "auto");
        let _ = canvas_style.set_property("pointer-events", "none");
        let _ = canvas_style.set_property("z-index", "0");

        if let Some(overlay) = overlay_canvas {
            let overlay_style = overlay.style();
            // Overlay sits on top of everything (z-index 2).
            let _ = overlay_style.set_property("inset", "auto");
            let _ = overlay_style.set_property("position", "absolute");
            let _ = overlay_style.set_property("top", "0");
            let _ = overlay_style.set_property("left", "0");
            let _ = overlay_style.set_property("right", "auto");
            let _ = overlay_style.set_property("bottom", "auto");
            let _ = overlay_style.set_property("pointer-events", "none");
            let _ = overlay_style.set_property("z-index", "2");
            // Promote to own GPU compositor layer so overlay repaints during scroll
            // don't trigger expensive recomposites of the parent layer.
            let _ = overlay_style.set_property("will-change", "transform");
        }

        // Tab bar: auto height at bottom
        let tab_style = tab_bar.style();
        let _ = tab_style.set_property("display", "flex");
        let _ = tab_style.set_property("align-items", "center");
        let _ = tab_style.set_property("gap", "1px");
        let _ = tab_style.set_property("padding", "4px 8px");
        let _ = tab_style.set_property("background", "#F3F3F3");
        let _ = tab_style.set_property("border-top", "1px solid #E0E0E0");
        let _ = tab_style.set_property("overflow-x", "auto");
        let _ = tab_style.set_property("flex-shrink", "0");

        // Build DOM structure.
        // Canvas is a direct child of flex_wrapper (behind the scroll container).
        // Scroll container sits on top with transparent background.
        let _ = parent.insert_before(&flex_wrapper, Some(canvas));
        let _ = flex_wrapper.append_child(canvas);
        let _ = scroll_container.append_child(&spacer);
        let _ = flex_wrapper.append_child(&scroll_container);
        if let Some(overlay) = overlay_canvas {
            let _ = flex_wrapper.append_child(overlay);
        }
        let _ = flex_wrapper.append_child(&tab_bar);

        // Scroll event: unconditionally request render on every scroll frame.
        // Since the canvas is viewport-sized and tiles are blitted directly,
        // every frame must re-composite to show the correct viewport position.
        let state_clone = Rc::clone(&state);
        let scroll_closure = Closure::wrap(Box::new(move |_event: web_sys::Event| {
            let mut s = state_clone.borrow_mut();
            s.last_scroll_ms = now_ms();
            s.needs_render = true;
            s.needs_overlay_render = true;
            if let Some(callback) = s.render_callback.clone() {
                drop(s);
                let _ = callback.call0(&JsValue::NULL);
            }
        }) as Box<dyn FnMut(web_sys::Event)>);

        let _ = scroll_container
            .add_event_listener_with_callback("scroll", scroll_closure.as_ref().unchecked_ref());

        (
            Some(flex_wrapper),
            Some(scroll_container),
            Some(spacer),
            Some(tab_bar),
            Some(scroll_closure),
        )
    }

    /// Update the DOM tab bar with current sheet names
    fn update_tab_bar(&self) {
        let Some(tab_bar) = &self.tab_bar else {
            return;
        };

        // Extract data from state then drop the borrow before any DOM mutations.
        // DOM changes (set_inner_html, append_child) can synchronously fire scroll
        // events whose closure needs borrow_mut().
        let (sheet_names, tab_colors, active_sheet) = {
            let s = self.state.borrow();
            (s.sheet_names.clone(), s.tab_colors.clone(), s.active_sheet)
        }; // borrow dropped here — safe to mutate DOM

        // Clear existing tabs
        tab_bar.set_inner_html("");

        let Some(document) = web_sys::window().and_then(|w| w.document()) else {
            return;
        };

        // Create tab buttons
        for (i, name) in sheet_names.iter().enumerate() {
            let Ok(button) = document.create_element("button") else {
                continue;
            };
            button.set_text_content(Some(name));

            let style = match button.dyn_ref::<HtmlElement>() {
                Some(el) => el.style(),
                None => continue,
            };

            // Base button styles
            let _ = style.set_property("padding", "4px 12px");
            let _ = style.set_property("border", "none");
            let _ = style.set_property("border-radius", "4px 4px 0 0");
            let _ = style.set_property("cursor", "pointer");
            let _ = style.set_property("font-size", "12px");
            let _ = style.set_property("font-family", "system-ui, -apple-system, sans-serif");
            let _ = style.set_property("white-space", "nowrap");

            // Active vs inactive styles
            if i == active_sheet {
                let _ = style.set_property("background", "#FFFFFF");
                let _ = style.set_property("font-weight", "500");
                let _ = style.set_property("border-bottom", "2px solid #217346");
            } else {
                let _ = style.set_property("background", "#E8E8E8");
                let _ = style.set_property("font-weight", "normal");

                // Apply custom tab color if present
                if let Some(Some(color)) = tab_colors.get(i) {
                    let _ = style.set_property("background", color);
                }
            }

            // Add click handler
            let state_clone = Rc::clone(&self.state);
            let closure = Closure::wrap(Box::new(move |_: web_sys::Event| {
                let callback = {
                    let mut s = state_clone.borrow_mut();
                    if s.active_sheet != i && Self::set_active_sheet_state(&mut s, i) {
                        s.needs_render = true;
                    }
                    s.render_callback.clone()
                };
                if let Some(callback) = callback {
                    let _ = callback.call0(&JsValue::NULL);
                }
            }) as Box<dyn FnMut(web_sys::Event)>);

            let _ =
                button.add_event_listener_with_callback("click", closure.as_ref().unchecked_ref());
            closure.forget(); // Leak the closure (it lives as long as the button)

            let _ = tab_bar.append_child(&button);
        }
    }

    /// Load a font (no-op for Canvas 2D, fonts loaded via CSS)
    #[wasm_bindgen]
    pub fn load_font(&mut self, _font_data: &[u8]) -> Result<(), JsValue> {
        let callback = {
            let mut s = self.state.borrow_mut();
            s.needs_render = true;
            s.render_callback.clone()
        };
        Self::invoke_render_callback(callback);
        Ok(())
    }

    /// Load an XLSX file from bytes
    #[wasm_bindgen]
    pub fn load(&mut self, data: &[u8]) -> Result<(), JsValue> {
        let xlsx_wb = offidized_xlsx::Workbook::from_bytes(data)
            .map_err(|e| JsValue::from_str(&e.to_string()))?;
        let mut workbook = crate::adapter::convert_workbook(&xlsx_wb);
        let render_styles: Vec<Option<CellStyleData>> =
            crate::adapter::build_render_style_cache(&xlsx_wb);
        let default_render_style = xlsx_wb.style(0).map(|s| {
            crate::adapter::style_to_render(s, xlsx_wb.theme_colors(), xlsx_wb.indexed_colors())
        });
        let layouts: Vec<Arc<SheetLayout>> = workbook
            .sheets
            .iter()
            .map(Self::compute_layout)
            .map(Arc::new)
            .collect();
        let sheet_names: Vec<String> = workbook.sheets.iter().map(|sh| sh.name.clone()).collect();
        let tab_colors: Vec<Option<String>> = workbook
            .sheets
            .iter()
            .map(|sh| sh.tab_color.clone())
            .collect();
        Self::prepare_workbook_caches(&mut workbook);
        let mut s = self.state.borrow_mut();
        // Initialize scroll to frozen boundary (for frozen panes support)
        if let Some(layout) = layouts.first() {
            s.viewport.scroll_x = layout.frozen_cols_width();
            s.viewport.scroll_y = layout.frozen_rows_height();
        } else {
            s.viewport.scroll_x = 0.0;
            s.viewport.scroll_y = 0.0;
        }
        s.layouts = layouts;
        s.workbook = Some(workbook);
        s.sheet_names = sheet_names;
        s.tab_colors = tab_colors;
        s.render_styles = render_styles;
        s.default_render_style = default_render_style;
        s.active_sheet = 0;
        Self::invalidate_visible_cell_cache(&mut s, true);
        Self::begin_idle_prewarm(&mut s);
        s.selection_start = None;
        s.selection_end = None;
        s.needs_render = true;
        let callback = s.render_callback.clone();
        drop(s);

        // Update scroll spacer to match content size
        self.update_scroll_spacer();

        // Update DOM tab bar
        self.update_tab_bar();

        Self::invoke_render_callback(callback);
        Ok(())
    }

    /// Update the scroll spacer size to match the current layout
    fn update_scroll_spacer(&self) {
        let Some(spacer) = &self.scroll_spacer else {
            return;
        };

        // Compute dimensions while holding the borrow, then drop before DOM mutations.
        // Setting scroll position fires a synchronous scroll event whose closure
        // needs borrow_mut(), so we must not hold any borrow at that point.
        let (scrollable_width, scrollable_height) = {
            let s = self.state.borrow();
            let Some(layout) = s.layouts.get(s.active_sheet) else {
                return;
            };

            // Spacer is sized to SCROLLABLE content only (total minus frozen regions).
            // This simplifies scroll coordinate mapping:
            // - container.scroll_left=0 corresponds to viewport at frozen boundary
            // - No need to add/subtract frozen offsets
            let frozen_w = layout.frozen_cols_width();
            let frozen_h = layout.frozen_rows_height();
            let header_w = if s.show_headers {
                s.header_config.row_header_width
            } else {
                0.0
            };
            let header_h = if s.show_headers {
                s.header_config.col_header_height
            } else {
                0.0
            };
            // Add header dimensions to the scrollable range so the native scroll
            // container can reach the true content edge.
            let w = (layout.total_width() - frozen_w).max(0.0) + header_w;
            let h = (layout.total_height() - frozen_h).max(0.0) + header_h;
            (w, h)
        }; // borrow dropped here — safe to mutate DOM

        let spacer_style = spacer.style();
        let _ = spacer_style.set_property("width", &format!("{}px", scrollable_width));
        let _ = spacer_style.set_property("height", &format!("{}px", scrollable_height));

        // Reset container scroll to 0,0 (start of scrollable area)
        if let Some(container) = &self.scroll_container {
            container.set_scroll_left(0);
            container.set_scroll_top(0);
        }
    }

    /// Sync viewport scroll from native scroll container
    fn sync_scroll_from_container(&self) {
        let Some(container) = &self.scroll_container else {
            return;
        };
        let mut s = self.state.borrow_mut();
        let Some(layout) = s.layouts.get(s.active_sheet) else {
            return;
        };
        let frozen_w = layout.frozen_cols_width();
        let frozen_h = layout.frozen_rows_height();
        let new_x = scroll_left_f64(container) as f32 + frozen_w;
        let new_y = scroll_top_f64(container) as f32 + frozen_h;
        s.viewport.scroll_x = new_x;
        s.viewport.scroll_y = new_y;
    }

    /// Load an XLSX file from bytes and return internal timing metrics.
    #[wasm_bindgen]
    pub fn load_with_metrics(&mut self, data: &[u8]) -> Result<JsValue, JsValue> {
        let total_start = now_ms();

        let parse_start = now_ms();
        let xlsx_wb = offidized_xlsx::Workbook::from_bytes(data)
            .map_err(|e| JsValue::from_str(&e.to_string()))?;
        let mut workbook = crate::adapter::convert_workbook(&xlsx_wb);
        let parse_ms = now_ms() - parse_start;

        let render_styles: Vec<Option<CellStyleData>> =
            crate::adapter::build_render_style_cache(&xlsx_wb);
        let default_render_style = xlsx_wb.style(0).map(|s| {
            crate::adapter::style_to_render(s, xlsx_wb.theme_colors(), xlsx_wb.indexed_colors())
        });

        let layout_start = now_ms();
        let layouts: Vec<Arc<SheetLayout>> = workbook
            .sheets
            .iter()
            .map(Self::compute_layout)
            .map(Arc::new)
            .collect();
        let layout_ms = now_ms() - layout_start;
        let sheet_names: Vec<String> = workbook.sheets.iter().map(|sh| sh.name.clone()).collect();
        let tab_colors: Vec<Option<String>> = workbook
            .sheets
            .iter()
            .map(|sh| sh.tab_color.clone())
            .collect();
        Self::prepare_workbook_caches(&mut workbook);

        let mut s = self.state.borrow_mut();
        if let Some(layout) = layouts.first() {
            s.viewport.scroll_x = layout.frozen_cols_width();
            s.viewport.scroll_y = layout.frozen_rows_height();
        } else {
            s.viewport.scroll_x = 0.0;
            s.viewport.scroll_y = 0.0;
        }
        s.layouts = layouts;
        s.workbook = Some(workbook);
        s.sheet_names = sheet_names;
        s.tab_colors = tab_colors;
        s.render_styles = render_styles;
        s.default_render_style = default_render_style;
        s.active_sheet = 0;
        Self::invalidate_visible_cell_cache(&mut s, true);
        Self::begin_idle_prewarm(&mut s);
        s.selection_start = None;
        s.selection_end = None;
        s.needs_render = true;
        let callback = s.render_callback.clone();
        drop(s);

        // Update scroll spacer to match content size
        self.update_scroll_spacer();

        Self::invoke_render_callback(callback);

        let total_ms = now_ms() - total_start;
        serde_wasm_bindgen::to_value(&LoadMetrics {
            parse_ms,
            layout_ms,
            total_ms,
        })
        .map_err(|e| JsValue::from_str(&format!("Serialization error: {e}")))
    }

    /// Resize the viewport (dimensions in physical pixels)
    #[wasm_bindgen]
    pub fn resize(&mut self, physical_width: u32, physical_height: u32, dpr: f32) {
        let logical_width = physical_width as f32 / dpr;
        let logical_height = physical_height as f32 / dpr;

        let callback = {
            let mut s = self.state.borrow_mut();
            s.width = physical_width;
            s.height = physical_height;
            s.dpr = dpr;
            s.viewport.width = logical_width;
            s.viewport.height = logical_height.max(0.0);
            s.needs_render = true;
            s.render_callback.clone()
        };

        // Canvas is viewport-sized (1x). Tiles are blitted directly.
        self.renderer.resize(physical_width, physical_height, dpr);
        self.renderer
            .set_canvas_css_size(logical_width, logical_height);

        // Overlay stays viewport-sized.
        if let Some(overlay) = self.overlay_renderer.as_mut() {
            overlay.resize(physical_width, physical_height, dpr);
        }

        // Update scroll spacer if workbook is loaded (layout exists)
        self.update_scroll_spacer();

        Self::invoke_render_callback(callback);
    }

    /// Get current scroll debug info (for debugging scroll coordinate issues)
    #[wasm_bindgen]
    pub fn get_scroll_debug(&self) -> JsValue {
        let s = self.state.borrow();
        let container_scroll = self
            .scroll_container
            .as_ref()
            .map(|c| (scroll_left_f64(c) as f32, scroll_top_f64(c) as f32));
        let layout_frozen = s
            .layouts
            .get(s.active_sheet)
            .map(|l| (l.frozen_cols_width(), l.frozen_rows_height()));
        let (visible_start_row, visible_end_row) = s
            .layouts
            .get(s.active_sheet)
            .map(|l| s.viewport.visible_rows(l))
            .unwrap_or((0, 0));

        // Build JS object directly to avoid serialization issues
        let obj = js_sys::Object::new();
        let _ = Reflect::set(
            &obj,
            &"viewport_scroll_x".into(),
            &s.viewport.scroll_x.into(),
        );
        let _ = Reflect::set(
            &obj,
            &"viewport_scroll_y".into(),
            &s.viewport.scroll_y.into(),
        );
        let _ = Reflect::set(&obj, &"viewport_width".into(), &s.viewport.width.into());
        let _ = Reflect::set(&obj, &"viewport_height".into(), &s.viewport.height.into());
        let _ = Reflect::set(
            &obj,
            &"container_scroll_left".into(),
            &container_scroll.map(|(x, _)| x).unwrap_or(0.0).into(),
        );
        let _ = Reflect::set(
            &obj,
            &"container_scroll_top".into(),
            &container_scroll.map(|(_, y)| y).unwrap_or(0.0).into(),
        );
        let _ = Reflect::set(
            &obj,
            &"frozen_cols_width".into(),
            &layout_frozen.map(|(x, _)| x).unwrap_or(0.0).into(),
        );
        let _ = Reflect::set(
            &obj,
            &"frozen_rows_height".into(),
            &layout_frozen.map(|(_, y)| y).unwrap_or(0.0).into(),
        );
        let _ = Reflect::set(&obj, &"visible_start_row".into(), &visible_start_row.into());
        let _ = Reflect::set(&obj, &"visible_end_row".into(), &visible_end_row.into());
        let _ = Reflect::set(&obj, &"show_headers".into(), &s.show_headers.into());
        obj.into()
    }

    /// Scroll by delta amounts
    #[wasm_bindgen]
    pub fn scroll(&mut self, delta_x: f32, delta_y: f32) {
        // With native scroll container, just update the container's scroll position
        // and let the scroll event handler trigger a render
        if let Some(container) = &self.scroll_container {
            let new_left = container.scroll_left() + delta_x as i32;
            let new_top = container.scroll_top() + delta_y as i32;
            container.set_scroll_left(new_left);
            container.set_scroll_top(new_top);
        } else {
            // Fallback for when no scroll container (shouldn't happen normally)
            let (callback, should_schedule) = (|| {
                let mut s = self.state.borrow_mut();
                if let Some((_dx, _dy)) = Self::scroll_state(&mut s, delta_x, delta_y) {
                    s.needs_render = true;
                    return (s.render_callback.clone(), true);
                }
                (None, false)
            })();
            if should_schedule {
                Self::schedule_scroll_settle_timeout(&self.state);
            }
            Self::invoke_render_callback(callback);
        }
    }

    /// Scroll by delta amounts and return scroll/clamp metrics for profiling.
    #[wasm_bindgen]
    pub fn scroll_with_metrics(&mut self, delta_x: f32, delta_y: f32) -> Result<JsValue, JsValue> {
        if let Some(container) = self.scroll_container.clone() {
            let (callback, should_schedule, metrics) = (|| {
                let mut s = self.state.borrow_mut();
                let active = s.active_sheet;
                let layout_bounds = s.layouts.get(active).map(|l| {
                    (
                        l.total_width(),
                        l.total_height(),
                        l.frozen_cols_width(),
                        l.frozen_rows_height(),
                    )
                });

                let mut applied_dx = 0.0;
                let mut applied_dy = 0.0;
                if let Some((dx, dy)) = Self::scroll_state(&mut s, delta_x, delta_y) {
                    applied_dx = dx;
                    applied_dy = dy;
                    s.needs_render = true;
                }

                let (total_w, total_h, frozen_w, frozen_h) =
                    layout_bounds.unwrap_or((0.0, 0.0, 0.0, 0.0));
                if frozen_w > 0.0 || frozen_h > 0.0 {
                    let target_left = (s.viewport.scroll_x - frozen_w).max(0.0);
                    let target_top = (s.viewport.scroll_y - frozen_h).max(0.0);
                    container.set_scroll_left(target_left as i32);
                    container.set_scroll_top(target_top as i32);
                } else {
                    container.set_scroll_left(s.viewport.scroll_x as i32);
                    container.set_scroll_top(s.viewport.scroll_y as i32);
                }

                let viewport_w = s.viewport.width;
                let viewport_h = s.viewport.height;
                let scrollable_w = total_w - frozen_w;
                let scrollable_h = total_h - frozen_h;
                let viewport_content_w = viewport_w - frozen_w;
                let viewport_content_h = viewport_h - frozen_h;
                let max_x = frozen_w + (scrollable_w - viewport_content_w).max(0.0);
                let max_y = frozen_h + (scrollable_h - viewport_content_h).max(0.0);

                let metrics = ScrollMetrics {
                    requested_dx: delta_x,
                    requested_dy: delta_y,
                    applied_dx,
                    applied_dy,
                    scroll_x: s.viewport.scroll_x,
                    scroll_y: s.viewport.scroll_y,
                    max_x,
                    max_y,
                    total_width: total_w,
                    total_height: total_h,
                    viewport_width: viewport_w,
                    viewport_height: viewport_h,
                };
                let should_schedule =
                    applied_dx.abs() > f32::EPSILON || applied_dy.abs() > f32::EPSILON;
                (s.render_callback.clone(), should_schedule, metrics)
            })();

            if should_schedule {
                Self::schedule_scroll_settle_timeout(&self.state);
            }
            Self::invoke_render_callback(callback);
            return serde_wasm_bindgen::to_value(&metrics)
                .map_err(|e| JsValue::from_str(&format!("Serialization error: {e}")));
        }

        let (callback, should_schedule, metrics) = (|| {
            let mut s = self.state.borrow_mut();
            let active = s.active_sheet;
            let layout_bounds = s.layouts.get(active).map(|l| {
                (
                    l.total_width(),
                    l.total_height(),
                    l.frozen_cols_width(),
                    l.frozen_rows_height(),
                )
            });

            let mut applied_dx = 0.0;
            let mut applied_dy = 0.0;
            if let Some((dx, dy)) = Self::scroll_state(&mut s, delta_x, delta_y) {
                applied_dx = dx;
                applied_dy = dy;
                s.needs_render = true;
            }

            let (total_w, total_h, frozen_w, frozen_h) =
                layout_bounds.unwrap_or((0.0, 0.0, 0.0, 0.0));
            let viewport_w = s.viewport.width;
            let viewport_h = s.viewport.height;
            let scrollable_w = total_w - frozen_w;
            let scrollable_h = total_h - frozen_h;
            let viewport_content_w = viewport_w - frozen_w;
            let viewport_content_h = viewport_h - frozen_h;
            let max_x = frozen_w + (scrollable_w - viewport_content_w).max(0.0);
            let max_y = frozen_h + (scrollable_h - viewport_content_h).max(0.0);

            let metrics = ScrollMetrics {
                requested_dx: delta_x,
                requested_dy: delta_y,
                applied_dx,
                applied_dy,
                scroll_x: s.viewport.scroll_x,
                scroll_y: s.viewport.scroll_y,
                max_x,
                max_y,
                total_width: total_w,
                total_height: total_h,
                viewport_width: viewport_w,
                viewport_height: viewport_h,
            };
            let should_schedule =
                applied_dx.abs() > f32::EPSILON || applied_dy.abs() > f32::EPSILON;
            (s.render_callback.clone(), should_schedule, metrics)
        })();

        if should_schedule {
            Self::schedule_scroll_settle_timeout(&self.state);
        }
        Self::invoke_render_callback(callback);
        serde_wasm_bindgen::to_value(&metrics)
            .map_err(|e| JsValue::from_str(&format!("Serialization error: {e}")))
    }

    /// Set absolute scroll position
    #[wasm_bindgen]
    pub fn set_scroll(&mut self, x: f32, y: f32) {
        if let Some(container) = self.scroll_container.clone() {
            let (callback, should_schedule) = (|| {
                let mut s = self.state.borrow_mut();
                let active = s.active_sheet;
                let layout_bounds = s.layouts.get(active).map(|l| {
                    (
                        l.total_width(),
                        l.total_height(),
                        l.frozen_cols_width(),
                        l.frozen_rows_height(),
                    )
                });
                let Some((total_w, total_h, frozen_w, frozen_h)) = layout_bounds else {
                    return (None, false);
                };

                let min_x = frozen_w;
                let min_y = frozen_h;
                let scrollable_w = total_w - frozen_w;
                let scrollable_h = total_h - frozen_h;
                let viewport_content_w = s.viewport.width - frozen_w;
                let viewport_content_h = s.viewport.height - frozen_h;
                let max_x = frozen_w + (scrollable_w - viewport_content_w).max(0.0);
                let max_y = frozen_h + (scrollable_h - viewport_content_h).max(0.0);

                let new_x = x.clamp(min_x, max_x);
                let new_y = y.clamp(min_y, max_y);
                let target_left = (new_x - frozen_w).max(0.0);
                let target_top = (new_y - frozen_h).max(0.0);
                container.set_scroll_left(target_left as i32);
                container.set_scroll_top(target_top as i32);

                let dx = new_x - s.viewport.scroll_x;
                let dy = new_y - s.viewport.scroll_y;
                if dx.abs() > f32::EPSILON || dy.abs() > f32::EPSILON {
                    s.viewport.scroll_x = new_x;
                    s.viewport.scroll_y = new_y;
                    s.needs_render = true;
                    return (s.render_callback.clone(), true);
                }
                (None, false)
            })();

            if should_schedule {
                Self::schedule_scroll_settle_timeout(&self.state);
            }
            Self::invoke_render_callback(callback);
            return;
        }

        let (callback, should_schedule) = (|| {
            let mut s = self.state.borrow_mut();
            let active = s.active_sheet;
            let layout_bounds = s.layouts.get(active).map(|l| {
                (
                    l.total_width(),
                    l.total_height(),
                    l.frozen_cols_width(),
                    l.frozen_rows_height(),
                )
            });
            if let Some((total_w, total_h, frozen_w, frozen_h)) = layout_bounds {
                // Minimum scroll is at the frozen boundary
                let min_x = frozen_w;
                let min_y = frozen_h;
                // Maximum scroll allows viewing the end of the content
                let scrollable_w = total_w - frozen_w;
                let scrollable_h = total_h - frozen_h;
                let viewport_content_w = s.viewport.width - frozen_w;
                let viewport_content_h = s.viewport.height - frozen_h;
                let max_x = frozen_w + (scrollable_w - viewport_content_w).max(0.0);
                let max_y = frozen_h + (scrollable_h - viewport_content_h).max(0.0);

                let new_x = x.clamp(min_x, max_x);
                let new_y = y.clamp(min_y, max_y);
                let dx = new_x - s.viewport.scroll_x;
                let dy = new_y - s.viewport.scroll_y;
                if dx.abs() > f32::EPSILON || dy.abs() > f32::EPSILON {
                    s.viewport.scroll_x = new_x;
                    s.viewport.scroll_y = new_y;
                    s.needs_render = true;
                    return (s.render_callback.clone(), true);
                }
            }
            (None, false)
        })();
        if should_schedule {
            Self::schedule_scroll_settle_timeout(&self.state);
        }
        Self::invoke_render_callback(callback);
    }

    /// Switch to a different sheet
    #[wasm_bindgen]
    pub fn set_active_sheet(&mut self, index: usize) {
        let callback = (|| {
            let mut s = self.state.borrow_mut();
            if Self::set_active_sheet_state(&mut s, index) {
                s.needs_render = true;
                return s.render_callback.clone();
            }
            None
        })();

        // Update scroll spacer for new sheet
        self.update_scroll_spacer();

        Self::invoke_render_callback(callback);
    }

    /// Handle click event (for tab switching and hyperlinks)
    #[wasm_bindgen]
    pub fn on_click(&mut self, x: f32, y: f32) -> Option<usize> {
        let (selected_sheet, url, callback) = {
            let mut s = self.state.borrow_mut();
            Self::handle_click_state(&mut s, x, y)
        };
        if let Some(url) = url {
            Self::open_url(&url);
        }
        Self::invoke_render_callback(callback);
        selected_sheet
    }

    /// Handle mouse down (public API, delegates to internal handler)
    #[wasm_bindgen]
    pub fn on_mouse_down(&mut self, x: f32, y: f32) {
        Self::internal_mouse_down(&self.state, x, y);
    }

    /// Handle mouse move (public API, delegates to internal handler)
    #[wasm_bindgen]
    pub fn on_mouse_move(&mut self, x: f32, y: f32) {
        Self::internal_mouse_move(&self.state, x, y);
    }

    /// Handle mouse up (public API, delegates to internal handler)
    #[wasm_bindgen]
    pub fn on_mouse_up(&mut self, _x: f32, _y: f32) {
        Self::internal_mouse_up(&self.state);
    }

    /// Get selection as [min_row, min_col, max_row, max_col]
    #[wasm_bindgen]
    pub fn get_selection(&self) -> Option<Vec<u32>> {
        let s = self.state.borrow();
        let start = s.selection_start?;
        let end = s.selection_end?;
        Some(vec![
            start.0.min(end.0),
            start.1.min(end.1),
            start.0.max(end.0),
            start.1.max(end.1),
        ])
    }

    /// Programmatically set the selection to a single cell.
    #[wasm_bindgen]
    pub fn set_selection(&mut self, row: u32, col: u32) {
        let callback = {
            let mut s = self.state.borrow_mut();
            s.selection_start = Some((row, col));
            s.selection_end = Some((row, col));
            s.needs_render = true;
            s.render_callback.clone()
        };
        Self::invoke_render_callback(callback);
    }

    /// Get the cell at a given viewport point as `[row, col]`.
    #[wasm_bindgen]
    pub fn cell_at_point(&self, x: f32, y: f32) -> Option<Vec<u32>> {
        let s = self.state.borrow();
        match Self::hit_test(&s, x, y) {
            HitTarget::Cell(row, col) => Some(vec![row, col]),
            _ => None,
        }
    }

    /// Get the on-screen rectangle for a cell as `[x, y, w, h]` in logical pixels.
    ///
    /// Returns `None` if no layout is loaded.
    #[wasm_bindgen]
    pub fn cell_rect(&self, row: u32, col: u32) -> Option<Vec<f32>> {
        let s = self.state.borrow();
        let layout = s.layouts.get(s.active_sheet)?;
        let rect = layout.cell_rect(row, col);
        if rect.skip {
            return None;
        }
        // Convert from sheet coordinates to viewport coordinates
        let header_w = if s.show_headers {
            s.header_config.row_header_width
        } else {
            0.0
        };
        let header_h = if s.show_headers {
            s.header_config.col_header_height
        } else {
            0.0
        };
        let vp_x = rect.x - s.viewport.scroll_x + header_w;
        let vp_y = rect.y - s.viewport.scroll_y + header_h;
        Some(vec![vp_x, vp_y, rect.width, rect.height])
    }

    /// Get the display value of a cell.
    #[wasm_bindgen]
    pub fn cell_value(&self, row: u32, col: u32) -> Option<String> {
        let s = self.state.borrow();
        let workbook = s.workbook.as_ref()?;
        let sheet = workbook.sheets.get(s.active_sheet)?;
        let idx = sheet.cell_index_at(row, col)?;
        let cd = sheet.cells.get(idx)?;
        cd.cell
            .cached_display
            .as_ref()
            .or(cd.cell.v.as_ref())
            .cloned()
    }

    /// Access shared state (for editor integration).
    #[cfg(feature = "editing")]
    pub(crate) fn shared_state(&self) -> &Rc<RefCell<SharedState>> {
        &self.state
    }

    /// Handle keyboard event (public API)
    #[wasm_bindgen]
    pub fn on_key_down(&mut self, key: &str, ctrl: bool, _shift: bool) -> bool {
        Self::internal_key_down(&self.state, key, ctrl)
    }

    /// Check if position is over a hyperlink (for cursor styling)
    /// Returns true if the position is over a clickable hyperlink
    #[wasm_bindgen]
    pub fn is_over_hyperlink(&self, x: f32, y: f32) -> bool {
        let s = self.state.borrow();
        let tab_y = s.viewport.height;

        // Only check cells area, not tab bar
        if y >= tab_y {
            return false;
        }

        let Some(workbook) = &s.workbook else {
            return false;
        };
        let Some(layout) = s.layouts.get(s.active_sheet) else {
            return false;
        };
        let Some(sheet) = workbook.sheets.get(s.active_sheet) else {
            return false;
        };

        let sheet_x = x + s.viewport.scroll_x;
        let sheet_y = y + s.viewport.scroll_y;

        let Some(col) = layout.col_at_x(sheet_x) else {
            return false;
        };
        let Some(row) = layout.row_at_y(sheet_y) else {
            return false;
        };

        let Some(cell_idx) = sheet.cell_index_at(row, col) else {
            return false;
        };
        let cell_data = &sheet.cells[cell_idx];
        cell_data
            .cell
            .hyperlink
            .as_ref()
            .is_some_and(|h| h.is_external)
    }

    /// Get the comment text at a position (for tooltip display)
    /// Returns None if no comment at position, or the comment text with optional author
    #[wasm_bindgen]
    pub fn get_comment_at(&self, x: f32, y: f32) -> Option<String> {
        let s = self.state.borrow();
        let tab_y = s.viewport.height;

        // Only check cells area, not tab bar
        if y >= tab_y {
            return None;
        }

        let workbook = s.workbook.as_ref()?;
        let layout = s.layouts.get(s.active_sheet)?;
        let sheet = workbook.sheets.get(s.active_sheet)?;

        let sheet_x = x + s.viewport.scroll_x;
        let sheet_y = y + s.viewport.scroll_y;

        let col = layout.col_at_x(sheet_x)?;
        let row = layout.row_at_y(sheet_y)?;

        // Check if this cell has a comment
        let cell_idx = sheet.cell_index_at(row, col)?;
        let cell_data = &sheet.cells[cell_idx];
        if cell_data.cell.has_comment != Some(true) {
            return None;
        }

        // Find the comment by cell reference
        let cell_ref = format!("{}{}", col_to_letter(col), row + 1);
        let comment_idx = sheet.comments_by_cell.get(&cell_ref)?;
        let comment = sheet.comments.get(*comment_idx)?;

        // Format: "Author: text" or just "text"
        if let Some(author) = &comment.author {
            Some(format!("{}: {}", author, comment.text))
        } else {
            Some(comment.text.clone())
        }
    }

    /// Render the current state to the canvas
    #[wasm_bindgen]
    pub fn render(&mut self) -> Result<(), JsValue> {
        // Sync scroll position from native scroll container (if present)
        self.sync_scroll_from_container();

        let mut s = self.state.borrow_mut();
        let has_overlay = self.overlay_renderer.is_some();
        let mut needs_base = s.needs_render;
        if s.needs_overlay_render && !has_overlay {
            needs_base = true;
        }
        if !needs_base && s.needs_overlay_render && s.visible_cells.is_empty() {
            needs_base = true;
        }
        let needs_overlay = s.needs_overlay_render || needs_base;
        if !needs_base && !needs_overlay {
            return Ok(());
        }

        let active_sheet = s.active_sheet;
        let selection = match (s.selection_start, s.selection_end) {
            (Some((sr, sc)), Some((er, ec))) => Some((sr, sc, er, ec)),
            _ => None,
        };
        let viewport = s.viewport.clone();
        let policy_now = now_ms();
        let scrolling_active = Self::is_scroll_active(s.last_scroll_ms, policy_now);
        let (tile_prefetch, max_prefetch_tile_renders) = Self::tile_render_policy(
            scrolling_active,
            s.prefetch_warmup_frames,
            s.last_base_draw_ms,
        );

        let Some(layout) = s.layouts.get(active_sheet).cloned() else {
            return Ok(());
        };

        if needs_base {
            let (row_ranges, col_ranges) =
                Self::visible_cell_ranges(&viewport, &layout, tile_prefetch);
            let reuse_visible_cells =
                Self::can_reuse_visible_cells(&s, active_sheet, &row_ranges, &col_ranges);

            if !reuse_visible_cells {
                let (workbook, visible_cells) = {
                    let s = &mut *s;
                    (s.workbook.as_mut(), &mut s.visible_cells)
                };
                let Some(workbook) = workbook else {
                    return Ok(());
                };
                let sheet = &mut workbook.sheets[active_sheet];
                Self::get_visible_cell_data(
                    sheet,
                    &row_ranges,
                    &col_ranges,
                    &workbook.shared_strings,
                    &workbook.numfmt_cache,
                    workbook.date1904,
                    visible_cells,
                );
            }

            s.last_visible_sheet = Some(active_sheet);
            s.last_visible_row_ranges = row_ranges;
            s.last_visible_col_ranges = col_ranges;
        }

        let Some(workbook) = s.workbook.as_ref() else {
            return Ok(());
        };
        let sheet = &workbook.sheets[active_sheet];
        let draw_start = now_ms();

        if let Some(overlay_renderer) = self.overlay_renderer.as_mut() {
            if needs_base {
                let base_params = RenderParams {
                    cells: &s.visible_cells,
                    layout: &layout,
                    viewport: &viewport,
                    style_cache: &s.render_styles,
                    default_style: &s.default_render_style,
                    sheet_names: &s.sheet_names,
                    tab_colors: &s.tab_colors,
                    active_sheet,
                    dpr: s.dpr,
                    selection,
                    drawings: &sheet.drawings,
                    images: &workbook.images,
                    charts: &sheet.charts,
                    data_validations: &sheet.data_validations,
                    conditional_formatting: &sheet.conditional_formatting,
                    conditional_formatting_cache: &sheet.conditional_formatting_cache,
                    sparkline_groups: &sheet.sparkline_groups,
                    major_font: workbook.theme.major_font.as_deref(),
                    minor_font: workbook.theme.minor_font.as_deref(),
                    dxf_styles: &workbook.dxf_styles,
                    auto_filter: sheet.auto_filter.as_ref(),
                    show_headers: s.show_headers,
                    header_config: &s.header_config,
                    header_selection: s.selection.as_ref(),
                    show_tab_bar: false,
                    tile_prefetch,
                    max_prefetch_tile_renders,
                };
                self.renderer
                    .render_base(&base_params)
                    .map_err(|e| JsValue::from_str(&e.to_string()))?;
            }
            if needs_overlay {
                let overlay_params = RenderParams {
                    cells: &s.visible_cells,
                    layout: &layout,
                    viewport: &viewport,
                    style_cache: &s.render_styles,
                    default_style: &s.default_render_style,
                    sheet_names: &s.sheet_names,
                    tab_colors: &s.tab_colors,
                    active_sheet,
                    dpr: s.dpr,
                    selection,
                    drawings: &sheet.drawings,
                    images: &workbook.images,
                    charts: &sheet.charts,
                    data_validations: &sheet.data_validations,
                    conditional_formatting: &sheet.conditional_formatting,
                    conditional_formatting_cache: &sheet.conditional_formatting_cache,
                    sparkline_groups: &sheet.sparkline_groups,
                    major_font: workbook.theme.major_font.as_deref(),
                    minor_font: workbook.theme.minor_font.as_deref(),
                    dxf_styles: &workbook.dxf_styles,
                    auto_filter: sheet.auto_filter.as_ref(),
                    show_headers: s.show_headers,
                    header_config: &s.header_config,
                    header_selection: s.selection.as_ref(),
                    show_tab_bar: false,
                    tile_prefetch,
                    max_prefetch_tile_renders,
                };
                overlay_renderer
                    .render_overlay(&overlay_params)
                    .map_err(|e| JsValue::from_str(&e.to_string()))?;
            }
        } else if needs_base || needs_overlay {
            let params = RenderParams {
                cells: &s.visible_cells,
                layout: &layout,
                viewport: &viewport,
                style_cache: &s.render_styles,
                default_style: &s.default_render_style,
                sheet_names: &s.sheet_names,
                tab_colors: &s.tab_colors,
                active_sheet,
                dpr: s.dpr,
                selection,
                drawings: &sheet.drawings,
                images: &workbook.images,
                charts: &sheet.charts,
                data_validations: &sheet.data_validations,
                conditional_formatting: &sheet.conditional_formatting,
                conditional_formatting_cache: &sheet.conditional_formatting_cache,
                sparkline_groups: &sheet.sparkline_groups,
                major_font: workbook.theme.major_font.as_deref(),
                minor_font: workbook.theme.minor_font.as_deref(),
                dxf_styles: &workbook.dxf_styles,
                auto_filter: sheet.auto_filter.as_ref(),
                show_headers: s.show_headers,
                header_config: &s.header_config,
                header_selection: s.selection.as_ref(),
                show_tab_bar: false,
                tile_prefetch,
                max_prefetch_tile_renders,
            };
            self.renderer
                .render(&params)
                .map_err(|e| JsValue::from_str(&e.to_string()))?;
        }
        let draw_ms = now_ms() - draw_start;
        if needs_base {
            s.last_base_draw_ms = draw_ms;
            if !scrolling_active && s.prefetch_warmup_frames > 0 {
                s.prefetch_warmup_frames = s.prefetch_warmup_frames.saturating_sub(1);
            }
        }
        let needs_prefetch_catchup = needs_base && self.renderer.has_deferred_prefetch_tiles();
        let callback = if needs_prefetch_catchup {
            s.needs_render = true;
            s.needs_overlay_render = false;
            s.render_callback.clone()
        } else {
            s.needs_render = false;
            s.needs_overlay_render = false;
            None
        };
        drop(s);
        Self::invoke_render_callback(callback);
        Ok(())
    }

    /// Render the current state and return internal timing metrics.
    #[wasm_bindgen]
    pub fn render_with_metrics(&mut self) -> Result<JsValue, JsValue> {
        let total_start = now_ms();
        self.sync_scroll_from_container();
        let mut s = self.state.borrow_mut();
        let has_overlay = self.overlay_renderer.is_some();
        let mut needs_base = s.needs_render;
        if s.needs_overlay_render && !has_overlay {
            needs_base = true;
        }
        if !needs_base && s.needs_overlay_render && s.visible_cells.is_empty() {
            needs_base = true;
        }
        let needs_overlay = s.needs_overlay_render || needs_base;
        if !needs_base && !needs_overlay {
            let metrics = RenderMetrics {
                prep_ms: 0.0,
                draw_ms: 0.0,
                total_ms: now_ms() - total_start,
                visible_cells: 0,
                skipped: true,
            };
            return serde_wasm_bindgen::to_value(&metrics)
                .map_err(|e| JsValue::from_str(&format!("Serialization error: {e}")));
        }

        let prep_start = now_ms();
        let active_sheet = s.active_sheet;
        let selection = match (s.selection_start, s.selection_end) {
            (Some((sr, sc)), Some((er, ec))) => Some((sr, sc, er, ec)),
            _ => None,
        };
        let viewport = s.viewport.clone();
        let scrolling_active = Self::is_scroll_active(s.last_scroll_ms, prep_start);
        let (tile_prefetch, max_prefetch_tile_renders) = Self::tile_render_policy(
            scrolling_active,
            s.prefetch_warmup_frames,
            s.last_base_draw_ms,
        );

        let Some(layout) = s.layouts.get(active_sheet).cloned() else {
            let metrics = RenderMetrics {
                prep_ms: 0.0,
                draw_ms: 0.0,
                total_ms: now_ms() - total_start,
                visible_cells: 0,
                skipped: true,
            };
            return serde_wasm_bindgen::to_value(&metrics)
                .map_err(|e| JsValue::from_str(&format!("Serialization error: {e}")));
        };

        let mut visible_cells = s.visible_cells.len() as u32;
        let mut prep_ms = 0.0;
        if needs_base {
            let (row_ranges, col_ranges) =
                Self::visible_cell_ranges(&viewport, &layout, tile_prefetch);
            let reuse_visible_cells =
                Self::can_reuse_visible_cells(&s, active_sheet, &row_ranges, &col_ranges);

            if !reuse_visible_cells {
                let (workbook, visible_cells_out) = {
                    let s = &mut *s;
                    (s.workbook.as_mut(), &mut s.visible_cells)
                };
                let Some(workbook) = workbook else {
                    let metrics = RenderMetrics {
                        prep_ms: 0.0,
                        draw_ms: 0.0,
                        total_ms: now_ms() - total_start,
                        visible_cells: 0,
                        skipped: true,
                    };
                    return serde_wasm_bindgen::to_value(&metrics)
                        .map_err(|e| JsValue::from_str(&format!("Serialization error: {e}")));
                };

                let sheet = &mut workbook.sheets[active_sheet];
                Self::get_visible_cell_data(
                    sheet,
                    &row_ranges,
                    &col_ranges,
                    &workbook.shared_strings,
                    &workbook.numfmt_cache,
                    workbook.date1904,
                    visible_cells_out,
                );
            }

            s.last_visible_sheet = Some(active_sheet);
            s.last_visible_row_ranges = row_ranges;
            s.last_visible_col_ranges = col_ranges;
            visible_cells = s.visible_cells.len() as u32;
            prep_ms = now_ms() - prep_start;
        }

        let Some(workbook) = s.workbook.as_ref() else {
            let metrics = RenderMetrics {
                prep_ms,
                draw_ms: 0.0,
                total_ms: now_ms() - total_start,
                visible_cells,
                skipped: true,
            };
            return serde_wasm_bindgen::to_value(&metrics)
                .map_err(|e| JsValue::from_str(&format!("Serialization error: {e}")));
        };
        let sheet = &workbook.sheets[active_sheet];

        let draw_start = now_ms();
        if let Some(overlay_renderer) = self.overlay_renderer.as_mut() {
            if needs_base {
                let base_params = RenderParams {
                    cells: &s.visible_cells,
                    layout: &layout,
                    viewport: &viewport,
                    style_cache: &s.render_styles,
                    default_style: &s.default_render_style,
                    sheet_names: &s.sheet_names,
                    tab_colors: &s.tab_colors,
                    active_sheet,
                    dpr: s.dpr,
                    selection,
                    drawings: &sheet.drawings,
                    images: &workbook.images,
                    charts: &sheet.charts,
                    data_validations: &sheet.data_validations,
                    conditional_formatting: &sheet.conditional_formatting,
                    conditional_formatting_cache: &sheet.conditional_formatting_cache,
                    sparkline_groups: &sheet.sparkline_groups,
                    major_font: workbook.theme.major_font.as_deref(),
                    minor_font: workbook.theme.minor_font.as_deref(),
                    dxf_styles: &workbook.dxf_styles,
                    auto_filter: sheet.auto_filter.as_ref(),
                    show_headers: s.show_headers,
                    header_config: &s.header_config,
                    header_selection: s.selection.as_ref(),
                    show_tab_bar: false,
                    tile_prefetch,
                    max_prefetch_tile_renders,
                };
                self.renderer
                    .render_base(&base_params)
                    .map_err(|e| JsValue::from_str(&e.to_string()))?;
            }
            if needs_overlay {
                let overlay_params = RenderParams {
                    cells: &s.visible_cells,
                    layout: &layout,
                    viewport: &viewport,
                    style_cache: &s.render_styles,
                    default_style: &s.default_render_style,
                    sheet_names: &s.sheet_names,
                    tab_colors: &s.tab_colors,
                    active_sheet,
                    dpr: s.dpr,
                    selection,
                    drawings: &sheet.drawings,
                    images: &workbook.images,
                    charts: &sheet.charts,
                    data_validations: &sheet.data_validations,
                    conditional_formatting: &sheet.conditional_formatting,
                    conditional_formatting_cache: &sheet.conditional_formatting_cache,
                    sparkline_groups: &sheet.sparkline_groups,
                    major_font: workbook.theme.major_font.as_deref(),
                    minor_font: workbook.theme.minor_font.as_deref(),
                    dxf_styles: &workbook.dxf_styles,
                    auto_filter: sheet.auto_filter.as_ref(),
                    show_headers: s.show_headers,
                    header_config: &s.header_config,
                    header_selection: s.selection.as_ref(),
                    show_tab_bar: false,
                    tile_prefetch,
                    max_prefetch_tile_renders,
                };
                overlay_renderer
                    .render_overlay(&overlay_params)
                    .map_err(|e| JsValue::from_str(&e.to_string()))?;
            }
        } else if needs_base || needs_overlay {
            let params = RenderParams {
                cells: &s.visible_cells,
                layout: &layout,
                viewport: &viewport,
                style_cache: &s.render_styles,
                default_style: &s.default_render_style,
                sheet_names: &s.sheet_names,
                tab_colors: &s.tab_colors,
                active_sheet,
                dpr: s.dpr,
                selection,
                drawings: &sheet.drawings,
                images: &workbook.images,
                charts: &sheet.charts,
                data_validations: &sheet.data_validations,
                conditional_formatting: &sheet.conditional_formatting,
                conditional_formatting_cache: &sheet.conditional_formatting_cache,
                sparkline_groups: &sheet.sparkline_groups,
                major_font: workbook.theme.major_font.as_deref(),
                minor_font: workbook.theme.minor_font.as_deref(),
                dxf_styles: &workbook.dxf_styles,
                auto_filter: sheet.auto_filter.as_ref(),
                show_headers: s.show_headers,
                header_config: &s.header_config,
                header_selection: s.selection.as_ref(),
                show_tab_bar: false,
                tile_prefetch,
                max_prefetch_tile_renders,
            };
            self.renderer
                .render(&params)
                .map_err(|e| JsValue::from_str(&e.to_string()))?;
        }
        let draw_ms = now_ms() - draw_start;
        let total_ms = now_ms() - total_start;
        if needs_base {
            s.last_base_draw_ms = draw_ms;
            if !scrolling_active && s.prefetch_warmup_frames > 0 {
                s.prefetch_warmup_frames = s.prefetch_warmup_frames.saturating_sub(1);
            }
        }

        let needs_prefetch_catchup = needs_base && self.renderer.has_deferred_prefetch_tiles();
        let callback = if needs_prefetch_catchup {
            s.needs_render = true;
            s.needs_overlay_render = false;
            s.render_callback.clone()
        } else {
            s.needs_render = false;
            s.needs_overlay_render = false;
            None
        };

        let metrics = RenderMetrics {
            prep_ms,
            draw_ms,
            total_ms,
            visible_cells,
            skipped: false,
        };
        let result = serde_wasm_bindgen::to_value(&metrics)
            .map_err(|e| JsValue::from_str(&format!("Serialization error: {e}")));
        drop(s);
        Self::invoke_render_callback(callback);
        result
    }

    /// Force a re-render
    #[wasm_bindgen]
    pub fn invalidate(&mut self) {
        let callback = {
            let mut s = self.state.borrow_mut();
            s.needs_render = true;
            s.render_callback.clone()
        };
        Self::invoke_render_callback(callback);
    }

    /// Invalidate all caches and re-render (call after data mutations).
    #[cfg(feature = "editing")]
    pub(crate) fn invalidate_data(&mut self) {
        self.renderer.clear_tile_cache();
        {
            let mut s = self.state.borrow_mut();
            s.visible_cells.clear();
            s.last_visible_row_ranges.clear();
            s.last_visible_col_ranges.clear();
            s.last_visible_sheet = None;
            s.needs_render = true;
        }
        let callback = { self.state.borrow().render_callback.clone() };
        Self::invoke_render_callback(callback);
    }

    /// Set whether row/column headers are visible
    #[wasm_bindgen]
    pub fn set_headers_visible(&mut self, visible: bool) {
        let callback = {
            let mut s = self.state.borrow_mut();
            if s.show_headers != visible {
                s.show_headers = visible;
                s.header_config.visible = visible;
                // Update layout header dimensions
                let header_width = if visible {
                    s.header_config.row_header_width
                } else {
                    0.0
                };
                let header_height = if visible {
                    s.header_config.col_header_height
                } else {
                    0.0
                };
                for layout in &mut s.layouts {
                    if let Some(layout) = Arc::get_mut(layout) {
                        layout.set_header_dimensions(header_width, header_height);
                    }
                }
                s.needs_render = true;
                s.render_callback.clone()
            } else {
                None
            }
        };
        Self::invoke_render_callback(callback);
    }

    /// Check if row/column headers are visible
    #[wasm_bindgen]
    pub fn headers_visible(&self) -> bool {
        self.state.borrow().show_headers
    }

    /// Get header configuration for debugging
    #[wasm_bindgen]
    pub fn get_header_config(&self) -> JsValue {
        let s = self.state.borrow();
        let obj = js_sys::Object::new();
        let _ = Reflect::set(&obj, &"visible".into(), &s.header_config.visible.into());
        let _ = Reflect::set(
            &obj,
            &"row_header_width".into(),
            &s.header_config.row_header_width.into(),
        );
        let _ = Reflect::set(
            &obj,
            &"col_header_height".into(),
            &s.header_config.col_header_height.into(),
        );
        obj.into()
    }

    /// Register a JS callback to request a render on the next animation frame.
    #[wasm_bindgen]
    pub fn set_render_callback(&mut self, callback: Option<Function>) {
        self.state.borrow_mut().render_callback = callback;
    }

    /// Get number of sheets
    #[wasm_bindgen]
    pub fn sheet_count(&self) -> usize {
        self.state
            .borrow()
            .workbook
            .as_ref()
            .map(|w| w.sheets.len())
            .unwrap_or(0)
    }

    /// Get sheet name by index
    #[wasm_bindgen]
    pub fn sheet_name(&self, index: usize) -> Option<String> {
        self.state
            .borrow()
            .workbook
            .as_ref()
            .and_then(|w| w.sheets.get(index))
            .map(|s| s.name.clone())
    }

    /// Get current active sheet index
    #[wasm_bindgen]
    pub fn active_sheet(&self) -> usize {
        self.state.borrow().active_sheet
    }

    /// Get all sheet names.
    #[wasm_bindgen]
    pub fn sheet_names(&self) -> Vec<String> {
        self.state.borrow().sheet_names.clone()
    }

    /// Get a reference to the scroll container element (for editor positioning).
    #[cfg(feature = "editing")]
    pub(crate) fn scroll_container_element(&self) -> Option<HtmlElement> {
        self.scroll_container.as_ref().map(|c| c.clone().into())
    }

    /// Get total content width
    #[wasm_bindgen]
    pub fn content_width(&self) -> f32 {
        let s = self.state.borrow();
        s.layouts
            .get(s.active_sheet)
            .map(|l| l.total_width())
            .unwrap_or(0.0)
    }

    /// Get total content height
    #[wasm_bindgen]
    pub fn content_height(&self) -> f32 {
        let s = self.state.borrow();
        s.layouts
            .get(s.active_sheet)
            .map(|l| l.total_height())
            .unwrap_or(0.0)
    }

    fn compute_layout(sheet: &crate::types::Sheet) -> SheetLayout {
        let mut col_widths_map: HashMap<u32, f32> = HashMap::new();
        for cw in &sheet.col_widths {
            // f64 to f32: pixel values are small, truncation is safe
            let width_px = (cw.width * 7.0).min(f64::from(f32::MAX)) as f32;
            col_widths_map.insert(cw.col, width_px);
        }
        let mut row_heights_map: HashMap<u32, f32> = HashMap::new();
        for rh in &sheet.row_heights {
            // f64 to f32: pixel values are small, truncation is safe
            let height_px = (rh.height * 1.33).min(f64::from(f32::MAX)) as f32;
            row_heights_map.insert(rh.row, height_px);
        }
        let hidden_cols: HashSet<u32> = sheet.hidden_cols.iter().copied().collect();
        let hidden_rows: HashSet<u32> = sheet.hidden_rows.iter().copied().collect();
        let merge_ranges: Vec<(u32, u32, u32, u32)> = sheet
            .merges
            .iter()
            .map(|m| (m.start_row, m.start_col, m.end_row, m.end_col))
            .collect();
        let mut max_row = sheet.cells.iter().map(|c| c.r).max().unwrap_or(0);
        let mut max_col = sheet.cells.iter().map(|c| c.c).max().unwrap_or(0);
        if sheet.max_row > 0 {
            max_row = max_row.max(sheet.max_row.saturating_sub(1));
        }
        if sheet.max_col > 0 {
            max_col = max_col.max(sheet.max_col.saturating_sub(1));
        }
        if let Some(rh_max) = sheet.row_heights.iter().map(|rh| rh.row).max() {
            max_row = max_row.max(rh_max);
        }
        if let Some(cw_max) = sheet.col_widths.iter().map(|cw| cw.col).max() {
            max_col = max_col.max(cw_max);
        }
        if let Some(hidden_row) = sheet.hidden_rows.iter().copied().max() {
            max_row = max_row.max(hidden_row);
        }
        if let Some(hidden_col) = sheet.hidden_cols.iter().copied().max() {
            max_col = max_col.max(hidden_col);
        }
        if let Some(merge_row) = sheet.merges.iter().map(|m| m.end_row).max() {
            max_row = max_row.max(merge_row);
        }
        if let Some(merge_col) = sheet.merges.iter().map(|m| m.end_col).max() {
            max_col = max_col.max(merge_col);
        }
        let max_row = max_row.max(100);
        let max_col = max_col.max(26);
        SheetLayout::new(
            max_row,
            max_col,
            &col_widths_map,
            &row_heights_map,
            &hidden_cols,
            &hidden_rows,
            &merge_ranges,
            sheet.frozen_rows,
            sheet.frozen_cols,
        )
    }

    fn get_visible_cell_data(
        sheet: &mut crate::types::Sheet,
        row_ranges: &[(u32, u32)],
        col_ranges: &[(u32, u32)],
        shared_strings: &[String],
        numfmt_cache: &[CompiledFormat],
        date1904: bool,
        cells_out: &mut Vec<CellRenderData>,
    ) {
        cells_out.clear();
        if sheet.cells_by_row.is_empty() {
            sheet.rebuild_cell_index();
        }

        if row_ranges.is_empty() || col_ranges.is_empty() {
            return;
        }

        for &(row_start, row_end) in row_ranges {
            if row_start > row_end {
                continue;
            }
            for row in row_start..=row_end {
                let Some(row_cells) = sheet.cells_by_row.get(row as usize) else {
                    continue;
                };
                if row_cells.is_empty() {
                    continue;
                }

                for &(start_col, end_col) in col_ranges {
                    if start_col > end_col {
                        continue;
                    }
                    let start_pos =
                        row_cells.partition_point(|&idx| sheet.cells[idx].c < start_col);
                    let end_pos = row_cells.partition_point(|&idx| sheet.cells[idx].c <= end_col);

                    for &cell_idx in &row_cells[start_pos..end_pos] {
                        let cell_data = &mut sheet.cells[cell_idx];
                        let value = Self::resolve_cell_display_value(
                            &mut cell_data.cell,
                            shared_strings,
                            numfmt_cache,
                            date1904,
                        );
                        let numeric_value = value.as_ref().and_then(|v| v.parse::<f64>().ok());
                        let style_override = cell_data
                            .cell
                            .s
                            .as_ref()
                            .map(|s| Self::style_ref_to_render(s));
                        if cell_data.cell.cached_rich_text.is_none() {
                            cell_data.cell.cached_rich_text =
                                cell_data.cell.rich_text.as_ref().map(|runs| {
                                    Rc::new(
                                        runs.iter()
                                            .map(|run| {
                                                let style = run.style.as_ref();
                                                TextRunData {
                                                    text: run.text.clone(),
                                                    bold: style.and_then(|s| s.bold),
                                                    italic: style.and_then(|s| s.italic),
                                                    font_size: style
                                                        .and_then(|s| s.font_size)
                                                        .map(|f| f as f32),
                                                    font_color: style
                                                        .and_then(|s| s.font_color.clone()),
                                                    font_family: style
                                                        .and_then(|s| s.font_family.clone()),
                                                    underline: style.and_then(|s| s.underline),
                                                    strikethrough: style
                                                        .and_then(|s| s.strikethrough),
                                                }
                                            })
                                            .collect(),
                                    )
                                });
                        }
                        let rich_text = cell_data.cell.cached_rich_text.clone();

                        cells_out.push(CellRenderData {
                            row,
                            col: cell_data.c,
                            value,
                            numeric_value,
                            style_idx: cell_data.cell.style_idx.map(|idx| idx as usize),
                            style_override,
                            has_hyperlink: cell_data.cell.hyperlink.is_some().then_some(true),
                            has_comment: cell_data.cell.has_comment,
                            rich_text,
                        });
                    }
                }
            }
        }
    }
}

#[allow(dead_code)]
impl XlView {
    fn prepare_workbook_caches(workbook: &mut Workbook) {
        for sheet in &mut workbook.sheets {
            sheet.rebuild_comment_index();

            sheet.conditional_formatting_cache = sheet
                .conditional_formatting
                .iter()
                .map(|cf| {
                    let ranges = parse_sqref(&cf.sqref);
                    let mut sorted_rule_indices: Vec<usize> = (0..cf.rules.len()).collect();
                    sorted_rule_indices.sort_by_key(|&i| {
                        cf.rules
                            .get(i)
                            .map(|rule| rule.priority)
                            .unwrap_or(u32::MAX)
                    });
                    crate::types::ConditionalFormattingCache {
                        ranges,
                        sorted_rule_indices,
                    }
                })
                .collect();
        }
    }

    #[allow(clippy::cast_possible_truncation)]
    fn style_ref_to_render(style: &StyleRef) -> CellStyleData {
        CellStyleData {
            bg_color: style.bg_color.clone(),
            font_color: style.font_color.clone(),
            font_size: style.font_size.map(|f| f as f32),
            font_family: style.font_family.clone(),
            bold: style.bold,
            italic: style.italic,
            underline: style.underline.is_some().then_some(true),
            strikethrough: style.strikethrough,
            rotation: style.rotation,
            indent: style.indent,
            align_h: style
                .align_h
                .as_ref()
                .map(|h| format!("{:?}", h).to_lowercase()),
            align_v: style
                .align_v
                .as_ref()
                .map(|v| format!("{:?}", v).to_lowercase()),
            wrap_text: style.wrap,
            border_top: style.border_top.as_ref().map(|b| BorderStyleData {
                style: Some(format!("{:?}", b.style).to_lowercase()),
                color: Some(b.color.clone()),
            }),
            border_right: style.border_right.as_ref().map(|b| BorderStyleData {
                style: Some(format!("{:?}", b.style).to_lowercase()),
                color: Some(b.color.clone()),
            }),
            border_bottom: style.border_bottom.as_ref().map(|b| BorderStyleData {
                style: Some(format!("{:?}", b.style).to_lowercase()),
                color: Some(b.color.clone()),
            }),
            border_left: style.border_left.as_ref().map(|b| BorderStyleData {
                style: Some(format!("{:?}", b.style).to_lowercase()),
                color: Some(b.color.clone()),
            }),
            border_diagonal_down: if style.diagonal_down == Some(true) {
                style.border_diagonal.as_ref().map(|b| BorderStyleData {
                    style: Some(format!("{:?}", b.style).to_lowercase()),
                    color: Some(b.color.clone()),
                })
            } else {
                None
            },
            border_diagonal_up: if style.diagonal_up == Some(true) {
                style.border_diagonal.as_ref().map(|b| BorderStyleData {
                    style: Some(format!("{:?}", b.style).to_lowercase()),
                    color: Some(b.color.clone()),
                })
            } else {
                None
            },
            pattern_type: style
                .pattern_type
                .as_ref()
                .map(|p| format!("{:?}", p).to_lowercase()),
            pattern_fg_color: style.fg_color.clone(),
            pattern_bg_color: style.bg_color.clone(),
        }
    }
}

// ============================================================================
// Non-WASM32 Implementation (for testing/CLI)
// ============================================================================

#[cfg(not(target_arch = "wasm32"))]
impl XlView {
    /// Create a new viewer (non-wasm version for testing)
    pub fn new_test(width: u32, height: u32, dpr: f32) -> Self {
        let logical_width = width as f32 / dpr;
        let logical_height = height as f32 / dpr;
        XlView {
            workbook: None,
            layouts: Vec::new(),
            viewport: Viewport {
                scroll_x: 0.0,
                scroll_y: 0.0,
                width: logical_width,
                height: logical_height.max(100.0),
                scale: 1.0,
                tab_scroll_x: 0.0,
            },
            active_sheet: 0,
            width,
            height,
            dpr,
            needs_render: true,
            needs_overlay_render: false,
            sheet_names: Vec::new(),
            tab_colors: Vec::new(),
            render_styles: Vec::new(),
            default_render_style: None,
            visible_cells: Vec::new(),
            selection_start: None,
            selection_end: None,
            is_selecting: false,
            is_resizing: false,
        }
    }

    /// Load an XLSX file from bytes.
    pub fn load(&mut self, data: &[u8]) -> crate::error::Result<()> {
        let xlsx_wb = offidized_xlsx::Workbook::from_bytes(data)?;
        let mut workbook = crate::adapter::convert_workbook(&xlsx_wb);
        let render_styles: Vec<Option<CellStyleData>> =
            crate::adapter::build_render_style_cache(&xlsx_wb);
        let default_render_style = xlsx_wb.style(0).map(|s| {
            crate::adapter::style_to_render(s, xlsx_wb.theme_colors(), xlsx_wb.indexed_colors())
        });
        let layouts: Vec<Arc<SheetLayout>> = workbook
            .sheets
            .iter()
            .map(Self::compute_layout)
            .map(Arc::new)
            .collect();
        let sheet_names: Vec<String> = workbook.sheets.iter().map(|sh| sh.name.clone()).collect();
        let tab_colors: Vec<Option<String>> = workbook
            .sheets
            .iter()
            .map(|sh| sh.tab_color.clone())
            .collect();
        Self::prepare_workbook_caches(&mut workbook);

        if let Some(layout) = layouts.first() {
            self.viewport.scroll_x = layout.frozen_cols_width();
            self.viewport.scroll_y = layout.frozen_rows_height();
        } else {
            self.viewport.scroll_x = 0.0;
            self.viewport.scroll_y = 0.0;
        }
        self.layouts = layouts;
        self.workbook = Some(workbook);
        self.sheet_names = sheet_names;
        self.tab_colors = tab_colors;
        self.render_styles = render_styles;
        self.default_render_style = default_render_style;
        self.active_sheet = 0;
        self.selection_start = None;
        self.selection_end = None;
        self.needs_render = true;
        Ok(())
    }

    /// Compute sheet layout from sheet data.
    /// The f64→f32 casts are safe: pixel values are clamped to f32::MAX before casting.
    #[allow(clippy::cast_possible_truncation, dead_code)]
    fn compute_layout(sheet: &crate::types::Sheet) -> SheetLayout {
        let mut col_widths_map: HashMap<u32, f32> = HashMap::new();
        for cw in &sheet.col_widths {
            let width_px = (cw.width * 7.0).min(f64::from(f32::MAX)) as f32;
            col_widths_map.insert(cw.col, width_px);
        }
        let mut row_heights_map: HashMap<u32, f32> = HashMap::new();
        for rh in &sheet.row_heights {
            let height_px = (rh.height * 1.33).min(f64::from(f32::MAX)) as f32;
            row_heights_map.insert(rh.row, height_px);
        }
        let hidden_cols: HashSet<u32> = sheet.hidden_cols.iter().copied().collect();
        let hidden_rows: HashSet<u32> = sheet.hidden_rows.iter().copied().collect();
        let merge_ranges: Vec<(u32, u32, u32, u32)> = sheet
            .merges
            .iter()
            .map(|m| (m.start_row, m.start_col, m.end_row, m.end_col))
            .collect();
        let mut max_row = sheet.cells.iter().map(|c| c.r).max().unwrap_or(0);
        let mut max_col = sheet.cells.iter().map(|c| c.c).max().unwrap_or(0);
        if sheet.max_row > 0 {
            max_row = max_row.max(sheet.max_row.saturating_sub(1));
        }
        if sheet.max_col > 0 {
            max_col = max_col.max(sheet.max_col.saturating_sub(1));
        }
        if let Some(rh_max) = sheet.row_heights.iter().map(|rh| rh.row).max() {
            max_row = max_row.max(rh_max);
        }
        if let Some(cw_max) = sheet.col_widths.iter().map(|cw| cw.col).max() {
            max_col = max_col.max(cw_max);
        }
        if let Some(hidden_row) = sheet.hidden_rows.iter().copied().max() {
            max_row = max_row.max(hidden_row);
        }
        if let Some(hidden_col) = sheet.hidden_cols.iter().copied().max() {
            max_col = max_col.max(hidden_col);
        }
        if let Some(merge_row) = sheet.merges.iter().map(|m| m.end_row).max() {
            max_row = max_row.max(merge_row);
        }
        if let Some(merge_col) = sheet.merges.iter().map(|m| m.end_col).max() {
            max_col = max_col.max(merge_col);
        }
        let max_row = max_row.max(100);
        let max_col = max_col.max(26);
        SheetLayout::new(
            max_row,
            max_col,
            &col_widths_map,
            &row_heights_map,
            &hidden_cols,
            &hidden_rows,
            &merge_ranges,
            sheet.frozen_rows,
            sheet.frozen_cols,
        )
    }

    pub fn scroll(&mut self, delta_x: f32, delta_y: f32) {
        if let Some(layout) = self.layouts.get(self.active_sheet) {
            self.viewport.scroll_by(delta_x, delta_y, layout);
            self.needs_render = true;
        }
    }

    pub fn set_active_sheet(&mut self, index: usize) {
        if let Some(workbook) = &self.workbook {
            if index < workbook.sheets.len() {
                self.active_sheet = index;
                // Initialize scroll to frozen boundary (for frozen panes support)
                if let Some(layout) = self.layouts.get(index) {
                    self.viewport.scroll_x = layout.frozen_cols_width();
                    self.viewport.scroll_y = layout.frozen_rows_height();
                } else {
                    self.viewport.scroll_x = 0.0;
                    self.viewport.scroll_y = 0.0;
                }
                self.selection_start = None;
                self.selection_end = None;
                self.needs_render = true;
            }
        }
    }

    pub fn sheet_count(&self) -> usize {
        self.workbook.as_ref().map(|w| w.sheets.len()).unwrap_or(0)
    }

    pub fn sheet_name(&self, index: usize) -> Option<String> {
        self.workbook
            .as_ref()
            .and_then(|w| w.sheets.get(index))
            .map(|s| s.name.clone())
    }

    pub fn active_sheet(&self) -> usize {
        self.active_sheet
    }

    pub fn sheet_names(&self) -> Vec<String> {
        self.workbook
            .as_ref()
            .map(|w| w.sheets.iter().map(|s| s.name.clone()).collect())
            .unwrap_or_default()
    }

    /// Programmatically set the selection to a single cell.
    pub fn set_selection(&mut self, row: u32, col: u32) {
        self.selection_start = Some((row, col));
        self.selection_end = Some((row, col));
        self.needs_render = true;
    }

    /// Mutable access to the workbook (for editor integration).
    #[cfg(feature = "editing")]
    pub(crate) fn workbook_mut(&mut self) -> Option<&mut Workbook> {
        self.workbook.as_mut()
    }

    /// Immutable access to the workbook (for editor integration).
    #[cfg(feature = "editing")]
    pub(crate) fn workbook_ref(&self) -> Option<&Workbook> {
        self.workbook.as_ref()
    }
}
