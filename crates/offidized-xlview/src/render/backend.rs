//! Render backend trait for pluggable rendering implementations.
//!
//! This module defines the `RenderBackend` trait that abstracts rendering
//! operations, allowing different backends (Canvas 2D, WebGPU/vello) to
//! be used interchangeably.

use crate::error::Result;
use crate::layout::{SheetLayout, Viewport};
use crate::types::chart::Chart;
use crate::types::content::DataValidationRange;
use crate::types::drawing::{Drawing, EmbeddedImage};
use crate::types::filter::AutoFilter;
use crate::types::formatting::ConditionalFormatting;
use crate::types::rich_text::TextRunData;
use crate::types::selection::{HeaderConfig, Selection};
use crate::types::sparkline::SparklineGroup;
use std::rc::Rc;

/// Data needed to render a single cell
#[derive(Debug, Clone)]
pub struct CellRenderData {
    pub row: u32,
    pub col: u32,
    pub value: Option<String>,
    pub numeric_value: Option<f64>,
    /// Style index into the render style cache.
    pub style_idx: Option<usize>,
    /// Inline style override (used when a cell has a resolved StyleRef).
    pub style_override: Option<CellStyleData>,
    pub has_hyperlink: Option<bool>,
    pub has_comment: Option<bool>,
    pub rich_text: Option<Rc<Vec<TextRunData>>>,
}

/// Style information for rendering a cell
#[derive(Debug, Clone, Default)]
pub struct CellStyleData {
    pub bg_color: Option<String>,
    pub font_color: Option<String>,
    pub font_size: Option<f32>,
    pub font_family: Option<String>,
    pub bold: Option<bool>,
    pub italic: Option<bool>,
    pub underline: Option<bool>,
    pub strikethrough: Option<bool>,
    pub rotation: Option<i32>,
    pub align_h: Option<String>,
    pub align_v: Option<String>,
    pub indent: Option<u32>,
    pub wrap_text: Option<bool>,
    pub border_top: Option<BorderStyleData>,
    pub border_right: Option<BorderStyleData>,
    pub border_bottom: Option<BorderStyleData>,
    pub border_left: Option<BorderStyleData>,
    pub border_diagonal_down: Option<BorderStyleData>,
    pub border_diagonal_up: Option<BorderStyleData>,
    pub pattern_type: Option<String>,
    pub pattern_fg_color: Option<String>,
    pub pattern_bg_color: Option<String>,
}

/// Border style information
#[derive(Debug, Clone)]
pub struct BorderStyleData {
    pub style: Option<String>,
    pub color: Option<String>,
}

impl BorderStyleData {
    /// Get border width based on style
    pub fn width(&self) -> f64 {
        match self.style.as_deref() {
            Some("thin") | Some("hair") => 1.0,
            Some("medium") => 2.0,
            Some("thick") => 3.0,
            Some("double") => 3.0,
            _ => 1.0,
        }
    }
}

/// Render parameters passed to the backend
pub struct RenderParams<'a> {
    pub cells: &'a [CellRenderData],
    pub layout: &'a SheetLayout,
    pub viewport: &'a Viewport,
    /// Render style cache indexed by style_idx.
    pub style_cache: &'a [Option<CellStyleData>],
    /// Default render style (for cells without explicit styles).
    pub default_style: &'a Option<CellStyleData>,
    pub sheet_names: &'a [String],
    pub tab_colors: &'a [Option<String>],
    pub active_sheet: usize,
    pub dpr: f32,
    /// Selection range: (start_row, start_col, end_row, end_col)
    pub selection: Option<(u32, u32, u32, u32)>,
    /// Drawings (images, charts, shapes) in the current sheet
    pub drawings: &'a [Drawing],
    /// Embedded images from the workbook (for resolving image_id references)
    pub images: &'a [EmbeddedImage],
    /// Charts in the current sheet
    pub charts: &'a [Chart],
    /// Data validations for the current sheet
    pub data_validations: &'a [DataValidationRange],
    /// Conditional formatting rules for the current sheet
    pub conditional_formatting: &'a [ConditionalFormatting],
    /// Preprocessed conditional formatting metadata for the current sheet
    pub conditional_formatting_cache: &'a [crate::types::ConditionalFormattingCache],
    /// Sparkline groups for the current sheet
    pub sparkline_groups: &'a [SparklineGroup],
    /// Theme major font (headings) - used as default for heading-style cells
    pub major_font: Option<&'a str>,
    /// Theme minor font (body) - used as default for body cells without explicit font
    pub minor_font: Option<&'a str>,
    /// DXF styles for conditional formatting (referenced by dxf_id in CFRule)
    pub dxf_styles: &'a [crate::types::DxfStyle],
    /// Auto-filter settings for the current sheet (for rendering filter dropdown buttons)
    pub auto_filter: Option<&'a AutoFilter>,
    /// Whether to show row/column headers
    pub show_headers: bool,
    /// Header configuration (colors, dimensions)
    pub header_config: &'a HeaderConfig,
    /// Current selection (for header highlighting)
    pub header_selection: Option<&'a Selection>,
    /// Whether to render the tab bar on canvas (false when using DOM tab bar)
    pub show_tab_bar: bool,
    /// Number of tile rings to prefetch outside the visible scrollable range.
    pub tile_prefetch: u32,
    /// Optional cap for new prefetch-only tile renders this frame.
    /// Visible tiles are always rendered even when this budget is exhausted.
    pub max_prefetch_tile_renders: Option<u32>,
}

/// Trait for render backends
///
/// Implementations handle the actual drawing operations for different
/// rendering technologies (Canvas 2D, WebGPU, etc.)
pub trait RenderBackend {
    /// Initialize the backend
    fn init(&mut self) -> Result<()>;

    /// Resize the render surface
    fn resize(&mut self, width: u32, height: u32, dpr: f32);

    /// Render a frame with the given parameters
    fn render(&mut self, params: &RenderParams<'_>) -> Result<()>;

    /// Load a font (for backends that need explicit font loading)
    fn load_font(&mut self, _font_data: &[u8]) -> Result<()> {
        // Default implementation does nothing - Canvas uses browser fonts
        Ok(())
    }

    /// Get the current width
    fn width(&self) -> u32;

    /// Get the current height
    fn height(&self) -> u32;
}
