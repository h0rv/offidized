//! Workbook/sheet types for the viewer data model.
//!
//! These types define the internal data model that the viewer's renderer
//! and layout engine operate on. The `adapter` module converts from
//! `offidized-xlsx` types into these viewer types.

#![allow(dead_code)]

use std::collections::HashMap;

use crate::types::chart::Chart;
use crate::types::content::DataValidationRange;
use crate::types::drawing::{Drawing, EmbeddedImage};
use crate::types::filter::AutoFilter;
use crate::types::formatting::{ConditionalFormatting, ConditionalFormattingCache, DxfStyle};
use crate::types::sparkline::SparklineGroup;
use crate::types::style::{ColWidth, MergeRange, RowHeight, StyleRef, Theme};
use crate::types::Hyperlink;

/// Cell type tag for OOXML serialization.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum CellType {
    /// String (inline or shared).
    String,
    /// Numeric value.
    Number,
    /// Boolean value.
    Boolean,
    /// Error value (e.g. "#REF!").
    Error,
    /// Date stored as serial number.
    Date,
}

/// Raw cell value variants.
#[derive(Debug, Clone)]
pub enum CellRawValue {
    /// Index into the shared strings table.
    SharedString(u32),
    /// Inline string value.
    String(String),
    /// Boolean value.
    Boolean(bool),
    /// Error value (e.g. "#REF!").
    Error(String),
    /// Numeric value.
    Number(f64),
    /// Date stored as a serial number.
    Date(f64),
}

/// A single cell's data.
#[derive(Debug, Clone)]
pub struct Cell {
    /// Display value (cached).
    pub v: Option<String>,
    /// Cached display string (formatted).
    pub cached_display: Option<String>,
    /// Raw value.
    pub raw: Option<CellRawValue>,
    /// Style index into the workbook's resolved styles.
    pub style_idx: Option<u32>,
    /// Hyperlink on this cell.
    pub hyperlink: Option<Hyperlink>,
    /// Whether this cell has a comment.
    pub has_comment: Option<bool>,
    /// Rich text runs (if any).
    pub rich_text: Option<Vec<crate::types::rich_text::RichTextRun>>,
    /// Inline resolved style (optional, used by renderer for override).
    pub s: Option<StyleRef>,
    /// Cached rich text data for rendering.
    pub cached_rich_text: Option<std::rc::Rc<Vec<crate::render::TextRunData>>>,
    /// Cell type tag (for serialization).
    pub t: CellType,
    /// Formula string (if any).
    pub formula: Option<String>,
}

/// Hyperlink definition with cell reference (for bulk storage on sheet).
#[derive(Debug, Clone)]
pub struct HyperlinkDef {
    /// Cell reference (e.g. "A1").
    pub cell_ref: String,
    /// The hyperlink data.
    pub hyperlink: Hyperlink,
}

/// A cell together with its row/column position.
#[derive(Debug, Clone)]
pub struct CellData {
    /// 0-based row index.
    pub r: u32,
    /// 0-based column index.
    pub c: u32,
    /// The cell itself.
    pub cell: Cell,
}

/// A comment on a cell.
#[derive(Debug, Clone)]
pub struct Comment {
    /// Comment text.
    pub text: String,
    /// Author (optional).
    pub author: Option<String>,
}

/// Re-export the real compiled number format from offidized-xlsx.
pub use offidized_xlsx::CompiledFormat;

/// A single worksheet.
#[derive(Debug)]
pub struct Sheet {
    /// Sheet name.
    pub name: String,
    /// Tab color (CSS color string).
    pub tab_color: Option<String>,
    /// All cell data.
    pub cells: Vec<CellData>,
    /// Column widths.
    pub col_widths: Vec<ColWidth>,
    /// Row heights.
    pub row_heights: Vec<RowHeight>,
    /// Hidden column indices.
    pub hidden_cols: Vec<u32>,
    /// Hidden row indices.
    pub hidden_rows: Vec<u32>,
    /// Merge ranges.
    pub merges: Vec<MergeRange>,
    /// Number of frozen rows.
    pub frozen_rows: u32,
    /// Number of frozen columns.
    pub frozen_cols: u32,
    /// Maximum row with data (1-based dimension).
    pub max_row: u32,
    /// Maximum column with data (1-based dimension).
    pub max_col: u32,
    /// Drawings.
    pub drawings: Vec<Drawing>,
    /// Charts.
    pub charts: Vec<Chart>,
    /// Data validations.
    pub data_validations: Vec<DataValidationRange>,
    /// Conditional formatting rules.
    pub conditional_formatting: Vec<ConditionalFormatting>,
    /// Preprocessed conditional formatting metadata.
    pub conditional_formatting_cache: Vec<ConditionalFormattingCache>,
    /// Sparkline groups.
    pub sparkline_groups: Vec<SparklineGroup>,
    /// Auto-filter.
    pub auto_filter: Option<AutoFilter>,
    /// Comments.
    pub comments: Vec<Comment>,
    /// Comments indexed by cell reference (e.g. "A1" -> index into comments).
    pub comments_by_cell: HashMap<String, usize>,
    /// Cell index by row (cells_by_row[row] = sorted indices into cells).
    pub cells_by_row: Vec<Vec<usize>>,
    /// Default column width (Excel character units).
    pub default_col_width: f64,
    /// Default row height (points).
    pub default_row_height: f64,
    /// Hyperlinks on this sheet (for serialization).
    pub hyperlinks: Vec<HyperlinkDef>,
}

impl Sheet {
    /// Look up a cell index at (row, col) using the cells_by_row index.
    pub fn cell_index_at(&self, row: u32, col: u32) -> Option<usize> {
        let row_cells = self.cells_by_row.get(row as usize)?;
        // Binary search within the row for the column
        row_cells
            .iter()
            .find(|&&idx| self.cells.get(idx).map(|cd| cd.c == col).unwrap_or(false))
            .copied()
    }

    /// Rebuild the cells_by_row index from the cells vector.
    pub fn rebuild_cell_index(&mut self) {
        let max_row = self.cells.iter().map(|c| c.r).max().unwrap_or(0);
        let mut by_row: Vec<Vec<usize>> = vec![Vec::new(); max_row as usize + 1];
        for (idx, cell_data) in self.cells.iter().enumerate() {
            if let Some(row_vec) = by_row.get_mut(cell_data.r as usize) {
                row_vec.push(idx);
            }
        }
        // Sort each row's indices by column
        for row_vec in &mut by_row {
            row_vec.sort_by_key(|&idx| self.cells.get(idx).map(|cd| cd.c).unwrap_or(u32::MAX));
        }
        self.cells_by_row = by_row;
    }

    /// Rebuild the comments_by_cell index.
    pub fn rebuild_comment_index(&mut self) {
        self.comments_by_cell.clear();
        for (idx, _comment) in self.comments.iter().enumerate() {
            // Comments need cell references — in xlview they were keyed by cell ref.
            // For now, the index is built externally by the adapter.
            let _ = idx;
        }
    }
}

/// The top-level workbook data model.
#[derive(Debug)]
pub struct Workbook {
    /// All sheets in the workbook.
    pub sheets: Vec<Sheet>,
    /// Shared strings table.
    pub shared_strings: Vec<String>,
    /// Compiled number format cache.
    pub numfmt_cache: Vec<CompiledFormat>,
    /// Whether the workbook uses the 1904 date system.
    pub date1904: bool,
    /// Embedded images.
    pub images: Vec<EmbeddedImage>,
    /// Theme data.
    pub theme: Theme,
    /// DXF styles for conditional formatting.
    pub dxf_styles: Vec<DxfStyle>,
    /// Resolved styles (indexed by style ID).
    pub resolved_styles: Vec<Option<StyleRef>>,
    /// Default style for cells without explicit styling.
    pub default_style: Option<StyleRef>,
    /// ZIP paths for each sheet (for roundtrip save).
    pub sheet_paths: Vec<String>,
}
