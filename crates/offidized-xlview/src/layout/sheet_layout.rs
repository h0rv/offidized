//! Pre-computed layout data for a sheet.
//!
//! This module computes cell positions once when a sheet is loaded,
//! enabling efficient O(log n) lookups for cell positions and hit testing.

use std::collections::HashMap;

/// Pre-computed layout data for a sheet
#[derive(Clone)]
pub struct SheetLayout {
    /// Cumulative column positions (`col_positions[i]` = x of column i's left edge)
    pub col_positions: Vec<f32>,
    /// Cumulative row positions (`row_positions[i]` = y of row i's top edge)
    pub row_positions: Vec<f32>,
    /// Column widths (0 for hidden columns)
    pub col_widths: Vec<f32>,
    /// Row heights (0 for hidden rows)
    pub row_heights: Vec<f32>,
    /// Merge info lookup by (row, col)
    pub merges: HashMap<(u32, u32), MergeInfo>,
    /// Precomputed skip ranges for vertical grid lines (indexed by column boundary)
    pub merge_vline_skips: Vec<Vec<(u32, u32)>>,
    /// Precomputed skip ranges for horizontal grid lines (indexed by row boundary)
    pub merge_hline_skips: Vec<Vec<(u32, u32)>>,
    /// Maximum row index
    pub max_row: u32,
    /// Maximum column index
    pub max_col: u32,
    /// Number of frozen rows (0 = no frozen rows)
    pub frozen_rows: u32,
    /// Number of frozen columns (0 = no frozen columns)
    pub frozen_cols: u32,
    /// Width of row headers in pixels (0 if headers not shown)
    pub row_header_width: f32,
    /// Height of column headers in pixels (0 if headers not shown)
    pub col_header_height: f32,
}

/// Information about a merged cell region
#[derive(Clone)]
pub struct MergeInfo {
    /// True if this cell is the top-left origin of the merge
    pub is_origin: bool,
    /// Row of the merge origin
    pub origin_row: u32,
    /// Column of the merge origin
    pub origin_col: u32,
    /// Number of rows in the merge
    pub row_span: u32,
    /// Number of columns in the merge
    pub col_span: u32,
}

/// Rectangle representing a cell's bounds
pub struct CellRect {
    /// X position (left edge)
    pub x: f32,
    /// Y position (top edge)
    pub y: f32,
    /// Width of the cell
    pub width: f32,
    /// Height of the cell
    pub height: f32,
    /// True if this cell should be skipped (part of merge but not origin)
    pub skip: bool,
}

/// Default column width in pixels (Excel default ~64px at 100% zoom)
pub const DEFAULT_COL_WIDTH: f32 = 64.0;

/// Default row height in pixels (Excel default ~20px at 100% zoom)
pub const DEFAULT_ROW_HEIGHT: f32 = 20.0;

impl SheetLayout {
    /// Create a new layout from sheet data
    ///
    /// # Arguments
    /// * `max_row` - Maximum row index in the sheet
    /// * `max_col` - Maximum column index in the sheet
    /// * `col_widths` - Map of column index to width
    /// * `row_heights` - Map of row index to height
    /// * `hidden_cols` - Set of hidden column indices
    /// * `hidden_rows` - Set of hidden row indices
    /// * `merges` - List of merge ranges as (start_row, start_col, end_row, end_col)
    /// * `frozen_rows` - Number of frozen rows (0 = none)
    /// * `frozen_cols` - Number of frozen columns (0 = none)
    #[allow(clippy::too_many_arguments)]
    pub fn new(
        max_row: u32,
        max_col: u32,
        col_widths_map: &HashMap<u32, f32>,
        row_heights_map: &HashMap<u32, f32>,
        hidden_cols: &std::collections::HashSet<u32>,
        hidden_rows: &std::collections::HashSet<u32>,
        merge_ranges: &[(u32, u32, u32, u32)],
        frozen_rows: u32,
        frozen_cols: u32,
    ) -> Self {
        // Pre-compute column positions
        let mut col_positions = Vec::with_capacity(max_col as usize + 2);
        let mut col_widths = Vec::with_capacity(max_col as usize + 1);
        let mut x: f32 = 0.0;

        for col in 0..=max_col {
            col_positions.push(x);
            let w = if hidden_cols.contains(&col) {
                0.0
            } else {
                col_widths_map
                    .get(&col)
                    .copied()
                    .unwrap_or(DEFAULT_COL_WIDTH)
            };
            col_widths.push(w);
            x += w;
        }
        col_positions.push(x); // Final edge

        // Pre-compute row positions
        let mut row_positions = Vec::with_capacity(max_row as usize + 2);
        let mut row_heights = Vec::with_capacity(max_row as usize + 1);
        let mut y: f32 = 0.0;

        for row in 0..=max_row {
            row_positions.push(y);
            let h = if hidden_rows.contains(&row) {
                0.0
            } else {
                row_heights_map
                    .get(&row)
                    .copied()
                    .unwrap_or(DEFAULT_ROW_HEIGHT)
            };
            row_heights.push(h);
            y += h;
        }
        row_positions.push(y); // Final edge

        // Build merge map
        let mut merges = HashMap::new();
        for &(start_row, start_col, end_row, end_col) in merge_ranges {
            let row_span = end_row.saturating_sub(start_row) + 1;
            let col_span = end_col.saturating_sub(start_col) + 1;

            for r in start_row..=end_row {
                for c in start_col..=end_col {
                    let is_origin = r == start_row && c == start_col;
                    merges.insert(
                        (r, c),
                        MergeInfo {
                            is_origin,
                            origin_row: start_row,
                            origin_col: start_col,
                            row_span,
                            col_span,
                        },
                    );
                }
            }
        }

        let mut merge_vline_skips: Vec<Vec<(u32, u32)>> = vec![Vec::new(); max_col as usize + 2];
        let mut merge_hline_skips: Vec<Vec<(u32, u32)>> = vec![Vec::new(); max_row as usize + 2];

        for &(start_row, start_col, end_row, end_col) in merge_ranges {
            if start_col < end_col {
                for col in (start_col + 1)..=end_col {
                    if let Some(list) = merge_vline_skips.get_mut(col as usize) {
                        list.push((start_row, end_row + 1));
                    }
                }
            }
            if start_row < end_row {
                for row in (start_row + 1)..=end_row {
                    if let Some(list) = merge_hline_skips.get_mut(row as usize) {
                        list.push((start_col, end_col + 1));
                    }
                }
            }
        }

        for ranges in &mut merge_vline_skips {
            merge_skip_ranges(ranges);
        }
        for ranges in &mut merge_hline_skips {
            merge_skip_ranges(ranges);
        }

        SheetLayout {
            col_positions,
            row_positions,
            col_widths,
            row_heights,
            merges,
            merge_vline_skips,
            merge_hline_skips,
            max_row,
            max_col,
            frozen_rows,
            frozen_cols,
            row_header_width: 0.0,
            col_header_height: 0.0,
        }
    }

    /// Get the width of row headers (0 if not shown)
    pub fn header_width(&self) -> f32 {
        self.row_header_width
    }

    /// Get the height of column headers (0 if not shown)
    pub fn header_height(&self) -> f32 {
        self.col_header_height
    }

    /// Set header dimensions
    pub fn set_header_dimensions(&mut self, row_header_width: f32, col_header_height: f32) {
        self.row_header_width = row_header_width;
        self.col_header_height = col_header_height;
    }

    /// Get cell bounds in sheet coordinates
    pub fn cell_rect(&self, row: u32, col: u32) -> CellRect {
        let x = self.col_positions.get(col as usize).copied().unwrap_or(0.0);
        let y = self.row_positions.get(row as usize).copied().unwrap_or(0.0);
        let mut w = self.col_widths.get(col as usize).copied().unwrap_or(0.0);
        let mut h = self.row_heights.get(row as usize).copied().unwrap_or(0.0);

        // Check for merge
        if let Some(merge) = self.merges.get(&(row, col)) {
            if !merge.is_origin {
                return CellRect {
                    x,
                    y,
                    width: w,
                    height: h,
                    skip: true,
                };
            }
            // Calculate merged size
            let end_col = col + merge.col_span;
            let end_row = row + merge.row_span;
            w = self
                .col_positions
                .get(end_col as usize)
                .copied()
                .unwrap_or(x)
                - x;
            h = self
                .row_positions
                .get(end_row as usize)
                .copied()
                .unwrap_or(y)
                - y;
        }

        CellRect {
            x,
            y,
            width: w,
            height: h,
            skip: false,
        }
    }

    /// Find row at y position (binary search)
    pub fn row_at_y(&self, y: f32) -> Option<u32> {
        if self.row_positions.is_empty() {
            return None;
        }
        match self
            .row_positions
            .binary_search_by(|pos| pos.partial_cmp(&y).unwrap_or(std::cmp::Ordering::Equal))
        {
            Ok(i) => u32::try_from(i).ok(),
            Err(i) => u32::try_from(i.saturating_sub(1)).ok(),
        }
    }

    /// Find column at x position (binary search)
    pub fn col_at_x(&self, x: f32) -> Option<u32> {
        if self.col_positions.is_empty() {
            return None;
        }
        match self
            .col_positions
            .binary_search_by(|pos| pos.partial_cmp(&x).unwrap_or(std::cmp::Ordering::Equal))
        {
            Ok(i) => u32::try_from(i).ok(),
            Err(i) => u32::try_from(i.saturating_sub(1)).ok(),
        }
    }

    /// Get total width of the sheet
    pub fn total_width(&self) -> f32 {
        self.col_positions.last().copied().unwrap_or(0.0)
    }

    /// Get total height of the sheet
    pub fn total_height(&self) -> f32 {
        self.row_positions.last().copied().unwrap_or(0.0)
    }

    /// Get column width at index
    pub fn col_width(&self, col: u32) -> f32 {
        self.col_widths
            .get(col as usize)
            .copied()
            .unwrap_or(DEFAULT_COL_WIDTH)
    }

    /// Get row height at index
    pub fn row_height(&self, row: u32) -> f32 {
        self.row_heights
            .get(row as usize)
            .copied()
            .unwrap_or(DEFAULT_ROW_HEIGHT)
    }

    /// Get the total height of frozen rows (returns 0 if no frozen rows)
    pub fn frozen_rows_height(&self) -> f32 {
        if self.frozen_rows == 0 {
            return 0.0;
        }
        self.row_positions
            .get(self.frozen_rows as usize)
            .copied()
            .unwrap_or(0.0)
    }

    /// Get the total width of frozen columns (returns 0 if no frozen columns)
    pub fn frozen_cols_width(&self) -> f32 {
        if self.frozen_cols == 0 {
            return 0.0;
        }
        self.col_positions
            .get(self.frozen_cols as usize)
            .copied()
            .unwrap_or(0.0)
    }
}

fn merge_skip_ranges(ranges: &mut Vec<(u32, u32)>) {
    if ranges.len() <= 1 {
        return;
    }

    ranges.sort_by_key(|r| r.0);
    let mut merged: Vec<(u32, u32)> = Vec::with_capacity(ranges.len());
    for (start, end) in ranges.drain(..) {
        if let Some(last) = merged.last_mut() {
            if start <= last.1 {
                if end > last.1 {
                    last.1 = end;
                }
                continue;
            }
        }
        merged.push((start, end));
    }
    *ranges = merged;
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
    use super::*;

    #[test]
    fn test_basic_layout() {
        let layout = SheetLayout::new(
            10,
            5,
            &HashMap::new(),
            &HashMap::new(),
            &std::collections::HashSet::new(),
            &std::collections::HashSet::new(),
            &[],
            0,
            0,
        );

        assert_eq!(layout.max_row, 10);
        assert_eq!(layout.max_col, 5);
        assert_eq!(layout.total_width(), DEFAULT_COL_WIDTH * 6.0);
        assert_eq!(layout.total_height(), DEFAULT_ROW_HEIGHT * 11.0);
    }

    #[test]
    fn test_cell_rect() {
        let layout = SheetLayout::new(
            10,
            5,
            &HashMap::new(),
            &HashMap::new(),
            &std::collections::HashSet::new(),
            &std::collections::HashSet::new(),
            &[],
            0,
            0,
        );

        let rect = layout.cell_rect(0, 0);
        assert_eq!(rect.x, 0.0);
        assert_eq!(rect.y, 0.0);
        assert_eq!(rect.width, DEFAULT_COL_WIDTH);
        assert_eq!(rect.height, DEFAULT_ROW_HEIGHT);
        assert!(!rect.skip);

        let rect = layout.cell_rect(1, 2);
        assert_eq!(rect.x, DEFAULT_COL_WIDTH * 2.0);
        assert_eq!(rect.y, DEFAULT_ROW_HEIGHT);
    }

    #[test]
    fn test_merged_cells() {
        // Merge A1:B2 (rows 0-1, cols 0-1)
        let layout = SheetLayout::new(
            10,
            5,
            &HashMap::new(),
            &HashMap::new(),
            &std::collections::HashSet::new(),
            &std::collections::HashSet::new(),
            &[(0, 0, 1, 1)],
            0,
            0,
        );

        // Origin cell should have full merged dimensions
        let rect = layout.cell_rect(0, 0);
        assert!(!rect.skip);
        assert_eq!(rect.width, DEFAULT_COL_WIDTH * 2.0);
        assert_eq!(rect.height, DEFAULT_ROW_HEIGHT * 2.0);

        // Non-origin cells should be skipped
        let rect = layout.cell_rect(0, 1);
        assert!(rect.skip);
        let rect = layout.cell_rect(1, 0);
        assert!(rect.skip);
        let rect = layout.cell_rect(1, 1);
        assert!(rect.skip);
    }

    #[test]
    fn test_row_at_y() {
        let layout = SheetLayout::new(
            10,
            5,
            &HashMap::new(),
            &HashMap::new(),
            &std::collections::HashSet::new(),
            &std::collections::HashSet::new(),
            &[],
            0,
            0,
        );

        assert_eq!(layout.row_at_y(0.0), Some(0));
        assert_eq!(layout.row_at_y(10.0), Some(0));
        assert_eq!(layout.row_at_y(DEFAULT_ROW_HEIGHT), Some(1));
        assert_eq!(layout.row_at_y(DEFAULT_ROW_HEIGHT * 2.5), Some(2));
    }

    #[test]
    fn test_col_at_x() {
        let layout = SheetLayout::new(
            10,
            5,
            &HashMap::new(),
            &HashMap::new(),
            &std::collections::HashSet::new(),
            &std::collections::HashSet::new(),
            &[],
            0,
            0,
        );

        assert_eq!(layout.col_at_x(0.0), Some(0));
        assert_eq!(layout.col_at_x(32.0), Some(0));
        assert_eq!(layout.col_at_x(DEFAULT_COL_WIDTH), Some(1));
        assert_eq!(layout.col_at_x(DEFAULT_COL_WIDTH * 2.5), Some(2));
    }
}
