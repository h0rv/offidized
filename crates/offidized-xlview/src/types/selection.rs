/// Type of selection for row/column headers
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum SelectionType {
    /// Standard cell selection (default)
    #[default]
    CellRange,
    /// Entire row(s) selected
    RowRange,
    /// Entire column(s) selected
    ColumnRange,
    /// All cells selected (corner click)
    All,
}

/// Selection state supporting cell, row, column, and all selection types
#[derive(Debug, Clone)]
pub struct Selection {
    pub selection_type: SelectionType,
    pub start_row: u32,
    pub start_col: u32,
    pub end_row: u32,
    pub end_col: u32,
}

impl Selection {
    /// Create a new cell range selection
    pub fn cell_range(start_row: u32, start_col: u32, end_row: u32, end_col: u32) -> Self {
        Self {
            selection_type: SelectionType::CellRange,
            start_row,
            start_col,
            end_row,
            end_col,
        }
    }

    /// Create a row range selection
    pub fn row_range(start_row: u32, end_row: u32) -> Self {
        Self {
            selection_type: SelectionType::RowRange,
            start_row,
            start_col: 0,
            end_row,
            end_col: u32::MAX,
        }
    }

    /// Create a column range selection
    pub fn column_range(start_col: u32, end_col: u32) -> Self {
        Self {
            selection_type: SelectionType::ColumnRange,
            start_row: 0,
            start_col,
            end_row: u32::MAX,
            end_col,
        }
    }

    /// Create a select-all selection
    pub fn all() -> Self {
        Self {
            selection_type: SelectionType::All,
            start_row: 0,
            start_col: 0,
            end_row: u32::MAX,
            end_col: u32::MAX,
        }
    }

    /// Get normalized bounds (min/max)
    pub fn bounds(&self) -> (u32, u32, u32, u32) {
        (
            self.start_row.min(self.end_row),
            self.start_col.min(self.end_col),
            self.start_row.max(self.end_row),
            self.start_col.max(self.end_col),
        )
    }
}

/// Configuration for row and column headers
#[derive(Debug, Clone)]
pub struct HeaderConfig {
    /// Whether headers are visible
    pub visible: bool,
    /// Width of row headers in pixels (~40px default)
    pub row_header_width: f32,
    /// Height of column headers in pixels (~20px default)
    pub col_header_height: f32,
    /// Background color for headers
    pub background_color: String,
    /// Text color for header labels
    pub text_color: String,
    /// Border color for headers
    pub border_color: String,
    /// Background color for selected headers
    pub selected_bg_color: String,
}

impl Default for HeaderConfig {
    fn default() -> Self {
        Self {
            visible: true,
            row_header_width: 40.0,
            col_header_height: 20.0,
            background_color: "#F3F3F3".to_string(),
            text_color: "#595959".to_string(),
            border_color: "#CCCCCC".to_string(),
            selected_bg_color: "#CFD8E8".to_string(),
        }
    }
}
