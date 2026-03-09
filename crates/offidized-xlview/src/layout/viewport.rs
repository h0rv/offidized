//! Viewport state management for scrolling and zoom.

use super::SheetLayout;

/// Viewport state - represents the visible area of the spreadsheet
#[derive(Clone)]
pub struct Viewport {
    /// Horizontal scroll position in sheet coordinates
    pub scroll_x: f32,
    /// Vertical scroll position in sheet coordinates
    pub scroll_y: f32,
    /// Viewport width in pixels
    pub width: f32,
    /// Viewport height in pixels
    pub height: f32,
    /// Zoom scale factor (1.0 = 100%)
    pub scale: f32,
    /// Tab bar horizontal scroll offset (pixels)
    pub tab_scroll_x: f32,
}

impl Default for Viewport {
    fn default() -> Self {
        Self::new()
    }
}

impl Viewport {
    /// Create a new viewport with default values
    pub fn new() -> Self {
        Self {
            scroll_x: 0.0,
            scroll_y: 0.0,
            width: 800.0,
            height: 600.0,
            scale: 1.0,
            tab_scroll_x: 0.0,
        }
    }

    /// Get visible scrollable row range (inclusive) based on current scroll position.
    pub fn visible_rows(&self, layout: &SheetLayout) -> (u32, u32) {
        self.visible_rows_in_height(layout, self.height)
    }

    /// Get visible scrollable column range (inclusive) based on current scroll position.
    pub fn visible_cols(&self, layout: &SheetLayout) -> (u32, u32) {
        self.visible_cols_in_width(layout, self.width)
    }

    /// Get visible scrollable row range (inclusive) for a given viewport height.
    /// This allows callers to use a content height that excludes headers/scrollbars.
    pub fn visible_rows_in_height(&self, layout: &SheetLayout, viewport_height: f32) -> (u32, u32) {
        let frozen_height = layout.frozen_rows_height();
        let scrollable_viewport_height = (viewport_height - frozen_height).max(0.0);

        let start = layout.row_at_y(self.scroll_y).unwrap_or(layout.max_row);
        let end = layout
            .row_at_y(self.scroll_y + scrollable_viewport_height)
            .unwrap_or(layout.max_row);
        (start.min(layout.max_row), end.min(layout.max_row))
    }

    /// Get visible scrollable column range (inclusive) for a given viewport width.
    /// This allows callers to use a content width that excludes headers/scrollbars.
    pub fn visible_cols_in_width(&self, layout: &SheetLayout, viewport_width: f32) -> (u32, u32) {
        let frozen_width = layout.frozen_cols_width();
        let scrollable_viewport_width = (viewport_width - frozen_width).max(0.0);

        let start = layout.col_at_x(self.scroll_x).unwrap_or(layout.max_col);
        let end = layout
            .col_at_x(self.scroll_x + scrollable_viewport_width)
            .unwrap_or(layout.max_col);
        (start.min(layout.max_col), end.min(layout.max_col))
    }

    /// Convert sheet coordinates to screen coordinates
    pub fn to_screen(&self, x: f32, y: f32) -> (f32, f32) {
        (
            (x - self.scroll_x) * self.scale,
            (y - self.scroll_y) * self.scale,
        )
    }

    /// Convert sheet coordinates to screen coordinates for a cell at (row, col),
    /// accounting for frozen panes.
    ///
    /// For cells in frozen rows/cols, the position is fixed (not affected by scroll).
    /// For cells in the scrollable area, the position accounts for both scroll and
    /// the space taken by frozen regions.
    ///
    /// Layout:
    /// - Frozen cells render at their natural layout position (no scroll)
    /// - Non-frozen cells render at: frozen_size + (layout_pos - scroll_pos) * scale
    ///   where scroll_pos starts at frozen_size (minimum scroll)
    pub fn to_screen_frozen(
        &self,
        x: f32,
        y: f32,
        row: u32,
        col: u32,
        layout: &SheetLayout,
    ) -> (f32, f32) {
        let frozen_width = layout.frozen_cols_width();
        let frozen_height = layout.frozen_rows_height();

        let screen_x = if col < layout.frozen_cols {
            // Cell is in frozen columns - fixed position (no scroll)
            x * self.scale
        } else {
            // Cell is in scrollable area
            // Screen position = frozen_width + (layout_x - scroll_x) * scale
            // scroll_x starts at frozen_width, so when scroll_x = frozen_width,
            // the cell at frozen_width appears at screen position frozen_width
            frozen_width * self.scale + (x - self.scroll_x) * self.scale
        };

        let screen_y = if row < layout.frozen_rows {
            // Cell is in frozen rows - fixed position (no scroll)
            y * self.scale
        } else {
            // Cell is in scrollable area
            frozen_height * self.scale + (y - self.scroll_y) * self.scale
        };

        (screen_x, screen_y)
    }

    /// Convert screen coordinates to sheet coordinates
    pub fn to_sheet(&self, screen_x: f32, screen_y: f32) -> (f32, f32) {
        (
            screen_x / self.scale + self.scroll_x,
            screen_y / self.scale + self.scroll_y,
        )
    }

    /// Clamp scroll position to valid range.
    ///
    /// For frozen panes, scroll positions are relative to the frozen region.
    /// scroll_x starts at the frozen column boundary, scroll_y starts at frozen row boundary.
    pub fn clamp_scroll(&mut self, layout: &SheetLayout) {
        let frozen_width = layout.frozen_cols_width();
        let frozen_height = layout.frozen_rows_height();

        // Minimum scroll is at the frozen boundary
        let min_x = frozen_width;
        let min_y = frozen_height;

        // Maximum scroll allows viewing the end of the content
        // The scrollable area width = total_width - frozen_width
        // We want to be able to scroll until the last content is visible
        let scrollable_width = layout.total_width() - frozen_width;
        let scrollable_height = layout.total_height() - frozen_height;
        let viewport_content_width = self.width - frozen_width;
        let viewport_content_height = self.height - frozen_height;

        let max_x = frozen_width + (scrollable_width - viewport_content_width).max(0.0);
        let max_y = frozen_height + (scrollable_height - viewport_content_height).max(0.0);

        self.scroll_x = self.scroll_x.clamp(min_x, max_x);
        self.scroll_y = self.scroll_y.clamp(min_y, max_y);
    }

    /// Scroll by delta amounts
    pub fn scroll_by(&mut self, delta_x: f32, delta_y: f32, layout: &SheetLayout) {
        self.scroll_x += delta_x;
        self.scroll_y += delta_y;
        self.clamp_scroll(layout);
    }

    /// Set absolute scroll position
    pub fn set_scroll(&mut self, x: f32, y: f32, layout: &SheetLayout) {
        self.scroll_x = x;
        self.scroll_y = y;
        self.clamp_scroll(layout);
    }

    /// Resize the viewport
    pub fn resize(&mut self, width: f32, height: f32) {
        self.width = width;
        self.height = height;
    }

    /// Convert x coordinate to screen x for a column boundary (grid lines).
    ///
    /// For grid lines, use `col <= frozen_cols` because the line at frozen_cols
    /// is the boundary between frozen and scrollable regions and should be fixed.
    pub fn screen_x_for_grid(&self, x: f32, col: u32, layout: &SheetLayout) -> f32 {
        if col <= layout.frozen_cols {
            // Line is at or before the frozen boundary - fixed position
            x * self.scale
        } else {
            // Line is in the scrollable region
            layout.frozen_cols_width() * self.scale + (x - self.scroll_x) * self.scale
        }
    }

    /// Convert y coordinate to screen y for a row boundary (grid lines).
    ///
    /// For grid lines, use `row <= frozen_rows` because the line at frozen_rows
    /// is the boundary between frozen and scrollable regions and should be fixed.
    pub fn screen_y_for_grid(&self, y: f32, row: u32, layout: &SheetLayout) -> f32 {
        if row <= layout.frozen_rows {
            // Line is at or before the frozen boundary - fixed position
            y * self.scale
        } else {
            // Line is in the scrollable region
            layout.frozen_rows_height() * self.scale + (y - self.scroll_y) * self.scale
        }
    }
}
