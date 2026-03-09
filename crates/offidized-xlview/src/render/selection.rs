//! Selection overlay helpers.
//!
//! These helpers keep selection math testable without depending on Canvas APIs.

use crate::layout::{SheetLayout, Viewport};

#[derive(Debug, Clone, Copy, PartialEq)]
pub struct SelectionRect {
    pub x: f64,
    pub y: f64,
    pub w: f64,
    pub h: f64,
    pub draw_top: bool,
    pub draw_bottom: bool,
    pub draw_left: bool,
    pub draw_right: bool,
}

fn to_screen_flags(
    x: f32,
    y: f32,
    row_frozen: bool,
    col_frozen: bool,
    layout: &SheetLayout,
    viewport: &Viewport,
) -> (f32, f32) {
    let frozen_width = layout.frozen_cols_width();
    let frozen_height = layout.frozen_rows_height();
    let scale = viewport.scale;
    let sx = if col_frozen {
        x * scale
    } else {
        frozen_width * scale + (x - viewport.scroll_x) * scale
    };
    let sy = if row_frozen {
        y * scale
    } else {
        frozen_height * scale + (y - viewport.scroll_y) * scale
    };
    (sx, sy)
}

#[allow(clippy::too_many_lines)]
pub fn selection_rects(
    selection: (u32, u32, u32, u32),
    layout: &SheetLayout,
    viewport: &Viewport,
) -> Vec<SelectionRect> {
    let (start_row, start_col, end_row, end_col) = selection;
    let min_row = start_row.min(end_row);
    let max_row = start_row.max(end_row);
    let min_col = start_col.min(end_col);
    let max_col = start_col.max(end_col);

    let frozen_rows = layout.frozen_rows;
    let frozen_cols = layout.frozen_cols;

    let frozen_row_range = if frozen_rows > 0 {
        let end = max_row.min(frozen_rows.saturating_sub(1));
        if min_row <= end {
            Some((min_row, end))
        } else {
            None
        }
    } else {
        None
    };
    let scroll_row_range = if max_row >= frozen_rows {
        let start = min_row.max(frozen_rows);
        if start <= max_row {
            Some((start, max_row))
        } else {
            None
        }
    } else {
        None
    };

    let frozen_col_range = if frozen_cols > 0 {
        let end = max_col.min(frozen_cols.saturating_sub(1));
        if min_col <= end {
            Some((min_col, end))
        } else {
            None
        }
    } else {
        None
    };
    let scroll_col_range = if max_col >= frozen_cols {
        let start = min_col.max(frozen_cols);
        if start <= max_col {
            Some((start, max_col))
        } else {
            None
        }
    } else {
        None
    };

    let mut rects = Vec::new();

    let mut push_rect =
        |row_range: (u32, u32), col_range: (u32, u32), row_frozen: bool, col_frozen: bool| {
            let (row_start, row_end) = row_range;
            let (col_start, col_end) = col_range;
            let x1 = layout
                .col_positions
                .get(col_start as usize)
                .copied()
                .unwrap_or(0.0);
            let y1 = layout
                .row_positions
                .get(row_start as usize)
                .copied()
                .unwrap_or(0.0);
            let x2 = layout
                .col_positions
                .get((col_end + 1) as usize)
                .copied()
                .unwrap_or(x1);
            let y2 = layout
                .row_positions
                .get((row_end + 1) as usize)
                .copied()
                .unwrap_or(y1);

            let (sx1, sy1) = to_screen_flags(x1, y1, row_frozen, col_frozen, layout, viewport);
            let (sx2, sy2) = to_screen_flags(x2, y2, row_frozen, col_frozen, layout, viewport);
            let w = (sx2 - sx1).max(0.0);
            let h = (sy2 - sy1).max(0.0);
            if w <= 0.0 || h <= 0.0 {
                return;
            }

            rects.push(SelectionRect {
                x: f64::from(sx1),
                y: f64::from(sy1),
                w: f64::from(w),
                h: f64::from(h),
                draw_top: row_start == min_row,
                draw_bottom: row_end == max_row,
                draw_left: col_start == min_col,
                draw_right: col_end == max_col,
            });
        };

    if let (Some(row_range), Some(col_range)) = (frozen_row_range, frozen_col_range) {
        push_rect(row_range, col_range, true, true);
    }
    if let (Some(row_range), Some(col_range)) = (frozen_row_range, scroll_col_range) {
        push_rect(row_range, col_range, true, false);
    }
    if let (Some(row_range), Some(col_range)) = (scroll_row_range, frozen_col_range) {
        push_rect(row_range, col_range, false, true);
    }
    if let (Some(row_range), Some(col_range)) = (scroll_row_range, scroll_col_range) {
        push_rect(row_range, col_range, false, false);
    }

    rects
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

    fn layout_with_frozen(rows: u32, cols: u32) -> SheetLayout {
        let col_widths = std::collections::HashMap::new();
        let row_heights = std::collections::HashMap::new();
        let hidden_cols = std::collections::HashSet::new();
        let hidden_rows = std::collections::HashSet::new();
        SheetLayout::new(
            10,
            10,
            &col_widths,
            &row_heights,
            &hidden_cols,
            &hidden_rows,
            &[],
            rows,
            cols,
        )
    }

    #[test]
    fn selection_rects_split_frozen_rows() {
        let layout = layout_with_frozen(1, 0);
        let mut viewport = Viewport::new();
        let row_height = layout.row_positions[1] - layout.row_positions[0];
        viewport.scroll_y = row_height * 2.0;
        let rects = selection_rects((0, 0, 2, 1), &layout, &viewport);
        assert_eq!(rects.len(), 2);

        let frozen = rects.iter().find(|r| r.draw_top).unwrap();
        let scroll = rects.iter().find(|r| r.draw_bottom).unwrap();
        assert_eq!(frozen.y, 0.0);
        let expected_scroll_y =
            f64::from(layout.frozen_rows_height()) + f64::from(row_height - viewport.scroll_y);
        assert!((scroll.y - expected_scroll_y).abs() < 0.1);
    }

    #[test]
    fn selection_rects_split_frozen_cols() {
        let layout = layout_with_frozen(0, 1);
        let mut viewport = Viewport::new();
        let col_width = layout.col_positions[1] - layout.col_positions[0];
        viewport.scroll_x = col_width * 2.0;
        let rects = selection_rects((0, 0, 1, 2), &layout, &viewport);
        assert_eq!(rects.len(), 2);

        let frozen = rects.iter().find(|r| r.draw_left).unwrap();
        let scroll = rects.iter().find(|r| r.draw_right).unwrap();
        assert_eq!(frozen.x, 0.0);
        let expected_scroll_x =
            f64::from(layout.frozen_cols_width()) + f64::from(col_width - viewport.scroll_x);
        assert!((scroll.x - expected_scroll_x).abs() < 0.1);
    }
}
