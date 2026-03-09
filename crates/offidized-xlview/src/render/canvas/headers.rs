//! Row and column header rendering for spreadsheet viewer.
//!
//! This module provides Excel/Google Sheets-style row and column headers
//! with support for:
//! - Column headers: A, B, C, ... Z, AA, AB, ...
//! - Row headers: 1, 2, 3, ...
//! - Selection highlighting for selected rows/columns
//! - Frozen pane support

use web_sys::CanvasRenderingContext2d;

use crate::layout::{SheetLayout, Viewport};
use crate::types::selection::{HeaderConfig, Selection, SelectionType};

/// Header color palette - matches Excel/Google Sheets styling
mod colors {
    /// Header background when column/row is fully selected (blue tint)
    pub const HEADER_ACTIVE_BG: &str = "#A8C7FA";
    /// Header text color when selected (blue)
    pub const HEADER_TEXT_SELECTED: &str = "#1A73E8";
    /// Header resize handle / divider color
    pub const HEADER_RESIZE_HANDLE: &str = "#80868B";
}

/// Convert a 0-based column index to Excel column letters (A, B, ..., Z, AA, AB, ...)
pub fn col_to_letter(col: u32) -> String {
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

/// Render column headers (A, B, C, ...)
#[allow(clippy::cast_possible_truncation)]
pub fn render_column_headers(
    ctx: &CanvasRenderingContext2d,
    layout: &SheetLayout,
    viewport: &Viewport,
    config: &HeaderConfig,
    selection: Option<&Selection>,
    content_width: f64,
) {
    if !config.visible || config.col_header_height <= 0.0 {
        return;
    }

    let header_height = f64::from(config.col_header_height);
    let header_width = f64::from(config.row_header_width);
    let canvas_width = content_width + header_width;

    // Determine which columns are selected (for highlighting)
    let (selected_cols, fully_selected_cols) = get_selected_columns_with_type(selection, layout);

    let frozen_width = f64::from(layout.frozen_cols_width());
    let (start_col, end_col) = viewport.visible_cols_in_width(layout, content_width as f32);

    // Set up text rendering - use system font stack for better cross-platform rendering
    ctx.set_font("500 11px -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif");
    ctx.set_text_align("center");
    ctx.set_text_baseline("middle");

    // === PASS 1: Render SCROLLABLE column headers (clipped to scrollable area) ===
    // This must come FIRST so frozen headers can render on top
    ctx.save();
    ctx.begin_path();
    // Clip to scrollable area only (after frozen columns)
    ctx.rect(
        header_width + frozen_width,
        0.0,
        canvas_width - header_width - frozen_width,
        header_height,
    );
    ctx.clip();

    // Background for scrollable area
    ctx.set_fill_style_str(&config.background_color);
    ctx.fill_rect(
        header_width + frozen_width,
        0.0,
        canvas_width - header_width - frozen_width,
        header_height,
    );

    // Render scrollable columns
    for col in start_col.max(layout.frozen_cols)..=end_col {
        let col_x = layout
            .col_positions
            .get(col as usize)
            .copied()
            .unwrap_or(0.0);
        let col_width = layout.col_widths.get(col as usize).copied().unwrap_or(0.0);
        if col_width <= 0.0 {
            continue;
        }

        // Screen position for scrollable area
        let screen_x = frozen_width + f64::from(col_x - viewport.scroll_x) + header_width;
        if screen_x + f64::from(col_width) < header_width + frozen_width {
            continue; // Off screen to the left
        }
        if screen_x > canvas_width {
            break; // Off screen to the right
        }

        render_single_column_header(
            ctx,
            col,
            screen_x,
            f64::from(col_width),
            header_height,
            config,
            selected_cols.contains(&col),
            fully_selected_cols.contains(&col),
        );
    }

    ctx.restore();

    // === PASS 2: Render FROZEN column headers (on top, in fixed position) ===
    if layout.frozen_cols > 0 {
        // Background for frozen column header area (opaque, covers any bleed-through)
        ctx.set_fill_style_str(&config.background_color);
        ctx.fill_rect(header_width, 0.0, frozen_width, header_height);

        // Render frozen column headers
        for col in 0..layout.frozen_cols {
            let col_x = layout
                .col_positions
                .get(col as usize)
                .copied()
                .unwrap_or(0.0);
            let col_width = layout.col_widths.get(col as usize).copied().unwrap_or(0.0);
            if col_width <= 0.0 {
                continue;
            }

            let screen_x = f64::from(col_x) + header_width;
            render_single_column_header(
                ctx,
                col,
                screen_x,
                f64::from(col_width),
                header_height,
                config,
                selected_cols.contains(&col),
                fully_selected_cols.contains(&col),
            );
        }
    }

    // Draw bottom border of header row (full width)
    ctx.set_stroke_style_str(&config.border_color);
    ctx.set_line_width(1.0);
    ctx.begin_path();
    ctx.move_to(header_width, header_height - 0.5);
    ctx.line_to(canvas_width, header_height - 0.5);
    ctx.stroke();

    // Draw frozen column divider in header if there are frozen columns
    if layout.frozen_cols > 0 {
        let divider_x = header_width + frozen_width;
        ctx.set_stroke_style_str(colors::HEADER_RESIZE_HANDLE);
        ctx.set_line_width(2.0);
        ctx.begin_path();
        ctx.move_to(divider_x, 0.0);
        ctx.line_to(divider_x, header_height);
        ctx.stroke();
    }
}

/// Render a single column header cell
#[allow(clippy::too_many_arguments)]
fn render_single_column_header(
    ctx: &CanvasRenderingContext2d,
    col: u32,
    x: f64,
    width: f64,
    height: f64,
    config: &HeaderConfig,
    is_selected: bool,
    is_fully_selected: bool,
) {
    // Draw background based on selection state
    if is_fully_selected {
        ctx.set_fill_style_str(colors::HEADER_ACTIVE_BG);
        ctx.fill_rect(x, 0.0, width, height);
    } else if is_selected {
        ctx.set_fill_style_str(&config.selected_bg_color);
        ctx.fill_rect(x, 0.0, width, height);
    }

    // Draw right border (cell separator)
    ctx.set_stroke_style_str(&config.border_color);
    ctx.set_line_width(1.0);
    ctx.begin_path();
    ctx.move_to(x + width - 0.5, 0.0);
    ctx.line_to(x + width - 0.5, height);
    ctx.stroke();

    // Draw column letter with appropriate color
    let text_color = if is_fully_selected {
        colors::HEADER_TEXT_SELECTED
    } else {
        &config.text_color
    };
    ctx.set_fill_style_str(text_color);

    let label = col_to_letter(col);

    // For narrow columns, truncate or skip text
    if width >= 20.0 {
        let _ = ctx.fill_text(&label, x + width / 2.0, height / 2.0);
    }
}

/// Render row headers (1, 2, 3, ...)
#[allow(clippy::cast_possible_truncation)]
pub fn render_row_headers(
    ctx: &CanvasRenderingContext2d,
    layout: &SheetLayout,
    viewport: &Viewport,
    config: &HeaderConfig,
    selection: Option<&Selection>,
    content_height: f64,
) {
    if !config.visible || config.row_header_width <= 0.0 {
        return;
    }

    let header_width = f64::from(config.row_header_width);
    let header_height = f64::from(config.col_header_height);
    let canvas_height = content_height + header_height;

    // Determine which rows are selected (for highlighting)
    let (selected_rows, fully_selected_rows) = get_selected_rows_with_type(selection, layout);

    let frozen_height = f64::from(layout.frozen_rows_height());
    let (start_row, end_row) = viewport.visible_rows_in_height(layout, content_height as f32);

    // Set up text rendering
    ctx.set_font("500 11px -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif");
    ctx.set_text_align("center");
    ctx.set_text_baseline("middle");

    // === PASS 1: Render SCROLLABLE row headers (clipped to scrollable area) ===
    // This must come FIRST so frozen headers can render on top
    ctx.save();
    ctx.begin_path();
    // Clip to scrollable area only (below frozen rows)
    ctx.rect(
        0.0,
        header_height + frozen_height,
        header_width,
        canvas_height - header_height - frozen_height,
    );
    ctx.clip();

    // Background for scrollable area
    ctx.set_fill_style_str(&config.background_color);
    ctx.fill_rect(
        0.0,
        header_height + frozen_height,
        header_width,
        canvas_height - header_height - frozen_height,
    );

    // Render scrollable rows
    for row in start_row.max(layout.frozen_rows)..=end_row {
        let row_y = layout
            .row_positions
            .get(row as usize)
            .copied()
            .unwrap_or(0.0);
        let row_height = layout.row_heights.get(row as usize).copied().unwrap_or(0.0);
        if row_height <= 0.0 {
            continue;
        }

        // Screen position for scrollable area
        let screen_y = frozen_height + f64::from(row_y - viewport.scroll_y) + header_height;
        if screen_y + f64::from(row_height) < header_height + frozen_height {
            continue; // Off screen above
        }
        if screen_y > canvas_height {
            break; // Off screen below
        }

        render_single_row_header(
            ctx,
            row,
            screen_y,
            header_width,
            f64::from(row_height),
            config,
            selected_rows.contains(&row),
            fully_selected_rows.contains(&row),
        );
    }

    ctx.restore();

    // === PASS 2: Render FROZEN row headers (on top, in fixed position) ===
    if layout.frozen_rows > 0 {
        // Background for frozen row header area (opaque, covers any bleed-through)
        ctx.set_fill_style_str(&config.background_color);
        ctx.fill_rect(0.0, header_height, header_width, frozen_height);

        // Render frozen row headers
        for row in 0..layout.frozen_rows {
            let row_y = layout
                .row_positions
                .get(row as usize)
                .copied()
                .unwrap_or(0.0);
            let row_height = layout.row_heights.get(row as usize).copied().unwrap_or(0.0);
            if row_height <= 0.0 {
                continue;
            }

            let screen_y = f64::from(row_y) + header_height;
            render_single_row_header(
                ctx,
                row,
                screen_y,
                header_width,
                f64::from(row_height),
                config,
                selected_rows.contains(&row),
                fully_selected_rows.contains(&row),
            );
        }
    }

    // Draw right border of header column (full height)
    ctx.set_stroke_style_str(&config.border_color);
    ctx.set_line_width(1.0);
    ctx.begin_path();
    ctx.move_to(header_width - 0.5, header_height);
    ctx.line_to(header_width - 0.5, canvas_height);
    ctx.stroke();

    // Draw frozen row divider in header if there are frozen rows
    if layout.frozen_rows > 0 {
        let divider_y = header_height + frozen_height;
        ctx.set_stroke_style_str(colors::HEADER_RESIZE_HANDLE);
        ctx.set_line_width(2.0);
        ctx.begin_path();
        ctx.move_to(0.0, divider_y);
        ctx.line_to(header_width, divider_y);
        ctx.stroke();
    }
}

/// Render a single row header cell
#[allow(clippy::too_many_arguments)]
fn render_single_row_header(
    ctx: &CanvasRenderingContext2d,
    row: u32,
    y: f64,
    width: f64,
    height: f64,
    config: &HeaderConfig,
    is_selected: bool,
    is_fully_selected: bool,
) {
    // Draw background based on selection state
    if is_fully_selected {
        ctx.set_fill_style_str(colors::HEADER_ACTIVE_BG);
        ctx.fill_rect(0.0, y, width, height);
    } else if is_selected {
        ctx.set_fill_style_str(&config.selected_bg_color);
        ctx.fill_rect(0.0, y, width, height);
    }

    // Draw bottom border (cell separator)
    ctx.set_stroke_style_str(&config.border_color);
    ctx.set_line_width(1.0);
    ctx.begin_path();
    ctx.move_to(0.0, y + height - 0.5);
    ctx.line_to(width, y + height - 0.5);
    ctx.stroke();

    // Draw row number with appropriate color
    let text_color = if is_fully_selected {
        colors::HEADER_TEXT_SELECTED
    } else {
        &config.text_color
    };
    ctx.set_fill_style_str(text_color);

    // Row numbers are 1-indexed for display
    let label = (row + 1).to_string();

    // For very short rows, skip text
    if height >= 12.0 {
        let _ = ctx.fill_text(&label, width / 2.0, y + height / 2.0);
    }
}

/// Render the corner cell (intersection of row and column headers)
pub fn render_header_corner(
    ctx: &CanvasRenderingContext2d,
    config: &HeaderConfig,
    all_selected: bool,
) {
    if !config.visible {
        return;
    }

    let width = f64::from(config.row_header_width);
    let height = f64::from(config.col_header_height);

    if width <= 0.0 || height <= 0.0 {
        return;
    }

    // Draw background
    let bg_color = if all_selected {
        colors::HEADER_ACTIVE_BG
    } else {
        &config.background_color
    };
    ctx.set_fill_style_str(bg_color);
    ctx.fill_rect(0.0, 0.0, width, height);

    // Draw "select all" triangle indicator (like Excel/Google Sheets)
    if !all_selected {
        ctx.set_fill_style_str(colors::HEADER_RESIZE_HANDLE);
        ctx.begin_path();
        // Small triangle in bottom-right corner
        let tri_size = 6.0;
        let margin = 4.0;
        ctx.move_to(width - margin, height - margin - tri_size);
        ctx.line_to(width - margin, height - margin);
        ctx.line_to(width - margin - tri_size, height - margin);
        ctx.close_path();
        ctx.fill();
    }

    // Draw borders
    ctx.set_stroke_style_str(&config.border_color);
    ctx.set_line_width(1.0);

    // Right border
    ctx.begin_path();
    ctx.move_to(width - 0.5, 0.0);
    ctx.line_to(width - 0.5, height);
    ctx.stroke();

    // Bottom border
    ctx.begin_path();
    ctx.move_to(0.0, height - 0.5);
    ctx.line_to(width, height - 0.5);
    ctx.stroke();
}

/// Get selected columns with distinction between partial and full selection
fn get_selected_columns_with_type(
    selection: Option<&Selection>,
    layout: &SheetLayout,
) -> (
    std::collections::HashSet<u32>,
    std::collections::HashSet<u32>,
) {
    let mut selected = std::collections::HashSet::new();
    let mut fully_selected = std::collections::HashSet::new();

    let Some(sel) = selection else {
        return (selected, fully_selected);
    };

    match sel.selection_type {
        SelectionType::ColumnRange => {
            // Entire columns are selected
            let (_, min_col, _, max_col) = sel.bounds();
            let max_col = max_col.min(layout.max_col);
            for col in min_col..=max_col {
                selected.insert(col);
                fully_selected.insert(col);
            }
        }
        SelectionType::All => {
            // All columns are fully selected
            for col in 0..=layout.max_col {
                selected.insert(col);
                fully_selected.insert(col);
            }
        }
        SelectionType::CellRange => {
            // Partial selection - columns contain selection but aren't fully selected
            let (_, min_col, _, max_col) = sel.bounds();
            let max_col = max_col.min(layout.max_col);
            for col in min_col..=max_col {
                selected.insert(col);
            }
        }
        SelectionType::RowRange => {
            // Row selection doesn't highlight column headers
        }
    }

    (selected, fully_selected)
}

/// Get selected rows with distinction between partial and full selection
fn get_selected_rows_with_type(
    selection: Option<&Selection>,
    layout: &SheetLayout,
) -> (
    std::collections::HashSet<u32>,
    std::collections::HashSet<u32>,
) {
    let mut selected = std::collections::HashSet::new();
    let mut fully_selected = std::collections::HashSet::new();

    let Some(sel) = selection else {
        return (selected, fully_selected);
    };

    match sel.selection_type {
        SelectionType::RowRange => {
            // Entire rows are selected
            let (min_row, _, max_row, _) = sel.bounds();
            let max_row = max_row.min(layout.max_row);
            for row in min_row..=max_row {
                selected.insert(row);
                fully_selected.insert(row);
            }
        }
        SelectionType::All => {
            // All rows are fully selected
            for row in 0..=layout.max_row {
                selected.insert(row);
                fully_selected.insert(row);
            }
        }
        SelectionType::CellRange => {
            // Partial selection - rows contain selection but aren't fully selected
            let (min_row, _, max_row, _) = sel.bounds();
            let max_row = max_row.min(layout.max_row);
            for row in min_row..=max_row {
                selected.insert(row);
            }
        }
        SelectionType::ColumnRange => {
            // Column selection doesn't highlight row headers
        }
    }

    (selected, fully_selected)
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_col_to_letter() {
        assert_eq!(col_to_letter(0), "A");
        assert_eq!(col_to_letter(1), "B");
        assert_eq!(col_to_letter(25), "Z");
        assert_eq!(col_to_letter(26), "AA");
        assert_eq!(col_to_letter(27), "AB");
        assert_eq!(col_to_letter(51), "AZ");
        assert_eq!(col_to_letter(52), "BA");
        assert_eq!(col_to_letter(701), "ZZ");
        assert_eq!(col_to_letter(702), "AAA");
    }
}
