//! Cell indicators (comments, validation, filter buttons, etc.) for Canvas 2D

use crate::cell_ref::parse_cell_range;
use crate::layout::{SheetLayout, Viewport};
use crate::render::CellRenderData;
use crate::types::content::{DataValidationRange, ValidationType};
use crate::types::filter::{AutoFilter, FilterType};
use web_sys::CanvasRenderingContext2d;

/// Render comment indicators (red triangles in top-right corner)
///
/// Note: The canvas context already has DPR scaling applied, so all coordinates
/// are in logical (CSS) pixels.
pub fn render_comment_indicators(
    ctx: &CanvasRenderingContext2d,
    cells: &[CellRenderData],
    layout: &SheetLayout,
    viewport: &Viewport,
    _dpr: f32, // Kept for API compatibility, but not used (context already scaled)
) {
    let triangle_size = 6.0;

    for cell in cells {
        if cell.has_comment != Some(true) {
            continue;
        }

        // Get cell rect from layout
        let rect = layout.cell_rect(cell.row, cell.col);
        if rect.skip || rect.width <= 0.0 || rect.height <= 0.0 {
            continue;
        }

        // Convert to screen coordinates (logical pixels), accounting for frozen panes
        let (sx, sy) = viewport.to_screen_frozen(rect.x, rect.y, cell.row, cell.col, layout);
        let screen_width = rect.width * viewport.scale;

        // Draw red triangle in top-right corner
        let right = f64::from(sx + screen_width);
        let top = f64::from(sy);

        ctx.save();
        ctx.set_fill_style_str("#FF0000"); // Red
        ctx.begin_path();
        ctx.move_to(right - triangle_size, top);
        ctx.line_to(right, top);
        ctx.line_to(right, top + triangle_size);
        ctx.close_path();
        ctx.fill();
        ctx.restore();
    }
}

/// Render data validation dropdown indicators (small triangles on right side of cells)
///
/// For cells with list-type data validation and show_dropdown=true, draws a small
/// downward-pointing gray triangle on the right side of the cell.
///
/// Note: The canvas context already has DPR scaling applied, so all coordinates
/// are in logical (CSS) pixels.
pub fn render_validation_indicators(
    ctx: &CanvasRenderingContext2d,
    data_validations: &[DataValidationRange],
    layout: &SheetLayout,
    viewport: &Viewport,
    content_width: f32,
    content_height: f32,
    _dpr: f32, // Kept for API compatibility, but not used (context already scaled)
) {
    // Arrow dimensions
    let arrow_width = 8.0;
    let arrow_height = 6.0;
    let right_padding = 4.0;

    // Get visible cell range
    let (start_row, end_row) = viewport.visible_rows_in_height(layout, content_height);
    let (start_col, end_col) = viewport.visible_cols_in_width(layout, content_width);

    for validation_range in data_validations {
        // Only render dropdown indicators for list-type validations with show_dropdown=true
        if !matches!(
            validation_range.validation.validation_type,
            ValidationType::List
        ) {
            continue;
        }
        if !validation_range.validation.show_dropdown {
            continue;
        }

        // Parse sqref to get cell ranges (can be multiple ranges separated by space)
        for range_str in validation_range.sqref.split_whitespace() {
            let Some((min_row, min_col, max_row, max_col)) = parse_cell_range(range_str) else {
                continue;
            };

            // Intersect the validation range with the visible range up front
            let row_start = min_row.max(start_row);
            let row_end = max_row.min(end_row);
            let col_start = min_col.max(start_col);
            let col_end = max_col.min(end_col);
            if row_start > row_end || col_start > col_end {
                continue;
            }

            // Iterate over visible cells in the intersected range
            for row in row_start..=row_end {
                for col in col_start..=col_end {
                    // Get cell rect from layout
                    let rect = layout.cell_rect(row, col);
                    if rect.skip || rect.width <= 0.0 || rect.height <= 0.0 {
                        continue;
                    }

                    // Convert to screen coordinates (logical pixels), accounting for frozen panes
                    let (sx, sy) = viewport.to_screen_frozen(rect.x, rect.y, row, col, layout);
                    let screen_width = rect.width * viewport.scale;
                    let screen_height = rect.height * viewport.scale;

                    // Calculate arrow position: right side of cell, vertically centered
                    let arrow_x = f64::from(sx + screen_width) - right_padding - arrow_width;
                    let arrow_y = f64::from(sy) + (f64::from(screen_height) - arrow_height) / 2.0;

                    // Draw downward-pointing triangle
                    ctx.save();
                    ctx.set_fill_style_str("#666666"); // Gray
                    ctx.begin_path();
                    ctx.move_to(arrow_x, arrow_y); // top-left
                    ctx.line_to(arrow_x + arrow_width, arrow_y); // top-right
                    ctx.line_to(arrow_x + arrow_width / 2.0, arrow_y + arrow_height); // bottom-center
                    ctx.close_path();
                    ctx.fill();
                    ctx.restore();
                }
            }
        }
    }
}

/// Render auto-filter dropdown buttons in the header row
///
/// For each column in the AutoFilter range, draws a small dropdown button in the
/// bottom-right corner of the header cell. If a filter is active on the column,
/// the button is shown in blue; otherwise it's shown in gray.
///
/// Note: The canvas context already has DPR scaling applied, so all coordinates
/// are in logical (CSS) pixels.
pub fn render_filter_buttons(
    ctx: &CanvasRenderingContext2d,
    auto_filter: Option<&AutoFilter>,
    layout: &SheetLayout,
    viewport: &Viewport,
    content_width: f32,
    _dpr: f32, // Kept for API compatibility, but not used (context already scaled)
) {
    let Some(filter) = auto_filter else {
        return;
    };

    // Button dimensions
    let button_size = 12.0;
    let arrow_size = 5.0;
    let padding = 2.0;

    // Get visible column range
    let (start_col, end_col) = viewport.visible_cols_in_width(layout, content_width);

    // The header row is the start_row of the filter range
    let header_row = filter.start_row;

    // Iterate over each column in the filter range
    for col_index in filter.start_col..=filter.end_col {
        // Skip if column is not visible
        if col_index < start_col || col_index > end_col {
            continue;
        }

        // Get cell rect for the header cell
        let rect = layout.cell_rect(header_row, col_index);
        if rect.skip || rect.width <= 0.0 || rect.height <= 0.0 {
            continue;
        }

        // Convert to screen coordinates (logical pixels), accounting for frozen panes
        let (sx, sy) = viewport.to_screen_frozen(rect.x, rect.y, header_row, col_index, layout);
        let screen_width = rect.width * viewport.scale;
        let screen_height = rect.height * viewport.scale;

        // Calculate button position: bottom-right corner of cell
        let button_x = f64::from(sx + screen_width) - button_size - padding;
        let button_y = f64::from(sy + screen_height) - button_size - padding;

        // Determine if this column has an active filter
        // Column index within the filter is relative to the filter's start_col
        let relative_col = col_index - filter.start_col;
        let has_active_filter = filter.filter_columns.iter().any(|fc| {
            fc.col_id == relative_col && fc.has_filter && fc.filter_type != FilterType::None
        });

        // Check if the button should be hidden for this column
        let button_hidden = filter
            .filter_columns
            .iter()
            .any(|fc| fc.col_id == relative_col && fc.show_button == Some(false));

        if button_hidden {
            continue;
        }

        // Choose colors based on filter state
        let (bg_color, arrow_color) = if has_active_filter {
            ("#4472C4", "#FFFFFF") // Blue background with white arrow (active filter)
        } else {
            ("#E0E0E0", "#666666") // Light gray background with dark gray arrow (inactive)
        };

        ctx.save();

        // Draw button background (rounded rectangle)
        let radius = 2.0;
        ctx.set_fill_style_str(bg_color);
        ctx.begin_path();
        ctx.move_to(button_x + radius, button_y);
        let _ = ctx.arc_to(
            button_x + button_size,
            button_y,
            button_x + button_size,
            button_y + button_size,
            radius,
        );
        let _ = ctx.arc_to(
            button_x + button_size,
            button_y + button_size,
            button_x,
            button_y + button_size,
            radius,
        );
        let _ = ctx.arc_to(button_x, button_y + button_size, button_x, button_y, radius);
        let _ = ctx.arc_to(button_x, button_y, button_x + button_size, button_y, radius);
        ctx.close_path();
        ctx.fill();

        // Draw dropdown arrow (downward-pointing triangle)
        let arrow_x = button_x + (button_size - arrow_size) / 2.0;
        let arrow_y = button_y + (button_size - arrow_size / 2.0) / 2.0;

        ctx.set_fill_style_str(arrow_color);
        ctx.begin_path();
        ctx.move_to(arrow_x, arrow_y); // top-left
        ctx.line_to(arrow_x + arrow_size, arrow_y); // top-right
        ctx.line_to(arrow_x + arrow_size / 2.0, arrow_y + arrow_size / 2.0 + 1.0); // bottom-center
        ctx.close_path();
        ctx.fill();

        ctx.restore();
    }
}
