//! Frozen panes rendering for Canvas 2D

use crate::layout::{SheetLayout, Viewport};
use web_sys::CanvasRenderingContext2d;

/// Helper to get crisp pixel position for 1px lines (same as renderer)
fn crisp(x: f64) -> f64 {
    x.floor() + 0.5
}

/// Render frozen pane divider lines
///
/// Uses a subtle gray line similar to Excel and Google Sheets.
/// Note: The canvas context already has DPR scaling applied, so all coordinates
/// are in logical (CSS) pixels.
pub fn render_frozen_dividers(
    ctx: &CanvasRenderingContext2d,
    layout: &SheetLayout,
    viewport: &Viewport,
    _dpr: f32, // Kept for API compatibility, but not used (context already scaled)
) {
    let frozen_rows = layout.frozen_rows;
    let frozen_cols = layout.frozen_cols;

    if frozen_rows == 0 && frozen_cols == 0 {
        return;
    }

    // Use a subtle gray color like Excel/Google Sheets
    let divider_color = "#BABABA";

    // Limit divider lines to actual data bounds (not past the sheet content)
    let data_width = f64::from(layout.total_width()).min(f64::from(viewport.width));
    let data_height = f64::from(layout.total_height()).min(f64::from(viewport.height));

    // Draw horizontal divider line (below frozen rows)
    // The frozen height is fixed at the top of the viewport
    if frozen_rows > 0 {
        let frozen_height = layout.frozen_rows_height();
        let y = crisp(f64::from(frozen_height));

        ctx.save();
        ctx.set_stroke_style_str(divider_color);
        ctx.set_line_width(1.0);
        ctx.begin_path();
        ctx.move_to(0.0, y);
        ctx.line_to(data_width, y);
        ctx.stroke();
        ctx.restore();
    }

    // Draw vertical divider line (right of frozen cols)
    // The frozen width is fixed at the left of the viewport
    if frozen_cols > 0 {
        let frozen_width = layout.frozen_cols_width();
        let x = crisp(f64::from(frozen_width));

        ctx.save();
        ctx.set_stroke_style_str(divider_color);
        ctx.set_line_width(1.0);
        ctx.begin_path();
        ctx.move_to(x, 0.0);
        ctx.line_to(x, data_height);
        ctx.stroke();
        ctx.restore();
    }
}
