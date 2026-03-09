//! Sparkline rendering for Canvas 2D backend.

use std::collections::HashMap;

use crate::cell_ref::{parse_cell_range, parse_cell_ref};
use crate::layout::{SheetLayout, Viewport};
use crate::render::backend::CellRenderData;
use crate::types::sparkline::SparklineGroup;

use super::renderer::CanvasRenderer;

impl CanvasRenderer {
    /// Render sparklines for the sheet
    pub(super) fn render_sparklines(
        &self,
        sparkline_groups: &[SparklineGroup],
        cells: &[&CellRenderData],
        layout: &SheetLayout,
        viewport: &Viewport,
    ) {
        if sparkline_groups.is_empty() {
            return;
        }

        let mut cell_map: HashMap<(u32, u32), &CellRenderData> =
            HashMap::with_capacity(cells.len());
        for &cell in cells {
            cell_map.insert((cell.row, cell.col), cell);
        }

        for group in sparkline_groups {
            for sparkline in &group.sparklines {
                // Get sparkline values from the data range
                let values = self.get_sparkline_values(&sparkline.data_range, &cell_map);
                if values.is_empty() {
                    continue;
                }

                // Parse location cell reference (e.g., "B1") to get row, col
                let Some((col, row)) = parse_cell_ref(&sparkline.location) else {
                    continue;
                };

                // Get cell rect for the sparkline location
                let rect = layout.cell_rect(row, col);
                if rect.skip || rect.width <= 0.0 || rect.height <= 0.0 {
                    continue;
                }

                let (sx, sy) = viewport.to_screen_frozen(rect.x, rect.y, row, col, layout);
                let x = f64::from(sx);
                let y = f64::from(sy);
                let w = f64::from(rect.width);
                let h = f64::from(rect.height);

                // Apply some padding within the cell
                let padding = 2.0;
                let plot_x = x + padding;
                let plot_y = y + padding;
                let plot_w = w - padding * 2.0;
                let plot_h = h - padding * 2.0;

                if plot_w <= 0.0 || plot_h <= 0.0 {
                    continue;
                }

                // Render based on sparkline type
                match group.sparkline_type.as_str() {
                    "line" => {
                        self.render_line_sparkline(&values, group, plot_x, plot_y, plot_w, plot_h)
                    }
                    "column" => {
                        self.render_column_sparkline(&values, group, plot_x, plot_y, plot_w, plot_h)
                    }
                    "stacked" => self
                        .render_stacked_sparkline(&values, group, plot_x, plot_y, plot_w, plot_h),
                    _ => self.render_line_sparkline(&values, group, plot_x, plot_y, plot_w, plot_h),
                }
            }
        }
    }

    /// Get numeric values from a sparkline data range
    fn get_sparkline_values(
        &self,
        data_range: &str,
        cells: &HashMap<(u32, u32), &CellRenderData>,
    ) -> Vec<Option<f64>> {
        // Parse the data range (e.g., "Sheet1!A1:A10" or "A1:F1")
        // Remove sheet prefix if present
        let range_part = data_range.split('!').next_back().unwrap_or(data_range);

        let Some(range) = parse_cell_range(range_part) else {
            return Vec::new();
        };

        let (start_row, start_col, end_row, end_col) = range;
        let mut values = Vec::new();

        // Determine if this is a row-wise or column-wise range
        let is_horizontal = end_col > start_col && end_row == start_row;

        if is_horizontal {
            // Iterate columns
            for col in start_col..=end_col {
                let value = cells.get(&(start_row, col)).and_then(|c| c.numeric_value);
                values.push(value);
            }
        } else {
            // Iterate rows
            for row in start_row..=end_row {
                let value = cells.get(&(row, start_col)).and_then(|c| c.numeric_value);
                values.push(value);
            }
        }

        values
    }

    /// Render a line sparkline
    fn render_line_sparkline(
        &self,
        values: &[Option<f64>],
        group: &SparklineGroup,
        x: f64,
        y: f64,
        w: f64,
        h: f64,
    ) {
        // Filter out None values for min/max calculation
        let valid_values: Vec<f64> = values.iter().filter_map(|v| *v).collect();
        if valid_values.is_empty() {
            return;
        }

        let min_val = valid_values.iter().copied().fold(f64::INFINITY, f64::min);
        let max_val = valid_values
            .iter()
            .copied()
            .fold(f64::NEG_INFINITY, f64::max);
        let range = (max_val - min_val).max(0.001); // Avoid division by zero

        // Get line color
        let line_color = group.colors.series.as_deref().unwrap_or("#2E75B6");
        let line_weight = group.line_weight.unwrap_or(0.75);

        // Calculate point positions
        let num_points = values.len();
        let point_spacing = if num_points > 1 {
            w / (num_points - 1) as f64
        } else {
            w
        };

        // Collect points for the line
        let mut points: Vec<(f64, f64, Option<f64>)> = Vec::new();
        for (i, value) in values.iter().enumerate() {
            let px = x + i as f64 * point_spacing;
            if let Some(v) = value {
                let normalized = (v - min_val) / range;
                let py = y + h - normalized * h;
                points.push((px, py, Some(*v)));
            } else {
                points.push((px, y + h / 2.0, None)); // Gap handling
            }
        }

        // Draw the line
        self.ctx.set_stroke_style_str(line_color);
        self.ctx.set_line_width(line_weight);
        self.ctx.begin_path();

        let mut first = true;
        for (px, py, value) in &points {
            if value.is_none() {
                // Gap - start new path segment
                first = true;
                continue;
            }
            if first {
                self.ctx.move_to(*px, *py);
                first = false;
            } else {
                self.ctx.line_to(*px, *py);
            }
        }
        self.ctx.stroke();

        // Draw markers if enabled
        if group.markers {
            let marker_color = group.colors.markers.as_deref().unwrap_or(line_color);
            self.ctx.set_fill_style_str(marker_color);
            for (px, py, value) in &points {
                if value.is_some() {
                    self.ctx.begin_path();
                    let _ = self.ctx.arc(*px, *py, 2.0, 0.0, std::f64::consts::PI * 2.0);
                    self.ctx.fill();
                }
            }
        }

        // Highlight special points
        self.render_sparkline_highlights(&points, &valid_values, group);

        // Draw X-axis if enabled
        if group.display_x_axis && min_val < 0.0 && max_val > 0.0 {
            let zero_y = y + h - ((-min_val) / range) * h;
            let axis_color = group.colors.axis.as_deref().unwrap_or("#000000");
            self.ctx.set_stroke_style_str(axis_color);
            self.ctx.set_line_width(0.5);
            self.ctx.begin_path();
            self.ctx.move_to(x, zero_y);
            self.ctx.line_to(x + w, zero_y);
            self.ctx.stroke();
        }
    }

    /// Render a column sparkline
    fn render_column_sparkline(
        &self,
        values: &[Option<f64>],
        group: &SparklineGroup,
        x: f64,
        y: f64,
        w: f64,
        h: f64,
    ) {
        let valid_values: Vec<f64> = values.iter().filter_map(|v| *v).collect();
        if valid_values.is_empty() {
            return;
        }

        let min_val = valid_values
            .iter()
            .copied()
            .fold(f64::INFINITY, f64::min)
            .min(0.0);
        let max_val = valid_values
            .iter()
            .copied()
            .fold(f64::NEG_INFINITY, f64::max)
            .max(0.0);
        let range = (max_val - min_val).max(0.001);

        let positive_color = group.colors.series.as_deref().unwrap_or("#2E75B6");
        let negative_color = group.colors.negative.as_deref().unwrap_or("#FF0000");

        let num_bars = values.len();
        let bar_gap = 1.0;
        let bar_width = (w - bar_gap * (num_bars - 1) as f64) / num_bars as f64;
        let bar_width = bar_width.max(1.0);

        // Find min/max values for highlighting
        let min_value = valid_values.iter().copied().fold(f64::INFINITY, f64::min);
        let max_value = valid_values
            .iter()
            .copied()
            .fold(f64::NEG_INFINITY, f64::max);

        // Calculate zero line position
        let zero_y = if min_val < 0.0 && max_val > 0.0 {
            y + h * (max_val / range)
        } else if max_val <= 0.0 {
            y
        } else {
            y + h
        };

        for (i, value) in values.iter().enumerate() {
            let Some(v) = value else {
                continue;
            };

            let bar_x = x + i as f64 * (bar_width + bar_gap);
            let bar_h = (v.abs() / range) * h;

            let (bar_y, bar_color) = if *v >= 0.0 {
                (zero_y - bar_h, positive_color)
            } else {
                (zero_y, negative_color)
            };

            // Determine if this bar should be highlighted
            let color = if group.high_point && (*v - max_value).abs() < 0.0001 {
                group.colors.high.as_deref().unwrap_or(positive_color)
            } else if group.low_point && (*v - min_value).abs() < 0.0001 {
                group.colors.low.as_deref().unwrap_or(positive_color)
            } else if group.first_point && i == 0 {
                group.colors.first.as_deref().unwrap_or(bar_color)
            } else if group.last_point && i == num_bars - 1 {
                group.colors.last.as_deref().unwrap_or(bar_color)
            } else if group.negative_points && *v < 0.0 {
                negative_color
            } else {
                bar_color
            };

            self.ctx.set_fill_style_str(color);
            self.ctx.fill_rect(bar_x, bar_y, bar_width, bar_h.max(1.0));
        }

        // Draw X-axis if enabled
        if group.display_x_axis {
            let axis_color = group.colors.axis.as_deref().unwrap_or("#000000");
            self.ctx.set_stroke_style_str(axis_color);
            self.ctx.set_line_width(0.5);
            self.ctx.begin_path();
            self.ctx.move_to(x, zero_y);
            self.ctx.line_to(x + w, zero_y);
            self.ctx.stroke();
        }
    }

    /// Render a stacked/win-loss sparkline
    fn render_stacked_sparkline(
        &self,
        values: &[Option<f64>],
        group: &SparklineGroup,
        x: f64,
        y: f64,
        w: f64,
        h: f64,
    ) {
        // Win/loss sparkline: bars are either above or below the center line
        // All positive bars have the same height, all negative bars have the same height
        let positive_color = group.colors.series.as_deref().unwrap_or("#2E75B6");
        let negative_color = group.colors.negative.as_deref().unwrap_or("#FF0000");

        let num_bars = values.len();
        if num_bars == 0 {
            return;
        }

        let bar_gap = 1.0;
        let bar_width = (w - bar_gap * (num_bars - 1) as f64) / num_bars as f64;
        let bar_width = bar_width.max(1.0);

        // Center line divides the cell in half
        let center_y = y + h / 2.0;
        let bar_h = h / 2.0 - 1.0; // Leave 1px margin from edges

        for (i, value) in values.iter().enumerate() {
            let Some(v) = value else {
                continue;
            };

            let bar_x = x + i as f64 * (bar_width + bar_gap);

            let (bar_y, color) = if *v >= 0.0 {
                (center_y - bar_h, positive_color)
            } else {
                (center_y, negative_color)
            };

            // Apply highlight colors
            let final_color = if group.first_point && i == 0 {
                group.colors.first.as_deref().unwrap_or(color)
            } else if group.last_point && i == num_bars - 1 {
                group.colors.last.as_deref().unwrap_or(color)
            } else if group.negative_points && *v < 0.0 {
                negative_color
            } else {
                color
            };

            self.ctx.set_fill_style_str(final_color);
            self.ctx.fill_rect(bar_x, bar_y, bar_width, bar_h);
        }

        // Draw center axis line
        if group.display_x_axis {
            let axis_color = group.colors.axis.as_deref().unwrap_or("#000000");
            self.ctx.set_stroke_style_str(axis_color);
            self.ctx.set_line_width(0.5);
            self.ctx.begin_path();
            self.ctx.move_to(x, center_y);
            self.ctx.line_to(x + w, center_y);
            self.ctx.stroke();
        }
    }

    /// Render highlight markers for special points in line sparklines
    fn render_sparkline_highlights(
        &self,
        points: &[(f64, f64, Option<f64>)],
        valid_values: &[f64],
        group: &SparklineGroup,
    ) {
        if valid_values.is_empty() {
            return;
        }

        let min_value = valid_values.iter().copied().fold(f64::INFINITY, f64::min);
        let max_value = valid_values
            .iter()
            .copied()
            .fold(f64::NEG_INFINITY, f64::max);

        let num_points = points.len();

        for (i, (px, py, value)) in points.iter().enumerate() {
            let Some(v) = value else {
                continue;
            };

            let mut should_highlight = false;
            let mut color = "#000000";

            if group.high_point && (*v - max_value).abs() < 0.0001 {
                should_highlight = true;
                color = group.colors.high.as_deref().unwrap_or("#00B050");
            } else if group.low_point && (*v - min_value).abs() < 0.0001 {
                should_highlight = true;
                color = group.colors.low.as_deref().unwrap_or("#FF0000");
            } else if group.first_point && i == 0 {
                should_highlight = true;
                color = group.colors.first.as_deref().unwrap_or("#FFBF00");
            } else if group.last_point && i == num_points - 1 {
                should_highlight = true;
                color = group.colors.last.as_deref().unwrap_or("#FFBF00");
            } else if group.negative_points && *v < 0.0 {
                should_highlight = true;
                color = group.colors.negative.as_deref().unwrap_or("#FF0000");
            }

            if should_highlight {
                self.ctx.set_fill_style_str(color);
                self.ctx.begin_path();
                let _ = self.ctx.arc(*px, *py, 3.0, 0.0, std::f64::consts::PI * 2.0);
                self.ctx.fill();
            }
        }
    }
}
