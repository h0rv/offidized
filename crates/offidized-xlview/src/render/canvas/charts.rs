//! Chart rendering for Canvas 2D backend.

use crate::layout::{SheetLayout, Viewport};
use crate::types::chart::{BarDirection, Chart, ChartType};

use super::renderer::{CanvasRenderer, SCROLLBAR_SIZE};

impl CanvasRenderer {
    /// Chart color palette
    const CHART_COLORS: [&'static str; 5] = ["#4472C4", "#ED7D31", "#A5A5A5", "#FFC000", "#5B9BD5"];
    const CHART_PADDING: f64 = 10.0;
    const CHART_AXIS_LEFT_GUTTER: f64 = 36.0;
    const CHART_AXIS_RIGHT_GUTTER: f64 = 6.0;
    const CHART_AXIS_BOTTOM_GUTTER: f64 = 24.0;
    const CHART_MIN_LABEL_SPACING: f64 = 56.0;

    /// Render charts embedded in the sheet
    pub(super) fn render_charts(
        &self,
        charts: &[Chart],
        layout: &SheetLayout,
        viewport: &Viewport,
    ) {
        for chart in charts {
            self.render_chart(chart, layout, viewport);
        }
    }

    /// Render a single chart
    /// EMU-to-pixel conversion factor (1 EMU = 1/914400 inch, at 96 DPI).
    const EMU_TO_PX: f64 = 1.0 / 9525.0;

    #[allow(clippy::cast_possible_truncation)]
    fn render_chart(&self, chart: &Chart, layout: &SheetLayout, viewport: &Viewport) {
        // Calculate chart bounds from anchor positions
        let from_col = chart.from_col.unwrap_or(0);
        let from_row = chart.from_row.unwrap_or(0);

        // Get sheet coordinates for the starting anchor
        let x1 = layout
            .col_positions
            .get(from_col as usize)
            .copied()
            .unwrap_or(0.0);
        let y1 = layout
            .row_positions
            .get(from_row as usize)
            .copied()
            .unwrap_or(0.0);

        let has_extent = matches!(
            (chart.extent_cx, chart.extent_cy),
            (Some(cx), Some(cy)) if cx > 0 && cy > 0
        );
        let has_two_cell_anchor = chart.to_col.is_some() || chart.to_row.is_some();

        // Only use extent-based sizing for real one-cell anchors. Two-cell anchors
        // should size from their `from`/`to` bounds even if the drawing XML also
        // contains a nested graphic-frame transform extent.
        let (width, height) = if has_extent && !has_two_cell_anchor {
            let (cx, cy) = (chart.extent_cx.unwrap_or(0), chart.extent_cy.unwrap_or(0));
            #[allow(clippy::cast_possible_truncation)]
            let w = (cx as f64 * Self::EMU_TO_PX).max(100.0) as f32;
            #[allow(clippy::cast_possible_truncation)]
            let h = (cy as f64 * Self::EMU_TO_PX).max(80.0) as f32;
            (w, h)
        } else {
            let to_col = chart.to_col.unwrap_or(from_col + 8);
            let to_row = chart.to_row.unwrap_or(from_row + 15);
            let x2 = layout
                .col_positions
                .get(to_col as usize)
                .copied()
                .unwrap_or(x1 + 400.0);
            let y2 = layout
                .row_positions
                .get(to_row as usize)
                .copied()
                .unwrap_or(y1 + 300.0);
            ((x2 - x1).max(100.0), (y2 - y1).max(80.0))
        };

        // Convert to screen coordinates
        let (screen_x, screen_y) = viewport.to_screen_frozen(x1, y1, from_row, from_col, layout);
        let screen_width = width * viewport.scale;
        let screen_height = height * viewport.scale;

        // Skip if chart is completely off-screen
        let content_width = f64::from(viewport.width) - SCROLLBAR_SIZE;
        let content_height = f64::from(viewport.height) - SCROLLBAR_SIZE;
        if f64::from(screen_x) > content_width
            || f64::from(screen_y) > content_height
            || f64::from(screen_x + screen_width) < 0.0
            || f64::from(screen_y + screen_height) < 0.0
        {
            return;
        }

        let x = f64::from(screen_x);
        let y = f64::from(screen_y);
        let w = f64::from(screen_width);
        let h = f64::from(screen_height);

        // Draw chart background
        self.ctx.set_fill_style_str("#FFFFFF");
        self.ctx.fill_rect(x, y, w, h);

        // Draw chart border
        self.ctx.set_stroke_style_str("#D9D9D9");
        self.ctx.set_line_width(1.0);
        self.ctx.stroke_rect(x, y, w, h);

        // Reserve space for title and legend
        let title_height = if chart.title.is_some() { 24.0 } else { 8.0 };
        let legend_height = if chart.legend.is_some() { 20.0 } else { 8.0 };
        let padding = Self::CHART_PADDING;
        let (axis_left_gutter, axis_right_gutter, axis_bottom_gutter) = match chart.chart_type {
            ChartType::Pie | ChartType::Doughnut => (0.0, 0.0, 0.0),
            _ => (
                Self::CHART_AXIS_LEFT_GUTTER,
                Self::CHART_AXIS_RIGHT_GUTTER,
                Self::CHART_AXIS_BOTTOM_GUTTER,
            ),
        };

        // Calculate plot area
        let plot_x = x + padding + axis_left_gutter;
        let plot_y = y + title_height;
        let plot_w = (w - padding * 2.0 - axis_left_gutter - axis_right_gutter).max(40.0);
        let plot_h = (h - title_height - legend_height - padding - axis_bottom_gutter).max(30.0);

        // Draw title if present
        if let Some(ref title) = chart.title {
            self.ctx.set_font("bold 12px Calibri, Arial, sans-serif");
            self.ctx.set_fill_style_str("#000000");
            self.ctx.set_text_align("center");
            let _ = self.ctx.fill_text(title, x + w / 2.0, y + 16.0);
            self.ctx.set_text_align("left");
        }

        // Render gridlines and axes first (for non-pie charts)
        match chart.chart_type {
            ChartType::Pie | ChartType::Doughnut => {}
            _ => self.render_chart_axes(chart, plot_x, plot_y, plot_w, plot_h),
        }

        // Render chart based on type
        match chart.chart_type {
            ChartType::Bar => self.render_bar_chart(chart, plot_x, plot_y, plot_w, plot_h),
            ChartType::Line => self.render_line_chart(chart, plot_x, plot_y, plot_w, plot_h),
            ChartType::Pie | ChartType::Doughnut => {
                self.render_pie_chart(chart, plot_x, plot_y, plot_w, plot_h)
            }
            ChartType::Area => self.render_area_chart(chart, plot_x, plot_y, plot_w, plot_h),
            ChartType::Scatter => self.render_scatter_chart(chart, plot_x, plot_y, plot_w, plot_h),
            ChartType::Bubble => self.render_bubble_chart(chart, plot_x, plot_y, plot_w, plot_h),
            ChartType::Radar => self.render_radar_chart(chart, plot_x, plot_y, plot_w, plot_h),
            ChartType::Stock => self.render_stock_chart(chart, plot_x, plot_y, plot_w, plot_h),
            ChartType::Surface => self.render_surface_chart(chart, plot_x, plot_y, plot_w, plot_h),
            ChartType::Combo => self.render_combo_chart(chart, plot_x, plot_y, plot_w, plot_h),
        }

        // Draw legend if present
        if let Some(ref legend) = chart.legend {
            self.render_chart_legend(chart, legend, x, y + h - legend_height, w, legend_height);
        }
    }

    /// Render a bar/column chart
    #[allow(clippy::indexing_slicing)] // Safe: modulo ensures index is within bounds
    fn render_bar_chart(&self, chart: &Chart, x: f64, y: f64, w: f64, h: f64) {
        let is_horizontal = chart.bar_direction == Some(BarDirection::Bar);

        // Collect all values to determine scale
        let mut all_values: Vec<f64> = Vec::new();
        for series in &chart.series {
            if let Some(ref values) = series.values {
                for v in values.num_values.iter().flatten() {
                    all_values.push(*v);
                }
            }
        }

        if all_values.is_empty() {
            return;
        }

        let min_val = all_values.iter().copied().fold(0.0_f64, f64::min);
        let max_val = all_values.iter().copied().fold(0.0_f64, f64::max);
        let range = (max_val - min_val).max(1.0);

        // Count data points (use first series as reference)
        let num_points = chart
            .series
            .first()
            .and_then(|s| s.values.as_ref())
            .map(|v| v.num_values.len())
            .unwrap_or(0);

        if num_points == 0 {
            return;
        }

        let num_series = chart.series.len();
        let group_gap = 0.2; // Gap between groups as fraction of group width
        let bar_gap = 0.05; // Gap between bars within group

        if is_horizontal {
            // Horizontal bars
            let group_height = h / num_points as f64;
            let bar_height = group_height * (1.0 - group_gap) / num_series as f64 * (1.0 - bar_gap);

            for (series_idx, series) in chart.series.iter().enumerate() {
                let color = Self::CHART_COLORS[series_idx % Self::CHART_COLORS.len()];
                self.ctx.set_fill_style_str(color);

                if let Some(ref values) = series.values {
                    for (point_idx, val) in values.num_values.iter().enumerate() {
                        if let Some(v) = val {
                            let bar_w = ((v - min_val) / range * w).max(0.0);
                            let bar_y = y
                                + point_idx as f64 * group_height
                                + group_height * group_gap / 2.0
                                + series_idx as f64 * (bar_height + bar_height * bar_gap);

                            self.ctx.fill_rect(x, bar_y, bar_w, bar_height);
                        }
                    }
                }
            }
        } else {
            // Vertical bars (column chart)
            let group_width = w / num_points as f64;
            let bar_width = group_width * (1.0 - group_gap) / num_series as f64 * (1.0 - bar_gap);

            for (series_idx, series) in chart.series.iter().enumerate() {
                let color = Self::CHART_COLORS[series_idx % Self::CHART_COLORS.len()];
                self.ctx.set_fill_style_str(color);

                if let Some(ref values) = series.values {
                    for (point_idx, val) in values.num_values.iter().enumerate() {
                        if let Some(v) = val {
                            let bar_h = ((v - min_val) / range * h).max(0.0);
                            let bar_x = x
                                + point_idx as f64 * group_width
                                + group_width * group_gap / 2.0
                                + series_idx as f64 * (bar_width + bar_width * bar_gap);
                            let bar_y = y + h - bar_h;

                            self.ctx.fill_rect(bar_x, bar_y, bar_width, bar_h);
                        }
                    }
                }
            }
        }

        // Draw axis lines
        self.ctx.set_stroke_style_str("#808080");
        self.ctx.set_line_width(1.0);
        self.ctx.begin_path();
        // Y-axis
        self.ctx.move_to(Self::crisp(x), y);
        self.ctx.line_to(Self::crisp(x), y + h);
        // X-axis
        self.ctx.move_to(x, Self::crisp(y + h));
        self.ctx.line_to(x + w, Self::crisp(y + h));
        self.ctx.stroke();
    }

    /// Render a line chart
    #[allow(clippy::indexing_slicing)] // Safe: modulo ensures index is within bounds
    fn render_line_chart(&self, chart: &Chart, x: f64, y: f64, w: f64, h: f64) {
        // Collect all values to determine scale
        let mut all_values: Vec<f64> = Vec::new();
        for series in &chart.series {
            if let Some(ref values) = series.values {
                for v in values.num_values.iter().flatten() {
                    all_values.push(*v);
                }
            }
        }

        if all_values.is_empty() {
            return;
        }

        let min_val = all_values.iter().copied().fold(f64::INFINITY, f64::min);
        let max_val = all_values.iter().copied().fold(f64::NEG_INFINITY, f64::max);
        let range = (max_val - min_val).max(1.0);

        // Draw each series as a line
        for (series_idx, series) in chart.series.iter().enumerate() {
            let color = Self::CHART_COLORS[series_idx % Self::CHART_COLORS.len()];
            self.ctx.set_stroke_style_str(color);
            self.ctx.set_line_width(2.0);
            self.ctx.begin_path();

            if let Some(ref values) = series.values {
                let num_points = values.num_values.len();
                if num_points == 0 {
                    continue;
                }

                let point_spacing = if num_points > 1 {
                    w / (num_points - 1) as f64
                } else {
                    w
                };
                let mut first = true;

                for (point_idx, val) in values.num_values.iter().enumerate() {
                    if let Some(v) = val {
                        let px = x + point_idx as f64 * point_spacing;
                        let py = y + h - ((v - min_val) / range * h);

                        if first {
                            self.ctx.move_to(px, py);
                            first = false;
                        } else {
                            self.ctx.line_to(px, py);
                        }
                    }
                }
                self.ctx.stroke();

                // Draw data points (circles)
                self.ctx.set_fill_style_str(color);
                for (point_idx, val) in values.num_values.iter().enumerate() {
                    if let Some(v) = val {
                        let px = x + point_idx as f64 * point_spacing;
                        let py = y + h - ((v - min_val) / range * h);

                        self.ctx.begin_path();
                        let _ = self.ctx.arc(px, py, 3.0, 0.0, std::f64::consts::PI * 2.0);
                        self.ctx.fill();
                    }
                }
            }
        }

        // Draw axis lines
        self.ctx.set_stroke_style_str("#808080");
        self.ctx.set_line_width(1.0);
        self.ctx.begin_path();
        // Y-axis
        self.ctx.move_to(Self::crisp(x), y);
        self.ctx.line_to(Self::crisp(x), y + h);
        // X-axis
        self.ctx.move_to(x, Self::crisp(y + h));
        self.ctx.line_to(x + w, Self::crisp(y + h));
        self.ctx.stroke();
    }

    fn category_label_stride(num_labels: usize, plot_width: f64) -> usize {
        if num_labels <= 1 {
            return 1;
        }

        let max_labels = ((plot_width / Self::CHART_MIN_LABEL_SPACING).floor() as usize).max(2);
        num_labels.div_ceil(max_labels).max(1)
    }

    fn compact_category_label(label: &str) -> String {
        let bytes = label.as_bytes();
        if bytes.len() == 10
            && bytes.get(4) == Some(&b'-')
            && bytes.get(7) == Some(&b'-')
            && bytes
                .iter()
                .enumerate()
                .all(|(idx, byte)| idx == 4 || idx == 7 || byte.is_ascii_digit())
        {
            return label.get(5..).unwrap_or(label).to_string();
        }

        label.to_string()
    }

    /// Render a pie chart (or doughnut)
    #[allow(clippy::indexing_slicing)] // Safe: modulo ensures index is within bounds
    fn render_pie_chart(&self, chart: &Chart, x: f64, y: f64, w: f64, h: f64) {
        // Use first series for pie chart
        let Some(series) = chart.series.first() else {
            return;
        };
        let Some(ref values) = series.values else {
            return;
        };

        // Filter out None and negative values, collect positive values
        let data: Vec<f64> = values
            .num_values
            .iter()
            .filter_map(|v| v.filter(|&x| x > 0.0))
            .collect();

        if data.is_empty() {
            return;
        }

        let total: f64 = data.iter().sum();
        if total <= 0.0 {
            return;
        }

        // Calculate center and radius
        let center_x = x + w / 2.0;
        let center_y = y + h / 2.0;
        let radius = (w.min(h) / 2.0) * 0.85; // Leave some margin

        // For doughnut, create inner radius
        let inner_radius = if chart.chart_type == ChartType::Doughnut {
            radius * 0.5
        } else {
            0.0
        };

        let mut start_angle = -std::f64::consts::FRAC_PI_2; // Start at top

        for (idx, &value) in data.iter().enumerate() {
            let sweep_angle = (value / total) * std::f64::consts::PI * 2.0;
            let end_angle = start_angle + sweep_angle;

            let color = Self::CHART_COLORS[idx % Self::CHART_COLORS.len()];
            self.ctx.set_fill_style_str(color);

            self.ctx.begin_path();

            if inner_radius > 0.0 {
                // Doughnut: draw arc with hole
                let _ = self
                    .ctx
                    .arc(center_x, center_y, radius, start_angle, end_angle);
                let _ = self
                    .ctx
                    .arc(center_x, center_y, inner_radius, end_angle, start_angle);
                self.ctx.close_path();
            } else {
                // Pie: draw slice from center
                self.ctx.move_to(center_x, center_y);
                let _ = self
                    .ctx
                    .arc(center_x, center_y, radius, start_angle, end_angle);
                self.ctx.close_path();
            }

            self.ctx.fill();

            // Draw slice border
            self.ctx.set_stroke_style_str("#FFFFFF");
            self.ctx.set_line_width(1.0);
            self.ctx.stroke();

            start_angle = end_angle;
        }
    }

    /// Render an area chart
    #[allow(clippy::indexing_slicing)] // Safe: modulo ensures index is within bounds
    fn render_area_chart(&self, chart: &Chart, x: f64, y: f64, w: f64, h: f64) {
        // Collect all values to determine scale
        let mut all_values: Vec<f64> = Vec::new();
        for series in &chart.series {
            if let Some(ref values) = series.values {
                for v in values.num_values.iter().flatten() {
                    all_values.push(*v);
                }
            }
        }

        if all_values.is_empty() {
            return;
        }

        let min_val = all_values
            .iter()
            .copied()
            .fold(f64::INFINITY, f64::min)
            .min(0.0);
        let max_val = all_values.iter().copied().fold(f64::NEG_INFINITY, f64::max);
        let range = (max_val - min_val).max(1.0);

        // Draw each series as a filled area
        for (series_idx, series) in chart.series.iter().enumerate() {
            let color = Self::CHART_COLORS[series_idx % Self::CHART_COLORS.len()];

            if let Some(ref values) = series.values {
                let num_points = values.num_values.len();
                if num_points == 0 {
                    continue;
                }

                let point_spacing = if num_points > 1 {
                    w / (num_points - 1) as f64
                } else {
                    w
                };

                // Draw filled area
                self.ctx.begin_path();
                self.ctx.move_to(x, y + h); // Start at bottom-left

                for (point_idx, val) in values.num_values.iter().enumerate() {
                    let v = val.unwrap_or(0.0);
                    let px = x + point_idx as f64 * point_spacing;
                    let py = y + h - ((v - min_val) / range * h);
                    self.ctx.line_to(px, py);
                }

                // Close path back to bottom
                self.ctx
                    .line_to(x + (num_points - 1) as f64 * point_spacing, y + h);
                self.ctx.close_path();

                // Fill with semi-transparent color
                self.ctx.set_fill_style_str(&format!("{}80", color)); // Add 50% alpha
                self.ctx.fill();

                // Draw line on top
                self.ctx.set_stroke_style_str(color);
                self.ctx.set_line_width(2.0);
                self.ctx.begin_path();
                let mut first = true;
                for (point_idx, val) in values.num_values.iter().enumerate() {
                    let v = val.unwrap_or(0.0);
                    let px = x + point_idx as f64 * point_spacing;
                    let py = y + h - ((v - min_val) / range * h);

                    if first {
                        self.ctx.move_to(px, py);
                        first = false;
                    } else {
                        self.ctx.line_to(px, py);
                    }
                }
                self.ctx.stroke();
            }
        }

        // Draw axis lines
        self.ctx.set_stroke_style_str("#808080");
        self.ctx.set_line_width(1.0);
        self.ctx.begin_path();
        self.ctx.move_to(Self::crisp(x), y);
        self.ctx.line_to(Self::crisp(x), y + h);
        self.ctx.move_to(x, Self::crisp(y + h));
        self.ctx.line_to(x + w, Self::crisp(y + h));
        self.ctx.stroke();
    }

    /// Render a placeholder for unsupported chart types
    /// Render chart axes with gridlines and labels
    fn render_chart_axes(&self, chart: &Chart, x: f64, y: f64, w: f64, h: f64) {
        // Draw gridlines
        self.ctx.set_stroke_style_str("#E0E0E0");
        self.ctx.set_line_width(0.5);

        // Find value axis for scale
        let value_axis = chart.axes.iter().find(|a| a.axis_type == "val");

        // Determine grid line count (default 5 major gridlines)
        let num_gridlines = 5;

        // Draw horizontal gridlines
        for i in 0..=num_gridlines {
            let grid_y = y + (h * i as f64 / num_gridlines as f64);
            self.ctx.begin_path();
            self.ctx.move_to(x, grid_y);
            self.ctx.line_to(x + w, grid_y);
            self.ctx.stroke();
        }

        // Draw vertical gridlines for category axis
        let num_points = chart
            .series
            .first()
            .and_then(|s| s.values.as_ref())
            .map(|v| v.num_values.len())
            .unwrap_or(5);
        let category_count = chart
            .series
            .first()
            .and_then(|s| s.categories.as_ref())
            .map(|c| c.str_values.len())
            .unwrap_or(num_points);
        let label_stride = Self::category_label_stride(category_count.max(num_points), w);

        if num_points > 1 {
            let divisor = match chart.chart_type {
                ChartType::Line | ChartType::Area => (num_points - 1) as f64,
                _ => num_points as f64,
            };

            for i in (0..num_points).step_by(label_stride) {
                let grid_x = x + (w * i as f64 / divisor.max(1.0));
                self.ctx.begin_path();
                self.ctx.move_to(grid_x, y);
                self.ctx.line_to(grid_x, y + h);
                self.ctx.stroke();
            }

            if matches!(chart.chart_type, ChartType::Line | ChartType::Area) {
                self.ctx.begin_path();
                self.ctx.move_to(x + w, y);
                self.ctx.line_to(x + w, y + h);
                self.ctx.stroke();
            }
        }

        // Draw axis lines
        self.ctx.set_stroke_style_str("#808080");
        self.ctx.set_line_width(1.0);

        // Y-axis (left)
        self.ctx.begin_path();
        self.ctx.move_to(x, y);
        self.ctx.line_to(x, y + h);
        self.ctx.stroke();

        // X-axis (bottom)
        self.ctx.begin_path();
        self.ctx.move_to(x, y + h);
        self.ctx.line_to(x + w, y + h);
        self.ctx.stroke();

        // Draw axis labels
        self.ctx.set_font("9px Calibri, Arial, sans-serif");
        self.ctx.set_fill_style_str("#606060");

        // Y-axis labels (values)
        if let Some(axis) = value_axis {
            // Collect all values for scale
            let mut all_values: Vec<f64> = Vec::new();
            for series in &chart.series {
                if let Some(ref values) = series.values {
                    for v in values.num_values.iter().flatten() {
                        all_values.push(*v);
                    }
                }
            }

            let min_val = axis
                .min
                .unwrap_or_else(|| all_values.iter().copied().fold(0.0_f64, f64::min));
            let max_val = axis
                .max
                .unwrap_or_else(|| all_values.iter().copied().fold(0.0_f64, f64::max));

            self.ctx.set_text_align("right");
            for i in 0..=num_gridlines {
                let val = max_val - (max_val - min_val) * i as f64 / num_gridlines as f64;
                let label = if val.abs() >= 1000.0 {
                    format!("{:.0}", val)
                } else if val.abs() >= 1.0 {
                    format!("{:.1}", val)
                } else {
                    format!("{:.2}", val)
                };
                let label_y = y + (h * i as f64 / num_gridlines as f64) + 3.0;
                let _ = self.ctx.fill_text(&label, x - 4.0, label_y);
            }
        }

        // X-axis labels (categories)
        self.ctx.set_text_align("center");
        if let Some(series) = chart.series.first() {
            if let Some(ref categories) = series.categories {
                let last_idx = categories.str_values.len().saturating_sub(1);
                for (i, cat) in categories.str_values.iter().enumerate() {
                    if num_points > 0 && (i % label_stride == 0 || i == last_idx) {
                        let label_x = match chart.chart_type {
                            ChartType::Line | ChartType::Area => {
                                if num_points > 1 {
                                    x + (w * i as f64 / (num_points - 1) as f64)
                                } else {
                                    x + w / 2.0
                                }
                            }
                            _ => x + (w * (i as f64 + 0.5) / num_points as f64),
                        };
                        let label = Self::compact_category_label(cat);
                        let _ = self.ctx.fill_text(&label, label_x, y + h + 14.0);
                    }
                }
            }
        }

        // Draw axis titles if present
        self.ctx.set_font("10px Calibri, Arial, sans-serif");
        self.ctx.set_fill_style_str("#404040");

        for axis in &chart.axes {
            if let Some(ref title) = axis.title {
                match axis.position.as_deref() {
                    Some("l") | Some("r") => {
                        // Y-axis title (rotated) - simplified: just draw at left
                        self.ctx.save();
                        self.ctx.translate(x - 30.0, y + h / 2.0).ok();
                        self.ctx.rotate(-std::f64::consts::FRAC_PI_2).ok();
                        self.ctx.set_text_align("center");
                        let _ = self.ctx.fill_text(title, 0.0, 0.0);
                        self.ctx.restore();
                    }
                    _ => {
                        // X-axis title
                        self.ctx.set_text_align("center");
                        let _ = self.ctx.fill_text(title, x + w / 2.0, y + h + 24.0);
                    }
                }
            }
        }

        self.ctx.set_text_align("left");
    }

    /// Render a scatter chart (XY plot)
    #[allow(clippy::indexing_slicing)]
    fn render_scatter_chart(&self, chart: &Chart, x: f64, y: f64, w: f64, h: f64) {
        // Collect all X and Y values to determine scale
        let mut all_x: Vec<f64> = Vec::new();
        let mut all_y: Vec<f64> = Vec::new();

        for series in &chart.series {
            if let Some(ref x_vals) = series.x_values {
                for v in x_vals.num_values.iter().flatten() {
                    all_x.push(*v);
                }
            }
            if let Some(ref y_vals) = series.values {
                for v in y_vals.num_values.iter().flatten() {
                    all_y.push(*v);
                }
            }
        }

        if all_x.is_empty() || all_y.is_empty() {
            // Fall back to index-based X if no X values
            for series in &chart.series {
                if let Some(ref y_vals) = series.values {
                    for (i, v) in y_vals.num_values.iter().enumerate() {
                        if v.is_some() {
                            all_x.push(i as f64);
                        }
                    }
                }
            }
        }

        if all_y.is_empty() {
            return;
        }

        let x_min = all_x.iter().copied().fold(f64::INFINITY, f64::min);
        let x_max = all_x.iter().copied().fold(f64::NEG_INFINITY, f64::max);
        let y_min = all_y.iter().copied().fold(0.0_f64, f64::min);
        let y_max = all_y.iter().copied().fold(0.0_f64, f64::max);

        let x_range = (x_max - x_min).max(1.0);
        let y_range = (y_max - y_min).max(1.0);

        // Draw each series
        for (series_idx, series) in chart.series.iter().enumerate() {
            let color = Self::CHART_COLORS[series_idx % Self::CHART_COLORS.len()];

            let y_vals = match &series.values {
                Some(v) => &v.num_values,
                None => continue,
            };

            // Get X values or generate indices
            let x_vals: Vec<f64> = if let Some(ref xv) = series.x_values {
                xv.num_values
                    .iter()
                    .enumerate()
                    .map(|(i, v)| v.unwrap_or(i as f64))
                    .collect()
            } else {
                (0..y_vals.len()).map(|i| i as f64).collect()
            };

            // Collect points
            let mut points: Vec<(f64, f64)> = Vec::new();
            for (i, y_val) in y_vals.iter().enumerate() {
                if let Some(yv) = y_val {
                    let xv = x_vals.get(i).copied().unwrap_or(i as f64);
                    let px = x + ((xv - x_min) / x_range) * w;
                    let py = y + h - ((yv - y_min) / y_range) * h;
                    points.push((px, py));
                }
            }

            // Draw connecting lines (if scatter style includes lines)
            if points.len() > 1 {
                self.ctx.set_stroke_style_str(color);
                self.ctx.set_line_width(1.5);
                self.ctx.begin_path();
                if let Some((first_x, first_y)) = points.first() {
                    self.ctx.move_to(*first_x, *first_y);
                    for (px, py) in points.iter().skip(1) {
                        self.ctx.line_to(*px, *py);
                    }
                }
                self.ctx.stroke();
            }

            // Draw markers
            self.ctx.set_fill_style_str(color);
            for (px, py) in &points {
                self.ctx.begin_path();
                let _ = self.ctx.arc(*px, *py, 4.0, 0.0, std::f64::consts::PI * 2.0);
                self.ctx.fill();
            }
        }
    }

    /// Render a bubble chart
    #[allow(clippy::indexing_slicing)]
    fn render_bubble_chart(&self, chart: &Chart, x: f64, y: f64, w: f64, h: f64) {
        // Collect all X, Y, and bubble size values
        let mut all_x: Vec<f64> = Vec::new();
        let mut all_y: Vec<f64> = Vec::new();
        let mut all_sizes: Vec<f64> = Vec::new();

        for series in &chart.series {
            if let Some(ref x_vals) = series.x_values {
                for v in x_vals.num_values.iter().flatten() {
                    all_x.push(*v);
                }
            }
            if let Some(ref y_vals) = series.values {
                for v in y_vals.num_values.iter().flatten() {
                    all_y.push(*v);
                }
            }
            if let Some(ref sizes) = series.bubble_sizes {
                for v in sizes.num_values.iter().flatten() {
                    all_sizes.push(*v);
                }
            }
        }

        // Fall back to index-based X if no X values
        if all_x.is_empty() {
            for series in &chart.series {
                if let Some(ref y_vals) = series.values {
                    for (i, v) in y_vals.num_values.iter().enumerate() {
                        if v.is_some() {
                            all_x.push(i as f64);
                        }
                    }
                }
            }
        }

        if all_y.is_empty() {
            return;
        }

        let x_min = all_x.iter().copied().fold(f64::INFINITY, f64::min);
        let x_max = all_x.iter().copied().fold(f64::NEG_INFINITY, f64::max);
        let y_min = all_y.iter().copied().fold(0.0_f64, f64::min);
        let y_max = all_y.iter().copied().fold(0.0_f64, f64::max);
        let size_max = all_sizes.iter().copied().fold(1.0_f64, f64::max);

        let x_range = (x_max - x_min).max(1.0);
        let y_range = (y_max - y_min).max(1.0);

        // Max bubble radius as fraction of chart size
        let max_radius = (w.min(h) * 0.15).min(40.0);

        // Draw each series
        for (series_idx, series) in chart.series.iter().enumerate() {
            let color = Self::CHART_COLORS[series_idx % Self::CHART_COLORS.len()];

            let y_vals = match &series.values {
                Some(v) => &v.num_values,
                None => continue,
            };

            // Get X values or generate indices
            let x_vals: Vec<f64> = if let Some(ref xv) = series.x_values {
                xv.num_values
                    .iter()
                    .enumerate()
                    .map(|(i, v)| v.unwrap_or(i as f64))
                    .collect()
            } else {
                (0..y_vals.len()).map(|i| i as f64).collect()
            };

            // Get bubble sizes or use constant size
            let sizes: Vec<f64> = if let Some(ref sv) = series.bubble_sizes {
                sv.num_values.iter().map(|v| v.unwrap_or(1.0)).collect()
            } else {
                vec![1.0; y_vals.len()]
            };

            // Draw bubbles
            self.ctx.set_fill_style_str(color);
            self.ctx.set_global_alpha(0.6);

            for (i, y_val) in y_vals.iter().enumerate() {
                if let Some(yv) = y_val {
                    let xv = x_vals.get(i).copied().unwrap_or(i as f64);
                    let size = sizes.get(i).copied().unwrap_or(1.0);

                    let px = x + ((xv - x_min) / x_range) * w;
                    let py = y + h - ((yv - y_min) / y_range) * h;
                    let radius = (size / size_max).sqrt() * max_radius;
                    let radius = radius.max(3.0); // Minimum visible size

                    self.ctx.begin_path();
                    let _ = self
                        .ctx
                        .arc(px, py, radius, 0.0, std::f64::consts::PI * 2.0);
                    self.ctx.fill();
                }
            }

            self.ctx.set_global_alpha(1.0);

            // Draw bubble borders
            self.ctx.set_stroke_style_str(color);
            self.ctx.set_line_width(1.0);

            for (i, y_val) in y_vals.iter().enumerate() {
                if let Some(yv) = y_val {
                    let xv = x_vals.get(i).copied().unwrap_or(i as f64);
                    let size = sizes.get(i).copied().unwrap_or(1.0);

                    let px = x + ((xv - x_min) / x_range) * w;
                    let py = y + h - ((yv - y_min) / y_range) * h;
                    let radius = (size / size_max).sqrt() * max_radius;
                    let radius = radius.max(3.0);

                    self.ctx.begin_path();
                    let _ = self
                        .ctx
                        .arc(px, py, radius, 0.0, std::f64::consts::PI * 2.0);
                    self.ctx.stroke();
                }
            }
        }
    }

    /// Render a radar/spider chart
    /// Polygonal chart with axes radiating from center
    #[allow(clippy::indexing_slicing)]
    fn render_radar_chart(&self, chart: &Chart, x: f64, y: f64, w: f64, h: f64) {
        // Find the maximum value across all series for scaling
        let mut max_val = 0.0_f64;
        let mut num_axes = 0;

        for series in &chart.series {
            if let Some(ref values) = series.values {
                num_axes = num_axes.max(values.num_values.len());
                for v in values.num_values.iter().flatten() {
                    max_val = max_val.max(*v);
                }
            }
        }

        if num_axes == 0 || max_val == 0.0 {
            return;
        }

        let center_x = x + w / 2.0;
        let center_y = y + h / 2.0;
        let radius = (w.min(h) / 2.0) * 0.85;
        let angle_step = std::f64::consts::PI * 2.0 / num_axes as f64;

        // Draw polygonal grid lines (3 levels)
        self.ctx.set_stroke_style_str("#D9D9D9");
        self.ctx.set_line_width(1.0);

        for level in 1..=3 {
            let level_radius = radius * level as f64 / 3.0;
            self.ctx.begin_path();

            for i in 0..num_axes {
                let angle = -std::f64::consts::FRAC_PI_2 + i as f64 * angle_step;
                let px = center_x + level_radius * angle.cos();
                let py = center_y + level_radius * angle.sin();

                if i == 0 {
                    self.ctx.move_to(px, py);
                } else {
                    self.ctx.line_to(px, py);
                }
            }
            self.ctx.close_path();
            self.ctx.stroke();
        }

        // Draw axis lines from center to each vertex
        self.ctx.set_stroke_style_str("#BFBFBF");
        for i in 0..num_axes {
            let angle = -std::f64::consts::FRAC_PI_2 + i as f64 * angle_step;
            let px = center_x + radius * angle.cos();
            let py = center_y + radius * angle.sin();

            self.ctx.begin_path();
            self.ctx.move_to(center_x, center_y);
            self.ctx.line_to(px, py);
            self.ctx.stroke();
        }

        // Draw category labels at each axis
        if let Some(first_series) = chart.series.first() {
            if let Some(ref categories) = first_series.categories {
                self.ctx.set_font("10px Calibri, Arial, sans-serif");
                self.ctx.set_fill_style_str("#404040");

                for (i, label) in categories.str_values.iter().enumerate().take(num_axes) {
                    let angle = -std::f64::consts::FRAC_PI_2 + i as f64 * angle_step;
                    let label_radius = radius + 15.0;
                    let lx = center_x + label_radius * angle.cos();
                    let ly = center_y + label_radius * angle.sin();

                    // Adjust text alignment based on position
                    if angle.cos().abs() < 0.1 {
                        self.ctx.set_text_align("center");
                    } else if angle.cos() > 0.0 {
                        self.ctx.set_text_align("left");
                    } else {
                        self.ctx.set_text_align("right");
                    }

                    let _ = self.ctx.fill_text(label, lx, ly + 4.0);
                }
                self.ctx.set_text_align("left");
            }
        }

        // Draw each series as a filled polygon
        for (series_idx, series) in chart.series.iter().enumerate() {
            let color = Self::CHART_COLORS[series_idx % Self::CHART_COLORS.len()];

            if let Some(ref values) = series.values {
                // Collect points
                let mut points: Vec<(f64, f64)> = Vec::new();

                for (i, val) in values.num_values.iter().enumerate().take(num_axes) {
                    let v = val.unwrap_or(0.0);
                    let angle = -std::f64::consts::FRAC_PI_2 + i as f64 * angle_step;
                    let point_radius = (v / max_val) * radius;
                    let px = center_x + point_radius * angle.cos();
                    let py = center_y + point_radius * angle.sin();
                    points.push((px, py));
                }

                if points.is_empty() {
                    continue;
                }

                // Draw filled polygon with transparency
                self.ctx.set_global_alpha(0.3);
                self.ctx.set_fill_style_str(color);
                self.ctx.begin_path();

                if let Some((first_x, first_y)) = points.first() {
                    self.ctx.move_to(*first_x, *first_y);
                    for (px, py) in points.iter().skip(1) {
                        self.ctx.line_to(*px, *py);
                    }
                }
                self.ctx.close_path();
                self.ctx.fill();

                // Draw polygon outline
                self.ctx.set_global_alpha(1.0);
                self.ctx.set_stroke_style_str(color);
                self.ctx.set_line_width(2.0);
                self.ctx.begin_path();

                if let Some((first_x, first_y)) = points.first() {
                    self.ctx.move_to(*first_x, *first_y);
                    for (px, py) in points.iter().skip(1) {
                        self.ctx.line_to(*px, *py);
                    }
                }
                self.ctx.close_path();
                self.ctx.stroke();

                // Draw data points
                self.ctx.set_fill_style_str(color);
                for (px, py) in &points {
                    self.ctx.begin_path();
                    let _ = self.ctx.arc(*px, *py, 4.0, 0.0, std::f64::consts::PI * 2.0);
                    self.ctx.fill();
                }
            }
        }
    }

    /// Render a stock chart (High-Low-Close or OHLC)
    /// Uses vertical lines from low to high with tick marks for open/close
    #[allow(clippy::indexing_slicing)]
    fn render_stock_chart(&self, chart: &Chart, x: f64, y: f64, w: f64, h: f64) {
        // Stock charts typically have 3 or 4 series: Open (optional), High, Low, Close
        // We'll interpret series based on count:
        // 3 series: High, Low, Close
        // 4 series: Open, High, Low, Close

        let series_count = chart.series.len();
        if series_count < 3 {
            // Not enough data for a stock chart, render as placeholder
            self.ctx.set_font("12px Calibri, Arial, sans-serif");
            self.ctx.set_fill_style_str("#808080");
            self.ctx.set_text_align("center");
            let _ = self.ctx.fill_text(
                "[Stock chart requires 3-4 series]",
                x + w / 2.0,
                y + h / 2.0,
            );
            self.ctx.set_text_align("left");
            return;
        }

        // Get the series data
        let (open_data, high_data, low_data, close_data) = if series_count >= 4 {
            (
                chart.series.first().and_then(|s| s.values.as_ref()),
                chart.series.get(1).and_then(|s| s.values.as_ref()),
                chart.series.get(2).and_then(|s| s.values.as_ref()),
                chart.series.get(3).and_then(|s| s.values.as_ref()),
            )
        } else {
            (
                None,
                chart.series.first().and_then(|s| s.values.as_ref()),
                chart.series.get(1).and_then(|s| s.values.as_ref()),
                chart.series.get(2).and_then(|s| s.values.as_ref()),
            )
        };

        let Some(high_data) = high_data else {
            return;
        };
        let Some(low_data) = low_data else {
            return;
        };
        let Some(close_data) = close_data else {
            return;
        };

        // Find min/max values for scaling
        let mut min_val = f64::INFINITY;
        let mut max_val = f64::NEG_INFINITY;

        for v in high_data.num_values.iter().flatten() {
            max_val = max_val.max(*v);
        }
        for v in low_data.num_values.iter().flatten() {
            min_val = min_val.min(*v);
        }

        if min_val == f64::INFINITY || max_val == f64::NEG_INFINITY {
            return;
        }

        let range = (max_val - min_val).max(1.0);
        let num_points = high_data.num_values.len();
        let bar_spacing = w / (num_points as f64 + 1.0);
        let tick_width = (bar_spacing * 0.4).min(8.0);

        // Draw each OHLC bar
        for i in 0..num_points {
            let bar_x = x + (i as f64 + 1.0) * bar_spacing;

            let high = high_data.num_values.get(i).and_then(|v| *v);
            let low = low_data.num_values.get(i).and_then(|v| *v);
            let close = close_data.num_values.get(i).and_then(|v| *v);
            let open = open_data.and_then(|d| d.num_values.get(i).and_then(|v| *v));

            if let (Some(hi), Some(lo), Some(cl)) = (high, low, close) {
                // Determine color based on close vs open (or previous close)
                let is_up = if let Some(op) = open {
                    cl >= op
                } else if i > 0 {
                    let prev_close = close_data.num_values.get(i - 1).and_then(|v| *v);
                    prev_close.is_none_or(|pc| cl >= pc)
                } else {
                    true
                };

                let color = if is_up { "#00B050" } else { "#FF0000" };
                self.ctx.set_stroke_style_str(color);
                self.ctx.set_line_width(1.5);

                // Calculate y positions
                let high_y = y + h - ((hi - min_val) / range * h);
                let low_y = y + h - ((lo - min_val) / range * h);
                let close_y = y + h - ((cl - min_val) / range * h);

                // Draw vertical line (high to low)
                self.ctx.begin_path();
                self.ctx.move_to(bar_x, high_y);
                self.ctx.line_to(bar_x, low_y);
                self.ctx.stroke();

                // Draw close tick (right side)
                self.ctx.begin_path();
                self.ctx.move_to(bar_x, close_y);
                self.ctx.line_to(bar_x + tick_width, close_y);
                self.ctx.stroke();

                // Draw open tick (left side) if we have open data
                if let Some(op) = open {
                    let open_y = y + h - ((op - min_val) / range * h);
                    self.ctx.begin_path();
                    self.ctx.move_to(bar_x - tick_width, open_y);
                    self.ctx.line_to(bar_x, open_y);
                    self.ctx.stroke();
                }
            }
        }

        // Draw axis lines
        self.ctx.set_stroke_style_str("#808080");
        self.ctx.set_line_width(1.0);
        self.ctx.begin_path();
        self.ctx.move_to(Self::crisp(x), y);
        self.ctx.line_to(Self::crisp(x), y + h);
        self.ctx.move_to(x, Self::crisp(y + h));
        self.ctx.line_to(x + w, Self::crisp(y + h));
        self.ctx.stroke();
    }

    /// Render a surface chart as a 2D heatmap/contour view
    /// Uses color gradients to represent 3D data in 2D
    #[allow(clippy::indexing_slicing)]
    fn render_surface_chart(&self, chart: &Chart, x: f64, y: f64, w: f64, h: f64) {
        // Collect all data points to create a grid
        // Each series represents a row, each value in the series represents a column
        let mut grid: Vec<Vec<f64>> = Vec::new();
        let mut min_val = f64::INFINITY;
        let mut max_val = f64::NEG_INFINITY;

        for series in &chart.series {
            if let Some(ref values) = series.values {
                let row: Vec<f64> = values.num_values.iter().map(|v| v.unwrap_or(0.0)).collect();
                for &v in &row {
                    min_val = min_val.min(v);
                    max_val = max_val.max(v);
                }
                grid.push(row);
            }
        }

        if grid.is_empty() {
            return;
        }

        let num_rows = grid.len();
        let num_cols = grid.iter().map(|r| r.len()).max().unwrap_or(1);
        let range = (max_val - min_val).max(1.0);

        let cell_width = w / num_cols as f64;
        let cell_height = h / num_rows as f64;

        // Draw heatmap cells
        for (row_idx, row) in grid.iter().enumerate() {
            for (col_idx, &value) in row.iter().enumerate() {
                let normalized = (value - min_val) / range;

                // Create a color gradient from blue (cold/low) to red (hot/high)
                let color = Self::value_to_heatmap_color(normalized);

                let cell_x = x + col_idx as f64 * cell_width;
                let cell_y = y + row_idx as f64 * cell_height;

                self.ctx.set_fill_style_str(&color);
                self.ctx.fill_rect(cell_x, cell_y, cell_width, cell_height);
            }
        }

        // Draw grid lines
        self.ctx.set_stroke_style_str("#FFFFFF");
        self.ctx.set_line_width(0.5);

        // Vertical lines
        for col in 0..=num_cols {
            let line_x = x + col as f64 * cell_width;
            self.ctx.begin_path();
            self.ctx.move_to(Self::crisp(line_x), y);
            self.ctx.line_to(Self::crisp(line_x), y + h);
            self.ctx.stroke();
        }

        // Horizontal lines
        for row in 0..=num_rows {
            let line_y = y + row as f64 * cell_height;
            self.ctx.begin_path();
            self.ctx.move_to(x, Self::crisp(line_y));
            self.ctx.line_to(x + w, Self::crisp(line_y));
            self.ctx.stroke();
        }

        // Draw a color scale legend on the right side
        let legend_width = 15.0;
        let legend_x = x + w + 5.0;
        let legend_height = h * 0.8;
        let legend_y = y + h * 0.1;

        for i in 0..20 {
            let normalized = 1.0 - (i as f64 / 19.0);
            let color = Self::value_to_heatmap_color(normalized);
            let segment_h = legend_height / 20.0;
            let segment_y = legend_y + i as f64 * segment_h;

            self.ctx.set_fill_style_str(&color);
            self.ctx
                .fill_rect(legend_x, segment_y, legend_width, segment_h + 1.0);
        }

        // Legend labels
        self.ctx.set_font("9px Calibri, Arial, sans-serif");
        self.ctx.set_fill_style_str("#404040");
        self.ctx.set_text_align("left");
        let _ = self.ctx.fill_text(
            &format!("{:.1}", max_val),
            legend_x + legend_width + 3.0,
            legend_y + 4.0,
        );
        let _ = self.ctx.fill_text(
            &format!("{:.1}", min_val),
            legend_x + legend_width + 3.0,
            legend_y + legend_height + 4.0,
        );
    }

    /// Convert a normalized value (0-1) to a heatmap color
    #[allow(clippy::cast_possible_truncation, clippy::cast_sign_loss)]
    fn value_to_heatmap_color(normalized: f64) -> String {
        // Blue -> Cyan -> Green -> Yellow -> Red gradient
        let n = normalized.clamp(0.0, 1.0);

        let (r, g, b) = if n < 0.25 {
            // Blue to Cyan
            let t = n / 0.25;
            (0.0, t, 1.0)
        } else if n < 0.5 {
            // Cyan to Green
            let t = (n - 0.25) / 0.25;
            (0.0, 1.0, 1.0 - t)
        } else if n < 0.75 {
            // Green to Yellow
            let t = (n - 0.5) / 0.25;
            (t, 1.0, 0.0)
        } else {
            // Yellow to Red
            let t = (n - 0.75) / 0.25;
            (1.0, 1.0 - t, 0.0)
        };

        format!(
            "#{:02X}{:02X}{:02X}",
            (r * 255.0).clamp(0.0, 255.0) as u8,
            (g * 255.0).clamp(0.0, 255.0) as u8,
            (b * 255.0).clamp(0.0, 255.0) as u8
        )
    }

    /// Render a combo chart (multiple chart types overlaid)
    /// Each series can have its own chart type specified in series_type
    #[allow(clippy::indexing_slicing)]
    fn render_combo_chart(&self, chart: &Chart, x: f64, y: f64, w: f64, h: f64) {
        // First, collect all values to determine scale
        let mut all_values: Vec<f64> = Vec::new();
        let mut num_points = 0;

        for series in &chart.series {
            if let Some(ref values) = series.values {
                num_points = num_points.max(values.num_values.len());
                for v in values.num_values.iter().flatten() {
                    all_values.push(*v);
                }
            }
        }

        if all_values.is_empty() || num_points == 0 {
            return;
        }

        let min_val = all_values.iter().copied().fold(0.0_f64, f64::min);
        let max_val = all_values.iter().copied().fold(0.0_f64, f64::max);
        let range = (max_val - min_val).max(1.0);

        // Count series by type for bar grouping
        let bar_series_count = chart
            .series
            .iter()
            .filter(|s| matches!(s.series_type.unwrap_or(ChartType::Bar), ChartType::Bar))
            .count();

        let group_width = w / num_points as f64;
        let bar_width = if bar_series_count > 0 {
            (group_width * 0.8) / bar_series_count as f64
        } else {
            group_width * 0.8
        };
        let group_gap = 0.1;

        let mut bar_series_idx = 0;

        // Render each series based on its type
        for (series_idx, series) in chart.series.iter().enumerate() {
            let color = Self::CHART_COLORS[series_idx % Self::CHART_COLORS.len()];
            let series_type = series.series_type.unwrap_or(ChartType::Bar);

            match series_type {
                ChartType::Bar => {
                    // Render as bar/column
                    if let Some(ref values) = series.values {
                        self.ctx.set_fill_style_str(color);

                        for (point_idx, val) in values.num_values.iter().enumerate() {
                            if let Some(v) = val {
                                let bar_h = ((v - min_val) / range * h).max(0.0);
                                let bar_x = x
                                    + point_idx as f64 * group_width
                                    + group_width * group_gap / 2.0
                                    + bar_series_idx as f64 * bar_width;
                                let bar_y = y + h - bar_h;

                                self.ctx.fill_rect(bar_x, bar_y, bar_width * 0.9, bar_h);
                            }
                        }
                    }
                    bar_series_idx += 1;
                }
                ChartType::Line | ChartType::Scatter => {
                    // Render as line
                    if let Some(ref values) = series.values {
                        let point_spacing = if num_points > 1 {
                            w / (num_points - 1) as f64
                        } else {
                            w
                        };

                        // Draw line
                        self.ctx.set_stroke_style_str(color);
                        self.ctx.set_line_width(2.0);
                        self.ctx.begin_path();

                        let mut first = true;
                        for (point_idx, val) in values.num_values.iter().enumerate() {
                            if let Some(v) = val {
                                let px = x + point_idx as f64 * point_spacing;
                                let py = y + h - ((v - min_val) / range * h);

                                if first {
                                    self.ctx.move_to(px, py);
                                    first = false;
                                } else {
                                    self.ctx.line_to(px, py);
                                }
                            }
                        }
                        self.ctx.stroke();

                        // Draw data points
                        self.ctx.set_fill_style_str(color);
                        for (point_idx, val) in values.num_values.iter().enumerate() {
                            if let Some(v) = val {
                                let px = x + point_idx as f64 * point_spacing;
                                let py = y + h - ((v - min_val) / range * h);

                                self.ctx.begin_path();
                                let _ = self.ctx.arc(px, py, 4.0, 0.0, std::f64::consts::PI * 2.0);
                                self.ctx.fill();
                            }
                        }
                    }
                }
                ChartType::Area => {
                    // Render as area
                    if let Some(ref values) = series.values {
                        let point_spacing = if num_points > 1 {
                            w / (num_points - 1) as f64
                        } else {
                            w
                        };

                        // Draw filled area
                        self.ctx.begin_path();
                        self.ctx.move_to(x, y + h);

                        for (point_idx, val) in values.num_values.iter().enumerate() {
                            let v = val.unwrap_or(min_val);
                            let px = x + point_idx as f64 * point_spacing;
                            let py = y + h - ((v - min_val) / range * h);
                            self.ctx.line_to(px, py);
                        }

                        self.ctx.line_to(x + w, y + h);
                        self.ctx.close_path();

                        self.ctx.set_global_alpha(0.4);
                        self.ctx.set_fill_style_str(color);
                        self.ctx.fill();
                        self.ctx.set_global_alpha(1.0);

                        // Draw line on top
                        self.ctx.set_stroke_style_str(color);
                        self.ctx.set_line_width(2.0);
                        self.ctx.begin_path();

                        let mut first = true;
                        for (point_idx, val) in values.num_values.iter().enumerate() {
                            if let Some(v) = val {
                                let px = x + point_idx as f64 * point_spacing;
                                let py = y + h - ((v - min_val) / range * h);

                                if first {
                                    self.ctx.move_to(px, py);
                                    first = false;
                                } else {
                                    self.ctx.line_to(px, py);
                                }
                            }
                        }
                        self.ctx.stroke();
                    }
                }
                _ => {
                    // For other types, default to line rendering
                    if let Some(ref values) = series.values {
                        let point_spacing = if num_points > 1 {
                            w / (num_points - 1) as f64
                        } else {
                            w
                        };

                        self.ctx.set_stroke_style_str(color);
                        self.ctx.set_line_width(2.0);
                        self.ctx.begin_path();

                        let mut first = true;
                        for (point_idx, val) in values.num_values.iter().enumerate() {
                            if let Some(v) = val {
                                let px = x + point_idx as f64 * point_spacing;
                                let py = y + h - ((v - min_val) / range * h);

                                if first {
                                    self.ctx.move_to(px, py);
                                    first = false;
                                } else {
                                    self.ctx.line_to(px, py);
                                }
                            }
                        }
                        self.ctx.stroke();
                    }
                }
            }
        }

        // Draw axis lines
        self.ctx.set_stroke_style_str("#808080");
        self.ctx.set_line_width(1.0);
        self.ctx.begin_path();
        self.ctx.move_to(Self::crisp(x), y);
        self.ctx.line_to(Self::crisp(x), y + h);
        self.ctx.move_to(x, Self::crisp(y + h));
        self.ctx.line_to(x + w, Self::crisp(y + h));
        self.ctx.stroke();
    }

    /// Render chart legend
    #[allow(clippy::indexing_slicing)] // Safe: modulo ensures index is within bounds
    fn render_chart_legend(
        &self,
        chart: &Chart,
        _legend: &crate::types::ChartLegend,
        x: f64,
        y: f64,
        w: f64,
        _h: f64,
    ) {
        self.ctx.set_font("10px Calibri, Arial, sans-serif");
        self.ctx.set_text_align("left");

        let mut legend_x = x + 10.0;
        let legend_y = y + 12.0;
        let box_size = 10.0;
        let spacing = 8.0;

        for (idx, series) in chart.series.iter().enumerate() {
            let color = Self::CHART_COLORS[idx % Self::CHART_COLORS.len()];
            let default_name = format!("Series {}", idx + 1);
            let name = series.name.as_deref().unwrap_or(&default_name);

            // Don't overflow the chart width
            let text_width = self
                .ctx
                .measure_text(name)
                .map(|m| m.width())
                .unwrap_or(50.0);
            if legend_x + box_size + spacing + text_width > x + w - 10.0 {
                break;
            }

            // Draw color box
            self.ctx.set_fill_style_str(color);
            self.ctx
                .fill_rect(legend_x, legend_y - box_size + 2.0, box_size, box_size);

            // Draw series name
            self.ctx.set_fill_style_str("#000000");
            let _ = self
                .ctx
                .fill_text(name, legend_x + box_size + 4.0, legend_y);

            legend_x += box_size + spacing + text_width + 15.0;
        }
    }
}
