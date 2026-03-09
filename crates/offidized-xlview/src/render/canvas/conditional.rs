//! Conditional formatting rendering for Canvas 2D backend.

use std::collections::HashMap;

use crate::cell_ref::parse_sqref;
use crate::layout::{SheetLayout, Viewport};
use crate::render::backend::CellRenderData;
use crate::render::colors::Rgb;
use crate::types::formatting::{CFRule, ConditionalFormatting};

use super::renderer::CanvasRenderer;

impl CanvasRenderer {
    /// Render conditional formatting visuals (color scales, data bars, icon sets, cellIs)
    /// This should be called BEFORE cell backgrounds so CF visuals appear under cell content
    pub(super) fn render_conditional_formatting(
        &self,
        cells: &[&CellRenderData],
        cf_rules: &[ConditionalFormatting],
        dxf_styles: &[crate::types::DxfStyle],
        layout: &SheetLayout,
        viewport: &Viewport,
        cf_cache: &[crate::types::ConditionalFormattingCache],
    ) {
        for (idx, cf) in cf_rules.iter().enumerate() {
            // Parse the sqref to get affected cell ranges (cached when available)
            let ranges_storage;
            let ranges = if let Some(cache) = cf_cache.get(idx) {
                &cache.ranges
            } else {
                ranges_storage = parse_sqref(&cf.sqref);
                &ranges_storage
            };

            // For each rule, sorted by priority (lower priority number = higher priority)
            let sorted_rules: Vec<&CFRule> = if let Some(cache) = cf_cache.get(idx) {
                cache
                    .sorted_rule_indices
                    .iter()
                    .filter_map(|&i| cf.rules.get(i))
                    .collect()
            } else {
                let mut rules: Vec<&CFRule> = cf.rules.iter().collect();
                rules.sort_by_key(|r| r.priority);
                rules
            };

            let range_values = Self::collect_range_values(cells, ranges);
            if range_values.is_empty() {
                continue;
            }

            let (min_val, max_val) = Self::get_min_max(&range_values);
            let range_span = max_val - min_val;

            for rule in sorted_rules {
                // Render based on rule type
                if let Some(ref color_scale) = rule.color_scale {
                    self.render_color_scale(
                        cells,
                        ranges,
                        color_scale,
                        min_val,
                        range_span,
                        layout,
                        viewport,
                    );
                } else if let Some(ref data_bar) = rule.data_bar {
                    self.render_data_bar(
                        cells, ranges, data_bar, min_val, range_span, layout, viewport,
                    );
                } else if let Some(ref icon_set) = rule.icon_set {
                    self.render_icon_set(
                        cells, ranges, icon_set, min_val, range_span, layout, viewport,
                    );
                } else if rule.rule_type == "cellIs" {
                    // Handle cellIs rules with DXF formatting
                    self.render_cell_is(cells, ranges, rule, dxf_styles, layout, viewport);
                }
            }
        }
    }

    /// Render cellIs conditional formatting rule
    /// These rules compare cell values to a formula/value and apply DXF formatting
    fn render_cell_is(
        &self,
        cells: &[&CellRenderData],
        ranges: &[(u32, u32, u32, u32)],
        rule: &CFRule,
        dxf_styles: &[crate::types::DxfStyle],
        layout: &SheetLayout,
        viewport: &Viewport,
    ) {
        // Get the DXF style for this rule
        let dxf = match rule.dxf_id {
            Some(id) => dxf_styles.get(id as usize),
            None => return,
        };

        let Some(dxf) = dxf else { return };

        // Get the comparison value from the formula
        let compare_value: Option<f64> = rule.formula.as_ref().and_then(|f| f.parse().ok());
        let compare_str = rule.formula.as_deref();

        // Get the operator
        let operator = rule.operator.as_deref().unwrap_or("equal");

        // Check each cell in the ranges
        for &cell in cells {
            // Check if cell is in any of the ranges
            let in_range = ranges.iter().any(|(sr, sc, er, ec)| {
                cell.row >= *sr && cell.row <= *er && cell.col >= *sc && cell.col <= *ec
            });

            if !in_range {
                continue;
            }

            // Get cell value
            let cell_value: Option<f64> = cell.numeric_value;
            let cell_str = cell.value.as_deref();

            // Evaluate the condition
            let matches = match operator {
                "equal" => {
                    if let (Some(cv), Some(cmp)) = (cell_value, compare_value) {
                        (cv - cmp).abs() < f64::EPSILON
                    } else {
                        cell_str == compare_str
                    }
                }
                "notEqual" => {
                    if let (Some(cv), Some(cmp)) = (cell_value, compare_value) {
                        (cv - cmp).abs() >= f64::EPSILON
                    } else {
                        cell_str != compare_str
                    }
                }
                "greaterThan" => {
                    if let (Some(cv), Some(cmp)) = (cell_value, compare_value) {
                        cv > cmp
                    } else {
                        false
                    }
                }
                "greaterThanOrEqual" => {
                    if let (Some(cv), Some(cmp)) = (cell_value, compare_value) {
                        cv >= cmp
                    } else {
                        false
                    }
                }
                "lessThan" => {
                    if let (Some(cv), Some(cmp)) = (cell_value, compare_value) {
                        cv < cmp
                    } else {
                        false
                    }
                }
                "lessThanOrEqual" => {
                    if let (Some(cv), Some(cmp)) = (cell_value, compare_value) {
                        cv <= cmp
                    } else {
                        false
                    }
                }
                "between" => {
                    // Between requires two values - would need second formula
                    false
                }
                "notBetween" => false,
                "containsText" => {
                    if let (Some(cv), Some(cmp)) = (cell_str, compare_str) {
                        cv.contains(cmp)
                    } else {
                        false
                    }
                }
                "notContainsText" | "notContains" => {
                    if let (Some(cv), Some(cmp)) = (cell_str, compare_str) {
                        !cv.contains(cmp)
                    } else {
                        true
                    }
                }
                "beginsWith" => {
                    if let (Some(cv), Some(cmp)) = (cell_str, compare_str) {
                        cv.starts_with(cmp)
                    } else {
                        false
                    }
                }
                "endsWith" => {
                    if let (Some(cv), Some(cmp)) = (cell_str, compare_str) {
                        cv.ends_with(cmp)
                    } else {
                        false
                    }
                }
                _ => false,
            };

            if !matches {
                continue;
            }

            // Get cell bounds
            let rect = layout.cell_rect(cell.row, cell.col);
            if rect.skip || rect.width <= 0.0 || rect.height <= 0.0 {
                continue;
            }

            let (sx, sy) = viewport.to_screen_frozen(rect.x, rect.y, cell.row, cell.col, layout);
            let x = f64::from(sx);
            let y = f64::from(sy);
            let w = f64::from(rect.width);
            let h = f64::from(rect.height);

            // Draw fill color if specified
            if let Some(ref fill_color) = dxf.fill_color {
                self.ctx.set_fill_style_str(fill_color);
                self.ctx.fill_rect(x, y, w, h);
            }
        }
    }

    /// Collect DXF style overrides for cells matching cellIs conditional formatting rules.
    /// Returns a map of (row, col) -> DxfStyle for cells that match cellIs rules.
    /// This allows font styling from CF rules to be applied during text rendering.
    pub(super) fn collect_cell_is_dxf_overrides<'a>(
        cells: &[&CellRenderData],
        cf_rules: &[ConditionalFormatting],
        dxf_styles: &'a [crate::types::DxfStyle],
        cf_cache: &[crate::types::ConditionalFormattingCache],
    ) -> HashMap<(u32, u32), &'a crate::types::DxfStyle> {
        let mut overrides: HashMap<(u32, u32), &'a crate::types::DxfStyle> = HashMap::new();

        for (idx, cf) in cf_rules.iter().enumerate() {
            let ranges_storage;
            let ranges = if let Some(cache) = cf_cache.get(idx) {
                &cache.ranges
            } else {
                ranges_storage = parse_sqref(&cf.sqref);
                &ranges_storage
            };

            // Sort rules by priority (lower priority number = higher priority, applied last)
            let sorted_rules: Vec<&CFRule> = if let Some(cache) = cf_cache.get(idx) {
                let mut rules: Vec<&CFRule> = cache
                    .sorted_rule_indices
                    .iter()
                    .filter_map(|&i| cf.rules.get(i))
                    .collect();
                rules.sort_by_key(|r| std::cmp::Reverse(r.priority));
                rules
            } else {
                let mut rules: Vec<&CFRule> = cf.rules.iter().collect();
                rules.sort_by_key(|r| std::cmp::Reverse(r.priority));
                rules
            };

            for rule in sorted_rules {
                if rule.rule_type != "cellIs" {
                    continue;
                }

                let dxf = match rule.dxf_id {
                    Some(id) => dxf_styles.get(id as usize),
                    None => continue,
                };

                let Some(dxf) = dxf else { continue };

                // Skip if no font styling properties
                if dxf.font_color.is_none()
                    && dxf.bold.is_none()
                    && dxf.italic.is_none()
                    && dxf.underline.is_none()
                    && dxf.strikethrough.is_none()
                {
                    continue;
                }

                let compare_value: Option<f64> = rule.formula.as_ref().and_then(|f| f.parse().ok());
                let compare_str = rule.formula.as_deref();
                let operator = rule.operator.as_deref().unwrap_or("equal");

                for &cell in cells {
                    let in_range = ranges.iter().any(|(sr, sc, er, ec)| {
                        cell.row >= *sr && cell.row <= *er && cell.col >= *sc && cell.col <= *ec
                    });

                    if !in_range {
                        continue;
                    }

                    let cell_value: Option<f64> = cell.numeric_value;
                    let cell_str = cell.value.as_deref();

                    let matches = match operator {
                        "equal" => {
                            if let (Some(cv), Some(cmp)) = (cell_value, compare_value) {
                                (cv - cmp).abs() < f64::EPSILON
                            } else {
                                cell_str == compare_str
                            }
                        }
                        "notEqual" => {
                            if let (Some(cv), Some(cmp)) = (cell_value, compare_value) {
                                (cv - cmp).abs() >= f64::EPSILON
                            } else {
                                cell_str != compare_str
                            }
                        }
                        "greaterThan" => {
                            if let (Some(cv), Some(cmp)) = (cell_value, compare_value) {
                                cv > cmp
                            } else {
                                false
                            }
                        }
                        "greaterThanOrEqual" => {
                            if let (Some(cv), Some(cmp)) = (cell_value, compare_value) {
                                cv >= cmp
                            } else {
                                false
                            }
                        }
                        "lessThan" => {
                            if let (Some(cv), Some(cmp)) = (cell_value, compare_value) {
                                cv < cmp
                            } else {
                                false
                            }
                        }
                        "lessThanOrEqual" => {
                            if let (Some(cv), Some(cmp)) = (cell_value, compare_value) {
                                cv <= cmp
                            } else {
                                false
                            }
                        }
                        "between" => false,
                        "notBetween" => false,
                        "containsText" => {
                            if let (Some(cv), Some(cmp)) = (cell_str, compare_str) {
                                cv.contains(cmp)
                            } else {
                                false
                            }
                        }
                        "notContainsText" | "notContains" => {
                            if let (Some(cv), Some(cmp)) = (cell_str, compare_str) {
                                !cv.contains(cmp)
                            } else {
                                true
                            }
                        }
                        "beginsWith" => {
                            if let (Some(cv), Some(cmp)) = (cell_str, compare_str) {
                                cv.starts_with(cmp)
                            } else {
                                false
                            }
                        }
                        "endsWith" => {
                            if let (Some(cv), Some(cmp)) = (cell_str, compare_str) {
                                cv.ends_with(cmp)
                            } else {
                                false
                            }
                        }
                        _ => false,
                    };

                    if matches {
                        overrides.insert((cell.row, cell.col), dxf);
                    }
                }
            }
        }

        overrides
    }

    /// Collect numeric values from cells within the given ranges
    fn collect_range_values(
        cells: &[&CellRenderData],
        ranges: &[(u32, u32, u32, u32)],
    ) -> Vec<(u32, u32, f64)> {
        let mut values = Vec::new();

        for &cell in cells {
            // Check if cell is in any of the ranges
            let in_range = ranges.iter().any(|(sr, sc, er, ec)| {
                cell.row >= *sr && cell.row <= *er && cell.col >= *sc && cell.col <= *ec
            });

            if !in_range {
                continue;
            }

            if let Some(num) = cell.numeric_value {
                values.push((cell.row, cell.col, num));
            }
        }

        values
    }

    /// Get min and max values from a collection
    fn get_min_max(values: &[(u32, u32, f64)]) -> (f64, f64) {
        let min = values
            .iter()
            .map(|(_, _, v)| *v)
            .fold(f64::INFINITY, f64::min);
        let max = values
            .iter()
            .map(|(_, _, v)| *v)
            .fold(f64::NEG_INFINITY, f64::max);
        (min, max)
    }

    /// Calculate normalized position (0.0 to 1.0) of a value within a range
    fn normalize_value(value: f64, min: f64, range_span: f64) -> f64 {
        if range_span <= 0.0 {
            0.5 // All values are the same
        } else {
            ((value - min) / range_span).clamp(0.0, 1.0)
        }
    }

    /// Interpolate between colors based on position (0.0 to 1.0)
    #[allow(
        clippy::indexing_slicing,
        clippy::cast_possible_truncation,
        clippy::cast_sign_loss
    )]
    fn interpolate_color(colors: &[String], position: f64) -> String {
        if colors.is_empty() {
            return "#FFFFFF".to_string();
        }
        if colors.len() == 1 {
            return colors[0].clone();
        }

        let position = position.clamp(0.0, 1.0);
        let num_segments = colors.len() - 1;
        let segment_size = 1.0 / num_segments as f64;

        // Find which segment we're in
        let segment_index = (position / segment_size).floor() as usize;
        let segment_index = segment_index.min(num_segments - 1);

        // Position within the segment (0.0 to 1.0)
        let segment_pos = (position - (segment_index as f64 * segment_size)) / segment_size;
        let segment_pos = segment_pos.clamp(0.0, 1.0);

        // Get the two colors to interpolate between
        let color1 = Rgb::from_hex(&colors[segment_index]).unwrap_or(Rgb::new(255, 255, 255));
        let color2 = Rgb::from_hex(&colors[segment_index + 1]).unwrap_or(Rgb::new(255, 255, 255));

        // Linear interpolation
        let r = Self::lerp_u8(color1.r, color2.r, segment_pos);
        let g = Self::lerp_u8(color1.g, color2.g, segment_pos);
        let b = Self::lerp_u8(color1.b, color2.b, segment_pos);

        format!("#{:02X}{:02X}{:02X}", r, g, b)
    }

    /// Linear interpolation for u8 values
    #[allow(clippy::cast_possible_truncation, clippy::cast_sign_loss)]
    fn lerp_u8(a: u8, b: u8, t: f64) -> u8 {
        let a = f64::from(a);
        let b = f64::from(b);
        (a + (b - a) * t).round().clamp(0.0, 255.0) as u8
    }

    /// Render color scale backgrounds
    #[allow(clippy::too_many_arguments)]
    fn render_color_scale(
        &self,
        cells: &[&CellRenderData],
        ranges: &[(u32, u32, u32, u32)],
        color_scale: &crate::types::ColorScale,
        min_val: f64,
        range_span: f64,
        layout: &SheetLayout,
        viewport: &Viewport,
    ) {
        for &cell in cells {
            // Check if cell is in any of the ranges
            let in_range = ranges.iter().any(|(sr, sc, er, ec)| {
                cell.row >= *sr && cell.row <= *er && cell.col >= *sc && cell.col <= *ec
            });

            if !in_range {
                continue;
            }

            let Some(value) = cell.numeric_value else {
                continue;
            };

            let rect = layout.cell_rect(cell.row, cell.col);
            if rect.skip || rect.width <= 0.0 || rect.height <= 0.0 {
                continue;
            }

            let (sx, sy) = viewport.to_screen_frozen(rect.x, rect.y, cell.row, cell.col, layout);
            let x = f64::from(sx);
            let y = f64::from(sy);
            let w = f64::from(rect.width);
            let h = f64::from(rect.height);

            // Calculate position in the range and get interpolated color
            let position = Self::normalize_value(value, min_val, range_span);
            let color = Self::interpolate_color(&color_scale.colors, position);

            // Draw the background
            self.fill_rect(x, y, w, h, &color);
        }
    }

    /// Render data bars
    #[allow(clippy::too_many_arguments)]
    fn render_data_bar(
        &self,
        cells: &[&CellRenderData],
        ranges: &[(u32, u32, u32, u32)],
        data_bar: &crate::types::DataBar,
        min_val: f64,
        range_span: f64,
        layout: &SheetLayout,
        viewport: &Viewport,
    ) {
        // Default bar color if not specified
        let bar_color = if data_bar.color.is_empty() {
            "#638EC6"
        } else {
            &data_bar.color
        };

        for &cell in cells {
            // Check if cell is in any of the ranges
            let in_range = ranges.iter().any(|(sr, sc, er, ec)| {
                cell.row >= *sr && cell.row <= *er && cell.col >= *sc && cell.col <= *ec
            });

            if !in_range {
                continue;
            }

            let Some(value) = cell.numeric_value else {
                continue;
            };

            let rect = layout.cell_rect(cell.row, cell.col);
            if rect.skip || rect.width <= 0.0 || rect.height <= 0.0 {
                continue;
            }

            let (sx, sy) = viewport.to_screen_frozen(rect.x, rect.y, cell.row, cell.col, layout);
            let x = f64::from(sx);
            let y = f64::from(sy);
            let w = f64::from(rect.width);
            let h = f64::from(rect.height);

            // Calculate bar width as percentage of cell
            let position = Self::normalize_value(value, min_val, range_span);

            // Apply min/max length constraints
            let min_length = data_bar.min_length.unwrap_or(10) as f64 / 100.0;
            let max_length = data_bar.max_length.unwrap_or(90) as f64 / 100.0;
            let bar_pct = min_length + position * (max_length - min_length);

            let bar_width = (w - 4.0) * bar_pct; // Leave 2px padding on each side
            let bar_height = h * 0.6; // 60% of cell height
            let bar_y = y + (h - bar_height) / 2.0; // Center vertically
            let bar_x = x + 2.0; // 2px left padding

            // Draw the bar
            self.fill_rect(bar_x, bar_y, bar_width, bar_height, bar_color);
        }
    }

    /// Render icon sets with proper icons based on icon set type
    #[allow(clippy::too_many_arguments)]
    fn render_icon_set(
        &self,
        cells: &[&CellRenderData],
        ranges: &[(u32, u32, u32, u32)],
        icon_set: &crate::types::IconSet,
        min_val: f64,
        range_span: f64,
        layout: &SheetLayout,
        viewport: &Viewport,
    ) {
        let reverse = icon_set.reverse.unwrap_or(false);
        let icon_set_name = &icon_set.icon_set;

        // Determine number of icons based on set name
        let num_icons = if icon_set_name.starts_with("5") {
            5
        } else if icon_set_name.starts_with("4") {
            4
        } else {
            3
        };

        for &cell in cells {
            // Check if cell is in any of the ranges
            let in_range = ranges.iter().any(|(sr, sc, er, ec)| {
                cell.row >= *sr && cell.row <= *er && cell.col >= *sc && cell.col <= *ec
            });

            if !in_range {
                continue;
            }

            let Some(value) = cell.numeric_value else {
                continue;
            };

            let rect = layout.cell_rect(cell.row, cell.col);
            if rect.skip || rect.width <= 0.0 || rect.height <= 0.0 {
                continue;
            }

            let (sx, sy) = viewport.to_screen_frozen(rect.x, rect.y, cell.row, cell.col, layout);
            let x = f64::from(sx);
            let y = f64::from(sy);
            let h = f64::from(rect.height);

            // Calculate position in the range (0.0 to 1.0)
            let position = Self::normalize_value(value, min_val, range_span);

            // Determine icon index based on position and number of icons
            let mut icon_index = if num_icons == 5 {
                if position >= 0.8 {
                    4
                } else if position >= 0.6 {
                    3
                } else if position >= 0.4 {
                    2
                } else if position >= 0.2 {
                    1
                } else {
                    0
                }
            } else if num_icons == 4 {
                if position >= 0.75 {
                    3
                } else if position >= 0.5 {
                    2
                } else if position >= 0.25 {
                    1
                } else {
                    0
                }
            } else {
                // 3 icons
                if position >= 0.67 {
                    2
                } else if position >= 0.33 {
                    1
                } else {
                    0
                }
            };

            // Reverse the icon index if needed
            if reverse {
                icon_index = (num_icons - 1) - icon_index;
            }

            // Icon position (left side of cell)
            let icon_size = 12.0;
            let icon_x = x + 4.0;
            let icon_y = y + (h - icon_size) / 2.0;

            // Draw the appropriate icon
            self.draw_icon(
                icon_set_name,
                icon_index,
                num_icons,
                icon_x,
                icon_y,
                icon_size,
            );
        }
    }

    /// Draw a specific icon from an icon set
    fn draw_icon(
        &self,
        icon_set_name: &str,
        icon_index: usize,
        num_icons: usize,
        x: f64,
        y: f64,
        size: f64,
    ) {
        // Colors for icons (green = high/good, yellow = medium, red = low/bad)
        let green = "#63BE7B";
        let yellow = "#FFEB84";
        let red = "#F8696B";
        let gray = "#808080";

        // Determine color based on icon index and number of icons
        let color = if num_icons == 5 {
            match icon_index {
                4 => green,
                3 => "#9ACD32", // yellow-green
                2 => yellow,
                1 => "#FFA500", // orange
                _ => red,
            }
        } else if num_icons == 4 {
            match icon_index {
                3 => green,
                2 => yellow,
                1 => "#FFA500", // orange
                _ => red,
            }
        } else {
            match icon_index {
                2 => green,
                1 => yellow,
                _ => red,
            }
        };

        // Draw based on icon set type
        if icon_set_name.contains("Arrow") {
            self.draw_arrow_icon(icon_index, num_icons, x, y, size, color);
        } else if icon_set_name.contains("TrafficLight") || icon_set_name.contains("Lights") {
            self.draw_traffic_light_icon(x, y, size, color);
        } else if icon_set_name.contains("Symbol") || icon_set_name.contains("Signs") {
            self.draw_symbol_icon(icon_index, num_icons, x, y, size, color);
        } else if icon_set_name.contains("Rating") || icon_set_name.contains("Bars") {
            self.draw_rating_icon(icon_index, num_icons, x, y, size, gray);
        } else if icon_set_name.contains("Quarter") || icon_set_name.contains("Pie") {
            self.draw_quarter_icon(icon_index, num_icons, x, y, size);
        } else if icon_set_name.contains("Flag") {
            self.draw_flag_icon(x, y, size, color);
        } else if icon_set_name.contains("Star") {
            self.draw_star_icon(icon_index, num_icons, x, y, size);
        } else if icon_set_name.contains("Triangle") {
            self.draw_triangle_icon(icon_index, num_icons, x, y, size, color);
        } else {
            // Default: draw traffic light (circle)
            self.draw_traffic_light_icon(x, y, size, color);
        }
    }

    /// Draw an arrow icon (up, right, or down)
    fn draw_arrow_icon(
        &self,
        icon_index: usize,
        num_icons: usize,
        x: f64,
        y: f64,
        size: f64,
        color: &str,
    ) {
        let cx = x + size / 2.0;
        let cy = y + size / 2.0;
        let half = size / 2.0 - 1.0;

        self.ctx.set_fill_style_str(color);
        self.ctx.begin_path();

        // Determine arrow direction
        let is_up = if num_icons == 5 {
            icon_index >= 3
        } else if num_icons == 4 {
            icon_index >= 2
        } else {
            icon_index == 2
        };

        let is_down = icon_index == 0;
        let _is_right = !is_up && !is_down;

        if is_up {
            // Up arrow
            self.ctx.move_to(cx, cy - half); // Top point
            self.ctx.line_to(cx + half, cy + half * 0.3); // Bottom right
            self.ctx.line_to(cx + half * 0.3, cy + half * 0.3);
            self.ctx.line_to(cx + half * 0.3, cy + half); // Stem bottom right
            self.ctx.line_to(cx - half * 0.3, cy + half); // Stem bottom left
            self.ctx.line_to(cx - half * 0.3, cy + half * 0.3);
            self.ctx.line_to(cx - half, cy + half * 0.3); // Bottom left
            self.ctx.close_path();
        } else if is_down {
            // Down arrow
            self.ctx.move_to(cx, cy + half); // Bottom point
            self.ctx.line_to(cx + half, cy - half * 0.3); // Top right
            self.ctx.line_to(cx + half * 0.3, cy - half * 0.3);
            self.ctx.line_to(cx + half * 0.3, cy - half); // Stem top right
            self.ctx.line_to(cx - half * 0.3, cy - half); // Stem top left
            self.ctx.line_to(cx - half * 0.3, cy - half * 0.3);
            self.ctx.line_to(cx - half, cy - half * 0.3); // Top left
            self.ctx.close_path();
        } else {
            // Right arrow (horizontal, for middle values)
            self.ctx.move_to(cx + half, cy); // Right point
            self.ctx.line_to(cx - half * 0.3, cy - half); // Top left
            self.ctx.line_to(cx - half * 0.3, cy - half * 0.3);
            self.ctx.line_to(cx - half, cy - half * 0.3); // Stem top
            self.ctx.line_to(cx - half, cy + half * 0.3); // Stem bottom
            self.ctx.line_to(cx - half * 0.3, cy + half * 0.3);
            self.ctx.line_to(cx - half * 0.3, cy + half); // Bottom left
            self.ctx.close_path();
        }

        self.ctx.fill();
    }

    /// Draw a traffic light icon (filled circle)
    fn draw_traffic_light_icon(&self, x: f64, y: f64, size: f64, color: &str) {
        let cx = x + size / 2.0;
        let cy = y + size / 2.0;
        let radius = size / 2.0 - 1.0;

        self.ctx.set_fill_style_str(color);
        self.ctx.begin_path();
        let _ = self
            .ctx
            .arc(cx, cy, radius, 0.0, std::f64::consts::PI * 2.0);
        self.ctx.fill();
    }

    /// Draw a symbol icon (check, exclamation, or X)
    fn draw_symbol_icon(
        &self,
        icon_index: usize,
        num_icons: usize,
        x: f64,
        y: f64,
        size: f64,
        color: &str,
    ) {
        let cx = x + size / 2.0;
        let cy = y + size / 2.0;
        let half = size / 2.0 - 2.0;

        // Draw circle background
        self.ctx.set_fill_style_str(color);
        self.ctx.begin_path();
        let _ = self
            .ctx
            .arc(cx, cy, size / 2.0 - 1.0, 0.0, std::f64::consts::PI * 2.0);
        self.ctx.fill();

        // Draw white symbol on top
        self.ctx.set_stroke_style_str("#FFFFFF");
        self.ctx.set_line_width(2.0);
        self.ctx.begin_path();

        let is_check = if num_icons == 3 {
            icon_index == 2
        } else {
            icon_index >= num_icons - 1
        };
        let is_x = icon_index == 0;

        if is_check {
            // Checkmark
            self.ctx.move_to(cx - half * 0.6, cy);
            self.ctx.line_to(cx - half * 0.1, cy + half * 0.5);
            self.ctx.line_to(cx + half * 0.6, cy - half * 0.5);
        } else if is_x {
            // X mark
            self.ctx.move_to(cx - half * 0.5, cy - half * 0.5);
            self.ctx.line_to(cx + half * 0.5, cy + half * 0.5);
            self.ctx.move_to(cx + half * 0.5, cy - half * 0.5);
            self.ctx.line_to(cx - half * 0.5, cy + half * 0.5);
        } else {
            // Exclamation mark
            self.ctx.move_to(cx, cy - half * 0.5);
            self.ctx.line_to(cx, cy + half * 0.1);
            self.ctx.move_to(cx, cy + half * 0.5);
            self.ctx.line_to(cx, cy + half * 0.5);
        }

        self.ctx.stroke();
    }

    /// Draw a rating icon (bars showing 1-4 or 1-5 levels)
    fn draw_rating_icon(
        &self,
        icon_index: usize,
        num_icons: usize,
        x: f64,
        y: f64,
        size: f64,
        color: &str,
    ) {
        let bar_width = size / (num_icons as f64 + 1.0);
        let gap = 1.0;
        let max_height = size - 2.0;

        for i in 0..num_icons {
            let bar_height = max_height * ((i + 1) as f64) / (num_icons as f64);
            let bar_x = x + 1.0 + (i as f64) * (bar_width + gap);
            let bar_y = y + size - 1.0 - bar_height;

            // Fill bar if it's at or below the current level
            if i <= icon_index {
                self.ctx.set_fill_style_str(color);
            } else {
                self.ctx.set_fill_style_str("#E0E0E0");
            }

            self.ctx.fill_rect(bar_x, bar_y, bar_width, bar_height);
        }
    }

    /// Draw a quarter/pie icon (0%, 25%, 50%, 75%, 100% filled)
    fn draw_quarter_icon(&self, icon_index: usize, num_icons: usize, x: f64, y: f64, size: f64) {
        let cx = x + size / 2.0;
        let cy = y + size / 2.0;
        let radius = size / 2.0 - 1.0;

        // Calculate fill percentage based on icon index
        let fill_percent = if num_icons == 5 {
            (icon_index as f64) / 4.0 // 0, 0.25, 0.5, 0.75, 1.0
        } else {
            (icon_index as f64) / ((num_icons - 1) as f64)
        };

        // Draw outer circle (gray outline)
        self.ctx.set_stroke_style_str("#808080");
        self.ctx.set_line_width(1.0);
        self.ctx.begin_path();
        let _ = self
            .ctx
            .arc(cx, cy, radius, 0.0, std::f64::consts::PI * 2.0);
        self.ctx.stroke();

        // Draw white background
        self.ctx.set_fill_style_str("#FFFFFF");
        self.ctx.begin_path();
        let _ = self
            .ctx
            .arc(cx, cy, radius - 0.5, 0.0, std::f64::consts::PI * 2.0);
        self.ctx.fill();

        // Draw filled portion (pie slice from top, clockwise)
        if fill_percent > 0.0 {
            self.ctx.set_fill_style_str("#808080");
            self.ctx.begin_path();
            self.ctx.move_to(cx, cy);

            // Start from top (-PI/2) and go clockwise
            let start_angle = -std::f64::consts::PI / 2.0;
            let end_angle = start_angle + (fill_percent * std::f64::consts::PI * 2.0);

            let _ = self.ctx.arc(cx, cy, radius - 0.5, start_angle, end_angle);
            self.ctx.close_path();
            self.ctx.fill();
        }
    }

    /// Draw a flag icon
    fn draw_flag_icon(&self, x: f64, y: f64, size: f64, color: &str) {
        let pole_x = x + 2.0;
        let flag_width = size - 4.0;
        let flag_height = size * 0.6;

        // Draw pole
        self.ctx.set_stroke_style_str("#404040");
        self.ctx.set_line_width(1.5);
        self.ctx.begin_path();
        self.ctx.move_to(pole_x, y + 1.0);
        self.ctx.line_to(pole_x, y + size - 1.0);
        self.ctx.stroke();

        // Draw flag
        self.ctx.set_fill_style_str(color);
        self.ctx.begin_path();
        self.ctx.move_to(pole_x, y + 1.0);
        self.ctx
            .line_to(pole_x + flag_width, y + 1.0 + flag_height / 2.0);
        self.ctx.line_to(pole_x, y + 1.0 + flag_height);
        self.ctx.close_path();
        self.ctx.fill();
    }

    /// Draw a star icon (filled or empty based on index)
    fn draw_star_icon(&self, icon_index: usize, num_icons: usize, x: f64, y: f64, size: f64) {
        let cx = x + size / 2.0;
        let cy = y + size / 2.0;
        let outer_radius = size / 2.0 - 1.0;
        let inner_radius = outer_radius * 0.4;

        // Determine fill based on index
        let fill_color = if icon_index >= num_icons / 2 {
            "#FFD700" // Gold for higher values
        } else {
            "#E0E0E0" // Gray for lower values
        };

        self.ctx.set_fill_style_str(fill_color);
        self.ctx.set_stroke_style_str("#808080");
        self.ctx.set_line_width(0.5);
        self.ctx.begin_path();

        // Draw 5-pointed star
        for i in 0..10 {
            let radius = if i % 2 == 0 {
                outer_radius
            } else {
                inner_radius
            };
            let angle = (i as f64) * std::f64::consts::PI / 5.0 - std::f64::consts::PI / 2.0;
            let px = cx + radius * angle.cos();
            let py = cy + radius * angle.sin();

            if i == 0 {
                self.ctx.move_to(px, py);
            } else {
                self.ctx.line_to(px, py);
            }
        }

        self.ctx.close_path();
        self.ctx.fill();
        self.ctx.stroke();
    }

    /// Draw a triangle icon (up or down)
    fn draw_triangle_icon(
        &self,
        icon_index: usize,
        num_icons: usize,
        x: f64,
        y: f64,
        size: f64,
        color: &str,
    ) {
        let cx = x + size / 2.0;
        let cy = y + size / 2.0;
        let half = size / 2.0 - 2.0;

        self.ctx.set_fill_style_str(color);
        self.ctx.begin_path();

        let is_up = icon_index >= num_icons / 2;

        if is_up {
            // Up triangle
            self.ctx.move_to(cx, cy - half);
            self.ctx.line_to(cx + half, cy + half);
            self.ctx.line_to(cx - half, cy + half);
        } else {
            // Down triangle
            self.ctx.move_to(cx, cy + half);
            self.ctx.line_to(cx + half, cy - half);
            self.ctx.line_to(cx - half, cy - half);
        }

        self.ctx.close_path();
        self.ctx.fill();
    }
}
