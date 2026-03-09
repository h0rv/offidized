//! Shape rendering for Canvas 2D backend.

use crate::layout::{SheetLayout, Viewport};
use crate::types::drawing::Drawing;

use super::renderer::{CanvasRenderer, SCROLLBAR_SIZE};

impl CanvasRenderer {
    /// Render shapes embedded in the sheet
    #[allow(clippy::cast_possible_truncation)]
    pub(super) fn render_shapes(
        &self,
        drawings: &[Drawing],
        layout: &SheetLayout,
        viewport: &Viewport,
    ) {
        for drawing in drawings {
            // Only render shapes, skip pictures and charts
            if drawing.drawing_type != "shape" {
                continue;
            }

            // Calculate the shape position and size based on anchor type
            let (x, y, width, height) = match drawing.anchor_type.as_str() {
                "twoCellAnchor" => self.calculate_two_cell_anchor_bounds(drawing, layout),
                "oneCellAnchor" => self.calculate_one_cell_anchor_bounds(drawing, layout),
                "absoluteAnchor" => self.calculate_absolute_anchor_bounds(drawing),
                _ => continue,
            };

            // Skip shapes with zero or negative size
            if width <= 0.0 || height <= 0.0 {
                continue;
            }

            // Convert to screen coordinates
            let from_col = drawing.from_col.unwrap_or(0);
            let from_row = drawing.from_row.unwrap_or(0);
            let (screen_x, screen_y) = viewport.to_screen_frozen(x, y, from_row, from_col, layout);
            let screen_width = width * viewport.scale;
            let screen_height = height * viewport.scale;

            // Skip if shape is completely off-screen
            let content_width = f64::from(viewport.width) - SCROLLBAR_SIZE;
            let content_height = f64::from(viewport.height) - SCROLLBAR_SIZE;
            if f64::from(screen_x) > content_width
                || f64::from(screen_y) > content_height
                || f64::from(screen_x + screen_width) < 0.0
                || f64::from(screen_y + screen_height) < 0.0
            {
                continue;
            }

            let sx = f64::from(screen_x);
            let sy = f64::from(screen_y);
            let sw = f64::from(screen_width);
            let sh = f64::from(screen_height);

            // Get shape type (default to rectangle if not specified)
            let shape_type = drawing.shape_type.as_deref().unwrap_or("rect");

            // Apply rotation if present
            let has_rotation = drawing.rotation.is_some() && drawing.rotation != Some(0);
            if has_rotation {
                self.ctx.save();
                // Rotation is in 1/60000th of a degree (EMU angle units)
                let rotation_deg = drawing.rotation.unwrap_or(0) as f64 / 60000.0;
                let rotation_rad = rotation_deg * std::f64::consts::PI / 180.0;

                // Translate to center, rotate, translate back
                let cx = sx + sw / 2.0;
                let cy = sy + sh / 2.0;
                let _ = self.ctx.translate(cx, cy);
                let _ = self.ctx.rotate(rotation_rad);
                let _ = self.ctx.translate(-cx, -cy);
            }

            // Render the shape based on type
            self.render_shape(drawing, shape_type, sx, sy, sw, sh);

            // Restore context if rotation was applied
            if has_rotation {
                self.ctx.restore();
            }
        }
    }

    /// Render a single shape
    fn render_shape(&self, drawing: &Drawing, shape_type: &str, x: f64, y: f64, w: f64, h: f64) {
        // Default colors
        let fill_color = drawing.fill_color.as_deref().unwrap_or("#4472C4");
        let line_color = drawing.line_color.as_deref().unwrap_or("#2F528F");
        let line_width = 1.0;

        match shape_type {
            // Rectangles
            "rect" | "rectangle" => {
                self.render_rectangle(x, y, w, h, fill_color, line_color, line_width);
            }
            // Rounded rectangles
            "roundRect" | "roundedRect" | "snip1Rect" | "snip2DiagRect" | "snip2SameRect"
            | "snipRoundRect" | "round1Rect" | "round2DiagRect" | "round2SameRect" => {
                let radius = (w.min(h) * 0.15).min(20.0); // 15% of smaller dimension, max 20px
                self.render_rounded_rectangle(
                    x, y, w, h, radius, fill_color, line_color, line_width,
                );
            }
            // Ellipses and circles
            "ellipse" | "oval" | "circle" => {
                self.render_ellipse(x, y, w, h, fill_color, line_color, line_width);
            }
            // Text boxes - rectangle with text
            "textBox" => {
                self.render_text_box(x, y, w, h, fill_color, line_color, line_width);
            }
            // Lines
            "line" | "straightConnector1" | "bentConnector2" | "bentConnector3"
            | "bentConnector4" | "bentConnector5" | "curvedConnector2" | "curvedConnector3"
            | "curvedConnector4" | "curvedConnector5" => {
                self.render_line(x, y, w, h, line_color, line_width, false);
            }
            // Arrows
            "rightArrow"
            | "leftArrow"
            | "upArrow"
            | "downArrow"
            | "leftRightArrow"
            | "upDownArrow"
            | "quadArrow"
            | "leftRightUpArrow"
            | "bentArrow"
            | "uturnArrow"
            | "leftUpArrow"
            | "bentUpArrow"
            | "curvedRightArrow"
            | "curvedLeftArrow"
            | "curvedUpArrow"
            | "curvedDownArrow"
            | "stripedRightArrow"
            | "notchedRightArrow"
            | "homePlate"
            | "chevron"
            | "rightArrowCallout"
            | "leftArrowCallout"
            | "upArrowCallout"
            | "downArrowCallout"
            | "leftRightArrowCallout"
            | "upDownArrowCallout"
            | "quadArrowCallout"
            | "circularArrow" => {
                self.render_arrow_shape(shape_type, x, y, w, h, fill_color, line_color, line_width);
            }
            // Triangles
            "triangle" | "rtTriangle" | "isoscelesTriangle" => {
                self.render_triangle(x, y, w, h, fill_color, line_color, line_width);
            }
            // Pentagons/Hexagons/Stars
            "pentagon" => {
                self.render_polygon(x, y, w, h, 5, fill_color, line_color, line_width);
            }
            "hexagon" => {
                self.render_polygon(x, y, w, h, 6, fill_color, line_color, line_width);
            }
            "star4" | "star5" | "star6" | "star7" | "star8" | "star10" | "star12" | "star16"
            | "star24" | "star32" => {
                let points = shape_type
                    .strip_prefix("star")
                    .and_then(|s| s.parse::<usize>().ok())
                    .unwrap_or(5);
                self.render_star(x, y, w, h, points, fill_color, line_color, line_width);
            }
            // Diamonds
            "diamond" => {
                self.render_diamond(x, y, w, h, fill_color, line_color, line_width);
            }
            // Parallelograms
            "parallelogram" => {
                self.render_parallelogram(x, y, w, h, fill_color, line_color, line_width);
            }
            // Trapezoids
            "trapezoid" => {
                self.render_trapezoid(x, y, w, h, fill_color, line_color, line_width);
            }
            // Plus signs
            "plus" | "cross" => {
                self.render_plus(x, y, w, h, fill_color, line_color, line_width);
            }
            // Default: render as rectangle
            _ => {
                self.render_rectangle(x, y, w, h, fill_color, line_color, line_width);
            }
        }

        // Render text content if present (for any shape)
        if let Some(ref text) = drawing.text_content {
            if !text.is_empty() {
                self.render_shape_text(text, x, y, w, h);
            }
        }
    }

    /// Render a simple rectangle
    #[allow(clippy::too_many_arguments)]
    fn render_rectangle(
        &self,
        x: f64,
        y: f64,
        w: f64,
        h: f64,
        fill: &str,
        stroke: &str,
        line_width: f64,
    ) {
        self.ctx.set_fill_style_str(fill);
        self.ctx.fill_rect(x, y, w, h);
        self.ctx.set_stroke_style_str(stroke);
        self.ctx.set_line_width(line_width);
        self.ctx.stroke_rect(x, y, w, h);
    }

    /// Render a rounded rectangle
    #[allow(clippy::too_many_arguments)]
    fn render_rounded_rectangle(
        &self,
        x: f64,
        y: f64,
        w: f64,
        h: f64,
        radius: f64,
        fill: &str,
        stroke: &str,
        line_width: f64,
    ) {
        self.ctx.begin_path();
        self.ctx.move_to(x + radius, y);
        let _ = self.ctx.arc_to(x + w, y, x + w, y + h, radius);
        let _ = self.ctx.arc_to(x + w, y + h, x, y + h, radius);
        let _ = self.ctx.arc_to(x, y + h, x, y, radius);
        let _ = self.ctx.arc_to(x, y, x + w, y, radius);
        self.ctx.close_path();

        self.ctx.set_fill_style_str(fill);
        self.ctx.fill();
        self.ctx.set_stroke_style_str(stroke);
        self.ctx.set_line_width(line_width);
        self.ctx.stroke();
    }

    /// Render an ellipse/oval
    #[allow(clippy::too_many_arguments)]
    fn render_ellipse(
        &self,
        x: f64,
        y: f64,
        w: f64,
        h: f64,
        fill: &str,
        stroke: &str,
        line_width: f64,
    ) {
        let cx = x + w / 2.0;
        let cy = y + h / 2.0;
        let rx = w / 2.0;
        let ry = h / 2.0;

        self.ctx.begin_path();
        let _ = self
            .ctx
            .ellipse(cx, cy, rx, ry, 0.0, 0.0, 2.0 * std::f64::consts::PI);

        self.ctx.set_fill_style_str(fill);
        self.ctx.fill();
        self.ctx.set_stroke_style_str(stroke);
        self.ctx.set_line_width(line_width);
        self.ctx.stroke();
    }

    /// Render a text box (rectangle with text)
    #[allow(clippy::too_many_arguments)]
    fn render_text_box(
        &self,
        x: f64,
        y: f64,
        w: f64,
        h: f64,
        fill: &str,
        stroke: &str,
        line_width: f64,
    ) {
        // Draw the rectangle background
        self.render_rectangle(x, y, w, h, fill, stroke, line_width);
    }

    /// Render a line between two points
    #[allow(clippy::too_many_arguments)]
    fn render_line(
        &self,
        x: f64,
        y: f64,
        w: f64,
        h: f64,
        stroke: &str,
        line_width: f64,
        with_arrow: bool,
    ) {
        let x1 = x;
        let y1 = y;
        let x2 = x + w;
        let y2 = y + h;

        self.ctx.begin_path();
        self.ctx.set_stroke_style_str(stroke);
        self.ctx.set_line_width(line_width);
        self.ctx.move_to(x1, y1);
        self.ctx.line_to(x2, y2);
        self.ctx.stroke();

        // Draw arrowhead if requested
        if with_arrow {
            self.draw_arrowhead(x1, y1, x2, y2, stroke);
        }
    }

    /// Draw an arrowhead at the end of a line
    fn draw_arrowhead(&self, _x1: f64, _y1: f64, x2: f64, y2: f64, color: &str) {
        let angle = (y2 - _y1).atan2(x2 - _x1);
        let arrow_length = 10.0;
        let arrow_width = 6.0;

        self.ctx.save();
        let _ = self.ctx.translate(x2, y2);
        let _ = self.ctx.rotate(angle);

        self.ctx.begin_path();
        self.ctx.move_to(0.0, 0.0);
        self.ctx.line_to(-arrow_length, -arrow_width / 2.0);
        self.ctx.line_to(-arrow_length, arrow_width / 2.0);
        self.ctx.close_path();

        self.ctx.set_fill_style_str(color);
        self.ctx.fill();

        self.ctx.restore();
    }

    /// Render an arrow shape
    #[allow(clippy::too_many_arguments)]
    fn render_arrow_shape(
        &self,
        shape_type: &str,
        x: f64,
        y: f64,
        w: f64,
        h: f64,
        fill: &str,
        stroke: &str,
        line_width: f64,
    ) {
        // Determine arrow direction
        let (is_right, is_left, is_up, is_down) = (
            shape_type.contains("right") || shape_type.contains("Right"),
            shape_type.contains("left") || shape_type.contains("Left"),
            shape_type.contains("up") || shape_type.contains("Up"),
            shape_type.contains("down") || shape_type.contains("Down"),
        );

        // Draw a simple arrow shape
        self.ctx.begin_path();

        if is_right && !is_left {
            // Right arrow
            let arrow_head = w * 0.3;
            let body_height = h * 0.4;
            let body_y = y + (h - body_height) / 2.0;

            self.ctx.move_to(x, body_y);
            self.ctx.line_to(x + w - arrow_head, body_y);
            self.ctx.line_to(x + w - arrow_head, y);
            self.ctx.line_to(x + w, y + h / 2.0);
            self.ctx.line_to(x + w - arrow_head, y + h);
            self.ctx.line_to(x + w - arrow_head, body_y + body_height);
            self.ctx.line_to(x, body_y + body_height);
        } else if is_left && !is_right {
            // Left arrow
            let arrow_head = w * 0.3;
            let body_height = h * 0.4;
            let body_y = y + (h - body_height) / 2.0;

            self.ctx.move_to(x + w, body_y);
            self.ctx.line_to(x + arrow_head, body_y);
            self.ctx.line_to(x + arrow_head, y);
            self.ctx.line_to(x, y + h / 2.0);
            self.ctx.line_to(x + arrow_head, y + h);
            self.ctx.line_to(x + arrow_head, body_y + body_height);
            self.ctx.line_to(x + w, body_y + body_height);
        } else if is_up && !is_down {
            // Up arrow
            let arrow_head = h * 0.3;
            let body_width = w * 0.4;
            let body_x = x + (w - body_width) / 2.0;

            self.ctx.move_to(body_x, y + h);
            self.ctx.line_to(body_x, y + arrow_head);
            self.ctx.line_to(x, y + arrow_head);
            self.ctx.line_to(x + w / 2.0, y);
            self.ctx.line_to(x + w, y + arrow_head);
            self.ctx.line_to(body_x + body_width, y + arrow_head);
            self.ctx.line_to(body_x + body_width, y + h);
        } else if is_down && !is_up {
            // Down arrow
            let arrow_head = h * 0.3;
            let body_width = w * 0.4;
            let body_x = x + (w - body_width) / 2.0;

            self.ctx.move_to(body_x, y);
            self.ctx.line_to(body_x, y + h - arrow_head);
            self.ctx.line_to(x, y + h - arrow_head);
            self.ctx.line_to(x + w / 2.0, y + h);
            self.ctx.line_to(x + w, y + h - arrow_head);
            self.ctx.line_to(body_x + body_width, y + h - arrow_head);
            self.ctx.line_to(body_x + body_width, y);
        } else {
            // Default: right arrow
            let arrow_head = w * 0.3;
            let body_height = h * 0.4;
            let body_y = y + (h - body_height) / 2.0;

            self.ctx.move_to(x, body_y);
            self.ctx.line_to(x + w - arrow_head, body_y);
            self.ctx.line_to(x + w - arrow_head, y);
            self.ctx.line_to(x + w, y + h / 2.0);
            self.ctx.line_to(x + w - arrow_head, y + h);
            self.ctx.line_to(x + w - arrow_head, body_y + body_height);
            self.ctx.line_to(x, body_y + body_height);
        }

        self.ctx.close_path();
        self.ctx.set_fill_style_str(fill);
        self.ctx.fill();
        self.ctx.set_stroke_style_str(stroke);
        self.ctx.set_line_width(line_width);
        self.ctx.stroke();
    }

    /// Render a triangle
    #[allow(clippy::too_many_arguments)]
    fn render_triangle(
        &self,
        x: f64,
        y: f64,
        w: f64,
        h: f64,
        fill: &str,
        stroke: &str,
        line_width: f64,
    ) {
        self.ctx.begin_path();
        self.ctx.move_to(x + w / 2.0, y);
        self.ctx.line_to(x + w, y + h);
        self.ctx.line_to(x, y + h);
        self.ctx.close_path();

        self.ctx.set_fill_style_str(fill);
        self.ctx.fill();
        self.ctx.set_stroke_style_str(stroke);
        self.ctx.set_line_width(line_width);
        self.ctx.stroke();
    }

    /// Render a regular polygon
    #[allow(clippy::too_many_arguments)]
    fn render_polygon(
        &self,
        x: f64,
        y: f64,
        w: f64,
        h: f64,
        sides: usize,
        fill: &str,
        stroke: &str,
        line_width: f64,
    ) {
        let cx = x + w / 2.0;
        let cy = y + h / 2.0;
        let rx = w / 2.0;
        let ry = h / 2.0;

        self.ctx.begin_path();
        for i in 0..sides {
            let angle =
                (i as f64 * 2.0 * std::f64::consts::PI / sides as f64) - std::f64::consts::PI / 2.0;
            let px = cx + rx * angle.cos();
            let py = cy + ry * angle.sin();
            if i == 0 {
                self.ctx.move_to(px, py);
            } else {
                self.ctx.line_to(px, py);
            }
        }
        self.ctx.close_path();

        self.ctx.set_fill_style_str(fill);
        self.ctx.fill();
        self.ctx.set_stroke_style_str(stroke);
        self.ctx.set_line_width(line_width);
        self.ctx.stroke();
    }

    /// Render a star shape
    #[allow(clippy::too_many_arguments)]
    fn render_star(
        &self,
        x: f64,
        y: f64,
        w: f64,
        h: f64,
        points: usize,
        fill: &str,
        stroke: &str,
        line_width: f64,
    ) {
        let cx = x + w / 2.0;
        let cy = y + h / 2.0;
        let outer_rx = w / 2.0;
        let outer_ry = h / 2.0;
        let inner_rx = outer_rx * 0.4;
        let inner_ry = outer_ry * 0.4;

        self.ctx.begin_path();
        for i in 0..(points * 2) {
            let angle =
                (i as f64 * std::f64::consts::PI / points as f64) - std::f64::consts::PI / 2.0;
            let (rx, ry) = if i % 2 == 0 {
                (outer_rx, outer_ry)
            } else {
                (inner_rx, inner_ry)
            };
            let px = cx + rx * angle.cos();
            let py = cy + ry * angle.sin();
            if i == 0 {
                self.ctx.move_to(px, py);
            } else {
                self.ctx.line_to(px, py);
            }
        }
        self.ctx.close_path();

        self.ctx.set_fill_style_str(fill);
        self.ctx.fill();
        self.ctx.set_stroke_style_str(stroke);
        self.ctx.set_line_width(line_width);
        self.ctx.stroke();
    }

    /// Render a diamond shape
    #[allow(clippy::too_many_arguments)]
    fn render_diamond(
        &self,
        x: f64,
        y: f64,
        w: f64,
        h: f64,
        fill: &str,
        stroke: &str,
        line_width: f64,
    ) {
        let cx = x + w / 2.0;
        let cy = y + h / 2.0;

        self.ctx.begin_path();
        self.ctx.move_to(cx, y);
        self.ctx.line_to(x + w, cy);
        self.ctx.line_to(cx, y + h);
        self.ctx.line_to(x, cy);
        self.ctx.close_path();

        self.ctx.set_fill_style_str(fill);
        self.ctx.fill();
        self.ctx.set_stroke_style_str(stroke);
        self.ctx.set_line_width(line_width);
        self.ctx.stroke();
    }

    /// Render a parallelogram
    #[allow(clippy::too_many_arguments)]
    fn render_parallelogram(
        &self,
        x: f64,
        y: f64,
        w: f64,
        h: f64,
        fill: &str,
        stroke: &str,
        line_width: f64,
    ) {
        let offset = w * 0.2;

        self.ctx.begin_path();
        self.ctx.move_to(x + offset, y);
        self.ctx.line_to(x + w, y);
        self.ctx.line_to(x + w - offset, y + h);
        self.ctx.line_to(x, y + h);
        self.ctx.close_path();

        self.ctx.set_fill_style_str(fill);
        self.ctx.fill();
        self.ctx.set_stroke_style_str(stroke);
        self.ctx.set_line_width(line_width);
        self.ctx.stroke();
    }

    /// Render a trapezoid
    #[allow(clippy::too_many_arguments)]
    fn render_trapezoid(
        &self,
        x: f64,
        y: f64,
        w: f64,
        h: f64,
        fill: &str,
        stroke: &str,
        line_width: f64,
    ) {
        let offset = w * 0.15;

        self.ctx.begin_path();
        self.ctx.move_to(x + offset, y);
        self.ctx.line_to(x + w - offset, y);
        self.ctx.line_to(x + w, y + h);
        self.ctx.line_to(x, y + h);
        self.ctx.close_path();

        self.ctx.set_fill_style_str(fill);
        self.ctx.fill();
        self.ctx.set_stroke_style_str(stroke);
        self.ctx.set_line_width(line_width);
        self.ctx.stroke();
    }

    /// Render a plus/cross shape
    #[allow(clippy::too_many_arguments)]
    fn render_plus(
        &self,
        x: f64,
        y: f64,
        w: f64,
        h: f64,
        fill: &str,
        stroke: &str,
        line_width: f64,
    ) {
        let arm_width = w * 0.3;
        let arm_height = h * 0.3;
        let hx = x + (w - arm_width) / 2.0;
        let hy = y + (h - arm_height) / 2.0;

        self.ctx.begin_path();
        // Top
        self.ctx.move_to(hx, y);
        self.ctx.line_to(hx + arm_width, y);
        // Right top
        self.ctx.line_to(hx + arm_width, hy);
        self.ctx.line_to(x + w, hy);
        // Right bottom
        self.ctx.line_to(x + w, hy + arm_height);
        self.ctx.line_to(hx + arm_width, hy + arm_height);
        // Bottom
        self.ctx.line_to(hx + arm_width, y + h);
        self.ctx.line_to(hx, y + h);
        // Left bottom
        self.ctx.line_to(hx, hy + arm_height);
        self.ctx.line_to(x, hy + arm_height);
        // Left top
        self.ctx.line_to(x, hy);
        self.ctx.line_to(hx, hy);
        self.ctx.close_path();

        self.ctx.set_fill_style_str(fill);
        self.ctx.fill();
        self.ctx.set_stroke_style_str(stroke);
        self.ctx.set_line_width(line_width);
        self.ctx.stroke();
    }

    /// Render text inside a shape
    fn render_shape_text(&self, text: &str, x: f64, y: f64, w: f64, h: f64) {
        let padding = 4.0;
        let text_x = x + padding;
        let text_y = y + h / 2.0;
        let max_width = w - padding * 2.0;

        self.ctx.set_fill_style_str("#000000");
        self.ctx.set_font("11px Calibri, Arial, sans-serif");
        self.ctx.set_text_align("left");
        self.ctx.set_text_baseline("middle");

        // Simple text rendering (truncate if too long)
        let metrics = self.ctx.measure_text(text).ok();
        let text_width = metrics.map(|m| m.width()).unwrap_or(0.0);

        if text_width <= max_width {
            let _ = self.ctx.fill_text(text, text_x, text_y);
        } else {
            // Truncate with ellipsis
            let mut truncated = text.to_string();
            while !truncated.is_empty() {
                truncated.pop();
                let test = format!("{}...", truncated);
                if let Ok(m) = self.ctx.measure_text(&test) {
                    if m.width() <= max_width {
                        let _ = self.ctx.fill_text(&test, text_x, text_y);
                        break;
                    }
                }
            }
        }
    }
}
