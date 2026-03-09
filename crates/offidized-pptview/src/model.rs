//! JSON-serializable view model for presentation rendering.
//!
//! All measurements are pre-converted to CSS points. The TypeScript renderer
//! consumes these structs as plain JSON objects.

use serde::Serialize;

/// Top-level view model for a PowerPoint presentation.
#[derive(Debug, Clone, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct PresentationViewModel {
    /// Ordered slides.
    pub slides: Vec<SlideModel>,
    /// Slide width in CSS points (default 720pt = 10").
    pub slide_width_pt: f64,
    /// Slide height in CSS points (default 540pt = 7.5").
    pub slide_height_pt: f64,
    /// Embedded images as base64 data URIs.
    pub images: Vec<ImageModel>,
}

/// A single slide.
#[derive(Debug, Clone, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct SlideModel {
    /// Shapes on this slide (in z-order).
    pub shapes: Vec<ShapeModel>,
    /// Slide background fill.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub background: Option<BackgroundModel>,
    /// Notes text.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub notes: Option<String>,
    /// Whether this slide is hidden.
    #[serde(skip_serializing_if = "is_false")]
    pub hidden: bool,
}

/// A shape on a slide.
#[derive(Debug, Clone, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct ShapeModel {
    /// X position in CSS points.
    pub x_pt: f64,
    /// Y position in CSS points.
    pub y_pt: f64,
    /// Width in CSS points.
    pub width_pt: f64,
    /// Height in CSS points.
    pub height_pt: f64,
    /// Rotation in CSS degrees.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub rotation: Option<f64>,
    /// Shape name.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub name: Option<String>,
    /// Preset geometry name (e.g. "rect", "ellipse", "roundRect").
    #[serde(skip_serializing_if = "Option::is_none")]
    pub preset_geometry: Option<String>,
    /// Shape fill.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub fill: Option<ShapeFillModel>,
    /// Shape outline.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub outline: Option<OutlineModel>,
    /// Text body (paragraphs with runs).
    #[serde(skip_serializing_if = "Option::is_none")]
    pub text: Option<TextBodyModel>,
    /// Index into the presentation images array (for picture shapes).
    #[serde(skip_serializing_if = "Option::is_none")]
    pub image_index: Option<usize>,
    /// Table content.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub table: Option<TableModel>,
    /// Whether the shape is hidden.
    #[serde(skip_serializing_if = "is_false")]
    pub hidden: bool,
}

/// Text body within a shape.
#[derive(Debug, Clone, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct TextBodyModel {
    /// Paragraphs.
    pub paragraphs: Vec<TextParagraphModel>,
    /// Vertical anchor: "top", "middle", "bottom".
    #[serde(skip_serializing_if = "Option::is_none")]
    pub anchor: Option<String>,
    /// Text insets in CSS points (left, top, right, bottom).
    #[serde(skip_serializing_if = "Option::is_none")]
    pub insets: Option<InsetsModel>,
}

/// Text insets (padding) in CSS points.
#[derive(Debug, Clone, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct InsetsModel {
    /// Left inset in points.
    pub left_pt: f64,
    /// Top inset in points.
    pub top_pt: f64,
    /// Right inset in points.
    pub right_pt: f64,
    /// Bottom inset in points.
    pub bottom_pt: f64,
}

/// A paragraph within a text body.
#[derive(Debug, Clone, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct TextParagraphModel {
    /// Runs of text.
    pub runs: Vec<TextRunModel>,
    /// Horizontal alignment: "left", "center", "right", "justify".
    #[serde(skip_serializing_if = "Option::is_none")]
    pub alignment: Option<String>,
    /// Indentation level (0-based).
    #[serde(skip_serializing_if = "Option::is_none")]
    pub level: Option<u32>,
    /// Spacing before in points.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub spacing_before_pt: Option<f64>,
    /// Spacing after in points.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub spacing_after_pt: Option<f64>,
    /// Line spacing multiplier (e.g. 1.0 = single, 1.5 = 1.5 spacing).
    #[serde(skip_serializing_if = "Option::is_none")]
    pub line_spacing: Option<f64>,
    /// Bullet information.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub bullet: Option<BulletModel>,
}

/// A run of styled text.
#[derive(Debug, Clone, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct TextRunModel {
    /// The text content.
    pub text: String,
    /// Bold formatting.
    #[serde(skip_serializing_if = "is_false")]
    pub bold: bool,
    /// Italic formatting.
    #[serde(skip_serializing_if = "is_false")]
    pub italic: bool,
    /// Underline formatting.
    #[serde(skip_serializing_if = "is_false")]
    pub underline: bool,
    /// Strikethrough formatting.
    #[serde(skip_serializing_if = "is_false")]
    pub strikethrough: bool,
    /// Font family name.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub font_family: Option<String>,
    /// Font size in CSS points.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub font_size_pt: Option<f64>,
    /// Text color as hex (e.g. "FF0000").
    #[serde(skip_serializing_if = "Option::is_none")]
    pub color: Option<String>,
    /// Hyperlink URL.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub hyperlink: Option<String>,
}

/// Bullet/numbering info for a paragraph.
#[derive(Debug, Clone, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct BulletModel {
    /// Bullet character (e.g. "\u{2022}").
    #[serde(skip_serializing_if = "Option::is_none")]
    pub char: Option<String>,
    /// Auto-number type name (e.g. "arabicPeriod").
    #[serde(skip_serializing_if = "Option::is_none")]
    pub auto_num_type: Option<String>,
    /// Bullet font family.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub font_family: Option<String>,
    /// Bullet color as hex.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub color: Option<String>,
}

/// Shape fill.
#[derive(Debug, Clone, Serialize)]
#[serde(tag = "type", rename_all = "camelCase")]
pub enum ShapeFillModel {
    /// Solid color fill.
    #[serde(rename = "solid")]
    Solid {
        /// Color as hex (e.g. "FF0000").
        color: String,
    },
    /// Gradient fill.
    #[serde(rename = "gradient")]
    Gradient {
        /// CSS gradient string (e.g. "linear-gradient(...)").
        css: String,
    },
    /// No fill.
    #[serde(rename = "none")]
    None,
}

/// Shape outline.
#[derive(Debug, Clone, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct OutlineModel {
    /// Width in CSS points.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub width_pt: Option<f64>,
    /// Color as hex.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub color: Option<String>,
    /// CSS dash style: "solid", "dashed", "dotted".
    #[serde(skip_serializing_if = "Option::is_none")]
    pub dash_style: Option<String>,
}

/// Table content within a shape.
#[derive(Debug, Clone, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct TableModel {
    /// Table rows.
    pub rows: Vec<TableRowModel>,
    /// Column widths in CSS points.
    pub column_widths_pt: Vec<f64>,
    /// Row heights in CSS points.
    pub row_heights_pt: Vec<f64>,
}

/// A table row.
#[derive(Debug, Clone, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct TableRowModel {
    /// Cells in this row.
    pub cells: Vec<TableCellModel>,
}

/// A table cell.
#[derive(Debug, Clone, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct TableCellModel {
    /// Cell text content.
    pub text: String,
    /// Fill color as hex.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub fill_color: Option<String>,
    /// Horizontal span (gridSpan).
    #[serde(skip_serializing_if = "is_one_u32")]
    pub grid_span: u32,
    /// Vertical span (rowSpan).
    #[serde(skip_serializing_if = "is_one_u32")]
    pub row_span: u32,
    /// Whether this cell is covered by a merge.
    #[serde(skip_serializing_if = "is_false")]
    pub v_merge: bool,
    /// Vertical alignment: "top", "middle", "bottom".
    #[serde(skip_serializing_if = "Option::is_none")]
    pub vertical_align: Option<String>,
}

/// Slide background.
#[derive(Debug, Clone, Serialize)]
#[serde(tag = "type", rename_all = "camelCase")]
pub enum BackgroundModel {
    /// Solid color background.
    #[serde(rename = "solid")]
    Solid {
        /// Color as hex.
        color: String,
    },
    /// Gradient background.
    #[serde(rename = "gradient")]
    Gradient {
        /// CSS gradient string.
        css: String,
    },
}

/// An embedded image.
#[derive(Debug, Clone, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct ImageModel {
    /// Base64 data URI (e.g. "data:image/png;base64,...").
    pub data_uri: String,
    /// MIME content type (e.g. "image/png").
    pub content_type: String,
}

// Helper functions for serde skip conditions.

fn is_false(v: &bool) -> bool {
    !(*v)
}

fn is_one_u32(v: &u32) -> bool {
    *v == 1
}
