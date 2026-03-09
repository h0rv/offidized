//! JSON-serializable view model for document rendering.
//!
//! All measurements are pre-converted to CSS points. The TypeScript renderer
//! consumes these structs as plain JSON objects.

use serde::Serialize;

/// Top-level view model for a Word document.
#[derive(Debug, Clone, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct DocViewModel {
    /// Ordered body items (paragraphs and tables).
    pub body: Vec<BodyItem>,
    /// Section definitions (page layout, margins, headers/footers).
    pub sections: Vec<SectionModel>,
    /// Embedded images as base64 data URIs.
    pub images: Vec<ImageModel>,
    /// Footnotes referenced in the document.
    pub footnotes: Vec<FootnoteModel>,
    /// Endnotes referenced in the document.
    pub endnotes: Vec<EndnoteModel>,
}

/// A body-level item: either a paragraph or a table.
#[derive(Debug, Clone, Serialize)]
#[serde(tag = "type", rename_all = "camelCase")]
pub enum BodyItem {
    /// A paragraph.
    #[serde(rename = "paragraph")]
    Paragraph(ParagraphModel),
    /// A table.
    #[serde(rename = "table")]
    Table(TableModel),
}

/// A rendered paragraph.
#[derive(Debug, Clone, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct ParagraphModel {
    /// Inline runs within this paragraph.
    pub runs: Vec<RunModel>,
    /// Heading level (1–9), or `None` for body text.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub heading_level: Option<u8>,
    /// Paragraph alignment.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub alignment: Option<String>,
    /// Spacing before this paragraph, in points.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub spacing_before_pt: Option<f64>,
    /// Spacing after this paragraph, in points.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub spacing_after_pt: Option<f64>,
    /// Line spacing value and rule.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub line_spacing: Option<LineSpacingModel>,
    /// Indentation settings.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub indents: Option<IndentsModel>,
    /// Numbering (bullet/list) info if this paragraph is part of a list.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub numbering: Option<NumberingModel>,
    /// Paragraph borders.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub borders: Option<BordersModel>,
    /// Background shading color (hex, e.g. "FFFF00").
    #[serde(skip_serializing_if = "Option::is_none")]
    pub shading_color: Option<String>,
    /// Whether a page break appears before this paragraph.
    #[serde(skip_serializing_if = "is_false")]
    pub page_break_before: bool,
    /// Keep with next paragraph on same page.
    #[serde(skip_serializing_if = "is_false")]
    pub keep_next: bool,
    /// Keep all lines of this paragraph on the same page.
    #[serde(skip_serializing_if = "is_false")]
    pub keep_lines: bool,
    /// Index into `sections` array.
    pub section_index: usize,
    /// Whether this paragraph ends its section (has section properties).
    #[serde(skip_serializing_if = "is_false")]
    pub ends_section: bool,
    /// Paragraph style ID.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub style_id: Option<String>,
}

/// An inline run of text with formatting.
#[derive(Debug, Clone, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct RunModel {
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
    /// Underline style (e.g. "single", "double", "wavy").
    #[serde(skip_serializing_if = "Option::is_none")]
    pub underline_type: Option<String>,
    /// Strikethrough formatting.
    #[serde(skip_serializing_if = "is_false")]
    pub strikethrough: bool,
    /// Superscript.
    #[serde(skip_serializing_if = "is_false")]
    pub superscript: bool,
    /// Subscript.
    #[serde(skip_serializing_if = "is_false")]
    pub subscript: bool,
    /// Small caps.
    #[serde(skip_serializing_if = "is_false")]
    pub small_caps: bool,
    /// Font family name.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub font_family: Option<String>,
    /// Font size in CSS points.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub font_size_pt: Option<f64>,
    /// Text color as hex (e.g. "FF0000").
    #[serde(skip_serializing_if = "Option::is_none")]
    pub color: Option<String>,
    /// Highlight color name (e.g. "yellow").
    #[serde(skip_serializing_if = "Option::is_none")]
    pub highlight: Option<String>,
    /// Hyperlink URL target.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub hyperlink: Option<String>,
    /// Hyperlink tooltip text.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub hyperlink_tooltip: Option<String>,
    /// Inline image reference.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub inline_image: Option<InlineImageModel>,
    /// Floating image reference.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub floating_image: Option<FloatingImageModel>,
    /// Footnote reference ID.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub footnote_ref: Option<u32>,
    /// Endnote reference ID.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub endnote_ref: Option<u32>,
    /// Run contains a tab character.
    #[serde(skip_serializing_if = "is_false")]
    pub has_tab: bool,
    /// Run contains a line break.
    #[serde(skip_serializing_if = "is_false")]
    pub has_break: bool,
}

/// An inline image within a run.
#[derive(Debug, Clone, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct InlineImageModel {
    /// Index into the `images` array.
    pub image_index: usize,
    /// Display width in points.
    pub width_pt: f64,
    /// Display height in points.
    pub height_pt: f64,
    /// Optional image name.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub name: Option<String>,
    /// Optional image description (alt text).
    #[serde(skip_serializing_if = "Option::is_none")]
    pub description: Option<String>,
}

/// A floating image (anchored, not inline).
#[derive(Debug, Clone, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct FloatingImageModel {
    /// Index into the `images` array.
    pub image_index: usize,
    /// Display width in points.
    pub width_pt: f64,
    /// Display height in points.
    pub height_pt: f64,
    /// Horizontal offset in points.
    pub offset_x_pt: f64,
    /// Vertical offset in points.
    pub offset_y_pt: f64,
    /// Optional image name.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub name: Option<String>,
    /// Optional image description (alt text).
    #[serde(skip_serializing_if = "Option::is_none")]
    pub description: Option<String>,
    /// Text wrapping type.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub wrap_type: Option<String>,
}

/// A table in the document body.
#[derive(Debug, Clone, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct TableModel {
    /// Table rows.
    pub rows: Vec<TableRowModel>,
    /// Total table width in points.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub width_pt: Option<f64>,
    /// Table alignment.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub alignment: Option<String>,
    /// Column widths in points.
    pub column_widths_pt: Vec<f64>,
    /// Table borders.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub borders: Option<BordersModel>,
    /// Section index this table belongs to.
    pub section_index: usize,
}

/// A single table row.
#[derive(Debug, Clone, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct TableRowModel {
    /// Cells in this row.
    pub cells: Vec<TableCellModel>,
    /// Row height in points.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub height_pt: Option<f64>,
    /// Row height rule (`auto`, `atLeast`, `exact`) when present.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub height_rule: Option<String>,
}

/// A single table cell.
#[derive(Debug, Clone, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct TableCellModel {
    /// Cell text content.
    pub text: String,
    /// Horizontal span (colspan).
    #[serde(skip_serializing_if = "is_one")]
    pub col_span: usize,
    /// Vertical span (rowspan), computed from vertical merge.
    #[serde(skip_serializing_if = "is_one")]
    pub row_span: usize,
    /// Background shading color.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub shading_color: Option<String>,
    /// Vertical text alignment.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub vertical_align: Option<String>,
    /// Cell width in points.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub width_pt: Option<f64>,
    /// Cell borders.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub borders: Option<BordersModel>,
    /// Whether this cell is covered by a merge and should not be rendered.
    #[serde(skip_serializing_if = "is_false")]
    pub is_covered: bool,
}

/// Page section layout.
#[derive(Debug, Clone, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct SectionModel {
    /// Page width in points.
    pub page_width_pt: f64,
    /// Page height in points.
    pub page_height_pt: f64,
    /// Page orientation.
    pub orientation: String,
    /// Page margins.
    pub margins: MarginsModel,
    /// Default header content (paragraphs as text).
    #[serde(skip_serializing_if = "Option::is_none")]
    pub header: Option<HeaderFooterModel>,
    /// Default footer content.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub footer: Option<HeaderFooterModel>,
    /// Number of text columns.
    #[serde(skip_serializing_if = "is_one_u16")]
    pub column_count: u16,
}

/// Page margins in points.
#[derive(Debug, Clone, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct MarginsModel {
    /// Top margin.
    pub top: f64,
    /// Right margin.
    pub right: f64,
    /// Bottom margin.
    pub bottom: f64,
    /// Left margin.
    pub left: f64,
}

/// Header or footer content.
#[derive(Debug, Clone, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct HeaderFooterModel {
    /// Paragraphs within the header/footer.
    pub paragraphs: Vec<ParagraphModel>,
}

/// Line spacing configuration.
#[derive(Debug, Clone, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct LineSpacingModel {
    /// Spacing value in points (or a multiplier for "auto" rule).
    pub value: f64,
    /// Rule: "auto", "exact", or "atLeast".
    pub rule: String,
}

/// Paragraph indentation.
#[derive(Debug, Clone, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct IndentsModel {
    /// Left indent in points.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub left_pt: Option<f64>,
    /// Right indent in points.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub right_pt: Option<f64>,
    /// First-line indent in points.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub first_line_pt: Option<f64>,
    /// Hanging indent in points (negative first-line).
    #[serde(skip_serializing_if = "Option::is_none")]
    pub hanging_pt: Option<f64>,
}

/// Numbering (list) information for a paragraph.
#[derive(Debug, Clone, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct NumberingModel {
    /// Numbering instance ID.
    pub num_id: u32,
    /// Nesting level (0-based).
    pub level: u8,
    /// Format string (e.g. "decimal", "bullet", "lowerLetter").
    pub format: String,
    /// Resolved display text (e.g. "1.", "a)", bullet char).
    pub text: String,
}

/// Border on all four sides.
#[derive(Debug, Clone, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct BordersModel {
    /// Top border.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub top: Option<BorderModel>,
    /// Right border.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub right: Option<BorderModel>,
    /// Bottom border.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub bottom: Option<BorderModel>,
    /// Left border.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub left: Option<BorderModel>,
}

/// A single border edge.
#[derive(Debug, Clone, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct BorderModel {
    /// Border style (e.g. "single", "double", "dashed").
    pub style: String,
    /// Border color as hex.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub color: Option<String>,
    /// Border width in points.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub width_pt: Option<f64>,
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

/// A footnote.
#[derive(Debug, Clone, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct FootnoteModel {
    /// Footnote ID.
    pub id: u32,
    /// Concatenated text content.
    pub text: String,
}

/// An endnote.
#[derive(Debug, Clone, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct EndnoteModel {
    /// Endnote ID.
    pub id: u32,
    /// Concatenated text content.
    pub text: String,
}

// Helper functions for serde skip conditions.

fn is_false(v: &bool) -> bool {
    !(*v)
}

fn is_one(v: &usize) -> bool {
    *v == 1
}

fn is_one_u16(v: &u16) -> bool {
    *v == 1
}
