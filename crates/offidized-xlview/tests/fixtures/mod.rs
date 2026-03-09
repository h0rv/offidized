//! Test fixtures for generating valid XLSX files in memory.
//!
//! This module provides builders for creating XLSX files programmatically,
//! useful for testing the offidized-xlview adapter pipeline with known inputs.
//!
//! # Example
//!
//! ```rust,ignore
//! use fixtures::{XlsxBuilder, StyleBuilder, SheetBuilder};
//!
//! let xlsx = XlsxBuilder::new()
//!     .sheet(SheetBuilder::new("Sheet1")
//!         .cell("A1", "Hello", Some(StyleBuilder::new().bold().build())))
//!     .build();
//!
//! let workbook = load_xlsx(&xlsx);
//! ```
#![allow(
    dead_code,
    clippy::unwrap_used,
    clippy::expect_used,
    clippy::indexing_slicing,
    clippy::float_cmp,
    clippy::panic,
    clippy::approx_constant,
    clippy::cast_possible_truncation,
    clippy::absurd_extreme_comparisons,
    clippy::cast_lossless
)]

use std::io::{Cursor, Write};
use zip::write::FileOptions;
use zip::ZipWriter;

// ============================================================================
// Style Builder
// ============================================================================

/// Vertical alignment for subscript/superscript text
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum FontVertAlign {
    Baseline,
    Subscript,
    Superscript,
}

impl FontVertAlign {
    pub fn as_str(&self) -> &'static str {
        match self {
            FontVertAlign::Baseline => "baseline",
            FontVertAlign::Subscript => "subscript",
            FontVertAlign::Superscript => "superscript",
        }
    }
}

/// Builder for creating cell styles.
#[derive(Debug, Clone, Default)]
pub struct StyleBuilder {
    // Font properties
    pub font_name: Option<String>,
    pub font_size: Option<f64>,
    pub font_color: Option<String>,
    pub bold: bool,
    pub italic: bool,
    pub underline: bool,
    pub strikethrough: bool,
    pub vert_align: Option<FontVertAlign>,

    // Fill properties
    pub bg_color: Option<String>,
    pub pattern_type: Option<String>,

    // Border properties
    pub border_top: Option<BorderSide>,
    pub border_right: Option<BorderSide>,
    pub border_bottom: Option<BorderSide>,
    pub border_left: Option<BorderSide>,

    // Alignment properties
    pub align_horizontal: Option<String>,
    pub align_vertical: Option<String>,
    pub wrap_text: bool,
    pub shrink_to_fit: bool,
    pub indent: Option<u32>,
    pub rotation: Option<i32>,

    // Number format
    pub number_format: Option<String>,
}

/// A border side definition.
#[derive(Debug, Clone, PartialEq)]
pub struct BorderSide {
    pub style: String,
    pub color: Option<String>,
}

impl BorderSide {
    /// Create a new border side with the given style.
    #[must_use]
    pub fn new(style: &str) -> Self {
        Self {
            style: style.to_string(),
            color: None,
        }
    }

    /// Set the border color.
    #[must_use]
    pub fn color(mut self, color: &str) -> Self {
        self.color = Some(color.to_string());
        self
    }
}

impl StyleBuilder {
    /// Create a new empty style builder.
    #[must_use]
    pub fn new() -> Self {
        Self::default()
    }

    // Font methods

    /// Set the font family name.
    #[must_use]
    pub fn font_name(mut self, name: &str) -> Self {
        self.font_name = Some(name.to_string());
        self
    }

    /// Set the font size in points.
    #[must_use]
    pub fn font_size(mut self, size: f64) -> Self {
        self.font_size = Some(size);
        self
    }

    /// Set the font color as #RRGGBB or AARRGGBB.
    #[must_use]
    pub fn font_color(mut self, color: &str) -> Self {
        self.font_color = Some(normalize_color(color));
        self
    }

    /// Make the font bold.
    #[must_use]
    pub fn bold(mut self) -> Self {
        self.bold = true;
        self
    }

    /// Make the font italic.
    #[must_use]
    pub fn italic(mut self) -> Self {
        self.italic = true;
        self
    }

    /// Add underline to the font.
    #[must_use]
    pub fn underline(mut self) -> Self {
        self.underline = true;
        self
    }

    /// Add strikethrough to the font.
    #[must_use]
    pub fn strikethrough(mut self) -> Self {
        self.strikethrough = true;
        self
    }

    /// Make text subscript.
    #[must_use]
    pub fn subscript(mut self) -> Self {
        self.vert_align = Some(FontVertAlign::Subscript);
        self
    }

    /// Make text superscript.
    #[must_use]
    pub fn superscript(mut self) -> Self {
        self.vert_align = Some(FontVertAlign::Superscript);
        self
    }

    /// Set vertical alignment (baseline, subscript, superscript).
    #[must_use]
    pub fn vert_align(mut self, align: FontVertAlign) -> Self {
        self.vert_align = Some(align);
        self
    }

    // Fill methods

    /// Set the background fill color (solid fill).
    #[must_use]
    pub fn bg_color(mut self, color: &str) -> Self {
        self.bg_color = Some(normalize_color(color));
        self.pattern_type = Some("solid".to_string());
        self
    }

    /// Set the fill pattern type.
    #[must_use]
    pub fn pattern(mut self, pattern_type: &str) -> Self {
        self.pattern_type = Some(pattern_type.to_string());
        self
    }

    // Border methods

    /// Set all borders to the same style.
    #[must_use]
    pub fn border_all(mut self, style: &str, color: Option<&str>) -> Self {
        let side = BorderSide {
            style: style.to_string(),
            color: color.map(normalize_color),
        };
        self.border_top = Some(side.clone());
        self.border_right = Some(side.clone());
        self.border_bottom = Some(side.clone());
        self.border_left = Some(side);
        self
    }

    /// Set the top border.
    #[must_use]
    pub fn border_top(mut self, side: BorderSide) -> Self {
        self.border_top = Some(side);
        self
    }

    /// Set the right border.
    #[must_use]
    pub fn border_right(mut self, side: BorderSide) -> Self {
        self.border_right = Some(side);
        self
    }

    /// Set the bottom border.
    #[must_use]
    pub fn border_bottom(mut self, side: BorderSide) -> Self {
        self.border_bottom = Some(side);
        self
    }

    /// Set the left border.
    #[must_use]
    pub fn border_left(mut self, side: BorderSide) -> Self {
        self.border_left = Some(side);
        self
    }

    // Alignment methods

    /// Set horizontal alignment (left, center, right, justify, fill, general).
    #[must_use]
    pub fn align_horizontal(mut self, align: &str) -> Self {
        self.align_horizontal = Some(align.to_string());
        self
    }

    /// Set vertical alignment (top, center, bottom).
    #[must_use]
    pub fn align_vertical(mut self, align: &str) -> Self {
        self.align_vertical = Some(align.to_string());
        self
    }

    /// Enable text wrapping.
    #[must_use]
    pub fn wrap_text(mut self) -> Self {
        self.wrap_text = true;
        self
    }

    /// Enable shrink to fit.
    #[must_use]
    pub fn shrink_to_fit(mut self) -> Self {
        self.shrink_to_fit = true;
        self
    }

    /// Set text indent level.
    #[must_use]
    pub fn indent(mut self, level: u32) -> Self {
        self.indent = Some(level);
        self
    }

    /// Set text rotation in degrees (-90 to 90, or 255 for vertical).
    #[must_use]
    pub fn rotation(mut self, degrees: i32) -> Self {
        self.rotation = Some(degrees);
        self
    }

    // Number format methods

    /// Set a custom number format code.
    #[must_use]
    pub fn number_format(mut self, format: &str) -> Self {
        self.number_format = Some(format.to_string());
        self
    }

    /// Build the style (returns self for use in cell creation).
    #[must_use]
    pub fn build(self) -> Self {
        self
    }
}

// ============================================================================
// Cell Value
// ============================================================================

/// Represents a cell value that can be added to a sheet.
#[derive(Debug, Clone)]
pub enum CellValue {
    /// A string value.
    String(String),
    /// A numeric value.
    Number(f64),
    /// A boolean value.
    Boolean(bool),
    /// An error value (e.g., "#DIV/0!").
    Error(String),
    /// An inline string (not shared).
    InlineString(String),
    /// An empty cell (style only).
    Empty,
}

impl From<&str> for CellValue {
    fn from(s: &str) -> Self {
        CellValue::String(s.to_string())
    }
}

impl From<String> for CellValue {
    fn from(s: String) -> Self {
        CellValue::String(s)
    }
}

impl From<f64> for CellValue {
    fn from(n: f64) -> Self {
        CellValue::Number(n)
    }
}

impl From<i32> for CellValue {
    fn from(n: i32) -> Self {
        CellValue::Number(f64::from(n))
    }
}

impl From<bool> for CellValue {
    fn from(b: bool) -> Self {
        CellValue::Boolean(b)
    }
}

// ============================================================================
// Sheet Builder
// ============================================================================

/// A cell in the sheet.
#[derive(Debug, Clone)]
pub struct CellEntry {
    pub cell_ref: String,
    pub value: CellValue,
    pub style: Option<StyleBuilder>,
}

/// A merge range.
#[derive(Debug, Clone)]
pub struct MergeEntry {
    pub range: String, // e.g., "A1:B2"
}

/// A column width definition.
#[derive(Debug, Clone)]
pub struct ColumnWidth {
    pub min: u32,
    pub max: u32,
    pub width: f64,
    pub hidden: bool,
}

/// A row height definition.
#[derive(Debug, Clone)]
pub struct RowHeight {
    pub row: u32,
    pub height: f64,
    pub hidden: bool,
}

/// Builder for a single worksheet.
#[derive(Debug, Clone, Default)]
pub struct SheetBuilder {
    pub name: String,
    pub cells: Vec<CellEntry>,
    pub merges: Vec<MergeEntry>,
    pub col_widths: Vec<ColumnWidth>,
    pub row_heights: Vec<RowHeight>,
    pub frozen_rows: u32,
    pub frozen_cols: u32,
}

impl SheetBuilder {
    /// Create a new sheet builder with the given name.
    #[must_use]
    pub fn new(name: &str) -> Self {
        Self {
            name: name.to_string(),
            cells: Vec::new(),
            merges: Vec::new(),
            col_widths: Vec::new(),
            row_heights: Vec::new(),
            frozen_rows: 0,
            frozen_cols: 0,
        }
    }

    /// Add a cell with a value and optional style.
    #[must_use]
    pub fn cell<V: Into<CellValue>>(
        mut self,
        cell_ref: &str,
        value: V,
        style: Option<StyleBuilder>,
    ) -> Self {
        self.cells.push(CellEntry {
            cell_ref: cell_ref.to_string(),
            value: value.into(),
            style,
        });
        self
    }

    /// Add an empty cell with only a style.
    #[must_use]
    pub fn styled_cell(mut self, cell_ref: &str, style: StyleBuilder) -> Self {
        self.cells.push(CellEntry {
            cell_ref: cell_ref.to_string(),
            value: CellValue::Empty,
            style: Some(style),
        });
        self
    }

    /// Add a merge range (e.g., "A1:B2").
    #[must_use]
    pub fn merge(mut self, range: &str) -> Self {
        self.merges.push(MergeEntry {
            range: range.to_string(),
        });
        self
    }

    /// Set column width for a range of columns.
    #[must_use]
    pub fn col_width(mut self, min: u32, max: u32, width: f64) -> Self {
        self.col_widths.push(ColumnWidth {
            min,
            max,
            width,
            hidden: false,
        });
        self
    }

    /// Hide columns in a range.
    #[must_use]
    pub fn hide_cols(mut self, min: u32, max: u32) -> Self {
        self.col_widths.push(ColumnWidth {
            min,
            max,
            width: 8.43,
            hidden: true,
        });
        self
    }

    /// Set row height.
    #[must_use]
    pub fn row_height(mut self, row: u32, height: f64) -> Self {
        self.row_heights.push(RowHeight {
            row,
            height,
            hidden: false,
        });
        self
    }

    /// Hide a row.
    #[must_use]
    pub fn hide_row(mut self, row: u32) -> Self {
        self.row_heights.push(RowHeight {
            row,
            height: 15.0,
            hidden: true,
        });
        self
    }

    /// Set frozen panes (freeze rows and/or columns).
    #[must_use]
    pub fn freeze_panes(mut self, rows: u32, cols: u32) -> Self {
        self.frozen_rows = rows;
        self.frozen_cols = cols;
        self
    }
}

// ============================================================================
// XLSX Builder
// ============================================================================

/// Builder for creating complete XLSX files.
#[derive(Debug, Default)]
pub struct XlsxBuilder {
    sheets: Vec<SheetBuilder>,
    theme_colors: Option<Vec<String>>,
}

impl XlsxBuilder {
    /// Create a new XLSX builder.
    #[must_use]
    pub fn new() -> Self {
        Self::default()
    }

    /// Add a new sheet with the given name.
    #[must_use]
    pub fn sheet(mut self, sheet: SheetBuilder) -> Self {
        self.sheets.push(sheet);
        self
    }

    /// Add a simple sheet by name (returns a builder for chaining).
    #[must_use]
    pub fn add_sheet(self, name: &str) -> XlsxSheetAdder {
        XlsxSheetAdder {
            builder: self,
            sheet: SheetBuilder::new(name),
        }
    }

    /// Set custom theme colors.
    #[must_use]
    pub fn theme_colors(mut self, colors: Vec<String>) -> Self {
        self.theme_colors = Some(colors);
        self
    }

    /// Build the XLSX file as bytes.
    #[must_use]
    pub fn build(self) -> Vec<u8> {
        let cursor = Cursor::new(Vec::new());
        let mut zip = ZipWriter::new(cursor);
        let options =
            FileOptions::<()>::default().compression_method(zip::CompressionMethod::Deflated);

        // Collect all unique styles and shared strings
        let mut styles_collector = StylesCollector::new();
        let mut shared_strings: Vec<String> = Vec::new();

        for sheet in &self.sheets {
            for cell in &sheet.cells {
                // Collect style
                if let Some(ref style) = cell.style {
                    styles_collector.add_style(style);
                }
                // Collect shared strings
                if let CellValue::String(ref s) = cell.value {
                    if !shared_strings.contains(s) {
                        shared_strings.push(s.clone());
                    }
                }
            }
        }

        let has_shared_strings = !shared_strings.is_empty();

        // Write [Content_Types].xml
        let _ = zip.start_file("[Content_Types].xml", options);
        let _ =
            zip.write_all(generate_content_types(self.sheets.len(), has_shared_strings).as_bytes());

        // Write _rels/.rels
        let _ = zip.start_file("_rels/.rels", options);
        let _ = zip.write_all(generate_rels().as_bytes());

        // Write xl/_rels/workbook.xml.rels
        let _ = zip.start_file("xl/_rels/workbook.xml.rels", options);
        let _ =
            zip.write_all(generate_workbook_rels(self.sheets.len(), has_shared_strings).as_bytes());

        // Write xl/workbook.xml
        let _ = zip.start_file("xl/workbook.xml", options);
        let _ = zip.write_all(generate_workbook(&self.sheets).as_bytes());

        // Write xl/styles.xml
        let _ = zip.start_file("xl/styles.xml", options);
        let _ = zip.write_all(styles_collector.generate_styles_xml().as_bytes());

        // Write xl/sharedStrings.xml if we have any
        if !shared_strings.is_empty() {
            let _ = zip.start_file("xl/sharedStrings.xml", options);
            let _ = zip.write_all(generate_shared_strings(&shared_strings).as_bytes());
        }

        // Write xl/theme/theme1.xml
        let _ = zip.start_file("xl/theme/theme1.xml", options);
        let _ = zip.write_all(generate_theme(self.theme_colors.as_deref()).as_bytes());

        // Write each sheet
        for (i, sheet) in self.sheets.iter().enumerate() {
            let path = format!("xl/worksheets/sheet{}.xml", i + 1);
            let _ = zip.start_file(&path, options);
            let _ = zip.write_all(
                generate_sheet_xml(sheet, &shared_strings, &styles_collector).as_bytes(),
            );
        }

        let cursor = zip.finish().expect("Failed to finish ZIP");
        cursor.into_inner()
    }
}

/// Helper for fluent sheet building within `XlsxBuilder`.
pub struct XlsxSheetAdder {
    builder: XlsxBuilder,
    sheet: SheetBuilder,
}

impl XlsxSheetAdder {
    /// Add a cell to the current sheet.
    #[must_use]
    pub fn add_cell<V: Into<CellValue>>(
        mut self,
        cell_ref: &str,
        value: V,
        style: Option<StyleBuilder>,
    ) -> Self {
        self.sheet = self.sheet.cell(cell_ref, value, style);
        self
    }

    /// Add a merge range to the current sheet.
    #[must_use]
    pub fn add_merge(mut self, range: &str) -> Self {
        self.sheet = self.sheet.merge(range);
        self
    }

    /// Finish the current sheet and return the builder.
    #[must_use]
    pub fn done(mut self) -> XlsxBuilder {
        self.builder.sheets.push(self.sheet);
        self.builder
    }

    /// Build the XLSX directly (finishes the current sheet automatically).
    #[must_use]
    pub fn build(self) -> Vec<u8> {
        self.done().build()
    }
}

// ============================================================================
// Styles Collector
// ============================================================================

/// Collects and deduplicates styles for the XLSX file.
#[derive(Debug, Default)]
struct StylesCollector {
    fonts: Vec<FontDef>,
    fills: Vec<FillDef>,
    borders: Vec<BorderDef>,
    num_fmts: Vec<(u32, String)>,
    cell_xfs: Vec<CellXfDef>,
    style_map: Vec<(StyleBuilder, u32)>, // Maps style to xf index
}

#[derive(Debug, Clone, PartialEq)]
struct FontDef {
    name: Option<String>,
    size: Option<f64>,
    color: Option<String>,
    bold: bool,
    italic: bool,
    underline: bool,
    strikethrough: bool,
    vert_align: Option<FontVertAlign>,
}

#[derive(Debug, Clone, PartialEq)]
struct FillDef {
    pattern_type: String,
    fg_color: Option<String>,
}

#[derive(Debug, Clone, PartialEq)]
struct BorderDef {
    top: Option<(String, Option<String>)>,
    right: Option<(String, Option<String>)>,
    bottom: Option<(String, Option<String>)>,
    left: Option<(String, Option<String>)>,
}

#[derive(Debug, Clone)]
struct CellXfDef {
    font_id: u32,
    fill_id: u32,
    border_id: u32,
    num_fmt_id: Option<u32>,
    alignment: Option<AlignmentDef>,
}

#[derive(Debug, Clone)]
struct AlignmentDef {
    horizontal: Option<String>,
    vertical: Option<String>,
    wrap_text: bool,
    shrink_to_fit: bool,
    indent: Option<u32>,
    rotation: Option<i32>,
}

impl StylesCollector {
    fn new() -> Self {
        let mut collector = Self::default();

        // Add default font (required)
        collector.fonts.push(FontDef {
            name: Some("Calibri".to_string()),
            size: Some(11.0),
            color: None,
            bold: false,
            italic: false,
            underline: false,
            strikethrough: false,
            vert_align: None,
        });

        // Add required fills (none and gray125)
        collector.fills.push(FillDef {
            pattern_type: "none".to_string(),
            fg_color: None,
        });
        collector.fills.push(FillDef {
            pattern_type: "gray125".to_string(),
            fg_color: None,
        });

        // Add default border (none)
        collector.borders.push(BorderDef {
            top: None,
            right: None,
            bottom: None,
            left: None,
        });

        // Add default cell format
        collector.cell_xfs.push(CellXfDef {
            font_id: 0,
            fill_id: 0,
            border_id: 0,
            num_fmt_id: None,
            alignment: None,
        });

        collector
    }

    fn add_style(&mut self, style: &StyleBuilder) -> u32 {
        // Check if we already have this style
        for (existing, idx) in &self.style_map {
            if styles_equal(existing, style) {
                return *idx;
            }
        }

        // Create new style components
        let font_id = self.add_font(style);
        let fill_id = self.add_fill(style);
        let border_id = self.add_border(style);
        let num_fmt_id = self.add_num_fmt(style);
        let alignment = self.make_alignment(style);

        let xf = CellXfDef {
            font_id,
            fill_id,
            border_id,
            num_fmt_id,
            alignment,
        };

        let idx = self.cell_xfs.len() as u32;
        self.cell_xfs.push(xf);
        self.style_map.push((style.clone(), idx));
        idx
    }

    fn get_style_index(&self, style: &StyleBuilder) -> u32 {
        for (existing, idx) in &self.style_map {
            if styles_equal(existing, style) {
                return *idx;
            }
        }
        0 // Default style
    }

    fn add_font(&mut self, style: &StyleBuilder) -> u32 {
        let font = FontDef {
            name: style
                .font_name
                .clone()
                .or_else(|| Some("Calibri".to_string())),
            size: style.font_size.or(Some(11.0)),
            color: style.font_color.clone(),
            bold: style.bold,
            italic: style.italic,
            underline: style.underline,
            strikethrough: style.strikethrough,
            vert_align: style.vert_align,
        };

        // Check if font already exists
        for (i, f) in self.fonts.iter().enumerate() {
            if f == &font {
                return i as u32;
            }
        }

        let idx = self.fonts.len() as u32;
        self.fonts.push(font);
        idx
    }

    fn add_fill(&mut self, style: &StyleBuilder) -> u32 {
        if style.bg_color.is_none() && style.pattern_type.is_none() {
            return 0; // Default no fill
        }

        let fill = FillDef {
            pattern_type: style
                .pattern_type
                .clone()
                .unwrap_or_else(|| "solid".to_string()),
            fg_color: style.bg_color.clone(),
        };

        // Check if fill already exists
        for (i, f) in self.fills.iter().enumerate() {
            if f == &fill {
                return i as u32;
            }
        }

        let idx = self.fills.len() as u32;
        self.fills.push(fill);
        idx
    }

    fn add_border(&mut self, style: &StyleBuilder) -> u32 {
        if style.border_top.is_none()
            && style.border_right.is_none()
            && style.border_bottom.is_none()
            && style.border_left.is_none()
        {
            return 0; // Default no border
        }

        let border = BorderDef {
            top: style
                .border_top
                .as_ref()
                .map(|s| (s.style.clone(), s.color.clone())),
            right: style
                .border_right
                .as_ref()
                .map(|s| (s.style.clone(), s.color.clone())),
            bottom: style
                .border_bottom
                .as_ref()
                .map(|s| (s.style.clone(), s.color.clone())),
            left: style
                .border_left
                .as_ref()
                .map(|s| (s.style.clone(), s.color.clone())),
        };

        // Check if border already exists
        for (i, b) in self.borders.iter().enumerate() {
            if b == &border {
                return i as u32;
            }
        }

        let idx = self.borders.len() as u32;
        self.borders.push(border);
        idx
    }

    fn add_num_fmt(&mut self, style: &StyleBuilder) -> Option<u32> {
        let format = style.number_format.as_ref()?;

        // Check if it's a built-in format
        if let Some(id) = get_builtin_format_id(format) {
            return Some(id);
        }

        // Check if we already have this custom format
        for (id, code) in &self.num_fmts {
            if code == format {
                return Some(*id);
            }
        }

        // Add new custom format (start at 164)
        let id = 164 + self.num_fmts.len() as u32;
        self.num_fmts.push((id, format.clone()));
        Some(id)
    }

    fn make_alignment(&self, style: &StyleBuilder) -> Option<AlignmentDef> {
        if style.align_horizontal.is_none()
            && style.align_vertical.is_none()
            && !style.wrap_text
            && !style.shrink_to_fit
            && style.indent.is_none()
            && style.rotation.is_none()
        {
            return None;
        }

        Some(AlignmentDef {
            horizontal: style.align_horizontal.clone(),
            vertical: style.align_vertical.clone(),
            wrap_text: style.wrap_text,
            shrink_to_fit: style.shrink_to_fit,
            indent: style.indent,
            rotation: style.rotation,
        })
    }

    fn generate_styles_xml(&self) -> String {
        let mut xml = String::new();
        xml.push_str(r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#);
        xml.push_str(
            r#"<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">"#,
        );

        // Number formats
        if !self.num_fmts.is_empty() {
            xml.push_str(&format!(r#"<numFmts count="{}">"#, self.num_fmts.len()));
            for (id, code) in &self.num_fmts {
                xml.push_str(&format!(
                    r#"<numFmt numFmtId="{}" formatCode="{}"/>"#,
                    id,
                    escape_xml(code)
                ));
            }
            xml.push_str("</numFmts>");
        }

        // Fonts
        xml.push_str(&format!(r#"<fonts count="{}">"#, self.fonts.len()));
        for font in &self.fonts {
            xml.push_str("<font>");
            if font.bold {
                xml.push_str("<b/>");
            }
            if font.italic {
                xml.push_str("<i/>");
            }
            if font.underline {
                xml.push_str("<u/>");
            }
            if font.strikethrough {
                xml.push_str("<strike/>");
            }
            if let Some(ref vert_align) = font.vert_align {
                xml.push_str(&format!(r#"<vertAlign val="{}"/>"#, vert_align.as_str()));
            }
            if let Some(size) = font.size {
                xml.push_str(&format!(r#"<sz val="{}"/>"#, size));
            }
            if let Some(ref color) = font.color {
                xml.push_str(&format!(r#"<color rgb="{}"/>"#, color));
            }
            if let Some(ref name) = font.name {
                xml.push_str(&format!(r#"<name val="{}"/>"#, escape_xml(name)));
            }
            xml.push_str("</font>");
        }
        xml.push_str("</fonts>");

        // Fills
        xml.push_str(&format!(r#"<fills count="{}">"#, self.fills.len()));
        for fill in &self.fills {
            xml.push_str("<fill>");
            xml.push_str(&format!(
                r#"<patternFill patternType="{}">"#,
                fill.pattern_type
            ));
            if let Some(ref color) = fill.fg_color {
                xml.push_str(&format!(r#"<fgColor rgb="{}"/>"#, color));
            }
            xml.push_str("</patternFill>");
            xml.push_str("</fill>");
        }
        xml.push_str("</fills>");

        // Borders
        xml.push_str(&format!(r#"<borders count="{}">"#, self.borders.len()));
        for border in &self.borders {
            xml.push_str("<border>");
            xml.push_str(&format_border_side("left", &border.left));
            xml.push_str(&format_border_side("right", &border.right));
            xml.push_str(&format_border_side("top", &border.top));
            xml.push_str(&format_border_side("bottom", &border.bottom));
            xml.push_str("<diagonal/>");
            xml.push_str("</border>");
        }
        xml.push_str("</borders>");

        // Cell style formats (cellStyleXfs) - required
        xml.push_str(r#"<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>"#);

        // Cell formats (cellXfs)
        xml.push_str(&format!(r#"<cellXfs count="{}">"#, self.cell_xfs.len()));
        for xf in &self.cell_xfs {
            let mut attrs = format!(
                r#"fontId="{}" fillId="{}" borderId="{}""#,
                xf.font_id, xf.fill_id, xf.border_id
            );

            if let Some(num_fmt_id) = xf.num_fmt_id {
                attrs.push_str(&format!(
                    r#" numFmtId="{}" applyNumberFormat="1""#,
                    num_fmt_id
                ));
            }

            if xf.font_id > 0 {
                attrs.push_str(r#" applyFont="1""#);
            }
            if xf.fill_id > 0 {
                attrs.push_str(r#" applyFill="1""#);
            }
            if xf.border_id > 0 {
                attrs.push_str(r#" applyBorder="1""#);
            }
            if xf.alignment.is_some() {
                attrs.push_str(r#" applyAlignment="1""#);
            }

            if let Some(ref align) = xf.alignment {
                xml.push_str(&format!("<xf {}>", attrs));
                let mut align_attrs = String::new();
                if let Some(ref h) = align.horizontal {
                    align_attrs.push_str(&format!(r#" horizontal="{}""#, h));
                }
                if let Some(ref v) = align.vertical {
                    align_attrs.push_str(&format!(r#" vertical="{}""#, v));
                }
                if align.wrap_text {
                    align_attrs.push_str(r#" wrapText="1""#);
                }
                if align.shrink_to_fit {
                    align_attrs.push_str(r#" shrinkToFit="1""#);
                }
                if let Some(indent) = align.indent {
                    align_attrs.push_str(&format!(r#" indent="{}""#, indent));
                }
                if let Some(rotation) = align.rotation {
                    align_attrs.push_str(&format!(r#" textRotation="{}""#, rotation));
                }
                xml.push_str(&format!("<alignment{}/>", align_attrs));
                xml.push_str("</xf>");
            } else {
                xml.push_str(&format!("<xf {}/>", attrs));
            }
        }
        xml.push_str("</cellXfs>");

        // Cell styles - required
        xml.push_str(r#"<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>"#);

        xml.push_str("</styleSheet>");
        xml
    }
}

// ============================================================================
// Helper Functions
// ============================================================================

/// Normalize color to ARGB format (without #).
fn normalize_color(color: &str) -> String {
    let color = color.trim_start_matches('#');
    if color.len() == 6 {
        format!("FF{}", color.to_uppercase())
    } else if color.len() == 8 {
        color.to_uppercase()
    } else {
        format!("FF{}", color.to_uppercase())
    }
}

/// Check if two styles are equal.
fn styles_equal(a: &StyleBuilder, b: &StyleBuilder) -> bool {
    a.font_name == b.font_name
        && a.font_size == b.font_size
        && a.font_color == b.font_color
        && a.bold == b.bold
        && a.italic == b.italic
        && a.underline == b.underline
        && a.strikethrough == b.strikethrough
        && a.vert_align == b.vert_align
        && a.bg_color == b.bg_color
        && a.pattern_type == b.pattern_type
        && a.border_top == b.border_top
        && a.border_right == b.border_right
        && a.border_bottom == b.border_bottom
        && a.border_left == b.border_left
        && a.align_horizontal == b.align_horizontal
        && a.align_vertical == b.align_vertical
        && a.wrap_text == b.wrap_text
        && a.shrink_to_fit == b.shrink_to_fit
        && a.indent == b.indent
        && a.rotation == b.rotation
        && a.number_format == b.number_format
}

/// Get built-in number format ID.
fn get_builtin_format_id(format: &str) -> Option<u32> {
    match format {
        "General" => Some(0),
        "0" => Some(1),
        "0.00" => Some(2),
        "#,##0" => Some(3),
        "#,##0.00" => Some(4),
        "0%" => Some(9),
        "0.00%" => Some(10),
        "0.00E+00" => Some(11),
        "# ?/?" => Some(12),
        "# ??/??" => Some(13),
        "mm-dd-yy" | "m/d/yy" => Some(14),
        "d-mmm-yy" => Some(15),
        "d-mmm" => Some(16),
        "mmm-yy" => Some(17),
        "h:mm AM/PM" => Some(18),
        "h:mm:ss AM/PM" => Some(19),
        "h:mm" => Some(20),
        "h:mm:ss" => Some(21),
        "m/d/yy h:mm" => Some(22),
        "@" => Some(49),
        _ => None,
    }
}

/// Escape XML special characters.
fn escape_xml(s: &str) -> String {
    s.replace('&', "&amp;")
        .replace('<', "&lt;")
        .replace('>', "&gt;")
        .replace('"', "&quot;")
        .replace('\'', "&apos;")
}

/// Format a border side element.
fn format_border_side(name: &str, side: &Option<(String, Option<String>)>) -> String {
    match side {
        Some((style, color)) => {
            let mut xml = format!(r#"<{} style="{}">"#, name, style);
            if let Some(c) = color {
                xml.push_str(&format!(r#"<color rgb="{}"/>"#, c));
            }
            xml.push_str(&format!("</{}>", name));
            xml
        }
        None => format!("<{}/>", name),
    }
}

/// Generate [Content_Types].xml
fn generate_content_types(sheet_count: usize, has_shared_strings: bool) -> String {
    let mut xml = String::new();
    xml.push_str(r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#);
    xml.push_str(r#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">"#);
    xml.push_str(r#"<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>"#);
    xml.push_str(r#"<Default Extension="xml" ContentType="application/xml"/>"#);
    xml.push_str(r#"<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>"#);
    xml.push_str(r#"<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>"#);
    if has_shared_strings {
        xml.push_str(r#"<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>"#);
    }
    xml.push_str(r#"<Override PartName="/xl/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>"#);

    for i in 1..=sheet_count {
        xml.push_str(&format!(
            r#"<Override PartName="/xl/worksheets/sheet{}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>"#,
            i
        ));
    }

    xml.push_str("</Types>");
    xml
}

/// Generate _rels/.rels
fn generate_rels() -> String {
    let mut xml = String::new();
    xml.push_str(r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#);
    xml.push_str(
        r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">"#,
    );
    xml.push_str(r#"<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>"#);
    xml.push_str("</Relationships>");
    xml
}

/// Generate xl/_rels/workbook.xml.rels
fn generate_workbook_rels(sheet_count: usize, has_shared_strings: bool) -> String {
    let mut xml = String::new();
    xml.push_str(r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#);
    xml.push_str(
        r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">"#,
    );

    let mut rid = 1;

    // Sheets
    for i in 1..=sheet_count {
        xml.push_str(&format!(
            r#"<Relationship Id="rId{}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet{}.xml"/>"#,
            rid, i
        ));
        rid += 1;
    }

    // Styles
    xml.push_str(&format!(
        r#"<Relationship Id="rId{}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>"#,
        rid
    ));
    rid += 1;

    // Shared strings (only if there are any)
    if has_shared_strings {
        xml.push_str(&format!(
            r#"<Relationship Id="rId{}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>"#,
            rid
        ));
        rid += 1;
    }

    // Theme
    xml.push_str(&format!(
        r#"<Relationship Id="rId{}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>"#,
        rid
    ));

    xml.push_str("</Relationships>");
    xml
}

/// Generate xl/workbook.xml
fn generate_workbook(sheets: &[SheetBuilder]) -> String {
    let mut xml = String::new();
    xml.push_str(r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#);
    xml.push_str(r#"<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">"#);
    xml.push_str("<sheets>");

    for (i, sheet) in sheets.iter().enumerate() {
        xml.push_str(&format!(
            r#"<sheet name="{}" sheetId="{}" r:id="rId{}"/>"#,
            escape_xml(&sheet.name),
            i + 1,
            i + 1
        ));
    }

    xml.push_str("</sheets>");
    xml.push_str("</workbook>");
    xml
}

/// Generate xl/sharedStrings.xml
fn generate_shared_strings(strings: &[String]) -> String {
    let mut xml = String::new();
    xml.push_str(r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#);
    xml.push_str(&format!(
        r#"<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="{}" uniqueCount="{}">"#,
        strings.len(),
        strings.len()
    ));

    for s in strings {
        // Add xml:space="preserve" to preserve leading/trailing whitespace
        xml.push_str(&format!(
            r#"<si><t xml:space="preserve">{}</t></si>"#,
            escape_xml(s)
        ));
    }

    xml.push_str("</sst>");
    xml
}

/// Generate xl/theme/theme1.xml
fn generate_theme(colors: Option<&[String]>) -> String {
    let default_colors = [
        "#000000", "#FFFFFF", "#44546A", "#E7E6E6", "#4472C4", "#ED7D31", "#A5A5A5", "#FFC000",
        "#5B9BD5", "#70AD47", "#0563C1", "#954F72",
    ];

    let colors = colors.unwrap_or(&[]);

    let mut xml = String::new();
    xml.push_str(r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#);
    xml.push_str(r#"<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office Theme">"#);
    xml.push_str("<a:themeElements>");
    xml.push_str(r#"<a:clrScheme name="Office">"#);

    let color_names = [
        "dk1", "lt1", "dk2", "lt2", "accent1", "accent2", "accent3", "accent4", "accent5",
        "accent6", "hlink", "folHlink",
    ];

    for (i, name) in color_names.iter().enumerate() {
        let color = if i < colors.len() {
            colors[i].trim_start_matches('#')
        } else if i < default_colors.len() {
            default_colors[i].trim_start_matches('#')
        } else {
            "000000"
        };

        xml.push_str(&format!(
            r#"<a:{}><a:srgbClr val="{}"/></a:{}>"#,
            name, color, name
        ));
    }

    xml.push_str("</a:clrScheme>");
    xml.push_str(r#"<a:fontScheme name="Office"><a:majorFont><a:latin typeface="Calibri Light"/></a:majorFont><a:minorFont><a:latin typeface="Calibri"/></a:minorFont></a:fontScheme>"#);
    xml.push_str(r#"<a:fmtScheme name="Office"><a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:fillStyleLst></a:fmtScheme>"#);
    xml.push_str("</a:themeElements>");
    xml.push_str("</a:theme>");
    xml
}

/// Convert column number (1-indexed) to letter(s).
fn col_num_to_letter(col: u32) -> String {
    let mut result = String::new();
    let mut n = col;
    while n > 0 {
        n -= 1;
        result.insert(0, (b'A' + (n % 26) as u8) as char);
        n /= 26;
    }
    if result.is_empty() {
        result.push('A');
    }
    result
}

/// Generate a sheet XML file
fn generate_sheet_xml(
    sheet: &SheetBuilder,
    shared_strings: &[String],
    styles: &StylesCollector,
) -> String {
    let mut xml = String::new();
    xml.push_str(r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#);
    xml.push_str(
        r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">"#,
    );

    // Sheet views with frozen panes
    if sheet.frozen_rows > 0 || sheet.frozen_cols > 0 {
        let top_left = format!(
            "{}{}",
            col_num_to_letter(sheet.frozen_cols + 1),
            sheet.frozen_rows + 1
        );

        // Determine active pane based on what's frozen
        let active_pane = if sheet.frozen_rows > 0 && sheet.frozen_cols > 0 {
            "bottomRight"
        } else if sheet.frozen_rows > 0 {
            "bottomLeft"
        } else {
            "topRight"
        };

        xml.push_str(r#"<sheetViews><sheetView tabSelected="1" workbookViewId="0">"#);

        // Build pane element
        let mut pane_attrs = String::new();
        if sheet.frozen_cols > 0 {
            pane_attrs.push_str(&format!(r#" xSplit="{}""#, sheet.frozen_cols));
        }
        if sheet.frozen_rows > 0 {
            pane_attrs.push_str(&format!(r#" ySplit="{}""#, sheet.frozen_rows));
        }
        pane_attrs.push_str(&format!(
            r#" topLeftCell="{}" activePane="{}" state="frozen""#,
            top_left, active_pane
        ));

        xml.push_str(&format!(r#"<pane{}/>"#, pane_attrs));
        xml.push_str(&format!(
            r#"<selection pane="{}" activeCell="{}" sqref="{}"/>"#,
            active_pane, top_left, top_left
        ));
        xml.push_str(r#"</sheetView></sheetViews>"#);
    }

    // Column definitions
    if !sheet.col_widths.is_empty() {
        xml.push_str("<cols>");
        for col in &sheet.col_widths {
            let hidden = if col.hidden { r#" hidden="1""# } else { "" };
            xml.push_str(&format!(
                r#"<col min="{}" max="{}" width="{}" customWidth="1"{}/ >"#,
                col.min, col.max, col.width, hidden
            ));
        }
        xml.push_str("</cols>");
    }

    // Sheet data
    xml.push_str("<sheetData>");

    // Group cells by row
    let mut rows: std::collections::BTreeMap<u32, Vec<&CellEntry>> =
        std::collections::BTreeMap::new();
    for cell in &sheet.cells {
        let (_, row) = parse_cell_ref(&cell.cell_ref);
        rows.entry(row).or_default().push(cell);
    }

    // Add row heights
    let row_height_map: std::collections::HashMap<u32, &RowHeight> =
        sheet.row_heights.iter().map(|rh| (rh.row, rh)).collect();

    for (row_num, cells) in rows {
        let mut row_attrs = format!(r#"r="{}""#, row_num);

        if let Some(rh) = row_height_map.get(&row_num) {
            row_attrs.push_str(&format!(r#" ht="{}" customHeight="1""#, rh.height));
            if rh.hidden {
                row_attrs.push_str(r#" hidden="1""#);
            }
        }

        xml.push_str(&format!("<row {}>", row_attrs));

        for cell in cells {
            let mut cell_attrs = format!(r#"r="{}""#, cell.cell_ref);

            // Add style reference
            if let Some(ref style) = cell.style {
                let style_idx = styles.get_style_index(style);
                if style_idx > 0 {
                    cell_attrs.push_str(&format!(r#" s="{}""#, style_idx));
                }
            }

            match &cell.value {
                CellValue::String(s) => {
                    let idx = shared_strings.iter().position(|x| x == s).unwrap_or(0);
                    cell_attrs.push_str(r#" t="s""#);
                    xml.push_str(&format!(r#"<c {}><v>{}</v></c>"#, cell_attrs, idx));
                }
                CellValue::Number(n) => {
                    xml.push_str(&format!(r#"<c {}><v>{}</v></c>"#, cell_attrs, n));
                }
                CellValue::Boolean(b) => {
                    cell_attrs.push_str(r#" t="b""#);
                    let v = if *b { "1" } else { "0" };
                    xml.push_str(&format!(r#"<c {}><v>{}</v></c>"#, cell_attrs, v));
                }
                CellValue::Error(e) => {
                    cell_attrs.push_str(r#" t="e""#);
                    xml.push_str(&format!(
                        r#"<c {}><v>{}</v></c>"#,
                        cell_attrs,
                        escape_xml(e)
                    ));
                }
                CellValue::InlineString(s) => {
                    cell_attrs.push_str(r#" t="inlineStr""#);
                    xml.push_str(&format!(
                        r#"<c {}><is><t>{}</t></is></c>"#,
                        cell_attrs,
                        escape_xml(s)
                    ));
                }
                CellValue::Empty => {
                    xml.push_str(&format!(r#"<c {}/>"#, cell_attrs));
                }
            }
        }

        xml.push_str("</row>");
    }

    xml.push_str("</sheetData>");

    // Merge cells
    if !sheet.merges.is_empty() {
        xml.push_str(&format!(r#"<mergeCells count="{}">"#, sheet.merges.len()));
        for merge in &sheet.merges {
            xml.push_str(&format!(r#"<mergeCell ref="{}"/>"#, merge.range));
        }
        xml.push_str("</mergeCells>");
    }

    xml.push_str("</worksheet>");
    xml
}

/// Parse a cell reference like "A1" into (col, row) as 1-indexed.
fn parse_cell_ref(cell_ref: &str) -> (u32, u32) {
    let mut col: u32 = 0;
    let mut row: u32 = 0;
    let mut in_letters = true;

    for c in cell_ref.chars() {
        if in_letters && c.is_ascii_alphabetic() {
            col = col * 26 + (c.to_ascii_uppercase() as u32 - 'A' as u32 + 1);
        } else {
            in_letters = false;
            if c.is_ascii_digit() {
                row = row * 10 + (c as u32 - '0' as u32);
            }
        }
    }

    (col, row)
}

// ============================================================================
// Convenience Functions
// ============================================================================

/// Create a minimal valid XLSX with a single empty sheet.
#[must_use]
pub fn minimal_xlsx() -> Vec<u8> {
    XlsxBuilder::new().add_sheet("Sheet1").build()
}

/// Create an XLSX with a single cell containing text.
#[must_use]
pub fn xlsx_with_text(text: &str) -> Vec<u8> {
    XlsxBuilder::new()
        .add_sheet("Sheet1")
        .add_cell("A1", text, None)
        .build()
}

/// Create an XLSX with a single cell containing a number.
#[must_use]
pub fn xlsx_with_number(value: f64) -> Vec<u8> {
    XlsxBuilder::new()
        .add_sheet("Sheet1")
        .add_cell("A1", value, None)
        .build()
}

/// Create an XLSX with a styled cell.
#[must_use]
pub fn xlsx_with_styled_cell<V: Into<CellValue>>(value: V, style: StyleBuilder) -> Vec<u8> {
    XlsxBuilder::new()
        .add_sheet("Sheet1")
        .add_cell("A1", value, Some(style))
        .build()
}

// ============================================================================
// All 19 Pattern Fill Types
// ============================================================================

/// All ECMA-376 pattern fill types
pub const ALL_PATTERN_FILLS: &[&str] = &[
    "none",
    "solid",
    "mediumGray",
    "darkGray",
    "lightGray",
    "darkHorizontal",
    "darkVertical",
    "darkDown",
    "darkUp",
    "darkGrid",
    "darkTrellis",
    "lightHorizontal",
    "lightVertical",
    "lightDown",
    "lightUp",
    "lightGrid",
    "lightTrellis",
    "gray125",
    "gray0625",
];

/// All ECMA-376 border styles
pub const ALL_BORDER_STYLES: &[&str] = &[
    "none",
    "thin",
    "medium",
    "dashed",
    "dotted",
    "thick",
    "double",
    "hair",
    "mediumDashed",
    "dashDot",
    "mediumDashDot",
    "dashDotDot",
    "mediumDashDotDot",
    "slantDashDot",
];

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_minimal_xlsx_is_valid_zip() {
        let bytes = minimal_xlsx();
        let cursor = Cursor::new(bytes);
        let archive = zip::ZipArchive::new(cursor);
        assert!(archive.is_ok());
    }

    #[test]
    fn test_style_builder_chaining() {
        let style = StyleBuilder::new()
            .bold()
            .italic()
            .font_size(14.0)
            .font_color("#FF0000")
            .bg_color("#FFFF00")
            .border_all("thin", Some("#000000"))
            .align_horizontal("center")
            .wrap_text()
            .build();

        assert!(style.bold);
        assert!(style.italic);
        assert_eq!(style.font_size, Some(14.0));
        assert_eq!(style.font_color, Some("FFFF0000".to_string()));
        assert_eq!(style.bg_color, Some("FFFFFF00".to_string()));
        assert!(style.border_top.is_some());
        assert_eq!(style.align_horizontal, Some("center".to_string()));
        assert!(style.wrap_text);
    }

    #[test]
    fn test_cell_value_conversions() {
        let _: CellValue = "hello".into();
        let _: CellValue = String::from("world").into();
        let _: CellValue = 42.0.into();
        let _: CellValue = 42.into();
        let _: CellValue = true.into();
    }
}
