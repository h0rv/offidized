//! Lower-level test helpers for creating XLSX files using raw ZIP + XML.
//!
//! These builders produce XLSX bytes at the ZIP level, useful for testing
//! edge cases in the parser (e.g., inline strings, shared strings, missing
//! parts). For most tests, prefer the higher-level `fixtures::XlsxBuilder`.
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

/// Create a minimal XLSX with just the required structure
pub fn create_minimal_xlsx() -> Vec<u8> {
    let mut buffer = Cursor::new(Vec::new());
    {
        let mut zip = ZipWriter::new(&mut buffer);
        let options =
            FileOptions::<()>::default().compression_method(zip::CompressionMethod::Deflated);

        // [Content_Types].xml
        zip.start_file("[Content_Types].xml", options).unwrap();
        zip.write_all(CONTENT_TYPES_XML.as_bytes()).unwrap();

        // _rels/.rels
        zip.start_file("_rels/.rels", options).unwrap();
        zip.write_all(RELS_XML.as_bytes()).unwrap();

        // xl/workbook.xml
        zip.start_file("xl/workbook.xml", options).unwrap();
        zip.write_all(WORKBOOK_XML.as_bytes()).unwrap();

        // xl/_rels/workbook.xml.rels
        zip.start_file("xl/_rels/workbook.xml.rels", options)
            .unwrap();
        zip.write_all(WORKBOOK_RELS_XML.as_bytes()).unwrap();

        // xl/worksheets/sheet1.xml
        zip.start_file("xl/worksheets/sheet1.xml", options).unwrap();
        zip.write_all(EMPTY_SHEET_XML.as_bytes()).unwrap();

        zip.finish().unwrap();
    }
    buffer.into_inner()
}

const CONTENT_TYPES_XML: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
</Types>"#;

const RELS_XML: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>"#;

const WORKBOOK_XML: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Sheet1" sheetId="1" r:id="rId1"/>
  </sheets>
</workbook>"#;

const WORKBOOK_RELS_XML: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>"#;

const EMPTY_SHEET_XML: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData/>
</worksheet>"#;

/// Create a minimal XLSX with custom shared strings and sheet XML.
///
/// Useful for testing rich text, inline strings, and other shared string
/// scenarios where you need full control over the XML content.
pub fn create_xlsx_with_shared_strings_and_sheet(
    shared_strings_xml: &str,
    sheet_xml: &str,
) -> Vec<u8> {
    let mut buffer = Cursor::new(Vec::new());
    {
        let mut zip = ZipWriter::new(&mut buffer);
        let options =
            FileOptions::<()>::default().compression_method(zip::CompressionMethod::Deflated);

        // [Content_Types].xml
        zip.start_file("[Content_Types].xml", options).unwrap();
        zip.write_all(
            br#"<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>
</Types>"#,
        )
        .unwrap();

        // _rels/.rels
        zip.start_file("_rels/.rels", options).unwrap();
        zip.write_all(
            br#"<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>"#,
        )
        .unwrap();

        // xl/_rels/workbook.xml.rels
        zip.start_file("xl/_rels/workbook.xml.rels", options)
            .unwrap();
        zip.write_all(
            br#"<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>
</Relationships>"#,
        )
        .unwrap();

        // xl/workbook.xml
        zip.start_file("xl/workbook.xml", options).unwrap();
        zip.write_all(
            br#"<?xml version="1.0" encoding="UTF-8"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheets>
    <sheet name="Sheet1" sheetId="1" r:id="rId1" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"/>
  </sheets>
</workbook>"#,
        )
        .unwrap();

        // xl/sharedStrings.xml
        zip.start_file("xl/sharedStrings.xml", options).unwrap();
        zip.write_all(shared_strings_xml.as_bytes()).unwrap();

        // xl/worksheets/sheet1.xml
        zip.start_file("xl/worksheets/sheet1.xml", options).unwrap();
        zip.write_all(sheet_xml.as_bytes()).unwrap();

        zip.finish().unwrap();
    }
    buffer.into_inner()
}

/// Cell data for building sheets
#[derive(Clone, Debug)]
pub struct CellData {
    pub cell_ref: String,
    pub value: String,
    pub cell_type: Option<String>,
    pub style_index: Option<u32>,
    pub formula: Option<String>,
}

impl CellData {
    pub fn new(cell_ref: &str, value: &str) -> Self {
        Self {
            cell_ref: cell_ref.to_string(),
            value: value.to_string(),
            cell_type: None,
            style_index: None,
            formula: None,
        }
    }

    pub fn with_type(mut self, cell_type: &str) -> Self {
        self.cell_type = Some(cell_type.to_string());
        self
    }

    pub fn with_style(mut self, style_index: u32) -> Self {
        self.style_index = Some(style_index);
        self
    }

    pub fn with_formula(mut self, formula: &str) -> Self {
        self.formula = Some(formula.to_string());
        self
    }

    /// Create a cell with shared string reference
    pub fn shared_string(cell_ref: &str, sst_index: u32) -> Self {
        Self {
            cell_ref: cell_ref.to_string(),
            value: sst_index.to_string(),
            cell_type: Some("s".to_string()),
            style_index: None,
            formula: None,
        }
    }

    /// Create a numeric cell
    pub fn number(cell_ref: &str, value: f64) -> Self {
        Self {
            cell_ref: cell_ref.to_string(),
            value: value.to_string(),
            cell_type: None,
            style_index: None,
            formula: None,
        }
    }

    /// Create an inline string cell
    pub fn inline_string(cell_ref: &str, value: &str) -> Self {
        Self {
            cell_ref: cell_ref.to_string(),
            value: value.to_string(),
            cell_type: Some("inlineStr".to_string()),
            style_index: None,
            formula: None,
        }
    }

    /// Create a boolean cell
    pub fn boolean(cell_ref: &str, value: bool) -> Self {
        Self {
            cell_ref: cell_ref.to_string(),
            value: if value {
                "1".to_string()
            } else {
                "0".to_string()
            },
            cell_type: Some("b".to_string()),
            style_index: None,
            formula: None,
        }
    }
}

/// Builder for creating worksheet XML
#[derive(Clone, Debug, Default)]
pub struct SheetBuilder {
    pub name: String,
    cells: Vec<CellData>,
    merges: Vec<String>,
    hyperlinks: Vec<(String, String, String)>, // (cell_ref, r:id, display)
    column_widths: Vec<(u32, u32, f64)>,       // (min, max, width)
    row_heights: Vec<(u32, f64)>,              // (row, height)
    frozen_rows: u32,
    frozen_cols: u32,
    dimension: Option<String>,
}

impl SheetBuilder {
    pub fn new(name: &str) -> Self {
        Self {
            name: name.to_string(),
            ..Default::default()
        }
    }

    /// Add a cell to the sheet
    pub fn add_cell(&mut self, cell: CellData) -> &mut Self {
        self.cells.push(cell);
        self
    }

    /// Add a simple value cell
    pub fn add_value(&mut self, cell_ref: &str, value: &str) -> &mut Self {
        self.cells.push(CellData::new(cell_ref, value));
        self
    }

    /// Add a numeric cell
    pub fn add_number(&mut self, cell_ref: &str, value: f64) -> &mut Self {
        self.cells.push(CellData::number(cell_ref, value));
        self
    }

    /// Add a shared string cell
    pub fn add_shared_string(&mut self, cell_ref: &str, sst_index: u32) -> &mut Self {
        self.cells
            .push(CellData::shared_string(cell_ref, sst_index));
        self
    }

    /// Add a merge range (e.g., "A1:C1")
    pub fn add_merge(&mut self, range: &str) -> &mut Self {
        self.merges.push(range.to_string());
        self
    }

    /// Add a hyperlink
    pub fn add_hyperlink(&mut self, cell_ref: &str, rel_id: &str, display: &str) -> &mut Self {
        self.hyperlinks.push((
            cell_ref.to_string(),
            rel_id.to_string(),
            display.to_string(),
        ));
        self
    }

    /// Set column width
    pub fn set_column_width(&mut self, min: u32, max: u32, width: f64) -> &mut Self {
        self.column_widths.push((min, max, width));
        self
    }

    /// Set row height
    pub fn set_row_height(&mut self, row: u32, height: f64) -> &mut Self {
        self.row_heights.push((row, height));
        self
    }

    /// Set frozen panes
    pub fn freeze_panes(&mut self, rows: u32, cols: u32) -> &mut Self {
        self.frozen_rows = rows;
        self.frozen_cols = cols;
        self
    }

    /// Set explicit dimension
    pub fn set_dimension(&mut self, dimension: &str) -> &mut Self {
        self.dimension = Some(dimension.to_string());
        self
    }

    /// Build the worksheet XML
    pub fn build_xml(&self) -> String {
        let mut xml = String::from(
            r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">"#,
        );

        // Dimension
        if let Some(ref dim) = self.dimension {
            xml.push_str(&format!(r#"<dimension ref="{}"/>"#, dim));
        }

        // Sheet views with freeze panes
        if self.frozen_rows > 0 || self.frozen_cols > 0 {
            let top_left = format!(
                "{}{}",
                col_to_letter(self.frozen_cols + 1),
                self.frozen_rows + 1
            );
            xml.push_str(r#"<sheetViews><sheetView tabSelected="1" workbookViewId="0">"#);
            xml.push_str(&format!(
                r#"<pane xSplit="{}" ySplit="{}" topLeftCell="{}" activePane="bottomRight" state="frozen"/>"#,
                self.frozen_cols, self.frozen_rows, top_left
            ));
            xml.push_str(r#"</sheetView></sheetViews>"#);
        }

        // Column widths
        if !self.column_widths.is_empty() {
            xml.push_str("<cols>");
            for (min, max, width) in &self.column_widths {
                xml.push_str(&format!(
                    r#"<col min="{}" max="{}" width="{}" customWidth="1"/>"#,
                    min, max, width
                ));
            }
            xml.push_str("</cols>");
        }

        // Sheet data
        xml.push_str("<sheetData>");

        // Group cells by row
        let mut rows: std::collections::BTreeMap<u32, Vec<&CellData>> =
            std::collections::BTreeMap::new();
        for cell in &self.cells {
            let row_num = parse_row_from_ref(&cell.cell_ref);
            rows.entry(row_num).or_default().push(cell);
        }

        for (row_num, cells) in rows {
            // Check if this row has a custom height
            let height_attr = self
                .row_heights
                .iter()
                .find(|(r, _)| *r == row_num)
                .map(|(_, h)| format!(r#" ht="{}" customHeight="1""#, h))
                .unwrap_or_default();

            xml.push_str(&format!(r#"<row r="{}"{}>"#, row_num, height_attr));

            for cell in cells {
                xml.push_str(&format!(r#"<c r="{}""#, cell.cell_ref));

                if let Some(ref t) = cell.cell_type {
                    xml.push_str(&format!(r#" t="{}""#, t));
                }

                if let Some(s) = cell.style_index {
                    xml.push_str(&format!(r#" s="{}""#, s));
                }

                xml.push('>');

                if let Some(ref formula) = cell.formula {
                    xml.push_str(&format!("<f>{}</f>", escape_xml(formula)));
                }

                if cell.cell_type.as_deref() == Some("inlineStr") {
                    xml.push_str(&format!("<is><t>{}</t></is>", escape_xml(&cell.value)));
                } else {
                    xml.push_str(&format!("<v>{}</v>", escape_xml(&cell.value)));
                }

                xml.push_str("</c>");
            }

            xml.push_str("</row>");
        }

        xml.push_str("</sheetData>");

        // Merge cells
        if !self.merges.is_empty() {
            xml.push_str(&format!(r#"<mergeCells count="{}">"#, self.merges.len()));
            for merge in &self.merges {
                xml.push_str(&format!(r#"<mergeCell ref="{}"/>"#, merge));
            }
            xml.push_str("</mergeCells>");
        }

        // Hyperlinks
        if !self.hyperlinks.is_empty() {
            xml.push_str("<hyperlinks>");
            for (cell_ref, rel_id, display) in &self.hyperlinks {
                xml.push_str(&format!(
                    r#"<hyperlink ref="{}" r:id="{}" display="{}"/>"#,
                    cell_ref,
                    rel_id,
                    escape_xml(display)
                ));
            }
            xml.push_str("</hyperlinks>");
        }

        xml.push_str("</worksheet>");
        xml
    }
}

/// XLSX builder for creating test fixtures (low-level ZIP approach)
#[derive(Clone, Debug, Default)]
pub struct XlsxBuilder {
    sheets: Vec<SheetBuilder>,
    shared_strings: Vec<String>,
    styles: Option<String>,
    theme: Option<String>,
    workbook_rels_extra: Vec<(String, String, String)>, // (id, type, target)
}

impl XlsxBuilder {
    pub fn new() -> Self {
        Self::default()
    }

    /// Add a sheet to the workbook
    pub fn add_sheet(&mut self, name: &str) -> &mut SheetBuilder {
        self.sheets.push(SheetBuilder::new(name));
        self.sheets.last_mut().unwrap()
    }

    /// Add a pre-built sheet
    pub fn with_sheet(mut self, sheet: SheetBuilder) -> Self {
        self.sheets.push(sheet);
        self
    }

    /// Set shared strings
    pub fn with_shared_strings(mut self, strings: Vec<String>) -> Self {
        self.shared_strings = strings;
        self
    }

    /// Add a shared string and return its index
    pub fn add_shared_string(&mut self, s: &str) -> u32 {
        let index = self.shared_strings.len() as u32;
        self.shared_strings.push(s.to_string());
        index
    }

    /// Set styles XML
    pub fn with_styles(mut self, xml: &str) -> Self {
        self.styles = Some(xml.to_string());
        self
    }

    /// Set theme XML
    pub fn with_theme(mut self, xml: &str) -> Self {
        self.theme = Some(xml.to_string());
        self
    }

    /// Add an extra relationship to workbook.xml.rels
    pub fn add_workbook_rel(&mut self, id: &str, rel_type: &str, target: &str) -> &mut Self {
        self.workbook_rels_extra
            .push((id.to_string(), rel_type.to_string(), target.to_string()));
        self
    }

    /// Build the XLSX file as bytes
    pub fn build(&self) -> Vec<u8> {
        let mut buffer = Cursor::new(Vec::new());
        {
            let mut zip = ZipWriter::new(&mut buffer);
            let options =
                FileOptions::<()>::default().compression_method(zip::CompressionMethod::Deflated);

            // Build content types
            zip.start_file("[Content_Types].xml", options).unwrap();
            zip.write_all(self.build_content_types().as_bytes())
                .unwrap();

            // _rels/.rels
            zip.start_file("_rels/.rels", options).unwrap();
            zip.write_all(self.build_root_rels().as_bytes()).unwrap();

            // xl/workbook.xml
            zip.start_file("xl/workbook.xml", options).unwrap();
            zip.write_all(self.build_workbook().as_bytes()).unwrap();

            // xl/_rels/workbook.xml.rels
            zip.start_file("xl/_rels/workbook.xml.rels", options)
                .unwrap();
            zip.write_all(self.build_workbook_rels().as_bytes())
                .unwrap();

            // Worksheets
            for (i, sheet) in self.sheets.iter().enumerate() {
                let path = format!("xl/worksheets/sheet{}.xml", i + 1);
                zip.start_file(&path, options).unwrap();
                zip.write_all(sheet.build_xml().as_bytes()).unwrap();
            }

            // If no sheets, add an empty one
            if self.sheets.is_empty() {
                zip.start_file("xl/worksheets/sheet1.xml", options).unwrap();
                zip.write_all(EMPTY_SHEET_XML.as_bytes()).unwrap();
            }

            // Shared strings
            if !self.shared_strings.is_empty() {
                zip.start_file("xl/sharedStrings.xml", options).unwrap();
                zip.write_all(self.build_shared_strings().as_bytes())
                    .unwrap();
            }

            // Styles
            if let Some(ref styles) = self.styles {
                zip.start_file("xl/styles.xml", options).unwrap();
                zip.write_all(styles.as_bytes()).unwrap();
            }

            // Theme
            if let Some(ref theme) = self.theme {
                zip.start_file("xl/theme/theme1.xml", options).unwrap();
                zip.write_all(theme.as_bytes()).unwrap();
            }

            zip.finish().unwrap();
        }
        buffer.into_inner()
    }

    fn build_content_types(&self) -> String {
        let mut xml = String::from(
            r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>"#,
        );

        let sheet_count = if self.sheets.is_empty() {
            1
        } else {
            self.sheets.len()
        };
        for i in 1..=sheet_count {
            xml.push_str(&format!(
                r#"
  <Override PartName="/xl/worksheets/sheet{}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>"#,
                i
            ));
        }

        if !self.shared_strings.is_empty() {
            xml.push_str(r#"
  <Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>"#);
        }

        if self.styles.is_some() {
            xml.push_str(r#"
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>"#);
        }

        if self.theme.is_some() {
            xml.push_str(r#"
  <Override PartName="/xl/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.theme+xml"/>"#);
        }

        xml.push_str("\n</Types>");
        xml
    }

    fn build_root_rels(&self) -> String {
        String::from(RELS_XML)
    }

    fn build_workbook(&self) -> String {
        let mut xml = String::from(
            r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>"#,
        );

        if self.sheets.is_empty() {
            xml.push_str(
                r#"
    <sheet name="Sheet1" sheetId="1" r:id="rId1"/>"#,
            );
        } else {
            for (i, sheet) in self.sheets.iter().enumerate() {
                xml.push_str(&format!(
                    r#"
    <sheet name="{}" sheetId="{}" r:id="rId{}"/>"#,
                    escape_xml(&sheet.name),
                    i + 1,
                    i + 1
                ));
            }
        }

        xml.push_str(
            r#"
  </sheets>
</workbook>"#,
        );
        xml
    }

    fn build_workbook_rels(&self) -> String {
        let mut xml = String::from(
            r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">"#,
        );

        let sheet_count = if self.sheets.is_empty() {
            1
        } else {
            self.sheets.len()
        };
        for i in 1..=sheet_count {
            xml.push_str(&format!(
                r#"
  <Relationship Id="rId{}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet{}.xml"/>"#,
                i, i
            ));
        }

        let mut next_id = sheet_count + 1;

        if !self.shared_strings.is_empty() {
            xml.push_str(&format!(
                r#"
  <Relationship Id="rId{}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>"#,
                next_id
            ));
            next_id += 1;
        }

        if self.styles.is_some() {
            xml.push_str(&format!(
                r#"
  <Relationship Id="rId{}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>"#,
                next_id
            ));
            next_id += 1;
        }

        if self.theme.is_some() {
            xml.push_str(&format!(
                r#"
  <Relationship Id="rId{}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>"#,
                next_id
            ));
            next_id += 1;
        }

        // Extra relationships
        for (id, rel_type, target) in &self.workbook_rels_extra {
            xml.push_str(&format!(
                r#"
  <Relationship Id="{}" Type="{}" Target="{}"/>"#,
                id, rel_type, target
            ));
        }
        let _ = next_id; // Suppress unused warning

        xml.push_str("\n</Relationships>");
        xml
    }

    fn build_shared_strings(&self) -> String {
        let mut xml = format!(
            r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="{}" uniqueCount="{}">"#,
            self.shared_strings.len(),
            self.shared_strings.len()
        );

        for s in &self.shared_strings {
            xml.push_str(&format!("<si><t>{}</t></si>", escape_xml(s)));
        }

        xml.push_str("</sst>");
        xml
    }
}

/// Create a minimal styles.xml
pub fn create_minimal_styles() -> String {
    r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="1">
    <font>
      <sz val="11"/>
      <color theme="1"/>
      <name val="Calibri"/>
      <family val="2"/>
      <scheme val="minor"/>
    </font>
  </fonts>
  <fills count="2">
    <fill><patternFill patternType="none"/></fill>
    <fill><patternFill patternType="gray125"/></fill>
  </fills>
  <borders count="1">
    <border>
      <left/><right/><top/><bottom/><diagonal/>
    </border>
  </borders>
  <cellStyleXfs count="1">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
  </cellStyleXfs>
  <cellXfs count="1">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
  </cellXfs>
</styleSheet>"#
        .to_string()
}

/// Create styles with custom number formats
pub fn create_styles_with_numfmts(formats: &[(u32, &str)]) -> String {
    let mut xml = String::from(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">"#,
    );

    // Number formats
    if !formats.is_empty() {
        xml.push_str(&format!(r#"<numFmts count="{}">"#, formats.len()));
        for (id, code) in formats {
            xml.push_str(&format!(
                r#"<numFmt numFmtId="{}" formatCode="{}"/>"#,
                id,
                escape_xml(code)
            ));
        }
        xml.push_str("</numFmts>");
    }

    xml.push_str(
        r#"
  <fonts count="1">
    <font>
      <sz val="11"/>
      <name val="Calibri"/>
    </font>
  </fonts>
  <fills count="2">
    <fill><patternFill patternType="none"/></fill>
    <fill><patternFill patternType="gray125"/></fill>
  </fills>
  <borders count="1">
    <border><left/><right/><top/><bottom/><diagonal/></border>
  </borders>
  <cellStyleXfs count="1">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
  </cellStyleXfs>
  <cellXfs count="1">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
  </cellXfs>
</styleSheet>"#,
    );

    xml
}

/// Create a minimal theme
pub fn create_minimal_theme() -> String {
    r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office Theme">
  <a:themeElements>
    <a:clrScheme name="Office">
      <a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1>
      <a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1>
      <a:dk2><a:srgbClr val="44546A"/></a:dk2>
      <a:lt2><a:srgbClr val="E7E6E6"/></a:lt2>
      <a:accent1><a:srgbClr val="4472C4"/></a:accent1>
      <a:accent2><a:srgbClr val="ED7D31"/></a:accent2>
      <a:accent3><a:srgbClr val="A5A5A5"/></a:accent3>
      <a:accent4><a:srgbClr val="FFC000"/></a:accent4>
      <a:accent5><a:srgbClr val="5B9BD5"/></a:accent5>
      <a:accent6><a:srgbClr val="70AD47"/></a:accent6>
      <a:hlink><a:srgbClr val="0563C1"/></a:hlink>
      <a:folHlink><a:srgbClr val="954F72"/></a:folHlink>
    </a:clrScheme>
    <a:fontScheme name="Office">
      <a:majorFont><a:latin typeface="Calibri Light"/></a:majorFont>
      <a:minorFont><a:latin typeface="Calibri"/></a:minorFont>
    </a:fontScheme>
    <a:fmtScheme name="Office"/>
  </a:themeElements>
</a:theme>"#
        .to_string()
}

// Helper functions

fn escape_xml(s: &str) -> String {
    s.replace('&', "&amp;")
        .replace('<', "&lt;")
        .replace('>', "&gt;")
        .replace('"', "&quot;")
        .replace('\'', "&apos;")
}

fn parse_row_from_ref(cell_ref: &str) -> u32 {
    let row_str: String = cell_ref.chars().filter(|c| c.is_ascii_digit()).collect();
    row_str.parse().unwrap_or(1)
}

fn col_to_letter(col: u32) -> String {
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

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_create_minimal_xlsx() {
        let data = create_minimal_xlsx();
        assert!(!data.is_empty());
        // Verify it's a valid ZIP
        let cursor = Cursor::new(data);
        let archive = zip::ZipArchive::new(cursor).unwrap();
        assert!(archive.len() >= 5);
    }

    #[test]
    fn test_xlsx_builder() {
        let mut builder = XlsxBuilder::new();
        builder
            .add_sheet("Test")
            .add_number("A1", 42.0)
            .add_number("B1", 3.14);

        let data = builder.build();
        assert!(!data.is_empty());

        let cursor = Cursor::new(data);
        let archive = zip::ZipArchive::new(cursor).unwrap();
        assert!(archive.len() >= 5);
    }

    #[test]
    fn test_xlsx_with_shared_strings() {
        let builder =
            XlsxBuilder::new().with_shared_strings(vec!["Hello".to_string(), "World".to_string()]);

        let data = builder.build();
        let cursor = Cursor::new(data);
        let mut archive = zip::ZipArchive::new(cursor).unwrap();

        // Verify shared strings file exists
        let sst = archive.by_name("xl/sharedStrings.xml");
        assert!(sst.is_ok());
    }

    #[test]
    fn test_col_to_letter() {
        assert_eq!(col_to_letter(1), "A");
        assert_eq!(col_to_letter(26), "Z");
        assert_eq!(col_to_letter(27), "AA");
        assert_eq!(col_to_letter(28), "AB");
    }

    #[test]
    fn test_parse_row_from_ref() {
        assert_eq!(parse_row_from_ref("A1"), 1);
        assert_eq!(parse_row_from_ref("B10"), 10);
        assert_eq!(parse_row_from_ref("AA100"), 100);
    }
}
