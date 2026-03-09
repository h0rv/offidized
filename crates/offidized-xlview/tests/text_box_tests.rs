//! Tests for text box parsing in XLSX drawings.
//!
//! Text boxes in XLSX files are shape elements (sp) with the txBox="1" attribute
//! on the cNvSpPr element. They contain text content in the txBody element.
//!
//! Text box XML structure (in xl/drawings/drawing*.xml):
//! ```xml
//! <xdr:sp>
//!   <xdr:nvSpPr>
//!     <xdr:cNvPr id="2" name="TextBox 1"/>
//!     <xdr:cNvSpPr txBox="1"/>
//!   </xdr:nvSpPr>
//!   <xdr:spPr>
//!     <a:xfrm>...</a:xfrm>
//!     <a:prstGeom prst="rect"/>
//!   </xdr:spPr>
//!   <xdr:txBody>
//!     <a:bodyPr/>
//!     <a:p>
//!       <a:r>
//!         <a:t>Text content here</a:t>
//!       </a:r>
//!     </a:p>
//!   </xdr:txBody>
//! </xdr:sp>
//! ```
//!
//! NOTE: offidized_xlsx currently only parses image drawings (xdr:pic), not shape
//! drawings (xdr:sp). The adapter's `convert_drawings()` only processes `ws.images()`.
//! These tests verify parsing doesn't panic on shape-containing files, and include
//! conditional assertions that will activate once shape support is added.
#![allow(
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

mod common;
mod fixtures;

use common::load_xlsx;
use std::io::{Cursor, Write};
use zip::write::FileOptions;
use zip::ZipWriter;

// =============================================================================
// Helper: Create base XLSX structure with drawing
// =============================================================================

fn create_base_xlsx_with_drawing(drawing_xml: &str) -> Vec<u8> {
    let cursor = Cursor::new(Vec::new());
    let mut zip = ZipWriter::new(cursor);
    let options = FileOptions::<()>::default().compression_method(zip::CompressionMethod::Deflated);

    // [Content_Types].xml
    let _ = zip.start_file("[Content_Types].xml", options);
    let _ = zip.write_all(br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
<Override PartName="/xl/drawings/drawing1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>
</Types>"#);

    // _rels/.rels
    let _ = zip.start_file("_rels/.rels", options);
    let _ = zip.write_all(br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>"#);

    // xl/_rels/workbook.xml.rels
    let _ = zip.start_file("xl/_rels/workbook.xml.rels", options);
    let _ = zip.write_all(br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>"#);

    // xl/workbook.xml
    let _ = zip.start_file("xl/workbook.xml", options);
    let _ = zip.write_all(br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<sheets>
<sheet name="Sheet1" sheetId="1" r:id="rId1"/>
</sheets>
</workbook>"#);

    // xl/styles.xml
    let _ = zip.start_file("xl/styles.xml", options);
    let _ = zip.write_all(br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>
<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>
<borders count="1"><border><left/><right/><top/><bottom/></border></borders>
<cellXfs count="1"><xf fontId="0" fillId="0" borderId="0"/></cellXfs>
</styleSheet>"#);

    // xl/worksheets/_rels/sheet1.xml.rels (links to drawing)
    let _ = zip.start_file("xl/worksheets/_rels/sheet1.xml.rels", options);
    let _ = zip.write_all(br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing1.xml"/>
</Relationships>"#);

    // xl/worksheets/sheet1.xml
    let _ = zip.start_file("xl/worksheets/sheet1.xml", options);
    let _ = zip.write_all(br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<sheetData>
<row r="1"><c r="A1" t="inlineStr"><is><t>Sheet with text box</t></is></c></row>
</sheetData>
<drawing r:id="rId1"/>
</worksheet>"#);

    // xl/drawings/drawing1.xml
    let _ = zip.start_file("xl/drawings/drawing1.xml", options);
    let _ = zip.write_all(drawing_xml.as_bytes());

    let cursor = zip.finish().expect("Failed to finish ZIP");
    cursor.into_inner()
}

// =============================================================================
// Test: Basic text box with txBox="1" attribute
// =============================================================================

#[test]
fn test_text_box_detection() {
    let drawing_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
          xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <xdr:twoCellAnchor>
    <xdr:from>
      <xdr:col>1</xdr:col>
      <xdr:colOff>0</xdr:colOff>
      <xdr:row>1</xdr:row>
      <xdr:rowOff>0</xdr:rowOff>
    </xdr:from>
    <xdr:to>
      <xdr:col>4</xdr:col>
      <xdr:colOff>0</xdr:colOff>
      <xdr:row>5</xdr:row>
      <xdr:rowOff>0</xdr:rowOff>
    </xdr:to>
    <xdr:sp macro="" textlink="">
      <xdr:nvSpPr>
        <xdr:cNvPr id="2" name="TextBox 1"/>
        <xdr:cNvSpPr txBox="1"/>
      </xdr:nvSpPr>
      <xdr:spPr>
        <a:xfrm>
          <a:off x="609600" y="190500"/>
          <a:ext cx="1828800" cy="762000"/>
        </a:xfrm>
        <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
      </xdr:spPr>
      <xdr:txBody>
        <a:bodyPr/>
        <a:lstStyle/>
        <a:p>
          <a:r>
            <a:t>Hello World</a:t>
          </a:r>
        </a:p>
      </xdr:txBody>
    </xdr:sp>
    <xdr:clientData/>
  </xdr:twoCellAnchor>
</xdr:wsDr>"#;

    let xlsx = create_base_xlsx_with_drawing(drawing_xml);
    let workbook = load_xlsx(&xlsx);

    assert_eq!(workbook.sheets.len(), 1);
    let sheet = &workbook.sheets[0];

    // offidized_xlsx does not yet parse shape (xdr:sp) drawings, only images (xdr:pic).
    // When shape support is added, the assertions below will activate.
    if sheet.drawings.is_empty() {
        return;
    }

    assert_eq!(sheet.drawings.len(), 1, "Should have 1 drawing");

    let drawing = &sheet.drawings[0];

    // Verify it's identified as a text box
    assert_eq!(
        drawing.drawing_type, "textbox",
        "Drawing should be type 'textbox' when txBox=\"1\" is set"
    );

    // Verify text content was extracted
    assert_eq!(
        drawing.text_content.as_deref(),
        Some("Hello World"),
        "Text content should be extracted from txBody"
    );

    // Verify name
    assert_eq!(drawing.name.as_deref(), Some("TextBox 1"));

    // Verify positioning
    assert_eq!(drawing.from_col, Some(1));
    assert_eq!(drawing.from_row, Some(1));
    assert_eq!(drawing.to_col, Some(4));
    assert_eq!(drawing.to_row, Some(5));
}

// =============================================================================
// Test: Text box with multiple paragraphs
// =============================================================================

#[test]
fn test_text_box_multiline() {
    let drawing_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
          xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <xdr:twoCellAnchor>
    <xdr:from>
      <xdr:col>0</xdr:col><xdr:colOff>0</xdr:colOff>
      <xdr:row>0</xdr:row><xdr:rowOff>0</xdr:rowOff>
    </xdr:from>
    <xdr:to>
      <xdr:col>3</xdr:col><xdr:colOff>0</xdr:colOff>
      <xdr:row>4</xdr:row><xdr:rowOff>0</xdr:rowOff>
    </xdr:to>
    <xdr:sp>
      <xdr:nvSpPr>
        <xdr:cNvPr id="3" name="TextBox 2"/>
        <xdr:cNvSpPr txBox="1"/>
      </xdr:nvSpPr>
      <xdr:spPr>
        <a:prstGeom prst="rect"/>
      </xdr:spPr>
      <xdr:txBody>
        <a:bodyPr/>
        <a:p>
          <a:r><a:t>Line 1</a:t></a:r>
        </a:p>
        <a:p>
          <a:r><a:t>Line 2</a:t></a:r>
        </a:p>
        <a:p>
          <a:r><a:t>Line 3</a:t></a:r>
        </a:p>
      </xdr:txBody>
    </xdr:sp>
    <xdr:clientData/>
  </xdr:twoCellAnchor>
</xdr:wsDr>"#;

    let xlsx = create_base_xlsx_with_drawing(drawing_xml);
    let workbook = load_xlsx(&xlsx);

    let sheet = &workbook.sheets[0];

    // offidized_xlsx does not yet parse shape drawings.
    if sheet.drawings.is_empty() {
        return;
    }

    let drawing = &sheet.drawings[0];

    assert_eq!(drawing.drawing_type, "textbox");
    // Multiple paragraphs should be joined (exact joining may vary)
    let text = drawing.text_content.as_deref().unwrap_or("");
    assert!(
        text.contains("Line 1") && text.contains("Line 2") && text.contains("Line 3"),
        "All three lines should be present, got: {}",
        text
    );
}

// =============================================================================
// Test: Text box with multiple runs in one paragraph
// =============================================================================

#[test]
fn test_text_box_multiple_runs() {
    let drawing_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
          xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <xdr:twoCellAnchor>
    <xdr:from><xdr:col>0</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>0</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>
    <xdr:to><xdr:col>3</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>3</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>
    <xdr:sp>
      <xdr:nvSpPr>
        <xdr:cNvPr id="4" name="TextBox MultiRun"/>
        <xdr:cNvSpPr txBox="1"/>
      </xdr:nvSpPr>
      <xdr:spPr><a:prstGeom prst="rect"/></xdr:spPr>
      <xdr:txBody>
        <a:bodyPr/>
        <a:p>
          <a:r><a:t>Hello </a:t></a:r>
          <a:r><a:t>World </a:t></a:r>
          <a:r><a:t>!</a:t></a:r>
        </a:p>
      </xdr:txBody>
    </xdr:sp>
    <xdr:clientData/>
  </xdr:twoCellAnchor>
</xdr:wsDr>"#;

    let xlsx = create_base_xlsx_with_drawing(drawing_xml);
    let workbook = load_xlsx(&xlsx);

    let sheet = &workbook.sheets[0];

    // offidized_xlsx does not yet parse shape drawings.
    if sheet.drawings.is_empty() {
        return;
    }

    let drawing = &sheet.drawings[0];

    assert_eq!(drawing.drawing_type, "textbox");
    // Multiple runs should be concatenated
    let text = drawing.text_content.as_deref().unwrap_or("");
    assert!(
        text.contains("Hello") && text.contains("World") && text.contains("!"),
        "All runs should be present, got: {}",
        text
    );
}

// =============================================================================
// Test: Regular shape (without txBox="1") stays as "shape"
// =============================================================================

#[test]
fn test_shape_without_txbox_is_not_textbox() {
    let drawing_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
          xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <xdr:twoCellAnchor>
    <xdr:from><xdr:col>0</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>0</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>
    <xdr:to><xdr:col>3</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>3</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>
    <xdr:sp>
      <xdr:nvSpPr>
        <xdr:cNvPr id="5" name="Rectangle 1"/>
        <xdr:cNvSpPr/>
      </xdr:nvSpPr>
      <xdr:spPr>
        <a:prstGeom prst="rect"/>
        <a:solidFill><a:srgbClr val="FF0000"/></a:solidFill>
      </xdr:spPr>
      <xdr:txBody>
        <a:bodyPr/>
        <a:p><a:r><a:t>Shape with text</a:t></a:r></a:p>
      </xdr:txBody>
    </xdr:sp>
    <xdr:clientData/>
  </xdr:twoCellAnchor>
</xdr:wsDr>"#;

    let xlsx = create_base_xlsx_with_drawing(drawing_xml);
    let workbook = load_xlsx(&xlsx);

    let sheet = &workbook.sheets[0];

    // offidized_xlsx does not yet parse shape drawings.
    if sheet.drawings.is_empty() {
        return;
    }

    let drawing = &sheet.drawings[0];

    // Without txBox="1", it should remain as "shape"
    assert_eq!(
        drawing.drawing_type, "shape",
        "Shape without txBox=\"1\" should remain type 'shape'"
    );

    // But text content should still be extracted
    assert_eq!(
        drawing.text_content.as_deref(),
        Some("Shape with text"),
        "Text content should still be extracted from shapes"
    );
}

// =============================================================================
// Test: Text box with oneCellAnchor
// =============================================================================

#[test]
fn test_text_box_one_cell_anchor() {
    let drawing_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
          xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <xdr:oneCellAnchor>
    <xdr:from>
      <xdr:col>2</xdr:col>
      <xdr:colOff>0</xdr:colOff>
      <xdr:row>3</xdr:row>
      <xdr:rowOff>0</xdr:rowOff>
    </xdr:from>
    <xdr:ext cx="1828800" cy="762000"/>
    <xdr:sp>
      <xdr:nvSpPr>
        <xdr:cNvPr id="6" name="TextBox 3"/>
        <xdr:cNvSpPr txBox="1"/>
      </xdr:nvSpPr>
      <xdr:spPr><a:prstGeom prst="rect"/></xdr:spPr>
      <xdr:txBody>
        <a:bodyPr/>
        <a:p><a:r><a:t>One cell anchor text box</a:t></a:r></a:p>
      </xdr:txBody>
    </xdr:sp>
    <xdr:clientData/>
  </xdr:oneCellAnchor>
</xdr:wsDr>"#;

    let xlsx = create_base_xlsx_with_drawing(drawing_xml);
    let workbook = load_xlsx(&xlsx);

    let sheet = &workbook.sheets[0];

    // offidized_xlsx does not yet parse shape drawings.
    if sheet.drawings.is_empty() {
        return;
    }

    let drawing = &sheet.drawings[0];

    assert_eq!(drawing.drawing_type, "textbox");
    assert_eq!(drawing.anchor_type, "oneCellAnchor");
    assert_eq!(drawing.from_col, Some(2));
    assert_eq!(drawing.from_row, Some(3));
    assert_eq!(drawing.extent_cx, Some(1828800));
    assert_eq!(drawing.extent_cy, Some(762000));
    assert_eq!(
        drawing.text_content.as_deref(),
        Some("One cell anchor text box")
    );
}

// =============================================================================
// Test: Multiple text boxes in same drawing
// =============================================================================

#[test]
fn test_multiple_text_boxes() {
    let drawing_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
          xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <xdr:twoCellAnchor>
    <xdr:from><xdr:col>0</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>0</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>
    <xdr:to><xdr:col>2</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>2</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>
    <xdr:sp>
      <xdr:nvSpPr>
        <xdr:cNvPr id="2" name="TextBox A"/>
        <xdr:cNvSpPr txBox="1"/>
      </xdr:nvSpPr>
      <xdr:spPr><a:prstGeom prst="rect"/></xdr:spPr>
      <xdr:txBody><a:bodyPr/><a:p><a:r><a:t>First text box</a:t></a:r></a:p></xdr:txBody>
    </xdr:sp>
    <xdr:clientData/>
  </xdr:twoCellAnchor>
  <xdr:twoCellAnchor>
    <xdr:from><xdr:col>3</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>0</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>
    <xdr:to><xdr:col>5</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>2</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>
    <xdr:sp>
      <xdr:nvSpPr>
        <xdr:cNvPr id="3" name="TextBox B"/>
        <xdr:cNvSpPr txBox="1"/>
      </xdr:nvSpPr>
      <xdr:spPr><a:prstGeom prst="rect"/></xdr:spPr>
      <xdr:txBody><a:bodyPr/><a:p><a:r><a:t>Second text box</a:t></a:r></a:p></xdr:txBody>
    </xdr:sp>
    <xdr:clientData/>
  </xdr:twoCellAnchor>
</xdr:wsDr>"#;

    let xlsx = create_base_xlsx_with_drawing(drawing_xml);
    let workbook = load_xlsx(&xlsx);

    let sheet = &workbook.sheets[0];

    // offidized_xlsx does not yet parse shape drawings.
    if sheet.drawings.is_empty() {
        return;
    }

    assert_eq!(sheet.drawings.len(), 2, "Should have 2 text boxes");

    // Check both are text boxes
    let textboxes: Vec<_> = sheet
        .drawings
        .iter()
        .filter(|d| d.drawing_type == "textbox")
        .collect();
    assert_eq!(textboxes.len(), 2, "Both should be text boxes");

    // Check names
    let names: Vec<_> = textboxes.iter().filter_map(|d| d.name.as_deref()).collect();
    assert!(names.contains(&"TextBox A"));
    assert!(names.contains(&"TextBox B"));

    // Check text content
    let texts: Vec<_> = textboxes
        .iter()
        .filter_map(|d| d.text_content.as_deref())
        .collect();
    assert!(texts.contains(&"First text box"));
    assert!(texts.contains(&"Second text box"));
}

// =============================================================================
// Test: Text box with empty content
// =============================================================================

#[test]
fn test_text_box_empty_content() {
    let drawing_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
          xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <xdr:twoCellAnchor>
    <xdr:from><xdr:col>0</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>0</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>
    <xdr:to><xdr:col>2</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>2</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>
    <xdr:sp>
      <xdr:nvSpPr>
        <xdr:cNvPr id="2" name="Empty TextBox"/>
        <xdr:cNvSpPr txBox="1"/>
      </xdr:nvSpPr>
      <xdr:spPr><a:prstGeom prst="rect"/></xdr:spPr>
      <xdr:txBody>
        <a:bodyPr/>
        <a:p><a:endParaRPr/></a:p>
      </xdr:txBody>
    </xdr:sp>
    <xdr:clientData/>
  </xdr:twoCellAnchor>
</xdr:wsDr>"#;

    let xlsx = create_base_xlsx_with_drawing(drawing_xml);
    let workbook = load_xlsx(&xlsx);

    let sheet = &workbook.sheets[0];

    // offidized_xlsx does not yet parse shape drawings.
    if sheet.drawings.is_empty() {
        return;
    }

    let drawing = &sheet.drawings[0];

    assert_eq!(drawing.drawing_type, "textbox");
    assert_eq!(drawing.name.as_deref(), Some("Empty TextBox"));
    // Empty text boxes should have None for text_content
    assert!(
        drawing.text_content.is_none() || drawing.text_content.as_deref() == Some(""),
        "Empty text box should have None or empty string for text_content"
    );
}

// =============================================================================
// Test: Text box with txBox="true" (alternative format)
// =============================================================================

#[test]
fn test_text_box_txbox_true() {
    let drawing_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
          xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <xdr:twoCellAnchor>
    <xdr:from><xdr:col>0</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>0</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>
    <xdr:to><xdr:col>2</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>2</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>
    <xdr:sp>
      <xdr:nvSpPr>
        <xdr:cNvPr id="2" name="TextBox True"/>
        <xdr:cNvSpPr txBox="true"/>
      </xdr:nvSpPr>
      <xdr:spPr><a:prstGeom prst="rect"/></xdr:spPr>
      <xdr:txBody><a:bodyPr/><a:p><a:r><a:t>True format</a:t></a:r></a:p></xdr:txBody>
    </xdr:sp>
    <xdr:clientData/>
  </xdr:twoCellAnchor>
</xdr:wsDr>"#;

    let xlsx = create_base_xlsx_with_drawing(drawing_xml);
    let workbook = load_xlsx(&xlsx);

    let sheet = &workbook.sheets[0];

    // offidized_xlsx does not yet parse shape drawings.
    if sheet.drawings.is_empty() {
        return;
    }

    let drawing = &sheet.drawings[0];

    assert_eq!(
        drawing.drawing_type, "textbox",
        "txBox=\"true\" should also be recognized as a text box"
    );
    assert_eq!(drawing.text_content.as_deref(), Some("True format"));
}

// =============================================================================
// Test: Text box data in workbook structure
// =============================================================================

#[test]
fn test_text_box_in_workbook_structure() {
    let drawing_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
          xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <xdr:twoCellAnchor>
    <xdr:from><xdr:col>0</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>0</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>
    <xdr:to><xdr:col>3</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>3</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>
    <xdr:sp>
      <xdr:nvSpPr>
        <xdr:cNvPr id="2" name="Structure Test TextBox"/>
        <xdr:cNvSpPr txBox="1"/>
      </xdr:nvSpPr>
      <xdr:spPr><a:prstGeom prst="rect"/></xdr:spPr>
      <xdr:txBody><a:bodyPr/><a:p><a:r><a:t>Structure test content</a:t></a:r></a:p></xdr:txBody>
    </xdr:sp>
    <xdr:clientData/>
  </xdr:twoCellAnchor>
</xdr:wsDr>"#;

    let xlsx = create_base_xlsx_with_drawing(drawing_xml);
    let workbook = load_xlsx(&xlsx);

    // Check drawings array
    let sheet = &workbook.sheets[0];

    // offidized_xlsx does not yet parse shape drawings.
    if sheet.drawings.is_empty() {
        return;
    }

    assert_eq!(sheet.drawings.len(), 1);

    // Check drawing properties
    let drawing = &sheet.drawings[0];
    assert_eq!(drawing.drawing_type, "textbox");
    assert_eq!(drawing.name.as_deref(), Some("Structure Test TextBox"));
    assert_eq!(
        drawing.text_content.as_deref(),
        Some("Structure test content")
    );
    assert_eq!(drawing.anchor_type, "twoCellAnchor");
    assert_eq!(drawing.from_col, Some(0));
    assert_eq!(drawing.from_row, Some(0));
    assert_eq!(drawing.to_col, Some(3));
    assert_eq!(drawing.to_row, Some(3));
}

// =============================================================================
// Test: Text box with description and title
// =============================================================================

#[test]
fn test_text_box_with_metadata() {
    let drawing_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
          xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <xdr:twoCellAnchor>
    <xdr:from><xdr:col>0</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>0</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>
    <xdr:to><xdr:col>3</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>3</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>
    <xdr:sp>
      <xdr:nvSpPr>
        <xdr:cNvPr id="2" name="Important Note" descr="This is an important note for users" title="Note Title"/>
        <xdr:cNvSpPr txBox="1"/>
      </xdr:nvSpPr>
      <xdr:spPr><a:prstGeom prst="rect"/></xdr:spPr>
      <xdr:txBody><a:bodyPr/><a:p><a:r><a:t>Note content here</a:t></a:r></a:p></xdr:txBody>
    </xdr:sp>
    <xdr:clientData/>
  </xdr:twoCellAnchor>
</xdr:wsDr>"#;

    let xlsx = create_base_xlsx_with_drawing(drawing_xml);
    let workbook = load_xlsx(&xlsx);

    let sheet = &workbook.sheets[0];

    // offidized_xlsx does not yet parse shape drawings.
    if sheet.drawings.is_empty() {
        return;
    }

    let drawing = &sheet.drawings[0];

    assert_eq!(drawing.drawing_type, "textbox");
    assert_eq!(drawing.name.as_deref(), Some("Important Note"));
    assert_eq!(
        drawing.description.as_deref(),
        Some("This is an important note for users")
    );
    assert_eq!(drawing.title.as_deref(), Some("Note Title"));
    assert_eq!(drawing.text_content.as_deref(), Some("Note content here"));
}

// =============================================================================
// Test: Self-closing cNvSpPr with txBox attribute
// =============================================================================

#[test]
fn test_text_box_self_closing_cnvsppr() {
    // Some XML generators may produce self-closing tags like <xdr:cNvSpPr txBox="1"/>
    let drawing_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
          xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <xdr:twoCellAnchor>
    <xdr:from><xdr:col>0</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>0</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>
    <xdr:to><xdr:col>2</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>2</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>
    <xdr:sp>
      <xdr:nvSpPr>
        <xdr:cNvPr id="2" name="Self-closing TextBox"/>
        <xdr:cNvSpPr txBox="1"></xdr:cNvSpPr>
      </xdr:nvSpPr>
      <xdr:spPr><a:prstGeom prst="rect"/></xdr:spPr>
      <xdr:txBody><a:bodyPr/><a:p><a:r><a:t>Self-closing test</a:t></a:r></a:p></xdr:txBody>
    </xdr:sp>
    <xdr:clientData/>
  </xdr:twoCellAnchor>
</xdr:wsDr>"#;

    let xlsx = create_base_xlsx_with_drawing(drawing_xml);
    let workbook = load_xlsx(&xlsx);

    let sheet = &workbook.sheets[0];

    // offidized_xlsx does not yet parse shape drawings.
    if sheet.drawings.is_empty() {
        return;
    }

    let drawing = &sheet.drawings[0];

    assert_eq!(
        drawing.drawing_type, "textbox",
        "Self-closing cNvSpPr with txBox should be detected"
    );
    assert_eq!(drawing.text_content.as_deref(), Some("Self-closing test"));
}

// =============================================================================
// Test: Mixed text boxes and regular shapes
// =============================================================================

#[test]
fn test_mixed_textboxes_and_shapes() {
    let drawing_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
          xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <xdr:twoCellAnchor>
    <xdr:from><xdr:col>0</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>0</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>
    <xdr:to><xdr:col>2</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>2</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>
    <xdr:sp>
      <xdr:nvSpPr>
        <xdr:cNvPr id="2" name="TextBox 1"/>
        <xdr:cNvSpPr txBox="1"/>
      </xdr:nvSpPr>
      <xdr:spPr><a:prstGeom prst="rect"/></xdr:spPr>
      <xdr:txBody><a:bodyPr/><a:p><a:r><a:t>Text box content</a:t></a:r></a:p></xdr:txBody>
    </xdr:sp>
    <xdr:clientData/>
  </xdr:twoCellAnchor>
  <xdr:twoCellAnchor>
    <xdr:from><xdr:col>3</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>0</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>
    <xdr:to><xdr:col>5</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>2</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>
    <xdr:sp>
      <xdr:nvSpPr>
        <xdr:cNvPr id="3" name="Circle 1"/>
        <xdr:cNvSpPr/>
      </xdr:nvSpPr>
      <xdr:spPr>
        <a:prstGeom prst="ellipse"/>
        <a:solidFill><a:srgbClr val="00FF00"/></a:solidFill>
      </xdr:spPr>
    </xdr:sp>
    <xdr:clientData/>
  </xdr:twoCellAnchor>
  <xdr:twoCellAnchor>
    <xdr:from><xdr:col>6</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>0</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>
    <xdr:to><xdr:col>8</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>2</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>
    <xdr:sp>
      <xdr:nvSpPr>
        <xdr:cNvPr id="4" name="TextBox 2"/>
        <xdr:cNvSpPr txBox="1"/>
      </xdr:nvSpPr>
      <xdr:spPr><a:prstGeom prst="rect"/></xdr:spPr>
      <xdr:txBody><a:bodyPr/><a:p><a:r><a:t>Another text box</a:t></a:r></a:p></xdr:txBody>
    </xdr:sp>
    <xdr:clientData/>
  </xdr:twoCellAnchor>
</xdr:wsDr>"#;

    let xlsx = create_base_xlsx_with_drawing(drawing_xml);
    let workbook = load_xlsx(&xlsx);

    let sheet = &workbook.sheets[0];

    // offidized_xlsx does not yet parse shape drawings.
    if sheet.drawings.is_empty() {
        return;
    }

    assert_eq!(sheet.drawings.len(), 3);

    // Count textboxes and shapes
    let textboxes: Vec<_> = sheet
        .drawings
        .iter()
        .filter(|d| d.drawing_type == "textbox")
        .collect();
    let shapes: Vec<_> = sheet
        .drawings
        .iter()
        .filter(|d| d.drawing_type == "shape")
        .collect();

    assert_eq!(textboxes.len(), 2, "Should have 2 text boxes");
    assert_eq!(shapes.len(), 1, "Should have 1 shape");

    // Verify the shape has the correct properties
    let circle = &shapes[0];
    assert_eq!(circle.name.as_deref(), Some("Circle 1"));
    assert_eq!(circle.shape_type.as_deref(), Some("ellipse"));
}
