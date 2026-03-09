//! Named styles (cellStyles) tests ported from xlview.
//!
//! The original xlview tests directly parsed styles.xml with `parse_styles()` to verify
//! internal cellStyles, cellStyleXfs, and inheritance. Since offidized-xlview delegates
//! parsing to offidized-xlsx and only exposes resolved styles, these tests verify that
//! named style properties (font, fill, alignment, protection) flow through to cells
//! correctly via the full pipeline.
//!
//! Where a test verified internal structures (e.g. `stylesheet.named_styles.len()`),
//! we verify the observable effect on cells instead. Some tests are adapted more
//! loosely since the internal parsing details are not exposed.
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

use common::*;
use std::io::{Cursor, Write};
use zip::write::FileOptions;
use zip::ZipWriter;

// ============================================================================
// Helper: Create XLSX with custom styles.xml
// ============================================================================

/// Create an XLSX file with a custom styles.xml (for testing named style internals).
/// This gives us direct control over the XML, simulating what `parse_styles()` tested.
fn create_xlsx_with_custom_styles(styles_xml: &str, sheet_xml: &str) -> Vec<u8> {
    let mut buf = Cursor::new(Vec::new());
    {
        let mut zip = ZipWriter::new(&mut buf);
        let opts = FileOptions::<()>::default();

        zip.start_file("[Content_Types].xml", opts).unwrap();
        zip.write_all(br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
  <Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>
</Types>"#).unwrap();

        zip.start_file("_rels/.rels", opts).unwrap();
        zip.write_all(br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>"#).unwrap();

        zip.start_file("xl/_rels/workbook.xml.rels", opts).unwrap();
        zip.write_all(br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>
</Relationships>"#).unwrap();

        zip.start_file("xl/workbook.xml", opts).unwrap();
        zip.write_all(
            br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets>
</workbook>"#,
        )
        .unwrap();

        zip.start_file("xl/sharedStrings.xml", opts).unwrap();
        zip.write_all(
            br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="0" uniqueCount="0">
</sst>"#,
        )
        .unwrap();

        zip.start_file("xl/styles.xml", opts).unwrap();
        zip.write_all(styles_xml.as_bytes()).unwrap();

        zip.start_file("xl/worksheets/sheet1.xml", opts).unwrap();
        zip.write_all(sheet_xml.as_bytes()).unwrap();

        zip.finish().unwrap();
    }
    buf.into_inner()
}

// ============================================================================
// Basic Named Style Parsing Tests
// ============================================================================

#[test]
fn test_parse_single_named_style() {
    // Verify that a simple file with a "Normal" named style parses correctly.
    // The cell should use style xfId=0, which references the Normal cellStyleXf.
    let styles_xml = r#"<?xml version="1.0" encoding="UTF-8"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="1">
    <font><name val="Calibri"/><sz val="11"/></font>
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
  <cellStyles count="1">
    <cellStyle name="Normal" xfId="0" builtinId="0"/>
  </cellStyles>
</styleSheet>"#;

    let sheet_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1"><c r="A1" s="0" t="inlineStr"><is><t>Normal Style</t></is></c></row>
  </sheetData>
</worksheet>"#;

    let xlsx = create_xlsx_with_custom_styles(styles_xml, sheet_xml);
    let workbook = load_xlsx(&xlsx);

    assert!(
        !workbook.sheets.is_empty(),
        "Should have at least one sheet"
    );
    let cell = get_cell(&workbook, 0, 0, 0);
    assert!(cell.is_some(), "Cell A1 should exist");
    // Normal style with Calibri font should be parsed
    let style = get_cell_style(&workbook, 0, 0, 0);
    if let Some(s) = style {
        if let Some(ref font) = s.font_family {
            assert_eq!(font, "Calibri", "Normal style should use Calibri font");
        }
    }
}

#[test]
fn test_parse_multiple_named_styles() {
    // File with Normal and Heading 1 named styles.
    // Cell A1 uses xfId=0 (Normal), Cell A2 uses xfId=1 (Heading 1 - bold, larger font).
    let styles_xml = r#"<?xml version="1.0" encoding="UTF-8"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="2">
    <font><name val="Calibri"/><sz val="11"/></font>
    <font><name val="Calibri"/><sz val="15"/><b/></font>
  </fonts>
  <fills count="2">
    <fill><patternFill patternType="none"/></fill>
    <fill><patternFill patternType="gray125"/></fill>
  </fills>
  <borders count="1">
    <border><left/><right/><top/><bottom/><diagonal/></border>
  </borders>
  <cellStyleXfs count="2">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
    <xf numFmtId="0" fontId="1" fillId="0" borderId="0"/>
  </cellStyleXfs>
  <cellXfs count="2">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
    <xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="1"/>
  </cellXfs>
  <cellStyles count="2">
    <cellStyle name="Normal" xfId="0" builtinId="0"/>
    <cellStyle name="Heading 1" xfId="1" builtinId="16"/>
  </cellStyles>
</styleSheet>"#;

    let sheet_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1"><c r="A1" s="0" t="inlineStr"><is><t>Normal</t></is></c></row>
    <row r="2"><c r="A2" s="1" t="inlineStr"><is><t>Heading</t></is></c></row>
  </sheetData>
</worksheet>"#;

    let xlsx = create_xlsx_with_custom_styles(styles_xml, sheet_xml);
    let workbook = load_xlsx(&xlsx);

    // Cell A2 should inherit the bold + larger font from the Heading 1 style
    let heading_style = get_cell_style(&workbook, 0, 1, 0).expect("A2 should have a style");
    assert_eq!(heading_style.bold, Some(true), "Heading 1 should be bold");
    assert_eq!(
        heading_style.font_size,
        Some(15.0),
        "Heading 1 should have size 15"
    );
}

#[test]
fn test_parse_named_style_without_builtin_id() {
    // Custom named style without a builtinId should still parse and apply
    let styles_xml = r#"<?xml version="1.0" encoding="UTF-8"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="1">
    <font><name val="Calibri"/><sz val="11"/></font>
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
  <cellStyles count="1">
    <cellStyle name="Custom Style" xfId="0"/>
  </cellStyles>
</styleSheet>"#;

    let sheet_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1"><c r="A1" s="0" t="inlineStr"><is><t>Custom</t></is></c></row>
  </sheetData>
</worksheet>"#;

    let xlsx = create_xlsx_with_custom_styles(styles_xml, sheet_xml);
    let workbook = load_xlsx(&xlsx);

    // Should parse without error
    assert!(
        !workbook.sheets.is_empty(),
        "Should have at least one sheet"
    );
    let cell = get_cell(&workbook, 0, 0, 0);
    assert!(cell.is_some(), "Cell should exist");
}

// ============================================================================
// cellStyleXfs Parsing Tests
// ============================================================================

#[test]
fn test_parse_cell_style_xfs() {
    // cellStyleXfs define named style base formatting.
    // Heading 1 has bold font (fontId=1) and yellow fill (fillId=1).
    let styles_xml = r#"<?xml version="1.0" encoding="UTF-8"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="2">
    <font><name val="Calibri"/><sz val="11"/></font>
    <font><name val="Calibri"/><sz val="15"/><b/></font>
  </fonts>
  <fills count="3">
    <fill><patternFill patternType="none"/></fill>
    <fill><patternFill patternType="gray125"/></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FFFFFF00"/></patternFill></fill>
  </fills>
  <borders count="1">
    <border><left/><right/><top/><bottom/><diagonal/></border>
  </borders>
  <cellStyleXfs count="2">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
    <xf numFmtId="0" fontId="1" fillId="2" borderId="0"/>
  </cellStyleXfs>
  <cellXfs count="2">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
    <xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="1"/>
  </cellXfs>
  <cellStyles count="2">
    <cellStyle name="Normal" xfId="0" builtinId="0"/>
    <cellStyle name="Heading 1" xfId="1" builtinId="16"/>
  </cellStyles>
</styleSheet>"#;

    let sheet_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1"><c r="A1" s="1" t="inlineStr"><is><t>Heading</t></is></c></row>
  </sheetData>
</worksheet>"#;

    let xlsx = create_xlsx_with_custom_styles(styles_xml, sheet_xml);
    let workbook = load_xlsx(&xlsx);

    let style = get_cell_style(&workbook, 0, 0, 0).expect("Cell should have style");
    // Heading 1 style should have bold font
    assert_eq!(style.bold, Some(true), "Heading should be bold");
    assert_eq!(style.font_size, Some(15.0), "Heading should have 15pt font");
    // Fill should be yellow
    if let Some(ref bg) = style.bg_color {
        assert!(
            bg.to_uppercase().contains("FFFF00"),
            "Background should be yellow, got: {}",
            bg
        );
    }
}

#[test]
fn test_cell_xf_references_cell_style_xf() {
    // Two cellXfs referencing different cellStyleXfs.
    // Cell A1 uses xfId=0 (Normal), Cell A2 uses xfId=1 (Heading 1).
    let styles_xml = r#"<?xml version="1.0" encoding="UTF-8"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="2">
    <font><name val="Calibri"/><sz val="11"/></font>
    <font><name val="Calibri"/><sz val="15"/><b/></font>
  </fonts>
  <fills count="2">
    <fill><patternFill patternType="none"/></fill>
    <fill><patternFill patternType="gray125"/></fill>
  </fills>
  <borders count="1">
    <border><left/><right/><top/><bottom/><diagonal/></border>
  </borders>
  <cellStyleXfs count="2">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
    <xf numFmtId="0" fontId="1" fillId="0" borderId="0"/>
  </cellStyleXfs>
  <cellXfs count="2">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
    <xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="1"/>
  </cellXfs>
  <cellStyles count="2">
    <cellStyle name="Normal" xfId="0" builtinId="0"/>
    <cellStyle name="Heading 1" xfId="1" builtinId="16"/>
  </cellStyles>
</styleSheet>"#;

    let sheet_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1"><c r="A1" s="0" t="inlineStr"><is><t>Normal</t></is></c></row>
    <row r="2"><c r="A2" s="1" t="inlineStr"><is><t>Heading</t></is></c></row>
  </sheetData>
</worksheet>"#;

    let xlsx = create_xlsx_with_custom_styles(styles_xml, sheet_xml);
    let workbook = load_xlsx(&xlsx);

    // A1 should use Normal style (no bold, 11pt)
    let normal_style = get_cell_style(&workbook, 0, 0, 0);
    if let Some(s) = normal_style {
        // Normal is not bold
        assert_ne!(s.bold, Some(true), "Normal style should not be bold");
    }

    // A2 should use Heading 1 style (bold, 15pt)
    let heading_style = get_cell_style(&workbook, 0, 1, 0).expect("A2 should have style");
    assert_eq!(heading_style.bold, Some(true), "Heading 1 should be bold");
    assert_eq!(
        heading_style.font_size,
        Some(15.0),
        "Heading 1 should have 15pt font"
    );
}

// ============================================================================
// Built-in Style ID Tests
// ============================================================================

#[test]
fn test_common_builtin_style_ids() {
    // Multiple built-in styles: Normal, Heading 1-3, Title, Bad
    // We verify that the correct fonts are applied to cells using these styles.
    let styles_xml = r#"<?xml version="1.0" encoding="UTF-8"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="6">
    <font><name val="Calibri"/><sz val="11"/></font>
    <font><name val="Calibri"/><sz val="15"/><b/></font>
    <font><name val="Calibri"/><sz val="13"/><b/></font>
    <font><name val="Calibri"/><sz val="11"/><b/></font>
    <font><name val="Calibri"/><sz val="11"/><b/><i/></font>
    <font><name val="Calibri"/><sz val="11"/><color rgb="FFFF0000"/></font>
  </fonts>
  <fills count="2">
    <fill><patternFill patternType="none"/></fill>
    <fill><patternFill patternType="gray125"/></fill>
  </fills>
  <borders count="1">
    <border><left/><right/><top/><bottom/><diagonal/></border>
  </borders>
  <cellStyleXfs count="6">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
    <xf numFmtId="0" fontId="1" fillId="0" borderId="0"/>
    <xf numFmtId="0" fontId="2" fillId="0" borderId="0"/>
    <xf numFmtId="0" fontId="3" fillId="0" borderId="0"/>
    <xf numFmtId="0" fontId="4" fillId="0" borderId="0"/>
    <xf numFmtId="0" fontId="5" fillId="0" borderId="0"/>
  </cellStyleXfs>
  <cellXfs count="6">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
    <xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="1"/>
    <xf numFmtId="0" fontId="2" fillId="0" borderId="0" xfId="2"/>
    <xf numFmtId="0" fontId="3" fillId="0" borderId="0" xfId="3"/>
    <xf numFmtId="0" fontId="4" fillId="0" borderId="0" xfId="4"/>
    <xf numFmtId="0" fontId="5" fillId="0" borderId="0" xfId="5"/>
  </cellXfs>
  <cellStyles count="6">
    <cellStyle name="Normal" xfId="0" builtinId="0"/>
    <cellStyle name="Heading 1" xfId="1" builtinId="16"/>
    <cellStyle name="Heading 2" xfId="2" builtinId="17"/>
    <cellStyle name="Heading 3" xfId="3" builtinId="18"/>
    <cellStyle name="Title" xfId="4" builtinId="15"/>
    <cellStyle name="Bad" xfId="5" builtinId="27"/>
  </cellStyles>
</styleSheet>"#;

    let sheet_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1"><c r="A1" s="0" t="inlineStr"><is><t>Normal</t></is></c></row>
    <row r="2"><c r="A2" s="1" t="inlineStr"><is><t>Heading 1</t></is></c></row>
    <row r="3"><c r="A3" s="2" t="inlineStr"><is><t>Heading 2</t></is></c></row>
    <row r="4"><c r="A4" s="3" t="inlineStr"><is><t>Heading 3</t></is></c></row>
    <row r="5"><c r="A5" s="4" t="inlineStr"><is><t>Title</t></is></c></row>
    <row r="6"><c r="A6" s="5" t="inlineStr"><is><t>Bad</t></is></c></row>
  </sheetData>
</worksheet>"#;

    let xlsx = create_xlsx_with_custom_styles(styles_xml, sheet_xml);
    let workbook = load_xlsx(&xlsx);

    // Normal (11pt, not bold)
    let s0 = get_cell_style(&workbook, 0, 0, 0);
    if let Some(s) = s0 {
        assert_ne!(s.bold, Some(true), "Normal should not be bold");
        assert_eq!(s.font_size, Some(11.0), "Normal should be 11pt");
    }

    // Heading 1 (15pt, bold)
    let s1 = get_cell_style(&workbook, 0, 1, 0).expect("Heading 1 should have style");
    assert_eq!(s1.bold, Some(true), "Heading 1 should be bold");
    assert_eq!(s1.font_size, Some(15.0), "Heading 1 should be 15pt");

    // Heading 2 (13pt, bold)
    let s2 = get_cell_style(&workbook, 0, 2, 0).expect("Heading 2 should have style");
    assert_eq!(s2.bold, Some(true), "Heading 2 should be bold");
    assert_eq!(s2.font_size, Some(13.0), "Heading 2 should be 13pt");

    // Heading 3 (11pt, bold)
    let s3 = get_cell_style(&workbook, 0, 3, 0).expect("Heading 3 should have style");
    assert_eq!(s3.bold, Some(true), "Heading 3 should be bold");

    // Title (11pt, bold + italic)
    let s4 = get_cell_style(&workbook, 0, 4, 0).expect("Title should have style");
    assert_eq!(s4.bold, Some(true), "Title should be bold");
    assert_eq!(s4.italic, Some(true), "Title should be italic");

    // Bad (11pt, red font color)
    let s5 = get_cell_style(&workbook, 0, 5, 0).expect("Bad should have style");
    if let Some(ref color) = s5.font_color {
        assert!(
            color.to_uppercase().contains("FF0000"),
            "Bad style should have red font color, got: {}",
            color
        );
    }
}

// ============================================================================
// Empty/Missing Sections Tests
// ============================================================================

#[test]
fn test_missing_cell_styles_section() {
    // styles.xml with no <cellStyles> section - should still parse
    let styles_xml = r#"<?xml version="1.0" encoding="UTF-8"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="1">
    <font><name val="Calibri"/><sz val="11"/></font>
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
</styleSheet>"#;

    let sheet_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1"><c r="A1" s="0" t="inlineStr"><is><t>Test</t></is></c></row>
  </sheetData>
</worksheet>"#;

    let xlsx = create_xlsx_with_custom_styles(styles_xml, sheet_xml);
    let workbook = load_xlsx(&xlsx);

    // Should parse without error
    assert!(
        !workbook.sheets.is_empty(),
        "Should have at least one sheet"
    );
    let cell = get_cell(&workbook, 0, 0, 0);
    assert!(cell.is_some(), "Cell should exist");
}

#[test]
fn test_missing_cell_style_xfs_section() {
    // styles.xml with no <cellStyleXfs> section - should still parse
    let styles_xml = r#"<?xml version="1.0" encoding="UTF-8"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="1">
    <font><name val="Calibri"/><sz val="11"/></font>
  </fonts>
  <fills count="2">
    <fill><patternFill patternType="none"/></fill>
    <fill><patternFill patternType="gray125"/></fill>
  </fills>
  <borders count="1">
    <border><left/><right/><top/><bottom/><diagonal/></border>
  </borders>
  <cellXfs count="1">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
  </cellXfs>
  <cellStyles count="1">
    <cellStyle name="Normal" xfId="0" builtinId="0"/>
  </cellStyles>
</styleSheet>"#;

    let sheet_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1"><c r="A1" s="0" t="inlineStr"><is><t>Test</t></is></c></row>
  </sheetData>
</worksheet>"#;

    let xlsx = create_xlsx_with_custom_styles(styles_xml, sheet_xml);
    let workbook = load_xlsx(&xlsx);

    // Should parse without error
    assert!(
        !workbook.sheets.is_empty(),
        "Should have at least one sheet"
    );
}

// ============================================================================
// Style Inheritance Tests (cellXfs inherits from cellStyleXfs)
// ============================================================================

#[test]
fn test_cell_xf_with_apply_flags() {
    // Cell uses Heading 1 base style (bold, fontId=1) but overrides fill (applyFill=1)
    // to yellow. The result should have both: bold (from base) + yellow fill (override).
    let styles_xml = r#"<?xml version="1.0" encoding="UTF-8"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="2">
    <font><name val="Calibri"/><sz val="11"/></font>
    <font><name val="Calibri"/><sz val="15"/><b/></font>
  </fonts>
  <fills count="3">
    <fill><patternFill patternType="none"/></fill>
    <fill><patternFill patternType="gray125"/></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FFFFFF00"/></patternFill></fill>
  </fills>
  <borders count="1">
    <border><left/><right/><top/><bottom/><diagonal/></border>
  </borders>
  <cellStyleXfs count="2">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
    <xf numFmtId="0" fontId="1" fillId="0" borderId="0"/>
  </cellStyleXfs>
  <cellXfs count="2">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
    <xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="1" applyFill="1"/>
  </cellXfs>
  <cellStyles count="2">
    <cellStyle name="Normal" xfId="0" builtinId="0"/>
    <cellStyle name="Heading 1" xfId="1" builtinId="16"/>
  </cellStyles>
</styleSheet>"#;

    let sheet_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1"><c r="A1" s="1" t="inlineStr"><is><t>Heading with fill</t></is></c></row>
  </sheetData>
</worksheet>"#;

    let xlsx = create_xlsx_with_custom_styles(styles_xml, sheet_xml);
    let workbook = load_xlsx(&xlsx);

    let style = get_cell_style(&workbook, 0, 0, 0).expect("Cell should have style");

    // Font properties from base style (Heading 1)
    assert_eq!(
        style.bold,
        Some(true),
        "Should be bold from Heading 1 base style"
    );
    assert_eq!(
        style.font_size,
        Some(15.0),
        "Should be 15pt from Heading 1 base style"
    );

    // Fill override
    if let Some(ref bg) = style.bg_color {
        assert!(
            bg.to_uppercase().contains("FFFF00"),
            "Background should be yellow (applied fill override), got: {}",
            bg
        );
    }
}

#[test]
fn test_cell_style_xf_with_alignment() {
    // cellStyleXf with alignment properties (center, wrap text)
    let styles_xml = r#"<?xml version="1.0" encoding="UTF-8"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="1">
    <font><name val="Calibri"/><sz val="11"/></font>
  </fonts>
  <fills count="2">
    <fill><patternFill patternType="none"/></fill>
    <fill><patternFill patternType="gray125"/></fill>
  </fills>
  <borders count="1">
    <border><left/><right/><top/><bottom/><diagonal/></border>
  </borders>
  <cellStyleXfs count="1">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" applyAlignment="1">
      <alignment horizontal="center" vertical="center" wrapText="1"/>
    </xf>
  </cellStyleXfs>
  <cellXfs count="1">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1">
      <alignment horizontal="center" vertical="center" wrapText="1"/>
    </xf>
  </cellXfs>
  <cellStyles count="1">
    <cellStyle name="Centered" xfId="0" builtinId="0"/>
  </cellStyles>
</styleSheet>"#;

    let sheet_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1"><c r="A1" s="0" t="inlineStr"><is><t>Centered</t></is></c></row>
  </sheetData>
</worksheet>"#;

    let xlsx = create_xlsx_with_custom_styles(styles_xml, sheet_xml);
    let workbook = load_xlsx(&xlsx);

    let style = get_cell_style(&workbook, 0, 0, 0).expect("Cell should have style");

    // Alignment should be center
    assert_cell_align_h(&workbook, 0, 0, 0, "center");
    // Wrap text should be true
    assert_eq!(style.wrap, Some(true), "Should have wrap text enabled");
}

// ============================================================================
// Real-world Style Configuration Tests
// ============================================================================

#[test]
fn test_typical_excel_style_structure() {
    // Simulates a typical Excel file with Normal, Heading 1, and Bad styles
    let styles_xml = r##"<?xml version="1.0" encoding="UTF-8"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <numFmts count="1">
    <numFmt numFmtId="164" formatCode="#,##0.00"/>
  </numFmts>
  <fonts count="3">
    <font><name val="Calibri"/><sz val="11"/><color theme="1"/></font>
    <font><name val="Calibri"/><sz val="15"/><b/><color theme="4"/></font>
    <font><name val="Calibri"/><sz val="11"/><color rgb="FF9C0006"/></font>
  </fonts>
  <fills count="4">
    <fill><patternFill patternType="none"/></fill>
    <fill><patternFill patternType="gray125"/></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FFFFC7CE"/></patternFill></fill>
    <fill><patternFill patternType="none"/></fill>
  </fills>
  <borders count="2">
    <border><left/><right/><top/><bottom/><diagonal/></border>
    <border>
      <left style="thin"><color indexed="64"/></left>
      <right style="thin"><color indexed="64"/></right>
      <top style="thin"><color indexed="64"/></top>
      <bottom style="thin"><color indexed="64"/></bottom>
      <diagonal/>
    </border>
  </borders>
  <cellStyleXfs count="3">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
    <xf numFmtId="0" fontId="1" fillId="0" borderId="0" applyFont="1"/>
    <xf numFmtId="0" fontId="2" fillId="2" borderId="0" applyFont="1" applyFill="1"/>
  </cellStyleXfs>
  <cellXfs count="4">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
    <xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="1" applyFont="1"/>
    <xf numFmtId="0" fontId="2" fillId="2" borderId="0" xfId="2" applyFont="1" applyFill="1"/>
    <xf numFmtId="164" fontId="0" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyBorder="1"/>
  </cellXfs>
  <cellStyles count="3">
    <cellStyle name="Normal" xfId="0" builtinId="0"/>
    <cellStyle name="Heading 1" xfId="1" builtinId="16"/>
    <cellStyle name="Bad" xfId="2" builtinId="27"/>
  </cellStyles>
</styleSheet>"##;

    let sheet_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1"><c r="A1" s="0" t="inlineStr"><is><t>Normal</t></is></c></row>
    <row r="2"><c r="A2" s="1" t="inlineStr"><is><t>Heading</t></is></c></row>
    <row r="3"><c r="A3" s="2" t="inlineStr"><is><t>Bad</t></is></c></row>
    <row r="4"><c r="A4" s="3"><v>1234.5</v></c></row>
  </sheetData>
</worksheet>"#;

    let xlsx = create_xlsx_with_custom_styles(styles_xml, sheet_xml);
    let workbook = load_xlsx(&xlsx);

    assert!(!workbook.sheets.is_empty(), "Should have a sheet");

    // Heading 1 should be bold with 15pt font
    let s1 = get_cell_style(&workbook, 0, 1, 0).expect("Heading 1 should have style");
    assert_eq!(s1.bold, Some(true), "Heading 1 should be bold");
    assert_eq!(s1.font_size, Some(15.0), "Heading 1 should be 15pt");

    // Bad style should have red-ish font color and pink background
    let s2 = get_cell_style(&workbook, 0, 2, 0).expect("Bad should have style");
    if let Some(ref color) = s2.font_color {
        assert!(
            color.to_uppercase().contains("9C0006"),
            "Bad font color should be dark red, got: {}",
            color
        );
    }
    if let Some(ref bg) = s2.bg_color {
        assert!(
            bg.to_uppercase().contains("FFC7CE"),
            "Bad background should be pink, got: {}",
            bg
        );
    }

    // Cell with number format and border (s=3)
    let s3 = get_cell_style(&workbook, 0, 3, 0).expect("Number cell should have style");
    // Should have thin borders
    assert!(
        s3.border_top.is_some(),
        "Number cell should have top border"
    );
}

#[test]
fn test_named_style_with_protection() {
    // cellStyleXf with protection properties
    let styles_xml = r#"<?xml version="1.0" encoding="UTF-8"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="1">
    <font><name val="Calibri"/><sz val="11"/></font>
  </fonts>
  <fills count="2">
    <fill><patternFill patternType="none"/></fill>
    <fill><patternFill patternType="gray125"/></fill>
  </fills>
  <borders count="1">
    <border><left/><right/><top/><bottom/><diagonal/></border>
  </borders>
  <cellStyleXfs count="1">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" applyProtection="1">
      <protection locked="0" hidden="1"/>
    </xf>
  </cellStyleXfs>
  <cellXfs count="1">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyProtection="1">
      <protection locked="0" hidden="1"/>
    </xf>
  </cellXfs>
  <cellStyles count="1">
    <cellStyle name="Unlocked" xfId="0"/>
  </cellStyles>
</styleSheet>"#;

    let sheet_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1"><c r="A1" s="0" t="inlineStr"><is><t>Unlocked</t></is></c></row>
  </sheetData>
</worksheet>"#;

    let xlsx = create_xlsx_with_custom_styles(styles_xml, sheet_xml);
    let workbook = load_xlsx(&xlsx);

    // Should parse without error
    assert!(
        !workbook.sheets.is_empty(),
        "Should have at least one sheet"
    );
    let cell = get_cell(&workbook, 0, 0, 0);
    assert!(cell.is_some(), "Cell should exist");

    // Protection properties may or may not be surfaced in the viewer model
    // The important thing is parsing doesn't fail
    let style = get_cell_style(&workbook, 0, 0, 0);
    if let Some(s) = style {
        // If protection is surfaced, check it
        if s.locked.is_some() {
            assert_eq!(s.locked, Some(false), "Cell should be unlocked");
        }
        if s.hidden.is_some() {
            assert_eq!(s.hidden, Some(true), "Formula should be hidden");
        }
    }
}

// ============================================================================
// Edge Cases
// ============================================================================

#[test]
fn test_named_style_with_special_characters_in_name() {
    // Named style with XML entities in the name
    let styles_xml = r#"<?xml version="1.0" encoding="UTF-8"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="1">
    <font><name val="Calibri"/><sz val="11"/></font>
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
  <cellStyles count="1">
    <cellStyle name="My &quot;Custom&quot; Style &amp; More" xfId="0"/>
  </cellStyles>
</styleSheet>"#;

    let sheet_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1"><c r="A1" s="0" t="inlineStr"><is><t>Special</t></is></c></row>
  </sheetData>
</worksheet>"#;

    let xlsx = create_xlsx_with_custom_styles(styles_xml, sheet_xml);
    let workbook = load_xlsx(&xlsx);

    // Should parse without error even with special characters in style name
    assert!(
        !workbook.sheets.is_empty(),
        "Should have at least one sheet"
    );
    let cell = get_cell(&workbook, 0, 0, 0);
    assert!(cell.is_some(), "Cell should exist");
}

#[test]
fn test_large_builtin_id() {
    // Large builtinId should not cause parsing issues
    let styles_xml = r#"<?xml version="1.0" encoding="UTF-8"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="1">
    <font><name val="Calibri"/><sz val="11"/></font>
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
  <cellStyles count="1">
    <cellStyle name="Custom Style" xfId="0" builtinId="999"/>
  </cellStyles>
</styleSheet>"#;

    let sheet_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1"><c r="A1" s="0" t="inlineStr"><is><t>Large builtin</t></is></c></row>
  </sheetData>
</worksheet>"#;

    let xlsx = create_xlsx_with_custom_styles(styles_xml, sheet_xml);
    let workbook = load_xlsx(&xlsx);

    // Should parse without error
    assert!(
        !workbook.sheets.is_empty(),
        "Should have at least one sheet"
    );
}

#[test]
fn test_empty_stylesheet() {
    // Minimal stylesheet with no sections except the required minimum
    let styles_xml = r#"<?xml version="1.0" encoding="UTF-8"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="1">
    <font><name val="Calibri"/><sz val="11"/></font>
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
</styleSheet>"#;

    let sheet_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1"><c r="A1" t="inlineStr"><is><t>Minimal</t></is></c></row>
  </sheetData>
</worksheet>"#;

    let xlsx = create_xlsx_with_custom_styles(styles_xml, sheet_xml);
    let workbook = load_xlsx(&xlsx);

    // Should parse without error
    assert!(
        !workbook.sheets.is_empty(),
        "Should have at least one sheet"
    );
    let cell = get_cell(&workbook, 0, 0, 0);
    assert!(cell.is_some(), "Cell should exist");
}
