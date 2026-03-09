//! Tests for gradient fill parsing in XLSX files, ported from xlview.
//!
//! Gradient fills in XLSX are defined in styles.xml under the fills section
//! using the `gradientFill` element instead of `patternFill`.
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

// =============================================================================
// Helper: Create XLSX with gradient fill
// =============================================================================

fn create_gradient_xlsx(gradient_xml: &str) -> Vec<u8> {
    let cursor = Cursor::new(Vec::new());
    let mut zip = ZipWriter::new(cursor);
    let options = FileOptions::<()>::default().compression_method(zip::CompressionMethod::Deflated);

    let _ = zip.start_file("[Content_Types].xml", options);
    let _ = zip.write_all(
        br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>
<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
<Override PartName="/xl/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>
</Types>"#,
    );

    let _ = zip.start_file("_rels/.rels", options);
    let _ = zip.write_all(
        br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>"#,
    );

    let _ = zip.start_file("xl/_rels/workbook.xml.rels", options);
    let _ = zip.write_all(
        br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>
<Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>
</Relationships>"#,
    );

    let _ = zip.start_file("xl/workbook.xml", options);
    let _ = zip.write_all(
        br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets>
</workbook>"#,
    );

    let _ = zip.start_file("xl/styles.xml", options);
    let styles_xml = format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>
  <fills count="3">
    <fill><patternFill patternType="none"/></fill>
    <fill><patternFill patternType="gray125"/></fill>
    <fill>{gradient_xml}</fill>
  </fills>
  <borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>
  <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
  <cellXfs count="2">
    <xf fontId="0" fillId="0" borderId="0"/>
    <xf fontId="0" fillId="2" borderId="0" applyFill="1"/>
  </cellXfs>
  <cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
</styleSheet>"#
    );
    let _ = zip.write_all(styles_xml.as_bytes());

    let _ = zip.start_file("xl/sharedStrings.xml", options);
    let _ = zip.write_all(
        br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="1" uniqueCount="1">
<si><t>Gradient Test</t></si>
</sst>"#,
    );

    let _ = zip.start_file("xl/theme/theme1.xml", options);
    let _ = zip.write_all(
        br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
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
<a:fontScheme name="Office"><a:majorFont><a:latin typeface="Calibri Light"/></a:majorFont><a:minorFont><a:latin typeface="Calibri"/></a:minorFont></a:fontScheme>
<a:fmtScheme name="Office"><a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:fillStyleLst></a:fmtScheme>
</a:themeElements>
</a:theme>"#,
    );

    let _ = zip.start_file("xl/worksheets/sheet1.xml", options);
    let _ = zip.write_all(
        br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<sheetData>
<row r="1"><c r="A1" t="s" s="1"><v>0</v></c></row>
</sheetData>
</worksheet>"#,
    );

    let cursor = zip.finish().expect("Failed to finish ZIP");
    cursor.into_inner()
}

// =============================================================================
// Linear Gradient Tests
// =============================================================================

#[test]
fn test_linear_gradient_horizontal() {
    let gradient_xml = r#"<gradientFill type="linear" degree="0">
        <stop position="0"><color rgb="FFFF0000"/></stop>
        <stop position="1"><color rgb="FF0000FF"/></stop>
    </gradientFill>"#;
    let xlsx = create_gradient_xlsx(gradient_xml);
    let workbook = load_xlsx(&xlsx);
    let style = get_cell_style(&workbook, 0, 0, 0).expect("Cell should have style");
    let gradient = style
        .gradient
        .as_ref()
        .expect("Style should have a gradient");
    assert_eq!(gradient.gradient_type, "linear");
    assert_eq!(gradient.degree, Some(0.0));
    assert_eq!(gradient.stops.len(), 2);
    assert_eq!(gradient.stops[0].position, 0.0);
    assert_eq!(gradient.stops[0].color, "#FF0000");
    assert_eq!(gradient.stops[1].position, 1.0);
    assert_eq!(gradient.stops[1].color, "#0000FF");
}

#[test]
fn test_linear_gradient_vertical() {
    let gradient_xml = r#"<gradientFill type="linear" degree="90">
        <stop position="0"><color rgb="FF00FF00"/></stop>
        <stop position="1"><color rgb="FFFFFF00"/></stop>
    </gradientFill>"#;
    let xlsx = create_gradient_xlsx(gradient_xml);
    let workbook = load_xlsx(&xlsx);
    let style = get_cell_style(&workbook, 0, 0, 0).expect("Cell should have style");
    let gradient = style
        .gradient
        .as_ref()
        .expect("Style should have a gradient");
    assert_eq!(gradient.gradient_type, "linear");
    assert_eq!(gradient.degree, Some(90.0));
    assert_eq!(gradient.stops.len(), 2);
    assert_eq!(gradient.stops[0].color, "#00FF00");
    assert_eq!(gradient.stops[1].color, "#FFFF00");
}

#[test]
fn test_linear_gradient_diagonal() {
    let gradient_xml = r#"<gradientFill type="linear" degree="45">
        <stop position="0"><color rgb="FFFF00FF"/></stop>
        <stop position="1"><color rgb="FF00FFFF"/></stop>
    </gradientFill>"#;
    let xlsx = create_gradient_xlsx(gradient_xml);
    let workbook = load_xlsx(&xlsx);
    let style = get_cell_style(&workbook, 0, 0, 0).expect("Cell should have style");
    let gradient = style
        .gradient
        .as_ref()
        .expect("Style should have a gradient");
    assert_eq!(gradient.gradient_type, "linear");
    assert_eq!(gradient.degree, Some(45.0));
    assert_eq!(gradient.stops.len(), 2);
}

#[test]
fn test_linear_gradient_no_type_defaults_to_linear() {
    let gradient_xml = r#"<gradientFill degree="90">
        <stop position="0"><color rgb="FFFF0000"/></stop>
        <stop position="1"><color rgb="FF0000FF"/></stop>
    </gradientFill>"#;
    let xlsx = create_gradient_xlsx(gradient_xml);
    let workbook = load_xlsx(&xlsx);
    let style = get_cell_style(&workbook, 0, 0, 0).expect("Cell should have style");
    let gradient = style
        .gradient
        .as_ref()
        .expect("Style should have a gradient");
    assert_eq!(gradient.gradient_type, "linear");
}

// =============================================================================
// Path (Radial) Gradient Tests
// =============================================================================

#[test]
fn test_path_gradient_center() {
    let gradient_xml = r#"<gradientFill type="path" left="0.5" right="0.5" top="0.5" bottom="0.5">
        <stop position="0"><color rgb="FFFFFFFF"/></stop>
        <stop position="1"><color rgb="FF000000"/></stop>
    </gradientFill>"#;
    let xlsx = create_gradient_xlsx(gradient_xml);
    let workbook = load_xlsx(&xlsx);
    let style = get_cell_style(&workbook, 0, 0, 0).expect("Cell should have style");
    let gradient = style
        .gradient
        .as_ref()
        .expect("Style should have a gradient");
    assert_eq!(gradient.gradient_type, "path");
    assert_eq!(gradient.left, Some(0.5));
    assert_eq!(gradient.right, Some(0.5));
    assert_eq!(gradient.top, Some(0.5));
    assert_eq!(gradient.bottom, Some(0.5));
    assert_eq!(gradient.stops.len(), 2);
    assert_eq!(gradient.stops[0].color, "#FFFFFF");
    assert_eq!(gradient.stops[1].color, "#000000");
}

#[test]
fn test_path_gradient_corner() {
    let gradient_xml = r#"<gradientFill type="path" left="0" right="1" top="0" bottom="1">
        <stop position="0"><color rgb="FF4472C4"/></stop>
        <stop position="1"><color rgb="FFED7D31"/></stop>
    </gradientFill>"#;
    let xlsx = create_gradient_xlsx(gradient_xml);
    let workbook = load_xlsx(&xlsx);
    let style = get_cell_style(&workbook, 0, 0, 0).expect("Cell should have style");
    let gradient = style
        .gradient
        .as_ref()
        .expect("Style should have a gradient");
    assert_eq!(gradient.gradient_type, "path");
    assert_eq!(gradient.left, Some(0.0));
    assert_eq!(gradient.right, Some(1.0));
}

// =============================================================================
// Multiple Color Stops Tests
// =============================================================================

#[test]
fn test_gradient_three_stops() {
    let gradient_xml = r#"<gradientFill type="linear" degree="0">
        <stop position="0"><color rgb="FFFF0000"/></stop>
        <stop position="0.5"><color rgb="FFFFFF00"/></stop>
        <stop position="1"><color rgb="FF00FF00"/></stop>
    </gradientFill>"#;
    let xlsx = create_gradient_xlsx(gradient_xml);
    let workbook = load_xlsx(&xlsx);
    let style = get_cell_style(&workbook, 0, 0, 0).expect("Cell should have style");
    let gradient = style
        .gradient
        .as_ref()
        .expect("Style should have a gradient");
    assert_eq!(gradient.stops.len(), 3);
    assert_eq!(gradient.stops[0].color, "#FF0000");
    assert_eq!(gradient.stops[1].color, "#FFFF00");
    assert_eq!(gradient.stops[2].color, "#00FF00");
}

#[test]
fn test_gradient_four_stops() {
    let gradient_xml = r#"<gradientFill type="linear" degree="90">
        <stop position="0"><color rgb="FFFF0000"/></stop>
        <stop position="0.33"><color rgb="FFFFFF00"/></stop>
        <stop position="0.66"><color rgb="FF00FF00"/></stop>
        <stop position="1"><color rgb="FF0000FF"/></stop>
    </gradientFill>"#;
    let xlsx = create_gradient_xlsx(gradient_xml);
    let workbook = load_xlsx(&xlsx);
    let style = get_cell_style(&workbook, 0, 0, 0).expect("Cell should have style");
    let gradient = style
        .gradient
        .as_ref()
        .expect("Style should have a gradient");
    assert_eq!(gradient.stops.len(), 4);
    assert_eq!(gradient.stops[0].color, "#FF0000");
    assert_eq!(gradient.stops[1].color, "#FFFF00");
    assert_eq!(gradient.stops[2].color, "#00FF00");
    assert_eq!(gradient.stops[3].color, "#0000FF");
}

// =============================================================================
// Theme Color Tests
// =============================================================================

#[test]
fn test_gradient_with_theme_colors() {
    let gradient_xml = r#"<gradientFill type="linear" degree="0">
        <stop position="0"><color theme="4"/></stop>
        <stop position="1"><color theme="5"/></stop>
    </gradientFill>"#;
    let xlsx = create_gradient_xlsx(gradient_xml);
    let workbook = load_xlsx(&xlsx);
    let style = get_cell_style(&workbook, 0, 0, 0).expect("Cell should have style");
    let gradient = style
        .gradient
        .as_ref()
        .expect("Style should have a gradient");
    assert_eq!(gradient.stops.len(), 2);
    assert_eq!(gradient.stops[0].color, "#4472C4");
    assert_eq!(gradient.stops[1].color, "#ED7D31");
}

#[test]
fn test_gradient_with_theme_tint() {
    let gradient_xml = r#"<gradientFill type="linear" degree="0">
        <stop position="0"><color theme="4" tint="0.4"/></stop>
        <stop position="1"><color theme="4" tint="-0.25"/></stop>
    </gradientFill>"#;
    let xlsx = create_gradient_xlsx(gradient_xml);
    let workbook = load_xlsx(&xlsx);
    let style = get_cell_style(&workbook, 0, 0, 0).expect("Cell should have style");
    let gradient = style
        .gradient
        .as_ref()
        .expect("Style should have a gradient");
    assert_eq!(gradient.stops.len(), 2);
    assert!(gradient.stops[0].color.starts_with('#'));
    assert!(gradient.stops[1].color.starts_with('#'));
}

// =============================================================================
// Edge Cases
// =============================================================================

#[test]
fn test_gradient_no_degree_defaults() {
    let gradient_xml = r#"<gradientFill type="linear">
        <stop position="0"><color rgb="FFFF0000"/></stop>
        <stop position="1"><color rgb="FF0000FF"/></stop>
    </gradientFill>"#;
    let xlsx = create_gradient_xlsx(gradient_xml);
    let workbook = load_xlsx(&xlsx);
    let style = get_cell_style(&workbook, 0, 0, 0).expect("Cell should have style");
    let gradient = style
        .gradient
        .as_ref()
        .expect("Style should have a gradient");
    assert_eq!(gradient.gradient_type, "linear");
    assert_eq!(gradient.degree, None);
    assert_eq!(gradient.stops.len(), 2);
}

#[test]
fn test_gradient_fractional_positions() {
    let gradient_xml = r#"<gradientFill type="linear" degree="0">
        <stop position="0"><color rgb="FFFF0000"/></stop>
        <stop position="0.25"><color rgb="FFFFA500"/></stop>
        <stop position="0.75"><color rgb="FF00FF00"/></stop>
        <stop position="1"><color rgb="FF0000FF"/></stop>
    </gradientFill>"#;
    let xlsx = create_gradient_xlsx(gradient_xml);
    let workbook = load_xlsx(&xlsx);
    let style = get_cell_style(&workbook, 0, 0, 0).expect("Cell should have style");
    let gradient = style
        .gradient
        .as_ref()
        .expect("Style should have a gradient");
    assert_eq!(gradient.stops.len(), 4);
    assert!((gradient.stops[1].position - 0.25).abs() < 0.001);
    assert!((gradient.stops[2].position - 0.75).abs() < 0.001);
}

#[test]
fn test_gradient_decimal_degree() {
    let gradient_xml = r#"<gradientFill type="linear" degree="135.5">
        <stop position="0"><color rgb="FFFF0000"/></stop>
        <stop position="1"><color rgb="FF0000FF"/></stop>
    </gradientFill>"#;
    let xlsx = create_gradient_xlsx(gradient_xml);
    let workbook = load_xlsx(&xlsx);
    let style = get_cell_style(&workbook, 0, 0, 0).expect("Cell should have style");
    let gradient = style
        .gradient
        .as_ref()
        .expect("Style should have a gradient");
    assert_eq!(gradient.degree, Some(135.5));
}

// =============================================================================
// Pattern Fill vs Gradient Fill Tests
// =============================================================================

#[test]
fn test_cell_with_pattern_fill_has_no_gradient() {
    let xlsx = XlsxBuilder::new()
        .sheet(SheetBuilder::new("Sheet1").cell(
            "A1",
            "Pattern Fill",
            Some(StyleBuilder::new().bg_color("#FF0000").build()),
        ))
        .build();

    let workbook = load_xlsx(&xlsx);
    let style = get_cell_style(&workbook, 0, 0, 0).expect("Cell should have a style");

    assert_eq!(style.bg_color, Some("#FF0000".to_string()));
    assert!(style.gradient.is_none());
}
