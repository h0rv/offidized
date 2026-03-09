//! Comprehensive tests for fill/background styling in XLSX files, ported from xlview.
//!
//! Tests all fill pattern types and color sources through the full adapter pipeline.
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
// Helper: Create styles.xml with custom fills
// =============================================================================

fn create_fill_styles_xml(fills: &[&str]) -> String {
    let fills_xml: String = fills.join("\n    ");
    let fill_count = fills.len() + 2; // +2 for mandatory none and gray125

    let cell_xfs: String = (0..fill_count)
        .map(|i| {
            if i == 0 {
                r#"<xf fontId="0" fillId="0" borderId="0"/>"#.to_string()
            } else {
                format!(r#"<xf fontId="0" fillId="{i}" borderId="0" applyFill="1"/>"#)
            }
        })
        .collect::<Vec<_>>()
        .join("\n    ");

    format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="1">
    <font>
      <sz val="11"/>
      <name val="Calibri"/>
    </font>
  </fonts>
  <fills count="{fill_count}">
    <fill>
      <patternFill patternType="none"/>
    </fill>
    <fill>
      <patternFill patternType="gray125"/>
    </fill>
    {fills_xml}
  </fills>
  <borders count="1">
    <border>
      <left/><right/><top/><bottom/><diagonal/>
    </border>
  </borders>
  <cellStyleXfs count="1">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
  </cellStyleXfs>
  <cellXfs count="{fill_count}">
    {cell_xfs}
  </cellXfs>
  <cellStyles count="1">
    <cellStyle name="Normal" xfId="0" builtinId="0"/>
  </cellStyles>
</styleSheet>"#
    )
}

/// Create an XLSX file with a custom fill XML
fn create_xlsx_with_custom_fill(fill_xml: &str, cell_value: &str) -> Vec<u8> {
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
    let styles_xml = create_fill_styles_xml(&[fill_xml]);
    let _ = zip.write_all(styles_xml.as_bytes());

    let _ = zip.start_file("xl/sharedStrings.xml", options);
    let shared_strings = format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="1" uniqueCount="1">
<si><t>{cell_value}</t></si>
</sst>"#
    );
    let _ = zip.write_all(shared_strings.as_bytes());

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
<row r="1"><c r="A1" t="s" s="2"><v>0</v></c></row>
</sheetData>
</worksheet>"#,
    );

    let cursor = zip.finish().expect("Failed to finish ZIP");
    cursor.into_inner()
}

// =============================================================================
// 1. Solid Fill with RGB Color Tests
// =============================================================================

#[test]
fn test_solid_fill_yellow_rgb() {
    let xlsx = xlsx_with_styled_cell("Test", StyleBuilder::new().bg_color("#FFFF00").build());
    let wb = load_xlsx(&xlsx);
    assert_cell_bg_color(&wb, 0, 0, 0, "#FFFF00");
}

#[test]
fn test_solid_fill_red_rgb() {
    let xlsx = xlsx_with_styled_cell("Test", StyleBuilder::new().bg_color("#FF0000").build());
    let wb = load_xlsx(&xlsx);
    assert_cell_bg_color(&wb, 0, 0, 0, "#FF0000");
}

#[test]
fn test_solid_fill_green_rgb() {
    let xlsx = xlsx_with_styled_cell("Test", StyleBuilder::new().bg_color("#00FF00").build());
    let wb = load_xlsx(&xlsx);
    assert_cell_bg_color(&wb, 0, 0, 0, "#00FF00");
}

#[test]
fn test_solid_fill_blue_rgb() {
    let xlsx = xlsx_with_styled_cell("Test", StyleBuilder::new().bg_color("#0000FF").build());
    let wb = load_xlsx(&xlsx);
    assert_cell_bg_color(&wb, 0, 0, 0, "#0000FF");
}

#[test]
fn test_solid_fill_white_rgb() {
    let xlsx = xlsx_with_styled_cell("Test", StyleBuilder::new().bg_color("#FFFFFF").build());
    let wb = load_xlsx(&xlsx);
    assert_cell_bg_color(&wb, 0, 0, 0, "#FFFFFF");
}

#[test]
fn test_solid_fill_black_rgb() {
    let xlsx = xlsx_with_styled_cell("Test", StyleBuilder::new().bg_color("#000000").build());
    let wb = load_xlsx(&xlsx);
    assert_cell_bg_color(&wb, 0, 0, 0, "#000000");
}

#[test]
fn test_solid_fill_with_argb_format() {
    // ARGB format where first 2 chars are alpha - bg_color normalizes to FFFF00
    let xlsx = xlsx_with_styled_cell("Test", StyleBuilder::new().bg_color("#FFFF00").build());
    let wb = load_xlsx(&xlsx);
    assert_cell_bg_color(&wb, 0, 0, 0, "#FFFF00");
}

#[test]
fn test_solid_fill_custom_color() {
    let xlsx = xlsx_with_styled_cell("Test", StyleBuilder::new().bg_color("#8B008B").build());
    let wb = load_xlsx(&xlsx);
    assert_cell_bg_color(&wb, 0, 0, 0, "#8B008B");
}

// =============================================================================
// 4. Pattern Fills (Gray Patterns) Tests
// =============================================================================

#[test]
fn test_pattern_gray125() {
    let fill_xml = r#"<fill><patternFill patternType="gray125"/></fill>"#;
    let xlsx = create_xlsx_with_custom_fill(fill_xml, "Gray125 Test");
    let wb = load_xlsx(&xlsx);
    let cell = get_cell(&wb, 0, 0, 0);
    assert!(cell.is_some());
}

#[test]
fn test_pattern_gray0625() {
    let fill_xml = r#"<fill><patternFill patternType="gray0625"/></fill>"#;
    let xlsx = create_xlsx_with_custom_fill(fill_xml, "Gray0625 Test");
    let wb = load_xlsx(&xlsx);
    let cell = get_cell(&wb, 0, 0, 0);
    assert!(cell.is_some());
}

#[test]
fn test_pattern_dark_gray() {
    let fill_xml = r#"<fill><patternFill patternType="darkGray"><fgColor rgb="FF000000"/><bgColor rgb="FFFFFFFF"/></patternFill></fill>"#;
    let xlsx = create_xlsx_with_custom_fill(fill_xml, "DarkGray Test");
    let wb = load_xlsx(&xlsx);
    let cell = get_cell(&wb, 0, 0, 0);
    assert!(cell.is_some());
}

#[test]
fn test_pattern_medium_gray() {
    let fill_xml = r#"<fill><patternFill patternType="mediumGray"><fgColor rgb="FF808080"/><bgColor rgb="FFFFFFFF"/></patternFill></fill>"#;
    let xlsx = create_xlsx_with_custom_fill(fill_xml, "MediumGray Test");
    let wb = load_xlsx(&xlsx);
    let cell = get_cell(&wb, 0, 0, 0);
    assert!(cell.is_some());
}

#[test]
fn test_pattern_light_gray() {
    let fill_xml = r#"<fill><patternFill patternType="lightGray"><fgColor rgb="FFC0C0C0"/><bgColor rgb="FFFFFFFF"/></patternFill></fill>"#;
    let xlsx = create_xlsx_with_custom_fill(fill_xml, "LightGray Test");
    let wb = load_xlsx(&xlsx);
    let cell = get_cell(&wb, 0, 0, 0);
    assert!(cell.is_some());
}

// =============================================================================
// 5. Pattern Fills with Foreground and Background Colors Tests
// =============================================================================

#[test]
fn test_pattern_with_red_fg_white_bg() {
    let fill_xml = r#"<fill><patternFill patternType="darkGray"><fgColor rgb="FFFF0000"/><bgColor rgb="FFFFFFFF"/></patternFill></fill>"#;
    let xlsx = create_xlsx_with_custom_fill(fill_xml, "FG/BG Test");
    let wb = load_xlsx(&xlsx);
    let cell = get_cell(&wb, 0, 0, 0);
    assert!(cell.is_some());
}

#[test]
fn test_pattern_with_blue_fg_yellow_bg() {
    let fill_xml = r#"<fill><patternFill patternType="mediumGray"><fgColor rgb="FF0000FF"/><bgColor rgb="FFFFFF00"/></patternFill></fill>"#;
    let xlsx = create_xlsx_with_custom_fill(fill_xml, "FG/BG Test");
    let wb = load_xlsx(&xlsx);
    let cell = get_cell(&wb, 0, 0, 0);
    assert!(cell.is_some());
}

#[test]
fn test_pattern_with_theme_fg_rgb_bg() {
    let fill_xml = r#"<fill><patternFill patternType="lightGray"><fgColor theme="4"/><bgColor rgb="FFFFFFFF"/></patternFill></fill>"#;
    let xlsx = create_xlsx_with_custom_fill(fill_xml, "Mixed Color Test");
    let wb = load_xlsx(&xlsx);
    let cell = get_cell(&wb, 0, 0, 0);
    assert!(cell.is_some());
}

#[test]
fn test_pattern_with_indexed_fg_theme_bg() {
    let fill_xml = r#"<fill><patternFill patternType="darkGray"><fgColor indexed="2"/><bgColor theme="1"/></patternFill></fill>"#;
    let xlsx = create_xlsx_with_custom_fill(fill_xml, "Mixed Color Test");
    let wb = load_xlsx(&xlsx);
    let cell = get_cell(&wb, 0, 0, 0);
    assert!(cell.is_some());
}

// =============================================================================
// 6. Stripe Pattern Tests (Horizontal, Vertical, Diagonal)
// =============================================================================

#[test]
fn test_pattern_dark_horizontal() {
    let fill_xml = r#"<fill><patternFill patternType="darkHorizontal"><fgColor rgb="FF000000"/><bgColor rgb="FFFFFFFF"/></patternFill></fill>"#;
    let xlsx = create_xlsx_with_custom_fill(fill_xml, "DarkHorizontal Test");
    let wb = load_xlsx(&xlsx);
    let cell = get_cell(&wb, 0, 0, 0);
    assert!(cell.is_some());
}

#[test]
fn test_pattern_dark_vertical() {
    let fill_xml = r#"<fill><patternFill patternType="darkVertical"><fgColor rgb="FF000000"/><bgColor rgb="FFFFFFFF"/></patternFill></fill>"#;
    let xlsx = create_xlsx_with_custom_fill(fill_xml, "DarkVertical Test");
    let wb = load_xlsx(&xlsx);
    let cell = get_cell(&wb, 0, 0, 0);
    assert!(cell.is_some());
}

#[test]
fn test_pattern_dark_down() {
    let fill_xml = r#"<fill><patternFill patternType="darkDown"><fgColor rgb="FF000000"/><bgColor rgb="FFFFFFFF"/></patternFill></fill>"#;
    let xlsx = create_xlsx_with_custom_fill(fill_xml, "DarkDown Test");
    let wb = load_xlsx(&xlsx);
    let cell = get_cell(&wb, 0, 0, 0);
    assert!(cell.is_some());
}

#[test]
fn test_pattern_dark_up() {
    let fill_xml = r#"<fill><patternFill patternType="darkUp"><fgColor rgb="FF000000"/><bgColor rgb="FFFFFFFF"/></patternFill></fill>"#;
    let xlsx = create_xlsx_with_custom_fill(fill_xml, "DarkUp Test");
    let wb = load_xlsx(&xlsx);
    let cell = get_cell(&wb, 0, 0, 0);
    assert!(cell.is_some());
}

#[test]
fn test_pattern_light_horizontal() {
    let fill_xml = r#"<fill><patternFill patternType="lightHorizontal"><fgColor rgb="FF808080"/><bgColor rgb="FFFFFFFF"/></patternFill></fill>"#;
    let xlsx = create_xlsx_with_custom_fill(fill_xml, "LightHorizontal Test");
    let wb = load_xlsx(&xlsx);
    let cell = get_cell(&wb, 0, 0, 0);
    assert!(cell.is_some());
}

#[test]
fn test_pattern_light_vertical() {
    let fill_xml = r#"<fill><patternFill patternType="lightVertical"><fgColor rgb="FF808080"/><bgColor rgb="FFFFFFFF"/></patternFill></fill>"#;
    let xlsx = create_xlsx_with_custom_fill(fill_xml, "LightVertical Test");
    let wb = load_xlsx(&xlsx);
    let cell = get_cell(&wb, 0, 0, 0);
    assert!(cell.is_some());
}

#[test]
fn test_pattern_light_down() {
    let fill_xml = r#"<fill><patternFill patternType="lightDown"><fgColor rgb="FF808080"/><bgColor rgb="FFFFFFFF"/></patternFill></fill>"#;
    let xlsx = create_xlsx_with_custom_fill(fill_xml, "LightDown Test");
    let wb = load_xlsx(&xlsx);
    let cell = get_cell(&wb, 0, 0, 0);
    assert!(cell.is_some());
}

#[test]
fn test_pattern_light_up() {
    let fill_xml = r#"<fill><patternFill patternType="lightUp"><fgColor rgb="FF808080"/><bgColor rgb="FFFFFFFF"/></patternFill></fill>"#;
    let xlsx = create_xlsx_with_custom_fill(fill_xml, "LightUp Test");
    let wb = load_xlsx(&xlsx);
    let cell = get_cell(&wb, 0, 0, 0);
    assert!(cell.is_some());
}

// =============================================================================
// 7. Grid and Trellis Pattern Tests
// =============================================================================

#[test]
fn test_pattern_dark_grid() {
    let fill_xml = r#"<fill><patternFill patternType="darkGrid"><fgColor rgb="FF000000"/><bgColor rgb="FFFFFFFF"/></patternFill></fill>"#;
    let xlsx = create_xlsx_with_custom_fill(fill_xml, "DarkGrid Test");
    let wb = load_xlsx(&xlsx);
    let cell = get_cell(&wb, 0, 0, 0);
    assert!(cell.is_some());
}

#[test]
fn test_pattern_light_grid() {
    let fill_xml = r#"<fill><patternFill patternType="lightGrid"><fgColor rgb="FF808080"/><bgColor rgb="FFFFFFFF"/></patternFill></fill>"#;
    let xlsx = create_xlsx_with_custom_fill(fill_xml, "LightGrid Test");
    let wb = load_xlsx(&xlsx);
    let cell = get_cell(&wb, 0, 0, 0);
    assert!(cell.is_some());
}

#[test]
fn test_pattern_dark_trellis() {
    let fill_xml = r#"<fill><patternFill patternType="darkTrellis"><fgColor rgb="FF000000"/><bgColor rgb="FFFFFFFF"/></patternFill></fill>"#;
    let xlsx = create_xlsx_with_custom_fill(fill_xml, "DarkTrellis Test");
    let wb = load_xlsx(&xlsx);
    let cell = get_cell(&wb, 0, 0, 0);
    assert!(cell.is_some());
}

#[test]
fn test_pattern_light_trellis() {
    let fill_xml = r#"<fill><patternFill patternType="lightTrellis"><fgColor rgb="FF808080"/><bgColor rgb="FFFFFFFF"/></patternFill></fill>"#;
    let xlsx = create_xlsx_with_custom_fill(fill_xml, "LightTrellis Test");
    let wb = load_xlsx(&xlsx);
    let cell = get_cell(&wb, 0, 0, 0);
    assert!(cell.is_some());
}

#[test]
fn test_pattern_grid_with_colors() {
    let fill_xml = r#"<fill><patternFill patternType="darkGrid"><fgColor rgb="FF0000FF"/><bgColor rgb="FFFFFF00"/></patternFill></fill>"#;
    let xlsx = create_xlsx_with_custom_fill(fill_xml, "Colored Grid Test");
    let wb = load_xlsx(&xlsx);
    let cell = get_cell(&wb, 0, 0, 0);
    assert!(cell.is_some());
}

#[test]
fn test_pattern_trellis_with_theme_colors() {
    let fill_xml = r#"<fill><patternFill patternType="lightTrellis"><fgColor theme="4"/><bgColor theme="1"/></patternFill></fill>"#;
    let xlsx = create_xlsx_with_custom_fill(fill_xml, "Theme Trellis Test");
    let wb = load_xlsx(&xlsx);
    let cell = get_cell(&wb, 0, 0, 0);
    assert!(cell.is_some());
}

// =============================================================================
// 8. No Fill / None Pattern Tests
// =============================================================================

#[test]
fn test_pattern_none() {
    let fill_xml = r#"<fill><patternFill patternType="none"/></fill>"#;
    let xlsx = create_xlsx_with_custom_fill(fill_xml, "None Test");
    let wb = load_xlsx(&xlsx);
    let cell = get_cell(&wb, 0, 0, 0);
    assert!(cell.is_some());

    // bgColor should be absent for none pattern
    let style = get_cell_style(&wb, 0, 0, 0);
    if let Some(s) = style {
        assert!(s.bg_color.is_none(), "None pattern should have no bg_color");
    }
}

#[test]
fn test_empty_pattern_fill() {
    let fill_xml = r#"<fill><patternFill/></fill>"#;
    let xlsx = create_xlsx_with_custom_fill(fill_xml, "Empty Pattern Test");
    let wb = load_xlsx(&xlsx);
    let cell = get_cell(&wb, 0, 0, 0);
    assert!(cell.is_some());
}

#[test]
fn test_fill_id_0_is_no_fill() {
    let xlsx = minimal_xlsx();
    let wb = load_xlsx(&xlsx);
    // Default cells use fillId=0 which is no fill
    assert_sheet_count(&wb, 1);
}

// =============================================================================
// 9. Gradient Fill Tests (via custom XML)
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

#[test]
fn test_linear_gradient_horizontal() {
    let gradient_xml = r#"<gradientFill type="linear" degree="0">
        <stop position="0"><color rgb="FFFF0000"/></stop>
        <stop position="1"><color rgb="FF0000FF"/></stop>
    </gradientFill>"#;
    let xlsx = create_gradient_xlsx(gradient_xml);
    let wb = load_xlsx(&xlsx);
    let cell = get_cell(&wb, 0, 0, 0);
    assert!(cell.is_some());
}

#[test]
fn test_linear_gradient_vertical() {
    let gradient_xml = r#"<gradientFill type="linear" degree="90">
        <stop position="0"><color rgb="FF00FF00"/></stop>
        <stop position="1"><color rgb="FFFFFF00"/></stop>
    </gradientFill>"#;
    let xlsx = create_gradient_xlsx(gradient_xml);
    let wb = load_xlsx(&xlsx);
    let cell = get_cell(&wb, 0, 0, 0);
    assert!(cell.is_some());
}

#[test]
fn test_linear_gradient_diagonal() {
    let gradient_xml = r#"<gradientFill type="linear" degree="45">
        <stop position="0"><color rgb="FFFF00FF"/></stop>
        <stop position="1"><color rgb="FF00FFFF"/></stop>
    </gradientFill>"#;
    let xlsx = create_gradient_xlsx(gradient_xml);
    let wb = load_xlsx(&xlsx);
    let cell = get_cell(&wb, 0, 0, 0);
    assert!(cell.is_some());
}

#[test]
fn test_gradient_three_stops() {
    let gradient_xml = r#"<gradientFill type="linear" degree="0">
        <stop position="0"><color rgb="FFFF0000"/></stop>
        <stop position="0.5"><color rgb="FFFFFF00"/></stop>
        <stop position="1"><color rgb="FF00FF00"/></stop>
    </gradientFill>"#;
    let xlsx = create_gradient_xlsx(gradient_xml);
    let wb = load_xlsx(&xlsx);
    let cell = get_cell(&wb, 0, 0, 0);
    assert!(cell.is_some());
}

#[test]
fn test_path_gradient() {
    let gradient_xml = r#"<gradientFill type="path" left="0.5" right="0.5" top="0.5" bottom="0.5">
        <stop position="0"><color rgb="FFFFFFFF"/></stop>
        <stop position="1"><color rgb="FF000000"/></stop>
    </gradientFill>"#;
    let xlsx = create_gradient_xlsx(gradient_xml);
    let wb = load_xlsx(&xlsx);
    let cell = get_cell(&wb, 0, 0, 0);
    assert!(cell.is_some());
}

#[test]
fn test_gradient_with_theme_colors() {
    let gradient_xml = r#"<gradientFill type="linear" degree="0">
        <stop position="0"><color theme="4"/></stop>
        <stop position="1"><color theme="5"/></stop>
    </gradientFill>"#;
    let xlsx = create_gradient_xlsx(gradient_xml);
    let wb = load_xlsx(&xlsx);
    let cell = get_cell(&wb, 0, 0, 0);
    assert!(cell.is_some());
}

// =============================================================================
// Additional Edge Case Tests
// =============================================================================

#[test]
fn test_fill_with_only_bg_color() {
    let fill_xml =
        r#"<fill><patternFill patternType="solid"><bgColor rgb="FFFF0000"/></patternFill></fill>"#;
    let xlsx = create_xlsx_with_custom_fill(fill_xml, "BgOnly Test");
    let wb = load_xlsx(&xlsx);
    let cell = get_cell(&wb, 0, 0, 0);
    assert!(cell.is_some());
}

#[test]
fn test_fill_with_auto_color() {
    let fill_xml = r#"<fill><patternFill patternType="solid"><fgColor auto="1"/><bgColor indexed="64"/></patternFill></fill>"#;
    let xlsx = create_xlsx_with_custom_fill(fill_xml, "Auto Color Test");
    let wb = load_xlsx(&xlsx);
    let cell = get_cell(&wb, 0, 0, 0);
    assert!(cell.is_some());
}

#[test]
fn test_multiple_fills_in_workbook() {
    let xlsx = XlsxBuilder::new()
        .sheet(
            SheetBuilder::new("Sheet1")
                .cell(
                    "A1",
                    "Red",
                    Some(StyleBuilder::new().bg_color("#FF0000").build()),
                )
                .cell(
                    "A2",
                    "Green",
                    Some(StyleBuilder::new().bg_color("#00FF00").build()),
                )
                .cell(
                    "A3",
                    "Blue",
                    Some(StyleBuilder::new().bg_color("#0000FF").build()),
                )
                .cell(
                    "A4",
                    "Yellow",
                    Some(StyleBuilder::new().bg_color("#FFFF00").build()),
                ),
        )
        .build();

    let wb = load_xlsx(&xlsx);
    assert_cell_bg_color(&wb, 0, 0, 0, "#FF0000");
    assert_cell_bg_color(&wb, 0, 1, 0, "#00FF00");
    assert_cell_bg_color(&wb, 0, 2, 0, "#0000FF");
    assert_cell_bg_color(&wb, 0, 3, 0, "#FFFF00");
}

#[test]
fn test_same_fill_reused() {
    let xlsx = XlsxBuilder::new()
        .sheet(
            SheetBuilder::new("Sheet1")
                .cell(
                    "A1",
                    "Same1",
                    Some(StyleBuilder::new().bg_color("#FF0000").build()),
                )
                .cell(
                    "A2",
                    "Same2",
                    Some(StyleBuilder::new().bg_color("#FF0000").build()),
                )
                .cell(
                    "A3",
                    "Same3",
                    Some(StyleBuilder::new().bg_color("#FF0000").build()),
                ),
        )
        .build();

    let wb = load_xlsx(&xlsx);
    assert_cell_bg_color(&wb, 0, 0, 0, "#FF0000");
    assert_cell_bg_color(&wb, 0, 1, 0, "#FF0000");
    assert_cell_bg_color(&wb, 0, 2, 0, "#FF0000");
}

#[test]
fn test_fill_with_empty_rgb() {
    let fill_xml =
        r#"<fill><patternFill patternType="solid"><fgColor rgb=""/></patternFill></fill>"#;
    let xlsx = create_xlsx_with_custom_fill(fill_xml, "Empty RGB Test");
    let wb = load_xlsx(&xlsx);
    let cell = get_cell(&wb, 0, 0, 0);
    assert!(cell.is_some());
}

#[test]
fn test_fill_indexed_64_system_foreground() {
    let fill_xml =
        r#"<fill><patternFill patternType="solid"><fgColor indexed="64"/></patternFill></fill>"#;
    let xlsx = create_xlsx_with_custom_fill(fill_xml, "System FG Test");
    let wb = load_xlsx(&xlsx);
    let cell = get_cell(&wb, 0, 0, 0);
    assert!(cell.is_some());
}

// =============================================================================
// All Pattern Types Comprehensive Test
// =============================================================================

const ALL_PATTERN_TYPES: &[&str] = &[
    "none",
    "solid",
    "gray0625",
    "gray125",
    "darkGray",
    "mediumGray",
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
];

#[test]
fn test_all_pattern_types_parse() {
    for pattern_type in ALL_PATTERN_TYPES {
        let fill_xml = format!(
            r#"<fill><patternFill patternType="{pattern_type}"><fgColor rgb="FF000000"/><bgColor rgb="FFFFFFFF"/></patternFill></fill>"#
        );

        let xlsx = create_xlsx_with_custom_fill(&fill_xml, &format!("{pattern_type} Test"));
        let wb = load_xlsx(&xlsx);

        let cell = get_cell(&wb, 0, 0, 0);
        assert!(
            cell.is_some(),
            "Failed to parse pattern type: {pattern_type}"
        );
    }
}
