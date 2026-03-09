//! Tests for conditional formatting parsing in XLSX files
//!
//! Conditional formatting in Excel allows cells to be formatted based on their values.
//! Rules are stored in the worksheet XML within `<conditionalFormatting>` elements.
//!
//! Types of conditional formatting:
//! - colorScale: Gradient color scale (2-color or 3-color)
//! - dataBar: Data bars showing relative values
//! - iconSet: Icon sets based on thresholds
//! - Cell value rules (greater than, less than, between, etc.)
//! - Formula-based rules

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

use std::io::{Cursor, Write};
use zip::write::FileOptions;
use zip::ZipWriter;

// =============================================================================
// Test Helpers
// =============================================================================

/// Create an XLSX with conditional formatting XML in the worksheet
fn create_xlsx_with_conditional_formatting(cf_xml: &str) -> Vec<u8> {
    let cursor = Cursor::new(Vec::new());
    let mut zip = ZipWriter::new(cursor);
    let options = FileOptions::<()>::default().compression_method(zip::CompressionMethod::Deflated);

    // [Content_Types].xml
    let _ = zip.start_file("[Content_Types].xml", options);
    let _ = zip.write_all(
        br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
</Types>"#,
    );

    // _rels/.rels
    let _ = zip.start_file("_rels/.rels", options);
    let _ = zip.write_all(
        br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>"#,
    );

    // xl/_rels/workbook.xml.rels
    let _ = zip.start_file("xl/_rels/workbook.xml.rels", options);
    let _ = zip.write_all(
        br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>"#,
    );

    // xl/workbook.xml
    let _ = zip.start_file("xl/workbook.xml", options);
    let _ = zip.write_all(
        br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<sheets>
<sheet name="Sheet1" sheetId="1" r:id="rId1"/>
</sheets>
</workbook>"#,
    );

    // xl/styles.xml
    let _ = zip.start_file("xl/styles.xml", options);
    let _ = zip.write_all(
        br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>
<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>
<borders count="1"><border><left/><right/><top/><bottom/></border></borders>
<cellXfs count="1"><xf fontId="0" fillId="0" borderId="0"/></cellXfs>
</styleSheet>"#,
    );

    // xl/worksheets/sheet1.xml with conditional formatting
    let _ = zip.start_file("xl/worksheets/sheet1.xml", options);
    let sheet_xml = format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<sheetData>
<row r="1"><c r="A1"><v>10</v></c><c r="B1"><v>20</v></c><c r="C1"><v>30</v></c></row>
<row r="2"><c r="A2"><v>40</v></c><c r="B2"><v>50</v></c><c r="C2"><v>60</v></c></row>
<row r="3"><c r="A3"><v>70</v></c><c r="B3"><v>80</v></c><c r="C3"><v>90</v></c></row>
</sheetData>
{cf_xml}
</worksheet>"#
    );
    let _ = zip.write_all(sheet_xml.as_bytes());

    let cursor = zip.finish().expect("Failed to finish ZIP");
    cursor.into_inner()
}

// =============================================================================
// Tests: Color Scale - 2 Color
// =============================================================================

#[test]
fn test_color_scale_two_color_min_max() {
    let cf_xml = r#"
<conditionalFormatting sqref="A1:C3">
  <cfRule type="colorScale" priority="1">
    <colorScale>
      <cfvo type="min"/>
      <cfvo type="max"/>
      <color rgb="FFF8696B"/>
      <color rgb="FF63BE7B"/>
    </colorScale>
  </cfRule>
</conditionalFormatting>"#;

    let xlsx = create_xlsx_with_conditional_formatting(cf_xml);
    let workbook = common::load_xlsx(&xlsx);

    let sheet = &workbook.sheets[0];
    assert!(
        !sheet.conditional_formatting.is_empty(),
        "Should have conditional formatting"
    );

    let cf = &sheet.conditional_formatting[0];
    assert_eq!(cf.sqref, "A1:C3");
    assert_eq!(cf.rules.len(), 1);

    let rule = &cf.rules[0];
    assert_eq!(rule.rule_type, "colorScale");
    assert_eq!(rule.priority, 1);

    // Check color scale
    assert!(rule.color_scale.is_some());
    let color_scale = rule.color_scale.as_ref().unwrap();

    // Should have 2 cfvo entries
    assert_eq!(color_scale.cfvo.len(), 2);
    assert_eq!(color_scale.cfvo[0].cfvo_type, "min");
    assert_eq!(color_scale.cfvo[1].cfvo_type, "max");

    // Should have 2 colors
    assert_eq!(color_scale.colors.len(), 2);
}

#[test]
fn test_color_scale_two_color_percentile() {
    let cf_xml = r#"
<conditionalFormatting sqref="A1:A100">
  <cfRule type="colorScale" priority="1">
    <colorScale>
      <cfvo type="percentile" val="10"/>
      <cfvo type="percentile" val="90"/>
      <color rgb="FFFF0000"/>
      <color rgb="FF00FF00"/>
    </colorScale>
  </cfRule>
</conditionalFormatting>"#;

    let xlsx = create_xlsx_with_conditional_formatting(cf_xml);
    let workbook = common::load_xlsx(&xlsx);

    let sheet = &workbook.sheets[0];
    let cf = &sheet.conditional_formatting[0];
    let rule = &cf.rules[0];
    let color_scale = rule.color_scale.as_ref().unwrap();

    assert_eq!(color_scale.cfvo[0].cfvo_type, "percentile");
    assert_eq!(color_scale.cfvo[0].val.as_deref(), Some("10"));
    assert_eq!(color_scale.cfvo[1].cfvo_type, "percentile");
    assert_eq!(color_scale.cfvo[1].val.as_deref(), Some("90"));
}

#[test]
fn test_color_scale_two_color_number() {
    let cf_xml = r#"
<conditionalFormatting sqref="B1:B50">
  <cfRule type="colorScale" priority="1">
    <colorScale>
      <cfvo type="num" val="0"/>
      <cfvo type="num" val="100"/>
      <color rgb="FFFFFFFF"/>
      <color rgb="FF0000FF"/>
    </colorScale>
  </cfRule>
</conditionalFormatting>"#;

    let xlsx = create_xlsx_with_conditional_formatting(cf_xml);
    let workbook = common::load_xlsx(&xlsx);

    let sheet = &workbook.sheets[0];
    let rule = &sheet.conditional_formatting[0].rules[0];
    let color_scale = rule.color_scale.as_ref().unwrap();

    assert_eq!(color_scale.cfvo[0].cfvo_type, "num");
    assert_eq!(color_scale.cfvo[0].val.as_deref(), Some("0"));
    assert_eq!(color_scale.cfvo[1].cfvo_type, "num");
    assert_eq!(color_scale.cfvo[1].val.as_deref(), Some("100"));
}

// =============================================================================
// Tests: Color Scale - 3 Color
// =============================================================================

#[test]
fn test_color_scale_three_color() {
    let cf_xml = r#"
<conditionalFormatting sqref="A1:C10">
  <cfRule type="colorScale" priority="1">
    <colorScale>
      <cfvo type="min"/>
      <cfvo type="percentile" val="50"/>
      <cfvo type="max"/>
      <color rgb="FFF8696B"/>
      <color rgb="FFFFEB84"/>
      <color rgb="FF63BE7B"/>
    </colorScale>
  </cfRule>
</conditionalFormatting>"#;

    let xlsx = create_xlsx_with_conditional_formatting(cf_xml);
    let workbook = common::load_xlsx(&xlsx);

    let sheet = &workbook.sheets[0];
    let rule = &sheet.conditional_formatting[0].rules[0];
    let color_scale = rule.color_scale.as_ref().unwrap();

    // Should have 3 cfvo entries
    assert_eq!(color_scale.cfvo.len(), 3);
    assert_eq!(color_scale.cfvo[0].cfvo_type, "min");
    assert_eq!(color_scale.cfvo[1].cfvo_type, "percentile");
    assert_eq!(color_scale.cfvo[1].val.as_deref(), Some("50"));
    assert_eq!(color_scale.cfvo[2].cfvo_type, "max");

    // Should have 3 colors
    assert_eq!(color_scale.colors.len(), 3);
}

#[test]
fn test_color_scale_three_color_custom_values() {
    let cf_xml = r#"
<conditionalFormatting sqref="D1:D100">
  <cfRule type="colorScale" priority="1">
    <colorScale>
      <cfvo type="num" val="0"/>
      <cfvo type="num" val="50"/>
      <cfvo type="num" val="100"/>
      <color rgb="FFFF0000"/>
      <color rgb="FFFFFF00"/>
      <color rgb="FF00FF00"/>
    </colorScale>
  </cfRule>
</conditionalFormatting>"#;

    let xlsx = create_xlsx_with_conditional_formatting(cf_xml);
    let workbook = common::load_xlsx(&xlsx);

    let sheet = &workbook.sheets[0];
    let rule = &sheet.conditional_formatting[0].rules[0];
    let color_scale = rule.color_scale.as_ref().unwrap();

    assert_eq!(color_scale.cfvo.len(), 3);
    assert_eq!(color_scale.cfvo[0].val.as_deref(), Some("0"));
    assert_eq!(color_scale.cfvo[1].val.as_deref(), Some("50"));
    assert_eq!(color_scale.cfvo[2].val.as_deref(), Some("100"));
}

// =============================================================================
// Tests: Data Bar
// =============================================================================

#[test]
fn test_data_bar_basic() {
    let cf_xml = r#"
<conditionalFormatting sqref="A1:A10">
  <cfRule type="dataBar" priority="1">
    <dataBar>
      <cfvo type="min"/>
      <cfvo type="max"/>
      <color rgb="FF638EC6"/>
    </dataBar>
  </cfRule>
</conditionalFormatting>"#;

    let xlsx = create_xlsx_with_conditional_formatting(cf_xml);
    let workbook = common::load_xlsx(&xlsx);

    let sheet = &workbook.sheets[0];
    let rule = &sheet.conditional_formatting[0].rules[0];

    assert_eq!(rule.rule_type, "dataBar");
    assert!(rule.data_bar.is_some());

    let data_bar = rule.data_bar.as_ref().unwrap();
    assert_eq!(data_bar.cfvo.len(), 2);
    assert!(!data_bar.color.is_empty());
}

#[test]
fn test_data_bar_with_percentile() {
    let cf_xml = r#"
<conditionalFormatting sqref="B1:B50">
  <cfRule type="dataBar" priority="1">
    <dataBar>
      <cfvo type="percentile" val="5"/>
      <cfvo type="percentile" val="95"/>
      <color rgb="FF00B050"/>
    </dataBar>
  </cfRule>
</conditionalFormatting>"#;

    let xlsx = create_xlsx_with_conditional_formatting(cf_xml);
    let workbook = common::load_xlsx(&xlsx);

    let sheet = &workbook.sheets[0];
    let rule = &sheet.conditional_formatting[0].rules[0];
    let data_bar = rule.data_bar.as_ref().unwrap();

    assert_eq!(data_bar.cfvo[0].cfvo_type, "percentile");
    assert_eq!(data_bar.cfvo[0].val.as_deref(), Some("5"));
    assert_eq!(data_bar.cfvo[1].cfvo_type, "percentile");
    assert_eq!(data_bar.cfvo[1].val.as_deref(), Some("95"));
}

#[test]
fn test_data_bar_with_number_values() {
    let cf_xml = r#"
<conditionalFormatting sqref="C1:C100">
  <cfRule type="dataBar" priority="1">
    <dataBar>
      <cfvo type="num" val="0"/>
      <cfvo type="num" val="1000"/>
      <color rgb="FFFFC000"/>
    </dataBar>
  </cfRule>
</conditionalFormatting>"#;

    let xlsx = create_xlsx_with_conditional_formatting(cf_xml);
    let workbook = common::load_xlsx(&xlsx);

    let sheet = &workbook.sheets[0];
    let rule = &sheet.conditional_formatting[0].rules[0];
    let data_bar = rule.data_bar.as_ref().unwrap();

    assert_eq!(data_bar.cfvo[0].cfvo_type, "num");
    assert_eq!(data_bar.cfvo[0].val.as_deref(), Some("0"));
    assert_eq!(data_bar.cfvo[1].cfvo_type, "num");
    assert_eq!(data_bar.cfvo[1].val.as_deref(), Some("1000"));
}

#[test]
fn test_data_bar_show_value() {
    let cf_xml = r#"
<conditionalFormatting sqref="D1:D20">
  <cfRule type="dataBar" priority="1">
    <dataBar showValue="0">
      <cfvo type="min"/>
      <cfvo type="max"/>
      <color rgb="FF5B9BD5"/>
    </dataBar>
  </cfRule>
</conditionalFormatting>"#;

    let xlsx = create_xlsx_with_conditional_formatting(cf_xml);
    let workbook = common::load_xlsx(&xlsx);

    let sheet = &workbook.sheets[0];
    let rule = &sheet.conditional_formatting[0].rules[0];
    let data_bar = rule.data_bar.as_ref().unwrap();

    // showValue="0" means the cell value is hidden (only bar shown)
    assert_eq!(data_bar.show_value, Some(false));
}

// =============================================================================
// Tests: Icon Set
// =============================================================================

#[test]
fn test_icon_set_3_arrows() {
    let cf_xml = r#"
<conditionalFormatting sqref="A1:A100">
  <cfRule type="iconSet" priority="1">
    <iconSet iconSet="3Arrows">
      <cfvo type="percent" val="0"/>
      <cfvo type="percent" val="33"/>
      <cfvo type="percent" val="67"/>
    </iconSet>
  </cfRule>
</conditionalFormatting>"#;

    let xlsx = create_xlsx_with_conditional_formatting(cf_xml);
    let workbook = common::load_xlsx(&xlsx);

    let sheet = &workbook.sheets[0];
    let rule = &sheet.conditional_formatting[0].rules[0];

    assert_eq!(rule.rule_type, "iconSet");
    assert!(rule.icon_set.is_some());

    let icon_set = rule.icon_set.as_ref().unwrap();
    assert_eq!(icon_set.icon_set.as_str(), "3Arrows");
    assert_eq!(icon_set.cfvo.len(), 3);
}

#[test]
fn test_icon_set_4_traffic_lights() {
    let cf_xml = r#"
<conditionalFormatting sqref="B1:B50">
  <cfRule type="iconSet" priority="1">
    <iconSet iconSet="4TrafficLights">
      <cfvo type="percent" val="0"/>
      <cfvo type="percent" val="25"/>
      <cfvo type="percent" val="50"/>
      <cfvo type="percent" val="75"/>
    </iconSet>
  </cfRule>
</conditionalFormatting>"#;

    let xlsx = create_xlsx_with_conditional_formatting(cf_xml);
    let workbook = common::load_xlsx(&xlsx);

    let sheet = &workbook.sheets[0];
    let rule = &sheet.conditional_formatting[0].rules[0];
    let icon_set = rule.icon_set.as_ref().unwrap();

    assert_eq!(icon_set.icon_set.as_str(), "4TrafficLights");
    assert_eq!(icon_set.cfvo.len(), 4);
}

#[test]
fn test_icon_set_5_ratings() {
    let cf_xml = r#"
<conditionalFormatting sqref="C1:C100">
  <cfRule type="iconSet" priority="1">
    <iconSet iconSet="5Rating">
      <cfvo type="percent" val="0"/>
      <cfvo type="percent" val="20"/>
      <cfvo type="percent" val="40"/>
      <cfvo type="percent" val="60"/>
      <cfvo type="percent" val="80"/>
    </iconSet>
  </cfRule>
</conditionalFormatting>"#;

    let xlsx = create_xlsx_with_conditional_formatting(cf_xml);
    let workbook = common::load_xlsx(&xlsx);

    let sheet = &workbook.sheets[0];
    let rule = &sheet.conditional_formatting[0].rules[0];
    let icon_set = rule.icon_set.as_ref().unwrap();

    assert_eq!(icon_set.icon_set.as_str(), "5Rating");
    assert_eq!(icon_set.cfvo.len(), 5);
}

#[test]
fn test_icon_set_reverse_order() {
    let cf_xml = r#"
<conditionalFormatting sqref="D1:D50">
  <cfRule type="iconSet" priority="1">
    <iconSet iconSet="3Symbols" reverse="1">
      <cfvo type="percent" val="0"/>
      <cfvo type="percent" val="33"/>
      <cfvo type="percent" val="67"/>
    </iconSet>
  </cfRule>
</conditionalFormatting>"#;

    let xlsx = create_xlsx_with_conditional_formatting(cf_xml);
    let workbook = common::load_xlsx(&xlsx);

    let sheet = &workbook.sheets[0];
    let rule = &sheet.conditional_formatting[0].rules[0];
    let icon_set = rule.icon_set.as_ref().unwrap();

    assert_eq!(icon_set.reverse, Some(true));
}

#[test]
fn test_icon_set_show_value() {
    let cf_xml = r#"
<conditionalFormatting sqref="E1:E30">
  <cfRule type="iconSet" priority="1">
    <iconSet iconSet="3Flags" showValue="0">
      <cfvo type="percent" val="0"/>
      <cfvo type="percent" val="33"/>
      <cfvo type="percent" val="67"/>
    </iconSet>
  </cfRule>
</conditionalFormatting>"#;

    let xlsx = create_xlsx_with_conditional_formatting(cf_xml);
    let workbook = common::load_xlsx(&xlsx);

    let sheet = &workbook.sheets[0];
    let rule = &sheet.conditional_formatting[0].rules[0];
    let icon_set = rule.icon_set.as_ref().unwrap();

    // showValue="0" means only icons shown, no values
    assert_eq!(icon_set.show_value, Some(false));
}

// =============================================================================
// Tests: Cell Ranges
// =============================================================================

#[test]
fn test_cf_single_cell_range() {
    let cf_xml = r#"
<conditionalFormatting sqref="A1">
  <cfRule type="colorScale" priority="1">
    <colorScale>
      <cfvo type="min"/>
      <cfvo type="max"/>
      <color rgb="FFFF0000"/>
      <color rgb="FF00FF00"/>
    </colorScale>
  </cfRule>
</conditionalFormatting>"#;

    let xlsx = create_xlsx_with_conditional_formatting(cf_xml);
    let workbook = common::load_xlsx(&xlsx);

    let sheet = &workbook.sheets[0];
    let cf = &sheet.conditional_formatting[0];
    assert_eq!(cf.sqref, "A1");
}

#[test]
fn test_cf_rectangular_range() {
    let cf_xml = r#"
<conditionalFormatting sqref="A1:Z100">
  <cfRule type="dataBar" priority="1">
    <dataBar>
      <cfvo type="min"/>
      <cfvo type="max"/>
      <color rgb="FF0000FF"/>
    </dataBar>
  </cfRule>
</conditionalFormatting>"#;

    let xlsx = create_xlsx_with_conditional_formatting(cf_xml);
    let workbook = common::load_xlsx(&xlsx);

    let sheet = &workbook.sheets[0];
    let cf = &sheet.conditional_formatting[0];
    assert_eq!(cf.sqref, "A1:Z100");
}

#[test]
fn test_cf_multiple_ranges() {
    let cf_xml = r#"
<conditionalFormatting sqref="A1:A10 C1:C10 E1:E10">
  <cfRule type="iconSet" priority="1">
    <iconSet iconSet="3Arrows">
      <cfvo type="percent" val="0"/>
      <cfvo type="percent" val="33"/>
      <cfvo type="percent" val="67"/>
    </iconSet>
  </cfRule>
</conditionalFormatting>"#;

    let xlsx = create_xlsx_with_conditional_formatting(cf_xml);
    let workbook = common::load_xlsx(&xlsx);

    let sheet = &workbook.sheets[0];
    let cf = &sheet.conditional_formatting[0];
    assert_eq!(cf.sqref, "A1:A10 C1:C10 E1:E10");
}

#[test]
fn test_cf_entire_column() {
    let cf_xml = r#"
<conditionalFormatting sqref="A:A">
  <cfRule type="colorScale" priority="1">
    <colorScale>
      <cfvo type="min"/>
      <cfvo type="max"/>
      <color rgb="FFFFFFFF"/>
      <color rgb="FF000000"/>
    </colorScale>
  </cfRule>
</conditionalFormatting>"#;

    let xlsx = create_xlsx_with_conditional_formatting(cf_xml);
    let workbook = common::load_xlsx(&xlsx);

    let sheet = &workbook.sheets[0];
    let cf = &sheet.conditional_formatting[0];
    assert_eq!(cf.sqref, "A:A");
}

// =============================================================================
// Tests: Priority Ordering
// =============================================================================

#[test]
fn test_cf_priority_ordering() {
    let cf_xml = r#"
<conditionalFormatting sqref="A1:A10">
  <cfRule type="colorScale" priority="2">
    <colorScale>
      <cfvo type="min"/>
      <cfvo type="max"/>
      <color rgb="FFFF0000"/>
      <color rgb="FF00FF00"/>
    </colorScale>
  </cfRule>
  <cfRule type="dataBar" priority="1">
    <dataBar>
      <cfvo type="min"/>
      <cfvo type="max"/>
      <color rgb="FF0000FF"/>
    </dataBar>
  </cfRule>
</conditionalFormatting>"#;

    let xlsx = create_xlsx_with_conditional_formatting(cf_xml);
    let workbook = common::load_xlsx(&xlsx);

    let sheet = &workbook.sheets[0];
    let cf = &sheet.conditional_formatting[0];

    // Should have 2 rules
    assert_eq!(cf.rules.len(), 2);

    // Check priorities
    let priorities: Vec<_> = cf.rules.iter().map(|r| r.priority).collect();
    assert!(priorities.contains(&1));
    assert!(priorities.contains(&2));
}

#[test]
fn test_cf_multiple_formatting_elements() {
    let cf_xml = r#"
<conditionalFormatting sqref="A1:A10">
  <cfRule type="colorScale" priority="1">
    <colorScale>
      <cfvo type="min"/>
      <cfvo type="max"/>
      <color rgb="FFFF0000"/>
      <color rgb="FF00FF00"/>
    </colorScale>
  </cfRule>
</conditionalFormatting>
<conditionalFormatting sqref="B1:B10">
  <cfRule type="dataBar" priority="2">
    <dataBar>
      <cfvo type="min"/>
      <cfvo type="max"/>
      <color rgb="FF0000FF"/>
    </dataBar>
  </cfRule>
</conditionalFormatting>
<conditionalFormatting sqref="C1:C10">
  <cfRule type="iconSet" priority="3">
    <iconSet iconSet="3Arrows">
      <cfvo type="percent" val="0"/>
      <cfvo type="percent" val="33"/>
      <cfvo type="percent" val="67"/>
    </iconSet>
  </cfRule>
</conditionalFormatting>"#;

    let xlsx = create_xlsx_with_conditional_formatting(cf_xml);
    let workbook = common::load_xlsx(&xlsx);

    let sheet = &workbook.sheets[0];

    // Should have 3 conditional formatting elements
    assert_eq!(sheet.conditional_formatting.len(), 3);

    // Check each range
    let sqrefs: Vec<_> = sheet
        .conditional_formatting
        .iter()
        .map(|cf| cf.sqref.as_str())
        .collect();
    assert!(sqrefs.contains(&"A1:A10"));
    assert!(sqrefs.contains(&"B1:B10"));
    assert!(sqrefs.contains(&"C1:C10"));
}

// =============================================================================
// Tests: Edge Cases
// =============================================================================

#[test]
fn test_no_conditional_formatting() {
    use fixtures::{SheetBuilder, XlsxBuilder};

    let xlsx = XlsxBuilder::new()
        .sheet(SheetBuilder::new("Sheet1").cell("A1", "No CF", None))
        .build();

    let workbook = common::load_xlsx(&xlsx);
    let sheet = &workbook.sheets[0];

    assert!(sheet.conditional_formatting.is_empty());
}

#[test]
fn test_cf_formula_type_value() {
    let cf_xml = r#"
<conditionalFormatting sqref="A1:A100">
  <cfRule type="colorScale" priority="1">
    <colorScale>
      <cfvo type="formula" val="MIN($A:$A)"/>
      <cfvo type="formula" val="MAX($A:$A)"/>
      <color rgb="FFFF0000"/>
      <color rgb="FF00FF00"/>
    </colorScale>
  </cfRule>
</conditionalFormatting>"#;

    let xlsx = create_xlsx_with_conditional_formatting(cf_xml);
    let workbook = common::load_xlsx(&xlsx);

    let sheet = &workbook.sheets[0];
    let rule = &sheet.conditional_formatting[0].rules[0];
    let color_scale = rule.color_scale.as_ref().unwrap();

    assert_eq!(color_scale.cfvo[0].cfvo_type, "formula");
    assert_eq!(color_scale.cfvo[0].val.as_deref(), Some("MIN($A:$A)"));
    assert_eq!(color_scale.cfvo[1].cfvo_type, "formula");
    assert_eq!(color_scale.cfvo[1].val.as_deref(), Some("MAX($A:$A)"));
}

#[test]
fn test_cf_automin_automax() {
    let cf_xml = r#"
<conditionalFormatting sqref="A1:A50">
  <cfRule type="dataBar" priority="1">
    <dataBar>
      <cfvo type="autoMin"/>
      <cfvo type="autoMax"/>
      <color rgb="FF5B9BD5"/>
    </dataBar>
  </cfRule>
</conditionalFormatting>"#;

    let xlsx = create_xlsx_with_conditional_formatting(cf_xml);
    let workbook = common::load_xlsx(&xlsx);

    let sheet = &workbook.sheets[0];
    let rule = &sheet.conditional_formatting[0].rules[0];
    let data_bar = rule.data_bar.as_ref().unwrap();

    assert_eq!(data_bar.cfvo[0].cfvo_type, "autoMin");
    assert_eq!(data_bar.cfvo[1].cfvo_type, "autoMax");
}

// =============================================================================
// Tests: Serialization
// =============================================================================

#[test]
fn test_cf_serialization_to_json() {
    let cf_xml = r#"
<conditionalFormatting sqref="A1:C10">
  <cfRule type="colorScale" priority="1">
    <colorScale>
      <cfvo type="min"/>
      <cfvo type="percentile" val="50"/>
      <cfvo type="max"/>
      <color rgb="FFF8696B"/>
      <color rgb="FFFFEB84"/>
      <color rgb="FF63BE7B"/>
    </colorScale>
  </cfRule>
</conditionalFormatting>"#;

    let xlsx = create_xlsx_with_conditional_formatting(cf_xml);
    let workbook = common::load_xlsx(&xlsx);

    // Serialize conditional formatting to JSON
    let cf_list = &workbook.sheets[0].conditional_formatting;
    let json = serde_json::to_value(cf_list).expect("Failed to serialize");

    assert!(json.is_array());
    assert_eq!(json.as_array().unwrap().len(), 1);

    let cf_item = &json[0];
    assert_eq!(cf_item["sqref"], "A1:C10");
    assert!(cf_item["rules"].is_array());
    assert_eq!(cf_item["rules"].as_array().unwrap().len(), 1);

    let rule = &cf_item["rules"][0];
    assert_eq!(rule["ruleType"], "colorScale");
    assert_eq!(rule["priority"], 1);
    assert!(rule["colorScale"].is_object());
}

#[test]
fn test_cf_data_bar_serialization() {
    let cf_xml = r#"
<conditionalFormatting sqref="B1:B100">
  <cfRule type="dataBar" priority="1">
    <dataBar showValue="0">
      <cfvo type="num" val="0"/>
      <cfvo type="num" val="100"/>
      <color rgb="FF00B050"/>
    </dataBar>
  </cfRule>
</conditionalFormatting>"#;

    let xlsx = create_xlsx_with_conditional_formatting(cf_xml);
    let workbook = common::load_xlsx(&xlsx);

    let cf_list = &workbook.sheets[0].conditional_formatting;
    let json = serde_json::to_value(cf_list).expect("Failed to serialize");

    let rule = &json[0]["rules"][0];
    assert_eq!(rule["ruleType"], "dataBar");
    assert!(rule["dataBar"].is_object());

    let data_bar = &rule["dataBar"];
    assert_eq!(data_bar["showValue"], false);
}

#[test]
fn test_cf_icon_set_serialization() {
    let cf_xml = r#"
<conditionalFormatting sqref="C1:C50">
  <cfRule type="iconSet" priority="1">
    <iconSet iconSet="4TrafficLights" reverse="1" showValue="0">
      <cfvo type="percent" val="0"/>
      <cfvo type="percent" val="25"/>
      <cfvo type="percent" val="50"/>
      <cfvo type="percent" val="75"/>
    </iconSet>
  </cfRule>
</conditionalFormatting>"#;

    let xlsx = create_xlsx_with_conditional_formatting(cf_xml);
    let workbook = common::load_xlsx(&xlsx);

    let cf_list = &workbook.sheets[0].conditional_formatting;
    let json = serde_json::to_value(cf_list).expect("Failed to serialize");

    let rule = &json[0]["rules"][0];
    let icon_set = &rule["iconSet"];

    assert_eq!(icon_set["iconSet"], "4TrafficLights");
    assert_eq!(icon_set["reverse"], true);
    assert_eq!(icon_set["showValue"], false);
}
