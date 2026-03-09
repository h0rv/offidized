//! Tests for conditional formatting data bars with negative values
//!
//! Excel data bars can handle negative values in several ways:
//! 1. cfvo type="num" with val="-10" (negative minimum)
//! 2. Extended data bar format (x14:dataBar) with negativeBarColorSameAsPositive, etc.
//! 3. Data bars spanning negative to positive ranges
//!
//! This module tests the parsing of data bars that may contain negative values.
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

use std::fs;
use std::io::{Cursor, Write};
use zip::write::FileOptions;
use zip::ZipWriter;

// =============================================================================
// Test Helpers
// =============================================================================

/// Create an XLSX with conditional formatting data bar using negative values
fn create_xlsx_with_negative_data_bar(
    min_val: &str,
    max_val: &str,
    min_type: &str,
    max_type: &str,
) -> Vec<u8> {
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

    // Build cfvo elements with optional val attributes
    let min_cfvo = if min_val.is_empty() {
        format!(r#"<cfvo type="{}"/>"#, min_type)
    } else {
        format!(r#"<cfvo type="{}" val="{}"/>"#, min_type, min_val)
    };

    let max_cfvo = if max_val.is_empty() {
        format!(r#"<cfvo type="{}"/>"#, max_type)
    } else {
        format!(r#"<cfvo type="{}" val="{}"/>"#, max_type, max_val)
    };

    // xl/worksheets/sheet1.xml with data bar containing negative values
    let sheet_xml = format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<sheetData>
<row r="1"><c r="A1"><v>-50</v></c></row>
<row r="2"><c r="A2"><v>-25</v></c></row>
<row r="3"><c r="A3"><v>0</v></c></row>
<row r="4"><c r="A4"><v>25</v></c></row>
<row r="5"><c r="A5"><v>50</v></c></row>
</sheetData>
<conditionalFormatting sqref="A1:A5">
  <cfRule type="dataBar" priority="1">
    <dataBar>
      {}
      {}
      <color rgb="FF638EC6"/>
    </dataBar>
  </cfRule>
</conditionalFormatting>
</worksheet>"#,
        min_cfvo, max_cfvo
    );

    let _ = zip.start_file("xl/worksheets/sheet1.xml", options);
    let _ = zip.write_all(sheet_xml.as_bytes());

    let cursor = zip.finish().expect("Failed to finish ZIP");
    cursor.into_inner()
}

// =============================================================================
// Tests: Data Bar with Negative Number Values
// =============================================================================

#[test]
fn test_data_bar_with_negative_min_value() {
    // Test data bar where min is a negative number
    let xlsx = create_xlsx_with_negative_data_bar("-100", "100", "num", "num");
    let workbook = common::load_xlsx(&xlsx);

    let sheet = &workbook.sheets[0];
    assert!(
        !sheet.conditional_formatting.is_empty(),
        "Should have CF rules"
    );

    let cf = &sheet.conditional_formatting[0];
    assert_eq!(cf.sqref, "A1:A5");

    let rule = &cf.rules[0];
    assert_eq!(rule.rule_type, "dataBar");
    assert!(rule.data_bar.is_some(), "Should have data bar");

    let data_bar = rule.data_bar.as_ref().unwrap();
    assert_eq!(data_bar.cfvo.len(), 2, "Should have 2 cfvo entries");

    // Check that the negative value is correctly parsed
    assert_eq!(data_bar.cfvo[0].cfvo_type, "num");
    assert_eq!(data_bar.cfvo[0].val.as_deref(), Some("-100"));

    assert_eq!(data_bar.cfvo[1].cfvo_type, "num");
    assert_eq!(data_bar.cfvo[1].val.as_deref(), Some("100"));
}

#[test]
fn test_data_bar_with_negative_max_value() {
    // Test edge case where max is also negative (all values negative)
    let xlsx = create_xlsx_with_negative_data_bar("-200", "-50", "num", "num");
    let workbook = common::load_xlsx(&xlsx);

    let sheet = &workbook.sheets[0];
    let cf = &sheet.conditional_formatting[0];
    let rule = &cf.rules[0];
    let data_bar = rule.data_bar.as_ref().unwrap();

    assert_eq!(data_bar.cfvo[0].val.as_deref(), Some("-200"));
    assert_eq!(data_bar.cfvo[1].val.as_deref(), Some("-50"));
}

#[test]
fn test_data_bar_spanning_zero() {
    // Test data bar that spans from negative through zero to positive
    let xlsx = create_xlsx_with_negative_data_bar("-50", "50", "num", "num");
    let workbook = common::load_xlsx(&xlsx);

    let sheet = &workbook.sheets[0];
    let cf = &sheet.conditional_formatting[0];
    let rule = &cf.rules[0];
    let data_bar = rule.data_bar.as_ref().unwrap();

    assert_eq!(data_bar.cfvo[0].val.as_deref(), Some("-50"));
    assert_eq!(data_bar.cfvo[1].val.as_deref(), Some("50"));
}

#[test]
fn test_data_bar_min_max_auto_with_negative_cells() {
    // Test data bar with min/max type - cells contain negative values
    let xlsx = create_xlsx_with_negative_data_bar("", "", "min", "max");
    let workbook = common::load_xlsx(&xlsx);

    let sheet = &workbook.sheets[0];
    let cf = &sheet.conditional_formatting[0];
    let rule = &cf.rules[0];
    let data_bar = rule.data_bar.as_ref().unwrap();

    // min/max types don't have val attributes
    assert_eq!(data_bar.cfvo[0].cfvo_type, "min");
    assert!(data_bar.cfvo[0].val.is_none());

    assert_eq!(data_bar.cfvo[1].cfvo_type, "max");
    assert!(data_bar.cfvo[1].val.is_none());
}

// =============================================================================
// Tests: Data Bar with Negative Percentile/Percent Values
// =============================================================================

#[test]
fn test_data_bar_percentile_spanning_negative_values() {
    // Percentile-based data bar on cells with negative values
    let xlsx = create_xlsx_with_negative_data_bar("10", "90", "percentile", "percentile");
    let workbook = common::load_xlsx(&xlsx);

    let sheet = &workbook.sheets[0];
    let cf = &sheet.conditional_formatting[0];
    let rule = &cf.rules[0];
    let data_bar = rule.data_bar.as_ref().unwrap();

    assert_eq!(data_bar.cfvo[0].cfvo_type, "percentile");
    assert_eq!(data_bar.cfvo[0].val.as_deref(), Some("10"));

    assert_eq!(data_bar.cfvo[1].cfvo_type, "percentile");
    assert_eq!(data_bar.cfvo[1].val.as_deref(), Some("90"));
}

#[test]
fn test_data_bar_percent_with_negative_cells() {
    // Percent-based cfvo types
    let xlsx = create_xlsx_with_negative_data_bar("0", "100", "percent", "percent");
    let workbook = common::load_xlsx(&xlsx);

    let sheet = &workbook.sheets[0];
    let cf = &sheet.conditional_formatting[0];
    let rule = &cf.rules[0];
    let data_bar = rule.data_bar.as_ref().unwrap();

    assert_eq!(data_bar.cfvo[0].cfvo_type, "percent");
    assert_eq!(data_bar.cfvo[0].val.as_deref(), Some("0"));
}

// =============================================================================
// Tests: Data Bar with Mixed cfvo Types
// =============================================================================

#[test]
fn test_data_bar_mixed_min_and_num() {
    // Mix of "min" type and explicit number
    let xlsx = create_xlsx_with_negative_data_bar("", "100", "min", "num");
    let workbook = common::load_xlsx(&xlsx);

    let sheet = &workbook.sheets[0];
    let cf = &sheet.conditional_formatting[0];
    let rule = &cf.rules[0];
    let data_bar = rule.data_bar.as_ref().unwrap();

    assert_eq!(data_bar.cfvo[0].cfvo_type, "min");
    assert!(data_bar.cfvo[0].val.is_none());

    assert_eq!(data_bar.cfvo[1].cfvo_type, "num");
    assert_eq!(data_bar.cfvo[1].val.as_deref(), Some("100"));
}

#[test]
fn test_data_bar_negative_num_to_max() {
    // Negative number minimum, auto max
    let xlsx = create_xlsx_with_negative_data_bar("-25", "", "num", "max");
    let workbook = common::load_xlsx(&xlsx);

    let sheet = &workbook.sheets[0];
    let cf = &sheet.conditional_formatting[0];
    let rule = &cf.rules[0];
    let data_bar = rule.data_bar.as_ref().unwrap();

    assert_eq!(data_bar.cfvo[0].cfvo_type, "num");
    assert_eq!(data_bar.cfvo[0].val.as_deref(), Some("-25"));

    assert_eq!(data_bar.cfvo[1].cfvo_type, "max");
}

// =============================================================================
// Tests: Data Bar Attributes with Negative Values
// =============================================================================

/// Create an XLSX with data bar including minLength and maxLength
fn create_xlsx_with_data_bar_lengths(
    min_val: &str,
    max_val: &str,
    min_length: u32,
    max_length: u32,
) -> Vec<u8> {
    let cursor = Cursor::new(Vec::new());
    let mut zip = ZipWriter::new(cursor);
    let options = FileOptions::<()>::default().compression_method(zip::CompressionMethod::Deflated);

    // Minimal XLSX setup
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
    let _ = zip.write_all(
        br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>
<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>
<borders count="1"><border><left/><right/><top/><bottom/></border></borders>
<cellXfs count="1"><xf fontId="0" fillId="0" borderId="0"/></cellXfs>
</styleSheet>"#,
    );

    let sheet_xml = format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<sheetData>
<row r="1"><c r="A1"><v>-75</v></c></row>
<row r="2"><c r="A2"><v>0</v></c></row>
<row r="3"><c r="A3"><v>75</v></c></row>
</sheetData>
<conditionalFormatting sqref="A1:A3">
  <cfRule type="dataBar" priority="1">
    <dataBar minLength="{}" maxLength="{}">
      <cfvo type="num" val="{}"/>
      <cfvo type="num" val="{}"/>
      <color rgb="FF00B050"/>
    </dataBar>
  </cfRule>
</conditionalFormatting>
</worksheet>"#,
        min_length, max_length, min_val, max_val
    );

    let _ = zip.start_file("xl/worksheets/sheet1.xml", options);
    let _ = zip.write_all(sheet_xml.as_bytes());

    let cursor = zip.finish().expect("Failed to finish ZIP");
    cursor.into_inner()
}

#[test]
fn test_data_bar_with_lengths_and_negative_range() {
    let xlsx = create_xlsx_with_data_bar_lengths("-100", "100", 5, 95);
    let workbook = common::load_xlsx(&xlsx);

    let sheet = &workbook.sheets[0];
    let cf = &sheet.conditional_formatting[0];
    let rule = &cf.rules[0];
    let data_bar = rule.data_bar.as_ref().unwrap();

    assert_eq!(data_bar.min_length, Some(5));
    assert_eq!(data_bar.max_length, Some(95));
    assert_eq!(data_bar.cfvo[0].val.as_deref(), Some("-100"));
    assert_eq!(data_bar.cfvo[1].val.as_deref(), Some("100"));
}

// =============================================================================
// Tests: Data Bar showValue with Negative Values
// =============================================================================

/// Create XLSX with data bar and showValue attribute
fn create_xlsx_with_data_bar_show_value(min_val: &str, max_val: &str, show_value: bool) -> Vec<u8> {
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
<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
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
    let _ = zip.write_all(
        br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>
<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>
<borders count="1"><border><left/><right/><top/><bottom/></border></borders>
<cellXfs count="1"><xf fontId="0" fillId="0" borderId="0"/></cellXfs>
</styleSheet>"#,
    );

    let show_value_attr = if show_value { "" } else { r#" showValue="0""# };

    let sheet_xml = format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<sheetData>
<row r="1"><c r="A1"><v>-30</v></c></row>
<row r="2"><c r="A2"><v>30</v></c></row>
</sheetData>
<conditionalFormatting sqref="A1:A2">
  <cfRule type="dataBar" priority="1">
    <dataBar{}>
      <cfvo type="num" val="{}"/>
      <cfvo type="num" val="{}"/>
      <color rgb="FF5B9BD5"/>
    </dataBar>
  </cfRule>
</conditionalFormatting>
</worksheet>"#,
        show_value_attr, min_val, max_val
    );

    let _ = zip.start_file("xl/worksheets/sheet1.xml", options);
    let _ = zip.write_all(sheet_xml.as_bytes());

    let cursor = zip.finish().expect("Failed to finish ZIP");
    cursor.into_inner()
}

#[test]
fn test_data_bar_negative_values_with_show_value_true() {
    let xlsx = create_xlsx_with_data_bar_show_value("-50", "50", true);
    let workbook = common::load_xlsx(&xlsx);

    let sheet = &workbook.sheets[0];
    let cf = &sheet.conditional_formatting[0];
    let rule = &cf.rules[0];
    let data_bar = rule.data_bar.as_ref().unwrap();

    // showValue defaults to true when not specified
    assert!(data_bar.show_value.is_none() || data_bar.show_value == Some(true));
    assert_eq!(data_bar.cfvo[0].val.as_deref(), Some("-50"));
}

#[test]
fn test_data_bar_negative_values_with_show_value_false() {
    let xlsx = create_xlsx_with_data_bar_show_value("-50", "50", false);
    let workbook = common::load_xlsx(&xlsx);

    let sheet = &workbook.sheets[0];
    let cf = &sheet.conditional_formatting[0];
    let rule = &cf.rules[0];
    let data_bar = rule.data_bar.as_ref().unwrap();

    assert_eq!(data_bar.show_value, Some(false));
    assert_eq!(data_bar.cfvo[0].val.as_deref(), Some("-50"));
}

// =============================================================================
// Tests: JSON Serialization of Negative Data Bar
// =============================================================================

#[test]
fn test_negative_data_bar_serialization() {
    let xlsx = create_xlsx_with_negative_data_bar("-100", "100", "num", "num");
    let workbook = common::load_xlsx(&xlsx);

    let cf_list = &workbook.sheets[0].conditional_formatting;
    let json = serde_json::to_value(cf_list).expect("Failed to serialize");

    assert!(json.is_array());

    let cf_item = &json[0];
    assert_eq!(cf_item["sqref"], "A1:A5");

    let rule = &cf_item["rules"][0];
    assert_eq!(rule["ruleType"], "dataBar");

    let data_bar = &rule["dataBar"];
    assert!(data_bar.is_object());

    let cfvo = &data_bar["cfvo"];
    assert!(cfvo.is_array());
    assert_eq!(cfvo.as_array().unwrap().len(), 2);

    // Check that negative value is preserved in JSON
    assert_eq!(cfvo[0]["cfvoType"], "num");
    assert_eq!(cfvo[0]["val"], "-100");

    assert_eq!(cfvo[1]["cfvoType"], "num");
    assert_eq!(cfvo[1]["val"], "100");
}

// =============================================================================
// Tests: Real XLSX File Parsing
// =============================================================================

#[test]
fn test_ms_cf_samples_data_bars() {
    // Test data bars from ms_cf_samples.xlsx
    let path = "crates/offidized-xlview/test/ms_cf_samples.xlsx";
    if !std::path::Path::new(path).exists() {
        println!("Skipping test: {} not found", path);
        return;
    }

    let data = fs::read(path).expect("Failed to read test file");
    let workbook = common::load_xlsx(&data);

    let mut data_bar_count = 0;
    let mut negative_cfvo_count = 0;

    for sheet in &workbook.sheets {
        for cf in &sheet.conditional_formatting {
            for rule in &cf.rules {
                if rule.rule_type == "dataBar" {
                    data_bar_count += 1;
                    if let Some(ref db) = rule.data_bar {
                        for cfvo in &db.cfvo {
                            if let Some(ref val) = cfvo.val {
                                if val.starts_with('-') {
                                    negative_cfvo_count += 1;
                                    println!(
                                        "Found negative cfvo in sheet '{}': type={}, val={}",
                                        sheet.name, cfvo.cfvo_type, val
                                    );
                                }
                            }
                        }
                    }
                }
            }
        }
    }

    println!(
        "ms_cf_samples.xlsx: {} data bars, {} negative cfvo values",
        data_bar_count, negative_cfvo_count
    );

    // Should have at least some data bars
    assert!(
        data_bar_count > 0,
        "Should find data bar rules in ms_cf_samples.xlsx"
    );
}

#[test]
fn test_kitchen_sink_v2_data_bars() {
    // Test data bars from kitchen_sink_v2.xlsx
    let path = "crates/offidized-xlview/test/kitchen_sink_v2.xlsx";
    if !std::path::Path::new(path).exists() {
        println!("Skipping test: {} not found", path);
        return;
    }

    let data = fs::read(path).expect("Failed to read test file");
    let workbook = common::load_xlsx(&data);

    let mut data_bar_count = 0;

    for sheet in &workbook.sheets {
        for cf in &sheet.conditional_formatting {
            for rule in &cf.rules {
                if rule.rule_type == "dataBar" {
                    data_bar_count += 1;
                    println!(
                        "Found data bar in sheet '{}', sqref: {}",
                        sheet.name, cf.sqref
                    );
                    if let Some(ref db) = rule.data_bar {
                        println!("  color: {}", db.color);
                        println!("  show_value: {:?}", db.show_value);
                        for (i, cfvo) in db.cfvo.iter().enumerate() {
                            println!("  cfvo[{}]: type={}, val={:?}", i, cfvo.cfvo_type, cfvo.val);
                        }
                    }
                }
            }
        }
    }

    println!("kitchen_sink_v2.xlsx: {} data bars", data_bar_count);
}

// =============================================================================
// Tests: Edge Cases for Negative Values
// =============================================================================

#[test]
fn test_data_bar_with_zero_as_boundary() {
    // Test data bar from 0 to positive
    let xlsx = create_xlsx_with_negative_data_bar("0", "100", "num", "num");
    let workbook = common::load_xlsx(&xlsx);

    let sheet = &workbook.sheets[0];
    let cf = &sheet.conditional_formatting[0];
    let rule = &cf.rules[0];
    let data_bar = rule.data_bar.as_ref().unwrap();

    assert_eq!(data_bar.cfvo[0].val.as_deref(), Some("0"));
    assert_eq!(data_bar.cfvo[1].val.as_deref(), Some("100"));
}

#[test]
fn test_data_bar_with_small_negative_fraction() {
    // Test data bar with small negative fraction
    let xlsx = create_xlsx_with_negative_data_bar("-0.5", "0.5", "num", "num");
    let workbook = common::load_xlsx(&xlsx);

    let sheet = &workbook.sheets[0];
    let cf = &sheet.conditional_formatting[0];
    let rule = &cf.rules[0];
    let data_bar = rule.data_bar.as_ref().unwrap();

    assert_eq!(data_bar.cfvo[0].val.as_deref(), Some("-0.5"));
    assert_eq!(data_bar.cfvo[1].val.as_deref(), Some("0.5"));
}

#[test]
fn test_data_bar_with_large_negative_value() {
    // Test data bar with large negative value
    let xlsx = create_xlsx_with_negative_data_bar("-1000000", "1000000", "num", "num");
    let workbook = common::load_xlsx(&xlsx);

    let sheet = &workbook.sheets[0];
    let cf = &sheet.conditional_formatting[0];
    let rule = &cf.rules[0];
    let data_bar = rule.data_bar.as_ref().unwrap();

    assert_eq!(data_bar.cfvo[0].val.as_deref(), Some("-1000000"));
    assert_eq!(data_bar.cfvo[1].val.as_deref(), Some("1000000"));
}
