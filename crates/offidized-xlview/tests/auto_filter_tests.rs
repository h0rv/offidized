//! Tests for auto-filter parsing in XLSX files
//!
//! Excel auto-filters allow users to filter data in a table by column values.
//! The autoFilter element can contain:
//! - A range reference (e.g., "A1:D100")
//! - filterColumn elements with filter criteria
//!
//! Filter types:
//! - Values filter: Show only specific values
//! - Custom filter: Operators like greaterThan, lessThan, contains, etc.
//! - Top10 filter: Show top/bottom N items or percent
//! - Dynamic filter: Date-based filters (thisWeek, thisMonth, etc.)
//! - Color filter: Filter by cell or font color
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

use common::load_xlsx;
use offidized_xlview::types::filter::{CustomFilterOperator, FilterType};

// ============================================================================
// Test Helper: Create XLSX with auto-filter
// ============================================================================

/// Creates a minimal XLSX file with an auto-filter in the sheet XML.
fn create_xlsx_with_auto_filter(auto_filter_xml: &str) -> Vec<u8> {
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
  <Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>
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
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>
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
  <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
  <cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellXfs>
</styleSheet>"#,
    );

    // xl/sharedStrings.xml
    let _ = zip.start_file("xl/sharedStrings.xml", options);
    let _ = zip.write_all(
        br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="4" uniqueCount="4">
  <si><t>Name</t></si>
  <si><t>Age</t></si>
  <si><t>City</t></si>
  <si><t>Score</t></si>
</sst>"#,
    );

    // xl/worksheets/sheet1.xml with auto-filter
    let _ = zip.start_file("xl/worksheets/sheet1.xml", options);
    let sheet_xml = format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1" t="s"><v>0</v></c>
      <c r="B1" t="s"><v>1</v></c>
      <c r="C1" t="s"><v>2</v></c>
      <c r="D1" t="s"><v>3</v></c>
    </row>
  </sheetData>
  {auto_filter_xml}
</worksheet>"#
    );
    let _ = zip.write_all(sheet_xml.as_bytes());

    let cursor = zip.finish().expect("Failed to finish ZIP");
    cursor.into_inner()
}

// ============================================================================
// Test 1: Simple auto-filter (range only, no active filters)
// ============================================================================

#[cfg(test)]
mod simple_auto_filter_tests {
    use super::*;

    #[test]
    fn test_simple_auto_filter_range_only() {
        let xml = r#"<autoFilter ref="A1:D100"/>"#;
        let xlsx = create_xlsx_with_auto_filter(xml);
        let workbook = load_xlsx(&xlsx);

        let sheet = &workbook.sheets[0];
        let af = sheet
            .auto_filter
            .as_ref()
            .expect("auto_filter should be present");
        assert_eq!(af.range, "A1:D100");
        assert_eq!(af.start_row, 0);
        assert_eq!(af.start_col, 0);
        assert_eq!(af.end_row, 99);
        assert_eq!(af.end_col, 3);
        assert!(af.filter_columns.is_empty());
    }

    #[test]
    fn test_auto_filter_range_parsing() {
        let xml = r#"<autoFilter ref="B2:F50"/>"#;
        let xlsx = create_xlsx_with_auto_filter(xml);
        let workbook = load_xlsx(&xlsx);

        let sheet = &workbook.sheets[0];
        let af = sheet
            .auto_filter
            .as_ref()
            .expect("auto_filter should be present");
        assert_eq!(af.range, "B2:F50");
        assert_eq!(af.start_row, 1);
        assert_eq!(af.start_col, 1);
        assert_eq!(af.end_row, 49);
        assert_eq!(af.end_col, 5);
    }

    #[test]
    fn test_auto_filter_single_column() {
        let xml = r#"<autoFilter ref="A1:A100"/>"#;
        let xlsx = create_xlsx_with_auto_filter(xml);
        let workbook = load_xlsx(&xlsx);

        let sheet = &workbook.sheets[0];
        let af = sheet
            .auto_filter
            .as_ref()
            .expect("auto_filter should be present");
        assert_eq!(af.range, "A1:A100");
        assert_eq!(af.start_col, 0);
        assert_eq!(af.end_col, 0);
    }
}

// ============================================================================
// Test 2: Filter by single value
// ============================================================================

#[cfg(test)]
mod single_value_filter_tests {
    use super::*;

    #[test]
    fn test_filter_by_single_string_value() {
        let xml = r#"<autoFilter ref="A1:D100">
  <filterColumn colId="0">
    <filters>
      <filter val="John"/>
    </filters>
  </filterColumn>
</autoFilter>"#;
        let xlsx = create_xlsx_with_auto_filter(xml);
        let workbook = load_xlsx(&xlsx);

        let sheet = &workbook.sheets[0];
        let af = sheet.auto_filter.as_ref().unwrap();
        assert_eq!(af.filter_columns.len(), 1);
        let fc = &af.filter_columns[0];
        assert_eq!(fc.col_id, 0);
        assert_eq!(fc.filter_type, FilterType::Values);
        assert!(fc.has_filter);
        assert_eq!(fc.values, vec!["John".to_string()]);
    }

    #[test]
    fn test_filter_by_single_numeric_value() {
        let xml = r#"<autoFilter ref="A1:D100">
  <filterColumn colId="1">
    <filters>
      <filter val="25"/>
    </filters>
  </filterColumn>
</autoFilter>"#;
        let xlsx = create_xlsx_with_auto_filter(xml);
        let workbook = load_xlsx(&xlsx);

        let sheet = &workbook.sheets[0];
        let af = sheet.auto_filter.as_ref().unwrap();
        let fc = &af.filter_columns[0];
        assert_eq!(fc.col_id, 1);
        assert_eq!(fc.filter_type, FilterType::Values);
        assert_eq!(fc.values, vec!["25".to_string()]);
    }

    #[test]
    fn test_filter_by_blank_value() {
        let xml = r#"<autoFilter ref="A1:D100">
  <filterColumn colId="0">
    <filters blank="1"/>
  </filterColumn>
</autoFilter>"#;
        let xlsx = create_xlsx_with_auto_filter(xml);
        let workbook = load_xlsx(&xlsx);

        let sheet = &workbook.sheets[0];
        let af = sheet.auto_filter.as_ref().unwrap();
        let fc = &af.filter_columns[0];
        assert_eq!(fc.col_id, 0);
        assert!(fc.has_filter);
    }
}

// ============================================================================
// Test 3: Filter by multiple values
// ============================================================================

#[cfg(test)]
mod multiple_value_filter_tests {
    use super::*;

    #[test]
    fn test_filter_by_multiple_string_values() {
        let xml = r#"<autoFilter ref="A1:D100">
  <filterColumn colId="0">
    <filters>
      <filter val="John"/>
      <filter val="Jane"/>
      <filter val="Bob"/>
    </filters>
  </filterColumn>
</autoFilter>"#;
        let xlsx = create_xlsx_with_auto_filter(xml);
        let workbook = load_xlsx(&xlsx);

        let sheet = &workbook.sheets[0];
        let af = sheet.auto_filter.as_ref().unwrap();
        let fc = &af.filter_columns[0];
        assert_eq!(fc.filter_type, FilterType::Values);
        assert_eq!(fc.values.len(), 3);
        assert!(fc.values.contains(&"John".to_string()));
        assert!(fc.values.contains(&"Jane".to_string()));
        assert!(fc.values.contains(&"Bob".to_string()));
    }

    #[test]
    fn test_filter_by_multiple_numeric_values() {
        let xml = r#"<autoFilter ref="A1:D100">
  <filterColumn colId="1">
    <filters>
      <filter val="20"/>
      <filter val="25"/>
      <filter val="30"/>
    </filters>
  </filterColumn>
</autoFilter>"#;
        let xlsx = create_xlsx_with_auto_filter(xml);
        let workbook = load_xlsx(&xlsx);

        let sheet = &workbook.sheets[0];
        let af = sheet.auto_filter.as_ref().unwrap();
        let fc = &af.filter_columns[0];
        assert_eq!(fc.values.len(), 3);
    }

    #[test]
    fn test_filter_values_including_blank() {
        let xml = r#"<autoFilter ref="A1:D100">
  <filterColumn colId="0">
    <filters blank="1">
      <filter val="Active"/>
      <filter val="Pending"/>
    </filters>
  </filterColumn>
</autoFilter>"#;
        let xlsx = create_xlsx_with_auto_filter(xml);
        let workbook = load_xlsx(&xlsx);

        let sheet = &workbook.sheets[0];
        let af = sheet.auto_filter.as_ref().unwrap();
        let fc = &af.filter_columns[0];
        assert!(fc.has_filter);
        assert!(fc.values.contains(&"Active".to_string()));
        assert!(fc.values.contains(&"Pending".to_string()));
    }
}

// ============================================================================
// Test 4: Custom filters (greater than, less than, contains, etc.)
// ============================================================================

#[cfg(test)]
mod custom_filter_tests {
    use super::*;

    #[test]
    fn test_custom_filter_greater_than() {
        let xml = r#"<autoFilter ref="A1:D100">
  <filterColumn colId="1">
    <customFilters>
      <customFilter operator="greaterThan" val="50"/>
    </customFilters>
  </filterColumn>
</autoFilter>"#;
        let xlsx = create_xlsx_with_auto_filter(xml);
        let workbook = load_xlsx(&xlsx);

        let sheet = &workbook.sheets[0];
        let af = sheet.auto_filter.as_ref().unwrap();
        let fc = &af.filter_columns[0];
        assert_eq!(fc.filter_type, FilterType::Custom);
        assert_eq!(fc.custom_filters.len(), 1);
        assert_eq!(
            fc.custom_filters[0].operator,
            CustomFilterOperator::GreaterThan
        );
        assert_eq!(fc.custom_filters[0].val, "50");
    }

    #[test]
    fn test_custom_filter_less_than() {
        let xml = r#"<autoFilter ref="A1:D100">
  <filterColumn colId="1">
    <customFilters>
      <customFilter operator="lessThan" val="100"/>
    </customFilters>
  </filterColumn>
</autoFilter>"#;
        let xlsx = create_xlsx_with_auto_filter(xml);
        let workbook = load_xlsx(&xlsx);

        let sheet = &workbook.sheets[0];
        let af = sheet.auto_filter.as_ref().unwrap();
        let fc = &af.filter_columns[0];
        assert_eq!(
            fc.custom_filters[0].operator,
            CustomFilterOperator::LessThan
        );
        assert_eq!(fc.custom_filters[0].val, "100");
    }

    #[test]
    fn test_custom_filter_greater_than_or_equal() {
        let xml = r#"<autoFilter ref="A1:D100">
  <filterColumn colId="1">
    <customFilters>
      <customFilter operator="greaterThanOrEqual" val="25"/>
    </customFilters>
  </filterColumn>
</autoFilter>"#;
        let xlsx = create_xlsx_with_auto_filter(xml);
        let workbook = load_xlsx(&xlsx);

        let sheet = &workbook.sheets[0];
        let af = sheet.auto_filter.as_ref().unwrap();
        let fc = &af.filter_columns[0];
        assert_eq!(
            fc.custom_filters[0].operator,
            CustomFilterOperator::GreaterThanOrEqual
        );
    }

    #[test]
    fn test_custom_filter_less_than_or_equal() {
        let xml = r#"<autoFilter ref="A1:D100">
  <filterColumn colId="1">
    <customFilters>
      <customFilter operator="lessThanOrEqual" val="75"/>
    </customFilters>
  </filterColumn>
</autoFilter>"#;
        let xlsx = create_xlsx_with_auto_filter(xml);
        let workbook = load_xlsx(&xlsx);

        let sheet = &workbook.sheets[0];
        let af = sheet.auto_filter.as_ref().unwrap();
        let fc = &af.filter_columns[0];
        assert_eq!(
            fc.custom_filters[0].operator,
            CustomFilterOperator::LessThanOrEqual
        );
    }

    #[test]
    fn test_custom_filter_equal() {
        let xml = r#"<autoFilter ref="A1:D100">
  <filterColumn colId="1">
    <customFilters>
      <customFilter operator="equal" val="50"/>
    </customFilters>
  </filterColumn>
</autoFilter>"#;
        let xlsx = create_xlsx_with_auto_filter(xml);
        let workbook = load_xlsx(&xlsx);

        let sheet = &workbook.sheets[0];
        let af = sheet.auto_filter.as_ref().unwrap();
        let fc = &af.filter_columns[0];
        assert_eq!(fc.custom_filters[0].operator, CustomFilterOperator::Equal);
    }

    #[test]
    fn test_custom_filter_not_equal() {
        let xml = r#"<autoFilter ref="A1:D100">
  <filterColumn colId="1">
    <customFilters>
      <customFilter operator="notEqual" val="0"/>
    </customFilters>
  </filterColumn>
</autoFilter>"#;
        let xlsx = create_xlsx_with_auto_filter(xml);
        let workbook = load_xlsx(&xlsx);

        let sheet = &workbook.sheets[0];
        let af = sheet.auto_filter.as_ref().unwrap();
        let fc = &af.filter_columns[0];
        assert_eq!(
            fc.custom_filters[0].operator,
            CustomFilterOperator::NotEqual
        );
    }

    #[test]
    fn test_custom_filter_between_and() {
        let xml = r#"<autoFilter ref="A1:D100">
  <filterColumn colId="1">
    <customFilters and="1">
      <customFilter operator="greaterThanOrEqual" val="20"/>
      <customFilter operator="lessThanOrEqual" val="50"/>
    </customFilters>
  </filterColumn>
</autoFilter>"#;
        let xlsx = create_xlsx_with_auto_filter(xml);
        let workbook = load_xlsx(&xlsx);

        let sheet = &workbook.sheets[0];
        let af = sheet.auto_filter.as_ref().unwrap();
        let fc = &af.filter_columns[0];
        assert_eq!(fc.custom_filters_and, Some(true));
        assert_eq!(fc.custom_filters.len(), 2);
        assert_eq!(
            fc.custom_filters[0].operator,
            CustomFilterOperator::GreaterThanOrEqual
        );
        assert_eq!(fc.custom_filters[0].val, "20");
        assert_eq!(
            fc.custom_filters[1].operator,
            CustomFilterOperator::LessThanOrEqual
        );
        assert_eq!(fc.custom_filters[1].val, "50");
    }

    #[test]
    fn test_custom_filter_or_logic() {
        let xml = r#"<autoFilter ref="A1:D100">
  <filterColumn colId="1">
    <customFilters>
      <customFilter operator="lessThan" val="10"/>
      <customFilter operator="greaterThan" val="90"/>
    </customFilters>
  </filterColumn>
</autoFilter>"#;
        let xlsx = create_xlsx_with_auto_filter(xml);
        let workbook = load_xlsx(&xlsx);

        let sheet = &workbook.sheets[0];
        let af = sheet.auto_filter.as_ref().unwrap();
        let fc = &af.filter_columns[0];
        // OR logic: custom_filters_and should be None or Some(false)
        assert!(fc.custom_filters_and.is_none() || fc.custom_filters_and == Some(false));
        assert_eq!(fc.custom_filters.len(), 2);
    }
}

// ============================================================================
// Test 5: Filter by color
// ============================================================================

#[cfg(test)]
mod color_filter_tests {
    use super::*;

    #[test]
    fn test_filter_by_cell_color() {
        let xml = r#"<autoFilter ref="A1:D100">
  <filterColumn colId="0">
    <colorFilter dxfId="0"/>
  </filterColumn>
</autoFilter>"#;
        let xlsx = create_xlsx_with_auto_filter(xml);
        let workbook = load_xlsx(&xlsx);

        let sheet = &workbook.sheets[0];
        let af = sheet.auto_filter.as_ref().unwrap();
        let fc = &af.filter_columns[0];
        assert_eq!(fc.filter_type, FilterType::Color);
        assert_eq!(fc.dxf_id, Some(0));
        // cellColor defaults to true
        assert!(fc.cell_color.is_none() || fc.cell_color == Some(true));
    }

    #[test]
    fn test_filter_by_font_color() {
        let xml = r#"<autoFilter ref="A1:D100">
  <filterColumn colId="0">
    <colorFilter dxfId="1" cellColor="0"/>
  </filterColumn>
</autoFilter>"#;
        let xlsx = create_xlsx_with_auto_filter(xml);
        let workbook = load_xlsx(&xlsx);

        let sheet = &workbook.sheets[0];
        let af = sheet.auto_filter.as_ref().unwrap();
        let fc = &af.filter_columns[0];
        assert_eq!(fc.cell_color, Some(false));
    }

    #[test]
    fn test_filter_by_no_fill() {
        let xml = r#"<autoFilter ref="A1:D100">
  <filterColumn colId="0">
    <colorFilter/>
  </filterColumn>
</autoFilter>"#;
        let xlsx = create_xlsx_with_auto_filter(xml);
        let workbook = load_xlsx(&xlsx);

        let sheet = &workbook.sheets[0];
        let af = sheet.auto_filter.as_ref().unwrap();
        let fc = &af.filter_columns[0];
        assert_eq!(fc.filter_type, FilterType::Color);
    }
}

// ============================================================================
// Test 6: Top 10 filter
// ============================================================================

#[cfg(test)]
mod top10_filter_tests {
    use super::*;

    #[test]
    fn test_top_10_items() {
        let xml = r#"<autoFilter ref="A1:D100">
  <filterColumn colId="3">
    <top10 top="1" val="10"/>
  </filterColumn>
</autoFilter>"#;
        let xlsx = create_xlsx_with_auto_filter(xml);
        let workbook = load_xlsx(&xlsx);

        let sheet = &workbook.sheets[0];
        let af = sheet.auto_filter.as_ref().unwrap();
        let fc = &af.filter_columns[0];
        assert_eq!(fc.filter_type, FilterType::Top10);
        assert_eq!(fc.top, Some(true));
        assert!(fc.percent.is_none() || fc.percent == Some(false));
        assert!((fc.top10_val.unwrap() - 10.0).abs() < 0.01);
    }

    #[test]
    fn test_bottom_10_items() {
        let xml = r#"<autoFilter ref="A1:D100">
  <filterColumn colId="3">
    <top10 top="0" val="10"/>
  </filterColumn>
</autoFilter>"#;
        let xlsx = create_xlsx_with_auto_filter(xml);
        let workbook = load_xlsx(&xlsx);

        let sheet = &workbook.sheets[0];
        let af = sheet.auto_filter.as_ref().unwrap();
        let fc = &af.filter_columns[0];
        assert_eq!(fc.top, Some(false));
    }

    #[test]
    fn test_top_5_items() {
        let xml = r#"<autoFilter ref="A1:D100">
  <filterColumn colId="3">
    <top10 top="1" val="5"/>
  </filterColumn>
</autoFilter>"#;
        let xlsx = create_xlsx_with_auto_filter(xml);
        let workbook = load_xlsx(&xlsx);

        let sheet = &workbook.sheets[0];
        let af = sheet.auto_filter.as_ref().unwrap();
        let fc = &af.filter_columns[0];
        assert!((fc.top10_val.unwrap() - 5.0).abs() < 0.01);
    }

    #[test]
    fn test_top_10_percent() {
        let xml = r#"<autoFilter ref="A1:D100">
  <filterColumn colId="3">
    <top10 top="1" percent="1" val="10"/>
  </filterColumn>
</autoFilter>"#;
        let xlsx = create_xlsx_with_auto_filter(xml);
        let workbook = load_xlsx(&xlsx);

        let sheet = &workbook.sheets[0];
        let af = sheet.auto_filter.as_ref().unwrap();
        let fc = &af.filter_columns[0];
        assert_eq!(fc.percent, Some(true));
        assert!((fc.top10_val.unwrap() - 10.0).abs() < 0.01);
    }

    #[test]
    fn test_bottom_25_percent() {
        let xml = r#"<autoFilter ref="A1:D100">
  <filterColumn colId="3">
    <top10 top="0" percent="1" val="25"/>
  </filterColumn>
</autoFilter>"#;
        let xlsx = create_xlsx_with_auto_filter(xml);
        let workbook = load_xlsx(&xlsx);

        let sheet = &workbook.sheets[0];
        let af = sheet.auto_filter.as_ref().unwrap();
        let fc = &af.filter_columns[0];
        assert_eq!(fc.top, Some(false));
        assert_eq!(fc.percent, Some(true));
        assert!((fc.top10_val.unwrap() - 25.0).abs() < 0.01);
    }

    #[test]
    fn test_top_1_item() {
        let xml = r#"<autoFilter ref="A1:D100">
  <filterColumn colId="3">
    <top10 top="1" val="1"/>
  </filterColumn>
</autoFilter>"#;
        let xlsx = create_xlsx_with_auto_filter(xml);
        let workbook = load_xlsx(&xlsx);

        let sheet = &workbook.sheets[0];
        let af = sheet.auto_filter.as_ref().unwrap();
        let fc = &af.filter_columns[0];
        assert!((fc.top10_val.unwrap() - 1.0).abs() < 0.01);
    }
}

// ============================================================================
// Test 7: Date filters (dynamic filters)
// ============================================================================

#[cfg(test)]
mod date_filter_tests {
    use super::*;

    fn assert_dynamic_filter(dynamic_type_xml: &str, expected_type: &str) {
        let xml = format!(
            r#"<autoFilter ref="A1:D100">
  <filterColumn colId="2">
    <dynamicFilter type="{}"/>
  </filterColumn>
</autoFilter>"#,
            dynamic_type_xml
        );
        let xlsx = create_xlsx_with_auto_filter(&xml);
        let workbook = load_xlsx(&xlsx);

        let sheet = &workbook.sheets[0];
        let af = sheet.auto_filter.as_ref().unwrap();
        let fc = &af.filter_columns[0];
        assert_eq!(fc.filter_type, FilterType::Dynamic);
        assert_eq!(fc.dynamic_type.as_deref(), Some(expected_type));
    }

    #[test]
    fn test_dynamic_filter_today() {
        assert_dynamic_filter("today", "today");
    }

    #[test]
    fn test_dynamic_filter_yesterday() {
        assert_dynamic_filter("yesterday", "yesterday");
    }

    #[test]
    fn test_dynamic_filter_tomorrow() {
        assert_dynamic_filter("tomorrow", "tomorrow");
    }

    #[test]
    fn test_dynamic_filter_this_week() {
        assert_dynamic_filter("thisWeek", "thisWeek");
    }

    #[test]
    fn test_dynamic_filter_last_week() {
        assert_dynamic_filter("lastWeek", "lastWeek");
    }

    #[test]
    fn test_dynamic_filter_next_week() {
        assert_dynamic_filter("nextWeek", "nextWeek");
    }

    #[test]
    fn test_dynamic_filter_this_month() {
        assert_dynamic_filter("thisMonth", "thisMonth");
    }

    #[test]
    fn test_dynamic_filter_last_month() {
        assert_dynamic_filter("lastMonth", "lastMonth");
    }

    #[test]
    fn test_dynamic_filter_next_month() {
        assert_dynamic_filter("nextMonth", "nextMonth");
    }

    #[test]
    fn test_dynamic_filter_this_quarter() {
        assert_dynamic_filter("thisQuarter", "thisQuarter");
    }

    #[test]
    fn test_dynamic_filter_this_year() {
        assert_dynamic_filter("thisYear", "thisYear");
    }

    #[test]
    fn test_dynamic_filter_last_year() {
        assert_dynamic_filter("lastYear", "lastYear");
    }

    #[test]
    fn test_dynamic_filter_year_to_date() {
        assert_dynamic_filter("yearToDate", "yearToDate");
    }

    #[test]
    fn test_dynamic_filter_above_average() {
        let xml = r#"<autoFilter ref="A1:D100">
  <filterColumn colId="3">
    <dynamicFilter type="aboveAverage"/>
  </filterColumn>
</autoFilter>"#;
        let xlsx = create_xlsx_with_auto_filter(xml);
        let workbook = load_xlsx(&xlsx);

        let sheet = &workbook.sheets[0];
        let af = sheet.auto_filter.as_ref().unwrap();
        let fc = &af.filter_columns[0];
        assert_eq!(fc.filter_type, FilterType::Dynamic);
        assert_eq!(fc.dynamic_type.as_deref(), Some("aboveAverage"));
    }

    #[test]
    fn test_dynamic_filter_below_average() {
        assert_dynamic_filter("belowAverage", "belowAverage");
    }

    #[test]
    fn test_date_grouping_filter() {
        let xml = r#"<autoFilter ref="A1:D100">
  <filterColumn colId="2">
    <filters>
      <dateGroupItem year="2024" month="1" dateTimeGrouping="month"/>
    </filters>
  </filterColumn>
</autoFilter>"#;
        let xlsx = create_xlsx_with_auto_filter(xml);
        let workbook = load_xlsx(&xlsx);

        let sheet = &workbook.sheets[0];
        let af = sheet.auto_filter.as_ref().unwrap();
        assert!(!af.filter_columns.is_empty());
    }

    #[test]
    fn test_date_grouping_by_year() {
        let xml = r#"<autoFilter ref="A1:D100">
  <filterColumn colId="2">
    <filters>
      <dateGroupItem year="2024" dateTimeGrouping="year"/>
    </filters>
  </filterColumn>
</autoFilter>"#;
        let xlsx = create_xlsx_with_auto_filter(xml);
        let workbook = load_xlsx(&xlsx);

        let sheet = &workbook.sheets[0];
        let af = sheet.auto_filter.as_ref().unwrap();
        assert!(!af.filter_columns.is_empty());
    }
}

// ============================================================================
// Test 8: Text filters (begins with, ends with, contains)
// ============================================================================

#[cfg(test)]
mod text_filter_tests {
    use super::*;

    #[test]
    fn test_text_begins_with() {
        let xml = r#"<autoFilter ref="A1:D100">
  <filterColumn colId="0">
    <customFilters>
      <customFilter operator="equal" val="A*"/>
    </customFilters>
  </filterColumn>
</autoFilter>"#;
        let xlsx = create_xlsx_with_auto_filter(xml);
        let workbook = load_xlsx(&xlsx);

        let sheet = &workbook.sheets[0];
        let af = sheet.auto_filter.as_ref().unwrap();
        let fc = &af.filter_columns[0];
        assert_eq!(fc.filter_type, FilterType::Custom);
        assert_eq!(fc.custom_filters[0].val, "A*");
    }

    #[test]
    fn test_text_ends_with() {
        let xml = r#"<autoFilter ref="A1:D100">
  <filterColumn colId="0">
    <customFilters>
      <customFilter operator="equal" val="*son"/>
    </customFilters>
  </filterColumn>
</autoFilter>"#;
        let xlsx = create_xlsx_with_auto_filter(xml);
        let workbook = load_xlsx(&xlsx);

        let sheet = &workbook.sheets[0];
        let af = sheet.auto_filter.as_ref().unwrap();
        let fc = &af.filter_columns[0];
        assert_eq!(fc.custom_filters[0].val, "*son");
    }

    #[test]
    fn test_text_contains() {
        let xml = r#"<autoFilter ref="A1:D100">
  <filterColumn colId="0">
    <customFilters>
      <customFilter operator="equal" val="*an*"/>
    </customFilters>
  </filterColumn>
</autoFilter>"#;
        let xlsx = create_xlsx_with_auto_filter(xml);
        let workbook = load_xlsx(&xlsx);

        let sheet = &workbook.sheets[0];
        let af = sheet.auto_filter.as_ref().unwrap();
        let fc = &af.filter_columns[0];
        assert_eq!(fc.custom_filters[0].val, "*an*");
    }

    #[test]
    fn test_text_does_not_contain() {
        let xml = r#"<autoFilter ref="A1:D100">
  <filterColumn colId="0">
    <customFilters>
      <customFilter operator="notEqual" val="*test*"/>
    </customFilters>
  </filterColumn>
</autoFilter>"#;
        let xlsx = create_xlsx_with_auto_filter(xml);
        let workbook = load_xlsx(&xlsx);

        let sheet = &workbook.sheets[0];
        let af = sheet.auto_filter.as_ref().unwrap();
        let fc = &af.filter_columns[0];
        assert_eq!(
            fc.custom_filters[0].operator,
            CustomFilterOperator::NotEqual
        );
        assert_eq!(fc.custom_filters[0].val, "*test*");
    }

    #[test]
    fn test_text_does_not_begin_with() {
        let xml = r#"<autoFilter ref="A1:D100">
  <filterColumn colId="0">
    <customFilters>
      <customFilter operator="notEqual" val="X*"/>
    </customFilters>
  </filterColumn>
</autoFilter>"#;
        let xlsx = create_xlsx_with_auto_filter(xml);
        let workbook = load_xlsx(&xlsx);

        let sheet = &workbook.sheets[0];
        let af = sheet.auto_filter.as_ref().unwrap();
        let fc = &af.filter_columns[0];
        assert_eq!(fc.custom_filters[0].val, "X*");
    }

    #[test]
    fn test_text_does_not_end_with() {
        let xml = r#"<autoFilter ref="A1:D100">
  <filterColumn colId="0">
    <customFilters>
      <customFilter operator="notEqual" val="*ing"/>
    </customFilters>
  </filterColumn>
</autoFilter>"#;
        let xlsx = create_xlsx_with_auto_filter(xml);
        let workbook = load_xlsx(&xlsx);

        let sheet = &workbook.sheets[0];
        let af = sheet.auto_filter.as_ref().unwrap();
        let fc = &af.filter_columns[0];
        assert_eq!(fc.custom_filters[0].val, "*ing");
    }

    #[test]
    fn test_text_equals_exact() {
        let xml = r#"<autoFilter ref="A1:D100">
  <filterColumn colId="0">
    <customFilters>
      <customFilter operator="equal" val="Exact Match"/>
    </customFilters>
  </filterColumn>
</autoFilter>"#;
        let xlsx = create_xlsx_with_auto_filter(xml);
        let workbook = load_xlsx(&xlsx);

        let sheet = &workbook.sheets[0];
        let af = sheet.auto_filter.as_ref().unwrap();
        let fc = &af.filter_columns[0];
        assert_eq!(fc.custom_filters[0].val, "Exact Match");
    }

    #[test]
    fn test_text_single_character_wildcard() {
        let xml = r#"<autoFilter ref="A1:D100">
  <filterColumn colId="0">
    <customFilters>
      <customFilter operator="equal" val="Jo?n"/>
    </customFilters>
  </filterColumn>
</autoFilter>"#;
        let xlsx = create_xlsx_with_auto_filter(xml);
        let workbook = load_xlsx(&xlsx);

        let sheet = &workbook.sheets[0];
        let af = sheet.auto_filter.as_ref().unwrap();
        let fc = &af.filter_columns[0];
        assert_eq!(fc.custom_filters[0].val, "Jo?n");
    }
}

// ============================================================================
// Test 9: Number filters
// ============================================================================

#[cfg(test)]
mod number_filter_tests {
    use super::*;

    #[test]
    fn test_number_equals() {
        let xml = r#"<autoFilter ref="A1:D100">
  <filterColumn colId="1">
    <customFilters>
      <customFilter operator="equal" val="100"/>
    </customFilters>
  </filterColumn>
</autoFilter>"#;
        let xlsx = create_xlsx_with_auto_filter(xml);
        let workbook = load_xlsx(&xlsx);

        let sheet = &workbook.sheets[0];
        let af = sheet.auto_filter.as_ref().unwrap();
        let fc = &af.filter_columns[0];
        assert_eq!(fc.custom_filters[0].operator, CustomFilterOperator::Equal);
        assert_eq!(fc.custom_filters[0].val, "100");
    }

    #[test]
    fn test_number_not_equals() {
        let xml = r#"<autoFilter ref="A1:D100">
  <filterColumn colId="1">
    <customFilters>
      <customFilter operator="notEqual" val="0"/>
    </customFilters>
  </filterColumn>
</autoFilter>"#;
        let xlsx = create_xlsx_with_auto_filter(xml);
        let workbook = load_xlsx(&xlsx);

        let sheet = &workbook.sheets[0];
        let af = sheet.auto_filter.as_ref().unwrap();
        let fc = &af.filter_columns[0];
        assert_eq!(
            fc.custom_filters[0].operator,
            CustomFilterOperator::NotEqual
        );
    }

    #[test]
    fn test_number_greater_than() {
        let xml = r#"<autoFilter ref="A1:D100">
  <filterColumn colId="1">
    <customFilters>
      <customFilter operator="greaterThan" val="1000"/>
    </customFilters>
  </filterColumn>
</autoFilter>"#;
        let xlsx = create_xlsx_with_auto_filter(xml);
        let workbook = load_xlsx(&xlsx);

        let sheet = &workbook.sheets[0];
        let af = sheet.auto_filter.as_ref().unwrap();
        let fc = &af.filter_columns[0];
        assert_eq!(
            fc.custom_filters[0].operator,
            CustomFilterOperator::GreaterThan
        );
    }

    #[test]
    fn test_number_less_than() {
        let xml = r#"<autoFilter ref="A1:D100">
  <filterColumn colId="1">
    <customFilters>
      <customFilter operator="lessThan" val="50"/>
    </customFilters>
  </filterColumn>
</autoFilter>"#;
        let xlsx = create_xlsx_with_auto_filter(xml);
        let workbook = load_xlsx(&xlsx);

        let sheet = &workbook.sheets[0];
        let af = sheet.auto_filter.as_ref().unwrap();
        let fc = &af.filter_columns[0];
        assert_eq!(
            fc.custom_filters[0].operator,
            CustomFilterOperator::LessThan
        );
    }

    #[test]
    fn test_number_between() {
        let xml = r#"<autoFilter ref="A1:D100">
  <filterColumn colId="1">
    <customFilters and="1">
      <customFilter operator="greaterThanOrEqual" val="10"/>
      <customFilter operator="lessThanOrEqual" val="100"/>
    </customFilters>
  </filterColumn>
</autoFilter>"#;
        let xlsx = create_xlsx_with_auto_filter(xml);
        let workbook = load_xlsx(&xlsx);

        let sheet = &workbook.sheets[0];
        let af = sheet.auto_filter.as_ref().unwrap();
        let fc = &af.filter_columns[0];
        assert_eq!(fc.custom_filters_and, Some(true));
        assert_eq!(fc.custom_filters.len(), 2);
    }

    #[test]
    fn test_number_not_between() {
        let xml = r#"<autoFilter ref="A1:D100">
  <filterColumn colId="1">
    <customFilters>
      <customFilter operator="lessThan" val="10"/>
      <customFilter operator="greaterThan" val="100"/>
    </customFilters>
  </filterColumn>
</autoFilter>"#;
        let xlsx = create_xlsx_with_auto_filter(xml);
        let workbook = load_xlsx(&xlsx);

        let sheet = &workbook.sheets[0];
        let af = sheet.auto_filter.as_ref().unwrap();
        let fc = &af.filter_columns[0];
        assert_eq!(fc.custom_filters.len(), 2);
    }

    #[test]
    fn test_number_with_decimal() {
        let xml = r#"<autoFilter ref="A1:D100">
  <filterColumn colId="3">
    <customFilters>
      <customFilter operator="greaterThan" val="99.5"/>
    </customFilters>
  </filterColumn>
</autoFilter>"#;
        let xlsx = create_xlsx_with_auto_filter(xml);
        let workbook = load_xlsx(&xlsx);

        let sheet = &workbook.sheets[0];
        let af = sheet.auto_filter.as_ref().unwrap();
        let fc = &af.filter_columns[0];
        assert_eq!(fc.custom_filters[0].val, "99.5");
    }

    #[test]
    fn test_number_negative() {
        let xml = r#"<autoFilter ref="A1:D100">
  <filterColumn colId="3">
    <customFilters>
      <customFilter operator="lessThan" val="-10"/>
    </customFilters>
  </filterColumn>
</autoFilter>"#;
        let xlsx = create_xlsx_with_auto_filter(xml);
        let workbook = load_xlsx(&xlsx);

        let sheet = &workbook.sheets[0];
        let af = sheet.auto_filter.as_ref().unwrap();
        let fc = &af.filter_columns[0];
        assert_eq!(fc.custom_filters[0].val, "-10");
    }
}

// ============================================================================
// Test 10: Multiple columns filtered
// ============================================================================

#[cfg(test)]
mod multiple_column_filter_tests {
    use super::*;

    #[test]
    fn test_two_columns_filtered() {
        let xml = r#"<autoFilter ref="A1:D100">
  <filterColumn colId="0">
    <filters>
      <filter val="Active"/>
    </filters>
  </filterColumn>
  <filterColumn colId="1">
    <customFilters>
      <customFilter operator="greaterThan" val="25"/>
    </customFilters>
  </filterColumn>
</autoFilter>"#;
        let xlsx = create_xlsx_with_auto_filter(xml);
        let workbook = load_xlsx(&xlsx);

        let sheet = &workbook.sheets[0];
        let af = sheet.auto_filter.as_ref().unwrap();
        assert_eq!(af.filter_columns.len(), 2);
    }

    #[test]
    fn test_three_columns_filtered() {
        let xml = r#"<autoFilter ref="A1:D100">
  <filterColumn colId="0">
    <filters>
      <filter val="John"/>
      <filter val="Jane"/>
    </filters>
  </filterColumn>
  <filterColumn colId="1">
    <customFilters>
      <customFilter operator="greaterThanOrEqual" val="18"/>
    </customFilters>
  </filterColumn>
  <filterColumn colId="2">
    <filters>
      <filter val="New York"/>
      <filter val="Los Angeles"/>
    </filters>
  </filterColumn>
</autoFilter>"#;
        let xlsx = create_xlsx_with_auto_filter(xml);
        let workbook = load_xlsx(&xlsx);

        let sheet = &workbook.sheets[0];
        let af = sheet.auto_filter.as_ref().unwrap();
        assert_eq!(af.filter_columns.len(), 3);
    }

    #[test]
    fn test_all_columns_filtered() {
        let xml = r#"<autoFilter ref="A1:D100">
  <filterColumn colId="0">
    <filters>
      <filter val="Smith"/>
    </filters>
  </filterColumn>
  <filterColumn colId="1">
    <top10 top="1" val="10"/>
  </filterColumn>
  <filterColumn colId="2">
    <dynamicFilter type="thisMonth"/>
  </filterColumn>
  <filterColumn colId="3">
    <customFilters>
      <customFilter operator="greaterThan" val="80"/>
    </customFilters>
  </filterColumn>
</autoFilter>"#;
        let xlsx = create_xlsx_with_auto_filter(xml);
        let workbook = load_xlsx(&xlsx);

        let sheet = &workbook.sheets[0];
        let af = sheet.auto_filter.as_ref().unwrap();
        assert_eq!(af.filter_columns.len(), 4);
    }

    #[test]
    fn test_mixed_filter_types() {
        let xml = r#"<autoFilter ref="A1:D100">
  <filterColumn colId="0">
    <colorFilter dxfId="0"/>
  </filterColumn>
  <filterColumn colId="3">
    <top10 top="1" percent="1" val="20"/>
  </filterColumn>
</autoFilter>"#;
        let xlsx = create_xlsx_with_auto_filter(xml);
        let workbook = load_xlsx(&xlsx);

        let sheet = &workbook.sheets[0];
        let af = sheet.auto_filter.as_ref().unwrap();
        assert_eq!(af.filter_columns.len(), 2);
        assert_eq!(af.filter_columns[0].filter_type, FilterType::Color);
        assert_eq!(af.filter_columns[1].filter_type, FilterType::Top10);
    }

    #[test]
    fn test_non_contiguous_column_filters() {
        let xml = r#"<autoFilter ref="A1:D100">
  <filterColumn colId="0">
    <filters>
      <filter val="Category A"/>
    </filters>
  </filterColumn>
  <filterColumn colId="3">
    <filters>
      <filter val="High"/>
    </filters>
  </filterColumn>
</autoFilter>"#;
        let xlsx = create_xlsx_with_auto_filter(xml);
        let workbook = load_xlsx(&xlsx);

        let sheet = &workbook.sheets[0];
        let af = sheet.auto_filter.as_ref().unwrap();
        assert_eq!(af.filter_columns.len(), 2);
        assert_eq!(af.filter_columns[0].col_id, 0);
        assert_eq!(af.filter_columns[1].col_id, 3);
    }
}

// ============================================================================
// Test 11: Filter column with hidden button
// ============================================================================

#[cfg(test)]
mod hidden_button_tests {
    use super::*;

    #[test]
    fn test_hidden_filter_button() {
        let xml = r#"<autoFilter ref="A1:D100">
  <filterColumn colId="0" hiddenButton="1"/>
</autoFilter>"#;
        let xlsx = create_xlsx_with_auto_filter(xml);
        let workbook = load_xlsx(&xlsx);

        let sheet = &workbook.sheets[0];
        let af = sheet.auto_filter.as_ref().unwrap();
        let fc = &af.filter_columns[0];
        assert_eq!(fc.show_button, Some(false));
    }

    #[test]
    fn test_visible_filter_button_explicit() {
        let xml = r#"<autoFilter ref="A1:D100">
  <filterColumn colId="0" hiddenButton="0">
    <filters>
      <filter val="Test"/>
    </filters>
  </filterColumn>
</autoFilter>"#;
        let xlsx = create_xlsx_with_auto_filter(xml);
        let workbook = load_xlsx(&xlsx);

        let sheet = &workbook.sheets[0];
        let af = sheet.auto_filter.as_ref().unwrap();
        let fc = &af.filter_columns[0];
        assert!(fc.show_button.is_none() || fc.show_button == Some(true));
    }

    #[test]
    fn test_multiple_columns_some_hidden() {
        let xml = r#"<autoFilter ref="A1:D100">
  <filterColumn colId="0" hiddenButton="1"/>
  <filterColumn colId="1">
    <filters>
      <filter val="Active"/>
    </filters>
  </filterColumn>
  <filterColumn colId="2" hiddenButton="1"/>
  <filterColumn colId="3">
    <top10 top="1" val="5"/>
  </filterColumn>
</autoFilter>"#;
        let xlsx = create_xlsx_with_auto_filter(xml);
        let workbook = load_xlsx(&xlsx);

        let sheet = &workbook.sheets[0];
        let af = sheet.auto_filter.as_ref().unwrap();
        assert_eq!(af.filter_columns.len(), 4);
    }

    #[test]
    fn test_all_buttons_hidden() {
        let xml = r#"<autoFilter ref="A1:D100">
  <filterColumn colId="0" hiddenButton="1"/>
  <filterColumn colId="1" hiddenButton="1"/>
  <filterColumn colId="2" hiddenButton="1"/>
  <filterColumn colId="3" hiddenButton="1"/>
</autoFilter>"#;
        let xlsx = create_xlsx_with_auto_filter(xml);
        let workbook = load_xlsx(&xlsx);

        let sheet = &workbook.sheets[0];
        let af = sheet.auto_filter.as_ref().unwrap();
        assert_eq!(af.filter_columns.len(), 4);
        for fc in &af.filter_columns {
            assert_eq!(fc.show_button, Some(false));
        }
    }
}

// ============================================================================
// Additional edge case tests
// ============================================================================

#[cfg(test)]
mod edge_case_tests {
    use super::*;

    #[test]
    fn test_empty_filter_column() {
        let xml = r#"<autoFilter ref="A1:D100">
  <filterColumn colId="0"/>
</autoFilter>"#;
        let xlsx = create_xlsx_with_auto_filter(xml);
        let workbook = load_xlsx(&xlsx);

        let sheet = &workbook.sheets[0];
        let af = sheet.auto_filter.as_ref().unwrap();
        let fc = &af.filter_columns[0];
        assert_eq!(fc.filter_type, FilterType::None);
        assert!(!fc.has_filter);
    }

    #[test]
    fn test_icon_filter() {
        let xml = r#"<autoFilter ref="A1:D100">
  <filterColumn colId="0">
    <iconFilter iconSet="3Arrows" iconId="0"/>
  </filterColumn>
</autoFilter>"#;
        let xlsx = create_xlsx_with_auto_filter(xml);
        let workbook = load_xlsx(&xlsx);

        let sheet = &workbook.sheets[0];
        let af = sheet.auto_filter.as_ref().unwrap();
        let fc = &af.filter_columns[0];
        assert_eq!(fc.filter_type, FilterType::Icon);
    }

    #[test]
    fn test_filter_with_special_characters() {
        let xml = r#"<autoFilter ref="A1:D100">
  <filterColumn colId="0">
    <filters>
      <filter val="O&apos;Brien"/>
      <filter val="Smith &amp; Co"/>
      <filter val="&lt;None&gt;"/>
    </filters>
  </filterColumn>
</autoFilter>"#;
        let xlsx = create_xlsx_with_auto_filter(xml);
        let workbook = load_xlsx(&xlsx);

        let sheet = &workbook.sheets[0];
        let af = sheet.auto_filter.as_ref().unwrap();
        let fc = &af.filter_columns[0];
        assert_eq!(fc.values.len(), 3);
    }

    #[test]
    fn test_large_range() {
        let xml = r#"<autoFilter ref="A1:ZZ1000000"/>"#;
        let xlsx = create_xlsx_with_auto_filter(xml);
        let workbook = load_xlsx(&xlsx);

        let sheet = &workbook.sheets[0];
        let af = sheet.auto_filter.as_ref().unwrap();
        assert_eq!(af.range, "A1:ZZ1000000");
    }

    #[test]
    fn test_filter_with_empty_string_value() {
        let xml = r#"<autoFilter ref="A1:D100">
  <filterColumn colId="0">
    <filters>
      <filter val=""/>
    </filters>
  </filterColumn>
</autoFilter>"#;
        let xlsx = create_xlsx_with_auto_filter(xml);
        let workbook = load_xlsx(&xlsx);

        let sheet = &workbook.sheets[0];
        let af = sheet.auto_filter.as_ref().unwrap();
        let fc = &af.filter_columns[0];
        assert!(fc.values.contains(&String::new()));
    }

    #[test]
    fn test_filter_with_unicode_value() {
        let xml = r#"<autoFilter ref="A1:D100">
  <filterColumn colId="0">
    <filters>
      <filter val="Tokyo"/>
      <filter val="Beijing"/>
      <filter val="Seoul"/>
    </filters>
  </filterColumn>
</autoFilter>"#;
        let xlsx = create_xlsx_with_auto_filter(xml);
        let workbook = load_xlsx(&xlsx);

        let sheet = &workbook.sheets[0];
        let af = sheet.auto_filter.as_ref().unwrap();
        let fc = &af.filter_columns[0];
        assert_eq!(fc.values.len(), 3);
    }

    #[test]
    fn test_q1_q4_dynamic_filters() {
        let xml = r#"<autoFilter ref="A1:D100">
  <filterColumn colId="2">
    <dynamicFilter type="Q1"/>
  </filterColumn>
</autoFilter>"#;
        let xlsx = create_xlsx_with_auto_filter(xml);
        let workbook = load_xlsx(&xlsx);

        let sheet = &workbook.sheets[0];
        let af = sheet.auto_filter.as_ref().unwrap();
        let fc = &af.filter_columns[0];
        assert_eq!(fc.filter_type, FilterType::Dynamic);
    }

    #[test]
    fn test_m1_m12_dynamic_filters() {
        let xml = r#"<autoFilter ref="A1:D100">
  <filterColumn colId="2">
    <dynamicFilter type="M1"/>
  </filterColumn>
</autoFilter>"#;
        let xlsx = create_xlsx_with_auto_filter(xml);
        let workbook = load_xlsx(&xlsx);

        let sheet = &workbook.sheets[0];
        let af = sheet.auto_filter.as_ref().unwrap();
        let fc = &af.filter_columns[0];
        assert_eq!(fc.filter_type, FilterType::Dynamic);
    }
}

// ============================================================================
// Documentation tests
// ============================================================================

#[cfg(test)]
mod documentation {
    #[test]
    fn document_auto_filter_structure() {
        // The autoFilter element can appear in worksheet.xml:
        //
        // <autoFilter ref="A1:D100">
        //   <filterColumn colId="0">
        //     <!-- One of the following filter types -->
        //     <!-- 1. Values filter / 2. Custom filter / 3. Top10 filter -->
        //     <!-- 4. Dynamic filter / 5. Color filter / 6. Icon filter -->
        //   </filterColumn>
        // </autoFilter>
    }

    #[test]
    fn document_custom_filter_operators() {
        // Valid operator values for customFilter:
        // equal, notEqual, greaterThan, greaterThanOrEqual, lessThan, lessThanOrEqual
        // Wildcard characters in val: * matches any sequence, ? matches single char
    }

    #[test]
    fn document_dynamic_filter_types() {
        // Valid type values for dynamicFilter:
        // today, yesterday, tomorrow, thisWeek, lastWeek, nextWeek,
        // thisMonth, lastMonth, nextMonth, thisQuarter, lastQuarter, nextQuarter,
        // thisYear, lastYear, nextYear, yearToDate, Q1-Q4, M1-M12,
        // aboveAverage, belowAverage
    }

    #[test]
    fn document_filter_column_attributes() {
        // filterColumn attributes:
        // colId (required): 0-based column index within the autoFilter range
        // hiddenButton: If "1", the dropdown button is hidden
        // showButton: If "0", the dropdown button is hidden (alternative)
    }
}
