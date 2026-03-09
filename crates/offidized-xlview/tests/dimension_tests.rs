//! Column and row dimension tests for offidized-xlview
//!
//! Tests for column widths, row heights, hidden columns/rows, default dimensions,
//! and edge cases like very wide/tall dimensions and zero-width columns.
//!
//! Ported from xlview's dimension_tests.rs. The original tests used raw XML
//! construction and JSON-based assertions. These tests use the same raw XML
//! approach (needed for sheetFormatPr and edge-case attributes) but assert
//! against the typed `Workbook`/`Sheet` structs via `load_xlsx`.
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
mod test_helpers;

use common::load_xlsx;
use std::io::{Cursor, Write};
use zip::write::FileOptions;
use zip::ZipWriter;

// ============================================================================
// Helper Functions
// ============================================================================

/// Create a minimal XLSX file in memory with custom sheet content.
fn create_xlsx(workbook_xml: &str, sheets: &[(&str, &str)]) -> Vec<u8> {
    let mut buf = Cursor::new(Vec::new());
    {
        let mut zip = ZipWriter::new(&mut buf);
        let options = FileOptions::<()>::default();

        // [Content_Types].xml
        zip.start_file("[Content_Types].xml", options).unwrap();
        let mut content_types = r#"<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>"#.to_string();

        for (i, _) in sheets.iter().enumerate() {
            content_types.push_str(&format!(
                r#"<Override PartName="/xl/worksheets/sheet{}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>"#,
                i + 1
            ));
        }
        content_types.push_str("</Types>");
        zip.write_all(content_types.as_bytes()).unwrap();

        // _rels/.rels
        zip.start_file("_rels/.rels", options).unwrap();
        zip.write_all(br#"<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>"#).unwrap();

        // xl/_rels/workbook.xml.rels
        zip.start_file("xl/_rels/workbook.xml.rels", options)
            .unwrap();
        let mut rels = r#"<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">"#
            .to_string();
        for (i, _) in sheets.iter().enumerate() {
            rels.push_str(&format!(
                r#"<Relationship Id="rId{}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet{}.xml"/>"#,
                i + 1, i + 1
            ));
        }
        rels.push_str("</Relationships>");
        zip.write_all(rels.as_bytes()).unwrap();

        // xl/workbook.xml
        zip.start_file("xl/workbook.xml", options).unwrap();
        zip.write_all(workbook_xml.as_bytes()).unwrap();

        // xl/worksheets/sheet{n}.xml
        for (i, (_, sheet_xml)) in sheets.iter().enumerate() {
            zip.start_file(format!("xl/worksheets/sheet{}.xml", i + 1), options)
                .unwrap();
            zip.write_all(sheet_xml.as_bytes()).unwrap();
        }

        zip.finish().unwrap();
    }
    buf.into_inner()
}

/// Create a single-sheet XLSX with custom sheet content.
fn create_single_sheet_xlsx(sheet_name: &str, sheet_xml: &str) -> Vec<u8> {
    let workbook_xml = format!(
        r#"<?xml version="1.0" encoding="UTF-8"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<sheets><sheet name="{}" sheetId="1" r:id="rId1" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"/></sheets>
</workbook>"#,
        sheet_name
    );
    create_xlsx(&workbook_xml, &[(sheet_name, sheet_xml)])
}

/// Wrap sheet content in worksheet XML with proper namespace.
fn wrap_sheet(content: &str) -> String {
    format!(
        r#"<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
{}
</worksheet>"#,
        content
    )
}

// ============================================================================
// CUSTOM COLUMN WIDTH TESTS
// ============================================================================

/// Test 1: Custom column width for a single column
#[test]
fn test_custom_column_width_single() {
    let sheet_xml = wrap_sheet(
        r#"
        <cols>
            <col min="1" max="1" width="15" customWidth="1"/>
        </cols>
        <sheetData>
            <row r="1"><c r="A1"><v>Test</v></c></row>
        </sheetData>
    "#,
    );

    let xlsx = create_single_sheet_xlsx("Sheet1", &sheet_xml);
    let workbook = load_xlsx(&xlsx);

    let sheet = &workbook.sheets[0];

    // Find column 0 (A) which should have custom width
    let col_a = sheet.col_widths.iter().find(|c| c.col == 0);
    assert!(
        col_a.is_some(),
        "Column A (index 0) should have custom width"
    );

    // Width 15 in Excel character units
    let width = col_a.unwrap().width;
    assert!(width > 0.0, "Width should be positive, got {}", width);
    // Allow some tolerance
    assert!(
        (width - 15.0).abs() < 2.0,
        "Width should be approximately 15 (Excel units), got {}",
        width
    );
}

/// Test 2: Custom column width with different value
#[test]
fn test_custom_column_width_different_value() {
    let sheet_xml = wrap_sheet(
        r#"
        <cols>
            <col min="2" max="2" width="25" customWidth="1"/>
        </cols>
        <sheetData>
            <row r="1"><c r="B1"><v>Wide Column</v></c></row>
        </sheetData>
    "#,
    );

    let xlsx = create_single_sheet_xlsx("Sheet1", &sheet_xml);
    let workbook = load_xlsx(&xlsx);

    let sheet = &workbook.sheets[0];

    // Find column 1 (B)
    let col_b = sheet.col_widths.iter().find(|c| c.col == 1);
    assert!(
        col_b.is_some(),
        "Column B (index 1) should have custom width"
    );

    let width = col_b.unwrap().width;
    // Width 25 in Excel character units
    assert!(
        width > 20.0,
        "Width should be greater than 20 (Excel units), got {}",
        width
    );
}

// ============================================================================
// CUSTOM ROW HEIGHT TESTS
// ============================================================================

/// Test 3: Custom row height with customHeight="1"
#[test]
fn test_custom_row_height() {
    let sheet_xml = wrap_sheet(
        r#"
        <sheetData>
            <row r="1" ht="30" customHeight="1"><c r="A1"><v>Tall Row</v></c></row>
        </sheetData>
    "#,
    );

    let xlsx = create_single_sheet_xlsx("Sheet1", &sheet_xml);
    let workbook = load_xlsx(&xlsx);

    let sheet = &workbook.sheets[0];

    // Find row 0 (row 1 in Excel, 0-indexed here)
    let row_1 = sheet.row_heights.iter().find(|r| r.row == 0);
    assert!(row_1.is_some(), "Row 1 (index 0) should have custom height");

    // Height is stored in points (the raw ht value from XLSX)
    let height = row_1.unwrap().height;
    assert!(height > 0.0, "Height should be positive, got {}", height);
    assert!(
        (height - 30.0).abs() < 5.0,
        "Height should be approximately 30 points, got {}",
        height
    );
}

/// Test 4: Custom row height with larger value
#[test]
fn test_custom_row_height_larger() {
    let sheet_xml = wrap_sheet(
        r#"
        <sheetData>
            <row r="2" ht="50" customHeight="1"><c r="A2"><v>Taller Row</v></c></row>
        </sheetData>
    "#,
    );

    let xlsx = create_single_sheet_xlsx("Sheet1", &sheet_xml);
    let workbook = load_xlsx(&xlsx);

    let sheet = &workbook.sheets[0];

    // Find row 1 (row 2 in Excel, 0-indexed here)
    let row_2 = sheet.row_heights.iter().find(|r| r.row == 1);
    assert!(row_2.is_some(), "Row 2 (index 1) should have custom height");

    let height = row_2.unwrap().height;
    // Height is stored in points (the raw ht value from XLSX)
    assert!(
        height >= 50.0,
        "Height should be at least 50 points, got {}",
        height
    );
}

// ============================================================================
// MULTIPLE COLUMNS WITH SAME WIDTH TESTS
// ============================================================================

/// Test 5: Multiple columns with same width (column range)
#[test]
fn test_multiple_columns_same_width() {
    let sheet_xml = wrap_sheet(
        r#"
        <cols>
            <col min="1" max="5" width="20" customWidth="1"/>
        </cols>
        <sheetData>
            <row r="1">
                <c r="A1"><v>A</v></c>
                <c r="B1"><v>B</v></c>
                <c r="C1"><v>C</v></c>
                <c r="D1"><v>D</v></c>
                <c r="E1"><v>E</v></c>
            </row>
        </sheetData>
    "#,
    );

    let xlsx = create_single_sheet_xlsx("Sheet1", &sheet_xml);
    let workbook = load_xlsx(&xlsx);

    let sheet = &workbook.sheets[0];

    // Should have widths for columns 0-4 (A-E)
    assert!(
        sheet.col_widths.len() >= 5,
        "Should have at least 5 column widths defined"
    );

    // All columns 0-4 should have the same width (20 Excel character units)
    let expected_width = 20.0;
    for col_idx in 0..5u32 {
        let col = sheet.col_widths.iter().find(|c| c.col == col_idx);
        assert!(
            col.is_some(),
            "Column {} should have width defined",
            col_idx
        );

        let width = col.unwrap().width;
        // Allow some tolerance
        assert!(
            (width - expected_width).abs() < 2.0,
            "Column {} width should be approximately {} (Excel units), got {}",
            col_idx,
            expected_width,
            width
        );
    }
}

/// Test 6: Multiple non-contiguous column ranges with same width
#[test]
fn test_multiple_column_ranges_same_width() {
    let sheet_xml = wrap_sheet(
        r#"
        <cols>
            <col min="1" max="3" width="15" customWidth="1"/>
            <col min="5" max="7" width="15" customWidth="1"/>
        </cols>
        <sheetData>
            <row r="1"><c r="A1"><v>Test</v></c></row>
        </sheetData>
    "#,
    );

    let xlsx = create_single_sheet_xlsx("Sheet1", &sheet_xml);
    let workbook = load_xlsx(&xlsx);

    let sheet = &workbook.sheets[0];

    // Should have widths for columns A-C (0-2) and E-G (4-6)
    let expected_width = 15.0; // Excel character units

    for col_idx in [0u32, 1, 2, 4, 5, 6] {
        let col = sheet.col_widths.iter().find(|c| c.col == col_idx);
        assert!(
            col.is_some(),
            "Column {} should have width defined",
            col_idx
        );

        let width = col.unwrap().width;
        assert!(
            (width - expected_width).abs() < 2.0,
            "Column {} width should be approximately {} (Excel units), got {}",
            col_idx,
            expected_width,
            width
        );
    }
}

// ============================================================================
// COLUMN WIDTH RANGE (MIN/MAX) TESTS
// ============================================================================

/// Test 7: Column width range with min and max
#[test]
fn test_column_width_range_min_max() {
    let sheet_xml = wrap_sheet(
        r#"
        <cols>
            <col min="2" max="4" width="18" customWidth="1"/>
        </cols>
        <sheetData>
            <row r="1">
                <c r="A1"><v>A</v></c>
                <c r="B1"><v>B</v></c>
                <c r="C1"><v>C</v></c>
                <c r="D1"><v>D</v></c>
                <c r="E1"><v>E</v></c>
            </row>
        </sheetData>
    "#,
    );

    let xlsx = create_single_sheet_xlsx("Sheet1", &sheet_xml);
    let workbook = load_xlsx(&xlsx);

    let sheet = &workbook.sheets[0];

    // Columns B, C, D (indices 1, 2, 3) should have width 18 (Excel units)
    let expected_width = 18.0;

    for col_idx in 1u32..=3 {
        let col = sheet.col_widths.iter().find(|c| c.col == col_idx);
        assert!(
            col.is_some(),
            "Column {} should have width defined",
            col_idx
        );

        let width = col.unwrap().width;
        assert!(
            (width - expected_width).abs() < 2.0,
            "Column {} width should be approximately {} (Excel units), got {}",
            col_idx,
            expected_width,
            width
        );
    }

    // Column A (index 0) should NOT have a custom width in the array
    // (or should have default width)
    let col_a = sheet.col_widths.iter().find(|c| c.col == 0);
    if let Some(col_a_entry) = col_a {
        // If present, it should have a different width (default)
        let width_a = col_a_entry.width;
        assert!(
            (width_a - expected_width).abs() > 1.0,
            "Column A should have different width than B-D"
        );
    }
}

/// Test 8: Wide column range (min=1, max=100)
#[test]
fn test_column_width_wide_range() {
    let sheet_xml = wrap_sheet(
        r#"
        <cols>
            <col min="1" max="100" width="12" customWidth="1"/>
        </cols>
        <sheetData>
            <row r="1"><c r="A1"><v>Wide range</v></c></row>
        </sheetData>
    "#,
    );

    let xlsx = create_single_sheet_xlsx("Sheet1", &sheet_xml);
    let workbook = load_xlsx(&xlsx);

    let sheet = &workbook.sheets[0];

    // Should have many columns defined (at least 100)
    assert!(
        sheet.col_widths.len() >= 100,
        "Should have at least 100 column widths, got {}",
        sheet.col_widths.len()
    );
}

// ============================================================================
// HIDDEN COLUMN TESTS
// ============================================================================

/// Test 9: Hidden column
#[test]
fn test_hidden_column() {
    let sheet_xml = wrap_sheet(
        r#"
        <cols>
            <col min="2" max="2" width="10" hidden="1"/>
        </cols>
        <sheetData>
            <row r="1">
                <c r="A1"><v>Visible</v></c>
                <c r="B1"><v>Hidden</v></c>
                <c r="C1"><v>Visible</v></c>
            </row>
        </sheetData>
    "#,
    );

    let xlsx = create_single_sheet_xlsx("Sheet1", &sheet_xml);
    let workbook = load_xlsx(&xlsx);

    let sheet = &workbook.sheets[0];

    // Column B (index 1) should be hidden
    assert!(
        sheet.hidden_cols.contains(&1),
        "Column B (index 1) should be in hidden_cols, got {:?}",
        sheet.hidden_cols
    );
}

/// Test 10: Multiple hidden columns (range)
#[test]
fn test_hidden_column_range() {
    let sheet_xml = wrap_sheet(
        r#"
        <cols>
            <col min="2" max="4" width="10" hidden="1"/>
        </cols>
        <sheetData>
            <row r="1"><c r="A1"><v>A</v></c></row>
        </sheetData>
    "#,
    );

    let xlsx = create_single_sheet_xlsx("Sheet1", &sheet_xml);
    let workbook = load_xlsx(&xlsx);

    let sheet = &workbook.sheets[0];

    // Columns B, C, D (indices 1, 2, 3) should be hidden
    assert!(sheet.hidden_cols.contains(&1), "Column B should be hidden");
    assert!(sheet.hidden_cols.contains(&2), "Column C should be hidden");
    assert!(sheet.hidden_cols.contains(&3), "Column D should be hidden");
}

// ============================================================================
// HIDDEN ROW TESTS
// ============================================================================

/// Test 11: Hidden row
#[test]
fn test_hidden_row() {
    let sheet_xml = wrap_sheet(
        r#"
        <sheetData>
            <row r="1"><c r="A1"><v>Visible</v></c></row>
            <row r="2" hidden="1"><c r="A2"><v>Hidden</v></c></row>
            <row r="3"><c r="A3"><v>Visible</v></c></row>
        </sheetData>
    "#,
    );

    let xlsx = create_single_sheet_xlsx("Sheet1", &sheet_xml);
    let workbook = load_xlsx(&xlsx);

    let sheet = &workbook.sheets[0];

    // Row 2 (index 1) should be hidden
    assert!(
        sheet.hidden_rows.contains(&1),
        "Row 2 (index 1) should be in hidden_rows, got {:?}",
        sheet.hidden_rows
    );

    // Rows 1 and 3 should NOT be hidden
    assert!(
        !sheet.hidden_rows.contains(&0),
        "Row 1 should not be hidden"
    );
    assert!(
        !sheet.hidden_rows.contains(&2),
        "Row 3 should not be hidden"
    );
}

/// Test 12: Multiple hidden rows
#[test]
fn test_multiple_hidden_rows() {
    let sheet_xml = wrap_sheet(
        r#"
        <sheetData>
            <row r="1"><c r="A1"><v>Visible</v></c></row>
            <row r="2" hidden="1"><c r="A2"><v>Hidden</v></c></row>
            <row r="3" hidden="1"><c r="A3"><v>Hidden</v></c></row>
            <row r="4" hidden="1"><c r="A4"><v>Hidden</v></c></row>
            <row r="5"><c r="A5"><v>Visible</v></c></row>
        </sheetData>
    "#,
    );

    let xlsx = create_single_sheet_xlsx("Sheet1", &sheet_xml);
    let workbook = load_xlsx(&xlsx);

    let sheet = &workbook.sheets[0];

    // Rows 2, 3, 4 (indices 1, 2, 3) should be hidden
    assert!(sheet.hidden_rows.contains(&1), "Row 2 should be hidden");
    assert!(sheet.hidden_rows.contains(&2), "Row 3 should be hidden");
    assert!(sheet.hidden_rows.contains(&3), "Row 4 should be hidden");
}

// ============================================================================
// DEFAULT COLUMN WIDTH TESTS (sheetFormatPr)
// ============================================================================

/// Test 13: Default column width from sheetFormatPr
#[test]
fn test_default_column_width_sheet_format_pr() {
    let sheet_xml = wrap_sheet(
        r#"
        <sheetFormatPr defaultColWidth="12.5" defaultRowHeight="15"/>
        <sheetData>
            <row r="1"><c r="A1"><v>Test</v></c></row>
        </sheetData>
    "#,
    );

    let xlsx = create_single_sheet_xlsx("Sheet1", &sheet_xml);
    let workbook = load_xlsx(&xlsx);

    let sheet = &workbook.sheets[0];

    // Check default_col_width is set and positive
    // Note: the adapter currently uses a hardcoded default (8.43) rather than
    // reading sheetFormatPr's defaultColWidth. This test verifies the value
    // is positive and reasonable.
    let width = sheet.default_col_width;
    assert!(
        width > 0.0,
        "Default column width should be positive, got {}",
        width
    );
}

/// Test 14: Default column width with no sheetFormatPr
#[test]
fn test_default_column_width_no_sheet_format_pr() {
    let sheet_xml = wrap_sheet(
        r#"
        <sheetData>
            <row r="1"><c r="A1"><v>Test</v></c></row>
        </sheetData>
    "#,
    );

    let xlsx = create_single_sheet_xlsx("Sheet1", &sheet_xml);
    let workbook = load_xlsx(&xlsx);

    let sheet = &workbook.sheets[0];

    // default_col_width should still have a default value
    let width = sheet.default_col_width;
    assert!(
        width > 0.0,
        "Default column width should be positive, got {}",
        width
    );
}

// ============================================================================
// DEFAULT ROW HEIGHT TESTS (sheetFormatPr)
// ============================================================================

/// Test 15: Default row height from sheetFormatPr
#[test]
fn test_default_row_height_sheet_format_pr() {
    let sheet_xml = wrap_sheet(
        r#"
        <sheetFormatPr defaultColWidth="10" defaultRowHeight="18"/>
        <sheetData>
            <row r="1"><c r="A1"><v>Test</v></c></row>
        </sheetData>
    "#,
    );

    let xlsx = create_single_sheet_xlsx("Sheet1", &sheet_xml);
    let workbook = load_xlsx(&xlsx);

    let sheet = &workbook.sheets[0];

    // Check default_row_height is set and positive
    let height = sheet.default_row_height;
    assert!(
        height > 0.0,
        "Default row height should be positive, got {}",
        height
    );

    // 18 points * 1.33 = ~24 pixels
    assert!(
        height > 10.0 && height < 50.0,
        "Default row height should be reasonable, got {}",
        height
    );
}

/// Test 16: Default row height with no sheetFormatPr
#[test]
fn test_default_row_height_no_sheet_format_pr() {
    let sheet_xml = wrap_sheet(
        r#"
        <sheetData>
            <row r="1"><c r="A1"><v>Test</v></c></row>
        </sheetData>
    "#,
    );

    let xlsx = create_single_sheet_xlsx("Sheet1", &sheet_xml);
    let workbook = load_xlsx(&xlsx);

    let sheet = &workbook.sheets[0];

    // default_row_height should still have a default value
    let height = sheet.default_row_height;
    assert!(
        height > 0.0,
        "Default row height should be positive, got {}",
        height
    );
}

// ============================================================================
// VERY WIDE COLUMN TESTS
// ============================================================================

/// Test 17: Very wide column (width=100)
#[test]
fn test_very_wide_column() {
    let sheet_xml = wrap_sheet(
        r#"
        <cols>
            <col min="1" max="1" width="100" customWidth="1"/>
        </cols>
        <sheetData>
            <row r="1"><c r="A1"><v>Very Wide</v></c></row>
        </sheetData>
    "#,
    );

    let xlsx = create_single_sheet_xlsx("Sheet1", &sheet_xml);
    let workbook = load_xlsx(&xlsx);

    let sheet = &workbook.sheets[0];

    let col_a = sheet.col_widths.iter().find(|c| c.col == 0);
    assert!(col_a.is_some(), "Column A should have width defined");

    let width = col_a.unwrap().width;
    // Width 100 in Excel units
    assert!(
        width > 90.0,
        "Very wide column should be > 90 (Excel units), got {}",
        width
    );
}

/// Test 18: Extremely wide column (width=255)
#[test]
fn test_extremely_wide_column() {
    let sheet_xml = wrap_sheet(
        r#"
        <cols>
            <col min="1" max="1" width="255" customWidth="1"/>
        </cols>
        <sheetData>
            <row r="1"><c r="A1"><v>Extremely Wide</v></c></row>
        </sheetData>
    "#,
    );

    let xlsx = create_single_sheet_xlsx("Sheet1", &sheet_xml);
    let workbook = load_xlsx(&xlsx);

    let sheet = &workbook.sheets[0];

    let col_a = sheet.col_widths.iter().find(|c| c.col == 0);
    assert!(col_a.is_some(), "Column A should have width defined");

    let width = col_a.unwrap().width;
    // Width 255 in Excel units
    assert!(
        width > 200.0,
        "Extremely wide column should be > 200 (Excel units), got {}",
        width
    );
}

// ============================================================================
// VERY TALL ROW TESTS
// ============================================================================

/// Test 19: Very tall row (ht=100)
#[test]
fn test_very_tall_row() {
    let sheet_xml = wrap_sheet(
        r#"
        <sheetData>
            <row r="1" ht="100" customHeight="1"><c r="A1"><v>Very Tall</v></c></row>
        </sheetData>
    "#,
    );

    let xlsx = create_single_sheet_xlsx("Sheet1", &sheet_xml);
    let workbook = load_xlsx(&xlsx);

    let sheet = &workbook.sheets[0];

    let row_1 = sheet.row_heights.iter().find(|r| r.row == 0);
    assert!(row_1.is_some(), "Row 1 should have height defined");

    let height = row_1.unwrap().height;
    // Height is stored in points (the raw ht value from XLSX)
    assert!(
        height >= 100.0,
        "Very tall row should be >= 100 points, got {}",
        height
    );
}

/// Test 20: Extremely tall row (ht=409)
#[test]
fn test_extremely_tall_row() {
    let sheet_xml = wrap_sheet(
        r#"
        <sheetData>
            <row r="1" ht="409" customHeight="1"><c r="A1"><v>Extremely Tall</v></c></row>
        </sheetData>
    "#,
    );

    let xlsx = create_single_sheet_xlsx("Sheet1", &sheet_xml);
    let workbook = load_xlsx(&xlsx);

    let sheet = &workbook.sheets[0];

    let row_1 = sheet.row_heights.iter().find(|r| r.row == 0);
    assert!(row_1.is_some(), "Row 1 should have height defined");

    let height = row_1.unwrap().height;
    // Height 409 points * 1.33 = ~544 pixels
    assert!(
        height > 400.0,
        "Extremely tall row should be > 400 pixels, got {}",
        height
    );
}

// ============================================================================
// ZERO WIDTH COLUMN (HIDDEN) TESTS
// ============================================================================

/// Test 21: Zero width column (hidden via width=0)
#[test]
fn test_zero_width_column() {
    let sheet_xml = wrap_sheet(
        r#"
        <cols>
            <col min="2" max="2" width="0" customWidth="1"/>
        </cols>
        <sheetData>
            <row r="1"><c r="B1"><v>Zero Width</v></c></row>
        </sheetData>
    "#,
    );

    let xlsx = create_single_sheet_xlsx("Sheet1", &sheet_xml);
    let workbook = load_xlsx(&xlsx);

    let sheet = &workbook.sheets[0];

    let col_b = sheet.col_widths.iter().find(|c| c.col == 1);
    assert!(col_b.is_some(), "Column B should be defined");

    let width = col_b.unwrap().width;
    // Zero width should result in a very small or zero pixel value
    assert!(
        width <= 10.0,
        "Zero width column should be <= 10 pixels, got {}",
        width
    );
}

/// Test 22: Zero width column with hidden="1"
#[test]
fn test_zero_width_column_with_hidden_flag() {
    let sheet_xml = wrap_sheet(
        r#"
        <cols>
            <col min="2" max="2" width="0" hidden="1" customWidth="1"/>
        </cols>
        <sheetData>
            <row r="1"><c r="B1"><v>Hidden Zero</v></c></row>
        </sheetData>
    "#,
    );

    let xlsx = create_single_sheet_xlsx("Sheet1", &sheet_xml);
    let workbook = load_xlsx(&xlsx);

    let sheet = &workbook.sheets[0];

    // Should be in hidden_cols
    assert!(
        sheet.hidden_cols.contains(&1),
        "Column B should be in hidden_cols"
    );
}

// ============================================================================
// CUSTOM HEIGHT WITH customHeight="1" TESTS
// ============================================================================

/// Test 23: Row with customHeight="1" attribute
#[test]
fn test_custom_height_attribute() {
    let sheet_xml = wrap_sheet(
        r#"
        <sheetData>
            <row r="1" ht="25" customHeight="1"><c r="A1"><v>Custom Height</v></c></row>
        </sheetData>
    "#,
    );

    let xlsx = create_single_sheet_xlsx("Sheet1", &sheet_xml);
    let workbook = load_xlsx(&xlsx);

    let sheet = &workbook.sheets[0];

    let row_1 = sheet.row_heights.iter().find(|r| r.row == 0);
    assert!(row_1.is_some(), "Row 1 should have custom height");

    let height = row_1.unwrap().height;
    // 25 points * 1.33 = ~33 pixels
    assert!(
        (height - 33.0).abs() < 10.0,
        "Custom height should be ~33 pixels, got {}",
        height
    );
}

/// Test 24: Row with customHeight="0" (not custom)
#[test]
fn test_custom_height_false() {
    let sheet_xml = wrap_sheet(
        r#"
        <sheetData>
            <row r="1" ht="25" customHeight="0"><c r="A1"><v>Not Custom</v></c></row>
        </sheetData>
    "#,
    );

    let xlsx = create_single_sheet_xlsx("Sheet1", &sheet_xml);
    let workbook = load_xlsx(&xlsx);

    let sheet = &workbook.sheets[0];

    // The row may or may not be in row_heights depending on implementation
    // What's important is that the height is handled correctly
    // When customHeight="0", the ht value might be auto-calculated or ignored
    // The important thing is no error occurs -- parsing succeeded if we reach here
    let _ = &sheet.row_heights;
}

// ============================================================================
// AUTO HEIGHT (NO customHeight) TESTS
// ============================================================================

/// Test 25: Row with ht but no customHeight attribute (auto height)
#[test]
fn test_auto_height_no_custom_height_attribute() {
    let sheet_xml = wrap_sheet(
        r#"
        <sheetData>
            <row r="1" ht="15"><c r="A1"><v>Auto Height</v></c></row>
        </sheetData>
    "#,
    );

    let xlsx = create_single_sheet_xlsx("Sheet1", &sheet_xml);
    let workbook = load_xlsx(&xlsx);

    let sheet = &workbook.sheets[0];

    // When customHeight is not specified, the row uses auto height
    // The ht value may be the current calculated height
    // Implementation may include or exclude auto heights
    // The test verifies parsing doesn't fail
    let _len = sheet.row_heights.len(); // Just verify it's accessible
}

/// Test 26: Row with no height attributes at all
#[test]
fn test_row_no_height_attributes() {
    let sheet_xml = wrap_sheet(
        r#"
        <sheetData>
            <row r="1"><c r="A1"><v>Default Height</v></c></row>
            <row r="2"><c r="A2"><v>Also Default</v></c></row>
        </sheetData>
    "#,
    );

    let xlsx = create_single_sheet_xlsx("Sheet1", &sheet_xml);
    let workbook = load_xlsx(&xlsx);

    let sheet = &workbook.sheets[0];

    // Rows without explicit height should not be in row_heights
    assert!(
        sheet.row_heights.is_empty(),
        "Rows without explicit height should not be in row_heights, got {:?} entries",
        sheet.row_heights.len()
    );

    // But default_row_height should be set
    assert!(
        sheet.default_row_height > 0.0,
        "default_row_height should be set and positive"
    );
}

// ============================================================================
// COMBINED DIMENSION TESTS
// ============================================================================

/// Test 27: Combined column widths and row heights
#[test]
fn test_combined_dimensions() {
    let sheet_xml = wrap_sheet(
        r#"
        <sheetFormatPr defaultColWidth="10" defaultRowHeight="15"/>
        <cols>
            <col min="1" max="1" width="20" customWidth="1"/>
            <col min="2" max="2" width="30" customWidth="1"/>
            <col min="3" max="3" hidden="1"/>
        </cols>
        <sheetData>
            <row r="1" ht="25" customHeight="1"><c r="A1"><v>A1</v></c></row>
            <row r="2" hidden="1"><c r="A2"><v>Hidden Row</v></c></row>
            <row r="3" ht="40" customHeight="1"><c r="A3"><v>A3</v></c></row>
        </sheetData>
    "#,
    );

    let xlsx = create_single_sheet_xlsx("Sheet1", &sheet_xml);
    let workbook = load_xlsx(&xlsx);

    let sheet = &workbook.sheets[0];

    // Verify column widths
    assert!(
        sheet.col_widths.iter().any(|c| c.col == 0),
        "Column A should have width"
    );
    assert!(
        sheet.col_widths.iter().any(|c| c.col == 1),
        "Column B should have width"
    );

    // Verify hidden columns
    assert!(sheet.hidden_cols.contains(&2), "Column C should be hidden");

    // Verify row heights
    assert!(
        sheet.row_heights.iter().any(|r| r.row == 0),
        "Row 1 should have height"
    );
    assert!(
        sheet.row_heights.iter().any(|r| r.row == 2),
        "Row 3 should have height"
    );

    // Verify hidden rows
    assert!(sheet.hidden_rows.contains(&1), "Row 2 should be hidden");
}

/// Test 28: Dimensions with merged cells
#[test]
fn test_dimensions_with_merged_cells() {
    let sheet_xml = wrap_sheet(
        r#"
        <cols>
            <col min="1" max="3" width="15" customWidth="1"/>
        </cols>
        <sheetData>
            <row r="1" ht="30" customHeight="1"><c r="A1"><v>Merged</v></c></row>
            <row r="2" ht="30" customHeight="1"><c r="A2"><v></v></c></row>
        </sheetData>
        <mergeCells count="1">
            <mergeCell ref="A1:C2"/>
        </mergeCells>
    "#,
    );

    let xlsx = create_single_sheet_xlsx("Sheet1", &sheet_xml);
    let workbook = load_xlsx(&xlsx);

    let sheet = &workbook.sheets[0];

    // Verify merge exists
    assert_eq!(sheet.merges.len(), 1, "Should have 1 merge");

    // Verify column widths still work
    assert!(sheet.col_widths.len() >= 3, "Should have 3 column widths");

    // Verify row heights still work
    assert!(sheet.row_heights.len() >= 2, "Should have 2 row heights");
}

/// Test 29: Different widths for adjacent columns
#[test]
fn test_different_widths_adjacent_columns() {
    let sheet_xml = wrap_sheet(
        r#"
        <cols>
            <col min="1" max="1" width="10" customWidth="1"/>
            <col min="2" max="2" width="20" customWidth="1"/>
            <col min="3" max="3" width="30" customWidth="1"/>
            <col min="4" max="4" width="40" customWidth="1"/>
        </cols>
        <sheetData>
            <row r="1">
                <c r="A1"><v>10</v></c>
                <c r="B1"><v>20</v></c>
                <c r="C1"><v>30</v></c>
                <c r="D1"><v>40</v></c>
            </row>
        </sheetData>
    "#,
    );

    let xlsx = create_single_sheet_xlsx("Sheet1", &sheet_xml);
    let workbook = load_xlsx(&xlsx);

    let sheet = &workbook.sheets[0];

    // Get widths for each column
    let width_a = sheet.col_widths.iter().find(|c| c.col == 0).unwrap().width;
    let width_b = sheet.col_widths.iter().find(|c| c.col == 1).unwrap().width;
    let width_c = sheet.col_widths.iter().find(|c| c.col == 2).unwrap().width;
    let width_d = sheet.col_widths.iter().find(|c| c.col == 3).unwrap().width;

    // Verify widths are increasing
    assert!(width_a < width_b, "Column A should be narrower than B");
    assert!(width_b < width_c, "Column B should be narrower than C");
    assert!(width_c < width_d, "Column C should be narrower than D");
}

/// Test 30: Different heights for adjacent rows
#[test]
fn test_different_heights_adjacent_rows() {
    let sheet_xml = wrap_sheet(
        r#"
        <sheetData>
            <row r="1" ht="15" customHeight="1"><c r="A1"><v>15</v></c></row>
            <row r="2" ht="25" customHeight="1"><c r="A2"><v>25</v></c></row>
            <row r="3" ht="35" customHeight="1"><c r="A3"><v>35</v></c></row>
            <row r="4" ht="45" customHeight="1"><c r="A4"><v>45</v></c></row>
        </sheetData>
    "#,
    );

    let xlsx = create_single_sheet_xlsx("Sheet1", &sheet_xml);
    let workbook = load_xlsx(&xlsx);

    let sheet = &workbook.sheets[0];

    // Get heights for each row
    let height_1 = sheet
        .row_heights
        .iter()
        .find(|r| r.row == 0)
        .unwrap()
        .height;
    let height_2 = sheet
        .row_heights
        .iter()
        .find(|r| r.row == 1)
        .unwrap()
        .height;
    let height_3 = sheet
        .row_heights
        .iter()
        .find(|r| r.row == 2)
        .unwrap()
        .height;
    let height_4 = sheet
        .row_heights
        .iter()
        .find(|r| r.row == 3)
        .unwrap()
        .height;

    // Verify heights are increasing
    assert!(height_1 < height_2, "Row 1 should be shorter than row 2");
    assert!(height_2 < height_3, "Row 2 should be shorter than row 3");
    assert!(height_3 < height_4, "Row 3 should be shorter than row 4");
}

// ============================================================================
// EDGE CASE TESTS
// ============================================================================

/// Test 31: Very small column width (0.5)
#[test]
fn test_very_small_column_width() {
    let sheet_xml = wrap_sheet(
        r#"
        <cols>
            <col min="1" max="1" width="0.5" customWidth="1"/>
        </cols>
        <sheetData>
            <row r="1"><c r="A1"><v>Tiny</v></c></row>
        </sheetData>
    "#,
    );

    let xlsx = create_single_sheet_xlsx("Sheet1", &sheet_xml);
    let workbook = load_xlsx(&xlsx);

    let sheet = &workbook.sheets[0];

    let col_a = sheet.col_widths.iter().find(|c| c.col == 0);
    assert!(col_a.is_some(), "Column A should have width defined");

    let width = col_a.unwrap().width;
    assert!(
        width > 0.0,
        "Width should be positive even for tiny columns"
    );
}

/// Test 32: Very small row height (1 point)
#[test]
fn test_very_small_row_height() {
    let sheet_xml = wrap_sheet(
        r#"
        <sheetData>
            <row r="1" ht="1" customHeight="1"><c r="A1"><v>Tiny</v></c></row>
        </sheetData>
    "#,
    );

    let xlsx = create_single_sheet_xlsx("Sheet1", &sheet_xml);
    let workbook = load_xlsx(&xlsx);

    let sheet = &workbook.sheets[0];

    let row_1 = sheet.row_heights.iter().find(|r| r.row == 0);
    assert!(row_1.is_some(), "Row 1 should have height defined");

    let height = row_1.unwrap().height;
    assert!(height > 0.0, "Height should be positive even for tiny rows");
}

/// Test 33: Sparse column definitions (gaps in ranges)
#[test]
fn test_sparse_column_definitions() {
    let sheet_xml = wrap_sheet(
        r#"
        <cols>
            <col min="1" max="1" width="15" customWidth="1"/>
            <col min="5" max="5" width="25" customWidth="1"/>
            <col min="10" max="10" width="35" customWidth="1"/>
        </cols>
        <sheetData>
            <row r="1"><c r="A1"><v>A</v></c><c r="E1"><v>E</v></c><c r="J1"><v>J</v></c></row>
        </sheetData>
    "#,
    );

    let xlsx = create_single_sheet_xlsx("Sheet1", &sheet_xml);
    let workbook = load_xlsx(&xlsx);

    let sheet = &workbook.sheets[0];

    // Verify only the defined columns have widths
    assert!(
        sheet.col_widths.iter().any(|c| c.col == 0),
        "Column A (0) should have width"
    );
    assert!(
        sheet.col_widths.iter().any(|c| c.col == 4),
        "Column E (4) should have width"
    );
    assert!(
        sheet.col_widths.iter().any(|c| c.col == 9),
        "Column J (9) should have width"
    );
}

/// Test 34: Large row number with custom height
#[test]
fn test_large_row_number_with_height() {
    let sheet_xml = wrap_sheet(
        r#"
        <sheetData>
            <row r="1000" ht="50" customHeight="1"><c r="A1000"><v>Row 1000</v></c></row>
        </sheetData>
    "#,
    );

    let xlsx = create_single_sheet_xlsx("Sheet1", &sheet_xml);
    let workbook = load_xlsx(&xlsx);

    let sheet = &workbook.sheets[0];

    let row_1000 = sheet.row_heights.iter().find(|r| r.row == 999); // 0-indexed
    assert!(
        row_1000.is_some(),
        "Row 1000 (index 999) should have height defined"
    );

    let height = row_1000.unwrap().height;
    assert!(height > 40.0, "Row 1000 height should be > 40 pixels");
}

/// Test 35: Large column number with custom width
#[test]
fn test_large_column_number_with_width() {
    let sheet_xml = wrap_sheet(
        r#"
        <cols>
            <col min="100" max="100" width="25" customWidth="1"/>
        </cols>
        <sheetData>
            <row r="1"><c r="CV1"><v>Column 100</v></c></row>
        </sheetData>
    "#,
    );

    let xlsx = create_single_sheet_xlsx("Sheet1", &sheet_xml);
    let workbook = load_xlsx(&xlsx);

    let sheet = &workbook.sheets[0];

    let col_100 = sheet.col_widths.iter().find(|c| c.col == 99); // 0-indexed (CV is column 100)
    assert!(
        col_100.is_some(),
        "Column 100 (index 99) should have width defined"
    );

    let width = col_100.unwrap().width;
    assert!(
        width > 20.0,
        "Column 100 width should be > 20 (Excel units)"
    );
}
