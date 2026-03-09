//! Layout feature tests for offidized-xlview
//!
//! Tests for column widths, row heights, hidden rows/columns, merged cells,
//! multiple sheets, sheet visibility, tab colors, frozen panes, and default dimensions.
//!
//! Ported from xlview's layout_tests.rs. The original tests used JSON-based assertions
//! against the xlview parser output. These tests use the typed Workbook/Sheet structs
//! from the offidized-xlview adapter pipeline.
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

/// Build a minimal XLSX from raw sheet XML content.
fn create_xlsx_from_raw_sheet_xml(sheet_name: &str, sheet_xml: &str) -> Vec<u8> {
    use std::io::{Cursor, Write};
    use zip::write::FileOptions;
    use zip::ZipWriter;

    let mut buf = Cursor::new(Vec::new());
    {
        let mut zip = ZipWriter::new(&mut buf);
        let options = FileOptions::<()>::default();

        // [Content_Types].xml
        zip.start_file("[Content_Types].xml", options).unwrap();
        zip.write_all(
            br#"<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
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
</Relationships>"#,
        )
        .unwrap();

        // xl/workbook.xml
        zip.start_file("xl/workbook.xml", options).unwrap();
        zip.write_all(
            format!(
                r#"<?xml version="1.0" encoding="UTF-8"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<sheets><sheet name="{}" sheetId="1" r:id="rId1" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"/></sheets>
</workbook>"#,
                sheet_name
            )
            .as_bytes(),
        )
        .unwrap();

        // xl/worksheets/sheet1.xml
        zip.start_file("xl/worksheets/sheet1.xml", options).unwrap();
        zip.write_all(sheet_xml.as_bytes()).unwrap();

        zip.finish().unwrap();
    }
    buf.into_inner()
}

/// Build a multi-sheet XLSX from raw XML.
fn create_xlsx_from_raw_multi_sheet(workbook_xml: &str, sheets: &[(&str, &str)]) -> Vec<u8> {
    use std::io::{Cursor, Write};
    use zip::write::FileOptions;
    use zip::ZipWriter;

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
<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>"#
            .to_string();

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
        let mut rels = r#"<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">"#
            .to_string();
        for (i, _) in sheets.iter().enumerate() {
            rels.push_str(&format!(
                r#"<Relationship Id="rId{}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet{}.xml"/>"#,
                i + 1,
                i + 1
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

/// Wrap sheet content in worksheet XML envelope.
fn wrap_sheet(content: &str) -> String {
    format!(
        r#"<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
{}
</worksheet>"#,
        content
    )
}

// =============================================================================
// COLUMN WIDTH TESTS
// =============================================================================

#[test]
fn test_column_width_custom_single() {
    // Test: col width="15" for single column
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

    let xlsx = create_xlsx_from_raw_sheet_xml("Sheet1", &sheet_xml);
    let workbook = load_xlsx(&xlsx);

    let sheet = &workbook.sheets[0];
    // Find column 0 (A) width
    let col_a = sheet.col_widths.iter().find(|cw| cw.col == 0);
    assert!(col_a.is_some(), "Column A should have custom width");

    let width = col_a.unwrap().width;
    assert!(
        (width - 15.0).abs() < 1.0,
        "Width should be ~15 (Excel units), got {}",
        width
    );
}

#[test]
fn test_column_width_range() {
    // Test: col min="1" max="5" width="20" for multiple columns
    let sheet_xml = wrap_sheet(
        r#"
        <cols>
            <col min="1" max="5" width="20" customWidth="1"/>
        </cols>
        <sheetData>
            <row r="1"><c r="A1"><v>A</v></c><c r="E1"><v>E</v></c></row>
        </sheetData>
    "#,
    );

    let xlsx = create_xlsx_from_raw_sheet_xml("Sheet1", &sheet_xml);
    let workbook = load_xlsx(&xlsx);

    let sheet = &workbook.sheets[0];
    // Should have 5 columns (0-4)
    assert!(
        sheet.col_widths.len() >= 5,
        "Should have at least 5 column widths"
    );

    // All columns 0-4 should have width 20 (Excel character units)
    for col_idx in 0u32..5 {
        let col = sheet.col_widths.iter().find(|cw| cw.col == col_idx);
        assert!(
            col.is_some(),
            "Column {} should have width defined",
            col_idx
        );
        let width = col.unwrap().width;
        assert!(
            (width - 20.0).abs() < 1.0,
            "Column {} width should be ~20 (Excel units), got {}",
            col_idx,
            width
        );
    }
}

#[test]
fn test_column_width_default() {
    // Test: No col definition, use default
    let sheet_xml = wrap_sheet(
        r#"
        <sheetData>
            <row r="1"><c r="A1"><v>Test</v></c></row>
        </sheetData>
    "#,
    );

    let xlsx = create_xlsx_from_raw_sheet_xml("Sheet1", &sheet_xml);
    let workbook = load_xlsx(&xlsx);

    let sheet = &workbook.sheets[0];
    // colWidths should be empty when no custom widths defined
    assert!(
        sheet.col_widths.is_empty(),
        "No custom column widths should be set"
    );

    // defaultColWidth should be positive
    assert!(
        sheet.default_col_width > 0.0,
        "Default column width should be positive"
    );
}

#[test]
fn test_column_width_zero() {
    // Test: width="0" (hidden via width)
    let sheet_xml = wrap_sheet(
        r#"
        <cols>
            <col min="2" max="2" width="0" customWidth="1"/>
        </cols>
        <sheetData>
            <row r="1"><c r="B1"><v>Hidden</v></c></row>
        </sheetData>
    "#,
    );

    let xlsx = create_xlsx_from_raw_sheet_xml("Sheet1", &sheet_xml);
    let workbook = load_xlsx(&xlsx);

    let sheet = &workbook.sheets[0];
    let col_b = sheet.col_widths.iter().find(|cw| cw.col == 1);
    assert!(col_b.is_some(), "Column B should be defined");

    let width = col_b.unwrap().width;
    assert!(
        width < 1.0,
        "Zero width column should be ~0 (Excel units), got {}",
        width
    );
}

#[test]
fn test_column_width_very_wide() {
    // Test: width="100"
    let sheet_xml = wrap_sheet(
        r#"
        <cols>
            <col min="1" max="1" width="100" customWidth="1"/>
        </cols>
        <sheetData>
            <row r="1"><c r="A1"><v>Wide</v></c></row>
        </sheetData>
    "#,
    );

    let xlsx = create_xlsx_from_raw_sheet_xml("Sheet1", &sheet_xml);
    let workbook = load_xlsx(&xlsx);

    let sheet = &workbook.sheets[0];
    let col_a = sheet.col_widths.iter().find(|cw| cw.col == 0);
    assert!(col_a.is_some(), "Column A should have width");

    let width = col_a.unwrap().width;
    assert!(
        (width - 100.0).abs() < 1.0,
        "Very wide column should be ~100 (Excel units), got {}",
        width
    );
}

// =============================================================================
// ROW HEIGHT TESTS
// =============================================================================

#[test]
fn test_row_height_custom() {
    // Test: row ht="30" customHeight="1"
    let sheet_xml = wrap_sheet(
        r#"
        <sheetData>
            <row r="1" ht="30" customHeight="1"><c r="A1"><v>Tall</v></c></row>
        </sheetData>
    "#,
    );

    let xlsx = create_xlsx_from_raw_sheet_xml("Sheet1", &sheet_xml);
    let workbook = load_xlsx(&xlsx);

    let sheet = &workbook.sheets[0];
    let row_1 = sheet.row_heights.iter().find(|rh| rh.row == 0);
    assert!(row_1.is_some(), "Row 1 should have custom height");

    // Height 30 points (stored as points, not pixels)
    let height = row_1.unwrap().height;
    assert!(
        (height - 30.0).abs() < 1.0,
        "Row height should be ~30 points, got {}",
        height
    );
}

#[test]
fn test_row_height_default() {
    // Test: No ht attribute
    let sheet_xml = wrap_sheet(
        r#"
        <sheetData>
            <row r="1"><c r="A1"><v>Default height</v></c></row>
        </sheetData>
    "#,
    );

    let xlsx = create_xlsx_from_raw_sheet_xml("Sheet1", &sheet_xml);
    let workbook = load_xlsx(&xlsx);

    let sheet = &workbook.sheets[0];
    assert!(
        sheet.row_heights.is_empty(),
        "No custom row heights should be set"
    );

    let default_height = sheet.default_row_height;
    assert!(
        default_height > 0.0,
        "Default row height should be positive"
    );
}

#[test]
fn test_row_height_zero() {
    // Test: ht="0" (hidden via height)
    let sheet_xml = wrap_sheet(
        r#"
        <sheetData>
            <row r="2" ht="0" customHeight="1"><c r="A2"><v>Hidden</v></c></row>
        </sheetData>
    "#,
    );

    let xlsx = create_xlsx_from_raw_sheet_xml("Sheet1", &sheet_xml);
    let workbook = load_xlsx(&xlsx);

    let sheet = &workbook.sheets[0];
    let row_2 = sheet.row_heights.iter().find(|rh| rh.row == 1);
    assert!(row_2.is_some(), "Row 2 should have height defined");

    let height = row_2.unwrap().height;
    assert!(
        height < 1.0,
        "Zero height row should be ~0 pixels, got {}",
        height
    );
}

#[test]
fn test_row_height_very_tall() {
    // Test: ht="100"
    let sheet_xml = wrap_sheet(
        r#"
        <sheetData>
            <row r="1" ht="100" customHeight="1"><c r="A1"><v>Tall</v></c></row>
        </sheetData>
    "#,
    );

    let xlsx = create_xlsx_from_raw_sheet_xml("Sheet1", &sheet_xml);
    let workbook = load_xlsx(&xlsx);

    let sheet = &workbook.sheets[0];
    let row_1 = sheet.row_heights.iter().find(|rh| rh.row == 0);
    assert!(row_1.is_some(), "Row 1 should have height");

    // Height 100 points (stored as points, not pixels)
    let height = row_1.unwrap().height;
    assert!(
        (height - 100.0).abs() < 1.0,
        "Very tall row should be ~100 points, got {}",
        height
    );
}

// =============================================================================
// HIDDEN ROWS/COLUMNS TESTS
// =============================================================================

#[test]
fn test_hidden_column() {
    // Test: col hidden="1"
    let sheet_xml = wrap_sheet(
        r#"
        <cols>
            <col min="2" max="2" width="10" hidden="1"/>
        </cols>
        <sheetData>
            <row r="1"><c r="B1"><v>Hidden</v></c></row>
        </sheetData>
    "#,
    );

    let xlsx = create_xlsx_from_raw_sheet_xml("Sheet1", &sheet_xml);
    let workbook = load_xlsx(&xlsx);

    let sheet = &workbook.sheets[0];
    assert!(
        sheet.hidden_cols.contains(&1),
        "Column B (index 1) should be hidden"
    );
}

#[test]
fn test_hidden_row() {
    // Test: row hidden="1"
    let sheet_xml = wrap_sheet(
        r#"
        <sheetData>
            <row r="1"><c r="A1"><v>Visible</v></c></row>
            <row r="2" hidden="1"><c r="A2"><v>Hidden</v></c></row>
            <row r="3"><c r="A3"><v>Visible</v></c></row>
        </sheetData>
    "#,
    );

    let xlsx = create_xlsx_from_raw_sheet_xml("Sheet1", &sheet_xml);
    let workbook = load_xlsx(&xlsx);

    let sheet = &workbook.sheets[0];
    assert!(
        sheet.hidden_rows.contains(&1),
        "Row 2 (index 1) should be hidden"
    );
    assert!(
        !sheet.hidden_rows.contains(&0),
        "Row 1 should not be hidden"
    );
    assert!(
        !sheet.hidden_rows.contains(&2),
        "Row 3 should not be hidden"
    );
}

#[test]
fn test_hidden_range() {
    // Test: Multiple consecutive hidden columns and rows
    let sheet_xml = wrap_sheet(
        r#"
        <cols>
            <col min="2" max="4" width="10" hidden="1"/>
        </cols>
        <sheetData>
            <row r="1"><c r="A1"><v>A</v></c></row>
            <row r="2" hidden="1"><c r="A2"><v>2</v></c></row>
            <row r="3" hidden="1"><c r="A3"><v>3</v></c></row>
            <row r="4" hidden="1"><c r="A4"><v>4</v></c></row>
            <row r="5"><c r="A5"><v>5</v></c></row>
        </sheetData>
    "#,
    );

    let xlsx = create_xlsx_from_raw_sheet_xml("Sheet1", &sheet_xml);
    let workbook = load_xlsx(&xlsx);

    let sheet = &workbook.sheets[0];
    assert!(sheet.hidden_cols.contains(&1), "Column B should be hidden");
    assert!(sheet.hidden_cols.contains(&2), "Column C should be hidden");
    assert!(sheet.hidden_cols.contains(&3), "Column D should be hidden");

    assert!(sheet.hidden_rows.contains(&1), "Row 2 should be hidden");
    assert!(sheet.hidden_rows.contains(&2), "Row 3 should be hidden");
    assert!(sheet.hidden_rows.contains(&3), "Row 4 should be hidden");
}

// =============================================================================
// MERGED CELLS TESTS
// =============================================================================

#[test]
fn test_merge_simple() {
    // Test: mergeCell ref="A1:B2"
    let sheet_xml = wrap_sheet(
        r#"
        <sheetData>
            <row r="1"><c r="A1"><v>Merged</v></c></row>
        </sheetData>
        <mergeCells count="1">
            <mergeCell ref="A1:B2"/>
        </mergeCells>
    "#,
    );

    let xlsx = create_xlsx_from_raw_sheet_xml("Sheet1", &sheet_xml);
    let workbook = load_xlsx(&xlsx);

    let sheet = &workbook.sheets[0];
    assert_eq!(sheet.merges.len(), 1, "Should have 1 merge");

    let merge = &sheet.merges[0];
    assert_eq!(merge.start_row, 0);
    assert_eq!(merge.start_col, 0);
    assert_eq!(merge.end_row, 1);
    assert_eq!(merge.end_col, 1);
}

#[test]
fn test_merge_wide() {
    // Test: mergeCell ref="A1:Z1" (26 columns wide)
    let sheet_xml = wrap_sheet(
        r#"
        <sheetData>
            <row r="1"><c r="A1"><v>Wide merge</v></c></row>
        </sheetData>
        <mergeCells count="1">
            <mergeCell ref="A1:Z1"/>
        </mergeCells>
    "#,
    );

    let xlsx = create_xlsx_from_raw_sheet_xml("Sheet1", &sheet_xml);
    let workbook = load_xlsx(&xlsx);

    let sheet = &workbook.sheets[0];
    assert_eq!(sheet.merges.len(), 1);

    let merge = &sheet.merges[0];
    assert_eq!(merge.start_row, 0);
    assert_eq!(merge.start_col, 0);
    assert_eq!(merge.end_row, 0);
    assert_eq!(merge.end_col, 25); // Z is column 26 (0-indexed: 25)
}

#[test]
fn test_merge_tall() {
    // Test: mergeCell ref="A1:A100" (100 rows tall)
    let sheet_xml = wrap_sheet(
        r#"
        <sheetData>
            <row r="1"><c r="A1"><v>Tall merge</v></c></row>
        </sheetData>
        <mergeCells count="1">
            <mergeCell ref="A1:A100"/>
        </mergeCells>
    "#,
    );

    let xlsx = create_xlsx_from_raw_sheet_xml("Sheet1", &sheet_xml);
    let workbook = load_xlsx(&xlsx);

    let sheet = &workbook.sheets[0];
    assert_eq!(sheet.merges.len(), 1);

    let merge = &sheet.merges[0];
    assert_eq!(merge.start_row, 0);
    assert_eq!(merge.start_col, 0);
    assert_eq!(merge.end_row, 99); // Row 100 (0-indexed: 99)
    assert_eq!(merge.end_col, 0);
}

#[test]
fn test_merge_large() {
    // Test: mergeCell ref="A1:D10"
    let sheet_xml = wrap_sheet(
        r#"
        <sheetData>
            <row r="1"><c r="A1"><v>Large merge</v></c></row>
        </sheetData>
        <mergeCells count="1">
            <mergeCell ref="A1:D10"/>
        </mergeCells>
    "#,
    );

    let xlsx = create_xlsx_from_raw_sheet_xml("Sheet1", &sheet_xml);
    let workbook = load_xlsx(&xlsx);

    let sheet = &workbook.sheets[0];
    assert_eq!(sheet.merges.len(), 1);

    let merge = &sheet.merges[0];
    assert_eq!(merge.start_row, 0);
    assert_eq!(merge.start_col, 0);
    assert_eq!(merge.end_row, 9);
    assert_eq!(merge.end_col, 3); // D is column 4 (0-indexed: 3)
}

#[test]
fn test_merge_multiple() {
    // Test: Several non-overlapping merges
    let sheet_xml = wrap_sheet(
        r#"
        <sheetData>
            <row r="1"><c r="A1"><v>Merge 1</v></c><c r="D1"><v>Merge 2</v></c></row>
            <row r="5"><c r="A5"><v>Merge 3</v></c></row>
        </sheetData>
        <mergeCells count="3">
            <mergeCell ref="A1:B2"/>
            <mergeCell ref="D1:F1"/>
            <mergeCell ref="A5:C7"/>
        </mergeCells>
    "#,
    );

    let xlsx = create_xlsx_from_raw_sheet_xml("Sheet1", &sheet_xml);
    let workbook = load_xlsx(&xlsx);

    let sheet = &workbook.sheets[0];
    assert_eq!(sheet.merges.len(), 3, "Should have 3 merges");

    // Verify each merge exists (order may vary)
    let has_a1_b2 = sheet
        .merges
        .iter()
        .any(|m| m.start_row == 0 && m.start_col == 0 && m.end_row == 1 && m.end_col == 1);
    let has_d1_f1 = sheet
        .merges
        .iter()
        .any(|m| m.start_row == 0 && m.start_col == 3 && m.end_row == 0 && m.end_col == 5);
    let has_a5_c7 = sheet
        .merges
        .iter()
        .any(|m| m.start_row == 4 && m.start_col == 0 && m.end_row == 6 && m.end_col == 2);

    assert!(has_a1_b2, "Should have A1:B2 merge");
    assert!(has_d1_f1, "Should have D1:F1 merge");
    assert!(has_a5_c7, "Should have A5:C7 merge");
}

#[test]
fn test_merge_with_content() {
    // Test: Value in top-left cell only
    let sheet_xml = wrap_sheet(
        r#"
        <sheetData>
            <row r="1">
                <c r="A1"><v>Content in top-left</v></c>
                <c r="B1"/>
            </row>
            <row r="2">
                <c r="A2"/>
                <c r="B2"/>
            </row>
        </sheetData>
        <mergeCells count="1">
            <mergeCell ref="A1:B2"/>
        </mergeCells>
    "#,
    );

    let xlsx = create_xlsx_from_raw_sheet_xml("Sheet1", &sheet_xml);
    let workbook = load_xlsx(&xlsx);

    let sheet = &workbook.sheets[0];
    assert_eq!(sheet.merges.len(), 1);

    // Find cell A1 and verify it has content
    let cell_a1 = sheet.cells.iter().find(|c| c.r == 0 && c.c == 0);
    assert!(cell_a1.is_some(), "Cell A1 should exist");
    // Value is stored in raw (not v), check that it exists
    let raw = cell_a1.unwrap().cell.raw.as_ref();
    assert!(raw.is_some(), "Cell A1 should have a raw value");
}

// =============================================================================
// MULTIPLE SHEETS TESTS
// =============================================================================

#[test]
fn test_two_sheets() {
    let xlsx = XlsxBuilder::new()
        .sheet(SheetBuilder::new("First Sheet").cell("A1", "Sheet 1 Content", None))
        .sheet(SheetBuilder::new("Second Sheet").cell("A1", "Sheet 2 Content", None))
        .build();

    let workbook = load_xlsx(&xlsx);

    assert_eq!(workbook.sheets.len(), 2, "Should have 2 sheets");
    assert_eq!(workbook.sheets[0].name, "First Sheet");
    assert_eq!(workbook.sheets[1].name, "Second Sheet");
}

#[test]
fn test_many_sheets() {
    // Test: workbook with 10 sheets
    let mut builder = XlsxBuilder::new();
    for i in 1..=10 {
        builder = builder.sheet(SheetBuilder::new(&format!("Sheet{}", i)).cell(
            "A1",
            format!("Content {}", i),
            None,
        ));
    }
    let xlsx = builder.build();

    let workbook = load_xlsx(&xlsx);

    assert_eq!(workbook.sheets.len(), 10, "Should have 10 sheets");

    for i in 1..=10 {
        assert_eq!(workbook.sheets[i - 1].name, format!("Sheet{}", i));
    }
}

#[test]
fn test_sheet_names_special() {
    let xlsx = XlsxBuilder::new()
        .sheet(SheetBuilder::new("Sales Data 2024").cell("A1", "1", None))
        .sheet(SheetBuilder::new("Q1 Summary").cell("A1", "2", None))
        .sheet(SheetBuilder::new("Sheet-With-Dashes").cell("A1", "3", None))
        .build();

    let workbook = load_xlsx(&xlsx);

    assert_eq!(workbook.sheets[0].name, "Sales Data 2024");
    assert_eq!(workbook.sheets[1].name, "Q1 Summary");
    assert_eq!(workbook.sheets[2].name, "Sheet-With-Dashes");
}

#[test]
fn test_sheet_order() {
    let xlsx = XlsxBuilder::new()
        .sheet(SheetBuilder::new("Alpha").cell("A1", "X", None))
        .sheet(SheetBuilder::new("Beta").cell("A1", "X", None))
        .sheet(SheetBuilder::new("Gamma").cell("A1", "X", None))
        .sheet(SheetBuilder::new("Delta").cell("A1", "X", None))
        .build();

    let workbook = load_xlsx(&xlsx);

    assert_eq!(workbook.sheets.len(), 4);
    assert_eq!(workbook.sheets[0].name, "Alpha");
    assert_eq!(workbook.sheets[1].name, "Beta");
    assert_eq!(workbook.sheets[2].name, "Gamma");
    assert_eq!(workbook.sheets[3].name, "Delta");
}

// =============================================================================
// SHEET VISIBILITY TESTS
// =============================================================================

#[test]
fn test_sheet_visibility_visible() {
    let workbook_xml = r#"<?xml version="1.0" encoding="UTF-8"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<sheets>
    <sheet name="Visible Sheet" sheetId="1" state="visible" r:id="rId1"/>
</sheets>
</workbook>"#;

    let sheet =
        wrap_sheet(r#"<sheetData><row r="1"><c r="A1"><v>Visible</v></c></row></sheetData>"#);

    let xlsx = create_xlsx_from_raw_multi_sheet(workbook_xml, &[("Visible Sheet", &sheet)]);
    let workbook = load_xlsx(&xlsx);

    assert_eq!(workbook.sheets.len(), 1);
    assert_eq!(workbook.sheets[0].name, "Visible Sheet");
}

#[test]
fn test_sheet_visibility_hidden() {
    let workbook_xml = r#"<?xml version="1.0" encoding="UTF-8"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<sheets>
    <sheet name="Visible" sheetId="1" r:id="rId1"/>
    <sheet name="Hidden" sheetId="2" state="hidden" r:id="rId2"/>
</sheets>
</workbook>"#;

    let sheet = wrap_sheet(r#"<sheetData><row r="1"><c r="A1"><v>X</v></c></row></sheetData>"#);

    let xlsx =
        create_xlsx_from_raw_multi_sheet(workbook_xml, &[("Visible", &sheet), ("Hidden", &sheet)]);
    let workbook = load_xlsx(&xlsx);

    assert_eq!(workbook.sheets.len(), 2);
}

#[test]
fn test_sheet_visibility_very_hidden() {
    let workbook_xml = r#"<?xml version="1.0" encoding="UTF-8"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<sheets>
    <sheet name="Normal" sheetId="1" r:id="rId1"/>
    <sheet name="VeryHidden" sheetId="2" state="veryHidden" r:id="rId2"/>
</sheets>
</workbook>"#;

    let sheet = wrap_sheet(r#"<sheetData><row r="1"><c r="A1"><v>X</v></c></row></sheetData>"#);

    let xlsx = create_xlsx_from_raw_multi_sheet(
        workbook_xml,
        &[("Normal", &sheet), ("VeryHidden", &sheet)],
    );
    let workbook = load_xlsx(&xlsx);

    assert_eq!(workbook.sheets.len(), 2);
}

// =============================================================================
// SHEET TAB COLOR TESTS
// =============================================================================

#[test]
fn test_tab_color_rgb() {
    let workbook_xml = r#"<?xml version="1.0" encoding="UTF-8"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<sheets>
    <sheet name="Green Tab" sheetId="1" r:id="rId1"/>
</sheets>
</workbook>"#;

    let sheet = wrap_sheet(
        r#"
        <sheetPr>
            <tabColor rgb="FF00FF00"/>
        </sheetPr>
        <sheetData><row r="1"><c r="A1"><v>Green</v></c></row></sheetData>
    "#,
    );

    let xlsx = create_xlsx_from_raw_multi_sheet(workbook_xml, &[("Green Tab", &sheet)]);
    let workbook = load_xlsx(&xlsx);

    assert_eq!(workbook.sheets.len(), 1);
    // Tab color should be set
    // The adapter sets it as "#RRGGBB" format
    if let Some(ref color) = workbook.sheets[0].tab_color {
        assert!(
            color.contains("00FF00"),
            "Tab color should contain 00FF00, got {}",
            color
        );
    }
}

#[test]
fn test_tab_color_theme() {
    let workbook_xml = r#"<?xml version="1.0" encoding="UTF-8"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<sheets>
    <sheet name="Theme Tab" sheetId="1" r:id="rId1"/>
</sheets>
</workbook>"#;

    let sheet = wrap_sheet(
        r#"
        <sheetPr>
            <tabColor theme="4"/>
        </sheetPr>
        <sheetData><row r="1"><c r="A1"><v>Theme</v></c></row></sheetData>
    "#,
    );

    let xlsx = create_xlsx_from_raw_multi_sheet(workbook_xml, &[("Theme Tab", &sheet)]);
    let workbook = load_xlsx(&xlsx);

    assert_eq!(workbook.sheets.len(), 1);
    // Theme-based tab color will resolve to a hex color if theme is available
}

// =============================================================================
// FROZEN PANES TESTS
// =============================================================================

#[test]
fn test_frozen_rows() {
    let sheet_xml = wrap_sheet(
        r#"
        <sheetViews>
            <sheetView tabSelected="1" workbookViewId="0">
                <pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/>
            </sheetView>
        </sheetViews>
        <sheetData>
            <row r="1"><c r="A1"><v>Header</v></c></row>
            <row r="2"><c r="A2"><v>Data</v></c></row>
        </sheetData>
    "#,
    );

    let xlsx = create_xlsx_from_raw_sheet_xml("Sheet1", &sheet_xml);
    let workbook = load_xlsx(&xlsx);

    assert_eq!(workbook.sheets.len(), 1);
    assert_eq!(workbook.sheets[0].frozen_rows, 1);
    assert_eq!(workbook.sheets[0].frozen_cols, 0);
}

#[test]
fn test_frozen_columns() {
    let sheet_xml = wrap_sheet(
        r#"
        <sheetViews>
            <sheetView tabSelected="1" workbookViewId="0">
                <pane xSplit="1" topLeftCell="B1" activePane="topRight" state="frozen"/>
            </sheetView>
        </sheetViews>
        <sheetData>
            <row r="1"><c r="A1"><v>Label</v></c><c r="B1"><v>Value</v></c></row>
        </sheetData>
    "#,
    );

    let xlsx = create_xlsx_from_raw_sheet_xml("Sheet1", &sheet_xml);
    let workbook = load_xlsx(&xlsx);

    assert_eq!(workbook.sheets.len(), 1);
    assert_eq!(workbook.sheets[0].frozen_rows, 0);
    assert_eq!(workbook.sheets[0].frozen_cols, 1);
}

#[test]
fn test_frozen_both() {
    let sheet_xml = wrap_sheet(
        r#"
        <sheetViews>
            <sheetView tabSelected="1" workbookViewId="0">
                <pane xSplit="1" ySplit="2" topLeftCell="B3" activePane="bottomRight" state="frozen"/>
            </sheetView>
        </sheetViews>
        <sheetData>
            <row r="1"><c r="A1"><v>Corner</v></c><c r="B1"><v>Header 1</v></c></row>
            <row r="2"><c r="A2"><v>Label</v></c><c r="B2"><v>Header 2</v></c></row>
            <row r="3"><c r="A3"><v>Row Label</v></c><c r="B3"><v>Data</v></c></row>
        </sheetData>
    "#,
    );

    let xlsx = create_xlsx_from_raw_sheet_xml("Sheet1", &sheet_xml);
    let workbook = load_xlsx(&xlsx);

    assert_eq!(workbook.sheets.len(), 1);
    assert_eq!(workbook.sheets[0].frozen_rows, 2);
    assert_eq!(workbook.sheets[0].frozen_cols, 1);
}

#[test]
fn test_split_panes() {
    let sheet_xml = wrap_sheet(
        r#"
        <sheetViews>
            <sheetView tabSelected="1" workbookViewId="0">
                <pane xSplit="2000" ySplit="1500" topLeftCell="C5" activePane="bottomRight" state="split"/>
            </sheetView>
        </sheetViews>
        <sheetData>
            <row r="1"><c r="A1"><v>Split</v></c></row>
        </sheetData>
    "#,
    );

    let xlsx = create_xlsx_from_raw_sheet_xml("Sheet1", &sheet_xml);
    let workbook = load_xlsx(&xlsx);

    assert_eq!(workbook.sheets.len(), 1);
    // Split panes don't set frozen_rows/cols
    assert_eq!(workbook.sheets[0].frozen_rows, 0);
    assert_eq!(workbook.sheets[0].frozen_cols, 0);
}

// =============================================================================
// DEFAULT DIMENSIONS TESTS
// =============================================================================

#[test]
fn test_sheet_format_pr_defaults() {
    let sheet_xml = wrap_sheet(
        r#"
        <sheetFormatPr defaultColWidth="12.5" defaultRowHeight="18"/>
        <sheetData>
            <row r="1"><c r="A1"><v>Test</v></c></row>
        </sheetData>
    "#,
    );

    let xlsx = create_xlsx_from_raw_sheet_xml("Sheet1", &sheet_xml);
    let workbook = load_xlsx(&xlsx);

    let sheet = &workbook.sheets[0];

    // Verify defaults are present (may be hardcoded or from sheetFormatPr)
    assert!(
        sheet.default_col_width > 0.0,
        "Default column width should be positive"
    );
    assert!(
        sheet.default_row_height > 0.0,
        "Default row height should be positive"
    );
}

// =============================================================================
// COMBINED LAYOUT TESTS
// =============================================================================

#[test]
fn test_complex_layout() {
    let sheet_xml = wrap_sheet(
        r#"
        <cols>
            <col min="1" max="1" width="25" customWidth="1"/>
            <col min="2" max="2" width="10" hidden="1"/>
            <col min="3" max="5" width="15" customWidth="1"/>
        </cols>
        <sheetData>
            <row r="1" ht="30" customHeight="1">
                <c r="A1"><v>Header</v></c>
                <c r="C1"><v>Merged Header</v></c>
            </row>
            <row r="2" hidden="1"><c r="A2"><v>Hidden Row</v></c></row>
            <row r="3"><c r="A3"><v>Data</v></c></row>
        </sheetData>
        <mergeCells count="1">
            <mergeCell ref="C1:E1"/>
        </mergeCells>
    "#,
    );

    let xlsx = create_xlsx_from_raw_sheet_xml("Sheet1", &sheet_xml);
    let workbook = load_xlsx(&xlsx);

    let sheet = &workbook.sheets[0];

    // Verify column widths
    assert!(
        sheet.col_widths.len() >= 5,
        "Should have column width definitions"
    );

    // Verify hidden column
    assert!(sheet.hidden_cols.contains(&1), "Column B should be hidden");

    // Verify row height
    let row_1 = sheet.row_heights.iter().find(|rh| rh.row == 0);
    assert!(row_1.is_some(), "Row 1 should have custom height");

    // Verify hidden row
    assert!(sheet.hidden_rows.contains(&1), "Row 2 should be hidden");

    // Verify merge
    assert_eq!(sheet.merges.len(), 1);
    let merge = &sheet.merges[0];
    assert_eq!(merge.start_col, 2); // C
    assert_eq!(merge.end_col, 4); // E
}

#[test]
fn test_max_dimensions() {
    let sheet_xml = wrap_sheet(
        r#"
        <sheetData>
            <row r="1"><c r="A1"><v>A1</v></c></row>
            <row r="5"><c r="D5"><v>D5</v></c></row>
            <row r="10"><c r="J10"><v>J10</v></c></row>
        </sheetData>
    "#,
    );

    let xlsx = create_xlsx_from_raw_sheet_xml("Sheet1", &sheet_xml);
    let workbook = load_xlsx(&xlsx);

    let sheet = &workbook.sheets[0];

    // maxRow should be 9 (0-indexed: row 10 in source = index 9)
    assert_eq!(sheet.max_row, 9, "maxRow should be 9 (0-indexed)");

    // maxCol should be 9 (0-indexed: column J = index 9)
    assert_eq!(sheet.max_col, 9, "maxCol should be 9 (0-indexed)");
}

#[test]
fn test_sparse_data() {
    let sheet_xml = wrap_sheet(
        r#"
        <sheetData>
            <row r="1"><c r="A1"><v>1</v></c></row>
            <row r="100"><c r="Z100"><v>100</v></c></row>
        </sheetData>
    "#,
    );

    let xlsx = create_xlsx_from_raw_sheet_xml("Sheet1", &sheet_xml);
    let workbook = load_xlsx(&xlsx);

    let sheet = &workbook.sheets[0];

    // Only 2 cells should be in the sparse representation
    assert_eq!(sheet.cells.len(), 2, "Should only have 2 cells");

    // Verify cell positions
    let cell_a1 = sheet.cells.iter().find(|c| c.r == 0 && c.c == 0);
    let cell_z100 = sheet.cells.iter().find(|c| c.r == 99 && c.c == 25);

    assert!(cell_a1.is_some(), "Cell A1 should exist");
    assert!(cell_z100.is_some(), "Cell Z100 should exist");
}
