//! Comprehensive tests for merge cell functionality in offidized-xlview
//!
//! These tests verify that merge cell definitions from xl/worksheets/sheetN.xml
//! are correctly parsed into the viewer workbook structure.

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

use common::{
    assert_cell_value, assert_merge_exists, assert_sheet_count, get_cell, get_cell_style, load_xlsx,
};
use fixtures::{SheetBuilder, StyleBuilder, XlsxBuilder};

// ============================================================================
// BASIC MERGE TESTS
// ============================================================================

/// Test 1: Simple 2x2 merge
#[test]
fn test_simple_2x2_merge() {
    let xlsx = XlsxBuilder::new()
        .sheet(
            SheetBuilder::new("Sheet1")
                .cell("A1", "Merged Header", None)
                .merge("A1:B2"),
        )
        .build();

    let workbook = load_xlsx(&xlsx);

    assert_sheet_count(&workbook, 1);
    assert_cell_value(&workbook, 0, 0, 0, "Merged Header");

    // Verify merge exists: A1:B2 = (row 0, col 0) to (row 1, col 1)
    assert_merge_exists(&workbook, 0, 0, 0, 1, 1);
}

/// Test 2: Wide merge (1 row, 10 columns)
#[test]
fn test_wide_merge_1_row_10_columns() {
    let xlsx = XlsxBuilder::new()
        .sheet(
            SheetBuilder::new("Sheet1")
                .cell("A1", "Wide Title Spanning Many Columns", None)
                .merge("A1:J1"),
        )
        .build();

    let workbook = load_xlsx(&xlsx);

    assert_cell_value(&workbook, 0, 0, 0, "Wide Title Spanning Many Columns");

    // Verify merge exists: A1:J1 = (row 0, col 0) to (row 0, col 9)
    assert_merge_exists(&workbook, 0, 0, 0, 0, 9);
}

/// Test 3: Tall merge (10 rows, 1 column)
#[test]
fn test_tall_merge_10_rows_1_column() {
    let xlsx = XlsxBuilder::new()
        .sheet(
            SheetBuilder::new("Sheet1")
                .cell("A1", "Row Label", None)
                .merge("A1:A10"),
        )
        .build();

    let workbook = load_xlsx(&xlsx);

    assert_cell_value(&workbook, 0, 0, 0, "Row Label");

    // Verify merge exists: A1:A10 = (row 0, col 0) to (row 9, col 0)
    assert_merge_exists(&workbook, 0, 0, 0, 9, 0);
}

/// Test 4: Large merge (5x5)
#[test]
fn test_large_5x5_merge() {
    let xlsx = XlsxBuilder::new()
        .sheet(
            SheetBuilder::new("Sheet1")
                .cell("A1", "Large Block", None)
                .merge("A1:E5"),
        )
        .build();

    let workbook = load_xlsx(&xlsx);

    assert_cell_value(&workbook, 0, 0, 0, "Large Block");

    // Verify merge exists: A1:E5 = (row 0, col 0) to (row 4, col 4)
    assert_merge_exists(&workbook, 0, 0, 0, 4, 4);
}

// ============================================================================
// MULTIPLE MERGES TESTS
// ============================================================================

/// Test 5: Multiple merges in one sheet
#[test]
fn test_multiple_merges_in_one_sheet() {
    let xlsx = XlsxBuilder::new()
        .sheet(
            SheetBuilder::new("Sheet1")
                .cell("A1", "Report Title", None)
                .merge("A1:D1")
                .cell("A3", "Section 1", None)
                .merge("A3:B3")
                .cell("C3", "Section 2", None)
                .merge("C3:D3")
                .cell("A5", "Group A", None)
                .merge("A5:A7")
                .cell("A8", "Group B", None)
                .merge("A8:A10"),
        )
        .build();

    let workbook = load_xlsx(&xlsx);

    // Verify all merges exist
    assert_merge_exists(&workbook, 0, 0, 0, 0, 3);
    assert_merge_exists(&workbook, 0, 2, 0, 2, 1);
    assert_merge_exists(&workbook, 0, 2, 2, 2, 3);
    assert_merge_exists(&workbook, 0, 4, 0, 6, 0);
    assert_merge_exists(&workbook, 0, 7, 0, 9, 0);

    // Verify the merge count
    let merges = &workbook.sheets[0].merges;
    assert_eq!(merges.len(), 5, "Should have 5 merges");
}

/// Test 6: Adjacent merges
#[test]
fn test_adjacent_merges() {
    let xlsx = XlsxBuilder::new()
        .sheet(
            SheetBuilder::new("Sheet1")
                .cell("A1", "Group 1", None)
                .merge("A1:B1")
                .cell("C1", "Group 2", None)
                .merge("C1:D1")
                .cell("E1", "Group 3", None)
                .merge("E1:F1")
                .cell("A2", "Sub A", None)
                .cell("B2", "Sub B", None)
                .cell("C2", "Sub C", None)
                .cell("D2", "Sub D", None)
                .cell("E2", "Sub E", None)
                .cell("F2", "Sub F", None),
        )
        .build();

    let workbook = load_xlsx(&xlsx);

    // Verify adjacent merges
    assert_merge_exists(&workbook, 0, 0, 0, 0, 1);
    assert_merge_exists(&workbook, 0, 0, 2, 0, 3);
    assert_merge_exists(&workbook, 0, 0, 4, 0, 5);

    // Verify values in adjacent cells
    assert_cell_value(&workbook, 0, 0, 0, "Group 1");
    assert_cell_value(&workbook, 0, 0, 2, "Group 2");
    assert_cell_value(&workbook, 0, 0, 4, "Group 3");
}

/// Test 6b: Vertically adjacent merges
#[test]
fn test_vertically_adjacent_merges() {
    let xlsx = XlsxBuilder::new()
        .sheet(
            SheetBuilder::new("Sheet1")
                .cell("A1", "Header 1", None)
                .merge("A1:A2")
                .cell("A3", "Header 2", None)
                .merge("A3:A4")
                .cell("A5", "Header 3", None)
                .merge("A5:A6"),
        )
        .build();

    let workbook = load_xlsx(&xlsx);

    assert_merge_exists(&workbook, 0, 0, 0, 1, 0);
    assert_merge_exists(&workbook, 0, 2, 0, 3, 0);
    assert_merge_exists(&workbook, 0, 4, 0, 5, 0);
}

// ============================================================================
// MERGE WITH CONTENT TESTS
// ============================================================================

/// Test 7: Merge with content in top-left cell
#[test]
fn test_merge_with_content_in_top_left() {
    let xlsx = XlsxBuilder::new()
        .sheet(
            SheetBuilder::new("Sheet1")
                .cell("A1", "Top-Left Content", None)
                .merge("A1:C3"),
        )
        .build();

    let workbook = load_xlsx(&xlsx);

    assert_cell_value(&workbook, 0, 0, 0, "Top-Left Content");
    assert_merge_exists(&workbook, 0, 0, 0, 2, 2);
}

/// Test 7b: Merge with numeric content
#[test]
fn test_merge_with_numeric_content() {
    let xlsx = XlsxBuilder::new()
        .sheet(
            SheetBuilder::new("Sheet1")
                .cell("A1", 12345.67, None)
                .merge("A1:B2"),
        )
        .build();

    let workbook = load_xlsx(&xlsx);

    // Verify numeric content -- the adapter may store it in raw, v, or cached_display
    let cd = get_cell(&workbook, 0, 0, 0).expect("Cell should exist");
    let has_value =
        cd.cell.cached_display.is_some() || cd.cell.v.is_some() || cd.cell.raw.is_some();
    assert!(
        has_value,
        "Cell should have a value (in raw, v, or cached_display)"
    );

    assert_merge_exists(&workbook, 0, 0, 0, 1, 1);
}

/// Test 7c: Merge with special characters in content
#[test]
fn test_merge_with_special_characters() {
    let xlsx = XlsxBuilder::new()
        .sheet(
            SheetBuilder::new("Sheet1")
                .cell("A1", "Special: <>&\"'", None)
                .merge("A1:B1"),
        )
        .build();

    let workbook = load_xlsx(&xlsx);

    assert_cell_value(&workbook, 0, 0, 0, "Special: <>&\"'");
    assert_merge_exists(&workbook, 0, 0, 0, 0, 1);
}

// ============================================================================
// MERGE WITH STYLE TESTS
// ============================================================================

/// Test 8: Merge with style applied
#[test]
fn test_merge_with_style_applied() {
    let style = StyleBuilder::new()
        .bold()
        .italic()
        .font_size(16.0)
        .bg_color("#FFFF00")
        .align_horizontal("center")
        .align_vertical("center")
        .build();

    let xlsx = XlsxBuilder::new()
        .sheet(
            SheetBuilder::new("Sheet1")
                .cell("A1", "Styled Merge", Some(style))
                .merge("A1:C3"),
        )
        .build();

    let workbook = load_xlsx(&xlsx);

    assert_cell_value(&workbook, 0, 0, 0, "Styled Merge");
    assert_merge_exists(&workbook, 0, 0, 0, 2, 2);

    // Verify style on top-left cell
    let cell_style = get_cell_style(&workbook, 0, 0, 0).expect("Style should exist");
    assert_eq!(cell_style.bold, Some(true));
    assert_eq!(cell_style.italic, Some(true));
}

/// Test 8b: Merge with border style
#[test]
fn test_merge_with_border_style() {
    let style = StyleBuilder::new()
        .border_all("thin", Some("#000000"))
        .build();

    let xlsx = XlsxBuilder::new()
        .sheet(
            SheetBuilder::new("Sheet1")
                .cell("A1", "Bordered Merge", Some(style))
                .merge("A1:B2"),
        )
        .build();

    let workbook = load_xlsx(&xlsx);

    assert_cell_value(&workbook, 0, 0, 0, "Bordered Merge");
    assert_merge_exists(&workbook, 0, 0, 0, 1, 1);

    // Verify border style exists
    let cell_style = get_cell_style(&workbook, 0, 0, 0).expect("Style should exist");
    assert!(cell_style.border_top.is_some());
    assert!(cell_style.border_bottom.is_some());
    assert!(cell_style.border_left.is_some());
    assert!(cell_style.border_right.is_some());
}

/// Test 8c: Merge with wrap text
#[test]
fn test_merge_with_wrap_text() {
    let style = StyleBuilder::new().wrap_text().build();

    let xlsx = XlsxBuilder::new()
        .sheet(
            SheetBuilder::new("Sheet1")
                .cell(
                    "A1",
                    "This is a long text that should wrap within the merged cell",
                    Some(style),
                )
                .merge("A1:B3"),
        )
        .build();

    let workbook = load_xlsx(&xlsx);

    assert_merge_exists(&workbook, 0, 0, 0, 2, 1);

    let cell_style = get_cell_style(&workbook, 0, 0, 0).expect("Style should exist");
    assert_eq!(cell_style.wrap, Some(true));
}

// ============================================================================
// MERGE WITH NUMBER FORMATTING TESTS
// ============================================================================

/// Test 9: Merge with number formatting
#[test]
fn test_merge_with_number_formatting() {
    let style = StyleBuilder::new().number_format("#,##0.00").build();

    let xlsx = XlsxBuilder::new()
        .sheet(
            SheetBuilder::new("Sheet1")
                .cell("A1", 1234567.89, Some(style))
                .merge("A1:C1"),
        )
        .build();

    let workbook = load_xlsx(&xlsx);

    assert_merge_exists(&workbook, 0, 0, 0, 0, 2);

    // Verify the cell has a value (may be in raw, v, or cached_display)
    let cd = get_cell(&workbook, 0, 0, 0).expect("Cell should exist");
    assert!(cd.cell.cached_display.is_some() || cd.cell.v.is_some() || cd.cell.raw.is_some());
}

/// Test 9b: Merge with percentage format
#[test]
fn test_merge_with_percentage_format() {
    let style = StyleBuilder::new().number_format("0.00%").build();

    let xlsx = XlsxBuilder::new()
        .sheet(
            SheetBuilder::new("Sheet1")
                .cell("A1", 0.7525, Some(style))
                .merge("A1:B1"),
        )
        .build();

    let workbook = load_xlsx(&xlsx);

    assert_merge_exists(&workbook, 0, 0, 0, 0, 1);

    let cd = get_cell(&workbook, 0, 0, 0).expect("Cell should exist");
    assert!(cd.cell.cached_display.is_some() || cd.cell.v.is_some() || cd.cell.raw.is_some());
}

/// Test 9c: Merge with date format
#[test]
fn test_merge_with_date_format() {
    let style = StyleBuilder::new().number_format("mm-dd-yy").build();

    let xlsx = XlsxBuilder::new()
        .sheet(
            SheetBuilder::new("Sheet1")
                .cell("A1", 45306.0, Some(style))
                .merge("A1:C1"),
        )
        .build();

    let workbook = load_xlsx(&xlsx);

    assert_merge_exists(&workbook, 0, 0, 0, 0, 2);
}

/// Test 9d: Merge with currency format
#[test]
fn test_merge_with_currency_format() {
    let style = StyleBuilder::new()
        .number_format("$#,##0.00")
        .align_horizontal("right")
        .build();

    let xlsx = XlsxBuilder::new()
        .sheet(
            SheetBuilder::new("Sheet1")
                .cell("A1", 9999.99, Some(style))
                .merge("A1:B1"),
        )
        .build();

    let workbook = load_xlsx(&xlsx);

    assert_merge_exists(&workbook, 0, 0, 0, 0, 1);
}

// ============================================================================
// STRESS TEST
// ============================================================================

/// Test 10: Many merges (stress test with 100 merges)
#[test]
fn test_many_merges_stress_test() {
    let mut sheet = SheetBuilder::new("Sheet1");

    for row in 0..10 {
        for col_group in 0..10 {
            let start_col = col_group * 2;
            let cell_ref = format!("{}{}", column_letter(start_col), row + 1);
            let end_ref = format!("{}{}", column_letter(start_col + 1), row + 1);
            let merge_ref = format!("{}:{}", cell_ref, end_ref);

            sheet = sheet
                .cell(&cell_ref, format!("Merge {}-{}", row, col_group), None)
                .merge(&merge_ref);
        }
    }

    let xlsx = XlsxBuilder::new().sheet(sheet).build();

    let workbook = load_xlsx(&xlsx);

    // Verify all 100 merges exist
    let merges = &workbook.sheets[0].merges;
    assert_eq!(merges.len(), 100, "Should have 100 merges");

    // Spot-check a few merges
    assert_merge_exists(&workbook, 0, 0, 0, 0, 1);
    assert_merge_exists(&workbook, 0, 4, 10, 4, 11);
    assert_merge_exists(&workbook, 0, 9, 18, 9, 19);
}

/// Test 10b: Mixed size merges stress test
#[test]
fn test_mixed_size_merges_stress_test() {
    let xlsx = XlsxBuilder::new()
        .sheet(
            SheetBuilder::new("Sheet1")
                .cell("A1", "1x1", None)
                .cell("A2", "1x2", None)
                .merge("A2:B2")
                .cell("A3", "2x1", None)
                .merge("A3:A4")
                .cell("C1", "2x2", None)
                .merge("C1:D2")
                .cell("E1", "3x3", None)
                .merge("E1:G3")
                .cell("A6", "1x5", None)
                .merge("A6:E6")
                .cell("H1", "5x1", None)
                .merge("H1:H5")
                .cell("I1", "4x4", None)
                .merge("I1:L4")
                .cell("A8", "2x8", None)
                .merge("A8:H9")
                .cell("M1", "8x2", None)
                .merge("M1:N8"),
        )
        .build();

    let workbook = load_xlsx(&xlsx);

    let merges = &workbook.sheets[0].merges;
    assert_eq!(merges.len(), 9, "Should have 9 merges");

    assert_merge_exists(&workbook, 0, 1, 0, 1, 1); // A2:B2
    assert_merge_exists(&workbook, 0, 2, 0, 3, 0); // A3:A4
    assert_merge_exists(&workbook, 0, 0, 2, 1, 3); // C1:D2
    assert_merge_exists(&workbook, 0, 0, 4, 2, 6); // E1:G3
    assert_merge_exists(&workbook, 0, 5, 0, 5, 4); // A6:E6
    assert_merge_exists(&workbook, 0, 0, 7, 4, 7); // H1:H5
    assert_merge_exists(&workbook, 0, 0, 8, 3, 11); // I1:L4
    assert_merge_exists(&workbook, 0, 7, 0, 8, 7); // A8:H9
    assert_merge_exists(&workbook, 0, 0, 12, 7, 13); // M1:N8
}

// ============================================================================
// EDGE CASE TESTS
// ============================================================================

/// Test: Merge starting at higher row/column
#[test]
fn test_merge_not_at_origin() {
    let xlsx = XlsxBuilder::new()
        .sheet(
            SheetBuilder::new("Sheet1")
                .cell("D5", "Offset Merge", None)
                .merge("D5:F7"),
        )
        .build();

    let workbook = load_xlsx(&xlsx);

    // D5:F7 = (4, 3) to (6, 5)
    assert_merge_exists(&workbook, 0, 4, 3, 6, 5);
}

/// Test: Merges in multiple sheets
#[test]
fn test_merges_in_multiple_sheets() {
    let xlsx = XlsxBuilder::new()
        .sheet(
            SheetBuilder::new("Sheet1")
                .cell("A1", "Sheet1 Merge", None)
                .merge("A1:B2"),
        )
        .sheet(
            SheetBuilder::new("Sheet2")
                .cell("C3", "Sheet2 Merge", None)
                .merge("C3:E5"),
        )
        .sheet(
            SheetBuilder::new("Sheet3")
                .cell("F1", "Sheet3 Merge", None)
                .merge("F1:H1"),
        )
        .build();

    let workbook = load_xlsx(&xlsx);

    assert_sheet_count(&workbook, 3);

    // Sheet1: A1:B2
    assert_merge_exists(&workbook, 0, 0, 0, 1, 1);
    assert_eq!(workbook.sheets[0].merges.len(), 1);

    // Sheet2: C3:E5
    assert_merge_exists(&workbook, 1, 2, 2, 4, 4);
    assert_eq!(workbook.sheets[1].merges.len(), 1);

    // Sheet3: F1:H1
    assert_merge_exists(&workbook, 2, 0, 5, 0, 7);
    assert_eq!(workbook.sheets[2].merges.len(), 1);
}

/// Test: Very wide merge (26+ columns, crossing letter boundary)
#[test]
fn test_very_wide_merge_crossing_z() {
    let xlsx = XlsxBuilder::new()
        .sheet(
            SheetBuilder::new("Sheet1")
                .cell("A1", "Super Wide", None)
                .merge("A1:AD1"),
        )
        .build();

    let workbook = load_xlsx(&xlsx);

    // A1:AD1 = (0, 0) to (0, 29)
    assert_merge_exists(&workbook, 0, 0, 0, 0, 29);
}

/// Test: Merge in AA column range
#[test]
fn test_merge_in_aa_column_range() {
    let xlsx = XlsxBuilder::new()
        .sheet(
            SheetBuilder::new("Sheet1")
                .cell("AA1", "Double Letter", None)
                .merge("AA1:AC3"),
        )
        .build();

    let workbook = load_xlsx(&xlsx);

    // AA1:AC3 = (0, 26) to (2, 28)
    assert_merge_exists(&workbook, 0, 0, 26, 2, 28);
}

/// Test: Large row number merge
#[test]
fn test_large_row_number_merge() {
    let xlsx = XlsxBuilder::new()
        .sheet(
            SheetBuilder::new("Sheet1")
                .cell("A1000", "High Row", None)
                .merge("A1000:C1005"),
        )
        .build();

    let workbook = load_xlsx(&xlsx);

    // A1000:C1005 = (999, 0) to (1004, 2)
    assert_merge_exists(&workbook, 0, 999, 0, 1004, 2);
}

/// Test: Single column horizontal extent merge
#[test]
fn test_single_column_multi_row_merge() {
    let xlsx = XlsxBuilder::new()
        .sheet(
            SheetBuilder::new("Sheet1")
                .cell("C5", "Vertical Only", None)
                .merge("C5:C15"),
        )
        .build();

    let workbook = load_xlsx(&xlsx);

    // C5:C15 = (4, 2) to (14, 2)
    assert_merge_exists(&workbook, 0, 4, 2, 14, 2);
}

/// Test: Single row vertical extent merge
#[test]
fn test_single_row_multi_column_merge() {
    let xlsx = XlsxBuilder::new()
        .sheet(
            SheetBuilder::new("Sheet1")
                .cell("B3", "Horizontal Only", None)
                .merge("B3:K3"),
        )
        .build();

    let workbook = load_xlsx(&xlsx);

    // B3:K3 = (2, 1) to (2, 10)
    assert_merge_exists(&workbook, 0, 2, 1, 2, 10);
}

/// Test: Complex style combination on merge
#[test]
fn test_merge_with_complex_style() {
    let style = StyleBuilder::new()
        .font_name("Arial")
        .font_size(18.0)
        .font_color("#FF0000")
        .bold()
        .italic()
        .underline()
        .bg_color("#FFFFCC")
        .border_all("medium", Some("#000000"))
        .align_horizontal("center")
        .align_vertical("center")
        .wrap_text()
        .indent(2)
        .build();

    let xlsx = XlsxBuilder::new()
        .sheet(
            SheetBuilder::new("Sheet1")
                .cell("B2", "Complex Style", Some(style))
                .merge("B2:E5"),
        )
        .build();

    let workbook = load_xlsx(&xlsx);

    // Verify merge
    assert_merge_exists(&workbook, 0, 1, 1, 4, 4);

    // Verify style properties
    let cell_style = get_cell_style(&workbook, 0, 1, 1).expect("Style should exist");
    assert_eq!(cell_style.bold, Some(true));
    assert_eq!(cell_style.italic, Some(true));
    assert!(cell_style.underline.is_some(), "Cell should be underlined");
    assert_eq!(cell_style.wrap, Some(true));
}

/// Test: Empty merge (no content)
#[test]
fn test_empty_merge_no_content() {
    let xlsx = XlsxBuilder::new()
        .sheet(SheetBuilder::new("Sheet1").merge("A1:C3"))
        .build();

    let workbook = load_xlsx(&xlsx);

    // Verify merge exists even without content
    assert_merge_exists(&workbook, 0, 0, 0, 2, 2);
}

/// Test: Merge with only style, no content
#[test]
fn test_merge_with_style_only() {
    let style = StyleBuilder::new().bg_color("#E0E0E0").build();

    let xlsx = XlsxBuilder::new()
        .sheet(
            SheetBuilder::new("Sheet1")
                .styled_cell("A1", style)
                .merge("A1:B2"),
        )
        .build();

    let workbook = load_xlsx(&xlsx);

    assert_merge_exists(&workbook, 0, 0, 0, 1, 1);
}

// ============================================================================
// REALISTIC PATTERNS TESTS
// ============================================================================

/// Test: Invoice header pattern
#[test]
fn test_invoice_header_pattern() {
    let title_style = StyleBuilder::new()
        .bold()
        .font_size(24.0)
        .align_horizontal("center")
        .build();

    let subtitle_style = StyleBuilder::new()
        .italic()
        .align_horizontal("center")
        .build();

    let xlsx = XlsxBuilder::new()
        .sheet(
            SheetBuilder::new("Invoice")
                .cell("A1", "COMPANY", Some(title_style.clone()))
                .merge("A1:B3")
                .cell("C1", "INVOICE", Some(title_style))
                .merge("C1:F1")
                .cell("C2", "Invoice #12345", Some(subtitle_style.clone()))
                .merge("C2:F2")
                .cell("C3", "Date: 2024-01-15", Some(subtitle_style))
                .merge("C3:F3"),
        )
        .build();

    let workbook = load_xlsx(&xlsx);

    // Verify all invoice header merges
    assert_merge_exists(&workbook, 0, 0, 0, 2, 1); // Logo area
    assert_merge_exists(&workbook, 0, 0, 2, 0, 5); // INVOICE title
    assert_merge_exists(&workbook, 0, 1, 2, 1, 5); // Invoice number
    assert_merge_exists(&workbook, 0, 2, 2, 2, 5); // Date
}

/// Test: Table with grouped headers
#[test]
fn test_table_grouped_headers() {
    let group_style = StyleBuilder::new()
        .bold()
        .bg_color("#4472C4")
        .font_color("#FFFFFF")
        .align_horizontal("center")
        .build();

    let subheader_style = StyleBuilder::new()
        .bold()
        .bg_color("#8EA9DB")
        .align_horizontal("center")
        .build();

    let xlsx = XlsxBuilder::new()
        .sheet(
            SheetBuilder::new("Report")
                .cell("A1", "ID", Some(group_style.clone()))
                .merge("A1:A2")
                .cell("B1", "Q1 2024", Some(group_style.clone()))
                .merge("B1:D1")
                .cell("E1", "Q2 2024", Some(group_style))
                .merge("E1:G1")
                .cell("B2", "Jan", Some(subheader_style.clone()))
                .cell("C2", "Feb", Some(subheader_style.clone()))
                .cell("D2", "Mar", Some(subheader_style.clone()))
                .cell("E2", "Apr", Some(subheader_style.clone()))
                .cell("F2", "May", Some(subheader_style.clone()))
                .cell("G2", "Jun", Some(subheader_style)),
        )
        .build();

    let workbook = load_xlsx(&xlsx);

    // Verify grouped headers
    assert_merge_exists(&workbook, 0, 0, 0, 1, 0); // ID spans rows
    assert_merge_exists(&workbook, 0, 0, 1, 0, 3); // Q1 spans columns
    assert_merge_exists(&workbook, 0, 0, 4, 0, 6); // Q2 spans columns
}

/// Test: Matrix-style report with row and column merges
#[test]
fn test_matrix_style_report() {
    let xlsx = XlsxBuilder::new()
        .sheet(
            SheetBuilder::new("Matrix")
                .merge("A1:B2")
                .cell("C1", "Region A", None)
                .merge("C1:D1")
                .cell("E1", "Region B", None)
                .merge("E1:F1")
                .cell("C2", "Sales", None)
                .cell("D2", "Returns", None)
                .cell("E2", "Sales", None)
                .cell("F2", "Returns", None)
                .cell("A3", "Product Category 1", None)
                .merge("A3:A5")
                .cell("A6", "Product Category 2", None)
                .merge("A6:A8")
                .cell("B3", "Type A", None)
                .cell("B4", "Type B", None)
                .cell("B5", "Type C", None)
                .cell("B6", "Type X", None)
                .cell("B7", "Type Y", None)
                .cell("B8", "Type Z", None),
        )
        .build();

    let workbook = load_xlsx(&xlsx);

    // Verify matrix merges
    assert_merge_exists(&workbook, 0, 0, 0, 1, 1); // Corner
    assert_merge_exists(&workbook, 0, 0, 2, 0, 3); // Region A
    assert_merge_exists(&workbook, 0, 0, 4, 0, 5); // Region B
    assert_merge_exists(&workbook, 0, 2, 0, 4, 0); // Category 1
    assert_merge_exists(&workbook, 0, 5, 0, 7, 0); // Category 2
}

// ============================================================================
// HELPER FUNCTIONS
// ============================================================================

/// Convert a 0-indexed column number to Excel column letter(s)
fn column_letter(col: u32) -> String {
    let mut result = String::new();
    let mut n = col;

    loop {
        let remainder = n % 26;
        result.insert(0, (b'A' + remainder as u8) as char);
        if n < 26 {
            break;
        }
        n = n / 26 - 1;
    }

    result
}

#[cfg(test)]
mod helper_tests {
    use super::*;

    #[test]
    fn test_column_letter_single() {
        assert_eq!(column_letter(0), "A");
        assert_eq!(column_letter(1), "B");
        assert_eq!(column_letter(25), "Z");
    }

    #[test]
    fn test_column_letter_double() {
        assert_eq!(column_letter(26), "AA");
        assert_eq!(column_letter(27), "AB");
        assert_eq!(column_letter(51), "AZ");
        assert_eq!(column_letter(52), "BA");
    }

    #[test]
    fn test_column_letter_triple() {
        assert_eq!(column_letter(702), "AAA");
    }
}
