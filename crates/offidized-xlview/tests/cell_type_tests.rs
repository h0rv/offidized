//! Cell type tests for offidized-xlview
//!
//! Tests for parsing different cell types including:
//! - String cells (shared string reference)
//! - Inline string cells
//! - Number cells
//! - Boolean cells
//! - Error cells
//! - Formula cells with string result
//! - Empty cells
//! - Cells with style but no value
//! - Cells with whitespace-only content
//! - Very large/small numbers
//! - Negative numbers

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

use common::{get_cell, load_xlsx};
use fixtures::{CellValue, SheetBuilder, StyleBuilder, XlsxBuilder};
use offidized_xlview::types::workbook::{CellRawValue, CellType};

// ============================================================================
// Helper Functions
// ============================================================================

/// Get cell type from parsed workbook.
fn get_cell_type(
    workbook: &offidized_xlview::types::workbook::Workbook,
    sheet: usize,
    row: u32,
    col: u32,
) -> Option<CellType> {
    let cd = get_cell(workbook, sheet, row, col)?;
    Some(cd.cell.t)
}

/// Get cell display value from parsed workbook.
fn get_cell_value(
    workbook: &offidized_xlview::types::workbook::Workbook,
    sheet: usize,
    row: u32,
    col: u32,
) -> Option<String> {
    let cd = get_cell(workbook, sheet, row, col)?;

    // Check cached display first
    if let Some(ref display) = cd.cell.cached_display {
        return Some(display.clone());
    }

    // Check explicit value
    if let Some(ref v) = cd.cell.v {
        return Some(v.clone());
    }

    // Resolve from raw value
    match cd.cell.raw.as_ref()? {
        CellRawValue::String(s) => Some(s.clone()),
        CellRawValue::Number(n) => Some(n.to_string()),
        CellRawValue::Boolean(b) => Some(if *b { "TRUE" } else { "FALSE" }.to_string()),
        CellRawValue::Error(e) => Some(e.clone()),
        CellRawValue::Date(n) => Some(n.to_string()),
        CellRawValue::SharedString(idx) => workbook.shared_strings.get(*idx as usize).cloned(),
    }
}

/// Check if cell exists in parsed workbook.
fn cell_exists(
    workbook: &offidized_xlview::types::workbook::Workbook,
    sheet: usize,
    row: u32,
    col: u32,
) -> bool {
    get_cell(workbook, sheet, row, col).is_some()
}

/// Check if cell has a style.
fn cell_has_style(
    workbook: &offidized_xlview::types::workbook::Workbook,
    sheet: usize,
    row: u32,
    col: u32,
) -> bool {
    let cd = match get_cell(workbook, sheet, row, col) {
        Some(c) => c,
        None => return false,
    };
    // Has inline style or a style index
    cd.cell.s.is_some() || cd.cell.style_idx.is_some()
}

// ============================================================================
// 1. String Cells (Shared String Reference)
// ============================================================================

#[test]
fn test_string_cell_simple() {
    let xlsx = XlsxBuilder::new()
        .add_sheet("Sheet1")
        .add_cell("A1", "Hello World", None)
        .build();

    let workbook = load_xlsx(&xlsx);
    assert!(matches!(
        get_cell_type(&workbook, 0, 0, 0),
        Some(CellType::String)
    ));
    assert_eq!(
        get_cell_value(&workbook, 0, 0, 0),
        Some("Hello World".to_string())
    );
}

#[test]
fn test_string_cell_empty_string() {
    let xlsx = XlsxBuilder::new()
        .add_sheet("Sheet1")
        .add_cell("A1", "", None)
        .build();

    let workbook = load_xlsx(&xlsx);
    assert!(matches!(
        get_cell_type(&workbook, 0, 0, 0),
        Some(CellType::String)
    ));
    assert_eq!(get_cell_value(&workbook, 0, 0, 0), Some("".to_string()));
}

#[test]
fn test_string_cell_with_numbers_as_text() {
    let xlsx = XlsxBuilder::new()
        .add_sheet("Sheet1")
        .add_cell("A1", "12345", None)
        .build();

    let workbook = load_xlsx(&xlsx);
    assert!(matches!(
        get_cell_type(&workbook, 0, 0, 0),
        Some(CellType::String)
    ));
    assert_eq!(
        get_cell_value(&workbook, 0, 0, 0),
        Some("12345".to_string())
    );
}

#[test]
fn test_string_cell_with_special_characters() {
    let xlsx = XlsxBuilder::new()
        .add_sheet("Sheet1")
        .add_cell("A1", "<>&\"'", None)
        .build();

    let workbook = load_xlsx(&xlsx);
    assert!(matches!(
        get_cell_type(&workbook, 0, 0, 0),
        Some(CellType::String)
    ));
    assert_eq!(
        get_cell_value(&workbook, 0, 0, 0),
        Some("<>&\"'".to_string())
    );
}

#[test]
fn test_string_cell_with_unicode() {
    let xlsx = XlsxBuilder::new()
        .add_sheet("Sheet1")
        .add_cell(
            "A1",
            "Hello \u{4e16}\u{754c} \u{041f}\u{0440}\u{0438}\u{0432}\u{0435}\u{0442}",
            None,
        )
        .build();

    let workbook = load_xlsx(&xlsx);
    assert!(matches!(
        get_cell_type(&workbook, 0, 0, 0),
        Some(CellType::String)
    ));
    assert_eq!(
        get_cell_value(&workbook, 0, 0, 0),
        Some("Hello \u{4e16}\u{754c} \u{041f}\u{0440}\u{0438}\u{0432}\u{0435}\u{0442}".to_string())
    );
}

#[test]
fn test_string_cell_with_newline() {
    let xlsx = XlsxBuilder::new()
        .add_sheet("Sheet1")
        .add_cell("A1", "Line1\nLine2\nLine3", None)
        .build();

    let workbook = load_xlsx(&xlsx);
    assert!(matches!(
        get_cell_type(&workbook, 0, 0, 0),
        Some(CellType::String)
    ));
    assert_eq!(
        get_cell_value(&workbook, 0, 0, 0),
        Some("Line1\nLine2\nLine3".to_string())
    );
}

#[test]
fn test_string_cell_very_long() {
    let long_string = "A".repeat(10000);
    let xlsx = XlsxBuilder::new()
        .add_sheet("Sheet1")
        .add_cell("A1", long_string.as_str(), None)
        .build();

    let workbook = load_xlsx(&xlsx);
    assert!(matches!(
        get_cell_type(&workbook, 0, 0, 0),
        Some(CellType::String)
    ));
    assert_eq!(get_cell_value(&workbook, 0, 0, 0), Some(long_string));
}

#[test]
fn test_shared_string_reuse() {
    // Same string value in multiple cells should share the string
    let xlsx = XlsxBuilder::new()
        .add_sheet("Sheet1")
        .add_cell("A1", "Shared Value", None)
        .add_cell("B1", "Shared Value", None)
        .add_cell("C1", "Shared Value", None)
        .build();

    let workbook = load_xlsx(&xlsx);
    assert_eq!(
        get_cell_value(&workbook, 0, 0, 0),
        Some("Shared Value".to_string())
    );
    assert_eq!(
        get_cell_value(&workbook, 0, 0, 1),
        Some("Shared Value".to_string())
    );
    assert_eq!(
        get_cell_value(&workbook, 0, 0, 2),
        Some("Shared Value".to_string())
    );
}

// ============================================================================
// 2. Inline String Cells
// ============================================================================

#[test]
fn test_inline_string_simple() {
    let xlsx = XlsxBuilder::new()
        .sheet(SheetBuilder::new("Sheet1").cell(
            "A1",
            CellValue::InlineString("Inline Text".to_string()),
            None,
        ))
        .build();

    let workbook = load_xlsx(&xlsx);
    // Inline strings should also be parsed as string type
    assert!(matches!(
        get_cell_type(&workbook, 0, 0, 0),
        Some(CellType::String)
    ));
    assert_eq!(
        get_cell_value(&workbook, 0, 0, 0),
        Some("Inline Text".to_string())
    );
}

#[test]
fn test_inline_string_empty() {
    let xlsx = XlsxBuilder::new()
        .sheet(SheetBuilder::new("Sheet1").cell(
            "A1",
            CellValue::InlineString("".to_string()),
            None,
        ))
        .build();

    let workbook = load_xlsx(&xlsx);
    assert!(matches!(
        get_cell_type(&workbook, 0, 0, 0),
        Some(CellType::String)
    ));
}

#[test]
fn test_inline_string_with_special_chars() {
    let xlsx = XlsxBuilder::new()
        .sheet(SheetBuilder::new("Sheet1").cell(
            "A1",
            CellValue::InlineString("Test <xml> & \"quotes\"".to_string()),
            None,
        ))
        .build();

    let workbook = load_xlsx(&xlsx);
    assert!(matches!(
        get_cell_type(&workbook, 0, 0, 0),
        Some(CellType::String)
    ));
    assert_eq!(
        get_cell_value(&workbook, 0, 0, 0),
        Some("Test <xml> & \"quotes\"".to_string())
    );
}

// ============================================================================
// 3. Number Cells
// ============================================================================

#[test]
fn test_number_cell_integer() {
    let xlsx = XlsxBuilder::new()
        .add_sheet("Sheet1")
        .add_cell("A1", 42.0, None)
        .build();

    let workbook = load_xlsx(&xlsx);
    assert!(matches!(
        get_cell_type(&workbook, 0, 0, 0),
        Some(CellType::Number)
    ));
    let value = get_cell_value(&workbook, 0, 0, 0);
    assert!(value.is_some());
}

#[test]
fn test_number_cell_decimal() {
    let xlsx = XlsxBuilder::new()
        .add_sheet("Sheet1")
        .add_cell("A1", 3.14159, None)
        .build();

    let workbook = load_xlsx(&xlsx);
    assert!(matches!(
        get_cell_type(&workbook, 0, 0, 0),
        Some(CellType::Number)
    ));
}

#[test]
fn test_number_cell_zero() {
    let xlsx = XlsxBuilder::new()
        .add_sheet("Sheet1")
        .add_cell("A1", 0.0, None)
        .build();

    let workbook = load_xlsx(&xlsx);
    assert!(matches!(
        get_cell_type(&workbook, 0, 0, 0),
        Some(CellType::Number)
    ));
}

#[test]
fn test_number_cell_negative() {
    let xlsx = XlsxBuilder::new()
        .add_sheet("Sheet1")
        .add_cell("A1", -42.5, None)
        .build();

    let workbook = load_xlsx(&xlsx);
    assert!(matches!(
        get_cell_type(&workbook, 0, 0, 0),
        Some(CellType::Number)
    ));
}

#[test]
fn test_number_cell_large_negative() {
    let xlsx = XlsxBuilder::new()
        .add_sheet("Sheet1")
        .add_cell("A1", -999999999.99, None)
        .build();

    let workbook = load_xlsx(&xlsx);
    assert!(matches!(
        get_cell_type(&workbook, 0, 0, 0),
        Some(CellType::Number)
    ));
}

// ============================================================================
// 4. Boolean Cells
// ============================================================================

#[test]
fn test_boolean_cell_true() {
    let xlsx = XlsxBuilder::new()
        .sheet(SheetBuilder::new("Sheet1").cell("A1", CellValue::Boolean(true), None))
        .build();

    let workbook = load_xlsx(&xlsx);
    assert!(matches!(
        get_cell_type(&workbook, 0, 0, 0),
        Some(CellType::Boolean)
    ));
    assert_eq!(get_cell_value(&workbook, 0, 0, 0), Some("TRUE".to_string()));
}

#[test]
fn test_boolean_cell_false() {
    let xlsx = XlsxBuilder::new()
        .sheet(SheetBuilder::new("Sheet1").cell("A1", CellValue::Boolean(false), None))
        .build();

    let workbook = load_xlsx(&xlsx);
    assert!(matches!(
        get_cell_type(&workbook, 0, 0, 0),
        Some(CellType::Boolean)
    ));
    assert_eq!(
        get_cell_value(&workbook, 0, 0, 0),
        Some("FALSE".to_string())
    );
}

#[test]
fn test_boolean_cells_multiple() {
    let xlsx = XlsxBuilder::new()
        .sheet(
            SheetBuilder::new("Sheet1")
                .cell("A1", CellValue::Boolean(true), None)
                .cell("B1", CellValue::Boolean(false), None)
                .cell("A2", CellValue::Boolean(false), None)
                .cell("B2", CellValue::Boolean(true), None),
        )
        .build();

    let workbook = load_xlsx(&xlsx);
    assert!(matches!(
        get_cell_type(&workbook, 0, 0, 0),
        Some(CellType::Boolean)
    ));
    assert!(matches!(
        get_cell_type(&workbook, 0, 0, 1),
        Some(CellType::Boolean)
    ));
    assert!(matches!(
        get_cell_type(&workbook, 0, 1, 0),
        Some(CellType::Boolean)
    ));
    assert!(matches!(
        get_cell_type(&workbook, 0, 1, 1),
        Some(CellType::Boolean)
    ));

    assert_eq!(get_cell_value(&workbook, 0, 0, 0), Some("TRUE".to_string()));
    assert_eq!(
        get_cell_value(&workbook, 0, 0, 1),
        Some("FALSE".to_string())
    );
    assert_eq!(
        get_cell_value(&workbook, 0, 1, 0),
        Some("FALSE".to_string())
    );
    assert_eq!(get_cell_value(&workbook, 0, 1, 1), Some("TRUE".to_string()));
}

// ============================================================================
// 5. Error Cells
// ============================================================================

#[test]
fn test_error_cell_value() {
    let xlsx = XlsxBuilder::new()
        .sheet(SheetBuilder::new("Sheet1").cell(
            "A1",
            CellValue::Error("#VALUE!".to_string()),
            None,
        ))
        .build();

    let workbook = load_xlsx(&xlsx);
    assert!(matches!(
        get_cell_type(&workbook, 0, 0, 0),
        Some(CellType::Error)
    ));
    assert_eq!(
        get_cell_value(&workbook, 0, 0, 0),
        Some("#VALUE!".to_string())
    );
}

#[test]
fn test_error_cell_ref() {
    let xlsx = XlsxBuilder::new()
        .sheet(SheetBuilder::new("Sheet1").cell("A1", CellValue::Error("#REF!".to_string()), None))
        .build();

    let workbook = load_xlsx(&xlsx);
    assert!(matches!(
        get_cell_type(&workbook, 0, 0, 0),
        Some(CellType::Error)
    ));
    assert_eq!(
        get_cell_value(&workbook, 0, 0, 0),
        Some("#REF!".to_string())
    );
}

#[test]
fn test_error_cell_name() {
    let xlsx = XlsxBuilder::new()
        .sheet(SheetBuilder::new("Sheet1").cell("A1", CellValue::Error("#NAME?".to_string()), None))
        .build();

    let workbook = load_xlsx(&xlsx);
    assert!(matches!(
        get_cell_type(&workbook, 0, 0, 0),
        Some(CellType::Error)
    ));
    assert_eq!(
        get_cell_value(&workbook, 0, 0, 0),
        Some("#NAME?".to_string())
    );
}

#[test]
fn test_error_cell_div0() {
    let xlsx = XlsxBuilder::new()
        .sheet(SheetBuilder::new("Sheet1").cell(
            "A1",
            CellValue::Error("#DIV/0!".to_string()),
            None,
        ))
        .build();

    let workbook = load_xlsx(&xlsx);
    assert!(matches!(
        get_cell_type(&workbook, 0, 0, 0),
        Some(CellType::Error)
    ));
    assert_eq!(
        get_cell_value(&workbook, 0, 0, 0),
        Some("#DIV/0!".to_string())
    );
}

#[test]
fn test_error_cell_null() {
    let xlsx = XlsxBuilder::new()
        .sheet(SheetBuilder::new("Sheet1").cell("A1", CellValue::Error("#NULL!".to_string()), None))
        .build();

    let workbook = load_xlsx(&xlsx);
    assert!(matches!(
        get_cell_type(&workbook, 0, 0, 0),
        Some(CellType::Error)
    ));
    assert_eq!(
        get_cell_value(&workbook, 0, 0, 0),
        Some("#NULL!".to_string())
    );
}

#[test]
fn test_error_cell_na() {
    let xlsx = XlsxBuilder::new()
        .sheet(SheetBuilder::new("Sheet1").cell("A1", CellValue::Error("#N/A".to_string()), None))
        .build();

    let workbook = load_xlsx(&xlsx);
    assert!(matches!(
        get_cell_type(&workbook, 0, 0, 0),
        Some(CellType::Error)
    ));
    assert_eq!(get_cell_value(&workbook, 0, 0, 0), Some("#N/A".to_string()));
}

#[test]
fn test_error_cell_num() {
    let xlsx = XlsxBuilder::new()
        .sheet(SheetBuilder::new("Sheet1").cell("A1", CellValue::Error("#NUM!".to_string()), None))
        .build();

    let workbook = load_xlsx(&xlsx);
    assert!(matches!(
        get_cell_type(&workbook, 0, 0, 0),
        Some(CellType::Error)
    ));
    assert_eq!(
        get_cell_value(&workbook, 0, 0, 0),
        Some("#NUM!".to_string())
    );
}

#[test]
fn test_error_cells_all_types() {
    let xlsx = XlsxBuilder::new()
        .sheet(
            SheetBuilder::new("Sheet1")
                .cell("A1", CellValue::Error("#VALUE!".to_string()), None)
                .cell("A2", CellValue::Error("#REF!".to_string()), None)
                .cell("A3", CellValue::Error("#NAME?".to_string()), None)
                .cell("A4", CellValue::Error("#DIV/0!".to_string()), None)
                .cell("A5", CellValue::Error("#NULL!".to_string()), None)
                .cell("A6", CellValue::Error("#N/A".to_string()), None)
                .cell("A7", CellValue::Error("#NUM!".to_string()), None),
        )
        .build();

    let workbook = load_xlsx(&xlsx);

    // All should be error type
    for row in 0..7 {
        assert!(
            matches!(get_cell_type(&workbook, 0, row, 0), Some(CellType::Error)),
            "Row {} should be error type",
            row
        );
    }

    // Check each error value
    assert_eq!(
        get_cell_value(&workbook, 0, 0, 0),
        Some("#VALUE!".to_string())
    );
    assert_eq!(
        get_cell_value(&workbook, 0, 1, 0),
        Some("#REF!".to_string())
    );
    assert_eq!(
        get_cell_value(&workbook, 0, 2, 0),
        Some("#NAME?".to_string())
    );
    assert_eq!(
        get_cell_value(&workbook, 0, 3, 0),
        Some("#DIV/0!".to_string())
    );
    assert_eq!(
        get_cell_value(&workbook, 0, 4, 0),
        Some("#NULL!".to_string())
    );
    assert_eq!(get_cell_value(&workbook, 0, 5, 0), Some("#N/A".to_string()));
    assert_eq!(
        get_cell_value(&workbook, 0, 6, 0),
        Some("#NUM!".to_string())
    );
}

// ============================================================================
// 6. Formula Cells with String Result
// ============================================================================

#[test]
fn test_string_formula_result_via_inline() {
    // Simulate a formula that returns a string by using inline string
    let xlsx = XlsxBuilder::new()
        .sheet(SheetBuilder::new("Sheet1").cell(
            "A1",
            CellValue::InlineString("Formula Result".to_string()),
            None,
        ))
        .build();

    let workbook = load_xlsx(&xlsx);
    // Should be treated as string type
    assert!(matches!(
        get_cell_type(&workbook, 0, 0, 0),
        Some(CellType::String)
    ));
}

// ============================================================================
// 7. Empty Cells
// ============================================================================

#[test]
fn test_empty_cell_no_style() {
    let xlsx = XlsxBuilder::new()
        .sheet(SheetBuilder::new("Sheet1").cell("A1", CellValue::Empty, None))
        .build();

    let workbook = load_xlsx(&xlsx);
    // Empty cell should still exist in the output
    assert!(cell_exists(&workbook, 0, 0, 0));
}

#[test]
fn test_empty_cells_sparse() {
    // Only cells A1 and C3 have values, B2 is empty
    let xlsx = XlsxBuilder::new()
        .add_sheet("Sheet1")
        .add_cell("A1", "First", None)
        .add_cell("C3", "Last", None)
        .build();

    let workbook = load_xlsx(&xlsx);

    // A1 and C3 should exist
    assert!(cell_exists(&workbook, 0, 0, 0));
    assert!(cell_exists(&workbook, 0, 2, 2));

    // B2 should not exist (sparse representation)
    assert!(!cell_exists(&workbook, 0, 1, 1));
}

// ============================================================================
// 8. Cells with Style but No Value
// ============================================================================

#[test]
fn test_styled_empty_cell() {
    let style = StyleBuilder::new().bold().bg_color("#FFFF00").build();

    let xlsx = XlsxBuilder::new()
        .sheet(SheetBuilder::new("Sheet1").styled_cell("A1", style))
        .build();

    let workbook = load_xlsx(&xlsx);

    // Cell should exist
    assert!(cell_exists(&workbook, 0, 0, 0));

    // Cell should have a style
    assert!(cell_has_style(&workbook, 0, 0, 0));
}

#[test]
fn test_styled_empty_cell_with_border() {
    let style = StyleBuilder::new()
        .border_all("thin", Some("#000000"))
        .build();

    let xlsx = XlsxBuilder::new()
        .sheet(SheetBuilder::new("Sheet1").styled_cell("A1", style))
        .build();

    let workbook = load_xlsx(&xlsx);
    assert!(cell_exists(&workbook, 0, 0, 0));
    assert!(cell_has_style(&workbook, 0, 0, 0));
}

#[test]
fn test_styled_empty_cell_with_alignment() {
    let style = StyleBuilder::new()
        .align_horizontal("center")
        .align_vertical("center")
        .build();

    let xlsx = XlsxBuilder::new()
        .sheet(SheetBuilder::new("Sheet1").styled_cell("A1", style))
        .build();

    let workbook = load_xlsx(&xlsx);
    assert!(cell_exists(&workbook, 0, 0, 0));
    assert!(cell_has_style(&workbook, 0, 0, 0));
}

// ============================================================================
// 9. Cells with Whitespace-Only Content
// ============================================================================

#[test]
fn test_whitespace_only_spaces() {
    let xlsx = XlsxBuilder::new()
        .add_sheet("Sheet1")
        .add_cell("A1", "   ", None)
        .build();

    let workbook = load_xlsx(&xlsx);
    assert!(matches!(
        get_cell_type(&workbook, 0, 0, 0),
        Some(CellType::String)
    ));
    assert_eq!(get_cell_value(&workbook, 0, 0, 0), Some("   ".to_string()));
}

#[test]
fn test_whitespace_only_tabs() {
    let xlsx = XlsxBuilder::new()
        .add_sheet("Sheet1")
        .add_cell("A1", "\t\t", None)
        .build();

    let workbook = load_xlsx(&xlsx);
    assert!(matches!(
        get_cell_type(&workbook, 0, 0, 0),
        Some(CellType::String)
    ));
    assert_eq!(get_cell_value(&workbook, 0, 0, 0), Some("\t\t".to_string()));
}

#[test]
fn test_whitespace_only_newlines() {
    let xlsx = XlsxBuilder::new()
        .add_sheet("Sheet1")
        .add_cell("A1", "\n\n", None)
        .build();

    let workbook = load_xlsx(&xlsx);
    assert!(matches!(
        get_cell_type(&workbook, 0, 0, 0),
        Some(CellType::String)
    ));
    assert_eq!(get_cell_value(&workbook, 0, 0, 0), Some("\n\n".to_string()));
}

#[test]
fn test_whitespace_mixed() {
    let xlsx = XlsxBuilder::new()
        .add_sheet("Sheet1")
        .add_cell("A1", " \t\n ", None)
        .build();

    let workbook = load_xlsx(&xlsx);
    assert!(matches!(
        get_cell_type(&workbook, 0, 0, 0),
        Some(CellType::String)
    ));
    assert_eq!(
        get_cell_value(&workbook, 0, 0, 0),
        Some(" \t\n ".to_string())
    );
}

#[test]
fn test_leading_trailing_whitespace() {
    let xlsx = XlsxBuilder::new()
        .add_sheet("Sheet1")
        .add_cell("A1", "  Text with spaces  ", None)
        .build();

    let workbook = load_xlsx(&xlsx);
    assert!(matches!(
        get_cell_type(&workbook, 0, 0, 0),
        Some(CellType::String)
    ));
    assert_eq!(
        get_cell_value(&workbook, 0, 0, 0),
        Some("  Text with spaces  ".to_string())
    );
}

// ============================================================================
// 10. Very Large Numbers
// ============================================================================

#[test]
fn test_large_number_integer() {
    let xlsx = XlsxBuilder::new()
        .add_sheet("Sheet1")
        .add_cell("A1", 9999999999999.0, None)
        .build();

    let workbook = load_xlsx(&xlsx);
    assert!(matches!(
        get_cell_type(&workbook, 0, 0, 0),
        Some(CellType::Number)
    ));
}

#[test]
fn test_large_number_scientific() {
    // Excel's maximum precision
    let xlsx = XlsxBuilder::new()
        .add_sheet("Sheet1")
        .add_cell("A1", 1.23e15, None)
        .build();

    let workbook = load_xlsx(&xlsx);
    assert!(matches!(
        get_cell_type(&workbook, 0, 0, 0),
        Some(CellType::Number)
    ));
}

#[test]
fn test_large_number_max_excel() {
    // Close to Excel's maximum value (9.99999999999999E+307)
    let xlsx = XlsxBuilder::new()
        .add_sheet("Sheet1")
        .add_cell("A1", 1.0e100, None)
        .build();

    let workbook = load_xlsx(&xlsx);
    assert!(matches!(
        get_cell_type(&workbook, 0, 0, 0),
        Some(CellType::Number)
    ));
}

#[test]
fn test_large_number_trillion() {
    let xlsx = XlsxBuilder::new()
        .add_sheet("Sheet1")
        .add_cell("A1", 1_000_000_000_000.0, None)
        .build();

    let workbook = load_xlsx(&xlsx);
    assert!(matches!(
        get_cell_type(&workbook, 0, 0, 0),
        Some(CellType::Number)
    ));
}

// ============================================================================
// 11. Very Small Numbers (Scientific Notation)
// ============================================================================

#[test]
fn test_small_number_decimal() {
    let xlsx = XlsxBuilder::new()
        .add_sheet("Sheet1")
        .add_cell("A1", 0.000001, None)
        .build();

    let workbook = load_xlsx(&xlsx);
    assert!(matches!(
        get_cell_type(&workbook, 0, 0, 0),
        Some(CellType::Number)
    ));
}

#[test]
fn test_small_number_scientific() {
    let xlsx = XlsxBuilder::new()
        .add_sheet("Sheet1")
        .add_cell("A1", 1.23e-10, None)
        .build();

    let workbook = load_xlsx(&xlsx);
    assert!(matches!(
        get_cell_type(&workbook, 0, 0, 0),
        Some(CellType::Number)
    ));
}

#[test]
fn test_small_number_very_small() {
    let xlsx = XlsxBuilder::new()
        .add_sheet("Sheet1")
        .add_cell("A1", 1.0e-50, None)
        .build();

    let workbook = load_xlsx(&xlsx);
    assert!(matches!(
        get_cell_type(&workbook, 0, 0, 0),
        Some(CellType::Number)
    ));
}

#[test]
fn test_small_number_min_excel() {
    // Close to Excel's minimum positive value
    let xlsx = XlsxBuilder::new()
        .add_sheet("Sheet1")
        .add_cell("A1", 1.0e-100, None)
        .build();

    let workbook = load_xlsx(&xlsx);
    assert!(matches!(
        get_cell_type(&workbook, 0, 0, 0),
        Some(CellType::Number)
    ));
}

// ============================================================================
// 12. Negative Numbers
// ============================================================================

#[test]
fn test_negative_integer() {
    let xlsx = XlsxBuilder::new()
        .add_sheet("Sheet1")
        .add_cell("A1", -1.0, None)
        .build();

    let workbook = load_xlsx(&xlsx);
    assert!(matches!(
        get_cell_type(&workbook, 0, 0, 0),
        Some(CellType::Number)
    ));
}

#[test]
fn test_negative_decimal() {
    let xlsx = XlsxBuilder::new()
        .add_sheet("Sheet1")
        .add_cell("A1", -123.456, None)
        .build();

    let workbook = load_xlsx(&xlsx);
    assert!(matches!(
        get_cell_type(&workbook, 0, 0, 0),
        Some(CellType::Number)
    ));
}

#[test]
fn test_negative_large() {
    let xlsx = XlsxBuilder::new()
        .add_sheet("Sheet1")
        .add_cell("A1", -1_000_000_000.0, None)
        .build();

    let workbook = load_xlsx(&xlsx);
    assert!(matches!(
        get_cell_type(&workbook, 0, 0, 0),
        Some(CellType::Number)
    ));
}

#[test]
fn test_negative_small() {
    let xlsx = XlsxBuilder::new()
        .add_sheet("Sheet1")
        .add_cell("A1", -0.0001, None)
        .build();

    let workbook = load_xlsx(&xlsx);
    assert!(matches!(
        get_cell_type(&workbook, 0, 0, 0),
        Some(CellType::Number)
    ));
}

#[test]
fn test_negative_scientific() {
    let xlsx = XlsxBuilder::new()
        .add_sheet("Sheet1")
        .add_cell("A1", -1.5e20, None)
        .build();

    let workbook = load_xlsx(&xlsx);
    assert!(matches!(
        get_cell_type(&workbook, 0, 0, 0),
        Some(CellType::Number)
    ));
}

#[test]
fn test_negative_scientific_small() {
    let xlsx = XlsxBuilder::new()
        .add_sheet("Sheet1")
        .add_cell("A1", -1.5e-20, None)
        .build();

    let workbook = load_xlsx(&xlsx);
    assert!(matches!(
        get_cell_type(&workbook, 0, 0, 0),
        Some(CellType::Number)
    ));
}

// ============================================================================
// Mixed Type Tests
// ============================================================================

#[test]
fn test_mixed_cell_types_single_row() {
    let xlsx = XlsxBuilder::new()
        .sheet(
            SheetBuilder::new("Sheet1")
                .cell("A1", CellValue::String("Text".to_string()), None)
                .cell("B1", CellValue::Number(123.0), None)
                .cell("C1", CellValue::Boolean(true), None)
                .cell("D1", CellValue::Error("#VALUE!".to_string()), None),
        )
        .build();

    let workbook = load_xlsx(&xlsx);

    assert!(matches!(
        get_cell_type(&workbook, 0, 0, 0),
        Some(CellType::String)
    ));
    assert!(matches!(
        get_cell_type(&workbook, 0, 0, 1),
        Some(CellType::Number)
    ));
    assert!(matches!(
        get_cell_type(&workbook, 0, 0, 2),
        Some(CellType::Boolean)
    ));
    assert!(matches!(
        get_cell_type(&workbook, 0, 0, 3),
        Some(CellType::Error)
    ));
}

#[test]
fn test_mixed_cell_types_multiple_sheets() {
    let xlsx = XlsxBuilder::new()
        .sheet(SheetBuilder::new("Strings").cell(
            "A1",
            CellValue::String("Hello".to_string()),
            None,
        ))
        .sheet(SheetBuilder::new("Numbers").cell("A1", CellValue::Number(42.0), None))
        .sheet(SheetBuilder::new("Booleans").cell("A1", CellValue::Boolean(true), None))
        .sheet(SheetBuilder::new("Errors").cell("A1", CellValue::Error("#N/A".to_string()), None))
        .build();

    let workbook = load_xlsx(&xlsx);

    assert!(matches!(
        get_cell_type(&workbook, 0, 0, 0),
        Some(CellType::String)
    ));
    assert!(matches!(
        get_cell_type(&workbook, 1, 0, 0),
        Some(CellType::Number)
    ));
    assert!(matches!(
        get_cell_type(&workbook, 2, 0, 0),
        Some(CellType::Boolean)
    ));
    assert!(matches!(
        get_cell_type(&workbook, 3, 0, 0),
        Some(CellType::Error)
    ));
}

// ============================================================================
// Edge Cases
// ============================================================================

#[test]
fn test_number_that_looks_like_date() {
    // Numbers like 44927 could be dates (days since 1900) but without date format
    // should be treated as numbers
    let xlsx = XlsxBuilder::new()
        .add_sheet("Sheet1")
        .add_cell("A1", 44927.0, None)
        .build();

    let workbook = load_xlsx(&xlsx);
    assert!(matches!(
        get_cell_type(&workbook, 0, 0, 0),
        Some(CellType::Number)
    ));
}

#[test]
fn test_cell_in_last_column() {
    // Excel column XFD (16384)
    let xlsx = XlsxBuilder::new()
        .add_sheet("Sheet1")
        .add_cell("XFD1", "Far right", None)
        .build();

    let workbook = load_xlsx(&xlsx);
    // Column XFD = 16384, so 0-indexed = 16383
    assert_eq!(
        get_cell_value(&workbook, 0, 0, 16383),
        Some("Far right".to_string())
    );
}

#[test]
fn test_cell_in_high_row() {
    let xlsx = XlsxBuilder::new()
        .add_sheet("Sheet1")
        .add_cell("A1000", "High row", None)
        .build();

    let workbook = load_xlsx(&xlsx);
    assert_eq!(
        get_cell_value(&workbook, 0, 999, 0),
        Some("High row".to_string())
    );
}

#[test]
fn test_multiple_cells_same_value_different_types() {
    // Same displayed value but different types
    let xlsx = XlsxBuilder::new()
        .sheet(
            SheetBuilder::new("Sheet1")
                .cell("A1", CellValue::String("1".to_string()), None)
                .cell("A2", CellValue::Number(1.0), None)
                .cell("A3", CellValue::Boolean(true), None),
        )
        .build();

    let workbook = load_xlsx(&xlsx);

    // All display "1" or "TRUE" but have different types
    assert!(matches!(
        get_cell_type(&workbook, 0, 0, 0),
        Some(CellType::String)
    ));
    assert!(matches!(
        get_cell_type(&workbook, 0, 1, 0),
        Some(CellType::Number)
    ));
    assert!(matches!(
        get_cell_type(&workbook, 0, 2, 0),
        Some(CellType::Boolean)
    ));
}
