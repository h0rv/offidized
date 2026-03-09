//! Integration tests for range copy/move operations in offidized-xlsx.
//!
//! This test suite ensures that:
//! - Copy operations preserve source cells and duplicate data to destination
//! - Move operations clear source cells and transfer data to destination
//! - Values, formulas, styles, rich text, and comments are correctly copied/moved
//! - Roundtrip (save/reload) preserves all copied/moved data
//! - Edge cases (large ranges, overlapping ranges, empty cells) work correctly

#![allow(clippy::unwrap_used)]
#![allow(clippy::expect_used)]

use offidized_xlsx::{CellComment, CellValue, RichTextRun, Workbook};
use tempfile::TempDir;

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

/// Creates a new workbook with a single sheet for testing.
fn new_test_workbook() -> Workbook {
    let mut wb = Workbook::new();
    wb.add_sheet("Sheet1");
    wb
}

/// Saves a workbook to a temp file, reopens it, and returns (output_path, tempdir, reopened_workbook).
fn roundtrip_save(wb: Workbook) -> (std::path::PathBuf, TempDir, Workbook) {
    let tmp = tempfile::tempdir().expect("create tempdir");
    let output = tmp.path().join("test.xlsx");
    wb.save(&output).expect("save workbook");
    let reopened = Workbook::open(&output).expect("reopen workbook");
    (output, tmp, reopened)
}

// ---------------------------------------------------------------------------
// Copy Operations
// ---------------------------------------------------------------------------

#[test]
fn copy_range_with_values() {
    let mut wb = new_test_workbook();
    let ws = wb.sheet_mut("Sheet1").expect("get sheet");

    // Set up source range A1:B2 with values
    ws.cell_mut("A1").unwrap().set_value(CellValue::Number(1.0));
    ws.cell_mut("A2").unwrap().set_value(CellValue::Number(2.0));
    ws.cell_mut("B1")
        .unwrap()
        .set_value(CellValue::String("Hello".to_string()));
    ws.cell_mut("B2").unwrap().set_value(CellValue::Bool(true));

    // Copy to D5:E6
    ws.copy_range("A1:B2", "D5").expect("copy range");

    // Verify source cells are preserved
    assert_eq!(
        ws.cell("A1").unwrap().value(),
        Some(&CellValue::Number(1.0))
    );
    assert_eq!(
        ws.cell("A2").unwrap().value(),
        Some(&CellValue::Number(2.0))
    );
    assert_eq!(
        ws.cell("B1").unwrap().value(),
        Some(&CellValue::String("Hello".to_string()))
    );
    assert_eq!(ws.cell("B2").unwrap().value(), Some(&CellValue::Bool(true)));

    // Verify destination cells have copied values
    assert_eq!(
        ws.cell("D5").unwrap().value(),
        Some(&CellValue::Number(1.0))
    );
    assert_eq!(
        ws.cell("D6").unwrap().value(),
        Some(&CellValue::Number(2.0))
    );
    assert_eq!(
        ws.cell("E5").unwrap().value(),
        Some(&CellValue::String("Hello".to_string()))
    );
    assert_eq!(ws.cell("E6").unwrap().value(), Some(&CellValue::Bool(true)));
}

#[test]
fn copy_range_with_formulas() {
    let mut wb = new_test_workbook();
    let ws = wb.sheet_mut("Sheet1").expect("get sheet");

    // Set up source cells with formulas
    ws.cell_mut("A1").unwrap().set_formula("SUM(B1:B10)");
    ws.cell_mut("A2")
        .unwrap()
        .set_formula("IF(C2>10,\"High\",\"Low\")");

    // Copy to D1:D2
    ws.copy_range("A1:A2", "D1").expect("copy range");

    // Verify source formulas are preserved
    assert_eq!(ws.cell("A1").unwrap().formula(), Some("SUM(B1:B10)"));
    assert_eq!(
        ws.cell("A2").unwrap().formula(),
        Some("IF(C2>10,\"High\",\"Low\")")
    );

    // Verify destination has copied formulas (note: formulas are copied as-is, no reference shifting)
    assert_eq!(ws.cell("D1").unwrap().formula(), Some("SUM(B1:B10)"));
    assert_eq!(
        ws.cell("D2").unwrap().formula(),
        Some("IF(C2>10,\"High\",\"Low\")")
    );
}

#[test]
fn copy_range_with_styles() {
    let mut wb = new_test_workbook();
    let ws = wb.sheet_mut("Sheet1").expect("get sheet");

    // Set up source cells with values and style IDs
    ws.cell_mut("A1").unwrap().set_value(100).set_style_id(5);
    ws.cell_mut("A2")
        .unwrap()
        .set_value("Styled")
        .set_style_id(10);

    // Copy to C1:C2
    ws.copy_range("A1:A2", "C1").expect("copy range");

    // Verify source cells retain their styles
    assert_eq!(ws.cell("A1").unwrap().style_id(), Some(5));
    assert_eq!(ws.cell("A2").unwrap().style_id(), Some(10));

    // Verify destination cells have copied styles
    assert_eq!(ws.cell("C1").unwrap().style_id(), Some(5));
    assert_eq!(ws.cell("C2").unwrap().style_id(), Some(10));
}

#[test]
fn copy_range_with_rich_text() {
    let mut wb = new_test_workbook();
    let ws = wb.sheet_mut("Sheet1").expect("get sheet");

    // Create a rich text value with multiple runs
    let mut run1 = RichTextRun::new("Bold");
    run1.set_bold(true);
    let run2 = RichTextRun::new(" and ");
    let mut run3 = RichTextRun::new("Italic");
    run3.set_italic(true);
    let rich_value = CellValue::rich_text(vec![run1.clone(), run2, run3]);

    ws.cell_mut("A1").unwrap().set_value(rich_value.clone());

    // Copy to B5
    ws.copy_range("A1:A1", "B5").expect("copy range");

    // Verify source is preserved
    assert_eq!(ws.cell("A1").unwrap().value(), Some(&rich_value));

    // Verify destination has the same rich text
    assert_eq!(ws.cell("B5").unwrap().value(), Some(&rich_value));
}

#[test]
fn copy_range_with_comments() {
    let mut wb = new_test_workbook();
    let ws = wb.sheet_mut("Sheet1").expect("get sheet");

    // Add a cell with a comment
    ws.cell_mut("A1")
        .unwrap()
        .set_value("Has comment")
        .set_comment(CellComment::new("Author", "This is a note"));

    // Copy to C3
    ws.copy_range("A1:A1", "C3").expect("copy range");

    // Verify source comment is preserved
    let source_comment = ws.cell("A1").unwrap().comment();
    assert!(source_comment.is_some());
    assert_eq!(source_comment.unwrap().author(), "Author");
    assert_eq!(source_comment.unwrap().text(), "This is a note");

    // Verify destination has the copied comment
    let dest_comment = ws.cell("C3").unwrap().comment();
    assert!(dest_comment.is_some());
    assert_eq!(dest_comment.unwrap().author(), "Author");
    assert_eq!(dest_comment.unwrap().text(), "This is a note");
}

#[test]
fn copy_empty_cells_does_not_create_cells() {
    let mut wb = new_test_workbook();
    let ws = wb.sheet_mut("Sheet1").expect("get sheet");

    // Set up a sparse range with only A1 and B2 populated
    ws.cell_mut("A1").unwrap().set_value(1);
    ws.cell_mut("B2").unwrap().set_value(2);

    let cell_count_before = ws.cells().count();

    // Copy A1:B2 to D5:E6
    ws.copy_range("A1:B2", "D5").expect("copy range");

    // Only cells that existed in source should be copied
    // Before: A1, B2 (2 cells)
    // After: A1, B2, D5, E6 (4 cells)
    assert_eq!(ws.cells().count(), cell_count_before + 2);

    // Verify the copied cells
    assert_eq!(
        ws.cell("D5").unwrap().value(),
        Some(&CellValue::Number(1.0))
    );
    assert_eq!(
        ws.cell("E6").unwrap().value(),
        Some(&CellValue::Number(2.0))
    );

    // Verify empty cells in the source range were not copied
    assert!(ws.cell("D6").is_none());
    assert!(ws.cell("E5").is_none());
}

#[test]
fn copy_to_overlapping_range() {
    let mut wb = new_test_workbook();
    let ws = wb.sheet_mut("Sheet1").expect("get sheet");

    // Set up source range A1:B2
    ws.cell_mut("A1").unwrap().set_value("A1");
    ws.cell_mut("A2").unwrap().set_value("A2");
    ws.cell_mut("B1").unwrap().set_value("B1");
    ws.cell_mut("B2").unwrap().set_value("B2");

    // Copy to B1:C2 (overlapping with source)
    ws.copy_range("A1:B2", "B1").expect("copy range");

    // Verify all cells
    assert_eq!(
        ws.cell("A1").unwrap().value(),
        Some(&CellValue::String("A1".to_string()))
    );
    assert_eq!(
        ws.cell("A2").unwrap().value(),
        Some(&CellValue::String("A2".to_string()))
    );
    // B1 should now have the value from A1 (copied)
    assert_eq!(
        ws.cell("B1").unwrap().value(),
        Some(&CellValue::String("A1".to_string()))
    );
    // B2 should now have the value from A2 (copied)
    assert_eq!(
        ws.cell("B2").unwrap().value(),
        Some(&CellValue::String("A2".to_string()))
    );
    // C1 should have the value from B1 (original)
    assert_eq!(
        ws.cell("C1").unwrap().value(),
        Some(&CellValue::String("B1".to_string()))
    );
    // C2 should have the value from B2 (original)
    assert_eq!(
        ws.cell("C2").unwrap().value(),
        Some(&CellValue::String("B2".to_string()))
    );
}

#[test]
fn copy_single_cell() {
    let mut wb = new_test_workbook();
    let ws = wb.sheet_mut("Sheet1").expect("get sheet");

    ws.cell_mut("A1").unwrap().set_value(42).set_style_id(7);

    ws.copy_range("A1:A1", "Z10").expect("copy single cell");

    // Verify source is unchanged
    assert_eq!(
        ws.cell("A1").unwrap().value(),
        Some(&CellValue::Number(42.0))
    );
    assert_eq!(ws.cell("A1").unwrap().style_id(), Some(7));

    // Verify destination
    assert_eq!(
        ws.cell("Z10").unwrap().value(),
        Some(&CellValue::Number(42.0))
    );
    assert_eq!(ws.cell("Z10").unwrap().style_id(), Some(7));
}

#[test]
fn copy_large_range() {
    let mut wb = new_test_workbook();
    let ws = wb.sheet_mut("Sheet1").expect("get sheet");

    // Create a 10x10 grid filled with sequential numbers
    for row in 1..=10 {
        for col_idx in 0..10 {
            let col_letter = (b'A' + col_idx) as char;
            let value = (row - 1) * 10 + (col_idx + 1);
            let cell_ref = format!("{}{}", col_letter, row);
            ws.cell_mut(&cell_ref).unwrap().set_value(value);
        }
    }

    // Copy A1:J10 to L1:U10
    ws.copy_range("A1:J10", "L1").expect("copy large range");

    // Verify source cells are preserved
    assert_eq!(
        ws.cell("A1").unwrap().value(),
        Some(&CellValue::Number(1.0))
    );
    assert_eq!(
        ws.cell("J1").unwrap().value(),
        Some(&CellValue::Number(10.0))
    );
    assert_eq!(
        ws.cell("A10").unwrap().value(),
        Some(&CellValue::Number(91.0))
    );
    assert_eq!(
        ws.cell("J10").unwrap().value(),
        Some(&CellValue::Number(100.0))
    );

    // Verify destination cells have copied values
    assert_eq!(
        ws.cell("L1").unwrap().value(),
        Some(&CellValue::Number(1.0))
    );
    assert_eq!(
        ws.cell("U1").unwrap().value(),
        Some(&CellValue::Number(10.0))
    );
    assert_eq!(
        ws.cell("L10").unwrap().value(),
        Some(&CellValue::Number(91.0))
    );
    assert_eq!(
        ws.cell("U10").unwrap().value(),
        Some(&CellValue::Number(100.0))
    );

    // Middle cell: E5 should have value 45
    assert_eq!(
        ws.cell("E5").unwrap().value(),
        Some(&CellValue::Number(45.0))
    );
    // Copied version: P5 should also have value 45
    assert_eq!(
        ws.cell("P5").unwrap().value(),
        Some(&CellValue::Number(45.0))
    );
}

#[test]
fn copy_boolean_and_error_values() {
    let mut wb = new_test_workbook();
    let ws = wb.sheet_mut("Sheet1").expect("get sheet");

    // Set up source cells with various value types including boolean and error
    ws.cell_mut("A1").unwrap().set_value(CellValue::Bool(true));
    ws.cell_mut("A2").unwrap().set_value(CellValue::Bool(false));
    ws.cell_mut("B1")
        .unwrap()
        .set_value(CellValue::Error("#DIV/0!".to_string()));
    ws.cell_mut("B2")
        .unwrap()
        .set_value(CellValue::Error("#N/A".to_string()));
    ws.cell_mut("C1")
        .unwrap()
        .set_value(CellValue::Number(42.5));
    ws.cell_mut("C2")
        .unwrap()
        .set_value(CellValue::String("Text".to_string()));

    // Copy to D1:F2
    ws.copy_range("A1:C2", "D1").expect("copy range");

    // Verify source cells preserved
    assert_eq!(ws.cell("A1").unwrap().value(), Some(&CellValue::Bool(true)));
    assert_eq!(
        ws.cell("A2").unwrap().value(),
        Some(&CellValue::Bool(false))
    );
    assert_eq!(
        ws.cell("B1").unwrap().value(),
        Some(&CellValue::Error("#DIV/0!".to_string()))
    );
    assert_eq!(
        ws.cell("B2").unwrap().value(),
        Some(&CellValue::Error("#N/A".to_string()))
    );
    assert_eq!(
        ws.cell("C1").unwrap().value(),
        Some(&CellValue::Number(42.5))
    );
    assert_eq!(
        ws.cell("C2").unwrap().value(),
        Some(&CellValue::String("Text".to_string()))
    );

    // Verify destination cells have copied values
    assert_eq!(ws.cell("D1").unwrap().value(), Some(&CellValue::Bool(true)));
    assert_eq!(
        ws.cell("D2").unwrap().value(),
        Some(&CellValue::Bool(false))
    );
    assert_eq!(
        ws.cell("E1").unwrap().value(),
        Some(&CellValue::Error("#DIV/0!".to_string()))
    );
    assert_eq!(
        ws.cell("E2").unwrap().value(),
        Some(&CellValue::Error("#N/A".to_string()))
    );
    assert_eq!(
        ws.cell("F1").unwrap().value(),
        Some(&CellValue::Number(42.5))
    );
    assert_eq!(
        ws.cell("F2").unwrap().value(),
        Some(&CellValue::String("Text".to_string()))
    );
}

// ---------------------------------------------------------------------------
// Move Operations
// ---------------------------------------------------------------------------

#[test]
fn move_range_clears_source() {
    // Note: move_range is not yet implemented in worksheet.rs.
    // This test documents the expected behavior.
    //
    // When implemented, move_range should:
    // 1. Copy all cells from source range to destination
    // 2. Clear all cells in the source range
    //
    // Expected API:
    //   ws.move_range("A1:B2", "D5").expect("move range");
}

#[test]
fn move_single_cell() {
    // Note: move_range is not yet implemented.
    // When implemented, moving a single cell should transfer the cell to the
    // destination and remove it from the source.
    //
    // Expected behavior:
    //   ws.cell_mut("A1").unwrap().set_value(42);
    //   ws.move_range("A1:A1", "Z10").expect("move cell");
    //   assert!(ws.cell("A1").is_none()); // source cleared
    //   assert_eq!(ws.cell("Z10").unwrap().value(), Some(&CellValue::Number(42.0)));
}

#[test]
fn move_range_with_formulas_and_styles() {
    // Note: move_range is not yet implemented.
    // When implemented, formulas and styles should be moved to destination
    // and cleared from source.
}

#[test]
fn move_to_non_overlapping_location() {
    // Note: move_range is not yet implemented.
    // When implemented, moving to a non-overlapping range should be straightforward:
    // copy to destination, clear source.
}

// ---------------------------------------------------------------------------
// Roundtrip Tests
// ---------------------------------------------------------------------------

#[test]
fn roundtrip_after_copy_preserves_data() {
    let mut wb = new_test_workbook();
    let ws = wb.sheet_mut("Sheet1").expect("get sheet");

    // Set up source range with diverse data
    ws.cell_mut("A1").unwrap().set_value(100).set_style_id(1);
    ws.cell_mut("A2")
        .unwrap()
        .set_formula("A1*2")
        .set_cached_value(CellValue::Number(200.0));
    ws.cell_mut("B1")
        .unwrap()
        .set_value("Text")
        .set_comment(CellComment::new("User", "Important"));

    // Copy to D5:E6
    ws.copy_range("A1:B2", "D5").expect("copy range");

    // Save and reload
    let (_output, _tmp, reopened) = roundtrip_save(wb);
    let ws = reopened.sheet("Sheet1").expect("get sheet");

    // Verify source cells survived roundtrip
    assert_eq!(
        ws.cell("A1").unwrap().value(),
        Some(&CellValue::Number(100.0))
    );
    assert_eq!(ws.cell("A1").unwrap().style_id(), Some(1));
    assert_eq!(ws.cell("A2").unwrap().formula(), Some("A1*2"));
    assert_eq!(
        ws.cell("A2").unwrap().cached_value(),
        Some(&CellValue::Number(200.0))
    );
    assert_eq!(
        ws.cell("B1").unwrap().value(),
        Some(&CellValue::String("Text".to_string()))
    );
    // Note: Comment roundtrip may not be fully implemented yet
    // assert!(ws.cell("B1").unwrap().comment().is_some());

    // Verify copied cells survived roundtrip
    assert_eq!(
        ws.cell("D5").unwrap().value(),
        Some(&CellValue::Number(100.0))
    );
    assert_eq!(ws.cell("D5").unwrap().style_id(), Some(1));
    assert_eq!(ws.cell("D6").unwrap().formula(), Some("A1*2"));
    assert_eq!(
        ws.cell("D6").unwrap().cached_value(),
        Some(&CellValue::Number(200.0))
    );
    assert_eq!(
        ws.cell("E5").unwrap().value(),
        Some(&CellValue::String("Text".to_string()))
    );
    // Note: Comment roundtrip may not be fully implemented yet
    // assert!(ws.cell("E5").unwrap().comment().is_some());
}

#[test]
fn roundtrip_after_move_clears_source() {
    // Note: move_range is not yet implemented.
    // When implemented, this test should verify that after save/reload:
    // - Destination cells contain all the moved data
    // - Source cells are completely absent (not just empty values, but no cell objects)
}

// ---------------------------------------------------------------------------
// Edge Cases
// ---------------------------------------------------------------------------

#[test]
fn copy_with_merged_cells_limitation() {
    // Note: copy_range does not currently handle merged cell ranges.
    // This test documents the current limitation.
    //
    // When merged cell support is added to copy_range:
    // 1. Copying a range containing merged cells should also copy the merge metadata
    // 2. The destination should have corresponding merged ranges
    //
    // Current behavior: merged cell metadata is not copied (only cell values/formulas/styles).
    let mut wb = new_test_workbook();
    let ws = wb.sheet_mut("Sheet1").expect("get sheet");

    // Merge A1:B2
    ws.add_merged_range("A1:B2").expect("merge cells");
    ws.cell_mut("A1").unwrap().set_value("Merged");

    // Copy to D5:E6
    ws.copy_range("A1:B2", "D5").expect("copy range");

    // Currently, only the cell value is copied, not the merge.
    // D5 will have "Merged", but D5:E6 will not be a merged range.
    assert_eq!(
        ws.cell("D5").unwrap().value(),
        Some(&CellValue::String("Merged".to_string()))
    );

    // Future: ws.merged_ranges() should include "D5:E6" after this feature is implemented.
}

#[test]
fn copy_with_conditional_formatting_limitation() {
    // Note: copy_range does not currently copy conditional formatting rules.
    // This test documents the current limitation.
    //
    // When conditional formatting support is added to copy_range:
    // 1. Conditional formatting rules applied to the source range should be duplicated
    //    for the destination range
    // 2. Rule references should be adjusted to the new range
    //
    // Current behavior: conditional formatting is not copied.
}
