//! Integration tests for frozen panes rendering.
//!
//! Tests verify that sheets with frozen rows/columns have those values
//! correctly parsed from XLSX files built with the fixtures module.
//!
//! Frozen panes in Excel:
//! - frozen_rows: number of rows frozen at top
//! - frozen_cols: number of columns frozen at left
//! - Frozen area stays visible while rest scrolls
//!
//! Ported from xlview's frozen_panes_rendering_tests.rs.
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
use common::load_xlsx;
use fixtures::*;

/// Test: Sheet with 1 frozen row
///
/// Verifies that a sheet with a single row frozen at the top
/// is correctly parsed with frozen_rows=1 and frozen_cols=0.
#[test]
fn test_sheet_with_frozen_row() {
    let xlsx = XlsxBuilder::new()
        .sheet(
            SheetBuilder::new("Sheet1")
                .freeze_panes(1, 0) // 1 row frozen, 0 columns
                .cell("A1", CellValue::String("Header".into()), None)
                .cell("A2", CellValue::String("Data Row 1".into()), None)
                .cell("A3", CellValue::String("Data Row 2".into()), None),
        )
        .build();

    let workbook = load_xlsx(&xlsx);

    assert_eq!(workbook.sheets.len(), 1);
    assert_eq!(workbook.sheets[0].frozen_rows, 1);
    assert_eq!(workbook.sheets[0].frozen_cols, 0);
}

/// Test: Sheet with 2 frozen columns
///
/// Verifies that a sheet with two columns frozen at the left
/// is correctly parsed with frozen_rows=0 and frozen_cols=2.
#[test]
fn test_sheet_with_frozen_columns() {
    let xlsx = XlsxBuilder::new()
        .sheet(
            SheetBuilder::new("Sheet1")
                .freeze_panes(0, 2) // 0 rows frozen, 2 columns
                .cell("A1", CellValue::String("ID".into()), None)
                .cell("B1", CellValue::String("Name".into()), None)
                .cell("C1", CellValue::String("Value".into()), None)
                .cell("D1", CellValue::String("Description".into()), None),
        )
        .build();

    let workbook = load_xlsx(&xlsx);

    assert_eq!(workbook.sheets.len(), 1);
    assert_eq!(workbook.sheets[0].frozen_rows, 0);
    assert_eq!(workbook.sheets[0].frozen_cols, 2);
}

/// Test: Sheet with both frozen rows and columns
///
/// Verifies that a sheet with both rows and columns frozen
/// (e.g., 2 rows, 1 column) is correctly parsed.
#[test]
fn test_sheet_with_frozen_rows_and_columns() {
    let xlsx = XlsxBuilder::new()
        .sheet(
            SheetBuilder::new("DataSheet")
                .freeze_panes(2, 1) // 2 rows frozen, 1 column
                .cell("A1", CellValue::String("Row Labels".into()), None)
                .cell("B1", CellValue::String("Jan".into()), None)
                .cell("C1", CellValue::String("Feb".into()), None)
                .cell("A2", CellValue::String("Category".into()), None)
                .cell("B2", CellValue::Number(100.0), None)
                .cell("C2", CellValue::Number(150.0), None)
                .cell("A3", CellValue::String("Product A".into()), None)
                .cell("B3", CellValue::Number(50.0), None)
                .cell("C3", CellValue::Number(75.0), None),
        )
        .build();

    let workbook = load_xlsx(&xlsx);

    assert_eq!(workbook.sheets.len(), 1);
    assert_eq!(workbook.sheets[0].frozen_rows, 2);
    assert_eq!(workbook.sheets[0].frozen_cols, 1);
}

/// Test: Sheet with no frozen panes (defaults to 0)
///
/// Verifies that a sheet without any frozen panes has
/// frozen_rows=0 and frozen_cols=0 by default.
#[test]
fn test_sheet_with_no_frozen_panes() {
    let xlsx = XlsxBuilder::new()
        .sheet(
            SheetBuilder::new("Sheet1")
                // No freeze_panes call - should default to no freezing
                .cell("A1", CellValue::String("Cell A1".into()), None)
                .cell("B1", CellValue::String("Cell B1".into()), None)
                .cell("A2", CellValue::String("Cell A2".into()), None)
                .cell("B2", CellValue::String("Cell B2".into()), None),
        )
        .build();

    let workbook = load_xlsx(&xlsx);

    assert_eq!(workbook.sheets.len(), 1);
    assert_eq!(workbook.sheets[0].frozen_rows, 0);
    assert_eq!(workbook.sheets[0].frozen_cols, 0);
}

/// Test: Sheet with large frozen area (5 rows, 3 columns)
///
/// Verifies that a sheet with a larger frozen area
/// (e.g., 5 rows and 3 columns) is correctly parsed.
#[test]
fn test_sheet_with_large_frozen_area() {
    let xlsx = XlsxBuilder::new()
        .sheet(
            SheetBuilder::new("LargeFreeze")
                .freeze_panes(5, 3) // 5 rows frozen, 3 columns
                // Header rows
                .cell("A1", CellValue::String("Level 1".into()), None)
                .cell("A2", CellValue::String("Level 2".into()), None)
                .cell("A3", CellValue::String("Level 3".into()), None)
                .cell("A4", CellValue::String("Level 4".into()), None)
                .cell("A5", CellValue::String("Level 5".into()), None)
                // Column headers
                .cell("B1", CellValue::String("Col B".into()), None)
                .cell("C1", CellValue::String("Col C".into()), None)
                .cell("D1", CellValue::String("Col D".into()), None)
                // Data starts at D6 (first non-frozen cell in data area)
                .cell("D6", CellValue::Number(999.0), None),
        )
        .build();

    let workbook = load_xlsx(&xlsx);

    assert_eq!(workbook.sheets.len(), 1);
    assert_eq!(workbook.sheets[0].frozen_rows, 5);
    assert_eq!(workbook.sheets[0].frozen_cols, 3);
}
