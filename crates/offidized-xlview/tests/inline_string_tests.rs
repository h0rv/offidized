//! Tests for inline string parsing in XLSX files
//!
//! Inline strings use `t="inlineStr"` with `<is><t>text</t></is>` structure
//! instead of shared string references.
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

// =============================================================================
// Basic Inline String Tests
// =============================================================================

#[test]
fn test_single_inline_string() {
    let xlsx = XlsxBuilder::new()
        .sheet(SheetBuilder::new("Sheet1").cell(
            "A1",
            CellValue::InlineString("Hello World".into()),
            None,
        ))
        .build();
    let wb = load_xlsx(&xlsx);

    assert_sheet_count(&wb, 1);
    let cd = get_cell(&wb, 0, 0, 0);
    assert!(cd.is_some(), "Should find cell A1");
    assert_cell_value(&wb, 0, 0, 0, "Hello World");
}

#[test]
fn test_multiple_inline_strings() {
    let xlsx = XlsxBuilder::new()
        .sheet(
            SheetBuilder::new("Sheet1")
                .cell("A1", CellValue::InlineString("First".into()), None)
                .cell("A2", CellValue::InlineString("Second".into()), None)
                .cell("A3", CellValue::InlineString("Third".into()), None)
                .cell("B1", CellValue::InlineString("Column B".into()), None),
        )
        .build();
    let wb = load_xlsx(&xlsx);

    assert_cell_value(&wb, 0, 0, 0, "First");
    assert_cell_value(&wb, 0, 1, 0, "Second");
    assert_cell_value(&wb, 0, 2, 0, "Third");
    assert_cell_value(&wb, 0, 0, 1, "Column B");
}

#[test]
fn test_inline_string_with_special_characters() {
    let xlsx = XlsxBuilder::new()
        .sheet(SheetBuilder::new("Sheet1").cell(
            "A1",
            CellValue::InlineString("Test & Value".into()),
            None,
        ))
        .build();
    let wb = load_xlsx(&xlsx);

    // The &amp; should be unescaped to &
    assert_cell_value(&wb, 0, 0, 0, "Test & Value");
}

#[test]
fn test_inline_string_unicode() {
    let xlsx = XlsxBuilder::new()
        .sheet(
            SheetBuilder::new("Sheet1")
                .cell("A1", CellValue::InlineString("日本語".into()), None)
                .cell("A2", CellValue::InlineString("中文".into()), None)
                .cell("A3", CellValue::InlineString("한국어".into()), None)
                .cell("A4", CellValue::InlineString("Ελληνικά".into()), None),
        )
        .build();
    let wb = load_xlsx(&xlsx);

    assert_cell_value(&wb, 0, 0, 0, "日本語");
    assert_cell_value(&wb, 0, 1, 0, "中文");
    assert_cell_value(&wb, 0, 2, 0, "한국어");
    assert_cell_value(&wb, 0, 3, 0, "Ελληνικά");
}

// =============================================================================
// Mixed Content Tests
// =============================================================================

#[test]
fn test_inline_strings_with_regular_values() {
    let xlsx = XlsxBuilder::new()
        .sheet(
            SheetBuilder::new("Sheet1")
                .cell("A1", CellValue::InlineString("Inline text".into()), None)
                .cell("B1", 42.0, None)
                .cell("C1", "Shared string", None),
        )
        .build();
    let wb = load_xlsx(&xlsx);

    assert_cell_value(&wb, 0, 0, 0, "Inline text");
    assert_cell_value(&wb, 0, 0, 1, "42");
    assert_cell_value(&wb, 0, 0, 2, "Shared string");
}

#[test]
fn test_inline_string_empty() {
    let xlsx = XlsxBuilder::new()
        .sheet(SheetBuilder::new("Sheet1").cell("A1", CellValue::InlineString("".into()), None))
        .build();
    let wb = load_xlsx(&xlsx);

    // Empty cell might not be included, or might have empty value
    let cd = get_cell(&wb, 0, 0, 0);
    if let Some(cd) = cd {
        assert!(
            cd.cell.v.is_none() || cd.cell.v.as_deref() == Some(""),
            "Expected empty value, got: {:?}",
            cd.cell.v
        );
    }
}

#[test]
fn test_inline_string_with_styled_cell() {
    let style = StyleBuilder::new().bold().font_size(14.0).build();
    let xlsx = XlsxBuilder::new()
        .sheet(SheetBuilder::new("Sheet1").cell(
            "A1",
            CellValue::InlineString("Styled inline".into()),
            Some(style),
        ))
        .build();
    let wb = load_xlsx(&xlsx);

    assert_cell_value(&wb, 0, 0, 0, "Styled inline");
    assert_cell_bold(&wb, 0, 0, 0);
    assert_cell_font_size(&wb, 0, 0, 0, 14.0);
}

#[test]
fn test_inline_string_multisheet() {
    let xlsx = XlsxBuilder::new()
        .sheet(SheetBuilder::new("Sheet1").cell(
            "A1",
            CellValue::InlineString("Sheet 1 data".into()),
            None,
        ))
        .sheet(SheetBuilder::new("Sheet2").cell(
            "A1",
            CellValue::InlineString("Sheet 2 data".into()),
            None,
        ))
        .build();
    let wb = load_xlsx(&xlsx);

    assert_sheet_count(&wb, 2);
    assert_cell_value(&wb, 0, 0, 0, "Sheet 1 data");
    assert_cell_value(&wb, 1, 0, 0, "Sheet 2 data");
}
