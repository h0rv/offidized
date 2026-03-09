//! Font family rendering tests ported from xlview.
//!
//! Integration tests that verify font family styles are correctly parsed from XLSX files
//! and available in the parsed cell data for rendering.
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

// ============================================================================
// Test 1: Cell with explicit font family (Arial) should use that font
// ============================================================================

#[test]
fn test_cell_with_arial_font() {
    let style = StyleBuilder::new().font_name("Arial").build();
    let xlsx = XlsxBuilder::new()
        .sheet(SheetBuilder::new("Sheet1").cell("A1", "Test", Some(style)))
        .build();

    let workbook = load_xlsx(&xlsx);
    assert!(
        !workbook.sheets.is_empty(),
        "Workbook should have at least one sheet"
    );
    assert!(
        !workbook.sheets[0].cells.is_empty(),
        "Sheet should have at least one cell"
    );

    let style = get_cell_style(&workbook, 0, 0, 0).expect("Cell should have a style");
    assert_eq!(
        style.font_family,
        Some("Arial".to_string()),
        "Cell should have Arial font family"
    );
}

#[test]
fn test_cell_with_times_new_roman_font() {
    let style = StyleBuilder::new().font_name("Times New Roman").build();
    let xlsx = XlsxBuilder::new()
        .sheet(SheetBuilder::new("Sheet1").cell("A1", "Serif Text", Some(style)))
        .build();

    let workbook = load_xlsx(&xlsx);
    let style = get_cell_style(&workbook, 0, 0, 0).expect("Cell should have style");
    assert_eq!(
        style.font_family,
        Some("Times New Roman".to_string()),
        "Cell should have Times New Roman font family"
    );
}

// ============================================================================
// Test 2: Cell without font family should use default font (Calibri)
// ============================================================================

#[test]
fn test_cell_without_font_family_uses_default() {
    let xlsx = XlsxBuilder::new()
        .sheet(SheetBuilder::new("Sheet1").cell("A1", "No Style", None))
        .build();

    let workbook = load_xlsx(&xlsx);
    let cell = get_cell(&workbook, 0, 0, 0);

    // Cell may have a default style applied or no style at all
    if let Some(cd) = cell {
        if let Some(ref style) = cd.cell.s {
            if style.font_family.is_some() {
                assert_eq!(
                    style.font_family,
                    Some("Calibri".to_string()),
                    "Default font family should be Calibri"
                );
            }
        }
    }
}

#[test]
fn test_cell_with_style_but_no_font_name() {
    let style = StyleBuilder::new().bold().build();
    let xlsx = XlsxBuilder::new()
        .sheet(SheetBuilder::new("Sheet1").cell("A1", "Bold Only", Some(style)))
        .build();

    let workbook = load_xlsx(&xlsx);
    let style = get_cell_style(&workbook, 0, 0, 0).expect("Cell should have style");
    assert_eq!(style.bold, Some(true), "Cell should be bold");
    // Font family should be Calibri (default from stylesheet)
    assert_eq!(
        style.font_family,
        Some("Calibri".to_string()),
        "Default font family should be Calibri"
    );
}

// ============================================================================
// Test 3: Cell with unknown font family should fall back gracefully
// ============================================================================

#[test]
fn test_cell_with_unknown_font_family() {
    let style = StyleBuilder::new()
        .font_name("NonExistentFont12345")
        .build();
    let xlsx = XlsxBuilder::new()
        .sheet(SheetBuilder::new("Sheet1").cell("A1", "Unknown Font", Some(style)))
        .build();

    let workbook = load_xlsx(&xlsx);
    let style = get_cell_style(&workbook, 0, 0, 0).expect("Cell should have style");
    assert_eq!(
        style.font_family,
        Some("NonExistentFont12345".to_string()),
        "Parser should preserve the specified font family even if unknown"
    );
}

#[test]
fn test_cell_with_empty_font_name() {
    let style = StyleBuilder::new().font_name("").build();
    let xlsx = XlsxBuilder::new()
        .sheet(SheetBuilder::new("Sheet1").cell("A1", "Empty Font", Some(style)))
        .build();

    let workbook = load_xlsx(&xlsx);
    // Should parse without error
    let _cell = get_cell(&workbook, 0, 0, 0);
}

// ============================================================================
// Test 4: Multiple cells with different font families
// ============================================================================

#[test]
fn test_multiple_cells_with_different_font_families() {
    let arial_style = StyleBuilder::new().font_name("Arial").build();
    let times_style = StyleBuilder::new().font_name("Times New Roman").build();
    let courier_style = StyleBuilder::new().font_name("Courier New").build();
    let verdana_style = StyleBuilder::new().font_name("Verdana").build();

    let xlsx = XlsxBuilder::new()
        .sheet(
            SheetBuilder::new("Sheet1")
                .cell("A1", "Arial Cell", Some(arial_style))
                .cell("A2", "Times Cell", Some(times_style))
                .cell("A3", "Courier Cell", Some(courier_style))
                .cell("A4", "Verdana Cell", Some(verdana_style)),
        )
        .build();

    let workbook = load_xlsx(&xlsx);
    let sheet = &workbook.sheets[0];
    assert_eq!(sheet.cells.len(), 4, "Should have 4 cells");

    let arial_style = get_cell_style(&workbook, 0, 0, 0).expect("Should find Arial cell style");
    let times_style = get_cell_style(&workbook, 0, 1, 0).expect("Should find Times cell style");
    let courier_style = get_cell_style(&workbook, 0, 2, 0).expect("Should find Courier cell style");
    let verdana_style = get_cell_style(&workbook, 0, 3, 0).expect("Should find Verdana cell style");

    assert_eq!(arial_style.font_family, Some("Arial".to_string()));
    assert_eq!(times_style.font_family, Some("Times New Roman".to_string()));
    assert_eq!(courier_style.font_family, Some("Courier New".to_string()));
    assert_eq!(verdana_style.font_family, Some("Verdana".to_string()));
}

#[test]
fn test_multiple_cells_same_font_different_styles() {
    let arial_bold = StyleBuilder::new().font_name("Arial").bold().build();
    let arial_italic = StyleBuilder::new().font_name("Arial").italic().build();
    let arial_both = StyleBuilder::new()
        .font_name("Arial")
        .bold()
        .italic()
        .build();

    let xlsx = XlsxBuilder::new()
        .sheet(
            SheetBuilder::new("Sheet1")
                .cell("A1", "Bold", Some(arial_bold))
                .cell("A2", "Italic", Some(arial_italic))
                .cell("A3", "Bold Italic", Some(arial_both)),
        )
        .build();

    let workbook = load_xlsx(&xlsx);

    let bold_style = get_cell_style(&workbook, 0, 0, 0).expect("Should find bold cell style");
    let italic_style = get_cell_style(&workbook, 0, 1, 0).expect("Should find italic cell style");
    let both_style =
        get_cell_style(&workbook, 0, 2, 0).expect("Should find bold+italic cell style");

    // All should have Arial font
    assert_eq!(bold_style.font_family, Some("Arial".to_string()));
    assert_eq!(italic_style.font_family, Some("Arial".to_string()));
    assert_eq!(both_style.font_family, Some("Arial".to_string()));

    // Check individual style properties
    assert_eq!(bold_style.bold, Some(true));
    assert_ne!(bold_style.italic, Some(true));

    assert_eq!(italic_style.italic, Some(true));
    assert_ne!(italic_style.bold, Some(true));

    assert_eq!(both_style.bold, Some(true));
    assert_eq!(both_style.italic, Some(true));
}

// ============================================================================
// Test 5: Font family combined with other style properties
// ============================================================================

#[test]
fn test_font_family_with_size_and_color() {
    let style = StyleBuilder::new()
        .font_name("Georgia")
        .font_size(14.0)
        .font_color("#FF0000")
        .build();

    let xlsx = XlsxBuilder::new()
        .sheet(SheetBuilder::new("Sheet1").cell("A1", "Styled", Some(style)))
        .build();

    let workbook = load_xlsx(&xlsx);
    let style = get_cell_style(&workbook, 0, 0, 0).expect("Cell should have style");

    assert_eq!(
        style.font_family,
        Some("Georgia".to_string()),
        "Font family should be Georgia"
    );
    assert_eq!(style.font_size, Some(14.0), "Font size should be 14");
    assert!(style.font_color.is_some(), "Font color should be set");
    let color = style.font_color.as_ref().unwrap();
    assert!(
        color.contains("FF0000") || color.contains("ff0000"),
        "Font color should be red, got: {}",
        color
    );
}
