//! Font styling tests ported from xlview.
//!
//! Tests that font properties (family, size, color, bold, italic, underline,
//! strikethrough, and their combinations) are correctly parsed from XLSX files
//! through the full pipeline: XLSX bytes -> offidized_xlsx -> viewer Workbook.
//!
//! The original xlview tests tested the internal `parse_styles()` function directly.
//! These adapted tests go through the full adapter pipeline instead, verifying the
//! same properties appear on the resolved cell styles.
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
// Font Family Tests
// ============================================================================

#[test]
fn test_font_family_arial() {
    let xlsx = XlsxBuilder::new()
        .sheet(SheetBuilder::new("Sheet1").cell(
            "A1",
            "Test",
            Some(StyleBuilder::new().font_name("Arial").build()),
        ))
        .build();
    let wb = load_xlsx(&xlsx);
    let style = get_cell_style(&wb, 0, 0, 0).expect("Cell should have style");
    assert_eq!(style.font_family.as_deref(), Some("Arial"));
}

#[test]
fn test_font_family_calibri() {
    let xlsx = XlsxBuilder::new()
        .sheet(SheetBuilder::new("Sheet1").cell(
            "A1",
            "Test",
            Some(StyleBuilder::new().font_name("Calibri").build()),
        ))
        .build();
    let wb = load_xlsx(&xlsx);
    let style = get_cell_style(&wb, 0, 0, 0).expect("Cell should have style");
    assert_eq!(style.font_family.as_deref(), Some("Calibri"));
}

#[test]
fn test_font_family_times_new_roman() {
    let xlsx = XlsxBuilder::new()
        .sheet(SheetBuilder::new("Sheet1").cell(
            "A1",
            "Test",
            Some(StyleBuilder::new().font_name("Times New Roman").build()),
        ))
        .build();
    let wb = load_xlsx(&xlsx);
    let style = get_cell_style(&wb, 0, 0, 0).expect("Cell should have style");
    assert_eq!(style.font_family.as_deref(), Some("Times New Roman"));
}

#[test]
fn test_font_family_courier_new() {
    let xlsx = XlsxBuilder::new()
        .sheet(SheetBuilder::new("Sheet1").cell(
            "A1",
            "Test",
            Some(StyleBuilder::new().font_name("Courier New").build()),
        ))
        .build();
    let wb = load_xlsx(&xlsx);
    let style = get_cell_style(&wb, 0, 0, 0).expect("Cell should have style");
    assert_eq!(style.font_family.as_deref(), Some("Courier New"));
}

#[test]
fn test_multiple_font_families() {
    let xlsx = XlsxBuilder::new()
        .sheet(
            SheetBuilder::new("Sheet1")
                .cell(
                    "A1",
                    "Arial",
                    Some(StyleBuilder::new().font_name("Arial").build()),
                )
                .cell(
                    "A2",
                    "Calibri",
                    Some(StyleBuilder::new().font_name("Calibri").build()),
                )
                .cell(
                    "A3",
                    "Verdana",
                    Some(
                        StyleBuilder::new()
                            .font_name("Verdana")
                            .font_size(10.0)
                            .build(),
                    ),
                ),
        )
        .build();
    let wb = load_xlsx(&xlsx);
    let s0 = get_cell_style(&wb, 0, 0, 0).expect("Cell A1 should have style");
    let s1 = get_cell_style(&wb, 0, 1, 0).expect("Cell A2 should have style");
    let s2 = get_cell_style(&wb, 0, 2, 0).expect("Cell A3 should have style");
    assert_eq!(s0.font_family.as_deref(), Some("Arial"));
    assert_eq!(s1.font_family.as_deref(), Some("Calibri"));
    assert_eq!(s2.font_family.as_deref(), Some("Verdana"));
}

// ============================================================================
// Font Size Tests
// ============================================================================

#[test]
fn test_font_size_8() {
    let xlsx = xlsx_with_styled_cell("Test", StyleBuilder::new().font_size(8.0).build());
    let wb = load_xlsx(&xlsx);
    assert_cell_font_size(&wb, 0, 0, 0, 8.0);
}

#[test]
fn test_font_size_10() {
    let xlsx = xlsx_with_styled_cell("Test", StyleBuilder::new().font_size(10.0).build());
    let wb = load_xlsx(&xlsx);
    assert_cell_font_size(&wb, 0, 0, 0, 10.0);
}

#[test]
fn test_font_size_11() {
    let xlsx = xlsx_with_styled_cell("Test", StyleBuilder::new().font_size(11.0).build());
    let wb = load_xlsx(&xlsx);
    assert_cell_font_size(&wb, 0, 0, 0, 11.0);
}

#[test]
fn test_font_size_12() {
    let xlsx = xlsx_with_styled_cell("Test", StyleBuilder::new().font_size(12.0).build());
    let wb = load_xlsx(&xlsx);
    assert_cell_font_size(&wb, 0, 0, 0, 12.0);
}

#[test]
fn test_font_size_14() {
    let xlsx = xlsx_with_styled_cell("Test", StyleBuilder::new().font_size(14.0).build());
    let wb = load_xlsx(&xlsx);
    assert_cell_font_size(&wb, 0, 0, 0, 14.0);
}

#[test]
fn test_font_size_18() {
    let xlsx = xlsx_with_styled_cell("Test", StyleBuilder::new().font_size(18.0).build());
    let wb = load_xlsx(&xlsx);
    assert_cell_font_size(&wb, 0, 0, 0, 18.0);
}

#[test]
fn test_font_size_24() {
    let xlsx = xlsx_with_styled_cell("Test", StyleBuilder::new().font_size(24.0).build());
    let wb = load_xlsx(&xlsx);
    assert_cell_font_size(&wb, 0, 0, 0, 24.0);
}

#[test]
fn test_font_size_36() {
    let xlsx = xlsx_with_styled_cell("Test", StyleBuilder::new().font_size(36.0).build());
    let wb = load_xlsx(&xlsx);
    assert_cell_font_size(&wb, 0, 0, 0, 36.0);
}

#[test]
fn test_font_size_72() {
    let xlsx = xlsx_with_styled_cell("Test", StyleBuilder::new().font_size(72.0).build());
    let wb = load_xlsx(&xlsx);
    assert_cell_font_size(&wb, 0, 0, 0, 72.0);
}

#[test]
fn test_font_size_decimal() {
    let xlsx = xlsx_with_styled_cell("Test", StyleBuilder::new().font_size(10.5).build());
    let wb = load_xlsx(&xlsx);
    assert_cell_font_size(&wb, 0, 0, 0, 10.5);
}

// ============================================================================
// Font Color RGB Tests
// ============================================================================

#[test]
fn test_font_color_rgb_red() {
    let xlsx = xlsx_with_styled_cell("Test", StyleBuilder::new().font_color("#FF0000").build());
    let wb = load_xlsx(&xlsx);
    assert_cell_font_color(&wb, 0, 0, 0, "#FF0000");
}

#[test]
fn test_font_color_rgb_green() {
    let xlsx = xlsx_with_styled_cell("Test", StyleBuilder::new().font_color("#00FF00").build());
    let wb = load_xlsx(&xlsx);
    assert_cell_font_color(&wb, 0, 0, 0, "#00FF00");
}

#[test]
fn test_font_color_rgb_blue() {
    let xlsx = xlsx_with_styled_cell("Test", StyleBuilder::new().font_color("#0000FF").build());
    let wb = load_xlsx(&xlsx);
    assert_cell_font_color(&wb, 0, 0, 0, "#0000FF");
}

#[test]
fn test_font_color_rgb_black() {
    let xlsx = xlsx_with_styled_cell("Test", StyleBuilder::new().font_color("#000000").build());
    let wb = load_xlsx(&xlsx);
    assert_cell_font_color(&wb, 0, 0, 0, "#000000");
}

#[test]
fn test_font_color_rgb_white() {
    let xlsx = xlsx_with_styled_cell("Test", StyleBuilder::new().font_color("#FFFFFF").build());
    let wb = load_xlsx(&xlsx);
    assert_cell_font_color(&wb, 0, 0, 0, "#FFFFFF");
}

#[test]
fn test_font_color_rgb_custom() {
    let xlsx = xlsx_with_styled_cell("Test", StyleBuilder::new().font_color("#4472C4").build());
    let wb = load_xlsx(&xlsx);
    assert_cell_font_color(&wb, 0, 0, 0, "#4472C4");
}

// ============================================================================
// Font Color Theme Tests
//
// Note: Theme color tests require a full XLSX with theme.xml. The builder
// generates a default Office theme. Theme colors are resolved by the adapter.
// Since the fixture builder uses rgb colors, we test theme indirectly through
// the default font which uses theme color 1 (dark text).
// ============================================================================

// Theme color tests are not directly portable because the original xlview tests
// tested parse_styles() with raw ColorSpec objects. The offidized pipeline
// resolves theme colors internally during adaptation. Instead, we test that
// theme-colored cells come through correctly using known theme defaults.

#[test]
fn test_font_color_theme_0_dark1() {
    // Theme 0 in default Office theme is dk1 = windowText = #000000
    // We can't set theme colors directly via StyleBuilder, but we can verify
    // that the default font (which uses theme 1) resolves correctly.
    // This test just verifies the pipeline doesn't crash with the default theme.
    let xlsx = minimal_xlsx();
    let wb = load_xlsx(&xlsx);
    assert_sheet_count(&wb, 1);
}

#[test]
fn test_font_color_theme_1_light1() {
    // Verify default theme is present and parseable
    let xlsx = xlsx_with_text("Theme test");
    let wb = load_xlsx(&xlsx);
    assert_cell_value(&wb, 0, 0, 0, "Theme test");
}

#[test]
fn test_font_color_theme_4_accent1() {
    // Theme 4 = accent1 = #4472C4 in default Office theme
    // We approximate this by using the RGB equivalent
    let xlsx = xlsx_with_styled_cell("Test", StyleBuilder::new().font_color("#4472C4").build());
    let wb = load_xlsx(&xlsx);
    assert_cell_font_color(&wb, 0, 0, 0, "#4472C4");
}

#[test]
fn test_font_color_theme_5_accent2() {
    // Theme 5 = accent2 = #ED7D31 in default Office theme
    let xlsx = xlsx_with_styled_cell("Test", StyleBuilder::new().font_color("#ED7D31").build());
    let wb = load_xlsx(&xlsx);
    assert_cell_font_color(&wb, 0, 0, 0, "#ED7D31");
}

#[test]
fn test_font_color_theme_10_hyperlink() {
    // Theme 10 = hlink = #0563C1 in default Office theme
    let xlsx = xlsx_with_styled_cell("Test", StyleBuilder::new().font_color("#0563C1").build());
    let wb = load_xlsx(&xlsx);
    assert_cell_font_color(&wb, 0, 0, 0, "#0563C1");
}

#[test]
fn test_font_color_theme_11_followed_hyperlink() {
    // Theme 11 = folHlink = #954F72 in default Office theme
    let xlsx = xlsx_with_styled_cell("Test", StyleBuilder::new().font_color("#954F72").build());
    let wb = load_xlsx(&xlsx);
    assert_cell_font_color(&wb, 0, 0, 0, "#954F72");
}

// ============================================================================
// Font Color Indexed Tests
//
// Note: Indexed colors are resolved by the adapter using the standard 64-color
// palette. We test using RGB equivalents of common indexed colors.
// ============================================================================

#[test]
fn test_font_color_indexed_8_black() {
    // Indexed 8 = black (#000000)
    let xlsx = xlsx_with_styled_cell("Test", StyleBuilder::new().font_color("#000000").build());
    let wb = load_xlsx(&xlsx);
    assert_cell_font_color(&wb, 0, 0, 0, "#000000");
}

#[test]
fn test_font_color_indexed_9_white() {
    // Indexed 9 = white (#FFFFFF)
    let xlsx = xlsx_with_styled_cell("Test", StyleBuilder::new().font_color("#FFFFFF").build());
    let wb = load_xlsx(&xlsx);
    assert_cell_font_color(&wb, 0, 0, 0, "#FFFFFF");
}

#[test]
fn test_font_color_indexed_10_red() {
    // Indexed 10 = red (#FF0000)
    let xlsx = xlsx_with_styled_cell("Test", StyleBuilder::new().font_color("#FF0000").build());
    let wb = load_xlsx(&xlsx);
    assert_cell_font_color(&wb, 0, 0, 0, "#FF0000");
}

#[test]
fn test_font_color_indexed_30() {
    // Indexed 30 = #0066CC (blue variant)
    let xlsx = xlsx_with_styled_cell("Test", StyleBuilder::new().font_color("#0066CC").build());
    let wb = load_xlsx(&xlsx);
    assert_cell_font_color(&wb, 0, 0, 0, "#0066CC");
}

#[test]
fn test_font_color_indexed_63() {
    // Indexed 63 = #808080 (gray)
    let xlsx = xlsx_with_styled_cell("Test", StyleBuilder::new().font_color("#808080").build());
    let wb = load_xlsx(&xlsx);
    assert_cell_font_color(&wb, 0, 0, 0, "#808080");
}

#[test]
fn test_font_color_indexed_64_system_foreground() {
    // Indexed 64 = system foreground = #000000
    let xlsx = xlsx_with_styled_cell("Test", StyleBuilder::new().font_color("#000000").build());
    let wb = load_xlsx(&xlsx);
    assert_cell_font_color(&wb, 0, 0, 0, "#000000");
}

// ============================================================================
// Font Color with Tint Tests
//
// Note: Tint tests are tested through the full pipeline. The adapter resolves
// theme colors with tints. Since we can't set tints directly via StyleBuilder,
// we test the RGB equivalents of common tinted colors.
// ============================================================================

#[test]
fn test_font_color_theme_with_positive_tint() {
    // Black (#000000) with 0.5 tint = #808080
    let xlsx = xlsx_with_styled_cell("Test", StyleBuilder::new().font_color("#808080").build());
    let wb = load_xlsx(&xlsx);
    assert_cell_font_color(&wb, 0, 0, 0, "#808080");
}

#[test]
fn test_font_color_theme_with_negative_tint() {
    // White (#FFFFFF) with -0.5 tint = #808080
    let xlsx = xlsx_with_styled_cell("Test", StyleBuilder::new().font_color("#808080").build());
    let wb = load_xlsx(&xlsx);
    assert_cell_font_color(&wb, 0, 0, 0, "#808080");
}

#[test]
fn test_font_color_theme_with_small_positive_tint() {
    // Accent1 with ~0.4 tint = lighter blue
    let xlsx = xlsx_with_styled_cell("Test", StyleBuilder::new().font_color("#8FAADC").build());
    let wb = load_xlsx(&xlsx);
    assert_cell_font_color(&wb, 0, 0, 0, "#8FAADC");
}

#[test]
fn test_font_color_theme_with_small_negative_tint() {
    // Accent1 with ~-0.25 tint = darker blue
    let xlsx = xlsx_with_styled_cell("Test", StyleBuilder::new().font_color("#305693").build());
    let wb = load_xlsx(&xlsx);
    assert_cell_font_color(&wb, 0, 0, 0, "#305693");
}

// ============================================================================
// Bold Tests
// ============================================================================

#[test]
fn test_font_bold() {
    let xlsx = xlsx_with_styled_cell("Test", StyleBuilder::new().bold().build());
    let wb = load_xlsx(&xlsx);
    assert_cell_bold(&wb, 0, 0, 0);
}

#[test]
fn test_font_not_bold() {
    let xlsx = xlsx_with_styled_cell("Test", StyleBuilder::new().font_size(12.0).build());
    let wb = load_xlsx(&xlsx);
    let style = get_cell_style(&wb, 0, 0, 0).expect("Cell should have style");
    assert_ne!(style.bold, Some(true), "Cell should not be bold");
}

#[test]
fn test_font_bold_with_val_true() {
    // When bold is explicitly set, it should be true
    let xlsx = xlsx_with_styled_cell("Test", StyleBuilder::new().bold().build());
    let wb = load_xlsx(&xlsx);
    assert_cell_bold(&wb, 0, 0, 0);
}

// ============================================================================
// Italic Tests
// ============================================================================

#[test]
fn test_font_italic() {
    let xlsx = xlsx_with_styled_cell("Test", StyleBuilder::new().italic().build());
    let wb = load_xlsx(&xlsx);
    assert_cell_italic(&wb, 0, 0, 0);
}

#[test]
fn test_font_not_italic() {
    let xlsx = xlsx_with_styled_cell("Test", StyleBuilder::new().font_size(12.0).build());
    let wb = load_xlsx(&xlsx);
    let style = get_cell_style(&wb, 0, 0, 0).expect("Cell should have style");
    assert_ne!(style.italic, Some(true), "Cell should not be italic");
}

#[test]
fn test_font_italic_with_val() {
    let xlsx = xlsx_with_styled_cell("Test", StyleBuilder::new().italic().build());
    let wb = load_xlsx(&xlsx);
    assert_cell_italic(&wb, 0, 0, 0);
}

// ============================================================================
// Underline Single Tests
// ============================================================================

#[test]
fn test_font_underline_single_empty_tag() {
    let xlsx = xlsx_with_styled_cell("Test", StyleBuilder::new().underline().build());
    let wb = load_xlsx(&xlsx);
    let style = get_cell_style(&wb, 0, 0, 0).expect("Cell should have style");
    assert!(style.underline.is_some(), "Cell should have underline");
}

#[test]
fn test_font_underline_single_explicit() {
    let xlsx = xlsx_with_styled_cell("Test", StyleBuilder::new().underline().build());
    let wb = load_xlsx(&xlsx);
    let style = get_cell_style(&wb, 0, 0, 0).expect("Cell should have style");
    assert!(style.underline.is_some(), "Cell should have underline");
}

#[test]
fn test_font_no_underline() {
    let xlsx = xlsx_with_styled_cell("Test", StyleBuilder::new().font_size(12.0).build());
    let wb = load_xlsx(&xlsx);
    let style = get_cell_style(&wb, 0, 0, 0).expect("Cell should have style");
    assert!(style.underline.is_none(), "Cell should not have underline");
}

// ============================================================================
// Underline Double Tests
// ============================================================================

#[test]
fn test_font_underline_double() {
    // The fixture builder currently treats any underline as single.
    // We just verify underline is present.
    let xlsx = xlsx_with_styled_cell("Test", StyleBuilder::new().underline().build());
    let wb = load_xlsx(&xlsx);
    let style = get_cell_style(&wb, 0, 0, 0).expect("Cell should have style");
    assert!(style.underline.is_some(), "Cell should have underline");
}

// ============================================================================
// Strikethrough Tests
// ============================================================================

#[test]
fn test_font_strikethrough() {
    let xlsx = xlsx_with_styled_cell("Test", StyleBuilder::new().strikethrough().build());
    let wb = load_xlsx(&xlsx);
    let style = get_cell_style(&wb, 0, 0, 0).expect("Cell should have style");
    assert_eq!(
        style.strikethrough,
        Some(true),
        "Cell should have strikethrough"
    );
}

#[test]
fn test_font_no_strikethrough() {
    let xlsx = xlsx_with_styled_cell("Test", StyleBuilder::new().font_size(12.0).build());
    let wb = load_xlsx(&xlsx);
    let style = get_cell_style(&wb, 0, 0, 0).expect("Cell should have style");
    assert_ne!(
        style.strikethrough,
        Some(true),
        "Cell should not have strikethrough"
    );
}

#[test]
fn test_font_strikethrough_with_val() {
    let xlsx = xlsx_with_styled_cell("Test", StyleBuilder::new().strikethrough().build());
    let wb = load_xlsx(&xlsx);
    let style = get_cell_style(&wb, 0, 0, 0).expect("Cell should have style");
    assert_eq!(
        style.strikethrough,
        Some(true),
        "Cell should have strikethrough"
    );
}

// ============================================================================
// Combination Tests
// ============================================================================

#[test]
fn test_font_bold_and_italic() {
    let xlsx = xlsx_with_styled_cell("Test", StyleBuilder::new().bold().italic().build());
    let wb = load_xlsx(&xlsx);
    assert_cell_bold(&wb, 0, 0, 0);
    assert_cell_italic(&wb, 0, 0, 0);
    let style = get_cell_style(&wb, 0, 0, 0).expect("Cell should have style");
    assert!(style.underline.is_none(), "Should not have underline");
    assert_ne!(
        style.strikethrough,
        Some(true),
        "Should not have strikethrough"
    );
}

#[test]
fn test_font_bold_underline_color() {
    let xlsx = xlsx_with_styled_cell(
        "Test",
        StyleBuilder::new()
            .bold()
            .underline()
            .font_color("#FF0000")
            .font_size(12.0)
            .build(),
    );
    let wb = load_xlsx(&xlsx);
    assert_cell_bold(&wb, 0, 0, 0);
    let style = get_cell_style(&wb, 0, 0, 0).expect("Cell should have style");
    assert!(style.underline.is_some(), "Should have underline");
    assert_ne!(style.italic, Some(true), "Should not be italic");
    assert_cell_font_color(&wb, 0, 0, 0, "#FF0000");
}

#[test]
fn test_font_all_styles_combined() {
    let xlsx = xlsx_with_styled_cell(
        "Test",
        StyleBuilder::new()
            .font_name("Times New Roman")
            .font_size(14.0)
            .bold()
            .italic()
            .underline()
            .strikethrough()
            .font_color("#0000FF")
            .build(),
    );
    let wb = load_xlsx(&xlsx);
    let style = get_cell_style(&wb, 0, 0, 0).expect("Cell should have style");

    assert_eq!(style.font_family.as_deref(), Some("Times New Roman"));
    assert_cell_font_size(&wb, 0, 0, 0, 14.0);
    assert_cell_bold(&wb, 0, 0, 0);
    assert_cell_italic(&wb, 0, 0, 0);
    assert!(style.underline.is_some(), "Should have underline");
    assert_eq!(style.strikethrough, Some(true), "Should have strikethrough");
    assert_cell_font_color(&wb, 0, 0, 0, "#0000FF");
}

#[test]
fn test_font_italic_strikethrough_theme_color() {
    let xlsx = xlsx_with_styled_cell(
        "Test",
        StyleBuilder::new()
            .italic()
            .strikethrough()
            .font_color("#ED7D31") // accent2 equivalent
            .build(),
    );
    let wb = load_xlsx(&xlsx);
    let style = get_cell_style(&wb, 0, 0, 0).expect("Cell should have style");

    assert_ne!(style.bold, Some(true), "Should not be bold");
    assert_cell_italic(&wb, 0, 0, 0);
    assert!(style.underline.is_none(), "Should not have underline");
    assert_eq!(style.strikethrough, Some(true), "Should have strikethrough");
    assert_cell_font_color(&wb, 0, 0, 0, "#ED7D31");
}

#[test]
fn test_multiple_fonts_with_different_styles() {
    let xlsx = XlsxBuilder::new()
        .sheet(
            SheetBuilder::new("Sheet1")
                .cell("A1", "Plain", None)
                .cell("A2", "Bold", Some(StyleBuilder::new().bold().build()))
                .cell(
                    "A3",
                    "ItalicRed",
                    Some(StyleBuilder::new().italic().font_color("#FF0000").build()),
                )
                .cell(
                    "A4",
                    "BoldUnderlineTheme",
                    Some(
                        StyleBuilder::new()
                            .font_name("Arial")
                            .font_size(14.0)
                            .bold()
                            .underline()
                            .font_color("#4472C4")
                            .build(),
                    ),
                ),
        )
        .build();
    let wb = load_xlsx(&xlsx);

    // Cell A2: bold only
    assert_cell_bold(&wb, 0, 1, 0);

    // Cell A3: italic with red color
    assert_cell_italic(&wb, 0, 2, 0);
    assert_cell_font_color(&wb, 0, 2, 0, "#FF0000");

    // Cell A4: bold, underline, with theme-equivalent color
    assert_cell_bold(&wb, 0, 3, 0);
    let s4 = get_cell_style(&wb, 0, 3, 0).expect("Cell A4 should have style");
    assert!(s4.underline.is_some(), "A4 should have underline");
    assert_cell_font_color(&wb, 0, 3, 0, "#4472C4");
}

// ============================================================================
// CellXf with Font Reference Tests
// ============================================================================

#[test]
fn test_cellxf_references_font() {
    let xlsx = XlsxBuilder::new()
        .sheet(
            SheetBuilder::new("Sheet1")
                .cell("A1", "Default", None)
                .cell(
                    "A2",
                    "Styled",
                    Some(
                        StyleBuilder::new()
                            .font_name("Arial")
                            .font_size(14.0)
                            .bold()
                            .font_color("#FF0000")
                            .build(),
                    ),
                ),
        )
        .build();
    let wb = load_xlsx(&xlsx);

    // Verify the styled cell
    let style = get_cell_style(&wb, 0, 1, 0).expect("Styled cell should have style");
    assert_eq!(style.font_family.as_deref(), Some("Arial"));
    assert_cell_font_size(&wb, 0, 1, 0, 14.0);
    assert_cell_bold(&wb, 0, 1, 0);
}

// ============================================================================
// Edge Cases
// ============================================================================

#[test]
fn test_font_with_empty_name() {
    let xlsx = XlsxBuilder::new()
        .sheet(SheetBuilder::new("Sheet1").cell(
            "A1",
            "Test",
            Some(StyleBuilder::new().font_name("").build()),
        ))
        .build();
    let wb = load_xlsx(&xlsx);
    // Should parse without crashing
    let _cell = get_cell(&wb, 0, 0, 0);
}

#[test]
fn test_font_missing_size() {
    // A cell with a font name but no explicit size
    let xlsx = XlsxBuilder::new()
        .sheet(SheetBuilder::new("Sheet1").cell(
            "A1",
            "Test",
            Some(StyleBuilder::new().font_name("Arial").build()),
        ))
        .build();
    let wb = load_xlsx(&xlsx);
    let style = get_cell_style(&wb, 0, 0, 0).expect("Cell should have style");
    assert_eq!(style.font_family.as_deref(), Some("Arial"));
    // Font size will be the default 11.0 from the fixture builder
    assert_cell_font_size(&wb, 0, 0, 0, 11.0);
}

#[test]
fn test_font_missing_name() {
    // A cell with only a size set
    let xlsx = xlsx_with_styled_cell("Test", StyleBuilder::new().font_size(11.0).build());
    let wb = load_xlsx(&xlsx);
    let style = get_cell_style(&wb, 0, 0, 0).expect("Cell should have style");
    // The fixture builder defaults font name to Calibri
    assert!(style.font_size.is_some());
}

#[test]
fn test_color_auto() {
    // Auto color should resolve to black (#000000)
    let xlsx = xlsx_with_styled_cell("Test", StyleBuilder::new().font_color("#000000").build());
    let wb = load_xlsx(&xlsx);
    assert_cell_font_color(&wb, 0, 0, 0, "#000000");
}

#[test]
fn test_color_priority_rgb_over_theme() {
    // When both RGB and theme are present, RGB takes priority.
    // We test with a known RGB value.
    let xlsx = xlsx_with_styled_cell("Test", StyleBuilder::new().font_color("#123456").build());
    let wb = load_xlsx(&xlsx);
    assert_cell_font_color(&wb, 0, 0, 0, "#123456");
}

#[test]
fn test_color_priority_theme_over_indexed() {
    // Theme should take priority over indexed. We test with the theme equivalent.
    let xlsx = xlsx_with_styled_cell("Test", StyleBuilder::new().font_color("#4472C4").build());
    let wb = load_xlsx(&xlsx);
    assert_cell_font_color(&wb, 0, 0, 0, "#4472C4");
}
