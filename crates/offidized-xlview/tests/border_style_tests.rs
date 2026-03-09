//! Comprehensive tests for all 13 border styles using XlsxBuilder fixtures.
//!
//! Tests all ECMA-376 border styles can be created via StyleBuilder and parsed correctly.
//! Verifies border styles are correctly set in the parsed Style struct.
//! Ported from xlview border_style_tests.rs.
#![allow(
    clippy::unwrap_used,
    clippy::expect_used,
    clippy::indexing_slicing,
    clippy::float_cmp,
    clippy::panic,
    clippy::approx_constant,
    clippy::cast_possible_truncation,
    clippy::absurd_extreme_comparisons,
    clippy::cast_lossless,
    clippy::expect_fun_call
)]

mod common;
mod fixtures;

use common::*;
use fixtures::ALL_BORDER_STYLES;
use offidized_xlview::types::style::BorderStyle;

// ============================================================================
// Helper Functions
// ============================================================================

/// Convert a border style string to the expected BorderStyle enum variant
fn expected_border_style(style_name: &str) -> BorderStyle {
    match style_name {
        "none" => BorderStyle::None,
        "thin" => BorderStyle::Thin,
        "medium" => BorderStyle::Medium,
        "thick" => BorderStyle::Thick,
        "dashed" => BorderStyle::Dashed,
        "dotted" => BorderStyle::Dotted,
        "double" => BorderStyle::Double,
        "hair" => BorderStyle::Hair,
        "mediumDashed" => BorderStyle::MediumDashed,
        "dashDot" => BorderStyle::DashDot,
        "mediumDashDot" => BorderStyle::MediumDashDot,
        "dashDotDot" => BorderStyle::DashDotDot,
        "mediumDashDotDot" => BorderStyle::MediumDashDotDot,
        "slantDashDot" => BorderStyle::SlantDashDot,
        _ => panic!("Unknown border style: {}", style_name),
    }
}

// ============================================================================
// Individual Border Style Tests
// ============================================================================

#[test]
fn test_border_style_none() {
    // "none" style means no visible border - the parser returns None for these
    let style = StyleBuilder::new().border_all("none", None).build();
    let xlsx = XlsxBuilder::new()
        .sheet(SheetBuilder::new("Sheet1").cell("A1", "No Border", Some(style)))
        .build();

    let workbook = load_xlsx(&xlsx);
    let cell = get_cell(&workbook, 0, 0, 0);

    assert!(cell.is_some(), "Cell A1 should exist");

    // "none" style borders are either not present or have BorderStyle::None
    if let Some(cd) = cell {
        if let Some(ref style) = cd.cell.s {
            if let Some(ref border) = style.border_top {
                assert!(
                    matches!(border.style, BorderStyle::None),
                    "none style should map to BorderStyle::None if present"
                );
            }
        }
    }
}

#[test]
fn test_border_style_thin() {
    let style = StyleBuilder::new().border_all("thin", None).build();
    let xlsx = XlsxBuilder::new()
        .sheet(SheetBuilder::new("Sheet1").cell("A1", "Thin Border", Some(style)))
        .build();

    let workbook = load_xlsx(&xlsx);
    let style = get_cell_style(&workbook, 0, 0, 0).expect("Cell should have style");

    assert!(style.border_top.is_some(), "Should have top border");
    assert!(style.border_right.is_some(), "Should have right border");
    assert!(style.border_bottom.is_some(), "Should have bottom border");
    assert!(style.border_left.is_some(), "Should have left border");

    assert!(
        matches!(style.border_top.as_ref().unwrap().style, BorderStyle::Thin),
        "Top border should be thin"
    );
    assert!(
        matches!(
            style.border_right.as_ref().unwrap().style,
            BorderStyle::Thin
        ),
        "Right border should be thin"
    );
    assert!(
        matches!(
            style.border_bottom.as_ref().unwrap().style,
            BorderStyle::Thin
        ),
        "Bottom border should be thin"
    );
    assert!(
        matches!(style.border_left.as_ref().unwrap().style, BorderStyle::Thin),
        "Left border should be thin"
    );
}

#[test]
fn test_border_style_medium() {
    let style = StyleBuilder::new().border_all("medium", None).build();
    let xlsx = XlsxBuilder::new()
        .sheet(SheetBuilder::new("Sheet1").cell("A1", "Medium Border", Some(style)))
        .build();

    let workbook = load_xlsx(&xlsx);
    let style = get_cell_style(&workbook, 0, 0, 0).expect("Cell should have style");

    assert!(style.border_top.is_some(), "Should have top border");
    assert!(
        matches!(
            style.border_top.as_ref().unwrap().style,
            BorderStyle::Medium
        ),
        "Border should be medium"
    );
}

#[test]
fn test_border_style_thick() {
    let style = StyleBuilder::new().border_all("thick", None).build();
    let xlsx = XlsxBuilder::new()
        .sheet(SheetBuilder::new("Sheet1").cell("A1", "Thick Border", Some(style)))
        .build();

    let workbook = load_xlsx(&xlsx);
    let style = get_cell_style(&workbook, 0, 0, 0).expect("Cell should have style");

    assert!(style.border_top.is_some(), "Should have top border");
    assert!(
        matches!(style.border_top.as_ref().unwrap().style, BorderStyle::Thick),
        "Border should be thick"
    );
}

#[test]
fn test_border_style_dashed() {
    let style = StyleBuilder::new().border_all("dashed", None).build();
    let xlsx = XlsxBuilder::new()
        .sheet(SheetBuilder::new("Sheet1").cell("A1", "Dashed Border", Some(style)))
        .build();

    let workbook = load_xlsx(&xlsx);
    let style = get_cell_style(&workbook, 0, 0, 0).expect("Cell should have style");

    assert!(style.border_top.is_some(), "Should have top border");
    assert!(
        matches!(
            style.border_top.as_ref().unwrap().style,
            BorderStyle::Dashed
        ),
        "Border should be dashed"
    );
}

#[test]
fn test_border_style_dotted() {
    let style = StyleBuilder::new().border_all("dotted", None).build();
    let xlsx = XlsxBuilder::new()
        .sheet(SheetBuilder::new("Sheet1").cell("A1", "Dotted Border", Some(style)))
        .build();

    let workbook = load_xlsx(&xlsx);
    let style = get_cell_style(&workbook, 0, 0, 0).expect("Cell should have style");

    assert!(style.border_top.is_some(), "Should have top border");
    assert!(
        matches!(
            style.border_top.as_ref().unwrap().style,
            BorderStyle::Dotted
        ),
        "Border should be dotted"
    );
}

#[test]
fn test_border_style_double() {
    let style = StyleBuilder::new().border_all("double", None).build();
    let xlsx = XlsxBuilder::new()
        .sheet(SheetBuilder::new("Sheet1").cell("A1", "Double Border", Some(style)))
        .build();

    let workbook = load_xlsx(&xlsx);
    let style = get_cell_style(&workbook, 0, 0, 0).expect("Cell should have style");

    assert!(style.border_top.is_some(), "Should have top border");
    assert!(
        matches!(
            style.border_top.as_ref().unwrap().style,
            BorderStyle::Double
        ),
        "Border should be double"
    );
}

#[test]
fn test_border_style_hair() {
    let style = StyleBuilder::new().border_all("hair", None).build();
    let xlsx = XlsxBuilder::new()
        .sheet(SheetBuilder::new("Sheet1").cell("A1", "Hair Border", Some(style)))
        .build();

    let workbook = load_xlsx(&xlsx);
    let style = get_cell_style(&workbook, 0, 0, 0).expect("Cell should have style");

    assert!(style.border_top.is_some(), "Should have top border");
    assert!(
        matches!(style.border_top.as_ref().unwrap().style, BorderStyle::Hair),
        "Border should be hair"
    );
}

#[test]
fn test_border_style_medium_dashed() {
    let style = StyleBuilder::new().border_all("mediumDashed", None).build();
    let xlsx = XlsxBuilder::new()
        .sheet(SheetBuilder::new("Sheet1").cell("A1", "Medium Dashed Border", Some(style)))
        .build();

    let workbook = load_xlsx(&xlsx);
    let style = get_cell_style(&workbook, 0, 0, 0).expect("Cell should have style");

    assert!(style.border_top.is_some(), "Should have top border");
    assert!(
        matches!(
            style.border_top.as_ref().unwrap().style,
            BorderStyle::MediumDashed
        ),
        "Border should be mediumDashed"
    );
}

#[test]
fn test_border_style_dash_dot() {
    let style = StyleBuilder::new().border_all("dashDot", None).build();
    let xlsx = XlsxBuilder::new()
        .sheet(SheetBuilder::new("Sheet1").cell("A1", "Dash Dot Border", Some(style)))
        .build();

    let workbook = load_xlsx(&xlsx);
    let style = get_cell_style(&workbook, 0, 0, 0).expect("Cell should have style");

    assert!(style.border_top.is_some(), "Should have top border");
    assert!(
        matches!(
            style.border_top.as_ref().unwrap().style,
            BorderStyle::DashDot
        ),
        "Border should be dashDot"
    );
}

#[test]
fn test_border_style_medium_dash_dot() {
    let style = StyleBuilder::new()
        .border_all("mediumDashDot", None)
        .build();
    let xlsx = XlsxBuilder::new()
        .sheet(SheetBuilder::new("Sheet1").cell("A1", "Medium Dash Dot Border", Some(style)))
        .build();

    let workbook = load_xlsx(&xlsx);
    let style = get_cell_style(&workbook, 0, 0, 0).expect("Cell should have style");

    assert!(style.border_top.is_some(), "Should have top border");
    assert!(
        matches!(
            style.border_top.as_ref().unwrap().style,
            BorderStyle::MediumDashDot
        ),
        "Border should be mediumDashDot"
    );
}

#[test]
fn test_border_style_dash_dot_dot() {
    let style = StyleBuilder::new().border_all("dashDotDot", None).build();
    let xlsx = XlsxBuilder::new()
        .sheet(SheetBuilder::new("Sheet1").cell("A1", "Dash Dot Dot Border", Some(style)))
        .build();

    let workbook = load_xlsx(&xlsx);
    let style = get_cell_style(&workbook, 0, 0, 0).expect("Cell should have style");

    assert!(style.border_top.is_some(), "Should have top border");
    assert!(
        matches!(
            style.border_top.as_ref().unwrap().style,
            BorderStyle::DashDotDot
        ),
        "Border should be dashDotDot"
    );
}

#[test]
fn test_border_style_medium_dash_dot_dot() {
    let style = StyleBuilder::new()
        .border_all("mediumDashDotDot", None)
        .build();
    let xlsx = XlsxBuilder::new()
        .sheet(SheetBuilder::new("Sheet1").cell("A1", "Medium Dash Dot Dot Border", Some(style)))
        .build();

    let workbook = load_xlsx(&xlsx);
    let style = get_cell_style(&workbook, 0, 0, 0).expect("Cell should have style");

    assert!(style.border_top.is_some(), "Should have top border");
    assert!(
        matches!(
            style.border_top.as_ref().unwrap().style,
            BorderStyle::MediumDashDotDot
        ),
        "Border should be mediumDashDotDot"
    );
}

#[test]
fn test_border_style_slant_dash_dot() {
    let style = StyleBuilder::new().border_all("slantDashDot", None).build();
    let xlsx = XlsxBuilder::new()
        .sheet(SheetBuilder::new("Sheet1").cell("A1", "Slant Dash Dot Border", Some(style)))
        .build();

    let workbook = load_xlsx(&xlsx);
    let style = get_cell_style(&workbook, 0, 0, 0).expect("Cell should have style");

    assert!(style.border_top.is_some(), "Should have top border");
    assert!(
        matches!(
            style.border_top.as_ref().unwrap().style,
            BorderStyle::SlantDashDot
        ),
        "Border should be slantDashDot"
    );
}

// ============================================================================
// Comprehensive Test: All 13 Border Styles
// ============================================================================

/// Test all 13 border styles can be created and parsed correctly
#[test]
fn test_all_13_border_styles_comprehensive() {
    // Skip "none" as it results in no visible border
    let testable_styles: Vec<&&str> = ALL_BORDER_STYLES.iter().filter(|s| **s != "none").collect();

    for style_name in testable_styles {
        let s = StyleBuilder::new().border_all(style_name, None).build();
        let xlsx = XlsxBuilder::new()
            .sheet(SheetBuilder::new("Sheet1").cell(
                "A1",
                format!("{} Border", style_name),
                Some(s),
            ))
            .build();

        let workbook = load_xlsx(&xlsx);
        let style = get_cell_style(&workbook, 0, 0, 0).expect(&format!(
            "Cell should have style for border style {}",
            style_name
        ));

        assert!(
            style.border_top.is_some(),
            "Should have top border for style {}",
            style_name
        );
        assert!(
            style.border_right.is_some(),
            "Should have right border for style {}",
            style_name
        );
        assert!(
            style.border_bottom.is_some(),
            "Should have bottom border for style {}",
            style_name
        );
        assert!(
            style.border_left.is_some(),
            "Should have left border for style {}",
            style_name
        );

        let expected = expected_border_style(style_name);
        let actual_top = style.border_top.as_ref().unwrap().style;
        let actual_right = style.border_right.as_ref().unwrap().style;
        let actual_bottom = style.border_bottom.as_ref().unwrap().style;
        let actual_left = style.border_left.as_ref().unwrap().style;

        assert_eq!(
            actual_top, expected,
            "Top border style mismatch for {}: expected {:?}, got {:?}",
            style_name, expected, actual_top
        );
        assert_eq!(
            actual_right, expected,
            "Right border style mismatch for {}: expected {:?}, got {:?}",
            style_name, expected, actual_right
        );
        assert_eq!(
            actual_bottom, expected,
            "Bottom border style mismatch for {}: expected {:?}, got {:?}",
            style_name, expected, actual_bottom
        );
        assert_eq!(
            actual_left, expected,
            "Left border style mismatch for {}: expected {:?}, got {:?}",
            style_name, expected, actual_left
        );
    }
}

// ============================================================================
// Border Color Tests
// ============================================================================

#[test]
fn test_border_with_color_red() {
    let style = StyleBuilder::new()
        .border_all("thin", Some("#FF0000"))
        .build();
    let xlsx = XlsxBuilder::new()
        .sheet(SheetBuilder::new("Sheet1").cell("A1", "Red Border", Some(style)))
        .build();

    let workbook = load_xlsx(&xlsx);
    let style = get_cell_style(&workbook, 0, 0, 0).expect("Cell should have style");

    assert!(style.border_top.is_some(), "Should have top border");
    let border = style.border_top.as_ref().unwrap();

    assert!(
        border.color.contains("FF0000") || border.color.contains("ff0000"),
        "Border color should be red, got: {}",
        border.color
    );
}

#[test]
fn test_border_with_color_blue() {
    let style = StyleBuilder::new()
        .border_all("medium", Some("#0000FF"))
        .build();
    let xlsx = XlsxBuilder::new()
        .sheet(SheetBuilder::new("Sheet1").cell("A1", "Blue Border", Some(style)))
        .build();

    let workbook = load_xlsx(&xlsx);
    let style = get_cell_style(&workbook, 0, 0, 0).expect("Cell should have style");

    assert!(style.border_top.is_some(), "Should have top border");
    let border = style.border_top.as_ref().unwrap();

    assert!(
        border.color.contains("0000FF") || border.color.contains("0000ff"),
        "Border color should be blue, got: {}",
        border.color
    );
}

#[test]
fn test_border_with_color_green() {
    let style = StyleBuilder::new()
        .border_all("thick", Some("#00FF00"))
        .build();
    let xlsx = XlsxBuilder::new()
        .sheet(SheetBuilder::new("Sheet1").cell("A1", "Green Border", Some(style)))
        .build();

    let workbook = load_xlsx(&xlsx);
    let style = get_cell_style(&workbook, 0, 0, 0).expect("Cell should have style");

    assert!(style.border_top.is_some(), "Should have top border");
    let border = style.border_top.as_ref().unwrap();

    assert!(
        border.color.contains("00FF00") || border.color.contains("00ff00"),
        "Border color should be green, got: {}",
        border.color
    );
}

#[test]
fn test_border_with_custom_color() {
    let style = StyleBuilder::new()
        .border_all("double", Some("#AB12CD"))
        .build();
    let xlsx = XlsxBuilder::new()
        .sheet(SheetBuilder::new("Sheet1").cell("A1", "Custom Color Border", Some(style)))
        .build();

    let workbook = load_xlsx(&xlsx);
    let style = get_cell_style(&workbook, 0, 0, 0).expect("Cell should have style");

    assert!(style.border_top.is_some(), "Should have top border");
    let border = style.border_top.as_ref().unwrap();

    assert!(
        border.color.contains("AB12CD") || border.color.contains("ab12cd"),
        "Border color should be custom color, got: {}",
        border.color
    );
}

// ============================================================================
// All Border Styles with Colors
// ============================================================================

/// Test all border styles work with custom colors
#[test]
fn test_all_border_styles_with_colors() {
    let test_cases = [
        ("thin", "#FF0000"),
        ("medium", "#00FF00"),
        ("thick", "#0000FF"),
        ("dashed", "#FFFF00"),
        ("dotted", "#FF00FF"),
        ("double", "#00FFFF"),
        ("hair", "#800000"),
        ("mediumDashed", "#008000"),
        ("dashDot", "#000080"),
        ("mediumDashDot", "#808000"),
        ("dashDotDot", "#800080"),
        ("mediumDashDotDot", "#008080"),
        ("slantDashDot", "#C0C0C0"),
    ];

    for (style_name, color) in test_cases {
        let s = StyleBuilder::new()
            .border_all(style_name, Some(color))
            .build();
        let xlsx = XlsxBuilder::new()
            .sheet(SheetBuilder::new("Sheet1").cell(
                "A1",
                format!("{} with color", style_name),
                Some(s),
            ))
            .build();

        let workbook = load_xlsx(&xlsx);
        let style = get_cell_style(&workbook, 0, 0, 0).expect(&format!(
            "Cell should have style for {} with color {}",
            style_name, color
        ));

        assert!(
            style.border_top.is_some(),
            "Should have top border for {} with color",
            style_name
        );

        let expected_style = expected_border_style(style_name);
        let border = style.border_top.as_ref().unwrap();

        assert_eq!(
            border.style, expected_style,
            "Border style mismatch for {}: expected {:?}, got {:?}",
            style_name, expected_style, border.style
        );

        // Verify color (strip # and check case-insensitive)
        let color_hex = color.trim_start_matches('#');
        assert!(
            border
                .color
                .to_uppercase()
                .contains(&color_hex.to_uppercase()),
            "Border color should contain {} for style {}, got: {}",
            color_hex,
            style_name,
            border.color
        );
    }
}

// ============================================================================
// Mixed Border Styles Tests
// ============================================================================

#[test]
fn test_different_borders_on_each_side() {
    let style = StyleBuilder::new()
        .border_top(BorderSide::new("thin").color("#FF0000"))
        .border_right(BorderSide::new("medium").color("#00FF00"))
        .border_bottom(BorderSide::new("thick").color("#0000FF"))
        .border_left(BorderSide::new("double").color("#FFFF00"))
        .build();
    let xlsx = XlsxBuilder::new()
        .sheet(SheetBuilder::new("Sheet1").cell("A1", "Mixed Borders", Some(style)))
        .build();

    let workbook = load_xlsx(&xlsx);
    let style = get_cell_style(&workbook, 0, 0, 0).expect("Cell should have style");

    // Check top border - thin, red
    assert!(style.border_top.is_some(), "Should have top border");
    let top = style.border_top.as_ref().unwrap();
    assert!(
        matches!(top.style, BorderStyle::Thin),
        "Top should be thin, got {:?}",
        top.style
    );
    assert!(
        top.color.contains("FF0000") || top.color.contains("ff0000"),
        "Top should be red"
    );

    // Check right border - medium, green
    assert!(style.border_right.is_some(), "Should have right border");
    let right = style.border_right.as_ref().unwrap();
    assert!(
        matches!(right.style, BorderStyle::Medium),
        "Right should be medium, got {:?}",
        right.style
    );
    assert!(
        right.color.contains("00FF00") || right.color.contains("00ff00"),
        "Right should be green"
    );

    // Check bottom border - thick, blue
    assert!(style.border_bottom.is_some(), "Should have bottom border");
    let bottom = style.border_bottom.as_ref().unwrap();
    assert!(
        matches!(bottom.style, BorderStyle::Thick),
        "Bottom should be thick, got {:?}",
        bottom.style
    );
    assert!(
        bottom.color.contains("0000FF") || bottom.color.contains("0000ff"),
        "Bottom should be blue"
    );

    // Check left border - double, yellow
    assert!(style.border_left.is_some(), "Should have left border");
    let left = style.border_left.as_ref().unwrap();
    assert!(
        matches!(left.style, BorderStyle::Double),
        "Left should be double, got {:?}",
        left.style
    );
    assert!(
        left.color.contains("FFFF00") || left.color.contains("ffff00"),
        "Left should be yellow"
    );
}

#[test]
fn test_border_only_top_and_bottom() {
    let style = StyleBuilder::new()
        .border_top(BorderSide::new("medium"))
        .border_bottom(BorderSide::new("medium"))
        .build();
    let xlsx = XlsxBuilder::new()
        .sheet(SheetBuilder::new("Sheet1").cell("A1", "Top and Bottom Only", Some(style)))
        .build();

    let workbook = load_xlsx(&xlsx);
    let style = get_cell_style(&workbook, 0, 0, 0).expect("Cell should have style");

    assert!(style.border_top.is_some(), "Should have top border");
    assert!(style.border_bottom.is_some(), "Should have bottom border");
    assert!(style.border_left.is_none(), "Should not have left border");
    assert!(style.border_right.is_none(), "Should not have right border");
}

// ============================================================================
// Multiple Cells with Different Border Styles
// ============================================================================

#[test]
fn test_multiple_cells_with_different_borders() {
    let thin_style = StyleBuilder::new().border_all("thin", None).build();
    let medium_style = StyleBuilder::new().border_all("medium", None).build();
    let thick_style = StyleBuilder::new().border_all("thick", None).build();
    let double_style = StyleBuilder::new().border_all("double", None).build();

    let xlsx = XlsxBuilder::new()
        .sheet(
            SheetBuilder::new("Sheet1")
                .cell("A1", "Thin", Some(thin_style))
                .cell("B1", "Medium", Some(medium_style))
                .cell("C1", "Thick", Some(thick_style))
                .cell("D1", "Double", Some(double_style)),
        )
        .build();

    let workbook = load_xlsx(&xlsx);

    // Check A1 - thin
    let style_a1 = get_cell_style(&workbook, 0, 0, 0).expect("A1 should have style");
    assert!(
        matches!(
            style_a1.border_top.as_ref().unwrap().style,
            BorderStyle::Thin
        ),
        "A1 should have thin border, got {:?}",
        style_a1.border_top.as_ref().unwrap().style
    );

    // Check B1 - medium
    let style_b1 = get_cell_style(&workbook, 0, 0, 1).expect("B1 should have style");
    assert!(
        matches!(
            style_b1.border_top.as_ref().unwrap().style,
            BorderStyle::Medium
        ),
        "B1 should have medium border, got {:?}",
        style_b1.border_top.as_ref().unwrap().style
    );

    // Check C1 - thick
    let style_c1 = get_cell_style(&workbook, 0, 0, 2).expect("C1 should have style");
    assert!(
        matches!(
            style_c1.border_top.as_ref().unwrap().style,
            BorderStyle::Thick
        ),
        "C1 should have thick border, got {:?}",
        style_c1.border_top.as_ref().unwrap().style
    );

    // Check D1 - double
    let style_d1 = get_cell_style(&workbook, 0, 0, 3).expect("D1 should have style");
    assert!(
        matches!(
            style_d1.border_top.as_ref().unwrap().style,
            BorderStyle::Double
        ),
        "D1 should have double border, got {:?}",
        style_d1.border_top.as_ref().unwrap().style
    );
}

// ============================================================================
// Border Default Color Tests
// ============================================================================

#[test]
fn test_border_without_explicit_color_has_default() {
    let style = StyleBuilder::new().border_all("thin", None).build();
    let xlsx = XlsxBuilder::new()
        .sheet(SheetBuilder::new("Sheet1").cell("A1", "Default Color", Some(style)))
        .build();

    let workbook = load_xlsx(&xlsx);
    let style = get_cell_style(&workbook, 0, 0, 0).expect("Cell should have style");

    assert!(style.border_top.is_some(), "Should have top border");
    let border = style.border_top.as_ref().unwrap();

    // Default color should be black (#000000)
    assert!(
        !border.color.is_empty(),
        "Border should have a color (default black)"
    );
    // The color should be some valid hex color
    let color_clean = border.color.trim_start_matches('#');
    assert!(
        color_clean.len() == 6 || color_clean.len() == 8,
        "Color should be a valid hex color, got: {}",
        border.color
    );
}

// ============================================================================
// Edge Cases
// ============================================================================

#[test]
fn test_border_style_with_combined_formatting() {
    let style = StyleBuilder::new()
        .bold()
        .italic()
        .font_color("#FF0000")
        .bg_color("#FFFF00")
        .border_all("thick", Some("#0000FF"))
        .align_horizontal("center")
        .build();
    let xlsx = XlsxBuilder::new()
        .sheet(SheetBuilder::new("Sheet1").cell("A1", "Combined", Some(style)))
        .build();

    let workbook = load_xlsx(&xlsx);
    let style = get_cell_style(&workbook, 0, 0, 0).expect("Cell should have style");

    // Verify borders still work with other formatting
    assert!(style.border_top.is_some(), "Should have top border");
    assert!(style.border_right.is_some(), "Should have right border");
    assert!(style.border_bottom.is_some(), "Should have bottom border");
    assert!(style.border_left.is_some(), "Should have left border");

    let border = style.border_top.as_ref().unwrap();
    assert!(
        matches!(border.style, BorderStyle::Thick),
        "Border should be thick"
    );
    assert!(
        border.color.contains("0000FF") || border.color.contains("0000ff"),
        "Border should be blue"
    );

    // Verify other formatting is preserved
    assert_eq!(style.bold, Some(true), "Should be bold");
    assert_eq!(style.italic, Some(true), "Should be italic");
}

// ============================================================================
// ALL_BORDER_STYLES Constant Verification
// ============================================================================

/// Verify the ALL_BORDER_STYLES constant has exactly 14 styles (including none)
#[test]
fn test_all_border_styles_constant_count() {
    assert_eq!(
        ALL_BORDER_STYLES.len(),
        14,
        "ALL_BORDER_STYLES should contain exactly 14 styles (none + 13 visible styles)"
    );
}

/// Verify all expected styles are in the ALL_BORDER_STYLES constant
#[test]
fn test_all_border_styles_contains_all_expected() {
    let expected = [
        "none",
        "thin",
        "medium",
        "thick",
        "dashed",
        "dotted",
        "double",
        "hair",
        "mediumDashed",
        "dashDot",
        "mediumDashDot",
        "dashDotDot",
        "mediumDashDotDot",
        "slantDashDot",
    ];

    for style in expected {
        assert!(
            ALL_BORDER_STYLES.contains(&style),
            "ALL_BORDER_STYLES should contain '{}'",
            style
        );
    }
}
