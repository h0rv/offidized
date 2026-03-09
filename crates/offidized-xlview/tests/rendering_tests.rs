//! Headless rendering tests.
//!
//! These tests verify that render data is generated correctly without requiring
//! a browser or Canvas. They test the data transformation from parsed XLSX to
//! render-ready data structures.
//!
//! Test categories:
//! - Font style parsing
//! - Color resolution
//! - Border style parsing
//! - Alignment parsing
//! - Pattern fill parsing
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
use fixtures::{SheetBuilder, StyleBuilder, XlsxBuilder};
use offidized_xlview::render::BorderStyleData;
use offidized_xlview::types::style::VertAlign;
use std::fs;

// ============================================================================
// Font Style Tests
// ============================================================================

#[test]
fn test_bold_cell_parses_correctly() {
    let xlsx = XlsxBuilder::new()
        .sheet(SheetBuilder::new("Sheet1").cell(
            "A1",
            "Bold Text",
            Some(StyleBuilder::new().bold().build()),
        ))
        .build();

    let workbook = load_xlsx(&xlsx);
    let style = common::get_cell_style(&workbook, 0, 0, 0);

    assert!(style.is_some(), "Cell should have style");
    let style = style.unwrap();
    assert_eq!(style.bold, Some(true), "Cell should be bold");
}

#[test]
fn test_italic_cell_parses_correctly() {
    let xlsx = XlsxBuilder::new()
        .sheet(SheetBuilder::new("Sheet1").cell(
            "A1",
            "Italic Text",
            Some(StyleBuilder::new().italic().build()),
        ))
        .build();

    let workbook = load_xlsx(&xlsx);
    let style = common::get_cell_style(&workbook, 0, 0, 0);

    assert!(style.is_some(), "Cell should have style");
    let style = style.unwrap();
    assert_eq!(style.italic, Some(true), "Cell should be italic");
}

#[test]
fn test_font_size_parses_correctly() {
    let xlsx = XlsxBuilder::new()
        .sheet(SheetBuilder::new("Sheet1").cell(
            "A1",
            "Large Text",
            Some(StyleBuilder::new().font_size(24.0).build()),
        ))
        .build();

    let workbook = load_xlsx(&xlsx);
    let style = common::get_cell_style(&workbook, 0, 0, 0);

    assert!(style.is_some(), "Cell should have style");
    let style = style.unwrap();
    assert_eq!(style.font_size, Some(24.0), "Font size should be 24");
}

#[test]
fn test_font_color_parses_correctly() {
    let xlsx = XlsxBuilder::new()
        .sheet(SheetBuilder::new("Sheet1").cell(
            "A1",
            "Red Text",
            Some(StyleBuilder::new().font_color("#FF0000").build()),
        ))
        .build();

    let workbook = load_xlsx(&xlsx);
    let style = common::get_cell_style(&workbook, 0, 0, 0);

    assert!(style.is_some(), "Cell should have style");
    let style = style.unwrap();
    assert!(style.font_color.is_some(), "Cell should have font color");

    let color = style.font_color.as_ref().unwrap();
    assert!(
        color.contains("FF0000") || color.contains("ff0000"),
        "Font color should be red, got {}",
        color
    );
}

#[test]
fn test_combined_font_styles() {
    let xlsx = XlsxBuilder::new()
        .sheet(
            SheetBuilder::new("Sheet1").cell(
                "A1",
                "Combined",
                Some(
                    StyleBuilder::new()
                        .bold()
                        .italic()
                        .underline()
                        .font_size(14.0)
                        .font_color("#0000FF")
                        .build(),
                ),
            ),
        )
        .build();

    let workbook = load_xlsx(&xlsx);
    let style = common::get_cell_style(&workbook, 0, 0, 0);

    assert!(style.is_some(), "Cell should have style");
    let style = style.unwrap();

    assert_eq!(style.bold, Some(true), "Should be bold");
    assert_eq!(style.italic, Some(true), "Should be italic");
    assert!(style.underline.is_some(), "Should be underlined");
    assert_eq!(style.font_size, Some(14.0), "Font size should be 14");
}

// ============================================================================
// Background Color Tests
// ============================================================================

#[test]
fn test_solid_fill_parses_correctly() {
    let xlsx = XlsxBuilder::new()
        .sheet(SheetBuilder::new("Sheet1").cell(
            "A1",
            "Yellow BG",
            Some(StyleBuilder::new().bg_color("#FFFF00").build()),
        ))
        .build();

    let workbook = load_xlsx(&xlsx);
    let style = common::get_cell_style(&workbook, 0, 0, 0);

    assert!(style.is_some(), "Cell should have style");
    let style = style.unwrap();
    assert!(style.bg_color.is_some(), "Cell should have bg_color");

    let color = style.bg_color.as_ref().unwrap();
    assert!(
        color.contains("FFFF00") || color.contains("ffff00"),
        "BG color should be yellow, got {}",
        color
    );
}

// ============================================================================
// Border Tests
// ============================================================================

#[test]
fn test_thin_border_parses_correctly() {
    let xlsx = XlsxBuilder::new()
        .sheet(SheetBuilder::new("Sheet1").cell(
            "A1",
            "Bordered",
            Some(StyleBuilder::new().border_all("thin", None).build()),
        ))
        .build();

    let workbook = load_xlsx(&xlsx);
    let style = common::get_cell_style(&workbook, 0, 0, 0);

    assert!(style.is_some(), "Cell should have style");
    let style = style.unwrap();

    assert!(style.border_top.is_some(), "Should have top border");
    assert!(style.border_right.is_some(), "Should have right border");
    assert!(style.border_bottom.is_some(), "Should have bottom border");
    assert!(style.border_left.is_some(), "Should have left border");
}

#[test]
fn test_colored_border_parses_correctly() {
    let xlsx = XlsxBuilder::new()
        .sheet(
            SheetBuilder::new("Sheet1").cell(
                "A1",
                "Red Border",
                Some(
                    StyleBuilder::new()
                        .border_all("thick", Some("#FF0000"))
                        .build(),
                ),
            ),
        )
        .build();

    let workbook = load_xlsx(&xlsx);
    let style = common::get_cell_style(&workbook, 0, 0, 0);

    assert!(style.is_some(), "Cell should have style");
    let style = style.unwrap();

    assert!(style.border_top.is_some(), "Should have top border");
    let border = style.border_top.as_ref().unwrap();
    let color = &border.color;
    assert!(
        color.contains("FF0000") || color.contains("ff0000"),
        "Border color should be red, got {}",
        color
    );
}

// ============================================================================
// Alignment Tests
// ============================================================================

#[test]
fn test_alignment_parses_correctly() {
    let xlsx = XlsxBuilder::new()
        .sheet(
            SheetBuilder::new("Sheet1").cell(
                "A1",
                "Centered",
                Some(
                    StyleBuilder::new()
                        .align_horizontal("center")
                        .align_vertical("center")
                        .build(),
                ),
            ),
        )
        .build();

    let workbook = load_xlsx(&xlsx);
    let style = common::get_cell_style(&workbook, 0, 0, 0);

    assert!(style.is_some(), "Cell should have style");
    let style = style.unwrap();

    assert!(style.align_h.is_some(), "Should have horizontal alignment");
    assert!(style.align_v.is_some(), "Should have vertical alignment");
}

#[test]
fn test_text_wrap_parses_correctly() {
    let xlsx = XlsxBuilder::new()
        .sheet(SheetBuilder::new("Sheet1").cell(
            "A1",
            "Long text that should wrap",
            Some(StyleBuilder::new().wrap_text().build()),
        ))
        .build();

    let workbook = load_xlsx(&xlsx);
    let style = common::get_cell_style(&workbook, 0, 0, 0);

    assert!(style.is_some(), "Cell should have style");
    let style = style.unwrap();

    assert_eq!(style.wrap, Some(true), "Text should wrap");
}

#[test]
fn test_rotation_parses_correctly() {
    let xlsx = XlsxBuilder::new()
        .sheet(SheetBuilder::new("Sheet1").cell(
            "A1",
            "Rotated",
            Some(StyleBuilder::new().rotation(45).build()),
        ))
        .build();

    let workbook = load_xlsx(&xlsx);
    let style = common::get_cell_style(&workbook, 0, 0, 0);

    assert!(style.is_some(), "Cell should have style");
    let style = style.unwrap();

    assert_eq!(style.rotation, Some(45), "Rotation should be 45 degrees");
}

// ============================================================================
// Real File Tests
// ============================================================================

#[test]
fn test_kitchen_sink_parses_without_panics() {
    let data = fs::read("test/kitchen_sink.xlsx").expect("Failed to read test file");
    let workbook = load_xlsx(&data);

    // Just verify we can iterate all cells without panicking
    let mut total_cells = 0;
    for sheet in &workbook.sheets {
        for _cell_data in &sheet.cells {
            total_cells += 1;
        }
    }

    assert!(total_cells > 0, "Should have parsed some cells");
}

#[test]
fn test_kitchen_sink_v2_style_counts() {
    let data = fs::read("test/kitchen_sink_v2.xlsx").expect("Failed to read test file");
    let workbook = load_xlsx(&data);

    // Count cells with various styles
    let mut bold_count = 0;
    let mut colored_bg_count = 0;
    let mut border_count = 0;

    for sheet in &workbook.sheets {
        for cell_data in &sheet.cells {
            if let Some(style_idx) = cell_data.cell.style_idx {
                if let Some(Some(ref style)) = workbook.resolved_styles.get(style_idx as usize) {
                    if style.bold == Some(true) {
                        bold_count += 1;
                    }
                    if style.bg_color.is_some() {
                        colored_bg_count += 1;
                    }
                    if style.border_top.is_some()
                        || style.border_right.is_some()
                        || style.border_bottom.is_some()
                        || style.border_left.is_some()
                    {
                        border_count += 1;
                    }
                }
            }
            // Also check inline style
            if let Some(ref style) = cell_data.cell.s {
                if style.bold == Some(true) {
                    bold_count += 1;
                }
                if style.bg_color.is_some() {
                    colored_bg_count += 1;
                }
                if style.border_top.is_some()
                    || style.border_right.is_some()
                    || style.border_bottom.is_some()
                    || style.border_left.is_some()
                {
                    border_count += 1;
                }
            }
        }
    }

    // Kitchen sink v2 should have various styled cells
    assert!(bold_count > 0, "Should have bold cells");
    assert!(
        colored_bg_count > 0,
        "Should have cells with colored backgrounds"
    );
    assert!(border_count > 0, "Should have cells with borders");
}

// ============================================================================
// Color Resolution Tests
// ============================================================================

#[test]
fn test_theme_colors_parsed() {
    let data = fs::read("test/kitchen_sink.xlsx").expect("Failed to read test file");
    let workbook = load_xlsx(&data);

    // Verify theme colors are available
    assert!(
        !workbook.theme.colors.is_empty(),
        "Should have theme colors parsed"
    );
}

#[test]
fn test_rgb_color_format() {
    let xlsx = XlsxBuilder::new()
        .sheet(SheetBuilder::new("Sheet1").cell(
            "A1",
            "Blue",
            Some(StyleBuilder::new().font_color("#0000FF").build()),
        ))
        .build();

    let workbook = load_xlsx(&xlsx);
    let style = common::get_cell_style(&workbook, 0, 0, 0);

    assert!(style.is_some(), "Cell should have style");
    let style = style.unwrap();

    if let Some(ref color) = style.font_color {
        // Color should be in a valid format (with or without # prefix, with or without alpha)
        let hex_part = color.strip_prefix('#').unwrap_or(color);
        assert!(
            hex_part.len() == 6 || hex_part.len() == 8,
            "Color should be 6 or 8 hex chars (optionally with #), got: {}",
            color
        );
    }
}

// ============================================================================
// Border Style Width Tests
// ============================================================================

#[test]
fn test_border_width_calculations() {
    // Test the BorderStyleData width calculations
    let thin = BorderStyleData {
        style: Some("thin".to_string()),
        color: None,
    };
    assert_eq!(thin.width(), 1.0, "Thin border should be 1px");

    let medium = BorderStyleData {
        style: Some("medium".to_string()),
        color: None,
    };
    assert_eq!(medium.width(), 2.0, "Medium border should be 2px");

    let thick = BorderStyleData {
        style: Some("thick".to_string()),
        color: None,
    };
    assert_eq!(thick.width(), 3.0, "Thick border should be 3px");

    let double = BorderStyleData {
        style: Some("double".to_string()),
        color: None,
    };
    assert_eq!(double.width(), 3.0, "Double border should be 3px");

    let hair = BorderStyleData {
        style: Some("hair".to_string()),
        color: None,
    };
    assert_eq!(hair.width(), 1.0, "Hair border should be 1px");

    let none = BorderStyleData {
        style: None,
        color: None,
    };
    assert_eq!(none.width(), 1.0, "No style defaults to 1px");
}

// ============================================================================
// Subscript and Superscript Tests
// ============================================================================

#[test]
fn test_subscript_cell_parses_correctly() {
    let xlsx = XlsxBuilder::new()
        .sheet(SheetBuilder::new("Sheet1").cell(
            "A1",
            "H2O",
            Some(StyleBuilder::new().subscript().build()),
        ))
        .build();

    let workbook = load_xlsx(&xlsx);
    let style = common::get_cell_style(&workbook, 0, 0, 0);

    assert!(style.is_some(), "Cell should have style");
    let style = style.unwrap();
    assert!(style.vert_align.is_some(), "Cell should have vert_align");

    let vert_align = style.vert_align.as_ref().unwrap();
    assert!(
        matches!(vert_align, VertAlign::Subscript),
        "Expected subscript, got {:?}",
        vert_align
    );
}

#[test]
fn test_superscript_cell_parses_correctly() {
    let xlsx = XlsxBuilder::new()
        .sheet(SheetBuilder::new("Sheet1").cell(
            "A1",
            "E=mc2",
            Some(StyleBuilder::new().superscript().build()),
        ))
        .build();

    let workbook = load_xlsx(&xlsx);
    let style = common::get_cell_style(&workbook, 0, 0, 0);

    assert!(style.is_some(), "Cell should have style");
    let style = style.unwrap();
    assert!(style.vert_align.is_some(), "Cell should have vert_align");

    let vert_align = style.vert_align.as_ref().unwrap();
    assert!(
        matches!(vert_align, VertAlign::Superscript),
        "Expected superscript, got {:?}",
        vert_align
    );
}

#[test]
fn test_subscript_with_other_styles() {
    let xlsx = XlsxBuilder::new()
        .sheet(
            SheetBuilder::new("Sheet1").cell(
                "A1",
                "Subscript Bold",
                Some(
                    StyleBuilder::new()
                        .subscript()
                        .bold()
                        .font_color("#FF0000")
                        .build(),
                ),
            ),
        )
        .build();

    let workbook = load_xlsx(&xlsx);
    let style = common::get_cell_style(&workbook, 0, 0, 0);

    assert!(style.is_some(), "Cell should have style");
    let style = style.unwrap();

    // Check all styles are applied
    assert!(style.vert_align.is_some(), "Cell should have vert_align");
    assert_eq!(style.bold, Some(true), "Cell should be bold");
    assert!(style.font_color.is_some(), "Cell should have font color");

    let vert_align = style.vert_align.as_ref().unwrap();
    assert!(
        matches!(vert_align, VertAlign::Subscript),
        "Expected subscript, got {:?}",
        vert_align
    );
}

#[test]
fn test_superscript_with_other_styles() {
    let xlsx = XlsxBuilder::new()
        .sheet(
            SheetBuilder::new("Sheet1").cell(
                "A1",
                "Superscript Italic",
                Some(
                    StyleBuilder::new()
                        .superscript()
                        .italic()
                        .font_size(12.0)
                        .build(),
                ),
            ),
        )
        .build();

    let workbook = load_xlsx(&xlsx);
    let style = common::get_cell_style(&workbook, 0, 0, 0);

    assert!(style.is_some(), "Cell should have style");
    let style = style.unwrap();

    // Check all styles are applied
    assert!(style.vert_align.is_some(), "Cell should have vert_align");
    assert_eq!(style.italic, Some(true), "Cell should be italic");
    assert_eq!(style.font_size, Some(12.0), "Font size should be 12");

    let vert_align = style.vert_align.as_ref().unwrap();
    assert!(
        matches!(vert_align, VertAlign::Superscript),
        "Expected superscript, got {:?}",
        vert_align
    );
}

#[test]
fn test_multiple_cells_with_different_vert_align() {
    let xlsx = XlsxBuilder::new()
        .sheet(
            SheetBuilder::new("Sheet1")
                .cell("A1", "Normal", None)
                .cell(
                    "A2",
                    "Subscript",
                    Some(StyleBuilder::new().subscript().build()),
                )
                .cell(
                    "A3",
                    "Superscript",
                    Some(StyleBuilder::new().superscript().build()),
                ),
        )
        .build();

    let workbook = load_xlsx(&xlsx);

    // Cell A1 - no vert_align
    let style_a1 = common::get_cell_style(&workbook, 0, 0, 0);
    if let Some(style) = style_a1 {
        assert!(
            style.vert_align.is_none(),
            "Cell A1 should not have vert_align"
        );
    }

    // Cell A2 - subscript
    let style_a2 = common::get_cell_style(&workbook, 0, 1, 0);
    assert!(style_a2.is_some(), "Cell A2 should have style");
    let style_a2 = style_a2.unwrap();
    assert!(
        matches!(style_a2.vert_align, Some(VertAlign::Subscript)),
        "Cell A2 should be subscript"
    );

    // Cell A3 - superscript
    let style_a3 = common::get_cell_style(&workbook, 0, 2, 0);
    assert!(style_a3.is_some(), "Cell A3 should have style");
    let style_a3 = style_a3.unwrap();
    assert!(
        matches!(style_a3.vert_align, Some(VertAlign::Superscript)),
        "Cell A3 should be superscript"
    );
}

#[test]
fn test_kitchen_sink_parses_vert_align_without_panics() {
    let data = fs::read("test/kitchen_sink.xlsx").expect("Failed to read test file");
    let workbook = load_xlsx(&data);

    // Count cells with vert_align (subscript/superscript)
    let mut subscript_count = 0;
    let mut superscript_count = 0;

    for sheet in &workbook.sheets {
        for cell_data in &sheet.cells {
            let style = cell_data.cell.s.as_ref().or_else(|| {
                cell_data
                    .cell
                    .style_idx
                    .and_then(|idx| workbook.resolved_styles.get(idx as usize))
                    .and_then(|opt| opt.as_ref())
            });
            if let Some(style) = style {
                match style.vert_align {
                    Some(VertAlign::Subscript) => subscript_count += 1,
                    Some(VertAlign::Superscript) => superscript_count += 1,
                    _ => {}
                }
            }
        }
    }

    // The kitchen sink file may or may not have subscript/superscript cells
    // This test just verifies we can check for them without panicking
    println!(
        "kitchen_sink.xlsx: {} subscript cells, {} superscript cells",
        subscript_count, superscript_count
    );
}

#[test]
fn test_kitchen_sink_v2_parses_vert_align_without_panics() {
    let data = fs::read("test/kitchen_sink_v2.xlsx").expect("Failed to read test file");
    let workbook = load_xlsx(&data);

    // Count cells with vert_align (subscript/superscript)
    let mut subscript_count = 0;
    let mut superscript_count = 0;

    for sheet in &workbook.sheets {
        for cell_data in &sheet.cells {
            let style = cell_data.cell.s.as_ref().or_else(|| {
                cell_data
                    .cell
                    .style_idx
                    .and_then(|idx| workbook.resolved_styles.get(idx as usize))
                    .and_then(|opt| opt.as_ref())
            });
            if let Some(style) = style {
                match style.vert_align {
                    Some(VertAlign::Subscript) => subscript_count += 1,
                    Some(VertAlign::Superscript) => superscript_count += 1,
                    _ => {}
                }
            }
        }
    }

    // The kitchen sink v2 file may or may not have subscript/superscript cells
    // This test just verifies we can check for them without panicking
    println!(
        "kitchen_sink_v2.xlsx: {} subscript cells, {} superscript cells",
        subscript_count, superscript_count
    );
}
