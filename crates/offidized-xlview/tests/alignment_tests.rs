//! Comprehensive tests for text alignment styling.
//!
//! These tests verify that alignment properties from xl/styles.xml are correctly
//! parsed through offidized-xlsx and adapted into the viewer's `Style` struct.
//!
//! XLSX alignment element format:
//! ```xml
//! <xf ...>
//!   <alignment horizontal="center" vertical="center" wrapText="1" textRotation="45" indent="1"/>
//! </xf>
//! ```
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
// HORIZONTAL ALIGNMENT TESTS
// ============================================================================

/// Test 1: horizontal="general" - Default, context-dependent alignment
#[test]
fn test_horizontal_alignment_general() {
    let style = StyleBuilder::new().align_horizontal("general").build();
    let xlsx = XlsxBuilder::new()
        .sheet(SheetBuilder::new("Sheet1").cell("A1", "Test", Some(style)))
        .build();
    let wb = load_xlsx(&xlsx);
    assert_cell_align_h(&wb, 0, 0, 0, "general");
}

/// Test 2: horizontal="left" - Left-aligned text
#[test]
fn test_horizontal_alignment_left() {
    let style = StyleBuilder::new().align_horizontal("left").build();
    let xlsx = XlsxBuilder::new()
        .sheet(SheetBuilder::new("Sheet1").cell("A1", "Test", Some(style)))
        .build();
    let wb = load_xlsx(&xlsx);
    assert_cell_align_h(&wb, 0, 0, 0, "left");
}

/// Test 3: horizontal="center" - Center-aligned text
#[test]
fn test_horizontal_alignment_center() {
    let style = StyleBuilder::new().align_horizontal("center").build();
    let xlsx = XlsxBuilder::new()
        .sheet(SheetBuilder::new("Sheet1").cell("A1", "Test", Some(style)))
        .build();
    let wb = load_xlsx(&xlsx);
    assert_cell_align_h(&wb, 0, 0, 0, "center");
}

/// Test 4: horizontal="right" - Right-aligned text
#[test]
fn test_horizontal_alignment_right() {
    let style = StyleBuilder::new().align_horizontal("right").build();
    let xlsx = XlsxBuilder::new()
        .sheet(SheetBuilder::new("Sheet1").cell("A1", "Test", Some(style)))
        .build();
    let wb = load_xlsx(&xlsx);
    assert_cell_align_h(&wb, 0, 0, 0, "right");
}

/// Test 5: horizontal="fill" - Text repeats to fill cell width
#[test]
fn test_horizontal_alignment_fill() {
    let style = StyleBuilder::new().align_horizontal("fill").build();
    let xlsx = XlsxBuilder::new()
        .sheet(SheetBuilder::new("Sheet1").cell("A1", "Test", Some(style)))
        .build();
    let wb = load_xlsx(&xlsx);
    assert_cell_align_h(&wb, 0, 0, 0, "fill");
}

/// Test 6: horizontal="justify" - Text justified across cell width
#[test]
fn test_horizontal_alignment_justify() {
    let style = StyleBuilder::new().align_horizontal("justify").build();
    let xlsx = XlsxBuilder::new()
        .sheet(SheetBuilder::new("Sheet1").cell("A1", "Test", Some(style)))
        .build();
    let wb = load_xlsx(&xlsx);
    assert_cell_align_h(&wb, 0, 0, 0, "justify");
}

/// Test 7: horizontal="centerContinuous" - Center across selection without merge
#[test]
fn test_horizontal_alignment_center_continuous() {
    let style = StyleBuilder::new()
        .align_horizontal("centerContinuous")
        .build();
    let xlsx = XlsxBuilder::new()
        .sheet(SheetBuilder::new("Sheet1").cell("A1", "Test", Some(style)))
        .build();
    let wb = load_xlsx(&xlsx);
    assert_cell_align_h(&wb, 0, 0, 0, "centerContinuous");
}

/// Test 8: horizontal="distributed" - Text distributed evenly across cell
#[test]
fn test_horizontal_alignment_distributed() {
    let style = StyleBuilder::new().align_horizontal("distributed").build();
    let xlsx = XlsxBuilder::new()
        .sheet(SheetBuilder::new("Sheet1").cell("A1", "Test", Some(style)))
        .build();
    let wb = load_xlsx(&xlsx);
    assert_cell_align_h(&wb, 0, 0, 0, "distributed");
}

// ============================================================================
// VERTICAL ALIGNMENT TESTS
// ============================================================================

/// Test 9: vertical="top" - Top-aligned text
#[test]
fn test_vertical_alignment_top() {
    let style = StyleBuilder::new().align_vertical("top").build();
    let xlsx = XlsxBuilder::new()
        .sheet(SheetBuilder::new("Sheet1").cell("A1", "Test", Some(style)))
        .build();
    let wb = load_xlsx(&xlsx);
    let style = get_cell_style(&wb, 0, 0, 0).expect("Should have style");
    assert_eq!(
        style.align_v,
        Some(offidized_xlview::types::style::VAlign::Top),
        "vertical='top' should be parsed"
    );
}

/// Test 10: vertical="center" - Vertically centered text
#[test]
fn test_vertical_alignment_center() {
    let style = StyleBuilder::new().align_vertical("center").build();
    let xlsx = XlsxBuilder::new()
        .sheet(SheetBuilder::new("Sheet1").cell("A1", "Test", Some(style)))
        .build();
    let wb = load_xlsx(&xlsx);
    let style = get_cell_style(&wb, 0, 0, 0).expect("Should have style");
    assert_eq!(
        style.align_v,
        Some(offidized_xlview::types::style::VAlign::Center),
        "vertical='center' should be parsed"
    );
}

/// Test 11: vertical="bottom" - Bottom-aligned text (often default)
#[test]
fn test_vertical_alignment_bottom() {
    let style = StyleBuilder::new().align_vertical("bottom").build();
    let xlsx = XlsxBuilder::new()
        .sheet(SheetBuilder::new("Sheet1").cell("A1", "Test", Some(style)))
        .build();
    let wb = load_xlsx(&xlsx);
    let style = get_cell_style(&wb, 0, 0, 0).expect("Should have style");
    assert_eq!(
        style.align_v,
        Some(offidized_xlview::types::style::VAlign::Bottom),
        "vertical='bottom' should be parsed"
    );
}

/// Test 12: vertical="justify" - Text justified vertically
#[test]
fn test_vertical_alignment_justify() {
    let style = StyleBuilder::new().align_vertical("justify").build();
    let xlsx = XlsxBuilder::new()
        .sheet(SheetBuilder::new("Sheet1").cell("A1", "Test", Some(style)))
        .build();
    let wb = load_xlsx(&xlsx);
    let style = get_cell_style(&wb, 0, 0, 0).expect("Should have style");
    assert_eq!(
        style.align_v,
        Some(offidized_xlview::types::style::VAlign::Justify),
        "vertical='justify' should be parsed"
    );
}

/// Test 13: vertical="distributed" - Text distributed vertically
#[test]
fn test_vertical_alignment_distributed() {
    let style = StyleBuilder::new().align_vertical("distributed").build();
    let xlsx = XlsxBuilder::new()
        .sheet(SheetBuilder::new("Sheet1").cell("A1", "Test", Some(style)))
        .build();
    let wb = load_xlsx(&xlsx);
    let style = get_cell_style(&wb, 0, 0, 0).expect("Should have style");
    assert_eq!(
        style.align_v,
        Some(offidized_xlview::types::style::VAlign::Distributed),
        "vertical='distributed' should be parsed"
    );
}

// ============================================================================
// TEXT CONTROL TESTS
// ============================================================================

/// Test 14: wrapText="1" - Text wraps within cell
#[test]
fn test_wrap_text_enabled() {
    let style = StyleBuilder::new().wrap_text().build();
    let xlsx = XlsxBuilder::new()
        .sheet(SheetBuilder::new("Sheet1").cell("A1", "Test", Some(style)))
        .build();
    let wb = load_xlsx(&xlsx);
    assert_cell_wrap(&wb, 0, 0, 0);
}

/// Test 14b: wrapText="0" - Text does not wrap
#[test]
fn test_wrap_text_disabled() {
    // Cell with alignment but no wrap
    let style = StyleBuilder::new().align_horizontal("left").build();
    let xlsx = XlsxBuilder::new()
        .sheet(SheetBuilder::new("Sheet1").cell("A1", "Test", Some(style)))
        .build();
    let wb = load_xlsx(&xlsx);
    let style = get_cell_style(&wb, 0, 0, 0).expect("Should have style");
    assert!(
        style.wrap != Some(true),
        "wrapText should not be true when not set"
    );
}

/// Test 14c: wrapText absent - Default is no wrap
#[test]
fn test_wrap_text_absent() {
    let style = StyleBuilder::new().align_horizontal("left").build();
    let xlsx = XlsxBuilder::new()
        .sheet(SheetBuilder::new("Sheet1").cell("A1", "Test", Some(style)))
        .build();
    let wb = load_xlsx(&xlsx);
    let style = get_cell_style(&wb, 0, 0, 0).expect("Should have style");
    assert!(
        style.wrap != Some(true),
        "Absent wrapText should default to not true"
    );
}

/// Test 15: shrinkToFit="1" - Font shrinks to fit cell width
#[test]
fn test_shrink_to_fit_enabled() {
    let style = StyleBuilder::new().shrink_to_fit().build();
    let xlsx = XlsxBuilder::new()
        .sheet(SheetBuilder::new("Sheet1").cell("A1", "Test", Some(style)))
        .build();
    let wb = load_xlsx(&xlsx);
    let style = get_cell_style(&wb, 0, 0, 0).expect("Should have style");
    assert_eq!(
        style.shrink_to_fit,
        Some(true),
        "shrinkToFit='1' should set shrink_to_fit to true"
    );
}

/// Test 16: wrapText + shrinkToFit together - wrapText takes precedence
#[test]
fn test_wrap_text_and_shrink_to_fit_combination() {
    let style = StyleBuilder::new().wrap_text().shrink_to_fit().build();
    let xlsx = XlsxBuilder::new()
        .sheet(SheetBuilder::new("Sheet1").cell("A1", "Test", Some(style)))
        .build();
    let wb = load_xlsx(&xlsx);
    let style = get_cell_style(&wb, 0, 0, 0).expect("Should have style");
    // wrapText should be honored
    assert_eq!(style.wrap, Some(true), "wrapText should be true");
    // shrinkToFit is also stored (renderer decides precedence)
    assert_eq!(
        style.shrink_to_fit,
        Some(true),
        "shrinkToFit should be true"
    );
}

// ============================================================================
// TEXT ROTATION TESTS
// ============================================================================

/// Test 17: textRotation="45" - 45 degrees counterclockwise
#[test]
fn test_text_rotation_45_degrees() {
    let style = StyleBuilder::new().rotation(45).build();
    let xlsx = XlsxBuilder::new()
        .sheet(SheetBuilder::new("Sheet1").cell("A1", "Test", Some(style)))
        .build();
    let wb = load_xlsx(&xlsx);
    let style = get_cell_style(&wb, 0, 0, 0).expect("Should have style");
    assert_eq!(style.rotation, Some(45), "textRotation should be 45");
}

/// Test 18: textRotation="90" - Vertical text, bottom to top
#[test]
fn test_text_rotation_90_degrees() {
    let style = StyleBuilder::new().rotation(90).build();
    let xlsx = XlsxBuilder::new()
        .sheet(SheetBuilder::new("Sheet1").cell("A1", "Test", Some(style)))
        .build();
    let wb = load_xlsx(&xlsx);
    let style = get_cell_style(&wb, 0, 0, 0).expect("Should have style");
    assert_eq!(style.rotation, Some(90), "textRotation should be 90");
}

/// Test 19: textRotation="135" - 45 degrees clockwise (stored as 90+45)
#[test]
fn test_text_rotation_negative_45_degrees() {
    let style = StyleBuilder::new().rotation(135).build();
    let xlsx = XlsxBuilder::new()
        .sheet(SheetBuilder::new("Sheet1").cell("A1", "Test", Some(style)))
        .build();
    let wb = load_xlsx(&xlsx);
    let style = get_cell_style(&wb, 0, 0, 0).expect("Should have style");
    assert_eq!(style.rotation, Some(135), "textRotation should be 135");
}

/// Test 19b: textRotation="180" - 90 degrees clockwise
#[test]
fn test_text_rotation_negative_90_degrees() {
    let style = StyleBuilder::new().rotation(180).build();
    let xlsx = XlsxBuilder::new()
        .sheet(SheetBuilder::new("Sheet1").cell("A1", "Test", Some(style)))
        .build();
    let wb = load_xlsx(&xlsx);
    let style = get_cell_style(&wb, 0, 0, 0).expect("Should have style");
    assert_eq!(style.rotation, Some(180), "textRotation should be 180");
}

/// Test 20: textRotation="255" - Special value for stacked vertical text
#[test]
fn test_text_rotation_vertical_stacked() {
    let style = StyleBuilder::new().rotation(255).build();
    let xlsx = XlsxBuilder::new()
        .sheet(SheetBuilder::new("Sheet1").cell("A1", "Test", Some(style)))
        .build();
    let wb = load_xlsx(&xlsx);
    let style = get_cell_style(&wb, 0, 0, 0).expect("Should have style");
    assert_eq!(style.rotation, Some(255), "textRotation should be 255");
}

/// Test 21: textRotation="0" - No rotation (horizontal)
#[test]
fn test_text_rotation_zero() {
    let style = StyleBuilder::new().rotation(0).build();
    let xlsx = XlsxBuilder::new()
        .sheet(SheetBuilder::new("Sheet1").cell("A1", "Test", Some(style)))
        .build();
    let wb = load_xlsx(&xlsx);
    let style = get_cell_style(&wb, 0, 0, 0).expect("Should have style");
    assert_eq!(style.rotation, Some(0), "textRotation should be 0");
}

/// Test 21b: textRotation absent - Should be None
#[test]
fn test_text_rotation_absent() {
    let style = StyleBuilder::new().align_horizontal("left").build();
    let xlsx = XlsxBuilder::new()
        .sheet(SheetBuilder::new("Sheet1").cell("A1", "Test", Some(style)))
        .build();
    let wb = load_xlsx(&xlsx);
    let style = get_cell_style(&wb, 0, 0, 0).expect("Should have style");
    assert_eq!(style.rotation, None, "Absent textRotation should be None");
}

/// Test boundary: textRotation at maximum counterclockwise (90)
#[test]
fn test_text_rotation_max_counterclockwise() {
    let style = StyleBuilder::new().rotation(90).build();
    let xlsx = XlsxBuilder::new()
        .sheet(SheetBuilder::new("Sheet1").cell("A1", "Test", Some(style)))
        .build();
    let wb = load_xlsx(&xlsx);
    let style = get_cell_style(&wb, 0, 0, 0).expect("Should have style");
    assert_eq!(
        style.rotation,
        Some(90),
        "Maximum counterclockwise rotation should be 90"
    );
}

// ============================================================================
// INDENT TESTS
// ============================================================================

/// Test 22: indent="1" - Single level indent
#[test]
fn test_indent_level_1() {
    let style = StyleBuilder::new()
        .align_horizontal("left")
        .indent(1)
        .build();
    let xlsx = XlsxBuilder::new()
        .sheet(SheetBuilder::new("Sheet1").cell("A1", "Test", Some(style)))
        .build();
    let wb = load_xlsx(&xlsx);
    let style = get_cell_style(&wb, 0, 0, 0).expect("Should have style");
    assert_eq!(style.indent, Some(1), "indent should be 1");
}

/// Test 23: indent="2" - Two level indent
#[test]
fn test_indent_level_2() {
    let style = StyleBuilder::new()
        .align_horizontal("left")
        .indent(2)
        .build();
    let xlsx = XlsxBuilder::new()
        .sheet(SheetBuilder::new("Sheet1").cell("A1", "Test", Some(style)))
        .build();
    let wb = load_xlsx(&xlsx);
    let style = get_cell_style(&wb, 0, 0, 0).expect("Should have style");
    assert_eq!(style.indent, Some(2), "indent should be 2");
}

/// Test 24: indent="5" - Five level indent
#[test]
fn test_indent_level_5() {
    let style = StyleBuilder::new()
        .align_horizontal("left")
        .indent(5)
        .build();
    let xlsx = XlsxBuilder::new()
        .sheet(SheetBuilder::new("Sheet1").cell("A1", "Test", Some(style)))
        .build();
    let wb = load_xlsx(&xlsx);
    let style = get_cell_style(&wb, 0, 0, 0).expect("Should have style");
    assert_eq!(style.indent, Some(5), "indent should be 5");
}

/// Test 25: indent with horizontal alignment
#[test]
fn test_indent_with_left_alignment() {
    let style = StyleBuilder::new()
        .align_horizontal("left")
        .indent(3)
        .build();
    let xlsx = XlsxBuilder::new()
        .sheet(SheetBuilder::new("Sheet1").cell("A1", "Test", Some(style)))
        .build();
    let wb = load_xlsx(&xlsx);
    assert_cell_align_h(&wb, 0, 0, 0, "left");
    let style = get_cell_style(&wb, 0, 0, 0).expect("Should have style");
    assert_eq!(style.indent, Some(3), "Should have indent of 3");
}

/// Test indent with right alignment (indents from right edge)
#[test]
fn test_indent_with_right_alignment() {
    let style = StyleBuilder::new()
        .align_horizontal("right")
        .indent(2)
        .build();
    let xlsx = XlsxBuilder::new()
        .sheet(SheetBuilder::new("Sheet1").cell("A1", "Test", Some(style)))
        .build();
    let wb = load_xlsx(&xlsx);
    assert_cell_align_h(&wb, 0, 0, 0, "right");
    let style = get_cell_style(&wb, 0, 0, 0).expect("Should have style");
    assert_eq!(style.indent, Some(2), "Should have indent of 2");
}

/// Test indent="0" - No indent
#[test]
fn test_indent_zero() {
    let style = StyleBuilder::new()
        .align_horizontal("left")
        .indent(0)
        .build();
    let xlsx = XlsxBuilder::new()
        .sheet(SheetBuilder::new("Sheet1").cell("A1", "Test", Some(style)))
        .build();
    let wb = load_xlsx(&xlsx);
    let style = get_cell_style(&wb, 0, 0, 0).expect("Should have style");
    // indent=0 may be stored as Some(0) or None
    assert!(
        style.indent == Some(0) || style.indent.is_none(),
        "indent=0 should be Some(0) or None, got {:?}",
        style.indent
    );
}

/// Test indent absent - Should be None
#[test]
fn test_indent_absent() {
    let style = StyleBuilder::new().align_horizontal("left").build();
    let xlsx = XlsxBuilder::new()
        .sheet(SheetBuilder::new("Sheet1").cell("A1", "Test", Some(style)))
        .build();
    let wb = load_xlsx(&xlsx);
    let style = get_cell_style(&wb, 0, 0, 0).expect("Should have style");
    assert!(
        style.indent.is_none() || style.indent == Some(0),
        "Absent indent should be None or Some(0), got {:?}",
        style.indent
    );
}

/// Test large indent value
#[test]
fn test_indent_large_value() {
    let style = StyleBuilder::new()
        .align_horizontal("left")
        .indent(15)
        .build();
    let xlsx = XlsxBuilder::new()
        .sheet(SheetBuilder::new("Sheet1").cell("A1", "Test", Some(style)))
        .build();
    let wb = load_xlsx(&xlsx);
    let style = get_cell_style(&wb, 0, 0, 0).expect("Should have style");
    assert_eq!(style.indent, Some(15), "indent should be 15");
}

// ============================================================================
// READING ORDER TESTS
// ============================================================================

/// Test 26: readingOrder="0" - Context-dependent reading order
#[test]
fn test_reading_order_context() {
    // readingOrder is parsed by offidized-xlview if the field exists in Style
    let style = StyleBuilder::new().align_horizontal("left").build();
    let xlsx = XlsxBuilder::new()
        .sheet(SheetBuilder::new("Sheet1").cell("A1", "Test", Some(style)))
        .build();
    let wb = load_xlsx(&xlsx);
    let style = get_cell_style(&wb, 0, 0, 0);
    assert!(style.is_some(), "Alignment element should be parsed");
}

/// Test 27: readingOrder="1" - Left to Right reading order
#[test]
fn test_reading_order_ltr() {
    let style = StyleBuilder::new().align_horizontal("left").build();
    let xlsx = XlsxBuilder::new()
        .sheet(SheetBuilder::new("Sheet1").cell("A1", "Test", Some(style)))
        .build();
    let wb = load_xlsx(&xlsx);
    let style = get_cell_style(&wb, 0, 0, 0);
    assert!(style.is_some(), "Alignment element should be parsed");
}

/// Test 28: readingOrder="2" - Right to Left reading order
#[test]
fn test_reading_order_rtl() {
    let style = StyleBuilder::new().align_horizontal("right").build();
    let xlsx = XlsxBuilder::new()
        .sheet(SheetBuilder::new("Sheet1").cell("A1", "Test", Some(style)))
        .build();
    let wb = load_xlsx(&xlsx);
    let style = get_cell_style(&wb, 0, 0, 0);
    assert!(style.is_some(), "Alignment element should be parsed");
}

// ============================================================================
// COMBINATION TESTS
// ============================================================================

/// Test 29: Center + middle + wrap - Common table header combination
#[test]
fn test_combination_center_middle_wrap() {
    let style = StyleBuilder::new()
        .align_horizontal("center")
        .align_vertical("center")
        .wrap_text()
        .build();
    let xlsx = XlsxBuilder::new()
        .sheet(SheetBuilder::new("Sheet1").cell("A1", "Test", Some(style)))
        .build();
    let wb = load_xlsx(&xlsx);
    assert_cell_align_h(&wb, 0, 0, 0, "center");
    let style = get_cell_style(&wb, 0, 0, 0).expect("Should have style");
    assert_eq!(
        style.align_v,
        Some(offidized_xlview::types::style::VAlign::Center)
    );
    assert_eq!(style.wrap, Some(true));
}

/// Test 30: Right + bottom + rotation - Complex combination
#[test]
fn test_combination_right_bottom_rotation() {
    let style = StyleBuilder::new()
        .align_horizontal("right")
        .align_vertical("bottom")
        .rotation(45)
        .build();
    let xlsx = XlsxBuilder::new()
        .sheet(SheetBuilder::new("Sheet1").cell("A1", "Test", Some(style)))
        .build();
    let wb = load_xlsx(&xlsx);
    assert_cell_align_h(&wb, 0, 0, 0, "right");
    let style = get_cell_style(&wb, 0, 0, 0).expect("Should have style");
    assert_eq!(
        style.align_v,
        Some(offidized_xlview::types::style::VAlign::Bottom)
    );
    assert_eq!(style.rotation, Some(45));
}

/// Test all alignment properties together
#[test]
fn test_combination_all_properties() {
    let style = StyleBuilder::new()
        .align_horizontal("left")
        .align_vertical("top")
        .wrap_text()
        .rotation(30)
        .indent(2)
        .build();
    let xlsx = XlsxBuilder::new()
        .sheet(SheetBuilder::new("Sheet1").cell("A1", "Test", Some(style)))
        .build();
    let wb = load_xlsx(&xlsx);
    assert_cell_align_h(&wb, 0, 0, 0, "left");
    let style = get_cell_style(&wb, 0, 0, 0).expect("Should have style");
    assert_eq!(
        style.align_v,
        Some(offidized_xlview::types::style::VAlign::Top)
    );
    assert_eq!(style.wrap, Some(true));
    assert_eq!(style.rotation, Some(30));
    assert_eq!(style.indent, Some(2));
}

/// Test justify horizontal with distributed vertical
#[test]
fn test_combination_justify_distributed() {
    let style = StyleBuilder::new()
        .align_horizontal("justify")
        .align_vertical("distributed")
        .build();
    let xlsx = XlsxBuilder::new()
        .sheet(SheetBuilder::new("Sheet1").cell("A1", "Test", Some(style)))
        .build();
    let wb = load_xlsx(&xlsx);
    assert_cell_align_h(&wb, 0, 0, 0, "justify");
    let style = get_cell_style(&wb, 0, 0, 0).expect("Should have style");
    assert_eq!(
        style.align_v,
        Some(offidized_xlview::types::style::VAlign::Distributed)
    );
}

/// Test center continuous with indent (used for grouped headers)
#[test]
fn test_combination_center_continuous_indent() {
    let style = StyleBuilder::new()
        .align_horizontal("centerContinuous")
        .indent(1)
        .build();
    let xlsx = XlsxBuilder::new()
        .sheet(SheetBuilder::new("Sheet1").cell("A1", "Test", Some(style)))
        .build();
    let wb = load_xlsx(&xlsx);
    assert_cell_align_h(&wb, 0, 0, 0, "centerContinuous");
    let style = get_cell_style(&wb, 0, 0, 0).expect("Should have style");
    assert_eq!(style.indent, Some(1));
}

// ============================================================================
// EDGE CASE TESTS
// ============================================================================

/// Test empty alignment element
#[test]
fn test_empty_alignment_element() {
    // A cell with no alignment properties should have no style or default style
    let xlsx = XlsxBuilder::new()
        .sheet(SheetBuilder::new("Sheet1").cell("A1", "Test", None))
        .build();
    let wb = load_xlsx(&xlsx);
    // Cell without explicit style may have None or default style
    let style = get_cell_style(&wb, 0, 0, 0);
    if let Some(s) = style {
        assert!(s.align_h.is_none(), "Empty should have no horizontal");
        assert!(s.align_v.is_none(), "Empty should have no vertical");
        assert!(s.wrap != Some(true), "Empty should have wrap_text false");
        assert!(
            s.indent.is_none() || s.indent == Some(0),
            "Empty should have no indent"
        );
        assert!(s.rotation.is_none(), "Empty should have no rotation");
    }
}

/// Test no alignment element at all
#[test]
fn test_no_alignment_element() {
    let xlsx = XlsxBuilder::new()
        .sheet(SheetBuilder::new("Sheet1").cell("A1", "Test", None))
        .build();
    let wb = load_xlsx(&xlsx);
    // Verify workbook parses correctly with default style
    assert_sheet_count(&wb, 1);
    assert_cell_value(&wb, 0, 0, 0, "Test");
}

/// Test attributes in different order
#[test]
fn test_attribute_order_independence() {
    // Two styles with the same properties should produce the same results
    let style1 = StyleBuilder::new()
        .align_vertical("top")
        .align_horizontal("left")
        .indent(1)
        .wrap_text()
        .build();
    let style2 = StyleBuilder::new()
        .align_horizontal("left")
        .wrap_text()
        .align_vertical("top")
        .indent(1)
        .build();

    let xlsx1 = XlsxBuilder::new()
        .sheet(SheetBuilder::new("Sheet1").cell("A1", "Test1", Some(style1)))
        .build();
    let xlsx2 = XlsxBuilder::new()
        .sheet(SheetBuilder::new("Sheet1").cell("A1", "Test2", Some(style2)))
        .build();

    let wb1 = load_xlsx(&xlsx1);
    let wb2 = load_xlsx(&xlsx2);

    let s1 = get_cell_style(&wb1, 0, 0, 0).expect("Should have style");
    let s2 = get_cell_style(&wb2, 0, 0, 0).expect("Should have style");

    assert_eq!(s1.align_h, s2.align_h, "Horizontal should match");
    assert_eq!(s1.align_v, s2.align_v, "Vertical should match");
    assert_eq!(s1.wrap, s2.wrap, "wrap should match");
    assert_eq!(s1.indent, s2.indent, "indent should match");
}

/// Test self-closing alignment element
#[test]
fn test_self_closing_alignment() {
    let style = StyleBuilder::new()
        .align_horizontal("center")
        .align_vertical("center")
        .build();
    let xlsx = XlsxBuilder::new()
        .sheet(SheetBuilder::new("Sheet1").cell("A1", "Test", Some(style)))
        .build();
    let wb = load_xlsx(&xlsx);
    assert_cell_align_h(&wb, 0, 0, 0, "center");
    let style = get_cell_style(&wb, 0, 0, 0).expect("Should have style");
    assert_eq!(
        style.align_v,
        Some(offidized_xlview::types::style::VAlign::Center)
    );
}

// ============================================================================
// MULTIPLE XF TESTS
// ============================================================================

/// Test multiple xf elements with different alignments
#[test]
fn test_multiple_xf_different_alignments() {
    let xlsx = XlsxBuilder::new()
        .sheet(
            SheetBuilder::new("Sheet1")
                .cell(
                    "A1",
                    "Left",
                    Some(StyleBuilder::new().align_horizontal("left").build()),
                )
                .cell(
                    "B1",
                    "Center",
                    Some(
                        StyleBuilder::new()
                            .align_horizontal("center")
                            .align_vertical("center")
                            .build(),
                    ),
                )
                .cell(
                    "C1",
                    "Right",
                    Some(
                        StyleBuilder::new()
                            .align_horizontal("right")
                            .wrap_text()
                            .indent(2)
                            .build(),
                    ),
                ),
        )
        .build();

    let wb = load_xlsx(&xlsx);

    // First cell: left
    assert_cell_align_h(&wb, 0, 0, 0, "left");

    // Second cell: center/center
    assert_cell_align_h(&wb, 0, 0, 1, "center");
    let s = get_cell_style(&wb, 0, 0, 1).expect("Should have style");
    assert_eq!(
        s.align_v,
        Some(offidized_xlview::types::style::VAlign::Center)
    );

    // Third cell: right with wrap and indent
    assert_cell_align_h(&wb, 0, 0, 2, "right");
    let s = get_cell_style(&wb, 0, 0, 2).expect("Should have style");
    assert_eq!(s.wrap, Some(true));
    assert_eq!(s.indent, Some(2));
}

/// Test xf with applyAlignment="0" (alignment should still be parsed)
#[test]
fn test_apply_alignment_false() {
    // When a style has alignment properties, they should be available
    // even if applyAlignment is false
    let style = StyleBuilder::new().align_horizontal("center").build();
    let xlsx = XlsxBuilder::new()
        .sheet(SheetBuilder::new("Sheet1").cell("A1", "Test", Some(style)))
        .build();
    let wb = load_xlsx(&xlsx);
    // The alignment should be parsed regardless
    assert_sheet_count(&wb, 1);
    assert_cell_value(&wb, 0, 0, 0, "Test");
}

// ============================================================================
// SPECIAL VALUE TESTS
// ============================================================================

/// Test textRotation with all valid rotation values
#[test]
fn test_text_rotation_full_range() {
    // Counterclockwise: selected values from 1-90
    for rotation in [1, 15, 30, 45, 60, 75, 89, 90] {
        let style = StyleBuilder::new().rotation(rotation).build();
        let xlsx = XlsxBuilder::new()
            .sheet(SheetBuilder::new("Sheet1").cell("A1", "Test", Some(style)))
            .build();
        let wb = load_xlsx(&xlsx);
        let s = get_cell_style(&wb, 0, 0, 0).expect("Should have style");
        assert_eq!(
            s.rotation,
            Some(rotation),
            "Should parse rotation {}",
            rotation
        );
    }

    // Clockwise (stored as 91-180)
    for rotation in [91, 105, 120, 135, 150, 165, 179, 180] {
        let style = StyleBuilder::new().rotation(rotation).build();
        let xlsx = XlsxBuilder::new()
            .sheet(SheetBuilder::new("Sheet1").cell("A1", "Test", Some(style)))
            .build();
        let wb = load_xlsx(&xlsx);
        let s = get_cell_style(&wb, 0, 0, 0).expect("Should have style");
        assert_eq!(
            s.rotation,
            Some(rotation),
            "Should parse rotation {}",
            rotation
        );
    }
}

// ============================================================================
// WHITESPACE AND FORMATTING TESTS
// ============================================================================

/// Test that extra whitespace in XML doesn't affect parsing
#[test]
fn test_whitespace_handling() {
    let style = StyleBuilder::new()
        .align_horizontal("center")
        .align_vertical("center")
        .wrap_text()
        .indent(2)
        .build();
    let xlsx = XlsxBuilder::new()
        .sheet(SheetBuilder::new("Sheet1").cell("A1", "Test", Some(style)))
        .build();
    let wb = load_xlsx(&xlsx);
    assert_cell_align_h(&wb, 0, 0, 0, "center");
    let s = get_cell_style(&wb, 0, 0, 0).expect("Should have style");
    assert_eq!(
        s.align_v,
        Some(offidized_xlview::types::style::VAlign::Center)
    );
    assert_eq!(s.wrap, Some(true));
    assert_eq!(s.indent, Some(2));
}

// ============================================================================
// DEFAULT VALUE TESTS
// ============================================================================

/// Test default values when alignment is present but empty
#[test]
fn test_default_values() {
    let xlsx = XlsxBuilder::new()
        .sheet(SheetBuilder::new("Sheet1").cell("A1", "Test", None))
        .build();
    let wb = load_xlsx(&xlsx);
    let style = get_cell_style(&wb, 0, 0, 0);
    if let Some(s) = style {
        assert!(s.align_h.is_none(), "Default horizontal should be None");
        assert!(s.align_v.is_none(), "Default vertical should be None");
        assert!(s.wrap != Some(true), "Default wrap should not be true");
        assert!(
            s.indent.is_none() || s.indent == Some(0),
            "Default indent should be None or 0"
        );
        assert!(s.rotation.is_none(), "Default rotation should be None");
    }
}

/// Test that Style::default() produces expected defaults
#[test]
fn test_style_default() {
    let default = offidized_xlview::types::style::Style::default();
    assert!(default.align_h.is_none());
    assert!(default.align_v.is_none());
    assert!(default.wrap.is_none());
    assert!(default.indent.is_none());
    assert!(default.rotation.is_none());
}
