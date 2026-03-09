//! Integration tests for text rotation in XLSX files.
//!
//! These tests verify that cells with rotation styles have the rotation value
//! correctly parsed and available through the adapter pipeline.
//!
//! ## Excel Rotation Values
//!
//! The XLSX format uses the following rotation encoding:
//! - 0 = no rotation (horizontal text)
//! - 1-90 = counterclockwise rotation (1 degree to 90 degrees)
//! - 91-180 = clockwise rotation (maps to -1 to -90 degrees in display)
//!   - 91 = 1 degree clockwise (-1 display)
//!   - 135 = 45 degrees clockwise (-45 display)
//!   - 180 = 90 degrees clockwise (-90 display)
//! - 255 = vertical stacked text (characters stacked vertically)
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
// Test 1: Cell with 0 rotation (horizontal text)
// =============================================================================

/// Test that a cell with textRotation="0" is parsed correctly as no rotation.
///
/// This is the default horizontal text orientation. A rotation of 0 means
/// the text is displayed normally without any angle.
#[test]
fn test_cell_with_0_rotation_horizontal() {
    let style = StyleBuilder::new().rotation(0).build();
    let xlsx = XlsxBuilder::new()
        .sheet(SheetBuilder::new("Sheet1").cell("A1", "Horizontal", Some(style)))
        .build();

    let wb = load_xlsx(&xlsx);

    assert_sheet_count(&wb, 1);
    let cd = get_cell(&wb, 0, 0, 0);
    assert!(cd.is_some(), "Should have cells");

    let style = get_cell_style(&wb, 0, 0, 0).expect("Cell should have a style");
    assert_eq!(
        style.rotation,
        Some(0),
        "Rotation should be 0 for horizontal text"
    );
}

// =============================================================================
// Test 2: Cell with 45 degree rotation (counterclockwise)
// =============================================================================

/// Test that a cell with textRotation="45" is parsed correctly.
///
/// A rotation of 45 means the text is rotated 45 degrees counterclockwise.
/// This is a common rotation used for angled column headers.
#[test]
fn test_cell_with_45_degree_rotation() {
    let style = StyleBuilder::new().rotation(45).build();
    let xlsx = XlsxBuilder::new()
        .sheet(SheetBuilder::new("Sheet1").cell("A1", "Rotated", Some(style)))
        .build();

    let wb = load_xlsx(&xlsx);

    assert_sheet_count(&wb, 1);
    let style = get_cell_style(&wb, 0, 0, 0).expect("Cell should have a style");
    assert_eq!(
        style.rotation,
        Some(45),
        "Rotation should be 45 degrees counterclockwise"
    );
}

// =============================================================================
// Test 3: Cell with 90 degree rotation (vertical, bottom to top)
// =============================================================================

/// Test that a cell with textRotation="90" is parsed correctly.
///
/// A rotation of 90 means the text is rotated 90 degrees counterclockwise,
/// resulting in vertical text reading from bottom to top.
#[test]
fn test_cell_with_90_degree_rotation_vertical() {
    let style = StyleBuilder::new().rotation(90).build();
    let xlsx = XlsxBuilder::new()
        .sheet(SheetBuilder::new("Sheet1").cell("A1", "Vertical", Some(style)))
        .build();

    let wb = load_xlsx(&xlsx);

    assert_sheet_count(&wb, 1);
    let style = get_cell_style(&wb, 0, 0, 0).expect("Cell should have a style");
    assert_eq!(
        style.rotation,
        Some(90),
        "Rotation should be 90 degrees (vertical, bottom to top)"
    );
}

// =============================================================================
// Test 4: Cell with 180 rotation (90 degrees clockwise, top to bottom)
// =============================================================================

/// Test that a cell with textRotation="180" is parsed correctly.
///
/// In Excel's encoding, values 91-180 represent clockwise rotation:
/// - 180 = 90 degrees clockwise (or -90 degrees)
/// - This results in vertical text reading from top to bottom
#[test]
fn test_cell_with_180_rotation_clockwise() {
    let style = StyleBuilder::new().rotation(180).build();
    let xlsx = XlsxBuilder::new()
        .sheet(SheetBuilder::new("Sheet1").cell("A1", "Clockwise", Some(style)))
        .build();

    let wb = load_xlsx(&xlsx);

    assert_sheet_count(&wb, 1);
    let style = get_cell_style(&wb, 0, 0, 0).expect("Cell should have a style");
    assert_eq!(
        style.rotation,
        Some(180),
        "Rotation should be 180 (90 degrees clockwise)"
    );
}

// =============================================================================
// Test 5: Cell with 255 rotation (vertical stacked text)
// =============================================================================

/// Test that a cell with textRotation="255" is parsed correctly.
///
/// The special value 255 indicates vertical stacked text, where each character
/// is displayed horizontally but stacked vertically (one character per line).
/// This is different from rotated text where characters are tilted.
#[test]
fn test_cell_with_255_rotation_vertical_stacked() {
    let style = StyleBuilder::new().rotation(255).build();
    let xlsx = XlsxBuilder::new()
        .sheet(SheetBuilder::new("Sheet1").cell("A1", "Stacked", Some(style)))
        .build();

    let wb = load_xlsx(&xlsx);

    assert_sheet_count(&wb, 1);
    let style = get_cell_style(&wb, 0, 0, 0).expect("Cell should have a style");
    assert_eq!(
        style.rotation,
        Some(255),
        "Rotation should be 255 for vertical stacked text"
    );
}
