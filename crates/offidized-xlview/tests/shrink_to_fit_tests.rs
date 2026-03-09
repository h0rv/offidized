//! Tests for shrink_to_fit alignment styling
//!
//! Shrink to fit is an Excel alignment feature that reduces the font size
//! to fit content within a cell without wrapping. When shrinkToFit="1" is
//! specified in the alignment element, the text is scaled down to fit.
//!
//! According to ECMA-376:
//! - shrinkToFit is mutually exclusive with wrapText
//! - When both are set, wrapText takes precedence
//! - The attribute is stored in xl/styles.xml in <alignment> elements
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
// Tests: Shrink to Fit Basic Parsing
// =============================================================================

#[test]
fn test_shrink_to_fit_enabled() {
    let style = StyleBuilder::new()
        .shrink_to_fit()
        .align_horizontal("center")
        .build();
    let xlsx = XlsxBuilder::new()
        .sheet(SheetBuilder::new("Sheet1").cell(
            "A1",
            "This is long text that needs shrinking",
            Some(style),
        ))
        .build();
    let wb = load_xlsx(&xlsx);

    let style = get_cell_style(&wb, 0, 0, 0).expect("Cell should have style");
    assert_eq!(
        style.shrink_to_fit,
        Some(true),
        "Cell A1 should have shrinkToFit=true"
    );
}

#[test]
fn test_shrink_to_fit_disabled() {
    let style = StyleBuilder::new().align_horizontal("center").build();
    let xlsx = XlsxBuilder::new()
        .sheet(SheetBuilder::new("Sheet1").cell("A1", "Text", Some(style)))
        .build();
    let wb = load_xlsx(&xlsx);

    let style = get_cell_style(&wb, 0, 0, 0).expect("Cell should have style");
    // When shrinkToFit is not set, it should be None or Some(false)
    if let Some(shrink) = style.shrink_to_fit {
        assert!(!shrink, "Cell A1 should not have shrinkToFit=true");
    }
}

#[test]
fn test_wrap_text_takes_precedence_over_shrink_to_fit() {
    // When both wrapText and shrinkToFit are set, wrapText takes precedence
    // according to ECMA-376
    let style = StyleBuilder::new()
        .shrink_to_fit()
        .wrap_text()
        .align_horizontal("center")
        .build();
    let xlsx = XlsxBuilder::new()
        .sheet(SheetBuilder::new("Sheet1").cell("A1", "Text", Some(style)))
        .build();
    let wb = load_xlsx(&xlsx);

    let style = get_cell_style(&wb, 0, 0, 0).expect("Cell should have style");
    // Both values should be parsed and stored
    assert_eq!(style.wrap, Some(true), "wrapText should be true");
    assert_eq!(
        style.shrink_to_fit,
        Some(true),
        "shrinkToFit should be true"
    );
    // Note: The renderer should handle the precedence - the parser stores both values
}

#[test]
fn test_cell_without_shrink_to_fit() {
    let xlsx = XlsxBuilder::new()
        .sheet(SheetBuilder::new("Sheet1").cell("A1", 42.0, None))
        .build();
    let wb = load_xlsx(&xlsx);

    // Cell without explicit style should not have shrink_to_fit
    let style = get_cell_style(&wb, 0, 0, 0);
    if let Some(s) = style {
        if let Some(shrink) = s.shrink_to_fit {
            assert!(!shrink, "Default style should not have shrinkToFit enabled");
        }
    }
}

// =============================================================================
// Tests: Shrink to Fit with Various Alignments
// =============================================================================

#[test]
fn test_shrink_to_fit_with_left_align() {
    let style = StyleBuilder::new()
        .shrink_to_fit()
        .align_horizontal("left")
        .build();
    let xlsx = XlsxBuilder::new()
        .sheet(SheetBuilder::new("Sheet1").cell("A1", "Text", Some(style)))
        .build();
    let wb = load_xlsx(&xlsx);

    let style = get_cell_style(&wb, 0, 0, 0).expect("Cell should have style");
    assert_eq!(style.shrink_to_fit, Some(true));
    assert!(
        style.align_h.is_some(),
        "Horizontal alignment should be set"
    );
}

#[test]
fn test_shrink_to_fit_with_center_align() {
    let style = StyleBuilder::new()
        .shrink_to_fit()
        .align_horizontal("center")
        .build();
    let xlsx = XlsxBuilder::new()
        .sheet(SheetBuilder::new("Sheet1").cell("A1", "Text", Some(style)))
        .build();
    let wb = load_xlsx(&xlsx);

    let style = get_cell_style(&wb, 0, 0, 0).expect("Cell should have style");
    assert_eq!(style.shrink_to_fit, Some(true));
}

#[test]
fn test_shrink_to_fit_with_right_align() {
    let style = StyleBuilder::new()
        .shrink_to_fit()
        .align_horizontal("right")
        .build();
    let xlsx = XlsxBuilder::new()
        .sheet(SheetBuilder::new("Sheet1").cell("A1", "Text", Some(style)))
        .build();
    let wb = load_xlsx(&xlsx);

    let style = get_cell_style(&wb, 0, 0, 0).expect("Cell should have style");
    assert_eq!(style.shrink_to_fit, Some(true));
}

// =============================================================================
// Tests: Real XLSX File Parsing
// =============================================================================

#[test]
fn test_kitchen_sink_v2_shrink_to_fit_cells() {
    let path = format!("{}/test/kitchen_sink_v2.xlsx", env!("CARGO_MANIFEST_DIR"));
    if !std::path::Path::new(&path).exists() {
        println!("Skipping test: {} not found", path);
        return;
    }

    let data = std::fs::read(&path).expect("Failed to read test file");
    let wb = load_xlsx(&data);

    let mut shrink_to_fit_count = 0;

    for sheet in &wb.sheets {
        for cell_data in &sheet.cells {
            if let Some(ref style) = cell_data.cell.s {
                if style.shrink_to_fit == Some(true) {
                    shrink_to_fit_count += 1;
                    println!(
                        "Found shrinkToFit cell in sheet '{}' at ({}, {})",
                        sheet.name, cell_data.r, cell_data.c
                    );
                }
            }
        }
    }

    println!(
        "kitchen_sink_v2.xlsx has {} cells with shrinkToFit=true",
        shrink_to_fit_count
    );
}

#[test]
fn test_ms_cf_samples_shrink_to_fit_cells() {
    let path = format!("{}/test/ms_cf_samples.xlsx", env!("CARGO_MANIFEST_DIR"));
    if !std::path::Path::new(&path).exists() {
        println!("Skipping test: {} not found", path);
        return;
    }

    let data = std::fs::read(&path).expect("Failed to read test file");
    let wb = load_xlsx(&data);

    let mut shrink_to_fit_count = 0;

    for sheet in &wb.sheets {
        for cell_data in &sheet.cells {
            if let Some(ref style) = cell_data.cell.s {
                if style.shrink_to_fit == Some(true) {
                    shrink_to_fit_count += 1;
                    println!(
                        "Found shrinkToFit cell in sheet '{}' at ({}, {})",
                        sheet.name, cell_data.r, cell_data.c
                    );
                }
            }
        }
    }

    println!(
        "ms_cf_samples.xlsx has {} cells with shrinkToFit=true",
        shrink_to_fit_count
    );
}

// =============================================================================
// Tests: Edge Cases
// =============================================================================

#[test]
fn test_shrink_to_fit_value_zero() {
    // When shrinkToFit is not enabled, it should be None or false
    let style = StyleBuilder::new().align_horizontal("left").build();
    let xlsx = XlsxBuilder::new()
        .sheet(SheetBuilder::new("Sheet1").cell("A1", "Test", Some(style)))
        .build();
    let wb = load_xlsx(&xlsx);

    let style = get_cell_style(&wb, 0, 0, 0).expect("Cell should have style");
    // shrinkToFit should be None or false when not set
    if let Some(shrink) = style.shrink_to_fit {
        assert!(!shrink, "shrinkToFit should be false when not enabled");
    }
}
