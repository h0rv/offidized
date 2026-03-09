//! Comprehensive tests for all 19 pattern fill types, ported from xlview.
//!
//! Tests that each ECMA-376 pattern fill type can be created via StyleBuilder,
//! written to a valid XLSX file, and parsed correctly through the adapter pipeline.
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
use offidized_xlview::types::style::PatternType;

// ============================================================================
// Helper Functions
// ============================================================================

fn expected_pattern_type(pattern_str: &str) -> PatternType {
    match pattern_str {
        "none" => PatternType::None,
        "solid" => PatternType::Solid,
        "gray125" => PatternType::Gray125,
        "gray0625" => PatternType::Gray0625,
        "darkGray" => PatternType::DarkGray,
        "mediumGray" => PatternType::MediumGray,
        "lightGray" => PatternType::LightGray,
        "darkHorizontal" => PatternType::DarkHorizontal,
        "darkVertical" => PatternType::DarkVertical,
        "darkDown" => PatternType::DarkDown,
        "darkUp" => PatternType::DarkUp,
        "darkGrid" => PatternType::DarkGrid,
        "darkTrellis" => PatternType::DarkTrellis,
        "lightHorizontal" => PatternType::LightHorizontal,
        "lightVertical" => PatternType::LightVertical,
        "lightDown" => PatternType::LightDown,
        "lightUp" => PatternType::LightUp,
        "lightGrid" => PatternType::LightGrid,
        "lightTrellis" => PatternType::LightTrellis,
        _ => panic!("Unknown pattern type: {}", pattern_str),
    }
}

fn create_xlsx_with_pattern(pattern_type: &str) -> Vec<u8> {
    XlsxBuilder::new()
        .sheet(SheetBuilder::new("Sheet1").cell(
            "A1",
            format!("Pattern: {}", pattern_type),
            Some(StyleBuilder::new().pattern(pattern_type).build()),
        ))
        .build()
}

// ============================================================================
// Test: All Pattern Types Parse Correctly
// ============================================================================

#[test]
fn test_all_pattern_types_via_fixture() {
    assert_eq!(
        ALL_PATTERN_FILLS.len(),
        19,
        "ALL_PATTERN_FILLS should have exactly 19 patterns"
    );

    for pattern_type in ALL_PATTERN_FILLS {
        let xlsx = create_xlsx_with_pattern(pattern_type);
        let workbook = load_xlsx(&xlsx);

        assert_eq!(
            workbook.sheets.len(),
            1,
            "Should have exactly one sheet for pattern '{}'",
            pattern_type
        );

        let sheet = &workbook.sheets[0];
        let cell = sheet
            .cells
            .iter()
            .find(|c| c.r == 0 && c.c == 0)
            .unwrap_or_else(|| panic!("Cell A1 should exist for pattern '{}'", pattern_type));

        let style = get_cell_style(&workbook, 0, 0, 0);
        assert!(
            style.is_some() || cell.cell.s.is_some(),
            "Cell should have style for pattern '{}'",
            pattern_type
        );

        let style_ref = cell.cell.s.as_ref().or(style);

        match *pattern_type {
            "none" => {
                if let Some(s) = style_ref {
                    if let Some(ref pt) = s.pattern_type {
                        assert_eq!(*pt, PatternType::None, "Pattern type mismatch for 'none'");
                    }
                }
            }
            "solid" => {
                if let Some(s) = style_ref {
                    assert!(
                        s.pattern_type.is_none(),
                        "Solid pattern should NOT set pattern_type (parser sets bg_color instead)"
                    );
                }
            }
            _ => {
                let s = style_ref.unwrap_or_else(|| {
                    panic!("Style should be present for pattern '{}'", pattern_type)
                });
                assert!(
                    s.pattern_type.is_some(),
                    "pattern_type should be set for pattern '{}', got None",
                    pattern_type
                );
                let expected = expected_pattern_type(pattern_type);
                let actual = s.pattern_type.as_ref().unwrap();
                assert_eq!(
                    *actual, expected,
                    "Pattern type mismatch for '{}': expected {:?}, got {:?}",
                    pattern_type, expected, actual
                );
            }
        }
    }
}

// ============================================================================
// Individual Pattern Type Tests
// ============================================================================

#[test]
fn test_pattern_none() {
    let xlsx = create_xlsx_with_pattern("none");
    let workbook = load_xlsx(&xlsx);
    let cell = get_cell(&workbook, 0, 0, 0);
    assert!(cell.is_some(), "Cell A1 should exist");
}

#[test]
fn test_pattern_solid() {
    let xlsx = create_xlsx_with_pattern("solid");
    let workbook = load_xlsx(&xlsx);
    let cell = get_cell(&workbook, 0, 0, 0);
    assert!(cell.is_some(), "Cell A1 should exist");
    let cd = cell.unwrap();
    if let Some(ref style) = cd.cell.s {
        assert!(
            style.pattern_type.is_none(),
            "Solid fills should NOT set pattern_type (parser uses bg_color instead)"
        );
    }
}

#[test]
fn test_pattern_gray125() {
    let xlsx = create_xlsx_with_pattern("gray125");
    let workbook = load_xlsx(&xlsx);
    let cd = get_cell(&workbook, 0, 0, 0).expect("Cell A1 should exist");
    let style = cd
        .cell
        .s
        .as_ref()
        .or_else(|| get_cell_style(&workbook, 0, 0, 0));
    let style = style.expect("Cell should have style");
    assert_eq!(
        style.pattern_type,
        Some(PatternType::Gray125),
        "Pattern should be gray125"
    );
}

#[test]
fn test_pattern_gray0625() {
    let xlsx = create_xlsx_with_pattern("gray0625");
    let workbook = load_xlsx(&xlsx);
    let cd = get_cell(&workbook, 0, 0, 0).expect("Cell A1 should exist");
    let style = cd
        .cell
        .s
        .as_ref()
        .or_else(|| get_cell_style(&workbook, 0, 0, 0));
    let style = style.expect("Cell should have style");
    assert_eq!(
        style.pattern_type,
        Some(PatternType::Gray0625),
        "Pattern should be gray0625"
    );
}

#[test]
fn test_pattern_dark_gray() {
    let xlsx = create_xlsx_with_pattern("darkGray");
    let workbook = load_xlsx(&xlsx);
    let cd = get_cell(&workbook, 0, 0, 0).expect("Cell A1 should exist");
    let style = cd
        .cell
        .s
        .as_ref()
        .or_else(|| get_cell_style(&workbook, 0, 0, 0));
    let style = style.expect("Cell should have style");
    assert_eq!(style.pattern_type, Some(PatternType::DarkGray));
}

#[test]
fn test_pattern_medium_gray() {
    let xlsx = create_xlsx_with_pattern("mediumGray");
    let workbook = load_xlsx(&xlsx);
    let cd = get_cell(&workbook, 0, 0, 0).expect("Cell A1 should exist");
    let style = cd
        .cell
        .s
        .as_ref()
        .or_else(|| get_cell_style(&workbook, 0, 0, 0));
    let style = style.expect("Cell should have style");
    assert_eq!(style.pattern_type, Some(PatternType::MediumGray));
}

#[test]
fn test_pattern_light_gray() {
    let xlsx = create_xlsx_with_pattern("lightGray");
    let workbook = load_xlsx(&xlsx);
    let cd = get_cell(&workbook, 0, 0, 0).expect("Cell A1 should exist");
    let style = cd
        .cell
        .s
        .as_ref()
        .or_else(|| get_cell_style(&workbook, 0, 0, 0));
    let style = style.expect("Cell should have style");
    assert_eq!(style.pattern_type, Some(PatternType::LightGray));
}

#[test]
fn test_pattern_dark_horizontal() {
    let xlsx = create_xlsx_with_pattern("darkHorizontal");
    let workbook = load_xlsx(&xlsx);
    let cd = get_cell(&workbook, 0, 0, 0).expect("Cell A1 should exist");
    let style = cd
        .cell
        .s
        .as_ref()
        .or_else(|| get_cell_style(&workbook, 0, 0, 0));
    let style = style.expect("Cell should have style");
    assert_eq!(style.pattern_type, Some(PatternType::DarkHorizontal));
}

#[test]
fn test_pattern_dark_vertical() {
    let xlsx = create_xlsx_with_pattern("darkVertical");
    let workbook = load_xlsx(&xlsx);
    let cd = get_cell(&workbook, 0, 0, 0).expect("Cell A1 should exist");
    let style = cd
        .cell
        .s
        .as_ref()
        .or_else(|| get_cell_style(&workbook, 0, 0, 0));
    let style = style.expect("Cell should have style");
    assert_eq!(style.pattern_type, Some(PatternType::DarkVertical));
}

#[test]
fn test_pattern_dark_down() {
    let xlsx = create_xlsx_with_pattern("darkDown");
    let workbook = load_xlsx(&xlsx);
    let cd = get_cell(&workbook, 0, 0, 0).expect("Cell A1 should exist");
    let style = cd
        .cell
        .s
        .as_ref()
        .or_else(|| get_cell_style(&workbook, 0, 0, 0));
    let style = style.expect("Cell should have style");
    assert_eq!(style.pattern_type, Some(PatternType::DarkDown));
}

#[test]
fn test_pattern_dark_up() {
    let xlsx = create_xlsx_with_pattern("darkUp");
    let workbook = load_xlsx(&xlsx);
    let cd = get_cell(&workbook, 0, 0, 0).expect("Cell A1 should exist");
    let style = cd
        .cell
        .s
        .as_ref()
        .or_else(|| get_cell_style(&workbook, 0, 0, 0));
    let style = style.expect("Cell should have style");
    assert_eq!(style.pattern_type, Some(PatternType::DarkUp));
}

#[test]
fn test_pattern_dark_grid() {
    let xlsx = create_xlsx_with_pattern("darkGrid");
    let workbook = load_xlsx(&xlsx);
    let cd = get_cell(&workbook, 0, 0, 0).expect("Cell A1 should exist");
    let style = cd
        .cell
        .s
        .as_ref()
        .or_else(|| get_cell_style(&workbook, 0, 0, 0));
    let style = style.expect("Cell should have style");
    assert_eq!(style.pattern_type, Some(PatternType::DarkGrid));
}

#[test]
fn test_pattern_dark_trellis() {
    let xlsx = create_xlsx_with_pattern("darkTrellis");
    let workbook = load_xlsx(&xlsx);
    let cd = get_cell(&workbook, 0, 0, 0).expect("Cell A1 should exist");
    let style = cd
        .cell
        .s
        .as_ref()
        .or_else(|| get_cell_style(&workbook, 0, 0, 0));
    let style = style.expect("Cell should have style");
    assert_eq!(style.pattern_type, Some(PatternType::DarkTrellis));
}

#[test]
fn test_pattern_light_horizontal() {
    let xlsx = create_xlsx_with_pattern("lightHorizontal");
    let workbook = load_xlsx(&xlsx);
    let cd = get_cell(&workbook, 0, 0, 0).expect("Cell A1 should exist");
    let style = cd
        .cell
        .s
        .as_ref()
        .or_else(|| get_cell_style(&workbook, 0, 0, 0));
    let style = style.expect("Cell should have style");
    assert_eq!(style.pattern_type, Some(PatternType::LightHorizontal));
}

#[test]
fn test_pattern_light_vertical() {
    let xlsx = create_xlsx_with_pattern("lightVertical");
    let workbook = load_xlsx(&xlsx);
    let cd = get_cell(&workbook, 0, 0, 0).expect("Cell A1 should exist");
    let style = cd
        .cell
        .s
        .as_ref()
        .or_else(|| get_cell_style(&workbook, 0, 0, 0));
    let style = style.expect("Cell should have style");
    assert_eq!(style.pattern_type, Some(PatternType::LightVertical));
}

#[test]
fn test_pattern_light_down() {
    let xlsx = create_xlsx_with_pattern("lightDown");
    let workbook = load_xlsx(&xlsx);
    let cd = get_cell(&workbook, 0, 0, 0).expect("Cell A1 should exist");
    let style = cd
        .cell
        .s
        .as_ref()
        .or_else(|| get_cell_style(&workbook, 0, 0, 0));
    let style = style.expect("Cell should have style");
    assert_eq!(style.pattern_type, Some(PatternType::LightDown));
}

#[test]
fn test_pattern_light_up() {
    let xlsx = create_xlsx_with_pattern("lightUp");
    let workbook = load_xlsx(&xlsx);
    let cd = get_cell(&workbook, 0, 0, 0).expect("Cell A1 should exist");
    let style = cd
        .cell
        .s
        .as_ref()
        .or_else(|| get_cell_style(&workbook, 0, 0, 0));
    let style = style.expect("Cell should have style");
    assert_eq!(style.pattern_type, Some(PatternType::LightUp));
}

#[test]
fn test_pattern_light_grid() {
    let xlsx = create_xlsx_with_pattern("lightGrid");
    let workbook = load_xlsx(&xlsx);
    let cd = get_cell(&workbook, 0, 0, 0).expect("Cell A1 should exist");
    let style = cd
        .cell
        .s
        .as_ref()
        .or_else(|| get_cell_style(&workbook, 0, 0, 0));
    let style = style.expect("Cell should have style");
    assert_eq!(style.pattern_type, Some(PatternType::LightGrid));
}

#[test]
fn test_pattern_light_trellis() {
    let xlsx = create_xlsx_with_pattern("lightTrellis");
    let workbook = load_xlsx(&xlsx);
    let cd = get_cell(&workbook, 0, 0, 0).expect("Cell A1 should exist");
    let style = cd
        .cell
        .s
        .as_ref()
        .or_else(|| get_cell_style(&workbook, 0, 0, 0));
    let style = style.expect("Cell should have style");
    assert_eq!(style.pattern_type, Some(PatternType::LightTrellis));
}

// ============================================================================
// Pattern Fill with Colors Tests
// ============================================================================

#[test]
fn test_solid_fill_with_color() {
    let xlsx = XlsxBuilder::new()
        .sheet(SheetBuilder::new("Sheet1").cell(
            "A1",
            "Yellow Background",
            Some(StyleBuilder::new().bg_color("#FFFF00").build()),
        ))
        .build();

    let workbook = load_xlsx(&xlsx);
    let style = get_cell_style(&workbook, 0, 0, 0).expect("Cell should have style");

    assert!(
        style.pattern_type.is_none(),
        "Solid fills should NOT set pattern_type (parser optimizes to bg_color)"
    );
    assert!(
        style.bg_color.is_some(),
        "Cell should have background color"
    );
    let bg = style.bg_color.as_ref().unwrap();
    assert!(
        bg.contains("FFFF00") || bg.contains("ffff00"),
        "Background color should be yellow, got: {}",
        bg
    );
}

#[test]
fn test_gray125_pattern_parses_with_style() {
    let xlsx = create_xlsx_with_pattern("gray125");
    let workbook = load_xlsx(&xlsx);
    let cd = get_cell(&workbook, 0, 0, 0).expect("Cell A1 should exist");
    let style = cd
        .cell
        .s
        .as_ref()
        .or_else(|| get_cell_style(&workbook, 0, 0, 0));
    let style = style.expect("Cell should have style");
    assert_eq!(style.pattern_type, Some(PatternType::Gray125));
}

// ============================================================================
// Multiple Patterns in Same Workbook
// ============================================================================

#[test]
fn test_multiple_patterns_in_single_sheet() {
    let xlsx = XlsxBuilder::new()
        .sheet(
            SheetBuilder::new("Sheet1")
                .cell(
                    "A1",
                    "Solid",
                    Some(StyleBuilder::new().pattern("solid").build()),
                )
                .cell(
                    "A2",
                    "Gray125",
                    Some(StyleBuilder::new().pattern("gray125").build()),
                )
                .cell(
                    "A3",
                    "DarkGray",
                    Some(StyleBuilder::new().pattern("darkGray").build()),
                )
                .cell(
                    "A4",
                    "LightHorizontal",
                    Some(StyleBuilder::new().pattern("lightHorizontal").build()),
                ),
        )
        .build();

    let workbook = load_xlsx(&xlsx);

    // A1 - solid (parser doesn't set pattern_type for solid)
    let cd_a1 = get_cell(&workbook, 0, 0, 0);
    assert!(cd_a1.is_some());

    // A2 - gray125
    let cd_a2 = get_cell(&workbook, 0, 1, 0);
    assert!(cd_a2.is_some());
    let s2 = cd_a2
        .unwrap()
        .cell
        .s
        .as_ref()
        .or_else(|| get_cell_style(&workbook, 0, 1, 0));
    if let Some(s) = s2 {
        assert_eq!(s.pattern_type, Some(PatternType::Gray125));
    }

    // A3 - darkGray
    let cd_a3 = get_cell(&workbook, 0, 2, 0);
    assert!(cd_a3.is_some());
    let s3 = cd_a3
        .unwrap()
        .cell
        .s
        .as_ref()
        .or_else(|| get_cell_style(&workbook, 0, 2, 0));
    if let Some(s) = s3 {
        assert_eq!(s.pattern_type, Some(PatternType::DarkGray));
    }

    // A4 - lightHorizontal
    let cd_a4 = get_cell(&workbook, 0, 3, 0);
    assert!(cd_a4.is_some());
    let s4 = cd_a4
        .unwrap()
        .cell
        .s
        .as_ref()
        .or_else(|| get_cell_style(&workbook, 0, 3, 0));
    if let Some(s) = s4 {
        assert_eq!(s.pattern_type, Some(PatternType::LightHorizontal));
    }
}

#[test]
fn test_all_19_patterns_in_single_workbook() {
    let mut sheet = SheetBuilder::new("AllPatterns");
    for (i, pattern_type) in ALL_PATTERN_FILLS.iter().enumerate() {
        let cell_ref = format!("A{}", i + 1);
        sheet = sheet.cell(
            &cell_ref,
            *pattern_type,
            Some(StyleBuilder::new().pattern(pattern_type).build()),
        );
    }

    let xlsx = XlsxBuilder::new().sheet(sheet).build();
    let workbook = load_xlsx(&xlsx);

    let sheet = &workbook.sheets[0];
    assert!(
        sheet.cells.len() >= 19,
        "Should have at least 19 cells, got {}",
        sheet.cells.len()
    );
}

// ============================================================================
// Pattern Fill Count Verification
// ============================================================================

#[test]
fn test_all_pattern_fills_constant_has_19_entries() {
    assert_eq!(
        ALL_PATTERN_FILLS.len(),
        19,
        "ALL_PATTERN_FILLS should contain exactly 19 pattern types per ECMA-376"
    );
}

#[test]
fn test_pattern_fills_contain_expected_values() {
    let expected_patterns = [
        "none",
        "solid",
        "mediumGray",
        "darkGray",
        "lightGray",
        "darkHorizontal",
        "darkVertical",
        "darkDown",
        "darkUp",
        "darkGrid",
        "darkTrellis",
        "lightHorizontal",
        "lightVertical",
        "lightDown",
        "lightUp",
        "lightGrid",
        "lightTrellis",
        "gray125",
        "gray0625",
    ];

    for pattern in &expected_patterns {
        assert!(
            ALL_PATTERN_FILLS.contains(pattern),
            "ALL_PATTERN_FILLS should contain '{}'",
            pattern
        );
    }
}

// ============================================================================
// Edge Cases
// ============================================================================

#[test]
fn test_pattern_with_additional_styling() {
    let xlsx = XlsxBuilder::new()
        .sheet(
            SheetBuilder::new("Sheet1").cell(
                "A1",
                "Styled Pattern",
                Some(
                    StyleBuilder::new()
                        .pattern("darkGrid")
                        .bold()
                        .italic()
                        .font_size(14.0)
                        .font_color("#FF0000")
                        .build(),
                ),
            ),
        )
        .build();

    let workbook = load_xlsx(&xlsx);
    let cd = get_cell(&workbook, 0, 0, 0).expect("Cell A1 should exist");
    let style = cd
        .cell
        .s
        .as_ref()
        .or_else(|| get_cell_style(&workbook, 0, 0, 0));
    let style = style.expect("Cell should have style");

    assert_eq!(
        style.pattern_type,
        Some(PatternType::DarkGrid),
        "Pattern should be darkGrid"
    );
    assert_eq!(style.bold, Some(true), "Should be bold");
    assert_eq!(style.italic, Some(true), "Should be italic");
    assert_eq!(style.font_size, Some(14.0), "Font size should be 14");
    assert!(style.font_color.is_some(), "Should have font color");
}

#[test]
fn test_empty_cell_with_pattern_only() {
    let xlsx = XlsxBuilder::new()
        .sheet(SheetBuilder::new("Sheet1").cell(
            "A1",
            "",
            Some(StyleBuilder::new().pattern("lightTrellis").build()),
        ))
        .build();

    let workbook = load_xlsx(&xlsx);
    let cell = get_cell(&workbook, 0, 0, 0);
    if let Some(cd) = cell {
        if let Some(ref style) = cd.cell.s {
            assert_eq!(style.pattern_type, Some(PatternType::LightTrellis));
        }
    }
}
