//! Common test utilities and assertion helpers for offidized-xlview.
//!
//! This module provides helper functions for testing the viewer adapter pipeline.
//! Unlike xlview (which parsed to JSON), these helpers work directly with the
//! viewer's typed `Workbook`, `Sheet`, `Cell`, and `StyleRef` types.
#![allow(
    dead_code,
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

// Re-export fixtures for convenience
pub use super::fixtures::*;

use offidized_xlview::types::style::{HAlign, StyleRef};
use offidized_xlview::types::workbook::{Cell, CellData, Workbook};

// ============================================================================
// Workbook Loading Helper
// ============================================================================

/// Load XLSX bytes into the viewer's Workbook type.
///
/// This is the primary entry point for tests: it parses the bytes with
/// `offidized_xlsx::Workbook` and then converts to the viewer's `Workbook`
/// via the adapter.
///
/// Panics on parse failure (this is test code).
#[must_use]
pub fn load_xlsx(data: &[u8]) -> Workbook {
    let wb = offidized_xlsx::Workbook::from_bytes(data)
        .expect("Failed to parse XLSX bytes with offidized_xlsx");
    offidized_xlview::adapter::convert_workbook(&wb)
}

// ============================================================================
// Cell Lookup Helpers
// ============================================================================

/// Get a cell at the given (0-indexed) position from the workbook.
///
/// Returns `None` if the sheet or cell doesn't exist.
pub fn get_cell(workbook: &Workbook, sheet: usize, row: u32, col: u32) -> Option<&CellData> {
    let s = workbook.sheets.get(sheet)?;
    s.cells.iter().find(|cd| cd.r == row && cd.c == col)
}

/// Get the resolved style for a cell at the given position.
///
/// Returns `None` if the cell doesn't exist or has no style.
pub fn get_cell_style(workbook: &Workbook, sheet: usize, row: u32, col: u32) -> Option<&StyleRef> {
    let cd = get_cell(workbook, sheet, row, col)?;

    // First check for inline style override
    if let Some(ref s) = cd.cell.s {
        return Some(s);
    }

    // Then check resolved styles by index
    let style_idx = cd.cell.style_idx? as usize;
    workbook
        .resolved_styles
        .get(style_idx)
        .and_then(|opt| opt.as_ref())
}

/// Get the display value of a cell.
///
/// Resolves shared strings and raw values to a display string.
fn cell_display_value(workbook: &Workbook, cell: &Cell) -> Option<String> {
    // Check cached display first
    if let Some(ref display) = cell.cached_display {
        return Some(display.clone());
    }

    // Check explicit value
    if let Some(ref v) = cell.v {
        return Some(v.clone());
    }

    // Resolve from raw value
    use offidized_xlview::types::workbook::CellRawValue;
    match cell.raw.as_ref()? {
        CellRawValue::String(s) => Some(s.clone()),
        CellRawValue::Number(n) => Some(n.to_string()),
        CellRawValue::Boolean(b) => Some(if *b { "TRUE" } else { "FALSE" }.to_string()),
        CellRawValue::Error(e) => Some(e.clone()),
        CellRawValue::Date(n) => Some(n.to_string()),
        CellRawValue::SharedString(idx) => workbook.shared_strings.get(*idx as usize).cloned(),
    }
}

// ============================================================================
// Assertion Helpers
// ============================================================================

/// Assert that a cell exists at the given position with the expected display value.
pub fn assert_cell_value(workbook: &Workbook, sheet: usize, row: u32, col: u32, expected: &str) {
    let cd = get_cell(workbook, sheet, row, col).unwrap_or_else(|| {
        panic!(
            "Cell at row={}, col={} not found in sheet {}",
            row, col, sheet
        )
    });

    let value = cell_display_value(workbook, &cd.cell)
        .unwrap_or_else(|| panic!("Cell at row={}, col={} has no value", row, col));

    pretty_assertions::assert_eq!(
        value.as_str(),
        expected,
        "Cell value mismatch at sheet={}, row={}, col={}",
        sheet,
        row,
        col
    );
}

/// Assert that a cell has bold formatting.
pub fn assert_cell_bold(workbook: &Workbook, sheet: usize, row: u32, col: u32) {
    let style = get_cell_style(workbook, sheet, row, col)
        .unwrap_or_else(|| panic!("Cell at row={}, col={} has no style", row, col));

    assert!(
        style.bold == Some(true),
        "Cell at row={}, col={} is not bold (bold={:?})",
        row,
        col,
        style.bold
    );
}

/// Assert that a cell has italic formatting.
pub fn assert_cell_italic(workbook: &Workbook, sheet: usize, row: u32, col: u32) {
    let style = get_cell_style(workbook, sheet, row, col)
        .unwrap_or_else(|| panic!("Cell at row={}, col={} has no style", row, col));

    assert!(
        style.italic == Some(true),
        "Cell at row={}, col={} is not italic (italic={:?})",
        row,
        col,
        style.italic
    );
}

/// Assert that a cell has a specific font size.
pub fn assert_cell_font_size(workbook: &Workbook, sheet: usize, row: u32, col: u32, expected: f64) {
    let style = get_cell_style(workbook, sheet, row, col)
        .unwrap_or_else(|| panic!("Cell at row={}, col={} has no style", row, col));

    let size = style
        .font_size
        .unwrap_or_else(|| panic!("Cell at row={}, col={} has no font size", row, col));

    assert!(
        (size - expected).abs() < 0.01,
        "Cell font size mismatch at row={}, col={}: expected {}, got {}",
        row,
        col,
        expected,
        size
    );
}

/// Assert that a cell has a specific font color.
pub fn assert_cell_font_color(
    workbook: &Workbook,
    sheet: usize,
    row: u32,
    col: u32,
    expected: &str,
) {
    let style = get_cell_style(workbook, sheet, row, col)
        .unwrap_or_else(|| panic!("Cell at row={}, col={} has no style", row, col));

    let color = style
        .font_color
        .as_deref()
        .unwrap_or_else(|| panic!("Cell at row={}, col={} has no font color", row, col));

    let expected_normalized = normalize_color_for_compare(expected);
    let actual_normalized = normalize_color_for_compare(color);

    pretty_assertions::assert_eq!(
        actual_normalized,
        expected_normalized,
        "Cell font color mismatch at row={}, col={}: expected {}, got {}",
        row,
        col,
        expected,
        color
    );
}

/// Assert that a cell has a specific background color.
pub fn assert_cell_bg_color(workbook: &Workbook, sheet: usize, row: u32, col: u32, expected: &str) {
    let style = get_cell_style(workbook, sheet, row, col)
        .unwrap_or_else(|| panic!("Cell at row={}, col={} has no style", row, col));

    let color = style
        .bg_color
        .as_deref()
        .unwrap_or_else(|| panic!("Cell at row={}, col={} has no background color", row, col));

    let expected_normalized = normalize_color_for_compare(expected);
    let actual_normalized = normalize_color_for_compare(color);

    pretty_assertions::assert_eq!(
        actual_normalized,
        expected_normalized,
        "Cell background color mismatch at row={}, col={}: expected {}, got {}",
        row,
        col,
        expected,
        color
    );
}

/// Assert that a cell has a specific horizontal alignment.
pub fn assert_cell_align_h(workbook: &Workbook, sheet: usize, row: u32, col: u32, expected: &str) {
    let style = get_cell_style(workbook, sheet, row, col)
        .unwrap_or_else(|| panic!("Cell at row={}, col={} has no style", row, col));

    let align = style.align_h.as_ref().unwrap_or_else(|| {
        panic!(
            "Cell at row={}, col={} has no horizontal alignment",
            row, col
        )
    });

    let actual_str = halign_to_str(align);
    pretty_assertions::assert_eq!(
        actual_str,
        expected,
        "Cell horizontal alignment mismatch at row={}, col={}",
        row,
        col
    );
}

/// Assert that a cell has text wrapping enabled.
pub fn assert_cell_wrap(workbook: &Workbook, sheet: usize, row: u32, col: u32) {
    let style = get_cell_style(workbook, sheet, row, col)
        .unwrap_or_else(|| panic!("Cell at row={}, col={} has no style", row, col));

    assert!(
        style.wrap == Some(true),
        "Cell at row={}, col={} does not have text wrap (wrap={:?})",
        row,
        col,
        style.wrap
    );
}

/// Assert that a merge range exists in the given sheet.
pub fn assert_merge_exists(
    workbook: &Workbook,
    sheet: usize,
    start_row: u32,
    start_col: u32,
    end_row: u32,
    end_col: u32,
) {
    let s = workbook
        .sheets
        .get(sheet)
        .unwrap_or_else(|| panic!("Sheet {} not found", sheet));

    let found = s.merges.iter().any(|m| {
        m.start_row == start_row
            && m.start_col == start_col
            && m.end_row == end_row
            && m.end_col == end_col
    });

    assert!(
        found,
        "Merge range {}:{} to {}:{} not found in sheet {}. Existing merges: {:?}",
        start_row,
        start_col,
        end_row,
        end_col,
        sheet,
        s.merges
            .iter()
            .map(|m| format!(
                "({}:{} -> {}:{})",
                m.start_row, m.start_col, m.end_row, m.end_col
            ))
            .collect::<Vec<_>>()
    );
}

/// Assert that a specific number of sheets exist.
pub fn assert_sheet_count(workbook: &Workbook, expected: usize) {
    pretty_assertions::assert_eq!(
        workbook.sheets.len(),
        expected,
        "Sheet count mismatch: expected {}, got {}",
        expected,
        workbook.sheets.len()
    );
}

/// Assert that a sheet has the expected name.
pub fn assert_sheet_name(workbook: &Workbook, sheet: usize, expected: &str) {
    let s = workbook
        .sheets
        .get(sheet)
        .unwrap_or_else(|| panic!("Sheet {} not found", sheet));

    pretty_assertions::assert_eq!(
        s.name.as_str(),
        expected,
        "Sheet name mismatch at index {}",
        sheet
    );
}

// ============================================================================
// Internal Helper Functions
// ============================================================================

/// Normalize color for comparison (strip # prefix, uppercase).
fn normalize_color_for_compare(color: &str) -> String {
    color.trim_start_matches('#').to_uppercase()
}

/// Convert HAlign enum to lowercase string for comparison.
fn halign_to_str(h: &HAlign) -> &'static str {
    match h {
        HAlign::General => "general",
        HAlign::Left => "left",
        HAlign::Center => "center",
        HAlign::Right => "right",
        HAlign::Fill => "fill",
        HAlign::Justify => "justify",
        HAlign::CenterContinuous => "centerContinuous",
        HAlign::Distributed => "distributed",
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_load_minimal_xlsx() {
        let xlsx = minimal_xlsx();
        let workbook = load_xlsx(&xlsx);

        assert_sheet_count(&workbook, 1);
        assert_sheet_name(&workbook, 0, "Sheet1");
    }

    #[test]
    fn test_load_xlsx_with_text() {
        let xlsx = xlsx_with_text("Hello, World!");
        let workbook = load_xlsx(&xlsx);

        assert_cell_value(&workbook, 0, 0, 0, "Hello, World!");
    }

    #[test]
    fn test_load_xlsx_with_styled_cell() {
        let style = StyleBuilder::new()
            .bold()
            .font_size(14.0)
            .bg_color("#FFFF00")
            .build();

        let xlsx = xlsx_with_styled_cell("Styled", style);
        let workbook = load_xlsx(&xlsx);

        assert_cell_value(&workbook, 0, 0, 0, "Styled");
        assert_cell_bold(&workbook, 0, 0, 0);
        assert_cell_font_size(&workbook, 0, 0, 0, 14.0);
        assert_cell_bg_color(&workbook, 0, 0, 0, "#FFFF00");
    }

    #[test]
    fn test_multiple_sheets() {
        let xlsx = XlsxBuilder::new()
            .sheet(
                SheetBuilder::new("Data")
                    .cell("A1", "First", None)
                    .cell("B1", 42.0, None),
            )
            .sheet(SheetBuilder::new("Summary").cell("A1", "Total", None))
            .build();

        let workbook = load_xlsx(&xlsx);
        assert_sheet_count(&workbook, 2);
        assert_sheet_name(&workbook, 0, "Data");
        assert_sheet_name(&workbook, 1, "Summary");
        assert_cell_value(&workbook, 0, 0, 0, "First");
        assert_cell_value(&workbook, 1, 0, 0, "Total");
    }

    #[test]
    fn test_merge_ranges() {
        let xlsx = XlsxBuilder::new()
            .sheet(
                SheetBuilder::new("Sheet1")
                    .cell("A1", "Merged", None)
                    .merge("A1:C1"),
            )
            .build();

        let workbook = load_xlsx(&xlsx);
        assert_merge_exists(&workbook, 0, 0, 0, 0, 2);
    }

    #[test]
    fn test_alignment() {
        let xlsx = XlsxBuilder::new()
            .sheet(SheetBuilder::new("Sheet1").cell(
                "A1",
                "Centered",
                Some(StyleBuilder::new().align_horizontal("center").build()),
            ))
            .build();

        let workbook = load_xlsx(&xlsx);
        assert_cell_align_h(&workbook, 0, 0, 0, "center");
    }

    #[test]
    fn test_wrap_text() {
        let xlsx = XlsxBuilder::new()
            .sheet(SheetBuilder::new("Sheet1").cell(
                "A1",
                "Wrapped",
                Some(StyleBuilder::new().wrap_text().build()),
            ))
            .build();

        let workbook = load_xlsx(&xlsx);
        assert_cell_wrap(&workbook, 0, 0, 0);
    }
}
