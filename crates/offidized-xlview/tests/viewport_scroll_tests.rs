//! Viewport and scroll coordinate tests
//!
//! Tests for verifying scroll position, visible row/column calculation,
//! and coordinate transformations.
//!
//! Ported from xlview's viewport_scroll_tests.rs.

#![allow(
    clippy::unwrap_used,
    clippy::expect_used,
    clippy::indexing_slicing,
    clippy::float_cmp,
    clippy::panic
)]

use offidized_xlview::layout::{SheetLayout, Viewport};
use std::collections::{HashMap, HashSet};

/// Create a simple layout with uniform row/column sizes
fn create_test_layout(rows: u32, cols: u32, row_height: f32, col_width: f32) -> SheetLayout {
    let mut col_widths_map = HashMap::new();
    let mut row_heights_map = HashMap::new();

    // Set uniform sizes
    for c in 0..=cols {
        col_widths_map.insert(c, col_width);
    }
    for r in 0..=rows {
        row_heights_map.insert(r, row_height);
    }

    SheetLayout::new(
        rows,
        cols,
        &col_widths_map,
        &row_heights_map,
        &HashSet::new(),
        &HashSet::new(),
        &[],
        0, // frozen_rows
        0, // frozen_cols
    )
}

/// Create a layout with frozen panes
fn create_frozen_layout(
    rows: u32,
    cols: u32,
    frozen_rows: u32,
    frozen_cols: u32,
    row_height: f32,
    col_width: f32,
) -> SheetLayout {
    let mut col_widths_map = HashMap::new();
    let mut row_heights_map = HashMap::new();

    for c in 0..=cols {
        col_widths_map.insert(c, col_width);
    }
    for r in 0..=rows {
        row_heights_map.insert(r, row_height);
    }

    SheetLayout::new(
        rows,
        cols,
        &col_widths_map,
        &row_heights_map,
        &HashSet::new(),
        &HashSet::new(),
        &[],
        frozen_rows,
        frozen_cols,
    )
}

// =============================================================================
// BASIC VIEWPORT TESTS
// =============================================================================

#[test]
fn test_viewport_initial_scroll_zero() {
    let viewport = Viewport::new();
    assert_eq!(viewport.scroll_x, 0.0, "Initial scroll_x should be 0");
    assert_eq!(viewport.scroll_y, 0.0, "Initial scroll_y should be 0");
}

#[test]
fn test_visible_rows_at_scroll_zero() {
    let layout = create_test_layout(100, 10, 20.0, 80.0);
    let mut viewport = Viewport::new();
    viewport.width = 800.0;
    viewport.height = 600.0;
    viewport.scroll_x = 0.0;
    viewport.scroll_y = 0.0;

    let (start_row, end_row) = viewport.visible_rows(&layout);

    assert_eq!(start_row, 0, "Start row should be 0 at scroll_y=0");
    // End row should be approximately viewport_height / row_height = 600 / 20 = 30
    assert!(
        end_row >= 29,
        "End row should be at least 29 for 600px viewport with 20px rows, got {}",
        end_row
    );
    assert!(
        end_row <= 31,
        "End row should be at most 31, got {}",
        end_row
    );
}

#[test]
fn test_visible_cols_at_scroll_zero() {
    let layout = create_test_layout(100, 100, 20.0, 80.0);
    let mut viewport = Viewport::new();
    viewport.width = 800.0;
    viewport.height = 600.0;
    viewport.scroll_x = 0.0;
    viewport.scroll_y = 0.0;

    let (start_col, end_col) = viewport.visible_cols(&layout);

    assert_eq!(start_col, 0, "Start col should be 0 at scroll_x=0");
    // End col should be approximately viewport_width / col_width = 800 / 80 = 10
    assert!(
        end_col >= 9,
        "End col should be at least 9 for 800px viewport with 80px cols, got {}",
        end_col
    );
    assert!(
        end_col <= 11,
        "End col should be at most 11, got {}",
        end_col
    );
}

#[test]
fn test_visible_rows_after_scroll() {
    let layout = create_test_layout(100, 10, 20.0, 80.0);
    let mut viewport = Viewport::new();
    viewport.width = 800.0;
    viewport.height = 600.0;
    // Scroll down by 10 rows (200px)
    viewport.scroll_y = 200.0;

    let (start_row, _end_row) = viewport.visible_rows(&layout);

    // At scroll_y=200, first visible row should be row 10 (200/20)
    assert_eq!(
        start_row, 10,
        "Start row should be 10 at scroll_y=200 with 20px rows"
    );
}

// =============================================================================
// FROZEN PANE TESTS
// =============================================================================

#[test]
fn test_visible_rows_with_frozen_rows() {
    let layout = create_frozen_layout(100, 10, 2, 0, 20.0, 80.0);
    let mut viewport = Viewport::new();
    viewport.width = 800.0;
    viewport.height = 600.0;
    // With frozen rows, scroll should start at frozen boundary
    viewport.scroll_y = layout.frozen_rows_height(); // 40.0

    let (start_row, _end_row) = viewport.visible_rows(&layout);

    // At scroll_y=40 (frozen boundary), visible scrollable rows start at row 2
    assert_eq!(start_row, 2, "Start row should be 2 (first non-frozen row)");
}

#[test]
fn test_scroll_clamp_with_frozen_panes() {
    let layout = create_frozen_layout(100, 100, 3, 2, 20.0, 80.0);
    let mut viewport = Viewport::new();
    viewport.width = 800.0;
    viewport.height = 600.0;

    // Try to scroll to 0,0 (before frozen boundary)
    viewport.scroll_x = 0.0;
    viewport.scroll_y = 0.0;
    viewport.clamp_scroll(&layout);

    // Should clamp to frozen boundaries
    let frozen_x = layout.frozen_cols_width(); // 2 * 80 = 160
    let frozen_y = layout.frozen_rows_height(); // 3 * 20 = 60

    assert_eq!(
        viewport.scroll_x, frozen_x,
        "scroll_x should clamp to frozen_cols_width"
    );
    assert_eq!(
        viewport.scroll_y, frozen_y,
        "scroll_y should clamp to frozen_rows_height"
    );
}

// =============================================================================
// COORDINATE TRANSFORMATION TESTS
// =============================================================================

#[test]
fn test_to_screen_at_scroll_zero() {
    let viewport = Viewport {
        scroll_x: 0.0,
        scroll_y: 0.0,
        width: 800.0,
        height: 600.0,
        scale: 1.0,
        tab_scroll_x: 0.0,
    };

    // Sheet coordinate (100, 50) should map to screen (100, 50) at scroll 0
    let (screen_x, screen_y) = viewport.to_screen(100.0, 50.0);

    assert_eq!(screen_x, 100.0, "Screen X should equal sheet X at scroll 0");
    assert_eq!(screen_y, 50.0, "Screen Y should equal sheet Y at scroll 0");
}

#[test]
fn test_to_screen_after_scroll() {
    let viewport = Viewport {
        scroll_x: 100.0,
        scroll_y: 200.0,
        width: 800.0,
        height: 600.0,
        scale: 1.0,
        tab_scroll_x: 0.0,
    };

    // Sheet coordinate (150, 250) should map to screen (50, 50) after scroll
    let (screen_x, screen_y) = viewport.to_screen(150.0, 250.0);

    assert_eq!(screen_x, 50.0, "Screen X should be sheet_x - scroll_x");
    assert_eq!(screen_y, 50.0, "Screen Y should be sheet_y - scroll_y");
}

#[test]
fn test_to_sheet_roundtrip() {
    let viewport = Viewport {
        scroll_x: 100.0,
        scroll_y: 200.0,
        width: 800.0,
        height: 600.0,
        scale: 1.0,
        tab_scroll_x: 0.0,
    };

    let original_sheet = (300.0, 400.0);
    let screen = viewport.to_screen(original_sheet.0, original_sheet.1);
    let back_to_sheet = viewport.to_sheet(screen.0, screen.1);

    assert!(
        (back_to_sheet.0 - original_sheet.0).abs() < 0.001,
        "X roundtrip should match"
    );
    assert!(
        (back_to_sheet.1 - original_sheet.1).abs() < 0.001,
        "Y roundtrip should match"
    );
}

// =============================================================================
// LAYOUT ROW/COL AT POSITION TESTS
// =============================================================================

#[test]
fn test_row_at_y_zero() {
    let layout = create_test_layout(100, 10, 20.0, 80.0);

    let row = layout.row_at_y(0.0);
    assert_eq!(row, Some(0), "row_at_y(0) should return row 0");
}

#[test]
fn test_row_at_y_middle_of_row() {
    let layout = create_test_layout(100, 10, 20.0, 80.0);

    // Middle of row 5 (y = 100 + 10 = 110)
    let row = layout.row_at_y(110.0);
    assert_eq!(
        row,
        Some(5),
        "row_at_y(110) should return row 5 for 20px rows"
    );
}

#[test]
fn test_col_at_x_zero() {
    let layout = create_test_layout(100, 10, 20.0, 80.0);

    let col = layout.col_at_x(0.0);
    assert_eq!(col, Some(0), "col_at_x(0) should return col 0");
}

#[test]
fn test_col_at_x_middle_of_col() {
    let layout = create_test_layout(100, 100, 20.0, 80.0);

    // Middle of col 3 (x = 240 + 40 = 280)
    let col = layout.col_at_x(280.0);
    assert_eq!(
        col,
        Some(3),
        "col_at_x(280) should return col 3 for 80px cols"
    );
}

// =============================================================================
// SPACER SIZE TESTS (for scroll range)
// =============================================================================

#[test]
fn test_scrollable_content_size_no_frozen() {
    let layout = create_test_layout(100, 50, 20.0, 80.0);

    let total_width = layout.total_width();
    let total_height = layout.total_height();
    let scrollable_width = total_width - layout.frozen_cols_width();
    let scrollable_height = total_height - layout.frozen_rows_height();

    // With no frozen panes, scrollable = total
    assert_eq!(
        scrollable_width, total_width,
        "No frozen cols: scrollable = total width"
    );
    assert_eq!(
        scrollable_height, total_height,
        "No frozen rows: scrollable = total height"
    );

    // Verify expected values: 51 cols * 80px = 4080, 101 rows * 20px = 2020 (includes extra position)
    assert!(
        total_width > 4000.0,
        "Total width should be > 4000, got {}",
        total_width
    );
    assert!(
        total_height > 2000.0,
        "Total height should be > 2000, got {}",
        total_height
    );
}

#[test]
fn test_scrollable_content_size_with_frozen() {
    let layout = create_frozen_layout(100, 50, 3, 2, 20.0, 80.0);

    let frozen_w = layout.frozen_cols_width(); // 2 * 80 = 160
    let frozen_h = layout.frozen_rows_height(); // 3 * 20 = 60

    assert_eq!(frozen_w, 160.0, "Frozen cols width should be 160");
    assert_eq!(frozen_h, 60.0, "Frozen rows height should be 60");

    let total_width = layout.total_width();
    let total_height = layout.total_height();
    let scrollable_width = total_width - frozen_w;
    let scrollable_height = total_height - frozen_h;

    assert!(
        scrollable_width > 0.0,
        "Scrollable width should be positive"
    );
    assert!(
        scrollable_height > 0.0,
        "Scrollable height should be positive"
    );
}

// =============================================================================
// SCROLL COORDINATE MAPPING TESTS
// These test the mapping between container scroll and viewport scroll
// =============================================================================

#[test]
fn test_container_scroll_to_viewport_no_frozen() {
    let layout = create_test_layout(100, 50, 20.0, 80.0);

    // Simulate: container.scroll_left = 100, container.scroll_top = 200
    let container_scroll_left: f32 = 100.0;
    let container_scroll_top: f32 = 200.0;

    // viewport.scroll = container.scroll + frozen
    let viewport_scroll_x = container_scroll_left + layout.frozen_cols_width();
    let viewport_scroll_y = container_scroll_top + layout.frozen_rows_height();

    // With no frozen panes, viewport.scroll = container.scroll
    assert_eq!(viewport_scroll_x, 100.0);
    assert_eq!(viewport_scroll_y, 200.0);
}

#[test]
fn test_container_scroll_to_viewport_with_frozen() {
    let layout = create_frozen_layout(100, 50, 3, 2, 20.0, 80.0);

    // Simulate: container.scroll_left = 0, container.scroll_top = 0
    // (start of scrollable area)
    let container_scroll_left: f32 = 0.0;
    let container_scroll_top: f32 = 0.0;

    // viewport.scroll = container.scroll + frozen
    let viewport_scroll_x = container_scroll_left + layout.frozen_cols_width();
    let viewport_scroll_y = container_scroll_top + layout.frozen_rows_height();

    // Should be at the frozen boundary
    assert_eq!(
        viewport_scroll_x, 160.0,
        "viewport.scroll_x = frozen_cols_width when container.scroll=0"
    );
    assert_eq!(
        viewport_scroll_y, 60.0,
        "viewport.scroll_y = frozen_rows_height when container.scroll=0"
    );
}

#[test]
fn test_visible_row_from_container_scroll_zero() {
    let layout = create_test_layout(100, 50, 20.0, 80.0);
    let mut viewport = Viewport::new();
    viewport.width = 800.0;
    viewport.height = 600.0;

    // Container scroll = 0, no frozen panes
    // viewport.scroll = 0 + 0 = 0
    viewport.scroll_x = 0.0 + layout.frozen_cols_width();
    viewport.scroll_y = 0.0 + layout.frozen_rows_height();

    let (start_row, _) = viewport.visible_rows(&layout);
    let (start_col, _) = viewport.visible_cols(&layout);

    assert_eq!(
        start_row, 0,
        "First visible row should be 0 when container scroll is 0"
    );
    assert_eq!(
        start_col, 0,
        "First visible col should be 0 when container scroll is 0"
    );
}

// =============================================================================
// HEADER-AWARE VIEWPORT TESTS
// Test that visible rows/cols are correct when headers are visible
// =============================================================================

#[test]
fn test_visible_rows_with_header_offset() {
    // Simulate viewport with headers visible
    // Headers don't affect the viewport.scroll calculation,
    // but they do affect the visible content area
    let layout = create_test_layout(100, 10, 20.0, 80.0);
    let mut viewport = Viewport::new();

    // Typical viewport with headers
    // Logical viewport: 1200x800
    // Header dimensions: row_header_width=40, col_header_height=20
    viewport.width = 1200.0;
    viewport.height = 800.0;
    viewport.scroll_x = 0.0;
    viewport.scroll_y = 0.0;

    let (start_row, end_row) = viewport.visible_rows(&layout);
    let (start_col, _end_col) = viewport.visible_cols(&layout);

    // With scroll at 0,0:
    // - First visible row should be 0
    // - First visible col should be 0
    assert_eq!(start_row, 0, "First visible row should be 0");
    assert_eq!(start_col, 0, "First visible col should be 0");

    // End row should be approximately viewport_height / row_height
    // 800 / 20 = 40 rows (but actually slightly less due to rounding)
    assert!(
        end_row >= 39,
        "End row should be at least 39 for 800px viewport with 20px rows, got {}",
        end_row
    );
}

#[test]
fn test_scroll_coordinate_after_header_area() {
    // Test that scrolling by a fixed amount gives the correct row
    let layout = create_test_layout(1000, 100, 20.0, 80.0);
    let mut viewport = Viewport::new();
    viewport.width = 1200.0;
    viewport.height = 800.0;

    // Scroll to row 30 (should require scroll_y = 600px)
    viewport.scroll_y = 600.0;

    let (start_row, _) = viewport.visible_rows(&layout);
    assert_eq!(
        start_row, 30,
        "At scroll_y=600 with 20px rows, first visible row should be 30"
    );

    // Scroll back to 0
    viewport.scroll_y = 0.0;
    let (start_row, _) = viewport.visible_rows(&layout);
    assert_eq!(start_row, 0, "At scroll_y=0, first visible row should be 0");
}

// =============================================================================
// REGRESSION TEST: Scroll offset bug
// When container scroll is 0, we should see row 0, not row 30
// =============================================================================

#[test]
fn test_no_initial_scroll_offset() {
    // Simulate the exact scenario from the bug report:
    // - Fresh load
    // - Container scroll = 0
    // - Should see row 0/col 0, not row 30

    let layout = create_test_layout(1000, 100, 20.0, 80.0);
    let mut viewport = Viewport::new();

    // Typical viewport size
    viewport.width = 1200.0;
    viewport.height = 800.0;

    // This is what sync_scroll_from_container does:
    // viewport.scroll = container.scroll + frozen
    let container_scroll_left = 0;
    let container_scroll_top = 0;
    viewport.scroll_x = container_scroll_left as f32 + layout.frozen_cols_width();
    viewport.scroll_y = container_scroll_top as f32 + layout.frozen_rows_height();

    let (start_row, _) = viewport.visible_rows(&layout);
    let (start_col, _) = viewport.visible_cols(&layout);

    // BUG: If this fails with start_row = 30 or similar, there's an offset bug
    assert_eq!(
        start_row, 0,
        "REGRESSION: First visible row should be 0, not ~30"
    );
    assert_eq!(start_col, 0, "REGRESSION: First visible col should be 0");
}
