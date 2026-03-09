//! Tests for tab bar positioning to ensure tabs are always visible.
//!
//! The tab bar should always be rendered at the bottom of the canvas,
//! regardless of viewport size.
//!
//! Ported from xlview's tab_bar_tests.rs.
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

/// TAB_BAR_HEIGHT constant (must match src/viewer.rs and src/render/canvas/renderer.rs)
const TAB_BAR_HEIGHT: f32 = 28.0;

/// SCROLLBAR_SIZE constant (must match src/render/canvas/renderer.rs)
const SCROLLBAR_SIZE: f32 = 14.0;

/// Simulates the viewport height calculation from viewer.rs
fn calculate_viewport_height(logical_canvas_height: f32) -> f32 {
    (logical_canvas_height - TAB_BAR_HEIGHT).max(0.0)
}

/// Simulates the tab bar Y position calculation from renderer.rs
fn calculate_tab_y(viewport_height: f32) -> f32 {
    viewport_height // Tab bar starts at bottom of content area
}

/// Simulates the total canvas height for clearing from renderer.rs
fn calculate_total_height(viewport_height: f32) -> f32 {
    viewport_height + TAB_BAR_HEIGHT
}

/// Simulates content area height calculation from renderer.rs
fn calculate_content_height(viewport_height: f32) -> f32 {
    viewport_height - SCROLLBAR_SIZE
}

#[test]
fn test_tab_bar_within_canvas_bounds() {
    // Test various canvas heights
    for canvas_height in [400.0, 600.0, 800.0, 1000.0, 1200.0] {
        let viewport_height = calculate_viewport_height(canvas_height);
        let tab_y = calculate_tab_y(viewport_height);
        let total_height = calculate_total_height(viewport_height);

        // Tab bar should start within the canvas
        assert!(
            tab_y < total_height,
            "Tab bar Y ({}) should be less than total height ({}) for canvas height {}",
            tab_y,
            total_height,
            canvas_height
        );

        // Tab bar should end exactly at canvas bottom
        let tab_bar_bottom = tab_y + TAB_BAR_HEIGHT;
        assert!(
            (tab_bar_bottom - total_height).abs() < 0.001,
            "Tab bar bottom ({}) should equal total height ({}) for canvas height {}",
            tab_bar_bottom,
            total_height,
            canvas_height
        );

        // Total height should equal original canvas height
        assert!(
            (total_height - canvas_height).abs() < 0.001,
            "Total height ({}) should equal canvas height ({}) after viewport calculation",
            total_height,
            canvas_height
        );
    }
}

#[test]
fn test_tab_bar_visible_for_small_viewports() {
    // Even with a very small canvas, tab bar should be positioned correctly
    let canvas_height = 100.0; // Very small
    let viewport_height = calculate_viewport_height(canvas_height);
    let tab_y = calculate_tab_y(viewport_height);
    let total_height = calculate_total_height(viewport_height);

    // viewport_height = 100 - 28 = 72
    assert_eq!(viewport_height, 72.0);

    // tab_y = 72 (starts at bottom of content)
    assert_eq!(tab_y, 72.0);

    // total_height = 72 + 28 = 100 (original canvas height)
    assert_eq!(total_height, 100.0);

    // Tab bar is from y=72 to y=100, which is within 0-100 canvas bounds
    assert!(tab_y >= 0.0, "Tab Y should be non-negative");
    assert!(
        tab_y + TAB_BAR_HEIGHT <= total_height,
        "Tab bar should fit within canvas"
    );
}

#[test]
fn test_content_area_does_not_overlap_tab_bar() {
    for canvas_height in [400.0, 600.0, 800.0] {
        let viewport_height = calculate_viewport_height(canvas_height);
        let content_height = calculate_content_height(viewport_height);
        let tab_y = calculate_tab_y(viewport_height);

        // Content area (0 to content_height) should not overlap with tab bar (tab_y to tab_y + TAB_BAR_HEIGHT)
        // Content ends at content_height = viewport_height - SCROLLBAR_SIZE
        // Tab bar starts at tab_y = viewport_height
        // So content_height < tab_y (there's a gap for the scrollbar)
        assert!(
            content_height < tab_y,
            "Content area ({}) should end before tab bar starts ({}) for canvas {}",
            content_height,
            tab_y,
            canvas_height
        );

        // The gap should be exactly SCROLLBAR_SIZE
        let gap = tab_y - content_height;
        assert!(
            (gap - SCROLLBAR_SIZE).abs() < 0.001,
            "Gap between content and tab bar should be {} (scrollbar), got {}",
            SCROLLBAR_SIZE,
            gap
        );
    }
}

#[test]
fn test_scrollbar_position() {
    let canvas_height = 600.0;
    let viewport_height = calculate_viewport_height(canvas_height);
    let content_height = calculate_content_height(viewport_height);

    // Horizontal scrollbar Y position (from renderer.rs logic)
    let h_scrollbar_y = content_height; // Scrollbar is at bottom of content area

    // Scrollbar should be above the tab bar
    let tab_y = calculate_tab_y(viewport_height);

    // Scrollbar is from h_scrollbar_y to h_scrollbar_y + SCROLLBAR_SIZE
    let scrollbar_bottom = h_scrollbar_y + SCROLLBAR_SIZE;

    assert!(
        (scrollbar_bottom - tab_y).abs() < 0.001,
        "Scrollbar bottom ({}) should meet tab bar top ({})",
        scrollbar_bottom,
        tab_y
    );
}

#[test]
fn test_minimum_viewport_height() {
    // When canvas is smaller than TAB_BAR_HEIGHT, viewport should be 0
    let canvas_height = 20.0; // Less than TAB_BAR_HEIGHT (28)
    let viewport_height = calculate_viewport_height(canvas_height);

    assert_eq!(
        viewport_height, 0.0,
        "Viewport height should be 0 when canvas is too small"
    );

    // Tab bar should still be at y=0
    let tab_y = calculate_tab_y(viewport_height);
    assert_eq!(tab_y, 0.0);

    // Total height should be TAB_BAR_HEIGHT (the minimum to show tab bar)
    let total_height = calculate_total_height(viewport_height);
    assert_eq!(total_height, TAB_BAR_HEIGHT);
}

/// Test that verifies the relationship between all the height values
#[test]
fn test_height_relationships() {
    let canvas_height = 800.0;

    // 1. Viewport height is canvas minus tab bar
    let viewport_height = calculate_viewport_height(canvas_height);
    assert_eq!(viewport_height, canvas_height - TAB_BAR_HEIGHT);

    // 2. Total height (for canvas clear) equals original canvas
    let total_height = calculate_total_height(viewport_height);
    assert_eq!(total_height, canvas_height);

    // 3. Content height is viewport minus scrollbar
    let content_height = calculate_content_height(viewport_height);
    assert_eq!(content_height, viewport_height - SCROLLBAR_SIZE);

    // 4. Tab Y equals viewport height (content area bottom)
    let tab_y = calculate_tab_y(viewport_height);
    assert_eq!(tab_y, viewport_height);

    // Layout from top to bottom:
    // [0, content_height) = cell content area
    // [content_height, content_height + SCROLLBAR_SIZE) = horizontal scrollbar
    // [viewport_height, viewport_height + TAB_BAR_HEIGHT) = tab bar
    //
    // Note: content_height + SCROLLBAR_SIZE = viewport_height
    assert_eq!(content_height + SCROLLBAR_SIZE, viewport_height);

    println!("Canvas height: {}", canvas_height);
    println!("Viewport height: {}", viewport_height);
    println!("Content height: {}", content_height);
    println!(
        "Scrollbar area: {} to {}",
        content_height,
        content_height + SCROLLBAR_SIZE
    );
    println!("Tab bar area: {} to {}", tab_y, tab_y + TAB_BAR_HEIGHT);
    println!("Total height: {}", total_height);
}
