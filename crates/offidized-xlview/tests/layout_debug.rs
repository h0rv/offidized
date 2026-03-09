//! Debug test for layout issues with kitchen_sink.xlsx
//!
//! Ported from xlview's layout_debug.rs. Uses offidized-xlview adapter + layout types.
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

use offidized_xlview::layout::SheetLayout;
use std::collections::{HashMap, HashSet};

mod common;
mod fixtures;
use common::load_xlsx;

#[test]
fn test_kitchen_sink_layout() {
    let Ok(data) = std::fs::read("test/kitchen_sink.xlsx") else {
        eprintln!("Skipping test: could not read test file");
        return;
    };

    let workbook = load_xlsx(&data);

    println!("\n=== Kitchen Sink Debug ===");
    println!("Sheets: {}", workbook.sheets.len());

    for (idx, sheet) in workbook.sheets.iter().enumerate() {
        println!("\n--- Sheet {}: {} ---", idx, sheet.name);
        println!("Cells: {}", sheet.cells.len());
        println!("Col widths defined: {}", sheet.col_widths.len());
        println!("Row heights defined: {}", sheet.row_heights.len());
        println!("Merges: {}", sheet.merges.len());

        if !sheet.col_widths.is_empty() {
            println!("Column widths:");
            for cw in &sheet.col_widths {
                println!(
                    "  Col {}: {} units -> {} px",
                    cw.col,
                    cw.width,
                    cw.width * 7.0
                );
            }
        }

        let mut col_widths_map: HashMap<u32, f32> = HashMap::new();
        for cw in &sheet.col_widths {
            let width_px = (cw.width * 7.0) as f32;
            col_widths_map.insert(cw.col, width_px);
        }

        let mut row_heights_map: HashMap<u32, f32> = HashMap::new();
        for rh in &sheet.row_heights {
            let height_px = (rh.height * 1.33) as f32;
            row_heights_map.insert(rh.row, height_px);
        }

        let hidden_cols: HashSet<u32> = sheet.hidden_cols.iter().copied().collect();
        let hidden_rows: HashSet<u32> = sheet.hidden_rows.iter().copied().collect();

        let merge_ranges: Vec<(u32, u32, u32, u32)> = sheet
            .merges
            .iter()
            .map(|m| (m.start_row, m.start_col, m.end_row, m.end_col))
            .collect();

        let max_row = sheet.cells.iter().map(|c| c.r).max().unwrap_or(0).max(20);
        let max_col = sheet.cells.iter().map(|c| c.c).max().unwrap_or(0).max(10);

        println!("Max row: {}, Max col: {}", max_row, max_col);

        let layout = SheetLayout::new(
            max_row,
            max_col,
            &col_widths_map,
            &row_heights_map,
            &hidden_cols,
            &hidden_rows,
            &merge_ranges,
            sheet.frozen_rows,
            sheet.frozen_cols,
        );

        println!("Total width: {} px", layout.total_width());
        println!("Total height: {} px", layout.total_height());

        println!("Column positions (first 10):");
        for col in 0..10.min(layout.col_positions.len() as u32) {
            let pos = layout
                .col_positions
                .get(col as usize)
                .copied()
                .unwrap_or(0.0);
            let width = layout.col_widths.get(col as usize).copied().unwrap_or(0.0);
            println!("  Col {}: x={:.1}, width={:.1}", col, pos, width);
        }

        println!("First 10 cells:");
        for cell in sheet.cells.iter().take(10) {
            let has_style = cell.cell.s.is_some() || cell.cell.style_idx.is_some();
            let has_bg = cell
                .cell
                .s
                .as_ref()
                .and_then(|s| s.bg_color.as_ref())
                .is_some();
            println!(
                "  ({}, {}): {:?} [style={}, bg={}]",
                cell.r,
                cell.c,
                cell.cell.v.as_ref().map(|s| if s.len() > 20 {
                    format!("{}...", &s[..20])
                } else {
                    s.clone()
                }),
                has_style,
                has_bg
            );
        }
    }
}

#[test]
fn test_visible_range_calculation() {
    let Ok(data) = std::fs::read("test/kitchen_sink.xlsx") else {
        eprintln!("Skipping test: could not read test file");
        return;
    };

    let workbook = load_xlsx(&data);

    let Some(sheet) = workbook.sheets.first() else {
        eprintln!("Skipping test: no sheets in workbook");
        return;
    };

    let mut col_widths_map: HashMap<u32, f32> = HashMap::new();
    for cw in &sheet.col_widths {
        col_widths_map.insert(cw.col, (cw.width * 7.0) as f32);
    }

    let layout = SheetLayout::new(
        100,
        26,
        &col_widths_map,
        &HashMap::new(),
        &HashSet::new(),
        &HashSet::new(),
        &[],
        sheet.frozen_rows,
        sheet.frozen_cols,
    );

    let viewport = offidized_xlview::layout::Viewport {
        scroll_x: 0.0,
        scroll_y: 0.0,
        width: 800.0,
        height: 600.0,
        scale: 1.0,
        tab_scroll_x: 0.0,
    };

    let (start_row, end_row) = viewport.visible_rows(&layout);
    let (start_col, end_col) = viewport.visible_cols(&layout);

    println!("\n=== Visible Range Test ===");
    println!("Viewport: {}x{}", viewport.width, viewport.height);
    println!("Visible rows: {} to {}", start_row, end_row);
    println!("Visible cols: {} to {}", start_col, end_col);
    println!("Total layout width: {}", layout.total_width());

    assert!(end_col > 0, "Should see more than just column 0!");
    assert!(end_row > 0, "Should see more than just row 0!");
}
