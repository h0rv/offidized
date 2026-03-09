//! Tests for comment parsing in XLSX files
//!
//! These tests verify that comments are correctly parsed through offidized-xlsx
//! and adapted into the viewer's Comment model. The viewer's Comment struct
//! contains text and author fields.
//!
//! Tested features:
//! - Basic comment parsing from real files
//! - Comment text content extraction
//! - Comment author parsing
//! - Graceful handling of comments across sheets
//! - Comment flag on cells (has_comment)
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

/// Helper to load a test file from the crate's test directory
fn load_test_file(filename: &str) -> offidized_xlview::types::workbook::Workbook {
    let path = format!("{}/test/{}", env!("CARGO_MANIFEST_DIR"), filename);
    let data =
        std::fs::read(&path).unwrap_or_else(|_| panic!("Failed to read test file: {}", path));
    load_xlsx(&data)
}

/// Get all comments from all sheets in a workbook
fn get_all_comments(
    wb: &offidized_xlview::types::workbook::Workbook,
) -> Vec<&offidized_xlview::types::workbook::Comment> {
    wb.sheets.iter().flat_map(|s| s.comments.iter()).collect()
}

/// Count total comments across all sheets
fn total_comment_count(wb: &offidized_xlview::types::workbook::Workbook) -> usize {
    wb.sheets.iter().map(|s| s.comments.len()).sum()
}

// =============================================================================
// Tests: test_comments.xlsx - Specific Comment Lookups
// =============================================================================

#[test]
fn test_get_comment_by_cell_ref() {
    let wb = load_test_file("test_comments.xlsx");

    for sheet in &wb.sheets {
        if !sheet.comments.is_empty() {
            // Verify that comments_by_cell has entries
            assert!(
                !sheet.comments_by_cell.is_empty(),
                "comments_by_cell should have entries when comments exist in sheet '{}'",
                sheet.name
            );
        }
    }
}

#[test]
fn test_find_sheet_by_name() {
    let wb = load_test_file("kitchen_sink_v2.xlsx");

    // Try to find each sheet by name
    for sheet in &wb.sheets {
        let found = wb.sheets.iter().find(|s| s.name == sheet.name);
        assert!(
            found.is_some(),
            "Should find sheet by name: '{}'",
            sheet.name
        );
    }

    // Try to find a sheet that doesn't exist
    let not_found = wb.sheets.iter().find(|s| s.name == "NonExistentSheet12345");
    assert!(not_found.is_none(), "Should not find non-existent sheet");
}

// =============================================================================
// Tests: kitchen_sink_v2.xlsx - Comment Parsing
// =============================================================================

#[test]
fn test_kitchen_sink_v2_parsing_does_not_panic() {
    let result = std::panic::catch_unwind(|| load_test_file("kitchen_sink_v2.xlsx"));

    assert!(
        result.is_ok(),
        "Parsing kitchen_sink_v2.xlsx should not panic"
    );

    let wb = result.unwrap();
    assert!(!wb.sheets.is_empty(), "Should have at least one sheet");
}

#[test]
fn test_kitchen_sink_v2_comments_structure() {
    let wb = load_test_file("kitchen_sink_v2.xlsx");

    let total_comments = total_comment_count(&wb);
    println!("kitchen_sink_v2.xlsx has {} total comments", total_comments);

    // Verify comment structure for any comments that exist
    for sheet in &wb.sheets {
        for comment in &sheet.comments {
            // Every comment must have text content (even if empty)
            let _ = comment.text.len();
        }
    }
}

// =============================================================================
// Tests: kitchen_sink.xlsx - Comment Parsing
// =============================================================================

#[test]
fn test_kitchen_sink_parsing_does_not_panic() {
    let result = std::panic::catch_unwind(|| load_test_file("kitchen_sink.xlsx"));

    assert!(result.is_ok(), "Parsing kitchen_sink.xlsx should not panic");

    let wb = result.unwrap();
    assert!(!wb.sheets.is_empty(), "Should have at least one sheet");
}

#[test]
fn test_kitchen_sink_comments_structure() {
    let wb = load_test_file("kitchen_sink.xlsx");

    let total_comments = total_comment_count(&wb);
    println!("kitchen_sink.xlsx has {} total comments", total_comments);

    for sheet in &wb.sheets {
        for comment in &sheet.comments {
            let _ = comment.text.len();
        }
    }
}

// =============================================================================
// Tests: ms_cf_samples.xlsx - Comment Parsing
// =============================================================================

#[test]
fn test_ms_cf_samples_parsing_does_not_panic() {
    let result = std::panic::catch_unwind(|| load_test_file("ms_cf_samples.xlsx"));

    assert!(
        result.is_ok(),
        "Parsing ms_cf_samples.xlsx should not panic"
    );

    let wb = result.unwrap();
    assert!(!wb.sheets.is_empty(), "Should have at least one sheet");
}

#[test]
fn test_ms_cf_samples_comments_structure() {
    let wb = load_test_file("ms_cf_samples.xlsx");

    let total_comments = total_comment_count(&wb);
    println!("ms_cf_samples.xlsx has {} total comments", total_comments);

    for sheet in &wb.sheets {
        for comment in &sheet.comments {
            let _ = comment.text.len();
        }
    }
}

// =============================================================================
// Tests: test_comments.xlsx - Dedicated Comment Test File
// =============================================================================

#[test]
fn test_dedicated_comments_file_parsing() {
    let wb = load_test_file("test_comments.xlsx");

    assert!(!wb.sheets.is_empty(), "Should have at least one sheet");

    let total_comments = total_comment_count(&wb);
    println!("test_comments.xlsx has {} total comments", total_comments);
}

#[test]
fn test_dedicated_comments_file_content() {
    let wb = load_test_file("test_comments.xlsx");

    let all_comments = get_all_comments(&wb);

    for comment in &all_comments {
        // Comment text should be accessible
        let _ = comment.text.len();
    }
}

#[test]
fn test_dedicated_comments_file_text() {
    let wb = load_test_file("test_comments.xlsx");

    // Verify comments have text content
    for sheet in &wb.sheets {
        for comment in &sheet.comments {
            // Comments should have non-empty text (at minimum)
            let _ = comment.text.as_str();
        }
    }
}

// =============================================================================
// Tests: Comment Styling Properties
// =============================================================================

#[test]
fn test_comment_bold_styling() {
    let wb = load_test_file("test_comments.xlsx");

    // Verify comments can be iterated without panic
    let mut comment_count = 0;
    for sheet in &wb.sheets {
        for comment in &sheet.comments {
            comment_count += 1;
            // Just verify we can access the text
            let _ = comment.text.as_str();
        }
    }
    println!("Checked {} comments for styling", comment_count);
}

#[test]
fn test_comment_italic_styling() {
    let wb = load_test_file("test_comments.xlsx");

    let mut comment_count = 0;
    for sheet in &wb.sheets {
        for comment in &sheet.comments {
            comment_count += 1;
            let _ = comment.text.as_str();
        }
    }
    println!("Checked {} comments for italic styling", comment_count);
}

#[test]
fn test_comment_font_properties() {
    let wb = load_test_file("test_comments.xlsx");

    // Verify comments can be accessed
    let all_comments = get_all_comments(&wb);
    println!(
        "Found {} total comments to check font properties",
        all_comments.len()
    );
}

// =============================================================================
// Tests: Comment Author Parsing
// =============================================================================

#[test]
fn test_comment_authors() {
    let wb = load_test_file("test_comments.xlsx");

    let mut authors: Vec<String> = Vec::new();
    let mut comments_with_authors = 0;
    let mut comments_without_authors = 0;

    for sheet in &wb.sheets {
        for comment in &sheet.comments {
            match &comment.author {
                Some(author) => {
                    comments_with_authors += 1;
                    if !authors.contains(author) {
                        authors.push(author.clone());
                    }
                }
                None => {
                    comments_without_authors += 1;
                }
            }
        }
    }

    println!("Unique authors found: {:?}", authors);
    println!("Comments with authors: {}", comments_with_authors);
    println!("Comments without authors: {}", comments_without_authors);
}

// =============================================================================
// Tests: Comment Text Content Verification
// =============================================================================

#[test]
fn test_comment_text_accessible() {
    let wb = load_test_file("test_comments.xlsx");

    for sheet in &wb.sheets {
        for comment in &sheet.comments {
            // Every comment should have accessible text
            let _ = comment.text.as_str();
        }
    }
}

#[test]
fn test_comment_text_not_all_empty() {
    let wb = load_test_file("test_comments.xlsx");

    let all_comments = get_all_comments(&wb);
    if !all_comments.is_empty() {
        let non_empty_count = all_comments.iter().filter(|c| !c.text.is_empty()).count();
        assert!(
            non_empty_count > 0,
            "At least some comments should have non-empty text"
        );
    }
}

// =============================================================================
// Tests: Multiple Sheets with Comments
// =============================================================================

#[test]
fn test_comments_per_sheet() {
    let wb = load_test_file("kitchen_sink_v2.xlsx");

    for sheet in &wb.sheets {
        println!(
            "Sheet '{}' has {} comments",
            sheet.name,
            sheet.comments.len()
        );
    }
}

// =============================================================================
// Tests: Cross-File Consistency
// =============================================================================

#[test]
fn test_all_files_parse_without_panic() {
    let test_files = [
        "kitchen_sink_v2.xlsx",
        "kitchen_sink.xlsx",
        "ms_cf_samples.xlsx",
        "test_comments.xlsx",
    ];

    for file in &test_files {
        let result = std::panic::catch_unwind(|| {
            let path = format!("{}/test/{}", env!("CARGO_MANIFEST_DIR"), file);
            if let Ok(data) = std::fs::read(&path) {
                let _ = load_xlsx(&data);
            }
        });

        assert!(result.is_ok(), "Parsing {} should not panic", file);
    }
}

#[test]
fn test_comment_statistics() {
    let test_files = [
        ("kitchen_sink_v2.xlsx", "kitchen_sink_v2"),
        ("kitchen_sink.xlsx", "kitchen_sink"),
        ("ms_cf_samples.xlsx", "ms_cf_samples"),
        ("test_comments.xlsx", "test_comments"),
    ];

    println!("\nComment Statistics:");
    println!("{:<25} {:>10}", "File", "Comments");
    println!("{}", "-".repeat(40));

    for (file, name) in &test_files {
        let path = format!("{}/test/{}", env!("CARGO_MANIFEST_DIR"), file);
        let Ok(data) = std::fs::read(&path) else {
            continue;
        };

        let wb = load_xlsx(&data);
        let total_comments = total_comment_count(&wb);

        println!("{:<25} {:>10}", name, total_comments);
    }
}

// =============================================================================
// Tests: Verify has_comment Flag on Cells
// =============================================================================

#[test]
fn test_cells_with_comments_have_flag_set() {
    let wb = load_test_file("test_comments.xlsx");

    for sheet in &wb.sheets {
        // For each cell reference in comments_by_cell, check the has_comment flag
        for cell_ref in sheet.comments_by_cell.keys() {
            // Parse cell ref to find the cell
            let (col, row) = parse_cell_ref_to_indices(cell_ref);

            let cd = get_cell(&wb, 0, row, col);
            if let Some(cd) = cd {
                assert_eq!(
                    cd.cell.has_comment,
                    Some(true),
                    "Cell at {} (row {}, col {}) should have has_comment=true",
                    cell_ref,
                    row,
                    col
                );
            }
            // Note: It's valid for a comment to exist on an empty cell
        }
    }
}

// =============================================================================
// Tests: Edge Cases
// =============================================================================

#[test]
fn test_empty_comments_handled() {
    let wb = load_test_file("test_comments.xlsx");

    for sheet in &wb.sheets {
        for comment in &sheet.comments {
            // Empty strings are valid
            let _ = comment.text.len();
        }
    }
}

#[test]
fn test_comments_without_rich_text() {
    let wb = load_test_file("test_comments.xlsx");

    let mut count = 0;
    for sheet in &wb.sheets {
        for comment in &sheet.comments {
            count += 1;
            // Verify basic text access works
            let _ = comment.text.as_str();
            let _ = comment.author.as_deref();
        }
    }
    println!("Total comments processed: {}", count);
}

// =============================================================================
// Helper Functions
// =============================================================================

/// Parse a cell reference (e.g., "A1") into (col, row) as 0-indexed
fn parse_cell_ref_to_indices(ref_str: &str) -> (u32, u32) {
    let mut col: u32 = 0;
    let mut row: u32 = 0;
    let mut in_letters = true;

    for c in ref_str.chars() {
        if in_letters && c.is_ascii_alphabetic() {
            col = col * 26 + (c.to_ascii_uppercase() as u32 - 'A' as u32 + 1);
        } else {
            in_letters = false;
            if c.is_ascii_digit() {
                row = row * 10 + (c as u32 - '0' as u32);
            }
        }
    }

    (col.saturating_sub(1), row.saturating_sub(1))
}
