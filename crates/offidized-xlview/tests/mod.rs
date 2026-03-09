//! Integration tests for offidized-xlview.
//!
//! This module provides the test infrastructure for testing the XLSX viewer
//! adapter pipeline: XLSX bytes -> offidized_xlsx::Workbook -> viewer Workbook.
//!
//! - `fixtures`: Builders for creating valid XLSX files in memory
//! - `common`: Assertion helpers and parsing utilities
//! - `test_helpers`: Lower-level XLSX builder (ZIP/XML)
//!
//! # Example Usage
//!
//! ```rust,ignore
//! use crate::fixtures::{XlsxBuilder, StyleBuilder, SheetBuilder};
//! use crate::common::{load_xlsx, assert_cell_value, assert_cell_bold};
//!
//! fn test_bold_text() {
//!     let xlsx = XlsxBuilder::new()
//!         .sheet(SheetBuilder::new("Sheet1")
//!             .cell("A1", "Bold Text", Some(StyleBuilder::new().bold().build())))
//!         .build();
//!
//!     let workbook = load_xlsx(&xlsx);
//!     assert_cell_value(&workbook, 0, 0, 0, "Bold Text");
//!     assert_cell_bold(&workbook, 0, 0, 0);
//! }
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

pub mod common;
pub mod fixtures;
pub mod test_helpers;

// NOTE: Individual test files (alignment_tests.rs, border_tests.rs, etc.) are
// standalone integration test binaries. They each declare their own `mod common;`
// and `mod fixtures;` to pull in shared infrastructure. Do NOT declare them here
// as submodules — that would cause path resolution conflicts.
