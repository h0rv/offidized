//! # offidized-xlview
//!
//! WASM-powered Excel spreadsheet viewer with Canvas 2D rendering.
//!
//! This crate provides a high-performance Excel file viewer that renders
//! spreadsheets using HTML5 Canvas 2D, with tile caching for smooth
//! scrolling at 120fps even with 100k+ cells.
//!
//! Built on top of `offidized-xlsx` for parsing and data model.

pub mod adapter;
pub mod cell_ref;
#[allow(dead_code)]
pub mod csv;
#[cfg(feature = "editing")]
pub mod editor;
pub mod error;
#[cfg(feature = "editing")]
pub mod export;
pub mod layout;
pub mod render;
pub mod types;
pub mod viewer;
