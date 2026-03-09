//! WASM-powered Word document viewer.
//!
//! Parses `.docx` files into a JSON-friendly view model that TypeScript
//! renders to styled HTML/CSS.

pub mod bridge;
pub mod convert;
#[cfg(feature = "editing")]
pub mod editor;
pub mod model;
pub mod units;
