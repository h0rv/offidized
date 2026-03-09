//! WASM-powered PowerPoint presentation viewer.
//!
//! Parses `.pptx` files into a JSON-friendly view model that TypeScript
//! renders to styled HTML/CSS with absolute positioning.

pub mod bridge;
pub mod convert;
pub mod model;
pub mod units;
