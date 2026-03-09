//! Rendering engine with pluggable backends.
//!
//! This module provides:
//! - Backend-agnostic rendering traits and types
//! - Canvas 2D backend (primary, stable)
//! - Color parsing utilities
//!
//! Future backends (vello/WebGPU) can be added by implementing RenderBackend.

pub mod backend;
pub mod blit;
pub mod cache;
pub mod canvas;
pub mod colors;
pub mod selection;

// Re-export commonly used types
pub use crate::types::rich_text::TextRunData;
pub use backend::{BorderStyleData, CellRenderData, CellStyleData, RenderBackend, RenderParams};
pub use canvas::CanvasRenderer;
pub use colors::{palette, parse_color, CssColor};
