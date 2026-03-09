//! Canvas 2D rendering backend.
//!
//! This module provides spreadsheet rendering using the HTML Canvas 2D API
//! via web-sys. It's simpler and more stable than WebGPU/vello, while being
//! perfectly suited for spreadsheet rendering (rectangles, lines, text).

mod charts;
mod conditional;
mod frozen;
pub mod headers;
mod indicators;
mod renderer;
mod shapes;
mod sparklines;

pub use renderer::CanvasRenderer;
