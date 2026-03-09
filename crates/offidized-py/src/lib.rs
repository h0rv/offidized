//! PyO3 bindings for the offidized Rust OOXML APIs.
//!
//! This module provides comprehensive Python bindings for all three document formats:
//! - Excel (.xlsx) via Workbook
//! - Word (.docx) via Document
//! - PowerPoint (.pptx) via Presentation
//! - IR derive/apply workflow via ir_derive/ir_apply

mod docx;
mod error;
mod ir;
mod pptx;
mod xlsx;

use error::register_exceptions;
use pyo3::prelude::*;

#[pymodule]
fn _native(py: Python<'_>, module: &Bound<'_, PyModule>) -> PyResult<()> {
    register_exceptions(py, module)?;
    xlsx::register(module)?;
    docx::register(module)?;
    pptx::register(module)?;
    ir::register(module)?;
    Ok(())
}
