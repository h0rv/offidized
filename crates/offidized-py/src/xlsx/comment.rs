//! Python bindings for cell comments from `offidized_xlsx`.
//!
//! Wraps [`Comment`] as a value type (`XlsxComment`) that can be constructed,
//! inspected, and passed to worksheet helper functions. The worksheet helpers
//! (`ws_comments`, `ws_add_comment`, `ws_remove_comment`, `ws_clear_comments`)
//! are called from the parent `Worksheet` `#[pymethods]` block.

use super::lock_wb;
use crate::error::{value_error, xlsx_error_to_py};
use offidized_xlsx::Comment as CoreComment;
use offidized_xlsx::Workbook as CoreWorkbook;
use pyo3::prelude::*;
use std::sync::{Arc, Mutex};

// =============================================================================
// XlsxComment
// =============================================================================

/// Python wrapper for a cell comment (value type).
///
/// Comments are attached to a cell reference and carry an author string, plain
/// text, and an optional visibility flag. Construct via the normal
/// ``__init__`` and add to a sheet with :py:meth:`Worksheet.add_comment`.
#[pyclass(module = "offidized._native", name = "XlsxComment", from_py_object)]
#[derive(Clone)]
pub struct XlsxComment {
    cell_ref: String,
    author: String,
    text: String,
    visible: bool,
}

impl XlsxComment {
    /// Convert a core `Comment` reference to an `XlsxComment` value.
    pub(super) fn from_core(c: &CoreComment) -> Self {
        Self {
            cell_ref: c.cell_ref().to_string(),
            author: c.author().to_string(),
            text: c.text().to_string(),
            visible: c.visible(),
        }
    }

    /// Convert this wrapper back to a core `Comment`, validating the cell
    /// reference in the process.
    pub(super) fn to_core(&self) -> PyResult<CoreComment> {
        let mut core =
            CoreComment::new(&self.cell_ref, &self.author, &self.text).map_err(xlsx_error_to_py)?;
        core.set_visible(self.visible);
        Ok(core)
    }
}

#[pymethods]
impl XlsxComment {
    /// Create a new comment.
    ///
    /// Args:
    ///     cell_ref: A1-style cell reference the comment is attached to (e.g. ``"B3"``).
    ///     author:   Display name of the comment author.
    ///     text:     Plain-text comment body.
    ///     visible:  When ``True`` the comment box is always shown; when ``False``
    ///               it only appears on hover. Defaults to ``False``.
    #[new]
    #[pyo3(signature = (cell_ref, author, text, visible = false))]
    pub fn new(cell_ref: &str, author: &str, text: &str, visible: bool) -> Self {
        Self {
            cell_ref: cell_ref.to_string(),
            author: author.to_string(),
            text: text.to_string(),
            visible,
        }
    }

    /// The A1-style cell reference this comment is attached to.
    #[getter]
    pub fn cell_ref(&self) -> &str {
        &self.cell_ref
    }

    /// Set the cell reference.
    #[setter]
    pub fn set_cell_ref(&mut self, cell_ref: String) {
        self.cell_ref = cell_ref;
    }

    /// The comment author.
    #[getter]
    pub fn author(&self) -> &str {
        &self.author
    }

    /// Set the comment author.
    #[setter]
    pub fn set_author(&mut self, author: String) {
        self.author = author;
    }

    /// The comment plain-text body.
    #[getter]
    pub fn text(&self) -> &str {
        &self.text
    }

    /// Set the comment text.
    #[setter]
    pub fn set_text(&mut self, text: String) {
        self.text = text;
    }

    /// Whether the comment box is always visible.
    #[getter]
    pub fn visible(&self) -> bool {
        self.visible
    }

    /// Set comment visibility.
    #[setter]
    pub fn set_visible(&mut self, visible: bool) {
        self.visible = visible;
    }
}

// =============================================================================
// Worksheet helper functions
// =============================================================================

/// Return all comments on the worksheet as a list of `XlsxComment` objects.
pub(super) fn ws_comments(
    workbook: &Arc<Mutex<CoreWorkbook>>,
    name_key: &str,
) -> PyResult<Vec<XlsxComment>> {
    let wb = lock_wb(workbook)?;
    let ws = wb
        .sheet(name_key)
        .ok_or_else(|| value_error(format!("worksheet '{name_key}' not found")))?;
    Ok(ws.comments().iter().map(XlsxComment::from_core).collect())
}

/// Add a comment to the worksheet.
pub(super) fn ws_add_comment(
    workbook: &Arc<Mutex<CoreWorkbook>>,
    name_key: &str,
    comment: &XlsxComment,
) -> PyResult<()> {
    let core = comment.to_core()?;
    let mut wb = lock_wb(workbook)?;
    let ws = wb
        .sheet_mut(name_key)
        .ok_or_else(|| value_error(format!("worksheet '{name_key}' not found")))?;
    ws.add_comment(core);
    Ok(())
}

/// Remove the comment at the given cell reference. Returns ``True`` if a
/// comment was found and removed.
pub(super) fn ws_remove_comment(
    workbook: &Arc<Mutex<CoreWorkbook>>,
    name_key: &str,
    cell_ref: &str,
) -> PyResult<bool> {
    let mut wb = lock_wb(workbook)?;
    let ws = wb
        .sheet_mut(name_key)
        .ok_or_else(|| value_error(format!("worksheet '{name_key}' not found")))?;
    Ok(ws.remove_comment(cell_ref))
}

/// Remove all comments from the worksheet.
pub(super) fn ws_clear_comments(
    workbook: &Arc<Mutex<CoreWorkbook>>,
    name_key: &str,
) -> PyResult<()> {
    let mut wb = lock_wb(workbook)?;
    let ws = wb
        .sheet_mut(name_key)
        .ok_or_else(|| value_error(format!("worksheet '{name_key}' not found")))?;
    ws.clear_comments();
    Ok(())
}

// =============================================================================
// Registration
// =============================================================================

/// Register all comment PyO3 types with the native module.
pub(super) fn register(module: &Bound<'_, PyModule>) -> PyResult<()> {
    module.add_class::<XlsxComment>()?;
    Ok(())
}
