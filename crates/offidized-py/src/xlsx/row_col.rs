//! Python bindings for row and column metadata from `offidized_xlsx`.
//!
//! Wraps [`Row`] and [`Column`] with PyO3 classes that mirror the core Rust
//! API. Instances hold a reference to the parent workbook mutex rather than
//! owning the underlying data, so mutations are reflected in the workbook
//! immediately without requiring a re-assignment.

use super::lock_wb;
use crate::error::{value_error, xlsx_error_to_py};
use offidized_xlsx::column::Column as CoreColumn;
use offidized_xlsx::row::Row as CoreRow;
use offidized_xlsx::Workbook as CoreWorkbook;
use pyo3::prelude::*;
use std::sync::{Arc, Mutex};

// =============================================================================
// XlsxRow
// =============================================================================

/// Python wrapper for row metadata within a worksheet.
///
/// Holds a reference to the parent workbook by name so that mutations reach
/// the underlying :class:`Workbook` directly. Row indices are 1-based,
/// matching the OOXML row numbering convention.
#[pyclass(module = "offidized._native", name = "XlsxRow")]
pub struct XlsxRow {
    workbook: Arc<Mutex<CoreWorkbook>>,
    sheet_name: String,
    row_index: u32,
}

impl XlsxRow {
    /// Construct a new `XlsxRow` wrapper.
    pub(super) fn new(
        workbook: Arc<Mutex<CoreWorkbook>>,
        sheet_name: String,
        row_index: u32,
    ) -> Self {
        Self {
            workbook,
            sheet_name,
            row_index,
        }
    }
}

#[pymethods]
impl XlsxRow {
    /// Return the 1-based row index.
    pub fn index(&self) -> u32 {
        self.row_index
    }

    /// Return the row height in points, or ``None`` if unset.
    pub fn height(&self) -> PyResult<Option<f64>> {
        let wb = lock_wb(&self.workbook)?;
        let ws = wb
            .sheet(&self.sheet_name)
            .ok_or_else(|| value_error(format!("worksheet '{}' not found", self.sheet_name)))?;
        Ok(ws.row(self.row_index).and_then(CoreRow::height))
    }

    /// Set the row height in points.
    pub fn set_height(&self, height: f64) -> PyResult<()> {
        let mut wb = lock_wb(&self.workbook)?;
        let ws = wb
            .sheet_mut(&self.sheet_name)
            .ok_or_else(|| value_error(format!("worksheet '{}' not found", self.sheet_name)))?;
        ws.row_mut(self.row_index)
            .map_err(xlsx_error_to_py)?
            .set_height(height);
        Ok(())
    }

    /// Return whether the row is hidden.
    pub fn is_hidden(&self) -> PyResult<bool> {
        let wb = lock_wb(&self.workbook)?;
        let ws = wb
            .sheet(&self.sheet_name)
            .ok_or_else(|| value_error(format!("worksheet '{}' not found", self.sheet_name)))?;
        Ok(ws.row(self.row_index).is_some_and(CoreRow::is_hidden))
    }

    /// Set the row hidden state.
    pub fn set_hidden(&self, hidden: bool) -> PyResult<()> {
        let mut wb = lock_wb(&self.workbook)?;
        let ws = wb
            .sheet_mut(&self.sheet_name)
            .ok_or_else(|| value_error(format!("worksheet '{}' not found", self.sheet_name)))?;
        ws.row_mut(self.row_index)
            .map_err(xlsx_error_to_py)?
            .set_hidden(hidden);
        Ok(())
    }

    /// Return whether the row has a custom height.
    pub fn custom_height(&self) -> PyResult<bool> {
        let wb = lock_wb(&self.workbook)?;
        let ws = wb
            .sheet(&self.sheet_name)
            .ok_or_else(|| value_error(format!("worksheet '{}' not found", self.sheet_name)))?;
        Ok(ws.row(self.row_index).is_some_and(CoreRow::custom_height))
    }

    /// Return the outline (grouping) level for this row (0–7).
    pub fn outline_level(&self) -> PyResult<u8> {
        let wb = lock_wb(&self.workbook)?;
        let ws = wb
            .sheet(&self.sheet_name)
            .ok_or_else(|| value_error(format!("worksheet '{}' not found", self.sheet_name)))?;
        Ok(ws.row(self.row_index).map_or(0u8, CoreRow::outline_level))
    }

    /// Set the outline (grouping) level for this row (0–7).
    pub fn set_outline_level(&self, level: u8) -> PyResult<()> {
        let mut wb = lock_wb(&self.workbook)?;
        let ws = wb
            .sheet_mut(&self.sheet_name)
            .ok_or_else(|| value_error(format!("worksheet '{}' not found", self.sheet_name)))?;
        ws.row_mut(self.row_index)
            .map_err(xlsx_error_to_py)?
            .set_outline_level(level);
        Ok(())
    }

    /// Return whether the row group is collapsed.
    pub fn is_collapsed(&self) -> PyResult<bool> {
        let wb = lock_wb(&self.workbook)?;
        let ws = wb
            .sheet(&self.sheet_name)
            .ok_or_else(|| value_error(format!("worksheet '{}' not found", self.sheet_name)))?;
        Ok(ws.row(self.row_index).is_some_and(CoreRow::is_collapsed))
    }

    /// Set whether the row group is collapsed.
    pub fn set_collapsed(&self, collapsed: bool) -> PyResult<()> {
        let mut wb = lock_wb(&self.workbook)?;
        let ws = wb
            .sheet_mut(&self.sheet_name)
            .ok_or_else(|| value_error(format!("worksheet '{}' not found", self.sheet_name)))?;
        ws.row_mut(self.row_index)
            .map_err(xlsx_error_to_py)?
            .set_collapsed(collapsed);
        Ok(())
    }
}

// =============================================================================
// XlsxColumn
// =============================================================================

/// Python wrapper for column metadata within a worksheet.
///
/// Holds a reference to the parent workbook by name so that mutations reach
/// the underlying :class:`Workbook` directly. Column indices are 1-based,
/// matching the OOXML column numbering convention.
#[pyclass(module = "offidized._native", name = "XlsxColumn")]
pub struct XlsxColumn {
    workbook: Arc<Mutex<CoreWorkbook>>,
    sheet_name: String,
    col_index: u32,
}

impl XlsxColumn {
    /// Construct a new `XlsxColumn` wrapper.
    pub(super) fn new(
        workbook: Arc<Mutex<CoreWorkbook>>,
        sheet_name: String,
        col_index: u32,
    ) -> Self {
        Self {
            workbook,
            sheet_name,
            col_index,
        }
    }
}

#[pymethods]
impl XlsxColumn {
    /// Return the 1-based column index.
    pub fn index(&self) -> u32 {
        self.col_index
    }

    /// Return the column width in character units, or ``None`` if unset.
    pub fn width(&self) -> PyResult<Option<f64>> {
        let wb = lock_wb(&self.workbook)?;
        let ws = wb
            .sheet(&self.sheet_name)
            .ok_or_else(|| value_error(format!("worksheet '{}' not found", self.sheet_name)))?;
        Ok(ws.column(self.col_index).and_then(CoreColumn::width))
    }

    /// Set the column width in character units.
    pub fn set_width(&self, width: f64) -> PyResult<()> {
        let mut wb = lock_wb(&self.workbook)?;
        let ws = wb
            .sheet_mut(&self.sheet_name)
            .ok_or_else(|| value_error(format!("worksheet '{}' not found", self.sheet_name)))?;
        ws.column_mut(self.col_index)
            .map_err(xlsx_error_to_py)?
            .set_width(width);
        Ok(())
    }

    /// Return whether the column is hidden.
    pub fn is_hidden(&self) -> PyResult<bool> {
        let wb = lock_wb(&self.workbook)?;
        let ws = wb
            .sheet(&self.sheet_name)
            .ok_or_else(|| value_error(format!("worksheet '{}' not found", self.sheet_name)))?;
        Ok(ws.column(self.col_index).is_some_and(CoreColumn::is_hidden))
    }

    /// Set the column hidden state.
    pub fn set_hidden(&self, hidden: bool) -> PyResult<()> {
        let mut wb = lock_wb(&self.workbook)?;
        let ws = wb
            .sheet_mut(&self.sheet_name)
            .ok_or_else(|| value_error(format!("worksheet '{}' not found", self.sheet_name)))?;
        ws.column_mut(self.col_index)
            .map_err(xlsx_error_to_py)?
            .set_hidden(hidden);
        Ok(())
    }

    /// Return the outline (grouping) level for this column (0–7).
    pub fn outline_level(&self) -> PyResult<u8> {
        let wb = lock_wb(&self.workbook)?;
        let ws = wb
            .sheet(&self.sheet_name)
            .ok_or_else(|| value_error(format!("worksheet '{}' not found", self.sheet_name)))?;
        Ok(ws
            .column(self.col_index)
            .map_or(0u8, CoreColumn::outline_level))
    }

    /// Set the outline (grouping) level for this column (0–7).
    pub fn set_outline_level(&self, level: u8) -> PyResult<()> {
        let mut wb = lock_wb(&self.workbook)?;
        let ws = wb
            .sheet_mut(&self.sheet_name)
            .ok_or_else(|| value_error(format!("worksheet '{}' not found", self.sheet_name)))?;
        ws.column_mut(self.col_index)
            .map_err(xlsx_error_to_py)?
            .set_outline_level(level);
        Ok(())
    }

    /// Return whether the column group is collapsed.
    pub fn is_collapsed(&self) -> PyResult<bool> {
        let wb = lock_wb(&self.workbook)?;
        let ws = wb
            .sheet(&self.sheet_name)
            .ok_or_else(|| value_error(format!("worksheet '{}' not found", self.sheet_name)))?;
        Ok(ws
            .column(self.col_index)
            .is_some_and(CoreColumn::is_collapsed))
    }

    /// Set whether the column group is collapsed.
    pub fn set_collapsed(&self, collapsed: bool) -> PyResult<()> {
        let mut wb = lock_wb(&self.workbook)?;
        let ws = wb
            .sheet_mut(&self.sheet_name)
            .ok_or_else(|| value_error(format!("worksheet '{}' not found", self.sheet_name)))?;
        ws.column_mut(self.col_index)
            .map_err(xlsx_error_to_py)?
            .set_collapsed(collapsed);
        Ok(())
    }

    /// Return the style index applied to this column, or ``None``.
    pub fn style_index(&self) -> PyResult<Option<u32>> {
        let wb = lock_wb(&self.workbook)?;
        let ws = wb
            .sheet(&self.sheet_name)
            .ok_or_else(|| value_error(format!("worksheet '{}' not found", self.sheet_name)))?;
        Ok(ws.column(self.col_index).and_then(CoreColumn::style_index))
    }

    /// Set the style index for this column.
    pub fn set_style_index(&self, index: u32) -> PyResult<()> {
        let mut wb = lock_wb(&self.workbook)?;
        let ws = wb
            .sheet_mut(&self.sheet_name)
            .ok_or_else(|| value_error(format!("worksheet '{}' not found", self.sheet_name)))?;
        ws.column_mut(self.col_index)
            .map_err(xlsx_error_to_py)?
            .set_style_index(index);
        Ok(())
    }

    /// Return whether the column width is a best-fit auto width.
    pub fn is_best_fit(&self) -> PyResult<bool> {
        let wb = lock_wb(&self.workbook)?;
        let ws = wb
            .sheet(&self.sheet_name)
            .ok_or_else(|| value_error(format!("worksheet '{}' not found", self.sheet_name)))?;
        Ok(ws
            .column(self.col_index)
            .is_some_and(CoreColumn::is_best_fit))
    }

    /// Set whether the column width is a best-fit auto width.
    pub fn set_best_fit(&self, best_fit: bool) -> PyResult<()> {
        let mut wb = lock_wb(&self.workbook)?;
        let ws = wb
            .sheet_mut(&self.sheet_name)
            .ok_or_else(|| value_error(format!("worksheet '{}' not found", self.sheet_name)))?;
        ws.column_mut(self.col_index)
            .map_err(xlsx_error_to_py)?
            .set_best_fit(best_fit);
        Ok(())
    }

    /// Return whether the column has a custom (non-default) width.
    pub fn custom_width(&self) -> PyResult<bool> {
        let wb = lock_wb(&self.workbook)?;
        let ws = wb
            .sheet(&self.sheet_name)
            .ok_or_else(|| value_error(format!("worksheet '{}' not found", self.sheet_name)))?;
        Ok(ws
            .column(self.col_index)
            .is_some_and(CoreColumn::custom_width))
    }

    /// Set whether the column has a custom (non-default) width.
    pub fn set_custom_width(&self, custom_width: bool) -> PyResult<()> {
        let mut wb = lock_wb(&self.workbook)?;
        let ws = wb
            .sheet_mut(&self.sheet_name)
            .ok_or_else(|| value_error(format!("worksheet '{}' not found", self.sheet_name)))?;
        ws.column_mut(self.col_index)
            .map_err(xlsx_error_to_py)?
            .set_custom_width(custom_width);
        Ok(())
    }
}

// =============================================================================
// Registration
// =============================================================================

/// Register all row/column PyO3 types with the native module.
pub(super) fn register(module: &Bound<'_, PyModule>) -> PyResult<()> {
    module.add_class::<XlsxRow>()?;
    module.add_class::<XlsxColumn>()?;
    Ok(())
}
