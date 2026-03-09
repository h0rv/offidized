//! Python bindings for worksheet table types from `offidized_xlsx`.
//!
//! Wraps [`WorksheetTable`] and [`TableColumn`] with PyO3 classes that mirror
//! the core Rust API. Worksheet helper functions (`ws_*`) are called from the
//! parent `Worksheet` `#[pymethods]` block.

use super::lock_wb;
use crate::error::{value_error, xlsx_error_to_py};
use offidized_xlsx::{
    TableColumn as CoreTableColumn, Workbook as CoreWorkbook, WorksheetTable as CoreWorksheetTable,
};
use pyo3::prelude::*;
use std::sync::{Arc, Mutex};

// =============================================================================
// XlsxTableColumn
// =============================================================================

/// Python wrapper for a single column definition within a worksheet table.
///
/// Maps to the ``<tableColumn>`` element in OOXML. Each column has a header
/// name and a 1-based integer ID. Optional totals-row label and formula
/// control what appears in the totals row when it is enabled.
#[pyclass(module = "offidized._native", name = "XlsxTableColumn", from_py_object)]
#[derive(Clone)]
pub struct XlsxTableColumn {
    inner: CoreTableColumn,
}

impl XlsxTableColumn {
    pub(super) fn from_core(col: CoreTableColumn) -> Self {
        Self { inner: col }
    }

    pub(super) fn into_core(self) -> CoreTableColumn {
        self.inner
    }
}

#[pymethods]
impl XlsxTableColumn {
    /// Create a new table column with the given header name and 1-based ID.
    #[new]
    pub fn new(name: &str, id: u32) -> Self {
        Self {
            inner: CoreTableColumn::new(name, id),
        }
    }

    /// The column header name.
    #[getter]
    pub fn name(&self) -> &str {
        self.inner.name()
    }

    /// Set the column header name.
    #[setter]
    pub fn set_name(&mut self, name: String) {
        self.inner.set_name(name);
    }

    /// The column ID (1-based). Read-only after construction.
    #[getter]
    pub fn id(&self) -> u32 {
        self.inner.id()
    }

    /// The label shown in the totals row for this column, or None.
    #[getter]
    pub fn totals_row_label(&self) -> Option<&str> {
        self.inner.totals_row_label()
    }

    /// Set the totals row label. Pass None to clear it.
    #[setter]
    pub fn set_totals_row_label(&mut self, label: Option<String>) {
        match label {
            Some(l) => {
                self.inner.set_totals_row_label(l);
            }
            None => {
                self.inner.clear_totals_row_label();
            }
        }
    }

    /// The custom formula used in the totals row for this column, or None.
    #[getter]
    pub fn totals_row_formula(&self) -> Option<&str> {
        self.inner.totals_row_formula()
    }

    /// Set the totals row formula. Pass None to clear it.
    #[setter]
    pub fn set_totals_row_formula(&mut self, formula: Option<String>) {
        match formula {
            Some(f) => {
                self.inner.set_totals_row_formula(f);
            }
            None => {
                self.inner.clear_totals_row_formula();
            }
        }
    }
}

// =============================================================================
// XlsxWorksheetTable
// =============================================================================

/// Python wrapper for a worksheet table definition.
///
/// A worksheet table (``<table>`` element in OOXML) defines a structured
/// range with optional header and totals rows, column definitions, and a
/// table style. Construct with :py:meth:`XlsxWorksheetTable.__init__` or
/// retrieve existing tables via :py:meth:`Worksheet.worksheet_tables`.
#[pyclass(
    module = "offidized._native",
    name = "XlsxWorksheetTable",
    from_py_object
)]
#[derive(Clone)]
pub struct XlsxWorksheetTable {
    inner: CoreWorksheetTable,
}

impl XlsxWorksheetTable {
    pub(super) fn from_core(table: CoreWorksheetTable) -> Self {
        Self { inner: table }
    }

    pub(super) fn into_core(self) -> CoreWorksheetTable {
        self.inner
    }
}

#[pymethods]
impl XlsxWorksheetTable {
    /// Create a new worksheet table.
    ///
    /// Args:
    ///     name: Unique table name (e.g. ``"SalesTable"``).
    ///     range: Cell range string the table covers (e.g. ``"A1:D10"``).
    #[new]
    pub fn new(name: &str, range: &str) -> PyResult<Self> {
        let inner = CoreWorksheetTable::new(name, range).map_err(xlsx_error_to_py)?;
        Ok(Self { inner })
    }

    /// The internal table name (used in formulas and references).
    #[getter]
    pub fn name(&self) -> &str {
        self.inner.name()
    }

    /// Set the table name.
    ///
    /// Raises :py:exc:`OffidizedValueError` if the name is invalid.
    #[setter]
    pub fn set_name(&mut self, name: String) -> PyResult<()> {
        self.inner.set_name(name).map_err(xlsx_error_to_py)?;
        Ok(())
    }

    /// The table range as a string formatted ``"start:end"`` (e.g. ``"A1:D10"``).
    #[getter]
    pub fn range(&self) -> String {
        let r = self.inner.range();
        format!("{}:{}", r.start(), r.end())
    }

    /// Whether the table has a header row.
    #[getter]
    pub fn show_header_row(&self) -> bool {
        self.inner.has_header_row()
    }

    /// Set whether the table has a header row.
    #[setter]
    pub fn set_show_header_row(&mut self, value: bool) {
        self.inner.set_header_row(value);
    }

    /// Whether the totals row is shown, or None if not explicitly set.
    #[getter]
    pub fn show_totals_row(&self) -> Option<bool> {
        self.inner.totals_row_shown()
    }

    /// Set whether the totals row is shown. Pass None to clear the setting.
    #[setter]
    pub fn set_show_totals_row(&mut self, value: Option<bool>) {
        match value {
            Some(v) => {
                self.inner.set_totals_row_shown(v);
            }
            None => {
                self.inner.clear_totals_row_shown();
            }
        }
    }

    /// The table style name (e.g. ``"TableStyleMedium9"``), or None.
    #[getter]
    pub fn style_name(&self) -> Option<&str> {
        self.inner.style_name()
    }

    /// Set the table style name. Pass None to clear it.
    #[setter]
    pub fn set_style_name(&mut self, name: Option<String>) {
        match name {
            Some(n) => {
                self.inner.set_style_name(n);
            }
            None => {
                self.inner.clear_style_name();
            }
        }
    }

    /// Return the column definitions as a list of :py:class:`XlsxTableColumn` objects.
    pub fn columns(&self) -> Vec<XlsxTableColumn> {
        self.inner
            .columns()
            .iter()
            .cloned()
            .map(XlsxTableColumn::from_core)
            .collect()
    }

    /// Append a column definition to the table.
    pub fn add_column(&mut self, column: &XlsxTableColumn) {
        self.inner.push_column(column.inner.clone());
    }
}

// =============================================================================
// Worksheet helper functions
// =============================================================================

/// Return all worksheet tables as a list of :py:class:`XlsxWorksheetTable` objects.
pub(super) fn ws_tables(
    workbook: &Arc<Mutex<CoreWorkbook>>,
    name_key: &str,
) -> PyResult<Vec<XlsxWorksheetTable>> {
    let wb = lock_wb(workbook)?;
    let ws = wb
        .sheet(name_key)
        .ok_or_else(|| value_error(format!("worksheet '{name_key}' not found")))?;
    Ok(ws
        .tables()
        .iter()
        .cloned()
        .map(XlsxWorksheetTable::from_core)
        .collect())
}

/// Add a pre-built :py:class:`XlsxWorksheetTable` to the worksheet.
pub(super) fn ws_add_table(
    workbook: &Arc<Mutex<CoreWorkbook>>,
    name_key: &str,
    table: XlsxWorksheetTable,
) -> PyResult<()> {
    let mut wb = lock_wb(workbook)?;
    let ws = wb
        .sheet_mut(name_key)
        .ok_or_else(|| value_error(format!("worksheet '{name_key}' not found")))?;
    ws.add_table(table.into_core());
    Ok(())
}

// =============================================================================
// Registration
// =============================================================================

/// Register all worksheet table PyO3 types with the native module.
pub(super) fn register(module: &Bound<'_, PyModule>) -> PyResult<()> {
    module.add_class::<XlsxTableColumn>()?;
    module.add_class::<XlsxWorksheetTable>()?;
    Ok(())
}
