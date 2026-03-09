//! Python bindings for sheet and workbook protection from `offidized_xlsx`.
//!
//! Wraps [`SheetProtection`] and [`WorkbookProtection`] as value types that can
//! be inspected and mutated via properties, then applied to a worksheet or
//! workbook through the helper functions at the bottom of this module.
//!
//! Worksheet helpers (`ws_protection_detail`, `ws_set_protection_detail`) and
//! workbook helpers (`wb_workbook_protection`, `wb_set_workbook_protection`,
//! `wb_clear_workbook_protection`) are called from the parent `Worksheet` and
//! `Workbook` `#[pymethods]` blocks respectively.

use super::lock_wb;
use crate::error::value_error;
use offidized_xlsx::SheetProtection as CoreSheetProtection;
use offidized_xlsx::Workbook as CoreWorkbook;
use offidized_xlsx::WorkbookProtection as CoreWorkbookProtection;
use pyo3::prelude::*;
use std::sync::{Arc, Mutex};

// =============================================================================
// XlsxSheetProtection
// =============================================================================

/// Python wrapper for worksheet protection settings (value type).
///
/// Each boolean property corresponds to a restriction that is enforced when
/// the sheet is protected. The ``sheet`` flag must be ``True`` for protection
/// to be active; all other flags default to ``False``.
///
/// Apply to a worksheet with :py:meth:`Worksheet.set_protection_detail`.
#[pyclass(module = "offidized._native", name = "XlsxSheetProtection")]
#[derive(Clone)]
pub struct XlsxSheetProtection {
    inner: CoreSheetProtection,
}

impl XlsxSheetProtection {
    fn from_core(p: &CoreSheetProtection) -> Self {
        Self { inner: p.clone() }
    }

    fn into_core(self) -> CoreSheetProtection {
        self.inner
    }
}

#[pymethods]
impl XlsxSheetProtection {
    /// Create a new sheet protection with the ``sheet`` flag enabled.
    #[new]
    pub fn new() -> Self {
        Self {
            inner: CoreSheetProtection::new(),
        }
    }

    /// Whether the sheet itself is protected.
    #[getter]
    pub fn sheet(&self) -> bool {
        self.inner.sheet()
    }

    /// Set whether the sheet itself is protected.
    #[setter]
    pub fn set_sheet(&mut self, value: bool) {
        self.inner.set_sheet(value);
    }

    /// Whether objects are protected.
    #[getter]
    pub fn objects(&self) -> bool {
        self.inner.objects()
    }

    /// Set whether objects are protected.
    #[setter]
    pub fn set_objects(&mut self, value: bool) {
        self.inner.set_objects(value);
    }

    /// Whether scenarios are protected.
    #[getter]
    pub fn scenarios(&self) -> bool {
        self.inner.scenarios()
    }

    /// Set whether scenarios are protected.
    #[setter]
    pub fn set_scenarios(&mut self, value: bool) {
        self.inner.set_scenarios(value);
    }

    /// Whether formatting cells is disallowed.
    #[getter]
    pub fn format_cells(&self) -> bool {
        self.inner.format_cells()
    }

    /// Set whether formatting cells is disallowed.
    #[setter]
    pub fn set_format_cells(&mut self, value: bool) {
        self.inner.set_format_cells(value);
    }

    /// Whether formatting columns is disallowed.
    #[getter]
    pub fn format_columns(&self) -> bool {
        self.inner.format_columns()
    }

    /// Set whether formatting columns is disallowed.
    #[setter]
    pub fn set_format_columns(&mut self, value: bool) {
        self.inner.set_format_columns(value);
    }

    /// Whether formatting rows is disallowed.
    #[getter]
    pub fn format_rows(&self) -> bool {
        self.inner.format_rows()
    }

    /// Set whether formatting rows is disallowed.
    #[setter]
    pub fn set_format_rows(&mut self, value: bool) {
        self.inner.set_format_rows(value);
    }

    /// Whether inserting columns is disallowed.
    #[getter]
    pub fn insert_columns(&self) -> bool {
        self.inner.insert_columns()
    }

    /// Set whether inserting columns is disallowed.
    #[setter]
    pub fn set_insert_columns(&mut self, value: bool) {
        self.inner.set_insert_columns(value);
    }

    /// Whether inserting rows is disallowed.
    #[getter]
    pub fn insert_rows(&self) -> bool {
        self.inner.insert_rows()
    }

    /// Set whether inserting rows is disallowed.
    #[setter]
    pub fn set_insert_rows(&mut self, value: bool) {
        self.inner.set_insert_rows(value);
    }

    /// Whether inserting hyperlinks is disallowed.
    #[getter]
    pub fn insert_hyperlinks(&self) -> bool {
        self.inner.insert_hyperlinks()
    }

    /// Set whether inserting hyperlinks is disallowed.
    #[setter]
    pub fn set_insert_hyperlinks(&mut self, value: bool) {
        self.inner.set_insert_hyperlinks(value);
    }

    /// Whether deleting columns is disallowed.
    #[getter]
    pub fn delete_columns(&self) -> bool {
        self.inner.delete_columns()
    }

    /// Set whether deleting columns is disallowed.
    #[setter]
    pub fn set_delete_columns(&mut self, value: bool) {
        self.inner.set_delete_columns(value);
    }

    /// Whether deleting rows is disallowed.
    #[getter]
    pub fn delete_rows(&self) -> bool {
        self.inner.delete_rows()
    }

    /// Set whether deleting rows is disallowed.
    #[setter]
    pub fn set_delete_rows(&mut self, value: bool) {
        self.inner.set_delete_rows(value);
    }

    /// Whether selecting locked cells is disallowed.
    #[getter]
    pub fn select_locked_cells(&self) -> bool {
        self.inner.select_locked_cells()
    }

    /// Set whether selecting locked cells is disallowed.
    #[setter]
    pub fn set_select_locked_cells(&mut self, value: bool) {
        self.inner.set_select_locked_cells(value);
    }

    /// Whether sorting is disallowed.
    #[getter]
    pub fn sort(&self) -> bool {
        self.inner.sort()
    }

    /// Set whether sorting is disallowed.
    #[setter]
    pub fn set_sort(&mut self, value: bool) {
        self.inner.set_sort(value);
    }

    /// Whether using auto-filter is disallowed.
    #[getter]
    pub fn auto_filter(&self) -> bool {
        self.inner.auto_filter()
    }

    /// Set whether using auto-filter is disallowed.
    #[setter]
    pub fn set_auto_filter(&mut self, value: bool) {
        self.inner.set_auto_filter(value);
    }

    /// Whether using pivot tables is disallowed.
    #[getter]
    pub fn pivot_tables(&self) -> bool {
        self.inner.pivot_tables()
    }

    /// Set whether using pivot tables is disallowed.
    #[setter]
    pub fn set_pivot_tables(&mut self, value: bool) {
        self.inner.set_pivot_tables(value);
    }

    /// Whether selecting unlocked cells is disallowed.
    #[getter]
    pub fn select_unlocked_cells(&self) -> bool {
        self.inner.select_unlocked_cells()
    }

    /// Set whether selecting unlocked cells is disallowed.
    #[setter]
    pub fn set_select_unlocked_cells(&mut self, value: bool) {
        self.inner.set_select_unlocked_cells(value);
    }

    /// The pre-hashed password string, or ``None`` if no password is set.
    #[getter]
    pub fn password_hash(&self) -> Option<&str> {
        self.inner.password_hash()
    }

    /// Set the pre-hashed password. Pass ``None`` to clear it.
    #[setter]
    pub fn set_password_hash(&mut self, hash: Option<String>) {
        match hash {
            Some(h) => {
                self.inner.set_password_hash(h);
            }
            None => {
                self.inner.clear_password_hash();
            }
        }
    }
}

// =============================================================================
// XlsxWorkbookProtection
// =============================================================================

/// Python wrapper for workbook-level protection settings (value type).
///
/// Controls whether the workbook structure (sheet tabs) or window positions
/// are locked. Apply with :py:meth:`Workbook.set_workbook_protection`.
#[pyclass(module = "offidized._native", name = "XlsxWorkbookProtection")]
#[derive(Clone)]
pub struct XlsxWorkbookProtection {
    inner: CoreWorkbookProtection,
}

impl XlsxWorkbookProtection {
    fn from_core(p: &CoreWorkbookProtection) -> Self {
        Self { inner: p.clone() }
    }

    fn into_core(self) -> CoreWorkbookProtection {
        self.inner
    }
}

#[pymethods]
impl XlsxWorkbookProtection {
    /// Create a new workbook protection with ``lock_structure`` enabled.
    #[new]
    pub fn new() -> Self {
        Self {
            inner: CoreWorkbookProtection::new(),
        }
    }

    /// Whether the workbook structure (sheet tabs) is locked.
    #[getter]
    pub fn lock_structure(&self) -> bool {
        self.inner.lock_structure()
    }

    /// Set whether the workbook structure is locked.
    #[setter]
    pub fn set_lock_structure(&mut self, value: bool) {
        self.inner.set_lock_structure(value);
    }

    /// Whether windows are locked.
    #[getter]
    pub fn lock_windows(&self) -> bool {
        self.inner.lock_windows()
    }

    /// Set whether windows are locked.
    #[setter]
    pub fn set_lock_windows(&mut self, value: bool) {
        self.inner.set_lock_windows(value);
    }

    /// The pre-hashed password string, or ``None`` if no password is set.
    #[getter]
    pub fn password_hash(&self) -> Option<&str> {
        self.inner.password_hash()
    }

    /// Set the pre-hashed password. Pass ``None`` to clear it.
    #[setter]
    pub fn set_password_hash(&mut self, hash: Option<String>) {
        match hash {
            Some(h) => {
                self.inner.set_password_hash(h);
            }
            None => {
                self.inner.clear_password_hash();
            }
        }
    }
}

// =============================================================================
// Worksheet helper functions
// =============================================================================

/// Return the detailed sheet protection settings as an `XlsxSheetProtection`,
/// or ``None`` if the sheet has no protection applied.
pub(super) fn ws_protection_detail(
    workbook: &Arc<Mutex<CoreWorkbook>>,
    name_key: &str,
) -> PyResult<Option<XlsxSheetProtection>> {
    let wb = lock_wb(workbook)?;
    let ws = wb
        .sheet(name_key)
        .ok_or_else(|| value_error(format!("worksheet '{name_key}' not found")))?;
    Ok(ws.protection().map(XlsxSheetProtection::from_core))
}

/// Apply detailed sheet protection settings. Replaces any existing protection.
pub(super) fn ws_set_protection_detail(
    workbook: &Arc<Mutex<CoreWorkbook>>,
    name_key: &str,
    protection: XlsxSheetProtection,
) -> PyResult<()> {
    let core = protection.into_core();
    let mut wb = lock_wb(workbook)?;
    let ws = wb
        .sheet_mut(name_key)
        .ok_or_else(|| value_error(format!("worksheet '{name_key}' not found")))?;
    ws.set_protection(core);
    Ok(())
}

// =============================================================================
// Workbook helper functions
// =============================================================================

/// Return the workbook protection settings, or ``None`` if none are set.
pub(super) fn wb_workbook_protection(
    workbook: &Arc<Mutex<CoreWorkbook>>,
) -> PyResult<Option<XlsxWorkbookProtection>> {
    let wb = lock_wb(workbook)?;
    Ok(wb
        .workbook_protection()
        .map(XlsxWorkbookProtection::from_core))
}

/// Apply workbook protection settings. Replaces any existing protection.
pub(super) fn wb_set_workbook_protection(
    workbook: &Arc<Mutex<CoreWorkbook>>,
    protection: XlsxWorkbookProtection,
) -> PyResult<()> {
    let core = protection.into_core();
    let mut wb = lock_wb(workbook)?;
    wb.set_workbook_protection(core);
    Ok(())
}

/// Clear workbook protection entirely.
pub(super) fn wb_clear_workbook_protection(workbook: &Arc<Mutex<CoreWorkbook>>) -> PyResult<()> {
    let mut wb = lock_wb(workbook)?;
    wb.clear_workbook_protection();
    Ok(())
}

// =============================================================================
// Registration
// =============================================================================

/// Register all protection PyO3 types with the native module.
pub(super) fn register(module: &Bound<'_, PyModule>) -> PyResult<()> {
    module.add_class::<XlsxSheetProtection>()?;
    module.add_class::<XlsxWorkbookProtection>()?;
    Ok(())
}
