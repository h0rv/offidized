//! Python bindings for data validation types from `offidized_xlsx`.
//!
//! Wraps [`DataValidation`] with a PyO3 class that mirrors the core Rust API.
//! Worksheet helper functions (`ws_*`) are called from the parent `Worksheet`
//! `#[pymethods]` block.

use super::lock_wb;
use crate::error::{value_error, xlsx_error_to_py};
use offidized_xlsx::{
    DataValidation as CoreDataValidation, DataValidationErrorStyle as CoreDataValidationErrorStyle,
    DataValidationType as CoreDataValidationType, Workbook as CoreWorkbook,
};
use pyo3::prelude::*;
use std::sync::{Arc, Mutex};

// =============================================================================
// Helpers
// =============================================================================

fn validation_type_to_str(vt: CoreDataValidationType) -> &'static str {
    match vt {
        CoreDataValidationType::List => "list",
        CoreDataValidationType::Whole => "whole",
        CoreDataValidationType::Decimal => "decimal",
        CoreDataValidationType::Date => "date",
        CoreDataValidationType::TextLength => "textLength",
        CoreDataValidationType::Custom => "custom",
        CoreDataValidationType::Time => "time",
    }
}

fn error_style_to_str(es: CoreDataValidationErrorStyle) -> &'static str {
    match es {
        CoreDataValidationErrorStyle::Stop => "stop",
        CoreDataValidationErrorStyle::Warning => "warning",
        CoreDataValidationErrorStyle::Information => "information",
    }
}

fn str_to_error_style(s: &str) -> PyResult<CoreDataValidationErrorStyle> {
    match s.to_lowercase().as_str() {
        "stop" => Ok(CoreDataValidationErrorStyle::Stop),
        "warning" => Ok(CoreDataValidationErrorStyle::Warning),
        "information" => Ok(CoreDataValidationErrorStyle::Information),
        _ => Err(value_error(format!(
            "Unknown error style '{s}': expected 'stop', 'warning', or 'information'"
        ))),
    }
}

// =============================================================================
// XlsxDataValidation
// =============================================================================

/// Data validation rule applied to one or more worksheet ranges.
#[pyclass(
    module = "offidized._native",
    name = "XlsxDataValidation",
    from_py_object
)]
#[derive(Clone)]
pub struct XlsxDataValidation {
    inner: CoreDataValidation,
}

impl XlsxDataValidation {
    pub(super) fn from_core(core: CoreDataValidation) -> Self {
        Self { inner: core }
    }

    pub(super) fn into_core(self) -> CoreDataValidation {
        self.inner
    }
}

#[pymethods]
impl XlsxDataValidation {
    /// Create a new data validation rule.
    ///
    /// `validation_type` must be one of: "list", "whole", "decimal", "date",
    /// "textLength", "custom", "time". `ranges` is a list of A1-notation range
    /// strings (e.g. ["A1:B10"]). `formula1` is the primary formula or value.
    #[new]
    pub fn new(validation_type: &str, ranges: Vec<String>, formula1: &str) -> PyResult<Self> {
        let refs: Vec<&str> = ranges.iter().map(|s| s.as_str()).collect();
        let core = match validation_type.to_lowercase().as_str() {
            "list" => CoreDataValidation::list(refs, formula1).map_err(xlsx_error_to_py)?,
            "whole" => CoreDataValidation::whole(refs, formula1).map_err(xlsx_error_to_py)?,
            "decimal" => CoreDataValidation::decimal(refs, formula1).map_err(xlsx_error_to_py)?,
            "date" => CoreDataValidation::date(refs, formula1).map_err(xlsx_error_to_py)?,
            "textlength" | "textLength" => {
                CoreDataValidation::text_length(refs, formula1).map_err(xlsx_error_to_py)?
            }
            "custom" => CoreDataValidation::custom(refs, formula1).map_err(xlsx_error_to_py)?,
            "time" => CoreDataValidation::time(refs, formula1).map_err(xlsx_error_to_py)?,
            _ => {
                return Err(value_error(format!(
                    "Unknown validation type '{validation_type}': expected one of \
                     'list', 'whole', 'decimal', 'date', 'textLength', 'custom', 'time'"
                )));
            }
        };
        Ok(Self { inner: core })
    }

    /// Return the validation type as a string.
    #[getter]
    pub fn validation_type(&self) -> &str {
        validation_type_to_str(self.inner.validation_type())
    }

    /// Return the sqref ranges as a list of "start:end" strings.
    #[getter]
    pub fn ranges(&self) -> Vec<String> {
        self.inner
            .sqref()
            .iter()
            .map(|r| format!("{}:{}", r.start(), r.end()))
            .collect()
    }

    /// Return the primary formula or value.
    #[getter]
    pub fn formula1(&self) -> &str {
        self.inner.formula1()
    }

    /// Return the secondary formula or value, or None.
    #[getter]
    pub fn formula2(&self) -> Option<String> {
        self.inner.formula2().map(|s| s.to_string())
    }

    /// Set the secondary formula or value.
    #[setter]
    pub fn set_formula2(&mut self, value: Option<String>) {
        if let Some(v) = value {
            self.inner.set_formula2(v);
        } else {
            self.inner.clear_formula2();
        }
    }

    /// Return the error dialog style as a string ("stop", "warning", "information"), or None.
    #[getter]
    pub fn error_style(&self) -> Option<String> {
        self.inner
            .error_style()
            .map(|s| error_style_to_str(s).to_string())
    }

    /// Set the error dialog style ("stop", "warning", "information").
    #[setter]
    pub fn set_error_style(&mut self, value: Option<String>) -> PyResult<()> {
        match value {
            None => {
                self.inner.clear_error_style();
            }
            Some(ref s) => {
                let style = str_to_error_style(s)?;
                self.inner.set_error_style(style);
            }
        }
        Ok(())
    }

    /// Return the error dialog title, or None.
    #[getter]
    pub fn error_title(&self) -> Option<String> {
        self.inner.error_title().map(|s| s.to_string())
    }

    /// Set the error dialog title.
    #[setter]
    pub fn set_error_title(&mut self, value: Option<String>) {
        if let Some(v) = value {
            self.inner.set_error_title(v);
        } else {
            self.inner.clear_error_title();
        }
    }

    /// Return the error dialog message, or None.
    #[getter]
    pub fn error_message(&self) -> Option<String> {
        self.inner.error_message().map(|s| s.to_string())
    }

    /// Set the error dialog message.
    #[setter]
    pub fn set_error_message(&mut self, value: Option<String>) {
        if let Some(v) = value {
            self.inner.set_error_message(v);
        } else {
            self.inner.clear_error_message();
        }
    }

    /// Return the input prompt title, or None.
    #[getter]
    pub fn prompt_title(&self) -> Option<String> {
        self.inner.prompt_title().map(|s| s.to_string())
    }

    /// Set the input prompt title.
    #[setter]
    pub fn set_prompt_title(&mut self, value: Option<String>) {
        if let Some(v) = value {
            self.inner.set_prompt_title(v);
        } else {
            self.inner.clear_prompt_title();
        }
    }

    /// Return the input prompt message, or None.
    #[getter]
    pub fn prompt_message(&self) -> Option<String> {
        self.inner.prompt_message().map(|s| s.to_string())
    }

    /// Set the input prompt message.
    #[setter]
    pub fn set_prompt_message(&mut self, value: Option<String>) {
        if let Some(v) = value {
            self.inner.set_prompt_message(v);
        } else {
            self.inner.clear_prompt_message();
        }
    }

    /// Return whether the input message prompt is shown when the cell is selected, or None.
    #[getter]
    pub fn show_input_message(&self) -> Option<bool> {
        self.inner.show_input_message()
    }

    /// Set whether to show the input message prompt.
    #[setter]
    pub fn set_show_input_message(&mut self, value: Option<bool>) {
        if let Some(v) = value {
            self.inner.set_show_input_message(v);
        } else {
            self.inner.clear_show_input_message();
        }
    }

    /// Return whether the error message dialog is shown for invalid data, or None.
    #[getter]
    pub fn show_error_message(&self) -> Option<bool> {
        self.inner.show_error_message()
    }

    /// Set whether to show the error message dialog.
    #[setter]
    pub fn set_show_error_message(&mut self, value: Option<bool>) {
        if let Some(v) = value {
            self.inner.set_show_error_message(v);
        } else {
            self.inner.clear_show_error_message();
        }
    }

    /// Add an additional range (e.g. "C1:C10") to this validation rule.
    pub fn add_range(&mut self, range: &str) -> PyResult<()> {
        self.inner.add_range(range).map_err(xlsx_error_to_py)?;
        Ok(())
    }

    // -- static constructors --------------------------------------------------

    /// Create a list validation rule.
    ///
    /// Args:
    ///     ranges: List of A1-notation range strings (e.g. ``["A1:A10"]``).
    ///     formula1: The list source formula (e.g. ``'"Yes,No"'``).
    #[staticmethod]
    pub fn list(ranges: Vec<String>, formula1: &str) -> PyResult<Self> {
        let refs: Vec<&str> = ranges.iter().map(|s| s.as_str()).collect();
        let core = CoreDataValidation::list(refs, formula1).map_err(xlsx_error_to_py)?;
        Ok(Self { inner: core })
    }

    /// Create a whole-number validation rule.
    #[staticmethod]
    pub fn whole(ranges: Vec<String>, formula1: &str) -> PyResult<Self> {
        let refs: Vec<&str> = ranges.iter().map(|s| s.as_str()).collect();
        let core = CoreDataValidation::whole(refs, formula1).map_err(xlsx_error_to_py)?;
        Ok(Self { inner: core })
    }

    /// Create a decimal validation rule.
    #[staticmethod]
    pub fn decimal(ranges: Vec<String>, formula1: &str) -> PyResult<Self> {
        let refs: Vec<&str> = ranges.iter().map(|s| s.as_str()).collect();
        let core = CoreDataValidation::decimal(refs, formula1).map_err(xlsx_error_to_py)?;
        Ok(Self { inner: core })
    }

    /// Create a date validation rule.
    #[staticmethod]
    pub fn date(ranges: Vec<String>, formula1: &str) -> PyResult<Self> {
        let refs: Vec<&str> = ranges.iter().map(|s| s.as_str()).collect();
        let core = CoreDataValidation::date(refs, formula1).map_err(xlsx_error_to_py)?;
        Ok(Self { inner: core })
    }

    /// Create a text-length validation rule.
    #[staticmethod]
    pub fn text_length(ranges: Vec<String>, formula1: &str) -> PyResult<Self> {
        let refs: Vec<&str> = ranges.iter().map(|s| s.as_str()).collect();
        let core = CoreDataValidation::text_length(refs, formula1).map_err(xlsx_error_to_py)?;
        Ok(Self { inner: core })
    }

    /// Create a custom formula validation rule.
    #[staticmethod]
    pub fn custom(ranges: Vec<String>, formula1: &str) -> PyResult<Self> {
        let refs: Vec<&str> = ranges.iter().map(|s| s.as_str()).collect();
        let core = CoreDataValidation::custom(refs, formula1).map_err(xlsx_error_to_py)?;
        Ok(Self { inner: core })
    }

    /// Create a time validation rule.
    #[staticmethod]
    pub fn time(ranges: Vec<String>, formula1: &str) -> PyResult<Self> {
        let refs: Vec<&str> = ranges.iter().map(|s| s.as_str()).collect();
        let core = CoreDataValidation::time(refs, formula1).map_err(xlsx_error_to_py)?;
        Ok(Self { inner: core })
    }
}

// =============================================================================
// Worksheet helper functions
// =============================================================================

pub(super) fn ws_data_validations(
    wb: &Arc<Mutex<CoreWorkbook>>,
    sheet_name: &str,
) -> PyResult<Vec<XlsxDataValidation>> {
    let wb = lock_wb(wb)?;
    let ws = wb
        .sheet(sheet_name)
        .ok_or_else(|| value_error(format!("worksheet '{sheet_name}' not found")))?;
    Ok(ws
        .data_validations()
        .iter()
        .cloned()
        .map(XlsxDataValidation::from_core)
        .collect())
}

pub(super) fn ws_add_data_validation(
    wb: &Arc<Mutex<CoreWorkbook>>,
    sheet_name: &str,
    dv: XlsxDataValidation,
) -> PyResult<()> {
    let mut wb = lock_wb(wb)?;
    let ws = wb
        .sheet_mut(sheet_name)
        .ok_or_else(|| value_error(format!("worksheet '{sheet_name}' not found")))?;
    ws.add_data_validation(dv.into_core());
    Ok(())
}

pub(super) fn ws_clear_data_validations(
    wb: &Arc<Mutex<CoreWorkbook>>,
    sheet_name: &str,
) -> PyResult<()> {
    let mut wb = lock_wb(wb)?;
    let ws = wb
        .sheet_mut(sheet_name)
        .ok_or_else(|| value_error(format!("worksheet '{sheet_name}' not found")))?;
    ws.clear_data_validations();
    Ok(())
}

// =============================================================================
// Registration
// =============================================================================

pub(super) fn register(module: &Bound<'_, PyModule>) -> PyResult<()> {
    module.add_class::<XlsxDataValidation>()?;
    Ok(())
}
