//! Python bindings for workbook lint types.
//!
//! Exposes [`XlsxLintFinding`] and [`XlsxLintReport`] as well as the
//! workbook-level helper [`wb_lint`] that drives [`WorkbookLintBuilder`].

use super::lock_wb;
use crate::error::value_error;
use offidized_xlsx::{LintFinding, LintReport, LintSeverity};
use pyo3::prelude::*;
use std::sync::{Arc, Mutex};

use offidized_xlsx::Workbook as CoreWorkbook;

// =============================================================================
// XlsxLintFinding
// =============================================================================

/// A single lint finding produced by the workbook linter.
///
/// Fields mirror [`offidized_xlsx::LintFinding`] but are flattened for
/// ergonomic Python access.
#[pyclass(module = "offidized._native", name = "XlsxLintFinding", from_py_object)]
#[derive(Clone)]
pub struct XlsxLintFinding {
    /// Severity string: `"error"`, `"warning"`, or `"info"`.
    pub severity: String,
    /// Machine-readable code identifying the rule that fired (e.g. `"broken_sheet_ref"`).
    pub code: String,
    /// Human-readable description of the problem.
    pub message: String,
    /// Worksheet name, if applicable.
    pub sheet: Option<String>,
    /// Cell reference (e.g. `"A1"`), if applicable.
    pub cell: Option<String>,
    /// Named object (e.g. pivot table name, defined name), if applicable.
    pub object: Option<String>,
}

impl XlsxLintFinding {
    /// Convert from the core [`LintFinding`] type.
    fn from_core(finding: &LintFinding) -> Self {
        let severity = match finding.severity {
            LintSeverity::Error => "error",
            LintSeverity::Warning => "warning",
            LintSeverity::Info => "info",
        }
        .to_string();
        Self {
            severity,
            code: finding.code.clone(),
            message: finding.message.clone(),
            sheet: finding.location.sheet.clone(),
            cell: finding.location.cell.clone(),
            object: finding.location.object.clone(),
        }
    }
}

#[pymethods]
impl XlsxLintFinding {
    /// Severity of the finding: `"error"`, `"warning"`, or `"info"`.
    #[getter]
    pub fn severity(&self) -> &str {
        &self.severity
    }

    /// Machine-readable rule code (e.g. `"broken_sheet_ref"`).
    #[getter]
    pub fn code(&self) -> &str {
        &self.code
    }

    /// Human-readable description of the problem.
    #[getter]
    pub fn message(&self) -> &str {
        &self.message
    }

    /// Worksheet name where the finding occurred, or `None`.
    #[getter]
    pub fn sheet(&self) -> Option<&str> {
        self.sheet.as_deref()
    }

    /// Cell reference where the finding occurred (e.g. `"A1"`), or `None`.
    #[getter]
    pub fn cell(&self) -> Option<&str> {
        self.cell.as_deref()
    }

    /// Named object associated with the finding (e.g. pivot table or defined
    /// name), or `None`.
    #[getter]
    pub fn object(&self) -> Option<&str> {
        self.object.as_deref()
    }

    fn __repr__(&self) -> String {
        format!(
            "XlsxLintFinding(severity={:?}, code={:?}, message={:?})",
            self.severity, self.code, self.message,
        )
    }
}

// =============================================================================
// XlsxLintReport
// =============================================================================

/// The result of running the workbook linter.
///
/// Wraps a collection of [`XlsxLintFinding`] values and provides summary
/// helpers for counting errors and warnings.
#[pyclass(module = "offidized._native", name = "XlsxLintReport")]
pub struct XlsxLintReport {
    findings: Vec<XlsxLintFinding>,
}

impl XlsxLintReport {
    /// Build from the core [`LintReport`].
    fn from_core(report: LintReport) -> Self {
        let findings = report
            .findings()
            .iter()
            .map(XlsxLintFinding::from_core)
            .collect();
        Self { findings }
    }
}

#[pymethods]
impl XlsxLintReport {
    /// Return all findings as a list of [`XlsxLintFinding`].
    pub fn findings(&self) -> Vec<XlsxLintFinding> {
        self.findings.clone()
    }

    /// Return the number of findings with severity `"error"`.
    pub fn error_count(&self) -> usize {
        self.findings
            .iter()
            .filter(|f| f.severity == "error")
            .count()
    }

    /// Return the number of findings with severity `"warning"`.
    pub fn warning_count(&self) -> usize {
        self.findings
            .iter()
            .filter(|f| f.severity == "warning")
            .count()
    }

    /// Return `True` when there are no findings at all.
    pub fn is_clean(&self) -> bool {
        self.findings.is_empty()
    }

    fn __repr__(&self) -> String {
        format!(
            "XlsxLintReport(findings={}, errors={}, warnings={})",
            self.findings.len(),
            self.error_count(),
            self.warning_count(),
        )
    }
}

// =============================================================================
// wb_lint helper
// =============================================================================

/// Run the workbook linter with the specified checks and return a report.
///
/// `checks` is a list of check names to enable. Passing an empty list runs
/// all available checks. Valid names are:
///
/// - `"broken_refs"` — detect formula references to missing sheets
/// - `"formula_consistency"` — detect cells whose cached value differs from the computed value
/// - `"pivot_sources"` — detect pivot tables with invalid or missing source ranges/fields
/// - `"named_ranges"` — detect defined names that reference missing sheets or invalid ranges
pub(super) fn wb_lint(
    workbook: &Arc<Mutex<CoreWorkbook>>,
    checks: Vec<String>,
) -> PyResult<XlsxLintReport> {
    let wb = lock_wb(workbook)?;

    let run_all = checks.is_empty();

    let mut builder = wb.lint();

    for check in &checks {
        match check.as_str() {
            "broken_refs" => {
                builder = builder.check_broken_refs();
            }
            "formula_consistency" => {
                builder = builder.check_formula_consistency();
            }
            "pivot_sources" => {
                builder = builder.check_pivot_sources();
            }
            "named_ranges" => {
                builder = builder.check_named_ranges();
            }
            other => {
                return Err(value_error(format!(
                    "Unknown lint check: {other:?}. Valid checks are: \
                     broken_refs, formula_consistency, pivot_sources, named_ranges"
                )));
            }
        }
    }

    if run_all {
        builder = builder
            .check_broken_refs()
            .check_formula_consistency()
            .check_pivot_sources()
            .check_named_ranges();
    }

    let report = builder.run();
    Ok(XlsxLintReport::from_core(report))
}

// =============================================================================
// Registration
// =============================================================================

/// Register lint types with the Python module.
pub(super) fn register(module: &Bound<'_, PyModule>) -> PyResult<()> {
    module.add_class::<XlsxLintFinding>()?;
    module.add_class::<XlsxLintReport>()?;
    Ok(())
}
