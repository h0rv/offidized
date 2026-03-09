//! Python bindings for the offidized-ir text IR derive/apply workflow.

use pyo3::prelude::*;
use pyo3::types::{PyDict, PyList};

use crate::error::value_error;

fn ir_error_to_py(error: offidized_ir::IrError) -> PyErr {
    value_error(error.to_string())
}

// =============================================================================
// Module-level functions
// =============================================================================

/// Derive IR text from an Office file.
#[pyfunction]
#[pyo3(signature = (path, mode="content", sheet=None, range=None))]
fn ir_derive(
    path: &str,
    mode: &str,
    sheet: Option<String>,
    range: Option<String>,
) -> PyResult<String> {
    let mode = offidized_ir::Mode::parse_str(mode).map_err(ir_error_to_py)?;
    let options = offidized_ir::DeriveOptions { mode, sheet, range };
    offidized_ir::derive(std::path::Path::new(path), options).map_err(ir_error_to_py)
}

/// Apply IR text to produce an output file.
#[pyfunction]
#[pyo3(signature = (ir, output, source_override=None, force=false))]
fn ir_apply(
    py: Python<'_>,
    ir: &str,
    output: &str,
    source_override: Option<String>,
    force: bool,
) -> PyResult<Py<PyAny>> {
    let options = offidized_ir::ApplyOptions {
        source_override: source_override.map(std::path::PathBuf::from),
        force,
    };
    let result =
        offidized_ir::apply(ir, std::path::Path::new(output), &options).map_err(ir_error_to_py)?;
    apply_result_to_py(py, &result)
}

/// Derive IR from in-memory bytes.
#[pyfunction]
#[pyo3(signature = (bytes, source_name, mode="content", sheet=None, range=None))]
fn ir_derive_from_bytes(
    bytes: &[u8],
    source_name: &str,
    mode: &str,
    sheet: Option<String>,
    range: Option<String>,
) -> PyResult<String> {
    let mode = offidized_ir::Mode::parse_str(mode).map_err(ir_error_to_py)?;
    let options = offidized_ir::DeriveOptions { mode, sheet, range };
    offidized_ir::derive_from_bytes(bytes, source_name, options).map_err(ir_error_to_py)
}

/// Apply IR to in-memory bytes, returning `(output_bytes, result_dict)`.
#[pyfunction]
#[pyo3(signature = (source_bytes, ir, source_override=None, force=false))]
fn ir_apply_to_bytes(
    py: Python<'_>,
    source_bytes: &[u8],
    ir: &str,
    source_override: Option<String>,
    force: bool,
) -> PyResult<(Vec<u8>, Py<PyAny>)> {
    let options = offidized_ir::ApplyOptions {
        source_override: source_override.map(std::path::PathBuf::from),
        force,
    };
    let (bytes, result) =
        offidized_ir::apply_to_bytes(source_bytes, ir, &options).map_err(ir_error_to_py)?;
    let result_dict = apply_result_to_py(py, &result)?;
    Ok((bytes, result_dict))
}

// =============================================================================
// Unified Document API
// =============================================================================

/// High-level unified document for structured node access and editing.
#[pyclass(module = "offidized._native", name = "UnifiedDocument")]
pub struct PyUnifiedDocument {
    inner: offidized_ir::UnifiedDocument,
}

#[pymethods]
impl PyUnifiedDocument {
    /// Derive a unified document from an Office file.
    #[staticmethod]
    #[pyo3(signature = (path, mode="content", sheet=None, range=None))]
    fn derive(
        path: &str,
        mode: &str,
        sheet: Option<String>,
        range: Option<String>,
    ) -> PyResult<Self> {
        let mode = offidized_ir::Mode::parse_str(mode).map_err(ir_error_to_py)?;
        let options = offidized_ir::UnifiedDeriveOptions { mode, sheet, range };
        let doc = offidized_ir::UnifiedDocument::derive(std::path::Path::new(path), options)
            .map_err(ir_error_to_py)?;
        Ok(Self { inner: doc })
    }

    /// Parse a unified document from an IR string.
    #[staticmethod]
    fn from_ir(ir: &str) -> PyResult<Self> {
        let doc = offidized_ir::UnifiedDocument::from_ir(ir).map_err(ir_error_to_py)?;
        Ok(Self { inner: doc })
    }

    /// Serialize back to IR text.
    fn to_ir(&self) -> String {
        self.inner.to_ir()
    }

    /// Return all content nodes as a list of dicts `{id, kind, text}`.
    fn nodes(&self, py: Python<'_>) -> PyResult<Py<PyAny>> {
        let nodes = self.inner.nodes();
        let list = PyList::empty(py);
        for node in nodes {
            let dict = PyDict::new(py);
            dict.set_item("id", node.id_string())?;
            dict.set_item("kind", format!("{:?}", node.kind))?;
            dict.set_item("text", &node.text)?;
            list.append(dict)?;
        }
        Ok(list.into())
    }

    /// Return document capabilities as a dict.
    fn capabilities(&self, py: Python<'_>) -> PyResult<Py<PyAny>> {
        let caps = self.inner.capabilities();
        let dict = PyDict::new(py);
        dict.set_item("text_nodes", caps.text_nodes)?;
        dict.set_item("table_cells", caps.table_cells)?;
        dict.set_item("chart_meta", caps.chart_meta)?;
        dict.set_item("style_nodes", caps.style_nodes)?;
        Ok(dict.into())
    }

    /// Apply a list of edits (each a dict `{id, text, group?}`).
    ///
    /// Returns an edit report dict `{requested, applied, skipped, diagnostics}`.
    fn apply_edits(&mut self, py: Python<'_>, edits: Vec<PyEditDict>) -> PyResult<Py<PyAny>> {
        let edits: Vec<offidized_ir::UnifiedEdit> = edits.into_iter().map(Into::into).collect();
        let report = self.inner.apply_edits(&edits).map_err(ir_error_to_py)?;
        edit_report_to_py(py, &report)
    }

    /// Lint edits without applying. Returns a list of diagnostic dicts.
    fn lint_edits(&self, py: Python<'_>, edits: Vec<PyEditDict>) -> PyResult<Py<PyAny>> {
        let edits: Vec<offidized_ir::UnifiedEdit> = edits.into_iter().map(Into::into).collect();
        let diags = self.inner.lint_edits(&edits);
        diagnostics_to_py(py, &diags)
    }

    /// Save the document to a file, returning an apply result dict.
    #[pyo3(signature = (output, source_override=None, force=false))]
    fn save_as(
        &self,
        py: Python<'_>,
        output: &str,
        source_override: Option<String>,
        force: bool,
    ) -> PyResult<Py<PyAny>> {
        let options = offidized_ir::ApplyOptions {
            source_override: source_override.map(std::path::PathBuf::from),
            force,
        };
        let result = self
            .inner
            .save_as(std::path::Path::new(output), &options)
            .map_err(ir_error_to_py)?;
        apply_result_to_py(py, &result)
    }
}

// =============================================================================
// Edit dict helper
// =============================================================================

/// A dict-like input for `UnifiedEdit`, extracted from Python kwargs.
#[derive(FromPyObject)]
pub struct PyEditDict {
    id: String,
    text: String,
    #[pyo3(attribute("group"))]
    group: Option<String>,
}

impl From<PyEditDict> for offidized_ir::UnifiedEdit {
    fn from(d: PyEditDict) -> Self {
        let mut edit = offidized_ir::UnifiedEdit::new(d.id, d.text);
        if let Some(group) = d.group {
            edit = edit.with_group(group);
        }
        edit
    }
}

// =============================================================================
// Conversion helpers
// =============================================================================

fn apply_result_to_py(py: Python<'_>, result: &offidized_ir::ApplyResult) -> PyResult<Py<PyAny>> {
    let dict = PyDict::new(py);
    dict.set_item("cells_updated", result.cells_updated)?;
    dict.set_item("cells_created", result.cells_created)?;
    dict.set_item("cells_cleared", result.cells_cleared)?;
    dict.set_item("charts_added", result.charts_added)?;
    dict.set_item("warnings", &result.warnings)?;
    Ok(dict.into())
}

fn edit_report_to_py(
    py: Python<'_>,
    report: &offidized_ir::UnifiedEditReport,
) -> PyResult<Py<PyAny>> {
    let dict = PyDict::new(py);
    dict.set_item("requested", report.requested)?;
    dict.set_item("applied", report.applied)?;
    dict.set_item("skipped", report.skipped)?;
    dict.set_item("diagnostics", diagnostics_to_py(py, &report.diagnostics)?)?;
    Ok(dict.into())
}

fn diagnostics_to_py(
    py: Python<'_>,
    diags: &[offidized_ir::UnifiedDiagnostic],
) -> PyResult<Py<PyAny>> {
    let list = PyList::empty(py);
    for d in diags {
        let dict = PyDict::new(py);
        dict.set_item(
            "severity",
            match d.severity {
                offidized_ir::UnifiedDiagnosticSeverity::Error => "error",
                offidized_ir::UnifiedDiagnosticSeverity::Warning => "warning",
            },
        )?;
        dict.set_item("code", &d.code)?;
        dict.set_item("message", &d.message)?;
        dict.set_item("id", &d.id)?;
        list.append(dict)?;
    }
    Ok(list.into())
}

// =============================================================================
// Module Registration
// =============================================================================

pub(crate) fn register(module: &Bound<'_, PyModule>) -> PyResult<()> {
    module.add_function(wrap_pyfunction!(ir_derive, module)?)?;
    module.add_function(wrap_pyfunction!(ir_apply, module)?)?;
    module.add_function(wrap_pyfunction!(ir_derive_from_bytes, module)?)?;
    module.add_function(wrap_pyfunction!(ir_apply_to_bytes, module)?)?;
    module.add_class::<PyUnifiedDocument>()?;
    Ok(())
}
