//! Python bindings for sparkline types from `offidized_xlsx`.
//!
//! Wraps [`SparklineGroup`] and [`Sparkline`] with PyO3 classes that mirror
//! the core Rust API. Worksheet helper functions (`ws_*`) are called from the
//! parent `Worksheet` `#[pymethods]` block.

use super::lock_wb;
use crate::error::value_error;
use offidized_xlsx::{
    Sparkline as CoreSparkline, SparklineGroup as CoreSparklineGroup,
    SparklineType as CoreSparklineType, Workbook as CoreWorkbook,
};
use pyo3::prelude::*;
use std::sync::{Arc, Mutex};

// =============================================================================
// String converters for SparklineType
// =============================================================================

fn sparkline_type_to_str(t: CoreSparklineType) -> &'static str {
    match t {
        CoreSparklineType::Line => "line",
        CoreSparklineType::Column => "column",
        CoreSparklineType::Stacked => "stacked",
    }
}

fn str_to_sparkline_type(s: &str) -> PyResult<CoreSparklineType> {
    match s.to_lowercase().as_str() {
        "line" => Ok(CoreSparklineType::Line),
        "column" => Ok(CoreSparklineType::Column),
        "stacked" => Ok(CoreSparklineType::Stacked),
        _ => Err(value_error(format!(
            "Unknown sparkline type '{s}': expected 'line', 'column', or 'stacked'"
        ))),
    }
}

// =============================================================================
// XlsxSparkline
// =============================================================================

/// Python wrapper for a single sparkline mapping a data range to a cell.
///
/// A sparkline renders a compact chart inside one cell. The ``location``
/// is the cell that displays the sparkline (e.g. ``"Sheet1!A1"``) and
/// ``data_range`` is the range of values it visualises (e.g.
/// ``"Sheet1!B1:B10"``).
#[pyclass(module = "offidized._native", name = "XlsxSparkline")]
#[derive(Clone)]
pub struct XlsxSparkline {
    inner: CoreSparkline,
}

impl XlsxSparkline {
    pub(super) fn from_core(sparkline: CoreSparkline) -> Self {
        Self { inner: sparkline }
    }

    pub(super) fn into_core(self) -> CoreSparkline {
        self.inner
    }
}

#[pymethods]
impl XlsxSparkline {
    /// Create a new sparkline.
    ///
    /// Args:
    ///     location: Cell reference where the sparkline is rendered
    ///         (e.g. ``"Sheet1!A1"``).
    ///     data_range: Data range the sparkline visualises
    ///         (e.g. ``"Sheet1!B1:B10"``).
    #[new]
    pub fn new(location: &str, data_range: &str) -> Self {
        Self {
            inner: CoreSparkline::new(location, data_range),
        }
    }

    /// The cell location where the sparkline is rendered.
    #[getter]
    pub fn location(&self) -> &str {
        self.inner.location()
    }

    /// Set the cell location.
    #[setter]
    pub fn set_location(&mut self, location: String) {
        self.inner.set_location(location);
    }

    /// The data range the sparkline visualises.
    #[getter]
    pub fn data_range(&self) -> &str {
        self.inner.data_range()
    }

    /// Set the data range.
    #[setter]
    pub fn set_data_range(&mut self, data_range: String) {
        self.inner.set_data_range(data_range);
    }
}

// =============================================================================
// XlsxSparklineGroup
// =============================================================================

/// Python wrapper for a sparkline group.
///
/// A sparkline group collects one or more :py:class:`XlsxSparkline` entries
/// that share a common type, colour scheme, and display options. Add groups
/// to a worksheet via :py:meth:`Worksheet.add_sparkline_group`.
#[pyclass(module = "offidized._native", name = "XlsxSparklineGroup")]
#[derive(Clone)]
pub struct XlsxSparklineGroup {
    inner: CoreSparklineGroup,
}

impl XlsxSparklineGroup {
    pub(super) fn from_core(group: CoreSparklineGroup) -> Self {
        Self { inner: group }
    }

    pub(super) fn into_core(self) -> CoreSparklineGroup {
        self.inner
    }
}

#[pymethods]
impl XlsxSparklineGroup {
    /// Create a new sparkline group with default settings (line type, no markers).
    #[new]
    pub fn new() -> Self {
        Self {
            inner: CoreSparklineGroup::new(),
        }
    }

    /// The sparkline type: ``"line"``, ``"column"``, or ``"stacked"``.
    #[getter]
    pub fn sparkline_type(&self) -> &str {
        sparkline_type_to_str(self.inner.sparkline_type())
    }

    /// Set the sparkline type.
    ///
    /// Accepted values: ``"line"``, ``"column"``, ``"stacked"``.
    #[setter]
    pub fn set_sparkline_type(&mut self, value: &str) -> PyResult<()> {
        let t = str_to_sparkline_type(value)?;
        self.inner.set_sparkline_type(t);
        Ok(())
    }

    /// Whether data-point markers are shown (line sparklines only).
    #[getter]
    pub fn markers(&self) -> bool {
        self.inner.markers()
    }

    /// Set whether data-point markers are shown.
    #[setter]
    pub fn set_markers(&mut self, value: bool) {
        self.inner.set_markers(value);
    }

    /// Whether the highest data point is highlighted.
    #[getter]
    pub fn high_point(&self) -> bool {
        self.inner.high_point()
    }

    /// Set whether the highest data point is highlighted.
    #[setter]
    pub fn set_high_point(&mut self, value: bool) {
        self.inner.set_high_point(value);
    }

    /// Whether the lowest data point is highlighted.
    #[getter]
    pub fn low_point(&self) -> bool {
        self.inner.low_point()
    }

    /// Set whether the lowest data point is highlighted.
    #[setter]
    pub fn set_low_point(&mut self, value: bool) {
        self.inner.set_low_point(value);
    }

    /// Whether the first data point is highlighted.
    #[getter]
    pub fn first_point(&self) -> bool {
        self.inner.first_point()
    }

    /// Set whether the first data point is highlighted.
    #[setter]
    pub fn set_first_point(&mut self, value: bool) {
        self.inner.set_first_point(value);
    }

    /// Whether the last data point is highlighted.
    #[getter]
    pub fn last_point(&self) -> bool {
        self.inner.last_point()
    }

    /// Set whether the last data point is highlighted.
    #[setter]
    pub fn set_last_point(&mut self, value: bool) {
        self.inner.set_last_point(value);
    }

    /// Whether negative data points are highlighted.
    #[getter]
    pub fn negative_points(&self) -> bool {
        self.inner.negative_points()
    }

    /// Set whether negative data points are highlighted.
    #[setter]
    pub fn set_negative_points(&mut self, value: bool) {
        self.inner.set_negative_points(value);
    }

    /// Line weight in points, or None to use the default (0.75 pt).
    #[getter]
    pub fn line_weight(&self) -> Option<f64> {
        self.inner.line_weight()
    }

    /// Set the line weight in points. Pass None to clear it.
    #[setter]
    pub fn set_line_weight(&mut self, value: Option<f64>) {
        match value {
            Some(v) => {
                self.inner.set_line_weight(v);
            }
            None => {
                // Clearing is not exposed on the core type; set a sentinel that
                // matches the OOXML default so serialization omits the attribute.
                self.inner.set_line_weight(0.75);
            }
        }
    }

    /// Manual minimum axis value, or None for automatic scaling.
    #[getter]
    pub fn manual_min(&self) -> Option<f64> {
        self.inner.manual_min()
    }

    /// Set the manual minimum axis value. Pass None to clear it.
    #[setter]
    pub fn set_manual_min(&mut self, value: Option<f64>) {
        if let Some(v) = value {
            self.inner.set_manual_min(v);
        }
    }

    /// Manual maximum axis value, or None for automatic scaling.
    #[getter]
    pub fn manual_max(&self) -> Option<f64> {
        self.inner.manual_max()
    }

    /// Set the manual maximum axis value. Pass None to clear it.
    #[setter]
    pub fn set_manual_max(&mut self, value: Option<f64>) {
        if let Some(v) = value {
            self.inner.set_manual_max(v);
        }
    }

    /// Return all sparklines in the group as a list of :py:class:`XlsxSparkline` objects.
    pub fn sparklines(&self) -> Vec<XlsxSparkline> {
        self.inner
            .sparklines()
            .iter()
            .cloned()
            .map(XlsxSparkline::from_core)
            .collect()
    }

    /// Append a sparkline to this group.
    pub fn add_sparkline(&mut self, sparkline: &XlsxSparkline) {
        self.inner.add_sparkline(sparkline.inner.clone());
    }

    /// Remove all sparklines from this group.
    pub fn clear_sparklines(&mut self) {
        self.inner.clear_sparklines();
    }
}

// =============================================================================
// Worksheet helper functions
// =============================================================================

/// Return all sparkline groups on the worksheet as a list of
/// :py:class:`XlsxSparklineGroup` objects.
pub(super) fn ws_sparkline_groups(
    workbook: &Arc<Mutex<CoreWorkbook>>,
    name_key: &str,
) -> PyResult<Vec<XlsxSparklineGroup>> {
    let wb = lock_wb(workbook)?;
    let ws = wb
        .sheet(name_key)
        .ok_or_else(|| value_error(format!("worksheet '{name_key}' not found")))?;
    Ok(ws
        .sparkline_groups()
        .iter()
        .cloned()
        .map(XlsxSparklineGroup::from_core)
        .collect())
}

/// Add a sparkline group to the worksheet.
pub(super) fn ws_add_sparkline_group(
    workbook: &Arc<Mutex<CoreWorkbook>>,
    name_key: &str,
    group: XlsxSparklineGroup,
) -> PyResult<()> {
    let mut wb = lock_wb(workbook)?;
    let ws = wb
        .sheet_mut(name_key)
        .ok_or_else(|| value_error(format!("worksheet '{name_key}' not found")))?;
    ws.add_sparkline_group(group.into_core());
    Ok(())
}

/// Remove all sparkline groups from the worksheet.
pub(super) fn ws_clear_sparkline_groups(
    workbook: &Arc<Mutex<CoreWorkbook>>,
    name_key: &str,
) -> PyResult<()> {
    let mut wb = lock_wb(workbook)?;
    let ws = wb
        .sheet_mut(name_key)
        .ok_or_else(|| value_error(format!("worksheet '{name_key}' not found")))?;
    ws.clear_sparkline_groups();
    Ok(())
}

// =============================================================================
// Registration
// =============================================================================

/// Register all sparkline PyO3 types with the native module.
pub(super) fn register(module: &Bound<'_, PyModule>) -> PyResult<()> {
    module.add_class::<XlsxSparkline>()?;
    module.add_class::<XlsxSparklineGroup>()?;
    Ok(())
}
