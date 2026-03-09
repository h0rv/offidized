//! Python bindings for chart types from `offidized_xlsx`.
//!
//! Wraps [`Chart`], [`ChartSeries`], [`ChartAxis`], [`ChartLegend`], and
//! [`ChartDataRef`] with PyO3 classes that mirror the core Rust API.
//! Worksheet helper functions (`ws_*`) are called from the parent
//! `Worksheet` `#[pymethods]` block.

use super::lock_wb;
use crate::error::value_error;
use offidized_xlsx::{
    BarDirection as CoreBarDirection, Chart as CoreChart, ChartAxis as CoreChartAxis,
    ChartDataRef as CoreChartDataRef, ChartGrouping as CoreChartGrouping,
    ChartLegend as CoreChartLegend, ChartSeries as CoreChartSeries, ChartType as CoreChartType,
    Workbook as CoreWorkbook,
};
use pyo3::prelude::*;
use std::sync::{Arc, Mutex};

// =============================================================================
// String converters
// =============================================================================

fn parse_chart_type(s: &str) -> PyResult<CoreChartType> {
    match s.to_lowercase().as_str() {
        "bar" | "barchart" => Ok(CoreChartType::Bar),
        "line" | "linechart" => Ok(CoreChartType::Line),
        "pie" | "piechart" => Ok(CoreChartType::Pie),
        "area" | "areachart" => Ok(CoreChartType::Area),
        "scatter" | "scatterchart" => Ok(CoreChartType::Scatter),
        "doughnut" | "doughnutchart" => Ok(CoreChartType::Doughnut),
        "radar" | "radarchart" => Ok(CoreChartType::Radar),
        "bubble" | "bubblechart" => Ok(CoreChartType::Bubble),
        "stock" | "stockchart" => Ok(CoreChartType::Stock),
        "surface" | "surfacechart" => Ok(CoreChartType::Surface),
        "combo" | "combochart" => Ok(CoreChartType::Combo),
        _ => Err(value_error(format!(
            "Unknown chart type '{s}': expected one of \
             'bar', 'line', 'pie', 'area', 'scatter', 'doughnut', \
             'radar', 'bubble', 'stock', 'surface', 'combo'"
        ))),
    }
}

fn parse_bar_direction(s: &str) -> PyResult<CoreBarDirection> {
    match s.to_lowercase().as_str() {
        "col" | "column" | "vertical" => Ok(CoreBarDirection::Column),
        "bar" | "horizontal" => Ok(CoreBarDirection::Bar),
        _ => Err(value_error(format!(
            "Unknown bar direction '{s}': expected 'col'/'column'/'vertical' or 'bar'/'horizontal'"
        ))),
    }
}

fn parse_grouping(s: &str) -> PyResult<CoreChartGrouping> {
    match s {
        "clustered" => Ok(CoreChartGrouping::Clustered),
        "stacked" => Ok(CoreChartGrouping::Stacked),
        "percentStacked" => Ok(CoreChartGrouping::PercentStacked),
        "standard" => Ok(CoreChartGrouping::Standard),
        _ => Err(value_error(format!(
            "Unknown grouping '{s}': expected 'clustered', 'stacked', 'percentStacked', or 'standard'"
        ))),
    }
}

// =============================================================================
// XlsxChartDataRef
// =============================================================================

/// Python wrapper for a chart data reference (formula + optional cached values).
#[pyclass(module = "offidized._native", name = "XlsxChartDataRef")]
#[derive(Clone)]
pub struct XlsxChartDataRef {
    inner: CoreChartDataRef,
}

impl XlsxChartDataRef {
    pub(super) fn from_core(data: CoreChartDataRef) -> Self {
        Self { inner: data }
    }

    pub(super) fn into_core(self) -> CoreChartDataRef {
        self.inner
    }
}

#[allow(clippy::new_without_default)]
#[pymethods]
impl XlsxChartDataRef {
    /// Create an empty data reference with no formula or cached values.
    #[new]
    pub fn new() -> Self {
        Self {
            inner: CoreChartDataRef::new(),
        }
    }

    /// Create a data reference from a cell range formula (e.g. ``"Sheet1!$A$1:$A$10"``).
    #[staticmethod]
    pub fn from_formula(formula: &str) -> Self {
        Self {
            inner: CoreChartDataRef::from_formula(formula),
        }
    }

    /// The cell range formula, or None if not set.
    #[getter]
    pub fn formula(&self) -> Option<String> {
        self.inner.formula().map(|s| s.to_string())
    }

    /// Set the cell range formula.
    #[setter]
    pub fn set_formula(&mut self, formula: Option<String>) {
        match formula {
            Some(f) => {
                self.inner.set_formula(f);
            }
            None => {
                self.inner.clear_formula();
            }
        }
    }

    /// The cached numeric values. ``None`` entries represent missing/blank cells.
    #[getter]
    pub fn num_values(&self) -> Vec<Option<f64>> {
        self.inner.num_values().to_vec()
    }

    /// Set the cached numeric values.
    #[setter]
    pub fn set_num_values(&mut self, values: Vec<Option<f64>>) {
        self.inner.set_num_values(values);
    }

    /// The cached string values.
    #[getter]
    pub fn str_values(&self) -> Vec<String> {
        self.inner.str_values().to_vec()
    }

    /// Set the cached string values.
    #[setter]
    pub fn set_str_values(&mut self, values: Vec<String>) {
        self.inner.set_str_values(values);
    }
}

// =============================================================================
// XlsxChartSeries
// =============================================================================

/// Python wrapper for a single chart data series.
#[pyclass(module = "offidized._native", name = "XlsxChartSeries")]
#[derive(Clone)]
pub struct XlsxChartSeries {
    inner: CoreChartSeries,
}

impl XlsxChartSeries {
    pub(super) fn from_core(series: CoreChartSeries) -> Self {
        Self { inner: series }
    }

    pub(super) fn into_core(self) -> CoreChartSeries {
        self.inner
    }
}

#[pymethods]
impl XlsxChartSeries {
    /// Create a new series with the given zero-based index and plot order.
    #[new]
    pub fn new(idx: u32, order: u32) -> Self {
        Self {
            inner: CoreChartSeries::new(idx, order),
        }
    }

    /// The zero-based series index.
    #[getter]
    pub fn idx(&self) -> u32 {
        self.inner.idx()
    }

    /// Set the zero-based series index.
    #[setter]
    pub fn set_idx(&mut self, idx: u32) {
        self.inner.set_idx(idx);
    }

    /// The plot order.
    #[getter]
    pub fn order(&self) -> u32 {
        self.inner.order()
    }

    /// Set the plot order.
    #[setter]
    pub fn set_order(&mut self, order: u32) {
        self.inner.set_order(order);
    }

    /// The display name for this series, or None.
    #[getter]
    pub fn name(&self) -> Option<String> {
        self.inner.name().map(|s| s.to_string())
    }

    /// Set the display name.
    #[setter]
    pub fn set_name(&mut self, name: Option<String>) {
        match name {
            Some(n) => {
                self.inner.set_name(n);
            }
            None => {
                self.inner.clear_name();
            }
        }
    }

    /// The fill color as a hex RGB string (e.g. ``"FF4472C4"``), or None.
    #[getter]
    pub fn fill_color(&self) -> Option<String> {
        self.inner.fill_color().map(|s| s.to_string())
    }

    /// Set the fill color as a hex RGB string.
    #[setter]
    pub fn set_fill_color(&mut self, color: Option<String>) {
        match color {
            Some(c) => {
                self.inner.set_fill_color(c);
            }
            None => {
                self.inner.clear_fill_color();
            }
        }
    }

    /// The line/border color as a hex RGB string, or None.
    #[getter]
    pub fn line_color(&self) -> Option<String> {
        self.inner.line_color().map(|s| s.to_string())
    }

    /// Set the line/border color as a hex RGB string.
    #[setter]
    pub fn set_line_color(&mut self, color: Option<String>) {
        match color {
            Some(c) => {
                self.inner.set_line_color(c);
            }
            None => {
                self.inner.clear_line_color();
            }
        }
    }

    /// The per-series chart type override as a string (for combo charts), or None.
    #[getter]
    pub fn series_type(&self) -> Option<String> {
        self.inner.series_type().map(|t| t.as_str().to_string())
    }

    /// Set the per-series chart type override.
    ///
    /// Pass ``None`` to clear the override. Valid values: ``"bar"``, ``"line"``,
    /// ``"pie"``, ``"area"``, ``"scatter"``, ``"doughnut"``, ``"radar"``,
    /// ``"bubble"``, ``"stock"``, ``"surface"``, ``"combo"``.
    #[setter]
    pub fn set_series_type(&mut self, series_type: Option<String>) -> PyResult<()> {
        match series_type {
            Some(s) => {
                let ct = parse_chart_type(&s)?;
                self.inner.set_series_type(ct);
            }
            None => {
                self.inner.clear_series_type();
            }
        }
        Ok(())
    }

    /// Return the category (X-axis label) data reference, or None.
    pub fn categories(&self) -> Option<XlsxChartDataRef> {
        self.inner
            .categories()
            .cloned()
            .map(XlsxChartDataRef::from_core)
    }

    /// Set the category data reference.
    pub fn set_categories(&mut self, data: &XlsxChartDataRef) {
        self.inner.set_categories(data.inner.clone());
    }

    /// Return the value (Y-axis) data reference, or None.
    pub fn values(&self) -> Option<XlsxChartDataRef> {
        self.inner
            .values()
            .cloned()
            .map(XlsxChartDataRef::from_core)
    }

    /// Set the value data reference.
    pub fn set_values(&mut self, data: &XlsxChartDataRef) {
        self.inner.set_values(data.inner.clone());
    }

    /// Return the X-values reference (for scatter/bubble charts), or None.
    pub fn x_values(&self) -> Option<XlsxChartDataRef> {
        self.inner
            .x_values()
            .cloned()
            .map(XlsxChartDataRef::from_core)
    }

    /// Set the X-values reference.
    pub fn set_x_values(&mut self, data: &XlsxChartDataRef) {
        self.inner.set_x_values(data.inner.clone());
    }

    /// Return the bubble sizes reference (for bubble charts), or None.
    pub fn bubble_sizes(&self) -> Option<XlsxChartDataRef> {
        self.inner
            .bubble_sizes()
            .cloned()
            .map(XlsxChartDataRef::from_core)
    }

    /// Set the bubble sizes reference.
    pub fn set_bubble_sizes(&mut self, data: &XlsxChartDataRef) {
        self.inner.set_bubble_sizes(data.inner.clone());
    }
}

// =============================================================================
// XlsxChartAxis
// =============================================================================

/// Python wrapper for a chart axis (category, value, date, or series axis).
#[pyclass(module = "offidized._native", name = "XlsxChartAxis")]
#[derive(Clone)]
pub struct XlsxChartAxis {
    inner: CoreChartAxis,
}

impl XlsxChartAxis {
    pub(super) fn from_core(axis: CoreChartAxis) -> Self {
        Self { inner: axis }
    }

    pub(super) fn into_core(self) -> CoreChartAxis {
        self.inner
    }
}

#[pymethods]
impl XlsxChartAxis {
    /// Create a category axis placed at the bottom (ID 1, crosses value axis 2).
    #[staticmethod]
    pub fn category() -> Self {
        Self {
            inner: CoreChartAxis::new_category(),
        }
    }

    /// Create a value axis placed on the left (ID 2, crosses category axis 1).
    #[staticmethod]
    pub fn value() -> Self {
        Self {
            inner: CoreChartAxis::new_value(),
        }
    }

    /// The axis ID.
    #[getter]
    pub fn id(&self) -> u32 {
        self.inner.id()
    }

    /// Set the axis ID.
    #[setter]
    pub fn set_id(&mut self, id: u32) {
        self.inner.set_id(id);
    }

    /// The axis type string (e.g. ``"catAx"``, ``"valAx"``).
    #[getter]
    pub fn axis_type(&self) -> String {
        self.inner.axis_type().to_string()
    }

    /// Set the axis type string.
    #[setter]
    pub fn set_axis_type(&mut self, axis_type: String) {
        self.inner.set_axis_type(axis_type);
    }

    /// The axis position (``"b"``, ``"l"``, ``"t"``, or ``"r"``).
    #[getter]
    pub fn position(&self) -> String {
        self.inner.position().to_string()
    }

    /// Set the axis position.
    #[setter]
    pub fn set_position(&mut self, position: String) {
        self.inner.set_position(position);
    }

    /// The axis title text, or None.
    #[getter]
    pub fn title(&self) -> Option<String> {
        self.inner.title().map(|s| s.to_string())
    }

    /// Set the axis title text.
    #[setter]
    pub fn set_title(&mut self, title: Option<String>) {
        match title {
            Some(t) => {
                self.inner.set_title(t);
            }
            None => {
                self.inner.clear_title();
            }
        }
    }

    /// The minimum scale value, or None (auto-scaling).
    #[getter]
    pub fn min(&self) -> Option<f64> {
        self.inner.min()
    }

    /// Set the minimum scale value.
    #[setter]
    pub fn set_min(&mut self, min: Option<f64>) {
        match min {
            Some(v) => {
                self.inner.set_min(v);
            }
            None => {
                self.inner.clear_min();
            }
        }
    }

    /// The maximum scale value, or None (auto-scaling).
    #[getter]
    pub fn max(&self) -> Option<f64> {
        self.inner.max()
    }

    /// Set the maximum scale value.
    #[setter]
    pub fn set_max(&mut self, max: Option<f64>) {
        match max {
            Some(v) => {
                self.inner.set_max(v);
            }
            None => {
                self.inner.clear_max();
            }
        }
    }

    /// The major tick/gridline interval, or None (auto interval).
    #[getter]
    pub fn major_unit(&self) -> Option<f64> {
        self.inner.major_unit()
    }

    /// Set the major tick/gridline interval.
    #[setter]
    pub fn set_major_unit(&mut self, unit: Option<f64>) {
        match unit {
            Some(v) => {
                self.inner.set_major_unit(v);
            }
            None => {
                self.inner.clear_major_unit();
            }
        }
    }

    /// The minor tick/gridline interval, or None (auto interval).
    #[getter]
    pub fn minor_unit(&self) -> Option<f64> {
        self.inner.minor_unit()
    }

    /// Set the minor tick/gridline interval.
    #[setter]
    pub fn set_minor_unit(&mut self, unit: Option<f64>) {
        match unit {
            Some(v) => {
                self.inner.set_minor_unit(v);
            }
            None => {
                self.inner.clear_minor_unit();
            }
        }
    }

    /// Whether major gridlines are displayed.
    #[getter]
    pub fn major_gridlines(&self) -> bool {
        self.inner.major_gridlines()
    }

    /// Set whether major gridlines are displayed.
    #[setter]
    pub fn set_major_gridlines(&mut self, show: bool) {
        self.inner.set_major_gridlines(show);
    }

    /// Whether minor gridlines are displayed.
    #[getter]
    pub fn minor_gridlines(&self) -> bool {
        self.inner.minor_gridlines()
    }

    /// Set whether minor gridlines are displayed.
    #[setter]
    pub fn set_minor_gridlines(&mut self, show: bool) {
        self.inner.set_minor_gridlines(show);
    }

    /// Whether the axis is hidden.
    #[getter]
    pub fn deleted(&self) -> bool {
        self.inner.deleted()
    }

    /// Set whether the axis is hidden.
    #[setter]
    pub fn set_deleted(&mut self, deleted: bool) {
        self.inner.set_deleted(deleted);
    }

    /// The number format string for axis labels, or None.
    #[getter]
    pub fn num_fmt(&self) -> Option<String> {
        self.inner.num_fmt().map(|s| s.to_string())
    }

    /// Set the number format string for axis labels.
    #[setter]
    pub fn set_num_fmt(&mut self, fmt: Option<String>) {
        match fmt {
            Some(f) => {
                self.inner.set_num_fmt(f);
            }
            None => {
                self.inner.clear_num_fmt();
            }
        }
    }
}

// =============================================================================
// XlsxChartLegend
// =============================================================================

/// Python wrapper for chart legend settings.
#[pyclass(module = "offidized._native", name = "XlsxChartLegend")]
#[derive(Clone)]
pub struct XlsxChartLegend {
    inner: CoreChartLegend,
}

impl XlsxChartLegend {
    pub(super) fn from_core(legend: CoreChartLegend) -> Self {
        Self { inner: legend }
    }

    pub(super) fn into_core(self) -> CoreChartLegend {
        self.inner
    }
}

#[allow(clippy::new_without_default)]
#[pymethods]
impl XlsxChartLegend {
    /// Create a new legend at the bottom of the chart.
    #[new]
    pub fn new() -> Self {
        Self {
            inner: CoreChartLegend::new(),
        }
    }

    /// The legend position (``"b"``, ``"t"``, ``"l"``, ``"r"``, or ``"tr"``).
    #[getter]
    pub fn position(&self) -> String {
        self.inner.position().to_string()
    }

    /// Set the legend position.
    #[setter]
    pub fn set_position(&mut self, position: String) {
        self.inner.set_position(position);
    }

    /// Whether the legend overlaps the plot area.
    #[getter]
    pub fn overlay(&self) -> bool {
        self.inner.overlay()
    }

    /// Set whether the legend overlaps the plot area.
    #[setter]
    pub fn set_overlay(&mut self, overlay: bool) {
        self.inner.set_overlay(overlay);
    }
}

// =============================================================================
// XlsxChart
// =============================================================================

/// Python wrapper for a chart embedded in a worksheet.
#[pyclass(module = "offidized._native", name = "XlsxChart")]
#[derive(Clone)]
pub struct XlsxChart {
    inner: CoreChart,
}

impl XlsxChart {
    pub(super) fn from_core(chart: CoreChart) -> Self {
        Self { inner: chart }
    }

    pub(super) fn into_core(self) -> CoreChart {
        self.inner
    }
}

#[pymethods]
impl XlsxChart {
    /// Create a new chart of the given type.
    ///
    /// Valid chart type strings: ``"bar"``, ``"line"``, ``"pie"``, ``"area"``,
    /// ``"scatter"``, ``"doughnut"``, ``"radar"``, ``"bubble"``, ``"stock"``,
    /// ``"surface"``, ``"combo"``.
    #[new]
    pub fn new(chart_type: &str) -> PyResult<Self> {
        let ct = parse_chart_type(chart_type)?;
        Ok(Self {
            inner: CoreChart::new(ct),
        })
    }

    /// The primary chart type as a string (e.g. ``"barChart"``).
    #[getter]
    pub fn chart_type(&self) -> String {
        self.inner.chart_type().as_str().to_string()
    }

    /// Set the primary chart type.
    ///
    /// Valid values: ``"bar"``, ``"line"``, ``"pie"``, ``"area"``, ``"scatter"``,
    /// ``"doughnut"``, ``"radar"``, ``"bubble"``, ``"stock"``, ``"surface"``,
    /// ``"combo"``.
    #[setter]
    pub fn set_chart_type(&mut self, chart_type: String) -> PyResult<()> {
        let ct = parse_chart_type(&chart_type)?;
        self.inner.set_chart_type(ct);
        Ok(())
    }

    /// The chart title text, or None.
    #[getter]
    pub fn title(&self) -> Option<String> {
        self.inner.title().map(|s| s.to_string())
    }

    /// Set the chart title text.
    #[setter]
    pub fn set_title(&mut self, title: Option<String>) {
        match title {
            Some(t) => {
                self.inner.set_title(t);
            }
            None => {
                self.inner.clear_title();
            }
        }
    }

    /// The bar direction as a string (``"col"`` or ``"bar"``), or None.
    #[getter]
    pub fn bar_direction(&self) -> Option<String> {
        self.inner.bar_direction().map(|d| d.as_str().to_string())
    }

    /// Set the bar direction.
    ///
    /// Valid values: ``"col"``/``"column"``/``"vertical"`` or
    /// ``"bar"``/``"horizontal"``. Pass ``None`` to clear.
    #[setter]
    pub fn set_bar_direction(&mut self, direction: Option<String>) -> PyResult<()> {
        match direction {
            Some(s) => {
                let dir = parse_bar_direction(&s)?;
                self.inner.set_bar_direction(dir);
            }
            None => {
                self.inner.clear_bar_direction();
            }
        }
        Ok(())
    }

    /// The chart grouping as a string (``"clustered"``, ``"stacked"``,
    /// ``"percentStacked"``, ``"standard"``), or None.
    #[getter]
    pub fn grouping(&self) -> Option<String> {
        self.inner.grouping().map(|g| g.as_str().to_string())
    }

    /// Set the chart grouping.
    ///
    /// Valid values: ``"clustered"``, ``"stacked"``, ``"percentStacked"``,
    /// ``"standard"``. Pass ``None`` to clear.
    #[setter]
    pub fn set_grouping(&mut self, grouping: Option<String>) -> PyResult<()> {
        match grouping {
            Some(s) => {
                let g = parse_grouping(&s)?;
                self.inner.set_grouping(g);
            }
            None => {
                self.inner.clear_grouping();
            }
        }
        Ok(())
    }

    /// Whether each data point gets a different color.
    #[getter]
    pub fn vary_colors(&self) -> bool {
        self.inner.vary_colors()
    }

    /// Set whether each data point gets a different color.
    #[setter]
    pub fn set_vary_colors(&mut self, vary: bool) {
        self.inner.set_vary_colors(vary);
    }

    /// The anchor starting column (zero-based).
    #[getter]
    #[allow(clippy::wrong_self_convention)]
    pub fn from_col(&self) -> u32 {
        self.inner.from_col()
    }

    /// Set the anchor starting column.
    #[setter]
    pub fn set_from_col(&mut self, col: u32) {
        self.inner.set_from_col(col);
    }

    /// The anchor starting row (zero-based).
    #[getter]
    #[allow(clippy::wrong_self_convention)]
    pub fn from_row(&self) -> u32 {
        self.inner.from_row()
    }

    /// Set the anchor starting row.
    #[setter]
    pub fn set_from_row(&mut self, row: u32) {
        self.inner.set_from_row(row);
    }

    /// The anchor ending column (zero-based).
    #[getter]
    pub fn to_col(&self) -> u32 {
        self.inner.to_col()
    }

    /// Set the anchor ending column.
    #[setter]
    pub fn set_to_col(&mut self, col: u32) {
        self.inner.set_to_col(col);
    }

    /// The anchor ending row (zero-based).
    #[getter]
    pub fn to_row(&self) -> u32 {
        self.inner.to_row()
    }

    /// Set the anchor ending row.
    #[setter]
    pub fn set_to_row(&mut self, row: u32) {
        self.inner.set_to_row(row);
    }

    /// The display name of the chart, or None.
    #[getter]
    pub fn name(&self) -> Option<String> {
        self.inner.name().map(|s| s.to_string())
    }

    /// Set the display name.
    #[setter]
    pub fn set_name(&mut self, name: Option<String>) {
        match name {
            Some(n) => {
                self.inner.set_name(n);
            }
            None => {
                self.inner.clear_name();
            }
        }
    }

    /// Return all data series as a list of :py:class:`XlsxChartSeries` objects.
    pub fn series(&self) -> Vec<XlsxChartSeries> {
        self.inner
            .series()
            .iter()
            .cloned()
            .map(XlsxChartSeries::from_core)
            .collect()
    }

    /// Append a data series to the chart.
    pub fn add_series(&mut self, series: &XlsxChartSeries) {
        self.inner.add_series(series.inner.clone());
    }

    /// Return all axes as a list of :py:class:`XlsxChartAxis` objects.
    pub fn axes(&self) -> Vec<XlsxChartAxis> {
        self.inner
            .axes()
            .iter()
            .cloned()
            .map(XlsxChartAxis::from_core)
            .collect()
    }

    /// Append an axis to the chart.
    pub fn add_axis(&mut self, axis: &XlsxChartAxis) {
        self.inner.add_axis(axis.inner.clone());
    }

    /// Return the legend, or None if not set.
    pub fn legend(&self) -> Option<XlsxChartLegend> {
        self.inner.legend().cloned().map(XlsxChartLegend::from_core)
    }

    /// Set the legend.
    pub fn set_legend(&mut self, legend: &XlsxChartLegend) {
        self.inner.set_legend(legend.inner.clone());
    }

    /// Set the two-cell anchor in one call.
    ///
    /// All coordinates are zero-based.
    pub fn set_anchor(&mut self, from_col: u32, from_row: u32, to_col: u32, to_row: u32) {
        self.inner.set_anchor(from_col, from_row, to_col, to_row);
    }
}

// =============================================================================
// Worksheet helper functions
// =============================================================================

/// Return all charts on the worksheet as a list of :py:class:`XlsxChart` objects.
pub(super) fn ws_charts(
    workbook: &Arc<Mutex<CoreWorkbook>>,
    name_key: &str,
) -> PyResult<Vec<XlsxChart>> {
    let wb = lock_wb(workbook)?;
    let ws = wb
        .sheet(name_key)
        .ok_or_else(|| value_error(format!("worksheet '{name_key}' not found")))?;
    Ok(ws
        .charts()
        .iter()
        .cloned()
        .map(XlsxChart::from_core)
        .collect())
}

/// Append a chart to the worksheet.
pub(super) fn ws_add_chart(
    workbook: &Arc<Mutex<CoreWorkbook>>,
    name_key: &str,
    chart: XlsxChart,
) -> PyResult<()> {
    let mut wb = lock_wb(workbook)?;
    let ws = wb
        .sheet_mut(name_key)
        .ok_or_else(|| value_error(format!("worksheet '{name_key}' not found")))?;
    ws.add_chart(chart.into_core());
    Ok(())
}

/// Remove all charts from the worksheet.
pub(super) fn ws_clear_charts(workbook: &Arc<Mutex<CoreWorkbook>>, name_key: &str) -> PyResult<()> {
    let mut wb = lock_wb(workbook)?;
    let ws = wb
        .sheet_mut(name_key)
        .ok_or_else(|| value_error(format!("worksheet '{name_key}' not found")))?;
    ws.clear_charts();
    Ok(())
}

// =============================================================================
// Registration
// =============================================================================

/// Register all chart PyO3 types with the native module.
pub(super) fn register(module: &Bound<'_, PyModule>) -> PyResult<()> {
    module.add_class::<XlsxChartDataRef>()?;
    module.add_class::<XlsxChartSeries>()?;
    module.add_class::<XlsxChartAxis>()?;
    module.add_class::<XlsxChartLegend>()?;
    module.add_class::<XlsxChart>()?;
    Ok(())
}
