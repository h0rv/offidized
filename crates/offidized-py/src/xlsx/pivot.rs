//! Python bindings for pivot table types from `offidized_xlsx`.
//!
//! Wraps [`PivotTable`], [`PivotField`], and [`PivotDataField`] with PyO3 classes
//! that mirror the core Rust API. Worksheet helper functions (`ws_*`) are
//! called from the parent `Worksheet` `#[pymethods]` block.

use super::lock_wb;
use crate::error::value_error;
use offidized_xlsx::{
    PivotDataField as CorePivotDataField, PivotField as CorePivotField,
    PivotFieldSort as CorePivotFieldSort, PivotSourceReference as CorePivotSourceReference,
    PivotSubtotalFunction as CorePivotSubtotalFunction, PivotTable as CorePivotTable,
    Workbook as CoreWorkbook,
};
use pyo3::prelude::*;
use std::sync::{Arc, Mutex};

// =============================================================================
// Internal utilities
// =============================================================================

/// Parse an A1-style cell reference (e.g. "F2", "$A$1") into 1-based (row, col).
///
/// Returns an error if the string does not look like a valid A1 reference.
fn parse_a1_to_row_col(reference: &str) -> PyResult<(u32, u32)> {
    let trimmed = reference.trim();
    let bytes = trimmed.as_bytes();
    let mut pos = 0;

    // Skip optional leading '$' for column absolute anchor.
    if bytes.get(pos) == Some(&b'$') {
        pos += 1;
    }

    // Collect column letters.
    let col_start = pos;
    while pos < bytes.len() && bytes[pos].is_ascii_alphabetic() {
        pos += 1;
    }
    let col_end = pos;
    if col_start == col_end {
        return Err(value_error(format!(
            "Invalid cell reference '{reference}': missing column letters"
        )));
    }

    // Skip optional '$' for row absolute anchor.
    if bytes.get(pos) == Some(&b'$') {
        pos += 1;
    }

    // Collect row digits.
    let row_start = pos;
    while pos < bytes.len() && bytes[pos].is_ascii_digit() {
        pos += 1;
    }
    if row_start == pos || pos != bytes.len() {
        return Err(value_error(format!(
            "Invalid cell reference '{reference}': missing or malformed row number"
        )));
    }

    // Convert column letters to 1-based index.
    let col_letters = &trimmed[col_start..col_end];
    let mut col: u32 = 0;
    for byte in col_letters.bytes() {
        let ch = byte.to_ascii_uppercase();
        col = col
            .checked_mul(26)
            .and_then(|v| v.checked_add(u32::from(ch - b'A' + 1)))
            .ok_or_else(|| value_error(format!("Column overflow in reference '{reference}'")))?;
    }

    let row: u32 = trimmed[row_start..]
        .parse()
        .map_err(|_| value_error(format!("Invalid row number in reference '{reference}'")))?;

    if row == 0 || col == 0 {
        return Err(value_error(format!(
            "Invalid cell reference '{reference}': row and column must be >= 1"
        )));
    }

    Ok((row, col))
}

// =============================================================================
// String converters
// =============================================================================

fn sort_type_to_str(sort: CorePivotFieldSort) -> &'static str {
    match sort {
        CorePivotFieldSort::Manual => "manual",
        CorePivotFieldSort::Ascending => "ascending",
        CorePivotFieldSort::Descending => "descending",
    }
}

fn str_to_sort_type(s: &str) -> PyResult<CorePivotFieldSort> {
    match s.to_lowercase().as_str() {
        "manual" => Ok(CorePivotFieldSort::Manual),
        "ascending" => Ok(CorePivotFieldSort::Ascending),
        "descending" => Ok(CorePivotFieldSort::Descending),
        _ => Err(value_error(format!(
            "Unknown sort type '{s}': expected 'manual', 'ascending', or 'descending'"
        ))),
    }
}

fn subtotal_to_str(f: CorePivotSubtotalFunction) -> &'static str {
    match f {
        CorePivotSubtotalFunction::Average => "average",
        CorePivotSubtotalFunction::Count => "count",
        CorePivotSubtotalFunction::CountNums => "countNums",
        CorePivotSubtotalFunction::Max => "max",
        CorePivotSubtotalFunction::Min => "min",
        CorePivotSubtotalFunction::Product => "product",
        CorePivotSubtotalFunction::StdDev => "stdDev",
        CorePivotSubtotalFunction::StdDevP => "stdDevP",
        CorePivotSubtotalFunction::Sum => "sum",
        CorePivotSubtotalFunction::Var => "var",
        CorePivotSubtotalFunction::VarP => "varP",
    }
}

fn str_to_subtotal(s: &str) -> PyResult<CorePivotSubtotalFunction> {
    match s.trim() {
        "average" => Ok(CorePivotSubtotalFunction::Average),
        "count" => Ok(CorePivotSubtotalFunction::Count),
        "countNums" => Ok(CorePivotSubtotalFunction::CountNums),
        "max" => Ok(CorePivotSubtotalFunction::Max),
        "min" => Ok(CorePivotSubtotalFunction::Min),
        "product" => Ok(CorePivotSubtotalFunction::Product),
        "stdDev" => Ok(CorePivotSubtotalFunction::StdDev),
        "stdDevP" => Ok(CorePivotSubtotalFunction::StdDevP),
        "sum" => Ok(CorePivotSubtotalFunction::Sum),
        "var" => Ok(CorePivotSubtotalFunction::Var),
        "varP" => Ok(CorePivotSubtotalFunction::VarP),
        _ => Err(value_error(format!(
            "Unknown subtotal function '{s}': expected one of \
             'average', 'count', 'countNums', 'max', 'min', 'product', \
             'stdDev', 'stdDevP', 'sum', 'var', 'varP'"
        ))),
    }
}

// =============================================================================
// XlsxPivotField
// =============================================================================

/// Python wrapper for a pivot row, column, or page field.
///
/// Represents a single dimension field placed on the row, column, or filter
/// axis of a pivot table. Holds the source column name and display options.
#[pyclass(module = "offidized._native", name = "XlsxPivotField", from_py_object)]
#[derive(Clone)]
pub struct XlsxPivotField {
    inner: CorePivotField,
}

impl XlsxPivotField {
    pub(super) fn from_core(field: CorePivotField) -> Self {
        Self { inner: field }
    }

    pub(super) fn into_core(self) -> CorePivotField {
        self.inner
    }
}

#[pymethods]
impl XlsxPivotField {
    /// Create a new pivot field with the given source column name.
    #[new]
    pub fn new(name: &str) -> Self {
        Self {
            inner: CorePivotField::new(name),
        }
    }

    /// The source column name (from the data range header row).
    #[getter]
    pub fn name(&self) -> &str {
        self.inner.name()
    }

    /// Set the source column name.
    #[setter]
    pub fn set_name(&mut self, name: String) {
        self.inner.set_name(name);
    }

    /// Custom label displayed instead of the field name, or None.
    #[getter]
    pub fn custom_label(&self) -> Option<&str> {
        self.inner.custom_label()
    }

    /// Set a custom display label for this field.
    #[setter]
    pub fn set_custom_label(&mut self, label: Option<String>) {
        match label {
            Some(l) => {
                self.inner.set_custom_label(l);
            }
            None => {
                self.inner.clear_custom_label();
            }
        }
    }

    /// Whether subtotals are shown for all items in this field.
    #[getter]
    pub fn show_all_subtotals(&self) -> bool {
        self.inner.show_all_subtotals()
    }

    /// Set whether to show subtotals for all items.
    #[setter]
    pub fn set_show_all_subtotals(&mut self, value: bool) {
        self.inner.set_show_all_subtotals(value);
    }

    /// Whether blank rows are inserted after each item group.
    #[getter]
    pub fn insert_blank_rows(&self) -> bool {
        self.inner.insert_blank_rows()
    }

    /// Set whether to insert blank rows after each item group.
    #[setter]
    pub fn set_insert_blank_rows(&mut self, value: bool) {
        self.inner.set_insert_blank_rows(value);
    }

    /// Sort order for this field: ``"manual"``, ``"ascending"``, or ``"descending"``.
    #[getter]
    pub fn sort_type(&self) -> &str {
        sort_type_to_str(self.inner.sort_type())
    }

    /// Set the sort order. Accepted values: ``"manual"``, ``"ascending"``, ``"descending"``.
    #[setter]
    pub fn set_sort_type(&mut self, value: &str) -> PyResult<()> {
        let sort = str_to_sort_type(value)?;
        self.inner.set_sort_type(sort);
        Ok(())
    }
}

// =============================================================================
// XlsxPivotDataField
// =============================================================================

/// Python wrapper for a pivot data field (aggregated value column).
///
/// Data fields are placed in the values area of a pivot table and summarize
/// the source data using a chosen aggregation function such as SUM or AVERAGE.
#[pyclass(
    module = "offidized._native",
    name = "XlsxPivotDataField",
    from_py_object
)]
#[derive(Clone)]
pub struct XlsxPivotDataField {
    inner: CorePivotDataField,
}

impl XlsxPivotDataField {
    pub(super) fn from_core(field: CorePivotDataField) -> Self {
        Self { inner: field }
    }

    pub(super) fn into_core(self) -> CorePivotDataField {
        self.inner
    }
}

#[pymethods]
impl XlsxPivotDataField {
    /// Create a new data field referencing the given source column name.
    #[new]
    pub fn new(field_name: &str) -> Self {
        Self {
            inner: CorePivotDataField::new(field_name),
        }
    }

    /// The source column name used as the aggregation input.
    #[getter]
    pub fn field_name(&self) -> &str {
        self.inner.field_name()
    }

    /// Set the source column name.
    #[setter]
    pub fn set_field_name(&mut self, name: String) {
        self.inner.set_field_name(name);
    }

    /// Custom display name shown in the pivot table header (e.g. "Sum of Sales"), or None.
    #[getter]
    pub fn custom_name(&self) -> Option<&str> {
        self.inner.custom_name()
    }

    /// Set a custom display name. Pass None to remove it.
    #[setter]
    pub fn set_custom_name(&mut self, name: Option<String>) {
        match name {
            Some(n) => {
                self.inner.set_custom_name(n);
            }
            None => {
                self.inner.clear_custom_name();
            }
        }
    }

    /// Aggregation function as a string (e.g. ``"sum"``, ``"average"``, ``"count"``).
    #[getter]
    pub fn subtotal(&self) -> &str {
        subtotal_to_str(self.inner.subtotal())
    }

    /// Set the aggregation function.
    ///
    /// Accepted values: ``"average"``, ``"count"``, ``"countNums"``, ``"max"``,
    /// ``"min"``, ``"product"``, ``"stdDev"``, ``"stdDevP"``, ``"sum"``,
    /// ``"var"``, ``"varP"``.
    #[setter]
    pub fn set_subtotal(&mut self, value: &str) -> PyResult<()> {
        let f = str_to_subtotal(value)?;
        self.inner.set_subtotal(f);
        Ok(())
    }

    /// Excel number format string applied to values (e.g. ``"#,##0.00"``), or None.
    #[getter]
    pub fn number_format(&self) -> Option<&str> {
        self.inner.number_format()
    }

    /// Set the number format. Pass None to remove it.
    #[setter]
    pub fn set_number_format(&mut self, fmt: Option<String>) {
        match fmt {
            Some(f) => {
                self.inner.set_number_format(f);
            }
            None => {
                self.inner.clear_number_format();
            }
        }
    }
}

// =============================================================================
// XlsxPivotTable
// =============================================================================

/// Python wrapper for a pivot table attached to a worksheet.
///
/// A pivot table summarizes data from a source range or named table. It
/// organises fields along row, column, and page (filter) axes and displays
/// aggregated values in the data area.
///
/// Construct via :py:meth:`XlsxPivotTable.__init__` or retrieve existing
/// tables through :py:meth:`Worksheet.pivot_tables`.
#[pyclass(module = "offidized._native", name = "XlsxPivotTable", from_py_object)]
#[derive(Clone)]
pub struct XlsxPivotTable {
    inner: CorePivotTable,
}

impl XlsxPivotTable {
    pub(super) fn from_core(table: CorePivotTable) -> Self {
        Self { inner: table }
    }

    pub(super) fn into_core(self) -> CorePivotTable {
        self.inner
    }
}

#[pymethods]
impl XlsxPivotTable {
    /// Create a new pivot table object.
    ///
    /// The table is not attached to any worksheet until passed to
    /// :py:meth:`Worksheet.add_pivot_table_obj`. To create and attach in one
    /// step use :py:meth:`Worksheet.add_pivot_table`.
    ///
    /// Args:
    ///     name: Unique name for the pivot table (e.g. ``"SalesSummary"``).
    ///     source_reference: Source range string (e.g. ``"Sheet1!$A$1:$D$100"``).
    ///     target: Top-left anchor cell reference (e.g. ``"F2"``). Parsed to
    ///         derive the 0-based row/column stored on the table.
    #[new]
    pub fn new(name: &str, source_reference: &str, target: &str) -> PyResult<Self> {
        let (row, col) = parse_a1_to_row_col(target)?;
        let src = CorePivotSourceReference::from_range(source_reference);
        let mut inner = CorePivotTable::new(name, src);
        inner.set_target(row.saturating_sub(1), col.saturating_sub(1));
        Ok(Self { inner })
    }

    /// The pivot table name.
    #[getter]
    pub fn name(&self) -> &str {
        self.inner.name()
    }

    /// Set the pivot table name.
    #[setter]
    pub fn set_name(&mut self, name: String) {
        self.inner.set_name(name);
    }

    /// The source data reference as a string.
    #[getter]
    pub fn source_reference(&self) -> &str {
        self.inner.source_reference().as_str()
    }

    /// Set the source data reference string.
    ///
    /// The string is always treated as a worksheet range reference
    /// (e.g. ``"Sheet1!$A$1:$D$100"``).
    #[setter]
    pub fn set_source_reference(&mut self, value: String) {
        self.inner
            .set_source_reference(CorePivotSourceReference::from_range(value));
    }

    /// Target cell row (0-indexed).
    #[getter]
    pub fn target_row(&self) -> u32 {
        self.inner.target_row()
    }

    /// Target cell column (0-indexed).
    #[getter]
    pub fn target_col(&self) -> u32 {
        self.inner.target_col()
    }

    /// Set the target cell position using 0-indexed row and column.
    pub fn set_target(&mut self, row: u32, col: u32) {
        self.inner.set_target(row, col);
    }

    /// Return the row fields as a list of :py:class:`XlsxPivotField` objects.
    #[getter]
    pub fn row_fields(&self) -> Vec<XlsxPivotField> {
        self.inner
            .row_fields()
            .iter()
            .cloned()
            .map(XlsxPivotField::from_core)
            .collect()
    }

    /// Return the column fields as a list of :py:class:`XlsxPivotField` objects.
    #[getter]
    pub fn column_fields(&self) -> Vec<XlsxPivotField> {
        self.inner
            .column_fields()
            .iter()
            .cloned()
            .map(XlsxPivotField::from_core)
            .collect()
    }

    /// Return the page (filter) fields as a list of :py:class:`XlsxPivotField` objects.
    #[getter]
    pub fn page_fields(&self) -> Vec<XlsxPivotField> {
        self.inner
            .page_fields()
            .iter()
            .cloned()
            .map(XlsxPivotField::from_core)
            .collect()
    }

    /// Return the data fields as a list of :py:class:`XlsxPivotDataField` objects.
    #[getter]
    pub fn data_fields(&self) -> Vec<XlsxPivotDataField> {
        self.inner
            .data_fields()
            .iter()
            .cloned()
            .map(XlsxPivotDataField::from_core)
            .collect()
    }

    /// Append a field to the row axis.
    pub fn add_row_field(&mut self, field: &XlsxPivotField) {
        self.inner.add_row_field(field.inner.clone());
    }

    /// Append a field to the column axis.
    pub fn add_column_field(&mut self, field: &XlsxPivotField) {
        self.inner.add_column_field(field.inner.clone());
    }

    /// Append a field to the page (filter) area.
    pub fn add_page_field(&mut self, field: &XlsxPivotField) {
        self.inner.add_page_field(field.inner.clone());
    }

    /// Append a data field to the values area.
    pub fn add_data_field(&mut self, field: &XlsxPivotDataField) {
        self.inner.add_data_field(field.inner.clone());
    }
}

// =============================================================================
// Worksheet helper functions
// =============================================================================

/// Return all pivot tables on the worksheet as a list of `XlsxPivotTable` objects.
pub(super) fn ws_pivot_tables(
    workbook: &Arc<Mutex<CoreWorkbook>>,
    name_key: &str,
) -> PyResult<Vec<XlsxPivotTable>> {
    let wb = lock_wb(workbook)?;
    let ws = wb
        .sheet(name_key)
        .ok_or_else(|| value_error(format!("worksheet '{name_key}' not found")))?;
    Ok(ws
        .pivot_tables()
        .iter()
        .cloned()
        .map(XlsxPivotTable::from_core)
        .collect())
}

/// Add a new pivot table to the worksheet, placing its top-left corner at `target`.
///
/// `target` is an A1-style cell reference (e.g. `"F2"`) that is parsed to
/// derive the 0-based row/column position stored in `PivotTable::set_target`.
pub(super) fn ws_add_pivot_table(
    workbook: &Arc<Mutex<CoreWorkbook>>,
    name_key: &str,
    name: &str,
    source_reference: &str,
    target: &str,
) -> PyResult<()> {
    let (target_row, target_col) = parse_a1_to_row_col(target)?;

    let src = CorePivotSourceReference::from_range(source_reference);
    let mut pivot = CorePivotTable::new(name, src);
    // set_target takes 0-indexed values; parse_a1_to_row_col returns 1-based.
    pivot.set_target(target_row.saturating_sub(1), target_col.saturating_sub(1));

    let mut wb = lock_wb(workbook)?;
    let ws = wb
        .sheet_mut(name_key)
        .ok_or_else(|| value_error(format!("worksheet '{name_key}' not found")))?;
    ws.add_pivot_table(pivot);
    Ok(())
}

/// Add a pre-built `XlsxPivotTable` object to the worksheet, optionally
/// updating its target position from an A1 reference string.
///
/// If `target` is `None` the placement stored on the object is used as-is.
pub(super) fn ws_add_pivot_table_obj(
    workbook: &Arc<Mutex<CoreWorkbook>>,
    name_key: &str,
    pivot: XlsxPivotTable,
    target: Option<&str>,
) -> PyResult<()> {
    let mut core = pivot.into_core();
    if let Some(t) = target {
        let (row, col) = parse_a1_to_row_col(t)?;
        core.set_target(row.saturating_sub(1), col.saturating_sub(1));
    }

    let mut wb = lock_wb(workbook)?;
    let ws = wb
        .sheet_mut(name_key)
        .ok_or_else(|| value_error(format!("worksheet '{name_key}' not found")))?;
    ws.add_pivot_table(core);
    Ok(())
}

/// Remove all pivot tables from the worksheet.
pub(super) fn ws_clear_pivot_tables(
    workbook: &Arc<Mutex<CoreWorkbook>>,
    name_key: &str,
) -> PyResult<()> {
    let mut wb = lock_wb(workbook)?;
    let ws = wb
        .sheet_mut(name_key)
        .ok_or_else(|| value_error(format!("worksheet '{name_key}' not found")))?;
    ws.clear_pivot_tables();
    Ok(())
}

// =============================================================================
// Registration
// =============================================================================

/// Register all pivot table PyO3 types with the native module.
pub(super) fn register(module: &Bound<'_, PyModule>) -> PyResult<()> {
    module.add_class::<XlsxPivotField>()?;
    module.add_class::<XlsxPivotDataField>()?;
    module.add_class::<XlsxPivotTable>()?;
    Ok(())
}
