//! Python bindings for the `offidized_xlsx` finance module.
//!
//! Exposes `wb_finance_model` and `wb_pivot_on` as free functions that operate on a
//! `Workbook` wrapper, plus the `xlsx_a1_to_r1c1` / `xlsx_r1c1_to_a1` reference-style
//! converters.

use crate::error::{value_error, xlsx_error_to_py};
use crate::xlsx::{lock_wb, Workbook};
use offidized_xlsx::{a1_to_r1c1, avg, r1c1_to_a1, sum, MeasureType, PivotValueSpec};
use pyo3::prelude::*;

// =============================================================================
// Helpers
// =============================================================================

/// Parse a `MeasureType` from a lowercase string.
///
/// Accepted values: `"currency"`, `"percentage"`, `"bps"`, `"multiple"`, `"number"`.
fn parse_measure_type(s: &str) -> PyResult<MeasureType> {
    match s.to_lowercase().as_str() {
        "currency" => Ok(MeasureType::Currency),
        "percentage" => Ok(MeasureType::Percentage),
        "bps" => Ok(MeasureType::Bps),
        "multiple" => Ok(MeasureType::Multiple),
        "number" => Ok(MeasureType::Number),
        _ => Err(value_error(format!(
            "Unknown measure type '{s}'. \
             Expected one of: currency, percentage, bps, multiple, number"
        ))),
    }
}

/// Parse a pivot aggregate function from a lowercase string.
///
/// Accepted values: `"sum"`, `"avg"` / `"average"`.
fn parse_pivot_value_spec(field_name: &str, function: &str) -> PyResult<PivotValueSpec> {
    match function.to_lowercase().as_str() {
        "sum" => Ok(sum(field_name.to_string())),
        "avg" | "average" => Ok(avg(field_name.to_string())),
        _ => Err(value_error(format!(
            "Unknown pivot value function '{function}'. Expected 'sum' or 'avg'."
        ))),
    }
}

// =============================================================================
// Finance-model builder (workbook-level)
// =============================================================================

/// Build a finance model in `workbook`, creating `"{name} Model"` and `"{name} Data"` sheets.
///
/// - `name`: Model name (used as a prefix for the two generated sheets).
/// - `dimensions`: Each entry is `(dimension_name, [member, ...])`.
/// - `measures`: Each entry is `(measure_name, measure_type_str)` where `measure_type_str`
///   is one of `"currency"`, `"percentage"`, `"bps"`, `"multiple"`, `"number"`.
/// - `scenarios`: Optional scenario labels. Pass an empty list for a single `"base"` scenario.
#[pyfunction]
pub fn wb_finance_model(
    workbook: &Workbook,
    name: &str,
    dimensions: Vec<(String, Vec<String>)>,
    measures: Vec<(String, String)>,
    scenarios: Vec<String>,
) -> PyResult<()> {
    let mut wb = lock_wb(&workbook.inner)?;

    let mut builder = wb.finance_model(name);

    for (dim_name, members) in dimensions {
        builder = builder.dimension(dim_name, members);
    }

    for (measure_name, measure_type_str) in &measures {
        let measure_type = parse_measure_type(measure_type_str)?;
        builder = builder.measure(measure_name.as_str(), measure_type);
    }

    for scenario in scenarios {
        builder = builder.scenario(scenario);
    }

    builder.build().map_err(xlsx_error_to_py)?;
    Ok(())
}

// =============================================================================
// Pivot builder (workbook-level)
// =============================================================================

/// Build and place a pivot table in `workbook`.
///
/// - `target_sheet`: Name of the sheet where the pivot table will be placed.
/// - `name`: Internal pivot table name.
/// - `source`: Source range reference, e.g. `"Data!A1:D100"`.
/// - `rows`: Fields to use as row labels.
/// - `cols`: Fields to use as column labels.
/// - `filters`: Fields to use as report filters / page fields.
/// - `values`: Each entry is `(field_name, function)` where `function` is `"sum"` or `"avg"`.
/// - `target_cell`: Top-left cell for the pivot table output (e.g. `"A4"`).
#[pyfunction]
#[allow(clippy::too_many_arguments)]
pub fn wb_pivot_on(
    workbook: &Workbook,
    target_sheet: &str,
    name: &str,
    source: &str,
    rows: Vec<String>,
    cols: Vec<String>,
    filters: Vec<String>,
    values: Vec<(String, String)>,
    target_cell: &str,
) -> PyResult<()> {
    let value_specs: Vec<PivotValueSpec> = values
        .iter()
        .map(|(field_name, function)| {
            parse_pivot_value_spec(field_name.as_str(), function.as_str())
        })
        .collect::<PyResult<_>>()?;

    let mut wb = lock_wb(&workbook.inner)?;

    wb.pivot_on(target_sheet, name)
        .source(source)
        .rows(rows)
        .cols(cols)
        .filters(filters)
        .values(value_specs)
        .place(target_cell)
        .map_err(xlsx_error_to_py)?;

    Ok(())
}

// =============================================================================
// Reference-style converters
// =============================================================================

/// Convert an A1-style cell reference to R1C1-style relative to a base cell.
///
/// - `reference`: A1 reference, e.g. `"$B$3"`, `"A1"`, `"$A1"`, `"A$1"`.
/// - `base_row`: 1-based row of the cell containing the formula.
/// - `base_col`: 1-based column of the cell containing the formula.
///
/// Returns the R1C1 string, or `None` if the input is invalid.
#[pyfunction]
pub fn xlsx_a1_to_r1c1(reference: &str, base_row: u32, base_col: u32) -> Option<String> {
    a1_to_r1c1(reference, base_row, base_col)
}

/// Convert an R1C1-style cell reference to A1-style relative to a base cell.
///
/// - `reference`: R1C1 reference, e.g. `"R1C1"`, `"R[-1]C[2]"`, `"RC[-1]"`.
/// - `base_row`: 1-based row of the cell containing the formula.
/// - `base_col`: 1-based column of the cell containing the formula.
///
/// Returns the A1 string with `$` for absolute parts, or `None` if the input is invalid.
#[pyfunction]
pub fn xlsx_r1c1_to_a1(reference: &str, base_row: u32, base_col: u32) -> Option<String> {
    r1c1_to_a1(reference, base_row, base_col)
}

// =============================================================================
// Registration
// =============================================================================

/// Register all finance-related free functions with the Python module.
pub(super) fn register(module: &Bound<'_, PyModule>) -> PyResult<()> {
    module.add_function(wrap_pyfunction!(wb_finance_model, module)?)?;
    module.add_function(wrap_pyfunction!(wb_pivot_on, module)?)?;
    module.add_function(wrap_pyfunction!(xlsx_a1_to_r1c1, module)?)?;
    module.add_function(wrap_pyfunction!(xlsx_r1c1_to_a1, module)?)?;
    Ok(())
}
