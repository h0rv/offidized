//! Python bindings for conditional formatting types from `offidized_xlsx`.
//!
//! Wraps [`ConditionalFormatting`] with a PyO3 class that exposes range
//! management, rule type, and formula inspection. Worksheet helper functions
//! (`ws_*`) are called from the parent `Worksheet` `#[pymethods]` block.

use super::lock_wb;
use crate::error::{value_error, xlsx_error_to_py};
use offidized_xlsx::{
    ConditionalFormatting as CoreConditionalFormatting,
    ConditionalFormattingRuleType as CoreRuleType, Workbook as CoreWorkbook,
};
use pyo3::prelude::*;
use std::sync::{Arc, Mutex};

// =============================================================================
// Helpers
// =============================================================================

fn rule_type_to_str(rt: CoreRuleType) -> &'static str {
    match rt {
        CoreRuleType::CellIs => "cellIs",
        CoreRuleType::Expression => "expression",
        CoreRuleType::ColorScale => "colorScale",
        CoreRuleType::DataBar => "dataBar",
        CoreRuleType::IconSet => "iconSet",
        CoreRuleType::Top10 => "top10",
        CoreRuleType::AboveAverage => "aboveAverage",
        CoreRuleType::TimePeriod => "timePeriod",
        CoreRuleType::DuplicateValues => "duplicateValues",
        CoreRuleType::UniqueValues => "uniqueValues",
        CoreRuleType::ContainsText => "containsText",
        CoreRuleType::NotContainsText => "notContainsText",
        CoreRuleType::BeginsWith => "beginsWith",
        CoreRuleType::EndsWith => "endsWith",
        CoreRuleType::ContainsBlanks => "containsBlanks",
        CoreRuleType::NotContainsBlanks => "notContainsBlanks",
        CoreRuleType::ContainsErrors => "containsErrors",
        CoreRuleType::NotContainsErrors => "notContainsErrors",
    }
}

// =============================================================================
// XlsxConditionalFormatting
// =============================================================================

/// Python wrapper for a conditional formatting rule.
///
/// A conditional formatting rule applies a visual style to one or more cell
/// ranges when a condition is met. Construct with a rule type, ranges, and
/// formulas, or retrieve existing rules via
/// :py:meth:`Worksheet.conditional_formattings`.
///
/// Currently, only ``"cellIs"`` and ``"expression"`` rule types can be
/// constructed from Python. Other rule types (``"colorScale"``, ``"dataBar"``,
/// etc.) are available when reading existing workbooks.
#[pyclass(
    module = "offidized._native",
    name = "XlsxConditionalFormatting",
    from_py_object
)]
#[derive(Clone)]
pub struct XlsxConditionalFormatting {
    inner: CoreConditionalFormatting,
}

impl XlsxConditionalFormatting {
    pub(super) fn from_core(cf: CoreConditionalFormatting) -> Self {
        Self { inner: cf }
    }

    pub(super) fn into_core(self) -> CoreConditionalFormatting {
        self.inner
    }
}

#[pymethods]
impl XlsxConditionalFormatting {
    /// Create a new conditional formatting rule.
    ///
    /// Args:
    ///     rule_type: The rule type string. Constructible types are
    ///         ``"cellIs"`` and ``"expression"``. Other types (``"colorScale"``,
    ///         ``"dataBar"``, ``"iconSet"``, etc.) are read-only and come from
    ///         existing workbooks.
    ///     ranges: A list of A1-notation range strings (e.g. ``["A1:B10"]``).
    ///     formulas: A list of formula strings (at least one required).
    #[new]
    pub fn new(rule_type: &str, ranges: Vec<String>, formulas: Vec<String>) -> PyResult<Self> {
        let refs: Vec<&str> = ranges.iter().map(|s| s.as_str()).collect();
        let formula_refs: Vec<&str> = formulas.iter().map(|s| s.as_str()).collect();
        let inner = match rule_type {
            "cellIs" => {
                CoreConditionalFormatting::cell_is(refs, formula_refs).map_err(xlsx_error_to_py)?
            }
            "expression" => CoreConditionalFormatting::expression(refs, formula_refs)
                .map_err(xlsx_error_to_py)?,
            _ => {
                return Err(value_error(format!(
                    "Cannot construct rule type '{rule_type}' from Python. \
                     Only 'cellIs' and 'expression' are constructible; other \
                     types are available when reading existing workbooks."
                )));
            }
        };
        Ok(Self { inner })
    }

    /// Return the rule type as a string (e.g. ``"cellIs"``, ``"expression"``).
    #[getter]
    pub fn rule_type(&self) -> &str {
        rule_type_to_str(self.inner.rule_type())
    }

    /// Return the cell ranges this rule applies to as a list of ``"start:end"`` strings.
    ///
    /// Single-cell ranges are returned as ``"A1:A1"``.
    pub fn ranges(&self) -> Vec<String> {
        self.inner
            .sqref()
            .iter()
            .map(|r| format!("{}:{}", r.start(), r.end()))
            .collect()
    }

    /// Return the formulas for this rule as a list of strings.
    #[getter]
    pub fn formulas(&self) -> Vec<String> {
        self.inner
            .formulas()
            .iter()
            .map(|s| s.to_string())
            .collect()
    }

    /// Add a cell range to this rule.
    ///
    /// Args:
    ///     range: A cell range string (e.g. ``"B2:D10"`` or ``"A1"``).
    pub fn add_range(&mut self, range: &str) -> PyResult<()> {
        self.inner.add_range(range).map_err(xlsx_error_to_py)?;
        Ok(())
    }

    /// Add a formula to this rule.
    ///
    /// Args:
    ///     formula: A formula string (must not be empty).
    pub fn add_formula(&mut self, formula: &str) -> PyResult<()> {
        self.inner.add_formula(formula).map_err(xlsx_error_to_py)?;
        Ok(())
    }
}

// =============================================================================
// Worksheet helper functions
// =============================================================================

/// Return all conditional formatting rules on the worksheet as a list of
/// :py:class:`XlsxConditionalFormatting` objects.
pub(super) fn ws_conditional_formattings(
    workbook: &Arc<Mutex<CoreWorkbook>>,
    name_key: &str,
) -> PyResult<Vec<XlsxConditionalFormatting>> {
    let wb = lock_wb(workbook)?;
    let ws = wb
        .sheet(name_key)
        .ok_or_else(|| value_error(format!("worksheet '{name_key}' not found")))?;
    Ok(ws
        .conditional_formattings()
        .iter()
        .cloned()
        .map(XlsxConditionalFormatting::from_core)
        .collect())
}

/// Add a conditional formatting rule to the worksheet.
pub(super) fn ws_add_conditional_formatting(
    workbook: &Arc<Mutex<CoreWorkbook>>,
    name_key: &str,
    cf: XlsxConditionalFormatting,
) -> PyResult<()> {
    let mut wb = lock_wb(workbook)?;
    let ws = wb
        .sheet_mut(name_key)
        .ok_or_else(|| value_error(format!("worksheet '{name_key}' not found")))?;
    ws.add_conditional_formatting(cf.into_core());
    Ok(())
}

/// Remove all conditional formatting rules from the worksheet.
pub(super) fn ws_clear_conditional_formattings(
    workbook: &Arc<Mutex<CoreWorkbook>>,
    name_key: &str,
) -> PyResult<()> {
    let mut wb = lock_wb(workbook)?;
    let ws = wb
        .sheet_mut(name_key)
        .ok_or_else(|| value_error(format!("worksheet '{name_key}' not found")))?;
    ws.clear_conditional_formattings();
    Ok(())
}

// =============================================================================
// Registration
// =============================================================================

/// Register all conditional formatting PyO3 types with the native module.
pub(super) fn register(module: &Bound<'_, PyModule>) -> PyResult<()> {
    module.add_class::<XlsxConditionalFormatting>()?;
    Ok(())
}
