//! Built-in Excel function implementations.
//!
//! This module provides a [`FunctionRegistry`] that maps function names to
//! their implementations, along with helper utilities for extracting and
//! coercing function arguments.

use std::collections::HashMap;

use crate::context::EvalContext;
use crate::reference::RangeRef;
use crate::value::{CalcValue, ScalarValue, XlError};

pub mod database;
pub mod date_time;
pub mod financial;
pub mod info;
pub mod logical;
pub mod lookup;
pub mod math;
pub mod stat;
pub mod text;

/// Whether a function parameter accepts a scalar value or a cell range.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ParamKind {
    /// The parameter expects a single scalar value.
    Scalar,
    /// The parameter expects a cell range reference.
    Range,
}

/// An argument passed to a function implementation.
pub enum FunctionArg<'a> {
    /// A resolved scalar (or scalar-coerced) value.
    Value(CalcValue),
    /// An unevaluated range reference with its evaluation context.
    Range {
        /// The range reference.
        range: &'a RangeRef,
        /// The evaluation context for resolving cell values.
        ctx: &'a EvalContext<'a>,
    },
}

/// The type signature for built-in function implementations.
pub type FormulaFn = fn(&[FunctionArg<'_>], &EvalContext<'_>) -> CalcValue;

/// Metadata for a registered function.
pub struct FunctionDef {
    /// The canonical (uppercase) name of the function.
    pub name: &'static str,
    /// Minimum number of arguments.
    pub min_args: usize,
    /// Maximum number of arguments. Use `usize::MAX` for unlimited.
    pub max_args: usize,
    /// Parameter kind pattern. Cycled for varargs.
    pub param_kinds: &'static [ParamKind],
    /// The function implementation.
    pub func: FormulaFn,
}

/// Registry of built-in and user-defined formula functions.
pub struct FunctionRegistry {
    functions: HashMap<String, FunctionDef>,
}

impl FunctionRegistry {
    /// Creates an empty function registry.
    pub fn new() -> Self {
        Self {
            functions: HashMap::new(),
        }
    }

    /// Registers a function definition, keyed by uppercase name.
    pub fn register(&mut self, def: FunctionDef) {
        self.functions.insert(def.name.to_uppercase(), def);
    }

    /// Looks up a function definition by name (case-insensitive).
    pub fn get(&self, name: &str) -> Option<&FunctionDef> {
        self.functions.get(&name.to_uppercase())
    }

    /// Creates a registry pre-populated with all built-in functions.
    pub fn with_builtins() -> Self {
        let mut reg = Self::new();
        math::register(&mut reg);
        stat::register(&mut reg);
        text::register(&mut reg);
        logical::register(&mut reg);
        lookup::register(&mut reg);
        date_time::register(&mut reg);
        info::register(&mut reg);
        database::register(&mut reg);
        financial::register(&mut reg);
        reg
    }
}

impl Default for FunctionRegistry {
    fn default() -> Self {
        Self::new()
    }
}

// ---------------------------------------------------------------------------
// Helper utilities for function implementations
// ---------------------------------------------------------------------------

/// Iterates over all cells in a range, calling the provider for each cell.
///
/// Cells are yielded row-by-row, column-by-column, from `start_row..=end_row`
/// and `start_col..=end_col`.
pub fn iter_range_values<'a>(
    range: &'a RangeRef,
    ctx: &'a EvalContext<'a>,
) -> impl Iterator<Item = ScalarValue> + 'a {
    let sheet = range.sheet.as_deref();
    (range.start_row..=range.end_row).flat_map(move |row| {
        (range.start_col..=range.end_col).map(move |col| ctx.provider.cell_value(sheet, row, col))
    })
}

/// Collects numeric values from function arguments, flattening ranges.
///
/// - For `FunctionArg::Value`: coerces the scalar to a number. Errors propagate.
///   Blanks become 0. Text that cannot parse returns `#VALUE!`.
/// - For `FunctionArg::Range`: iterates the range. Blanks and text are skipped.
///   Errors propagate.
///
/// Returns `Ok(numbers)` or `Err(error_value)` if an error was encountered.
pub fn collect_numbers(
    args: &[FunctionArg<'_>],
    ctx: &EvalContext<'_>,
) -> Result<Vec<f64>, CalcValue> {
    let _ = ctx; // ctx is available if needed for future use
    let mut numbers = Vec::new();
    for arg in args {
        match arg {
            FunctionArg::Value(v) => {
                let scalar = v.as_scalar();
                if let ScalarValue::Error(e) = scalar {
                    return Err(CalcValue::error(*e));
                }
                let coerced = scalar.to_number();
                match coerced {
                    ScalarValue::Number(n) => numbers.push(n),
                    ScalarValue::Error(e) => return Err(CalcValue::error(e)),
                    _ => {}
                }
            }
            FunctionArg::Range { range, ctx: rctx } => {
                for val in iter_range_values(range, rctx) {
                    match val {
                        ScalarValue::Error(e) => return Err(CalcValue::error(e)),
                        ScalarValue::Number(n) => numbers.push(n),
                        ScalarValue::Bool(b) => numbers.push(if b { 1.0 } else { 0.0 }),
                        ScalarValue::Blank | ScalarValue::Text(_) => {
                            // Skip blanks and text in ranges
                        }
                    }
                }
            }
        }
    }
    Ok(numbers)
}

/// Extracts a scalar value from a function argument.
///
/// For `FunctionArg::Value`, returns a reference to the inner scalar.
/// For `FunctionArg::Range`, returns `#VALUE!` (should not occur when
/// `param_kinds` is correctly set to `Scalar`).
pub fn require_scalar<'a>(arg: &'a FunctionArg<'_>) -> &'a ScalarValue {
    match arg {
        FunctionArg::Value(v) => v.as_scalar(),
        FunctionArg::Range { .. } => &ScalarValue::Error(XlError::Value),
    }
}

/// Coerces a function argument to a number. Propagates errors.
pub fn require_number(arg: &FunctionArg<'_>) -> Result<f64, CalcValue> {
    let scalar = require_scalar(arg);
    if let ScalarValue::Error(e) = scalar {
        return Err(CalcValue::error(*e));
    }
    match scalar.to_number() {
        ScalarValue::Number(n) => Ok(n),
        ScalarValue::Error(e) => Err(CalcValue::error(e)),
        _ => Err(CalcValue::error(XlError::Value)),
    }
}

/// Coerces a function argument to text. Propagates errors.
pub fn require_text(arg: &FunctionArg<'_>) -> Result<String, CalcValue> {
    let scalar = require_scalar(arg);
    if let ScalarValue::Error(e) = scalar {
        return Err(CalcValue::error(*e));
    }
    match scalar.to_text() {
        ScalarValue::Text(s) => Ok(s),
        ScalarValue::Error(e) => Err(CalcValue::error(e)),
        _ => Err(CalcValue::error(XlError::Value)),
    }
}

/// Coerces a function argument to a boolean. Propagates errors.
pub fn require_bool(arg: &FunctionArg<'_>) -> Result<bool, CalcValue> {
    let scalar = require_scalar(arg);
    if let ScalarValue::Error(e) = scalar {
        return Err(CalcValue::error(*e));
    }
    match scalar.to_bool() {
        ScalarValue::Bool(b) => Ok(b),
        ScalarValue::Error(e) => Err(CalcValue::error(e)),
        _ => Err(CalcValue::error(XlError::Value)),
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn registry_case_insensitive_lookup() {
        let mut reg = FunctionRegistry::new();
        reg.register(FunctionDef {
            name: "SUM",
            min_args: 1,
            max_args: usize::MAX,
            param_kinds: &[ParamKind::Range],
            func: |_, _| CalcValue::number(0.0),
        });
        assert!(reg.get("SUM").is_some());
        assert!(reg.get("sum").is_some());
        assert!(reg.get("Sum").is_some());
        assert!(reg.get("NOTFOUND").is_none());
    }

    #[test]
    fn with_builtins_has_functions() {
        let reg = FunctionRegistry::with_builtins();
        // Spot-check a few functions from different modules
        assert!(reg.get("SUM").is_some());
        assert!(reg.get("AVERAGE").is_some());
        assert!(reg.get("LEFT").is_some());
        assert!(reg.get("IF").is_some());
        assert!(reg.get("VLOOKUP").is_some());
        assert!(reg.get("TODAY").is_some());
        assert!(reg.get("ISBLANK").is_some());
        assert!(reg.get("DAVERAGE").is_some());
        assert!(reg.get("DSUM").is_some());
        assert!(reg.get("FV").is_some());
        assert!(reg.get("PMT").is_some());
        assert!(reg.get("IPMT").is_some());
    }
}
