//! Excel formula parser and evaluator.
//!
//! This crate provides a standalone formula engine that can parse and evaluate
//! Excel formulas. It has no dependency on `offidized-xlsx` or `offidized-opc` —
//! consumers implement the [`CellDataProvider`] trait to bridge their data model.
//!
//! # Function Coverage
//!
//! **227 Excel functions** spanning all major categories:
//!
//! - **Math & Trig** — Arithmetic, rounding, trigonometry, factorials, base conversion
//! - **Statistical** — Averages, distributions, standard deviations, variance, criteria counting
//! - **Text** — String manipulation, formatting, searching, concatenation
//! - **Logical** — Conditional logic, boolean operations, error handling
//! - **Date & Time** — Date arithmetic, extraction, formatting, business day calculations
//! - **Lookup & Reference** — Lookups, cell references, array operations
//! - **Information** — Type checking, error type identification, cell metadata
//! - **Database** — Criteria-based aggregation on structured data
//! - **Financial** — Basic time-value-of-money calculations
//!
//! # Quick start
//!
//! ```
//! use offidized_formula::{evaluate, CellDataProvider, EvalContext, ScalarValue};
//!
//! struct MyData;
//! impl CellDataProvider for MyData {
//!     fn cell_value(&self, _sheet: Option<&str>, _row: u32, _col: u32) -> ScalarValue {
//!         ScalarValue::Number(10.0)
//!     }
//! }
//!
//! let ctx = EvalContext::new(&MyData, None::<String>, 1, 1);
//! let result = evaluate("=1+2+3", &ctx);
//! assert_eq!(result.into_scalar(), ScalarValue::Number(6.0));
//! ```
//!
//! # Features
//!
//! - **Hand-written recursive descent parser** — no proc macros, fast compilation
//! - **Tree-walking evaluator** — simple and debuggable
//! - **Excel-compatible semantics** — follows Excel's type coercion, error propagation, and edge cases
//! - **Extensible** — register custom functions via [`FunctionRegistry`](functions::FunctionRegistry)
//! - **Zero dependencies on offidized** — use standalone or integrate with any data source

pub mod ast;
pub mod context;
pub mod error;
pub mod eval;
pub mod functions;
pub mod lexer;
pub mod parser;
pub mod reference;
pub mod token;
pub mod value;

pub use context::{CellDataProvider, EvalContext};
pub use error::{FormulaError, Result};
pub use reference::{CellRef, RangeRef};
pub use value::{CalcValue, ScalarValue, XlError};

use functions::FunctionRegistry;

/// Parses and evaluates a formula string, returning the computed value.
///
/// This is the main entry point for formula evaluation. The formula string
/// may optionally start with `=`.
pub fn evaluate(formula: &str, ctx: &EvalContext<'_>) -> CalcValue {
    let registry = FunctionRegistry::with_builtins();
    evaluate_with_registry(formula, ctx, &registry)
}

/// Parses and evaluates a formula string using a custom function registry.
pub fn evaluate_with_registry(
    formula: &str,
    ctx: &EvalContext<'_>,
    registry: &FunctionRegistry,
) -> CalcValue {
    let tokens = match lexer::tokenize(formula) {
        Ok(t) => t,
        Err(e) => {
            tracing::warn!("formula lex error: {e}");
            return CalcValue::error(XlError::Name);
        }
    };
    let expr = match parser::parse(tokens) {
        Ok(e) => e,
        Err(e) => {
            tracing::warn!("formula parse error: {e}");
            return CalcValue::error(XlError::Name);
        }
    };
    eval::evaluate_expr(&expr, ctx, registry)
}
