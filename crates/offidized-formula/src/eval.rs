use crate::ast::{BinaryOp, Expr, UnaryOp};
use crate::context::EvalContext;
use crate::functions::{FunctionArg, FunctionRegistry, ParamKind};
use crate::reference::RangeRef;
use crate::value::{compare_values, CalcValue, ScalarValue, XlError};

/// Evaluates a parsed expression tree against the given context and function registry.
///
/// This is the core tree-walking evaluator. Each `Expr` variant is handled by
/// dispatching to the appropriate evaluation logic. Errors in operands are
/// propagated immediately (short-circuit).
pub fn evaluate_expr(expr: &Expr, ctx: &EvalContext<'_>, registry: &FunctionRegistry) -> CalcValue {
    match expr {
        Expr::Literal(v) => CalcValue::Scalar(v.clone()),
        Expr::CellRef(r) => {
            let val = ctx.provider.cell_value(r.sheet.as_deref(), r.row, r.col);
            CalcValue::Scalar(val)
        }
        Expr::RangeRef(_) => {
            // Standalone range references are not valid outside of function arguments.
            CalcValue::error(XlError::Value)
        }
        Expr::Name(name) => eval_name(name, ctx, registry),
        Expr::Unary { op, expr: inner } => eval_unary(*op, inner, ctx, registry),
        Expr::Binary { op, left, right } => eval_binary(*op, left, right, ctx, registry),
        Expr::Function { name, args } => eval_function(name, args, ctx, registry),
    }
}

/// Evaluates a defined name reference.
fn eval_name(name: &str, ctx: &EvalContext<'_>, registry: &FunctionRegistry) -> CalcValue {
    match ctx.provider.resolve_name(name) {
        Some(resolved) => {
            // The resolved value is a string that could be a literal number, boolean,
            // error, or a formula. For now, treat it as a simple value.
            if let Ok(n) = resolved.parse::<f64>() {
                CalcValue::number(n)
            } else if resolved.eq_ignore_ascii_case("TRUE") {
                CalcValue::bool(true)
            } else if resolved.eq_ignore_ascii_case("FALSE") {
                CalcValue::bool(false)
            } else if let Some(err) = XlError::parse(&resolved) {
                CalcValue::error(err)
            } else {
                // Try to parse and evaluate as a formula expression.
                // If we have a lexer and parser available, use them.
                // For now, if it looks like a formula, try to evaluate it.
                try_eval_resolved(&resolved, ctx, registry)
            }
        }
        None => CalcValue::error(XlError::Name),
    }
}

/// Attempts to evaluate a resolved name value as a formula expression.
fn try_eval_resolved(
    resolved: &str,
    ctx: &EvalContext<'_>,
    registry: &FunctionRegistry,
) -> CalcValue {
    // If it starts with '=' strip it and try to tokenize/parse.
    let formula_text = resolved.strip_prefix('=').unwrap_or(resolved);

    // Try to lex and parse.
    let tokens = match crate::lexer::tokenize(formula_text) {
        Ok(t) => t,
        Err(_) => return CalcValue::text(resolved.to_string()),
    };
    match crate::parser::parse(tokens) {
        Ok(expr) => evaluate_expr(&expr, ctx, registry),
        Err(_) => CalcValue::text(resolved.to_string()),
    }
}

/// Evaluates a unary operation.
fn eval_unary(
    op: UnaryOp,
    inner: &Expr,
    ctx: &EvalContext<'_>,
    registry: &FunctionRegistry,
) -> CalcValue {
    let val = evaluate_expr(inner, ctx, registry).into_scalar();

    // Propagate errors.
    if let ScalarValue::Error(_) = &val {
        return CalcValue::Scalar(val);
    }

    match op {
        UnaryOp::Plus => {
            let coerced = val.to_number();
            CalcValue::Scalar(coerced)
        }
        UnaryOp::Negate => {
            let coerced = val.to_number();
            match coerced {
                ScalarValue::Number(n) => CalcValue::number(-n),
                other => CalcValue::Scalar(other),
            }
        }
        UnaryOp::Percent => {
            let coerced = val.to_number();
            match coerced {
                ScalarValue::Number(n) => CalcValue::number(n / 100.0),
                other => CalcValue::Scalar(other),
            }
        }
    }
}

/// Evaluates a binary operation.
fn eval_binary(
    op: BinaryOp,
    left: &Expr,
    right: &Expr,
    ctx: &EvalContext<'_>,
    registry: &FunctionRegistry,
) -> CalcValue {
    let lval = evaluate_expr(left, ctx, registry).into_scalar();

    // Propagate left error immediately.
    if let ScalarValue::Error(_) = &lval {
        return CalcValue::Scalar(lval);
    }

    let rval = evaluate_expr(right, ctx, registry).into_scalar();

    // Propagate right error.
    if let ScalarValue::Error(_) = &rval {
        return CalcValue::Scalar(rval);
    }

    match op {
        BinaryOp::Add | BinaryOp::Sub | BinaryOp::Mul | BinaryOp::Div | BinaryOp::Pow => {
            eval_arithmetic(op, &lval, &rval)
        }
        BinaryOp::Eq | BinaryOp::Ne | BinaryOp::Lt | BinaryOp::Le | BinaryOp::Gt | BinaryOp::Ge => {
            eval_comparison(op, &lval, &rval)
        }
        BinaryOp::Concat => eval_concat(&lval, &rval),
        BinaryOp::Range => {
            // Range operator is not supported as a standalone binary op.
            CalcValue::error(XlError::Value)
        }
    }
}

/// Evaluates arithmetic binary operations (+, -, *, /, ^).
fn eval_arithmetic(op: BinaryOp, lval: &ScalarValue, rval: &ScalarValue) -> CalcValue {
    let ln = lval.to_number();
    let rn = rval.to_number();

    // Propagate coercion errors.
    if let ScalarValue::Error(_) = &ln {
        return CalcValue::Scalar(ln);
    }
    if let ScalarValue::Error(_) = &rn {
        return CalcValue::Scalar(rn);
    }

    let left_n = match ln.as_number() {
        Some(n) => n,
        None => return CalcValue::error(XlError::Value),
    };
    let right_n = match rn.as_number() {
        Some(n) => n,
        None => return CalcValue::error(XlError::Value),
    };

    match op {
        BinaryOp::Add => CalcValue::number(left_n + right_n),
        BinaryOp::Sub => CalcValue::number(left_n - right_n),
        BinaryOp::Mul => CalcValue::number(left_n * right_n),
        BinaryOp::Div => {
            if right_n == 0.0 {
                CalcValue::error(XlError::Div0)
            } else {
                CalcValue::number(left_n / right_n)
            }
        }
        BinaryOp::Pow => {
            let result = left_n.powf(right_n);
            if result.is_nan() || result.is_infinite() {
                CalcValue::error(XlError::Num)
            } else {
                CalcValue::number(result)
            }
        }
        _ => CalcValue::error(XlError::Value),
    }
}

/// Evaluates comparison operators (=, <>, <, <=, >, >=).
fn eval_comparison(op: BinaryOp, lval: &ScalarValue, rval: &ScalarValue) -> CalcValue {
    let ord = compare_values(lval, rval);

    let result = match op {
        BinaryOp::Eq => ord == std::cmp::Ordering::Equal,
        BinaryOp::Ne => ord != std::cmp::Ordering::Equal,
        BinaryOp::Lt => ord == std::cmp::Ordering::Less,
        BinaryOp::Le => ord != std::cmp::Ordering::Greater,
        BinaryOp::Gt => ord == std::cmp::Ordering::Greater,
        BinaryOp::Ge => ord != std::cmp::Ordering::Less,
        _ => return CalcValue::error(XlError::Value),
    };

    CalcValue::bool(result)
}

/// Evaluates string concatenation (&).
fn eval_concat(lval: &ScalarValue, rval: &ScalarValue) -> CalcValue {
    let lt = lval.to_text();
    let rt = rval.to_text();

    // Propagate coercion errors.
    if let ScalarValue::Error(_) = &lt {
        return CalcValue::Scalar(lt);
    }
    if let ScalarValue::Error(_) = &rt {
        return CalcValue::Scalar(rt);
    }

    let left_s = match lt.as_text() {
        Some(s) => s,
        None => return CalcValue::error(XlError::Value),
    };
    let right_s = match rt.as_text() {
        Some(s) => s,
        None => return CalcValue::error(XlError::Value),
    };

    CalcValue::text(format!("{left_s}{right_s}"))
}

/// Evaluates a function call.
fn eval_function(
    name: &str,
    args: &[Expr],
    ctx: &EvalContext<'_>,
    registry: &FunctionRegistry,
) -> CalcValue {
    // Look up the function (case-insensitive).
    let upper = name.to_uppercase();
    let func_def = match registry.get(&upper) {
        Some(def) => def,
        None => return CalcValue::error(XlError::Name),
    };

    // Validate argument count.
    if args.len() < func_def.min_args || args.len() > func_def.max_args {
        return CalcValue::error(XlError::Value);
    }

    // Build the argument list.
    // Store converted CellRef->RangeRef so they live long enough to be referenced
    let mut temp_ranges: Vec<RangeRef> = Vec::new();

    // First pass: identify which args need conversion and pre-allocate temp_ranges
    let mut range_indices: Vec<Option<usize>> = Vec::with_capacity(args.len());
    for (i, arg) in args.iter().enumerate() {
        let kind = if func_def.param_kinds.is_empty() {
            &ParamKind::Scalar
        } else {
            let idx = if i < func_def.param_kinds.len() {
                i
            } else {
                func_def.param_kinds.len() - 1
            };
            &func_def.param_kinds[idx]
        };

        if matches!(kind, ParamKind::Range) {
            if let Expr::CellRef(cell_ref) = arg {
                let range_ref = RangeRef {
                    sheet: cell_ref.sheet.clone(),
                    start_row: cell_ref.row,
                    start_col: cell_ref.col,
                    end_row: cell_ref.row,
                    end_col: cell_ref.col,
                };
                let idx = temp_ranges.len();
                temp_ranges.push(range_ref);
                range_indices.push(Some(idx));
                continue;
            }
        }
        range_indices.push(None);
    }

    // Second pass: build func_args with stable references
    let mut func_args: Vec<FunctionArg<'_>> = Vec::with_capacity(args.len());
    for (i, arg) in args.iter().enumerate() {
        let kind = if func_def.param_kinds.is_empty() {
            &ParamKind::Scalar
        } else {
            let idx = if i < func_def.param_kinds.len() {
                i
            } else {
                func_def.param_kinds.len() - 1
            };
            &func_def.param_kinds[idx]
        };

        match kind {
            ParamKind::Range => {
                if let Some(idx) = range_indices[i] {
                    // CellRef converted to RangeRef in first pass
                    func_args.push(FunctionArg::Range {
                        range: &temp_ranges[idx],
                        ctx,
                    });
                } else if let Expr::RangeRef(range_ref) = arg {
                    // Direct RangeRef from AST
                    func_args.push(FunctionArg::Range {
                        range: range_ref,
                        ctx,
                    });
                } else {
                    // Otherwise evaluate and pass as a value
                    let val = evaluate_expr(arg, ctx, registry);
                    func_args.push(FunctionArg::Value(val));
                }
            }
            ParamKind::Scalar => {
                let val = evaluate_expr(arg, ctx, registry);
                func_args.push(FunctionArg::Value(val));
            }
        }
    }

    (func_def.func)(&func_args, ctx)
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::ast::Expr;
    use crate::context::{CellDataProvider, EvalContext};
    use crate::reference::{CellRef, RangeRef};
    use crate::value::{CalcValue, ScalarValue, XlError};
    use pretty_assertions::assert_eq;

    /// A mock cell data provider for testing.
    struct MockProvider {
        /// Fixed values keyed by (row, col).
        cells: Vec<(u32, u32, ScalarValue)>,
    }

    impl MockProvider {
        fn new(cells: Vec<(u32, u32, ScalarValue)>) -> Self {
            Self { cells }
        }
    }

    impl CellDataProvider for MockProvider {
        fn cell_value(&self, _sheet: Option<&str>, row: u32, col: u32) -> ScalarValue {
            for (r, c, v) in &self.cells {
                if *r == row && *c == col {
                    return v.clone();
                }
            }
            ScalarValue::Blank
        }
    }

    fn empty_registry() -> FunctionRegistry {
        FunctionRegistry::new()
    }

    fn make_ctx(provider: &dyn CellDataProvider) -> EvalContext<'_> {
        EvalContext::new(provider, None::<String>, 1, 1)
    }

    // --- Literal evaluation ---

    #[test]
    fn eval_literal_number() {
        let provider = MockProvider::new(vec![]);
        let ctx = make_ctx(&provider);
        let registry = empty_registry();
        let expr = Expr::Literal(ScalarValue::Number(42.0));
        assert_eq!(
            evaluate_expr(&expr, &ctx, &registry),
            CalcValue::number(42.0)
        );
    }

    #[test]
    fn eval_literal_text() {
        let provider = MockProvider::new(vec![]);
        let ctx = make_ctx(&provider);
        let registry = empty_registry();
        let expr = Expr::Literal(ScalarValue::Text("hello".to_string()));
        assert_eq!(
            evaluate_expr(&expr, &ctx, &registry),
            CalcValue::text("hello")
        );
    }

    #[test]
    fn eval_literal_bool() {
        let provider = MockProvider::new(vec![]);
        let ctx = make_ctx(&provider);
        let registry = empty_registry();
        let expr = Expr::Literal(ScalarValue::Bool(true));
        assert_eq!(evaluate_expr(&expr, &ctx, &registry), CalcValue::bool(true));
    }

    #[test]
    fn eval_literal_error() {
        let provider = MockProvider::new(vec![]);
        let ctx = make_ctx(&provider);
        let registry = empty_registry();
        let expr = Expr::Literal(ScalarValue::Error(XlError::Na));
        assert_eq!(
            evaluate_expr(&expr, &ctx, &registry),
            CalcValue::error(XlError::Na)
        );
    }

    #[test]
    fn eval_literal_blank() {
        let provider = MockProvider::new(vec![]);
        let ctx = make_ctx(&provider);
        let registry = empty_registry();
        let expr = Expr::Literal(ScalarValue::Blank);
        assert_eq!(evaluate_expr(&expr, &ctx, &registry), CalcValue::blank());
    }

    // --- Cell reference lookup ---

    #[test]
    fn eval_cell_ref_found() {
        let provider = MockProvider::new(vec![(1, 1, ScalarValue::Number(99.0))]);
        let ctx = make_ctx(&provider);
        let registry = empty_registry();
        let expr = Expr::CellRef(CellRef {
            sheet: None,
            col: 1,
            row: 1,
            col_absolute: false,
            row_absolute: false,
        });
        assert_eq!(
            evaluate_expr(&expr, &ctx, &registry),
            CalcValue::number(99.0)
        );
    }

    #[test]
    fn eval_cell_ref_blank() {
        let provider = MockProvider::new(vec![]);
        let ctx = make_ctx(&provider);
        let registry = empty_registry();
        let expr = Expr::CellRef(CellRef {
            sheet: None,
            col: 5,
            row: 5,
            col_absolute: false,
            row_absolute: false,
        });
        assert_eq!(evaluate_expr(&expr, &ctx, &registry), CalcValue::blank());
    }

    #[test]
    fn eval_cell_ref_with_sheet() {
        let provider = MockProvider::new(vec![(3, 2, ScalarValue::Text("data".to_string()))]);
        let ctx = make_ctx(&provider);
        let registry = empty_registry();
        let expr = Expr::CellRef(CellRef {
            sheet: Some("Sheet2".to_string()),
            col: 2,
            row: 3,
            col_absolute: true,
            row_absolute: true,
        });
        // The mock ignores sheet, but the plumbing still works.
        assert_eq!(
            evaluate_expr(&expr, &ctx, &registry),
            CalcValue::text("data")
        );
    }

    // --- Range ref (standalone) ---

    #[test]
    fn eval_standalone_range_ref() {
        let provider = MockProvider::new(vec![]);
        let ctx = make_ctx(&provider);
        let registry = empty_registry();
        let expr = Expr::RangeRef(RangeRef {
            sheet: None,
            start_col: 1,
            start_row: 1,
            end_col: 3,
            end_row: 3,
        });
        assert_eq!(
            evaluate_expr(&expr, &ctx, &registry),
            CalcValue::error(XlError::Value)
        );
    }

    // --- Arithmetic operations ---

    #[test]
    fn eval_addition() {
        let provider = MockProvider::new(vec![]);
        let ctx = make_ctx(&provider);
        let registry = empty_registry();
        let expr = Expr::Binary {
            op: BinaryOp::Add,
            left: Box::new(Expr::Literal(ScalarValue::Number(10.0))),
            right: Box::new(Expr::Literal(ScalarValue::Number(20.0))),
        };
        assert_eq!(
            evaluate_expr(&expr, &ctx, &registry),
            CalcValue::number(30.0)
        );
    }

    #[test]
    fn eval_subtraction() {
        let provider = MockProvider::new(vec![]);
        let ctx = make_ctx(&provider);
        let registry = empty_registry();
        let expr = Expr::Binary {
            op: BinaryOp::Sub,
            left: Box::new(Expr::Literal(ScalarValue::Number(50.0))),
            right: Box::new(Expr::Literal(ScalarValue::Number(8.0))),
        };
        assert_eq!(
            evaluate_expr(&expr, &ctx, &registry),
            CalcValue::number(42.0)
        );
    }

    #[test]
    fn eval_multiplication() {
        let provider = MockProvider::new(vec![]);
        let ctx = make_ctx(&provider);
        let registry = empty_registry();
        let expr = Expr::Binary {
            op: BinaryOp::Mul,
            left: Box::new(Expr::Literal(ScalarValue::Number(6.0))),
            right: Box::new(Expr::Literal(ScalarValue::Number(7.0))),
        };
        assert_eq!(
            evaluate_expr(&expr, &ctx, &registry),
            CalcValue::number(42.0)
        );
    }

    #[test]
    fn eval_division() {
        let provider = MockProvider::new(vec![]);
        let ctx = make_ctx(&provider);
        let registry = empty_registry();
        let expr = Expr::Binary {
            op: BinaryOp::Div,
            left: Box::new(Expr::Literal(ScalarValue::Number(84.0))),
            right: Box::new(Expr::Literal(ScalarValue::Number(2.0))),
        };
        assert_eq!(
            evaluate_expr(&expr, &ctx, &registry),
            CalcValue::number(42.0)
        );
    }

    #[test]
    fn eval_division_by_zero() {
        let provider = MockProvider::new(vec![]);
        let ctx = make_ctx(&provider);
        let registry = empty_registry();
        let expr = Expr::Binary {
            op: BinaryOp::Div,
            left: Box::new(Expr::Literal(ScalarValue::Number(1.0))),
            right: Box::new(Expr::Literal(ScalarValue::Number(0.0))),
        };
        assert_eq!(
            evaluate_expr(&expr, &ctx, &registry),
            CalcValue::error(XlError::Div0)
        );
    }

    #[test]
    fn eval_power() {
        let provider = MockProvider::new(vec![]);
        let ctx = make_ctx(&provider);
        let registry = empty_registry();
        let expr = Expr::Binary {
            op: BinaryOp::Pow,
            left: Box::new(Expr::Literal(ScalarValue::Number(2.0))),
            right: Box::new(Expr::Literal(ScalarValue::Number(10.0))),
        };
        assert_eq!(
            evaluate_expr(&expr, &ctx, &registry),
            CalcValue::number(1024.0)
        );
    }

    #[test]
    fn eval_power_error() {
        let provider = MockProvider::new(vec![]);
        let ctx = make_ctx(&provider);
        let registry = empty_registry();
        // (-1)^0.5 = NaN → #NUM!
        let expr = Expr::Binary {
            op: BinaryOp::Pow,
            left: Box::new(Expr::Literal(ScalarValue::Number(-1.0))),
            right: Box::new(Expr::Literal(ScalarValue::Number(0.5))),
        };
        assert_eq!(
            evaluate_expr(&expr, &ctx, &registry),
            CalcValue::error(XlError::Num)
        );
    }

    // --- Error propagation in arithmetic ---

    #[test]
    fn eval_error_propagation_left() {
        let provider = MockProvider::new(vec![]);
        let ctx = make_ctx(&provider);
        let registry = empty_registry();
        let expr = Expr::Binary {
            op: BinaryOp::Add,
            left: Box::new(Expr::Literal(ScalarValue::Error(XlError::Div0))),
            right: Box::new(Expr::Literal(ScalarValue::Number(1.0))),
        };
        assert_eq!(
            evaluate_expr(&expr, &ctx, &registry),
            CalcValue::error(XlError::Div0)
        );
    }

    #[test]
    fn eval_error_propagation_right() {
        let provider = MockProvider::new(vec![]);
        let ctx = make_ctx(&provider);
        let registry = empty_registry();
        let expr = Expr::Binary {
            op: BinaryOp::Add,
            left: Box::new(Expr::Literal(ScalarValue::Number(1.0))),
            right: Box::new(Expr::Literal(ScalarValue::Error(XlError::Na))),
        };
        assert_eq!(
            evaluate_expr(&expr, &ctx, &registry),
            CalcValue::error(XlError::Na)
        );
    }

    #[test]
    fn eval_error_propagation_both() {
        let provider = MockProvider::new(vec![]);
        let ctx = make_ctx(&provider);
        let registry = empty_registry();
        // Left error should win (short-circuit).
        let expr = Expr::Binary {
            op: BinaryOp::Mul,
            left: Box::new(Expr::Literal(ScalarValue::Error(XlError::Ref))),
            right: Box::new(Expr::Literal(ScalarValue::Error(XlError::Na))),
        };
        assert_eq!(
            evaluate_expr(&expr, &ctx, &registry),
            CalcValue::error(XlError::Ref)
        );
    }

    #[test]
    fn eval_text_coercion_error_in_arithmetic() {
        let provider = MockProvider::new(vec![]);
        let ctx = make_ctx(&provider);
        let registry = empty_registry();
        let expr = Expr::Binary {
            op: BinaryOp::Add,
            left: Box::new(Expr::Literal(ScalarValue::Text("abc".to_string()))),
            right: Box::new(Expr::Literal(ScalarValue::Number(1.0))),
        };
        assert_eq!(
            evaluate_expr(&expr, &ctx, &registry),
            CalcValue::error(XlError::Value)
        );
    }

    // --- Comparison operators ---

    #[test]
    fn eval_eq_true() {
        let provider = MockProvider::new(vec![]);
        let ctx = make_ctx(&provider);
        let registry = empty_registry();
        let expr = Expr::Binary {
            op: BinaryOp::Eq,
            left: Box::new(Expr::Literal(ScalarValue::Number(5.0))),
            right: Box::new(Expr::Literal(ScalarValue::Number(5.0))),
        };
        assert_eq!(evaluate_expr(&expr, &ctx, &registry), CalcValue::bool(true));
    }

    #[test]
    fn eval_eq_false() {
        let provider = MockProvider::new(vec![]);
        let ctx = make_ctx(&provider);
        let registry = empty_registry();
        let expr = Expr::Binary {
            op: BinaryOp::Eq,
            left: Box::new(Expr::Literal(ScalarValue::Number(5.0))),
            right: Box::new(Expr::Literal(ScalarValue::Number(6.0))),
        };
        assert_eq!(
            evaluate_expr(&expr, &ctx, &registry),
            CalcValue::bool(false)
        );
    }

    #[test]
    fn eval_ne() {
        let provider = MockProvider::new(vec![]);
        let ctx = make_ctx(&provider);
        let registry = empty_registry();
        let expr = Expr::Binary {
            op: BinaryOp::Ne,
            left: Box::new(Expr::Literal(ScalarValue::Number(5.0))),
            right: Box::new(Expr::Literal(ScalarValue::Number(6.0))),
        };
        assert_eq!(evaluate_expr(&expr, &ctx, &registry), CalcValue::bool(true));
    }

    #[test]
    fn eval_lt() {
        let provider = MockProvider::new(vec![]);
        let ctx = make_ctx(&provider);
        let registry = empty_registry();
        let expr = Expr::Binary {
            op: BinaryOp::Lt,
            left: Box::new(Expr::Literal(ScalarValue::Number(3.0))),
            right: Box::new(Expr::Literal(ScalarValue::Number(5.0))),
        };
        assert_eq!(evaluate_expr(&expr, &ctx, &registry), CalcValue::bool(true));
    }

    #[test]
    fn eval_le() {
        let provider = MockProvider::new(vec![]);
        let ctx = make_ctx(&provider);
        let registry = empty_registry();
        let expr = Expr::Binary {
            op: BinaryOp::Le,
            left: Box::new(Expr::Literal(ScalarValue::Number(5.0))),
            right: Box::new(Expr::Literal(ScalarValue::Number(5.0))),
        };
        assert_eq!(evaluate_expr(&expr, &ctx, &registry), CalcValue::bool(true));
    }

    #[test]
    fn eval_gt() {
        let provider = MockProvider::new(vec![]);
        let ctx = make_ctx(&provider);
        let registry = empty_registry();
        let expr = Expr::Binary {
            op: BinaryOp::Gt,
            left: Box::new(Expr::Literal(ScalarValue::Number(10.0))),
            right: Box::new(Expr::Literal(ScalarValue::Number(5.0))),
        };
        assert_eq!(evaluate_expr(&expr, &ctx, &registry), CalcValue::bool(true));
    }

    #[test]
    fn eval_ge() {
        let provider = MockProvider::new(vec![]);
        let ctx = make_ctx(&provider);
        let registry = empty_registry();
        let expr = Expr::Binary {
            op: BinaryOp::Ge,
            left: Box::new(Expr::Literal(ScalarValue::Number(5.0))),
            right: Box::new(Expr::Literal(ScalarValue::Number(5.0))),
        };
        assert_eq!(evaluate_expr(&expr, &ctx, &registry), CalcValue::bool(true));
    }

    #[test]
    fn eval_comparison_different_types() {
        let provider = MockProvider::new(vec![]);
        let ctx = make_ctx(&provider);
        let registry = empty_registry();
        // Number < Text in Excel comparison semantics.
        let expr = Expr::Binary {
            op: BinaryOp::Lt,
            left: Box::new(Expr::Literal(ScalarValue::Number(100.0))),
            right: Box::new(Expr::Literal(ScalarValue::Text("a".to_string()))),
        };
        assert_eq!(evaluate_expr(&expr, &ctx, &registry), CalcValue::bool(true));
    }

    #[test]
    fn eval_comparison_case_insensitive() {
        let provider = MockProvider::new(vec![]);
        let ctx = make_ctx(&provider);
        let registry = empty_registry();
        let expr = Expr::Binary {
            op: BinaryOp::Eq,
            left: Box::new(Expr::Literal(ScalarValue::Text("Hello".to_string()))),
            right: Box::new(Expr::Literal(ScalarValue::Text("hello".to_string()))),
        };
        assert_eq!(evaluate_expr(&expr, &ctx, &registry), CalcValue::bool(true));
    }

    // --- String concatenation ---

    #[test]
    fn eval_concat_strings() {
        let provider = MockProvider::new(vec![]);
        let ctx = make_ctx(&provider);
        let registry = empty_registry();
        let expr = Expr::Binary {
            op: BinaryOp::Concat,
            left: Box::new(Expr::Literal(ScalarValue::Text("Hello".to_string()))),
            right: Box::new(Expr::Literal(ScalarValue::Text(" World".to_string()))),
        };
        assert_eq!(
            evaluate_expr(&expr, &ctx, &registry),
            CalcValue::text("Hello World")
        );
    }

    #[test]
    fn eval_concat_number_and_text() {
        let provider = MockProvider::new(vec![]);
        let ctx = make_ctx(&provider);
        let registry = empty_registry();
        let expr = Expr::Binary {
            op: BinaryOp::Concat,
            left: Box::new(Expr::Literal(ScalarValue::Number(42.0))),
            right: Box::new(Expr::Literal(ScalarValue::Text(" items".to_string()))),
        };
        assert_eq!(
            evaluate_expr(&expr, &ctx, &registry),
            CalcValue::text("42 items")
        );
    }

    #[test]
    fn eval_concat_blank() {
        let provider = MockProvider::new(vec![]);
        let ctx = make_ctx(&provider);
        let registry = empty_registry();
        let expr = Expr::Binary {
            op: BinaryOp::Concat,
            left: Box::new(Expr::Literal(ScalarValue::Text("prefix".to_string()))),
            right: Box::new(Expr::Literal(ScalarValue::Blank)),
        };
        assert_eq!(
            evaluate_expr(&expr, &ctx, &registry),
            CalcValue::text("prefix")
        );
    }

    // --- Unary operators ---

    #[test]
    fn eval_unary_plus() {
        let provider = MockProvider::new(vec![]);
        let ctx = make_ctx(&provider);
        let registry = empty_registry();
        let expr = Expr::Unary {
            op: UnaryOp::Plus,
            expr: Box::new(Expr::Literal(ScalarValue::Number(5.0))),
        };
        assert_eq!(
            evaluate_expr(&expr, &ctx, &registry),
            CalcValue::number(5.0)
        );
    }

    #[test]
    fn eval_unary_negate() {
        let provider = MockProvider::new(vec![]);
        let ctx = make_ctx(&provider);
        let registry = empty_registry();
        let expr = Expr::Unary {
            op: UnaryOp::Negate,
            expr: Box::new(Expr::Literal(ScalarValue::Number(5.0))),
        };
        assert_eq!(
            evaluate_expr(&expr, &ctx, &registry),
            CalcValue::number(-5.0)
        );
    }

    #[test]
    fn eval_unary_percent() {
        let provider = MockProvider::new(vec![]);
        let ctx = make_ctx(&provider);
        let registry = empty_registry();
        let expr = Expr::Unary {
            op: UnaryOp::Percent,
            expr: Box::new(Expr::Literal(ScalarValue::Number(50.0))),
        };
        assert_eq!(
            evaluate_expr(&expr, &ctx, &registry),
            CalcValue::number(0.5)
        );
    }

    #[test]
    fn eval_unary_negate_blank() {
        let provider = MockProvider::new(vec![]);
        let ctx = make_ctx(&provider);
        let registry = empty_registry();
        // -Blank = -0.0 = 0.0 (well, -0.0 in IEEE)
        let expr = Expr::Unary {
            op: UnaryOp::Negate,
            expr: Box::new(Expr::Literal(ScalarValue::Blank)),
        };
        let result = evaluate_expr(&expr, &ctx, &registry);
        // -0.0 == 0.0 in f64, but let's just check it's a number.
        assert_eq!(result, CalcValue::number(-0.0));
    }

    #[test]
    fn eval_unary_error_propagation() {
        let provider = MockProvider::new(vec![]);
        let ctx = make_ctx(&provider);
        let registry = empty_registry();
        let expr = Expr::Unary {
            op: UnaryOp::Negate,
            expr: Box::new(Expr::Literal(ScalarValue::Error(XlError::Null))),
        };
        assert_eq!(
            evaluate_expr(&expr, &ctx, &registry),
            CalcValue::error(XlError::Null)
        );
    }

    #[test]
    fn eval_unary_plus_coerces_bool() {
        let provider = MockProvider::new(vec![]);
        let ctx = make_ctx(&provider);
        let registry = empty_registry();
        let expr = Expr::Unary {
            op: UnaryOp::Plus,
            expr: Box::new(Expr::Literal(ScalarValue::Bool(true))),
        };
        assert_eq!(
            evaluate_expr(&expr, &ctx, &registry),
            CalcValue::number(1.0)
        );
    }

    // --- Name resolution ---

    struct NameProvider;
    impl CellDataProvider for NameProvider {
        fn cell_value(&self, _sheet: Option<&str>, _row: u32, _col: u32) -> ScalarValue {
            ScalarValue::Blank
        }
        fn resolve_name(&self, name: &str) -> Option<String> {
            match name {
                "TaxRate" => Some("0.08".to_string()),
                "MyBool" => Some("TRUE".to_string()),
                _ => None,
            }
        }
    }

    #[test]
    fn eval_name_resolved_number() {
        let provider = NameProvider;
        let ctx = make_ctx(&provider);
        let registry = empty_registry();
        let expr = Expr::Name("TaxRate".to_string());
        assert_eq!(
            evaluate_expr(&expr, &ctx, &registry),
            CalcValue::number(0.08)
        );
    }

    #[test]
    fn eval_name_resolved_bool() {
        let provider = NameProvider;
        let ctx = make_ctx(&provider);
        let registry = empty_registry();
        let expr = Expr::Name("MyBool".to_string());
        assert_eq!(evaluate_expr(&expr, &ctx, &registry), CalcValue::bool(true));
    }

    #[test]
    fn eval_name_not_found() {
        let provider = NameProvider;
        let ctx = make_ctx(&provider);
        let registry = empty_registry();
        let expr = Expr::Name("Unknown".to_string());
        assert_eq!(
            evaluate_expr(&expr, &ctx, &registry),
            CalcValue::error(XlError::Name)
        );
    }

    // --- Function call (with empty registry) ---

    #[test]
    fn eval_unknown_function() {
        let provider = MockProvider::new(vec![]);
        let ctx = make_ctx(&provider);
        let registry = empty_registry();
        let expr = Expr::Function {
            name: "NONEXISTENT".to_string(),
            args: vec![],
        };
        assert_eq!(
            evaluate_expr(&expr, &ctx, &registry),
            CalcValue::error(XlError::Name)
        );
    }

    // --- Nested expressions ---

    #[test]
    fn eval_nested_arithmetic() {
        let provider = MockProvider::new(vec![]);
        let ctx = make_ctx(&provider);
        let registry = empty_registry();
        // (2 + 3) * 4 = 20
        let expr = Expr::Binary {
            op: BinaryOp::Mul,
            left: Box::new(Expr::Binary {
                op: BinaryOp::Add,
                left: Box::new(Expr::Literal(ScalarValue::Number(2.0))),
                right: Box::new(Expr::Literal(ScalarValue::Number(3.0))),
            }),
            right: Box::new(Expr::Literal(ScalarValue::Number(4.0))),
        };
        assert_eq!(
            evaluate_expr(&expr, &ctx, &registry),
            CalcValue::number(20.0)
        );
    }

    #[test]
    fn eval_cell_ref_in_arithmetic() {
        // A1=10, B1=20 → A1+B1 = 30
        let provider = MockProvider::new(vec![
            (1, 1, ScalarValue::Number(10.0)),
            (1, 2, ScalarValue::Number(20.0)),
        ]);
        let ctx = make_ctx(&provider);
        let registry = empty_registry();
        let expr = Expr::Binary {
            op: BinaryOp::Add,
            left: Box::new(Expr::CellRef(CellRef {
                sheet: None,
                col: 1,
                row: 1,
                col_absolute: false,
                row_absolute: false,
            })),
            right: Box::new(Expr::CellRef(CellRef {
                sheet: None,
                col: 2,
                row: 1,
                col_absolute: false,
                row_absolute: false,
            })),
        };
        assert_eq!(
            evaluate_expr(&expr, &ctx, &registry),
            CalcValue::number(30.0)
        );
    }

    #[test]
    fn eval_blank_in_arithmetic() {
        // Blank + 5 = 0 + 5 = 5 (blank coerces to 0)
        let provider = MockProvider::new(vec![]);
        let ctx = make_ctx(&provider);
        let registry = empty_registry();
        let expr = Expr::Binary {
            op: BinaryOp::Add,
            left: Box::new(Expr::Literal(ScalarValue::Blank)),
            right: Box::new(Expr::Literal(ScalarValue::Number(5.0))),
        };
        assert_eq!(
            evaluate_expr(&expr, &ctx, &registry),
            CalcValue::number(5.0)
        );
    }

    #[test]
    fn eval_bool_in_arithmetic() {
        // TRUE + 1 = 1 + 1 = 2
        let provider = MockProvider::new(vec![]);
        let ctx = make_ctx(&provider);
        let registry = empty_registry();
        let expr = Expr::Binary {
            op: BinaryOp::Add,
            left: Box::new(Expr::Literal(ScalarValue::Bool(true))),
            right: Box::new(Expr::Literal(ScalarValue::Number(1.0))),
        };
        assert_eq!(
            evaluate_expr(&expr, &ctx, &registry),
            CalcValue::number(2.0)
        );
    }

    // --- Range operator standalone ---

    #[test]
    fn eval_range_operator_standalone() {
        let provider = MockProvider::new(vec![]);
        let ctx = make_ctx(&provider);
        let registry = empty_registry();
        let expr = Expr::Binary {
            op: BinaryOp::Range,
            left: Box::new(Expr::Literal(ScalarValue::Number(1.0))),
            right: Box::new(Expr::Literal(ScalarValue::Number(10.0))),
        };
        assert_eq!(
            evaluate_expr(&expr, &ctx, &registry),
            CalcValue::error(XlError::Value)
        );
    }
}
