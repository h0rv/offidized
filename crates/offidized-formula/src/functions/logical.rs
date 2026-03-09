//! Logical functions: IF, AND, OR, NOT, IFERROR, TRUE, FALSE.

use crate::context::EvalContext;
use crate::value::{CalcValue, ScalarValue, XlError};

use super::{
    iter_range_values, require_bool, require_scalar, FunctionArg, FunctionDef, FunctionRegistry,
    ParamKind,
};

/// `IF(condition, value_if_true, [value_if_false=FALSE])`
///
/// Returns one of two values depending on a condition.
pub fn fn_if(args: &[FunctionArg<'_>], _ctx: &EvalContext<'_>) -> CalcValue {
    let condition = match args.first() {
        Some(a) => match require_bool(a) {
            Ok(b) => b,
            Err(e) => return e,
        },
        None => return CalcValue::error(XlError::Value),
    };

    if condition {
        match args.get(1) {
            Some(FunctionArg::Value(v)) => v.clone(),
            Some(FunctionArg::Range { .. }) => CalcValue::error(XlError::Value),
            None => CalcValue::bool(true),
        }
    } else {
        match args.get(2) {
            Some(FunctionArg::Value(v)) => v.clone(),
            Some(FunctionArg::Range { .. }) => CalcValue::error(XlError::Value),
            None => CalcValue::bool(false),
        }
    }
}

/// `AND(logical1, [logical2], ...)`
///
/// Returns TRUE if all arguments evaluate to TRUE. Accepts ranges.
/// Blanks in ranges are skipped.
pub fn fn_and(args: &[FunctionArg<'_>], _ctx: &EvalContext<'_>) -> CalcValue {
    let mut found_any = false;
    for arg in args {
        match arg {
            FunctionArg::Value(v) => {
                let scalar = v.as_scalar();
                if let ScalarValue::Error(e) = scalar {
                    return CalcValue::error(*e);
                }
                match scalar.to_bool() {
                    ScalarValue::Bool(b) => {
                        found_any = true;
                        if !b {
                            return CalcValue::bool(false);
                        }
                    }
                    ScalarValue::Error(e) => return CalcValue::error(e),
                    _ => {}
                }
            }
            FunctionArg::Range { range, ctx } => {
                for val in iter_range_values(range, ctx) {
                    match val {
                        ScalarValue::Error(e) => return CalcValue::error(e),
                        ScalarValue::Blank => {
                            // Skip blanks in ranges
                        }
                        ScalarValue::Text(_) => {
                            // Skip text in ranges (Excel AND ignores text in ranges)
                        }
                        other => match other.to_bool() {
                            ScalarValue::Bool(b) => {
                                found_any = true;
                                if !b {
                                    return CalcValue::bool(false);
                                }
                            }
                            ScalarValue::Error(e) => return CalcValue::error(e),
                            _ => {}
                        },
                    }
                }
            }
        }
    }
    if !found_any {
        return CalcValue::error(XlError::Value);
    }
    CalcValue::bool(true)
}

/// `OR(logical1, [logical2], ...)`
///
/// Returns TRUE if any argument evaluates to TRUE. Accepts ranges.
/// Blanks in ranges are skipped.
pub fn fn_or(args: &[FunctionArg<'_>], _ctx: &EvalContext<'_>) -> CalcValue {
    let mut found_any = false;
    for arg in args {
        match arg {
            FunctionArg::Value(v) => {
                let scalar = v.as_scalar();
                if let ScalarValue::Error(e) = scalar {
                    return CalcValue::error(*e);
                }
                match scalar.to_bool() {
                    ScalarValue::Bool(b) => {
                        found_any = true;
                        if b {
                            return CalcValue::bool(true);
                        }
                    }
                    ScalarValue::Error(e) => return CalcValue::error(e),
                    _ => {}
                }
            }
            FunctionArg::Range { range, ctx } => {
                for val in iter_range_values(range, ctx) {
                    match val {
                        ScalarValue::Error(e) => return CalcValue::error(e),
                        ScalarValue::Blank | ScalarValue::Text(_) => {
                            // Skip blanks and text in ranges
                        }
                        other => match other.to_bool() {
                            ScalarValue::Bool(b) => {
                                found_any = true;
                                if b {
                                    return CalcValue::bool(true);
                                }
                            }
                            ScalarValue::Error(e) => return CalcValue::error(e),
                            _ => {}
                        },
                    }
                }
            }
        }
    }
    if !found_any {
        return CalcValue::error(XlError::Value);
    }
    CalcValue::bool(false)
}

/// `NOT(logical)`
///
/// Negates a boolean value.
pub fn fn_not(args: &[FunctionArg<'_>], _ctx: &EvalContext<'_>) -> CalcValue {
    let b = match args.first() {
        Some(a) => match require_bool(a) {
            Ok(v) => v,
            Err(e) => return e,
        },
        None => return CalcValue::error(XlError::Value),
    };
    CalcValue::bool(!b)
}

/// `IFERROR(value, value_if_error)`
///
/// Returns `value` if it is not an error, otherwise returns `value_if_error`.
pub fn fn_iferror(args: &[FunctionArg<'_>], _ctx: &EvalContext<'_>) -> CalcValue {
    let value_scalar = match args.first() {
        Some(a) => require_scalar(a).clone(),
        None => return CalcValue::error(XlError::Value),
    };
    if value_scalar.is_error() {
        match args.get(1) {
            Some(FunctionArg::Value(v)) => v.clone(),
            Some(FunctionArg::Range { .. }) => CalcValue::error(XlError::Value),
            None => CalcValue::error(XlError::Value),
        }
    } else {
        CalcValue::Scalar(value_scalar)
    }
}

/// `TRUE()`
///
/// Returns the boolean value TRUE.
pub fn fn_true(_args: &[FunctionArg<'_>], _ctx: &EvalContext<'_>) -> CalcValue {
    CalcValue::bool(true)
}

/// `FALSE()`
///
/// Returns the boolean value FALSE.
pub fn fn_false(_args: &[FunctionArg<'_>], _ctx: &EvalContext<'_>) -> CalcValue {
    CalcValue::bool(false)
}

/// Registers all logical functions into the given registry.
pub fn register(registry: &mut FunctionRegistry) {
    registry.register(FunctionDef {
        name: "IF",
        min_args: 2,
        max_args: 3,
        param_kinds: &[ParamKind::Scalar, ParamKind::Scalar, ParamKind::Scalar],
        func: fn_if,
    });
    registry.register(FunctionDef {
        name: "AND",
        min_args: 1,
        max_args: usize::MAX,
        param_kinds: &[ParamKind::Range],
        func: fn_and,
    });
    registry.register(FunctionDef {
        name: "OR",
        min_args: 1,
        max_args: usize::MAX,
        param_kinds: &[ParamKind::Range],
        func: fn_or,
    });
    registry.register(FunctionDef {
        name: "NOT",
        min_args: 1,
        max_args: 1,
        param_kinds: &[ParamKind::Scalar],
        func: fn_not,
    });
    registry.register(FunctionDef {
        name: "IFERROR",
        min_args: 2,
        max_args: 2,
        param_kinds: &[ParamKind::Scalar, ParamKind::Scalar],
        func: fn_iferror,
    });
    registry.register(FunctionDef {
        name: "TRUE",
        min_args: 0,
        max_args: 0,
        param_kinds: &[],
        func: fn_true,
    });
    registry.register(FunctionDef {
        name: "FALSE",
        min_args: 0,
        max_args: 0,
        param_kinds: &[],
        func: fn_false,
    });
}
