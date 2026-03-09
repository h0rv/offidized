//! Information functions: ISBLANK, ISNUMBER, ISTEXT, ISERROR, ISERR, ISNA,
//! ISLOGICAL, ISNONTEXT, ISEVEN, ISODD, ISREF, N, NA, TYPE, ERROR.TYPE.

use crate::context::EvalContext;
use crate::value::{CalcValue, ScalarValue, XlError};

use super::{require_scalar, FunctionArg, FunctionDef, FunctionRegistry, ParamKind};

/// `ISBLANK(value)`
///
/// Returns TRUE if the value is blank (empty cell).
pub fn fn_isblank(args: &[FunctionArg<'_>], _ctx: &EvalContext<'_>) -> CalcValue {
    let scalar = match args.first() {
        Some(a) => require_scalar(a),
        None => return CalcValue::bool(false),
    };
    CalcValue::bool(scalar.is_blank())
}

/// `ISNUMBER(value)`
///
/// Returns TRUE if the value is a number.
pub fn fn_isnumber(args: &[FunctionArg<'_>], _ctx: &EvalContext<'_>) -> CalcValue {
    let scalar = match args.first() {
        Some(a) => require_scalar(a),
        None => return CalcValue::bool(false),
    };
    CalcValue::bool(matches!(scalar, ScalarValue::Number(_)))
}

/// `ISTEXT(value)`
///
/// Returns TRUE if the value is text.
pub fn fn_istext(args: &[FunctionArg<'_>], _ctx: &EvalContext<'_>) -> CalcValue {
    let scalar = match args.first() {
        Some(a) => require_scalar(a),
        None => return CalcValue::bool(false),
    };
    CalcValue::bool(matches!(scalar, ScalarValue::Text(_)))
}

/// `ISERROR(value)`
///
/// Returns TRUE if the value is any error type.
pub fn fn_iserror(args: &[FunctionArg<'_>], _ctx: &EvalContext<'_>) -> CalcValue {
    let scalar = match args.first() {
        Some(a) => require_scalar(a),
        None => return CalcValue::bool(false),
    };
    CalcValue::bool(scalar.is_error())
}

/// `ISERR(value)`
///
/// Returns TRUE if the value is any error except #N/A.
pub fn fn_iserr(args: &[FunctionArg<'_>], _ctx: &EvalContext<'_>) -> CalcValue {
    let scalar = match args.first() {
        Some(a) => require_scalar(a),
        None => return CalcValue::bool(false),
    };
    match scalar {
        ScalarValue::Error(e) => CalcValue::bool(*e != XlError::Na),
        _ => CalcValue::bool(false),
    }
}

/// `ISNA(value)`
///
/// Returns TRUE if the value is the #N/A error.
pub fn fn_isna(args: &[FunctionArg<'_>], _ctx: &EvalContext<'_>) -> CalcValue {
    let scalar = match args.first() {
        Some(a) => require_scalar(a),
        None => return CalcValue::bool(false),
    };
    match scalar {
        ScalarValue::Error(XlError::Na) => CalcValue::bool(true),
        _ => CalcValue::bool(false),
    }
}

/// `ISLOGICAL(value)`
///
/// Returns TRUE if the value is a logical value (TRUE or FALSE).
pub fn fn_islogical(args: &[FunctionArg<'_>], _ctx: &EvalContext<'_>) -> CalcValue {
    let scalar = match args.first() {
        Some(a) => require_scalar(a),
        None => return CalcValue::bool(false),
    };
    CalcValue::bool(matches!(scalar, ScalarValue::Bool(_)))
}

/// `ISNONTEXT(value)`
///
/// Returns TRUE if the value is not text (including blank, number, logical, error).
pub fn fn_isnontext(args: &[FunctionArg<'_>], _ctx: &EvalContext<'_>) -> CalcValue {
    let scalar = match args.first() {
        Some(a) => require_scalar(a),
        None => return CalcValue::bool(true),
    };
    CalcValue::bool(!matches!(scalar, ScalarValue::Text(_)))
}

/// `ISEVEN(value)`
///
/// Returns TRUE if the value is an even number. Logical values return #VALUE!.
pub fn fn_iseven(args: &[FunctionArg<'_>], _ctx: &EvalContext<'_>) -> CalcValue {
    let scalar = match args.first() {
        Some(a) => require_scalar(a),
        None => return CalcValue::error(XlError::Value),
    };

    // Blank cells return #N/A for ISEVEN/ISODD
    if scalar.is_blank() {
        return CalcValue::error(XlError::Na);
    }

    // Logical values return #VALUE!
    if matches!(scalar, ScalarValue::Bool(_)) {
        return CalcValue::error(XlError::Value);
    }

    // Convert to number
    let num = match scalar.to_number() {
        ScalarValue::Number(n) => n,
        ScalarValue::Error(e) => return CalcValue::error(e),
        _ => return CalcValue::error(XlError::Value),
    };

    // Truncate and check if even
    CalcValue::bool(num.trunc() % 2.0 == 0.0)
}

/// `ISODD(value)`
///
/// Returns TRUE if the value is an odd number. Logical values return #VALUE!.
pub fn fn_isodd(args: &[FunctionArg<'_>], _ctx: &EvalContext<'_>) -> CalcValue {
    let scalar = match args.first() {
        Some(a) => require_scalar(a),
        None => return CalcValue::error(XlError::Value),
    };

    // Blank cells return #N/A for ISEVEN/ISODD
    if scalar.is_blank() {
        return CalcValue::error(XlError::Na);
    }

    // Logical values return #VALUE!
    if matches!(scalar, ScalarValue::Bool(_)) {
        return CalcValue::error(XlError::Value);
    }

    // Convert to number
    let num = match scalar.to_number() {
        ScalarValue::Number(n) => n,
        ScalarValue::Error(e) => return CalcValue::error(e),
        _ => return CalcValue::error(XlError::Value),
    };

    // Truncate and check if odd
    CalcValue::bool(num.trunc() % 2.0 != 0.0)
}

/// `ISREF(value)`
///
/// Returns TRUE if the value is a reference. In our current scalar-only
/// implementation, this checks if the argument is a range.
pub fn fn_isref(args: &[FunctionArg<'_>], _ctx: &EvalContext<'_>) -> CalcValue {
    match args.first() {
        Some(FunctionArg::Range { .. }) => CalcValue::bool(true),
        Some(FunctionArg::Value(_)) => CalcValue::bool(false),
        None => CalcValue::bool(false),
    }
}

/// `N(value)`
///
/// Converts a value to a number, returning 0 for blank and text, 1/0 for
/// logical, and propagating errors.
pub fn fn_n(args: &[FunctionArg<'_>], _ctx: &EvalContext<'_>) -> CalcValue {
    let scalar = match args.first() {
        Some(a) => require_scalar(a),
        None => return CalcValue::number(0.0),
    };

    match scalar {
        ScalarValue::Number(n) => CalcValue::number(*n),
        ScalarValue::Bool(b) => CalcValue::number(if *b { 1.0 } else { 0.0 }),
        ScalarValue::Error(e) => CalcValue::error(*e),
        ScalarValue::Blank | ScalarValue::Text(_) => CalcValue::number(0.0),
    }
}

/// `NA()`
///
/// Returns the #N/A error value.
pub fn fn_na(_args: &[FunctionArg<'_>], _ctx: &EvalContext<'_>) -> CalcValue {
    CalcValue::error(XlError::Na)
}

/// `TYPE(value)`
///
/// Returns a number indicating the data type of a value:
/// - 1 = Number (or blank)
/// - 2 = Text
/// - 4 = Logical
/// - 16 = Error
/// - 64 = Array (not yet supported in this implementation)
pub fn fn_type(args: &[FunctionArg<'_>], _ctx: &EvalContext<'_>) -> CalcValue {
    match args.first() {
        Some(FunctionArg::Range { .. }) => {
            // In Excel, a reference to a single cell returns the type of that cell's value.
            // A reference to multiple cells returns 64 (array). For now, we return 64.
            CalcValue::number(64.0)
        }
        Some(FunctionArg::Value(v)) => {
            let scalar = v.as_scalar();
            match scalar {
                ScalarValue::Blank | ScalarValue::Number(_) => CalcValue::number(1.0),
                ScalarValue::Text(_) => CalcValue::number(2.0),
                ScalarValue::Bool(_) => CalcValue::number(4.0),
                ScalarValue::Error(_) => CalcValue::number(16.0),
            }
        }
        None => CalcValue::number(1.0),
    }
}

/// `ERROR.TYPE(error_val)`
///
/// Returns a number corresponding to the error type:
/// - 1 = #NULL!
/// - 2 = #DIV/0!
/// - 3 = #VALUE!
/// - 4 = #REF!
/// - 5 = #NAME?
/// - 6 = #NUM!
/// - 7 = #N/A
///
/// If the value is not an error, returns #N/A.
pub fn fn_error_type(args: &[FunctionArg<'_>], _ctx: &EvalContext<'_>) -> CalcValue {
    let scalar = match args.first() {
        Some(a) => require_scalar(a),
        None => return CalcValue::error(XlError::Na),
    };

    match scalar {
        ScalarValue::Error(e) => {
            let error_code = match e {
                XlError::Null => 1,
                XlError::Div0 => 2,
                XlError::Value => 3,
                XlError::Ref => 4,
                XlError::Name => 5,
                XlError::Num => 6,
                XlError::Na => 7,
            };
            CalcValue::number(error_code as f64)
        }
        _ => CalcValue::error(XlError::Na),
    }
}

/// `CELL(info_type, [reference])`
///
/// Returns information about a cell. Supported info types:
/// - "address" → Cell address as text (e.g., "$A$1" or "Sheet1!$B$2")
/// - "col" → Column number (1-based)
/// - "row" → Row number (1-based)
/// - "contents" → Cell value as text
/// - "type" → "b" (blank), "l" (label/text), "v" (value)
/// - "format" → Number format code string
/// - "width" → Column width (approximate character width)
///
/// If reference is omitted, uses the last changed cell (we use formula cell).
pub fn fn_cell(args: &[FunctionArg<'_>], ctx: &EvalContext<'_>) -> CalcValue {
    // First argument: info_type (required)
    let info_type = match args.first() {
        Some(a) => {
            let scalar = require_scalar(a);
            match scalar {
                ScalarValue::Text(s) => s.to_lowercase(),
                _ => return CalcValue::error(XlError::Value),
            }
        }
        None => return CalcValue::error(XlError::Value),
    };

    // Second argument: reference (optional, defaults to formula cell)
    let (sheet, row, col) = match args.get(1) {
        Some(FunctionArg::Range { range, .. }) => {
            // Use top-left cell of range
            let row = range.start_row;
            let col = range.start_col;
            // Use range's sheet or fallback to current sheet
            let sheet = range.sheet.as_deref().or(ctx.current_sheet.as_deref());
            (sheet, row, col)
        }
        Some(FunctionArg::Value(_)) => {
            // If a value is passed, it's an error
            return CalcValue::error(XlError::Value);
        }
        None => {
            // No reference provided, use formula cell
            (
                ctx.current_sheet.as_deref(),
                ctx.formula_row,
                ctx.formula_col,
            )
        }
    };

    // Query the provider for cell info
    match ctx.provider.cell_info(sheet, row, col, &info_type) {
        Some(result) => CalcValue::text(result),
        None => CalcValue::error(XlError::Value),
    }
}

/// `INFO(type_text)`
///
/// Returns information about the environment. Supported types:
/// - "numfile" → Number of worksheets
/// - "recalc" → "Automatic" (always auto-recalc)
/// - "release" → Version string
/// - "system" → "mac" or "pcdos"
/// - "osversion" → OS name
/// - "directory" → Returns #N/A (no filesystem context)
/// - "origin" → Returns #N/A (no UI context)
pub fn fn_info(args: &[FunctionArg<'_>], ctx: &EvalContext<'_>) -> CalcValue {
    let info_type = match args.first() {
        Some(a) => {
            let scalar = require_scalar(a);
            match scalar {
                ScalarValue::Text(s) => s.to_lowercase(),
                _ => return CalcValue::error(XlError::Value),
            }
        }
        None => return CalcValue::error(XlError::Value),
    };

    // Special cases that don't need provider
    match info_type.as_str() {
        "directory" | "origin" => return CalcValue::error(XlError::Na),
        _ => {}
    }

    // Query the provider for workbook info
    match ctx.provider.workbook_info(&info_type) {
        Some(result) => CalcValue::text(result),
        None => CalcValue::error(XlError::Value),
    }
}

/// Registers all information functions into the given registry.
pub fn register(registry: &mut FunctionRegistry) {
    registry.register(FunctionDef {
        name: "ISBLANK",
        min_args: 1,
        max_args: 1,
        param_kinds: &[ParamKind::Scalar],
        func: fn_isblank,
    });
    registry.register(FunctionDef {
        name: "ISNUMBER",
        min_args: 1,
        max_args: 1,
        param_kinds: &[ParamKind::Scalar],
        func: fn_isnumber,
    });
    registry.register(FunctionDef {
        name: "ISTEXT",
        min_args: 1,
        max_args: 1,
        param_kinds: &[ParamKind::Scalar],
        func: fn_istext,
    });
    registry.register(FunctionDef {
        name: "ISERROR",
        min_args: 1,
        max_args: 1,
        param_kinds: &[ParamKind::Scalar],
        func: fn_iserror,
    });
    registry.register(FunctionDef {
        name: "ISERR",
        min_args: 1,
        max_args: 1,
        param_kinds: &[ParamKind::Scalar],
        func: fn_iserr,
    });
    registry.register(FunctionDef {
        name: "ISNA",
        min_args: 1,
        max_args: 1,
        param_kinds: &[ParamKind::Scalar],
        func: fn_isna,
    });
    registry.register(FunctionDef {
        name: "ISLOGICAL",
        min_args: 1,
        max_args: 1,
        param_kinds: &[ParamKind::Scalar],
        func: fn_islogical,
    });
    registry.register(FunctionDef {
        name: "ISNONTEXT",
        min_args: 1,
        max_args: 1,
        param_kinds: &[ParamKind::Scalar],
        func: fn_isnontext,
    });
    registry.register(FunctionDef {
        name: "ISEVEN",
        min_args: 1,
        max_args: 1,
        param_kinds: &[ParamKind::Scalar],
        func: fn_iseven,
    });
    registry.register(FunctionDef {
        name: "ISODD",
        min_args: 1,
        max_args: 1,
        param_kinds: &[ParamKind::Scalar],
        func: fn_isodd,
    });
    registry.register(FunctionDef {
        name: "ISREF",
        min_args: 1,
        max_args: 1,
        param_kinds: &[ParamKind::Range], // Must accept range to detect if it's a reference
        func: fn_isref,
    });
    registry.register(FunctionDef {
        name: "N",
        min_args: 1,
        max_args: 1,
        param_kinds: &[ParamKind::Scalar],
        func: fn_n,
    });
    registry.register(FunctionDef {
        name: "NA",
        min_args: 0,
        max_args: 0,
        param_kinds: &[],
        func: fn_na,
    });
    registry.register(FunctionDef {
        name: "TYPE",
        min_args: 1,
        max_args: 1,
        param_kinds: &[ParamKind::Range], // Must accept range to detect array vs scalar
        func: fn_type,
    });
    registry.register(FunctionDef {
        name: "ERROR.TYPE",
        min_args: 1,
        max_args: 1,
        param_kinds: &[ParamKind::Scalar],
        func: fn_error_type,
    });
    registry.register(FunctionDef {
        name: "CELL",
        min_args: 1,
        max_args: 2,
        param_kinds: &[ParamKind::Scalar, ParamKind::Range],
        func: fn_cell,
    });
    registry.register(FunctionDef {
        name: "INFO",
        min_args: 1,
        max_args: 1,
        param_kinds: &[ParamKind::Scalar],
        func: fn_info,
    });
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::context::CellDataProvider;
    use pretty_assertions::assert_eq;

    struct DummyProvider;
    impl CellDataProvider for DummyProvider {
        fn cell_value(&self, _sheet: Option<&str>, _row: u32, _col: u32) -> ScalarValue {
            ScalarValue::Blank
        }
    }

    fn test_ctx() -> EvalContext<'static> {
        static PROVIDER: DummyProvider = DummyProvider;
        EvalContext::new(&PROVIDER, None::<String>, 1, 1)
    }

    #[test]
    fn test_isblank() {
        let ctx = test_ctx();
        assert_eq!(
            fn_isblank(&[FunctionArg::Value(CalcValue::blank())], &ctx),
            CalcValue::bool(true)
        );
        assert_eq!(
            fn_isblank(&[FunctionArg::Value(CalcValue::number(42.0))], &ctx),
            CalcValue::bool(false)
        );
        assert_eq!(
            fn_isblank(&[FunctionArg::Value(CalcValue::text(""))], &ctx),
            CalcValue::bool(false)
        );
    }

    #[test]
    fn test_isnumber() {
        let ctx = test_ctx();
        assert_eq!(
            fn_isnumber(&[FunctionArg::Value(CalcValue::number(42.0))], &ctx),
            CalcValue::bool(true)
        );
        assert_eq!(
            fn_isnumber(&[FunctionArg::Value(CalcValue::text("42"))], &ctx),
            CalcValue::bool(false)
        );
        assert_eq!(
            fn_isnumber(&[FunctionArg::Value(CalcValue::bool(true))], &ctx),
            CalcValue::bool(false)
        );
    }

    #[test]
    fn test_istext() {
        let ctx = test_ctx();
        assert_eq!(
            fn_istext(&[FunctionArg::Value(CalcValue::text("hello"))], &ctx),
            CalcValue::bool(true)
        );
        assert_eq!(
            fn_istext(&[FunctionArg::Value(CalcValue::number(42.0))], &ctx),
            CalcValue::bool(false)
        );
        assert_eq!(
            fn_istext(&[FunctionArg::Value(CalcValue::blank())], &ctx),
            CalcValue::bool(false)
        );
    }

    #[test]
    fn test_iserror() {
        let ctx = test_ctx();
        assert_eq!(
            fn_iserror(&[FunctionArg::Value(CalcValue::error(XlError::Div0))], &ctx),
            CalcValue::bool(true)
        );
        assert_eq!(
            fn_iserror(&[FunctionArg::Value(CalcValue::error(XlError::Na))], &ctx),
            CalcValue::bool(true)
        );
        assert_eq!(
            fn_iserror(&[FunctionArg::Value(CalcValue::number(42.0))], &ctx),
            CalcValue::bool(false)
        );
    }

    #[test]
    fn test_iserr() {
        let ctx = test_ctx();
        assert_eq!(
            fn_iserr(&[FunctionArg::Value(CalcValue::error(XlError::Div0))], &ctx),
            CalcValue::bool(true)
        );
        assert_eq!(
            fn_iserr(&[FunctionArg::Value(CalcValue::error(XlError::Na))], &ctx),
            CalcValue::bool(false)
        );
        assert_eq!(
            fn_iserr(&[FunctionArg::Value(CalcValue::number(42.0))], &ctx),
            CalcValue::bool(false)
        );
    }

    #[test]
    fn test_isna() {
        let ctx = test_ctx();
        assert_eq!(
            fn_isna(&[FunctionArg::Value(CalcValue::error(XlError::Na))], &ctx),
            CalcValue::bool(true)
        );
        assert_eq!(
            fn_isna(&[FunctionArg::Value(CalcValue::error(XlError::Div0))], &ctx),
            CalcValue::bool(false)
        );
        assert_eq!(
            fn_isna(&[FunctionArg::Value(CalcValue::number(42.0))], &ctx),
            CalcValue::bool(false)
        );
    }

    #[test]
    fn test_islogical() {
        let ctx = test_ctx();
        assert_eq!(
            fn_islogical(&[FunctionArg::Value(CalcValue::bool(true))], &ctx),
            CalcValue::bool(true)
        );
        assert_eq!(
            fn_islogical(&[FunctionArg::Value(CalcValue::bool(false))], &ctx),
            CalcValue::bool(true)
        );
        assert_eq!(
            fn_islogical(&[FunctionArg::Value(CalcValue::number(1.0))], &ctx),
            CalcValue::bool(false)
        );
        assert_eq!(
            fn_islogical(&[FunctionArg::Value(CalcValue::text("TRUE"))], &ctx),
            CalcValue::bool(false)
        );
    }

    #[test]
    fn test_isnontext() {
        let ctx = test_ctx();
        assert_eq!(
            fn_isnontext(&[FunctionArg::Value(CalcValue::text("hello"))], &ctx),
            CalcValue::bool(false)
        );
        assert_eq!(
            fn_isnontext(&[FunctionArg::Value(CalcValue::number(42.0))], &ctx),
            CalcValue::bool(true)
        );
        assert_eq!(
            fn_isnontext(&[FunctionArg::Value(CalcValue::bool(true))], &ctx),
            CalcValue::bool(true)
        );
        assert_eq!(
            fn_isnontext(&[FunctionArg::Value(CalcValue::blank())], &ctx),
            CalcValue::bool(true)
        );
        assert_eq!(
            fn_isnontext(&[FunctionArg::Value(CalcValue::error(XlError::Na))], &ctx),
            CalcValue::bool(true)
        );
    }

    #[test]
    fn test_iseven() {
        let ctx = test_ctx();
        assert_eq!(
            fn_iseven(&[FunctionArg::Value(CalcValue::number(2.0))], &ctx),
            CalcValue::bool(true)
        );
        assert_eq!(
            fn_iseven(&[FunctionArg::Value(CalcValue::number(3.0))], &ctx),
            CalcValue::bool(false)
        );
        assert_eq!(
            fn_iseven(&[FunctionArg::Value(CalcValue::number(-2.0))], &ctx),
            CalcValue::bool(true)
        );
        assert_eq!(
            fn_iseven(&[FunctionArg::Value(CalcValue::number(2.5))], &ctx),
            CalcValue::bool(true)
        ); // Truncates to 2
        assert_eq!(
            fn_iseven(&[FunctionArg::Value(CalcValue::blank())], &ctx),
            CalcValue::error(XlError::Na)
        );
        assert_eq!(
            fn_iseven(&[FunctionArg::Value(CalcValue::bool(true))], &ctx),
            CalcValue::error(XlError::Value)
        );
    }

    #[test]
    fn test_isodd() {
        let ctx = test_ctx();
        assert_eq!(
            fn_isodd(&[FunctionArg::Value(CalcValue::number(3.0))], &ctx),
            CalcValue::bool(true)
        );
        assert_eq!(
            fn_isodd(&[FunctionArg::Value(CalcValue::number(2.0))], &ctx),
            CalcValue::bool(false)
        );
        assert_eq!(
            fn_isodd(&[FunctionArg::Value(CalcValue::number(-3.0))], &ctx),
            CalcValue::bool(true)
        );
        assert_eq!(
            fn_isodd(&[FunctionArg::Value(CalcValue::number(3.7))], &ctx),
            CalcValue::bool(true)
        ); // Truncates to 3
        assert_eq!(
            fn_isodd(&[FunctionArg::Value(CalcValue::blank())], &ctx),
            CalcValue::error(XlError::Na)
        );
        assert_eq!(
            fn_isodd(&[FunctionArg::Value(CalcValue::bool(false))], &ctx),
            CalcValue::error(XlError::Value)
        );
    }

    #[test]
    fn test_n() {
        let ctx = test_ctx();
        assert_eq!(
            fn_n(&[FunctionArg::Value(CalcValue::number(42.5))], &ctx),
            CalcValue::number(42.5)
        );
        assert_eq!(
            fn_n(&[FunctionArg::Value(CalcValue::bool(true))], &ctx),
            CalcValue::number(1.0)
        );
        assert_eq!(
            fn_n(&[FunctionArg::Value(CalcValue::bool(false))], &ctx),
            CalcValue::number(0.0)
        );
        assert_eq!(
            fn_n(&[FunctionArg::Value(CalcValue::text("hello"))], &ctx),
            CalcValue::number(0.0)
        );
        assert_eq!(
            fn_n(&[FunctionArg::Value(CalcValue::blank())], &ctx),
            CalcValue::number(0.0)
        );
        assert_eq!(
            fn_n(&[FunctionArg::Value(CalcValue::error(XlError::Div0))], &ctx),
            CalcValue::error(XlError::Div0)
        );
    }

    #[test]
    fn test_na() {
        let ctx = test_ctx();
        assert_eq!(fn_na(&[], &ctx), CalcValue::error(XlError::Na));
    }

    #[test]
    fn test_type() {
        let ctx = test_ctx();
        assert_eq!(
            fn_type(&[FunctionArg::Value(CalcValue::number(42.0))], &ctx),
            CalcValue::number(1.0)
        );
        assert_eq!(
            fn_type(&[FunctionArg::Value(CalcValue::blank())], &ctx),
            CalcValue::number(1.0)
        );
        assert_eq!(
            fn_type(&[FunctionArg::Value(CalcValue::text("hello"))], &ctx),
            CalcValue::number(2.0)
        );
        assert_eq!(
            fn_type(&[FunctionArg::Value(CalcValue::bool(true))], &ctx),
            CalcValue::number(4.0)
        );
        assert_eq!(
            fn_type(&[FunctionArg::Value(CalcValue::error(XlError::Na))], &ctx),
            CalcValue::number(16.0)
        );
    }

    #[test]
    fn test_error_type() {
        let ctx = test_ctx();
        assert_eq!(
            fn_error_type(&[FunctionArg::Value(CalcValue::error(XlError::Null))], &ctx),
            CalcValue::number(1.0)
        );
        assert_eq!(
            fn_error_type(&[FunctionArg::Value(CalcValue::error(XlError::Div0))], &ctx),
            CalcValue::number(2.0)
        );
        assert_eq!(
            fn_error_type(
                &[FunctionArg::Value(CalcValue::error(XlError::Value))],
                &ctx
            ),
            CalcValue::number(3.0)
        );
        assert_eq!(
            fn_error_type(&[FunctionArg::Value(CalcValue::error(XlError::Ref))], &ctx),
            CalcValue::number(4.0)
        );
        assert_eq!(
            fn_error_type(&[FunctionArg::Value(CalcValue::error(XlError::Name))], &ctx),
            CalcValue::number(5.0)
        );
        assert_eq!(
            fn_error_type(&[FunctionArg::Value(CalcValue::error(XlError::Num))], &ctx),
            CalcValue::number(6.0)
        );
        assert_eq!(
            fn_error_type(&[FunctionArg::Value(CalcValue::error(XlError::Na))], &ctx),
            CalcValue::number(7.0)
        );
        assert_eq!(
            fn_error_type(&[FunctionArg::Value(CalcValue::number(42.0))], &ctx),
            CalcValue::error(XlError::Na)
        );
    }

    #[test]
    fn test_cell_basic() {
        // CELL function with no provider support returns #VALUE!
        let ctx = test_ctx();
        assert_eq!(
            fn_cell(&[FunctionArg::Value(CalcValue::text("col"))], &ctx),
            CalcValue::error(XlError::Value)
        );
    }

    #[test]
    fn test_cell_invalid_info_type() {
        let ctx = test_ctx();
        assert_eq!(
            fn_cell(&[FunctionArg::Value(CalcValue::number(42.0))], &ctx),
            CalcValue::error(XlError::Value)
        );
    }

    #[test]
    fn test_info_basic() {
        // INFO function with no provider support returns #VALUE!
        let ctx = test_ctx();
        assert_eq!(
            fn_info(&[FunctionArg::Value(CalcValue::text("recalc"))], &ctx),
            CalcValue::error(XlError::Value)
        );
    }

    #[test]
    fn test_info_special_cases() {
        let ctx = test_ctx();
        // "directory" and "origin" always return #N/A
        assert_eq!(
            fn_info(&[FunctionArg::Value(CalcValue::text("directory"))], &ctx),
            CalcValue::error(XlError::Na)
        );
        assert_eq!(
            fn_info(&[FunctionArg::Value(CalcValue::text("origin"))], &ctx),
            CalcValue::error(XlError::Na)
        );
    }

    #[test]
    fn test_info_invalid_type() {
        let ctx = test_ctx();
        assert_eq!(
            fn_info(&[FunctionArg::Value(CalcValue::number(42.0))], &ctx),
            CalcValue::error(XlError::Value)
        );
    }
}
