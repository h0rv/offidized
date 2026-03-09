//! Mathematical and trigonometric functions.

use super::{
    collect_numbers, require_number, FunctionArg, FunctionDef, FunctionRegistry, ParamKind,
};
use crate::context::EvalContext;
use crate::value::{CalcValue, XlError};

mod aggregate;
mod trig;

// Re-export extended functions
pub use aggregate::*;
pub use trig::*;

// ============================================================================
// BASIC MATH FUNCTIONS (original 13)
// ============================================================================

pub fn fn_sum(args: &[FunctionArg<'_>], ctx: &EvalContext<'_>) -> CalcValue {
    match collect_numbers(args, ctx) {
        Ok(nums) => CalcValue::number(nums.iter().sum()),
        Err(e) => e,
    }
}

pub fn fn_abs(args: &[FunctionArg<'_>], _ctx: &EvalContext<'_>) -> CalcValue {
    let n = match args.first() {
        Some(a) => match require_number(a) {
            Ok(v) => v,
            Err(e) => return e,
        },
        None => return CalcValue::error(XlError::Value),
    };
    CalcValue::number(n.abs())
}

pub fn fn_round(args: &[FunctionArg<'_>], _ctx: &EvalContext<'_>) -> CalcValue {
    let (n, digits) = match extract_two_numbers(args) {
        Ok(pair) => pair,
        Err(e) => return e,
    };
    CalcValue::number(round_half_away(n, digits.trunc() as i32))
}

pub fn fn_roundup(args: &[FunctionArg<'_>], _ctx: &EvalContext<'_>) -> CalcValue {
    let (n, digits) = match extract_two_numbers(args) {
        Ok(pair) => pair,
        Err(e) => return e,
    };
    let d = digits.trunc() as i32;
    let factor = 10_f64.powi(d);
    let result = if n >= 0.0 {
        (n * factor).ceil() / factor
    } else {
        (n * factor).floor() / factor
    };
    CalcValue::number(result)
}

pub fn fn_rounddown(args: &[FunctionArg<'_>], _ctx: &EvalContext<'_>) -> CalcValue {
    let (n, digits) = match extract_two_numbers(args) {
        Ok(pair) => pair,
        Err(e) => return e,
    };
    let d = digits.trunc() as i32;
    let factor = 10_f64.powi(d);
    let result = if n >= 0.0 {
        (n * factor).floor() / factor
    } else {
        (n * factor).ceil() / factor
    };
    CalcValue::number(result)
}

pub fn fn_int(args: &[FunctionArg<'_>], _ctx: &EvalContext<'_>) -> CalcValue {
    let n = match args.first() {
        Some(a) => match require_number(a) {
            Ok(v) => v,
            Err(e) => return e,
        },
        None => return CalcValue::error(XlError::Value),
    };
    CalcValue::number(n.floor())
}

pub fn fn_mod(args: &[FunctionArg<'_>], _ctx: &EvalContext<'_>) -> CalcValue {
    let (n, d) = match extract_two_numbers(args) {
        Ok(pair) => pair,
        Err(e) => return e,
    };
    if d == 0.0 {
        return CalcValue::error(XlError::Div0);
    }
    let result = n - d * (n / d).floor();
    CalcValue::number(result)
}

pub fn fn_power(args: &[FunctionArg<'_>], _ctx: &EvalContext<'_>) -> CalcValue {
    let (base, exp) = match extract_two_numbers(args) {
        Ok(pair) => pair,
        Err(e) => return e,
    };
    if base == 0.0 && exp == 0.0 {
        return CalcValue::error(XlError::Num);
    }
    if base == 0.0 && exp < 0.0 {
        return CalcValue::error(XlError::Div0);
    }
    if base < 0.0 && exp.fract() != 0.0 {
        return CalcValue::error(XlError::Num);
    }
    let result = base.powf(exp);
    if result.is_infinite() || result.is_nan() {
        return CalcValue::error(XlError::Num);
    }
    CalcValue::number(result)
}

pub fn fn_sqrt(args: &[FunctionArg<'_>], _ctx: &EvalContext<'_>) -> CalcValue {
    let n = match args.first() {
        Some(a) => match require_number(a) {
            Ok(v) => v,
            Err(e) => return e,
        },
        None => return CalcValue::error(XlError::Value),
    };
    if n < 0.0 {
        return CalcValue::error(XlError::Num);
    }
    CalcValue::number(n.sqrt())
}

pub fn fn_pi(_args: &[FunctionArg<'_>], _ctx: &EvalContext<'_>) -> CalcValue {
    CalcValue::number(std::f64::consts::PI)
}

pub fn fn_sign(args: &[FunctionArg<'_>], _ctx: &EvalContext<'_>) -> CalcValue {
    let n = match args.first() {
        Some(a) => match require_number(a) {
            Ok(v) => v,
            Err(e) => return e,
        },
        None => return CalcValue::error(XlError::Value),
    };
    let result = if n > 0.0 {
        1.0
    } else if n < 0.0 {
        -1.0
    } else {
        0.0
    };
    CalcValue::number(result)
}

pub fn fn_log(args: &[FunctionArg<'_>], _ctx: &EvalContext<'_>) -> CalcValue {
    let n = match args.first() {
        Some(a) => match require_number(a) {
            Ok(v) => v,
            Err(e) => return e,
        },
        None => return CalcValue::error(XlError::Value),
    };
    let base = match args.get(1) {
        Some(a) => match require_number(a) {
            Ok(v) => v,
            Err(e) => return e,
        },
        None => 10.0,
    };
    if n <= 0.0 || base <= 0.0 {
        return CalcValue::error(XlError::Num);
    }
    if (base - 1.0).abs() < 1e-10 {
        return CalcValue::error(XlError::Num);
    }
    CalcValue::number(n.ln() / base.ln())
}

pub fn fn_ln(args: &[FunctionArg<'_>], _ctx: &EvalContext<'_>) -> CalcValue {
    let n = match args.first() {
        Some(a) => match require_number(a) {
            Ok(v) => v,
            Err(e) => return e,
        },
        None => return CalcValue::error(XlError::Value),
    };
    if n <= 0.0 {
        return CalcValue::error(XlError::Num);
    }
    CalcValue::number(n.ln())
}

// ============================================================================
// HELPERS
// ============================================================================

fn extract_two_numbers(args: &[FunctionArg<'_>]) -> Result<(f64, f64), CalcValue> {
    let a = match args.first() {
        Some(a) => require_number(a)?,
        None => return Err(CalcValue::error(XlError::Value)),
    };
    let b = match args.get(1) {
        Some(a) => require_number(a)?,
        None => return Err(CalcValue::error(XlError::Value)),
    };
    Ok((a, b))
}

fn round_half_away(value: f64, digits: i32) -> f64 {
    if digits < 0 {
        let coef = 10_f64.powi(-digits);
        let shifted = value / coef;
        let shifted = if shifted.abs() >= 0.0 {
            (shifted.abs() + 0.5).floor().copysign(shifted)
        } else {
            shifted
        };
        if shifted == 0.0 {
            return 0.0;
        }
        return shifted * coef;
    }
    let factor = 10_f64.powi(digits);
    let scaled = value * factor;
    let rounded = (scaled.abs() + 0.5).floor().copysign(scaled);
    let result = rounded / factor;
    if digits >= 15 {
        value
    } else {
        result
    }
}

// ============================================================================
// REGISTRATION
// ============================================================================

pub fn register(registry: &mut FunctionRegistry) {
    // Original 13 functions
    registry.register(FunctionDef {
        name: "SUM",
        min_args: 1,
        max_args: usize::MAX,
        param_kinds: &[ParamKind::Range],
        func: fn_sum,
    });
    registry.register(FunctionDef {
        name: "ABS",
        min_args: 1,
        max_args: 1,
        param_kinds: &[ParamKind::Scalar],
        func: fn_abs,
    });
    registry.register(FunctionDef {
        name: "ROUND",
        min_args: 2,
        max_args: 2,
        param_kinds: &[ParamKind::Scalar, ParamKind::Scalar],
        func: fn_round,
    });
    registry.register(FunctionDef {
        name: "ROUNDUP",
        min_args: 2,
        max_args: 2,
        param_kinds: &[ParamKind::Scalar, ParamKind::Scalar],
        func: fn_roundup,
    });
    registry.register(FunctionDef {
        name: "ROUNDDOWN",
        min_args: 2,
        max_args: 2,
        param_kinds: &[ParamKind::Scalar, ParamKind::Scalar],
        func: fn_rounddown,
    });
    registry.register(FunctionDef {
        name: "INT",
        min_args: 1,
        max_args: 1,
        param_kinds: &[ParamKind::Scalar],
        func: fn_int,
    });
    registry.register(FunctionDef {
        name: "MOD",
        min_args: 2,
        max_args: 2,
        param_kinds: &[ParamKind::Scalar, ParamKind::Scalar],
        func: fn_mod,
    });
    registry.register(FunctionDef {
        name: "POWER",
        min_args: 2,
        max_args: 2,
        param_kinds: &[ParamKind::Scalar, ParamKind::Scalar],
        func: fn_power,
    });
    registry.register(FunctionDef {
        name: "SQRT",
        min_args: 1,
        max_args: 1,
        param_kinds: &[ParamKind::Scalar],
        func: fn_sqrt,
    });
    registry.register(FunctionDef {
        name: "PI",
        min_args: 0,
        max_args: 0,
        param_kinds: &[],
        func: fn_pi,
    });
    registry.register(FunctionDef {
        name: "SIGN",
        min_args: 1,
        max_args: 1,
        param_kinds: &[ParamKind::Scalar],
        func: fn_sign,
    });
    registry.register(FunctionDef {
        name: "LOG",
        min_args: 1,
        max_args: 2,
        param_kinds: &[ParamKind::Scalar, ParamKind::Scalar],
        func: fn_log,
    });
    registry.register(FunctionDef {
        name: "LN",
        min_args: 1,
        max_args: 1,
        param_kinds: &[ParamKind::Scalar],
        func: fn_ln,
    });

    // Trigonometric (22 functions)
    registry.register(FunctionDef {
        name: "ACOS",
        min_args: 1,
        max_args: 1,
        param_kinds: &[ParamKind::Scalar],
        func: fn_acos,
    });
    registry.register(FunctionDef {
        name: "ACOSH",
        min_args: 1,
        max_args: 1,
        param_kinds: &[ParamKind::Scalar],
        func: fn_acosh,
    });
    registry.register(FunctionDef {
        name: "ACOT",
        min_args: 1,
        max_args: 1,
        param_kinds: &[ParamKind::Scalar],
        func: fn_acot,
    });
    registry.register(FunctionDef {
        name: "ACOTH",
        min_args: 1,
        max_args: 1,
        param_kinds: &[ParamKind::Scalar],
        func: fn_acoth,
    });
    registry.register(FunctionDef {
        name: "ASIN",
        min_args: 1,
        max_args: 1,
        param_kinds: &[ParamKind::Scalar],
        func: fn_asin,
    });
    registry.register(FunctionDef {
        name: "ASINH",
        min_args: 1,
        max_args: 1,
        param_kinds: &[ParamKind::Scalar],
        func: fn_asinh,
    });
    registry.register(FunctionDef {
        name: "ATAN",
        min_args: 1,
        max_args: 1,
        param_kinds: &[ParamKind::Scalar],
        func: fn_atan,
    });
    registry.register(FunctionDef {
        name: "ATAN2",
        min_args: 2,
        max_args: 2,
        param_kinds: &[ParamKind::Scalar, ParamKind::Scalar],
        func: fn_atan2,
    });
    registry.register(FunctionDef {
        name: "ATANH",
        min_args: 1,
        max_args: 1,
        param_kinds: &[ParamKind::Scalar],
        func: fn_atanh,
    });
    registry.register(FunctionDef {
        name: "COS",
        min_args: 1,
        max_args: 1,
        param_kinds: &[ParamKind::Scalar],
        func: fn_cos,
    });
    registry.register(FunctionDef {
        name: "COSH",
        min_args: 1,
        max_args: 1,
        param_kinds: &[ParamKind::Scalar],
        func: fn_cosh,
    });
    registry.register(FunctionDef {
        name: "COT",
        min_args: 1,
        max_args: 1,
        param_kinds: &[ParamKind::Scalar],
        func: fn_cot,
    });
    registry.register(FunctionDef {
        name: "COTH",
        min_args: 1,
        max_args: 1,
        param_kinds: &[ParamKind::Scalar],
        func: fn_coth,
    });
    registry.register(FunctionDef {
        name: "CSC",
        min_args: 1,
        max_args: 1,
        param_kinds: &[ParamKind::Scalar],
        func: fn_csc,
    });
    registry.register(FunctionDef {
        name: "CSCH",
        min_args: 1,
        max_args: 1,
        param_kinds: &[ParamKind::Scalar],
        func: fn_csch,
    });
    registry.register(FunctionDef {
        name: "SIN",
        min_args: 1,
        max_args: 1,
        param_kinds: &[ParamKind::Scalar],
        func: fn_sin,
    });
    registry.register(FunctionDef {
        name: "SINH",
        min_args: 1,
        max_args: 1,
        param_kinds: &[ParamKind::Scalar],
        func: fn_sinh,
    });
    registry.register(FunctionDef {
        name: "TAN",
        min_args: 1,
        max_args: 1,
        param_kinds: &[ParamKind::Scalar],
        func: fn_tan,
    });
    registry.register(FunctionDef {
        name: "TANH",
        min_args: 1,
        max_args: 1,
        param_kinds: &[ParamKind::Scalar],
        func: fn_tanh,
    });
    registry.register(FunctionDef {
        name: "SEC",
        min_args: 1,
        max_args: 1,
        param_kinds: &[ParamKind::Scalar],
        func: fn_sec,
    });
    registry.register(FunctionDef {
        name: "SECH",
        min_args: 1,
        max_args: 1,
        param_kinds: &[ParamKind::Scalar],
        func: fn_sech,
    });
    registry.register(FunctionDef {
        name: "DEGREES",
        min_args: 1,
        max_args: 1,
        param_kinds: &[ParamKind::Scalar],
        func: fn_degrees,
    });
    registry.register(FunctionDef {
        name: "RADIANS",
        min_args: 1,
        max_args: 1,
        param_kinds: &[ParamKind::Scalar],
        func: fn_radians,
    });

    // Rounding & Special (10 functions)
    registry.register(FunctionDef {
        name: "CEILING",
        min_args: 2,
        max_args: 2,
        param_kinds: &[ParamKind::Scalar, ParamKind::Scalar],
        func: fn_ceiling,
    });
    registry.register(FunctionDef {
        name: "CEILING.MATH",
        min_args: 1,
        max_args: 3,
        param_kinds: &[ParamKind::Scalar, ParamKind::Scalar, ParamKind::Scalar],
        func: fn_ceiling_math,
    });
    registry.register(FunctionDef {
        name: "FLOOR",
        min_args: 2,
        max_args: 2,
        param_kinds: &[ParamKind::Scalar, ParamKind::Scalar],
        func: fn_floor,
    });
    registry.register(FunctionDef {
        name: "FLOOR.MATH",
        min_args: 1,
        max_args: 3,
        param_kinds: &[ParamKind::Scalar, ParamKind::Scalar, ParamKind::Scalar],
        func: fn_floor_math,
    });
    registry.register(FunctionDef {
        name: "EVEN",
        min_args: 1,
        max_args: 1,
        param_kinds: &[ParamKind::Scalar],
        func: fn_even,
    });
    registry.register(FunctionDef {
        name: "ODD",
        min_args: 1,
        max_args: 1,
        param_kinds: &[ParamKind::Scalar],
        func: fn_odd,
    });
    registry.register(FunctionDef {
        name: "TRUNC",
        min_args: 1,
        max_args: 2,
        param_kinds: &[ParamKind::Scalar, ParamKind::Scalar],
        func: fn_trunc,
    });
    registry.register(FunctionDef {
        name: "MROUND",
        min_args: 2,
        max_args: 2,
        param_kinds: &[ParamKind::Scalar, ParamKind::Scalar],
        func: fn_mround,
    });
    registry.register(FunctionDef {
        name: "EXP",
        min_args: 1,
        max_args: 1,
        param_kinds: &[ParamKind::Scalar],
        func: fn_exp,
    });
    registry.register(FunctionDef {
        name: "LOG10",
        min_args: 1,
        max_args: 1,
        param_kinds: &[ParamKind::Scalar],
        func: fn_log10,
    });
    registry.register(FunctionDef {
        name: "SQRTPI",
        min_args: 1,
        max_args: 1,
        param_kinds: &[ParamKind::Scalar],
        func: fn_sqrtpi,
    });

    // Factorial & Combinatorics (7 functions)
    registry.register(FunctionDef {
        name: "FACT",
        min_args: 1,
        max_args: 1,
        param_kinds: &[ParamKind::Scalar],
        func: fn_fact,
    });
    registry.register(FunctionDef {
        name: "FACTDOUBLE",
        min_args: 1,
        max_args: 1,
        param_kinds: &[ParamKind::Scalar],
        func: fn_factdouble,
    });
    registry.register(FunctionDef {
        name: "COMBIN",
        min_args: 2,
        max_args: 2,
        param_kinds: &[ParamKind::Scalar, ParamKind::Scalar],
        func: fn_combin,
    });
    registry.register(FunctionDef {
        name: "COMBINA",
        min_args: 2,
        max_args: 2,
        param_kinds: &[ParamKind::Scalar, ParamKind::Scalar],
        func: fn_combina,
    });
    registry.register(FunctionDef {
        name: "MULTINOMIAL",
        min_args: 1,
        max_args: usize::MAX,
        param_kinds: &[ParamKind::Range],
        func: fn_multinomial,
    });
    registry.register(FunctionDef {
        name: "GCD",
        min_args: 1,
        max_args: usize::MAX,
        param_kinds: &[ParamKind::Range],
        func: fn_gcd,
    });
    registry.register(FunctionDef {
        name: "LCM",
        min_args: 1,
        max_args: usize::MAX,
        param_kinds: &[ParamKind::Range],
        func: fn_lcm,
    });

    // Base Conversion (4 functions)
    registry.register(FunctionDef {
        name: "ARABIC",
        min_args: 1,
        max_args: 1,
        param_kinds: &[ParamKind::Scalar],
        func: fn_arabic,
    });
    registry.register(FunctionDef {
        name: "BASE",
        min_args: 2,
        max_args: 3,
        param_kinds: &[ParamKind::Scalar, ParamKind::Scalar, ParamKind::Scalar],
        func: fn_base,
    });
    registry.register(FunctionDef {
        name: "DECIMAL",
        min_args: 2,
        max_args: 2,
        param_kinds: &[ParamKind::Scalar, ParamKind::Scalar],
        func: fn_decimal,
    });
    registry.register(FunctionDef {
        name: "ROMAN",
        min_args: 1,
        max_args: 2,
        param_kinds: &[ParamKind::Scalar, ParamKind::Scalar],
        func: fn_roman,
    });

    // Product & Quotient (2 functions)
    registry.register(FunctionDef {
        name: "PRODUCT",
        min_args: 1,
        max_args: usize::MAX,
        param_kinds: &[ParamKind::Range],
        func: fn_product,
    });
    registry.register(FunctionDef {
        name: "QUOTIENT",
        min_args: 2,
        max_args: 2,
        param_kinds: &[ParamKind::Scalar, ParamKind::Scalar],
        func: fn_quotient,
    });

    // Random (2 functions)
    registry.register(FunctionDef {
        name: "RAND",
        min_args: 0,
        max_args: 0,
        param_kinds: &[],
        func: fn_rand,
    });
    registry.register(FunctionDef {
        name: "RANDBETWEEN",
        min_args: 2,
        max_args: 2,
        param_kinds: &[ParamKind::Scalar, ParamKind::Scalar],
        func: fn_randbetween,
    });

    // Aggregate (7 functions)
    registry.register(FunctionDef {
        name: "SUMPRODUCT",
        min_args: 1,
        max_args: 30,
        param_kinds: &[ParamKind::Range],
        func: fn_sumproduct,
    });
    registry.register(FunctionDef {
        name: "SUMSQ",
        min_args: 1,
        max_args: usize::MAX,
        param_kinds: &[ParamKind::Range],
        func: fn_sumsq,
    });
    registry.register(FunctionDef {
        name: "SUMX2MY2",
        min_args: 2,
        max_args: 2,
        param_kinds: &[ParamKind::Range, ParamKind::Range],
        func: fn_sumx2my2,
    });
    registry.register(FunctionDef {
        name: "SUMX2PY2",
        min_args: 2,
        max_args: 2,
        param_kinds: &[ParamKind::Range, ParamKind::Range],
        func: fn_sumx2py2,
    });
    registry.register(FunctionDef {
        name: "SUMXMY2",
        min_args: 2,
        max_args: 2,
        param_kinds: &[ParamKind::Range, ParamKind::Range],
        func: fn_sumxmy2,
    });
    registry.register(FunctionDef {
        name: "SERIESSUM",
        min_args: 4,
        max_args: 4,
        param_kinds: &[
            ParamKind::Scalar,
            ParamKind::Scalar,
            ParamKind::Scalar,
            ParamKind::Range,
        ],
        func: fn_seriessum,
    });
    registry.register(FunctionDef {
        name: "SUBTOTAL",
        min_args: 2,
        max_args: usize::MAX,
        param_kinds: &[ParamKind::Scalar, ParamKind::Range],
        func: fn_subtotal,
    });

    // Conditional sums (placeholders - require criteria engine)
    registry.register(FunctionDef {
        name: "SUMIF",
        min_args: 2,
        max_args: 3,
        param_kinds: &[ParamKind::Range, ParamKind::Scalar, ParamKind::Range],
        func: fn_sumif,
    });
    registry.register(FunctionDef {
        name: "SUMIFS",
        min_args: 3,
        max_args: usize::MAX,
        param_kinds: &[ParamKind::Range],
        func: fn_sumifs,
    });

    // Matrix operations (placeholders)
    registry.register(FunctionDef {
        name: "MDETERM",
        min_args: 1,
        max_args: 1,
        param_kinds: &[ParamKind::Range],
        func: fn_mdeterm,
    });
    registry.register(FunctionDef {
        name: "MINVERSE",
        min_args: 1,
        max_args: 1,
        param_kinds: &[ParamKind::Range],
        func: fn_minverse,
    });
    registry.register(FunctionDef {
        name: "MMULT",
        min_args: 2,
        max_args: 2,
        param_kinds: &[ParamKind::Range, ParamKind::Range],
        func: fn_mmult,
    });
}
