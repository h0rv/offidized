//! Extended mathematical and trigonometric functions.
//!
//! This module contains additional math functions beyond the basic set:
//! Trigonometric, hyperbolic, rounding, factorials, combinatorics, base conversion,
//! matrix operations, conditional sums, and more.

use super::super::{iter_range_values, require_number, FunctionArg};
use crate::context::EvalContext;
use crate::value::{CalcValue, ScalarValue, XlError};

const MAX_DOUBLE_INT: f64 = 9007199254740991.0; // 2^53 - 1

// ============================================================================
// TRIGONOMETRIC FUNCTIONS
// ============================================================================

pub fn fn_acos(args: &[FunctionArg<'_>], _ctx: &EvalContext<'_>) -> CalcValue {
    let n = match require_number(&args[0]) {
        Ok(v) => v,
        Err(e) => return e,
    };
    if n.abs() > 1.0 {
        return CalcValue::error(XlError::Num);
    }
    CalcValue::number(n.acos())
}

pub fn fn_acosh(args: &[FunctionArg<'_>], _ctx: &EvalContext<'_>) -> CalcValue {
    let n = match require_number(&args[0]) {
        Ok(v) => v,
        Err(e) => return e,
    };
    if n < 1.0 {
        return CalcValue::error(XlError::Num);
    }
    CalcValue::number((n + (n * n - 1.0).sqrt()).ln())
}

pub fn fn_acot(args: &[FunctionArg<'_>], _ctx: &EvalContext<'_>) -> CalcValue {
    let angle = match require_number(&args[0]) {
        Ok(v) => v,
        Err(e) => return e,
    };
    if angle == 0.0 {
        return CalcValue::number(std::f64::consts::PI / 2.0);
    }
    let mut acot = (1.0 / angle).atan();
    while acot < 0.0 {
        acot += std::f64::consts::PI;
    }
    CalcValue::number(acot)
}

pub fn fn_acoth(args: &[FunctionArg<'_>], _ctx: &EvalContext<'_>) -> CalcValue {
    let angle = match require_number(&args[0]) {
        Ok(v) => v,
        Err(e) => return e,
    };
    if angle.abs() < 1.0 {
        return CalcValue::error(XlError::Num);
    }
    CalcValue::number(0.5 * ((angle + 1.0) / (angle - 1.0)).ln())
}

pub fn fn_asin(args: &[FunctionArg<'_>], _ctx: &EvalContext<'_>) -> CalcValue {
    let n = match require_number(&args[0]) {
        Ok(v) => v,
        Err(e) => return e,
    };
    if n.abs() > 1.0 {
        return CalcValue::error(XlError::Num);
    }
    CalcValue::number(n.asin())
}

pub fn fn_asinh(args: &[FunctionArg<'_>], _ctx: &EvalContext<'_>) -> CalcValue {
    let n = match require_number(&args[0]) {
        Ok(v) => v,
        Err(e) => return e,
    };
    CalcValue::number((n + (n * n + 1.0).sqrt()).ln())
}

pub fn fn_atan(args: &[FunctionArg<'_>], _ctx: &EvalContext<'_>) -> CalcValue {
    let n = match require_number(&args[0]) {
        Ok(v) => v,
        Err(e) => return e,
    };
    CalcValue::number(n.atan())
}

pub fn fn_atan2(args: &[FunctionArg<'_>], _ctx: &EvalContext<'_>) -> CalcValue {
    let x = match require_number(&args[0]) {
        Ok(v) => v,
        Err(e) => return e,
    };
    let y = match require_number(&args[1]) {
        Ok(v) => v,
        Err(e) => return e,
    };
    if x == 0.0 && y == 0.0 {
        return CalcValue::error(XlError::Div0);
    }
    CalcValue::number(y.atan2(x))
}

pub fn fn_atanh(args: &[FunctionArg<'_>], _ctx: &EvalContext<'_>) -> CalcValue {
    let n = match require_number(&args[0]) {
        Ok(v) => v,
        Err(e) => return e,
    };
    if n.abs() >= 1.0 {
        return CalcValue::error(XlError::Num);
    }
    CalcValue::number(((1.0 + n) / (1.0 - n)).ln() / 2.0)
}

pub fn fn_cos(args: &[FunctionArg<'_>], _ctx: &EvalContext<'_>) -> CalcValue {
    let n = match require_number(&args[0]) {
        Ok(v) => v,
        Err(e) => return e,
    };
    CalcValue::number(n.cos())
}

pub fn fn_cosh(args: &[FunctionArg<'_>], _ctx: &EvalContext<'_>) -> CalcValue {
    let n = match require_number(&args[0]) {
        Ok(v) => v,
        Err(e) => return e,
    };
    let result = n.cosh();
    if result.is_infinite() {
        return CalcValue::error(XlError::Num);
    }
    CalcValue::number(result)
}

pub fn fn_cot(args: &[FunctionArg<'_>], _ctx: &EvalContext<'_>) -> CalcValue {
    let angle = match require_number(&args[0]) {
        Ok(v) => v,
        Err(e) => return e,
    };
    let tan = angle.tan();
    if tan == 0.0 {
        return CalcValue::error(XlError::Div0);
    }
    CalcValue::number(1.0 / tan)
}

pub fn fn_coth(args: &[FunctionArg<'_>], _ctx: &EvalContext<'_>) -> CalcValue {
    let angle = match require_number(&args[0]) {
        Ok(v) => v,
        Err(e) => return e,
    };
    if angle == 0.0 {
        return CalcValue::error(XlError::Div0);
    }
    CalcValue::number(1.0 / angle.tanh())
}

pub fn fn_csc(args: &[FunctionArg<'_>], _ctx: &EvalContext<'_>) -> CalcValue {
    let angle = match require_number(&args[0]) {
        Ok(v) => v,
        Err(e) => return e,
    };
    if angle == 0.0 {
        return CalcValue::error(XlError::Div0);
    }
    CalcValue::number(1.0 / angle.sin())
}

pub fn fn_csch(args: &[FunctionArg<'_>], _ctx: &EvalContext<'_>) -> CalcValue {
    let angle = match require_number(&args[0]) {
        Ok(v) => v,
        Err(e) => return e,
    };
    if angle == 0.0 {
        return CalcValue::error(XlError::Div0);
    }
    CalcValue::number(1.0 / angle.sinh())
}

pub fn fn_sin(args: &[FunctionArg<'_>], _ctx: &EvalContext<'_>) -> CalcValue {
    let n = match require_number(&args[0]) {
        Ok(v) => v,
        Err(e) => return e,
    };
    CalcValue::number(n.sin())
}

pub fn fn_sinh(args: &[FunctionArg<'_>], _ctx: &EvalContext<'_>) -> CalcValue {
    let n = match require_number(&args[0]) {
        Ok(v) => v,
        Err(e) => return e,
    };
    let result = n.sinh();
    if result.is_infinite() {
        return CalcValue::error(XlError::Num);
    }
    CalcValue::number(result)
}

pub fn fn_tan(args: &[FunctionArg<'_>], _ctx: &EvalContext<'_>) -> CalcValue {
    let radians = match require_number(&args[0]) {
        Ok(v) => v,
        Err(e) => return e,
    };
    if radians.abs() >= 134217728.0 {
        return CalcValue::error(XlError::Num);
    }
    CalcValue::number(radians.tan())
}

pub fn fn_tanh(args: &[FunctionArg<'_>], _ctx: &EvalContext<'_>) -> CalcValue {
    let n = match require_number(&args[0]) {
        Ok(v) => v,
        Err(e) => return e,
    };
    CalcValue::number(n.tanh())
}

pub fn fn_sec(args: &[FunctionArg<'_>], _ctx: &EvalContext<'_>) -> CalcValue {
    let angle = match require_number(&args[0]) {
        Ok(v) => v,
        Err(e) => return e,
    };
    CalcValue::number(1.0 / angle.cos())
}

pub fn fn_sech(args: &[FunctionArg<'_>], _ctx: &EvalContext<'_>) -> CalcValue {
    let angle = match require_number(&args[0]) {
        Ok(v) => v,
        Err(e) => return e,
    };
    CalcValue::number(1.0 / angle.cosh())
}

pub fn fn_degrees(args: &[FunctionArg<'_>], _ctx: &EvalContext<'_>) -> CalcValue {
    let n = match require_number(&args[0]) {
        Ok(v) => v,
        Err(e) => return e,
    };
    CalcValue::number(n * 180.0 / std::f64::consts::PI)
}

pub fn fn_radians(args: &[FunctionArg<'_>], _ctx: &EvalContext<'_>) -> CalcValue {
    let angle = match require_number(&args[0]) {
        Ok(v) => v,
        Err(e) => return e,
    };
    CalcValue::number(angle * std::f64::consts::PI / 180.0)
}

// ============================================================================
// ROUNDING & SPECIAL FUNCTIONS
// ============================================================================

pub fn fn_ceiling(args: &[FunctionArg<'_>], _ctx: &EvalContext<'_>) -> CalcValue {
    let number = match require_number(&args[0]) {
        Ok(v) => v,
        Err(e) => return e,
    };
    let significance = match require_number(&args[1]) {
        Ok(v) => v,
        Err(e) => return e,
    };

    if significance == 0.0 {
        return CalcValue::number(0.0);
    }
    if significance < 0.0 && number > 0.0 {
        return CalcValue::error(XlError::Num);
    }

    if number < 0.0 {
        return CalcValue::number(-(-number / -significance).ceil() * -significance);
    }
    CalcValue::number((number / significance).ceil() * significance)
}

pub fn fn_ceiling_math(args: &[FunctionArg<'_>], _ctx: &EvalContext<'_>) -> CalcValue {
    let number = match require_number(&args[0]) {
        Ok(v) => v,
        Err(e) => return e,
    };
    let significance = if args.len() > 1 {
        match require_number(&args[1]) {
            Ok(v) => v,
            Err(e) => return e,
        }
    } else {
        1.0
    };
    let mode = if args.len() > 2 {
        match require_number(&args[2]) {
            Ok(v) => v,
            Err(e) => return e,
        }
    } else {
        0.0
    };

    if significance == 0.0 {
        return CalcValue::number(0.0);
    }
    let sig = significance.abs();

    if number < 0.0 && mode != 0.0 {
        return CalcValue::number((number / sig).floor() * sig);
    }
    CalcValue::number((number / sig).ceil() * sig)
}

pub fn fn_floor(args: &[FunctionArg<'_>], _ctx: &EvalContext<'_>) -> CalcValue {
    let number = match require_number(&args[0]) {
        Ok(v) => v,
        Err(e) => return e,
    };
    let significance = match require_number(&args[1]) {
        Ok(v) => v,
        Err(e) => return e,
    };

    if number == 0.0 {
        return CalcValue::number(0.0);
    }
    if number > 0.0 && significance < 0.0 {
        return CalcValue::error(XlError::Num);
    }
    if significance == 0.0 {
        return CalcValue::error(XlError::Div0);
    }

    if significance < 0.0 {
        return CalcValue::number(-(-number / -significance).floor() * -significance);
    }
    CalcValue::number((number / significance).floor() * significance)
}

pub fn fn_floor_math(args: &[FunctionArg<'_>], _ctx: &EvalContext<'_>) -> CalcValue {
    let number = match require_number(&args[0]) {
        Ok(v) => v,
        Err(e) => return e,
    };
    let significance = if args.len() > 1 {
        match require_number(&args[1]) {
            Ok(v) => v,
            Err(e) => return e,
        }
    } else {
        1.0
    };
    let mode = if args.len() > 2 {
        match require_number(&args[2]) {
            Ok(v) => v,
            Err(e) => return e,
        }
    } else {
        0.0
    };

    if significance == 0.0 {
        return CalcValue::number(0.0);
    }
    let sig = significance.abs();

    if number >= 0.0 {
        return CalcValue::number((number / sig).floor() * sig);
    }

    if mode == 0.0 {
        CalcValue::number((number / sig).floor() * sig)
    } else {
        CalcValue::number((number / sig).trunc() * sig)
    }
}

pub fn fn_even(args: &[FunctionArg<'_>], _ctx: &EvalContext<'_>) -> CalcValue {
    let number = match require_number(&args[0]) {
        Ok(v) => v,
        Err(e) => return e,
    };
    let num = number.ceil();
    let add_value = if num >= 0.0 { 1.0 } else { -1.0 };
    let is_even = num % 2.0 == 0.0;
    CalcValue::number(if is_even { num } else { num + add_value })
}

pub fn fn_odd(args: &[FunctionArg<'_>], _ctx: &EvalContext<'_>) -> CalcValue {
    let number = match require_number(&args[0]) {
        Ok(v) => v,
        Err(e) => return e,
    };
    let num = number.ceil();
    let add_value = if num >= 0.0 { 1.0 } else { -1.0 };
    let is_odd = num % 2.0 != 0.0;
    CalcValue::number(if is_odd { num } else { num + add_value })
}

pub fn fn_trunc(args: &[FunctionArg<'_>], _ctx: &EvalContext<'_>) -> CalcValue {
    let number = match require_number(&args[0]) {
        Ok(v) => v,
        Err(e) => return e,
    };
    let digits = if args.len() > 1 {
        match require_number(&args[1]) {
            Ok(v) => v,
            Err(e) => return e,
        }
    } else {
        0.0
    };

    let scaling = 10_f64.powf(digits);
    CalcValue::number((number * scaling).trunc() / scaling)
}

pub fn fn_mround(args: &[FunctionArg<'_>], _ctx: &EvalContext<'_>) -> CalcValue {
    let number = match require_number(&args[0]) {
        Ok(v) => v,
        Err(e) => return e,
    };
    let multiple = match require_number(&args[1]) {
        Ok(v) => v,
        Err(e) => return e,
    };

    if multiple == 0.0 {
        return CalcValue::number(0.0);
    }
    if number == 0.0 {
        return CalcValue::number(0.0);
    }
    if number.signum() != multiple.signum() {
        return CalcValue::error(XlError::Num);
    }

    let result = (number / multiple).round() * multiple;
    CalcValue::number(result)
}

pub fn fn_exp(args: &[FunctionArg<'_>], _ctx: &EvalContext<'_>) -> CalcValue {
    let number = match require_number(&args[0]) {
        Ok(v) => v,
        Err(e) => return e,
    };
    let exp = number.exp();
    if exp.is_infinite() {
        return CalcValue::error(XlError::Num);
    }
    CalcValue::number(exp)
}

pub fn fn_log10(args: &[FunctionArg<'_>], _ctx: &EvalContext<'_>) -> CalcValue {
    let x = match require_number(&args[0]) {
        Ok(v) => v,
        Err(e) => return e,
    };
    if x <= 0.0 {
        return CalcValue::error(XlError::Num);
    }
    CalcValue::number(x.log10())
}

pub fn fn_sqrtpi(args: &[FunctionArg<'_>], _ctx: &EvalContext<'_>) -> CalcValue {
    let number = match require_number(&args[0]) {
        Ok(v) => v,
        Err(e) => return e,
    };
    if number < 0.0 {
        return CalcValue::error(XlError::Num);
    }
    CalcValue::number((std::f64::consts::PI * number).sqrt())
}

// ============================================================================
// FACTORIAL & COMBINATORICS
// ============================================================================

fn factorial(n: f64) -> f64 {
    let mut n = n.trunc();
    let mut result = 1.0;
    while n > 1.0 {
        result *= n;
        n -= 1.0;
        if result.is_infinite() {
            return result;
        }
    }
    result
}

pub fn fn_fact(args: &[FunctionArg<'_>], _ctx: &EvalContext<'_>) -> CalcValue {
    let n = match require_number(&args[0]) {
        Ok(v) => v,
        Err(e) => return e,
    };
    if !(0.0..171.0).contains(&n) {
        return CalcValue::error(XlError::Num);
    }
    CalcValue::number(factorial(n.floor()))
}

pub fn fn_factdouble(args: &[FunctionArg<'_>], _ctx: &EvalContext<'_>) -> CalcValue {
    let n = match require_number(&args[0]) {
        Ok(v) => v,
        Err(e) => return e,
    };
    let num = n.floor();
    if num < -1.0 {
        return CalcValue::error(XlError::Num);
    }

    let mut fact = 1.0;
    if num > 1.0 {
        let start = if num % 2.0 == 0.0 { 2.0 } else { 1.0 };
        let mut i = start;
        while i <= num {
            fact *= i;
            i += 2.0;
            if fact.is_infinite() {
                return CalcValue::error(XlError::Num);
            }
        }
    }
    CalcValue::number(fact)
}

fn combin(n: f64, k: f64) -> f64 {
    if k == 0.0 {
        return 1.0;
    }
    let mut result = 1.0;
    let mut n = n;
    for i in 1..=(k as i32) {
        result *= n;
        result /= i as f64;
        n -= 1.0;
    }
    result
}

pub fn fn_combin(args: &[FunctionArg<'_>], _ctx: &EvalContext<'_>) -> CalcValue {
    let number = match require_number(&args[0]) {
        Ok(v) => v,
        Err(e) => return e,
    };
    let number_chosen = match require_number(&args[1]) {
        Ok(v) => v,
        Err(e) => return e,
    };

    if number < 0.0 || number_chosen < 0.0 {
        return CalcValue::error(XlError::Num);
    }

    let n = number.floor();
    let k = number_chosen.floor();

    if n >= i32::MAX as f64 || k >= i32::MAX as f64 {
        return CalcValue::error(XlError::Num);
    }
    if n < k {
        return CalcValue::error(XlError::Num);
    }

    let combinations = combin(n, k);
    if combinations.is_infinite() || combinations.is_nan() {
        return CalcValue::error(XlError::Num);
    }
    CalcValue::number(combinations)
}

pub fn fn_combina(args: &[FunctionArg<'_>], _ctx: &EvalContext<'_>) -> CalcValue {
    let number = match require_number(&args[0]) {
        Ok(v) => v,
        Err(e) => return e,
    };
    let chosen = match require_number(&args[1]) {
        Ok(v) => v,
        Err(e) => return e,
    };

    let number = number.trunc();
    let chosen = chosen.trunc();

    if number < 0.0 || chosen < 0.0 {
        return CalcValue::error(XlError::Num);
    }

    let n = number + chosen - 1.0;
    if n > i32::MAX as f64 {
        return CalcValue::error(XlError::Num);
    }

    let k = number - 1.0;
    if chosen == 0.0 || k == 0.0 {
        return CalcValue::number(1.0);
    }
    CalcValue::number(combin(n, k))
}

pub fn fn_multinomial(args: &[FunctionArg<'_>], _ctx: &EvalContext<'_>) -> CalcValue {
    let mut numbers_sum = 0.0;
    let mut denominator = 1.0;

    for arg in args {
        match arg {
            FunctionArg::Value(v) => {
                let scalar = v.as_scalar();
                if let ScalarValue::Error(e) = scalar {
                    return CalcValue::error(*e);
                }
                if let ScalarValue::Bool(_) = scalar {
                    return CalcValue::error(XlError::Value);
                }
                if let ScalarValue::Number(n) = scalar.to_number() {
                    if n < 0.0 {
                        return CalcValue::error(XlError::Num);
                    }
                    let num = n.trunc();
                    numbers_sum += num;
                    denominator *= factorial(num);
                    if denominator.is_infinite() {
                        return CalcValue::error(XlError::Num);
                    }
                }
            }
            FunctionArg::Range { range, ctx: rctx } => {
                for val in iter_range_values(range, rctx) {
                    if let ScalarValue::Error(e) = val {
                        return CalcValue::error(e);
                    }
                    if let ScalarValue::Bool(_) = val {
                        return CalcValue::error(XlError::Value);
                    }
                    if let ScalarValue::Number(n) = val.to_number() {
                        if n < 0.0 {
                            return CalcValue::error(XlError::Num);
                        }
                        let num = n.trunc();
                        numbers_sum += num;
                        denominator *= factorial(num);
                        if denominator.is_infinite() {
                            return CalcValue::error(XlError::Num);
                        }
                    }
                }
            }
        }
    }

    let numerator = factorial(numbers_sum);
    if numerator.is_infinite() {
        return CalcValue::error(XlError::Num);
    }
    CalcValue::number(numerator / denominator)
}

// ============================================================================
// GCD & LCM
// ============================================================================

fn gcd_two(a: f64, b: f64) -> f64 {
    let mut a = a.trunc();
    let mut b = b.trunc();
    while b != 0.0 {
        let temp = b;
        b = a % b;
        a = temp;
    }
    a
}

pub fn fn_gcd(args: &[FunctionArg<'_>], _ctx: &EvalContext<'_>) -> CalcValue {
    let mut result = 0.0;

    for arg in args {
        match arg {
            FunctionArg::Value(v) => {
                let scalar = v.as_scalar();
                if let ScalarValue::Error(e) = scalar {
                    return CalcValue::error(*e);
                }
                if let ScalarValue::Bool(_) = scalar {
                    return CalcValue::error(XlError::Value);
                }
                if let ScalarValue::Number(n) = scalar.to_number() {
                    if !(0.0..=MAX_DOUBLE_INT).contains(&n) {
                        return CalcValue::error(XlError::Num);
                    }
                    result = gcd_two(n, result);
                }
            }
            FunctionArg::Range { range, ctx: rctx } => {
                for val in iter_range_values(range, rctx) {
                    if let ScalarValue::Error(e) = val {
                        return CalcValue::error(e);
                    }
                    if let ScalarValue::Bool(_) = val {
                        return CalcValue::error(XlError::Value);
                    }
                    if let ScalarValue::Number(n) = val.to_number() {
                        if !(0.0..=MAX_DOUBLE_INT).contains(&n) {
                            return CalcValue::error(XlError::Num);
                        }
                        result = gcd_two(n, result);
                    }
                }
            }
        }
    }

    CalcValue::number(result)
}

fn lcm_two(a: f64, b: f64) -> f64 {
    if a == 0.0 || b == 0.0 {
        return 0.0;
    }
    a * (b / gcd_two(a, b))
}

pub fn fn_lcm(args: &[FunctionArg<'_>], _ctx: &EvalContext<'_>) -> CalcValue {
    let mut result = 1.0;

    for arg in args {
        match arg {
            FunctionArg::Value(v) => {
                let scalar = v.as_scalar();
                if let ScalarValue::Error(e) = scalar {
                    return CalcValue::error(*e);
                }
                if let ScalarValue::Bool(_) = scalar {
                    return CalcValue::error(XlError::Value);
                }
                if let ScalarValue::Number(n) = scalar.to_number() {
                    if !(0.0..=MAX_DOUBLE_INT).contains(&n) {
                        return CalcValue::error(XlError::Num);
                    }
                    result = lcm_two(result, n.trunc());
                }
            }
            FunctionArg::Range { range, ctx: rctx } => {
                for val in iter_range_values(range, rctx) {
                    if let ScalarValue::Error(e) = val {
                        return CalcValue::error(e);
                    }
                    if let ScalarValue::Bool(_) = val {
                        return CalcValue::error(XlError::Value);
                    }
                    if let ScalarValue::Number(n) = val.to_number() {
                        if !(0.0..=MAX_DOUBLE_INT).contains(&n) {
                            return CalcValue::error(XlError::Num);
                        }
                        result = lcm_two(result, n.trunc());
                    }
                }
            }
        }
    }

    CalcValue::number(result)
}

// Continued in next file part...
