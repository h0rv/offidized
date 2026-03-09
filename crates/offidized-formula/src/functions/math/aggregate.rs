//! Extended math functions part 2: Base conversion, aggregates, random, etc.

use super::super::{collect_numbers, iter_range_values, require_number, require_text, FunctionArg};
use crate::context::EvalContext;
use crate::value::{CalcValue, ScalarValue, XlError};
use std::sync::Mutex;

const MAX_DOUBLE_INT: f64 = 9007199254740991.0;

// ============================================================================
// BASE CONVERSION
// ============================================================================

static ROMAN_SYMBOL_VALUES: &[(char, i32)] = &[
    ('I', 1),
    ('V', 5),
    ('X', 10),
    ('L', 50),
    ('C', 100),
    ('D', 500),
    ('M', 1000),
];

pub fn fn_arabic(args: &[FunctionArg<'_>], _ctx: &EvalContext<'_>) -> CalcValue {
    let input = match require_text(&args[0]) {
        Ok(v) => v,
        Err(e) => return e,
    };

    if input.len() > 255 {
        return CalcValue::error(XlError::Value);
    }

    let text = input.trim();
    let minus_sign = text.starts_with('-');
    let text = if minus_sign { &text[1..] } else { text };

    let mut total = 0i32;
    let chars: Vec<char> = text.chars().collect();

    let mut i = chars.len();
    while i > 0 {
        i -= 1;
        let add_symbol = chars[i].to_ascii_uppercase();
        let add_value = match ROMAN_SYMBOL_VALUES.iter().find(|(c, _)| *c == add_symbol) {
            Some((_, v)) => *v,
            None => return CalcValue::error(XlError::Value),
        };

        total += add_value;

        while i > 0 {
            let subtract_symbol = chars[i - 1].to_ascii_uppercase();
            let subtract_value = match ROMAN_SYMBOL_VALUES
                .iter()
                .find(|(c, _)| *c == subtract_symbol)
            {
                Some((_, v)) => *v,
                None => return CalcValue::error(XlError::Value),
            };

            if subtract_value >= add_value {
                break;
            }

            total -= subtract_value;
            i -= 1;
        }
    }

    if minus_sign && total == 0 {
        return CalcValue::error(XlError::Num);
    }

    CalcValue::number(if minus_sign {
        -total as f64
    } else {
        total as f64
    })
}

pub fn fn_base(args: &[FunctionArg<'_>], _ctx: &EvalContext<'_>) -> CalcValue {
    let mut number = match require_number(&args[0]) {
        Ok(v) => v,
        Err(e) => return e,
    };
    let radix = match require_number(&args[1]) {
        Ok(v) => v,
        Err(e) => return e,
    };
    let min_length = if args.len() > 2 {
        match require_number(&args[2]) {
            Ok(v) => v,
            Err(e) => return e,
        }
    } else {
        0.0
    };

    number = number.trunc();
    let radix = radix.trunc();
    let min_length = min_length.trunc();

    if !(0.0..=MAX_DOUBLE_INT).contains(&number)
        || !(2.0..=36.0).contains(&radix)
        || !(0.0..=255.0).contains(&min_length)
    {
        return CalcValue::error(XlError::Num);
    }

    let mut result = String::new();
    let mut num = number;
    while num > 0.0 {
        let digit = (num % radix) as i32;
        num = (num / radix).floor();

        let digit_char = if digit < 10 {
            (b'0' + digit as u8) as char
        } else {
            (b'A' + (digit - 10) as u8) as char
        };
        result.insert(0, digit_char);
    }

    while result.len() < min_length as usize {
        result.insert(0, '0');
    }

    if result.is_empty() {
        result = "0".to_string();
    }

    CalcValue::text(result)
}

pub fn fn_decimal(args: &[FunctionArg<'_>], _ctx: &EvalContext<'_>) -> CalcValue {
    let text = match require_text(&args[0]) {
        Ok(v) => v,
        Err(e) => return e,
    };
    let radix = match require_number(&args[1]) {
        Ok(v) => v,
        Err(e) => return e,
    };

    let radix = radix.trunc();
    if !(2.0..=36.0).contains(&radix) {
        return CalcValue::error(XlError::Num);
    }

    if text.len() > 255 {
        return CalcValue::error(XlError::Value);
    }

    let mut result = 0.0;
    for ch in text.trim_start().chars() {
        let digit_number = match ch {
            '0'..='9' => (ch as i32) - ('0' as i32),
            'A'..='Z' => (ch as i32) - ('A' as i32) + 10,
            'a'..='z' => (ch as i32) - ('a' as i32) + 10,
            _ => return CalcValue::error(XlError::Num),
        };

        if digit_number as f64 > radix - 1.0 {
            return CalcValue::error(XlError::Num);
        }

        result = result * radix + digit_number as f64;

        if result.is_infinite() {
            return CalcValue::error(XlError::Num);
        }
    }

    CalcValue::number(result)
}

pub fn fn_roman(args: &[FunctionArg<'_>], _ctx: &EvalContext<'_>) -> CalcValue {
    let number = match require_number(&args[0]) {
        Ok(v) => v,
        Err(e) => return e,
    };
    let form = if args.len() > 1 {
        match require_number(&args[1]) {
            Ok(v) => v,
            Err(e) => return e,
        }
    } else {
        0.0
    };

    if number == 0.0 {
        return CalcValue::text(String::new());
    }

    if !(0.0..=3999.0).contains(&number) {
        return CalcValue::error(XlError::Value);
    }

    let form = form.trunc() as i32;
    if !(0..=4).contains(&form) {
        return CalcValue::error(XlError::Value);
    }

    // Simplified Roman numeral conversion (form 0 - classic)
    let mut number = number as i32;
    let mut result = String::new();

    let values = [
        (1000, "M"),
        (900, "CM"),
        (500, "D"),
        (400, "CD"),
        (100, "C"),
        (90, "XC"),
        (50, "L"),
        (40, "XL"),
        (10, "X"),
        (9, "IX"),
        (5, "V"),
        (4, "IV"),
        (1, "I"),
    ];

    for (value, symbol) in &values {
        while number >= *value {
            result.push_str(symbol);
            number -= *value;
        }
    }

    CalcValue::text(result)
}

// ============================================================================
// PRODUCT, QUOTIENT
// ============================================================================

pub fn fn_product(args: &[FunctionArg<'_>], _ctx: &EvalContext<'_>) -> CalcValue {
    let mut result = 1.0;
    let mut has_values = false;

    for arg in args {
        match arg {
            FunctionArg::Value(v) => {
                let scalar = v.as_scalar();
                if let ScalarValue::Error(e) = scalar {
                    return CalcValue::error(*e);
                }
                if let ScalarValue::Number(n) = scalar.to_number() {
                    result *= n;
                    has_values = true;
                }
            }
            FunctionArg::Range { range, ctx: rctx } => {
                for val in iter_range_values(range, rctx) {
                    if let ScalarValue::Error(e) = val {
                        return CalcValue::error(e);
                    }
                    if let ScalarValue::Number(n) = val {
                        result *= n;
                        has_values = true;
                    }
                }
            }
        }
    }

    CalcValue::number(if has_values { result } else { 0.0 })
}

pub fn fn_quotient(args: &[FunctionArg<'_>], _ctx: &EvalContext<'_>) -> CalcValue {
    let dividend = match require_number(&args[0]) {
        Ok(v) => v,
        Err(e) => return e,
    };
    let divisor = match require_number(&args[1]) {
        Ok(v) => v,
        Err(e) => return e,
    };

    if divisor == 0.0 {
        return CalcValue::error(XlError::Div0);
    }

    CalcValue::number((dividend / divisor).trunc())
}

// ============================================================================
// RANDOM
// ============================================================================

static RNG: Mutex<Option<fastrand::Rng>> = Mutex::new(None);

fn get_rng() -> f64 {
    let Ok(mut rng) = RNG.lock() else {
        // If mutex is poisoned, create a new RNG instance
        return fastrand::Rng::new().f64();
    };
    if rng.is_none() {
        *rng = Some(fastrand::Rng::new());
    }
    if let Some(r) = rng.as_mut() {
        r.f64()
    } else {
        // Should never happen, but fallback to new RNG
        fastrand::Rng::new().f64()
    }
}

pub fn fn_rand(_args: &[FunctionArg<'_>], _ctx: &EvalContext<'_>) -> CalcValue {
    CalcValue::number(get_rng())
}

pub fn fn_randbetween(args: &[FunctionArg<'_>], _ctx: &EvalContext<'_>) -> CalcValue {
    let lower_bound = match require_number(&args[0]) {
        Ok(v) => v,
        Err(e) => return e,
    };
    let upper_bound = match require_number(&args[1]) {
        Ok(v) => v,
        Err(e) => return e,
    };

    if lower_bound > upper_bound {
        return CalcValue::error(XlError::Num);
    }

    let lower = lower_bound.ceil();
    let upper = upper_bound.ceil();

    let range = upper - lower;
    let random_value = lower + (get_rng() * range).round();

    CalcValue::number(random_value)
}

// ============================================================================
// AGGREGATE FUNCTIONS
// ============================================================================

pub fn fn_sumproduct(args: &[FunctionArg<'_>], _ctx: &EvalContext<'_>) -> CalcValue {
    if args.is_empty() {
        return CalcValue::error(XlError::Value);
    }

    // Collect all arrays
    let mut arrays: Vec<Vec<ScalarValue>> = Vec::new();
    let mut width = 0;
    let mut height = 0;

    for arg in args {
        let mut array_values = Vec::new();
        let (w, h) = match arg {
            FunctionArg::Value(v) => {
                let scalar = v.as_scalar();
                if scalar == &ScalarValue::Blank {
                    return CalcValue::error(XlError::Value);
                }
                array_values.push(scalar.clone());
                (1, 1)
            }
            FunctionArg::Range { range, ctx: rctx } => {
                let w = (range.end_col - range.start_col + 1) as usize;
                let h = (range.end_row - range.start_row + 1) as usize;
                for val in iter_range_values(range, rctx) {
                    array_values.push(val);
                }
                (w, h)
            }
        };

        if width == 0 {
            width = w;
        }
        if height == 0 {
            height = h;
        }

        if width != w || height != h {
            return CalcValue::error(XlError::Value);
        }

        arrays.push(array_values);
    }

    let mut sum = 0.0;
    for idx in 0..(width * height) {
        let mut product = 1.0;
        for array in &arrays {
            let scalar = &array[idx];

            if let ScalarValue::Error(e) = scalar {
                return CalcValue::error(*e);
            }

            let number = match scalar {
                ScalarValue::Number(n) => *n,
                _ => 0.0,
            };

            product *= number;
        }
        sum += product;
    }

    CalcValue::number(sum)
}

pub fn fn_sumsq(args: &[FunctionArg<'_>], ctx: &EvalContext<'_>) -> CalcValue {
    match collect_numbers(args, ctx) {
        Ok(nums) => {
            let sum: f64 = nums.iter().map(|n| n * n).sum();
            CalcValue::number(sum)
        }
        Err(e) => e,
    }
}

pub fn fn_sumx2my2(args: &[FunctionArg<'_>], ctx: &EvalContext<'_>) -> CalcValue {
    let array_x = collect_array(&args[0], ctx);
    let array_y = collect_array(&args[1], ctx);

    if array_x.len() != array_y.len() {
        return CalcValue::error(XlError::Na);
    }

    let mut sum = 0.0;
    for (x, y) in array_x.iter().zip(array_y.iter()) {
        sum += x * x - y * y;
    }

    CalcValue::number(sum)
}

pub fn fn_sumx2py2(args: &[FunctionArg<'_>], ctx: &EvalContext<'_>) -> CalcValue {
    let array_x = collect_array(&args[0], ctx);
    let array_y = collect_array(&args[1], ctx);

    if array_x.len() != array_y.len() {
        return CalcValue::error(XlError::Na);
    }

    let mut sum = 0.0;
    for (x, y) in array_x.iter().zip(array_y.iter()) {
        sum += x * x + y * y;
    }

    CalcValue::number(sum)
}

pub fn fn_sumxmy2(args: &[FunctionArg<'_>], ctx: &EvalContext<'_>) -> CalcValue {
    let array_x = collect_array(&args[0], ctx);
    let array_y = collect_array(&args[1], ctx);

    if array_x.len() != array_y.len() {
        return CalcValue::error(XlError::Na);
    }

    let mut sum = 0.0;
    for (x, y) in array_x.iter().zip(array_y.iter()) {
        let diff = x - y;
        sum += diff * diff;
    }

    CalcValue::number(sum)
}

pub fn fn_seriessum(args: &[FunctionArg<'_>], ctx: &EvalContext<'_>) -> CalcValue {
    let input = match require_number(&args[0]) {
        Ok(v) => v,
        Err(e) => return e,
    };
    let initial = match require_number(&args[1]) {
        Ok(v) => v,
        Err(e) => return e,
    };
    let step = match require_number(&args[2]) {
        Ok(v) => v,
        Err(e) => return e,
    };

    let coefficients = collect_array(&args[3], ctx);

    let mut total = 0.0;
    for (i, coef) in coefficients.iter().enumerate() {
        let power = input.powf(initial + i as f64 * step);
        total += coef * power;
        if total.is_infinite() {
            return CalcValue::error(XlError::Num);
        }
    }

    CalcValue::number(total)
}

fn collect_array(arg: &FunctionArg<'_>, _ctx: &EvalContext<'_>) -> Vec<f64> {
    let mut result = Vec::new();
    match arg {
        FunctionArg::Value(v) => {
            if let ScalarValue::Number(n) = v.as_scalar().to_number() {
                result.push(n);
            }
        }
        FunctionArg::Range { range, ctx: rctx } => {
            for val in iter_range_values(range, rctx) {
                if let ScalarValue::Number(n) = val.to_number() {
                    result.push(n);
                } else if val != ScalarValue::Blank {
                    result.push(0.0);
                }
            }
        }
    }
    result
}

// SUBTOTAL placeholder - requires statistical functions
pub fn fn_subtotal(_args: &[FunctionArg<'_>], _ctx: &EvalContext<'_>) -> CalcValue {
    CalcValue::error(XlError::Value) // Placeholder
}

// SUMIF/SUMIFS placeholders - require criteria engine
pub fn fn_sumif(_args: &[FunctionArg<'_>], _ctx: &EvalContext<'_>) -> CalcValue {
    CalcValue::error(XlError::Value) // Placeholder
}

pub fn fn_sumifs(_args: &[FunctionArg<'_>], _ctx: &EvalContext<'_>) -> CalcValue {
    CalcValue::error(XlError::Value) // Placeholder
}

// Matrix operations placeholders
pub fn fn_mdeterm(_args: &[FunctionArg<'_>], _ctx: &EvalContext<'_>) -> CalcValue {
    CalcValue::error(XlError::Value) // Placeholder
}

pub fn fn_minverse(_args: &[FunctionArg<'_>], _ctx: &EvalContext<'_>) -> CalcValue {
    CalcValue::error(XlError::Value) // Placeholder
}

pub fn fn_mmult(_args: &[FunctionArg<'_>], _ctx: &EvalContext<'_>) -> CalcValue {
    CalcValue::error(XlError::Value) // Placeholder
}
