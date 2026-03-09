//! Text functions: LEFT, RIGHT, MID, LEN, TRIM, UPPER, LOWER, CONCATENATE,
//! ASC, CHAR, CLEAN, CODE, CONCAT, DOLLAR, EXACT, FIND, FIXED, LEFTB,
//! NUMBERVALUE, PROPER, REPLACE, REPT, SEARCH, SUBSTITUTE, T, TEXT, TEXTJOIN, VALUE.

use crate::context::EvalContext;
use crate::value::{CalcValue, ScalarValue, XlError};

use super::{
    require_number, require_scalar, require_text, FunctionArg, FunctionDef, FunctionRegistry,
    ParamKind,
};

/// Windows-1252 encoding table for characters 0x80-0xFF
const WINDOWS_1252: &str =
    "\u{20AC}\u{0081}\u{201A}\u{0192}\u{201E}\u{2026}\u{2020}\u{2021}\u{02C6}\u{2030}\u{0160}\u{2039}\u{0152}\u{008D}\u{017D}\u{008F}\
     \u{0090}\u{2018}\u{2019}\u{201C}\u{201D}\u{2022}\u{2013}\u{2014}\u{02DC}\u{2122}\u{0161}\u{203A}\u{0153}\u{009D}\u{017E}\u{0178}\
     \u{00A0}\u{00A1}\u{00A2}\u{00A3}\u{00A4}\u{00A5}\u{00A6}\u{00A7}\u{00A8}\u{00A9}\u{00AA}\u{00AB}\u{00AC}\u{00AD}\u{00AE}\u{00AF}\
     \u{00B0}\u{00B1}\u{00B2}\u{00B3}\u{00B4}\u{00B5}\u{00B6}\u{00B7}\u{00B8}\u{00B9}\u{00BA}\u{00BB}\u{00BC}\u{00BD}\u{00BE}\u{00BF}\
     \u{00C0}\u{00C1}\u{00C2}\u{00C3}\u{00C4}\u{00C5}\u{00C6}\u{00C7}\u{00C8}\u{00C9}\u{00CA}\u{00CB}\u{00CC}\u{00CD}\u{00CE}\u{00CF}\
     \u{00D0}\u{00D1}\u{00D2}\u{00D3}\u{00D4}\u{00D5}\u{00D6}\u{00D7}\u{00D8}\u{00D9}\u{00DA}\u{00DB}\u{00DC}\u{00DD}\u{00DE}\u{00DF}\
     \u{00E0}\u{00E1}\u{00E2}\u{00E3}\u{00E4}\u{00E5}\u{00E6}\u{00E7}\u{00E8}\u{00E9}\u{00EA}\u{00EB}\u{00EC}\u{00ED}\u{00EE}\u{00EF}\
     \u{00F0}\u{00F1}\u{00F2}\u{00F3}\u{00F4}\u{00F5}\u{00F6}\u{00F7}\u{00F8}\u{00F9}\u{00FA}\u{00FB}\u{00FC}\u{00FD}\u{00FE}\u{00FF}";

/// Add thousands separator to a number string.
fn add_thousands_separator(s: &str) -> String {
    let (sign, rest) = if let Some(stripped) = s.strip_prefix('-') {
        ("-", stripped)
    } else {
        ("", s)
    };

    let (integer, decimal) = if let Some(pos) = rest.find('.') {
        (&rest[..pos], &rest[pos..])
    } else {
        (rest, "")
    };

    let mut result = String::new();
    let len = integer.len();
    for (i, ch) in integer.chars().enumerate() {
        if i > 0 && (len - i) % 3 == 0 {
            result.push(',');
        }
        result.push(ch);
    }

    format!("{}{}{}", sign, result, decimal)
}

/// `LEFT(text, [num_chars=1])`
///
/// Returns the leftmost characters of a string.
pub fn fn_left(args: &[FunctionArg<'_>], _ctx: &EvalContext<'_>) -> CalcValue {
    let text = match args.first() {
        Some(a) => match require_text(a) {
            Ok(s) => s,
            Err(e) => return e,
        },
        None => return CalcValue::error(XlError::Value),
    };
    let num_chars = match args.get(1) {
        Some(a) => match require_number(a) {
            Ok(n) => n as i64,
            Err(e) => return e,
        },
        None => 1,
    };
    if num_chars < 0 {
        return CalcValue::error(XlError::Value);
    }
    let n = num_chars as usize;
    let result: String = text.chars().take(n).collect();
    CalcValue::text(result)
}

/// `RIGHT(text, [num_chars=1])`
///
/// Returns the rightmost characters of a string.
pub fn fn_right(args: &[FunctionArg<'_>], _ctx: &EvalContext<'_>) -> CalcValue {
    let text = match args.first() {
        Some(a) => match require_text(a) {
            Ok(s) => s,
            Err(e) => return e,
        },
        None => return CalcValue::error(XlError::Value),
    };
    let num_chars = match args.get(1) {
        Some(a) => match require_number(a) {
            Ok(n) => n as i64,
            Err(e) => return e,
        },
        None => 1,
    };
    if num_chars < 0 {
        return CalcValue::error(XlError::Value);
    }
    let n = num_chars as usize;
    let char_count = text.chars().count();
    let skip = char_count.saturating_sub(n);
    let result: String = text.chars().skip(skip).collect();
    CalcValue::text(result)
}

/// `MID(text, start_num, num_chars)`
///
/// Returns characters from the middle of a string. `start_num` is 1-based.
pub fn fn_mid(args: &[FunctionArg<'_>], _ctx: &EvalContext<'_>) -> CalcValue {
    let text = match args.first() {
        Some(a) => match require_text(a) {
            Ok(s) => s,
            Err(e) => return e,
        },
        None => return CalcValue::error(XlError::Value),
    };
    let start_num = match args.get(1) {
        Some(a) => match require_number(a) {
            Ok(n) => n as i64,
            Err(e) => return e,
        },
        None => return CalcValue::error(XlError::Value),
    };
    let num_chars = match args.get(2) {
        Some(a) => match require_number(a) {
            Ok(n) => n as i64,
            Err(e) => return e,
        },
        None => return CalcValue::error(XlError::Value),
    };
    if start_num < 1 {
        return CalcValue::error(XlError::Value);
    }
    if num_chars < 0 {
        return CalcValue::error(XlError::Value);
    }
    let start = (start_num - 1) as usize;
    let n = num_chars as usize;
    let result: String = text.chars().skip(start).take(n).collect();
    CalcValue::text(result)
}

/// `LEN(text)`
///
/// Returns the number of characters in a string.
pub fn fn_len(args: &[FunctionArg<'_>], _ctx: &EvalContext<'_>) -> CalcValue {
    let text = match args.first() {
        Some(a) => match require_text(a) {
            Ok(s) => s,
            Err(e) => return e,
        },
        None => return CalcValue::error(XlError::Value),
    };
    CalcValue::number(text.len() as f64)
}

/// `TRIM(text)`
///
/// Removes leading/trailing spaces and collapses internal runs of spaces to a single space.
pub fn fn_trim(args: &[FunctionArg<'_>], _ctx: &EvalContext<'_>) -> CalcValue {
    let text = match args.first() {
        Some(a) => match require_text(a) {
            Ok(s) => s,
            Err(e) => return e,
        },
        None => return CalcValue::error(XlError::Value),
    };
    let trimmed: Vec<&str> = text.split(' ').filter(|s| !s.is_empty()).collect();
    CalcValue::text(trimmed.join(" "))
}

/// `UPPER(text)`
///
/// Converts text to uppercase.
pub fn fn_upper(args: &[FunctionArg<'_>], _ctx: &EvalContext<'_>) -> CalcValue {
    let text = match args.first() {
        Some(a) => match require_text(a) {
            Ok(s) => s,
            Err(e) => return e,
        },
        None => return CalcValue::error(XlError::Value),
    };
    CalcValue::text(text.to_uppercase())
}

/// `LOWER(text)`
///
/// Converts text to lowercase.
pub fn fn_lower(args: &[FunctionArg<'_>], _ctx: &EvalContext<'_>) -> CalcValue {
    let text = match args.first() {
        Some(a) => match require_text(a) {
            Ok(s) => s,
            Err(e) => return e,
        },
        None => return CalcValue::error(XlError::Value),
    };
    CalcValue::text(text.to_lowercase())
}

/// `CONCATENATE(text1, text2, ...)`
///
/// Concatenates all arguments as text.
pub fn fn_concatenate(args: &[FunctionArg<'_>], _ctx: &EvalContext<'_>) -> CalcValue {
    let mut result = String::new();
    for arg in args {
        match require_text(arg) {
            Ok(s) => {
                result.push_str(&s);
                if result.len() > 32767 {
                    return CalcValue::error(XlError::Value);
                }
            }
            Err(e) => return e,
        }
    }
    CalcValue::text(result)
}

/// `ASC(text)`
///
/// Converts full-width characters to half-width.
pub fn fn_asc(args: &[FunctionArg<'_>], _ctx: &EvalContext<'_>) -> CalcValue {
    let text = match args.first() {
        Some(a) => match require_text(a) {
            Ok(s) => s,
            Err(e) => return e,
        },
        None => return CalcValue::error(XlError::Value),
    };

    let result: String = text.chars().map(to_half_width).collect();
    CalcValue::text(result)
}

fn to_half_width(c: char) -> char {
    let cp = c as u32;
    let converted = match cp {
        // Katakana conversions
        0x30A1..=0x30AA if cp.is_multiple_of(2) => (cp - 0x30A2) / 2 + 0xFF71,
        0x30A1..=0x30AA => (cp - 0x30A1) / 2 + 0xFF67,
        0x30AB..=0x30C2 if cp % 2 == 1 => (cp - 0x30AB) / 2 + 0xFF76,
        0x30AB..=0x30C2 => (cp - 0x30AC) / 2 + 0xFF76,
        0x30C3 => 0xFF6F,
        0x30C4..=0x30C9 if cp.is_multiple_of(2) => (cp - 0x30C4) / 2 + 0xFF82,
        0x30C4..=0x30C9 => (cp - 0x30C5) / 2 + 0xFF82,
        0x30CA..=0x30CE => cp - 0x30CA + 0xFF85,
        0x30CF..=0x30DD if cp.is_multiple_of(3) => (cp - 0x30CF) / 3 + 0xFF8A,
        0x30CF..=0x30DD if cp % 3 == 1 => (cp - 0x30D0) / 3 + 0xFF8A,
        0x30CF..=0x30DD => (cp - 0x30D1) / 3 + 0xFF8A,
        0x30DE..=0x30E2 => cp - 0x30DE + 0xFF8F,
        0x30E3..=0x30E8 if cp.is_multiple_of(2) => (cp - 0x30E4) / 2 + 0xFF94,
        0x30E3..=0x30E8 => (cp - 0x30E3) / 2 + 0xFF6C,
        0x30E9..=0x30ED => cp - 0x30E9 + 0xFF97,
        0x30EF => 0xFF9C,
        0x30F2 => 0xFF66,
        0x30F3 => 0xFF9D,
        // ASCII fullwidth
        0xFF01..=0xFF5E => cp - 0xFF01 + 0x0021,
        // Special characters
        0x2015 => 0xFF70,
        0x2018 => 0x0060,
        0x2019 => 0x0027,
        0x201D => 0x0022,
        0x3001 => 0xFF64,
        0x3002 => 0xFF61,
        0x300C => 0xFF62,
        0x300D => 0xFF63,
        0x309B => 0xFF9E,
        0x309C => 0xFF9F,
        0x30FB => 0xFF65,
        0x30FC => 0xFF70,
        0xFFE5 => 0x005C,
        _ => cp,
    };

    char::from_u32(converted).unwrap_or(c)
}

/// `CHAR(number)`
///
/// Returns the character specified by the code number (1-255, Windows-1252).
pub fn fn_char(args: &[FunctionArg<'_>], _ctx: &EvalContext<'_>) -> CalcValue {
    let number = match args.first() {
        Some(a) => match require_number(a) {
            Ok(n) => n.trunc(),
            Err(e) => return e,
        },
        None => return CalcValue::error(XlError::Value),
    };

    if !(1.0..=255.0).contains(&number) {
        return CalcValue::error(XlError::Value);
    }

    let code = number as u8;
    let ch = if code < 128 {
        code as char
    } else {
        WINDOWS_1252
            .chars()
            .nth((code - 128) as usize)
            .unwrap_or('?')
    };

    CalcValue::text(ch.to_string())
}

/// `CLEAN(text)`
///
/// Removes all nonprintable characters (0x00-0x1F and 0x80-0x9F).
pub fn fn_clean(args: &[FunctionArg<'_>], _ctx: &EvalContext<'_>) -> CalcValue {
    let text = match args.first() {
        Some(a) => match require_text(a) {
            Ok(s) => s,
            Err(e) => return e,
        },
        None => return CalcValue::error(XlError::Value),
    };

    let result: String = text
        .chars()
        .filter(|&c| {
            let cp = c as u32;
            !(cp <= 0x1F || (0x80..=0x9F).contains(&cp))
        })
        .collect();

    CalcValue::text(result)
}

/// `CODE(text)`
///
/// Returns the numeric code for the first character (Windows-1252).
pub fn fn_code(args: &[FunctionArg<'_>], _ctx: &EvalContext<'_>) -> CalcValue {
    let text = match args.first() {
        Some(a) => match require_text(a) {
            Ok(s) => s,
            Err(e) => return e,
        },
        None => return CalcValue::error(XlError::Value),
    };

    if text.is_empty() {
        return CalcValue::error(XlError::Value);
    }

    let first = text.chars().next().unwrap_or('?');
    let code = if (first as u32) < 128 {
        first as u32
    } else {
        // Find in Windows-1252 table
        WINDOWS_1252
            .chars()
            .position(|c| c == first)
            .map(|pos| (pos + 128) as u32)
            .unwrap_or(63) // '?' fallback
    };

    CalcValue::number(code as f64)
}

/// `CONCAT(text1, text2, ...)`
///
/// Concatenates text from arguments (newer version of CONCATENATE).
pub fn fn_concat(args: &[FunctionArg<'_>], _ctx: &EvalContext<'_>) -> CalcValue {
    let mut result = String::new();
    for arg in args {
        let scalar = require_scalar(arg);
        if let ScalarValue::Error(e) = scalar {
            return CalcValue::error(*e);
        }
        let text = scalar.to_text();
        match text {
            ScalarValue::Text(s) => {
                result.push_str(&s);
                if result.len() > 32767 {
                    return CalcValue::error(XlError::Value);
                }
            }
            ScalarValue::Error(e) => return CalcValue::error(e),
            _ => {}
        }
    }
    CalcValue::text(result)
}

/// `DOLLAR(number, [decimals=2])`
///
/// Formats a number with currency format.
pub fn fn_dollar(args: &[FunctionArg<'_>], _ctx: &EvalContext<'_>) -> CalcValue {
    let number = match args.first() {
        Some(a) => match require_number(a) {
            Ok(n) => n,
            Err(e) => return e,
        },
        None => return CalcValue::error(XlError::Value),
    };

    let decimals = match args.get(1) {
        Some(a) => match require_number(a) {
            Ok(n) => n.trunc() as i32,
            Err(e) => return e,
        },
        None => 2,
    };

    if decimals > 99 {
        return CalcValue::error(XlError::Value);
    }

    let formatted = if decimals >= 0 {
        let base = format!("{:.prec$}", number.abs(), prec = decimals as usize);
        let with_commas = add_thousands_separator(&base);
        format!("${}{}", if number < 0.0 { "-" } else { "" }, with_commas)
    } else {
        let factor = 10_f64.powi(-decimals);
        let rounded = (number / factor).round() * factor;
        let base = format!("{:.0}", rounded.abs());
        let with_commas = add_thousands_separator(&base);
        format!("${}{}", if rounded < 0.0 { "-" } else { "" }, with_commas)
    };

    CalcValue::text(formatted)
}

/// `EXACT(text1, text2)`
///
/// Checks if two text values are identical (case-sensitive).
pub fn fn_exact(args: &[FunctionArg<'_>], _ctx: &EvalContext<'_>) -> CalcValue {
    let text1 = match args.first() {
        Some(a) => match require_text(a) {
            Ok(s) => s,
            Err(e) => return e,
        },
        None => return CalcValue::error(XlError::Value),
    };

    let text2 = match args.get(1) {
        Some(a) => match require_text(a) {
            Ok(s) => s,
            Err(e) => return e,
        },
        None => return CalcValue::error(XlError::Value),
    };

    CalcValue::bool(text1 == text2)
}

/// `FIND(find_text, within_text, [start_num=1])`
///
/// Finds one text value within another (case-sensitive). Returns 1-based position.
pub fn fn_find(args: &[FunctionArg<'_>], _ctx: &EvalContext<'_>) -> CalcValue {
    let find_text = match args.first() {
        Some(a) => match require_text(a) {
            Ok(s) => s,
            Err(e) => return e,
        },
        None => return CalcValue::error(XlError::Value),
    };

    let within_text = match args.get(1) {
        Some(a) => match require_text(a) {
            Ok(s) => s,
            Err(e) => return e,
        },
        None => return CalcValue::error(XlError::Value),
    };

    let start_num = match args.get(2) {
        Some(a) => match require_number(a) {
            Ok(n) => n.trunc() as i32,
            Err(e) => return e,
        },
        None => 1,
    };

    if start_num < 1 {
        return CalcValue::error(XlError::Value);
    }

    let start_idx = (start_num - 1) as usize;
    if start_idx > within_text.len() {
        return CalcValue::error(XlError::Value);
    }

    match within_text[start_idx..].find(&find_text) {
        Some(pos) => CalcValue::number((start_idx + pos + 1) as f64),
        None => CalcValue::error(XlError::Value),
    }
}

/// `FIXED(number, [decimals=2], [no_commas=false])`
///
/// Formats a number as text with fixed number of decimals.
pub fn fn_fixed(args: &[FunctionArg<'_>], _ctx: &EvalContext<'_>) -> CalcValue {
    let number = match args.first() {
        Some(a) => match require_number(a) {
            Ok(n) => n,
            Err(e) => return e,
        },
        None => return CalcValue::error(XlError::Value),
    };

    let decimals = match args.get(1) {
        Some(a) => match require_number(a) {
            Ok(n) => n.trunc() as i32,
            Err(e) => return e,
        },
        None => 2,
    };

    let no_commas = match args.get(2) {
        Some(a) => {
            let scalar = require_scalar(a);
            matches!(scalar.to_bool(), ScalarValue::Bool(true))
        }
        None => false,
    };

    if decimals > 99 {
        return CalcValue::error(XlError::Value);
    }

    let dec_places = decimals.max(0) as usize;
    let rounded = if decimals >= 0 {
        let factor = 10_f64.powi(decimals);
        (number * factor).round() / factor
    } else {
        let factor = 10_f64.powi(-decimals);
        (number / factor).round() * factor
    };

    let formatted = if no_commas {
        format!("{:.prec$}", rounded, prec = dec_places)
    } else {
        let base = format!("{:.prec$}", rounded, prec = dec_places);
        add_thousands_separator(&base)
    };

    CalcValue::text(formatted)
}

/// `LEFTB(text, [num_bytes=1])`
///
/// Returns leftmost bytes (treating each char as 1 or 2 bytes).
pub fn fn_leftb(args: &[FunctionArg<'_>], _ctx: &EvalContext<'_>) -> CalcValue {
    let text = match args.first() {
        Some(a) => match require_text(a) {
            Ok(s) => s,
            Err(e) => return e,
        },
        None => return CalcValue::error(XlError::Value),
    };

    let num_bytes = match args.get(1) {
        Some(a) => match require_number(a) {
            Ok(n) => n as i64,
            Err(e) => return e,
        },
        None => 1,
    };

    if num_bytes < 0 {
        return CalcValue::error(XlError::Value);
    }

    // Simplified: treat ASCII as 1 byte, others as 2
    let mut byte_count = 0;
    let result: String = text
        .chars()
        .take_while(|&c| {
            let size = if (c as u32) < 128 { 1 } else { 2 };
            if byte_count + size <= num_bytes as usize {
                byte_count += size;
                true
            } else {
                false
            }
        })
        .collect();

    CalcValue::text(result)
}

/// `NUMBERVALUE(text, [decimal_separator="."], [group_separator=","])`
///
/// Converts text to a number.
pub fn fn_numbervalue(args: &[FunctionArg<'_>], _ctx: &EvalContext<'_>) -> CalcValue {
    let text = match args.first() {
        Some(a) => match require_text(a) {
            Ok(s) => s,
            Err(e) => return e,
        },
        None => return CalcValue::error(XlError::Value),
    };

    let decimal_sep = match args.get(1) {
        Some(a) => match require_text(a) {
            Ok(s) => s,
            Err(e) => return e,
        },
        None => ".".to_string(),
    };

    let group_sep = match args.get(2) {
        Some(a) => match require_text(a) {
            Ok(s) => s,
            Err(e) => return e,
        },
        None => ",".to_string(),
    };

    if decimal_sep.is_empty() || group_sep.is_empty() {
        return CalcValue::error(XlError::Value);
    }

    if decimal_sep == group_sep {
        return CalcValue::error(XlError::Value);
    }

    if text.is_empty() {
        return CalcValue::number(0.0);
    }

    let dec_char = decimal_sep.chars().next().unwrap_or('.');
    let grp_char = group_sep.chars().next().unwrap_or(',');

    // Process text
    let mut processed = String::new();
    let mut decimal_seen = false;

    for c in text.chars() {
        if c == dec_char {
            if !decimal_seen {
                processed.push('.');
                decimal_seen = true;
            } else {
                processed.push(c);
            }
        } else if c == grp_char && !decimal_seen {
            // Skip group separators before decimal
        } else if !c.is_whitespace() {
            processed.push(c);
        }
    }

    // Add leading zero if starts with '.'
    if processed.starts_with('.') {
        processed.insert(0, '0');
    }

    // Count trailing % signs
    let percent_count = processed.chars().rev().take_while(|&c| c == '%').count();
    if percent_count > 0 {
        processed.truncate(processed.len() - percent_count);
    }

    match processed.parse::<f64>() {
        Ok(mut n) => {
            for _ in 0..percent_count {
                n /= 100.0;
            }
            CalcValue::number(n)
        }
        Err(_) => CalcValue::error(XlError::Value),
    }
}

/// `PROPER(text)`
///
/// Capitalizes the first letter of each word.
pub fn fn_proper(args: &[FunctionArg<'_>], _ctx: &EvalContext<'_>) -> CalcValue {
    let text = match args.first() {
        Some(a) => match require_text(a) {
            Ok(s) => s,
            Err(e) => return e,
        },
        None => return CalcValue::error(XlError::Value),
    };

    if text.is_empty() {
        return CalcValue::text(String::new());
    }

    let mut result = String::with_capacity(text.len());
    let mut prev_was_letter = false;

    for c in text.chars() {
        let cased = if prev_was_letter {
            c.to_lowercase().collect::<String>()
        } else {
            c.to_uppercase().collect::<String>()
        };
        result.push_str(&cased);
        prev_was_letter = c.is_alphabetic();
    }

    CalcValue::text(result)
}

/// `REPLACE(old_text, start_num, num_chars, new_text)`
///
/// Replaces part of a text string with another string.
pub fn fn_replace(args: &[FunctionArg<'_>], _ctx: &EvalContext<'_>) -> CalcValue {
    let old_text = match args.first() {
        Some(a) => match require_text(a) {
            Ok(s) => s,
            Err(e) => return e,
        },
        None => return CalcValue::error(XlError::Value),
    };

    let start_num = match args.get(1) {
        Some(a) => match require_number(a) {
            Ok(n) => n.trunc() as i32,
            Err(e) => return e,
        },
        None => return CalcValue::error(XlError::Value),
    };

    let num_chars = match args.get(2) {
        Some(a) => match require_number(a) {
            Ok(n) => n.trunc() as i32,
            Err(e) => return e,
        },
        None => return CalcValue::error(XlError::Value),
    };

    let new_text = match args.get(3) {
        Some(a) => match require_text(a) {
            Ok(s) => s,
            Err(e) => return e,
        },
        None => return CalcValue::error(XlError::Value),
    };

    if start_num < 1 || num_chars < 0 {
        return CalcValue::error(XlError::Value);
    }

    let start_idx = (start_num - 1) as usize;
    let delete_len = num_chars as usize;

    let prefix_len = start_idx.min(old_text.len());
    let delete_end = (start_idx + delete_len).min(old_text.len());

    let mut result = String::new();
    result.push_str(&old_text[..prefix_len]);
    result.push_str(&new_text);
    if delete_end < old_text.len() {
        result.push_str(&old_text[delete_end..]);
    }

    CalcValue::text(result)
}

/// `REPT(text, number_times)`
///
/// Repeats text a given number of times.
pub fn fn_rept(args: &[FunctionArg<'_>], _ctx: &EvalContext<'_>) -> CalcValue {
    let text = match args.first() {
        Some(a) => match require_text(a) {
            Ok(s) => s,
            Err(e) => return e,
        },
        None => return CalcValue::error(XlError::Value),
    };

    let number = match args.get(1) {
        Some(a) => match require_number(a) {
            Ok(n) => n.trunc() as i32,
            Err(e) => return e,
        },
        None => return CalcValue::error(XlError::Value),
    };

    if number < 0 {
        return CalcValue::error(XlError::Value);
    }

    if text.is_empty() {
        return CalcValue::text(String::new());
    }

    let total_len = text.len() * number as usize;
    if total_len > 32767 {
        return CalcValue::error(XlError::Value);
    }

    CalcValue::text(text.repeat(number as usize))
}

/// `SEARCH(find_text, within_text, [start_num=1])`
///
/// Finds text using wildcards (case-insensitive). Returns 1-based position.
pub fn fn_search(args: &[FunctionArg<'_>], _ctx: &EvalContext<'_>) -> CalcValue {
    let find_text = match args.first() {
        Some(a) => match require_text(a) {
            Ok(s) => s,
            Err(e) => return e,
        },
        None => return CalcValue::error(XlError::Value),
    };

    let within_text = match args.get(1) {
        Some(a) => match require_text(a) {
            Ok(s) => s,
            Err(e) => return e,
        },
        None => return CalcValue::error(XlError::Value),
    };

    let start_num = match args.get(2) {
        Some(a) => match require_number(a) {
            Ok(n) => n.trunc() as i32,
            Err(e) => return e,
        },
        None => 1,
    };

    if within_text.is_empty() || start_num < 1 {
        return CalcValue::error(XlError::Value);
    }

    let start_idx = (start_num - 1) as usize;
    if start_idx >= within_text.len() {
        return CalcValue::error(XlError::Value);
    }

    let search_text = &within_text[start_idx..];

    match wildcard_search(&find_text, search_text) {
        Some(pos) => CalcValue::number((start_idx + pos + 1) as f64),
        None => CalcValue::error(XlError::Value),
    }
}

/// Simple wildcard search (supports * and ? and ~ escape).
fn wildcard_search(pattern: &str, text: &str) -> Option<usize> {
    if pattern.len() > 255 {
        return None;
    }

    for i in 0..=text.len() {
        if wildcard_match(pattern, &text[i..]) {
            return Some(i);
        }
    }
    None
}

fn wildcard_match(pattern: &str, text: &str) -> bool {
    let pat_chars: Vec<char> = pattern.chars().collect();
    let txt_chars: Vec<char> = text.chars().collect();

    let mut ti = 0; // text index
    let mut pi = 0; // pattern index
    let mut star_idx = None; // last * position
    let mut match_idx = 0; // text position at last *

    while ti < txt_chars.len() {
        if pi < pat_chars.len() {
            if pat_chars[pi] == '?' {
                ti += 1;
                pi += 1;
                continue;
            }

            if pat_chars[pi] == '*' {
                star_idx = Some(pi);
                match_idx = ti;
                pi += 1;
                continue;
            }

            let mut pat_char = pat_chars[pi];
            if pat_char == '~' && pi + 1 < pat_chars.len() {
                pi += 1;
                pat_char = pat_chars[pi];
            }

            if pat_char.to_uppercase().to_string() == txt_chars[ti].to_uppercase().to_string() {
                ti += 1;
                pi += 1;
                continue;
            }
        }

        // Mismatch - try backtracking
        if let Some(si) = star_idx {
            match_idx += 1;
            ti = match_idx;
            pi = si + 1;
        } else {
            return false;
        }
    }

    // Consume remaining *
    while pi < pat_chars.len() && pat_chars[pi] == '*' {
        pi += 1;
    }

    pi == pat_chars.len()
}

/// `SUBSTITUTE(text, old_text, new_text, [instance_num])`
///
/// Substitutes new text for old text in a string.
pub fn fn_substitute(args: &[FunctionArg<'_>], _ctx: &EvalContext<'_>) -> CalcValue {
    let text = match args.first() {
        Some(a) => match require_text(a) {
            Ok(s) => s,
            Err(e) => return e,
        },
        None => return CalcValue::error(XlError::Value),
    };

    let old_text = match args.get(1) {
        Some(a) => match require_text(a) {
            Ok(s) => s,
            Err(e) => return e,
        },
        None => return CalcValue::error(XlError::Value),
    };

    let new_text = match args.get(2) {
        Some(a) => match require_text(a) {
            Ok(s) => s,
            Err(e) => return e,
        },
        None => return CalcValue::error(XlError::Value),
    };

    let instance_num = match args.get(3) {
        Some(a) => {
            let n = match require_number(a) {
                Ok(n) => n.trunc() as i32,
                Err(e) => return e,
            };
            if n < 1 {
                return CalcValue::error(XlError::Value);
            }
            Some(n as usize)
        }
        None => None,
    };

    if text.is_empty() || old_text.is_empty() {
        return CalcValue::text(text);
    }

    let result = if let Some(instance) = instance_num {
        // Replace only specific occurrence
        let mut count = 0;
        let mut pos = 0;
        let mut found_pos = None;

        while let Some(idx) = text[pos..].find(&old_text) {
            count += 1;
            if count == instance {
                found_pos = Some(pos + idx);
                break;
            }
            pos += idx + old_text.len();
        }

        if let Some(idx) = found_pos {
            let mut result = String::new();
            result.push_str(&text[..idx]);
            result.push_str(&new_text);
            result.push_str(&text[idx + old_text.len()..]);
            result
        } else {
            text
        }
    } else {
        // Replace all occurrences
        text.replace(&old_text, &new_text)
    };

    CalcValue::text(result)
}

/// `T(value)`
///
/// Returns the text if the value is text, otherwise returns empty string.
pub fn fn_t(args: &[FunctionArg<'_>], _ctx: &EvalContext<'_>) -> CalcValue {
    let value = match args.first() {
        Some(a) => require_scalar(a),
        None => return CalcValue::error(XlError::Value),
    };

    if let ScalarValue::Error(e) = value {
        return CalcValue::error(*e);
    }

    match value {
        ScalarValue::Text(s) => CalcValue::text(s.clone()),
        _ => CalcValue::text(String::new()),
    }
}

/// `TEXT(value, format_text)`
///
/// Formats a number as text using a format code.
pub fn fn_text(args: &[FunctionArg<'_>], _ctx: &EvalContext<'_>) -> CalcValue {
    let value = match args.first() {
        Some(a) => require_scalar(a),
        None => return CalcValue::error(XlError::Value),
    };

    let format_text = match args.get(1) {
        Some(a) => match require_text(a) {
            Ok(s) => s,
            Err(e) => return e,
        },
        None => return CalcValue::error(XlError::Value),
    };

    // If value is not convertible to number or is logical, convert to text
    let number = match value.to_number() {
        ScalarValue::Number(n) => {
            if matches!(value, ScalarValue::Bool(_)) {
                return value.to_text().into();
            }
            n
        }
        ScalarValue::Error(e) => return CalcValue::error(e),
        _ => return value.to_text().into(),
    };

    if format_text.trim().is_empty() {
        return CalcValue::text(format_text);
    }

    // Simple format code parsing (basic implementation)
    let formatted = simple_number_format(number, &format_text);

    CalcValue::text(formatted)
}

/// Simple number formatter (handles basic Excel format codes).
fn simple_number_format(value: f64, format: &str) -> String {
    // Basic format code handling
    if format.contains("0.00") || format.contains("#.##") {
        format!("{:.2}", value)
    } else if format.contains("0.0") || format.contains("#.#") {
        format!("{:.1}", value)
    } else if format.contains("0") || format.contains("#") {
        format!("{:.0}", value)
    } else if format.contains('%') {
        format!("{:.0}%", value * 100.0)
    } else if format.contains("$") {
        format!("${:.2}", value)
    } else {
        format!("{}", value)
    }
}

/// `TEXTJOIN(delimiter, ignore_empty, text1, [text2], ...)`
///
/// Joins text with a delimiter.
pub fn fn_textjoin(args: &[FunctionArg<'_>], _ctx: &EvalContext<'_>) -> CalcValue {
    let delimiter = match args.first() {
        Some(a) => match require_text(a) {
            Ok(s) => s,
            Err(e) => return e,
        },
        None => return CalcValue::error(XlError::Value),
    };

    let ignore_empty = match args.get(1) {
        Some(a) => {
            let scalar = require_scalar(a);
            matches!(scalar.to_bool(), ScalarValue::Bool(true))
        }
        None => return CalcValue::error(XlError::Value),
    };

    let mut parts = Vec::new();

    for arg in args.iter().skip(2) {
        let scalar = require_scalar(arg);
        if let ScalarValue::Error(e) = scalar {
            return CalcValue::error(*e);
        }

        match scalar.to_text() {
            ScalarValue::Text(s) => {
                if !ignore_empty || !s.is_empty() {
                    parts.push(s);
                }
            }
            ScalarValue::Error(e) => return CalcValue::error(e),
            _ => {}
        }
    }

    let result = parts.join(&delimiter);
    if result.len() > 32767 {
        return CalcValue::error(XlError::Value);
    }

    CalcValue::text(result)
}

/// `VALUE(text)`
///
/// Converts text to a number.
pub fn fn_value(args: &[FunctionArg<'_>], _ctx: &EvalContext<'_>) -> CalcValue {
    let scalar = match args.first() {
        Some(a) => require_scalar(a),
        None => return CalcValue::error(XlError::Value),
    };

    if let ScalarValue::Blank = scalar {
        return CalcValue::number(0.0);
    }

    if let ScalarValue::Number(n) = scalar {
        return CalcValue::number(*n);
    }

    let text = match scalar {
        ScalarValue::Text(s) => s.clone(),
        ScalarValue::Error(e) => return CalcValue::error(*e),
        _ => return CalcValue::error(XlError::Value),
    };

    // Try parsing with percent
    let is_percent = text.contains('%');
    let text_clean = if is_percent {
        text.replace('%', "")
    } else {
        text
    };

    match text_clean.trim().parse::<f64>() {
        Ok(n) => {
            let result = if is_percent { n / 100.0 } else { n };
            CalcValue::number(result)
        }
        Err(_) => CalcValue::error(XlError::Value),
    }
}

/// Registers all text functions into the given registry.
pub fn register(registry: &mut FunctionRegistry) {
    registry.register(FunctionDef {
        name: "LEFT",
        min_args: 1,
        max_args: 2,
        param_kinds: &[ParamKind::Scalar, ParamKind::Scalar],
        func: fn_left,
    });
    registry.register(FunctionDef {
        name: "RIGHT",
        min_args: 1,
        max_args: 2,
        param_kinds: &[ParamKind::Scalar, ParamKind::Scalar],
        func: fn_right,
    });
    registry.register(FunctionDef {
        name: "MID",
        min_args: 3,
        max_args: 3,
        param_kinds: &[ParamKind::Scalar, ParamKind::Scalar, ParamKind::Scalar],
        func: fn_mid,
    });
    registry.register(FunctionDef {
        name: "LEN",
        min_args: 1,
        max_args: 1,
        param_kinds: &[ParamKind::Scalar],
        func: fn_len,
    });
    registry.register(FunctionDef {
        name: "TRIM",
        min_args: 1,
        max_args: 1,
        param_kinds: &[ParamKind::Scalar],
        func: fn_trim,
    });
    registry.register(FunctionDef {
        name: "UPPER",
        min_args: 1,
        max_args: 1,
        param_kinds: &[ParamKind::Scalar],
        func: fn_upper,
    });
    registry.register(FunctionDef {
        name: "LOWER",
        min_args: 1,
        max_args: 1,
        param_kinds: &[ParamKind::Scalar],
        func: fn_lower,
    });
    registry.register(FunctionDef {
        name: "CONCATENATE",
        min_args: 1,
        max_args: usize::MAX,
        param_kinds: &[ParamKind::Scalar],
        func: fn_concatenate,
    });
    registry.register(FunctionDef {
        name: "ASC",
        min_args: 1,
        max_args: 1,
        param_kinds: &[ParamKind::Scalar],
        func: fn_asc,
    });
    registry.register(FunctionDef {
        name: "CHAR",
        min_args: 1,
        max_args: 1,
        param_kinds: &[ParamKind::Scalar],
        func: fn_char,
    });
    registry.register(FunctionDef {
        name: "CLEAN",
        min_args: 1,
        max_args: 1,
        param_kinds: &[ParamKind::Scalar],
        func: fn_clean,
    });
    registry.register(FunctionDef {
        name: "CODE",
        min_args: 1,
        max_args: 1,
        param_kinds: &[ParamKind::Scalar],
        func: fn_code,
    });
    registry.register(FunctionDef {
        name: "CONCAT",
        min_args: 1,
        max_args: usize::MAX,
        param_kinds: &[ParamKind::Scalar],
        func: fn_concat,
    });
    registry.register(FunctionDef {
        name: "DOLLAR",
        min_args: 1,
        max_args: 2,
        param_kinds: &[ParamKind::Scalar, ParamKind::Scalar],
        func: fn_dollar,
    });
    registry.register(FunctionDef {
        name: "EXACT",
        min_args: 2,
        max_args: 2,
        param_kinds: &[ParamKind::Scalar, ParamKind::Scalar],
        func: fn_exact,
    });
    registry.register(FunctionDef {
        name: "FIND",
        min_args: 2,
        max_args: 3,
        param_kinds: &[ParamKind::Scalar, ParamKind::Scalar, ParamKind::Scalar],
        func: fn_find,
    });
    registry.register(FunctionDef {
        name: "FIXED",
        min_args: 1,
        max_args: 3,
        param_kinds: &[ParamKind::Scalar, ParamKind::Scalar, ParamKind::Scalar],
        func: fn_fixed,
    });
    registry.register(FunctionDef {
        name: "LEFTB",
        min_args: 1,
        max_args: 2,
        param_kinds: &[ParamKind::Scalar, ParamKind::Scalar],
        func: fn_leftb,
    });
    registry.register(FunctionDef {
        name: "NUMBERVALUE",
        min_args: 1,
        max_args: 3,
        param_kinds: &[ParamKind::Scalar, ParamKind::Scalar, ParamKind::Scalar],
        func: fn_numbervalue,
    });
    registry.register(FunctionDef {
        name: "PROPER",
        min_args: 1,
        max_args: 1,
        param_kinds: &[ParamKind::Scalar],
        func: fn_proper,
    });
    registry.register(FunctionDef {
        name: "REPLACE",
        min_args: 4,
        max_args: 4,
        param_kinds: &[
            ParamKind::Scalar,
            ParamKind::Scalar,
            ParamKind::Scalar,
            ParamKind::Scalar,
        ],
        func: fn_replace,
    });
    registry.register(FunctionDef {
        name: "REPT",
        min_args: 2,
        max_args: 2,
        param_kinds: &[ParamKind::Scalar, ParamKind::Scalar],
        func: fn_rept,
    });
    registry.register(FunctionDef {
        name: "SEARCH",
        min_args: 2,
        max_args: 3,
        param_kinds: &[ParamKind::Scalar, ParamKind::Scalar, ParamKind::Scalar],
        func: fn_search,
    });
    registry.register(FunctionDef {
        name: "SUBSTITUTE",
        min_args: 3,
        max_args: 4,
        param_kinds: &[
            ParamKind::Scalar,
            ParamKind::Scalar,
            ParamKind::Scalar,
            ParamKind::Scalar,
        ],
        func: fn_substitute,
    });
    registry.register(FunctionDef {
        name: "T",
        min_args: 1,
        max_args: 1,
        param_kinds: &[ParamKind::Scalar],
        func: fn_t,
    });
    registry.register(FunctionDef {
        name: "TEXT",
        min_args: 2,
        max_args: 2,
        param_kinds: &[ParamKind::Scalar, ParamKind::Scalar],
        func: fn_text,
    });
    registry.register(FunctionDef {
        name: "TEXTJOIN",
        min_args: 3,
        max_args: usize::MAX,
        param_kinds: &[ParamKind::Scalar, ParamKind::Scalar, ParamKind::Scalar],
        func: fn_textjoin,
    });
    registry.register(FunctionDef {
        name: "VALUE",
        min_args: 1,
        max_args: 1,
        param_kinds: &[ParamKind::Scalar],
        func: fn_value,
    });
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::context::{CellDataProvider, EvalContext};
    use pretty_assertions::assert_eq;

    struct TestProvider;
    impl CellDataProvider for TestProvider {
        fn cell_value(&self, _sheet: Option<&str>, _row: u32, _col: u32) -> ScalarValue {
            ScalarValue::Blank
        }
    }

    fn test_ctx() -> EvalContext<'static> {
        static PROVIDER: TestProvider = TestProvider;
        EvalContext::new(&PROVIDER, None::<String>, 1, 1)
    }

    fn arg(val: CalcValue) -> FunctionArg<'static> {
        FunctionArg::Value(val)
    }

    #[test]
    fn test_char() {
        let ctx = test_ctx();
        assert_eq!(
            fn_char(&[arg(CalcValue::number(65.0))], &ctx),
            CalcValue::text("A")
        );
        assert_eq!(
            fn_char(&[arg(CalcValue::number(97.0))], &ctx),
            CalcValue::text("a")
        );
    }

    #[test]
    fn test_code() {
        let ctx = test_ctx();
        assert_eq!(
            fn_code(&[arg(CalcValue::text("A"))], &ctx),
            CalcValue::number(65.0)
        );
        assert_eq!(
            fn_code(&[arg(CalcValue::text("a"))], &ctx),
            CalcValue::number(97.0)
        );
    }

    #[test]
    fn test_clean() {
        let ctx = test_ctx();
        let input = "Hello\x01World\x1F!";
        assert_eq!(
            fn_clean(&[arg(CalcValue::text(input))], &ctx),
            CalcValue::text("HelloWorld!")
        );
    }

    #[test]
    fn test_exact() {
        let ctx = test_ctx();
        assert_eq!(
            fn_exact(
                &[arg(CalcValue::text("ABC")), arg(CalcValue::text("ABC"))],
                &ctx
            ),
            CalcValue::bool(true)
        );
        assert_eq!(
            fn_exact(
                &[arg(CalcValue::text("ABC")), arg(CalcValue::text("abc"))],
                &ctx
            ),
            CalcValue::bool(false)
        );
    }

    #[test]
    fn test_find() {
        let ctx = test_ctx();
        assert_eq!(
            fn_find(
                &[arg(CalcValue::text("o")), arg(CalcValue::text("Hello"))],
                &ctx
            ),
            CalcValue::number(5.0)
        );
    }

    #[test]
    fn test_proper() {
        let ctx = test_ctx();
        assert_eq!(
            fn_proper(&[arg(CalcValue::text("hello world"))], &ctx),
            CalcValue::text("Hello World")
        );
    }

    #[test]
    fn test_replace() {
        let ctx = test_ctx();
        assert_eq!(
            fn_replace(
                &[
                    arg(CalcValue::text("abcdef")),
                    arg(CalcValue::number(2.0)),
                    arg(CalcValue::number(3.0)),
                    arg(CalcValue::text("XYZ"))
                ],
                &ctx
            ),
            CalcValue::text("aXYZef")
        );
    }

    #[test]
    fn test_rept() {
        let ctx = test_ctx();
        assert_eq!(
            fn_rept(
                &[arg(CalcValue::text("x")), arg(CalcValue::number(5.0))],
                &ctx
            ),
            CalcValue::text("xxxxx")
        );
    }

    #[test]
    fn test_substitute() {
        let ctx = test_ctx();
        assert_eq!(
            fn_substitute(
                &[
                    arg(CalcValue::text("hello hello")),
                    arg(CalcValue::text("hello")),
                    arg(CalcValue::text("hi"))
                ],
                &ctx
            ),
            CalcValue::text("hi hi")
        );
    }

    #[test]
    fn test_value() {
        let ctx = test_ctx();
        assert_eq!(
            fn_value(&[arg(CalcValue::text("123"))], &ctx),
            CalcValue::number(123.0)
        );
        assert_eq!(
            fn_value(&[arg(CalcValue::text("50%"))], &ctx),
            CalcValue::number(0.5)
        );
    }

    #[test]
    fn test_textjoin() {
        let ctx = test_ctx();
        assert_eq!(
            fn_textjoin(
                &[
                    arg(CalcValue::text(",")),
                    arg(CalcValue::bool(false)),
                    arg(CalcValue::text("a")),
                    arg(CalcValue::text("b")),
                    arg(CalcValue::text("c"))
                ],
                &ctx
            ),
            CalcValue::text("a,b,c")
        );
    }
}
