//! Number format parsing and application
//!
//! Excel number format codes are a mini-language. We support common formats
//! and fall back to raw values for unknown formats.
//!
//! Format codes can have up to 4 sections separated by semicolons:
//! - `positive;negative;zero;text`
//! - `positive;negative` (zero uses positive, text uses General)
//! - `positive` (all values use this)
//!
//! Format codes can include:
//! - Color specifications: `[Red]`, `[Blue]`, `[Color1]` through `[Color56]`
//! - Conditions: `[>100]`, `[<=50]`
//! - Accounting alignment: `_` (skip width), `*` (repeat fill)

/// Result of formatting a value, including optional color
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct FormattedValue {
    /// The formatted text
    pub text: String,
    /// Optional color in #RRGGBB format
    pub color: Option<String>,
}

impl FormattedValue {
    /// Create a new FormattedValue with just text
    pub fn new(text: String) -> Self {
        Self { text, color: None }
    }

    /// Create a new FormattedValue with text and color
    pub fn with_color(text: String, color: String) -> Self {
        Self {
            text,
            color: Some(color),
        }
    }
}

/// Parsed format sections from a format code
#[derive(Debug, Clone, Default)]
pub struct FormatSections {
    /// Format for positive numbers (or all numbers if only one section)
    pub positive: String,
    /// Format for negative numbers
    pub negative: Option<String>,
    /// Format for zero values
    pub zero: Option<String>,
    /// Format for text values
    pub text: Option<String>,
}

/// A condition parsed from a format code (e.g., `[>100]`)
#[derive(Debug, Clone, PartialEq)]
pub struct FormatCondition {
    /// The comparison operator
    pub operator: ConditionOperator,
    /// The value to compare against
    pub value: f64,
}

/// Comparison operators for format conditions
#[derive(Debug, Clone, Copy, PartialEq)]
pub enum ConditionOperator {
    /// `<` comparison.
    LessThan,
    /// `<=` comparison.
    LessOrEqual,
    /// `>` comparison.
    GreaterThan,
    /// `>=` comparison.
    GreaterOrEqual,
    /// `=` comparison.
    Equal,
    /// `<>` comparison.
    NotEqual,
}

impl FormatCondition {
    /// Check if a value matches this condition
    pub fn matches(&self, value: f64) -> bool {
        match self.operator {
            ConditionOperator::LessThan => value < self.value,
            ConditionOperator::LessOrEqual => value <= self.value,
            ConditionOperator::GreaterThan => value > self.value,
            ConditionOperator::GreaterOrEqual => value >= self.value,
            ConditionOperator::Equal => (value - self.value).abs() < f64::EPSILON,
            ConditionOperator::NotEqual => (value - self.value).abs() >= f64::EPSILON,
        }
    }
}

/// Parsed information from a single format section
#[derive(Debug, Clone, Default)]
#[allow(dead_code)]
struct ParsedSection {
    /// The format code with color/condition codes removed
    format: String,
    /// Optional color specification
    color: Option<String>,
    /// Optional condition
    condition: Option<FormatCondition>,
}

/// Built-in number format IDs (0-49 are predefined by Excel)
/// See: ECMA-376 Part 1, Section 18.8.30
pub const fn get_builtin_format(id: u32) -> Option<&'static str> {
    match id {
        0 => Some("General"),
        1 => Some("0"),
        2 => Some("0.00"),
        3 => Some("#,##0"),
        4 => Some("#,##0.00"),
        // Currency formats (5-8)
        5 => Some("$#,##0_);($#,##0)"),
        6 => Some("$#,##0_);[Red]($#,##0)"),
        7 => Some("$#,##0.00_);($#,##0.00)"),
        8 => Some("$#,##0.00_);[Red]($#,##0.00)"),
        9 => Some("0%"),
        10 => Some("0.00%"),
        11 => Some("0.00E+00"),
        12 => Some("# ?/?"),
        13 => Some("# ??/??"),
        14 => Some("mm-dd-yy"),
        15 => Some("d-mmm-yy"),
        16 => Some("d-mmm"),
        17 => Some("mmm-yy"),
        18 => Some("h:mm AM/PM"),
        19 => Some("h:mm:ss AM/PM"),
        20 => Some("h:mm"),
        21 => Some("h:mm:ss"),
        22 => Some("m/d/yy h:mm"),
        37 => Some("#,##0 ;(#,##0)"),
        38 => Some("#,##0 ;[Red](#,##0)"),
        39 => Some("#,##0.00;(#,##0.00)"),
        40 => Some("#,##0.00;[Red](#,##0.00)"),
        // Accounting formats (41-44)
        41 => Some("_(* #,##0_);_(* (#,##0);_(* \"-\"_);_(@_)"),
        42 => Some("_($* #,##0_);_($* (#,##0);_($* \"-\"_);_(@_)"),
        43 => Some("_(* #,##0.00_);_(* (#,##0.00);_(* \"-\"??_);_(@_)"),
        44 => Some("_($* #,##0.00_);_($* (#,##0.00);_($* \"-\"??_);_(@_)"),
        45 => Some("mm:ss"),
        46 => Some("[h]:mm:ss"),
        47 => Some("mmss.0"),
        48 => Some("##0.0E+0"),
        49 => Some("@"),
        _ => None,
    }
}

/// Check if a format code is a date/time format
pub fn is_date_format(format_code: &str) -> bool {
    let lower = format_code.to_lowercase();

    // Skip text in quotes and brackets
    let mut in_quotes = false;
    let mut in_brackets = false;
    let mut cleaned = String::new();

    for c in lower.chars() {
        match c {
            '"' => in_quotes = !in_quotes,
            '[' => in_brackets = true,
            ']' => in_brackets = false,
            _ if !in_quotes && !in_brackets => cleaned.push(c),
            _ => {}
        }
    }

    // Check for date/time tokens
    cleaned.contains('y') ||
    cleaned.contains('m') && !cleaned.contains('#') ||  // m is month if no # (number format)
    cleaned.contains('d') ||
    cleaned.contains('h') ||
    cleaned.contains('s') && cleaned.contains(':') // s is seconds if with colon
}

/// Check if a format code is a scientific notation format
pub fn is_scientific_format(format_code: &str) -> bool {
    let upper = format_code.to_uppercase();
    upper.contains("E+") || upper.contains("E-")
}

/// Check if a format code is a fraction format
pub fn is_fraction_format(format_code: &str) -> bool {
    // Fraction formats contain ? or digits around a /
    // Examples: "# ?/?", "# ??/??", "# ?/8", "# ?/16"
    format_code.contains('/')
        && (format_code.contains('?') || {
            // Check for fixed denominator format like "# ?/8"
            let parts: Vec<&str> = format_code.split('/').collect();
            parts.len() == 2
                && parts
                    .get(1)
                    .is_some_and(|p| p.trim().chars().all(|c| c.is_ascii_digit()))
        })
}

/// Pre-compiled number format for efficient repeated formatting.
#[derive(Debug, Clone)]
pub enum CompiledFormat {
    /// General (default) display format.
    General,
    /// Scientific notation format (e.g. `0.00E+00`).
    Scientific(String),
    /// Fraction format (e.g. `# ?/?`).
    Fraction(String),
    /// Date/time format.
    Date(DateFormat),
    /// Numeric format with decimals, thousands, percent, and/or currency.
    Numeric(NumericFormat),
}

/// Compiled date format with pre-parsed tokens.
#[derive(Debug, Clone)]
pub struct DateFormat {
    tokens: Vec<DateToken>,
    has_ampm: bool,
}

/// Compiled numeric format with pre-parsed display options.
#[derive(Debug, Clone)]
pub struct NumericFormat {
    decimals: usize,
    has_thousands: bool,
    percent: bool,
    currency: Option<char>,
}

/// Compile a format code string into a [`CompiledFormat`] for repeated use.
pub fn compile_format_code(format_code: &str) -> CompiledFormat {
    let code = format_code.trim();

    if code.eq_ignore_ascii_case("General") || code == "@" {
        return CompiledFormat::General;
    }

    if is_scientific_format(code) {
        return CompiledFormat::Scientific(code.to_string());
    }

    if is_fraction_format(code) {
        return CompiledFormat::Fraction(code.to_string());
    }

    if is_date_format(code) {
        let code_lower = code.to_lowercase();
        let has_ampm = code_lower.contains("am/pm") || code_lower.contains("a/p");
        let tokens = parse_date_format_tokens(code);
        return CompiledFormat::Date(DateFormat { tokens, has_ampm });
    }

    let percent = code.contains('%');
    let has_thousands = code.contains(',');
    let decimals = if percent {
        code.matches('0').count().saturating_sub(1)
    } else {
        code.find('.')
            .map_or(0, |pos| code[pos..].matches('0').count())
    };
    let currency = if code.contains('$') {
        Some('$')
    } else if code.contains('€') {
        Some('€')
    } else if code.contains('£') {
        Some('£')
    } else {
        None
    };

    CompiledFormat::Numeric(NumericFormat {
        decimals,
        has_thousands,
        percent,
        currency,
    })
}

/// Format a numeric value using a pre-compiled format.
pub fn format_number_compiled(value: f64, compiled: &CompiledFormat, date1904: bool) -> String {
    match compiled {
        CompiledFormat::General => format_general(value),
        CompiledFormat::Scientific(code) => format_scientific(value, code),
        CompiledFormat::Fraction(code) => format_fraction(value, code),
        CompiledFormat::Date(fmt) => format_date_compiled(value, fmt, date1904),
        CompiledFormat::Numeric(fmt) => format_numeric_compiled(value, fmt),
    }
}

/// Format a numeric value using a format code
/// Returns the formatted string
pub fn format_number(value: f64, format_code: &str, date1904: bool) -> String {
    let code = format_code.trim();

    // Handle General format
    if code.eq_ignore_ascii_case("General") || code == "@" {
        return format_general(value);
    }

    // Check for scientific notation format
    if is_scientific_format(code) {
        return format_scientific(value, code);
    }

    // Check for fraction format
    if is_fraction_format(code) {
        return format_fraction(value, code);
    }

    // Check if it's a date format
    if is_date_format(code) {
        return format_date(value, code, date1904);
    }

    // Try to parse and apply numeric format
    format_numeric(value, code)
}

/// General format - smart number display
#[allow(clippy::float_cmp)]
#[allow(clippy::cast_possible_truncation)]
fn format_general(value: f64) -> String {
    if value == value.floor() && value.abs() < 1e11 {
        // Integer display
        format!("{}", value as i64)
    } else if value.abs() >= 1e11 || (value.abs() < 1e-4 && value != 0.0) {
        // Scientific notation for very large/small
        format!("{value:.5E}")
    } else {
        // Standard decimal, trim trailing zeros
        let s = format!("{value:.10}");
        let s = s.trim_end_matches('0');
        let s = s.trim_end_matches('.');
        s.to_string()
    }
}

/// Format a number in scientific notation
/// Supports formats like "0.00E+00", "0.00E-00", "##0.0E+0"
#[allow(clippy::cast_possible_truncation)]
#[allow(clippy::cast_sign_loss)]
fn format_scientific(value: f64, format_code: &str) -> String {
    let upper = format_code.to_uppercase();

    // Determine if we always show the sign (E+) or only for negative (E-)
    let always_show_sign = upper.contains("E+");

    // Find the position of E in the format
    let e_pos = upper.find('E').unwrap_or(format_code.len());
    let mantissa_part = &format_code[..e_pos];
    let exponent_part = &format_code[e_pos..];

    // Count decimal places in mantissa (count 0s and ?s after decimal point)
    let mantissa_decimals = mantissa_part.find('.').map_or(0, |pos| {
        mantissa_part[pos..]
            .chars()
            .filter(|&c| c == '0' || c == '?')
            .count()
    });

    // Count digits in exponent part (determines minimum exponent width)
    // Excel always uses at least 2 digits for the exponent
    let exponent_width = exponent_part
        .chars()
        .filter(|&c| c == '0' || c == '#')
        .count()
        .max(2);

    // Handle special case of zero
    if value == 0.0 {
        let exp_sign = if always_show_sign { "+" } else { "" };
        let exp_zeros = "0".repeat(exponent_width);
        if mantissa_decimals > 0 {
            let zeros = "0".repeat(mantissa_decimals);
            return format!("0.{zeros}E{exp_sign}{exp_zeros}");
        }
        return format!("0E{exp_sign}{exp_zeros}");
    }

    let is_negative = value < 0.0;
    let abs_value = value.abs();

    // Calculate the exponent
    let exponent = abs_value.log10().floor() as i32;
    let mantissa = abs_value / 10_f64.powi(exponent);

    // Format the mantissa with the specified decimal places
    let mantissa_str = format!("{:.prec$}", mantissa, prec = mantissa_decimals);

    // Format the exponent with sign and padding
    let exp_sign = if exponent >= 0 {
        if always_show_sign {
            "+"
        } else {
            ""
        }
    } else {
        "-"
    };
    let exp_abs = exponent.unsigned_abs() as usize;
    let exp_str = format!("{:0>width$}", exp_abs, width = exponent_width);

    let result = format!("{mantissa_str}E{exp_sign}{exp_str}");

    if is_negative {
        format!("-{result}")
    } else {
        result
    }
}

/// Find the best fraction approximation using the continued fraction algorithm
/// Returns (numerator, denominator) for the fractional part
///
/// # Safety of casts
/// The algorithm uses i64 for intermediate values, then converts to i32/u32 for return.
/// - Denominator: Always <= max_denom (a u32), so i64→u32 is safe
/// - Numerator: For fractions in [0,1] with denominator d, numerator <= d, so also safe
#[allow(clippy::cast_possible_truncation, clippy::cast_sign_loss)]
fn continued_fraction_approximation(x: f64, max_denom: u32) -> (i32, u32) {
    if x == 0.0 {
        return (0, 1);
    }

    let is_negative = x < 0.0;
    let x = x.abs();

    // Use the Stern-Brocot tree / mediant method for best approximation
    let mut lower_n: i64 = 0;
    let mut lower_d: i64 = 1;
    let mut upper_n: i64 = 1;
    let mut upper_d: i64 = 1;

    let max_denom = max_denom as i64;

    loop {
        let mid_n = lower_n + upper_n;
        let mid_d = lower_d + upper_d;

        if mid_d > max_denom {
            break;
        }

        let mid_val = mid_n as f64 / mid_d as f64;

        if (mid_val - x).abs() < 1e-12 {
            // Exact match
            let num = if is_negative {
                -(mid_n as i32)
            } else {
                mid_n as i32
            };
            return (num, mid_d as u32);
        }

        if mid_val < x {
            lower_n = mid_n;
            lower_d = mid_d;
        } else {
            upper_n = mid_n;
            upper_d = mid_d;
        }
    }

    // Choose the closer approximation between lower and upper bounds
    let lower_err = (lower_n as f64 / lower_d as f64 - x).abs();
    let upper_err = (upper_n as f64 / upper_d as f64 - x).abs();

    let (best_n, best_d) = if lower_err <= upper_err {
        (lower_n, lower_d)
    } else {
        (upper_n, upper_d)
    };

    let num = if is_negative {
        -(best_n as i32)
    } else {
        best_n as i32
    };
    (num, best_d as u32)
}

/// Convert a decimal value to a fraction
/// Returns (whole_part, numerator, denominator)
#[allow(clippy::cast_possible_truncation)]
fn to_fraction(value: f64, max_denominator: u32) -> (i32, i32, u32) {
    if value.is_nan() || value.is_infinite() {
        return (0, 0, 1);
    }

    let is_negative = value < 0.0;
    let abs_value = value.abs();

    // Extract whole part
    let whole_part = abs_value.floor() as i32;
    let fractional_part = abs_value.fract();

    // Handle case where there's no fractional part
    if fractional_part < 1e-10 {
        let signed_whole = if is_negative { -whole_part } else { whole_part };
        return (signed_whole, 0, 1);
    }

    // Find the best fraction approximation for the fractional part
    let (mut numerator, denominator) =
        continued_fraction_approximation(fractional_part, max_denominator);

    // Handle case where fraction rounds to 1 (e.g., 0.9999... with small max_denom)
    // numerator is always non-negative here (fraction of absolute value)
    if u32::try_from(numerator).is_ok_and(|n| n == denominator) {
        let signed_whole = if is_negative {
            -(whole_part + 1)
        } else {
            whole_part + 1
        };
        return (signed_whole, 0, 1);
    }

    let signed_whole = if is_negative { -whole_part } else { whole_part };
    if is_negative && numerator > 0 {
        numerator = -numerator;
    }

    (signed_whole, numerator, denominator)
}

/// Convert a value to a fraction with a fixed denominator
/// Returns (whole_part, numerator, denominator)
#[allow(clippy::cast_possible_truncation)]
fn to_fraction_fixed_denom(value: f64, fixed_denom: u32) -> (i32, i32, u32) {
    if value.is_nan() || value.is_infinite() || fixed_denom == 0 {
        return (0, 0, 1);
    }

    let is_negative = value < 0.0;
    let abs_value = value.abs();

    // Extract whole part
    let whole_part = abs_value.floor() as i32;
    let fractional_part = abs_value.fract();

    // Calculate numerator by rounding
    let numerator = (fractional_part * fixed_denom as f64).round() as i32;

    // Handle case where numerator equals denominator
    // fixed_denom is typically small (like 2, 4, 8, 16) so this comparison is safe
    if i32::try_from(fixed_denom).is_ok_and(|d| numerator == d) {
        let signed_whole = if is_negative {
            -(whole_part + 1)
        } else {
            whole_part + 1
        };
        return (signed_whole, 0, fixed_denom);
    }

    let signed_whole = if is_negative { -whole_part } else { whole_part };
    let signed_num = if is_negative && numerator > 0 {
        -numerator
    } else {
        numerator
    };

    (signed_whole, signed_num, fixed_denom)
}

/// Parse fraction format to determine max denominator or fixed denominator
/// Returns (max_denominator, fixed_denominator_option)
fn parse_fraction_format(format_code: &str) -> (u32, Option<u32>) {
    // Split by '/'
    let parts: Vec<&str> = format_code.split('/').collect();
    let Some(denom_part) = parts.get(1).map(|s| s.trim()) else {
        return (9, None); // Default to single digit
    };

    // Check if it's a fixed denominator (all digits)
    if denom_part.chars().all(|c| c.is_ascii_digit()) && !denom_part.is_empty() {
        if let Ok(fixed) = denom_part.parse::<u32>() {
            if fixed > 0 {
                return (fixed, Some(fixed));
            }
        }
    }

    // Count the number of ? characters to determine max denominator digits
    let question_marks = denom_part.chars().filter(|&c| c == '?').count();

    let max_denom = match question_marks {
        0 => 9,    // Default
        1 => 9,    // Single digit: 1-9
        2 => 99,   // Two digits: 1-99
        3 => 999,  // Three digits: 1-999
        _ => 9999, // Four or more: 1-9999
    };

    (max_denom, None)
}

/// Format a number as a fraction
/// Supports formats like "# ?/?", "# ??/??", "# ?/8", "# ?/16"
fn format_fraction(value: f64, format_code: &str) -> String {
    // Handle special cases
    if value.is_nan() {
        return "NaN".to_string();
    }
    if value.is_infinite() {
        return if value > 0.0 { "Inf" } else { "-Inf" }.to_string();
    }

    let (max_denom, fixed_denom) = parse_fraction_format(format_code);

    let (whole, numerator, denominator) = if let Some(fixed) = fixed_denom {
        to_fraction_fixed_denom(value, fixed)
    } else {
        to_fraction(value, max_denom)
    };

    // Determine the width for numerator and denominator based on format
    let parts: Vec<&str> = format_code.split('/').collect();
    let num_part = parts.first().copied().unwrap_or("");
    // Check if format has a whole number placeholder (# before the fraction)
    let _has_whole_placeholder = num_part.contains('#');
    let denom_width = parts.get(1).map_or(1, |denom_part| {
        let denom_part = denom_part.trim();
        if denom_part.chars().all(|c| c.is_ascii_digit()) {
            denom_part.len()
        } else {
            denom_part.chars().filter(|&c| c == '?').count().max(1)
        }
    });
    // For fixed denominator formats, numerator width should match denominator width
    // For variable denominator formats, use the count of ? placeholders
    let num_width = if fixed_denom.is_some() {
        denom_width
    } else {
        num_part.chars().filter(|&c| c == '?').count().max(1)
    };

    // Build the result string
    let is_negative = value < 0.0;
    let abs_whole = whole.abs();
    let abs_num = numerator.abs();

    if abs_num == 0 {
        // No fractional part, just show the whole number
        if abs_whole == 0 {
            return "0".to_string();
        }
        return format!("{}", if is_negative { -abs_whole } else { abs_whole });
    }

    // Format numerator and denominator with proper spacing
    let num_str = format!("{:>width$}", abs_num, width = num_width);
    let denom_str = format!("{:<width$}", denominator, width = denom_width);

    if abs_whole == 0 {
        // Only fractional part
        let sign = if is_negative { "-" } else { "" };
        format!("{sign}{num_str}/{denom_str}")
    } else {
        // Whole part and fractional part
        let sign = if is_negative { "-" } else { "" };
        format!("{sign}{abs_whole} {num_str}/{denom_str}")
    }
}

/// Format a date value (Excel serial date)
///
/// # Cast safety
/// The casts in this function are safe because:
/// - Excel dates are typically in range 1-2958465 (years 1900-9999)
/// - Elapsed hours/minutes/seconds are bounded by practical spreadsheet limits
#[allow(
    clippy::cast_possible_truncation,
    clippy::cast_sign_loss,
    clippy::cast_possible_wrap
)]
fn format_date(value: f64, format_code: &str, date1904: bool) -> String {
    // Excel dates: days since 1899-12-30 (1900 system) or 1904-01-01 (1904 system)
    let (year, month, day, hour, minute, second) = excel_date_to_components(value, date1904);

    // Calculate day of week (0 = Sunday, 1 = Monday, ... 6 = Saturday)
    // Jan 1, 1900 was a Monday (day 1), but we need to account for the leap year bug
    let days = value.floor() as i32;
    let day_of_week = ((days + 6) % 7) as u32; // Adjust so day 1 (Jan 1, 1900) is Monday (1)

    // Check if AM/PM is present in the format (case insensitive)
    let code_lower = format_code.to_lowercase();
    let has_ampm = code_lower.contains("am/pm") || code_lower.contains("a/p");

    // Calculate display hour for 12-hour format
    let display_hour = if has_ampm {
        match hour {
            0 => 12,
            1..=12 => hour,
            _ => hour - 12,
        }
    } else {
        hour
    };

    let ampm_str = if hour >= 12 { "PM" } else { "AM" };
    let ap_str = if hour >= 12 { "P" } else { "A" };

    // Parse the format code token by token
    let tokens = parse_date_format_tokens(format_code);

    let mut result = String::new();

    for token in tokens {
        match token {
            DateToken::Year4 => result.push_str(&format!("{year:04}")),
            DateToken::Year2 => result.push_str(&format!("{:02}", year % 100)),
            DateToken::Month1 => result.push_str(&format!("{month}")),
            DateToken::Month2 => result.push_str(&format!("{month:02}")),
            DateToken::Month3 => result.push_str(month_abbrev(month)),
            DateToken::Month4 => result.push_str(month_full(month)),
            DateToken::Month5 => result.push_str(month_letter(month)),
            DateToken::Day1 => result.push_str(&format!("{day}")),
            DateToken::Day2 => result.push_str(&format!("{day:02}")),
            DateToken::Day3 => result.push_str(day_abbrev(day_of_week)),
            DateToken::Day4 => result.push_str(day_full(day_of_week)),
            DateToken::Hour1 => {
                result.push_str(&format!("{}", if has_ampm { display_hour } else { hour }))
            }
            DateToken::Hour2 => result.push_str(&format!(
                "{:02}",
                if has_ampm { display_hour } else { hour }
            )),
            DateToken::Minute1 => result.push_str(&format!("{minute}")),
            DateToken::Minute2 => result.push_str(&format!("{minute:02}")),
            DateToken::Second1 => result.push_str(&format!("{second}")),
            DateToken::Second2 => result.push_str(&format!("{second:02}")),
            DateToken::AmPm => result.push_str(ampm_str),
            DateToken::AP => result.push_str(ap_str),
            DateToken::Literal(s) => result.push_str(&s),
            DateToken::ElapsedHours => {
                // [h] format - total hours including days
                let total_hours = (value * 24.0).floor() as u32;
                result.push_str(&format!("{total_hours}"));
            }
            DateToken::ElapsedMinutes => {
                // [m] format - total minutes
                let total_minutes = (value * 24.0 * 60.0).floor() as u32;
                result.push_str(&format!("{total_minutes}"));
            }
            DateToken::ElapsedSeconds => {
                // [s] format - total seconds
                let total_seconds = (value * 86400.0).floor() as u32;
                result.push_str(&format!("{total_seconds}"));
            }
        }
    }

    if result.is_empty() {
        // Fallback: ISO-ish format
        format!("{year:04}-{month:02}-{day:02}")
    } else {
        result
    }
}

#[allow(
    clippy::cast_possible_truncation,
    clippy::cast_sign_loss,
    clippy::cast_possible_wrap
)]
fn format_date_compiled(value: f64, fmt: &DateFormat, date1904: bool) -> String {
    let (year, month, day, hour, minute, second) = excel_date_to_components(value, date1904);
    let days = value.floor() as i32;
    let day_of_week = ((days + 6) % 7) as u32;

    let display_hour = if fmt.has_ampm {
        match hour {
            0 => 12,
            1..=12 => hour,
            _ => hour - 12,
        }
    } else {
        hour
    };

    let ampm_str = if hour >= 12 { "PM" } else { "AM" };
    let ap_str = if hour >= 12 { "P" } else { "A" };

    let mut result = String::new();
    for token in &fmt.tokens {
        match token {
            DateToken::Year4 => result.push_str(&format!("{year:04}")),
            DateToken::Year2 => result.push_str(&format!("{:02}", year % 100)),
            DateToken::Month1 => result.push_str(&format!("{month}")),
            DateToken::Month2 => result.push_str(&format!("{month:02}")),
            DateToken::Month3 => result.push_str(month_abbrev(month)),
            DateToken::Month4 => result.push_str(month_full(month)),
            DateToken::Month5 => result.push_str(month_letter(month)),
            DateToken::Day1 => result.push_str(&format!("{day}")),
            DateToken::Day2 => result.push_str(&format!("{day:02}")),
            DateToken::Day3 => result.push_str(day_abbrev(day_of_week)),
            DateToken::Day4 => result.push_str(day_full(day_of_week)),
            DateToken::Hour1 => result.push_str(&format!(
                "{}",
                if fmt.has_ampm { display_hour } else { hour }
            )),
            DateToken::Hour2 => result.push_str(&format!(
                "{:02}",
                if fmt.has_ampm { display_hour } else { hour }
            )),
            DateToken::Minute1 => result.push_str(&format!("{minute}")),
            DateToken::Minute2 => result.push_str(&format!("{minute:02}")),
            DateToken::Second1 => result.push_str(&format!("{second}")),
            DateToken::Second2 => result.push_str(&format!("{second:02}")),
            DateToken::AmPm => result.push_str(ampm_str),
            DateToken::AP => result.push_str(ap_str),
            DateToken::Literal(s) => result.push_str(s),
            DateToken::ElapsedHours => {
                let total_hours = (value * 24.0).floor() as u32;
                result.push_str(&format!("{total_hours}"));
            }
            DateToken::ElapsedMinutes => {
                let total_minutes = (value * 24.0 * 60.0).floor() as u32;
                result.push_str(&format!("{total_minutes}"));
            }
            DateToken::ElapsedSeconds => {
                let total_seconds = (value * 86400.0).floor() as u32;
                result.push_str(&format!("{total_seconds}"));
            }
        }
    }

    if result.is_empty() {
        format!("{year:04}-{month:02}-{day:02}")
    } else {
        result
    }
}

/// Token types for date format parsing
#[derive(Debug, Clone)]
enum DateToken {
    Year4,           // yyyy
    Year2,           // yy or y
    Month1,          // m (month, single digit)
    Month2,          // mm (month, zero-padded)
    Month3,          // mmm (abbreviated month name)
    Month4,          // mmmm (full month name)
    Month5,          // mmmmm (first letter of month)
    Day1,            // d (day, single digit)
    Day2,            // dd (day, zero-padded)
    Day3,            // ddd (abbreviated day name)
    Day4,            // dddd (full day name)
    Hour1,           // h (hour, single digit)
    Hour2,           // hh (hour, zero-padded)
    Minute1,         // m (minute, single digit) - context dependent
    Minute2,         // mm (minute, zero-padded) - context dependent
    Second1,         // s (second, single digit)
    Second2,         // ss (second, zero-padded)
    AmPm,            // AM/PM
    AP,              // A/P
    Literal(String), // literal text or separator
    ElapsedHours,    // [h]
    ElapsedMinutes,  // [m]
    ElapsedSeconds,  // [s]
}

/// Parse a date format code into tokens
///
/// # Indexing safety
/// This function uses manual index bounds checking throughout (e.g., `i + 2 < chars.len()`).
/// Converting all indexing to `.get()` would significantly harm readability without benefit.
#[allow(clippy::indexing_slicing)]
fn parse_date_format_tokens(format_code: &str) -> Vec<DateToken> {
    let mut tokens = Vec::new();
    let chars: Vec<char> = format_code.chars().collect();
    let mut i = 0;
    let mut in_time_context = false; // true if we've seen h/H and not yet seen s/S

    while i < chars.len() {
        let c = chars[i];
        let c_lower = c.to_ascii_lowercase();

        // Check for quoted strings
        if c == '"' {
            let mut literal = String::new();
            i += 1;
            while i < chars.len() && chars[i] != '"' {
                literal.push(chars[i]);
                i += 1;
            }
            i += 1; // skip closing quote
            tokens.push(DateToken::Literal(literal));
            continue;
        }

        // Check for escaped character
        if c == '\\' && i + 1 < chars.len() {
            tokens.push(DateToken::Literal(chars[i + 1].to_string()));
            i += 2;
            continue;
        }

        // Check for elapsed time formats [h], [m], [s]
        if c == '[' {
            if i + 2 < chars.len() && chars[i + 2] == ']' {
                let inner = chars[i + 1].to_ascii_lowercase();
                match inner {
                    'h' => {
                        tokens.push(DateToken::ElapsedHours);
                        i += 3;
                        in_time_context = true;
                        continue;
                    }
                    'm' => {
                        tokens.push(DateToken::ElapsedMinutes);
                        i += 3;
                        continue;
                    }
                    's' => {
                        tokens.push(DateToken::ElapsedSeconds);
                        i += 3;
                        in_time_context = false;
                        continue;
                    }
                    _ => {}
                }
            }
            // Skip color codes and conditions like [Red], [>100], etc.
            let mut j = i + 1;
            while j < chars.len() && chars[j] != ']' {
                j += 1;
            }
            i = j + 1;
            continue;
        }

        // Check for AM/PM or am/pm
        if c_lower == 'a' {
            // Check for AM/PM
            if i + 4 < chars.len() {
                let is_ampm = chars[i + 1].eq_ignore_ascii_case(&'m')
                    && chars[i + 2] == '/'
                    && chars[i + 3].eq_ignore_ascii_case(&'p')
                    && chars[i + 4].eq_ignore_ascii_case(&'m');
                if is_ampm {
                    tokens.push(DateToken::AmPm);
                    i += 5;
                    continue;
                }
            }
            // Check for A/P
            if i + 2 < chars.len() {
                let is_ap = chars[i + 1] == '/' && chars[i + 2].eq_ignore_ascii_case(&'p');
                if is_ap {
                    tokens.push(DateToken::AP);
                    i += 3;
                    continue;
                }
            }
            // Just a literal 'a'
            tokens.push(DateToken::Literal(c.to_string()));
            i += 1;
            continue;
        }

        // Count consecutive same characters
        let mut count = 1;
        while i + count < chars.len() && chars[i + count].to_ascii_lowercase() == c_lower {
            count += 1;
        }

        match c_lower {
            'y' => {
                if count >= 4 {
                    tokens.push(DateToken::Year4);
                } else {
                    tokens.push(DateToken::Year2);
                }
                i += count;
            }
            'm' => {
                // m can be month or minute depending on context
                // It's minutes if:
                // - After h/hh (in_time_context is true)
                // - Before s/ss (look ahead)
                let is_minute = in_time_context || is_followed_by_seconds(&chars, i + count);

                if is_minute {
                    if count >= 2 {
                        tokens.push(DateToken::Minute2);
                    } else {
                        tokens.push(DateToken::Minute1);
                    }
                } else {
                    match count {
                        1 => tokens.push(DateToken::Month1),
                        2 => tokens.push(DateToken::Month2),
                        3 => tokens.push(DateToken::Month3),
                        4 => tokens.push(DateToken::Month4),
                        _ => tokens.push(DateToken::Month5), // 5 or more m's
                    }
                }
                i += count;
            }
            'd' => {
                match count {
                    1 => tokens.push(DateToken::Day1),
                    2 => tokens.push(DateToken::Day2),
                    3 => tokens.push(DateToken::Day3),
                    _ => tokens.push(DateToken::Day4), // 4 or more d's
                }
                i += count;
            }
            'h' => {
                in_time_context = true;
                if count >= 2 {
                    tokens.push(DateToken::Hour2);
                } else {
                    tokens.push(DateToken::Hour1);
                }
                i += count;
            }
            's' => {
                in_time_context = false;
                if count >= 2 {
                    tokens.push(DateToken::Second2);
                } else {
                    tokens.push(DateToken::Second1);
                }
                i += count;
            }
            // Skip formatting characters we don't handle
            '_' | '*' => {
                // _ = skip width of next char, * = repeat next char
                if i + 1 < chars.len() {
                    i += 2;
                } else {
                    i += 1;
                }
            }
            // Common separators and literals
            _ => {
                tokens.push(DateToken::Literal(c.to_string()));
                i += 1;
            }
        }
    }

    tokens
}

/// Check if 'm' at position is followed by 's' (making it minutes, not months)
fn is_followed_by_seconds(chars: &[char], start: usize) -> bool {
    let mut i = start;
    while let Some(&ch) = chars.get(i) {
        let c = ch.to_ascii_lowercase();
        match c {
            's' => return true,
            'h' | 'y' | 'd' => return false, // hit another date/time component
            'm' => return false,             // another m without s in between
            _ => i += 1,                     // skip separators and other chars
        }
    }
    false
}

fn month_abbrev(month: u32) -> &'static str {
    match month {
        1 => "Jan",
        2 => "Feb",
        3 => "Mar",
        4 => "Apr",
        5 => "May",
        6 => "Jun",
        7 => "Jul",
        8 => "Aug",
        9 => "Sep",
        10 => "Oct",
        11 => "Nov",
        12 => "Dec",
        _ => "???",
    }
}

fn month_full(month: u32) -> &'static str {
    match month {
        1 => "January",
        2 => "February",
        3 => "March",
        4 => "April",
        5 => "May",
        6 => "June",
        7 => "July",
        8 => "August",
        9 => "September",
        10 => "October",
        11 => "November",
        12 => "December",
        _ => "???",
    }
}

fn month_letter(month: u32) -> &'static str {
    match month {
        1 => "J",
        2 => "F",
        3 => "M",
        4 => "A",
        5 => "M",
        6 => "J",
        7 => "J",
        8 => "A",
        9 => "S",
        10 => "O",
        11 => "N",
        12 => "D",
        _ => "?",
    }
}

fn day_abbrev(day_of_week: u32) -> &'static str {
    match day_of_week {
        0 => "Sun",
        1 => "Mon",
        2 => "Tue",
        3 => "Wed",
        4 => "Thu",
        5 => "Fri",
        6 => "Sat",
        _ => "???",
    }
}

fn day_full(day_of_week: u32) -> &'static str {
    match day_of_week {
        0 => "Sunday",
        1 => "Monday",
        2 => "Tuesday",
        3 => "Wednesday",
        4 => "Thursday",
        5 => "Friday",
        6 => "Saturday",
        _ => "???",
    }
}

/// Convert Excel serial date to (year, month, day, hour, minute, second)
#[allow(clippy::cast_possible_truncation)]
#[allow(clippy::cast_sign_loss)]
fn excel_date_to_components(serial: f64, date1904: bool) -> (i32, u32, u32, u32, u32, u32) {
    let days = serial.floor() as i32;
    let time_frac = serial.fract().abs(); // Use abs for negative dates

    // Convert Excel serial to Julian Day Number
    // Excel 1900 system: serial 1 = Jan 1, 1900 = JDN 2415021
    // Excel 1904 system: serial 0 = Jan 1, 1904 = JDN 2416481
    let jdn = if date1904 {
        // 1904 system: Day 0 = Jan 1, 1904 (no leap year bug)
        days + 2_416_481
    } else {
        // 1900 system: Excel's epoch has serial 1 = Jan 1, 1900
        // But Excel incorrectly thinks 1900 was a leap year (Feb 29, 1900 = serial 60)
        // Days 1-59 are Jan 1 - Feb 28, 1900
        // Day 60 is the fake Feb 29, 1900
        // Days 61+ are March 1, 1900 onward
        // JDN for Dec 31, 1899 is 2415020, so serial 1 (Jan 1, 1900) = 2415021
        if days <= 60 {
            days + 2_415_020
        } else {
            // After the phantom Feb 29, we need to subtract 1 to account for the fake day
            days + 2_415_019
        }
    };

    // Convert Julian Day Number to (year, month, day)
    let (year, month, day_of_month) = jdn_to_ymd(jdn);

    // Time components
    let total_seconds = (time_frac * 86400.0).round() as u32;
    let hour = total_seconds / 3600;
    let minute = (total_seconds % 3600) / 60;
    let second = total_seconds % 60;

    (year, month, day_of_month, hour, minute, second)
}

/// Convert Julian Day Number to (year, month, day) in proleptic Gregorian calendar
#[allow(clippy::cast_possible_truncation)]
#[allow(clippy::cast_sign_loss)]
fn jdn_to_ymd(jdn: i32) -> (i32, u32, u32) {
    // Algorithm from: https://en.wikipedia.org/wiki/Julian_day#Julian_or_Gregorian_calendar_from_Julian_day_number
    // This is the proleptic Gregorian calendar algorithm

    let y = 4716;
    let j = 1401;
    let m = 2;
    let n = 12;
    let r = 4;
    let p = 1461;
    let v = 3;
    let u = 5;
    let s = 153;
    let w = 2;
    let b = 274277;
    let c = -38;

    let jdn_i64 = jdn as i64;

    let f = jdn_i64 + j + (((4 * jdn_i64 + b) / 146097) * 3) / 4 + c;
    let e = r * f + v;
    let g = (e % p) / r;
    let h = u * g + w;

    let day = (h % s) / u + 1;
    let month = ((h / s + m) % n) + 1;
    let year = (e / p) - y + (n + m - month) / n;

    (year as i32, month as u32, day as u32)
}

/// Format a numeric value with a format code
fn format_numeric(value: f64, format_code: &str) -> String {
    // Handle percentage
    if format_code.contains('%') {
        let pct = value * 100.0;
        let decimals = format_code.matches('0').count().saturating_sub(1);
        return format!("{:.prec$}%", pct, prec = decimals.min(10));
    }

    // Handle thousands separator
    let has_thousands = format_code.contains(',');

    // Count decimal places (0s after the decimal point)
    let decimals = format_code
        .find('.')
        .map_or(0, |pos| format_code[pos..].matches('0').count());

    // Format the number
    let formatted = if has_thousands {
        format_with_thousands(value, decimals)
    } else {
        format!("{:.prec$}", value, prec = decimals.min(10))
    };

    // Handle currency prefix/suffix
    let mut result = formatted;

    if format_code.contains('$') {
        result = format!("${result}");
    } else if format_code.contains('€') {
        result = format!("€{result}");
    } else if format_code.contains('£') {
        result = format!("£{result}");
    }

    result
}

fn format_numeric_compiled(value: f64, fmt: &NumericFormat) -> String {
    if let Some(fast) = format_numeric_fast(value, fmt) {
        return fast;
    }

    if fmt.percent {
        let pct = value * 100.0;
        return format!("{:.prec$}%", pct, prec = fmt.decimals.min(10));
    }

    let formatted = if fmt.has_thousands {
        format_with_thousands(value, fmt.decimals)
    } else {
        format!("{:.prec$}", value, prec = fmt.decimals.min(10))
    };

    let mut result = formatted;
    if let Some(currency) = fmt.currency {
        result.insert(0, currency);
    }

    result
}

#[allow(clippy::cast_possible_truncation, clippy::cast_sign_loss)]
fn format_numeric_fast(value: f64, fmt: &NumericFormat) -> Option<String> {
    if !value.is_finite() || fmt.decimals > 2 {
        return None;
    }

    let mut base = value;
    if fmt.percent {
        base *= 100.0;
    }

    let factor = match fmt.decimals {
        0 => 1.0,
        1 => 10.0,
        2 => 100.0,
        _ => return None,
    };

    let scaled = base * factor;
    let rounded = scaled.round();
    if (scaled - rounded).abs() > 1e-9 {
        return None;
    }
    if rounded.abs() > (i64::MAX as f64) {
        return None;
    }

    let scaled_i = rounded as i64;
    let negative = scaled_i < 0;
    let abs_scaled = scaled_i.abs();
    let factor_i = factor as i64;
    let int_part = abs_scaled / factor_i;
    let frac_part = abs_scaled % factor_i;

    let mut int_str = if fmt.has_thousands {
        format_int_with_thousands(int_part)
    } else {
        int_part.to_string()
    };

    if negative {
        int_str.insert(0, '-');
    }

    let mut out = if fmt.decimals == 0 {
        int_str
    } else {
        let width = fmt.decimals;
        let frac_str = format!("{:0width$}", frac_part, width = width);
        format!("{int_str}.{frac_str}")
    };

    if let Some(currency) = fmt.currency {
        out.insert(0, currency);
    }
    if fmt.percent {
        out.push('%');
    }

    Some(out)
}

fn format_int_with_thousands(value: i64) -> String {
    let digits = value.to_string();
    let len = digits.len();
    if len <= 3 {
        return digits;
    }

    let mut out = String::with_capacity(len + (len - 1) / 3);
    let mut first = len % 3;
    if first == 0 {
        first = 3;
    }
    out.push_str(&digits[..first]);
    let mut i = first;
    while i < len {
        out.push(',');
        out.push_str(&digits[i..i + 3]);
        i += 3;
    }
    out
}

/// Format number with thousands separators
fn format_with_thousands(value: f64, decimals: usize) -> String {
    let is_negative = value < 0.0;
    let abs_value = value.abs();

    let formatted = format!("{:.prec$}", abs_value, prec = decimals.min(10));
    let parts: Vec<&str> = formatted.split('.').collect();

    // split() always returns at least one element for a non-empty string
    let int_part = parts.first().copied().unwrap_or("0");
    let dec_part = parts.get(1);

    // Add thousands separators
    let mut with_sep = String::new();
    for (i, c) in int_part.chars().rev().enumerate() {
        if i > 0 && i % 3 == 0 {
            with_sep.push(',');
        }
        with_sep.push(c);
    }
    let int_with_sep: String = with_sep.chars().rev().collect();

    let result = if let Some(dec) = dec_part {
        format!("{int_with_sep}.{dec}")
    } else {
        int_with_sep
    };

    if is_negative {
        format!("-{result}")
    } else {
        result
    }
}

/// Format number in accounting style (with currency symbol alignment)
#[allow(dead_code)]
fn format_accounting(value: f64, format_code: &str) -> String {
    // For now, just format with currency symbol
    // Full implementation would align currency symbols
    format_numeric(value, format_code)
}

/// Format negative numbers with parentheses instead of minus sign
#[allow(dead_code)]
fn format_with_parentheses(value: f64, format_code: &str) -> String {
    if value < 0.0 {
        format!("({})", format_numeric(-value, format_code))
    } else {
        format_numeric(value, format_code)
    }
}

/// Map an Excel color name to its hex value.
fn color_name_to_hex(name: &str) -> Option<String> {
    match name.to_lowercase().as_str() {
        "black" => Some("#000000".to_string()),
        "blue" => Some("#0000FF".to_string()),
        "cyan" => Some("#00FFFF".to_string()),
        "green" => Some("#00FF00".to_string()),
        "magenta" => Some("#FF00FF".to_string()),
        "red" => Some("#FF0000".to_string()),
        "white" => Some("#FFFFFF".to_string()),
        "yellow" => Some("#FFFF00".to_string()),
        _ => {
            // Color1-Color56 mapping (Color1 = black, Color3 = red, etc.)
            if let Some(idx_str) = name.strip_prefix("Color") {
                if let Ok(idx) = idx_str.parse::<u32>() {
                    return match idx {
                        1 => Some("#000000".to_string()),
                        2 => Some("#FFFFFF".to_string()),
                        3 => Some("#FF0000".to_string()),
                        4 => Some("#00FF00".to_string()),
                        5 => Some("#0000FF".to_string()),
                        6 => Some("#FFFF00".to_string()),
                        7 => Some("#FF00FF".to_string()),
                        8 => Some("#00FFFF".to_string()),
                        _ => None,
                    };
                }
            }
            None
        }
    }
}

/// Extract a color code like `[Red]` or `[Color3]` from a format section.
///
/// Returns `(optional_color_hex, cleaned_format_string)`.
fn extract_color_code(section: &str) -> (Option<String>, String) {
    if let Some(bracket_end) = section.find(']') {
        let inside = section.get(1..bracket_end).unwrap_or_default();
        if let Some(hex) = color_name_to_hex(inside) {
            return (
                Some(hex),
                section
                    .get(bracket_end + 1..)
                    .unwrap_or_default()
                    .to_string(),
            );
        }
    }
    (None, section.to_string())
}

/// Split a format code string into sections separated by `;`.
///
/// Excel format codes can have up to 4 sections: positive;negative;zero;text.
fn split_format_sections(format_code: &str) -> Vec<&str> {
    // Simple split - brackets don't contain semicolons in practice
    format_code.split(';').collect()
}

/// Select the appropriate format section based on the value.
///
/// With 1 section: used for all numbers.
/// With 2 sections: first for positive/zero, second for negative.
/// With 3 sections: positive, negative, zero.
/// With 4 sections: positive, negative, zero, text.
fn select_section<'a>(value: f64, sections: &[&'a str]) -> &'a str {
    match sections.len() {
        0 => "General",
        1 => sections.first().copied().unwrap_or("General"),
        2 => {
            if value >= 0.0 {
                sections.first().copied().unwrap_or("General")
            } else {
                sections.get(1).copied().unwrap_or("General")
            }
        }
        _ => {
            if value > 0.0 {
                sections.first().copied().unwrap_or("General")
            } else if value < 0.0 {
                sections.get(1).copied().unwrap_or("General")
            } else {
                sections.get(2).copied().unwrap_or("General")
            }
        }
    }
}

/// Format a numeric value using an Excel format code.
///
/// This is the main entry point for number formatting. It handles section-based
/// formats (positive;negative;zero;text), color codes like `[Red]`, and all
/// standard Excel format types.
pub fn format_value(value: f64, format_code: &str, date1904: bool) -> FormattedValue {
    if format_code.is_empty() || format_code == "General" {
        return FormattedValue {
            text: format_general(value),
            color: None,
        };
    }

    let sections = split_format_sections(format_code);
    let section = select_section(value, &sections);
    let (color, clean_format) = if section.starts_with('[') {
        extract_color_code(section)
    } else {
        (None, section.to_string())
    };

    let format_value_abs = if sections.len() >= 2 && value < 0.0 {
        value.abs()
    } else {
        value
    };

    let text = if clean_format.is_empty() || clean_format == "General" {
        format_general(format_value_abs)
    } else {
        format_number(format_value_abs, &clean_format, date1904)
    };

    FormattedValue { text, color }
}

/// Compile a format code for repeated use.
///
/// Pre-compiling format codes avoids re-parsing on every call, which is
/// important when formatting many cells with the same format.
pub fn compile_format(format_code: &str) -> CompiledFormat {
    compile_format_code(format_code)
}

/// Format a value using a pre-compiled format.
///
/// This is the fast path for repeated formatting with the same format code.
pub fn format_value_compiled(
    value: f64,
    compiled: &CompiledFormat,
    date1904: bool,
) -> FormattedValue {
    let text = format_number_compiled(value, compiled, date1904);
    FormattedValue { text, color: None }
}

#[cfg(test)]
#[allow(
    clippy::unwrap_used,
    clippy::expect_used,
    clippy::indexing_slicing,
    clippy::float_cmp,
    clippy::panic
)]
mod tests {
    use super::*;

    #[test]
    #[allow(clippy::approx_constant)]
    fn test_general_format() {
        assert_eq!(format_general(42.0), "42");
        assert_eq!(format_general(3.14159), "3.14159");
    }

    #[test]
    fn test_percentage() {
        assert_eq!(format_number(0.5, "0%", false), "50%");
        assert_eq!(format_number(0.125, "0.00%", false), "12.50%");
    }

    #[test]
    fn test_thousands() {
        assert_eq!(format_with_thousands(1_234_567.89, 2), "1,234,567.89");
    }

    #[test]
    fn test_date_detection() {
        assert!(is_date_format("yyyy-mm-dd"));
        assert!(is_date_format("m/d/yy"));
        assert!(!is_date_format("#,##0.00"));
    }

    // Scientific notation tests
    #[test]
    fn test_scientific_format_detection() {
        assert!(is_scientific_format("0.00E+00"));
        assert!(is_scientific_format("0.00E-00"));
        assert!(is_scientific_format("##0.0E+0"));
        assert!(!is_scientific_format("#,##0.00"));
        assert!(!is_scientific_format("0%"));
    }

    #[test]
    fn test_scientific_format_basic() {
        assert_eq!(format_scientific(1234567.0, "0.00E+00"), "1.23E+06");
        assert_eq!(format_scientific(0.000123, "0.00E+00"), "1.23E-04");
        assert_eq!(format_scientific(1.0, "0.00E+00"), "1.00E+00");
    }

    #[test]
    fn test_scientific_format_negative() {
        assert_eq!(format_scientific(-1234567.0, "0.00E+00"), "-1.23E+06");
        assert_eq!(format_scientific(-0.000123, "0.00E+00"), "-1.23E-04");
    }

    #[test]
    fn test_scientific_format_zero() {
        assert_eq!(format_scientific(0.0, "0.00E+00"), "0.00E+00");
        // Excel always uses at least 2 digits for the exponent
        assert_eq!(format_scientific(0.0, "0.0E+0"), "0.0E+00");
    }

    #[test]
    fn test_scientific_format_e_minus() {
        // E- only shows sign for negative exponents
        assert_eq!(format_scientific(1234567.0, "0.00E-00"), "1.23E06");
        assert_eq!(format_scientific(0.000123, "0.00E-00"), "1.23E-04");
    }

    #[test]
    fn test_scientific_format_varying_precision() {
        assert_eq!(format_scientific(1234567.0, "0.0E+00"), "1.2E+06");
        assert_eq!(format_scientific(1234567.0, "0.000E+00"), "1.235E+06");
        assert_eq!(format_scientific(1234567.0, "0E+00"), "1E+06");
    }

    // Fraction format tests
    #[test]
    fn test_fraction_format_detection() {
        assert!(is_fraction_format("# ?/?"));
        assert!(is_fraction_format("# ??/??"));
        assert!(is_fraction_format("# ?/8"));
        assert!(is_fraction_format("# ?/16"));
        assert!(!is_fraction_format("#,##0.00"));
        assert!(!is_fraction_format("0%"));
    }

    #[test]
    fn test_continued_fraction_simple() {
        // 0.5 = 1/2
        let (num, denom) = continued_fraction_approximation(0.5, 9);
        assert_eq!(num, 1);
        assert_eq!(denom, 2);

        // 0.25 = 1/4
        let (num, denom) = continued_fraction_approximation(0.25, 9);
        assert_eq!(num, 1);
        assert_eq!(denom, 4);

        // 0.75 = 3/4
        let (num, denom) = continued_fraction_approximation(0.75, 9);
        assert_eq!(num, 3);
        assert_eq!(denom, 4);
    }

    #[test]
    fn test_continued_fraction_thirds() {
        // 0.333... ~ 1/3
        let (num, denom) = continued_fraction_approximation(0.333333333, 9);
        assert_eq!(num, 1);
        assert_eq!(denom, 3);

        // 0.666... ~ 2/3
        let (num, denom) = continued_fraction_approximation(0.666666666, 9);
        assert_eq!(num, 2);
        assert_eq!(denom, 3);
    }

    #[test]
    fn test_to_fraction_basic() {
        // 1.5 = 1 1/2
        let (whole, num, denom) = to_fraction(1.5, 9);
        assert_eq!(whole, 1);
        assert_eq!(num, 1);
        assert_eq!(denom, 2);

        // 2.25 = 2 1/4
        let (whole, num, denom) = to_fraction(2.25, 9);
        assert_eq!(whole, 2);
        assert_eq!(num, 1);
        assert_eq!(denom, 4);
    }

    #[test]
    fn test_to_fraction_negative() {
        // -1.5 = -1 1/2
        let (whole, num, denom) = to_fraction(-1.5, 9);
        assert_eq!(whole, -1);
        assert_eq!(num, -1);
        assert_eq!(denom, 2);
    }

    #[test]
    fn test_to_fraction_fixed_denom() {
        // 0.375 with denominator 8 = 3/8
        let (whole, num, denom) = to_fraction_fixed_denom(0.375, 8);
        assert_eq!(whole, 0);
        assert_eq!(num, 3);
        assert_eq!(denom, 8);

        // 0.5 with denominator 16 = 8/16
        let (whole, num, denom) = to_fraction_fixed_denom(0.5, 16);
        assert_eq!(whole, 0);
        assert_eq!(num, 8);
        assert_eq!(denom, 16);
    }

    #[test]
    fn test_format_fraction_basic() {
        assert_eq!(format_fraction(0.5, "# ?/?"), "1/2");
        assert_eq!(format_fraction(1.5, "# ?/?"), "1 1/2");
        assert_eq!(format_fraction(2.0, "# ?/?"), "2");
    }

    #[test]
    fn test_format_fraction_fixed_denom() {
        assert_eq!(format_fraction(0.375, "# ?/8"), "3/8");
        assert_eq!(format_fraction(0.5, "# ?/16"), " 8/16");
    }

    #[test]
    fn test_format_fraction_two_digits() {
        // 0.333... with two digit denominator
        let result = format_fraction(0.333333333, "# ??/??");
        assert!(result.contains("1/") && result.contains("3"));
    }

    #[test]
    fn test_format_number_routes_scientific() {
        assert_eq!(format_number(1234567.0, "0.00E+00", false), "1.23E+06");
    }

    #[test]
    fn test_format_number_routes_fraction() {
        assert_eq!(format_number(0.5, "# ?/?", false), "1/2");
    }

    #[test]
    fn test_builtin_format_12() {
        // numFmtId 12 is "# ?/?"
        let format = get_builtin_format(12).unwrap();
        assert_eq!(format_number(0.5, format, false), "1/2");
    }

    #[test]
    fn test_builtin_format_13() {
        // numFmtId 13 is "# ??/??"
        let format = get_builtin_format(13).unwrap();
        let result = format_number(0.333333333, format, false);
        assert!(result.contains("1/") && result.contains("3"));
    }

    #[test]
    fn test_builtin_format_11() {
        // numFmtId 11 is "0.00E+00"
        let format = get_builtin_format(11).unwrap();
        assert_eq!(format_number(1234567.0, format, false), "1.23E+06");
    }

    #[test]
    fn test_builtin_format_48() {
        // numFmtId 48 is "##0.0E+0"
        let format = get_builtin_format(48).unwrap();
        assert_eq!(format_number(1234567.0, format, false), "1.2E+06");
    }

    #[test]
    fn test_date1904_system() {
        // Test that 1904 system produces dates ~4 years later than 1900 system for same serial
        // The difference should be 1462 days (4 years + 1 day for the phantom Feb 29, 1900)
        let result_1900 = format_number(1000.0, "yyyy-mm-dd", false);
        let result_1904 = format_number(1000.0, "yyyy-mm-dd", true);

        // 1904 dates should be about 4 years later than 1900 dates for the same serial
        // Serial 1000 in 1900 = ~Sep 26, 1902
        // Serial 1000 in 1904 = ~Sep 27, 1906
        assert!(result_1900.starts_with("190")); // 1900s
        assert!(result_1904.starts_with("190")); // 1900s
        assert_ne!(result_1900, result_1904); // Different dates
    }
}
