//! Database functions: DAVERAGE, DCOUNT, DCOUNTA, DGET, DMAX, DMIN,
//! DPRODUCT, DSTDEV, DSTDEVP, DSUM, DVAR, DVARP.
//!
//! These functions operate on database-like ranges with criteria matching.

use crate::context::EvalContext;
use crate::reference::RangeRef;
use crate::value::{CalcValue, ScalarValue, XlError};

use super::{require_scalar, FunctionArg, FunctionDef, FunctionRegistry, ParamKind};

/// Database function algorithm trait.
///
/// Each database function implements this trait to define its specific
/// aggregation behavior (average, count, sum, etc.).
trait DStarAlgorithm {
    /// Process a matching value from the database.
    ///
    /// Returns `true` to continue iteration, `false` to stop early.
    fn process_match(&mut self, value: &ScalarValue) -> bool;

    /// Get the final result after all matches have been processed.
    fn result(&self) -> CalcValue;

    /// Whether this algorithm allows an empty match field (column name).
    ///
    /// DCOUNT and DCOUNTA allow empty field, others require a field name.
    fn allow_empty_match_field(&self) -> bool {
        false
    }
}

/// Runner for all D* functions.
///
/// This function implements the common database query logic:
/// 1. Extract database range, field column, and criteria range
/// 2. Iterate through database rows
/// 3. Check each row against criteria
/// 4. If row matches, pass the field value to the algorithm
/// 5. Return the algorithm's result
fn run_dstar<A: DStarAlgorithm>(
    args: &[FunctionArg<'_>],
    ctx: &EvalContext<'_>,
    mut algorithm: A,
) -> CalcValue {
    // Validate we have exactly 3 arguments
    if args.len() != 3 {
        return CalcValue::error(XlError::Value);
    }

    // Extract database range
    let db_range = match &args[0] {
        FunctionArg::Range { range, .. } => range,
        _ => return CalcValue::error(XlError::Value),
    };

    // Extract field column identifier (number or text)
    let field_arg = require_scalar(&args[1]);
    if let ScalarValue::Error(e) = field_arg {
        return CalcValue::error(*e);
    }

    // Extract criteria range
    let criteria_range = match &args[2] {
        FunctionArg::Range { range, ctx: _ } => range,
        _ => return CalcValue::error(XlError::Value),
    };

    // Determine field column index
    let field_col_idx = match resolve_field_column(field_arg, db_range, ctx) {
        Ok(idx) => idx,
        Err(e) => {
            // Allow empty field for DCOUNT/DCOUNTA
            if algorithm.allow_empty_match_field() && matches!(field_arg, ScalarValue::Blank) {
                None
            } else {
                return e;
            }
        }
    };

    // Validate database has at least 2 rows (header + data)
    let db_height = (db_range.end_row - db_range.start_row + 1) as usize;
    if db_height < 2 {
        return CalcValue::error(XlError::Value);
    }

    // Iterate through database rows (skip header row 0)
    for row_idx in 1..db_height {
        let absolute_row = db_range.start_row + row_idx as u32;

        // Check if this row matches all criteria
        match fulfills_conditions(db_range, row_idx, criteria_range, ctx) {
            Ok(true) => {
                // Row matches, get the field value
                let value = if let Some(col_idx) = field_col_idx {
                    let absolute_col = db_range.start_col + col_idx as u32;
                    ctx.provider
                        .cell_value(db_range.sheet.as_deref(), absolute_row, absolute_col)
                } else {
                    // Empty field for DCOUNT/DCOUNTA: treat as 0 if not numeric
                    ScalarValue::Number(0.0)
                };

                // Process the match
                if !algorithm.process_match(&value) {
                    break; // Algorithm wants to stop early
                }
            }
            Ok(false) => {
                // Row doesn't match, continue
            }
            Err(e) => {
                return e;
            }
        }
    }

    algorithm.result()
}

/// Resolves the field column identifier to a zero-based column index.
///
/// The field can be:
/// - A number: 1-based column index (1 = first column)
/// - A text: column header name to match (case-insensitive)
fn resolve_field_column(
    field: &ScalarValue,
    db_range: &RangeRef,
    ctx: &EvalContext<'_>,
) -> Result<Option<usize>, CalcValue> {
    match field {
        ScalarValue::Number(n) => {
            // 1-based column index
            let col_num = *n as i32 - 1;
            let db_width = (db_range.end_col - db_range.start_col + 1) as i32;
            if col_num < 0 || col_num >= db_width {
                Err(CalcValue::error(XlError::Value))
            } else {
                Ok(Some(col_num as usize))
            }
        }
        ScalarValue::Text(name) => {
            // Column header name - search in first row
            let db_width = (db_range.end_col - db_range.start_col + 1) as usize;
            for col_idx in 0..db_width {
                let absolute_col = db_range.start_col + col_idx as u32;
                let header_value = ctx.provider.cell_value(
                    db_range.sheet.as_deref(),
                    db_range.start_row,
                    absolute_col,
                );

                match header_value {
                    ScalarValue::Blank | ScalarValue::Error(_) => continue,
                    ScalarValue::Text(ref header_text) => {
                        if header_text.eq_ignore_ascii_case(name) {
                            return Ok(Some(col_idx));
                        }
                    }
                    ScalarValue::Number(n) => {
                        if name == &n.to_string() {
                            return Ok(Some(col_idx));
                        }
                    }
                    ScalarValue::Bool(b) => {
                        let bool_str = if b { "TRUE" } else { "FALSE" };
                        if name.eq_ignore_ascii_case(bool_str) {
                            return Ok(Some(col_idx));
                        }
                    }
                }
            }
            // Column not found
            Err(CalcValue::error(XlError::Value))
        }
        ScalarValue::Blank => Err(CalcValue::error(XlError::Value)),
        ScalarValue::Bool(_) => Err(CalcValue::error(XlError::Value)),
        ScalarValue::Error(e) => Err(CalcValue::error(*e)),
    }
}

/// Checks if a database row fulfills all criteria.
///
/// Criteria are organized as:
/// - First row: column headers (matching database headers or formulas)
/// - Subsequent rows: conditions (ORed across rows, ANDed within a row)
fn fulfills_conditions(
    db_range: &RangeRef,
    db_row_idx: usize,
    criteria_range: &RangeRef,
    ctx: &EvalContext<'_>,
) -> Result<bool, CalcValue> {
    let criteria_height = (criteria_range.end_row - criteria_range.start_row + 1) as usize;
    let criteria_width = (criteria_range.end_col - criteria_range.start_col + 1) as usize;

    // Must have at least a header row
    if criteria_height < 1 {
        return Ok(true); // No criteria = all rows match
    }

    // If only header row (no criteria rows), all rows match
    if criteria_height == 1 {
        return Ok(true);
    }

    // Iterate through criteria rows (skip header row 0)
    for crit_row_idx in 1..criteria_height {
        let mut row_matches = true;

        // All conditions in this row must match (AND)
        for crit_col_idx in 0..criteria_width {
            let crit_abs_row = criteria_range.start_row + crit_row_idx as u32;
            let crit_abs_col = criteria_range.start_col + crit_col_idx as u32;

            // Get the condition value
            let condition = ctx.provider.cell_value(
                criteria_range.sheet.as_deref(),
                crit_abs_row,
                crit_abs_col,
            );

            // Empty condition always matches
            if matches!(condition, ScalarValue::Blank) {
                continue;
            }

            // Get the column header from criteria
            let header_abs_row = criteria_range.start_row;
            let header_value = ctx.provider.cell_value(
                criteria_range.sheet.as_deref(),
                header_abs_row,
                crit_abs_col,
            );

            // Find matching column in database
            let db_col_idx = match find_db_column(&header_value, db_range, ctx) {
                Some(idx) => idx,
                None => {
                    // Header doesn't match any database column
                    // This could be a formula condition (not implemented)
                    return Err(CalcValue::error(XlError::Value));
                }
            };

            // Get the database value for this row/column
            let db_abs_row = db_range.start_row + db_row_idx as u32;
            let db_abs_col = db_range.start_col + db_col_idx as u32;
            let db_value =
                ctx.provider
                    .cell_value(db_range.sheet.as_deref(), db_abs_row, db_abs_col);

            // Test the condition
            if !test_condition(&db_value, &condition) {
                row_matches = false;
                break;
            }
        }

        // If any criteria row matches, the database row matches (OR)
        if row_matches {
            return Ok(true);
        }
    }

    // No criteria row matched
    Ok(false)
}

/// Finds a database column index matching the given header value.
fn find_db_column(
    header: &ScalarValue,
    db_range: &RangeRef,
    ctx: &EvalContext<'_>,
) -> Option<usize> {
    if matches!(header, ScalarValue::Blank | ScalarValue::Error(_)) {
        return None;
    }

    let db_width = (db_range.end_col - db_range.start_col + 1) as usize;
    let header_text = match header {
        ScalarValue::Text(s) => s.clone(),
        ScalarValue::Number(n) => n.to_string(),
        ScalarValue::Bool(b) => (if *b { "TRUE" } else { "FALSE" }).to_string(),
        _ => return None,
    };

    for col_idx in 0..db_width {
        let absolute_col = db_range.start_col + col_idx as u32;
        let db_header =
            ctx.provider
                .cell_value(db_range.sheet.as_deref(), db_range.start_row, absolute_col);

        let db_header_text = match db_header {
            ScalarValue::Text(s) => s,
            ScalarValue::Number(n) => n.to_string(),
            ScalarValue::Bool(b) => (if b { "TRUE" } else { "FALSE" }).to_string(),
            ScalarValue::Blank | ScalarValue::Error(_) => continue,
        };

        if header_text.eq_ignore_ascii_case(&db_header_text) {
            return Some(col_idx);
        }
    }

    None
}

/// Tests if a value satisfies a condition.
///
/// Conditions can be:
/// - Comparison operators: `<`, `>`, `<=`, `>=`, `=`, `<>`
/// - Wildcards: `*` (any sequence), `?` (any single char)
/// - Plain text: case-insensitive exact match or prefix match
fn test_condition(value: &ScalarValue, condition: &ScalarValue) -> bool {
    match condition {
        ScalarValue::Text(cond_str) => {
            // Parse condition string for operators
            if let Some(rest) = cond_str.strip_prefix('<') {
                if let Some(rest) = rest.strip_prefix('=') {
                    // <=
                    test_numeric_condition(value, rest, |v, c| v <= c)
                } else if let Some(rest) = rest.strip_prefix('>') {
                    // <>
                    if is_number(rest) {
                        test_numeric_condition(value, rest, |v, c| v != c)
                    } else {
                        test_string_condition(value, rest, |v, c| !v.eq_ignore_ascii_case(c))
                    }
                } else {
                    // <
                    test_numeric_condition(value, rest, |v, c| v < c)
                }
            } else if let Some(rest) = cond_str.strip_prefix('>') {
                if let Some(rest) = rest.strip_prefix('=') {
                    // >=
                    test_numeric_condition(value, rest, |v, c| v >= c)
                } else {
                    // >
                    test_numeric_condition(value, rest, |v, c| v > c)
                }
            } else if let Some(rest) = cond_str.strip_prefix('=') {
                // =
                if rest.is_empty() {
                    // "=" matches blank
                    return matches!(value, ScalarValue::Blank);
                }
                if is_number(rest) {
                    test_numeric_condition(value, rest, |v, c| v == c)
                } else {
                    test_string_condition(value, rest, |v, c| v.eq_ignore_ascii_case(c))
                }
            } else {
                // No operator: wildcard text match
                if cond_str.is_empty() {
                    return matches!(value, ScalarValue::Text(_));
                }
                test_wildcard(value, cond_str)
            }
        }
        ScalarValue::Number(cond_num) => {
            // Numeric condition: exact match
            if let Some(val_num) = get_number_from_value(value) {
                return val_num == *cond_num;
            }
            false
        }
        ScalarValue::Error(cond_err) => {
            // Error condition: exact match
            if let ScalarValue::Error(val_err) = value {
                return val_err == cond_err;
            }
            false
        }
        ScalarValue::Bool(_) | ScalarValue::Blank => {
            // Boolean or blank conditions not typically used
            false
        }
    }
}

/// Tests a numeric comparison condition.
fn test_numeric_condition<F>(value: &ScalarValue, cond_str: &str, op: F) -> bool
where
    F: Fn(f64, f64) -> bool,
{
    let Some(val_num) = get_number_from_value(value) else {
        return false;
    };

    let Ok(cond_num) = cond_str.parse::<f64>() else {
        return false;
    };

    op(val_num, cond_num)
}

/// Tests a string comparison condition.
fn test_string_condition<F>(value: &ScalarValue, cond_str: &str, op: F) -> bool
where
    F: Fn(&str, &str) -> bool,
{
    let val_str = match value {
        ScalarValue::Text(s) => s.as_str(),
        ScalarValue::Blank => "",
        _ => return false,
    };

    op(val_str, cond_str)
}

/// Tests wildcard pattern matching.
///
/// Supports:
/// - `*` matches any sequence of characters
/// - `?` matches any single character
fn test_wildcard(value: &ScalarValue, pattern: &str) -> bool {
    let val_str = match value {
        ScalarValue::Text(s) => s.to_lowercase(),
        ScalarValue::Blank => String::new(),
        _ => return false,
    };

    let pattern = pattern.to_lowercase();

    // Convert Excel wildcard pattern to regex-like matching
    wildcard_match(&val_str, &pattern)
}

/// Simple wildcard matching without regex dependency.
fn wildcard_match(text: &str, pattern: &str) -> bool {
    let text_chars: Vec<char> = text.chars().collect();
    let pattern_chars: Vec<char> = pattern.chars().collect();

    let mut ti = 0;
    let mut pi = 0;
    let mut star_idx = None;
    let mut match_idx = 0;

    while ti < text_chars.len() {
        if pi < pattern_chars.len() {
            match pattern_chars[pi] {
                '*' => {
                    star_idx = Some(pi);
                    match_idx = ti;
                    pi += 1;
                    continue;
                }
                '?' => {
                    ti += 1;
                    pi += 1;
                    continue;
                }
                c if c == text_chars[ti] => {
                    ti += 1;
                    pi += 1;
                    continue;
                }
                _ => {}
            }
        }

        if let Some(si) = star_idx {
            pi = si + 1;
            match_idx += 1;
            ti = match_idx;
        } else {
            return false;
        }
    }

    while pi < pattern_chars.len() && pattern_chars[pi] == '*' {
        pi += 1;
    }

    pi == pattern_chars.len()
}

/// Checks if a string represents a number.
fn is_number(s: &str) -> bool {
    s.parse::<f64>().is_ok()
}

/// Extracts a number from a ScalarValue.
fn get_number_from_value(value: &ScalarValue) -> Option<f64> {
    match value {
        ScalarValue::Number(n) => Some(*n),
        ScalarValue::Text(s) => s.parse::<f64>().ok(),
        _ => None,
    }
}

// ---------------------------------------------------------------------------
// Algorithm implementations
// ---------------------------------------------------------------------------

/// DAVERAGE algorithm: calculates average of matching numeric values.
struct DAverage {
    sum: f64,
    count: usize,
}

impl DAverage {
    fn new() -> Self {
        Self { sum: 0.0, count: 0 }
    }
}

impl DStarAlgorithm for DAverage {
    fn process_match(&mut self, value: &ScalarValue) -> bool {
        if let ScalarValue::Number(n) = value {
            self.sum += n;
            self.count += 1;
        }
        true
    }

    fn result(&self) -> CalcValue {
        if self.count == 0 {
            CalcValue::number(0.0)
        } else {
            CalcValue::number(self.sum / self.count as f64)
        }
    }
}

/// DCOUNT algorithm: counts matching numeric values.
struct DCount {
    count: usize,
}

impl DCount {
    fn new() -> Self {
        Self { count: 0 }
    }
}

impl DStarAlgorithm for DCount {
    fn process_match(&mut self, value: &ScalarValue) -> bool {
        if matches!(value, ScalarValue::Number(_)) {
            self.count += 1;
        }
        true
    }

    fn result(&self) -> CalcValue {
        CalcValue::number(self.count as f64)
    }

    fn allow_empty_match_field(&self) -> bool {
        true
    }
}

/// DCOUNTA algorithm: counts matching non-blank values.
struct DCountA {
    count: usize,
}

impl DCountA {
    fn new() -> Self {
        Self { count: 0 }
    }
}

impl DStarAlgorithm for DCountA {
    fn process_match(&mut self, value: &ScalarValue) -> bool {
        if !matches!(value, ScalarValue::Blank) {
            self.count += 1;
        }
        true
    }

    fn result(&self) -> CalcValue {
        CalcValue::number(self.count as f64)
    }

    fn allow_empty_match_field(&self) -> bool {
        true
    }
}

/// DGET algorithm: extracts single matching value.
struct DGet {
    result: Option<ScalarValue>,
    multiple: bool,
}

impl DGet {
    fn new() -> Self {
        Self {
            result: None,
            multiple: false,
        }
    }
}

impl DStarAlgorithm for DGet {
    fn process_match(&mut self, value: &ScalarValue) -> bool {
        match &self.result {
            None => {
                self.result = Some(value.clone());
                true
            }
            Some(prev) => {
                // Check for multiple non-blank matches
                if matches!(prev, ScalarValue::Blank) {
                    self.result = Some(value.clone());
                    true
                } else if !matches!(value, ScalarValue::Blank) {
                    // Multiple non-blank values found
                    self.multiple = true;
                    false // Stop iteration
                } else {
                    true
                }
            }
        }
    }

    fn result(&self) -> CalcValue {
        if self.multiple {
            return CalcValue::error(XlError::Num);
        }

        match &self.result {
            None => CalcValue::error(XlError::Value),
            Some(ScalarValue::Blank) => CalcValue::error(XlError::Value),
            Some(ScalarValue::Text(s)) if s.is_empty() => CalcValue::error(XlError::Value),
            Some(v) => CalcValue::from_scalar(v.clone()),
        }
    }
}

/// DMAX algorithm: finds maximum of matching numeric values.
struct DMax {
    max: Option<f64>,
}

impl DMax {
    fn new() -> Self {
        Self { max: None }
    }
}

impl DStarAlgorithm for DMax {
    fn process_match(&mut self, value: &ScalarValue) -> bool {
        if let ScalarValue::Number(n) = value {
            self.max = Some(self.max.map_or(*n, |m| m.max(*n)));
        }
        true
    }

    fn result(&self) -> CalcValue {
        CalcValue::number(self.max.unwrap_or(0.0))
    }
}

/// DMIN algorithm: finds minimum of matching numeric values.
struct DMin {
    min: Option<f64>,
}

impl DMin {
    fn new() -> Self {
        Self { min: None }
    }
}

impl DStarAlgorithm for DMin {
    fn process_match(&mut self, value: &ScalarValue) -> bool {
        if let ScalarValue::Number(n) = value {
            self.min = Some(self.min.map_or(*n, |m| m.min(*n)));
        }
        true
    }

    fn result(&self) -> CalcValue {
        CalcValue::number(self.min.unwrap_or(0.0))
    }
}

/// DPRODUCT algorithm: multiplies matching numeric values.
struct DProduct {
    product: f64,
    initialized: bool,
}

impl DProduct {
    fn new() -> Self {
        Self {
            product: 0.0,
            initialized: false,
        }
    }
}

impl DStarAlgorithm for DProduct {
    fn process_match(&mut self, value: &ScalarValue) -> bool {
        if let ScalarValue::Number(n) = value {
            if self.initialized {
                self.product *= n;
            } else {
                self.product = *n;
                self.initialized = true;
            }
        }
        true
    }

    fn result(&self) -> CalcValue {
        CalcValue::number(self.product)
    }
}

/// DSUM algorithm: sums matching numeric values.
struct DSum {
    sum: f64,
}

impl DSum {
    fn new() -> Self {
        Self { sum: 0.0 }
    }
}

impl DStarAlgorithm for DSum {
    fn process_match(&mut self, value: &ScalarValue) -> bool {
        if let ScalarValue::Number(n) = value {
            self.sum += n;
        }
        true
    }

    fn result(&self) -> CalcValue {
        CalcValue::number(self.sum)
    }
}

/// DSTDEV algorithm: calculates sample standard deviation.
struct DStdev {
    values: Vec<f64>,
}

impl DStdev {
    fn new() -> Self {
        Self { values: Vec::new() }
    }
}

impl DStarAlgorithm for DStdev {
    fn process_match(&mut self, value: &ScalarValue) -> bool {
        if let ScalarValue::Number(n) = value {
            self.values.push(*n);
        }
        true
    }

    fn result(&self) -> CalcValue {
        if self.values.len() < 2 {
            return CalcValue::error(XlError::Div0);
        }

        let devsq = calc_devsq(&self.values);
        let variance = devsq / (self.values.len() - 1) as f64;
        CalcValue::number(variance.sqrt())
    }
}

/// DSTDEVP algorithm: calculates population standard deviation.
struct DStdevP {
    values: Vec<f64>,
}

impl DStdevP {
    fn new() -> Self {
        Self { values: Vec::new() }
    }
}

impl DStarAlgorithm for DStdevP {
    fn process_match(&mut self, value: &ScalarValue) -> bool {
        if let ScalarValue::Number(n) = value {
            self.values.push(*n);
        }
        true
    }

    fn result(&self) -> CalcValue {
        if self.values.len() < 2 {
            return CalcValue::error(XlError::Div0);
        }

        let devsq = calc_devsq(&self.values);
        let variance = devsq / self.values.len() as f64;
        CalcValue::number(variance.sqrt())
    }
}

/// DVAR algorithm: calculates sample variance.
struct DVar {
    values: Vec<f64>,
}

impl DVar {
    fn new() -> Self {
        Self { values: Vec::new() }
    }
}

impl DStarAlgorithm for DVar {
    fn process_match(&mut self, value: &ScalarValue) -> bool {
        if let ScalarValue::Number(n) = value {
            self.values.push(*n);
        }
        true
    }

    fn result(&self) -> CalcValue {
        if self.values.len() < 2 {
            return CalcValue::error(XlError::Div0);
        }

        let devsq = calc_devsq(&self.values);
        let variance = devsq / (self.values.len() - 1) as f64;
        CalcValue::number(variance)
    }
}

/// DVARP algorithm: calculates population variance.
struct DVarP {
    values: Vec<f64>,
}

impl DVarP {
    fn new() -> Self {
        Self { values: Vec::new() }
    }
}

impl DStarAlgorithm for DVarP {
    fn process_match(&mut self, value: &ScalarValue) -> bool {
        if let ScalarValue::Number(n) = value {
            self.values.push(*n);
        }
        true
    }

    fn result(&self) -> CalcValue {
        if self.values.len() < 2 {
            return CalcValue::error(XlError::Div0);
        }

        let devsq = calc_devsq(&self.values);
        let variance = devsq / self.values.len() as f64;
        CalcValue::number(variance)
    }
}

/// Calculates the sum of squared deviations from the mean.
fn calc_devsq(values: &[f64]) -> f64 {
    if values.is_empty() {
        return 0.0;
    }

    let mean = values.iter().sum::<f64>() / values.len() as f64;
    values.iter().map(|v| (v - mean).powi(2)).sum()
}

// ---------------------------------------------------------------------------
// Function definitions
// ---------------------------------------------------------------------------

/// `DAVERAGE(database, field, criteria)`
pub fn fn_daverage(args: &[FunctionArg<'_>], ctx: &EvalContext<'_>) -> CalcValue {
    run_dstar(args, ctx, DAverage::new())
}

/// `DCOUNT(database, field, criteria)`
pub fn fn_dcount(args: &[FunctionArg<'_>], ctx: &EvalContext<'_>) -> CalcValue {
    run_dstar(args, ctx, DCount::new())
}

/// `DCOUNTA(database, field, criteria)`
pub fn fn_dcounta(args: &[FunctionArg<'_>], ctx: &EvalContext<'_>) -> CalcValue {
    run_dstar(args, ctx, DCountA::new())
}

/// `DGET(database, field, criteria)`
pub fn fn_dget(args: &[FunctionArg<'_>], ctx: &EvalContext<'_>) -> CalcValue {
    run_dstar(args, ctx, DGet::new())
}

/// `DMAX(database, field, criteria)`
pub fn fn_dmax(args: &[FunctionArg<'_>], ctx: &EvalContext<'_>) -> CalcValue {
    run_dstar(args, ctx, DMax::new())
}

/// `DMIN(database, field, criteria)`
pub fn fn_dmin(args: &[FunctionArg<'_>], ctx: &EvalContext<'_>) -> CalcValue {
    run_dstar(args, ctx, DMin::new())
}

/// `DPRODUCT(database, field, criteria)`
pub fn fn_dproduct(args: &[FunctionArg<'_>], ctx: &EvalContext<'_>) -> CalcValue {
    run_dstar(args, ctx, DProduct::new())
}

/// `DSTDEV(database, field, criteria)`
pub fn fn_dstdev(args: &[FunctionArg<'_>], ctx: &EvalContext<'_>) -> CalcValue {
    run_dstar(args, ctx, DStdev::new())
}

/// `DSTDEVP(database, field, criteria)`
pub fn fn_dstdevp(args: &[FunctionArg<'_>], ctx: &EvalContext<'_>) -> CalcValue {
    run_dstar(args, ctx, DStdevP::new())
}

/// `DSUM(database, field, criteria)`
pub fn fn_dsum(args: &[FunctionArg<'_>], ctx: &EvalContext<'_>) -> CalcValue {
    run_dstar(args, ctx, DSum::new())
}

/// `DVAR(database, field, criteria)`
pub fn fn_dvar(args: &[FunctionArg<'_>], ctx: &EvalContext<'_>) -> CalcValue {
    run_dstar(args, ctx, DVar::new())
}

/// `DVARP(database, field, criteria)`
pub fn fn_dvarp(args: &[FunctionArg<'_>], ctx: &EvalContext<'_>) -> CalcValue {
    run_dstar(args, ctx, DVarP::new())
}

/// Registers all database functions into the given registry.
pub fn register(registry: &mut FunctionRegistry) {
    registry.register(FunctionDef {
        name: "DAVERAGE",
        min_args: 3,
        max_args: 3,
        param_kinds: &[ParamKind::Range, ParamKind::Scalar, ParamKind::Range],
        func: fn_daverage,
    });
    registry.register(FunctionDef {
        name: "DCOUNT",
        min_args: 3,
        max_args: 3,
        param_kinds: &[ParamKind::Range, ParamKind::Scalar, ParamKind::Range],
        func: fn_dcount,
    });
    registry.register(FunctionDef {
        name: "DCOUNTA",
        min_args: 3,
        max_args: 3,
        param_kinds: &[ParamKind::Range, ParamKind::Scalar, ParamKind::Range],
        func: fn_dcounta,
    });
    registry.register(FunctionDef {
        name: "DGET",
        min_args: 3,
        max_args: 3,
        param_kinds: &[ParamKind::Range, ParamKind::Scalar, ParamKind::Range],
        func: fn_dget,
    });
    registry.register(FunctionDef {
        name: "DMAX",
        min_args: 3,
        max_args: 3,
        param_kinds: &[ParamKind::Range, ParamKind::Scalar, ParamKind::Range],
        func: fn_dmax,
    });
    registry.register(FunctionDef {
        name: "DMIN",
        min_args: 3,
        max_args: 3,
        param_kinds: &[ParamKind::Range, ParamKind::Scalar, ParamKind::Range],
        func: fn_dmin,
    });
    registry.register(FunctionDef {
        name: "DPRODUCT",
        min_args: 3,
        max_args: 3,
        param_kinds: &[ParamKind::Range, ParamKind::Scalar, ParamKind::Range],
        func: fn_dproduct,
    });
    registry.register(FunctionDef {
        name: "DSTDEV",
        min_args: 3,
        max_args: 3,
        param_kinds: &[ParamKind::Range, ParamKind::Scalar, ParamKind::Range],
        func: fn_dstdev,
    });
    registry.register(FunctionDef {
        name: "DSTDEVP",
        min_args: 3,
        max_args: 3,
        param_kinds: &[ParamKind::Range, ParamKind::Scalar, ParamKind::Range],
        func: fn_dstdevp,
    });
    registry.register(FunctionDef {
        name: "DSUM",
        min_args: 3,
        max_args: 3,
        param_kinds: &[ParamKind::Range, ParamKind::Scalar, ParamKind::Range],
        func: fn_dsum,
    });
    registry.register(FunctionDef {
        name: "DVAR",
        min_args: 3,
        max_args: 3,
        param_kinds: &[ParamKind::Range, ParamKind::Scalar, ParamKind::Range],
        func: fn_dvar,
    });
    registry.register(FunctionDef {
        name: "DVARP",
        min_args: 3,
        max_args: 3,
        param_kinds: &[ParamKind::Range, ParamKind::Scalar, ParamKind::Range],
        func: fn_dvarp,
    });
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::context::CellDataProvider;
    use std::collections::HashMap;

    struct TestData {
        cells: HashMap<(Option<String>, u32, u32), ScalarValue>,
    }

    impl TestData {
        fn new() -> Self {
            Self {
                cells: HashMap::new(),
            }
        }

        fn set(&mut self, sheet: Option<&str>, row: u32, col: u32, value: ScalarValue) {
            self.cells
                .insert((sheet.map(|s| s.to_string()), row, col), value);
        }
    }

    impl CellDataProvider for TestData {
        fn cell_value(&self, sheet: Option<&str>, row: u32, col: u32) -> ScalarValue {
            self.cells
                .get(&(sheet.map(|s| s.to_string()), row, col))
                .cloned()
                .unwrap_or(ScalarValue::Blank)
        }
    }

    #[test]
    fn test_daverage_basic() {
        let mut data = TestData::new();

        // Database: A1:C4
        // Headers
        data.set(None, 1, 1, ScalarValue::Text("Tree".to_string()));
        data.set(None, 1, 2, ScalarValue::Text("Height".to_string()));
        data.set(None, 1, 3, ScalarValue::Text("Age".to_string()));

        // Data rows
        data.set(None, 2, 1, ScalarValue::Text("Apple".to_string()));
        data.set(None, 2, 2, ScalarValue::Number(18.0));
        data.set(None, 2, 3, ScalarValue::Number(20.0));

        data.set(None, 3, 1, ScalarValue::Text("Pear".to_string()));
        data.set(None, 3, 2, ScalarValue::Number(12.0));
        data.set(None, 3, 3, ScalarValue::Number(12.0));

        data.set(None, 4, 1, ScalarValue::Text("Apple".to_string()));
        data.set(None, 4, 2, ScalarValue::Number(14.0));
        data.set(None, 4, 3, ScalarValue::Number(15.0));

        // Criteria: E1:E2
        data.set(None, 1, 5, ScalarValue::Text("Tree".to_string()));
        data.set(None, 2, 5, ScalarValue::Text("Apple".to_string()));

        let ctx = EvalContext::new(&data, None::<String>, 1, 1);

        let db_range = RangeRef {
            sheet: None,
            start_row: 1,
            start_col: 1,
            end_row: 4,
            end_col: 3,
        };

        let criteria_range = RangeRef {
            sheet: None,
            start_row: 1,
            start_col: 5,
            end_row: 2,
            end_col: 5,
        };

        let args = vec![
            FunctionArg::Range {
                range: &db_range,
                ctx: &ctx,
            },
            FunctionArg::Value(CalcValue::text("Height")),
            FunctionArg::Range {
                range: &criteria_range,
                ctx: &ctx,
            },
        ];

        let result = fn_daverage(&args, &ctx);
        assert_eq!(result.into_scalar(), ScalarValue::Number(16.0)); // (18 + 14) / 2
    }

    #[test]
    fn test_dcount_with_numeric_field() {
        let mut data = TestData::new();

        // Database
        data.set(None, 1, 1, ScalarValue::Text("Tree".to_string()));
        data.set(None, 1, 2, ScalarValue::Text("Height".to_string()));

        data.set(None, 2, 1, ScalarValue::Text("Apple".to_string()));
        data.set(None, 2, 2, ScalarValue::Number(18.0));

        data.set(None, 3, 1, ScalarValue::Text("Pear".to_string()));
        data.set(None, 3, 2, ScalarValue::Number(12.0));

        // Criteria (empty = all rows)
        data.set(None, 1, 4, ScalarValue::Text("Tree".to_string()));

        let ctx = EvalContext::new(&data, None::<String>, 1, 1);

        let db_range = RangeRef {
            sheet: None,
            start_row: 1,
            start_col: 1,
            end_row: 3,
            end_col: 2,
        };

        let criteria_range = RangeRef {
            sheet: None,
            start_row: 1,
            start_col: 4,
            end_row: 1,
            end_col: 4,
        };

        let args = vec![
            FunctionArg::Range {
                range: &db_range,
                ctx: &ctx,
            },
            FunctionArg::Value(CalcValue::number(2.0)), // Column 2 = Height
            FunctionArg::Range {
                range: &criteria_range,
                ctx: &ctx,
            },
        ];

        let result = fn_dcount(&args, &ctx);
        assert_eq!(result.into_scalar(), ScalarValue::Number(2.0));
    }

    #[test]
    fn test_dsum() {
        let mut data = TestData::new();

        data.set(None, 1, 1, ScalarValue::Text("Item".to_string()));
        data.set(None, 1, 2, ScalarValue::Text("Value".to_string()));

        data.set(None, 2, 1, ScalarValue::Text("A".to_string()));
        data.set(None, 2, 2, ScalarValue::Number(10.0));

        data.set(None, 3, 1, ScalarValue::Text("B".to_string()));
        data.set(None, 3, 2, ScalarValue::Number(20.0));

        data.set(None, 4, 1, ScalarValue::Text("A".to_string()));
        data.set(None, 4, 2, ScalarValue::Number(30.0));

        // Criteria: Item = "A"
        data.set(None, 1, 4, ScalarValue::Text("Item".to_string()));
        data.set(None, 2, 4, ScalarValue::Text("A".to_string()));

        let ctx = EvalContext::new(&data, None::<String>, 1, 1);

        let db_range = RangeRef {
            sheet: None,
            start_row: 1,
            start_col: 1,
            end_row: 4,
            end_col: 2,
        };

        let criteria_range = RangeRef {
            sheet: None,
            start_row: 1,
            start_col: 4,
            end_row: 2,
            end_col: 4,
        };

        let args = vec![
            FunctionArg::Range {
                range: &db_range,
                ctx: &ctx,
            },
            FunctionArg::Value(CalcValue::text("Value")),
            FunctionArg::Range {
                range: &criteria_range,
                ctx: &ctx,
            },
        ];

        let result = fn_dsum(&args, &ctx);
        assert_eq!(result.into_scalar(), ScalarValue::Number(40.0)); // 10 + 30
    }

    #[test]
    fn test_wildcard_matching() {
        assert!(wildcard_match("apple", "app*"));
        assert!(wildcard_match("apple", "*le"));
        assert!(wildcard_match("apple", "a*e"));
        assert!(wildcard_match("apple", "a??le"));
        assert!(!wildcard_match("apple", "a?e"));
        assert!(wildcard_match("apple", "apple"));
        assert!(wildcard_match("", "*"));
        assert!(!wildcard_match("apple", "orange"));
    }
}
