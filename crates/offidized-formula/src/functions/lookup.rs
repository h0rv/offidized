//! Lookup and reference functions.

use crate::context::EvalContext;
use crate::reference::{column_index_to_letters, RangeRef};
use crate::value::{compare_values, CalcValue, ScalarValue, XlError};

use super::{
    require_bool, require_number, require_scalar, require_text, FunctionArg, FunctionDef,
    FunctionRegistry, ParamKind,
};

/// `ADDRESS(row_num, column_num, [abs_num=1], [a1=TRUE], [sheet_text])`
///
/// Returns a cell reference as text, given row and column numbers.
///
/// - `abs_num`: 1 (default) = absolute ($A$1), 2 = row absolute ($A1),
///   3 = column absolute (A$1), 4 = relative (A1)
/// - `a1`: TRUE (default) = A1 style, FALSE = R1C1 style
/// - `sheet_text`: optional sheet name to prepend
pub fn fn_address(args: &[FunctionArg<'_>], _ctx: &EvalContext<'_>) -> CalcValue {
    // Arg 0: row_num (1-based)
    let row_num = match args.first() {
        Some(a) => match require_number(a) {
            Ok(n) => {
                if !(1.0..=1048576.0).contains(&n) {
                    return CalcValue::error(XlError::Value);
                }
                n as u32
            }
            Err(e) => return e,
        },
        None => return CalcValue::error(XlError::Value),
    };

    // Arg 1: column_num (1-based)
    let col_num = match args.get(1) {
        Some(a) => match require_number(a) {
            Ok(n) => {
                if !(1.0..=16384.0).contains(&n) {
                    return CalcValue::error(XlError::Value);
                }
                n as u32
            }
            Err(e) => return e,
        },
        None => return CalcValue::error(XlError::Value),
    };

    // Arg 2: abs_num (1-4, default 1)
    let abs_num = match args.get(2) {
        Some(a) => match require_number(a) {
            Ok(n) => {
                let val = n as i32;
                if !(1..=4).contains(&val) {
                    return CalcValue::error(XlError::Value);
                }
                val
            }
            Err(e) => return e,
        },
        None => 1,
    };

    // Arg 3: a1 (bool, default TRUE)
    let a1_style = match args.get(3) {
        Some(a) => match require_bool(a) {
            Ok(b) => b,
            Err(e) => return e,
        },
        None => true,
    };

    // Arg 4: sheet_text (optional)
    let sheet_text = match args.get(4) {
        Some(a) => match require_text(a) {
            Ok(s) => Some(s),
            Err(e) => return e,
        },
        None => None,
    };

    let mut result = String::new();

    // Add sheet name if provided
    if let Some(ref sheet) = sheet_text {
        if sheet.contains(' ') || sheet.contains('!') {
            result.push('\'');
            result.push_str(sheet);
            result.push('\'');
        } else {
            result.push_str(sheet);
        }
        result.push('!');
    }

    if a1_style {
        // A1 style
        let col_absolute = abs_num == 1 || abs_num == 2;
        let row_absolute = abs_num == 1 || abs_num == 3;

        if col_absolute {
            result.push('$');
        }
        if let Some(letters) = column_index_to_letters(col_num) {
            result.push_str(&letters);
        }
        if row_absolute {
            result.push('$');
        }
        result.push_str(&row_num.to_string());
    } else {
        // R1C1 style
        result.push('R');
        match abs_num {
            1 | 2 => {
                // Row absolute
                result.push_str(&row_num.to_string());
            }
            3 | 4 => {
                // Row relative
                result.push('[');
                result.push_str(&row_num.to_string());
                result.push(']');
            }
            _ => unreachable!(),
        }
        result.push('C');
        match abs_num {
            1 | 3 => {
                // Column absolute
                result.push_str(&col_num.to_string());
            }
            2 | 4 => {
                // Column relative
                result.push('[');
                result.push_str(&col_num.to_string());
                result.push(']');
            }
            _ => unreachable!(),
        }
    }

    CalcValue::text(result)
}

/// `AREAS(reference)`
///
/// Returns the number of areas in a reference. Currently returns 1 for any
/// valid range (multi-area references not yet supported).
pub fn fn_areas(args: &[FunctionArg<'_>], _ctx: &EvalContext<'_>) -> CalcValue {
    match args.first() {
        Some(FunctionArg::Range { .. }) => CalcValue::number(1.0),
        Some(FunctionArg::Value(v)) => {
            // Check if it's an error
            if let ScalarValue::Error(e) = v.as_scalar() {
                CalcValue::error(*e)
            } else {
                CalcValue::error(XlError::Value)
            }
        }
        None => CalcValue::error(XlError::Value),
    }
}

/// `CHOOSE(index_num, value1, [value2], ...)`
///
/// Chooses a value from a list based on index number (1-based).
/// Returns `#VALUE!` if index is out of bounds or invalid.
pub fn fn_choose(args: &[FunctionArg<'_>], _ctx: &EvalContext<'_>) -> CalcValue {
    if args.is_empty() {
        return CalcValue::error(XlError::Value);
    }

    // Arg 0: index_num
    let index = match require_number(&args[0]) {
        Ok(n) => {
            if n < 1.0 {
                return CalcValue::error(XlError::Value);
            }
            n as usize
        }
        Err(e) => return e,
    };

    // Check bounds
    if index > args.len() - 1 {
        return CalcValue::error(XlError::Value);
    }

    // Return the chosen value
    match &args[index] {
        FunctionArg::Value(v) => v.clone(),
        FunctionArg::Range { .. } => {
            // For simplicity, return #VALUE! for range arguments
            // (Excel can handle this with implicit intersection)
            CalcValue::error(XlError::Value)
        }
    }
}

/// `COLUMN([reference])`
///
/// Returns the column number of a reference. If called without arguments,
/// returns the column of the cell containing the formula.
pub fn fn_column(args: &[FunctionArg<'_>], ctx: &EvalContext<'_>) -> CalcValue {
    match args.first() {
        None => {
            // No argument: return the column of the formula cell
            CalcValue::number(ctx.formula_col as f64)
        }
        Some(FunctionArg::Range { range, .. }) => {
            // Return the first column of the range
            CalcValue::number(range.start_col as f64)
        }
        Some(FunctionArg::Value(v)) => {
            if let ScalarValue::Error(e) = v.as_scalar() {
                CalcValue::error(*e)
            } else {
                CalcValue::error(XlError::Value)
            }
        }
    }
}

/// `COLUMNS(array)`
///
/// Returns the number of columns in a reference or array.
pub fn fn_columns(args: &[FunctionArg<'_>], _ctx: &EvalContext<'_>) -> CalcValue {
    match args.first() {
        Some(FunctionArg::Range { range, .. }) => {
            let count = range.end_col - range.start_col + 1;
            CalcValue::number(count as f64)
        }
        Some(FunctionArg::Value(v)) => {
            if let ScalarValue::Error(e) = v.as_scalar() {
                CalcValue::error(*e)
            } else {
                // A single value has 1 column
                CalcValue::number(1.0)
            }
        }
        None => CalcValue::error(XlError::Value),
    }
}

/// `FORMULATEXT(reference)`
///
/// Returns the formula at the given reference as text.
/// Currently returns #N/A as formula text retrieval is not yet implemented.
pub fn fn_formulatext(args: &[FunctionArg<'_>], ctx: &EvalContext<'_>) -> CalcValue {
    match args.first() {
        Some(FunctionArg::Range { range, .. }) => {
            let sheet = range.sheet.as_deref();
            let row = range.start_row;
            let col = range.start_col;

            if let Some(formula) = ctx.provider.cell_formula(sheet, row, col) {
                CalcValue::text(formula)
            } else {
                CalcValue::error(XlError::Na)
            }
        }
        Some(FunctionArg::Value(v)) => {
            if let ScalarValue::Error(e) = v.as_scalar() {
                CalcValue::error(*e)
            } else {
                CalcValue::error(XlError::Na)
            }
        }
        None => CalcValue::error(XlError::Value),
    }
}

/// `GETPIVOTDATA(...)`
///
/// Returns data stored in a PivotTable report.
/// Not implemented - returns #REF! error.
pub fn fn_getpivotdata(_args: &[FunctionArg<'_>], _ctx: &EvalContext<'_>) -> CalcValue {
    // PivotTable support is out of scope for Phase 1
    CalcValue::error(XlError::Ref)
}

/// `HLOOKUP(lookup_value, table_array, row_index_num, [range_lookup=TRUE])`
///
/// Searches for a value in the top row of a table and returns a value in
/// the same column from a specified row.
///
/// - `range_lookup` TRUE (default): approximate match (sorted ascending)
/// - `range_lookup` FALSE: exact match
pub fn fn_hlookup(args: &[FunctionArg<'_>], _ctx: &EvalContext<'_>) -> CalcValue {
    // Arg 0: lookup_value (scalar)
    let mut lookup_val = match args.first() {
        Some(a) => require_scalar(a).clone(),
        None => return CalcValue::error(XlError::Value),
    };
    if let ScalarValue::Error(e) = &lookup_val {
        return CalcValue::error(*e);
    }
    // Convert blank to 0
    if lookup_val.is_blank() {
        lookup_val = ScalarValue::Number(0.0);
    }

    // Arg 1: table_array (range)
    let (range, rctx) = match args.get(1) {
        Some(FunctionArg::Range { range, ctx }) => (*range, *ctx),
        _ => return CalcValue::error(XlError::Value),
    };

    // Arg 2: row_index_num (scalar)
    let row_index = match args.get(2) {
        Some(a) => match require_number(a) {
            Ok(n) => n as i64,
            Err(e) => return e,
        },
        None => return CalcValue::error(XlError::Value),
    };
    if row_index < 1 {
        return CalcValue::error(XlError::Value);
    }
    let row_idx = row_index as u32;
    let table_height = range.end_row - range.start_row + 1;
    if row_idx > table_height {
        return CalcValue::error(XlError::Ref);
    }

    // Arg 3: range_lookup (scalar, default TRUE)
    let range_lookup = match args.get(3) {
        Some(a) => {
            let scalar = require_scalar(a);
            if let ScalarValue::Error(e) = scalar {
                return CalcValue::error(*e);
            }
            match scalar.to_bool() {
                ScalarValue::Bool(b) => b,
                ScalarValue::Error(e) => return CalcValue::error(e),
                _ => true,
            }
        }
        None => true,
    };

    let sheet = range.sheet.as_deref();
    let first_row = range.start_row;
    let result_row = range.start_row + row_idx - 1;

    if range_lookup {
        // Approximate match
        hlookup_approximate(&lookup_val, range, sheet, first_row, result_row, rctx)
    } else {
        // Exact match
        hlookup_exact(&lookup_val, range, sheet, first_row, result_row, rctx)
    }
}

/// Approximate match HLOOKUP (range_lookup=TRUE).
fn hlookup_approximate(
    lookup_val: &ScalarValue,
    range: &RangeRef,
    sheet: Option<&str>,
    first_row: u32,
    result_row: u32,
    ctx: &EvalContext<'_>,
) -> CalcValue {
    let mut best_col: Option<u32> = None;
    for col in range.start_col..=range.end_col {
        let cell_val = ctx.provider.cell_value(sheet, first_row, col);
        if cell_val.is_blank() || cell_val.is_error() {
            continue;
        }
        let ord = compare_values(&cell_val, lookup_val);
        match ord {
            std::cmp::Ordering::Equal => {
                let result = ctx.provider.cell_value(sheet, result_row, col);
                return CalcValue::Scalar(result);
            }
            std::cmp::Ordering::Less => {
                best_col = Some(col);
            }
            std::cmp::Ordering::Greater => {
                break;
            }
        }
    }
    match best_col {
        Some(col) => {
            let result = ctx.provider.cell_value(sheet, result_row, col);
            CalcValue::Scalar(result)
        }
        None => CalcValue::error(XlError::Na),
    }
}

/// Exact match HLOOKUP (range_lookup=FALSE).
fn hlookup_exact(
    lookup_val: &ScalarValue,
    range: &RangeRef,
    sheet: Option<&str>,
    first_row: u32,
    result_row: u32,
    ctx: &EvalContext<'_>,
) -> CalcValue {
    for col in range.start_col..=range.end_col {
        let cell_val = ctx.provider.cell_value(sheet, first_row, col);
        if compare_values(&cell_val, lookup_val) == std::cmp::Ordering::Equal {
            let result = ctx.provider.cell_value(sheet, result_row, col);
            return CalcValue::Scalar(result);
        }
    }
    CalcValue::error(XlError::Na)
}

/// `HYPERLINK(link_location, [friendly_name])`
///
/// Creates a hyperlink. Returns the friendly name if provided, otherwise
/// returns the link location.
pub fn fn_hyperlink(args: &[FunctionArg<'_>], _ctx: &EvalContext<'_>) -> CalcValue {
    // Arg 0: link_location
    let _link = match args.first() {
        Some(a) => match require_text(a) {
            Ok(s) => s,
            Err(e) => return e,
        },
        None => return CalcValue::error(XlError::Value),
    };

    // Arg 1: friendly_name (optional)
    match args.get(1) {
        Some(a) => {
            let scalar = require_scalar(a);
            if let ScalarValue::Error(e) = scalar {
                CalcValue::error(*e)
            } else {
                CalcValue::Scalar(scalar.clone())
            }
        }
        None => CalcValue::text(_link),
    }
}

/// `INDIRECT(ref_text, [a1=TRUE])`
///
/// Returns the reference specified by a text string.
/// Not fully implemented - returns #REF! error.
pub fn fn_indirect(_args: &[FunctionArg<'_>], _ctx: &EvalContext<'_>) -> CalcValue {
    // INDIRECT requires parsing reference text and evaluating it dynamically.
    // This needs special handling in the evaluator context.
    // For now, return #REF!
    CalcValue::error(XlError::Ref)
}

/// `LOOKUP(lookup_value, lookup_vector, [result_vector])`
///
/// Looks up a value in a vector. This is a simplified vector form.
/// Array form not yet supported.
pub fn fn_lookup(args: &[FunctionArg<'_>], _ctx: &EvalContext<'_>) -> CalcValue {
    // Arg 0: lookup_value
    let lookup_val = match args.first() {
        Some(a) => require_scalar(a).clone(),
        None => return CalcValue::error(XlError::Value),
    };
    if let ScalarValue::Error(e) = &lookup_val {
        return CalcValue::error(*e);
    }

    // Arg 1: lookup_vector (range)
    let (lookup_range, lookup_ctx) = match args.get(1) {
        Some(FunctionArg::Range { range, ctx }) => (*range, *ctx),
        _ => return CalcValue::error(XlError::Value),
    };

    // Arg 2: result_vector (range, optional)
    let result_info = args.get(2).and_then(|a| match a {
        FunctionArg::Range { range, ctx } => Some((*range, *ctx)),
        _ => None,
    });

    // Build lookup values
    let sheet = lookup_range.sheet.as_deref();
    let is_column = lookup_range.start_col == lookup_range.end_col;

    let lookup_values: Vec<ScalarValue> = if is_column {
        (lookup_range.start_row..=lookup_range.end_row)
            .map(|row| {
                lookup_ctx
                    .provider
                    .cell_value(sheet, row, lookup_range.start_col)
            })
            .collect()
    } else {
        (lookup_range.start_col..=lookup_range.end_col)
            .map(|col| {
                lookup_ctx
                    .provider
                    .cell_value(sheet, lookup_range.start_row, col)
            })
            .collect()
    };

    // Find the position (approximate match, largest value <= lookup_value)
    let mut best_pos: Option<usize> = None;
    for (i, val) in lookup_values.iter().enumerate() {
        if val.is_error() || val.is_blank() {
            continue;
        }
        let ord = compare_values(val, &lookup_val);
        match ord {
            std::cmp::Ordering::Equal => {
                best_pos = Some(i);
                break;
            }
            std::cmp::Ordering::Less => {
                best_pos = Some(i);
            }
            std::cmp::Ordering::Greater => {
                break;
            }
        }
    }

    let pos = match best_pos {
        Some(p) => p,
        None => return CalcValue::error(XlError::Na),
    };

    // Get result value
    if let Some((result_range, result_ctx)) = result_info {
        let result_sheet = result_range.sheet.as_deref();
        let result_is_column = result_range.start_col == result_range.end_col;

        if result_is_column {
            let row = result_range.start_row + pos as u32;
            if row > result_range.end_row {
                return CalcValue::error(XlError::Ref);
            }
            let val = result_ctx
                .provider
                .cell_value(result_sheet, row, result_range.start_col);
            CalcValue::Scalar(val)
        } else {
            let col = result_range.start_col + pos as u32;
            if col > result_range.end_col {
                return CalcValue::error(XlError::Ref);
            }
            let val = result_ctx
                .provider
                .cell_value(result_sheet, result_range.start_row, col);
            CalcValue::Scalar(val)
        }
    } else {
        // Return from lookup vector itself
        CalcValue::Scalar(lookup_values[pos].clone())
    }
}

/// `OFFSET(reference, rows, cols, [height], [width])`
///
/// Returns a reference offset from a starting reference.
/// Not fully implemented - returns #REF! error.
pub fn fn_offset(_args: &[FunctionArg<'_>], _ctx: &EvalContext<'_>) -> CalcValue {
    // OFFSET requires returning a reference, not a value.
    // This needs special handling in the evaluator to return a new RangeRef.
    // For now, return #REF!
    CalcValue::error(XlError::Ref)
}

/// `ROW([reference])`
///
/// Returns the row number of a reference. If called without arguments,
/// returns the row of the cell containing the formula.
pub fn fn_row(args: &[FunctionArg<'_>], ctx: &EvalContext<'_>) -> CalcValue {
    match args.first() {
        None => {
            // No argument: return the row of the formula cell
            CalcValue::number(ctx.formula_row as f64)
        }
        Some(FunctionArg::Range { range, .. }) => {
            // Return the first row of the range
            CalcValue::number(range.start_row as f64)
        }
        Some(FunctionArg::Value(v)) => {
            if let ScalarValue::Error(e) = v.as_scalar() {
                CalcValue::error(*e)
            } else {
                CalcValue::error(XlError::Value)
            }
        }
    }
}

/// `ROWS(array)`
///
/// Returns the number of rows in a reference or array.
pub fn fn_rows(args: &[FunctionArg<'_>], _ctx: &EvalContext<'_>) -> CalcValue {
    match args.first() {
        Some(FunctionArg::Range { range, .. }) => {
            let count = range.end_row - range.start_row + 1;
            CalcValue::number(count as f64)
        }
        Some(FunctionArg::Value(v)) => {
            if let ScalarValue::Error(e) = v.as_scalar() {
                CalcValue::error(*e)
            } else {
                // A single value has 1 row
                CalcValue::number(1.0)
            }
        }
        None => CalcValue::error(XlError::Value),
    }
}

/// `RTD(...)`
///
/// Retrieves real-time data from a COM automation program.
/// Not implemented - returns #N/A error.
pub fn fn_rtd(_args: &[FunctionArg<'_>], _ctx: &EvalContext<'_>) -> CalcValue {
    // RTD support is out of scope
    CalcValue::error(XlError::Na)
}

/// `TRANSPOSE(array)`
///
/// Returns the transpose of an array. Currently returns #VALUE! as array
/// operations are not yet fully supported.
pub fn fn_transpose(args: &[FunctionArg<'_>], _ctx: &EvalContext<'_>) -> CalcValue {
    match args.first() {
        Some(FunctionArg::Range { .. }) => {
            // Array operations not yet supported in Phase 1
            CalcValue::error(XlError::Value)
        }
        Some(FunctionArg::Value(v)) => {
            if let ScalarValue::Error(e) = v.as_scalar() {
                CalcValue::error(*e)
            } else {
                // Transposing a single value returns the value
                v.clone()
            }
        }
        None => CalcValue::error(XlError::Value),
    }
}

/// `VLOOKUP(lookup_value, table_range, col_index, [range_lookup=TRUE])`
///
/// Searches for a value in the first column of a range and returns a value in
/// the same row from a specified column.
///
/// - `range_lookup` TRUE (default): finds the largest value <= `lookup_value`.
///   The first column must be sorted ascending.
/// - `range_lookup` FALSE: exact match. Not found returns `#N/A`.
/// - `col_index < 1` returns `#VALUE!`.
/// - `col_index > table width` returns `#REF!`.
pub fn fn_vlookup(args: &[FunctionArg<'_>], _ctx: &EvalContext<'_>) -> CalcValue {
    // Arg 0: lookup_value (scalar)
    let lookup_val = match args.first() {
        Some(a) => require_scalar(a).clone(),
        None => return CalcValue::error(XlError::Value),
    };
    if let ScalarValue::Error(e) = &lookup_val {
        return CalcValue::error(*e);
    }

    // Arg 1: table_range (range)
    let (range, rctx) = match args.get(1) {
        Some(FunctionArg::Range { range, ctx }) => (*range, *ctx),
        _ => return CalcValue::error(XlError::Value),
    };

    // Arg 2: col_index (scalar)
    let col_index = match args.get(2) {
        Some(a) => match require_number(a) {
            Ok(n) => n as i64,
            Err(e) => return e,
        },
        None => return CalcValue::error(XlError::Value),
    };
    if col_index < 1 {
        return CalcValue::error(XlError::Value);
    }
    let col_idx = col_index as u32;
    let table_width = range.end_col - range.start_col + 1;
    if col_idx > table_width {
        return CalcValue::error(XlError::Ref);
    }

    // Arg 3: range_lookup (scalar, default TRUE)
    let range_lookup = match args.get(3) {
        Some(a) => {
            let scalar = require_scalar(a);
            if let ScalarValue::Error(e) = scalar {
                return CalcValue::error(*e);
            }
            match scalar.to_bool() {
                ScalarValue::Bool(b) => b,
                ScalarValue::Error(e) => return CalcValue::error(e),
                _ => true,
            }
        }
        None => true,
    };

    let sheet = range.sheet.as_deref();
    let first_col = range.start_col;
    let result_col = range.start_col + col_idx - 1;

    if range_lookup {
        // Approximate match: find largest value <= lookup_value
        // Assumes first column is sorted ascending
        vlookup_approximate(&lookup_val, range, sheet, first_col, result_col, rctx)
    } else {
        // Exact match: scan all rows
        vlookup_exact(&lookup_val, range, sheet, first_col, result_col, rctx)
    }
}

/// Approximate match VLOOKUP (range_lookup=TRUE).
fn vlookup_approximate(
    lookup_val: &ScalarValue,
    range: &RangeRef,
    sheet: Option<&str>,
    first_col: u32,
    result_col: u32,
    ctx: &EvalContext<'_>,
) -> CalcValue {
    let mut best_row: Option<u32> = None;
    for row in range.start_row..=range.end_row {
        let cell_val = ctx.provider.cell_value(sheet, row, first_col);
        if cell_val.is_blank() || cell_val.is_error() {
            continue;
        }
        let ord = compare_values(&cell_val, lookup_val);
        match ord {
            std::cmp::Ordering::Equal => {
                // Exact match found, return immediately
                let result = ctx.provider.cell_value(sheet, row, result_col);
                return CalcValue::Scalar(result);
            }
            std::cmp::Ordering::Less => {
                best_row = Some(row);
            }
            std::cmp::Ordering::Greater => {
                // Since data is sorted, we've gone past the lookup value
                break;
            }
        }
    }
    match best_row {
        Some(row) => {
            let result = ctx.provider.cell_value(sheet, row, result_col);
            CalcValue::Scalar(result)
        }
        None => CalcValue::error(XlError::Na),
    }
}

/// Exact match VLOOKUP (range_lookup=FALSE).
fn vlookup_exact(
    lookup_val: &ScalarValue,
    range: &RangeRef,
    sheet: Option<&str>,
    first_col: u32,
    result_col: u32,
    ctx: &EvalContext<'_>,
) -> CalcValue {
    for row in range.start_row..=range.end_row {
        let cell_val = ctx.provider.cell_value(sheet, row, first_col);
        if compare_values(&cell_val, lookup_val) == std::cmp::Ordering::Equal {
            let result = ctx.provider.cell_value(sheet, row, result_col);
            return CalcValue::Scalar(result);
        }
    }
    CalcValue::error(XlError::Na)
}

/// `INDEX(range, row_num, [col_num=1])`
///
/// Returns the value at a given position within a range. Indices are 1-based.
/// Out-of-bounds returns `#REF!`. `row_num=0` returns `#VALUE!` (array form
/// not supported in Phase 1).
pub fn fn_index(args: &[FunctionArg<'_>], _ctx: &EvalContext<'_>) -> CalcValue {
    // Arg 0: range
    let (range, rctx) = match args.first() {
        Some(FunctionArg::Range { range, ctx }) => (*range, *ctx),
        _ => return CalcValue::error(XlError::Value),
    };

    // Arg 1: row_num (scalar)
    let row_num = match args.get(1) {
        Some(a) => match require_number(a) {
            Ok(n) => n as i64,
            Err(e) => return e,
        },
        None => return CalcValue::error(XlError::Value),
    };

    // Arg 2: col_num (scalar, default 1)
    let col_num = match args.get(2) {
        Some(a) => match require_number(a) {
            Ok(n) => n as i64,
            Err(e) => return e,
        },
        None => 1,
    };

    // row_num=0 or col_num=0 would return an entire row/column (array)
    // Not supported in Phase 1
    if row_num == 0 || col_num == 0 {
        return CalcValue::error(XlError::Value);
    }
    if row_num < 0 || col_num < 0 {
        return CalcValue::error(XlError::Value);
    }

    let range_rows = range.end_row - range.start_row + 1;
    let range_cols = range.end_col - range.start_col + 1;

    if row_num as u32 > range_rows || col_num as u32 > range_cols {
        return CalcValue::error(XlError::Ref);
    }

    let target_row = range.start_row + (row_num as u32) - 1;
    let target_col = range.start_col + (col_num as u32) - 1;
    let sheet = range.sheet.as_deref();
    let result = rctx.provider.cell_value(sheet, target_row, target_col);
    CalcValue::Scalar(result)
}

/// `MATCH(lookup_value, lookup_range, [match_type=1])`
///
/// Returns the relative position (1-based) of an item in a range.
///
/// - `match_type = 1`: largest value <= `lookup_value` (range sorted ascending)
/// - `match_type = 0`: exact match (first found)
/// - `match_type = -1`: smallest value >= `lookup_value` (range sorted descending)
/// - Not found returns `#N/A`.
pub fn fn_match(args: &[FunctionArg<'_>], _ctx: &EvalContext<'_>) -> CalcValue {
    // Arg 0: lookup_value (scalar)
    let lookup_val = match args.first() {
        Some(a) => {
            let s = require_scalar(a).clone();
            if let ScalarValue::Error(e) = &s {
                return CalcValue::error(*e);
            }
            s
        }
        None => return CalcValue::error(XlError::Value),
    };

    // Arg 1: lookup_range (range)
    let (range, rctx) = match args.get(1) {
        Some(FunctionArg::Range { range, ctx }) => (*range, *ctx),
        _ => return CalcValue::error(XlError::Value),
    };

    // Arg 2: match_type (scalar, default 1)
    let match_type = match args.get(2) {
        Some(a) => match require_number(a) {
            Ok(n) => n as i64,
            Err(e) => return e,
        },
        None => 1,
    };

    // Build the list of values from the range (1-D)
    let sheet = range.sheet.as_deref();
    let is_column = range.start_col == range.end_col;

    let values: Vec<ScalarValue> = if is_column {
        // Vertical range
        (range.start_row..=range.end_row)
            .map(|row| rctx.provider.cell_value(sheet, row, range.start_col))
            .collect()
    } else {
        // Horizontal range
        (range.start_col..=range.end_col)
            .map(|col| rctx.provider.cell_value(sheet, range.start_row, col))
            .collect()
    };

    match match_type {
        0 => match_exact(&lookup_val, &values),
        1 => match_ascending(&lookup_val, &values),
        -1 => match_descending(&lookup_val, &values),
        _ => {
            // Non-standard match_type: treat positive as 1, negative as -1
            if match_type > 0 {
                match_ascending(&lookup_val, &values)
            } else {
                match_descending(&lookup_val, &values)
            }
        }
    }
}

/// Exact match (match_type=0): returns 1-based position of first exact match.
fn match_exact(lookup: &ScalarValue, values: &[ScalarValue]) -> CalcValue {
    for (i, val) in values.iter().enumerate() {
        if val.is_error() || val.is_blank() {
            continue;
        }
        if compare_values(val, lookup) == std::cmp::Ordering::Equal {
            return CalcValue::number((i + 1) as f64);
        }
    }
    CalcValue::error(XlError::Na)
}

/// Ascending match (match_type=1): largest value <= lookup_value.
fn match_ascending(lookup: &ScalarValue, values: &[ScalarValue]) -> CalcValue {
    let mut best_pos: Option<usize> = None;
    for (i, val) in values.iter().enumerate() {
        if val.is_error() || val.is_blank() {
            continue;
        }
        let ord = compare_values(val, lookup);
        match ord {
            std::cmp::Ordering::Equal => return CalcValue::number((i + 1) as f64),
            std::cmp::Ordering::Less => {
                best_pos = Some(i);
            }
            std::cmp::Ordering::Greater => {
                break;
            }
        }
    }
    match best_pos {
        Some(pos) => CalcValue::number((pos + 1) as f64),
        None => CalcValue::error(XlError::Na),
    }
}

/// Descending match (match_type=-1): smallest value >= lookup_value.
fn match_descending(lookup: &ScalarValue, values: &[ScalarValue]) -> CalcValue {
    let mut best_pos: Option<usize> = None;
    for (i, val) in values.iter().enumerate() {
        if val.is_error() || val.is_blank() {
            continue;
        }
        let ord = compare_values(val, lookup);
        match ord {
            std::cmp::Ordering::Equal => return CalcValue::number((i + 1) as f64),
            std::cmp::Ordering::Greater => {
                best_pos = Some(i);
            }
            std::cmp::Ordering::Less => {
                break;
            }
        }
    }
    match best_pos {
        Some(pos) => CalcValue::number((pos + 1) as f64),
        None => CalcValue::error(XlError::Na),
    }
}

/// Registers all lookup functions into the given registry.
pub fn register(registry: &mut FunctionRegistry) {
    registry.register(FunctionDef {
        name: "ADDRESS",
        min_args: 2,
        max_args: 5,
        param_kinds: &[
            ParamKind::Scalar,
            ParamKind::Scalar,
            ParamKind::Scalar,
            ParamKind::Scalar,
            ParamKind::Scalar,
        ],
        func: fn_address,
    });
    registry.register(FunctionDef {
        name: "AREAS",
        min_args: 1,
        max_args: 1,
        param_kinds: &[ParamKind::Range],
        func: fn_areas,
    });
    registry.register(FunctionDef {
        name: "CHOOSE",
        min_args: 2,
        max_args: usize::MAX,
        param_kinds: &[ParamKind::Scalar],
        func: fn_choose,
    });
    registry.register(FunctionDef {
        name: "COLUMN",
        min_args: 0,
        max_args: 1,
        param_kinds: &[ParamKind::Range],
        func: fn_column,
    });
    registry.register(FunctionDef {
        name: "COLUMNS",
        min_args: 1,
        max_args: 1,
        param_kinds: &[ParamKind::Range],
        func: fn_columns,
    });
    registry.register(FunctionDef {
        name: "FORMULATEXT",
        min_args: 1,
        max_args: 1,
        param_kinds: &[ParamKind::Range],
        func: fn_formulatext,
    });
    registry.register(FunctionDef {
        name: "GETPIVOTDATA",
        min_args: 2,
        max_args: usize::MAX,
        param_kinds: &[ParamKind::Scalar],
        func: fn_getpivotdata,
    });
    registry.register(FunctionDef {
        name: "HLOOKUP",
        min_args: 3,
        max_args: 4,
        param_kinds: &[
            ParamKind::Scalar,
            ParamKind::Range,
            ParamKind::Scalar,
            ParamKind::Scalar,
        ],
        func: fn_hlookup,
    });
    registry.register(FunctionDef {
        name: "HYPERLINK",
        min_args: 1,
        max_args: 2,
        param_kinds: &[ParamKind::Scalar, ParamKind::Scalar],
        func: fn_hyperlink,
    });
    registry.register(FunctionDef {
        name: "INDEX",
        min_args: 2,
        max_args: 3,
        param_kinds: &[ParamKind::Range, ParamKind::Scalar, ParamKind::Scalar],
        func: fn_index,
    });
    registry.register(FunctionDef {
        name: "INDIRECT",
        min_args: 1,
        max_args: 2,
        param_kinds: &[ParamKind::Scalar, ParamKind::Scalar],
        func: fn_indirect,
    });
    registry.register(FunctionDef {
        name: "LOOKUP",
        min_args: 2,
        max_args: 3,
        param_kinds: &[ParamKind::Scalar, ParamKind::Range, ParamKind::Range],
        func: fn_lookup,
    });
    registry.register(FunctionDef {
        name: "MATCH",
        min_args: 2,
        max_args: 3,
        param_kinds: &[ParamKind::Scalar, ParamKind::Range, ParamKind::Scalar],
        func: fn_match,
    });
    registry.register(FunctionDef {
        name: "OFFSET",
        min_args: 3,
        max_args: 5,
        param_kinds: &[
            ParamKind::Range,
            ParamKind::Scalar,
            ParamKind::Scalar,
            ParamKind::Scalar,
            ParamKind::Scalar,
        ],
        func: fn_offset,
    });
    registry.register(FunctionDef {
        name: "ROW",
        min_args: 0,
        max_args: 1,
        param_kinds: &[ParamKind::Range],
        func: fn_row,
    });
    registry.register(FunctionDef {
        name: "ROWS",
        min_args: 1,
        max_args: 1,
        param_kinds: &[ParamKind::Range],
        func: fn_rows,
    });
    registry.register(FunctionDef {
        name: "RTD",
        min_args: 1,
        max_args: usize::MAX,
        param_kinds: &[ParamKind::Scalar],
        func: fn_rtd,
    });
    registry.register(FunctionDef {
        name: "TRANSPOSE",
        min_args: 1,
        max_args: 1,
        param_kinds: &[ParamKind::Range],
        func: fn_transpose,
    });
    registry.register(FunctionDef {
        name: "VLOOKUP",
        min_args: 3,
        max_args: 4,
        param_kinds: &[
            ParamKind::Scalar,
            ParamKind::Range,
            ParamKind::Scalar,
            ParamKind::Scalar,
        ],
        func: fn_vlookup,
    });
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::context::CellDataProvider;
    use pretty_assertions::assert_eq;

    struct TestProvider;
    impl CellDataProvider for TestProvider {
        fn cell_value(&self, _sheet: Option<&str>, row: u32, col: u32) -> ScalarValue {
            // Simple test data for lookup functions
            match (row, col) {
                (1, 1) => ScalarValue::Number(1.0),
                (1, 2) => ScalarValue::Text("A".to_string()),
                (2, 1) => ScalarValue::Number(2.0),
                (2, 2) => ScalarValue::Text("B".to_string()),
                (3, 1) => ScalarValue::Number(3.0),
                (3, 2) => ScalarValue::Text("C".to_string()),
                _ => ScalarValue::Blank,
            }
        }
    }

    #[test]
    fn test_address() {
        let provider = TestProvider;
        let ctx = EvalContext::new(&provider, None::<String>, 1, 1);

        // Basic A1 style
        let args = vec![
            FunctionArg::Value(CalcValue::number(1.0)),
            FunctionArg::Value(CalcValue::number(1.0)),
        ];
        let result = fn_address(&args, &ctx);
        assert_eq!(result, CalcValue::text("$A$1"));

        // With abs_num = 4 (relative)
        let args = vec![
            FunctionArg::Value(CalcValue::number(2.0)),
            FunctionArg::Value(CalcValue::number(3.0)),
            FunctionArg::Value(CalcValue::number(4.0)),
        ];
        let result = fn_address(&args, &ctx);
        assert_eq!(result, CalcValue::text("C2"));

        // With sheet name
        let args = vec![
            FunctionArg::Value(CalcValue::number(5.0)),
            FunctionArg::Value(CalcValue::number(10.0)),
            FunctionArg::Value(CalcValue::number(1.0)),
            FunctionArg::Value(CalcValue::bool(true)),
            FunctionArg::Value(CalcValue::text("Sheet1")),
        ];
        let result = fn_address(&args, &ctx);
        assert_eq!(result, CalcValue::text("Sheet1!$J$5"));
    }

    #[test]
    fn test_areas() {
        let provider = TestProvider;
        let ctx = EvalContext::new(&provider, None::<String>, 1, 1);
        let range = RangeRef {
            sheet: None,
            start_row: 1,
            start_col: 1,
            end_row: 3,
            end_col: 2,
        };

        let args = vec![FunctionArg::Range {
            range: &range,
            ctx: &ctx,
        }];
        let result = fn_areas(&args, &ctx);
        assert_eq!(result, CalcValue::number(1.0));
    }

    #[test]
    fn test_choose() {
        let provider = TestProvider;
        let ctx = EvalContext::new(&provider, None::<String>, 1, 1);

        let args = vec![
            FunctionArg::Value(CalcValue::number(2.0)),
            FunctionArg::Value(CalcValue::text("First")),
            FunctionArg::Value(CalcValue::text("Second")),
            FunctionArg::Value(CalcValue::text("Third")),
        ];
        let result = fn_choose(&args, &ctx);
        assert_eq!(result, CalcValue::text("Second"));

        // Out of bounds
        let args = vec![
            FunctionArg::Value(CalcValue::number(5.0)),
            FunctionArg::Value(CalcValue::text("First")),
        ];
        let result = fn_choose(&args, &ctx);
        assert_eq!(result, CalcValue::error(XlError::Value));
    }

    #[test]
    fn test_column() {
        let provider = TestProvider;
        let ctx = EvalContext::new(&provider, None::<String>, 5, 3);

        // No argument - returns formula cell column
        let args = vec![];
        let result = fn_column(&args, &ctx);
        assert_eq!(result, CalcValue::number(3.0));

        // With range argument
        let range = RangeRef {
            sheet: None,
            start_row: 1,
            start_col: 5,
            end_row: 10,
            end_col: 10,
        };
        let args = vec![FunctionArg::Range {
            range: &range,
            ctx: &ctx,
        }];
        let result = fn_column(&args, &ctx);
        assert_eq!(result, CalcValue::number(5.0));
    }

    #[test]
    fn test_columns() {
        let provider = TestProvider;
        let ctx = EvalContext::new(&provider, None::<String>, 1, 1);
        let range = RangeRef {
            sheet: None,
            start_row: 1,
            start_col: 1,
            end_row: 3,
            end_col: 5,
        };

        let args = vec![FunctionArg::Range {
            range: &range,
            ctx: &ctx,
        }];
        let result = fn_columns(&args, &ctx);
        assert_eq!(result, CalcValue::number(5.0));
    }

    #[test]
    fn test_hlookup() {
        let provider = TestProvider;
        let ctx = EvalContext::new(&provider, None::<String>, 1, 1);
        let range = RangeRef {
            sheet: None,
            start_row: 1,
            start_col: 1,
            end_row: 3,
            end_col: 2,
        };

        // Exact match
        let args = vec![
            FunctionArg::Value(CalcValue::number(1.0)),
            FunctionArg::Range {
                range: &range,
                ctx: &ctx,
            },
            FunctionArg::Value(CalcValue::number(2.0)),
            FunctionArg::Value(CalcValue::bool(false)),
        ];
        let result = fn_hlookup(&args, &ctx);
        assert_eq!(result, CalcValue::number(2.0));
    }

    #[test]
    fn test_hyperlink() {
        let provider = TestProvider;
        let ctx = EvalContext::new(&provider, None::<String>, 1, 1);

        // With friendly name
        let args = vec![
            FunctionArg::Value(CalcValue::text("https://example.com")),
            FunctionArg::Value(CalcValue::text("Example")),
        ];
        let result = fn_hyperlink(&args, &ctx);
        assert_eq!(result, CalcValue::text("Example"));

        // Without friendly name
        let args = vec![FunctionArg::Value(CalcValue::text("https://example.com"))];
        let result = fn_hyperlink(&args, &ctx);
        assert_eq!(result, CalcValue::text("https://example.com"));
    }

    #[test]
    fn test_lookup() {
        let provider = TestProvider;
        let ctx = EvalContext::new(&provider, None::<String>, 1, 1);
        let lookup_range = RangeRef {
            sheet: None,
            start_row: 1,
            start_col: 1,
            end_row: 3,
            end_col: 1,
        };
        let result_range = RangeRef {
            sheet: None,
            start_row: 1,
            start_col: 2,
            end_row: 3,
            end_col: 2,
        };

        // Two-vector form
        let args = vec![
            FunctionArg::Value(CalcValue::number(2.0)),
            FunctionArg::Range {
                range: &lookup_range,
                ctx: &ctx,
            },
            FunctionArg::Range {
                range: &result_range,
                ctx: &ctx,
            },
        ];
        let result = fn_lookup(&args, &ctx);
        assert_eq!(result, CalcValue::text("B"));
    }

    #[test]
    fn test_row() {
        let provider = TestProvider;
        let ctx = EvalContext::new(&provider, None::<String>, 7, 2);

        // No argument - returns formula cell row
        let args = vec![];
        let result = fn_row(&args, &ctx);
        assert_eq!(result, CalcValue::number(7.0));

        // With range argument
        let range = RangeRef {
            sheet: None,
            start_row: 10,
            start_col: 1,
            end_row: 20,
            end_col: 5,
        };
        let args = vec![FunctionArg::Range {
            range: &range,
            ctx: &ctx,
        }];
        let result = fn_row(&args, &ctx);
        assert_eq!(result, CalcValue::number(10.0));
    }

    #[test]
    fn test_rows() {
        let provider = TestProvider;
        let ctx = EvalContext::new(&provider, None::<String>, 1, 1);
        let range = RangeRef {
            sheet: None,
            start_row: 1,
            start_col: 1,
            end_row: 10,
            end_col: 2,
        };

        let args = vec![FunctionArg::Range {
            range: &range,
            ctx: &ctx,
        }];
        let result = fn_rows(&args, &ctx);
        assert_eq!(result, CalcValue::number(10.0));
    }
}
