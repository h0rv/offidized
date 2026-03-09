use std::str::FromStr;

use crate::cell::{normalize_cell_reference, Cell, CellValue};
use crate::error::{Result, XlsxError};
use crate::worksheet::Worksheet;
use offidized_formula::{
    ast::{BinaryOp, Expr, UnaryOp},
    lexer, parser, CellRef, RangeRef,
};

/// Adjusts a formula by offsetting relative cell references.
///
/// This is used during copy/move operations. Relative references (e.g., `A1`)
/// are adjusted by the given column and row offsets, while absolute references
/// (e.g., `$A$1`) remain unchanged.
fn adjust_formula(formula: &str, col_offset: i64, row_offset: i64) -> Result<String> {
    // Strip leading '=' if present
    let formula = formula.strip_prefix('=').unwrap_or(formula);

    // Parse the formula
    let tokens = lexer::tokenize(formula)
        .map_err(|e| XlsxError::InvalidFormula(format!("Tokenize error: {e}")))?;
    let expr = parser::parse(tokens)
        .map_err(|e| XlsxError::InvalidFormula(format!("Parse error: {e}")))?;

    // Adjust references in the AST
    let adjusted = adjust_expr(&expr, col_offset, row_offset)?;

    // Serialize back to string
    Ok(expr_to_string(&adjusted))
}

/// Recursively adjusts cell and range references in an expression.
fn adjust_expr(expr: &Expr, col_offset: i64, row_offset: i64) -> Result<Expr> {
    match expr {
        Expr::CellRef(cell_ref) => {
            let adjusted = adjust_cell_ref(cell_ref, col_offset, row_offset)?;
            Ok(Expr::CellRef(adjusted))
        }
        Expr::RangeRef(range_ref) => {
            let adjusted = adjust_range_ref(range_ref, col_offset, row_offset)?;
            Ok(Expr::RangeRef(adjusted))
        }
        Expr::Unary { op, expr: inner } => {
            let adjusted_inner = adjust_expr(inner, col_offset, row_offset)?;
            Ok(Expr::Unary {
                op: *op,
                expr: Box::new(adjusted_inner),
            })
        }
        Expr::Binary { op, left, right } => {
            let adjusted_left = adjust_expr(left, col_offset, row_offset)?;
            let adjusted_right = adjust_expr(right, col_offset, row_offset)?;
            Ok(Expr::Binary {
                op: *op,
                left: Box::new(adjusted_left),
                right: Box::new(adjusted_right),
            })
        }
        Expr::Function { name, args } => {
            let adjusted_args: Result<Vec<_>> = args
                .iter()
                .map(|arg| adjust_expr(arg, col_offset, row_offset))
                .collect();
            Ok(Expr::Function {
                name: name.clone(),
                args: adjusted_args?,
            })
        }
        Expr::Literal(_) | Expr::Name(_) => Ok(expr.clone()),
    }
}

/// Adjusts a single cell reference.
fn adjust_cell_ref(cell_ref: &CellRef, col_offset: i64, row_offset: i64) -> Result<CellRef> {
    let new_col = if cell_ref.col_absolute {
        cell_ref.col
    } else {
        apply_offset(cell_ref.col, col_offset, "column")?
    };

    let new_row = if cell_ref.row_absolute {
        cell_ref.row
    } else {
        apply_offset(cell_ref.row, row_offset, "row")?
    };

    Ok(CellRef {
        sheet: cell_ref.sheet.clone(),
        col: new_col,
        row: new_row,
        col_absolute: cell_ref.col_absolute,
        row_absolute: cell_ref.row_absolute,
    })
}

/// Adjusts a range reference.
fn adjust_range_ref(range_ref: &RangeRef, col_offset: i64, row_offset: i64) -> Result<RangeRef> {
    // Note: RangeRef doesn't have absolute flags, so we always adjust
    // This matches Excel's behavior for range references in formulas
    let new_start_col = apply_offset(range_ref.start_col, col_offset, "start column")?;
    let new_start_row = apply_offset(range_ref.start_row, row_offset, "start row")?;
    let new_end_col = apply_offset(range_ref.end_col, col_offset, "end column")?;
    let new_end_row = apply_offset(range_ref.end_row, row_offset, "end row")?;

    Ok(RangeRef {
        sheet: range_ref.sheet.clone(),
        start_col: new_start_col,
        start_row: new_start_row,
        end_col: new_end_col,
        end_row: new_end_row,
    })
}

/// Applies an offset to a 1-based index, checking for overflow.
fn apply_offset(index: u32, offset: i64, context: &str) -> Result<u32> {
    let new_index = index as i64 + offset;
    if new_index < 1 {
        return Err(XlsxError::InvalidFormula(format!(
            "{context} offset would result in invalid index: {new_index}"
        )));
    }
    u32::try_from(new_index)
        .map_err(|_| XlsxError::InvalidFormula(format!("{context} offset overflow: {new_index}")))
}

/// Converts an expression AST back to a formula string (without leading '=').
fn expr_to_string(expr: &Expr) -> String {
    match expr {
        Expr::Literal(val) => format!("{val}"),
        Expr::CellRef(cell_ref) => format!("{cell_ref}"),
        Expr::RangeRef(range_ref) => format!("{range_ref}"),
        Expr::Unary { op, expr: inner } => {
            let inner_str = expr_to_string(inner);
            match op {
                UnaryOp::Plus => format!("+{inner_str}"),
                UnaryOp::Negate => format!("-{inner_str}"),
                UnaryOp::Percent => format!("{inner_str}%"),
            }
        }
        Expr::Binary { op, left, right } => {
            let left_str = expr_to_string(left);
            let right_str = expr_to_string(right);
            let op_str = match op {
                BinaryOp::Add => "+",
                BinaryOp::Sub => "-",
                BinaryOp::Mul => "*",
                BinaryOp::Div => "/",
                BinaryOp::Pow => "^",
                BinaryOp::Concat => "&",
                BinaryOp::Eq => "=",
                BinaryOp::Ne => "<>",
                BinaryOp::Lt => "<",
                BinaryOp::Le => "<=",
                BinaryOp::Gt => ">",
                BinaryOp::Ge => ">=",
                BinaryOp::Range => ":",
            };
            format!("{left_str}{op_str}{right_str}")
        }
        Expr::Function { name, args } => {
            let args_str = args
                .iter()
                .map(expr_to_string)
                .collect::<Vec<_>>()
                .join(",");
            format!("{name}({args_str})")
        }
        Expr::Name(name) => name.clone(),
    }
}

/// A rectangular cell range identified by start/end references.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct CellRange {
    start: String,
    end: String,
    start_column: u32,
    start_row: u32,
    end_column: u32,
    end_row: u32,
}

impl CellRange {
    pub fn new(start: &str, end: &str) -> Result<Self> {
        let (start_column, start_row) = parse_cell_reference(start)?;
        let (end_column, end_row) = parse_cell_reference(end)?;

        let normalized_start_column = start_column.min(end_column);
        let normalized_end_column = start_column.max(end_column);
        let normalized_start_row = start_row.min(end_row);
        let normalized_end_row = start_row.max(end_row);

        let normalized_start = build_cell_reference(normalized_start_column, normalized_start_row)?;
        let normalized_end = build_cell_reference(normalized_end_column, normalized_end_row)?;

        Ok(Self {
            start: normalized_start,
            end: normalized_end,
            start_column: normalized_start_column,
            start_row: normalized_start_row,
            end_column: normalized_end_column,
            end_row: normalized_end_row,
        })
    }

    /// Parses a range in A1 notation, e.g. `A1:B3` or `C7`.
    pub fn parse(range: &str) -> Result<Self> {
        let trimmed = range.trim();
        if trimmed.is_empty() {
            return Err(XlsxError::InvalidCellReference(range.to_string()));
        }

        let mut parts = trimmed.split(':');
        let start = parts.next();
        let end = parts.next();
        let extra = parts.next();

        match (start, end, extra) {
            (Some(single), None, None) => Self::new(single.trim(), single.trim()),
            (Some(left), Some(right), None) => Self::new(left.trim(), right.trim()),
            _ => Err(XlsxError::InvalidCellReference(trimmed.to_string())),
        }
    }

    pub fn start(&self) -> &str {
        &self.start
    }

    pub fn end(&self) -> &str {
        &self.end
    }

    pub fn width(&self) -> u32 {
        self.end_column - self.start_column + 1
    }

    pub fn height(&self) -> u32 {
        self.end_row - self.start_row + 1
    }

    pub fn contains(&self, reference: &str) -> bool {
        parse_cell_reference(reference)
            .map(|(column, row)| self.contains_coordinates(column, row))
            .unwrap_or(false)
    }

    /// Iterates all cell references in row-major order.
    pub fn iter(&self) -> CellRangeIter {
        CellRangeIter::new(
            self.start_column,
            self.start_row,
            self.end_column,
            self.end_row,
        )
    }

    /// Gets a cell from a worksheet when the provided reference is inside this range.
    pub fn get<'a>(&self, worksheet: &'a Worksheet, reference: &str) -> Option<&'a Cell> {
        if !self.contains(reference) {
            return None;
        }
        worksheet.cell(reference)
    }

    /// Sets a cell value when the provided reference is inside this range.
    ///
    /// Returns `Ok(true)` when a value was set and `Ok(false)` when the reference is outside
    /// this range.
    pub fn set_value(
        &self,
        worksheet: &mut Worksheet,
        reference: &str,
        value: impl Into<CellValue>,
    ) -> Result<bool> {
        let (column, row) = parse_cell_reference(reference)?;
        if !self.contains_coordinates(column, row) {
            return Ok(false);
        }

        worksheet.cell_mut(reference)?.set_value(value);
        Ok(true)
    }

    /// Returns the 1-based start column index.
    pub fn start_column(&self) -> u32 {
        self.start_column
    }

    /// Returns the 1-based start row index.
    pub fn start_row(&self) -> u32 {
        self.start_row
    }

    /// Returns the 1-based end column index.
    pub fn end_column(&self) -> u32 {
        self.end_column
    }

    /// Returns the 1-based end row index.
    pub fn end_row(&self) -> u32 {
        self.end_row
    }

    /// Returns a new range with rows and columns swapped.
    ///
    /// The start column becomes the start row and vice versa.
    pub fn transpose(&self) -> Result<Self> {
        let new_start = build_cell_reference(self.start_row, self.start_column)?;
        let new_end = build_cell_reference(self.end_row, self.end_column)?;
        Self::new(&new_start, &new_end)
    }

    fn contains_coordinates(&self, column: u32, row: u32) -> bool {
        (self.start_column..=self.end_column).contains(&column)
            && (self.start_row..=self.end_row).contains(&row)
    }

    /// Copies this range to a destination starting cell.
    ///
    /// Copies cell values, formulas (with adjusted references), and styles.
    /// Note: This does not copy merged cells, conditional formatting, or data validation yet.
    ///
    /// # Arguments
    /// * `worksheet` - The worksheet to copy cells within (source and destination must be same sheet)
    /// * `dest_start` - The top-left cell reference where the range should be copied to (e.g., "D5")
    ///
    /// # Returns
    /// `Ok(())` if the copy succeeded, or an error if references are invalid.
    pub fn copy_to(&self, worksheet: &mut Worksheet, dest_start: &str) -> Result<()> {
        let (dest_start_col, dest_start_row) = parse_cell_reference(dest_start)?;

        let col_offset = dest_start_col as i64 - self.start_column as i64;
        let row_offset = dest_start_row as i64 - self.start_row as i64;

        // First pass: collect all cell data to avoid borrow checker issues
        let mut cells_to_copy = Vec::new();
        for row in self.start_row..=self.end_row {
            for col in self.start_column..=self.end_column {
                let src_ref = build_cell_reference(col, row)?;
                if let Some(cell) = worksheet.cell(&src_ref) {
                    let value = cell.value().cloned();
                    let formula = cell.formula().map(String::from);
                    let style_id = cell.style_id();
                    cells_to_copy.push((col, row, value, formula, style_id));
                }
            }
        }

        // Second pass: write to destination
        for (src_col, src_row, value, formula, style_id) in cells_to_copy {
            let dest_col = (src_col as i64 + col_offset) as u32;
            let dest_row = (src_row as i64 + row_offset) as u32;
            let dest_ref = build_cell_reference(dest_col, dest_row)?;

            let dest_cell = worksheet.cell_mut(&dest_ref)?;
            if let Some(v) = value {
                dest_cell.set_value(v);
            }
            if let Some(f) = formula {
                // Adjust formula references when copying
                let adjusted_formula = adjust_formula(&f, col_offset, row_offset)?;
                dest_cell.set_formula(adjusted_formula);
            }
            if let Some(sid) = style_id {
                dest_cell.set_style_id(sid);
            }
        }

        Ok(())
    }

    /// Moves this range to a destination starting cell.
    ///
    /// Equivalent to copy_to() followed by clearing the source range.
    ///
    /// # Arguments
    /// * `worksheet` - The worksheet to move cells within
    /// * `dest_start` - The top-left cell reference where the range should be moved to
    ///
    /// # Returns
    /// `Ok(())` if the move succeeded, or an error if references are invalid.
    pub fn move_to(&self, worksheet: &mut Worksheet, dest_start: &str) -> Result<()> {
        self.copy_to(worksheet, dest_start)?;

        // Clear source range - only clear cells that actually exist
        for row in self.start_row..=self.end_row {
            for col in self.start_column..=self.end_column {
                let src_ref = build_cell_reference(col, row)?;
                // Check if cell exists before clearing to avoid creating empty cells
                if worksheet.cell(&src_ref).is_some() {
                    if let Ok(cell) = worksheet.cell_mut(&src_ref) {
                        cell.clear();
                    }
                }
            }
        }

        Ok(())
    }
}

impl FromStr for CellRange {
    type Err = XlsxError;

    fn from_str(value: &str) -> std::result::Result<Self, Self::Err> {
        Self::parse(value)
    }
}

#[derive(Debug, Clone)]
pub struct CellRangeIter {
    start_column: u32,
    end_column: u32,
    end_row: u32,
    next_column: u32,
    next_row: u32,
    exhausted: bool,
}

impl CellRangeIter {
    fn new(start_column: u32, start_row: u32, end_column: u32, end_row: u32) -> Self {
        Self {
            start_column,
            end_column,
            end_row,
            next_column: start_column,
            next_row: start_row,
            exhausted: false,
        }
    }
}

impl Iterator for CellRangeIter {
    type Item = String;

    fn next(&mut self) -> Option<Self::Item> {
        if self.exhausted {
            return None;
        }

        let reference = build_cell_reference(self.next_column, self.next_row).ok()?;

        if self.next_column == self.end_column {
            if self.next_row == self.end_row {
                self.exhausted = true;
            } else {
                self.next_column = self.start_column;
                self.next_row = self.next_row.checked_add(1)?;
            }
        } else {
            self.next_column = self.next_column.checked_add(1)?;
        }

        Some(reference)
    }
}

pub(crate) fn parse_cell_reference(reference: &str) -> Result<(u32, u32)> {
    let normalized = normalize_cell_reference(reference)?;
    let split_index = normalized
        .char_indices()
        .find_map(|(index, ch)| {
            if ch.is_ascii_digit() {
                Some(index)
            } else {
                None
            }
        })
        .ok_or_else(|| XlsxError::InvalidCellReference(reference.trim().to_string()))?;
    let (column_name, row_text) = normalized.split_at(split_index);

    let column_index = column_name.chars().try_fold(0_u32, |acc, ch| {
        let value = u32::from((ch as u8) - b'A' + 1);
        acc.checked_mul(26).and_then(|v| v.checked_add(value))
    });
    let column_index =
        column_index.ok_or_else(|| XlsxError::InvalidCellReference(reference.to_string()))?;

    let row_index = row_text
        .parse::<u32>()
        .map_err(|_| XlsxError::InvalidCellReference(reference.to_string()))?;

    Ok((column_index, row_index))
}

pub(crate) fn build_cell_reference(column_index: u32, row_index: u32) -> Result<String> {
    let column_name = column_index_to_name(column_index)?;
    if row_index == 0 {
        return Err(XlsxError::InvalidCellReference(format!(
            "{column_name}{row_index}"
        )));
    }
    Ok(format!("{column_name}{row_index}"))
}

fn column_index_to_name(mut column_index: u32) -> Result<String> {
    if column_index == 0 {
        return Err(XlsxError::InvalidCellReference("0".to_string()));
    }

    let mut letters = Vec::new();
    while column_index > 0 {
        let remainder = (column_index - 1) % 26;
        letters.push((b'A' + remainder as u8) as char);
        column_index = (column_index - 1) / 26;
    }
    letters.reverse();
    Ok(letters.into_iter().collect())
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::CellComment;

    #[test]
    fn parse_normalizes_bounds() {
        let range = CellRange::parse(" b3 : a1 ").expect("range should parse");
        assert_eq!(range.start(), "A1");
        assert_eq!(range.end(), "B3");
        assert_eq!(range.width(), 2);
        assert_eq!(range.height(), 3);
    }

    #[test]
    fn parse_single_cell_is_supported() {
        let range = CellRange::parse(" c7 ").expect("range should parse");
        assert_eq!(range.start(), "C7");
        assert_eq!(range.end(), "C7");
        assert_eq!(range.width(), 1);
        assert_eq!(range.height(), 1);
    }

    #[test]
    fn parse_rejects_invalid_input() {
        assert!(CellRange::parse("").is_err());
        assert!(CellRange::parse("A1:B2:C3").is_err());
        assert!(CellRange::parse("A0:B2").is_err());
        assert!(CellRange::parse("A1:").is_err());
    }

    #[test]
    fn iteration_is_row_major() {
        let range = CellRange::parse("A1:B3").expect("range should parse");
        let cells: Vec<String> = range.iter().collect();
        assert_eq!(cells, vec!["A1", "B1", "A2", "B2", "A3", "B3"]);
    }

    #[test]
    fn get_and_set_helpers_work() {
        let mut worksheet = Worksheet::new("Data");
        let range = CellRange::parse("A1:B2").expect("range should parse");

        assert!(range
            .set_value(&mut worksheet, "B2", 42)
            .expect("set should succeed"));
        assert!(!range
            .set_value(&mut worksheet, "C2", 9)
            .expect("set should succeed"));
        assert!(range.set_value(&mut worksheet, "bad", 9).is_err());

        assert_eq!(
            range.get(&worksheet, "B2").and_then(Cell::value),
            Some(&CellValue::Number(42.0))
        );
        assert!(range.get(&worksheet, "C2").is_none());
        assert!(range.get(&worksheet, "bad").is_none());
    }

    #[test]
    fn copy_to_copies_values_and_styles() {
        let mut worksheet = Worksheet::new("Data");

        // Set up source range A1:B2
        worksheet.cell_mut("A1").expect("cell").set_value("Header1");
        worksheet.cell_mut("A1").expect("cell").set_style_id(1);
        worksheet.cell_mut("B1").expect("cell").set_value("Header2");
        worksheet.cell_mut("A2").expect("cell").set_value(100);
        worksheet.cell_mut("A2").expect("cell").set_style_id(2);
        worksheet.cell_mut("B2").expect("cell").set_value(200);

        let range = CellRange::parse("A1:B2").expect("range should parse");
        range
            .copy_to(&mut worksheet, "D5")
            .expect("copy should succeed");

        // Verify source is unchanged
        assert_eq!(
            worksheet.cell("A1").and_then(Cell::value),
            Some(&CellValue::String("Header1".to_string()))
        );
        assert_eq!(worksheet.cell("A1").and_then(Cell::style_id), Some(1));

        // Verify destination has copied values
        assert_eq!(
            worksheet.cell("D5").and_then(Cell::value),
            Some(&CellValue::String("Header1".to_string()))
        );
        assert_eq!(worksheet.cell("D5").and_then(Cell::style_id), Some(1));
        assert_eq!(
            worksheet.cell("E5").and_then(Cell::value),
            Some(&CellValue::String("Header2".to_string()))
        );
        assert_eq!(
            worksheet.cell("D6").and_then(Cell::value),
            Some(&CellValue::Number(100.0))
        );
        assert_eq!(worksheet.cell("D6").and_then(Cell::style_id), Some(2));
        assert_eq!(
            worksheet.cell("E6").and_then(Cell::value),
            Some(&CellValue::Number(200.0))
        );
    }

    #[test]
    fn move_to_clears_source_range() {
        let mut worksheet = Worksheet::new("Data");

        // Set up source range A1:B2
        worksheet.cell_mut("A1").expect("cell").set_value("Test");
        worksheet.cell_mut("B1").expect("cell").set_value(42);
        worksheet.cell_mut("A2").expect("cell").set_value("Data");
        worksheet.cell_mut("B2").expect("cell").set_value(99);

        let range = CellRange::parse("A1:B2").expect("range should parse");
        range
            .move_to(&mut worksheet, "C3")
            .expect("move should succeed");

        // Verify source is cleared
        assert_eq!(worksheet.cell("A1").and_then(Cell::value), None);
        assert_eq!(worksheet.cell("B1").and_then(Cell::value), None);
        assert_eq!(worksheet.cell("A2").and_then(Cell::value), None);
        assert_eq!(worksheet.cell("B2").and_then(Cell::value), None);

        // Verify destination has moved values
        assert_eq!(
            worksheet.cell("C3").and_then(Cell::value),
            Some(&CellValue::String("Test".to_string()))
        );
        assert_eq!(
            worksheet.cell("D3").and_then(Cell::value),
            Some(&CellValue::Number(42.0))
        );
        assert_eq!(
            worksheet.cell("C4").and_then(Cell::value),
            Some(&CellValue::String("Data".to_string()))
        );
        assert_eq!(
            worksheet.cell("D4").and_then(Cell::value),
            Some(&CellValue::Number(99.0))
        );
    }

    #[test]
    fn copy_to_handles_formulas() {
        let mut worksheet = Worksheet::new("Data");

        worksheet.cell_mut("A1").expect("cell").set_value(10);
        worksheet.cell_mut("B1").expect("cell").set_formula("=A1*2");

        let range = CellRange::parse("A1:B1").expect("range should parse");
        range
            .copy_to(&mut worksheet, "A3")
            .expect("copy should succeed");

        assert_eq!(
            worksheet.cell("A3").and_then(Cell::value),
            Some(&CellValue::Number(10.0))
        );
        assert_eq!(
            worksheet.cell("B3").and_then(Cell::formula),
            Some("A3*2") // Formula adjusted correctly
        );
    }

    #[test]
    fn copy_empty_range_succeeds() {
        let mut worksheet = Worksheet::new("Data");

        // Create a range with no cells set
        let range = CellRange::parse("A1:B2").expect("range should parse");
        range
            .copy_to(&mut worksheet, "D5")
            .expect("copy should succeed");

        // Destination should also be empty
        assert!(worksheet.cell("D5").is_none());
        assert!(worksheet.cell("E5").is_none());
        assert!(worksheet.cell("D6").is_none());
        assert!(worksheet.cell("E6").is_none());
    }

    #[test]
    fn move_empty_range_succeeds() {
        let mut worksheet = Worksheet::new("Data");

        // Create a range with no cells set
        let range = CellRange::parse("A1:B2").expect("range should parse");
        range
            .move_to(&mut worksheet, "D5")
            .expect("move should succeed");

        // Both source and destination should be empty
        assert!(worksheet.cell("A1").is_none());
        assert!(worksheet.cell("B1").is_none());
        assert!(worksheet.cell("D5").is_none());
        assert!(worksheet.cell("E5").is_none());
    }

    #[test]
    fn copy_single_cell_range() {
        let mut worksheet = Worksheet::new("Data");

        worksheet.cell_mut("A1").expect("cell").set_value("Single");
        worksheet.cell_mut("A1").expect("cell").set_style_id(5);

        let range = CellRange::parse("A1").expect("range should parse");
        range
            .copy_to(&mut worksheet, "C3")
            .expect("copy should succeed");

        // Source should remain
        assert_eq!(
            worksheet.cell("A1").and_then(Cell::value),
            Some(&CellValue::String("Single".to_string()))
        );
        assert_eq!(worksheet.cell("A1").and_then(Cell::style_id), Some(5));

        // Destination should have the copy
        assert_eq!(
            worksheet.cell("C3").and_then(Cell::value),
            Some(&CellValue::String("Single".to_string()))
        );
        assert_eq!(worksheet.cell("C3").and_then(Cell::style_id), Some(5));
    }

    #[test]
    fn move_single_cell_range() {
        let mut worksheet = Worksheet::new("Data");

        worksheet.cell_mut("A1").expect("cell").set_value("Move me");
        worksheet.cell_mut("A1").expect("cell").set_style_id(3);

        let range = CellRange::parse("A1").expect("range should parse");
        range
            .move_to(&mut worksheet, "B2")
            .expect("move should succeed");

        // Source should be cleared
        assert!(worksheet.cell("A1").and_then(Cell::value).is_none());
        assert!(worksheet.cell("A1").and_then(Cell::style_id).is_none());

        // Destination should have the value
        assert_eq!(
            worksheet.cell("B2").and_then(Cell::value),
            Some(&CellValue::String("Move me".to_string()))
        );
        assert_eq!(worksheet.cell("B2").and_then(Cell::style_id), Some(3));
    }

    #[test]
    fn copy_with_sparse_cells() {
        let mut worksheet = Worksheet::new("Data");

        // Only set some cells in the range
        worksheet.cell_mut("A1").expect("cell").set_value("A1");
        worksheet.cell_mut("B2").expect("cell").set_value("B2");
        // A2 and B1 remain empty

        let range = CellRange::parse("A1:B2").expect("range should parse");
        range
            .copy_to(&mut worksheet, "D5")
            .expect("copy should succeed");

        // Check copied cells
        assert_eq!(
            worksheet.cell("D5").and_then(Cell::value),
            Some(&CellValue::String("A1".to_string()))
        );
        assert_eq!(
            worksheet.cell("E6").and_then(Cell::value),
            Some(&CellValue::String("B2".to_string()))
        );

        // Check that empty cells remain empty
        assert!(worksheet.cell("D6").is_none());
        assert!(worksheet.cell("E5").is_none());
    }

    #[test]
    fn overlapping_copy_copies_correctly() {
        let mut worksheet = Worksheet::new("Data");

        // Set up source A1:A3
        worksheet.cell_mut("A1").expect("cell").set_value(1);
        worksheet.cell_mut("A2").expect("cell").set_value(2);
        worksheet.cell_mut("A3").expect("cell").set_value(3);

        let range = CellRange::parse("A1:A3").expect("range should parse");
        // Copy to A2:A4 (overlapping)
        range
            .copy_to(&mut worksheet, "A2")
            .expect("copy should succeed");

        // A1 should remain unchanged
        assert_eq!(
            worksheet.cell("A1").and_then(Cell::value),
            Some(&CellValue::Number(1.0))
        );
        // A2, A3, A4 should have copied values
        assert_eq!(
            worksheet.cell("A2").and_then(Cell::value),
            Some(&CellValue::Number(1.0))
        );
        assert_eq!(
            worksheet.cell("A3").and_then(Cell::value),
            Some(&CellValue::Number(2.0))
        );
        assert_eq!(
            worksheet.cell("A4").and_then(Cell::value),
            Some(&CellValue::Number(3.0))
        );
    }

    #[test]
    fn cell_clear_removes_all_data() {
        let mut cell = Cell::new();

        cell.set_value("test")
            .set_formula("SUM(A1:A5)")
            .set_cached_value(CellValue::Number(42.0))
            .set_style_id(5)
            .set_comment(CellComment::new("Author", "Note"))
            .set_array_formula(true)
            .set_shared_formula_index(1);

        cell.clear();

        assert!(cell.value().is_none());
        assert!(cell.formula().is_none());
        assert!(cell.cached_value().is_none());
        assert!(cell.style_id().is_none());
        assert!(cell.comment().is_none());
        assert!(!cell.is_array_formula());
        assert!(cell.array_range().is_none());
        assert!(cell.shared_formula_index().is_none());
    }

    #[test]
    fn adjust_formula_basic() {
        // Test adjusting a simple formula
        let result = adjust_formula("A1*2", 0, 2).expect("should adjust");
        assert_eq!(result, "A3*2");

        // Test with leading =
        let result = adjust_formula("=A1*2", 0, 2).expect("should adjust");
        assert_eq!(result, "A3*2");

        // Test with column offset
        let result = adjust_formula("A1+B1", 2, 0).expect("should adjust");
        assert_eq!(result, "C1+D1");

        // Test with both offsets
        let result = adjust_formula("A1+B2", 1, 1).expect("should adjust");
        assert_eq!(result, "B2+C3");
    }

    #[test]
    fn adjust_formula_absolute_refs() {
        // Absolute column
        let result = adjust_formula("$A1*2", 2, 3).expect("should adjust");
        assert_eq!(result, "$A4*2");

        // Absolute row
        let result = adjust_formula("A$1*2", 2, 3).expect("should adjust");
        assert_eq!(result, "C$1*2");

        // Both absolute
        let result = adjust_formula("$A$1*2", 2, 3).expect("should adjust");
        assert_eq!(result, "$A$1*2");
    }

    #[test]
    fn copy_complex_formulas() {
        let mut worksheet = Worksheet::new("Data");

        // Set up source data
        worksheet.cell_mut("A1").unwrap().set_value(10);
        worksheet.cell_mut("A2").unwrap().set_value(20);
        worksheet.cell_mut("A3").unwrap().set_value(30);
        worksheet.cell_mut("B1").unwrap().set_formula("=SUM(A1:A3)");
        worksheet.cell_mut("B2").unwrap().set_formula("=A1+A2");

        let range = CellRange::parse("A1:B2").expect("range should parse");
        range
            .copy_to(&mut worksheet, "D5")
            .expect("copy should succeed");

        // Check copied formulas are adjusted
        assert_eq!(
            worksheet.cell("E5").and_then(Cell::formula),
            Some("SUM(D5:D7)") // A1:A3 -> D5:D7
        );
        assert_eq!(
            worksheet.cell("E6").and_then(Cell::formula),
            Some("D5+D6") // A1+A2 -> D5+D6
        );
    }

    #[test]
    fn move_with_formulas() {
        let mut worksheet = Worksheet::new("Data");

        worksheet.cell_mut("A1").unwrap().set_value(5);
        worksheet.cell_mut("A2").unwrap().set_formula("=A1*3");

        let range = CellRange::parse("A1:A2").expect("range should parse");
        range
            .move_to(&mut worksheet, "C5")
            .expect("move should succeed");

        // Source should be cleared
        assert!(worksheet.cell("A1").and_then(Cell::value).is_none());
        assert!(worksheet.cell("A2").and_then(Cell::formula).is_none());

        // Destination should have moved data
        assert_eq!(
            worksheet.cell("C5").and_then(Cell::value),
            Some(&CellValue::Number(5.0))
        );
        assert_eq!(
            worksheet.cell("C6").and_then(Cell::formula),
            Some("C5*3") // A1*3 -> C5*3
        );
    }

    #[test]
    fn copy_with_absolute_references() {
        let mut worksheet = Worksheet::new("Data");

        worksheet.cell_mut("A1").unwrap().set_value(100);
        worksheet.cell_mut("B1").unwrap().set_formula("=$A$1*2"); // Absolute ref
        worksheet.cell_mut("B2").unwrap().set_formula("=A$1*2"); // Row absolute
        worksheet.cell_mut("B3").unwrap().set_formula("=$A1*2"); // Col absolute

        let range = CellRange::parse("B1:B3").expect("range should parse");
        range
            .copy_to(&mut worksheet, "D5")
            .expect("copy should succeed");

        // Absolute references should not change
        assert_eq!(worksheet.cell("D5").and_then(Cell::formula), Some("$A$1*2"));
        // Row absolute should not change row
        assert_eq!(
            worksheet.cell("D6").and_then(Cell::formula),
            Some("C$1*2") // Column adjusted (A->C), row stays $1
        );
        // Column absolute should not change column
        assert_eq!(
            worksheet.cell("D7").and_then(Cell::formula),
            Some("$A5*2") // Column stays $A, row adjusted (1->5)
        );
    }
}
