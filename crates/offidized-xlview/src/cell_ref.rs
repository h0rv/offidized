//! Cell reference parsing utilities.
//!
//! Parses Excel-style cell references like "A1", "B2:D5", and sqref
//! ranges like "A1:B2 C3:D4".

/// Parse a single cell reference like "A1" into (col, row) both 0-indexed.
///
/// Returns `None` if the reference is invalid.
pub fn parse_cell_ref(s: &str) -> Option<(u32, u32)> {
    let s = s.trim().replace('$', "");
    // Split at the boundary between letters and digits
    let col_end = s.bytes().position(|b| b.is_ascii_digit())?;
    if col_end == 0 {
        return None;
    }
    let col_str = s.get(..col_end)?;
    let row_str = s.get(col_end..)?;

    let col = col_letters_to_index(col_str)?;
    let row: u32 = row_str.parse().ok()?;
    if row == 0 {
        return None;
    }
    Some((col, row - 1)) // Convert to 0-indexed row
}

/// Parse a cell range like "A1:B2" into (min_row, min_col, max_row, max_col) all 0-indexed.
///
/// Also handles single cell references like "A1" (treated as a 1x1 range).
pub fn parse_cell_range(s: &str) -> Option<(u32, u32, u32, u32)> {
    let s = s.trim();
    // Strip optional sheet name prefix (e.g. "Sheet1!A1:B2")
    let range_part = if let Some(idx) = s.find('!') {
        s.get(idx + 1..)?
    } else {
        s
    };

    if let Some((left, right)) = range_part.split_once(':') {
        let (col1, row1) = parse_cell_ref(left)?;
        let (col2, row2) = parse_cell_ref(right)?;
        Some((
            row1.min(row2),
            col1.min(col2),
            row1.max(row2),
            col2.max(col2),
        ))
    } else {
        // Single cell
        let (col, row) = parse_cell_ref(range_part)?;
        Some((row, col, row, col))
    }
}

/// Parse an sqref string (space-separated list of cell ranges) into a list of
/// (min_row, min_col, max_row, max_col) tuples, all 0-indexed.
pub fn parse_sqref(sqref: &str) -> Vec<(u32, u32, u32, u32)> {
    sqref
        .split_whitespace()
        .filter_map(parse_cell_range)
        .collect()
}

/// Convert column letters (e.g. "A", "AB", "ZZ") to a 0-indexed column number.
fn col_letters_to_index(letters: &str) -> Option<u32> {
    if letters.is_empty() {
        return None;
    }
    let mut result: u32 = 0;
    for byte in letters.bytes() {
        if !byte.is_ascii_alphabetic() {
            return None;
        }
        let digit = u32::from(byte.to_ascii_uppercase() - b'A') + 1;
        result = result.checked_mul(26)?.checked_add(digit)?;
    }
    Some(result - 1) // Convert to 0-indexed
}

#[cfg(test)]
#[allow(
    clippy::unwrap_used,
    clippy::expect_used,
    clippy::indexing_slicing,
    clippy::panic
)]
mod tests {
    use super::*;

    #[test]
    fn test_parse_cell_ref() {
        assert_eq!(parse_cell_ref("A1"), Some((0, 0)));
        assert_eq!(parse_cell_ref("B2"), Some((1, 1)));
        assert_eq!(parse_cell_ref("Z1"), Some((25, 0)));
        assert_eq!(parse_cell_ref("AA1"), Some((26, 0)));
        assert_eq!(parse_cell_ref("$A$1"), Some((0, 0)));
    }

    #[test]
    fn test_parse_cell_range() {
        assert_eq!(parse_cell_range("A1:B2"), Some((0, 0, 1, 1)));
        assert_eq!(parse_cell_range("A1"), Some((0, 0, 0, 0)));
    }

    #[test]
    fn test_parse_sqref() {
        let ranges = parse_sqref("A1:B2 C3:D4");
        assert_eq!(ranges.len(), 2);
        assert_eq!(ranges[0], (0, 0, 1, 1));
        assert_eq!(ranges[1], (2, 2, 3, 3));
    }

    #[test]
    fn test_col_letters_to_index() {
        assert_eq!(col_letters_to_index("A"), Some(0));
        assert_eq!(col_letters_to_index("B"), Some(1));
        assert_eq!(col_letters_to_index("Z"), Some(25));
        assert_eq!(col_letters_to_index("AA"), Some(26));
        assert_eq!(col_letters_to_index("AB"), Some(27));
    }
}
