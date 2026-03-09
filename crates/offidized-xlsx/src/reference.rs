//! R1C1 and A1 reference style conversion utilities.
//!
//! Excel supports two reference styles:
//! - **A1 style**: column letters + row number (e.g. `B3`, `$A$1`, `A$1`, `$A1`)
//! - **R1C1 style**: `R` + row + `C` + column, with optional `[offset]` for relative references
//!
//! This module provides conversion functions between the two styles.

/// Converts an A1-style cell reference to R1C1-style relative to a base cell.
///
/// Supports absolute references (`$A$1` -> `R1C1`), relative references
/// (`A1` from base (3,2) -> `R[-2]C[-1]`), and mixed references
/// (`$A1` -> `R[-2]C1`, `A$1` -> `R1C[-1]`).
///
/// - `reference`: The A1-style reference (e.g. "A1", "$B$3", "A$1", "$A1")
/// - `base_row`: The 1-based row of the cell containing the formula
/// - `base_col`: The 1-based column of the cell containing the formula
///
/// Returns the R1C1-style string.
pub fn a1_to_r1c1(reference: &str, base_row: u32, base_col: u32) -> Option<String> {
    let trimmed = reference.trim();
    if trimmed.is_empty() {
        return None;
    }

    let bytes = trimmed.as_bytes();
    let mut pos = 0;

    // Check if column is absolute
    let col_absolute = bytes.get(pos) == Some(&b'$');
    if col_absolute {
        pos += 1;
    }

    // Parse column letters
    let col_start = pos;
    while pos < bytes.len() && bytes[pos].is_ascii_alphabetic() {
        pos += 1;
    }
    let col_end = pos;
    if col_start == col_end {
        return None;
    }

    // Check if row is absolute
    let row_absolute = bytes.get(pos) == Some(&b'$');
    if row_absolute {
        pos += 1;
    }

    // Parse row digits
    let row_start = pos;
    while pos < bytes.len() && bytes[pos].is_ascii_digit() {
        pos += 1;
    }
    if row_start == pos || pos != bytes.len() {
        return None;
    }

    let col_letters = &trimmed[col_start..col_end];
    let col_index = column_letters_to_index(col_letters)?;
    let row_index: u32 = trimmed[row_start..].parse().ok()?;

    if row_index == 0 || col_index == 0 {
        return None;
    }

    let row_part = if row_absolute {
        format!("R{row_index}")
    } else {
        let offset = row_index as i64 - base_row as i64;
        if offset == 0 {
            "R".to_string()
        } else {
            format!("R[{offset}]")
        }
    };

    let col_part = if col_absolute {
        format!("C{col_index}")
    } else {
        let offset = col_index as i64 - base_col as i64;
        if offset == 0 {
            "C".to_string()
        } else {
            format!("C[{offset}]")
        }
    };

    Some(format!("{row_part}{col_part}"))
}

/// Converts an R1C1-style cell reference to A1-style relative to a base cell.
///
/// - `reference`: The R1C1-style reference (e.g. "R1C1", "R[-1]C[2]", "RC[-1]", "R1C")
/// - `base_row`: The 1-based row of the cell containing the formula
/// - `base_col`: The 1-based column of the cell containing the formula
///
/// Returns the A1-style string with `$` for absolute references.
pub fn r1c1_to_a1(reference: &str, base_row: u32, base_col: u32) -> Option<String> {
    let trimmed = reference.trim();
    if trimmed.is_empty() {
        return None;
    }

    let bytes = trimmed.as_bytes();
    let mut pos = 0;

    // Expect 'R' or 'r'
    if !matches!(bytes.get(pos), Some(b'R' | b'r')) {
        return None;
    }
    pos += 1;

    // Parse row part
    let (row_value, row_absolute) = parse_r1c1_part(bytes, &mut pos, base_row)?;

    // Expect 'C' or 'c'
    if !matches!(bytes.get(pos), Some(b'C' | b'c')) {
        return None;
    }
    pos += 1;

    // Parse column part
    let (col_value, col_absolute) = parse_r1c1_part(bytes, &mut pos, base_col)?;

    if pos != bytes.len() {
        return None;
    }

    if row_value == 0 || col_value == 0 {
        return None;
    }

    let col_letters = column_index_to_letters(col_value)?;

    let col_str = if col_absolute {
        format!("${col_letters}")
    } else {
        col_letters
    };

    let row_str = if row_absolute {
        format!("${row_value}")
    } else {
        row_value.to_string()
    };

    Some(format!("{col_str}{row_str}"))
}

/// Parses a row or column part from an R1C1 reference.
///
/// Returns (absolute_value, is_absolute).
fn parse_r1c1_part(bytes: &[u8], pos: &mut usize, base: u32) -> Option<(u32, bool)> {
    if *pos >= bytes.len() || bytes[*pos] == b'C' || bytes[*pos] == b'c' {
        // No number or bracket -> relative offset of 0
        return Some((base, false));
    }

    if bytes[*pos] == b'[' {
        // Relative reference: [offset]
        *pos += 1;
        let start = *pos;
        // Allow optional negative sign
        if *pos < bytes.len() && bytes[*pos] == b'-' {
            *pos += 1;
        }
        while *pos < bytes.len() && bytes[*pos].is_ascii_digit() {
            *pos += 1;
        }
        if *pos >= bytes.len() || bytes[*pos] != b']' {
            return None;
        }
        let offset_str = std::str::from_utf8(&bytes[start..*pos]).ok()?;
        let offset: i64 = offset_str.parse().ok()?;
        *pos += 1; // skip ']'

        let result = (base as i64).checked_add(offset)?;
        if result < 1 {
            return None;
        }
        Some((result as u32, false))
    } else if bytes[*pos].is_ascii_digit() {
        // Absolute reference: just digits
        let start = *pos;
        while *pos < bytes.len() && bytes[*pos].is_ascii_digit() {
            *pos += 1;
        }
        let value: u32 = std::str::from_utf8(&bytes[start..*pos])
            .ok()?
            .parse()
            .ok()?;
        Some((value, true))
    } else {
        None
    }
}

fn column_letters_to_index(letters: &str) -> Option<u32> {
    let mut index = 0_u32;
    for byte in letters.bytes() {
        let ch = byte.to_ascii_uppercase();
        if !ch.is_ascii_uppercase() {
            return None;
        }
        index = index.checked_mul(26)?;
        index = index.checked_add(u32::from(ch - b'A' + 1))?;
    }
    if index == 0 {
        None
    } else {
        Some(index)
    }
}

fn column_index_to_letters(mut index: u32) -> Option<String> {
    if index == 0 {
        return None;
    }
    let mut letters = Vec::new();
    while index > 0 {
        let remainder = (index - 1) % 26;
        letters.push((b'A' + remainder as u8) as char);
        index = (index - 1) / 26;
    }
    letters.reverse();
    Some(letters.into_iter().collect())
}

#[cfg(test)]
mod tests {
    use super::*;

    // ===== Feature 10: R1C1 reference style =====

    #[test]
    fn a1_to_r1c1_absolute_references() {
        assert_eq!(a1_to_r1c1("$A$1", 1, 1), Some("R1C1".to_string()));
        assert_eq!(a1_to_r1c1("$B$3", 1, 1), Some("R3C2".to_string()));
        assert_eq!(a1_to_r1c1("$Z$26", 1, 1), Some("R26C26".to_string()));
        assert_eq!(a1_to_r1c1("$AA$1", 1, 1), Some("R1C27".to_string()));
    }

    #[test]
    fn a1_to_r1c1_relative_references() {
        // From base cell (3, 2) = B3
        assert_eq!(a1_to_r1c1("A1", 3, 2), Some("R[-2]C[-1]".to_string()));
        assert_eq!(a1_to_r1c1("B3", 3, 2), Some("RC".to_string()));
        assert_eq!(a1_to_r1c1("C5", 3, 2), Some("R[2]C[1]".to_string()));
        assert_eq!(a1_to_r1c1("A3", 3, 2), Some("RC[-1]".to_string()));
        assert_eq!(a1_to_r1c1("B1", 3, 2), Some("R[-2]C".to_string()));
    }

    #[test]
    fn a1_to_r1c1_mixed_references() {
        // $A1 from base (3, 2) -> absolute col, relative row
        assert_eq!(a1_to_r1c1("$A1", 3, 2), Some("R[-2]C1".to_string()));
        // A$1 from base (3, 2) -> relative col, absolute row
        assert_eq!(a1_to_r1c1("A$1", 3, 2), Some("R1C[-1]".to_string()));
    }

    #[test]
    fn a1_to_r1c1_invalid_inputs() {
        assert!(a1_to_r1c1("", 1, 1).is_none());
        assert!(a1_to_r1c1("123", 1, 1).is_none());
        assert!(a1_to_r1c1("$", 1, 1).is_none());
    }

    #[test]
    fn r1c1_to_a1_absolute_references() {
        assert_eq!(r1c1_to_a1("R1C1", 1, 1), Some("$A$1".to_string()));
        assert_eq!(r1c1_to_a1("R3C2", 1, 1), Some("$B$3".to_string()));
        assert_eq!(r1c1_to_a1("R26C26", 5, 5), Some("$Z$26".to_string()));
        assert_eq!(r1c1_to_a1("R1C27", 1, 1), Some("$AA$1".to_string()));
    }

    #[test]
    fn r1c1_to_a1_relative_references() {
        // From base cell (3, 2) = B3
        assert_eq!(r1c1_to_a1("R[-2]C[-1]", 3, 2), Some("A1".to_string()));
        assert_eq!(r1c1_to_a1("RC", 3, 2), Some("B3".to_string()));
        assert_eq!(r1c1_to_a1("R[2]C[1]", 3, 2), Some("C5".to_string()));
        assert_eq!(r1c1_to_a1("RC[-1]", 3, 2), Some("A3".to_string()));
        assert_eq!(r1c1_to_a1("R[-2]C", 3, 2), Some("B1".to_string()));
    }

    #[test]
    fn r1c1_to_a1_mixed_references() {
        // R[-2]C1 from base (3, 2) -> absolute col, relative row
        assert_eq!(r1c1_to_a1("R[-2]C1", 3, 2), Some("$A1".to_string()));
        // R1C[-1] from base (3, 2) -> relative col, absolute row
        assert_eq!(r1c1_to_a1("R1C[-1]", 3, 2), Some("A$1".to_string()));
    }

    #[test]
    fn r1c1_to_a1_invalid_inputs() {
        assert!(r1c1_to_a1("", 1, 1).is_none());
        assert!(r1c1_to_a1("A1", 1, 1).is_none());
        assert!(r1c1_to_a1("R", 1, 1).is_none()); // missing C part
        assert!(r1c1_to_a1("RC[-999]", 1, 1).is_none()); // would result in col 0 or negative
    }

    #[test]
    fn a1_r1c1_roundtrip() {
        // Absolute roundtrip
        let r1c1 = a1_to_r1c1("$C$5", 1, 1).unwrap();
        assert_eq!(r1c1, "R5C3");
        let a1 = r1c1_to_a1(&r1c1, 1, 1).unwrap();
        assert_eq!(a1, "$C$5");

        // Relative roundtrip from base (3, 2)
        let r1c1 = a1_to_r1c1("D7", 3, 2).unwrap();
        let a1 = r1c1_to_a1(&r1c1, 3, 2).unwrap();
        assert_eq!(a1, "D7");
    }

    #[test]
    fn column_letters_to_index_works() {
        assert_eq!(column_letters_to_index("A"), Some(1));
        assert_eq!(column_letters_to_index("Z"), Some(26));
        assert_eq!(column_letters_to_index("AA"), Some(27));
        assert_eq!(column_letters_to_index("AZ"), Some(52));
        assert_eq!(column_letters_to_index("BA"), Some(53));
    }

    #[test]
    fn column_index_to_letters_works() {
        assert_eq!(column_index_to_letters(1), Some("A".to_string()));
        assert_eq!(column_index_to_letters(26), Some("Z".to_string()));
        assert_eq!(column_index_to_letters(27), Some("AA".to_string()));
        assert_eq!(column_index_to_letters(52), Some("AZ".to_string()));
        assert_eq!(column_index_to_letters(0), None);
    }
}
