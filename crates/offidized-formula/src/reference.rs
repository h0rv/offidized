use std::fmt;

/// A reference to a single cell.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct CellRef {
    /// Optional sheet name (e.g. `Sheet1`).
    pub sheet: Option<String>,
    /// 1-based column index.
    pub col: u32,
    /// 1-based row index.
    pub row: u32,
    /// Whether the column is absolute (`$A`).
    pub col_absolute: bool,
    /// Whether the row is absolute (`$1`).
    pub row_absolute: bool,
}

impl fmt::Display for CellRef {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        if let Some(ref sheet) = self.sheet {
            // Quote sheet names that contain spaces
            if sheet.contains(' ') {
                write!(f, "'{sheet}'!")?;
            } else {
                write!(f, "{sheet}!")?;
            }
        }
        if self.col_absolute {
            write!(f, "$")?;
        }
        if let Some(letters) = column_index_to_letters(self.col) {
            write!(f, "{letters}")?;
        }
        if self.row_absolute {
            write!(f, "$")?;
        }
        write!(f, "{}", self.row)
    }
}

/// A reference to a rectangular range of cells.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct RangeRef {
    /// Optional sheet name.
    pub sheet: Option<String>,
    /// 1-based start column.
    pub start_col: u32,
    /// 1-based start row.
    pub start_row: u32,
    /// 1-based end column.
    pub end_col: u32,
    /// 1-based end row.
    pub end_row: u32,
}

impl fmt::Display for RangeRef {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        if let Some(ref sheet) = self.sheet {
            if sheet.contains(' ') {
                write!(f, "'{sheet}'!")?;
            } else {
                write!(f, "{sheet}!")?;
            }
        }
        let start_col = column_index_to_letters(self.start_col).unwrap_or_default();
        let end_col = column_index_to_letters(self.end_col).unwrap_or_default();
        write!(
            f,
            "{}{}:{}{}",
            start_col, self.start_row, end_col, self.end_row
        )
    }
}

/// Converts a 1-based column index to Excel column letters.
///
/// e.g. 1 → "A", 26 → "Z", 27 → "AA"
pub fn column_index_to_letters(mut index: u32) -> Option<String> {
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

/// Converts Excel column letters to a 1-based column index.
///
/// e.g. "A" → 1, "Z" → 26, "AA" → 27
pub fn column_letters_to_index(letters: &str) -> Option<u32> {
    if letters.is_empty() {
        return None;
    }
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

#[cfg(test)]
mod tests {
    use super::*;
    use pretty_assertions::assert_eq;

    #[test]
    fn column_letters_to_index_basic() {
        assert_eq!(column_letters_to_index("A"), Some(1));
        assert_eq!(column_letters_to_index("Z"), Some(26));
        assert_eq!(column_letters_to_index("AA"), Some(27));
        assert_eq!(column_letters_to_index("AZ"), Some(52));
        assert_eq!(column_letters_to_index("BA"), Some(53));
        assert_eq!(column_letters_to_index("XFD"), Some(16384));
    }

    #[test]
    fn column_index_to_letters_basic() {
        assert_eq!(column_index_to_letters(1), Some("A".to_string()));
        assert_eq!(column_index_to_letters(26), Some("Z".to_string()));
        assert_eq!(column_index_to_letters(27), Some("AA".to_string()));
        assert_eq!(column_index_to_letters(52), Some("AZ".to_string()));
        assert_eq!(column_index_to_letters(16384), Some("XFD".to_string()));
    }

    #[test]
    fn column_conversion_roundtrip() {
        for i in 1..=16384 {
            let letters = column_index_to_letters(i).unwrap();
            assert_eq!(column_letters_to_index(&letters), Some(i));
        }
    }

    #[test]
    fn column_edge_cases() {
        assert_eq!(column_index_to_letters(0), None);
        assert_eq!(column_letters_to_index(""), None);
        assert_eq!(column_letters_to_index("1"), None);
    }

    #[test]
    fn cell_ref_display() {
        let r = CellRef {
            sheet: None,
            col: 1,
            row: 1,
            col_absolute: false,
            row_absolute: false,
        };
        assert_eq!(r.to_string(), "A1");

        let r = CellRef {
            sheet: Some("Sheet1".to_string()),
            col: 3,
            row: 5,
            col_absolute: true,
            row_absolute: true,
        };
        assert_eq!(r.to_string(), "Sheet1!$C$5");

        let r = CellRef {
            sheet: Some("My Sheet".to_string()),
            col: 2,
            row: 10,
            col_absolute: false,
            row_absolute: false,
        };
        assert_eq!(r.to_string(), "'My Sheet'!B10");
    }

    #[test]
    fn range_ref_display() {
        let r = RangeRef {
            sheet: None,
            start_col: 1,
            start_row: 1,
            end_col: 3,
            end_row: 5,
        };
        assert_eq!(r.to_string(), "A1:C5");

        let r = RangeRef {
            sheet: Some("Data".to_string()),
            start_col: 2,
            start_row: 3,
            end_col: 10,
            end_row: 20,
        };
        assert_eq!(r.to_string(), "Data!B3:J20");
    }
}
