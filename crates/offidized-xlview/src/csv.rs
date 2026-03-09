//! Minimal CSV/TSV parser that produces a [`Workbook`] with a single sheet.

use crate::error::Result;
use crate::types::style::Theme;
use crate::types::workbook::{Cell, CellData, CellType, CompiledFormat, Sheet, Workbook};

/// Delimiter for parsing.
#[derive(Clone, Copy)]
pub(crate) enum Delimiter {
    /// Comma-separated values.
    Comma,
    /// Tab-separated values.
    Tab,
}

/// Parse CSV/TSV bytes into a [`Workbook`] with one sheet.
pub(crate) fn parse_delimited(data: &[u8], delim: Delimiter) -> Result<Workbook> {
    let text = String::from_utf8_lossy(data);
    let sep = match delim {
        Delimiter::Comma => ',',
        Delimiter::Tab => '\t',
    };

    let mut cells: Vec<CellData> = Vec::new();
    let mut max_row: u32 = 0;
    let mut max_col: u32 = 0;

    for (row_idx, line) in text.lines().enumerate() {
        if line.is_empty() {
            continue;
        }
        #[allow(clippy::cast_possible_truncation)]
        let row = row_idx as u32;
        for (col_idx, field) in split_csv_line(line, sep).iter().enumerate() {
            #[allow(clippy::cast_possible_truncation)]
            let col = col_idx as u32;
            let value = field.trim().to_string();
            if value.is_empty() {
                continue;
            }

            // Try to detect numbers
            let (cell_type, display) = if value.parse::<f64>().is_ok() {
                (CellType::Number, value)
            } else {
                (CellType::String, value)
            };

            cells.push(CellData {
                r: row,
                c: col,
                cell: Cell {
                    v: Some(display),
                    t: cell_type,
                    s: None,
                    style_idx: None,
                    raw: None,
                    cached_display: None,
                    rich_text: None,
                    cached_rich_text: None,
                    has_comment: None,
                    hyperlink: None,
                    formula: None,
                },
            });

            if col > max_col {
                max_col = col;
            }
        }
        if row > max_row {
            max_row = row;
        }
    }

    let name = match delim {
        Delimiter::Comma => "CSV".to_string(),
        Delimiter::Tab => "TSV".to_string(),
    };

    let mut sheet = Sheet {
        name,
        tab_color: None,
        cells,
        cells_by_row: Vec::new(),
        merges: Vec::new(),
        col_widths: Vec::new(),
        row_heights: Vec::new(),
        default_col_width: 8.43,
        default_row_height: 15.0,
        hidden_cols: Vec::new(),
        hidden_rows: Vec::new(),
        max_row,
        max_col,
        frozen_rows: 0,
        frozen_cols: 0,
        drawings: Vec::new(),
        charts: Vec::new(),
        data_validations: Vec::new(),
        conditional_formatting: Vec::new(),
        conditional_formatting_cache: Vec::new(),
        sparkline_groups: Vec::new(),
        auto_filter: None,
        comments: Vec::new(),
        comments_by_cell: Default::default(),
        hyperlinks: Vec::new(),
    };
    sheet.rebuild_cell_index();

    Ok(Workbook {
        sheets: vec![sheet],
        theme: Theme {
            colors: Vec::new(),
            major_font: None,
            minor_font: None,
        },
        date1904: false,
        images: Vec::new(),
        dxf_styles: Vec::new(),
        shared_strings: Vec::new(),
        resolved_styles: Vec::new(),
        default_style: None,
        numfmt_cache: vec![CompiledFormat::General],
        sheet_paths: Vec::new(),
    })
}

/// Split a CSV line respecting quoted fields.
fn split_csv_line(line: &str, sep: char) -> Vec<String> {
    let mut fields = Vec::new();
    let mut current = String::new();
    let mut in_quotes = false;
    let mut chars = line.chars().peekable();

    while let Some(ch) = chars.next() {
        if in_quotes {
            if ch == '"' {
                if chars.peek() == Some(&'"') {
                    // Escaped quote
                    current.push('"');
                    chars.next();
                } else {
                    in_quotes = false;
                }
            } else {
                current.push(ch);
            }
        } else if ch == '"' {
            in_quotes = true;
        } else if ch == sep {
            fields.push(current.clone());
            current.clear();
        } else {
            current.push(ch);
        }
    }
    fields.push(current);
    fields
}

#[cfg(test)]
#[allow(
    clippy::unwrap_used,
    clippy::expect_used,
    clippy::indexing_slicing,
    clippy::float_cmp
)]
mod tests {
    use super::*;

    #[test]
    fn test_parse_csv_basic() {
        let data = b"Name,Age,City\nAlice,30,NYC\nBob,25,LA";
        let wb = parse_delimited(data, Delimiter::Comma).unwrap();
        assert_eq!(wb.sheets.len(), 1);
        let sheet = &wb.sheets[0];
        assert_eq!(sheet.name, "CSV");
        assert_eq!(sheet.max_row, 2);
        assert_eq!(sheet.max_col, 2);
        // "Alice" is at (1, 0)
        let alice = sheet.cells.iter().find(|c| c.r == 1 && c.c == 0).unwrap();
        assert_eq!(alice.cell.v.as_deref(), Some("Alice"));
        // "30" is parsed as number
        let age = sheet.cells.iter().find(|c| c.r == 1 && c.c == 1).unwrap();
        assert!(matches!(age.cell.t, CellType::Number));
    }

    #[test]
    fn test_parse_tsv() {
        let data = b"A\tB\n1\t2";
        let wb = parse_delimited(data, Delimiter::Tab).unwrap();
        assert_eq!(wb.sheets[0].name, "TSV");
        assert_eq!(wb.sheets[0].cells.len(), 4);
    }

    #[test]
    fn test_quoted_csv() {
        let data = b"\"Hello, World\",42\n\"She said \"\"hi\"\"\",0";
        let wb = parse_delimited(data, Delimiter::Comma).unwrap();
        let cell = wb.sheets[0]
            .cells
            .iter()
            .find(|c| c.r == 0 && c.c == 0)
            .unwrap();
        assert_eq!(cell.cell.v.as_deref(), Some("Hello, World"));
        let cell2 = wb.sheets[0]
            .cells
            .iter()
            .find(|c| c.r == 1 && c.c == 0)
            .unwrap();
        assert_eq!(cell2.cell.v.as_deref(), Some("She said \"hi\""));
    }

    #[test]
    fn test_empty_csv() {
        let data = b"";
        let wb = parse_delimited(data, Delimiter::Comma).unwrap();
        assert_eq!(wb.sheets[0].cells.len(), 0);
    }
}
