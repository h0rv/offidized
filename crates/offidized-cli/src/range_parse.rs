use anyhow::{bail, Result};

/// Parsed sheet + cell range from a string like `"Sheet1!A1:D10"` or `"Sheet1!A1"`.
#[derive(Debug, Clone)]
pub struct SheetRange {
    pub sheet: String,
    pub range: String,
}

/// Parse a range string in the form `"Sheet1!A1:D10"` or `"Sheet1!A1"`.
///
/// If no `!` separator is present, the entire string is treated as a cell/range
/// reference on the first sheet.
pub fn parse_range(input: &str) -> Result<SheetRange> {
    if let Some((sheet, range)) = input.split_once('!') {
        if sheet.is_empty() {
            bail!("empty sheet name in range: {input:?}");
        }
        if range.is_empty() {
            bail!("empty cell reference in range: {input:?}");
        }
        Ok(SheetRange {
            sheet: sheet.to_string(),
            range: range.to_string(),
        })
    } else {
        bail!("range must include sheet name separated by `!`, e.g. \"Sheet1!A1:D10\"; got: {input:?}");
    }
}

/// Parse a simple paragraph range like `"1-5"` (0-based, inclusive on both ends).
pub fn parse_paragraph_range(input: &str) -> Result<(usize, usize)> {
    if let Some((start_str, end_str)) = input.split_once('-') {
        let start: usize = start_str
            .trim()
            .parse()
            .map_err(|_| anyhow::anyhow!("invalid paragraph range start: {start_str:?}"))?;
        let end: usize = end_str
            .trim()
            .parse()
            .map_err(|_| anyhow::anyhow!("invalid paragraph range end: {end_str:?}"))?;
        Ok((start, end))
    } else {
        let index: usize = input
            .trim()
            .parse()
            .map_err(|_| anyhow::anyhow!("invalid paragraph index: {input:?}"))?;
        Ok((index, index))
    }
}
