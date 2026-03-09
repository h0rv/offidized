//! XLSX export pipeline.
//!
//! Produces a modified XLSX by patching the original ZIP archive.
//! Only dirty (edited) sheets are re-serialized; everything else is
//! passed through byte-identical.

pub(crate) mod sheet_writer;
pub(crate) mod zip_patcher;

use std::collections::HashSet;

use crate::error::Result;
use crate::types::Workbook;

/// Save a workbook to XLSX bytes.
///
/// `original_bytes` is the original XLSX data (needed for ZIP roundtrip).
/// `dirty_sheets` is the set of sheet indices that were modified.
///
/// Returns the new XLSX file as `Vec<u8>`.
pub(crate) fn save_xlsx(
    original_bytes: &[u8],
    workbook: &Workbook,
    dirty_sheets: &HashSet<usize>,
) -> Result<Vec<u8>> {
    if dirty_sheets.is_empty() {
        // Nothing changed -- return original bytes
        return Ok(original_bytes.to_vec());
    }

    zip_patcher::patch_zip(original_bytes, workbook, dirty_sheets)
}
