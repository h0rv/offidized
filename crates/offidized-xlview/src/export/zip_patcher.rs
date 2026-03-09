//! Patch an XLSX ZIP archive with modified sheet XML.
//!
//! Unmodified entries are copied via `raw_copy_file` (zero recompression cost).
//! Only dirty sheets get new XML generated and written.

use std::collections::HashSet;
use std::io::{Cursor, Write};
use zip::{ZipArchive, ZipWriter};

use crate::error::Result;
use crate::types::Workbook;

use super::sheet_writer::write_sheet_xml;

/// Patch the original XLSX bytes, replacing only sheets in `dirty_sheets`.
///
/// Returns the new XLSX file as `Vec<u8>`.
pub(crate) fn patch_zip(
    original_data: &[u8],
    workbook: &Workbook,
    dirty_sheets: &HashSet<usize>,
) -> Result<Vec<u8>> {
    let cursor = Cursor::new(original_data);
    let mut archive = ZipArchive::new(cursor)?;

    // Build set of ZIP paths that need replacement
    let dirty_paths: HashSet<&str> = dirty_sheets
        .iter()
        .filter_map(|&idx| workbook.sheet_paths.get(idx).map(String::as_str))
        .collect();

    let buf: Vec<u8> = Vec::with_capacity(original_data.len());
    let mut writer = ZipWriter::new(Cursor::new(buf));

    // Copy all entries, replacing dirty ones
    for i in 0..archive.len() {
        let entry = archive.by_index_raw(i)?;
        let name = entry.name().to_string();

        if dirty_paths.contains(name.as_str()) {
            // Find which sheet index this path corresponds to
            if let Some(sheet_idx) = workbook.sheet_paths.iter().position(|p| p == &name) {
                if let Some(sheet) = workbook.sheets.get(sheet_idx) {
                    let xml = write_sheet_xml(sheet)?;
                    let options: zip::write::FileOptions<'_, ()> =
                        zip::write::FileOptions::default()
                            .compression_method(zip::CompressionMethod::Deflated);
                    writer.start_file(&name, options)?;
                    writer.write_all(xml.as_bytes())?;
                    continue;
                }
            }
        }

        // Pass through unmodified entry (raw copy, no re-compression)
        writer.raw_copy_file(entry)?;
    }

    let cursor = writer.finish()?;
    Ok(cursor.into_inner())
}
