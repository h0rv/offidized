//! Cell editing mutations.
//!
//! Applies user edits to the in-memory workbook model.

use crate::error::Result;
use crate::types::{Cell, CellData, CellRawValue, CellType, Workbook};

/// Apply a cell edit to the workbook.
///
/// Detects the value type automatically:
/// - Empty string -> clears the cell
/// - Parseable as f64 -> Number
/// - "true"/"false" (case-insensitive) -> Boolean
/// - Otherwise -> String
pub(crate) fn apply_cell_edit(
    workbook: &mut Workbook,
    sheet_idx: usize,
    row: u32,
    col: u32,
    value: &str,
) -> Result<()> {
    let sheet = workbook
        .sheets
        .get_mut(sheet_idx)
        .ok_or_else(|| crate::error::XlViewError::Other("sheet index out of range".into()))?;

    let trimmed = value.trim();

    if trimmed.is_empty() {
        // Clear cell: remove it from the cells vec
        if let Some(idx) = sheet.cell_index_at(row, col) {
            sheet.cells.remove(idx);
            sheet.rebuild_cell_index();
        }
        return Ok(());
    }

    // Detect type
    let (cell_type, raw_value, display) = detect_cell_type(trimmed);

    if let Some(idx) = sheet.cell_index_at(row, col) {
        // Update existing cell
        if let Some(cd) = sheet.cells.get_mut(idx) {
            cd.cell.t = cell_type;
            cd.cell.v = Some(display.clone());
            cd.cell.raw = Some(raw_value);
            cd.cell.cached_display = Some(display);
            cd.cell.cached_rich_text = None;
            cd.cell.rich_text = None;
            cd.cell.formula = None; // Editing clears the formula
        }
    } else {
        // Insert new cell
        sheet.cells.push(CellData {
            r: row,
            c: col,
            cell: Cell {
                v: Some(display.clone()),
                t: cell_type,
                s: None,
                style_idx: None,
                raw: Some(raw_value),
                cached_display: Some(display),
                rich_text: None,
                cached_rich_text: None,
                has_comment: None,
                hyperlink: None,
                formula: None,
            },
        });

        // Update max_row / max_col
        if row + 1 > sheet.max_row {
            sheet.max_row = row + 1;
        }
        if col + 1 > sheet.max_col {
            sheet.max_col = col + 1;
        }

        sheet.rebuild_cell_index();
    }

    Ok(())
}

/// Restore a cell to a previous state (for undo/redo).
///
/// Returns the cell that was at the position before restoration (for the reverse operation).
/// If `cell` is `Some`, replaces or inserts the cell. If `None`, removes the cell.
pub(crate) fn restore_cell(
    workbook: &mut Workbook,
    sheet_idx: usize,
    row: u32,
    col: u32,
    cell: Option<Cell>,
) -> Result<Option<Cell>> {
    let sheet = workbook
        .sheets
        .get_mut(sheet_idx)
        .ok_or_else(|| crate::error::XlViewError::Other("sheet index out of range".into()))?;

    // Snapshot the current cell at this position
    let current = sheet
        .cell_index_at(row, col)
        .and_then(|idx| sheet.cells.get(idx))
        .map(|cd| cd.cell.clone());

    match cell {
        Some(new_cell) => {
            // Replace or insert
            if let Some(idx) = sheet.cell_index_at(row, col) {
                if let Some(cd) = sheet.cells.get_mut(idx) {
                    cd.cell = new_cell;
                }
            } else {
                sheet.cells.push(CellData {
                    r: row,
                    c: col,
                    cell: new_cell,
                });
                sheet.rebuild_cell_index();
            }
        }
        None => {
            // Remove the cell
            if let Some(idx) = sheet.cell_index_at(row, col) {
                sheet.cells.remove(idx);
                sheet.rebuild_cell_index();
            }
        }
    }

    Ok(current)
}

/// Detect the appropriate cell type and create raw value.
fn detect_cell_type(value: &str) -> (CellType, CellRawValue, String) {
    // Boolean
    if value.eq_ignore_ascii_case("true") {
        return (
            CellType::Boolean,
            CellRawValue::Boolean(true),
            "TRUE".into(),
        );
    }
    if value.eq_ignore_ascii_case("false") {
        return (
            CellType::Boolean,
            CellRawValue::Boolean(false),
            "FALSE".into(),
        );
    }

    // Number
    if let Ok(n) = value.parse::<f64>() {
        return (CellType::Number, CellRawValue::Number(n), value.into());
    }

    // Default: String
    (
        CellType::String,
        CellRawValue::String(value.into()),
        value.into(),
    )
}
