//! Clipboard and cell-value resolution for `XlView`.
//!
//! Handles Ctrl+C copy, TSV formatting, and converting raw cell values
//! into display strings.

#[cfg(target_arch = "wasm32")]
use super::{SharedState, XlView};
#[cfg(target_arch = "wasm32")]
use crate::types::workbook::{Cell, CellRawValue, CompiledFormat};

/// Format a number using the compiled format from offidized-xlsx.
#[cfg(target_arch = "wasm32")]
fn format_number_compiled(n: f64, fmt: &CompiledFormat, date1904: bool) -> String {
    offidized_xlsx::format_value_compiled(n, fmt, date1904).text
}

#[cfg(target_arch = "wasm32")]
impl XlView {
    pub(crate) fn get_selected_values_from_state(s: &mut SharedState) -> Option<String> {
        let start = s.selection_start?;
        let end = s.selection_end?;
        let sel_min_row = start.0.min(end.0);
        let sel_max_row = start.0.max(end.0);
        let sel_min_col = start.1.min(end.1);
        let sel_max_col = start.1.max(end.1);
        let workbook = s.workbook.as_mut()?;
        let date1904 = workbook.date1904;
        let shared_strings = &workbook.shared_strings;
        let numfmt_cache = &workbook.numfmt_cache;
        let sheet = workbook.sheets.get_mut(s.active_sheet)?;

        // Find the actual content bounds within the selection (trim trailing empty cells)
        // This matches Excel/Sheets behavior: only copy up to the last non-empty cell
        let mut content_max_row = sel_min_row;
        let mut content_max_col = sel_min_col;

        for cell_data in &sheet.cells {
            let row = cell_data.r;
            let col = cell_data.c;
            // Check if cell is within selection bounds
            if row >= sel_min_row && row <= sel_max_row && col >= sel_min_col && col <= sel_max_col
            {
                // Check if cell has content
                let has_content = cell_data.cell.raw.is_some()
                    || cell_data.cell.v.is_some()
                    || cell_data.cell.cached_display.is_some();
                if has_content {
                    content_max_row = content_max_row.max(row);
                    content_max_col = content_max_col.max(col);
                }
            }
        }

        // Use the trimmed bounds
        let min_row = sel_min_row;
        let max_row = content_max_row;
        let min_col = sel_min_col;
        let max_col = content_max_col;

        // If no content found, return empty
        if max_row < min_row || max_col < min_col {
            return Some(String::new());
        }

        let mut result = String::new();
        for row in min_row..=max_row {
            for col in min_col..=max_col {
                if col > min_col {
                    result.push('\t');
                }
                if let Some(cell_idx) = sheet.cell_index_at(row, col) {
                    if let Some(cell_data) = sheet.cells.get_mut(cell_idx) {
                        if let Some(v) = Self::resolve_cell_display_value(
                            &mut cell_data.cell,
                            shared_strings,
                            numfmt_cache,
                            date1904,
                        ) {
                            // Escape cell value for TSV format (like Excel does)
                            let escaped = Self::escape_cell_value(&v);
                            result.push_str(&escaped);
                        }
                    }
                }
            }
            if row < max_row {
                result.push('\n');
            }
        }
        Some(result)
    }

    /// Escape a cell value for TSV/clipboard format
    /// If the value contains tabs, newlines, or quotes, wrap in quotes and escape internal quotes
    pub(crate) fn escape_cell_value(value: &str) -> String {
        // Check if value needs quoting (contains tab, newline, or quote)
        let needs_quoting = value.contains('\t')
            || value.contains('\n')
            || value.contains('\r')
            || value.contains('"');

        if needs_quoting {
            // Wrap in quotes and double any internal quotes
            let escaped = value.replace('"', "\"\"");
            format!("\"{}\"", escaped)
        } else {
            value.to_string()
        }
    }

    pub(crate) fn copy_to_clipboard_internal(text: &str) {
        if let Some(window) = web_sys::window() {
            let clipboard = window.navigator().clipboard();
            let _ = clipboard.write_text(text);
        }
    }

    pub(crate) fn resolve_cell_display_value(
        cell: &mut Cell,
        shared_strings: &[String],
        numfmt_cache: &[CompiledFormat],
        date1904: bool,
    ) -> Option<String> {
        if let Some(v) = cell.v.as_ref() {
            return Some(v.clone());
        }
        if let Some(cached) = cell.cached_display.as_ref() {
            return Some(cached.clone());
        }
        let raw = cell.raw.as_ref()?;
        let display = match raw {
            CellRawValue::SharedString(idx) => shared_strings.get(*idx as usize).cloned(),
            CellRawValue::String(s) => Some(s.clone()),
            CellRawValue::Boolean(b) => Some(if *b {
                "TRUE".to_string()
            } else {
                "FALSE".to_string()
            }),
            CellRawValue::Error(e) => Some(e.clone()),
            CellRawValue::Number(n) => {
                let compiled = cell
                    .style_idx
                    .and_then(|idx| numfmt_cache.get(idx as usize));
                let formatted = match compiled {
                    Some(fmt) if matches!(fmt, CompiledFormat::General) => {
                        cell.cached_display.clone().unwrap_or_else(|| n.to_string())
                    }
                    Some(fmt) => format_number_compiled(*n, fmt, date1904),
                    None => cell.cached_display.clone().unwrap_or_else(|| n.to_string()),
                };
                cell.cached_display = Some(formatted.clone());
                Some(formatted)
            }
            CellRawValue::Date(n) => {
                let compiled = cell
                    .style_idx
                    .and_then(|idx| numfmt_cache.get(idx as usize));
                let formatted = match compiled {
                    Some(fmt) => format_number_compiled(*n, fmt, date1904),
                    None => n.to_string(),
                };
                cell.cached_display = Some(formatted.clone());
                Some(formatted)
            }
        };
        display
    }
}
