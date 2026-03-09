//! Generates worksheet XML from a `Sheet` struct.
//!
//! Modified sheets use inline strings (`t="inlineStr"`) instead of shared
//! string references, avoiding the need to rebuild the shared string table.

use crate::error::Result;
use crate::render::canvas::headers::col_to_letter;
use crate::types::{Cell, CellType, Sheet};

/// Write a complete worksheet XML string from a `Sheet`.
pub(crate) fn write_sheet_xml(sheet: &Sheet) -> Result<String> {
    let mut out = String::with_capacity(4096);
    out.push_str(r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#);
    out.push('\n');
    out.push_str(
        r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" "#,
    );
    out.push_str(
        r#"xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">"#,
    );
    out.push('\n');

    // <dimension>
    if sheet.max_row > 0 || sheet.max_col > 0 {
        let end_col = col_to_letter(sheet.max_col.saturating_sub(1));
        write_escaped_fmt(
            &mut out,
            &format!("<dimension ref=\"A1:{}{}\"/>", end_col, sheet.max_row),
        );
        out.push('\n');
    }

    // <sheetViews> -- frozen panes
    if sheet.frozen_rows > 0 || sheet.frozen_cols > 0 {
        out.push_str("<sheetViews><sheetView tabSelected=\"1\" workbookViewId=\"0\">");
        let top_left = format!(
            "{}{}",
            col_to_letter(sheet.frozen_cols),
            sheet.frozen_rows + 1
        );
        out.push_str(&format!(
            "<pane xSplit=\"{}\" ySplit=\"{}\" topLeftCell=\"{}\" state=\"frozen\"/>",
            sheet.frozen_cols, sheet.frozen_rows, top_left
        ));
        out.push_str("</sheetView></sheetViews>\n");
    }

    // <sheetFormatPr>
    out.push_str(&format!(
        "<sheetFormatPr defaultRowHeight=\"{:.2}\" defaultColWidth=\"{:.4}\"/>\n",
        sheet.default_row_height, sheet.default_col_width
    ));

    // <cols>
    if !sheet.col_widths.is_empty() || !sheet.hidden_cols.is_empty() {
        out.push_str("<cols>\n");
        for cw in &sheet.col_widths {
            let col1 = cw.col + 1; // XLSX is 1-based
            let hidden = if sheet.hidden_cols.contains(&cw.col) {
                " hidden=\"1\""
            } else {
                ""
            };
            out.push_str(&format!(
                "<col min=\"{}\" max=\"{}\" width=\"{:.4}\" customWidth=\"1\"{}/>\n",
                col1,
                col1,
                cw.width / (7.0_f64 / 0.75), // px back to Excel character units (approx)
                hidden
            ));
        }
        out.push_str("</cols>\n");
    }

    // <sheetData>
    out.push_str("<sheetData>\n");
    write_sheet_data(&mut out, sheet);
    out.push_str("</sheetData>\n");

    // <mergeCells>
    if !sheet.merges.is_empty() {
        out.push_str(&format!("<mergeCells count=\"{}\">\n", sheet.merges.len()));
        for merge in &sheet.merges {
            let start_col = col_to_letter(merge.start_col);
            let end_col = col_to_letter(merge.end_col);
            out.push_str(&format!(
                "<mergeCell ref=\"{}{}:{}{}\"/>\n",
                start_col,
                merge.start_row + 1,
                end_col,
                merge.end_row + 1
            ));
        }
        out.push_str("</mergeCells>\n");
    }

    // <hyperlinks>
    if !sheet.hyperlinks.is_empty() {
        out.push_str("<hyperlinks>\n");
        for (idx, hdef) in sheet.hyperlinks.iter().enumerate() {
            if hdef.hyperlink.is_external {
                out.push_str(&format!(
                    "<hyperlink ref=\"{}\" r:id=\"rId{}\"/>\n",
                    xml_escape(&hdef.cell_ref),
                    idx + 1
                ));
            } else if let Some(ref loc) = hdef.hyperlink.location {
                out.push_str(&format!(
                    "<hyperlink ref=\"{}\" location=\"{}\"/>\n",
                    xml_escape(&hdef.cell_ref),
                    xml_escape(loc)
                ));
            }
        }
        out.push_str("</hyperlinks>\n");
    }

    out.push_str("</worksheet>");
    Ok(out)
}

/// Write all cell rows into `<sheetData>`.
fn write_sheet_data(out: &mut String, sheet: &Sheet) {
    if sheet.cells.is_empty() {
        return;
    }

    // Group cells by row
    let mut rows: Vec<(u32, Vec<usize>)> = Vec::new();
    for (idx, cd) in sheet.cells.iter().enumerate() {
        if let Some(last) = rows.last_mut() {
            if last.0 == cd.r {
                last.1.push(idx);
                continue;
            }
        }
        rows.push((cd.r, vec![idx]));
    }

    for (row, cell_indices) in &rows {
        // Check for custom row height
        let ht = sheet
            .row_heights
            .iter()
            .find(|rh| rh.row == *row)
            .map(|rh| rh.height);
        let hidden = sheet.hidden_rows.contains(row);

        out.push_str(&format!("<row r=\"{}\"", row + 1));
        if let Some(h) = ht {
            // Convert pixels back to points
            #[allow(clippy::cast_possible_truncation)]
            let pts = h * (72.0 / 96.0);
            out.push_str(&format!(" ht=\"{pts:.2}\" customHeight=\"1\""));
        }
        if hidden {
            out.push_str(" hidden=\"1\"");
        }
        out.push('>');

        for &idx in cell_indices {
            if let Some(cd) = sheet.cells.get(idx) {
                write_cell(out, cd.r, cd.c, &cd.cell);
            }
        }

        out.push_str("</row>\n");
    }
}

/// Write a single `<c>` element.
fn write_cell(out: &mut String, row: u32, col: u32, cell: &Cell) {
    let col_letter = col_to_letter(col);
    let cell_ref = format!("{}{}", col_letter, row + 1);

    out.push_str(&format!("<c r=\"{}\"", cell_ref));

    // Style index
    if let Some(si) = cell.style_idx {
        out.push_str(&format!(" s=\"{}\"", si));
    }

    // Determine type attribute and value writing strategy
    match cell.t {
        CellType::String => {
            // Use inline string to avoid shared string table rebuild
            out.push_str(" t=\"inlineStr\">");
            // Formula (if any)
            if let Some(ref f) = cell.formula {
                out.push_str("<f>");
                out.push_str(&xml_escape(f));
                out.push_str("</f>");
            }
            // Inline string value
            let display = cell_display_value(cell);
            out.push_str("<is><t>");
            out.push_str(&xml_escape(&display));
            out.push_str("</t></is>");
        }
        CellType::Number | CellType::Date => {
            if cell.t == CellType::Date {
                out.push_str(" t=\"d\"");
            }
            out.push('>');
            if let Some(ref f) = cell.formula {
                out.push_str("<f>");
                out.push_str(&xml_escape(f));
                out.push_str("</f>");
            }
            let val = cell_raw_number(cell);
            out.push_str(&format!("<v>{val}</v>"));
        }
        CellType::Boolean => {
            out.push_str(" t=\"b\">");
            if let Some(ref f) = cell.formula {
                out.push_str("<f>");
                out.push_str(&xml_escape(f));
                out.push_str("</f>");
            }
            let val = match cell.v.as_deref() {
                Some("TRUE") | Some("true") | Some("1") => "1",
                _ => "0",
            };
            out.push_str(&format!("<v>{val}</v>"));
        }
        CellType::Error => {
            out.push_str(" t=\"e\">");
            if let Some(ref f) = cell.formula {
                out.push_str("<f>");
                out.push_str(&xml_escape(f));
                out.push_str("</f>");
            }
            let val = cell.v.as_deref().unwrap_or("#VALUE!");
            out.push_str(&format!("<v>{}</v>", xml_escape(val)));
        }
    }

    out.push_str("</c>");
}

/// Get display value from a cell (prefers cached_display, then v).
fn cell_display_value(cell: &Cell) -> String {
    cell.cached_display
        .as_ref()
        .or(cell.v.as_ref())
        .cloned()
        .unwrap_or_default()
}

/// Get the raw numeric value from a cell for `<v>` output.
fn cell_raw_number(cell: &Cell) -> String {
    use crate::types::CellRawValue;
    if let Some(ref raw) = cell.raw {
        match raw {
            CellRawValue::Number(n) | CellRawValue::Date(n) => return n.to_string(),
            CellRawValue::Boolean(b) => return if *b { "1".into() } else { "0".into() },
            CellRawValue::SharedString(idx) => return idx.to_string(),
            CellRawValue::String(s) | CellRawValue::Error(s) => return s.clone(),
        }
    }
    cell.v.as_deref().unwrap_or("0").to_string()
}

/// Minimal XML escaping for attribute/text content.
fn xml_escape(s: &str) -> String {
    let mut out = String::with_capacity(s.len());
    for c in s.chars() {
        match c {
            '&' => out.push_str("&amp;"),
            '<' => out.push_str("&lt;"),
            '>' => out.push_str("&gt;"),
            '"' => out.push_str("&quot;"),
            '\'' => out.push_str("&apos;"),
            _ => out.push(c),
        }
    }
    out
}

/// Helper: write a pre-escaped string directly (used for dimension ref etc.)
fn write_escaped_fmt(out: &mut String, s: &str) {
    out.push_str(s);
}
