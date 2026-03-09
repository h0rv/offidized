use std::path::Path;

use anyhow::{bail, Result};
use offidized_xlsx::CellValue;

use crate::format::FileFormat;
use crate::output::{CellOutput, ParagraphOutput, SheetCells};
use crate::range_parse::{parse_paragraph_range, parse_range};

pub fn run(path: &Path, range: Option<&str>, format: &str, paragraphs: Option<&str>) -> Result<()> {
    let file_format = FileFormat::detect(path)?;
    match file_format {
        FileFormat::Xlsx => read_xlsx(path, range, format),
        FileFormat::Docx => read_docx(path, paragraphs),
        FileFormat::Pptx => bail!("pptx read is not yet implemented"),
    }
}

fn cell_value_to_output(cell_ref: &str, value: Option<&CellValue>) -> CellOutput {
    match value {
        None | Some(CellValue::Blank) => CellOutput {
            cell_ref: cell_ref.to_string(),
            value: serde_json::Value::Null,
            value_type: "blank".to_string(),
        },
        Some(CellValue::String(s)) => CellOutput {
            cell_ref: cell_ref.to_string(),
            value: serde_json::Value::String(s.clone()),
            value_type: "string".to_string(),
        },
        Some(CellValue::Number(n)) => CellOutput {
            cell_ref: cell_ref.to_string(),
            value: serde_json::json!(*n),
            value_type: "number".to_string(),
        },
        Some(CellValue::Bool(b)) => CellOutput {
            cell_ref: cell_ref.to_string(),
            value: serde_json::Value::Bool(*b),
            value_type: "bool".to_string(),
        },
        Some(CellValue::Date(d)) => CellOutput {
            cell_ref: cell_ref.to_string(),
            value: serde_json::Value::String(d.clone()),
            value_type: "date".to_string(),
        },
        Some(CellValue::DateTime(serial)) => CellOutput {
            cell_ref: cell_ref.to_string(),
            value: serde_json::json!(*serial),
            value_type: "datetime".to_string(),
        },
        Some(CellValue::Error(e)) => CellOutput {
            cell_ref: cell_ref.to_string(),
            value: serde_json::Value::String(e.clone()),
            value_type: "error".to_string(),
        },
        Some(CellValue::RichText(runs)) => {
            let text: String = runs.iter().map(|r| r.text()).collect();
            CellOutput {
                cell_ref: cell_ref.to_string(),
                value: serde_json::Value::String(text),
                value_type: "richtext".to_string(),
            }
        }
    }
}

fn read_xlsx(path: &Path, range: Option<&str>, format: &str) -> Result<()> {
    let wb = offidized_xlsx::Workbook::open(path)?;

    if let Some(range_str) = range {
        let parsed = parse_range(range_str)?;
        let ws = wb
            .worksheets()
            .iter()
            .find(|ws| ws.name() == parsed.sheet)
            .ok_or_else(|| anyhow::anyhow!("sheet not found: {:?}", parsed.sheet))?;

        let cells: Vec<CellOutput> = ws
            .cells()
            .filter(|(cell_ref, _)| cell_ref_in_range(cell_ref, &parsed.range))
            .map(|(cell_ref, cell)| cell_value_to_output(cell_ref, cell.value()))
            .collect();

        if format == "csv" {
            print_csv(&cells)?;
        } else {
            println!("{}", serde_json::to_string(&cells)?);
        }
    } else {
        // Read all sheets.
        let sheets: Vec<SheetCells> = wb
            .worksheets()
            .iter()
            .map(|ws| {
                let cells: Vec<CellOutput> = ws
                    .cells()
                    .map(|(cell_ref, cell)| cell_value_to_output(cell_ref, cell.value()))
                    .collect();
                SheetCells {
                    name: ws.name().to_string(),
                    cells,
                }
            })
            .collect();

        if format == "csv" {
            for sheet in &sheets {
                print_csv(&sheet.cells)?;
            }
        } else {
            println!(
                "{}",
                serde_json::to_string(&serde_json::json!({"sheets": sheets}))?
            );
        }
    }

    Ok(())
}

/// Check if a cell reference like "A1" is within a range like "A1:D10" or a single cell "A1".
fn cell_ref_in_range(cell_ref: &str, range: &str) -> bool {
    if let Some((start, end)) = range.split_once(':') {
        let (start_col, start_row) = split_cell_ref(start);
        let (end_col, end_row) = split_cell_ref(end);
        let (cell_col, cell_row) = split_cell_ref(cell_ref);

        cell_col >= start_col && cell_col <= end_col && cell_row >= start_row && cell_row <= end_row
    } else {
        cell_ref.eq_ignore_ascii_case(range)
    }
}

/// Split a cell reference like "AB12" or "$AB$12" into (column_number, row_number).
fn split_cell_ref(cell_ref: &str) -> (u32, u32) {
    let stripped = cell_ref.replace('$', "");
    let col_end = stripped
        .find(|c: char| c.is_ascii_digit())
        .unwrap_or(stripped.len());
    let col_str = &stripped[..col_end];
    let row_str = &stripped[col_end..];

    let col = col_str.bytes().fold(0u32, |acc, b| {
        acc * 26 + u32::from(b.to_ascii_uppercase() - b'A') + 1
    });
    let row = row_str.parse::<u32>().unwrap_or(0);
    (col, row)
}

fn print_csv(cells: &[CellOutput]) -> Result<()> {
    for cell in cells {
        let val = match &cell.value {
            serde_json::Value::Null => String::new(),
            serde_json::Value::String(s) => s.clone(),
            other => other.to_string(),
        };
        println!("{},{}", cell.cell_ref, val);
    }
    Ok(())
}

fn read_docx(path: &Path, paragraphs_filter: Option<&str>) -> Result<()> {
    let doc = offidized_docx::Document::open(path)?;
    let paragraphs = doc.paragraphs();

    let (start, end) = if let Some(filter) = paragraphs_filter {
        parse_paragraph_range(filter)?
    } else {
        (0, paragraphs.len().saturating_sub(1))
    };

    let output: Vec<ParagraphOutput> = paragraphs
        .iter()
        .enumerate()
        .filter(|(i, _)| *i >= start && *i <= end)
        .map(|(i, p)| ParagraphOutput {
            index: i,
            text: p.text(),
            style: p.style_id().map(String::from),
        })
        .collect();

    println!("{}", serde_json::to_string(&output)?);
    Ok(())
}
