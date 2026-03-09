use std::path::Path;

use anyhow::Result;
use offidized_xlsx::Workbook;
use serde::Serialize;

#[derive(Serialize)]
struct PivotTableInfo {
    sheet: String,
    name: String,
    source_reference: String,
    target_row: u32,
    target_col: u32,
    row_fields: Vec<String>,
    column_fields: Vec<String>,
    page_fields: Vec<String>,
    data_fields: Vec<String>,
}

pub fn run(path: &Path, sheet_name: Option<&str>) -> Result<()> {
    let wb = Workbook::open(path)?;

    let mut pivot_infos = Vec::new();

    for ws in wb.worksheets() {
        if let Some(filter_sheet) = sheet_name {
            if ws.name() != filter_sheet {
                continue;
            }
        }

        for pivot in ws.pivot_tables() {
            pivot_infos.push(PivotTableInfo {
                sheet: ws.name().to_string(),
                name: pivot.name().to_string(),
                source_reference: format!("{:?}", pivot.source_reference()),
                target_row: pivot.target_row(),
                target_col: pivot.target_col(),
                row_fields: pivot
                    .row_fields()
                    .iter()
                    .map(|f| f.name().to_string())
                    .collect(),
                column_fields: pivot
                    .column_fields()
                    .iter()
                    .map(|f| f.name().to_string())
                    .collect(),
                page_fields: pivot
                    .page_fields()
                    .iter()
                    .map(|f| f.name().to_string())
                    .collect(),
                data_fields: pivot
                    .data_fields()
                    .iter()
                    .map(|f| f.field_name().to_string())
                    .collect(),
            });
        }
    }

    println!("{}", serde_json::to_string_pretty(&pivot_infos)?);
    Ok(())
}
