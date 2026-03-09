use std::path::Path;

use anyhow::Result;
use offidized_xlsx::{CellValue, Workbook};
use serde::Serialize;

#[derive(Serialize)]
struct EvalResult {
    formula: String,
    result: serde_json::Value,
    #[serde(rename = "type")]
    result_type: String,
}

pub fn run(path: &Path, formula: &str, sheet_name: Option<&str>) -> Result<()> {
    let wb = Workbook::open(path)?;

    let sheet = match sheet_name {
        Some(name) => name.to_string(),
        None => wb
            .worksheets()
            .first()
            .ok_or_else(|| anyhow::anyhow!("workbook has no sheets"))?
            .name()
            .to_string(),
    };

    let result = wb.evaluate_formula(formula, &sheet, 1, 1);

    let (result_value, result_type) = cell_value_to_json(&result);

    let eval_result = EvalResult {
        formula: formula.to_string(),
        result: result_value,
        result_type,
    };

    println!("{}", serde_json::to_string_pretty(&eval_result)?);
    Ok(())
}

fn cell_value_to_json(value: &CellValue) -> (serde_json::Value, String) {
    match value {
        CellValue::Blank => (serde_json::Value::Null, "blank".to_string()),
        CellValue::String(s) => (serde_json::json!(s), "string".to_string()),
        CellValue::Number(n) => (serde_json::json!(n), "number".to_string()),
        CellValue::Bool(b) => (serde_json::json!(b), "bool".to_string()),
        CellValue::Date(d) => (serde_json::json!(d), "date".to_string()),
        CellValue::DateTime(dt) => (serde_json::json!(dt), "datetime".to_string()),
        CellValue::Error(e) => (serde_json::json!(e), "error".to_string()),
        CellValue::RichText(runs) => {
            let text: String = runs.iter().map(|r| r.text()).collect();
            (serde_json::json!(text), "richtext".to_string())
        }
    }
}
