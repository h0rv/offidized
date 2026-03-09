//! Bridge between `offidized-formula` and the xlsx workbook model.
//!
//! Implements [`CellDataProvider`] so that formula evaluation can read cell
//! values directly from a [`Workbook`].

use crate::cell::CellValue;
use crate::workbook::Workbook;
use crate::worksheet::Worksheet;
use offidized_formula::{CellDataProvider, ScalarValue, XlError};

/// Provides cell data from a [`Workbook`] to the formula evaluator.
pub(crate) struct WorkbookProvider<'a> {
    workbook: &'a Workbook,
}

impl<'a> WorkbookProvider<'a> {
    pub(crate) fn new(workbook: &'a Workbook) -> Self {
        Self { workbook }
    }

    /// Resolves a sheet by name, or returns the first sheet if `None`.
    fn resolve_sheet(&self, sheet: Option<&str>) -> Option<&'a Worksheet> {
        match sheet {
            Some(name) => self.workbook.sheet(name),
            None => self.workbook.worksheets().first(),
        }
    }
}

impl CellDataProvider for WorkbookProvider<'_> {
    fn cell_value(&self, sheet: Option<&str>, row: u32, col: u32) -> ScalarValue {
        let Some(ws) = self.resolve_sheet(sheet) else {
            return ScalarValue::Error(XlError::Ref);
        };

        let Some(col_letters) = offidized_formula::reference::column_index_to_letters(col) else {
            return ScalarValue::Error(XlError::Ref);
        };
        let cell_ref = format!("{col_letters}{row}");

        let Some(cell) = ws.cell(&cell_ref) else {
            return ScalarValue::Blank;
        };

        match cell.value() {
            None => ScalarValue::Blank,
            Some(cv) => cell_value_to_scalar(cv),
        }
    }

    fn cell_formula(&self, sheet: Option<&str>, row: u32, col: u32) -> Option<String> {
        let ws = self.resolve_sheet(sheet)?;
        let col_letters = offidized_formula::reference::column_index_to_letters(col)?;
        let cell_ref = format!("{col_letters}{row}");
        let cell = ws.cell(&cell_ref)?;
        cell.formula().map(String::from)
    }

    fn resolve_name(&self, name: &str) -> Option<String> {
        self.workbook
            .defined_names()
            .iter()
            .find(|dn| dn.name().eq_ignore_ascii_case(name))
            .map(|dn| dn.reference().to_string())
    }

    fn sheet_name(&self, index: usize) -> Option<String> {
        self.workbook
            .worksheets()
            .get(index)
            .map(|ws| ws.name().to_string())
    }

    fn sheet_index(&self, name: &str) -> Option<usize> {
        self.workbook
            .worksheets()
            .iter()
            .position(|ws| ws.name() == name)
    }

    fn cell_info(
        &self,
        sheet: Option<&str>,
        row: u32,
        col: u32,
        info_type: &str,
    ) -> Option<String> {
        let ws = self.resolve_sheet(sheet)?;
        let col_letters = offidized_formula::reference::column_index_to_letters(col)?;
        let cell_ref = format!("{col_letters}{row}");

        match info_type {
            "address" => {
                // Return absolute address with sheet name if different from current
                let addr = if let Some(sheet_name) = sheet {
                    format!("{}!${col_letters}${row}", sheet_name)
                } else {
                    format!("${col_letters}${row}")
                };
                Some(addr)
            }
            "col" => Some(col.to_string()),
            "row" => Some(row.to_string()),
            "contents" => {
                let cell = ws.cell(&cell_ref)?;
                let value = cell.value()?;
                let text = match value {
                    CellValue::Blank => String::new(),
                    CellValue::String(s) => s.clone(),
                    CellValue::Number(n) => n.to_string(),
                    CellValue::Bool(b) => if *b { "TRUE" } else { "FALSE" }.to_string(),
                    CellValue::Date(d) => d.clone(),
                    CellValue::DateTime(dt) => dt.to_string(),
                    CellValue::Error(e) => e.clone(),
                    CellValue::RichText(runs) => runs.iter().map(|r| r.text()).collect(),
                };
                Some(text)
            }
            "type" => {
                let cell = ws.cell(&cell_ref);
                let type_char = match cell {
                    None => "b", // blank
                    Some(c) => match c.value() {
                        None => "b",
                        Some(CellValue::Blank) => "b",
                        Some(CellValue::String(_) | CellValue::RichText(_)) => "l", // label
                        Some(_) => "v",                                             // value
                    },
                };
                Some(type_char.to_string())
            }
            "format" => {
                let cell = ws.cell(&cell_ref)?;
                let style_id = cell.style_id()?;
                let style = self.workbook.styles().style(style_id)?;
                let format = style.number_format()?.to_string();
                Some(format)
            }
            "width" => {
                let column = ws.column(col)?;
                let width = column.width()?;
                Some(width.to_string())
            }
            _ => None,
        }
    }

    fn workbook_info(&self, info_type: &str) -> Option<String> {
        match info_type {
            "numfile" => {
                let count = self.workbook.worksheets().len();
                Some(count.to_string())
            }
            "recalc" => Some("Automatic".to_string()),
            "release" => Some("offidized-0.1.0".to_string()),
            "system" => {
                #[cfg(target_os = "macos")]
                return Some("mac".to_string());
                #[cfg(target_os = "ios")]
                return Some("mac".to_string());
                #[cfg(not(any(target_os = "macos", target_os = "ios")))]
                return Some("pcdos".to_string());
            }
            "osversion" => {
                let os = std::env::consts::OS;
                Some(os.to_string())
            }
            _ => None,
        }
    }
}

/// Converts an xlsx `CellValue` to a formula engine `ScalarValue`.
fn cell_value_to_scalar(cv: &CellValue) -> ScalarValue {
    match cv {
        CellValue::Blank => ScalarValue::Blank,
        CellValue::String(s) => ScalarValue::Text(s.clone()),
        CellValue::Number(n) => ScalarValue::Number(*n),
        CellValue::Bool(b) => ScalarValue::Bool(*b),
        CellValue::Date(_) => ScalarValue::Text(String::new()),
        CellValue::DateTime(serial) => ScalarValue::Number(*serial),
        CellValue::Error(e) => match XlError::parse(e) {
            Some(xl) => ScalarValue::Error(xl),
            None => ScalarValue::Error(XlError::Value),
        },
        CellValue::RichText(runs) => {
            let text: String = runs.iter().map(|r| r.text()).collect();
            ScalarValue::Text(text)
        }
    }
}

/// Converts a formula engine `ScalarValue` back to an xlsx `CellValue`.
pub(crate) fn scalar_to_cell_value(sv: &ScalarValue) -> CellValue {
    match sv {
        ScalarValue::Blank => CellValue::Blank,
        ScalarValue::Bool(b) => CellValue::Bool(*b),
        ScalarValue::Number(n) => CellValue::Number(*n),
        ScalarValue::Text(s) => CellValue::String(s.clone()),
        ScalarValue::Error(e) => CellValue::Error(e.to_string()),
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::workbook::Workbook;

    #[test]
    fn cell_value_to_scalar_conversions() {
        assert_eq!(cell_value_to_scalar(&CellValue::Blank), ScalarValue::Blank);
        assert_eq!(
            cell_value_to_scalar(&CellValue::Number(42.0)),
            ScalarValue::Number(42.0)
        );
        assert_eq!(
            cell_value_to_scalar(&CellValue::String("hello".to_string())),
            ScalarValue::Text("hello".to_string())
        );
        assert_eq!(
            cell_value_to_scalar(&CellValue::Bool(true)),
            ScalarValue::Bool(true)
        );
        assert_eq!(
            cell_value_to_scalar(&CellValue::Error("#DIV/0!".to_string())),
            ScalarValue::Error(XlError::Div0)
        );
        assert_eq!(
            cell_value_to_scalar(&CellValue::DateTime(44927.5)),
            ScalarValue::Number(44927.5)
        );
    }

    #[test]
    fn scalar_to_cell_value_conversions() {
        assert_eq!(scalar_to_cell_value(&ScalarValue::Blank), CellValue::Blank);
        assert_eq!(
            scalar_to_cell_value(&ScalarValue::Number(42.0)),
            CellValue::Number(42.0)
        );
        assert_eq!(
            scalar_to_cell_value(&ScalarValue::Text("hello".to_string())),
            CellValue::String("hello".to_string())
        );
        assert_eq!(
            scalar_to_cell_value(&ScalarValue::Bool(true)),
            CellValue::Bool(true)
        );
        assert_eq!(
            scalar_to_cell_value(&ScalarValue::Error(XlError::Na)),
            CellValue::Error("#N/A".to_string())
        );
    }

    #[test]
    fn evaluate_simple_formula_via_workbook() {
        let mut wb = Workbook::new();
        {
            let ws = wb.add_sheet("Sheet1");
            ws.cell_mut("A1").ok().map(|c| c.set_value(10.0));
            ws.cell_mut("A2").ok().map(|c| c.set_value(20.0));
            ws.cell_mut("A3").ok().map(|c| c.set_value(30.0));
        }

        let provider = WorkbookProvider::new(&wb);
        let ctx = offidized_formula::EvalContext::new(&provider, Some("Sheet1".to_string()), 4, 1);
        let result = offidized_formula::evaluate("=SUM(A1:A3)", &ctx);
        assert_eq!(result, offidized_formula::CalcValue::number(60.0));
    }

    #[test]
    fn evaluate_cell_ref_formula() {
        let mut wb = Workbook::new();
        {
            let ws = wb.add_sheet("Sheet1");
            ws.cell_mut("B2").ok().map(|c| c.set_value(42.0));
        }

        let provider = WorkbookProvider::new(&wb);
        let ctx = offidized_formula::EvalContext::new(&provider, Some("Sheet1".to_string()), 1, 1);
        let result = offidized_formula::evaluate("=B2*2", &ctx);
        assert_eq!(result, offidized_formula::CalcValue::number(84.0));
    }

    #[test]
    fn test_cell_function_col() {
        let mut wb = Workbook::new();
        wb.add_sheet("Sheet1");

        let provider = WorkbookProvider::new(&wb);
        let ctx = offidized_formula::EvalContext::new(&provider, Some("Sheet1".to_string()), 5, 3);
        // CELL("col", C5) should return 3
        let result = offidized_formula::evaluate("=CELL(\"col\", C5)", &ctx);
        assert_eq!(result, offidized_formula::CalcValue::text("3"));
    }

    #[test]
    fn test_cell_function_row() {
        let mut wb = Workbook::new();
        wb.add_sheet("Sheet1");

        let provider = WorkbookProvider::new(&wb);
        let ctx = offidized_formula::EvalContext::new(&provider, Some("Sheet1".to_string()), 5, 3);
        // CELL("row", C5) should return 5
        let result = offidized_formula::evaluate("=CELL(\"row\", C5)", &ctx);
        assert_eq!(result, offidized_formula::CalcValue::text("5"));
    }

    #[test]
    fn test_cell_function_address() {
        let mut wb = Workbook::new();
        wb.add_sheet("Sheet1");

        let provider = WorkbookProvider::new(&wb);
        let ctx = offidized_formula::EvalContext::new(&provider, Some("Sheet1".to_string()), 1, 1);
        // CELL("address", B3) should return "$B$3"
        let result = offidized_formula::evaluate("=CELL(\"address\", B3)", &ctx);
        assert_eq!(result, offidized_formula::CalcValue::text("Sheet1!$B$3"));
    }

    #[test]
    fn test_cell_function_contents() {
        let mut wb = Workbook::new();
        {
            let ws = wb.add_sheet("Sheet1");
            ws.cell_mut("A1").ok().map(|c| c.set_value("Hello"));
        }

        let provider = WorkbookProvider::new(&wb);
        let ctx = offidized_formula::EvalContext::new(&provider, Some("Sheet1".to_string()), 1, 1);
        // CELL("contents", A1) should return "Hello"
        let result = offidized_formula::evaluate("=CELL(\"contents\", A1)", &ctx);
        assert_eq!(result, offidized_formula::CalcValue::text("Hello"));
    }

    #[test]
    fn test_cell_function_type() {
        let mut wb = Workbook::new();
        {
            let ws = wb.add_sheet("Sheet1");
            ws.cell_mut("A1").ok().map(|c| c.set_value("text"));
            ws.cell_mut("B1").ok().map(|c| c.set_value(42.0));
        }

        let provider = WorkbookProvider::new(&wb);
        let ctx = offidized_formula::EvalContext::new(&provider, Some("Sheet1".to_string()), 1, 1);
        // CELL("type", A1) should return "l" (label)
        let result = offidized_formula::evaluate("=CELL(\"type\", A1)", &ctx);
        assert_eq!(result, offidized_formula::CalcValue::text("l"));
        // CELL("type", B1) should return "v" (value)
        let result = offidized_formula::evaluate("=CELL(\"type\", B1)", &ctx);
        assert_eq!(result, offidized_formula::CalcValue::text("v"));
        // CELL("type", C1) should return "b" (blank)
        let result = offidized_formula::evaluate("=CELL(\"type\", C1)", &ctx);
        assert_eq!(result, offidized_formula::CalcValue::text("b"));
    }

    #[test]
    fn test_info_function_numfile() {
        let mut wb = Workbook::new();
        wb.add_sheet("Sheet1");
        wb.add_sheet("Sheet2");
        wb.add_sheet("Sheet3");

        let provider = WorkbookProvider::new(&wb);
        let ctx = offidized_formula::EvalContext::new(&provider, Some("Sheet1".to_string()), 1, 1);
        // INFO("numfile") should return 3
        let result = offidized_formula::evaluate("=INFO(\"numfile\")", &ctx);
        assert_eq!(result, offidized_formula::CalcValue::text("3"));
    }

    #[test]
    fn test_info_function_recalc() {
        let mut wb = Workbook::new();
        wb.add_sheet("Sheet1");

        let provider = WorkbookProvider::new(&wb);
        let ctx = offidized_formula::EvalContext::new(&provider, Some("Sheet1".to_string()), 1, 1);
        // INFO("recalc") should return "Automatic"
        let result = offidized_formula::evaluate("=INFO(\"recalc\")", &ctx);
        assert_eq!(result, offidized_formula::CalcValue::text("Automatic"));
    }

    #[test]
    fn test_info_function_release() {
        let mut wb = Workbook::new();
        wb.add_sheet("Sheet1");

        let provider = WorkbookProvider::new(&wb);
        let ctx = offidized_formula::EvalContext::new(&provider, Some("Sheet1".to_string()), 1, 1);
        // INFO("release") should return version string
        let result = offidized_formula::evaluate("=INFO(\"release\")", &ctx);
        assert_eq!(
            result,
            offidized_formula::CalcValue::text("offidized-0.1.0")
        );
    }

    #[test]
    fn test_info_function_system() {
        let mut wb = Workbook::new();
        wb.add_sheet("Sheet1");

        let provider = WorkbookProvider::new(&wb);
        let ctx = offidized_formula::EvalContext::new(&provider, Some("Sheet1".to_string()), 1, 1);
        // INFO("system") should return "mac" or "pcdos"
        let result = offidized_formula::evaluate("=INFO(\"system\")", &ctx);
        let text = result.as_scalar();
        match text {
            ScalarValue::Text(s) => {
                assert!(
                    s == "mac" || s == "pcdos",
                    "Expected 'mac' or 'pcdos', got '{}'",
                    s
                );
            }
            _ => panic!("Expected text value"),
        }
    }

    #[test]
    fn test_info_function_directory() {
        let mut wb = Workbook::new();
        wb.add_sheet("Sheet1");

        let provider = WorkbookProvider::new(&wb);
        let ctx = offidized_formula::EvalContext::new(&provider, Some("Sheet1".to_string()), 1, 1);
        // INFO("directory") should return #N/A
        let result = offidized_formula::evaluate("=INFO(\"directory\")", &ctx);
        assert_eq!(result, offidized_formula::CalcValue::error(XlError::Na));
    }
}
