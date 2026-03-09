use crate::cell::CellValue;
use crate::error::XlsxError;
use crate::range::CellRange;
use crate::workbook::Workbook;

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum LintSeverity {
    Error,
    Warning,
    Info,
}

#[derive(Debug, Clone, PartialEq, Eq, Default)]
pub struct LintLocation {
    pub sheet: Option<String>,
    pub cell: Option<String>,
    pub object: Option<String>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct LintFinding {
    pub severity: LintSeverity,
    pub code: String,
    pub message: String,
    pub location: LintLocation,
}

impl LintFinding {
    fn error(code: impl Into<String>, message: impl Into<String>, location: LintLocation) -> Self {
        Self {
            severity: LintSeverity::Error,
            code: code.into(),
            message: message.into(),
            location,
        }
    }

    fn warning(
        code: impl Into<String>,
        message: impl Into<String>,
        location: LintLocation,
    ) -> Self {
        Self {
            severity: LintSeverity::Warning,
            code: code.into(),
            message: message.into(),
            location,
        }
    }
}

#[derive(Debug, Clone, Default)]
pub struct LintReport {
    findings: Vec<LintFinding>,
}

impl LintReport {
    pub fn findings(&self) -> &[LintFinding] {
        self.findings.as_slice()
    }

    pub fn error_count(&self) -> usize {
        self.findings
            .iter()
            .filter(|finding| finding.severity == LintSeverity::Error)
            .count()
    }

    pub fn warning_count(&self) -> usize {
        self.findings
            .iter()
            .filter(|finding| finding.severity == LintSeverity::Warning)
            .count()
    }

    pub fn is_clean(&self) -> bool {
        self.findings.is_empty()
    }
}

pub struct WorkbookLintBuilder<'a> {
    workbook: &'a Workbook,
    broken_refs: bool,
    formula_consistency: bool,
    pivot_sources: bool,
    named_ranges: bool,
}

impl Workbook {
    pub fn lint(&self) -> WorkbookLintBuilder<'_> {
        WorkbookLintBuilder {
            workbook: self,
            broken_refs: false,
            formula_consistency: false,
            pivot_sources: false,
            named_ranges: false,
        }
    }
}

impl<'a> WorkbookLintBuilder<'a> {
    pub fn check_broken_refs(mut self) -> Self {
        self.broken_refs = true;
        self
    }

    pub fn check_formula_consistency(mut self) -> Self {
        self.formula_consistency = true;
        self
    }

    pub fn check_pivot_sources(mut self) -> Self {
        self.pivot_sources = true;
        self
    }

    pub fn check_named_ranges(mut self) -> Self {
        self.named_ranges = true;
        self
    }

    pub fn run(self) -> LintReport {
        let mut report = LintReport::default();

        if self.broken_refs {
            report
                .findings
                .extend(check_broken_formula_references(self.workbook));
        }
        if self.formula_consistency {
            report
                .findings
                .extend(check_formula_cached_values(self.workbook));
        }
        if self.pivot_sources {
            report
                .findings
                .extend(check_pivot_source_ranges(self.workbook));
        }
        if self.named_ranges {
            report.findings.extend(check_named_ranges(self.workbook));
        }

        report
    }
}

fn check_broken_formula_references(workbook: &Workbook) -> Vec<LintFinding> {
    let mut findings = Vec::new();
    for ws in workbook.worksheets() {
        for (cell_ref, cell) in ws.cells() {
            let Some(formula) = cell.formula() else {
                continue;
            };
            let tokens = match offidized_formula::lexer::tokenize(formula) {
                Ok(tokens) => tokens,
                Err(error) => {
                    findings.push(LintFinding::error(
                        "broken_formula_syntax",
                        format!("formula tokenize failed: {error}"),
                        LintLocation {
                            sheet: Some(ws.name().to_string()),
                            cell: Some(cell_ref.to_string()),
                            object: None,
                        },
                    ));
                    continue;
                }
            };

            for token in tokens {
                if let offidized_formula::token::Token::CellReference {
                    sheet: Some(sheet_name),
                    ..
                } = token
                {
                    if workbook.sheet(sheet_name.as_str()).is_none() {
                        findings.push(LintFinding::error(
                            "broken_sheet_ref",
                            format!("formula references missing sheet '{sheet_name}'"),
                            LintLocation {
                                sheet: Some(ws.name().to_string()),
                                cell: Some(cell_ref.to_string()),
                                object: None,
                            },
                        ));
                    }
                }
            }
        }
    }
    findings
}

fn check_formula_cached_values(workbook: &Workbook) -> Vec<LintFinding> {
    let mut findings = Vec::new();

    for ws in workbook.worksheets() {
        for (cell_ref, cell) in ws.cells() {
            let Some(formula) = cell.formula() else {
                continue;
            };
            let Some(cached) = cell.value() else {
                continue;
            };
            let (col, row) = match parse_cell_reference(cell_ref) {
                Ok(pos) => pos,
                Err(_) => continue,
            };
            let computed = workbook.evaluate_formula(formula, ws.name(), row, col);
            if !cell_values_equivalent(cached, &computed) {
                findings.push(LintFinding::warning(
                    "formula_cache_mismatch",
                    "cached formula value differs from computed value",
                    LintLocation {
                        sheet: Some(ws.name().to_string()),
                        cell: Some(cell_ref.to_string()),
                        object: None,
                    },
                ));
            }
        }
    }

    findings
}

fn check_pivot_source_ranges(workbook: &Workbook) -> Vec<LintFinding> {
    let mut findings = Vec::new();

    for ws in workbook.worksheets() {
        for pivot in ws.pivot_tables() {
            let source_text = pivot.source_reference().as_str();
            let Some((sheet_name, range)) = split_sheet_range(source_text) else {
                findings.push(LintFinding::error(
                    "pivot_source_invalid",
                    format!("pivot source '{source_text}' is not a worksheet range"),
                    LintLocation {
                        sheet: Some(ws.name().to_string()),
                        cell: None,
                        object: Some(pivot.name().to_string()),
                    },
                ));
                continue;
            };

            let Some(source_ws) = workbook.sheet(sheet_name.as_str()) else {
                findings.push(LintFinding::error(
                    "pivot_source_missing_sheet",
                    format!("pivot source sheet '{sheet_name}' not found"),
                    LintLocation {
                        sheet: Some(ws.name().to_string()),
                        cell: None,
                        object: Some(pivot.name().to_string()),
                    },
                ));
                continue;
            };

            let headers = match source_ws.source_headers_from_range(range.as_str()) {
                Ok(headers) => headers,
                Err(error) => {
                    findings.push(LintFinding::error(
                        "pivot_source_range_invalid",
                        format!("pivot source range parse failed: {error}"),
                        LintLocation {
                            sheet: Some(ws.name().to_string()),
                            cell: None,
                            object: Some(pivot.name().to_string()),
                        },
                    ));
                    continue;
                }
            };

            for field in pivot
                .row_fields()
                .iter()
                .chain(pivot.column_fields().iter())
                .chain(pivot.page_fields().iter())
                .map(|field| field.name())
            {
                if !headers.iter().any(|header| header == field) {
                    findings.push(LintFinding::error(
                        "pivot_field_missing_header",
                        format!("pivot field '{field}' not found in source headers"),
                        LintLocation {
                            sheet: Some(ws.name().to_string()),
                            cell: None,
                            object: Some(pivot.name().to_string()),
                        },
                    ));
                }
            }

            for field in pivot.data_fields().iter().map(|field| field.field_name()) {
                if !headers.iter().any(|header| header == field) {
                    findings.push(LintFinding::error(
                        "pivot_value_missing_header",
                        format!("pivot value field '{field}' not found in source headers"),
                        LintLocation {
                            sheet: Some(ws.name().to_string()),
                            cell: None,
                            object: Some(pivot.name().to_string()),
                        },
                    ));
                }
            }
        }
    }

    findings
}

fn check_named_ranges(workbook: &Workbook) -> Vec<LintFinding> {
    let mut findings = Vec::new();

    for defined_name in workbook.defined_names() {
        let reference = defined_name.reference();
        let Some((sheet_name, range)) = split_sheet_range(reference) else {
            findings.push(LintFinding::warning(
                "named_range_not_a1_range",
                format!("named range '{}' is not an A1 range", defined_name.name()),
                LintLocation {
                    sheet: None,
                    cell: None,
                    object: Some(defined_name.name().to_string()),
                },
            ));
            continue;
        };

        if workbook.sheet(sheet_name.as_str()).is_none() {
            findings.push(LintFinding::error(
                "named_range_missing_sheet",
                format!(
                    "named range '{}' references missing sheet '{}'",
                    defined_name.name(),
                    sheet_name
                ),
                LintLocation {
                    sheet: None,
                    cell: None,
                    object: Some(defined_name.name().to_string()),
                },
            ));
            continue;
        }

        if CellRange::parse(range.as_str()).is_err() {
            findings.push(LintFinding::error(
                "named_range_invalid_range",
                format!("named range '{}' has invalid range", defined_name.name()),
                LintLocation {
                    sheet: Some(sheet_name),
                    cell: None,
                    object: Some(defined_name.name().to_string()),
                },
            ));
        }
    }

    findings
}

fn split_sheet_range(reference: &str) -> Option<(String, String)> {
    let (sheet_raw, range) = reference.split_once('!')?;
    let sheet = if sheet_raw.starts_with('\'') && sheet_raw.ends_with('\'') && sheet_raw.len() > 1 {
        sheet_raw[1..sheet_raw.len() - 1].replace("''", "'")
    } else {
        sheet_raw.to_string()
    };
    Some((sheet, range.to_string()))
}

fn parse_cell_reference(reference: &str) -> Result<(u32, u32), XlsxError> {
    let normalized = crate::cell::normalize_cell_reference(reference)?;
    let split_index = normalized
        .char_indices()
        .find_map(|(index, ch)| ch.is_ascii_digit().then_some(index))
        .ok_or_else(|| XlsxError::InvalidCellReference(reference.to_string()))?;
    let (column_name, row_text) = normalized.split_at(split_index);
    let col = column_name
        .bytes()
        .try_fold(0_u32, |acc, byte| {
            acc.checked_mul(26)
                .and_then(|value| value.checked_add(u32::from(byte - b'A' + 1)))
        })
        .ok_or_else(|| XlsxError::InvalidCellReference(reference.to_string()))?;
    let row = row_text
        .parse::<u32>()
        .map_err(|_| XlsxError::InvalidCellReference(reference.to_string()))?;
    Ok((col, row))
}

fn cell_values_equivalent(left: &CellValue, right: &CellValue) -> bool {
    match (left, right) {
        (CellValue::Number(a), CellValue::Number(b)) => (a - b).abs() < 1e-9,
        (CellValue::String(a), CellValue::String(b)) => a == b,
        (CellValue::Bool(a), CellValue::Bool(b)) => a == b,
        (CellValue::Blank, CellValue::Blank) => true,
        (CellValue::Date(a), CellValue::Date(b)) => a == b,
        (CellValue::DateTime(a), CellValue::DateTime(b)) => (a - b).abs() < 1e-9,
        (CellValue::Error(a), CellValue::Error(b)) => a == b,
        _ => false,
    }
}

#[cfg(test)]
mod tests {
    use crate::{sum, Workbook};

    #[test]
    fn lint_finds_missing_sheet_in_formula() {
        let mut wb = Workbook::new();
        let ws = wb.add_sheet("Sheet1");
        ws.cell_mut("A1")
            .expect("cell")
            .set_formula("=Missing!A1")
            .set_value(1);

        let report = wb.lint().check_broken_refs().run();
        assert!(report
            .findings()
            .iter()
            .any(|finding| finding.code == "broken_sheet_ref"));
    }

    #[test]
    fn lint_finds_formula_cache_mismatch() {
        let mut wb = Workbook::new();
        let ws = wb.add_sheet("Sheet1");
        ws.cell_mut("A1").unwrap().set_value(2);
        ws.cell_mut("A2").unwrap().set_value(3);
        ws.cell_mut("A3")
            .unwrap()
            .set_formula("=SUM(A1:A2)")
            .set_value(100);

        let report = wb.lint().check_formula_consistency().run();
        assert!(report
            .findings()
            .iter()
            .any(|finding| finding.code == "formula_cache_mismatch"));
    }

    #[test]
    fn lint_finds_invalid_pivot_source_field() {
        let mut wb = Workbook::new();
        let data = wb.add_sheet("Data");
        data.cell_mut("A1").unwrap().set_value("Region");
        data.cell_mut("B1").unwrap().set_value("Revenue");

        wb.add_sheet("Pivot")
            .pivot("P")
            .source("Data!A1:B10")
            .rows(["Desk"])
            .values([sum("Revenue")])
            .place("A4")
            .expect("pivot write should succeed without pre-validate");

        let report = wb.lint().check_pivot_sources().run();
        assert!(report
            .findings()
            .iter()
            .any(|finding| finding.code == "pivot_field_missing_header"));
    }

    #[test]
    fn lint_finds_named_range_with_missing_sheet() {
        let mut wb = Workbook::new();
        wb.add_defined_name("Positions", "Missing!A1:A10");

        let report = wb.lint().check_named_ranges().run();
        assert!(report
            .findings()
            .iter()
            .any(|finding| finding.code == "named_range_missing_sheet"));
    }
}
