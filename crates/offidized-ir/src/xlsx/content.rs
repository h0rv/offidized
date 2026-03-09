//! xlsx content mode: derive and apply cell values/formulas.
//!
//! Content mode emits one line per cell in row-major order. Each line is
//! `CELLREF: VALUE` where VALUE follows the encoding rules documented in
//! the IR specification.
//!
//! Derive intentionally does not emit chart sections. Chart sections are
//! accepted by apply to create new charts.

use crate::{ApplyResult, DeriveOptions, IrError, Result};
use offidized_xlsx::{
    BarDirection, CellValue, Chart, ChartAxis, ChartDataRef, ChartGrouping, ChartLegend,
    ChartSeries, ChartType, Workbook, Worksheet,
};

// ---------------------------------------------------------------------------
// Derive
// ---------------------------------------------------------------------------

/// Append the xlsx content-mode body to `output`.
pub(crate) fn derive_content(wb: &Workbook, options: &DeriveOptions, output: &mut String) {
    for ws in wb.worksheets() {
        // Filter by sheet name if specified
        if let Some(ref sheet_filter) = options.sheet {
            if ws.name() != sheet_filter.as_str() {
                continue;
            }
        }

        output.push('\n');
        output.push_str("=== Sheet: ");
        output.push_str(ws.name());
        output.push_str(" ===\n");

        // Annotate hidden sheets
        match ws.visibility() {
            offidized_xlsx::SheetVisibility::Hidden => {
                output.push_str("# hidden\n");
            }
            offidized_xlsx::SheetVisibility::VeryHidden => {
                output.push_str("# very-hidden\n");
            }
            offidized_xlsx::SheetVisibility::Visible => {}
        }

        derive_sheet_cells(ws, output);
    }
}

/// Derive all cells from a worksheet in row-major order.
fn derive_sheet_cells(ws: &Worksheet, output: &mut String) {
    // Collect cells and sort by (row, col) for row-major order.
    // BTreeMap iterates lexicographically ("A1" < "A10" < "A2"), not row-major.
    let mut cells: Vec<(&str, &offidized_xlsx::Cell)> = ws.cells().collect();
    cells.sort_by_key(|(cell_ref, _)| cell_sort_key(cell_ref));

    for (cell_ref, cell) in cells {
        if let Some(formatted) = format_cell(cell) {
            output.push_str(cell_ref);
            output.push_str(": ");
            output.push_str(&formatted);
            output.push('\n');
        }
    }
}

/// Parse a chart type from an IR string.
fn parse_ir_chart_type(s: &str) -> Option<ChartType> {
    match s.to_lowercase().as_str() {
        "bar" | "column" => Some(ChartType::Bar),
        "line" => Some(ChartType::Line),
        "pie" => Some(ChartType::Pie),
        "area" => Some(ChartType::Area),
        "scatter" => Some(ChartType::Scatter),
        "doughnut" => Some(ChartType::Doughnut),
        "radar" => Some(ChartType::Radar),
        "bubble" => Some(ChartType::Bubble),
        "stock" => Some(ChartType::Stock),
        "surface" => Some(ChartType::Surface),
        "combo" => Some(ChartType::Combo),
        _ => None,
    }
}

/// Parse an anchor string `"from_col,from_row,to_col,to_row"` (zero-based).
fn parse_ir_anchor(s: &str) -> Option<(u32, u32, u32, u32)> {
    let parts: Vec<&str> = s.split(',').collect();
    if parts.len() != 4 {
        return None;
    }
    let nums: Vec<u32> = parts
        .iter()
        .filter_map(|p| p.trim().parse::<u32>().ok())
        .collect();
    if nums.len() != 4 {
        return None;
    }
    Some((nums[0], nums[1], nums[2], nums[3]))
}

/// Parse a series spec: `"name | [cats] | vals"` or `"name | vals"`.
fn parse_ir_series(s: &str) -> Option<(String, Option<String>, String)> {
    let parts: Vec<&str> = s.splitn(3, " | ").collect();
    match parts.as_slice() {
        [name, cats, vals] => {
            let cats_opt = if cats.trim().is_empty() {
                None
            } else {
                Some(cats.trim().to_string())
            };
            Some((name.trim().to_string(), cats_opt, vals.trim().to_string()))
        }
        [name, vals] => Some((name.trim().to_string(), None, vals.trim().to_string())),
        _ => None,
    }
}

/// Builder accumulating properties for a single chart section.
struct ChartBuilder {
    name: String,
    sheet: Option<String>,
    chart_type: Option<ChartType>,
    from_col: u32,
    from_row: u32,
    to_col: u32,
    to_row: u32,
    title: Option<String>,
    bar_direction: Option<BarDirection>,
    grouping: Option<ChartGrouping>,
    legend_pos: Option<String>,
    /// (name, cats_formula, vals_formula)
    series: Vec<(String, Option<String>, String)>,
}

impl ChartBuilder {
    fn new(name: String) -> Self {
        Self {
            name,
            sheet: None,
            chart_type: None,
            from_col: 0,
            from_row: 0,
            to_col: 9,
            to_row: 14,
            title: None,
            bar_direction: None,
            grouping: None,
            legend_pos: None,
            series: Vec::new(),
        }
    }

    /// Apply a `key: value` property line to this builder.
    fn apply_property(&mut self, line: &str) {
        let Some((key, value)) = line.split_once(": ") else {
            return;
        };
        let key = key.trim();
        let value = value.trim();
        match key {
            "sheet" => self.sheet = Some(value.to_string()),
            "type" => self.chart_type = parse_ir_chart_type(value),
            "anchor" => {
                if let Some((fc, fr, tc, tr)) = parse_ir_anchor(value) {
                    self.from_col = fc;
                    self.from_row = fr;
                    self.to_col = tc;
                    self.to_row = tr;
                }
            }
            "title" => self.title = Some(value.to_string()),
            "bar-direction" => {
                self.bar_direction = match value {
                    "col" | "column" => Some(BarDirection::Column),
                    "bar" => Some(BarDirection::Bar),
                    _ => None,
                };
            }
            "grouping" => self.grouping = ChartGrouping::from_xml_value(value),
            "legend" => self.legend_pos = Some(value.to_string()),
            "series" => {
                if let Some(spec) = parse_ir_series(value) {
                    self.series.push(spec);
                }
            }
            _ => {}
        }
    }

    /// Build the chart and add it to `wb`. Returns `true` on success.
    fn finalize(self, wb: &mut Workbook) -> bool {
        let Some(sheet_name) = self.sheet else {
            return false;
        };
        let Some(chart_type) = self.chart_type else {
            return false;
        };

        if !wb.contains_sheet(&sheet_name) {
            wb.add_sheet(&sheet_name);
        }
        let Some(ws) = wb.sheet_mut(&sheet_name) else {
            return false;
        };

        let mut chart = Chart::new(chart_type);
        chart.set_anchor(self.from_col, self.from_row, self.to_col, self.to_row);
        if !self.name.is_empty() {
            chart.set_name(&self.name);
        }
        if let Some(title) = self.title {
            chart.set_title(title);
        }
        if let Some(dir) = self.bar_direction {
            chart.set_bar_direction(dir);
        }
        if let Some(grp) = self.grouping {
            chart.set_grouping(grp);
        }
        if let Some(pos) = self.legend_pos {
            let mut legend = ChartLegend::new();
            legend.set_position(pos);
            chart.set_legend(legend);
        }

        for (i, (name, cats, vals)) in self.series.into_iter().enumerate() {
            let idx = i as u32;
            let mut series = ChartSeries::new(idx, idx);
            if !name.is_empty() {
                series.set_name(name);
            }
            if let Some(cats_formula) = cats {
                series.set_categories(ChartDataRef::from_formula(cats_formula));
            }
            if !vals.is_empty() {
                series.set_values(ChartDataRef::from_formula(vals));
            }
            chart.add_series(series);
        }

        chart.add_axis(ChartAxis::new_category());
        chart.add_axis(ChartAxis::new_value());

        ws.add_chart(chart);
        true
    }
}

/// Format a cell for the IR. Returns `None` for blank/empty cells.
fn format_cell(cell: &offidized_xlsx::Cell) -> Option<String> {
    // Formula takes precedence over value
    if let Some(formula) = cell.formula() {
        // Escape newlines to preserve one-line-per-cell invariant
        let escaped = formula.replace('\n', "\\n").replace('\r', "\\r");
        return Some(format!("={escaped}"));
    }

    match cell.value()? {
        CellValue::Blank => None,
        CellValue::Number(n) => Some(format_number(*n)),
        CellValue::Bool(b) => Some(if *b {
            "true".to_string()
        } else {
            "false".to_string()
        }),
        CellValue::Error(e) => Some(e.clone()),
        CellValue::String(s) => Some(format_string_value(s)),
        CellValue::Date(d) => Some(format_string_value(d)),
        CellValue::DateTime(serial) => Some(format_number(*serial)),
        CellValue::RichText(runs) => {
            let text: String = runs.iter().map(|r| r.text()).collect::<String>();
            Some(format_string_value(&text))
        }
    }
}

/// Format a number, using integer representation when possible.
fn format_number(n: f64) -> String {
    // Handle zero (including negative zero)
    if n == 0.0 {
        return "0".to_string();
    }
    // Whole numbers within safe integer range: emit without decimal point
    if n.fract() == 0.0 && n.is_finite() && n.abs() < 1e15 {
        #[allow(clippy::cast_possible_truncation)]
        return format!("{}", n as i64);
    }
    format!("{n}")
}

/// Format a string value, quoting it if it could be misinterpreted during parsing.
fn format_string_value(s: &str) -> String {
    if needs_quoting(s) {
        let escaped = s
            .replace('"', "\"\"")
            .replace('\n', "\\n")
            .replace('\r', "\\r");
        format!("\"{escaped}\"")
    } else {
        s.to_string()
    }
}

/// Returns true if a string value needs quoting to avoid ambiguity.
fn needs_quoting(s: &str) -> bool {
    // Empty string must be quoted to distinguish from blank
    if s.is_empty() {
        return true;
    }
    // Looks like a number
    if s.parse::<f64>().is_ok() {
        return true;
    }
    // Looks like a boolean
    if s.eq_ignore_ascii_case("true") || s.eq_ignore_ascii_case("false") {
        return true;
    }
    // Looks like a formula
    if s.starts_with('=') {
        return true;
    }
    // Starts with # (could be confused with error values or comment lines)
    if s.starts_with('#') {
        return true;
    }
    // Is the empty/delete marker
    if s == "<empty>" {
        return true;
    }
    // Has leading/trailing whitespace
    if s != s.trim() {
        return true;
    }
    // Contains newlines or carriage returns
    if s.contains('\n') || s.contains('\r') {
        return true;
    }
    // Contains quotes (would be ambiguous with quoted strings)
    if s.contains('"') {
        return true;
    }
    false
}

/// Parse (row, col) from a cell reference for sorting in row-major order.
fn cell_sort_key(cell_ref: &str) -> (u32, u32) {
    let col_end = cell_ref
        .find(|c: char| c.is_ascii_digit())
        .unwrap_or(cell_ref.len());
    let col_str = &cell_ref[..col_end];
    let row_str = &cell_ref[col_end..];

    let col = col_str.bytes().fold(0u32, |acc, b| {
        acc * 26 + u32::from(b.to_ascii_uppercase() - b'A') + 1
    });
    let row: u32 = row_str.parse().unwrap_or(0);
    (row, col)
}

// ---------------------------------------------------------------------------
// Apply
// ---------------------------------------------------------------------------

/// Parsed cell value from the IR text.
enum ParsedValue {
    Number(f64),
    Bool(bool),
    String(String),
    Formula(String),
    Error(String),
    Empty,
}

/// Apply the IR body to a workbook, updating cell values and adding charts.
pub(crate) fn apply_content(body: &str, wb: &mut Workbook) -> Result<ApplyResult> {
    let mut result = ApplyResult::default();
    let mut current_sheet: Option<String> = None;
    let mut current_chart: Option<ChartBuilder> = None;
    let mut in_chart_section = false;

    for line in body.lines() {
        let line = line.trim_end_matches('\r');

        // Chart header: === Chart: NAME ===
        if let Some(name) = parse_chart_header(line) {
            // Finalize any pending chart builder
            if let Some(cb) = current_chart.take() {
                if cb.finalize(wb) {
                    result.charts_added += 1;
                }
            }
            current_chart = Some(ChartBuilder::new(name.to_string()));
            in_chart_section = true;
            continue;
        }

        // Sheet header: === Sheet: NAME ===
        if let Some(name) = parse_sheet_header(line) {
            // Finalize any pending chart builder
            if let Some(cb) = current_chart.take() {
                if cb.finalize(wb) {
                    result.charts_added += 1;
                }
            }
            current_sheet = Some(name.to_string());
            in_chart_section = false;
            continue;
        }

        // Skip comments and blank lines
        if line.is_empty() || line.starts_with('#') {
            continue;
        }

        // Inside a chart section: accumulate properties
        if in_chart_section {
            if let Some(ref mut cb) = current_chart {
                cb.apply_property(line);
            }
            continue;
        }

        // Cell line: REF: VALUE
        let Some(sheet_name) = current_sheet.as_deref() else {
            continue;
        };

        let Some((cell_ref, raw_value)) = parse_cell_line(line) else {
            continue;
        };

        let parsed = parse_value(raw_value);

        // Auto-create the sheet if it doesn't exist yet.
        if !wb.contains_sheet(sheet_name) {
            wb.add_sheet(sheet_name);
        }

        let Some(ws) = wb.sheet_mut(sheet_name) else {
            continue;
        };

        // Check if cell already exists
        let cell_exists = ws.cell(cell_ref).is_some();

        let cell = ws
            .cell_mut(cell_ref)
            .map_err(|e| IrError::InvalidBody(format!("invalid cell reference {cell_ref}: {e}")))?;

        match parsed {
            ParsedValue::Number(n) => {
                cell.clear_formula().set_value(CellValue::Number(n));
            }
            ParsedValue::Bool(b) => {
                cell.clear_formula().set_value(CellValue::Bool(b));
            }
            ParsedValue::String(s) => {
                cell.clear_formula().set_value(CellValue::String(s));
            }
            ParsedValue::Formula(f) => {
                cell.set_formula(f);
            }
            ParsedValue::Error(e) => {
                cell.clear_formula().set_value(CellValue::Error(e));
            }
            ParsedValue::Empty => {
                cell.clear_value().clear_formula();
                result.cells_cleared += 1;
                continue; // Don't count as updated/created
            }
        }

        if cell_exists {
            result.cells_updated += 1;
        } else {
            result.cells_created += 1;
        }
    }

    // Finalize any chart still pending at end of input
    if let Some(cb) = current_chart.take() {
        if cb.finalize(wb) {
            result.charts_added += 1;
        }
    }

    Ok(result)
}

/// Parse a sheet header line like `=== Sheet: Revenue ===`.
fn parse_sheet_header(line: &str) -> Option<&str> {
    let line = line.trim();
    let rest = line.strip_prefix("=== Sheet: ")?;
    let name = rest.strip_suffix(" ===")?;
    if name.is_empty() {
        return None;
    }
    Some(name)
}

/// Parse a chart header line like `=== Chart: Revenue Trend ===`.
fn parse_chart_header(line: &str) -> Option<&str> {
    let line = line.trim();
    let rest = line.strip_prefix("=== Chart: ")?;
    let name = rest.strip_suffix(" ===")?;
    if name.is_empty() {
        return None;
    }
    Some(name)
}

/// Parse a cell line like `A1: hello world`.
/// Returns `(cell_ref, raw_value)`.
fn parse_cell_line(line: &str) -> Option<(&str, &str)> {
    let (cell_ref, value) = line.split_once(": ")?;
    let cell_ref = cell_ref.trim();

    // Validate that it looks like a cell reference (letters followed by digits)
    if cell_ref.is_empty() {
        return None;
    }
    let first = cell_ref.as_bytes().first()?;
    if !first.is_ascii_alphabetic() {
        return None;
    }

    Some((cell_ref, value))
}

/// Parse a raw value string into a typed `ParsedValue`.
fn parse_value(s: &str) -> ParsedValue {
    // Empty/delete marker
    if s == "<empty>" {
        return ParsedValue::Empty;
    }

    // Formula (starts with =)
    if let Some(formula) = s.strip_prefix('=') {
        // Unescape \n back to real newlines (formulas can contain newlines
        // in column references like Table1[Purchase\n price])
        let unescaped = formula.replace("\\n", "\n").replace("\\r", "\r");
        return ParsedValue::Formula(unescaped);
    }

    // Quoted string
    if s.starts_with('"') && s.ends_with('"') && s.len() >= 2 {
        let inner = &s[1..s.len() - 1];
        let unescaped = inner
            .replace("\"\"", "\"")
            .replace("\\n", "\n")
            .replace("\\r", "\r");
        return ParsedValue::String(unescaped);
    }

    // Boolean
    if s == "true" {
        return ParsedValue::Bool(true);
    }
    if s == "false" {
        return ParsedValue::Bool(false);
    }

    // Error value (like #REF!, #NAME?, #VALUE!, #DIV/0!, #N/A, #NULL!)
    if s.starts_with('#') {
        return ParsedValue::Error(s.to_string());
    }

    // Number
    if let Ok(n) = s.parse::<f64>() {
        return ParsedValue::Number(n);
    }

    // Everything else is a string
    ParsedValue::String(s.to_string())
}

// ---------------------------------------------------------------------------
// Tests
// ---------------------------------------------------------------------------

#[cfg(test)]
mod tests {
    #![allow(clippy::panic_in_result_fn, clippy::approx_constant)]

    use super::*;
    use offidized_xlsx::Workbook;

    type TestResult = std::result::Result<(), Box<dyn std::error::Error>>;

    // --- Value encoding/decoding roundtrip tests ---

    #[test]
    fn format_number_integers() {
        assert_eq!(format_number(42000.0), "42000");
        assert_eq!(format_number(-5.0), "-5");
        assert_eq!(format_number(0.0), "0");
        assert_eq!(format_number(1.0), "1");
    }

    #[test]
    fn format_number_decimals() {
        assert_eq!(format_number(3.14), "3.14");
        assert_eq!(format_number(0.1), "0.1");
        assert_eq!(format_number(-0.5), "-0.5");
    }

    #[test]
    fn string_quoting_numbers() {
        // String that looks like a number must be quoted
        assert_eq!(format_string_value("42"), "\"42\"");
        assert_eq!(format_string_value("3.14"), "\"3.14\"");
        assert_eq!(format_string_value("-5"), "\"-5\"");
    }

    #[test]
    fn string_quoting_booleans() {
        assert_eq!(format_string_value("true"), "\"true\"");
        assert_eq!(format_string_value("false"), "\"false\"");
        assert_eq!(format_string_value("TRUE"), "\"TRUE\"");
    }

    #[test]
    fn string_quoting_formulas() {
        assert_eq!(format_string_value("=SUM(A1:A3)"), "\"=SUM(A1:A3)\"");
    }

    #[test]
    fn string_quoting_errors() {
        assert_eq!(format_string_value("#REF!"), "\"#REF!\"");
        assert_eq!(format_string_value("#NAME?"), "\"#NAME?\"");
    }

    #[test]
    fn string_quoting_empty_marker() {
        assert_eq!(format_string_value("<empty>"), "\"<empty>\"");
    }

    #[test]
    fn string_quoting_whitespace() {
        assert_eq!(format_string_value("  indented"), "\"  indented\"");
        assert_eq!(format_string_value("trailing "), "\"trailing \"");
    }

    #[test]
    fn string_quoting_newlines() {
        assert_eq!(format_string_value("line1\nline2"), "\"line1\\nline2\"");
    }

    #[test]
    fn string_no_quoting_plain() {
        assert_eq!(format_string_value("Category"), "Category");
        assert_eq!(format_string_value("Hello World"), "Hello World");
        assert_eq!(format_string_value("Revenue"), "Revenue");
    }

    #[test]
    fn string_quoting_with_quotes() {
        assert_eq!(
            format_string_value("has \"quotes\""),
            "\"has \"\"quotes\"\"\"",
        );
    }

    // --- Parse value tests ---

    #[test]
    fn parse_value_empty() {
        assert!(matches!(parse_value("<empty>"), ParsedValue::Empty));
    }

    #[test]
    fn parse_value_formula() {
        match parse_value("=SUM(B3:B5)") {
            ParsedValue::Formula(f) => assert_eq!(f, "SUM(B3:B5)"),
            _ => panic!("expected formula"),
        }
    }

    #[test]
    fn parse_value_quoted_string() {
        match parse_value("\"42\"") {
            ParsedValue::String(s) => assert_eq!(s, "42"),
            _ => panic!("expected string"),
        }
        match parse_value("\"true\"") {
            ParsedValue::String(s) => assert_eq!(s, "true"),
            _ => panic!("expected string"),
        }
        match parse_value("\"has \"\"quotes\"\"\"") {
            ParsedValue::String(s) => assert_eq!(s, "has \"quotes\""),
            _ => panic!("expected string"),
        }
        match parse_value("\"line1\\nline2\"") {
            ParsedValue::String(s) => assert_eq!(s, "line1\nline2"),
            _ => panic!("expected string"),
        }
    }

    #[test]
    fn parse_value_bool() {
        assert!(matches!(parse_value("true"), ParsedValue::Bool(true)));
        assert!(matches!(parse_value("false"), ParsedValue::Bool(false)));
    }

    #[test]
    fn parse_value_error() {
        match parse_value("#REF!") {
            ParsedValue::Error(e) => assert_eq!(e, "#REF!"),
            _ => panic!("expected error"),
        }
    }

    #[test]
    fn parse_value_number() {
        match parse_value("42000") {
            ParsedValue::Number(n) => assert!((n - 42000.0).abs() < f64::EPSILON),
            _ => panic!("expected number"),
        }
        match parse_value("3.14") {
            ParsedValue::Number(n) => assert!((n - 3.14).abs() < f64::EPSILON),
            _ => panic!("expected number"),
        }
        match parse_value("-5") {
            ParsedValue::Number(n) => assert!((n - (-5.0)).abs() < f64::EPSILON),
            _ => panic!("expected number"),
        }
    }

    #[test]
    fn parse_value_plain_string() {
        match parse_value("Category") {
            ParsedValue::String(s) => assert_eq!(s, "Category"),
            _ => panic!("expected string"),
        }
    }

    // --- Derive/apply roundtrip ---

    #[test]
    fn parse_sheet_header_valid() {
        assert_eq!(
            parse_sheet_header("=== Sheet: Revenue ==="),
            Some("Revenue")
        );
        assert_eq!(
            parse_sheet_header("=== Sheet: My Sheet ==="),
            Some("My Sheet"),
        );
    }

    #[test]
    fn parse_sheet_header_invalid() {
        assert_eq!(parse_sheet_header("Not a header"), None);
        assert_eq!(parse_sheet_header("=== Sheet:  ==="), None);
    }

    #[test]
    fn parse_chart_header_valid() {
        assert_eq!(
            parse_chart_header("=== Chart: Revenue ==="),
            Some("Revenue")
        );
        assert_eq!(
            parse_chart_header("=== Chart: My Chart 1 ==="),
            Some("My Chart 1"),
        );
    }

    #[test]
    fn parse_chart_header_invalid() {
        assert_eq!(parse_chart_header("Not a header"), None);
        assert_eq!(parse_chart_header("=== Chart:  ==="), None);
    }

    #[test]
    fn parse_cell_line_valid() {
        assert_eq!(parse_cell_line("A1: hello"), Some(("A1", "hello")));
        assert_eq!(
            parse_cell_line("B2: =SUM(A1:A3)"),
            Some(("B2", "=SUM(A1:A3)"))
        );
        assert_eq!(
            parse_cell_line("AA100: some value"),
            Some(("AA100", "some value")),
        );
    }

    #[test]
    fn parse_cell_line_value_with_colon() {
        // Value containing ": " should work (split on first occurrence)
        assert_eq!(
            parse_cell_line("A1: time: 12:30"),
            Some(("A1", "time: 12:30")),
        );
    }

    #[test]
    fn cell_sort_key_ordering() {
        // Row-major: A1, B1, C1, A2, B2, C2
        let mut refs = vec!["C2", "A1", "B2", "A2", "C1", "B1"];
        refs.sort_by_key(|r| cell_sort_key(r));
        assert_eq!(refs, vec!["A1", "B1", "C1", "A2", "B2", "C2"]);
    }

    #[test]
    fn derive_basic_workbook() -> TestResult {
        let mut wb = Workbook::new();
        let ws = wb.add_sheet("Sales");
        ws.cell_mut("A1")?.set_value("Product");
        ws.cell_mut("B1")?.set_value(42000);
        ws.cell_mut("A2")?.set_value(true);
        ws.cell_mut("B2")?.set_formula("SUM(B1:B1)");

        let mut output = String::new();
        derive_content(&wb, &DeriveOptions::default(), &mut output);

        assert!(output.contains("=== Sheet: Sales ==="));
        assert!(output.contains("A1: Product"));
        assert!(output.contains("B1: 42000"));
        assert!(output.contains("A2: true"));
        assert!(output.contains("B2: =SUM(B1:B1)"));

        Ok(())
    }

    #[test]
    fn derive_ignores_existing_charts() -> TestResult {
        use offidized_xlsx::{Chart, ChartAxis, ChartDataRef, ChartSeries, ChartType};

        let mut wb = Workbook::new();
        let ws = wb.add_sheet("Sheet1");
        ws.cell_mut("A1")?.set_value("Category");
        ws.cell_mut("B1")?.set_value("Value");
        ws.cell_mut("A2")?.set_value("Q1");
        ws.cell_mut("B2")?.set_value(100);
        ws.cell_mut("A3")?.set_value("Q2");
        ws.cell_mut("B3")?.set_value(200);

        let mut chart = Chart::new(ChartType::Bar);
        chart.set_anchor(1, 4, 8, 20);
        chart.set_name("Revenue");
        let mut series = ChartSeries::new(0, 0);
        series.set_name("Revenue");
        series.set_categories(ChartDataRef::from_formula("'Sheet1'!$A$2:$A$3"));
        series.set_values(ChartDataRef::from_formula("'Sheet1'!$B$2:$B$3"));
        chart.add_series(series);
        chart.add_axis(ChartAxis::new_category());
        chart.add_axis(ChartAxis::new_value());
        ws.add_chart(chart);

        let mut output = String::new();
        derive_content(&wb, &DeriveOptions::default(), &mut output);

        assert!(output.contains("=== Sheet: Sheet1 ==="));
        assert!(!output.contains("=== Chart: "));

        Ok(())
    }

    #[test]
    fn derive_string_that_looks_like_number() -> TestResult {
        let mut wb = Workbook::new();
        let ws = wb.add_sheet("Sheet1");
        ws.cell_mut("A1")?.set_value(CellValue::String("42".into()));
        ws.cell_mut("A2")?
            .set_value(CellValue::String("true".into()));
        ws.cell_mut("A3")?
            .set_value(CellValue::String("=SUM".into()));
        ws.cell_mut("A4")?
            .set_value(CellValue::String("#REF!".into()));
        ws.cell_mut("A5")?
            .set_value(CellValue::String("<empty>".into()));

        let mut output = String::new();
        derive_content(&wb, &DeriveOptions::default(), &mut output);

        assert!(output.contains("A1: \"42\""));
        assert!(output.contains("A2: \"true\""));
        assert!(output.contains("A3: \"=SUM\""));
        assert!(output.contains("A4: \"#REF!\""));
        assert!(output.contains("A5: \"<empty>\""));

        Ok(())
    }

    #[test]
    fn derive_rich_text_flattened() -> TestResult {
        use offidized_xlsx::RichTextRun;
        let mut wb = Workbook::new();
        let ws = wb.add_sheet("Sheet1");
        ws.cell_mut("A1")?.set_value(CellValue::RichText(vec![
            RichTextRun::new("Hello "),
            RichTextRun::new("World"),
        ]));

        let mut output = String::new();
        derive_content(&wb, &DeriveOptions::default(), &mut output);

        assert!(output.contains("A1: Hello World"));

        Ok(())
    }

    #[test]
    fn derive_hidden_sheet() -> TestResult {
        let mut wb = Workbook::new();
        let ws = wb.add_sheet("Hidden");
        ws.set_visibility(offidized_xlsx::SheetVisibility::Hidden);
        ws.cell_mut("A1")?.set_value("secret");

        let mut output = String::new();
        derive_content(&wb, &DeriveOptions::default(), &mut output);

        assert!(output.contains("=== Sheet: Hidden ==="));
        assert!(output.contains("# hidden"));
        assert!(output.contains("A1: secret"));

        Ok(())
    }

    #[test]
    fn derive_sheet_filter() -> TestResult {
        let mut wb = Workbook::new();
        wb.add_sheet("Sheet1").cell_mut("A1")?.set_value("one");
        wb.add_sheet("Sheet2").cell_mut("A1")?.set_value("two");

        let options = DeriveOptions {
            sheet: Some("Sheet2".to_string()),
            ..DeriveOptions::default()
        };
        let mut output = String::new();
        derive_content(&wb, &options, &mut output);

        assert!(!output.contains("Sheet1"));
        assert!(output.contains("=== Sheet: Sheet2 ==="));
        assert!(output.contains("A1: two"));

        Ok(())
    }

    #[test]
    fn apply_basic() -> TestResult {
        let mut wb = Workbook::new();
        let ws = wb.add_sheet("Sheet1");
        ws.cell_mut("A1")?.set_value("original");
        ws.cell_mut("B1")?.set_value(100);

        let body = "\n=== Sheet: Sheet1 ===\nA1: updated\nB1: 200\n";
        let result = apply_content(body, &mut wb)?;

        assert_eq!(result.cells_updated, 2);
        assert_eq!(result.cells_created, 0);

        let ws = wb.sheet("Sheet1").ok_or("sheet missing")?;
        assert_eq!(
            ws.cell("A1").and_then(|c| c.value()),
            Some(&CellValue::String("updated".into())),
        );
        assert_eq!(
            ws.cell("B1").and_then(|c| c.value()),
            Some(&CellValue::Number(200.0)),
        );

        Ok(())
    }

    #[test]
    fn apply_creates_new_cell() -> TestResult {
        let mut wb = Workbook::new();
        wb.add_sheet("Sheet1");

        let body = "\n=== Sheet: Sheet1 ===\nA1: new value\n";
        let result = apply_content(body, &mut wb)?;

        assert_eq!(result.cells_created, 1);
        assert_eq!(result.cells_updated, 0);

        let ws = wb.sheet("Sheet1").ok_or("sheet missing")?;
        assert_eq!(
            ws.cell("A1").and_then(|c| c.value()),
            Some(&CellValue::String("new value".into())),
        );

        Ok(())
    }

    #[test]
    fn apply_clear_cell() -> TestResult {
        let mut wb = Workbook::new();
        let ws = wb.add_sheet("Sheet1");
        ws.cell_mut("A1")?.set_value("will be cleared");

        let body = "\n=== Sheet: Sheet1 ===\nA1: <empty>\n";
        let result = apply_content(body, &mut wb)?;

        assert_eq!(result.cells_cleared, 1);

        let ws = wb.sheet("Sheet1").ok_or("sheet missing")?;
        let cell = ws.cell("A1").ok_or("cell missing")?;
        assert_eq!(cell.value(), None);
        assert_eq!(cell.formula(), None);

        Ok(())
    }

    #[test]
    fn apply_formula() -> TestResult {
        let mut wb = Workbook::new();
        let ws = wb.add_sheet("Sheet1");
        ws.cell_mut("A1")?.set_value(10);

        let body = "\n=== Sheet: Sheet1 ===\nB1: =SUM(A1:A1)\n";
        let result = apply_content(body, &mut wb)?;

        assert_eq!(result.cells_created, 1);

        let ws = wb.sheet("Sheet1").ok_or("sheet missing")?;
        let cell = ws.cell("B1").ok_or("cell missing")?;
        assert_eq!(cell.formula(), Some("SUM(A1:A1)"));

        Ok(())
    }

    #[test]
    fn apply_preserves_untouched_cells() -> TestResult {
        let mut wb = Workbook::new();
        let ws = wb.add_sheet("Sheet1");
        ws.cell_mut("A1")?.set_value("keep me");
        ws.cell_mut("B1")?.set_value(42);
        ws.cell_mut("A2")?.set_value("also keep");

        // Only update B1
        let body = "\n=== Sheet: Sheet1 ===\nB1: 99\n";
        apply_content(body, &mut wb)?;

        let ws = wb.sheet("Sheet1").ok_or("sheet missing")?;
        assert_eq!(
            ws.cell("A1").and_then(|c| c.value()),
            Some(&CellValue::String("keep me".into())),
        );
        assert_eq!(
            ws.cell("B1").and_then(|c| c.value()),
            Some(&CellValue::Number(99.0)),
        );
        assert_eq!(
            ws.cell("A2").and_then(|c| c.value()),
            Some(&CellValue::String("also keep".into())),
        );

        Ok(())
    }

    #[test]
    fn apply_creates_missing_sheet() -> TestResult {
        let mut wb = Workbook::new();
        wb.add_sheet("Sheet1");

        let body = "\n=== Sheet: NewSheet ===\nA1: value\n";
        let result = apply_content(body, &mut wb)?;

        // Sheet should be created automatically, cell written, no warnings.
        assert_eq!(result.cells_created, 1);
        assert_eq!(result.cells_updated, 0);
        assert!(result.warnings.is_empty());

        let ws = wb.sheet("NewSheet").ok_or("sheet was not created")?;
        assert_eq!(
            ws.cell("A1").and_then(|c| c.value()),
            Some(&CellValue::String("value".into())),
        );

        Ok(())
    }

    #[test]
    fn apply_chart_section_creates_chart() -> TestResult {
        let mut wb = Workbook::new();
        wb.add_sheet("Sheet1");

        let body = r#"
=== Chart: Revenue ===
sheet: Sheet1
type: bar
anchor: 1,4,8,20
title: Revenue by Quarter
bar-direction: col
grouping: clustered
legend: r
series: Revenue | 'Sheet1'!$A$2:$A$3 | 'Sheet1'!$B$2:$B$3
"#;
        let result = apply_content(body, &mut wb)?;

        assert_eq!(result.charts_added, 1);
        assert_eq!(result.cells_created, 0);
        assert_eq!(result.cells_updated, 0);

        let ws = wb.sheet("Sheet1").ok_or("sheet missing")?;
        assert_eq!(ws.charts().len(), 1);
        assert_eq!(ws.charts()[0].name(), Some("Revenue"));
        assert_eq!(ws.charts()[0].title(), Some("Revenue by Quarter"));
        assert_eq!(ws.charts()[0].from_col(), 1);
        assert_eq!(ws.charts()[0].from_row(), 4);
        assert_eq!(ws.charts()[0].to_col(), 8);
        assert_eq!(ws.charts()[0].to_row(), 20);

        Ok(())
    }

    #[test]
    fn value_encoding_roundtrip() -> TestResult {
        // Test that special string values survive derive→parse roundtrip
        let test_cases = vec![
            (CellValue::String("true".into()), "\"true\""),
            (CellValue::String("42".into()), "\"42\""),
            (CellValue::String("=SUM".into()), "\"=SUM\""),
            (CellValue::String("#REF!".into()), "\"#REF!\""),
            (CellValue::String("<empty>".into()), "\"<empty>\""),
            (CellValue::Number(42.0), "42"),
            (CellValue::Bool(true), "true"),
            (CellValue::Bool(false), "false"),
            (CellValue::Error("#DIV/0!".into()), "#DIV/0!"),
        ];

        for (cell_value, expected_ir) in &test_cases {
            let mut wb = Workbook::new();
            let ws = wb.add_sheet("Sheet1");
            ws.cell_mut("A1")?.set_value(cell_value.clone());

            // Derive
            let mut output = String::new();
            derive_content(&wb, &DeriveOptions::default(), &mut output);

            // Check IR contains expected encoding
            let expected_line = format!("A1: {expected_ir}");
            assert!(
                output.contains(&expected_line),
                "Expected IR to contain {expected_line:?}, got:\n{output}",
            );

            // Apply back to a fresh workbook
            let mut wb2 = Workbook::new();
            wb2.add_sheet("Sheet1");
            let body = format!("\n=== Sheet: Sheet1 ===\n{expected_line}\n");
            apply_content(&body, &mut wb2)?;

            let ws2 = wb2.sheet("Sheet1").ok_or("sheet missing")?;
            let cell2 = ws2.cell("A1").ok_or("cell missing")?;

            // Verify value roundtripped correctly
            match cell_value {
                CellValue::String(s) => {
                    assert_eq!(cell2.value(), Some(&CellValue::String(s.clone())));
                }
                CellValue::Number(n) => {
                    if let Some(CellValue::Number(got)) = cell2.value() {
                        assert!(
                            (got - n).abs() < f64::EPSILON,
                            "Number mismatch: expected {n}, got {got}",
                        );
                    } else {
                        return Err(format!("Expected Number({n}), got {:?}", cell2.value()).into());
                    }
                }
                CellValue::Bool(b) => {
                    assert_eq!(cell2.value(), Some(&CellValue::Bool(*b)));
                }
                CellValue::Error(e) => {
                    assert_eq!(cell2.value(), Some(&CellValue::Error(e.clone())));
                }
                _ => {}
            }
        }

        Ok(())
    }

    #[test]
    fn derive_apply_file_roundtrip() -> TestResult {
        let dir = tempfile::tempdir()?;
        let path = dir.path().join("test.xlsx");

        // Create workbook with various cell types
        let mut wb = Workbook::new();
        let ws = wb.add_sheet("Data");
        ws.cell_mut("A1")?.set_value("Category");
        ws.cell_mut("B1")?.set_value("Amount");
        ws.cell_mut("A2")?.set_value("Sales");
        ws.cell_mut("B2")?.set_value(42000);
        ws.cell_mut("A3")?.set_value("Active");
        ws.cell_mut("B3")?.set_value(true);
        ws.cell_mut("B4")?.set_formula("SUM(B2:B3)");
        wb.save(&path)?;

        // Derive
        let ir = crate::derive(&path, DeriveOptions::default())?;

        // Apply unchanged → should produce identical content
        let output_path = dir.path().join("output.xlsx");
        let result = crate::apply(
            &ir,
            &output_path,
            &crate::ApplyOptions {
                force: true, // source path in header won't match
                ..Default::default()
            },
        )?;

        assert_eq!(result.cells_updated, 7); // A1,B1,A2,B2,A3,B3 (values) + B4 (formula)
        assert_eq!(result.cells_created, 0);

        // Derive the output and compare
        let ir2 = crate::derive(&output_path, DeriveOptions::default())?;

        // Strip headers (checksums will differ) and compare bodies
        let body1 = strip_header(&ir);
        let body2 = strip_header(&ir2);
        pretty_assertions::assert_eq!(body1, body2);

        Ok(())
    }

    /// Helper to strip the TOML header, returning just the body.
    fn strip_header(ir: &str) -> &str {
        // Find the second +++
        let first = ir.find("+++").unwrap_or(0);
        let rest = &ir[first + 3..];
        let second = rest.find("+++").unwrap_or(0);
        let after_second = &rest[second + 3..];
        // Skip the newline after +++
        after_second.strip_prefix('\n').unwrap_or(after_second)
    }
}
