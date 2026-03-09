use std::collections::HashMap;
use std::fmt::Display;

use offidized_docx::Document as CoreDocument;
use offidized_ir::{
    self as ir, ApplyOptions, DeriveOptions, Mode, UnifiedDiagnostic, UnifiedDiagnosticSeverity,
    UnifiedDocument, UnifiedEdit, UnifiedEditPayload, UnifiedNode,
};
use offidized_pptx::Presentation as CorePresentation;
use offidized_xlsx::Workbook as CoreWorkbook;
use serde::{Deserialize, Serialize};
use wasm_bindgen::prelude::wasm_bindgen;
#[cfg(target_arch = "wasm32")]
use wasm_bindgen::JsError;
use wasm_bindgen::JsValue;

fn to_js_error(error: impl Display) -> JsValue {
    to_js_error_message(&error.to_string())
}

#[cfg(target_arch = "wasm32")]
fn to_js_error_message(message: &str) -> JsValue {
    JsError::new(message).into()
}

#[cfg(not(target_arch = "wasm32"))]
fn to_js_error_message(_message: &str) -> JsValue {
    JsValue::UNDEFINED
}

fn missing_sheet_message(sheet: &str) -> String {
    format!("worksheet `{sheet}` not found")
}

#[wasm_bindgen]
pub struct Workbook {
    inner: CoreWorkbook,
}

impl Default for Workbook {
    fn default() -> Self {
        Self::new()
    }
}

#[wasm_bindgen]
impl Workbook {
    #[wasm_bindgen(constructor)]
    pub fn new() -> Self {
        Self {
            inner: CoreWorkbook::new(),
        }
    }

    #[wasm_bindgen(js_name = addSheet)]
    pub fn add_sheet(&mut self, name: &str) {
        self.inner.add_sheet(name);
    }

    #[wasm_bindgen(js_name = setCellString)]
    pub fn set_cell_string(&mut self, sheet: &str, cell: &str, value: &str) -> Result<(), JsValue> {
        self.try_set_cell_string(sheet, cell, value)
            .map_err(to_js_error)
    }

    #[wasm_bindgen(js_name = toBytes)]
    pub fn to_bytes(&self) -> Result<Vec<u8>, JsValue> {
        self.inner.to_bytes().map_err(to_js_error)
    }

    pub fn save(&self, path: &str) -> Result<(), JsValue> {
        self.inner.save(path).map_err(to_js_error)
    }
}

impl Workbook {
    fn try_set_cell_string(&mut self, sheet: &str, cell: &str, value: &str) -> Result<(), String> {
        let worksheet = self
            .inner
            .sheet_mut(sheet)
            .ok_or_else(|| missing_sheet_message(sheet))?;

        worksheet
            .cell_mut(cell)
            .map_err(|error| error.to_string())?
            .set_value(value);
        Ok(())
    }
}

#[wasm_bindgen]
pub struct Document {
    inner: CoreDocument,
}

impl Default for Document {
    fn default() -> Self {
        Self::new()
    }
}

#[wasm_bindgen]
impl Document {
    #[wasm_bindgen(constructor)]
    pub fn new() -> Self {
        Self {
            inner: CoreDocument::new(),
        }
    }

    #[wasm_bindgen(js_name = addParagraph)]
    pub fn add_paragraph(&mut self, text: &str) {
        self.inner.add_paragraph(text);
    }

    #[wasm_bindgen(js_name = addHeading)]
    pub fn add_heading(&mut self, text: &str, level: u8) {
        self.inner.add_heading(text, level);
    }

    #[wasm_bindgen(js_name = toBytes)]
    pub fn to_bytes(&self) -> Result<Vec<u8>, JsValue> {
        self.inner.to_bytes().map_err(to_js_error)
    }

    pub fn save(&self, path: &str) -> Result<(), JsValue> {
        self.inner.save(path).map_err(to_js_error)
    }
}

#[wasm_bindgen]
pub struct Presentation {
    inner: CorePresentation,
}

impl Default for Presentation {
    fn default() -> Self {
        Self::new()
    }
}

#[wasm_bindgen]
impl Presentation {
    #[wasm_bindgen(constructor)]
    pub fn new() -> Self {
        Self {
            inner: CorePresentation::new(),
        }
    }

    #[wasm_bindgen(js_name = addSlideWithTitle)]
    pub fn add_slide_with_title(&mut self, title: &str) {
        self.inner.add_slide_with_title(title);
    }

    pub fn save(&mut self, path: &str) -> Result<(), JsValue> {
        self.inner.save(path).map_err(to_js_error)
    }
}

#[cfg(test)]
mod tests {
    use std::path::Path;

    use offidized_docx::Document as CoreDocument;
    use offidized_pptx::Presentation as CorePresentation;
    use offidized_xlsx::{CellValue, Workbook as CoreWorkbook};
    use tempfile::tempdir;

    use super::{missing_sheet_message, Document, Presentation, Workbook};

    fn path_to_str(path: &Path) -> &str {
        match path.to_str() {
            Some(path) => path,
            None => panic!("temporary path should be valid UTF-8"),
        }
    }

    #[test]
    fn workbook_wrapper_sets_string_cells() {
        let mut workbook = Workbook::new();
        workbook.add_sheet("Data");
        assert!(workbook.set_cell_string("Data", "A1", "hello").is_ok());

        let maybe_value = workbook
            .inner
            .sheet("Data")
            .and_then(|sheet| sheet.cell("A1"))
            .and_then(|cell| cell.value());
        assert_eq!(maybe_value, Some(&CellValue::String("hello".to_string())));
    }

    #[test]
    fn workbook_wrapper_reports_sheet_lookup_errors() {
        let mut workbook = Workbook::new();
        let result = workbook.try_set_cell_string("Missing", "A1", "value");

        assert!(result.is_err());
        assert_eq!(result.err(), Some(missing_sheet_message("Missing")));
    }

    #[test]
    fn workbook_wrapper_save_roundtrips() {
        let dir = match tempdir() {
            Ok(dir) => dir,
            Err(error) => panic!("failed to create temp dir: {error}"),
        };
        let output = dir.path().join("workbook.xlsx");

        let mut workbook = Workbook::new();
        workbook.add_sheet("Data");
        assert!(workbook.set_cell_string("Data", "B2", "value").is_ok());
        assert!(workbook.save(path_to_str(&output)).is_ok());

        let reopened = match CoreWorkbook::open(&output) {
            Ok(workbook) => workbook,
            Err(error) => panic!("failed to open saved workbook: {error}"),
        };
        let maybe_value = reopened
            .sheet("Data")
            .and_then(|sheet| sheet.cell("B2"))
            .and_then(|cell| cell.value());
        assert_eq!(maybe_value, Some(&CellValue::String("value".to_string())));
    }

    #[test]
    fn document_wrapper_appends_paragraph_and_heading() {
        let mut document = Document::new();
        document.add_paragraph("body");
        document.add_heading("title", 12);

        assert_eq!(document.inner.paragraphs().len(), 2);
        assert_eq!(document.inner.paragraphs()[0].runs()[0].text(), "body");
        assert_eq!(document.inner.paragraphs()[1].runs()[0].text(), "title");
        assert_eq!(document.inner.paragraphs()[1].heading_level(), Some(9));
    }

    #[test]
    fn document_wrapper_save_roundtrips() {
        let dir = match tempdir() {
            Ok(dir) => dir,
            Err(error) => panic!("failed to create temp dir: {error}"),
        };
        let output = dir.path().join("document.docx");

        let mut document = Document::new();
        document.add_paragraph("one");
        document.add_heading("two", 2);
        assert!(document.save(path_to_str(&output)).is_ok());

        let reopened = match CoreDocument::open(&output) {
            Ok(document) => document,
            Err(error) => panic!("failed to open saved document: {error}"),
        };
        assert_eq!(reopened.paragraphs().len(), 2);
        assert_eq!(reopened.paragraphs()[0].runs()[0].text(), "one");
        assert_eq!(reopened.paragraphs()[1].runs()[0].text(), "two");
        assert_eq!(reopened.paragraphs()[1].heading_level(), Some(2));
    }

    #[test]
    fn presentation_wrapper_adds_titled_slides() {
        let mut presentation = Presentation::new();
        presentation.add_slide_with_title("Intro");

        assert_eq!(presentation.inner.slide_count(), 1);
        assert_eq!(
            presentation.inner.slide(0).map(|slide| slide.title()),
            Some("Intro")
        );
    }

    #[test]
    fn presentation_wrapper_save_roundtrips() {
        let dir = match tempdir() {
            Ok(dir) => dir,
            Err(error) => panic!("failed to create temp dir: {error}"),
        };
        let output = dir.path().join("deck.pptx");

        let mut presentation = Presentation::new();
        presentation.add_slide_with_title("Agenda");
        assert!(presentation.save(path_to_str(&output)).is_ok());

        let reopened = match CorePresentation::open(&output) {
            Ok(presentation) => presentation,
            Err(error) => panic!("failed to open saved presentation: {error}"),
        };
        assert_eq!(reopened.slide_count(), 1);
        assert_eq!(reopened.slide(0).map(|slide| slide.title()), Some("Agenda"));
    }
}

#[wasm_bindgen(js_name = deriveFile)]
pub fn derive_file(data: &[u8], filename: &str, mode: &str) -> Result<String, JsValue> {
    let ir_mode = match mode {
        "content" => Mode::Content,
        "style" => Mode::Style,
        "full" => Mode::Full,
        _ => return Err(to_js_error(format!("unknown mode: {mode}"))),
    };

    let options = DeriveOptions {
        mode: ir_mode,
        sheet: None,
        range: None,
    };

    ir::derive_from_bytes(data, filename, options).map_err(to_js_error)
}

#[wasm_bindgen]
pub struct ApplyResult {
    #[wasm_bindgen(getter_with_clone)]
    pub bytes: Vec<u8>,
    pub cells_updated: usize,
    pub cells_created: usize,
    pub cells_cleared: usize,
    #[wasm_bindgen(getter_with_clone)]
    pub warnings: Vec<String>,
}

#[wasm_bindgen(js_name = applyFile)]
pub fn apply_file(source: &[u8], ir: &str) -> Result<ApplyResult, JsValue> {
    let options = ApplyOptions {
        source_override: None,
        force: false,
    };

    let (bytes, result) = ir::apply_to_bytes(source, ir, &options).map_err(to_js_error)?;

    Ok(ApplyResult {
        bytes,
        cells_updated: result.cells_updated,
        cells_created: result.cells_created,
        cells_cleared: result.cells_cleared,
        warnings: result.warnings,
    })
}

#[derive(Serialize)]
struct UnifiedNodeOut {
    id: String,
    kind: String,
    text: String,
}

#[derive(Serialize)]
struct UnifiedCapabilitiesOut {
    text_nodes: bool,
    table_cells: bool,
    chart_meta: bool,
    style_nodes: bool,
}

#[derive(Serialize)]
struct UnifiedDiagnosticOut {
    severity: String,
    code: String,
    message: String,
    id: Option<String>,
}

#[derive(Serialize)]
struct UnifiedEditReportOut {
    requested: usize,
    applied: usize,
    skipped: usize,
    diagnostics: Vec<UnifiedDiagnosticOut>,
}

#[derive(Serialize)]
struct ChangedNodeOut {
    id: String,
    kind: String,
    before: String,
    after: String,
}

#[derive(Serialize)]
struct UnifiedTargetsOut {
    mode: String,
    capabilities: UnifiedCapabilitiesOut,
    nodes: Vec<UnifiedNodeOut>,
}

#[derive(Serialize)]
struct UnifiedLintOut {
    mode: String,
    edit_count: usize,
    diagnostics: Vec<UnifiedDiagnosticOut>,
}

#[derive(Serialize)]
struct UnifiedPreviewOut {
    mode: String,
    report: UnifiedEditReportOut,
    lint_diagnostics: Vec<UnifiedDiagnosticOut>,
    changed_nodes: Vec<ChangedNodeOut>,
    ir: String,
}

#[derive(Deserialize)]
struct JsonUnifiedEdit {
    id: String,
    #[serde(default)]
    text: Option<String>,
    #[serde(default)]
    group: Option<String>,
    #[serde(default)]
    payload: Option<JsonUnifiedPayload>,
}

#[derive(Deserialize)]
#[serde(tag = "kind", rename_all = "snake_case")]
enum JsonUnifiedPayload {
    XlsxCellStyle {
        #[serde(default)]
        bold: Option<bool>,
        #[serde(default)]
        italic: Option<bool>,
        #[serde(default)]
        number_format: Option<String>,
    },
    PptxTextStyle {
        #[serde(default)]
        bold: Option<bool>,
        #[serde(default)]
        italic: Option<bool>,
        #[serde(default)]
        font_size: Option<u32>,
        #[serde(default)]
        font_color: Option<String>,
        #[serde(default)]
        font_name: Option<String>,
    },
}

fn parse_mode(mode: &str) -> Result<Mode, JsValue> {
    match mode {
        "content" => Ok(Mode::Content),
        "full" => Ok(Mode::Full),
        "style" => Ok(Mode::Style),
        _ => Err(to_js_error(format!("unknown mode: {mode}"))),
    }
}

fn diagnostics_out(diagnostics: Vec<UnifiedDiagnostic>) -> Vec<UnifiedDiagnosticOut> {
    diagnostics
        .into_iter()
        .map(|diag| UnifiedDiagnosticOut {
            severity: match diag.severity {
                UnifiedDiagnosticSeverity::Error => "error".to_string(),
                UnifiedDiagnosticSeverity::Warning => "warning".to_string(),
            },
            code: diag.code,
            message: diag.message,
            id: diag.id,
        })
        .collect()
}

fn nodes_out(nodes: &[UnifiedNode]) -> Vec<UnifiedNodeOut> {
    nodes
        .iter()
        .map(|node| UnifiedNodeOut {
            id: node.id.to_string(),
            kind: format!("{:?}", node.kind),
            text: node.text.clone(),
        })
        .collect()
}

fn parse_unified_edits(edits_json: &str) -> Result<Vec<UnifiedEdit>, JsValue> {
    let parsed: Vec<JsonUnifiedEdit> = serde_json::from_str(edits_json)
        .map_err(|error| to_js_error(format!("invalid edits json: {error}")))?;
    let mut edits = Vec::with_capacity(parsed.len());

    for item in parsed {
        let mut edit = UnifiedEdit::new(item.id, item.text.unwrap_or_default());
        if let Some(group) = item.group {
            edit = edit.with_group(group);
        }
        if let Some(payload) = item.payload {
            let mapped = match payload {
                JsonUnifiedPayload::XlsxCellStyle {
                    bold,
                    italic,
                    number_format,
                } => {
                    let mut patch = ir::CellStylePatch::new();
                    if let Some(value) = bold {
                        patch.set_bold(value);
                    }
                    if let Some(value) = italic {
                        patch.set_italic(value);
                    }
                    if let Some(value) = number_format {
                        patch.set_number_format(value);
                    }
                    UnifiedEditPayload::XlsxCellStyle(patch)
                }
                JsonUnifiedPayload::PptxTextStyle {
                    bold,
                    italic,
                    font_size,
                    font_color,
                    font_name,
                } => {
                    let mut patch = ir::PptxTextStylePatch::new();
                    if let Some(value) = bold {
                        patch.set_bold(value);
                    }
                    if let Some(value) = italic {
                        patch.set_italic(value);
                    }
                    if let Some(value) = font_size {
                        patch.set_font_size(value);
                    }
                    if let Some(value) = font_color {
                        patch.set_font_color(value);
                    }
                    if let Some(value) = font_name {
                        patch.set_font_name(value);
                    }
                    UnifiedEditPayload::PptxTextStyle(patch)
                }
            };
            edit = edit.with_payload(mapped);
        }
        edits.push(edit);
    }

    Ok(edits)
}

fn derive_unified_doc(data: &[u8], filename: &str, mode: &str) -> Result<UnifiedDocument, JsValue> {
    let parsed_mode = parse_mode(mode)?;
    let ir = ir::derive_from_bytes(
        data,
        filename,
        DeriveOptions {
            mode: parsed_mode,
            sheet: None,
            range: None,
        },
    )
    .map_err(to_js_error)?;
    UnifiedDocument::from_ir(&ir).map_err(to_js_error)
}

fn preview_unified_edits_inner(
    data: &[u8],
    filename: &str,
    mode: &str,
    edits_json: &str,
    lint: bool,
) -> Result<UnifiedPreviewOut, JsValue> {
    let mut doc = derive_unified_doc(data, filename, mode)?;
    let before_nodes = nodes_out(doc.nodes());
    let edits = parse_unified_edits(edits_json)?;

    let lint_diagnostics = if lint {
        diagnostics_out(doc.lint_edits(edits.as_slice()))
    } else {
        Vec::new()
    };

    let report = doc.apply_edits(edits.as_slice()).map_err(to_js_error)?;
    let report_out = UnifiedEditReportOut {
        requested: report.requested,
        applied: report.applied,
        skipped: report.skipped,
        diagnostics: diagnostics_out(report.diagnostics),
    };

    let after_nodes = nodes_out(doc.nodes());
    let mut before_by_id = HashMap::<String, UnifiedNodeOut>::new();
    for node in before_nodes {
        before_by_id.insert(node.id.clone(), node);
    }

    let mut changed_nodes = Vec::new();
    for node in after_nodes {
        if let Some(before) = before_by_id.get(node.id.as_str()) {
            if before.text != node.text {
                changed_nodes.push(ChangedNodeOut {
                    id: node.id,
                    kind: node.kind,
                    before: before.text.clone(),
                    after: node.text,
                });
            }
        }
    }

    Ok(UnifiedPreviewOut {
        mode: mode.to_string(),
        report: report_out,
        lint_diagnostics,
        changed_nodes,
        ir: doc.to_ir(),
    })
}

#[wasm_bindgen(js_name = deriveUnifiedTargets)]
pub fn derive_unified_targets(data: &[u8], filename: &str, mode: &str) -> Result<String, JsValue> {
    let doc = derive_unified_doc(data, filename, mode)?;
    let caps = doc.capabilities();

    let out = UnifiedTargetsOut {
        mode: mode.to_string(),
        capabilities: UnifiedCapabilitiesOut {
            text_nodes: caps.text_nodes,
            table_cells: caps.table_cells,
            chart_meta: caps.chart_meta,
            style_nodes: caps.style_nodes,
        },
        nodes: nodes_out(doc.nodes()),
    };
    serde_json::to_string(&out).map_err(to_js_error)
}

#[wasm_bindgen(js_name = lintUnifiedEdits)]
pub fn lint_unified_edits(
    data: &[u8],
    filename: &str,
    mode: &str,
    edits_json: &str,
) -> Result<String, JsValue> {
    let doc = derive_unified_doc(data, filename, mode)?;
    let edits = parse_unified_edits(edits_json)?;
    let diagnostics = diagnostics_out(doc.lint_edits(edits.as_slice()));
    let out = UnifiedLintOut {
        mode: mode.to_string(),
        edit_count: edits.len(),
        diagnostics,
    };
    serde_json::to_string(&out).map_err(to_js_error)
}

#[wasm_bindgen(js_name = previewUnifiedEdits)]
pub fn preview_unified_edits(
    data: &[u8],
    filename: &str,
    mode: &str,
    edits_json: &str,
    lint: bool,
) -> Result<String, JsValue> {
    let out = preview_unified_edits_inner(data, filename, mode, edits_json, lint)?;
    serde_json::to_string(&out).map_err(to_js_error)
}
