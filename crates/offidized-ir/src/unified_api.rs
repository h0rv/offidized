use std::fmt;
use std::path::Path;
use std::str::FromStr;

use crate::{
    apply, derive, ApplyOptions, ApplyResult, DeriveOptions, Format, IrError, IrHeader, Mode,
    Result,
};

const FULL_MODE_SEPARATOR: &str = "\n--- style ---\n";

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum UnifiedNodeKind {
    SpreadsheetCell,
    XlsxTableCell,
    XlsxChartTitle,
    XlsxChartSeriesName,
    XlsxCellStyle,
    Paragraph,
    DocxParagraphStyle,
    DocxTableCell,
    SlideTitle,
    SlideSubtitle,
    SlideNotes,
    SlideShape,
    PptxShapeStyle,
    PptxTableCell,
    PptxChartTitle,
    PptxChartSeriesName,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub enum UnifiedNodeId {
    SpreadsheetCell {
        sheet: String,
        cell: String,
    },
    XlsxTableCell {
        sheet: String,
        table: String,
        cell: String,
    },
    XlsxChartTitle {
        sheet: String,
        chart: usize,
    },
    XlsxChartSeriesName {
        sheet: String,
        chart: usize,
        series: usize,
    },
    XlsxCellStyle {
        sheet: String,
        cell: String,
    },
    Paragraph {
        index: usize,
    },
    DocxParagraphStyle {
        index: usize,
    },
    DocxTableCell {
        table: usize,
        row: usize,
        col: usize,
    },
    SlideTitle {
        slide: usize,
    },
    SlideSubtitle {
        slide: usize,
    },
    SlideNotes {
        slide: usize,
    },
    SlideShape {
        slide: usize,
        anchor: String,
    },
    PptxShapeStyle {
        slide: usize,
        anchor: String,
    },
    PptxTableCell {
        slide: usize,
        table: usize,
        row: usize,
        col: usize,
    },
    PptxChartTitle {
        slide: usize,
        chart: usize,
    },
    PptxChartSeriesName {
        slide: usize,
        chart: usize,
        series: usize,
    },
}

impl fmt::Display for UnifiedNodeId {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::SpreadsheetCell { sheet, cell } => write!(f, "sheet:{sheet}/cell:{cell}"),
            Self::XlsxTableCell { sheet, table, cell } => {
                write!(f, "sheet:{sheet}/table:{table}/cell:{cell}")
            }
            Self::XlsxChartTitle { sheet, chart } => {
                write!(f, "sheet:{sheet}/chart:{chart}/title")
            }
            Self::XlsxChartSeriesName {
                sheet,
                chart,
                series,
            } => write!(f, "sheet:{sheet}/chart:{chart}/series:{series}/name"),
            Self::XlsxCellStyle { sheet, cell } => write!(f, "sheet:{sheet}/cell:{cell}/style"),
            Self::Paragraph { index } => write!(f, "paragraph:{index}"),
            Self::DocxParagraphStyle { index } => write!(f, "paragraph:{index}/style"),
            Self::DocxTableCell { table, row, col } => {
                write!(f, "docx_table:{table}/cell:{row},{col}")
            }
            Self::SlideTitle { slide } => write!(f, "slide:{slide}/title"),
            Self::SlideSubtitle { slide } => write!(f, "slide:{slide}/subtitle"),
            Self::SlideNotes { slide } => write!(f, "slide:{slide}/notes"),
            Self::SlideShape { slide, anchor } => write!(f, "slide:{slide}/shape:{anchor}"),
            Self::PptxShapeStyle { slide, anchor } => {
                write!(f, "slide:{slide}/shape:{anchor}/style")
            }
            Self::PptxTableCell {
                slide,
                table,
                row,
                col,
            } => write!(f, "slide:{slide}/table:{table}/cell:{row},{col}"),
            Self::PptxChartTitle { slide, chart } => write!(f, "slide:{slide}/chart:{chart}/title"),
            Self::PptxChartSeriesName {
                slide,
                chart,
                series,
            } => write!(f, "slide:{slide}/chart:{chart}/series:{series}/name"),
        }
    }
}

impl FromStr for UnifiedNodeId {
    type Err = IrError;

    fn from_str(s: &str) -> Result<Self> {
        parse_unified_node_id(s)
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct UnifiedNode {
    pub id: UnifiedNodeId,
    pub kind: UnifiedNodeKind,
    pub text: String,
}

impl UnifiedNode {
    pub fn id_string(&self) -> String {
        self.id.to_string()
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct UnifiedEdit {
    pub id: String,
    pub text: String,
    pub group: Option<String>,
    pub payload: Option<UnifiedEditPayload>,
}

impl UnifiedEdit {
    pub fn new(id: impl Into<String>, text: impl Into<String>) -> Self {
        Self {
            id: id.into(),
            text: text.into(),
            group: None,
            payload: None,
        }
    }

    pub fn with_group(mut self, group: impl Into<String>) -> Self {
        let group = group.into();
        self.group = if group.trim().is_empty() {
            None
        } else {
            Some(group)
        };
        self
    }

    pub fn with_payload(mut self, payload: UnifiedEditPayload) -> Self {
        self.payload = Some(payload);
        self
    }

    pub fn typed_id(&self) -> Result<UnifiedNodeId> {
        UnifiedNodeId::from_str(self.id.as_str())
    }
}

#[derive(Debug, Clone, PartialEq, Eq, Default)]
pub struct UnifiedCapabilities {
    pub text_nodes: bool,
    pub table_cells: bool,
    pub chart_meta: bool,
    pub style_nodes: bool,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum UnifiedDiagnosticSeverity {
    Error,
    Warning,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct UnifiedDiagnostic {
    pub severity: UnifiedDiagnosticSeverity,
    pub code: String,
    pub message: String,
    pub id: Option<String>,
}

#[derive(Debug, Clone, PartialEq, Eq, Default)]
pub struct UnifiedEditReport {
    pub requested: usize,
    pub applied: usize,
    pub skipped: usize,
    pub diagnostics: Vec<UnifiedDiagnostic>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct CellStylePatch {
    bold: Option<bool>,
    italic: Option<bool>,
    number_format: Option<String>,
}

impl CellStylePatch {
    pub fn new() -> Self {
        Self {
            bold: None,
            italic: None,
            number_format: None,
        }
    }

    pub fn set_bold(&mut self, value: bool) {
        self.bold = Some(value);
    }

    pub fn set_italic(&mut self, value: bool) {
        self.italic = Some(value);
    }

    pub fn set_number_format(&mut self, value: impl Into<String>) {
        self.number_format = Some(value.into());
    }

    fn is_empty(&self) -> bool {
        self.bold.is_none() && self.italic.is_none() && self.number_format.is_none()
    }
}

impl Default for CellStylePatch {
    fn default() -> Self {
        Self::new()
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct PptxTextStylePatch {
    bold: Option<bool>,
    italic: Option<bool>,
    font_size: Option<u32>,
    font_color: Option<String>,
    font_name: Option<String>,
}

impl PptxTextStylePatch {
    pub fn new() -> Self {
        Self {
            bold: None,
            italic: None,
            font_size: None,
            font_color: None,
            font_name: None,
        }
    }

    pub fn set_bold(&mut self, value: bool) {
        self.bold = Some(value);
    }

    pub fn set_italic(&mut self, value: bool) {
        self.italic = Some(value);
    }

    pub fn set_font_size(&mut self, value: u32) {
        self.font_size = Some(value);
    }

    pub fn set_font_color(&mut self, value: impl Into<String>) {
        self.font_color = Some(value.into());
    }

    pub fn set_font_name(&mut self, value: impl Into<String>) {
        self.font_name = Some(value.into());
    }

    fn is_empty(&self) -> bool {
        self.bold.is_none()
            && self.italic.is_none()
            && self.font_size.is_none()
            && self.font_color.is_none()
            && self.font_name.is_none()
    }
}

impl Default for PptxTextStylePatch {
    fn default() -> Self {
        Self::new()
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub enum UnifiedEditPayload {
    XlsxCellStyle(CellStylePatch),
    PptxTextStyle(PptxTextStylePatch),
}

#[derive(Debug, Clone, PartialEq, Eq)]
enum DirectMutation {
    XlsxChartTitle {
        sheet: String,
        chart: usize,
        title: String,
    },
    XlsxChartSeriesName {
        sheet: String,
        chart: usize,
        series: usize,
        name: String,
    },
    XlsxCellStyle {
        sheet: String,
        cell: String,
        patch: CellStylePatch,
    },
    DocxParagraphStyle {
        index: usize,
        style_id: Option<String>,
    },
    PptxChartTitle {
        slide: usize,
        chart: usize,
        title: String,
    },
    PptxChartSeriesName {
        slide: usize,
        chart: usize,
        series: usize,
        name: String,
    },
    PptxShapeStyle {
        slide: usize,
        shape_name: String,
        patch: PptxTextStylePatch,
    },
}

#[derive(Debug, Clone)]
pub struct UnifiedDocument {
    header: IrHeader,
    content_body: String,
    style_body: Option<String>,
    nodes: Vec<UnifiedNode>,
    capabilities: UnifiedCapabilities,
    pending_direct_mutations: Vec<DirectMutation>,
}

impl UnifiedDocument {
    pub fn derive(path: &Path, options: UnifiedDeriveOptions) -> Result<Self> {
        let ir = derive(
            path,
            DeriveOptions {
                mode: options.mode,
                sheet: options.sheet,
                range: options.range,
            },
        )?;
        let mut doc = Self::from_ir(&ir)?;
        doc.augment_from_source(path)?;
        Ok(doc)
    }

    pub fn from_ir(ir: &str) -> Result<Self> {
        let (header, body) = IrHeader::parse(ir)?;

        let (content_body, style_body) = match header.mode {
            Mode::Content => (body, None),
            Mode::Full => {
                let (content, style) = split_full_body(&body);
                (content.to_string(), Some(style.to_string()))
            }
            Mode::Style => {
                return Err(IrError::UnsupportedMode {
                    format: header.format.as_str().to_string(),
                    mode: header.mode.as_str().to_string(),
                })
            }
        };

        let nodes = list_nodes_for_format(header.format, content_body.as_str())?;
        let capabilities = capabilities_for_format(header.format, header.mode);

        Ok(Self {
            header,
            content_body,
            style_body,
            nodes,
            capabilities,
            pending_direct_mutations: Vec::new(),
        })
    }

    pub fn nodes(&self) -> &[UnifiedNode] {
        self.nodes.as_slice()
    }

    pub fn capabilities(&self) -> &UnifiedCapabilities {
        &self.capabilities
    }

    pub fn apply_edits(&mut self, edits: &[UnifiedEdit]) -> Result<UnifiedEditReport> {
        if edits.is_empty() {
            return Ok(UnifiedEditReport::default());
        }

        let (updated, report) = apply_edits_for_format(
            self.header.format,
            self.content_body.as_str(),
            edits,
            &self.capabilities,
            &mut self.pending_direct_mutations,
        )?;
        self.content_body = updated;
        self.nodes = list_nodes_for_format(self.header.format, self.content_body.as_str())?;
        Ok(report)
    }

    pub fn lint_edits(&self, edits: &[UnifiedEdit]) -> Vec<UnifiedDiagnostic> {
        lint_unified_edits(self.nodes.as_slice(), edits)
    }

    pub fn to_ir(&self) -> String {
        let mut out = self.header.write();
        out.push_str(self.content_body.as_str());
        if self.header.mode == Mode::Full {
            if let Some(style) = &self.style_body {
                out.push_str(FULL_MODE_SEPARATOR);
                out.push_str(style.as_str());
            }
        }
        out
    }

    pub fn save_as(&self, output: &Path, apply_options: &ApplyOptions) -> Result<ApplyResult> {
        if self.pending_direct_mutations.is_empty() {
            return apply(self.to_ir().as_str(), output, apply_options);
        }

        let tmp = tempfile::NamedTempFile::new()?;
        let base_result = apply(self.to_ir().as_str(), tmp.path(), apply_options)?;
        apply_direct_mutations(
            tmp.path(),
            self.header.format,
            &self.pending_direct_mutations,
        )?;

        std::fs::copy(tmp.path(), output)?;
        Ok(base_result)
    }

    fn augment_from_source(&mut self, path: &Path) -> Result<()> {
        match self.header.format {
            Format::Xlsx => augment_xlsx_nodes(path, &mut self.nodes, &mut self.capabilities)?,
            Format::Docx => augment_docx_nodes(path, &mut self.nodes, &mut self.capabilities)?,
            Format::Pptx => augment_pptx_nodes(path, &mut self.nodes, &mut self.capabilities)?,
        }
        Ok(())
    }
}

#[derive(Debug, Clone)]
pub struct UnifiedDeriveOptions {
    pub mode: Mode,
    pub sheet: Option<String>,
    pub range: Option<String>,
}

impl Default for UnifiedDeriveOptions {
    fn default() -> Self {
        Self {
            mode: Mode::Content,
            sheet: None,
            range: None,
        }
    }
}

pub fn derive_content_nodes(path: &Path) -> Result<Vec<UnifiedNode>> {
    let doc = UnifiedDocument::derive(path, UnifiedDeriveOptions::default())?;
    Ok(doc.nodes().to_vec())
}

pub fn list_nodes_from_ir(ir: &str) -> Result<Vec<UnifiedNode>> {
    let doc = UnifiedDocument::from_ir(ir)?;
    Ok(doc.nodes().to_vec())
}

pub fn apply_edits_to_ir(ir: &str, edits: &[UnifiedEdit]) -> Result<String> {
    let mut doc = UnifiedDocument::from_ir(ir)?;
    let _ = doc.apply_edits(edits)?;
    Ok(doc.to_ir())
}

pub fn edit_file_content(
    source: &Path,
    output: &Path,
    edits: &[UnifiedEdit],
    apply_options: &ApplyOptions,
) -> Result<ApplyResult> {
    let mut doc = UnifiedDocument::derive(source, UnifiedDeriveOptions::default())?;
    let _ = doc.apply_edits(edits)?;
    doc.save_as(output, apply_options)
}

fn capabilities_for_format(format: Format, mode: Mode) -> UnifiedCapabilities {
    match format {
        Format::Xlsx => UnifiedCapabilities {
            text_nodes: true,
            table_cells: true,
            chart_meta: true,
            style_nodes: true,
        },
        Format::Docx => UnifiedCapabilities {
            text_nodes: true,
            table_cells: true,
            chart_meta: false,
            style_nodes: mode != Mode::Content,
        },
        Format::Pptx => UnifiedCapabilities {
            text_nodes: true,
            table_cells: true,
            chart_meta: true,
            style_nodes: mode != Mode::Content,
        },
    }
}

fn list_nodes_for_format(format: Format, body: &str) -> Result<Vec<UnifiedNode>> {
    match format {
        Format::Xlsx => list_xlsx_nodes(body),
        Format::Docx => list_docx_nodes(body),
        Format::Pptx => list_pptx_nodes(body),
    }
}

fn lint_unified_edits(nodes: &[UnifiedNode], edits: &[UnifiedEdit]) -> Vec<UnifiedDiagnostic> {
    use std::collections::HashMap;

    let mut node_counts: HashMap<String, usize> = HashMap::new();
    let mut docx_table_bounds: HashMap<usize, (usize, usize)> = HashMap::new();
    let mut pptx_table_bounds: HashMap<(usize, usize), (usize, usize)> = HashMap::new();

    for node in nodes {
        let key = node.id.to_string();
        *node_counts.entry(key).or_insert(0) += 1;
        match &node.id {
            UnifiedNodeId::DocxTableCell { table, row, col } => {
                let entry = docx_table_bounds.entry(*table).or_insert((0, 0));
                entry.0 = entry.0.max(*row);
                entry.1 = entry.1.max(*col);
            }
            UnifiedNodeId::PptxTableCell {
                slide,
                table,
                row,
                col,
            } => {
                let entry = pptx_table_bounds.entry((*slide, *table)).or_insert((0, 0));
                entry.0 = entry.0.max(*row);
                entry.1 = entry.1.max(*col);
            }
            _ => {}
        }
    }

    let mut diagnostics = Vec::new();
    for edit in edits {
        let typed_id = match edit.typed_id() {
            Ok(id) => id,
            Err(error) => {
                diagnostics.push(UnifiedDiagnostic {
                    severity: UnifiedDiagnosticSeverity::Error,
                    code: "invalid_id".to_string(),
                    message: format!("invalid unified node id '{}': {error}", edit.id),
                    id: Some(edit.id.clone()),
                });
                continue;
            }
        };
        let key = typed_id.to_string();
        let count = node_counts.get(key.as_str()).copied().unwrap_or(0);
        if count == 0 {
            let (code, message) = match typed_id {
                UnifiedNodeId::DocxTableCell { table, row, col } => {
                    if let Some((max_row, max_col)) = docx_table_bounds.get(&table) {
                        if row > *max_row || col > *max_col {
                            (
                                "invalid_table_coordinates",
                                format!(
                                    "docx table coordinates out of bounds: row={row}, col={col}, max_row={max_row}, max_col={max_col}"
                                ),
                            )
                        } else {
                            ("missing_target", "target id not found".to_string())
                        }
                    } else {
                        ("missing_target", "target id not found".to_string())
                    }
                }
                UnifiedNodeId::PptxTableCell {
                    slide,
                    table,
                    row,
                    col,
                } => {
                    if let Some((max_row, max_col)) = pptx_table_bounds.get(&(slide, table)) {
                        if row > *max_row || col > *max_col {
                            (
                                "invalid_table_coordinates",
                                format!(
                                    "pptx table coordinates out of bounds: row={row}, col={col}, max_row={max_row}, max_col={max_col}"
                                ),
                            )
                        } else {
                            ("missing_target", "target id not found".to_string())
                        }
                    } else {
                        ("missing_target", "target id not found".to_string())
                    }
                }
                _ => ("missing_target", "target id not found".to_string()),
            };
            diagnostics.push(UnifiedDiagnostic {
                severity: UnifiedDiagnosticSeverity::Warning,
                code: code.to_string(),
                message,
                id: Some(edit.id.clone()),
            });
            continue;
        }

        if count > 1
            && matches!(
                typed_id,
                UnifiedNodeId::SlideShape { .. } | UnifiedNodeId::PptxShapeStyle { .. }
            )
        {
            diagnostics.push(UnifiedDiagnostic {
                severity: UnifiedDiagnosticSeverity::Warning,
                code: "ambiguous_anchor".to_string(),
                message: "shape anchor resolves to multiple nodes".to_string(),
                id: Some(edit.id.clone()),
            });
        }
    }

    diagnostics
}

fn apply_edits_for_format(
    format: Format,
    body: &str,
    edits: &[UnifiedEdit],
    caps: &UnifiedCapabilities,
    pending_direct_mutations: &mut Vec<DirectMutation>,
) -> Result<(String, UnifiedEditReport)> {
    let mut report = UnifiedEditReport {
        requested: edits.len(),
        ..UnifiedEditReport::default()
    };

    let mut updated = body.to_string();
    let mut processed = vec![false; edits.len()];
    let mut index = 0_usize;
    while index < edits.len() {
        if processed[index] {
            index += 1;
            continue;
        }
        let edit = &edits[index];
        if let Some(group) = edit.group.as_deref() {
            let mut group_indices = Vec::new();
            for (i, candidate) in edits.iter().enumerate().skip(index) {
                if !processed[i] && candidate.group.as_deref() == Some(group) {
                    group_indices.push(i);
                }
            }
            for i in &group_indices {
                processed[*i] = true;
            }

            let mut group_body = updated.clone();
            let mut group_mutations = pending_direct_mutations.clone();
            let mut group_applied = 0_usize;
            let mut group_failed = false;
            let mut fail_message = String::new();

            for edit_idx in &group_indices {
                let candidate = &edits[*edit_idx];
                let outcome = apply_single_edit(
                    format,
                    group_body.as_str(),
                    candidate,
                    caps,
                    &mut group_mutations,
                );
                match outcome {
                    Ok(Some(next)) => {
                        group_body = next;
                        group_applied += 1;
                    }
                    Ok(None) => {
                        group_failed = true;
                        fail_message = "target id not found in derived content".to_string();
                        break;
                    }
                    Err(error) => {
                        group_failed = true;
                        fail_message = error.to_string();
                        break;
                    }
                }
            }

            if group_failed {
                report.skipped += group_indices.len();
                for edit_idx in group_indices {
                    report.diagnostics.push(UnifiedDiagnostic {
                        severity: UnifiedDiagnosticSeverity::Warning,
                        code: "group_aborted".to_string(),
                        message: format!("edit group '{group}' aborted: {fail_message}"),
                        id: Some(edits[edit_idx].id.clone()),
                    });
                }
            } else {
                updated = group_body;
                *pending_direct_mutations = group_mutations;
                report.applied += group_applied;
            }
        } else {
            processed[index] = true;
            match apply_single_edit(
                format,
                updated.as_str(),
                edit,
                caps,
                pending_direct_mutations,
            ) {
                Ok(Some(next)) => {
                    updated = next;
                    report.applied += 1;
                }
                Ok(None) => {
                    report.skipped += 1;
                    report.diagnostics.push(UnifiedDiagnostic {
                        severity: UnifiedDiagnosticSeverity::Warning,
                        code: "target_not_found".to_string(),
                        message: "target id not found in derived content".to_string(),
                        id: Some(edit.id.clone()),
                    });
                }
                Err(error) => {
                    report.skipped += 1;
                    report.diagnostics.push(UnifiedDiagnostic {
                        severity: UnifiedDiagnosticSeverity::Error,
                        code: "invalid_edit".to_string(),
                        message: error.to_string(),
                        id: Some(edit.id.clone()),
                    });
                }
            }
        }
        index += 1;
    }

    Ok((updated, report))
}

fn apply_single_edit(
    format: Format,
    body: &str,
    edit: &UnifiedEdit,
    caps: &UnifiedCapabilities,
    pending_direct_mutations: &mut Vec<DirectMutation>,
) -> Result<Option<String>> {
    let typed_id = match edit.typed_id() {
        Ok(id) => id,
        Err(error) => {
            return Err(IrError::InvalidBody(format!(
                "invalid unified node id '{}': {error}",
                edit.id
            )))
        }
    };

    match format {
        Format::Xlsx => apply_xlsx_edit(body, &typed_id, edit, pending_direct_mutations),
        Format::Docx => apply_docx_edit(body, &typed_id, edit, caps, pending_direct_mutations),
        Format::Pptx => apply_pptx_edit(body, &typed_id, edit, caps, pending_direct_mutations),
    }
}

fn apply_xlsx_edit(
    body: &str,
    id: &UnifiedNodeId,
    edit: &UnifiedEdit,
    pending_direct_mutations: &mut Vec<DirectMutation>,
) -> Result<Option<String>> {
    let text = edit.text.as_str();
    match id {
        UnifiedNodeId::XlsxChartTitle { sheet, chart } => {
            pending_direct_mutations.push(DirectMutation::XlsxChartTitle {
                sheet: sheet.clone(),
                chart: *chart,
                title: text.to_string(),
            });
            return Ok(Some(body.to_string()));
        }
        UnifiedNodeId::XlsxChartSeriesName {
            sheet,
            chart,
            series,
        } => {
            pending_direct_mutations.push(DirectMutation::XlsxChartSeriesName {
                sheet: sheet.clone(),
                chart: *chart,
                series: *series,
                name: text.to_string(),
            });
            return Ok(Some(body.to_string()));
        }
        UnifiedNodeId::XlsxCellStyle { sheet, cell } => {
            let patch =
                if let Some(UnifiedEditPayload::XlsxCellStyle(payload)) = edit.payload.as_ref() {
                    payload.clone()
                } else {
                    parse_cell_style_patch(text)?
                };
            if !patch.is_empty() {
                pending_direct_mutations.push(DirectMutation::XlsxCellStyle {
                    sheet: sheet.clone(),
                    cell: cell.clone(),
                    patch,
                });
            }
            return Ok(Some(body.to_string()));
        }
        _ => {}
    }

    let (target_sheet, target_cell) = match id {
        UnifiedNodeId::SpreadsheetCell { sheet, cell } => (sheet, cell),
        UnifiedNodeId::XlsxTableCell { sheet, cell, .. } => (sheet, cell),
        _ => return Ok(None),
    };

    let mut lines: Vec<String> = body.lines().map(ToString::to_string).collect();
    let mut current_sheet: Option<String> = None;

    for line in &mut lines {
        if let Some(sheet_name) = parse_sheet_header(line.as_str()) {
            current_sheet = Some(sheet_name.to_string());
            continue;
        }

        let Some((left, _right)) = line.split_once(": ") else {
            continue;
        };
        if !looks_like_cell_reference(left) {
            continue;
        }

        if current_sheet.as_deref() == Some(target_sheet.as_str()) && left == target_cell.as_str() {
            *line = format!("{left}: {text}");
            return Ok(Some(join_lines(lines.as_slice(), body.ends_with('\n'))));
        }
    }

    Ok(None)
}

fn apply_docx_edit(
    body: &str,
    id: &UnifiedNodeId,
    edit: &UnifiedEdit,
    caps: &UnifiedCapabilities,
    pending_direct_mutations: &mut Vec<DirectMutation>,
) -> Result<Option<String>> {
    let text = edit.text.as_str();
    let mut lines: Vec<String> = body.lines().map(ToString::to_string).collect();

    match id {
        UnifiedNodeId::Paragraph { index } => {
            for line in &mut lines {
                if let Some((line_index, _)) = parse_docx_paragraph_line(line.as_str()) {
                    if &line_index == index {
                        *line = format!("[p{index}] {text}");
                        return Ok(Some(join_lines(lines.as_slice(), body.ends_with('\n'))));
                    }
                }
            }
            Ok(None)
        }
        UnifiedNodeId::DocxParagraphStyle { index } => {
            let style_id = if text.trim().is_empty() {
                None
            } else {
                Some(text.trim().to_string())
            };
            pending_direct_mutations.push(DirectMutation::DocxParagraphStyle {
                index: *index,
                style_id,
            });
            Ok(Some(body.to_string()))
        }
        UnifiedNodeId::DocxTableCell { table, row, col } => {
            if !caps.table_cells {
                return Ok(None);
            }
            edit_docx_table_cell(lines, *table, *row, *col, text, body.ends_with('\n'))
        }
        _ => Ok(None),
    }
}

fn apply_pptx_edit(
    body: &str,
    id: &UnifiedNodeId,
    edit: &UnifiedEdit,
    caps: &UnifiedCapabilities,
    pending_direct_mutations: &mut Vec<DirectMutation>,
) -> Result<Option<String>> {
    let text = edit.text.as_str();
    match id {
        UnifiedNodeId::PptxChartTitle { slide, chart } => {
            pending_direct_mutations.push(DirectMutation::PptxChartTitle {
                slide: *slide,
                chart: *chart,
                title: text.to_string(),
            });
            return Ok(Some(body.to_string()));
        }
        UnifiedNodeId::PptxChartSeriesName {
            slide,
            chart,
            series,
        } => {
            pending_direct_mutations.push(DirectMutation::PptxChartSeriesName {
                slide: *slide,
                chart: *chart,
                series: *series,
                name: text.to_string(),
            });
            return Ok(Some(body.to_string()));
        }
        UnifiedNodeId::PptxShapeStyle { slide, anchor } => {
            let patch =
                if let Some(UnifiedEditPayload::PptxTextStyle(payload)) = edit.payload.as_ref() {
                    payload.clone()
                } else {
                    parse_pptx_text_style_patch(text)?
                };
            if !patch.is_empty() {
                let shape_name =
                    parse_shape_name_from_anchor(anchor.as_str()).ok_or_else(|| {
                        IrError::InvalidBody(format!("invalid shape anchor for style id: {anchor}"))
                    })?;
                pending_direct_mutations.push(DirectMutation::PptxShapeStyle {
                    slide: *slide,
                    shape_name,
                    patch,
                });
            }
            return Ok(Some(body.to_string()));
        }
        _ => {}
    }

    let mut lines: Vec<String> = body.lines().map(ToString::to_string).collect();

    match id {
        UnifiedNodeId::SlideTitle { slide }
        | UnifiedNodeId::SlideSubtitle { slide }
        | UnifiedNodeId::SlideNotes { slide }
        | UnifiedNodeId::SlideShape { slide, .. }
        | UnifiedNodeId::PptxTableCell { slide, .. } => {
            let mut current_slide: Option<usize> = None;
            let mut i = 0_usize;
            while i < lines.len() {
                if let Some(s) = parse_slide_header(lines[i].as_str()) {
                    current_slide = Some(s);
                    i += 1;
                    continue;
                }
                if current_slide != Some(*slide) {
                    i += 1;
                    continue;
                }

                match id {
                    UnifiedNodeId::SlideTitle { .. } if lines[i].starts_with("[title] ") => {
                        lines[i] = format!("[title] {text}");
                        return Ok(Some(join_lines(lines.as_slice(), body.ends_with('\n'))));
                    }
                    UnifiedNodeId::SlideSubtitle { .. } if lines[i].starts_with("[subtitle] ") => {
                        lines[i] = format!("[subtitle] {text}");
                        return Ok(Some(join_lines(lines.as_slice(), body.ends_with('\n'))));
                    }
                    UnifiedNodeId::SlideNotes { .. } if lines[i].starts_with("[notes] ") => {
                        lines[i] = format!("[notes] {text}");
                        return Ok(Some(join_lines(lines.as_slice(), body.ends_with('\n'))));
                    }
                    UnifiedNodeId::SlideShape { anchor, .. } if lines[i] == *anchor => {
                        let start = i + 1;
                        let mut end = start;
                        while end < lines.len()
                            && !lines[end].starts_with("[shape ")
                            && !lines[end].starts_with("[title]")
                            && !lines[end].starts_with("[subtitle]")
                            && !lines[end].starts_with("[notes]")
                            && !lines[end].starts_with("--- slide ")
                            && !lines[end].starts_with("[table]")
                        {
                            if lines[end].is_empty() {
                                break;
                            }
                            end += 1;
                        }
                        let replacement: Vec<String> = if text.is_empty() {
                            vec![String::new()]
                        } else {
                            text.lines().map(ToString::to_string).collect()
                        };
                        lines.splice(start..end, replacement);
                        return Ok(Some(join_lines(lines.as_slice(), body.ends_with('\n'))));
                    }
                    UnifiedNodeId::PptxTableCell {
                        table, row, col, ..
                    } => {
                        if !caps.table_cells {
                            return Ok(None);
                        }
                        if lines[i].starts_with("[table]") {
                            let edited = edit_pptx_table_cell_for_slide(
                                lines,
                                *slide,
                                *table,
                                *row,
                                *col,
                                text,
                                body.ends_with('\n'),
                            )?;
                            return Ok(edited);
                        }
                    }
                    _ => {}
                }

                i += 1;
            }

            Ok(None)
        }
        _ => Ok(None),
    }
}

fn edit_docx_table_cell(
    mut lines: Vec<String>,
    table_idx: usize,
    row_idx: usize,
    col_idx: usize,
    text: &str,
    trailing_newline: bool,
) -> Result<Option<String>> {
    let mut i = 0_usize;
    while i < lines.len() {
        let is_target_header = lines[i] == format!("[t{table_idx}]");
        if !is_target_header {
            i += 1;
            continue;
        }

        let mut row_number = 0_usize;
        let mut j = i + 1;
        while j < lines.len() && lines[j].starts_with('|') {
            if is_markdown_separator_row(lines[j].as_str()) {
                j += 1;
                continue;
            }
            if row_number == row_idx {
                let mut cells = parse_markdown_cells(lines[j].as_str());
                if col_idx >= cells.len() {
                    return Ok(None);
                }
                cells[col_idx] = text.to_string();
                lines[j] = format_markdown_cells(cells.as_slice());
                return Ok(Some(join_lines(lines.as_slice(), trailing_newline)));
            }
            row_number += 1;
            j += 1;
        }

        return Ok(None);
    }

    Ok(None)
}

fn edit_pptx_table_cell_for_slide(
    mut lines: Vec<String>,
    target_slide: usize,
    target_table: usize,
    row_idx: usize,
    col_idx: usize,
    text: &str,
    trailing_newline: bool,
) -> Result<Option<String>> {
    let mut i = 0_usize;
    let mut current_slide: Option<usize> = None;
    let mut table_counter = 0_usize;

    while i < lines.len() {
        if let Some(s) = parse_slide_header(lines[i].as_str()) {
            current_slide = Some(s);
            table_counter = 0;
            i += 1;
            continue;
        }

        if current_slide == Some(target_slide) && lines[i].starts_with("[table]") {
            table_counter += 1;
            if table_counter == target_table {
                let mut row_number = 0_usize;
                let mut j = i + 1;
                while j < lines.len() && lines[j].starts_with('|') {
                    if is_markdown_separator_row(lines[j].as_str()) {
                        j += 1;
                        continue;
                    }
                    if row_number == row_idx {
                        let mut cells = parse_markdown_cells(lines[j].as_str());
                        if col_idx >= cells.len() {
                            return Ok(None);
                        }
                        cells[col_idx] = text.to_string();
                        lines[j] = format_markdown_cells(cells.as_slice());
                        return Ok(Some(join_lines(lines.as_slice(), trailing_newline)));
                    }
                    row_number += 1;
                    j += 1;
                }
                return Ok(None);
            }
        }

        i += 1;
    }

    Ok(None)
}

fn list_xlsx_nodes(body: &str) -> Result<Vec<UnifiedNode>> {
    let mut nodes = Vec::new();
    let mut current_sheet: Option<String> = None;

    for line in body.lines() {
        if let Some(name) = parse_sheet_header(line) {
            current_sheet = Some(name.to_string());
            continue;
        }

        let Some((left, right)) = line.split_once(": ") else {
            continue;
        };
        if !looks_like_cell_reference(left) {
            continue;
        }
        let sheet = current_sheet.as_deref().ok_or_else(|| {
            IrError::InvalidBody("cell line found before sheet header".to_string())
        })?;
        nodes.push(UnifiedNode {
            id: UnifiedNodeId::SpreadsheetCell {
                sheet: sheet.to_string(),
                cell: left.to_string(),
            },
            kind: UnifiedNodeKind::SpreadsheetCell,
            text: right.to_string(),
        });
    }

    Ok(nodes)
}

fn list_docx_nodes(body: &str) -> Result<Vec<UnifiedNode>> {
    let mut nodes = Vec::new();
    let lines: Vec<&str> = body.lines().collect();

    for line in &lines {
        if let Some((index, text)) = parse_docx_paragraph_line(line) {
            nodes.push(UnifiedNode {
                id: UnifiedNodeId::Paragraph { index },
                kind: UnifiedNodeKind::Paragraph,
                text: text.to_string(),
            });
        }
    }

    let mut i = 0_usize;
    while i < lines.len() {
        if let Some(table_idx) = parse_docx_table_header(lines[i]) {
            let mut row_number = 0_usize;
            let mut j = i + 1;
            while j < lines.len() && lines[j].starts_with('|') {
                if !is_markdown_separator_row(lines[j]) {
                    let cells = parse_markdown_cells(lines[j]);
                    for (col_number, cell) in cells.iter().enumerate() {
                        nodes.push(UnifiedNode {
                            id: UnifiedNodeId::DocxTableCell {
                                table: table_idx,
                                row: row_number,
                                col: col_number,
                            },
                            kind: UnifiedNodeKind::DocxTableCell,
                            text: cell.to_string(),
                        });
                    }
                    row_number += 1;
                }
                j += 1;
            }
            i = j;
            continue;
        }
        i += 1;
    }

    Ok(nodes)
}

fn list_pptx_nodes(body: &str) -> Result<Vec<UnifiedNode>> {
    let mut nodes = Vec::new();
    let mut current_slide: Option<usize> = None;
    let mut table_counter = 0_usize;
    let lines: Vec<&str> = body.lines().collect();
    let mut i = 0_usize;

    while i < lines.len() {
        let line = lines[i];
        if let Some(slide) = parse_slide_header(line) {
            current_slide = Some(slide);
            table_counter = 0;
            i += 1;
            continue;
        }

        let Some(slide) = current_slide else {
            i += 1;
            continue;
        };

        if let Some(text) = line.strip_prefix("[title] ") {
            nodes.push(UnifiedNode {
                id: UnifiedNodeId::SlideTitle { slide },
                kind: UnifiedNodeKind::SlideTitle,
                text: text.to_string(),
            });
            i += 1;
            continue;
        }
        if let Some(text) = line.strip_prefix("[subtitle] ") {
            nodes.push(UnifiedNode {
                id: UnifiedNodeId::SlideSubtitle { slide },
                kind: UnifiedNodeKind::SlideSubtitle,
                text: text.to_string(),
            });
            i += 1;
            continue;
        }
        if let Some(text) = line.strip_prefix("[notes] ") {
            nodes.push(UnifiedNode {
                id: UnifiedNodeId::SlideNotes { slide },
                kind: UnifiedNodeKind::SlideNotes,
                text: text.to_string(),
            });
            i += 1;
            continue;
        }

        if line.starts_with("[shape ") {
            let anchor = line.trim().to_string();
            let mut block_lines = Vec::new();
            i += 1;
            while i < lines.len()
                && !lines[i].starts_with("[shape ")
                && !lines[i].starts_with("[title]")
                && !lines[i].starts_with("[subtitle]")
                && !lines[i].starts_with("[notes]")
                && !lines[i].starts_with("--- slide ")
                && !lines[i].starts_with("[table]")
            {
                if lines[i].is_empty() {
                    break;
                }
                block_lines.push(lines[i].to_string());
                i += 1;
            }
            nodes.push(UnifiedNode {
                id: UnifiedNodeId::SlideShape { slide, anchor },
                kind: UnifiedNodeKind::SlideShape,
                text: block_lines.join("\n"),
            });
            continue;
        }

        if line.starts_with("[table]") {
            table_counter += 1;
            let table_idx = table_counter;
            let mut row_number = 0_usize;
            i += 1;
            while i < lines.len() && lines[i].starts_with('|') {
                if !is_markdown_separator_row(lines[i]) {
                    let cells = parse_markdown_cells(lines[i]);
                    for (col_number, cell) in cells.iter().enumerate() {
                        nodes.push(UnifiedNode {
                            id: UnifiedNodeId::PptxTableCell {
                                slide,
                                table: table_idx,
                                row: row_number,
                                col: col_number,
                            },
                            kind: UnifiedNodeKind::PptxTableCell,
                            text: cell.to_string(),
                        });
                    }
                    row_number += 1;
                }
                i += 1;
            }
            continue;
        }

        i += 1;
    }

    Ok(nodes)
}

fn split_full_body(body: &str) -> (&str, &str) {
    if let Some(pos) = body.find(FULL_MODE_SEPARATOR) {
        let content = &body[..pos];
        let style = &body[pos + FULL_MODE_SEPARATOR.len()..];
        (content, style)
    } else {
        (body, "")
    }
}

fn parse_unified_node_id(s: &str) -> Result<UnifiedNodeId> {
    if let Some(rest) = s.strip_prefix("sheet:") {
        let (sheet, tail) = rest
            .split_once('/')
            .ok_or_else(|| IrError::InvalidBody(format!("invalid xlsx id: {s}")))?;
        if sheet.is_empty() {
            return Err(IrError::InvalidBody(format!("invalid xlsx id: {s}")));
        }

        if let Some(cell) = tail
            .strip_prefix("cell:")
            .and_then(|v| v.strip_suffix("/style"))
        {
            if looks_like_cell_reference(cell) {
                return Ok(UnifiedNodeId::XlsxCellStyle {
                    sheet: sheet.to_string(),
                    cell: cell.to_string(),
                });
            }
        }

        if let Some(cell_part) = tail.strip_prefix("cell:") {
            if looks_like_cell_reference(cell_part) {
                return Ok(UnifiedNodeId::SpreadsheetCell {
                    sheet: sheet.to_string(),
                    cell: cell_part.to_string(),
                });
            }
        }

        if let Some(table_tail) = tail.strip_prefix("table:") {
            if let Some((table, cell_part)) = table_tail.split_once("/cell:") {
                if !table.is_empty() && looks_like_cell_reference(cell_part) {
                    return Ok(UnifiedNodeId::XlsxTableCell {
                        sheet: sheet.to_string(),
                        table: table.to_string(),
                        cell: cell_part.to_string(),
                    });
                }
            }
        }

        if let Some(chart_tail) = tail.strip_prefix("chart:") {
            let (chart_text, after_chart) = chart_tail
                .split_once('/')
                .ok_or_else(|| IrError::InvalidBody(format!("invalid xlsx chart id: {s}")))?;
            let chart = chart_text
                .parse::<usize>()
                .map_err(|_| IrError::InvalidBody(format!("invalid xlsx chart id: {s}")))?;
            if after_chart == "title" {
                return Ok(UnifiedNodeId::XlsxChartTitle {
                    sheet: sheet.to_string(),
                    chart,
                });
            }
            if let Some(series_tail) = after_chart.strip_prefix("series:") {
                let (series_text, suffix) = series_tail.split_once('/').ok_or_else(|| {
                    IrError::InvalidBody(format!("invalid xlsx chart series id: {s}"))
                })?;
                if suffix != "name" {
                    return Err(IrError::InvalidBody(format!(
                        "invalid xlsx chart series id: {s}"
                    )));
                }
                let series = series_text.parse::<usize>().map_err(|_| {
                    IrError::InvalidBody(format!("invalid xlsx chart series id: {s}"))
                })?;
                return Ok(UnifiedNodeId::XlsxChartSeriesName {
                    sheet: sheet.to_string(),
                    chart,
                    series,
                });
            }
        }
    }

    if let Some(rest) = s.strip_prefix("paragraph:") {
        if let Some(idx_text) = rest.strip_suffix("/style") {
            let index = idx_text
                .parse::<usize>()
                .map_err(|_| IrError::InvalidBody(format!("invalid paragraph style id: {s}")))?;
            return Ok(UnifiedNodeId::DocxParagraphStyle { index });
        }
        let index = rest
            .parse::<usize>()
            .map_err(|_| IrError::InvalidBody(format!("invalid paragraph id: {s}")))?;
        return Ok(UnifiedNodeId::Paragraph { index });
    }

    if let Some(rest) = s.strip_prefix("docx_table:") {
        if let Some((table_text, cell_part)) = rest.split_once("/cell:") {
            let table = table_text
                .parse::<usize>()
                .map_err(|_| IrError::InvalidBody(format!("invalid docx table id: {s}")))?;
            let (row, col) = parse_row_col(cell_part, s)?;
            return Ok(UnifiedNodeId::DocxTableCell { table, row, col });
        }
    }

    if let Some(rest) = s.strip_prefix("slide:") {
        let (slide_text, tail) = rest
            .split_once('/')
            .ok_or_else(|| IrError::InvalidBody(format!("invalid slide id: {s}")))?;
        let slide = slide_text
            .parse::<usize>()
            .map_err(|_| IrError::InvalidBody(format!("invalid slide id: {s}")))?;

        if tail == "title" {
            return Ok(UnifiedNodeId::SlideTitle { slide });
        }
        if tail == "subtitle" {
            return Ok(UnifiedNodeId::SlideSubtitle { slide });
        }
        if tail == "notes" {
            return Ok(UnifiedNodeId::SlideNotes { slide });
        }
        if let Some(anchor) = tail
            .strip_prefix("shape:")
            .and_then(|value| value.strip_suffix("/style"))
        {
            return Ok(UnifiedNodeId::PptxShapeStyle {
                slide,
                anchor: anchor.to_string(),
            });
        }
        if let Some(anchor) = tail.strip_prefix("shape:") {
            return Ok(UnifiedNodeId::SlideShape {
                slide,
                anchor: anchor.to_string(),
            });
        }
        if let Some(table_rest) = tail.strip_prefix("table:") {
            let (table_text, cell_part) = table_rest
                .split_once("/cell:")
                .ok_or_else(|| IrError::InvalidBody(format!("invalid pptx table id: {s}")))?;
            let table = table_text
                .parse::<usize>()
                .map_err(|_| IrError::InvalidBody(format!("invalid pptx table id: {s}")))?;
            let (row, col) = parse_row_col(cell_part, s)?;
            return Ok(UnifiedNodeId::PptxTableCell {
                slide,
                table,
                row,
                col,
            });
        }
        if let Some(chart_rest) = tail.strip_prefix("chart:") {
            let (chart_text, after_chart) = chart_rest
                .split_once('/')
                .ok_or_else(|| IrError::InvalidBody(format!("invalid pptx chart id: {s}")))?;
            let chart = chart_text
                .parse::<usize>()
                .map_err(|_| IrError::InvalidBody(format!("invalid pptx chart id: {s}")))?;
            if after_chart == "title" {
                return Ok(UnifiedNodeId::PptxChartTitle { slide, chart });
            }
            if let Some(series_tail) = after_chart.strip_prefix("series:") {
                let (series_text, suffix) = series_tail.split_once('/').ok_or_else(|| {
                    IrError::InvalidBody(format!("invalid pptx chart series id: {s}"))
                })?;
                if suffix != "name" {
                    return Err(IrError::InvalidBody(format!(
                        "invalid pptx chart series id: {s}"
                    )));
                }
                let series = series_text.parse::<usize>().map_err(|_| {
                    IrError::InvalidBody(format!("invalid pptx chart series id: {s}"))
                })?;
                return Ok(UnifiedNodeId::PptxChartSeriesName {
                    slide,
                    chart,
                    series,
                });
            }
        }
    }

    Err(IrError::InvalidBody(format!(
        "invalid unified node id: {s}"
    )))
}

fn parse_row_col(cell_part: &str, full: &str) -> Result<(usize, usize)> {
    let (row_text, col_text) = cell_part
        .split_once(',')
        .ok_or_else(|| IrError::InvalidBody(format!("invalid row,col cell id: {full}")))?;
    let row = row_text
        .parse::<usize>()
        .map_err(|_| IrError::InvalidBody(format!("invalid row in id: {full}")))?;
    let col = col_text
        .parse::<usize>()
        .map_err(|_| IrError::InvalidBody(format!("invalid col in id: {full}")))?;
    Ok((row, col))
}

fn parse_sheet_header(line: &str) -> Option<&str> {
    line.strip_prefix("=== Sheet: ")?.strip_suffix(" ===")
}

fn parse_docx_paragraph_line(line: &str) -> Option<(usize, &str)> {
    let rest = line.strip_prefix("[p")?;
    let (idx, text) = rest.split_once("] ")?;
    let index = idx.parse::<usize>().ok()?;
    Some((index, text))
}

fn parse_docx_table_header(line: &str) -> Option<usize> {
    let rest = line.strip_prefix("[t")?;
    let idx_text = rest.strip_suffix(']')?;
    idx_text.parse::<usize>().ok()
}

fn parse_slide_header(line: &str) -> Option<usize> {
    let rest = line.strip_prefix("--- slide ")?;
    let num_text: String = rest.chars().take_while(|ch| ch.is_ascii_digit()).collect();
    if num_text.is_empty() {
        return None;
    }
    num_text.parse::<usize>().ok()
}

fn looks_like_cell_reference(s: &str) -> bool {
    let mut seen_letters = false;
    let mut seen_digits = false;

    for ch in s.chars() {
        if ch.is_ascii_uppercase() && !seen_digits {
            seen_letters = true;
            continue;
        }
        if ch.is_ascii_digit() && seen_letters {
            seen_digits = true;
            continue;
        }
        return false;
    }

    seen_letters && seen_digits
}

fn parse_markdown_cells(line: &str) -> Vec<String> {
    let trimmed = line.trim();
    let core = trimmed.trim_start_matches('|').trim_end_matches('|');
    core.split('|')
        .map(|cell| cell.trim().to_string())
        .collect()
}

fn format_markdown_cells(cells: &[String]) -> String {
    let mut out = String::new();
    out.push('|');
    for cell in cells {
        out.push(' ');
        out.push_str(cell.as_str());
        out.push_str(" |");
    }
    out
}

fn is_markdown_separator_row(line: &str) -> bool {
    let cells = parse_markdown_cells(line);
    if cells.is_empty() {
        return false;
    }
    cells
        .iter()
        .all(|cell| !cell.is_empty() && cell.chars().all(|ch| ch == '-' || ch == ':'))
}

fn join_lines(lines: &[String], trailing_newline: bool) -> String {
    let mut out = lines.join("\n");
    if trailing_newline {
        out.push('\n');
    }
    out
}

fn parse_cell_style_patch(text: &str) -> Result<CellStylePatch> {
    let mut patch = CellStylePatch {
        bold: None,
        italic: None,
        number_format: None,
    };

    for raw in text.split(';') {
        let part = raw.trim();
        if part.is_empty() {
            continue;
        }
        let Some((key, value)) = part.split_once('=') else {
            continue;
        };
        let key = key.trim();
        let value = value.trim();
        match key {
            "bold" => match value {
                "true" => patch.bold = Some(true),
                "false" => patch.bold = Some(false),
                _ => {
                    return Err(IrError::InvalidBody(format!(
                        "invalid style patch boolean for bold: {value}"
                    )))
                }
            },
            "italic" => match value {
                "true" => patch.italic = Some(true),
                "false" => patch.italic = Some(false),
                _ => {
                    return Err(IrError::InvalidBody(format!(
                        "invalid style patch boolean for italic: {value}"
                    )))
                }
            },
            "number_format" => patch.number_format = Some(value.to_string()),
            _ => {}
        }
    }

    Ok(patch)
}

fn format_cell_style_patch(style: Option<&offidized_xlsx::Style>) -> Option<String> {
    let style = style?;
    let mut parts = Vec::new();
    if let Some(font) = style.font() {
        if let Some(bold) = font.bold() {
            parts.push(format!("bold={bold}"));
        }
        if let Some(italic) = font.italic() {
            parts.push(format!("italic={italic}"));
        }
    }
    if let Some(fmt) = style.custom_format().or_else(|| style.number_format()) {
        if !fmt.is_empty() {
            parts.push(format!("number_format={fmt}"));
        }
    }
    if parts.is_empty() {
        None
    } else {
        Some(parts.join(";"))
    }
}

fn parse_pptx_text_style_patch(text: &str) -> Result<PptxTextStylePatch> {
    let mut patch = PptxTextStylePatch::new();
    for raw in text.split(';') {
        let part = raw.trim();
        if part.is_empty() {
            continue;
        }
        let Some((key, value)) = part.split_once('=') else {
            continue;
        };
        let key = key.trim();
        let value = value.trim();
        match key {
            "bold" => match value {
                "true" => patch.set_bold(true),
                "false" => patch.set_bold(false),
                _ => {
                    return Err(IrError::InvalidBody(format!(
                        "invalid style patch boolean for bold: {value}"
                    )))
                }
            },
            "italic" => match value {
                "true" => patch.set_italic(true),
                "false" => patch.set_italic(false),
                _ => {
                    return Err(IrError::InvalidBody(format!(
                        "invalid style patch boolean for italic: {value}"
                    )))
                }
            },
            "font_size" => {
                let font_size = value.parse::<u32>().map_err(|_| {
                    IrError::InvalidBody(format!("invalid style patch font_size: {value}"))
                })?;
                patch.set_font_size(font_size);
            }
            "font_color" => patch.set_font_color(value),
            "font_name" => patch.set_font_name(value),
            _ => {}
        }
    }
    Ok(patch)
}

fn format_pptx_shape_style_patch(shape: &offidized_pptx::Shape) -> Option<String> {
    let paragraph = shape.paragraphs().first()?;
    let run = paragraph.runs().first()?;
    let mut parts = Vec::new();
    if run.properties().bold.is_some() {
        parts.push(format!("bold={}", run.is_bold()));
    }
    if run.properties().italic.is_some() {
        parts.push(format!("italic={}", run.is_italic()));
    }
    if let Some(font_size) = run.font_size() {
        parts.push(format!("font_size={font_size}"));
    }
    if let Some(font_color) = run.font_color() {
        parts.push(format!("font_color={font_color}"));
    }
    if let Some(font_name) = run.font_name() {
        parts.push(format!("font_name={font_name}"));
    }
    if parts.is_empty() {
        None
    } else {
        Some(parts.join(";"))
    }
}

fn parse_shape_name_from_anchor(anchor: &str) -> Option<String> {
    let body = anchor
        .strip_prefix("[shape \"")
        .and_then(|rest| rest.strip_suffix("\"]"))?;
    Some(body.to_string())
}

fn augment_xlsx_nodes(
    path: &Path,
    nodes: &mut Vec<UnifiedNode>,
    capabilities: &mut UnifiedCapabilities,
) -> Result<()> {
    let wb = offidized_xlsx::Workbook::open(path)?;
    capabilities.table_cells = true;
    capabilities.chart_meta = true;
    capabilities.style_nodes = true;

    for ws in wb.worksheets() {
        for table in ws.tables() {
            let range = table.range();
            let (start_col, start_row) = parse_cell_reference(range.start())?;
            let (end_col, end_row) = parse_cell_reference(range.end())?;
            for row in start_row..=end_row {
                for col in start_col..=end_col {
                    let cell_ref = build_cell_reference(col, row)?;
                    let text = ws
                        .cell(cell_ref.as_str())
                        .and_then(|cell| {
                            cell.formula()
                                .map(|f| format!("={f}"))
                                .or_else(|| cell.value().map(format_cell_value_for_node))
                        })
                        .unwrap_or_default();
                    nodes.push(UnifiedNode {
                        id: UnifiedNodeId::XlsxTableCell {
                            sheet: ws.name().to_string(),
                            table: table.name().to_string(),
                            cell: cell_ref,
                        },
                        kind: UnifiedNodeKind::XlsxTableCell,
                        text,
                    });
                }
            }
        }

        for (chart_idx, chart) in ws.charts().iter().enumerate() {
            let chart_num = chart_idx + 1;
            nodes.push(UnifiedNode {
                id: UnifiedNodeId::XlsxChartTitle {
                    sheet: ws.name().to_string(),
                    chart: chart_num,
                },
                kind: UnifiedNodeKind::XlsxChartTitle,
                text: chart.title().unwrap_or_default().to_string(),
            });

            for (series_idx, series) in chart.series().iter().enumerate() {
                nodes.push(UnifiedNode {
                    id: UnifiedNodeId::XlsxChartSeriesName {
                        sheet: ws.name().to_string(),
                        chart: chart_num,
                        series: series_idx + 1,
                    },
                    kind: UnifiedNodeKind::XlsxChartSeriesName,
                    text: series.name().unwrap_or_default().to_string(),
                });
            }
        }

        for (cell_ref, cell) in ws.cells() {
            let Some(style_id) = cell.style_id() else {
                continue;
            };
            if style_id == 0 {
                continue;
            }
            let Some(style_text) = format_cell_style_patch(wb.style(style_id)) else {
                continue;
            };
            nodes.push(UnifiedNode {
                id: UnifiedNodeId::XlsxCellStyle {
                    sheet: ws.name().to_string(),
                    cell: cell_ref.to_string(),
                },
                kind: UnifiedNodeKind::XlsxCellStyle,
                text: style_text,
            });
        }
    }

    Ok(())
}

fn augment_docx_nodes(
    path: &Path,
    nodes: &mut Vec<UnifiedNode>,
    capabilities: &mut UnifiedCapabilities,
) -> Result<()> {
    let doc = offidized_docx::Document::open(path)?;
    capabilities.style_nodes = true;

    for (idx, paragraph) in doc.paragraphs().iter().enumerate() {
        let Some(style_id) = paragraph.style_id() else {
            continue;
        };
        nodes.push(UnifiedNode {
            id: UnifiedNodeId::DocxParagraphStyle { index: idx + 1 },
            kind: UnifiedNodeKind::DocxParagraphStyle,
            text: style_id.to_string(),
        });
    }

    Ok(())
}

fn augment_pptx_nodes(
    path: &Path,
    nodes: &mut Vec<UnifiedNode>,
    capabilities: &mut UnifiedCapabilities,
) -> Result<()> {
    let prs = offidized_pptx::Presentation::open(path)?;
    capabilities.chart_meta = true;
    capabilities.style_nodes = true;

    for (slide_idx, slide) in prs.slides().iter().enumerate() {
        let slide_num = slide_idx + 1;
        for (chart_idx, chart) in slide.charts().iter().enumerate() {
            let chart_num = chart_idx + 1;
            nodes.push(UnifiedNode {
                id: UnifiedNodeId::PptxChartTitle {
                    slide: slide_num,
                    chart: chart_num,
                },
                kind: UnifiedNodeKind::PptxChartTitle,
                text: chart.title().to_string(),
            });
            for (series_idx, series) in chart.additional_series().iter().enumerate() {
                nodes.push(UnifiedNode {
                    id: UnifiedNodeId::PptxChartSeriesName {
                        slide: slide_num,
                        chart: chart_num,
                        series: series_idx + 1,
                    },
                    kind: UnifiedNodeKind::PptxChartSeriesName,
                    text: series.name().to_string(),
                });
            }
        }

        for shape in slide.shapes() {
            let Some(style_text) = format_pptx_shape_style_patch(shape) else {
                continue;
            };
            let anchor = format!("[shape \"{}\"]", shape.name());
            nodes.push(UnifiedNode {
                id: UnifiedNodeId::PptxShapeStyle {
                    slide: slide_num,
                    anchor,
                },
                kind: UnifiedNodeKind::PptxShapeStyle,
                text: style_text,
            });
        }
    }

    Ok(())
}

fn apply_direct_mutations(path: &Path, format: Format, mutations: &[DirectMutation]) -> Result<()> {
    match format {
        Format::Xlsx => apply_direct_mutations_xlsx(path, mutations),
        Format::Pptx => apply_direct_mutations_pptx(path, mutations),
        Format::Docx => apply_direct_mutations_docx(path, mutations),
    }
}

fn apply_direct_mutations_xlsx(path: &Path, mutations: &[DirectMutation]) -> Result<()> {
    let mut wb = offidized_xlsx::Workbook::open(path)?;
    for mutation in mutations {
        match mutation {
            DirectMutation::XlsxChartTitle {
                sheet,
                chart,
                title,
            } => {
                let Some(ws) = wb.sheet_mut(sheet.as_str()) else {
                    continue;
                };
                let chart_idx = chart.saturating_sub(1);
                if let Some(target) = ws.charts_mut().get_mut(chart_idx) {
                    target.set_title(title.as_str());
                }
            }
            DirectMutation::XlsxChartSeriesName {
                sheet,
                chart,
                series,
                name,
            } => {
                let Some(ws) = wb.sheet_mut(sheet.as_str()) else {
                    continue;
                };
                let chart_idx = chart.saturating_sub(1);
                if let Some(target) = ws.charts_mut().get_mut(chart_idx) {
                    let series_idx = series.saturating_sub(1);
                    if let Some(ser) = target.series_mut().get_mut(series_idx) {
                        ser.set_name(name.as_str());
                    }
                }
            }
            DirectMutation::XlsxCellStyle { sheet, cell, patch } => {
                let mut style = offidized_xlsx::Style::new();
                if let Some(number_format) = &patch.number_format {
                    style.set_custom_format(number_format.to_string());
                }
                let mut font = offidized_xlsx::Font::new();
                let mut has_font = false;
                if let Some(bold) = patch.bold {
                    font.set_bold(bold);
                    has_font = true;
                }
                if let Some(italic) = patch.italic {
                    font.set_italic(italic);
                    has_font = true;
                }
                if has_font {
                    style.set_font(font);
                }
                let style_id = wb.add_style(style)?;
                let Some(ws) = wb.sheet_mut(sheet.as_str()) else {
                    continue;
                };
                ws.cell_mut(cell.as_str())?.set_style_id(style_id);
            }
            _ => {}
        }
    }
    wb.save(path)?;
    Ok(())
}

fn apply_direct_mutations_pptx(path: &Path, mutations: &[DirectMutation]) -> Result<()> {
    let mut prs = offidized_pptx::Presentation::open(path)?;
    for mutation in mutations {
        match mutation {
            DirectMutation::PptxChartTitle {
                slide,
                chart,
                title,
            } => {
                let slide_idx = slide.saturating_sub(1);
                let chart_idx = chart.saturating_sub(1);
                if let Some(slide) = prs.slide_mut(slide_idx) {
                    if let Some(target) = slide.charts_mut().get_mut(chart_idx) {
                        target.set_title(title.as_str());
                    }
                }
            }
            DirectMutation::PptxChartSeriesName {
                slide,
                chart,
                series,
                name,
            } => {
                let slide_idx = slide.saturating_sub(1);
                let chart_idx = chart.saturating_sub(1);
                if let Some(slide) = prs.slide_mut(slide_idx) {
                    if let Some(target) = slide.charts_mut().get_mut(chart_idx) {
                        let total = target.additional_series().len();
                        if *series > 0 && *series <= total {
                            let mut series_list = Vec::with_capacity(total);
                            for _ in 0..total {
                                if let Some(item) = target.remove_series(0) {
                                    series_list.push(item);
                                }
                            }
                            let idx = series - 1;
                            if let Some(item) = series_list.get_mut(idx) {
                                item.set_name(name.as_str());
                            }
                            for item in series_list {
                                target.add_series(item);
                            }
                        }
                    }
                }
            }
            DirectMutation::PptxShapeStyle {
                slide,
                shape_name,
                patch,
            } => {
                let slide_idx = slide.saturating_sub(1);
                let Some(slide) = prs.slide_mut(slide_idx) else {
                    continue;
                };
                let Some(shape) = slide
                    .shapes_mut()
                    .iter_mut()
                    .find(|s| s.name() == shape_name)
                else {
                    continue;
                };
                if shape.paragraphs().is_empty() {
                    shape.add_paragraph();
                }
                if shape
                    .paragraphs()
                    .first()
                    .is_some_and(|p| p.runs().is_empty())
                {
                    shape.paragraphs_mut()[0].add_run("");
                }
                let run = &mut shape.paragraphs_mut()[0].runs_mut()[0];
                if let Some(value) = patch.bold {
                    run.set_bold(value);
                }
                if let Some(value) = patch.italic {
                    run.set_italic(value);
                }
                if let Some(value) = patch.font_size {
                    run.set_font_size(value);
                }
                if let Some(value) = &patch.font_color {
                    run.set_font_color(value.as_str());
                }
                if let Some(value) = &patch.font_name {
                    run.set_font_name(value.as_str());
                }
            }
            _ => {}
        }
    }
    prs.save(path)?;
    Ok(())
}

fn apply_direct_mutations_docx(path: &Path, mutations: &[DirectMutation]) -> Result<()> {
    let mut doc = offidized_docx::Document::open(path)?;
    for mutation in mutations {
        if let DirectMutation::DocxParagraphStyle { index, style_id } = mutation {
            let paragraph_idx = index.saturating_sub(1);
            let Some(paragraph) = doc.paragraphs_mut().get_mut(paragraph_idx) else {
                continue;
            };
            if let Some(style_id) = style_id {
                paragraph.set_style_id(style_id.as_str());
            } else {
                paragraph.clear_style_id();
            }
        }
    }
    doc.save(path)?;
    Ok(())
}

fn parse_cell_reference(reference: &str) -> Result<(u32, u32)> {
    let trimmed = reference.trim();
    if trimmed.is_empty() {
        return Err(IrError::InvalidBody(format!(
            "invalid cell reference: {reference}"
        )));
    }
    let split_index = trimmed
        .char_indices()
        .find_map(|(index, ch)| ch.is_ascii_digit().then_some(index))
        .ok_or_else(|| IrError::InvalidBody(format!("invalid cell reference: {reference}")))?;
    let (column_text, row_text) = trimmed.split_at(split_index);
    if column_text.is_empty()
        || row_text.is_empty()
        || !column_text.chars().all(|ch| ch.is_ascii_alphabetic())
        || !row_text.chars().all(|ch| ch.is_ascii_digit())
        || row_text.starts_with('0')
    {
        return Err(IrError::InvalidBody(format!(
            "invalid cell reference: {reference}"
        )));
    }
    let normalized_column: String = column_text
        .chars()
        .map(|ch| ch.to_ascii_uppercase())
        .collect();
    let normalized = format!("{normalized_column}{row_text}");
    let split_index = normalized
        .char_indices()
        .find_map(|(index, ch)| ch.is_ascii_digit().then_some(index))
        .ok_or_else(|| IrError::InvalidBody(format!("invalid cell reference: {reference}")))?;
    let (column_name, row_text) = normalized.split_at(split_index);
    let col = column_name
        .bytes()
        .try_fold(0_u32, |acc, byte| {
            acc.checked_mul(26)
                .and_then(|value| value.checked_add(u32::from(byte - b'A' + 1)))
        })
        .ok_or_else(|| IrError::InvalidBody(format!("invalid cell reference: {reference}")))?;
    let row = row_text
        .parse::<u32>()
        .map_err(|_| IrError::InvalidBody(format!("invalid cell reference: {reference}")))?;
    Ok((col, row))
}

fn build_cell_reference(column: u32, row: u32) -> Result<String> {
    if column == 0 || row == 0 {
        return Err(IrError::InvalidBody("invalid row/column index".to_string()));
    }
    let mut current = column;
    let mut label = String::new();
    while current > 0 {
        let remainder = (current - 1) % 26;
        label.push(char::from(b'A' + remainder as u8));
        current = (current - 1) / 26;
    }
    let column_label: String = label.chars().rev().collect();
    Ok(format!("{column_label}{row}"))
}

fn format_cell_value_for_node(value: &offidized_xlsx::CellValue) -> String {
    match value {
        offidized_xlsx::CellValue::Blank => String::new(),
        offidized_xlsx::CellValue::String(v) => v.clone(),
        offidized_xlsx::CellValue::Number(v) => v.to_string(),
        offidized_xlsx::CellValue::Bool(v) => {
            if *v {
                "true".to_string()
            } else {
                "false".to_string()
            }
        }
        offidized_xlsx::CellValue::Date(v) => v.clone(),
        offidized_xlsx::CellValue::Error(v) => v.clone(),
        offidized_xlsx::CellValue::DateTime(v) => v.to_string(),
        offidized_xlsx::CellValue::RichText(runs) => runs.iter().map(|r| r.text()).collect(),
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn wrap_ir(format: &str, body: &str) -> String {
        format!(
            "+++\nsource = \"test.{format}\"\nformat = \"{format}\"\nmode = \"content\"\nversion = 1\nchecksum = \"sha256:test\"\n+++\n{body}"
        )
    }

    #[test]
    fn id_roundtrip_display_parse() {
        let ids = vec![
            UnifiedNodeId::SpreadsheetCell {
                sheet: "Data".to_string(),
                cell: "B2".to_string(),
            },
            UnifiedNodeId::XlsxTableCell {
                sheet: "Data".to_string(),
                table: "t_fin".to_string(),
                cell: "C4".to_string(),
            },
            UnifiedNodeId::XlsxChartTitle {
                sheet: "Data".to_string(),
                chart: 2,
            },
            UnifiedNodeId::XlsxChartSeriesName {
                sheet: "Data".to_string(),
                chart: 2,
                series: 3,
            },
            UnifiedNodeId::XlsxCellStyle {
                sheet: "Data".to_string(),
                cell: "D7".to_string(),
            },
            UnifiedNodeId::Paragraph { index: 4 },
            UnifiedNodeId::DocxParagraphStyle { index: 4 },
            UnifiedNodeId::DocxTableCell {
                table: 1,
                row: 2,
                col: 3,
            },
            UnifiedNodeId::SlideTitle { slide: 1 },
            UnifiedNodeId::SlideShape {
                slide: 2,
                anchor: "[shape \"Body 2\"]".to_string(),
            },
            UnifiedNodeId::PptxShapeStyle {
                slide: 2,
                anchor: "[shape \"Body 2\"]".to_string(),
            },
            UnifiedNodeId::PptxTableCell {
                slide: 3,
                table: 1,
                row: 0,
                col: 1,
            },
            UnifiedNodeId::PptxChartTitle { slide: 3, chart: 1 },
            UnifiedNodeId::PptxChartSeriesName {
                slide: 3,
                chart: 1,
                series: 2,
            },
        ];

        for id in ids {
            let text = id.to_string();
            let reparsed = UnifiedNodeId::from_str(text.as_str()).expect("parse id");
            assert_eq!(id, reparsed);
        }
    }

    #[test]
    fn xlsx_nodes_and_edit_roundtrip() {
        let ir = wrap_ir("xlsx", "\n=== Sheet: Data ===\nA1: Region\nB2: 123\n");

        let nodes = list_nodes_from_ir(ir.as_str()).expect("list nodes");
        assert!(nodes.iter().any(|n| {
            n.id == UnifiedNodeId::SpreadsheetCell {
                sheet: "Data".to_string(),
                cell: "A1".to_string(),
            }
        }));

        let edited = apply_edits_to_ir(
            ir.as_str(),
            &[UnifiedEdit::new("sheet:Data/cell:B2", "456")],
        )
        .expect("edit ir");

        assert!(edited.contains("B2: 456"));
    }

    #[test]
    fn docx_nodes_include_table_cells_and_edit() {
        let ir = wrap_ir(
            "docx",
            "\n[p1] hello\n[t1]\n| H1 | H2 |\n|---|---|\n| a | b |\n",
        );
        let nodes = list_nodes_from_ir(ir.as_str()).expect("list nodes");
        assert!(nodes
            .iter()
            .any(|n| n.id == UnifiedNodeId::Paragraph { index: 1 }));
        assert!(nodes.iter().any(|n| {
            n.id == UnifiedNodeId::DocxTableCell {
                table: 1,
                row: 1,
                col: 1,
            }
        }));

        let edited = apply_edits_to_ir(
            ir.as_str(),
            &[UnifiedEdit::new("docx_table:1/cell:1,1", "updated")],
        )
        .expect("edit ir");
        assert!(edited.contains("| a | updated |"));
    }

    #[test]
    fn pptx_nodes_include_table_cells_and_edit() {
        let body = "\n--- slide 1 [Title and Content] ---\n[title] Q1 Review\n[shape \"Body 2\"]\nold line\n[table]\n| H1 | H2 |\n|---|---|\n| x | y |\n[notes] old note\n";
        let ir = wrap_ir("pptx", body);

        let nodes = list_nodes_from_ir(ir.as_str()).expect("list nodes");
        assert!(nodes
            .iter()
            .any(|n| n.id == UnifiedNodeId::SlideTitle { slide: 1 }));
        assert!(nodes.iter().any(|n| {
            n.id == UnifiedNodeId::PptxTableCell {
                slide: 1,
                table: 1,
                row: 1,
                col: 1,
            }
        }));

        let edited = apply_edits_to_ir(
            ir.as_str(),
            &[
                UnifiedEdit::new("slide:1/title", "Q2 Review"),
                UnifiedEdit::new("slide:1/shape:[shape \"Body 2\"]", "line A\nline B"),
                UnifiedEdit::new("slide:1/table:1/cell:1,1", "z"),
                UnifiedEdit::new("slide:1/notes", "updated note"),
            ],
        )
        .expect("edit ir");

        assert!(edited.contains("[title] Q2 Review"));
        assert!(edited.contains("line A\nline B"));
        assert!(edited.contains("| x | z |"));
        assert!(edited.contains("[notes] updated note"));
    }

    #[test]
    fn unified_document_returns_report_with_missing_target() {
        let ir = wrap_ir("xlsx", "\n=== Sheet: Data ===\nA1: 1\n");
        let mut doc = UnifiedDocument::from_ir(ir.as_str()).expect("from ir");
        let report = doc
            .apply_edits(&[UnifiedEdit::new("sheet:Data/cell:Z9", "2")])
            .expect("apply edits");

        assert_eq!(report.requested, 1);
        assert_eq!(report.applied, 0);
        assert_eq!(report.skipped, 1);
        assert_eq!(report.diagnostics.len(), 1);
        assert_eq!(report.diagnostics[0].code, "target_not_found");
    }

    #[test]
    fn grouped_edits_are_atomic() {
        let ir = wrap_ir("xlsx", "\n=== Sheet: Data ===\nA1: 1\n");
        let mut doc = UnifiedDocument::from_ir(ir.as_str()).expect("from ir");
        let edits = vec![
            UnifiedEdit::new("sheet:Data/cell:A1", "2").with_group("txn-1"),
            UnifiedEdit::new("sheet:Data/cell:Z9", "3").with_group("txn-1"),
        ];
        let report = doc.apply_edits(edits.as_slice()).expect("apply edits");

        assert_eq!(report.applied, 0);
        assert_eq!(report.skipped, 2);
        assert!(report
            .diagnostics
            .iter()
            .any(|diag| diag.code == "group_aborted"));
        assert!(doc.to_ir().contains("A1: 1"));
    }

    #[test]
    fn xlsx_roundtrip_keeps_pivot_parts_and_updates_chart_title() {
        use offidized_xlsx::{
            Chart, ChartAxis, ChartDataRef, ChartSeries, ChartType, PivotDataField, PivotField,
            PivotSourceReference, PivotSubtotalFunction, PivotTable, Workbook,
        };
        use std::io::Read;

        let dir = tempfile::tempdir().expect("tempdir");
        let source = dir.path().join("finance.xlsx");
        let output = dir.path().join("finance.updated.xlsx");

        let mut workbook = Workbook::new();
        let sheet = workbook.add_sheet("Sheet1");
        sheet.cell_mut("A1").expect("A1").set_value("Region");
        sheet.cell_mut("B1").expect("B1").set_value("Revenue");
        sheet.cell_mut("A2").expect("A2").set_value("North");
        sheet.cell_mut("B2").expect("B2").set_value(1000);

        let mut chart = Chart::new(ChartType::Bar);
        chart.set_title("Revenue");
        chart.set_anchor(3, 1, 10, 15);
        let mut series = ChartSeries::new(0, 0);
        series
            .set_name_ref("Sheet1!$B$1")
            .set_categories(ChartDataRef::from_formula("Sheet1!$A$2:$A$2"))
            .set_values(ChartDataRef::from_formula("Sheet1!$B$2:$B$2"));
        chart.add_series(series);
        chart.add_axis(ChartAxis::new_category());
        chart.add_axis(ChartAxis::new_value());
        sheet.charts_mut().push(chart);

        let pivot_sheet = workbook.add_sheet("Pivot");
        let source_ref = PivotSourceReference::from_range("Sheet1!$A$1:$B$2");
        let mut pivot = PivotTable::new("Pivot1", source_ref);
        pivot.set_target(0, 0);
        pivot.add_row_field(PivotField::new("Region"));
        let mut revenue = PivotDataField::new("Revenue");
        revenue.set_subtotal(PivotSubtotalFunction::Sum);
        pivot.add_data_field(revenue);
        pivot_sheet.add_pivot_table(pivot);

        workbook.save(&source).expect("save source");

        let mut doc =
            UnifiedDocument::derive(&source, UnifiedDeriveOptions::default()).expect("derive");
        let report = doc
            .apply_edits(&[UnifiedEdit::new(
                "sheet:Sheet1/chart:1/title",
                "Revenue Updated",
            )])
            .expect("apply edits");
        assert_eq!(report.applied, 1);
        doc.save_as(&output, &ApplyOptions::default())
            .expect("save updated");

        let loaded = Workbook::open(&output).expect("open updated workbook");
        let ws = loaded
            .sheet("Sheet1")
            .expect("sheet should exist after unified edit");
        assert_eq!(ws.charts()[0].title(), Some("Revenue Updated"));

        let file = std::fs::File::open(&output).expect("open zip");
        let mut archive = zip::ZipArchive::new(file).expect("read zip");
        let mut pivot_def = archive
            .by_name("xl/pivotCache/pivotCacheDefinition1.xml")
            .expect("pivot cache definition should remain");
        let mut xml = String::new();
        pivot_def.read_to_string(&mut xml).expect("read cache");
        assert!(xml.contains("Sheet1"));
    }

    #[test]
    fn docx_paragraph_style_node_roundtrip() {
        let dir = tempfile::tempdir().expect("tempdir");
        let source = dir.path().join("styled.docx");
        let output = dir.path().join("styled.updated.docx");

        let mut docx = offidized_docx::Document::new();
        docx.add_paragraph_with_style("Hello", "Heading1");
        docx.save(&source).expect("save source docx");

        let mut doc =
            UnifiedDocument::derive(&source, UnifiedDeriveOptions::default()).expect("derive");
        assert!(doc.nodes().iter().any(|node| {
            node.id == UnifiedNodeId::DocxParagraphStyle { index: 1 } && node.text == "Heading1"
        }));
        let report = doc
            .apply_edits(&[UnifiedEdit::new("paragraph:1/style", "Normal")])
            .expect("apply style edit");
        assert_eq!(report.applied, 1);
        doc.save_as(&output, &ApplyOptions::default())
            .expect("save output");

        let reopened = offidized_docx::Document::open(&output).expect("open output docx");
        assert_eq!(reopened.paragraphs()[0].style_id(), Some("Normal"));
    }

    #[test]
    fn pptx_shape_style_node_roundtrip() {
        let dir = tempfile::tempdir().expect("tempdir");
        let source = dir.path().join("styled.pptx");
        let output = dir.path().join("styled.updated.pptx");

        let mut prs = offidized_pptx::Presentation::new();
        let slide = prs.add_slide_with_title("Title");
        let shape = slide.add_shape("Body 2");
        let run = shape.add_paragraph().add_run("Line");
        run.set_bold(true);
        run.set_font_color("FF0000");
        prs.save(&source).expect("save source pptx");

        let mut doc =
            UnifiedDocument::derive(&source, UnifiedDeriveOptions::default()).expect("derive");
        assert!(doc.nodes().iter().any(|node| {
            node.id
                == UnifiedNodeId::PptxShapeStyle {
                    slide: 1,
                    anchor: "[shape \"Body 2\"]".to_string(),
                }
        }));

        let mut patch = PptxTextStylePatch::new();
        patch.set_italic(true);
        patch.set_font_name("Arial");
        let edit = UnifiedEdit::new("slide:1/shape:[shape \"Body 2\"]/style", "")
            .with_payload(UnifiedEditPayload::PptxTextStyle(patch));
        let report = doc.apply_edits(&[edit]).expect("apply style edit");
        assert_eq!(report.applied, 1);
        doc.save_as(&output, &ApplyOptions::default())
            .expect("save output");

        let reopened = offidized_pptx::Presentation::open(&output).expect("open output pptx");
        let shape = &reopened.slides()[0].shapes()[0];
        let run = &shape.paragraphs()[0].runs()[0];
        assert!(run.is_italic());
        assert_eq!(run.font_name(), Some("Arial"));
    }

    #[test]
    fn lint_reports_invalid_table_coordinates() {
        let ir = wrap_ir("docx", "\n[t1]\n| H1 | H2 |\n|---|---|\n| a | b |\n");
        let doc = UnifiedDocument::from_ir(ir.as_str()).expect("from ir");
        let diagnostics = doc.lint_edits(&[UnifiedEdit::new("docx_table:1/cell:9,9", "x")]);
        assert!(diagnostics
            .iter()
            .any(|diag| diag.code == "invalid_table_coordinates"));
    }
}
