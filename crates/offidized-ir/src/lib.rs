//! Bidirectional lossless text IR for Office files.
//!
//! The IR provides a text format that AI agents can read and edit as source code,
//! with lossless bidirectional sync to the binary Office file.
//!
//! Core loop: `derive` (binary → text) → agent edits text → `apply` (text → binary).
//! Everything untouched survives via offidized's roundtrip layer.
//!
//! # Example
//!
//! ```no_run
//! use offidized_ir::{derive, apply, DeriveOptions, ApplyOptions};
//! use std::path::Path;
//!
//! // Derive text IR from an xlsx file
//! let ir = derive(Path::new("report.xlsx"), DeriveOptions::default())?;
//! println!("{ir}");
//!
//! // Apply edited IR back to produce updated file
//! let result = apply(&ir, Path::new("updated.xlsx"), &ApplyOptions::default())?;
//! println!("Updated {} cells", result.cells_updated);
//! # Ok::<(), offidized_ir::IrError>(())
//! ```

mod docx;
mod header;
mod pptx;
mod unified_api;
mod xlsx;

use std::path::{Path, PathBuf};

use sha2::{Digest, Sha256};

pub use header::IrHeader;
pub use unified_api::{
    apply_edits_to_ir, derive_content_nodes, edit_file_content, list_nodes_from_ir, CellStylePatch,
    PptxTextStylePatch, UnifiedCapabilities, UnifiedDeriveOptions, UnifiedDiagnostic,
    UnifiedDiagnosticSeverity, UnifiedDocument, UnifiedEdit, UnifiedEditPayload, UnifiedEditReport,
    UnifiedNode, UnifiedNodeId, UnifiedNodeKind,
};

/// IR operation mode controlling how much detail is included.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum Mode {
    /// Cell values and formulas only. The 80% case.
    #[default]
    Content,
    /// Formatting and layout only (no content).
    Style,
    /// Content and style combined.
    Full,
}

impl Mode {
    /// Returns the string representation used in IR headers.
    pub fn as_str(self) -> &'static str {
        match self {
            Self::Content => "content",
            Self::Style => "style",
            Self::Full => "full",
        }
    }

    /// Parse a mode string from an IR header.
    pub fn parse_str(s: &str) -> Result<Self> {
        match s {
            "content" => Ok(Self::Content),
            "style" => Ok(Self::Style),
            "full" => Ok(Self::Full),
            _ => Err(IrError::InvalidHeader(format!("unknown mode: {s}"))),
        }
    }
}

/// Source file format.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Format {
    /// Excel spreadsheet (.xlsx).
    Xlsx,
    /// Word document (.docx).
    Docx,
    /// PowerPoint presentation (.pptx).
    Pptx,
}

impl Format {
    /// Returns the string representation used in IR headers.
    pub fn as_str(self) -> &'static str {
        match self {
            Self::Xlsx => "xlsx",
            Self::Docx => "docx",
            Self::Pptx => "pptx",
        }
    }

    /// Parse a format string from an IR header.
    pub fn parse_str(s: &str) -> Result<Self> {
        match s {
            "xlsx" => Ok(Self::Xlsx),
            "docx" => Ok(Self::Docx),
            "pptx" => Ok(Self::Pptx),
            _ => Err(IrError::InvalidHeader(format!("unknown format: {s}"))),
        }
    }

    /// Detect format from file extension.
    pub fn detect(path: &Path) -> Result<Self> {
        match path
            .extension()
            .and_then(|e| e.to_str())
            .map(|e| e.to_ascii_lowercase())
            .as_deref()
        {
            Some("xlsx") => Ok(Self::Xlsx),
            Some("docx") => Ok(Self::Docx),
            Some("pptx") => Ok(Self::Pptx),
            Some(ext) => Err(IrError::UnsupportedFormat(format!(".{ext}"))),
            None => Err(IrError::UnsupportedFormat("no extension".into())),
        }
    }

    /// Detect format from a filename string.
    pub fn detect_from_name(name: &str) -> Result<Self> {
        let name_lower = name.to_ascii_lowercase();
        if name_lower.ends_with(".xlsx") {
            Ok(Self::Xlsx)
        } else if name_lower.ends_with(".docx") {
            Ok(Self::Docx)
        } else if name_lower.ends_with(".pptx") {
            Ok(Self::Pptx)
        } else {
            Err(IrError::UnsupportedFormat(format!(
                "unknown extension in: {name}"
            )))
        }
    }
}

/// Options for deriving an IR from an Office file.
#[derive(Debug, Clone, Default)]
pub struct DeriveOptions {
    /// IR mode (content, style, full).
    pub mode: Mode,
    /// Filter to a single sheet by name (xlsx only).
    pub sheet: Option<String>,
    /// Filter to a cell range within a sheet (xlsx only, e.g. "A1:D20").
    pub range: Option<String>,
}

/// Options for applying an IR to an Office file.
#[derive(Debug, Clone, Default)]
pub struct ApplyOptions {
    /// Override the source file path (default: from IR header).
    pub source_override: Option<PathBuf>,
    /// Skip checksum validation.
    pub force: bool,
}

/// Result of applying an IR, reporting what changed.
#[derive(Debug, Clone, Default)]
pub struct ApplyResult {
    /// Number of existing cells whose values were updated.
    pub cells_updated: usize,
    /// Number of new cells created.
    pub cells_created: usize,
    /// Number of cells cleared via `<empty>`.
    pub cells_cleared: usize,
    /// Number of charts added (xlsx only).
    pub charts_added: usize,
    /// Warnings encountered during apply.
    pub warnings: Vec<String>,
}

/// Errors from IR operations.
#[derive(Debug, thiserror::Error)]
pub enum IrError {
    /// Error from the xlsx layer.
    #[error("xlsx error: {0}")]
    Xlsx(#[from] offidized_xlsx::XlsxError),

    /// Error from the docx layer.
    #[error("docx error: {0}")]
    Docx(#[from] offidized_docx::DocxError),

    /// Error from the pptx layer.
    #[error("pptx error: {0}")]
    Pptx(#[from] offidized_pptx::PptxError),

    /// Malformed IR header.
    #[error("invalid IR header: {0}")]
    InvalidHeader(String),

    /// Malformed IR body.
    #[error("invalid IR body: {0}")]
    InvalidBody(String),

    /// File format not supported.
    #[error("unsupported format: {0}")]
    UnsupportedFormat(String),

    /// Mode not yet implemented for the given format.
    #[error("{mode} mode not yet implemented for {format}")]
    UnsupportedMode {
        /// The file format.
        format: String,
        /// The requested mode.
        mode: String,
    },

    /// Source file has changed since the IR was derived.
    #[error("checksum mismatch: source has changed since derive (expected {expected}, got {actual}). Use --force to override.")]
    ChecksumMismatch {
        /// Checksum at derive time.
        expected: String,
        /// Current checksum.
        actual: String,
    },

    /// I/O error.
    #[error("io error: {0}")]
    Io(#[from] std::io::Error),
}

/// Alias for `std::result::Result<T, IrError>`.
pub type Result<T> = std::result::Result<T, IrError>;

/// Separator line between content and style sections in full-mode IR.
const FULL_MODE_SEPARATOR: &str = "\n--- style ---\n";

/// Split a full-mode IR body into content and style parts.
fn split_full_body(body: &str) -> (&str, &str) {
    if let Some(pos) = body.find("\n--- style ---\n") {
        let content = &body[..pos];
        let style = &body[pos + "\n--- style ---\n".len()..];
        (content, style)
    } else {
        // No separator found — treat entire body as content, empty style
        (body, "")
    }
}

/// Merge a style `ApplyResult` into a content `ApplyResult`.
fn merge_apply_results(target: &mut ApplyResult, style: &ApplyResult) {
    target.cells_updated += style.cells_updated;
    target.cells_created += style.cells_created;
    target.cells_cleared += style.cells_cleared;
    target.charts_added += style.charts_added;
    target.warnings.extend(style.warnings.iter().cloned());
}

/// Compute SHA-256 checksum of a file, returned as `"sha256:<hex>"`.
fn compute_checksum(path: &Path) -> Result<String> {
    let bytes = std::fs::read(path)?;
    compute_checksum_bytes(&bytes)
}

/// Compute SHA-256 checksum of bytes, returned as `"sha256:<hex>"`.
fn compute_checksum_bytes(bytes: &[u8]) -> Result<String> {
    let mut hasher = Sha256::new();
    hasher.update(bytes);
    let result = hasher.finalize();
    Ok(format!("sha256:{result:x}"))
}

/// Derive a text IR from an Office file.
///
/// Returns the complete IR as a string (header + body). Writes to stdout
/// by default when used via CLI; the library returns the string directly.
pub fn derive(path: &Path, options: DeriveOptions) -> Result<String> {
    let format = Format::detect(path)?;
    let checksum = compute_checksum(path)?;
    let source_str = path.to_string_lossy().to_string();

    let header = IrHeader {
        source: source_str,
        format,
        mode: options.mode,
        version: 1,
        checksum,
    };
    let mut output = header.write();

    match (format, options.mode) {
        (Format::Xlsx, Mode::Content) => {
            let wb = offidized_xlsx::Workbook::open(path)?;
            xlsx::content::derive_content(&wb, &options, &mut output);
        }
        (Format::Docx, Mode::Content) => {
            let doc = offidized_docx::Document::open(path)?;
            docx::content::derive_content(&doc, &mut output);
        }
        (Format::Pptx, Mode::Content) => {
            let prs = offidized_pptx::Presentation::open(path)?;
            pptx::content::derive_content(&prs, &mut output);
        }
        (Format::Xlsx, Mode::Style) => {
            let wb = offidized_xlsx::Workbook::open(path)?;
            xlsx::style::derive_style(&wb, &options, &mut output);
        }
        (Format::Docx, Mode::Style) => {
            let doc = offidized_docx::Document::open(path)?;
            docx::style::derive_style(&doc, &mut output);
        }
        (Format::Pptx, Mode::Style) => {
            let prs = offidized_pptx::Presentation::open(path)?;
            pptx::style::derive_style(&prs, &mut output);
        }
        (Format::Xlsx, Mode::Full) => {
            let wb = offidized_xlsx::Workbook::open(path)?;
            xlsx::content::derive_content(&wb, &options, &mut output);
            output.push_str(FULL_MODE_SEPARATOR);
            xlsx::style::derive_style(&wb, &options, &mut output);
        }
        (Format::Docx, Mode::Full) => {
            let doc = offidized_docx::Document::open(path)?;
            docx::content::derive_content(&doc, &mut output);
            output.push_str(FULL_MODE_SEPARATOR);
            docx::style::derive_style(&doc, &mut output);
        }
        (Format::Pptx, Mode::Full) => {
            let prs = offidized_pptx::Presentation::open(path)?;
            pptx::content::derive_content(&prs, &mut output);
            output.push_str(FULL_MODE_SEPARATOR);
            pptx::style::derive_style(&prs, &mut output);
        }
    }

    Ok(output)
}

/// Apply an IR to produce an updated Office file.
///
/// Reads the source file path from the IR header (or uses `options.source_override`),
/// applies the changes described in the IR body, and saves to `output`.
pub fn apply(ir: &str, output: &Path, options: &ApplyOptions) -> Result<ApplyResult> {
    let (header, body) = IrHeader::parse(ir)?;

    let source = if let Some(ref src) = options.source_override {
        src.clone()
    } else {
        PathBuf::from(&header.source)
    };

    // Validate checksum unless forced
    if !options.force {
        let current_checksum = compute_checksum(&source)?;
        if current_checksum != header.checksum {
            return Err(IrError::ChecksumMismatch {
                expected: header.checksum,
                actual: current_checksum,
            });
        }
    }

    match (header.format, header.mode) {
        (Format::Xlsx, Mode::Content) => {
            let mut wb = offidized_xlsx::Workbook::open(&source)?;
            let result = xlsx::content::apply_content(&body, &mut wb)?;
            wb.save(output)?;
            Ok(result)
        }
        (Format::Docx, Mode::Content) => {
            let mut doc = offidized_docx::Document::open(&source)?;
            let result = docx::content::apply_content(&body, &mut doc)?;
            doc.save(output)?;
            Ok(result)
        }
        (Format::Pptx, Mode::Content) => {
            let mut prs = offidized_pptx::Presentation::open(&source)?;
            let result = pptx::content::apply_content(&body, &mut prs)?;
            prs.save(output)?;
            Ok(result)
        }
        (Format::Xlsx, Mode::Style) => {
            let mut wb = offidized_xlsx::Workbook::open(&source)?;
            let result = xlsx::style::apply_style(&body, &mut wb)?;
            wb.save(output)?;
            Ok(result)
        }
        (Format::Docx, Mode::Style) => {
            let mut doc = offidized_docx::Document::open(&source)?;
            let result = docx::style::apply_style(&body, &mut doc)?;
            doc.save(output)?;
            Ok(result)
        }
        (Format::Pptx, Mode::Style) => {
            let mut prs = offidized_pptx::Presentation::open(&source)?;
            let result = pptx::style::apply_style(&body, &mut prs)?;
            prs.save(output)?;
            Ok(result)
        }
        (Format::Xlsx, Mode::Full) => {
            let (content_body, style_body) = split_full_body(&body);
            let mut wb = offidized_xlsx::Workbook::open(&source)?;
            let mut result = xlsx::content::apply_content(content_body, &mut wb)?;
            let style_result = xlsx::style::apply_style(style_body, &mut wb)?;
            merge_apply_results(&mut result, &style_result);
            wb.save(output)?;
            Ok(result)
        }
        (Format::Docx, Mode::Full) => {
            let (content_body, style_body) = split_full_body(&body);
            let mut doc = offidized_docx::Document::open(&source)?;
            let mut result = docx::content::apply_content(content_body, &mut doc)?;
            let style_result = docx::style::apply_style(style_body, &mut doc)?;
            merge_apply_results(&mut result, &style_result);
            doc.save(output)?;
            Ok(result)
        }
        (Format::Pptx, Mode::Full) => {
            let (content_body, style_body) = split_full_body(&body);
            let mut prs = offidized_pptx::Presentation::open(&source)?;
            let mut result = pptx::content::apply_content(content_body, &mut prs)?;
            let style_result = pptx::style::apply_style(style_body, &mut prs)?;
            merge_apply_results(&mut result, &style_result);
            prs.save(output)?;
            Ok(result)
        }
    }
}

/// Derive a text IR from in-memory Office file bytes.
///
/// Returns the complete IR as a string (header + body). The `source_name` is used
/// for the IR header's source field (can be any identifier, e.g., "document.xlsx").
pub fn derive_from_bytes(
    bytes: &[u8],
    source_name: &str,
    options: DeriveOptions,
) -> Result<String> {
    let format = Format::detect_from_name(source_name)?;
    let checksum = compute_checksum_bytes(bytes)?;

    let header = IrHeader {
        source: source_name.to_string(),
        format,
        mode: options.mode,
        version: 1,
        checksum,
    };
    let mut output = header.write();

    match (format, options.mode) {
        (Format::Xlsx, Mode::Content) => {
            let wb = offidized_xlsx::Workbook::from_bytes(bytes)?;
            xlsx::content::derive_content(&wb, &options, &mut output);
        }
        (Format::Docx, Mode::Content) => {
            let doc = offidized_docx::Document::from_bytes(bytes)?;
            docx::content::derive_content(&doc, &mut output);
        }
        (Format::Pptx, Mode::Content) => {
            let prs = offidized_pptx::Presentation::from_bytes(bytes)?;
            pptx::content::derive_content(&prs, &mut output);
        }
        (Format::Xlsx, Mode::Style) => {
            let wb = offidized_xlsx::Workbook::from_bytes(bytes)?;
            xlsx::style::derive_style(&wb, &options, &mut output);
        }
        (Format::Docx, Mode::Style) => {
            let doc = offidized_docx::Document::from_bytes(bytes)?;
            docx::style::derive_style(&doc, &mut output);
        }
        (Format::Pptx, Mode::Style) => {
            let prs = offidized_pptx::Presentation::from_bytes(bytes)?;
            pptx::style::derive_style(&prs, &mut output);
        }
        (Format::Xlsx, Mode::Full) => {
            let wb = offidized_xlsx::Workbook::from_bytes(bytes)?;
            xlsx::content::derive_content(&wb, &options, &mut output);
            output.push_str(FULL_MODE_SEPARATOR);
            xlsx::style::derive_style(&wb, &options, &mut output);
        }
        (Format::Docx, Mode::Full) => {
            let doc = offidized_docx::Document::from_bytes(bytes)?;
            docx::content::derive_content(&doc, &mut output);
            output.push_str(FULL_MODE_SEPARATOR);
            docx::style::derive_style(&doc, &mut output);
        }
        (Format::Pptx, Mode::Full) => {
            let prs = offidized_pptx::Presentation::from_bytes(bytes)?;
            pptx::content::derive_content(&prs, &mut output);
            output.push_str(FULL_MODE_SEPARATOR);
            pptx::style::derive_style(&prs, &mut output);
        }
    }

    Ok(output)
}

/// Apply an IR to in-memory Office file bytes, returning updated bytes.
///
/// Takes the source file bytes and the IR string, applies the changes,
/// and returns the updated file as bytes.
pub fn apply_to_bytes(
    source_bytes: &[u8],
    ir: &str,
    options: &ApplyOptions,
) -> Result<(Vec<u8>, ApplyResult)> {
    let (header, body) = IrHeader::parse(ir)?;

    // Validate checksum unless forced
    if !options.force {
        let current_checksum = compute_checksum_bytes(source_bytes)?;
        if current_checksum != header.checksum {
            return Err(IrError::ChecksumMismatch {
                expected: header.checksum,
                actual: current_checksum,
            });
        }
    }

    match (header.format, header.mode) {
        (Format::Xlsx, Mode::Content) => {
            let mut wb = offidized_xlsx::Workbook::from_bytes(source_bytes)?;
            let result = xlsx::content::apply_content(&body, &mut wb)?;
            let bytes = wb.to_bytes()?;
            Ok((bytes, result))
        }
        (Format::Docx, Mode::Content) => {
            let mut doc = offidized_docx::Document::from_bytes(source_bytes)?;
            let result = docx::content::apply_content(&body, &mut doc)?;
            let bytes = docx_save_to_bytes(&doc)?;
            Ok((bytes, result))
        }
        (Format::Pptx, Mode::Content) => {
            let mut prs = offidized_pptx::Presentation::from_bytes(source_bytes)?;
            let result = pptx::content::apply_content(&body, &mut prs)?;
            let bytes = pptx_save_to_bytes(&mut prs)?;
            Ok((bytes, result))
        }
        (Format::Xlsx, Mode::Style) => {
            let mut wb = offidized_xlsx::Workbook::from_bytes(source_bytes)?;
            let result = xlsx::style::apply_style(&body, &mut wb)?;
            let bytes = wb.to_bytes()?;
            Ok((bytes, result))
        }
        (Format::Docx, Mode::Style) => {
            let mut doc = offidized_docx::Document::from_bytes(source_bytes)?;
            let result = docx::style::apply_style(&body, &mut doc)?;
            let bytes = docx_save_to_bytes(&doc)?;
            Ok((bytes, result))
        }
        (Format::Pptx, Mode::Style) => {
            let mut prs = offidized_pptx::Presentation::from_bytes(source_bytes)?;
            let result = pptx::style::apply_style(&body, &mut prs)?;
            let bytes = pptx_save_to_bytes(&mut prs)?;
            Ok((bytes, result))
        }
        (Format::Xlsx, Mode::Full) => {
            let (content_body, style_body) = split_full_body(&body);
            let mut wb = offidized_xlsx::Workbook::from_bytes(source_bytes)?;
            let mut result = xlsx::content::apply_content(content_body, &mut wb)?;
            let style_result = xlsx::style::apply_style(style_body, &mut wb)?;
            merge_apply_results(&mut result, &style_result);
            let bytes = wb.to_bytes()?;
            Ok((bytes, result))
        }
        (Format::Docx, Mode::Full) => {
            let (content_body, style_body) = split_full_body(&body);
            let mut doc = offidized_docx::Document::from_bytes(source_bytes)?;
            let mut result = docx::content::apply_content(content_body, &mut doc)?;
            let style_result = docx::style::apply_style(style_body, &mut doc)?;
            merge_apply_results(&mut result, &style_result);
            let bytes = docx_save_to_bytes(&doc)?;
            Ok((bytes, result))
        }
        (Format::Pptx, Mode::Full) => {
            let (content_body, style_body) = split_full_body(&body);
            let mut prs = offidized_pptx::Presentation::from_bytes(source_bytes)?;
            let mut result = pptx::content::apply_content(content_body, &mut prs)?;
            let style_result = pptx::style::apply_style(style_body, &mut prs)?;
            merge_apply_results(&mut result, &style_result);
            let bytes = pptx_save_to_bytes(&mut prs)?;
            Ok((bytes, result))
        }
    }
}

fn docx_save_to_bytes(doc: &offidized_docx::Document) -> Result<Vec<u8>> {
    use tempfile::NamedTempFile;
    let tmp = NamedTempFile::new()?;
    doc.save(tmp.path())?;
    Ok(std::fs::read(tmp.path())?)
}

fn pptx_save_to_bytes(prs: &mut offidized_pptx::Presentation) -> Result<Vec<u8>> {
    use tempfile::NamedTempFile;
    let tmp = NamedTempFile::new()?;
    prs.save(tmp.path())?;
    Ok(std::fs::read(tmp.path())?)
}
