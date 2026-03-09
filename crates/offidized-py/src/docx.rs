//! Python bindings for the offidized-docx Word document API.

use std::sync::{Arc, Mutex};

use offidized_docx::{
    Bookmark, Comment, Document as CoreDocument, DocumentProtection, Endnote, Footnote,
    HeaderFooter, ParagraphAlignment, StyleKind, TabStop, TabStopAlignment, TableAlignment,
    TableLayout, VerticalAlignment, VerticalMerge,
};
use pyo3::prelude::*;

use crate::error::{docx_error_to_py, value_error};

// =============================================================================
// Helpers
// =============================================================================

fn lock_doc(doc: &Arc<Mutex<CoreDocument>>) -> PyResult<std::sync::MutexGuard<'_, CoreDocument>> {
    doc.lock()
        .map_err(|e| value_error(format!("Failed to lock document: {e}")))
}

fn alignment_to_string(a: ParagraphAlignment) -> &'static str {
    match a {
        ParagraphAlignment::Left => "left",
        ParagraphAlignment::Center => "center",
        ParagraphAlignment::Right => "right",
        ParagraphAlignment::Justified => "justified",
    }
}

fn string_to_alignment(s: &str) -> PyResult<ParagraphAlignment> {
    match s.to_lowercase().as_str() {
        "left" => Ok(ParagraphAlignment::Left),
        "center" => Ok(ParagraphAlignment::Center),
        "right" => Ok(ParagraphAlignment::Right),
        "justified" | "justify" => Ok(ParagraphAlignment::Justified),
        _ => Err(value_error(format!("Unknown alignment: {s}"))),
    }
}

// =============================================================================
// Document
// =============================================================================

/// Python wrapper for `offidized_docx::Document`.
#[pyclass(module = "offidized._native", name = "Document")]
pub struct Document {
    inner: Arc<Mutex<CoreDocument>>,
}

impl Default for Document {
    fn default() -> Self {
        Self::new()
    }
}

#[pymethods]
impl Document {
    #[new]
    pub fn new() -> Self {
        Self {
            inner: Arc::new(Mutex::new(CoreDocument::new())),
        }
    }

    /// Open an existing document from a file path.
    #[staticmethod]
    pub fn open(path: &str) -> PyResult<Self> {
        let document = CoreDocument::open(path).map_err(docx_error_to_py)?;
        Ok(Self {
            inner: Arc::new(Mutex::new(document)),
        })
    }

    /// Create a document from in-memory bytes.
    #[staticmethod]
    pub fn from_bytes(bytes: &[u8]) -> PyResult<Self> {
        let document = CoreDocument::from_bytes(bytes).map_err(docx_error_to_py)?;
        Ok(Self {
            inner: Arc::new(Mutex::new(document)),
        })
    }

    /// Save document to a `.docx` path.
    pub fn save(&self, path: &str) -> PyResult<()> {
        let doc = lock_doc(&self.inner)?;
        doc.save(path).map_err(docx_error_to_py)
    }

    /// Serialize document to bytes.
    pub fn to_bytes(&self) -> PyResult<Vec<u8>> {
        let doc = lock_doc(&self.inner)?;
        let dir = std::env::temp_dir();
        let path = dir.join(format!("offidized-{}.docx", std::process::id()));
        doc.save(&path).map_err(docx_error_to_py)?;
        let bytes = std::fs::read(&path)
            .map_err(|e| value_error(format!("Failed to read temp file: {e}")))?;
        let _ = std::fs::remove_file(&path);
        Ok(bytes)
    }

    // -- Paragraphs ----------------------------------------------------------

    /// Get the number of paragraphs.
    pub fn paragraph_count(&self) -> PyResult<usize> {
        let doc = lock_doc(&self.inner)?;
        Ok(doc.paragraphs().len())
    }

    /// Get a paragraph by index.
    pub fn paragraph(&self, index: usize) -> PyResult<DocxParagraph> {
        let doc = lock_doc(&self.inner)?;
        if index >= doc.paragraphs().len() {
            return Err(value_error(format!(
                "Paragraph index {index} out of range (0..{})",
                doc.paragraphs().len()
            )));
        }
        Ok(DocxParagraph {
            document: Arc::clone(&self.inner),
            index,
        })
    }

    /// Get all paragraphs as a list.
    pub fn paragraphs(&self) -> PyResult<Vec<DocxParagraph>> {
        let doc = lock_doc(&self.inner)?;
        Ok((0..doc.paragraphs().len())
            .map(|i| DocxParagraph {
                document: Arc::clone(&self.inner),
                index: i,
            })
            .collect())
    }

    /// Add a paragraph with plain text. Returns the new paragraph.
    pub fn add_paragraph(&mut self, text: &str) -> PyResult<DocxParagraph> {
        let mut doc = lock_doc(&self.inner)?;
        doc.add_paragraph(text);
        let index = doc.paragraphs().len() - 1;
        Ok(DocxParagraph {
            document: Arc::clone(&self.inner),
            index,
        })
    }

    /// Add a paragraph with a style. Returns the new paragraph.
    pub fn add_paragraph_with_style(
        &mut self,
        text: &str,
        style_id: &str,
    ) -> PyResult<DocxParagraph> {
        let mut doc = lock_doc(&self.inner)?;
        doc.add_paragraph_with_style(text, style_id);
        let index = doc.paragraphs().len() - 1;
        Ok(DocxParagraph {
            document: Arc::clone(&self.inner),
            index,
        })
    }

    /// Add a heading paragraph at the given level (1..=9).
    pub fn add_heading(&mut self, text: &str, level: u8) -> PyResult<DocxParagraph> {
        if !(1..=9).contains(&level) {
            return Err(value_error("level must be between 1 and 9"));
        }
        let mut doc = lock_doc(&self.inner)?;
        doc.add_heading(text, level);
        let index = doc.paragraphs().len() - 1;
        Ok(DocxParagraph {
            document: Arc::clone(&self.inner),
            index,
        })
    }

    /// Add a bulleted paragraph. Returns the new paragraph.
    pub fn add_bulleted_paragraph(&mut self, text: &str) -> PyResult<DocxParagraph> {
        let mut doc = lock_doc(&self.inner)?;
        doc.add_bulleted_paragraph(text);
        let index = doc.paragraphs().len() - 1;
        Ok(DocxParagraph {
            document: Arc::clone(&self.inner),
            index,
        })
    }

    /// Add a numbered paragraph. Returns the new paragraph.
    pub fn add_numbered_paragraph(&mut self, text: &str) -> PyResult<DocxParagraph> {
        let mut doc = lock_doc(&self.inner)?;
        doc.add_numbered_paragraph(text);
        let index = doc.paragraphs().len() - 1;
        Ok(DocxParagraph {
            document: Arc::clone(&self.inner),
            index,
        })
    }

    // -- Tables --------------------------------------------------------------

    /// Get the number of tables.
    pub fn table_count(&self) -> PyResult<usize> {
        let doc = lock_doc(&self.inner)?;
        Ok(doc.tables().len())
    }

    /// Get a table by index.
    pub fn table(&self, index: usize) -> PyResult<DocxTable> {
        let doc = lock_doc(&self.inner)?;
        if index >= doc.tables().len() {
            return Err(value_error(format!(
                "Table index {index} out of range (0..{})",
                doc.tables().len()
            )));
        }
        Ok(DocxTable {
            document: Arc::clone(&self.inner),
            table_index: index,
        })
    }

    /// Add a table with the given number of rows and columns.
    pub fn add_table(&mut self, rows: usize, columns: usize) -> PyResult<DocxTable> {
        let mut doc = lock_doc(&self.inner)?;
        doc.add_table(rows, columns);
        let index = doc.tables().len() - 1;
        Ok(DocxTable {
            document: Arc::clone(&self.inner),
            table_index: index,
        })
    }

    /// Get all tables as a list.
    pub fn tables(&self) -> PyResult<Vec<DocxTable>> {
        let doc = lock_doc(&self.inner)?;
        Ok((0..doc.tables().len())
            .map(|i| DocxTable {
                document: Arc::clone(&self.inner),
                table_index: i,
            })
            .collect())
    }

    // -- Images --------------------------------------------------------------

    /// Add an image from bytes with a content type. Returns the image index.
    pub fn add_image(&mut self, bytes: Vec<u8>, content_type: &str) -> PyResult<usize> {
        let mut doc = lock_doc(&self.inner)?;
        Ok(doc.add_image(bytes, content_type))
    }

    /// Get the number of images.
    pub fn image_count(&self) -> PyResult<usize> {
        let doc = lock_doc(&self.inner)?;
        Ok(doc.images().len())
    }

    // -- Section -------------------------------------------------------------

    /// Get the document section (page layout).
    pub fn section(&self) -> PyResult<DocxSection> {
        Ok(DocxSection {
            document: Arc::clone(&self.inner),
        })
    }

    // -- Document Properties -------------------------------------------------

    /// Get document properties (metadata).
    pub fn document_properties(&self) -> PyResult<DocxDocumentProperties> {
        Ok(DocxDocumentProperties {
            document: Arc::clone(&self.inner),
        })
    }

    // -- Counts --------------------------------------------------------------

    /// Get the number of comments.
    pub fn comment_count(&self) -> PyResult<usize> {
        let doc = lock_doc(&self.inner)?;
        Ok(doc.comments().len())
    }

    /// Get the number of footnotes.
    pub fn footnote_count(&self) -> PyResult<usize> {
        let doc = lock_doc(&self.inner)?;
        Ok(doc.footnotes().len())
    }

    /// Get the number of bookmarks.
    pub fn bookmark_count(&self) -> PyResult<usize> {
        let doc = lock_doc(&self.inner)?;
        Ok(doc.bookmarks().len())
    }

    // -- Comments ------------------------------------------------------------

    /// Get comment texts as a list of (id, author, text) tuples.
    pub fn comments(&self) -> PyResult<Vec<(u32, String, String)>> {
        let doc = lock_doc(&self.inner)?;
        Ok(doc
            .comments()
            .iter()
            .map(|c| (c.id(), c.author().to_owned(), c.text()))
            .collect())
    }

    /// Add a comment with the given id, author, and text. Returns the comment id.
    pub fn add_comment(&mut self, id: u32, author: &str, text: &str) -> PyResult<u32> {
        let mut doc = lock_doc(&self.inner)?;
        let comment = Comment::from_text(id, author, text);
        let added = doc.add_comment(comment);
        Ok(added.id())
    }

    // -- Footnotes -----------------------------------------------------------

    /// Get footnote texts as a list of (id, text) tuples.
    pub fn footnotes(&self) -> PyResult<Vec<(u32, String)>> {
        let doc = lock_doc(&self.inner)?;
        Ok(doc.footnotes().iter().map(|f| (f.id(), f.text())).collect())
    }

    /// Add a footnote with the given id and text. Returns the footnote id.
    pub fn add_footnote(&mut self, id: u32, text: &str) -> PyResult<u32> {
        let mut doc = lock_doc(&self.inner)?;
        let footnote = Footnote::from_text(id, text);
        let added = doc.add_footnote(footnote);
        Ok(added.id())
    }

    // -- Bookmarks -----------------------------------------------------------

    /// Get bookmarks as a list of (id, name) tuples.
    pub fn bookmarks(&self) -> PyResult<Vec<(u32, String)>> {
        let doc = lock_doc(&self.inner)?;
        Ok(doc
            .bookmarks()
            .iter()
            .map(|b| (b.id(), b.name().to_owned()))
            .collect())
    }

    /// Add a bookmark. Returns the bookmark id.
    pub fn add_bookmark(
        &mut self,
        id: u32,
        name: &str,
        start_para: usize,
        end_para: usize,
    ) -> PyResult<u32> {
        let mut doc = lock_doc(&self.inner)?;
        let bookmark = Bookmark::new(id, name, start_para, end_para);
        let added = doc.add_bookmark(bookmark);
        Ok(added.id())
    }

    // -- Content Controls ----------------------------------------------------

    /// Get the number of content controls (structured document tags).
    pub fn content_control_count(&self) -> PyResult<usize> {
        let doc = lock_doc(&self.inner)?;
        Ok(doc.content_controls().len())
    }

    // -- Protection ----------------------------------------------------------

    /// Get document protection as (edit_type, enforcement), or None if unset.
    pub fn protection(&self) -> PyResult<Option<(String, bool)>> {
        let doc = lock_doc(&self.inner)?;
        Ok(doc.protection().map(|p| (p.edit.clone(), p.enforcement)))
    }

    /// Set document protection. edit_type: "readOnly", "comments", "trackedChanges", "forms".
    pub fn set_protection(&mut self, edit_type: &str, enforcement: bool) -> PyResult<()> {
        let mut doc = lock_doc(&self.inner)?;
        doc.set_protection(DocumentProtection::new(edit_type, enforcement));
        Ok(())
    }

    /// Clear document protection.
    pub fn clear_protection(&mut self) -> PyResult<()> {
        let mut doc = lock_doc(&self.inner)?;
        doc.clear_protection();
        Ok(())
    }

    // -- Styles --------------------------------------------------------------

    /// Get the number of styles in the style registry.
    pub fn style_count(&self) -> PyResult<usize> {
        let doc = lock_doc(&self.inner)?;
        Ok(doc.styles().styles().len())
    }

    /// Get all style IDs as (kind, style_id, name) tuples.
    /// kind is "paragraph", "character", or "table".
    pub fn styles(&self) -> PyResult<Vec<(String, String, Option<String>)>> {
        let doc = lock_doc(&self.inner)?;
        Ok(doc
            .styles()
            .styles()
            .iter()
            .map(|s| {
                let kind = match s.kind() {
                    StyleKind::Paragraph => "paragraph",
                    StyleKind::Character => "character",
                    StyleKind::Table => "table",
                }
                .to_owned();
                (kind, s.style_id().to_owned(), s.name().map(String::from))
            })
            .collect())
    }

    /// Add a paragraph style. Returns the style_id.
    pub fn add_paragraph_style(&mut self, style_id: &str) -> PyResult<String> {
        let mut doc = lock_doc(&self.inner)?;
        doc.styles_mut().add_paragraph_style(style_id);
        Ok(style_id.to_owned())
    }

    /// Add a character style. Returns the style_id.
    pub fn add_character_style(&mut self, style_id: &str) -> PyResult<String> {
        let mut doc = lock_doc(&self.inner)?;
        doc.styles_mut().add_character_style(style_id);
        Ok(style_id.to_owned())
    }

    /// Add a table style. Returns the style_id.
    pub fn add_table_style(&mut self, style_id: &str) -> PyResult<String> {
        let mut doc = lock_doc(&self.inner)?;
        doc.styles_mut().add_table_style(style_id);
        Ok(style_id.to_owned())
    }

    // -- Endnotes ------------------------------------------------------------

    /// Get the number of endnotes.
    pub fn endnote_count(&self) -> PyResult<usize> {
        let doc = lock_doc(&self.inner)?;
        Ok(doc.endnotes().len())
    }

    /// Get endnote texts as a list of (id, text) tuples.
    pub fn endnotes(&self) -> PyResult<Vec<(u32, String)>> {
        let doc = lock_doc(&self.inner)?;
        Ok(doc.endnotes().iter().map(|e| (e.id(), e.text())).collect())
    }

    /// Add an endnote with the given id and text. Returns the endnote id.
    pub fn add_endnote(&mut self, id: u32, text: &str) -> PyResult<u32> {
        let mut doc = lock_doc(&self.inner)?;
        let endnote = Endnote::from_text(id, text);
        let added = doc.add_endnote(endnote);
        Ok(added.id())
    }

    // -- Body Items ----------------------------------------------------------

    /// Get body items in document order as a list of ("paragraph", index) or ("table", index).
    pub fn body_items(&self) -> PyResult<Vec<(String, usize)>> {
        let doc = lock_doc(&self.inner)?;
        let mut items = Vec::new();
        let mut para_idx = 0usize;
        let mut table_idx = 0usize;
        for item in doc.body_items() {
            match item {
                offidized_docx::BodyItem::Paragraph(_) => {
                    items.push(("paragraph".to_owned(), para_idx));
                    para_idx += 1;
                }
                offidized_docx::BodyItem::Table(_) => {
                    items.push(("table".to_owned(), table_idx));
                    table_idx += 1;
                }
            }
        }
        Ok(items)
    }
}

// =============================================================================
// DocxParagraph
// =============================================================================

/// Python wrapper for a paragraph within a Document.
#[pyclass(module = "offidized._native", name = "DocxParagraph", from_py_object)]
#[derive(Clone)]
pub struct DocxParagraph {
    document: Arc<Mutex<CoreDocument>>,
    index: usize,
}

#[pymethods]
impl DocxParagraph {
    /// Get the concatenated text of all runs.
    pub fn text(&self) -> PyResult<String> {
        let doc = lock_doc(&self.document)?;
        let para = doc
            .paragraphs()
            .get(self.index)
            .ok_or_else(|| value_error("Paragraph no longer exists"))?;
        Ok(para.text())
    }

    /// Replace all runs with a single run containing the given text.
    pub fn set_text(&mut self, text: &str) -> PyResult<()> {
        let mut doc = lock_doc(&self.document)?;
        let para = doc
            .paragraphs_mut()
            .get_mut(self.index)
            .ok_or_else(|| value_error("Paragraph no longer exists"))?;
        para.set_text(text);
        Ok(())
    }

    /// Get all runs as a list.
    pub fn runs(&self) -> PyResult<Vec<DocxRun>> {
        let doc = lock_doc(&self.document)?;
        let para = doc
            .paragraphs()
            .get(self.index)
            .ok_or_else(|| value_error("Paragraph no longer exists"))?;
        Ok((0..para.runs().len())
            .map(|i| DocxRun {
                document: Arc::clone(&self.document),
                para_index: self.index,
                run_index: i,
            })
            .collect())
    }

    /// Get the number of runs.
    pub fn run_count(&self) -> PyResult<usize> {
        let doc = lock_doc(&self.document)?;
        let para = doc
            .paragraphs()
            .get(self.index)
            .ok_or_else(|| value_error("Paragraph no longer exists"))?;
        Ok(para.runs().len())
    }

    /// Get a run by index.
    pub fn run(&self, index: usize) -> PyResult<DocxRun> {
        let doc = lock_doc(&self.document)?;
        let para = doc
            .paragraphs()
            .get(self.index)
            .ok_or_else(|| value_error("Paragraph no longer exists"))?;
        if index >= para.runs().len() {
            return Err(value_error(format!(
                "Run index {index} out of range (0..{})",
                para.runs().len()
            )));
        }
        Ok(DocxRun {
            document: Arc::clone(&self.document),
            para_index: self.index,
            run_index: index,
        })
    }

    /// Add a run with text. Returns the new run.
    pub fn add_run(&mut self, text: &str) -> PyResult<DocxRun> {
        let mut doc = lock_doc(&self.document)?;
        let para = doc
            .paragraphs_mut()
            .get_mut(self.index)
            .ok_or_else(|| value_error("Paragraph no longer exists"))?;
        para.add_run(text);
        let run_index = para.runs().len() - 1;
        Ok(DocxRun {
            document: Arc::clone(&self.document),
            para_index: self.index,
            run_index,
        })
    }

    /// Add a run with a character style. Returns the new run.
    pub fn add_run_with_style(&mut self, text: &str, style_id: &str) -> PyResult<DocxRun> {
        let mut doc = lock_doc(&self.document)?;
        let para = doc
            .paragraphs_mut()
            .get_mut(self.index)
            .ok_or_else(|| value_error("Paragraph no longer exists"))?;
        para.add_run_with_style(text, style_id);
        let run_index = para.runs().len() - 1;
        Ok(DocxRun {
            document: Arc::clone(&self.document),
            para_index: self.index,
            run_index,
        })
    }

    /// Add a hyperlink run. Returns the new run.
    pub fn add_hyperlink(&mut self, text: &str, url: &str) -> PyResult<DocxRun> {
        let mut doc = lock_doc(&self.document)?;
        let para = doc
            .paragraphs_mut()
            .get_mut(self.index)
            .ok_or_else(|| value_error("Paragraph no longer exists"))?;
        para.add_hyperlink(text, url);
        let run_index = para.runs().len() - 1;
        Ok(DocxRun {
            document: Arc::clone(&self.document),
            para_index: self.index,
            run_index,
        })
    }

    /// Add an inline image run. Returns the new run.
    pub fn add_inline_image(
        &mut self,
        image_index: usize,
        width_emu: u32,
        height_emu: u32,
    ) -> PyResult<DocxRun> {
        let mut doc = lock_doc(&self.document)?;
        let para = doc
            .paragraphs_mut()
            .get_mut(self.index)
            .ok_or_else(|| value_error("Paragraph no longer exists"))?;
        para.add_inline_image(image_index, width_emu, height_emu);
        let run_index = para.runs().len() - 1;
        Ok(DocxRun {
            document: Arc::clone(&self.document),
            para_index: self.index,
            run_index,
        })
    }

    /// Get the paragraph style ID, if set.
    pub fn style_id(&self) -> PyResult<Option<String>> {
        let doc = lock_doc(&self.document)?;
        let para = doc
            .paragraphs()
            .get(self.index)
            .ok_or_else(|| value_error("Paragraph no longer exists"))?;
        Ok(para.style_id().map(String::from))
    }

    /// Set the paragraph style ID.
    pub fn set_style_id(&mut self, style_id: &str) -> PyResult<()> {
        let mut doc = lock_doc(&self.document)?;
        let para = doc
            .paragraphs_mut()
            .get_mut(self.index)
            .ok_or_else(|| value_error("Paragraph no longer exists"))?;
        para.set_style_id(style_id);
        Ok(())
    }

    /// Clear the paragraph style ID.
    pub fn clear_style_id(&mut self) -> PyResult<()> {
        let mut doc = lock_doc(&self.document)?;
        let para = doc
            .paragraphs_mut()
            .get_mut(self.index)
            .ok_or_else(|| value_error("Paragraph no longer exists"))?;
        para.clear_style_id();
        Ok(())
    }

    /// Get alignment as a string: "left", "center", "right", "justified", or None.
    pub fn alignment(&self) -> PyResult<Option<String>> {
        let doc = lock_doc(&self.document)?;
        let para = doc
            .paragraphs()
            .get(self.index)
            .ok_or_else(|| value_error("Paragraph no longer exists"))?;
        Ok(para.alignment().map(|a| alignment_to_string(a).to_owned()))
    }

    /// Set alignment from a string: "left", "center", "right", "justified".
    pub fn set_alignment(&mut self, alignment: &str) -> PyResult<()> {
        let align = string_to_alignment(alignment)?;
        let mut doc = lock_doc(&self.document)?;
        let para = doc
            .paragraphs_mut()
            .get_mut(self.index)
            .ok_or_else(|| value_error("Paragraph no longer exists"))?;
        para.set_alignment(align);
        Ok(())
    }

    /// Clear the paragraph alignment.
    pub fn clear_alignment(&mut self) -> PyResult<()> {
        let mut doc = lock_doc(&self.document)?;
        let para = doc
            .paragraphs_mut()
            .get_mut(self.index)
            .ok_or_else(|| value_error("Paragraph no longer exists"))?;
        para.clear_alignment();
        Ok(())
    }

    /// Get the heading level, if this is a heading paragraph.
    pub fn heading_level(&self) -> PyResult<Option<u8>> {
        let doc = lock_doc(&self.document)?;
        let para = doc
            .paragraphs()
            .get(self.index)
            .ok_or_else(|| value_error("Paragraph no longer exists"))?;
        Ok(para.heading_level())
    }

    // -- Spacing -------------------------------------------------------------

    /// Get spacing before in twips.
    pub fn spacing_before_twips(&self) -> PyResult<Option<u32>> {
        let doc = lock_doc(&self.document)?;
        let para = doc
            .paragraphs()
            .get(self.index)
            .ok_or_else(|| value_error("Paragraph no longer exists"))?;
        Ok(para.spacing_before_twips())
    }

    /// Set spacing before in twips.
    pub fn set_spacing_before_twips(&mut self, value: u32) -> PyResult<()> {
        let mut doc = lock_doc(&self.document)?;
        let para = doc
            .paragraphs_mut()
            .get_mut(self.index)
            .ok_or_else(|| value_error("Paragraph no longer exists"))?;
        para.set_spacing_before_twips(value);
        Ok(())
    }

    /// Get spacing after in twips.
    pub fn spacing_after_twips(&self) -> PyResult<Option<u32>> {
        let doc = lock_doc(&self.document)?;
        let para = doc
            .paragraphs()
            .get(self.index)
            .ok_or_else(|| value_error("Paragraph no longer exists"))?;
        Ok(para.spacing_after_twips())
    }

    /// Set spacing after in twips.
    pub fn set_spacing_after_twips(&mut self, value: u32) -> PyResult<()> {
        let mut doc = lock_doc(&self.document)?;
        let para = doc
            .paragraphs_mut()
            .get_mut(self.index)
            .ok_or_else(|| value_error("Paragraph no longer exists"))?;
        para.set_spacing_after_twips(value);
        Ok(())
    }

    /// Get line spacing in twips.
    pub fn line_spacing_twips(&self) -> PyResult<Option<u32>> {
        let doc = lock_doc(&self.document)?;
        let para = doc
            .paragraphs()
            .get(self.index)
            .ok_or_else(|| value_error("Paragraph no longer exists"))?;
        Ok(para.line_spacing_twips())
    }

    /// Set line spacing in twips.
    pub fn set_line_spacing_twips(&mut self, value: u32) -> PyResult<()> {
        let mut doc = lock_doc(&self.document)?;
        let para = doc
            .paragraphs_mut()
            .get_mut(self.index)
            .ok_or_else(|| value_error("Paragraph no longer exists"))?;
        para.set_line_spacing_twips(value);
        Ok(())
    }

    // -- Indents -------------------------------------------------------------

    /// Get left indent in twips.
    pub fn indent_left_twips(&self) -> PyResult<Option<u32>> {
        let doc = lock_doc(&self.document)?;
        let para = doc
            .paragraphs()
            .get(self.index)
            .ok_or_else(|| value_error("Paragraph no longer exists"))?;
        Ok(para.indent_left_twips())
    }

    /// Set left indent in twips.
    pub fn set_indent_left_twips(&mut self, value: u32) -> PyResult<()> {
        let mut doc = lock_doc(&self.document)?;
        let para = doc
            .paragraphs_mut()
            .get_mut(self.index)
            .ok_or_else(|| value_error("Paragraph no longer exists"))?;
        para.set_indent_left_twips(value);
        Ok(())
    }

    /// Get right indent in twips.
    pub fn indent_right_twips(&self) -> PyResult<Option<u32>> {
        let doc = lock_doc(&self.document)?;
        let para = doc
            .paragraphs()
            .get(self.index)
            .ok_or_else(|| value_error("Paragraph no longer exists"))?;
        Ok(para.indent_right_twips())
    }

    /// Set right indent in twips.
    pub fn set_indent_right_twips(&mut self, value: u32) -> PyResult<()> {
        let mut doc = lock_doc(&self.document)?;
        let para = doc
            .paragraphs_mut()
            .get_mut(self.index)
            .ok_or_else(|| value_error("Paragraph no longer exists"))?;
        para.set_indent_right_twips(value);
        Ok(())
    }

    /// Get first-line indent in twips.
    pub fn indent_first_line_twips(&self) -> PyResult<Option<u32>> {
        let doc = lock_doc(&self.document)?;
        let para = doc
            .paragraphs()
            .get(self.index)
            .ok_or_else(|| value_error("Paragraph no longer exists"))?;
        Ok(para.indent_first_line_twips())
    }

    /// Set first-line indent in twips.
    pub fn set_indent_first_line_twips(&mut self, value: u32) -> PyResult<()> {
        let mut doc = lock_doc(&self.document)?;
        let para = doc
            .paragraphs_mut()
            .get_mut(self.index)
            .ok_or_else(|| value_error("Paragraph no longer exists"))?;
        para.set_indent_first_line_twips(value);
        Ok(())
    }

    /// Get hanging indent in twips.
    pub fn indent_hanging_twips(&self) -> PyResult<Option<u32>> {
        let doc = lock_doc(&self.document)?;
        let para = doc
            .paragraphs()
            .get(self.index)
            .ok_or_else(|| value_error("Paragraph no longer exists"))?;
        Ok(para.indent_hanging_twips())
    }

    /// Set hanging indent in twips.
    pub fn set_indent_hanging_twips(&mut self, value: u32) -> PyResult<()> {
        let mut doc = lock_doc(&self.document)?;
        let para = doc
            .paragraphs_mut()
            .get_mut(self.index)
            .ok_or_else(|| value_error("Paragraph no longer exists"))?;
        para.set_indent_hanging_twips(value);
        Ok(())
    }

    // -- Numbering -----------------------------------------------------------

    /// Get the numbering num_id, if set.
    pub fn numbering_num_id(&self) -> PyResult<Option<u32>> {
        let doc = lock_doc(&self.document)?;
        let para = doc
            .paragraphs()
            .get(self.index)
            .ok_or_else(|| value_error("Paragraph no longer exists"))?;
        Ok(para.numbering_num_id())
    }

    /// Get the numbering indent level, if set.
    pub fn numbering_ilvl(&self) -> PyResult<Option<u8>> {
        let doc = lock_doc(&self.document)?;
        let para = doc
            .paragraphs()
            .get(self.index)
            .ok_or_else(|| value_error("Paragraph no longer exists"))?;
        Ok(para.numbering_ilvl())
    }

    /// Set numbering with num_id and indent level.
    pub fn set_numbering(&mut self, num_id: u32, ilvl: u8) -> PyResult<()> {
        let mut doc = lock_doc(&self.document)?;
        let para = doc
            .paragraphs_mut()
            .get_mut(self.index)
            .ok_or_else(|| value_error("Paragraph no longer exists"))?;
        para.set_numbering(num_id, ilvl);
        Ok(())
    }

    /// Clear numbering.
    pub fn clear_numbering(&mut self) -> PyResult<()> {
        let mut doc = lock_doc(&self.document)?;
        let para = doc
            .paragraphs_mut()
            .get_mut(self.index)
            .ok_or_else(|| value_error("Paragraph no longer exists"))?;
        para.clear_numbering();
        Ok(())
    }

    // -- Pagination flags ----------------------------------------------------

    /// Get keep-with-next flag.
    pub fn keep_next(&self) -> PyResult<bool> {
        let doc = lock_doc(&self.document)?;
        let para = doc
            .paragraphs()
            .get(self.index)
            .ok_or_else(|| value_error("Paragraph no longer exists"))?;
        Ok(para.keep_next())
    }

    /// Set keep-with-next flag.
    pub fn set_keep_next(&mut self, value: bool) -> PyResult<()> {
        let mut doc = lock_doc(&self.document)?;
        let para = doc
            .paragraphs_mut()
            .get_mut(self.index)
            .ok_or_else(|| value_error("Paragraph no longer exists"))?;
        para.set_keep_next(value);
        Ok(())
    }

    /// Get keep-lines-together flag.
    pub fn keep_lines(&self) -> PyResult<bool> {
        let doc = lock_doc(&self.document)?;
        let para = doc
            .paragraphs()
            .get(self.index)
            .ok_or_else(|| value_error("Paragraph no longer exists"))?;
        Ok(para.keep_lines())
    }

    /// Set keep-lines-together flag.
    pub fn set_keep_lines(&mut self, value: bool) -> PyResult<()> {
        let mut doc = lock_doc(&self.document)?;
        let para = doc
            .paragraphs_mut()
            .get_mut(self.index)
            .ok_or_else(|| value_error("Paragraph no longer exists"))?;
        para.set_keep_lines(value);
        Ok(())
    }

    /// Get page-break-before flag.
    pub fn page_break_before(&self) -> PyResult<bool> {
        let doc = lock_doc(&self.document)?;
        let para = doc
            .paragraphs()
            .get(self.index)
            .ok_or_else(|| value_error("Paragraph no longer exists"))?;
        Ok(para.page_break_before())
    }

    /// Set page-break-before flag.
    pub fn set_page_break_before(&mut self, value: bool) -> PyResult<()> {
        let mut doc = lock_doc(&self.document)?;
        let para = doc
            .paragraphs_mut()
            .get_mut(self.index)
            .ok_or_else(|| value_error("Paragraph no longer exists"))?;
        para.set_page_break_before(value);
        Ok(())
    }

    // -- Shading -------------------------------------------------------------

    /// Get shading/background color as hex string, or None.
    pub fn shading_color(&self) -> PyResult<Option<String>> {
        let doc = lock_doc(&self.document)?;
        let para = doc
            .paragraphs()
            .get(self.index)
            .ok_or_else(|| value_error("Paragraph no longer exists"))?;
        Ok(para.shading_color().map(String::from))
    }

    /// Set shading/background color as hex string.
    pub fn set_shading_color(&mut self, color: &str) -> PyResult<()> {
        let mut doc = lock_doc(&self.document)?;
        let para = doc
            .paragraphs_mut()
            .get_mut(self.index)
            .ok_or_else(|| value_error("Paragraph no longer exists"))?;
        para.set_shading_color(color);
        Ok(())
    }

    /// Clear shading color.
    pub fn clear_shading_color(&mut self) -> PyResult<()> {
        let mut doc = lock_doc(&self.document)?;
        let para = doc
            .paragraphs_mut()
            .get_mut(self.index)
            .ok_or_else(|| value_error("Paragraph no longer exists"))?;
        para.clear_shading_color();
        Ok(())
    }

    // -- Tab Stops -----------------------------------------------------------

    /// Get tab stops as list of (position_twips, alignment) tuples.
    pub fn tab_stops(&self) -> PyResult<Vec<(u32, String)>> {
        let doc = lock_doc(&self.document)?;
        let para = doc
            .paragraphs()
            .get(self.index)
            .ok_or_else(|| value_error("Paragraph no longer exists"))?;
        Ok(para
            .tab_stops()
            .iter()
            .map(|ts| {
                let align = match ts.alignment() {
                    TabStopAlignment::Left => "left",
                    TabStopAlignment::Center => "center",
                    TabStopAlignment::Right => "right",
                    TabStopAlignment::Decimal => "decimal",
                    TabStopAlignment::Bar => "bar",
                    TabStopAlignment::Clear => "clear",
                }
                .to_owned();
                (ts.position_twips(), align)
            })
            .collect())
    }

    /// Add a tab stop at the given position with alignment.
    /// alignment: "left", "center", "right", "decimal", "bar", "clear".
    pub fn add_tab_stop(&mut self, position_twips: u32, alignment: &str) -> PyResult<()> {
        let align = match alignment.to_lowercase().as_str() {
            "left" => TabStopAlignment::Left,
            "center" => TabStopAlignment::Center,
            "right" => TabStopAlignment::Right,
            "decimal" => TabStopAlignment::Decimal,
            "bar" => TabStopAlignment::Bar,
            "clear" => TabStopAlignment::Clear,
            _ => {
                return Err(value_error(format!(
                    "Unknown tab stop alignment: {alignment}"
                )))
            }
        };
        let mut doc = lock_doc(&self.document)?;
        let para = doc
            .paragraphs_mut()
            .get_mut(self.index)
            .ok_or_else(|| value_error("Paragraph no longer exists"))?;
        para.add_tab_stop(TabStop::new(position_twips, align));
        Ok(())
    }

    /// Clear all tab stops.
    pub fn clear_tab_stops(&mut self) -> PyResult<()> {
        let mut doc = lock_doc(&self.document)?;
        let para = doc
            .paragraphs_mut()
            .get_mut(self.index)
            .ok_or_else(|| value_error("Paragraph no longer exists"))?;
        para.clear_tab_stops();
        Ok(())
    }
}

// =============================================================================
// DocxRun
// =============================================================================

/// Python wrapper for a text run within a paragraph.
#[pyclass(module = "offidized._native", name = "DocxRun", from_py_object)]
#[derive(Clone)]
pub struct DocxRun {
    document: Arc<Mutex<CoreDocument>>,
    para_index: usize,
    run_index: usize,
}

/// Helper macro to reduce boilerplate for run getters/setters.
macro_rules! with_run {
    ($self:ident, $doc:ident, $run:ident, $body:expr) => {{
        let $doc = lock_doc(&$self.document)?;
        let para = $doc
            .paragraphs()
            .get($self.para_index)
            .ok_or_else(|| value_error("Paragraph no longer exists"))?;
        let $run = para
            .runs()
            .get($self.run_index)
            .ok_or_else(|| value_error("Run no longer exists"))?;
        $body
    }};
}

macro_rules! with_run_mut {
    ($self:ident, $doc:ident, $run:ident, $body:expr) => {{
        let mut $doc = lock_doc(&$self.document)?;
        let para = $doc
            .paragraphs_mut()
            .get_mut($self.para_index)
            .ok_or_else(|| value_error("Paragraph no longer exists"))?;
        let $run = para
            .runs_mut()
            .get_mut($self.run_index)
            .ok_or_else(|| value_error("Run no longer exists"))?;
        $body
    }};
}

#[pymethods]
impl DocxRun {
    /// Get the run text.
    pub fn text(&self) -> PyResult<String> {
        with_run!(self, doc, run, Ok(run.text().to_owned()))
    }

    /// Set the run text.
    pub fn set_text(&mut self, text: &str) -> PyResult<()> {
        with_run_mut!(self, doc, run, {
            run.set_text(text);
            Ok(())
        })
    }

    // -- Font formatting -----------------------------------------------------

    /// Check if bold.
    pub fn is_bold(&self) -> PyResult<bool> {
        with_run!(self, doc, run, Ok(run.is_bold()))
    }

    /// Set bold.
    pub fn set_bold(&mut self, bold: bool) -> PyResult<()> {
        with_run_mut!(self, doc, run, {
            run.set_bold(bold);
            Ok(())
        })
    }

    /// Check if italic.
    pub fn is_italic(&self) -> PyResult<bool> {
        with_run!(self, doc, run, Ok(run.is_italic()))
    }

    /// Set italic.
    pub fn set_italic(&mut self, italic: bool) -> PyResult<()> {
        with_run_mut!(self, doc, run, {
            run.set_italic(italic);
            Ok(())
        })
    }

    /// Check if underlined.
    pub fn is_underline(&self) -> PyResult<bool> {
        with_run!(self, doc, run, Ok(run.is_underline()))
    }

    /// Set underline.
    pub fn set_underline(&mut self, underline: bool) -> PyResult<()> {
        with_run_mut!(self, doc, run, {
            run.set_underline(underline);
            Ok(())
        })
    }

    /// Check if strikethrough.
    pub fn is_strikethrough(&self) -> PyResult<bool> {
        with_run!(self, doc, run, Ok(run.is_strikethrough()))
    }

    /// Set strikethrough.
    pub fn set_strikethrough(&mut self, value: bool) -> PyResult<()> {
        with_run_mut!(self, doc, run, {
            run.set_strikethrough(value);
            Ok(())
        })
    }

    /// Check if hidden.
    pub fn is_hidden(&self) -> PyResult<bool> {
        with_run!(self, doc, run, Ok(run.is_hidden()))
    }

    /// Set hidden.
    pub fn set_hidden(&mut self, value: bool) -> PyResult<()> {
        with_run_mut!(self, doc, run, {
            run.set_hidden(value);
            Ok(())
        })
    }

    /// Get font family, if set.
    pub fn font_family(&self) -> PyResult<Option<String>> {
        with_run!(self, doc, run, Ok(run.font_family().map(String::from)))
    }

    /// Set font family.
    pub fn set_font_family(&mut self, family: &str) -> PyResult<()> {
        with_run_mut!(self, doc, run, {
            run.set_font_family(family);
            Ok(())
        })
    }

    /// Get font size in half-points.
    pub fn font_size_half_points(&self) -> PyResult<Option<u16>> {
        with_run!(self, doc, run, Ok(run.font_size_half_points()))
    }

    /// Set font size in half-points (e.g. 24 = 12pt).
    pub fn set_font_size_half_points(&mut self, size: u16) -> PyResult<()> {
        with_run_mut!(self, doc, run, {
            run.set_font_size_half_points(size);
            Ok(())
        })
    }

    /// Get text color as hex string (e.g. "FF0000"), if set.
    pub fn color(&self) -> PyResult<Option<String>> {
        with_run!(self, doc, run, Ok(run.color().map(String::from)))
    }

    /// Set text color as hex string (e.g. "FF0000").
    pub fn set_color(&mut self, color: &str) -> PyResult<()> {
        with_run_mut!(self, doc, run, {
            run.set_color(color);
            Ok(())
        })
    }

    /// Clear the text color.
    pub fn clear_color(&mut self) -> PyResult<()> {
        with_run_mut!(self, doc, run, {
            run.clear_color();
            Ok(())
        })
    }

    // -- Hyperlink -----------------------------------------------------------

    /// Get the hyperlink URL, if set.
    pub fn hyperlink(&self) -> PyResult<Option<String>> {
        with_run!(self, doc, run, Ok(run.hyperlink().map(String::from)))
    }

    /// Set a hyperlink URL on this run.
    pub fn set_hyperlink(&mut self, url: &str) -> PyResult<()> {
        with_run_mut!(self, doc, run, {
            run.set_hyperlink(url);
            Ok(())
        })
    }

    /// Clear the hyperlink.
    pub fn clear_hyperlink(&mut self) -> PyResult<()> {
        with_run_mut!(self, doc, run, {
            run.clear_hyperlink();
            Ok(())
        })
    }

    // -- Style ---------------------------------------------------------------

    /// Get the run (character) style ID, if set.
    pub fn style_id(&self) -> PyResult<Option<String>> {
        with_run!(self, doc, run, Ok(run.style_id().map(String::from)))
    }

    /// Set the run style ID.
    pub fn set_style_id(&mut self, style_id: &str) -> PyResult<()> {
        with_run_mut!(self, doc, run, {
            run.set_style_id(style_id);
            Ok(())
        })
    }

    // -- Caps ----------------------------------------------------------------

    /// Check if small caps.
    pub fn is_small_caps(&self) -> PyResult<bool> {
        with_run!(self, doc, run, Ok(run.is_small_caps()))
    }

    /// Set small caps.
    pub fn set_small_caps(&mut self, value: bool) -> PyResult<()> {
        with_run_mut!(self, doc, run, {
            run.set_small_caps(value);
            Ok(())
        })
    }

    /// Check if all caps.
    pub fn is_all_caps(&self) -> PyResult<bool> {
        with_run!(self, doc, run, Ok(run.is_all_caps()))
    }

    /// Set all caps.
    pub fn set_all_caps(&mut self, value: bool) -> PyResult<()> {
        with_run_mut!(self, doc, run, {
            run.set_all_caps(value);
            Ok(())
        })
    }

    // -- Sub/superscript -----------------------------------------------------

    /// Check if subscript.
    pub fn is_subscript(&self) -> PyResult<bool> {
        with_run!(self, doc, run, Ok(run.is_subscript()))
    }

    /// Set subscript.
    pub fn set_subscript(&mut self, value: bool) -> PyResult<()> {
        with_run_mut!(self, doc, run, {
            run.set_subscript(value);
            Ok(())
        })
    }

    /// Check if superscript.
    pub fn is_superscript(&self) -> PyResult<bool> {
        with_run!(self, doc, run, Ok(run.is_superscript()))
    }

    /// Set superscript.
    pub fn set_superscript(&mut self, value: bool) -> PyResult<()> {
        with_run_mut!(self, doc, run, {
            run.set_superscript(value);
            Ok(())
        })
    }

    // -- Highlight -----------------------------------------------------------

    /// Get highlight color name (e.g. "yellow"), if set.
    pub fn highlight_color(&self) -> PyResult<Option<String>> {
        with_run!(self, doc, run, Ok(run.highlight_color().map(String::from)))
    }

    /// Set highlight color by name (e.g. "yellow", "green", "cyan").
    pub fn set_highlight_color(&mut self, color: &str) -> PyResult<()> {
        with_run_mut!(self, doc, run, {
            run.set_highlight_color(color);
            Ok(())
        })
    }
}

// =============================================================================
// DocxTable
// =============================================================================

/// Python wrapper for a table within a Document.
#[pyclass(module = "offidized._native", name = "DocxTable", from_py_object)]
#[derive(Clone)]
pub struct DocxTable {
    document: Arc<Mutex<CoreDocument>>,
    table_index: usize,
}

#[pymethods]
impl DocxTable {
    /// Get the number of rows.
    pub fn rows(&self) -> PyResult<usize> {
        let doc = lock_doc(&self.document)?;
        let table = doc
            .tables()
            .get(self.table_index)
            .ok_or_else(|| value_error("Table no longer exists"))?;
        Ok(table.rows())
    }

    /// Get the number of columns.
    pub fn columns(&self) -> PyResult<usize> {
        let doc = lock_doc(&self.document)?;
        let table = doc
            .tables()
            .get(self.table_index)
            .ok_or_else(|| value_error("Table no longer exists"))?;
        Ok(table.columns())
    }

    /// Get cell text at (row, col), or None if out of range.
    pub fn cell_text(&self, row: usize, col: usize) -> PyResult<Option<String>> {
        let doc = lock_doc(&self.document)?;
        let table = doc
            .tables()
            .get(self.table_index)
            .ok_or_else(|| value_error("Table no longer exists"))?;
        Ok(table.cell_text(row, col).map(String::from))
    }

    /// Set cell text at (row, col).
    pub fn set_cell_text(&mut self, row: usize, col: usize, text: &str) -> PyResult<()> {
        let mut doc = lock_doc(&self.document)?;
        let table = doc
            .tables_mut()
            .get_mut(self.table_index)
            .ok_or_else(|| value_error("Table no longer exists"))?;
        table.set_cell_text(row, col, text);
        Ok(())
    }

    /// Get the table style ID, if set.
    pub fn style_id(&self) -> PyResult<Option<String>> {
        let doc = lock_doc(&self.document)?;
        let table = doc
            .tables()
            .get(self.table_index)
            .ok_or_else(|| value_error("Table no longer exists"))?;
        Ok(table.style_id().map(String::from))
    }

    /// Set the table style ID.
    pub fn set_style_id(&mut self, style_id: &str) -> PyResult<()> {
        let mut doc = lock_doc(&self.document)?;
        let table = doc
            .tables_mut()
            .get_mut(self.table_index)
            .ok_or_else(|| value_error("Table no longer exists"))?;
        table.set_style_id(style_id);
        Ok(())
    }

    /// Merge cells horizontally in a row. Returns true if successful.
    pub fn merge_cells_horizontally(
        &mut self,
        row: usize,
        start_col: usize,
        span: usize,
    ) -> PyResult<bool> {
        let mut doc = lock_doc(&self.document)?;
        let table = doc
            .tables_mut()
            .get_mut(self.table_index)
            .ok_or_else(|| value_error("Table no longer exists"))?;
        Ok(table.merge_cells_horizontally(row, start_col, span))
    }

    /// Clear horizontal merge on a cell.
    pub fn clear_horizontal_merge(&mut self, row: usize, col: usize) -> PyResult<bool> {
        let mut doc = lock_doc(&self.document)?;
        let table = doc
            .tables_mut()
            .get_mut(self.table_index)
            .ok_or_else(|| value_error("Table no longer exists"))?;
        Ok(table.clear_horizontal_merge(row, col))
    }

    // -- Add/Remove Rows/Columns ---------------------------------------------

    /// Add a new empty row at the end. Returns the new row count.
    pub fn add_row(&mut self) -> PyResult<usize> {
        let mut doc = lock_doc(&self.document)?;
        let table = doc
            .tables_mut()
            .get_mut(self.table_index)
            .ok_or_else(|| value_error("Table no longer exists"))?;
        table.add_row();
        Ok(table.rows())
    }

    /// Insert a new row at the given index.
    pub fn insert_row(&mut self, index: usize) -> PyResult<()> {
        let mut doc = lock_doc(&self.document)?;
        let table = doc
            .tables_mut()
            .get_mut(self.table_index)
            .ok_or_else(|| value_error("Table no longer exists"))?;
        table.insert_row(index);
        Ok(())
    }

    /// Remove a row by index. Returns true if successful.
    pub fn remove_row(&mut self, index: usize) -> PyResult<bool> {
        let mut doc = lock_doc(&self.document)?;
        let table = doc
            .tables_mut()
            .get_mut(self.table_index)
            .ok_or_else(|| value_error("Table no longer exists"))?;
        Ok(table.remove_row(index))
    }

    /// Add a new column at the end. Returns the new column count.
    pub fn add_column(&mut self) -> PyResult<usize> {
        let mut doc = lock_doc(&self.document)?;
        let table = doc
            .tables_mut()
            .get_mut(self.table_index)
            .ok_or_else(|| value_error("Table no longer exists"))?;
        table.add_column();
        Ok(table.columns())
    }

    // -- Width, Layout, Column Widths ----------------------------------------

    /// Get table width in twips, or None.
    pub fn width_twips(&self) -> PyResult<Option<u32>> {
        let doc = lock_doc(&self.document)?;
        let table = doc
            .tables()
            .get(self.table_index)
            .ok_or_else(|| value_error("Table no longer exists"))?;
        Ok(table.width_twips())
    }

    /// Set table width in twips.
    pub fn set_width_twips(&mut self, width: u32) -> PyResult<()> {
        let mut doc = lock_doc(&self.document)?;
        let table = doc
            .tables_mut()
            .get_mut(self.table_index)
            .ok_or_else(|| value_error("Table no longer exists"))?;
        table.set_width_twips(width);
        Ok(())
    }

    /// Clear table width.
    pub fn clear_width_twips(&mut self) -> PyResult<()> {
        let mut doc = lock_doc(&self.document)?;
        let table = doc
            .tables_mut()
            .get_mut(self.table_index)
            .ok_or_else(|| value_error("Table no longer exists"))?;
        table.clear_width_twips();
        Ok(())
    }

    /// Get table layout: "fixed" or "autofit", or None.
    pub fn layout(&self) -> PyResult<Option<String>> {
        let doc = lock_doc(&self.document)?;
        let table = doc
            .tables()
            .get(self.table_index)
            .ok_or_else(|| value_error("Table no longer exists"))?;
        Ok(table.layout().map(|l| {
            match l {
                TableLayout::Fixed => "fixed",
                TableLayout::AutoFit => "autofit",
            }
            .to_owned()
        }))
    }

    /// Set table layout: "fixed" or "autofit".
    pub fn set_layout(&mut self, layout: &str) -> PyResult<()> {
        let tl = match layout.to_lowercase().as_str() {
            "fixed" => TableLayout::Fixed,
            "autofit" => TableLayout::AutoFit,
            _ => return Err(value_error(format!("Unknown table layout: {layout}"))),
        };
        let mut doc = lock_doc(&self.document)?;
        let table = doc
            .tables_mut()
            .get_mut(self.table_index)
            .ok_or_else(|| value_error("Table no longer exists"))?;
        table.set_layout(tl);
        Ok(())
    }

    /// Get column widths in twips.
    pub fn column_widths_twips(&self) -> PyResult<Vec<u32>> {
        let doc = lock_doc(&self.document)?;
        let table = doc
            .tables()
            .get(self.table_index)
            .ok_or_else(|| value_error("Table no longer exists"))?;
        Ok(table.column_widths_twips().to_vec())
    }

    /// Set column widths in twips.
    pub fn set_column_widths_twips(&mut self, widths: Vec<u32>) -> PyResult<()> {
        let mut doc = lock_doc(&self.document)?;
        let table = doc
            .tables_mut()
            .get_mut(self.table_index)
            .ok_or_else(|| value_error("Table no longer exists"))?;
        table.set_column_widths_twips(widths);
        Ok(())
    }

    /// Get table alignment: "left", "center", "right", or None.
    pub fn alignment(&self) -> PyResult<Option<String>> {
        let doc = lock_doc(&self.document)?;
        let table = doc
            .tables()
            .get(self.table_index)
            .ok_or_else(|| value_error("Table no longer exists"))?;
        Ok(table.alignment().map(|a| {
            match a {
                TableAlignment::Left => "left",
                TableAlignment::Center => "center",
                TableAlignment::Right => "right",
            }
            .to_owned()
        }))
    }

    /// Set table alignment: "left", "center", "right".
    pub fn set_alignment(&mut self, alignment: &str) -> PyResult<()> {
        let align = match alignment.to_lowercase().as_str() {
            "left" => TableAlignment::Left,
            "center" => TableAlignment::Center,
            "right" => TableAlignment::Right,
            _ => return Err(value_error(format!("Unknown table alignment: {alignment}"))),
        };
        let mut doc = lock_doc(&self.document)?;
        let table = doc
            .tables_mut()
            .get_mut(self.table_index)
            .ok_or_else(|| value_error("Table no longer exists"))?;
        table.set_alignment(align);
        Ok(())
    }

    /// Get a cell wrapper for the given (row, col).
    pub fn cell(&self, row: usize, col: usize) -> PyResult<DocxTableCell> {
        let doc = lock_doc(&self.document)?;
        let table = doc
            .tables()
            .get(self.table_index)
            .ok_or_else(|| value_error("Table no longer exists"))?;
        if table.cell(row, col).is_none() {
            return Err(value_error(format!("Cell ({row}, {col}) out of range")));
        }
        Ok(DocxTableCell {
            document: Arc::clone(&self.document),
            table_index: self.table_index,
            row,
            col,
        })
    }
}

// =============================================================================
// DocxTableCell
// =============================================================================

/// Python wrapper for a single table cell.
#[pyclass(module = "offidized._native", name = "DocxTableCell", from_py_object)]
#[derive(Clone)]
pub struct DocxTableCell {
    document: Arc<Mutex<CoreDocument>>,
    table_index: usize,
    row: usize,
    col: usize,
}

#[pymethods]
impl DocxTableCell {
    /// Get cell text.
    pub fn text(&self) -> PyResult<String> {
        let doc = lock_doc(&self.document)?;
        let table = doc
            .tables()
            .get(self.table_index)
            .ok_or_else(|| value_error("Table no longer exists"))?;
        let cell = table
            .cell(self.row, self.col)
            .ok_or_else(|| value_error("Cell no longer exists"))?;
        Ok(cell.text().to_owned())
    }

    /// Set cell text.
    pub fn set_text(&mut self, text: &str) -> PyResult<()> {
        let mut doc = lock_doc(&self.document)?;
        let table = doc
            .tables_mut()
            .get_mut(self.table_index)
            .ok_or_else(|| value_error("Table no longer exists"))?;
        let cell = table
            .cell_mut(self.row, self.col)
            .ok_or_else(|| value_error("Cell no longer exists"))?;
        cell.set_text(text);
        Ok(())
    }

    /// Get cell shading color as hex string, if set.
    pub fn shading_color(&self) -> PyResult<Option<String>> {
        let doc = lock_doc(&self.document)?;
        let table = doc
            .tables()
            .get(self.table_index)
            .ok_or_else(|| value_error("Table no longer exists"))?;
        let cell = table
            .cell(self.row, self.col)
            .ok_or_else(|| value_error("Cell no longer exists"))?;
        Ok(cell.shading_color().map(String::from))
    }

    /// Set cell shading color as hex string.
    pub fn set_shading_color(&mut self, color: &str) -> PyResult<()> {
        let mut doc = lock_doc(&self.document)?;
        let table = doc
            .tables_mut()
            .get_mut(self.table_index)
            .ok_or_else(|| value_error("Table no longer exists"))?;
        let cell = table
            .cell_mut(self.row, self.col)
            .ok_or_else(|| value_error("Cell no longer exists"))?;
        cell.set_shading_color(color);
        Ok(())
    }

    /// Get cell vertical alignment: "top", "center", "bottom", or None.
    pub fn vertical_alignment(&self) -> PyResult<Option<String>> {
        let doc = lock_doc(&self.document)?;
        let table = doc
            .tables()
            .get(self.table_index)
            .ok_or_else(|| value_error("Table no longer exists"))?;
        let cell = table
            .cell(self.row, self.col)
            .ok_or_else(|| value_error("Cell no longer exists"))?;
        Ok(cell.vertical_alignment().map(|a| {
            match a {
                VerticalAlignment::Top => "top",
                VerticalAlignment::Center => "center",
                VerticalAlignment::Bottom => "bottom",
            }
            .to_owned()
        }))
    }

    /// Set cell vertical alignment: "top", "center", "bottom".
    pub fn set_vertical_alignment(&mut self, alignment: &str) -> PyResult<()> {
        let align = match alignment.to_lowercase().as_str() {
            "top" => VerticalAlignment::Top,
            "center" => VerticalAlignment::Center,
            "bottom" => VerticalAlignment::Bottom,
            _ => {
                return Err(value_error(format!(
                    "Unknown vertical alignment: {alignment}"
                )))
            }
        };
        let mut doc = lock_doc(&self.document)?;
        let table = doc
            .tables_mut()
            .get_mut(self.table_index)
            .ok_or_else(|| value_error("Table no longer exists"))?;
        let cell = table
            .cell_mut(self.row, self.col)
            .ok_or_else(|| value_error("Cell no longer exists"))?;
        cell.set_vertical_alignment(align);
        Ok(())
    }

    // -- Vertical Merge ------------------------------------------------------

    /// Get vertical merge state: "restart", "continue", or None.
    pub fn vertical_merge(&self) -> PyResult<Option<String>> {
        let doc = lock_doc(&self.document)?;
        let table = doc
            .tables()
            .get(self.table_index)
            .ok_or_else(|| value_error("Table no longer exists"))?;
        let cell = table
            .cell(self.row, self.col)
            .ok_or_else(|| value_error("Cell no longer exists"))?;
        Ok(cell.vertical_merge().map(|vm| {
            match vm {
                VerticalMerge::Restart => "restart",
                VerticalMerge::Continue => "continue",
            }
            .to_owned()
        }))
    }

    /// Set vertical merge: "restart" or "continue".
    pub fn set_vertical_merge(&mut self, merge: &str) -> PyResult<()> {
        let vm = match merge.to_lowercase().as_str() {
            "restart" => VerticalMerge::Restart,
            "continue" => VerticalMerge::Continue,
            _ => return Err(value_error(format!("Unknown vertical merge: {merge}"))),
        };
        let mut doc = lock_doc(&self.document)?;
        let table = doc
            .tables_mut()
            .get_mut(self.table_index)
            .ok_or_else(|| value_error("Table no longer exists"))?;
        let cell = table
            .cell_mut(self.row, self.col)
            .ok_or_else(|| value_error("Cell no longer exists"))?;
        cell.set_vertical_merge(vm);
        Ok(())
    }

    /// Clear vertical merge.
    pub fn clear_vertical_merge(&mut self) -> PyResult<()> {
        let mut doc = lock_doc(&self.document)?;
        let table = doc
            .tables_mut()
            .get_mut(self.table_index)
            .ok_or_else(|| value_error("Table no longer exists"))?;
        let cell = table
            .cell_mut(self.row, self.col)
            .ok_or_else(|| value_error("Cell no longer exists"))?;
        cell.clear_vertical_merge();
        Ok(())
    }
}

// =============================================================================
// DocxSection
// =============================================================================

/// Python wrapper for the document section (page layout).
#[pyclass(module = "offidized._native", name = "DocxSection", from_py_object)]
#[derive(Clone)]
pub struct DocxSection {
    document: Arc<Mutex<CoreDocument>>,
}

#[pymethods]
impl DocxSection {
    /// Get page width in twips.
    pub fn page_width_twips(&self) -> PyResult<Option<u32>> {
        let doc = lock_doc(&self.document)?;
        Ok(doc.section().page_width_twips())
    }

    /// Get page height in twips.
    pub fn page_height_twips(&self) -> PyResult<Option<u32>> {
        let doc = lock_doc(&self.document)?;
        Ok(doc.section().page_height_twips())
    }

    /// Set page size in twips.
    pub fn set_page_size_twips(&mut self, width: u32, height: u32) -> PyResult<()> {
        let mut doc = lock_doc(&self.document)?;
        doc.section_mut().set_page_size_twips(width, height);
        Ok(())
    }

    /// Get page orientation: "portrait", "landscape", or None.
    pub fn page_orientation(&self) -> PyResult<Option<String>> {
        let doc = lock_doc(&self.document)?;
        Ok(doc.section().page_orientation().map(|o| {
            match o {
                offidized_docx::PageOrientation::Portrait => "portrait",
                offidized_docx::PageOrientation::Landscape => "landscape",
            }
            .to_owned()
        }))
    }

    /// Set page orientation: "portrait" or "landscape".
    pub fn set_page_orientation(&mut self, orientation: &str) -> PyResult<()> {
        let orient = match orientation.to_lowercase().as_str() {
            "portrait" => offidized_docx::PageOrientation::Portrait,
            "landscape" => offidized_docx::PageOrientation::Landscape,
            _ => return Err(value_error(format!("Unknown orientation: {orientation}"))),
        };
        let mut doc = lock_doc(&self.document)?;
        doc.section_mut().set_page_orientation(orient);
        Ok(())
    }

    // -- Page Margins --------------------------------------------------------

    /// Get top margin in twips.
    pub fn margin_top_twips(&self) -> PyResult<Option<u32>> {
        let doc = lock_doc(&self.document)?;
        Ok(doc.section().page_margins().top_twips())
    }

    /// Set top margin in twips.
    pub fn set_margin_top_twips(&mut self, value: u32) -> PyResult<()> {
        let mut doc = lock_doc(&self.document)?;
        doc.section_mut().page_margins_mut().set_top_twips(value);
        Ok(())
    }

    /// Get right margin in twips.
    pub fn margin_right_twips(&self) -> PyResult<Option<u32>> {
        let doc = lock_doc(&self.document)?;
        Ok(doc.section().page_margins().right_twips())
    }

    /// Set right margin in twips.
    pub fn set_margin_right_twips(&mut self, value: u32) -> PyResult<()> {
        let mut doc = lock_doc(&self.document)?;
        doc.section_mut().page_margins_mut().set_right_twips(value);
        Ok(())
    }

    /// Get bottom margin in twips.
    pub fn margin_bottom_twips(&self) -> PyResult<Option<u32>> {
        let doc = lock_doc(&self.document)?;
        Ok(doc.section().page_margins().bottom_twips())
    }

    /// Set bottom margin in twips.
    pub fn set_margin_bottom_twips(&mut self, value: u32) -> PyResult<()> {
        let mut doc = lock_doc(&self.document)?;
        doc.section_mut().page_margins_mut().set_bottom_twips(value);
        Ok(())
    }

    /// Get left margin in twips.
    pub fn margin_left_twips(&self) -> PyResult<Option<u32>> {
        let doc = lock_doc(&self.document)?;
        Ok(doc.section().page_margins().left_twips())
    }

    /// Set left margin in twips.
    pub fn set_margin_left_twips(&mut self, value: u32) -> PyResult<()> {
        let mut doc = lock_doc(&self.document)?;
        doc.section_mut().page_margins_mut().set_left_twips(value);
        Ok(())
    }

    /// Get header margin in twips.
    pub fn margin_header_twips(&self) -> PyResult<Option<u32>> {
        let doc = lock_doc(&self.document)?;
        Ok(doc.section().page_margins().header_twips())
    }

    /// Set header margin in twips.
    pub fn set_margin_header_twips(&mut self, value: u32) -> PyResult<()> {
        let mut doc = lock_doc(&self.document)?;
        doc.section_mut().page_margins_mut().set_header_twips(value);
        Ok(())
    }

    /// Get footer margin in twips.
    pub fn margin_footer_twips(&self) -> PyResult<Option<u32>> {
        let doc = lock_doc(&self.document)?;
        Ok(doc.section().page_margins().footer_twips())
    }

    /// Set footer margin in twips.
    pub fn set_margin_footer_twips(&mut self, value: u32) -> PyResult<()> {
        let mut doc = lock_doc(&self.document)?;
        doc.section_mut().page_margins_mut().set_footer_twips(value);
        Ok(())
    }

    /// Get gutter margin in twips.
    pub fn margin_gutter_twips(&self) -> PyResult<Option<u32>> {
        let doc = lock_doc(&self.document)?;
        Ok(doc.section().page_margins().gutter_twips())
    }

    /// Set gutter margin in twips.
    pub fn set_margin_gutter_twips(&mut self, value: u32) -> PyResult<()> {
        let mut doc = lock_doc(&self.document)?;
        doc.section_mut().page_margins_mut().set_gutter_twips(value);
        Ok(())
    }

    // -- Headers and Footers -------------------------------------------------

    /// Get the default header text (concatenated paragraph texts).
    pub fn header_text(&self) -> PyResult<Option<String>> {
        let doc = lock_doc(&self.document)?;
        Ok(doc.section().header().map(|hf| {
            hf.paragraphs()
                .iter()
                .map(|p| p.text())
                .collect::<Vec<_>>()
                .join("\n")
        }))
    }

    /// Set the default header with plain text.
    pub fn set_header_text(&mut self, text: &str) -> PyResult<()> {
        let mut doc = lock_doc(&self.document)?;
        doc.section_mut().set_header(HeaderFooter::from_text(text));
        Ok(())
    }

    /// Clear the default header.
    pub fn clear_header(&mut self) -> PyResult<()> {
        let mut doc = lock_doc(&self.document)?;
        doc.section_mut().clear_header();
        Ok(())
    }

    /// Get the default footer text (concatenated paragraph texts).
    pub fn footer_text(&self) -> PyResult<Option<String>> {
        let doc = lock_doc(&self.document)?;
        Ok(doc.section().footer().map(|hf| {
            hf.paragraphs()
                .iter()
                .map(|p| p.text())
                .collect::<Vec<_>>()
                .join("\n")
        }))
    }

    /// Set the default footer with plain text.
    pub fn set_footer_text(&mut self, text: &str) -> PyResult<()> {
        let mut doc = lock_doc(&self.document)?;
        doc.section_mut().set_footer(HeaderFooter::from_text(text));
        Ok(())
    }

    /// Clear the default footer.
    pub fn clear_footer(&mut self) -> PyResult<()> {
        let mut doc = lock_doc(&self.document)?;
        doc.section_mut().clear_footer();
        Ok(())
    }

    /// Get first page header text.
    pub fn first_page_header_text(&self) -> PyResult<Option<String>> {
        let doc = lock_doc(&self.document)?;
        Ok(doc.section().first_page_header().map(|hf| {
            hf.paragraphs()
                .iter()
                .map(|p| p.text())
                .collect::<Vec<_>>()
                .join("\n")
        }))
    }

    /// Set first page header with plain text.
    pub fn set_first_page_header_text(&mut self, text: &str) -> PyResult<()> {
        let mut doc = lock_doc(&self.document)?;
        doc.section_mut()
            .set_first_page_header(HeaderFooter::from_text(text));
        Ok(())
    }

    /// Clear first page header.
    pub fn clear_first_page_header(&mut self) -> PyResult<()> {
        let mut doc = lock_doc(&self.document)?;
        doc.section_mut().clear_first_page_header();
        Ok(())
    }

    /// Get first page footer text.
    pub fn first_page_footer_text(&self) -> PyResult<Option<String>> {
        let doc = lock_doc(&self.document)?;
        Ok(doc.section().first_page_footer().map(|hf| {
            hf.paragraphs()
                .iter()
                .map(|p| p.text())
                .collect::<Vec<_>>()
                .join("\n")
        }))
    }

    /// Set first page footer with plain text.
    pub fn set_first_page_footer_text(&mut self, text: &str) -> PyResult<()> {
        let mut doc = lock_doc(&self.document)?;
        doc.section_mut()
            .set_first_page_footer(HeaderFooter::from_text(text));
        Ok(())
    }

    /// Clear first page footer.
    pub fn clear_first_page_footer(&mut self) -> PyResult<()> {
        let mut doc = lock_doc(&self.document)?;
        doc.section_mut().clear_first_page_footer();
        Ok(())
    }

    /// Get even page header text.
    pub fn even_page_header_text(&self) -> PyResult<Option<String>> {
        let doc = lock_doc(&self.document)?;
        Ok(doc.section().even_page_header().map(|hf| {
            hf.paragraphs()
                .iter()
                .map(|p| p.text())
                .collect::<Vec<_>>()
                .join("\n")
        }))
    }

    /// Set even page header with plain text.
    pub fn set_even_page_header_text(&mut self, text: &str) -> PyResult<()> {
        let mut doc = lock_doc(&self.document)?;
        doc.section_mut()
            .set_even_page_header(HeaderFooter::from_text(text));
        Ok(())
    }

    /// Clear even page header.
    pub fn clear_even_page_header(&mut self) -> PyResult<()> {
        let mut doc = lock_doc(&self.document)?;
        doc.section_mut().clear_even_page_header();
        Ok(())
    }

    /// Get even page footer text.
    pub fn even_page_footer_text(&self) -> PyResult<Option<String>> {
        let doc = lock_doc(&self.document)?;
        Ok(doc.section().even_page_footer().map(|hf| {
            hf.paragraphs()
                .iter()
                .map(|p| p.text())
                .collect::<Vec<_>>()
                .join("\n")
        }))
    }

    /// Set even page footer with plain text.
    pub fn set_even_page_footer_text(&mut self, text: &str) -> PyResult<()> {
        let mut doc = lock_doc(&self.document)?;
        doc.section_mut()
            .set_even_page_footer(HeaderFooter::from_text(text));
        Ok(())
    }

    /// Clear even page footer.
    pub fn clear_even_page_footer(&mut self) -> PyResult<()> {
        let mut doc = lock_doc(&self.document)?;
        doc.section_mut().clear_even_page_footer();
        Ok(())
    }
}

// =============================================================================
// DocxDocumentProperties
// =============================================================================

/// Python wrapper for document properties (metadata).
#[pyclass(
    module = "offidized._native",
    name = "DocxDocumentProperties",
    from_py_object
)]
#[derive(Clone)]
pub struct DocxDocumentProperties {
    document: Arc<Mutex<CoreDocument>>,
}

#[pymethods]
impl DocxDocumentProperties {
    /// Get the document title.
    pub fn title(&self) -> PyResult<Option<String>> {
        let doc = lock_doc(&self.document)?;
        Ok(doc.document_properties().title().map(String::from))
    }

    /// Set the document title.
    pub fn set_title(&mut self, value: &str) -> PyResult<()> {
        let mut doc = lock_doc(&self.document)?;
        doc.document_properties_mut().set_title(value);
        Ok(())
    }

    /// Get the document subject.
    pub fn subject(&self) -> PyResult<Option<String>> {
        let doc = lock_doc(&self.document)?;
        Ok(doc.document_properties().subject().map(String::from))
    }

    /// Set the document subject.
    pub fn set_subject(&mut self, value: &str) -> PyResult<()> {
        let mut doc = lock_doc(&self.document)?;
        doc.document_properties_mut().set_subject(value);
        Ok(())
    }

    /// Get the document creator.
    pub fn creator(&self) -> PyResult<Option<String>> {
        let doc = lock_doc(&self.document)?;
        Ok(doc.document_properties().creator().map(String::from))
    }

    /// Set the document creator.
    pub fn set_creator(&mut self, value: &str) -> PyResult<()> {
        let mut doc = lock_doc(&self.document)?;
        doc.document_properties_mut().set_creator(value);
        Ok(())
    }

    /// Get the document description.
    pub fn description(&self) -> PyResult<Option<String>> {
        let doc = lock_doc(&self.document)?;
        Ok(doc.document_properties().description().map(String::from))
    }

    /// Set the document description.
    pub fn set_description(&mut self, value: &str) -> PyResult<()> {
        let mut doc = lock_doc(&self.document)?;
        doc.document_properties_mut().set_description(value);
        Ok(())
    }

    /// Get the document keywords.
    pub fn keywords(&self) -> PyResult<Option<String>> {
        let doc = lock_doc(&self.document)?;
        Ok(doc.document_properties().keywords().map(String::from))
    }

    /// Set the document keywords.
    pub fn set_keywords(&mut self, value: &str) -> PyResult<()> {
        let mut doc = lock_doc(&self.document)?;
        doc.document_properties_mut().set_keywords(value);
        Ok(())
    }
}

// =============================================================================
// Module Registration
// =============================================================================

pub(crate) fn register(module: &Bound<'_, PyModule>) -> PyResult<()> {
    module.add_class::<Document>()?;
    module.add_class::<DocxParagraph>()?;
    module.add_class::<DocxRun>()?;
    module.add_class::<DocxTable>()?;
    module.add_class::<DocxTableCell>()?;
    module.add_class::<DocxSection>()?;
    module.add_class::<DocxDocumentProperties>()?;
    Ok(())
}
