//! Import pipeline: `.docx` → CRDT document.
//!
//! Converts a parsed [`offidized_docx::Document`] into a [`CrdtDoc`],
//! populating the CRDT's body array, footnotes, endnotes, and image blobs.

use std::collections::HashMap;
use std::sync::Arc;

use yrs::{Any, Array, Map, MapPrelim, Text, TextPrelim, Transact};

use super::crdt_doc::CrdtDoc;
use super::para_id::ParaId;
use super::tokens::{self, TokenType, SENTINEL};

/// Errors during document import.
#[derive(Debug, thiserror::Error)]
pub enum ImportError {
    /// A general import failure.
    #[error("import failed: {0}")]
    General(String),
}

/// Result type for import operations.
pub type Result<T> = std::result::Result<T, ImportError>;

/// Import a parsed docx [`Document`](offidized_docx::Document) into a CRDT document.
///
/// Populates the CRDT with the document's content:
/// - Body paragraphs and tables into the body `YArray`
/// - Image binary blobs into the side-channel blob store
/// - Footnotes and endnotes into their respective `YArray`s
///
/// The original `Document` should be retained separately for export
/// (the CRDT only holds the editable surface).
pub fn import_document(doc: &offidized_docx::Document, crdt: &mut CrdtDoc) -> Result<()> {
    // Collect paragraph registrations to apply after the transaction ends.
    // This avoids borrowing `crdt` mutably while the transaction (which
    // borrows `crdt.doc()`) is still alive.
    let mut registrations: Vec<(ParaId, usize)> = Vec::new();

    // Phase 1: CRDT writes inside a single transaction.
    {
        let body = crdt.body();
        let footnotes_arr = crdt.footnotes();
        let endnotes_arr = crdt.endnotes();
        let mut txn = crdt.doc().transact_mut();

        for (body_index, item) in doc.body_items().enumerate() {
            match item {
                offidized_docx::BodyItem::Paragraph(para) => {
                    let para_id = ParaId::new();

                    // Push a paragraph map into the body array.
                    let prelim = MapPrelim::from([
                        ("type".to_string(), Any::String(Arc::from("paragraph"))),
                        (
                            "id".to_string(),
                            Any::String(Arc::from(para_id.to_string())),
                        ),
                    ]);
                    body.push_back(&mut txn, prelim);

                    // Retrieve the just-inserted MapRef.
                    let map_value = body.get(&txn, body.len(&txn) - 1).ok_or_else(|| {
                        ImportError::General("failed to get inserted paragraph map".into())
                    })?;
                    let map_ref = map_value
                        .cast::<yrs::MapRef>()
                        .map_err(|_| ImportError::General("body element is not a map".into()))?;

                    // Create a nested YText for the paragraph's content.
                    map_ref.insert(&mut txn, "text", TextPrelim::new(""));
                    let text_value = map_ref
                        .get(&txn, "text")
                        .ok_or_else(|| ImportError::General("failed to get text ref".into()))?;
                    let text_ref = text_value
                        .cast::<yrs::TextRef>()
                        .map_err(|_| ImportError::General("text is not a TextRef".into()))?;

                    // Flatten runs into the YText.
                    import_paragraph_content(para, &text_ref, &mut txn);

                    // Set paragraph-level properties on the map.
                    set_paragraph_properties(doc, para, &map_ref, &mut txn);

                    // Defer the registration until after the transaction drops.
                    registrations.push((para_id, body_index));
                }
                offidized_docx::BodyItem::Table(table) => {
                    let para_id = ParaId::new();

                    let prelim = MapPrelim::from([
                        ("type".to_string(), Any::String(Arc::from("table"))),
                        (
                            "id".to_string(),
                            Any::String(Arc::from(para_id.to_string())),
                        ),
                        ("rows".to_string(), Any::Number(table.rows() as f64)),
                        ("columns".to_string(), Any::Number(table.columns() as f64)),
                    ]);
                    body.push_back(&mut txn, prelim);

                    // Retrieve the table map and populate cells.
                    let map_value = body.get(&txn, body.len(&txn) - 1).ok_or_else(|| {
                        ImportError::General("failed to get inserted table map".into())
                    })?;
                    let map_ref = map_value
                        .cast::<yrs::MapRef>()
                        .map_err(|_| ImportError::General("table element is not a map".into()))?;

                    import_table_cells(table, &map_ref, &mut txn)?;

                    registrations.push((para_id, body_index));
                }
            }
        }

        // Import footnotes into the CRDT.
        for footnote in doc.footnotes() {
            let prelim = MapPrelim::from([
                ("id".to_string(), Any::Number(f64::from(footnote.id()))),
                ("text".to_string(), Any::String(Arc::from(footnote.text()))),
            ]);
            footnotes_arr.push_back(&mut txn, prelim);
        }

        // Import endnotes into the CRDT.
        for endnote in doc.endnotes() {
            let prelim = MapPrelim::from([
                ("id".to_string(), Any::Number(f64::from(endnote.id()))),
                ("text".to_string(), Any::String(Arc::from(endnote.text()))),
            ]);
            endnotes_arr.push_back(&mut txn, prelim);
        }

        // Transaction drops here, releasing the borrow on crdt.doc().
    }

    // Phase 2: Apply deferred registrations (needs &mut crdt).
    for (para_id, original_index) in registrations {
        crdt.register_paragraph(para_id, original_index);
    }

    // Phase 3: Import images into the blob store.
    for (index, image) in doc.images().iter().enumerate() {
        let key = format!("img:{index}");
        crdt.image_blobs_mut().insert(key, image.bytes().to_vec());
    }

    Ok(())
}

/// Flatten a paragraph's runs into a YText with formatting attributes.
fn import_paragraph_content(
    para: &offidized_docx::paragraph::Paragraph,
    text_ref: &yrs::TextRef,
    txn: &mut yrs::TransactionMut<'_>,
) {
    let mut position: u32 = 0;
    let sentinel_str = SENTINEL.to_string();

    for run in para.runs() {
        let fmt_attrs = run_formatting_attrs(run);

        // --- Text content ---
        let raw_text = run.text();
        if !raw_text.is_empty() {
            let stripped = tokens::strip_sentinels(raw_text);
            if !stripped.is_empty() {
                text_ref.insert_with_attributes(txn, position, &stripped, fmt_attrs.clone());
                position += stripped.chars().count() as u32;
            }
        }

        // --- Tab ---
        if run.has_tab() {
            let mut attrs = tokens::token_to_attrs(&TokenType::Tab);
            merge_formatting_attrs(&mut attrs, &fmt_attrs);
            text_ref.insert_with_attributes(txn, position, &sentinel_str, attrs);
            position += 1;
        }

        // --- Line break ---
        if run.has_break() {
            let token = TokenType::LineBreak { break_type: None };
            let mut attrs = tokens::token_to_attrs(&token);
            merge_formatting_attrs(&mut attrs, &fmt_attrs);
            text_ref.insert_with_attributes(txn, position, &sentinel_str, attrs);
            position += 1;
        }

        // --- Footnote reference ---
        if let Some(id) = run.footnote_reference_id() {
            let token = TokenType::FootnoteRef { id };
            let mut attrs = tokens::token_to_attrs(&token);
            merge_formatting_attrs(&mut attrs, &fmt_attrs);
            text_ref.insert_with_attributes(txn, position, &sentinel_str, attrs);
            position += 1;
        }

        // --- Endnote reference ---
        if let Some(id) = run.endnote_reference_id() {
            let token = TokenType::EndnoteRef { id };
            let mut attrs = tokens::token_to_attrs(&token);
            merge_formatting_attrs(&mut attrs, &fmt_attrs);
            text_ref.insert_with_attributes(txn, position, &sentinel_str, attrs);
            position += 1;
        }

        // --- Simple field ---
        if let Some(instr) = run.field_simple() {
            let field_type = extract_field_type(instr);
            let presentation = run.text().to_string();
            let token = TokenType::FieldSimple {
                field_type,
                instr: instr.to_string(),
                presentation,
            };
            let mut attrs = tokens::token_to_attrs(&token);
            merge_formatting_attrs(&mut attrs, &fmt_attrs);
            text_ref.insert_with_attributes(txn, position, &sentinel_str, attrs);
            position += 1;
        }

        // --- Complex field (collapsed into FieldCode by the Run API) ---
        if let Some(field_code) = run.field_code() {
            let instr = field_code.instruction();
            let field_type = extract_field_type(instr);
            let token = TokenType::FieldSimple {
                field_type,
                instr: instr.to_string(),
                presentation: field_code.result().to_string(),
            };
            let mut attrs = tokens::token_to_attrs(&token);
            merge_formatting_attrs(&mut attrs, &fmt_attrs);
            text_ref.insert_with_attributes(txn, position, &sentinel_str, attrs);
            position += 1;
        }

        // --- Inline image ---
        if let Some(img) = run.inline_image() {
            let image_ref = format!("img:{}", img.image_index());
            let token = TokenType::InlineImage {
                image_ref,
                width: i64::from(img.width_emu()),
                height: i64::from(img.height_emu()),
            };
            let mut attrs = tokens::token_to_attrs(&token);
            if let Some(name) = img.name() {
                attrs.insert(Arc::from("name"), Any::String(Arc::from(name)));
            }
            if let Some(description) = img.description() {
                attrs.insert(
                    Arc::from("description"),
                    Any::String(Arc::from(description)),
                );
            }
            merge_formatting_attrs(&mut attrs, &fmt_attrs);
            text_ref.insert_with_attributes(txn, position, &sentinel_str, attrs);
            position += 1;
        }
    }
}

/// Set paragraph-level properties on a CRDT paragraph map.
fn set_paragraph_properties(
    doc: &offidized_docx::Document,
    para: &offidized_docx::paragraph::Paragraph,
    map_ref: &yrs::MapRef,
    txn: &mut yrs::TransactionMut<'_>,
) {
    if let Some(alignment) = para.alignment() {
        let s = match alignment {
            offidized_docx::paragraph::ParagraphAlignment::Left => "left",
            offidized_docx::paragraph::ParagraphAlignment::Center => "center",
            offidized_docx::paragraph::ParagraphAlignment::Right => "right",
            offidized_docx::paragraph::ParagraphAlignment::Justified => "justify",
        };
        map_ref.insert(txn, "alignment", Any::String(Arc::from(s)));
    }

    if let Some(level) = para.heading_level() {
        map_ref.insert(txn, "headingLevel", Any::Number(f64::from(level)));
    }

    if let Some(style) = para.style_id() {
        map_ref.insert(txn, "styleId", Any::String(Arc::from(style)));
    }

    if let Some(val) = para.spacing_before_twips() {
        map_ref.insert(txn, "spacingBeforeTwips", Any::Number(f64::from(val)));
    }

    if let Some(val) = para.spacing_after_twips() {
        map_ref.insert(txn, "spacingAfterTwips", Any::Number(f64::from(val)));
    }

    if let Some(val) = para.line_spacing_twips() {
        map_ref.insert(txn, "lineSpacingTwips", Any::Number(f64::from(val)));
    }

    if let Some(rule) = para.line_spacing_rule() {
        let s = match rule {
            offidized_docx::paragraph::LineSpacingRule::Auto => "auto",
            offidized_docx::paragraph::LineSpacingRule::Exact => "exact",
            offidized_docx::paragraph::LineSpacingRule::AtLeast => "atLeast",
        };
        map_ref.insert(txn, "lineSpacingRule", Any::String(Arc::from(s)));
    }

    if let Some(val) = para.indent_left_twips() {
        map_ref.insert(txn, "indentLeftTwips", Any::Number(f64::from(val)));
    }

    if let Some(val) = para.indent_right_twips() {
        map_ref.insert(txn, "indentRightTwips", Any::Number(f64::from(val)));
    }

    if let Some(val) = para.indent_first_line_twips() {
        map_ref.insert(txn, "indentFirstLineTwips", Any::Number(f64::from(val)));
    }

    if let Some(val) = para.indent_hanging_twips() {
        map_ref.insert(txn, "indentHangingTwips", Any::Number(f64::from(val)));
    }

    if let Some(num_id) = para.numbering_num_id() {
        map_ref.insert(txn, "numberingNumId", Any::Number(f64::from(num_id)));
    }

    if let Some(ilvl) = para.numbering_ilvl() {
        map_ref.insert(txn, "numberingIlvl", Any::Number(f64::from(ilvl)));
    }

    if let Some(kind) = resolve_paragraph_numbering_kind(doc, para) {
        map_ref.insert(txn, "numberingKind", Any::String(Arc::from(kind)));
    }

    if para.page_break_before() {
        map_ref.insert(txn, "pageBreakBefore", Any::Bool(true));
    }

    if para.keep_next() {
        map_ref.insert(txn, "keepNext", Any::Bool(true));
    }

    if para.keep_lines() {
        map_ref.insert(txn, "keepLines", Any::Bool(true));
    }
}

fn resolve_paragraph_numbering_kind(
    doc: &offidized_docx::Document,
    para: &offidized_docx::paragraph::Paragraph,
) -> Option<String> {
    let num_id = para.numbering_num_id()?;
    let level = para.numbering_ilvl().unwrap_or(0);
    let instance = doc
        .numbering_instances()
        .iter()
        .find(|inst| inst.num_id() == num_id)?;
    let definition = doc
        .numbering_definitions()
        .iter()
        .find(|def| def.abstract_num_id() == instance.abstract_num_id())?;
    Some(definition.level(level)?.format().to_string())
}

/// Import table cells into the CRDT table map.
///
/// Each cell is stored as `"cell:{row}:{col}"` key in the table map,
/// with a nested `TextPrelim` holding the cell's plain text.
fn import_table_cells(
    table: &offidized_docx::table::Table,
    map_ref: &yrs::MapRef,
    txn: &mut yrs::TransactionMut<'_>,
) -> Result<()> {
    for row in 0..table.rows() {
        for col in 0..table.columns() {
            if let Some(cell) = table.cell(row, col) {
                let key = format!("cell:{row}:{col}");
                map_ref.insert(txn, key.as_str(), TextPrelim::new(cell.text()));
            }
        }
    }
    Ok(())
}

/// Build formatting attributes from a run's character properties.
///
/// Only set attributes are included (absent formatting is omitted,
/// not set to `Any::Null`).
fn run_formatting_attrs(run: &offidized_docx::run::Run) -> HashMap<Arc<str>, Any> {
    let mut attrs = HashMap::new();
    if run.is_bold() {
        attrs.insert(Arc::from("bold"), Any::Bool(true));
    }
    if run.is_italic() {
        attrs.insert(Arc::from("italic"), Any::Bool(true));
    }
    if run.is_underline() {
        attrs.insert(Arc::from("underline"), Any::Bool(true));
    }
    if run.is_strikethrough() {
        attrs.insert(Arc::from("strike"), Any::Bool(true));
    }
    if run.is_superscript() {
        attrs.insert(Arc::from("superscript"), Any::Bool(true));
    }
    if run.is_subscript() {
        attrs.insert(Arc::from("subscript"), Any::Bool(true));
    }
    if run.is_small_caps() {
        attrs.insert(Arc::from("smallCaps"), Any::Bool(true));
    }
    if let Some(font) = run.font_family() {
        attrs.insert(Arc::from("fontFamily"), Any::String(Arc::from(font)));
    }
    if let Some(size) = run.font_size_half_points() {
        attrs.insert(Arc::from("fontSize"), Any::Number(f64::from(size)));
    }
    if let Some(color) = run.color() {
        attrs.insert(Arc::from("color"), Any::String(Arc::from(color)));
    }
    if let Some(highlight) = run.highlight_color() {
        attrs.insert(Arc::from("highlight"), Any::String(Arc::from(highlight)));
    }
    if let Some(link) = run.hyperlink() {
        attrs.insert(Arc::from("hyperlink"), Any::String(Arc::from(link)));
    }
    attrs
}

/// Merge formatting attributes into a token attribute map.
///
/// Token-specific attributes (like `tokenType`) take priority;
/// formatting attrs are added only when the key is not already present.
fn merge_formatting_attrs(
    token_attrs: &mut HashMap<Arc<str>, Any>,
    fmt_attrs: &HashMap<Arc<str>, Any>,
) {
    for (key, value) in fmt_attrs {
        token_attrs
            .entry(Arc::clone(key))
            .or_insert_with(|| value.clone());
    }
}

/// Extract the field type keyword from a field instruction string.
///
/// For example, `" PAGE \\* MERGEFORMAT"` returns `"PAGE"`.
fn extract_field_type(instr: &str) -> String {
    instr
        .split_whitespace()
        .next()
        .unwrap_or("UNKNOWN")
        .to_string()
}

#[cfg(test)]
#[allow(clippy::unwrap_used, clippy::expect_used, clippy::panic)]
mod tests {
    use super::*;
    use offidized_docx::Document;
    use yrs::{GetString, Transact};

    /// Extract an [`Out::Any`] value from a yrs [`Out`].
    fn out_to_any(out: yrs::Out) -> Option<Any> {
        match out {
            yrs::Out::Any(a) => Some(a),
            _ => None,
        }
    }

    #[test]
    fn extract_field_type_basic() {
        assert_eq!(extract_field_type(" PAGE \\* MERGEFORMAT"), "PAGE");
        assert_eq!(extract_field_type("DATE \\@ \"M/d/yyyy\""), "DATE");
        assert_eq!(extract_field_type(""), "UNKNOWN");
        assert_eq!(extract_field_type("   "), "UNKNOWN");
    }

    #[test]
    fn import_empty_document() {
        let doc = Document::new();
        let mut crdt = CrdtDoc::new();
        import_document(&doc, &mut crdt).expect("import empty doc");

        let txn = crdt.doc().transact();
        assert_eq!(crdt.body().len(&txn), 0);
        assert_eq!(crdt.footnotes().len(&txn), 0);
        assert_eq!(crdt.endnotes().len(&txn), 0);
        assert!(crdt.image_blobs().is_empty());
    }

    #[test]
    fn import_single_paragraph() {
        let mut doc = Document::new();
        doc.add_paragraph("Hello world");
        let mut crdt = CrdtDoc::new();
        import_document(&doc, &mut crdt).expect("import single para");

        let txn = crdt.doc().transact();
        assert_eq!(crdt.body().len(&txn), 1);

        // Verify the paragraph map.
        let val = crdt.body().get(&txn, 0).expect("body[0] exists");
        let map = val.cast::<yrs::MapRef>().expect("is a map");
        let type_val = map.get(&txn, "type").expect("has type");
        let type_any = out_to_any(type_val).expect("is Any");
        assert_eq!(type_any, Any::String(Arc::from("paragraph")));

        // Verify the text content.
        let text_val = map.get(&txn, "text").expect("has text");
        let text_ref = text_val.cast::<yrs::TextRef>().expect("is TextRef");
        assert_eq!(text_ref.get_string(&txn), "Hello world");

        // Verify para_index_map has an entry.
        assert_eq!(crdt.para_index_map().len(), 1);
    }

    #[test]
    fn import_paragraph_with_bold_run() {
        let mut doc = Document::new();
        let para = doc.add_paragraph("Bold text");
        // Make the existing run bold.
        if let Some(run) = para.runs_mut().first_mut() {
            run.set_bold(true);
        }

        let mut crdt = CrdtDoc::new();
        import_document(&doc, &mut crdt).expect("import bold para");

        let txn = crdt.doc().transact();
        let val = crdt.body().get(&txn, 0).expect("body[0]");
        let map = val.cast::<yrs::MapRef>().expect("is map");
        let text_val = map.get(&txn, "text").expect("has text");
        let text_ref = text_val.cast::<yrs::TextRef>().expect("is TextRef");
        assert_eq!(text_ref.get_string(&txn), "Bold text");

        // Verify formatting attributes are present via diff.
        let diff = text_ref.diff(&txn, yrs::types::text::YChange::identity);
        assert_eq!(diff.len(), 1);
        let attrs = diff[0].attributes.as_ref().expect("has attrs");
        assert_eq!(attrs.get("bold"), Some(&Any::Bool(true)));
    }

    #[test]
    fn import_paragraph_with_tab() {
        let mut doc = Document::new();
        // Create a paragraph, then set a tab on the existing run.
        let para = doc.add_paragraph("");
        if let Some(run) = para.runs_mut().first_mut() {
            run.set_has_tab(true);
        }

        let mut crdt = CrdtDoc::new();
        import_document(&doc, &mut crdt).expect("import tab para");

        let txn = crdt.doc().transact();
        let val = crdt.body().get(&txn, 0).expect("body[0]");
        let map = val.cast::<yrs::MapRef>().expect("is map");
        let text_val = map.get(&txn, "text").expect("has text");
        let text_ref = text_val.cast::<yrs::TextRef>().expect("is TextRef");

        // Should contain the sentinel character (empty text run contributes nothing,
        // but the tab contributes one sentinel).
        let content = text_ref.get_string(&txn);
        assert_eq!(content, SENTINEL.to_string());

        // Verify the token type in the diff.
        let diff = text_ref.diff(&txn, yrs::types::text::YChange::identity);
        assert_eq!(diff.len(), 1);
        let attrs = diff[0].attributes.as_ref().expect("has attrs");
        assert_eq!(attrs.get("tokenType"), Some(&Any::String(Arc::from("tab"))));
    }

    #[test]
    fn import_multiple_body_items() {
        let mut doc = Document::new();
        doc.add_paragraph("First paragraph");
        doc.add_paragraph("Second paragraph");

        let mut crdt = CrdtDoc::new();
        import_document(&doc, &mut crdt).expect("import multi para");

        let txn = crdt.doc().transact();
        assert_eq!(crdt.body().len(&txn), 2);
        assert_eq!(crdt.para_index_map().len(), 2);
    }

    #[test]
    fn import_footnotes_and_endnotes() {
        let mut doc = Document::new();
        doc.footnotes_mut()
            .push(offidized_docx::Footnote::from_text(1, "Footnote one"));
        doc.endnotes_mut()
            .push(offidized_docx::Endnote::from_text(1, "Endnote one"));

        let mut crdt = CrdtDoc::new();
        import_document(&doc, &mut crdt).expect("import fn/en");

        let txn = crdt.doc().transact();
        assert_eq!(crdt.footnotes().len(&txn), 1);
        assert_eq!(crdt.endnotes().len(&txn), 1);
    }

    #[test]
    fn import_images_stored_in_blobs() {
        let mut doc = Document::new();
        doc.add_image(vec![0xFF_u8, 0xD8], "image/jpeg");

        let mut crdt = CrdtDoc::new();
        import_document(&doc, &mut crdt).expect("import images");

        assert_eq!(crdt.image_blobs().len(), 1);
        assert_eq!(crdt.image_blobs().get("img:0"), Some(&vec![0xFF_u8, 0xD8]));
    }

    #[test]
    fn import_paragraph_heading_level() {
        let mut doc = Document::new();
        doc.add_heading("Heading", 2);

        let mut crdt = CrdtDoc::new();
        import_document(&doc, &mut crdt).expect("import heading");

        let txn = crdt.doc().transact();
        let val = crdt.body().get(&txn, 0).expect("body[0]");
        let map = val.cast::<yrs::MapRef>().expect("is map");
        let level = map.get(&txn, "headingLevel").expect("has headingLevel");
        let level_any = out_to_any(level).expect("is Any");
        assert_eq!(level_any, Any::Number(2.0));
    }

    #[test]
    fn run_formatting_attrs_captures_all_flags() {
        let mut run = offidized_docx::run::Run::new("test");
        run.set_bold(true);
        run.set_italic(true);

        let attrs = run_formatting_attrs(&run);
        assert_eq!(attrs.get("bold" as &str), Some(&Any::Bool(true)));
        assert_eq!(attrs.get("italic" as &str), Some(&Any::Bool(true)));
        assert!(!attrs.contains_key("underline" as &str));
    }

    #[test]
    fn merge_formatting_preserves_token_keys() {
        let mut token_attrs = HashMap::new();
        token_attrs.insert(Arc::from("tokenType"), Any::String(Arc::from("tab")));
        token_attrs.insert(Arc::from("bold"), Any::Bool(false));

        let mut fmt = HashMap::new();
        fmt.insert(Arc::from("bold"), Any::Bool(true));
        fmt.insert(Arc::from("italic"), Any::Bool(true));

        merge_formatting_attrs(&mut token_attrs, &fmt);

        // Token's "bold" should not be overwritten.
        assert_eq!(token_attrs.get("bold" as &str), Some(&Any::Bool(false)));
        // "italic" should be added.
        assert_eq!(token_attrs.get("italic" as &str), Some(&Any::Bool(true)));
        // "tokenType" should be preserved.
        assert!(token_attrs.contains_key("tokenType" as &str));
    }
}
