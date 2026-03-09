//! Export pipeline: CRDT document state to `.docx` bytes via clone-and-patch.
//!
//! The export does **not** build a new [`Document`] from scratch. Instead it:
//!
//! 1. Re-parses the original `.docx` bytes into a fresh mutable [`Document`].
//! 2. Walks only the paragraphs the CRDT marked as **dirty**.
//! 3. For each dirty paragraph, reads the CRDT rich text via `TextRef::diff()`
//!    and converts the diff chunks back into [`Run`] objects.
//! 4. Replaces the runs on the corresponding paragraph in the clone.
//! 5. Serializes the clone to bytes via [`Document::to_bytes()`].
//!
//! Everything else -- styles, themes, unknown XML elements, non-dirty paragraphs
//! -- survives unchanged because it lives on the re-parsed original objects.
//!
//! [`Document`]: offidized_docx::Document
//! [`Run`]: offidized_docx::run::Run

use std::collections::HashMap;
use std::io::{Cursor, Read, Write};
use std::sync::Arc;

use base64::Engine;
use yrs::types::text::YChange;
use yrs::{Any, Array, GetString, Map, Out, ReadTxn, Text, Transact};
use zip::write::SimpleFileOptions;
use zip::{CompressionMethod, ZipArchive, ZipWriter};

use offidized_docx::paragraph::{LineSpacingRule, ParagraphAlignment};
use offidized_docx::run::{FieldCode, Run};
use offidized_docx::table::Table;
use offidized_docx::{
    BodyItem as DocxBodyItem, Document, NumberingDefinition, NumberingInstance, Paragraph,
};

use super::crdt_doc::CrdtDoc;
use super::tokens::{self, TokenType};

// ---------------------------------------------------------------------------
// Error types
// ---------------------------------------------------------------------------

/// Errors during document export.
#[derive(Debug, thiserror::Error)]
pub enum ExportError {
    /// A general export failure.
    #[error("export failed: {0}")]
    General(String),
    /// The underlying docx layer reported an error.
    #[error("docx error: {0}")]
    Docx(#[from] offidized_docx::DocxError),
}

/// Result type for export operations.
pub type Result<T> = std::result::Result<T, ExportError>;

// ---------------------------------------------------------------------------
// Public API
// ---------------------------------------------------------------------------

/// Export CRDT state to `.docx` bytes using clone-and-patch strategy.
///
/// Clones the original [`Document`] by re-parsing its saved bytes, patches
/// only dirty paragraphs from CRDT state, and serializes the result.
/// Non-dirty content, unknown elements, styles, and themes survive unchanged
/// from the original.
///
/// After a successful export, call [`CrdtDoc::clear_dirty()`] to reset the
/// dirty set.
pub fn export_to_docx(crdt: &CrdtDoc, original_doc: &Document) -> Result<Vec<u8>> {
    let original_bytes = original_doc.to_bytes()?;
    let mut doc = Document::from_bytes(&original_bytes)?;
    if crdt.dirty_paragraphs().is_empty() {
        return Ok(original_bytes);
    }

    let txn = crdt.doc().transact();
    let body = crdt.body();
    let body_len = body.len(&txn);
    let original_body_lookup = collect_original_body_lookup(original_doc);
    let mut numbering_ctx = NumberingExportContext::default();
    let mut image_ctx = ExportImageContext::default();
    let mut new_body_items = Vec::<(u32, NewBodyItem)>::new();

    for i in 0..body_len {
        let Some(entry) = body.get(&txn, i) else {
            continue;
        };
        let Ok(map_ref) = entry.cast::<yrs::MapRef>() else {
            continue;
        };
        let Some(body_item_type) = read_body_item_type(&map_ref, &txn) else {
            continue;
        };
        let Some(body_item_id) = read_body_item_id(&map_ref, &txn) else {
            continue;
        };
        let is_dirty = crdt
            .dirty_paragraphs()
            .iter()
            .any(|pid| pid.to_string() == body_item_id);
        let original_body_position = crdt
            .para_index_map()
            .iter()
            .find(|(pid, _)| pid.to_string() == body_item_id)
            .map(|(_, idx)| idx)
            .copied();

        match (body_item_type.as_str(), original_body_position) {
            ("paragraph", Some(body_pos)) if is_dirty => {
                let Some(OriginalBodyItemRef::Paragraph(paragraph_index)) =
                    original_body_lookup.get(&body_pos).copied()
                else {
                    continue;
                };
                let Some(text_out) = map_ref.get(&txn, "text") else {
                    continue;
                };
                let Ok(text_ref) = text_out.cast::<yrs::TextRef>() else {
                    continue;
                };
                let runs = crdt_text_to_runs(&text_ref, &txn, crdt, &mut doc, &mut image_ctx)?;
                let numbering =
                    resolve_paragraph_numbering(&mut doc, &mut numbering_ctx, &map_ref, &txn);
                if let Some(paragraph) = doc.paragraphs_mut().get_mut(paragraph_index) {
                    paragraph.set_runs(runs);
                    apply_paragraph_properties(paragraph, numbering, &map_ref, &txn);
                }
            }
            ("table", Some(body_pos)) if is_dirty => {
                let Some(OriginalBodyItemRef::Table(table_index)) =
                    original_body_lookup.get(&body_pos).copied()
                else {
                    continue;
                };
                let base_table = doc.tables().get(table_index);
                let exported_table = crdt_table_to_docx(&map_ref, &txn, base_table);
                if let Some(table) = doc.tables_mut().get_mut(table_index) {
                    *table = exported_table;
                }
            }
            ("paragraph", None) => {
                let Some(text_out) = map_ref.get(&txn, "text") else {
                    continue;
                };
                let Ok(text_ref) = text_out.cast::<yrs::TextRef>() else {
                    continue;
                };
                let runs = crdt_text_to_runs(&text_ref, &txn, crdt, &mut doc, &mut image_ctx)?;
                let mut paragraph = Paragraph::new();
                paragraph.set_runs(runs);
                let numbering =
                    resolve_paragraph_numbering(&mut doc, &mut numbering_ctx, &map_ref, &txn);
                apply_paragraph_properties(&mut paragraph, numbering, &map_ref, &txn);
                new_body_items.push((i, NewBodyItem::Paragraph(paragraph)));
            }
            ("table", None) => {
                let table = crdt_table_to_docx(&map_ref, &txn, None);
                new_body_items.push((i, NewBodyItem::Table(table)));
            }
            _ => {}
        }
    }

    for (crdt_idx, new_item) in new_body_items {
        let body_pos =
            find_insert_position(crdt, &txn, &body, crdt_idx, &doc, &original_body_lookup);
        match new_item {
            NewBodyItem::Paragraph(paragraph) => doc.insert_paragraph_at(body_pos, paragraph),
            NewBodyItem::Table(table) => doc.insert_table_at(body_pos, table),
        }
    }

    let bytes = doc.to_bytes()?;
    patch_numbering_package(bytes, &doc)
}

// ---------------------------------------------------------------------------
// Insertion position helpers
// ---------------------------------------------------------------------------

#[derive(Debug, Clone, Copy)]
enum OriginalBodyItemRef {
    Paragraph(usize),
    Table(usize),
}

enum NewBodyItem {
    Paragraph(Paragraph),
    Table(Table),
}

fn collect_original_body_lookup(doc: &Document) -> HashMap<usize, OriginalBodyItemRef> {
    let mut lookup = HashMap::new();
    let mut paragraph_index = 0_usize;
    let mut table_index = 0_usize;
    for (body_pos, item) in doc.body_items().enumerate() {
        match item {
            DocxBodyItem::Paragraph(_) => {
                lookup.insert(body_pos, OriginalBodyItemRef::Paragraph(paragraph_index));
                paragraph_index += 1;
            }
            DocxBodyItem::Table(_) => {
                lookup.insert(body_pos, OriginalBodyItemRef::Table(table_index));
                table_index += 1;
            }
        }
    }
    lookup
}

fn read_body_item_type(map_ref: &yrs::MapRef, txn: &impl ReadTxn) -> Option<String> {
    map_ref.get(txn, "type").and_then(|value| match value {
        Out::Any(Any::String(value)) => Some(value.to_string()),
        _ => None,
    })
}

fn read_body_item_id(map_ref: &yrs::MapRef, txn: &impl ReadTxn) -> Option<String> {
    map_ref.get(txn, "id").and_then(|value| match value {
        Out::Any(Any::String(value)) => Some(value.to_string()),
        _ => None,
    })
}

fn crdt_table_to_docx(
    map_ref: &yrs::MapRef,
    txn: &impl ReadTxn,
    base_table: Option<&Table>,
) -> Table {
    let rows = map_ref
        .get(txn, "rows")
        .and_then(|value| match value {
            Out::Any(Any::Number(value)) => Some(value.round().max(0.0) as usize),
            _ => None,
        })
        .or_else(|| base_table.map(Table::rows))
        .unwrap_or(0);
    let columns = map_ref
        .get(txn, "columns")
        .and_then(|value| match value {
            Out::Any(Any::Number(value)) => Some(value.round().max(0.0) as usize),
            _ => None,
        })
        .or_else(|| base_table.map(Table::columns))
        .unwrap_or(0);

    let mut table = match base_table {
        Some(existing) if existing.rows() == rows && existing.columns() == columns => {
            existing.clone()
        }
        _ => Table::new(rows, columns),
    };

    for row in 0..rows {
        for col in 0..columns {
            let key = format!("cell:{row}:{col}");
            let text = match map_ref.get(txn, key.as_str()) {
                Some(out) => match out.cast::<yrs::TextRef>() {
                    Ok(text_ref) => text_ref.get_string(txn),
                    Err(_) => String::new(),
                },
                None => String::new(),
            };
            let _ = table.set_cell_text(row, col, text);
        }
    }

    table
}

/// Find the body position at which a new body item should be inserted.
fn find_insert_position(
    crdt: &CrdtDoc,
    txn: &impl ReadTxn,
    body: &yrs::ArrayRef,
    new_body_crdt_idx: u32,
    doc: &Document,
    original_body_lookup: &HashMap<usize, OriginalBodyItemRef>,
) -> usize {
    let mut prior_new_items = 0_usize;
    for i in (0..new_body_crdt_idx).rev() {
        let Some(entry) = body.get(txn, i) else {
            continue;
        };
        let Ok(map_ref) = entry.cast::<yrs::MapRef>() else {
            continue;
        };
        let Some(id_str) = map_ref.get(txn, "id").and_then(|v| match v {
            Out::Any(Any::String(s)) => Some(s),
            _ => None,
        }) else {
            continue;
        };

        if let Some(&orig_idx) = crdt
            .para_index_map()
            .iter()
            .find(|(pid, _)| pid.to_string() == id_str.as_ref())
            .map(|(_, idx)| idx)
        {
            let base_position = match original_body_lookup.get(&orig_idx).copied() {
                Some(OriginalBodyItemRef::Paragraph(paragraph_index)) => doc
                    .body_position_of_paragraph(paragraph_index)
                    .map(|body_pos| body_pos + 1)
                    .unwrap_or(orig_idx + 1),
                Some(OriginalBodyItemRef::Table(table_index)) => doc
                    .body_position_of_table(table_index)
                    .map(|body_pos| body_pos + 1)
                    .unwrap_or(orig_idx + 1),
                None => orig_idx + 1,
            };
            return base_position + prior_new_items;
        }

        if read_body_item_id(&map_ref, txn).is_some() {
            prior_new_items += 1;
        }
    }
    prior_new_items
}

#[derive(Default)]
struct NumberingExportContext {
    local_to_actual: HashMap<(String, u32), u32>,
}

#[derive(Default)]
struct ExportImageContext {
    local_to_actual: HashMap<String, usize>,
}

fn apply_paragraph_properties(
    paragraph: &mut Paragraph,
    numbering: Option<(u32, u8)>,
    map_ref: &yrs::MapRef,
    txn: &impl ReadTxn,
) {
    match map_ref.get(txn, "styleId") {
        Some(Out::Any(Any::String(style_id))) => paragraph.set_style_id(style_id.as_ref()),
        _ => paragraph.clear_style_id(),
    }

    let heading_level = map_ref
        .get(txn, "headingLevel")
        .and_then(|value| match value {
            Out::Any(Any::Number(level)) => Some(level.round().clamp(1.0, 9.0) as u8),
            _ => None,
        });
    if let Some(level) = heading_level {
        paragraph.set_style_id(format!("Heading{level}"));
    }

    match paragraph_alignment_from_crdt(map_ref, txn) {
        Some(alignment) => paragraph.set_alignment(alignment),
        None => paragraph.clear_alignment(),
    }

    match paragraph_twips_attr(map_ref, txn, "spacingBeforeTwips") {
        Some(value) => paragraph.set_spacing_before_twips(value),
        None => paragraph.clear_spacing_before_twips(),
    }

    match paragraph_twips_attr(map_ref, txn, "spacingAfterTwips") {
        Some(value) => paragraph.set_spacing_after_twips(value),
        None => paragraph.clear_spacing_after_twips(),
    }

    match paragraph_twips_attr(map_ref, txn, "lineSpacingTwips") {
        Some(value) => paragraph.set_line_spacing_twips(value),
        None => paragraph.clear_line_spacing_twips(),
    }

    match paragraph_line_spacing_rule_from_crdt(map_ref, txn) {
        Some(rule) => paragraph.set_line_spacing_rule(rule),
        None => paragraph.clear_line_spacing_rule(),
    }

    match paragraph_twips_attr(map_ref, txn, "indentLeftTwips") {
        Some(value) => paragraph.set_indent_left_twips(value),
        None => paragraph.clear_indent_left_twips(),
    }

    match paragraph_twips_attr(map_ref, txn, "indentFirstLineTwips") {
        Some(value) => {
            paragraph.clear_indent_hanging_twips();
            paragraph.set_indent_first_line_twips(value);
        }
        None => paragraph.clear_indent_first_line_twips(),
    }

    match numbering {
        Some((num_id, level)) => paragraph.set_numbering(num_id, level),
        _ => paragraph.clear_numbering(),
    }
}

fn paragraph_alignment_from_crdt(
    map_ref: &yrs::MapRef,
    txn: &impl ReadTxn,
) -> Option<ParagraphAlignment> {
    map_ref.get(txn, "alignment").and_then(|value| match value {
        Out::Any(Any::String(alignment)) => match alignment.trim().to_ascii_lowercase().as_str() {
            "left" => Some(ParagraphAlignment::Left),
            "center" => Some(ParagraphAlignment::Center),
            "right" => Some(ParagraphAlignment::Right),
            "justify" | "justified" => Some(ParagraphAlignment::Justified),
            _ => None,
        },
        _ => None,
    })
}

fn paragraph_line_spacing_rule_from_crdt(
    map_ref: &yrs::MapRef,
    txn: &impl ReadTxn,
) -> Option<LineSpacingRule> {
    map_ref
        .get(txn, "lineSpacingRule")
        .and_then(|value| match value {
            Out::Any(Any::String(rule)) => match rule.trim() {
                "auto" => Some(LineSpacingRule::Auto),
                "exact" => Some(LineSpacingRule::Exact),
                "atLeast" => Some(LineSpacingRule::AtLeast),
                _ => None,
            },
            _ => None,
        })
}

fn paragraph_twips_attr(map_ref: &yrs::MapRef, txn: &impl ReadTxn, key: &str) -> Option<u32> {
    map_ref.get(txn, key).and_then(|value| match value {
        Out::Any(Any::Number(value)) => Some(value.round().max(0.0) as u32),
        _ => None,
    })
}

fn resolve_paragraph_numbering(
    doc: &mut Document,
    numbering_ctx: &mut NumberingExportContext,
    map_ref: &yrs::MapRef,
    txn: &impl ReadTxn,
) -> Option<(u32, u8)> {
    let numbering_kind = map_ref
        .get(txn, "numberingKind")
        .and_then(|value| match value {
            Out::Any(Any::String(kind)) => Some(kind.to_string()),
            _ => None,
        });
    let numbering_num_id = map_ref
        .get(txn, "numberingNumId")
        .and_then(|value| match value {
            Out::Any(Any::Number(num_id)) => Some(num_id.round().max(1.0) as u32),
            _ => None,
        });
    let numbering_ilvl = map_ref
        .get(txn, "numberingIlvl")
        .and_then(|value| match value {
            Out::Any(Any::Number(level)) => Some(level.round().clamp(0.0, 8.0) as u8),
            _ => None,
        });

    match (numbering_kind.as_deref(), numbering_num_id, numbering_ilvl) {
        (Some(kind), Some(local_num_id), Some(level)) => Some((
            resolve_export_numbering_num_id(doc, numbering_ctx, local_num_id, kind),
            level,
        )),
        (None, Some(num_id), Some(level)) => Some((num_id, level)),
        _ => None,
    }
}

fn resolve_export_numbering_num_id(
    doc: &mut Document,
    numbering_ctx: &mut NumberingExportContext,
    local_num_id: u32,
    kind: &str,
) -> u32 {
    let cache_key = (kind.to_string(), local_num_id);
    if let Some(actual) = numbering_ctx.local_to_actual.get(&cache_key) {
        return *actual;
    }

    let abstract_num_id = ensure_numbering_definition_for_kind(doc, kind);
    let actual_num_id = if numbering_instance_matches(doc, local_num_id, abstract_num_id) {
        local_num_id
    } else if doc
        .numbering_instances()
        .iter()
        .all(|instance| instance.num_id() != local_num_id)
    {
        doc.numbering_instances_mut()
            .push(NumberingInstance::new(local_num_id, abstract_num_id));
        local_num_id
    } else {
        let next_num_id = doc
            .numbering_instances()
            .iter()
            .map(|instance| instance.num_id())
            .max()
            .map(|max_id| max_id + 1)
            .unwrap_or(1);
        doc.numbering_instances_mut()
            .push(NumberingInstance::new(next_num_id, abstract_num_id));
        next_num_id
    };

    numbering_ctx
        .local_to_actual
        .insert(cache_key, actual_num_id);
    actual_num_id
}

fn numbering_instance_matches(doc: &Document, num_id: u32, abstract_num_id: u32) -> bool {
    doc.numbering_instances()
        .iter()
        .find(|instance| instance.num_id() == num_id)
        .is_some_and(|instance| instance.abstract_num_id() == abstract_num_id)
}

fn ensure_numbering_definition_for_kind(doc: &mut Document, kind: &str) -> u32 {
    if let Some(existing) = doc.numbering_definitions().iter().find(|definition| {
        definition
            .level(0)
            .map(|level| level.format() == kind)
            .unwrap_or(false)
    }) {
        return existing.abstract_num_id();
    }

    let next_abstract_num_id = doc
        .numbering_definitions()
        .iter()
        .map(|definition| definition.abstract_num_id())
        .max()
        .map(|max_id| max_id + 1)
        .unwrap_or(0);
    let definition = match kind {
        "bullet" => NumberingDefinition::create_bullet(next_abstract_num_id),
        _ => NumberingDefinition::create_numbered(next_abstract_num_id),
    };
    doc.numbering_definitions_mut().push(definition);
    next_abstract_num_id
}

fn patch_numbering_package(bytes: Vec<u8>, doc: &Document) -> Result<Vec<u8>> {
    if doc.numbering_definitions().is_empty() || doc.numbering_instances().is_empty() {
        return Ok(bytes);
    }

    let numbering_xml = serialize_numbering_xml(doc);
    let mut archive = ZipArchive::new(Cursor::new(bytes))
        .map_err(|e| ExportError::General(format!("zip open failed: {e}")))?;
    let mut entries = Vec::<(String, Vec<u8>)>::new();
    let mut content_types = None::<Vec<u8>>;
    let mut document_rels = None::<Vec<u8>>;

    for i in 0..archive.len() {
        let mut file = archive
            .by_index(i)
            .map_err(|e| ExportError::General(format!("zip read failed: {e}")))?;
        let mut data = Vec::new();
        file.read_to_end(&mut data)
            .map_err(|e| ExportError::General(format!("zip extract failed: {e}")))?;
        let name = file.name().to_string();
        match name.as_str() {
            "[Content_Types].xml" => content_types = Some(data),
            "word/_rels/document.xml.rels" => document_rels = Some(data),
            "word/numbering.xml" => {}
            _ => entries.push((name, data)),
        }
    }

    let content_types = ensure_numbering_content_type(
        &String::from_utf8(
            content_types
                .ok_or_else(|| ExportError::General("missing [Content_Types].xml".into()))?,
        )
        .map_err(|e| ExportError::General(format!("content types utf8 failed: {e}")))?,
    );
    let document_rels =
        ensure_numbering_relationship(
            &String::from_utf8(document_rels.ok_or_else(|| {
                ExportError::General("missing word/_rels/document.xml.rels".into())
            })?)
            .map_err(|e| ExportError::General(format!("document rels utf8 failed: {e}")))?,
        );

    let mut out = Cursor::new(Vec::<u8>::new());
    {
        let mut writer = ZipWriter::new(&mut out);
        let options = SimpleFileOptions::default()
            .compression_method(CompressionMethod::Deflated)
            .unix_permissions(0o644);
        for (name, data) in entries {
            writer
                .start_file(name, options)
                .map_err(|e| ExportError::General(format!("zip write failed: {e}")))?;
            writer
                .write_all(&data)
                .map_err(|e| ExportError::General(format!("zip write failed: {e}")))?;
        }
        writer
            .start_file("[Content_Types].xml", options)
            .map_err(|e| ExportError::General(format!("zip write failed: {e}")))?;
        writer
            .write_all(content_types.as_bytes())
            .map_err(|e| ExportError::General(format!("zip write failed: {e}")))?;
        writer
            .start_file("word/_rels/document.xml.rels", options)
            .map_err(|e| ExportError::General(format!("zip write failed: {e}")))?;
        writer
            .write_all(document_rels.as_bytes())
            .map_err(|e| ExportError::General(format!("zip write failed: {e}")))?;
        writer
            .start_file("word/numbering.xml", options)
            .map_err(|e| ExportError::General(format!("zip write failed: {e}")))?;
        writer
            .write_all(numbering_xml.as_bytes())
            .map_err(|e| ExportError::General(format!("zip write failed: {e}")))?;
        writer
            .finish()
            .map_err(|e| ExportError::General(format!("zip finalize failed: {e}")))?;
    }
    Ok(out.into_inner())
}

fn ensure_numbering_content_type(content_types_xml: &str) -> String {
    if content_types_xml.contains("/word/numbering.xml") {
        return content_types_xml.to_string();
    }
    let override_xml = r#"<Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>"#;
    content_types_xml.replacen("</Types>", &format!("{override_xml}</Types>"), 1)
}

fn ensure_numbering_relationship(rels_xml: &str) -> String {
    if rels_xml.contains(WORD_NUMBERING_REL_TYPE) || rels_xml.contains("Target=\"numbering.xml\"") {
        return rels_xml.to_string();
    }
    let relationship_xml = format!(
        r#"<Relationship Id="rIdNumbering" Type="{WORD_NUMBERING_REL_TYPE}" Target="numbering.xml"/>"#
    );
    rels_xml.replacen(
        "</Relationships>",
        &format!("{relationship_xml}</Relationships>"),
        1,
    )
}

const WORD_NUMBERING_REL_TYPE: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering";

fn serialize_numbering_xml(doc: &Document) -> String {
    let mut xml = String::from(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">"#,
    );

    for definition in doc.numbering_definitions() {
        xml.push_str(&format!(
            r#"<w:abstractNum w:abstractNumId="{}">"#,
            definition.abstract_num_id()
        ));
        for level in definition.levels() {
            xml.push_str(&format!(r#"<w:lvl w:ilvl="{}">"#, level.level()));
            xml.push_str(&format!(r#"<w:start w:val="{}"/>"#, level.start()));
            xml.push_str(&format!(
                r#"<w:numFmt w:val="{}"/>"#,
                xml_escape(level.format())
            ));
            xml.push_str(&format!(
                r#"<w:lvlText w:val="{}"/>"#,
                xml_escape(level.text())
            ));
            if let Some(alignment) = level.alignment() {
                xml.push_str(&format!(r#"<w:lvlJc w:val="{}"/>"#, xml_escape(alignment)));
            }
            if level.indent_left_twips().is_some() || level.indent_hanging_twips().is_some() {
                xml.push_str("<w:pPr><w:ind");
                if let Some(left) = level.indent_left_twips() {
                    xml.push_str(&format!(r#" w:left="{left}""#));
                }
                if let Some(hanging) = level.indent_hanging_twips() {
                    xml.push_str(&format!(r#" w:hanging="{hanging}""#));
                }
                xml.push_str("/></w:pPr>");
            }
            xml.push_str("</w:lvl>");
        }
        xml.push_str("</w:abstractNum>");
    }

    for instance in doc.numbering_instances() {
        xml.push_str(&format!(
            r#"<w:num w:numId="{}"><w:abstractNumId w:val="{}"/>"#,
            instance.num_id(),
            instance.abstract_num_id()
        ));
        for override_level in instance.level_overrides() {
            xml.push_str(&format!(
                r#"<w:lvlOverride w:ilvl="{}">"#,
                override_level.level()
            ));
            if let Some(start) = override_level.start_override() {
                xml.push_str(&format!(r#"<w:startOverride w:val="{}"/>"#, start));
            }
            xml.push_str("</w:lvlOverride>");
        }
        xml.push_str("</w:num>");
    }

    xml.push_str("</w:numbering>");
    xml
}

fn xml_escape(value: &str) -> String {
    value
        .replace('&', "&amp;")
        .replace('"', "&quot;")
        .replace('<', "&lt;")
        .replace('>', "&gt;")
        .replace('\'', "&apos;")
}

// ---------------------------------------------------------------------------
// CRDT text -> Vec<Run> conversion
// ---------------------------------------------------------------------------

/// Convert a CRDT [`TextRef`]'s content back into a vector of [`Run`]s.
///
/// Walks the `TextRef::diff()` output. Each diff chunk becomes one or more
/// runs depending on whether it contains sentinel tokens.
fn crdt_text_to_runs(
    text_ref: &yrs::TextRef,
    txn: &impl ReadTxn,
    crdt: &CrdtDoc,
    doc: &mut Document,
    image_ctx: &mut ExportImageContext,
) -> Result<Vec<Run>> {
    let diff = text_ref.diff(txn, YChange::identity);
    let mut runs: Vec<Run> = Vec::new();

    for chunk in &diff {
        let attrs = chunk
            .attributes
            .as_ref()
            .map(|a| a.as_ref().clone())
            .unwrap_or_default();

        // Extract the text content from the diff chunk.
        let text = match &chunk.insert {
            Out::Any(Any::String(s)) => s.as_ref().to_string(),
            _ => continue,
        };

        if let Some(Any::String(_)) = attrs.get(tokens::ATTR_TOKEN_TYPE as &str) {
            if let Some(token) = tokens::attrs_to_token(&attrs) {
                let sentinel_count = text.chars().filter(|c| tokens::is_sentinel(*c)).count();
                let count = sentinel_count.max(1);
                for _ in 0..count {
                    if let Some(run) = token_to_run(&token, &attrs, crdt, doc, txn, image_ctx)? {
                        runs.push(run);
                    }
                }
                continue;
            }
        }

        // Regular text chunk -- build a run with formatting.
        // The text might contain multiple sentinel chars mixed with text
        // (though import guarantees they are separate chunks). Handle both.
        let clean_text = tokens::strip_sentinels(&text);
        if clean_text.is_empty() {
            continue;
        }

        let mut run = Run::new(&clean_text);
        apply_formatting_attrs(&mut run, &attrs);
        runs.push(run);
    }

    Ok(runs)
}

/// Convert a CRDT token back into a [`Run`].
///
/// Returns `None` for token types that cannot be represented as runs in
/// the current MVP (e.g., complex field markers, bookmarks, comments).
fn token_to_run(
    token: &TokenType,
    attrs: &HashMap<Arc<str>, Any>,
    crdt: &CrdtDoc,
    doc: &mut Document,
    txn: &impl ReadTxn,
    image_ctx: &mut ExportImageContext,
) -> Result<Option<Run>> {
    let mut run = Run::new("");
    apply_formatting_attrs(&mut run, attrs);

    match token {
        TokenType::Tab => {
            run.set_has_tab(true);
            Ok(Some(run))
        }
        TokenType::LineBreak { .. } => {
            run.set_has_break(true);
            Ok(Some(run))
        }
        TokenType::FootnoteRef { id } => {
            run.set_footnote_reference_id(*id);
            Ok(Some(run))
        }
        TokenType::EndnoteRef { id } => {
            run.set_endnote_reference_id(*id);
            Ok(Some(run))
        }
        TokenType::FieldSimple { instr, .. } => {
            run.set_field_simple(instr);
            Ok(Some(run))
        }
        TokenType::FieldBegin { .. }
        | TokenType::FieldCode { .. }
        | TokenType::FieldSeparate { .. }
        | TokenType::FieldEnd { .. } => {
            // Complex field markers: cannot round-trip perfectly as a single
            // fldSimple run. For MVP, emit a field_code run for FieldCode
            // chunks and skip the structural markers.
            if let TokenType::FieldCode { instr, .. } = token {
                run.set_field_code(FieldCode::new(instr, ""));
                Ok(Some(run))
            } else {
                // Structural markers (begin/separate/end) are not emitted
                // as separate runs; they are implied by the serializer.
                Ok(None)
            }
        }
        TokenType::BookmarkStart { .. }
        | TokenType::BookmarkEnd { .. }
        | TokenType::CommentStart { .. }
        | TokenType::CommentEnd { .. } => {
            // Bookmark and comment markers are paragraph-level or run-adjacent
            // elements that the original document preserves on non-dirty
            // paragraphs. On dirty paragraphs they are lost in the MVP.
            Ok(None)
        }
        TokenType::InlineImage {
            image_ref,
            width,
            height,
        } => {
            let Some(image_index) =
                resolve_export_image_index(doc, crdt, txn, image_ctx, image_ref)?
            else {
                return Ok(None);
            };
            let mut inline = offidized_docx::InlineImage::new(
                image_index,
                (*width).max(1) as u32,
                (*height).max(1) as u32,
            );
            if let Some((name, description)) = export_image_labels(attrs, crdt, txn, image_ref) {
                if let Some(name) = name {
                    inline.set_name(name);
                }
                if let Some(description) = description {
                    inline.set_description(description);
                }
            }
            run.set_inline_image(inline);
            Ok(Some(run))
        }
        TokenType::Opaque { .. } => {
            // Opaque tokens carry raw XML that we cannot inject back into
            // a Run through the high-level API. Skip for MVP.
            Ok(None)
        }
    }
}

/// Apply CRDT formatting attributes to a [`Run`].
///
/// This is the reverse of `run_formatting_attrs()` in the import module.
fn apply_formatting_attrs(run: &mut Run, attrs: &HashMap<Arc<str>, Any>) {
    if matches!(attrs.get("bold" as &str), Some(Any::Bool(true))) {
        run.set_bold(true);
    }
    if matches!(attrs.get("italic" as &str), Some(Any::Bool(true))) {
        run.set_italic(true);
    }
    if matches!(attrs.get("underline" as &str), Some(Any::Bool(true))) {
        run.set_underline(true);
    }
    if matches!(attrs.get("strike" as &str), Some(Any::Bool(true))) {
        run.set_strikethrough(true);
    }
    if matches!(attrs.get("superscript" as &str), Some(Any::Bool(true))) {
        run.set_superscript(true);
    }
    if matches!(attrs.get("subscript" as &str), Some(Any::Bool(true))) {
        run.set_subscript(true);
    }
    if matches!(attrs.get("smallCaps" as &str), Some(Any::Bool(true))) {
        run.set_small_caps(true);
    }
    if let Some(Any::String(s)) = attrs.get("fontFamily" as &str) {
        run.set_font_family(s.as_ref());
    }
    if let Some(Any::Number(n)) = attrs.get("fontSize" as &str) {
        run.set_font_size_half_points(*n as u16);
    }
    if let Some(Any::String(s)) = attrs.get("color" as &str) {
        run.set_color(s.as_ref());
    }
    if let Some(Any::String(s)) = attrs.get("highlight" as &str) {
        run.set_highlight_color(s.as_ref());
    }
    if let Some(Any::String(s)) = attrs.get("hyperlink" as &str) {
        run.set_hyperlink(s.as_ref());
    }
}

fn export_image_labels(
    attrs: &HashMap<Arc<str>, Any>,
    crdt: &CrdtDoc,
    txn: &impl ReadTxn,
    image_ref: &str,
) -> Option<(Option<String>, Option<String>)> {
    let attrs_name = attrs.get("name").and_then(|value| match value {
        Any::String(name) => Some(name.to_string()),
        _ => None,
    });
    let attrs_description = attrs.get("description").and_then(|value| match value {
        Any::String(description) => Some(description.to_string()),
        _ => None,
    });
    if attrs_name.is_some() || attrs_description.is_some() {
        return Some((attrs_name, attrs_description));
    }

    let images = crdt.images_map();
    let map_ref = images.get(txn, image_ref)?.cast::<yrs::MapRef>().ok()?;
    let name = map_ref.get(txn, "name").and_then(|value| match value {
        Out::Any(Any::String(name)) => Some(name.to_string()),
        _ => None,
    });
    let description = map_ref
        .get(txn, "description")
        .and_then(|value| match value {
            Out::Any(Any::String(description)) => Some(description.to_string()),
            _ => None,
        });
    Some((name, description))
}

fn resolve_export_image_index(
    doc: &mut Document,
    crdt: &CrdtDoc,
    txn: &impl ReadTxn,
    image_ctx: &mut ExportImageContext,
    image_ref: &str,
) -> Result<Option<usize>> {
    if let Some(index) = image_ref
        .strip_prefix("img:")
        .filter(|value| !value.starts_with("local:"))
        .and_then(|value| value.parse::<usize>().ok())
    {
        return Ok((index < doc.images().len()).then_some(index));
    }

    if let Some(&index) = image_ctx.local_to_actual.get(image_ref) {
        return Ok(Some(index));
    }

    let images = crdt.images_map();
    let Some(map_ref) = images
        .get(txn, image_ref)
        .and_then(|value| value.cast::<yrs::MapRef>().ok())
    else {
        return Ok(None);
    };
    let Some(content_type) = map_ref
        .get(txn, "contentType")
        .and_then(|value| match value {
            Out::Any(Any::String(content_type)) => Some(content_type.to_string()),
            _ => None,
        })
    else {
        return Ok(None);
    };
    let bytes = if let Some(bytes) = crdt.image_blobs().get(image_ref) {
        bytes.clone()
    } else if let Some(Out::Any(Any::String(data_uri))) = map_ref.get(txn, "dataUri") {
        decode_data_uri(data_uri.as_ref())
            .map(|(_, bytes)| bytes)
            .map_err(ExportError::General)?
    } else {
        return Ok(None);
    };

    let index = doc.add_image(bytes, content_type);
    image_ctx
        .local_to_actual
        .insert(image_ref.to_string(), index);
    Ok(Some(index))
}

fn decode_data_uri(data_uri: &str) -> std::result::Result<(String, Vec<u8>), String> {
    let Some((header, body)) = data_uri.split_once(',') else {
        return Err("image data URI missing comma".to_string());
    };
    if !header.starts_with("data:") || !header.ends_with(";base64") {
        return Err("image data URI must be base64-encoded".to_string());
    }
    let content_type = header
        .trim_start_matches("data:")
        .trim_end_matches(";base64")
        .trim();
    if content_type.is_empty() {
        return Err("image data URI missing content type".to_string());
    }
    let bytes = base64::engine::general_purpose::STANDARD
        .decode(body)
        .map_err(|err| format!("image base64 decode failed: {err}"))?;
    Ok((content_type.to_string(), bytes))
}

// ---------------------------------------------------------------------------
// Tests
// ---------------------------------------------------------------------------

#[cfg(test)]
#[allow(clippy::unwrap_used, clippy::expect_used, clippy::panic)]
mod tests {
    use super::*;
    use crate::editor::import::import_document;
    use crate::editor::tokens::SENTINEL;
    use offidized_docx::Document;
    use yrs::{GetString, Text, Transact};
    use zip::ZipArchive;

    fn empty_export_context() -> (Document, CrdtDoc, ExportImageContext) {
        (
            Document::new(),
            CrdtDoc::new(),
            ExportImageContext::default(),
        )
    }

    #[test]
    fn export_unchanged_roundtrips() {
        let mut doc = Document::new();
        doc.add_paragraph("Hello world");
        doc.add_paragraph("Second paragraph");

        let mut crdt = CrdtDoc::new();
        import_document(&doc, &mut crdt).expect("import");

        // No edits => no dirty paragraphs.
        assert!(crdt.dirty_paragraphs().is_empty());

        let bytes = export_to_docx(&crdt, &doc).expect("export");
        let reopened = Document::from_bytes(&bytes).expect("reopen");
        assert_eq!(reopened.paragraphs().len(), 2);
        assert_eq!(reopened.paragraphs()[0].text(), "Hello world");
        assert_eq!(reopened.paragraphs()[1].text(), "Second paragraph");
    }

    #[test]
    fn export_with_dirty_paragraph_patches_text() {
        let mut doc = Document::new();
        doc.add_paragraph("Original text");
        doc.add_paragraph("Untouched");

        let mut crdt = CrdtDoc::new();
        import_document(&doc, &mut crdt).expect("import");

        // Simulate an edit: replace the first paragraph's text.
        let txn_r = crdt.doc().transact();
        let body = crdt.body();
        let entry = body.get(&txn_r, 0).expect("body[0]");
        let map_ref = entry.cast::<yrs::MapRef>().expect("is map");
        let para_id_str = match map_ref.get(&txn_r, "id").expect("has id") {
            Out::Any(Any::String(s)) => s.to_string(),
            _ => panic!("id is not string"),
        };
        let text_val = map_ref.get(&txn_r, "text").expect("has text");
        let text_ref = text_val.cast::<yrs::TextRef>().expect("is TextRef");
        let old_len = text_ref.get_string(&txn_r).len() as u32;
        drop(txn_r);

        // Write new content.
        {
            let mut txn = crdt.doc().transact_mut();
            text_ref.remove_range(&mut txn, 0, old_len);
            text_ref.insert(&mut txn, 0, "Modified text");
        }

        // Mark dirty via the para ID string.
        let para_id = crdt
            .para_index_map()
            .keys()
            .find(|pid| pid.to_string() == para_id_str)
            .expect("find para id")
            .clone();
        crdt.mark_dirty(&para_id);

        let bytes = export_to_docx(&crdt, &doc).expect("export");
        let reopened = Document::from_bytes(&bytes).expect("reopen");
        assert_eq!(reopened.paragraphs().len(), 2);
        assert_eq!(reopened.paragraphs()[0].text(), "Modified text");
        assert_eq!(reopened.paragraphs()[1].text(), "Untouched");
    }

    #[test]
    fn crdt_text_to_runs_plain_text() {
        let yrs_doc = yrs::Doc::new();
        let text_ref = yrs_doc.get_or_insert_text("test");
        {
            let mut txn = yrs_doc.transact_mut();
            text_ref.insert(&mut txn, 0, "Hello world");
        }
        let txn = yrs_doc.transact();
        let (mut doc, crdt, mut image_ctx) = empty_export_context();
        let runs = crdt_text_to_runs(&text_ref, &txn, &crdt, &mut doc, &mut image_ctx)
            .expect("plain text runs");
        assert_eq!(runs.len(), 1);
        assert_eq!(runs[0].text(), "Hello world");
        assert!(!runs[0].is_bold());
    }

    #[test]
    fn crdt_text_to_runs_formatted_text() {
        let yrs_doc = yrs::Doc::new();
        let text_ref = yrs_doc.get_or_insert_text("test");
        {
            let mut txn = yrs_doc.transact_mut();
            let mut attrs = HashMap::new();
            attrs.insert(Arc::from("bold"), Any::Bool(true));
            attrs.insert(Arc::from("italic"), Any::Bool(true));
            text_ref.insert_with_attributes(&mut txn, 0, "bold italic", attrs);
        }
        let txn = yrs_doc.transact();
        let (mut doc, crdt, mut image_ctx) = empty_export_context();
        let runs = crdt_text_to_runs(&text_ref, &txn, &crdt, &mut doc, &mut image_ctx)
            .expect("formatted runs");
        assert_eq!(runs.len(), 1);
        assert_eq!(runs[0].text(), "bold italic");
        assert!(runs[0].is_bold());
        assert!(runs[0].is_italic());
    }

    #[test]
    fn crdt_text_to_runs_with_tab_token() {
        // Use the import pipeline to build realistic CRDT content
        // (avoids yrs position-insert quirks with standalone text refs).
        let mut doc = Document::new();
        let para = doc.add_paragraph("");
        if let Some(run) = para.runs_mut().first_mut() {
            run.set_has_tab(true);
        }

        let mut crdt = CrdtDoc::new();
        import_document(&doc, &mut crdt).expect("import");

        // Read back the text ref from the first paragraph in the CRDT body.
        let txn = crdt.doc().transact();
        let body = crdt.body();
        let entry = body.get(&txn, 0).expect("body[0]");
        let map_ref = entry.cast::<yrs::MapRef>().expect("is map");
        let text_out = map_ref.get(&txn, "text").expect("has text");
        let text_ref = text_out.cast::<yrs::TextRef>().expect("is TextRef");

        let mut export_doc = Document::new();
        let mut image_ctx = ExportImageContext::default();
        let runs = crdt_text_to_runs(&text_ref, &txn, &crdt, &mut export_doc, &mut image_ctx)
            .expect("tab token runs");
        assert_eq!(runs.len(), 1);
        assert!(runs[0].has_tab());
        assert_eq!(runs[0].text(), "");
    }

    #[test]
    fn crdt_text_to_runs_with_footnote_ref() {
        let yrs_doc = yrs::Doc::new();
        let text_ref = yrs_doc.get_or_insert_text("test");
        {
            let mut txn = yrs_doc.transact_mut();
            let token = TokenType::FootnoteRef { id: 7 };
            let attrs = tokens::token_to_attrs(&token);
            text_ref.insert_with_attributes(&mut txn, 0, &SENTINEL.to_string(), attrs);
        }
        let txn = yrs_doc.transact();
        let (mut doc, crdt, mut image_ctx) = empty_export_context();
        let runs = crdt_text_to_runs(&text_ref, &txn, &crdt, &mut doc, &mut image_ctx)
            .expect("footnote runs");
        assert_eq!(runs.len(), 1);
        assert_eq!(runs[0].footnote_reference_id(), Some(7));
    }

    #[test]
    fn crdt_text_to_runs_mixed_formatting() {
        let yrs_doc = yrs::Doc::new();
        let text_ref = yrs_doc.get_or_insert_text("test");
        {
            let mut txn = yrs_doc.transact_mut();
            // Plain text
            text_ref.insert(&mut txn, 0, "plain ");
            // Bold text
            let mut bold_attrs = HashMap::new();
            bold_attrs.insert(Arc::from("bold"), Any::Bool(true));
            text_ref.insert_with_attributes(&mut txn, 6, "bold", bold_attrs);
        }
        let txn = yrs_doc.transact();
        let (mut doc, crdt, mut image_ctx) = empty_export_context();
        let runs = crdt_text_to_runs(&text_ref, &txn, &crdt, &mut doc, &mut image_ctx)
            .expect("mixed formatting runs");
        assert_eq!(runs.len(), 2);
        assert_eq!(runs[0].text(), "plain ");
        assert!(!runs[0].is_bold());
        assert_eq!(runs[1].text(), "bold");
        assert!(runs[1].is_bold());
    }

    #[test]
    fn apply_formatting_attrs_all_properties() {
        let mut attrs = HashMap::new();
        attrs.insert(Arc::from("bold"), Any::Bool(true));
        attrs.insert(Arc::from("italic"), Any::Bool(true));
        attrs.insert(Arc::from("underline"), Any::Bool(true));
        attrs.insert(Arc::from("strike"), Any::Bool(true));
        attrs.insert(Arc::from("superscript"), Any::Bool(true));
        attrs.insert(Arc::from("smallCaps"), Any::Bool(true));
        attrs.insert(Arc::from("fontFamily"), Any::String(Arc::from("Calibri")));
        attrs.insert(Arc::from("fontSize"), Any::Number(28.0));
        attrs.insert(Arc::from("color"), Any::String(Arc::from("FF0000")));
        attrs.insert(Arc::from("highlight"), Any::String(Arc::from("yellow")));
        attrs.insert(
            Arc::from("hyperlink"),
            Any::String(Arc::from("https://example.com")),
        );

        let mut run = Run::new("test");
        apply_formatting_attrs(&mut run, &attrs);

        assert!(run.is_bold());
        assert!(run.is_italic());
        assert!(run.is_underline());
        assert!(run.is_strikethrough());
        assert!(run.is_superscript());
        assert!(run.is_small_caps());
        assert_eq!(run.font_family(), Some("Calibri"));
        assert_eq!(run.font_size_half_points(), Some(28));
        assert_eq!(run.color(), Some("FF0000"));
        assert_eq!(run.highlight_color(), Some("yellow"));
        assert_eq!(run.hyperlink(), Some("https://example.com"));
    }

    #[test]
    fn token_to_run_tab() {
        let token = TokenType::Tab;
        let attrs = tokens::token_to_attrs(&token);
        let (mut doc, crdt, mut image_ctx) = empty_export_context();
        let txn = crdt.doc().transact();
        let run = token_to_run(&token, &attrs, &crdt, &mut doc, &txn, &mut image_ctx)
            .expect("tab export")
            .expect("tab produces run");
        assert!(run.has_tab());
        assert_eq!(run.text(), "");
    }

    #[test]
    fn token_to_run_line_break() {
        let token = TokenType::LineBreak { break_type: None };
        let attrs = tokens::token_to_attrs(&token);
        let (mut doc, crdt, mut image_ctx) = empty_export_context();
        let txn = crdt.doc().transact();
        let run = token_to_run(&token, &attrs, &crdt, &mut doc, &txn, &mut image_ctx)
            .expect("line break export")
            .expect("break produces run");
        assert!(run.has_break());
    }

    #[test]
    fn token_to_run_field_simple() {
        let token = TokenType::FieldSimple {
            field_type: "PAGE".to_string(),
            instr: "PAGE \\* MERGEFORMAT".to_string(),
            presentation: "3".to_string(),
        };
        let attrs = tokens::token_to_attrs(&token);
        let (mut doc, crdt, mut image_ctx) = empty_export_context();
        let txn = crdt.doc().transact();
        let run = token_to_run(&token, &attrs, &crdt, &mut doc, &txn, &mut image_ctx)
            .expect("field export")
            .expect("field produces run");
        assert_eq!(run.field_simple(), Some("PAGE \\* MERGEFORMAT"));
    }

    #[test]
    fn token_to_run_inline_image_uses_original_doc_image() {
        let mut doc = Document::new();
        let index = doc.add_image(vec![1_u8, 2, 3], "image/png");
        let crdt = CrdtDoc::new();
        let mut image_ctx = ExportImageContext::default();
        let token = TokenType::InlineImage {
            image_ref: format!("img:{index}"),
            width: 914400,
            height: 457200,
        };
        let attrs = tokens::token_to_attrs(&token);
        let txn = crdt.doc().transact();
        let run = token_to_run(&token, &attrs, &crdt, &mut doc, &txn, &mut image_ctx)
            .expect("inline image export")
            .expect("image run");
        let inline = run.inline_image().expect("inline image");
        assert_eq!(inline.image_index(), index);
        assert_eq!(inline.width_emu(), 914400);
    }

    #[test]
    fn token_to_run_opaque_skipped() {
        let token = TokenType::Opaque {
            opaque_id: "some-uuid".to_string(),
            xml: "<w:sym/>".to_string(),
        };
        let attrs = tokens::token_to_attrs(&token);
        let (mut doc, crdt, mut image_ctx) = empty_export_context();
        let txn = crdt.doc().transact();
        assert!(
            token_to_run(&token, &attrs, &crdt, &mut doc, &txn, &mut image_ctx)
                .expect("opaque export")
                .is_none()
        );
    }

    #[test]
    fn export_bold_paragraph_preserves_formatting() {
        let mut doc = Document::new();
        let para = doc.add_paragraph("");
        if let Some(run) = para.runs_mut().first_mut() {
            run.set_text("Bold text");
            run.set_bold(true);
        }

        let mut crdt = CrdtDoc::new();
        import_document(&doc, &mut crdt).expect("import");

        // Mark first paragraph dirty to force re-export from CRDT.
        let txn = crdt.doc().transact();
        let body = crdt.body();
        let entry = body.get(&txn, 0).expect("body[0]");
        let map_ref = entry.cast::<yrs::MapRef>().expect("is map");
        let para_id_str = match map_ref.get(&txn, "id").expect("has id") {
            Out::Any(Any::String(s)) => s.to_string(),
            _ => panic!("id is not string"),
        };
        drop(txn);

        let para_id = crdt
            .para_index_map()
            .keys()
            .find(|pid| pid.to_string() == para_id_str)
            .expect("find para id")
            .clone();
        crdt.mark_dirty(&para_id);

        let bytes = export_to_docx(&crdt, &doc).expect("export");
        let reopened = Document::from_bytes(&bytes).expect("reopen");
        assert_eq!(reopened.paragraphs().len(), 1);
        let runs = reopened.paragraphs()[0].runs();
        assert_eq!(runs.len(), 1);
        assert_eq!(runs[0].text(), "Bold text");
        assert!(runs[0].is_bold());
    }

    #[test]
    fn export_new_paragraph_from_split() {
        use crate::editor::para_id::ParaId;
        use yrs::{Array, Map, MapPrelim, TextPrelim};

        // 1. Create a document with one paragraph: "Hello".
        let mut doc = Document::new();
        doc.add_paragraph("Hello");

        let mut crdt = CrdtDoc::new();
        import_document(&doc, &mut crdt).expect("import");

        // 2. Simulate InsertParagraph at offset 5 (end of "Hello").
        //    This splits into "Hello" (existing) and a new empty paragraph.
        let body = crdt.body();

        // Mark the first paragraph dirty (its text didn't actually change,
        // but InsertParagraph marks both sides dirty).
        let txn_r = crdt.doc().transact();
        let entry0 = body.get(&txn_r, 0).expect("body[0]");
        let map0 = entry0.cast::<yrs::MapRef>().expect("is map");
        let para0_id_str = match map0.get(&txn_r, "id").expect("has id") {
            Out::Any(Any::String(s)) => s.to_string(),
            _ => panic!("id is not string"),
        };
        drop(txn_r);

        let para0_id = crdt
            .para_index_map()
            .keys()
            .find(|pid| pid.to_string() == para0_id_str)
            .expect("find para0 id")
            .clone();
        crdt.mark_dirty(&para0_id);

        // Create new paragraph in the CRDT body at index 1.
        let new_id = ParaId::new();
        {
            let mut txn = crdt.doc().transact_mut();
            let prelim = MapPrelim::from([
                ("type".to_string(), Any::String(Arc::from("paragraph"))),
                ("id".to_string(), Any::String(Arc::from(new_id.to_string()))),
            ]);
            body.insert(&mut txn, 1, prelim);

            // Create TextRef on the new paragraph.
            let new_entry = body.get(&txn, 1).expect("body[1]");
            let new_map = new_entry.cast::<yrs::MapRef>().expect("is map");
            new_map.insert(&mut txn, "text", TextPrelim::new(""));
            let text_val = new_map.get(&txn, "text").expect("has text");
            let text_ref = text_val.cast::<yrs::TextRef>().expect("is TextRef");

            // 3. Insert "World" into the new paragraph.
            text_ref.insert(&mut txn, 0, "World");
        }

        // Register the new paragraph (dirty, no original index).
        crdt.register_new_paragraph(new_id);

        // 4. Export and reload.
        let bytes = export_to_docx(&crdt, &doc).expect("export");
        let reopened = Document::from_bytes(&bytes).expect("reopen");

        // 5. Verify both paragraphs are present with correct text.
        assert_eq!(reopened.paragraphs().len(), 2);
        assert_eq!(reopened.paragraphs()[0].text(), "Hello");
        assert_eq!(reopened.paragraphs()[1].text(), "World");
    }

    #[test]
    fn export_preserves_existing_bullet_numbering_on_dirty_paragraph() {
        let mut doc = Document::new();
        doc.add_bulleted_paragraph("Bullet item");

        let mut crdt = CrdtDoc::new();
        import_document(&doc, &mut crdt).expect("import");

        let txn = crdt.doc().transact();
        let body = crdt.body();
        let entry = body.get(&txn, 0).expect("body[0]");
        let map_ref = entry.cast::<yrs::MapRef>().expect("is map");
        let para_id_str = match map_ref.get(&txn, "id").expect("has id") {
            Out::Any(Any::String(s)) => s.to_string(),
            _ => panic!("id is not string"),
        };
        drop(txn);

        let para_id = crdt
            .para_index_map()
            .keys()
            .find(|pid| pid.to_string() == para_id_str)
            .expect("find para id")
            .clone();
        crdt.mark_dirty(&para_id);

        let bytes = export_to_docx(&crdt, &doc).expect("export");
        let reopened = Document::from_bytes(&bytes).expect("reopen");
        let paragraph = &reopened.paragraphs()[0];
        assert_eq!(paragraph.text(), "Bullet item");
        assert_eq!(paragraph.numbering_ilvl(), Some(0));
        assert!(paragraph.numbering_num_id().is_some());
    }

    #[test]
    fn export_preserves_existing_decimal_numbering_on_dirty_paragraph() {
        let mut doc = Document::new();
        doc.add_numbered_paragraph("Numbered item");

        let mut crdt = CrdtDoc::new();
        import_document(&doc, &mut crdt).expect("import");

        let txn = crdt.doc().transact();
        let body = crdt.body();
        let entry = body.get(&txn, 0).expect("body[0]");
        let map_ref = entry.cast::<yrs::MapRef>().expect("is map");
        let para_id_str = match map_ref.get(&txn, "id").expect("has id") {
            Out::Any(Any::String(s)) => s.to_string(),
            _ => panic!("id is not string"),
        };
        drop(txn);

        let para_id = crdt
            .para_index_map()
            .keys()
            .find(|pid| pid.to_string() == para_id_str)
            .expect("find para id")
            .clone();
        crdt.mark_dirty(&para_id);

        let bytes = export_to_docx(&crdt, &doc).expect("export");
        let reopened = Document::from_bytes(&bytes).expect("reopen");
        let paragraph = &reopened.paragraphs()[0];
        assert_eq!(paragraph.text(), "Numbered item");
        assert_eq!(paragraph.numbering_ilvl(), Some(0));
        assert!(paragraph.numbering_num_id().is_some());
    }

    #[test]
    fn export_writes_numbering_package_parts_when_lists_exist() {
        let mut doc = Document::new();
        doc.add_bulleted_paragraph("Bullet item");

        let mut crdt = CrdtDoc::new();
        import_document(&doc, &mut crdt).expect("import");

        let txn = crdt.doc().transact();
        let body = crdt.body();
        let entry = body.get(&txn, 0).expect("body[0]");
        let map_ref = entry.cast::<yrs::MapRef>().expect("is map");
        let para_id_str = match map_ref.get(&txn, "id").expect("has id") {
            Out::Any(Any::String(s)) => s.to_string(),
            _ => panic!("id is not string"),
        };
        drop(txn);

        let para_id = crdt
            .para_index_map()
            .keys()
            .find(|pid| pid.to_string() == para_id_str)
            .expect("find para id")
            .clone();
        crdt.mark_dirty(&para_id);

        let bytes = export_to_docx(&crdt, &doc).expect("export");
        let mut archive = ZipArchive::new(Cursor::new(bytes)).expect("zip");

        let mut numbering_xml = String::new();
        archive
            .by_name("word/numbering.xml")
            .expect("numbering part")
            .read_to_string(&mut numbering_xml)
            .expect("read numbering");
        assert!(numbering_xml.contains("<w:numbering"));
        assert!(numbering_xml.contains("w:abstractNum"));
        assert!(numbering_xml.contains("w:num"));

        let mut content_types = String::new();
        archive
            .by_name("[Content_Types].xml")
            .expect("content types")
            .read_to_string(&mut content_types)
            .expect("read content types");
        assert!(content_types.contains("/word/numbering.xml"));

        let mut rels = String::new();
        archive
            .by_name("word/_rels/document.xml.rels")
            .expect("document rels")
            .read_to_string(&mut rels)
            .expect("read rels");
        assert!(rels.contains("relationships/numbering"));
        assert!(rels.contains("Target=\"numbering.xml\""));
    }
}
