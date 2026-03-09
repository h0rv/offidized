//! CRDT-to-view-model conversion.
//!
//! Reads from the CRDT state (populated by the import pipeline) and produces
//! a [`DocViewModel`] — the same shape the existing TypeScript renderer consumes.
//!
//! This module is the editing-mode counterpart of [`crate::convert`]: the latter
//! converts directly from parsed `offidized_docx` types, while this one reads
//! from the CRDT (the source of truth during editing) and produces identical
//! output.

use std::collections::HashMap;
use std::sync::Arc;

use base64::Engine;
use yrs::types::text::YChange;
use yrs::{Any, Array, GetString, Map, MapRef, Out, ReadTxn, Text, Transact};

use super::crdt_doc::CrdtDoc;
use super::tokens::ATTR_TOKEN_TYPE;
use crate::convert::{format_number, map_bullet_char, resolve_numbering};
use crate::model::{
    BodyItem, DocViewModel, EndnoteModel, FootnoteModel, ImageModel, InlineImageModel,
    NumberingModel, ParagraphModel, RunModel, SectionModel, TableCellModel, TableModel,
    TableRowModel,
};
use crate::units::twips_to_pt;

type ImageIndexMap = HashMap<String, usize>;

/// Errors during CRDT-to-view-model conversion.
#[derive(Debug, thiserror::Error)]
pub enum ViewError {
    /// A required shared type or field was missing.
    #[error("missing CRDT field: {0}")]
    MissingField(String),

    /// A body element had an unexpected type.
    #[error("unexpected body element type: {0}")]
    UnexpectedType(String),

    /// Cast from yrs `Out` to the expected shared type failed.
    #[error("type cast failed: {0}")]
    CastFailed(String),
}

/// An incremental view update.
///
/// Sent from WASM to TypeScript after an edit to update the DOM
/// without a full re-render.
#[derive(Debug, Clone, serde::Serialize)]
#[serde(tag = "type", rename_all = "camelCase")]
pub enum ViewPatch {
    /// Replace a paragraph at the given index.
    ReplaceParagraph {
        /// The body index of the paragraph to replace.
        index: usize,
        /// The new paragraph model.
        model: ParagraphModel,
    },
    /// Insert a new paragraph at the given index.
    InsertParagraph {
        /// The body index to insert at.
        index: usize,
        /// The new paragraph model.
        model: ParagraphModel,
    },
    /// Remove the paragraph at the given index.
    RemoveParagraph {
        /// The body index of the paragraph to remove.
        index: usize,
    },
}

/// Convert CRDT document state into a [`DocViewModel`].
///
/// This reads from the CRDT (the source of truth during editing) and produces
/// the same `DocViewModel` shape that the existing `convert_document` function
/// produces from `offidized_docx` types. The TypeScript renderer is unchanged.
///
/// The `original_doc` parameter provides sections and image data from the
/// original parsed document (these are not stored in the CRDT).
pub fn crdt_to_view_model(
    crdt: &CrdtDoc,
    original_doc: &offidized_docx::Document,
) -> Result<DocViewModel, ViewError> {
    let txn = crdt.doc().transact();

    // Sections come from the original document (not in the CRDT).
    let sections = collect_sections(original_doc);

    // Images come from the original document.
    let (images, image_index_map) = convert_images(crdt, original_doc, &txn);

    // Body items from the CRDT.
    let body_arr = crdt.body();
    let body_len = body_arr.len(&txn);
    let mut body = Vec::with_capacity(body_len as usize);
    let mut numbering_counters = HashMap::<(u32, u8), u32>::new();

    for i in 0..body_len {
        let item = body_arr
            .get(&txn, i)
            .ok_or_else(|| ViewError::MissingField(format!("body[{i}]")))?;
        let map_ref = item
            .cast::<MapRef>()
            .map_err(|_| ViewError::CastFailed(format!("body[{i}] is not a YMap")))?;

        let type_str = read_string(&map_ref, &txn, "type")
            .ok_or_else(|| ViewError::MissingField(format!("body[{i}].type")))?;

        match type_str.as_str() {
            "paragraph" => {
                let para = read_paragraph(
                    &map_ref,
                    &txn,
                    original_doc,
                    &image_index_map,
                    &mut numbering_counters,
                )?;
                body.push(BodyItem::Paragraph(para));
            }
            "table" => {
                let table = read_table(&map_ref, &txn)?;
                body.push(BodyItem::Table(table));
            }
            other => {
                return Err(ViewError::UnexpectedType(other.to_string()));
            }
        }
    }

    // Footnotes from the CRDT (may have been edited).
    let footnotes = read_footnotes(crdt, &txn)?;

    // Endnotes from the CRDT (may have been edited).
    let endnotes = read_endnotes(crdt, &txn)?;

    Ok(DocViewModel {
        body,
        sections,
        images,
        footnotes,
        endnotes,
    })
}

// ---------------------------------------------------------------------------
// Helpers: read primitive values from a YMap
// ---------------------------------------------------------------------------

/// Extract a string value from a map entry.
fn read_string(map: &MapRef, txn: &impl ReadTxn, key: &str) -> Option<String> {
    match map.get(txn, key)? {
        Out::Any(Any::String(s)) => Some(s.to_string()),
        _ => None,
    }
}

/// Extract a number value from a map entry.
fn read_number(map: &MapRef, txn: &impl ReadTxn, key: &str) -> Option<f64> {
    match map.get(txn, key)? {
        Out::Any(Any::Number(n)) => Some(n),
        _ => None,
    }
}

/// Extract a bool value from a map entry.
fn read_bool(map: &MapRef, txn: &impl ReadTxn, key: &str) -> Option<bool> {
    match map.get(txn, key)? {
        Out::Any(Any::Bool(b)) => Some(b),
        _ => None,
    }
}

// ---------------------------------------------------------------------------
// Paragraph reading
// ---------------------------------------------------------------------------

/// Read a paragraph from a CRDT map and produce a [`ParagraphModel`].
fn read_paragraph(
    map: &MapRef,
    txn: &impl ReadTxn,
    original_doc: &offidized_docx::Document,
    image_index_map: &ImageIndexMap,
    numbering_counters: &mut HashMap<(u32, u8), u32>,
) -> Result<ParagraphModel, ViewError> {
    // Read text content via diff to get formatted chunks.
    let runs = match map.get(txn, "text") {
        Some(text_out) => {
            let text_ref = text_out
                .cast::<yrs::TextRef>()
                .map_err(|_| ViewError::CastFailed("paragraph text is not a TextRef".into()))?;
            let diffs = text_ref.diff(txn, YChange::identity);
            diffs_to_runs(&diffs, image_index_map)
        }
        None => Vec::new(),
    };

    // Read paragraph-level properties.
    let alignment = read_string(map, txn, "alignment");
    let heading_level = read_number(map, txn, "headingLevel").map(|n| n as u8);
    let style_id = read_string(map, txn, "styleId");

    let spacing_before_pt =
        read_number(map, txn, "spacingBeforeTwips").map(|t| twips_to_pt(t as u32));
    let spacing_after_pt =
        read_number(map, txn, "spacingAfterTwips").map(|t| twips_to_pt(t as u32));

    let line_spacing = read_number(map, txn, "lineSpacingTwips").map(|val| {
        let rule_str =
            read_string(map, txn, "lineSpacingRule").unwrap_or_else(|| "auto".to_string());
        let value = match rule_str.as_str() {
            "auto" => val / 240.0,
            "exact" | "atLeast" => twips_to_pt(val as u32),
            _ => val / 240.0,
        };
        crate::model::LineSpacingModel {
            value,
            rule: rule_str,
        }
    });

    let indents = {
        let left = read_number(map, txn, "indentLeftTwips").map(|t| twips_to_pt(t as u32));
        let right = read_number(map, txn, "indentRightTwips").map(|t| twips_to_pt(t as u32));
        let first_line =
            read_number(map, txn, "indentFirstLineTwips").map(|t| twips_to_pt(t as u32));
        let hanging = read_number(map, txn, "indentHangingTwips").map(|t| twips_to_pt(t as u32));
        if left.is_some() || right.is_some() || first_line.is_some() || hanging.is_some() {
            Some(crate::model::IndentsModel {
                left_pt: left,
                right_pt: right,
                first_line_pt: first_line,
                hanging_pt: hanging,
            })
        } else {
            None
        }
    };

    let numbering = {
        let num_id = read_number(map, txn, "numberingNumId");
        let ilvl = read_number(map, txn, "numberingIlvl");
        let numbering_kind = read_string(map, txn, "numberingKind");
        match (num_id, ilvl) {
            (Some(nid), Some(lvl)) => resolve_numbering(
                Some(nid as u32),
                Some(lvl as u8),
                original_doc,
                numbering_counters,
            )
            .or_else(|| {
                fallback_numbering_model(
                    nid as u32,
                    lvl as u8,
                    numbering_kind.as_deref(),
                    numbering_counters,
                )
            }),
            _ => None,
        }
    };

    let page_break_before = read_bool(map, txn, "pageBreakBefore").unwrap_or(false);
    let keep_next = read_bool(map, txn, "keepNext").unwrap_or(false);
    let keep_lines = read_bool(map, txn, "keepLines").unwrap_or(false);

    Ok(ParagraphModel {
        runs,
        heading_level,
        alignment,
        spacing_before_pt,
        spacing_after_pt,
        line_spacing,
        indents,
        numbering,
        borders: None,
        shading_color: None,
        page_break_before,
        keep_next,
        keep_lines,
        section_index: 0,
        ends_section: false,
        style_id,
    })
}

fn fallback_numbering_model(
    num_id: u32,
    level: u8,
    numbering_kind: Option<&str>,
    counters: &mut HashMap<(u32, u8), u32>,
) -> Option<NumberingModel> {
    let format = numbering_kind?.to_string();
    let text = if format == "bullet" {
        map_bullet_char("\u{2022}", "Symbol")
    } else {
        let counter = counters.entry((num_id, level)).or_insert(0);
        *counter += 1;
        format!("{}.", format_number(*counter, &format))
    };
    Some(NumberingModel {
        num_id,
        level,
        format,
        text,
    })
}

// ---------------------------------------------------------------------------
// Diff → RunModel conversion
// ---------------------------------------------------------------------------

/// Convert YText diff chunks into a list of [`RunModel`]s.
///
/// When multiple adjacent sentinels share the same attributes, yrs merges
/// them into a single diff chunk (e.g. two line breaks → one chunk with
/// text "￼￼").  We must split these back into individual runs so that
/// each sentinel produces its own `<br>` or tab element in the DOM.
fn diffs_to_runs(
    diffs: &[yrs::types::text::Diff<YChange>],
    image_index_map: &ImageIndexMap,
) -> Vec<RunModel> {
    use super::tokens;

    let mut runs = Vec::with_capacity(diffs.len());
    for diff in diffs {
        let attrs = diff
            .attributes
            .as_ref()
            .map(|boxed| boxed.as_ref().clone())
            .unwrap_or_default();

        // Check for a sentinel token.
        if let Some(Any::String(token_type)) = attrs.get(ATTR_TOKEN_TYPE as &str) {
            // Count sentinel characters in this chunk — yrs may have merged
            // multiple adjacent sentinels with identical attributes into one.
            let text = extract_text_from_out(&diff.insert);
            let sentinel_count = text.chars().filter(|c| tokens::is_sentinel(*c)).count();
            let count = sentinel_count.max(1);
            for _ in 0..count {
                runs.push(token_to_run_model(token_type, &attrs, image_index_map));
            }
        } else {
            // Regular text chunk.
            let text = extract_text_from_out(&diff.insert);
            runs.push(attrs_to_run_model(text, &attrs));
        }
    }
    runs
}

/// Build a [`RunModel`] for a sentinel token.
fn token_to_run_model(
    token_type: &Arc<str>,
    attrs: &HashMap<Arc<str>, Any>,
    image_index_map: &ImageIndexMap,
) -> RunModel {
    let mut run = attrs_to_run_formatting(attrs);

    match token_type.as_ref() {
        "tab" => {
            // Text left empty — the renderer uses has_tab to produce
            // a <span class="docview-tab"> element.  Setting text to "\t"
            // would create an extra text node that double-counts the
            // sentinel in offset calculations and, with white-space:
            // pre-wrap, renders a visible tab character.
            run.has_tab = true;
        }
        "lineBreak" => {
            // Text left empty — the renderer uses has_break to produce
            // a <br> element.  Setting text to "\n" would create an
            // extra text node that double-counts the sentinel in offset
            // calculations and, with white-space: pre-wrap, renders a
            // double line break.
            run.has_break = true;
        }
        "footnoteRef" => {
            if let Some(Any::Number(id)) = attrs.get("id" as &str) {
                run.footnote_ref = Some(*id as u32);
            }
            run.text = String::new();
        }
        "endnoteRef" => {
            if let Some(Any::Number(id)) = attrs.get("id" as &str) {
                run.endnote_ref = Some(*id as u32);
            }
            run.text = String::new();
        }
        "fieldSimple" => {
            if let Some(Any::String(pres)) = attrs.get("presentation" as &str) {
                run.text = pres.to_string();
            } else {
                run.text = String::new();
            }
        }
        "inlineImage" => {
            let image_ref = attrs
                .get("imageRef" as &str)
                .and_then(|v| match v {
                    Any::String(s) => Some(s.to_string()),
                    _ => None,
                })
                .unwrap_or_default();

            let width_emu = attrs
                .get("width" as &str)
                .and_then(|v| match v {
                    Any::Number(n) => Some(*n as i64),
                    _ => None,
                })
                .unwrap_or(0);

            let height_emu = attrs
                .get("height" as &str)
                .and_then(|v| match v {
                    Any::Number(n) => Some(*n as i64),
                    _ => None,
                })
                .unwrap_or(0);
            let name = read_attr_string(attrs, "name");
            let description = read_attr_string(attrs, "description");

            if let Some(&image_index) = image_index_map.get(&image_ref) {
                run.inline_image = Some(InlineImageModel {
                    image_index,
                    width_pt: crate::units::emu_to_pt(width_emu.max(0) as u32),
                    height_pt: crate::units::emu_to_pt(height_emu.max(0) as u32),
                    name,
                    description,
                });
            }
            run.text = String::new();
        }
        // Other token types (bookmarks, comments, field parts, opaque) are
        // preserved in the CRDT but invisible in the view model for now.
        _ => {
            run.text = String::new();
        }
    }

    run
}

fn read_attr_string(attrs: &HashMap<Arc<str>, Any>, key: &str) -> Option<String> {
    attrs.get(key).and_then(|value| match value {
        Any::String(value) => Some(value.to_string()),
        _ => None,
    })
}

/// Build a [`RunModel`] for a regular text chunk with formatting attrs.
fn attrs_to_run_model(text: String, attrs: &HashMap<Arc<str>, Any>) -> RunModel {
    let mut run = attrs_to_run_formatting(attrs);
    run.text = text;
    run
}

/// Build a [`RunModel`] with formatting properties from CRDT attrs.
///
/// Sets bold, italic, underline, font, etc. from the attribute map.
/// Text and token-specific fields are left at defaults.
fn attrs_to_run_formatting(attrs: &HashMap<Arc<str>, Any>) -> RunModel {
    let bold = matches!(attrs.get("bold" as &str), Some(Any::Bool(true)));
    let italic = matches!(attrs.get("italic" as &str), Some(Any::Bool(true)));
    let underline = matches!(attrs.get("underline" as &str), Some(Any::Bool(true)));
    let strikethrough = matches!(attrs.get("strike" as &str), Some(Any::Bool(true)));
    let superscript = matches!(attrs.get("superscript" as &str), Some(Any::Bool(true)));
    let subscript = matches!(attrs.get("subscript" as &str), Some(Any::Bool(true)));
    let small_caps = matches!(attrs.get("smallCaps" as &str), Some(Any::Bool(true)));

    let font_family = attrs.get("fontFamily" as &str).and_then(|v| match v {
        Any::String(s) => Some(s.to_string()),
        _ => None,
    });

    // Font size in the CRDT is stored as half-points. Convert to CSS points.
    let font_size_pt = attrs.get("fontSize" as &str).and_then(|v| match v {
        Any::Number(n) => Some(*n / 2.0),
        _ => None,
    });

    let color = attrs.get("color" as &str).and_then(|v| match v {
        Any::String(s) => Some(s.to_string()),
        _ => None,
    });

    let highlight = attrs.get("highlight" as &str).and_then(|v| match v {
        Any::String(s) => Some(s.to_string()),
        _ => None,
    });

    let hyperlink = attrs.get("hyperlink" as &str).and_then(|v| match v {
        Any::String(s) => Some(s.to_string()),
        _ => None,
    });

    RunModel {
        text: String::new(),
        bold,
        italic,
        underline,
        underline_type: None,
        strikethrough,
        superscript,
        subscript,
        small_caps,
        font_family,
        font_size_pt,
        color,
        highlight,
        hyperlink,
        hyperlink_tooltip: None,
        inline_image: None,
        floating_image: None,
        footnote_ref: None,
        endnote_ref: None,
        has_tab: false,
        has_break: false,
    }
}

/// Extract the text string from a yrs `Out` value.
fn extract_text_from_out(out: &Out) -> String {
    match out {
        Out::Any(Any::String(s)) => s.to_string(),
        _ => String::new(),
    }
}

// ---------------------------------------------------------------------------
// Table reading
// ---------------------------------------------------------------------------

/// Read a table from a CRDT map and produce a [`TableModel`].
fn read_table(map: &MapRef, txn: &impl ReadTxn) -> Result<TableModel, ViewError> {
    let num_rows = read_number(map, txn, "rows").unwrap_or(0.0) as usize;
    let num_cols = read_number(map, txn, "columns").unwrap_or(0.0) as usize;

    let mut rows = Vec::with_capacity(num_rows);
    for row in 0..num_rows {
        let mut cells = Vec::with_capacity(num_cols);
        for col in 0..num_cols {
            let key = format!("cell:{row}:{col}");
            let cell_text = match map.get(txn, key.as_str()) {
                Some(out) => match out.cast::<yrs::TextRef>() {
                    Ok(text_ref) => text_ref.get_string(txn),
                    Err(_) => String::new(),
                },
                None => String::new(),
            };

            cells.push(TableCellModel {
                text: cell_text,
                col_span: 1,
                row_span: 1,
                shading_color: None,
                vertical_align: None,
                width_pt: None,
                borders: None,
                is_covered: false,
            });
        }
        rows.push(TableRowModel {
            cells,
            height_pt: None,
            height_rule: None,
        });
    }

    Ok(TableModel {
        rows,
        width_pt: None,
        alignment: None,
        column_widths_pt: Vec::new(),
        borders: None,
        section_index: 0,
    })
}

// ---------------------------------------------------------------------------
// Footnotes / Endnotes from CRDT
// ---------------------------------------------------------------------------

/// Read footnotes from the CRDT array.
fn read_footnotes(crdt: &CrdtDoc, txn: &impl ReadTxn) -> Result<Vec<FootnoteModel>, ViewError> {
    let arr = crdt.footnotes();
    let len = arr.len(txn);
    let mut result = Vec::with_capacity(len as usize);
    for i in 0..len {
        let item = arr
            .get(txn, i)
            .ok_or_else(|| ViewError::MissingField(format!("footnotes[{i}]")))?;
        let map_ref = item
            .cast::<MapRef>()
            .map_err(|_| ViewError::CastFailed(format!("footnotes[{i}] is not a YMap")))?;
        let id = read_number(&map_ref, txn, "id").unwrap_or(0.0) as u32;
        let text = read_string(&map_ref, txn, "text").unwrap_or_default();
        result.push(FootnoteModel { id, text });
    }
    Ok(result)
}

/// Read endnotes from the CRDT array.
fn read_endnotes(crdt: &CrdtDoc, txn: &impl ReadTxn) -> Result<Vec<EndnoteModel>, ViewError> {
    let arr = crdt.endnotes();
    let len = arr.len(txn);
    let mut result = Vec::with_capacity(len as usize);
    for i in 0..len {
        let item = arr
            .get(txn, i)
            .ok_or_else(|| ViewError::MissingField(format!("endnotes[{i}]")))?;
        let map_ref = item
            .cast::<MapRef>()
            .map_err(|_| ViewError::CastFailed(format!("endnotes[{i}] is not a YMap")))?;
        let id = read_number(&map_ref, txn, "id").unwrap_or(0.0) as u32;
        let text = read_string(&map_ref, txn, "text").unwrap_or_default();
        result.push(EndnoteModel { id, text });
    }
    Ok(result)
}

// ---------------------------------------------------------------------------
// Sections (from original document, not CRDT)
// ---------------------------------------------------------------------------

/// Collect section definitions from the original document.
///
/// Sections are not stored in the CRDT. This uses the same logic as
/// [`crate::convert::collect_sections`] but is duplicated here to avoid
/// coupling to that module's internal API.
fn collect_sections(doc: &offidized_docx::Document) -> Vec<SectionModel> {
    use offidized_docx::BodyItem as DocBodyItem;

    let mut sections = Vec::new();

    // Sections from paragraph section breaks.
    for item in doc.body_items() {
        if let DocBodyItem::Paragraph(para) = item {
            if let Some(sec) = para.section_properties() {
                sections.push(convert_section(sec));
            }
        }
    }

    // Final section from document-level section properties.
    sections.push(convert_section(doc.section()));

    sections
}

/// Convert a single section to the view model.
fn convert_section(sec: &offidized_docx::section::Section) -> SectionModel {
    use offidized_docx::section::PageOrientation;

    let page_width_pt = sec.page_width_twips().map(twips_to_pt).unwrap_or(612.0);
    let page_height_pt = sec.page_height_twips().map(twips_to_pt).unwrap_or(792.0);
    let orientation = match sec.page_orientation() {
        Some(PageOrientation::Landscape) => "landscape",
        _ => "portrait",
    };

    let margins = sec.page_margins();
    let margins_model = crate::model::MarginsModel {
        top: margins.top_twips().map(twips_to_pt).unwrap_or(72.0),
        right: margins.right_twips().map(twips_to_pt).unwrap_or(72.0),
        bottom: margins.bottom_twips().map(twips_to_pt).unwrap_or(72.0),
        left: margins.left_twips().map(twips_to_pt).unwrap_or(72.0),
    };

    let header = sec.header().map(convert_header_footer);
    let footer = sec.footer().map(convert_header_footer);

    SectionModel {
        page_width_pt,
        page_height_pt,
        orientation: orientation.to_string(),
        margins: margins_model,
        header,
        footer,
        column_count: sec.column_count().unwrap_or(1),
    }
}

/// Convert a header/footer to the view model.
fn convert_header_footer(
    hf: &offidized_docx::section::HeaderFooter,
) -> crate::model::HeaderFooterModel {
    let paragraphs = hf
        .paragraphs()
        .iter()
        .map(|p| {
            // Simple paragraph conversion for headers/footers (no CRDT involved).
            let runs: Vec<RunModel> = p
                .runs()
                .iter()
                .map(|run| RunModel {
                    text: run.text().to_string(),
                    bold: run.is_bold(),
                    italic: run.is_italic(),
                    underline: run.is_underline(),
                    underline_type: None,
                    strikethrough: run.is_strikethrough(),
                    superscript: run.is_superscript(),
                    subscript: run.is_subscript(),
                    small_caps: run.is_small_caps(),
                    font_family: run.font_family().map(|s| s.to_string()),
                    font_size_pt: run.font_size_half_points().map(|hp| f64::from(hp) / 2.0),
                    color: run.color().map(|s| s.to_string()),
                    highlight: run.highlight_color().map(|s| s.to_string()),
                    hyperlink: run.hyperlink().map(|s| s.to_string()),
                    hyperlink_tooltip: run.hyperlink_tooltip().map(|s| s.to_string()),
                    inline_image: None,
                    floating_image: None,
                    footnote_ref: None,
                    endnote_ref: None,
                    has_tab: run.has_tab(),
                    has_break: run.has_break(),
                })
                .collect();

            ParagraphModel {
                runs,
                heading_level: p.heading_level(),
                alignment: p.alignment().map(|a| {
                    match a {
                        offidized_docx::paragraph::ParagraphAlignment::Left => "left",
                        offidized_docx::paragraph::ParagraphAlignment::Center => "center",
                        offidized_docx::paragraph::ParagraphAlignment::Right => "right",
                        offidized_docx::paragraph::ParagraphAlignment::Justified => "justify",
                    }
                    .to_string()
                }),
                spacing_before_pt: p.spacing_before_twips().map(twips_to_pt),
                spacing_after_pt: p.spacing_after_twips().map(twips_to_pt),
                line_spacing: None,
                indents: None,
                numbering: None,
                borders: None,
                shading_color: None,
                page_break_before: false,
                keep_next: false,
                keep_lines: false,
                section_index: 0,
                ends_section: false,
                style_id: p.style_id().map(|s| s.to_string()),
            }
        })
        .collect();
    crate::model::HeaderFooterModel { paragraphs }
}

// ---------------------------------------------------------------------------
// Images (from original document, not CRDT)
// ---------------------------------------------------------------------------

/// Convert images from the original document and local CRDT state to the view model.
fn convert_images(
    crdt: &CrdtDoc,
    doc: &offidized_docx::Document,
    txn: &impl ReadTxn,
) -> (Vec<ImageModel>, ImageIndexMap) {
    let engine = base64::engine::general_purpose::STANDARD;
    let mut images = Vec::new();
    let mut image_index_map = HashMap::new();

    for (index, img) in doc.images().iter().enumerate() {
        let encoded = engine.encode(img.bytes());
        let data_uri = format!("data:{};base64,{encoded}", img.content_type());
        image_index_map.insert(format!("img:{index}"), images.len());
        images.push(ImageModel {
            data_uri,
            content_type: img.content_type().to_string(),
        });
    }

    let images_map = crdt.images_map();
    for (key, value) in images_map.iter(txn) {
        let key_str: &str = key;
        let Ok(map_ref) = value.cast::<MapRef>() else {
            continue;
        };
        let Some(Out::Any(Any::String(content_type))) = map_ref.get(txn, "contentType") else {
            continue;
        };
        let data_uri = if let Some(bytes) = crdt.image_blobs().get(key_str) {
            let encoded = engine.encode(bytes);
            format!("data:{};base64,{encoded}", content_type.as_ref())
        } else if let Some(Out::Any(Any::String(data_uri))) = map_ref.get(txn, "dataUri") {
            data_uri.to_string()
        } else {
            continue;
        };
        image_index_map.insert(key_str.to_string(), images.len());
        images.push(ImageModel {
            data_uri,
            content_type: content_type.to_string(),
        });
    }

    (images, image_index_map)
}

#[cfg(test)]
#[allow(clippy::unwrap_used, clippy::expect_used, clippy::panic)]
mod tests {
    use super::*;
    use crate::editor::import::import_document;

    #[test]
    fn empty_document_roundtrip() {
        let doc = offidized_docx::Document::new();
        let mut crdt = CrdtDoc::new();
        import_document(&doc, &mut crdt).expect("import");

        let view = crdt_to_view_model(&crdt, &doc).expect("view model");
        assert!(view.body.is_empty());
        assert!(view.footnotes.is_empty());
        assert!(view.endnotes.is_empty());
        // Should have at least one section (the document-level section).
        assert!(!view.sections.is_empty());
    }

    #[test]
    fn paragraph_with_bold_text() {
        let mut doc = offidized_docx::Document::new();
        let para = doc.add_paragraph("Bold text");
        if let Some(run) = para.runs_mut().first_mut() {
            run.set_bold(true);
        }

        let mut crdt = CrdtDoc::new();
        import_document(&doc, &mut crdt).expect("import");

        let view = crdt_to_view_model(&crdt, &doc).expect("view model");
        assert_eq!(view.body.len(), 1);

        match &view.body[0] {
            BodyItem::Paragraph(p) => {
                assert_eq!(p.runs.len(), 1);
                assert_eq!(p.runs[0].text, "Bold text");
                assert!(p.runs[0].bold);
                assert!(!p.runs[0].italic);
            }
            BodyItem::Table(_) => panic!("expected paragraph, got table"),
        }
    }

    #[test]
    fn paragraph_with_tab_token() {
        let mut doc = offidized_docx::Document::new();
        let para = doc.add_paragraph("");
        if let Some(run) = para.runs_mut().first_mut() {
            run.set_has_tab(true);
        }

        let mut crdt = CrdtDoc::new();
        import_document(&doc, &mut crdt).expect("import");

        let view = crdt_to_view_model(&crdt, &doc).expect("view model");
        assert_eq!(view.body.len(), 1);

        match &view.body[0] {
            BodyItem::Paragraph(p) => {
                assert_eq!(p.runs.len(), 1);
                assert!(p.runs[0].has_tab);
                assert!(p.runs[0].text.is_empty());
            }
            BodyItem::Table(_) => panic!("expected paragraph, got table"),
        }
    }

    #[test]
    fn sections_from_original_doc() {
        let doc = offidized_docx::Document::new();
        let mut crdt = CrdtDoc::new();
        import_document(&doc, &mut crdt).expect("import");

        let view = crdt_to_view_model(&crdt, &doc).expect("view model");

        // The document-level section should always be present.
        assert!(!view.sections.is_empty());
        // Default US Letter.
        let sec = &view.sections[0];
        assert!((sec.page_width_pt - 612.0).abs() < 0.01);
        assert!((sec.page_height_pt - 792.0).abs() < 0.01);
        assert_eq!(sec.orientation, "portrait");
    }

    #[test]
    fn footnotes_from_crdt() {
        let mut doc = offidized_docx::Document::new();
        doc.footnotes_mut()
            .push(offidized_docx::Footnote::from_text(1, "Note one"));

        let mut crdt = CrdtDoc::new();
        import_document(&doc, &mut crdt).expect("import");

        let view = crdt_to_view_model(&crdt, &doc).expect("view model");
        assert_eq!(view.footnotes.len(), 1);
        assert_eq!(view.footnotes[0].id, 1);
        assert_eq!(view.footnotes[0].text, "Note one");
    }

    #[test]
    fn paragraph_heading_level() {
        let mut doc = offidized_docx::Document::new();
        doc.add_heading("Title", 1);

        let mut crdt = CrdtDoc::new();
        import_document(&doc, &mut crdt).expect("import");

        let view = crdt_to_view_model(&crdt, &doc).expect("view model");
        assert_eq!(view.body.len(), 1);

        match &view.body[0] {
            BodyItem::Paragraph(p) => {
                assert_eq!(p.heading_level, Some(1));
            }
            BodyItem::Table(_) => panic!("expected paragraph"),
        }
    }

    #[test]
    fn paragraph_formatting_attrs() {
        let mut doc = offidized_docx::Document::new();
        let para = doc.add_paragraph("Styled");
        if let Some(run) = para.runs_mut().first_mut() {
            run.set_italic(true);
            run.set_font_family("Arial");
            run.set_font_size_half_points(24); // 12pt
        }

        let mut crdt = CrdtDoc::new();
        import_document(&doc, &mut crdt).expect("import");

        let view = crdt_to_view_model(&crdt, &doc).expect("view model");
        match &view.body[0] {
            BodyItem::Paragraph(p) => {
                let r = &p.runs[0];
                assert!(r.italic);
                assert_eq!(r.font_family.as_deref(), Some("Arial"));
                assert!((r.font_size_pt.unwrap() - 12.0).abs() < f64::EPSILON);
            }
            BodyItem::Table(_) => panic!("expected paragraph"),
        }
    }

    #[test]
    fn paragraph_numbering_resolves_to_visible_marker() {
        let mut doc = offidized_docx::Document::new();
        doc.add_bulleted_paragraph("Bullet item");
        doc.add_numbered_paragraph("Numbered item");

        let mut crdt = CrdtDoc::new();
        import_document(&doc, &mut crdt).expect("import");

        let view = crdt_to_view_model(&crdt, &doc).expect("view model");
        assert_eq!(view.body.len(), 2);

        match &view.body[0] {
            BodyItem::Paragraph(p) => {
                let numbering = p.numbering.as_ref().expect("bullet numbering");
                assert_eq!(numbering.format, "bullet");
                assert_eq!(numbering.text, "\u{2022}");
            }
            BodyItem::Table(_) => panic!("expected paragraph"),
        }

        match &view.body[1] {
            BodyItem::Paragraph(p) => {
                let numbering = p.numbering.as_ref().expect("decimal numbering");
                assert_eq!(numbering.format, "decimal");
                assert_eq!(numbering.text, "1.");
            }
            BodyItem::Table(_) => panic!("expected paragraph"),
        }
    }

    #[test]
    fn multiple_paragraphs() {
        let mut doc = offidized_docx::Document::new();
        doc.add_paragraph("First");
        doc.add_paragraph("Second");

        let mut crdt = CrdtDoc::new();
        import_document(&doc, &mut crdt).expect("import");

        let view = crdt_to_view_model(&crdt, &doc).expect("view model");
        assert_eq!(view.body.len(), 2);
    }

    #[test]
    fn images_from_original_doc() {
        let mut doc = offidized_docx::Document::new();
        doc.add_image(vec![0xFF, 0xD8], "image/jpeg");

        let mut crdt = CrdtDoc::new();
        import_document(&doc, &mut crdt).expect("import");

        let view = crdt_to_view_model(&crdt, &doc).expect("view model");
        assert_eq!(view.images.len(), 1);
        assert_eq!(view.images[0].content_type, "image/jpeg");
        assert!(view.images[0]
            .data_uri
            .starts_with("data:image/jpeg;base64,"));
    }
}
