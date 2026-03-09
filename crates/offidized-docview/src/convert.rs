//! Converts an [`offidized_docx::Document`] into a [`DocViewModel`](crate::model::DocViewModel).

use std::collections::HashMap;

use base64::Engine;
use offidized_docx::paragraph::{LineSpacingRule, ParagraphAlignment, ParagraphBorder};
use offidized_docx::section::{HeaderFooter, PageOrientation, Section};
use offidized_docx::table::{TableAlignment, TableBorder, VerticalAlignment, VerticalMerge};
use offidized_docx::{BodyItem, Document};

use crate::model::{
    BodyItem as ViewBodyItem, BorderModel, BordersModel, DocViewModel, EndnoteModel,
    FloatingImageModel, FootnoteModel, HeaderFooterModel, ImageModel, IndentsModel,
    InlineImageModel, LineSpacingModel, MarginsModel, NumberingModel, ParagraphModel, RunModel,
    SectionModel, TableCellModel, TableModel, TableRowModel,
};
use crate::units::{emu_to_pt, half_points_to_pt, signed_emu_to_pt, twips_to_pt};

/// Errors that can occur during document conversion.
#[derive(Debug, thiserror::Error)]
pub enum ConvertError {
    /// Base64 encoding failed.
    #[error("base64 encoding failed: {0}")]
    Base64(String),
}

/// Result type for conversion operations.
pub type Result<T> = std::result::Result<T, ConvertError>;

/// Convert a parsed `Document` into a renderer-friendly `DocViewModel`.
pub fn convert_document(doc: &Document) -> Result<DocViewModel> {
    let images = convert_images(doc);
    let sections = collect_sections(doc);
    let footnotes = convert_footnotes(doc);
    let endnotes = convert_endnotes(doc);

    // Build numbering counters for resolving list prefixes.
    let mut numbering_counters: HashMap<(u32, u8), u32> = HashMap::new();

    // Walk body items, tracking current section index.
    let mut body = Vec::new();
    let mut section_index: usize = 0;

    for item in doc.body_items() {
        match item {
            BodyItem::Paragraph(para) => {
                let ends_section = para.section_properties().is_some();
                let numbering = resolve_numbering(
                    para.numbering_num_id(),
                    para.numbering_ilvl(),
                    doc,
                    &mut numbering_counters,
                );
                let paragraph_model =
                    convert_paragraph(para, section_index, ends_section, numbering);
                body.push(ViewBodyItem::Paragraph(paragraph_model));
                if ends_section && section_index + 1 < sections.len() {
                    section_index += 1;
                }
            }
            BodyItem::Table(table) => {
                let table_model = convert_table(table, section_index);
                body.push(ViewBodyItem::Table(table_model));
            }
        }
    }

    Ok(DocViewModel {
        body,
        sections,
        images,
        footnotes,
        endnotes,
    })
}

// ---------------------------------------------------------------------------
// Images
// ---------------------------------------------------------------------------

fn convert_images(doc: &Document) -> Vec<ImageModel> {
    let engine = base64::engine::general_purpose::STANDARD;
    doc.images()
        .iter()
        .map(|img| {
            let encoded = engine.encode(img.bytes());
            let data_uri = format!("data:{};base64,{encoded}", img.content_type());
            ImageModel {
                data_uri,
                content_type: img.content_type().to_string(),
            }
        })
        .collect()
}

// ---------------------------------------------------------------------------
// Sections
// ---------------------------------------------------------------------------

fn collect_sections(doc: &Document) -> Vec<SectionModel> {
    let mut sections = Vec::new();

    // Sections from paragraph section breaks.
    for item in doc.body_items() {
        if let BodyItem::Paragraph(para) = item {
            if let Some(sec) = para.section_properties() {
                sections.push(convert_section(sec));
            }
        }
    }

    // Final section from document section.
    sections.push(convert_section(doc.section()));

    sections
}

fn convert_section(sec: &Section) -> SectionModel {
    // Default US Letter dimensions: 8.5 × 11 inches = 612 × 792 pt.
    let page_width_pt = sec.page_width_twips().map(twips_to_pt).unwrap_or(612.0);
    let page_height_pt = sec.page_height_twips().map(twips_to_pt).unwrap_or(792.0);
    let orientation = match sec.page_orientation() {
        Some(PageOrientation::Landscape) => "landscape",
        _ => "portrait",
    };

    let margins = sec.page_margins();
    let margins_model = MarginsModel {
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

fn convert_header_footer(hf: &HeaderFooter) -> HeaderFooterModel {
    let paragraphs = hf
        .paragraphs()
        .iter()
        .map(|p| convert_paragraph(p, 0, false, None))
        .collect();
    HeaderFooterModel { paragraphs }
}

// ---------------------------------------------------------------------------
// Paragraphs
// ---------------------------------------------------------------------------

fn convert_paragraph(
    para: &offidized_docx::paragraph::Paragraph,
    section_index: usize,
    ends_section: bool,
    numbering: Option<NumberingModel>,
) -> ParagraphModel {
    let runs: Vec<RunModel> = para.runs().iter().map(convert_run).collect();

    let alignment = para.alignment().map(|a| {
        match a {
            ParagraphAlignment::Left => "left",
            ParagraphAlignment::Center => "center",
            ParagraphAlignment::Right => "right",
            ParagraphAlignment::Justified => "justify",
        }
        .to_string()
    });

    let spacing_before_pt = para.spacing_before_twips().map(twips_to_pt);
    let spacing_after_pt = para.spacing_after_twips().map(twips_to_pt);

    let line_spacing = para.line_spacing_twips().map(|val| {
        let rule = para.line_spacing_rule().unwrap_or(LineSpacingRule::Auto);
        let (value, rule_str) = match rule {
            LineSpacingRule::Auto => {
                // Auto: value is in 240ths of a line; convert to a multiplier.
                (val as f64 / 240.0, "auto")
            }
            LineSpacingRule::Exact => (twips_to_pt(val), "exact"),
            LineSpacingRule::AtLeast => (twips_to_pt(val), "atLeast"),
        };
        LineSpacingModel {
            value,
            rule: rule_str.to_string(),
        }
    });

    let indents = {
        let left = para.indent_left_twips().map(twips_to_pt);
        let right = para.indent_right_twips().map(twips_to_pt);
        let first_line = para.indent_first_line_twips().map(twips_to_pt);
        let hanging = para.indent_hanging_twips().map(twips_to_pt);
        if left.is_some() || right.is_some() || first_line.is_some() || hanging.is_some() {
            Some(IndentsModel {
                left_pt: left,
                right_pt: right,
                first_line_pt: first_line,
                hanging_pt: hanging,
            })
        } else {
            None
        }
    };

    let borders = convert_paragraph_borders(para.borders());
    let shading_color = para.shading_color().map(|s| s.to_string());

    ParagraphModel {
        runs,
        heading_level: para.heading_level(),
        alignment,
        spacing_before_pt,
        spacing_after_pt,
        line_spacing,
        indents,
        numbering,
        borders,
        shading_color,
        page_break_before: para.page_break_before(),
        keep_next: para.keep_next(),
        keep_lines: para.keep_lines(),
        section_index,
        ends_section,
        style_id: para.style_id().map(|s| s.to_string()),
    }
}

// ---------------------------------------------------------------------------
// Runs
// ---------------------------------------------------------------------------

fn convert_run(run: &offidized_docx::run::Run) -> RunModel {
    let underline_type = run
        .underline_type()
        .map(|ut| format!("{ut:?}").to_lowercase());

    let inline_image = run.inline_image().map(|img| InlineImageModel {
        image_index: img.image_index(),
        width_pt: emu_to_pt(img.width_emu()),
        height_pt: emu_to_pt(img.height_emu()),
        name: img.name().map(|s| s.to_string()),
        description: img.description().map(|s| s.to_string()),
    });

    let floating_image = run.floating_image().map(|img| {
        let wrap_type = img.wrap_type().map(|w| format!("{w:?}"));
        FloatingImageModel {
            image_index: img.image_index(),
            width_pt: emu_to_pt(img.width_emu()),
            height_pt: emu_to_pt(img.height_emu()),
            offset_x_pt: signed_emu_to_pt(img.offset_x_emu()),
            offset_y_pt: signed_emu_to_pt(img.offset_y_emu()),
            name: img.name().map(|s| s.to_string()),
            description: img.description().map(|s| s.to_string()),
            wrap_type,
        }
    });

    RunModel {
        text: run.text().to_string(),
        bold: run.is_bold(),
        italic: run.is_italic(),
        underline: run.is_underline(),
        underline_type,
        strikethrough: run.is_strikethrough(),
        superscript: run.is_superscript(),
        subscript: run.is_subscript(),
        small_caps: run.is_small_caps(),
        font_family: run.font_family().map(|s| s.to_string()),
        font_size_pt: run.font_size_half_points().map(half_points_to_pt),
        color: run.color().map(|s| s.to_string()),
        highlight: run.highlight_color().map(|s| s.to_string()),
        hyperlink: run.hyperlink().map(|s| s.to_string()),
        hyperlink_tooltip: run.hyperlink_tooltip().map(|s| s.to_string()),
        inline_image,
        floating_image,
        footnote_ref: run.footnote_reference_id(),
        endnote_ref: run.endnote_reference_id(),
        has_tab: run.has_tab(),
        has_break: run.has_break(),
    }
}

// ---------------------------------------------------------------------------
// Tables
// ---------------------------------------------------------------------------

fn convert_table(table: &offidized_docx::table::Table, section_index: usize) -> TableModel {
    let num_rows = table.rows();
    let num_cols = table.columns();
    let column_widths_pt: Vec<f64> = table
        .column_widths_twips()
        .iter()
        .map(|&tw| twips_to_pt(tw))
        .collect();

    // Compute vertical merge spans.
    // For each (row, col) that starts a vertical merge (Restart),
    // count how many Continue cells follow.
    let mut vmerge_spans: HashMap<(usize, usize), usize> = HashMap::new();
    let mut covered: HashMap<(usize, usize), bool> = HashMap::new();

    for col in 0..num_cols {
        let mut merge_start: Option<usize> = None;
        for row in 0..num_rows {
            let vm = table.cell(row, col).and_then(|c| c.vertical_merge());
            match vm {
                Some(VerticalMerge::Restart) => {
                    merge_start = Some(row);
                }
                Some(VerticalMerge::Continue) => {
                    covered.insert((row, col), true);
                    if let Some(start) = merge_start {
                        let count = vmerge_spans.entry((start, col)).or_insert(1);
                        *count += 1;
                    }
                }
                None => {
                    merge_start = None;
                }
            }
        }
    }

    let mut rows = Vec::with_capacity(num_rows);
    for row in 0..num_rows {
        let row_height_pt = table
            .row_properties(row)
            .and_then(|rp| rp.height_twips())
            .map(twips_to_pt);

        let mut cells = Vec::with_capacity(num_cols);
        for col in 0..num_cols {
            if let Some(cell) = table.cell(row, col) {
                let is_covered = covered.get(&(row, col)).copied().unwrap_or(false)
                    || cell.is_horizontal_merge_continuation();
                let row_span = vmerge_spans.get(&(row, col)).copied().unwrap_or(1);
                let col_span = cell.horizontal_span();
                let cell_borders = convert_cell_borders(cell.borders());

                cells.push(TableCellModel {
                    text: cell.text().to_string(),
                    col_span,
                    row_span,
                    shading_color: cell.shading_color().map(|s| s.to_string()),
                    vertical_align: cell.vertical_alignment().map(|va| {
                        match va {
                            VerticalAlignment::Top => "top",
                            VerticalAlignment::Center => "center",
                            VerticalAlignment::Bottom => "bottom",
                        }
                        .to_string()
                    }),
                    width_pt: cell.cell_width_twips().map(twips_to_pt),
                    borders: cell_borders,
                    is_covered,
                });
            }
        }

        rows.push(TableRowModel {
            cells,
            height_pt: row_height_pt,
        });
    }

    let width_pt = table.width_twips().map(twips_to_pt);
    let alignment = table.alignment().map(|a| {
        match a {
            TableAlignment::Left => "left",
            TableAlignment::Center => "center",
            TableAlignment::Right => "right",
        }
        .to_string()
    });

    let table_borders = convert_table_borders(table.borders());

    TableModel {
        rows,
        width_pt,
        alignment,
        column_widths_pt,
        borders: table_borders,
        section_index,
    }
}

// ---------------------------------------------------------------------------
// Numbering
// ---------------------------------------------------------------------------

pub(crate) fn resolve_numbering(
    num_id: Option<u32>,
    ilvl: Option<u8>,
    doc: &Document,
    counters: &mut HashMap<(u32, u8), u32>,
) -> Option<NumberingModel> {
    let num_id = num_id?;
    let level = ilvl.unwrap_or(0);

    // Find the numbering instance → abstract definition → level.
    let instance = doc
        .numbering_instances()
        .iter()
        .find(|ni| ni.num_id() == num_id)?;
    let abstract_id = instance.abstract_num_id();

    // Check for start override.
    let start_override = instance
        .level_overrides()
        .iter()
        .find(|lo| lo.level() == level)
        .and_then(|lo| lo.start_override());

    let definition = doc
        .numbering_definitions()
        .iter()
        .find(|nd| nd.abstract_num_id() == abstract_id)?;
    let num_level = definition.level(level)?;

    let format = num_level.format().to_string();
    let level_text = num_level.text().to_string();

    // Resolve the display text.
    let text = if format == "bullet" {
        // Bullet: map PUA characters from Symbol/Wingdings fonts to standard Unicode.
        // Word stores bullet chars as PUA codepoints (U+F000–U+F0FF) that only render
        // in their specific font. We map common ones to standard Unicode equivalents.
        let font = num_level.font_family().unwrap_or("");
        if level_text.is_empty() {
            "\u{2022}".to_string()
        } else {
            map_bullet_char(&level_text, font)
        }
    } else {
        // Numbered: increment counter and format.
        let start = start_override.unwrap_or(num_level.start());
        let counter = counters
            .entry((num_id, level))
            .or_insert(start.saturating_sub(1));
        *counter += 1;
        let current = *counter;

        // Replace %N placeholders in the level text with the actual number.
        let formatted_number = format_number(current, &format);
        // Level text like "%1." — replace the %N pattern.
        let mut result = level_text;
        let placeholder = format!("%{}", level + 1);
        result = result.replace(&placeholder, &formatted_number);
        result
    };

    Some(NumberingModel {
        num_id,
        level,
        format,
        text,
    })
}

/// Map Word bullet characters (often PUA codepoints from Symbol/Wingdings) to
/// standard Unicode characters that render in any web font.
pub(crate) fn map_bullet_char(text: &str, font: &str) -> String {
    // If the text is a single character in the PUA range, map it.
    let mut chars = text.chars();
    if let Some(ch) = chars.next() {
        if chars.next().is_none() && ('\u{F000}'..='\u{F0FF}').contains(&ch) {
            let code = ch as u32 & 0xFF;
            let font_lower = font.to_lowercase();
            let mapped = if font_lower.contains("wingdings") {
                match code {
                    0x6C => '\u{25CF}', // filled circle
                    0x6E => '\u{25A0}', // filled square
                    0x71 => '\u{25C6}', // filled diamond (checkbox empty in Wingdings but commonly diamond)
                    0x76 => '\u{2714}', // check mark
                    0x77 => '\u{2718}', // cross mark
                    0xA7 => '\u{25A0}', // filled square (common Wingdings bullet)
                    0xA8 => '\u{25CB}', // white circle
                    0xD8 => '\u{27A2}', // right arrow
                    0xFC => '\u{2714}', // check mark
                    0xFB => '\u{25CF}', // filled circle
                    _ => '\u{2022}',    // fallback: standard bullet
                }
            } else if font_lower.contains("symbol") {
                match code {
                    0xB7 => '\u{2022}', // bullet (·)
                    0x6F => '\u{25CB}', // white circle (o)
                    0xA7 => '\u{2666}', // diamond
                    _ => '\u{2022}',    // fallback: standard bullet
                }
            } else {
                // Unknown font with PUA char — use standard bullet.
                '\u{2022}'
            };
            return mapped.to_string();
        }
    }
    // Not a PUA char — return as-is.
    text.to_string()
}

pub(crate) fn format_number(n: u32, format: &str) -> String {
    match format {
        "decimal" => n.to_string(),
        "lowerLetter" => {
            if n == 0 {
                return String::new();
            }
            let idx = ((n - 1) % 26) as u8;
            let c = b'a' + idx;
            String::from(c as char)
        }
        "upperLetter" => {
            if n == 0 {
                return String::new();
            }
            let idx = ((n - 1) % 26) as u8;
            let c = b'A' + idx;
            String::from(c as char)
        }
        "lowerRoman" => to_roman(n).to_lowercase(),
        "upperRoman" => to_roman(n),
        _ => n.to_string(),
    }
}

fn to_roman(mut n: u32) -> String {
    let table = [
        (1000, "M"),
        (900, "CM"),
        (500, "D"),
        (400, "CD"),
        (100, "C"),
        (90, "XC"),
        (50, "L"),
        (40, "XL"),
        (10, "X"),
        (9, "IX"),
        (5, "V"),
        (4, "IV"),
        (1, "I"),
    ];
    let mut result = String::new();
    for &(value, numeral) in &table {
        while n >= value {
            result.push_str(numeral);
            n -= value;
        }
    }
    result
}

// ---------------------------------------------------------------------------
// Borders
// ---------------------------------------------------------------------------

fn convert_paragraph_border(border: &ParagraphBorder) -> Option<BorderModel> {
    let style = border.line_type()?;
    if style == "none" || style == "nil" {
        return None;
    }
    Some(BorderModel {
        style: style.to_string(),
        color: border.color().map(|s| s.to_string()),
        width_pt: border.size_eighth_points().map(|s| s as f64 / 8.0),
    })
}

fn convert_paragraph_borders(
    borders: &offidized_docx::paragraph::ParagraphBorders,
) -> Option<BordersModel> {
    let top = borders.top().and_then(convert_paragraph_border);
    let right = borders.right().and_then(convert_paragraph_border);
    let bottom = borders.bottom().and_then(convert_paragraph_border);
    let left = borders.left().and_then(convert_paragraph_border);

    if top.is_none() && right.is_none() && bottom.is_none() && left.is_none() {
        return None;
    }

    Some(BordersModel {
        top,
        right,
        bottom,
        left,
    })
}

fn convert_table_border(border: &TableBorder) -> Option<BorderModel> {
    let style = border.line_type()?;
    if style == "none" || style == "nil" {
        return None;
    }
    Some(BorderModel {
        style: style.to_string(),
        color: border.color().map(|s| s.to_string()),
        width_pt: border.size_eighth_points().map(|s| s as f64 / 8.0),
    })
}

fn convert_table_borders(borders: &offidized_docx::table::TableBorders) -> Option<BordersModel> {
    let top = borders.top().and_then(convert_table_border);
    let right = borders.right().and_then(convert_table_border);
    let bottom = borders.bottom().and_then(convert_table_border);
    let left = borders.left().and_then(convert_table_border);

    if top.is_none() && right.is_none() && bottom.is_none() && left.is_none() {
        return None;
    }

    Some(BordersModel {
        top,
        right,
        bottom,
        left,
    })
}

fn convert_cell_borders(borders: &offidized_docx::table::CellBorders) -> Option<BordersModel> {
    let top = borders.top().and_then(convert_table_border);
    let right = borders.right().and_then(convert_table_border);
    let bottom = borders.bottom().and_then(convert_table_border);
    let left = borders.left().and_then(convert_table_border);

    if top.is_none() && right.is_none() && bottom.is_none() && left.is_none() {
        return None;
    }

    Some(BordersModel {
        top,
        right,
        bottom,
        left,
    })
}

// ---------------------------------------------------------------------------
// Footnotes / Endnotes
// ---------------------------------------------------------------------------

fn convert_footnotes(doc: &Document) -> Vec<FootnoteModel> {
    doc.footnotes()
        .iter()
        .map(|fn_| FootnoteModel {
            id: fn_.id(),
            text: fn_.text(),
        })
        .collect()
}

fn convert_endnotes(doc: &Document) -> Vec<EndnoteModel> {
    doc.endnotes()
        .iter()
        .map(|en| EndnoteModel {
            id: en.id(),
            text: en.text(),
        })
        .collect()
}
