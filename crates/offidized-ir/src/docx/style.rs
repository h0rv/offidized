//! docx style mode: derive and apply section + paragraph formatting properties.
//!
//! Style mode emits layout/formatting properties without touching text content.
//! Properties use additive semantics: properties present in the IR are applied;
//! properties absent are left unchanged.

use crate::{ApplyResult, IrError, Result};
use offidized_docx::{Document, PageOrientation, ParagraphAlignment};

// ---------------------------------------------------------------------------
// Derive
// ---------------------------------------------------------------------------

/// Append the docx style-mode body to `output`.
pub(crate) fn derive_style(doc: &Document, output: &mut String) {
    // Section properties
    let section = doc.section();
    let mut section_lines = Vec::new();

    if let Some(w) = section.page_width_twips() {
        section_lines.push(format!("page-width: {w}"));
    }
    if let Some(h) = section.page_height_twips() {
        section_lines.push(format!("page-height: {h}"));
    }
    if let Some(orient) = section.page_orientation() {
        section_lines.push(format!(
            "orientation: {}",
            match orient {
                PageOrientation::Portrait => "portrait",
                PageOrientation::Landscape => "landscape",
            }
        ));
    }

    let margins = section.page_margins();
    if let Some(v) = margins.top_twips() {
        section_lines.push(format!("margin-top: {v}"));
    }
    if let Some(v) = margins.bottom_twips() {
        section_lines.push(format!("margin-bottom: {v}"));
    }
    if let Some(v) = margins.left_twips() {
        section_lines.push(format!("margin-left: {v}"));
    }
    if let Some(v) = margins.right_twips() {
        section_lines.push(format!("margin-right: {v}"));
    }

    if !section_lines.is_empty() {
        output.push_str("\n# Section\n");
        for line in &section_lines {
            output.push_str(line);
            output.push('\n');
        }
    }

    // Paragraphs
    let paragraphs = doc.paragraphs();
    if paragraphs.is_empty() {
        return;
    }

    let mut para_lines = Vec::new();
    for (i, para) in paragraphs.iter().enumerate() {
        let idx = i + 1;
        let mut props = Vec::new();

        if let Some(sid) = para.style_id() {
            props.push(format!("style=\"{sid}\""));
        }
        if let Some(align) = para.alignment() {
            props.push(format!("align={}", alignment_to_ir(align)));
        }
        if let Some(v) = para.spacing_before_twips() {
            props.push(format!("spacing-before={v}"));
        }
        if let Some(v) = para.spacing_after_twips() {
            props.push(format!("spacing-after={v}"));
        }
        if let Some(v) = para.line_spacing_twips() {
            props.push(format!("line-spacing={v}"));
        }
        if let Some(v) = para.indent_left_twips() {
            props.push(format!("indent-left={v}"));
        }
        if let Some(v) = para.indent_right_twips() {
            props.push(format!("indent-right={v}"));
        }
        if let Some(v) = para.indent_first_line_twips() {
            props.push(format!("indent-first={v}"));
        }
        if let Some(v) = para.indent_hanging_twips() {
            props.push(format!("indent-hanging={v}"));
        }
        if para.keep_next() {
            props.push("keep-next".to_string());
        }
        if para.keep_lines() {
            props.push("keep-lines".to_string());
        }
        if para.page_break_before() {
            props.push("page-break-before".to_string());
        }

        if !props.is_empty() {
            para_lines.push(format!("[p{idx}] {}", props.join(", ")));
        }
    }

    if !para_lines.is_empty() {
        output.push_str("\n# Paragraphs\n");
        for line in &para_lines {
            output.push_str(line);
            output.push('\n');
        }
    }
}

// ---------------------------------------------------------------------------
// Apply
// ---------------------------------------------------------------------------

/// Apply the style IR body to a document, updating formatting.
pub(crate) fn apply_style(body: &str, doc: &mut Document) -> Result<ApplyResult> {
    let mut result = ApplyResult::default();

    for line in body.lines() {
        let line = line.trim();

        // Skip comments, blank lines, section headers
        if line.is_empty() || line.starts_with('#') {
            continue;
        }

        // Section property lines
        if let Some(rest) = line.strip_prefix("page-width: ") {
            if let Ok(v) = rest.trim().parse::<u32>() {
                let h = doc.section().page_height_twips().unwrap_or(15840);
                doc.section_mut().set_page_size_twips(v, h);
            }
            continue;
        }
        if let Some(rest) = line.strip_prefix("page-height: ") {
            if let Ok(v) = rest.trim().parse::<u32>() {
                let w = doc.section().page_width_twips().unwrap_or(12240);
                doc.section_mut().set_page_size_twips(w, v);
            }
            continue;
        }
        if let Some(rest) = line.strip_prefix("orientation: ") {
            let orient = match rest.trim() {
                "portrait" => PageOrientation::Portrait,
                "landscape" => PageOrientation::Landscape,
                _ => continue,
            };
            doc.section_mut().set_page_orientation(orient);
            continue;
        }
        if let Some(rest) = line.strip_prefix("margin-top: ") {
            if let Ok(v) = rest.trim().parse::<u32>() {
                doc.section_mut().page_margins_mut().set_top_twips(v);
            }
            continue;
        }
        if let Some(rest) = line.strip_prefix("margin-bottom: ") {
            if let Ok(v) = rest.trim().parse::<u32>() {
                doc.section_mut().page_margins_mut().set_bottom_twips(v);
            }
            continue;
        }
        if let Some(rest) = line.strip_prefix("margin-left: ") {
            if let Ok(v) = rest.trim().parse::<u32>() {
                doc.section_mut().page_margins_mut().set_left_twips(v);
            }
            continue;
        }
        if let Some(rest) = line.strip_prefix("margin-right: ") {
            if let Ok(v) = rest.trim().parse::<u32>() {
                doc.section_mut().page_margins_mut().set_right_twips(v);
            }
            continue;
        }

        // Paragraph lines: [pN] props...
        if let Some((para_idx, props_str)) = parse_paragraph_line(line) {
            let para_i = para_idx.checked_sub(1).ok_or_else(|| {
                IrError::InvalidBody(format!("paragraph index must be >= 1: {para_idx}"))
            })?;

            let paragraphs = doc.paragraphs_mut();
            if para_i >= paragraphs.len() {
                result
                    .warnings
                    .push(format!("paragraph index out of range: p{para_idx}"));
                continue;
            }

            let para = &mut paragraphs[para_i];
            apply_paragraph_properties(para, props_str)?;
            result.cells_updated += 1;
        }
    }

    Ok(result)
}

/// Parse a paragraph line like `[p1] style="Normal", align=center`.
fn parse_paragraph_line(line: &str) -> Option<(usize, &str)> {
    let rest = line.strip_prefix("[p")?;
    let bracket_end = rest.find(']')?;
    let idx: usize = rest[..bracket_end].parse().ok()?;
    let props = rest[bracket_end + 1..].trim();
    Some((idx, props))
}

/// Apply comma-separated property strings to a paragraph.
fn apply_paragraph_properties(para: &mut offidized_docx::Paragraph, props_str: &str) -> Result<()> {
    for prop in split_properties(props_str) {
        let prop = prop.trim();
        if prop.is_empty() {
            continue;
        }

        // Boolean flags
        match prop {
            "keep-next" => {
                para.set_keep_next(true);
                continue;
            }
            "keep-lines" => {
                para.set_keep_lines(true);
                continue;
            }
            "page-break-before" => {
                para.set_page_break_before(true);
                continue;
            }
            _ => {}
        }

        // Key=value properties
        if let Some((key, value)) = prop.split_once('=') {
            let key = key.trim();
            let value = value.trim();

            match key {
                "style" => {
                    let sid = value.trim_matches('"');
                    para.set_style_id(sid);
                }
                "align" => {
                    if let Some(a) = ir_to_alignment(value) {
                        para.set_alignment(a);
                    }
                }
                "spacing-before" => {
                    if let Ok(v) = value.parse::<u32>() {
                        para.set_spacing_before_twips(v);
                    }
                }
                "spacing-after" => {
                    if let Ok(v) = value.parse::<u32>() {
                        para.set_spacing_after_twips(v);
                    }
                }
                "line-spacing" => {
                    if let Ok(v) = value.parse::<u32>() {
                        para.set_line_spacing_twips(v);
                    }
                }
                "indent-left" => {
                    if let Ok(v) = value.parse::<u32>() {
                        para.set_indent_left_twips(v);
                    }
                }
                "indent-right" => {
                    if let Ok(v) = value.parse::<u32>() {
                        para.set_indent_right_twips(v);
                    }
                }
                "indent-first" => {
                    if let Ok(v) = value.parse::<u32>() {
                        para.set_indent_first_line_twips(v);
                    }
                }
                "indent-hanging" => {
                    if let Ok(v) = value.parse::<u32>() {
                        para.set_indent_hanging_twips(v);
                    }
                }
                _ => {
                    // Unknown property — ignore silently
                }
            }
        }
    }

    Ok(())
}

/// Quote-aware comma splitting (same logic as xlsx).
fn split_properties(s: &str) -> Vec<&str> {
    let mut result = Vec::new();
    let mut start = 0;
    let mut in_quotes = false;
    let bytes = s.as_bytes();

    for i in 0..bytes.len() {
        match bytes[i] {
            b'"' => in_quotes = !in_quotes,
            b',' if !in_quotes => {
                let part = s[start..i].trim();
                if !part.is_empty() {
                    result.push(part);
                }
                start = i + 1;
            }
            _ => {}
        }
    }

    let last = s[start..].trim();
    if !last.is_empty() {
        result.push(last);
    }

    result
}

fn alignment_to_ir(a: ParagraphAlignment) -> &'static str {
    match a {
        ParagraphAlignment::Left => "left",
        ParagraphAlignment::Center => "center",
        ParagraphAlignment::Right => "right",
        ParagraphAlignment::Justified => "justified",
    }
}

fn ir_to_alignment(s: &str) -> Option<ParagraphAlignment> {
    match s {
        "left" => Some(ParagraphAlignment::Left),
        "center" => Some(ParagraphAlignment::Center),
        "right" => Some(ParagraphAlignment::Right),
        "justified" | "justify" => Some(ParagraphAlignment::Justified),
        _ => None,
    }
}

// ---------------------------------------------------------------------------
// Tests
// ---------------------------------------------------------------------------

#[cfg(test)]
mod tests {
    #![allow(clippy::panic_in_result_fn)]

    use super::*;

    type TestResult = std::result::Result<(), Box<dyn std::error::Error>>;

    #[test]
    fn derive_empty_document() {
        let doc = Document::new();
        let mut output = String::new();
        derive_style(&doc, &mut output);
        // Should have section properties at minimum (from default section)
        // May or may not have content depending on defaults
        assert!(!output.contains("# Paragraphs"));
    }

    #[test]
    fn derive_section_properties() {
        let mut doc = Document::new();
        doc.section_mut().set_page_size_twips(12240, 15840);
        doc.section_mut()
            .set_page_orientation(PageOrientation::Portrait);
        doc.section_mut().page_margins_mut().set_top_twips(1440);
        doc.section_mut().page_margins_mut().set_bottom_twips(1440);

        let mut output = String::new();
        derive_style(&doc, &mut output);
        assert!(output.contains("page-width: 12240"));
        assert!(output.contains("page-height: 15840"));
        assert!(output.contains("orientation: portrait"));
        assert!(output.contains("margin-top: 1440"));
        assert!(output.contains("margin-bottom: 1440"));
    }

    #[test]
    fn derive_paragraph_style() {
        let mut doc = Document::new();
        let p = doc.add_paragraph("Hello");
        p.set_style_id("Heading1");
        p.set_alignment(ParagraphAlignment::Center);

        let mut output = String::new();
        derive_style(&doc, &mut output);
        assert!(output.contains("[p1] style=\"Heading1\", align=center"));
    }

    #[test]
    fn derive_paragraph_spacing() {
        let mut doc = Document::new();
        let p = doc.add_paragraph("Test");
        p.set_spacing_after_twips(240);
        p.set_line_spacing_twips(276);

        let mut output = String::new();
        derive_style(&doc, &mut output);
        assert!(output.contains("spacing-after=240"));
        assert!(output.contains("line-spacing=276"));
    }

    #[test]
    fn apply_section_margins() -> TestResult {
        let mut doc = Document::new();

        let body = "# Section\nmargin-top: 720\nmargin-bottom: 720\n";
        apply_style(body, &mut doc)?;

        assert_eq!(doc.section().page_margins().top_twips(), Some(720));
        assert_eq!(doc.section().page_margins().bottom_twips(), Some(720));

        Ok(())
    }

    #[test]
    fn apply_paragraph_style_id() -> TestResult {
        let mut doc = Document::new();
        doc.add_paragraph("test");

        let body = "# Paragraphs\n[p1] style=\"Normal\", align=center\n";
        let result = apply_style(body, &mut doc)?;

        assert_eq!(result.cells_updated, 1);

        let para = &doc.paragraphs()[0];
        assert_eq!(para.style_id(), Some("Normal"));
        assert_eq!(para.alignment(), Some(ParagraphAlignment::Center));

        Ok(())
    }

    #[test]
    fn apply_paragraph_indent() -> TestResult {
        let mut doc = Document::new();
        doc.add_paragraph("test");

        let body = "[p1] indent-left=720, indent-hanging=360\n";
        apply_style(body, &mut doc)?;

        let para = &doc.paragraphs()[0];
        assert_eq!(para.indent_left_twips(), Some(720));
        assert_eq!(para.indent_hanging_twips(), Some(360));

        Ok(())
    }

    #[test]
    fn apply_paragraph_keep_flags() -> TestResult {
        let mut doc = Document::new();
        doc.add_paragraph("test");

        let body = "[p1] keep-next, keep-lines, page-break-before\n";
        apply_style(body, &mut doc)?;

        let para = &doc.paragraphs()[0];
        assert!(para.keep_next());
        assert!(para.keep_lines());
        assert!(para.page_break_before());

        Ok(())
    }

    #[test]
    fn apply_out_of_range_paragraph() -> TestResult {
        let mut doc = Document::new();
        doc.add_paragraph("test");

        let body = "[p99] align=center\n";
        let result = apply_style(body, &mut doc)?;

        assert_eq!(result.warnings.len(), 1);
        assert!(result.warnings[0].contains("out of range"));

        Ok(())
    }
}
