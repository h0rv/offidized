//! pptx content mode: derive and apply slide/shape text.
//!
//! Content mode extracts text from slides in a structured format:
//! - `--- slide N [Layout] ---` delimiters
//! - `[title]` / `[subtitle]` for well-known placeholders
//! - `[shape "Name"]` for named shapes
//! - `[shape #N]` for unnamed shapes (by index)
//! - `(chart)` / `(image)` / `(table)` for non-text content
//! - `[notes]` for slide notes

use crate::{ApplyResult, Result};
use offidized_pptx::{PlaceholderType, Presentation, Shape};

// ---------------------------------------------------------------------------
// Derive
// ---------------------------------------------------------------------------

/// Append the pptx content-mode body to `output`.
pub(crate) fn derive_content(prs: &Presentation, output: &mut String) {
    for (slide_idx, slide) in prs.slides().iter().enumerate() {
        let slide_num = slide_idx + 1;

        // Get layout name
        let layout_name = slide
            .layout_reference()
            .and_then(|(mi, li)| prs.layout(mi, li))
            .map(|l| l.name().to_string())
            .unwrap_or_else(|| "Unknown".to_string());

        output.push('\n');
        output.push_str(&format!("--- slide {slide_num} [{layout_name}] ---\n"));

        // Derive shapes
        for (shape_idx, shape) in slide.shapes().iter().enumerate() {
            derive_shape(shape, shape_idx, output);
        }

        // Derive grouped shapes
        for group in slide.grouped_shapes() {
            for (shape_idx, shape) in group.shapes().iter().enumerate() {
                derive_shape(shape, shape_idx, output);
            }
        }

        // Derive tables
        for table in slide.tables() {
            output.push('\n');
            // Find the shape name for the table if available
            output.push_str("[table]\n");
            derive_pptx_table(table, output);
        }

        // Derive notes
        if let Some(notes) = slide.notes_text() {
            if !notes.is_empty() {
                output.push('\n');
                output.push_str("[notes] ");
                output.push_str(notes);
                output.push('\n');
            }
        }
    }
}

/// Derive a single shape's text content.
fn derive_shape(shape: &Shape, shape_idx: usize, output: &mut String) {
    // Skip shapes with no text content
    let has_text = shape
        .paragraphs()
        .iter()
        .any(|p| p.runs().iter().any(|r| !r.text().is_empty()));

    // Determine the shape anchor
    let anchor = match shape.placeholder_type() {
        Some(PlaceholderType::Title | PlaceholderType::CenteredTitle) => "[title]".to_string(),
        Some(PlaceholderType::Subtitle) => "[subtitle]".to_string(),
        Some(PlaceholderType::Body | PlaceholderType::Object) => {
            let name = shape.name();
            if name.is_empty() {
                format!("[shape #{}]", shape_idx)
            } else {
                format!("[shape \"{name}\"]")
            }
        }
        Some(PlaceholderType::Chart) => {
            if !has_text {
                output.push_str("\n(chart)\n");
                return;
            }
            format!("[shape \"{}\"]", shape.name())
        }
        Some(PlaceholderType::Table) => {
            output.push_str("\n(table)\n");
            return;
        }
        Some(PlaceholderType::Media | PlaceholderType::ClipArt) => {
            output.push_str("\n(media)\n");
            return;
        }
        Some(
            PlaceholderType::DateAndTime
            | PlaceholderType::SlideNumber
            | PlaceholderType::Footer
            | PlaceholderType::Header,
        ) => {
            // Skip metadata placeholders
            return;
        }
        Some(_) => {
            let name = shape.name();
            if name.is_empty() {
                format!("[shape #{}]", shape_idx)
            } else {
                format!("[shape \"{name}\"]")
            }
        }
        None => {
            if !has_text {
                // Check for image/media
                if shape.is_connector() || shape.is_smartart() {
                    return;
                }
                // Shape with no text and no placeholder - likely an image or decorative shape
                return;
            }
            let name = shape.name();
            if name.is_empty() {
                format!("[shape #{}]", shape_idx)
            } else {
                format!("[shape \"{name}\"]")
            }
        }
    };

    // Collect text from paragraphs
    let paragraphs: Vec<String> = shape
        .paragraphs()
        .iter()
        .map(|para| {
            let text: String = para.runs().iter().map(|r| r.text()).collect();
            let is_bullet =
                para.properties().bullet.style.is_some() || para.level().unwrap_or(0) > 0;
            if is_bullet && !text.is_empty() {
                format!("- {text}")
            } else {
                text
            }
        })
        .collect();

    // Skip if all paragraphs are empty
    if paragraphs.iter().all(|p| p.is_empty()) {
        return;
    }

    output.push('\n');
    output.push_str(&anchor);

    // For title/subtitle, emit on same line
    if anchor == "[title]" || anchor == "[subtitle]" {
        let text: String = paragraphs.join("\n");
        let trimmed = text.trim();
        if !trimmed.is_empty() {
            output.push(' ');
            output.push_str(trimmed);
        }
        output.push('\n');
    } else {
        // For other shapes, emit paragraphs on subsequent lines
        output.push('\n');
        for para in &paragraphs {
            output.push_str(para);
            output.push('\n');
        }
    }
}

/// Render a pptx table as markdown.
fn derive_pptx_table(table: &offidized_pptx::Table, output: &mut String) {
    let rows = table.rows();
    let cols = table.cols();

    if rows == 0 || cols == 0 {
        return;
    }

    for r in 0..rows {
        output.push('|');
        for c in 0..cols {
            let text = table.cell_text(r, c).unwrap_or("");
            let escaped = text.replace('|', "\\|");
            output.push(' ');
            output.push_str(&escaped);
            output.push_str(" |");
        }
        output.push('\n');

        if r == 0 {
            output.push('|');
            for _ in 0..cols {
                output.push_str("---|");
            }
            output.push('\n');
        }
    }
}

// ---------------------------------------------------------------------------
// Apply
// ---------------------------------------------------------------------------

/// Apply the IR body to a presentation, updating shape text.
pub(crate) fn apply_content(body: &str, prs: &mut Presentation) -> Result<ApplyResult> {
    let mut result = ApplyResult::default();

    let slides = parse_body(body);

    for slide_section in &slides {
        let slide_idx = slide_section.index - 1; // 0-based
        let Some(slide) = prs.slide_mut(slide_idx) else {
            result
                .warnings
                .push(format!("slide {} not found", slide_section.index));
            continue;
        };

        for item in &slide_section.items {
            match item {
                SlideItem::Shape { anchor, text } => {
                    if apply_shape_text(slide, anchor, text) {
                        result.cells_updated += 1;
                    } else {
                        result.warnings.push(format!(
                            "slide {}: shape {:?} not found",
                            slide_section.index, anchor,
                        ));
                    }
                }
                SlideItem::Table { rows } => {
                    // Apply table text to the first table on the slide
                    if let Some(table) = slide.tables_mut().first_mut() {
                        for (r, row) in rows.iter().enumerate() {
                            for (c, cell_text) in row.iter().enumerate() {
                                if table.set_cell_text(r, c, cell_text) {
                                    result.cells_updated += 1;
                                }
                            }
                        }
                    }
                }
                SlideItem::Notes { text } => {
                    slide.set_notes_text(text);
                    result.cells_updated += 1;
                }
            }
        }
    }

    Ok(result)
}

/// Apply text to a shape matched by anchor.
fn apply_shape_text(slide: &mut offidized_pptx::Slide, anchor: &ShapeAnchor, text: &str) -> bool {
    match anchor {
        ShapeAnchor::Title => {
            if let Some(shape) = slide.placeholder_mut(PlaceholderType::Title) {
                set_shape_text(shape, text);
                return true;
            }
            if let Some(shape) = slide.placeholder_mut(PlaceholderType::CenteredTitle) {
                set_shape_text(shape, text);
                return true;
            }
            false
        }
        ShapeAnchor::Subtitle => {
            if let Some(shape) = slide.placeholder_mut(PlaceholderType::Subtitle) {
                set_shape_text(shape, text);
                return true;
            }
            false
        }
        ShapeAnchor::Named(name) => {
            // Find shape by name
            for shape in slide.shapes_mut() {
                if shape.name() == name.as_str() {
                    set_shape_text(shape, text);
                    return true;
                }
            }
            false
        }
        ShapeAnchor::Index(idx) => {
            let shapes = slide.shapes_mut();
            if *idx < shapes.len() {
                set_shape_text(&mut shapes[*idx], text);
                return true;
            }
            false
        }
    }
}

/// Set the text content of a shape, preserving formatting of existing runs where possible.
fn set_shape_text(shape: &mut Shape, text: &str) {
    let lines: Vec<&str> = text.lines().collect();

    if lines.is_empty() {
        return;
    }

    // First pass: update existing paragraphs
    let para_count = shape.paragraph_count();
    for (i, line) in lines.iter().enumerate() {
        let line_text = line.strip_prefix("- ").unwrap_or(line);

        if i < para_count {
            let paragraphs = shape.paragraphs_mut();
            let runs = paragraphs[i].runs_mut();
            if runs.is_empty() {
                paragraphs[i].add_run(line_text);
            } else if runs.len() == 1 {
                runs[0].set_text(line_text);
            } else {
                runs[0].set_text(line_text);
                for run in &mut runs[1..] {
                    run.set_text("");
                }
            }
        }
    }

    // Second pass: add new paragraphs beyond existing count
    for line in lines.iter().skip(para_count) {
        let line_text = line.strip_prefix("- ").unwrap_or(line);
        shape.add_paragraph_with_text(line_text);
    }
}

// ---------------------------------------------------------------------------
// Body parser
// ---------------------------------------------------------------------------

/// A parsed slide section from the IR.
struct SlideSection {
    index: usize, // 1-based
    items: Vec<SlideItem>,
}

/// Items within a slide section.
enum SlideItem {
    Shape { anchor: ShapeAnchor, text: String },
    Table { rows: Vec<Vec<String>> },
    Notes { text: String },
}

/// How a shape is identified.
#[derive(Debug)]
enum ShapeAnchor {
    Title,
    Subtitle,
    Named(String),
    Index(usize),
}

/// Parse the IR body into slide sections.
fn parse_body(body: &str) -> Vec<SlideSection> {
    let mut slides = Vec::new();
    let mut current_slide: Option<SlideSection> = None;
    let mut current_shape: Option<(ShapeAnchor, Vec<String>)> = None;
    let mut in_table = false;
    let mut table_rows: Vec<Vec<String>> = Vec::new();

    for line in body.lines() {
        let line = line.trim_end_matches('\r');

        // Slide header: --- slide N [Layout] ---
        if let Some(idx) = parse_slide_header(line) {
            // Flush current shape and table into current slide
            if let Some(ref mut slide) = current_slide {
                flush_shape_into(&mut current_shape, slide);
                if in_table && !table_rows.is_empty() {
                    slide.items.push(SlideItem::Table {
                        rows: std::mem::take(&mut table_rows),
                    });
                }
                in_table = false;
            }
            // Flush current slide
            if let Some(slide) = current_slide.take() {
                slides.push(slide);
            }
            current_slide = Some(SlideSection {
                index: idx,
                items: Vec::new(),
            });
            continue;
        }

        let Some(ref mut slide) = current_slide else {
            continue;
        };

        // Table header
        if line.trim() == "[table]" {
            flush_shape_into(&mut current_shape, slide);
            in_table = true;
            table_rows.clear();
            continue;
        }

        // In table: collect rows
        if in_table {
            if line.starts_with('|') {
                let trimmed = line.trim_matches('|').trim();
                if !trimmed.is_empty()
                    && trimmed
                        .chars()
                        .all(|c| c == '-' || c == '|' || c == ':' || c == ' ')
                {
                    continue;
                }
                let cells: Vec<String> = line
                    .split('|')
                    .filter(|s| !s.is_empty())
                    .map(|s| s.trim().replace("\\|", "|"))
                    .collect();
                if !cells.is_empty() {
                    table_rows.push(cells);
                }
                continue;
            }
            // End of table
            if !table_rows.is_empty() {
                slide.items.push(SlideItem::Table {
                    rows: std::mem::take(&mut table_rows),
                });
            }
            in_table = false;
            // Fall through to process this line
        }

        // Notes: [notes] text
        if let Some(rest) = line.strip_prefix("[notes] ") {
            flush_shape_into(&mut current_shape, slide);
            slide.items.push(SlideItem::Notes {
                text: rest.to_string(),
            });
            continue;
        }
        if line.trim() == "[notes]" {
            flush_shape_into(&mut current_shape, slide);
            continue;
        }

        // Shape anchors
        if let Some(anchor) = parse_shape_anchor(line) {
            flush_shape_into(&mut current_shape, slide);

            let after_anchor = extract_after_anchor(line);
            if !after_anchor.is_empty() {
                current_shape = Some((anchor, vec![after_anchor.to_string()]));
            } else {
                current_shape = Some((anchor, Vec::new()));
            }
            continue;
        }

        // Content lines (belong to current shape)
        if let Some((_, ref mut lines_vec)) = current_shape {
            if !line.is_empty() {
                lines_vec.push(line.to_string());
            }
        }
    }

    // Flush remaining state
    if let Some(ref mut slide) = current_slide {
        flush_shape_into(&mut current_shape, slide);
        if in_table && !table_rows.is_empty() {
            slide.items.push(SlideItem::Table { rows: table_rows });
        }
    }
    if let Some(slide) = current_slide {
        slides.push(slide);
    }

    slides
}

/// Parse a slide header like `--- slide 1 [Title Slide] ---`.
fn parse_slide_header(line: &str) -> Option<usize> {
    let line = line.trim();
    let rest = line.strip_prefix("--- slide ")?;
    // Find the space or [ after the number
    let num_end = rest.find(|c: char| !c.is_ascii_digit())?;
    let num: usize = rest[..num_end].parse().ok()?;
    // Verify it ends with ---
    if rest.ends_with("---") {
        Some(num)
    } else {
        None
    }
}

/// Parse a shape anchor from a line.
fn parse_shape_anchor(line: &str) -> Option<ShapeAnchor> {
    let trimmed = line.trim();

    if trimmed.starts_with("[title]") {
        return Some(ShapeAnchor::Title);
    }
    if trimmed.starts_with("[subtitle]") {
        return Some(ShapeAnchor::Subtitle);
    }

    // [shape "Name"]
    if let Some(rest) = trimmed.strip_prefix("[shape \"") {
        if let Some(end) = rest.find("\"]") {
            let name = &rest[..end];
            return Some(ShapeAnchor::Named(name.to_string()));
        }
    }

    // [shape #N]
    if let Some(rest) = trimmed.strip_prefix("[shape #") {
        if let Some(end) = rest.find(']') {
            if let Ok(idx) = rest[..end].parse::<usize>() {
                return Some(ShapeAnchor::Index(idx));
            }
        }
    }

    None
}

/// Extract the text after a shape anchor on the same line.
fn extract_after_anchor(line: &str) -> &str {
    let trimmed = line.trim();
    // Try common patterns
    for prefix in &["[title] ", "[subtitle] "] {
        if let Some(rest) = trimmed.strip_prefix(prefix) {
            return rest;
        }
    }
    // [shape "Name"] text or [shape #N] text
    if let Some(bracket_end) = trimmed.find(']') {
        let after = &trimmed[bracket_end + 1..];
        return after.trim_start();
    }
    ""
}

fn flush_shape_into(
    current_shape: &mut Option<(ShapeAnchor, Vec<String>)>,
    slide: &mut SlideSection,
) {
    if let Some((anchor, lines)) = current_shape.take() {
        let text = lines.join("\n");
        if !text.is_empty() {
            slide.items.push(SlideItem::Shape { anchor, text });
        }
    }
}

// ---------------------------------------------------------------------------
// Tests
// ---------------------------------------------------------------------------

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn parse_slide_header_valid() {
        assert_eq!(parse_slide_header("--- slide 1 [Title Slide] ---"), Some(1));
        assert_eq!(
            parse_slide_header("--- slide 12 [Title and Content] ---"),
            Some(12),
        );
        assert_eq!(parse_slide_header("--- slide 3 [Blank] ---"), Some(3));
    }

    #[test]
    fn parse_slide_header_invalid() {
        assert_eq!(parse_slide_header("Not a header"), None);
        assert_eq!(parse_slide_header("--- slide ---"), None);
    }

    #[test]
    fn parse_shape_anchor_title() {
        assert!(matches!(
            parse_shape_anchor("[title] Q4 Review"),
            Some(ShapeAnchor::Title),
        ));
    }

    #[test]
    fn parse_shape_anchor_subtitle() {
        assert!(matches!(
            parse_shape_anchor("[subtitle] December 2025"),
            Some(ShapeAnchor::Subtitle),
        ));
    }

    #[test]
    fn parse_shape_anchor_named() {
        match parse_shape_anchor("[shape \"Key Info\"]") {
            Some(ShapeAnchor::Named(name)) => assert_eq!(name, "Key Info"),
            other => panic!("expected Named, got {other:?}"),
        }
    }

    #[test]
    fn parse_shape_anchor_indexed() {
        match parse_shape_anchor("[shape #3]") {
            Some(ShapeAnchor::Index(idx)) => assert_eq!(idx, 3),
            other => panic!("expected Index, got {other:?}"),
        }
    }

    #[test]
    fn extract_after_anchor_title() {
        assert_eq!(extract_after_anchor("[title] My Title"), "My Title");
    }

    #[test]
    fn extract_after_anchor_shape() {
        assert_eq!(
            extract_after_anchor("[shape \"Info\"] Some text"),
            "Some text",
        );
    }

    #[test]
    fn parse_body_single_slide() {
        let body = "\n--- slide 1 [Title Slide] ---\n\n[title] Hello World\n[subtitle] Subtitle\n";
        let slides = parse_body(body);
        assert_eq!(slides.len(), 1);
        assert_eq!(slides[0].index, 1);
        assert_eq!(slides[0].items.len(), 2);
    }

    #[test]
    fn parse_body_shape_with_content() {
        let body = "\n--- slide 1 [Title and Content] ---\n\n[shape \"Key Info\"]\n- 15% growth\n- 12 contracts\n";
        let slides = parse_body(body);
        assert_eq!(slides.len(), 1);
        assert_eq!(slides[0].items.len(), 1);
        match &slides[0].items[0] {
            SlideItem::Shape { text, .. } => {
                assert!(text.contains("15% growth"));
                assert!(text.contains("12 contracts"));
            }
            _ => panic!("expected Shape"),
        }
    }
}
