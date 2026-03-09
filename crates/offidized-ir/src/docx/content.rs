//! docx content mode: derive and apply paragraph text with markdown formatting.
//!
//! Content mode derives a markdown representation of the document body:
//! - `[pN]` anchors for each paragraph (1-based positional index)
//! - `[tN]` anchors for each table
//! - Heading levels via `#` prefixes
//! - Bold/italic via `**`/`*` markers
//! - Hyperlinks via `[text](url)`
//! - Bulleted/numbered lists via `-`/`1.`
//! - Block quotes via `>`
//! - Tables via markdown pipe syntax

use crate::{ApplyResult, Result};
use offidized_docx::{BodyItem, Document};

// ---------------------------------------------------------------------------
// Derive
// ---------------------------------------------------------------------------

/// Append the docx content-mode body to `output`.
pub(crate) fn derive_content(doc: &Document, output: &mut String) {
    let mut para_idx: usize = 0;
    let mut table_idx: usize = 0;

    for item in doc.body_items() {
        match item {
            BodyItem::Paragraph(para) => {
                para_idx += 1;
                output.push('\n');

                // Determine paragraph prefix based on type
                let heading_level = para.heading_level();
                let is_bullet = is_bullet_list(doc, para);
                let is_numbered = is_numbered_list(doc, para);
                let is_quote = is_block_quote(para);

                output.push_str(&format!("[p{para_idx}] "));

                // Heading prefix
                if let Some(level) = heading_level {
                    for _ in 0..level {
                        output.push('#');
                    }
                    output.push(' ');
                }

                // List prefix
                if is_bullet {
                    output.push_str("- ");
                } else if is_numbered {
                    output.push_str("1. ");
                }

                // Block quote prefix
                if is_quote {
                    output.push_str("> ");
                }

                // Render runs with inline formatting
                derive_runs(para.runs(), output);

                output.push('\n');
            }
            BodyItem::Table(table) => {
                table_idx += 1;
                output.push('\n');
                output.push_str(&format!("[t{table_idx}]\n"));
                derive_table(table, output);
            }
        }
    }
}

/// Render runs as markdown inline formatting.
fn derive_runs(runs: &[offidized_docx::Run], output: &mut String) {
    for run in runs {
        let text = run.text();
        if text.is_empty() {
            continue;
        }

        // Image placeholder
        if run.inline_image().is_some() || run.floating_image().is_some() {
            output.push_str("(image)");
            continue;
        }

        // Field code placeholder
        if run.field_code().is_some() {
            continue;
        }

        let bold = run.is_bold();
        let italic = run.is_italic();
        let hyperlink = run.hyperlink();

        if let Some(url) = hyperlink {
            output.push('[');
            output.push_str(text);
            output.push_str("](");
            output.push_str(url);
            output.push(')');
        } else if bold && italic {
            output.push_str("***");
            output.push_str(text);
            output.push_str("***");
        } else if bold {
            output.push_str("**");
            output.push_str(text);
            output.push_str("**");
        } else if italic {
            output.push('*');
            output.push_str(text);
            output.push('*');
        } else {
            output.push_str(text);
        }
    }
}

/// Render a table as a markdown pipe table.
fn derive_table(table: &offidized_docx::Table, output: &mut String) {
    let rows = table.rows();
    let cols = table.columns();

    if rows == 0 || cols == 0 {
        return;
    }

    for r in 0..rows {
        output.push('|');
        for c in 0..cols {
            let text = table.cell_text(r, c).unwrap_or("");
            // Escape pipes in cell text
            let escaped = text.replace('|', "\\|");
            output.push(' ');
            output.push_str(&escaped);
            output.push_str(" |");
        }
        output.push('\n');

        // Separator row after header
        if r == 0 {
            output.push('|');
            for _ in 0..cols {
                output.push_str("---|");
            }
            output.push('\n');
        }
    }
}

/// Check if a paragraph is a bullet list item.
fn is_bullet_list(doc: &Document, para: &offidized_docx::Paragraph) -> bool {
    let Some(num_id) = para.numbering_num_id() else {
        return false;
    };
    let ilvl = para.numbering_ilvl().unwrap_or(0);

    // Look up the numbering definition to check if it's a bullet
    if let Some(format) = resolve_list_format(doc, num_id, ilvl) {
        return format == "bullet";
    }

    false
}

/// Check if a paragraph is a numbered list item.
fn is_numbered_list(doc: &Document, para: &offidized_docx::Paragraph) -> bool {
    let Some(num_id) = para.numbering_num_id() else {
        return false;
    };
    let ilvl = para.numbering_ilvl().unwrap_or(0);

    if let Some(format) = resolve_list_format(doc, num_id, ilvl) {
        return format != "bullet";
    }

    // If we can't resolve, assume numbered (it has a numbering reference)
    true
}

/// Resolve the list format string for a numbering reference.
fn resolve_list_format(doc: &Document, num_id: u32, ilvl: u8) -> Option<String> {
    let inst = doc
        .numbering_instances()
        .iter()
        .find(|i| i.num_id() == num_id)?;
    let def = doc
        .numbering_definitions()
        .iter()
        .find(|d| d.abstract_num_id() == inst.abstract_num_id())?;
    let level = def.level(ilvl)?;
    Some(level.format().to_string())
}

/// Check if a paragraph has a block quote style.
fn is_block_quote(para: &offidized_docx::Paragraph) -> bool {
    matches!(
        para.style_id(),
        Some("IntenseQuote" | "Quote" | "BlockText"),
    )
}

// ---------------------------------------------------------------------------
// Apply
// ---------------------------------------------------------------------------

/// Apply the IR body to a document, updating paragraph text.
pub(crate) fn apply_content(body: &str, doc: &mut Document) -> Result<ApplyResult> {
    let mut result = ApplyResult::default();

    let sections = parse_body(body);

    for section in &sections {
        match section {
            BodySection::Paragraph { index, content } => {
                apply_paragraph(doc, *index, content, &mut result);
            }
            BodySection::Table { index, rows } => {
                apply_table(doc, *index, rows, &mut result);
            }
        }
    }

    Ok(result)
}

/// Parsed body section from the IR.
enum BodySection {
    Paragraph {
        index: usize, // 1-based
        content: ParagraphContent,
    },
    Table {
        index: usize, // 1-based
        rows: Vec<Vec<String>>,
    },
}

/// Parsed paragraph content.
struct ParagraphContent {
    heading_level: Option<u8>,
    is_bullet: bool,
    is_numbered: bool,
    is_quote: bool,
    runs: Vec<InlineRun>,
}

/// An inline run parsed from markdown.
struct InlineRun {
    text: String,
    bold: bool,
    italic: bool,
    hyperlink: Option<String>,
}

/// Parse the IR body into sections.
fn parse_body(body: &str) -> Vec<BodySection> {
    let mut sections = Vec::new();
    let mut current_table_index: Option<usize> = None;
    let mut table_rows: Vec<Vec<String>> = Vec::new();

    for line in body.lines() {
        let line = line.trim_end_matches('\r');

        // Table anchor: [tN]
        if let Some(idx) = parse_table_anchor(line) {
            // Flush any previous table
            if let Some(tidx) = current_table_index.take() {
                sections.push(BodySection::Table {
                    index: tidx,
                    rows: std::mem::take(&mut table_rows),
                });
            }
            current_table_index = Some(idx);
            continue;
        }

        // If we're in a table section, collect rows
        if current_table_index.is_some() {
            if line.starts_with('|') {
                // Skip separator rows (|---|---|)
                let trimmed = line.trim_matches('|').trim();
                if !trimmed.is_empty()
                    && trimmed
                        .chars()
                        .all(|c| c == '-' || c == '|' || c == ':' || c == ' ')
                {
                    continue;
                }
                // Parse table row
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
            // Non-table line ends the table section
            if let Some(tidx) = current_table_index.take() {
                sections.push(BodySection::Table {
                    index: tidx,
                    rows: std::mem::take(&mut table_rows),
                });
            }
        }

        // Paragraph anchor: [pN]
        if let Some((idx, content)) = parse_paragraph_line(line) {
            sections.push(BodySection::Paragraph {
                index: idx,
                content,
            });
            continue;
        }

        // Skip blank lines and unrecognized content
    }

    // Flush trailing table
    if let Some(tidx) = current_table_index {
        sections.push(BodySection::Table {
            index: tidx,
            rows: table_rows,
        });
    }

    sections
}

/// Parse a `[tN]` anchor line.
fn parse_table_anchor(line: &str) -> Option<usize> {
    let line = line.trim();
    let rest = line.strip_prefix("[t")?;
    let rest = rest.strip_suffix(']')?;
    rest.parse().ok()
}

/// Parse a `[pN] content...` line.
fn parse_paragraph_line(line: &str) -> Option<(usize, ParagraphContent)> {
    let line = line.trim();
    let rest = line.strip_prefix("[p")?;
    let bracket_end = rest.find(']')?;
    let idx: usize = rest[..bracket_end].parse().ok()?;
    let after = rest[bracket_end + 1..].trim_start();

    // Parse paragraph type from prefixes
    let (heading_level, after) = parse_heading_prefix(after);
    let (is_bullet, after) = parse_bullet_prefix(after);
    let (is_numbered, after) = parse_numbered_prefix(after);
    let (is_quote, after) = parse_quote_prefix(after);

    // Parse inline markdown into runs
    let runs = parse_inline_markdown(after);

    Some((
        idx,
        ParagraphContent {
            heading_level,
            is_bullet,
            is_numbered,
            is_quote,
            runs,
        },
    ))
}

fn parse_heading_prefix(s: &str) -> (Option<u8>, &str) {
    let mut level = 0u8;
    let bytes = s.as_bytes();
    while (level as usize) < bytes.len() && bytes[level as usize] == b'#' {
        level += 1;
    }
    if level > 0 && (level as usize) < bytes.len() && bytes[level as usize] == b' ' {
        (Some(level), &s[level as usize + 1..])
    } else {
        (None, s)
    }
}

fn parse_bullet_prefix(s: &str) -> (bool, &str) {
    if let Some(rest) = s.strip_prefix("- ") {
        (true, rest)
    } else {
        (false, s)
    }
}

fn parse_numbered_prefix(s: &str) -> (bool, &str) {
    // Match "N. " where N is one or more digits
    let bytes = s.as_bytes();
    let mut i = 0;
    while i < bytes.len() && bytes[i].is_ascii_digit() {
        i += 1;
    }
    if i > 0 && i + 1 < bytes.len() && bytes[i] == b'.' && bytes[i + 1] == b' ' {
        (true, &s[i + 2..])
    } else {
        (false, s)
    }
}

fn parse_quote_prefix(s: &str) -> (bool, &str) {
    if let Some(rest) = s.strip_prefix("> ") {
        (true, rest)
    } else {
        (false, s)
    }
}

/// Parse markdown inline formatting into runs.
fn parse_inline_markdown(s: &str) -> Vec<InlineRun> {
    let mut runs = Vec::new();
    let mut chars = s.char_indices().peekable();
    let mut current_text = String::new();

    while let Some(&(i, c)) = chars.peek() {
        match c {
            '*' => {
                // Flush current text
                if !current_text.is_empty() {
                    runs.push(InlineRun {
                        text: std::mem::take(&mut current_text),
                        bold: false,
                        italic: false,
                        hyperlink: None,
                    });
                }

                // Count asterisks
                let mut count = 0;
                while chars.peek().is_some_and(|&(_, ch)| ch == '*') {
                    chars.next();
                    count += 1;
                }

                let (bold, italic) = match count {
                    1 => (false, true),
                    2 => (true, false),
                    _ => (true, true), // 3+ = bold+italic
                };

                // Find closing marker
                let closing = "*".repeat(count);
                let rest = &s[i + count..];
                if let Some(end) = rest.find(&closing) {
                    let inner = &rest[..end];
                    runs.push(InlineRun {
                        text: inner.to_string(),
                        bold,
                        italic,
                        hyperlink: None,
                    });
                    // Skip past closing marker
                    let skip_to = i + count + end + count;
                    while chars.peek().is_some_and(|&(idx, _)| idx < skip_to) {
                        chars.next();
                    }
                } else {
                    // No closing marker found, treat asterisks as literal
                    current_text.push_str(&"*".repeat(count));
                }
            }
            '[' => {
                // Flush current text
                if !current_text.is_empty() {
                    runs.push(InlineRun {
                        text: std::mem::take(&mut current_text),
                        bold: false,
                        italic: false,
                        hyperlink: None,
                    });
                }

                // Try to parse [text](url)
                let rest = &s[i..];
                if let Some((text, url, consumed)) = parse_markdown_link(rest) {
                    runs.push(InlineRun {
                        text,
                        bold: false,
                        italic: false,
                        hyperlink: Some(url),
                    });
                    let skip_to = i + consumed;
                    while chars.peek().is_some_and(|&(idx, _)| idx < skip_to) {
                        chars.next();
                    }
                } else {
                    chars.next();
                    current_text.push('[');
                }
            }
            _ => {
                chars.next();
                current_text.push(c);
            }
        }
    }

    // Flush remaining text
    if !current_text.is_empty() {
        runs.push(InlineRun {
            text: current_text,
            bold: false,
            italic: false,
            hyperlink: None,
        });
    }

    // If no runs were parsed, ensure at least an empty one for the empty string case
    if runs.is_empty() && s.is_empty() {
        runs.push(InlineRun {
            text: String::new(),
            bold: false,
            italic: false,
            hyperlink: None,
        });
    }

    runs
}

/// Try to parse `[text](url)` at the start of `s`.
/// Returns (link_text, url, total_chars_consumed).
fn parse_markdown_link(s: &str) -> Option<(String, String, usize)> {
    let rest = s.strip_prefix('[')?;
    let bracket_close = rest.find(']')?;
    let text = &rest[..bracket_close];
    let after_bracket = &rest[bracket_close + 1..];
    let after_paren_open = after_bracket.strip_prefix('(')?;
    let paren_close = after_paren_open.find(')')?;
    let url = &after_paren_open[..paren_close];
    let consumed = 1 + bracket_close + 1 + 1 + paren_close + 1; // [text](url)
    Some((text.to_string(), url.to_string(), consumed))
}

/// Apply a parsed paragraph to the document.
fn apply_paragraph(
    doc: &mut Document,
    index: usize, // 1-based
    content: &ParagraphContent,
    result: &mut ApplyResult,
) {
    let para_idx = index - 1; // 0-based
    let paragraphs = doc.paragraphs_mut();

    if para_idx < paragraphs.len() {
        // Update existing paragraph
        let para = &mut paragraphs[para_idx];

        // Set heading if specified
        if let Some(level) = content.heading_level {
            para.set_style_id(format!("Heading{level}"));
        }

        // Rebuild runs from the parsed inline content
        rebuild_runs(para, &content.runs);

        result.cells_updated += 1;
    } else {
        // Append new paragraph
        let text: String = content.runs.iter().map(|r| r.text.as_str()).collect();

        let para = if let Some(level) = content.heading_level {
            doc.add_heading(&text, level)
        } else if content.is_bullet {
            doc.add_bulleted_paragraph(&text)
        } else if content.is_numbered {
            doc.add_numbered_paragraph(&text)
        } else {
            doc.add_paragraph(&text)
        };

        if content.is_quote {
            para.set_style_id("IntenseQuote");
        }

        // Apply inline formatting to the new paragraph's runs
        // Since add_paragraph creates a single run, we need to rebuild
        if content.runs.len() > 1
            || content
                .runs
                .iter()
                .any(|r| r.bold || r.italic || r.hyperlink.is_some())
        {
            rebuild_runs(para, &content.runs);
        }

        result.cells_created += 1;
    }
}

/// Rebuild the runs of a paragraph from parsed inline content.
fn rebuild_runs(para: &mut offidized_docx::Paragraph, inline_runs: &[InlineRun]) {
    // Concatenate all text for set_text, then re-add formatting via runs
    // set_text replaces all runs with a single plain run
    let full_text: String = inline_runs.iter().map(|r| r.text.as_str()).collect();
    para.set_text(&full_text);

    // If all runs are plain text, we're done
    if inline_runs.len() <= 1
        && !inline_runs
            .iter()
            .any(|r| r.bold || r.italic || r.hyperlink.is_some())
    {
        return;
    }

    // Need to create separate runs with formatting.
    // Clear current runs and rebuild.
    para.set_text(""); // Clear runs
    for run_spec in inline_runs {
        if let Some(ref url) = run_spec.hyperlink {
            let run = para.add_hyperlink(&run_spec.text, url);
            if run_spec.bold {
                run.set_bold(true);
            }
            if run_spec.italic {
                run.set_italic(true);
            }
        } else {
            let run = para.add_run(&run_spec.text);
            if run_spec.bold {
                run.set_bold(true);
            }
            if run_spec.italic {
                run.set_italic(true);
            }
        }
    }
}

/// Apply a parsed table to the document.
fn apply_table(
    doc: &mut Document,
    index: usize, // 1-based
    rows: &[Vec<String>],
    result: &mut ApplyResult,
) {
    let table_idx = index - 1; // 0-based
    let tables = doc.tables_mut();

    if table_idx < tables.len() {
        let table = &mut tables[table_idx];
        for (r, row) in rows.iter().enumerate() {
            for (c, cell_text) in row.iter().enumerate() {
                if table.set_cell_text(r, c, cell_text) {
                    result.cells_updated += 1;
                }
            }
        }
    }
    // If table doesn't exist, we could create one, but that's complex.
    // For now, skip with a warning (the result captures nothing).
}

// ---------------------------------------------------------------------------
// Tests
// ---------------------------------------------------------------------------

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn parse_heading_levels() {
        let (level, rest) = parse_heading_prefix("# Title");
        assert_eq!(level, Some(1));
        assert_eq!(rest, "Title");

        let (level, rest) = parse_heading_prefix("## Subtitle");
        assert_eq!(level, Some(2));
        assert_eq!(rest, "Subtitle");

        let (level, rest) = parse_heading_prefix("Not a heading");
        assert_eq!(level, None);
        assert_eq!(rest, "Not a heading");
    }

    #[test]
    fn parse_bullet_and_numbered() {
        let (is_bullet, rest) = parse_bullet_prefix("- Item");
        assert!(is_bullet);
        assert_eq!(rest, "Item");

        let (is_num, rest) = parse_numbered_prefix("1. First");
        assert!(is_num);
        assert_eq!(rest, "First");

        let (is_num, rest) = parse_numbered_prefix("12. Twelfth");
        assert!(is_num);
        assert_eq!(rest, "Twelfth");
    }

    #[test]
    fn parse_quote() {
        let (is_quote, rest) = parse_quote_prefix("> Quoted text");
        assert!(is_quote);
        assert_eq!(rest, "Quoted text");
    }

    #[test]
    fn inline_markdown_plain() {
        let runs = parse_inline_markdown("Hello world");
        assert_eq!(runs.len(), 1);
        assert_eq!(runs[0].text, "Hello world");
        assert!(!runs[0].bold);
        assert!(!runs[0].italic);
    }

    #[test]
    fn inline_markdown_bold() {
        let runs = parse_inline_markdown("Hello **bold** world");
        assert_eq!(runs.len(), 3);
        assert_eq!(runs[0].text, "Hello ");
        assert_eq!(runs[1].text, "bold");
        assert!(runs[1].bold);
        assert!(!runs[1].italic);
        assert_eq!(runs[2].text, " world");
    }

    #[test]
    fn inline_markdown_italic() {
        let runs = parse_inline_markdown("Hello *italic* world");
        assert_eq!(runs.len(), 3);
        assert_eq!(runs[1].text, "italic");
        assert!(!runs[1].bold);
        assert!(runs[1].italic);
    }

    #[test]
    fn inline_markdown_bold_italic() {
        let runs = parse_inline_markdown("***both***");
        assert_eq!(runs.len(), 1);
        assert_eq!(runs[0].text, "both");
        assert!(runs[0].bold);
        assert!(runs[0].italic);
    }

    #[test]
    fn inline_markdown_link() {
        let runs = parse_inline_markdown("Click [here](https://example.com) please");
        assert_eq!(runs.len(), 3);
        assert_eq!(runs[0].text, "Click ");
        assert_eq!(runs[1].text, "here");
        assert_eq!(runs[1].hyperlink, Some("https://example.com".to_string()),);
        assert_eq!(runs[2].text, " please");
    }

    #[test]
    fn inline_markdown_mixed() {
        let runs = parse_inline_markdown("Normal **bold** and *italic* [link](url)");
        assert_eq!(runs.len(), 6);
        assert_eq!(runs[0].text, "Normal ");
        assert!(runs[1].bold);
        assert_eq!(runs[2].text, " and ");
        assert!(runs[3].italic);
        assert_eq!(runs[4].text, " ");
        assert!(runs[5].hyperlink.is_some());
    }

    #[test]
    fn parse_table_anchor_valid() {
        assert_eq!(parse_table_anchor("[t1]"), Some(1));
        assert_eq!(parse_table_anchor("[t42]"), Some(42));
        assert_eq!(parse_table_anchor("  [t3]  "), Some(3));
    }

    #[test]
    fn parse_paragraph_line_valid() {
        let (idx, content) = parse_paragraph_line("[p1] # Title").expect("should parse");
        assert_eq!(idx, 1);
        assert_eq!(content.heading_level, Some(1));
        assert_eq!(content.runs[0].text, "Title");
    }

    #[test]
    fn parse_paragraph_line_bullet() {
        let (idx, content) = parse_paragraph_line("[p4] - Bullet item").expect("should parse");
        assert_eq!(idx, 4);
        assert!(content.is_bullet);
        assert_eq!(content.runs[0].text, "Bullet item");
    }

    #[test]
    fn derive_basic_document() {
        let mut doc = Document::new();
        doc.add_heading("Title", 1);
        doc.add_paragraph("Normal text");
        doc.add_bulleted_paragraph("Bullet");

        let mut output = String::new();
        derive_content(&doc, &mut output);

        assert!(output.contains("[p1] # Title"));
        assert!(output.contains("[p2] Normal text"));
        // Bullet detection depends on numbering definitions being set up
    }

    #[test]
    fn derive_table() {
        let mut doc = Document::new();
        doc.add_paragraph("Before table");
        let table = doc.add_table(2, 2);
        table.set_cell_text(0, 0, "A");
        table.set_cell_text(0, 1, "B");
        table.set_cell_text(1, 0, "C");
        table.set_cell_text(1, 1, "D");
        doc.add_paragraph("After table");

        let mut output = String::new();
        derive_content(&doc, &mut output);

        assert!(output.contains("[p1] Before table"));
        assert!(output.contains("[t1]"));
        assert!(output.contains("| A | B |"));
        assert!(output.contains("| C | D |"));
        assert!(output.contains("[p2] After table"));
    }
}
