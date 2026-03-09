//! pptx style mode: derive and apply slide background + shape geometry/fill.
//!
//! Style mode emits layout/formatting properties without touching text content.
//! Properties use additive semantics: properties present in the IR are applied;
//! properties absent are left unchanged.

use crate::{ApplyResult, IrError, Result};
use offidized_pptx::{Presentation, ShapeGeometry, SlideBackground};

// ---------------------------------------------------------------------------
// Derive
// ---------------------------------------------------------------------------

/// Append the pptx style-mode body to `output`.
pub(crate) fn derive_style(prs: &Presentation, output: &mut String) {
    for (i, slide) in prs.slides().iter().enumerate() {
        let slide_num = i + 1;

        output.push('\n');
        output.push_str(&format!("--- slide {slide_num} ---\n"));

        // Background
        if let Some(SlideBackground::Solid(color)) = slide.background() {
            output.push_str(&format!("background: #{color}\n"));
        }

        // Shapes
        output.push('\n');
        for shape in slide.shapes() {
            let name = shape.name();
            let anchor = derive_shape_anchor(name);

            let mut props = Vec::new();

            if let Some(geo) = shape.geometry() {
                props.push(format!("x={}", geo.x()));
                props.push(format!("y={}", geo.y()));
                props.push(format!("w={}", geo.cx()));
                props.push(format!("h={}", geo.cy()));
            }

            if let Some(fill) = shape.solid_fill_srgb() {
                props.push(format!("fill=#{fill}"));
            }

            if let Some(rotation) = shape.rotation() {
                if rotation != 0 {
                    props.push(format!("rotation={rotation}"));
                }
            }

            if !props.is_empty() {
                output.push_str(&format!("{anchor} {}\n", props.join(", ")));
            }
        }
    }
}

/// Format a shape anchor string from its name.
fn derive_shape_anchor(name: &str) -> String {
    let lower = name.to_lowercase();
    if lower.contains("title") && !lower.contains("subtitle") {
        "[title]".to_string()
    } else if lower.contains("subtitle") {
        "[subtitle]".to_string()
    } else {
        format!("[shape \"{name}\"]")
    }
}

// ---------------------------------------------------------------------------
// Apply
// ---------------------------------------------------------------------------

/// Apply the style IR body to a presentation, updating layout/formatting.
pub(crate) fn apply_style(body: &str, prs: &mut Presentation) -> Result<ApplyResult> {
    let mut result = ApplyResult::default();
    let mut current_slide: Option<usize> = None;

    for line in body.lines() {
        let line = line.trim();

        // Skip blank/comment lines
        if line.is_empty() || line.starts_with('#') {
            continue;
        }

        // Slide header: --- slide N ---
        if let Some(num) = parse_slide_header(line) {
            current_slide = Some(num);
            continue;
        }

        let Some(slide_num) = current_slide else {
            continue;
        };

        let slide_idx = slide_num.checked_sub(1).ok_or_else(|| {
            IrError::InvalidBody(format!("slide number must be >= 1: {slide_num}"))
        })?;

        // Background line
        if let Some(rest) = line.strip_prefix("background: ") {
            let color = rest.trim().trim_start_matches('#');
            if let Some(slide) = prs.slide_mut(slide_idx) {
                slide.set_background(SlideBackground::Solid(color.to_string()));
            }
            continue;
        }

        // Shape property line: [anchor] props...
        if let Some((anchor, props_str)) = parse_shape_line(line) {
            let slide = match prs.slide_mut(slide_idx) {
                Some(s) => s,
                None => {
                    result.warnings.push(format!("slide {slide_num} not found"));
                    continue;
                }
            };

            let shape = match &anchor {
                ShapeAnchor::Title => slide.shapes_mut().iter_mut().find(|s| {
                    let lower = s.name().to_lowercase();
                    lower.contains("title") && !lower.contains("subtitle")
                }),
                ShapeAnchor::Subtitle => slide
                    .shapes_mut()
                    .iter_mut()
                    .find(|s| s.name().to_lowercase().contains("subtitle")),
                ShapeAnchor::Named(name) => slide.find_shape_by_name_mut(name),
                ShapeAnchor::Index(idx) => slide.shapes_mut().get_mut(*idx),
            };

            let Some(shape) = shape else {
                result.warnings.push(format!("shape not found: {anchor:?}"));
                continue;
            };

            apply_shape_properties(shape, props_str)?;
            result.cells_updated += 1;
        }
    }

    Ok(result)
}

/// Parse a slide header line.
fn parse_slide_header(line: &str) -> Option<usize> {
    let line = line.trim();
    let rest = line.strip_prefix("--- slide ")?;
    let num_end = rest.find(|c: char| !c.is_ascii_digit())?;
    let num: usize = rest[..num_end].parse().ok()?;
    if rest.ends_with("---") {
        Some(num)
    } else {
        None
    }
}

#[derive(Debug)]
enum ShapeAnchor {
    Title,
    Subtitle,
    Named(String),
    Index(usize),
}

/// Parse a shape property line like `[title] x=914400, y=2286000`.
fn parse_shape_line(line: &str) -> Option<(ShapeAnchor, &str)> {
    let trimmed = line.trim();

    if let Some(rest) = trimmed.strip_prefix("[title]") {
        return Some((ShapeAnchor::Title, rest.trim()));
    }
    if let Some(rest) = trimmed.strip_prefix("[subtitle]") {
        return Some((ShapeAnchor::Subtitle, rest.trim()));
    }

    // [shape "Name"] props
    if let Some(rest) = trimmed.strip_prefix("[shape \"") {
        if let Some(end) = rest.find("\"]") {
            let name = &rest[..end];
            let props = rest[end + 2..].trim();
            return Some((ShapeAnchor::Named(name.to_string()), props));
        }
    }

    // [shape #N] props
    if let Some(rest) = trimmed.strip_prefix("[shape #") {
        if let Some(end) = rest.find(']') {
            let idx: usize = rest[..end].parse().ok()?;
            let props = rest[end + 1..].trim();
            return Some((ShapeAnchor::Index(idx), props));
        }
    }

    None
}

/// Apply property strings to a shape.
fn apply_shape_properties(shape: &mut offidized_pptx::Shape, props_str: &str) -> Result<()> {
    let mut x: Option<i64> = None;
    let mut y: Option<i64> = None;
    let mut w: Option<i64> = None;
    let mut h: Option<i64> = None;

    for prop in split_properties(props_str) {
        let prop = prop.trim();
        if prop.is_empty() {
            continue;
        }

        if let Some((key, value)) = prop.split_once('=') {
            let key = key.trim();
            let value = value.trim();

            match key {
                "x" => {
                    if let Ok(v) = value.parse::<i64>() {
                        x = Some(v);
                    }
                }
                "y" => {
                    if let Ok(v) = value.parse::<i64>() {
                        y = Some(v);
                    }
                }
                "w" => {
                    if let Ok(v) = value.parse::<i64>() {
                        w = Some(v);
                    }
                }
                "h" => {
                    if let Ok(v) = value.parse::<i64>() {
                        h = Some(v);
                    }
                }
                "fill" => {
                    let color = value.trim_start_matches('#');
                    shape.set_solid_fill_srgb(color);
                }
                "rotation" => {
                    if let Ok(v) = value.parse::<i32>() {
                        shape.set_rotation(v);
                    }
                }
                _ => {
                    // Unknown property — ignore
                }
            }
        }
    }

    // Apply geometry: merge with existing or create new
    if x.is_some() || y.is_some() || w.is_some() || h.is_some() {
        let current = shape
            .geometry()
            .unwrap_or_else(|| ShapeGeometry::new(0, 0, 0, 0));
        let geo = ShapeGeometry::new(
            x.unwrap_or_else(|| current.x()),
            y.unwrap_or_else(|| current.y()),
            w.unwrap_or_else(|| current.cx()),
            h.unwrap_or_else(|| current.cy()),
        );
        shape.set_geometry(geo);
    }

    Ok(())
}

/// Quote-aware comma splitting (same logic as xlsx/docx).
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

// ---------------------------------------------------------------------------
// Tests
// ---------------------------------------------------------------------------

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn parse_slide_header_valid() {
        assert_eq!(parse_slide_header("--- slide 1 ---"), Some(1));
        assert_eq!(parse_slide_header("--- slide 12 ---"), Some(12));
    }

    #[test]
    fn parse_slide_header_invalid() {
        assert_eq!(parse_slide_header("Not a header"), None);
        assert_eq!(parse_slide_header("--- slide ---"), None);
    }

    #[test]
    fn parse_shape_line_title() {
        let (anchor, props) = parse_shape_line("[title] x=100, y=200").expect("parse");
        assert!(matches!(anchor, ShapeAnchor::Title));
        assert_eq!(props, "x=100, y=200");
    }

    #[test]
    fn parse_shape_line_named() {
        let (anchor, props) = parse_shape_line("[shape \"Key Info\"] fill=#FFFFFF").expect("parse");
        match anchor {
            ShapeAnchor::Named(name) => assert_eq!(name, "Key Info"),
            other => panic!("expected Named, got {other:?}"),
        }
        assert_eq!(props, "fill=#FFFFFF");
    }

    #[test]
    fn parse_shape_line_indexed() {
        let (anchor, props) = parse_shape_line("[shape #3] x=0").expect("parse");
        match anchor {
            ShapeAnchor::Index(idx) => assert_eq!(idx, 3),
            other => panic!("expected Index, got {other:?}"),
        }
        assert_eq!(props, "x=0");
    }

    #[test]
    fn derive_shape_anchor_detection() {
        assert_eq!(derive_shape_anchor("Title 1"), "[title]");
        assert_eq!(derive_shape_anchor("Subtitle 2"), "[subtitle]");
        assert_eq!(
            derive_shape_anchor("Content Placeholder 3"),
            "[shape \"Content Placeholder 3\"]"
        );
    }
}
