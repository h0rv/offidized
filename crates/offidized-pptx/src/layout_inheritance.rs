//! Layout and Master Slide Inheritance
//!
//! PowerPoint slides inherit properties from their layout, and layouts inherit
//! from their master. This module implements the inheritance resolution logic
//! for placeholders, backgrounds, shapes, and text properties.
//!
//! Based on ShapeCrawler's inheritance patterns:
//! - ReferencedPShape: resolve shape transforms and properties
//! - ReferencedFont: resolve font sizes, colors, and styles
//! - Placeholder matching by type and/or index

use crate::shape::{PlaceholderType, Shape};
use offidized_opc::RawXmlNode;
use quick_xml::events::Event;
use quick_xml::Reader;

/// Result of a placeholder search in a shape tree (layout or master).
#[derive(Debug, Clone)]
pub struct ReferencedPlaceholder {
    /// The placeholder type found.
    pub placeholder_type: Option<PlaceholderType>,
    /// The placeholder index.
    pub placeholder_idx: Option<u32>,
    /// Whether this placeholder was found.
    pub found: bool,
}

/// Matches a placeholder by type and index in the target shape tree.
///
/// This implements the ShapeCrawler matching logic:
/// 1. Try to match both type and index
/// 2. If not found, try to match type only
/// 3. Handle special index values (e.g., 0xFFFFFFFF)
pub fn find_placeholder_in_shapes(
    shapes: &[Shape],
    source_type: Option<&PlaceholderType>,
    source_index: Option<u32>,
) -> Option<usize> {
    // First pass: match both type and index
    if let Some(idx) = find_by_type_and_index(shapes, source_type, source_index) {
        return Some(idx);
    }

    // Second pass: match type only
    if let Some(placeholder_type) = source_type {
        return find_by_type_only(shapes, placeholder_type);
    }

    // Special case: index 0xFFFFFFFF means "find placeholder with index 1"
    // Reference: https://answers.microsoft.com/en-us/msoffice/forum/all/placeholder-master/0d51dcec-f982-4098-b6b6-94785304607a?page=3
    if source_index == Some(0xFFFF_FFFF) {
        return shapes.iter().position(|s| s.placeholder_idx() == Some(1));
    }

    None
}

/// Match placeholder by both type and index.
fn find_by_type_and_index(
    shapes: &[Shape],
    source_type: Option<&PlaceholderType>,
    source_index: Option<u32>,
) -> Option<usize> {
    if source_type.is_none() && source_index.is_none() {
        return None;
    }

    shapes.iter().position(|shape| {
        let target_type = shape.placeholder_type();
        let target_index = shape.placeholder_idx();

        // Check type match (if present)
        let type_matches = match (source_type, target_type) {
            (Some(src), Some(tgt)) => are_types_compatible(src, tgt),
            (None, _) | (_, None) => true, // If either is None, don't filter by type
        };

        // Check index match (if present)
        let index_matches = match (source_index, target_index) {
            (Some(src), Some(tgt)) => src == tgt,
            (None, _) | (_, None) => true, // If either is None, don't filter by index
        };

        type_matches && index_matches
    })
}

/// Check if two placeholder types are compatible.
///
/// Implements ShapeCrawler's type matching rules:
/// - Title matches CenteredTitle and vice versa
/// - Body matches Body with same index
/// - Exact type matches
fn are_types_compatible(source: &PlaceholderType, target: &PlaceholderType) -> bool {
    match (source, target) {
        // Exact match
        (a, b) if a == b => true,
        // Title and CenteredTitle are interchangeable
        (PlaceholderType::Title, PlaceholderType::CenteredTitle)
        | (PlaceholderType::CenteredTitle, PlaceholderType::Title) => true,
        // No other cross-type matches
        _ => false,
    }
}

/// Match placeholder by type only.
fn find_by_type_only(shapes: &[Shape], source_type: &PlaceholderType) -> Option<usize> {
    shapes.iter().position(|shape| {
        if let Some(target_type) = shape.placeholder_type() {
            are_types_compatible(source_type, target_type)
        } else {
            false
        }
    })
}

/// Resolves a shape's transform from its layout or master placeholder.
///
/// If a shape is a placeholder on a slide, it may inherit positioning
/// and transform properties from the corresponding placeholder on the layout,
/// or from the master if not found on the layout.
#[derive(Debug, Clone)]
pub struct ResolvedTransform {
    /// X position in EMUs (English Metric Units).
    pub x: Option<i64>,
    /// Y position in EMUs.
    pub y: Option<i64>,
    /// Width in EMUs.
    pub width: Option<i64>,
    /// Height in EMUs.
    pub height: Option<i64>,
    /// Rotation in 60000ths of a degree.
    pub rotation: Option<i32>,
}

impl ResolvedTransform {
    pub fn empty() -> Self {
        Self {
            x: None,
            y: None,
            width: None,
            height: None,
            rotation: None,
        }
    }

    /// Merge with fallback values (use self's values, fall back to other).
    pub fn merge_with(&mut self, other: &ResolvedTransform) {
        if self.x.is_none() {
            self.x = other.x;
        }
        if self.y.is_none() {
            self.y = other.y;
        }
        if self.width.is_none() {
            self.width = other.width;
        }
        if self.height.is_none() {
            self.height = other.height;
        }
        if self.rotation.is_none() {
            self.rotation = other.rotation;
        }
    }
}

/// Resolves font properties from layout and master shapes.
///
/// Font properties (size, bold, color, family) can be inherited from:
/// 1. The placeholder on the layout
/// 2. The placeholder on the master
/// 3. The master text styles (bodyStyle, titleStyle, otherStyle)
#[derive(Debug, Clone)]
pub struct ResolvedFont {
    /// Font size in points.
    pub size: Option<f64>,
    /// Bold flag.
    pub bold: Option<bool>,
    /// Italic flag.
    pub italic: Option<bool>,
    /// Font family name.
    pub font_family: Option<String>,
    /// Font color (hex RGB).
    pub color: Option<String>,
}

impl ResolvedFont {
    pub fn empty() -> Self {
        Self {
            size: None,
            bold: None,
            italic: None,
            font_family: None,
            color: None,
        }
    }

    /// Merge with fallback values (use self's values, fall back to other).
    pub fn merge_with(&mut self, other: &ResolvedFont) {
        if self.size.is_none() {
            self.size = other.size;
        }
        if self.bold.is_none() {
            self.bold = other.bold;
        }
        if self.italic.is_none() {
            self.italic = other.italic;
        }
        if self.font_family.is_none() {
            self.font_family = other.font_family.clone();
        }
        if self.color.is_none() {
            self.color = other.color.clone();
        }
    }
}

/// Resolves background properties from layout or master.
#[derive(Debug, Clone)]
pub enum ResolvedBackground {
    /// No background (transparent).
    None,
    /// Solid color fill (hex RGB).
    SolidColor(String),
    /// Picture fill (relationship ID).
    Picture(String),
    /// Gradient fill (raw XML node for now).
    Gradient(RawXmlNode),
    /// Pattern fill (raw XML node for now).
    Pattern(RawXmlNode),
}

/// Inheritance resolver for a slide, layout, or master.
///
/// This struct provides methods to resolve inherited properties by walking
/// the slide → layout → master chain.
pub struct InheritanceResolver<'a> {
    /// Shapes from the layout (if available).
    pub layout_shapes: Option<&'a [Shape]>,
    /// Shapes from the master (if available).
    pub master_shapes: Option<&'a [Shape]>,
}

impl<'a> InheritanceResolver<'a> {
    /// Create a new resolver with layout and master shapes.
    pub fn new(layout_shapes: Option<&'a [Shape]>, master_shapes: Option<&'a [Shape]>) -> Self {
        Self {
            layout_shapes,
            master_shapes,
        }
    }

    /// Resolve a placeholder shape from layout or master.
    ///
    /// Given a placeholder on a slide, find the corresponding placeholder
    /// on the layout or master.
    pub fn resolve_placeholder(
        &self,
        placeholder_type: Option<&PlaceholderType>,
        placeholder_idx: Option<u32>,
    ) -> Option<&'a Shape> {
        // Try layout first
        if let Some(shapes) = self.layout_shapes {
            if let Some(idx) = find_placeholder_in_shapes(shapes, placeholder_type, placeholder_idx)
            {
                return Some(&shapes[idx]);
            }
        }

        // Fall back to master
        if let Some(shapes) = self.master_shapes {
            if let Some(idx) = find_placeholder_in_shapes(shapes, placeholder_type, placeholder_idx)
            {
                return Some(&shapes[idx]);
            }
        }

        None
    }

    /// Resolve transform properties by walking layout → master.
    pub fn resolve_transform(
        &self,
        placeholder_type: Option<&PlaceholderType>,
        placeholder_idx: Option<u32>,
    ) -> ResolvedTransform {
        let transform = ResolvedTransform::empty();

        // Try layout first
        if let Some(shapes) = self.layout_shapes {
            if let Some(_idx) =
                find_placeholder_in_shapes(shapes, placeholder_type, placeholder_idx)
            {
                // In a real implementation, we'd extract x/y/width/height from the shape
                // For now, this is a placeholder
            }
        }

        // Try master if layout didn't have complete info
        if let Some(shapes) = self.master_shapes {
            if let Some(_idx) =
                find_placeholder_in_shapes(shapes, placeholder_type, placeholder_idx)
            {
                // Extract transform from master shape
            }
        }

        transform
    }

    /// Resolve font properties by walking layout → master → text styles.
    pub fn resolve_font(
        &self,
        placeholder_type: Option<&PlaceholderType>,
        placeholder_idx: Option<u32>,
        _indent_level: usize,
    ) -> ResolvedFont {
        let font = ResolvedFont::empty();

        // Try layout placeholder
        if let Some(shapes) = self.layout_shapes {
            if let Some(_idx) =
                find_placeholder_in_shapes(shapes, placeholder_type, placeholder_idx)
            {
                // Extract font properties from layout shape
                // This would involve parsing the shape's text properties
            }
        }

        // Try master placeholder
        if let Some(shapes) = self.master_shapes {
            if let Some(_idx) =
                find_placeholder_in_shapes(shapes, placeholder_type, placeholder_idx)
            {
                // Extract font properties from master shape
            }
        }

        // Fall back to master text styles (bodyStyle, titleStyle, etc.)
        // This would require parsing the master's <p:txStyles> element

        font
    }
}

/// Parse a placeholder element from XML.
///
/// Extracts the type and index attributes from a `<p:ph>` element.
pub fn parse_placeholder_from_xml(xml: &str) -> (Option<PlaceholderType>, Option<u32>) {
    let mut reader = Reader::from_str(xml);
    reader.config_mut().trim_text(true);

    let mut placeholder_type = None;
    let mut placeholder_idx = None;

    loop {
        match reader.read_event() {
            Ok(Event::Start(e)) | Ok(Event::Empty(e)) => {
                let name = e.name();
                let local = std::str::from_utf8(name.as_ref()).unwrap_or("");

                if local.ends_with(":ph") || local == "ph" {
                    // Parse attributes
                    for attr in e.attributes().flatten() {
                        let key = std::str::from_utf8(attr.key.as_ref()).unwrap_or("");
                        let value = std::str::from_utf8(&attr.value).unwrap_or("");

                        match key {
                            "type" => {
                                placeholder_type = Some(PlaceholderType::from_xml(value));
                            }
                            "idx" => {
                                if let Ok(idx) = value.parse::<u32>() {
                                    placeholder_idx = Some(idx);
                                }
                            }
                            _ => {}
                        }
                    }
                }
            }
            Ok(Event::Eof) => break,
            Err(_) => break,
            _ => {}
        }
    }

    (placeholder_type, placeholder_idx)
}

#[cfg(test)]
mod tests {
    use super::*;

    fn make_shape_with_placeholder(name: &str, ptype: PlaceholderType, pidx: Option<u32>) -> Shape {
        let mut shape = Shape::new(name);
        shape.set_placeholder_type(ptype);
        if let Some(idx) = pidx {
            shape.set_placeholder_idx(idx);
        }
        shape
    }

    #[test]
    fn test_placeholder_matching_type_and_index() {
        let shapes = vec![
            make_shape_with_placeholder("Shape1", PlaceholderType::Body, Some(0)),
            make_shape_with_placeholder("Shape2", PlaceholderType::Body, Some(1)),
            make_shape_with_placeholder("Shape3", PlaceholderType::Title, None),
        ];

        // Match by type and index
        let idx = find_placeholder_in_shapes(&shapes, Some(&PlaceholderType::Body), Some(1));
        assert_eq!(idx, Some(1));

        // Match by type only
        let idx = find_placeholder_in_shapes(&shapes, Some(&PlaceholderType::Title), None);
        assert_eq!(idx, Some(2));
    }

    #[test]
    fn test_placeholder_matching_title_centered_title() {
        let shapes = vec![make_shape_with_placeholder(
            "Shape1",
            PlaceholderType::CenteredTitle,
            None,
        )];

        // Title should match CenteredTitle
        let idx = find_placeholder_in_shapes(&shapes, Some(&PlaceholderType::Title), None);
        assert_eq!(idx, Some(0));
    }

    #[test]
    fn test_placeholder_matching_special_index() {
        let shapes = vec![
            make_shape_with_placeholder("Shape1", PlaceholderType::Body, Some(1)),
            make_shape_with_placeholder("Shape2", PlaceholderType::Body, Some(2)),
        ];

        // Index 0xFFFFFFFF should match index 1
        let idx =
            find_placeholder_in_shapes(&shapes, Some(&PlaceholderType::Body), Some(0xFFFF_FFFF));
        assert_eq!(idx, Some(0));
    }

    #[test]
    fn test_resolved_transform_merge() {
        let mut t1 = ResolvedTransform {
            x: Some(100),
            y: None,
            width: Some(200),
            height: None,
            rotation: None,
        };

        let t2 = ResolvedTransform {
            x: Some(999), // Should be ignored (t1 has value)
            y: Some(150),
            width: Some(888), // Should be ignored
            height: Some(250),
            rotation: Some(90),
        };

        t1.merge_with(&t2);

        assert_eq!(t1.x, Some(100)); // Original value preserved
        assert_eq!(t1.y, Some(150)); // Filled from t2
        assert_eq!(t1.width, Some(200)); // Original value preserved
        assert_eq!(t1.height, Some(250)); // Filled from t2
        assert_eq!(t1.rotation, Some(90)); // Filled from t2
    }

    #[test]
    fn test_resolved_font_merge() {
        let mut f1 = ResolvedFont {
            size: Some(12.0),
            bold: None,
            italic: Some(true),
            font_family: None,
            color: None,
        };

        let f2 = ResolvedFont {
            size: Some(18.0), // Should be ignored
            bold: Some(true),
            italic: Some(false), // Should be ignored
            font_family: Some("Arial".to_string()),
            color: Some("FF0000".to_string()),
        };

        f1.merge_with(&f2);

        assert_eq!(f1.size, Some(12.0)); // Original preserved
        assert_eq!(f1.bold, Some(true)); // Filled from f2
        assert_eq!(f1.italic, Some(true)); // Original preserved
        assert_eq!(f1.font_family, Some("Arial".to_string())); // Filled from f2
        assert_eq!(f1.color, Some("FF0000".to_string())); // Filled from f2
    }

    #[test]
    fn test_parse_placeholder_from_xml() {
        let xml = r#"<p:ph type="title" idx="0"/>"#;
        let (ptype, pidx) = parse_placeholder_from_xml(xml);
        assert_eq!(ptype, Some(PlaceholderType::Title));
        assert_eq!(pidx, Some(0));

        let xml2 = r#"<p:ph type="body"/>"#;
        let (ptype2, pidx2) = parse_placeholder_from_xml(xml2);
        assert_eq!(ptype2, Some(PlaceholderType::Body));
        assert_eq!(pidx2, None);
    }
}
