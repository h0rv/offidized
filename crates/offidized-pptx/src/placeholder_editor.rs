//! Placeholder customization and content insertion API.
//!
//! Provides functionality for editing placeholder properties, inserting
//! content, and creating standard placeholder shapes.

use crate::error::Result;
use crate::shape::{PlaceholderType, Shape, ShapeGeometry};
use crate::slide::Slide;

/// Builder for customizing placeholder properties.
///
/// Allows fine-grained control over placeholder type, index, position,
/// and size via a chainable API.
pub struct PlaceholderEditor<'a> {
    shape: &'a mut Shape,
}

impl<'a> PlaceholderEditor<'a> {
    /// Create a new placeholder editor for the given shape.
    pub fn new(shape: &'a mut Shape) -> Self {
        Self { shape }
    }

    /// Set the placeholder type.
    ///
    /// Changes the semantic type of the placeholder (Title, Body, Footer, etc.).
    /// This affects how the placeholder inherits properties from layout/master slides.
    pub fn set_type(self, placeholder_type: PlaceholderType) -> Self {
        self.shape.set_placeholder_type(placeholder_type);
        self
    }

    /// Set the placeholder index.
    ///
    /// The index is used to match placeholders across slide, layout, and master levels.
    pub fn set_index(self, idx: u32) -> Self {
        self.shape.set_placeholder_idx(idx);
        self
    }

    /// Clear the placeholder index.
    pub fn clear_index(self) -> Self {
        self.shape.clear_placeholder_idx();
        self
    }

    /// Set the placeholder position and size in EMUs.
    ///
    /// 1 inch = 914400 EMUs.
    pub fn set_geometry(self, x: i64, y: i64, width: i64, height: i64) -> Self {
        self.shape
            .set_geometry(ShapeGeometry::new(x, y, width, height));
        self
    }

    /// Set just the position, preserving existing size.
    ///
    /// If no geometry exists, creates one with zero width/height.
    pub fn set_position(self, x: i64, y: i64) -> Self {
        let (cx, cy) = self
            .shape
            .geometry()
            .map(|g| (g.cx(), g.cy()))
            .unwrap_or((0, 0));
        self.shape.set_geometry(ShapeGeometry::new(x, y, cx, cy));
        self
    }

    /// Set just the size, preserving existing position.
    ///
    /// If no geometry exists, creates one at position (0, 0).
    pub fn set_size(self, width: i64, height: i64) -> Self {
        let (x, y) = self
            .shape
            .geometry()
            .map(|g| (g.x(), g.y()))
            .unwrap_or((0, 0));
        self.shape
            .set_geometry(ShapeGeometry::new(x, y, width, height));
        self
    }

    /// Consume the editor and return the modified shape.
    pub fn apply(self) -> &'a mut Shape {
        self.shape
    }
}

/// Placeholder content insertion utilities.
pub struct PlaceholderContent;

impl PlaceholderContent {
    /// Insert text into a placeholder by adding a paragraph with the given text.
    ///
    /// Adds a new paragraph with a single run containing the text. To replace
    /// existing content, use the shape's paragraph API directly.
    pub fn insert_text(shape: &mut Shape, text: impl Into<String>) {
        let para = shape.add_paragraph();
        para.add_run(text);
    }

    /// Check if a shape is a text-capable placeholder.
    pub fn is_text_placeholder(shape: &Shape) -> bool {
        matches!(
            shape.placeholder_type(),
            Some(PlaceholderType::Title)
                | Some(PlaceholderType::Body)
                | Some(PlaceholderType::CenteredTitle)
                | Some(PlaceholderType::Subtitle)
                | Some(PlaceholderType::Footer)
                | Some(PlaceholderType::Header)
                | Some(PlaceholderType::DateAndTime)
                | Some(PlaceholderType::SlideNumber)
        )
    }

    /// Check if a shape is a media/image placeholder.
    pub fn is_media_placeholder(shape: &Shape) -> bool {
        matches!(
            shape.placeholder_type(),
            Some(PlaceholderType::ClipArt)
                | Some(PlaceholderType::SlideImage)
                | Some(PlaceholderType::Media)
                | Some(PlaceholderType::Object)
        )
    }
}

/// Placeholder inheritance override utilities.
///
/// In OOXML, slide placeholders inherit properties (position, size, formatting)
/// from matching placeholders in the layout and master slides.
pub struct PlaceholderInheritance;

impl PlaceholderInheritance {
    /// Check if a placeholder has an explicit geometry override.
    pub fn has_geometry_override(shape: &Shape) -> bool {
        shape.geometry().is_some()
    }

    /// Override inherited geometry with explicit values.
    pub fn override_geometry(shape: &mut Shape, x: i64, y: i64, width: i64, height: i64) {
        shape.set_geometry(ShapeGeometry::new(x, y, width, height));
    }

    /// Clear the geometry override (reverts to inherited values).
    pub fn clear_geometry_override(shape: &mut Shape) {
        shape.clear_geometry();
    }
}

/// Factory for creating standard placeholder shapes with default positions.
///
/// Position/size values follow typical PowerPoint defaults for a 10"x7.5" slide
/// (9144000 x 6858000 EMUs).
pub struct PlaceholderFactory;

impl PlaceholderFactory {
    /// Create a title placeholder with standard position.
    pub fn create_title(slide: &mut Slide) -> Result<&mut Shape> {
        let shape = slide.add_shape("Title Placeholder");
        shape.set_placeholder_type(PlaceholderType::Title);
        shape.set_placeholder_idx(0);
        shape.set_geometry(ShapeGeometry::new(457200, 274638, 8229600, 1143000));
        Ok(shape)
    }

    /// Create a body/content placeholder with standard position.
    pub fn create_body(slide: &mut Slide) -> Result<&mut Shape> {
        let shape = slide.add_shape("Content Placeholder");
        shape.set_placeholder_type(PlaceholderType::Body);
        shape.set_placeholder_idx(1);
        shape.set_geometry(ShapeGeometry::new(457200, 1600200, 8229600, 4525963));
        Ok(shape)
    }

    /// Create a subtitle placeholder with standard position.
    pub fn create_subtitle(slide: &mut Slide) -> Result<&mut Shape> {
        let shape = slide.add_shape("Subtitle Placeholder");
        shape.set_placeholder_type(PlaceholderType::Subtitle);
        shape.set_placeholder_idx(1);
        shape.set_geometry(ShapeGeometry::new(1371600, 2895600, 6858000, 1325563));
        Ok(shape)
    }

    /// Create a footer placeholder with standard position.
    pub fn create_footer(slide: &mut Slide) -> Result<&mut Shape> {
        let shape = slide.add_shape("Footer Placeholder");
        shape.set_placeholder_type(PlaceholderType::Footer);
        shape.set_placeholder_idx(10);
        shape.set_geometry(ShapeGeometry::new(457200, 6356350, 4038600, 369332));
        Ok(shape)
    }

    /// Create a date/time placeholder with standard position.
    pub fn create_date_time(slide: &mut Slide) -> Result<&mut Shape> {
        let shape = slide.add_shape("Date Placeholder");
        shape.set_placeholder_type(PlaceholderType::DateAndTime);
        shape.set_placeholder_idx(11);
        shape.set_geometry(ShapeGeometry::new(457200, 6356350, 2971800, 369332));
        Ok(shape)
    }

    /// Create a slide number placeholder with standard position.
    pub fn create_slide_number(slide: &mut Slide) -> Result<&mut Shape> {
        let shape = slide.add_shape("Slide Number Placeholder");
        shape.set_placeholder_type(PlaceholderType::SlideNumber);
        shape.set_placeholder_idx(12);
        shape.set_geometry(ShapeGeometry::new(8534400, 6356350, 1524000, 369332));
        Ok(shape)
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn placeholder_editor_set_type_and_index() {
        let mut shape = Shape::new("Test");
        let shape = PlaceholderEditor::new(&mut shape)
            .set_type(PlaceholderType::Title)
            .set_index(0)
            .apply();

        assert_eq!(shape.placeholder_type(), Some(&PlaceholderType::Title));
        assert_eq!(shape.placeholder_idx(), Some(0));
    }

    #[test]
    fn placeholder_editor_set_geometry() {
        let mut shape = Shape::new("Test");
        let shape = PlaceholderEditor::new(&mut shape)
            .set_geometry(100, 200, 300, 400)
            .apply();

        let geo = shape.geometry().expect("geometry should be set");
        assert_eq!(geo.x(), 100);
        assert_eq!(geo.y(), 200);
        assert_eq!(geo.cx(), 300);
        assert_eq!(geo.cy(), 400);
    }

    #[test]
    fn placeholder_editor_set_position_preserves_size() {
        let mut shape = Shape::new("Test");
        shape.set_geometry(ShapeGeometry::new(0, 0, 500, 600));

        let shape = PlaceholderEditor::new(&mut shape)
            .set_position(100, 200)
            .apply();

        let geo = shape.geometry().expect("geometry should be set");
        assert_eq!(geo.x(), 100);
        assert_eq!(geo.y(), 200);
        assert_eq!(geo.cx(), 500);
        assert_eq!(geo.cy(), 600);
    }

    #[test]
    fn placeholder_editor_set_size_preserves_position() {
        let mut shape = Shape::new("Test");
        shape.set_geometry(ShapeGeometry::new(100, 200, 0, 0));

        let shape = PlaceholderEditor::new(&mut shape)
            .set_size(500, 600)
            .apply();

        let geo = shape.geometry().expect("geometry should be set");
        assert_eq!(geo.x(), 100);
        assert_eq!(geo.y(), 200);
        assert_eq!(geo.cx(), 500);
        assert_eq!(geo.cy(), 600);
    }

    #[test]
    fn placeholder_content_is_text_placeholder() {
        let mut shape = Shape::new("Test");
        shape.set_placeholder_type(PlaceholderType::Title);
        assert!(PlaceholderContent::is_text_placeholder(&shape));

        shape.set_placeholder_type(PlaceholderType::Body);
        assert!(PlaceholderContent::is_text_placeholder(&shape));

        shape.clear_placeholder_type();
        assert!(!PlaceholderContent::is_text_placeholder(&shape));
    }

    #[test]
    fn placeholder_content_is_media_placeholder() {
        let mut shape = Shape::new("Test");
        shape.set_placeholder_type(PlaceholderType::Media);
        assert!(PlaceholderContent::is_media_placeholder(&shape));

        shape.set_placeholder_type(PlaceholderType::Title);
        assert!(!PlaceholderContent::is_media_placeholder(&shape));
    }

    #[test]
    fn placeholder_inheritance_geometry_check() {
        let mut shape = Shape::new("Test");
        assert!(!PlaceholderInheritance::has_geometry_override(&shape));

        PlaceholderInheritance::override_geometry(&mut shape, 100, 200, 300, 400);
        assert!(PlaceholderInheritance::has_geometry_override(&shape));

        PlaceholderInheritance::clear_geometry_override(&mut shape);
        assert!(!PlaceholderInheritance::has_geometry_override(&shape));
    }

    #[test]
    fn placeholder_factory_creates_title() {
        let mut slide = Slide::new("Test Slide");
        PlaceholderFactory::create_title(&mut slide).expect("should create title");

        assert_eq!(slide.shapes().len(), 1);
        let shape = &slide.shapes()[0];
        assert_eq!(shape.placeholder_type(), Some(&PlaceholderType::Title));
        assert!(shape.geometry().is_some());
    }

    #[test]
    fn placeholder_factory_creates_body() {
        let mut slide = Slide::new("Test Slide");
        PlaceholderFactory::create_body(&mut slide).expect("should create body");

        let shape = &slide.shapes()[0];
        assert_eq!(shape.placeholder_type(), Some(&PlaceholderType::Body));
    }
}
