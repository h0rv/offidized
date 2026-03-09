use crate::shape::Shape;
use crate::slide::SlideBackground;

/// A mutable slide layout that can be modified and saved.
///
/// Slide layouts define the arrangement and formatting of content placeholders
/// for slides that use this layout. Each layout belongs to a slide master.
#[derive(Debug, Clone, PartialEq)]
pub struct SlideLayout {
    /// Layout name (e.g., "Title Slide", "Title and Content").
    name: String,
    /// Layout type (e.g., "title", "titleOnly", "twoObj").
    layout_type: Option<String>,
    /// Whether this layout should be preserved in the presentation.
    preserve: bool,
    /// Shapes on this layout (placeholders, graphics, etc.).
    shapes: Vec<Shape>,
    /// Background fill for slides using this layout.
    background: Option<SlideBackground>,
    /// Relationship ID to the parent slide master.
    master_relationship_id: String,
    /// Part URI for this layout (e.g., "/ppt/slideLayouts/slideLayout1.xml").
    part_uri: String,
    /// Relationship ID from the presentation to this layout.
    relationship_id: String,
    /// Original XML bytes for roundtrip preservation.
    original_xml: Option<Vec<u8>>,
    /// Whether this layout has been modified.
    dirty: bool,
}

impl SlideLayout {
    /// Creates a new slide layout with the given name.
    ///
    /// # Arguments
    ///
    /// * `name` - The layout name (e.g., "Title Slide")
    /// * `master_relationship_id` - The relationship ID to the parent slide master
    /// * `part_uri` - The part URI for this layout
    /// * `relationship_id` - The relationship ID from the presentation
    pub fn new(
        name: impl Into<String>,
        master_relationship_id: impl Into<String>,
        part_uri: impl Into<String>,
        relationship_id: impl Into<String>,
    ) -> Self {
        Self {
            name: name.into(),
            layout_type: None,
            preserve: false,
            shapes: Vec::new(),
            background: None,
            master_relationship_id: master_relationship_id.into(),
            part_uri: part_uri.into(),
            relationship_id: relationship_id.into(),
            original_xml: None,
            dirty: false,
        }
    }

    /// Returns the layout name.
    pub fn name(&self) -> &str {
        &self.name
    }

    /// Sets the layout name.
    pub fn set_name(&mut self, name: impl Into<String>) {
        self.name = name.into();
        self.dirty = true;
    }

    /// Returns the layout type (e.g., "title", "titleOnly", "twoObj").
    pub fn layout_type(&self) -> Option<&str> {
        self.layout_type.as_deref()
    }

    /// Sets the layout type.
    pub fn set_layout_type(&mut self, layout_type: impl Into<String>) {
        self.layout_type = Some(layout_type.into());
        self.dirty = true;
    }

    /// Clears the layout type.
    pub fn clear_layout_type(&mut self) {
        self.layout_type = None;
        self.dirty = true;
    }

    /// Returns whether this layout should be preserved.
    pub fn preserve(&self) -> bool {
        self.preserve
    }

    /// Sets whether this layout should be preserved.
    pub fn set_preserve(&mut self, preserve: bool) {
        self.preserve = preserve;
        self.dirty = true;
    }

    /// Returns a reference to the shapes on this layout.
    pub fn shapes(&self) -> &[Shape] {
        &self.shapes
    }

    /// Returns a mutable reference to the shapes on this layout.
    pub fn shapes_mut(&mut self) -> &mut Vec<Shape> {
        self.dirty = true;
        &mut self.shapes
    }

    /// Adds a shape to this layout.
    pub fn add_shape(&mut self, shape: Shape) {
        self.shapes.push(shape);
        self.dirty = true;
    }

    /// Removes a shape at the given index.
    ///
    /// Returns `Some(shape)` if a shape was removed, or `None` if the index was out of bounds.
    pub fn remove_shape(&mut self, index: usize) -> Option<Shape> {
        if index < self.shapes.len() {
            self.dirty = true;
            Some(self.shapes.remove(index))
        } else {
            None
        }
    }

    /// Returns the background fill for this layout.
    pub fn background(&self) -> Option<&SlideBackground> {
        self.background.as_ref()
    }

    /// Sets the background fill for this layout.
    pub fn set_background(&mut self, background: SlideBackground) {
        self.background = Some(background);
        self.dirty = true;
    }

    /// Clears the background fill for this layout.
    pub fn clear_background(&mut self) {
        self.background = None;
        self.dirty = true;
    }

    /// Returns the relationship ID to the parent slide master.
    pub fn master_relationship_id(&self) -> &str {
        &self.master_relationship_id
    }

    /// Returns the part URI for this layout.
    pub fn part_uri(&self) -> &str {
        &self.part_uri
    }

    /// Returns the relationship ID from the presentation to this layout.
    pub fn relationship_id(&self) -> &str {
        &self.relationship_id
    }

    /// Returns whether this layout has been modified.
    pub fn is_dirty(&self) -> bool {
        self.dirty
    }

    /// Marks this layout as clean (not modified).
    pub(crate) fn mark_clean(&mut self) {
        self.dirty = false;
    }

    /// Returns the original XML bytes for roundtrip preservation.
    #[allow(dead_code)]
    pub(crate) fn original_xml(&self) -> Option<&[u8]> {
        self.original_xml.as_deref()
    }

    /// Sets the original XML bytes for roundtrip preservation.
    pub(crate) fn set_original_xml(&mut self, xml: Vec<u8>) {
        self.original_xml = Some(xml);
    }

    /// Creates a new layout with a title placeholder.
    ///
    /// This is a convenience constructor for creating a "Title Slide" layout
    /// with a centered title placeholder.
    pub fn with_title_placeholder(
        name: impl Into<String>,
        master_relationship_id: impl Into<String>,
        part_uri: impl Into<String>,
        relationship_id: impl Into<String>,
    ) -> Self {
        use crate::shape::{PlaceholderType, Shape, ShapeGeometry};

        let mut layout = Self::new(name, master_relationship_id, part_uri, relationship_id);
        layout.set_layout_type("title");

        // Add a centered title placeholder with typical dimensions
        let mut title_shape = Shape::new("Title 1");
        title_shape.set_placeholder_type(PlaceholderType::CenteredTitle);
        title_shape.set_geometry(ShapeGeometry::new(
            685800,  // x: left margin
            2130425, // y: upper portion of slide
            7772400, // cx: width
            1470025, // cy: height
        ));
        layout.add_shape(title_shape);

        layout
    }

    /// Creates a new layout with title and content placeholders.
    ///
    /// This is a convenience constructor for creating a "Title and Content" layout.
    pub fn with_title_and_content(
        name: impl Into<String>,
        master_relationship_id: impl Into<String>,
        part_uri: impl Into<String>,
        relationship_id: impl Into<String>,
    ) -> Self {
        use crate::shape::{PlaceholderType, Shape, ShapeGeometry};

        let mut layout = Self::new(name, master_relationship_id, part_uri, relationship_id);
        layout.set_layout_type("obj");

        // Add a title placeholder at the top
        let mut title_shape = Shape::new("Title 1");
        title_shape.set_placeholder_type(PlaceholderType::Title);
        title_shape.set_geometry(ShapeGeometry::new(
            457200,  // x
            274638,  // y: top of slide
            8229600, // cx
            1143000, // cy
        ));
        layout.add_shape(title_shape);

        // Add a content placeholder below
        let mut content_shape = Shape::new("Content Placeholder 2");
        content_shape.set_placeholder_type(PlaceholderType::Object);
        content_shape.set_placeholder_idx(1);
        content_shape.set_geometry(ShapeGeometry::new(
            457200,  // x
            1600200, // y: below title
            8229600, // cx
            4525963, // cy: remaining height
        ));
        layout.add_shape(content_shape);

        layout
    }

    /// Adds a placeholder to this layout and returns its index.
    ///
    /// This is a convenience method for adding a placeholder shape with
    /// the specified type and bounds.
    ///
    /// # Arguments
    ///
    /// * `placeholder_type` - The type of placeholder (Title, Object, etc.)
    /// * `x` - X position in EMUs
    /// * `y` - Y position in EMUs
    /// * `cx` - Width in EMUs
    /// * `cy` - Height in EMUs
    ///
    /// Returns the index of the added shape.
    pub fn add_placeholder(
        &mut self,
        placeholder_type: crate::shape::PlaceholderType,
        x: i64,
        y: i64,
        cx: i64,
        cy: i64,
    ) -> usize {
        use crate::shape::{Shape, ShapeGeometry};

        let shape_count = self.shapes.len() + 1;
        let mut shape = Shape::new(format!("Placeholder {shape_count}"));
        shape.set_placeholder_type(placeholder_type);
        shape.set_geometry(ShapeGeometry::new(x, y, cx, cy));
        self.add_shape(shape);
        self.shapes.len() - 1
    }
}

impl Eq for SlideLayout {}
