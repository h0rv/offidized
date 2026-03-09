use crate::shape::Shape;
use crate::slide::SlideBackground;
use crate::slide_layout::SlideLayout;
use crate::theme::ThemeColorScheme;

/// Mutable slide master with full write API.
#[derive(Debug, Clone, PartialEq)]
pub struct SlideMaster {
    /// Relationship ID referencing this master from the presentation.
    relationship_id: String,
    /// Part URI of the master (e.g., "/ppt/slideMasters/slideMaster1.xml").
    part_uri: String,
    /// Slide layouts associated with this master.
    layouts: Vec<SlideLayout>,
    /// Shapes on this master.
    shapes: Vec<Shape>,
    /// Theme color scheme for this master.
    theme: Option<ThemeColorScheme>,
    /// Background fill for this master.
    background: Option<SlideBackground>,
    /// Preserve flag (whether to preserve the master when unused).
    preserve: bool,
    /// Whether the master has been modified.
    dirty: bool,
}

impl SlideMaster {
    /// Creates a new slide master.
    pub fn new(relationship_id: String, part_uri: String) -> Self {
        Self {
            relationship_id,
            part_uri,
            layouts: Vec::new(),
            shapes: Vec::new(),
            theme: None,
            background: None,
            preserve: false,
            dirty: false,
        }
    }

    /// Creates a slide master from parsed metadata.
    pub(crate) fn from_metadata(
        relationship_id: String,
        part_uri: String,
        layouts: Vec<SlideLayout>,
        shapes: Vec<Shape>,
    ) -> Self {
        Self {
            relationship_id,
            part_uri,
            layouts,
            shapes,
            theme: None,
            background: None,
            preserve: false,
            dirty: false,
        }
    }

    /// Returns the relationship ID.
    pub fn relationship_id(&self) -> &str {
        &self.relationship_id
    }

    /// Returns the part URI.
    pub fn part_uri(&self) -> &str {
        &self.part_uri
    }

    /// Returns a reference to the layouts.
    pub fn layouts(&self) -> &[SlideLayout] {
        &self.layouts
    }

    /// Returns a mutable reference to the layouts.
    pub fn layouts_mut(&mut self) -> &mut Vec<SlideLayout> {
        self.dirty = true;
        &mut self.layouts
    }

    /// Returns a layout by index.
    pub fn layout(&self, index: usize) -> Option<&SlideLayout> {
        self.layouts.get(index)
    }

    /// Returns a mutable layout by index.
    pub fn layout_mut(&mut self, index: usize) -> Option<&mut SlideLayout> {
        self.dirty = true;
        self.layouts.get_mut(index)
    }

    /// Adds a layout to the master.
    pub fn add_layout(&mut self, layout: SlideLayout) {
        self.layouts.push(layout);
        self.dirty = true;
    }

    /// Removes a layout by index.
    pub fn remove_layout(&mut self, index: usize) -> Option<SlideLayout> {
        if index < self.layouts.len() {
            self.dirty = true;
            Some(self.layouts.remove(index))
        } else {
            None
        }
    }

    /// Returns a reference to the shapes.
    pub fn shapes(&self) -> &[Shape] {
        &self.shapes
    }

    /// Returns a mutable reference to the shapes.
    pub fn shapes_mut(&mut self) -> &mut Vec<Shape> {
        self.dirty = true;
        &mut self.shapes
    }

    /// Adds a shape to the master.
    pub fn add_shape(&mut self, shape: Shape) {
        self.shapes.push(shape);
        self.dirty = true;
    }

    /// Removes a shape by index.
    pub fn remove_shape(&mut self, index: usize) -> Option<Shape> {
        if index < self.shapes.len() {
            self.dirty = true;
            Some(self.shapes.remove(index))
        } else {
            None
        }
    }

    /// Returns the theme color scheme.
    pub fn theme(&self) -> Option<&ThemeColorScheme> {
        self.theme.as_ref()
    }

    /// Sets the theme color scheme.
    pub fn set_theme(&mut self, theme: ThemeColorScheme) {
        self.theme = Some(theme);
        self.dirty = true;
    }

    /// Clears the theme color scheme.
    pub fn clear_theme(&mut self) {
        self.theme = None;
        self.dirty = true;
    }

    /// Returns the background fill.
    pub fn background(&self) -> Option<&SlideBackground> {
        self.background.as_ref()
    }

    /// Sets the background fill.
    pub fn set_background(&mut self, background: SlideBackground) {
        self.background = Some(background);
        self.dirty = true;
    }

    /// Clears the background fill.
    pub fn clear_background(&mut self) {
        self.background = None;
        self.dirty = true;
    }

    /// Returns the preserve flag.
    pub fn preserve(&self) -> bool {
        self.preserve
    }

    /// Sets the preserve flag.
    pub fn set_preserve(&mut self, preserve: bool) {
        self.preserve = preserve;
        self.dirty = true;
    }

    /// Applies a theme color scheme to all layouts in this master.
    /// This is a convenience method to propagate theme changes.
    pub fn apply_theme_to_all_layouts(&mut self, theme: ThemeColorScheme) {
        self.set_theme(theme);
        // Note: Individual layout theme application would need to be implemented
        // in the serialization layer, as layouts reference the master's theme.
        self.dirty = true;
    }

    /// Applies a background to all layouts in this master.
    /// This is a convenience method to propagate background changes.
    pub fn apply_background_to_all_layouts(&mut self, background: SlideBackground) {
        self.set_background(background);
        // Note: Individual layout background application would need to be implemented
        // in the serialization layer, as layouts can inherit from the master.
        self.dirty = true;
    }

    /// Returns whether the master has been modified.
    pub fn is_dirty(&self) -> bool {
        self.dirty || self.layouts.iter().any(|l| l.is_dirty())
    }

    /// Marks the master as clean (not modified).
    #[allow(dead_code)]
    pub(crate) fn mark_clean(&mut self) {
        self.dirty = false;
        for layout in &mut self.layouts {
            layout.mark_clean();
        }
    }

    /// Creates a new slide master with a default layout.
    ///
    /// This is a convenience constructor that creates a master with a single
    /// "Title Slide" layout, which is the minimum required for a valid presentation.
    ///
    /// # Arguments
    ///
    /// * `relationship_id` - Relationship ID from the presentation to this master
    /// * `part_uri` - Part URI for this master (e.g., "/ppt/slideMasters/slideMaster1.xml")
    ///
    /// # Example
    ///
    /// ```ignore
    /// let master = SlideMaster::with_default_layout("rId1".to_string(), "/ppt/slideMasters/slideMaster1.xml".to_string());
    /// ```
    pub fn with_default_layout(relationship_id: String, part_uri: String) -> Self {
        let mut master = Self::new(relationship_id, part_uri);

        // Add a default "Title Slide" layout
        let layout = SlideLayout::new(
            "Title Slide",
            "rId1",
            "/ppt/slideLayouts/slideLayout1.xml",
            "rId1",
        );
        master.add_layout(layout);
        master.dirty = true;
        master
    }
}
