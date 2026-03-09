use crate::chart::Chart;
use crate::comment::SlideComment;
use crate::image::Image;
use crate::shape::{GradientFill, PlaceholderType, Shape};
use crate::table::Table;
use crate::text::TextRun;
use crate::timing::SlideTiming;
use crate::transition::SlideTransition;
use offidized_opc::RawXmlNode;

// ── Feature #7: Slide background ──

/// Slide background fill.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum SlideBackground {
    /// Solid fill with sRGB hex color.
    Solid(String),
    /// Gradient fill.
    Gradient(GradientFill),
    /// Pattern background fill.
    Pattern {
        pattern_type: String,
        foreground_color: String,
        background_color: String,
    },
    /// Image background fill.
    Image { relationship_id: String },
}

// ── Feature #15: Footer configuration ──

/// Header/footer configuration parsed from `<p:hf>`.
#[derive(Debug, Clone, PartialEq, Eq, Default)]
pub struct SlideHeaderFooter {
    /// Whether the slide number is shown (`sldNum` attribute, inverted: hf has `sldNum="0"` to hide).
    pub show_slide_number: Option<bool>,
    /// Whether the date/time is shown (`dt` attribute).
    pub show_date_time: Option<bool>,
    /// Whether the header is shown (`hdr` attribute).
    pub show_header: Option<bool>,
    /// Whether the footer is shown (`ftr` attribute).
    pub show_footer: Option<bool>,
    /// Actual footer text content from the footer placeholder shape.
    pub footer_text: Option<String>,
    /// Actual date/time text content from the date placeholder shape.
    pub date_time_text: Option<String>,
}

impl SlideHeaderFooter {
    pub fn is_set(&self) -> bool {
        self.show_slide_number.is_some()
            || self.show_date_time.is_some()
            || self.show_header.is_some()
            || self.show_footer.is_some()
            || self.footer_text.is_some()
            || self.date_time_text.is_some()
    }
}

#[derive(Debug, Clone, PartialEq)]
pub struct Slide {
    title: String,
    shapes: Vec<Shape>,
    text_runs: Vec<TextRun>,
    tables: Vec<Table>,
    images: Vec<Image>,
    charts: Vec<Chart>,
    comments: Vec<SlideComment>,
    transition: Option<SlideTransition>,
    timing: Option<SlideTiming>,
    notes_text: Option<String>,
    /// Slide background (Feature #7).
    background: Option<SlideBackground>,
    /// Whether the slide is hidden (Feature #8).
    hidden: bool,
    /// Header/footer configuration (Feature #15).
    header_footer: Option<SlideHeaderFooter>,
    /// Grouped shapes (Feature #12).
    grouped_shapes: Vec<ShapeGroup>,
    /// Layout reference (master_index, layout_index) for applying layouts to slides.
    layout_reference: Option<(usize, usize)>,
    unknown_children: Vec<RawXmlNode>,
    /// Extra namespace declarations from the original `<p:sld>` element,
    /// preserved so that unknown elements/attributes using those prefixes remain
    /// valid XML on dirty save.
    extra_namespace_declarations: Vec<(String, String)>,
    original_part_bytes: Option<(String, Vec<u8>)>,
    dirty: bool,
}

// ── Feature #12: Grouped shapes ──

/// A group of shapes parsed from `<p:grpSp>`.
#[derive(Debug, Clone, PartialEq)]
pub struct ShapeGroup {
    /// Group name from `p:cNvPr`.
    name: String,
    /// Child shapes within the group.
    shapes: Vec<Shape>,
}

impl ShapeGroup {
    pub fn new(name: impl Into<String>) -> Self {
        Self {
            name: name.into(),
            shapes: Vec::new(),
        }
    }

    pub fn name(&self) -> &str {
        &self.name
    }

    pub fn shapes(&self) -> &[Shape] {
        &self.shapes
    }

    pub fn shapes_mut(&mut self) -> &mut Vec<Shape> {
        &mut self.shapes
    }

    pub fn add_shape(&mut self, shape: Shape) {
        self.shapes.push(shape);
    }

    /// Consume the group and return its shapes.
    pub fn into_shapes(self) -> Vec<Shape> {
        self.shapes
    }

    /// Remove a shape from the group by index.
    pub fn remove_shape(&mut self, index: usize) -> Option<Shape> {
        if index < self.shapes.len() {
            Some(self.shapes.remove(index))
        } else {
            None
        }
    }
}

/// A presentation section parsed from `p14:sectionLst`.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct PresentationSection {
    /// Section name.
    name: String,
    /// Slide IDs belonging to this section.
    slide_ids: Vec<u32>,
}

impl PresentationSection {
    pub fn new(name: impl Into<String>) -> Self {
        Self {
            name: name.into(),
            slide_ids: Vec::new(),
        }
    }

    pub fn name(&self) -> &str {
        &self.name
    }

    pub fn set_name(&mut self, name: impl Into<String>) {
        self.name = name.into();
    }

    pub fn slide_ids(&self) -> &[u32] {
        &self.slide_ids
    }

    pub fn add_slide_id(&mut self, slide_id: u32) {
        self.slide_ids.push(slide_id);
    }

    pub fn set_slide_ids(&mut self, ids: Vec<u32>) {
        self.slide_ids = ids;
    }
}

impl Slide {
    pub fn new(title: impl Into<String>) -> Self {
        Self {
            title: title.into(),
            shapes: Vec::new(),
            text_runs: Vec::new(),
            tables: Vec::new(),
            images: Vec::new(),
            charts: Vec::new(),
            comments: Vec::new(),
            transition: None,
            timing: None,
            notes_text: None,
            background: None,
            hidden: false,
            header_footer: None,
            grouped_shapes: Vec::new(),
            layout_reference: None,
            unknown_children: Vec::new(),
            extra_namespace_declarations: Vec::new(),
            original_part_bytes: None,
            dirty: true,
        }
    }

    pub fn title(&self) -> &str {
        &self.title
    }

    pub fn set_title(&mut self, title: impl Into<String>) {
        self.title = title.into();
        self.mark_dirty();
    }

    pub fn add_shape(&mut self, name: impl Into<String>) -> &mut Shape {
        self.shapes.push(Shape::new(name));
        self.mark_dirty();
        let index = self.shapes.len().saturating_sub(1);
        &mut self.shapes[index]
    }

    pub fn shapes(&self) -> &[Shape] {
        &self.shapes
    }

    pub fn shapes_mut(&mut self) -> &mut [Shape] {
        self.mark_dirty();
        &mut self.shapes
    }

    pub fn shape_count(&self) -> usize {
        self.shapes.len()
    }

    // ── Placeholder collection ──

    /// Returns shapes that have a non-None `placeholder_type`.
    pub fn placeholders(&self) -> Vec<&Shape> {
        self.shapes
            .iter()
            .filter(|s| s.placeholder_type().is_some())
            .collect()
    }

    /// Returns the first shape matching the given placeholder type.
    pub fn placeholder(&self, ph_type: PlaceholderType) -> Option<&Shape> {
        self.shapes
            .iter()
            .find(|s| s.placeholder_type() == Some(&ph_type))
    }

    /// Returns a mutable reference to the first shape matching the given placeholder type.
    pub fn placeholder_mut(&mut self, ph_type: PlaceholderType) -> Option<&mut Shape> {
        self.mark_dirty();
        self.shapes
            .iter_mut()
            .find(|s| s.placeholder_type() == Some(&ph_type))
    }

    /// Convenience: returns the title placeholder shape (either `Title` or `CenteredTitle`).
    pub fn title_placeholder(&self) -> Option<&Shape> {
        self.placeholder(PlaceholderType::Title)
            .or_else(|| self.placeholder(PlaceholderType::CenteredTitle))
    }

    /// Convenience: returns the body placeholder shape.
    pub fn body_placeholder(&self) -> Option<&Shape> {
        self.placeholder(PlaceholderType::Body)
    }

    // ── Shape z-order ──

    /// Move the shape at `index` to the front (last position, renders on top).
    pub fn move_shape_to_front(&mut self, index: usize) {
        if index < self.shapes.len() {
            let shape = self.shapes.remove(index);
            self.shapes.push(shape);
            self.mark_dirty();
        }
    }

    /// Move the shape at `index` to the back (first position, renders on bottom).
    pub fn move_shape_to_back(&mut self, index: usize) {
        if index < self.shapes.len() {
            let shape = self.shapes.remove(index);
            self.shapes.insert(0, shape);
            self.mark_dirty();
        }
    }

    /// Move the shape at `index` one position forward (toward the front).
    pub fn move_shape_forward(&mut self, index: usize) {
        if index < self.shapes.len().saturating_sub(1) {
            self.shapes.swap(index, index + 1);
            self.mark_dirty();
        }
    }

    /// Move the shape at `index` one position backward (toward the back).
    pub fn move_shape_backward(&mut self, index: usize) {
        if index > 0 && index < self.shapes.len() {
            self.shapes.swap(index, index - 1);
            self.mark_dirty();
        }
    }

    pub fn add_text_run(&mut self, text: impl Into<String>) -> &mut TextRun {
        self.text_runs.push(TextRun::new(text));
        self.mark_dirty();
        let index = self.text_runs.len().saturating_sub(1);
        &mut self.text_runs[index]
    }

    pub fn text_runs(&self) -> &[TextRun] {
        &self.text_runs
    }

    pub fn text_runs_mut(&mut self) -> &mut [TextRun] {
        self.mark_dirty();
        &mut self.text_runs
    }

    pub fn text_run_count(&self) -> usize {
        self.text_runs.len()
    }

    pub fn add_table(&mut self, rows: usize, cols: usize) -> &mut Table {
        self.tables.push(Table::new(rows, cols));
        self.mark_dirty();
        let index = self.tables.len().saturating_sub(1);
        &mut self.tables[index]
    }

    pub fn tables(&self) -> &[Table] {
        &self.tables
    }

    pub fn tables_mut(&mut self) -> &mut [Table] {
        self.mark_dirty();
        &mut self.tables
    }

    pub fn table_count(&self) -> usize {
        self.tables.len()
    }

    pub fn add_image(
        &mut self,
        bytes: impl Into<Vec<u8>>,
        content_type: impl Into<String>,
    ) -> &mut Image {
        self.images.push(Image::new(bytes, content_type));
        self.mark_dirty();
        let index = self.images.len().saturating_sub(1);
        &mut self.images[index]
    }

    pub fn images(&self) -> &[Image] {
        &self.images
    }

    pub fn images_mut(&mut self) -> &mut [Image] {
        self.mark_dirty();
        &mut self.images
    }

    pub fn image_count(&self) -> usize {
        self.images.len()
    }

    /// Replace the image at `index` with new binary data and content type.
    ///
    /// Preserves the existing image's name, relationship ID, and crop settings.
    /// Returns `false` if the index is out of bounds.
    pub fn replace_image(
        &mut self,
        index: usize,
        bytes: impl Into<Vec<u8>>,
        content_type: impl Into<String>,
    ) -> bool {
        if index >= self.images.len() {
            return false;
        }
        let new_bytes = bytes.into();
        let new_ct = content_type.into();
        let existing = &mut self.images[index];
        // Create a replacement preserving metadata
        let mut replacement = Image::new(new_bytes, new_ct);
        if let Some(name) = existing.name() {
            replacement.set_name(Some(name.to_string()));
        }
        if let Some(rid) = existing.relationship_id() {
            replacement.set_relationship_id(Some(rid.to_string()));
        }
        if let Some(crop) = existing.crop().cloned() {
            replacement.set_crop(Some(crop));
        }
        self.images[index] = replacement;
        self.mark_dirty();
        true
    }

    pub(crate) fn push_image(&mut self, image: Image) {
        self.images.push(image);
    }

    pub fn add_chart(&mut self, title: impl Into<String>) -> &mut Chart {
        self.charts.push(Chart::new(title));
        self.mark_dirty();
        let index = self.charts.len().saturating_sub(1);
        &mut self.charts[index]
    }

    pub fn charts(&self) -> &[Chart] {
        &self.charts
    }

    pub fn charts_mut(&mut self) -> &mut [Chart] {
        self.mark_dirty();
        &mut self.charts
    }

    pub fn chart_count(&self) -> usize {
        self.charts.len()
    }

    pub fn add_comment(
        &mut self,
        author: impl Into<String>,
        text: impl Into<String>,
    ) -> &mut SlideComment {
        self.comments.push(SlideComment::new(author, text));
        self.mark_dirty();
        let index = self.comments.len().saturating_sub(1);
        &mut self.comments[index]
    }

    pub fn comments(&self) -> &[SlideComment] {
        &self.comments
    }

    pub fn comments_mut(&mut self) -> &mut [SlideComment] {
        self.mark_dirty();
        &mut self.comments
    }

    pub fn comment_count(&self) -> usize {
        self.comments.len()
    }

    pub(crate) fn push_chart(&mut self, chart: Chart) {
        self.charts.push(chart);
    }

    pub(crate) fn push_comment(&mut self, comment: SlideComment) {
        self.comments.push(comment);
    }

    pub fn transition(&self) -> Option<&SlideTransition> {
        self.transition.as_ref()
    }

    pub fn transition_mut(&mut self) -> Option<&mut SlideTransition> {
        self.mark_dirty();
        self.transition.as_mut()
    }

    pub fn set_transition(&mut self, transition: SlideTransition) {
        self.transition = Some(transition);
        self.mark_dirty();
    }

    pub fn clear_transition(&mut self) {
        self.transition = None;
        self.mark_dirty();
    }

    pub fn timing(&self) -> Option<&SlideTiming> {
        self.timing.as_ref()
    }

    pub fn timing_mut(&mut self) -> Option<&mut SlideTiming> {
        self.mark_dirty();
        self.timing.as_mut()
    }

    pub fn set_timing(&mut self, timing: SlideTiming) {
        self.timing = Some(timing);
        self.mark_dirty();
    }

    pub fn clear_timing(&mut self) {
        self.timing = None;
        self.mark_dirty();
    }

    pub fn notes_text(&self) -> Option<&str> {
        self.notes_text.as_deref()
    }

    pub fn set_notes_text(&mut self, notes_text: impl Into<String>) {
        self.notes_text = Some(notes_text.into());
        self.mark_dirty();
    }

    pub fn clear_notes_text(&mut self) {
        self.notes_text = None;
        self.mark_dirty();
    }

    // ── Feature #7: Slide background ──

    /// Slide background fill.
    pub fn background(&self) -> Option<&SlideBackground> {
        self.background.as_ref()
    }

    /// Set slide background.
    pub fn set_background(&mut self, background: SlideBackground) {
        self.background = Some(background);
        self.mark_dirty();
    }

    /// Clear slide background.
    pub fn clear_background(&mut self) {
        self.background = None;
        self.mark_dirty();
    }

    // ── Feature #8: Slide hidden state ──

    /// Whether the slide is hidden (show="0" on p:sld).
    pub fn is_hidden(&self) -> bool {
        self.hidden
    }

    /// Set the hidden flag.
    pub fn set_hidden(&mut self, hidden: bool) {
        self.hidden = hidden;
        self.mark_dirty();
    }

    // ── Feature #12: Grouped shapes ──

    /// Grouped shapes on the slide.
    pub fn grouped_shapes(&self) -> &[ShapeGroup] {
        &self.grouped_shapes
    }

    pub fn grouped_shapes_mut(&mut self) -> &mut Vec<ShapeGroup> {
        self.mark_dirty();
        &mut self.grouped_shapes
    }

    pub(crate) fn push_grouped_shape(&mut self, group: ShapeGroup) {
        self.grouped_shapes.push(group);
    }

    // ── Feature #15: Footer configuration ──

    /// Header/footer configuration.
    pub fn header_footer(&self) -> Option<&SlideHeaderFooter> {
        self.header_footer.as_ref()
    }

    /// Set header/footer configuration.
    pub fn set_header_footer(&mut self, hf: SlideHeaderFooter) {
        self.header_footer = Some(hf);
        self.mark_dirty();
    }

    /// Clear header/footer configuration.
    pub fn clear_header_footer(&mut self) {
        self.header_footer = None;
        self.mark_dirty();
    }

    // ── Layout application ──

    /// Get the layout reference for this slide (master_index, layout_index).
    ///
    /// Returns `None` if no specific layout has been assigned (will use default).
    pub fn layout_reference(&self) -> Option<(usize, usize)> {
        self.layout_reference
    }

    /// Apply a specific layout to this slide.
    ///
    /// # Arguments
    ///
    /// * `master_index` - Index of the slide master (0-based)
    /// * `layout_index` - Index of the layout within that master (0-based)
    ///
    /// # Example
    ///
    /// ```ignore
    /// slide.apply_layout(0, 2); // Use layout #2 from master #0
    /// ```
    pub fn apply_layout(&mut self, master_index: usize, layout_index: usize) {
        self.layout_reference = Some((master_index, layout_index));
        self.mark_dirty();
    }

    /// Clear the layout reference (slide will use default layout).
    pub fn clear_layout(&mut self) {
        self.layout_reference = None;
        self.mark_dirty();
    }

    /// Unknown XML children preserved for roundtrip fidelity.
    pub fn unknown_children(&self) -> &[RawXmlNode] {
        &self.unknown_children
    }

    pub(crate) fn set_unknown_children(&mut self, children: Vec<RawXmlNode>) {
        self.unknown_children = children;
    }

    pub(crate) fn original_part_bytes(&self) -> Option<(&str, &[u8])> {
        self.original_part_bytes
            .as_ref()
            .map(|(part_uri, bytes)| (part_uri.as_str(), bytes.as_slice()))
    }

    pub(crate) fn set_original_part_bytes(&mut self, part_uri: String, bytes: Vec<u8>) {
        self.original_part_bytes = Some((part_uri, bytes));
        self.dirty = false;
    }

    pub(crate) fn extra_namespace_declarations(&self) -> &[(String, String)] {
        &self.extra_namespace_declarations
    }

    pub(crate) fn set_extra_namespace_declarations(&mut self, declarations: Vec<(String, String)>) {
        self.extra_namespace_declarations = declarations;
    }

    pub(crate) fn dirty(&self) -> bool {
        self.dirty
    }

    // ── Shape removal, duplication, and find ──

    /// Remove a shape by index. Returns the removed shape if the index is valid.
    pub fn remove_shape(&mut self, index: usize) -> Option<Shape> {
        if index < self.shapes.len() {
            self.mark_dirty();
            Some(self.shapes.remove(index))
        } else {
            None
        }
    }

    /// Remove the first shape with the given name. Returns the removed shape if found.
    pub fn remove_shape_by_name(&mut self, name: &str) -> Option<Shape> {
        let pos = self.shapes.iter().position(|s| s.name() == name)?;
        self.mark_dirty();
        Some(self.shapes.remove(pos))
    }

    /// Duplicate the shape at `index`, appending the clone to the end.
    ///
    /// Returns a mutable reference to the new shape, or `None` if the index is out of bounds.
    pub fn duplicate_shape(&mut self, index: usize) -> Option<&mut Shape> {
        if index >= self.shapes.len() {
            return None;
        }
        let cloned = self.shapes[index].clone();
        self.shapes.push(cloned);
        self.mark_dirty();
        self.shapes.last_mut()
    }

    /// Find a shape by name.
    pub fn find_shape_by_name(&self, name: &str) -> Option<&Shape> {
        self.shapes.iter().find(|s| s.name() == name)
    }

    /// Find a shape by name (mutable).
    pub fn find_shape_by_name_mut(&mut self, name: &str) -> Option<&mut Shape> {
        self.mark_dirty();
        self.shapes.iter_mut().find(|s| s.name() == name)
    }

    // ── Find and replace text ──

    /// Search all shapes' text for occurrences of `needle`.
    ///
    /// Returns tuples of `(shape_index, paragraph_index, run_index)` for each match.
    pub fn find_text(&self, needle: &str) -> Vec<(usize, usize, usize)> {
        let mut results = Vec::new();
        for (si, shape) in self.shapes.iter().enumerate() {
            for (pi, para) in shape.paragraphs().iter().enumerate() {
                for (ri, run) in para.runs().iter().enumerate() {
                    if run.text().contains(needle) {
                        results.push((si, pi, ri));
                    }
                }
            }
        }
        results
    }

    /// Replace all occurrences of `old` with `new` across all shapes' text.
    ///
    /// Returns the number of replacements made.
    pub fn replace_text(&mut self, old: &str, new: &str) -> usize {
        let mut count = 0;
        for shape in &mut self.shapes {
            for para in shape.paragraphs_mut() {
                for run in para.runs_mut() {
                    let text = run.text().to_string();
                    if text.contains(old) {
                        let replaced = text.replace(old, new);
                        count += text.matches(old).count();
                        run.set_text(replaced);
                    }
                }
            }
        }
        if count > 0 {
            self.mark_dirty();
        }
        count
    }

    // ── Group shape operations ──

    /// Group the shapes at the given indices into a new group.
    ///
    /// The shapes are removed from the top-level list and placed into a new [`ShapeGroup`].
    /// Indices are processed in reverse order to avoid shifting. Returns `true` if the group was created.
    pub fn group_shapes(&mut self, indices: &[usize], name: impl Into<String>) -> bool {
        let max_index = self.shapes.len();
        if indices.len() < 2 || indices.iter().any(|&i| i >= max_index) {
            return false;
        }
        let mut sorted_indices: Vec<usize> = indices.to_vec();
        sorted_indices.sort_unstable();
        sorted_indices.dedup();
        if sorted_indices.len() < 2 {
            return false;
        }

        let mut group = ShapeGroup::new(name);
        // Remove in reverse order to keep indices valid
        for &idx in sorted_indices.iter().rev() {
            let shape = self.shapes.remove(idx);
            group.shapes_mut().insert(0, shape);
        }
        self.grouped_shapes.push(group);
        self.mark_dirty();
        true
    }

    /// Dissolve a grouped shape, returning its children to the top-level shape list.
    ///
    /// Returns `true` if the group was found and ungrouped.
    pub fn ungroup_shapes(&mut self, group_index: usize) -> bool {
        if group_index >= self.grouped_shapes.len() {
            return false;
        }
        let group = self.grouped_shapes.remove(group_index);
        for shape in group.into_shapes() {
            self.shapes.push(shape);
        }
        self.mark_dirty();
        true
    }

    fn mark_dirty(&mut self) {
        self.dirty = true;
    }
}

#[cfg(test)]
mod tests {
    use super::{PresentationSection, Slide, SlideBackground, SlideHeaderFooter};
    use crate::timing::SlideTiming;
    use crate::transition::{SlideTransition, SlideTransitionKind};

    #[test]
    fn owns_descriptor_collections() {
        let mut slide = Slide::new("Title");
        slide.add_shape("Headline box");
        slide.add_text_run("Bullet one");
        slide.add_table(2, 3);
        slide.add_image(vec![0_u8, 1, 2], "image/png");
        slide.add_chart("Revenue");
        slide.add_comment("Alice", "Review this slide");

        assert_eq!(slide.shape_count(), 1);
        assert_eq!(slide.text_run_count(), 1);
        assert_eq!(slide.table_count(), 1);
        assert_eq!(slide.image_count(), 1);
        assert_eq!(slide.chart_count(), 1);
        assert_eq!(slide.comment_count(), 1);
        assert_eq!(slide.shapes()[0].name(), "Headline box");
        assert_eq!(slide.text_runs()[0].text(), "Bullet one");
        assert_eq!(slide.tables()[0].rows(), 2);
        assert_eq!(slide.tables()[0].cols(), 3);
        assert_eq!(slide.images()[0].content_type(), "image/png");
        assert_eq!(slide.charts()[0].title(), "Revenue");
        assert_eq!(slide.comments()[0].author(), "Alice");
        assert_eq!(slide.comments()[0].text(), "Review this slide");
    }

    #[test]
    fn supports_transition_timing_and_notes_metadata() {
        let mut slide = Slide::new("Title");
        let mut transition = SlideTransition::new(SlideTransitionKind::Wipe);
        transition.set_advance_on_click(Some(false));
        transition.set_advance_after_ms(Some(3_000));
        let timing = SlideTiming::new("<p:tnLst/>");

        slide.set_transition(transition.clone());
        slide.set_timing(timing.clone());
        slide.set_notes_text("Presenter notes");

        assert_eq!(slide.transition(), Some(&transition));
        assert_eq!(slide.timing(), Some(&timing));
        assert_eq!(slide.notes_text(), Some("Presenter notes"));

        slide.clear_transition();
        slide.clear_timing();
        slide.clear_notes_text();
        assert!(slide.transition().is_none());
        assert!(slide.timing().is_none());
        assert!(slide.notes_text().is_none());
    }

    // ── Shape z-order tests ──

    #[test]
    fn move_shape_to_front() {
        let mut slide = Slide::new("Title");
        slide.add_shape("A");
        slide.add_shape("B");
        slide.add_shape("C");

        slide.move_shape_to_front(0);
        assert_eq!(slide.shapes()[0].name(), "B");
        assert_eq!(slide.shapes()[1].name(), "C");
        assert_eq!(slide.shapes()[2].name(), "A");
    }

    #[test]
    fn move_shape_to_back() {
        let mut slide = Slide::new("Title");
        slide.add_shape("A");
        slide.add_shape("B");
        slide.add_shape("C");

        slide.move_shape_to_back(2);
        assert_eq!(slide.shapes()[0].name(), "C");
        assert_eq!(slide.shapes()[1].name(), "A");
        assert_eq!(slide.shapes()[2].name(), "B");
    }

    #[test]
    fn move_shape_forward() {
        let mut slide = Slide::new("Title");
        slide.add_shape("A");
        slide.add_shape("B");
        slide.add_shape("C");

        slide.move_shape_forward(0);
        assert_eq!(slide.shapes()[0].name(), "B");
        assert_eq!(slide.shapes()[1].name(), "A");
        assert_eq!(slide.shapes()[2].name(), "C");
    }

    #[test]
    fn move_shape_backward() {
        let mut slide = Slide::new("Title");
        slide.add_shape("A");
        slide.add_shape("B");
        slide.add_shape("C");

        slide.move_shape_backward(2);
        assert_eq!(slide.shapes()[0].name(), "A");
        assert_eq!(slide.shapes()[1].name(), "C");
        assert_eq!(slide.shapes()[2].name(), "B");
    }

    #[test]
    fn move_shape_noop_on_boundary() {
        let mut slide = Slide::new("Title");
        slide.add_shape("A");
        slide.add_shape("B");

        // Move forward at last position is a no-op.
        slide.move_shape_forward(1);
        assert_eq!(slide.shapes()[0].name(), "A");
        assert_eq!(slide.shapes()[1].name(), "B");

        // Move backward at first position is a no-op.
        slide.move_shape_backward(0);
        assert_eq!(slide.shapes()[0].name(), "A");
        assert_eq!(slide.shapes()[1].name(), "B");
    }

    // ── Slide background pattern and image tests ──

    #[test]
    fn slide_background_pattern() {
        let mut slide = Slide::new("Title");
        slide.set_background(SlideBackground::Pattern {
            pattern_type: "pct20".to_string(),
            foreground_color: "FF0000".to_string(),
            background_color: "FFFFFF".to_string(),
        });

        match slide.background() {
            Some(SlideBackground::Pattern {
                pattern_type,
                foreground_color,
                background_color,
            }) => {
                assert_eq!(pattern_type, "pct20");
                assert_eq!(foreground_color, "FF0000");
                assert_eq!(background_color, "FFFFFF");
            }
            _ => panic!("expected Pattern background"),
        }
    }

    #[test]
    fn slide_background_image() {
        let mut slide = Slide::new("Title");
        slide.set_background(SlideBackground::Image {
            relationship_id: "rId2".to_string(),
        });

        match slide.background() {
            Some(SlideBackground::Image { relationship_id }) => {
                assert_eq!(relationship_id, "rId2");
            }
            _ => panic!("expected Image background"),
        }
    }

    // ── Footer content tests ──

    #[test]
    fn footer_content_text() {
        let mut slide = Slide::new("Title");
        let hf = SlideHeaderFooter {
            show_footer: Some(true),
            footer_text: Some("Confidential".to_string()),
            date_time_text: Some("2024-01-15".to_string()),
            ..Default::default()
        };
        slide.set_header_footer(hf);

        let hf = slide.header_footer().unwrap();
        assert_eq!(hf.footer_text.as_deref(), Some("Confidential"));
        assert_eq!(hf.date_time_text.as_deref(), Some("2024-01-15"));
        assert!(hf.is_set());
    }

    #[test]
    fn footer_content_only_text_is_set() {
        let mut hf = SlideHeaderFooter::default();
        assert!(!hf.is_set());
        hf.footer_text = Some("Footer".to_string());
        assert!(hf.is_set());
    }

    // ── Presentation section tests ──

    #[test]
    fn presentation_section_roundtrip() {
        let mut section = PresentationSection::new("Introduction");
        assert_eq!(section.name(), "Introduction");
        assert!(section.slide_ids().is_empty());

        section.add_slide_id(256);
        section.add_slide_id(257);

        assert_eq!(section.slide_ids(), &[256, 257]);

        section.set_name("Overview");
        assert_eq!(section.name(), "Overview");

        section.set_slide_ids(vec![100, 200, 300]);
        assert_eq!(section.slide_ids(), &[100, 200, 300]);
    }

    #[test]
    fn presentation_section_empty() {
        let section = PresentationSection::new("Empty");
        assert_eq!(section.name(), "Empty");
        assert!(section.slide_ids().is_empty());
    }

    // ── Placeholder collection tests ──

    #[test]
    fn placeholders_returns_shapes_with_placeholder_type() {
        let mut slide = Slide::new("Title");
        slide.add_shape("Title 1");
        slide.add_shape("Body 1");
        slide.add_shape("Decoration");

        // Set placeholder types on first two shapes.
        slide.shapes_mut()[0].set_placeholder_type(crate::shape::PlaceholderType::Title);
        slide.shapes_mut()[1].set_placeholder_type(crate::shape::PlaceholderType::Body);

        let placeholders = slide.placeholders();
        assert_eq!(placeholders.len(), 2);
        assert_eq!(placeholders[0].name(), "Title 1");
        assert_eq!(placeholders[1].name(), "Body 1");
    }

    #[test]
    fn placeholders_empty_when_none_set() {
        let mut slide = Slide::new("Title");
        slide.add_shape("A");
        slide.add_shape("B");

        assert!(slide.placeholders().is_empty());
    }

    #[test]
    fn placeholder_finds_by_type() {
        let mut slide = Slide::new("Title");
        slide.add_shape("Title 1");
        slide.add_shape("Body 1");
        slide.shapes_mut()[0].set_placeholder_type(crate::shape::PlaceholderType::Title);
        slide.shapes_mut()[1].set_placeholder_type(crate::shape::PlaceholderType::Body);

        let title = slide.placeholder(crate::shape::PlaceholderType::Title);
        assert!(title.is_some());
        assert_eq!(title.unwrap().name(), "Title 1");

        let body = slide.placeholder(crate::shape::PlaceholderType::Body);
        assert!(body.is_some());
        assert_eq!(body.unwrap().name(), "Body 1");

        let footer = slide.placeholder(crate::shape::PlaceholderType::Footer);
        assert!(footer.is_none());
    }

    #[test]
    fn placeholder_mut_finds_and_modifies() {
        let mut slide = Slide::new("Title");
        slide.add_shape("Title 1");
        slide.shapes_mut()[0].set_placeholder_type(crate::shape::PlaceholderType::Title);

        let title = slide.placeholder_mut(crate::shape::PlaceholderType::Title);
        assert!(title.is_some());
        let title_shape = title.unwrap();
        title_shape.add_paragraph_with_text("New title text");
        assert_eq!(title_shape.paragraph_count(), 1);
    }

    #[test]
    fn title_placeholder_finds_title_or_centered_title() {
        let mut slide = Slide::new("Title");
        slide.add_shape("CTitle");
        slide.shapes_mut()[0].set_placeholder_type(crate::shape::PlaceholderType::CenteredTitle);

        // Should find CenteredTitle via fallback.
        let title = slide.title_placeholder();
        assert!(title.is_some());
        assert_eq!(title.unwrap().name(), "CTitle");

        // Add a regular title; it should take priority.
        slide.add_shape("Title 1");
        let idx = slide.shapes().len() - 1;
        slide.shapes_mut()[idx].set_placeholder_type(crate::shape::PlaceholderType::Title);

        let title = slide.title_placeholder();
        assert!(title.is_some());
        assert_eq!(title.unwrap().name(), "Title 1");
    }

    #[test]
    fn body_placeholder_convenience() {
        let mut slide = Slide::new("Title");
        slide.add_shape("Body 1");
        slide.shapes_mut()[0].set_placeholder_type(crate::shape::PlaceholderType::Body);

        let body = slide.body_placeholder();
        assert!(body.is_some());
        assert_eq!(body.unwrap().name(), "Body 1");

        // No body placeholder set.
        let mut slide2 = Slide::new("Title");
        slide2.add_shape("Decoration");
        assert!(slide2.body_placeholder().is_none());
    }

    // ── Phase 2A: Shape removal and duplication tests ──

    #[test]
    fn remove_shape_by_index() {
        let mut slide = Slide::new("Title");
        slide.add_shape("A");
        slide.add_shape("B");
        slide.add_shape("C");

        let removed = slide.remove_shape(1);
        assert!(removed.is_some());
        assert_eq!(removed.unwrap().name(), "B");
        assert_eq!(slide.shape_count(), 2);
        assert_eq!(slide.shapes()[0].name(), "A");
        assert_eq!(slide.shapes()[1].name(), "C");
    }

    #[test]
    fn remove_shape_out_of_bounds() {
        let mut slide = Slide::new("Title");
        slide.add_shape("A");
        assert!(slide.remove_shape(5).is_none());
        assert_eq!(slide.shape_count(), 1);
    }

    #[test]
    fn remove_shape_by_name_found() {
        let mut slide = Slide::new("Title");
        slide.add_shape("Alpha");
        slide.add_shape("Beta");
        slide.add_shape("Gamma");

        let removed = slide.remove_shape_by_name("Beta");
        assert!(removed.is_some());
        assert_eq!(removed.unwrap().name(), "Beta");
        assert_eq!(slide.shape_count(), 2);
    }

    #[test]
    fn remove_shape_by_name_not_found() {
        let mut slide = Slide::new("Title");
        slide.add_shape("Alpha");
        assert!(slide.remove_shape_by_name("Missing").is_none());
        assert_eq!(slide.shape_count(), 1);
    }

    #[test]
    fn duplicate_shape_clones_at_end() {
        let mut slide = Slide::new("Title");
        slide.add_shape("Original");
        slide.add_shape("Other");

        let dup = slide.duplicate_shape(0);
        assert!(dup.is_some());
        assert_eq!(slide.shape_count(), 3);
        assert_eq!(slide.shapes()[2].name(), "Original");
    }

    #[test]
    fn duplicate_shape_out_of_bounds() {
        let mut slide = Slide::new("Title");
        slide.add_shape("A");
        assert!(slide.duplicate_shape(10).is_none());
        assert_eq!(slide.shape_count(), 1);
    }

    #[test]
    fn find_shape_by_name_returns_ref() {
        let mut slide = Slide::new("Title");
        slide.add_shape("Target");
        slide.add_shape("Other");

        assert!(slide.find_shape_by_name("Target").is_some());
        assert_eq!(slide.find_shape_by_name("Target").unwrap().name(), "Target");
        assert!(slide.find_shape_by_name("Missing").is_none());
    }

    #[test]
    fn find_shape_by_name_mut_allows_modification() {
        let mut slide = Slide::new("Title");
        slide.add_shape("Target");

        let shape = slide.find_shape_by_name_mut("Target").unwrap();
        shape.add_paragraph_with_text("Hello");
        assert_eq!(slide.shapes()[0].paragraphs().len(), 1);
    }

    // ── Phase 2B: Find and replace text tests ──

    #[test]
    fn find_text_in_shapes() {
        let mut slide = Slide::new("Title");
        slide.add_shape("Shape1");
        slide.shapes_mut()[0].add_paragraph_with_text("Hello World");
        slide.add_shape("Shape2");
        slide.shapes_mut()[1].add_paragraph_with_text("Goodbye World");

        let results = slide.find_text("World");
        assert_eq!(results.len(), 2);
        assert_eq!(results[0].0, 0); // shape 0
        assert_eq!(results[1].0, 1); // shape 1
    }

    #[test]
    fn find_text_no_match() {
        let mut slide = Slide::new("Title");
        slide.add_shape("Shape1");
        slide.shapes_mut()[0].add_paragraph_with_text("Hello");

        let results = slide.find_text("Missing");
        assert!(results.is_empty());
    }

    #[test]
    fn replace_text_across_shapes() {
        let mut slide = Slide::new("Title");
        slide.add_shape("Shape1");
        slide.shapes_mut()[0].add_paragraph_with_text("Hello World");
        slide.add_shape("Shape2");
        slide.shapes_mut()[1].add_paragraph_with_text("World Peace");

        let count = slide.replace_text("World", "Earth");
        assert_eq!(count, 2);
        assert_eq!(
            slide.shapes()[0].paragraphs()[0].runs()[0].text(),
            "Hello Earth"
        );
        assert_eq!(
            slide.shapes()[1].paragraphs()[0].runs()[0].text(),
            "Earth Peace"
        );
    }

    #[test]
    fn replace_text_no_match_returns_zero() {
        let mut slide = Slide::new("Title");
        slide.add_shape("Shape1");
        slide.shapes_mut()[0].add_paragraph_with_text("Hello");

        let count = slide.replace_text("Missing", "Found");
        assert_eq!(count, 0);
    }

    // ── Phase 2C: Group shape operations tests ──

    #[test]
    fn group_shapes_creates_group() {
        let mut slide = Slide::new("Title");
        slide.add_shape("A");
        slide.add_shape("B");
        slide.add_shape("C");

        let result = slide.group_shapes(&[0, 2], "MyGroup");
        assert!(result);
        assert_eq!(slide.shape_count(), 1); // B remains
        assert_eq!(slide.shapes()[0].name(), "B");
        assert_eq!(slide.grouped_shapes().len(), 1);
        assert_eq!(slide.grouped_shapes()[0].name(), "MyGroup");
        assert_eq!(slide.grouped_shapes()[0].shapes().len(), 2);
    }

    #[test]
    fn group_shapes_requires_at_least_two() {
        let mut slide = Slide::new("Title");
        slide.add_shape("A");
        slide.add_shape("B");

        assert!(!slide.group_shapes(&[0], "G"));
        assert!(!slide.group_shapes(&[], "G"));
        assert_eq!(slide.shape_count(), 2);
    }

    #[test]
    fn group_shapes_rejects_out_of_bounds() {
        let mut slide = Slide::new("Title");
        slide.add_shape("A");
        slide.add_shape("B");

        assert!(!slide.group_shapes(&[0, 5], "G"));
        assert_eq!(slide.shape_count(), 2);
    }

    #[test]
    fn ungroup_shapes_returns_children() {
        let mut slide = Slide::new("Title");
        slide.add_shape("A");
        slide.add_shape("B");
        slide.add_shape("C");
        slide.group_shapes(&[0, 1], "Group1");

        // After grouping: shapes = [C], groups = [Group1(A, B)]
        assert_eq!(slide.shape_count(), 1);
        assert_eq!(slide.grouped_shapes().len(), 1);

        let result = slide.ungroup_shapes(0);
        assert!(result);
        assert_eq!(slide.shape_count(), 3); // C + A + B
        assert_eq!(slide.grouped_shapes().len(), 0);
    }

    #[test]
    fn ungroup_shapes_out_of_bounds() {
        let mut slide = Slide::new("Title");
        assert!(!slide.ungroup_shapes(0));
    }
}
