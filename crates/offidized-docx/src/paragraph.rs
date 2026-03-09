use crate::image::{FloatingImage, InlineImage};
use crate::run::Run;
use offidized_opc::RawXmlNode;

/// Line spacing rule for a paragraph (`w:lineRule` attribute on `w:spacing`).
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum LineSpacingRule {
    /// Spacing is automatically determined by the consumer (default). The value
    /// in `line_spacing_twips` is interpreted as 240ths of a line.
    Auto,
    /// Spacing is an exact value in twips. Lines taller than the value are clipped.
    Exact,
    /// Spacing is at least the given value in twips. Lines taller than the value
    /// expand the spacing.
    AtLeast,
}

impl LineSpacingRule {
    pub(crate) fn from_xml_value(value: &str) -> Option<Self> {
        match value {
            "auto" => Some(Self::Auto),
            "exact" => Some(Self::Exact),
            "atLeast" => Some(Self::AtLeast),
            _ => None,
        }
    }

    pub(crate) fn to_xml_value(self) -> &'static str {
        match self {
            Self::Auto => "auto",
            Self::Exact => "exact",
            Self::AtLeast => "atLeast",
        }
    }
}

/// Paragraph alignment.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ParagraphAlignment {
    Left,
    Center,
    Right,
    Justified,
}

impl ParagraphAlignment {
    pub(crate) fn from_xml_value(value: &str) -> Option<Self> {
        match value {
            "left" | "start" => Some(Self::Left),
            "center" => Some(Self::Center),
            "right" | "end" => Some(Self::Right),
            "both" | "justify" => Some(Self::Justified),
            _ => None,
        }
    }

    pub(crate) fn to_xml_value(self) -> &'static str {
        match self {
            Self::Left => "left",
            Self::Center => "center",
            Self::Right => "right",
            Self::Justified => "both",
        }
    }
}

/// Tab stop alignment within a paragraph.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum TabStopAlignment {
    /// Left-aligned tab stop.
    Left,
    /// Center-aligned tab stop.
    Center,
    /// Right-aligned tab stop.
    Right,
    /// Decimal-aligned tab stop.
    Decimal,
    /// Bar tab stop.
    Bar,
    /// Clear an inherited tab stop at this position.
    Clear,
}

impl TabStopAlignment {
    pub(crate) fn from_xml_value(value: &str) -> Option<Self> {
        match value {
            "left" | "start" => Some(Self::Left),
            "center" => Some(Self::Center),
            "right" | "end" => Some(Self::Right),
            "decimal" | "num" => Some(Self::Decimal),
            "bar" => Some(Self::Bar),
            "clear" => Some(Self::Clear),
            _ => None,
        }
    }

    pub(crate) fn to_xml_value(self) -> &'static str {
        match self {
            Self::Left => "left",
            Self::Center => "center",
            Self::Right => "right",
            Self::Decimal => "decimal",
            Self::Bar => "bar",
            Self::Clear => "clear",
        }
    }
}

/// Tab stop leader character.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum TabStopLeader {
    /// No leader.
    None,
    /// Dot leader.
    Dot,
    /// Hyphen leader.
    Hyphen,
    /// Underscore leader.
    Underscore,
    /// Heavy leader (thick line).
    Heavy,
    /// Middle dot leader.
    MiddleDot,
}

impl TabStopLeader {
    pub(crate) fn from_xml_value(value: &str) -> Option<Self> {
        match value {
            "none" => Some(Self::None),
            "dot" => Some(Self::Dot),
            "hyphen" => Some(Self::Hyphen),
            "underscore" => Some(Self::Underscore),
            "heavy" => Some(Self::Heavy),
            "middleDot" => Some(Self::MiddleDot),
            _ => None,
        }
    }

    pub(crate) fn to_xml_value(self) -> &'static str {
        match self {
            Self::None => "none",
            Self::Dot => "dot",
            Self::Hyphen => "hyphen",
            Self::Underscore => "underscore",
            Self::Heavy => "heavy",
            Self::MiddleDot => "middleDot",
        }
    }
}

/// A single tab stop definition in a paragraph.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct TabStop {
    position_twips: u32,
    alignment: TabStopAlignment,
    leader: Option<TabStopLeader>,
    /// Whether this is a numbering tab (`w:tab` with `w:val="num"`-like behaviour flag).
    num_tab: bool,
}

impl TabStop {
    /// Create a new tab stop at the given position with alignment.
    pub fn new(position_twips: u32, alignment: TabStopAlignment) -> Self {
        Self {
            position_twips,
            alignment,
            leader: None,
            num_tab: false,
        }
    }

    /// Create a new tab stop with a leader character.
    pub fn with_leader(
        position_twips: u32,
        alignment: TabStopAlignment,
        leader: TabStopLeader,
    ) -> Self {
        Self {
            position_twips,
            alignment,
            leader: Some(leader),
            num_tab: false,
        }
    }

    /// Tab stop position, in twips.
    pub fn position_twips(&self) -> u32 {
        self.position_twips
    }

    /// Set tab stop position, in twips.
    pub fn set_position_twips(&mut self, position: u32) {
        self.position_twips = position;
    }

    /// Tab stop alignment.
    pub fn alignment(&self) -> TabStopAlignment {
        self.alignment
    }

    /// Set tab stop alignment.
    pub fn set_alignment(&mut self, alignment: TabStopAlignment) {
        self.alignment = alignment;
    }

    /// Tab stop leader character.
    pub fn leader(&self) -> Option<TabStopLeader> {
        self.leader
    }

    /// Set tab stop leader character.
    pub fn set_leader(&mut self, leader: TabStopLeader) {
        self.leader = Some(leader);
    }

    /// Clear leader character.
    pub fn clear_leader(&mut self) {
        self.leader = None;
    }

    /// Whether this tab stop is a numbering tab.
    pub fn num_tab(&self) -> bool {
        self.num_tab
    }

    /// Set whether this tab stop is a numbering tab.
    pub fn set_num_tab(&mut self, num_tab: bool) {
        self.num_tab = num_tab;
    }
}

/// A single border edge definition for a paragraph border.
#[derive(Debug, Clone, PartialEq, Eq, Default)]
pub struct ParagraphBorder {
    line_type: Option<String>,
    size_eighth_points: Option<u16>,
    color: Option<String>,
    space_points: Option<u32>,
}

impl ParagraphBorder {
    /// Create a border with the given line type.
    pub fn new(line_type: impl Into<String>) -> Self {
        let mut border = Self::default();
        border.set_line_type(line_type);
        border
    }

    /// Border line type (e.g., `"single"`, `"double"`, `"dotted"`).
    pub fn line_type(&self) -> Option<&str> {
        self.line_type.as_deref()
    }

    /// Set border line type.
    pub fn set_line_type(&mut self, line_type: impl Into<String>) {
        self.line_type = normalize_optional_text(line_type.into());
    }

    /// Clear border line type.
    pub fn clear_line_type(&mut self) {
        self.line_type = None;
    }

    /// Border size in eighth-points.
    pub fn size_eighth_points(&self) -> Option<u16> {
        self.size_eighth_points
    }

    /// Set border size in eighth-points.
    pub fn set_size_eighth_points(&mut self, size: u16) {
        self.size_eighth_points = Some(size);
    }

    /// Clear border size.
    pub fn clear_size_eighth_points(&mut self) {
        self.size_eighth_points = None;
    }

    /// Border color, uppercase hex without `#`.
    pub fn color(&self) -> Option<&str> {
        self.color.as_deref()
    }

    /// Set border color.
    pub fn set_color(&mut self, color: impl Into<String>) {
        self.color = normalize_color_value(color.into().as_str());
    }

    /// Clear border color.
    pub fn clear_color(&mut self) {
        self.color = None;
    }

    /// Space between border and text, in points.
    pub fn space_points(&self) -> Option<u32> {
        self.space_points
    }

    /// Set space between border and text, in points.
    pub fn set_space_points(&mut self, space: u32) {
        self.space_points = Some(space);
    }

    /// Clear space.
    pub fn clear_space_points(&mut self) {
        self.space_points = None;
    }

    pub(crate) fn set_line_type_option(&mut self, line_type: Option<String>) {
        self.line_type = line_type.and_then(normalize_optional_text);
    }
}

/// Paragraph border collection (`w:pBdr`).
#[derive(Debug, Clone, PartialEq, Eq, Default)]
pub struct ParagraphBorders {
    top: Option<ParagraphBorder>,
    left: Option<ParagraphBorder>,
    bottom: Option<ParagraphBorder>,
    right: Option<ParagraphBorder>,
    between: Option<ParagraphBorder>,
}

impl ParagraphBorders {
    /// Create an empty border set.
    pub fn new() -> Self {
        Self::default()
    }

    /// Top border.
    pub fn top(&self) -> Option<&ParagraphBorder> {
        self.top.as_ref()
    }

    /// Set top border.
    pub fn set_top(&mut self, border: ParagraphBorder) {
        self.top = Some(border);
    }

    /// Clear top border.
    pub fn clear_top(&mut self) {
        self.top = None;
    }

    /// Left border.
    pub fn left(&self) -> Option<&ParagraphBorder> {
        self.left.as_ref()
    }

    /// Set left border.
    pub fn set_left(&mut self, border: ParagraphBorder) {
        self.left = Some(border);
    }

    /// Clear left border.
    pub fn clear_left(&mut self) {
        self.left = None;
    }

    /// Bottom border.
    pub fn bottom(&self) -> Option<&ParagraphBorder> {
        self.bottom.as_ref()
    }

    /// Set bottom border.
    pub fn set_bottom(&mut self, border: ParagraphBorder) {
        self.bottom = Some(border);
    }

    /// Clear bottom border.
    pub fn clear_bottom(&mut self) {
        self.bottom = None;
    }

    /// Right border.
    pub fn right(&self) -> Option<&ParagraphBorder> {
        self.right.as_ref()
    }

    /// Set right border.
    pub fn set_right(&mut self, border: ParagraphBorder) {
        self.right = Some(border);
    }

    /// Clear right border.
    pub fn clear_right(&mut self) {
        self.right = None;
    }

    /// Between border (between consecutive paragraphs with same settings).
    pub fn between(&self) -> Option<&ParagraphBorder> {
        self.between.as_ref()
    }

    /// Set between border.
    pub fn set_between(&mut self, border: ParagraphBorder) {
        self.between = Some(border);
    }

    /// Clear between border.
    pub fn clear_between(&mut self) {
        self.between = None;
    }

    /// Clear all borders.
    pub fn clear(&mut self) {
        *self = Self::default();
    }

    pub(crate) fn is_empty(&self) -> bool {
        self.top.is_none()
            && self.left.is_none()
            && self.bottom.is_none()
            && self.right.is_none()
            && self.between.is_none()
    }
}

/// A paragraph in a Word document.
#[derive(Debug, Clone, PartialEq, Eq, Default)]
pub struct Paragraph {
    runs: Vec<Run>,
    heading_level: Option<u8>,
    style_id: Option<String>,
    alignment: Option<ParagraphAlignment>,
    spacing_before_twips: Option<u32>,
    spacing_after_twips: Option<u32>,
    line_spacing_twips: Option<u32>,
    indent_left_twips: Option<u32>,
    indent_right_twips: Option<u32>,
    indent_first_line_twips: Option<u32>,
    indent_hanging_twips: Option<u32>,
    numbering_num_id: Option<u32>,
    numbering_ilvl: Option<u8>,
    tab_stops: Vec<TabStop>,
    borders: ParagraphBorders,
    shading_color: Option<String>,
    shading_pattern: Option<String>,
    shading_color_attribute: Option<String>,
    line_spacing_rule: Option<LineSpacingRule>,
    before_autospacing: Option<bool>,
    after_autospacing: Option<bool>,
    keep_next: bool,
    keep_lines: bool,
    /// Whether a page break is inserted before this paragraph (`w:pageBreakBefore`).
    page_break_before: bool,
    /// Whether contextual spacing is enabled (`w:contextualSpacing`).
    /// When true, spacing between paragraphs of the same style is suppressed.
    contextual_spacing: bool,
    widow_control: Option<bool>,
    /// Outline level for this paragraph (`w:outlineLvl`), values 0-9.
    /// Used to define heading levels for table-of-contents generation.
    outline_level: Option<u8>,
    /// Bidirectional (RTL) paragraph direction (`w:bidi`).
    bidi: bool,
    /// Comment range start ids associated with this paragraph.
    comment_range_start_ids: Vec<u32>,
    /// Comment range end ids associated with this paragraph.
    comment_range_end_ids: Vec<u32>,
    /// Section properties attached to this paragraph (for section breaks).
    /// When present, this paragraph is the last paragraph of a non-final section.
    section_properties: Option<Box<crate::section::Section>>,
    unknown_children: Vec<RawXmlNode>,
    unknown_property_children: Vec<RawXmlNode>,
}

impl Paragraph {
    /// Create an empty paragraph.
    pub fn new() -> Self {
        Self::default()
    }

    /// Create a paragraph initialized with one text run.
    pub fn from_text(text: impl Into<String>) -> Self {
        let mut paragraph = Self::new();
        paragraph.add_run(text);
        paragraph
    }

    /// Create a heading paragraph.
    pub fn heading(text: impl Into<String>, level: u8) -> Self {
        let mut paragraph = Self::from_text(text);
        paragraph.heading_level = Some(level.clamp(1, 9));
        paragraph
    }

    /// Add a run and return a mutable reference to it.
    pub fn add_run(&mut self, text: impl Into<String>) -> &mut Run {
        self.runs.push(Run::new(text));
        let idx = self.runs.len().saturating_sub(1);
        &mut self.runs[idx]
    }

    /// Add a run with an explicit character style identifier.
    pub fn add_run_with_style(
        &mut self,
        text: impl Into<String>,
        style_id: impl Into<String>,
    ) -> &mut Run {
        let run = self.add_run(text);
        run.set_style_id(style_id);
        run
    }

    /// Add a hyperlink run and return a mutable reference to it.
    pub fn add_hyperlink(
        &mut self,
        text: impl Into<String>,
        hyperlink: impl Into<String>,
    ) -> &mut Run {
        let run = self.add_run(text);
        run.set_hyperlink(hyperlink);
        run
    }

    /// Add an inline image run and return a mutable reference to it.
    pub fn add_inline_image(
        &mut self,
        image_index: usize,
        width_emu: u32,
        height_emu: u32,
    ) -> &mut Run {
        let run = self.add_run("");
        run.set_inline_image(InlineImage::new(image_index, width_emu, height_emu));
        run
    }

    /// Add a floating image run and return a mutable reference to it.
    pub fn add_floating_image(
        &mut self,
        image_index: usize,
        width_emu: u32,
        height_emu: u32,
    ) -> &mut Run {
        let run = self.add_run("");
        run.set_floating_image(FloatingImage::new(image_index, width_emu, height_emu));
        run
    }

    /// All runs in this paragraph.
    pub fn runs(&self) -> &[Run] {
        &self.runs
    }

    /// Mutable access to all runs.
    pub fn runs_mut(&mut self) -> &mut [Run] {
        &mut self.runs
    }

    /// Replace all runs with a single run containing the given text.
    pub fn set_text(&mut self, text: impl Into<String>) {
        self.runs.clear();
        self.runs.push(Run::new(text));
    }

    /// Replace all runs with the given vector of runs.
    ///
    /// Any existing runs (and their unknown children) are dropped.
    pub fn set_runs(&mut self, runs: Vec<Run>) {
        self.runs = runs;
    }

    /// Heading level when this paragraph is a heading.
    pub fn heading_level(&self) -> Option<u8> {
        self.heading_level
    }

    pub(crate) fn set_heading_level(&mut self, heading_level: Option<u8>) {
        self.heading_level = heading_level.map(|level| level.clamp(1, 9));
    }

    /// Paragraph style identifier (`w:pStyle`).
    pub fn style_id(&self) -> Option<&str> {
        self.style_id.as_deref()
    }

    /// Set paragraph style identifier (`w:pStyle`).
    pub fn set_style_id(&mut self, style_id: impl Into<String>) {
        let style_id = style_id.into();
        let trimmed = style_id.trim();
        self.style_id = if trimmed.is_empty() {
            None
        } else {
            Some(trimmed.to_string())
        };
    }

    /// Clear paragraph style identifier.
    pub fn clear_style_id(&mut self) {
        self.style_id = None;
    }

    pub(crate) fn set_style_id_option(&mut self, style_id: Option<String>) {
        self.style_id = style_id.and_then(|value| {
            let trimmed = value.trim();
            if trimmed.is_empty() {
                None
            } else {
                Some(trimmed.to_string())
            }
        });
    }

    /// Paragraph alignment.
    pub fn alignment(&self) -> Option<ParagraphAlignment> {
        self.alignment
    }

    /// Set paragraph alignment.
    pub fn set_alignment(&mut self, alignment: ParagraphAlignment) {
        self.alignment = Some(alignment);
    }

    /// Clear paragraph alignment.
    pub fn clear_alignment(&mut self) {
        self.alignment = None;
    }

    /// Spacing before paragraph, in twips.
    pub fn spacing_before_twips(&self) -> Option<u32> {
        self.spacing_before_twips
    }

    /// Set spacing before paragraph, in twips.
    pub fn set_spacing_before_twips(&mut self, value: u32) {
        self.spacing_before_twips = Some(value);
    }

    /// Clear spacing before paragraph.
    pub fn clear_spacing_before_twips(&mut self) {
        self.spacing_before_twips = None;
    }

    /// Spacing after paragraph, in twips.
    pub fn spacing_after_twips(&self) -> Option<u32> {
        self.spacing_after_twips
    }

    /// Set spacing after paragraph, in twips.
    pub fn set_spacing_after_twips(&mut self, value: u32) {
        self.spacing_after_twips = Some(value);
    }

    /// Clear spacing after paragraph.
    pub fn clear_spacing_after_twips(&mut self) {
        self.spacing_after_twips = None;
    }

    /// Line spacing for this paragraph, in twips.
    pub fn line_spacing_twips(&self) -> Option<u32> {
        self.line_spacing_twips
    }

    /// Set line spacing, in twips.
    pub fn set_line_spacing_twips(&mut self, value: u32) {
        self.line_spacing_twips = Some(value);
    }

    /// Clear line spacing.
    pub fn clear_line_spacing_twips(&mut self) {
        self.line_spacing_twips = None;
    }

    /// Line spacing rule (`w:lineRule` attribute on `w:spacing`).
    /// Controls how the `line_spacing_twips` value is interpreted.
    pub fn line_spacing_rule(&self) -> Option<LineSpacingRule> {
        self.line_spacing_rule
    }

    /// Set line spacing rule.
    pub fn set_line_spacing_rule(&mut self, rule: LineSpacingRule) {
        self.line_spacing_rule = Some(rule);
    }

    /// Clear line spacing rule (use default).
    pub fn clear_line_spacing_rule(&mut self) {
        self.line_spacing_rule = None;
    }

    /// Whether automatic spacing before paragraph is enabled (`w:beforeAutospacing`).
    pub fn before_autospacing(&self) -> Option<bool> {
        self.before_autospacing
    }

    /// Set automatic spacing before paragraph.
    pub fn set_before_autospacing(&mut self, value: bool) {
        self.before_autospacing = Some(value);
    }

    /// Clear automatic spacing before paragraph.
    pub fn clear_before_autospacing(&mut self) {
        self.before_autospacing = None;
    }

    /// Whether automatic spacing after paragraph is enabled (`w:afterAutospacing`).
    pub fn after_autospacing(&self) -> Option<bool> {
        self.after_autospacing
    }

    /// Set automatic spacing after paragraph.
    pub fn set_after_autospacing(&mut self, value: bool) {
        self.after_autospacing = Some(value);
    }

    /// Clear automatic spacing after paragraph.
    pub fn clear_after_autospacing(&mut self) {
        self.after_autospacing = None;
    }

    /// Left indentation, in twips.
    pub fn indent_left_twips(&self) -> Option<u32> {
        self.indent_left_twips
    }

    /// Set left indentation, in twips.
    pub fn set_indent_left_twips(&mut self, value: u32) {
        self.indent_left_twips = Some(value);
    }

    /// Clear left indentation.
    pub fn clear_indent_left_twips(&mut self) {
        self.indent_left_twips = None;
    }

    /// Right indentation, in twips.
    pub fn indent_right_twips(&self) -> Option<u32> {
        self.indent_right_twips
    }

    /// Set right indentation, in twips.
    pub fn set_indent_right_twips(&mut self, value: u32) {
        self.indent_right_twips = Some(value);
    }

    /// Clear right indentation.
    pub fn clear_indent_right_twips(&mut self) {
        self.indent_right_twips = None;
    }

    /// First-line indentation, in twips.
    pub fn indent_first_line_twips(&self) -> Option<u32> {
        self.indent_first_line_twips
    }

    /// Set first-line indentation, in twips.
    pub fn set_indent_first_line_twips(&mut self, value: u32) {
        self.indent_first_line_twips = Some(value);
    }

    /// Clear first-line indentation.
    pub fn clear_indent_first_line_twips(&mut self) {
        self.indent_first_line_twips = None;
    }

    /// Hanging indentation, in twips.
    pub fn indent_hanging_twips(&self) -> Option<u32> {
        self.indent_hanging_twips
    }

    /// Set hanging indentation, in twips.
    pub fn set_indent_hanging_twips(&mut self, value: u32) {
        self.indent_hanging_twips = Some(value);
    }

    /// Clear hanging indentation.
    pub fn clear_indent_hanging_twips(&mut self) {
        self.indent_hanging_twips = None;
    }

    /// Numbering definition id for this paragraph (`w:numId`).
    pub fn numbering_num_id(&self) -> Option<u32> {
        self.numbering_num_id
    }

    /// Numbering indentation level for this paragraph (`w:ilvl`).
    pub fn numbering_ilvl(&self) -> Option<u8> {
        self.numbering_ilvl
    }

    /// Set paragraph list numbering (`w:numPr`).
    pub fn set_numbering(&mut self, num_id: u32, ilvl: u8) {
        self.numbering_num_id = Some(num_id);
        self.numbering_ilvl = Some(ilvl);
    }

    /// Remove paragraph list numbering (`w:numPr`).
    pub fn clear_numbering(&mut self) {
        self.numbering_num_id = None;
        self.numbering_ilvl = None;
    }

    /// Set the bullet/numbering indentation level without changing the num ID.
    ///
    /// This is a convenience for adjusting the nesting level of an already-numbered paragraph.
    pub fn set_bullet_level(&mut self, level: u32) {
        self.numbering_ilvl = Some(level as u8);
    }

    pub(crate) fn set_numbering_num_id(&mut self, num_id: Option<u32>) {
        self.numbering_num_id = num_id;
    }

    pub(crate) fn set_numbering_ilvl(&mut self, ilvl: Option<u8>) {
        self.numbering_ilvl = ilvl;
    }

    /// Tab stops for this paragraph (`w:tabs`).
    pub fn tab_stops(&self) -> &[TabStop] {
        &self.tab_stops
    }

    /// Add a tab stop.
    pub fn add_tab_stop(&mut self, tab_stop: TabStop) {
        self.tab_stops.push(tab_stop);
    }

    /// Replace all tab stops.
    pub fn set_tab_stops(&mut self, tab_stops: Vec<TabStop>) {
        self.tab_stops = tab_stops;
    }

    /// Clear all tab stops.
    pub fn clear_tab_stops(&mut self) {
        self.tab_stops.clear();
    }

    /// Paragraph borders (`w:pBdr`).
    pub fn borders(&self) -> &ParagraphBorders {
        &self.borders
    }

    /// Mutable paragraph borders.
    pub fn borders_mut(&mut self) -> &mut ParagraphBorders {
        &mut self.borders
    }

    /// Set paragraph borders.
    pub fn set_borders(&mut self, borders: ParagraphBorders) {
        self.borders = borders;
    }

    /// Clear all paragraph borders.
    pub fn clear_borders(&mut self) {
        self.borders.clear();
    }

    /// Paragraph shading fill color (`w:shd w:fill`), uppercase hex without `#`.
    pub fn shading_color(&self) -> Option<&str> {
        self.shading_color.as_deref()
    }

    /// Set paragraph shading fill color.
    pub fn set_shading_color(&mut self, color: impl Into<String>) {
        self.shading_color = normalize_color_value(color.into().as_str());
    }

    /// Clear paragraph shading fill color.
    pub fn clear_shading_color(&mut self) {
        self.shading_color = None;
    }

    /// Paragraph shading pattern (`w:shd w:val`), e.g. `"clear"`.
    pub fn shading_pattern(&self) -> Option<&str> {
        self.shading_pattern.as_deref()
    }

    /// Set paragraph shading pattern.
    pub fn set_shading_pattern(&mut self, pattern: impl Into<String>) {
        self.shading_pattern = normalize_optional_text(pattern.into());
    }

    /// Clear paragraph shading pattern.
    pub fn clear_shading_pattern(&mut self) {
        self.shading_pattern = None;
    }

    pub(crate) fn set_shading_color_option(&mut self, color: Option<String>) {
        self.shading_color = color;
    }

    pub(crate) fn set_shading_pattern_option(&mut self, pattern: Option<String>) {
        self.shading_pattern = pattern;
    }

    /// The `w:color` attribute on paragraph shading (`w:shd`).
    ///
    /// This is the foreground/pattern color, not the fill. Returns `None` when the
    /// attribute was not present in the original XML.
    pub fn shading_color_attribute(&self) -> Option<&str> {
        self.shading_color_attribute.as_deref()
    }

    /// Set the `w:color` attribute for paragraph shading.
    pub fn set_shading_color_attribute(&mut self, color: impl Into<String>) {
        let color = color.into();
        self.shading_color_attribute = if color.trim().is_empty() {
            None
        } else {
            Some(color)
        };
    }

    /// Clear the shading `w:color` attribute (defaults to `"auto"` on serialize).
    pub fn clear_shading_color_attribute(&mut self) {
        self.shading_color_attribute = None;
    }

    pub(crate) fn set_shading_color_attribute_option(&mut self, color: Option<String>) {
        self.shading_color_attribute = color;
    }

    /// Whether this paragraph should keep with the next paragraph (`w:keepNext`).
    pub fn keep_next(&self) -> bool {
        self.keep_next
    }

    /// Set keep with next paragraph.
    pub fn set_keep_next(&mut self, keep_next: bool) {
        self.keep_next = keep_next;
    }

    /// Whether all lines of this paragraph should stay on the same page (`w:keepLines`).
    pub fn keep_lines(&self) -> bool {
        self.keep_lines
    }

    /// Set keep lines together.
    pub fn set_keep_lines(&mut self, keep_lines: bool) {
        self.keep_lines = keep_lines;
    }

    /// Whether a page break is inserted before this paragraph (`w:pageBreakBefore`).
    pub fn page_break_before(&self) -> bool {
        self.page_break_before
    }

    /// Set page break before paragraph.
    pub fn set_page_break_before(&mut self, page_break_before: bool) {
        self.page_break_before = page_break_before;
    }

    /// Whether contextual spacing is enabled for this paragraph (`w:contextualSpacing`).
    /// When true, spacing between paragraphs of the same style is suppressed.
    pub fn contextual_spacing(&self) -> bool {
        self.contextual_spacing
    }

    /// Set contextual spacing.
    pub fn set_contextual_spacing(&mut self, contextual_spacing: bool) {
        self.contextual_spacing = contextual_spacing;
    }

    /// Outline level for this paragraph (`w:outlineLvl`), values 0-9.
    /// Used to define heading levels for table-of-contents generation.
    pub fn outline_level(&self) -> Option<u8> {
        self.outline_level
    }

    /// Set outline level (clamped to 0-9).
    pub fn set_outline_level(&mut self, level: u8) {
        self.outline_level = Some(level.clamp(0, 9));
    }

    /// Clear outline level.
    pub fn clear_outline_level(&mut self) {
        self.outline_level = None;
    }

    /// Widow/orphan control for this paragraph (`w:widowControl`).
    /// `None` means use default (typically enabled), `Some(true)` enables, `Some(false)` disables.
    pub fn widow_control(&self) -> Option<bool> {
        self.widow_control
    }

    /// Set widow/orphan control.
    pub fn set_widow_control(&mut self, widow_control: bool) {
        self.widow_control = Some(widow_control);
    }

    /// Clear explicit widow/orphan control (use default).
    pub fn clear_widow_control(&mut self) {
        self.widow_control = None;
    }

    /// Whether this paragraph has bidirectional (RTL) direction (`w:bidi`).
    pub fn is_bidi(&self) -> bool {
        self.bidi
    }

    /// Set bidirectional (RTL) paragraph direction.
    pub fn set_bidi(&mut self, bidi: bool) {
        self.bidi = bidi;
    }

    /// Comment range start ids associated with this paragraph.
    pub fn comment_range_start_ids(&self) -> &[u32] {
        &self.comment_range_start_ids
    }

    /// Add a comment range start id.
    pub fn add_comment_range_start(&mut self, comment_id: u32) {
        self.comment_range_start_ids.push(comment_id);
    }

    /// Clear all comment range start ids.
    pub fn clear_comment_range_starts(&mut self) {
        self.comment_range_start_ids.clear();
    }

    /// Comment range end ids associated with this paragraph.
    pub fn comment_range_end_ids(&self) -> &[u32] {
        &self.comment_range_end_ids
    }

    /// Add a comment range end id.
    pub fn add_comment_range_end(&mut self, comment_id: u32) {
        self.comment_range_end_ids.push(comment_id);
    }

    /// Clear all comment range end ids.
    pub fn clear_comment_range_ends(&mut self) {
        self.comment_range_end_ids.clear();
    }

    /// Section properties attached to this paragraph (for section breaks).
    /// When present, this paragraph is the last paragraph of a non-final section.
    pub fn section_properties(&self) -> Option<&crate::section::Section> {
        self.section_properties.as_deref()
    }

    /// Set section properties for this paragraph (creating a section break).
    pub fn set_section_properties(&mut self, section: crate::section::Section) {
        self.section_properties = Some(Box::new(section));
    }

    /// Clear section properties from this paragraph.
    pub fn clear_section_properties(&mut self) {
        self.section_properties = None;
    }

    #[allow(dead_code)]
    pub(crate) fn set_section_properties_option(
        &mut self,
        section: Option<crate::section::Section>,
    ) {
        self.section_properties = section.map(Box::new);
    }

    pub(crate) fn has_properties(&self) -> bool {
        self.heading_level.is_some()
            || self.style_id.is_some()
            || self.alignment.is_some()
            || self.spacing_before_twips.is_some()
            || self.spacing_after_twips.is_some()
            || self.line_spacing_twips.is_some()
            || self.indent_left_twips.is_some()
            || self.indent_right_twips.is_some()
            || self.indent_first_line_twips.is_some()
            || self.indent_hanging_twips.is_some()
            || self.numbering_num_id.is_some()
            || self.numbering_ilvl.is_some()
            || !self.tab_stops.is_empty()
            || !self.borders.is_empty()
            || self.shading_color.is_some()
            || self.line_spacing_rule.is_some()
            || self.before_autospacing.is_some()
            || self.after_autospacing.is_some()
            || self.keep_next
            || self.keep_lines
            || self.page_break_before
            || self.contextual_spacing
            || self.widow_control.is_some()
            || self.outline_level.is_some()
            || self.bidi
            || self.section_properties.is_some()
            || !self.unknown_property_children.is_empty()
    }

    pub(crate) fn unknown_children(&self) -> &[RawXmlNode] {
        self.unknown_children.as_slice()
    }

    pub(crate) fn push_unknown_child(&mut self, node: RawXmlNode) {
        self.unknown_children.push(node);
    }

    pub(crate) fn unknown_property_children(&self) -> &[RawXmlNode] {
        self.unknown_property_children.as_slice()
    }

    pub(crate) fn push_unknown_property_child(&mut self, node: RawXmlNode) {
        self.unknown_property_children.push(node);
    }

    /// Concatenated plain text for this paragraph.
    pub fn text(&self) -> String {
        let mut text = String::new();
        for run in &self.runs {
            text.push_str(run.text());
        }
        text
    }
}

fn normalize_optional_text(value: String) -> Option<String> {
    let trimmed = value.trim();
    if trimmed.is_empty() {
        None
    } else {
        Some(trimmed.to_string())
    }
}

fn normalize_color_value(value: &str) -> Option<String> {
    let trimmed = value.trim();
    let normalized = trimmed.trim_start_matches('#').to_ascii_uppercase();
    if normalized.is_empty() || normalized == "AUTO" {
        None
    } else {
        Some(normalized)
    }
}

#[cfg(test)]
mod tests {
    use super::{
        LineSpacingRule, Paragraph, ParagraphAlignment, ParagraphBorder, ParagraphBorders, TabStop,
        TabStopAlignment, TabStopLeader,
    };
    use crate::image::{FloatingImage, InlineImage};

    #[test]
    fn creates_paragraph_from_text() {
        let paragraph = Paragraph::from_text("Hello world");

        assert_eq!(paragraph.text(), "Hello world");
        assert_eq!(paragraph.runs().len(), 1);
        assert_eq!(paragraph.runs()[0].text(), "Hello world");
        assert_eq!(paragraph.heading_level(), None);
        assert_eq!(paragraph.style_id(), None);
        assert_eq!(paragraph.alignment(), None);
        assert_eq!(paragraph.spacing_before_twips(), None);
        assert_eq!(paragraph.spacing_after_twips(), None);
        assert_eq!(paragraph.line_spacing_twips(), None);
        assert_eq!(paragraph.indent_left_twips(), None);
        assert_eq!(paragraph.indent_right_twips(), None);
        assert_eq!(paragraph.indent_first_line_twips(), None);
        assert_eq!(paragraph.indent_hanging_twips(), None);
        assert_eq!(paragraph.numbering_num_id(), None);
        assert_eq!(paragraph.numbering_ilvl(), None);
    }

    #[test]
    fn paragraph_formatting_and_hyperlinks_can_be_set() {
        let mut paragraph = Paragraph::new();
        paragraph.set_style_id("BodyText");
        paragraph.set_alignment(ParagraphAlignment::Justified);
        paragraph.set_spacing_before_twips(120);
        paragraph.set_spacing_after_twips(80);
        paragraph.set_line_spacing_twips(300);
        paragraph.set_indent_left_twips(240);
        paragraph.set_indent_right_twips(60);
        paragraph.set_indent_first_line_twips(180);
        paragraph.set_indent_hanging_twips(90);
        paragraph.set_numbering(42, 2);

        paragraph.add_hyperlink("offidized", "https://example.com");
        paragraph.add_inline_image(0, 640_000, 480_000);
        paragraph.add_floating_image(1, 720_000, 540_000);

        assert_eq!(paragraph.style_id(), Some("BodyText"));
        assert_eq!(paragraph.alignment(), Some(ParagraphAlignment::Justified));
        assert_eq!(paragraph.spacing_before_twips(), Some(120));
        assert_eq!(paragraph.spacing_after_twips(), Some(80));
        assert_eq!(paragraph.line_spacing_twips(), Some(300));
        assert_eq!(paragraph.indent_left_twips(), Some(240));
        assert_eq!(paragraph.indent_right_twips(), Some(60));
        assert_eq!(paragraph.indent_first_line_twips(), Some(180));
        assert_eq!(paragraph.indent_hanging_twips(), Some(90));
        assert_eq!(paragraph.numbering_num_id(), Some(42));
        assert_eq!(paragraph.numbering_ilvl(), Some(2));
        assert_eq!(paragraph.runs()[0].hyperlink(), Some("https://example.com"));
        assert_eq!(
            paragraph.runs()[1]
                .inline_image()
                .map(InlineImage::image_index),
            Some(0)
        );
        assert_eq!(
            paragraph.runs()[2]
                .floating_image()
                .map(FloatingImage::image_index),
            Some(1)
        );
    }

    #[test]
    fn add_run_with_style_sets_character_style() {
        let mut paragraph = Paragraph::new();

        paragraph.add_run_with_style("Styled", "Emphasis");

        assert_eq!(paragraph.runs().len(), 1);
        assert_eq!(paragraph.runs()[0].text(), "Styled");
        assert_eq!(paragraph.runs()[0].style_id(), Some("Emphasis"));
    }

    #[test]
    fn paragraph_style_id_can_be_cleared() {
        let mut paragraph = Paragraph::new();
        paragraph.set_style_id("Quote");
        assert_eq!(paragraph.style_id(), Some("Quote"));
        paragraph.clear_style_id();
        assert_eq!(paragraph.style_id(), None);
    }

    #[test]
    fn tab_stops_can_be_added_and_cleared() {
        let mut paragraph = Paragraph::new();
        assert!(paragraph.tab_stops().is_empty());

        paragraph.add_tab_stop(TabStop::new(720, TabStopAlignment::Left));
        paragraph.add_tab_stop(TabStop::with_leader(
            4320,
            TabStopAlignment::Right,
            TabStopLeader::Dot,
        ));

        assert_eq!(paragraph.tab_stops().len(), 2);
        assert_eq!(paragraph.tab_stops()[0].position_twips(), 720);
        assert_eq!(paragraph.tab_stops()[0].alignment(), TabStopAlignment::Left);
        assert_eq!(paragraph.tab_stops()[0].leader(), None);
        assert_eq!(paragraph.tab_stops()[1].position_twips(), 4320);
        assert_eq!(
            paragraph.tab_stops()[1].alignment(),
            TabStopAlignment::Right
        );
        assert_eq!(paragraph.tab_stops()[1].leader(), Some(TabStopLeader::Dot));

        paragraph.clear_tab_stops();
        assert!(paragraph.tab_stops().is_empty());
    }

    #[test]
    fn paragraph_borders_can_be_set_and_cleared() {
        let mut paragraph = Paragraph::new();
        assert!(paragraph.borders().is_empty());

        let mut top = ParagraphBorder::new("single");
        top.set_size_eighth_points(8);
        top.set_color("#FF0000");
        top.set_space_points(1);
        paragraph.borders_mut().set_top(top);

        assert_eq!(
            paragraph
                .borders()
                .top()
                .and_then(ParagraphBorder::line_type),
            Some("single")
        );
        assert_eq!(
            paragraph
                .borders()
                .top()
                .and_then(ParagraphBorder::size_eighth_points),
            Some(8)
        );
        assert_eq!(
            paragraph.borders().top().and_then(ParagraphBorder::color),
            Some("FF0000")
        );
        assert_eq!(
            paragraph
                .borders()
                .top()
                .and_then(ParagraphBorder::space_points),
            Some(1)
        );

        paragraph.clear_borders();
        assert!(paragraph.borders().is_empty());
    }

    #[test]
    fn paragraph_shading_can_be_set_and_cleared() {
        let mut paragraph = Paragraph::new();
        assert_eq!(paragraph.shading_color(), None);

        paragraph.set_shading_color("#FFAA00");
        assert_eq!(paragraph.shading_color(), Some("FFAA00"));

        paragraph.set_shading_pattern("clear");
        assert_eq!(paragraph.shading_pattern(), Some("clear"));

        paragraph.clear_shading_color();
        paragraph.clear_shading_pattern();
        assert_eq!(paragraph.shading_color(), None);
        assert_eq!(paragraph.shading_pattern(), None);
    }

    #[test]
    fn keep_next_and_keep_lines_and_widow_control() {
        let mut paragraph = Paragraph::new();
        assert!(!paragraph.keep_next());
        assert!(!paragraph.keep_lines());
        assert_eq!(paragraph.widow_control(), None);

        paragraph.set_keep_next(true);
        paragraph.set_keep_lines(true);
        paragraph.set_widow_control(false);

        assert!(paragraph.keep_next());
        assert!(paragraph.keep_lines());
        assert_eq!(paragraph.widow_control(), Some(false));

        paragraph.set_keep_next(false);
        paragraph.set_keep_lines(false);
        paragraph.clear_widow_control();

        assert!(!paragraph.keep_next());
        assert!(!paragraph.keep_lines());
        assert_eq!(paragraph.widow_control(), None);
    }

    #[test]
    fn tab_stop_alignment_xml_roundtrip() {
        assert_eq!(
            TabStopAlignment::from_xml_value("left"),
            Some(TabStopAlignment::Left)
        );
        assert_eq!(
            TabStopAlignment::from_xml_value("center"),
            Some(TabStopAlignment::Center)
        );
        assert_eq!(
            TabStopAlignment::from_xml_value("right"),
            Some(TabStopAlignment::Right)
        );
        assert_eq!(
            TabStopAlignment::from_xml_value("decimal"),
            Some(TabStopAlignment::Decimal)
        );
        assert_eq!(
            TabStopAlignment::from_xml_value("bar"),
            Some(TabStopAlignment::Bar)
        );
        assert_eq!(
            TabStopAlignment::from_xml_value("clear"),
            Some(TabStopAlignment::Clear)
        );
        assert_eq!(TabStopAlignment::from_xml_value("invalid"), None);

        assert_eq!(TabStopAlignment::Left.to_xml_value(), "left");
        assert_eq!(TabStopAlignment::Right.to_xml_value(), "right");
    }

    #[test]
    fn tab_stop_leader_xml_roundtrip() {
        assert_eq!(
            TabStopLeader::from_xml_value("dot"),
            Some(TabStopLeader::Dot)
        );
        assert_eq!(
            TabStopLeader::from_xml_value("hyphen"),
            Some(TabStopLeader::Hyphen)
        );
        assert_eq!(
            TabStopLeader::from_xml_value("underscore"),
            Some(TabStopLeader::Underscore)
        );
        assert_eq!(
            TabStopLeader::from_xml_value("none"),
            Some(TabStopLeader::None)
        );
        assert_eq!(TabStopLeader::from_xml_value("invalid"), None);

        assert_eq!(TabStopLeader::Dot.to_xml_value(), "dot");
        assert_eq!(TabStopLeader::Hyphen.to_xml_value(), "hyphen");
    }

    #[test]
    fn new_properties_affect_has_properties() {
        let mut paragraph = Paragraph::new();
        assert!(!paragraph.has_properties());

        paragraph.set_keep_next(true);
        assert!(paragraph.has_properties());
        paragraph.set_keep_next(false);
        assert!(!paragraph.has_properties());

        paragraph.set_keep_lines(true);
        assert!(paragraph.has_properties());
        paragraph.set_keep_lines(false);

        paragraph.set_widow_control(false);
        assert!(paragraph.has_properties());
        paragraph.clear_widow_control();

        paragraph.add_tab_stop(TabStop::new(720, TabStopAlignment::Left));
        assert!(paragraph.has_properties());
        paragraph.clear_tab_stops();

        paragraph
            .borders_mut()
            .set_top(ParagraphBorder::new("single"));
        assert!(paragraph.has_properties());
        paragraph.clear_borders();

        paragraph.set_shading_color("FFAA00");
        assert!(paragraph.has_properties());
        paragraph.clear_shading_color();

        assert!(!paragraph.has_properties());
    }

    #[test]
    fn paragraph_section_properties_can_be_set_and_cleared() {
        use crate::section::Section;

        let mut paragraph = Paragraph::new();
        assert_eq!(paragraph.section_properties(), None);

        let mut section = Section::new();
        section.set_page_size_twips(12_240, 15_840);
        paragraph.set_section_properties(section);
        assert!(paragraph.section_properties().is_some());
        assert_eq!(
            paragraph
                .section_properties()
                .and_then(|s| s.page_width_twips()),
            Some(12_240)
        );

        paragraph.clear_section_properties();
        assert_eq!(paragraph.section_properties(), None);
    }

    #[test]
    fn paragraph_borders_between_can_be_set() {
        let mut borders = ParagraphBorders::new();
        borders.set_between(ParagraphBorder::new("single"));
        assert_eq!(
            borders.between().and_then(ParagraphBorder::line_type),
            Some("single")
        );
        borders.clear_between();
        assert_eq!(borders.between(), None);
    }

    #[test]
    fn bidi_can_be_set_and_read() {
        let mut paragraph = Paragraph::new();
        assert!(!paragraph.is_bidi());

        paragraph.set_bidi(true);
        assert!(paragraph.is_bidi());
        assert!(paragraph.has_properties());

        paragraph.set_bidi(false);
        assert!(!paragraph.is_bidi());
    }

    #[test]
    fn bidi_affects_has_properties() {
        let mut paragraph = Paragraph::new();
        assert!(!paragraph.has_properties());

        paragraph.set_bidi(true);
        assert!(paragraph.has_properties());

        paragraph.set_bidi(false);
        assert!(!paragraph.has_properties());
    }

    #[test]
    fn comment_range_start_ids_can_be_added_and_cleared() {
        let mut paragraph = Paragraph::new();
        assert!(paragraph.comment_range_start_ids().is_empty());

        paragraph.add_comment_range_start(1);
        paragraph.add_comment_range_start(5);
        assert_eq!(paragraph.comment_range_start_ids(), &[1, 5]);

        paragraph.clear_comment_range_starts();
        assert!(paragraph.comment_range_start_ids().is_empty());
    }

    #[test]
    fn comment_range_end_ids_can_be_added_and_cleared() {
        let mut paragraph = Paragraph::new();
        assert!(paragraph.comment_range_end_ids().is_empty());

        paragraph.add_comment_range_end(1);
        paragraph.add_comment_range_end(5);
        assert_eq!(paragraph.comment_range_end_ids(), &[1, 5]);

        paragraph.clear_comment_range_ends();
        assert!(paragraph.comment_range_end_ids().is_empty());
    }

    #[test]
    fn page_break_before_can_be_set_and_read() {
        let mut paragraph = Paragraph::new();
        assert!(!paragraph.page_break_before());

        paragraph.set_page_break_before(true);
        assert!(paragraph.page_break_before());
        assert!(paragraph.has_properties());

        paragraph.set_page_break_before(false);
        assert!(!paragraph.page_break_before());
        assert!(!paragraph.has_properties());
    }

    #[test]
    fn line_spacing_rule_can_be_set_and_cleared() {
        let mut paragraph = Paragraph::new();
        assert_eq!(paragraph.line_spacing_rule(), None);

        paragraph.set_line_spacing_rule(LineSpacingRule::Exact);
        assert_eq!(paragraph.line_spacing_rule(), Some(LineSpacingRule::Exact));
        assert!(paragraph.has_properties());

        paragraph.set_line_spacing_rule(LineSpacingRule::AtLeast);
        assert_eq!(
            paragraph.line_spacing_rule(),
            Some(LineSpacingRule::AtLeast)
        );

        paragraph.set_line_spacing_rule(LineSpacingRule::Auto);
        assert_eq!(paragraph.line_spacing_rule(), Some(LineSpacingRule::Auto));

        paragraph.clear_line_spacing_rule();
        assert_eq!(paragraph.line_spacing_rule(), None);
        assert!(!paragraph.has_properties());
    }

    #[test]
    fn line_spacing_rule_xml_roundtrip() {
        assert_eq!(
            LineSpacingRule::from_xml_value("auto"),
            Some(LineSpacingRule::Auto)
        );
        assert_eq!(
            LineSpacingRule::from_xml_value("exact"),
            Some(LineSpacingRule::Exact)
        );
        assert_eq!(
            LineSpacingRule::from_xml_value("atLeast"),
            Some(LineSpacingRule::AtLeast)
        );
        assert_eq!(LineSpacingRule::from_xml_value("invalid"), None);

        assert_eq!(LineSpacingRule::Auto.to_xml_value(), "auto");
        assert_eq!(LineSpacingRule::Exact.to_xml_value(), "exact");
        assert_eq!(LineSpacingRule::AtLeast.to_xml_value(), "atLeast");
    }

    #[test]
    fn outline_level_can_be_set_and_cleared() {
        let mut paragraph = Paragraph::new();
        assert_eq!(paragraph.outline_level(), None);

        paragraph.set_outline_level(0);
        assert_eq!(paragraph.outline_level(), Some(0));
        assert!(paragraph.has_properties());

        paragraph.set_outline_level(5);
        assert_eq!(paragraph.outline_level(), Some(5));

        paragraph.set_outline_level(9);
        assert_eq!(paragraph.outline_level(), Some(9));

        paragraph.clear_outline_level();
        assert_eq!(paragraph.outline_level(), None);
        assert!(!paragraph.has_properties());
    }

    #[test]
    fn outline_level_is_clamped_to_valid_range() {
        let mut paragraph = Paragraph::new();

        paragraph.set_outline_level(15);
        assert_eq!(paragraph.outline_level(), Some(9));

        paragraph.set_outline_level(255);
        assert_eq!(paragraph.outline_level(), Some(9));
    }

    #[test]
    fn contextual_spacing_can_be_set_and_read() {
        let mut paragraph = Paragraph::new();
        assert!(!paragraph.contextual_spacing());

        paragraph.set_contextual_spacing(true);
        assert!(paragraph.contextual_spacing());
        assert!(paragraph.has_properties());

        paragraph.set_contextual_spacing(false);
        assert!(!paragraph.contextual_spacing());
        assert!(!paragraph.has_properties());
    }
}
