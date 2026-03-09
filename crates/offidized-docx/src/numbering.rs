use offidized_opc::RawXmlNode;

/// A single level definition within a numbering definition.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct NumberingLevel {
    level: u8,
    start: u32,
    format: String,
    text: String,
    alignment: Option<String>,
    /// Left indentation in twips for this level (`w:pPr/w:ind/@w:left`).
    indent_left_twips: Option<u32>,
    /// Hanging indentation in twips for this level (`w:pPr/w:ind/@w:hanging`).
    indent_hanging_twips: Option<u32>,
    /// Tab stop position in twips for this level (`w:pPr/w:tabs/w:tab/@w:pos`).
    tab_stop_twips: Option<u32>,
    /// Suffix after number: tab (default), space, or nothing (`w:suff/@w:val`).
    suffix: Option<String>,
    /// Font family for the number text (`w:rPr/w:rFonts/@w:ascii`).
    font_family: Option<String>,
    /// Font size in half-points for the number text (`w:rPr/w:sz/@w:val`).
    font_size_half_points: Option<u16>,
    /// Whether the number text is bold (`w:rPr/w:b`).
    bold: Option<bool>,
    /// Whether the number text is italic (`w:rPr/w:i`).
    italic: Option<bool>,
    /// Color of the number text (`w:rPr/w:color/@w:val`).
    color: Option<String>,
    /// Unknown children captured for roundtrip fidelity.
    unknown_children: Vec<RawXmlNode>,
}

impl NumberingLevel {
    /// Create a new numbering level.
    pub fn new(level: u8, start: u32, format: impl Into<String>, text: impl Into<String>) -> Self {
        Self {
            level,
            start,
            format: format.into(),
            text: text.into(),
            alignment: None,
            indent_left_twips: None,
            indent_hanging_twips: None,
            tab_stop_twips: None,
            suffix: None,
            font_family: None,
            font_size_half_points: None,
            bold: None,
            italic: None,
            color: None,
            unknown_children: Vec::new(),
        }
    }

    /// The 0-based level index (`w:lvl w:ilvl`).
    pub fn level(&self) -> u8 {
        self.level
    }

    /// Set the level index.
    pub fn set_level(&mut self, level: u8) {
        self.level = level;
    }

    /// Starting number for this level (`w:start w:val`).
    pub fn start(&self) -> u32 {
        self.start
    }

    /// Set starting number.
    pub fn set_start(&mut self, start: u32) {
        self.start = start;
    }

    /// Number format for this level (`w:numFmt w:val`), e.g., `"decimal"`, `"bullet"`, `"lowerRoman"`.
    pub fn format(&self) -> &str {
        &self.format
    }

    /// Set number format.
    pub fn set_format(&mut self, format: impl Into<String>) {
        self.format = format.into();
    }

    /// Level text template (`w:lvlText w:val`), e.g., `"%1."`, `"%1.%2"`.
    pub fn text(&self) -> &str {
        &self.text
    }

    /// Set level text template.
    pub fn set_text(&mut self, text: impl Into<String>) {
        self.text = text.into();
    }

    /// Alignment for this level (`w:lvlJc w:val`), e.g., `"left"`, `"center"`, `"right"`.
    pub fn alignment(&self) -> Option<&str> {
        self.alignment.as_deref()
    }

    /// Set alignment.
    pub fn set_alignment(&mut self, alignment: impl Into<String>) {
        let alignment = alignment.into();
        self.alignment = if alignment.trim().is_empty() {
            None
        } else {
            Some(alignment)
        };
    }

    /// Clear alignment.
    pub fn clear_alignment(&mut self) {
        self.alignment = None;
    }

    /// Left indentation in twips for this level (`w:pPr/w:ind/@w:left`).
    pub fn indent_left_twips(&self) -> Option<u32> {
        self.indent_left_twips
    }

    /// Set left indentation in twips.
    pub fn set_indent_left_twips(&mut self, value: u32) {
        self.indent_left_twips = Some(value);
    }

    /// Clear left indentation.
    pub fn clear_indent_left_twips(&mut self) {
        self.indent_left_twips = None;
    }

    /// Hanging indentation in twips for this level (`w:pPr/w:ind/@w:hanging`).
    pub fn indent_hanging_twips(&self) -> Option<u32> {
        self.indent_hanging_twips
    }

    /// Set hanging indentation in twips.
    pub fn set_indent_hanging_twips(&mut self, value: u32) {
        self.indent_hanging_twips = Some(value);
    }

    /// Clear hanging indentation.
    pub fn clear_indent_hanging_twips(&mut self) {
        self.indent_hanging_twips = None;
    }

    /// Tab stop position in twips for this level (`w:pPr/w:tabs/w:tab/@w:pos`).
    pub fn tab_stop_twips(&self) -> Option<u32> {
        self.tab_stop_twips
    }

    /// Set tab stop position in twips.
    pub fn set_tab_stop_twips(&mut self, value: u32) {
        self.tab_stop_twips = Some(value);
    }

    /// Clear tab stop position.
    pub fn clear_tab_stop_twips(&mut self) {
        self.tab_stop_twips = None;
    }

    /// Suffix after number: tab (default), space, or nothing (`w:suff/@w:val`).
    pub fn suffix(&self) -> Option<&str> {
        self.suffix.as_deref()
    }

    /// Set suffix.
    pub fn set_suffix(&mut self, suffix: impl Into<String>) {
        self.suffix = Some(suffix.into());
    }

    /// Clear suffix.
    pub fn clear_suffix(&mut self) {
        self.suffix = None;
    }

    /// Font family for the number text (`w:rPr/w:rFonts/@w:ascii`).
    pub fn font_family(&self) -> Option<&str> {
        self.font_family.as_deref()
    }

    /// Set font family.
    pub fn set_font_family(&mut self, font: impl Into<String>) {
        self.font_family = Some(font.into());
    }

    /// Clear font family.
    pub fn clear_font_family(&mut self) {
        self.font_family = None;
    }

    /// Font size in half-points for the number text (`w:rPr/w:sz/@w:val`).
    pub fn font_size_half_points(&self) -> Option<u16> {
        self.font_size_half_points
    }

    /// Set font size in half-points.
    pub fn set_font_size_half_points(&mut self, size: u16) {
        self.font_size_half_points = Some(size);
    }

    /// Clear font size.
    pub fn clear_font_size_half_points(&mut self) {
        self.font_size_half_points = None;
    }

    /// Whether the number text is bold (`w:rPr/w:b`).
    pub fn bold(&self) -> Option<bool> {
        self.bold
    }

    /// Set bold.
    pub fn set_bold(&mut self, bold: bool) {
        self.bold = Some(bold);
    }

    /// Clear bold.
    pub fn clear_bold(&mut self) {
        self.bold = None;
    }

    /// Whether the number text is italic (`w:rPr/w:i`).
    pub fn italic(&self) -> Option<bool> {
        self.italic
    }

    /// Set italic.
    pub fn set_italic(&mut self, italic: bool) {
        self.italic = Some(italic);
    }

    /// Clear italic.
    pub fn clear_italic(&mut self) {
        self.italic = None;
    }

    /// Color of the number text (`w:rPr/w:color/@w:val`).
    pub fn color(&self) -> Option<&str> {
        self.color.as_deref()
    }

    /// Set color.
    pub fn set_color(&mut self, color: impl Into<String>) {
        self.color = Some(color.into());
    }

    /// Clear color.
    pub fn clear_color(&mut self) {
        self.color = None;
    }

    /// Unknown children captured for roundtrip fidelity.
    #[allow(dead_code)]
    pub(crate) fn unknown_children(&self) -> &[RawXmlNode] {
        self.unknown_children.as_slice()
    }

    /// Push an unknown child node.
    #[allow(dead_code)]
    pub(crate) fn push_unknown_child(&mut self, node: RawXmlNode) {
        self.unknown_children.push(node);
    }
}

/// A numbering definition parsed from `numbering.xml` (`w:abstractNum`).
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct NumberingDefinition {
    abstract_num_id: u32,
    levels: Vec<NumberingLevel>,
    /// Unknown children captured for roundtrip fidelity.
    unknown_children: Vec<RawXmlNode>,
}

impl NumberingDefinition {
    /// Create a new numbering definition with the given abstract num id.
    pub fn new(abstract_num_id: u32) -> Self {
        Self {
            abstract_num_id,
            levels: Vec::new(),
            unknown_children: Vec::new(),
        }
    }

    /// Abstract numbering definition id (`w:abstractNum w:abstractNumId`).
    pub fn abstract_num_id(&self) -> u32 {
        self.abstract_num_id
    }

    /// Set abstract numbering definition id.
    pub fn set_abstract_num_id(&mut self, id: u32) {
        self.abstract_num_id = id;
    }

    /// Levels in this numbering definition.
    pub fn levels(&self) -> &[NumberingLevel] {
        &self.levels
    }

    /// Mutable levels.
    pub fn levels_mut(&mut self) -> &mut [NumberingLevel] {
        &mut self.levels
    }

    /// Add a level to this numbering definition.
    pub fn add_level(&mut self, level: NumberingLevel) {
        self.levels.push(level);
    }

    /// Replace all levels.
    pub fn set_levels(&mut self, levels: Vec<NumberingLevel>) {
        self.levels = levels;
    }

    /// Clear all levels.
    pub fn clear_levels(&mut self) {
        self.levels.clear();
    }

    /// Get a level by its index.
    pub fn level(&self, ilvl: u8) -> Option<&NumberingLevel> {
        self.levels.iter().find(|level| level.level == ilvl)
    }

    /// Unknown children captured for roundtrip fidelity.
    #[allow(dead_code)]
    pub(crate) fn unknown_children(&self) -> &[RawXmlNode] {
        self.unknown_children.as_slice()
    }

    /// Push an unknown child node.
    #[allow(dead_code)]
    pub(crate) fn push_unknown_child(&mut self, node: RawXmlNode) {
        self.unknown_children.push(node);
    }

    /// Create a standard bullet numbering definition.
    ///
    /// Produces a 9-level definition using the `"bullet"` number format.
    pub fn create_bullet(abstract_num_id: u32) -> Self {
        let mut def = Self::new(abstract_num_id);
        let fonts = [
            "\u{2022}", "o", "\u{25AA}", "\u{2022}", "o", "\u{25AA}", "\u{2022}", "o", "\u{25AA}",
        ];
        for (i, symbol) in fonts.iter().enumerate() {
            let ilvl = i as u8;
            let indent = (ilvl as u32 + 1) * 360;
            let mut level = NumberingLevel::new(ilvl, 1, "bullet", *symbol);
            level.set_alignment("left");
            level.set_indent_left_twips(indent);
            level.set_indent_hanging_twips(360);
            def.add_level(level);
        }
        def
    }

    /// Create a standard decimal numbered list definition.
    ///
    /// Produces a 9-level definition using `"decimal"` format with patterns like `"%1."`, `"%1.%2."`.
    pub fn create_numbered(abstract_num_id: u32) -> Self {
        let mut def = Self::new(abstract_num_id);
        for i in 0u8..9 {
            let indent = (i as u32 + 1) * 360;
            let text = if i == 0 {
                "%1.".to_string()
            } else {
                let parts: Vec<String> = (1..=i as u32 + 1).map(|n| format!("%{n}")).collect();
                format!("{}.", parts.join("."))
            };
            let mut level = NumberingLevel::new(i, 1, "decimal", text);
            level.set_alignment("left");
            level.set_indent_left_twips(indent);
            level.set_indent_hanging_twips(360);
            def.add_level(level);
        }
        def
    }
}

/// A single level override within a numbering instance.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct NumberingLevelOverride {
    level: u8,
    start_override: Option<u32>,
}

impl NumberingLevelOverride {
    /// Create a new level override.
    pub fn new(level: u8) -> Self {
        Self {
            level,
            start_override: None,
        }
    }

    /// The level index being overridden.
    pub fn level(&self) -> u8 {
        self.level
    }

    /// Set the level index.
    pub fn set_level(&mut self, level: u8) {
        self.level = level;
    }

    /// The starting number override (`w:startOverride w:val`).
    pub fn start_override(&self) -> Option<u32> {
        self.start_override
    }

    /// Set the starting number override.
    pub fn set_start_override(&mut self, start: u32) {
        self.start_override = Some(start);
    }

    /// Clear the starting number override.
    pub fn clear_start_override(&mut self) {
        self.start_override = None;
    }
}

/// A numbering instance (`w:num`) that references an abstract numbering definition.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct NumberingInstance {
    num_id: u32,
    abstract_num_id: u32,
    level_overrides: Vec<NumberingLevelOverride>,
}

impl NumberingInstance {
    /// Create a new numbering instance.
    pub fn new(num_id: u32, abstract_num_id: u32) -> Self {
        Self {
            num_id,
            abstract_num_id,
            level_overrides: Vec::new(),
        }
    }

    /// Numbering instance id (`w:num w:numId`).
    pub fn num_id(&self) -> u32 {
        self.num_id
    }

    /// Set numbering instance id.
    pub fn set_num_id(&mut self, id: u32) {
        self.num_id = id;
    }

    /// Abstract numbering definition id this instance references.
    pub fn abstract_num_id(&self) -> u32 {
        self.abstract_num_id
    }

    /// Set abstract numbering definition id.
    pub fn set_abstract_num_id(&mut self, id: u32) {
        self.abstract_num_id = id;
    }

    /// Level overrides for this instance.
    pub fn level_overrides(&self) -> &[NumberingLevelOverride] {
        &self.level_overrides
    }

    /// Add a level override.
    pub fn add_level_override(&mut self, override_val: NumberingLevelOverride) {
        self.level_overrides.push(override_val);
    }

    /// Replace all level overrides.
    pub fn set_level_overrides(&mut self, overrides: Vec<NumberingLevelOverride>) {
        self.level_overrides = overrides;
    }

    /// Clear all level overrides.
    pub fn clear_level_overrides(&mut self) {
        self.level_overrides.clear();
    }
}

#[cfg(test)]
mod tests {
    use super::{NumberingDefinition, NumberingLevel};

    #[test]
    fn numbering_level_stores_all_properties() {
        let mut level = NumberingLevel::new(0, 1, "decimal", "%1.");
        assert_eq!(level.level(), 0);
        assert_eq!(level.start(), 1);
        assert_eq!(level.format(), "decimal");
        assert_eq!(level.text(), "%1.");
        assert_eq!(level.alignment(), None);

        level.set_alignment("left");
        assert_eq!(level.alignment(), Some("left"));

        level.set_level(2);
        level.set_start(5);
        level.set_format("lowerRoman");
        level.set_text("%1.%2.%3");
        assert_eq!(level.level(), 2);
        assert_eq!(level.start(), 5);
        assert_eq!(level.format(), "lowerRoman");
        assert_eq!(level.text(), "%1.%2.%3");

        level.clear_alignment();
        assert_eq!(level.alignment(), None);
    }

    #[test]
    fn numbering_definition_stores_levels() {
        let mut def = NumberingDefinition::new(0);
        assert_eq!(def.abstract_num_id(), 0);
        assert!(def.levels().is_empty());

        def.add_level(NumberingLevel::new(0, 1, "decimal", "%1."));
        def.add_level(NumberingLevel::new(1, 1, "lowerLetter", "%2)"));

        assert_eq!(def.levels().len(), 2);
        assert_eq!(def.level(0).map(NumberingLevel::format), Some("decimal"));
        assert_eq!(
            def.level(1).map(NumberingLevel::format),
            Some("lowerLetter")
        );
        assert_eq!(def.level(2), None);
    }

    #[test]
    fn numbering_definition_can_be_cleared() {
        let mut def = NumberingDefinition::new(1);
        def.add_level(NumberingLevel::new(0, 1, "bullet", ""));
        assert!(!def.levels().is_empty());

        def.clear_levels();
        assert!(def.levels().is_empty());
    }

    #[test]
    fn numbering_definition_id_can_be_set() {
        let mut def = NumberingDefinition::new(0);
        def.set_abstract_num_id(42);
        assert_eq!(def.abstract_num_id(), 42);
    }

    #[test]
    fn numbering_level_alignment_whitespace_normalized() {
        let mut level = NumberingLevel::new(0, 1, "decimal", "%1.");
        level.set_alignment("  ");
        assert_eq!(level.alignment(), None);
    }
}
