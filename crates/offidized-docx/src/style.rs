/// Supported style categories for Word styles.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum StyleKind {
    Paragraph,
    Character,
    Table,
}

impl StyleKind {
    pub(crate) fn from_xml_value(value: &str) -> Option<Self> {
        match value {
            "paragraph" => Some(Self::Paragraph),
            "character" => Some(Self::Character),
            "table" => Some(Self::Table),
            _ => None,
        }
    }

    pub(crate) fn to_xml_value(self) -> &'static str {
        match self {
            Self::Paragraph => "paragraph",
            Self::Character => "character",
            Self::Table => "table",
        }
    }
}

/// Basic style definition used by `/word/styles.xml`.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Style {
    kind: StyleKind,
    style_id: String,
    name: Option<String>,
    based_on: Option<String>,
    next_style: Option<String>,
    ui_priority: Option<u32>,
    quick_style: bool,
    semi_hidden: bool,
    unhide_when_used: bool,
    locked: bool,
    paragraph_properties_xml: Option<String>,
    run_properties_xml: Option<String>,
    table_properties_xml: Option<String>,
    table_style_properties_xml: Vec<String>,
}

impl Style {
    pub fn new(kind: StyleKind, style_id: impl Into<String>) -> Self {
        Self {
            kind,
            style_id: normalize_style_id(style_id.into()),
            name: None,
            based_on: None,
            next_style: None,
            ui_priority: None,
            quick_style: false,
            semi_hidden: false,
            unhide_when_used: false,
            locked: false,
            paragraph_properties_xml: None,
            run_properties_xml: None,
            table_properties_xml: None,
            table_style_properties_xml: Vec::new(),
        }
    }

    pub fn paragraph(style_id: impl Into<String>) -> Self {
        Self::new(StyleKind::Paragraph, style_id)
    }

    pub fn character(style_id: impl Into<String>) -> Self {
        Self::new(StyleKind::Character, style_id)
    }

    pub fn table(style_id: impl Into<String>) -> Self {
        Self::new(StyleKind::Table, style_id)
    }

    pub fn kind(&self) -> StyleKind {
        self.kind
    }

    pub fn style_id(&self) -> &str {
        self.style_id.as_str()
    }

    pub fn set_style_id(&mut self, style_id: impl Into<String>) {
        self.style_id = normalize_style_id(style_id.into());
    }

    pub fn name(&self) -> Option<&str> {
        self.name.as_deref()
    }

    pub fn set_name(&mut self, name: impl Into<String>) {
        self.name = normalize_optional_text(name.into());
    }

    pub fn clear_name(&mut self) {
        self.name = None;
    }

    /// The style identifier this style inherits from (`w:basedOn`).
    pub fn based_on(&self) -> Option<&str> {
        self.based_on.as_deref()
    }

    /// Set the parent style this style inherits from.
    pub fn set_based_on(&mut self, style_id: impl Into<String>) {
        self.based_on = normalize_optional_text(style_id.into());
    }

    /// Clear style inheritance.
    pub fn clear_based_on(&mut self) {
        self.based_on = None;
    }

    /// The style that should follow this one (`w:next`).
    pub fn next_style(&self) -> Option<&str> {
        self.next_style.as_deref()
    }

    /// Set the next paragraph style.
    pub fn set_next_style(&mut self, style_id: impl Into<String>) {
        self.next_style = normalize_optional_text(style_id.into());
    }

    /// Clear next paragraph style.
    pub fn clear_next_style(&mut self) {
        self.next_style = None;
    }

    /// The UI sort priority for this style (`w:uiPriority`).
    pub fn ui_priority(&self) -> Option<u32> {
        self.ui_priority
    }

    /// Set the UI sort priority for this style.
    pub fn set_ui_priority(&mut self, priority: u32) {
        self.ui_priority = Some(priority);
    }

    /// Clear the UI sort priority.
    pub fn clear_ui_priority(&mut self) {
        self.ui_priority = None;
    }

    /// Whether this style appears in the Quick Styles gallery (`w:qFormat`).
    pub fn is_quick_style(&self) -> bool {
        self.quick_style
    }

    /// Set whether this style appears in the Quick Styles gallery.
    pub fn set_quick_style(&mut self, value: bool) {
        self.quick_style = value;
    }

    /// Whether this style is semi-hidden from the UI (`w:semiHidden`).
    pub fn is_semi_hidden(&self) -> bool {
        self.semi_hidden
    }

    /// Set whether this style is semi-hidden.
    pub fn set_semi_hidden(&mut self, value: bool) {
        self.semi_hidden = value;
    }

    /// Whether this style becomes visible when used (`w:unhideWhenUsed`).
    pub fn is_unhide_when_used(&self) -> bool {
        self.unhide_when_used
    }

    /// Set whether this style becomes visible when used.
    pub fn set_unhide_when_used(&mut self, value: bool) {
        self.unhide_when_used = value;
    }

    /// Whether this style is locked from editing (`w:locked`).
    pub fn is_locked(&self) -> bool {
        self.locked
    }

    /// Set whether this style is locked from editing.
    pub fn set_locked(&mut self, value: bool) {
        self.locked = value;
    }

    pub fn paragraph_properties_xml(&self) -> Option<&str> {
        self.paragraph_properties_xml.as_deref()
    }

    pub fn set_paragraph_properties_xml(&mut self, xml: impl Into<String>) {
        self.paragraph_properties_xml = normalize_optional_xml_snippet(xml.into());
    }

    pub fn clear_paragraph_properties_xml(&mut self) {
        self.paragraph_properties_xml = None;
    }

    pub fn run_properties_xml(&self) -> Option<&str> {
        self.run_properties_xml.as_deref()
    }

    pub fn set_run_properties_xml(&mut self, xml: impl Into<String>) {
        self.run_properties_xml = normalize_optional_xml_snippet(xml.into());
    }

    pub fn clear_run_properties_xml(&mut self) {
        self.run_properties_xml = None;
    }

    pub fn table_properties_xml(&self) -> Option<&str> {
        self.table_properties_xml.as_deref()
    }

    pub fn set_table_properties_xml(&mut self, xml: impl Into<String>) {
        self.table_properties_xml = normalize_optional_xml_snippet(xml.into());
    }

    pub fn clear_table_properties_xml(&mut self) {
        self.table_properties_xml = None;
    }

    pub fn table_style_properties_xml(&self) -> &[String] {
        &self.table_style_properties_xml
    }

    pub fn add_table_style_properties_xml(&mut self, xml: impl Into<String>) {
        if let Some(snippet) = normalize_optional_xml_snippet(xml.into()) {
            self.table_style_properties_xml.push(snippet);
        }
    }

    pub fn clear_table_style_properties_xml(&mut self) {
        self.table_style_properties_xml.clear();
    }
}

/// Collection of styles to be serialized in `/word/styles.xml`.
#[derive(Debug, Clone, PartialEq, Eq, Default)]
pub struct StyleRegistry {
    styles: Vec<Style>,
}

impl StyleRegistry {
    pub fn new() -> Self {
        Self::default()
    }

    pub fn is_empty(&self) -> bool {
        self.styles.is_empty()
    }

    pub fn styles(&self) -> &[Style] {
        &self.styles
    }

    pub fn paragraph_styles(&self) -> impl Iterator<Item = &Style> {
        self.styles
            .iter()
            .filter(|style| style.kind == StyleKind::Paragraph)
    }

    pub fn character_styles(&self) -> impl Iterator<Item = &Style> {
        self.styles
            .iter()
            .filter(|style| style.kind == StyleKind::Character)
    }

    pub fn table_styles(&self) -> impl Iterator<Item = &Style> {
        self.styles
            .iter()
            .filter(|style| style.kind == StyleKind::Table)
    }

    pub fn style(&self, kind: StyleKind, style_id: &str) -> Option<&Style> {
        let normalized = normalize_style_id(style_id.to_string());
        self.styles
            .iter()
            .find(|style| style.kind == kind && style.style_id == normalized)
    }

    pub fn style_mut(&mut self, kind: StyleKind, style_id: &str) -> Option<&mut Style> {
        let normalized = normalize_style_id(style_id.to_string());
        self.styles
            .iter_mut()
            .find(|style| style.kind == kind && style.style_id == normalized)
    }

    pub fn paragraph_style(&self, style_id: &str) -> Option<&Style> {
        self.style(StyleKind::Paragraph, style_id)
    }

    pub fn character_style(&self, style_id: &str) -> Option<&Style> {
        self.style(StyleKind::Character, style_id)
    }

    pub fn table_style(&self, style_id: &str) -> Option<&Style> {
        self.style(StyleKind::Table, style_id)
    }

    pub fn paragraph_style_mut(&mut self, style_id: &str) -> Option<&mut Style> {
        self.style_mut(StyleKind::Paragraph, style_id)
    }

    pub fn character_style_mut(&mut self, style_id: &str) -> Option<&mut Style> {
        self.style_mut(StyleKind::Character, style_id)
    }

    pub fn table_style_mut(&mut self, style_id: &str) -> Option<&mut Style> {
        self.style_mut(StyleKind::Table, style_id)
    }

    pub fn add_style(&mut self, style: Style) -> &mut Style {
        let key_kind = style.kind;
        let key_style_id = style.style_id.clone();

        if let Some(index) = self
            .styles
            .iter()
            .position(|existing| existing.kind == key_kind && existing.style_id == key_style_id)
        {
            self.styles[index] = style;
            return &mut self.styles[index];
        }

        self.styles.push(style);
        let index = self.styles.len().saturating_sub(1);
        &mut self.styles[index]
    }

    pub fn add_paragraph_style(&mut self, style_id: impl Into<String>) -> &mut Style {
        self.add_style(Style::paragraph(style_id))
    }

    pub fn add_character_style(&mut self, style_id: impl Into<String>) -> &mut Style {
        self.add_style(Style::character(style_id))
    }

    pub fn add_table_style(&mut self, style_id: impl Into<String>) -> &mut Style {
        self.add_style(Style::table(style_id))
    }

    pub fn ensure_style(&mut self, kind: StyleKind, style_id: &str) -> &mut Style {
        let normalized = normalize_style_id(style_id.to_string());
        if let Some(index) = self
            .styles
            .iter()
            .position(|style| style.kind == kind && style.style_id == normalized)
        {
            return &mut self.styles[index];
        }

        self.styles.push(Style::new(kind, normalized));
        let index = self.styles.len().saturating_sub(1);
        &mut self.styles[index]
    }

    pub fn ensure_paragraph_style(&mut self, style_id: &str) -> &mut Style {
        self.ensure_style(StyleKind::Paragraph, style_id)
    }

    pub fn ensure_character_style(&mut self, style_id: &str) -> &mut Style {
        self.ensure_style(StyleKind::Character, style_id)
    }

    pub fn ensure_table_style(&mut self, style_id: &str) -> &mut Style {
        self.ensure_style(StyleKind::Table, style_id)
    }
}

fn normalize_style_id(style_id: String) -> String {
    style_id.trim().to_string()
}

fn normalize_optional_text(value: String) -> Option<String> {
    let trimmed = value.trim();
    if trimmed.is_empty() {
        None
    } else {
        Some(trimmed.to_string())
    }
}

fn normalize_optional_xml_snippet(value: String) -> Option<String> {
    let trimmed = value.trim();
    if trimmed.is_empty() {
        None
    } else {
        Some(trimmed.to_string())
    }
}

#[cfg(test)]
mod tests {
    use super::{Style, StyleKind, StyleRegistry};

    #[test]
    fn style_registry_stores_and_updates_styles_by_key() {
        let mut registry = StyleRegistry::new();
        let body_text = registry.add_paragraph_style("BodyText");
        body_text.set_name("Body Text");
        registry.add_style(Style::new(StyleKind::Paragraph, "BodyText"));

        assert_eq!(registry.styles().len(), 1);
        assert_eq!(
            registry.style(StyleKind::Paragraph, "BodyText"),
            Some(&Style::new(StyleKind::Paragraph, "BodyText"))
        );
    }

    #[test]
    fn ensure_style_returns_existing_or_creates_new() {
        let mut registry = StyleRegistry::new();
        registry.ensure_style(StyleKind::Character, "Emphasis");
        registry.ensure_style(StyleKind::Character, "Emphasis");

        assert_eq!(registry.styles().len(), 1);
        assert_eq!(
            registry
                .style(StyleKind::Character, "Emphasis")
                .map(Style::style_id),
            Some("Emphasis")
        );
    }

    #[test]
    fn table_style_helpers_are_consistent() {
        let mut registry = StyleRegistry::new();
        registry.add_table_style("TableGrid").set_name("Table Grid");

        assert_eq!(registry.table_styles().count(), 1);
        assert_eq!(
            registry.table_style("TableGrid").and_then(Style::name),
            Some("Table Grid")
        );

        registry.ensure_table_style("TableGrid");
        assert_eq!(registry.styles().len(), 1);
    }

    #[test]
    fn style_inheritance_based_on_and_next() {
        let mut style = Style::paragraph("Heading2");
        assert_eq!(style.based_on(), None);
        assert_eq!(style.next_style(), None);

        style.set_based_on("Heading1");
        style.set_next_style("Normal");
        assert_eq!(style.based_on(), Some("Heading1"));
        assert_eq!(style.next_style(), Some("Normal"));

        style.clear_based_on();
        style.clear_next_style();
        assert_eq!(style.based_on(), None);
        assert_eq!(style.next_style(), None);
    }

    #[test]
    fn style_ui_priority() {
        let mut style = Style::paragraph("Normal");
        assert_eq!(style.ui_priority(), None);

        style.set_ui_priority(1);
        assert_eq!(style.ui_priority(), Some(1));

        style.set_ui_priority(99);
        assert_eq!(style.ui_priority(), Some(99));

        style.clear_ui_priority();
        assert_eq!(style.ui_priority(), None);
    }

    #[test]
    fn style_quick_style() {
        let mut style = Style::paragraph("Normal");
        assert!(!style.is_quick_style());

        style.set_quick_style(true);
        assert!(style.is_quick_style());

        style.set_quick_style(false);
        assert!(!style.is_quick_style());
    }

    #[test]
    fn style_semi_hidden() {
        let mut style = Style::paragraph("Normal");
        assert!(!style.is_semi_hidden());

        style.set_semi_hidden(true);
        assert!(style.is_semi_hidden());

        style.set_semi_hidden(false);
        assert!(!style.is_semi_hidden());
    }

    #[test]
    fn style_unhide_when_used() {
        let mut style = Style::paragraph("Normal");
        assert!(!style.is_unhide_when_used());

        style.set_unhide_when_used(true);
        assert!(style.is_unhide_when_used());

        style.set_unhide_when_used(false);
        assert!(!style.is_unhide_when_used());
    }

    #[test]
    fn style_locked() {
        let mut style = Style::paragraph("Normal");
        assert!(!style.is_locked());

        style.set_locked(true);
        assert!(style.is_locked());

        style.set_locked(false);
        assert!(!style.is_locked());
    }

    #[test]
    fn style_new_defaults_for_new_fields() {
        let style = Style::character("Emphasis");
        assert_eq!(style.ui_priority(), None);
        assert!(!style.is_quick_style());
        assert!(!style.is_semi_hidden());
        assert!(!style.is_unhide_when_used());
        assert!(!style.is_locked());
    }

    #[test]
    fn style_stores_xml_property_snippets() {
        let mut style = Style::table("TableGrid");
        style.set_paragraph_properties_xml("<w:pPr><w:spacing w:before=\"120\"/></w:pPr>");
        style.set_run_properties_xml("<w:rPr><w:b/></w:rPr>");
        style.set_table_properties_xml("<w:tblPr><w:tblW w:w=\"0\" w:type=\"auto\"/></w:tblPr>");
        style.add_table_style_properties_xml(
            "<w:tblStylePr w:type=\"firstRow\"><w:rPr><w:b/></w:rPr></w:tblStylePr>",
        );
        style.add_table_style_properties_xml(
            "<w:tblStylePr w:type=\"lastRow\"><w:rPr><w:i/></w:rPr></w:tblStylePr>",
        );

        assert_eq!(
            style.paragraph_properties_xml(),
            Some("<w:pPr><w:spacing w:before=\"120\"/></w:pPr>")
        );
        assert_eq!(style.run_properties_xml(), Some("<w:rPr><w:b/></w:rPr>"));
        assert_eq!(
            style.table_properties_xml(),
            Some("<w:tblPr><w:tblW w:w=\"0\" w:type=\"auto\"/></w:tblPr>")
        );
        assert_eq!(style.table_style_properties_xml().len(), 2);

        style.clear_paragraph_properties_xml();
        style.clear_run_properties_xml();
        style.clear_table_properties_xml();
        style.clear_table_style_properties_xml();
        assert_eq!(style.paragraph_properties_xml(), None);
        assert_eq!(style.run_properties_xml(), None);
        assert_eq!(style.table_properties_xml(), None);
        assert!(style.table_style_properties_xml().is_empty());
    }
}
