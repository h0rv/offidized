use crate::paragraph::Paragraph;
use offidized_opc::RawXmlNode;

/// The type of content a structured document tag (content control) contains.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ContentControlType {
    /// Rich text content (default SDT behavior).
    RichText,
    /// Plain text only (`w:sdtPr > w:text`).
    PlainText,
    /// Picture content (`w:sdtPr > w:picture`).
    Picture,
    /// Combo box allowing typed or selected values (`w:sdtPr > w:comboBox`).
    ComboBox,
    /// Drop-down list with fixed options (`w:sdtPr > w:dropDownList`).
    DropDownList,
    /// Date picker (`w:sdtPr > w:date`).
    Date,
    /// Building block gallery (`w:sdtPr > w:docPartList`).
    BuildingBlock,
    /// Checkbox (`w14:checkbox`).
    Checkbox,
    /// Repeating section (`w15:repeatingSectionItem`).
    RepeatingSection,
}

/// An item in a combo box or drop-down list content control.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ListItem {
    /// The display text shown to the user.
    pub display_text: String,
    /// The underlying value stored in the document.
    pub value: String,
}

impl ListItem {
    /// Create a new list item.
    pub fn new(display_text: impl Into<String>, value: impl Into<String>) -> Self {
        Self {
            display_text: display_text.into(),
            value: value.into(),
        }
    }
}

/// A structured document tag (content control) at the block level (`w:sdt`).
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ContentControl {
    tag: Option<String>,
    alias: Option<String>,
    control_type: Option<ContentControlType>,
    lock_content: bool,
    lock_sdt: bool,
    content: Vec<Paragraph>,
    list_items: Vec<ListItem>,
    unknown_sdt_pr_children: Vec<RawXmlNode>,
}

impl ContentControl {
    /// Create a new empty content control.
    pub fn new() -> Self {
        Self {
            tag: None,
            alias: None,
            control_type: None,
            lock_content: false,
            lock_sdt: false,
            content: Vec::new(),
            list_items: Vec::new(),
            unknown_sdt_pr_children: Vec::new(),
        }
    }

    /// Content control tag (`w:sdtPr > w:tag w:val`).
    pub fn tag(&self) -> Option<&str> {
        self.tag.as_deref()
    }

    /// Set content control tag.
    pub fn set_tag(&mut self, tag: impl Into<String>) {
        let tag = tag.into();
        self.tag = if tag.trim().is_empty() {
            None
        } else {
            Some(tag)
        };
    }

    /// Clear content control tag.
    pub fn clear_tag(&mut self) {
        self.tag = None;
    }

    /// Content control alias/title (`w:sdtPr > w:alias w:val`).
    pub fn alias(&self) -> Option<&str> {
        self.alias.as_deref()
    }

    /// Set content control alias.
    pub fn set_alias(&mut self, alias: impl Into<String>) {
        let alias = alias.into();
        self.alias = if alias.trim().is_empty() {
            None
        } else {
            Some(alias)
        };
    }

    /// Clear content control alias.
    pub fn clear_alias(&mut self) {
        self.alias = None;
    }

    /// The type of content this content control contains.
    pub fn control_type(&self) -> Option<ContentControlType> {
        self.control_type
    }

    /// Set the type of content this content control contains.
    pub fn set_control_type(&mut self, control_type: ContentControlType) {
        self.control_type = Some(control_type);
    }

    /// Clear the content control type.
    pub fn clear_control_type(&mut self) {
        self.control_type = None;
    }

    /// Whether the content of this SDT is locked from editing (`w:sdtPr > w:lock w:val="sdtContentLocked"`).
    pub fn lock_content(&self) -> bool {
        self.lock_content
    }

    /// Set whether the content of this SDT is locked from editing.
    pub fn set_lock_content(&mut self, value: bool) {
        self.lock_content = value;
    }

    /// Whether this SDT itself is locked from deletion (`w:sdtPr > w:lock w:val="sdtLocked"`).
    pub fn lock_sdt(&self) -> bool {
        self.lock_sdt
    }

    /// Set whether this SDT is locked from deletion.
    pub fn set_lock_sdt(&mut self, value: bool) {
        self.lock_sdt = value;
    }

    /// Paragraphs inside this content control (`w:sdtContent`).
    pub fn content(&self) -> &[Paragraph] {
        &self.content
    }

    /// Mutable paragraphs inside this content control.
    pub fn content_mut(&mut self) -> &mut [Paragraph] {
        &mut self.content
    }

    /// Add a paragraph to the content control.
    pub fn add_paragraph(&mut self, text: impl Into<String>) -> &mut Paragraph {
        self.content.push(Paragraph::from_text(text));
        let index = self.content.len().saturating_sub(1);
        &mut self.content[index]
    }

    /// Replace all content paragraphs.
    pub fn set_content(&mut self, paragraphs: Vec<Paragraph>) {
        self.content = paragraphs;
    }

    /// Clear all content paragraphs.
    pub fn clear_content(&mut self) {
        self.content.clear();
    }

    /// Returns the list items for combo box or drop-down content controls.
    pub fn list_items(&self) -> &[ListItem] {
        &self.list_items
    }

    /// Add an item to the combo box or drop-down list.
    pub fn add_list_item(&mut self, display_text: impl Into<String>, value: impl Into<String>) {
        self.list_items.push(ListItem::new(display_text, value));
    }

    /// Remove all list items.
    pub fn clear_list_items(&mut self) {
        self.list_items.clear();
    }

    /// Unknown sdtPr children captured for roundtrip fidelity.
    #[allow(dead_code)]
    pub(crate) fn unknown_sdt_pr_children(&self) -> &[RawXmlNode] {
        self.unknown_sdt_pr_children.as_slice()
    }

    /// Push an unknown sdtPr child node.
    pub(crate) fn push_unknown_sdt_pr_child(&mut self, node: RawXmlNode) {
        self.unknown_sdt_pr_children.push(node);
    }

    /// Concatenated plain text of the content control paragraphs.
    pub fn text(&self) -> String {
        self.content
            .iter()
            .map(|p| p.text())
            .collect::<Vec<_>>()
            .join("\n")
    }
}

impl Default for ContentControl {
    fn default() -> Self {
        Self::new()
    }
}

#[cfg(test)]
mod tests {
    use super::{ContentControl, ContentControlType};

    #[test]
    fn content_control_stores_tag_alias_and_content() {
        let mut sdt = ContentControl::new();
        assert_eq!(sdt.tag(), None);
        assert_eq!(sdt.alias(), None);
        assert!(sdt.content().is_empty());

        sdt.set_tag("myTag");
        sdt.set_alias("My Content Control");
        sdt.add_paragraph("Hello");
        sdt.add_paragraph("World");

        assert_eq!(sdt.tag(), Some("myTag"));
        assert_eq!(sdt.alias(), Some("My Content Control"));
        assert_eq!(sdt.content().len(), 2);
        assert_eq!(sdt.text(), "Hello\nWorld");
    }

    #[test]
    fn content_control_can_be_cleared() {
        let mut sdt = ContentControl::new();
        sdt.set_tag("tag1");
        sdt.set_alias("alias1");
        sdt.add_paragraph("content");

        sdt.clear_tag();
        sdt.clear_alias();
        sdt.clear_content();

        assert_eq!(sdt.tag(), None);
        assert_eq!(sdt.alias(), None);
        assert!(sdt.content().is_empty());
    }

    #[test]
    fn content_control_whitespace_tag_normalized() {
        let mut sdt = ContentControl::new();
        sdt.set_tag("  ");
        assert_eq!(sdt.tag(), None);

        sdt.set_alias("");
        assert_eq!(sdt.alias(), None);
    }

    #[test]
    fn content_control_type_get_set_clear() {
        let mut sdt = ContentControl::new();
        assert_eq!(sdt.control_type(), None);

        sdt.set_control_type(ContentControlType::RichText);
        assert_eq!(sdt.control_type(), Some(ContentControlType::RichText));

        sdt.set_control_type(ContentControlType::PlainText);
        assert_eq!(sdt.control_type(), Some(ContentControlType::PlainText));

        sdt.set_control_type(ContentControlType::Picture);
        assert_eq!(sdt.control_type(), Some(ContentControlType::Picture));

        sdt.set_control_type(ContentControlType::ComboBox);
        assert_eq!(sdt.control_type(), Some(ContentControlType::ComboBox));

        sdt.set_control_type(ContentControlType::DropDownList);
        assert_eq!(sdt.control_type(), Some(ContentControlType::DropDownList));

        sdt.set_control_type(ContentControlType::Date);
        assert_eq!(sdt.control_type(), Some(ContentControlType::Date));

        sdt.set_control_type(ContentControlType::BuildingBlock);
        assert_eq!(sdt.control_type(), Some(ContentControlType::BuildingBlock));

        sdt.set_control_type(ContentControlType::Checkbox);
        assert_eq!(sdt.control_type(), Some(ContentControlType::Checkbox));

        sdt.set_control_type(ContentControlType::RepeatingSection);
        assert_eq!(
            sdt.control_type(),
            Some(ContentControlType::RepeatingSection)
        );

        sdt.clear_control_type();
        assert_eq!(sdt.control_type(), None);
    }

    #[test]
    fn content_control_lock_content() {
        let mut sdt = ContentControl::new();
        assert!(!sdt.lock_content());

        sdt.set_lock_content(true);
        assert!(sdt.lock_content());

        sdt.set_lock_content(false);
        assert!(!sdt.lock_content());
    }

    #[test]
    fn content_control_lock_sdt() {
        let mut sdt = ContentControl::new();
        assert!(!sdt.lock_sdt());

        sdt.set_lock_sdt(true);
        assert!(sdt.lock_sdt());

        sdt.set_lock_sdt(false);
        assert!(!sdt.lock_sdt());
    }

    #[test]
    fn content_control_new_defaults() {
        let sdt = ContentControl::new();
        assert_eq!(sdt.tag(), None);
        assert_eq!(sdt.alias(), None);
        assert_eq!(sdt.control_type(), None);
        assert!(!sdt.lock_content());
        assert!(!sdt.lock_sdt());
        assert!(sdt.content().is_empty());
    }
}
