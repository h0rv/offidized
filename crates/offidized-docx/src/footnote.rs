use crate::paragraph::Paragraph;
use offidized_opc::RawXmlNode;

/// A footnote in a Word document, stored in `footnotes.xml`.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Footnote {
    id: u32,
    paragraphs: Vec<Paragraph>,
    /// Unknown children captured for roundtrip fidelity.
    unknown_children: Vec<RawXmlNode>,
}

impl Footnote {
    /// Create a new footnote with the given id.
    pub fn new(id: u32) -> Self {
        Self {
            id,
            paragraphs: Vec::new(),
            unknown_children: Vec::new(),
        }
    }

    /// Create a footnote with a single text paragraph.
    pub fn from_text(id: u32, text: impl Into<String>) -> Self {
        let mut footnote = Self::new(id);
        footnote.add_paragraph(text);
        footnote
    }

    /// Footnote id (`w:footnote w:id`).
    pub fn id(&self) -> u32 {
        self.id
    }

    /// Set footnote id.
    pub fn set_id(&mut self, id: u32) {
        self.id = id;
    }

    /// Paragraphs in this footnote.
    pub fn paragraphs(&self) -> &[Paragraph] {
        &self.paragraphs
    }

    /// Mutable paragraphs in this footnote.
    pub fn paragraphs_mut(&mut self) -> &mut [Paragraph] {
        &mut self.paragraphs
    }

    /// Add a paragraph to this footnote.
    pub fn add_paragraph(&mut self, text: impl Into<String>) -> &mut Paragraph {
        self.paragraphs.push(Paragraph::from_text(text));
        let index = self.paragraphs.len().saturating_sub(1);
        &mut self.paragraphs[index]
    }

    /// Replace all paragraphs.
    pub fn set_paragraphs(&mut self, paragraphs: Vec<Paragraph>) {
        self.paragraphs = paragraphs;
    }

    /// Clear all paragraphs.
    pub fn clear(&mut self) {
        self.paragraphs.clear();
    }

    /// Concatenated plain text for this footnote.
    pub fn text(&self) -> String {
        self.paragraphs
            .iter()
            .map(|p| p.text())
            .collect::<Vec<_>>()
            .join("\n")
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

/// An endnote in a Word document, stored in `endnotes.xml`.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Endnote {
    id: u32,
    paragraphs: Vec<Paragraph>,
    /// Unknown children captured for roundtrip fidelity.
    unknown_children: Vec<RawXmlNode>,
}

impl Endnote {
    /// Create a new endnote with the given id.
    pub fn new(id: u32) -> Self {
        Self {
            id,
            paragraphs: Vec::new(),
            unknown_children: Vec::new(),
        }
    }

    /// Create an endnote with a single text paragraph.
    pub fn from_text(id: u32, text: impl Into<String>) -> Self {
        let mut endnote = Self::new(id);
        endnote.add_paragraph(text);
        endnote
    }

    /// Endnote id (`w:endnote w:id`).
    pub fn id(&self) -> u32 {
        self.id
    }

    /// Set endnote id.
    pub fn set_id(&mut self, id: u32) {
        self.id = id;
    }

    /// Paragraphs in this endnote.
    pub fn paragraphs(&self) -> &[Paragraph] {
        &self.paragraphs
    }

    /// Mutable paragraphs in this endnote.
    pub fn paragraphs_mut(&mut self) -> &mut [Paragraph] {
        &mut self.paragraphs
    }

    /// Add a paragraph to this endnote.
    pub fn add_paragraph(&mut self, text: impl Into<String>) -> &mut Paragraph {
        self.paragraphs.push(Paragraph::from_text(text));
        let index = self.paragraphs.len().saturating_sub(1);
        &mut self.paragraphs[index]
    }

    /// Replace all paragraphs.
    pub fn set_paragraphs(&mut self, paragraphs: Vec<Paragraph>) {
        self.paragraphs = paragraphs;
    }

    /// Clear all paragraphs.
    pub fn clear(&mut self) {
        self.paragraphs.clear();
    }

    /// Concatenated plain text for this endnote.
    pub fn text(&self) -> String {
        self.paragraphs
            .iter()
            .map(|p| p.text())
            .collect::<Vec<_>>()
            .join("\n")
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

#[cfg(test)]
mod tests {
    use super::{Endnote, Footnote};

    #[test]
    fn footnote_stores_id_and_paragraphs() {
        let mut footnote = Footnote::new(1);
        assert_eq!(footnote.id(), 1);
        assert!(footnote.paragraphs().is_empty());

        footnote.add_paragraph("First paragraph");
        footnote.add_paragraph("Second paragraph");
        assert_eq!(footnote.paragraphs().len(), 2);
        assert_eq!(footnote.text(), "First paragraph\nSecond paragraph");

        footnote.set_id(42);
        assert_eq!(footnote.id(), 42);
    }

    #[test]
    fn footnote_from_text_creates_with_paragraph() {
        let footnote = Footnote::from_text(3, "Footnote text");
        assert_eq!(footnote.id(), 3);
        assert_eq!(footnote.paragraphs().len(), 1);
        assert_eq!(footnote.text(), "Footnote text");
    }

    #[test]
    fn footnote_can_be_cleared() {
        let mut footnote = Footnote::from_text(1, "content");
        assert!(!footnote.paragraphs().is_empty());

        footnote.clear();
        assert!(footnote.paragraphs().is_empty());
    }

    #[test]
    fn endnote_stores_id_and_paragraphs() {
        let mut endnote = Endnote::new(1);
        assert_eq!(endnote.id(), 1);
        assert!(endnote.paragraphs().is_empty());

        endnote.add_paragraph("First paragraph");
        endnote.add_paragraph("Second paragraph");
        assert_eq!(endnote.paragraphs().len(), 2);
        assert_eq!(endnote.text(), "First paragraph\nSecond paragraph");

        endnote.set_id(99);
        assert_eq!(endnote.id(), 99);
    }

    #[test]
    fn endnote_from_text_creates_with_paragraph() {
        let endnote = Endnote::from_text(7, "Endnote text");
        assert_eq!(endnote.id(), 7);
        assert_eq!(endnote.paragraphs().len(), 1);
        assert_eq!(endnote.text(), "Endnote text");
    }

    #[test]
    fn endnote_can_be_cleared() {
        let mut endnote = Endnote::from_text(1, "content");
        assert!(!endnote.paragraphs().is_empty());

        endnote.clear();
        assert!(endnote.paragraphs().is_empty());
    }
}
