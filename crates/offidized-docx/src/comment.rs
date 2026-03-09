use crate::paragraph::Paragraph;
use offidized_opc::RawXmlNode;

/// A comment in a Word document, parsed from `comments.xml`.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Comment {
    id: u32,
    author: String,
    date: Option<String>,
    paragraphs: Vec<Paragraph>,
    /// Unknown children captured for roundtrip fidelity.
    unknown_children: Vec<RawXmlNode>,
}

impl Comment {
    /// Create a new comment with the given id and author.
    pub fn new(id: u32, author: impl Into<String>) -> Self {
        Self {
            id,
            author: author.into(),
            date: None,
            paragraphs: Vec::new(),
            unknown_children: Vec::new(),
        }
    }

    /// Create a comment with a single text paragraph.
    pub fn from_text(id: u32, author: impl Into<String>, text: impl Into<String>) -> Self {
        let mut comment = Self::new(id, author);
        comment.add_paragraph(text);
        comment
    }

    /// Comment id (`w:comment w:id`).
    pub fn id(&self) -> u32 {
        self.id
    }

    /// Set comment id.
    pub fn set_id(&mut self, id: u32) {
        self.id = id;
    }

    /// Comment author (`w:comment w:author`).
    pub fn author(&self) -> &str {
        &self.author
    }

    /// Set comment author.
    pub fn set_author(&mut self, author: impl Into<String>) {
        self.author = author.into();
    }

    /// Comment date (`w:comment w:date`), in ISO 8601 format.
    pub fn date(&self) -> Option<&str> {
        self.date.as_deref()
    }

    /// Set comment date.
    pub fn set_date(&mut self, date: impl Into<String>) {
        let date = date.into();
        self.date = if date.trim().is_empty() {
            None
        } else {
            Some(date)
        };
    }

    /// Clear comment date.
    pub fn clear_date(&mut self) {
        self.date = None;
    }

    /// Paragraphs in this comment.
    pub fn paragraphs(&self) -> &[Paragraph] {
        &self.paragraphs
    }

    /// Mutable paragraphs in this comment.
    pub fn paragraphs_mut(&mut self) -> &mut [Paragraph] {
        &mut self.paragraphs
    }

    /// Add a paragraph to this comment.
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

    /// Concatenated plain text for this comment.
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
    use super::Comment;

    #[test]
    fn comment_stores_id_author_and_paragraphs() {
        let mut comment = Comment::new(1, "John Doe");
        assert_eq!(comment.id(), 1);
        assert_eq!(comment.author(), "John Doe");
        assert_eq!(comment.date(), None);
        assert!(comment.paragraphs().is_empty());

        comment.add_paragraph("This is a comment.");
        comment.add_paragraph("Second paragraph.");
        assert_eq!(comment.paragraphs().len(), 2);
        assert_eq!(comment.text(), "This is a comment.\nSecond paragraph.");
    }

    #[test]
    fn comment_from_text_creates_with_paragraph() {
        let comment = Comment::from_text(5, "Jane", "Review this section");
        assert_eq!(comment.id(), 5);
        assert_eq!(comment.author(), "Jane");
        assert_eq!(comment.paragraphs().len(), 1);
        assert_eq!(comment.text(), "Review this section");
    }

    #[test]
    fn comment_date_can_be_set_and_cleared() {
        let mut comment = Comment::new(1, "Author");
        assert_eq!(comment.date(), None);

        comment.set_date("2024-01-15T10:30:00Z");
        assert_eq!(comment.date(), Some("2024-01-15T10:30:00Z"));

        comment.clear_date();
        assert_eq!(comment.date(), None);
    }

    #[test]
    fn comment_can_be_modified() {
        let mut comment = Comment::from_text(1, "Original Author", "Original text");

        comment.set_id(42);
        comment.set_author("New Author");
        comment.set_date("2024-06-15T14:00:00Z");

        assert_eq!(comment.id(), 42);
        assert_eq!(comment.author(), "New Author");
        assert_eq!(comment.date(), Some("2024-06-15T14:00:00Z"));
    }

    #[test]
    fn comment_paragraphs_can_be_cleared() {
        let mut comment = Comment::from_text(1, "Author", "content");
        assert!(!comment.paragraphs().is_empty());

        comment.clear();
        assert!(comment.paragraphs().is_empty());
    }
}
