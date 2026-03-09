#[derive(Debug, Clone, PartialEq, Eq)]
pub struct SlideComment {
    author: String,
    text: String,
}

impl SlideComment {
    pub fn new(author: impl Into<String>, text: impl Into<String>) -> Self {
        Self {
            author: author.into(),
            text: text.into(),
        }
    }

    pub fn author(&self) -> &str {
        &self.author
    }

    pub fn text(&self) -> &str {
        &self.text
    }

    pub fn set_author(&mut self, author: impl Into<String>) {
        self.author = author.into();
    }

    pub fn set_text(&mut self, text: impl Into<String>) {
        self.text = text.into();
    }
}

#[cfg(test)]
mod tests {
    use super::SlideComment;

    #[test]
    fn stores_author_and_text() {
        let mut comment = SlideComment::new("Alice", "Looks good");
        comment.set_author("Bob");
        comment.set_text("Update this number");

        assert_eq!(comment.author(), "Bob");
        assert_eq!(comment.text(), "Update this number");
    }
}
