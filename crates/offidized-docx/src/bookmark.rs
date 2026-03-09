/// A bookmark in a Word document, parsed from `w:bookmarkStart` and `w:bookmarkEnd` elements.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Bookmark {
    id: u32,
    name: String,
    start_paragraph_index: usize,
    end_paragraph_index: usize,
}

impl Bookmark {
    /// Create a new bookmark spanning from `start_paragraph_index` to `end_paragraph_index`.
    pub fn new(
        id: u32,
        name: impl Into<String>,
        start_paragraph_index: usize,
        end_paragraph_index: usize,
    ) -> Self {
        Self {
            id,
            name: name.into(),
            start_paragraph_index,
            end_paragraph_index,
        }
    }

    /// Bookmark id (`w:bookmarkStart w:id`).
    pub fn id(&self) -> u32 {
        self.id
    }

    /// Set bookmark id.
    pub fn set_id(&mut self, id: u32) {
        self.id = id;
    }

    /// Bookmark name (`w:bookmarkStart w:name`).
    pub fn name(&self) -> &str {
        &self.name
    }

    /// Set bookmark name.
    pub fn set_name(&mut self, name: impl Into<String>) {
        self.name = name.into();
    }

    /// Index of the paragraph where this bookmark starts.
    pub fn start_paragraph_index(&self) -> usize {
        self.start_paragraph_index
    }

    /// Set the start paragraph index.
    pub fn set_start_paragraph_index(&mut self, index: usize) {
        self.start_paragraph_index = index;
    }

    /// Index of the paragraph where this bookmark ends.
    pub fn end_paragraph_index(&self) -> usize {
        self.end_paragraph_index
    }

    /// Set the end paragraph index.
    pub fn set_end_paragraph_index(&mut self, index: usize) {
        self.end_paragraph_index = index;
    }

    /// Whether this bookmark spans exactly one paragraph (start == end).
    pub fn is_single_paragraph(&self) -> bool {
        self.start_paragraph_index == self.end_paragraph_index
    }
}

#[cfg(test)]
mod tests {
    use super::Bookmark;

    #[test]
    fn bookmark_stores_id_name_and_range() {
        let bookmark = Bookmark::new(0, "_GoBack", 2, 5);
        assert_eq!(bookmark.id(), 0);
        assert_eq!(bookmark.name(), "_GoBack");
        assert_eq!(bookmark.start_paragraph_index(), 2);
        assert_eq!(bookmark.end_paragraph_index(), 5);
        assert!(!bookmark.is_single_paragraph());
    }

    #[test]
    fn bookmark_can_be_modified() {
        let mut bookmark = Bookmark::new(1, "Introduction", 0, 0);
        assert!(bookmark.is_single_paragraph());

        bookmark.set_id(42);
        bookmark.set_name("Chapter1");
        bookmark.set_start_paragraph_index(3);
        bookmark.set_end_paragraph_index(10);

        assert_eq!(bookmark.id(), 42);
        assert_eq!(bookmark.name(), "Chapter1");
        assert_eq!(bookmark.start_paragraph_index(), 3);
        assert_eq!(bookmark.end_paragraph_index(), 10);
        assert!(!bookmark.is_single_paragraph());
    }

    #[test]
    fn single_paragraph_bookmark() {
        let bookmark = Bookmark::new(5, "Marker", 7, 7);
        assert!(bookmark.is_single_paragraph());
    }
}
