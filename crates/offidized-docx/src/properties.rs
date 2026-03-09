/// Core document properties parsed from `docProps/core.xml`.
#[derive(Debug, Clone, PartialEq, Eq, Default)]
pub struct DocumentProperties {
    title: Option<String>,
    subject: Option<String>,
    creator: Option<String>,
    description: Option<String>,
    keywords: Option<String>,
    last_modified_by: Option<String>,
    created: Option<String>,
    modified: Option<String>,
}

impl DocumentProperties {
    /// Create empty document properties.
    pub fn new() -> Self {
        Self::default()
    }

    /// Document title (`dc:title`).
    pub fn title(&self) -> Option<&str> {
        self.title.as_deref()
    }

    /// Set document title.
    pub fn set_title(&mut self, title: impl Into<String>) {
        self.title = normalize_optional_text(title.into());
    }

    /// Clear document title.
    pub fn clear_title(&mut self) {
        self.title = None;
    }

    /// Document subject (`dc:subject`).
    pub fn subject(&self) -> Option<&str> {
        self.subject.as_deref()
    }

    /// Set document subject.
    pub fn set_subject(&mut self, subject: impl Into<String>) {
        self.subject = normalize_optional_text(subject.into());
    }

    /// Clear document subject.
    pub fn clear_subject(&mut self) {
        self.subject = None;
    }

    /// Document creator/author (`dc:creator`).
    pub fn creator(&self) -> Option<&str> {
        self.creator.as_deref()
    }

    /// Set document creator.
    pub fn set_creator(&mut self, creator: impl Into<String>) {
        self.creator = normalize_optional_text(creator.into());
    }

    /// Clear document creator.
    pub fn clear_creator(&mut self) {
        self.creator = None;
    }

    /// Document description (`dc:description`).
    pub fn description(&self) -> Option<&str> {
        self.description.as_deref()
    }

    /// Set document description.
    pub fn set_description(&mut self, description: impl Into<String>) {
        self.description = normalize_optional_text(description.into());
    }

    /// Clear document description.
    pub fn clear_description(&mut self) {
        self.description = None;
    }

    /// Document keywords (`cp:keywords`).
    pub fn keywords(&self) -> Option<&str> {
        self.keywords.as_deref()
    }

    /// Set document keywords.
    pub fn set_keywords(&mut self, keywords: impl Into<String>) {
        self.keywords = normalize_optional_text(keywords.into());
    }

    /// Clear document keywords.
    pub fn clear_keywords(&mut self) {
        self.keywords = None;
    }

    /// Last person who modified the document (`cp:lastModifiedBy`).
    pub fn last_modified_by(&self) -> Option<&str> {
        self.last_modified_by.as_deref()
    }

    /// Set last modified by.
    pub fn set_last_modified_by(&mut self, last_modified_by: impl Into<String>) {
        self.last_modified_by = normalize_optional_text(last_modified_by.into());
    }

    /// Clear last modified by.
    pub fn clear_last_modified_by(&mut self) {
        self.last_modified_by = None;
    }

    /// Document creation date (`dcterms:created`), typically ISO 8601.
    pub fn created(&self) -> Option<&str> {
        self.created.as_deref()
    }

    /// Set creation date.
    pub fn set_created(&mut self, created: impl Into<String>) {
        self.created = normalize_optional_text(created.into());
    }

    /// Clear creation date.
    pub fn clear_created(&mut self) {
        self.created = None;
    }

    /// Document modification date (`dcterms:modified`), typically ISO 8601.
    pub fn modified(&self) -> Option<&str> {
        self.modified.as_deref()
    }

    /// Set modification date.
    pub fn set_modified(&mut self, modified: impl Into<String>) {
        self.modified = normalize_optional_text(modified.into());
    }

    /// Clear modification date.
    pub fn clear_modified(&mut self) {
        self.modified = None;
    }

    /// Whether all properties are empty/unset.
    pub fn is_empty(&self) -> bool {
        self.title.is_none()
            && self.subject.is_none()
            && self.creator.is_none()
            && self.description.is_none()
            && self.keywords.is_none()
            && self.last_modified_by.is_none()
            && self.created.is_none()
            && self.modified.is_none()
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

#[cfg(test)]
mod tests {
    use super::DocumentProperties;

    #[test]
    fn document_properties_start_empty() {
        let props = DocumentProperties::new();
        assert!(props.is_empty());
        assert_eq!(props.title(), None);
        assert_eq!(props.subject(), None);
        assert_eq!(props.creator(), None);
        assert_eq!(props.description(), None);
        assert_eq!(props.keywords(), None);
        assert_eq!(props.last_modified_by(), None);
        assert_eq!(props.created(), None);
        assert_eq!(props.modified(), None);
    }

    #[test]
    fn document_properties_can_be_set_and_cleared() {
        let mut props = DocumentProperties::new();

        props.set_title("My Document");
        props.set_subject("Report");
        props.set_creator("John Doe");
        props.set_description("A quarterly report");
        props.set_keywords("finance, quarterly");
        props.set_last_modified_by("Jane Smith");
        props.set_created("2024-01-15T10:00:00Z");
        props.set_modified("2024-03-20T14:30:00Z");

        assert!(!props.is_empty());
        assert_eq!(props.title(), Some("My Document"));
        assert_eq!(props.subject(), Some("Report"));
        assert_eq!(props.creator(), Some("John Doe"));
        assert_eq!(props.description(), Some("A quarterly report"));
        assert_eq!(props.keywords(), Some("finance, quarterly"));
        assert_eq!(props.last_modified_by(), Some("Jane Smith"));
        assert_eq!(props.created(), Some("2024-01-15T10:00:00Z"));
        assert_eq!(props.modified(), Some("2024-03-20T14:30:00Z"));

        props.clear_title();
        props.clear_subject();
        props.clear_creator();
        props.clear_description();
        props.clear_keywords();
        props.clear_last_modified_by();
        props.clear_created();
        props.clear_modified();

        assert!(props.is_empty());
    }

    #[test]
    fn whitespace_only_values_are_normalized_to_none() {
        let mut props = DocumentProperties::new();
        props.set_title("   ");
        props.set_creator("");
        assert_eq!(props.title(), None);
        assert_eq!(props.creator(), None);
        assert!(props.is_empty());
    }
}
