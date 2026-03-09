use std::io::Cursor;

use offidized_opc::RawXmlNode;
use quick_xml::events::{BytesDecl, BytesEnd, BytesStart, BytesText, Event};
use quick_xml::{Reader, Writer};

use crate::error::Result;

/// Presentation document properties from `docProps/core.xml` and `docProps/app.xml`.
///
/// Based on ShapeCrawler's `IPresentationProperties` interface. Properties are split between:
/// - Core properties (`docProps/core.xml`): author, title, subject, keywords, category,
///   comments, last_modified_by, created, modified, revision
/// - Extended properties (`docProps/app.xml`): company, manager, hyperlink_base, app_version
#[derive(Debug, Clone, PartialEq, Eq, Default)]
pub struct PresentationProperties {
    // ── Core properties (docProps/core.xml) ──
    /// Document author/creator (`dc:creator`).
    author: Option<String>,
    /// Document title (`dc:title`).
    title: Option<String>,
    /// Document subject (`dc:subject`).
    subject: Option<String>,
    /// Document keywords/tags (`cp:keywords`).
    keywords: Option<String>,
    /// Document category (`cp:category`).
    category: Option<String>,
    /// Document comments/description (`dc:description`).
    comments: Option<String>,
    /// Last person who modified the document (`cp:lastModifiedBy`).
    last_modified_by: Option<String>,
    /// Document creation date (`dcterms:created`), typically ISO 8601.
    created: Option<String>,
    /// Document modification date (`dcterms:modified`), typically ISO 8601.
    modified: Option<String>,
    /// Revision number (`cp:revision`).
    revision: Option<u32>,

    // ── Extended properties (docProps/app.xml) ──
    /// Company name.
    company: Option<String>,
    /// Manager name.
    manager: Option<String>,
    /// Hyperlink base URI.
    hyperlink_base: Option<String>,
    /// Application version that created/modified the document.
    app_version: Option<String>,

    // ── Unknown properties for roundtrip fidelity ──
    /// Unknown elements from `docProps/core.xml`, preserved for roundtrip.
    unknown_core_properties: Vec<RawXmlNode>,
    /// Unknown elements from `docProps/app.xml`, preserved for roundtrip.
    unknown_app_properties: Vec<RawXmlNode>,
}

impl PresentationProperties {
    /// Create empty presentation properties.
    pub fn new() -> Self {
        Self::default()
    }

    // ── Core properties: author ──

    /// Document author/creator (`dc:creator`).
    pub fn author(&self) -> Option<&str> {
        self.author.as_deref()
    }

    /// Set document author.
    pub fn set_author(&mut self, author: impl Into<String>) {
        self.author = normalize_optional_text(author.into());
    }

    /// Clear document author.
    pub fn clear_author(&mut self) {
        self.author = None;
    }

    // ── Core properties: title ──

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

    // ── Core properties: subject ──

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

    // ── Core properties: keywords ──

    /// Document keywords/tags (`cp:keywords`).
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

    // ── Core properties: category ──

    /// Document category (`cp:category`).
    pub fn category(&self) -> Option<&str> {
        self.category.as_deref()
    }

    /// Set document category.
    pub fn set_category(&mut self, category: impl Into<String>) {
        self.category = normalize_optional_text(category.into());
    }

    /// Clear document category.
    pub fn clear_category(&mut self) {
        self.category = None;
    }

    // ── Core properties: comments ──

    /// Document comments/description (`dc:description`).
    pub fn comments(&self) -> Option<&str> {
        self.comments.as_deref()
    }

    /// Set document comments.
    pub fn set_comments(&mut self, comments: impl Into<String>) {
        self.comments = normalize_optional_text(comments.into());
    }

    /// Clear document comments.
    pub fn clear_comments(&mut self) {
        self.comments = None;
    }

    // ── Core properties: last_modified_by ──

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

    // ── Core properties: created ──

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

    // ── Core properties: modified ──

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

    // ── Core properties: revision ──

    /// Revision number (`cp:revision`).
    pub fn revision(&self) -> Option<u32> {
        self.revision
    }

    /// Set revision number.
    pub fn set_revision(&mut self, revision: u32) {
        self.revision = Some(revision);
    }

    /// Clear revision number.
    pub fn clear_revision(&mut self) {
        self.revision = None;
    }

    // ── Extended properties: company ──

    /// Company name (from `docProps/app.xml`).
    pub fn company(&self) -> Option<&str> {
        self.company.as_deref()
    }

    /// Set company name.
    pub fn set_company(&mut self, company: impl Into<String>) {
        self.company = normalize_optional_text(company.into());
    }

    /// Clear company name.
    pub fn clear_company(&mut self) {
        self.company = None;
    }

    // ── Extended properties: manager ──

    /// Manager name (from `docProps/app.xml`).
    pub fn manager(&self) -> Option<&str> {
        self.manager.as_deref()
    }

    /// Set manager name.
    pub fn set_manager(&mut self, manager: impl Into<String>) {
        self.manager = normalize_optional_text(manager.into());
    }

    /// Clear manager name.
    pub fn clear_manager(&mut self) {
        self.manager = None;
    }

    // ── Extended properties: hyperlink_base ──

    /// Hyperlink base URI (from `docProps/app.xml`).
    pub fn hyperlink_base(&self) -> Option<&str> {
        self.hyperlink_base.as_deref()
    }

    /// Set hyperlink base.
    pub fn set_hyperlink_base(&mut self, hyperlink_base: impl Into<String>) {
        self.hyperlink_base = normalize_optional_text(hyperlink_base.into());
    }

    /// Clear hyperlink base.
    pub fn clear_hyperlink_base(&mut self) {
        self.hyperlink_base = None;
    }

    // ── Extended properties: app_version ──

    /// Application version (from `docProps/app.xml`).
    pub fn app_version(&self) -> Option<&str> {
        self.app_version.as_deref()
    }

    /// Set application version.
    pub fn set_app_version(&mut self, app_version: impl Into<String>) {
        self.app_version = normalize_optional_text(app_version.into());
    }

    /// Clear application version.
    pub fn clear_app_version(&mut self) {
        self.app_version = None;
    }

    // ── Helpers ──

    /// Whether all properties are empty/unset.
    pub fn is_empty(&self) -> bool {
        self.author.is_none()
            && self.title.is_none()
            && self.subject.is_none()
            && self.keywords.is_none()
            && self.category.is_none()
            && self.comments.is_none()
            && self.last_modified_by.is_none()
            && self.created.is_none()
            && self.modified.is_none()
            && self.revision.is_none()
            && self.company.is_none()
            && self.manager.is_none()
            && self.hyperlink_base.is_none()
            && self.app_version.is_none()
            && self.unknown_core_properties.is_empty()
            && self.unknown_app_properties.is_empty()
    }

    // ── Parsing ──

    /// Parse core properties from `docProps/core.xml`.
    pub fn parse_core_xml(&mut self, xml: &[u8]) -> Result<()> {
        let mut reader = Reader::from_reader(Cursor::new(xml));
        reader.config_mut().trim_text(false);

        let mut current_element: Option<String> = None;
        let mut current_text = String::new();
        let mut buffer = Vec::new();
        let mut in_root = false;

        loop {
            match reader.read_event_into(&mut buffer)? {
                Event::Start(ref event) => {
                    let name_bytes = event.name();
                    let local = local_name(name_bytes.as_ref());
                    if local == b"coreProperties" {
                        in_root = true;
                    } else if in_root {
                        match local {
                            b"title" | b"subject" | b"creator" | b"description" | b"keywords"
                            | b"category" | b"lastModifiedBy" | b"created" | b"modified"
                            | b"revision" => {
                                current_element = Some(String::from_utf8_lossy(local).into_owned());
                                current_text.clear();
                            }
                            _ => {
                                // Unknown element — capture for roundtrip.
                                self.unknown_core_properties
                                    .push(RawXmlNode::read_element(&mut reader, event)?);
                            }
                        }
                    }
                }
                Event::Text(ref event) => {
                    if current_element.is_some() {
                        if let Ok(text) = event.unescape() {
                            current_text.push_str(text.as_ref());
                        }
                    }
                }
                Event::End(ref event) => {
                    let name_bytes = event.name();
                    let local = local_name(name_bytes.as_ref());
                    if local == b"coreProperties" {
                        in_root = false;
                    } else if let Some(ref elem) = current_element {
                        let local_str = std::str::from_utf8(local).unwrap_or("");
                        if local_str == elem.as_str() {
                            match elem.as_str() {
                                "title" => self.set_title(current_text.clone()),
                                "subject" => self.set_subject(current_text.clone()),
                                "creator" => self.set_author(current_text.clone()),
                                "description" => self.set_comments(current_text.clone()),
                                "keywords" => self.set_keywords(current_text.clone()),
                                "category" => self.set_category(current_text.clone()),
                                "lastModifiedBy" => {
                                    self.set_last_modified_by(current_text.clone());
                                }
                                "created" => self.set_created(current_text.clone()),
                                "modified" => self.set_modified(current_text.clone()),
                                "revision" => {
                                    if let Ok(rev) = current_text.trim().parse::<u32>() {
                                        self.set_revision(rev);
                                    }
                                }
                                _ => {}
                            }
                            current_element = None;
                            current_text.clear();
                        }
                    }
                }
                Event::Eof => break,
                _ => {}
            }
            buffer.clear();
        }

        Ok(())
    }

    /// Parse extended properties from `docProps/app.xml`.
    pub fn parse_app_xml(&mut self, xml: &[u8]) -> Result<()> {
        let mut reader = Reader::from_reader(Cursor::new(xml));
        reader.config_mut().trim_text(false);

        let mut current_element: Option<String> = None;
        let mut current_text = String::new();
        let mut buffer = Vec::new();
        let mut in_root = false;

        loop {
            match reader.read_event_into(&mut buffer)? {
                Event::Start(ref event) => {
                    let name_bytes = event.name();
                    let local = local_name(name_bytes.as_ref());
                    if local == b"Properties" {
                        in_root = true;
                    } else if in_root {
                        match local {
                            b"Company" | b"Manager" | b"HyperlinkBase" | b"AppVersion" => {
                                current_element = Some(String::from_utf8_lossy(local).into_owned());
                                current_text.clear();
                            }
                            _ => {
                                // Unknown element — capture for roundtrip.
                                self.unknown_app_properties
                                    .push(RawXmlNode::read_element(&mut reader, event)?);
                            }
                        }
                    }
                }
                Event::Text(ref event) => {
                    if current_element.is_some() {
                        if let Ok(text) = event.unescape() {
                            current_text.push_str(text.as_ref());
                        }
                    }
                }
                Event::End(ref event) => {
                    let name_bytes = event.name();
                    let local = local_name(name_bytes.as_ref());
                    if local == b"Properties" {
                        in_root = false;
                    } else if let Some(ref elem) = current_element {
                        let local_str = std::str::from_utf8(local).unwrap_or("");
                        if local_str == elem.as_str() {
                            match elem.as_str() {
                                "Company" => self.set_company(current_text.clone()),
                                "Manager" => self.set_manager(current_text.clone()),
                                "HyperlinkBase" => {
                                    self.set_hyperlink_base(current_text.clone());
                                }
                                "AppVersion" => self.set_app_version(current_text.clone()),
                                _ => {}
                            }
                            current_element = None;
                            current_text.clear();
                        }
                    }
                }
                Event::Eof => break,
                _ => {}
            }
            buffer.clear();
        }

        Ok(())
    }

    /// Write core properties to `docProps/core.xml` format.
    pub fn write_core_xml(&self) -> Result<Vec<u8>> {
        let mut writer = Writer::new(Cursor::new(Vec::new()));
        writer.write_event(Event::Decl(BytesDecl::new("1.0", Some("UTF-8"), None)))?;

        let mut root = BytesStart::new("cp:coreProperties");
        root.push_attribute((
            "xmlns:cp",
            "http://schemas.openxmlformats.org/package/2006/metadata/core-properties",
        ));
        root.push_attribute(("xmlns:dc", "http://purl.org/dc/elements/1.1/"));
        root.push_attribute(("xmlns:dcterms", "http://purl.org/dc/terms/"));
        root.push_attribute(("xmlns:xsi", "http://www.w3.org/2001/XMLSchema-instance"));
        writer.write_event(Event::Start(root))?;

        if let Some(ref title) = self.title {
            writer.write_event(Event::Start(BytesStart::new("dc:title")))?;
            writer.write_event(Event::Text(BytesText::new(title)))?;
            writer.write_event(Event::End(BytesEnd::new("dc:title")))?;
        }

        if let Some(ref subject) = self.subject {
            writer.write_event(Event::Start(BytesStart::new("dc:subject")))?;
            writer.write_event(Event::Text(BytesText::new(subject)))?;
            writer.write_event(Event::End(BytesEnd::new("dc:subject")))?;
        }

        if let Some(ref author) = self.author {
            writer.write_event(Event::Start(BytesStart::new("dc:creator")))?;
            writer.write_event(Event::Text(BytesText::new(author)))?;
            writer.write_event(Event::End(BytesEnd::new("dc:creator")))?;
        }

        if let Some(ref keywords) = self.keywords {
            writer.write_event(Event::Start(BytesStart::new("cp:keywords")))?;
            writer.write_event(Event::Text(BytesText::new(keywords)))?;
            writer.write_event(Event::End(BytesEnd::new("cp:keywords")))?;
        }

        if let Some(ref comments) = self.comments {
            writer.write_event(Event::Start(BytesStart::new("dc:description")))?;
            writer.write_event(Event::Text(BytesText::new(comments)))?;
            writer.write_event(Event::End(BytesEnd::new("dc:description")))?;
        }

        if let Some(ref last_modified_by) = self.last_modified_by {
            writer.write_event(Event::Start(BytesStart::new("cp:lastModifiedBy")))?;
            writer.write_event(Event::Text(BytesText::new(last_modified_by)))?;
            writer.write_event(Event::End(BytesEnd::new("cp:lastModifiedBy")))?;
        }

        if let Some(ref revision) = self.revision {
            writer.write_event(Event::Start(BytesStart::new("cp:revision")))?;
            writer.write_event(Event::Text(BytesText::new(&revision.to_string())))?;
            writer.write_event(Event::End(BytesEnd::new("cp:revision")))?;
        }

        if let Some(ref category) = self.category {
            writer.write_event(Event::Start(BytesStart::new("cp:category")))?;
            writer.write_event(Event::Text(BytesText::new(category)))?;
            writer.write_event(Event::End(BytesEnd::new("cp:category")))?;
        }

        if let Some(ref created) = self.created {
            let mut created_elem = BytesStart::new("dcterms:created");
            created_elem.push_attribute(("xsi:type", "dcterms:W3CDTF"));
            writer.write_event(Event::Start(created_elem))?;
            writer.write_event(Event::Text(BytesText::new(created)))?;
            writer.write_event(Event::End(BytesEnd::new("dcterms:created")))?;
        }

        if let Some(ref modified) = self.modified {
            let mut modified_elem = BytesStart::new("dcterms:modified");
            modified_elem.push_attribute(("xsi:type", "dcterms:W3CDTF"));
            writer.write_event(Event::Start(modified_elem))?;
            writer.write_event(Event::Text(BytesText::new(modified)))?;
            writer.write_event(Event::End(BytesEnd::new("dcterms:modified")))?;
        }

        // Replay unknown core properties for roundtrip fidelity.
        for node in &self.unknown_core_properties {
            node.write_to(&mut writer)?;
        }

        writer.write_event(Event::End(BytesEnd::new("cp:coreProperties")))?;

        Ok(writer.into_inner().into_inner())
    }

    /// Write extended properties to `docProps/app.xml` format.
    pub fn write_app_xml(&self) -> Result<Vec<u8>> {
        let mut writer = Writer::new(Cursor::new(Vec::new()));
        writer.write_event(Event::Decl(BytesDecl::new("1.0", Some("UTF-8"), None)))?;

        let mut root = BytesStart::new("Properties");
        root.push_attribute((
            "xmlns",
            "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties",
        ));
        root.push_attribute((
            "xmlns:vt",
            "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes",
        ));
        writer.write_event(Event::Start(root))?;

        if let Some(ref company) = self.company {
            writer.write_event(Event::Start(BytesStart::new("Company")))?;
            writer.write_event(Event::Text(BytesText::new(company)))?;
            writer.write_event(Event::End(BytesEnd::new("Company")))?;
        }

        if let Some(ref manager) = self.manager {
            writer.write_event(Event::Start(BytesStart::new("Manager")))?;
            writer.write_event(Event::Text(BytesText::new(manager)))?;
            writer.write_event(Event::End(BytesEnd::new("Manager")))?;
        }

        if let Some(ref hyperlink_base) = self.hyperlink_base {
            writer.write_event(Event::Start(BytesStart::new("HyperlinkBase")))?;
            writer.write_event(Event::Text(BytesText::new(hyperlink_base)))?;
            writer.write_event(Event::End(BytesEnd::new("HyperlinkBase")))?;
        }

        if let Some(ref app_version) = self.app_version {
            writer.write_event(Event::Start(BytesStart::new("AppVersion")))?;
            writer.write_event(Event::Text(BytesText::new(app_version)))?;
            writer.write_event(Event::End(BytesEnd::new("AppVersion")))?;
        }

        // Replay unknown app properties for roundtrip fidelity.
        for node in &self.unknown_app_properties {
            node.write_to(&mut writer)?;
        }

        writer.write_event(Event::End(BytesEnd::new("Properties")))?;

        Ok(writer.into_inner().into_inner())
    }
}

/// Normalize optional text by trimming and converting empty strings to None.
fn normalize_optional_text(value: String) -> Option<String> {
    let trimmed = value.trim();
    if trimmed.is_empty() {
        None
    } else {
        Some(trimmed.to_string())
    }
}

/// Extract the local name from a qualified XML name (e.g., "cp:title" -> "title").
fn local_name(qualified: &[u8]) -> &[u8] {
    if let Some(colon_pos) = qualified.iter().position(|&b| b == b':') {
        &qualified[colon_pos + 1..]
    } else {
        qualified
    }
}

#[cfg(test)]
mod tests {
    use super::PresentationProperties;

    #[test]
    fn presentation_properties_start_empty() {
        let props = PresentationProperties::new();
        assert!(props.is_empty());
        assert_eq!(props.author(), None);
        assert_eq!(props.title(), None);
        assert_eq!(props.subject(), None);
        assert_eq!(props.keywords(), None);
        assert_eq!(props.category(), None);
        assert_eq!(props.comments(), None);
        assert_eq!(props.last_modified_by(), None);
        assert_eq!(props.created(), None);
        assert_eq!(props.modified(), None);
        assert_eq!(props.revision(), None);
        assert_eq!(props.company(), None);
        assert_eq!(props.manager(), None);
        assert_eq!(props.hyperlink_base(), None);
        assert_eq!(props.app_version(), None);
    }

    #[test]
    fn core_properties_can_be_set_and_cleared() {
        let mut props = PresentationProperties::new();

        props.set_author("Jane Smith");
        props.set_title("Q4 Results");
        props.set_subject("Quarterly Report");
        props.set_keywords("finance, quarterly");
        props.set_category("Reports");
        props.set_comments("Final version for board review");
        props.set_last_modified_by("John Doe");
        props.set_created("2024-01-15T10:00:00Z");
        props.set_modified("2024-03-20T14:30:00Z");
        props.set_revision(5);

        assert!(!props.is_empty());
        assert_eq!(props.author(), Some("Jane Smith"));
        assert_eq!(props.title(), Some("Q4 Results"));
        assert_eq!(props.subject(), Some("Quarterly Report"));
        assert_eq!(props.keywords(), Some("finance, quarterly"));
        assert_eq!(props.category(), Some("Reports"));
        assert_eq!(props.comments(), Some("Final version for board review"));
        assert_eq!(props.last_modified_by(), Some("John Doe"));
        assert_eq!(props.created(), Some("2024-01-15T10:00:00Z"));
        assert_eq!(props.modified(), Some("2024-03-20T14:30:00Z"));
        assert_eq!(props.revision(), Some(5));

        props.clear_author();
        props.clear_title();
        props.clear_subject();
        props.clear_keywords();
        props.clear_category();
        props.clear_comments();
        props.clear_last_modified_by();
        props.clear_created();
        props.clear_modified();
        props.clear_revision();

        assert!(props.is_empty());
    }

    #[test]
    fn extended_properties_can_be_set_and_cleared() {
        let mut props = PresentationProperties::new();

        props.set_company("Acme Corp");
        props.set_manager("Alice Johnson");
        props.set_hyperlink_base("https://example.com/docs/");
        props.set_app_version("16.0.0");

        assert!(!props.is_empty());
        assert_eq!(props.company(), Some("Acme Corp"));
        assert_eq!(props.manager(), Some("Alice Johnson"));
        assert_eq!(props.hyperlink_base(), Some("https://example.com/docs/"));
        assert_eq!(props.app_version(), Some("16.0.0"));

        props.clear_company();
        props.clear_manager();
        props.clear_hyperlink_base();
        props.clear_app_version();

        assert!(props.is_empty());
    }

    #[test]
    fn whitespace_only_values_are_normalized_to_none() {
        let mut props = PresentationProperties::new();
        props.set_title("   ");
        props.set_author("");
        props.set_company("\t\n  ");
        assert_eq!(props.title(), None);
        assert_eq!(props.author(), None);
        assert_eq!(props.company(), None);
        assert!(props.is_empty());
    }

    #[test]
    fn values_are_trimmed() {
        let mut props = PresentationProperties::new();
        props.set_title("  My Title  ");
        props.set_author("\tJohn Doe\n");
        assert_eq!(props.title(), Some("My Title"));
        assert_eq!(props.author(), Some("John Doe"));
    }

    #[test]
    fn parse_core_xml_reads_all_properties() {
        let xml = br#"<?xml version="1.0" encoding="UTF-8"?>
        <cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
            xmlns:dc="http://purl.org/dc/elements/1.1/"
            xmlns:dcterms="http://purl.org/dc/terms/"
            xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
          <dc:title>Q4 Results</dc:title>
          <dc:subject>Quarterly Report</dc:subject>
          <dc:creator>Jane Smith</dc:creator>
          <cp:keywords>finance, quarterly</cp:keywords>
          <dc:description>Final version for board review</dc:description>
          <cp:lastModifiedBy>John Doe</cp:lastModifiedBy>
          <dcterms:created xsi:type="dcterms:W3CDTF">2024-01-15T10:00:00Z</dcterms:created>
          <dcterms:modified xsi:type="dcterms:W3CDTF">2024-03-20T14:30:00Z</dcterms:modified>
          <cp:category>Reports</cp:category>
          <cp:revision>5</cp:revision>
        </cp:coreProperties>"#;

        let mut props = PresentationProperties::new();
        props.parse_core_xml(xml).expect("parse core properties");

        assert_eq!(props.title(), Some("Q4 Results"));
        assert_eq!(props.subject(), Some("Quarterly Report"));
        assert_eq!(props.author(), Some("Jane Smith"));
        assert_eq!(props.keywords(), Some("finance, quarterly"));
        assert_eq!(props.comments(), Some("Final version for board review"));
        assert_eq!(props.last_modified_by(), Some("John Doe"));
        assert_eq!(props.created(), Some("2024-01-15T10:00:00Z"));
        assert_eq!(props.modified(), Some("2024-03-20T14:30:00Z"));
        assert_eq!(props.category(), Some("Reports"));
        assert_eq!(props.revision(), Some(5));
    }

    #[test]
    fn parse_app_xml_reads_all_properties() {
        let xml = br#"<?xml version="1.0" encoding="UTF-8"?>
        <Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"
            xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
          <Company>Acme Corp</Company>
          <Manager>Alice Johnson</Manager>
          <HyperlinkBase>https://example.com/docs/</HyperlinkBase>
          <AppVersion>16.0.0</AppVersion>
        </Properties>"#;

        let mut props = PresentationProperties::new();
        props.parse_app_xml(xml).expect("parse app properties");

        assert_eq!(props.company(), Some("Acme Corp"));
        assert_eq!(props.manager(), Some("Alice Johnson"));
        assert_eq!(props.hyperlink_base(), Some("https://example.com/docs/"));
        assert_eq!(props.app_version(), Some("16.0.0"));
    }

    #[test]
    fn write_core_xml_produces_valid_xml() {
        let mut props = PresentationProperties::new();
        props.set_title("Q4 Results");
        props.set_author("Jane Smith");
        props.set_subject("Quarterly Report");
        props.set_keywords("finance, quarterly");
        props.set_category("Reports");
        props.set_comments("Final version");
        props.set_last_modified_by("John Doe");
        props.set_created("2024-01-15T10:00:00Z");
        props.set_modified("2024-03-20T14:30:00Z");
        props.set_revision(5);

        let xml = props.write_core_xml().expect("write core xml");
        let xml_str = String::from_utf8_lossy(&xml);

        assert!(xml_str.contains("<dc:title>Q4 Results</dc:title>"));
        assert!(xml_str.contains("<dc:creator>Jane Smith</dc:creator>"));
        assert!(xml_str.contains("<dc:subject>Quarterly Report</dc:subject>"));
        assert!(xml_str.contains("<cp:keywords>finance, quarterly</cp:keywords>"));
        assert!(xml_str.contains("<cp:category>Reports</cp:category>"));
        assert!(xml_str.contains("<dc:description>Final version</dc:description>"));
        assert!(xml_str.contains("<cp:lastModifiedBy>John Doe</cp:lastModifiedBy>"));
        assert!(xml_str.contains("<cp:revision>5</cp:revision>"));
        assert!(xml_str.contains("<dcterms:created"));
        assert!(xml_str.contains("2024-01-15T10:00:00Z"));
        assert!(xml_str.contains("<dcterms:modified"));
        assert!(xml_str.contains("2024-03-20T14:30:00Z"));
    }

    #[test]
    fn write_app_xml_produces_valid_xml() {
        let mut props = PresentationProperties::new();
        props.set_company("Acme Corp");
        props.set_manager("Alice Johnson");
        props.set_hyperlink_base("https://example.com/docs/");
        props.set_app_version("16.0.0");

        let xml = props.write_app_xml().expect("write app xml");
        let xml_str = String::from_utf8_lossy(&xml);

        assert!(xml_str.contains("<Company>Acme Corp</Company>"));
        assert!(xml_str.contains("<Manager>Alice Johnson</Manager>"));
        assert!(xml_str.contains("<HyperlinkBase>https://example.com/docs/</HyperlinkBase>"));
        assert!(xml_str.contains("<AppVersion>16.0.0</AppVersion>"));
    }

    #[test]
    fn roundtrip_core_xml_preserves_all_properties() {
        let mut props1 = PresentationProperties::new();
        props1.set_title("Test Title");
        props1.set_author("Test Author");
        props1.set_subject("Test Subject");
        props1.set_keywords("test, keywords");
        props1.set_category("Test Category");
        props1.set_comments("Test Comments");
        props1.set_last_modified_by("Test User");
        props1.set_created("2024-01-01T00:00:00Z");
        props1.set_modified("2024-02-01T00:00:00Z");
        props1.set_revision(3);

        let xml = props1.write_core_xml().expect("write xml");
        let mut props2 = PresentationProperties::new();
        props2.parse_core_xml(&xml).expect("parse xml");

        assert_eq!(props1.title(), props2.title());
        assert_eq!(props1.author(), props2.author());
        assert_eq!(props1.subject(), props2.subject());
        assert_eq!(props1.keywords(), props2.keywords());
        assert_eq!(props1.category(), props2.category());
        assert_eq!(props1.comments(), props2.comments());
        assert_eq!(props1.last_modified_by(), props2.last_modified_by());
        assert_eq!(props1.created(), props2.created());
        assert_eq!(props1.modified(), props2.modified());
        assert_eq!(props1.revision(), props2.revision());
    }

    #[test]
    fn roundtrip_app_xml_preserves_all_properties() {
        let mut props1 = PresentationProperties::new();
        props1.set_company("Test Company");
        props1.set_manager("Test Manager");
        props1.set_hyperlink_base("https://test.com/");
        props1.set_app_version("1.2.3");

        let xml = props1.write_app_xml().expect("write xml");
        let mut props2 = PresentationProperties::new();
        props2.parse_app_xml(&xml).expect("parse xml");

        assert_eq!(props1.company(), props2.company());
        assert_eq!(props1.manager(), props2.manager());
        assert_eq!(props1.hyperlink_base(), props2.hyperlink_base());
        assert_eq!(props1.app_version(), props2.app_version());
    }
}
