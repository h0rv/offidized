//! Package properties — OPC core, extended, and custom properties.
//!
//! Maps to `IPackageProperties` in the Open XML SDK. Provides typed access
//! to `docProps/core.xml` and `docProps/app.xml` metadata.

use quick_xml::events::{BytesDecl, BytesEnd, BytesRef, BytesStart, BytesText, Event};
use quick_xml::{Reader, Writer};
use std::io::{BufRead, Write};

use crate::error::Result;
use crate::raw::RawXmlNode;

/// Core package properties from `docProps/core.xml`.
///
/// These correspond to the Dublin Core and OPC core properties defined
/// in ECMA-376 Part 2. Maps 1:1 to `IPackageProperties` in the Open XML SDK.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct CoreProperties {
    /// Document title.
    pub title: Option<String>,
    /// Topic/subject of the contents.
    pub subject: Option<String>,
    /// Primary creator (author).
    pub creator: Option<String>,
    /// Keywords for searching and indexing.
    pub keywords: Option<String>,
    /// Description or abstract.
    pub description: Option<String>,
    /// User who last modified the document.
    pub last_modified_by: Option<String>,
    /// Revision number.
    pub revision: Option<String>,
    /// Date/time of creation (ISO 8601 string).
    pub created: Option<String>,
    /// Date/time of last modification (ISO 8601 string).
    pub modified: Option<String>,
    /// Date/time of last printing (ISO 8601 string).
    pub last_printed: Option<String>,
    /// Category.
    pub category: Option<String>,
    /// Content type (not MIME — e.g., "Whitepaper", "Exam").
    pub content_type: Option<String>,
    /// Content status (e.g., "Draft", "Final").
    pub content_status: Option<String>,
    /// Unique identifier.
    pub identifier: Option<String>,
    /// Primary language (RFC 3066 tag, e.g., "en-US").
    pub language: Option<String>,
    /// Version number.
    pub version: Option<String>,
    /// Original raw XML bytes for no-op roundtrip.
    raw_xml: Option<Vec<u8>>,
    /// Whether any property has been modified since loading.
    dirty: bool,
}

impl CoreProperties {
    /// Create empty core properties.
    pub fn new() -> Self {
        Self::default()
    }

    /// Parse from `docProps/core.xml` bytes, preserving original for roundtrip.
    pub fn from_xml_bytes(bytes: Vec<u8>) -> Result<Self> {
        let mut props = Self::from_xml(std::io::Cursor::new(bytes.as_slice()))?;
        props.raw_xml = Some(bytes);
        props.dirty = false;
        Ok(props)
    }

    /// Parse from an XML reader.
    pub fn from_xml<R: BufRead>(reader: R) -> Result<Self> {
        let mut xml = Reader::from_reader(reader);
        xml.config_mut().trim_text(true);

        let mut props = CoreProperties::new();
        let mut buf = Vec::new();
        let mut current_element: Option<String> = None;
        let mut current_text = String::new();

        loop {
            match xml.read_event_into(&mut buf) {
                Ok(Event::Start(ref e)) => {
                    let name_bytes = e.name();
                    let local = local_name(name_bytes.as_ref());
                    current_element = Some(String::from_utf8_lossy(local).into_owned());
                    current_text.clear();
                }
                Ok(Event::Text(ref e)) => {
                    if current_element.is_some() {
                        current_text.push_str(&decode_text_event(e)?);
                    }
                }
                Ok(Event::GeneralRef(ref e)) => {
                    if current_element.is_some() {
                        current_text.push_str(&decode_general_ref(e)?);
                    }
                }
                Ok(Event::End(_)) => {
                    if let Some(elem) = current_element.take() {
                        if !current_text.is_empty() {
                            set_core_property(&mut props, &elem, std::mem::take(&mut current_text));
                        } else {
                            current_text.clear();
                        }
                    }
                    current_element = None;
                }
                Ok(Event::Eof) => break,
                Err(e) => return Err(e.into()),
                _ => {}
            }
            buf.clear();
        }

        props.dirty = false;
        Ok(props)
    }

    /// Serialize to XML. Uses original bytes on no-op roundtrip.
    pub fn to_xml<W: Write>(&self, writer: W) -> Result<()> {
        if !self.dirty {
            if let Some(ref raw) = self.raw_xml {
                let mut w = writer;
                w.write_all(raw)?;
                return Ok(());
            }
        }

        let mut xml = Writer::new_with_indent(writer, b' ', 2);

        xml.write_event(Event::Decl(BytesDecl::new(
            "1.0",
            Some("UTF-8"),
            Some("yes"),
        )))?;

        let mut root = BytesStart::new("cp:coreProperties");
        root.push_attribute((
            "xmlns:cp",
            "http://schemas.openxmlformats.org/package/2006/metadata/core-properties",
        ));
        root.push_attribute(("xmlns:dc", "http://purl.org/dc/elements/1.1/"));
        root.push_attribute(("xmlns:dcterms", "http://purl.org/dc/terms/"));
        root.push_attribute(("xmlns:dcmitype", "http://purl.org/dc/dcmitype/"));
        root.push_attribute(("xmlns:xsi", "http://www.w3.org/2001/XMLSchema-instance"));
        xml.write_event(Event::Start(root))?;

        write_dc_element(&mut xml, "dc:title", self.title.as_deref())?;
        write_dc_element(&mut xml, "dc:subject", self.subject.as_deref())?;
        write_dc_element(&mut xml, "dc:creator", self.creator.as_deref())?;
        write_element(&mut xml, "cp:keywords", self.keywords.as_deref())?;
        write_dc_element(&mut xml, "dc:description", self.description.as_deref())?;
        write_element(
            &mut xml,
            "cp:lastModifiedBy",
            self.last_modified_by.as_deref(),
        )?;
        write_element(&mut xml, "cp:revision", self.revision.as_deref())?;
        write_dcterms_element(&mut xml, "dcterms:created", self.created.as_deref())?;
        write_dcterms_element(&mut xml, "dcterms:modified", self.modified.as_deref())?;
        write_element(&mut xml, "cp:lastPrinted", self.last_printed.as_deref())?;
        write_element(&mut xml, "cp:category", self.category.as_deref())?;
        write_element(&mut xml, "cp:contentType", self.content_type.as_deref())?;
        write_element(&mut xml, "cp:contentStatus", self.content_status.as_deref())?;
        write_dc_element(&mut xml, "dc:identifier", self.identifier.as_deref())?;
        write_dc_element(&mut xml, "dc:language", self.language.as_deref())?;
        write_element(&mut xml, "cp:version", self.version.as_deref())?;

        xml.write_event(Event::End(BytesEnd::new("cp:coreProperties")))?;

        Ok(())
    }

    /// Whether any property has been set since loading.
    pub fn is_dirty(&self) -> bool {
        self.dirty
    }

    /// Mark properties as clean (not modified).
    pub fn mark_clean(&mut self) {
        self.dirty = false;
    }

    /// Whether any property has a value.
    pub fn is_empty(&self) -> bool {
        self.title.is_none()
            && self.subject.is_none()
            && self.creator.is_none()
            && self.keywords.is_none()
            && self.description.is_none()
            && self.last_modified_by.is_none()
            && self.revision.is_none()
            && self.created.is_none()
            && self.modified.is_none()
            && self.last_printed.is_none()
            && self.category.is_none()
            && self.content_type.is_none()
            && self.content_status.is_none()
            && self.identifier.is_none()
            && self.language.is_none()
            && self.version.is_none()
    }

    /// Set any field and mark dirty.
    pub fn set_title(&mut self, value: impl Into<String>) {
        self.title = Some(value.into());
        self.dirty = true;
    }
    pub fn set_subject(&mut self, value: impl Into<String>) {
        self.subject = Some(value.into());
        self.dirty = true;
    }
    pub fn set_creator(&mut self, value: impl Into<String>) {
        self.creator = Some(value.into());
        self.dirty = true;
    }
    pub fn set_keywords(&mut self, value: impl Into<String>) {
        self.keywords = Some(value.into());
        self.dirty = true;
    }
    pub fn set_description(&mut self, value: impl Into<String>) {
        self.description = Some(value.into());
        self.dirty = true;
    }
    pub fn set_last_modified_by(&mut self, value: impl Into<String>) {
        self.last_modified_by = Some(value.into());
        self.dirty = true;
    }
    pub fn set_revision(&mut self, value: impl Into<String>) {
        self.revision = Some(value.into());
        self.dirty = true;
    }
    pub fn set_created(&mut self, value: impl Into<String>) {
        self.created = Some(value.into());
        self.dirty = true;
    }
    pub fn set_modified(&mut self, value: impl Into<String>) {
        self.modified = Some(value.into());
        self.dirty = true;
    }
    pub fn set_last_printed(&mut self, value: impl Into<String>) {
        self.last_printed = Some(value.into());
        self.dirty = true;
    }
    pub fn set_category(&mut self, value: impl Into<String>) {
        self.category = Some(value.into());
        self.dirty = true;
    }
    pub fn set_content_type(&mut self, value: impl Into<String>) {
        self.content_type = Some(value.into());
        self.dirty = true;
    }
    pub fn set_content_status(&mut self, value: impl Into<String>) {
        self.content_status = Some(value.into());
        self.dirty = true;
    }
    pub fn set_identifier(&mut self, value: impl Into<String>) {
        self.identifier = Some(value.into());
        self.dirty = true;
    }
    pub fn set_language(&mut self, value: impl Into<String>) {
        self.language = Some(value.into());
        self.dirty = true;
    }
    pub fn set_version(&mut self, value: impl Into<String>) {
        self.version = Some(value.into());
        self.dirty = true;
    }
}

/// Extended properties from `docProps/app.xml`.
///
/// Application-specific metadata like total editing time, pages, words, etc.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct ExtendedProperties {
    /// Application that created the document (e.g., "Microsoft Office Word").
    pub application: Option<String>,
    /// Application version (e.g., "16.0000").
    pub app_version: Option<String>,
    /// Document template used.
    pub template: Option<String>,
    /// Manager.
    pub manager: Option<String>,
    /// Company.
    pub company: Option<String>,
    /// Total editing time in minutes.
    pub total_time: Option<String>,
    /// Number of pages.
    pub pages: Option<String>,
    /// Number of words.
    pub words: Option<String>,
    /// Number of characters.
    pub characters: Option<String>,
    /// Number of characters with spaces.
    pub characters_with_spaces: Option<String>,
    /// Number of lines.
    pub lines: Option<String>,
    /// Number of paragraphs.
    pub paragraphs: Option<String>,
    /// Number of slides.
    pub slides: Option<String>,
    /// Number of notes.
    pub notes: Option<String>,
    /// Number of hidden slides.
    pub hidden_slides: Option<String>,
    /// Document security level.
    pub doc_security: Option<String>,
    /// Original raw XML bytes for no-op roundtrip.
    raw_xml: Option<Vec<u8>>,
    /// Whether any property has been modified since loading.
    dirty: bool,
}

impl ExtendedProperties {
    /// Create empty extended properties.
    pub fn new() -> Self {
        Self::default()
    }

    /// Parse from `docProps/app.xml` bytes, preserving original for roundtrip.
    pub fn from_xml_bytes(bytes: Vec<u8>) -> Result<Self> {
        let mut props = Self::from_xml(std::io::Cursor::new(bytes.as_slice()))?;
        props.raw_xml = Some(bytes);
        props.dirty = false;
        Ok(props)
    }

    /// Parse from an XML reader.
    pub fn from_xml<R: BufRead>(reader: R) -> Result<Self> {
        let mut xml = Reader::from_reader(reader);
        xml.config_mut().trim_text(true);

        let mut props = ExtendedProperties::new();
        let mut buf = Vec::new();
        let mut current_element: Option<String> = None;
        let mut current_text = String::new();

        loop {
            match xml.read_event_into(&mut buf) {
                Ok(Event::Start(ref e)) => {
                    let name_bytes = e.name();
                    let local = local_name(name_bytes.as_ref());
                    current_element = Some(String::from_utf8_lossy(local).into_owned());
                    current_text.clear();
                }
                Ok(Event::Text(ref e)) => {
                    if current_element.is_some() {
                        current_text.push_str(&decode_text_event(e)?);
                    }
                }
                Ok(Event::GeneralRef(ref e)) => {
                    if current_element.is_some() {
                        current_text.push_str(&decode_general_ref(e)?);
                    }
                }
                Ok(Event::End(_)) => {
                    if let Some(elem) = current_element.take() {
                        if !current_text.is_empty() {
                            set_extended_property(
                                &mut props,
                                &elem,
                                std::mem::take(&mut current_text),
                            );
                        } else {
                            current_text.clear();
                        }
                    }
                    current_element = None;
                }
                Ok(Event::Eof) => break,
                Err(e) => return Err(e.into()),
                _ => {}
            }
            buf.clear();
        }

        props.dirty = false;
        Ok(props)
    }

    /// Serialize to XML. Uses original bytes on no-op roundtrip.
    pub fn to_xml<W: Write>(&self, writer: W) -> Result<()> {
        if !self.dirty {
            if let Some(ref raw) = self.raw_xml {
                let mut w = writer;
                w.write_all(raw)?;
                return Ok(());
            }
        }

        let mut xml = Writer::new_with_indent(writer, b' ', 2);

        xml.write_event(Event::Decl(BytesDecl::new(
            "1.0",
            Some("UTF-8"),
            Some("yes"),
        )))?;

        let mut root = BytesStart::new("Properties");
        root.push_attribute((
            "xmlns",
            "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties",
        ));
        root.push_attribute((
            "xmlns:vt",
            "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes",
        ));
        xml.write_event(Event::Start(root))?;

        write_element(&mut xml, "Application", self.application.as_deref())?;
        write_element(&mut xml, "AppVersion", self.app_version.as_deref())?;
        write_element(&mut xml, "Template", self.template.as_deref())?;
        write_element(&mut xml, "Manager", self.manager.as_deref())?;
        write_element(&mut xml, "Company", self.company.as_deref())?;
        write_element(&mut xml, "TotalTime", self.total_time.as_deref())?;
        write_element(&mut xml, "Pages", self.pages.as_deref())?;
        write_element(&mut xml, "Words", self.words.as_deref())?;
        write_element(&mut xml, "Characters", self.characters.as_deref())?;
        write_element(
            &mut xml,
            "CharactersWithSpaces",
            self.characters_with_spaces.as_deref(),
        )?;
        write_element(&mut xml, "Lines", self.lines.as_deref())?;
        write_element(&mut xml, "Paragraphs", self.paragraphs.as_deref())?;
        write_element(&mut xml, "Slides", self.slides.as_deref())?;
        write_element(&mut xml, "Notes", self.notes.as_deref())?;
        write_element(&mut xml, "HiddenSlides", self.hidden_slides.as_deref())?;
        write_element(&mut xml, "DocSecurity", self.doc_security.as_deref())?;

        xml.write_event(Event::End(BytesEnd::new("Properties")))?;

        Ok(())
    }

    /// Whether any property has been modified.
    pub fn is_dirty(&self) -> bool {
        self.dirty
    }

    /// Mark as clean.
    pub fn mark_clean(&mut self) {
        self.dirty = false;
    }

    /// Set any field and mark dirty.
    pub fn set_application(&mut self, value: impl Into<String>) {
        self.application = Some(value.into());
        self.dirty = true;
    }
    pub fn set_app_version(&mut self, value: impl Into<String>) {
        self.app_version = Some(value.into());
        self.dirty = true;
    }
    pub fn set_template(&mut self, value: impl Into<String>) {
        self.template = Some(value.into());
        self.dirty = true;
    }
    pub fn set_manager(&mut self, value: impl Into<String>) {
        self.manager = Some(value.into());
        self.dirty = true;
    }
    pub fn set_company(&mut self, value: impl Into<String>) {
        self.company = Some(value.into());
        self.dirty = true;
    }
    pub fn set_total_time(&mut self, value: impl Into<String>) {
        self.total_time = Some(value.into());
        self.dirty = true;
    }
    pub fn set_doc_security(&mut self, value: impl Into<String>) {
        self.doc_security = Some(value.into());
        self.dirty = true;
    }
}

/// Standard FMTID for custom properties.
const CUSTOM_PROPERTY_FMTID: &str = "{D5CDD505-2E9C-101B-9397-08002B2CF9AE}";

/// Namespace for custom properties root element.
const NS_CUSTOM_PROPERTIES: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/custom-properties";

/// Namespace for value types (vt:).
const NS_VT: &str = "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes";

/// Value of a custom property.
#[derive(Debug, Clone, PartialEq)]
pub enum CustomPropertyValue {
    /// A string value (`vt:lpwstr`).
    String(String),
    /// A 32-bit integer value (`vt:i4`).
    Int(i32),
    /// A 64-bit float value (`vt:r8`).
    Float(f64),
    /// A boolean value (`vt:bool`).
    Bool(bool),
    /// A date/time value as ISO 8601 string (`vt:filetime`).
    DateTime(String),
    /// An unknown/unsupported value type preserved as raw XML.
    Raw(RawXmlNode),
}

/// A single custom property.
#[derive(Debug, Clone, PartialEq)]
pub struct CustomProperty {
    /// The property name.
    pub name: String,
    /// The property value.
    pub value: CustomPropertyValue,
    /// The format ID (always `{D5CDD505-2E9C-101B-9397-08002B2CF9AE}`).
    pub fmtid: String,
    /// The property ID (auto-assigned, starting from 2).
    pub pid: u32,
}

/// Custom properties from `docProps/custom.xml`.
///
/// User-defined key-value metadata. Each property has a name and a typed value.
#[derive(Debug, Clone, Default, PartialEq)]
pub struct CustomProperties {
    /// The custom properties.
    pub properties: Vec<CustomProperty>,
    /// Original raw XML bytes for no-op roundtrip.
    raw_xml: Option<Vec<u8>>,
    /// Whether any property has been modified since loading.
    dirty: bool,
}

impl CustomProperties {
    /// Create empty custom properties.
    pub fn new() -> Self {
        Self::default()
    }

    /// Parse from `docProps/custom.xml` bytes, preserving original for roundtrip.
    pub fn from_xml_bytes(bytes: Vec<u8>) -> Result<Self> {
        let mut props = Self::from_xml(std::io::Cursor::new(bytes.as_slice()))?;
        props.raw_xml = Some(bytes);
        props.dirty = false;
        Ok(props)
    }

    /// Parse from an XML reader.
    pub fn from_xml<R: BufRead>(reader: R) -> Result<Self> {
        let mut xml = Reader::from_reader(reader);
        xml.config_mut().trim_text(true);

        let mut props = CustomProperties::new();
        let mut buf = Vec::new();

        loop {
            match xml.read_event_into(&mut buf) {
                Ok(Event::Start(ref e)) => {
                    let name_bytes = e.name();
                    let local = local_name(name_bytes.as_ref());
                    if local == b"property" {
                        if let Some(prop) = parse_custom_property(&mut xml, e)? {
                            props.properties.push(prop);
                        }
                    }
                }
                Ok(Event::Eof) => break,
                Err(e) => return Err(e.into()),
                _ => {}
            }
            buf.clear();
        }

        props.dirty = false;
        Ok(props)
    }

    /// Serialize to XML. Uses original bytes on no-op roundtrip.
    pub fn to_xml<W: Write>(&self, writer: W) -> Result<()> {
        if !self.dirty {
            if let Some(ref raw) = self.raw_xml {
                let mut w = writer;
                w.write_all(raw)?;
                return Ok(());
            }
        }

        let mut xml = Writer::new_with_indent(writer, b' ', 2);

        xml.write_event(Event::Decl(BytesDecl::new(
            "1.0",
            Some("UTF-8"),
            Some("yes"),
        )))?;

        let mut root = BytesStart::new("Properties");
        root.push_attribute(("xmlns", NS_CUSTOM_PROPERTIES));
        root.push_attribute(("xmlns:vt", NS_VT));
        xml.write_event(Event::Start(root))?;

        for prop in &self.properties {
            let mut elem = BytesStart::new("property");
            elem.push_attribute(("fmtid", prop.fmtid.as_str()));
            elem.push_attribute(("pid", prop.pid.to_string().as_str()));
            elem.push_attribute(("name", prop.name.as_str()));
            xml.write_event(Event::Start(elem))?;

            match &prop.value {
                CustomPropertyValue::String(s) => {
                    xml.write_event(Event::Start(BytesStart::new("vt:lpwstr")))?;
                    xml.write_event(Event::Text(BytesText::new(s)))?;
                    xml.write_event(Event::End(BytesEnd::new("vt:lpwstr")))?;
                }
                CustomPropertyValue::Int(n) => {
                    xml.write_event(Event::Start(BytesStart::new("vt:i4")))?;
                    xml.write_event(Event::Text(BytesText::new(&n.to_string())))?;
                    xml.write_event(Event::End(BytesEnd::new("vt:i4")))?;
                }
                CustomPropertyValue::Float(f) => {
                    xml.write_event(Event::Start(BytesStart::new("vt:r8")))?;
                    xml.write_event(Event::Text(BytesText::new(&f.to_string())))?;
                    xml.write_event(Event::End(BytesEnd::new("vt:r8")))?;
                }
                CustomPropertyValue::Bool(b) => {
                    xml.write_event(Event::Start(BytesStart::new("vt:bool")))?;
                    xml.write_event(Event::Text(BytesText::new(if *b {
                        "true"
                    } else {
                        "false"
                    })))?;
                    xml.write_event(Event::End(BytesEnd::new("vt:bool")))?;
                }
                CustomPropertyValue::DateTime(dt) => {
                    xml.write_event(Event::Start(BytesStart::new("vt:filetime")))?;
                    xml.write_event(Event::Text(BytesText::new(dt)))?;
                    xml.write_event(Event::End(BytesEnd::new("vt:filetime")))?;
                }
                CustomPropertyValue::Raw(raw_node) => {
                    raw_node
                        .write_to(&mut xml)
                        .map_err(crate::error::OpcError::from)?;
                }
            }

            xml.write_event(Event::End(BytesEnd::new("property")))?;
        }

        xml.write_event(Event::End(BytesEnd::new("Properties")))?;

        Ok(())
    }

    /// Get a custom property by name.
    pub fn get(&self, name: &str) -> Option<&CustomProperty> {
        self.properties.iter().find(|p| p.name == name)
    }

    /// Set a custom property by name. If a property with the given name already
    /// exists, its value is updated. Otherwise a new property is appended with
    /// an auto-assigned pid.
    pub fn set(&mut self, name: &str, value: CustomPropertyValue) {
        self.dirty = true;
        if let Some(existing) = self.properties.iter_mut().find(|p| p.name == name) {
            existing.value = value;
        } else {
            let pid = self.next_pid();
            self.properties.push(CustomProperty {
                name: name.to_string(),
                value,
                fmtid: CUSTOM_PROPERTY_FMTID.to_string(),
                pid,
            });
        }
    }

    /// Remove a custom property by name. Returns the removed property if found.
    pub fn remove(&mut self, name: &str) -> Option<CustomProperty> {
        if let Some(pos) = self.properties.iter().position(|p| p.name == name) {
            self.dirty = true;
            Some(self.properties.remove(pos))
        } else {
            None
        }
    }

    /// Iterate over all custom properties.
    pub fn iter(&self) -> impl Iterator<Item = &CustomProperty> {
        self.properties.iter()
    }

    /// Number of custom properties.
    pub fn len(&self) -> usize {
        self.properties.len()
    }

    /// Whether there are no custom properties.
    pub fn is_empty(&self) -> bool {
        self.properties.is_empty()
    }

    /// Whether any property has been modified.
    pub fn is_dirty(&self) -> bool {
        self.dirty
    }

    /// Mark as clean.
    pub fn mark_clean(&mut self) {
        self.dirty = false;
    }

    /// Compute the next available pid (max existing pid + 1, or 2 if empty).
    fn next_pid(&self) -> u32 {
        self.properties
            .iter()
            .map(|p| p.pid)
            .max()
            .map(|m| m + 1)
            .unwrap_or(2)
    }
}

/// Parse a single `<property>` element into a `CustomProperty`.
fn parse_custom_property<R: BufRead>(
    xml: &mut Reader<R>,
    start: &BytesStart<'_>,
) -> Result<Option<CustomProperty>> {
    let mut name = String::new();
    let mut fmtid = CUSTOM_PROPERTY_FMTID.to_string();
    let mut pid: u32 = 2;

    for attr in start.attributes().flatten() {
        let key = String::from_utf8_lossy(attr.key.as_ref());
        let val = String::from_utf8_lossy(&attr.value);
        match key.as_ref() {
            "name" => name = val.into_owned(),
            "fmtid" => fmtid = val.into_owned(),
            "pid" => pid = val.parse().unwrap_or(2),
            _ => {}
        }
    }

    let mut buf = Vec::new();
    let mut value: Option<CustomPropertyValue> = None;

    loop {
        match xml.read_event_into(&mut buf) {
            Ok(Event::Start(ref e)) => {
                let name_bytes = e.name();
                let local = local_name(name_bytes.as_ref());
                match local {
                    b"lpwstr" => {
                        let text = read_text_content(xml, name_bytes.as_ref())?;
                        value = Some(CustomPropertyValue::String(text));
                    }
                    b"i4" => {
                        let text = read_text_content(xml, name_bytes.as_ref())?;
                        value = Some(CustomPropertyValue::Int(text.parse().unwrap_or(0)));
                    }
                    b"r8" => {
                        let text = read_text_content(xml, name_bytes.as_ref())?;
                        value = Some(CustomPropertyValue::Float(text.parse().unwrap_or(0.0)));
                    }
                    b"bool" => {
                        let text = read_text_content(xml, name_bytes.as_ref())?;
                        value = Some(CustomPropertyValue::Bool(text == "true" || text == "1"));
                    }
                    b"filetime" => {
                        let text = read_text_content(xml, name_bytes.as_ref())?;
                        value = Some(CustomPropertyValue::DateTime(text));
                    }
                    _ => {
                        // Unknown value type — preserve as RawXmlNode
                        let raw = RawXmlNode::read_element(xml, e)?;
                        value = Some(CustomPropertyValue::Raw(raw));
                    }
                }
            }
            Ok(Event::End(ref e)) => {
                let end_name = e.name();
                let local = local_name(end_name.as_ref());
                if local == b"property" {
                    break;
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => return Err(e.into()),
            _ => {}
        }
        buf.clear();
    }

    match value {
        Some(v) => Ok(Some(CustomProperty {
            name,
            value: v,
            fmtid,
            pid,
        })),
        None => Ok(None),
    }
}

/// Read the text content of a simple element and consume its closing tag.
fn read_text_content<R: BufRead>(xml: &mut Reader<R>, end_name: &[u8]) -> Result<String> {
    let mut buf = Vec::new();
    let mut text = String::new();

    loop {
        match xml.read_event_into(&mut buf) {
            Ok(Event::Text(ref e)) => {
                text.push_str(&decode_text_event(e)?);
            }
            Ok(Event::GeneralRef(ref e)) => {
                text.push_str(&decode_general_ref(e)?);
            }
            Ok(Event::End(ref e)) if e.name().as_ref() == end_name => {
                break;
            }
            Ok(Event::Eof) => break,
            Err(e) => return Err(e.into()),
            _ => {}
        }
        buf.clear();
    }

    Ok(text)
}

// --- Helpers ---

fn local_name(name: &[u8]) -> &[u8] {
    name.rsplit(|byte| *byte == b':').next().unwrap_or(name)
}

fn decode_text_event(event: &BytesText<'_>) -> Result<String> {
    event
        .xml_content()
        .map(|text| text.into_owned())
        .map_err(quick_xml::Error::from)
        .map_err(Into::into)
}

fn decode_general_ref(event: &BytesRef<'_>) -> Result<String> {
    let reference = event
        .decode()
        .map_err(quick_xml::Error::from)
        .map_err(crate::error::OpcError::from)?;
    let escaped = format!("&{};", reference);
    quick_xml::escape::unescape(&escaped)
        .map(|text| text.into_owned())
        .map_err(quick_xml::Error::from)
        .map_err(Into::into)
}

fn set_core_property(props: &mut CoreProperties, elem: &str, text: String) {
    match elem {
        "title" => props.title = Some(text),
        "subject" => props.subject = Some(text),
        "creator" => props.creator = Some(text),
        "keywords" => props.keywords = Some(text),
        "description" => props.description = Some(text),
        "lastModifiedBy" => props.last_modified_by = Some(text),
        "revision" => props.revision = Some(text),
        "created" => props.created = Some(text),
        "modified" => props.modified = Some(text),
        "lastPrinted" => props.last_printed = Some(text),
        "category" => props.category = Some(text),
        "contentType" => props.content_type = Some(text),
        "contentStatus" => props.content_status = Some(text),
        "identifier" => props.identifier = Some(text),
        "language" => props.language = Some(text),
        "version" => props.version = Some(text),
        _ => {}
    }
}

fn set_extended_property(props: &mut ExtendedProperties, elem: &str, text: String) {
    match elem {
        "Application" => props.application = Some(text),
        "AppVersion" => props.app_version = Some(text),
        "Template" => props.template = Some(text),
        "Manager" => props.manager = Some(text),
        "Company" => props.company = Some(text),
        "TotalTime" => props.total_time = Some(text),
        "Pages" => props.pages = Some(text),
        "Words" => props.words = Some(text),
        "Characters" => props.characters = Some(text),
        "CharactersWithSpaces" => props.characters_with_spaces = Some(text),
        "Lines" => props.lines = Some(text),
        "Paragraphs" => props.paragraphs = Some(text),
        "Slides" => props.slides = Some(text),
        "Notes" => props.notes = Some(text),
        "HiddenSlides" => props.hidden_slides = Some(text),
        "DocSecurity" => props.doc_security = Some(text),
        _ => {}
    }
}

fn write_element<W: Write>(xml: &mut Writer<W>, name: &str, value: Option<&str>) -> Result<()> {
    if let Some(val) = value {
        xml.write_event(Event::Start(BytesStart::new(name)))?;
        xml.write_event(Event::Text(BytesText::new(val)))?;
        xml.write_event(Event::End(BytesEnd::new(name)))?;
    }
    Ok(())
}

fn write_dc_element<W: Write>(xml: &mut Writer<W>, name: &str, value: Option<&str>) -> Result<()> {
    if let Some(val) = value {
        xml.write_event(Event::Start(BytesStart::new(name)))?;
        xml.write_event(Event::Text(BytesText::new(val)))?;
        xml.write_event(Event::End(BytesEnd::new(name)))?;
    }
    Ok(())
}

fn write_dcterms_element<W: Write>(
    xml: &mut Writer<W>,
    name: &str,
    value: Option<&str>,
) -> Result<()> {
    if let Some(val) = value {
        let mut elem = BytesStart::new(name);
        elem.push_attribute(("xsi:type", "dcterms:W3CDTF"));
        xml.write_event(Event::Start(elem))?;
        xml.write_event(Event::Text(BytesText::new(val)))?;
        xml.write_event(Event::End(BytesEnd::new(name)))?;
    }
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn parse_core_properties() {
        let xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
    xmlns:dc="http://purl.org/dc/elements/1.1/"
    xmlns:dcterms="http://purl.org/dc/terms/"
    xmlns:dcmitype="http://purl.org/dc/dcmitype/"
    xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
  <dc:title>Test Document</dc:title>
  <dc:subject>Testing</dc:subject>
  <dc:creator>John Doe</dc:creator>
  <cp:keywords>test keyword</cp:keywords>
  <dc:description>A test document.</dc:description>
  <cp:lastModifiedBy>Jane Doe</cp:lastModifiedBy>
  <cp:revision>3</cp:revision>
  <dcterms:created xsi:type="dcterms:W3CDTF">2024-01-15T10:30:00Z</dcterms:created>
  <dcterms:modified xsi:type="dcterms:W3CDTF">2024-06-20T14:00:00Z</dcterms:modified>
  <cp:category>Report</cp:category>
  <dc:language>en-US</dc:language>
</cp:coreProperties>"#;

        let props = CoreProperties::from_xml_bytes(xml.to_vec()).unwrap();

        assert_eq!(props.title.as_deref(), Some("Test Document"));
        assert_eq!(props.subject.as_deref(), Some("Testing"));
        assert_eq!(props.creator.as_deref(), Some("John Doe"));
        assert_eq!(props.keywords.as_deref(), Some("test keyword"));
        assert_eq!(props.description.as_deref(), Some("A test document."));
        assert_eq!(props.last_modified_by.as_deref(), Some("Jane Doe"));
        assert_eq!(props.revision.as_deref(), Some("3"));
        assert_eq!(props.created.as_deref(), Some("2024-01-15T10:30:00Z"));
        assert_eq!(props.modified.as_deref(), Some("2024-06-20T14:00:00Z"));
        assert_eq!(props.category.as_deref(), Some("Report"));
        assert_eq!(props.language.as_deref(), Some("en-US"));
    }

    #[test]
    fn core_properties_no_op_roundtrip() {
        let xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
    xmlns:dc="http://purl.org/dc/elements/1.1/">
  <dc:title>Hello</dc:title>
  <dc:creator>Test</dc:creator>
</cp:coreProperties>"#;

        let props = CoreProperties::from_xml_bytes(xml.to_vec()).unwrap();
        let mut output = Vec::new();
        props.to_xml(&mut output).unwrap();
        assert_eq!(output, xml, "no-op roundtrip should preserve bytes");
    }

    #[test]
    fn core_properties_dirty_regenerates_xml() {
        let mut props = CoreProperties::new();
        props.set_title("My Document");
        props.set_creator("offidized");
        props.set_created("2024-01-01T00:00:00Z");

        let mut output = Vec::new();
        props.to_xml(&mut output).unwrap();
        let text = String::from_utf8(output).unwrap();

        assert!(text.contains("<dc:title>My Document</dc:title>"));
        assert!(text.contains("<dc:creator>offidized</dc:creator>"));
        assert!(text.contains("dcterms:W3CDTF"));
    }

    #[test]
    fn parse_extended_properties() {
        let xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties">
  <Application>Microsoft Office Word</Application>
  <AppVersion>16.0000</AppVersion>
  <Company>Contoso</Company>
  <TotalTime>42</TotalTime>
  <Pages>5</Pages>
  <Words>1234</Words>
  <Characters>7890</Characters>
  <DocSecurity>0</DocSecurity>
</Properties>"#;

        let props = ExtendedProperties::from_xml_bytes(xml.to_vec()).unwrap();

        assert_eq!(props.application.as_deref(), Some("Microsoft Office Word"));
        assert_eq!(props.app_version.as_deref(), Some("16.0000"));
        assert_eq!(props.company.as_deref(), Some("Contoso"));
        assert_eq!(props.total_time.as_deref(), Some("42"));
        assert_eq!(props.pages.as_deref(), Some("5"));
        assert_eq!(props.words.as_deref(), Some("1234"));
        assert_eq!(props.characters.as_deref(), Some("7890"));
        assert_eq!(props.doc_security.as_deref(), Some("0"));
    }

    #[test]
    fn extended_properties_no_op_roundtrip() {
        let xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties">
  <Application>Excel</Application>
</Properties>"#;

        let props = ExtendedProperties::from_xml_bytes(xml.to_vec()).unwrap();
        let mut output = Vec::new();
        props.to_xml(&mut output).unwrap();
        assert_eq!(output, xml, "no-op roundtrip should preserve bytes");
    }

    #[test]
    fn remove_by_id_on_relationships() {
        use crate::relationship::{Relationships, TargetMode};

        let mut rels = Relationships::new();
        rels.add_new(
            "type/a".to_string(),
            "a.xml".to_string(),
            TargetMode::Internal,
        );
        rels.add_new(
            "type/b".to_string(),
            "b.xml".to_string(),
            TargetMode::Internal,
        );
        rels.add_new(
            "type/c".to_string(),
            "c.xml".to_string(),
            TargetMode::Internal,
        );

        assert_eq!(rels.len(), 3);

        let removed = rels.remove_by_id("rId2");
        assert!(removed.is_some());
        assert_eq!(removed.unwrap().target, "b.xml");
        assert_eq!(rels.len(), 2);

        assert!(rels.get_by_id("rId1").is_some());
        assert!(rels.get_by_id("rId2").is_none());
        assert!(rels.get_by_id("rId3").is_some());
    }

    #[test]
    fn remove_by_type_on_relationships() {
        use crate::relationship::{Relationships, TargetMode};

        let mut rels = Relationships::new();
        rels.add_new(
            "type/a".to_string(),
            "a.xml".to_string(),
            TargetMode::Internal,
        );
        rels.add_new(
            "type/b".to_string(),
            "b1.xml".to_string(),
            TargetMode::Internal,
        );
        rels.add_new(
            "type/b".to_string(),
            "b2.xml".to_string(),
            TargetMode::Internal,
        );
        rels.add_new(
            "type/c".to_string(),
            "c.xml".to_string(),
            TargetMode::Internal,
        );

        let removed = rels.remove_by_type("type/b");
        assert_eq!(removed.len(), 2);
        assert_eq!(rels.len(), 2);

        assert!(rels.get_by_id("rId1").is_some());
        assert!(rels.get_by_id("rId4").is_some());
        assert!(rels.get_by_type("type/b").is_empty());
    }

    #[test]
    fn parse_custom_properties() {
        let xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties"
    xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
  <property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="2" name="MyProp">
    <vt:lpwstr>Hello World</vt:lpwstr>
  </property>
  <property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="3" name="Count">
    <vt:i4>42</vt:i4>
  </property>
  <property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="4" name="IsActive">
    <vt:bool>true</vt:bool>
  </property>
</Properties>"#;

        let props = CustomProperties::from_xml_bytes(xml.to_vec()).unwrap();

        assert_eq!(props.len(), 3);
        assert!(!props.is_empty());
        assert!(!props.is_dirty());

        let my_prop = props.get("MyProp").unwrap();
        assert_eq!(my_prop.name, "MyProp");
        assert_eq!(my_prop.pid, 2);
        assert_eq!(my_prop.fmtid, "{D5CDD505-2E9C-101B-9397-08002B2CF9AE}");
        assert_eq!(
            my_prop.value,
            CustomPropertyValue::String("Hello World".to_string())
        );

        let count = props.get("Count").unwrap();
        assert_eq!(count.value, CustomPropertyValue::Int(42));
        assert_eq!(count.pid, 3);

        let active = props.get("IsActive").unwrap();
        assert_eq!(active.value, CustomPropertyValue::Bool(true));
        assert_eq!(active.pid, 4);

        assert!(props.get("NonExistent").is_none());
    }

    #[test]
    fn custom_properties_no_op_roundtrip() {
        let xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties"
    xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
  <property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="2" name="MyProp">
    <vt:lpwstr>Hello World</vt:lpwstr>
  </property>
</Properties>"#;

        let props = CustomProperties::from_xml_bytes(xml.to_vec()).unwrap();
        assert!(!props.is_dirty());

        let mut output = Vec::new();
        props.to_xml(&mut output).unwrap();
        assert_eq!(output, xml, "no-op roundtrip should preserve bytes");
    }

    #[test]
    fn custom_properties_dirty_regenerates_xml() {
        let mut props = CustomProperties::new();
        props.set("Author", CustomPropertyValue::String("Alice".to_string()));
        props.set("Version", CustomPropertyValue::Int(3));
        props.set("Rating", CustomPropertyValue::Float(4.5));
        props.set("Published", CustomPropertyValue::Bool(true));
        props.set(
            "ReleaseDate",
            CustomPropertyValue::DateTime("2024-01-01T00:00:00Z".to_string()),
        );

        assert!(props.is_dirty());
        assert_eq!(props.len(), 5);

        let mut output = Vec::new();
        props.to_xml(&mut output).unwrap();
        let text = String::from_utf8(output).unwrap();

        assert!(text.contains("<vt:lpwstr>Alice</vt:lpwstr>"));
        assert!(text.contains("<vt:i4>3</vt:i4>"));
        assert!(text.contains("<vt:r8>4.5</vt:r8>"));
        assert!(text.contains("<vt:bool>true</vt:bool>"));
        assert!(text.contains("<vt:filetime>2024-01-01T00:00:00Z</vt:filetime>"));
        assert!(text.contains("name=\"Author\""));
        assert!(text.contains("name=\"Version\""));
    }

    #[test]
    fn custom_properties_get_set_remove() {
        let mut props = CustomProperties::new();
        assert!(props.is_empty());
        assert_eq!(props.len(), 0);

        // Set new properties
        props.set("Name", CustomPropertyValue::String("Test".to_string()));
        assert_eq!(props.len(), 1);
        assert_eq!(props.get("Name").unwrap().pid, 2);

        props.set("Count", CustomPropertyValue::Int(10));
        assert_eq!(props.len(), 2);
        assert_eq!(props.get("Count").unwrap().pid, 3);

        // Update existing property
        props.set("Name", CustomPropertyValue::String("Updated".to_string()));
        assert_eq!(props.len(), 2);
        assert_eq!(
            props.get("Name").unwrap().value,
            CustomPropertyValue::String("Updated".to_string())
        );

        // Remove property
        let removed = props.remove("Count");
        assert!(removed.is_some());
        assert_eq!(removed.unwrap().value, CustomPropertyValue::Int(10));
        assert_eq!(props.len(), 1);
        assert!(props.get("Count").is_none());

        // Remove non-existent
        assert!(props.remove("NoSuch").is_none());
        assert_eq!(props.len(), 1);
    }

    #[test]
    fn custom_properties_iter() {
        let mut props = CustomProperties::new();
        props.set("A", CustomPropertyValue::String("a".to_string()));
        props.set("B", CustomPropertyValue::Int(2));

        let names: Vec<&str> = props.iter().map(|p| p.name.as_str()).collect();
        assert_eq!(names, vec!["A", "B"]);
    }

    #[test]
    fn custom_properties_mark_clean() {
        let mut props = CustomProperties::new();
        props.set("X", CustomPropertyValue::Bool(false));
        assert!(props.is_dirty());

        props.mark_clean();
        assert!(!props.is_dirty());
    }

    #[test]
    fn custom_properties_pid_auto_increment() {
        let mut props = CustomProperties::new();
        props.set("First", CustomPropertyValue::Int(1));
        props.set("Second", CustomPropertyValue::Int(2));
        props.set("Third", CustomPropertyValue::Int(3));

        assert_eq!(props.get("First").unwrap().pid, 2);
        assert_eq!(props.get("Second").unwrap().pid, 3);
        assert_eq!(props.get("Third").unwrap().pid, 4);
    }

    #[test]
    fn custom_properties_roundtrip_parse_regenerate() {
        let xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties"
    xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
  <property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="2" name="Greeting">
    <vt:lpwstr>Hello</vt:lpwstr>
  </property>
  <property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="3" name="Number">
    <vt:i4>7</vt:i4>
  </property>
</Properties>"#;

        let mut props = CustomProperties::from_xml_bytes(xml.to_vec()).unwrap();

        // Modify to make dirty — regeneration should produce valid XML
        props.set("Greeting", CustomPropertyValue::String("World".to_string()));
        assert!(props.is_dirty());

        let mut output = Vec::new();
        props.to_xml(&mut output).unwrap();
        let text = String::from_utf8(output.clone()).unwrap();

        assert!(text.contains("<vt:lpwstr>World</vt:lpwstr>"));
        assert!(text.contains("<vt:i4>7</vt:i4>"));

        // Re-parse the generated XML to verify it's valid
        let reparsed = CustomProperties::from_xml_bytes(output).unwrap();
        assert_eq!(reparsed.len(), 2);
        assert_eq!(
            reparsed.get("Greeting").unwrap().value,
            CustomPropertyValue::String("World".to_string())
        );
        assert_eq!(
            reparsed.get("Number").unwrap().value,
            CustomPropertyValue::Int(7)
        );
    }
}
