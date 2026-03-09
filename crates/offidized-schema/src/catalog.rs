//! Comprehensive schema/part catalogs generated from Open-XML-SDK data.
//!
//! These registries are generated during build by `offidized-codegen` and
//! embedded via `include!` from `OUT_DIR`.

use std::fmt;

mod presentationml_main {
    include!(concat!(env!("OUT_DIR"), "/presentationml_main_registry.rs"));
}

pub mod presentationml_main_typed_api {
    include!(concat!(
        env!("OUT_DIR"),
        "/presentationml_main_typed_api.rs"
    ));
}

mod shared_parts {
    include!(concat!(env!("OUT_DIR"), "/shared_part_registry.rs"));
}

mod spreadsheetml_main {
    include!(concat!(env!("OUT_DIR"), "/spreadsheetml_main_registry.rs"));
}

pub mod spreadsheetml_main_typed_api {
    include!(concat!(env!("OUT_DIR"), "/spreadsheetml_main_typed_api.rs"));
}

mod wordprocessingml_main {
    include!(concat!(
        env!("OUT_DIR"),
        "/wordprocessingml_main_registry.rs"
    ));
}

pub mod wordprocessingml_main_typed_api {
    include!(concat!(
        env!("OUT_DIR"),
        "/wordprocessingml_main_typed_api.rs"
    ));
}

#[derive(Debug)]
pub enum TypedElementXmlError {
    Xml(quick_xml::Error),
    Io(std::io::Error),
    Utf8(std::string::FromUtf8Error),
    MissingRootElement,
    UnexpectedEof,
    UnknownQualifiedName(String),
    UnknownClassName(String),
    UnexpectedRoot {
        expected: &'static str,
        found: String,
    },
}

impl fmt::Display for TypedElementXmlError {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::Xml(error) => write!(f, "xml error: {error}"),
            Self::Io(error) => write!(f, "I/O error: {error}"),
            Self::Utf8(error) => write!(f, "utf-8 error: {error}"),
            Self::MissingRootElement => write!(f, "missing root element in XML snippet"),
            Self::UnexpectedEof => write!(f, "unexpected EOF while parsing XML snippet"),
            Self::UnknownQualifiedName(qualified_name) => {
                write!(f, "unknown typed element qualified name: {qualified_name}")
            }
            Self::UnknownClassName(class_name) => {
                write!(f, "unknown typed element class name: {class_name}")
            }
            Self::UnexpectedRoot { expected, found } => {
                write!(
                    f,
                    "unexpected root element `{found}`; expected `{expected}`"
                )
            }
        }
    }
}

impl std::error::Error for TypedElementXmlError {}

impl From<quick_xml::Error> for TypedElementXmlError {
    fn from(value: quick_xml::Error) -> Self {
        Self::Xml(value)
    }
}

impl From<std::io::Error> for TypedElementXmlError {
    fn from(value: std::io::Error) -> Self {
        Self::Io(value)
    }
}

impl From<std::string::FromUtf8Error> for TypedElementXmlError {
    fn from(value: std::string::FromUtf8Error) -> Self {
        Self::Utf8(value)
    }
}

pub type PresentationmlMainEntry = presentationml_main::TypedElementDescriptor;
pub type SpreadsheetmlMainEntry = spreadsheetml_main::TypedElementDescriptor;
pub type WordprocessingmlMainEntry = wordprocessingml_main::TypedElementDescriptor;

pub type PresentationmlMainTypedElement = presentationml_main::TypedElement;
pub type SpreadsheetmlMainTypedElement = spreadsheetml_main::TypedElement;
pub type WordprocessingmlMainTypedElement = wordprocessingml_main::TypedElement;

pub type PresentationmlMainTypeKind = presentationml_main::TypedElementKind;
pub type SpreadsheetmlMainTypeKind = spreadsheetml_main::TypedElementKind;
pub type WordprocessingmlMainTypeKind = wordprocessingml_main::TypedElementKind;

pub type PartEntry = shared_parts::PartRegistryEntry;
pub type PartChildEntry = shared_parts::PartChildRegistryEntry;
pub type PartPathEntry = shared_parts::PartPathRegistryEntry;

#[must_use]
pub fn spreadsheetml_main_elements() -> &'static [SpreadsheetmlMainEntry] {
    spreadsheetml_main::SPREADSHEETML_MAIN_REGISTRY
}

#[must_use]
pub fn wordprocessingml_main_elements() -> &'static [WordprocessingmlMainEntry] {
    wordprocessingml_main::WORDPROCESSINGML_MAIN_REGISTRY
}

#[must_use]
pub fn presentationml_main_elements() -> &'static [PresentationmlMainEntry] {
    presentationml_main::PRESENTATIONML_MAIN_REGISTRY
}

#[must_use]
pub fn spreadsheetml_main_namespace_uri() -> &'static str {
    spreadsheetml_main::NAMESPACE_URI
}

#[must_use]
pub fn spreadsheetml_main_namespace_prefix() -> &'static str {
    spreadsheetml_main::NAMESPACE_PREFIX
}

#[must_use]
pub fn wordprocessingml_main_namespace_uri() -> &'static str {
    wordprocessingml_main::NAMESPACE_URI
}

#[must_use]
pub fn wordprocessingml_main_namespace_prefix() -> &'static str {
    wordprocessingml_main::NAMESPACE_PREFIX
}

#[must_use]
pub fn presentationml_main_namespace_uri() -> &'static str {
    presentationml_main::NAMESPACE_URI
}

#[must_use]
pub fn presentationml_main_namespace_prefix() -> &'static str {
    presentationml_main::NAMESPACE_PREFIX
}

#[must_use]
pub fn find_spreadsheetml_main_by_qualified_name(
    qualified_name: &str,
) -> Option<&'static SpreadsheetmlMainEntry> {
    spreadsheetml_main::SPREADSHEETML_MAIN_REGISTRY
        .iter()
        .find(|entry| entry.qualified_name == qualified_name)
}

#[must_use]
pub fn find_wordprocessingml_main_by_qualified_name(
    qualified_name: &str,
) -> Option<&'static WordprocessingmlMainEntry> {
    wordprocessingml_main::WORDPROCESSINGML_MAIN_REGISTRY
        .iter()
        .find(|entry| entry.qualified_name == qualified_name)
}

#[must_use]
pub fn find_presentationml_main_by_qualified_name(
    qualified_name: &str,
) -> Option<&'static PresentationmlMainEntry> {
    presentationml_main::PRESENTATIONML_MAIN_REGISTRY
        .iter()
        .find(|entry| entry.qualified_name == qualified_name)
}

#[must_use]
pub fn new_spreadsheetml_main_typed_element_by_class(
    class_name: &str,
) -> Option<SpreadsheetmlMainTypedElement> {
    spreadsheetml_main::new_typed_element_by_class(class_name)
}

#[must_use]
pub fn new_spreadsheetml_main_typed_element_by_qualified_name(
    qualified_name: &str,
) -> Option<SpreadsheetmlMainTypedElement> {
    spreadsheetml_main::new_typed_element_by_qualified_name(qualified_name)
}

#[must_use]
pub fn new_wordprocessingml_main_typed_element_by_class(
    class_name: &str,
) -> Option<WordprocessingmlMainTypedElement> {
    wordprocessingml_main::new_typed_element_by_class(class_name)
}

#[must_use]
pub fn new_wordprocessingml_main_typed_element_by_qualified_name(
    qualified_name: &str,
) -> Option<WordprocessingmlMainTypedElement> {
    wordprocessingml_main::new_typed_element_by_qualified_name(qualified_name)
}

#[must_use]
pub fn new_presentationml_main_typed_element_by_class(
    class_name: &str,
) -> Option<PresentationmlMainTypedElement> {
    presentationml_main::new_typed_element_by_class(class_name)
}

#[must_use]
pub fn new_presentationml_main_typed_element_by_qualified_name(
    qualified_name: &str,
) -> Option<PresentationmlMainTypedElement> {
    presentationml_main::new_typed_element_by_qualified_name(qualified_name)
}

pub fn parse_spreadsheetml_main_typed_element(
    snippet: &str,
) -> Result<SpreadsheetmlMainTypedElement, TypedElementXmlError> {
    spreadsheetml_main::TypedElement::from_xml_snippet(snippet)
}

pub fn parse_wordprocessingml_main_typed_element(
    snippet: &str,
) -> Result<WordprocessingmlMainTypedElement, TypedElementXmlError> {
    wordprocessingml_main::TypedElement::from_xml_snippet(snippet)
}

pub fn parse_presentationml_main_typed_element(
    snippet: &str,
) -> Result<PresentationmlMainTypedElement, TypedElementXmlError> {
    presentationml_main::TypedElement::from_xml_snippet(snippet)
}

pub fn parse_spreadsheetml_main_typed_element_for_class(
    class_name: &str,
    snippet: &str,
) -> Result<SpreadsheetmlMainTypedElement, TypedElementXmlError> {
    spreadsheetml_main::parse_typed_element_for_class(class_name, snippet)
}

pub fn parse_wordprocessingml_main_typed_element_for_class(
    class_name: &str,
    snippet: &str,
) -> Result<WordprocessingmlMainTypedElement, TypedElementXmlError> {
    wordprocessingml_main::parse_typed_element_for_class(class_name, snippet)
}

pub fn parse_presentationml_main_typed_element_for_class(
    class_name: &str,
    snippet: &str,
) -> Result<PresentationmlMainTypedElement, TypedElementXmlError> {
    presentationml_main::parse_typed_element_for_class(class_name, snippet)
}

#[must_use]
pub fn shared_parts() -> &'static [PartEntry] {
    shared_parts::SHARED_PART_REGISTRY
}

#[must_use]
pub fn find_spreadsheetml_main_by_class(
    class_name: &str,
) -> Option<&'static SpreadsheetmlMainEntry> {
    spreadsheetml_main::SPREADSHEETML_MAIN_REGISTRY
        .iter()
        .find(|entry| entry.class_name == class_name)
}

#[must_use]
pub fn find_wordprocessingml_main_by_class(
    class_name: &str,
) -> Option<&'static WordprocessingmlMainEntry> {
    wordprocessingml_main::WORDPROCESSINGML_MAIN_REGISTRY
        .iter()
        .find(|entry| entry.class_name == class_name)
}

#[must_use]
pub fn find_presentationml_main_by_class(
    class_name: &str,
) -> Option<&'static PresentationmlMainEntry> {
    presentationml_main::PRESENTATIONML_MAIN_REGISTRY
        .iter()
        .find(|entry| entry.class_name == class_name)
}

#[must_use]
pub fn find_part_by_name(name: &str) -> Option<&'static PartEntry> {
    shared_parts::SHARED_PART_REGISTRY
        .iter()
        .find(|entry| entry.name == name)
}

#[must_use]
pub fn find_part_by_relationship_type(relationship_type: &str) -> Option<&'static PartEntry> {
    shared_parts::SHARED_PART_REGISTRY.iter().find(|entry| {
        entry
            .relationship_type
            .is_some_and(|candidate| candidate == relationship_type)
    })
}

#[cfg(test)]
mod tests {
    use super::{
        find_part_by_name, find_presentationml_main_by_class,
        find_presentationml_main_by_qualified_name, find_spreadsheetml_main_by_class,
        find_spreadsheetml_main_by_qualified_name, find_wordprocessingml_main_by_class,
        find_wordprocessingml_main_by_qualified_name,
        new_wordprocessingml_main_typed_element_by_class,
        parse_wordprocessingml_main_typed_element_for_class, presentationml_main_elements,
        presentationml_main_namespace_prefix, presentationml_main_namespace_uri,
        presentationml_main_typed_api, shared_parts, spreadsheetml_main_elements,
        spreadsheetml_main_namespace_prefix, spreadsheetml_main_namespace_uri,
        spreadsheetml_main_typed_api, wordprocessingml_main_elements,
        wordprocessingml_main_namespace_prefix, wordprocessingml_main_namespace_uri,
        wordprocessingml_main_typed_api, PresentationmlMainTypeKind, SpreadsheetmlMainTypeKind,
        WordprocessingmlMainTypeKind,
    };
    use crate::raw::RawXmlNode;

    #[test]
    fn generated_catalogs_are_non_empty() {
        assert!(!spreadsheetml_main_elements().is_empty());
        assert!(!wordprocessingml_main_elements().is_empty());
        assert!(!presentationml_main_elements().is_empty());
        assert!(!shared_parts().is_empty());
        assert_eq!(
            spreadsheetml_main_namespace_uri(),
            "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
        );
        assert_eq!(
            wordprocessingml_main_namespace_uri(),
            "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
        );
        assert_eq!(
            presentationml_main_namespace_uri(),
            "http://schemas.openxmlformats.org/presentationml/2006/main"
        );
        assert_eq!(spreadsheetml_main_namespace_prefix(), "x");
        assert_eq!(wordprocessingml_main_namespace_prefix(), "w");
        assert_eq!(presentationml_main_namespace_prefix(), "p");
    }

    #[test]
    fn class_and_qualified_name_lookups_work() {
        let workbook = find_spreadsheetml_main_by_class("Workbook")
            .expect("SpreadsheetML class lookup should find Workbook");
        assert_eq!(workbook.qualified_name, "x:workbook");
        let workbook_by_qname = find_spreadsheetml_main_by_qualified_name("x:workbook")
            .expect("SpreadsheetML qualified name lookup should find Workbook");
        assert_eq!(workbook_by_qname.class_name, "Workbook");

        let document = find_wordprocessingml_main_by_class("Document")
            .expect("WordprocessingML class lookup should find Document");
        assert_eq!(document.qualified_name, "w:document");
        let document_by_qname = find_wordprocessingml_main_by_qualified_name("w:document")
            .expect("WordprocessingML qualified name lookup should find Document");
        assert_eq!(document_by_qname.class_name, "Document");

        let presentation = find_presentationml_main_by_class("Presentation")
            .expect("PresentationML class lookup should find Presentation");
        assert_eq!(presentation.qualified_name, "p:presentation");
        let presentation_by_qname = find_presentationml_main_by_qualified_name("p:presentation")
            .expect("PresentationML qualified name lookup should find Presentation");
        assert_eq!(presentation_by_qname.class_name, "Presentation");

        assert!(find_part_by_name("WorkbookPart").is_some());
    }

    #[test]
    fn known_type_kinds_match_expected_values() {
        let workbook = find_spreadsheetml_main_by_class("Workbook")
            .expect("SpreadsheetML class lookup should find Workbook");
        assert_eq!(workbook.type_kind, SpreadsheetmlMainTypeKind::Composite);

        let document = find_wordprocessingml_main_by_class("Document")
            .expect("WordprocessingML class lookup should find Document");
        assert_eq!(document.type_kind, WordprocessingmlMainTypeKind::Composite);

        let presentation = find_presentationml_main_by_class("Presentation")
            .expect("PresentationML class lookup should find Presentation");
        assert_eq!(
            presentation.type_kind,
            PresentationmlMainTypeKind::Composite
        );

        let text = find_wordprocessingml_main_by_class("Text")
            .expect("WordprocessingML class lookup should find Text");
        assert_eq!(text.type_kind, WordprocessingmlMainTypeKind::TypedLeaf);
        assert!(text.known_attributes.contains(&"xml:space"));
    }

    #[test]
    fn typed_element_roundtrip_preserves_unknown_children_and_attrs() {
        let created = new_wordprocessingml_main_typed_element_by_class("Text")
            .expect("WordprocessingML typed element constructor should find Text");
        assert_eq!(created.descriptor.qualified_name, "w:t");

        let snippet =
            "<w:t xml:space=\"preserve\" customAttr=\"z\"><w:unknown customChild=\"1\"/></w:t>";
        let parsed = parse_wordprocessingml_main_typed_element_for_class("Text", snippet)
            .expect("WordprocessingML typed element parser should parse Text");
        assert_eq!(parsed.descriptor.class_name, "Text");
        assert_eq!(
            parsed.known_attrs.get("xml:space").map(String::as_str),
            Some("preserve")
        );
        assert!(parsed
            .unknown_attrs
            .iter()
            .any(|(name, value)| name == "customAttr" && value == "z"));
        assert_eq!(parsed.unknown_children.len(), 1);
        match &parsed.unknown_children[0] {
            RawXmlNode::Element {
                name, attributes, ..
            } => {
                assert_eq!(name, "w:unknown");
                assert!(attributes
                    .iter()
                    .any(|(key, value)| key == "customChild" && value == "1"));
            }
            _ => panic!("expected unknown child element"),
        }

        let serialized = parsed
            .to_xml_snippet()
            .expect("WordprocessingML typed element should serialize");
        let reparsed = parse_wordprocessingml_main_typed_element_for_class("Text", &serialized)
            .expect("WordprocessingML typed element should deserialize after serialize");
        assert_eq!(
            reparsed.known_attrs.get("xml:space").map(String::as_str),
            Some("preserve")
        );
        assert!(reparsed
            .unknown_attrs
            .iter()
            .any(|(name, value)| name == "customAttr" && value == "z"));
        assert_eq!(reparsed.unknown_children.len(), 1);
    }

    #[test]
    fn typed_wrappers_exist_and_roundtrip_unknown_payload() {
        let workbook = spreadsheetml_main_typed_api::Workbook::parse(
            "<x:workbook conformance=\"strict\" customAttr=\"z\"><x:unknown customChild=\"1\"/></x:workbook>",
        )
        .expect("SpreadsheetML Workbook wrapper should parse");
        assert_eq!(
            spreadsheetml_main_typed_api::Workbook::descriptor()
                .expect("SpreadsheetML Workbook descriptor should exist")
                .qualified_name,
            "x:workbook"
        );
        assert_eq!(workbook.conformance(), Some("strict"));
        assert!(workbook
            .inner()
            .unknown_attrs
            .iter()
            .any(|(name, value)| name == "customAttr" && value == "z"));
        assert_eq!(workbook.inner().unknown_children.len(), 1);
        let workbook_xml = workbook
            .to_xml()
            .expect("SpreadsheetML Workbook wrapper should serialize");
        let workbook_reparsed = spreadsheetml_main_typed_api::Workbook::parse(&workbook_xml)
            .expect("SpreadsheetML Workbook wrapper should deserialize");
        assert!(workbook_reparsed
            .inner()
            .unknown_attrs
            .iter()
            .any(|(name, value)| name == "customAttr" && value == "z"));

        let document = wordprocessingml_main_typed_api::Document::parse(
            "<w:document customAttr=\"z\"><w:unknown customChild=\"1\"/></w:document>",
        )
        .expect("WordprocessingML Document wrapper should parse");
        assert_eq!(
            wordprocessingml_main_typed_api::Document::descriptor()
                .expect("WordprocessingML Document descriptor should exist")
                .qualified_name,
            "w:document"
        );
        assert!(document
            .inner()
            .unknown_attrs
            .iter()
            .any(|(name, value)| name == "customAttr" && value == "z"));
        assert_eq!(document.inner().unknown_children.len(), 1);
        let document_xml = document
            .to_xml()
            .expect("WordprocessingML Document wrapper should serialize");
        let document_reparsed = wordprocessingml_main_typed_api::Document::parse(&document_xml)
            .expect("WordprocessingML Document wrapper should deserialize");
        assert!(document_reparsed
            .inner()
            .unknown_attrs
            .iter()
            .any(|(name, value)| name == "customAttr" && value == "z"));

        let presentation = presentationml_main_typed_api::Presentation::parse(
            "<p:presentation customAttr=\"z\"><p:unknown customChild=\"1\"/></p:presentation>",
        )
        .expect("PresentationML Presentation wrapper should parse");
        assert_eq!(
            presentationml_main_typed_api::Presentation::descriptor()
                .expect("PresentationML Presentation descriptor should exist")
                .qualified_name,
            "p:presentation"
        );
        assert!(presentation
            .inner()
            .unknown_attrs
            .iter()
            .any(|(name, value)| name == "customAttr" && value == "z"));
        assert_eq!(presentation.inner().unknown_children.len(), 1);
        let presentation_xml = presentation
            .to_xml()
            .expect("PresentationML Presentation wrapper should serialize");
        let presentation_reparsed =
            presentationml_main_typed_api::Presentation::parse(&presentation_xml)
                .expect("PresentationML Presentation wrapper should deserialize");
        assert!(presentation_reparsed
            .inner()
            .unknown_attrs
            .iter()
            .any(|(name, value)| name == "customAttr" && value == "z"));

        match &workbook_reparsed.inner().unknown_children[0] {
            RawXmlNode::Element {
                name, attributes, ..
            } => {
                assert_eq!(name, "x:unknown");
                assert!(attributes
                    .iter()
                    .any(|(key, value)| key == "customChild" && value == "1"));
            }
            _ => panic!("expected unknown workbook child element"),
        }
    }
}
