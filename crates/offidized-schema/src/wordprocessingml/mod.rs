pub use crate::catalog::wordprocessingml_main_typed_api::*;
pub use crate::catalog::WordprocessingmlMainEntry as ElementDescriptor;
pub use crate::catalog::WordprocessingmlMainTypeKind as ElementTypeKind;
pub use crate::catalog::WordprocessingmlMainTypedElement as TypedElement;

/// Returns all registered WordprocessingML element descriptors.
#[must_use]
pub fn elements() -> &'static [ElementDescriptor] {
    crate::catalog::wordprocessingml_main_elements()
}

/// Returns the number of registered WordprocessingML elements.
#[must_use]
pub fn count() -> usize {
    crate::catalog::wordprocessingml_main_elements().len()
}

/// Finds a WordprocessingML element descriptor by schema path.
#[must_use]
pub fn find_by_path(schema_path: &str) -> Option<&'static ElementDescriptor> {
    elements()
        .iter()
        .find(|descriptor| descriptor.schema_path == schema_path)
}

/// Finds a WordprocessingML element descriptor by class name.
#[must_use]
pub fn find_by_class(class_name: &str) -> Option<&'static ElementDescriptor> {
    crate::catalog::find_wordprocessingml_main_by_class(class_name)
}

/// Finds a WordprocessingML element descriptor by qualified XML name.
#[must_use]
pub fn find_by_qualified_name(qualified_name: &str) -> Option<&'static ElementDescriptor> {
    crate::catalog::find_wordprocessingml_main_by_qualified_name(qualified_name)
}

/// Constructs an empty typed element payload by class name.
#[must_use]
pub fn new_typed_element_by_class(class_name: &str) -> Option<TypedElement> {
    crate::catalog::new_wordprocessingml_main_typed_element_by_class(class_name)
}

/// Constructs an empty typed element payload by qualified XML name.
#[must_use]
pub fn new_typed_element_by_qualified_name(qualified_name: &str) -> Option<TypedElement> {
    crate::catalog::new_wordprocessingml_main_typed_element_by_qualified_name(qualified_name)
}

/// Parses a WordprocessingML typed element from an XML snippet.
pub fn parse_typed_element(
    snippet: &str,
) -> Result<TypedElement, crate::catalog::TypedElementXmlError> {
    crate::catalog::parse_wordprocessingml_main_typed_element(snippet)
}

/// Parses a WordprocessingML typed element for a specific class.
pub fn parse_typed_element_for_class(
    class_name: &str,
    snippet: &str,
) -> Result<TypedElement, crate::catalog::TypedElementXmlError> {
    crate::catalog::parse_wordprocessingml_main_typed_element_for_class(class_name, snippet)
}

#[cfg(test)]
mod tests {
    use super::{
        count, elements, find_by_class, find_by_path, find_by_qualified_name,
        parse_typed_element_for_class, Document, ElementTypeKind,
    };
    use crate::raw::RawXmlNode;

    #[test]
    fn registry_is_not_empty() {
        assert!(!elements().is_empty());
        assert_eq!(count(), elements().len());
    }

    #[test]
    fn path_lookup_returns_descriptor() {
        let descriptor = find_by_path("/w:document");
        assert!(descriptor.is_some());
        if let Some(descriptor) = descriptor {
            assert_eq!(descriptor.class_name, "Document");
        }
    }

    #[test]
    fn class_lookup_returns_descriptor() {
        let descriptor = find_by_class("Table");
        assert!(descriptor.is_some());
        if let Some(descriptor) = descriptor {
            assert_eq!(descriptor.schema_path, "/w:tbl");
        }
    }

    #[test]
    fn qualified_name_lookup_and_kind_match() {
        let descriptor = find_by_qualified_name("w:t");
        assert!(descriptor.is_some());
        if let Some(descriptor) = descriptor {
            assert_eq!(descriptor.class_name, "Text");
            assert_eq!(descriptor.type_kind, ElementTypeKind::TypedLeaf);
        }
    }

    #[test]
    fn typed_element_known_and_unknown_attribute_split() {
        let parsed =
            parse_typed_element_for_class("Text", "<w:t xml:space=\"preserve\" custom=\"v\" />")
                .expect("WordprocessingML Text snippet should parse");
        assert_eq!(
            parsed.known_attrs.get("xml:space").map(String::as_str),
            Some("preserve")
        );
        assert!(parsed
            .unknown_attrs
            .iter()
            .any(|(name, value)| name == "custom" && value == "v"));
    }

    #[test]
    fn document_wrapper_roundtrip_preserves_unknown_payload() {
        assert_eq!(
            Document::descriptor()
                .expect("WordprocessingML Document descriptor should exist")
                .qualified_name,
            "w:document"
        );

        let parsed = Document::parse(
            "<w:document customAttr=\"z\"><w:unknown customChild=\"1\"/></w:document>",
        )
        .expect("WordprocessingML Document wrapper should parse");
        assert!(parsed
            .inner()
            .unknown_attrs
            .iter()
            .any(|(name, value)| name == "customAttr" && value == "z"));
        assert_eq!(parsed.inner().unknown_children.len(), 1);

        let serialized = parsed
            .to_xml()
            .expect("WordprocessingML Document wrapper should serialize");
        let reparsed = Document::parse(&serialized)
            .expect("WordprocessingML Document wrapper should deserialize");
        assert!(reparsed
            .inner()
            .unknown_attrs
            .iter()
            .any(|(name, value)| name == "customAttr" && value == "z"));
        assert_eq!(reparsed.inner().unknown_children.len(), 1);

        match &reparsed.inner().unknown_children[0] {
            RawXmlNode::Element {
                name, attributes, ..
            } => {
                assert_eq!(name, "w:unknown");
                assert!(attributes
                    .iter()
                    .any(|(key, value)| key == "customChild" && value == "1"));
            }
            _ => panic!("expected unknown document child element"),
        }
    }
}
