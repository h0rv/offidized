pub use crate::catalog::presentationml_main_typed_api::*;
pub use crate::catalog::PresentationmlMainEntry as ElementDescriptor;
pub use crate::catalog::PresentationmlMainTypeKind as ElementTypeKind;
pub use crate::catalog::PresentationmlMainTypedElement as TypedElement;

/// Returns all registered PresentationML element descriptors.
#[must_use]
pub fn elements() -> &'static [ElementDescriptor] {
    crate::catalog::presentationml_main_elements()
}

/// Returns the number of registered PresentationML elements.
#[must_use]
pub fn count() -> usize {
    crate::catalog::presentationml_main_elements().len()
}

/// Finds a PresentationML element descriptor by schema path.
#[must_use]
pub fn find_by_path(schema_path: &str) -> Option<&'static ElementDescriptor> {
    elements()
        .iter()
        .find(|descriptor| descriptor.schema_path == schema_path)
}

/// Finds a PresentationML element descriptor by class name.
#[must_use]
pub fn find_by_class(class_name: &str) -> Option<&'static ElementDescriptor> {
    crate::catalog::find_presentationml_main_by_class(class_name)
}

/// Finds a PresentationML element descriptor by qualified XML name.
#[must_use]
pub fn find_by_qualified_name(qualified_name: &str) -> Option<&'static ElementDescriptor> {
    crate::catalog::find_presentationml_main_by_qualified_name(qualified_name)
}

/// Constructs an empty typed element payload by class name.
#[must_use]
pub fn new_typed_element_by_class(class_name: &str) -> Option<TypedElement> {
    crate::catalog::new_presentationml_main_typed_element_by_class(class_name)
}

/// Constructs an empty typed element payload by qualified XML name.
#[must_use]
pub fn new_typed_element_by_qualified_name(qualified_name: &str) -> Option<TypedElement> {
    crate::catalog::new_presentationml_main_typed_element_by_qualified_name(qualified_name)
}

/// Parses a PresentationML typed element from an XML snippet.
pub fn parse_typed_element(
    snippet: &str,
) -> Result<TypedElement, crate::catalog::TypedElementXmlError> {
    crate::catalog::parse_presentationml_main_typed_element(snippet)
}

/// Parses a PresentationML typed element for a specific class.
pub fn parse_typed_element_for_class(
    class_name: &str,
    snippet: &str,
) -> Result<TypedElement, crate::catalog::TypedElementXmlError> {
    crate::catalog::parse_presentationml_main_typed_element_for_class(class_name, snippet)
}

#[cfg(test)]
mod tests {
    use super::{
        count, elements, find_by_class, find_by_path, find_by_qualified_name,
        parse_typed_element_for_class, ElementTypeKind, Presentation,
    };
    use crate::raw::RawXmlNode;

    #[test]
    fn registry_is_not_empty() {
        assert!(!elements().is_empty());
        assert_eq!(count(), elements().len());
    }

    #[test]
    fn path_lookup_returns_descriptor() {
        let descriptor = find_by_path("/p:presentation");
        assert!(descriptor.is_some());
        if let Some(descriptor) = descriptor {
            assert_eq!(descriptor.class_name, "Presentation");
        }
    }

    #[test]
    fn class_lookup_returns_descriptor() {
        let descriptor = find_by_class("Presentation");
        assert!(descriptor.is_some());
        if let Some(descriptor) = descriptor {
            assert_eq!(descriptor.schema_path, "/p:presentation");
        }
    }

    #[test]
    fn qualified_name_lookup_and_kind_match() {
        let descriptor = find_by_qualified_name("p:presentation");
        assert!(descriptor.is_some());
        if let Some(descriptor) = descriptor {
            assert_eq!(descriptor.class_name, "Presentation");
            assert_eq!(descriptor.type_kind, ElementTypeKind::Composite);
        }
    }

    #[test]
    fn typed_element_known_and_unknown_attribute_split() {
        let parsed =
            parse_typed_element_for_class("SlideId", "<p:sldId id=\"256\" custom=\"v\" />")
                .expect("PresentationML SlideId snippet should parse");
        assert_eq!(
            parsed.known_attrs.get("id").map(String::as_str),
            Some("256")
        );
        assert!(parsed
            .unknown_attrs
            .iter()
            .any(|(name, value)| name == "custom" && value == "v"));
    }

    #[test]
    fn presentation_wrapper_roundtrip_preserves_unknown_payload() {
        assert_eq!(
            Presentation::descriptor()
                .expect("PresentationML Presentation descriptor should exist")
                .qualified_name,
            "p:presentation"
        );

        let parsed = Presentation::parse(
            "<p:presentation customAttr=\"z\"><p:unknown customChild=\"1\"/></p:presentation>",
        )
        .expect("PresentationML Presentation wrapper should parse");
        assert!(parsed
            .inner()
            .unknown_attrs
            .iter()
            .any(|(name, value)| name == "customAttr" && value == "z"));
        assert_eq!(parsed.inner().unknown_children.len(), 1);

        let serialized = parsed
            .to_xml()
            .expect("PresentationML Presentation wrapper should serialize");
        let reparsed = Presentation::parse(&serialized)
            .expect("PresentationML Presentation wrapper should deserialize");
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
                assert_eq!(name, "p:unknown");
                assert!(attributes
                    .iter()
                    .any(|(key, value)| key == "customChild" && value == "1"));
            }
            _ => panic!("expected unknown presentation child element"),
        }
    }
}
