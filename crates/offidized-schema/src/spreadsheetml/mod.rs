pub use crate::catalog::spreadsheetml_main_typed_api::*;
pub use crate::catalog::SpreadsheetmlMainEntry as ElementDescriptor;
pub use crate::catalog::SpreadsheetmlMainTypeKind as ElementTypeKind;
pub use crate::catalog::SpreadsheetmlMainTypedElement as TypedElement;

/// Returns all registered SpreadsheetML element descriptors.
#[must_use]
pub fn elements() -> &'static [ElementDescriptor] {
    crate::catalog::spreadsheetml_main_elements()
}

/// Returns the number of registered SpreadsheetML elements.
#[must_use]
pub fn count() -> usize {
    crate::catalog::spreadsheetml_main_elements().len()
}

/// Finds a SpreadsheetML element descriptor by schema path.
#[must_use]
pub fn find_by_path(schema_path: &str) -> Option<&'static ElementDescriptor> {
    elements()
        .iter()
        .find(|descriptor| descriptor.schema_path == schema_path)
}

/// Finds a SpreadsheetML element descriptor by class name.
#[must_use]
pub fn find_by_class(class_name: &str) -> Option<&'static ElementDescriptor> {
    crate::catalog::find_spreadsheetml_main_by_class(class_name)
}

/// Finds a SpreadsheetML element descriptor by qualified XML name.
#[must_use]
pub fn find_by_qualified_name(qualified_name: &str) -> Option<&'static ElementDescriptor> {
    crate::catalog::find_spreadsheetml_main_by_qualified_name(qualified_name)
}

/// Constructs an empty typed element payload by class name.
#[must_use]
pub fn new_typed_element_by_class(class_name: &str) -> Option<TypedElement> {
    crate::catalog::new_spreadsheetml_main_typed_element_by_class(class_name)
}

/// Constructs an empty typed element payload by qualified XML name.
#[must_use]
pub fn new_typed_element_by_qualified_name(qualified_name: &str) -> Option<TypedElement> {
    crate::catalog::new_spreadsheetml_main_typed_element_by_qualified_name(qualified_name)
}

/// Parses a SpreadsheetML typed element from an XML snippet.
pub fn parse_typed_element(
    snippet: &str,
) -> Result<TypedElement, crate::catalog::TypedElementXmlError> {
    crate::catalog::parse_spreadsheetml_main_typed_element(snippet)
}

/// Parses a SpreadsheetML typed element for a specific class.
pub fn parse_typed_element_for_class(
    class_name: &str,
    snippet: &str,
) -> Result<TypedElement, crate::catalog::TypedElementXmlError> {
    crate::catalog::parse_spreadsheetml_main_typed_element_for_class(class_name, snippet)
}

#[cfg(test)]
mod tests {
    use super::{
        count, elements, find_by_class, find_by_path, find_by_qualified_name,
        parse_typed_element_for_class, ElementTypeKind, Workbook,
    };
    use crate::raw::RawXmlNode;

    #[test]
    fn registry_is_not_empty() {
        assert!(!elements().is_empty());
        assert_eq!(count(), elements().len());
    }

    #[test]
    fn path_lookup_returns_descriptor() {
        let descriptor = find_by_path("/x:workbook");
        assert!(descriptor.is_some());
        if let Some(descriptor) = descriptor {
            assert_eq!(descriptor.class_name, "Workbook");
        }
    }

    #[test]
    fn class_lookup_returns_descriptor() {
        let descriptor = find_by_class("Worksheet");
        assert!(descriptor.is_some());
        if let Some(descriptor) = descriptor {
            assert_eq!(descriptor.schema_path, "/x:worksheet");
        }
    }

    #[test]
    fn qualified_name_lookup_and_kind_match() {
        let descriptor = find_by_qualified_name("x:workbook");
        assert!(descriptor.is_some());
        if let Some(descriptor) = descriptor {
            assert_eq!(descriptor.class_name, "Workbook");
            assert_eq!(descriptor.type_kind, ElementTypeKind::Composite);
        }
    }

    #[test]
    fn typed_element_known_and_unknown_attribute_split() {
        let parsed = parse_typed_element_for_class(
            "Workbook",
            "<x:workbook conformance=\"strict\" custom=\"v\" />",
        )
        .expect("SpreadsheetML Workbook snippet should parse");
        assert_eq!(
            parsed.known_attrs.get("conformance").map(String::as_str),
            Some("strict")
        );
        assert!(parsed
            .unknown_attrs
            .iter()
            .any(|(name, value)| name == "custom" && value == "v"));
    }

    #[test]
    fn workbook_wrapper_roundtrip_preserves_unknown_payload() {
        assert_eq!(
            Workbook::descriptor()
                .expect("SpreadsheetML Workbook descriptor should exist")
                .qualified_name,
            "x:workbook"
        );

        let parsed = Workbook::parse(
            "<x:workbook conformance=\"strict\" customAttr=\"z\"><x:unknown customChild=\"1\"/></x:workbook>",
        )
        .expect("SpreadsheetML Workbook wrapper should parse");
        assert_eq!(parsed.conformance(), Some("strict"));
        assert!(parsed
            .inner()
            .unknown_attrs
            .iter()
            .any(|(name, value)| name == "customAttr" && value == "z"));
        assert_eq!(parsed.inner().unknown_children.len(), 1);

        let serialized = parsed
            .to_xml()
            .expect("SpreadsheetML Workbook wrapper should serialize");
        let reparsed = Workbook::parse(&serialized)
            .expect("SpreadsheetML Workbook wrapper should deserialize");
        assert_eq!(reparsed.conformance(), Some("strict"));
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
                assert_eq!(name, "x:unknown");
                assert!(attributes
                    .iter()
                    .any(|(key, value)| key == "customChild" && value == "1"));
            }
            _ => panic!("expected unknown workbook child element"),
        }
    }
}
