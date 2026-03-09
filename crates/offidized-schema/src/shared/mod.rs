pub use crate::generated::shared::SharedElementDescriptor as ElementDescriptor;

/// Returns all registered shared OOXML element descriptors.
#[must_use]
pub fn elements() -> &'static [ElementDescriptor] {
    crate::generated::shared::ELEMENTS
}

/// Returns the number of registered shared OOXML elements.
#[must_use]
pub const fn count() -> usize {
    crate::generated::shared::ELEMENTS.len()
}

/// Finds a shared element descriptor by schema path.
#[must_use]
pub fn find_by_path(schema_path: &str) -> Option<&'static ElementDescriptor> {
    crate::generated::shared::ELEMENTS
        .iter()
        .find(|descriptor| descriptor.schema_path == schema_path)
}

/// Finds a shared element descriptor by class name.
#[must_use]
pub fn find_by_class(class_name: &str) -> Option<&'static ElementDescriptor> {
    crate::generated::shared::ELEMENTS
        .iter()
        .find(|descriptor| descriptor.class_name == class_name)
}

#[cfg(test)]
mod tests {
    use super::{count, elements, find_by_class, find_by_path};

    #[test]
    fn registry_is_not_empty() {
        assert!(!elements().is_empty());
        assert_eq!(count(), elements().len());
    }

    #[test]
    fn path_lookup_returns_descriptor() {
        let descriptor = find_by_path("/a:txBody/a:p/a:r/a:t");
        assert!(descriptor.is_some());
        if let Some(descriptor) = descriptor {
            assert_eq!(descriptor.class_name, "DrawingText");
        }
    }

    #[test]
    fn class_lookup_returns_descriptor() {
        let descriptor = find_by_class("AlternateContent");
        assert!(descriptor.is_some());
        if let Some(descriptor) = descriptor {
            assert_eq!(descriptor.schema_path, "/mc:AlternateContent");
        }
    }
}
