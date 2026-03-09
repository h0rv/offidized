//! Convenience type for hyperlink relationships.
//!
//! In the Open XML SDK, `HyperlinkRelationship` is a typed wrapper around
//! external relationships whose type is the OPC hyperlink URI. This module
//! provides an equivalent helper for creating and querying hyperlinks.

use crate::relationship::{Relationship, RelationshipType, Relationships, TargetMode};

/// A convenience wrapper for hyperlink relationships.
///
/// Hyperlinks in OPC are external relationships with the well-known type
/// `http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink`.
#[derive(Debug, Clone)]
pub struct HyperlinkRelationship {
    /// The relationship ID (e.g., "rId5").
    pub id: String,
    /// The hyperlink URI (e.g., "https://example.com").
    pub uri: String,
}

impl HyperlinkRelationship {
    /// Create a new hyperlink relationship with the given URI.
    pub fn new(id: impl Into<String>, uri: impl Into<String>) -> Self {
        Self {
            id: id.into(),
            uri: uri.into(),
        }
    }

    /// Extract hyperlink relationships from a `Relationships` collection.
    pub fn from_relationships(rels: &Relationships) -> Vec<Self> {
        rels.get_by_type(RelationshipType::HYPERLINK)
            .into_iter()
            .filter(|r| r.target_mode == TargetMode::External)
            .map(|r| Self {
                id: r.id.clone(),
                uri: r.target.clone(),
            })
            .collect()
    }

    /// Add this hyperlink to a `Relationships` collection.
    /// Returns the generated relationship ID.
    pub fn add_to(uri: impl Into<String>, rels: &mut Relationships) -> String {
        let rel = rels.add_new(
            RelationshipType::HYPERLINK.to_string(),
            uri.into(),
            TargetMode::External,
        );
        rel.id.clone()
    }

    /// Convert to a raw `Relationship`.
    pub fn to_relationship(&self) -> Relationship {
        Relationship {
            id: self.id.clone(),
            rel_type: RelationshipType::HYPERLINK.to_string(),
            target: self.uri.clone(),
            target_mode: TargetMode::External,
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn add_and_extract_hyperlinks() {
        let mut rels = Relationships::new();

        let id1 = HyperlinkRelationship::add_to("https://example.com", &mut rels);
        let id2 = HyperlinkRelationship::add_to("https://rust-lang.org", &mut rels);

        let hyperlinks = HyperlinkRelationship::from_relationships(&rels);
        assert_eq!(hyperlinks.len(), 2);
        assert_eq!(hyperlinks[0].id, id1);
        assert_eq!(hyperlinks[0].uri, "https://example.com");
        assert_eq!(hyperlinks[1].id, id2);
        assert_eq!(hyperlinks[1].uri, "https://rust-lang.org");
    }

    #[test]
    fn to_relationship_produces_external_hyperlink() {
        let link = HyperlinkRelationship::new("rId1", "https://example.com");
        let rel = link.to_relationship();

        assert_eq!(rel.id, "rId1");
        assert_eq!(rel.rel_type, RelationshipType::HYPERLINK);
        assert_eq!(rel.target, "https://example.com");
        assert_eq!(rel.target_mode, TargetMode::External);
    }

    #[test]
    fn ignores_internal_hyperlink_type_relationships() {
        let mut rels = Relationships::new();
        // Add an internal relationship with hyperlink type (unusual but possible)
        rels.add(Relationship {
            id: "rId1".to_string(),
            rel_type: RelationshipType::HYPERLINK.to_string(),
            target: "/some/part.xml".to_string(),
            target_mode: TargetMode::Internal,
        });
        // Add a normal external hyperlink
        HyperlinkRelationship::add_to("https://example.com", &mut rels);

        let hyperlinks = HyperlinkRelationship::from_relationships(&rels);
        assert_eq!(hyperlinks.len(), 1);
        assert_eq!(hyperlinks[0].uri, "https://example.com");
    }
}
