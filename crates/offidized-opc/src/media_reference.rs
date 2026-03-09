//! Typed media reference relationship helpers.
//!
//! In the Open XML SDK, `AudioReferenceRelationship`,
//! `VideoReferenceRelationship`, and `DataPartReferenceRelationship` are
//! typed wrappers around internal relationships that point to media parts.
//! This module provides equivalent helpers.

use crate::relationship::{Relationship, RelationshipType, Relationships, TargetMode};

/// A convenience wrapper for audio reference relationships.
///
/// Points to an internal audio media part with the well-known type
/// `http://schemas.openxmlformats.org/officeDocument/2006/relationships/audio`.
#[derive(Debug, Clone)]
pub struct AudioReferenceRelationship {
    /// The relationship ID (e.g., "rId3").
    pub id: String,
    /// The target part URI (e.g., "/ppt/media/audio1.wav").
    pub target: String,
}

impl AudioReferenceRelationship {
    /// Create a new audio reference relationship.
    pub fn new(id: impl Into<String>, target: impl Into<String>) -> Self {
        Self {
            id: id.into(),
            target: target.into(),
        }
    }

    /// Extract audio reference relationships from a `Relationships` collection.
    pub fn from_relationships(rels: &Relationships) -> Vec<Self> {
        rels.get_by_type(RelationshipType::AUDIO)
            .into_iter()
            .filter(|r| r.target_mode == TargetMode::Internal)
            .map(|r| Self {
                id: r.id.clone(),
                target: r.target.clone(),
            })
            .collect()
    }

    /// Add an audio reference to a `Relationships` collection.
    /// Returns the generated relationship ID.
    pub fn add_to(target: impl Into<String>, rels: &mut Relationships) -> String {
        let rel = rels.add_new(
            RelationshipType::AUDIO.to_string(),
            target.into(),
            TargetMode::Internal,
        );
        rel.id.clone()
    }

    /// Convert to a raw `Relationship`.
    pub fn to_relationship(&self) -> Relationship {
        Relationship {
            id: self.id.clone(),
            rel_type: RelationshipType::AUDIO.to_string(),
            target: self.target.clone(),
            target_mode: TargetMode::Internal,
        }
    }
}

/// A convenience wrapper for video reference relationships.
///
/// Points to an internal video media part with the well-known type
/// `http://schemas.openxmlformats.org/officeDocument/2006/relationships/video`.
#[derive(Debug, Clone)]
pub struct VideoReferenceRelationship {
    /// The relationship ID (e.g., "rId4").
    pub id: String,
    /// The target part URI (e.g., "/ppt/media/video1.mp4").
    pub target: String,
}

impl VideoReferenceRelationship {
    /// Create a new video reference relationship.
    pub fn new(id: impl Into<String>, target: impl Into<String>) -> Self {
        Self {
            id: id.into(),
            target: target.into(),
        }
    }

    /// Extract video reference relationships from a `Relationships` collection.
    pub fn from_relationships(rels: &Relationships) -> Vec<Self> {
        rels.get_by_type(RelationshipType::VIDEO)
            .into_iter()
            .filter(|r| r.target_mode == TargetMode::Internal)
            .map(|r| Self {
                id: r.id.clone(),
                target: r.target.clone(),
            })
            .collect()
    }

    /// Add a video reference to a `Relationships` collection.
    /// Returns the generated relationship ID.
    pub fn add_to(target: impl Into<String>, rels: &mut Relationships) -> String {
        let rel = rels.add_new(
            RelationshipType::VIDEO.to_string(),
            target.into(),
            TargetMode::Internal,
        );
        rel.id.clone()
    }

    /// Convert to a raw `Relationship`.
    pub fn to_relationship(&self) -> Relationship {
        Relationship {
            id: self.id.clone(),
            rel_type: RelationshipType::VIDEO.to_string(),
            target: self.target.clone(),
            target_mode: TargetMode::Internal,
        }
    }
}

/// A convenience wrapper for data part reference relationships.
///
/// Points to an internal binary data part using the generic media type
/// `http://schemas.microsoft.com/office/2007/relationships/media`.
/// Used for embedded media that doesn't fit the standard audio/video types.
#[derive(Debug, Clone)]
pub struct DataPartReferenceRelationship {
    /// The relationship ID (e.g., "rId5").
    pub id: String,
    /// The target part URI (e.g., "/ppt/media/media1.bin").
    pub target: String,
}

impl DataPartReferenceRelationship {
    /// Create a new data part reference relationship.
    pub fn new(id: impl Into<String>, target: impl Into<String>) -> Self {
        Self {
            id: id.into(),
            target: target.into(),
        }
    }

    /// Extract data part reference relationships from a `Relationships` collection.
    pub fn from_relationships(rels: &Relationships) -> Vec<Self> {
        rels.get_by_type(RelationshipType::MEDIA)
            .into_iter()
            .filter(|r| r.target_mode == TargetMode::Internal)
            .map(|r| Self {
                id: r.id.clone(),
                target: r.target.clone(),
            })
            .collect()
    }

    /// Add a data part reference to a `Relationships` collection.
    /// Returns the generated relationship ID.
    pub fn add_to(target: impl Into<String>, rels: &mut Relationships) -> String {
        let rel = rels.add_new(
            RelationshipType::MEDIA.to_string(),
            target.into(),
            TargetMode::Internal,
        );
        rel.id.clone()
    }

    /// Convert to a raw `Relationship`.
    pub fn to_relationship(&self) -> Relationship {
        Relationship {
            id: self.id.clone(),
            rel_type: RelationshipType::MEDIA.to_string(),
            target: self.target.clone(),
            target_mode: TargetMode::Internal,
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn add_and_extract_audio_references() {
        let mut rels = Relationships::new();

        let id1 = AudioReferenceRelationship::add_to("media/audio1.wav", &mut rels);
        let id2 = AudioReferenceRelationship::add_to("media/audio2.mp3", &mut rels);

        let audio_refs = AudioReferenceRelationship::from_relationships(&rels);
        assert_eq!(audio_refs.len(), 2);
        assert_eq!(audio_refs[0].id, id1);
        assert_eq!(audio_refs[0].target, "media/audio1.wav");
        assert_eq!(audio_refs[1].id, id2);
        assert_eq!(audio_refs[1].target, "media/audio2.mp3");
    }

    #[test]
    fn add_and_extract_video_references() {
        let mut rels = Relationships::new();

        let id = VideoReferenceRelationship::add_to("media/video1.mp4", &mut rels);

        let video_refs = VideoReferenceRelationship::from_relationships(&rels);
        assert_eq!(video_refs.len(), 1);
        assert_eq!(video_refs[0].id, id);
        assert_eq!(video_refs[0].target, "media/video1.mp4");
    }

    #[test]
    fn add_and_extract_data_part_references() {
        let mut rels = Relationships::new();

        let id = DataPartReferenceRelationship::add_to("media/media1.bin", &mut rels);

        let data_refs = DataPartReferenceRelationship::from_relationships(&rels);
        assert_eq!(data_refs.len(), 1);
        assert_eq!(data_refs[0].id, id);
        assert_eq!(data_refs[0].target, "media/media1.bin");
    }

    #[test]
    fn audio_to_relationship_produces_internal_audio() {
        let audio = AudioReferenceRelationship::new("rId1", "media/audio1.wav");
        let rel = audio.to_relationship();

        assert_eq!(rel.id, "rId1");
        assert_eq!(rel.rel_type, RelationshipType::AUDIO);
        assert_eq!(rel.target, "media/audio1.wav");
        assert_eq!(rel.target_mode, TargetMode::Internal);
    }

    #[test]
    fn video_to_relationship_produces_internal_video() {
        let video = VideoReferenceRelationship::new("rId2", "media/video1.mp4");
        let rel = video.to_relationship();

        assert_eq!(rel.id, "rId2");
        assert_eq!(rel.rel_type, RelationshipType::VIDEO);
        assert_eq!(rel.target, "media/video1.mp4");
        assert_eq!(rel.target_mode, TargetMode::Internal);
    }

    #[test]
    fn ignores_external_media_relationships() {
        let mut rels = Relationships::new();
        // Add an external relationship with audio type (e.g., linked audio)
        rels.add(Relationship {
            id: "rId1".to_string(),
            rel_type: RelationshipType::AUDIO.to_string(),
            target: "https://example.com/audio.mp3".to_string(),
            target_mode: TargetMode::External,
        });
        // Add a normal internal audio reference
        AudioReferenceRelationship::add_to("media/audio1.wav", &mut rels);

        let audio_refs = AudioReferenceRelationship::from_relationships(&rels);
        assert_eq!(audio_refs.len(), 1);
        assert_eq!(audio_refs[0].target, "media/audio1.wav");
    }
}
