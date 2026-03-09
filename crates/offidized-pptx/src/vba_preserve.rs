//! VBA macro preservation for roundtrip fidelity.
//!
//! This module provides support for preserving VBA macros (VBA project parts) when
//! opening and saving presentations. It does NOT execute or modify VBA code - it only
//! ensures that existing VBA projects survive a roundtrip unchanged.
//!
//! VBA macros in PowerPoint are stored in:
//! - `/ppt/vbaProject.bin` - The VBA project binary
//! - Content type: `application/vnd.ms-office.vbaProject`
//! - Relationships from presentation part
//!
//! ## Safety and Limitations
//!
//! - This module is read-only for VBA content preservation
//! - No VBA execution or modification is supported
//! - VBA signatures are preserved but not validated
//! - Users are responsible for macro security settings

use std::collections::HashMap;

/// VBA project metadata and binary content.
///
/// Represents a VBA project embedded in a presentation. The binary content
/// is stored as opaque bytes and preserved during roundtrip operations.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct VbaProject {
    /// The relationship ID linking the presentation to the VBA project.
    relationship_id: String,
    /// The part name (typically "vbaProject.bin").
    part_name: String,
    /// Content type (typically "application/vnd.ms-office.vbaProject").
    content_type: String,
    /// Raw binary content of the VBA project.
    /// This is preserved as-is for roundtrip fidelity.
    binary_content: Vec<u8>,
}

impl VbaProject {
    pub fn new(
        relationship_id: impl Into<String>,
        part_name: impl Into<String>,
        content_type: impl Into<String>,
        binary_content: Vec<u8>,
    ) -> Self {
        Self {
            relationship_id: relationship_id.into(),
            part_name: part_name.into(),
            content_type: content_type.into(),
            binary_content,
        }
    }

    pub fn relationship_id(&self) -> &str {
        &self.relationship_id
    }

    pub fn set_relationship_id(&mut self, relationship_id: impl Into<String>) {
        self.relationship_id = relationship_id.into();
    }

    pub fn part_name(&self) -> &str {
        &self.part_name
    }

    pub fn set_part_name(&mut self, part_name: impl Into<String>) {
        self.part_name = part_name.into();
    }

    pub fn content_type(&self) -> &str {
        &self.content_type
    }

    pub fn set_content_type(&mut self, content_type: impl Into<String>) {
        self.content_type = content_type.into();
    }

    pub fn binary_content(&self) -> &[u8] {
        &self.binary_content
    }

    pub fn set_binary_content(&mut self, binary_content: Vec<u8>) {
        self.binary_content = binary_content;
    }

    /// Check if this VBA project has any content.
    pub fn is_empty(&self) -> bool {
        self.binary_content.is_empty()
    }

    /// Get the size of the VBA project in bytes.
    pub fn size(&self) -> usize {
        self.binary_content.len()
    }
}

/// VBA project signature metadata.
///
/// Represents digital signatures attached to VBA projects. These are preserved
/// but not validated by this library.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct VbaSignature {
    /// Relationship ID for the signature part.
    relationship_id: String,
    /// Content type (typically "application/vnd.ms-office.vbaProjectSignature").
    content_type: String,
    /// Raw signature binary content.
    binary_content: Vec<u8>,
}

impl VbaSignature {
    pub fn new(
        relationship_id: impl Into<String>,
        content_type: impl Into<String>,
        binary_content: Vec<u8>,
    ) -> Self {
        Self {
            relationship_id: relationship_id.into(),
            content_type: content_type.into(),
            binary_content,
        }
    }

    pub fn relationship_id(&self) -> &str {
        &self.relationship_id
    }

    pub fn content_type(&self) -> &str {
        &self.content_type
    }

    pub fn binary_content(&self) -> &[u8] {
        &self.binary_content
    }

    pub fn size(&self) -> usize {
        self.binary_content.len()
    }
}

/// VBA macro container managing projects and signatures.
///
/// Provides roundtrip preservation of VBA macros without execution or modification.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct VbaMacroContainer {
    /// The VBA project (if present).
    project: Option<VbaProject>,
    /// VBA project signature (if present).
    signature: Option<VbaSignature>,
    /// Additional VBA-related parts (for future extensibility).
    additional_parts: HashMap<String, Vec<u8>>,
}

impl Default for VbaMacroContainer {
    fn default() -> Self {
        Self::new()
    }
}

impl VbaMacroContainer {
    pub fn new() -> Self {
        Self {
            project: None,
            signature: None,
            additional_parts: HashMap::new(),
        }
    }

    /// Check if any VBA content is present.
    pub fn has_macros(&self) -> bool {
        self.project.is_some()
    }

    pub fn project(&self) -> Option<&VbaProject> {
        self.project.as_ref()
    }

    pub fn project_mut(&mut self) -> Option<&mut VbaProject> {
        self.project.as_mut()
    }

    pub fn set_project(&mut self, project: Option<VbaProject>) {
        self.project = project;
    }

    pub fn signature(&self) -> Option<&VbaSignature> {
        self.signature.as_ref()
    }

    pub fn set_signature(&mut self, signature: Option<VbaSignature>) {
        self.signature = signature;
    }

    pub fn additional_parts(&self) -> &HashMap<String, Vec<u8>> {
        &self.additional_parts
    }

    pub fn additional_parts_mut(&mut self) -> &mut HashMap<String, Vec<u8>> {
        &mut self.additional_parts
    }

    /// Add an additional VBA-related part.
    pub fn add_additional_part(&mut self, name: impl Into<String>, content: Vec<u8>) {
        self.additional_parts.insert(name.into(), content);
    }

    /// Remove an additional VBA-related part.
    pub fn remove_additional_part(&mut self, name: &str) -> Option<Vec<u8>> {
        self.additional_parts.remove(name)
    }

    /// Clear all VBA content (removes project, signature, and additional parts).
    pub fn clear(&mut self) {
        self.project = None;
        self.signature = None;
        self.additional_parts.clear();
    }

    /// Get total size of all VBA content in bytes.
    pub fn total_size(&self) -> usize {
        let project_size = self.project.as_ref().map(|p| p.size()).unwrap_or(0);
        let signature_size = self.signature.as_ref().map(|s| s.size()).unwrap_or(0);
        let additional_size: usize = self.additional_parts.values().map(|v| v.len()).sum();

        project_size + signature_size + additional_size
    }
}

/// VBA-related content type constants.
pub mod content_types {
    /// VBA project binary content type.
    pub const VBA_PROJECT: &str = "application/vnd.ms-office.vbaProject";

    /// VBA project signature content type.
    pub const VBA_PROJECT_SIGNATURE: &str = "application/vnd.ms-office.vbaProjectSignature";

    /// VBA data content type (used for VBA storage).
    pub const VBA_DATA: &str = "application/vnd.ms-office.vbaData";
}

/// VBA-related part name constants.
pub mod part_names {
    /// Default VBA project part name.
    pub const VBA_PROJECT: &str = "/ppt/vbaProject.bin";

    /// Default VBA project signature part name.
    pub const VBA_PROJECT_SIGNATURE: &str = "/ppt/_xmlsignatures/vbaProjectSignature.bin";
}

/// VBA relationship type constants.
pub mod relationship_types {
    /// VBA project relationship type.
    pub const VBA_PROJECT: &str =
        "http://schemas.microsoft.com/office/2006/relationships/vbaProject";
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn vba_project_roundtrip() {
        let binary = vec![0x50, 0x4B, 0x03, 0x04]; // Fake VBA binary
        let project = VbaProject::new(
            "rId1",
            "vbaProject.bin",
            content_types::VBA_PROJECT,
            binary.clone(),
        );

        assert_eq!(project.relationship_id(), "rId1");
        assert_eq!(project.part_name(), "vbaProject.bin");
        assert_eq!(project.content_type(), content_types::VBA_PROJECT);
        assert_eq!(project.binary_content(), binary.as_slice());
        assert!(!project.is_empty());
        assert_eq!(project.size(), 4);
    }

    #[test]
    fn vba_project_empty() {
        let project = VbaProject::new("rId1", "vbaProject.bin", content_types::VBA_PROJECT, vec![]);

        assert!(project.is_empty());
        assert_eq!(project.size(), 0);
    }

    #[test]
    fn vba_signature_roundtrip() {
        let sig_data = vec![0xAB, 0xCD, 0xEF];
        let signature = VbaSignature::new(
            "rId2",
            content_types::VBA_PROJECT_SIGNATURE,
            sig_data.clone(),
        );

        assert_eq!(signature.relationship_id(), "rId2");
        assert_eq!(
            signature.content_type(),
            content_types::VBA_PROJECT_SIGNATURE
        );
        assert_eq!(signature.binary_content(), sig_data.as_slice());
        assert_eq!(signature.size(), 3);
    }

    #[test]
    fn vba_macro_container_empty() {
        let container = VbaMacroContainer::new();

        assert!(!container.has_macros());
        assert!(container.project().is_none());
        assert!(container.signature().is_none());
        assert_eq!(container.total_size(), 0);
    }

    #[test]
    fn vba_macro_container_with_project() {
        let mut container = VbaMacroContainer::new();
        let binary = vec![0x01, 0x02, 0x03, 0x04, 0x05];
        let project = VbaProject::new("rId1", "vbaProject.bin", content_types::VBA_PROJECT, binary);

        container.set_project(Some(project));

        assert!(container.has_macros());
        assert!(container.project().is_some());
        assert_eq!(container.total_size(), 5);
    }

    #[test]
    fn vba_macro_container_with_signature() {
        let mut container = VbaMacroContainer::new();
        let project = VbaProject::new(
            "rId1",
            "vbaProject.bin",
            content_types::VBA_PROJECT,
            vec![0x01, 0x02],
        );
        let signature = VbaSignature::new(
            "rId2",
            content_types::VBA_PROJECT_SIGNATURE,
            vec![0xAB, 0xCD],
        );

        container.set_project(Some(project));
        container.set_signature(Some(signature));

        assert!(container.has_macros());
        assert!(container.signature().is_some());
        assert_eq!(container.total_size(), 4); // 2 + 2
    }

    #[test]
    fn vba_macro_container_additional_parts() {
        let mut container = VbaMacroContainer::new();
        container.add_additional_part("custom.bin", vec![0xFF, 0xFE]);

        assert_eq!(container.additional_parts().len(), 1);
        assert_eq!(container.total_size(), 2);

        let removed = container.remove_additional_part("custom.bin");
        assert!(removed.is_some());
        assert_eq!(removed.unwrap(), vec![0xFF, 0xFE]);
        assert_eq!(container.additional_parts().len(), 0);
    }

    #[test]
    fn vba_macro_container_clear() {
        let mut container = VbaMacroContainer::new();
        let project = VbaProject::new(
            "rId1",
            "vbaProject.bin",
            content_types::VBA_PROJECT,
            vec![0x01],
        );
        let signature = VbaSignature::new("rId2", content_types::VBA_PROJECT_SIGNATURE, vec![0x02]);

        container.set_project(Some(project));
        container.set_signature(Some(signature));
        container.add_additional_part("custom.bin", vec![0x03]);

        assert!(container.has_macros());
        assert_eq!(container.total_size(), 3);

        container.clear();

        assert!(!container.has_macros());
        assert!(container.project().is_none());
        assert!(container.signature().is_none());
        assert_eq!(container.additional_parts().len(), 0);
        assert_eq!(container.total_size(), 0);
    }

    #[test]
    fn vba_content_type_constants() {
        assert_eq!(
            content_types::VBA_PROJECT,
            "application/vnd.ms-office.vbaProject"
        );
        assert_eq!(
            content_types::VBA_PROJECT_SIGNATURE,
            "application/vnd.ms-office.vbaProjectSignature"
        );
        assert_eq!(content_types::VBA_DATA, "application/vnd.ms-office.vbaData");
    }

    #[test]
    fn vba_part_name_constants() {
        assert_eq!(part_names::VBA_PROJECT, "/ppt/vbaProject.bin");
        assert_eq!(
            part_names::VBA_PROJECT_SIGNATURE,
            "/ppt/_xmlsignatures/vbaProjectSignature.bin"
        );
    }

    #[test]
    fn vba_relationship_type_constants() {
        assert_eq!(
            relationship_types::VBA_PROJECT,
            "http://schemas.microsoft.com/office/2006/relationships/vbaProject"
        );
    }
}
