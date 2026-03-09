//! Package — the main entry point for reading and writing OPC packages.
//!
//! An OPC package is a ZIP archive containing parts (XML and binary),
//! relationships (how parts connect), and content types (what each part is).

use std::collections::{HashMap, HashSet, VecDeque};
use std::io::{BufReader, Cursor, Read, Seek, Write};
use std::path::Path;

use zip::read::ZipArchive;
use zip::write::FileOptions;
use zip::ZipWriter;

use crate::content_types::ContentTypes;
use crate::error::Result;
use crate::part::Part;
use crate::properties::{CoreProperties, CustomProperties, ExtendedProperties};
use crate::relationship::{RelationshipType, Relationships, TargetMode};
use crate::uri::{PartUri, CONTENT_TYPES_URI, PACKAGE_RELS_URI};

/// ZIP compression level for parts in a package.
///
/// Maps to `System.IO.Packaging.CompressionOption` in .NET.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum CompressionOption {
    /// Default compression balance between size and speed.
    #[default]
    Normal,
    /// Best compression ratio, slower.
    Maximum,
    /// Fast compression, larger output.
    Fast,
    /// Minimal compression, fastest.
    SuperFast,
    /// No compression — store raw bytes.
    NotCompressed,
}

impl CompressionOption {
    /// Convert to the corresponding `zip` crate compression method.
    pub fn to_zip_method(self) -> zip::CompressionMethod {
        match self {
            CompressionOption::Normal
            | CompressionOption::Maximum
            | CompressionOption::Fast
            | CompressionOption::SuperFast => zip::CompressionMethod::Deflated,
            CompressionOption::NotCompressed => zip::CompressionMethod::Stored,
        }
    }
}

/// Whether an OOXML package uses Strict or Transitional conformance.
///
/// ECMA-376 defines two conformance classes:
/// - **Transitional**: backward-compatible with older Office versions, allows legacy features.
/// - **Strict**: fully ISO/IEC 29500 conformant, different namespace URIs.
///
/// Most real-world files are Transitional. Strict is used by some non-Microsoft tools
/// and is required for certain government/regulatory workflows.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum DocumentConformance {
    /// Transitional conformance (ECMA-376 Transitional).
    #[default]
    Transitional,
    /// Strict conformance (ISO/IEC 29500 Strict).
    Strict,
}

/// An OPC package — represents an .xlsx, .docx, or .pptx file.
#[derive(Debug, Clone)]
pub struct Package {
    /// All parts in the package, keyed by their URI.
    parts: HashMap<String, Part>,

    /// Package-level content types.
    content_types: ContentTypes,

    /// Whether `[Content_Types].xml` existed in the loaded package.
    had_content_types_part: bool,

    /// Package-level relationships (from `/_rels/.rels`).
    relationships: Relationships,

    /// Relationship parts without a corresponding source part in `parts`.
    ///
    /// These can appear in malformed or partially-authored packages. We keep
    /// them so a no-op roundtrip does not drop ZIP entries.
    orphan_relationship_parts: HashMap<String, Vec<u8>>,

    /// Default compression level for newly written parts.
    compression_option: CompressionOption,

    /// Whether the package uses Strict or Transitional conformance.
    ///
    /// Detected during open by inspecting relationship namespace URIs.
    document_conformance: DocumentConformance,
}

impl Package {
    /// Create a new, empty package.
    pub fn new() -> Self {
        Self {
            parts: HashMap::new(),
            content_types: ContentTypes::new(),
            had_content_types_part: true,
            relationships: Relationships::new(),
            orphan_relationship_parts: HashMap::new(),
            compression_option: CompressionOption::Normal,
            document_conformance: DocumentConformance::default(),
        }
    }

    /// Open a package from a file path.
    pub fn open(path: impl AsRef<Path>) -> Result<Self> {
        let file = std::fs::File::open(path)?;
        let reader = BufReader::new(file);
        Self::from_reader(reader)
    }

    /// Open a package from bytes.
    pub fn from_bytes(bytes: &[u8]) -> Result<Self> {
        let cursor = Cursor::new(bytes);
        Self::from_reader(cursor)
    }

    /// Open a package from any reader.
    pub fn from_reader<R: Read + Seek>(reader: R) -> Result<Self> {
        let mut archive = ZipArchive::new(reader)?;
        let mut package = Package::new();
        package.had_content_types_part = false;
        package.compression_option = CompressionOption::Normal;

        // First pass: read all entries into parts
        let mut raw_parts: HashMap<String, Vec<u8>> = HashMap::new();
        for i in 0..archive.len() {
            let mut entry = archive.by_index(i)?;
            let name = entry.name().to_string();

            // Skip directories
            if entry.is_dir() {
                continue;
            }

            let mut data = Vec::with_capacity(entry.size() as usize);
            entry.read_to_end(&mut data)?;
            raw_parts.insert(name, data);
        }

        // Parse [Content_Types].xml
        if let Some(ct_data) = raw_parts.remove(CONTENT_TYPES_URI.trim_start_matches('/')) {
            package.content_types = ContentTypes::from_xml_bytes(ct_data)?;
            package.had_content_types_part = true;
        } else {
            // Preserve malformed inputs that omit [Content_Types].xml on no-op saves.
            package.content_types.mark_clean();
        }

        // Parse package-level relationships
        if let Some(rels_data) = raw_parts.remove(PACKAGE_RELS_URI.trim_start_matches('/')) {
            package.relationships = Relationships::from_xml_bytes(rels_data)?;
        }

        // Detect Strict vs Transitional conformance by checking relationship types
        for rel in package.relationships.iter() {
            if rel.rel_type.starts_with("http://purl.oclc.org/ooxml/") {
                package.document_conformance = DocumentConformance::Strict;
                break;
            }
        }

        // Process remaining parts
        let mut part_rels: HashMap<String, Vec<u8>> = HashMap::new();
        let mut regular_parts: HashMap<String, Vec<u8>> = HashMap::new();

        for (name, data) in raw_parts {
            if PartUri::is_part_relationship_zip_path(&name) {
                part_rels.insert(name, data);
            } else {
                regular_parts.insert(name, data);
            }
        }

        // Create parts with their relationships
        for (zip_path, data) in regular_parts {
            let uri = PartUri::from_zip_path(&zip_path)?;
            let content_type = package
                .content_types
                .get(uri.as_str())
                .map(|s| s.to_string());

            let is_xml = content_type
                .as_deref()
                .map(|ct| ct.contains("xml") || ct.contains("+xml"))
                .unwrap_or_else(|| {
                    uri.extension()
                        .map(|e| e == "xml" || e == "rels")
                        .unwrap_or(false)
                });

            let mut part = if is_xml {
                Part::new_xml(uri.clone(), data)
            } else {
                Part::new(uri.clone(), data)
            };

            part.content_type = content_type;

            // Look for this part's relationships
            let rels_zip_path = uri.relationship_zip_path();
            if let Some(rels_data) = part_rels.remove(&rels_zip_path) {
                part.relationships = Relationships::from_xml_bytes(rels_data)?;
            }

            package.parts.insert(uri.as_str().to_string(), part);
        }

        // Preserve orphan relationship parts so no-op saves don't drop them.
        package.orphan_relationship_parts = part_rels;

        Ok(package)
    }

    /// Save the package to a file path.
    pub fn save(&self, path: impl AsRef<Path>) -> Result<()> {
        let file = std::fs::File::create(path)?;
        self.write_to(file)
    }

    /// Save the package to bytes.
    pub fn to_bytes(&self) -> Result<Vec<u8>> {
        let mut buf = Vec::new();
        self.write_to(Cursor::new(&mut buf))?;
        Ok(buf)
    }

    /// Write the package to any writer.
    pub fn write_to<W: Write + Seek>(&self, writer: W) -> Result<()> {
        let mut zip = ZipWriter::new(writer);
        let options = FileOptions::<()>::default()
            .compression_method(self.compression_option.to_zip_method());

        // Write [Content_Types].xml
        if self.had_content_types_part || self.content_types.is_dirty() {
            zip.start_file(CONTENT_TYPES_URI.trim_start_matches('/'), options)?;
            self.content_types.to_xml(&mut zip)?;
        }

        // Write package-level relationships
        if self.relationships.should_write_xml() {
            zip.start_file(PACKAGE_RELS_URI.trim_start_matches('/'), options)?;
            self.relationships.to_xml(&mut zip)?;
        }

        // Write all parts (sorted for deterministic output)
        let mut part_uris: Vec<&String> = self.parts.keys().collect();
        part_uris.sort();

        for uri in part_uris {
            let part = &self.parts[uri];
            let zip_path = part.uri.to_zip_path();

            zip.start_file(zip_path, options)?;
            zip.write_all(part.data.as_bytes())?;

            // Write part relationships if any
            if part.relationships.should_write_xml() {
                let rels_path = part.uri.relationship_zip_path();
                zip.start_file(rels_path, options)?;
                part.relationships.to_xml(&mut zip)?;
            }
        }

        // Write orphan relationship parts (sorted for deterministic output).
        let mut orphan_paths: Vec<&String> = self.orphan_relationship_parts.keys().collect();
        orphan_paths.sort();
        for path in orphan_paths {
            zip.start_file(path, options)?;
            if let Some(bytes) = self.orphan_relationship_parts.get(path) {
                zip.write_all(bytes)?;
            }
        }

        zip.finish()?;
        Ok(())
    }

    // --- Accessors ---

    /// Get a part by URI.
    pub fn get_part(&self, uri: &str) -> Option<&Part> {
        self.parts.get(uri)
    }

    /// Get a mutable part by URI.
    pub fn get_part_mut(&mut self, uri: &str) -> Option<&mut Part> {
        self.parts.get_mut(uri)
    }

    /// Add or replace a part.
    pub fn set_part(&mut self, mut part: Part) {
        if let Some(ref ct) = part.content_type {
            self.content_types
                .add_override(part.uri.as_str(), ct.clone());
        }

        // If this part had an orphan .rels entry, attach it when possible.
        if part.relationships.is_empty() {
            let rels_zip_path = part.uri.relationship_zip_path();
            if let Some(rels_data) = self.orphan_relationship_parts.remove(&rels_zip_path) {
                if let Ok(rels) = Relationships::from_xml_bytes(rels_data) {
                    part.relationships = rels;
                }
            }
        }

        self.parts.insert(part.uri.as_str().to_string(), part);
    }

    /// Remove a part.
    pub fn remove_part(&mut self, uri: &str) -> Option<Part> {
        self.content_types.remove_override(uri);
        if let Ok(part_uri) = PartUri::new(uri) {
            self.orphan_relationship_parts
                .remove(&part_uri.relationship_zip_path());
        }
        self.parts.remove(uri)
    }

    /// Get package-level relationships.
    pub fn relationships(&self) -> &Relationships {
        &self.relationships
    }

    /// Get mutable package-level relationships.
    pub fn relationships_mut(&mut self) -> &mut Relationships {
        &mut self.relationships
    }

    /// Get content types.
    pub fn content_types(&self) -> &ContentTypes {
        &self.content_types
    }

    /// Get mutable content types.
    pub fn content_types_mut(&mut self) -> &mut ContentTypes {
        &mut self.content_types
    }

    /// Iterate all parts.
    pub fn parts(&self) -> impl Iterator<Item = &Part> {
        self.parts.values()
    }

    /// Get all part URIs.
    pub fn part_uris(&self) -> Vec<&str> {
        self.parts.keys().map(|s| s.as_str()).collect()
    }

    /// Check whether a part exists at the given URI.
    pub fn has_part(&self, uri: &str) -> bool {
        self.parts.contains_key(uri)
    }

    /// Number of parts in the package.
    pub fn part_count(&self) -> usize {
        self.parts.len()
    }

    /// Read core properties from `docProps/core.xml` if present.
    pub fn core_properties(&self) -> Result<Option<CoreProperties>> {
        // Find the core properties part via relationship or well-known path
        let core_uri = self
            .relationships
            .get_first_by_type(RelationshipType::CORE_PROPERTIES)
            .map(|r| {
                if r.target.starts_with('/') {
                    r.target.clone()
                } else {
                    format!("/{}", r.target)
                }
            })
            .unwrap_or_else(|| "/docProps/core.xml".to_string());

        match self.parts.get(&core_uri) {
            Some(part) => {
                let props = CoreProperties::from_xml_bytes(part.data.as_bytes().to_vec())?;
                Ok(Some(props))
            }
            None => Ok(None),
        }
    }

    /// Read extended properties from `docProps/app.xml` if present.
    pub fn extended_properties(&self) -> Result<Option<ExtendedProperties>> {
        let ext_uri = self
            .relationships
            .get_first_by_type(RelationshipType::EXTENDED_PROPERTIES)
            .map(|r| {
                if r.target.starts_with('/') {
                    r.target.clone()
                } else {
                    format!("/{}", r.target)
                }
            })
            .unwrap_or_else(|| "/docProps/app.xml".to_string());

        match self.parts.get(&ext_uri) {
            Some(part) => {
                let props = ExtendedProperties::from_xml_bytes(part.data.as_bytes().to_vec())?;
                Ok(Some(props))
            }
            None => Ok(None),
        }
    }

    /// Get the package-wide compression option.
    pub fn compression_option(&self) -> CompressionOption {
        self.compression_option
    }

    /// Set the package-wide compression option.
    pub fn set_compression_option(&mut self, option: CompressionOption) {
        self.compression_option = option;
    }

    /// Get the detected document conformance (Strict or Transitional).
    ///
    /// For packages opened from files, this is detected by inspecting
    /// relationship type URIs. Strict OOXML uses `http://purl.oclc.org/ooxml/`
    /// prefixed URIs instead of the standard `http://schemas.openxmlformats.org/`.
    pub fn document_conformance(&self) -> DocumentConformance {
        self.document_conformance
    }

    /// Whether this package uses Strict OOXML conformance.
    pub fn is_strict(&self) -> bool {
        self.document_conformance == DocumentConformance::Strict
    }

    /// Iterate all parts mutably.
    pub fn parts_mut(&mut self) -> impl Iterator<Item = &mut Part> {
        self.parts.values_mut()
    }

    /// Read custom properties from `docProps/custom.xml` if present.
    pub fn custom_properties(&self) -> Result<Option<CustomProperties>> {
        let custom_uri = self
            .relationships
            .get_first_by_type(RelationshipType::CUSTOM_PROPERTIES)
            .map(|r| {
                if r.target.starts_with('/') {
                    r.target.clone()
                } else {
                    format!("/{}", r.target)
                }
            })
            .unwrap_or_else(|| "/docProps/custom.xml".to_string());

        match self.parts.get(&custom_uri) {
            Some(part) => {
                let props = CustomProperties::from_xml_bytes(part.data.as_bytes().to_vec())?;
                Ok(Some(props))
            }
            None => Ok(None),
        }
    }

    /// Delete a part and recursively remove all parts it references.
    ///
    /// This follows the Open XML SDK's cascade-delete pattern: when you remove
    /// a part, all parts referenced only through that part's relationships are
    /// also removed. Package-level relationships pointing to the deleted part
    /// are cleaned up as well.
    pub fn delete_part_recursive(&mut self, uri: &str) -> usize {
        let mut to_delete = vec![uri.to_string()];
        let mut deleted_count = 0;

        while let Some(current_uri) = to_delete.pop() {
            if let Some(part) = self.parts.remove(&current_uri) {
                self.content_types.remove_override(&current_uri);
                if let Ok(part_uri) = PartUri::new(&current_uri) {
                    self.orphan_relationship_parts
                        .remove(&part_uri.relationship_zip_path());
                }
                deleted_count += 1;

                // Queue all internal targets for deletion
                for rel in part.relationships.iter() {
                    if rel.target_mode == TargetMode::Internal {
                        let target_uri = if rel.target.starts_with('/') {
                            rel.target.clone()
                        } else if let Ok(part_uri) = PartUri::new(&current_uri) {
                            match part_uri.resolve_relative(&rel.target) {
                                Ok(resolved) => resolved.as_str().to_string(),
                                Err(_) => continue,
                            }
                        } else {
                            continue;
                        };

                        // Only delete if the part is not referenced by other remaining parts
                        if self.parts.contains_key(&target_uri) {
                            to_delete.push(target_uri);
                        }
                    }
                }
            }
        }

        // Clean up package-level relationships pointing to deleted parts
        let orphaned_rels: Vec<String> = self
            .relationships
            .iter()
            .filter(|r| {
                if r.target_mode == TargetMode::External {
                    return false;
                }
                let target = if r.target.starts_with('/') {
                    r.target.clone()
                } else {
                    format!("/{}", r.target)
                };
                !self.parts.contains_key(&target)
            })
            .map(|r| r.id.clone())
            .collect();

        for id in orphaned_rels {
            self.relationships.remove_by_id(&id);
        }

        deleted_count
    }

    /// Get all parts reachable from the root via BFS relationship traversal.
    ///
    /// Starting from the package-level relationships, follows all internal
    /// relationship targets and their sub-relationships recursively.
    /// Returns part URIs in BFS visit order.
    pub fn get_all_parts_bfs(&self) -> Vec<&str> {
        let mut visited = HashSet::new();
        let mut queue = VecDeque::new();
        let mut result = Vec::new();

        // Seed with package-level relationship targets
        for rel in self.relationships.iter() {
            if rel.target_mode == TargetMode::Internal {
                let uri = if rel.target.starts_with('/') {
                    rel.target.clone()
                } else {
                    format!("/{}", rel.target)
                };
                if !visited.contains(&uri) {
                    visited.insert(uri.clone());
                    queue.push_back(uri);
                }
            }
        }

        while let Some(uri) = queue.pop_front() {
            if let Some(part) = self.parts.get(&uri) {
                result.push(part.uri.as_str());

                // Follow this part's relationships
                for rel in part.relationships.iter() {
                    if rel.target_mode == TargetMode::Internal {
                        let target_uri = if rel.target.starts_with('/') {
                            rel.target.clone()
                        } else if let Ok(part_uri) = PartUri::new(&uri) {
                            match part_uri.resolve_relative(&rel.target) {
                                Ok(resolved) => resolved.as_str().to_string(),
                                Err(_) => continue,
                            }
                        } else {
                            continue;
                        };

                        if !visited.contains(&target_uri) {
                            visited.insert(target_uri.clone());
                            queue.push_back(target_uri);
                        }
                    }
                }
            }
        }

        result
    }

    /// Get all parts that reference the given part URI via a relationship.
    ///
    /// Performs a reverse lookup: for each part in the package, checks if
    /// any of its relationships point to `target_uri`.
    pub fn get_parent_parts(&self, target_uri: &str) -> Vec<&str> {
        let mut parents = Vec::new();
        let target_normalized = if target_uri.starts_with('/') {
            target_uri.to_string()
        } else {
            format!("/{target_uri}")
        };

        // Check package-level relationships
        for rel in self.relationships.iter() {
            if rel.target_mode == TargetMode::Internal {
                let resolved = if rel.target.starts_with('/') {
                    rel.target.clone()
                } else {
                    format!("/{}", rel.target)
                };
                if resolved == target_normalized {
                    // Package itself is a parent — represented as empty string
                    // but we skip since package isn't a "part"
                }
            }
        }

        // Check all part-level relationships
        for (uri, part) in &self.parts {
            for rel in part.relationships.iter() {
                if rel.target_mode == TargetMode::Internal {
                    let resolved = if rel.target.starts_with('/') {
                        rel.target.clone()
                    } else if let Ok(part_uri) = PartUri::new(uri) {
                        match part_uri.resolve_relative(&rel.target) {
                            Ok(r) => r.as_str().to_string(),
                            Err(_) => continue,
                        }
                    } else {
                        continue;
                    };

                    if resolved == target_normalized {
                        parents.push(part.uri.as_str());
                        break; // Only add each parent once
                    }
                }
            }
        }

        parents
    }

    /// Check if the given bytes represent an OLE compound file (encrypted OOXML).
    ///
    /// Encrypted Office documents are stored as OLE compound files with the
    /// magic bytes `D0 CF 11 E0 A1 B1 1A E1` at the start.
    pub fn is_encrypted_package(bytes: &[u8]) -> bool {
        const OLE_SIGNATURE: [u8; 8] = [0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1];
        bytes.len() >= 8 && bytes[..8] == OLE_SIGNATURE
    }

    /// Check if a file at the given path is an encrypted OLE compound file.
    pub fn is_encrypted_file(path: impl AsRef<Path>) -> std::io::Result<bool> {
        let mut file = std::fs::File::open(path)?;
        let mut header = [0u8; 8];
        match file.read_exact(&mut header) {
            Ok(()) => Ok(Self::is_encrypted_package(&header)),
            Err(e) if e.kind() == std::io::ErrorKind::UnexpectedEof => Ok(false),
            Err(e) => Err(e),
        }
    }

    /// Get a part by looking up a relationship ID in the package-level relationships.
    ///
    /// Resolves the relationship target to a part URI, then returns the part.
    /// This mirrors Open XML SDK's `GetPartById()`.
    pub fn get_part_by_rel_id(&self, rel_id: &str) -> Option<&Part> {
        let rel = self.relationships.get_by_id(rel_id)?;
        if rel.target_mode == TargetMode::External {
            return None;
        }
        let uri = if rel.target.starts_with('/') {
            rel.target.clone()
        } else {
            format!("/{}", rel.target)
        };
        self.parts.get(&uri)
    }

    /// Get the relationship ID that points to the given part URI.
    ///
    /// Searches the package-level relationships for one whose target matches
    /// the given URI. This mirrors Open XML SDK's `GetIdOfPart()`.
    pub fn get_id_of_part(&self, part_uri: &str) -> Option<&str> {
        let normalized = if part_uri.starts_with('/') {
            part_uri.to_string()
        } else {
            format!("/{part_uri}")
        };

        for rel in self.relationships.iter() {
            if rel.target_mode == TargetMode::External {
                continue;
            }
            let target = if rel.target.starts_with('/') {
                &rel.target
            } else {
                // Relative to package root = prepend /
                &format!("/{}", rel.target)
            };
            if *target == normalized {
                return Some(&rel.id);
            }
        }
        None
    }

    /// Clone a part from this package into another package.
    ///
    /// Copies the part data, content type, and relationships.
    /// Returns `true` if the part was found and copied.
    pub fn clone_part_to(&self, uri: &str, target_package: &mut Package) -> bool {
        if let Some(part) = self.parts.get(uri) {
            target_package.set_part(part.clone());
            true
        } else {
            false
        }
    }

    /// Remove all parts whose URI starts with any of the given prefixes.
    /// Also removes corresponding content-type overrides.
    pub fn remove_parts_by_prefix(&mut self, prefixes: &[&str]) -> usize {
        let uris: Vec<String> = self
            .parts
            .keys()
            .filter(|uri| prefixes.iter().any(|p| uri.starts_with(p)))
            .cloned()
            .collect();
        let count = uris.len();
        for uri in &uris {
            self.parts.remove(uri);
            self.content_types.remove_override(uri);
            if let Ok(part_uri) = PartUri::new(uri) {
                self.orphan_relationship_parts
                    .remove(&part_uri.relationship_zip_path());
            }
        }
        count
    }
}

impl Default for Package {
    fn default() -> Self {
        Self::new()
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::relationship::TargetMode;
    use std::collections::BTreeMap;

    fn relationship_xml(target: &str) -> Vec<u8> {
        let mut relationships = Relationships::new();
        relationships.add_new(
            "http://example.com/relationship".to_string(),
            target.to_string(),
            TargetMode::Internal,
        );

        let mut xml = Vec::new();
        relationships.to_xml(&mut xml).unwrap();
        xml
    }

    fn zip_entries(bytes: &[u8]) -> BTreeMap<String, Vec<u8>> {
        let mut archive = ZipArchive::new(Cursor::new(bytes)).expect("zip should parse");
        let mut entries = BTreeMap::new();
        for index in 0..archive.len() {
            let mut entry = archive.by_index(index).expect("zip entry should open");
            if entry.is_dir() {
                continue;
            }
            let mut data = Vec::new();
            entry
                .read_to_end(&mut data)
                .expect("zip entry bytes should read");
            entries.insert(entry.name().to_string(), data);
        }
        entries
    }

    #[test]
    fn test_from_reader_detects_root_and_nested_relationship_parts() {
        let mut bytes = Vec::new();
        {
            let mut zip = ZipWriter::new(Cursor::new(&mut bytes));
            let options = FileOptions::<()>::default();

            zip.start_file("root.xml", options).unwrap();
            zip.write_all(b"<root/>").unwrap();

            zip.start_file("_rels/root.xml.rels", options).unwrap();
            zip.write_all(&relationship_xml("child.xml")).unwrap();

            zip.start_file("xl/workbook.xml", options).unwrap();
            zip.write_all(b"<workbook/>").unwrap();

            zip.start_file("xl/_rels/workbook.xml.rels", options)
                .unwrap();
            zip.write_all(&relationship_xml("worksheets/sheet1.xml"))
                .unwrap();

            zip.finish().unwrap();
        }

        let package = Package::from_bytes(&bytes).unwrap();

        let root = package.get_part("/root.xml").unwrap();
        assert_eq!(root.relationships.len(), 1);

        let workbook = package.get_part("/xl/workbook.xml").unwrap();
        assert_eq!(workbook.relationships.len(), 1);
    }

    #[test]
    fn test_root_level_relationship_roundtrip_and_zip_path() {
        let mut package = Package::new();

        let mut part = Part::new_xml(PartUri::new("/root.xml").unwrap(), b"<root/>".to_vec());
        part.relationships.add_new(
            "http://example.com/relationship".to_string(),
            "child.xml".to_string(),
            TargetMode::Internal,
        );
        package.set_part(part);

        let bytes = package.to_bytes().unwrap();

        let mut archive = ZipArchive::new(Cursor::new(bytes.as_slice())).unwrap();
        let mut names = Vec::new();
        for i in 0..archive.len() {
            let entry = archive.by_index(i).unwrap();
            names.push(entry.name().to_string());
        }

        assert!(names.iter().any(|name| name == "_rels/root.xml.rels"));
        assert!(!names.iter().any(|name| name == "/_rels/root.xml.rels"));

        let loaded = Package::from_bytes(&bytes).unwrap();
        let loaded_part = loaded.get_part("/root.xml").unwrap();
        assert_eq!(loaded_part.relationships.len(), 1);
    }

    #[test]
    fn test_clone_preserves_all_parts_and_relationships() {
        let mut package = Package::new();

        let mut part = Part::new_xml(PartUri::new("/root.xml").unwrap(), b"<root/>".to_vec());
        part.content_type = Some("application/xml".to_string());
        part.relationships.add_new(
            "http://example.com/relationship".to_string(),
            "child.xml".to_string(),
            TargetMode::Internal,
        );
        package.set_part(part);

        let part2 = Part::new(PartUri::new("/media/image.png").unwrap(), vec![1, 2, 3]);
        package.set_part(part2);

        package.relationships_mut().add_new(
            "http://example.com/root".to_string(),
            "/root.xml".to_string(),
            TargetMode::Internal,
        );

        let cloned = package.clone();

        assert_eq!(cloned.part_uris().len(), package.part_uris().len());
        assert!(cloned.get_part("/root.xml").is_some());
        assert!(cloned.get_part("/media/image.png").is_some());
        assert_eq!(cloned.relationships().len(), package.relationships().len());
        assert_eq!(cloned.get_part("/root.xml").unwrap().relationships.len(), 1);
    }

    #[test]
    fn test_remove_parts_by_prefix() {
        let mut package = Package::new();

        for uri in &[
            "/xl/workbook.xml",
            "/xl/worksheets/sheet1.xml",
            "/xl/worksheets/sheet2.xml",
            "/xl/theme/theme1.xml",
            "/xl/styles.xml",
            "/docProps/core.xml",
        ] {
            let mut part = Part::new_xml(PartUri::new(*uri).unwrap(), b"<x/>".to_vec());
            part.content_type = Some("application/xml".to_string());
            package.set_part(part);
        }

        assert_eq!(package.part_uris().len(), 6);

        let removed = package.remove_parts_by_prefix(&["/xl/worksheets/", "/xl/styles.xml"]);
        assert_eq!(removed, 3);
        assert_eq!(package.part_uris().len(), 3);

        // These should survive
        assert!(package.get_part("/xl/workbook.xml").is_some());
        assert!(package.get_part("/xl/theme/theme1.xml").is_some());
        assert!(package.get_part("/docProps/core.xml").is_some());

        // These should be removed
        assert!(package.get_part("/xl/worksheets/sheet1.xml").is_none());
        assert!(package.get_part("/xl/worksheets/sheet2.xml").is_none());
        assert!(package.get_part("/xl/styles.xml").is_none());

        // Content type overrides should also be removed
        assert!(package
            .content_types()
            .get_override("/xl/worksheets/sheet1.xml")
            .is_none());
        assert!(package
            .content_types()
            .get_override("/xl/styles.xml")
            .is_none());
    }

    #[test]
    fn test_remove_parts_by_prefix_empty() {
        let mut package = Package::new();
        let part = Part::new_xml(PartUri::new("/test.xml").unwrap(), b"<x/>".to_vec());
        package.set_part(part);

        let removed = package.remove_parts_by_prefix(&["/nonexistent/"]);
        assert_eq!(removed, 0);
        assert_eq!(package.part_uris().len(), 1);
    }

    #[test]
    fn no_op_roundtrip_preserves_original_content_types_and_relationship_xml() {
        let content_types = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default ContentType="application/xml" Extension="xml"/>
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Override ContentType="application/custom+xml" PartName="/root.xml"/>
</Types>"#;
        let package_rels = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Target="/root.xml" Type="http://example.com/root" Id="rId9"/>
</Relationships>"#;
        let part_rels = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Type="http://example.com/child" Id="rId3" Target="child.xml"/>
</Relationships>"#;

        let mut input = Vec::new();
        {
            let mut zip = ZipWriter::new(Cursor::new(&mut input));
            let options = FileOptions::<()>::default();
            zip.start_file("[Content_Types].xml", options).unwrap();
            zip.write_all(content_types).unwrap();
            zip.start_file("_rels/.rels", options).unwrap();
            zip.write_all(package_rels).unwrap();
            zip.start_file("root.xml", options).unwrap();
            zip.write_all(b"<root/>").unwrap();
            zip.start_file("_rels/root.xml.rels", options).unwrap();
            zip.write_all(part_rels).unwrap();
            zip.finish().unwrap();
        }

        let package = Package::from_bytes(&input).expect("package should open");
        let output = package.to_bytes().expect("package should save");
        let entries = zip_entries(output.as_slice());

        assert_eq!(
            entries.get("[Content_Types].xml").map(Vec::as_slice),
            Some(content_types.as_slice()),
            "no-op roundtrip should preserve original [Content_Types].xml bytes",
        );
        assert_eq!(
            entries.get("_rels/.rels").map(Vec::as_slice),
            Some(package_rels.as_slice()),
            "no-op roundtrip should preserve original package .rels bytes",
        );
        assert_eq!(
            entries.get("_rels/root.xml.rels").map(Vec::as_slice),
            Some(part_rels.as_slice()),
            "no-op roundtrip should preserve original part .rels bytes",
        );
    }

    #[test]
    fn no_op_roundtrip_preserves_empty_relationship_parts() {
        let content_types = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default ContentType="application/xml" Extension="xml"/>
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Override ContentType="application/custom+xml" PartName="/root.xml"/>
</Types>"#;
        let package_rels = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>"#;
        let part_rels = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>"#;

        let mut input = Vec::new();
        {
            let mut zip = ZipWriter::new(Cursor::new(&mut input));
            let options = FileOptions::<()>::default();
            zip.start_file("[Content_Types].xml", options).unwrap();
            zip.write_all(content_types).unwrap();
            zip.start_file("_rels/.rels", options).unwrap();
            zip.write_all(package_rels).unwrap();
            zip.start_file("root.xml", options).unwrap();
            zip.write_all(b"<root/>").unwrap();
            zip.start_file("_rels/root.xml.rels", options).unwrap();
            zip.write_all(part_rels).unwrap();
            zip.finish().unwrap();
        }

        let package = Package::from_bytes(&input).expect("package should open");
        let output = package.to_bytes().expect("package should save");
        let entries = zip_entries(output.as_slice());

        assert_eq!(
            entries.get("_rels/.rels").map(Vec::as_slice),
            Some(package_rels.as_slice()),
            "empty package .rels should be preserved",
        );
        assert_eq!(
            entries.get("_rels/root.xml.rels").map(Vec::as_slice),
            Some(part_rels.as_slice()),
            "empty part .rels should be preserved",
        );
    }

    #[test]
    fn no_op_roundtrip_preserves_orphan_relationship_entries() {
        let content_types = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/xml"/>
</Types>"#;
        let orphan_rels = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://example.com/sheet" Target="../worksheets/sheet1.xml"/>
</Relationships>"#;

        let mut input = Vec::new();
        {
            let mut zip = ZipWriter::new(Cursor::new(&mut input));
            let options = FileOptions::<()>::default();
            zip.start_file("[Content_Types].xml", options).unwrap();
            zip.write_all(content_types).unwrap();
            zip.start_file("xl/workbook.xml", options).unwrap();
            zip.write_all(b"<workbook/>").unwrap();
            zip.start_file("xl/worksheets/_rels/sheet1.xml.rels", options)
                .unwrap();
            zip.write_all(orphan_rels).unwrap();
            zip.finish().unwrap();
        }

        let package = Package::from_bytes(&input).expect("package should open");
        let output = package.to_bytes().expect("package should save");
        let entries = zip_entries(output.as_slice());

        assert_eq!(
            entries
                .get("xl/worksheets/_rels/sheet1.xml.rels")
                .map(Vec::as_slice),
            Some(orphan_rels.as_slice()),
            "orphan .rels entry should be preserved on no-op roundtrip",
        );
    }

    #[test]
    fn set_part_attaches_matching_orphan_relationship_entry() {
        let content_types = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/xml"/>
</Types>"#;
        let orphan_rels = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId7" Type="http://example.com/sheet" Target="../worksheets/sheet1.xml"/>
</Relationships>"#;

        let mut input = Vec::new();
        {
            let mut zip = ZipWriter::new(Cursor::new(&mut input));
            let options = FileOptions::<()>::default();
            zip.start_file("[Content_Types].xml", options).unwrap();
            zip.write_all(content_types).unwrap();
            zip.start_file("xl/workbook.xml", options).unwrap();
            zip.write_all(b"<workbook/>").unwrap();
            zip.start_file("xl/worksheets/_rels/sheet1.xml.rels", options)
                .unwrap();
            zip.write_all(orphan_rels).unwrap();
            zip.finish().unwrap();
        }

        let mut package = Package::from_bytes(&input).expect("package should open");
        let part = Part::new_xml(
            PartUri::new("/xl/worksheets/sheet1.xml").expect("valid part URI"),
            br#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"/>"#
                .to_vec(),
        );
        package.set_part(part);

        let attached = package
            .get_part("/xl/worksheets/sheet1.xml")
            .expect("sheet part should exist");
        assert_eq!(attached.relationships.len(), 1);
        assert!(attached.relationships.get_by_id("rId7").is_some());
    }

    #[test]
    fn remove_part_cleans_content_type_override_and_orphan_relationship_part() {
        let mut package = Package::new();
        let mut part = Part::new_xml(PartUri::new("/root.xml").unwrap(), b"<root/>".to_vec());
        part.content_type = Some("application/custom+xml".to_string());
        package.set_part(part);

        package.orphan_relationship_parts.insert(
            "_rels/root.xml.rels".to_string(),
            br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>"#
                .to_vec(),
        );

        let removed = package.remove_part("/root.xml");
        assert!(removed.is_some());
        assert!(
            package.content_types().get_override("/root.xml").is_none(),
            "removing part should remove override",
        );
        assert!(
            !package
                .orphan_relationship_parts
                .contains_key("_rels/root.xml.rels"),
            "removing part should remove matching orphan relationship entry",
        );
    }

    #[test]
    fn no_op_roundtrip_preserves_missing_content_types_manifest() {
        let mut input = Vec::new();
        {
            let mut zip = ZipWriter::new(Cursor::new(&mut input));
            let options = FileOptions::<()>::default();
            zip.start_file("root.xml", options).unwrap();
            zip.write_all(b"<root/>").unwrap();
            zip.finish().unwrap();
        }

        let package = Package::from_bytes(&input).expect("package should open");
        let output = package.to_bytes().expect("package should save");
        let entries = zip_entries(output.as_slice());

        assert!(
            !entries.contains_key("[Content_Types].xml"),
            "missing manifest should stay missing on no-op malformed-package roundtrip",
        );
    }

    #[test]
    fn get_part_by_rel_id_resolves_internal_relationship() {
        let mut package = Package::new();
        let part = Part::new_xml(PartUri::new("/xl/workbook.xml").unwrap(), b"<wb/>".to_vec());
        package.set_part(part);
        package.relationships_mut().add_new(
            RelationshipType::WORKBOOK.to_string(),
            "/xl/workbook.xml".to_string(),
            TargetMode::Internal,
        );

        let found = package.get_part_by_rel_id("rId1");
        assert!(found.is_some());
        assert_eq!(found.unwrap().uri.as_str(), "/xl/workbook.xml");
    }

    #[test]
    fn get_part_by_rel_id_returns_none_for_external() {
        let mut package = Package::new();
        package.relationships_mut().add_new(
            RelationshipType::HYPERLINK.to_string(),
            "https://example.com".to_string(),
            TargetMode::External,
        );

        assert!(package.get_part_by_rel_id("rId1").is_none());
    }

    #[test]
    fn get_part_by_rel_id_returns_none_for_missing_id() {
        let package = Package::new();
        assert!(package.get_part_by_rel_id("rId999").is_none());
    }

    #[test]
    fn get_part_by_rel_id_resolves_relative_target() {
        let mut package = Package::new();
        let part = Part::new_xml(PartUri::new("/xl/workbook.xml").unwrap(), b"<wb/>".to_vec());
        package.set_part(part);
        package.relationships_mut().add_new(
            RelationshipType::WORKBOOK.to_string(),
            "xl/workbook.xml".to_string(), // relative (no leading /)
            TargetMode::Internal,
        );

        let found = package.get_part_by_rel_id("rId1");
        assert!(found.is_some());
        assert_eq!(found.unwrap().uri.as_str(), "/xl/workbook.xml");
    }

    #[test]
    fn get_id_of_part_finds_relationship_id() {
        let mut package = Package::new();
        let part = Part::new_xml(PartUri::new("/xl/workbook.xml").unwrap(), b"<wb/>".to_vec());
        package.set_part(part);
        package.relationships_mut().add_new(
            RelationshipType::WORKBOOK.to_string(),
            "/xl/workbook.xml".to_string(),
            TargetMode::Internal,
        );

        assert_eq!(package.get_id_of_part("/xl/workbook.xml"), Some("rId1"));
    }

    #[test]
    fn get_id_of_part_normalizes_uri() {
        let mut package = Package::new();
        package.relationships_mut().add_new(
            RelationshipType::WORKBOOK.to_string(),
            "xl/workbook.xml".to_string(), // relative
            TargetMode::Internal,
        );

        // Should find even when querying with absolute path
        assert_eq!(package.get_id_of_part("/xl/workbook.xml"), Some("rId1"));
        // And without leading slash
        assert_eq!(package.get_id_of_part("xl/workbook.xml"), Some("rId1"));
    }

    #[test]
    fn get_id_of_part_returns_none_for_missing() {
        let package = Package::new();
        assert!(package.get_id_of_part("/nonexistent.xml").is_none());
    }

    #[test]
    fn get_id_of_part_skips_external_relationships() {
        let mut package = Package::new();
        package.relationships_mut().add_new(
            RelationshipType::HYPERLINK.to_string(),
            "https://example.com".to_string(),
            TargetMode::External,
        );

        assert!(package.get_id_of_part("https://example.com").is_none());
    }

    #[test]
    fn default_conformance_is_transitional() {
        let package = Package::new();
        assert_eq!(
            package.document_conformance(),
            DocumentConformance::Transitional
        );
        assert!(!package.is_strict());
    }

    #[test]
    fn strict_conformance_detected_from_relationships() {
        let strict_rels = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://purl.oclc.org/ooxml/officeDocument/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>"#;

        let mut input = Vec::new();
        {
            let mut zip = ZipWriter::new(Cursor::new(&mut input));
            let options = FileOptions::<()>::default();
            zip.start_file("[Content_Types].xml", options).unwrap();
            zip.write_all(
                br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="xml" ContentType="application/xml"/>
</Types>"#,
            )
            .unwrap();
            zip.start_file("_rels/.rels", options).unwrap();
            zip.write_all(strict_rels).unwrap();
            zip.start_file("xl/workbook.xml", options).unwrap();
            zip.write_all(b"<workbook/>").unwrap();
            zip.finish().unwrap();
        }

        let package = Package::from_bytes(&input).unwrap();
        assert_eq!(package.document_conformance(), DocumentConformance::Strict);
        assert!(package.is_strict());
    }

    #[test]
    fn transitional_conformance_detected_from_relationships() {
        let transitional_rels = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>"#;

        let mut input = Vec::new();
        {
            let mut zip = ZipWriter::new(Cursor::new(&mut input));
            let options = FileOptions::<()>::default();
            zip.start_file("[Content_Types].xml", options).unwrap();
            zip.write_all(
                br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="xml" ContentType="application/xml"/>
</Types>"#,
            )
            .unwrap();
            zip.start_file("_rels/.rels", options).unwrap();
            zip.write_all(transitional_rels).unwrap();
            zip.start_file("xl/workbook.xml", options).unwrap();
            zip.write_all(b"<workbook/>").unwrap();
            zip.finish().unwrap();
        }

        let package = Package::from_bytes(&input).unwrap();
        assert_eq!(
            package.document_conformance(),
            DocumentConformance::Transitional
        );
        assert!(!package.is_strict());
    }

    #[test]
    fn mutating_package_with_missing_content_types_writes_manifest() {
        let mut input = Vec::new();
        {
            let mut zip = ZipWriter::new(Cursor::new(&mut input));
            let options = FileOptions::<()>::default();
            zip.start_file("root.xml", options).unwrap();
            zip.write_all(b"<root/>").unwrap();
            zip.finish().unwrap();
        }

        let mut package = Package::from_bytes(&input).expect("package should open");
        let mut part = Part::new_xml(PartUri::new("/extra.xml").unwrap(), b"<extra/>".to_vec());
        part.content_type = Some("application/xml".to_string());
        package.set_part(part);

        let output = package.to_bytes().expect("package should save");
        let entries = zip_entries(output.as_slice());

        assert!(
            entries.contains_key("[Content_Types].xml"),
            "mutated package should emit content types manifest",
        );
    }
}
