//! Core CRDT document model.
//!
//! Wraps a [`yrs::Doc`] with a typed interface for document editing.
//! The CRDT is the single source of truth during editing sessions.

use std::collections::{HashMap, HashSet};
use yrs::types::array::ArrayRef;
use yrs::types::map::MapRef;
use yrs::{Doc, OffsetKind, Options};

use super::para_id::ParaId;

/// Errors from CRDT document operations.
#[derive(Debug, thiserror::Error)]
pub enum CrdtDocError {
    /// Failed to create or access a shared type.
    #[error("failed to access shared type '{0}'")]
    SharedType(String),
    /// Failed to read a value from the CRDT.
    #[error("missing expected value: {0}")]
    MissingValue(String),
    /// Transaction error.
    #[error("transaction error: {0}")]
    Transaction(String),
}

/// Result type for CRDT document operations.
pub type Result<T> = std::result::Result<T, CrdtDocError>;

/// CRDT-backed document model.
///
/// Wraps a `yrs::Doc` with a typed interface matching the document structure.
/// The CRDT is the single source of truth during editing.
///
/// # Shared types
///
/// The document exposes several named shared types:
/// - `"body"` ([`ArrayRef`]) -- the ordered list of paragraphs
/// - `"images"` ([`MapRef`]) -- image metadata keyed by content hash
/// - `"footnotes"` ([`ArrayRef`]) -- footnote content
/// - `"endnotes"` ([`ArrayRef`]) -- endnote content
/// - `"styles"` ([`MapRef`]) -- style definitions
///
/// All shared types are eagerly initialized in [`CrdtDoc::new`] and cached
/// as stored refs, so accessor methods never open an internal write transaction.
/// This means they are safe to call while a read transaction is active.
pub struct CrdtDoc {
    /// The yrs document.
    doc: Doc,
    /// Cached body array ref.
    body: ArrayRef,
    /// Cached images map ref.
    images: MapRef,
    /// Cached footnotes array ref.
    footnotes: ArrayRef,
    /// Cached endnotes array ref.
    endnotes: ArrayRef,
    /// Cached styles map ref.
    styles: MapRef,
    /// Mapping from app-level ParaId to original paragraph index.
    /// Used for dirty tracking during export.
    para_index_map: HashMap<ParaId, usize>,
    /// Image binary blobs, keyed by content hash.
    /// These are NOT stored in the CRDT (too large for replication).
    image_blobs: HashMap<String, Vec<u8>>,
    /// Set of dirty paragraph IDs (modified since last export).
    dirty_paragraphs: HashSet<ParaId>,
}

impl CrdtDoc {
    /// Create a new empty CRDT document with shared types initialized.
    ///
    /// All shared types (`body`, `images`, `footnotes`, `endnotes`, `styles`)
    /// are eagerly created and cached so that accessors never need to open
    /// an internal write transaction.
    pub fn new() -> Self {
        let doc = Doc::with_options(Options {
            // Browser selection/input offsets are UTF-16 code units.
            // Keep Yrs in the same unit to avoid offset conversion bugs.
            offset_kind: OffsetKind::Utf16,
            ..Options::default()
        });
        // Initialize all shared types eagerly. Each call internally creates
        // a short-lived transact_mut, which is fine since no other
        // transaction is active yet.
        let body = doc.get_or_insert_array("body");
        let images = doc.get_or_insert_map("images");
        let footnotes = doc.get_or_insert_array("footnotes");
        let endnotes = doc.get_or_insert_array("endnotes");
        let styles = doc.get_or_insert_map("styles");

        Self {
            doc,
            body,
            images,
            footnotes,
            endnotes,
            styles,
            para_index_map: HashMap::new(),
            image_blobs: HashMap::new(),
            dirty_paragraphs: HashSet::new(),
        }
    }

    /// Access the inner yrs [`Doc`].
    pub fn doc(&self) -> &Doc {
        &self.doc
    }

    /// Get the body array (ordered list of paragraphs).
    ///
    /// Safe to call while a read transaction is active.
    pub fn body(&self) -> ArrayRef {
        self.body.clone()
    }

    /// Get the images metadata map.
    ///
    /// Safe to call while a read transaction is active.
    pub fn images_map(&self) -> MapRef {
        self.images.clone()
    }

    /// Get the footnotes array.
    ///
    /// Safe to call while a read transaction is active.
    pub fn footnotes(&self) -> ArrayRef {
        self.footnotes.clone()
    }

    /// Get the endnotes array.
    ///
    /// Safe to call while a read transaction is active.
    pub fn endnotes(&self) -> ArrayRef {
        self.endnotes.clone()
    }

    /// Get the styles map.
    ///
    /// Safe to call while a read transaction is active.
    pub fn styles(&self) -> MapRef {
        self.styles.clone()
    }

    /// Get the paragraph-to-original-index mapping.
    ///
    /// Used during export to determine which paragraphs are new
    /// vs. modified from the original document.
    pub fn para_index_map(&self) -> &HashMap<ParaId, usize> {
        &self.para_index_map
    }

    /// Get image blobs (keyed by content hash).
    pub fn image_blobs(&self) -> &HashMap<String, Vec<u8>> {
        &self.image_blobs
    }

    /// Get mutable access to image blobs.
    pub fn image_blobs_mut(&mut self) -> &mut HashMap<String, Vec<u8>> {
        &mut self.image_blobs
    }

    /// Get the set of paragraph IDs modified since last export.
    pub fn dirty_paragraphs(&self) -> &HashSet<ParaId> {
        &self.dirty_paragraphs
    }

    /// Mark a paragraph as dirty (modified since last export).
    pub fn mark_dirty(&mut self, id: &ParaId) {
        self.dirty_paragraphs.insert(id.clone());
    }

    /// Clear the dirty set after a successful export.
    pub fn clear_dirty(&mut self) {
        self.dirty_paragraphs.clear();
    }

    /// Register a paragraph mapping from app ID to original document index.
    ///
    /// Called during import to track which CRDT paragraphs correspond
    /// to which original XML paragraph indices.
    pub fn register_paragraph(&mut self, id: ParaId, original_index: usize) {
        self.para_index_map.insert(id, original_index);
    }

    /// Register a newly created paragraph (no original index).
    ///
    /// The paragraph is automatically marked as dirty since it has
    /// no corresponding original content.
    pub fn register_new_paragraph(&mut self, id: ParaId) {
        self.dirty_paragraphs.insert(id);
    }
}

impl Default for CrdtDoc {
    fn default() -> Self {
        Self::new()
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use yrs::{Array, Transact};

    #[test]
    fn new_doc_has_empty_body() {
        let crdt = CrdtDoc::new();
        let txn = crdt.doc().transact();
        assert_eq!(crdt.body().len(&txn), 0);
    }

    #[test]
    fn shared_types_accessible() {
        let crdt = CrdtDoc::new();
        let txn = crdt.doc().transact();
        // All shared types should be accessible without deadlock
        assert_eq!(crdt.body().len(&txn), 0);
        assert_eq!(crdt.footnotes().len(&txn), 0);
        assert_eq!(crdt.endnotes().len(&txn), 0);
    }

    #[test]
    fn register_and_lookup_paragraph() {
        let mut crdt = CrdtDoc::new();
        let id = ParaId::new();
        crdt.register_paragraph(id.clone(), 5);
        assert_eq!(crdt.para_index_map().get(&id), Some(&5));
    }

    #[test]
    fn dirty_tracking() {
        let mut crdt = CrdtDoc::new();
        let id = ParaId::new();
        assert!(crdt.dirty_paragraphs().is_empty());

        crdt.mark_dirty(&id);
        assert!(crdt.dirty_paragraphs().contains(&id));

        crdt.clear_dirty();
        assert!(crdt.dirty_paragraphs().is_empty());
    }

    #[test]
    fn new_paragraph_auto_dirty() {
        let mut crdt = CrdtDoc::new();
        let id = ParaId::new();
        crdt.register_new_paragraph(id.clone());
        assert!(crdt.dirty_paragraphs().contains(&id));
        // New paragraphs should NOT be in the index map
        assert!(crdt.para_index_map().get(&id).is_none());
    }

    #[test]
    fn image_blobs_crud() {
        let mut crdt = CrdtDoc::new();
        assert!(crdt.image_blobs().is_empty());

        crdt.image_blobs_mut()
            .insert("sha256:abc".to_string(), vec![1, 2, 3]);
        assert_eq!(crdt.image_blobs().get("sha256:abc"), Some(&vec![1, 2, 3]));
    }

    #[test]
    fn default_creates_empty_doc() {
        let crdt = CrdtDoc::default();
        let txn = crdt.doc().transact();
        assert_eq!(crdt.body().len(&txn), 0);
    }
}
