//! Part abstraction — a single entry in an OPC package.
//!
//! Parts can contain XML (parsed or raw) or binary data (images, etc.).
//! The key to roundtrip fidelity: parts we don't understand are stored
//! as raw bytes and written back unchanged.

use crate::relationship::Relationships;
use crate::uri::PartUri;

/// Data stored in a part.
#[derive(Debug, Clone)]
pub enum PartData {
    /// Raw bytes — used for binary parts (images, etc.) and XML parts
    /// we haven't parsed yet or don't need to modify.
    Raw(Vec<u8>),

    /// XML that has been parsed and potentially modified.
    /// Stored as raw bytes that will be re-serialized on save.
    Xml(Vec<u8>),
}

impl PartData {
    pub fn as_bytes(&self) -> &[u8] {
        match self {
            PartData::Raw(bytes) | PartData::Xml(bytes) => bytes,
        }
    }

    pub fn into_bytes(self) -> Vec<u8> {
        match self {
            PartData::Raw(bytes) | PartData::Xml(bytes) => bytes,
        }
    }

    pub fn len(&self) -> usize {
        self.as_bytes().len()
    }

    pub fn is_empty(&self) -> bool {
        self.as_bytes().is_empty()
    }
}

/// A single part within an OPC package.
#[derive(Debug, Clone)]
pub struct Part {
    /// The absolute URI of this part within the package.
    pub uri: PartUri,

    /// The content type of this part (e.g., "application/xml").
    pub content_type: Option<String>,

    /// The part's data.
    pub data: PartData,

    /// Relationships owned by this part (from its `.rels` file).
    pub relationships: Relationships,

    /// Whether this part has been modified since loading.
    pub dirty: bool,
}

impl Part {
    /// Create a new part with raw data.
    pub fn new(uri: PartUri, data: Vec<u8>) -> Self {
        Self {
            uri,
            content_type: None,
            data: PartData::Raw(data),
            relationships: Relationships::new(),
            dirty: false,
        }
    }

    /// Create a new part with XML data.
    pub fn new_xml(uri: PartUri, data: Vec<u8>) -> Self {
        Self {
            uri,
            content_type: None,
            data: PartData::Xml(data),
            relationships: Relationships::new(),
            dirty: false,
        }
    }

    /// Mark this part as modified.
    pub fn mark_dirty(&mut self) {
        self.dirty = true;
    }

    /// Get the data as a UTF-8 string (for XML parts).
    pub fn as_str(&self) -> Option<&str> {
        std::str::from_utf8(self.data.as_bytes()).ok()
    }

    /// Replace the part's data with new bytes (stream loading).
    ///
    /// This is the equivalent of `Part.FeedData()` in the Open XML SDK.
    /// Marks the part as dirty.
    pub fn feed_data(&mut self, data: Vec<u8>) {
        let is_xml = matches!(self.data, PartData::Xml(_));
        self.data = if is_xml {
            PartData::Xml(data)
        } else {
            PartData::Raw(data)
        };
        self.dirty = true;
    }

    /// Replace the part's data from a reader (stream loading).
    pub fn feed_data_from<R: std::io::Read>(&mut self, reader: &mut R) -> std::io::Result<()> {
        let mut data = Vec::new();
        reader.read_to_end(&mut data)?;
        self.feed_data(data);
        Ok(())
    }

    /// Get the size of the part data in bytes.
    pub fn size(&self) -> usize {
        self.data.len()
    }

    /// Whether this part contains XML data.
    pub fn is_xml(&self) -> bool {
        matches!(self.data, PartData::Xml(_))
    }

    /// Whether this part contains binary data.
    pub fn is_binary(&self) -> bool {
        matches!(self.data, PartData::Raw(_))
    }
}
