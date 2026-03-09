//! # offidized-opc
//!
//! Open Packaging Convention (OPC) implementation for OOXML files.
//!
//! This crate handles the ZIP-based package format shared by all OOXML documents
//! (.xlsx, .docx, .pptx). It provides:
//!
//! - Reading and writing ZIP packages
//! - Relationship parsing and management (`_rels/.rels`)
//! - Content type resolution (`[Content_Types].xml`)
//! - Part URI resolution
//! - Raw XML preservation for roundtrip fidelity
//!
//! ## Architecture
//!
//! An OOXML file is a ZIP archive containing:
//! - XML parts (the actual content)
//! - Relationship parts (`.rels` files describing how parts connect)
//! - A content types manifest (`[Content_Types].xml`)
//! - Optional binary parts (images, embedded objects)
//!
//! This crate models the package structure without knowing anything about
//! spreadsheets, documents, or presentations. Format-specific logic lives
//! in `offidized-xlsx`, `offidized-docx`, and `offidized-pptx`.

pub mod content_types;
pub mod data_part;
pub mod error;
pub mod flat_opc;
pub mod hyperlink;
pub mod media_reference;
pub mod open_settings;
pub mod package;
pub mod part;
pub mod part_extension;
pub mod part_uri_helper;
pub mod properties;
pub mod raw;
pub mod relationship;
pub mod uri;
pub mod xml_util;

pub use content_types::{ContentTypeValue, ContentTypes};
pub use data_part::{content_type_from_extension, DataPart, MediaCategory, MediaDataPart};
pub use error::OpcError;
pub use flat_opc::{from_flat_opc, to_flat_opc};
pub use hyperlink::HyperlinkRelationship;
pub use media_reference::{
    AudioReferenceRelationship, DataPartReferenceRelationship, VideoReferenceRelationship,
};
pub use open_settings::{CompatibilityLevel, MarkupCompatibilityProcessMode, OpenSettings};
pub use package::{CompressionOption, DocumentConformance, Package};
pub use part::{Part, PartData};
pub use part_extension::PartExtensionMap;
pub use part_uri_helper::PartUriHelper;
pub use properties::{
    CoreProperties, CustomProperties, CustomProperty, CustomPropertyValue, ExtendedProperties,
};
pub use raw::RawXmlNode;
pub use relationship::{Relationship, RelationshipType, Relationships, TargetMode};
pub use uri::PartUri;
