//! # offidized-codegen
//!
//! Code generator that reads OOXML schema definitions and produces Rust types.
//!
//! ## Strategy
//!
//! The OOXML spec defines hundreds of XML element types across three markup languages:
//! - SpreadsheetML (.xlsx) — worksheets, cells, styles, charts
//! - WordprocessingML (.docx) — paragraphs, runs, tables, sections
//! - PresentationML (.pptx) — slides, shapes, animations
//!
//! Plus shared schemas: DrawingML (shapes, images), Office Math, VML, etc.
//!
//! Rather than hand-writing Rust structs for each (the .NET SDK has ~669k lines of
//! generated code), we generate them from the same schema data Microsoft uses.
//!
//! ## Schema Sources
//!
//! Two options, both viable:
//!
//! 1. **JSON schema data** from .NET Open XML SDK (`data/` directory)
//!    - 128 part definitions, 155 schema definitions
//!    - Already parsed and structured, easy to consume
//!    - Licensed MIT (same as Open XML SDK)
//!
//! 2. **XSD schemas** from ECMA-376 spec
//!    - The canonical source
//!    - More verbose to parse but complete
//!
//! We use approach #1 (JSON) for pragmatism, same as ooxmlsdk (Rust) does.
//!
//! ## Generated Output
//!
//! For each XML element type, we generate:
//! - A Rust struct with typed fields for attributes and known children
//! - A `Vec<RawXmlNode>` field for unknown/unrecognized children (roundtrip preservation)
//! - `from_xml()` and `to_xml()` methods using quick-xml
//! - Builder pattern for construction
//! - Per-class typed wrapper APIs around generated `TypedElement` registries
//!
//! Example generated output for `<x:cell>`:
//! ```ignore
//! #[derive(Debug, Clone)]
//! pub struct Cell {
//!     /// Cell reference (e.g., "A1")
//!     pub reference: Option<String>,
//!     /// Style index
//!     pub style_index: Option<u32>,
//!     /// Data type
//!     pub data_type: Option<CellDataType>,
//!     /// Cell value
//!     pub value: Option<CellValue>,
//!     /// Cell formula
//!     pub formula: Option<CellFormula>,
//!     /// Unknown children preserved for roundtrip
//!     pub unknown_children: Vec<RawXmlNode>,
//!     /// Unknown attributes preserved for roundtrip
//!     pub unknown_attrs: Vec<(String, String)>,
//! }
//! ```

pub mod generator;
pub mod schema;

use std::path::{Path, PathBuf};

use anyhow::Result;

pub use generator::{generate_schema_registries, GeneratorOutput};
pub use schema::{
    load_openxml_sdk_data, OpenXmlSdkData, PRESENTATIONML_MAIN_URI, SPREADSHEETML_MAIN_URI,
    WORDPROCESSINGML_MAIN_URI,
};

pub fn generate_from_data_root<P: AsRef<Path>, Q: AsRef<Path>>(
    data_root: P,
    output_dir: Q,
) -> Result<GeneratorOutput> {
    let data = load_openxml_sdk_data(data_root)?;
    generate_schema_registries(&data, output_dir)
}

pub fn generate_default_into<P: AsRef<Path>>(output_dir: P) -> Result<GeneratorOutput> {
    let data_root =
        PathBuf::from(env!("CARGO_MANIFEST_DIR")).join("../../references/Open-XML-SDK/data");
    generate_from_data_root(data_root, output_dir)
}
