//! # offidized-docx
//!
//! High-level Word API inspired by python-docx.
//!
//! ```ignore
//! use offidized_docx::Document;
//!
//! // Create from scratch
//! let mut doc = Document::new();
//! doc.add_heading("Quarterly Report", 1);
//! doc.add_paragraph("Revenue increased by 15% this quarter.");
//!
//! let table = doc.add_table(3, 4);
//! table.cell(0, 0).set_text("Product");
//! table.cell(0, 1).set_text("Q1");
//!
//! doc.save("report.docx")?;
//!
//! // Open and modify
//! let mut doc = Document::open("report.docx")?;
//! for para in doc.paragraphs() {
//!     if para.text().contains("15%") {
//!         // Modify while preserving all formatting
//!     }
//! }
//! doc.save("report_updated.docx")?;
//! ```

pub mod bookmark;
pub mod comment;
pub mod content_control;
pub mod document;
pub mod error;
pub mod footnote;
pub mod image;
pub mod numbering;
pub mod paragraph;
pub mod properties;
pub mod run;
pub mod section;
pub mod style;
pub mod table;
pub mod text;

pub use bookmark::Bookmark;
pub use comment::Comment;
pub use content_control::{ContentControl, ContentControlType, ListItem};
pub use document::{BodyItem, BodyItems, Document, DocumentProtection};
pub use error::{DocxError, Result};
pub use footnote::{Endnote, Footnote};
pub use image::{FloatingImage, Image, InlineImage, WrapType};
pub use numbering::{
    NumberingDefinition, NumberingInstance, NumberingLevel, NumberingLevelOverride,
};
pub use paragraph::{
    LineSpacingRule, Paragraph, ParagraphAlignment, ParagraphBorder, ParagraphBorders, TabStop,
    TabStopAlignment, TabStopLeader,
};
pub use properties::DocumentProperties;
pub use run::{FieldCode, Run, UnderlineType};
pub use section::{
    HeaderFooter, LineNumberRestart, PageMargins, PageOrientation, Section, SectionBreakType,
    SectionVerticalAlignment,
};
pub use style::{Style, StyleKind, StyleRegistry};
pub use table::{
    CellBorders, CellMargins, Table, TableAlignment, TableBorder, TableBorders, TableLayout,
    TableRowProperties, TableWidthType, VerticalAlignment, VerticalMerge,
};
