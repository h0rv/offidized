//! # offidized
//!
//! **Office, oxidized.** A Rust-native OOXML library for reading, writing,
//! and manipulating Excel (.xlsx), Word (.docx), and PowerPoint (.pptx)
//! files with full roundtrip fidelity.
//!
//! This is the umbrella crate that re-exports the format-specific crates.
//! You can also depend on individual crates directly if you only need one format.
//!
//! ## Quick Start
//!
//! ```ignore
//! use offidized::{Workbook, Document, Presentation};
//!
//! // Excel
//! let mut wb = Workbook::new();
//! wb.add_sheet("Data").cell("A1").set_value("Hello");
//! wb.save("hello.xlsx")?;
//!
//! // Word
//! let mut doc = Document::new();
//! doc.add_heading("Title", 1);
//! doc.add_paragraph("Body text");
//! doc.save("hello.docx")?;
//!
//! // PowerPoint
//! let mut prs = Presentation::new();
//! prs.add_slide_with_title("Hello World");
//! prs.save("hello.pptx")?;
//! ```

pub use offidized_opc as opc;

#[cfg(feature = "xlsx")]
pub use offidized_xlsx as xlsx;
#[cfg(feature = "xlsx")]
pub use offidized_xlsx::Workbook;

#[cfg(feature = "docx")]
pub use offidized_docx as docx;
#[cfg(feature = "docx")]
pub use offidized_docx::Document;

#[cfg(feature = "pptx")]
pub use offidized_pptx as pptx;
#[cfg(feature = "pptx")]
pub use offidized_pptx::Presentation;
