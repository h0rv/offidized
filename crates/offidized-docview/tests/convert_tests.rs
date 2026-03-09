#![allow(clippy::unwrap_used)]
#![allow(clippy::expect_used)]

use offidized_docview::convert::convert_document;
use offidized_docview::model::BodyItem;
use offidized_docview::units::{
    emu_to_pt, half_points_to_pt, signed_emu_to_pt, signed_twips_to_pt, twips_to_pt,
};
use offidized_docx::Document;

// ---------------------------------------------------------------------------
// Unit conversion tests
// ---------------------------------------------------------------------------

#[test]
fn twips_to_pt_basic() {
    assert!((twips_to_pt(240) - 12.0).abs() < f64::EPSILON);
    assert!((twips_to_pt(0) - 0.0).abs() < f64::EPSILON);
    assert!((twips_to_pt(1440) - 72.0).abs() < f64::EPSILON); // 1 inch
}

#[test]
fn half_points_to_pt_basic() {
    assert!((half_points_to_pt(24) - 12.0).abs() < f64::EPSILON);
    assert!((half_points_to_pt(22) - 11.0).abs() < f64::EPSILON);
}

#[test]
fn emu_to_pt_basic() {
    // 1 inch = 914400 EMU = 72 pt
    assert!((emu_to_pt(914_400) - 72.0).abs() < 0.01);
    assert!((emu_to_pt(12_700) - 1.0).abs() < f64::EPSILON);
}

#[test]
fn signed_conversions() {
    assert!((signed_twips_to_pt(-240) - -12.0).abs() < f64::EPSILON);
    assert!((signed_emu_to_pt(-12_700) - -1.0).abs() < f64::EPSILON);
}

// ---------------------------------------------------------------------------
// Document conversion tests
// ---------------------------------------------------------------------------

#[test]
fn convert_simple_paragraphs() {
    let mut doc = Document::new();
    doc.add_paragraph("Hello World");
    doc.add_heading("Title", 1);

    let model = convert_document(&doc).expect("conversion should succeed");

    assert_eq!(model.body.len(), 2, "should have 2 body items");

    // First item: paragraph with "Hello World"
    match &model.body[0] {
        BodyItem::Paragraph(p) => {
            assert_eq!(p.runs.len(), 1);
            assert_eq!(p.runs[0].text, "Hello World");
            assert!(p.heading_level.is_none());
        }
        BodyItem::Table(_) => panic!("expected paragraph, got table"),
    }

    // Second item: heading
    match &model.body[1] {
        BodyItem::Paragraph(p) => {
            assert_eq!(p.heading_level, Some(1));
            assert_eq!(p.runs[0].text, "Title");
        }
        BodyItem::Table(_) => panic!("expected paragraph, got table"),
    }

    // At least one section
    assert!(
        !model.sections.is_empty(),
        "should have at least one section"
    );
}

#[test]
fn convert_table() {
    let mut doc = Document::new();
    doc.add_table(2, 3);

    let model = convert_document(&doc).expect("conversion should succeed");

    // Find the table in body
    let table = model.body.iter().find_map(|item| match item {
        BodyItem::Table(t) => Some(t),
        _ => None,
    });
    assert!(table.is_some(), "should have a table in body");
    let table = table.unwrap();

    assert_eq!(table.rows.len(), 2, "table should have 2 rows");
    assert_eq!(
        table.rows[0].cells.len(),
        3,
        "first row should have 3 cells"
    );
    assert_eq!(
        table.rows[1].cells.len(),
        3,
        "second row should have 3 cells"
    );
}

#[test]
fn convert_bulleted_paragraph() {
    let mut doc = Document::new();
    doc.add_bulleted_paragraph("Item 1");

    let model = convert_document(&doc).expect("conversion should succeed");

    // Find the bulleted paragraph
    let para = model.body.iter().find_map(|item| match item {
        BodyItem::Paragraph(p) => Some(p),
        _ => None,
    });
    assert!(para.is_some(), "should have a paragraph");
    let para = para.unwrap();

    assert!(
        para.numbering.is_some(),
        "bulleted paragraph should have numbering"
    );
    let numbering = para.numbering.as_ref().unwrap();
    assert_eq!(numbering.format, "bullet");
}

#[test]
fn model_serializes_to_json() {
    let mut doc = Document::new();
    doc.add_paragraph("test");

    let model = convert_document(&doc).expect("conversion should succeed");
    let json = serde_json::to_string(&model).expect("JSON serialization should succeed");

    assert!(json.contains("\"type\":\"paragraph\""));
    assert!(json.contains("\"text\":\"test\""));
}

#[test]
fn sections_have_default_dimensions() {
    let doc = Document::new();
    let model = convert_document(&doc).expect("conversion should succeed");

    assert!(!model.sections.is_empty());
    let section = &model.sections[0];

    // Should have reasonable defaults (US Letter)
    assert!(section.page_width_pt > 0.0);
    assert!(section.page_height_pt > 0.0);
    assert!(section.margins.top >= 0.0);
    assert!(section.margins.left >= 0.0);
}

// ---------------------------------------------------------------------------
// Real file conversion tests
// ---------------------------------------------------------------------------

fn fixture_dir() -> std::path::PathBuf {
    std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
        .parent()
        .expect("crates dir")
        .parent()
        .expect("workspace root")
        .join("Open-XML-SDK/test/DocumentFormat.OpenXml.Tests.Assets/assets/TestFiles")
}

#[test]
fn convert_real_hello_world_docx() {
    let path = fixture_dir().join("HelloWorld.docx");
    if !path.exists() {
        // Skip if reference files not available.
        eprintln!("Skipping: {path:?} not found");
        return;
    }

    let doc = Document::open(&path).expect("should open HelloWorld.docx");
    let model = convert_document(&doc).expect("conversion should succeed");

    assert!(!model.body.is_empty(), "body should not be empty");
    assert!(!model.sections.is_empty(), "sections should not be empty");
}

#[test]
fn convert_real_complex_docx() {
    let path = fixture_dir().join("complex2010.docx");
    if !path.exists() {
        eprintln!("Skipping: {path:?} not found");
        return;
    }

    let doc = Document::open(&path).expect("should open complex2010.docx");
    let model = convert_document(&doc).expect("conversion should succeed");

    assert!(!model.body.is_empty(), "body should not be empty");
    assert!(!model.sections.is_empty(), "sections should not be empty");
}

#[test]
fn from_bytes_roundtrip() {
    let mut doc = Document::new();
    doc.add_paragraph("bytes test");

    // Save to bytes via tempfile, then reload via from_bytes.
    let dir = tempfile::tempdir().expect("should create temp dir");
    let path = dir.path().join("test.docx");
    doc.save(&path).expect("should save");

    let bytes = std::fs::read(&path).expect("should read bytes");
    let doc2 = Document::from_bytes(&bytes).expect("should parse from bytes");
    let model = convert_document(&doc2).expect("conversion should succeed");

    assert!(!model.body.is_empty());
    match &model.body[0] {
        BodyItem::Paragraph(p) => {
            assert_eq!(p.runs[0].text, "bytes test");
        }
        BodyItem::Table(_) => panic!("expected paragraph"),
    }
}
