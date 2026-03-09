#![cfg(feature = "editing")]
#![allow(clippy::unwrap_used)]
#![allow(clippy::expect_used)]
#![allow(clippy::panic)]

//! Integration tests for the collaborative editing pipeline.
//!
//! Exercises the full cycle: parse .docx -> import to CRDT -> edit via yrs
//! -> view model -> export -> re-parse -> verify.

use std::collections::HashMap;
use std::sync::Arc;

use offidized_docview::editor::bridge::DocEdit;
use offidized_docview::editor::crdt_doc::CrdtDoc;
use offidized_docview::editor::export::export_to_docx;
use offidized_docview::editor::import::import_document;
use offidized_docview::editor::intent::EditIntent;
use offidized_docview::editor::tokens;
use offidized_docview::editor::view::crdt_to_view_model;
use offidized_docview::model::BodyItem;
use offidized_docx::Document;
use yrs::{Any, Array, GetString, Map, Out, Text, Transact};

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

/// Extract the text content from a paragraph body item, panicking if not a paragraph.
fn paragraph_text(item: &BodyItem) -> String {
    match item {
        BodyItem::Paragraph(p) => p.runs.iter().map(|r| r.text.as_str()).collect::<String>(),
        BodyItem::Table(_) => panic!("expected paragraph, got table"),
    }
}

/// Extract a ParagraphModel from a body item, panicking if not a paragraph.
fn as_paragraph(item: &BodyItem) -> &offidized_docview::model::ParagraphModel {
    match item {
        BodyItem::Paragraph(p) => p,
        BodyItem::Table(_) => panic!("expected paragraph, got table"),
    }
}

/// Get the TextRef for a paragraph at `body_index` in the CRDT.
fn get_text_ref(crdt: &CrdtDoc, body_index: u32) -> yrs::TextRef {
    let txn = crdt.doc().transact();
    let body = crdt.body();
    let entry = body.get(&txn, body_index).expect("body entry exists");
    let map_ref = entry.cast::<yrs::MapRef>().expect("is map");
    let text_out = map_ref.get(&txn, "text").expect("has text");
    text_out.cast::<yrs::TextRef>().expect("is TextRef")
}

/// Get the paragraph ID string for a paragraph at `body_index` in the CRDT.
fn get_para_id_str(crdt: &CrdtDoc, body_index: u32) -> String {
    let txn = crdt.doc().transact();
    let body = crdt.body();
    let entry = body.get(&txn, body_index).expect("body entry exists");
    let map_ref = entry.cast::<yrs::MapRef>().expect("is map");
    match map_ref.get(&txn, "id").expect("has id") {
        Out::Any(Any::String(s)) => s.to_string(),
        _ => panic!("id is not string"),
    }
}

/// Find and clone the ParaId that matches the given string representation.
fn find_para_id(crdt: &CrdtDoc, id_str: &str) -> offidized_docview::editor::para_id::ParaId {
    crdt.para_index_map()
        .keys()
        .find(|pid| pid.to_string() == id_str)
        .expect("find para id")
        .clone()
}

// ---------------------------------------------------------------------------
// Test 1: import_and_view_model_roundtrip
// ---------------------------------------------------------------------------

#[test]
fn import_and_view_model_roundtrip() {
    // Create a multi-paragraph Document.
    let mut doc = Document::new();
    doc.add_paragraph("First paragraph");
    doc.add_heading("A Heading", 2);
    doc.add_paragraph("Third paragraph");

    // Import into CRDT.
    let mut crdt = CrdtDoc::new();
    import_document(&doc, &mut crdt).expect("import");

    // Convert to view model.
    let vm = crdt_to_view_model(&crdt, &doc).expect("view model");

    // Verify paragraph count.
    assert_eq!(vm.body.len(), 3, "should have 3 body items");

    // Verify text content.
    assert_eq!(paragraph_text(&vm.body[0]), "First paragraph");
    assert_eq!(paragraph_text(&vm.body[1]), "A Heading");
    assert_eq!(paragraph_text(&vm.body[2]), "Third paragraph");

    // Verify structure: the second item should be a heading level 2.
    let heading = as_paragraph(&vm.body[1]);
    assert_eq!(heading.heading_level, Some(2));

    // The first and third should not have heading levels.
    assert!(as_paragraph(&vm.body[0]).heading_level.is_none());
    assert!(as_paragraph(&vm.body[2]).heading_level.is_none());

    // At least one section should be present (the document-level section).
    assert!(!vm.sections.is_empty());

    // The para_index_map should have 3 entries.
    assert_eq!(crdt.para_index_map().len(), 3);
}

// ---------------------------------------------------------------------------
// Test 2: edit_text_and_export_roundtrip
// ---------------------------------------------------------------------------

#[test]
fn edit_text_and_export_roundtrip() {
    // Create a simple Document with "Hello World".
    let mut doc = Document::new();
    doc.add_paragraph("Hello World");
    let bytes = doc.to_bytes().expect("to_bytes");
    let doc = Document::from_bytes(&bytes).expect("from_bytes");

    // Import into CRDT.
    let mut crdt = CrdtDoc::new();
    import_document(&doc, &mut crdt).expect("import");

    // Get the TextRef for the first paragraph.
    let text_ref = get_text_ref(&crdt, 0);

    // Verify initial content.
    {
        let txn = crdt.doc().transact();
        assert_eq!(text_ref.get_string(&txn), "Hello World");
    }

    // Edit the text: insert " beautiful" after "Hello" (offset 5).
    {
        let mut txn = crdt.doc().transact_mut();
        text_ref.insert(&mut txn, 5, " beautiful");
    }

    // Mark paragraph dirty.
    let para_id_str = get_para_id_str(&crdt, 0);
    let para_id = find_para_id(&crdt, &para_id_str);
    crdt.mark_dirty(&para_id);

    // Export to .docx bytes.
    let output_bytes = export_to_docx(&crdt, &doc).expect("export");

    // Re-parse the exported .docx.
    let reopened = Document::from_bytes(&output_bytes).expect("reopen");

    // Verify the text is "Hello beautiful World".
    assert_eq!(reopened.paragraphs().len(), 1);
    assert_eq!(reopened.paragraphs()[0].text(), "Hello beautiful World");
}

// ---------------------------------------------------------------------------
// Test 3: format_and_view_model
// ---------------------------------------------------------------------------

#[test]
fn format_and_view_model() {
    // Create Document with text.
    let mut doc = Document::new();
    doc.add_paragraph("Hello World");

    // Import into CRDT.
    let mut crdt = CrdtDoc::new();
    import_document(&doc, &mut crdt).expect("import");

    // Get the TextRef for the first paragraph.
    let text_ref = get_text_ref(&crdt, 0);

    // Apply bold formatting to "Hello" (range [0, 5)).
    {
        let mut attrs: HashMap<Arc<str>, Any> = HashMap::new();
        attrs.insert(Arc::from("bold"), Any::Bool(true));
        let mut txn = crdt.doc().transact_mut();
        text_ref.format(&mut txn, 0, 5, attrs);
    }

    // Get view model and verify the runs.
    let vm = crdt_to_view_model(&crdt, &doc).expect("view model");
    let para = as_paragraph(&vm.body[0]);

    // Should have at least 2 runs: bold "Hello" and non-bold " World".
    assert!(
        para.runs.len() >= 2,
        "expected at least 2 runs, got {}",
        para.runs.len()
    );

    // Find the bold run.
    let bold_run = para.runs.iter().find(|r| r.bold);
    assert!(bold_run.is_some(), "should have a bold run");
    assert_eq!(bold_run.unwrap().text, "Hello");

    // Find the non-bold run.
    let plain_run = para.runs.iter().find(|r| !r.bold && !r.text.is_empty());
    assert!(plain_run.is_some(), "should have a non-bold run");
    assert_eq!(plain_run.unwrap().text, " World");
}

// ---------------------------------------------------------------------------
// Test 4: multi_paragraph_selective_edit
// ---------------------------------------------------------------------------

#[test]
fn multi_paragraph_selective_edit() {
    // Create Document with 3 paragraphs.
    let mut doc = Document::new();
    doc.add_paragraph("First paragraph");
    doc.add_paragraph("Middle paragraph");
    doc.add_paragraph("Third paragraph");
    let bytes = doc.to_bytes().expect("to_bytes");
    let doc = Document::from_bytes(&bytes).expect("from_bytes");

    // Import into CRDT.
    let mut crdt = CrdtDoc::new();
    import_document(&doc, &mut crdt).expect("import");

    // Edit only the middle paragraph: replace "Middle" with "Edited".
    let text_ref = get_text_ref(&crdt, 1);

    {
        let txn = crdt.doc().transact();
        let content = text_ref.get_string(&txn);
        assert_eq!(content, "Middle paragraph");
    }

    {
        let mut txn = crdt.doc().transact_mut();
        // Remove "Middle" (6 chars).
        text_ref.remove_range(&mut txn, 0, 6);
    }
    {
        let mut txn = crdt.doc().transact_mut();
        text_ref.insert(&mut txn, 0, "Edited");
    }

    // Mark only the middle paragraph dirty.
    let para_id_str = get_para_id_str(&crdt, 1);
    let para_id = find_para_id(&crdt, &para_id_str);
    crdt.mark_dirty(&para_id);

    // Export and verify.
    let output_bytes = export_to_docx(&crdt, &doc).expect("export");
    let reopened = Document::from_bytes(&output_bytes).expect("reopen");

    assert_eq!(reopened.paragraphs().len(), 3);
    // First paragraph unchanged.
    assert_eq!(reopened.paragraphs()[0].text(), "First paragraph");
    // Middle paragraph has the edit.
    assert_eq!(reopened.paragraphs()[1].text(), "Edited paragraph");
    // Third paragraph unchanged.
    assert_eq!(reopened.paragraphs()[2].text(), "Third paragraph");
}

// ---------------------------------------------------------------------------
// Test 5: export_preserves_non_text_content
// ---------------------------------------------------------------------------

#[test]
fn export_preserves_non_text_content() {
    // Create Document with a paragraph and a table.
    let mut doc = Document::new();
    doc.add_paragraph("Before table");
    doc.add_table(2, 3);
    let bytes = doc.to_bytes().expect("to_bytes");
    let doc = Document::from_bytes(&bytes).expect("from_bytes");

    // Import into CRDT (tables should pass through).
    let mut crdt = CrdtDoc::new();
    import_document(&doc, &mut crdt).expect("import");

    // Verify the CRDT has 2 body items (paragraph + table).
    {
        let txn = crdt.doc().transact();
        assert_eq!(crdt.body().len(&txn), 2);
    }

    // Verify the view model shows the table.
    let vm = crdt_to_view_model(&crdt, &doc).expect("view model");
    assert_eq!(vm.body.len(), 2);
    match &vm.body[0] {
        BodyItem::Paragraph(p) => {
            assert_eq!(
                p.runs.iter().map(|r| r.text.as_str()).collect::<String>(),
                "Before table"
            );
        }
        BodyItem::Table(_) => panic!("expected paragraph first"),
    }
    match &vm.body[1] {
        BodyItem::Table(t) => {
            assert_eq!(t.rows.len(), 2, "table should have 2 rows");
            assert_eq!(t.rows[0].cells.len(), 3, "first row should have 3 cells");
        }
        BodyItem::Paragraph(_) => panic!("expected table second"),
    }

    // Export and verify the table is preserved. Since no paragraphs are dirty,
    // the export should just return the original bytes (which contain the table).
    let output_bytes = export_to_docx(&crdt, &doc).expect("export");
    let reopened = Document::from_bytes(&output_bytes).expect("reopen");

    // The original doc has a paragraph and a table. Verify the paragraph is intact.
    assert_eq!(reopened.paragraphs()[0].text(), "Before table");

    // Verify the table exists by checking body items.
    let mut found_table = false;
    for item in reopened.body_items() {
        if let offidized_docx::BodyItem::Table(t) = item {
            assert_eq!(t.rows(), 2);
            assert_eq!(t.columns(), 3);
            found_table = true;
        }
    }
    assert!(found_table, "table should be preserved in export");
}

#[test]
fn insert_table_via_intent_roundtrips() {
    let editor = DocEdit::blank().expect("blank editor");
    let intent = EditIntent::InsertTable {
        anchor: DocEdit::encode_position_js(0, 0),
        rows: 2,
        columns: 2,
    };
    let json = serde_json::to_string(&intent).expect("serialize insertTable");
    editor.apply_intent(&json).expect("apply insertTable");

    let output_bytes = editor.save().expect("save");
    let reopened = Document::from_bytes(&output_bytes).expect("reopen");

    let mut body_items = reopened.body_items();
    let Some(offidized_docx::BodyItem::Paragraph(first_paragraph)) = body_items.next() else {
        panic!("expected first body item to be a paragraph");
    };
    assert_eq!(first_paragraph.text(), "");

    let Some(offidized_docx::BodyItem::Table(inserted_table)) = body_items.next() else {
        panic!("expected second body item to be a table");
    };
    assert_eq!(inserted_table.rows(), 2);
    assert_eq!(inserted_table.columns(), 2);
    assert_eq!(inserted_table.cell_text(0, 0), Some(""));
    assert_eq!(inserted_table.cell_text(1, 1), Some(""));
}

#[test]
fn edited_table_cell_via_intent_roundtrips() {
    let mut doc = Document::new();
    doc.add_paragraph("Before table");
    let table = doc.add_table(1, 2);
    assert!(table.set_cell_text(0, 0, "A1"));
    assert!(table.set_cell_text(0, 1, "B1"));
    let bytes = doc.to_bytes().expect("to_bytes");

    let editor = DocEdit::new(&bytes).expect("editor");
    let intent = EditIntent::SetTableCellText {
        body_index: 1,
        row: 0,
        col: 1,
        text: "Edited".to_string(),
    };
    let json = serde_json::to_string(&intent).expect("serialize setTableCellText");
    editor.apply_intent(&json).expect("apply setTableCellText");

    let output_bytes = editor.save().expect("save");
    let reopened = Document::from_bytes(&output_bytes).expect("reopen");

    let mut body_items = reopened.body_items();
    let Some(offidized_docx::BodyItem::Paragraph(first_paragraph)) = body_items.next() else {
        panic!("expected first body item to be a paragraph");
    };
    assert_eq!(first_paragraph.text(), "Before table");

    let Some(offidized_docx::BodyItem::Table(reopened_table)) = body_items.next() else {
        panic!("expected second body item to be a table");
    };
    assert_eq!(reopened_table.rows(), 1);
    assert_eq!(reopened_table.columns(), 2);
    assert_eq!(reopened_table.cell_text(0, 0), Some("A1"));
    assert_eq!(reopened_table.cell_text(0, 1), Some("Edited"));
}

// ---------------------------------------------------------------------------
// Test 6: sentinel_survives_roundtrip
// ---------------------------------------------------------------------------

#[test]
fn sentinel_survives_roundtrip() {
    // Create Document with tab runs.
    let mut doc = Document::new();
    let para = doc.add_paragraph("Before");
    if let Some(run) = para.runs_mut().first_mut() {
        run.set_has_tab(true);
    }
    let bytes = doc.to_bytes().expect("to_bytes");
    let doc = Document::from_bytes(&bytes).expect("from_bytes");

    // Import into CRDT (tabs become sentinels).
    let mut crdt = CrdtDoc::new();
    import_document(&doc, &mut crdt).expect("import");

    // Verify sentinel is present in the CRDT text.
    {
        let txn = crdt.doc().transact();
        let text_ref = get_text_ref(&crdt, 0);
        let content = text_ref.get_string(&txn);
        assert!(
            content.contains(tokens::SENTINEL),
            "CRDT text should contain sentinel"
        );
    }

    // View model should show has_tab.
    let vm = crdt_to_view_model(&crdt, &doc).expect("view model");
    let para_model = as_paragraph(&vm.body[0]);
    let tab_run = para_model.runs.iter().find(|r| r.has_tab);
    assert!(tab_run.is_some(), "view model should have a tab run");

    // Mark dirty and export. The export should reconstruct the tab run.
    let para_id_str = get_para_id_str(&crdt, 0);
    let para_id = find_para_id(&crdt, &para_id_str);
    crdt.mark_dirty(&para_id);

    let output_bytes = export_to_docx(&crdt, &doc).expect("export");
    let reopened = Document::from_bytes(&output_bytes).expect("reopen");

    // Verify tabs are preserved.
    assert_eq!(reopened.paragraphs().len(), 1);
    let runs = reopened.paragraphs()[0].runs();
    let has_tab = runs.iter().any(|r| r.has_tab());
    assert!(has_tab, "tab should be preserved after export roundtrip");
}

// ---------------------------------------------------------------------------
// Test 7: empty_document_roundtrip
// ---------------------------------------------------------------------------

#[test]
fn empty_document_roundtrip() {
    // Create empty Document.
    let doc = Document::new();

    // Import into CRDT.
    let mut crdt = CrdtDoc::new();
    import_document(&doc, &mut crdt).expect("import empty doc");

    // CRDT body should be empty.
    {
        let txn = crdt.doc().transact();
        assert_eq!(crdt.body().len(&txn), 0);
    }

    // View model should work without panicking.
    let vm = crdt_to_view_model(&crdt, &doc).expect("view model");
    assert!(vm.body.is_empty());
    assert!(
        !vm.sections.is_empty(),
        "should have document-level section"
    );

    // Export should succeed.
    let output_bytes = export_to_docx(&crdt, &doc).expect("export");
    let reopened = Document::from_bytes(&output_bytes).expect("reopen");

    // Re-importing the exported empty doc should also work.
    let mut crdt2 = CrdtDoc::new();
    import_document(&reopened, &mut crdt2).expect("re-import empty doc");
    let vm2 = crdt_to_view_model(&crdt2, &reopened).expect("view model 2");
    assert!(vm2.body.is_empty());
}

// ---------------------------------------------------------------------------
// Test 8: dirty_tracking_only_exports_changed
// ---------------------------------------------------------------------------

#[test]
fn dirty_tracking_only_exports_changed() {
    // Create multi-paragraph doc.
    let mut doc = Document::new();
    doc.add_paragraph("Alpha");
    doc.add_paragraph("Beta");
    doc.add_paragraph("Gamma");

    // Import into CRDT.
    let mut crdt = CrdtDoc::new();
    import_document(&doc, &mut crdt).expect("import");

    // Nothing should be dirty initially.
    assert!(crdt.dirty_paragraphs().is_empty());

    // Edit one paragraph and mark dirty.
    let text_ref_1 = get_text_ref(&crdt, 0);
    {
        let mut txn = crdt.doc().transact_mut();
        text_ref_1.insert(&mut txn, 5, "!");
    }
    let para_id_str_0 = get_para_id_str(&crdt, 0);
    let para_id_0 = find_para_id(&crdt, &para_id_str_0);
    crdt.mark_dirty(&para_id_0);

    assert_eq!(crdt.dirty_paragraphs().len(), 1);
    assert!(crdt.dirty_paragraphs().contains(&para_id_0));

    // Clear dirty.
    crdt.clear_dirty();
    assert!(crdt.dirty_paragraphs().is_empty());

    // Edit another paragraph, mark dirty.
    let text_ref_2 = get_text_ref(&crdt, 2);
    {
        let mut txn = crdt.doc().transact_mut();
        text_ref_2.insert(&mut txn, 5, "?");
    }
    let para_id_str_2 = get_para_id_str(&crdt, 2);
    let para_id_2 = find_para_id(&crdt, &para_id_str_2);
    crdt.mark_dirty(&para_id_2);

    // Verify only the third paragraph is dirty.
    assert_eq!(crdt.dirty_paragraphs().len(), 1);
    assert!(crdt.dirty_paragraphs().contains(&para_id_2));
    assert!(!crdt.dirty_paragraphs().contains(&para_id_0));
}
