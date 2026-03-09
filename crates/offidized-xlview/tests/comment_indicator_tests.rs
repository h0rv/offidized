//! Integration tests for comment indicator flags.
//!
//! These tests verify that cells with comments have the `has_comment` flag properly set.
//! The comment indicator is rendered as a small red triangle in the top-right corner
//! of the cell (rendering is tested separately in E2E tests).
//!
//! Comment indicator behavior:
//! - Cells with comments should have `has_comment: Some(true)`
//! - Cells without comments should have `has_comment: None` or `Some(false)`
#![allow(
    clippy::unwrap_used,
    clippy::expect_used,
    clippy::indexing_slicing,
    clippy::float_cmp,
    clippy::panic,
    clippy::approx_constant,
    clippy::cast_possible_truncation,
    clippy::absurd_extreme_comparisons,
    clippy::cast_lossless
)]

mod common;
mod fixtures;

use common::load_xlsx;
use std::io::{Cursor, Write};
use zip::write::FileOptions;
use zip::ZipWriter;

/// Create an XLSX file with a comment on a specific cell.
fn create_xlsx_with_comment(cell_ref: &str, comment_text: &str, author: &str) -> Vec<u8> {
    let cursor = Cursor::new(Vec::new());
    let mut zip = ZipWriter::new(cursor);
    let options = FileOptions::<()>::default().compression_method(zip::CompressionMethod::Deflated);

    let _ = zip.start_file("[Content_Types].xml", options);
    let _ = zip.write_all(
        br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
<Override PartName="/xl/comments1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml"/>
</Types>"#,
    );

    let _ = zip.start_file("_rels/.rels", options);
    let _ = zip.write_all(
        br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>"#,
    );

    let _ = zip.start_file("xl/_rels/workbook.xml.rels", options);
    let _ = zip.write_all(
        br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>"#,
    );

    let _ = zip.start_file("xl/workbook.xml", options);
    let _ = zip.write_all(
        br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<sheets>
<sheet name="Sheet1" sheetId="1" r:id="rId1"/>
</sheets>
</workbook>"#,
    );

    let _ = zip.start_file("xl/styles.xml", options);
    let _ = zip.write_all(
        br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>
<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>
<borders count="1"><border><left/><right/><top/><bottom/></border></borders>
<cellXfs count="1"><xf fontId="0" fillId="0" borderId="0"/></cellXfs>
</styleSheet>"#,
    );

    let _ = zip.start_file("xl/worksheets/_rels/sheet1.xml.rels", options);
    let _ = zip.write_all(
        br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="../comments1.xml"/>
</Relationships>"#,
    );

    let _ = zip.start_file("xl/worksheets/sheet1.xml", options);
    let sheet_xml = format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<sheetData>
<row r="1">
<c r="{cell_ref}" t="inlineStr"><is><t>Cell with comment</t></is></c>
</row>
</sheetData>
</worksheet>"#
    );
    let _ = zip.write_all(sheet_xml.as_bytes());

    let _ = zip.start_file("xl/comments1.xml", options);
    let comments_xml = format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<comments xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<authors>
<author>{author}</author>
</authors>
<commentList>
<comment ref="{cell_ref}" authorId="0">
<text>
<r><t>{comment_text}</t></r>
</text>
</comment>
</commentList>
</comments>"#
    );
    let _ = zip.write_all(comments_xml.as_bytes());

    let cursor = zip.finish().expect("Failed to finish ZIP");
    cursor.into_inner()
}

/// Create an XLSX file without any comments.
fn create_xlsx_without_comments() -> Vec<u8> {
    let cursor = Cursor::new(Vec::new());
    let mut zip = ZipWriter::new(cursor);
    let options = FileOptions::<()>::default().compression_method(zip::CompressionMethod::Deflated);

    let _ = zip.start_file("[Content_Types].xml", options);
    let _ = zip.write_all(
        br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
</Types>"#,
    );

    let _ = zip.start_file("_rels/.rels", options);
    let _ = zip.write_all(
        br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>"#,
    );

    let _ = zip.start_file("xl/_rels/workbook.xml.rels", options);
    let _ = zip.write_all(
        br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>"#,
    );

    let _ = zip.start_file("xl/workbook.xml", options);
    let _ = zip.write_all(
        br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<sheets>
<sheet name="Sheet1" sheetId="1" r:id="rId1"/>
</sheets>
</workbook>"#,
    );

    let _ = zip.start_file("xl/styles.xml", options);
    let _ = zip.write_all(
        br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>
<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>
<borders count="1"><border><left/><right/><top/><bottom/></border></borders>
<cellXfs count="1"><xf fontId="0" fillId="0" borderId="0"/></cellXfs>
</styleSheet>"#,
    );

    let _ = zip.start_file("xl/worksheets/sheet1.xml", options);
    let _ = zip.write_all(
        br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<sheetData>
<row r="1"><c r="A1" t="inlineStr"><is><t>Cell without comment</t></is></c></row>
</sheetData>
</worksheet>"#,
    );

    let cursor = zip.finish().expect("Failed to finish ZIP");
    cursor.into_inner()
}

/// Create an XLSX file with multiple cells, some with comments and some without.
fn create_xlsx_with_mixed_comments() -> Vec<u8> {
    let cursor = Cursor::new(Vec::new());
    let mut zip = ZipWriter::new(cursor);
    let options = FileOptions::<()>::default().compression_method(zip::CompressionMethod::Deflated);

    let _ = zip.start_file("[Content_Types].xml", options);
    let _ = zip.write_all(
        br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
<Override PartName="/xl/comments1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml"/>
</Types>"#,
    );

    let _ = zip.start_file("_rels/.rels", options);
    let _ = zip.write_all(
        br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>"#,
    );

    let _ = zip.start_file("xl/_rels/workbook.xml.rels", options);
    let _ = zip.write_all(
        br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>"#,
    );

    let _ = zip.start_file("xl/workbook.xml", options);
    let _ = zip.write_all(
        br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<sheets>
<sheet name="Sheet1" sheetId="1" r:id="rId1"/>
</sheets>
</workbook>"#,
    );

    let _ = zip.start_file("xl/styles.xml", options);
    let _ = zip.write_all(
        br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>
<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>
<borders count="1"><border><left/><right/><top/><bottom/></border></borders>
<cellXfs count="1"><xf fontId="0" fillId="0" borderId="0"/></cellXfs>
</styleSheet>"#,
    );

    let _ = zip.start_file("xl/worksheets/_rels/sheet1.xml.rels", options);
    let _ = zip.write_all(
        br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="../comments1.xml"/>
</Relationships>"#,
    );

    // xl/worksheets/sheet1.xml - 4 cells: A1, B1, A2, B2
    let _ = zip.start_file("xl/worksheets/sheet1.xml", options);
    let _ = zip.write_all(
        br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<sheetData>
<row r="1">
<c r="A1" t="inlineStr"><is><t>Has comment</t></is></c>
<c r="B1" t="inlineStr"><is><t>No comment</t></is></c>
</row>
<row r="2">
<c r="A2" t="inlineStr"><is><t>No comment</t></is></c>
<c r="B2" t="inlineStr"><is><t>Has comment</t></is></c>
</row>
</sheetData>
</worksheet>"#,
    );

    // xl/comments1.xml - Comments only on A1 and B2
    let _ = zip.start_file("xl/comments1.xml", options);
    let _ = zip.write_all(
        br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<comments xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<authors>
<author>Test Author</author>
</authors>
<commentList>
<comment ref="A1" authorId="0">
<text><r><t>Comment on A1</t></r></text>
</comment>
<comment ref="B2" authorId="0">
<text><r><t>Comment on B2</t></r></text>
</comment>
</commentList>
</comments>"#,
    );

    let cursor = zip.finish().expect("Failed to finish ZIP");
    cursor.into_inner()
}

/// Create an XLSX file with a comment containing rich text (multiple formatted runs).
fn create_xlsx_with_rich_text_comment() -> Vec<u8> {
    let cursor = Cursor::new(Vec::new());
    let mut zip = ZipWriter::new(cursor);
    let options = FileOptions::<()>::default().compression_method(zip::CompressionMethod::Deflated);

    let _ = zip.start_file("[Content_Types].xml", options);
    let _ = zip.write_all(
        br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
<Override PartName="/xl/comments1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml"/>
</Types>"#,
    );

    let _ = zip.start_file("_rels/.rels", options);
    let _ = zip.write_all(
        br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>"#,
    );

    let _ = zip.start_file("xl/_rels/workbook.xml.rels", options);
    let _ = zip.write_all(
        br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>"#,
    );

    let _ = zip.start_file("xl/workbook.xml", options);
    let _ = zip.write_all(
        br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<sheets>
<sheet name="Sheet1" sheetId="1" r:id="rId1"/>
</sheets>
</workbook>"#,
    );

    let _ = zip.start_file("xl/styles.xml", options);
    let _ = zip.write_all(
        br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>
<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>
<borders count="1"><border><left/><right/><top/><bottom/></border></borders>
<cellXfs count="1"><xf fontId="0" fillId="0" borderId="0"/></cellXfs>
</styleSheet>"#,
    );

    let _ = zip.start_file("xl/worksheets/_rels/sheet1.xml.rels", options);
    let _ = zip.write_all(
        br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="../comments1.xml"/>
</Relationships>"#,
    );

    let _ = zip.start_file("xl/worksheets/sheet1.xml", options);
    let _ = zip.write_all(
        br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<sheetData>
<row r="1"><c r="A1" t="inlineStr"><is><t>Cell with rich text comment</t></is></c></row>
</sheetData>
</worksheet>"#,
    );

    // xl/comments1.xml - Comment with multiple rich text runs (bold, italic, etc.)
    let _ = zip.start_file("xl/comments1.xml", options);
    let _ = zip.write_all(
        br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<comments xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<authors>
<author>Rich Text Author</author>
</authors>
<commentList>
<comment ref="A1" authorId="0">
<text>
<r><rPr><b/></rPr><t>Bold text </t></r>
<r><rPr><i/></rPr><t>italic text </t></r>
<r><t>normal text</t></r>
</text>
</comment>
</commentList>
</comments>"#,
    );

    let cursor = zip.finish().expect("Failed to finish ZIP");
    cursor.into_inner()
}

// =============================================================================
// Tests
// =============================================================================

/// Test 1: Cell with a comment has has_comment = true
#[test]
fn test_cell_with_comment_has_indicator_flag() {
    let xlsx = create_xlsx_with_comment("A1", "This is a comment", "Author");
    let workbook = load_xlsx(&xlsx);

    assert_eq!(workbook.sheets.len(), 1);
    let sheet = &workbook.sheets[0];

    // Find cell A1 (row 0, col 0)
    let cell = sheet
        .cells
        .iter()
        .find(|c| c.r == 0 && c.c == 0)
        .expect("Cell A1 should exist");

    assert_eq!(
        cell.cell.has_comment,
        Some(true),
        "Cell with comment should have has_comment = Some(true)"
    );
}

/// Test 2: Cell without comment has has_comment = false/None
#[test]
fn test_cell_without_comment_has_no_indicator_flag() {
    let xlsx = create_xlsx_without_comments();
    let workbook = load_xlsx(&xlsx);

    assert_eq!(workbook.sheets.len(), 1);
    let sheet = &workbook.sheets[0];

    // Find cell A1 (row 0, col 0)
    let cell = sheet
        .cells
        .iter()
        .find(|c| c.r == 0 && c.c == 0)
        .expect("Cell A1 should exist");

    // Cell without comment should have has_comment as None or Some(false)
    assert!(
        cell.cell.has_comment.is_none() || cell.cell.has_comment == Some(false),
        "Cell without comment should have has_comment = None or Some(false), got {:?}",
        cell.cell.has_comment
    );
}

/// Test 3: Multiple cells, only some with comments
#[test]
fn test_mixed_cells_only_commented_have_indicator_flag() {
    let xlsx = create_xlsx_with_mixed_comments();
    let workbook = load_xlsx(&xlsx);

    assert_eq!(workbook.sheets.len(), 1);
    let sheet = &workbook.sheets[0];

    // A1 (row 0, col 0) - has comment
    let cell_a1 = sheet
        .cells
        .iter()
        .find(|c| c.r == 0 && c.c == 0)
        .expect("Cell A1 should exist");
    assert_eq!(
        cell_a1.cell.has_comment,
        Some(true),
        "Cell A1 should have has_comment = Some(true)"
    );

    // B1 (row 0, col 1) - no comment
    let cell_b1 = sheet
        .cells
        .iter()
        .find(|c| c.r == 0 && c.c == 1)
        .expect("Cell B1 should exist");
    assert!(
        cell_b1.cell.has_comment.is_none() || cell_b1.cell.has_comment == Some(false),
        "Cell B1 should have has_comment = None or Some(false), got {:?}",
        cell_b1.cell.has_comment
    );

    // A2 (row 1, col 0) - no comment
    let cell_a2 = sheet
        .cells
        .iter()
        .find(|c| c.r == 1 && c.c == 0)
        .expect("Cell A2 should exist");
    assert!(
        cell_a2.cell.has_comment.is_none() || cell_a2.cell.has_comment == Some(false),
        "Cell A2 should have has_comment = None or Some(false), got {:?}",
        cell_a2.cell.has_comment
    );

    // B2 (row 1, col 1) - has comment
    let cell_b2 = sheet
        .cells
        .iter()
        .find(|c| c.r == 1 && c.c == 1)
        .expect("Cell B2 should exist");
    assert_eq!(
        cell_b2.cell.has_comment,
        Some(true),
        "Cell B2 should have has_comment = Some(true)"
    );
}

/// Test 4: Comment with rich text still sets has_comment
#[test]
fn test_rich_text_comment_sets_indicator_flag() {
    let xlsx = create_xlsx_with_rich_text_comment();
    let workbook = load_xlsx(&xlsx);

    assert_eq!(workbook.sheets.len(), 1);
    let sheet = &workbook.sheets[0];

    // Verify the comment was parsed correctly (rich text runs concatenated)
    assert_eq!(sheet.comments.len(), 1);
    let comment = &sheet.comments[0];
    // Rich text runs should be concatenated
    assert!(
        comment.text.contains("Bold text"),
        "Comment should contain 'Bold text'"
    );
    assert!(
        comment.text.contains("italic text"),
        "Comment should contain 'italic text'"
    );
    assert!(
        comment.text.contains("normal text"),
        "Comment should contain 'normal text'"
    );

    // Find cell A1 (row 0, col 0)
    let cell = sheet
        .cells
        .iter()
        .find(|c| c.r == 0 && c.c == 0)
        .expect("Cell A1 should exist");

    // Even with rich text formatting, the has_comment flag should be set
    assert_eq!(
        cell.cell.has_comment,
        Some(true),
        "Cell with rich text comment should have has_comment = Some(true)"
    );
}
