//! Tests for comments/notes parsing in XLSX files
//!
//! Excel comments (called "notes" in newer versions) are stored in separate XML files
//! within the XLSX package. Each sheet can have its own comments file (e.g., xl/comments1.xml)
//! linked via the sheet's relationship file (xl/worksheets/_rels/sheet1.xml.rels).
//!
//! Comment structure in xl/commentsN.xml:
//! ```xml
//! <comments>
//!   <authors>
//!     <author>John Doe</author>
//!   </authors>
//!   <commentList>
//!     <comment ref="A1" authorId="0">
//!       <text>
//!         <r><t>This is a comment</t></r>
//!       </text>
//!     </comment>
//!   </commentList>
//! </comments>
//! ```
//!
//! The relationship in xl/worksheets/_rels/sheet1.xml.rels:
//! ```xml
//! <Relationship Id="rId1"
//!   Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments"
//!   Target="../comments1.xml"/>
//! ```
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

/// Create an XLSX file with a comment on a cell
fn create_xlsx_with_comment(cell_ref: &str, comment_text: &str, author: &str) -> Vec<u8> {
    let cursor = Cursor::new(Vec::new());
    let mut zip = ZipWriter::new(cursor);
    let options = FileOptions::<()>::default().compression_method(zip::CompressionMethod::Deflated);

    // [Content_Types].xml
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

    // _rels/.rels
    let _ = zip.start_file("_rels/.rels", options);
    let _ = zip.write_all(
        br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>"#,
    );

    // xl/_rels/workbook.xml.rels
    let _ = zip.start_file("xl/_rels/workbook.xml.rels", options);
    let _ = zip.write_all(
        br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>"#,
    );

    // xl/workbook.xml
    let _ = zip.start_file("xl/workbook.xml", options);
    let _ = zip.write_all(
        br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<sheets>
<sheet name="Sheet1" sheetId="1" r:id="rId1"/>
</sheets>
</workbook>"#,
    );

    // xl/styles.xml
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

    // xl/worksheets/_rels/sheet1.xml.rels (links to comments file)
    let _ = zip.start_file("xl/worksheets/_rels/sheet1.xml.rels", options);
    let _ = zip.write_all(
        br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="../comments1.xml"/>
</Relationships>"#,
    );

    // xl/worksheets/sheet1.xml
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

    // xl/comments1.xml
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

/// Create an XLSX file with multiple comments
fn create_xlsx_with_multiple_comments(
    comments: &[(&str, &str, &str)], // (cell_ref, comment_text, author)
) -> Vec<u8> {
    let cursor = Cursor::new(Vec::new());
    let mut zip = ZipWriter::new(cursor);
    let options = FileOptions::<()>::default().compression_method(zip::CompressionMethod::Deflated);

    // [Content_Types].xml
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

    // _rels/.rels
    let _ = zip.start_file("_rels/.rels", options);
    let _ = zip.write_all(
        br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>"#,
    );

    // xl/_rels/workbook.xml.rels
    let _ = zip.start_file("xl/_rels/workbook.xml.rels", options);
    let _ = zip.write_all(
        br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>"#,
    );

    // xl/workbook.xml
    let _ = zip.start_file("xl/workbook.xml", options);
    let _ = zip.write_all(
        br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<sheets>
<sheet name="Sheet1" sheetId="1" r:id="rId1"/>
</sheets>
</workbook>"#,
    );

    // xl/styles.xml
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

    // xl/worksheets/_rels/sheet1.xml.rels
    let _ = zip.start_file("xl/worksheets/_rels/sheet1.xml.rels", options);
    let _ = zip.write_all(
        br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="../comments1.xml"/>
</Relationships>"#,
    );

    // xl/worksheets/sheet1.xml
    let _ = zip.start_file("xl/worksheets/sheet1.xml", options);
    let mut rows = String::new();
    for (cell_ref, _, _) in comments {
        // Extract row number from cell ref (e.g., "A1" -> 1)
        let row_num: u32 = cell_ref
            .chars()
            .skip_while(|c| c.is_ascii_alphabetic())
            .collect::<String>()
            .parse()
            .unwrap_or(1);
        rows.push_str(&format!(
            r#"<row r="{}"><c r="{}" t="inlineStr"><is><t>Cell {}</t></is></c></row>"#,
            row_num, cell_ref, cell_ref
        ));
    }
    let sheet_xml = format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<sheetData>
{rows}
</sheetData>
</worksheet>"#
    );
    let _ = zip.write_all(sheet_xml.as_bytes());

    // xl/comments1.xml
    let _ = zip.start_file("xl/comments1.xml", options);

    // Collect unique authors
    let mut authors: Vec<&str> = comments.iter().map(|(_, _, a)| *a).collect();
    authors.sort();
    authors.dedup();

    let authors_xml: String = authors
        .iter()
        .map(|a| format!("<author>{a}</author>"))
        .collect();

    let comments_xml_entries: String = comments
        .iter()
        .map(|(cell_ref, text, author)| {
            let author_id = authors.iter().position(|a| a == author).unwrap_or(0);
            format!(
                r#"<comment ref="{cell_ref}" authorId="{author_id}">
<text><r><t>{text}</t></r></text>
</comment>"#
            )
        })
        .collect();

    let comments_xml = format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<comments xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<authors>
{authors_xml}
</authors>
<commentList>
{comments_xml_entries}
</commentList>
</comments>"#
    );
    let _ = zip.write_all(comments_xml.as_bytes());

    let cursor = zip.finish().expect("Failed to finish ZIP");
    cursor.into_inner()
}

/// Create an XLSX file without any comments
fn create_xlsx_without_comments() -> Vec<u8> {
    let cursor = Cursor::new(Vec::new());
    let mut zip = ZipWriter::new(cursor);
    let options = FileOptions::<()>::default().compression_method(zip::CompressionMethod::Deflated);

    // [Content_Types].xml
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

    // _rels/.rels
    let _ = zip.start_file("_rels/.rels", options);
    let _ = zip.write_all(
        br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>"#,
    );

    // xl/_rels/workbook.xml.rels
    let _ = zip.start_file("xl/_rels/workbook.xml.rels", options);
    let _ = zip.write_all(
        br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>"#,
    );

    // xl/workbook.xml
    let _ = zip.start_file("xl/workbook.xml", options);
    let _ = zip.write_all(
        br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<sheets>
<sheet name="Sheet1" sheetId="1" r:id="rId1"/>
</sheets>
</workbook>"#,
    );

    // xl/styles.xml
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

    // xl/worksheets/sheet1.xml (no comments relationship file)
    let _ = zip.start_file("xl/worksheets/sheet1.xml", options);
    let _ = zip.write_all(
        br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<sheetData>
<row r="1"><c r="A1" t="inlineStr"><is><t>Hello</t></is></c></row>
</sheetData>
</worksheet>"#,
    );

    let cursor = zip.finish().expect("Failed to finish ZIP");
    cursor.into_inner()
}

// =============================================================================
// Tests
// =============================================================================

#[test]
fn test_parse_single_comment() {
    let xlsx = create_xlsx_with_comment("A1", "This is a test comment", "John Doe");
    let workbook = load_xlsx(&xlsx);

    assert_eq!(workbook.sheets.len(), 1);
    let sheet = &workbook.sheets[0];

    // Check that comments are parsed
    assert_eq!(sheet.comments.len(), 1);

    let comment = &sheet.comments[0];
    assert_eq!(comment.author.as_deref(), Some("John Doe"));
    assert_eq!(comment.text, "This is a test comment");

    // Verify the comment is indexed by cell ref
    assert!(sheet.comments_by_cell.contains_key("A1"));
}

#[test]
fn test_cell_has_comment_flag() {
    let xlsx = create_xlsx_with_comment("A1", "Comment text", "Author");
    let workbook = load_xlsx(&xlsx);

    let sheet = &workbook.sheets[0];

    // Find the cell at A1 (row 0, col 0)
    let cell = sheet.cells.iter().find(|c| c.r == 0 && c.c == 0);
    assert!(cell.is_some(), "Cell A1 should exist");

    let cell = cell.unwrap();
    assert_eq!(
        cell.cell.has_comment,
        Some(true),
        "Cell A1 should have has_comment = true"
    );
}

#[test]
fn test_parse_multiple_comments() {
    let comments = vec![
        ("A1", "First comment", "Alice"),
        ("B2", "Second comment", "Bob"),
        ("C3", "Third comment", "Alice"),
    ];
    let xlsx = create_xlsx_with_multiple_comments(&comments);
    let workbook = load_xlsx(&xlsx);

    let sheet = &workbook.sheets[0];

    // Check that all comments are parsed
    assert_eq!(sheet.comments.len(), 3);

    // Verify comment contents via comments_by_cell index
    // Comments are indexed by uppercase cell ref
    let idx_a1 = sheet.comments_by_cell.get("A1");
    assert!(idx_a1.is_some(), "A1 should have a comment index");
    let comment_a1 = &sheet.comments[*idx_a1.unwrap()];
    assert_eq!(comment_a1.text, "First comment");
    assert_eq!(comment_a1.author.as_deref(), Some("Alice"));

    let idx_b2 = sheet.comments_by_cell.get("B2");
    assert!(idx_b2.is_some(), "B2 should have a comment index");
    let comment_b2 = &sheet.comments[*idx_b2.unwrap()];
    assert_eq!(comment_b2.text, "Second comment");
    assert_eq!(comment_b2.author.as_deref(), Some("Bob"));

    let idx_c3 = sheet.comments_by_cell.get("C3");
    assert!(idx_c3.is_some(), "C3 should have a comment index");
    let comment_c3 = &sheet.comments[*idx_c3.unwrap()];
    assert_eq!(comment_c3.text, "Third comment");
    assert_eq!(comment_c3.author.as_deref(), Some("Alice"));
}

#[test]
fn test_multiple_comments_has_comment_flags() {
    let comments = vec![("A1", "Comment 1", "Author"), ("B2", "Comment 2", "Author")];
    let xlsx = create_xlsx_with_multiple_comments(&comments);
    let workbook = load_xlsx(&xlsx);

    let sheet = &workbook.sheets[0];

    // Check A1 (row 0, col 0)
    let cell_a1 = sheet.cells.iter().find(|c| c.r == 0 && c.c == 0);
    assert!(cell_a1.is_some());
    assert_eq!(cell_a1.unwrap().cell.has_comment, Some(true));

    // Check B2 (row 1, col 1)
    let cell_b2 = sheet.cells.iter().find(|c| c.r == 1 && c.c == 1);
    assert!(cell_b2.is_some());
    assert_eq!(cell_b2.unwrap().cell.has_comment, Some(true));
}

#[test]
fn test_sheet_without_comments() {
    let xlsx = create_xlsx_without_comments();
    let workbook = load_xlsx(&xlsx);

    let sheet = &workbook.sheets[0];

    // Sheet should have no comments
    assert!(sheet.comments.is_empty());

    // Cell should not have has_comment flag set
    let cell = sheet.cells.iter().find(|c| c.r == 0 && c.c == 0);
    assert!(cell.is_some());
    assert!(
        cell.unwrap().cell.has_comment.is_none() || cell.unwrap().cell.has_comment == Some(false)
    );
}

#[test]
fn test_comment_with_rich_text() {
    // Create a comment with multiple <r> elements (rich text)
    let cursor = Cursor::new(Vec::new());
    let mut zip = ZipWriter::new(cursor);
    let options = FileOptions::<()>::default().compression_method(zip::CompressionMethod::Deflated);

    // Basic XLSX structure
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
<sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets>
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
<sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>Cell</t></is></c></row></sheetData>
</worksheet>"#,
    );

    // Comment with multiple rich text runs
    let _ = zip.start_file("xl/comments1.xml", options);
    let _ = zip.write_all(
        br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<comments xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<authors><author>Test Author</author></authors>
<commentList>
<comment ref="A1" authorId="0">
<text>
<r><t>First part </t></r>
<r><t>second part </t></r>
<r><t>third part</t></r>
</text>
</comment>
</commentList>
</comments>"#,
    );

    let cursor = zip.finish().expect("Failed to finish ZIP");
    let xlsx = cursor.into_inner();

    let workbook = load_xlsx(&xlsx);
    let sheet = &workbook.sheets[0];

    assert_eq!(sheet.comments.len(), 1);
    let comment = &sheet.comments[0];

    // The rich text runs should be concatenated
    assert_eq!(comment.text, "First part second part third part");
}

#[test]
fn test_comment_in_workbook_structure() {
    let xlsx = create_xlsx_with_comment("A1", "Test comment", "John Doe");
    let workbook = load_xlsx(&xlsx);

    // Verify workbook has the expected structure
    assert_eq!(workbook.sheets.len(), 1);
    let sheet = &workbook.sheets[0];

    // Check comments list
    assert_eq!(sheet.comments.len(), 1);
    let comment = &sheet.comments[0];
    assert_eq!(comment.text, "Test comment");
    assert_eq!(comment.author.as_deref(), Some("John Doe"));

    // Check comments_by_cell index
    assert!(sheet.comments_by_cell.contains_key("A1"));

    // Check hasComment flag on cell
    let cell = sheet.cells.iter().find(|c| c.r == 0 && c.c == 0);
    assert!(cell.is_some());
    assert_eq!(cell.unwrap().cell.has_comment, Some(true));
}
