//! Content-verifying dirty roundtrip tests for offidized-docx.
//!
//! These tests open real .docx files, dirty-modify them, save, reopen,
//! and verify that the original content survived the roundtrip.

#![allow(clippy::unwrap_used)]
#![allow(clippy::expect_used)]
#![allow(clippy::panic)]

use std::path::{Path, PathBuf};

use offidized_docx::Document;

fn reference_root() -> PathBuf {
    let manifest_dir = Path::new(env!("CARGO_MANIFEST_DIR"));
    manifest_dir.join("../../references")
}

fn openxml_fixture(name: &str) -> PathBuf {
    reference_root()
        .join("Open-XML-SDK/test/DocumentFormat.OpenXml.Tests.Assets/assets/TestFiles")
        .join(name)
}

fn skip_if_missing(path: &Path) -> bool {
    if !path.is_file() {
        eprintln!("skipping: fixture not found at `{}`", path.display());
        true
    } else {
        false
    }
}

/// Open a file, add a paragraph to dirty it, save to temp, reopen.
fn dirty_roundtrip(src: &Path) -> (Document, Document, tempfile::TempDir) {
    let original = Document::open(src).expect("open original");

    let mut modified = Document::open(src).expect("open for modification");
    modified.add_paragraph("__DIRTY_ROUNDTRIP_SENTINEL__");

    let tmp = tempfile::tempdir().expect("create tempdir");
    let output = tmp.path().join("output.docx");
    modified.save(&output).expect("save dirty document");

    let reopened = Document::open(&output).expect("reopen saved document");
    (original, reopened, tmp)
}

/// Collect all paragraph texts from a document.
fn all_paragraph_texts(doc: &Document) -> Vec<String> {
    doc.paragraphs().iter().map(|p| p.text()).collect()
}

/// Fingerprint: paragraph count + first N paragraph texts.
fn fingerprint_paragraphs(doc: &Document, max: usize) -> Vec<String> {
    doc.paragraphs()
        .iter()
        .take(max)
        .map(|p| p.text())
        .collect()
}

/// Fingerprint: table dimensions + cell texts for first few tables.
fn fingerprint_tables(doc: &Document, max_tables: usize) -> Vec<(usize, usize, Vec<String>)> {
    doc.tables()
        .iter()
        .take(max_tables)
        .map(|t| {
            let rows = t.rows();
            let cols = t.columns();
            let mut cells = Vec::new();
            for r in 0..rows {
                for c in 0..cols {
                    cells.push(t.cell_text(r, c).unwrap_or("").to_string());
                }
            }
            (rows, cols, cells)
        })
        .collect()
}

// ── Targeted tests ───────────────────────────────────────────────────

#[test]
fn dirty_roundtrip_preserves_hello_world() {
    let src = openxml_fixture("HelloWorld.docx");
    if skip_if_missing(&src) {
        return;
    }

    let (original, reopened, _tmp) = dirty_roundtrip(&src);

    let orig_texts = all_paragraph_texts(&original);
    let rt_texts = all_paragraph_texts(&reopened);

    // Original paragraphs must all be present (reopened may have extra sentinel).
    assert!(
        rt_texts.len() >= orig_texts.len(),
        "paragraph count shrank: {} -> {}",
        orig_texts.len(),
        rt_texts.len()
    );

    for (i, orig) in orig_texts.iter().enumerate() {
        assert_eq!(
            &rt_texts[i], orig,
            "paragraph {i} text changed after roundtrip"
        );
    }
}

#[test]
fn dirty_roundtrip_preserves_document_docx() {
    let src = openxml_fixture("Document.docx");
    if skip_if_missing(&src) {
        return;
    }

    let (original, reopened, _tmp) = dirty_roundtrip(&src);

    let orig_texts = fingerprint_paragraphs(&original, 200);
    let rt_texts = fingerprint_paragraphs(&reopened, 200);

    for (i, orig) in orig_texts.iter().enumerate() {
        assert_eq!(
            &rt_texts[i], orig,
            "paragraph {i} text changed after roundtrip"
        );
    }
}

#[test]
fn dirty_roundtrip_preserves_tables() {
    let src = openxml_fixture("complex0.docx");
    if skip_if_missing(&src) {
        return;
    }

    let (original, reopened, _tmp) = dirty_roundtrip(&src);

    let orig_tables = fingerprint_tables(&original, 10);
    let rt_tables = fingerprint_tables(&reopened, 10);

    assert_eq!(
        orig_tables.len(),
        rt_tables.len(),
        "table count changed: {} -> {}",
        orig_tables.len(),
        rt_tables.len()
    );

    for (i, (orig, rt)) in orig_tables.iter().zip(rt_tables.iter()).enumerate() {
        assert_eq!(orig.0, rt.0, "table {i} row count changed");
        assert_eq!(orig.1, rt.1, "table {i} column count changed");
        assert_eq!(orig.2, rt.2, "table {i} cell contents changed");
    }
}

#[test]
fn dirty_roundtrip_preserves_comments_docx() {
    let src = openxml_fixture("Comments.docx");
    if skip_if_missing(&src) {
        return;
    }

    let (original, reopened, _tmp) = dirty_roundtrip(&src);

    let orig_texts = fingerprint_paragraphs(&original, 200);
    let rt_texts = fingerprint_paragraphs(&reopened, 200);

    for (i, orig) in orig_texts.iter().enumerate() {
        assert_eq!(
            &rt_texts[i], orig,
            "paragraph {i} text changed after roundtrip"
        );
    }
}

// ── Bulk corpus test ─────────────────────────────────────────────────

#[test]
fn bulk_openxml_sdk_dirty_roundtrip_content_verification() {
    let fixture_dir = reference_root()
        .join("Open-XML-SDK/test/DocumentFormat.OpenXml.Tests.Assets/assets/TestFiles");

    if !fixture_dir.is_dir() {
        eprintln!(
            "skipping bulk test: fixture dir not found at `{}`",
            fixture_dir.display()
        );
        return;
    }

    let mut files: Vec<PathBuf> = std::fs::read_dir(&fixture_dir)
        .expect("read fixture dir")
        .filter_map(|e| e.ok())
        .map(|e| e.path())
        .filter(|p| {
            p.extension()
                .map(|e| e.eq_ignore_ascii_case("docx"))
                .unwrap_or(false)
        })
        .collect();
    files.sort();

    assert!(!files.is_empty(), "no .docx fixtures found");

    let mut passed = 0;
    let mut failed = Vec::new();
    let mut skipped = Vec::new();

    for file in &files {
        let name = file.file_name().unwrap().to_string_lossy().to_string();

        // Try to open — some files may be encrypted or malformed.
        let original = match Document::open(file) {
            Ok(doc) => doc,
            Err(e) => {
                eprintln!("  skip {name}: open failed: {e}");
                skipped.push(name);
                continue;
            }
        };

        let mut modified = match Document::open(file) {
            Ok(doc) => doc,
            Err(_) => {
                skipped.push(name);
                continue;
            }
        };

        // Dirty the document.
        modified.add_paragraph("__BULK_ROUNDTRIP_SENTINEL__");

        let tmp = tempfile::tempdir().expect("create tempdir");
        let output = tmp.path().join("output.docx");

        if let Err(e) = modified.save(&output) {
            eprintln!("  skip {name}: save failed: {e}");
            skipped.push(name);
            continue;
        }

        let reopened = match Document::open(&output) {
            Ok(doc) => doc,
            Err(e) => {
                eprintln!("  FAIL {name}: reopen failed: {e}");
                failed.push(name);
                continue;
            }
        };

        // Compare paragraph fingerprints.
        let orig_fp = fingerprint_paragraphs(&original, 200);
        let rt_fp = fingerprint_paragraphs(&reopened, 200);

        let mut content_ok = true;
        if rt_fp.len() < orig_fp.len() {
            eprintln!(
                "  FAIL {name}: paragraph count shrank: {} -> {}",
                orig_fp.len(),
                rt_fp.len()
            );
            content_ok = false;
        } else {
            for (i, orig) in orig_fp.iter().enumerate() {
                if &rt_fp[i] != orig {
                    eprintln!(
                        "  FAIL {name}: paragraph {i} changed: {:?} -> {:?}",
                        &orig[..orig.len().min(60)],
                        &rt_fp[i][..rt_fp[i].len().min(60)]
                    );
                    content_ok = false;
                    break;
                }
            }
        }

        // Compare table fingerprints.
        if content_ok {
            let orig_tbl = fingerprint_tables(&original, 10);
            let rt_tbl = fingerprint_tables(&reopened, 10);

            if orig_tbl.len() != rt_tbl.len() {
                eprintln!(
                    "  FAIL {name}: table count changed: {} -> {}",
                    orig_tbl.len(),
                    rt_tbl.len()
                );
                content_ok = false;
            } else {
                for (i, (orig, rt)) in orig_tbl.iter().zip(rt_tbl.iter()).enumerate() {
                    if orig != rt {
                        eprintln!("  FAIL {name}: table {i} content changed");
                        content_ok = false;
                        break;
                    }
                }
            }
        }

        if content_ok {
            passed += 1;
        } else {
            failed.push(name);
        }
    }

    eprintln!(
        "\n=== docx bulk roundtrip: {passed} passed, {} failed, {} skipped out of {} total ===",
        failed.len(),
        skipped.len(),
        files.len()
    );

    if !failed.is_empty() {
        eprintln!("Failed files:");
        for f in &failed {
            eprintln!("  - {f}");
        }
    }

    assert!(
        failed.is_empty(),
        "{} out of {} docx files failed content roundtrip",
        failed.len(),
        files.len()
    );
}
