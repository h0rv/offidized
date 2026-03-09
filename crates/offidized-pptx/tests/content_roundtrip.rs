//! Content-verifying dirty roundtrip tests for offidized-pptx.
//!
//! These tests open real .pptx files, dirty-modify them, save, reopen,
//! and verify that the original content survived the roundtrip.

#![allow(clippy::unwrap_used)]
#![allow(clippy::expect_used)]
#![allow(clippy::panic)]

use std::path::{Path, PathBuf};

use offidized_pptx::Presentation;

fn reference_root() -> PathBuf {
    let manifest_dir = Path::new(env!("CARGO_MANIFEST_DIR"));
    manifest_dir.join("../../references")
}

fn openxml_fixture(name: &str) -> PathBuf {
    reference_root()
        .join("Open-XML-SDK/test/DocumentFormat.OpenXml.Tests.Assets/assets/TestFiles")
        .join(name)
}

fn shapecrawler_fixture(name: &str) -> PathBuf {
    reference_root()
        .join("ShapeCrawler/tests/ShapeCrawler.DevTests/Assets")
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

/// Open a file, add a slide to dirty it, save to temp, reopen.
fn dirty_roundtrip(src: &Path) -> (Presentation, Presentation, tempfile::TempDir) {
    let original = Presentation::open(src).expect("open original");

    let mut modified = Presentation::open(src).expect("open for modification");
    modified.add_slide_with_title("__DIRTY_ROUNDTRIP_SENTINEL__");

    let tmp = tempfile::tempdir().expect("create tempdir");
    let output = tmp.path().join("output.pptx");
    modified.save(&output).expect("save dirty presentation");

    let reopened = Presentation::open(&output).expect("reopen saved presentation");
    (original, reopened, tmp)
}

/// Collect all text from all shapes on a slide.
fn slide_text_fingerprint(prs: &Presentation, slide_idx: usize) -> Vec<String> {
    let slide = match prs.slide(slide_idx) {
        Some(s) => s,
        None => return Vec::new(),
    };

    let mut texts = Vec::new();
    for shape in slide.shapes() {
        let mut shape_text = String::new();
        for para in shape.paragraphs() {
            for run in para.runs() {
                shape_text.push_str(run.text());
            }
            shape_text.push('\n');
        }
        let trimmed = shape_text.trim().to_string();
        if !trimmed.is_empty() {
            texts.push(trimmed);
        }
    }
    texts
}

/// Fingerprint table content on a slide.
fn slide_table_fingerprint(
    prs: &Presentation,
    slide_idx: usize,
) -> Vec<(usize, usize, Vec<String>)> {
    let slide = match prs.slide(slide_idx) {
        Some(s) => s,
        None => return Vec::new(),
    };

    slide
        .tables()
        .iter()
        .map(|t| {
            let rows = t.rows();
            let cols = t.cols();
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
fn dirty_roundtrip_preserves_slide_count() {
    let src = openxml_fixture("mcppt.pptx");
    if skip_if_missing(&src) {
        return;
    }

    let (original, reopened, _tmp) = dirty_roundtrip(&src);

    // Reopened should have original slides + 1 sentinel slide.
    assert_eq!(
        reopened.slide_count(),
        original.slide_count() + 1,
        "expected original + 1 sentinel slide"
    );
}

#[test]
fn dirty_roundtrip_preserves_slide_text() {
    let src = openxml_fixture("mcppt.pptx");
    if skip_if_missing(&src) {
        return;
    }

    let (original, reopened, _tmp) = dirty_roundtrip(&src);

    for i in 0..original.slide_count() {
        let orig_texts = slide_text_fingerprint(&original, i);
        let rt_texts = slide_text_fingerprint(&reopened, i);

        assert_eq!(
            orig_texts, rt_texts,
            "slide {i} text content changed after roundtrip"
        );
    }
}

#[test]
fn dirty_roundtrip_preserves_shapecrawler_table() {
    let src = shapecrawler_fixture("tables/table-case001.pptx");
    if skip_if_missing(&src) {
        return;
    }

    let (original, reopened, _tmp) = dirty_roundtrip(&src);

    for i in 0..original.slide_count() {
        let orig_tables = slide_table_fingerprint(&original, i);
        let rt_tables = slide_table_fingerprint(&reopened, i);

        assert_eq!(
            orig_tables.len(),
            rt_tables.len(),
            "slide {i}: table count changed"
        );

        for (j, (orig, rt)) in orig_tables.iter().zip(rt_tables.iter()).enumerate() {
            assert_eq!(orig.0, rt.0, "slide {i}, table {j}: row count changed");
            assert_eq!(orig.1, rt.1, "slide {i}, table {j}: col count changed");
            assert_eq!(orig.2, rt.2, "slide {i}, table {j}: cell contents changed");
        }
    }
}

#[test]
fn dirty_roundtrip_preserves_animation_pptx() {
    let src = openxml_fixture("animation.pptx");
    if skip_if_missing(&src) {
        return;
    }

    let (original, reopened, _tmp) = dirty_roundtrip(&src);

    for i in 0..original.slide_count() {
        let orig_texts = slide_text_fingerprint(&original, i);
        let rt_texts = slide_text_fingerprint(&reopened, i);

        assert_eq!(
            orig_texts, rt_texts,
            "slide {i} text content changed after roundtrip"
        );
    }
}

// ── Bulk corpus tests ────────────────────────────────────────────────

fn run_bulk_roundtrip(fixture_dir: &Path, label: &str) {
    if !fixture_dir.is_dir() {
        eprintln!(
            "skipping {label} bulk test: fixture dir not found at `{}`",
            fixture_dir.display()
        );
        return;
    }

    let mut files: Vec<PathBuf> = walkdir(fixture_dir);
    files.sort();

    assert!(!files.is_empty(), "no .pptx fixtures found in {label}");

    let mut passed = 0;
    let mut failed = Vec::new();
    let mut skipped = Vec::new();

    for file in &files {
        let name = file
            .strip_prefix(fixture_dir)
            .unwrap_or(file)
            .display()
            .to_string();

        let original = match Presentation::open(file) {
            Ok(prs) => prs,
            Err(e) => {
                eprintln!("  skip {name}: open failed: {e}");
                skipped.push(name);
                continue;
            }
        };

        let mut modified = match Presentation::open(file) {
            Ok(prs) => prs,
            Err(_) => {
                skipped.push(name);
                continue;
            }
        };

        modified.add_slide_with_title("__BULK_SENTINEL__");

        let tmp = tempfile::tempdir().expect("create tempdir");
        let output = tmp.path().join("output.pptx");

        if let Err(e) = modified.save(&output) {
            eprintln!("  skip {name}: save failed: {e}");
            skipped.push(name);
            continue;
        }

        let reopened = match Presentation::open(&output) {
            Ok(prs) => prs,
            Err(e) => {
                eprintln!("  FAIL {name}: reopen failed: {e}");
                failed.push(name);
                continue;
            }
        };

        let mut content_ok = true;

        // Verify original slide count is preserved (reopened = original + 1 sentinel).
        if reopened.slide_count() != original.slide_count() + 1 {
            eprintln!(
                "  FAIL {name}: slide count: expected {} + 1, got {}",
                original.slide_count(),
                reopened.slide_count()
            );
            content_ok = false;
        }

        // Verify text content on original slides.
        if content_ok {
            for i in 0..original.slide_count() {
                let orig_texts = slide_text_fingerprint(&original, i);
                let rt_texts = slide_text_fingerprint(&reopened, i);

                if orig_texts != rt_texts {
                    eprintln!("  FAIL {name}: slide {i} text changed");
                    content_ok = false;
                    break;
                }
            }
        }

        // Verify table content on original slides.
        if content_ok {
            for i in 0..original.slide_count() {
                let orig_tables = slide_table_fingerprint(&original, i);
                let rt_tables = slide_table_fingerprint(&reopened, i);

                if orig_tables != rt_tables {
                    eprintln!("  FAIL {name}: slide {i} table content changed");
                    content_ok = false;
                    break;
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
        "\n=== {label} bulk roundtrip: {passed} passed, {} failed, {} skipped out of {} total ===",
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
        "{} out of {} pptx files failed content roundtrip in {label}",
        failed.len(),
        files.len()
    );
}

/// Recursively find all .pptx files.
fn walkdir(dir: &Path) -> Vec<PathBuf> {
    let mut results = Vec::new();
    if let Ok(entries) = std::fs::read_dir(dir) {
        for entry in entries.flatten() {
            let path = entry.path();
            if path.is_dir() {
                results.extend(walkdir(&path));
            } else if path
                .extension()
                .map(|e| e.eq_ignore_ascii_case("pptx"))
                .unwrap_or(false)
            {
                results.push(path);
            }
        }
    }
    results
}

#[test]
fn bulk_openxml_sdk_dirty_roundtrip_content_verification() {
    let fixture_dir = reference_root()
        .join("Open-XML-SDK/test/DocumentFormat.OpenXml.Tests.Assets/assets/TestFiles");
    run_bulk_roundtrip(&fixture_dir, "OpenXML-SDK");
}

#[test]
fn bulk_shapecrawler_dirty_roundtrip_content_verification() {
    let fixture_dir = reference_root().join("ShapeCrawler/tests/ShapeCrawler.DevTests/Assets");
    run_bulk_roundtrip(&fixture_dir, "ShapeCrawler");
}
