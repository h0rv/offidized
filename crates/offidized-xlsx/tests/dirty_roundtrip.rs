//! Dirty roundtrip tests: open a real file, modify a cell (making the worksheet
//! dirty), save, reopen, and verify that all OPC parts survive and the workbook
//! is still valid.
//!
//! These tests specifically catch regressions where the worksheet serializer
//! drops elements or relationships it doesn't model (charts, VML drawings,
//! sparklines, conditional formatting extensions, etc.).

#![allow(clippy::unwrap_used)]
#![allow(clippy::expect_used)]

use std::collections::BTreeSet;
use std::path::{Path, PathBuf};

use offidized_opc::Package;
use offidized_xlsx::{CellValue, Workbook};

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

fn reference_root() -> PathBuf {
    let manifest_dir = Path::new(env!("CARGO_MANIFEST_DIR"));
    manifest_dir.join("../../references")
}

fn openxml_fixture(name: &str) -> PathBuf {
    reference_root()
        .join("Open-XML-SDK/test/DocumentFormat.OpenXml.Tests.Assets/assets/TestFiles")
        .join(name)
}

fn closedxml_fixture(rel: &str) -> PathBuf {
    reference_root()
        .join("ClosedXML/ClosedXML.Tests/Resource")
        .join(rel)
}

fn skip_if_missing(path: &Path) -> bool {
    if !path.is_file() {
        eprintln!("skipping: fixture not found at `{}`", path.display());
        true
    } else {
        false
    }
}

/// Return the set of part URIs in a package.
fn part_uri_set(path: &Path) -> BTreeSet<String> {
    let pkg = Package::open(path).unwrap();
    pkg.part_uris().into_iter().map(String::from).collect()
}

/// Check if a specific part's XML contains a given substring.
fn part_xml_contains(path: &Path, part_uri: &str, needle: &str) -> bool {
    let pkg = Package::open(path).unwrap();
    pkg.get_part(part_uri)
        .and_then(|p| p.as_str())
        .is_some_and(|xml| xml.contains(needle))
}

/// Open a workbook, modify cell A1 on the first sheet, save to a temp file,
/// reopen and return (original_parts, output_path, _tmpdir, reopened_workbook).
fn dirty_roundtrip(src: &Path) -> (BTreeSet<String>, PathBuf, tempfile::TempDir, Workbook) {
    let original_parts = part_uri_set(src);

    let mut wb = Workbook::open(src).expect("open workbook");
    let sheet_name = wb.sheet_names()[0].to_string();
    let ws = wb.sheet_mut(&sheet_name).expect("get first sheet");
    ws.cell_mut("A1")
        .expect("get cell A1")
        .set_value(CellValue::String("DIRTY_ROUNDTRIP_TEST".to_string()));

    let tmp = tempfile::tempdir().expect("create tempdir");
    let output = tmp.path().join("output.xlsx");
    wb.save(&output).expect("save dirty workbook");

    let reopened = Workbook::open(&output).expect("reopen saved workbook");
    (original_parts, output, tmp, reopened)
}

/// Assert that no original parts are missing from the output.
fn assert_no_parts_lost(original: &BTreeSet<String>, output: &BTreeSet<String>, label: &str) {
    let lost: BTreeSet<_> = original.difference(output).collect();
    assert!(
        lost.is_empty(),
        "{label}: lost {} parts: {:?}",
        lost.len(),
        lost,
    );
}

// ---------------------------------------------------------------------------
// Chart preservation (regression for drawing relationship passthrough)
// ---------------------------------------------------------------------------

#[test]
fn dirty_roundtrip_preserves_chart_parts() {
    let src = closedxml_fixture("Other/Charts/PreserveCharts/inputfile.xlsx");
    if skip_if_missing(&src) {
        return;
    }

    let (orig_parts, output, _tmp, reopened) = dirty_roundtrip(&src);
    let out_parts = part_uri_set(&output);

    assert_no_parts_lost(&orig_parts, &out_parts, "charts");

    // Chart parts must survive.
    assert!(
        out_parts.iter().any(|p| p.contains("chart")),
        "chart parts missing from output"
    );

    // The drawing part must still exist.
    assert!(
        out_parts.iter().any(|p| p.contains("drawing")),
        "drawing parts missing from output"
    );

    // Verify the workbook reopened with both sheets.
    assert_eq!(reopened.sheet_names().len(), 2);
}

// ---------------------------------------------------------------------------
// VML / legacyDrawing preservation (regression for comments roundtrip)
// ---------------------------------------------------------------------------

#[test]
fn dirty_roundtrip_preserves_vml_drawing_for_comments() {
    let src = openxml_fixture("Comments.xlsx");
    if skip_if_missing(&src) {
        return;
    }

    let (orig_parts, output, _tmp, _reopened) = dirty_roundtrip(&src);
    let out_parts = part_uri_set(&output);

    assert_no_parts_lost(&orig_parts, &out_parts, "comments");

    // The legacyDrawing element must survive in the worksheet XML.
    assert!(
        part_xml_contains(&output, "/xl/worksheets/sheet1.xml", "legacyDrawing"),
        "sheet1 lost <legacyDrawing> element",
    );

    // VML drawing part must still exist.
    assert!(
        out_parts.iter().any(|p| p.contains("vmlDrawing")),
        "vmlDrawing part missing from output",
    );
}

// ---------------------------------------------------------------------------
// Complex file with charts + diagrams + images
// ---------------------------------------------------------------------------

#[test]
fn dirty_roundtrip_complex01() {
    let src = openxml_fixture("Complex01.xlsx");
    if skip_if_missing(&src) {
        return;
    }

    let (orig_parts, output, _tmp, reopened) = dirty_roundtrip(&src);
    let out_parts = part_uri_set(&output);

    assert_no_parts_lost(&orig_parts, &out_parts, "complex01");
    assert_eq!(reopened.sheet_names().len(), 2);
}

// ---------------------------------------------------------------------------
// Sparklines (stored in extLst which should be an unknown child)
// ---------------------------------------------------------------------------

#[test]
fn dirty_roundtrip_preserves_sparklines() {
    let src = closedxml_fixture("Examples/Sparklines/SampleSparklines.xlsx");
    if skip_if_missing(&src) {
        return;
    }

    let (orig_parts, output, _tmp, _reopened) = dirty_roundtrip(&src);
    let out_parts = part_uri_set(&output);

    assert_no_parts_lost(&orig_parts, &out_parts, "sparklines");

    // Sparklines live in extLst.
    assert!(
        part_xml_contains(&output, "/xl/worksheets/sheet1.xml", "extLst"),
        "sheet1 lost <extLst> containing sparklines",
    );
}

// ---------------------------------------------------------------------------
// Conditional formatting (data bars)
// ---------------------------------------------------------------------------

#[test]
fn dirty_roundtrip_preserves_conditional_formatting() {
    let src = closedxml_fixture("Examples/ConditionalFormatting/CFDataBar.xlsx");
    if skip_if_missing(&src) {
        return;
    }

    let (orig_parts, output, _tmp, _reopened) = dirty_roundtrip(&src);
    let out_parts = part_uri_set(&output);

    assert_no_parts_lost(&orig_parts, &out_parts, "cf_databar");

    assert!(
        part_xml_contains(
            &output,
            "/xl/worksheets/sheet1.xml",
            "conditionalFormatting"
        ),
        "sheet1 lost conditionalFormatting",
    );
}

// ---------------------------------------------------------------------------
// Images / drawings
// ---------------------------------------------------------------------------

#[test]
fn dirty_roundtrip_preserves_images() {
    let src = closedxml_fixture("Examples/ImageHandling/ImageFormats.xlsx");
    if skip_if_missing(&src) {
        return;
    }

    let (orig_parts, output, _tmp, _reopened) = dirty_roundtrip(&src);
    let out_parts = part_uri_set(&output);

    assert_no_parts_lost(&orig_parts, &out_parts, "images");

    // Media parts must survive (may include duplicates from re-serialized sheets).
    let media_count = out_parts.iter().filter(|p| p.contains("media/")).count();
    let orig_media = orig_parts.iter().filter(|p| p.contains("media/")).count();
    assert!(
        media_count >= orig_media,
        "media parts lost: output has {media_count} but original had {orig_media}",
    );
}

// ---------------------------------------------------------------------------
// Pivot tables
// ---------------------------------------------------------------------------

#[test]
fn dirty_roundtrip_preserves_pivot_tables() {
    let src = closedxml_fixture("Examples/PivotTables/PivotTables.xlsx");
    if skip_if_missing(&src) {
        return;
    }

    let (orig_parts, output, _tmp, _reopened) = dirty_roundtrip(&src);
    let out_parts = part_uri_set(&output);

    assert_no_parts_lost(&orig_parts, &out_parts, "pivottables");

    let pivot_count = out_parts
        .iter()
        .filter(|p| p.contains("pivotTable"))
        .count();
    assert!(pivot_count > 0, "pivot table parts missing");
}

// ---------------------------------------------------------------------------
// Tables
// ---------------------------------------------------------------------------

#[test]
fn dirty_roundtrip_preserves_tables() {
    let src = closedxml_fixture("Examples/Tables/UsingTables.xlsx");
    if skip_if_missing(&src) {
        return;
    }

    let (_orig_parts, _output, _tmp, reopened) = dirty_roundtrip(&src);

    let table_count: usize = reopened
        .worksheets()
        .iter()
        .map(|ws| ws.tables().len())
        .sum();
    assert!(table_count > 0, "no tables found after dirty roundtrip");
}

// ---------------------------------------------------------------------------
// Hyperlinks
// ---------------------------------------------------------------------------

#[test]
fn dirty_roundtrip_preserves_hyperlinks() {
    let src = closedxml_fixture("Examples/Misc/Hyperlinks.xlsx");
    if skip_if_missing(&src) {
        return;
    }

    let (orig_parts, output, _tmp, reopened) = dirty_roundtrip(&src);
    let out_parts = part_uri_set(&output);

    assert_no_parts_lost(&orig_parts, &out_parts, "hyperlinks");

    let hyperlink_count: usize = reopened
        .worksheets()
        .iter()
        .map(|ws| ws.hyperlinks().len())
        .sum();
    assert!(
        hyperlink_count > 0,
        "no hyperlinks found after dirty roundtrip"
    );
}

// ---------------------------------------------------------------------------
// Data validation
// ---------------------------------------------------------------------------

#[test]
fn dirty_roundtrip_preserves_data_validation() {
    let src = closedxml_fixture("Examples/Misc/DataValidation.xlsx");
    if skip_if_missing(&src) {
        return;
    }

    let (orig_parts, output, _tmp, _reopened) = dirty_roundtrip(&src);
    let out_parts = part_uri_set(&output);

    assert_no_parts_lost(&orig_parts, &out_parts, "datavalidation");

    assert!(
        part_xml_contains(&output, "/xl/worksheets/sheet1.xml", "dataValidation"),
        "sheet1 lost dataValidation elements",
    );
}

// ---------------------------------------------------------------------------
// Auto filter
// ---------------------------------------------------------------------------

#[test]
fn dirty_roundtrip_preserves_autofilter() {
    let src = closedxml_fixture("Examples/AutoFilter/RegularAutoFilter.xlsx");
    if skip_if_missing(&src) {
        return;
    }

    let (orig_parts, output, _tmp, reopened) = dirty_roundtrip(&src);
    let out_parts = part_uri_set(&output);

    assert_no_parts_lost(&orig_parts, &out_parts, "autofilter");

    let has_af = reopened
        .worksheets()
        .iter()
        .any(|ws| ws.auto_filter().is_some());
    assert!(has_af, "auto filters lost after dirty roundtrip");
}

// ---------------------------------------------------------------------------
// Formulas
// ---------------------------------------------------------------------------

#[test]
fn dirty_roundtrip_preserves_formulas() {
    let src = closedxml_fixture("Examples/Misc/Formulas.xlsx");
    if skip_if_missing(&src) {
        return;
    }

    let (orig_parts, output, _tmp, _reopened) = dirty_roundtrip(&src);
    let out_parts = part_uri_set(&output);

    assert_no_parts_lost(&orig_parts, &out_parts, "formulas");
}

// ---------------------------------------------------------------------------
// Merge cells
// ---------------------------------------------------------------------------

#[test]
fn dirty_roundtrip_preserves_merged_cells() {
    let src = closedxml_fixture("Examples/Misc/MergeCells.xlsx");
    if skip_if_missing(&src) {
        return;
    }

    let (_orig_parts, _output, _tmp, reopened) = dirty_roundtrip(&src);

    let merge_count: usize = reopened
        .worksheets()
        .iter()
        .map(|ws| ws.merged_ranges().len())
        .sum();
    assert!(
        merge_count > 0,
        "no merged ranges found after dirty roundtrip"
    );
}

// ---------------------------------------------------------------------------
// Sheet protection
// ---------------------------------------------------------------------------

#[test]
fn dirty_roundtrip_preserves_protection() {
    let src = closedxml_fixture("Examples/Misc/SheetProtection.xlsx");
    if skip_if_missing(&src) {
        return;
    }

    let (orig_parts, output, _tmp, _reopened) = dirty_roundtrip(&src);
    let out_parts = part_uri_set(&output);

    assert_no_parts_lost(&orig_parts, &out_parts, "protection");

    assert!(
        part_xml_contains(&output, "/xl/worksheets/sheet1.xml", "sheetProtection"),
        "sheet1 lost sheetProtection",
    );
}

// ---------------------------------------------------------------------------
// External links
// ---------------------------------------------------------------------------

#[test]
fn dirty_roundtrip_preserves_external_links() {
    let src = closedxml_fixture("Other/ExternalLinks/WorkbookWithExternalLink.xlsx");
    if skip_if_missing(&src) {
        return;
    }

    let (orig_parts, output, _tmp, _reopened) = dirty_roundtrip(&src);
    let out_parts = part_uri_set(&output);

    assert_no_parts_lost(&orig_parts, &out_parts, "extlinks");
}

// ---------------------------------------------------------------------------
// Bulk dirty roundtrip: modify cell on EVERY sheet
// ---------------------------------------------------------------------------

#[test]
fn dirty_roundtrip_all_sheets_modified() {
    let src = closedxml_fixture("Examples/Misc/ShowCase.xlsx");
    if skip_if_missing(&src) {
        return;
    }

    let original_parts = part_uri_set(&src);

    let mut wb = Workbook::open(&src).expect("open showcase");
    let names: Vec<String> = wb.sheet_names().into_iter().map(String::from).collect();
    for name in &names {
        let ws = wb.sheet_mut(name).expect("get sheet");
        ws.cell_mut("A1")
            .expect("get A1")
            .set_value(CellValue::String(format!("DIRTY_{name}")));
    }

    let tmp = tempfile::tempdir().expect("tempdir");
    let output = tmp.path().join("showcase_dirty.xlsx");
    wb.save(&output).expect("save");

    let out_parts = part_uri_set(&output);
    assert_no_parts_lost(&original_parts, &out_parts, "showcase_all_dirty");

    let reopened = Workbook::open(&output).expect("reopen");
    assert_eq!(reopened.sheet_names().len(), names.len());
}
