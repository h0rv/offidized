//! Content-verifying dirty roundtrip tests.
//!
//! Unlike the existing dirty_roundtrip.rs tests which only check "do ZIP parts
//! still exist", these tests verify that actual cell values, formulas, and types
//! survive a dirty save + reopen cycle.
//!
//! This is the real test of roundtrip fidelity.

#![allow(clippy::unwrap_used)]
#![allow(clippy::expect_used)]
#![allow(clippy::panic)]

use std::path::{Path, PathBuf};

use offidized_xlsx::{CellValue, Workbook};

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

fn reference_root() -> PathBuf {
    let manifest_dir = Path::new(env!("CARGO_MANIFEST_DIR"));
    manifest_dir.join("../../references")
}

fn closedxml_fixture(rel: &str) -> PathBuf {
    reference_root()
        .join("ClosedXML/ClosedXML.Tests/Resource")
        .join(rel)
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

/// Open a file, dirty cell A1 on first sheet, save to temp, reopen.
fn dirty_roundtrip(src: &Path) -> (Workbook, Workbook, tempfile::TempDir) {
    let original = Workbook::open(src).expect("open original");

    let mut modified = Workbook::open(src).expect("open for modification");
    let sheet_name = modified.sheet_names()[0].to_string();
    let ws = modified.sheet_mut(&sheet_name).expect("get first sheet");
    ws.cell_mut("ZZ1")
        .expect("get cell ZZ1")
        .set_value(CellValue::String("DIRTY".into()));

    let tmp = tempfile::tempdir().expect("create tempdir");
    let output = tmp.path().join("output.xlsx");
    modified.save(&output).expect("save dirty workbook");

    let reopened = Workbook::open(&output).expect("reopen saved workbook");
    (original, reopened, tmp)
}

/// Assert a cell has a specific string value.
fn assert_cell_str(wb: &Workbook, sheet: &str, cell_ref: &str, expected: &str, ctx: &str) {
    let ws = wb
        .sheet(sheet)
        .unwrap_or_else(|| panic!("{ctx}: sheet '{sheet}' not found"));
    let cell = ws
        .cell(cell_ref)
        .unwrap_or_else(|| panic!("{ctx}: cell {cell_ref} not found"));
    match cell.value() {
        Some(CellValue::String(s)) => {
            assert_eq!(s.as_str(), expected, "{ctx}: {cell_ref} value mismatch")
        }
        other => panic!("{ctx}: {cell_ref} expected String(\"{expected}\"), got {other:?}"),
    }
}

/// Assert a cell has a specific numeric value.
fn assert_cell_num(wb: &Workbook, sheet: &str, cell_ref: &str, expected: f64, ctx: &str) {
    let ws = wb
        .sheet(sheet)
        .unwrap_or_else(|| panic!("{ctx}: sheet '{sheet}' not found"));
    let cell = ws
        .cell(cell_ref)
        .unwrap_or_else(|| panic!("{ctx}: cell {cell_ref} not found"));
    match cell.value() {
        Some(CellValue::Number(n)) => {
            assert!(
                (n - expected).abs() < 1e-6,
                "{ctx}: {cell_ref} expected {expected}, got {n}"
            );
        }
        Some(CellValue::DateTime(n)) => {
            assert!(
                (n - expected).abs() < 1e-6,
                "{ctx}: {cell_ref} expected {expected}, got DateTime({n})"
            );
        }
        other => panic!("{ctx}: {cell_ref} expected Number({expected}), got {other:?}"),
    }
}

/// Assert a cell has a specific bool value.
fn assert_cell_bool(wb: &Workbook, sheet: &str, cell_ref: &str, expected: bool, ctx: &str) {
    let ws = wb
        .sheet(sheet)
        .unwrap_or_else(|| panic!("{ctx}: sheet '{sheet}' not found"));
    let cell = ws
        .cell(cell_ref)
        .unwrap_or_else(|| panic!("{ctx}: cell {cell_ref} not found"));
    match cell.value() {
        Some(CellValue::Bool(b)) => assert_eq!(*b, expected, "{ctx}: {cell_ref} value mismatch"),
        other => panic!("{ctx}: {cell_ref} expected Bool({expected}), got {other:?}"),
    }
}

/// Assert a cell has a specific formula.
fn assert_cell_formula(wb: &Workbook, sheet: &str, cell_ref: &str, expected: &str, ctx: &str) {
    let ws = wb
        .sheet(sheet)
        .unwrap_or_else(|| panic!("{ctx}: sheet '{sheet}' not found"));
    let cell = ws
        .cell(cell_ref)
        .unwrap_or_else(|| panic!("{ctx}: cell {cell_ref} not found"));
    let formula = cell
        .formula()
        .unwrap_or_else(|| panic!("{ctx}: {cell_ref} has no formula"));
    assert_eq!(formula, expected, "{ctx}: {cell_ref} formula mismatch");
}

/// Assert a formula cell's cached value is a specific number.
fn assert_cell_cached_num(wb: &Workbook, sheet: &str, cell_ref: &str, expected: f64, ctx: &str) {
    let ws = wb
        .sheet(sheet)
        .unwrap_or_else(|| panic!("{ctx}: sheet '{sheet}' not found"));
    let cell = ws
        .cell(cell_ref)
        .unwrap_or_else(|| panic!("{ctx}: cell {cell_ref} not found"));
    match cell.cached_value() {
        Some(CellValue::Number(n)) => {
            assert!(
                (n - expected).abs() < 1e-6,
                "{ctx}: {cell_ref} cached expected {expected}, got {n}"
            );
        }
        other => panic!("{ctx}: {cell_ref} expected cached Number({expected}), got {other:?}"),
    }
}

/// Assert a cell has an error value.
fn assert_cell_error(wb: &Workbook, sheet: &str, cell_ref: &str, expected: &str, ctx: &str) {
    let ws = wb
        .sheet(sheet)
        .unwrap_or_else(|| panic!("{ctx}: sheet '{sheet}' not found"));
    let cell = ws
        .cell(cell_ref)
        .unwrap_or_else(|| panic!("{ctx}: cell {cell_ref} not found"));
    match cell.value() {
        Some(CellValue::Error(e)) => {
            assert_eq!(e.as_str(), expected, "{ctx}: {cell_ref} error mismatch")
        }
        other => panic!("{ctx}: {cell_ref} expected Error(\"{expected}\"), got {other:?}"),
    }
}

// ---------------------------------------------------------------------------
// Tests
// ---------------------------------------------------------------------------

/// CellValues.xlsx has strings, numbers, bools, dates, errors.
/// Verify ALL of these survive a dirty roundtrip.
#[test]
fn dirty_roundtrip_preserves_cell_values() {
    let src = closedxml_fixture("Examples/Misc/CellValues.xlsx");
    if skip_if_missing(&src) {
        return;
    }

    let (original, reopened, _tmp) = dirty_roundtrip(&src);
    let sheet = "Cell Values";

    // Headers
    assert_cell_str(&original, sheet, "B2", "Initial Value", "original");
    assert_cell_str(&reopened, sheet, "B2", "Initial Value", "roundtrip");

    // Number: 1234.567
    assert_cell_num(&original, sheet, "B5", 1234.567, "original");
    assert_cell_num(&reopened, sheet, "B5", 1234.567, "roundtrip");

    // Bool: true
    assert_cell_bool(&original, sheet, "B4", true, "original");
    assert_cell_bool(&reopened, sheet, "B4", true, "roundtrip");

    // String
    assert_cell_str(&original, sheet, "B6", "Test Case", "original");
    assert_cell_str(&reopened, sheet, "B6", "Test Case", "roundtrip");

    // Date serial (40423 = 9/2/2010)
    assert_cell_num(&original, sheet, "B3", 40423.0, "original");
    assert_cell_num(&reopened, sheet, "B3", 40423.0, "roundtrip");

    // Error: #DIV/0!
    assert_cell_error(&original, sheet, "B8", "#DIV/0!", "original");
    assert_cell_error(&reopened, sheet, "B8", "#DIV/0!", "roundtrip");
}

/// Errors sheet has multiple error types. Verify they all survive.
#[test]
fn dirty_roundtrip_preserves_error_types() {
    let src = closedxml_fixture("Examples/Misc/CellValues.xlsx");
    if skip_if_missing(&src) {
        return;
    }

    let (original, reopened, _tmp) = dirty_roundtrip(&src);
    let sheet = "Errors";

    let errors = [
        ("B3", "#REF!"),
        ("B4", "#VALUE!"),
        ("B5", "#DIV/0!"),
        ("B6", "#NAME?"),
        ("B7", "#N/A"),
        ("B8", "#NULL!"),
        ("B9", "#NUM!"),
    ];

    for (cell_ref, expected) in &errors {
        assert_cell_error(&original, sheet, cell_ref, expected, "original");
        assert_cell_error(&reopened, sheet, cell_ref, expected, "roundtrip");
    }
}

/// Errors sheet also has formula errors with formulas. Verify formulas survive.
#[test]
fn dirty_roundtrip_preserves_error_formulas() {
    let src = closedxml_fixture("Examples/Misc/CellValues.xlsx");
    if skip_if_missing(&src) {
        return;
    }

    let (original, reopened, _tmp) = dirty_roundtrip(&src);
    let sheet = "Errors";

    // These cells have formulas that produce errors
    let formulas = [("C3", "#REF!+1"), ("C5", "1/0")];

    for (cell_ref, expected) in &formulas {
        assert_cell_formula(&original, sheet, cell_ref, expected, "original");
        assert_cell_formula(&reopened, sheet, cell_ref, expected, "roundtrip");
    }
}

/// FormulasWithEvaluation.xlsx has formulas with cached values.
/// Verify formulas AND their cached values survive.
#[test]
fn dirty_roundtrip_preserves_formulas_and_cached_values() {
    let src = closedxml_fixture("Examples/Misc/FormulasWithEvaluation.xlsx");
    if skip_if_missing(&src) {
        return;
    }

    let (original, reopened, _tmp) = dirty_roundtrip(&src);
    let sheet = "Formulas";

    // Source data
    assert_cell_num(&original, sheet, "A2", 1.0, "original");
    assert_cell_num(&reopened, sheet, "A2", 1.0, "roundtrip");

    assert_cell_num(&original, sheet, "B2", 2.0, "original");
    assert_cell_num(&reopened, sheet, "B2", 2.0, "roundtrip");

    // Formula: =A2+$B$2
    assert_cell_formula(&original, sheet, "C2", "A2+$B$2", "original");
    assert_cell_formula(&reopened, sheet, "C2", "A2+$B$2", "roundtrip");

    // Cached value lives in cached_value(), not value(), for formula cells
    assert_cell_cached_num(&original, sheet, "C2", 3.0, "original cached");
    assert_cell_cached_num(&reopened, sheet, "C2", 3.0, "roundtrip cached");

    // String concatenation formula
    assert_cell_str(&original, sheet, "A4", "A", "original");
    assert_cell_str(&reopened, sheet, "A4", "A", "roundtrip");
}

/// BasicTable.xlsx has a table with mixed types.
/// Verify the table data survives dirty roundtrip.
#[test]
fn dirty_roundtrip_preserves_table_data() {
    let src = closedxml_fixture("Examples/Misc/BasicTable.xlsx");
    if skip_if_missing(&src) {
        return;
    }

    let (original, reopened, _tmp) = dirty_roundtrip(&src);
    let sheet = "Contacts";

    // Header
    assert_cell_str(&original, sheet, "B3", "FName", "original");
    assert_cell_str(&reopened, sheet, "B3", "FName", "roundtrip");

    // Data rows - strings
    assert_cell_str(&original, sheet, "B4", "John", "original");
    assert_cell_str(&reopened, sheet, "B4", "John", "roundtrip");

    assert_cell_str(&original, sheet, "C4", "Galt", "original");
    assert_cell_str(&reopened, sheet, "C4", "Galt", "roundtrip");

    assert_cell_str(&original, sheet, "B5", "Hank", "original");
    assert_cell_str(&reopened, sheet, "B5", "Hank", "roundtrip");

    // Bool column
    assert_cell_bool(&original, sheet, "D4", true, "original");
    assert_cell_bool(&reopened, sheet, "D4", true, "roundtrip");

    assert_cell_bool(&original, sheet, "D5", false, "original");
    assert_cell_bool(&reopened, sheet, "D5", false, "roundtrip");

    // Numeric column
    assert_cell_num(&original, sheet, "F4", 2000.0, "original");
    assert_cell_num(&reopened, sheet, "F4", 2000.0, "roundtrip");

    assert_cell_num(&original, sheet, "F5", 40000.0, "original");
    assert_cell_num(&reopened, sheet, "F5", 40000.0, "roundtrip");
}

/// ShowCase.xlsx has a table with a SUBTOTAL formula in the totals row.
/// Verify the formula survives.
#[test]
fn dirty_roundtrip_preserves_table_formula() {
    let src = closedxml_fixture("Examples/Misc/ShowCase.xlsx");
    if skip_if_missing(&src) {
        return;
    }

    let (original, reopened, _tmp) = dirty_roundtrip(&src);
    let sheet = "Contacts";

    // Table data still intact
    assert_cell_str(&original, sheet, "B4", "John", "original");
    assert_cell_str(&reopened, sheet, "B4", "John", "roundtrip");

    // Totals row label
    assert_cell_str(&original, sheet, "E7", "Average:", "original");
    assert_cell_str(&reopened, sheet, "E7", "Average:", "roundtrip");

    // SUBTOTAL formula in totals row
    assert_cell_formula(&original, sheet, "F7", "SUBTOTAL(101,[Income])", "original");
    assert_cell_formula(
        &reopened,
        sheet,
        "F7",
        "SUBTOTAL(101,[Income])",
        "roundtrip",
    );
}

/// Open a complex file, modify ALL sheets, verify cell content from every sheet.
#[test]
fn dirty_roundtrip_all_sheets_preserves_content() {
    let src = closedxml_fixture("Examples/Misc/CellValues.xlsx");
    if skip_if_missing(&src) {
        return;
    }

    let original = Workbook::open(&src).expect("open original");

    // Dirty ALL sheets
    let mut modified = Workbook::open(&src).expect("open for modification");
    let names: Vec<String> = modified
        .sheet_names()
        .into_iter()
        .map(String::from)
        .collect();
    for name in &names {
        let ws = modified.sheet_mut(name).expect("get sheet");
        ws.cell_mut("ZZ1")
            .expect("get ZZ1")
            .set_value(CellValue::String(format!("DIRTY_{name}")));
    }

    let tmp = tempfile::tempdir().expect("create tempdir");
    let output = tmp.path().join("output.xlsx");
    modified.save(&output).expect("save");

    let reopened = Workbook::open(&output).expect("reopen");

    // Sheet count preserved
    assert_eq!(original.sheet_names().len(), reopened.sheet_names().len());

    // Content on "Cell Values" sheet
    assert_cell_num(
        &reopened,
        "Cell Values",
        "B5",
        1234.567,
        "all-dirty roundtrip",
    );
    assert_cell_bool(&reopened, "Cell Values", "B4", true, "all-dirty roundtrip");
    assert_cell_str(
        &reopened,
        "Cell Values",
        "B6",
        "Test Case",
        "all-dirty roundtrip",
    );

    // Content on "Errors" sheet
    assert_cell_error(&reopened, "Errors", "B3", "#REF!", "all-dirty roundtrip");
    assert_cell_error(&reopened, "Errors", "B5", "#DIV/0!", "all-dirty roundtrip");

    // Verify our dirty markers are there
    for name in &names {
        let ws = reopened.sheet(name).expect("sheet exists");
        let cell = ws.cell("ZZ1").expect("dirty marker exists");
        match cell.value() {
            Some(CellValue::String(s)) => assert_eq!(s, &format!("DIRTY_{name}")),
            other => panic!("dirty marker for {name}: expected String, got {other:?}"),
        }
    }
}

// ---------------------------------------------------------------------------
// Bulk corpus: dirty roundtrip every ClosedXML example and compare ALL cells
// ---------------------------------------------------------------------------

/// Fingerprint: for each cell, store (value_debug, formula) so we can compare.
fn fingerprint_sheet(wb: &Workbook, sheet_name: &str) -> Vec<(String, String, String)> {
    let ws = match wb.sheet(sheet_name) {
        Some(ws) => ws,
        None => return vec![],
    };
    let mut cells = Vec::new();
    let cols: Vec<&str> = vec![
        "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P",
    ];
    for row in 1..=100 {
        for col in &cols {
            let ref_str = format!("{col}{row}");
            if let Some(cell) = ws.cell(&ref_str) {
                let val = format!("{:?}", cell.value());
                let formula = cell.formula().unwrap_or("").to_string();
                if val != "None" && val != "Some(Blank)" {
                    cells.push((ref_str, val, formula));
                }
            }
        }
    }
    cells
}

/// Walk all ClosedXML Example fixtures, dirty roundtrip each one,
/// compare cell fingerprints. Report every mismatch.
#[test]
fn bulk_closedxml_dirty_roundtrip_content_verification() {
    let examples_dir = reference_root().join("ClosedXML/ClosedXML.Tests/Resource/Examples");
    if !examples_dir.is_dir() {
        eprintln!(
            "skipping: ClosedXML Examples not found at {}",
            examples_dir.display()
        );
        return;
    }

    let mut files: Vec<PathBuf> = Vec::new();
    collect_xlsx_files(&examples_dir, &mut files);
    files.sort();

    if files.is_empty() {
        eprintln!("skipping: no xlsx files found in ClosedXML Examples");
        return;
    }

    let mut total = 0;
    let mut passed = 0;
    let mut open_failed = Vec::new();
    let mut save_failed = Vec::new();
    let mut content_mismatches = Vec::new();

    for file in &files {
        total += 1;
        let name = file.file_name().unwrap().to_string_lossy().to_string();

        // Open original
        let original = match Workbook::open(file) {
            Ok(wb) => wb,
            Err(e) => {
                open_failed.push(format!("{name}: {e}"));
                continue;
            }
        };

        // Dirty roundtrip
        let mut modified = match Workbook::open(file) {
            Ok(wb) => wb,
            Err(e) => {
                open_failed.push(format!("{name}: {e}"));
                continue;
            }
        };
        let sheet_name = modified.sheet_names()[0].to_string();
        if let Ok(cell) = modified.sheet_mut(&sheet_name).unwrap().cell_mut("ZZ1") {
            cell.set_value(CellValue::String("DIRTY".into()));
        }

        let tmp = tempfile::tempdir().expect("tempdir");
        let output = tmp.path().join("output.xlsx");
        if let Err(e) = modified.save(&output) {
            save_failed.push(format!("{name}: {e}"));
            continue;
        }

        let reopened = match Workbook::open(&output) {
            Ok(wb) => wb,
            Err(e) => {
                save_failed.push(format!("{name}: reopen failed: {e}"));
                continue;
            }
        };

        // Compare fingerprints for ALL sheets
        let mut file_mismatches = Vec::new();
        for sname in original
            .sheet_names()
            .into_iter()
            .map(String::from)
            .collect::<Vec<_>>()
        {
            let orig_fp = fingerprint_sheet(&original, &sname);
            let reopen_fp = fingerprint_sheet(&reopened, &sname);

            // Find cells in original but different/missing in reopened
            for (cell_ref, orig_val, orig_formula) in &orig_fp {
                let reopen_entry = reopen_fp.iter().find(|(r, _, _)| r == cell_ref);
                match reopen_entry {
                    Some((_, reopen_val, reopen_formula)) => {
                        if orig_val != reopen_val || orig_formula != reopen_formula {
                            file_mismatches.push(format!(
                                "  [{sname}!{cell_ref}] orig=({orig_val}, f={orig_formula}) reopen=({reopen_val}, f={reopen_formula})"
                            ));
                        }
                    }
                    None => {
                        file_mismatches.push(format!(
                            "  [{sname}!{cell_ref}] LOST: was ({orig_val}, f={orig_formula})"
                        ));
                    }
                }
            }

            // Find cells in reopened but not in original (spurious additions)
            for (cell_ref, reopen_val, _) in &reopen_fp {
                if cell_ref == "ZZ1" {
                    continue;
                } // our dirty marker
                if !orig_fp.iter().any(|(r, _, _)| r == cell_ref) {
                    file_mismatches.push(format!("  [{sname}!{cell_ref}] ADDED: ({reopen_val})"));
                }
            }
        }

        if file_mismatches.is_empty() {
            passed += 1;
        } else {
            content_mismatches.push(format!("{name}:\n{}", file_mismatches.join("\n")));
        }
    }

    // Report
    eprintln!("\n=== BULK CONTENT ROUNDTRIP RESULTS ===");
    eprintln!("Total files: {total}");
    eprintln!("Passed (content identical): {passed}");
    eprintln!("Open failures: {}", open_failed.len());
    eprintln!("Save failures: {}", save_failed.len());
    eprintln!("Content mismatches: {}", content_mismatches.len());

    if !open_failed.is_empty() {
        eprintln!("\n--- OPEN FAILURES ---");
        for f in &open_failed {
            eprintln!("  {f}");
        }
    }
    if !save_failed.is_empty() {
        eprintln!("\n--- SAVE FAILURES ---");
        for f in &save_failed {
            eprintln!("  {f}");
        }
    }
    if !content_mismatches.is_empty() {
        eprintln!("\n--- CONTENT MISMATCHES ---");
        for m in &content_mismatches {
            eprintln!("{m}");
        }
    }

    // Hard fail on open/save failures — those are bugs
    assert!(
        open_failed.is_empty(),
        "{} files failed to open:\n{}",
        open_failed.len(),
        open_failed.join("\n")
    );
    assert!(
        save_failed.is_empty(),
        "{} files failed to save:\n{}",
        save_failed.len(),
        save_failed.join("\n")
    );

    // Report content mismatches but don't hard-fail — we want to see all of them
    if !content_mismatches.is_empty() {
        panic!(
            "{} files had content mismatches after dirty roundtrip (see above)",
            content_mismatches.len()
        );
    }
}

fn collect_xlsx_files(dir: &Path, out: &mut Vec<PathBuf>) {
    if let Ok(entries) = std::fs::read_dir(dir) {
        for entry in entries.flatten() {
            let path = entry.path();
            if path.is_dir() {
                collect_xlsx_files(&path, out);
            } else if path.extension().is_some_and(|e| e == "xlsx") {
                out.push(path);
            }
        }
    }
}

/// Open Spreadsheet.xlsx from OpenXML SDK — a file we didn't create.
/// Verify it roundtrips without content loss.
#[test]
fn dirty_roundtrip_openxml_sdk_spreadsheet() {
    let src = openxml_fixture("Spreadsheet.xlsx");
    if skip_if_missing(&src) {
        return;
    }

    let original = Workbook::open(&src).expect("open original");
    let orig_names = original
        .sheet_names()
        .into_iter()
        .map(String::from)
        .collect::<Vec<_>>();

    let (_, reopened, _tmp) = dirty_roundtrip(&src);
    let reopen_names = reopened
        .sheet_names()
        .into_iter()
        .map(String::from)
        .collect::<Vec<_>>();

    assert_eq!(orig_names, reopen_names, "sheet names changed");

    // Compare cell values on first sheet
    let orig_ws = original.sheet(&orig_names[0]).expect("original sheet");
    let reopen_ws = reopened.sheet(&orig_names[0]).expect("reopened sheet");

    // Check first 20 rows x 10 cols for value equality
    let cols = ["A", "B", "C", "D", "E", "F", "G", "H", "I", "J"];
    for row in 1..=20 {
        for col in &cols {
            let ref_str = format!("{col}{row}");
            let orig_cell = orig_ws.cell(&ref_str);
            let reopen_cell = reopen_ws.cell(&ref_str);

            match (orig_cell, reopen_cell) {
                (Some(o), Some(r)) => {
                    assert_eq!(
                        o.value(),
                        r.value(),
                        "cell {ref_str} value mismatch: orig={:?} reopened={:?}",
                        o.value(),
                        r.value()
                    );
                    assert_eq!(o.formula(), r.formula(), "cell {ref_str} formula mismatch");
                }
                (None, None) => {} // both empty, fine
                (orig, reopen) => {
                    // One exists, one doesn't — check if the existing one is blank
                    let orig_val = orig.and_then(|c| c.value().cloned());
                    let reopen_val = reopen.and_then(|c| c.value().cloned());
                    let orig_blank = orig_val.is_none() || orig_val == Some(CellValue::Blank);
                    let reopen_blank = reopen_val.is_none() || reopen_val == Some(CellValue::Blank);
                    if !orig_blank || !reopen_blank {
                        panic!(
                            "cell {ref_str}: orig={:?} reopened={:?}",
                            orig_val, reopen_val
                        );
                    }
                }
            }
        }
    }
}
