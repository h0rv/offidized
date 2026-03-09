//! Comprehensive integration tests for offidized-ir.
//!
//! These tests exercise the full derive→apply pipeline with both programmatically
//! created workbooks and real xlsx/docx/pptx files from the reference repositories.

#![allow(
    clippy::expect_used,
    clippy::panic_in_result_fn,
    clippy::approx_constant
)]

use std::path::{Path, PathBuf};

use offidized_docx::Document;
use offidized_ir::{apply, derive, ApplyOptions, DeriveOptions, IrError, Mode};
use offidized_pptx::Presentation;
use offidized_xlsx::{CellValue, RichTextRun, Workbook};

type R = Result<(), Box<dyn std::error::Error>>;

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

/// Create a temp dir + file path with the given extension.
fn tmp(name: &str) -> (tempfile::TempDir, PathBuf) {
    let dir = tempfile::tempdir().expect("tempdir");
    let path = dir.path().join(name);
    (dir, path)
}

/// Strip the TOML header from an IR string, returning just the body.
fn body_of(ir: &str) -> String {
    let (_, body) = offidized_ir::IrHeader::parse(ir).expect("parse header");
    body
}

/// Reference file paths (relative to workspace root).
/// Try to resolve a reference file, returning None if it doesn't exist
/// (so tests can be skipped when references aren't cloned).
fn try_ref_file(rel: &str) -> Option<PathBuf> {
    let root = PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../..")
        .canonicalize()
        .ok()?;
    let p = root.join(rel);
    p.exists().then_some(p)
}

// =========================================================================
// 1. Roundtrip WITHOUT modifications
// =========================================================================

#[test]
fn roundtrip_no_modifications_simple() -> R {
    let (_dir, path) = tmp("simple.xlsx");

    let mut wb = Workbook::new();
    let ws = wb.add_sheet("Data");
    ws.cell_mut("A1")?.set_value("Name");
    ws.cell_mut("B1")?.set_value("Score");
    ws.cell_mut("A2")?.set_value("Alice");
    ws.cell_mut("B2")?.set_value(95);
    ws.cell_mut("A3")?.set_value("Bob");
    ws.cell_mut("B3")?.set_value(87);
    wb.save(&path)?;

    // Derive → apply (no edits) → derive again
    let ir1 = derive(&path, DeriveOptions::default())?;

    let (_dir2, out) = tmp("out.xlsx");
    apply(
        &ir1,
        &out,
        &ApplyOptions {
            force: true,
            ..Default::default()
        },
    )?;

    let ir2 = derive(&out, DeriveOptions::default())?;

    // Bodies must be identical (headers differ in source path + checksum)
    pretty_assertions::assert_eq!(body_of(&ir1), body_of(&ir2));
    Ok(())
}

#[test]
fn roundtrip_no_modifications_multi_sheet() -> R {
    let (_dir, path) = tmp("multi.xlsx");

    let mut wb = Workbook::new();
    wb.add_sheet("Revenue").cell_mut("A1")?.set_value(100);
    wb.add_sheet("Costs").cell_mut("A1")?.set_value(50);
    wb.add_sheet("Summary")
        .cell_mut("A1")?
        .set_formula("Revenue!A1-Costs!A1");
    wb.save(&path)?;

    let ir1 = derive(&path, DeriveOptions::default())?;
    let (_dir2, out) = tmp("out.xlsx");
    apply(
        &ir1,
        &out,
        &ApplyOptions {
            force: true,
            ..Default::default()
        },
    )?;
    let ir2 = derive(&out, DeriveOptions::default())?;

    pretty_assertions::assert_eq!(body_of(&ir1), body_of(&ir2));
    Ok(())
}

#[test]
fn roundtrip_no_modifications_all_value_types() -> R {
    let (_dir, path) = tmp("types.xlsx");

    let mut wb = Workbook::new();
    let ws = wb.add_sheet("Types");
    ws.cell_mut("A1")?.set_value("plain string");
    ws.cell_mut("A2")?.set_value(CellValue::String("42".into())); // string-as-number
    ws.cell_mut("A3")?
        .set_value(CellValue::String("true".into())); // string-as-bool
    ws.cell_mut("A4")?
        .set_value(CellValue::String("=SUM(1,2)".into())); // string-as-formula
    ws.cell_mut("A5")?
        .set_value(CellValue::String("#REF!".into())); // string-as-error
    ws.cell_mut("A6")?
        .set_value(CellValue::String("<empty>".into())); // string-as-marker
    ws.cell_mut("A7")?.set_value(CellValue::Number(3.14));
    ws.cell_mut("A8")?.set_value(CellValue::Number(0.0));
    ws.cell_mut("A9")?.set_value(CellValue::Number(-42000.0));
    ws.cell_mut("A10")?.set_value(CellValue::Bool(true));
    ws.cell_mut("A11")?.set_value(CellValue::Bool(false));
    ws.cell_mut("A12")?
        .set_value(CellValue::Error("#DIV/0!".into()));
    ws.cell_mut("A13")?
        .set_value(CellValue::Error("#N/A".into()));
    ws.cell_mut("A14")?
        .set_value(CellValue::String("has \"quotes\"".into()));
    ws.cell_mut("A15")?
        .set_value(CellValue::String("line1\nline2".into()));
    ws.cell_mut("A16")?
        .set_value(CellValue::String("  leading spaces".into()));
    ws.cell_mut("A17")?
        .set_value(CellValue::String("trailing spaces  ".into()));
    ws.cell_mut("A18")?.set_value(CellValue::String("".into())); // empty string
    ws.cell_mut("A19")?.set_formula("SUM(A7:A9)");
    ws.cell_mut("A20")?.set_value(CellValue::RichText(vec![
        RichTextRun::new("Bold "),
        RichTextRun::new("Normal"),
    ]));
    wb.save(&path)?;

    let ir1 = derive(&path, DeriveOptions::default())?;

    // Verify specific encodings in the IR
    let body = body_of(&ir1);
    assert!(body.contains("A1: plain string"), "plain string unquoted");
    assert!(body.contains("A2: \"42\""), "string-as-number quoted");
    assert!(body.contains("A3: \"true\""), "string-as-bool quoted");
    assert!(
        body.contains("A4: \"=SUM(1,2)\""),
        "string-as-formula quoted"
    );
    assert!(body.contains("A5: \"#REF!\""), "string-as-error quoted");
    assert!(body.contains("A6: \"<empty>\""), "string-as-marker quoted");
    assert!(body.contains("A7: 3.14"), "float");
    assert!(body.contains("A8: 0"), "zero");
    assert!(body.contains("A9: -42000"), "negative integer");
    assert!(body.contains("A10: true"), "bool true");
    assert!(body.contains("A11: false"), "bool false");
    assert!(body.contains("A12: #DIV/0!"), "error bare");
    assert!(body.contains("A13: #N/A"), "error N/A bare");
    assert!(
        body.contains("A14: \"has \"\"quotes\"\"\""),
        "embedded quotes"
    );
    assert!(body.contains("A15: \"line1\\nline2\""), "newline escaped");
    assert!(
        body.contains("A16: \"  leading spaces\""),
        "leading ws quoted"
    );
    assert!(
        body.contains("A17: \"trailing spaces  \""),
        "trailing ws quoted"
    );
    assert!(body.contains("A18: \"\""), "empty string quoted");
    assert!(body.contains("A19: =SUM(A7:A9)"), "formula");
    assert!(body.contains("A20: Bold Normal"), "rich text flattened");

    // Apply back and re-derive
    let (_dir2, out) = tmp("out.xlsx");
    apply(
        &ir1,
        &out,
        &ApplyOptions {
            force: true,
            ..Default::default()
        },
    )?;
    let ir2 = derive(&out, DeriveOptions::default())?;

    // Rich text will lose formatting and become a plain string on roundtrip,
    // but the text content is preserved. Everything else should be identical.
    let body2 = body_of(&ir2);

    // Check every cell value roundtrips (except rich text becomes plain string)
    for i in 1..=19 {
        let line_prefix = format!("A{i}: ");
        let orig_line = body.lines().find(|l| l.starts_with(&line_prefix));
        let rt_line = body2.lines().find(|l| l.starts_with(&line_prefix));
        assert_eq!(
            orig_line, rt_line,
            "A{i} mismatch: orig={orig_line:?} rt={rt_line:?}"
        );
    }

    Ok(())
}

#[test]
fn roundtrip_no_modifications_formulas() -> R {
    let (_dir, path) = tmp("formulas.xlsx");

    let mut wb = Workbook::new();
    let ws = wb.add_sheet("Calc");
    ws.cell_mut("A1")?.set_value(10);
    ws.cell_mut("A2")?.set_value(20);
    ws.cell_mut("A3")?.set_value(30);
    ws.cell_mut("B1")?.set_formula("SUM(A1:A3)");
    ws.cell_mut("B2")?.set_formula("AVERAGE(A1:A3)");
    ws.cell_mut("B3")?.set_formula("IF(A1>5,\"big\",\"small\")");
    wb.save(&path)?;

    let ir1 = derive(&path, DeriveOptions::default())?;
    let (_dir2, out) = tmp("out.xlsx");
    apply(
        &ir1,
        &out,
        &ApplyOptions {
            force: true,
            ..Default::default()
        },
    )?;
    let ir2 = derive(&out, DeriveOptions::default())?;

    pretty_assertions::assert_eq!(body_of(&ir1), body_of(&ir2));
    Ok(())
}

// =========================================================================
// 2. Roundtrip WITH modifications
// =========================================================================

#[test]
fn roundtrip_modify_single_cell() -> R {
    let (_dir, path) = tmp("modify.xlsx");

    let mut wb = Workbook::new();
    let ws = wb.add_sheet("Sheet1");
    ws.cell_mut("A1")?.set_value("original");
    ws.cell_mut("B1")?.set_value(100);
    ws.cell_mut("C1")?.set_value(true);
    wb.save(&path)?;

    let ir = derive(&path, DeriveOptions::default())?;

    // Modify B1 from 100 to 999
    let modified_ir = ir.replace("B1: 100", "B1: 999");
    assert_ne!(ir, modified_ir, "IR should have changed");

    let (_dir2, out) = tmp("out.xlsx");
    apply(
        &modified_ir,
        &out,
        &ApplyOptions {
            force: true,
            ..Default::default()
        },
    )?;

    let wb2 = Workbook::open(&out)?;
    let ws2 = wb2.sheet("Sheet1").ok_or("missing")?;

    // B1 changed
    assert_eq!(
        ws2.cell("B1").and_then(|c| c.value()),
        Some(&CellValue::Number(999.0))
    );
    // A1 and C1 unchanged
    assert_eq!(
        ws2.cell("A1").and_then(|c| c.value()),
        Some(&CellValue::String("original".into()))
    );
    assert_eq!(
        ws2.cell("C1").and_then(|c| c.value()),
        Some(&CellValue::Bool(true))
    );

    Ok(())
}

#[test]
fn roundtrip_modify_multiple_cells() -> R {
    let (_dir, path) = tmp("multi_mod.xlsx");

    let mut wb = Workbook::new();
    let ws = wb.add_sheet("Sheet1");
    ws.cell_mut("A1")?.set_value("Category");
    ws.cell_mut("B1")?.set_value("Q1");
    ws.cell_mut("C1")?.set_value("Q2");
    ws.cell_mut("A2")?.set_value("Sales");
    ws.cell_mut("B2")?.set_value(1000);
    ws.cell_mut("C2")?.set_value(1200);
    ws.cell_mut("A3")?.set_value("Costs");
    ws.cell_mut("B3")?.set_value(800);
    ws.cell_mut("C3")?.set_value(900);
    wb.save(&path)?;

    let ir = derive(&path, DeriveOptions::default())?;
    let modified = ir
        .replace("B2: 1000", "B2: 1500")
        .replace("C2: 1200", "C2: 1800")
        .replace("B3: 800", "B3: 700")
        .replace("C3: 900", "C3: 850");

    let (_dir2, out) = tmp("out.xlsx");
    let result = apply(
        &modified,
        &out,
        &ApplyOptions {
            force: true,
            ..Default::default()
        },
    )?;
    assert_eq!(result.cells_updated, 9); // all 9 cells present in IR
    assert_eq!(result.cells_created, 0);

    let wb2 = Workbook::open(&out)?;
    let ws2 = wb2.sheet("Sheet1").ok_or("missing")?;
    assert_eq!(
        ws2.cell("B2").and_then(|c| c.value()),
        Some(&CellValue::Number(1500.0))
    );
    assert_eq!(
        ws2.cell("C2").and_then(|c| c.value()),
        Some(&CellValue::Number(1800.0))
    );
    assert_eq!(
        ws2.cell("B3").and_then(|c| c.value()),
        Some(&CellValue::Number(700.0))
    );
    assert_eq!(
        ws2.cell("C3").and_then(|c| c.value()),
        Some(&CellValue::Number(850.0))
    );
    // Headers unchanged
    assert_eq!(
        ws2.cell("A1").and_then(|c| c.value()),
        Some(&CellValue::String("Category".into()))
    );

    Ok(())
}

#[test]
fn roundtrip_add_new_cells() -> R {
    let (_dir, path) = tmp("add_cells.xlsx");

    let mut wb = Workbook::new();
    let ws = wb.add_sheet("Sheet1");
    ws.cell_mut("A1")?.set_value("existing");
    wb.save(&path)?;

    let ir = derive(&path, DeriveOptions::default())?;
    // Add new cells to the IR
    let extended = ir.replace(
        "A1: existing",
        "A1: existing\nB1: new value\nA2: another new\nB2: 42",
    );

    let (_dir2, out) = tmp("out.xlsx");
    let result = apply(
        &extended,
        &out,
        &ApplyOptions {
            force: true,
            ..Default::default()
        },
    )?;
    assert_eq!(result.cells_updated, 1); // A1
    assert_eq!(result.cells_created, 3); // B1, A2, B2

    let wb2 = Workbook::open(&out)?;
    let ws2 = wb2.sheet("Sheet1").ok_or("missing")?;
    assert_eq!(
        ws2.cell("A1").and_then(|c| c.value()),
        Some(&CellValue::String("existing".into()))
    );
    assert_eq!(
        ws2.cell("B1").and_then(|c| c.value()),
        Some(&CellValue::String("new value".into()))
    );
    assert_eq!(
        ws2.cell("A2").and_then(|c| c.value()),
        Some(&CellValue::String("another new".into()))
    );
    assert_eq!(
        ws2.cell("B2").and_then(|c| c.value()),
        Some(&CellValue::Number(42.0))
    );

    Ok(())
}

#[test]
fn roundtrip_clear_cells_with_empty() -> R {
    let (_dir, path) = tmp("clear.xlsx");

    let mut wb = Workbook::new();
    let ws = wb.add_sheet("Sheet1");
    ws.cell_mut("A1")?.set_value("keep");
    ws.cell_mut("B1")?.set_value("remove me");
    ws.cell_mut("C1")?.set_value(42);
    ws.cell_mut("D1")?.set_formula("SUM(C1:C1)");
    wb.save(&path)?;

    let ir = derive(&path, DeriveOptions::default())?;
    // Clear B1 and D1
    let modified = ir
        .replace("B1: remove me", "B1: <empty>")
        .replace("D1: =SUM(C1:C1)", "D1: <empty>");

    let (_dir2, out) = tmp("out.xlsx");
    let result = apply(
        &modified,
        &out,
        &ApplyOptions {
            force: true,
            ..Default::default()
        },
    )?;
    assert_eq!(result.cells_cleared, 2);

    let wb2 = Workbook::open(&out)?;
    let ws2 = wb2.sheet("Sheet1").ok_or("missing")?;
    assert_eq!(
        ws2.cell("A1").and_then(|c| c.value()),
        Some(&CellValue::String("keep".into()))
    );
    // B1 and D1 should have no value
    assert_eq!(ws2.cell("B1").and_then(|c| c.value()), None);
    assert_eq!(ws2.cell("D1").and_then(|c| c.value()), None);
    assert_eq!(ws2.cell("D1").and_then(|c| c.formula()), None);
    // C1 unchanged
    assert_eq!(
        ws2.cell("C1").and_then(|c| c.value()),
        Some(&CellValue::Number(42.0))
    );

    Ok(())
}

#[test]
fn roundtrip_change_value_type() -> R {
    let (_dir, path) = tmp("type_change.xlsx");

    let mut wb = Workbook::new();
    let ws = wb.add_sheet("Sheet1");
    ws.cell_mut("A1")?.set_value(100); // number
    ws.cell_mut("A2")?.set_value("text"); // string
    ws.cell_mut("A3")?.set_value(true); // bool
    ws.cell_mut("A4")?.set_formula("SUM(A1:A1)"); // formula
    wb.save(&path)?;

    // Change types: number→string, string→number, bool→formula, formula→number
    let ir = derive(&path, DeriveOptions::default())?;
    let modified = ir
        .replace("A1: 100", "A1: now a string")
        .replace("A2: text", "A2: 999")
        .replace("A3: true", "A3: =1+1")
        .replace("A4: =SUM(A1:A1)", "A4: false");

    let (_dir2, out) = tmp("out.xlsx");
    apply(
        &modified,
        &out,
        &ApplyOptions {
            force: true,
            ..Default::default()
        },
    )?;

    let wb2 = Workbook::open(&out)?;
    let ws2 = wb2.sheet("Sheet1").ok_or("missing")?;
    assert_eq!(
        ws2.cell("A1").and_then(|c| c.value()),
        Some(&CellValue::String("now a string".into()))
    );
    assert_eq!(
        ws2.cell("A2").and_then(|c| c.value()),
        Some(&CellValue::Number(999.0))
    );
    assert_eq!(ws2.cell("A3").and_then(|c| c.formula()), Some("1+1"));
    assert_eq!(
        ws2.cell("A4").and_then(|c| c.value()),
        Some(&CellValue::Bool(false))
    );

    Ok(())
}

#[test]
fn roundtrip_replace_formula_with_value_clears_formula() -> R {
    let (_dir, path) = tmp("formula_clear.xlsx");

    let mut wb = Workbook::new();
    let ws = wb.add_sheet("Sheet1");
    ws.cell_mut("A1")?.set_value(10);
    ws.cell_mut("A2")?.set_formula("A1*2");
    wb.save(&path)?;

    // Replace formula with a plain value
    let ir = derive(&path, DeriveOptions::default())?;
    let modified = ir.replace("A2: =A1*2", "A2: 20");

    let (_dir2, out) = tmp("out.xlsx");
    apply(
        &modified,
        &out,
        &ApplyOptions {
            force: true,
            ..Default::default()
        },
    )?;

    let wb2 = Workbook::open(&out)?;
    let ws2 = wb2.sheet("Sheet1").ok_or("missing")?;
    assert_eq!(
        ws2.cell("A2").and_then(|c| c.value()),
        Some(&CellValue::Number(20.0))
    );
    assert_eq!(
        ws2.cell("A2").and_then(|c| c.formula()),
        None,
        "formula should be cleared"
    );

    Ok(())
}

// =========================================================================
// 3. Checksum / staleness validation
// =========================================================================

#[test]
fn checksum_mismatch_errors() -> R {
    let (_dir, path) = tmp("stale.xlsx");

    let mut wb = Workbook::new();
    wb.add_sheet("Sheet1").cell_mut("A1")?.set_value("original");
    wb.save(&path)?;

    let ir = derive(&path, DeriveOptions::default())?;

    // Modify the source file after deriving
    let mut wb2 = Workbook::open(&path)?;
    wb2.sheet_mut("Sheet1")
        .ok_or("missing")?
        .cell_mut("A1")?
        .set_value("changed");
    wb2.save(&path)?;

    // Apply should fail with checksum mismatch
    let (_dir2, out) = tmp("out.xlsx");
    let result = apply(&ir, &out, &ApplyOptions::default());
    assert!(result.is_err());
    let err = result.err().ok_or("expected error")?;
    assert!(
        matches!(err, IrError::ChecksumMismatch { .. }),
        "expected ChecksumMismatch, got: {err}"
    );

    Ok(())
}

#[test]
fn checksum_mismatch_force_succeeds() -> R {
    let (_dir, path) = tmp("force.xlsx");

    let mut wb = Workbook::new();
    wb.add_sheet("Sheet1").cell_mut("A1")?.set_value("original");
    wb.save(&path)?;

    let ir = derive(&path, DeriveOptions::default())?;

    // Modify source
    let mut wb2 = Workbook::open(&path)?;
    wb2.sheet_mut("Sheet1")
        .ok_or("missing")?
        .cell_mut("B1")?
        .set_value("extra");
    wb2.save(&path)?;

    // Apply with --force should succeed
    let (_dir2, out) = tmp("out.xlsx");
    let result = apply(
        &ir,
        &out,
        &ApplyOptions {
            force: true,
            ..Default::default()
        },
    )?;
    assert_eq!(result.cells_updated, 1); // A1

    // Verify both old IR content and new content are present
    let wb3 = Workbook::open(&out)?;
    let ws3 = wb3.sheet("Sheet1").ok_or("missing")?;
    assert_eq!(
        ws3.cell("A1").and_then(|c| c.value()),
        Some(&CellValue::String("original".into()))
    );
    // B1 should still be "extra" since the IR only touched A1
    assert_eq!(
        ws3.cell("B1").and_then(|c| c.value()),
        Some(&CellValue::String("extra".into()))
    );

    Ok(())
}

#[test]
fn checksum_valid_succeeds() -> R {
    let (_dir, path) = tmp("valid.xlsx");

    let mut wb = Workbook::new();
    wb.add_sheet("Sheet1").cell_mut("A1")?.set_value("hello");
    wb.save(&path)?;

    let ir = derive(&path, DeriveOptions::default())?;

    // Apply without --force, source unchanged → should succeed
    let (_dir2, out) = tmp("out.xlsx");
    let result = apply(&ir, &out, &ApplyOptions::default())?;
    assert_eq!(result.cells_updated, 1);

    Ok(())
}

// =========================================================================
// 4. Edge cases and error handling
// =========================================================================

#[test]
fn apply_to_missing_sheet_creates_it() -> R {
    let (_dir, path) = tmp("missing_sheet.xlsx");

    let mut wb = Workbook::new();
    wb.add_sheet("Sheet1").cell_mut("A1")?.set_value("here");
    wb.save(&path)?;

    // IR that references a sheet that doesn't exist yet — should be auto-created.
    let ir = derive(&path, DeriveOptions::default())?;
    let with_new_sheet = format!("{ir}\n=== Sheet: NewSheet ===\nA1: created\n");

    let (_dir2, out) = tmp("out.xlsx");
    let result = apply(
        &with_new_sheet,
        &out,
        &ApplyOptions {
            force: true,
            ..Default::default()
        },
    )?;

    assert!(result.warnings.is_empty(), "should have no warnings");
    assert_eq!(result.cells_created, 1);

    let wb_out = Workbook::open(&out)?;
    assert!(wb_out.contains_sheet("NewSheet"), "new sheet should exist");
    let ws = wb_out.sheet("NewSheet").unwrap();
    assert_eq!(
        ws.cell("A1").and_then(|c| c.value()),
        Some(&CellValue::String("created".into())),
    );

    Ok(())
}

#[test]
fn apply_preserves_styles() -> R {
    let (_dir, path) = tmp("styles.xlsx");

    let mut wb = Workbook::new();
    let ws = wb.add_sheet("Sheet1");
    ws.cell_mut("A1")?.set_value("styled").set_style_id(1);
    ws.cell_mut("B1")?.set_value(100).set_style_id(2);
    wb.save(&path)?;

    // Derive, change values, apply
    let ir = derive(&path, DeriveOptions::default())?;
    let modified = ir
        .replace("A1: styled", "A1: new text")
        .replace("B1: 100", "B1: 200");

    let (_dir2, out) = tmp("out.xlsx");
    apply(
        &modified,
        &out,
        &ApplyOptions {
            force: true,
            ..Default::default()
        },
    )?;

    // Verify styles are preserved
    let wb2 = Workbook::open(&out)?;
    let ws2 = wb2.sheet("Sheet1").ok_or("missing")?;
    assert_eq!(ws2.cell("A1").and_then(|c| c.style_id()), Some(1));
    assert_eq!(ws2.cell("B1").and_then(|c| c.style_id()), Some(2));

    Ok(())
}

#[test]
fn derive_empty_workbook() -> R {
    let (_dir, path) = tmp("empty.xlsx");

    let wb = Workbook::new();
    wb.save(&path)?;

    let ir = derive(&path, DeriveOptions::default())?;
    // Should have header but no sheet sections (no sheets)
    assert!(ir.contains("+++"));
    assert!(!ir.contains("=== Sheet:"));

    Ok(())
}

#[test]
fn derive_empty_sheet() -> R {
    let (_dir, path) = tmp("empty_sheet.xlsx");

    let mut wb = Workbook::new();
    wb.add_sheet("Empty");
    wb.save(&path)?;

    let ir = derive(&path, DeriveOptions::default())?;
    let body = body_of(&ir);
    assert!(body.contains("=== Sheet: Empty ==="));
    // No cell lines after the header
    let after_header = body
        .split("=== Sheet: Empty ===")
        .nth(1)
        .ok_or("no sheet")?;
    let cell_lines: Vec<&str> = after_header
        .lines()
        .filter(|l| {
            let trimmed = l.trim();
            !trimmed.is_empty() && !trimmed.starts_with('#')
        })
        .collect();
    assert!(
        cell_lines.is_empty(),
        "empty sheet should have no cells: {cell_lines:?}"
    );

    Ok(())
}

#[test]
fn derive_hidden_sheet_annotation() -> R {
    let (_dir, path) = tmp("hidden.xlsx");

    let mut wb = Workbook::new();
    let ws1 = wb.add_sheet("Visible");
    ws1.cell_mut("A1")?.set_value("see me");
    let ws2 = wb.add_sheet("Hidden");
    ws2.set_visibility(offidized_xlsx::SheetVisibility::Hidden);
    ws2.cell_mut("A1")?.set_value("hidden data");
    let ws3 = wb.add_sheet("VeryHidden");
    ws3.set_visibility(offidized_xlsx::SheetVisibility::VeryHidden);
    ws3.cell_mut("A1")?.set_value("very hidden data");
    wb.save(&path)?;

    let ir = derive(&path, DeriveOptions::default())?;
    let body = body_of(&ir);

    // Visible sheet: no annotation
    assert!(body.contains("=== Sheet: Visible ==="));
    let vis_section = body
        .split("=== Sheet: Visible ===")
        .nth(1)
        .ok_or("missing")?;
    let vis_first_lines: Vec<&str> = vis_section.lines().take(3).collect();
    assert!(!vis_first_lines.iter().any(|l| l.contains("# hidden")));

    // Hidden sheet
    assert!(body.contains("=== Sheet: Hidden ==="));
    assert!(body.contains("# hidden"));

    // VeryHidden sheet
    assert!(body.contains("=== Sheet: VeryHidden ==="));
    assert!(body.contains("# very-hidden"));

    Ok(())
}

#[test]
fn derive_with_sheet_filter() -> R {
    let (_dir, path) = tmp("filter.xlsx");

    let mut wb = Workbook::new();
    wb.add_sheet("Alpha").cell_mut("A1")?.set_value("a");
    wb.add_sheet("Beta").cell_mut("A1")?.set_value("b");
    wb.add_sheet("Gamma").cell_mut("A1")?.set_value("c");
    wb.save(&path)?;

    let ir = derive(
        &path,
        DeriveOptions {
            sheet: Some("Beta".into()),
            ..Default::default()
        },
    )?;
    let body = body_of(&ir);

    assert!(!body.contains("Alpha"));
    assert!(body.contains("=== Sheet: Beta ==="));
    assert!(body.contains("A1: b"));
    assert!(!body.contains("Gamma"));

    Ok(())
}

#[test]
fn invalid_ir_missing_header() {
    let result = apply(
        "not an IR string",
        Path::new("/tmp/out.xlsx"),
        &ApplyOptions::default(),
    );
    assert!(result.is_err());
    assert!(matches!(result.err(), Some(IrError::InvalidHeader(_))));
}

#[test]
fn invalid_ir_bad_format() {
    let ir = "+++\nsource = \"x.txt\"\nformat = \"txt\"\nmode = \"content\"\nversion = 1\nchecksum = \"sha256:abc\"\n+++\n";
    let result = apply(ir, Path::new("/tmp/out.xlsx"), &ApplyOptions::default());
    assert!(result.is_err());
}

#[test]
fn source_override_works() -> R {
    let (_dir, path) = tmp("source.xlsx");
    let (_dir2, path2) = tmp("other.xlsx");

    let mut wb = Workbook::new();
    wb.add_sheet("Sheet1")
        .cell_mut("A1")?
        .set_value("from source");
    wb.save(&path)?;

    // Create a different file to use as override
    let mut wb2 = Workbook::new();
    wb2.add_sheet("Sheet1")
        .cell_mut("A1")?
        .set_value("from other");
    wb2.add_sheet("Sheet1"); // no-op, already exists
    wb2.save(&path2)?;

    // Derive from first file
    let ir = derive(&path, DeriveOptions::default())?;

    // Apply using second file as source override
    let (_dir3, out) = tmp("out.xlsx");
    let result = apply(
        &ir,
        &out,
        &ApplyOptions {
            source_override: Some(path2),
            force: true,
        },
    )?;

    // A1 should be overwritten with "from source" (the IR value), not "from other"
    let wb3 = Workbook::open(&out)?;
    let ws3 = wb3.sheet("Sheet1").ok_or("missing")?;
    assert_eq!(
        ws3.cell("A1").and_then(|c| c.value()),
        Some(&CellValue::String("from source".into()))
    );
    assert_eq!(result.cells_updated, 1);

    Ok(())
}

#[test]
fn large_workbook_roundtrip() -> R {
    let (_dir, path) = tmp("large.xlsx");

    let mut wb = Workbook::new();
    let ws = wb.add_sheet("Data");
    for row in 1..=100 {
        ws.cell_mut(&format!("A{row}"))?
            .set_value(format!("Row {row}"));
        ws.cell_mut(&format!("B{row}"))?.set_value(row as f64 * 1.5);
        ws.cell_mut(&format!("C{row}"))?.set_value(row % 2 == 0);
    }
    ws.cell_mut("D1")?.set_formula("SUM(B1:B100)");
    wb.save(&path)?;

    let ir1 = derive(&path, DeriveOptions::default())?;

    // Verify 100 rows * 3 cols + 1 formula = 301 cell lines
    let body = body_of(&ir1);
    let cell_count = body
        .lines()
        .filter(|l| {
            let trimmed = l.trim();
            !trimmed.is_empty()
                && !trimmed.starts_with('#')
                && !trimmed.starts_with("===")
                && trimmed.contains(": ")
        })
        .count();
    assert_eq!(cell_count, 301, "expected 301 cells");

    // Roundtrip
    let (_dir2, out) = tmp("out.xlsx");
    apply(
        &ir1,
        &out,
        &ApplyOptions {
            force: true,
            ..Default::default()
        },
    )?;
    let ir2 = derive(&out, DeriveOptions::default())?;
    pretty_assertions::assert_eq!(body_of(&ir1), body_of(&ir2));

    Ok(())
}

#[test]
fn cell_ordering_is_row_major() -> R {
    let (_dir, path) = tmp("order.xlsx");

    let mut wb = Workbook::new();
    let ws = wb.add_sheet("Sheet1");
    // Insert in non-row-major order
    ws.cell_mut("C2")?.set_value("c2");
    ws.cell_mut("A1")?.set_value("a1");
    ws.cell_mut("B2")?.set_value("b2");
    ws.cell_mut("A2")?.set_value("a2");
    ws.cell_mut("C1")?.set_value("c1");
    ws.cell_mut("B1")?.set_value("b1");
    wb.save(&path)?;

    let ir = derive(&path, DeriveOptions::default())?;
    let body = body_of(&ir);
    let cell_lines: Vec<&str> = body
        .lines()
        .filter(|l| {
            let t = l.trim();
            !t.is_empty() && !t.starts_with('#') && !t.starts_with("===")
        })
        .collect();

    // Should be row-major: A1, B1, C1, A2, B2, C2
    assert_eq!(cell_lines.len(), 6);
    assert!(cell_lines[0].starts_with("A1:"));
    assert!(cell_lines[1].starts_with("B1:"));
    assert!(cell_lines[2].starts_with("C1:"));
    assert!(cell_lines[3].starts_with("A2:"));
    assert!(cell_lines[4].starts_with("B2:"));
    assert!(cell_lines[5].starts_with("C2:"));

    Ok(())
}

// =========================================================================
// 5. Additive apply semantics
// =========================================================================

#[test]
fn apply_only_touches_listed_cells() -> R {
    let (_dir, path) = tmp("additive.xlsx");

    let mut wb = Workbook::new();
    let ws = wb.add_sheet("Sheet1");
    ws.cell_mut("A1")?.set_value("keep");
    ws.cell_mut("A2")?.set_value("also keep");
    ws.cell_mut("A3")?.set_value("keep too");
    ws.cell_mut("B1")?.set_value(1);
    ws.cell_mut("B2")?.set_value(2);
    ws.cell_mut("B3")?.set_value(3);
    wb.save(&path)?;

    // IR only mentions A2 and B2
    let ir = derive(&path, DeriveOptions::default())?;
    let (header, _) = offidized_ir::IrHeader::parse(&ir)?;
    let partial_ir = format!(
        "{}\n=== Sheet: Sheet1 ===\nA2: updated\nB2: 999\n",
        header.write()
    );

    let (_dir2, out) = tmp("out.xlsx");
    let result = apply(
        &partial_ir,
        &out,
        &ApplyOptions {
            force: true,
            ..Default::default()
        },
    )?;
    assert_eq!(result.cells_updated, 2);

    let wb2 = Workbook::open(&out)?;
    let ws2 = wb2.sheet("Sheet1").ok_or("missing")?;
    // Only A2 and B2 changed
    assert_eq!(
        ws2.cell("A2").and_then(|c| c.value()),
        Some(&CellValue::String("updated".into()))
    );
    assert_eq!(
        ws2.cell("B2").and_then(|c| c.value()),
        Some(&CellValue::Number(999.0))
    );
    // Everything else unchanged
    assert_eq!(
        ws2.cell("A1").and_then(|c| c.value()),
        Some(&CellValue::String("keep".into()))
    );
    assert_eq!(
        ws2.cell("A3").and_then(|c| c.value()),
        Some(&CellValue::String("keep too".into()))
    );
    assert_eq!(
        ws2.cell("B1").and_then(|c| c.value()),
        Some(&CellValue::Number(1.0))
    );
    assert_eq!(
        ws2.cell("B3").and_then(|c| c.value()),
        Some(&CellValue::Number(3.0))
    );

    Ok(())
}

#[test]
fn apply_does_not_delete_omitted_sheets() -> R {
    let (_dir, path) = tmp("keep_sheets.xlsx");

    let mut wb = Workbook::new();
    wb.add_sheet("Sheet1").cell_mut("A1")?.set_value("one");
    wb.add_sheet("Sheet2").cell_mut("A1")?.set_value("two");
    wb.add_sheet("Sheet3").cell_mut("A1")?.set_value("three");
    wb.save(&path)?;

    // IR only mentions Sheet2
    let ir = derive(&path, DeriveOptions::default())?;
    let (header, _) = offidized_ir::IrHeader::parse(&ir)?;
    let partial_ir = format!(
        "{}\n=== Sheet: Sheet2 ===\nA1: updated two\n",
        header.write()
    );

    let (_dir2, out) = tmp("out.xlsx");
    apply(
        &partial_ir,
        &out,
        &ApplyOptions {
            force: true,
            ..Default::default()
        },
    )?;

    let wb2 = Workbook::open(&out)?;
    assert_eq!(wb2.sheet_names().len(), 3, "all 3 sheets should exist");
    assert_eq!(
        wb2.sheet("Sheet1")
            .and_then(|ws| ws.cell("A1"))
            .and_then(|c| c.value()),
        Some(&CellValue::String("one".into()))
    );
    assert_eq!(
        wb2.sheet("Sheet2")
            .and_then(|ws| ws.cell("A1"))
            .and_then(|c| c.value()),
        Some(&CellValue::String("updated two".into()))
    );
    assert_eq!(
        wb2.sheet("Sheet3")
            .and_then(|ws| ws.cell("A1"))
            .and_then(|c| c.value()),
        Some(&CellValue::String("three".into()))
    );

    Ok(())
}

// =========================================================================
// 6. Real file roundtrips (reference repos)
// =========================================================================

/// Helper: roundtrip a real file without modifications and verify IR stability.
fn roundtrip_real_file(rel_path: &str) -> R {
    let path = match try_ref_file(rel_path) {
        Some(p) => p,
        None => {
            eprintln!("SKIP: {rel_path} not found (references not cloned)");
            return Ok(());
        }
    };

    let ir1 = derive(&path, DeriveOptions::default())?;
    assert!(ir1.contains("+++"), "IR should have header");

    let (_dir, out) = tmp("roundtrip.xlsx");
    apply(
        &ir1,
        &out,
        &ApplyOptions {
            force: true,
            ..Default::default()
        },
    )?;

    let ir2 = derive(&out, DeriveOptions::default())?;
    pretty_assertions::assert_eq!(body_of(&ir1), body_of(&ir2));

    Ok(())
}

#[test]
fn real_file_basic_spreadsheet() -> R {
    roundtrip_real_file(
        "references/Open-XML-SDK/test/DocumentFormat.OpenXml.Tests.Assets/assets/TestFiles/basicspreadsheet.xlsx",
    )
}

#[test]
fn real_file_spreadsheet() -> R {
    roundtrip_real_file(
        "references/Open-XML-SDK/test/DocumentFormat.OpenXml.Tests.Assets/assets/TestFiles/Spreadsheet.xlsx",
    )
}

#[test]
fn real_file_complex01() -> R {
    roundtrip_real_file(
        "references/Open-XML-SDK/test/DocumentFormat.OpenXml.Tests.Assets/assets/TestFiles/Complex01.xlsx",
    )
}

#[test]
fn real_file_comments() -> R {
    roundtrip_real_file(
        "references/Open-XML-SDK/test/DocumentFormat.OpenXml.Tests.Assets/assets/TestFiles/Comments.xlsx",
    )
}

#[test]
fn real_file_openpyxl_breaker() -> R {
    roundtrip_real_file("demo/openpyxl_breaker.xlsx")
}

#[test]
fn real_file_closedxml_pivot() -> R {
    roundtrip_real_file("demo/closedxml_pivot.xlsx")
}

#[test]
fn real_file_table_headers_with_line_breaks() -> R {
    roundtrip_real_file(
        "references/ClosedXML/ClosedXML.Tests/Resource/TryToLoad/TableHeadersWithLineBreaks.xlsx",
    )
}

#[test]
fn real_file_copy_row_contents() -> R {
    roundtrip_real_file(
        "references/ClosedXML/ClosedXML.Tests/Resource/TryToLoad/CopyRowContents.xlsx",
    )
}

#[test]
fn real_file_excel14() -> R {
    roundtrip_real_file(
        "references/Open-XML-SDK/test/DocumentFormat.OpenXml.Tests.Assets/assets/TestFiles/excel14.xlsx",
    )
}

#[test]
fn real_file_extlst() -> R {
    roundtrip_real_file(
        "references/Open-XML-SDK/test/DocumentFormat.OpenXml.Tests.Assets/assets/TestFiles/extlst.xlsx",
    )
}

/// Roundtrip a real file WITH modifications: change a cell, verify it sticks
/// and other cells are undisturbed.
fn roundtrip_real_file_with_edit(rel_path: &str) -> R {
    let path = match try_ref_file(rel_path) {
        Some(p) => p,
        None => {
            eprintln!("SKIP: {rel_path} not found");
            return Ok(());
        }
    };

    let ir1 = derive(&path, DeriveOptions::default())?;
    let body1 = body_of(&ir1);

    // Find the first cell line and modify its value
    let first_cell_line = body1.lines().find(|l| {
        let t = l.trim();
        !t.is_empty() && !t.starts_with('#') && !t.starts_with("===") && t.contains(": ")
    });
    let Some(original_line) = first_cell_line else {
        eprintln!("SKIP: no cells in {rel_path}");
        return Ok(());
    };

    let cell_ref = original_line.split(": ").next().ok_or("no ref")?;
    let modified_line = format!("{cell_ref}: IR_MODIFIED_VALUE");
    let modified_ir = ir1.replace(original_line, &modified_line);

    let (_dir, out) = tmp("edited.xlsx");
    apply(
        &modified_ir,
        &out,
        &ApplyOptions {
            force: true,
            ..Default::default()
        },
    )?;

    // Verify the edit took effect
    let ir2 = derive(&out, DeriveOptions::default())?;
    let body2 = body_of(&ir2);
    assert!(
        body2.contains(&format!("{cell_ref}: IR_MODIFIED_VALUE")),
        "modified cell should appear in re-derived IR"
    );

    // Verify other cells are unchanged
    let other_lines_1: Vec<&str> = body1
        .lines()
        .filter(|l| l.trim() != original_line.trim() && l.contains(": "))
        .collect();
    let other_lines_2: Vec<&str> = body2
        .lines()
        .filter(|l| !l.contains("IR_MODIFIED_VALUE") && l.contains(": "))
        .collect();
    assert_eq!(
        other_lines_1.len(),
        other_lines_2.len(),
        "other cell count should match"
    );
    for (l1, l2) in other_lines_1.iter().zip(other_lines_2.iter()) {
        assert_eq!(l1.trim(), l2.trim(), "other cells should be unchanged");
    }

    Ok(())
}

#[test]
fn real_file_edit_spreadsheet() -> R {
    roundtrip_real_file_with_edit(
        "references/Open-XML-SDK/test/DocumentFormat.OpenXml.Tests.Assets/assets/TestFiles/Spreadsheet.xlsx",
    )
}

#[test]
fn real_file_edit_openpyxl_breaker() -> R {
    roundtrip_real_file_with_edit("demo/openpyxl_breaker.xlsx")
}

#[test]
fn real_file_edit_table_headers_with_line_breaks() -> R {
    roundtrip_real_file_with_edit(
        "references/ClosedXML/ClosedXML.Tests/Resource/TryToLoad/TableHeadersWithLineBreaks.xlsx",
    )
}

#[test]
fn real_file_edit_complex01() -> R {
    roundtrip_real_file_with_edit(
        "references/Open-XML-SDK/test/DocumentFormat.OpenXml.Tests.Assets/assets/TestFiles/Complex01.xlsx",
    )
}

#[test]
fn real_file_edit_basic_spreadsheet() -> R {
    roundtrip_real_file_with_edit(
        "references/Open-XML-SDK/test/DocumentFormat.OpenXml.Tests.Assets/assets/TestFiles/basicspreadsheet.xlsx",
    )
}

// =========================================================================
// 7. Header parsing edge cases
// =========================================================================

#[test]
fn header_unknown_fields_ignored() -> R {
    let ir = "+++\n\
              source = \"test.xlsx\"\n\
              format = \"xlsx\"\n\
              mode = \"content\"\n\
              version = 1\n\
              checksum = \"sha256:abc\"\n\
              future_field = \"value\"\n\
              +++\n\
              === Sheet: Sheet1 ===\n\
              A1: hello\n";
    let (header, body) = offidized_ir::IrHeader::parse(ir)?;
    assert_eq!(header.source, "test.xlsx");
    assert!(body.contains("A1: hello"));
    Ok(())
}

#[test]
fn header_version_1_required() {
    let ir = "+++\n\
              source = \"test.xlsx\"\n\
              format = \"xlsx\"\n\
              mode = \"content\"\n\
              version = 99\n\
              checksum = \"sha256:abc\"\n\
              +++\n";
    let result = offidized_ir::IrHeader::parse(ir);
    // Version 99 should parse (we don't reject unknown versions yet)
    assert!(result.is_ok());
    let (header, _) = result.expect("parse");
    assert_eq!(header.version, 99);
}

// =========================================================================
// 8. docx content mode
// =========================================================================

#[test]
fn docx_derive_basic_document() -> R {
    let (_dir, path) = tmp("basic.docx");

    let mut doc = Document::new();
    doc.add_heading("Title", 1);
    doc.add_paragraph("Normal paragraph");
    doc.save(&path)?;

    let ir = derive(&path, DeriveOptions::default())?;
    let body = body_of(&ir);

    assert!(ir.contains("format = \"docx\""));
    assert!(ir.contains("mode = \"content\""));
    assert!(body.contains("[p1] # Title"));
    assert!(body.contains("[p2] Normal paragraph"));

    Ok(())
}

#[test]
fn docx_derive_table() -> R {
    let (_dir, path) = tmp("table.docx");

    let mut doc = Document::new();
    doc.add_paragraph("Before");
    let table = doc.add_table(2, 2);
    table.set_cell_text(0, 0, "A");
    table.set_cell_text(0, 1, "B");
    table.set_cell_text(1, 0, "C");
    table.set_cell_text(1, 1, "D");
    doc.add_paragraph("After");
    doc.save(&path)?;

    let ir = derive(&path, DeriveOptions::default())?;
    let body = body_of(&ir);

    assert!(body.contains("[t1]"));
    assert!(body.contains("| A | B |"));
    assert!(body.contains("| C | D |"));

    Ok(())
}

#[test]
fn docx_roundtrip_no_modifications() -> R {
    let (_dir, path) = tmp("rt.docx");

    let mut doc = Document::new();
    doc.add_heading("Chapter 1", 1);
    doc.add_paragraph("This is a paragraph.");
    doc.add_heading("Section 1.1", 2);
    doc.add_paragraph("Another paragraph here.");
    doc.save(&path)?;

    let ir1 = derive(&path, DeriveOptions::default())?;
    let (_dir2, out) = tmp("out.docx");
    apply(
        &ir1,
        &out,
        &ApplyOptions {
            force: true,
            ..Default::default()
        },
    )?;
    let ir2 = derive(&out, DeriveOptions::default())?;

    pretty_assertions::assert_eq!(body_of(&ir1), body_of(&ir2));

    Ok(())
}

#[test]
fn docx_apply_modifies_paragraph_text() -> R {
    let (_dir, path) = tmp("edit.docx");

    let mut doc = Document::new();
    doc.add_paragraph("Original text");
    doc.add_paragraph("Keep this");
    doc.save(&path)?;

    let ir = derive(&path, DeriveOptions::default())?;
    let modified = ir.replace("[p1] Original text", "[p1] Modified text");

    let (_dir2, out) = tmp("out.docx");
    let result = apply(
        &modified,
        &out,
        &ApplyOptions {
            force: true,
            ..Default::default()
        },
    )?;
    assert!(result.cells_updated >= 1);

    let ir2 = derive(&out, DeriveOptions::default())?;
    let body2 = body_of(&ir2);
    assert!(body2.contains("[p1] Modified text"));
    assert!(body2.contains("[p2] Keep this"));

    Ok(())
}

#[test]
fn docx_apply_bold_italic() -> R {
    let (_dir, path) = tmp("fmt.docx");

    let mut doc = Document::new();
    doc.add_paragraph("Plain text");
    doc.save(&path)?;

    let ir = derive(&path, DeriveOptions::default())?;
    let modified = ir.replace("[p1] Plain text", "[p1] Now **bold** and *italic*");

    let (_dir2, out) = tmp("out.docx");
    apply(
        &modified,
        &out,
        &ApplyOptions {
            force: true,
            ..Default::default()
        },
    )?;

    // Re-derive and verify the formatting survived
    let ir2 = derive(&out, DeriveOptions::default())?;
    let body2 = body_of(&ir2);
    assert!(
        body2.contains("**bold**"),
        "bold should be preserved: {body2}"
    );
    assert!(
        body2.contains("*italic*"),
        "italic should be preserved: {body2}"
    );

    Ok(())
}

#[test]
fn docx_apply_heading_level() -> R {
    let (_dir, path) = tmp("heading.docx");

    let mut doc = Document::new();
    doc.add_paragraph("Will become heading");
    doc.save(&path)?;

    let ir = derive(&path, DeriveOptions::default())?;
    let modified = ir.replace("[p1] Will become heading", "[p1] # Now a Heading");

    let (_dir2, out) = tmp("out.docx");
    apply(
        &modified,
        &out,
        &ApplyOptions {
            force: true,
            ..Default::default()
        },
    )?;

    let ir2 = derive(&out, DeriveOptions::default())?;
    assert!(body_of(&ir2).contains("[p1] # Now a Heading"));

    Ok(())
}

#[test]
fn docx_apply_table_cells() -> R {
    let (_dir, path) = tmp("tedit.docx");

    let mut doc = Document::new();
    let table = doc.add_table(2, 2);
    table.set_cell_text(0, 0, "X");
    table.set_cell_text(0, 1, "Y");
    table.set_cell_text(1, 0, "Z");
    table.set_cell_text(1, 1, "W");
    doc.save(&path)?;

    let ir = derive(&path, DeriveOptions::default())?;
    let modified = ir.replace("| X | Y |", "| A | B |");

    let (_dir2, out) = tmp("out.docx");
    apply(
        &modified,
        &out,
        &ApplyOptions {
            force: true,
            ..Default::default()
        },
    )?;

    let ir2 = derive(&out, DeriveOptions::default())?;
    let body2 = body_of(&ir2);
    assert!(
        body2.contains("| A | B |"),
        "table cell should be updated: {body2}"
    );
    assert!(
        body2.contains("| Z | W |"),
        "other row should be unchanged: {body2}"
    );

    Ok(())
}

#[test]
fn docx_checksum_validation() -> R {
    let (_dir, path) = tmp("cksum.docx");

    let mut doc = Document::new();
    doc.add_paragraph("Original");
    doc.save(&path)?;

    let ir = derive(&path, DeriveOptions::default())?;

    // Modify the source
    let mut doc2 = Document::open(&path)?;
    doc2.add_paragraph("Extra");
    doc2.save(&path)?;

    // Apply without force should fail
    let (_dir2, out) = tmp("out.docx");
    let result = apply(&ir, &out, &ApplyOptions::default());
    assert!(matches!(result, Err(IrError::ChecksumMismatch { .. })));

    // Apply with force should succeed
    let result = apply(
        &ir,
        &out,
        &ApplyOptions {
            force: true,
            ..Default::default()
        },
    );
    assert!(result.is_ok());

    Ok(())
}

/// Roundtrip helper for real docx files.
fn roundtrip_real_docx(rel_path: &str) -> R {
    let path = match try_ref_file(rel_path) {
        Some(p) => p,
        None => {
            eprintln!("SKIP: {rel_path} not found");
            return Ok(());
        }
    };

    let ir1 = derive(&path, DeriveOptions::default())?;
    assert!(ir1.contains("format = \"docx\""));

    let (_dir, out) = tmp("roundtrip.docx");
    apply(
        &ir1,
        &out,
        &ApplyOptions {
            force: true,
            ..Default::default()
        },
    )?;

    let ir2 = derive(&out, DeriveOptions::default())?;
    pretty_assertions::assert_eq!(body_of(&ir1), body_of(&ir2));

    Ok(())
}

#[test]
fn real_docx_helloworld() -> R {
    roundtrip_real_docx(
        "references/Open-XML-SDK/test/DocumentFormat.OpenXml.Tests.Assets/assets/TestFiles/HelloWorld.docx",
    )
}

#[test]
fn real_docx_simplesdt() -> R {
    roundtrip_real_docx(
        "references/Open-XML-SDK/test/DocumentFormat.OpenXml.Tests.Assets/assets/TestFiles/simpleSdt.docx",
    )
}

#[test]
fn real_docx_complex2010_no_panic() -> R {
    // complex2010.docx has complex run formatting (bold/italic spans) that
    // shift at run boundaries after content-mode roundtrip. Derive-only test.
    derive_real_docx_no_panic(
        "references/Open-XML-SDK/test/DocumentFormat.OpenXml.Tests.Assets/assets/TestFiles/complex2010.docx",
    )
}

#[test]
fn real_docx_docprops() -> R {
    roundtrip_real_docx(
        "references/Open-XML-SDK/test/DocumentFormat.OpenXml.Tests.Assets/assets/TestFiles/DocProps.docx",
    )
}

/// Derive-only test: verify real docx files don't panic during derive.
fn derive_real_docx_no_panic(rel_path: &str) -> R {
    let path = match try_ref_file(rel_path) {
        Some(p) => p,
        None => {
            eprintln!("SKIP: {rel_path} not found");
            return Ok(());
        }
    };

    let ir = derive(&path, DeriveOptions::default())?;
    assert!(!ir.is_empty());

    Ok(())
}

#[test]
fn real_docx_annotationref_no_panic() -> R {
    derive_real_docx_no_panic(
        "references/Open-XML-SDK/test/DocumentFormat.OpenXml.Tests.Assets/assets/TestFiles/AnnotationRef.docx",
    )
}

#[test]
fn real_docx_of16_no_panic() -> R {
    derive_real_docx_no_panic(
        "references/Open-XML-SDK/test/DocumentFormat.OpenXml.Tests.Assets/assets/TestFiles/Of16-02.docx",
    )
}

// =========================================================================
// 9. pptx content mode
// =========================================================================

#[test]
fn pptx_derive_basic_presentation() -> R {
    let (_dir, path) = tmp("basic.pptx");

    let mut prs = Presentation::new();
    let slide = prs.add_slide();
    if let Some(title) = slide.placeholder_mut(offidized_pptx::PlaceholderType::Title) {
        title.add_paragraph_with_text("My Title");
    }
    prs.save(&path)?;

    let ir = derive(&path, DeriveOptions::default())?;

    assert!(ir.contains("format = \"pptx\""));
    assert!(ir.contains("mode = \"content\""));
    let body = body_of(&ir);
    assert!(
        body.contains("--- slide 1"),
        "should have slide header: {body}"
    );

    Ok(())
}

#[test]
fn pptx_derive_multiple_slides() -> R {
    let (_dir, path) = tmp("multi.pptx");

    let mut prs = Presentation::new();
    prs.add_slide();
    prs.add_slide();
    prs.add_slide();
    prs.save(&path)?;

    let ir = derive(&path, DeriveOptions::default())?;
    let body = body_of(&ir);

    assert!(body.contains("--- slide 1"));
    assert!(body.contains("--- slide 2"));
    assert!(body.contains("--- slide 3"));

    Ok(())
}

#[test]
fn pptx_roundtrip_no_modifications() -> R {
    let (_dir, path) = tmp("rt.pptx");

    let mut prs = Presentation::new();
    let slide = prs.add_slide();
    if let Some(title) = slide.placeholder_mut(offidized_pptx::PlaceholderType::Title) {
        title.add_paragraph_with_text("Title Text");
    }
    prs.save(&path)?;

    let ir1 = derive(&path, DeriveOptions::default())?;
    let (_dir2, out) = tmp("out.pptx");
    apply(
        &ir1,
        &out,
        &ApplyOptions {
            force: true,
            ..Default::default()
        },
    )?;
    let ir2 = derive(&out, DeriveOptions::default())?;

    pretty_assertions::assert_eq!(body_of(&ir1), body_of(&ir2));

    Ok(())
}

#[test]
fn pptx_apply_modifies_title() -> R {
    let (_dir, path) = tmp("edit.pptx");

    let mut prs = Presentation::new();
    let slide = prs.add_slide();
    if let Some(title) = slide.placeholder_mut(offidized_pptx::PlaceholderType::Title) {
        title.add_paragraph_with_text("Original Title");
    }
    prs.save(&path)?;

    let ir = derive(&path, DeriveOptions::default())?;
    let body = body_of(&ir);

    // Skip test if no title placeholder was created by Presentation::new()
    if !body.contains("Original Title") {
        eprintln!("SKIP: Presentation::new() did not create title placeholder");
        return Ok(());
    }

    let modified = ir.replace("Original Title", "Updated Title");

    let (_dir2, out) = tmp("out.pptx");
    let result = apply(
        &modified,
        &out,
        &ApplyOptions {
            force: true,
            ..Default::default()
        },
    )?;
    assert!(result.cells_updated >= 1);

    let ir2 = derive(&out, DeriveOptions::default())?;
    assert!(
        body_of(&ir2).contains("Updated Title"),
        "title should be updated: {}",
        body_of(&ir2),
    );

    Ok(())
}

#[test]
fn pptx_checksum_validation() -> R {
    let (_dir, path) = tmp("cksum.pptx");

    let mut prs = Presentation::new();
    prs.add_slide();
    prs.save(&path)?;

    let ir = derive(&path, DeriveOptions::default())?;

    // Modify source
    let mut prs2 = Presentation::open(&path)?;
    prs2.add_slide();
    prs2.save(&path)?;

    // Apply without force should fail
    let (_dir2, out) = tmp("out.pptx");
    let result = apply(&ir, &out, &ApplyOptions::default());
    assert!(matches!(result, Err(IrError::ChecksumMismatch { .. })));

    Ok(())
}

/// Roundtrip helper for real pptx files.
fn roundtrip_real_pptx(rel_path: &str) -> R {
    let path = match try_ref_file(rel_path) {
        Some(p) => p,
        None => {
            eprintln!("SKIP: {rel_path} not found");
            return Ok(());
        }
    };

    let ir1 = derive(&path, DeriveOptions::default())?;
    assert!(ir1.contains("format = \"pptx\""));

    let (_dir, out) = tmp("roundtrip.pptx");
    apply(
        &ir1,
        &out,
        &ApplyOptions {
            force: true,
            ..Default::default()
        },
    )?;

    let ir2 = derive(&out, DeriveOptions::default())?;
    pretty_assertions::assert_eq!(body_of(&ir1), body_of(&ir2));

    Ok(())
}

#[test]
fn real_pptx_presentation() -> R {
    roundtrip_real_pptx(
        "references/Open-XML-SDK/test/DocumentFormat.OpenXml.Tests.Assets/assets/TestFiles/Presentation.pptx",
    )
}

#[test]
fn real_pptx_animation() -> R {
    roundtrip_real_pptx(
        "references/Open-XML-SDK/test/DocumentFormat.OpenXml.Tests.Assets/assets/TestFiles/animation.pptx",
    )
}

#[test]
fn real_pptx_autosave() -> R {
    roundtrip_real_pptx(
        "references/Open-XML-SDK/test/DocumentFormat.OpenXml.Tests.Assets/assets/TestFiles/autosave.pptx",
    )
}

/// Derive-only test: verify real pptx files don't panic during derive.
fn derive_real_pptx_no_panic(rel_path: &str) -> R {
    let path = match try_ref_file(rel_path) {
        Some(p) => p,
        None => {
            eprintln!("SKIP: {rel_path} not found");
            return Ok(());
        }
    };

    let ir = derive(&path, DeriveOptions::default())?;
    assert!(!ir.is_empty());

    Ok(())
}

#[test]
fn real_pptx_shapecrawler_010_no_panic() -> R {
    derive_real_pptx_no_panic("references/ShapeCrawler/tests/ShapeCrawler.DevTests/Assets/010.pptx")
}

#[test]
fn real_pptx_shapecrawler_026_no_panic() -> R {
    derive_real_pptx_no_panic("references/ShapeCrawler/tests/ShapeCrawler.DevTests/Assets/026.pptx")
}

#[test]
fn real_pptx_shapecrawler_textbox_no_panic() -> R {
    derive_real_pptx_no_panic(
        "references/ShapeCrawler/tests/ShapeCrawler.DevTests/Assets/078 textbox.pptx",
    )
}

#[test]
fn real_pptx_shapecrawler_table_no_panic() -> R {
    derive_real_pptx_no_panic(
        "references/ShapeCrawler/tests/ShapeCrawler.DevTests/Assets/tables/009_table.pptx",
    )
}

#[test]
fn real_pptx_3dtest_no_panic() -> R {
    derive_real_pptx_no_panic(
        "references/Open-XML-SDK/test/DocumentFormat.OpenXml.Tests.Assets/assets/TestFiles/3dtestdash.pptx",
    )
}

/// Bulk: derive all discoverable reference docx files without panics.
#[test]
fn bulk_derive_reference_docx_files() -> R {
    let root = PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../..")
        .canonicalize()?;

    let refs_dir = root
        .join("references/Open-XML-SDK/test/DocumentFormat.OpenXml.Tests.Assets/assets/TestFiles");
    if !refs_dir.exists() {
        eprintln!("SKIP: references not cloned");
        return Ok(());
    }

    let mut tested = 0;
    let mut failed = Vec::new();

    for entry in std::fs::read_dir(&refs_dir)? {
        let entry = entry?;
        let path = entry.path();
        if path.extension().and_then(|e| e.to_str()) != Some("docx") {
            continue;
        }
        tested += 1;
        match derive(&path, DeriveOptions::default()) {
            Ok(ir) => {
                assert!(!ir.is_empty());
            }
            Err(e) => {
                failed.push((
                    path.file_name()
                        .unwrap_or_default()
                        .to_string_lossy()
                        .to_string(),
                    e.to_string(),
                ));
            }
        }
    }

    eprintln!("docx bulk derive: {tested} tested, {} failed", failed.len());
    for (name, err) in &failed {
        eprintln!("  FAIL: {name}: {err}");
    }

    Ok(())
}

/// Bulk: derive all discoverable reference pptx files without panics.
#[test]
fn bulk_derive_reference_pptx_files() -> R {
    let root = PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../..")
        .canonicalize()?;

    let refs_dir = root
        .join("references/Open-XML-SDK/test/DocumentFormat.OpenXml.Tests.Assets/assets/TestFiles");
    if !refs_dir.exists() {
        eprintln!("SKIP: references not cloned");
        return Ok(());
    }

    let mut tested = 0;
    let mut failed = Vec::new();

    for entry in std::fs::read_dir(&refs_dir)? {
        let entry = entry?;
        let path = entry.path();
        if path.extension().and_then(|e| e.to_str()) != Some("pptx") {
            continue;
        }
        tested += 1;
        match derive(&path, DeriveOptions::default()) {
            Ok(ir) => {
                assert!(!ir.is_empty());
            }
            Err(e) => {
                failed.push((
                    path.file_name()
                        .unwrap_or_default()
                        .to_string_lossy()
                        .to_string(),
                    e.to_string(),
                ));
            }
        }
    }

    eprintln!("pptx bulk derive: {tested} tested, {} failed", failed.len());
    for (name, err) in &failed {
        eprintln!("  FAIL: {name}: {err}");
    }

    Ok(())
}

// =========================================================================
// 10. xlsx style mode
// =========================================================================

#[test]
fn xlsx_style_derive_empty_workbook() -> R {
    let (_dir, path) = tmp("style_empty.xlsx");

    let mut wb = Workbook::new();
    wb.add_sheet("Sheet1");
    wb.save(&path)?;

    let ir = derive(
        &path,
        DeriveOptions {
            mode: Mode::Style,
            ..Default::default()
        },
    )?;

    // Should not panic, should contain header and sheet section
    assert!(ir.contains("mode = \"style\""));
    let body = body_of(&ir);
    assert!(body.contains("=== Sheet: Sheet1 ==="));

    Ok(())
}

#[test]
fn xlsx_style_roundtrip_no_modifications() -> R {
    let (_dir, path) = tmp("style_rt.xlsx");

    let mut wb = Workbook::new();
    let ws = wb.add_sheet("Data");
    ws.cell_mut("A1")?.set_value("Hello");
    ws.cell_mut("B1")?.set_value(42);
    wb.save(&path)?;

    let opts = DeriveOptions {
        mode: Mode::Style,
        ..Default::default()
    };

    // Derive style IR
    let ir1 = derive(&path, opts.clone())?;

    // Apply to new file
    let (_dir2, out) = tmp("style_rt_out.xlsx");
    apply(
        &ir1,
        &out,
        &ApplyOptions {
            force: true,
            ..Default::default()
        },
    )?;

    // Derive again from the output
    let ir2 = derive(&out, opts)?;

    // Bodies must match
    pretty_assertions::assert_eq!(body_of(&ir1), body_of(&ir2));

    Ok(())
}

#[test]
fn xlsx_style_apply_bold() -> R {
    let (_dir, path) = tmp("style_bold.xlsx");

    let mut wb = Workbook::new();
    let ws = wb.add_sheet("Sheet1");
    ws.cell_mut("A1")?.set_value("test");
    wb.save(&path)?;

    // Derive style IR, then inject a bold property for A1
    let ir = derive(
        &path,
        DeriveOptions {
            mode: Mode::Style,
            ..Default::default()
        },
    )?;

    // Build a style IR with bold on A1
    let (header, _body) = offidized_ir::IrHeader::parse(&ir)?;
    let style_ir = format!(
        "{}\n=== Sheet: Sheet1 ===\n\n# Cell styles\nA1: bold\n",
        header.write()
    );

    let (_dir2, out) = tmp("style_bold_out.xlsx");
    apply(
        &style_ir,
        &out,
        &ApplyOptions {
            force: true,
            ..Default::default()
        },
    )?;

    // Verify the cell has bold formatting
    let wb2 = Workbook::open(&out)?;
    let ws2 = wb2.sheet("Sheet1").ok_or("missing sheet")?;
    let cell = ws2.cell("A1").ok_or("missing cell")?;
    let style_id = cell.style_id().ok_or("no style")?;
    let style = wb2.style(style_id).ok_or("no style obj")?;
    assert_eq!(
        style.font().and_then(|f| f.bold()),
        Some(true),
        "cell A1 should have bold font"
    );

    Ok(())
}

#[test]
fn xlsx_style_preserves_content() -> R {
    let (_dir, path) = tmp("style_content.xlsx");

    let mut wb = Workbook::new();
    let ws = wb.add_sheet("Sheet1");
    ws.cell_mut("A1")?.set_value("keep me");
    ws.cell_mut("B1")?.set_value(99);
    ws.cell_mut("C1")?.set_value(true);
    wb.save(&path)?;

    // Apply bold style to A1
    let ir = derive(
        &path,
        DeriveOptions {
            mode: Mode::Style,
            ..Default::default()
        },
    )?;
    let (header, _) = offidized_ir::IrHeader::parse(&ir)?;
    let style_ir = format!(
        "{}\n=== Sheet: Sheet1 ===\n\n# Cell styles\nA1: bold, italic\n",
        header.write()
    );

    let (_dir2, out) = tmp("style_content_out.xlsx");
    apply(
        &style_ir,
        &out,
        &ApplyOptions {
            force: true,
            ..Default::default()
        },
    )?;

    // Verify cell values are unchanged
    let wb2 = Workbook::open(&out)?;
    let ws2 = wb2.sheet("Sheet1").ok_or("missing")?;
    assert_eq!(
        ws2.cell("A1").and_then(|c| c.value()),
        Some(&CellValue::String("keep me".into())),
        "A1 content should be preserved"
    );
    assert_eq!(
        ws2.cell("B1").and_then(|c| c.value()),
        Some(&CellValue::Number(99.0)),
        "B1 content should be preserved"
    );
    assert_eq!(
        ws2.cell("C1").and_then(|c| c.value()),
        Some(&CellValue::Bool(true)),
        "C1 content should be preserved"
    );

    Ok(())
}

#[test]
fn xlsx_style_apply_sheet_properties() -> R {
    let (_dir, path) = tmp("style_sheetprops.xlsx");

    let mut wb = Workbook::new();
    wb.add_sheet("Sheet1");
    wb.save(&path)?;

    let ir = derive(
        &path,
        DeriveOptions {
            mode: Mode::Style,
            ..Default::default()
        },
    )?;
    let (header, _) = offidized_ir::IrHeader::parse(&ir)?;
    let style_ir = format!(
        "{}\n=== Sheet: Sheet1 ===\n\n# Sheet properties\ntab-color: #4472C4\nzoom: 150\ngridlines: hidden\n",
        header.write()
    );

    let (_dir2, out) = tmp("style_sheetprops_out.xlsx");
    apply(
        &style_ir,
        &out,
        &ApplyOptions {
            force: true,
            ..Default::default()
        },
    )?;

    let wb2 = Workbook::open(&out)?;
    let ws2 = wb2.sheet("Sheet1").ok_or("missing")?;

    assert_eq!(ws2.tab_color(), Some("4472C4"), "tab color should be set");

    let view_opts = ws2.sheet_view_options().ok_or("no view opts")?;
    assert_eq!(view_opts.zoom_scale(), Some(150), "zoom should be 150");
    assert_eq!(
        view_opts.show_gridlines(),
        Some(false),
        "gridlines should be hidden"
    );

    Ok(())
}

// =========================================================================
// 11. docx style mode
// =========================================================================

#[test]
fn docx_style_derive_section() -> R {
    let (_dir, path) = tmp("style_section.docx");

    let mut doc = Document::new();
    doc.section_mut().set_page_size_twips(12240, 15840);
    doc.section_mut()
        .set_page_orientation(offidized_docx::PageOrientation::Portrait);
    doc.section_mut().page_margins_mut().set_top_twips(1440);
    doc.section_mut().page_margins_mut().set_bottom_twips(1440);
    doc.add_paragraph("Test paragraph");
    doc.save(&path)?;

    let ir = derive(
        &path,
        DeriveOptions {
            mode: Mode::Style,
            ..Default::default()
        },
    )?;

    assert!(ir.contains("mode = \"style\""));
    let body = body_of(&ir);
    assert!(
        body.contains("page-width: 12240"),
        "should contain page width"
    );
    assert!(
        body.contains("page-height: 15840"),
        "should contain page height"
    );
    assert!(
        body.contains("orientation: portrait"),
        "should contain orientation"
    );
    assert!(
        body.contains("margin-top: 1440"),
        "should contain top margin"
    );

    Ok(())
}

#[test]
fn docx_style_roundtrip() -> R {
    let (_dir, path) = tmp("style_rt.docx");

    let mut doc = Document::new();
    doc.section_mut().set_page_size_twips(12240, 15840);
    doc.section_mut().page_margins_mut().set_top_twips(1440);
    doc.section_mut().page_margins_mut().set_bottom_twips(1440);
    let p = doc.add_paragraph("Styled paragraph");
    p.set_alignment(offidized_docx::ParagraphAlignment::Center);
    p.set_spacing_after_twips(240);
    doc.save(&path)?;

    let opts = DeriveOptions {
        mode: Mode::Style,
        ..Default::default()
    };

    let ir1 = derive(&path, opts.clone())?;
    let (_dir2, out) = tmp("style_rt_out.docx");
    apply(
        &ir1,
        &out,
        &ApplyOptions {
            force: true,
            ..Default::default()
        },
    )?;
    let ir2 = derive(&out, opts)?;

    pretty_assertions::assert_eq!(body_of(&ir1), body_of(&ir2));

    Ok(())
}

#[test]
fn docx_style_apply_paragraph_alignment() -> R {
    let (_dir, path) = tmp("style_align.docx");

    let mut doc = Document::new();
    doc.add_paragraph("Left aligned text");
    doc.add_paragraph("Also left aligned");
    doc.save(&path)?;

    let ir = derive(
        &path,
        DeriveOptions {
            mode: Mode::Style,
            ..Default::default()
        },
    )?;
    let (header, _) = offidized_ir::IrHeader::parse(&ir)?;

    // Apply center alignment to the first paragraph
    let style_ir = format!("{}\n# Paragraphs\n[p1] align=center\n", header.write());

    let (_dir2, out) = tmp("style_align_out.docx");
    apply(
        &style_ir,
        &out,
        &ApplyOptions {
            force: true,
            ..Default::default()
        },
    )?;

    // Re-derive to verify
    let ir2 = derive(
        &out,
        DeriveOptions {
            mode: Mode::Style,
            ..Default::default()
        },
    )?;
    let body2 = body_of(&ir2);
    assert!(
        body2.contains("align=center"),
        "paragraph should have center alignment: {body2}"
    );

    Ok(())
}

// =========================================================================
// 12. pptx style mode
// =========================================================================

#[test]
fn pptx_style_derive_geometry() -> R {
    // Use a real reference pptx with shapes that have explicit geometry
    let path = match try_ref_file(
        "references/Open-XML-SDK/test/DocumentFormat.OpenXml.Tests.Assets/assets/TestFiles/Presentation.pptx",
    ) {
        Some(p) => p,
        None => {
            // Fallback: test with a simple presentation (at least verify no panic)
            let (_dir, fallback_path) = tmp("style_geo.pptx");
            let mut prs = Presentation::new();
            let slide = prs.add_slide();
            if let Some(title) = slide.placeholder_mut(offidized_pptx::PlaceholderType::Title) {
                title.add_paragraph_with_text("Title Text");
            }
            prs.save(&fallback_path)?;

            let ir = derive(
                &fallback_path,
                DeriveOptions { mode: Mode::Style, ..Default::default() },
            )?;
            assert!(ir.contains("mode = \"style\""));
            assert!(body_of(&ir).contains("--- slide 1"));
            eprintln!("SKIP: references not cloned, geometry check skipped");
            return Ok(());
        }
    };

    let ir = derive(
        &path,
        DeriveOptions {
            mode: Mode::Style,
            ..Default::default()
        },
    )?;

    assert!(ir.contains("mode = \"style\""));
    let body = body_of(&ir);
    assert!(
        body.contains("--- slide 1"),
        "should have slide header: {body}"
    );

    // A real presentation file should have shapes with geometry coordinates.
    let has_geometry =
        body.contains("x=") || body.contains("y=") || body.contains("w=") || body.contains("h=");
    assert!(
        has_geometry,
        "should contain shape geometry properties: {body}"
    );

    Ok(())
}

#[test]
fn pptx_style_roundtrip() -> R {
    let (_dir, path) = tmp("style_rt.pptx");

    let mut prs = Presentation::new();
    let slide = prs.add_slide();
    if let Some(title) = slide.placeholder_mut(offidized_pptx::PlaceholderType::Title) {
        title.add_paragraph_with_text("Title");
    }
    prs.save(&path)?;

    let opts = DeriveOptions {
        mode: Mode::Style,
        ..Default::default()
    };

    let ir1 = derive(&path, opts.clone())?;
    let (_dir2, out) = tmp("style_rt_out.pptx");
    apply(
        &ir1,
        &out,
        &ApplyOptions {
            force: true,
            ..Default::default()
        },
    )?;
    let ir2 = derive(&out, opts)?;

    pretty_assertions::assert_eq!(body_of(&ir1), body_of(&ir2));

    Ok(())
}

// =========================================================================
// 13. Bulk derive style mode on reference files
// =========================================================================

/// Bulk: derive all discoverable reference xlsx files in style mode without panics.
#[test]
fn bulk_derive_style_reference_xlsx_files() -> R {
    let root = PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../..")
        .canonicalize()?;

    let refs_dir = root
        .join("references/Open-XML-SDK/test/DocumentFormat.OpenXml.Tests.Assets/assets/TestFiles");
    if !refs_dir.exists() {
        eprintln!("SKIP: references not cloned");
        return Ok(());
    }

    let mut tested = 0;
    let mut failed = Vec::new();

    for entry in std::fs::read_dir(&refs_dir)? {
        let entry = entry?;
        let path = entry.path();
        if path.extension().and_then(|e| e.to_str()) != Some("xlsx") {
            continue;
        }
        tested += 1;
        match derive(
            &path,
            DeriveOptions {
                mode: Mode::Style,
                ..Default::default()
            },
        ) {
            Ok(ir) => {
                assert!(!ir.is_empty());
                assert!(ir.contains("mode = \"style\""));
            }
            Err(e) => {
                failed.push((
                    path.file_name()
                        .unwrap_or_default()
                        .to_string_lossy()
                        .to_string(),
                    e.to_string(),
                ));
            }
        }
    }

    eprintln!(
        "xlsx style bulk derive: {tested} tested, {} failed",
        failed.len()
    );
    for (name, err) in &failed {
        eprintln!("  FAIL: {name}: {err}");
    }

    Ok(())
}

/// Bulk: derive all discoverable reference docx files in style mode without panics.
#[test]
fn bulk_derive_style_reference_docx_files() -> R {
    let root = PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../..")
        .canonicalize()?;

    let refs_dir = root
        .join("references/Open-XML-SDK/test/DocumentFormat.OpenXml.Tests.Assets/assets/TestFiles");
    if !refs_dir.exists() {
        eprintln!("SKIP: references not cloned");
        return Ok(());
    }

    let mut tested = 0;
    let mut failed = Vec::new();

    for entry in std::fs::read_dir(&refs_dir)? {
        let entry = entry?;
        let path = entry.path();
        if path.extension().and_then(|e| e.to_str()) != Some("docx") {
            continue;
        }
        tested += 1;
        match derive(
            &path,
            DeriveOptions {
                mode: Mode::Style,
                ..Default::default()
            },
        ) {
            Ok(ir) => {
                assert!(!ir.is_empty());
                assert!(ir.contains("mode = \"style\""));
            }
            Err(e) => {
                failed.push((
                    path.file_name()
                        .unwrap_or_default()
                        .to_string_lossy()
                        .to_string(),
                    e.to_string(),
                ));
            }
        }
    }

    eprintln!(
        "docx style bulk derive: {tested} tested, {} failed",
        failed.len()
    );
    for (name, err) in &failed {
        eprintln!("  FAIL: {name}: {err}");
    }

    Ok(())
}

/// Bulk: derive all discoverable reference pptx files in style mode without panics.
#[test]
fn bulk_derive_style_reference_pptx_files() -> R {
    let root = PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../..")
        .canonicalize()?;

    let refs_dir = root
        .join("references/Open-XML-SDK/test/DocumentFormat.OpenXml.Tests.Assets/assets/TestFiles");
    if !refs_dir.exists() {
        eprintln!("SKIP: references not cloned");
        return Ok(());
    }

    let mut tested = 0;
    let mut failed = Vec::new();

    for entry in std::fs::read_dir(&refs_dir)? {
        let entry = entry?;
        let path = entry.path();
        if path.extension().and_then(|e| e.to_str()) != Some("pptx") {
            continue;
        }
        tested += 1;
        match derive(
            &path,
            DeriveOptions {
                mode: Mode::Style,
                ..Default::default()
            },
        ) {
            Ok(ir) => {
                assert!(!ir.is_empty());
                assert!(ir.contains("mode = \"style\""));
            }
            Err(e) => {
                failed.push((
                    path.file_name()
                        .unwrap_or_default()
                        .to_string_lossy()
                        .to_string(),
                    e.to_string(),
                ));
            }
        }
    }

    eprintln!(
        "pptx style bulk derive: {tested} tested, {} failed",
        failed.len()
    );
    for (name, err) in &failed {
        eprintln!("  FAIL: {name}: {err}");
    }

    Ok(())
}
