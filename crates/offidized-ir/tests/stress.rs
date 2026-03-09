//! Stress tests targeting edge cases and potential bugs in the IR layer.

#![allow(clippy::expect_used, clippy::panic_in_result_fn)]

use std::path::PathBuf;

use offidized_ir::{apply, derive, ApplyOptions, DeriveOptions, IrError};
use offidized_xlsx::{CellValue, Workbook};

type R = Result<(), Box<dyn std::error::Error>>;

fn tmp(name: &str) -> (tempfile::TempDir, PathBuf) {
    let dir = tempfile::tempdir().expect("tempdir");
    let path = dir.path().join(name);
    (dir, path)
}

fn body_of(ir: &str) -> String {
    offidized_ir::IrHeader::parse(ir).expect("header").1
}

fn try_ref_file(rel: &str) -> Option<PathBuf> {
    let root = PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../..")
        .canonicalize()
        .ok()?;
    let p = root.join(rel);
    p.exists().then_some(p)
}

// =========================================================================
// Hash-prefixed strings (the # quoting problem)
// =========================================================================

#[test]
fn string_starting_with_hash_roundtrips() -> R {
    // Strings starting with # could be confused with error values.
    // CellValue::String("#N/A") must not become CellValue::Error("#N/A").
    let (_dir, path) = tmp("hash.xlsx");

    let mut wb = Workbook::new();
    let ws = wb.add_sheet("Sheet1");
    ws.cell_mut("A1")?
        .set_value(CellValue::String("#N/A".into()));
    ws.cell_mut("A2")?
        .set_value(CellValue::String("#hashtag".into()));
    ws.cell_mut("A3")?
        .set_value(CellValue::String("#123".into()));
    ws.cell_mut("A4")?
        .set_value(CellValue::String("# comment".into()));
    ws.cell_mut("A5")?.set_value(CellValue::String("#".into()));
    wb.save(&path)?;

    let ir = derive(&path, DeriveOptions::default())?;
    let body = body_of(&ir);

    // These should all be quoted to avoid confusion with error values
    assert!(
        body.contains(r##"A1: "#N/A""##),
        "String #N/A must be quoted, got:\n{body}"
    );
    assert!(
        body.contains(r##"A2: "#hashtag""##),
        "String #hashtag must be quoted"
    );
    assert!(
        body.contains(r##"A3: "#123""##),
        "String #123 must be quoted"
    );
    assert!(
        body.contains(r##"A4: "# comment""##),
        "String '# comment' must be quoted"
    );
    assert!(body.contains(r##"A5: "#""##), "String '#' must be quoted");

    // Apply back and verify they come back as String, not Error
    let (_dir2, out) = tmp("out.xlsx");
    apply(
        &ir,
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
        Some(&CellValue::String("#N/A".into()))
    );
    assert_eq!(
        ws2.cell("A2").and_then(|c| c.value()),
        Some(&CellValue::String("#hashtag".into()))
    );
    assert_eq!(
        ws2.cell("A3").and_then(|c| c.value()),
        Some(&CellValue::String("#123".into()))
    );
    assert_eq!(
        ws2.cell("A4").and_then(|c| c.value()),
        Some(&CellValue::String("# comment".into()))
    );
    assert_eq!(
        ws2.cell("A5").and_then(|c| c.value()),
        Some(&CellValue::String("#".into()))
    );

    Ok(())
}

#[test]
fn actual_error_values_roundtrip() -> R {
    // Real error values must survive as errors (not become quoted strings).
    let (_dir, path) = tmp("errors.xlsx");

    let mut wb = Workbook::new();
    let ws = wb.add_sheet("Sheet1");
    ws.cell_mut("A1")?
        .set_value(CellValue::Error("#NULL!".into()));
    ws.cell_mut("A2")?
        .set_value(CellValue::Error("#DIV/0!".into()));
    ws.cell_mut("A3")?
        .set_value(CellValue::Error("#VALUE!".into()));
    ws.cell_mut("A4")?
        .set_value(CellValue::Error("#REF!".into()));
    ws.cell_mut("A5")?
        .set_value(CellValue::Error("#NAME?".into()));
    ws.cell_mut("A6")?
        .set_value(CellValue::Error("#NUM!".into()));
    ws.cell_mut("A7")?
        .set_value(CellValue::Error("#N/A".into()));
    wb.save(&path)?;

    let ir = derive(&path, DeriveOptions::default())?;
    let (_dir2, out) = tmp("out.xlsx");
    apply(
        &ir,
        &out,
        &ApplyOptions {
            force: true,
            ..Default::default()
        },
    )?;

    let wb2 = Workbook::open(&out)?;
    let ws2 = wb2.sheet("Sheet1").ok_or("missing")?;
    for (i, expected) in [
        "#NULL!", "#DIV/0!", "#VALUE!", "#REF!", "#NAME?", "#NUM!", "#N/A",
    ]
    .iter()
    .enumerate()
    {
        let cell_ref = format!("A{}", i + 1);
        assert_eq!(
            ws2.cell(&cell_ref).and_then(|c| c.value()),
            Some(&CellValue::Error((*expected).to_string())),
            "error {expected} should roundtrip at {cell_ref}"
        );
    }

    Ok(())
}

// =========================================================================
// Sheet names with special characters
// =========================================================================

#[test]
fn sheet_name_with_spaces() -> R {
    let (_dir, path) = tmp("spaces.xlsx");

    let mut wb = Workbook::new();
    wb.add_sheet("My Sheet Name")
        .cell_mut("A1")?
        .set_value("data");
    wb.save(&path)?;

    let ir = derive(&path, DeriveOptions::default())?;
    assert!(body_of(&ir).contains("=== Sheet: My Sheet Name ==="));

    let (_dir2, out) = tmp("out.xlsx");
    apply(
        &ir,
        &out,
        &ApplyOptions {
            force: true,
            ..Default::default()
        },
    )?;
    let wb2 = Workbook::open(&out)?;
    assert_eq!(
        wb2.sheet("My Sheet Name")
            .and_then(|ws| ws.cell("A1"))
            .and_then(|c| c.value()),
        Some(&CellValue::String("data".into())),
    );
    Ok(())
}

#[test]
fn sheet_name_with_special_chars() -> R {
    let (_dir, path) = tmp("special.xlsx");

    let mut wb = Workbook::new();
    // Excel allows these characters in sheet names
    wb.add_sheet("Q1 (2025)").cell_mut("A1")?.set_value("q1");
    wb.add_sheet("Revenue & Costs")
        .cell_mut("A1")?
        .set_value("r");
    wb.add_sheet("Sheet-2").cell_mut("A1")?.set_value("s2");
    wb.add_sheet("100%").cell_mut("A1")?.set_value("pct");
    wb.save(&path)?;

    let ir = derive(&path, DeriveOptions::default())?;
    let body = body_of(&ir);
    assert!(body.contains("=== Sheet: Q1 (2025) ==="));
    assert!(body.contains("=== Sheet: Revenue & Costs ==="));
    assert!(body.contains("=== Sheet: Sheet-2 ==="));
    assert!(body.contains("=== Sheet: 100% ==="));

    let (_dir2, out) = tmp("out.xlsx");
    apply(
        &ir,
        &out,
        &ApplyOptions {
            force: true,
            ..Default::default()
        },
    )?;
    let wb2 = Workbook::open(&out)?;
    assert_eq!(wb2.sheet_names().len(), 4);
    assert_eq!(
        wb2.sheet("100%")
            .and_then(|ws| ws.cell("A1"))
            .and_then(|c| c.value()),
        Some(&CellValue::String("pct".into())),
    );
    Ok(())
}

// =========================================================================
// Unicode
// =========================================================================

#[test]
fn unicode_values_roundtrip() -> R {
    let (_dir, path) = tmp("unicode.xlsx");

    let mut wb = Workbook::new();
    let ws = wb.add_sheet("Sheet1");
    ws.cell_mut("A1")?.set_value("日本語テスト");
    ws.cell_mut("A2")?.set_value("Ñoño café");
    ws.cell_mut("A3")?.set_value("emoji 🎉🚀");
    ws.cell_mut("A4")?.set_value("中文数据");
    ws.cell_mut("A5")?.set_value("한국어");
    ws.cell_mut("A6")?.set_value("العربية");
    wb.save(&path)?;

    let ir = derive(&path, DeriveOptions::default())?;
    let (_dir2, out) = tmp("out.xlsx");
    apply(
        &ir,
        &out,
        &ApplyOptions {
            force: true,
            ..Default::default()
        },
    )?;

    let ir2 = derive(&out, DeriveOptions::default())?;
    pretty_assertions::assert_eq!(body_of(&ir), body_of(&ir2));

    let wb2 = Workbook::open(&out)?;
    let ws2 = wb2.sheet("Sheet1").ok_or("missing")?;
    assert_eq!(
        ws2.cell("A1").and_then(|c| c.value()),
        Some(&CellValue::String("日本語テスト".into()))
    );
    assert_eq!(
        ws2.cell("A3").and_then(|c| c.value()),
        Some(&CellValue::String("emoji 🎉🚀".into()))
    );
    Ok(())
}

#[test]
fn unicode_sheet_name() -> R {
    let (_dir, path) = tmp("unicode_sheet.xlsx");

    let mut wb = Workbook::new();
    wb.add_sheet("数据表").cell_mut("A1")?.set_value("中文");
    wb.save(&path)?;

    let ir = derive(&path, DeriveOptions::default())?;
    assert!(body_of(&ir).contains("=== Sheet: 数据表 ==="));

    let (_dir2, out) = tmp("out.xlsx");
    apply(
        &ir,
        &out,
        &ApplyOptions {
            force: true,
            ..Default::default()
        },
    )?;
    let wb2 = Workbook::open(&out)?;
    assert!(wb2.sheet("数据表").is_some());
    Ok(())
}

// =========================================================================
// Number edge cases
// =========================================================================

#[test]
fn number_edge_cases_roundtrip() -> R {
    let (_dir, path) = tmp("numbers.xlsx");

    let mut wb = Workbook::new();
    let ws = wb.add_sheet("Sheet1");
    ws.cell_mut("A1")?.set_value(CellValue::Number(0.0));
    ws.cell_mut("A2")?.set_value(CellValue::Number(-0.0));
    ws.cell_mut("A3")?.set_value(CellValue::Number(0.1));
    ws.cell_mut("A4")?.set_value(CellValue::Number(0.1 + 0.2)); // 0.30000000000000004
    ws.cell_mut("A5")?.set_value(CellValue::Number(1e10));
    ws.cell_mut("A6")?.set_value(CellValue::Number(1e-10));
    ws.cell_mut("A7")?.set_value(CellValue::Number(f64::MAX));
    ws.cell_mut("A8")?.set_value(CellValue::Number(f64::MIN));
    ws.cell_mut("A9")?
        .set_value(CellValue::Number(f64::MIN_POSITIVE));
    ws.cell_mut("A10")?
        .set_value(CellValue::Number(999999999999999.0)); // 15 digits
    ws.cell_mut("A11")?
        .set_value(CellValue::Number(-999999999999999.0));
    ws.cell_mut("A12")?.set_value(CellValue::Number(1.0));
    ws.cell_mut("A13")?.set_value(CellValue::Number(-1.0));
    wb.save(&path)?;

    let ir = derive(&path, DeriveOptions::default())?;
    let body = body_of(&ir);

    // Zero and negative zero both → "0"
    let zero_lines: Vec<&str> = body
        .lines()
        .filter(|l| l.starts_with("A1:") || l.starts_with("A2:"))
        .collect();
    for line in &zero_lines {
        assert!(line.ends_with(": 0"), "zero should be '0': {line}");
    }

    // Roundtrip
    let (_dir2, out) = tmp("out.xlsx");
    apply(
        &ir,
        &out,
        &ApplyOptions {
            force: true,
            ..Default::default()
        },
    )?;

    let wb2 = Workbook::open(&out)?;
    let ws2 = wb2.sheet("Sheet1").ok_or("missing")?;

    // Verify precision-sensitive values
    if let Some(CellValue::Number(n)) = ws2.cell("A4").and_then(|c| c.value()) {
        assert!(
            (*n - (0.1 + 0.2)).abs() < f64::EPSILON,
            "0.1+0.2 precision lost: {n}"
        );
    }
    if let Some(CellValue::Number(n)) = ws2.cell("A6").and_then(|c| c.value()) {
        assert!(
            (*n - 1e-10).abs() < f64::EPSILON,
            "1e-10 precision lost: {n}"
        );
    }

    Ok(())
}

// =========================================================================
// Carriage returns and mixed line endings
// =========================================================================

#[test]
fn carriage_return_in_cell_value() -> R {
    let (_dir, path) = tmp("cr.xlsx");

    let mut wb = Workbook::new();
    let ws = wb.add_sheet("Sheet1");
    ws.cell_mut("A1")?
        .set_value(CellValue::String("line1\r\nline2".into()));
    ws.cell_mut("A2")?
        .set_value(CellValue::String("just\rcarriage".into()));
    wb.save(&path)?;

    let ir = derive(&path, DeriveOptions::default())?;
    // Should not have literal \r or \n in cell lines (would break parsing)
    for line in body_of(&ir).lines() {
        if line.starts_with("A1:") || line.starts_with("A2:") {
            assert!(
                !line.contains('\r'),
                "cell line should not contain literal CR: {line:?}"
            );
        }
    }

    let (_dir2, out) = tmp("out.xlsx");
    apply(
        &ir,
        &out,
        &ApplyOptions {
            force: true,
            ..Default::default()
        },
    )?;

    // Verify content survived (even if \r might be normalized)
    let wb2 = Workbook::open(&out)?;
    let ws2 = wb2.sheet("Sheet1").ok_or("missing")?;
    let a1 = ws2.cell("A1").and_then(|c| c.value());
    assert!(a1.is_some(), "A1 should have a value");

    Ok(())
}

// =========================================================================
// Idempotent apply
// =========================================================================

#[test]
fn apply_same_ir_twice_is_idempotent() -> R {
    let (_dir, path) = tmp("idempotent.xlsx");

    let mut wb = Workbook::new();
    let ws = wb.add_sheet("Sheet1");
    ws.cell_mut("A1")?.set_value("hello");
    ws.cell_mut("B1")?.set_value(42);
    ws.cell_mut("C1")?.set_formula("A1&B1");
    wb.save(&path)?;

    let ir = derive(&path, DeriveOptions::default())?;

    // Apply once
    let (_dir2, out1) = tmp("out1.xlsx");
    apply(
        &ir,
        &out1,
        &ApplyOptions {
            force: true,
            ..Default::default()
        },
    )?;
    let ir_after_1 = derive(&out1, DeriveOptions::default())?;

    // Apply again to the result
    let (_dir3, out2) = tmp("out2.xlsx");
    apply(
        &ir_after_1,
        &out2,
        &ApplyOptions {
            force: true,
            ..Default::default()
        },
    )?;
    let ir_after_2 = derive(&out2, DeriveOptions::default())?;

    pretty_assertions::assert_eq!(body_of(&ir_after_1), body_of(&ir_after_2));

    Ok(())
}

// =========================================================================
// Long/pathological values
// =========================================================================

#[test]
fn very_long_cell_value() -> R {
    let (_dir, path) = tmp("long.xlsx");
    let long_string = "x".repeat(10_000);

    let mut wb = Workbook::new();
    wb.add_sheet("Sheet1")
        .cell_mut("A1")?
        .set_value(CellValue::String(long_string.clone()));
    wb.save(&path)?;

    let ir = derive(&path, DeriveOptions::default())?;
    let (_dir2, out) = tmp("out.xlsx");
    apply(
        &ir,
        &out,
        &ApplyOptions {
            force: true,
            ..Default::default()
        },
    )?;

    let wb2 = Workbook::open(&out)?;
    assert_eq!(
        wb2.sheet("Sheet1")
            .and_then(|ws| ws.cell("A1"))
            .and_then(|c| c.value()),
        Some(&CellValue::String(long_string)),
    );
    Ok(())
}

#[test]
fn cell_value_with_many_special_chars() -> R {
    let (_dir, path) = tmp("special_val.xlsx");
    let tricky = "tab\there | pipe | =not_formula | \"quotes\" | #not_error | <empty> but not | true but string";

    let mut wb = Workbook::new();
    wb.add_sheet("Sheet1")
        .cell_mut("A1")?
        .set_value(CellValue::String(tricky.into()));
    wb.save(&path)?;

    let ir = derive(&path, DeriveOptions::default())?;
    let (_dir2, out) = tmp("out.xlsx");
    apply(
        &ir,
        &out,
        &ApplyOptions {
            force: true,
            ..Default::default()
        },
    )?;

    let wb2 = Workbook::open(&out)?;
    assert_eq!(
        wb2.sheet("Sheet1")
            .and_then(|ws| ws.cell("A1"))
            .and_then(|c| c.value()),
        Some(&CellValue::String(tricky.into())),
    );
    Ok(())
}

// =========================================================================
// Empty string vs blank cell
// =========================================================================

#[test]
fn empty_string_distinct_from_blank() -> R {
    let (_dir, path) = tmp("empty_vs_blank.xlsx");

    let mut wb = Workbook::new();
    let ws = wb.add_sheet("Sheet1");
    ws.cell_mut("A1")?.set_value(CellValue::String("".into())); // empty string
                                                                // A2 is just not set (blank/missing)
    ws.cell_mut("A3")?.set_value("not empty");
    wb.save(&path)?;

    let ir = derive(&path, DeriveOptions::default())?;
    let body = body_of(&ir);

    // A1 should appear as quoted empty string
    assert!(
        body.contains("A1: \"\""),
        "empty string should be quoted: {body}"
    );
    // A2 should not appear at all (blank)
    assert!(!body.contains("A2:"), "blank cell should not appear");

    // Roundtrip
    let (_dir2, out) = tmp("out.xlsx");
    apply(
        &ir,
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
        Some(&CellValue::String("".into()))
    );
    // A2 might exist as a default cell from the IR (it shouldn't)
    // Actually A2 isn't in the IR so it should be untouched from source
    assert_eq!(
        ws2.cell("A3").and_then(|c| c.value()),
        Some(&CellValue::String("not empty".into()))
    );

    Ok(())
}

// =========================================================================
// Formula with newlines (structured references)
// =========================================================================

#[test]
fn formula_with_newline_roundtrips() -> R {
    let (_dir, path) = tmp("formula_nl.xlsx");

    let mut wb = Workbook::new();
    let ws = wb.add_sheet("Sheet1");
    // Structured reference with newline in column name (as seen in ClosedXML test)
    ws.cell_mut("A1")?
        .set_formula("SUBTOTAL(109,Table1[Purchase\n price])");
    ws.cell_mut("A2")?.set_formula("SUM(A1:A1)"); // normal formula for comparison
    wb.save(&path)?;

    let ir = derive(&path, DeriveOptions::default())?;
    let body = body_of(&ir);

    // Formula should be on one line (newline escaped)
    for line in body.lines() {
        if line.starts_with("A1:") {
            assert!(
                line.contains("\\n"),
                "formula newline should be escaped: {line}"
            );
            assert!(
                !line.contains('\n') || line == line.trim(),
                "should not have literal newline"
            );
        }
    }

    // Roundtrip
    let (_dir2, out) = tmp("out.xlsx");
    apply(
        &ir,
        &out,
        &ApplyOptions {
            force: true,
            ..Default::default()
        },
    )?;

    let wb2 = Workbook::open(&out)?;
    let ws2 = wb2.sheet("Sheet1").ok_or("missing")?;
    assert_eq!(
        ws2.cell("A1").and_then(|c| c.formula()),
        Some("SUBTOTAL(109,Table1[Purchase\n price])"),
        "formula with newline should roundtrip"
    );
    assert_eq!(ws2.cell("A2").and_then(|c| c.formula()), Some("SUM(A1:A1)"));

    Ok(())
}

// =========================================================================
// IR body with comment lines and blank lines
// =========================================================================

#[test]
fn apply_ignores_comments_and_blanks() -> R {
    let (_dir, path) = tmp("comments.xlsx");

    let mut wb = Workbook::new();
    wb.add_sheet("Sheet1").cell_mut("A1")?.set_value("orig");
    wb.save(&path)?;

    let ir = derive(&path, DeriveOptions::default())?;
    let (header, _) = offidized_ir::IrHeader::parse(&ir)?;

    // Build IR with lots of comments and blank lines
    let custom_ir = format!(
        "{}\n\
         # This is a comment\n\
         \n\
         === Sheet: Sheet1 ===\n\
         # Another comment\n\
         \n\
         A1: updated\n\
         # More comments\n\
         B1: new\n\
         \n\
         \n",
        header.write()
    );

    let (_dir2, out) = tmp("out.xlsx");
    let result = apply(
        &custom_ir,
        &out,
        &ApplyOptions {
            force: true,
            ..Default::default()
        },
    )?;
    assert_eq!(result.cells_updated, 1); // A1
    assert_eq!(result.cells_created, 1); // B1

    let wb2 = Workbook::open(&out)?;
    let ws2 = wb2.sheet("Sheet1").ok_or("missing")?;
    assert_eq!(
        ws2.cell("A1").and_then(|c| c.value()),
        Some(&CellValue::String("updated".into()))
    );
    assert_eq!(
        ws2.cell("B1").and_then(|c| c.value()),
        Some(&CellValue::String("new".into()))
    );

    Ok(())
}

// =========================================================================
// Case sensitivity in cell references
// =========================================================================

#[test]
fn cell_ref_case_insensitive_on_apply() -> R {
    let (_dir, path) = tmp("case.xlsx");

    let mut wb = Workbook::new();
    wb.add_sheet("Sheet1").cell_mut("A1")?.set_value("orig");
    wb.save(&path)?;

    let ir = derive(&path, DeriveOptions::default())?;
    let (header, _) = offidized_ir::IrHeader::parse(&ir)?;

    // Apply using lowercase cell ref
    let custom_ir = format!(
        "{}\n=== Sheet: Sheet1 ===\na1: from lowercase\n",
        header.write()
    );

    let (_dir2, out) = tmp("out.xlsx");
    apply(
        &custom_ir,
        &out,
        &ApplyOptions {
            force: true,
            ..Default::default()
        },
    )?;

    let wb2 = Workbook::open(&out)?;
    let ws2 = wb2.sheet("Sheet1").ok_or("missing")?;
    // Should work regardless of case
    let a1 = ws2.cell("A1").and_then(|c| c.value());
    assert_eq!(a1, Some(&CellValue::String("from lowercase".into())));

    Ok(())
}

// =========================================================================
// Many sheets
// =========================================================================

#[test]
fn many_sheets_roundtrip() -> R {
    let (_dir, path) = tmp("many_sheets.xlsx");

    let mut wb = Workbook::new();
    for i in 1..=20 {
        let name = format!("Sheet{i}");
        wb.add_sheet(&name)
            .cell_mut("A1")?
            .set_value(format!("data_{i}"));
    }
    wb.save(&path)?;

    let ir = derive(&path, DeriveOptions::default())?;
    let body = body_of(&ir);

    // All 20 sheets should appear
    for i in 1..=20 {
        assert!(body.contains(&format!("=== Sheet: Sheet{i} ===")));
        assert!(body.contains(&format!("A1: data_{i}")));
    }

    // Roundtrip
    let (_dir2, out) = tmp("out.xlsx");
    apply(
        &ir,
        &out,
        &ApplyOptions {
            force: true,
            ..Default::default()
        },
    )?;
    let ir2 = derive(&out, DeriveOptions::default())?;
    pretty_assertions::assert_eq!(body_of(&ir), body_of(&ir2));

    Ok(())
}

// =========================================================================
// Sparse data (cells far apart)
// =========================================================================

#[test]
fn sparse_cells_roundtrip() -> R {
    let (_dir, path) = tmp("sparse.xlsx");

    let mut wb = Workbook::new();
    let ws = wb.add_sheet("Sheet1");
    ws.cell_mut("A1")?.set_value("top-left");
    ws.cell_mut("Z1")?.set_value("top-right");
    ws.cell_mut("A100")?.set_value("bottom-left");
    ws.cell_mut("Z100")?.set_value("bottom-right");
    ws.cell_mut("M50")?.set_value("middle");
    wb.save(&path)?;

    let ir = derive(&path, DeriveOptions::default())?;
    let body = body_of(&ir);
    let cell_lines: Vec<&str> = body
        .lines()
        .filter(|l| l.contains(": ") && !l.starts_with("===") && !l.starts_with('#'))
        .collect();

    // Should only have 5 cells (sparse)
    assert_eq!(cell_lines.len(), 5);

    // Row-major ordering check
    assert!(cell_lines[0].starts_with("A1:"));
    assert!(cell_lines[1].starts_with("Z1:"));
    assert!(cell_lines[2].starts_with("M50:"));
    assert!(cell_lines[3].starts_with("A100:"));
    assert!(cell_lines[4].starts_with("Z100:"));

    // Roundtrip
    let (_dir2, out) = tmp("out.xlsx");
    apply(
        &ir,
        &out,
        &ApplyOptions {
            force: true,
            ..Default::default()
        },
    )?;
    let ir2 = derive(&out, DeriveOptions::default())?;
    pretty_assertions::assert_eq!(body_of(&ir), body_of(&ir2));

    Ok(())
}

// =========================================================================
// Real file stress: derive every reference xlsx, verify no panics
// =========================================================================

fn derive_without_panic(rel_path: &str) {
    let Some(path) = try_ref_file(rel_path) else {
        return;
    };
    // Just derive — don't apply. We want to ensure no panics on any file.
    let _ = derive(&path, DeriveOptions::default());
}

#[test]
fn no_panic_on_reference_files() {
    let files = [
        "references/Open-XML-SDK/test/DocumentFormat.OpenXml.Tests.Assets/assets/TestFiles/basicspreadsheet.xlsx",
        "references/Open-XML-SDK/test/DocumentFormat.OpenXml.Tests.Assets/assets/TestFiles/Spreadsheet.xlsx",
        "references/Open-XML-SDK/test/DocumentFormat.OpenXml.Tests.Assets/assets/TestFiles/Complex01.xlsx",
        "references/Open-XML-SDK/test/DocumentFormat.OpenXml.Tests.Assets/assets/TestFiles/Comments.xlsx",
        "references/Open-XML-SDK/test/DocumentFormat.OpenXml.Tests.Assets/assets/TestFiles/excel14.xlsx",
        "references/Open-XML-SDK/test/DocumentFormat.OpenXml.Tests.Assets/assets/TestFiles/extlst.xlsx",
        "references/Open-XML-SDK/test/DocumentFormat.OpenXml.Tests.Assets/assets/TestFiles/MCExecl.xlsx",
        "references/Open-XML-SDK/test/DocumentFormat.OpenXml.Tests.Assets/assets/TestFiles/vmldrawingroot.xlsx",
        "references/Open-XML-SDK/test/DocumentFormat.OpenXml.Tests.Assets/assets/TestFiles/Youtube.xlsx",
        "references/ClosedXML/ClosedXML.Tests/Resource/TryToLoad/TableHeadersWithLineBreaks.xlsx",
        "references/ClosedXML/ClosedXML.Tests/Resource/TryToLoad/CopyRowContents.xlsx",
        "references/ClosedXML/ClosedXML.Tests/Resource/TryToLoad/AllShapes.xlsx",
        "references/ClosedXML/ClosedXML.Tests/Resource/TryToLoad/EmptyCellValue.xlsx",
        "references/ClosedXML/ClosedXML.Tests/Resource/TryToLoad/EmptyTable.xlsx",
        "references/ClosedXML/ClosedXML.Tests/Resource/TryToLoad/EmptyStyles.xlsx",
        "references/ClosedXML/ClosedXML.Tests/Resource/TryToLoad/BaseColumnWidth.xlsx",
        "references/ClosedXML/ClosedXML.Tests/Resource/TryToLoad/Date1904System.xlsx",
        "references/ClosedXML/ClosedXML.Tests/Resource/TryToLoad/LoadSheetsWithCommas.xlsx",
        "references/ClosedXML/ClosedXML.Tests/Resource/TryToLoad/PivotTableWithTableSource.xlsx",
        "demo/openpyxl_breaker.xlsx",
        "demo/closedxml_pivot.xlsx",
    ];

    for file in &files {
        derive_without_panic(file);
    }
}

/// Roundtrip every reference file that opens successfully.
#[test]
fn bulk_roundtrip_reference_files() {
    let files = [
        "references/Open-XML-SDK/test/DocumentFormat.OpenXml.Tests.Assets/assets/TestFiles/basicspreadsheet.xlsx",
        "references/Open-XML-SDK/test/DocumentFormat.OpenXml.Tests.Assets/assets/TestFiles/Spreadsheet.xlsx",
        "references/Open-XML-SDK/test/DocumentFormat.OpenXml.Tests.Assets/assets/TestFiles/Complex01.xlsx",
        "references/Open-XML-SDK/test/DocumentFormat.OpenXml.Tests.Assets/assets/TestFiles/Comments.xlsx",
        "references/Open-XML-SDK/test/DocumentFormat.OpenXml.Tests.Assets/assets/TestFiles/excel14.xlsx",
        "references/Open-XML-SDK/test/DocumentFormat.OpenXml.Tests.Assets/assets/TestFiles/extlst.xlsx",
        "references/ClosedXML/ClosedXML.Tests/Resource/TryToLoad/TableHeadersWithLineBreaks.xlsx",
        "references/ClosedXML/ClosedXML.Tests/Resource/TryToLoad/CopyRowContents.xlsx",
        "references/ClosedXML/ClosedXML.Tests/Resource/TryToLoad/EmptyCellValue.xlsx",
        "references/ClosedXML/ClosedXML.Tests/Resource/TryToLoad/EmptyTable.xlsx",
        "references/ClosedXML/ClosedXML.Tests/Resource/TryToLoad/BaseColumnWidth.xlsx",
        "references/ClosedXML/ClosedXML.Tests/Resource/TryToLoad/LoadSheetsWithCommas.xlsx",
        "demo/openpyxl_breaker.xlsx",
        "demo/closedxml_pivot.xlsx",
    ];

    let mut passed = 0;
    let mut skipped = 0;

    for file in &files {
        let Some(path) = try_ref_file(file) else {
            skipped += 1;
            continue;
        };

        let Ok(ir1) = derive(&path, DeriveOptions::default()) else {
            continue; // File couldn't be opened
        };

        let dir = tempfile::tempdir().expect("tempdir");
        let out = dir.path().join("roundtrip.xlsx");

        let Ok(_) = apply(
            &ir1,
            &out,
            &ApplyOptions {
                force: true,
                ..Default::default()
            },
        ) else {
            panic!("apply failed for {file}");
        };

        let Ok(ir2) = derive(&out, DeriveOptions::default()) else {
            panic!("re-derive failed for {file}");
        };

        let b1 = body_of(&ir1);
        let b2 = body_of(&ir2);
        assert_eq!(b1, b2, "roundtrip mismatch for {file}");
        passed += 1;
    }

    eprintln!("bulk roundtrip: {passed} passed, {skipped} skipped");
    assert!(passed > 0, "at least some files should pass");
}

// =========================================================================
// Derive unsupported format gives clear error
// =========================================================================

#[test]
fn derive_unsupported_format() {
    let result = derive(std::path::Path::new("file.txt"), DeriveOptions::default());
    assert!(matches!(result, Err(IrError::UnsupportedFormat(_))));
}

#[test]
fn derive_full_mode_contains_both_sections() {
    let (_dir, path) = tmp("test.xlsx");
    let mut wb = Workbook::new();
    let ws = wb.add_sheet("Sheet1");
    ws.cell_mut("A1").expect("cell").set_value("hello");
    wb.save(&path).expect("save");
    let ir = derive(
        &path,
        DeriveOptions {
            mode: offidized_ir::Mode::Full,
            ..Default::default()
        },
    )
    .expect("derive full");
    // Should contain the content section
    assert!(ir.contains("=== Sheet: Sheet1 ==="));
    assert!(ir.contains("hello"));
    // Should contain the style separator
    assert!(ir.contains("--- style ---"));
}

// =========================================================================
// Apply with Windows-style line endings in IR
// =========================================================================

#[test]
fn apply_handles_crlf_ir() -> R {
    let (_dir, path) = tmp("crlf.xlsx");

    let mut wb = Workbook::new();
    wb.add_sheet("Sheet1").cell_mut("A1")?.set_value("orig");
    wb.save(&path)?;

    let ir = derive(&path, DeriveOptions::default())?;
    // Convert all \n to \r\n (simulate Windows-authored IR)
    let crlf_ir = ir.replace('\n', "\r\n");

    let (_dir2, out) = tmp("out.xlsx");
    apply(
        &crlf_ir,
        &out,
        &ApplyOptions {
            force: true,
            ..Default::default()
        },
    )?;

    let wb2 = Workbook::open(&out)?;
    assert_eq!(
        wb2.sheet("Sheet1")
            .and_then(|ws| ws.cell("A1"))
            .and_then(|c| c.value()),
        Some(&CellValue::String("orig".into())),
    );
    Ok(())
}
