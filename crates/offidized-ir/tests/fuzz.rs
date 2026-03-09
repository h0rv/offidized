//! Property-based fuzz tests for the IR layer.
//!
//! Randomly generates valid xlsx workbooks with diverse content types,
//! derives them to IR, applies the IR back, and verifies roundtrip fidelity.
//! This explores the huge space of valid Excel files to find edge cases.

#![allow(
    clippy::expect_used,
    clippy::panic_in_result_fn,
    clippy::approx_constant
)]

use std::path::PathBuf;

use offidized_ir::{apply, derive, ApplyOptions, DeriveOptions};
use offidized_xlsx::{CellValue, RichTextRun, Workbook};
use proptest::prelude::*;

type R = Result<(), Box<dyn std::error::Error>>;

fn tmp(name: &str) -> (tempfile::TempDir, PathBuf) {
    let dir = tempfile::tempdir().expect("tempdir");
    let path = dir.path().join(name);
    (dir, path)
}

fn body_of(ir: &str) -> String {
    offidized_ir::IrHeader::parse(ir).expect("header").1
}

// =========================================================================
// Strategy: generate random cell values
// =========================================================================

/// Generate a random CellValue.
fn arb_cell_value() -> impl Strategy<Value = CellValue> {
    prop_oneof![
        // Numbers: wide range including edge cases
        prop_oneof![
            any::<f64>()
                .prop_filter("must be finite", |n| n.is_finite())
                .prop_map(CellValue::Number),
            Just(CellValue::Number(0.0)),
            Just(CellValue::Number(-0.0)),
            Just(CellValue::Number(1e15)),
            Just(CellValue::Number(1e-15)),
            Just(CellValue::Number(f64::MIN_POSITIVE)),
        ],
        // Booleans
        any::<bool>().prop_map(CellValue::Bool),
        // Plain strings (no special chars)
        "[a-zA-Z][a-zA-Z0-9 ]{0,30}".prop_map(CellValue::String),
        // Strings that could be confused with other types
        prop_oneof![
            // Looks like a number
            any::<i32>().prop_map(|n| CellValue::String(n.to_string())),
            // Looks like a boolean
            prop_oneof![Just("true"), Just("false"), Just("TRUE"), Just("FALSE")]
                .prop_map(|s| CellValue::String(s.to_string())),
            // Looks like a formula
            "=[A-Z]{1,3}[0-9]{1,3}".prop_map(CellValue::String),
            // Looks like an error
            prop_oneof![
                Just("#REF!"),
                Just("#NAME?"),
                Just("#VALUE!"),
                Just("#DIV/0!"),
                Just("#N/A"),
                Just("#NULL!"),
                Just("#NUM!"),
            ]
            .prop_map(|s| CellValue::String(s.to_string())),
            // Looks like <empty> marker
            Just(CellValue::String("<empty>".to_string())),
            // Has leading/trailing whitespace
            "[ \t]{1,3}[a-z]{1,10}[ \t]{0,3}".prop_map(CellValue::String),
            // Contains newlines
            "[a-z]{1,5}\n[a-z]{1,5}".prop_map(CellValue::String),
            // Contains carriage returns
            "[a-z]{1,5}\r\n[a-z]{1,5}".prop_map(CellValue::String),
            "[a-z]{1,5}\r[a-z]{1,5}".prop_map(CellValue::String),
            // Contains quotes
            "[a-z]{1,5}\"[a-z]{1,5}".prop_map(CellValue::String),
            // Starts with #
            "#[a-z]{1,10}".prop_map(CellValue::String),
            // Empty string
            Just(CellValue::String(String::new())),
        ],
        // Error values
        prop_oneof![
            Just("#REF!"),
            Just("#NAME?"),
            Just("#VALUE!"),
            Just("#DIV/0!"),
            Just("#N/A"),
            Just("#NULL!"),
            Just("#NUM!"),
        ]
        .prop_map(|s| CellValue::Error(s.to_string())),
        // Rich text (flattened to plain text in content mode)
        prop::collection::vec("[a-zA-Z ]{1,10}", 1..=4).prop_map(|parts| {
            CellValue::RichText(parts.into_iter().map(RichTextRun::new).collect())
        }),
    ]
}

/// Generate a random cell reference within reasonable bounds.
fn arb_cell_ref() -> impl Strategy<Value = String> {
    // Column: A-Z (single letter for simplicity)
    // Row: 1-200
    (0..26u8, 1..200u32).prop_map(|(col, row)| {
        let col_char = (b'A' + col) as char;
        format!("{col_char}{row}")
    })
}

/// Generate a random sheet name.
fn arb_sheet_name() -> impl Strategy<Value = String> {
    prop_oneof![
        // Simple names
        "[A-Z][a-zA-Z0-9]{1,15}",
        // Names with spaces
        "[A-Z][a-z]{1,5} [A-Z][a-z]{1,5}",
        // Names with special chars (Excel-allowed)
        "[A-Z][a-z]{1,5}[&()%-][a-z]{1,5}",
        // Unicode names
        prop_oneof![
            Just("数据表".to_string()),
            Just("Données".to_string()),
            Just("Daten 2025".to_string()),
        ],
    ]
}

/// Generate a random formula.
fn arb_formula() -> impl Strategy<Value = String> {
    prop_oneof![
        // Simple cell reference
        arb_cell_ref().prop_map(|r| r),
        // SUM/AVERAGE/COUNT
        (arb_cell_ref(), arb_cell_ref()).prop_map(|(a, b)| format!("SUM({a}:{b})")),
        (arb_cell_ref(), arb_cell_ref()).prop_map(|(a, b)| format!("AVERAGE({a}:{b})")),
        // Arithmetic
        (arb_cell_ref(), arb_cell_ref()).prop_map(|(a, b)| format!("{a}+{b}")),
        // IF formula
        (arb_cell_ref(),).prop_map(|(a,)| format!("IF({a}>0,{a},0)")),
    ]
}

// =========================================================================
// Core roundtrip property: derive → apply → derive must be idempotent
// =========================================================================

proptest! {
    #![proptest_config(ProptestConfig::with_cases(50))]

    /// Any randomly generated workbook must survive derive → apply → derive roundtrip.
    #[test]
    fn fuzz_xlsx_content_roundtrip(
        cells in prop::collection::vec(
            (arb_cell_ref(), arb_cell_value()),
            1..=50,
        ),
    ) {
        let (_dir, path) = tmp("fuzz.xlsx");

        let mut wb = Workbook::new();
        let ws = wb.add_sheet("Sheet1");

        for (cell_ref, value) in &cells {
            // Ignoring errors from invalid cell refs
            if let Ok(cell) = ws.cell_mut(cell_ref) {
                cell.set_value(value.clone());
            }
        }

        wb.save(&path).expect("save");

        let ir1 = derive(&path, DeriveOptions::default()).expect("derive");

        let (_dir2, out) = tmp("fuzz_out.xlsx");
        apply(&ir1, &out, &ApplyOptions { force: true, ..Default::default() })
            .expect("apply");

        let ir2 = derive(&out, DeriveOptions::default()).expect("re-derive");

        prop_assert_eq!(body_of(&ir1), body_of(&ir2));
    }

    /// Workbooks with multiple sheets must survive roundtrip.
    #[test]
    fn fuzz_xlsx_multi_sheet_roundtrip(
        sheet_data in prop::collection::vec(
            (
                arb_sheet_name(),
                prop::collection::vec(
                    (arb_cell_ref(), arb_cell_value()),
                    1..=20,
                ),
            ),
            1..=5,
        ),
    ) {
        let (_dir, path) = tmp("fuzz_multi.xlsx");

        let mut wb = Workbook::new();

        // Deduplicate sheet names (can't have two sheets with the same name)
        let mut seen_names = std::collections::HashSet::new();
        for (name, cells) in &sheet_data {
            let name = if seen_names.contains(name.as_str()) {
                format!("{name}_{}", seen_names.len())
            } else {
                name.clone()
            };
            seen_names.insert(name.clone());

            let ws = wb.add_sheet(&name);
            for (cell_ref, value) in cells {
                if let Ok(cell) = ws.cell_mut(cell_ref) {
                    cell.set_value(value.clone());
                }
            }
        }

        wb.save(&path).expect("save");

        let ir1 = derive(&path, DeriveOptions::default()).expect("derive");

        let (_dir2, out) = tmp("fuzz_multi_out.xlsx");
        apply(&ir1, &out, &ApplyOptions { force: true, ..Default::default() })
            .expect("apply");

        let ir2 = derive(&out, DeriveOptions::default()).expect("re-derive");

        prop_assert_eq!(body_of(&ir1), body_of(&ir2));
    }

    /// Workbooks with formulas must survive roundtrip.
    #[test]
    fn fuzz_xlsx_formulas_roundtrip(
        formulas in prop::collection::vec(
            (arb_cell_ref(), arb_formula()),
            1..=20,
        ),
        values in prop::collection::vec(
            (arb_cell_ref(), arb_cell_value()),
            0..=20,
        ),
    ) {
        let (_dir, path) = tmp("fuzz_formulas.xlsx");

        let mut wb = Workbook::new();
        let ws = wb.add_sheet("Sheet1");

        // Set values first, then formulas (formulas override values on same cell)
        for (cell_ref, value) in &values {
            if let Ok(cell) = ws.cell_mut(cell_ref) {
                cell.set_value(value.clone());
            }
        }
        for (cell_ref, formula) in &formulas {
            if let Ok(cell) = ws.cell_mut(cell_ref) {
                cell.set_formula(formula);
            }
        }

        wb.save(&path).expect("save");

        let ir1 = derive(&path, DeriveOptions::default()).expect("derive");

        let (_dir2, out) = tmp("fuzz_formulas_out.xlsx");
        apply(&ir1, &out, &ApplyOptions { force: true, ..Default::default() })
            .expect("apply");

        let ir2 = derive(&out, DeriveOptions::default()).expect("re-derive");

        prop_assert_eq!(body_of(&ir1), body_of(&ir2));
    }

    /// The IR value encoding must roundtrip: for any CellValue, derive→parse→apply→derive
    /// must produce the same value.
    #[test]
    fn fuzz_cell_value_encoding_roundtrip(
        value in arb_cell_value(),
    ) {
        let (_dir, path) = tmp("fuzz_val.xlsx");

        let mut wb = Workbook::new();
        let ws = wb.add_sheet("Sheet1");
        ws.cell_mut("A1").expect("cell").set_value(value.clone());
        wb.save(&path).expect("save");

        let ir1 = derive(&path, DeriveOptions::default()).expect("derive");
        let body1 = body_of(&ir1);

        let (_dir2, out) = tmp("fuzz_val_out.xlsx");
        apply(&ir1, &out, &ApplyOptions { force: true, ..Default::default() })
            .expect("apply");

        let ir2 = derive(&out, DeriveOptions::default()).expect("re-derive");
        let body2 = body_of(&ir2);

        // The A1 line should be identical
        let line1 = body1.lines().find(|l| l.starts_with("A1:"));
        let line2 = body2.lines().find(|l| l.starts_with("A1:"));
        prop_assert_eq!(line1, line2,
            "Cell value roundtrip mismatch for {:?}", value);
    }

    /// Randomly generated IR text must be parseable and applicable without panics.
    /// This tests the apply side's robustness against varied inputs.
    #[test]
    fn fuzz_apply_no_panic(
        cell_count in 1..30usize,
        values in prop::collection::vec(
            prop_oneof![
                // Bare numbers
                any::<f64>()
                    .prop_filter("finite", |n| n.is_finite())
                    .prop_map(|n| format!("{n}")),
                // Bare strings
                "[a-zA-Z][a-zA-Z0-9 ]{0,20}".prop_map(|s| s),
                // Quoted strings
                "[a-zA-Z0-9 ]{0,20}".prop_map(|s| format!("\"{s}\"")),
                // Formulas
                "=[A-Z][0-9]\\+[A-Z][0-9]".prop_map(|s| s),
                // Booleans
                prop_oneof![Just("true".to_string()), Just("false".to_string())],
                // Errors
                prop_oneof![
                    Just("#REF!".to_string()),
                    Just("#N/A".to_string()),
                    Just("#VALUE!".to_string()),
                ],
                // Empty marker
                Just("<empty>".to_string()),
            ],
            1..30,
        ),
    ) {
        let (_dir, path) = tmp("fuzz_apply.xlsx");

        // Create a workbook with some pre-existing cells
        let mut wb = Workbook::new();
        wb.add_sheet("Sheet1");
        wb.save(&path).expect("save");

        let base_ir = derive(&path, DeriveOptions::default()).expect("derive");
        let (header, _) = offidized_ir::IrHeader::parse(&base_ir).expect("parse header");

        // Build IR body with random cell values
        let mut body = String::new();
        body.push_str("\n=== Sheet: Sheet1 ===\n");
        for (i, value) in values.iter().take(cell_count).enumerate() {
            let col = (b'A' + (i % 26) as u8) as char;
            let row = i / 26 + 1;
            body.push_str(&format!("{col}{row}: {value}\n"));
        }

        let ir = format!("{}{body}", header.write());

        let (_dir2, out) = tmp("fuzz_apply_out.xlsx");
        // Should not panic — errors are OK, panics are not
        let _ = apply(&ir, &out, &ApplyOptions { force: true, ..Default::default() });
    }

    /// Sheet names with special characters must survive roundtrip.
    #[test]
    fn fuzz_sheet_names_roundtrip(
        names in prop::collection::vec(arb_sheet_name(), 1..=8),
    ) {
        let (_dir, path) = tmp("fuzz_names.xlsx");

        let mut wb = Workbook::new();
        let mut seen = std::collections::HashSet::new();
        for name in &names {
            let name = if seen.contains(name.as_str()) {
                format!("{name}_{}", seen.len())
            } else {
                name.clone()
            };
            seen.insert(name.clone());

            if let Ok(cell) = wb.add_sheet(&name).cell_mut("A1") {
                cell.set_value("data");
            }
        }

        wb.save(&path).expect("save");

        let ir1 = derive(&path, DeriveOptions::default()).expect("derive");

        let (_dir2, out) = tmp("fuzz_names_out.xlsx");
        apply(&ir1, &out, &ApplyOptions { force: true, ..Default::default() })
            .expect("apply");

        let ir2 = derive(&out, DeriveOptions::default()).expect("re-derive");
        prop_assert_eq!(body_of(&ir1), body_of(&ir2));
    }
}

// =========================================================================
// Deterministic fuzz-like tests for specific categories
// =========================================================================

/// Test all possible CellValue::String contents that could be confused
/// with other types, in bulk.
#[test]
fn bulk_confusable_strings() -> R {
    let confusable_strings = vec![
        // Numbers
        "0",
        "1",
        "-1",
        "3.14",
        "-0",
        "1e10",
        "1E10",
        "1e-5",
        "NaN",
        "Infinity",
        "-Infinity",
        "inf",
        "-inf",
        "+1",
        "+3.14",
        "0.0",
        ".5",
        "1.",
        "00",
        "007",
        // Booleans
        "true",
        "false",
        "TRUE",
        "FALSE",
        "True",
        "False",
        // Formulas
        "=A1",
        "=SUM(A1:B2)",
        "=",
        "==",
        "=TRUE",
        // Errors
        "#REF!",
        "#NAME?",
        "#VALUE!",
        "#DIV/0!",
        "#N/A",
        "#NULL!",
        "#NUM!",
        "#ref!",
        "#CUSTOM!",
        "#",
        "##",
        // Special markers
        "<empty>",
        "<EMPTY>",
        // Whitespace
        " ",
        "  ",
        "\t",
        " hello",
        "hello ",
        " hello ",
        // Newlines/CR
        "a\nb",
        "a\rb",
        "a\r\nb",
        "\n",
        "\r",
        "\r\n",
        // Quotes
        "\"",
        "\"\"",
        "a\"b",
        "\"hello\"",
        // Empty
        "",
        // Mixed
        "=true",
        "#true",
        "\"42\"",
    ];

    let (_dir, path) = tmp("confusable.xlsx");

    let mut wb = Workbook::new();
    let ws = wb.add_sheet("Sheet1");

    for (i, s) in confusable_strings.iter().enumerate() {
        let cell_ref = format!("A{}", i + 1);
        ws.cell_mut(&cell_ref)?
            .set_value(CellValue::String((*s).to_string()));
    }
    wb.save(&path)?;

    let ir1 = derive(&path, DeriveOptions::default())?;

    let (_dir2, out) = tmp("confusable_out.xlsx");
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

    // Also verify each cell came back as String (not Number, Bool, etc.)
    let wb2 = Workbook::open(&out)?;
    let ws2 = wb2.sheet("Sheet1").ok_or("missing")?;

    for (i, expected) in confusable_strings.iter().enumerate() {
        let cell_ref = format!("A{}", i + 1);
        let actual = ws2.cell(&cell_ref).and_then(|c| c.value());
        match actual {
            Some(CellValue::String(s)) => {
                assert_eq!(
                    s.as_str(),
                    *expected,
                    "String roundtrip failed for cell {cell_ref}: expected {:?}, got {:?}",
                    expected,
                    s
                );
            }
            other => {
                panic!(
                    "Cell {cell_ref} should be String({:?}), got {:?}",
                    expected, other,
                );
            }
        }
    }

    Ok(())
}

/// Test number formatting/parsing edge cases in bulk.
#[test]
fn bulk_number_roundtrip() -> R {
    let numbers: Vec<f64> = vec![
        0.0,
        -0.0,
        1.0,
        -1.0,
        0.1,
        0.01,
        0.001,
        0.1 + 0.2, // 0.30000000000000004
        1.0 / 3.0, // 0.333...
        std::f64::consts::PI,
        std::f64::consts::E,
        42.0,
        -42.0,
        999999999999999.0, // 15 digits (fits i64)
        -999999999999999.0,
        1e10,
        1e-10,
        1e100,
        1e-100,
        f64::MAX,
        f64::MIN,
        f64::MIN_POSITIVE,
        f64::EPSILON,
        1.7976931348623157e308, // Near MAX
        5e-324,                 // MIN_POSITIVE subnormal
    ];

    let (_dir, path) = tmp("numbers.xlsx");
    let mut wb = Workbook::new();
    let ws = wb.add_sheet("Sheet1");

    for (i, &n) in numbers.iter().enumerate() {
        let cell_ref = format!("A{}", i + 1);
        ws.cell_mut(&cell_ref)?.set_value(CellValue::Number(n));
    }
    wb.save(&path)?;

    let ir1 = derive(&path, DeriveOptions::default())?;

    let (_dir2, out) = tmp("numbers_out.xlsx");
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

    // Verify values
    let wb2 = Workbook::open(&out)?;
    let ws2 = wb2.sheet("Sheet1").ok_or("missing")?;

    for (i, &expected) in numbers.iter().enumerate() {
        let cell_ref = format!("A{}", i + 1);
        match ws2.cell(&cell_ref).and_then(|c| c.value()) {
            Some(CellValue::Number(actual)) => {
                // Both zero: OK (negative zero collapses to zero)
                if expected == 0.0 && *actual == 0.0 {
                    continue;
                }
                assert!(
                    (actual - expected).abs() < f64::EPSILON * expected.abs().max(1.0),
                    "Number mismatch at {cell_ref}: expected {expected}, got {actual}",
                );
            }
            other => {
                panic!("Cell {cell_ref} should be Number({expected}), got {other:?}",);
            }
        }
    }

    Ok(())
}

/// Mixed content: values, formulas, errors, booleans, rich text on the same sheet.
#[test]
fn mixed_content_stress() -> R {
    let (_dir, path) = tmp("mixed.xlsx");
    let mut wb = Workbook::new();
    let ws = wb.add_sheet("Mixed");

    // Values of every type
    ws.cell_mut("A1")?.set_value("plain string");
    ws.cell_mut("A2")?.set_value(CellValue::String("42".into()));
    ws.cell_mut("A3")?
        .set_value(CellValue::String("true".into()));
    ws.cell_mut("A4")?
        .set_value(CellValue::String("=SUM".into()));
    ws.cell_mut("A5")?
        .set_value(CellValue::String("#N/A".into()));
    ws.cell_mut("A6")?
        .set_value(CellValue::String("<empty>".into()));
    ws.cell_mut("A7")?
        .set_value(CellValue::String("line1\nline2".into()));
    ws.cell_mut("A8")?
        .set_value(CellValue::String("has \"quotes\"".into()));
    ws.cell_mut("A9")?.set_value(CellValue::String("".into()));
    ws.cell_mut("A10")?
        .set_value(CellValue::String("  spaces  ".into()));
    ws.cell_mut("A11")?.set_value(CellValue::Number(42.0));
    ws.cell_mut("A12")?.set_value(CellValue::Number(3.14));
    ws.cell_mut("A13")?.set_value(CellValue::Number(-0.0));
    ws.cell_mut("A14")?.set_value(CellValue::Bool(true));
    ws.cell_mut("A15")?.set_value(CellValue::Bool(false));
    ws.cell_mut("A16")?
        .set_value(CellValue::Error("#REF!".into()));
    ws.cell_mut("A17")?
        .set_value(CellValue::Error("#DIV/0!".into()));
    ws.cell_mut("A18")?.set_value(CellValue::RichText(vec![
        RichTextRun::new("bold "),
        RichTextRun::new("normal"),
    ]));
    ws.cell_mut("B1")?.set_formula("SUM(A11:A13)");
    ws.cell_mut("B2")?.set_formula("IF(A14,\"yes\",\"no\")");

    wb.save(&path)?;

    let ir1 = derive(&path, DeriveOptions::default())?;

    let (_dir2, out) = tmp("mixed_out.xlsx");
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

/// Fuzz the apply side: generate random modifications to an existing workbook.
#[test]
fn fuzz_apply_modifications() -> R {
    let (_dir, path) = tmp("fuzz_mod.xlsx");

    // Create a workbook with initial data
    let mut wb = Workbook::new();
    let ws = wb.add_sheet("Data");
    for row in 1..=10 {
        for col in b'A'..=b'E' {
            let cell_ref = format!("{}{row}", col as char);
            ws.cell_mut(&cell_ref)?.set_value(format!("r{row}c{col}"));
        }
    }
    wb.save(&path)?;

    let base_ir = derive(&path, DeriveOptions::default())?;
    let (header, _) = offidized_ir::IrHeader::parse(&base_ir)?;

    // Apply various modifications
    let modifications = vec![
        ("A1", "modified"),
        ("A2", "42"),
        ("A3", "true"),
        ("A4", "=B4+C4"),
        ("A5", "<empty>"),
        ("A6", "\"quoted number 42\""),
        ("A7", "#REF!"),
        ("A8", "\"string with \\nnewline\""),
        ("F1", "new column"), // new cell
        ("A11", "new row"),   // new cell
    ];

    let mut body = String::new();
    body.push_str("\n=== Sheet: Data ===\n");
    for (cell_ref, value) in &modifications {
        body.push_str(&format!("{cell_ref}: {value}\n"));
    }

    let ir = format!("{}{body}", header.write());

    let (_dir2, out) = tmp("fuzz_mod_out.xlsx");
    let result = apply(
        &ir,
        &out,
        &ApplyOptions {
            force: true,
            ..Default::default()
        },
    )?;

    // Verify the results
    assert!(result.cells_updated > 0);
    assert!(result.cells_created > 0);
    assert!(result.cells_cleared > 0);

    // Verify untouched cells survived
    let wb2 = Workbook::open(&out)?;
    let ws2 = wb2.sheet("Data").ok_or("missing")?;

    // B1 should be untouched
    assert_eq!(
        ws2.cell("B1").and_then(|c| c.value()),
        Some(&CellValue::String("r1c66".into())),
    );

    // A1 should be modified
    assert_eq!(
        ws2.cell("A1").and_then(|c| c.value()),
        Some(&CellValue::String("modified".into())),
    );

    // A5 should be cleared
    let a5 = ws2.cell("A5").and_then(|c| c.value());
    assert!(a5.is_none() || matches!(a5, Some(CellValue::Blank)));

    Ok(())
}
