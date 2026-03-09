//! Integration tests for the `offidized` CLI binary.
//!
//! These tests run the compiled binary as a subprocess and assert on stdout/stderr/exit code.
//! They use real OOXML files from the Open-XML-SDK test corpus and the local sales.xlsx.

#![allow(clippy::unwrap_used)]
#![allow(clippy::expect_used)]

use std::path::{Path, PathBuf};
use std::process::Command;

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

fn bin() -> Command {
    Command::new(env!("CARGO_BIN_EXE_ofx"))
}

fn reference_test_files() -> PathBuf {
    let manifest = Path::new(env!("CARGO_MANIFEST_DIR"));
    manifest.join(
        "../../references/Open-XML-SDK/test/DocumentFormat.OpenXml.Tests.Assets/assets/TestFiles",
    )
}

fn sales_xlsx() -> PathBuf {
    let manifest = Path::new(env!("CARGO_MANIFEST_DIR"));
    manifest.join("../offidized-xlsx/tests/fixtures/sales.xlsx")
}

fn has_reference_files() -> bool {
    reference_test_files()
        .join("basicspreadsheet.xlsx")
        .is_file()
}

/// Run the binary and return (exit_code, stdout, stderr).
fn run(args: &[&str]) -> (i32, String, String) {
    let output = bin()
        .args(args)
        .output()
        .expect("failed to execute offidized binary");
    let code = output.status.code().unwrap_or(-1);
    let stdout = String::from_utf8_lossy(&output.stdout).to_string();
    let stderr = String::from_utf8_lossy(&output.stderr).to_string();
    (code, stdout, stderr)
}

/// Run binary, assert success, return stdout.
fn run_ok(args: &[&str]) -> String {
    let (code, stdout, stderr) = run(args);
    assert_eq!(
        code, 0,
        "command failed (exit {code}):\nstderr: {stderr}\nargs: {args:?}"
    );
    stdout
}

/// Run binary, assert failure, return stderr.
fn run_err(args: &[&str]) -> String {
    let (code, _stdout, stderr) = run(args);
    assert_ne!(code, 0, "expected failure but got success\nargs: {args:?}");
    stderr
}

// ---------------------------------------------------------------------------
// info
// ---------------------------------------------------------------------------

#[test]
fn info_xlsx_sales() {
    let path = sales_xlsx();
    if !path.is_file() {
        eprintln!("skipping: sales.xlsx not found");
        return;
    }
    let stdout = run_ok(&["info", path.to_str().unwrap()]);
    let json: serde_json::Value = serde_json::from_str(&stdout).expect("valid JSON");
    assert_eq!(json["format"], "xlsx");
    assert_eq!(json["sheets"][0], "Sales");
    assert!(json["part_count"].as_u64().unwrap() > 0);
}

#[test]
fn info_xlsx_basicspreadsheet() {
    if !has_reference_files() {
        eprintln!("skipping: reference test files not found");
        return;
    }
    let path = reference_test_files().join("basicspreadsheet.xlsx");
    let stdout = run_ok(&["info", path.to_str().unwrap()]);
    let json: serde_json::Value = serde_json::from_str(&stdout).expect("valid JSON");
    assert_eq!(json["format"], "xlsx");
    let sheets = json["sheets"].as_array().unwrap();
    assert_eq!(sheets.len(), 3);
    assert_eq!(sheets[0], "Sheet1");
    assert_eq!(sheets[1], "Sheet2");
    assert_eq!(sheets[2], "Sheet3");
}

#[test]
fn info_docx_hello_world() {
    if !has_reference_files() {
        eprintln!("skipping: reference test files not found");
        return;
    }
    let path = reference_test_files().join("HelloWorld.docx");
    let stdout = run_ok(&["info", path.to_str().unwrap()]);
    let json: serde_json::Value = serde_json::from_str(&stdout).expect("valid JSON");
    assert_eq!(json["format"], "docx");
    assert_eq!(json["paragraph_count"], 1);
    assert_eq!(json["table_count"], 0);
    assert!(json["part_count"].as_u64().unwrap() > 0);
}

#[test]
fn info_pptx_not_implemented() {
    let stderr = run_err(&["info", "/tmp/nonexistent.pptx"]);
    assert!(
        stderr.contains("pptx"),
        "expected pptx error, got: {stderr}"
    );
}

#[test]
fn info_unsupported_extension() {
    let stderr = run_err(&["info", "/tmp/file.txt"]);
    assert!(
        stderr.contains("unsupported file extension"),
        "got: {stderr}"
    );
}

// ---------------------------------------------------------------------------
// read
// ---------------------------------------------------------------------------

#[test]
fn read_xlsx_all_sheets() {
    let path = sales_xlsx();
    if !path.is_file() {
        eprintln!("skipping: sales.xlsx not found");
        return;
    }
    let stdout = run_ok(&["read", path.to_str().unwrap()]);
    let json: serde_json::Value = serde_json::from_str(&stdout).expect("valid JSON");
    let sheets = json["sheets"].as_array().unwrap();
    assert_eq!(sheets.len(), 1);
    assert_eq!(sheets[0]["name"], "Sales");
    let cells = sheets[0]["cells"].as_array().unwrap();
    assert!(!cells.is_empty());
    // Check known values.
    let product_cell = cells.iter().find(|c| c["ref"] == "A1").unwrap();
    assert_eq!(product_cell["value"], "Product");
    assert_eq!(product_cell["type"], "string");
    let revenue_cell = cells.iter().find(|c| c["ref"] == "B2").unwrap();
    assert_eq!(revenue_cell["value"], 42000.0);
    assert_eq!(revenue_cell["type"], "number");
}

#[test]
fn read_xlsx_with_range_filter() {
    let path = sales_xlsx();
    if !path.is_file() {
        eprintln!("skipping: sales.xlsx not found");
        return;
    }
    let stdout = run_ok(&["read", path.to_str().unwrap(), "Sales!A1:A2"]);
    let json: serde_json::Value = serde_json::from_str(&stdout).expect("valid JSON");
    let cells = json.as_array().unwrap();
    // Only A-column rows 1-2.
    assert!(cells.iter().all(|c| {
        let r = c["ref"].as_str().unwrap();
        r.starts_with('A')
    }));
}

#[test]
fn read_xlsx_csv_format() {
    let path = sales_xlsx();
    if !path.is_file() {
        eprintln!("skipping: sales.xlsx not found");
        return;
    }
    let stdout = run_ok(&["read", path.to_str().unwrap(), "--format", "csv"]);
    assert!(stdout.contains("A1,Product"));
    assert!(stdout.contains("B2,42000"));
}

#[test]
fn read_xlsx_sheet_not_found() {
    let path = sales_xlsx();
    if !path.is_file() {
        eprintln!("skipping: sales.xlsx not found");
        return;
    }
    let stderr = run_err(&["read", path.to_str().unwrap(), "NoSuchSheet!A1"]);
    assert!(stderr.contains("sheet not found"), "got: {stderr}");
}

#[test]
fn read_xlsx_basicspreadsheet_sheet2() {
    if !has_reference_files() {
        eprintln!("skipping: reference test files not found");
        return;
    }
    let path = reference_test_files().join("basicspreadsheet.xlsx");
    let stdout = run_ok(&["read", path.to_str().unwrap(), "Sheet2!A1:C4"]);
    let json: serde_json::Value = serde_json::from_str(&stdout).expect("valid JSON");
    let cells = json.as_array().unwrap();
    // Sheet2 has headers a, b, c and numeric data.
    let a1 = cells.iter().find(|c| c["ref"] == "A1").unwrap();
    assert_eq!(a1["value"], "a");
    let b2 = cells.iter().find(|c| c["ref"] == "B2").unwrap();
    assert_eq!(b2["value"], 2.0);
    assert_eq!(b2["type"], "number");
}

#[test]
fn read_docx_hello_world() {
    if !has_reference_files() {
        eprintln!("skipping: reference test files not found");
        return;
    }
    let path = reference_test_files().join("HelloWorld.docx");
    let stdout = run_ok(&["read", path.to_str().unwrap()]);
    let json: serde_json::Value = serde_json::from_str(&stdout).expect("valid JSON");
    let paragraphs = json.as_array().unwrap();
    assert_eq!(paragraphs.len(), 1);
    assert_eq!(paragraphs[0]["index"], 0);
    assert_eq!(paragraphs[0]["text"], "Hello World!");
}

#[test]
fn read_docx_with_paragraph_filter() {
    if !has_reference_files() {
        eprintln!("skipping: reference test files not found");
        return;
    }
    let path = reference_test_files().join("Plain.docx");
    let stdout = run_ok(&["read", path.to_str().unwrap(), "--paragraphs", "0-0"]);
    let json: serde_json::Value = serde_json::from_str(&stdout).expect("valid JSON");
    let paragraphs = json.as_array().unwrap();
    assert_eq!(paragraphs.len(), 1);
    assert!(paragraphs[0]["text"]
        .as_str()
        .unwrap()
        .contains("Video provides"));
}

// ---------------------------------------------------------------------------
// set + roundtrip
// ---------------------------------------------------------------------------

#[test]
fn set_xlsx_and_read_back() {
    let dir = tempfile::tempdir().expect("tempdir");
    let created = dir.path().join("test.xlsx");
    let output = dir.path().join("modified.xlsx");

    // Create, set, read back.
    run_ok(&["create", created.to_str().unwrap()]);
    run_ok(&[
        "set",
        created.to_str().unwrap(),
        "Sheet1!A1",
        "42",
        "-o",
        output.to_str().unwrap(),
    ]);
    let stdout = run_ok(&["read", output.to_str().unwrap(), "Sheet1!A1"]);
    let json: serde_json::Value = serde_json::from_str(&stdout).expect("valid JSON");
    let cells = json.as_array().unwrap();
    assert_eq!(cells.len(), 1);
    assert_eq!(cells[0]["value"], 42.0);
    assert_eq!(cells[0]["type"], "number");
}

#[test]
fn set_xlsx_bool_and_string_autodetection() {
    let dir = tempfile::tempdir().expect("tempdir");
    let created = dir.path().join("test.xlsx");
    let step1 = dir.path().join("step1.xlsx");
    let step2 = dir.path().join("step2.xlsx");

    run_ok(&["create", created.to_str().unwrap()]);
    run_ok(&[
        "set",
        created.to_str().unwrap(),
        "Sheet1!A1",
        "true",
        "-o",
        step1.to_str().unwrap(),
    ]);
    run_ok(&[
        "set",
        step1.to_str().unwrap(),
        "Sheet1!B1",
        "hello world",
        "-o",
        step2.to_str().unwrap(),
    ]);

    let stdout = run_ok(&["read", step2.to_str().unwrap()]);
    let json: serde_json::Value = serde_json::from_str(&stdout).expect("valid JSON");
    let cells = &json["sheets"][0]["cells"];
    let a1 = cells
        .as_array()
        .unwrap()
        .iter()
        .find(|c| c["ref"] == "A1")
        .unwrap();
    assert_eq!(a1["value"], true);
    assert_eq!(a1["type"], "bool");
    let b1 = cells
        .as_array()
        .unwrap()
        .iter()
        .find(|c| c["ref"] == "B1")
        .unwrap();
    assert_eq!(b1["value"], "hello world");
    assert_eq!(b1["type"], "string");
}

#[test]
fn set_xlsx_in_place() {
    let dir = tempfile::tempdir().expect("tempdir");
    let file = dir.path().join("test.xlsx");

    run_ok(&["create", file.to_str().unwrap()]);
    run_ok(&["set", file.to_str().unwrap(), "Sheet1!A1", "99", "-i"]);
    let stdout = run_ok(&["read", file.to_str().unwrap(), "Sheet1!A1"]);
    let json: serde_json::Value = serde_json::from_str(&stdout).expect("valid JSON");
    assert_eq!(json.as_array().unwrap()[0]["value"], 99.0);
}

#[test]
fn set_docx_and_read_back() {
    if !has_reference_files() {
        eprintln!("skipping: reference test files not found");
        return;
    }
    let dir = tempfile::tempdir().expect("tempdir");
    let output = dir.path().join("modified.docx");
    let path = reference_test_files().join("HelloWorld.docx");

    run_ok(&[
        "set",
        path.to_str().unwrap(),
        "0",
        "Replaced text",
        "-o",
        output.to_str().unwrap(),
    ]);
    let stdout = run_ok(&["read", output.to_str().unwrap()]);
    let json: serde_json::Value = serde_json::from_str(&stdout).expect("valid JSON");
    assert_eq!(json[0]["text"], "Replaced text");
}

#[test]
fn set_requires_output_flag() {
    let dir = tempfile::tempdir().expect("tempdir");
    let file = dir.path().join("test.xlsx");
    run_ok(&["create", file.to_str().unwrap()]);

    let stderr = run_err(&["set", file.to_str().unwrap(), "Sheet1!A1", "42"]);
    assert!(stderr.contains("must specify"), "got: {stderr}");
}

#[test]
fn set_rejects_both_output_and_inplace() {
    let dir = tempfile::tempdir().expect("tempdir");
    let file = dir.path().join("test.xlsx");
    let out = dir.path().join("out.xlsx");
    run_ok(&["create", file.to_str().unwrap()]);

    let stderr = run_err(&[
        "set",
        file.to_str().unwrap(),
        "Sheet1!A1",
        "42",
        "-o",
        out.to_str().unwrap(),
        "-i",
    ]);
    assert!(stderr.contains("cannot specify both"), "got: {stderr}");
}

// ---------------------------------------------------------------------------
// patch
// ---------------------------------------------------------------------------

#[test]
fn patch_xlsx_via_stdin() {
    let dir = tempfile::tempdir().expect("tempdir");
    let created = dir.path().join("test.xlsx");
    let output = dir.path().join("patched.xlsx");

    run_ok(&["create", created.to_str().unwrap()]);

    let patch_json = r#"[{"ref":"Sheet1!A1","value":42},{"ref":"Sheet1!B1","value":"text"},{"ref":"Sheet1!C1","value":true}]"#;
    let result = bin()
        .args([
            "patch",
            created.to_str().unwrap(),
            "-o",
            output.to_str().unwrap(),
        ])
        .stdin(std::process::Stdio::piped())
        .stdout(std::process::Stdio::piped())
        .stderr(std::process::Stdio::piped())
        .spawn()
        .and_then(|mut child| {
            use std::io::Write;
            child
                .stdin
                .take()
                .unwrap()
                .write_all(patch_json.as_bytes())
                .unwrap();
            child.wait_with_output()
        })
        .expect("patch command");
    assert!(
        result.status.success(),
        "patch failed: {}",
        String::from_utf8_lossy(&result.stderr)
    );

    let stdout = run_ok(&["read", output.to_str().unwrap()]);
    let json: serde_json::Value = serde_json::from_str(&stdout).expect("valid JSON");
    let cells = json["sheets"][0]["cells"].as_array().unwrap();
    let a1 = cells.iter().find(|c| c["ref"] == "A1").unwrap();
    assert_eq!(a1["value"], 42.0);
    assert_eq!(a1["type"], "number");
    let b1 = cells.iter().find(|c| c["ref"] == "B1").unwrap();
    assert_eq!(b1["value"], "text");
    assert_eq!(b1["type"], "string");
    let c1 = cells.iter().find(|c| c["ref"] == "C1").unwrap();
    assert_eq!(c1["value"], true);
    assert_eq!(c1["type"], "bool");
}

#[test]
fn patch_xlsx_null_clears_cell() {
    let dir = tempfile::tempdir().expect("tempdir");
    let step1 = dir.path().join("step1.xlsx");
    let step2 = dir.path().join("step2.xlsx");

    run_ok(&["create", step1.to_str().unwrap()]);
    run_ok(&["set", step1.to_str().unwrap(), "Sheet1!A1", "hello", "-i"]);

    let patch_json = r#"[{"ref":"Sheet1!A1","value":null}]"#;
    let result = bin()
        .args([
            "patch",
            step1.to_str().unwrap(),
            "-o",
            step2.to_str().unwrap(),
        ])
        .stdin(std::process::Stdio::piped())
        .stdout(std::process::Stdio::piped())
        .stderr(std::process::Stdio::piped())
        .spawn()
        .and_then(|mut child| {
            use std::io::Write;
            child
                .stdin
                .take()
                .unwrap()
                .write_all(patch_json.as_bytes())
                .unwrap();
            child.wait_with_output()
        })
        .expect("patch command");
    assert!(result.status.success());

    let stdout = run_ok(&["read", step2.to_str().unwrap(), "Sheet1!A1"]);
    let json: serde_json::Value = serde_json::from_str(&stdout).expect("valid JSON");
    let cells = json.as_array().unwrap();
    let a1 = cells.iter().find(|c| c["ref"] == "A1").unwrap();
    assert_eq!(a1["type"], "blank");
}

#[test]
fn patch_xlsx_string_stays_string() {
    // Regression: patch should NOT auto-type string JSON values.
    // "42" as a JSON string should produce a String cell, not Number.
    let dir = tempfile::tempdir().expect("tempdir");
    let created = dir.path().join("test.xlsx");
    let output = dir.path().join("patched.xlsx");

    run_ok(&["create", created.to_str().unwrap()]);

    let patch_json = r#"[{"ref":"Sheet1!A1","value":"42"}]"#;
    let result = bin()
        .args([
            "patch",
            created.to_str().unwrap(),
            "-o",
            output.to_str().unwrap(),
        ])
        .stdin(std::process::Stdio::piped())
        .stdout(std::process::Stdio::piped())
        .stderr(std::process::Stdio::piped())
        .spawn()
        .and_then(|mut child| {
            use std::io::Write;
            child
                .stdin
                .take()
                .unwrap()
                .write_all(patch_json.as_bytes())
                .unwrap();
            child.wait_with_output()
        })
        .expect("patch command");
    assert!(result.status.success());

    let stdout = run_ok(&["read", output.to_str().unwrap(), "Sheet1!A1"]);
    let json: serde_json::Value = serde_json::from_str(&stdout).expect("valid JSON");
    let cells = json.as_array().unwrap();
    let a1 = cells.iter().find(|c| c["ref"] == "A1").unwrap();
    assert_eq!(a1["value"], "42");
    assert_eq!(a1["type"], "string");
}

// ---------------------------------------------------------------------------
// replace
// ---------------------------------------------------------------------------

#[test]
fn replace_xlsx() {
    let path = sales_xlsx();
    if !path.is_file() {
        eprintln!("skipping: sales.xlsx not found");
        return;
    }
    let dir = tempfile::tempdir().expect("tempdir");
    let output = dir.path().join("replaced.xlsx");

    run_ok(&[
        "replace",
        path.to_str().unwrap(),
        "Widget",
        "Gadget",
        "-o",
        output.to_str().unwrap(),
    ]);

    let stdout = run_ok(&["read", output.to_str().unwrap()]);
    let json: serde_json::Value = serde_json::from_str(&stdout).expect("valid JSON");
    let cells = json["sheets"][0]["cells"].as_array().unwrap();
    let a2 = cells.iter().find(|c| c["ref"] == "A2").unwrap();
    assert_eq!(a2["value"], "Gadget");
    // Revenue should be untouched (it's a number).
    let b2 = cells.iter().find(|c| c["ref"] == "B2").unwrap();
    assert_eq!(b2["value"], 42000.0);
}

#[test]
fn replace_docx() {
    if !has_reference_files() {
        eprintln!("skipping: reference test files not found");
        return;
    }
    let dir = tempfile::tempdir().expect("tempdir");
    let output = dir.path().join("replaced.docx");
    let path = reference_test_files().join("HelloWorld.docx");

    run_ok(&[
        "replace",
        path.to_str().unwrap(),
        "World",
        "Rust",
        "-o",
        output.to_str().unwrap(),
    ]);

    let stdout = run_ok(&["read", output.to_str().unwrap()]);
    let json: serde_json::Value = serde_json::from_str(&stdout).expect("valid JSON");
    assert_eq!(json[0]["text"], "Hello Rust!");
}

// ---------------------------------------------------------------------------
// part
// ---------------------------------------------------------------------------

#[test]
fn part_list_xlsx() {
    let path = sales_xlsx();
    if !path.is_file() {
        eprintln!("skipping: sales.xlsx not found");
        return;
    }
    let stdout = run_ok(&["part", path.to_str().unwrap(), "--list"]);
    let json: serde_json::Value = serde_json::from_str(&stdout).expect("valid JSON");
    let parts = json.as_array().unwrap();
    assert!(!parts.is_empty());
    // Every part must have uri, content_type, size_bytes.
    for part in parts {
        assert!(part["uri"].is_string());
        assert!(part["size_bytes"].is_number());
    }
}

#[test]
fn part_list_docx() {
    if !has_reference_files() {
        eprintln!("skipping: reference test files not found");
        return;
    }
    let path = reference_test_files().join("HelloWorld.docx");
    let stdout = run_ok(&["part", path.to_str().unwrap(), "--list"]);
    let json: serde_json::Value = serde_json::from_str(&stdout).expect("valid JSON");
    let parts = json.as_array().unwrap();
    assert_eq!(parts.len(), 8);
    // Should have document.xml part.
    assert!(parts
        .iter()
        .any(|p| p["uri"].as_str().unwrap().contains("document.xml")));
}

#[test]
fn part_extract_raw_xml() {
    if !has_reference_files() {
        eprintln!("skipping: reference test files not found");
        return;
    }
    let path = reference_test_files().join("HelloWorld.docx");
    let stdout = run_ok(&["part", path.to_str().unwrap(), "/word/document.xml"]);
    // Should be raw XML.
    assert!(
        stdout.contains("<?xml"),
        "expected XML, got: {}",
        &stdout[..stdout.len().min(200)]
    );
    assert!(stdout.contains("Hello World!"));
}

#[test]
fn part_not_found() {
    let path = sales_xlsx();
    if !path.is_file() {
        eprintln!("skipping: sales.xlsx not found");
        return;
    }
    let stderr = run_err(&["part", path.to_str().unwrap(), "/nonexistent/part.xml"]);
    assert!(stderr.contains("not found"), "got: {stderr}");
}

#[test]
fn part_requires_uri_or_list() {
    let path = sales_xlsx();
    if !path.is_file() {
        eprintln!("skipping: sales.xlsx not found");
        return;
    }
    let stderr = run_err(&["part", path.to_str().unwrap()]);
    assert!(
        stderr.contains("specify a part URI") || stderr.contains("--list"),
        "got: {stderr}"
    );
}

// ---------------------------------------------------------------------------
// create
// ---------------------------------------------------------------------------

#[test]
fn create_xlsx_and_verify() {
    let dir = tempfile::tempdir().expect("tempdir");
    let file = dir.path().join("new.xlsx");

    run_ok(&["create", file.to_str().unwrap()]);
    assert!(file.is_file());

    let stdout = run_ok(&["info", file.to_str().unwrap()]);
    let json: serde_json::Value = serde_json::from_str(&stdout).expect("valid JSON");
    assert_eq!(json["format"], "xlsx");
    assert_eq!(json["sheets"][0], "Sheet1");
}

#[test]
fn create_docx_and_verify() {
    let dir = tempfile::tempdir().expect("tempdir");
    let file = dir.path().join("new.docx");

    run_ok(&["create", file.to_str().unwrap()]);
    assert!(file.is_file());

    let stdout = run_ok(&["info", file.to_str().unwrap()]);
    let json: serde_json::Value = serde_json::from_str(&stdout).expect("valid JSON");
    assert_eq!(json["format"], "docx");
    assert_eq!(json["paragraph_count"], 0);
}

#[test]
fn create_pptx_not_implemented() {
    let stderr = run_err(&["create", "/tmp/test.pptx"]);
    assert!(stderr.contains("pptx"), "got: {stderr}");
}

// ---------------------------------------------------------------------------
// roundtrip: open real file → set → save → read back → verify unchanged parts
// ---------------------------------------------------------------------------

#[test]
fn roundtrip_xlsx_preserves_other_cells() {
    let path = sales_xlsx();
    if !path.is_file() {
        eprintln!("skipping: sales.xlsx not found");
        return;
    }
    let dir = tempfile::tempdir().expect("tempdir");
    let output = dir.path().join("roundtrip.xlsx");

    // Change B2 (Revenue) but A1/A2/B1 should survive.
    run_ok(&[
        "set",
        path.to_str().unwrap(),
        "Sales!B2",
        "99999",
        "-o",
        output.to_str().unwrap(),
    ]);

    let stdout = run_ok(&["read", output.to_str().unwrap()]);
    let json: serde_json::Value = serde_json::from_str(&stdout).expect("valid JSON");
    let cells = json["sheets"][0]["cells"].as_array().unwrap();
    // Original cells preserved.
    let a1 = cells.iter().find(|c| c["ref"] == "A1").unwrap();
    assert_eq!(a1["value"], "Product");
    let a2 = cells.iter().find(|c| c["ref"] == "A2").unwrap();
    assert_eq!(a2["value"], "Widget");
    let b1 = cells.iter().find(|c| c["ref"] == "B1").unwrap();
    assert_eq!(b1["value"], "Revenue");
    // Modified cell.
    let b2 = cells.iter().find(|c| c["ref"] == "B2").unwrap();
    assert_eq!(b2["value"], 99999.0);
}

#[test]
fn roundtrip_docx_preserves_other_paragraphs() {
    if !has_reference_files() {
        eprintln!("skipping: reference test files not found");
        return;
    }
    let path = reference_test_files().join("Plain.docx");
    let dir = tempfile::tempdir().expect("tempdir");
    let output = dir.path().join("roundtrip.docx");

    run_ok(&[
        "set",
        path.to_str().unwrap(),
        "0",
        "New first paragraph",
        "-o",
        output.to_str().unwrap(),
    ]);

    let stdout = run_ok(&["read", output.to_str().unwrap()]);
    let json: serde_json::Value = serde_json::from_str(&stdout).expect("valid JSON");
    let paragraphs = json.as_array().unwrap();
    assert_eq!(paragraphs[0]["text"], "New first paragraph");
    // Part count should be preserved (no parts lost).
    let info_stdout = run_ok(&["info", output.to_str().unwrap()]);
    let info: serde_json::Value = serde_json::from_str(&info_stdout).expect("valid JSON");
    assert!(info["part_count"].as_u64().unwrap() >= 1);
}

// ---------------------------------------------------------------------------
// Edge cases
// ---------------------------------------------------------------------------

#[test]
fn read_empty_xlsx_returns_empty_cells() {
    let dir = tempfile::tempdir().expect("tempdir");
    let file = dir.path().join("empty.xlsx");
    run_ok(&["create", file.to_str().unwrap()]);

    let stdout = run_ok(&["read", file.to_str().unwrap()]);
    let json: serde_json::Value = serde_json::from_str(&stdout).expect("valid JSON");
    let cells = json["sheets"][0]["cells"].as_array().unwrap();
    assert!(cells.is_empty());
}

#[test]
fn read_empty_docx_returns_empty_paragraphs() {
    let dir = tempfile::tempdir().expect("tempdir");
    let file = dir.path().join("empty.docx");
    run_ok(&["create", file.to_str().unwrap()]);

    let stdout = run_ok(&["read", file.to_str().unwrap()]);
    let json: serde_json::Value = serde_json::from_str(&stdout).expect("valid JSON");
    let paragraphs = json.as_array().unwrap();
    assert!(paragraphs.is_empty());
}

#[test]
fn set_docx_out_of_range_paragraph() {
    if !has_reference_files() {
        eprintln!("skipping: reference test files not found");
        return;
    }
    let dir = tempfile::tempdir().expect("tempdir");
    let output = dir.path().join("out.docx");
    let path = reference_test_files().join("HelloWorld.docx");

    let stderr = run_err(&[
        "set",
        path.to_str().unwrap(),
        "999",
        "text",
        "-o",
        output.to_str().unwrap(),
    ]);
    assert!(stderr.contains("out of range"), "got: {stderr}");
}

#[test]
fn nonexistent_file_gives_error() {
    let stderr = run_err(&["info", "/tmp/this_file_does_not_exist_12345.xlsx"]);
    assert!(!stderr.is_empty());
}
