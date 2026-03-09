use std::path::{Path, PathBuf};

use offidized_xlsx::Workbook;

const CURATED_FIXTURES: &[&str] = &[
    "basicspreadsheet.xlsx",
    "Spreadsheet.xlsx",
    "Comments.xlsx",
    "Complex01.xlsx",
    "excel14.xlsx",
    "missingcalcchainpart.xlsx",
];
const STRESS_FIXTURES: &[&str] = &[
    "Comments.xlsx",
    "Complex01.xlsx",
    "MCExecl.xlsx",
    "Revision_NameCommentChange.xlsx",
    "Spreadsheet.xlsx",
    "Youtube.xlsx",
    "basicspreadsheet.xlsx",
    "excel14.xlsx",
    "extlst.xlsx",
    "malformed_uri.xlsx",
    "malformed_uri_long.xlsx",
    "missingcalcchainpart.xlsx",
    "npoi/testcases/test-data/openxml4j/ExcelWithHyperlinks.xlsx",
    "vmldrawingroot.xlsx",
];
const MIN_STRESS_SUCCESSES: usize = 10;

#[test]
fn curated_reference_fixtures_roundtrip_open_save_open() {
    let mut failures = Vec::new();
    for fixture in CURATED_FIXTURES {
        if let Err(error) = roundtrip_fixture(fixture) {
            failures.push(error);
        }
    }

    assert!(
        failures.is_empty(),
        "curated fixture failures ({}):\n{}",
        failures.len(),
        failures.join("\n")
    );
}

#[test]
#[ignore = "Stress corpus scan. Run manually with --ignored to review failures."]
fn stress_reference_corpus_roundtrip_open_save_open() {
    let mut successes = 0_usize;
    let mut failures = Vec::new();
    for fixture in STRESS_FIXTURES {
        match roundtrip_fixture(fixture) {
            Ok(()) => successes += 1,
            Err(error) => failures.push(error),
        }
    }

    eprintln!(
        "reference corpus roundtrip summary: successes={successes}, failures={}",
        failures.len()
    );
    if !failures.is_empty() {
        eprintln!("failed fixtures:\n{}", failures.join("\n"));
    }

    assert!(
        successes >= MIN_STRESS_SUCCESSES,
        "expected at least {MIN_STRESS_SUCCESSES} successful roundtrips, got {successes}"
    );
}

#[test]
fn closedxml_regular_autofilter_range_survives_open_save_open() {
    let fixture_path = reference_root()
        .join("ClosedXML/ClosedXML.Tests/Resource/Examples/AutoFilter/RegularAutoFilter.xlsx");
    if !fixture_path.is_file() {
        eprintln!(
            "skipping test: ClosedXML RegularAutoFilter fixture not found at `{}`",
            fixture_path.display()
        );
        return;
    }

    let workbook = Workbook::open(&fixture_path).expect("open ClosedXML RegularAutoFilter fixture");
    let before = worksheet_auto_filter_ranges(&workbook);
    let expected_ranges = vec![
        (
            "Single Column Numbers".to_string(),
            "A1".to_string(),
            "B8".to_string(),
        ),
        (
            "Single Column Strings".to_string(),
            "A1".to_string(),
            "A7".to_string(),
        ),
        (
            "Single Column Mixed".to_string(),
            "A1".to_string(),
            "A7".to_string(),
        ),
        (
            "Multi Column".to_string(),
            "A1".to_string(),
            "C7".to_string(),
        ),
    ];
    assert_eq!(before, expected_ranges);

    let output_dir = tempfile::tempdir().expect("create tempdir");
    let output_path = output_dir.path().join("regular-autofilter-roundtrip.xlsx");
    workbook
        .save(&output_path)
        .expect("save roundtripped workbook");

    let reopened = Workbook::open(&output_path).expect("reopen roundtripped workbook");
    let after = worksheet_auto_filter_ranges(&reopened);
    assert_eq!(after, expected_ranges);
}

fn roundtrip_fixture(fixture_name: &str) -> Result<(), String> {
    let fixture_path = fixture_path(fixture_name);
    if !fixture_path.is_file() {
        return Err(format!(
            "{fixture_name}: fixture not found at `{}`",
            fixture_path.display()
        ));
    }

    let workbook = Workbook::open(&fixture_path)
        .map_err(|error| format!("{fixture_name}: initial open failed: {error}"))?;
    let output_dir =
        tempfile::tempdir().map_err(|error| format!("{fixture_name}: tempdir failed: {error}"))?;
    let output_file_name = Path::new(fixture_name)
        .file_name()
        .and_then(|file_name| file_name.to_str())
        .filter(|file_name| !file_name.trim().is_empty())
        .unwrap_or("roundtrip.xlsx");
    let output_path = output_dir.path().join(output_file_name);

    workbook
        .save(&output_path)
        .map_err(|error| format!("{fixture_name}: save failed: {error}"))?;
    Workbook::open(&output_path)
        .map(|_| ())
        .map_err(|error| format!("{fixture_name}: reopen failed: {error}"))
}

fn fixture_path(fixture_name: &str) -> PathBuf {
    if fixture_name.contains('/') || fixture_name.contains('\\') {
        reference_root().join(fixture_name)
    } else {
        reference_fixture_dir().join(fixture_name)
    }
}

fn worksheet_auto_filter_ranges(workbook: &Workbook) -> Vec<(String, String, String)> {
    workbook
        .worksheets()
        .iter()
        .filter_map(|worksheet| {
            worksheet.auto_filter().and_then(|af| {
                af.range().map(|range| {
                    (
                        worksheet.name().to_string(),
                        range.start().to_string(),
                        range.end().to_string(),
                    )
                })
            })
        })
        .collect::<Vec<_>>()
}

fn reference_fixture_dir() -> PathBuf {
    reference_root().join("Open-XML-SDK/test/DocumentFormat.OpenXml.Tests.Assets/assets/TestFiles")
}

fn reference_root() -> PathBuf {
    let manifest_dir = Path::new(env!("CARGO_MANIFEST_DIR"));
    manifest_dir.join("../../references")
}
