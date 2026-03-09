//! Tests for frozen panes and split panes parsing in offidized-xlview
//!
//! Tests frozen rows, frozen columns, split panes, and various pane states
//! including state="frozen", state="frozenSplit", and state="split".
//!
//! Ported from xlview's frozen_panes_tests.rs. The original tests checked
//! PaneState, split_row, and split_col fields on the xlview Sheet type.
//! The offidized-xlview Sheet only exposes frozen_rows and frozen_cols,
//! so split/pane state assertions verify the frozen values instead.
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

use std::io::{Cursor, Write};
use zip::write::FileOptions;
use zip::ZipWriter;

mod common;
mod fixtures;
use common::load_xlsx;

// ============================================================================
// Test Helpers
// ============================================================================

/// Create a minimal XLSX file with custom sheet XML content for pane testing
fn create_xlsx_with_sheet_xml(sheet_xml: &str) -> Vec<u8> {
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

    // xl/styles.xml - minimal
    let _ = zip.start_file("xl/styles.xml", options);
    let _ = zip.write_all(
        br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>
<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>
<borders count="1"><border><left/><right/><top/><bottom/></border></borders>
<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
<cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellXfs>
</styleSheet>"#,
    );

    // xl/worksheets/sheet1.xml - custom content
    let _ = zip.start_file("xl/worksheets/sheet1.xml", options);
    let _ = zip.write_all(sheet_xml.as_bytes());

    let cursor = zip.finish().expect("Failed to finish ZIP");
    cursor.into_inner()
}

/// Parse XLSX bytes into the viewer's Workbook type
fn parse_xlsx(data: &[u8]) -> offidized_xlview::types::workbook::Workbook {
    load_xlsx(data)
}

// ============================================================================
// FROZEN ROWS ONLY TESTS
// ============================================================================

mod frozen_rows_only {
    use super::*;

    #[test]
    fn test_freeze_1_row() {
        let sheet_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<sheetViews>
<sheetView tabSelected="1" workbookViewId="0">
<pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/>
<selection pane="bottomLeft" activeCell="A2" sqref="A2"/>
</sheetView>
</sheetViews>
<sheetData/>
</worksheet>"#;

        let xlsx = create_xlsx_with_sheet_xml(sheet_xml);
        let workbook = parse_xlsx(&xlsx);

        assert_eq!(workbook.sheets.len(), 1);
        let sheet = &workbook.sheets[0];

        assert_eq!(sheet.frozen_rows, 1);
        assert_eq!(sheet.frozen_cols, 0);
    }

    #[test]
    fn test_freeze_2_rows() {
        let sheet_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<sheetViews>
<sheetView tabSelected="1" workbookViewId="0">
<pane ySplit="2" topLeftCell="A3" activePane="bottomLeft" state="frozen"/>
<selection pane="bottomLeft" activeCell="A3" sqref="A3"/>
</sheetView>
</sheetViews>
<sheetData/>
</worksheet>"#;

        let xlsx = create_xlsx_with_sheet_xml(sheet_xml);
        let workbook = parse_xlsx(&xlsx);

        let sheet = &workbook.sheets[0];

        assert_eq!(sheet.frozen_rows, 2);
        assert_eq!(sheet.frozen_cols, 0);
    }

    #[test]
    fn test_freeze_5_rows() {
        let sheet_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<sheetViews>
<sheetView tabSelected="1" workbookViewId="0">
<pane ySplit="5" topLeftCell="A6" activePane="bottomLeft" state="frozen"/>
<selection pane="bottomLeft" activeCell="A6" sqref="A6"/>
</sheetView>
</sheetViews>
<sheetData/>
</worksheet>"#;

        let xlsx = create_xlsx_with_sheet_xml(sheet_xml);
        let workbook = parse_xlsx(&xlsx);

        let sheet = &workbook.sheets[0];

        assert_eq!(sheet.frozen_rows, 5);
        assert_eq!(sheet.frozen_cols, 0);
    }
}

// ============================================================================
// FROZEN COLUMNS ONLY TESTS
// ============================================================================

mod frozen_cols_only {
    use super::*;

    #[test]
    fn test_freeze_1_column() {
        let sheet_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<sheetViews>
<sheetView tabSelected="1" workbookViewId="0">
<pane xSplit="1" topLeftCell="B1" activePane="topRight" state="frozen"/>
<selection pane="topRight" activeCell="B1" sqref="B1"/>
</sheetView>
</sheetViews>
<sheetData/>
</worksheet>"#;

        let xlsx = create_xlsx_with_sheet_xml(sheet_xml);
        let workbook = parse_xlsx(&xlsx);

        assert_eq!(workbook.sheets.len(), 1);
        let sheet = &workbook.sheets[0];

        assert_eq!(sheet.frozen_rows, 0);
        assert_eq!(sheet.frozen_cols, 1);
    }

    #[test]
    fn test_freeze_2_columns() {
        let sheet_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<sheetViews>
<sheetView tabSelected="1" workbookViewId="0">
<pane xSplit="2" topLeftCell="C1" activePane="topRight" state="frozen"/>
<selection pane="topRight" activeCell="C1" sqref="C1"/>
</sheetView>
</sheetViews>
<sheetData/>
</worksheet>"#;

        let xlsx = create_xlsx_with_sheet_xml(sheet_xml);
        let workbook = parse_xlsx(&xlsx);

        let sheet = &workbook.sheets[0];

        assert_eq!(sheet.frozen_rows, 0);
        assert_eq!(sheet.frozen_cols, 2);
    }

    #[test]
    fn test_freeze_5_columns() {
        let sheet_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<sheetViews>
<sheetView tabSelected="1" workbookViewId="0">
<pane xSplit="5" topLeftCell="F1" activePane="topRight" state="frozen"/>
<selection pane="topRight" activeCell="F1" sqref="F1"/>
</sheetView>
</sheetViews>
<sheetData/>
</worksheet>"#;

        let xlsx = create_xlsx_with_sheet_xml(sheet_xml);
        let workbook = parse_xlsx(&xlsx);

        let sheet = &workbook.sheets[0];

        assert_eq!(sheet.frozen_rows, 0);
        assert_eq!(sheet.frozen_cols, 5);
    }
}

// ============================================================================
// BOTH FROZEN ROWS AND COLUMNS TESTS
// ============================================================================

mod frozen_rows_and_cols {
    use super::*;

    #[test]
    fn test_freeze_1_row_1_column() {
        let sheet_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<sheetViews>
<sheetView tabSelected="1" workbookViewId="0">
<pane xSplit="1" ySplit="1" topLeftCell="B2" activePane="bottomRight" state="frozen"/>
<selection pane="topRight" activeCell="B1" sqref="B1"/>
<selection pane="bottomLeft" activeCell="A2" sqref="A2"/>
<selection pane="bottomRight" activeCell="B2" sqref="B2"/>
</sheetView>
</sheetViews>
<sheetData/>
</worksheet>"#;

        let xlsx = create_xlsx_with_sheet_xml(sheet_xml);
        let workbook = parse_xlsx(&xlsx);

        let sheet = &workbook.sheets[0];

        assert_eq!(sheet.frozen_rows, 1);
        assert_eq!(sheet.frozen_cols, 1);
    }

    #[test]
    fn test_freeze_2_rows_3_columns() {
        let sheet_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<sheetViews>
<sheetView tabSelected="1" workbookViewId="0">
<pane xSplit="3" ySplit="2" topLeftCell="D3" activePane="bottomRight" state="frozen"/>
<selection pane="topRight" activeCell="D1" sqref="D1"/>
<selection pane="bottomLeft" activeCell="A3" sqref="A3"/>
<selection pane="bottomRight" activeCell="D3" sqref="D3"/>
</sheetView>
</sheetViews>
<sheetData/>
</worksheet>"#;

        let xlsx = create_xlsx_with_sheet_xml(sheet_xml);
        let workbook = parse_xlsx(&xlsx);

        let sheet = &workbook.sheets[0];

        assert_eq!(sheet.frozen_rows, 2);
        assert_eq!(sheet.frozen_cols, 3);
    }

    #[test]
    fn test_freeze_5_rows_4_columns() {
        let sheet_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<sheetViews>
<sheetView tabSelected="1" workbookViewId="0">
<pane xSplit="4" ySplit="5" topLeftCell="E6" activePane="bottomRight" state="frozen"/>
<selection pane="topRight" activeCell="E1" sqref="E1"/>
<selection pane="bottomLeft" activeCell="A6" sqref="A6"/>
<selection pane="bottomRight" activeCell="E6" sqref="E6"/>
</sheetView>
</sheetViews>
<sheetData/>
</worksheet>"#;

        let xlsx = create_xlsx_with_sheet_xml(sheet_xml);
        let workbook = parse_xlsx(&xlsx);

        let sheet = &workbook.sheets[0];

        assert_eq!(sheet.frozen_rows, 5);
        assert_eq!(sheet.frozen_cols, 4);
    }
}

// ============================================================================
// DIFFERENT SPLIT POSITIONS TESTS
// ============================================================================

mod different_split_positions {
    use super::*;

    #[test]
    fn test_freeze_at_row_10() {
        let sheet_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<sheetViews>
<sheetView tabSelected="1" workbookViewId="0">
<pane ySplit="10" topLeftCell="A11" activePane="bottomLeft" state="frozen"/>
<selection pane="bottomLeft" activeCell="A11" sqref="A11"/>
</sheetView>
</sheetViews>
<sheetData/>
</worksheet>"#;

        let xlsx = create_xlsx_with_sheet_xml(sheet_xml);
        let workbook = parse_xlsx(&xlsx);

        let sheet = &workbook.sheets[0];

        assert_eq!(sheet.frozen_rows, 10);
        assert_eq!(sheet.frozen_cols, 0);
    }

    #[test]
    fn test_freeze_at_column_z() {
        // Freezing at column Z (26th column)
        let sheet_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<sheetViews>
<sheetView tabSelected="1" workbookViewId="0">
<pane xSplit="26" topLeftCell="AA1" activePane="topRight" state="frozen"/>
<selection pane="topRight" activeCell="AA1" sqref="AA1"/>
</sheetView>
</sheetViews>
<sheetData/>
</worksheet>"#;

        let xlsx = create_xlsx_with_sheet_xml(sheet_xml);
        let workbook = parse_xlsx(&xlsx);

        let sheet = &workbook.sheets[0];

        assert_eq!(sheet.frozen_rows, 0);
        assert_eq!(sheet.frozen_cols, 26);
    }

    #[test]
    fn test_freeze_at_row_100_col_10() {
        let sheet_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<sheetViews>
<sheetView tabSelected="1" workbookViewId="0">
<pane xSplit="10" ySplit="100" topLeftCell="K101" activePane="bottomRight" state="frozen"/>
<selection pane="bottomRight" activeCell="K101" sqref="K101"/>
</sheetView>
</sheetViews>
<sheetData/>
</worksheet>"#;

        let xlsx = create_xlsx_with_sheet_xml(sheet_xml);
        let workbook = parse_xlsx(&xlsx);

        let sheet = &workbook.sheets[0];

        assert_eq!(sheet.frozen_rows, 100);
        assert_eq!(sheet.frozen_cols, 10);
    }
}

// ============================================================================
// STATE="FROZEN" VS STATE="FROZENSPLIT" TESTS
// ============================================================================

mod frozen_vs_frozen_split {
    use super::*;

    #[test]
    fn test_state_frozen() {
        let sheet_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<sheetViews>
<sheetView tabSelected="1" workbookViewId="0">
<pane xSplit="2" ySplit="3" topLeftCell="C4" activePane="bottomRight" state="frozen"/>
<selection pane="bottomRight" activeCell="C4" sqref="C4"/>
</sheetView>
</sheetViews>
<sheetData/>
</worksheet>"#;

        let xlsx = create_xlsx_with_sheet_xml(sheet_xml);
        let workbook = parse_xlsx(&xlsx);

        let sheet = &workbook.sheets[0];

        assert_eq!(sheet.frozen_rows, 3);
        assert_eq!(sheet.frozen_cols, 2);
    }

    #[test]
    fn test_state_frozen_split() {
        // frozenSplit is used when panes are both frozen and have been split
        let sheet_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<sheetViews>
<sheetView tabSelected="1" workbookViewId="0">
<pane xSplit="2" ySplit="3" topLeftCell="C4" activePane="bottomRight" state="frozenSplit"/>
<selection pane="bottomRight" activeCell="C4" sqref="C4"/>
</sheetView>
</sheetViews>
<sheetData/>
</worksheet>"#;

        let xlsx = create_xlsx_with_sheet_xml(sheet_xml);
        let workbook = parse_xlsx(&xlsx);

        let sheet = &workbook.sheets[0];

        assert_eq!(sheet.frozen_rows, 3);
        assert_eq!(sheet.frozen_cols, 2);
    }

    #[test]
    fn test_frozen_split_single_dimension() {
        // frozenSplit with only rows frozen
        let sheet_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<sheetViews>
<sheetView tabSelected="1" workbookViewId="0">
<pane ySplit="5" topLeftCell="A6" activePane="bottomLeft" state="frozenSplit"/>
<selection pane="bottomLeft" activeCell="A6" sqref="A6"/>
</sheetView>
</sheetViews>
<sheetData/>
</worksheet>"#;

        let xlsx = create_xlsx_with_sheet_xml(sheet_xml);
        let workbook = parse_xlsx(&xlsx);

        let sheet = &workbook.sheets[0];

        assert_eq!(sheet.frozen_rows, 5);
        assert_eq!(sheet.frozen_cols, 0);
    }
}

// ============================================================================
// STATE="SPLIT" (NON-FROZEN SPLIT) TESTS
// ============================================================================

mod split_panes {
    use super::*;

    #[test]
    fn test_split_horizontal_only() {
        // Split panes use twips (1/20th of a point) for positioning
        let sheet_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<sheetViews>
<sheetView tabSelected="1" workbookViewId="0">
<pane ySplit="2400" topLeftCell="A10" activePane="bottomLeft" state="split"/>
<selection pane="bottomLeft" activeCell="A10" sqref="A10"/>
</sheetView>
</sheetViews>
<sheetData/>
</worksheet>"#;

        let xlsx = create_xlsx_with_sheet_xml(sheet_xml);
        let workbook = parse_xlsx(&xlsx);

        let sheet = &workbook.sheets[0];

        // For split state, frozen_rows/cols should be 0
        assert_eq!(sheet.frozen_rows, 0);
        assert_eq!(sheet.frozen_cols, 0);
    }

    #[test]
    fn test_split_vertical_only() {
        let sheet_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<sheetViews>
<sheetView tabSelected="1" workbookViewId="0">
<pane xSplit="3600" topLeftCell="E1" activePane="topRight" state="split"/>
<selection pane="topRight" activeCell="E1" sqref="E1"/>
</sheetView>
</sheetViews>
<sheetData/>
</worksheet>"#;

        let xlsx = create_xlsx_with_sheet_xml(sheet_xml);
        let workbook = parse_xlsx(&xlsx);

        let sheet = &workbook.sheets[0];

        assert_eq!(sheet.frozen_rows, 0);
        assert_eq!(sheet.frozen_cols, 0);
    }

    #[test]
    fn test_split_both_dimensions() {
        let sheet_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<sheetViews>
<sheetView tabSelected="1" workbookViewId="0">
<pane xSplit="2000" ySplit="1500" topLeftCell="D8" activePane="bottomRight" state="split"/>
<selection pane="topRight" activeCell="D1" sqref="D1"/>
<selection pane="bottomLeft" activeCell="A8" sqref="A8"/>
<selection pane="bottomRight" activeCell="D8" sqref="D8"/>
</sheetView>
</sheetViews>
<sheetData/>
</worksheet>"#;

        let xlsx = create_xlsx_with_sheet_xml(sheet_xml);
        let workbook = parse_xlsx(&xlsx);

        let sheet = &workbook.sheets[0];

        assert_eq!(sheet.frozen_rows, 0);
        assert_eq!(sheet.frozen_cols, 0);
    }

    #[test]
    fn test_split_without_state_attribute() {
        // When state attribute is missing but xSplit/ySplit are present
        let sheet_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<sheetViews>
<sheetView tabSelected="1" workbookViewId="0">
<pane xSplit="1800" ySplit="900" topLeftCell="C5" activePane="bottomRight"/>
<selection pane="bottomRight" activeCell="C5" sqref="C5"/>
</sheetView>
</sheetViews>
<sheetData/>
</worksheet>"#;

        let xlsx = create_xlsx_with_sheet_xml(sheet_xml);
        let workbook = parse_xlsx(&xlsx);

        let sheet = &workbook.sheets[0];

        // Should default to split behavior when no state is specified
        // (large split values indicate split, not frozen)
        assert_eq!(sheet.frozen_rows, 0);
        assert_eq!(sheet.frozen_cols, 0);
    }
}

// ============================================================================
// LARGE FREEZE PANES TESTS
// ============================================================================

mod large_freeze_panes {
    use super::*;

    #[test]
    fn test_freeze_10_rows() {
        let sheet_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<sheetViews>
<sheetView tabSelected="1" workbookViewId="0">
<pane ySplit="10" topLeftCell="A11" activePane="bottomLeft" state="frozen"/>
<selection pane="bottomLeft" activeCell="A11" sqref="A11"/>
</sheetView>
</sheetViews>
<sheetData/>
</worksheet>"#;

        let xlsx = create_xlsx_with_sheet_xml(sheet_xml);
        let workbook = parse_xlsx(&xlsx);

        let sheet = &workbook.sheets[0];

        assert_eq!(sheet.frozen_rows, 10);
        assert_eq!(sheet.frozen_cols, 0);
    }

    #[test]
    fn test_freeze_50_rows() {
        let sheet_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<sheetViews>
<sheetView tabSelected="1" workbookViewId="0">
<pane ySplit="50" topLeftCell="A51" activePane="bottomLeft" state="frozen"/>
<selection pane="bottomLeft" activeCell="A51" sqref="A51"/>
</sheetView>
</sheetViews>
<sheetData/>
</worksheet>"#;

        let xlsx = create_xlsx_with_sheet_xml(sheet_xml);
        let workbook = parse_xlsx(&xlsx);

        let sheet = &workbook.sheets[0];

        assert_eq!(sheet.frozen_rows, 50);
        assert_eq!(sheet.frozen_cols, 0);
    }

    #[test]
    fn test_freeze_10_columns() {
        let sheet_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<sheetViews>
<sheetView tabSelected="1" workbookViewId="0">
<pane xSplit="10" topLeftCell="K1" activePane="topRight" state="frozen"/>
<selection pane="topRight" activeCell="K1" sqref="K1"/>
</sheetView>
</sheetViews>
<sheetData/>
</worksheet>"#;

        let xlsx = create_xlsx_with_sheet_xml(sheet_xml);
        let workbook = parse_xlsx(&xlsx);

        let sheet = &workbook.sheets[0];

        assert_eq!(sheet.frozen_rows, 0);
        assert_eq!(sheet.frozen_cols, 10);
    }

    #[test]
    fn test_freeze_10_rows_10_columns() {
        let sheet_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<sheetViews>
<sheetView tabSelected="1" workbookViewId="0">
<pane xSplit="10" ySplit="10" topLeftCell="K11" activePane="bottomRight" state="frozen"/>
<selection pane="bottomRight" activeCell="K11" sqref="K11"/>
</sheetView>
</sheetViews>
<sheetData/>
</worksheet>"#;

        let xlsx = create_xlsx_with_sheet_xml(sheet_xml);
        let workbook = parse_xlsx(&xlsx);

        let sheet = &workbook.sheets[0];

        assert_eq!(sheet.frozen_rows, 10);
        assert_eq!(sheet.frozen_cols, 10);
    }
}

// ============================================================================
// NO FROZEN PANES TESTS (DEFAULT STATE)
// ============================================================================

mod no_frozen_panes {
    use super::*;

    #[test]
    fn test_no_pane_element() {
        let sheet_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<sheetViews>
<sheetView tabSelected="1" workbookViewId="0">
<selection activeCell="A1" sqref="A1"/>
</sheetView>
</sheetViews>
<sheetData/>
</worksheet>"#;

        let xlsx = create_xlsx_with_sheet_xml(sheet_xml);
        let workbook = parse_xlsx(&xlsx);

        let sheet = &workbook.sheets[0];

        assert_eq!(sheet.frozen_rows, 0);
        assert_eq!(sheet.frozen_cols, 0);
    }

    #[test]
    fn test_no_sheet_views() {
        let sheet_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<sheetData/>
</worksheet>"#;

        let xlsx = create_xlsx_with_sheet_xml(sheet_xml);
        let workbook = parse_xlsx(&xlsx);

        let sheet = &workbook.sheets[0];

        assert_eq!(sheet.frozen_rows, 0);
        assert_eq!(sheet.frozen_cols, 0);
    }

    #[test]
    fn test_empty_sheet_view() {
        let sheet_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<sheetViews>
<sheetView tabSelected="1" workbookViewId="0"/>
</sheetViews>
<sheetData/>
</worksheet>"#;

        let xlsx = create_xlsx_with_sheet_xml(sheet_xml);
        let workbook = parse_xlsx(&xlsx);

        let sheet = &workbook.sheets[0];

        assert_eq!(sheet.frozen_rows, 0);
        assert_eq!(sheet.frozen_cols, 0);
    }

    #[test]
    fn test_multiple_sheet_views_no_pane() {
        let sheet_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<sheetViews>
<sheetView workbookViewId="0">
<selection activeCell="A1" sqref="A1"/>
</sheetView>
<sheetView workbookViewId="1">
<selection activeCell="B2" sqref="B2"/>
</sheetView>
</sheetViews>
<sheetData/>
</worksheet>"#;

        let xlsx = create_xlsx_with_sheet_xml(sheet_xml);
        let workbook = parse_xlsx(&xlsx);

        let sheet = &workbook.sheets[0];

        // Without a pane element, there should be no frozen panes
        assert_eq!(sheet.frozen_rows, 0);
        assert_eq!(sheet.frozen_cols, 0);
    }
}

// ============================================================================
// ACTIVE PANE SETTINGS TESTS
// ============================================================================

mod active_pane_settings {
    use super::*;

    #[test]
    fn test_active_pane_top_right() {
        // When only columns are frozen, activePane should be topRight
        let sheet_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<sheetViews>
<sheetView tabSelected="1" workbookViewId="0">
<pane xSplit="2" topLeftCell="C1" activePane="topRight" state="frozen"/>
<selection pane="topRight" activeCell="C1" sqref="C1"/>
</sheetView>
</sheetViews>
<sheetData/>
</worksheet>"#;

        let xlsx = create_xlsx_with_sheet_xml(sheet_xml);
        let workbook = parse_xlsx(&xlsx);

        let sheet = &workbook.sheets[0];

        assert_eq!(sheet.frozen_rows, 0);
        assert_eq!(sheet.frozen_cols, 2);
    }

    #[test]
    fn test_active_pane_bottom_left() {
        // When only rows are frozen, activePane should be bottomLeft
        let sheet_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<sheetViews>
<sheetView tabSelected="1" workbookViewId="0">
<pane ySplit="3" topLeftCell="A4" activePane="bottomLeft" state="frozen"/>
<selection pane="bottomLeft" activeCell="A4" sqref="A4"/>
</sheetView>
</sheetViews>
<sheetData/>
</worksheet>"#;

        let xlsx = create_xlsx_with_sheet_xml(sheet_xml);
        let workbook = parse_xlsx(&xlsx);

        let sheet = &workbook.sheets[0];

        assert_eq!(sheet.frozen_rows, 3);
        assert_eq!(sheet.frozen_cols, 0);
    }

    #[test]
    fn test_active_pane_bottom_right() {
        // When both rows and columns are frozen, activePane should be bottomRight
        let sheet_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<sheetViews>
<sheetView tabSelected="1" workbookViewId="0">
<pane xSplit="2" ySplit="3" topLeftCell="C4" activePane="bottomRight" state="frozen"/>
<selection pane="topRight" activeCell="C1" sqref="C1"/>
<selection pane="bottomLeft" activeCell="A4" sqref="A4"/>
<selection pane="bottomRight" activeCell="C4" sqref="C4"/>
</sheetView>
</sheetViews>
<sheetData/>
</worksheet>"#;

        let xlsx = create_xlsx_with_sheet_xml(sheet_xml);
        let workbook = parse_xlsx(&xlsx);

        let sheet = &workbook.sheets[0];

        assert_eq!(sheet.frozen_rows, 3);
        assert_eq!(sheet.frozen_cols, 2);
    }

    #[test]
    fn test_top_left_cell_reference() {
        // topLeftCell indicates the first visible cell in the bottom-right pane
        let sheet_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<sheetViews>
<sheetView tabSelected="1" workbookViewId="0">
<pane xSplit="5" ySplit="10" topLeftCell="F11" activePane="bottomRight" state="frozen"/>
<selection pane="bottomRight" activeCell="F11" sqref="F11"/>
</sheetView>
</sheetViews>
<sheetData/>
</worksheet>"#;

        let xlsx = create_xlsx_with_sheet_xml(sheet_xml);
        let workbook = parse_xlsx(&xlsx);

        let sheet = &workbook.sheets[0];

        // The topLeftCell F11 confirms 5 columns (A-E) and 10 rows (1-10) are frozen
        assert_eq!(sheet.frozen_rows, 10);
        assert_eq!(sheet.frozen_cols, 5);
    }

    #[test]
    fn test_multiple_selection_elements() {
        // Multiple selection elements for different panes
        let sheet_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<sheetViews>
<sheetView tabSelected="1" workbookViewId="0">
<pane xSplit="1" ySplit="1" topLeftCell="B2" activePane="bottomRight" state="frozen"/>
<selection pane="topLeft" activeCell="A1" sqref="A1"/>
<selection pane="topRight" activeCell="B1" sqref="B1"/>
<selection pane="bottomLeft" activeCell="A2" sqref="A2"/>
<selection pane="bottomRight" activeCell="B2" sqref="B2"/>
</sheetView>
</sheetViews>
<sheetData/>
</worksheet>"#;

        let xlsx = create_xlsx_with_sheet_xml(sheet_xml);
        let workbook = parse_xlsx(&xlsx);

        let sheet = &workbook.sheets[0];

        assert_eq!(sheet.frozen_rows, 1);
        assert_eq!(sheet.frozen_cols, 1);
    }
}

// ============================================================================
// EDGE CASES
// ============================================================================

mod edge_cases {
    use super::*;

    #[test]
    fn test_zero_split_values() {
        // Edge case: xSplit and ySplit are both 0 with frozen state
        let sheet_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<sheetViews>
<sheetView tabSelected="1" workbookViewId="0">
<pane xSplit="0" ySplit="0" topLeftCell="A1" state="frozen"/>
<selection activeCell="A1" sqref="A1"/>
</sheetView>
</sheetViews>
<sheetData/>
</worksheet>"#;

        let xlsx = create_xlsx_with_sheet_xml(sheet_xml);
        let workbook = parse_xlsx(&xlsx);

        let sheet = &workbook.sheets[0];

        // Zero values should result in no frozen rows/cols
        assert_eq!(sheet.frozen_rows, 0);
        assert_eq!(sheet.frozen_cols, 0);
    }

    #[test]
    fn test_decimal_split_values() {
        // In split state, values can be decimal (twips)
        let sheet_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<sheetViews>
<sheetView tabSelected="1" workbookViewId="0">
<pane xSplit="1234.5" ySplit="5678.9" topLeftCell="E20" activePane="bottomRight" state="split"/>
<selection pane="bottomRight" activeCell="E20" sqref="E20"/>
</sheetView>
</sheetViews>
<sheetData/>
</worksheet>"#;

        let xlsx = create_xlsx_with_sheet_xml(sheet_xml);
        let workbook = parse_xlsx(&xlsx);

        let sheet = &workbook.sheets[0];

        // Split state: frozen_rows/cols should be 0
        assert_eq!(sheet.frozen_rows, 0);
        assert_eq!(sheet.frozen_cols, 0);
    }

    #[test]
    fn test_pane_element_with_only_state() {
        // Pane element with state but no xSplit or ySplit
        let sheet_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<sheetViews>
<sheetView tabSelected="1" workbookViewId="0">
<pane state="frozen"/>
<selection activeCell="A1" sqref="A1"/>
</sheetView>
</sheetViews>
<sheetData/>
</worksheet>"#;

        let xlsx = create_xlsx_with_sheet_xml(sheet_xml);
        let workbook = parse_xlsx(&xlsx);

        let sheet = &workbook.sheets[0];

        // State is frozen but no split values, so nothing is actually frozen
        assert_eq!(sheet.frozen_rows, 0);
        assert_eq!(sheet.frozen_cols, 0);
    }

    #[test]
    fn test_pane_with_data() {
        // Pane element with actual cell data in the sheet
        let sheet_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<sheetViews>
<sheetView tabSelected="1" workbookViewId="0">
<pane xSplit="2" ySplit="1" topLeftCell="C2" activePane="bottomRight" state="frozen"/>
<selection pane="bottomRight" activeCell="C2" sqref="C2"/>
</sheetView>
</sheetViews>
<sheetData>
<row r="1">
<c r="A1" t="s"><v>0</v></c>
<c r="B1" t="s"><v>1</v></c>
<c r="C1" t="s"><v>2</v></c>
</row>
<row r="2">
<c r="A2"><v>1</v></c>
<c r="B2"><v>2</v></c>
<c r="C2"><v>3</v></c>
</row>
</sheetData>
</worksheet>"#;

        let xlsx = create_xlsx_with_sheet_xml(sheet_xml);
        let workbook = parse_xlsx(&xlsx);

        let sheet = &workbook.sheets[0];

        assert_eq!(sheet.frozen_rows, 1);
        assert_eq!(sheet.frozen_cols, 2);
        // Verify cells are still parsed
        assert!(!sheet.cells.is_empty());
    }

    #[test]
    fn test_very_large_freeze() {
        // Test freezing a very large number of rows
        let sheet_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<sheetViews>
<sheetView tabSelected="1" workbookViewId="0">
<pane ySplit="1000" topLeftCell="A1001" activePane="bottomLeft" state="frozen"/>
<selection pane="bottomLeft" activeCell="A1001" sqref="A1001"/>
</sheetView>
</sheetViews>
<sheetData/>
</worksheet>"#;

        let xlsx = create_xlsx_with_sheet_xml(sheet_xml);
        let workbook = parse_xlsx(&xlsx);

        let sheet = &workbook.sheets[0];

        assert_eq!(sheet.frozen_rows, 1000);
        assert_eq!(sheet.frozen_cols, 0);
    }

    #[test]
    fn test_unknown_state_attribute() {
        // Test with an unknown state attribute (should not crash)
        let sheet_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<sheetViews>
<sheetView tabSelected="1" workbookViewId="0">
<pane xSplit="2" ySplit="2" topLeftCell="C3" state="unknown"/>
<selection activeCell="C3" sqref="C3"/>
</sheetView>
</sheetViews>
<sheetData/>
</worksheet>"#;

        let xlsx = create_xlsx_with_sheet_xml(sheet_xml);
        let workbook = parse_xlsx(&xlsx);

        let sheet = &workbook.sheets[0];

        // Unknown state should default to split behavior (not frozen)
        assert_eq!(sheet.frozen_rows, 0);
        assert_eq!(sheet.frozen_cols, 0);
    }
}

// ============================================================================
// COMBINED FEATURES TESTS
// ============================================================================

mod combined_features {
    use super::*;

    #[test]
    fn test_frozen_panes_with_hidden_rows() {
        // Test frozen panes in a sheet that also has hidden rows
        let sheet_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<sheetViews>
<sheetView tabSelected="1" workbookViewId="0">
<pane ySplit="3" topLeftCell="A4" activePane="bottomLeft" state="frozen"/>
<selection pane="bottomLeft" activeCell="A4" sqref="A4"/>
</sheetView>
</sheetViews>
<sheetData>
<row r="1"><c r="A1"><v>1</v></c></row>
<row r="2" hidden="1"><c r="A2"><v>2</v></c></row>
<row r="3"><c r="A3"><v>3</v></c></row>
<row r="4"><c r="A4"><v>4</v></c></row>
</sheetData>
</worksheet>"#;

        let xlsx = create_xlsx_with_sheet_xml(sheet_xml);
        let workbook = parse_xlsx(&xlsx);

        let sheet = &workbook.sheets[0];

        assert_eq!(sheet.frozen_rows, 3);
        // Hidden rows should still be tracked
        assert!(sheet.hidden_rows.contains(&1)); // Row 2 is hidden (0-indexed)
    }

    #[test]
    fn test_frozen_panes_with_column_widths() {
        // Test frozen panes with custom column widths
        let sheet_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<sheetViews>
<sheetView tabSelected="1" workbookViewId="0">
<pane xSplit="2" topLeftCell="C1" activePane="topRight" state="frozen"/>
<selection pane="topRight" activeCell="C1" sqref="C1"/>
</sheetView>
</sheetViews>
<cols>
<col min="1" max="1" width="20" customWidth="1"/>
<col min="2" max="2" width="30" customWidth="1"/>
</cols>
<sheetData/>
</worksheet>"#;

        let xlsx = create_xlsx_with_sheet_xml(sheet_xml);
        let workbook = parse_xlsx(&xlsx);

        let sheet = &workbook.sheets[0];

        assert_eq!(sheet.frozen_cols, 2);
        // Column widths should still be parsed
        assert_eq!(sheet.col_widths.len(), 2);
    }

    #[test]
    fn test_frozen_panes_with_merged_cells() {
        // Test frozen panes with merged cells
        let sheet_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<sheetViews>
<sheetView tabSelected="1" workbookViewId="0">
<pane xSplit="1" ySplit="1" topLeftCell="B2" activePane="bottomRight" state="frozen"/>
<selection pane="bottomRight" activeCell="B2" sqref="B2"/>
</sheetView>
</sheetViews>
<sheetData>
<row r="1">
<c r="A1" t="s"><v>0</v></c>
</row>
</sheetData>
<mergeCells count="1">
<mergeCell ref="B2:C3"/>
</mergeCells>
</worksheet>"#;

        let xlsx = create_xlsx_with_sheet_xml(sheet_xml);
        let workbook = parse_xlsx(&xlsx);

        let sheet = &workbook.sheets[0];

        assert_eq!(sheet.frozen_rows, 1);
        assert_eq!(sheet.frozen_cols, 1);
        // Merged cells should still be parsed
        assert_eq!(sheet.merges.len(), 1);
    }

    #[test]
    fn test_frozen_panes_in_protected_sheet() {
        // Test frozen panes in a protected sheet
        let sheet_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<sheetViews>
<sheetView tabSelected="1" workbookViewId="0">
<pane xSplit="1" ySplit="2" topLeftCell="B3" activePane="bottomRight" state="frozen"/>
<selection pane="bottomRight" activeCell="B3" sqref="B3"/>
</sheetView>
</sheetViews>
<sheetProtection sheet="1" objects="1" scenarios="1"/>
<sheetData/>
</worksheet>"#;

        let xlsx = create_xlsx_with_sheet_xml(sheet_xml);
        let workbook = parse_xlsx(&xlsx);

        let sheet = &workbook.sheets[0];

        // Frozen panes should still work in a protected sheet
        assert_eq!(sheet.frozen_rows, 2);
        assert_eq!(sheet.frozen_cols, 1);
        // Note: offidized-xlview Sheet does not expose is_protected field,
        // so we only verify frozen panes are correctly parsed
    }
}
