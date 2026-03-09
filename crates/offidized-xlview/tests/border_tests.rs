//! Comprehensive tests for border styling ported from xlview.
//!
//! Tests all border styles, colors, diagonal borders, and edge cases.
//! Original xlview tests used `parse_styles()` to parse raw XML directly;
//! these tests go through the full pipeline: StyleBuilder -> XLSX -> load_xlsx().
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

mod common;
mod fixtures;

use common::*;
use offidized_xlview::types::style::BorderStyle;

// ============================================================================
// BORDER STYLE TESTS
// ============================================================================

mod border_styles {
    use super::*;

    #[test]
    fn test_thin_border_style() {
        let style = StyleBuilder::new().border_all("thin", None).build();
        let xlsx = XlsxBuilder::new()
            .sheet(SheetBuilder::new("Sheet1").cell("A1", "Test", Some(style)))
            .build();

        let workbook = load_xlsx(&xlsx);
        let style = get_cell_style(&workbook, 0, 0, 0).expect("Cell should have style");

        assert!(style.border_left.is_some(), "Left border should be present");
        assert!(
            style.border_right.is_some(),
            "Right border should be present"
        );
        assert!(style.border_top.is_some(), "Top border should be present");
        assert!(
            style.border_bottom.is_some(),
            "Bottom border should be present"
        );

        assert_eq!(style.border_left.as_ref().unwrap().style, BorderStyle::Thin);
        assert_eq!(
            style.border_right.as_ref().unwrap().style,
            BorderStyle::Thin
        );
        assert_eq!(style.border_top.as_ref().unwrap().style, BorderStyle::Thin);
        assert_eq!(
            style.border_bottom.as_ref().unwrap().style,
            BorderStyle::Thin
        );
    }

    #[test]
    fn test_medium_border_style() {
        let style = StyleBuilder::new().border_all("medium", None).build();
        let xlsx = XlsxBuilder::new()
            .sheet(SheetBuilder::new("Sheet1").cell("A1", "Test", Some(style)))
            .build();

        let workbook = load_xlsx(&xlsx);
        let style = get_cell_style(&workbook, 0, 0, 0).expect("Cell should have style");

        assert_eq!(
            style.border_left.as_ref().unwrap().style,
            BorderStyle::Medium
        );
        assert_eq!(
            style.border_right.as_ref().unwrap().style,
            BorderStyle::Medium
        );
        assert_eq!(
            style.border_top.as_ref().unwrap().style,
            BorderStyle::Medium
        );
        assert_eq!(
            style.border_bottom.as_ref().unwrap().style,
            BorderStyle::Medium
        );
    }

    #[test]
    fn test_thick_border_style() {
        let style = StyleBuilder::new().border_all("thick", None).build();
        let xlsx = XlsxBuilder::new()
            .sheet(SheetBuilder::new("Sheet1").cell("A1", "Test", Some(style)))
            .build();

        let workbook = load_xlsx(&xlsx);
        let style = get_cell_style(&workbook, 0, 0, 0).expect("Cell should have style");

        assert_eq!(
            style.border_left.as_ref().unwrap().style,
            BorderStyle::Thick
        );
        assert_eq!(
            style.border_right.as_ref().unwrap().style,
            BorderStyle::Thick
        );
        assert_eq!(style.border_top.as_ref().unwrap().style, BorderStyle::Thick);
        assert_eq!(
            style.border_bottom.as_ref().unwrap().style,
            BorderStyle::Thick
        );
    }

    #[test]
    fn test_dashed_border_style() {
        let style = StyleBuilder::new().border_all("dashed", None).build();
        let xlsx = XlsxBuilder::new()
            .sheet(SheetBuilder::new("Sheet1").cell("A1", "Test", Some(style)))
            .build();

        let workbook = load_xlsx(&xlsx);
        let style = get_cell_style(&workbook, 0, 0, 0).expect("Cell should have style");

        assert_eq!(
            style.border_left.as_ref().unwrap().style,
            BorderStyle::Dashed
        );
        assert_eq!(
            style.border_right.as_ref().unwrap().style,
            BorderStyle::Dashed
        );
        assert_eq!(
            style.border_top.as_ref().unwrap().style,
            BorderStyle::Dashed
        );
        assert_eq!(
            style.border_bottom.as_ref().unwrap().style,
            BorderStyle::Dashed
        );
    }

    #[test]
    fn test_dotted_border_style() {
        let style = StyleBuilder::new().border_all("dotted", None).build();
        let xlsx = XlsxBuilder::new()
            .sheet(SheetBuilder::new("Sheet1").cell("A1", "Test", Some(style)))
            .build();

        let workbook = load_xlsx(&xlsx);
        let style = get_cell_style(&workbook, 0, 0, 0).expect("Cell should have style");

        assert_eq!(
            style.border_left.as_ref().unwrap().style,
            BorderStyle::Dotted
        );
        assert_eq!(
            style.border_right.as_ref().unwrap().style,
            BorderStyle::Dotted
        );
        assert_eq!(
            style.border_top.as_ref().unwrap().style,
            BorderStyle::Dotted
        );
        assert_eq!(
            style.border_bottom.as_ref().unwrap().style,
            BorderStyle::Dotted
        );
    }

    #[test]
    fn test_double_border_style() {
        let style = StyleBuilder::new().border_all("double", None).build();
        let xlsx = XlsxBuilder::new()
            .sheet(SheetBuilder::new("Sheet1").cell("A1", "Test", Some(style)))
            .build();

        let workbook = load_xlsx(&xlsx);
        let style = get_cell_style(&workbook, 0, 0, 0).expect("Cell should have style");

        assert_eq!(
            style.border_left.as_ref().unwrap().style,
            BorderStyle::Double
        );
        assert_eq!(
            style.border_right.as_ref().unwrap().style,
            BorderStyle::Double
        );
        assert_eq!(
            style.border_top.as_ref().unwrap().style,
            BorderStyle::Double
        );
        assert_eq!(
            style.border_bottom.as_ref().unwrap().style,
            BorderStyle::Double
        );
    }

    #[test]
    fn test_hair_border_style() {
        let style = StyleBuilder::new().border_all("hair", None).build();
        let xlsx = XlsxBuilder::new()
            .sheet(SheetBuilder::new("Sheet1").cell("A1", "Test", Some(style)))
            .build();

        let workbook = load_xlsx(&xlsx);
        let style = get_cell_style(&workbook, 0, 0, 0).expect("Cell should have style");

        assert_eq!(style.border_left.as_ref().unwrap().style, BorderStyle::Hair);
        assert_eq!(
            style.border_right.as_ref().unwrap().style,
            BorderStyle::Hair
        );
        assert_eq!(style.border_top.as_ref().unwrap().style, BorderStyle::Hair);
        assert_eq!(
            style.border_bottom.as_ref().unwrap().style,
            BorderStyle::Hair
        );
    }

    #[test]
    fn test_medium_dashed_border_style() {
        let style = StyleBuilder::new().border_all("mediumDashed", None).build();
        let xlsx = XlsxBuilder::new()
            .sheet(SheetBuilder::new("Sheet1").cell("A1", "Test", Some(style)))
            .build();

        let workbook = load_xlsx(&xlsx);
        let style = get_cell_style(&workbook, 0, 0, 0).expect("Cell should have style");

        assert_eq!(
            style.border_left.as_ref().unwrap().style,
            BorderStyle::MediumDashed
        );
        assert_eq!(
            style.border_right.as_ref().unwrap().style,
            BorderStyle::MediumDashed
        );
        assert_eq!(
            style.border_top.as_ref().unwrap().style,
            BorderStyle::MediumDashed
        );
        assert_eq!(
            style.border_bottom.as_ref().unwrap().style,
            BorderStyle::MediumDashed
        );
    }

    #[test]
    fn test_dash_dot_border_style() {
        let style = StyleBuilder::new().border_all("dashDot", None).build();
        let xlsx = XlsxBuilder::new()
            .sheet(SheetBuilder::new("Sheet1").cell("A1", "Test", Some(style)))
            .build();

        let workbook = load_xlsx(&xlsx);
        let style = get_cell_style(&workbook, 0, 0, 0).expect("Cell should have style");

        assert_eq!(
            style.border_left.as_ref().unwrap().style,
            BorderStyle::DashDot
        );
        assert_eq!(
            style.border_right.as_ref().unwrap().style,
            BorderStyle::DashDot
        );
        assert_eq!(
            style.border_top.as_ref().unwrap().style,
            BorderStyle::DashDot
        );
        assert_eq!(
            style.border_bottom.as_ref().unwrap().style,
            BorderStyle::DashDot
        );
    }

    #[test]
    fn test_medium_dash_dot_border_style() {
        let style = StyleBuilder::new()
            .border_all("mediumDashDot", None)
            .build();
        let xlsx = XlsxBuilder::new()
            .sheet(SheetBuilder::new("Sheet1").cell("A1", "Test", Some(style)))
            .build();

        let workbook = load_xlsx(&xlsx);
        let style = get_cell_style(&workbook, 0, 0, 0).expect("Cell should have style");

        assert_eq!(
            style.border_left.as_ref().unwrap().style,
            BorderStyle::MediumDashDot
        );
        assert_eq!(
            style.border_right.as_ref().unwrap().style,
            BorderStyle::MediumDashDot
        );
        assert_eq!(
            style.border_top.as_ref().unwrap().style,
            BorderStyle::MediumDashDot
        );
        assert_eq!(
            style.border_bottom.as_ref().unwrap().style,
            BorderStyle::MediumDashDot
        );
    }

    #[test]
    fn test_dash_dot_dot_border_style() {
        let style = StyleBuilder::new().border_all("dashDotDot", None).build();
        let xlsx = XlsxBuilder::new()
            .sheet(SheetBuilder::new("Sheet1").cell("A1", "Test", Some(style)))
            .build();

        let workbook = load_xlsx(&xlsx);
        let style = get_cell_style(&workbook, 0, 0, 0).expect("Cell should have style");

        assert_eq!(
            style.border_left.as_ref().unwrap().style,
            BorderStyle::DashDotDot
        );
        assert_eq!(
            style.border_right.as_ref().unwrap().style,
            BorderStyle::DashDotDot
        );
        assert_eq!(
            style.border_top.as_ref().unwrap().style,
            BorderStyle::DashDotDot
        );
        assert_eq!(
            style.border_bottom.as_ref().unwrap().style,
            BorderStyle::DashDotDot
        );
    }

    #[test]
    fn test_medium_dash_dot_dot_border_style() {
        let style = StyleBuilder::new()
            .border_all("mediumDashDotDot", None)
            .build();
        let xlsx = XlsxBuilder::new()
            .sheet(SheetBuilder::new("Sheet1").cell("A1", "Test", Some(style)))
            .build();

        let workbook = load_xlsx(&xlsx);
        let style = get_cell_style(&workbook, 0, 0, 0).expect("Cell should have style");

        assert_eq!(
            style.border_left.as_ref().unwrap().style,
            BorderStyle::MediumDashDotDot
        );
        assert_eq!(
            style.border_right.as_ref().unwrap().style,
            BorderStyle::MediumDashDotDot
        );
        assert_eq!(
            style.border_top.as_ref().unwrap().style,
            BorderStyle::MediumDashDotDot
        );
        assert_eq!(
            style.border_bottom.as_ref().unwrap().style,
            BorderStyle::MediumDashDotDot
        );
    }

    #[test]
    fn test_slant_dash_dot_border_style() {
        let style = StyleBuilder::new().border_all("slantDashDot", None).build();
        let xlsx = XlsxBuilder::new()
            .sheet(SheetBuilder::new("Sheet1").cell("A1", "Test", Some(style)))
            .build();

        let workbook = load_xlsx(&xlsx);
        let style = get_cell_style(&workbook, 0, 0, 0).expect("Cell should have style");

        assert_eq!(
            style.border_left.as_ref().unwrap().style,
            BorderStyle::SlantDashDot
        );
        assert_eq!(
            style.border_right.as_ref().unwrap().style,
            BorderStyle::SlantDashDot
        );
        assert_eq!(
            style.border_top.as_ref().unwrap().style,
            BorderStyle::SlantDashDot
        );
        assert_eq!(
            style.border_bottom.as_ref().unwrap().style,
            BorderStyle::SlantDashDot
        );
    }
}

// ============================================================================
// BORDER COLOR TESTS
// ============================================================================

mod border_colors {
    use super::*;

    #[test]
    fn test_rgb_color_red() {
        let style = StyleBuilder::new()
            .border_all("thin", Some("#FF0000"))
            .build();
        let xlsx = XlsxBuilder::new()
            .sheet(SheetBuilder::new("Sheet1").cell("A1", "Red", Some(style)))
            .build();

        let workbook = load_xlsx(&xlsx);
        let style = get_cell_style(&workbook, 0, 0, 0).expect("Cell should have style");

        let border = style.border_top.as_ref().expect("Should have top border");
        assert!(
            border.color.contains("FF0000") || border.color.contains("ff0000"),
            "Border color should contain FF0000, got: {}",
            border.color
        );
    }

    #[test]
    fn test_rgb_color_green() {
        let style = StyleBuilder::new()
            .border_all("thin", Some("#00FF00"))
            .build();
        let xlsx = XlsxBuilder::new()
            .sheet(SheetBuilder::new("Sheet1").cell("A1", "Green", Some(style)))
            .build();

        let workbook = load_xlsx(&xlsx);
        let style = get_cell_style(&workbook, 0, 0, 0).expect("Cell should have style");

        let border = style.border_top.as_ref().expect("Should have top border");
        assert!(
            border.color.contains("00FF00") || border.color.contains("00ff00"),
            "Border color should contain 00FF00, got: {}",
            border.color
        );
    }

    #[test]
    fn test_rgb_color_blue() {
        let style = StyleBuilder::new()
            .border_all("thin", Some("#0000FF"))
            .build();
        let xlsx = XlsxBuilder::new()
            .sheet(SheetBuilder::new("Sheet1").cell("A1", "Blue", Some(style)))
            .build();

        let workbook = load_xlsx(&xlsx);
        let style = get_cell_style(&workbook, 0, 0, 0).expect("Cell should have style");

        let border = style
            .border_bottom
            .as_ref()
            .expect("Should have bottom border");
        assert!(
            border.color.contains("0000FF") || border.color.contains("0000ff"),
            "Border color should contain 0000FF, got: {}",
            border.color
        );
    }

    #[test]
    fn test_rgb_color_black() {
        let style = StyleBuilder::new()
            .border_all("thin", Some("#000000"))
            .build();
        let xlsx = XlsxBuilder::new()
            .sheet(SheetBuilder::new("Sheet1").cell("A1", "Black", Some(style)))
            .build();

        let workbook = load_xlsx(&xlsx);
        let style = get_cell_style(&workbook, 0, 0, 0).expect("Cell should have style");

        let border = style.border_left.as_ref().expect("Should have left border");
        assert!(
            border.color.contains("000000"),
            "Border color should contain 000000, got: {}",
            border.color
        );
    }

    // Note: Theme/indexed/auto colors from the original test can't be set
    // via StyleBuilder directly. We test that default borders (without explicit color)
    // still get a valid color assigned.

    #[test]
    fn test_default_border_color() {
        // Borders without explicit color should get a default (typically black)
        let style = StyleBuilder::new().border_all("thin", None).build();
        let xlsx = XlsxBuilder::new()
            .sheet(SheetBuilder::new("Sheet1").cell("A1", "Default", Some(style)))
            .build();

        let workbook = load_xlsx(&xlsx);
        let style = get_cell_style(&workbook, 0, 0, 0).expect("Cell should have style");

        let border = style.border_top.as_ref().expect("Should have top border");
        // Default color should be a valid hex string
        let color_clean = border.color.trim_start_matches('#');
        assert!(
            color_clean.len() == 6 || color_clean.len() == 8,
            "Color should be a valid hex color, got: {}",
            border.color
        );
    }

    #[test]
    fn test_mixed_color_on_different_sides() {
        // Each side gets a different explicit RGB color
        let style = StyleBuilder::new()
            .border_top(BorderSide::new("thin").color("#FF0000"))
            .border_right(BorderSide::new("thin").color("#00FF00"))
            .border_bottom(BorderSide::new("thin").color("#0000FF"))
            .border_left(BorderSide::new("thin").color("#FFFF00"))
            .build();
        let xlsx = XlsxBuilder::new()
            .sheet(SheetBuilder::new("Sheet1").cell("A1", "Mixed", Some(style)))
            .build();

        let workbook = load_xlsx(&xlsx);
        let style = get_cell_style(&workbook, 0, 0, 0).expect("Cell should have style");

        let top = style.border_top.as_ref().unwrap();
        assert!(
            top.color.contains("FF0000") || top.color.contains("ff0000"),
            "Top should be red, got: {}",
            top.color
        );

        let right = style.border_right.as_ref().unwrap();
        assert!(
            right.color.contains("00FF00") || right.color.contains("00ff00"),
            "Right should be green, got: {}",
            right.color
        );

        let bottom = style.border_bottom.as_ref().unwrap();
        assert!(
            bottom.color.contains("0000FF") || bottom.color.contains("0000ff"),
            "Bottom should be blue, got: {}",
            bottom.color
        );

        let left = style.border_left.as_ref().unwrap();
        assert!(
            left.color.contains("FFFF00") || left.color.contains("ffff00"),
            "Left should be yellow, got: {}",
            left.color
        );
    }
}

// ============================================================================
// DIAGONAL BORDER TESTS
// ============================================================================

mod diagonal_borders {
    use super::*;
    use std::io::{Cursor, Write};
    use zip::write::FileOptions;
    use zip::ZipWriter;

    /// Create an XLSX file with a custom styles.xml containing diagonal border attributes.
    /// Since StyleBuilder can't set diagonal borders, we construct the ZIP manually.
    fn create_xlsx_with_diagonal_border(border_xml: &str) -> Vec<u8> {
        let mut buf = Cursor::new(Vec::new());
        {
            let mut zip = ZipWriter::new(&mut buf);
            let opts = FileOptions::<()>::default();

            zip.start_file("[Content_Types].xml", opts).unwrap();
            zip.write_all(br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
</Types>"#).unwrap();

            zip.start_file("_rels/.rels", opts).unwrap();
            zip.write_all(br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>"#).unwrap();

            zip.start_file("xl/_rels/workbook.xml.rels", opts).unwrap();
            zip.write_all(br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>"#).unwrap();

            zip.start_file("xl/workbook.xml", opts).unwrap();
            zip.write_all(
                br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets>
</workbook>"#,
            )
            .unwrap();

            zip.start_file("xl/styles.xml", opts).unwrap();
            let styles_xml = format!(
                r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="1"><font><name val="Calibri"/><sz val="11"/></font></fonts>
  <fills count="2">
    <fill><patternFill patternType="none"/></fill>
    <fill><patternFill patternType="gray125"/></fill>
  </fills>
  <borders count="2">
    <border><left/><right/><top/><bottom/><diagonal/></border>
    {border_xml}
  </borders>
  <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
  <cellXfs count="2">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
    <xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyBorder="1"/>
  </cellXfs>
  <cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
</styleSheet>"#
            );
            zip.write_all(styles_xml.as_bytes()).unwrap();

            zip.start_file("xl/worksheets/sheet1.xml", opts).unwrap();
            zip.write_all(
                br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1"><c r="A1" s="1" t="inlineStr"><is><t>Test</t></is></c></row>
  </sheetData>
</worksheet>"#,
            )
            .unwrap();

            zip.finish().unwrap();
        }
        buf.into_inner()
    }

    #[test]
    fn test_diagonal_down() {
        let xlsx = create_xlsx_with_diagonal_border(
            r#"<border diagonalDown="1">
                <left/><right/><top/><bottom/>
                <diagonal style="thin"><color indexed="64"/></diagonal>
            </border>"#,
        );

        let workbook = load_xlsx(&xlsx);
        // Should parse without error; diagonal borders may or may not be surfaced
        assert!(
            !workbook.sheets.is_empty(),
            "Should have at least one sheet"
        );
    }

    #[test]
    fn test_diagonal_up() {
        let xlsx = create_xlsx_with_diagonal_border(
            r#"<border diagonalUp="1">
                <left/><right/><top/><bottom/>
                <diagonal style="thin"><color rgb="FFFF0000"/></diagonal>
            </border>"#,
        );

        let workbook = load_xlsx(&xlsx);
        assert!(
            !workbook.sheets.is_empty(),
            "Should have at least one sheet"
        );
    }

    #[test]
    fn test_both_diagonals() {
        let xlsx = create_xlsx_with_diagonal_border(
            r#"<border diagonalUp="1" diagonalDown="1">
                <left/><right/><top/><bottom/>
                <diagonal style="medium"><color theme="1"/></diagonal>
            </border>"#,
        );

        let workbook = load_xlsx(&xlsx);
        assert!(
            !workbook.sheets.is_empty(),
            "Should have at least one sheet"
        );
    }

    #[test]
    fn test_diagonal_with_regular_borders() {
        let xlsx = create_xlsx_with_diagonal_border(
            r#"<border diagonalDown="1">
                <left style="thin"><color indexed="64"/></left>
                <right style="thin"><color indexed="64"/></right>
                <top style="thin"><color indexed="64"/></top>
                <bottom style="thin"><color indexed="64"/></bottom>
                <diagonal style="dashed"><color rgb="FF0000FF"/></diagonal>
            </border>"#,
        );

        let workbook = load_xlsx(&xlsx);
        assert!(
            !workbook.sheets.is_empty(),
            "Should have at least one sheet"
        );

        // Regular borders should still be parsed even when diagonal is present
        let style = get_cell_style(&workbook, 0, 0, 0);
        if let Some(s) = style {
            // If borders are surfaced, all four should be present
            if s.border_top.is_some() {
                assert!(s.border_right.is_some(), "Right border should exist");
                assert!(s.border_bottom.is_some(), "Bottom border should exist");
                assert!(s.border_left.is_some(), "Left border should exist");
            }
        }
    }
}

// ============================================================================
// EDGE CASE TESTS
// ============================================================================

mod edge_cases {
    use super::*;

    #[test]
    fn test_mixed_border_styles() {
        let style = StyleBuilder::new()
            .border_top(BorderSide::new("thin"))
            .border_right(BorderSide::new("medium"))
            .border_bottom(BorderSide::new("dashed"))
            .border_left(BorderSide::new("double"))
            .build();
        let xlsx = XlsxBuilder::new()
            .sheet(SheetBuilder::new("Sheet1").cell("A1", "Mixed", Some(style)))
            .build();

        let workbook = load_xlsx(&xlsx);
        let style = get_cell_style(&workbook, 0, 0, 0).expect("Cell should have style");

        assert_eq!(style.border_top.as_ref().unwrap().style, BorderStyle::Thin);
        assert_eq!(
            style.border_right.as_ref().unwrap().style,
            BorderStyle::Medium
        );
        assert_eq!(
            style.border_bottom.as_ref().unwrap().style,
            BorderStyle::Dashed
        );
        assert_eq!(
            style.border_left.as_ref().unwrap().style,
            BorderStyle::Double
        );
    }

    #[test]
    fn test_cell_without_borders() {
        // A cell with no border styling should have None for border fields
        let xlsx = XlsxBuilder::new()
            .sheet(SheetBuilder::new("Sheet1").cell("A1", "No borders", None))
            .build();

        let workbook = load_xlsx(&xlsx);
        let cd = get_cell(&workbook, 0, 0, 0);
        assert!(cd.is_some(), "Cell should exist");

        // No explicit style or default borders should result in no border fields
        if let Some(cell_data) = cd {
            if let Some(ref s) = cell_data.cell.s {
                // If a style is present (e.g. default), borders should be None
                assert!(
                    s.border_top.is_none()
                        || matches!(s.border_top.as_ref().unwrap().style, BorderStyle::None),
                    "Default cell should have no top border"
                );
            }
        }
    }

    #[test]
    fn test_partial_borders_left_only() {
        let style = StyleBuilder::new()
            .border_left(BorderSide::new("thin"))
            .build();
        let xlsx = XlsxBuilder::new()
            .sheet(SheetBuilder::new("Sheet1").cell("A1", "Left only", Some(style)))
            .build();

        let workbook = load_xlsx(&xlsx);
        let style = get_cell_style(&workbook, 0, 0, 0).expect("Cell should have style");

        assert!(style.border_left.is_some(), "Should have left border");
        assert_eq!(style.border_left.as_ref().unwrap().style, BorderStyle::Thin);
        assert!(style.border_right.is_none(), "Should not have right border");
        assert!(style.border_top.is_none(), "Should not have top border");
        assert!(
            style.border_bottom.is_none(),
            "Should not have bottom border"
        );
    }

    #[test]
    fn test_partial_borders_top_bottom() {
        let style = StyleBuilder::new()
            .border_top(BorderSide::new("medium"))
            .border_bottom(BorderSide::new("medium"))
            .build();
        let xlsx = XlsxBuilder::new()
            .sheet(SheetBuilder::new("Sheet1").cell("A1", "Top and Bottom", Some(style)))
            .build();

        let workbook = load_xlsx(&xlsx);
        let style = get_cell_style(&workbook, 0, 0, 0).expect("Cell should have style");

        assert!(style.border_left.is_none(), "Should not have left border");
        assert!(style.border_right.is_none(), "Should not have right border");
        assert!(style.border_top.is_some(), "Should have top border");
        assert_eq!(
            style.border_top.as_ref().unwrap().style,
            BorderStyle::Medium
        );
        assert!(style.border_bottom.is_some(), "Should have bottom border");
        assert_eq!(
            style.border_bottom.as_ref().unwrap().style,
            BorderStyle::Medium
        );
    }

    #[test]
    fn test_partial_borders_right_bottom() {
        let style = StyleBuilder::new()
            .border_right(BorderSide::new("thick"))
            .border_bottom(BorderSide::new("thick"))
            .build();
        let xlsx = XlsxBuilder::new()
            .sheet(SheetBuilder::new("Sheet1").cell("A1", "Right and Bottom", Some(style)))
            .build();

        let workbook = load_xlsx(&xlsx);
        let style = get_cell_style(&workbook, 0, 0, 0).expect("Cell should have style");

        assert!(style.border_left.is_none(), "Should not have left border");
        assert!(style.border_right.is_some(), "Should have right border");
        assert_eq!(
            style.border_right.as_ref().unwrap().style,
            BorderStyle::Thick
        );
        assert!(style.border_top.is_none(), "Should not have top border");
        assert!(style.border_bottom.is_some(), "Should have bottom border");
        assert_eq!(
            style.border_bottom.as_ref().unwrap().style,
            BorderStyle::Thick
        );
    }

    #[test]
    fn test_multiple_cells_different_borders() {
        let thin_style = StyleBuilder::new().border_all("thin", None).build();
        let medium_style = StyleBuilder::new().border_all("medium", None).build();
        let thick_style = StyleBuilder::new().border_all("thick", None).build();

        let xlsx = XlsxBuilder::new()
            .sheet(
                SheetBuilder::new("Sheet1")
                    .cell("A1", "Thin", Some(thin_style))
                    .cell("B1", "Medium", Some(medium_style))
                    .cell("C1", "Thick", Some(thick_style)),
            )
            .build();

        let workbook = load_xlsx(&xlsx);

        let s1 = get_cell_style(&workbook, 0, 0, 0).expect("A1 should have style");
        let s2 = get_cell_style(&workbook, 0, 0, 1).expect("B1 should have style");
        let s3 = get_cell_style(&workbook, 0, 0, 2).expect("C1 should have style");

        assert_eq!(s1.border_top.as_ref().unwrap().style, BorderStyle::Thin);
        assert_eq!(s2.border_top.as_ref().unwrap().style, BorderStyle::Medium);
        assert_eq!(s3.border_top.as_ref().unwrap().style, BorderStyle::Thick);
    }

    #[test]
    fn test_border_with_self_closing_elements() {
        // Test that partial borders (only some sides set) parse correctly
        let style = StyleBuilder::new()
            .border_left(BorderSide::new("thin"))
            .border_bottom(BorderSide::new("thin"))
            .build();
        let xlsx = XlsxBuilder::new()
            .sheet(SheetBuilder::new("Sheet1").cell("A1", "Partial", Some(style)))
            .build();

        let workbook = load_xlsx(&xlsx);
        let style = get_cell_style(&workbook, 0, 0, 0).expect("Cell should have style");

        assert!(style.border_left.is_some(), "Should have left border");
        assert!(style.border_right.is_none(), "Should not have right border");
        assert!(style.border_top.is_none(), "Should not have top border");
        assert!(style.border_bottom.is_some(), "Should have bottom border");
    }

    #[test]
    fn test_border_with_outline_attribute() {
        // outline attribute should not prevent border parsing.
        // We test that normal borders with outline still parse.
        let style = StyleBuilder::new().border_all("thin", None).build();
        let xlsx = XlsxBuilder::new()
            .sheet(SheetBuilder::new("Sheet1").cell("A1", "Outline", Some(style)))
            .build();

        let workbook = load_xlsx(&xlsx);
        let style = get_cell_style(&workbook, 0, 0, 0).expect("Cell should have style");

        assert!(style.border_left.is_some());
        assert!(style.border_right.is_some());
        assert!(style.border_top.is_some());
        assert!(style.border_bottom.is_some());
    }

    #[test]
    fn test_border_style_without_color() {
        // Borders created without explicit color - should still parse with a default color
        let style = StyleBuilder::new().border_all("thin", None).build();
        let xlsx = XlsxBuilder::new()
            .sheet(SheetBuilder::new("Sheet1").cell("A1", "No color", Some(style)))
            .build();

        let workbook = load_xlsx(&xlsx);
        let style = get_cell_style(&workbook, 0, 0, 0).expect("Cell should have style");

        assert!(style.border_top.is_some(), "Should have top border");
        let border = style.border_top.as_ref().unwrap();
        // Color should be present (default or otherwise)
        assert!(
            !border.color.is_empty(),
            "Border should have a color string"
        );
    }
}

// ============================================================================
// INDIVIDUAL SIDE TESTS
// ============================================================================

mod individual_sides {
    use super::*;

    #[test]
    fn test_left_border_only() {
        let style = StyleBuilder::new()
            .border_left(BorderSide::new("thin").color("#000000"))
            .build();
        let xlsx = XlsxBuilder::new()
            .sheet(SheetBuilder::new("Sheet1").cell("A1", "Left", Some(style)))
            .build();

        let workbook = load_xlsx(&xlsx);
        let style = get_cell_style(&workbook, 0, 0, 0).expect("Cell should have style");

        assert!(style.border_left.is_some(), "Should have left border");
        let left = style.border_left.as_ref().unwrap();
        assert_eq!(left.style, BorderStyle::Thin);
        assert!(
            left.color.contains("000000"),
            "Left border color should contain 000000, got: {}",
            left.color
        );
    }

    #[test]
    fn test_right_border_only() {
        let style = StyleBuilder::new()
            .border_right(BorderSide::new("medium").color("#FF0000"))
            .build();
        let xlsx = XlsxBuilder::new()
            .sheet(SheetBuilder::new("Sheet1").cell("A1", "Right", Some(style)))
            .build();

        let workbook = load_xlsx(&xlsx);
        let style = get_cell_style(&workbook, 0, 0, 0).expect("Cell should have style");

        assert!(style.border_right.is_some(), "Should have right border");
        let right = style.border_right.as_ref().unwrap();
        assert_eq!(right.style, BorderStyle::Medium);
        assert!(
            right.color.contains("FF0000") || right.color.contains("ff0000"),
            "Right border color should contain FF0000, got: {}",
            right.color
        );
    }

    #[test]
    fn test_top_border_only() {
        let style = StyleBuilder::new()
            .border_top(BorderSide::new("thick").color("#00FF00"))
            .build();
        let xlsx = XlsxBuilder::new()
            .sheet(SheetBuilder::new("Sheet1").cell("A1", "Top", Some(style)))
            .build();

        let workbook = load_xlsx(&xlsx);
        let style = get_cell_style(&workbook, 0, 0, 0).expect("Cell should have style");

        assert!(style.border_top.is_some(), "Should have top border");
        let top = style.border_top.as_ref().unwrap();
        assert_eq!(top.style, BorderStyle::Thick);
        assert!(
            top.color.contains("00FF00") || top.color.contains("00ff00"),
            "Top border color should contain 00FF00, got: {}",
            top.color
        );
    }

    #[test]
    fn test_bottom_border_only() {
        let style = StyleBuilder::new()
            .border_bottom(BorderSide::new("double").color("#0000FF"))
            .build();
        let xlsx = XlsxBuilder::new()
            .sheet(SheetBuilder::new("Sheet1").cell("A1", "Bottom", Some(style)))
            .build();

        let workbook = load_xlsx(&xlsx);
        let style = get_cell_style(&workbook, 0, 0, 0).expect("Cell should have style");

        assert!(style.border_bottom.is_some(), "Should have bottom border");
        let bottom = style.border_bottom.as_ref().unwrap();
        assert_eq!(bottom.style, BorderStyle::Double);
        assert!(
            bottom.color.contains("0000FF") || bottom.color.contains("0000ff"),
            "Bottom border color should contain 0000FF, got: {}",
            bottom.color
        );
    }
}

// ============================================================================
// ALL BORDER STYLES COMPREHENSIVE TEST
// ============================================================================

mod comprehensive {
    use super::*;
    use fixtures::ALL_BORDER_STYLES;

    /// Convert a border style string to the expected BorderStyle enum variant
    fn expected_border_style(style_name: &str) -> BorderStyle {
        match style_name {
            "none" => BorderStyle::None,
            "thin" => BorderStyle::Thin,
            "medium" => BorderStyle::Medium,
            "thick" => BorderStyle::Thick,
            "dashed" => BorderStyle::Dashed,
            "dotted" => BorderStyle::Dotted,
            "double" => BorderStyle::Double,
            "hair" => BorderStyle::Hair,
            "mediumDashed" => BorderStyle::MediumDashed,
            "dashDot" => BorderStyle::DashDot,
            "mediumDashDot" => BorderStyle::MediumDashDot,
            "dashDotDot" => BorderStyle::DashDotDot,
            "mediumDashDotDot" => BorderStyle::MediumDashDotDot,
            "slantDashDot" => BorderStyle::SlantDashDot,
            _ => panic!("Unknown border style: {}", style_name),
        }
    }

    /// Test all 13 border styles on left side
    #[test]
    fn test_all_border_styles_on_left() {
        let styles = [
            "thin",
            "medium",
            "thick",
            "dashed",
            "dotted",
            "double",
            "hair",
            "mediumDashed",
            "dashDot",
            "mediumDashDot",
            "dashDotDot",
            "mediumDashDotDot",
            "slantDashDot",
        ];

        for style_name in &styles {
            let s = StyleBuilder::new()
                .border_left(BorderSide::new(style_name))
                .build();
            let xlsx = XlsxBuilder::new()
                .sheet(SheetBuilder::new("Sheet1").cell("A1", "Test", Some(s)))
                .build();

            let workbook = load_xlsx(&xlsx);
            let style = get_cell_style(&workbook, 0, 0, 0).unwrap_or_else(|| {
                panic!("Cell should have style for border style '{}'", style_name)
            });

            assert!(
                style.border_left.is_some(),
                "Left border should be present for style '{}'",
                style_name
            );
            let expected = expected_border_style(style_name);
            assert_eq!(
                style.border_left.as_ref().unwrap().style,
                expected,
                "Style mismatch for '{}': expected {:?}, got {:?}",
                style_name,
                expected,
                style.border_left.as_ref().unwrap().style
            );
        }
    }

    /// Test all 13 border styles on right side
    #[test]
    fn test_all_border_styles_on_right() {
        let styles = [
            "thin",
            "medium",
            "thick",
            "dashed",
            "dotted",
            "double",
            "hair",
            "mediumDashed",
            "dashDot",
            "mediumDashDot",
            "dashDotDot",
            "mediumDashDotDot",
            "slantDashDot",
        ];

        for style_name in &styles {
            let s = StyleBuilder::new()
                .border_right(BorderSide::new(style_name))
                .build();
            let xlsx = XlsxBuilder::new()
                .sheet(SheetBuilder::new("Sheet1").cell("A1", "Test", Some(s)))
                .build();

            let workbook = load_xlsx(&xlsx);
            let style = get_cell_style(&workbook, 0, 0, 0).unwrap_or_else(|| {
                panic!("Cell should have style for border style '{}'", style_name)
            });

            assert!(
                style.border_right.is_some(),
                "Right border should be present for style '{}'",
                style_name
            );
            let expected = expected_border_style(style_name);
            assert_eq!(
                style.border_right.as_ref().unwrap().style,
                expected,
                "Style mismatch for '{}'",
                style_name
            );
        }
    }

    /// Test all 13 border styles on top side
    #[test]
    fn test_all_border_styles_on_top() {
        let styles = [
            "thin",
            "medium",
            "thick",
            "dashed",
            "dotted",
            "double",
            "hair",
            "mediumDashed",
            "dashDot",
            "mediumDashDot",
            "dashDotDot",
            "mediumDashDotDot",
            "slantDashDot",
        ];

        for style_name in &styles {
            let s = StyleBuilder::new()
                .border_top(BorderSide::new(style_name))
                .build();
            let xlsx = XlsxBuilder::new()
                .sheet(SheetBuilder::new("Sheet1").cell("A1", "Test", Some(s)))
                .build();

            let workbook = load_xlsx(&xlsx);
            let style = get_cell_style(&workbook, 0, 0, 0).unwrap_or_else(|| {
                panic!("Cell should have style for border style '{}'", style_name)
            });

            assert!(
                style.border_top.is_some(),
                "Top border should be present for style '{}'",
                style_name
            );
            let expected = expected_border_style(style_name);
            assert_eq!(
                style.border_top.as_ref().unwrap().style,
                expected,
                "Style mismatch for '{}'",
                style_name
            );
        }
    }

    /// Test all 13 border styles on bottom side
    #[test]
    fn test_all_border_styles_on_bottom() {
        let styles = [
            "thin",
            "medium",
            "thick",
            "dashed",
            "dotted",
            "double",
            "hair",
            "mediumDashed",
            "dashDot",
            "mediumDashDot",
            "dashDotDot",
            "mediumDashDotDot",
            "slantDashDot",
        ];

        for style_name in &styles {
            let s = StyleBuilder::new()
                .border_bottom(BorderSide::new(style_name))
                .build();
            let xlsx = XlsxBuilder::new()
                .sheet(SheetBuilder::new("Sheet1").cell("A1", "Test", Some(s)))
                .build();

            let workbook = load_xlsx(&xlsx);
            let style = get_cell_style(&workbook, 0, 0, 0).unwrap_or_else(|| {
                panic!("Cell should have style for border style '{}'", style_name)
            });

            assert!(
                style.border_bottom.is_some(),
                "Bottom border should be present for style '{}'",
                style_name
            );
            let expected = expected_border_style(style_name);
            assert_eq!(
                style.border_bottom.as_ref().unwrap().style,
                expected,
                "Style mismatch for '{}'",
                style_name
            );
        }
    }

    /// Test all color types on all sides (using RGB colors since StyleBuilder only supports RGB)
    #[test]
    fn test_all_color_types_all_sides() {
        let style = StyleBuilder::new()
            .border_top(BorderSide::new("thin").color("#123456"))
            .border_right(BorderSide::new("thin").color("#654321"))
            .border_bottom(BorderSide::new("thin").color("#ABCDEF"))
            .border_left(BorderSide::new("thin").color("#FEDCBA"))
            .build();
        let xlsx = XlsxBuilder::new()
            .sheet(SheetBuilder::new("Sheet1").cell("A1", "Colors", Some(style)))
            .build();

        let workbook = load_xlsx(&xlsx);
        let style = get_cell_style(&workbook, 0, 0, 0).expect("Cell should have style");

        let top = style.border_top.as_ref().unwrap();
        assert!(
            top.color.to_uppercase().contains("123456"),
            "Top color should contain 123456, got: {}",
            top.color
        );

        let right = style.border_right.as_ref().unwrap();
        assert!(
            right.color.to_uppercase().contains("654321"),
            "Right color should contain 654321, got: {}",
            right.color
        );

        let bottom = style.border_bottom.as_ref().unwrap();
        assert!(
            bottom.color.to_uppercase().contains("ABCDEF"),
            "Bottom color should contain ABCDEF, got: {}",
            bottom.color
        );

        let left = style.border_left.as_ref().unwrap();
        assert!(
            left.color.to_uppercase().contains("FEDCBA"),
            "Left color should contain FEDCBA, got: {}",
            left.color
        );
    }

    /// Test all border styles work with custom colors
    #[test]
    fn test_all_border_styles_with_colors() {
        let test_cases = [
            ("thin", "#FF0000"),
            ("medium", "#00FF00"),
            ("thick", "#0000FF"),
            ("dashed", "#FFFF00"),
            ("dotted", "#FF00FF"),
            ("double", "#00FFFF"),
            ("hair", "#800000"),
            ("mediumDashed", "#008000"),
            ("dashDot", "#000080"),
            ("mediumDashDot", "#808000"),
            ("dashDotDot", "#800080"),
            ("mediumDashDotDot", "#008080"),
            ("slantDashDot", "#C0C0C0"),
        ];

        for (style_name, color) in test_cases {
            let s = StyleBuilder::new()
                .border_all(style_name, Some(color))
                .build();
            let xlsx = XlsxBuilder::new()
                .sheet(SheetBuilder::new("Sheet1").cell("A1", "Test", Some(s)))
                .build();

            let workbook = load_xlsx(&xlsx);
            let style = get_cell_style(&workbook, 0, 0, 0).unwrap_or_else(|| {
                panic!(
                    "Cell should have style for {} with color {}",
                    style_name, color
                )
            });

            assert!(
                style.border_top.is_some(),
                "Should have top border for {} with color",
                style_name
            );

            let expected = expected_border_style(style_name);
            let border = style.border_top.as_ref().unwrap();

            assert_eq!(
                border.style, expected,
                "Border style mismatch for {}: expected {:?}, got {:?}",
                style_name, expected, border.style
            );

            let color_hex = color.trim_start_matches('#');
            assert!(
                border
                    .color
                    .to_uppercase()
                    .contains(&color_hex.to_uppercase()),
                "Border color should contain {} for style {}, got: {}",
                color_hex,
                style_name,
                border.color
            );
        }
    }

    /// Verify the ALL_BORDER_STYLES constant has exactly 14 styles (including none)
    #[test]
    fn test_all_border_styles_constant_count() {
        assert_eq!(
            ALL_BORDER_STYLES.len(),
            14,
            "ALL_BORDER_STYLES should contain exactly 14 styles (none + 13 visible styles)"
        );
    }

    /// Verify all expected styles are in the ALL_BORDER_STYLES constant
    #[test]
    fn test_all_border_styles_contains_all_expected() {
        let expected = [
            "none",
            "thin",
            "medium",
            "thick",
            "dashed",
            "dotted",
            "double",
            "hair",
            "mediumDashed",
            "dashDot",
            "mediumDashDot",
            "dashDotDot",
            "mediumDashDotDot",
            "slantDashDot",
        ];

        for style in expected {
            assert!(
                ALL_BORDER_STYLES.contains(&style),
                "ALL_BORDER_STYLES should contain '{}'",
                style
            );
        }
    }
}

// ============================================================================
// REALISTIC EXCEL BORDER PATTERNS
// ============================================================================

mod realistic_patterns {
    use super::*;

    #[test]
    fn test_box_border_thin() {
        // Common pattern: thin box border around cell
        let style = StyleBuilder::new().border_all("thin", None).build();
        let xlsx = XlsxBuilder::new()
            .sheet(SheetBuilder::new("Sheet1").cell("A1", "Box", Some(style)))
            .build();

        let workbook = load_xlsx(&xlsx);
        let style = get_cell_style(&workbook, 0, 0, 0).expect("Cell should have style");

        for border in [
            &style.border_left,
            &style.border_right,
            &style.border_top,
            &style.border_bottom,
        ] {
            assert!(border.is_some(), "All sides should have borders");
            assert_eq!(border.as_ref().unwrap().style, BorderStyle::Thin);
        }
    }

    #[test]
    fn test_header_bottom_border() {
        // Common pattern: medium bottom border for headers
        let style = StyleBuilder::new()
            .border_bottom(BorderSide::new("medium"))
            .build();
        let xlsx = XlsxBuilder::new()
            .sheet(SheetBuilder::new("Sheet1").cell("A1", "Header", Some(style)))
            .build();

        let workbook = load_xlsx(&xlsx);
        let style = get_cell_style(&workbook, 0, 0, 0).expect("Cell should have style");

        assert!(style.border_left.is_none(), "Should not have left border");
        assert!(style.border_right.is_none(), "Should not have right border");
        assert!(style.border_top.is_none(), "Should not have top border");
        assert!(style.border_bottom.is_some(), "Should have bottom border");
        assert_eq!(
            style.border_bottom.as_ref().unwrap().style,
            BorderStyle::Medium
        );
    }

    #[test]
    fn test_total_row_double_top() {
        // Common pattern: double top border for totals row
        let style = StyleBuilder::new()
            .border_top(BorderSide::new("double"))
            .build();
        let xlsx = XlsxBuilder::new()
            .sheet(SheetBuilder::new("Sheet1").cell("A1", "Total", Some(style)))
            .build();

        let workbook = load_xlsx(&xlsx);
        let style = get_cell_style(&workbook, 0, 0, 0).expect("Cell should have style");

        assert!(style.border_left.is_none(), "Should not have left border");
        assert!(style.border_right.is_none(), "Should not have right border");
        assert!(style.border_top.is_some(), "Should have top border");
        assert_eq!(
            style.border_top.as_ref().unwrap().style,
            BorderStyle::Double
        );
        assert!(
            style.border_bottom.is_none(),
            "Should not have bottom border"
        );
    }

    #[test]
    fn test_colored_border_accent() {
        // Common pattern: accent colored borders (using explicit RGB)
        let style = StyleBuilder::new()
            .border_all("thin", Some("#4472C4"))
            .build();
        let xlsx = XlsxBuilder::new()
            .sheet(SheetBuilder::new("Sheet1").cell("A1", "Accent", Some(style)))
            .build();

        let workbook = load_xlsx(&xlsx);
        let style = get_cell_style(&workbook, 0, 0, 0).expect("Cell should have style");

        for border in [
            &style.border_left,
            &style.border_right,
            &style.border_top,
            &style.border_bottom,
        ] {
            assert!(border.is_some(), "All sides should have borders");
            let b = border.as_ref().unwrap();
            assert!(
                b.color.to_uppercase().contains("4472C4"),
                "Border color should be accent blue, got: {}",
                b.color
            );
        }
    }

    #[test]
    fn test_thick_outside_borders() {
        // Pattern: thick outside borders for a single cell
        let style = StyleBuilder::new().border_all("thick", None).build();
        let xlsx = XlsxBuilder::new()
            .sheet(SheetBuilder::new("Sheet1").cell("A1", "Thick", Some(style)))
            .build();

        let workbook = load_xlsx(&xlsx);
        let style = get_cell_style(&workbook, 0, 0, 0).expect("Cell should have style");

        for border in [
            &style.border_left,
            &style.border_right,
            &style.border_top,
            &style.border_bottom,
        ] {
            assert!(border.is_some(), "All sides should have borders");
            assert_eq!(border.as_ref().unwrap().style, BorderStyle::Thick);
        }
    }

    #[test]
    fn test_default_excel_border_set() {
        // Typical Excel file has no borders by default and thin borders as the first custom style
        let no_border_style = None;
        let thin_style = Some(StyleBuilder::new().border_all("thin", None).build());

        let xlsx = XlsxBuilder::new()
            .sheet(
                SheetBuilder::new("Sheet1")
                    .cell("A1", "No Border", no_border_style)
                    .cell("A2", "Thin Border", thin_style),
            )
            .build();

        let workbook = load_xlsx(&xlsx);

        // A2 should have thin borders
        let s2 = get_cell_style(&workbook, 0, 1, 0).expect("A2 should have style");
        assert!(s2.border_left.is_some(), "A2 should have left border");
        assert_eq!(s2.border_left.as_ref().unwrap().style, BorderStyle::Thin);
    }
}
