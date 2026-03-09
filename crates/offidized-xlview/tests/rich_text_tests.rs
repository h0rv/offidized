//! Rich Text Tests for offidized-xlview
//!
//! Tests for parsing rich text content in XLSX files.
//! Rich text allows multiple formats within a single cell, using <r> (run) elements
//! with <rPr> (run properties) to define formatting for each segment.
//!
//! These tests use the offidized-xlview adapter pipeline:
//! XLSX bytes -> offidized_xlsx::Workbook -> viewer Workbook
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
mod test_helpers;

use common::*;
use test_helpers::*;

/// Helper to get the display value of a cell, checking v, cached_display, and raw.
///
/// When cells reference shared strings with rich text runs, the adapter
/// stores the concatenated text in `raw` as `CellRawValue::String`,
/// so we must check all three value locations.
fn get_display_value(
    wb: &offidized_xlview::types::workbook::Workbook,
    cd: &offidized_xlview::types::workbook::CellData,
) -> String {
    if let Some(ref display) = cd.cell.cached_display {
        return display.clone();
    }
    if let Some(ref v) = cd.cell.v {
        return v.clone();
    }
    use offidized_xlview::types::workbook::CellRawValue;
    match cd.cell.raw.as_ref() {
        Some(CellRawValue::String(s)) => s.clone(),
        Some(CellRawValue::Number(n)) => n.to_string(),
        Some(CellRawValue::Boolean(b)) => if *b { "TRUE" } else { "FALSE" }.to_string(),
        Some(CellRawValue::Error(e)) => e.clone(),
        Some(CellRawValue::Date(n)) => n.to_string(),
        Some(CellRawValue::SharedString(idx)) => wb
            .shared_strings
            .get(*idx as usize)
            .cloned()
            .unwrap_or_default(),
        None => String::new(),
    }
}

/// Helper to create a minimal XLSX file for testing rich text
fn create_test_xlsx(shared_strings_xml: &str, sheet_xml: &str) -> Vec<u8> {
    create_xlsx_with_shared_strings_and_sheet(shared_strings_xml, sheet_xml)
}

/// Helper to create XLSX with inline strings in the sheet
fn create_test_xlsx_inline(sheet_xml: &str) -> Vec<u8> {
    let empty_shared_strings = r#"<?xml version="1.0" encoding="UTF-8"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="0" uniqueCount="0">
</sst>"#;
    create_test_xlsx(empty_shared_strings, sheet_xml)
}

// ============================================================================
// SHARED STRINGS RICH TEXT TESTS
// ============================================================================

mod shared_strings_rich_text {
    use super::*;

    /// Test 1: Simple rich text - Part bold, part normal
    /// Expected: Text content "Bold Normal" should be extracted (concatenated)
    #[test]
    fn test_simple_rich_text_bold_and_normal() {
        let shared_strings = r#"<?xml version="1.0" encoding="UTF-8"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="1" uniqueCount="1">
  <si>
    <r><rPr><b/></rPr><t>Bold</t></r>
    <r><t> Normal</t></r>
  </si>
</sst>"#;

        let sheet = r#"<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1" t="s"><v>0</v></c>
    </row>
  </sheetData>
</worksheet>"#;

        let xlsx_data = create_test_xlsx(shared_strings, sheet);
        let wb = load_xlsx(&xlsx_data);

        assert_sheet_count(&wb, 1);
        let cd = get_cell(&wb, 0, 0, 0).expect("Cell should exist");
        // Rich text should be concatenated to plain text
        // The value may be in v, cached_display, or raw depending on the adapter
        assert_cell_value(&wb, 0, 0, 0, "Bold Normal");
        assert!(
            cd.cell.v.is_some() || cd.cell.cached_display.is_some() || cd.cell.raw.is_some(),
            "Cell should have a value in v, cached_display, or raw"
        );
    }

    /// Test 2: Multiple formats - Bold, italic, and colored runs
    /// Expected: Text "Bold Italic Red" concatenated
    #[test]
    fn test_multiple_format_runs() {
        let shared_strings = r#"<?xml version="1.0" encoding="UTF-8"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="1" uniqueCount="1">
  <si>
    <r><rPr><b/></rPr><t>Bold </t></r>
    <r><rPr><i/></rPr><t>Italic </t></r>
    <r><rPr><color rgb="FFFF0000"/></rPr><t>Red</t></r>
  </si>
</sst>"#;

        let sheet = r#"<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1" t="s"><v>0</v></c>
    </row>
  </sheetData>
</worksheet>"#;

        let xlsx_data = create_test_xlsx(shared_strings, sheet);
        let wb = load_xlsx(&xlsx_data);
        assert_cell_value(&wb, 0, 0, 0, "Bold Italic Red");
    }

    /// Test 3: Font size changes within cell
    /// Expected: Text "Small Large" concatenated
    #[test]
    fn test_font_size_changes() {
        let shared_strings = r#"<?xml version="1.0" encoding="UTF-8"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="1" uniqueCount="1">
  <si>
    <r><rPr><sz val="8"/></rPr><t>Small </t></r>
    <r><rPr><sz val="14"/></rPr><t>Large</t></r>
  </si>
</sst>"#;

        let sheet = r#"<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1" t="s"><v>0</v></c>
    </row>
  </sheetData>
</worksheet>"#;

        let xlsx_data = create_test_xlsx(shared_strings, sheet);
        let wb = load_xlsx(&xlsx_data);
        assert_cell_value(&wb, 0, 0, 0, "Small Large");
    }

    /// Test 4: Different font families
    /// Expected: Text "Arial Courier" concatenated
    #[test]
    fn test_font_family_changes() {
        let shared_strings = r#"<?xml version="1.0" encoding="UTF-8"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="1" uniqueCount="1">
  <si>
    <r><rPr><rFont val="Arial"/></rPr><t>Arial </t></r>
    <r><rPr><rFont val="Courier New"/></rPr><t>Courier</t></r>
  </si>
</sst>"#;

        let sheet = r#"<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1" t="s"><v>0</v></c>
    </row>
  </sheetData>
</worksheet>"#;

        let xlsx_data = create_test_xlsx(shared_strings, sheet);
        let wb = load_xlsx(&xlsx_data);
        assert_cell_value(&wb, 0, 0, 0, "Arial Courier");
    }

    /// Test 5: Partial underline in rich text
    /// Expected: Text "Normal Underlined" concatenated
    #[test]
    fn test_partial_underline() {
        let shared_strings = r#"<?xml version="1.0" encoding="UTF-8"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="1" uniqueCount="1">
  <si>
    <r><t>Normal </t></r>
    <r><rPr><u/></rPr><t>Underlined</t></r>
  </si>
</sst>"#;

        let sheet = r#"<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1" t="s"><v>0</v></c>
    </row>
  </sheetData>
</worksheet>"#;

        let xlsx_data = create_test_xlsx(shared_strings, sheet);
        let wb = load_xlsx(&xlsx_data);
        assert_cell_value(&wb, 0, 0, 0, "Normal Underlined");
    }

    /// Test 6: Partial strikethrough in rich text
    /// Expected: Text "Normal Strikethrough" concatenated
    #[test]
    fn test_partial_strikethrough() {
        let shared_strings = r#"<?xml version="1.0" encoding="UTF-8"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="1" uniqueCount="1">
  <si>
    <r><t>Normal </t></r>
    <r><rPr><strike/></rPr><t>Strikethrough</t></r>
  </si>
</sst>"#;

        let sheet = r#"<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1" t="s"><v>0</v></c>
    </row>
  </sheetData>
</worksheet>"#;

        let xlsx_data = create_test_xlsx(shared_strings, sheet);
        let wb = load_xlsx(&xlsx_data);
        assert_cell_value(&wb, 0, 0, 0, "Normal Strikethrough");
    }

    /// Test 7: Subscript text using vertAlign
    /// Expected: Text "H2O" concatenated (subscript 2)
    #[test]
    fn test_subscript() {
        let shared_strings = r#"<?xml version="1.0" encoding="UTF-8"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="1" uniqueCount="1">
  <si>
    <r><t>H</t></r>
    <r><rPr><vertAlign val="subscript"/></rPr><t>2</t></r>
    <r><t>O</t></r>
  </si>
</sst>"#;

        let sheet = r#"<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1" t="s"><v>0</v></c>
    </row>
  </sheetData>
</worksheet>"#;

        let xlsx_data = create_test_xlsx(shared_strings, sheet);
        let wb = load_xlsx(&xlsx_data);
        assert_cell_value(&wb, 0, 0, 0, "H2O");
    }

    /// Test 8: Superscript text using vertAlign
    /// Expected: Text "x2" concatenated (superscript 2)
    #[test]
    fn test_superscript() {
        let shared_strings = r#"<?xml version="1.0" encoding="UTF-8"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="1" uniqueCount="1">
  <si>
    <r><t>x</t></r>
    <r><rPr><vertAlign val="superscript"/></rPr><t>2</t></r>
  </si>
</sst>"#;

        let sheet = r#"<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1" t="s"><v>0</v></c>
    </row>
  </sheetData>
</worksheet>"#;

        let xlsx_data = create_test_xlsx(shared_strings, sheet);
        let wb = load_xlsx(&xlsx_data);
        assert_cell_value(&wb, 0, 0, 0, "x2");
    }

    /// Test 9: Complex combination - Multiple properties in one run
    /// Expected: Text "Bold+Italic+Red+Size" concatenated
    #[test]
    fn test_complex_combination() {
        let shared_strings = r#"<?xml version="1.0" encoding="UTF-8"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="1" uniqueCount="1">
  <si>
    <r>
      <rPr>
        <b/>
        <i/>
        <color rgb="FFFF0000"/>
        <sz val="14"/>
        <rFont val="Arial"/>
        <u/>
      </rPr>
      <t>Bold+Italic+Red+Size</t>
    </r>
  </si>
</sst>"#;

        let sheet = r#"<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1" t="s"><v>0</v></c>
    </row>
  </sheetData>
</worksheet>"#;

        let xlsx_data = create_test_xlsx(shared_strings, sheet);
        let wb = load_xlsx(&xlsx_data);
        assert_cell_value(&wb, 0, 0, 0, "Bold+Italic+Red+Size");
    }

    /// Test: Mixed plain and rich text in shared strings table
    #[test]
    fn test_mixed_plain_and_rich_text() {
        let shared_strings = r#"<?xml version="1.0" encoding="UTF-8"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="3" uniqueCount="3">
  <si><t>Plain text</t></si>
  <si>
    <r><rPr><b/></rPr><t>Rich</t></r>
    <r><t> text</t></r>
  </si>
  <si><t>Another plain</t></si>
</sst>"#;

        let sheet = r#"<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1" t="s"><v>0</v></c>
      <c r="B1" t="s"><v>1</v></c>
      <c r="C1" t="s"><v>2</v></c>
    </row>
  </sheetData>
</worksheet>"#;

        let xlsx_data = create_test_xlsx(shared_strings, sheet);
        let wb = load_xlsx(&xlsx_data);

        assert_cell_value(&wb, 0, 0, 0, "Plain text");
        assert_cell_value(&wb, 0, 0, 1, "Rich text");
        assert_cell_value(&wb, 0, 0, 2, "Another plain");
    }

    /// Test: Empty runs should be handled gracefully
    #[test]
    fn test_empty_rich_text_runs() {
        let shared_strings = r#"<?xml version="1.0" encoding="UTF-8"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="1" uniqueCount="1">
  <si>
    <r><rPr><b/></rPr><t></t></r>
    <r><t>Content</t></r>
    <r><rPr><i/></rPr><t></t></r>
  </si>
</sst>"#;

        let sheet = r#"<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1" t="s"><v>0</v></c>
    </row>
  </sheetData>
</worksheet>"#;

        let xlsx_data = create_test_xlsx(shared_strings, sheet);
        let wb = load_xlsx(&xlsx_data);
        assert_cell_value(&wb, 0, 0, 0, "Content");
    }

    /// Test: Whitespace preservation in rich text runs
    #[test]
    fn test_whitespace_preservation() {
        let shared_strings = r#"<?xml version="1.0" encoding="UTF-8"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="1" uniqueCount="1">
  <si>
    <r><rPr><b/></rPr><t xml:space="preserve">Bold </t></r>
    <r><t xml:space="preserve"> and </t></r>
    <r><rPr><i/></rPr><t xml:space="preserve"> Italic</t></r>
  </si>
</sst>"#;

        let sheet = r#"<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1" t="s"><v>0</v></c>
    </row>
  </sheetData>
</worksheet>"#;

        let xlsx_data = create_test_xlsx(shared_strings, sheet);
        let wb = load_xlsx(&xlsx_data);

        let cd = get_cell(&wb, 0, 0, 0).expect("Cell should exist");
        let value = get_display_value(&wb, cd);
        assert!(
            value.contains("Bold") && value.contains("and") && value.contains("Italic"),
            "Expected all text parts, got: {}",
            value
        );
    }

    /// Test: Theme colors in rich text run properties
    #[test]
    fn test_theme_color_in_rich_text() {
        let shared_strings = r#"<?xml version="1.0" encoding="UTF-8"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="1" uniqueCount="1">
  <si>
    <r><rPr><color theme="4"/></rPr><t>Accent1 Color</t></r>
  </si>
</sst>"#;

        let sheet = r#"<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1" t="s"><v>0</v></c>
    </row>
  </sheetData>
</worksheet>"#;

        let xlsx_data = create_test_xlsx(shared_strings, sheet);
        let wb = load_xlsx(&xlsx_data);
        assert_cell_value(&wb, 0, 0, 0, "Accent1 Color");
    }

    /// Test: Double underline in rich text
    #[test]
    fn test_double_underline_rich_text() {
        let shared_strings = r#"<?xml version="1.0" encoding="UTF-8"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="1" uniqueCount="1">
  <si>
    <r><rPr><u val="double"/></rPr><t>Double Underline</t></r>
  </si>
</sst>"#;

        let sheet = r#"<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1" t="s"><v>0</v></c>
    </row>
  </sheetData>
</worksheet>"#;

        let xlsx_data = create_test_xlsx(shared_strings, sheet);
        let wb = load_xlsx(&xlsx_data);
        assert_cell_value(&wb, 0, 0, 0, "Double Underline");
    }

    /// Test: Rich text with character set (charset) and font family type
    #[test]
    fn test_rich_text_with_charset() {
        let shared_strings = r#"<?xml version="1.0" encoding="UTF-8"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="1" uniqueCount="1">
  <si>
    <r>
      <rPr>
        <rFont val="Calibri"/>
        <charset val="1"/>
        <family val="2"/>
        <scheme val="minor"/>
      </rPr>
      <t>With Charset</t>
    </r>
  </si>
</sst>"#;

        let sheet = r#"<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1" t="s"><v>0</v></c>
    </row>
  </sheetData>
</worksheet>"#;

        let xlsx_data = create_test_xlsx(shared_strings, sheet);
        let wb = load_xlsx(&xlsx_data);
        assert_cell_value(&wb, 0, 0, 0, "With Charset");
    }
}

// ============================================================================
// INLINE STRING RICH TEXT TESTS
// ============================================================================

mod inline_string_rich_text {
    use super::*;

    /// Test 10: Inline rich text in cell (t="inlineStr")
    #[test]
    #[ignore = "TODO: Inline rich text run concatenation not yet implemented"]
    fn test_inline_rich_text() {
        let sheet = r#"<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1" t="inlineStr">
        <is>
          <r><rPr><b/></rPr><t>Bold</t></r>
          <r><t> Normal</t></r>
        </is>
      </c>
    </row>
  </sheetData>
</worksheet>"#;

        let xlsx_data = create_test_xlsx_inline(sheet);
        let wb = load_xlsx(&xlsx_data);

        assert_sheet_count(&wb, 1);
        assert_cell_value(&wb, 0, 0, 0, "Bold Normal");
    }

    /// Test: Inline string with plain text (no rich text runs)
    #[test]
    fn test_inline_plain_text() {
        let sheet = r#"<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1" t="inlineStr">
        <is>
          <t>Plain inline string</t>
        </is>
      </c>
    </row>
  </sheetData>
</worksheet>"#;

        let xlsx_data = create_test_xlsx_inline(sheet);
        let wb = load_xlsx(&xlsx_data);
        assert_cell_value(&wb, 0, 0, 0, "Plain inline string");
    }

    /// Test: Inline rich text with multiple formatting runs
    #[test]
    #[ignore = "TODO: Inline rich text run concatenation not yet implemented"]
    fn test_inline_multiple_runs() {
        let sheet = r#"<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1" t="inlineStr">
        <is>
          <r><rPr><b/></rPr><t>Bold </t></r>
          <r><rPr><i/></rPr><t>Italic </t></r>
          <r><rPr><u/></rPr><t>Underline</t></r>
        </is>
      </c>
    </row>
  </sheetData>
</worksheet>"#;

        let xlsx_data = create_test_xlsx_inline(sheet);
        let wb = load_xlsx(&xlsx_data);
        assert_cell_value(&wb, 0, 0, 0, "Bold Italic Underline");
    }

    /// Test: Inline rich text with subscript/superscript
    #[test]
    #[ignore = "TODO: Inline rich text run concatenation not yet implemented"]
    fn test_inline_sub_superscript() {
        let sheet = r#"<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1" t="inlineStr">
        <is>
          <r><t>E=mc</t></r>
          <r><rPr><vertAlign val="superscript"/></rPr><t>2</t></r>
        </is>
      </c>
    </row>
  </sheetData>
</worksheet>"#;

        let xlsx_data = create_test_xlsx_inline(sheet);
        let wb = load_xlsx(&xlsx_data);
        assert_cell_value(&wb, 0, 0, 0, "E=mc2");
    }

    /// Test: Empty inline string
    #[test]
    fn test_empty_inline_string() {
        let sheet = r#"<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1" t="inlineStr">
        <is>
          <t></t>
        </is>
      </c>
    </row>
  </sheetData>
</worksheet>"#;

        let xlsx_data = create_test_xlsx_inline(sheet);
        let wb = load_xlsx(&xlsx_data);

        // Empty cell might not be included, or might have empty value
        let cd = get_cell(&wb, 0, 0, 0);
        if let Some(cd) = cd {
            assert!(
                cd.cell.v.is_none() || cd.cell.v.as_deref() == Some(""),
                "Expected empty value, got: {:?}",
                cd.cell.v
            );
        }
    }
}

// ============================================================================
// EDGE CASES AND ERROR HANDLING
// ============================================================================

mod edge_cases {
    use super::*;

    /// Test: Malformed rich text (missing <t> element)
    #[test]
    fn test_rich_text_missing_text_element() {
        let shared_strings = r#"<?xml version="1.0" encoding="UTF-8"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="1" uniqueCount="1">
  <si>
    <r><rPr><b/></rPr></r>
    <r><t>Only this</t></r>
  </si>
</sst>"#;

        let sheet = r#"<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1" t="s"><v>0</v></c>
    </row>
  </sheetData>
</worksheet>"#;

        let xlsx_data = create_test_xlsx(shared_strings, sheet);
        let wb = load_xlsx(&xlsx_data);
        assert_cell_value(&wb, 0, 0, 0, "Only this");
    }

    /// Test: Rich text with nested elements (like phonetic reading <rPh>)
    #[test]
    fn test_rich_text_with_phonetic_reading() {
        let shared_strings = r#"<?xml version="1.0" encoding="UTF-8"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="1" uniqueCount="1">
  <si>
    <t>Main Text</t>
    <rPh sb="0" eb="4">
      <t>phonetic</t>
    </rPh>
  </si>
</sst>"#;

        let sheet = r#"<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1" t="s"><v>0</v></c>
    </row>
  </sheetData>
</worksheet>"#;

        let xlsx_data = create_test_xlsx(shared_strings, sheet);
        let wb = load_xlsx(&xlsx_data);

        let cd = get_cell(&wb, 0, 0, 0).expect("Cell should exist");
        let value = get_display_value(&wb, cd);
        assert!(
            value.contains("Main Text"),
            "Expected main text, got: {}",
            value
        );
    }

    /// Test: Very long rich text string (many runs)
    #[test]
    fn test_many_rich_text_runs() {
        let runs: String = (0..10)
            .map(|i| {
                if i % 2 == 0 {
                    format!(r#"<r><rPr><b/></rPr><t>Run{} </t></r>"#, i)
                } else {
                    format!(r#"<r><rPr><i/></rPr><t>Run{} </t></r>"#, i)
                }
            })
            .collect();

        let shared_strings = format!(
            r#"<?xml version="1.0" encoding="UTF-8"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="1" uniqueCount="1">
  <si>
    {}
  </si>
</sst>"#,
            runs
        );

        let sheet = r#"<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1" t="s"><v>0</v></c>
    </row>
  </sheetData>
</worksheet>"#;

        let xlsx_data = create_test_xlsx(&shared_strings, sheet);
        let wb = load_xlsx(&xlsx_data);

        let cd = get_cell(&wb, 0, 0, 0).expect("Cell should exist");
        let value = get_display_value(&wb, cd);

        for i in 0..10 {
            assert!(
                value.contains(&format!("Run{}", i)),
                "Missing Run{} in: {}",
                i,
                value
            );
        }
    }

    /// Test: Unicode text in rich text runs
    #[test]
    fn test_unicode_in_rich_text() {
        let shared_strings = r#"<?xml version="1.0" encoding="UTF-8"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="1" uniqueCount="1">
  <si>
    <r><rPr><b/></rPr><t>Hello </t></r>
    <r><t>&#x4e16;&#x754c; </t></r>
    <r><rPr><i/></rPr><t>&#x41f;&#x440;&#x438;&#x432;&#x435;&#x442;</t></r>
  </si>
</sst>"#;

        let sheet = r#"<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1" t="s"><v>0</v></c>
    </row>
  </sheetData>
</worksheet>"#;

        let xlsx_data = create_test_xlsx(shared_strings, sheet);
        let wb = load_xlsx(&xlsx_data);

        let cd = get_cell(&wb, 0, 0, 0).expect("Cell should exist");
        let value = get_display_value(&wb, cd);

        assert!(value.contains("Hello"), "Missing English text");
    }

    /// Test: Special XML characters in rich text
    #[test]
    fn test_xml_entities_in_rich_text() {
        let shared_strings = r#"<?xml version="1.0" encoding="UTF-8"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="1" uniqueCount="1">
  <si>
    <r><rPr><b/></rPr><t>&lt;bold&gt;</t></r>
    <r><t> &amp; </t></r>
    <r><rPr><i/></rPr><t>&quot;quoted&quot;</t></r>
  </si>
</sst>"#;

        let sheet = r#"<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1" t="s"><v>0</v></c>
    </row>
  </sheetData>
</worksheet>"#;

        let xlsx_data = create_test_xlsx(shared_strings, sheet);
        let wb = load_xlsx(&xlsx_data);

        let cd = get_cell(&wb, 0, 0, 0).expect("Cell should exist");
        let value = get_display_value(&wb, cd);

        assert!(value.contains("<bold>"), "Expected unescaped <bold>");
        assert!(value.contains("&"), "Expected unescaped &");
        assert!(value.contains("\"quoted\""), "Expected unescaped quotes");
    }

    /// Test: Newlines within rich text runs
    #[test]
    fn test_newlines_in_rich_text() {
        let shared_strings = r#"<?xml version="1.0" encoding="UTF-8"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="1" uniqueCount="1">
  <si>
    <r><rPr><b/></rPr><t>Line1
Line2</t></r>
    <r><t>
Line3</t></r>
  </si>
</sst>"#;

        let sheet = r#"<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1" t="s"><v>0</v></c>
    </row>
  </sheetData>
</worksheet>"#;

        let xlsx_data = create_test_xlsx(shared_strings, sheet);
        let wb = load_xlsx(&xlsx_data);

        let cd = get_cell(&wb, 0, 0, 0).expect("Cell should exist");
        let value = get_display_value(&wb, cd);

        assert!(value.contains("Line1"), "Missing Line1");
        assert!(value.contains("Line2"), "Missing Line2");
        assert!(value.contains("Line3"), "Missing Line3");
    }
}

// ============================================================================
// FUTURE RICH TEXT DATA MODEL TESTS
// ============================================================================

mod future_rich_text_model {
    /// Document what the future data model should look like
    #[test]
    #[ignore = "Future enhancement: rich text run data model"]
    fn test_future_rich_text_runs_model() {
        // When rich text is fully supported, a cell should provide:
        // 1. Plain text (concatenated) - for backward compatibility
        // 2. Rich text runs with individual formatting
    }

    /// Document expected HTML conversion for rich text
    #[test]
    #[ignore = "Future enhancement: rich text to HTML conversion"]
    fn test_future_rich_text_to_html() {
        // Future: Rich text should be convertible to HTML for rendering
    }

    /// Document expected behavior for nested formatting
    #[test]
    #[ignore = "Future enhancement: rich text run property resolution"]
    fn test_future_nested_run_properties() {
        // Rich text runs can have multiple properties
        // The parser should correctly capture all of them
    }
}
