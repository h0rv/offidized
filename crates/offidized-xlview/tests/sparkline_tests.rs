//! Tests for sparkline parsing in XLSX files
//!
//! Sparklines are mini-charts embedded in cells, introduced in Excel 2010.
//! They are stored in the worksheet's extLst (extension list) element using
//! the x14 namespace (SpreadsheetML 2009/9/main extensions).
//!
//! Structure in sheetN.xml:
//! ```xml
//! <extLst>
//!   <ext uri="{05C60535-1F16-4fd2-B633-F4F36F0B64E0}">
//!     <x14:sparklineGroups xmlns:x14="...">
//!       <x14:sparklineGroup type="line" displayEmptyCellsAs="gap">
//!         <x14:colorSeries rgb="FF376092"/>
//!         <x14:colorNegative rgb="FFD00000"/>
//!         <x14:colorAxis rgb="FF000000"/>
//!         <x14:colorMarkers rgb="FFD00000"/>
//!         <x14:colorFirst rgb="FFD00000"/>
//!         <x14:colorLast rgb="FFD00000"/>
//!         <x14:colorHigh rgb="FFD00000"/>
//!         <x14:colorLow rgb="FFD00000"/>
//!         <x14:sparklines>
//!           <x14:sparkline>
//!             <xm:f>Sheet1!A1:A10</xm:f>
//!             <xm:sqref>B1</xm:sqref>
//!           </x14:sparkline>
//!         </x14:sparklines>
//!       </x14:sparklineGroup>
//!     </x14:sparklineGroups>
//!   </ext>
//! </extLst>
//! ```
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

use std::io::{Cursor, Write};
use zip::write::FileOptions;
use zip::ZipWriter;

// =============================================================================
// Test Helper: Create XLSX with sparklines
// =============================================================================

/// Create base XLSX structure with custom sheet content
fn create_xlsx_with_sheet_content(sheet_xml: &str) -> Vec<u8> {
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

    // xl/styles.xml
    let _ = zip.start_file("xl/styles.xml", options);
    let _ = zip.write_all(
        br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>
<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>
<borders count="1"><border><left/><right/><top/><bottom/></border></borders>
<cellXfs count="1"><xf fontId="0" fillId="0" borderId="0"/></cellXfs>
</styleSheet>"#,
    );

    // xl/worksheets/sheet1.xml
    let _ = zip.start_file("xl/worksheets/sheet1.xml", options);
    let _ = zip.write_all(sheet_xml.as_bytes());

    let cursor = zip.finish().expect("Failed to finish ZIP");
    cursor.into_inner()
}

/// Generate sparkline group XML
fn sparkline_group_xml(
    sparkline_type: &str,
    data_range: &str,
    location: &str,
    colors: Option<&SparklineColorConfig>,
    options: Option<&SparklineOptions>,
) -> String {
    let opts = options.cloned().unwrap_or_default();
    let colors = colors.cloned().unwrap_or_default();

    let mut attrs = format!(r#"type="{}""#, sparkline_type);

    if let Some(ref empty) = opts.display_empty_cells_as {
        attrs.push_str(&format!(r#" displayEmptyCellsAs="{}""#, empty));
    }
    if opts.markers {
        attrs.push_str(r#" markers="1""#);
    }
    if opts.high {
        attrs.push_str(r#" high="1""#);
    }
    if opts.low {
        attrs.push_str(r#" low="1""#);
    }
    if opts.first {
        attrs.push_str(r#" first="1""#);
    }
    if opts.last {
        attrs.push_str(r#" last="1""#);
    }
    if opts.negative {
        attrs.push_str(r#" negative="1""#);
    }
    if opts.display_x_axis {
        attrs.push_str(r#" displayXAxis="1""#);
    }
    if opts.right_to_left {
        attrs.push_str(r#" rightToLeft="1""#);
    }
    if let Some(ref axis_type) = opts.min_axis_type {
        attrs.push_str(&format!(r#" minAxisType="{}""#, axis_type));
    }
    if let Some(ref axis_type) = opts.max_axis_type {
        attrs.push_str(&format!(r#" maxAxisType="{}""#, axis_type));
    }
    if let Some(val) = opts.manual_min {
        attrs.push_str(&format!(r#" manualMin="{}""#, val));
    }
    if let Some(val) = opts.manual_max {
        attrs.push_str(&format!(r#" manualMax="{}""#, val));
    }

    let mut color_elements = String::new();
    if let Some(ref c) = colors.series {
        color_elements.push_str(&format!(r#"<x14:colorSeries rgb="{}"/>"#, c));
    }
    if let Some(ref c) = colors.negative {
        color_elements.push_str(&format!(r#"<x14:colorNegative rgb="{}"/>"#, c));
    }
    if let Some(ref c) = colors.axis {
        color_elements.push_str(&format!(r#"<x14:colorAxis rgb="{}"/>"#, c));
    }
    if let Some(ref c) = colors.markers {
        color_elements.push_str(&format!(r#"<x14:colorMarkers rgb="{}"/>"#, c));
    }
    if let Some(ref c) = colors.first {
        color_elements.push_str(&format!(r#"<x14:colorFirst rgb="{}"/>"#, c));
    }
    if let Some(ref c) = colors.last {
        color_elements.push_str(&format!(r#"<x14:colorLast rgb="{}"/>"#, c));
    }
    if let Some(ref c) = colors.high {
        color_elements.push_str(&format!(r#"<x14:colorHigh rgb="{}"/>"#, c));
    }
    if let Some(ref c) = colors.low {
        color_elements.push_str(&format!(r#"<x14:colorLow rgb="{}"/>"#, c));
    }

    format!(
        r#"<x14:sparklineGroup {attrs}>
{color_elements}
<x14:sparklines>
<x14:sparkline>
<xm:f>{data_range}</xm:f>
<xm:sqref>{location}</xm:sqref>
</x14:sparkline>
</x14:sparklines>
</x14:sparklineGroup>"#
    )
}

/// Generate multiple sparklines in a group
fn sparkline_group_multi_xml(
    sparkline_type: &str,
    sparklines: &[(&str, &str)], // (data_range, location)
    colors: Option<&SparklineColorConfig>,
) -> String {
    let colors = colors.cloned().unwrap_or_default();

    let mut color_elements = String::new();
    if let Some(ref c) = colors.series {
        color_elements.push_str(&format!(r#"<x14:colorSeries rgb="{}"/>"#, c));
    }

    let sparkline_elements: String = sparklines
        .iter()
        .map(|(data, loc)| {
            format!(
                r#"<x14:sparkline>
<xm:f>{}</xm:f>
<xm:sqref>{}</xm:sqref>
</x14:sparkline>"#,
                data, loc
            )
        })
        .collect();

    format!(
        r#"<x14:sparklineGroup type="{}">
{color_elements}
<x14:sparklines>
{sparkline_elements}
</x14:sparklines>
</x14:sparklineGroup>"#,
        sparkline_type
    )
}

/// Wrap sparkline groups in sheet XML
fn sheet_with_sparklines(groups: &[String]) -> String {
    let groups_xml: String = groups.join("\n");

    format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
           xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
           xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
           xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"
           xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main"
           mc:Ignorable="x14">
<sheetData>
<row r="1">
<c r="A1"><v>1</v></c>
<c r="B1"><v>2</v></c>
<c r="C1"><v>3</v></c>
<c r="D1"><v>4</v></c>
<c r="E1"><v>5</v></c>
</row>
<row r="2">
<c r="A2"><v>-1</v></c>
<c r="B2"><v>3</v></c>
<c r="C2"><v>-2</v></c>
<c r="D2"><v>4</v></c>
<c r="E2"><v>-3</v></c>
</row>
</sheetData>
<extLst>
<ext uri="{{05C60535-1F16-4FD2-B633-F4F36F0B64E0}}"
     xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main">
<x14:sparklineGroups xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main">
{groups_xml}
</x14:sparklineGroups>
</ext>
</extLst>
</worksheet>"#
    )
}

/// Sheet without sparklines
fn sheet_without_sparklines() -> String {
    r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<sheetData>
<row r="1">
<c r="A1"><v>1</v></c>
<c r="B1"><v>2</v></c>
</row>
</sheetData>
</worksheet>"#
        .to_string()
}

// =============================================================================
// Test Configuration Structs
// =============================================================================

#[derive(Clone, Default)]
struct SparklineColorConfig {
    series: Option<String>,
    negative: Option<String>,
    axis: Option<String>,
    markers: Option<String>,
    first: Option<String>,
    last: Option<String>,
    high: Option<String>,
    low: Option<String>,
}

#[derive(Clone, Default)]
struct SparklineOptions {
    display_empty_cells_as: Option<String>,
    markers: bool,
    high: bool,
    low: bool,
    first: bool,
    last: bool,
    negative: bool,
    display_x_axis: bool,
    right_to_left: bool,
    min_axis_type: Option<String>,
    max_axis_type: Option<String>,
    manual_min: Option<f64>,
    manual_max: Option<f64>,
}

// =============================================================================
// Tests: Line Sparklines
// =============================================================================

#[test]
fn test_line_sparkline_basic() {
    let group = sparkline_group_xml("line", "Sheet1!A1:E1", "F1", None, None);
    let sheet = sheet_with_sparklines(&[group]);
    let xlsx = create_xlsx_with_sheet_content(&sheet);

    let workbook = common::load_xlsx(&xlsx);

    assert_eq!(workbook.sheets.len(), 1);
    let sheet = &workbook.sheets[0];

    assert!(
        !sheet.sparkline_groups.is_empty(),
        "Should have sparkline groups"
    );
    let group = &sheet.sparkline_groups[0];

    assert_eq!(group.sparkline_type, "line");
    assert_eq!(group.sparklines.len(), 1);
    assert_eq!(group.sparklines[0].data_range, "Sheet1!A1:E1");
    assert_eq!(group.sparklines[0].location, "F1");
}

#[test]
fn test_line_sparkline_with_markers() {
    let opts = SparklineOptions {
        markers: true,
        ..Default::default()
    };
    let group = sparkline_group_xml("line", "Sheet1!A1:E1", "F1", None, Some(&opts));
    let sheet = sheet_with_sparklines(&[group]);
    let xlsx = create_xlsx_with_sheet_content(&sheet);

    let workbook = common::load_xlsx(&xlsx);
    let group = &workbook.sheets[0].sparkline_groups[0];

    assert!(group.markers, "Markers should be enabled");
}

#[test]
fn test_line_sparkline_with_all_point_markers() {
    let opts = SparklineOptions {
        markers: true,
        high: true,
        low: true,
        first: true,
        last: true,
        negative: true,
        ..Default::default()
    };
    let group = sparkline_group_xml("line", "Sheet1!A1:E1", "F1", None, Some(&opts));
    let sheet = sheet_with_sparklines(&[group]);
    let xlsx = create_xlsx_with_sheet_content(&sheet);

    let workbook = common::load_xlsx(&xlsx);
    let group = &workbook.sheets[0].sparkline_groups[0];

    assert!(group.markers, "Markers should be enabled");
    assert!(group.high_point, "High point should be enabled");
    assert!(group.low_point, "Low point should be enabled");
    assert!(group.first_point, "First point should be enabled");
    assert!(group.last_point, "Last point should be enabled");
    assert!(group.negative_points, "Negative points should be enabled");
}

// =============================================================================
// Tests: Column Sparklines
// =============================================================================

#[test]
fn test_column_sparkline_basic() {
    let group = sparkline_group_xml("column", "Sheet1!A1:E1", "F1", None, None);
    let sheet = sheet_with_sparklines(&[group]);
    let xlsx = create_xlsx_with_sheet_content(&sheet);

    let workbook = common::load_xlsx(&xlsx);
    let group = &workbook.sheets[0].sparkline_groups[0];

    assert_eq!(group.sparkline_type, "column");
}

#[test]
fn test_column_sparkline_with_negative_values() {
    let opts = SparklineOptions {
        negative: true,
        ..Default::default()
    };
    let colors = SparklineColorConfig {
        series: Some("FF376092".to_string()),
        negative: Some("FFD00000".to_string()),
        ..Default::default()
    };
    let group = sparkline_group_xml("column", "Sheet1!A2:E2", "F2", Some(&colors), Some(&opts));
    let sheet = sheet_with_sparklines(&[group]);
    let xlsx = create_xlsx_with_sheet_content(&sheet);

    let workbook = common::load_xlsx(&xlsx);
    let group = &workbook.sheets[0].sparkline_groups[0];

    assert!(
        group.negative_points,
        "Negative points should be highlighted"
    );
    assert!(
        group.colors.negative.is_some(),
        "Negative color should be set"
    );
}

// =============================================================================
// Tests: Win/Loss Sparklines
// =============================================================================

#[test]
fn test_winloss_sparkline_basic() {
    let group = sparkline_group_xml("stacked", "Sheet1!A2:E2", "F2", None, None);
    let sheet = sheet_with_sparklines(&[group]);
    let xlsx = create_xlsx_with_sheet_content(&sheet);

    let workbook = common::load_xlsx(&xlsx);
    let group = &workbook.sheets[0].sparkline_groups[0];

    // Win/loss is represented as "stacked" type in XLSX
    assert_eq!(group.sparkline_type, "stacked");
}

#[test]
fn test_winloss_sparkline_with_colors() {
    let colors = SparklineColorConfig {
        series: Some("FF00B050".to_string()),   // Green for wins
        negative: Some("FFFF0000".to_string()), // Red for losses
        ..Default::default()
    };
    let group = sparkline_group_xml("stacked", "Sheet1!A2:E2", "F2", Some(&colors), None);
    let sheet = sheet_with_sparklines(&[group]);
    let xlsx = create_xlsx_with_sheet_content(&sheet);

    let workbook = common::load_xlsx(&xlsx);
    let group = &workbook.sheets[0].sparkline_groups[0];

    assert_eq!(group.colors.series.as_deref(), Some("#00B050"));
    assert_eq!(group.colors.negative.as_deref(), Some("#FF0000"));
}

// =============================================================================
// Tests: Sparkline Colors
// =============================================================================

#[test]
fn test_sparkline_series_color() {
    let colors = SparklineColorConfig {
        series: Some("FF4472C4".to_string()),
        ..Default::default()
    };
    let group = sparkline_group_xml("line", "Sheet1!A1:E1", "F1", Some(&colors), None);
    let sheet = sheet_with_sparklines(&[group]);
    let xlsx = create_xlsx_with_sheet_content(&sheet);

    let workbook = common::load_xlsx(&xlsx);
    let group = &workbook.sheets[0].sparkline_groups[0];

    assert_eq!(group.colors.series.as_deref(), Some("#4472C4"));
}

#[test]
fn test_sparkline_negative_color() {
    let colors = SparklineColorConfig {
        negative: Some("FFD00000".to_string()),
        ..Default::default()
    };
    let opts = SparklineOptions {
        negative: true,
        ..Default::default()
    };
    let group = sparkline_group_xml("column", "Sheet1!A2:E2", "F2", Some(&colors), Some(&opts));
    let sheet = sheet_with_sparklines(&[group]);
    let xlsx = create_xlsx_with_sheet_content(&sheet);

    let workbook = common::load_xlsx(&xlsx);
    let group = &workbook.sheets[0].sparkline_groups[0];

    assert_eq!(group.colors.negative.as_deref(), Some("#D00000"));
}

#[test]
fn test_sparkline_markers_color() {
    let colors = SparklineColorConfig {
        markers: Some("FFED7D31".to_string()),
        ..Default::default()
    };
    let opts = SparklineOptions {
        markers: true,
        ..Default::default()
    };
    let group = sparkline_group_xml("line", "Sheet1!A1:E1", "F1", Some(&colors), Some(&opts));
    let sheet = sheet_with_sparklines(&[group]);
    let xlsx = create_xlsx_with_sheet_content(&sheet);

    let workbook = common::load_xlsx(&xlsx);
    let group = &workbook.sheets[0].sparkline_groups[0];

    assert_eq!(group.colors.markers.as_deref(), Some("#ED7D31"));
}

#[test]
fn test_sparkline_first_point_color() {
    let colors = SparklineColorConfig {
        first: Some("FF70AD47".to_string()),
        ..Default::default()
    };
    let opts = SparklineOptions {
        first: true,
        ..Default::default()
    };
    let group = sparkline_group_xml("line", "Sheet1!A1:E1", "F1", Some(&colors), Some(&opts));
    let sheet = sheet_with_sparklines(&[group]);
    let xlsx = create_xlsx_with_sheet_content(&sheet);

    let workbook = common::load_xlsx(&xlsx);
    let group = &workbook.sheets[0].sparkline_groups[0];

    assert_eq!(group.colors.first.as_deref(), Some("#70AD47"));
}

#[test]
fn test_sparkline_last_point_color() {
    let colors = SparklineColorConfig {
        last: Some("FF5B9BD5".to_string()),
        ..Default::default()
    };
    let opts = SparklineOptions {
        last: true,
        ..Default::default()
    };
    let group = sparkline_group_xml("line", "Sheet1!A1:E1", "F1", Some(&colors), Some(&opts));
    let sheet = sheet_with_sparklines(&[group]);
    let xlsx = create_xlsx_with_sheet_content(&sheet);

    let workbook = common::load_xlsx(&xlsx);
    let group = &workbook.sheets[0].sparkline_groups[0];

    assert_eq!(group.colors.last.as_deref(), Some("#5B9BD5"));
}

#[test]
fn test_sparkline_high_point_color() {
    let colors = SparklineColorConfig {
        high: Some("FFFFC000".to_string()),
        ..Default::default()
    };
    let opts = SparklineOptions {
        high: true,
        ..Default::default()
    };
    let group = sparkline_group_xml("line", "Sheet1!A1:E1", "F1", Some(&colors), Some(&opts));
    let sheet = sheet_with_sparklines(&[group]);
    let xlsx = create_xlsx_with_sheet_content(&sheet);

    let workbook = common::load_xlsx(&xlsx);
    let group = &workbook.sheets[0].sparkline_groups[0];

    assert_eq!(group.colors.high.as_deref(), Some("#FFC000"));
}

#[test]
fn test_sparkline_low_point_color() {
    let colors = SparklineColorConfig {
        low: Some("FF7030A0".to_string()),
        ..Default::default()
    };
    let opts = SparklineOptions {
        low: true,
        ..Default::default()
    };
    let group = sparkline_group_xml("line", "Sheet1!A1:E1", "F1", Some(&colors), Some(&opts));
    let sheet = sheet_with_sparklines(&[group]);
    let xlsx = create_xlsx_with_sheet_content(&sheet);

    let workbook = common::load_xlsx(&xlsx);
    let group = &workbook.sheets[0].sparkline_groups[0];

    assert_eq!(group.colors.low.as_deref(), Some("#7030A0"));
}

#[test]
fn test_sparkline_all_colors() {
    let colors = SparklineColorConfig {
        series: Some("FF4472C4".to_string()),
        negative: Some("FFD00000".to_string()),
        axis: Some("FF000000".to_string()),
        markers: Some("FFED7D31".to_string()),
        first: Some("FF70AD47".to_string()),
        last: Some("FF5B9BD5".to_string()),
        high: Some("FFFFC000".to_string()),
        low: Some("FF7030A0".to_string()),
    };
    let opts = SparklineOptions {
        markers: true,
        first: true,
        last: true,
        high: true,
        low: true,
        negative: true,
        display_x_axis: true,
        ..Default::default()
    };
    let group = sparkline_group_xml("line", "Sheet1!A1:E1", "F1", Some(&colors), Some(&opts));
    let sheet = sheet_with_sparklines(&[group]);
    let xlsx = create_xlsx_with_sheet_content(&sheet);

    let workbook = common::load_xlsx(&xlsx);
    let group = &workbook.sheets[0].sparkline_groups[0];

    assert_eq!(group.colors.series.as_deref(), Some("#4472C4"));
    assert_eq!(group.colors.negative.as_deref(), Some("#D00000"));
    assert_eq!(group.colors.axis.as_deref(), Some("#000000"));
    assert_eq!(group.colors.markers.as_deref(), Some("#ED7D31"));
    assert_eq!(group.colors.first.as_deref(), Some("#70AD47"));
    assert_eq!(group.colors.last.as_deref(), Some("#5B9BD5"));
    assert_eq!(group.colors.high.as_deref(), Some("#FFC000"));
    assert_eq!(group.colors.low.as_deref(), Some("#7030A0"));
}

// =============================================================================
// Tests: Multiple Sparklines in a Group
// =============================================================================

#[test]
fn test_multiple_sparklines_in_group() {
    let sparklines = vec![
        ("Sheet1!A1:E1", "F1"),
        ("Sheet1!A2:E2", "F2"),
        ("Sheet1!A3:E3", "F3"),
    ];
    let group = sparkline_group_multi_xml("line", &sparklines, None);
    let sheet = sheet_with_sparklines(&[group]);
    let xlsx = create_xlsx_with_sheet_content(&sheet);

    let workbook = common::load_xlsx(&xlsx);
    let group = &workbook.sheets[0].sparkline_groups[0];

    assert_eq!(group.sparklines.len(), 3);
    assert_eq!(group.sparklines[0].location, "F1");
    assert_eq!(group.sparklines[1].location, "F2");
    assert_eq!(group.sparklines[2].location, "F3");
}

#[test]
fn test_multiple_sparkline_groups() {
    let group1 = sparkline_group_xml("line", "Sheet1!A1:E1", "F1", None, None);
    let group2 = sparkline_group_xml("column", "Sheet1!A2:E2", "F2", None, None);
    let sheet = sheet_with_sparklines(&[group1, group2]);
    let xlsx = create_xlsx_with_sheet_content(&sheet);

    let workbook = common::load_xlsx(&xlsx);

    assert_eq!(workbook.sheets[0].sparkline_groups.len(), 2);
    assert_eq!(
        workbook.sheets[0].sparkline_groups[0].sparkline_type,
        "line"
    );
    assert_eq!(
        workbook.sheets[0].sparkline_groups[1].sparkline_type,
        "column"
    );
}

#[test]
fn test_sparklines_shared_colors_in_group() {
    let colors = SparklineColorConfig {
        series: Some("FF4472C4".to_string()),
        ..Default::default()
    };
    let sparklines = vec![("Sheet1!A1:E1", "F1"), ("Sheet1!A2:E2", "F2")];
    let group = sparkline_group_multi_xml("column", &sparklines, Some(&colors));
    let sheet = sheet_with_sparklines(&[group]);
    let xlsx = create_xlsx_with_sheet_content(&sheet);

    let workbook = common::load_xlsx(&xlsx);
    let group = &workbook.sheets[0].sparkline_groups[0];

    // All sparklines in the group share the same color settings
    assert_eq!(group.colors.series.as_deref(), Some("#4472C4"));
    assert_eq!(group.sparklines.len(), 2);
}

// =============================================================================
// Tests: Empty Cell Handling
// =============================================================================

#[test]
fn test_empty_cells_as_gaps() {
    let opts = SparklineOptions {
        display_empty_cells_as: Some("gap".to_string()),
        ..Default::default()
    };
    let group = sparkline_group_xml("line", "Sheet1!A1:E1", "F1", None, Some(&opts));
    let sheet = sheet_with_sparklines(&[group]);
    let xlsx = create_xlsx_with_sheet_content(&sheet);

    let workbook = common::load_xlsx(&xlsx);
    let group = &workbook.sheets[0].sparkline_groups[0];

    assert_eq!(group.display_empty_cells_as.as_deref(), Some("gap"));
}

#[test]
fn test_empty_cells_as_zero() {
    let opts = SparklineOptions {
        display_empty_cells_as: Some("zero".to_string()),
        ..Default::default()
    };
    let group = sparkline_group_xml("line", "Sheet1!A1:E1", "F1", None, Some(&opts));
    let sheet = sheet_with_sparklines(&[group]);
    let xlsx = create_xlsx_with_sheet_content(&sheet);

    let workbook = common::load_xlsx(&xlsx);
    let group = &workbook.sheets[0].sparkline_groups[0];

    assert_eq!(group.display_empty_cells_as.as_deref(), Some("zero"));
}

#[test]
fn test_empty_cells_as_span() {
    let opts = SparklineOptions {
        display_empty_cells_as: Some("connect".to_string()),
        ..Default::default()
    };
    let group = sparkline_group_xml("line", "Sheet1!A1:E1", "F1", None, Some(&opts));
    let sheet = sheet_with_sparklines(&[group]);
    let xlsx = create_xlsx_with_sheet_content(&sheet);

    let workbook = common::load_xlsx(&xlsx);
    let group = &workbook.sheets[0].sparkline_groups[0];

    assert_eq!(group.display_empty_cells_as.as_deref(), Some("connect"));
}

// =============================================================================
// Tests: Axis Settings
// =============================================================================

#[test]
fn test_display_x_axis() {
    let opts = SparklineOptions {
        display_x_axis: true,
        ..Default::default()
    };
    let group = sparkline_group_xml("line", "Sheet1!A1:E1", "F1", None, Some(&opts));
    let sheet = sheet_with_sparklines(&[group]);
    let xlsx = create_xlsx_with_sheet_content(&sheet);

    let workbook = common::load_xlsx(&xlsx);
    let group = &workbook.sheets[0].sparkline_groups[0];

    assert!(group.display_x_axis, "X axis should be displayed");
}

#[test]
fn test_axis_color() {
    let colors = SparklineColorConfig {
        axis: Some("FF000000".to_string()),
        ..Default::default()
    };
    let opts = SparklineOptions {
        display_x_axis: true,
        ..Default::default()
    };
    let group = sparkline_group_xml("line", "Sheet1!A1:E1", "F1", Some(&colors), Some(&opts));
    let sheet = sheet_with_sparklines(&[group]);
    let xlsx = create_xlsx_with_sheet_content(&sheet);

    let workbook = common::load_xlsx(&xlsx);
    let group = &workbook.sheets[0].sparkline_groups[0];

    assert_eq!(group.colors.axis.as_deref(), Some("#000000"));
}

#[test]
fn test_min_axis_type_individual() {
    let opts = SparklineOptions {
        min_axis_type: Some("individual".to_string()),
        ..Default::default()
    };
    let group = sparkline_group_xml("line", "Sheet1!A1:E1", "F1", None, Some(&opts));
    let sheet = sheet_with_sparklines(&[group]);
    let xlsx = create_xlsx_with_sheet_content(&sheet);

    let workbook = common::load_xlsx(&xlsx);
    let group = &workbook.sheets[0].sparkline_groups[0];

    assert_eq!(group.min_axis_type.as_deref(), Some("individual"));
}

#[test]
fn test_max_axis_type_group() {
    let opts = SparklineOptions {
        max_axis_type: Some("group".to_string()),
        ..Default::default()
    };
    let group = sparkline_group_xml("line", "Sheet1!A1:E1", "F1", None, Some(&opts));
    let sheet = sheet_with_sparklines(&[group]);
    let xlsx = create_xlsx_with_sheet_content(&sheet);

    let workbook = common::load_xlsx(&xlsx);
    let group = &workbook.sheets[0].sparkline_groups[0];

    assert_eq!(group.max_axis_type.as_deref(), Some("group"));
}

#[test]
fn test_manual_min_max() {
    let opts = SparklineOptions {
        min_axis_type: Some("custom".to_string()),
        max_axis_type: Some("custom".to_string()),
        manual_min: Some(-10.0),
        manual_max: Some(100.0),
        ..Default::default()
    };
    let group = sparkline_group_xml("line", "Sheet1!A1:E1", "F1", None, Some(&opts));
    let sheet = sheet_with_sparklines(&[group]);
    let xlsx = create_xlsx_with_sheet_content(&sheet);

    let workbook = common::load_xlsx(&xlsx);
    let group = &workbook.sheets[0].sparkline_groups[0];

    assert_eq!(group.manual_min, Some(-10.0));
    assert_eq!(group.manual_max, Some(100.0));
}

// =============================================================================
// Tests: Right-to-Left Sparklines
// =============================================================================

#[test]
fn test_right_to_left_sparkline() {
    let opts = SparklineOptions {
        right_to_left: true,
        ..Default::default()
    };
    let group = sparkline_group_xml("line", "Sheet1!A1:E1", "F1", None, Some(&opts));
    let sheet = sheet_with_sparklines(&[group]);
    let xlsx = create_xlsx_with_sheet_content(&sheet);

    let workbook = common::load_xlsx(&xlsx);
    let group = &workbook.sheets[0].sparkline_groups[0];

    assert!(group.right_to_left, "Sparkline should be right-to-left");
}

#[test]
fn test_left_to_right_sparkline_default() {
    let group = sparkline_group_xml("line", "Sheet1!A1:E1", "F1", None, None);
    let sheet = sheet_with_sparklines(&[group]);
    let xlsx = create_xlsx_with_sheet_content(&sheet);

    let workbook = common::load_xlsx(&xlsx);
    let group = &workbook.sheets[0].sparkline_groups[0];

    assert!(!group.right_to_left, "Default should be left-to-right");
}

// =============================================================================
// Tests: Edge Cases
// =============================================================================

#[test]
fn test_sheet_without_sparklines() {
    let sheet = sheet_without_sparklines();
    let xlsx = create_xlsx_with_sheet_content(&sheet);

    let workbook = common::load_xlsx(&xlsx);

    assert!(
        workbook.sheets[0].sparkline_groups.is_empty(),
        "Sheet should have no sparkline groups"
    );
}

#[test]
fn test_sparkline_with_cross_sheet_reference() {
    // Data is on a different sheet than the sparkline display location
    let group = sparkline_group_xml("line", "Data!A1:A10", "B1", None, None);
    let sheet = sheet_with_sparklines(&[group]);
    let xlsx = create_xlsx_with_sheet_content(&sheet);

    let workbook = common::load_xlsx(&xlsx);
    let group = &workbook.sheets[0].sparkline_groups[0];

    assert_eq!(group.sparklines[0].data_range, "Data!A1:A10");
}

#[test]
fn test_sparkline_with_column_range() {
    // Vertical data range (column) instead of row
    let group = sparkline_group_xml("line", "Sheet1!A1:A10", "B1", None, None);
    let sheet = sheet_with_sparklines(&[group]);
    let xlsx = create_xlsx_with_sheet_content(&sheet);

    let workbook = common::load_xlsx(&xlsx);
    let group = &workbook.sheets[0].sparkline_groups[0];

    assert_eq!(group.sparklines[0].data_range, "Sheet1!A1:A10");
}

#[test]
fn test_sparkline_json_serialization() {
    let colors = SparklineColorConfig {
        series: Some("FF4472C4".to_string()),
        ..Default::default()
    };
    let group = sparkline_group_xml("line", "Sheet1!A1:E1", "F1", Some(&colors), None);
    let sheet = sheet_with_sparklines(&[group]);
    let xlsx = create_xlsx_with_sheet_content(&sheet);

    let workbook = common::load_xlsx(&xlsx);

    // Serialize sparkline groups to JSON and verify structure
    let sparkline_groups = &workbook.sheets[0].sparkline_groups;
    let json = serde_json::to_value(sparkline_groups).expect("Failed to serialize");

    assert!(json.is_array(), "sparklineGroups should be an array");
    assert_eq!(json.as_array().unwrap().len(), 1);

    let group = &json[0];
    assert_eq!(group["sparklineType"], "line");
    assert!(group["sparklines"].is_array());
}

// =============================================================================
// Tests: Complex Scenarios
// =============================================================================

#[test]
fn test_line_sparkline_full_configuration() {
    let colors = SparklineColorConfig {
        series: Some("FF4472C4".to_string()),
        negative: Some("FFD00000".to_string()),
        axis: Some("FF000000".to_string()),
        markers: Some("FFED7D31".to_string()),
        first: Some("FF70AD47".to_string()),
        last: Some("FF5B9BD5".to_string()),
        high: Some("FFFFC000".to_string()),
        low: Some("FF7030A0".to_string()),
    };
    let opts = SparklineOptions {
        display_empty_cells_as: Some("gap".to_string()),
        markers: true,
        high: true,
        low: true,
        first: true,
        last: true,
        negative: true,
        display_x_axis: true,
        right_to_left: false,
        min_axis_type: Some("individual".to_string()),
        max_axis_type: Some("group".to_string()),
        manual_min: None,
        manual_max: None,
    };
    let group = sparkline_group_xml("line", "Sheet1!A1:E1", "F1", Some(&colors), Some(&opts));
    let sheet = sheet_with_sparklines(&[group]);
    let xlsx = create_xlsx_with_sheet_content(&sheet);

    let workbook = common::load_xlsx(&xlsx);
    let group = &workbook.sheets[0].sparkline_groups[0];

    // Verify all settings
    assert_eq!(group.sparkline_type, "line");
    assert_eq!(group.display_empty_cells_as.as_deref(), Some("gap"));
    assert!(group.markers);
    assert!(group.high_point);
    assert!(group.low_point);
    assert!(group.first_point);
    assert!(group.last_point);
    assert!(group.negative_points);
    assert!(group.display_x_axis);
    assert!(!group.right_to_left);
    assert_eq!(group.min_axis_type.as_deref(), Some("individual"));
    assert_eq!(group.max_axis_type.as_deref(), Some("group"));

    // Verify all colors
    assert!(group.colors.series.is_some());
    assert!(group.colors.negative.is_some());
    assert!(group.colors.axis.is_some());
    assert!(group.colors.markers.is_some());
    assert!(group.colors.first.is_some());
    assert!(group.colors.last.is_some());
    assert!(group.colors.high.is_some());
    assert!(group.colors.low.is_some());
}

#[test]
fn test_mixed_sparkline_types_in_sheet() {
    let line_group = sparkline_group_xml("line", "Sheet1!A1:E1", "F1", None, None);
    let column_group = sparkline_group_xml("column", "Sheet1!A2:E2", "F2", None, None);
    let winloss_group = sparkline_group_xml("stacked", "Sheet1!A3:E3", "F3", None, None);

    let sheet = sheet_with_sparklines(&[line_group, column_group, winloss_group]);
    let xlsx = create_xlsx_with_sheet_content(&sheet);

    let workbook = common::load_xlsx(&xlsx);

    assert_eq!(workbook.sheets[0].sparkline_groups.len(), 3);
    assert_eq!(
        workbook.sheets[0].sparkline_groups[0].sparkline_type,
        "line"
    );
    assert_eq!(
        workbook.sheets[0].sparkline_groups[1].sparkline_type,
        "column"
    );
    assert_eq!(
        workbook.sheets[0].sparkline_groups[2].sparkline_type,
        "stacked"
    );
}

#[test]
fn test_large_sparkline_group() {
    // Create a group with many sparklines
    let sparklines: Vec<(&str, &str)> = (1..=20)
        .map(|i| match i {
            1 => ("Sheet1!A1:E1", "F1"),
            2 => ("Sheet1!A2:E2", "F2"),
            3 => ("Sheet1!A3:E3", "F3"),
            4 => ("Sheet1!A4:E4", "F4"),
            5 => ("Sheet1!A5:E5", "F5"),
            6 => ("Sheet1!A6:E6", "F6"),
            7 => ("Sheet1!A7:E7", "F7"),
            8 => ("Sheet1!A8:E8", "F8"),
            9 => ("Sheet1!A9:E9", "F9"),
            10 => ("Sheet1!A10:E10", "F10"),
            11 => ("Sheet1!A11:E11", "F11"),
            12 => ("Sheet1!A12:E12", "F12"),
            13 => ("Sheet1!A13:E13", "F13"),
            14 => ("Sheet1!A14:E14", "F14"),
            15 => ("Sheet1!A15:E15", "F15"),
            16 => ("Sheet1!A16:E16", "F16"),
            17 => ("Sheet1!A17:E17", "F17"),
            18 => ("Sheet1!A18:E18", "F18"),
            19 => ("Sheet1!A19:E19", "F19"),
            _ => ("Sheet1!A20:E20", "F20"),
        })
        .collect();

    let group = sparkline_group_multi_xml("line", &sparklines, None);
    let sheet = sheet_with_sparklines(&[group]);
    let xlsx = create_xlsx_with_sheet_content(&sheet);

    let workbook = common::load_xlsx(&xlsx);
    let group = &workbook.sheets[0].sparkline_groups[0];

    assert_eq!(group.sparklines.len(), 20);
}
