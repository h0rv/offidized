//! Tests for chart parsing in XLSX files
//!
//! Charts in XLSX files are stored in xl/charts/chartN.xml files and referenced
//! through drawings. This module tests parsing of various chart types including:
//! - Bar charts (clustered, stacked, 100% stacked)
//! - Line charts
//! - Pie charts
//! - Chart titles and legends
//! - Chart positioning within sheets
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
// Test Helpers
// =============================================================================

/// Create an XLSX with a chart
fn create_xlsx_with_chart(drawing_xml: &str, chart_xml: &str, drawing_rels_xml: &str) -> Vec<u8> {
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
<Override PartName="/xl/drawings/drawing1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>
<Override PartName="/xl/charts/chart1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>
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

    // xl/worksheets/_rels/sheet1.xml.rels
    let _ = zip.start_file("xl/worksheets/_rels/sheet1.xml.rels", options);
    let _ = zip.write_all(
        br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing1.xml"/>
</Relationships>"#,
    );

    // xl/worksheets/sheet1.xml with data for chart
    let _ = zip.start_file("xl/worksheets/sheet1.xml", options);
    let _ = zip.write_all(
        br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<sheetData>
<row r="1"><c r="A1" t="inlineStr"><is><t>Category</t></is></c><c r="B1" t="inlineStr"><is><t>Sales</t></is></c></row>
<row r="2"><c r="A2" t="inlineStr"><is><t>Q1</t></is></c><c r="B2"><v>100</v></c></row>
<row r="3"><c r="A3" t="inlineStr"><is><t>Q2</t></is></c><c r="B3"><v>150</v></c></row>
<row r="4"><c r="A4" t="inlineStr"><is><t>Q3</t></is></c><c r="B4"><v>200</v></c></row>
<row r="5"><c r="A5" t="inlineStr"><is><t>Q4</t></is></c><c r="B5"><v>175</v></c></row>
</sheetData>
<drawing r:id="rId1"/>
</worksheet>"#,
    );

    // xl/drawings/drawing1.xml
    let _ = zip.start_file("xl/drawings/drawing1.xml", options);
    let _ = zip.write_all(drawing_xml.as_bytes());

    // xl/drawings/_rels/drawing1.xml.rels
    let _ = zip.start_file("xl/drawings/_rels/drawing1.xml.rels", options);
    let _ = zip.write_all(drawing_rels_xml.as_bytes());

    // xl/charts/chart1.xml
    let _ = zip.start_file("xl/charts/chart1.xml", options);
    let _ = zip.write_all(chart_xml.as_bytes());

    let cursor = zip.finish().expect("Failed to finish ZIP");
    cursor.into_inner()
}

/// Standard drawing XML for a chart
fn standard_chart_drawing_xml() -> &'static str {
    r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
          xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
          xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
  <xdr:twoCellAnchor>
    <xdr:from><xdr:col>4</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>1</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>
    <xdr:to><xdr:col>12</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>15</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>
    <xdr:graphicFrame macro="">
      <xdr:nvGraphicFramePr>
        <xdr:cNvPr id="2" name="Chart 1"/>
        <xdr:cNvGraphicFramePr/>
      </xdr:nvGraphicFramePr>
      <xdr:xfrm>
        <a:off x="0" y="0"/>
        <a:ext cx="4572000" cy="2743200"/>
      </xdr:xfrm>
      <a:graphic>
        <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">
          <c:chart r:id="rId1"/>
        </a:graphicData>
      </a:graphic>
    </xdr:graphicFrame>
    <xdr:clientData/>
  </xdr:twoCellAnchor>
</xdr:wsDr>"#
}

/// Standard drawing rels XML
fn standard_chart_drawing_rels_xml() -> &'static str {
    r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart1.xml"/>
</Relationships>"#
}

// =============================================================================
// Tests: Bar Chart Parsing
// =============================================================================

#[test]
fn test_bar_chart_clustered() {
    let chart_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
              xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <c:chart>
    <c:title>
      <c:tx>
        <c:rich>
          <a:bodyPr/>
          <a:lstStyle/>
          <a:p>
            <a:r><a:t>Quarterly Sales</a:t></a:r>
          </a:p>
        </c:rich>
      </c:tx>
    </c:title>
    <c:plotArea>
      <c:barChart>
        <c:barDir val="col"/>
        <c:grouping val="clustered"/>
        <c:ser>
          <c:idx val="0"/>
          <c:order val="0"/>
          <c:tx>
            <c:strRef>
              <c:f>Sheet1!$B$1</c:f>
            </c:strRef>
          </c:tx>
          <c:cat>
            <c:strRef>
              <c:f>Sheet1!$A$2:$A$5</c:f>
            </c:strRef>
          </c:cat>
          <c:val>
            <c:numRef>
              <c:f>Sheet1!$B$2:$B$5</c:f>
            </c:numRef>
          </c:val>
        </c:ser>
        <c:axId val="1"/>
        <c:axId val="2"/>
      </c:barChart>
      <c:catAx>
        <c:axId val="1"/>
        <c:scaling><c:orientation val="minMax"/></c:scaling>
        <c:axPos val="b"/>
        <c:title>
          <c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>Quarter</a:t></a:r></a:p></c:rich></c:tx>
        </c:title>
      </c:catAx>
      <c:valAx>
        <c:axId val="2"/>
        <c:scaling><c:orientation val="minMax"/></c:scaling>
        <c:axPos val="l"/>
        <c:title>
          <c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>Sales ($)</a:t></a:r></a:p></c:rich></c:tx>
        </c:title>
      </c:valAx>
    </c:plotArea>
    <c:legend>
      <c:legendPos val="r"/>
    </c:legend>
  </c:chart>
</c:chartSpace>"#;

    let xlsx = create_xlsx_with_chart(
        standard_chart_drawing_xml(),
        chart_xml,
        standard_chart_drawing_rels_xml(),
    );
    let workbook = common::load_xlsx(&xlsx);

    let sheet = &workbook.sheets[0];
    assert!(!sheet.drawings.is_empty(), "Should have drawings");

    // Find the chart drawing
    let chart_drawing = sheet.drawings.iter().find(|d| d.drawing_type == "chart");
    assert!(chart_drawing.is_some(), "Should have a chart drawing");

    let chart = chart_drawing.unwrap();
    assert_eq!(chart.name.as_deref(), Some("Chart 1"));

    // Check chart is linked
    assert!(chart.chart_id.is_some());
}

#[test]
fn test_bar_chart_stacked() {
    let chart_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
              xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <c:chart>
    <c:title>
      <c:tx>
        <c:rich>
          <a:bodyPr/><a:lstStyle/>
          <a:p><a:r><a:t>Stacked Bar Chart</a:t></a:r></a:p>
        </c:rich>
      </c:tx>
    </c:title>
    <c:plotArea>
      <c:barChart>
        <c:barDir val="col"/>
        <c:grouping val="stacked"/>
        <c:ser>
          <c:idx val="0"/>
          <c:order val="0"/>
          <c:val>
            <c:numRef>
              <c:f>Sheet1!$B$2:$B$5</c:f>
            </c:numRef>
          </c:val>
        </c:ser>
        <c:axId val="1"/>
        <c:axId val="2"/>
      </c:barChart>
    </c:plotArea>
  </c:chart>
</c:chartSpace>"#;

    let xlsx = create_xlsx_with_chart(
        standard_chart_drawing_xml(),
        chart_xml,
        standard_chart_drawing_rels_xml(),
    );
    let workbook = common::load_xlsx(&xlsx);

    let sheet = &workbook.sheets[0];
    let chart = sheet.drawings.iter().find(|d| d.drawing_type == "chart");
    assert!(chart.is_some());
}

#[test]
fn test_bar_chart_horizontal() {
    let chart_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
              xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <c:chart>
    <c:plotArea>
      <c:barChart>
        <c:barDir val="bar"/>
        <c:grouping val="clustered"/>
        <c:ser>
          <c:idx val="0"/>
          <c:order val="0"/>
        </c:ser>
      </c:barChart>
    </c:plotArea>
  </c:chart>
</c:chartSpace>"#;

    let xlsx = create_xlsx_with_chart(
        standard_chart_drawing_xml(),
        chart_xml,
        standard_chart_drawing_rels_xml(),
    );
    let workbook = common::load_xlsx(&xlsx);

    let sheet = &workbook.sheets[0];
    assert!(!sheet.drawings.is_empty());
}

// =============================================================================
// Tests: Line Chart Parsing
// =============================================================================

#[test]
fn test_line_chart_basic() {
    let chart_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
              xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <c:chart>
    <c:title>
      <c:tx>
        <c:rich>
          <a:bodyPr/><a:lstStyle/>
          <a:p><a:r><a:t>Sales Trend</a:t></a:r></a:p>
        </c:rich>
      </c:tx>
    </c:title>
    <c:plotArea>
      <c:lineChart>
        <c:grouping val="standard"/>
        <c:ser>
          <c:idx val="0"/>
          <c:order val="0"/>
          <c:tx>
            <c:strRef>
              <c:f>Sheet1!$B$1</c:f>
            </c:strRef>
          </c:tx>
          <c:cat>
            <c:strRef>
              <c:f>Sheet1!$A$2:$A$5</c:f>
            </c:strRef>
          </c:cat>
          <c:val>
            <c:numRef>
              <c:f>Sheet1!$B$2:$B$5</c:f>
            </c:numRef>
          </c:val>
          <c:marker>
            <c:symbol val="circle"/>
            <c:size val="5"/>
          </c:marker>
        </c:ser>
        <c:axId val="1"/>
        <c:axId val="2"/>
      </c:lineChart>
      <c:catAx>
        <c:axId val="1"/>
        <c:scaling><c:orientation val="minMax"/></c:scaling>
        <c:axPos val="b"/>
      </c:catAx>
      <c:valAx>
        <c:axId val="2"/>
        <c:scaling><c:orientation val="minMax"/></c:scaling>
        <c:axPos val="l"/>
      </c:valAx>
    </c:plotArea>
    <c:legend>
      <c:legendPos val="b"/>
    </c:legend>
  </c:chart>
</c:chartSpace>"#;

    let xlsx = create_xlsx_with_chart(
        standard_chart_drawing_xml(),
        chart_xml,
        standard_chart_drawing_rels_xml(),
    );
    let workbook = common::load_xlsx(&xlsx);

    let sheet = &workbook.sheets[0];
    let chart = sheet.drawings.iter().find(|d| d.drawing_type == "chart");
    assert!(chart.is_some());
}

#[test]
fn test_line_chart_with_markers() {
    let chart_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
              xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <c:chart>
    <c:plotArea>
      <c:lineChart>
        <c:grouping val="standard"/>
        <c:ser>
          <c:idx val="0"/>
          <c:order val="0"/>
          <c:marker>
            <c:symbol val="diamond"/>
            <c:size val="7"/>
          </c:marker>
        </c:ser>
        <c:marker val="1"/>
      </c:lineChart>
    </c:plotArea>
  </c:chart>
</c:chartSpace>"#;

    let xlsx = create_xlsx_with_chart(
        standard_chart_drawing_xml(),
        chart_xml,
        standard_chart_drawing_rels_xml(),
    );
    let workbook = common::load_xlsx(&xlsx);

    let sheet = &workbook.sheets[0];
    assert!(!sheet.drawings.is_empty());
}

// =============================================================================
// Tests: Pie Chart Parsing
// =============================================================================

#[test]
fn test_pie_chart_basic() {
    let chart_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
              xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <c:chart>
    <c:title>
      <c:tx>
        <c:rich>
          <a:bodyPr/><a:lstStyle/>
          <a:p><a:r><a:t>Market Share</a:t></a:r></a:p>
        </c:rich>
      </c:tx>
    </c:title>
    <c:plotArea>
      <c:pieChart>
        <c:ser>
          <c:idx val="0"/>
          <c:order val="0"/>
          <c:tx>
            <c:strRef>
              <c:f>Sheet1!$B$1</c:f>
            </c:strRef>
          </c:tx>
          <c:cat>
            <c:strRef>
              <c:f>Sheet1!$A$2:$A$5</c:f>
            </c:strRef>
          </c:cat>
          <c:val>
            <c:numRef>
              <c:f>Sheet1!$B$2:$B$5</c:f>
            </c:numRef>
          </c:val>
        </c:ser>
      </c:pieChart>
    </c:plotArea>
    <c:legend>
      <c:legendPos val="r"/>
    </c:legend>
  </c:chart>
</c:chartSpace>"#;

    let xlsx = create_xlsx_with_chart(
        standard_chart_drawing_xml(),
        chart_xml,
        standard_chart_drawing_rels_xml(),
    );
    let workbook = common::load_xlsx(&xlsx);

    let sheet = &workbook.sheets[0];
    let chart = sheet.drawings.iter().find(|d| d.drawing_type == "chart");
    assert!(chart.is_some());
}

#[test]
fn test_pie_chart_3d() {
    let chart_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
              xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <c:chart>
    <c:title>
      <c:tx>
        <c:rich>
          <a:bodyPr/><a:lstStyle/>
          <a:p><a:r><a:t>3D Pie Chart</a:t></a:r></a:p>
        </c:rich>
      </c:tx>
    </c:title>
    <c:view3D>
      <c:rotX val="30"/>
      <c:rotY val="0"/>
      <c:rAngAx val="0"/>
      <c:perspective val="30"/>
    </c:view3D>
    <c:plotArea>
      <c:pie3DChart>
        <c:ser>
          <c:idx val="0"/>
          <c:order val="0"/>
        </c:ser>
      </c:pie3DChart>
    </c:plotArea>
  </c:chart>
</c:chartSpace>"#;

    let xlsx = create_xlsx_with_chart(
        standard_chart_drawing_xml(),
        chart_xml,
        standard_chart_drawing_rels_xml(),
    );
    let workbook = common::load_xlsx(&xlsx);

    let sheet = &workbook.sheets[0];
    assert!(!sheet.drawings.is_empty());
}

#[test]
fn test_doughnut_chart() {
    let chart_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
              xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <c:chart>
    <c:plotArea>
      <c:doughnutChart>
        <c:ser>
          <c:idx val="0"/>
          <c:order val="0"/>
        </c:ser>
        <c:holeSize val="50"/>
      </c:doughnutChart>
    </c:plotArea>
  </c:chart>
</c:chartSpace>"#;

    let xlsx = create_xlsx_with_chart(
        standard_chart_drawing_xml(),
        chart_xml,
        standard_chart_drawing_rels_xml(),
    );
    let workbook = common::load_xlsx(&xlsx);

    let sheet = &workbook.sheets[0];
    assert!(!sheet.drawings.is_empty());
}

// =============================================================================
// Tests: Chart Titles and Legends
// =============================================================================

#[test]
fn test_chart_with_title() {
    let chart_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
              xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <c:chart>
    <c:title>
      <c:tx>
        <c:rich>
          <a:bodyPr/>
          <a:lstStyle/>
          <a:p>
            <a:pPr>
              <a:defRPr sz="1400" b="1"/>
            </a:pPr>
            <a:r>
              <a:rPr lang="en-US" sz="1400" b="1"/>
              <a:t>Annual Revenue Report</a:t>
            </a:r>
          </a:p>
        </c:rich>
      </c:tx>
      <c:layout/>
      <c:overlay val="0"/>
    </c:title>
    <c:plotArea>
      <c:barChart>
        <c:barDir val="col"/>
        <c:grouping val="clustered"/>
        <c:ser><c:idx val="0"/><c:order val="0"/></c:ser>
      </c:barChart>
    </c:plotArea>
  </c:chart>
</c:chartSpace>"#;

    let xlsx = create_xlsx_with_chart(
        standard_chart_drawing_xml(),
        chart_xml,
        standard_chart_drawing_rels_xml(),
    );
    let workbook = common::load_xlsx(&xlsx);

    let sheet = &workbook.sheets[0];
    assert!(!sheet.drawings.is_empty());
}

#[test]
fn test_chart_legend_positions() {
    // Test legend at bottom
    let chart_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
              xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <c:chart>
    <c:plotArea>
      <c:barChart>
        <c:barDir val="col"/>
        <c:grouping val="clustered"/>
        <c:ser><c:idx val="0"/><c:order val="0"/></c:ser>
      </c:barChart>
    </c:plotArea>
    <c:legend>
      <c:legendPos val="b"/>
      <c:layout/>
      <c:overlay val="0"/>
    </c:legend>
  </c:chart>
</c:chartSpace>"#;

    let xlsx = create_xlsx_with_chart(
        standard_chart_drawing_xml(),
        chart_xml,
        standard_chart_drawing_rels_xml(),
    );
    let workbook = common::load_xlsx(&xlsx);
    assert!(!workbook.sheets[0].drawings.is_empty());
}

#[test]
fn test_chart_without_legend() {
    let chart_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
              xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <c:chart>
    <c:plotArea>
      <c:pieChart>
        <c:ser><c:idx val="0"/><c:order val="0"/></c:ser>
      </c:pieChart>
    </c:plotArea>
  </c:chart>
</c:chartSpace>"#;

    let xlsx = create_xlsx_with_chart(
        standard_chart_drawing_xml(),
        chart_xml,
        standard_chart_drawing_rels_xml(),
    );
    let workbook = common::load_xlsx(&xlsx);
    assert!(!workbook.sheets[0].drawings.is_empty());
}

// =============================================================================
// Tests: Chart Position
// =============================================================================

#[test]
fn test_chart_position_two_cell_anchor() {
    let drawing_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
          xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
          xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
  <xdr:twoCellAnchor>
    <xdr:from>
      <xdr:col>2</xdr:col>
      <xdr:colOff>100000</xdr:colOff>
      <xdr:row>5</xdr:row>
      <xdr:rowOff>50000</xdr:rowOff>
    </xdr:from>
    <xdr:to>
      <xdr:col>10</xdr:col>
      <xdr:colOff>200000</xdr:colOff>
      <xdr:row>20</xdr:row>
      <xdr:rowOff>100000</xdr:rowOff>
    </xdr:to>
    <xdr:graphicFrame macro="">
      <xdr:nvGraphicFramePr>
        <xdr:cNvPr id="2" name="Positioned Chart"/>
        <xdr:cNvGraphicFramePr/>
      </xdr:nvGraphicFramePr>
      <xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm>
      <a:graphic>
        <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">
          <c:chart r:id="rId1"/>
        </a:graphicData>
      </a:graphic>
    </xdr:graphicFrame>
    <xdr:clientData/>
  </xdr:twoCellAnchor>
</xdr:wsDr>"#;

    let chart_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
  <c:chart>
    <c:plotArea>
      <c:barChart>
        <c:barDir val="col"/>
        <c:grouping val="clustered"/>
        <c:ser><c:idx val="0"/><c:order val="0"/></c:ser>
      </c:barChart>
    </c:plotArea>
  </c:chart>
</c:chartSpace>"#;

    let xlsx = create_xlsx_with_chart(drawing_xml, chart_xml, standard_chart_drawing_rels_xml());
    let workbook = common::load_xlsx(&xlsx);

    let sheet = &workbook.sheets[0];
    let chart = sheet.drawings.iter().find(|d| d.drawing_type == "chart");
    assert!(chart.is_some());

    let chart = chart.unwrap();
    assert_eq!(chart.from_col, Some(2));
    assert_eq!(chart.from_row, Some(5));
    assert_eq!(chart.from_col_off, Some(100000));
    assert_eq!(chart.from_row_off, Some(50000));
    assert_eq!(chart.to_col, Some(10));
    assert_eq!(chart.to_row, Some(20));
}

// =============================================================================
// Tests: Multiple Series
// =============================================================================

#[test]
fn test_chart_multiple_series() {
    let chart_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
              xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <c:chart>
    <c:plotArea>
      <c:barChart>
        <c:barDir val="col"/>
        <c:grouping val="clustered"/>
        <c:ser>
          <c:idx val="0"/>
          <c:order val="0"/>
          <c:tx><c:v>Series 1</c:v></c:tx>
        </c:ser>
        <c:ser>
          <c:idx val="1"/>
          <c:order val="1"/>
          <c:tx><c:v>Series 2</c:v></c:tx>
        </c:ser>
        <c:ser>
          <c:idx val="2"/>
          <c:order val="2"/>
          <c:tx><c:v>Series 3</c:v></c:tx>
        </c:ser>
      </c:barChart>
    </c:plotArea>
    <c:legend>
      <c:legendPos val="r"/>
    </c:legend>
  </c:chart>
</c:chartSpace>"#;

    let xlsx = create_xlsx_with_chart(
        standard_chart_drawing_xml(),
        chart_xml,
        standard_chart_drawing_rels_xml(),
    );
    let workbook = common::load_xlsx(&xlsx);

    let sheet = &workbook.sheets[0];
    assert!(!sheet.drawings.is_empty());
}

// =============================================================================
// Tests: Other Chart Types
// =============================================================================

#[test]
fn test_area_chart() {
    let chart_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
              xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <c:chart>
    <c:plotArea>
      <c:areaChart>
        <c:grouping val="standard"/>
        <c:ser>
          <c:idx val="0"/>
          <c:order val="0"/>
        </c:ser>
      </c:areaChart>
    </c:plotArea>
  </c:chart>
</c:chartSpace>"#;

    let xlsx = create_xlsx_with_chart(
        standard_chart_drawing_xml(),
        chart_xml,
        standard_chart_drawing_rels_xml(),
    );
    let workbook = common::load_xlsx(&xlsx);
    assert!(!workbook.sheets[0].drawings.is_empty());
}

#[test]
fn test_scatter_chart() {
    let chart_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
              xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <c:chart>
    <c:plotArea>
      <c:scatterChart>
        <c:scatterStyle val="lineMarker"/>
        <c:ser>
          <c:idx val="0"/>
          <c:order val="0"/>
          <c:xVal>
            <c:numRef>
              <c:f>Sheet1!$A$2:$A$5</c:f>
            </c:numRef>
          </c:xVal>
          <c:yVal>
            <c:numRef>
              <c:f>Sheet1!$B$2:$B$5</c:f>
            </c:numRef>
          </c:yVal>
        </c:ser>
      </c:scatterChart>
    </c:plotArea>
  </c:chart>
</c:chartSpace>"#;

    let xlsx = create_xlsx_with_chart(
        standard_chart_drawing_xml(),
        chart_xml,
        standard_chart_drawing_rels_xml(),
    );
    let workbook = common::load_xlsx(&xlsx);
    assert!(!workbook.sheets[0].drawings.is_empty());
}

#[test]
fn test_radar_chart() {
    let chart_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
              xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <c:chart>
    <c:plotArea>
      <c:radarChart>
        <c:radarStyle val="marker"/>
        <c:ser>
          <c:idx val="0"/>
          <c:order val="0"/>
        </c:ser>
      </c:radarChart>
    </c:plotArea>
  </c:chart>
</c:chartSpace>"#;

    let xlsx = create_xlsx_with_chart(
        standard_chart_drawing_xml(),
        chart_xml,
        standard_chart_drawing_rels_xml(),
    );
    let workbook = common::load_xlsx(&xlsx);
    assert!(!workbook.sheets[0].drawings.is_empty());
}

// =============================================================================
// Tests: Edge Cases
// =============================================================================

#[test]
fn test_sheet_without_charts() {
    let xlsx = fixtures::XlsxBuilder::new()
        .sheet(fixtures::SheetBuilder::new("Sheet1").cell("A1", "No charts", None))
        .build();

    let workbook = common::load_xlsx(&xlsx);
    let sheet = &workbook.sheets[0];

    // No chart drawings
    let charts: Vec<_> = sheet
        .drawings
        .iter()
        .filter(|d| d.drawing_type == "chart")
        .collect();
    assert!(charts.is_empty());
}

#[test]
fn test_chart_serialization_to_json() {
    let chart_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
  <c:chart>
    <c:plotArea>
      <c:barChart>
        <c:barDir val="col"/>
        <c:grouping val="clustered"/>
        <c:ser><c:idx val="0"/><c:order val="0"/></c:ser>
      </c:barChart>
    </c:plotArea>
  </c:chart>
</c:chartSpace>"#;

    let xlsx = create_xlsx_with_chart(
        standard_chart_drawing_xml(),
        chart_xml,
        standard_chart_drawing_rels_xml(),
    );
    let workbook = common::load_xlsx(&xlsx);

    // Serialize drawings to JSON and check chart drawing
    let sheet = &workbook.sheets[0];
    assert!(!sheet.drawings.is_empty());

    let chart_drawing = sheet.drawings.iter().find(|d| d.drawing_type == "chart");
    assert!(chart_drawing.is_some());

    let json = serde_json::to_value(chart_drawing.unwrap()).expect("Failed to serialize");
    assert_eq!(json["name"], "Chart 1");
    assert!(json["chartId"].is_string());
}
