//! Integration test for chart and sparkline roundtrip preservation.

use offidized_xlsx::{
    CellValue, Chart, ChartDataRef, ChartSeries, ChartType, Sparkline, SparklineGroup,
    SparklineType, Workbook,
};
use tempfile::tempdir;

#[test]
fn chart_roundtrip_programmatic() {
    // Create a workbook with a chart programmatically
    let mut wb = Workbook::new();
    let ws = wb.add_sheet("ChartData");

    // Add data
    ws.cell_mut("A1").expect("A1").set_value("Month");
    ws.cell_mut("B1").expect("B1").set_value("Sales");
    ws.cell_mut("A2").expect("A2").set_value("Jan");
    ws.cell_mut("B2").expect("B2").set_value(100);
    ws.cell_mut("A3").expect("A3").set_value("Feb");
    ws.cell_mut("B3").expect("B3").set_value(200);
    ws.cell_mut("A4").expect("A4").set_value("Mar");
    ws.cell_mut("B4").expect("B4").set_value(150);

    // Create a chart
    let mut chart = Chart::new(ChartType::Bar);
    chart.set_title("Monthly Sales");
    chart.set_anchor(3, 1, 10, 15); // D2:K16

    let mut series = ChartSeries::new(0, 0);
    series
        .set_name_ref("Sheet1!$B$1")
        .set_categories(ChartDataRef::from_formula("Sheet1!$A$2:$A$4"))
        .set_values(ChartDataRef::from_formula("Sheet1!$B$2:$B$4"));
    chart.add_series(series);

    // Bar charts come with default axes (catAx + valAx), no need to add manually
    ws.charts_mut().push(chart);

    // Save to temp file
    let temp_dir = tempdir().expect("create temp dir");
    let file_path = temp_dir.path().join("chart_test.xlsx");
    wb.save(&file_path).expect("should save workbook");

    // Load it back
    let loaded = Workbook::open(&file_path).expect("should load workbook");
    let loaded_ws = &loaded.worksheets()[0];

    // Verify chart survived roundtrip
    assert_eq!(loaded_ws.charts().len(), 1, "should have 1 chart");
    let loaded_chart = &loaded_ws.charts()[0];

    assert_eq!(loaded_chart.chart_type(), ChartType::Bar);
    assert_eq!(loaded_chart.title(), Some("Monthly Sales"));
    assert_eq!(loaded_chart.series().len(), 1);
    assert_eq!(loaded_chart.axes().len(), 2);

    let loaded_series = &loaded_chart.series()[0];
    assert_eq!(loaded_series.name_ref(), Some("Sheet1!$B$1"));
    assert_eq!(
        loaded_series.categories().unwrap().formula(),
        Some("Sheet1!$A$2:$A$4")
    );
    assert_eq!(
        loaded_series.values().unwrap().formula(),
        Some("Sheet1!$B$2:$B$4")
    );

    // Verify data also survived
    assert_eq!(
        loaded_ws.cell("A1").expect("A1").value(),
        Some(&CellValue::String("Month".to_string()))
    );
    assert_eq!(
        loaded_ws.cell("B2").expect("B2").value(),
        Some(&CellValue::Number(100.0))
    );
}

#[test]
fn sparkline_roundtrip_programmatic() {
    let mut wb = Workbook::new();
    let ws = wb.add_sheet("SparklineData");

    // Add data
    for i in 1..=10 {
        ws.cell_mut(&format!("A{i}"))
            .expect("cell")
            .set_value(i * 10);
    }

    // Create sparkline group
    let mut group = SparklineGroup::new();
    group
        .set_sparkline_type(SparklineType::Line)
        .add_sparkline(Sparkline::new("B1", "Sheet1!$A$1:$A$10"));

    ws.sparkline_groups_mut().push(group);

    // Save to temp file
    let temp_dir = tempdir().expect("create temp dir");
    let file_path = temp_dir.path().join("sparkline_test.xlsx");
    wb.save(&file_path).expect("should save workbook");

    // Load it back
    let loaded = Workbook::open(&file_path).expect("should load workbook");
    let loaded_ws = &loaded.worksheets()[0];

    // Verify sparkline survived roundtrip
    assert_eq!(
        loaded_ws.sparkline_groups().len(),
        1,
        "should have 1 sparkline group"
    );
    let loaded_group = &loaded_ws.sparkline_groups()[0];

    assert_eq!(loaded_group.sparkline_type(), SparklineType::Line);
    assert_eq!(loaded_group.sparklines().len(), 1);

    let loaded_sparkline = &loaded_group.sparklines()[0];
    assert_eq!(loaded_sparkline.location(), "B1");
    assert_eq!(loaded_sparkline.data_range(), "Sheet1!$A$1:$A$10");

    // Verify data also survived
    assert_eq!(
        loaded_ws.cell("A5").expect("A5").value(),
        Some(&CellValue::Number(50.0))
    );
}
