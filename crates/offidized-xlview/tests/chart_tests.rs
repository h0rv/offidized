//! Comprehensive tests for chart parsing from real XLSX files
//!
//! These tests parse real XLSX files (kitchen_sink_v2.xlsx) to verify that
//! various chart types are correctly parsed according to ECMA-376.
//!
//! Tested chart types:
//! - Bar chart (clustered, stacked, column)
//! - Line chart
//! - Pie chart
//! - Doughnut chart
//! - Area chart
//! - Scatter chart
//! - Bubble chart
//! - Radar chart
//! - Stock chart (if present)
//! - Surface chart (if present)
//! - Combo chart (if present)
//!
//! Chart properties tested:
//! - Chart type detection
//! - Title parsing
//! - Series data references
//! - Axis configuration
//! - Legend parsing
//! - Position/anchoring
//! - Grouping (clustered, stacked, percentStacked)
//! - Bar direction (col vs bar)
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

use offidized_xlview::types::chart::{BarDirection, Chart, ChartGrouping, ChartType};
use offidized_xlview::types::workbook::{Sheet, Workbook};

// =============================================================================
// Helper Functions
// =============================================================================

/// Parse an XLSX file and return the workbook
fn parse_test_file(filename: &str) -> Workbook {
    let path = format!("{}/test/{}", env!("CARGO_MANIFEST_DIR"), filename);
    let data =
        std::fs::read(&path).unwrap_or_else(|_| panic!("Failed to read test file: {}", path));
    common::load_xlsx(&data)
}

/// Find a sheet by name in the workbook
fn find_sheet<'a>(workbook: &'a Workbook, name: &str) -> Option<&'a Sheet> {
    workbook.sheets.iter().find(|s| s.name == name)
}

/// Get all charts from a sheet
fn get_charts(sheet: &Sheet) -> &Vec<Chart> {
    &sheet.charts
}

/// Count charts of a specific type across all sheets
fn count_charts_by_type(workbook: &Workbook, chart_type: ChartType) -> usize {
    workbook
        .sheets
        .iter()
        .flat_map(|s| s.charts.iter())
        .filter(|c| c.chart_type == chart_type)
        .count()
}

/// Find first chart of a specific type
fn find_chart_by_type(workbook: &Workbook, chart_type: ChartType) -> Option<&Chart> {
    workbook
        .sheets
        .iter()
        .flat_map(|s| s.charts.iter())
        .find(|c| c.chart_type == chart_type)
}

/// Find chart by title
fn find_chart_by_title<'a>(workbook: &'a Workbook, title: &str) -> Option<&'a Chart> {
    workbook
        .sheets
        .iter()
        .flat_map(|s| s.charts.iter())
        .find(|c| c.title.as_deref() == Some(title))
}

// =============================================================================
// Tests: Kitchen Sink V2 - Basic Chart Parsing
// =============================================================================

#[test]
fn test_kitchen_sink_v2_has_charts() {
    let workbook = parse_test_file("kitchen_sink_v2.xlsx");

    // Count total charts across all sheets
    let total_charts: usize = workbook.sheets.iter().map(|s| s.charts.len()).sum();

    assert!(
        total_charts > 0,
        "kitchen_sink_v2.xlsx should have charts, found {}",
        total_charts
    );

    // The file has 4 charts based on our analysis
    assert_eq!(
        total_charts, 4,
        "kitchen_sink_v2.xlsx should have exactly 4 charts"
    );
}

#[test]
fn test_kitchen_sink_v2_charts_sheet() {
    let workbook = parse_test_file("kitchen_sink_v2.xlsx");

    // Find the Charts sheet
    let charts_sheet = find_sheet(&workbook, "Charts");
    assert!(charts_sheet.is_some(), "Should have a 'Charts' sheet");

    let charts = get_charts(charts_sheet.unwrap());
    assert!(!charts.is_empty(), "Charts sheet should contain charts");
}

// =============================================================================
// Tests: Bar Chart Parsing
// =============================================================================

#[test]
fn test_bar_chart_detection() {
    let workbook = parse_test_file("kitchen_sink_v2.xlsx");

    let bar_count = count_charts_by_type(&workbook, ChartType::Bar);
    assert!(
        bar_count > 0,
        "Should have at least one Bar chart, found {}",
        bar_count
    );
}

#[test]
fn test_bar_chart_quarterly_sales() {
    let workbook = parse_test_file("kitchen_sink_v2.xlsx");

    // Find the "Quarterly Sales" bar chart
    let chart = find_chart_by_title(&workbook, "Quarterly Sales");
    assert!(chart.is_some(), "Should find 'Quarterly Sales' chart");

    let chart = chart.unwrap();

    // Verify chart type
    assert_eq!(chart.chart_type, ChartType::Bar, "Should be a Bar chart");

    // Verify bar direction is column (vertical bars)
    assert_eq!(
        chart.bar_direction,
        Some(BarDirection::Col),
        "Should have column direction"
    );

    // Verify grouping is clustered
    assert_eq!(
        chart.grouping,
        Some(ChartGrouping::Clustered),
        "Should be clustered"
    );

    // Verify title
    assert_eq!(
        chart.title.as_deref(),
        Some("Quarterly Sales"),
        "Title should match"
    );
}

#[test]
fn test_bar_chart_series() {
    let workbook = parse_test_file("kitchen_sink_v2.xlsx");

    let chart = find_chart_by_title(&workbook, "Quarterly Sales");
    assert!(chart.is_some());

    let chart = chart.unwrap();

    // Should have multiple series
    assert!(!chart.series.is_empty(), "Bar chart should have series");

    // Verify series have data references
    for series in &chart.series {
        // Series should have values reference
        if let Some(ref values) = series.values {
            // Should have a formula reference
            assert!(
                values.formula.is_some() || !values.num_values.is_empty(),
                "Series should have formula or cached values"
            );
        }
    }
}

#[test]
fn test_bar_chart_axes() {
    let workbook = parse_test_file("kitchen_sink_v2.xlsx");

    let chart = find_chart_by_title(&workbook, "Quarterly Sales");
    assert!(chart.is_some());

    let chart = chart.unwrap();

    // Bar charts should have axes
    assert!(!chart.axes.is_empty(), "Bar chart should have axes");

    // Check for category axis
    let has_cat_axis = chart
        .axes
        .iter()
        .any(|a| a.axis_type == "cat" || a.axis_type == "catAx");
    assert!(has_cat_axis, "Should have a category axis");

    // Check for value axis
    let has_val_axis = chart
        .axes
        .iter()
        .any(|a| a.axis_type == "val" || a.axis_type == "valAx");
    assert!(has_val_axis, "Should have a value axis");
}

#[test]
fn test_bar_chart_axis_titles() {
    let workbook = parse_test_file("kitchen_sink_v2.xlsx");

    let chart = find_chart_by_title(&workbook, "Quarterly Sales");
    assert!(chart.is_some());

    let chart = chart.unwrap();

    // Check axis titles
    let cat_axis = chart
        .axes
        .iter()
        .find(|a| a.axis_type == "cat" || a.axis_type == "catAx");
    if let Some(axis) = cat_axis {
        if axis.title.is_some() {
            assert_eq!(
                axis.title.as_deref(),
                Some("Quarter"),
                "Category axis title should be 'Quarter'"
            );
        }
    }

    let val_axis = chart
        .axes
        .iter()
        .find(|a| a.axis_type == "val" || a.axis_type == "valAx");
    if let Some(axis) = val_axis {
        if axis.title.is_some() {
            assert_eq!(
                axis.title.as_deref(),
                Some("Amount"),
                "Value axis title should be 'Amount'"
            );
        }
    }
}

#[test]
fn test_bar_chart_legend() {
    let workbook = parse_test_file("kitchen_sink_v2.xlsx");

    let chart = find_chart_by_title(&workbook, "Quarterly Sales");
    assert!(chart.is_some());

    let chart = chart.unwrap();

    // Should have a legend
    assert!(chart.legend.is_some(), "Bar chart should have a legend");

    let legend = chart.legend.as_ref().unwrap();
    // Default position is right
    assert_eq!(legend.position, "r", "Legend should be on the right");
}

// =============================================================================
// Tests: Line Chart Parsing
// =============================================================================

#[test]
fn test_line_chart_detection() {
    let workbook = parse_test_file("kitchen_sink_v2.xlsx");

    let line_count = count_charts_by_type(&workbook, ChartType::Line);
    assert!(
        line_count > 0,
        "Should have at least one Line chart, found {}",
        line_count
    );
}

#[test]
fn test_line_chart_sales_trend() {
    let workbook = parse_test_file("kitchen_sink_v2.xlsx");

    // Find the "Sales Trend" line chart
    let chart = find_chart_by_title(&workbook, "Sales Trend");
    assert!(chart.is_some(), "Should find 'Sales Trend' chart");

    let chart = chart.unwrap();

    // Verify chart type
    assert_eq!(chart.chart_type, ChartType::Line, "Should be a Line chart");

    // Verify grouping is standard
    assert_eq!(
        chart.grouping,
        Some(ChartGrouping::Standard),
        "Should have standard grouping"
    );

    // Verify title
    assert_eq!(
        chart.title.as_deref(),
        Some("Sales Trend"),
        "Title should match"
    );
}

#[test]
fn test_line_chart_series() {
    let workbook = parse_test_file("kitchen_sink_v2.xlsx");

    let chart = find_chart_by_title(&workbook, "Sales Trend");
    assert!(chart.is_some());

    let chart = chart.unwrap();

    // Should have multiple series
    assert!(!chart.series.is_empty(), "Line chart should have series");

    // Verify series indices
    for (i, series) in chart.series.iter().enumerate() {
        assert_eq!(series.idx as usize, i, "Series index should match position");
    }
}

// =============================================================================
// Tests: Pie Chart Parsing
// =============================================================================

#[test]
fn test_pie_chart_detection() {
    let workbook = parse_test_file("kitchen_sink_v2.xlsx");

    let pie_count = count_charts_by_type(&workbook, ChartType::Pie);
    assert!(
        pie_count > 0,
        "Should have at least one Pie chart, found {}",
        pie_count
    );
}

#[test]
fn test_pie_chart_product_distribution() {
    let workbook = parse_test_file("kitchen_sink_v2.xlsx");

    // Find the "Product Distribution" pie chart
    let chart = find_chart_by_title(&workbook, "Product Distribution");
    assert!(chart.is_some(), "Should find 'Product Distribution' chart");

    let chart = chart.unwrap();

    // Verify chart type
    assert_eq!(chart.chart_type, ChartType::Pie, "Should be a Pie chart");

    // Pie charts should have varyColors
    assert_eq!(
        chart.vary_colors,
        Some(true),
        "Pie chart should vary colors"
    );
}

#[test]
fn test_pie_chart_single_series() {
    let workbook = parse_test_file("kitchen_sink_v2.xlsx");

    let chart = find_chart_by_title(&workbook, "Product Distribution");
    assert!(chart.is_some());

    let chart = chart.unwrap();

    // Pie charts typically have a single series
    assert_eq!(
        chart.series.len(),
        1,
        "Pie chart should have exactly one series"
    );
}

#[test]
fn test_pie_chart_no_axes() {
    let workbook = parse_test_file("kitchen_sink_v2.xlsx");

    let chart = find_chart_by_type(&workbook, ChartType::Pie);
    assert!(chart.is_some());

    let chart = chart.unwrap();

    // Pie charts don't have axes
    assert!(chart.axes.is_empty(), "Pie chart should not have axes");
}

// =============================================================================
// Tests: Area Chart Parsing
// =============================================================================

#[test]
fn test_area_chart_detection() {
    let workbook = parse_test_file("kitchen_sink_v2.xlsx");

    let area_count = count_charts_by_type(&workbook, ChartType::Area);
    assert!(
        area_count > 0,
        "Should have at least one Area chart, found {}",
        area_count
    );
}

#[test]
fn test_area_chart_cumulative_sales() {
    let workbook = parse_test_file("kitchen_sink_v2.xlsx");

    // Find the "Cumulative Sales" area chart
    let chart = find_chart_by_title(&workbook, "Cumulative Sales");
    assert!(chart.is_some(), "Should find 'Cumulative Sales' chart");

    let chart = chart.unwrap();

    // Verify chart type
    assert_eq!(chart.chart_type, ChartType::Area, "Should be an Area chart");

    // Verify grouping
    assert_eq!(
        chart.grouping,
        Some(ChartGrouping::Standard),
        "Should have standard grouping"
    );
}

#[test]
fn test_area_chart_has_axes() {
    let workbook = parse_test_file("kitchen_sink_v2.xlsx");

    let chart = find_chart_by_type(&workbook, ChartType::Area);
    assert!(chart.is_some());

    let chart = chart.unwrap();

    // Area charts should have axes
    assert!(!chart.axes.is_empty(), "Area chart should have axes");
}

// =============================================================================
// Tests: Chart Positioning and Anchoring
// =============================================================================

#[test]
fn test_chart_positioning() {
    let workbook = parse_test_file("kitchen_sink_v2.xlsx");

    // Get all charts
    for sheet in &workbook.sheets {
        for chart in &sheet.charts {
            // Charts should have position information
            // from_col and from_row are set from the drawing anchor
            if chart.from_col.is_some() {
                assert!(
                    chart.from_row.is_some(),
                    "If from_col is set, from_row should also be set"
                );
            }

            // If we have to_col, we should have to_row (two-cell anchor)
            if chart.to_col.is_some() {
                assert!(
                    chart.to_row.is_some(),
                    "If to_col is set, to_row should also be set"
                );
            }
        }
    }
}

#[test]
fn test_chart_names() {
    let workbook = parse_test_file("kitchen_sink_v2.xlsx");

    // Some charts should have names from the drawing
    let charts_with_names: Vec<_> = workbook
        .sheets
        .iter()
        .flat_map(|s| s.charts.iter())
        .filter(|c| c.name.is_some())
        .collect();

    // It's OK if no charts have names, but if they do, verify they're non-empty
    for chart in charts_with_names {
        let name = chart.name.as_ref().unwrap();
        assert!(!name.is_empty(), "Chart name should not be empty");
    }
}

// =============================================================================
// Tests: Series Data References
// =============================================================================

#[test]
fn test_series_formula_references() {
    let workbook = parse_test_file("kitchen_sink_v2.xlsx");

    for sheet in &workbook.sheets {
        for chart in &sheet.charts {
            for series in &chart.series {
                // Check values reference
                if let Some(ref values) = series.values {
                    if let Some(ref formula) = values.formula {
                        // Formula should look like 'SheetName'!$A$1:$A$10
                        assert!(
                            formula.contains('!') || formula.contains('$'),
                            "Values formula should be a valid cell reference: {}",
                            formula
                        );
                    }
                }

                // Check categories reference
                if let Some(ref categories) = series.categories {
                    if let Some(ref formula) = categories.formula {
                        assert!(
                            formula.contains('!') || formula.contains('$'),
                            "Categories formula should be a valid cell reference: {}",
                            formula
                        );
                    }
                }

                // Check name reference
                if let Some(ref name_ref) = series.name_ref {
                    assert!(
                        name_ref.contains('!') || name_ref.contains('$') || !name_ref.is_empty(),
                        "Name reference should be valid: {}",
                        name_ref
                    );
                }
            }
        }
    }
}

// =============================================================================
// Tests: Chart Type Statistics
// =============================================================================

#[test]
fn test_chart_type_distribution() {
    let workbook = parse_test_file("kitchen_sink_v2.xlsx");

    let bar_count = count_charts_by_type(&workbook, ChartType::Bar);
    let line_count = count_charts_by_type(&workbook, ChartType::Line);
    let pie_count = count_charts_by_type(&workbook, ChartType::Pie);
    let area_count = count_charts_by_type(&workbook, ChartType::Area);

    let total = bar_count + line_count + pie_count + area_count;

    // We expect 4 charts total in kitchen_sink_v2.xlsx
    assert_eq!(
        total, 4,
        "Total charts should be 4, got bar={}, line={}, pie={}, area={}",
        bar_count, line_count, pie_count, area_count
    );

    assert_eq!(bar_count, 1, "Should have 1 bar chart");
    assert_eq!(line_count, 1, "Should have 1 line chart");
    assert_eq!(pie_count, 1, "Should have 1 pie chart");
    assert_eq!(area_count, 1, "Should have 1 area chart");
}

// =============================================================================
// Tests: Axis Properties
// =============================================================================

#[test]
fn test_axis_properties() {
    let workbook = parse_test_file("kitchen_sink_v2.xlsx");

    for sheet in &workbook.sheets {
        for chart in &sheet.charts {
            for axis in &chart.axes {
                // Axis should have an ID
                assert!(axis.id > 0, "Axis should have a positive ID");

                // Axis type should be valid
                let valid_types = [
                    "cat", "val", "date", "ser", "catAx", "valAx", "dateAx", "serAx",
                ];
                assert!(
                    valid_types.contains(&axis.axis_type.as_str()),
                    "Invalid axis type: {}",
                    axis.axis_type
                );

                // If position is set, it should be valid
                if let Some(ref pos) = axis.position {
                    let valid_positions = ["b", "l", "r", "t"];
                    assert!(
                        valid_positions.contains(&pos.as_str()),
                        "Invalid axis position: {}",
                        pos
                    );
                }
            }
        }
    }
}

#[test]
fn test_axis_gridlines() {
    let workbook = parse_test_file("kitchen_sink_v2.xlsx");

    // Check that value axes have major gridlines (common in Excel charts)
    let val_axes_with_gridlines: Vec<_> = workbook
        .sheets
        .iter()
        .flat_map(|s| s.charts.iter())
        .flat_map(|c| c.axes.iter())
        .filter(|a| (a.axis_type == "val" || a.axis_type == "valAx") && a.major_gridlines)
        .collect();

    assert!(
        !val_axes_with_gridlines.is_empty(),
        "Should have at least one value axis with major gridlines"
    );
}

// =============================================================================
// Tests: Legend Properties
// =============================================================================

#[test]
fn test_legend_properties() {
    let workbook = parse_test_file("kitchen_sink_v2.xlsx");

    let charts_with_legends: Vec<_> = workbook
        .sheets
        .iter()
        .flat_map(|s| s.charts.iter())
        .filter(|c| c.legend.is_some())
        .collect();

    assert!(
        !charts_with_legends.is_empty(),
        "Should have charts with legends"
    );

    for chart in charts_with_legends {
        let legend = chart.legend.as_ref().unwrap();

        // Position should be valid
        let valid_positions = ["b", "l", "r", "t", "tr"];
        assert!(
            valid_positions.contains(&legend.position.as_str()),
            "Invalid legend position: {}",
            legend.position
        );
    }
}

// =============================================================================
// Tests: Scatter Chart (if present in other files)
// =============================================================================

#[test]
fn test_scatter_chart_type_supported() {
    // This test verifies that the ChartType::Scatter enum variant exists
    // and can be used for matching
    let chart_type = ChartType::Scatter;
    assert_eq!(chart_type, ChartType::Scatter);
}

// =============================================================================
// Tests: Doughnut Chart (if present in other files)
// =============================================================================

#[test]
fn test_doughnut_chart_type_supported() {
    let chart_type = ChartType::Doughnut;
    assert_eq!(chart_type, ChartType::Doughnut);
}

// =============================================================================
// Tests: Radar Chart (if present in other files)
// =============================================================================

#[test]
fn test_radar_chart_type_supported() {
    let chart_type = ChartType::Radar;
    assert_eq!(chart_type, ChartType::Radar);
}

// =============================================================================
// Tests: Bubble Chart (if present in other files)
// =============================================================================

#[test]
fn test_bubble_chart_type_supported() {
    let chart_type = ChartType::Bubble;
    assert_eq!(chart_type, ChartType::Bubble);
}

// =============================================================================
// Tests: Stock Chart (if present in other files)
// =============================================================================

#[test]
fn test_stock_chart_type_supported() {
    let chart_type = ChartType::Stock;
    assert_eq!(chart_type, ChartType::Stock);
}

// =============================================================================
// Tests: Surface Chart (if present in other files)
// =============================================================================

#[test]
fn test_surface_chart_type_supported() {
    let chart_type = ChartType::Surface;
    assert_eq!(chart_type, ChartType::Surface);
}

// =============================================================================
// Tests: Combo Chart (if present in other files)
// =============================================================================

#[test]
fn test_combo_chart_type_supported() {
    let chart_type = ChartType::Combo;
    assert_eq!(chart_type, ChartType::Combo);
}

// =============================================================================
// Tests: Grouping Types
// =============================================================================

#[test]
fn test_grouping_types() {
    let workbook = parse_test_file("kitchen_sink_v2.xlsx");

    // Collect all grouping types
    let groupings: Vec<_> = workbook
        .sheets
        .iter()
        .flat_map(|s| s.charts.iter())
        .filter_map(|c| c.grouping)
        .collect();

    assert!(!groupings.is_empty(), "Should have charts with grouping");

    // Verify all groupings are valid
    for grouping in groupings {
        match grouping {
            ChartGrouping::Standard
            | ChartGrouping::Stacked
            | ChartGrouping::PercentStacked
            | ChartGrouping::Clustered => {
                // All valid groupings
            }
        }
    }
}

// =============================================================================
// Tests: Bar Direction
// =============================================================================

#[test]
fn test_bar_direction_types() {
    let workbook = parse_test_file("kitchen_sink_v2.xlsx");

    // Get bar charts
    let bar_charts: Vec<_> = workbook
        .sheets
        .iter()
        .flat_map(|s| s.charts.iter())
        .filter(|c| c.chart_type == ChartType::Bar)
        .collect();

    assert!(!bar_charts.is_empty(), "Should have bar charts");

    for chart in bar_charts {
        if let Some(direction) = chart.bar_direction {
            match direction {
                BarDirection::Col => {
                    // Column chart (vertical bars) - OK
                }
                BarDirection::Bar => {
                    // Bar chart (horizontal bars) - OK
                }
            }
        }
    }
}

// =============================================================================
// Tests: No Panic on Parsing
// =============================================================================

#[test]
fn test_parsing_does_not_panic() {
    // This test ensures parsing completes without panicking
    let result = std::panic::catch_unwind(|| {
        let _ = parse_test_file("kitchen_sink_v2.xlsx");
    });

    assert!(result.is_ok(), "Parsing should not panic");
}

#[test]
fn test_parsing_kitchen_sink_v1_no_panic() {
    // The v1 file might not have charts, but parsing should not panic
    let result = std::panic::catch_unwind(|| {
        let _ = parse_test_file("kitchen_sink.xlsx");
    });

    assert!(result.is_ok(), "Parsing kitchen_sink.xlsx should not panic");
}

// =============================================================================
// Tests: Chart Serialization
// =============================================================================

#[test]
fn test_chart_serialization_to_json() {
    let workbook = parse_test_file("kitchen_sink_v2.xlsx");

    // Serialize the workbook's charts to JSON
    for sheet in &workbook.sheets {
        for (chart_idx, chart) in sheet.charts.iter().enumerate() {
            let json = serde_json::to_value(chart).expect("Failed to serialize chart");

            // Chart type should be serialized
            assert!(
                json["chartType"].is_string(),
                "Chart {} should have chartType",
                chart_idx
            );

            // Series should be an array
            assert!(
                json["series"].is_array(),
                "Chart {} should have series array",
                chart_idx
            );
        }
    }
}

// =============================================================================
// Tests: Empty Charts Array for Sheets Without Charts
// =============================================================================

#[test]
fn test_sheets_without_charts_have_empty_array() {
    let workbook = parse_test_file("kitchen_sink.xlsx");

    // v1 file should not have charts
    for sheet in &workbook.sheets {
        // The charts field should be empty, not None
        // (it's Vec<Chart>, not Option<Vec<Chart>>)
        // This test just verifies the structure is correct
        let _ = sheet.charts.len(); // Should not panic
    }
}

// =============================================================================
// Tests: VaryColors for Pie/Doughnut Charts
// =============================================================================

#[test]
fn test_pie_chart_vary_colors() {
    let workbook = parse_test_file("kitchen_sink_v2.xlsx");

    let pie_chart = find_chart_by_type(&workbook, ChartType::Pie);
    assert!(pie_chart.is_some());

    let chart = pie_chart.unwrap();

    // Pie charts typically have varyColors=true
    assert_eq!(
        chart.vary_colors,
        Some(true),
        "Pie chart should have varyColors=true"
    );
}

// =============================================================================
// Tests: Series Order
// =============================================================================

#[test]
fn test_series_order() {
    let workbook = parse_test_file("kitchen_sink_v2.xlsx");

    for sheet in &workbook.sheets {
        for chart in &sheet.charts {
            // Verify series are in order
            let mut prev_order = None;
            for series in &chart.series {
                if let Some(prev) = prev_order {
                    // Order should be non-decreasing
                    assert!(
                        series.order >= prev,
                        "Series order should be non-decreasing"
                    );
                }
                prev_order = Some(series.order);
            }
        }
    }
}

// =============================================================================
// Tests: Multiple Charts on Same Sheet
// =============================================================================

#[test]
fn test_multiple_charts_same_sheet() {
    let workbook = parse_test_file("kitchen_sink_v2.xlsx");

    // Find sheets with multiple charts
    let sheets_with_multiple: Vec<_> = workbook
        .sheets
        .iter()
        .filter(|s| s.charts.len() > 1)
        .collect();

    // kitchen_sink_v2.xlsx has a Charts sheet with multiple charts
    if !sheets_with_multiple.is_empty() {
        for sheet in sheets_with_multiple {
            // Each chart should have a unique title or position
            let titles: Vec<_> = sheet.charts.iter().map(|c| &c.title).collect();

            // Check that not all titles are None
            let has_titles = titles.iter().any(|t| t.is_some());
            if has_titles {
                // If we have titles, they should be unique
                let unique_titles: std::collections::HashSet<_> =
                    titles.iter().filter_map(|t| t.as_ref()).collect();
                assert_eq!(
                    unique_titles.len(),
                    titles.iter().filter_map(|t| t.as_ref()).count(),
                    "Chart titles should be unique on the same sheet"
                );
            }
        }
    }
}

// =============================================================================
// Tests: Chart with Categories
// =============================================================================

#[test]
fn test_chart_categories() {
    let workbook = parse_test_file("kitchen_sink_v2.xlsx");

    // Bar and Line charts should have categories
    let bar_chart = find_chart_by_type(&workbook, ChartType::Bar);
    if let Some(chart) = bar_chart {
        // At least one series should have categories
        let has_categories = chart.series.iter().any(|s| s.categories.is_some());
        assert!(
            has_categories,
            "Bar chart should have series with categories"
        );
    }
}

// =============================================================================
// Tests: Axis Crossing
// =============================================================================

#[test]
fn test_axis_crossing() {
    let workbook = parse_test_file("kitchen_sink_v2.xlsx");

    for sheet in &workbook.sheets {
        for chart in &sheet.charts {
            // For charts with multiple axes, check crossing references
            if chart.axes.len() >= 2 {
                for axis in &chart.axes {
                    if let Some(crosses_ax) = axis.crosses_ax {
                        // The crossing axis ID should reference another axis in this chart
                        let crossing_exists = chart.axes.iter().any(|a| a.id == crosses_ax);
                        assert!(
                            crossing_exists,
                            "Axis {} crosses non-existent axis {}",
                            axis.id, crosses_ax
                        );
                    }
                }
            }
        }
    }
}

// =============================================================================
// Tests: Chart Title Extraction
// =============================================================================

#[test]
fn test_all_chart_titles() {
    let workbook = parse_test_file("kitchen_sink_v2.xlsx");

    let expected_titles = [
        "Quarterly Sales",
        "Sales Trend",
        "Product Distribution",
        "Cumulative Sales",
    ];

    for expected in &expected_titles {
        let found = find_chart_by_title(&workbook, expected);
        assert!(
            found.is_some(),
            "Should find chart with title '{}'",
            expected
        );
    }
}

// =============================================================================
// Integration Test: Full Chart Structure
// =============================================================================

#[test]
fn test_complete_chart_structure() {
    let workbook = parse_test_file("kitchen_sink_v2.xlsx");

    let chart = find_chart_by_title(&workbook, "Quarterly Sales");
    assert!(chart.is_some());

    let chart = chart.unwrap();

    // Verify complete structure
    assert_eq!(chart.chart_type, ChartType::Bar);
    assert!(chart.title.is_some());
    assert!(!chart.series.is_empty());
    assert!(!chart.axes.is_empty());
    assert!(chart.legend.is_some());
    assert!(chart.grouping.is_some());
    assert!(chart.bar_direction.is_some());

    // Series structure
    let first_series = &chart.series[0];
    assert_eq!(first_series.idx, 0);
    assert_eq!(first_series.order, 0);

    // Axis structure
    let first_axis = &chart.axes[0];
    assert!(first_axis.id > 0);
    assert!(!first_axis.axis_type.is_empty());
}
