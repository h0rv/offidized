use std::path::Path;

use anyhow::Result;
use offidized_xlsx::Workbook;
use serde::Serialize;

#[derive(Serialize)]
struct ChartInfo {
    sheet: String,
    chart_type: String,
    title: Option<String>,
    series_count: usize,
    series: Vec<SeriesInfo>,
    has_legend: bool,
    bar_direction: Option<String>,
    grouping: Option<String>,
}

#[derive(Serialize)]
struct SeriesInfo {
    index: usize,
    category_ref: Option<String>,
    value_ref: Option<String>,
}

pub fn run(path: &Path, sheet_name: Option<&str>) -> Result<()> {
    let wb = Workbook::open(path)?;

    let mut chart_infos = Vec::new();

    for ws in wb.worksheets() {
        if let Some(filter_sheet) = sheet_name {
            if ws.name() != filter_sheet {
                continue;
            }
        }

        for chart in ws.charts() {
            let series: Vec<SeriesInfo> = chart
                .series()
                .iter()
                .enumerate()
                .map(|(i, s)| SeriesInfo {
                    index: i,
                    category_ref: s.categories().as_ref().map(|r| format!("{:?}", r)),
                    value_ref: s.values().as_ref().map(|r| format!("{:?}", r)),
                })
                .collect();

            chart_infos.push(ChartInfo {
                sheet: ws.name().to_string(),
                chart_type: format!("{:?}", chart.chart_type()),
                title: chart.title().map(String::from),
                series_count: chart.series().len(),
                series,
                has_legend: chart.legend().is_some(),
                bar_direction: chart.bar_direction().map(|d| format!("{:?}", d)),
                grouping: chart.grouping().map(|g| format!("{:?}", g)),
            });
        }
    }

    println!("{}", serde_json::to_string_pretty(&chart_infos)?);
    Ok(())
}
