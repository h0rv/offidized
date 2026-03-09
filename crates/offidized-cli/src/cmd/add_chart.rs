use std::path::Path;

use anyhow::{bail, Result};
use offidized_xlsx::{
    BarDirection, Chart, ChartAxis, ChartDataRef, ChartGrouping, ChartLegend, ChartSeries,
    ChartType, Workbook,
};

/// Parsed series spec from `--series "Name:cats_range:vals_range"`.
struct SeriesSpec {
    name: String,
    categories: Option<String>,
    values: String,
}

fn parse_series(s: &str) -> Result<SeriesSpec> {
    let parts: Vec<&str> = s.splitn(3, ':').collect();
    match parts.as_slice() {
        [name, cats, vals] => Ok(SeriesSpec {
            name: name.to_string(),
            categories: if cats.is_empty() {
                None
            } else {
                Some(cats.to_string())
            },
            values: vals.to_string(),
        }),
        [name, vals] => Ok(SeriesSpec {
            name: name.to_string(),
            categories: None,
            values: vals.to_string(),
        }),
        _ => bail!(
            "invalid --series format; expected \"Name:categories_range:values_range\" or \"Name:values_range\""
        ),
    }
}

fn parse_chart_type(s: &str) -> Result<ChartType> {
    match s.to_lowercase().as_str() {
        "bar" | "column" => Ok(ChartType::Bar),
        "line" => Ok(ChartType::Line),
        "pie" => Ok(ChartType::Pie),
        "area" => Ok(ChartType::Area),
        "scatter" => Ok(ChartType::Scatter),
        "doughnut" => Ok(ChartType::Doughnut),
        "radar" => Ok(ChartType::Radar),
        _ => bail!(
            "unknown chart type {s:?}; supported: bar, column, line, pie, area, scatter, doughnut, radar"
        ),
    }
}

fn parse_anchor(s: &str) -> Result<(u32, u32, u32, u32)> {
    let parts: Vec<&str> = s.split(',').collect();
    if parts.len() != 4 {
        bail!("--anchor expects \"from_col,from_row,to_col,to_row\" (zero-based)");
    }
    let nums: Vec<u32> = parts
        .iter()
        .map(|p| p.trim().parse::<u32>())
        .collect::<std::result::Result<Vec<_>, _>>()
        .map_err(|_| anyhow::anyhow!("anchor values must be non-negative integers"))?;
    Ok((nums[0], nums[1], nums[2], nums[3]))
}

#[allow(clippy::too_many_arguments)]
pub fn run(
    path: &Path,
    sheet_name: &str,
    chart_type_str: &str,
    title: Option<&str>,
    series_specs: &[String],
    anchor: &str,
    bar_direction: Option<&str>,
    grouping: Option<&str>,
    legend_pos: Option<&str>,
    output: Option<&Path>,
    in_place: bool,
) -> Result<()> {
    let dest = match (output, in_place) {
        (Some(_), true) => bail!("cannot specify both -o and -i"),
        (Some(o), false) => o.to_path_buf(),
        (None, true) => path.to_path_buf(),
        (None, false) => bail!("must specify -o <path> or -i for in-place edit"),
    };

    let mut wb = Workbook::open(path)?;

    let chart_type = parse_chart_type(chart_type_str)?;
    let (from_col, from_row, to_col, to_row) = parse_anchor(anchor)?;

    let mut chart = Chart::new(chart_type);
    chart.set_anchor(from_col, from_row, to_col, to_row);

    if let Some(t) = title {
        chart.set_title(t);
    }

    if let Some(dir) = bar_direction {
        let d = match dir.to_lowercase().as_str() {
            "col" | "column" => BarDirection::Column,
            "bar" => BarDirection::Bar,
            _ => bail!("--bar-direction must be \"col\" or \"bar\""),
        };
        chart.set_bar_direction(d);
    }

    if let Some(grp) = grouping {
        let g = match grp.to_lowercase().as_str() {
            "clustered" => ChartGrouping::Clustered,
            "stacked" => ChartGrouping::Stacked,
            "percent-stacked" | "percentstacked" => ChartGrouping::PercentStacked,
            "standard" => ChartGrouping::Standard,
            _ => bail!("--grouping must be clustered, stacked, percent-stacked, or standard"),
        };
        chart.set_grouping(g);
    }

    if let Some(pos) = legend_pos {
        let mut legend = ChartLegend::new();
        legend.set_position(pos);
        chart.set_legend(legend);
    }

    for (i, spec_str) in series_specs.iter().enumerate() {
        let spec = parse_series(spec_str)?;
        let idx = i as u32;
        let mut series = ChartSeries::new(idx, idx);
        series.set_name(spec.name);

        if let Some(cats) = spec.categories {
            series.set_categories(ChartDataRef::from_formula(cats));
        }
        series.set_values(ChartDataRef::from_formula(spec.values));
        chart.add_series(series);
    }

    chart.add_axis(ChartAxis::new_category());
    chart.add_axis(ChartAxis::new_value());

    let ws = wb
        .sheet_mut(sheet_name)
        .ok_or_else(|| anyhow::anyhow!("sheet not found: {sheet_name}"))?;
    ws.add_chart(chart);

    wb.save(&dest)?;
    eprintln!("added {} chart to sheet {:?}", chart_type_str, sheet_name);
    Ok(())
}
