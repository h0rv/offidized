use serde::{Deserialize, Serialize};

/// Type of chart
#[derive(Debug, Serialize, Deserialize, Clone, Copy, PartialEq, Eq, Default)]
#[serde(rename_all = "camelCase")]
pub enum ChartType {
    /// Bar chart (vertical bars)
    #[default]
    Bar,
    /// Line chart
    Line,
    /// Pie chart
    Pie,
    /// Area chart
    Area,
    /// Scatter chart (XY plot)
    Scatter,
    /// Doughnut chart
    Doughnut,
    /// Radar chart
    Radar,
    /// Bubble chart
    Bubble,
    /// Stock chart
    Stock,
    /// Surface chart
    Surface,
    /// Combo chart (multiple chart types)
    Combo,
}

/// Bar chart direction
#[derive(Debug, Serialize, Deserialize, Clone, Copy, PartialEq, Eq, Default)]
#[serde(rename_all = "camelCase")]
pub enum BarDirection {
    /// Vertical bars (column chart)
    #[default]
    Col,
    /// Horizontal bars
    Bar,
}

/// Chart grouping type
#[derive(Debug, Serialize, Deserialize, Clone, Copy, PartialEq, Eq, Default)]
#[serde(rename_all = "camelCase")]
pub enum ChartGrouping {
    /// Standard grouping
    #[default]
    Standard,
    /// Stacked grouping
    Stacked,
    /// Percent stacked
    PercentStacked,
    /// Clustered grouping
    Clustered,
}

/// A data series in a chart
#[derive(Debug, Serialize, Deserialize, Clone)]
#[serde(rename_all = "camelCase")]
pub struct ChartSeries {
    /// Series index (order in the chart)
    pub idx: u32,
    /// Series order (display order)
    pub order: u32,
    /// Series name/label
    #[serde(skip_serializing_if = "Option::is_none")]
    pub name: Option<String>,
    /// Cell reference for the series name (e.g., "Sheet1!$A$1")
    #[serde(skip_serializing_if = "Option::is_none")]
    pub name_ref: Option<String>,
    /// Category values (X-axis labels)
    #[serde(skip_serializing_if = "Option::is_none")]
    pub categories: Option<ChartDataRef>,
    /// Data values (Y-axis values)
    #[serde(skip_serializing_if = "Option::is_none")]
    pub values: Option<ChartDataRef>,
    /// X values (for scatter/bubble charts)
    #[serde(skip_serializing_if = "Option::is_none")]
    pub x_values: Option<ChartDataRef>,
    /// Bubble sizes (for bubble charts)
    #[serde(skip_serializing_if = "Option::is_none")]
    pub bubble_sizes: Option<ChartDataRef>,
    /// Series fill color
    #[serde(skip_serializing_if = "Option::is_none")]
    pub fill_color: Option<String>,
    /// Series line/border color
    #[serde(skip_serializing_if = "Option::is_none")]
    pub line_color: Option<String>,
    /// Series-specific chart type (for combo charts)
    #[serde(skip_serializing_if = "Option::is_none")]
    pub series_type: Option<ChartType>,
}

/// Reference to chart data (formula reference or cached values)
#[derive(Debug, Serialize, Deserialize, Clone, Default)]
#[serde(rename_all = "camelCase")]
pub struct ChartDataRef {
    /// Cell range formula (e.g., "Sheet1!$B$2:$B$5")
    #[serde(skip_serializing_if = "Option::is_none")]
    pub formula: Option<String>,
    /// Cached numeric values
    #[serde(skip_serializing_if = "Vec::is_empty")]
    pub num_values: Vec<Option<f64>>,
    /// Cached string values (for categories)
    #[serde(skip_serializing_if = "Vec::is_empty")]
    pub str_values: Vec<String>,
}

/// Chart axis information
#[derive(Debug, Serialize, Deserialize, Clone)]
#[serde(rename_all = "camelCase")]
pub struct ChartAxis {
    /// Axis ID
    pub id: u32,
    /// Axis type: "cat" (category), "val" (value), "date", "ser" (series)
    pub axis_type: String,
    /// Axis position: "b" (bottom), "l" (left), "r" (right), "t" (top)
    #[serde(skip_serializing_if = "Option::is_none")]
    pub position: Option<String>,
    /// Axis title
    #[serde(skip_serializing_if = "Option::is_none")]
    pub title: Option<String>,
    /// Minimum value (for value axis)
    #[serde(skip_serializing_if = "Option::is_none")]
    pub min: Option<f64>,
    /// Maximum value (for value axis)
    #[serde(skip_serializing_if = "Option::is_none")]
    pub max: Option<f64>,
    /// Major tick unit
    #[serde(skip_serializing_if = "Option::is_none")]
    pub major_unit: Option<f64>,
    /// Minor tick unit
    #[serde(skip_serializing_if = "Option::is_none")]
    pub minor_unit: Option<f64>,
    /// Whether to show major gridlines
    #[serde(default)]
    pub major_gridlines: bool,
    /// Whether to show minor gridlines
    #[serde(default)]
    pub minor_gridlines: bool,
    /// Crossing axis ID
    #[serde(skip_serializing_if = "Option::is_none")]
    pub crosses_ax: Option<u32>,
    /// Number format for axis labels
    #[serde(skip_serializing_if = "Option::is_none")]
    pub num_fmt: Option<String>,
    /// Whether axis is deleted/hidden
    #[serde(default)]
    pub deleted: bool,
}

/// Chart legend information
#[derive(Debug, Serialize, Deserialize, Clone)]
#[serde(rename_all = "camelCase")]
pub struct ChartLegend {
    /// Legend position: "b" (bottom), "l" (left), "r" (right), "t" (top), "tr" (top-right)
    pub position: String,
    /// Whether legend overlays the chart
    #[serde(default)]
    pub overlay: bool,
}

/// A complete chart object
#[derive(Debug, Serialize, Deserialize, Clone)]
#[serde(rename_all = "camelCase")]
pub struct Chart {
    /// Type of chart
    pub chart_type: ChartType,
    /// Bar direction (for bar/column charts)
    #[serde(skip_serializing_if = "Option::is_none")]
    pub bar_direction: Option<BarDirection>,
    /// Chart grouping (stacked, clustered, etc.)
    #[serde(skip_serializing_if = "Option::is_none")]
    pub grouping: Option<ChartGrouping>,
    /// Chart title
    #[serde(skip_serializing_if = "Option::is_none")]
    pub title: Option<String>,
    /// Data series in the chart
    pub series: Vec<ChartSeries>,
    /// Chart axes
    #[serde(skip_serializing_if = "Vec::is_empty")]
    pub axes: Vec<ChartAxis>,
    /// Legend information
    #[serde(skip_serializing_if = "Option::is_none")]
    pub legend: Option<ChartLegend>,
    /// Whether to vary colors by point (for pie/doughnut)
    #[serde(skip_serializing_if = "Option::is_none")]
    pub vary_colors: Option<bool>,
    /// Anchor position - starting column (0-indexed)
    #[serde(skip_serializing_if = "Option::is_none")]
    pub from_col: Option<u32>,
    /// Anchor position - starting row (0-indexed)
    #[serde(skip_serializing_if = "Option::is_none")]
    pub from_row: Option<u32>,
    /// Anchor position - ending column (0-indexed)
    #[serde(skip_serializing_if = "Option::is_none")]
    pub to_col: Option<u32>,
    /// Anchor position - ending row (0-indexed)
    #[serde(skip_serializing_if = "Option::is_none")]
    pub to_row: Option<u32>,
    /// Chart name from drawing
    #[serde(skip_serializing_if = "Option::is_none")]
    pub name: Option<String>,
    /// Extent width in EMUs (from one-cell anchor `<ext cx="..."/>`).
    #[serde(skip_serializing_if = "Option::is_none")]
    pub extent_cx: Option<i64>,
    /// Extent height in EMUs (from one-cell anchor `<ext cy="..."/>`).
    #[serde(skip_serializing_if = "Option::is_none")]
    pub extent_cy: Option<i64>,
}
