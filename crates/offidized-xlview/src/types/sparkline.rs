use serde::{Deserialize, Serialize};

/// Type of sparkline chart
#[derive(Debug, Serialize, Deserialize, Clone, Copy, Default, PartialEq, Eq)]
#[serde(rename_all = "camelCase")]
pub enum SparklineType {
    /// Line chart (default)
    #[default]
    Line,
    /// Column/bar chart
    Column,
    /// Win/loss chart (stacked)
    Stacked,
}

/// How to handle empty cells in sparkline data
#[derive(Debug, Serialize, Deserialize, Clone, Copy, Default, PartialEq, Eq)]
#[serde(rename_all = "camelCase")]
pub enum SparklineEmptyCells {
    /// Treat empty cells as gaps (default)
    #[default]
    Gap,
    /// Treat empty cells as zero
    Zero,
    /// Connect data points across empty cells
    Connect,
}

/// Axis type for sparklines
#[derive(Debug, Serialize, Deserialize, Clone, Copy, Default, PartialEq, Eq)]
#[serde(rename_all = "camelCase")]
pub enum SparklineAxisType {
    /// Individual axis per sparkline (default)
    #[default]
    Individual,
    /// Shared axis across the group
    Group,
    /// Custom min/max values
    Custom,
}

/// Colors used in a sparkline
#[derive(Debug, Serialize, Deserialize, Clone, Default)]
#[serde(rename_all = "camelCase")]
pub struct SparklineColors {
    /// Series/line color (main color)
    #[serde(skip_serializing_if = "Option::is_none")]
    pub series: Option<String>,
    /// Color for negative values
    #[serde(skip_serializing_if = "Option::is_none")]
    pub negative: Option<String>,
    /// Axis color
    #[serde(skip_serializing_if = "Option::is_none")]
    pub axis: Option<String>,
    /// Markers color (for line sparklines)
    #[serde(skip_serializing_if = "Option::is_none")]
    pub markers: Option<String>,
    /// First point color
    #[serde(skip_serializing_if = "Option::is_none")]
    pub first: Option<String>,
    /// Last point color
    #[serde(skip_serializing_if = "Option::is_none")]
    pub last: Option<String>,
    /// Highest value color
    #[serde(skip_serializing_if = "Option::is_none")]
    pub high: Option<String>,
    /// Lowest value color
    #[serde(skip_serializing_if = "Option::is_none")]
    pub low: Option<String>,
}

/// A single sparkline within a group
#[derive(Debug, Serialize, Deserialize, Clone)]
#[serde(rename_all = "camelCase")]
pub struct Sparkline {
    /// Cell reference where the sparkline is displayed (e.g., "A1")
    pub location: String,
    /// Data range for the sparkline (e.g., "Sheet1!B1:F1")
    pub data_range: String,
}

/// A group of sparklines with shared settings
#[derive(Debug, Serialize, Deserialize, Clone)]
#[serde(rename_all = "camelCase")]
pub struct SparklineGroup {
    /// Type of sparkline (line, column, stacked)
    pub sparkline_type: String,
    /// Individual sparklines in this group
    pub sparklines: Vec<Sparkline>,
    /// Colors for the sparklines
    pub colors: SparklineColors,
    /// How to handle empty cells
    #[serde(skip_serializing_if = "Option::is_none")]
    pub display_empty_cells_as: Option<String>,
    /// Whether to show markers (for line sparklines)
    #[serde(default)]
    pub markers: bool,
    /// Whether to show the high point
    #[serde(default)]
    pub high_point: bool,
    /// Whether to show the low point
    #[serde(default)]
    pub low_point: bool,
    /// Whether to show the first point
    #[serde(default)]
    pub first_point: bool,
    /// Whether to show the last point
    #[serde(default)]
    pub last_point: bool,
    /// Whether to show negative values in different color
    #[serde(default)]
    pub negative_points: bool,
    /// Whether to show the axis
    #[serde(default)]
    pub display_x_axis: bool,
    /// Whether to display sparklines for hidden data
    #[serde(skip_serializing_if = "Option::is_none")]
    pub display_hidden: Option<bool>,
    /// Whether data is right-to-left
    #[serde(default)]
    pub right_to_left: bool,
    /// Line weight in points (for line sparklines)
    #[serde(skip_serializing_if = "Option::is_none")]
    pub line_weight: Option<f64>,
    /// Minimum axis type
    #[serde(skip_serializing_if = "Option::is_none")]
    pub min_axis_type: Option<String>,
    /// Maximum axis type
    #[serde(skip_serializing_if = "Option::is_none")]
    pub max_axis_type: Option<String>,
    /// Custom minimum value (when min_axis_type is Custom)
    #[serde(skip_serializing_if = "Option::is_none")]
    pub manual_min: Option<f64>,
    /// Custom maximum value (when max_axis_type is Custom)
    #[serde(skip_serializing_if = "Option::is_none")]
    pub manual_max: Option<f64>,
    /// Whether dates are used for the axis
    #[serde(skip_serializing_if = "Option::is_none")]
    pub date_axis: Option<bool>,
}
