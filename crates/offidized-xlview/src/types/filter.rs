use serde::{Deserialize, Serialize};

/// Auto-filter definition for a range of cells
#[derive(Debug, Serialize, Deserialize, Clone)]
#[serde(rename_all = "camelCase")]
pub struct AutoFilter {
    /// The range covered by the filter (e.g., "A1:D100")
    pub range: String,
    /// Start row of the filter range (0-indexed)
    pub start_row: u32,
    /// Start column of the filter range (0-indexed)
    pub start_col: u32,
    /// End row of the filter range (0-indexed)
    pub end_row: u32,
    /// End column of the filter range (0-indexed)
    pub end_col: u32,
    /// Columns that have filter criteria applied
    pub filter_columns: Vec<FilterColumn>,
}

/// A column with filter settings within an auto-filter range
#[derive(Debug, Serialize, Deserialize, Clone)]
#[serde(rename_all = "camelCase")]
pub struct FilterColumn {
    /// 0-based column index within the auto-filter range
    pub col_id: u32,
    /// True if a filter is actively applied (not showing all values)
    pub has_filter: bool,
    /// The type of filter applied to this column
    pub filter_type: FilterType,
    /// Whether to show the filter button (default true)
    #[serde(skip_serializing_if = "Option::is_none")]
    pub show_button: Option<bool>,
    /// Whether blank cells are included in the filter
    #[serde(skip_serializing_if = "Option::is_none")]
    pub blank: Option<bool>,
    /// Values for Values filter type
    #[serde(skip_serializing_if = "Vec::is_empty", default)]
    pub values: Vec<String>,
    /// Custom filters (for Custom filter type)
    #[serde(skip_serializing_if = "Vec::is_empty", default)]
    pub custom_filters: Vec<CustomFilter>,
    /// Whether custom filters use AND logic (true) or OR logic (false)
    #[serde(skip_serializing_if = "Option::is_none")]
    pub custom_filters_and: Option<bool>,
    /// DxfId for color filter
    #[serde(skip_serializing_if = "Option::is_none")]
    pub dxf_id: Option<u32>,
    /// Whether color filter applies to cell color (true) or font color (false)
    #[serde(skip_serializing_if = "Option::is_none")]
    pub cell_color: Option<bool>,
    /// Icon set index for icon filter
    #[serde(skip_serializing_if = "Option::is_none")]
    pub icon_set: Option<u32>,
    /// Icon ID within the icon set
    #[serde(skip_serializing_if = "Option::is_none")]
    pub icon_id: Option<u32>,
    /// Dynamic filter type name
    #[serde(skip_serializing_if = "Option::is_none")]
    pub dynamic_type: Option<String>,
    /// Top10 filter: whether to filter top (true) or bottom (false)
    #[serde(skip_serializing_if = "Option::is_none")]
    pub top: Option<bool>,
    /// Top10 filter: whether to use percent (true) or count (false)
    #[serde(skip_serializing_if = "Option::is_none")]
    pub percent: Option<bool>,
    /// Top10 filter: the value (count or percent)
    #[serde(skip_serializing_if = "Option::is_none")]
    pub top10_val: Option<f64>,
}

/// Custom filter condition
#[derive(Debug, Serialize, Deserialize, Clone)]
#[serde(rename_all = "camelCase")]
pub struct CustomFilter {
    /// Operator for the filter
    pub operator: CustomFilterOperator,
    /// Value to compare against
    pub val: String,
}

/// Operators for custom filters
#[derive(Debug, Serialize, Deserialize, Clone, Copy, PartialEq, Eq, Default)]
#[serde(rename_all = "camelCase")]
pub enum CustomFilterOperator {
    #[default]
    Equal,
    NotEqual,
    GreaterThan,
    GreaterThanOrEqual,
    LessThan,
    LessThanOrEqual,
}

/// Types of filters that can be applied to a column
#[derive(Debug, Serialize, Deserialize, Clone, Copy, PartialEq, Eq)]
#[serde(rename_all = "camelCase")]
pub enum FilterType {
    /// No filter (show all values)
    None,
    /// Specific values selected
    Values,
    /// Custom filter criteria (operators like greaterThan, lessThan, etc.)
    Custom,
    /// Top/bottom N filter
    Top10,
    /// Dynamic filter (dates, above average, etc.)
    Dynamic,
    /// Filter by cell color
    Color,
    /// Filter by icon
    Icon,
}
