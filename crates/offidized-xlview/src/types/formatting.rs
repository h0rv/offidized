use serde::{Deserialize, Serialize};

/// Known conditional formatting rule types.
///
/// Serializes to/from the same camelCase strings Excel uses (e.g. "colorScale", "cellIs").
/// Unknown types round-trip through `Other(String)`.
#[derive(Debug, Serialize, Deserialize, Clone, PartialEq, Eq, Hash)]
pub enum CFRuleType {
    #[serde(rename = "colorScale")]
    ColorScale,
    #[serde(rename = "dataBar")]
    DataBar,
    #[serde(rename = "iconSet")]
    IconSet,
    #[serde(rename = "cellIs")]
    CellIs,
    #[serde(rename = "expression")]
    Expression,
    #[serde(rename = "top10")]
    Top10,
    #[serde(rename = "aboveAverage")]
    AboveAverage,
    #[serde(rename = "timePeriod")]
    TimePeriod,
    #[serde(rename = "duplicateValues")]
    DuplicateValues,
    #[serde(rename = "uniqueValues")]
    UniqueValues,
    #[serde(rename = "containsBlanks")]
    ContainsBlanks,
    #[serde(rename = "notContainsBlanks")]
    NotContainsBlanks,
    #[serde(rename = "containsText")]
    ContainsText,
    #[serde(rename = "notContainsText")]
    NotContainsText,
    #[serde(rename = "beginsWith")]
    BeginsWith,
    #[serde(rename = "endsWith")]
    EndsWith,
    #[serde(rename = "containsErrors")]
    ContainsErrors,
    #[serde(rename = "notContainsErrors")]
    NotContainsErrors,
    /// Catch-all for future/unknown rule types.
    #[serde(untagged)]
    Other(String),
}

impl CFRuleType {
    /// Parse from a string (e.g. XML attribute value).
    pub fn from_str_val(s: &str) -> Self {
        match s {
            "colorScale" => Self::ColorScale,
            "dataBar" => Self::DataBar,
            "iconSet" => Self::IconSet,
            "cellIs" => Self::CellIs,
            "expression" => Self::Expression,
            "top10" => Self::Top10,
            "aboveAverage" => Self::AboveAverage,
            "timePeriod" => Self::TimePeriod,
            "duplicateValues" => Self::DuplicateValues,
            "uniqueValues" => Self::UniqueValues,
            "containsBlanks" => Self::ContainsBlanks,
            "notContainsBlanks" => Self::NotContainsBlanks,
            "containsText" => Self::ContainsText,
            "notContainsText" => Self::NotContainsText,
            "beginsWith" => Self::BeginsWith,
            "endsWith" => Self::EndsWith,
            "containsErrors" => Self::ContainsErrors,
            "notContainsErrors" => Self::NotContainsErrors,
            other => Self::Other(other.to_string()),
        }
    }
}

impl PartialEq<&str> for CFRuleType {
    fn eq(&self, other: &&str) -> bool {
        match (self, *other) {
            (Self::ColorScale, "colorScale")
            | (Self::DataBar, "dataBar")
            | (Self::IconSet, "iconSet")
            | (Self::CellIs, "cellIs")
            | (Self::Expression, "expression")
            | (Self::Top10, "top10")
            | (Self::AboveAverage, "aboveAverage")
            | (Self::TimePeriod, "timePeriod")
            | (Self::DuplicateValues, "duplicateValues")
            | (Self::UniqueValues, "uniqueValues")
            | (Self::ContainsBlanks, "containsBlanks")
            | (Self::NotContainsBlanks, "notContainsBlanks")
            | (Self::ContainsText, "containsText")
            | (Self::NotContainsText, "notContainsText")
            | (Self::BeginsWith, "beginsWith")
            | (Self::EndsWith, "endsWith")
            | (Self::ContainsErrors, "containsErrors")
            | (Self::NotContainsErrors, "notContainsErrors") => true,
            (Self::Other(s), o) => s == o,
            _ => false,
        }
    }
}

/// Conditional formatting rules applied to a range of cells
#[derive(Debug, Serialize, Deserialize, Clone)]
#[serde(rename_all = "camelCase")]
pub struct ConditionalFormatting {
    /// Cell range like "A1:A10" or multiple ranges "A1:A10 B1:B10"
    pub sqref: String,
    /// Rules to apply (evaluated in priority order)
    pub rules: Vec<CFRule>,
}

/// Cached conditional formatting metadata for rendering.
#[derive(Debug, Clone, Default)]
pub struct ConditionalFormattingCache {
    /// Parsed ranges for sqref entries.
    pub ranges: Vec<(u32, u32, u32, u32)>,
    /// Rule indices sorted by priority (ascending).
    pub sorted_rule_indices: Vec<usize>,
}

/// Differential formatting style (DXF) for conditional formatting
/// These are partial styles that only contain the properties that differ from the base style
#[derive(Debug, Serialize, Deserialize, Clone, Default)]
#[serde(rename_all = "camelCase")]
pub struct DxfStyle {
    /// Fill/background color
    #[serde(skip_serializing_if = "Option::is_none")]
    pub fill_color: Option<String>,
    /// Font color
    #[serde(skip_serializing_if = "Option::is_none")]
    pub font_color: Option<String>,
    /// Bold
    #[serde(skip_serializing_if = "Option::is_none")]
    pub bold: Option<bool>,
    /// Italic
    #[serde(skip_serializing_if = "Option::is_none")]
    pub italic: Option<bool>,
    /// Underline
    #[serde(skip_serializing_if = "Option::is_none")]
    pub underline: Option<bool>,
    /// Strikethrough
    #[serde(skip_serializing_if = "Option::is_none")]
    pub strikethrough: Option<bool>,
    /// Border color (all sides)
    #[serde(skip_serializing_if = "Option::is_none")]
    pub border_color: Option<String>,
    /// Border style (all sides)
    #[serde(skip_serializing_if = "Option::is_none")]
    pub border_style: Option<String>,
}

/// A single conditional formatting rule
#[derive(Debug, Serialize, Deserialize, Clone)]
#[serde(rename_all = "camelCase")]
pub struct CFRule {
    /// Rule type discriminant.
    pub rule_type: CFRuleType,
    /// Priority (lower = higher priority)
    pub priority: u32,
    /// Color scale definition (for type="colorScale")
    #[serde(skip_serializing_if = "Option::is_none")]
    pub color_scale: Option<ColorScale>,
    /// Data bar definition (for type="dataBar")
    #[serde(skip_serializing_if = "Option::is_none")]
    pub data_bar: Option<DataBar>,
    /// Icon set definition (for type="iconSet")
    #[serde(skip_serializing_if = "Option::is_none")]
    pub icon_set: Option<IconSet>,
    /// Formula for expression-based rules
    #[serde(skip_serializing_if = "Option::is_none")]
    pub formula: Option<String>,
    /// Operator for cellIs rules (equal, notEqual, lessThan, greaterThan, etc.)
    #[serde(skip_serializing_if = "Option::is_none")]
    pub operator: Option<String>,
    /// Reference to differential formatting in dxfs array
    #[serde(skip_serializing_if = "Option::is_none")]
    pub dxf_id: Option<u32>,

    // Top10 rule attributes (type="top10")
    /// Rank value for top10 rules (number of items to highlight)
    #[serde(skip_serializing_if = "Option::is_none")]
    pub rank: Option<u32>,
    /// Whether rank is a percentage (true) or count (false) for top10 rules
    #[serde(skip_serializing_if = "Option::is_none")]
    pub percent: Option<bool>,
    /// Whether to show bottom values instead of top for top10 rules
    #[serde(skip_serializing_if = "Option::is_none")]
    pub bottom: Option<bool>,

    // AboveAverage rule attributes (type="aboveAverage")
    /// Whether the rule is for above average (true) or below average (false)
    #[serde(skip_serializing_if = "Option::is_none")]
    pub above_average: Option<bool>,
    /// Whether to include values equal to the average
    #[serde(skip_serializing_if = "Option::is_none")]
    pub equal_average: Option<bool>,
    /// Standard deviation value for above/below average rules (1, 2, or 3)
    #[serde(skip_serializing_if = "Option::is_none")]
    pub std_dev: Option<u32>,

    // TimePeriod rule attributes (type="timePeriod")
    /// Time period value: today, yesterday, tomorrow, last7Days, thisWeek, lastWeek, nextWeek, thisMonth, lastMonth, nextMonth, thisYear, lastYear, nextYear
    #[serde(skip_serializing_if = "Option::is_none")]
    pub time_period: Option<String>,
}

/// Color scale definition (2 or 3 color gradient)
#[derive(Debug, Serialize, Deserialize, Clone)]
#[serde(rename_all = "camelCase")]
pub struct ColorScale {
    /// Value objects defining the scale points (2 or 3)
    pub cfvo: Vec<CFValueObject>,
    /// Colors for each scale point
    pub colors: Vec<String>,
}

/// Data bar definition
#[derive(Debug, Serialize, Deserialize, Clone)]
#[serde(rename_all = "camelCase")]
pub struct DataBar {
    /// Value objects for min and max
    pub cfvo: Vec<CFValueObject>,
    /// Bar color
    pub color: String,
    /// Show value in cell (default true)
    #[serde(skip_serializing_if = "Option::is_none")]
    pub show_value: Option<bool>,
    /// Minimum bar length as percentage (0-100)
    #[serde(skip_serializing_if = "Option::is_none")]
    pub min_length: Option<u32>,
    /// Maximum bar length as percentage (0-100)
    #[serde(skip_serializing_if = "Option::is_none")]
    pub max_length: Option<u32>,
}

/// Icon set definition
#[derive(Debug, Serialize, Deserialize, Clone)]
#[serde(rename_all = "camelCase")]
pub struct IconSet {
    /// Icon set name (3Arrows, 3TrafficLights, 4Rating, 5Quarters, etc.)
    pub icon_set: String,
    /// Value objects for each threshold
    pub cfvo: Vec<CFValueObject>,
    /// Show icon only (hide value)
    #[serde(skip_serializing_if = "Option::is_none")]
    pub show_value: Option<bool>,
    /// Reverse icon order
    #[serde(skip_serializing_if = "Option::is_none")]
    pub reverse: Option<bool>,
}

/// Conditional formatting value object (cfvo)
#[derive(Debug, Serialize, Deserialize, Clone)]
#[serde(rename_all = "camelCase")]
pub struct CFValueObject {
    /// Type: min, max, num, percent, percentile, formula
    pub cfvo_type: String,
    /// Value (for num, percent, percentile, formula types)
    #[serde(skip_serializing_if = "Option::is_none")]
    pub val: Option<String>,
}
