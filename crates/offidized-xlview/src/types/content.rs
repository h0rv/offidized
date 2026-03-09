use serde::{Deserialize, Serialize};

/// Type of data validation
#[derive(Debug, Serialize, Deserialize, Clone, Copy, Default)]
#[serde(rename_all = "camelCase")]
pub enum ValidationType {
    #[default]
    None,
    Whole,   // Whole number
    Decimal, // Decimal number
    List,    // Dropdown list
    Date,
    Time,
    TextLength,
    Custom, // Custom formula
}

/// Operator for data validation comparisons
#[derive(Debug, Serialize, Deserialize, Clone, Copy, Default)]
#[serde(rename_all = "camelCase")]
pub enum ValidationOperator {
    #[default]
    Between,
    NotBetween,
    Equal,
    NotEqual,
    LessThan,
    LessThanOrEqual,
    GreaterThan,
    GreaterThanOrEqual,
}

/// Data validation rule for a cell
#[derive(Debug, Serialize, Deserialize, Clone)]
#[serde(rename_all = "camelCase")]
pub struct DataValidation {
    pub validation_type: ValidationType,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub operator: Option<ValidationOperator>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub formula1: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub formula2: Option<String>,
    pub allow_blank: bool,
    pub show_dropdown: bool, // For list type
    pub show_input_message: bool,
    pub show_error_message: bool,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub error_title: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub error_message: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub prompt_title: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub prompt_message: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub list_values: Option<Vec<String>>, // For list type with explicit values (parsed from formula1)
}

/// Data validation applied to a range of cells
#[derive(Debug, Serialize, Deserialize, Clone)]
#[serde(rename_all = "camelCase")]
pub struct DataValidationRange {
    pub sqref: String, // Cell range like "A1:A100"
    pub validation: DataValidation,
}

/// Outline level for a row or column (for grouping/collapsing)
#[derive(Debug, Serialize, Deserialize, Clone)]
#[serde(rename_all = "camelCase")]
pub struct OutlineLevel {
    /// Row or column index (0-indexed)
    pub index: u32,
    /// Outline level (1-7)
    pub level: u8,
    /// Is this group collapsed
    pub collapsed: bool,
    /// Is this row/column hidden due to collapse
    pub hidden: bool,
}
