//! Render-specific types for the Excel viewer.

pub mod chart;
pub mod content;
pub mod drawing;
pub mod filter;
pub mod formatting;
pub mod page;
pub mod rich_text;
pub mod selection;
pub mod sparkline;
pub mod style;
pub mod workbook;

// Re-export commonly used types at the module level
pub use chart::ChartLegend;
pub use formatting::{
    CFRule, CFRuleType, CFValueObject, ColorScale, ConditionalFormatting,
    ConditionalFormattingCache, DataBar, DxfStyle, IconSet,
};
pub use rich_text::TextRunData;
pub use selection::SelectionType;
pub use workbook::{
    Cell, CellData, CellRawValue, CellType, Comment, CompiledFormat, HyperlinkDef, Sheet, Workbook,
};

/// Hyperlink data for a cell or drawing.
#[derive(Debug, Clone, serde::Serialize, serde::Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct Hyperlink {
    /// Target URL or internal reference
    pub target: String,
    /// Whether the link is external (URL) vs internal (sheet reference)
    pub is_external: bool,
    /// Optional display text
    #[serde(skip_serializing_if = "Option::is_none")]
    pub display: Option<String>,
    /// Optional tooltip
    #[serde(skip_serializing_if = "Option::is_none")]
    pub tooltip: Option<String>,
    /// Bookmark/location within target (e.g. cell reference for internal links).
    #[serde(skip_serializing_if = "Option::is_none")]
    pub location: Option<String>,
}
