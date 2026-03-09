use serde::{Deserialize, Deserializer, Serialize, Serializer};
use std::ops::Deref;
use std::sync::Arc;

/// Resolved cell style
#[derive(Debug, Serialize, Deserialize, Default, Clone)]
#[serde(rename_all = "camelCase")]
pub struct Style {
    // Font
    #[serde(skip_serializing_if = "Option::is_none")]
    pub font_family: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub font_size: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub font_color: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub bold: Option<bool>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub italic: Option<bool>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub underline: Option<UnderlineStyle>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub strikethrough: Option<bool>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub vert_align: Option<VertAlign>,

    // Fill
    #[serde(skip_serializing_if = "Option::is_none")]
    pub bg_color: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub pattern_type: Option<PatternType>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub fg_color: Option<String>, // Pattern foreground color
    /// Gradient fill (if this cell uses a gradient instead of a solid/pattern fill)
    #[serde(skip_serializing_if = "Option::is_none")]
    pub gradient: Option<GradientFill>,

    // Borders
    #[serde(skip_serializing_if = "Option::is_none")]
    pub border_top: Option<Border>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub border_right: Option<Border>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub border_bottom: Option<Border>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub border_left: Option<Border>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub border_diagonal: Option<Border>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub diagonal_up: Option<bool>, // Line from bottom-left to top-right
    #[serde(skip_serializing_if = "Option::is_none")]
    pub diagonal_down: Option<bool>, // Line from top-left to bottom-right

    // Alignment
    #[serde(skip_serializing_if = "Option::is_none")]
    pub align_h: Option<HAlign>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub align_v: Option<VAlign>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub wrap: Option<bool>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub shrink_to_fit: Option<bool>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub indent: Option<u32>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub rotation: Option<i32>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub reading_order: Option<u8>, // 0=context, 1=LTR, 2=RTL

    // Protection
    #[serde(skip_serializing_if = "Option::is_none")]
    pub locked: Option<bool>, // Cell is locked (default true when sheet is protected)
    #[serde(skip_serializing_if = "Option::is_none")]
    pub hidden: Option<bool>, // Formula is hidden when sheet is protected
}

#[derive(Debug, Clone)]
pub struct StyleRef(pub Arc<Style>);

impl Deref for StyleRef {
    type Target = Style;

    fn deref(&self) -> &Self::Target {
        &self.0
    }
}

impl Serialize for StyleRef {
    fn serialize<S>(&self, serializer: S) -> Result<S::Ok, S::Error>
    where
        S: Serializer,
    {
        self.0.serialize(serializer)
    }
}

impl<'de> Deserialize<'de> for StyleRef {
    fn deserialize<D>(deserializer: D) -> Result<Self, D::Error>
    where
        D: Deserializer<'de>,
    {
        let style = Style::deserialize(deserializer)?;
        Ok(Self(Arc::new(style)))
    }
}

#[derive(Debug, Serialize, Deserialize, Clone)]
pub struct Border {
    pub style: BorderStyle,
    pub color: String,
}

#[derive(Debug, Serialize, Deserialize, Clone, Copy, Default, PartialEq, Eq)]
#[serde(rename_all = "lowercase")]
pub enum BorderStyle {
    #[default]
    None,
    Thin,
    Medium,
    Thick,
    Dashed,
    Dotted,
    Double,
    Hair,
    MediumDashed,
    DashDot,
    MediumDashDot,
    DashDotDot,
    MediumDashDotDot,
    SlantDashDot,
}

#[derive(Debug, Serialize, Deserialize, Clone, Copy, PartialEq)]
#[serde(rename_all = "camelCase")]
pub enum HAlign {
    General,
    Left,
    Center,
    Right,
    Fill,
    Justify,
    CenterContinuous,
    Distributed,
}

#[derive(Debug, Serialize, Deserialize, Clone, Copy, PartialEq)]
#[serde(rename_all = "camelCase")]
pub enum VAlign {
    Top,
    Center, // Note: Excel uses "center" not "middle"
    Bottom,
    Justify,
    Distributed,
}

/// Pane state for frozen/split panes
#[derive(Debug, Serialize, Deserialize, Clone, Copy, PartialEq)]
#[serde(rename_all = "camelCase")]
pub enum PaneState {
    Frozen,
    FrozenSplit,
    Split,
}

/// Underline style for font formatting
#[derive(Debug, Serialize, Deserialize, Clone, Copy)]
#[serde(rename_all = "camelCase")]
pub enum UnderlineStyle {
    Single,
    Double,
    SingleAccounting,
    DoubleAccounting,
    None,
}

/// Vertical alignment for text (subscript/superscript)
#[derive(Debug, Serialize, Deserialize, Clone, Copy)]
#[serde(rename_all = "camelCase")]
pub enum VertAlign {
    Baseline,
    Subscript,
    Superscript,
}

/// Pattern fill types from ECMA-376 Part 1, Section 18.18.55
#[derive(Debug, Serialize, Deserialize, Clone, Copy, PartialEq, Eq)]
#[serde(rename_all = "camelCase")]
pub enum PatternType {
    None,
    Solid,
    Gray125,
    Gray0625,
    DarkGray,
    MediumGray,
    LightGray,
    DarkHorizontal,
    DarkVertical,
    DarkDown,
    DarkUp,
    DarkGrid,
    DarkTrellis,
    LightHorizontal,
    LightVertical,
    LightDown,
    LightUp,
    LightGrid,
    LightTrellis,
}

#[derive(Debug, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct MergeRange {
    pub start_row: u32,
    pub start_col: u32,
    pub end_row: u32,
    pub end_col: u32,
}

#[derive(Debug, Serialize, Deserialize)]
pub struct ColWidth {
    pub col: u32,
    pub width: f64,
}

#[derive(Debug, Serialize, Deserialize)]
pub struct RowHeight {
    pub row: u32,
    pub height: f64,
}

/// Theme colors and fonts extracted from theme1.xml
#[derive(Debug, Serialize, Deserialize, Default)]
pub struct Theme {
    /// 12 theme colors: dk1, lt1, dk2, lt2, accent1-6, hlink, folHlink
    pub colors: Vec<String>,
    /// Major font (headings) from fontScheme
    #[serde(skip_serializing_if = "Option::is_none")]
    pub major_font: Option<String>,
    /// Minor font (body) from fontScheme
    #[serde(skip_serializing_if = "Option::is_none")]
    pub minor_font: Option<String>,
}

/// A gradient stop with position and color
#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct GradientStop {
    /// Position of the stop (0.0 to 1.0)
    pub position: f64,
    /// Color at this stop position
    pub color: String,
}

/// Gradient fill definition
#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct GradientFill {
    /// Gradient type: "linear" or "path"
    pub gradient_type: String,
    /// Angle in degrees for linear gradients (0 = left-to-right, 90 = top-to-bottom)
    #[serde(skip_serializing_if = "Option::is_none")]
    pub degree: Option<f64>,
    /// Left position for path gradients (0.0 to 1.0)
    #[serde(skip_serializing_if = "Option::is_none")]
    pub left: Option<f64>,
    /// Right position for path gradients (0.0 to 1.0)
    #[serde(skip_serializing_if = "Option::is_none")]
    pub right: Option<f64>,
    /// Top position for path gradients (0.0 to 1.0)
    #[serde(skip_serializing_if = "Option::is_none")]
    pub top: Option<f64>,
    /// Bottom position for path gradients (0.0 to 1.0)
    #[serde(skip_serializing_if = "Option::is_none")]
    pub bottom: Option<f64>,
    /// Color stops defining the gradient
    pub stops: Vec<GradientStop>,
}
