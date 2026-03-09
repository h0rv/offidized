use crate::error::{Result, XlsxError};

// ===== Font vertical alignment =====

/// Vertical alignment for font characters (superscript, subscript, baseline).
///
/// Maps to the OOXML `<vertAlign val="..."/>` element inside a font definition.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum FontVerticalAlign {
    /// Characters appear as superscript (raised, smaller).
    Superscript,
    /// Characters appear as subscript (lowered, smaller).
    Subscript,
    /// Characters appear at the normal baseline position.
    Baseline,
}

impl FontVerticalAlign {
    pub(crate) fn as_xml_value(self) -> &'static str {
        match self {
            Self::Superscript => "superscript",
            Self::Subscript => "subscript",
            Self::Baseline => "baseline",
        }
    }

    pub(crate) fn from_xml_value(value: &str) -> Option<Self> {
        match value {
            "superscript" => Some(Self::Superscript),
            "subscript" => Some(Self::Subscript),
            "baseline" => Some(Self::Baseline),
            _ => None,
        }
    }
}

// ===== Font scheme =====

/// Font scheme classification.
///
/// Maps to the OOXML `<scheme val="..."/>` element inside a font definition.
/// This tells Excel whether the font is the major (heading) or minor (body) font
/// from the theme, or not associated with either.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum FontScheme {
    /// The font is the major (heading) theme font.
    Major,
    /// The font is the minor (body) theme font.
    Minor,
    /// The font is not associated with a theme font scheme.
    None,
}

impl FontScheme {
    pub(crate) fn as_xml_value(self) -> &'static str {
        match self {
            Self::Major => "major",
            Self::Minor => "minor",
            Self::None => "none",
        }
    }

    #[allow(unused_qualifications)]
    pub(crate) fn from_xml_value(value: &str) -> Option<Self> {
        match value {
            "major" => Some(Self::Major),
            "minor" => Some(Self::Minor),
            "none" => Some(Self::None),
            _ => Option::None,
        }
    }
}

// ===== Feature 9: Theme colors =====

/// Theme color indices corresponding to the OOXML theme color slots.
///
/// These map to the `theme="X"` attribute in OOXML color references.
/// In a standard theme, these are:
/// - dk1 (0) = Dark 1 (usually black)
/// - lt1 (1) = Light 1 (usually white)
/// - dk2 (2) = Dark 2
/// - lt2 (3) = Light 2
/// - accent1-6 (4-9) = Accent colors
/// - hlink (10) = Hyperlink
/// - folHlink (11) = Followed Hyperlink
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ThemeColor {
    Dark1,
    Light1,
    Dark2,
    Light2,
    Accent1,
    Accent2,
    Accent3,
    Accent4,
    Accent5,
    Accent6,
    Hyperlink,
    FollowedHyperlink,
}

impl ThemeColor {
    /// Returns the zero-based theme index used in XML attributes.
    pub fn index(self) -> u32 {
        match self {
            Self::Dark1 => 0,
            Self::Light1 => 1,
            Self::Dark2 => 2,
            Self::Light2 => 3,
            Self::Accent1 => 4,
            Self::Accent2 => 5,
            Self::Accent3 => 6,
            Self::Accent4 => 7,
            Self::Accent5 => 8,
            Self::Accent6 => 9,
            Self::Hyperlink => 10,
            Self::FollowedHyperlink => 11,
        }
    }

    /// Creates a `ThemeColor` from a zero-based theme index.
    pub fn from_index(index: u32) -> Option<Self> {
        match index {
            0 => Some(Self::Dark1),
            1 => Some(Self::Light1),
            2 => Some(Self::Dark2),
            3 => Some(Self::Light2),
            4 => Some(Self::Accent1),
            5 => Some(Self::Accent2),
            6 => Some(Self::Accent3),
            7 => Some(Self::Accent4),
            8 => Some(Self::Accent5),
            9 => Some(Self::Accent6),
            10 => Some(Self::Hyperlink),
            11 => Some(Self::FollowedHyperlink),
            _ => None,
        }
    }

    /// Returns the OOXML name for this theme color slot.
    pub fn name(self) -> &'static str {
        match self {
            Self::Dark1 => "dk1",
            Self::Light1 => "lt1",
            Self::Dark2 => "dk2",
            Self::Light2 => "lt2",
            Self::Accent1 => "accent1",
            Self::Accent2 => "accent2",
            Self::Accent3 => "accent3",
            Self::Accent4 => "accent4",
            Self::Accent5 => "accent5",
            Self::Accent6 => "accent6",
            Self::Hyperlink => "hlink",
            Self::FollowedHyperlink => "folHlink",
        }
    }
}

/// A color reference that can be an RGB value, a theme color, or a theme color with tint.
///
/// OOXML color references support `rgb`, `theme`, `tint`, `indexed`, and `auto` attributes.
/// `tint` is a value from -1.0 to 1.0 that lightens (positive) or darkens (negative)
/// the base theme color. `indexed` refers to the legacy indexed color palette, and
/// `auto` indicates an automatic (system-default) color.
#[derive(Debug, Clone, PartialEq)]
pub struct ColorReference {
    rgb: Option<String>,
    theme: Option<ThemeColor>,
    tint: Option<f64>,
    indexed: Option<u32>,
    auto: Option<bool>,
}

impl ColorReference {
    /// Creates an empty color reference with no attributes set.
    pub fn empty() -> Self {
        Self {
            rgb: None,
            theme: None,
            tint: None,
            indexed: None,
            auto: None,
        }
    }

    /// Creates a color reference from an ARGB hex string (e.g. "FF0000FF").
    pub fn from_rgb(rgb: impl Into<String>) -> Self {
        Self {
            rgb: Some(rgb.into()),
            theme: None,
            tint: None,
            indexed: None,
            auto: None,
        }
    }

    /// Creates a color reference from a theme color.
    pub fn from_theme(theme: ThemeColor) -> Self {
        Self {
            rgb: None,
            theme: Some(theme),
            tint: None,
            indexed: None,
            auto: None,
        }
    }

    /// Creates a color reference from a theme color with a tint value.
    pub fn from_theme_with_tint(theme: ThemeColor, tint: f64) -> Self {
        Self {
            rgb: None,
            theme: Some(theme),
            tint: Some(tint),
            indexed: None,
            auto: None,
        }
    }

    /// Creates a color reference from an indexed color palette value.
    pub fn from_indexed(indexed: u32) -> Self {
        Self {
            rgb: None,
            theme: None,
            tint: None,
            indexed: Some(indexed),
            auto: None,
        }
    }

    /// Returns the RGB value if set.
    pub fn rgb(&self) -> Option<&str> {
        self.rgb.as_deref()
    }

    /// Returns the theme color if set.
    pub fn theme(&self) -> Option<ThemeColor> {
        self.theme
    }

    /// Returns the tint value if set.
    pub fn tint(&self) -> Option<f64> {
        self.tint
    }

    /// Sets the RGB value.
    pub fn set_rgb(&mut self, rgb: impl Into<String>) -> &mut Self {
        self.rgb = Some(rgb.into());
        self
    }

    /// Sets the theme color.
    pub fn set_theme(&mut self, theme: ThemeColor) -> &mut Self {
        self.theme = Some(theme);
        self
    }

    /// Sets the tint value (-1.0 to 1.0).
    pub fn set_tint(&mut self, tint: f64) -> &mut Self {
        self.tint = Some(tint);
        self
    }

    /// Clears the tint value.
    pub fn clear_tint(&mut self) -> &mut Self {
        self.tint = None;
        self
    }

    /// Returns the indexed color palette value if set.
    pub fn indexed(&self) -> Option<u32> {
        self.indexed
    }

    /// Sets the indexed color palette value.
    pub fn set_indexed(&mut self, indexed: u32) -> &mut Self {
        self.indexed = Some(indexed);
        self
    }

    /// Clears the indexed color palette value.
    pub fn clear_indexed(&mut self) -> &mut Self {
        self.indexed = None;
        self
    }

    /// Returns the auto color flag if set.
    pub fn auto(&self) -> Option<bool> {
        self.auto
    }

    /// Sets the auto color flag.
    pub fn set_auto(&mut self, auto: bool) -> &mut Self {
        self.auto = Some(auto);
        self
    }

    /// Clears the auto color flag.
    pub fn clear_auto(&mut self) -> &mut Self {
        self.auto = None;
        self
    }

    /// Resolve this color reference to an `#RRGGBB` hex string.
    ///
    /// This is a convenience wrapper around [`crate::color::resolve_color`].
    ///
    /// # Arguments
    /// - `theme_colors` - The workbook's theme color palette (up to 12 entries).
    /// - `indexed_colors` - Optional custom indexed color palette.
    pub fn resolve(
        &self,
        theme_colors: &[String],
        indexed_colors: Option<&[String]>,
    ) -> Option<String> {
        crate::color::resolve_color(self, theme_colors, indexed_colors)
    }
}

// ===== Feature 7: Gradient fill =====

/// The type of gradient fill.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum GradientFillType {
    /// A linear gradient.
    Linear,
    /// A path (rectangular) gradient.
    Path,
}

#[allow(dead_code)]
impl GradientFillType {
    pub(crate) fn as_xml_value(self) -> &'static str {
        match self {
            Self::Linear => "linear",
            Self::Path => "path",
        }
    }

    pub(crate) fn from_xml_value(value: &str) -> Option<Self> {
        match value {
            "linear" => Some(Self::Linear),
            "path" => Some(Self::Path),
            _ => None,
        }
    }
}

/// A single color stop in a gradient fill.
#[derive(Debug, Clone, PartialEq)]
pub struct GradientStop {
    position: f64,
    color: String,
    color_ref: Option<ColorReference>,
}

impl GradientStop {
    /// Creates a new gradient stop at the given position (0.0 to 1.0) with an ARGB color.
    pub fn new(position: f64, color: impl Into<String>) -> Self {
        Self {
            position: position.clamp(0.0, 1.0),
            color: color.into(),
            color_ref: None,
        }
    }

    /// Creates a new gradient stop with a `ColorReference` for theme/indexed color resolution.
    pub fn with_color_ref(position: f64, color_ref: ColorReference) -> Self {
        let color = color_ref.rgb().unwrap_or_default().to_string();
        Self {
            position: position.clamp(0.0, 1.0),
            color,
            color_ref: Some(color_ref),
        }
    }

    /// Returns the position of this stop (0.0 to 1.0).
    pub fn position(&self) -> f64 {
        self.position
    }

    /// Returns the color (ARGB hex string).
    pub fn color(&self) -> &str {
        self.color.as_str()
    }

    /// Returns the color reference, if available (for theme/indexed color resolution).
    pub fn color_ref(&self) -> Option<&ColorReference> {
        self.color_ref.as_ref()
    }
}

/// A gradient fill definition.
///
/// Gradient fills support linear and path (rectangular) types, a rotation degree,
/// path bounds (left/right/top/bottom), and multiple color stops.
#[derive(Debug, Clone, PartialEq)]
pub struct GradientFill {
    gradient_type: GradientFillType,
    degree: Option<f64>,
    left: Option<f64>,
    right: Option<f64>,
    top: Option<f64>,
    bottom: Option<f64>,
    stops: Vec<GradientStop>,
}

impl GradientFill {
    /// Creates a new linear gradient fill with the given degree (0-360).
    pub fn linear(degree: f64) -> Self {
        Self {
            gradient_type: GradientFillType::Linear,
            degree: Some(degree),
            left: None,
            right: None,
            top: None,
            bottom: None,
            stops: Vec::new(),
        }
    }

    /// Creates a new path (rectangular) gradient fill.
    pub fn path() -> Self {
        Self {
            gradient_type: GradientFillType::Path,
            degree: None,
            left: None,
            right: None,
            top: None,
            bottom: None,
            stops: Vec::new(),
        }
    }

    /// Returns the gradient type.
    pub fn gradient_type(&self) -> GradientFillType {
        self.gradient_type
    }

    /// Sets the gradient type.
    pub fn set_gradient_type(&mut self, gradient_type: GradientFillType) -> &mut Self {
        self.gradient_type = gradient_type;
        self
    }

    /// Returns the degree of rotation for linear gradients.
    pub fn degree(&self) -> Option<f64> {
        self.degree
    }

    /// Sets the degree of rotation.
    pub fn set_degree(&mut self, degree: f64) -> &mut Self {
        self.degree = Some(degree);
        self
    }

    /// Clears the degree.
    pub fn clear_degree(&mut self) -> &mut Self {
        self.degree = None;
        self
    }

    /// Returns the left position for path gradients (0.0 to 1.0).
    pub fn left(&self) -> Option<f64> {
        self.left
    }

    /// Sets the left position for path gradients.
    pub fn set_left(&mut self, left: f64) -> &mut Self {
        self.left = Some(left);
        self
    }

    /// Returns the right position for path gradients (0.0 to 1.0).
    pub fn right(&self) -> Option<f64> {
        self.right
    }

    /// Sets the right position for path gradients.
    pub fn set_right(&mut self, right: f64) -> &mut Self {
        self.right = Some(right);
        self
    }

    /// Returns the top position for path gradients (0.0 to 1.0).
    pub fn top(&self) -> Option<f64> {
        self.top
    }

    /// Sets the top position for path gradients.
    pub fn set_top(&mut self, top: f64) -> &mut Self {
        self.top = Some(top);
        self
    }

    /// Returns the bottom position for path gradients (0.0 to 1.0).
    pub fn bottom(&self) -> Option<f64> {
        self.bottom
    }

    /// Sets the bottom position for path gradients.
    pub fn set_bottom(&mut self, bottom: f64) -> &mut Self {
        self.bottom = Some(bottom);
        self
    }

    /// Returns the color stops.
    pub fn stops(&self) -> &[GradientStop] {
        self.stops.as_slice()
    }

    /// Adds a color stop at the given position.
    pub fn add_stop(&mut self, position: f64, color: impl Into<String>) -> &mut Self {
        self.stops.push(GradientStop::new(position, color));
        self
    }

    /// Adds a color stop with a `ColorReference` for theme/indexed color resolution.
    pub fn add_stop_with_color_ref(
        &mut self,
        position: f64,
        color_ref: ColorReference,
    ) -> &mut Self {
        self.stops
            .push(GradientStop::with_color_ref(position, color_ref));
        self
    }

    /// Clears all color stops.
    pub fn clear_stops(&mut self) -> &mut Self {
        self.stops.clear();
        self
    }
}

/// Horizontal cell alignment values.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum HorizontalAlignment {
    General,
    Left,
    Center,
    Right,
    Fill,
    Justify,
    CenterContinuous,
    Distributed,
}

impl HorizontalAlignment {
    pub(crate) fn as_xml_value(self) -> &'static str {
        match self {
            Self::General => "general",
            Self::Left => "left",
            Self::Center => "center",
            Self::Right => "right",
            Self::Fill => "fill",
            Self::Justify => "justify",
            Self::CenterContinuous => "centerContinuous",
            Self::Distributed => "distributed",
        }
    }

    pub(crate) fn from_xml_value(value: &str) -> Option<Self> {
        match value {
            "general" => Some(Self::General),
            "left" => Some(Self::Left),
            "center" => Some(Self::Center),
            "right" => Some(Self::Right),
            "fill" => Some(Self::Fill),
            "justify" => Some(Self::Justify),
            "centerContinuous" => Some(Self::CenterContinuous),
            "distributed" => Some(Self::Distributed),
            _ => None,
        }
    }
}

/// Vertical cell alignment values.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum VerticalAlignment {
    Top,
    Center,
    Bottom,
    Justify,
    Distributed,
}

impl VerticalAlignment {
    pub(crate) fn as_xml_value(self) -> &'static str {
        match self {
            Self::Top => "top",
            Self::Center => "center",
            Self::Bottom => "bottom",
            Self::Justify => "justify",
            Self::Distributed => "distributed",
        }
    }

    pub(crate) fn from_xml_value(value: &str) -> Option<Self> {
        match value {
            "top" => Some(Self::Top),
            "center" => Some(Self::Center),
            "bottom" => Some(Self::Bottom),
            "justify" => Some(Self::Justify),
            "distributed" => Some(Self::Distributed),
            _ => None,
        }
    }
}

/// Cell alignment metadata.
#[derive(Debug, Clone, PartialEq, Eq, Default)]
pub struct Alignment {
    horizontal: Option<HorizontalAlignment>,
    vertical: Option<VerticalAlignment>,
    wrap_text: Option<bool>,
    indent: Option<u32>,
    text_rotation: Option<u32>,
    shrink_to_fit: Option<bool>,
    reading_order: Option<u32>,
}

impl Alignment {
    pub fn new() -> Self {
        Self::default()
    }

    pub fn horizontal(&self) -> Option<HorizontalAlignment> {
        self.horizontal
    }

    pub fn set_horizontal(&mut self, horizontal: HorizontalAlignment) -> &mut Self {
        self.horizontal = Some(horizontal);
        self
    }

    pub fn clear_horizontal(&mut self) -> &mut Self {
        self.horizontal = None;
        self
    }

    pub fn vertical(&self) -> Option<VerticalAlignment> {
        self.vertical
    }

    pub fn set_vertical(&mut self, vertical: VerticalAlignment) -> &mut Self {
        self.vertical = Some(vertical);
        self
    }

    pub fn clear_vertical(&mut self) -> &mut Self {
        self.vertical = None;
        self
    }

    pub fn wrap_text(&self) -> Option<bool> {
        self.wrap_text
    }

    pub fn set_wrap_text(&mut self, wrap_text: bool) -> &mut Self {
        self.wrap_text = Some(wrap_text);
        self
    }

    pub fn clear_wrap_text(&mut self) -> &mut Self {
        self.wrap_text = None;
        self
    }

    /// Returns the cell indent level.
    pub fn indent(&self) -> Option<u32> {
        self.indent
    }

    /// Sets the cell indent level.
    pub fn set_indent(&mut self, indent: u32) -> &mut Self {
        self.indent = Some(indent);
        self
    }

    /// Clears the cell indent level.
    pub fn clear_indent(&mut self) -> &mut Self {
        self.indent = None;
        self
    }

    /// Returns the text rotation in degrees (0-180, where 255=vertical text).
    pub fn text_rotation(&self) -> Option<u32> {
        self.text_rotation
    }

    /// Sets the text rotation in degrees (0-180, where 255=vertical text).
    pub fn set_text_rotation(&mut self, degrees: u32) -> &mut Self {
        self.text_rotation = Some(degrees);
        self
    }

    /// Clears the text rotation.
    pub fn clear_text_rotation(&mut self) -> &mut Self {
        self.text_rotation = None;
        self
    }

    /// Returns whether to shrink text to fit the cell.
    pub fn shrink_to_fit(&self) -> Option<bool> {
        self.shrink_to_fit
    }

    /// Sets whether to shrink text to fit the cell.
    pub fn set_shrink_to_fit(&mut self, value: bool) -> &mut Self {
        self.shrink_to_fit = Some(value);
        self
    }

    /// Clears the shrink-to-fit setting.
    pub fn clear_shrink_to_fit(&mut self) -> &mut Self {
        self.shrink_to_fit = None;
        self
    }

    /// Returns the reading order (0=context, 1=left-to-right, 2=right-to-left).
    pub fn reading_order(&self) -> Option<u32> {
        self.reading_order
    }

    /// Sets the reading order.
    pub fn set_reading_order(&mut self, order: u32) -> &mut Self {
        self.reading_order = Some(order);
        self
    }

    /// Clears the reading order.
    pub fn clear_reading_order(&mut self) -> &mut Self {
        self.reading_order = None;
        self
    }

    pub(crate) fn has_metadata(&self) -> bool {
        self.horizontal.is_some()
            || self.vertical.is_some()
            || self.wrap_text.is_some()
            || self.indent.is_some()
            || self.text_rotation.is_some()
            || self.shrink_to_fit.is_some()
            || self.reading_order.is_some()
    }
}

/// Cell font metadata.
#[derive(Debug, Clone, PartialEq, Default)]
pub struct Font {
    name: Option<String>,
    size: Option<String>,
    bold: Option<bool>,
    italic: Option<bool>,
    underline: Option<bool>,
    color: Option<String>,
    color_ref: Option<ColorReference>,
    strikethrough: Option<bool>,
    double_strikethrough: Option<bool>,
    shadow: Option<bool>,
    outline: Option<bool>,
    subscript: Option<bool>,
    superscript: Option<bool>,
    vertical_align: Option<FontVerticalAlign>,
    font_scheme: Option<FontScheme>,
}

impl Font {
    pub fn new() -> Self {
        Self::default()
    }

    pub fn name(&self) -> Option<&str> {
        self.name.as_deref()
    }

    pub fn set_name(&mut self, name: impl Into<String>) -> &mut Self {
        self.name = Some(name.into());
        self
    }

    pub fn clear_name(&mut self) -> &mut Self {
        self.name = None;
        self
    }

    pub fn size(&self) -> Option<&str> {
        self.size.as_deref()
    }

    pub fn set_size(&mut self, size: impl Into<String>) -> &mut Self {
        self.size = Some(size.into());
        self
    }

    pub fn clear_size(&mut self) -> &mut Self {
        self.size = None;
        self
    }

    pub fn bold(&self) -> Option<bool> {
        self.bold
    }

    pub fn set_bold(&mut self, bold: bool) -> &mut Self {
        self.bold = Some(bold);
        self
    }

    pub fn clear_bold(&mut self) -> &mut Self {
        self.bold = None;
        self
    }

    pub fn italic(&self) -> Option<bool> {
        self.italic
    }

    pub fn set_italic(&mut self, italic: bool) -> &mut Self {
        self.italic = Some(italic);
        self
    }

    pub fn clear_italic(&mut self) -> &mut Self {
        self.italic = None;
        self
    }

    pub fn underline(&self) -> Option<bool> {
        self.underline
    }

    pub fn set_underline(&mut self, underline: bool) -> &mut Self {
        self.underline = Some(underline);
        self
    }

    pub fn clear_underline(&mut self) -> &mut Self {
        self.underline = None;
        self
    }

    pub fn color(&self) -> Option<&str> {
        self.color.as_deref()
    }

    pub fn set_color(&mut self, color: impl Into<String>) -> &mut Self {
        self.color = Some(color.into());
        self
    }

    pub fn clear_color(&mut self) -> &mut Self {
        self.color = None;
        self
    }

    /// Returns the structured color reference if set.
    pub fn color_ref(&self) -> Option<&ColorReference> {
        self.color_ref.as_ref()
    }

    /// Sets the structured color reference.
    pub fn set_color_ref(&mut self, color_ref: ColorReference) -> &mut Self {
        self.color_ref = Some(color_ref);
        self
    }

    /// Clears the structured color reference.
    pub fn clear_color_ref(&mut self) -> &mut Self {
        self.color_ref = None;
        self
    }

    /// Returns whether single strikethrough is enabled.
    pub fn strikethrough(&self) -> Option<bool> {
        self.strikethrough
    }

    /// Sets single strikethrough.
    pub fn set_strikethrough(&mut self, value: bool) -> &mut Self {
        self.strikethrough = Some(value);
        self
    }

    /// Clears single strikethrough.
    pub fn clear_strikethrough(&mut self) -> &mut Self {
        self.strikethrough = None;
        self
    }

    /// Returns whether double strikethrough is enabled.
    pub fn double_strikethrough(&self) -> Option<bool> {
        self.double_strikethrough
    }

    /// Sets double strikethrough.
    pub fn set_double_strikethrough(&mut self, value: bool) -> &mut Self {
        self.double_strikethrough = Some(value);
        self
    }

    /// Clears double strikethrough.
    pub fn clear_double_strikethrough(&mut self) -> &mut Self {
        self.double_strikethrough = None;
        self
    }

    /// Returns whether text shadow is enabled.
    pub fn shadow(&self) -> Option<bool> {
        self.shadow
    }

    /// Sets text shadow.
    pub fn set_shadow(&mut self, value: bool) -> &mut Self {
        self.shadow = Some(value);
        self
    }

    /// Clears text shadow.
    pub fn clear_shadow(&mut self) -> &mut Self {
        self.shadow = None;
        self
    }

    /// Returns whether text outline is enabled.
    pub fn outline(&self) -> Option<bool> {
        self.outline
    }

    /// Sets text outline.
    pub fn set_outline(&mut self, value: bool) -> &mut Self {
        self.outline = Some(value);
        self
    }

    /// Clears text outline.
    pub fn clear_outline(&mut self) -> &mut Self {
        self.outline = None;
        self
    }

    /// Returns whether subscript is enabled.
    pub fn subscript(&self) -> Option<bool> {
        self.subscript
    }

    /// Sets subscript.
    pub fn set_subscript(&mut self, value: bool) -> &mut Self {
        self.subscript = Some(value);
        self
    }

    /// Clears subscript.
    pub fn clear_subscript(&mut self) -> &mut Self {
        self.subscript = None;
        self
    }

    /// Returns whether superscript is enabled.
    pub fn superscript(&self) -> Option<bool> {
        self.superscript
    }

    /// Sets superscript.
    pub fn set_superscript(&mut self, value: bool) -> &mut Self {
        self.superscript = Some(value);
        self
    }

    /// Clears superscript.
    pub fn clear_superscript(&mut self) -> &mut Self {
        self.superscript = None;
        self
    }

    /// Returns the font vertical alignment (superscript/subscript/baseline).
    pub fn vertical_align(&self) -> Option<FontVerticalAlign> {
        self.vertical_align
    }

    /// Sets the font vertical alignment.
    pub fn set_vertical_align(&mut self, value: FontVerticalAlign) -> &mut Self {
        self.vertical_align = Some(value);
        self
    }

    /// Clears the font vertical alignment.
    pub fn clear_vertical_align(&mut self) -> &mut Self {
        self.vertical_align = None;
        self
    }

    /// Returns the font scheme (major/minor/none).
    pub fn font_scheme(&self) -> Option<FontScheme> {
        self.font_scheme
    }

    /// Sets the font scheme.
    pub fn set_font_scheme(&mut self, value: FontScheme) -> &mut Self {
        self.font_scheme = Some(value);
        self
    }

    /// Clears the font scheme.
    pub fn clear_font_scheme(&mut self) -> &mut Self {
        self.font_scheme = None;
        self
    }

    pub(crate) fn has_metadata(&self) -> bool {
        self.name.is_some()
            || self.size.is_some()
            || self.bold.is_some()
            || self.italic.is_some()
            || self.underline.is_some()
            || self.color.is_some()
            || self.color_ref.is_some()
            || self.strikethrough.is_some()
            || self.double_strikethrough.is_some()
            || self.shadow.is_some()
            || self.outline.is_some()
            || self.subscript.is_some()
            || self.superscript.is_some()
            || self.vertical_align.is_some()
            || self.font_scheme.is_some()
    }
}

// ===== Fill patterns =====

/// Pattern fill type for cell backgrounds.
///
/// Maps to the `patternType` attribute of the OOXML `<patternFill>` element.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum PatternFillType {
    None,
    Solid,
    DarkDown,
    DarkGray,
    DarkGrid,
    DarkHorizontal,
    DarkTrellis,
    DarkUp,
    DarkVertical,
    Gray0625,
    Gray125,
    LightDown,
    LightGray,
    LightGrid,
    LightHorizontal,
    LightTrellis,
    LightUp,
    LightVertical,
    MediumGray,
}

impl PatternFillType {
    pub(crate) fn as_xml_value(self) -> &'static str {
        match self {
            Self::None => "none",
            Self::Solid => "solid",
            Self::DarkDown => "darkDown",
            Self::DarkGray => "darkGray",
            Self::DarkGrid => "darkGrid",
            Self::DarkHorizontal => "darkHorizontal",
            Self::DarkTrellis => "darkTrellis",
            Self::DarkUp => "darkUp",
            Self::DarkVertical => "darkVertical",
            Self::Gray0625 => "gray0625",
            Self::Gray125 => "gray125",
            Self::LightDown => "lightDown",
            Self::LightGray => "lightGray",
            Self::LightGrid => "lightGrid",
            Self::LightHorizontal => "lightHorizontal",
            Self::LightTrellis => "lightTrellis",
            Self::LightUp => "lightUp",
            Self::LightVertical => "lightVertical",
            Self::MediumGray => "mediumGray",
        }
    }

    #[allow(unused_qualifications)]
    pub(crate) fn from_xml_value(value: &str) -> Option<Self> {
        match value {
            "none" => Some(Self::None),
            "solid" => Some(Self::Solid),
            "darkDown" => Some(Self::DarkDown),
            "darkGray" => Some(Self::DarkGray),
            "darkGrid" => Some(Self::DarkGrid),
            "darkHorizontal" => Some(Self::DarkHorizontal),
            "darkTrellis" => Some(Self::DarkTrellis),
            "darkUp" => Some(Self::DarkUp),
            "darkVertical" => Some(Self::DarkVertical),
            "gray0625" => Some(Self::Gray0625),
            "gray125" => Some(Self::Gray125),
            "lightDown" => Some(Self::LightDown),
            "lightGray" => Some(Self::LightGray),
            "lightGrid" => Some(Self::LightGrid),
            "lightHorizontal" => Some(Self::LightHorizontal),
            "lightTrellis" => Some(Self::LightTrellis),
            "lightUp" => Some(Self::LightUp),
            "lightVertical" => Some(Self::LightVertical),
            "mediumGray" => Some(Self::MediumGray),
            _ => Option::None,
        }
    }
}

/// A structured pattern fill with typed pattern type and optional foreground/background colors.
///
/// This provides a typed alternative to the string-based pattern on `Fill`.
/// The `fg_color` and `bg_color` use `ColorReference` for richer color support
/// (RGB, theme colors, tints).
#[derive(Debug, Clone, PartialEq)]
pub struct PatternFill {
    pattern_type: PatternFillType,
    fg_color: Option<ColorReference>,
    bg_color: Option<ColorReference>,
}

impl PatternFill {
    /// Creates a new pattern fill with the given type.
    pub fn new(pattern_type: PatternFillType) -> Self {
        Self {
            pattern_type,
            fg_color: None,
            bg_color: None,
        }
    }

    /// Returns the pattern fill type.
    pub fn pattern_type(&self) -> PatternFillType {
        self.pattern_type
    }

    /// Sets the pattern fill type.
    pub fn set_pattern_type(&mut self, pattern_type: PatternFillType) -> &mut Self {
        self.pattern_type = pattern_type;
        self
    }

    /// Returns the foreground color, if set.
    pub fn fg_color(&self) -> Option<&ColorReference> {
        self.fg_color.as_ref()
    }

    /// Sets the foreground color.
    pub fn set_fg_color(&mut self, color: ColorReference) -> &mut Self {
        self.fg_color = Some(color);
        self
    }

    /// Clears the foreground color.
    pub fn clear_fg_color(&mut self) -> &mut Self {
        self.fg_color = None;
        self
    }

    /// Returns the background color, if set.
    pub fn bg_color(&self) -> Option<&ColorReference> {
        self.bg_color.as_ref()
    }

    /// Sets the background color.
    pub fn set_bg_color(&mut self, color: ColorReference) -> &mut Self {
        self.bg_color = Some(color);
        self
    }

    /// Clears the background color.
    pub fn clear_bg_color(&mut self) -> &mut Self {
        self.bg_color = None;
        self
    }
}

/// Cell fill metadata.
#[derive(Debug, Clone, PartialEq, Default)]
pub struct Fill {
    pattern: Option<String>,
    foreground_color: Option<String>,
    background_color: Option<String>,
    pattern_fill: Option<PatternFill>,
}

impl Fill {
    pub fn new() -> Self {
        Self::default()
    }

    pub fn pattern(&self) -> Option<&str> {
        self.pattern.as_deref()
    }

    pub fn set_pattern(&mut self, pattern: impl Into<String>) -> &mut Self {
        self.pattern = Some(pattern.into());
        self
    }

    pub fn clear_pattern(&mut self) -> &mut Self {
        self.pattern = None;
        self
    }

    pub fn foreground_color(&self) -> Option<&str> {
        self.foreground_color.as_deref()
    }

    pub fn set_foreground_color(&mut self, color: impl Into<String>) -> &mut Self {
        self.foreground_color = Some(color.into());
        self
    }

    pub fn clear_foreground_color(&mut self) -> &mut Self {
        self.foreground_color = None;
        self
    }

    pub fn background_color(&self) -> Option<&str> {
        self.background_color.as_deref()
    }

    pub fn set_background_color(&mut self, color: impl Into<String>) -> &mut Self {
        self.background_color = Some(color.into());
        self
    }

    pub fn clear_background_color(&mut self) -> &mut Self {
        self.background_color = None;
        self
    }

    /// Returns the structured pattern fill, if set.
    pub fn pattern_fill(&self) -> Option<&PatternFill> {
        self.pattern_fill.as_ref()
    }

    /// Sets a structured pattern fill on this fill.
    pub fn set_pattern_fill(&mut self, pattern_fill: PatternFill) -> &mut Self {
        self.pattern_fill = Some(pattern_fill);
        self
    }

    /// Clears the structured pattern fill.
    pub fn clear_pattern_fill(&mut self) -> &mut Self {
        self.pattern_fill = None;
        self
    }

    pub(crate) fn has_metadata(&self) -> bool {
        self.pattern.is_some()
            || self.foreground_color.is_some()
            || self.background_color.is_some()
            || self.pattern_fill.is_some()
    }
}

/// Cell border side metadata.
#[derive(Debug, Clone, PartialEq, Default)]
pub struct BorderSide {
    style: Option<String>,
    color: Option<String>,
    color_ref: Option<ColorReference>,
}

impl BorderSide {
    pub fn new() -> Self {
        Self::default()
    }

    pub fn style(&self) -> Option<&str> {
        self.style.as_deref()
    }

    pub fn set_style(&mut self, style: impl Into<String>) -> &mut Self {
        self.style = Some(style.into());
        self
    }

    pub fn clear_style(&mut self) -> &mut Self {
        self.style = None;
        self
    }

    pub fn color(&self) -> Option<&str> {
        self.color.as_deref()
    }

    pub fn set_color(&mut self, color: impl Into<String>) -> &mut Self {
        self.color = Some(color.into());
        self
    }

    pub fn clear_color(&mut self) -> &mut Self {
        self.color = None;
        self
    }

    /// Returns the structured color reference if set.
    pub fn color_ref(&self) -> Option<&ColorReference> {
        self.color_ref.as_ref()
    }

    /// Sets the structured color reference.
    pub fn set_color_ref(&mut self, color_ref: ColorReference) -> &mut Self {
        self.color_ref = Some(color_ref);
        self
    }

    /// Clears the structured color reference.
    pub fn clear_color_ref(&mut self) -> &mut Self {
        self.color_ref = None;
        self
    }

    pub(crate) fn has_metadata(&self) -> bool {
        self.style.is_some() || self.color.is_some() || self.color_ref.is_some()
    }
}

/// Cell border metadata.
#[derive(Debug, Clone, PartialEq, Default)]
pub struct Border {
    left: Option<BorderSide>,
    right: Option<BorderSide>,
    top: Option<BorderSide>,
    bottom: Option<BorderSide>,
    diagonal: Option<BorderSide>,
    diagonal_up: Option<bool>,
    diagonal_down: Option<bool>,
}

impl Border {
    pub fn new() -> Self {
        Self::default()
    }

    pub fn left(&self) -> Option<&BorderSide> {
        self.left.as_ref()
    }

    pub fn set_left(&mut self, side: BorderSide) -> &mut Self {
        self.left = Some(side);
        self
    }

    pub fn clear_left(&mut self) -> &mut Self {
        self.left = None;
        self
    }

    pub fn right(&self) -> Option<&BorderSide> {
        self.right.as_ref()
    }

    pub fn set_right(&mut self, side: BorderSide) -> &mut Self {
        self.right = Some(side);
        self
    }

    pub fn clear_right(&mut self) -> &mut Self {
        self.right = None;
        self
    }

    pub fn top(&self) -> Option<&BorderSide> {
        self.top.as_ref()
    }

    pub fn set_top(&mut self, side: BorderSide) -> &mut Self {
        self.top = Some(side);
        self
    }

    pub fn clear_top(&mut self) -> &mut Self {
        self.top = None;
        self
    }

    pub fn bottom(&self) -> Option<&BorderSide> {
        self.bottom.as_ref()
    }

    pub fn set_bottom(&mut self, side: BorderSide) -> &mut Self {
        self.bottom = Some(side);
        self
    }

    pub fn clear_bottom(&mut self) -> &mut Self {
        self.bottom = None;
        self
    }

    /// Returns the diagonal border side, if set.
    pub fn diagonal(&self) -> Option<&BorderSide> {
        self.diagonal.as_ref()
    }

    /// Sets the diagonal border side.
    pub fn set_diagonal(&mut self, side: BorderSide) -> &mut Self {
        self.diagonal = Some(side);
        self
    }

    /// Clears the diagonal border side.
    pub fn clear_diagonal(&mut self) -> &mut Self {
        self.diagonal = None;
        self
    }

    /// Returns whether the diagonal border goes from bottom-left to top-right.
    pub fn diagonal_up(&self) -> Option<bool> {
        self.diagonal_up
    }

    /// Sets whether the diagonal border goes from bottom-left to top-right.
    pub fn set_diagonal_up(&mut self, value: bool) -> &mut Self {
        self.diagonal_up = Some(value);
        self
    }

    /// Clears the diagonal up setting.
    pub fn clear_diagonal_up(&mut self) -> &mut Self {
        self.diagonal_up = None;
        self
    }

    /// Returns whether the diagonal border goes from top-left to bottom-right.
    pub fn diagonal_down(&self) -> Option<bool> {
        self.diagonal_down
    }

    /// Sets whether the diagonal border goes from top-left to bottom-right.
    pub fn set_diagonal_down(&mut self, value: bool) -> &mut Self {
        self.diagonal_down = Some(value);
        self
    }

    /// Clears the diagonal down setting.
    pub fn clear_diagonal_down(&mut self) -> &mut Self {
        self.diagonal_down = None;
        self
    }

    pub(crate) fn has_metadata(&self) -> bool {
        self.left.as_ref().is_some_and(BorderSide::has_metadata)
            || self.right.as_ref().is_some_and(BorderSide::has_metadata)
            || self.top.as_ref().is_some_and(BorderSide::has_metadata)
            || self.bottom.as_ref().is_some_and(BorderSide::has_metadata)
            || self.diagonal.as_ref().is_some_and(BorderSide::has_metadata)
            || self.diagonal_up.is_some()
            || self.diagonal_down.is_some()
    }
}

/// Cell-level protection settings.
///
/// In Excel, cells are locked by default. The `locked` and `hidden` flags
/// only take effect when the sheet is protected via `SheetProtection`.
#[derive(Debug, Clone, PartialEq, Eq, Default)]
pub struct CellProtection {
    locked: Option<bool>,
    hidden: Option<bool>,
}

impl CellProtection {
    /// Creates a new cell protection with default settings.
    pub fn new() -> Self {
        Self::default()
    }

    /// Returns whether the cell is locked (default is true in Excel).
    pub fn locked(&self) -> Option<bool> {
        self.locked
    }

    /// Sets whether the cell is locked.
    pub fn set_locked(&mut self, value: bool) -> &mut Self {
        self.locked = Some(value);
        self
    }

    /// Clears the locked setting.
    pub fn clear_locked(&mut self) -> &mut Self {
        self.locked = None;
        self
    }

    /// Returns whether the cell's formula is hidden in the formula bar.
    pub fn hidden(&self) -> Option<bool> {
        self.hidden
    }

    /// Sets whether the cell's formula is hidden.
    pub fn set_hidden(&mut self, value: bool) -> &mut Self {
        self.hidden = Some(value);
        self
    }

    /// Clears the hidden setting.
    pub fn clear_hidden(&mut self) -> &mut Self {
        self.hidden = None;
        self
    }

    /// Returns true if any protection property is set.
    pub(crate) fn has_metadata(&self) -> bool {
        self.locked.is_some() || self.hidden.is_some()
    }
}

/// Cell style metadata.
#[derive(Debug, Clone, PartialEq, Default)]
pub struct Style {
    number_format: Option<String>,
    /// Custom format string (e.g. "#,##0.00", "yyyy-mm-dd").
    ///
    /// This is stored separately from `number_format` so that the user
    /// can set an arbitrary format code that maps to a custom `numFmt`
    /// in the styles XML.  When both `number_format` and `custom_format`
    /// are set, `custom_format` takes precedence during serialization.
    custom_format: Option<String>,
    alignment: Option<Alignment>,
    font: Option<Font>,
    fill: Option<Fill>,
    gradient_fill: Option<GradientFill>,
    border: Option<Border>,
    protection: Option<CellProtection>,
}

impl Style {
    pub fn new() -> Self {
        Self::default()
    }

    pub fn number_format(&self) -> Option<&str> {
        self.number_format.as_deref()
    }

    pub fn set_number_format(&mut self, number_format: impl Into<String>) -> &mut Self {
        self.number_format = Some(number_format.into());
        self
    }

    pub fn clear_number_format(&mut self) -> &mut Self {
        self.number_format = None;
        self
    }

    pub fn alignment(&self) -> Option<&Alignment> {
        self.alignment.as_ref()
    }

    pub fn set_alignment(&mut self, alignment: Alignment) -> &mut Self {
        self.alignment = Some(alignment);
        self
    }

    pub fn clear_alignment(&mut self) -> &mut Self {
        self.alignment = None;
        self
    }

    pub fn font(&self) -> Option<&Font> {
        self.font.as_ref()
    }

    pub fn set_font(&mut self, font: Font) -> &mut Self {
        self.font = Some(font);
        self
    }

    pub fn clear_font(&mut self) -> &mut Self {
        self.font = None;
        self
    }

    pub fn fill(&self) -> Option<&Fill> {
        self.fill.as_ref()
    }

    pub fn set_fill(&mut self, fill: Fill) -> &mut Self {
        self.fill = Some(fill);
        self
    }

    pub fn clear_fill(&mut self) -> &mut Self {
        self.fill = None;
        self
    }

    pub fn border(&self) -> Option<&Border> {
        self.border.as_ref()
    }

    pub fn set_border(&mut self, border: Border) -> &mut Self {
        self.border = Some(border);
        self
    }

    pub fn clear_border(&mut self) -> &mut Self {
        self.border = None;
        self
    }

    // ---- Cell-level protection ----

    /// Returns the cell protection settings, if set.
    pub fn protection(&self) -> Option<&CellProtection> {
        self.protection.as_ref()
    }

    /// Sets the cell protection settings.
    pub fn set_protection(&mut self, protection: CellProtection) -> &mut Self {
        self.protection = Some(protection);
        self
    }

    /// Clears the cell protection settings.
    pub fn clear_protection(&mut self) -> &mut Self {
        self.protection = None;
        self
    }

    // ---- Feature 6: Custom number format ----

    /// Returns the custom format string if set (e.g. "#,##0.00", "yyyy-mm-dd").
    pub fn custom_format(&self) -> Option<&str> {
        self.custom_format.as_deref()
    }

    /// Sets a custom number format string.
    ///
    /// Custom format strings are arbitrary Excel format codes like `"#,##0.00"`,
    /// `"yyyy-mm-dd"`, `"0.00%"`, etc. They are serialized as `numFmt` entries
    /// in the styles XML with IDs starting at 164.
    pub fn set_custom_format(&mut self, format: impl Into<String>) -> &mut Self {
        self.custom_format = Some(format.into());
        self
    }

    /// Clears the custom format string.
    pub fn clear_custom_format(&mut self) -> &mut Self {
        self.custom_format = None;
        self
    }

    // ---- Feature 7: Gradient fill ----

    /// Returns the gradient fill, if set.
    pub fn gradient_fill(&self) -> Option<&GradientFill> {
        self.gradient_fill.as_ref()
    }

    /// Sets a gradient fill on this style.
    pub fn set_gradient_fill(&mut self, gradient: GradientFill) -> &mut Self {
        self.gradient_fill = Some(gradient);
        self
    }

    /// Clears the gradient fill.
    pub fn clear_gradient_fill(&mut self) -> &mut Self {
        self.gradient_fill = None;
        self
    }
}

/// Workbook style table indexed by cell `style_id`.
#[derive(Debug, Clone, PartialEq)]
pub struct StyleTable {
    styles: Vec<Style>,
}

impl Default for StyleTable {
    fn default() -> Self {
        Self::new()
    }
}

impl StyleTable {
    pub fn new() -> Self {
        Self {
            styles: vec![Style::new()],
        }
    }

    pub fn len(&self) -> usize {
        self.styles.len()
    }

    pub fn is_empty(&self) -> bool {
        self.styles.is_empty()
    }

    pub fn styles(&self) -> &[Style] {
        self.styles.as_slice()
    }

    pub fn styles_mut(&mut self) -> &mut [Style] {
        self.styles.as_mut_slice()
    }

    pub fn style(&self, style_id: u32) -> Option<&Style> {
        self.styles.get(usize::try_from(style_id).ok()?)
    }

    pub fn style_mut(&mut self, style_id: u32) -> Option<&mut Style> {
        self.styles.get_mut(usize::try_from(style_id).ok()?)
    }

    pub fn add_style(&mut self, style: Style) -> Result<u32> {
        let style_id = u32::try_from(self.styles.len()).map_err(|_| {
            XlsxError::InvalidWorkbookState(
                "style table cannot exceed u32::MAX entries".to_string(),
            )
        })?;
        self.styles.push(style);
        Ok(style_id)
    }

    pub fn clear_custom_styles(&mut self) -> &mut Self {
        self.styles.truncate(1);
        self
    }

    pub(crate) fn from_styles(styles: Vec<Style>) -> Self {
        if styles.is_empty() {
            return Self::new();
        }
        Self { styles }
    }

    pub(crate) fn ensure_len(&mut self, len: usize) {
        if self.styles.len() < len {
            self.styles.resize(len, Style::new());
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn style_and_alignment_setters_work() {
        let mut style = Style::new();
        let mut alignment = Alignment::new();
        let mut font = Font::new();
        let mut fill = Fill::new();
        let mut left_border = BorderSide::new();
        let mut border = Border::new();
        alignment
            .set_horizontal(HorizontalAlignment::Center)
            .set_vertical(VerticalAlignment::Top)
            .set_wrap_text(true);
        font.set_name("Aptos")
            .set_size("12")
            .set_bold(true)
            .set_italic(false)
            .set_underline(true)
            .set_color("FFFF0000");
        fill.set_pattern("solid")
            .set_foreground_color("FFFFFF00")
            .set_background_color("FF000000");
        left_border.set_style("thin").set_color("FF00FF00");
        border.set_left(left_border.clone());

        style
            .set_number_format("yyyy-mm-dd")
            .set_alignment(alignment.clone())
            .set_font(font.clone())
            .set_fill(fill.clone())
            .set_border(border.clone());

        assert_eq!(style.number_format(), Some("yyyy-mm-dd"));
        assert_eq!(style.alignment(), Some(&alignment));
        assert_eq!(style.font(), Some(&font));
        assert_eq!(style.fill(), Some(&fill));
        assert_eq!(style.border(), Some(&border));
        assert!(alignment.has_metadata());
        assert!(font.has_metadata());
        assert!(fill.has_metadata());
        assert!(left_border.has_metadata());
        assert!(border.has_metadata());

        style
            .clear_number_format()
            .clear_alignment()
            .clear_font()
            .clear_fill()
            .clear_border();
        assert_eq!(style.number_format(), None);
        assert_eq!(style.alignment(), None);
        assert_eq!(style.font(), None);
        assert_eq!(style.fill(), None);
        assert_eq!(style.border(), None);
    }

    #[test]
    fn style_table_assigns_style_ids() {
        let mut styles = StyleTable::new();
        assert_eq!(styles.len(), 1);

        let mut style = Style::new();
        style.set_number_format("0.00");
        let style_id = styles.add_style(style).expect("style id should fit u32");

        assert_eq!(style_id, 1);
        assert_eq!(
            styles.style(style_id).and_then(Style::number_format),
            Some("0.00")
        );

        styles.clear_custom_styles();
        assert_eq!(styles.len(), 1);
    }

    // ===== Feature 10: Cell indentation and text rotation =====

    #[test]
    fn alignment_indent_accessors_work() {
        let mut alignment = Alignment::new();
        assert!(alignment.indent().is_none());

        alignment.set_indent(2);
        assert_eq!(alignment.indent(), Some(2));
        assert!(alignment.has_metadata());

        alignment.clear_indent();
        assert!(alignment.indent().is_none());
    }

    #[test]
    fn alignment_text_rotation_accessors_work() {
        let mut alignment = Alignment::new();
        assert!(alignment.text_rotation().is_none());

        alignment.set_text_rotation(90);
        assert_eq!(alignment.text_rotation(), Some(90));
        assert!(alignment.has_metadata());

        alignment.set_text_rotation(255); // vertical text
        assert_eq!(alignment.text_rotation(), Some(255));

        alignment.clear_text_rotation();
        assert!(alignment.text_rotation().is_none());
    }

    #[test]
    fn alignment_shrink_to_fit_accessors_work() {
        let mut alignment = Alignment::new();
        assert!(alignment.shrink_to_fit().is_none());

        alignment.set_shrink_to_fit(true);
        assert_eq!(alignment.shrink_to_fit(), Some(true));
        assert!(alignment.has_metadata());

        alignment.set_shrink_to_fit(false);
        assert_eq!(alignment.shrink_to_fit(), Some(false));

        alignment.clear_shrink_to_fit();
        assert!(alignment.shrink_to_fit().is_none());
    }

    #[test]
    fn alignment_reading_order_accessors_work() {
        let mut alignment = Alignment::new();
        assert!(alignment.reading_order().is_none());

        alignment.set_reading_order(1); // left-to-right
        assert_eq!(alignment.reading_order(), Some(1));
        assert!(alignment.has_metadata());

        alignment.set_reading_order(2); // right-to-left
        assert_eq!(alignment.reading_order(), Some(2));

        alignment.clear_reading_order();
        assert!(alignment.reading_order().is_none());
    }

    #[test]
    fn alignment_has_metadata_checks_all_fields() {
        let mut alignment = Alignment::new();
        assert!(!alignment.has_metadata());

        // Test each field individually triggers has_metadata
        let mut a1 = Alignment::new();
        a1.set_indent(1);
        assert!(a1.has_metadata());

        let mut a2 = Alignment::new();
        a2.set_text_rotation(45);
        assert!(a2.has_metadata());

        let mut a3 = Alignment::new();
        a3.set_shrink_to_fit(true);
        assert!(a3.has_metadata());

        let mut a4 = Alignment::new();
        a4.set_reading_order(0);
        assert!(a4.has_metadata());

        // Original fields still work
        alignment.set_horizontal(HorizontalAlignment::Left);
        assert!(alignment.has_metadata());
    }

    // ===== Feature 11: Border diagonal =====

    #[test]
    fn border_diagonal_accessors_work() {
        let mut border = Border::new();
        assert!(border.diagonal().is_none());
        assert!(border.diagonal_up().is_none());
        assert!(border.diagonal_down().is_none());

        let mut diagonal = BorderSide::new();
        diagonal.set_style("thin").set_color("FFFF0000");
        border.set_diagonal(diagonal);
        border.set_diagonal_up(true);
        border.set_diagonal_down(false);

        assert!(border.diagonal().is_some());
        assert_eq!(border.diagonal().unwrap().style(), Some("thin"));
        assert_eq!(border.diagonal().unwrap().color(), Some("FFFF0000"));
        assert_eq!(border.diagonal_up(), Some(true));
        assert_eq!(border.diagonal_down(), Some(false));
        assert!(border.has_metadata());
    }

    #[test]
    fn border_diagonal_clear_works() {
        let mut border = Border::new();
        let mut diag = BorderSide::new();
        diag.set_style("thin");
        border.set_diagonal(diag);
        border.set_diagonal_up(true);
        border.set_diagonal_down(true);

        border.clear_diagonal();
        border.clear_diagonal_up();
        border.clear_diagonal_down();

        assert!(border.diagonal().is_none());
        assert!(border.diagonal_up().is_none());
        assert!(border.diagonal_down().is_none());
        assert!(!border.has_metadata());
    }

    #[test]
    fn border_has_metadata_includes_diagonal_fields() {
        let mut border = Border::new();
        assert!(!border.has_metadata());

        // diagonal_up alone triggers has_metadata
        border.set_diagonal_up(true);
        assert!(border.has_metadata());

        let mut border2 = Border::new();
        border2.set_diagonal_down(true);
        assert!(border2.has_metadata());

        let mut border3 = Border::new();
        let mut diag = BorderSide::new();
        diag.set_style("double");
        border3.set_diagonal(diag);
        assert!(border3.has_metadata());
    }

    // ===== Feature 6: Custom number format strings =====

    #[test]
    fn custom_format_accessors_work() {
        let mut style = Style::new();
        assert!(style.custom_format().is_none());

        style.set_custom_format("#,##0.00");
        assert_eq!(style.custom_format(), Some("#,##0.00"));

        style.set_custom_format("yyyy-mm-dd");
        assert_eq!(style.custom_format(), Some("yyyy-mm-dd"));

        style.clear_custom_format();
        assert!(style.custom_format().is_none());
    }

    #[test]
    fn custom_format_coexists_with_number_format() {
        let mut style = Style::new();
        style.set_number_format("General");
        style.set_custom_format("0.00%");

        assert_eq!(style.number_format(), Some("General"));
        assert_eq!(style.custom_format(), Some("0.00%"));
    }

    // ===== Feature 7: Gradient fill =====

    #[test]
    fn gradient_fill_linear_construction() {
        let mut gradient = GradientFill::linear(90.0);
        assert_eq!(gradient.gradient_type(), GradientFillType::Linear);
        assert_eq!(gradient.degree(), Some(90.0));
        assert!(gradient.stops().is_empty());

        gradient.add_stop(0.0, "FF000000").add_stop(1.0, "FFFFFFFF");
        assert_eq!(gradient.stops().len(), 2);
        assert_eq!(gradient.stops()[0].position(), 0.0);
        assert_eq!(gradient.stops()[0].color(), "FF000000");
        assert_eq!(gradient.stops()[1].position(), 1.0);
        assert_eq!(gradient.stops()[1].color(), "FFFFFFFF");
    }

    #[test]
    fn gradient_fill_path_construction() {
        let mut gradient = GradientFill::path();
        assert_eq!(gradient.gradient_type(), GradientFillType::Path);
        assert!(gradient.degree().is_none());

        gradient
            .add_stop(0.0, "FFFF0000")
            .add_stop(0.5, "FF00FF00")
            .add_stop(1.0, "FF0000FF");
        assert_eq!(gradient.stops().len(), 3);
        assert_eq!(gradient.stops()[1].position(), 0.5);
    }

    #[test]
    fn gradient_fill_clear_and_modify() {
        let mut gradient = GradientFill::linear(45.0);
        gradient.add_stop(0.0, "FF000000");

        gradient.clear_degree();
        assert!(gradient.degree().is_none());

        gradient.set_degree(180.0);
        assert_eq!(gradient.degree(), Some(180.0));

        gradient.set_gradient_type(GradientFillType::Path);
        assert_eq!(gradient.gradient_type(), GradientFillType::Path);

        gradient.clear_stops();
        assert!(gradient.stops().is_empty());
    }

    #[test]
    fn gradient_stop_position_clamped() {
        let stop = GradientStop::new(-0.5, "FF000000");
        assert_eq!(stop.position(), 0.0);

        let stop2 = GradientStop::new(1.5, "FFFFFFFF");
        assert_eq!(stop2.position(), 1.0);
    }

    #[test]
    fn style_gradient_fill_set_and_clear() {
        let mut style = Style::new();
        assert!(style.gradient_fill().is_none());

        let mut gradient = GradientFill::linear(90.0);
        gradient.add_stop(0.0, "FF000000").add_stop(1.0, "FFFFFFFF");
        style.set_gradient_fill(gradient);

        assert!(style.gradient_fill().is_some());
        assert_eq!(
            style.gradient_fill().unwrap().gradient_type(),
            GradientFillType::Linear
        );
        assert_eq!(style.gradient_fill().unwrap().stops().len(), 2);

        style.clear_gradient_fill();
        assert!(style.gradient_fill().is_none());
    }

    #[test]
    fn gradient_fill_type_xml_roundtrip() {
        assert_eq!(GradientFillType::Linear.as_xml_value(), "linear");
        assert_eq!(GradientFillType::Path.as_xml_value(), "path");
        assert_eq!(
            GradientFillType::from_xml_value("linear"),
            Some(GradientFillType::Linear)
        );
        assert_eq!(
            GradientFillType::from_xml_value("path"),
            Some(GradientFillType::Path)
        );
        assert_eq!(GradientFillType::from_xml_value("unknown"), None);
    }

    // ===== Feature 8: Border diagonal support =====

    // (Border diagonal tests are already above)

    // ===== Feature 9: Theme colors =====

    #[test]
    fn theme_color_index_roundtrip() {
        for (theme, expected_idx) in [
            (ThemeColor::Dark1, 0),
            (ThemeColor::Light1, 1),
            (ThemeColor::Dark2, 2),
            (ThemeColor::Light2, 3),
            (ThemeColor::Accent1, 4),
            (ThemeColor::Accent2, 5),
            (ThemeColor::Accent3, 6),
            (ThemeColor::Accent4, 7),
            (ThemeColor::Accent5, 8),
            (ThemeColor::Accent6, 9),
            (ThemeColor::Hyperlink, 10),
            (ThemeColor::FollowedHyperlink, 11),
        ] {
            assert_eq!(theme.index(), expected_idx);
            assert_eq!(ThemeColor::from_index(expected_idx), Some(theme));
        }
    }

    #[test]
    fn theme_color_from_index_returns_none_for_invalid() {
        assert!(ThemeColor::from_index(12).is_none());
        assert!(ThemeColor::from_index(255).is_none());
    }

    #[test]
    fn theme_color_names() {
        assert_eq!(ThemeColor::Dark1.name(), "dk1");
        assert_eq!(ThemeColor::Light1.name(), "lt1");
        assert_eq!(ThemeColor::Accent1.name(), "accent1");
        assert_eq!(ThemeColor::Hyperlink.name(), "hlink");
        assert_eq!(ThemeColor::FollowedHyperlink.name(), "folHlink");
    }

    #[test]
    fn color_reference_from_rgb() {
        let color = ColorReference::from_rgb("FFFF0000");
        assert_eq!(color.rgb(), Some("FFFF0000"));
        assert!(color.theme().is_none());
        assert!(color.tint().is_none());
    }

    #[test]
    fn color_reference_from_theme() {
        let color = ColorReference::from_theme(ThemeColor::Accent1);
        assert!(color.rgb().is_none());
        assert_eq!(color.theme(), Some(ThemeColor::Accent1));
        assert!(color.tint().is_none());
    }

    #[test]
    fn color_reference_from_theme_with_tint() {
        let color = ColorReference::from_theme_with_tint(ThemeColor::Dark1, 0.5);
        assert!(color.rgb().is_none());
        assert_eq!(color.theme(), Some(ThemeColor::Dark1));
        assert_eq!(color.tint(), Some(0.5));
    }

    #[test]
    fn color_reference_mutators() {
        let mut color = ColorReference::from_rgb("FF000000");
        color.set_theme(ThemeColor::Light2).set_tint(-0.25);
        assert_eq!(color.theme(), Some(ThemeColor::Light2));
        assert_eq!(color.tint(), Some(-0.25));

        color.clear_tint();
        assert!(color.tint().is_none());

        color.set_rgb("FF00FF00");
        assert_eq!(color.rgb(), Some("FF00FF00"));
    }

    // ===== Font effects =====

    #[test]
    fn font_strikethrough_accessors() {
        let mut font = Font::new();
        assert!(font.strikethrough().is_none());

        font.set_strikethrough(true);
        assert_eq!(font.strikethrough(), Some(true));
        assert!(font.has_metadata());

        font.clear_strikethrough();
        assert!(font.strikethrough().is_none());
    }

    #[test]
    fn font_double_strikethrough_accessors() {
        let mut font = Font::new();
        assert!(font.double_strikethrough().is_none());

        font.set_double_strikethrough(true);
        assert_eq!(font.double_strikethrough(), Some(true));
        assert!(font.has_metadata());

        font.clear_double_strikethrough();
        assert!(font.double_strikethrough().is_none());
    }

    #[test]
    fn font_shadow_accessors() {
        let mut font = Font::new();
        assert!(font.shadow().is_none());

        font.set_shadow(true);
        assert_eq!(font.shadow(), Some(true));
        assert!(font.has_metadata());

        font.clear_shadow();
        assert!(font.shadow().is_none());
    }

    #[test]
    fn font_outline_accessors() {
        let mut font = Font::new();
        assert!(font.outline().is_none());

        font.set_outline(true);
        assert_eq!(font.outline(), Some(true));
        assert!(font.has_metadata());

        font.clear_outline();
        assert!(font.outline().is_none());
    }

    #[test]
    fn font_subscript_accessors() {
        let mut font = Font::new();
        assert!(font.subscript().is_none());

        font.set_subscript(true);
        assert_eq!(font.subscript(), Some(true));
        assert!(font.has_metadata());

        font.clear_subscript();
        assert!(font.subscript().is_none());
    }

    #[test]
    fn font_superscript_accessors() {
        let mut font = Font::new();
        assert!(font.superscript().is_none());

        font.set_superscript(true);
        assert_eq!(font.superscript(), Some(true));
        assert!(font.has_metadata());

        font.clear_superscript();
        assert!(font.superscript().is_none());
    }

    #[test]
    fn font_effects_has_metadata() {
        let mut font = Font::new();
        assert!(!font.has_metadata());

        font.set_strikethrough(true);
        assert!(font.has_metadata());

        let mut font2 = Font::new();
        font2.set_shadow(true);
        assert!(font2.has_metadata());

        let mut font3 = Font::new();
        font3.set_superscript(true);
        assert!(font3.has_metadata());
    }

    // ===== Cell-level protection =====

    #[test]
    fn cell_protection_accessors() {
        let mut protection = CellProtection::new();
        assert!(protection.locked().is_none());
        assert!(protection.hidden().is_none());
        assert!(!protection.has_metadata());

        protection.set_locked(false);
        assert_eq!(protection.locked(), Some(false));
        assert!(protection.has_metadata());

        protection.set_hidden(true);
        assert_eq!(protection.hidden(), Some(true));

        protection.clear_locked();
        protection.clear_hidden();
        assert!(!protection.has_metadata());
    }

    #[test]
    fn style_protection_set_and_clear() {
        let mut style = Style::new();
        assert!(style.protection().is_none());

        let mut protection = CellProtection::new();
        protection.set_locked(false);
        style.set_protection(protection);

        assert!(style.protection().is_some());
        assert_eq!(style.protection().unwrap().locked(), Some(false));

        style.clear_protection();
        assert!(style.protection().is_none());
    }

    // ===== FontVerticalAlign enum =====

    #[test]
    fn font_vertical_align_xml_roundtrip() {
        assert_eq!(FontVerticalAlign::Superscript.as_xml_value(), "superscript");
        assert_eq!(FontVerticalAlign::Subscript.as_xml_value(), "subscript");
        assert_eq!(FontVerticalAlign::Baseline.as_xml_value(), "baseline");

        assert_eq!(
            FontVerticalAlign::from_xml_value("superscript"),
            Some(FontVerticalAlign::Superscript)
        );
        assert_eq!(
            FontVerticalAlign::from_xml_value("subscript"),
            Some(FontVerticalAlign::Subscript)
        );
        assert_eq!(
            FontVerticalAlign::from_xml_value("baseline"),
            Some(FontVerticalAlign::Baseline)
        );
        assert_eq!(FontVerticalAlign::from_xml_value("unknown"), None);
    }

    #[test]
    fn font_vertical_align_accessors() {
        let mut font = Font::new();
        assert!(font.vertical_align().is_none());

        font.set_vertical_align(FontVerticalAlign::Superscript);
        assert_eq!(font.vertical_align(), Some(FontVerticalAlign::Superscript));
        assert!(font.has_metadata());

        font.set_vertical_align(FontVerticalAlign::Subscript);
        assert_eq!(font.vertical_align(), Some(FontVerticalAlign::Subscript));

        font.set_vertical_align(FontVerticalAlign::Baseline);
        assert_eq!(font.vertical_align(), Some(FontVerticalAlign::Baseline));

        font.clear_vertical_align();
        assert!(font.vertical_align().is_none());
    }

    // ===== FontScheme enum =====

    #[test]
    fn font_scheme_xml_roundtrip() {
        assert_eq!(FontScheme::Major.as_xml_value(), "major");
        assert_eq!(FontScheme::Minor.as_xml_value(), "minor");
        assert_eq!(FontScheme::None.as_xml_value(), "none");

        assert_eq!(FontScheme::from_xml_value("major"), Some(FontScheme::Major));
        assert_eq!(FontScheme::from_xml_value("minor"), Some(FontScheme::Minor));
        assert_eq!(FontScheme::from_xml_value("none"), Some(FontScheme::None));
        assert_eq!(FontScheme::from_xml_value("unknown"), None);
    }

    #[test]
    fn font_scheme_accessors() {
        let mut font = Font::new();
        assert!(font.font_scheme().is_none());

        font.set_font_scheme(FontScheme::Minor);
        assert_eq!(font.font_scheme(), Some(FontScheme::Minor));
        assert!(font.has_metadata());

        font.set_font_scheme(FontScheme::Major);
        assert_eq!(font.font_scheme(), Some(FontScheme::Major));

        font.clear_font_scheme();
        assert!(font.font_scheme().is_none());
    }

    // ===== PatternFillType enum =====

    #[test]
    fn pattern_fill_type_xml_roundtrip() {
        let variants = [
            (PatternFillType::None, "none"),
            (PatternFillType::Solid, "solid"),
            (PatternFillType::DarkDown, "darkDown"),
            (PatternFillType::DarkGray, "darkGray"),
            (PatternFillType::DarkGrid, "darkGrid"),
            (PatternFillType::DarkHorizontal, "darkHorizontal"),
            (PatternFillType::DarkTrellis, "darkTrellis"),
            (PatternFillType::DarkUp, "darkUp"),
            (PatternFillType::DarkVertical, "darkVertical"),
            (PatternFillType::Gray0625, "gray0625"),
            (PatternFillType::Gray125, "gray125"),
            (PatternFillType::LightDown, "lightDown"),
            (PatternFillType::LightGray, "lightGray"),
            (PatternFillType::LightGrid, "lightGrid"),
            (PatternFillType::LightHorizontal, "lightHorizontal"),
            (PatternFillType::LightTrellis, "lightTrellis"),
            (PatternFillType::LightUp, "lightUp"),
            (PatternFillType::LightVertical, "lightVertical"),
            (PatternFillType::MediumGray, "mediumGray"),
        ];
        for (variant, expected) in &variants {
            assert_eq!(variant.as_xml_value(), *expected);
            assert_eq!(PatternFillType::from_xml_value(expected), Some(*variant));
        }
        assert_eq!(PatternFillType::from_xml_value("invalid"), None);
    }

    // ===== PatternFill struct =====

    #[test]
    fn pattern_fill_construction_and_accessors() {
        let mut pf = PatternFill::new(PatternFillType::Solid);
        assert_eq!(pf.pattern_type(), PatternFillType::Solid);
        assert!(pf.fg_color().is_none());
        assert!(pf.bg_color().is_none());

        pf.set_fg_color(ColorReference::from_rgb("FFFF0000"));
        assert_eq!(pf.fg_color().unwrap().rgb(), Some("FFFF0000"));

        pf.set_bg_color(ColorReference::from_theme(ThemeColor::Light1));
        assert_eq!(pf.bg_color().unwrap().theme(), Some(ThemeColor::Light1));

        pf.set_pattern_type(PatternFillType::DarkGray);
        assert_eq!(pf.pattern_type(), PatternFillType::DarkGray);

        pf.clear_fg_color();
        assert!(pf.fg_color().is_none());

        pf.clear_bg_color();
        assert!(pf.bg_color().is_none());
    }

    #[test]
    fn fill_pattern_fill_set_and_clear() {
        let mut fill = Fill::new();
        assert!(fill.pattern_fill().is_none());

        let pf = PatternFill::new(PatternFillType::Solid);
        fill.set_pattern_fill(pf);
        assert!(fill.pattern_fill().is_some());
        assert_eq!(
            fill.pattern_fill().unwrap().pattern_type(),
            PatternFillType::Solid
        );
        assert!(fill.has_metadata());

        fill.clear_pattern_fill();
        assert!(fill.pattern_fill().is_none());
    }
}
