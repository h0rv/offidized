use crate::color::ShapeColor;
use crate::table::TextDirection;
use crate::text::TextRun;
use offidized_opc::RawXmlNode;

#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum ShapeType {
    #[default]
    AutoShape,
    TextBox,
    Rectangle,
    Ellipse,
    Triangle,
    Diamond,
    Pentagon,
    Hexagon,
    Octagon,
    Star5,
    RightArrow,
}

// ── Feature #1: Shape outline/border ──

/// Dash style for shape outlines, mapped to `a:ln prstDash` attribute values.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum LineDashStyle {
    Solid,
    Dot,
    Dash,
    LargeDash,
    DashDot,
    LargeDashDot,
    LargeDashDotDot,
    SystemDash,
    SystemDot,
    SystemDashDot,
    SystemDashDotDot,
}

impl LineDashStyle {
    pub fn from_xml(value: &str) -> Option<Self> {
        match value {
            "solid" => Some(Self::Solid),
            "dot" => Some(Self::Dot),
            "dash" => Some(Self::Dash),
            "lgDash" => Some(Self::LargeDash),
            "dashDot" => Some(Self::DashDot),
            "lgDashDot" => Some(Self::LargeDashDot),
            "lgDashDotDot" => Some(Self::LargeDashDotDot),
            "sysDash" => Some(Self::SystemDash),
            "sysDot" => Some(Self::SystemDot),
            "sysDashDot" => Some(Self::SystemDashDot),
            "sysDashDotDot" => Some(Self::SystemDashDotDot),
            _ => None,
        }
    }

    pub fn to_xml(self) -> &'static str {
        match self {
            Self::Solid => "solid",
            Self::Dot => "dot",
            Self::Dash => "dash",
            Self::LargeDash => "lgDash",
            Self::DashDot => "dashDot",
            Self::LargeDashDot => "lgDashDot",
            Self::LargeDashDotDot => "lgDashDotDot",
            Self::SystemDash => "sysDash",
            Self::SystemDot => "sysDot",
            Self::SystemDashDot => "sysDashDot",
            Self::SystemDashDotDot => "sysDashDotDot",
        }
    }
}

/// Compound line style for shape outlines.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum LineCompoundStyle {
    Single,
    Double,
    ThickThin,
    ThinThick,
    Triple,
}

impl LineCompoundStyle {
    pub fn from_xml(value: &str) -> Option<Self> {
        match value {
            "sng" => Some(Self::Single),
            "dbl" => Some(Self::Double),
            "thickThin" => Some(Self::ThickThin),
            "thinThick" => Some(Self::ThinThick),
            "tri" => Some(Self::Triple),
            _ => None,
        }
    }

    pub fn to_xml(self) -> &'static str {
        match self {
            Self::Single => "sng",
            Self::Double => "dbl",
            Self::ThickThin => "thickThin",
            Self::ThinThick => "thinThick",
            Self::Triple => "tri",
        }
    }
}

// ── Line Arrows ──

/// Arrow head/tail type for lines and connectors.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ArrowType {
    None,
    Triangle,
    Stealth,
    Diamond,
    Oval,
    Arrow,
    Open,
}

impl ArrowType {
    pub fn from_xml(value: &str) -> Option<Self> {
        match value {
            "none" => Some(Self::None),
            "triangle" => Some(Self::Triangle),
            "stealth" => Some(Self::Stealth),
            "diamond" => Some(Self::Diamond),
            "oval" => Some(Self::Oval),
            "arrow" => Some(Self::Arrow),
            "open" => Some(Self::Open),
            _ => None,
        }
    }

    pub fn to_xml(self) -> &'static str {
        match self {
            Self::None => "none",
            Self::Triangle => "triangle",
            Self::Stealth => "stealth",
            Self::Diamond => "diamond",
            Self::Oval => "oval",
            Self::Arrow => "arrow",
            Self::Open => "open",
        }
    }
}

/// Arrow size (small, medium, large) for width and length.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ArrowSize {
    Small,
    Medium,
    Large,
}

impl ArrowSize {
    pub fn from_xml(value: &str) -> Option<Self> {
        match value {
            "sm" => Some(Self::Small),
            "med" => Some(Self::Medium),
            "lg" => Some(Self::Large),
            _ => None,
        }
    }

    pub fn to_xml(self) -> &'static str {
        match self {
            Self::Small => "sm",
            Self::Medium => "med",
            Self::Large => "lg",
        }
    }
}

/// Arrow properties for a line head or tail end.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct LineArrow {
    /// Arrow type.
    pub arrow_type: ArrowType,
    /// Arrow width size.
    pub width: ArrowSize,
    /// Arrow length size.
    pub length: ArrowSize,
}

impl LineArrow {
    pub fn new(arrow_type: ArrowType) -> Self {
        Self {
            arrow_type,
            width: ArrowSize::Medium,
            length: ArrowSize::Medium,
        }
    }
}

/// Shape outline properties parsed from `<a:ln>` inside `<p:spPr>`.
#[derive(Debug, Clone, PartialEq, Eq, Default)]
pub struct ShapeOutline {
    /// Line width in EMUs (`w` attribute). 12700 EMU = 1pt.
    pub width_emu: Option<i64>,
    /// Solid fill color as sRGB hex (`<a:solidFill><a:srgbClr val="..."/>`).
    pub color_srgb: Option<String>,
    /// Dash style (`<a:prstDash val="..."/>`).
    pub dash_style: Option<LineDashStyle>,
    /// Compound line style (`cmpd` attribute).
    pub compound_style: Option<LineCompoundStyle>,
    /// Head arrow properties (`<a:headEnd>`).
    pub head_arrow: Option<LineArrow>,
    /// Tail arrow properties (`<a:tailEnd>`).
    pub tail_arrow: Option<LineArrow>,
    /// Line color alpha/opacity as percentage (0-100, where 100 = fully opaque).
    /// Parsed from `<a:srgbClr><a:alpha val="..."/></a:srgbClr>`.
    pub alpha: Option<u8>,
    /// Full color model (supports sRGB, scheme colors, and transforms).
    /// When set, this takes precedence over `color_srgb` and `alpha` during serialization.
    pub color: Option<ShapeColor>,
}

impl ShapeOutline {
    pub fn new() -> Self {
        Self::default()
    }

    /// Returns true if any outline property is set.
    pub fn is_set(&self) -> bool {
        self.width_emu.is_some()
            || self.color_srgb.is_some()
            || self.color.is_some()
            || self.dash_style.is_some()
            || self.compound_style.is_some()
            || self.head_arrow.is_some()
            || self.tail_arrow.is_some()
            || self.alpha.is_some()
    }
}

// ── Feature #2: Shape fill improvements ──

/// Gradient fill type.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum GradientFillType {
    Linear,
    Radial,
    Rectangular,
    Path,
}

impl GradientFillType {
    pub fn from_xml(value: &str) -> Option<Self> {
        match value {
            "linear" | "lin" => Some(Self::Linear),
            "radial" => Some(Self::Radial),
            "rect" => Some(Self::Rectangular),
            "path" => Some(Self::Path),
            _ => None,
        }
    }

    pub fn to_xml(self) -> &'static str {
        match self {
            Self::Linear => "lin",
            Self::Radial => "path",
            Self::Rectangular => "rect",
            Self::Path => "path",
        }
    }
}

/// A single gradient stop with color and position.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct GradientStop {
    /// Position in thousandths of a percent (0-100000). 0 = start, 100000 = end.
    pub position: u32,
    /// Color as sRGB hex.
    pub color_srgb: String,
    /// Full color model (supports sRGB, scheme colors, and transforms).
    /// When set, this takes precedence over `color_srgb` during serialization.
    pub color: Option<ShapeColor>,
}

/// Gradient fill properties parsed from `<a:gradFill>` inside `<p:spPr>`.
#[derive(Debug, Clone, PartialEq, Eq, Default)]
pub struct GradientFill {
    /// Gradient type (linear, radial, etc.).
    pub fill_type: Option<GradientFillType>,
    /// Linear gradient angle in 60000ths of a degree.
    pub linear_angle: Option<i32>,
    /// Gradient stops.
    pub stops: Vec<GradientStop>,
}

impl GradientFill {
    pub fn new() -> Self {
        Self::default()
    }

    pub fn is_set(&self) -> bool {
        !self.stops.is_empty()
    }
}

/// Pattern fill type.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum PatternFillType {
    Percent5,
    Percent10,
    Percent20,
    Percent25,
    Percent30,
    Percent40,
    Percent50,
    Percent60,
    Percent70,
    Percent75,
    Percent80,
    Percent90,
    Horizontal,
    Vertical,
    LightHorizontal,
    LightVertical,
    DarkHorizontal,
    DarkVertical,
    NarrowHorizontal,
    NarrowVertical,
    DashedHorizontal,
    DashedVertical,
    Cross,
    DownwardDiagonal,
    UpwardDiagonal,
    LightDownwardDiagonal,
    LightUpwardDiagonal,
    DarkDownwardDiagonal,
    DarkUpwardDiagonal,
    WideDownwardDiagonal,
    WideUpwardDiagonal,
    DashedDownwardDiagonal,
    DashedUpwardDiagonal,
    DiagonalCross,
    SmallCheckerBoard,
    LargeCheckerBoard,
    SmallGrid,
    LargeGrid,
    DottedGrid,
    SmallConfetti,
    LargeConfetti,
    HorizontalBrick,
    DiagonalBrick,
    SolidDiamond,
    OpenDiamond,
    DottedDiamond,
    Plaid,
    Sphere,
    Weave,
    Divot,
    Shingle,
    Wave,
    Trellis,
    ZigZag,
    Other(String),
}

impl PatternFillType {
    pub fn from_xml(value: &str) -> Self {
        match value {
            "pct5" => Self::Percent5,
            "pct10" => Self::Percent10,
            "pct20" => Self::Percent20,
            "pct25" => Self::Percent25,
            "pct30" => Self::Percent30,
            "pct40" => Self::Percent40,
            "pct50" => Self::Percent50,
            "pct60" => Self::Percent60,
            "pct70" => Self::Percent70,
            "pct75" => Self::Percent75,
            "pct80" => Self::Percent80,
            "pct90" => Self::Percent90,
            "horz" => Self::Horizontal,
            "vert" => Self::Vertical,
            "ltHorz" => Self::LightHorizontal,
            "ltVert" => Self::LightVertical,
            "dkHorz" => Self::DarkHorizontal,
            "dkVert" => Self::DarkVertical,
            "narHorz" => Self::NarrowHorizontal,
            "narVert" => Self::NarrowVertical,
            "dashHorz" => Self::DashedHorizontal,
            "dashVert" => Self::DashedVertical,
            "cross" => Self::Cross,
            "dnDiag" => Self::DownwardDiagonal,
            "upDiag" => Self::UpwardDiagonal,
            "ltDnDiag" => Self::LightDownwardDiagonal,
            "ltUpDiag" => Self::LightUpwardDiagonal,
            "dkDnDiag" => Self::DarkDownwardDiagonal,
            "dkUpDiag" => Self::DarkUpwardDiagonal,
            "wdDnDiag" => Self::WideDownwardDiagonal,
            "wdUpDiag" => Self::WideUpwardDiagonal,
            "dashDnDiag" => Self::DashedDownwardDiagonal,
            "dashUpDiag" => Self::DashedUpwardDiagonal,
            "diagCross" => Self::DiagonalCross,
            "smCheck" => Self::SmallCheckerBoard,
            "lgCheck" => Self::LargeCheckerBoard,
            "smGrid" => Self::SmallGrid,
            "lgGrid" => Self::LargeGrid,
            "dotGrid" => Self::DottedGrid,
            "smConfetti" => Self::SmallConfetti,
            "lgConfetti" => Self::LargeConfetti,
            "horzBrick" => Self::HorizontalBrick,
            "diagBrick" => Self::DiagonalBrick,
            "solidDmnd" => Self::SolidDiamond,
            "openDmnd" => Self::OpenDiamond,
            "dotDmnd" => Self::DottedDiamond,
            "plaid" => Self::Plaid,
            "sphere" => Self::Sphere,
            "weave" => Self::Weave,
            "divot" => Self::Divot,
            "shingle" => Self::Shingle,
            "wave" => Self::Wave,
            "trellis" => Self::Trellis,
            "zigZag" => Self::ZigZag,
            other => Self::Other(other.to_string()),
        }
    }

    pub fn to_xml(&self) -> &str {
        match self {
            Self::Percent5 => "pct5",
            Self::Percent10 => "pct10",
            Self::Percent20 => "pct20",
            Self::Percent25 => "pct25",
            Self::Percent30 => "pct30",
            Self::Percent40 => "pct40",
            Self::Percent50 => "pct50",
            Self::Percent60 => "pct60",
            Self::Percent70 => "pct70",
            Self::Percent75 => "pct75",
            Self::Percent80 => "pct80",
            Self::Percent90 => "pct90",
            Self::Horizontal => "horz",
            Self::Vertical => "vert",
            Self::LightHorizontal => "ltHorz",
            Self::LightVertical => "ltVert",
            Self::DarkHorizontal => "dkHorz",
            Self::DarkVertical => "dkVert",
            Self::NarrowHorizontal => "narHorz",
            Self::NarrowVertical => "narVert",
            Self::DashedHorizontal => "dashHorz",
            Self::DashedVertical => "dashVert",
            Self::Cross => "cross",
            Self::DownwardDiagonal => "dnDiag",
            Self::UpwardDiagonal => "upDiag",
            Self::LightDownwardDiagonal => "ltDnDiag",
            Self::LightUpwardDiagonal => "ltUpDiag",
            Self::DarkDownwardDiagonal => "dkDnDiag",
            Self::DarkUpwardDiagonal => "dkUpDiag",
            Self::WideDownwardDiagonal => "wdDnDiag",
            Self::WideUpwardDiagonal => "wdUpDiag",
            Self::DashedDownwardDiagonal => "dashDnDiag",
            Self::DashedUpwardDiagonal => "dashUpDiag",
            Self::DiagonalCross => "diagCross",
            Self::SmallCheckerBoard => "smCheck",
            Self::LargeCheckerBoard => "lgCheck",
            Self::SmallGrid => "smGrid",
            Self::LargeGrid => "lgGrid",
            Self::DottedGrid => "dotGrid",
            Self::SmallConfetti => "smConfetti",
            Self::LargeConfetti => "lgConfetti",
            Self::HorizontalBrick => "horzBrick",
            Self::DiagonalBrick => "diagBrick",
            Self::SolidDiamond => "solidDmnd",
            Self::OpenDiamond => "openDmnd",
            Self::DottedDiamond => "dotDmnd",
            Self::Plaid => "plaid",
            Self::Sphere => "sphere",
            Self::Weave => "weave",
            Self::Divot => "divot",
            Self::Shingle => "shingle",
            Self::Wave => "wave",
            Self::Trellis => "trellis",
            Self::ZigZag => "zigZag",
            Self::Other(ref name) => name.as_str(),
        }
    }
}

/// Pattern fill properties parsed from `<a:pattFill>` inside `<p:spPr>`.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct PatternFill {
    /// Pattern preset type (`prst` attribute).
    pub pattern_type: PatternFillType,
    /// Foreground color as sRGB hex.
    pub foreground_srgb: Option<String>,
    /// Background color as sRGB hex.
    pub background_srgb: Option<String>,
    /// Foreground color (full model, supports scheme colors and transforms).
    pub foreground_color: Option<ShapeColor>,
    /// Background color (full model, supports scheme colors and transforms).
    pub background_color: Option<ShapeColor>,
}

impl PatternFill {
    /// Create a new pattern fill with the given pattern type.
    pub fn new(pattern_type: PatternFillType) -> Self {
        Self {
            pattern_type,
            foreground_srgb: None,
            background_srgb: None,
            foreground_color: None,
            background_color: None,
        }
    }
}

/// Picture fill properties parsed from `<a:blipFill>` inside `<p:spPr>`.
#[derive(Debug, Clone, PartialEq)]
pub struct PictureFill {
    /// Relationship ID referencing the image part.
    pub relationship_id: String,
    /// Whether the image is stretched to fill the shape.
    pub stretch: bool,
    /// Image crop settings (from `<a:srcRect>`).
    pub crop: Option<crate::image::ImageCrop>,
}

impl PictureFill {
    pub fn new(relationship_id: impl Into<String>) -> Self {
        Self {
            relationship_id: relationship_id.into(),
            stretch: true,
            crop: None,
        }
    }

    /// Sets crop boundaries for the image.
    ///
    /// # Arguments
    /// * `crop` - Crop settings, or None to remove cropping
    pub fn set_crop(&mut self, crop: Option<crate::image::ImageCrop>) {
        self.crop = crop;
    }

    /// Gets the current crop settings.
    pub fn crop(&self) -> Option<&crate::image::ImageCrop> {
        self.crop.as_ref()
    }
}

/// Unified shape fill, representing the mutually exclusive fill options.
#[derive(Debug, Clone, PartialEq)]
pub enum ShapeFill {
    /// Solid fill with sRGB hex color.
    Solid(String),
    /// Gradient fill.
    Gradient(GradientFill),
    /// Pattern fill.
    Pattern(PatternFill),
    /// Picture fill (blipFill).
    Picture(PictureFill),
    /// No fill (`<a:noFill/>`).
    NoFill,
}

// ── Shape effects (shadow, glow, reflection) ──

/// Outer shadow effect parsed from `<a:outerShdw>` inside `<a:effectLst>`.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ShapeShadow {
    /// Horizontal offset in EMUs.
    pub offset_x: i64,
    /// Vertical offset in EMUs.
    pub offset_y: i64,
    /// Blur radius in EMUs.
    pub blur_radius: i64,
    /// Shadow color as sRGB hex.
    pub color: String,
    /// Alpha (opacity) as 0-100 percentage.
    pub alpha: Option<u8>,
    /// Full color model (supports sRGB, scheme colors, and transforms).
    /// When set, this takes precedence over `color` and `alpha` during serialization.
    pub color_full: Option<ShapeColor>,
}

impl ShapeShadow {
    /// Create a new outer shadow with the given offsets, blur radius, and sRGB color.
    pub fn new(offset_x: i64, offset_y: i64, blur_radius: i64, color: impl Into<String>) -> Self {
        Self {
            offset_x,
            offset_y,
            blur_radius,
            color: color.into(),
            alpha: None,
            color_full: None,
        }
    }
}

/// Glow effect parsed from `<a:glow>` inside `<a:effectLst>`.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ShapeGlow {
    /// Glow radius in EMUs.
    pub radius: i64,
    /// Glow color as sRGB hex.
    pub color: String,
    /// Alpha (opacity) as 0-100 percentage.
    pub alpha: Option<u8>,
    /// Full color model (supports sRGB, scheme colors, and transforms).
    /// When set, this takes precedence over `color` and `alpha` during serialization.
    pub color_full: Option<ShapeColor>,
}

impl ShapeGlow {
    /// Create a new glow effect with the given radius and sRGB color.
    pub fn new(radius: i64, color: impl Into<String>) -> Self {
        Self {
            radius,
            color: color.into(),
            alpha: None,
            color_full: None,
        }
    }
}

/// Reflection effect parsed from `<a:reflection>` inside `<a:effectLst>`.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ShapeReflection {
    /// Blur radius in EMUs.
    pub blur_radius: i64,
    /// Start opacity as 0-100 percentage (`stA` attribute, in thousandths of a percent in XML).
    pub start_alpha: Option<u8>,
    /// End opacity as 0-100 percentage (`endA` attribute, in thousandths of a percent in XML).
    pub end_alpha: Option<u8>,
    /// Distance from shape in EMUs.
    pub distance: i64,
    /// Direction in 60000ths of a degree.
    pub direction: Option<i64>,
}

impl ShapeReflection {
    pub fn new(blur_radius: i64, distance: i64) -> Self {
        Self {
            blur_radius,
            start_alpha: None,
            end_alpha: None,
            distance,
            direction: None,
        }
    }
}

// ── Text anchoring and auto-fit ──

/// Vertical text anchor position within the shape body, parsed from the
/// `anchor` attribute on `<a:bodyPr>`.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum TextAnchor {
    Top,
    Middle,
    Bottom,
    TopCentered,
    MiddleCentered,
    BottomCentered,
}

impl TextAnchor {
    /// Parse from the `anchor` XML attribute value plus the `anchorCtr` flag.
    pub fn from_xml(anchor: &str, anchor_ctr: bool) -> Option<Self> {
        match (anchor, anchor_ctr) {
            ("t", false) => Some(Self::Top),
            ("t", true) => Some(Self::TopCentered),
            ("ctr", false) => Some(Self::Middle),
            ("ctr", true) => Some(Self::MiddleCentered),
            ("b", false) => Some(Self::Bottom),
            ("b", true) => Some(Self::BottomCentered),
            _ => None,
        }
    }

    /// Convert to the XML `anchor` attribute value.
    pub fn to_xml_anchor(self) -> &'static str {
        match self {
            Self::Top | Self::TopCentered => "t",
            Self::Middle | Self::MiddleCentered => "ctr",
            Self::Bottom | Self::BottomCentered => "b",
        }
    }

    /// Whether `anchorCtr` should be `"1"` for this anchor type.
    pub fn is_centered(self) -> bool {
        matches!(
            self,
            Self::TopCentered | Self::MiddleCentered | Self::BottomCentered
        )
    }
}

/// Auto-fit mode for text within the shape body, parsed from child elements
/// of `<a:bodyPr>`.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum AutoFitType {
    /// No auto-fit (`<a:noAutofit/>`).
    None,
    /// Normal auto-fit, text shrinks to fit (`<a:normAutofit/>`).
    Normal,
    /// Shape auto-sizes to fit text (`<a:spAutoFit/>`).
    ShrinkOnOverflow,
}

impl AutoFitType {
    pub fn from_xml_tag(local_name: &[u8]) -> Option<Self> {
        match local_name {
            b"noAutofit" => Some(Self::None),
            b"normAutofit" => Some(Self::Normal),
            b"spAutoFit" => Some(Self::ShrinkOnOverflow),
            _ => None,
        }
    }

    pub fn to_xml_tag(self) -> &'static str {
        match self {
            Self::None => "a:noAutofit",
            Self::Normal => "a:normAutofit",
            Self::ShrinkOnOverflow => "a:spAutoFit",
        }
    }
}

// ── Feature #11: Bullets and numbering ──

/// Bullet style for a paragraph.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum BulletStyle {
    /// No bullets (`<a:buNone/>`).
    None,
    /// Character bullet (`<a:buChar char="..."/>`).
    Char(String),
    /// Auto-numbered bullet (`<a:buAutoNum type="..."/>`).
    AutoNum(String),
}

/// Bullet properties for a paragraph.
#[derive(Debug, Clone, PartialEq, Eq, Default)]
pub struct BulletProperties {
    /// The bullet style.
    pub style: Option<BulletStyle>,
    /// Bullet font name (`<a:buFont typeface="..."/>`).
    pub font_name: Option<String>,
    /// Bullet size as percentage of text size in thousandths of a percent.
    /// (`<a:buSzPct val="..."/>`).
    pub size_percent: Option<u32>,
    /// Bullet color as sRGB hex (`<a:buClr><a:srgbClr val="..."/></a:buClr>`).
    pub color_srgb: Option<String>,
    /// Bullet color (full model, supports scheme colors and transforms).
    /// When set, this takes precedence over `color_srgb` during serialization.
    pub color: Option<ShapeColor>,
}

/// Horizontal text alignment within a paragraph.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum TextAlignment {
    Left,
    Center,
    Right,
    Justified,
    Distributed,
}

impl TextAlignment {
    pub fn from_xml(value: &str) -> Option<Self> {
        match value {
            "l" => Some(Self::Left),
            "ctr" => Some(Self::Center),
            "r" => Some(Self::Right),
            "just" => Some(Self::Justified),
            "dist" => Some(Self::Distributed),
            _ => None,
        }
    }

    pub fn to_xml(self) -> &'static str {
        match self {
            Self::Left => "l",
            Self::Center => "ctr",
            Self::Right => "r",
            Self::Justified => "just",
            Self::Distributed => "dist",
        }
    }
}

// ── Text line spacing types ──

/// Unit for line spacing values within a paragraph.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum LineSpacingUnit {
    /// Percentage of the font size (thousandths of a percent in OOXML).
    /// 100000 = single spacing, 150000 = 1.5 spacing.
    Percent,
    /// Absolute points (hundredths of a point in OOXML).
    Points,
}

/// Line spacing for a paragraph (`<a:lnSpc>`).
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct LineSpacing {
    /// The spacing value. Interpretation depends on `unit`:
    /// - `Percent`: thousandths of a percent (100000 = single spacing)
    /// - `Points`: hundredths of a point (1200 = 12pt)
    pub value: i32,
    /// Whether the value is a percentage or an absolute point size.
    pub unit: LineSpacingUnit,
}

impl LineSpacing {
    /// Create a line spacing in percentage units.
    pub fn percent(value: i32) -> Self {
        Self {
            value,
            unit: LineSpacingUnit::Percent,
        }
    }

    /// Create a line spacing in point units.
    pub fn points(value: i32) -> Self {
        Self {
            value,
            unit: LineSpacingUnit::Points,
        }
    }
}

/// Unit for space-before / space-after values.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum SpacingUnit {
    /// Percentage of the font size (thousandths of a percent in OOXML).
    Percent,
    /// Absolute points (hundredths of a point in OOXML).
    Points,
}

/// Spacing value for space-before or space-after (`<a:spcBef>`, `<a:spcAft>`).
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct SpacingValue {
    /// The spacing value. Interpretation depends on `unit`.
    pub value: i32,
    /// Whether the value is a percentage or an absolute point size.
    pub unit: SpacingUnit,
}

impl SpacingValue {
    /// Create a spacing value in percentage units.
    pub fn percent(value: i32) -> Self {
        Self {
            value,
            unit: SpacingUnit::Percent,
        }
    }

    /// Create a spacing value in point units.
    pub fn points(value: i32) -> Self {
        Self {
            value,
            unit: SpacingUnit::Points,
        }
    }
}

/// Paragraph-level properties parsed from `<a:pPr>`.
#[derive(Debug, Clone, PartialEq, Eq, Default)]
pub struct ParagraphProperties {
    /// Horizontal alignment (`algn` attribute).
    pub alignment: Option<TextAlignment>,
    /// Indentation level 0-8 (`lvl` attribute).
    pub level: Option<u32>,
    /// Left margin in EMUs (`marL` attribute).
    pub margin_left_emu: Option<i64>,
    /// Right margin in EMUs (`marR` attribute).
    pub margin_right_emu: Option<i64>,
    /// First-line indent in EMUs (`indent` attribute). Negative = hanging.
    pub indent_emu: Option<i64>,
    /// Line spacing in hundredths of a percent (`<a:lnSpc><a:spcPct val="150000"/>`).
    /// 100000 = single spacing, 150000 = 1.5 spacing.
    pub line_spacing_pct: Option<u32>,
    /// Line spacing in hundredths of a point (`<a:lnSpc><a:spcPts val="1200"/>`).
    pub line_spacing_pts: Option<u32>,
    /// Space before paragraph in hundredths of a point.
    pub space_before_pts: Option<u32>,
    /// Space after paragraph in hundredths of a point.
    pub space_after_pts: Option<u32>,
    /// Typed line spacing (`<a:lnSpc>`), combining value and unit.
    pub line_spacing: Option<LineSpacing>,
    /// Typed space before paragraph (`<a:spcBef>`), combining value and unit.
    pub space_before: Option<SpacingValue>,
    /// Typed space after paragraph (`<a:spcAft>`), combining value and unit.
    pub space_after: Option<SpacingValue>,
    /// Bullet/numbering properties.
    pub bullet: BulletProperties,
}

#[derive(Debug, Clone, PartialEq, Eq, Default)]
pub struct ShapeParagraph {
    runs: Vec<TextRun>,
    properties: ParagraphProperties,
}

impl ShapeParagraph {
    pub fn new() -> Self {
        Self::default()
    }

    pub fn add_run(&mut self, text: impl Into<String>) -> &mut TextRun {
        self.runs.push(TextRun::new(text));
        let index = self.runs.len().saturating_sub(1);
        &mut self.runs[index]
    }

    pub fn runs(&self) -> &[TextRun] {
        &self.runs
    }

    pub fn runs_mut(&mut self) -> &mut [TextRun] {
        &mut self.runs
    }

    pub fn run_count(&self) -> usize {
        self.runs.len()
    }

    /// Access the paragraph properties.
    pub fn properties(&self) -> &ParagraphProperties {
        &self.properties
    }

    /// Mutable access to the paragraph properties.
    pub fn properties_mut(&mut self) -> &mut ParagraphProperties {
        &mut self.properties
    }

    pub(crate) fn set_properties(&mut self, properties: ParagraphProperties) {
        self.properties = properties;
    }

    // Convenience getters/setters.

    pub fn alignment(&self) -> Option<TextAlignment> {
        self.properties.alignment
    }

    pub fn set_alignment(&mut self, alignment: TextAlignment) -> &mut Self {
        self.properties.alignment = Some(alignment);
        self
    }

    pub fn clear_alignment(&mut self) -> &mut Self {
        self.properties.alignment = None;
        self
    }

    pub fn level(&self) -> Option<u32> {
        self.properties.level
    }

    pub fn set_level(&mut self, level: u32) -> &mut Self {
        self.properties.level = Some(level);
        self
    }

    pub fn clear_level(&mut self) -> &mut Self {
        self.properties.level = None;
        self
    }
}

// ── Feature: Placeholder type enum ──

/// Typed placeholder type replacing string-based `placeholder_kind`.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum PlaceholderType {
    Title,
    Body,
    CenteredTitle,
    Subtitle,
    DateAndTime,
    SlideNumber,
    Footer,
    Header,
    Object,
    Chart,
    Table,
    ClipArt,
    Diagram,
    Media,
    SlideImage,
    Other(String),
}

impl PlaceholderType {
    /// Parse from the `type` attribute value on `<p:ph>`.
    pub fn from_xml(value: &str) -> Self {
        match value {
            "title" => Self::Title,
            "body" => Self::Body,
            "ctrTitle" => Self::CenteredTitle,
            "subTitle" => Self::Subtitle,
            "dt" => Self::DateAndTime,
            "sldNum" => Self::SlideNumber,
            "ftr" => Self::Footer,
            "hdr" => Self::Header,
            "obj" => Self::Object,
            "chart" => Self::Chart,
            "tbl" => Self::Table,
            "clipArt" => Self::ClipArt,
            "dgm" => Self::Diagram,
            "media" => Self::Media,
            "sldImg" => Self::SlideImage,
            other => Self::Other(other.to_string()),
        }
    }

    /// Convert to the XML attribute value.
    pub fn to_xml(&self) -> &str {
        match self {
            Self::Title => "title",
            Self::Body => "body",
            Self::CenteredTitle => "ctrTitle",
            Self::Subtitle => "subTitle",
            Self::DateAndTime => "dt",
            Self::SlideNumber => "sldNum",
            Self::Footer => "ftr",
            Self::Header => "hdr",
            Self::Object => "obj",
            Self::Chart => "chart",
            Self::Table => "tbl",
            Self::ClipArt => "clipArt",
            Self::Diagram => "dgm",
            Self::Media => "media",
            Self::SlideImage => "sldImg",
            Self::Other(ref name) => name.as_str(),
        }
    }
}

// ── Feature: Audio/video media ──

/// Media type for audio/video shapes.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum MediaType {
    Audio,
    Video,
}

// ── Feature: Connector shapes ──

/// Connection endpoint information for connector shapes.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ConnectionInfo {
    /// The shape ID this connector attaches to.
    pub shape_id: u32,
    /// The connection point index on the target shape.
    pub connection_point_index: u32,
}

impl ConnectionInfo {
    pub fn new(shape_id: u32, connection_point_index: u32) -> Self {
        Self {
            shape_id,
            connection_point_index,
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub struct ShapeGeometry {
    x: i64,
    y: i64,
    cx: i64,
    cy: i64,
}

impl ShapeGeometry {
    pub fn new(x: i64, y: i64, cx: i64, cy: i64) -> Self {
        Self { x, y, cx, cy }
    }

    pub fn x(&self) -> i64 {
        self.x
    }

    pub fn y(&self) -> i64 {
        self.y
    }

    pub fn cx(&self) -> i64 {
        self.cx
    }

    pub fn cy(&self) -> i64 {
        self.cy
    }
}

#[derive(Debug, Clone, PartialEq)]
pub struct Shape {
    name: String,
    paragraphs: Vec<ShapeParagraph>,
    placeholder_kind: Option<String>,
    placeholder_idx: Option<u32>,
    /// Typed placeholder type (replaces string-based placeholder_kind for typed access).
    placeholder_type: Option<PlaceholderType>,
    shape_type: ShapeType,
    geometry: Option<ShapeGeometry>,
    preset_geometry: Option<String>,
    /// Custom geometry as a raw XML node (`a:custGeom`), preserved for roundtrip.
    custom_geometry_raw: Option<RawXmlNode>,
    /// Preset geometry adjustment values (`a:avLst` children), preserved for roundtrip.
    preset_geometry_adjustments: Option<RawXmlNode>,
    solid_fill_srgb: Option<String>,
    /// Solid fill alpha/opacity as percentage (0-100, where 100 = fully opaque).
    solid_fill_alpha: Option<u8>,
    /// Full solid fill color model (supports sRGB, scheme colors, and transforms).
    /// When set, this takes precedence over `solid_fill_srgb`/`solid_fill_alpha` during serialization.
    solid_fill_color: Option<ShapeColor>,
    /// Shape outline/border properties (Feature #1).
    outline: Option<ShapeOutline>,
    /// Gradient fill (Feature #2).
    gradient_fill: Option<GradientFill>,
    /// Pattern fill (Feature #2).
    pattern_fill: Option<PatternFill>,
    /// Picture fill (blipFill).
    picture_fill: Option<PictureFill>,
    /// No fill flag (Feature #2).
    no_fill: bool,
    /// Rotation in 60000ths of a degree (Feature #3).
    rotation: Option<i32>,
    /// Horizontal flip flag from `flipH` attribute on `<a:xfrm>`.
    flip_h: bool,
    /// Vertical flip flag from `flipV` attribute on `<a:xfrm>`.
    flip_v: bool,
    /// Hidden flag from p:cNvPr (Feature #4).
    hidden: bool,
    /// Hyperlink click target relationship id for the shape (Feature #10, basic).
    hyperlink_click_rid: Option<String>,
    /// Alt text description from `p:cNvPr descr` attribute.
    alt_text: Option<String>,
    /// Alt text title from `p:cNvPr title` attribute.
    alt_text_title: Option<String>,
    /// Whether this shape is a SmartArt (contains `dgm:relIds`).
    is_smartart: bool,
    /// Whether this shape is a connector (`p:cxnSp`).
    is_connector: bool,
    /// Start connection for connector shapes.
    start_connection: Option<ConnectionInfo>,
    /// End connection for connector shapes.
    end_connection: Option<ConnectionInfo>,
    /// Media type and relationship ID for audio/video shapes.
    media: Option<(MediaType, String)>,
    /// Outer shadow effect.
    shadow: Option<ShapeShadow>,
    /// Glow effect.
    glow: Option<ShapeGlow>,
    /// Reflection effect.
    reflection: Option<ShapeReflection>,
    /// Vertical text anchor position.
    text_anchor: Option<TextAnchor>,
    /// Auto-fit mode for text body.
    auto_fit: Option<AutoFitType>,
    /// Text direction (`vert` attribute on `<a:bodyPr>`).
    text_direction: Option<TextDirection>,
    /// Number of text columns (`numCol` attribute on `<a:bodyPr>`).
    text_columns: Option<u32>,
    /// Spacing between text columns in EMUs (`spcCol` attribute on `<a:bodyPr>`).
    text_column_spacing: Option<i64>,
    /// Left text inset in EMUs (`lIns` attribute on `<a:bodyPr>`).
    text_inset_left: Option<i64>,
    /// Right text inset in EMUs (`rIns` attribute on `<a:bodyPr>`).
    text_inset_right: Option<i64>,
    /// Top text inset in EMUs (`tIns` attribute on `<a:bodyPr>`).
    text_inset_top: Option<i64>,
    /// Bottom text inset in EMUs (`bIns` attribute on `<a:bodyPr>`).
    text_inset_bottom: Option<i64>,
    /// Word wrap mode (`wrap` attribute on `<a:bodyPr>`; "square" = true, "none" = false).
    word_wrap: Option<bool>,
    /// Body text rotation in 60000ths of a degree (`rot` attribute on `<a:bodyPr>`).
    body_pr_rot: Option<i32>,
    /// Right-to-left column ordering (`rtlCol` attribute on `<a:bodyPr>`).
    body_pr_rtl_col: Option<bool>,
    /// WordArt text body (`fromWordArt` attribute on `<a:bodyPr>`).
    body_pr_from_word_art: Option<bool>,
    /// Force anti-alias (`forceAA` attribute on `<a:bodyPr>`).
    body_pr_force_aa: Option<bool>,
    /// Compatible line spacing (`compatLnSpc` attribute on `<a:bodyPr>`).
    body_pr_compat_ln_spc: Option<bool>,
    /// Shape action settings (click/hover actions).
    action: Option<crate::actions::ShapeAction>,
    /// Embedded OLE object data.
    embedded_object: Option<crate::actions::EmbeddedObject>,
    /// Action button preset type.
    action_button_type: Option<crate::actions::ActionButtonType>,
    unknown_attrs: Vec<(String, String)>,
    unknown_children: Vec<RawXmlNode>,
    /// Raw `<a:lstStyle>` element from `txBody`, preserved for roundtrip fidelity.
    lst_style_raw: Option<RawXmlNode>,
    /// Unknown children of `<a:bodyPr>`, preserved for roundtrip fidelity.
    body_pr_unknown_children: Vec<RawXmlNode>,
    /// Unknown attributes on `<a:bodyPr>`, preserved for roundtrip fidelity.
    body_pr_unknown_attrs: Vec<(String, String)>,
}

impl Shape {
    pub fn new(name: impl Into<String>) -> Self {
        Self {
            name: name.into(),
            paragraphs: Vec::new(),
            placeholder_kind: None,
            placeholder_idx: None,
            placeholder_type: None,
            shape_type: ShapeType::AutoShape,
            geometry: None,
            preset_geometry: None,
            custom_geometry_raw: None,
            preset_geometry_adjustments: None,
            solid_fill_srgb: None,
            solid_fill_alpha: None,
            solid_fill_color: None,
            outline: None,
            gradient_fill: None,
            pattern_fill: None,
            picture_fill: None,
            no_fill: false,
            rotation: None,
            flip_h: false,
            flip_v: false,
            hidden: false,
            hyperlink_click_rid: None,
            alt_text: None,
            alt_text_title: None,
            is_smartart: false,
            is_connector: false,
            start_connection: None,
            end_connection: None,
            media: None,
            shadow: None,
            glow: None,
            reflection: None,
            text_anchor: None,
            auto_fit: None,
            text_direction: None,
            text_columns: None,
            text_column_spacing: None,
            text_inset_left: None,
            text_inset_right: None,
            text_inset_top: None,
            text_inset_bottom: None,
            word_wrap: None,
            body_pr_rot: None,
            body_pr_rtl_col: None,
            body_pr_from_word_art: None,
            body_pr_force_aa: None,
            body_pr_compat_ln_spc: None,
            action: None,
            embedded_object: None,
            action_button_type: None,
            unknown_attrs: Vec::new(),
            unknown_children: Vec::new(),
            lst_style_raw: None,
            body_pr_unknown_children: Vec::new(),
            body_pr_unknown_attrs: Vec::new(),
        }
    }

    pub fn name(&self) -> &str {
        &self.name
    }

    pub fn add_paragraph(&mut self) -> &mut ShapeParagraph {
        self.paragraphs.push(ShapeParagraph::new());
        let index = self.paragraphs.len().saturating_sub(1);
        &mut self.paragraphs[index]
    }

    pub fn add_paragraph_with_text(&mut self, text: impl Into<String>) -> &mut ShapeParagraph {
        let paragraph = self.add_paragraph();
        paragraph.add_run(text);
        paragraph
    }

    pub fn paragraphs(&self) -> &[ShapeParagraph] {
        &self.paragraphs
    }

    pub fn paragraphs_mut(&mut self) -> &mut [ShapeParagraph] {
        &mut self.paragraphs
    }

    pub fn paragraph_count(&self) -> usize {
        self.paragraphs.len()
    }

    pub fn placeholder_kind(&self) -> Option<&str> {
        self.placeholder_kind.as_deref()
    }

    pub fn set_placeholder_kind(&mut self, placeholder_kind: impl Into<String>) {
        self.placeholder_kind = Some(placeholder_kind.into());
    }

    pub fn clear_placeholder_kind(&mut self) {
        self.placeholder_kind = None;
    }

    pub fn placeholder_idx(&self) -> Option<u32> {
        self.placeholder_idx
    }

    pub fn set_placeholder_idx(&mut self, placeholder_idx: u32) {
        self.placeholder_idx = Some(placeholder_idx);
    }

    pub fn clear_placeholder_idx(&mut self) {
        self.placeholder_idx = None;
    }

    pub fn shape_type(&self) -> ShapeType {
        self.shape_type
    }

    pub fn set_shape_type(&mut self, shape_type: ShapeType) {
        self.shape_type = shape_type;
    }

    pub fn geometry(&self) -> Option<ShapeGeometry> {
        self.geometry
    }

    pub fn set_geometry(&mut self, geometry: ShapeGeometry) {
        self.geometry = Some(geometry);
    }

    pub fn clear_geometry(&mut self) {
        self.geometry = None;
    }

    pub fn preset_geometry(&self) -> Option<&str> {
        self.preset_geometry.as_deref()
    }

    pub fn set_preset_geometry(&mut self, preset_geometry: impl Into<String>) {
        self.preset_geometry = Some(preset_geometry.into());
    }

    pub fn clear_preset_geometry(&mut self) {
        self.preset_geometry = None;
    }

    /// Custom geometry raw XML node (`a:custGeom`), preserved for roundtrip.
    pub(crate) fn custom_geometry_raw(&self) -> Option<&RawXmlNode> {
        self.custom_geometry_raw.as_ref()
    }

    /// Set custom geometry raw XML node.
    pub(crate) fn set_custom_geometry_raw(&mut self, node: RawXmlNode) {
        self.custom_geometry_raw = Some(node);
    }

    /// Preset geometry adjustment values (`a:avLst` from `a:prstGeom`), preserved for roundtrip.
    pub(crate) fn preset_geometry_adjustments(&self) -> Option<&RawXmlNode> {
        self.preset_geometry_adjustments.as_ref()
    }

    /// Set preset geometry adjustment values.
    pub(crate) fn set_preset_geometry_adjustments(&mut self, node: RawXmlNode) {
        self.preset_geometry_adjustments = Some(node);
    }

    pub fn solid_fill_srgb(&self) -> Option<&str> {
        self.solid_fill_srgb.as_deref()
    }

    pub fn set_solid_fill_srgb(&mut self, solid_fill_srgb: impl Into<String>) {
        self.solid_fill_srgb = Some(solid_fill_srgb.into());
    }

    pub fn clear_solid_fill_srgb(&mut self) {
        self.solid_fill_srgb = None;
    }

    // ── Solid fill alpha/opacity ──

    /// Solid fill alpha/opacity as percentage (0-100, where 100 = fully opaque).
    pub fn solid_fill_alpha(&self) -> Option<u8> {
        self.solid_fill_alpha
    }

    /// Set solid fill alpha/opacity (0-100, where 100 = fully opaque).
    pub fn set_solid_fill_alpha(&mut self, alpha: u8) {
        self.solid_fill_alpha = Some(alpha);
    }

    /// Clear solid fill alpha.
    pub fn clear_solid_fill_alpha(&mut self) {
        self.solid_fill_alpha = None;
    }

    /// Full solid fill color (supports sRGB, scheme colors, and transforms).
    pub fn solid_fill_color(&self) -> Option<&ShapeColor> {
        self.solid_fill_color.as_ref()
    }

    /// Set the full solid fill color model.
    pub fn set_solid_fill_color(&mut self, color: ShapeColor) {
        // Also populate the legacy field for backward compat.
        if let Some(srgb_val) = color.srgb_value() {
            self.solid_fill_srgb = Some(srgb_val.to_string());
        }
        self.solid_fill_alpha = color.alpha();
        self.solid_fill_color = Some(color);
    }

    /// Clear the full solid fill color model.
    pub fn clear_solid_fill_color(&mut self) {
        self.solid_fill_color = None;
    }

    // ── Feature #1: Shape outline/border ──

    /// Shape outline/border properties.
    pub fn outline(&self) -> Option<&ShapeOutline> {
        self.outline.as_ref()
    }

    /// Set the shape outline.
    pub fn set_outline(&mut self, outline: ShapeOutline) {
        self.outline = Some(outline);
    }

    /// Clear the shape outline.
    pub fn clear_outline(&mut self) {
        self.outline = None;
    }

    // ── Feature #2: Shape fill improvements ──

    /// Gradient fill properties.
    pub fn gradient_fill(&self) -> Option<&GradientFill> {
        self.gradient_fill.as_ref()
    }

    /// Set gradient fill.
    pub fn set_gradient_fill(&mut self, fill: GradientFill) {
        self.gradient_fill = Some(fill);
        // Gradient fill is mutually exclusive with solid/pattern/no fill.
        self.solid_fill_srgb = None;
        self.pattern_fill = None;
        self.no_fill = false;
    }

    /// Clear gradient fill.
    pub fn clear_gradient_fill(&mut self) {
        self.gradient_fill = None;
    }

    /// Pattern fill properties.
    pub fn pattern_fill(&self) -> Option<&PatternFill> {
        self.pattern_fill.as_ref()
    }

    /// Set pattern fill.
    pub fn set_pattern_fill(&mut self, fill: PatternFill) {
        self.pattern_fill = Some(fill);
        self.solid_fill_srgb = None;
        self.gradient_fill = None;
        self.no_fill = false;
    }

    /// Clear pattern fill.
    pub fn clear_pattern_fill(&mut self) {
        self.pattern_fill = None;
    }

    /// Whether the shape has no fill (`<a:noFill/>`).
    pub fn is_no_fill(&self) -> bool {
        self.no_fill
    }

    /// Set no fill.
    pub fn set_no_fill(&mut self, no_fill: bool) {
        self.no_fill = no_fill;
        if no_fill {
            self.solid_fill_srgb = None;
            self.gradient_fill = None;
            self.pattern_fill = None;
        }
    }

    /// Get the unified fill representation.
    pub fn fill(&self) -> Option<ShapeFill> {
        if self.no_fill {
            Some(ShapeFill::NoFill)
        } else if let Some(ref color) = self.solid_fill_srgb {
            Some(ShapeFill::Solid(color.clone()))
        } else if let Some(ref gradient) = self.gradient_fill {
            Some(ShapeFill::Gradient(gradient.clone()))
        } else if let Some(ref picture) = self.picture_fill {
            Some(ShapeFill::Picture(picture.clone()))
        } else {
            self.pattern_fill
                .as_ref()
                .map(|pattern| ShapeFill::Pattern(pattern.clone()))
        }
    }

    // ── Picture fill (blipFill) ──

    /// Picture fill properties.
    pub fn picture_fill(&self) -> Option<&PictureFill> {
        self.picture_fill.as_ref()
    }

    /// Set picture fill.
    pub fn set_picture_fill(&mut self, fill: PictureFill) {
        self.picture_fill = Some(fill);
        self.solid_fill_srgb = None;
        self.gradient_fill = None;
        self.pattern_fill = None;
        self.no_fill = false;
    }

    /// Clear picture fill.
    pub fn clear_picture_fill(&mut self) {
        self.picture_fill = None;
    }

    // ── Feature #3: Shape rotation ──

    /// Rotation in 60000ths of a degree. 5400000 = 90 degrees.
    pub fn rotation(&self) -> Option<i32> {
        self.rotation
    }

    /// Set rotation in 60000ths of a degree.
    pub fn set_rotation(&mut self, rotation: i32) {
        self.rotation = Some(rotation);
    }

    /// Clear rotation.
    pub fn clear_rotation(&mut self) {
        self.rotation = None;
    }

    // ── Flip attributes ──

    /// Whether the shape is flipped horizontally (`flipH` on `<a:xfrm>`).
    pub fn flip_h(&self) -> bool {
        self.flip_h
    }

    /// Set horizontal flip.
    pub fn set_flip_h(&mut self, flip: bool) {
        self.flip_h = flip;
    }

    /// Whether the shape is flipped vertically (`flipV` on `<a:xfrm>`).
    pub fn flip_v(&self) -> bool {
        self.flip_v
    }

    /// Set vertical flip.
    pub fn set_flip_v(&mut self, flip: bool) {
        self.flip_v = flip;
    }

    // ── Feature #4: Shape visibility ──

    /// Whether the shape is hidden.
    pub fn is_hidden(&self) -> bool {
        self.hidden
    }

    /// Set hidden flag.
    pub fn set_hidden(&mut self, hidden: bool) {
        self.hidden = hidden;
    }

    // ── Feature #10: Hyperlinks (basic shape-level) ──

    /// Hyperlink click relationship ID.
    pub fn hyperlink_click_rid(&self) -> Option<&str> {
        self.hyperlink_click_rid.as_deref()
    }

    pub fn set_hyperlink_click_rid(&mut self, rid: impl Into<String>) {
        self.hyperlink_click_rid = Some(rid.into());
    }

    pub fn clear_hyperlink_click_rid(&mut self) {
        self.hyperlink_click_rid = None;
    }

    // ── Placeholder type enum ──

    /// Typed placeholder type.
    pub fn placeholder_type(&self) -> Option<&PlaceholderType> {
        self.placeholder_type.as_ref()
    }

    /// Set the typed placeholder type.
    pub fn set_placeholder_type(&mut self, placeholder_type: PlaceholderType) {
        self.placeholder_kind = Some(placeholder_type.to_xml().to_string());
        self.placeholder_type = Some(placeholder_type);
    }

    /// Clear the typed placeholder type.
    pub fn clear_placeholder_type(&mut self) {
        self.placeholder_type = None;
        self.placeholder_kind = None;
    }

    // ── Alt text on shapes ──

    /// Alt text description (`descr` attribute on `p:cNvPr`).
    pub fn alt_text(&self) -> Option<&str> {
        self.alt_text.as_deref()
    }

    /// Set alt text description.
    pub fn set_alt_text(&mut self, text: impl Into<String>) {
        self.alt_text = Some(text.into());
    }

    /// Clear alt text description.
    pub fn clear_alt_text(&mut self) {
        self.alt_text = None;
    }

    /// Alt text title (`title` attribute on `p:cNvPr`).
    pub fn alt_text_title(&self) -> Option<&str> {
        self.alt_text_title.as_deref()
    }

    /// Set alt text title.
    pub fn set_alt_text_title(&mut self, title: impl Into<String>) {
        self.alt_text_title = Some(title.into());
    }

    /// Clear alt text title.
    pub fn clear_alt_text_title(&mut self) {
        self.alt_text_title = None;
    }

    // ── SmartArt detection ──

    /// Whether this shape is a SmartArt graphic.
    pub fn is_smartart(&self) -> bool {
        self.is_smartart
    }

    /// Set SmartArt flag.
    pub fn set_smartart(&mut self, is_smartart: bool) {
        self.is_smartart = is_smartart;
    }

    // ── Connector shapes ──

    /// Whether this shape is a connector.
    pub fn is_connector(&self) -> bool {
        self.is_connector
    }

    /// Set connector flag.
    pub fn set_connector(&mut self, is_connector: bool) {
        self.is_connector = is_connector;
    }

    /// Start connection for connector shapes.
    pub fn start_connection(&self) -> Option<&ConnectionInfo> {
        self.start_connection.as_ref()
    }

    /// Set start connection.
    pub fn set_start_connection(&mut self, connection: ConnectionInfo) {
        self.start_connection = Some(connection);
    }

    /// Clear start connection.
    pub fn clear_start_connection(&mut self) {
        self.start_connection = None;
    }

    /// End connection for connector shapes.
    pub fn end_connection(&self) -> Option<&ConnectionInfo> {
        self.end_connection.as_ref()
    }

    /// Set end connection.
    pub fn set_end_connection(&mut self, connection: ConnectionInfo) {
        self.end_connection = Some(connection);
    }

    /// Clear end connection.
    pub fn clear_end_connection(&mut self) {
        self.end_connection = None;
    }

    // ── Audio/video media ──

    /// Media type and relationship ID for audio/video shapes.
    pub fn media(&self) -> Option<(&MediaType, &str)> {
        self.media.as_ref().map(|(t, rid)| (t, rid.as_str()))
    }

    /// Set media type and relationship ID.
    pub fn set_media(&mut self, media_type: MediaType, relationship_id: impl Into<String>) {
        self.media = Some((media_type, relationship_id.into()));
    }

    /// Clear media.
    pub fn clear_media(&mut self) {
        self.media = None;
    }

    // ── Shape effects ──

    /// Outer shadow effect.
    pub fn shadow(&self) -> Option<&ShapeShadow> {
        self.shadow.as_ref()
    }

    /// Set the outer shadow effect.
    pub fn set_shadow(&mut self, shadow: ShapeShadow) {
        self.shadow = Some(shadow);
    }

    /// Clear the outer shadow effect.
    pub fn clear_shadow(&mut self) {
        self.shadow = None;
    }

    /// Glow effect.
    pub fn glow(&self) -> Option<&ShapeGlow> {
        self.glow.as_ref()
    }

    /// Set the glow effect.
    pub fn set_glow(&mut self, glow: ShapeGlow) {
        self.glow = Some(glow);
    }

    /// Clear the glow effect.
    pub fn clear_glow(&mut self) {
        self.glow = None;
    }

    /// Reflection effect.
    pub fn reflection(&self) -> Option<&ShapeReflection> {
        self.reflection.as_ref()
    }

    /// Set the reflection effect.
    pub fn set_reflection(&mut self, reflection: ShapeReflection) {
        self.reflection = Some(reflection);
    }

    /// Clear the reflection effect.
    pub fn clear_reflection(&mut self) {
        self.reflection = None;
    }

    // ── Text anchoring and auto-fit ──

    /// Vertical text anchor position.
    pub fn text_anchor(&self) -> Option<TextAnchor> {
        self.text_anchor
    }

    /// Set the vertical text anchor.
    pub fn set_text_anchor(&mut self, anchor: TextAnchor) {
        self.text_anchor = Some(anchor);
    }

    /// Clear the text anchor.
    pub fn clear_text_anchor(&mut self) {
        self.text_anchor = None;
    }

    /// Auto-fit mode for text body.
    pub fn auto_fit(&self) -> Option<AutoFitType> {
        self.auto_fit
    }

    /// Set the auto-fit mode.
    pub fn set_auto_fit(&mut self, auto_fit: AutoFitType) {
        self.auto_fit = Some(auto_fit);
    }

    /// Clear the auto-fit mode.
    pub fn clear_auto_fit(&mut self) {
        self.auto_fit = None;
    }

    // ── Text direction ──

    /// Text direction (`vert` attribute on `<a:bodyPr>`).
    pub fn text_direction(&self) -> Option<TextDirection> {
        self.text_direction
    }

    /// Set the text direction.
    pub fn set_text_direction(&mut self, direction: TextDirection) {
        self.text_direction = Some(direction);
    }

    /// Clear the text direction.
    pub fn clear_text_direction(&mut self) {
        self.text_direction = None;
    }

    // ── Text columns ──

    /// Number of text columns (`numCol` attribute on `<a:bodyPr>`).
    pub fn text_columns(&self) -> Option<u32> {
        self.text_columns
    }

    /// Set the number of text columns.
    pub fn set_text_columns(&mut self, columns: u32) {
        self.text_columns = Some(columns);
    }

    /// Clear the number of text columns.
    pub fn clear_text_columns(&mut self) {
        self.text_columns = None;
    }

    /// Spacing between text columns in EMUs (`spcCol` attribute on `<a:bodyPr>`).
    pub fn text_column_spacing(&self) -> Option<i64> {
        self.text_column_spacing
    }

    /// Set the spacing between text columns in EMUs.
    pub fn set_text_column_spacing(&mut self, spacing: i64) {
        self.text_column_spacing = Some(spacing);
    }

    /// Clear the text column spacing.
    pub fn clear_text_column_spacing(&mut self) {
        self.text_column_spacing = None;
    }

    // ── Text insets/margins ──

    /// Left text inset in EMUs (`lIns` attribute on `<a:bodyPr>`).
    pub fn text_inset_left(&self) -> Option<i64> {
        self.text_inset_left
    }

    /// Set the left text inset in EMUs.
    pub fn set_text_inset_left(&mut self, inset: i64) {
        self.text_inset_left = Some(inset);
    }

    /// Clear the left text inset.
    pub fn clear_text_inset_left(&mut self) {
        self.text_inset_left = None;
    }

    /// Right text inset in EMUs (`rIns` attribute on `<a:bodyPr>`).
    pub fn text_inset_right(&self) -> Option<i64> {
        self.text_inset_right
    }

    /// Set the right text inset in EMUs.
    pub fn set_text_inset_right(&mut self, inset: i64) {
        self.text_inset_right = Some(inset);
    }

    /// Clear the right text inset.
    pub fn clear_text_inset_right(&mut self) {
        self.text_inset_right = None;
    }

    /// Top text inset in EMUs (`tIns` attribute on `<a:bodyPr>`).
    pub fn text_inset_top(&self) -> Option<i64> {
        self.text_inset_top
    }

    /// Set the top text inset in EMUs.
    pub fn set_text_inset_top(&mut self, inset: i64) {
        self.text_inset_top = Some(inset);
    }

    /// Clear the top text inset.
    pub fn clear_text_inset_top(&mut self) {
        self.text_inset_top = None;
    }

    /// Bottom text inset in EMUs (`bIns` attribute on `<a:bodyPr>`).
    pub fn text_inset_bottom(&self) -> Option<i64> {
        self.text_inset_bottom
    }

    /// Set the bottom text inset in EMUs.
    pub fn set_text_inset_bottom(&mut self, inset: i64) {
        self.text_inset_bottom = Some(inset);
    }

    /// Clear the bottom text inset.
    pub fn clear_text_inset_bottom(&mut self) {
        self.text_inset_bottom = None;
    }

    // ── Word wrap ──

    /// Whether word wrap is enabled (`wrap` attribute on `<a:bodyPr>`;
    /// "square" = true, "none" = false).
    pub fn word_wrap(&self) -> Option<bool> {
        self.word_wrap
    }

    /// Set word wrap mode.
    pub fn set_word_wrap(&mut self, wrap: bool) {
        self.word_wrap = Some(wrap);
    }

    /// Clear word wrap mode.
    pub fn clear_word_wrap(&mut self) {
        self.word_wrap = None;
    }

    // ── bodyPr extended attributes ──

    /// Body text rotation in 60000ths of a degree (`rot` attribute on `<a:bodyPr>`).
    pub fn body_pr_rot(&self) -> Option<i32> {
        self.body_pr_rot
    }

    /// Set body text rotation.
    pub fn set_body_pr_rot(&mut self, rot: i32) {
        self.body_pr_rot = Some(rot);
    }

    /// Right-to-left column ordering (`rtlCol` attribute on `<a:bodyPr>`).
    pub fn body_pr_rtl_col(&self) -> Option<bool> {
        self.body_pr_rtl_col
    }

    /// Set right-to-left column ordering.
    pub fn set_body_pr_rtl_col(&mut self, rtl: bool) {
        self.body_pr_rtl_col = Some(rtl);
    }

    /// WordArt text body (`fromWordArt` attribute on `<a:bodyPr>`).
    pub fn body_pr_from_word_art(&self) -> Option<bool> {
        self.body_pr_from_word_art
    }

    /// Set WordArt text body flag.
    pub fn set_body_pr_from_word_art(&mut self, from_word_art: bool) {
        self.body_pr_from_word_art = Some(from_word_art);
    }

    /// Force anti-alias (`forceAA` attribute on `<a:bodyPr>`).
    pub fn body_pr_force_aa(&self) -> Option<bool> {
        self.body_pr_force_aa
    }

    /// Set force anti-alias flag.
    pub fn set_body_pr_force_aa(&mut self, force_aa: bool) {
        self.body_pr_force_aa = Some(force_aa);
    }

    /// Compatible line spacing (`compatLnSpc` attribute on `<a:bodyPr>`).
    pub fn body_pr_compat_ln_spc(&self) -> Option<bool> {
        self.body_pr_compat_ln_spc
    }

    /// Set compatible line spacing flag.
    pub fn set_body_pr_compat_ln_spc(&mut self, compat: bool) {
        self.body_pr_compat_ln_spc = Some(compat);
    }

    // ── Shape actions (click/hover) ──

    /// Gets the shape action settings (click and hover actions).
    pub fn action(&self) -> Option<&crate::actions::ShapeAction> {
        self.action.as_ref()
    }

    /// Sets the shape action settings.
    pub fn set_action(&mut self, action: crate::actions::ShapeAction) {
        self.action = Some(action);
    }

    /// Clears the shape action settings.
    pub fn clear_action(&mut self) {
        self.action = None;
    }

    // ── Embedded OLE objects ──

    /// Gets the embedded OLE object data.
    pub fn embedded_object(&self) -> Option<&crate::actions::EmbeddedObject> {
        self.embedded_object.as_ref()
    }

    /// Sets the embedded OLE object data.
    pub fn set_embedded_object(&mut self, obj: crate::actions::EmbeddedObject) {
        self.embedded_object = Some(obj);
    }

    /// Clears the embedded OLE object data.
    pub fn clear_embedded_object(&mut self) {
        self.embedded_object = None;
    }

    // ── Action button type ──

    /// Gets the action button preset type.
    pub fn action_button_type(&self) -> Option<crate::actions::ActionButtonType> {
        self.action_button_type
    }

    /// Sets the action button preset type.
    pub fn set_action_button_type(&mut self, button_type: crate::actions::ActionButtonType) {
        self.action_button_type = Some(button_type);
    }

    /// Clears the action button type.
    pub fn clear_action_button_type(&mut self) {
        self.action_button_type = None;
    }

    pub(crate) fn unknown_attrs(&self) -> &[(String, String)] {
        self.unknown_attrs.as_slice()
    }

    pub(crate) fn set_unknown_attrs(&mut self, attrs: Vec<(String, String)>) {
        self.unknown_attrs = attrs;
    }

    pub(crate) fn unknown_children(&self) -> &[RawXmlNode] {
        self.unknown_children.as_slice()
    }

    pub(crate) fn push_unknown_child(&mut self, node: RawXmlNode) {
        self.unknown_children.push(node);
    }

    /// Raw `<a:lstStyle>` element from `txBody`, preserved for roundtrip.
    pub(crate) fn lst_style_raw(&self) -> Option<&RawXmlNode> {
        self.lst_style_raw.as_ref()
    }

    /// Set the raw `<a:lstStyle>` element for roundtrip.
    pub(crate) fn set_lst_style_raw(&mut self, node: RawXmlNode) {
        self.lst_style_raw = Some(node);
    }

    /// Unknown children of `<a:bodyPr>`, preserved for roundtrip.
    pub(crate) fn body_pr_unknown_children(&self) -> &[RawXmlNode] {
        self.body_pr_unknown_children.as_slice()
    }

    /// Push an unknown child of `<a:bodyPr>`.
    pub(crate) fn push_body_pr_unknown_child(&mut self, node: RawXmlNode) {
        self.body_pr_unknown_children.push(node);
    }

    /// Unknown attributes on `<a:bodyPr>`, preserved for roundtrip.
    pub(crate) fn body_pr_unknown_attrs(&self) -> &[(String, String)] {
        self.body_pr_unknown_attrs.as_slice()
    }

    /// Set unknown attributes on `<a:bodyPr>`.
    pub(crate) fn set_body_pr_unknown_attrs(&mut self, attrs: Vec<(String, String)>) {
        self.body_pr_unknown_attrs = attrs;
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn shape_text_frame_stores_paragraph_runs() {
        let mut shape = Shape::new("Headline");
        shape.add_paragraph_with_text("Line 1");
        let paragraph = shape.add_paragraph();
        paragraph.add_run("Line");
        paragraph.add_run(" 2");

        assert_eq!(shape.paragraph_count(), 2);
        assert_eq!(shape.paragraphs()[0].run_count(), 1);
        assert_eq!(shape.paragraphs()[0].runs()[0].text(), "Line 1");
        assert_eq!(shape.paragraphs()[1].run_count(), 2);
        assert_eq!(shape.paragraphs()[1].runs()[0].text(), "Line");
        assert_eq!(shape.paragraphs()[1].runs()[1].text(), " 2");
    }

    #[test]
    fn shape_supports_placeholder_and_textbox_metadata() {
        let mut shape = Shape::new("Body Placeholder");

        assert_eq!(shape.shape_type(), ShapeType::AutoShape);
        assert_eq!(shape.placeholder_kind(), None);
        assert_eq!(shape.placeholder_idx(), None);
        assert_eq!(shape.geometry(), None);
        assert_eq!(shape.preset_geometry(), None);
        assert_eq!(shape.solid_fill_srgb(), None);

        shape.set_shape_type(ShapeType::TextBox);
        shape.set_placeholder_kind("body");
        shape.set_placeholder_idx(3);
        shape.set_geometry(ShapeGeometry::new(10, 20, 30, 40));
        shape.set_preset_geometry("roundRect");
        shape.set_solid_fill_srgb("AABBCC");

        assert_eq!(shape.shape_type(), ShapeType::TextBox);
        assert_eq!(shape.placeholder_kind(), Some("body"));
        assert_eq!(shape.placeholder_idx(), Some(3));
        assert_eq!(shape.geometry(), Some(ShapeGeometry::new(10, 20, 30, 40)));
        assert_eq!(shape.preset_geometry(), Some("roundRect"));
        assert_eq!(shape.solid_fill_srgb(), Some("AABBCC"));

        shape.clear_placeholder_kind();
        shape.clear_placeholder_idx();
        shape.clear_geometry();
        shape.clear_preset_geometry();
        shape.clear_solid_fill_srgb();

        assert_eq!(shape.placeholder_kind(), None);
        assert_eq!(shape.placeholder_idx(), None);
        assert_eq!(shape.geometry(), None);
        assert_eq!(shape.preset_geometry(), None);
        assert_eq!(shape.solid_fill_srgb(), None);
    }

    #[test]
    fn shape_outline_roundtrip() {
        let mut shape = Shape::new("Box");
        assert!(shape.outline().is_none());

        let mut outline = ShapeOutline::new();
        outline.width_emu = Some(12700);
        outline.color_srgb = Some("FF0000".to_string());
        outline.dash_style = Some(LineDashStyle::Dash);
        outline.compound_style = Some(LineCompoundStyle::Double);
        shape.set_outline(outline);

        let outline = shape.outline().unwrap();
        assert_eq!(outline.width_emu, Some(12700));
        assert_eq!(outline.color_srgb.as_deref(), Some("FF0000"));
        assert_eq!(outline.dash_style, Some(LineDashStyle::Dash));
        assert_eq!(outline.compound_style, Some(LineCompoundStyle::Double));

        shape.clear_outline();
        assert!(shape.outline().is_none());
    }

    #[test]
    fn line_dash_style_xml_roundtrip() {
        for (xml, style) in [
            ("solid", LineDashStyle::Solid),
            ("dot", LineDashStyle::Dot),
            ("dash", LineDashStyle::Dash),
            ("lgDash", LineDashStyle::LargeDash),
            ("dashDot", LineDashStyle::DashDot),
        ] {
            assert_eq!(LineDashStyle::from_xml(xml), Some(style));
            assert_eq!(style.to_xml(), xml);
        }
    }

    #[test]
    fn gradient_fill_roundtrip() {
        let mut shape = Shape::new("Box");
        assert!(shape.gradient_fill().is_none());

        let mut gradient = GradientFill::new();
        gradient.fill_type = Some(GradientFillType::Linear);
        gradient.linear_angle = Some(5400000);
        gradient.stops.push(GradientStop {
            position: 0,
            color_srgb: "FF0000".to_string(),
            color: None,
        });
        gradient.stops.push(GradientStop {
            position: 100000,
            color_srgb: "0000FF".to_string(),
            color: None,
        });
        shape.set_gradient_fill(gradient);

        assert!(shape.solid_fill_srgb().is_none());
        assert!(shape.pattern_fill().is_none());
        assert!(!shape.is_no_fill());

        let gradient = shape.gradient_fill().unwrap();
        assert_eq!(gradient.fill_type, Some(GradientFillType::Linear));
        assert_eq!(gradient.linear_angle, Some(5400000));
        assert_eq!(gradient.stops.len(), 2);
        assert_eq!(gradient.stops[0].color_srgb, "FF0000");
        assert_eq!(gradient.stops[1].position, 100000);

        shape.clear_gradient_fill();
        assert!(shape.gradient_fill().is_none());
    }

    #[test]
    fn pattern_fill_roundtrip() {
        let mut shape = Shape::new("Box");
        let mut pattern = PatternFill::new(PatternFillType::Cross);
        pattern.foreground_srgb = Some("FF0000".to_string());
        pattern.background_srgb = Some("FFFFFF".to_string());
        shape.set_pattern_fill(pattern);

        assert!(shape.solid_fill_srgb().is_none());
        assert!(shape.gradient_fill().is_none());

        let pattern = shape.pattern_fill().unwrap();
        assert_eq!(pattern.pattern_type.to_xml(), "cross");
        assert_eq!(pattern.foreground_srgb.as_deref(), Some("FF0000"));
        assert_eq!(pattern.background_srgb.as_deref(), Some("FFFFFF"));

        shape.clear_pattern_fill();
        assert!(shape.pattern_fill().is_none());
    }

    #[test]
    fn no_fill_flag() {
        let mut shape = Shape::new("Box");
        shape.set_solid_fill_srgb("AABBCC");
        shape.set_no_fill(true);
        assert!(shape.is_no_fill());
        assert!(shape.solid_fill_srgb().is_none());
        assert_eq!(shape.fill(), Some(ShapeFill::NoFill));

        shape.set_no_fill(false);
        assert!(!shape.is_no_fill());
    }

    #[test]
    fn rotation_roundtrip() {
        let mut shape = Shape::new("Box");
        assert_eq!(shape.rotation(), None);

        shape.set_rotation(5400000); // 90 degrees
        assert_eq!(shape.rotation(), Some(5400000));

        shape.clear_rotation();
        assert_eq!(shape.rotation(), None);
    }

    #[test]
    fn hidden_flag() {
        let mut shape = Shape::new("Box");
        assert!(!shape.is_hidden());

        shape.set_hidden(true);
        assert!(shape.is_hidden());

        shape.set_hidden(false);
        assert!(!shape.is_hidden());
    }

    #[test]
    fn bullet_properties_default() {
        let paragraph = ShapeParagraph::new();
        assert!(paragraph.properties().bullet.style.is_none());
        assert!(paragraph.properties().bullet.font_name.is_none());
        assert!(paragraph.properties().bullet.size_percent.is_none());
        assert!(paragraph.properties().bullet.color_srgb.is_none());
    }

    #[test]
    fn bullet_char_style() {
        let mut paragraph = ShapeParagraph::new();
        paragraph.properties_mut().bullet.style = Some(BulletStyle::Char("\u{2022}".to_string()));
        paragraph.properties_mut().bullet.font_name = Some("Arial".to_string());
        paragraph.properties_mut().bullet.size_percent = Some(100000);
        paragraph.properties_mut().bullet.color_srgb = Some("FF0000".to_string());

        let bullet = &paragraph.properties().bullet;
        assert_eq!(
            bullet.style,
            Some(BulletStyle::Char("\u{2022}".to_string()))
        );
        assert_eq!(bullet.font_name.as_deref(), Some("Arial"));
        assert_eq!(bullet.size_percent, Some(100000));
        assert_eq!(bullet.color_srgb.as_deref(), Some("FF0000"));
    }

    #[test]
    fn fill_unified_accessor() {
        let mut shape = Shape::new("Box");
        assert_eq!(shape.fill(), None);

        shape.set_solid_fill_srgb("AABBCC");
        assert_eq!(shape.fill(), Some(ShapeFill::Solid("AABBCC".to_string())));
    }

    // ── Picture fill tests ──

    #[test]
    fn picture_fill_roundtrip() {
        let mut shape = Shape::new("Box");
        assert!(shape.picture_fill().is_none());

        let fill = PictureFill::new("rId5");
        shape.set_picture_fill(fill);

        assert!(shape.solid_fill_srgb().is_none());
        assert!(shape.gradient_fill().is_none());
        assert!(shape.pattern_fill().is_none());
        assert!(!shape.is_no_fill());

        let pf = shape.picture_fill().unwrap();
        assert_eq!(pf.relationship_id, "rId5");
        assert!(pf.stretch);

        assert_eq!(
            shape.fill(),
            Some(ShapeFill::Picture(PictureFill {
                relationship_id: "rId5".to_string(),
                stretch: true,
                crop: None,
            }))
        );

        shape.clear_picture_fill();
        assert!(shape.picture_fill().is_none());
    }

    #[test]
    fn picture_fill_not_stretch() {
        let mut fill = PictureFill::new("rId3");
        fill.stretch = false;
        assert!(!fill.stretch);
        assert_eq!(fill.relationship_id, "rId3");
    }

    // ── Placeholder type enum tests ──

    #[test]
    fn placeholder_type_xml_roundtrip() {
        for (xml, expected) in [
            ("title", PlaceholderType::Title),
            ("body", PlaceholderType::Body),
            ("ctrTitle", PlaceholderType::CenteredTitle),
            ("subTitle", PlaceholderType::Subtitle),
            ("dt", PlaceholderType::DateAndTime),
            ("sldNum", PlaceholderType::SlideNumber),
            ("ftr", PlaceholderType::Footer),
            ("hdr", PlaceholderType::Header),
            ("obj", PlaceholderType::Object),
            ("chart", PlaceholderType::Chart),
            ("tbl", PlaceholderType::Table),
            ("clipArt", PlaceholderType::ClipArt),
            ("dgm", PlaceholderType::Diagram),
            ("media", PlaceholderType::Media),
            ("sldImg", PlaceholderType::SlideImage),
        ] {
            let parsed = PlaceholderType::from_xml(xml);
            assert_eq!(parsed, expected);
            assert_eq!(parsed.to_xml(), xml);
        }
    }

    #[test]
    fn placeholder_type_other_roundtrip() {
        let ph = PlaceholderType::from_xml("customType");
        assert_eq!(ph, PlaceholderType::Other("customType".to_string()));
        assert_eq!(ph.to_xml(), "customType");
    }

    #[test]
    fn shape_placeholder_type_accessors() {
        let mut shape = Shape::new("Title");
        assert!(shape.placeholder_type().is_none());

        shape.set_placeholder_type(PlaceholderType::Title);
        assert_eq!(shape.placeholder_type(), Some(&PlaceholderType::Title));
        // Also syncs the string-based kind.
        assert_eq!(shape.placeholder_kind(), Some("title"));

        shape.clear_placeholder_type();
        assert!(shape.placeholder_type().is_none());
        assert!(shape.placeholder_kind().is_none());
    }

    // ── Alt text tests ──

    #[test]
    fn alt_text_roundtrip() {
        let mut shape = Shape::new("Image1");
        assert!(shape.alt_text().is_none());
        assert!(shape.alt_text_title().is_none());

        shape.set_alt_text("A sunset over mountains");
        shape.set_alt_text_title("Sunset Photo");

        assert_eq!(shape.alt_text(), Some("A sunset over mountains"));
        assert_eq!(shape.alt_text_title(), Some("Sunset Photo"));

        shape.clear_alt_text();
        shape.clear_alt_text_title();
        assert!(shape.alt_text().is_none());
        assert!(shape.alt_text_title().is_none());
    }

    #[test]
    fn alt_text_independent_of_other_shape_props() {
        let mut shape = Shape::new("Box");
        shape.set_alt_text("Description");
        shape.set_hidden(true);

        assert_eq!(shape.alt_text(), Some("Description"));
        assert!(shape.is_hidden());
    }

    // ── SmartArt detection tests ──

    #[test]
    fn smartart_detection_flag() {
        let mut shape = Shape::new("Diagram");
        assert!(!shape.is_smartart());

        shape.set_smartart(true);
        assert!(shape.is_smartart());

        shape.set_smartart(false);
        assert!(!shape.is_smartart());
    }

    #[test]
    fn smartart_flag_independent() {
        let mut shape = Shape::new("SmartArt");
        shape.set_smartart(true);
        shape.set_hidden(false);
        assert!(shape.is_smartart());
        assert!(!shape.is_hidden());
    }

    // ── Connector shapes tests ──

    #[test]
    fn connector_shape_roundtrip() {
        let mut shape = Shape::new("Connector 1");
        assert!(!shape.is_connector());
        assert!(shape.start_connection().is_none());
        assert!(shape.end_connection().is_none());

        shape.set_connector(true);
        shape.set_start_connection(ConnectionInfo::new(5, 0));
        shape.set_end_connection(ConnectionInfo::new(7, 2));

        assert!(shape.is_connector());
        let start = shape.start_connection().unwrap();
        assert_eq!(start.shape_id, 5);
        assert_eq!(start.connection_point_index, 0);
        let end = shape.end_connection().unwrap();
        assert_eq!(end.shape_id, 7);
        assert_eq!(end.connection_point_index, 2);

        shape.clear_start_connection();
        shape.clear_end_connection();
        assert!(shape.start_connection().is_none());
        assert!(shape.end_connection().is_none());
    }

    #[test]
    fn connector_without_connections() {
        let mut shape = Shape::new("Line");
        shape.set_connector(true);
        assert!(shape.is_connector());
        assert!(shape.start_connection().is_none());
        assert!(shape.end_connection().is_none());
    }

    // ── Audio/video media tests ──

    #[test]
    fn media_roundtrip() {
        let mut shape = Shape::new("Video 1");
        assert!(shape.media().is_none());

        shape.set_media(MediaType::Video, "rId3");
        let (media_type, rid) = shape.media().unwrap();
        assert_eq!(*media_type, MediaType::Video);
        assert_eq!(rid, "rId3");

        shape.clear_media();
        assert!(shape.media().is_none());
    }

    #[test]
    fn media_audio_type() {
        let mut shape = Shape::new("Audio 1");
        shape.set_media(MediaType::Audio, "rId7");
        let (media_type, rid) = shape.media().unwrap();
        assert_eq!(*media_type, MediaType::Audio);
        assert_eq!(rid, "rId7");
    }

    // ── Shadow effect tests ──

    #[test]
    fn shadow_roundtrip() {
        let mut shape = Shape::new("Box");
        assert!(shape.shadow().is_none());

        let mut shadow = ShapeShadow::new(50800, 50800, 63500, "000000");
        shadow.alpha = Some(50);
        shape.set_shadow(shadow);

        let shadow = shape.shadow().unwrap();
        assert_eq!(shadow.offset_x, 50800);
        assert_eq!(shadow.offset_y, 50800);
        assert_eq!(shadow.blur_radius, 63500);
        assert_eq!(shadow.color, "000000");
        assert_eq!(shadow.alpha, Some(50));

        shape.clear_shadow();
        assert!(shape.shadow().is_none());
    }

    #[test]
    fn shadow_without_alpha() {
        let shadow = ShapeShadow::new(25400, 25400, 38100, "FF0000");
        assert_eq!(shadow.alpha, None);
        assert_eq!(shadow.offset_x, 25400);
        assert_eq!(shadow.color, "FF0000");
    }

    // ── Glow effect tests ──

    #[test]
    fn glow_roundtrip() {
        let mut shape = Shape::new("Box");
        assert!(shape.glow().is_none());

        let mut glow = ShapeGlow::new(101600, "FFC000");
        glow.alpha = Some(75);
        shape.set_glow(glow);

        let glow = shape.glow().unwrap();
        assert_eq!(glow.radius, 101600);
        assert_eq!(glow.color, "FFC000");
        assert_eq!(glow.alpha, Some(75));

        shape.clear_glow();
        assert!(shape.glow().is_none());
    }

    #[test]
    fn glow_without_alpha() {
        let glow = ShapeGlow::new(50800, "00FF00");
        assert_eq!(glow.alpha, None);
        assert_eq!(glow.radius, 50800);
    }

    // ── Reflection effect tests ──

    #[test]
    fn reflection_roundtrip() {
        let mut shape = Shape::new("Box");
        assert!(shape.reflection().is_none());

        let mut reflection = ShapeReflection::new(6350, 0);
        reflection.start_alpha = Some(50);
        reflection.end_alpha = Some(0);
        reflection.direction = Some(5400000);
        shape.set_reflection(reflection);

        let reflection = shape.reflection().unwrap();
        assert_eq!(reflection.blur_radius, 6350);
        assert_eq!(reflection.distance, 0);
        assert_eq!(reflection.start_alpha, Some(50));
        assert_eq!(reflection.end_alpha, Some(0));
        assert_eq!(reflection.direction, Some(5400000));

        shape.clear_reflection();
        assert!(shape.reflection().is_none());
    }

    #[test]
    fn reflection_minimal() {
        let reflection = ShapeReflection::new(12700, 25400);
        assert_eq!(reflection.blur_radius, 12700);
        assert_eq!(reflection.distance, 25400);
        assert_eq!(reflection.start_alpha, None);
        assert_eq!(reflection.end_alpha, None);
        assert_eq!(reflection.direction, None);
    }

    // ── Text anchor tests ──

    #[test]
    fn text_anchor_roundtrip() {
        let mut shape = Shape::new("Box");
        assert!(shape.text_anchor().is_none());

        shape.set_text_anchor(TextAnchor::Middle);
        assert_eq!(shape.text_anchor(), Some(TextAnchor::Middle));

        shape.set_text_anchor(TextAnchor::BottomCentered);
        assert_eq!(shape.text_anchor(), Some(TextAnchor::BottomCentered));

        shape.clear_text_anchor();
        assert!(shape.text_anchor().is_none());
    }

    #[test]
    fn text_anchor_from_xml_all_variants() {
        assert_eq!(TextAnchor::from_xml("t", false), Some(TextAnchor::Top));
        assert_eq!(
            TextAnchor::from_xml("t", true),
            Some(TextAnchor::TopCentered)
        );
        assert_eq!(TextAnchor::from_xml("ctr", false), Some(TextAnchor::Middle));
        assert_eq!(
            TextAnchor::from_xml("ctr", true),
            Some(TextAnchor::MiddleCentered)
        );
        assert_eq!(TextAnchor::from_xml("b", false), Some(TextAnchor::Bottom));
        assert_eq!(
            TextAnchor::from_xml("b", true),
            Some(TextAnchor::BottomCentered)
        );
        assert_eq!(TextAnchor::from_xml("unknown", false), None);
    }

    #[test]
    fn text_anchor_to_xml_roundtrip() {
        for (anchor, expected_xml, expected_ctr) in [
            (TextAnchor::Top, "t", false),
            (TextAnchor::TopCentered, "t", true),
            (TextAnchor::Middle, "ctr", false),
            (TextAnchor::MiddleCentered, "ctr", true),
            (TextAnchor::Bottom, "b", false),
            (TextAnchor::BottomCentered, "b", true),
        ] {
            assert_eq!(anchor.to_xml_anchor(), expected_xml);
            assert_eq!(anchor.is_centered(), expected_ctr);
        }
    }

    // ── Auto-fit tests ──

    #[test]
    fn auto_fit_roundtrip() {
        let mut shape = Shape::new("Box");
        assert!(shape.auto_fit().is_none());

        shape.set_auto_fit(AutoFitType::Normal);
        assert_eq!(shape.auto_fit(), Some(AutoFitType::Normal));

        shape.set_auto_fit(AutoFitType::ShrinkOnOverflow);
        assert_eq!(shape.auto_fit(), Some(AutoFitType::ShrinkOnOverflow));

        shape.set_auto_fit(AutoFitType::None);
        assert_eq!(shape.auto_fit(), Some(AutoFitType::None));

        shape.clear_auto_fit();
        assert!(shape.auto_fit().is_none());
    }

    #[test]
    fn auto_fit_type_from_xml_tag() {
        assert_eq!(
            AutoFitType::from_xml_tag(b"noAutofit"),
            Some(AutoFitType::None)
        );
        assert_eq!(
            AutoFitType::from_xml_tag(b"normAutofit"),
            Some(AutoFitType::Normal)
        );
        assert_eq!(
            AutoFitType::from_xml_tag(b"spAutoFit"),
            Some(AutoFitType::ShrinkOnOverflow)
        );
        assert_eq!(AutoFitType::from_xml_tag(b"unknown"), None);
    }

    #[test]
    fn auto_fit_type_to_xml_tag() {
        assert_eq!(AutoFitType::None.to_xml_tag(), "a:noAutofit");
        assert_eq!(AutoFitType::Normal.to_xml_tag(), "a:normAutofit");
        assert_eq!(AutoFitType::ShrinkOnOverflow.to_xml_tag(), "a:spAutoFit");
    }

    // ── Combined effects test ──

    #[test]
    fn shape_with_all_effects() {
        let mut shape = Shape::new("Fancy Box");
        shape.set_shadow(ShapeShadow::new(50800, 50800, 63500, "000000"));
        shape.set_glow(ShapeGlow::new(101600, "FFC000"));
        shape.set_reflection(ShapeReflection::new(6350, 0));
        shape.set_text_anchor(TextAnchor::Middle);
        shape.set_auto_fit(AutoFitType::Normal);

        assert!(shape.shadow().is_some());
        assert!(shape.glow().is_some());
        assert!(shape.reflection().is_some());
        assert_eq!(shape.text_anchor(), Some(TextAnchor::Middle));
        assert_eq!(shape.auto_fit(), Some(AutoFitType::Normal));
    }

    // ── Arrow type tests ──

    #[test]
    fn arrow_type_xml_roundtrip() {
        for (xml, expected) in [
            ("none", ArrowType::None),
            ("triangle", ArrowType::Triangle),
            ("stealth", ArrowType::Stealth),
            ("diamond", ArrowType::Diamond),
            ("oval", ArrowType::Oval),
            ("arrow", ArrowType::Arrow),
            ("open", ArrowType::Open),
        ] {
            assert_eq!(ArrowType::from_xml(xml), Some(expected));
            assert_eq!(expected.to_xml(), xml);
        }
        assert_eq!(ArrowType::from_xml("unknown"), None);
    }

    #[test]
    fn arrow_size_xml_roundtrip() {
        for (xml, expected) in [
            ("sm", ArrowSize::Small),
            ("med", ArrowSize::Medium),
            ("lg", ArrowSize::Large),
        ] {
            assert_eq!(ArrowSize::from_xml(xml), Some(expected));
            assert_eq!(expected.to_xml(), xml);
        }
        assert_eq!(ArrowSize::from_xml("unknown"), None);
    }

    #[test]
    fn line_arrow_defaults() {
        let arrow = LineArrow::new(ArrowType::Triangle);
        assert_eq!(arrow.arrow_type, ArrowType::Triangle);
        assert_eq!(arrow.width, ArrowSize::Medium);
        assert_eq!(arrow.length, ArrowSize::Medium);
    }

    #[test]
    fn outline_arrow_roundtrip() {
        let mut outline = ShapeOutline::new();
        assert!(outline.head_arrow.is_none());
        assert!(outline.tail_arrow.is_none());

        let head = LineArrow {
            arrow_type: ArrowType::Triangle,
            width: ArrowSize::Large,
            length: ArrowSize::Small,
        };
        let tail = LineArrow {
            arrow_type: ArrowType::Stealth,
            width: ArrowSize::Medium,
            length: ArrowSize::Large,
        };
        outline.head_arrow = Some(head);
        outline.tail_arrow = Some(tail);

        assert!(outline.is_set());
        let head = outline.head_arrow.unwrap();
        assert_eq!(head.arrow_type, ArrowType::Triangle);
        assert_eq!(head.width, ArrowSize::Large);
        assert_eq!(head.length, ArrowSize::Small);
        let tail = outline.tail_arrow.unwrap();
        assert_eq!(tail.arrow_type, ArrowType::Stealth);
        assert_eq!(tail.width, ArrowSize::Medium);
        assert_eq!(tail.length, ArrowSize::Large);
    }

    // ── Transparency/opacity tests ──

    #[test]
    fn solid_fill_alpha_roundtrip() {
        let mut shape = Shape::new("Box");
        assert_eq!(shape.solid_fill_alpha(), None);

        shape.set_solid_fill_srgb("FF0000");
        shape.set_solid_fill_alpha(50);
        assert_eq!(shape.solid_fill_alpha(), Some(50));

        shape.clear_solid_fill_alpha();
        assert_eq!(shape.solid_fill_alpha(), None);
    }

    #[test]
    fn outline_alpha_roundtrip() {
        let mut outline = ShapeOutline::new();
        assert_eq!(outline.alpha, None);

        outline.alpha = Some(75);
        assert_eq!(outline.alpha, Some(75));
        assert!(outline.is_set());
    }

    // ── Line spacing type tests ──

    #[test]
    fn line_spacing_percent() {
        let ls = LineSpacing::percent(150000);
        assert_eq!(ls.value, 150000);
        assert_eq!(ls.unit, LineSpacingUnit::Percent);
    }

    #[test]
    fn line_spacing_points() {
        let ls = LineSpacing::points(1200);
        assert_eq!(ls.value, 1200);
        assert_eq!(ls.unit, LineSpacingUnit::Points);
    }

    #[test]
    fn spacing_value_percent() {
        let sv = SpacingValue::percent(50000);
        assert_eq!(sv.value, 50000);
        assert_eq!(sv.unit, SpacingUnit::Percent);
    }

    #[test]
    fn spacing_value_points() {
        let sv = SpacingValue::points(600);
        assert_eq!(sv.value, 600);
        assert_eq!(sv.unit, SpacingUnit::Points);
    }

    #[test]
    fn paragraph_typed_spacing_fields() {
        let mut props = ParagraphProperties::default();
        assert!(props.line_spacing.is_none());
        assert!(props.space_before.is_none());
        assert!(props.space_after.is_none());

        props.line_spacing = Some(LineSpacing::percent(150000));
        props.space_before = Some(SpacingValue::points(600));
        props.space_after = Some(SpacingValue::percent(20000));

        let ls = props.line_spacing.unwrap();
        assert_eq!(ls.value, 150000);
        assert_eq!(ls.unit, LineSpacingUnit::Percent);

        let sb = props.space_before.unwrap();
        assert_eq!(sb.value, 600);
        assert_eq!(sb.unit, SpacingUnit::Points);

        let sa = props.space_after.unwrap();
        assert_eq!(sa.value, 20000);
        assert_eq!(sa.unit, SpacingUnit::Percent);
    }

    // ── Text direction tests ──

    #[test]
    fn text_direction_roundtrip() {
        let mut shape = Shape::new("Box");
        assert!(shape.text_direction().is_none());

        shape.set_text_direction(TextDirection::Rotate270);
        assert_eq!(shape.text_direction(), Some(TextDirection::Rotate270));

        shape.set_text_direction(TextDirection::Horizontal);
        assert_eq!(shape.text_direction(), Some(TextDirection::Horizontal));

        shape.clear_text_direction();
        assert!(shape.text_direction().is_none());
    }

    #[test]
    fn text_direction_all_variants() {
        let mut shape = Shape::new("Box");
        for direction in [
            TextDirection::Horizontal,
            TextDirection::Rotate90,
            TextDirection::Rotate270,
            TextDirection::Stacked,
        ] {
            shape.set_text_direction(direction);
            assert_eq!(shape.text_direction(), Some(direction));
        }
    }

    // ── Text columns tests ──

    #[test]
    fn text_columns_roundtrip() {
        let mut shape = Shape::new("Box");
        assert!(shape.text_columns().is_none());
        assert!(shape.text_column_spacing().is_none());

        shape.set_text_columns(2);
        shape.set_text_column_spacing(457200);
        assert_eq!(shape.text_columns(), Some(2));
        assert_eq!(shape.text_column_spacing(), Some(457200));

        shape.clear_text_columns();
        shape.clear_text_column_spacing();
        assert!(shape.text_columns().is_none());
        assert!(shape.text_column_spacing().is_none());
    }

    // ── Text insets/margins tests ──

    #[test]
    fn text_insets_roundtrip() {
        let mut shape = Shape::new("Box");
        assert!(shape.text_inset_left().is_none());
        assert!(shape.text_inset_right().is_none());
        assert!(shape.text_inset_top().is_none());
        assert!(shape.text_inset_bottom().is_none());

        shape.set_text_inset_left(91440);
        shape.set_text_inset_right(91440);
        shape.set_text_inset_top(45720);
        shape.set_text_inset_bottom(45720);

        assert_eq!(shape.text_inset_left(), Some(91440));
        assert_eq!(shape.text_inset_right(), Some(91440));
        assert_eq!(shape.text_inset_top(), Some(45720));
        assert_eq!(shape.text_inset_bottom(), Some(45720));

        shape.clear_text_inset_left();
        shape.clear_text_inset_right();
        shape.clear_text_inset_top();
        shape.clear_text_inset_bottom();

        assert!(shape.text_inset_left().is_none());
        assert!(shape.text_inset_right().is_none());
        assert!(shape.text_inset_top().is_none());
        assert!(shape.text_inset_bottom().is_none());
    }

    #[test]
    fn text_insets_zero_is_valid() {
        let mut shape = Shape::new("Box");
        shape.set_text_inset_left(0);
        shape.set_text_inset_top(0);
        assert_eq!(shape.text_inset_left(), Some(0));
        assert_eq!(shape.text_inset_top(), Some(0));
    }

    // ── Word wrap tests ──

    #[test]
    fn word_wrap_roundtrip() {
        let mut shape = Shape::new("Box");
        assert!(shape.word_wrap().is_none());

        shape.set_word_wrap(true);
        assert_eq!(shape.word_wrap(), Some(true));

        shape.set_word_wrap(false);
        assert_eq!(shape.word_wrap(), Some(false));

        shape.clear_word_wrap();
        assert!(shape.word_wrap().is_none());
    }

    // ── Combined body properties test ──

    #[test]
    fn shape_with_all_body_properties() {
        let mut shape = Shape::new("Text Box");
        shape.set_text_anchor(TextAnchor::Middle);
        shape.set_auto_fit(AutoFitType::Normal);
        shape.set_text_direction(TextDirection::Rotate270);
        shape.set_text_columns(3);
        shape.set_text_column_spacing(228600);
        shape.set_text_inset_left(91440);
        shape.set_text_inset_right(91440);
        shape.set_text_inset_top(45720);
        shape.set_text_inset_bottom(45720);
        shape.set_word_wrap(true);

        assert_eq!(shape.text_anchor(), Some(TextAnchor::Middle));
        assert_eq!(shape.auto_fit(), Some(AutoFitType::Normal));
        assert_eq!(shape.text_direction(), Some(TextDirection::Rotate270));
        assert_eq!(shape.text_columns(), Some(3));
        assert_eq!(shape.text_column_spacing(), Some(228600));
        assert_eq!(shape.text_inset_left(), Some(91440));
        assert_eq!(shape.text_inset_right(), Some(91440));
        assert_eq!(shape.text_inset_top(), Some(45720));
        assert_eq!(shape.text_inset_bottom(), Some(45720));
        assert_eq!(shape.word_wrap(), Some(true));
    }
}
