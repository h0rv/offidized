// Text run formatting and content for PresentationML shapes.

use crate::color::ShapeColor;

/// Underline style for a text run.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum UnderlineStyle {
    Single,
    Double,
    Heavy,
    Dotted,
    DottedHeavy,
    Dash,
    DashHeavy,
    DashLong,
    DashLongHeavy,
    DotDashHeavy,
    DotDotDash,
    DotDotDashHeavy,
    Wavy,
    WavyHeavy,
    WavyDouble,
}

impl UnderlineStyle {
    /// Convert from the XML attribute value.
    pub fn from_xml(value: &str) -> Option<Self> {
        match value {
            "sng" => Some(Self::Single),
            "dbl" => Some(Self::Double),
            "heavy" => Some(Self::Heavy),
            "dotted" => Some(Self::Dotted),
            "dottedHeavy" => Some(Self::DottedHeavy),
            "dash" => Some(Self::Dash),
            "dashHeavy" => Some(Self::DashHeavy),
            "dashLong" => Some(Self::DashLong),
            "dashLongHeavy" => Some(Self::DashLongHeavy),
            "dotDashHeavy" => Some(Self::DotDashHeavy),
            "dotDotDash" => Some(Self::DotDotDash),
            "dotDotDashHeavy" => Some(Self::DotDotDashHeavy),
            "wavy" => Some(Self::Wavy),
            "wavyHeavy" => Some(Self::WavyHeavy),
            "wavyDbl" => Some(Self::WavyDouble),
            _ => None,
        }
    }

    /// Convert to the XML attribute value.
    pub fn to_xml(self) -> &'static str {
        match self {
            Self::Single => "sng",
            Self::Double => "dbl",
            Self::Heavy => "heavy",
            Self::Dotted => "dotted",
            Self::DottedHeavy => "dottedHeavy",
            Self::Dash => "dash",
            Self::DashHeavy => "dashHeavy",
            Self::DashLong => "dashLong",
            Self::DashLongHeavy => "dashLongHeavy",
            Self::DotDashHeavy => "dotDashHeavy",
            Self::DotDotDash => "dotDotDash",
            Self::DotDotDashHeavy => "dotDotDashHeavy",
            Self::Wavy => "wavy",
            Self::WavyHeavy => "wavyHeavy",
            Self::WavyDouble => "wavyDbl",
        }
    }
}

/// Strikethrough style for a text run.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum StrikethroughStyle {
    Single,
    Double,
}

impl StrikethroughStyle {
    pub fn from_xml(value: &str) -> Option<Self> {
        match value {
            "sngStrike" => Some(Self::Single),
            "dblStrike" => Some(Self::Double),
            _ => None,
        }
    }

    pub fn to_xml(self) -> &'static str {
        match self {
            Self::Single => "sngStrike",
            Self::Double => "dblStrike",
        }
    }
}

/// Run-level font and formatting properties parsed from `<a:rPr>`.
#[derive(Debug, Clone, PartialEq, Eq, Default)]
pub struct RunProperties {
    /// Bold (`b` attribute).
    pub bold: Option<bool>,
    /// Italic (`i` attribute).
    pub italic: Option<bool>,
    /// Underline style (`u` attribute).
    pub underline: Option<UnderlineStyle>,
    /// Strikethrough style (`strike` attribute).
    pub strikethrough: Option<StrikethroughStyle>,
    /// Font size in hundredths of a point (`sz` attribute). 2400 = 24pt.
    pub font_size: Option<u32>,
    /// Font color as sRGB hex (`<a:solidFill><a:srgbClr val="FF0000"/>...`).
    pub font_color_srgb: Option<String>,
    /// Font color (full model, supports scheme colors and transforms).
    /// When set, this takes precedence over `font_color_srgb` during serialization.
    pub font_color: Option<ShapeColor>,
    /// Latin font name (`<a:latin typeface="Arial"/>`).
    pub font_name: Option<String>,
    /// East Asian font name (`<a:ea typeface="..."/>`).
    pub font_name_east_asian: Option<String>,
    /// Complex Script font name (`<a:cs typeface="..."/>`).
    pub font_name_complex_script: Option<String>,
    /// Language tag (`lang` attribute).
    pub language: Option<String>,
    /// Hyperlink click relationship ID (`<a:hlinkClick r:id="..."/>`). (Feature #10).
    pub hyperlink_click_rid: Option<String>,
    /// Resolved hyperlink URL (from the relationship target). When set, this is the
    /// actual URL the hyperlink points to, resolved from the slide's relationships.
    pub hyperlink_url: Option<String>,
    /// Hyperlink tooltip text (`tooltip` attribute on `<a:hlinkClick>`).
    pub hyperlink_tooltip: Option<String>,
    /// Character spacing in hundredths of a point (`spc` attribute on `<a:rPr>`).
    /// Negative values condense spacing, positive values expand it.
    pub character_spacing: Option<i32>,
    /// Kerning size threshold in hundredths of a point (`kern` attribute on `<a:rPr>`).
    /// When set, kerning is enabled for characters at or above this font size.
    /// A value of 0 disables kerning.
    pub kerning: Option<i32>,
    /// Baseline offset as a percentage. Positive values create superscript,
    /// negative values create subscript. (`baseline` attribute on `<a:rPr>`).
    /// Typical values: 30000 for superscript, -25000 for subscript (in 1/1000ths of %).
    pub baseline: Option<i32>,
}

#[derive(Debug, Clone, PartialEq, Eq, Default)]
pub struct TextRun {
    text: String,
    properties: RunProperties,
}

impl TextRun {
    pub fn new(text: impl Into<String>) -> Self {
        Self {
            text: text.into(),
            properties: RunProperties::default(),
        }
    }

    pub fn text(&self) -> &str {
        &self.text
    }

    pub fn set_text(&mut self, text: impl Into<String>) {
        self.text = text.into();
    }

    /// Access the run formatting properties.
    pub fn properties(&self) -> &RunProperties {
        &self.properties
    }

    /// Mutable access to the run formatting properties.
    pub fn properties_mut(&mut self) -> &mut RunProperties {
        &mut self.properties
    }

    pub(crate) fn set_properties(&mut self, properties: RunProperties) {
        self.properties = properties;
    }

    // Convenience getters/setters for the most common properties.

    pub fn is_bold(&self) -> bool {
        self.properties.bold.unwrap_or(false)
    }

    pub fn set_bold(&mut self, bold: bool) -> &mut Self {
        self.properties.bold = Some(bold);
        self
    }

    pub fn is_italic(&self) -> bool {
        self.properties.italic.unwrap_or(false)
    }

    pub fn set_italic(&mut self, italic: bool) -> &mut Self {
        self.properties.italic = Some(italic);
        self
    }

    pub fn underline(&self) -> Option<UnderlineStyle> {
        self.properties.underline
    }

    pub fn set_underline(&mut self, style: UnderlineStyle) -> &mut Self {
        self.properties.underline = Some(style);
        self
    }

    pub fn clear_underline(&mut self) -> &mut Self {
        self.properties.underline = None;
        self
    }

    pub fn strikethrough(&self) -> Option<StrikethroughStyle> {
        self.properties.strikethrough
    }

    pub fn set_strikethrough(&mut self, style: StrikethroughStyle) -> &mut Self {
        self.properties.strikethrough = Some(style);
        self
    }

    pub fn clear_strikethrough(&mut self) -> &mut Self {
        self.properties.strikethrough = None;
        self
    }

    /// Font size in hundredths of a point (e.g., 2400 = 24pt).
    pub fn font_size(&self) -> Option<u32> {
        self.properties.font_size
    }

    /// Set font size in hundredths of a point (e.g., 2400 = 24pt).
    pub fn set_font_size(&mut self, hundredths_of_point: u32) -> &mut Self {
        self.properties.font_size = Some(hundredths_of_point);
        self
    }

    pub fn clear_font_size(&mut self) -> &mut Self {
        self.properties.font_size = None;
        self
    }

    /// Font color as sRGB hex string (e.g., "FF0000" for red).
    pub fn font_color(&self) -> Option<&str> {
        self.properties.font_color_srgb.as_deref()
    }

    pub fn set_font_color(&mut self, srgb_hex: impl Into<String>) -> &mut Self {
        self.properties.font_color_srgb = Some(srgb_hex.into());
        self
    }

    pub fn clear_font_color(&mut self) -> &mut Self {
        self.properties.font_color_srgb = None;
        self
    }

    /// Latin font name (e.g., "Arial", "Calibri").
    pub fn font_name(&self) -> Option<&str> {
        self.properties.font_name.as_deref()
    }

    pub fn set_font_name(&mut self, name: impl Into<String>) -> &mut Self {
        self.properties.font_name = Some(name.into());
        self
    }

    pub fn clear_font_name(&mut self) -> &mut Self {
        self.properties.font_name = None;
        self
    }

    /// Language tag (e.g., "en-US").
    pub fn language(&self) -> Option<&str> {
        self.properties.language.as_deref()
    }

    pub fn set_language(&mut self, lang: impl Into<String>) -> &mut Self {
        self.properties.language = Some(lang.into());
        self
    }

    pub fn clear_language(&mut self) -> &mut Self {
        self.properties.language = None;
        self
    }

    // ── Feature #10: Hyperlinks ──

    /// Hyperlink click relationship ID.
    pub fn hyperlink_click_rid(&self) -> Option<&str> {
        self.properties.hyperlink_click_rid.as_deref()
    }

    pub fn set_hyperlink_click_rid(&mut self, rid: impl Into<String>) -> &mut Self {
        self.properties.hyperlink_click_rid = Some(rid.into());
        self
    }

    pub fn clear_hyperlink_click_rid(&mut self) -> &mut Self {
        self.properties.hyperlink_click_rid = None;
        self
    }

    // ── URL Hyperlinks ──

    /// Resolved hyperlink URL.
    pub fn hyperlink_url(&self) -> Option<&str> {
        self.properties.hyperlink_url.as_deref()
    }

    pub fn set_hyperlink_url(&mut self, url: impl Into<String>) -> &mut Self {
        self.properties.hyperlink_url = Some(url.into());
        self
    }

    pub fn clear_hyperlink_url(&mut self) -> &mut Self {
        self.properties.hyperlink_url = None;
        self
    }

    /// Hyperlink tooltip text.
    pub fn hyperlink_tooltip(&self) -> Option<&str> {
        self.properties.hyperlink_tooltip.as_deref()
    }

    pub fn set_hyperlink_tooltip(&mut self, tooltip: impl Into<String>) -> &mut Self {
        self.properties.hyperlink_tooltip = Some(tooltip.into());
        self
    }

    pub fn clear_hyperlink_tooltip(&mut self) -> &mut Self {
        self.properties.hyperlink_tooltip = None;
        self
    }

    // ── Character Spacing ──

    /// Character spacing in hundredths of a point. Negative = condensed.
    pub fn character_spacing(&self) -> Option<i32> {
        self.properties.character_spacing
    }

    pub fn set_character_spacing(&mut self, hundredths_of_point: i32) -> &mut Self {
        self.properties.character_spacing = Some(hundredths_of_point);
        self
    }

    pub fn clear_character_spacing(&mut self) -> &mut Self {
        self.properties.character_spacing = None;
        self
    }

    // ── Kerning ──

    /// Kerning size threshold in hundredths of a point.
    pub fn kerning(&self) -> Option<i32> {
        self.properties.kerning
    }

    pub fn set_kerning(&mut self, hundredths_of_point: i32) -> &mut Self {
        self.properties.kerning = Some(hundredths_of_point);
        self
    }

    pub fn clear_kerning(&mut self) -> &mut Self {
        self.properties.kerning = None;
        self
    }

    // ── Subscript/Superscript ──

    /// Baseline offset as a percentage (in 1/1000ths of %).
    /// Positive = superscript, negative = subscript.
    pub fn baseline(&self) -> Option<i32> {
        self.properties.baseline
    }

    pub fn set_baseline(&mut self, baseline: i32) -> &mut Self {
        self.properties.baseline = Some(baseline);
        self
    }

    pub fn clear_baseline(&mut self) -> &mut Self {
        self.properties.baseline = None;
        self
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn set_text_updates_value() {
        let mut run = TextRun::new("Old");
        run.set_text("New");

        assert_eq!(run.text(), "New");
    }

    #[test]
    fn formatting_defaults_are_unset() {
        let run = TextRun::new("Hello");
        assert!(!run.is_bold());
        assert!(!run.is_italic());
        assert_eq!(run.underline(), None);
        assert_eq!(run.strikethrough(), None);
        assert_eq!(run.font_size(), None);
        assert_eq!(run.font_color(), None);
        assert_eq!(run.font_name(), None);
        assert_eq!(run.language(), None);
    }

    #[test]
    fn formatting_roundtrip() {
        let mut run = TextRun::new("Styled");
        run.set_bold(true)
            .set_italic(true)
            .set_underline(UnderlineStyle::Single)
            .set_strikethrough(StrikethroughStyle::Single)
            .set_font_size(2400)
            .set_font_color("FF0000")
            .set_font_name("Arial")
            .set_language("en-US");

        assert!(run.is_bold());
        assert!(run.is_italic());
        assert_eq!(run.underline(), Some(UnderlineStyle::Single));
        assert_eq!(run.strikethrough(), Some(StrikethroughStyle::Single));
        assert_eq!(run.font_size(), Some(2400));
        assert_eq!(run.font_color(), Some("FF0000"));
        assert_eq!(run.font_name(), Some("Arial"));
        assert_eq!(run.language(), Some("en-US"));

        run.clear_underline()
            .clear_strikethrough()
            .clear_font_size()
            .clear_font_color()
            .clear_font_name()
            .clear_language();

        assert_eq!(run.underline(), None);
        assert_eq!(run.strikethrough(), None);
        assert_eq!(run.font_size(), None);
        assert_eq!(run.font_color(), None);
        assert_eq!(run.font_name(), None);
        assert_eq!(run.language(), None);
    }

    #[test]
    fn underline_xml_roundtrip() {
        for (xml, style) in [
            ("sng", UnderlineStyle::Single),
            ("dbl", UnderlineStyle::Double),
            ("heavy", UnderlineStyle::Heavy),
            ("wavy", UnderlineStyle::Wavy),
            ("wavyDbl", UnderlineStyle::WavyDouble),
        ] {
            assert_eq!(UnderlineStyle::from_xml(xml), Some(style));
            assert_eq!(style.to_xml(), xml);
        }
    }

    #[test]
    fn strikethrough_xml_roundtrip() {
        assert_eq!(
            StrikethroughStyle::from_xml("sngStrike"),
            Some(StrikethroughStyle::Single)
        );
        assert_eq!(
            StrikethroughStyle::from_xml("dblStrike"),
            Some(StrikethroughStyle::Double)
        );
        assert_eq!(StrikethroughStyle::Single.to_xml(), "sngStrike");
        assert_eq!(StrikethroughStyle::Double.to_xml(), "dblStrike");
    }

    #[test]
    fn hyperlink_url_roundtrip() {
        let mut run = TextRun::new("Click here");
        assert_eq!(run.hyperlink_url(), None);
        assert_eq!(run.hyperlink_tooltip(), None);

        run.set_hyperlink_url("https://example.com");
        run.set_hyperlink_tooltip("Example site");

        assert_eq!(run.hyperlink_url(), Some("https://example.com"));
        assert_eq!(run.hyperlink_tooltip(), Some("Example site"));

        run.clear_hyperlink_url();
        run.clear_hyperlink_tooltip();
        assert_eq!(run.hyperlink_url(), None);
        assert_eq!(run.hyperlink_tooltip(), None);
    }

    #[test]
    fn character_spacing_roundtrip() {
        let mut run = TextRun::new("Spaced");
        assert_eq!(run.character_spacing(), None);

        run.set_character_spacing(200); // 2pt expanded
        assert_eq!(run.character_spacing(), Some(200));

        run.set_character_spacing(-100); // 1pt condensed
        assert_eq!(run.character_spacing(), Some(-100));

        run.clear_character_spacing();
        assert_eq!(run.character_spacing(), None);
    }

    #[test]
    fn baseline_subscript_superscript() {
        let mut run = TextRun::new("H2O");
        assert_eq!(run.baseline(), None);

        // Subscript
        run.set_baseline(-25000);
        assert_eq!(run.baseline(), Some(-25000));

        // Superscript
        run.set_baseline(30000);
        assert_eq!(run.baseline(), Some(30000));

        run.clear_baseline();
        assert_eq!(run.baseline(), None);
    }

    #[test]
    fn formatting_defaults_include_new_fields() {
        let run = TextRun::new("Test");
        assert_eq!(run.hyperlink_url(), None);
        assert_eq!(run.hyperlink_tooltip(), None);
        assert_eq!(run.character_spacing(), None);
        assert_eq!(run.kerning(), None);
        assert_eq!(run.baseline(), None);
    }

    #[test]
    fn kerning_roundtrip() {
        let mut run = TextRun::new("Kerned");
        assert_eq!(run.kerning(), None);

        run.set_kerning(1200); // 12pt threshold
        assert_eq!(run.kerning(), Some(1200));

        run.set_kerning(0); // disable
        assert_eq!(run.kerning(), Some(0));

        run.clear_kerning();
        assert_eq!(run.kerning(), None);
    }

    #[test]
    fn kerning_in_properties() {
        let mut props = RunProperties::default();
        assert_eq!(props.kerning, None);

        props.kerning = Some(2400);
        assert_eq!(props.kerning, Some(2400));
    }
}
