// ── Feature #14: Theme colors ──

/// Color scheme parsed from theme1.xml (`<a:clrScheme>`).
#[derive(Debug, Clone, PartialEq, Eq, Default)]
pub struct ThemeColorScheme {
    /// Dark 1 color as sRGB hex.
    pub dark1: Option<String>,
    /// Light 1 color as sRGB hex.
    pub light1: Option<String>,
    /// Dark 2 color as sRGB hex.
    pub dark2: Option<String>,
    /// Light 2 color as sRGB hex.
    pub light2: Option<String>,
    /// Accent 1 color as sRGB hex.
    pub accent1: Option<String>,
    /// Accent 2 color as sRGB hex.
    pub accent2: Option<String>,
    /// Accent 3 color as sRGB hex.
    pub accent3: Option<String>,
    /// Accent 4 color as sRGB hex.
    pub accent4: Option<String>,
    /// Accent 5 color as sRGB hex.
    pub accent5: Option<String>,
    /// Accent 6 color as sRGB hex.
    pub accent6: Option<String>,
    /// Hyperlink color as sRGB hex.
    pub hyperlink: Option<String>,
    /// Followed hyperlink color as sRGB hex.
    pub followed_hyperlink: Option<String>,
    /// Scheme name.
    pub name: Option<String>,
}

impl ThemeColorScheme {
    pub fn new() -> Self {
        Self::default()
    }

    /// Look up a color by its theme color name (dk1, lt1, dk2, lt2, accent1-6, hlink, folHlink).
    pub fn color_by_name(&self, name: &str) -> Option<&str> {
        match name {
            "dk1" => self.dark1.as_deref(),
            "lt1" => self.light1.as_deref(),
            "dk2" => self.dark2.as_deref(),
            "lt2" => self.light2.as_deref(),
            "accent1" => self.accent1.as_deref(),
            "accent2" => self.accent2.as_deref(),
            "accent3" => self.accent3.as_deref(),
            "accent4" => self.accent4.as_deref(),
            "accent5" => self.accent5.as_deref(),
            "accent6" => self.accent6.as_deref(),
            "hlink" => self.hyperlink.as_deref(),
            "folHlink" => self.followed_hyperlink.as_deref(),
            _ => None,
        }
    }

    /// Set a color by its theme color name.
    pub fn set_color_by_name(&mut self, name: &str, value: impl Into<String>) -> bool {
        let value = value.into();
        match name {
            "dk1" => self.dark1 = Some(value),
            "lt1" => self.light1 = Some(value),
            "dk2" => self.dark2 = Some(value),
            "lt2" => self.light2 = Some(value),
            "accent1" => self.accent1 = Some(value),
            "accent2" => self.accent2 = Some(value),
            "accent3" => self.accent3 = Some(value),
            "accent4" => self.accent4 = Some(value),
            "accent5" => self.accent5 = Some(value),
            "accent6" => self.accent6 = Some(value),
            "hlink" => self.hyperlink = Some(value),
            "folHlink" => self.followed_hyperlink = Some(value),
            _ => return false,
        }
        true
    }
}

/// Font scheme parsed from `a:fontScheme` in theme XML.
#[derive(Debug, Clone, PartialEq, Eq, Default)]
pub struct ThemeFontScheme {
    /// Major latin font (headings).
    pub major_latin: String,
    /// Minor latin font (body).
    pub minor_latin: String,
    /// Major east asian font.
    pub major_east_asian: Option<String>,
    /// Minor east asian font.
    pub minor_east_asian: Option<String>,
}

impl ThemeFontScheme {
    pub fn new(major_latin: impl Into<String>, minor_latin: impl Into<String>) -> Self {
        Self {
            major_latin: major_latin.into(),
            minor_latin: minor_latin.into(),
            major_east_asian: None,
            minor_east_asian: None,
        }
    }
}

/// Theme color reference for elements that use theme colors instead of sRGB.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum ThemeColorRef {
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

impl ThemeColorRef {
    pub fn from_xml(value: &str) -> Option<Self> {
        match value {
            "dk1" => Some(Self::Dark1),
            "lt1" => Some(Self::Light1),
            "dk2" => Some(Self::Dark2),
            "lt2" => Some(Self::Light2),
            "accent1" => Some(Self::Accent1),
            "accent2" => Some(Self::Accent2),
            "accent3" => Some(Self::Accent3),
            "accent4" => Some(Self::Accent4),
            "accent5" => Some(Self::Accent5),
            "accent6" => Some(Self::Accent6),
            "hlink" => Some(Self::Hyperlink),
            "folHlink" => Some(Self::FollowedHyperlink),
            _ => None,
        }
    }

    pub fn to_xml(&self) -> &'static str {
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

    /// Resolve to an sRGB hex color using a color scheme.
    pub fn resolve<'a>(&self, scheme: &'a ThemeColorScheme) -> Option<&'a str> {
        scheme.color_by_name(self.to_xml())
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn theme_color_scheme_lookup() {
        let mut scheme = ThemeColorScheme::new();
        scheme.set_color_by_name("dk1", "000000");
        scheme.set_color_by_name("lt1", "FFFFFF");
        scheme.set_color_by_name("accent1", "4472C4");

        assert_eq!(scheme.color_by_name("dk1"), Some("000000"));
        assert_eq!(scheme.color_by_name("lt1"), Some("FFFFFF"));
        assert_eq!(scheme.color_by_name("accent1"), Some("4472C4"));
        assert_eq!(scheme.color_by_name("nonexistent"), None);
    }

    #[test]
    fn theme_color_ref_resolve() {
        let mut scheme = ThemeColorScheme::new();
        scheme.accent1 = Some("4472C4".to_string());

        let color_ref = ThemeColorRef::from_xml("accent1").unwrap();
        assert_eq!(color_ref.resolve(&scheme), Some("4472C4"));
        assert_eq!(color_ref.to_xml(), "accent1");
    }

    #[test]
    fn theme_color_ref_xml_roundtrip() {
        for (xml, expected) in [
            ("dk1", ThemeColorRef::Dark1),
            ("lt1", ThemeColorRef::Light1),
            ("accent1", ThemeColorRef::Accent1),
            ("hlink", ThemeColorRef::Hyperlink),
        ] {
            let parsed = ThemeColorRef::from_xml(xml).unwrap();
            assert_eq!(parsed, expected);
            assert_eq!(parsed.to_xml(), xml);
        }
    }

    // ── Theme font scheme tests ──

    #[test]
    fn theme_font_scheme_new() {
        let scheme = ThemeFontScheme::new("Calibri Light", "Calibri");
        assert_eq!(scheme.major_latin, "Calibri Light");
        assert_eq!(scheme.minor_latin, "Calibri");
        assert!(scheme.major_east_asian.is_none());
        assert!(scheme.minor_east_asian.is_none());
    }

    #[test]
    fn theme_font_scheme_with_east_asian() {
        let mut scheme = ThemeFontScheme::new("Arial", "Verdana");
        scheme.major_east_asian = Some("MS Gothic".to_string());
        scheme.minor_east_asian = Some("MS Mincho".to_string());

        assert_eq!(scheme.major_east_asian.as_deref(), Some("MS Gothic"));
        assert_eq!(scheme.minor_east_asian.as_deref(), Some("MS Mincho"));
    }

    #[test]
    fn theme_font_scheme_default() {
        let scheme = ThemeFontScheme::default();
        assert!(scheme.major_latin.is_empty());
        assert!(scheme.minor_latin.is_empty());
    }
}
