//! Color parsing utilities for spreadsheet rendering.
//!
//! This module provides backend-agnostic color handling using CSS color strings,
//! which are directly usable by Canvas 2D and can be converted for other backends.

/// A CSS color string (e.g., "#FF0000", "rgba(255, 0, 0, 0.5)")
pub type CssColor = String;

/// RGB color with u8 components for efficient color manipulation.
///
/// This type avoids repeated string parsing when doing color math operations
/// like lightening, darkening, or calculating luminance.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Rgb {
    pub r: u8,
    pub g: u8,
    pub b: u8,
}

impl Rgb {
    /// Create a new RGB color.
    pub const fn new(r: u8, g: u8, b: u8) -> Self {
        Self { r, g, b }
    }

    /// Parse from a hex string (with or without #).
    /// Returns None if the format is invalid.
    pub fn from_hex(s: &str) -> Option<Self> {
        let hex = s.trim().strip_prefix('#').unwrap_or(s.trim());
        if hex.len() != 6 {
            return None;
        }
        let r = u8::from_str_radix(hex.get(0..2)?, 16).ok()?;
        let g = u8::from_str_radix(hex.get(2..4)?, 16).ok()?;
        let b = u8::from_str_radix(hex.get(4..6)?, 16).ok()?;
        Some(Self { r, g, b })
    }

    /// Convert to CSS hex string (#RRGGBB).
    pub fn to_hex(self) -> String {
        format!("#{:02X}{:02X}{:02X}", self.r, self.g, self.b)
    }

    /// Lighten the color by blending with white.
    /// Factor of 0.0 = no change, 1.0 = pure white.
    pub fn lighten(self, factor: f64) -> Self {
        Self {
            r: Self::blend_component(self.r, 255, factor),
            g: Self::blend_component(self.g, 255, factor),
            b: Self::blend_component(self.b, 255, factor),
        }
    }

    /// Darken the color by blending with black.
    /// Factor of 0.0 = no change, 1.0 = pure black.
    pub fn darken(self, factor: f64) -> Self {
        Self {
            r: Self::blend_component(self.r, 0, factor),
            g: Self::blend_component(self.g, 0, factor),
            b: Self::blend_component(self.b, 0, factor),
        }
    }

    /// Calculate relative luminance (0.0 to 1.0).
    /// Uses simplified formula: 0.299*R + 0.587*G + 0.114*B
    pub fn luminance(self) -> f64 {
        let r = f64::from(self.r);
        let g = f64::from(self.g);
        let b = f64::from(self.b);
        (0.299 * r + 0.587 * g + 0.114 * b) / 255.0
    }

    /// Check if this is a light color (luminance > 0.5).
    pub fn is_light(self) -> bool {
        self.luminance() > 0.5
    }

    /// Blend a single color component toward a target.
    /// The cast is safe because we clamp to [0, 255] before converting.
    #[allow(clippy::cast_possible_truncation, clippy::cast_sign_loss)]
    fn blend_component(from: u8, to: u8, factor: f64) -> u8 {
        let from = f64::from(from);
        let to = f64::from(to);
        let blended = from + (to - from) * factor.clamp(0.0, 1.0);
        blended.clamp(0.0, 255.0).round() as u8
    }
}

impl Default for Rgb {
    fn default() -> Self {
        Self::new(0, 0, 0)
    }
}

/// Parse a color string and normalize it to CSS format.
///
/// Supports formats:
/// - "#RRGGBB" (hex without alpha)
/// - "#AARRGGBB" (Excel format: alpha first)
/// - "RRGGBB" (hex without # prefix)
/// - "rgb(r, g, b)" (CSS-style)
/// - "rgba(r, g, b, a)" (CSS-style with alpha)
pub fn parse_color(s: &str) -> Option<CssColor> {
    let s = s.trim();

    if s.starts_with('#') {
        parse_hex_color(s)
    } else if s.starts_with("rgb") {
        // Already in CSS format, just validate and return
        Some(s.to_string())
    } else {
        // Try as plain hex without #
        parse_hex_color(&format!("#{}", s))
    }
}

fn parse_hex_color(s: &str) -> Option<CssColor> {
    let hex = s.strip_prefix('#')?;

    match hex.len() {
        6 => {
            // #RRGGBB - validate and return as-is
            let _r = u8::from_str_radix(hex.get(0..2)?, 16).ok()?;
            let _g = u8::from_str_radix(hex.get(2..4)?, 16).ok()?;
            let _b = u8::from_str_radix(hex.get(4..6)?, 16).ok()?;
            Some(format!("#{}", hex))
        }
        8 => {
            // #AARRGGBB (Excel format) - convert to rgba()
            let a = u8::from_str_radix(hex.get(0..2)?, 16).ok()?;
            let r = u8::from_str_radix(hex.get(2..4)?, 16).ok()?;
            let g = u8::from_str_radix(hex.get(4..6)?, 16).ok()?;
            let b = u8::from_str_radix(hex.get(6..8)?, 16).ok()?;

            if a == 255 {
                // Fully opaque, use simple hex
                Some(format!("#{:02X}{:02X}{:02X}", r, g, b))
            } else {
                // Has transparency, use rgba
                let alpha = f64::from(a) / 255.0;
                Some(format!("rgba({}, {}, {}, {:.2})", r, g, b, alpha))
            }
        }
        _ => None,
    }
}

/// Parse color and return RGBA components (0-255 for RGB, 0.0-1.0 for alpha)
pub fn parse_color_rgba(s: &str) -> Option<(u8, u8, u8, f64)> {
    let s = s.trim();

    if s.starts_with('#') {
        parse_hex_rgba(s)
    } else if s.starts_with("rgba(") {
        parse_rgba_string(s)
    } else if s.starts_with("rgb(") {
        parse_rgb_string(s)
    } else {
        // Try as plain hex
        parse_hex_rgba(&format!("#{}", s))
    }
}

fn parse_hex_rgba(s: &str) -> Option<(u8, u8, u8, f64)> {
    let hex = s.strip_prefix('#')?;

    match hex.len() {
        6 => {
            let r = u8::from_str_radix(hex.get(0..2)?, 16).ok()?;
            let g = u8::from_str_radix(hex.get(2..4)?, 16).ok()?;
            let b = u8::from_str_radix(hex.get(4..6)?, 16).ok()?;
            Some((r, g, b, 1.0))
        }
        8 => {
            // Excel AARRGGBB format
            let a = u8::from_str_radix(hex.get(0..2)?, 16).ok()?;
            let r = u8::from_str_radix(hex.get(2..4)?, 16).ok()?;
            let g = u8::from_str_radix(hex.get(4..6)?, 16).ok()?;
            let b = u8::from_str_radix(hex.get(6..8)?, 16).ok()?;
            Some((r, g, b, f64::from(a) / 255.0))
        }
        _ => None,
    }
}

fn parse_rgb_string(s: &str) -> Option<(u8, u8, u8, f64)> {
    let inner = s.strip_prefix("rgb(")?.strip_suffix(')')?;
    let mut parts = inner.split(',').map(|p| p.trim());
    let r: u8 = parts.next()?.parse().ok()?;
    let g: u8 = parts.next()?.parse().ok()?;
    let b: u8 = parts.next()?.parse().ok()?;
    if parts.next().is_some() {
        return None;
    }
    Some((r, g, b, 1.0))
}

fn parse_rgba_string(s: &str) -> Option<(u8, u8, u8, f64)> {
    let inner = s.strip_prefix("rgba(")?.strip_suffix(')')?;
    let mut parts = inner.split(',').map(|p| p.trim());
    let r: u8 = parts.next()?.parse().ok()?;
    let g: u8 = parts.next()?.parse().ok()?;
    let b: u8 = parts.next()?.parse().ok()?;
    let a: f64 = parts.next()?.parse().ok()?;
    if parts.next().is_some() {
        return None;
    }
    Some((r, g, b, a))
}

/// Common colors used in spreadsheet rendering (CSS format)
pub mod palette {
    pub const WHITE: &str = "#FFFFFF";
    pub const BLACK: &str = "#000000";

    /// Grid line color (light gray)
    pub const GRID_LINE: &str = "#E0E0E0";

    // Tab bar colors - Google Sheets inspired
    /// Tab bar background (dark gray)
    pub const TAB_BG: &str = "#F1F3F4";

    /// Active tab background (white)
    pub const TAB_ACTIVE: &str = "#FFFFFF";

    /// Inactive tab background (slightly transparent)
    pub const TAB_INACTIVE: &str = "#E8EAED";

    /// Tab hover background
    pub const TAB_HOVER: &str = "#DEE1E6";

    /// Tab border color
    pub const TAB_BORDER: &str = "#DADCE0";

    /// Tab text color (dark gray)
    pub const TAB_TEXT: &str = "#3C4043";

    /// Tab text color for active tab
    pub const TAB_TEXT_ACTIVE: &str = "#202124";

    /// Scroll button background
    pub const TAB_SCROLL_BG: &str = "#E8EAED";

    /// Scroll button hover background
    pub const TAB_SCROLL_HOVER: &str = "#DADCE0";

    /// Scroll button icon color
    pub const TAB_SCROLL_ICON: &str = "#5F6368";

    /// Scrollbar track color
    pub const SCROLLBAR_TRACK: &str = "#F5F5F5";

    /// Scrollbar thumb color
    pub const SCROLLBAR_THUMB: &str = "#B4B4B4";
}

#[cfg(test)]
#[allow(
    clippy::unwrap_used,
    clippy::expect_used,
    clippy::indexing_slicing,
    clippy::float_cmp,
    clippy::panic
)]
mod tests {
    use super::*;

    #[test]
    fn test_parse_hex_6() {
        let color = parse_color("#FF0000").unwrap();
        assert_eq!(color, "#FF0000");
    }

    #[test]
    fn test_parse_hex_8_opaque() {
        let color = parse_color("#FFFF0000").unwrap();
        // Fully opaque (FF alpha) should return simple hex
        assert_eq!(color, "#FF0000");
    }

    #[test]
    fn test_parse_hex_8_transparent() {
        let color = parse_color("#80FF0000").unwrap();
        // 50% alpha should return rgba
        assert!(color.starts_with("rgba(255, 0, 0,"));
    }

    #[test]
    fn test_parse_rgb() {
        let color = parse_color("rgb(255, 128, 64)").unwrap();
        assert_eq!(color, "rgb(255, 128, 64)");
    }

    #[test]
    fn test_parse_rgba() {
        let color = parse_color("rgba(255, 128, 64, 0.5)").unwrap();
        assert_eq!(color, "rgba(255, 128, 64, 0.5)");
    }

    #[test]
    fn test_parse_color_rgba() {
        let (r, g, b, a) = parse_color_rgba("#FF8040").unwrap();
        assert_eq!((r, g, b), (255, 128, 64));
        assert!((a - 1.0).abs() < 0.001);
    }

    #[test]
    fn test_parse_without_hash() {
        let color = parse_color("FF0000").unwrap();
        assert_eq!(color, "#FF0000");
    }
}
