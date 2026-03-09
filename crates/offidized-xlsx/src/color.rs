//! Color resolution utilities for Excel OOXML files.
//!
//! Handles theme colors, indexed colors, RGB, and tint/shade calculations.
//! This module provides the resolution logic to convert abstract [`ColorReference`]
//! values into concrete `#RRGGBB` hex strings.

use crate::style::ColorReference;

/// Excel's 64 indexed colors (legacy palette).
///
/// This is the default indexed color palette defined by the OOXML specification.
/// Workbooks may override individual entries via the `<colors><indexedColors>` element.
pub const INDEXED_COLORS: [&str; 64] = [
    "#000000", "#FFFFFF", "#FF0000", "#00FF00", "#0000FF", "#FFFF00", "#FF00FF", "#00FFFF",
    "#000000", "#FFFFFF", "#FF0000", "#00FF00", "#0000FF", "#FFFF00", "#FF00FF", "#00FFFF",
    "#800000", "#008000", "#000080", "#808000", "#800080", "#008080", "#C0C0C0", "#808080",
    "#9999FF", "#993366", "#FFFFCC", "#CCFFFF", "#660066", "#FF8080", "#0066CC", "#CCCCFF",
    "#000080", "#FF00FF", "#FFFF00", "#00FFFF", "#800080", "#800000", "#008080", "#0000FF",
    "#00CCFF", "#CCFFFF", "#CCFFCC", "#FFFF99", "#99CCFF", "#FF99CC", "#CC99FF", "#FFCC99",
    "#3366FF", "#33CCCC", "#99CC00", "#FFCC00", "#FF9900", "#FF6600", "#666699", "#969696",
    "#003366", "#339966", "#003300", "#333300", "#993300", "#993366", "#333399", "#333333",
];

/// Default theme colors (Office theme) used when no theme is present.
///
/// Indices correspond to [`ThemeColor::index()`](crate::ThemeColor::index) values
/// (per ECMA-376, matching the XML `<a:clrScheme>` element order):
/// - 0: dk1 (Dark 1 / Text 1) - typically black
/// - 1: lt1 (Light 1 / Background 1) - typically white
/// - 2: dk2 (Dark 2 / Text 2)
/// - 3: lt2 (Light 2 / Background 2)
/// - 4-9: accent1-accent6
/// - 10: hlink (hyperlink)
/// - 11: folHlink (followed hyperlink)
pub const DEFAULT_THEME_COLORS: [&str; 12] = [
    "#000000", // 0: dk1 (Dark 1 - typically black)
    "#FFFFFF", // 1: lt1 (Light 1 - typically white)
    "#44546A", // 2: dk2 (Dark 2)
    "#E7E6E6", // 3: lt2 (Light 2)
    "#4472C4", // 4: accent1
    "#ED7D31", // 5: accent2
    "#A5A5A5", // 6: accent3
    "#FFC000", // 7: accent4
    "#5B9BD5", // 8: accent5
    "#70AD47", // 9: accent6
    "#0563C1", // 10: hlink
    "#954F72", // 11: folHlink
];

/// Resolve a [`ColorReference`] to an `#RRGGBB` hex string.
///
/// The resolution priority follows the OOXML specification:
/// 1. `rgb` (explicit RGB or ARGB value)
/// 2. `theme` (theme color index, optionally modified by `tint`)
/// 3. `indexed` (legacy indexed color palette)
/// 4. `auto` (automatic/system color, defaults to black)
///
/// # Arguments
/// - `color` - The color reference to resolve.
/// - `theme_colors` - The workbook's theme color palette (up to 12 entries).
///   Falls back to [`DEFAULT_THEME_COLORS`] for missing indices.
/// - `indexed_colors` - Optional custom indexed color palette. Falls back to
///   [`INDEXED_COLORS`] if `None` or if the index is out of range.
///
/// Returns `None` if the color reference has no resolvable color information.
pub fn resolve_color(
    color: &ColorReference,
    theme_colors: &[String],
    indexed_colors: Option<&[String]>,
) -> Option<String> {
    // Priority: rgb > theme > indexed > auto

    if let Some(rgb) = color.rgb() {
        // Excel sometimes uses ARGB (8 chars), we want RGB (6 chars)
        let rgb = rgb.trim_start_matches('#');
        if rgb.len() == 8 {
            return Some(format!("#{}", &rgb[2..]));
        }
        return Some(format!("#{rgb}"));
    }

    if let Some(theme) = color.theme() {
        let idx = theme.index() as usize;
        let base_color = theme_colors
            .get(idx)
            .map(String::as_str)
            .or_else(|| DEFAULT_THEME_COLORS.get(idx).copied())?;

        // Apply tint if present
        if let Some(tint) = color.tint() {
            return Some(apply_tint(base_color, tint));
        }
        return Some(base_color.to_string());
    }

    if let Some(indexed) = color.indexed() {
        if indexed == 64 {
            // 64 is "system foreground" - usually black
            return Some("#000000".to_string());
        }

        let idx = indexed as usize;

        // Use custom palette if available, otherwise fall back to default
        if let Some(custom_palette) = indexed_colors {
            if let Some(c) = custom_palette.get(idx) {
                return Some(c.clone());
            }
        }

        // Fall back to default palette
        if let Some(c) = INDEXED_COLORS.get(idx) {
            return Some((*c).to_string());
        }
    }

    if color.auto() == Some(true) {
        // Auto color - default to black for text, white for background
        return Some("#000000".to_string());
    }

    None
}

/// Apply a tint value to a color, returning the modified `#RRGGBB` string.
///
/// Per the OOXML specification (ECMA-376 Part 1, Section 18.8.19):
/// - `tint < 0`: shade (darken). The luminance is multiplied by `(1 + tint)`.
/// - `tint > 0`: tint (lighten). The luminance is moved toward 1.0 by `tint`.
///
/// The algorithm converts RGB to HSL, adjusts the lightness component, then
/// converts back to RGB.
#[allow(clippy::many_single_char_names)]
pub fn apply_tint(hex_color: &str, tint: f64) -> String {
    let hex = hex_color.trim_start_matches('#');

    let r = u8::from_str_radix(hex.get(0..2).unwrap_or("00"), 16).unwrap_or(0);
    let g = u8::from_str_radix(hex.get(2..4).unwrap_or("00"), 16).unwrap_or(0);
    let b = u8::from_str_radix(hex.get(4..6).unwrap_or("00"), 16).unwrap_or(0);

    let (h, s, l) = rgb_to_hsl(r, g, b);

    let new_l = if tint < 0.0 {
        // Shade: darken
        l * (1.0 + tint)
    } else {
        // Tint: lighten
        (1.0 - l).mul_add(tint, l)
    };

    let (r, g, b) = hsl_to_rgb(h, s, new_l.clamp(0.0, 1.0));

    format!("#{r:02X}{g:02X}{b:02X}")
}

/// Convert RGB (0-255 per channel) to HSL (all components 0.0-1.0).
#[allow(clippy::many_single_char_names)]
fn rgb_to_hsl(r: u8, g: u8, b: u8) -> (f64, f64, f64) {
    let r = f64::from(r) / 255.0;
    let g = f64::from(g) / 255.0;
    let b = f64::from(b) / 255.0;

    let max = r.max(g).max(b);
    let min = r.min(g).min(b);
    let l = f64::midpoint(max, min);

    if (max - min).abs() < f64::EPSILON {
        return (0.0, 0.0, l);
    }

    let d = max - min;
    let s = if l > 0.5 {
        d / (2.0 - max - min)
    } else {
        d / (max + min)
    };

    let h = if (max - r).abs() < f64::EPSILON {
        (g - b) / d + if g < b { 6.0 } else { 0.0 }
    } else if (max - g).abs() < f64::EPSILON {
        (b - r) / d + 2.0
    } else {
        (r - g) / d + 4.0
    };

    (h / 6.0, s, l)
}

/// Convert HSL (all components 0.0-1.0) to RGB (0-255 per channel).
#[allow(clippy::many_single_char_names)]
#[allow(clippy::cast_possible_truncation)]
#[allow(clippy::cast_sign_loss)]
fn hsl_to_rgb(h: f64, s: f64, l: f64) -> (u8, u8, u8) {
    if s.abs() < f64::EPSILON {
        let v = (l * 255.0).round() as u8;
        return (v, v, v);
    }

    let q = if l < 0.5 {
        l * (1.0 + s)
    } else {
        l.mul_add(-s, l + s)
    };
    let p = 2.0f64.mul_add(l, -q);

    let r = hue_to_rgb(p, q, h + 1.0 / 3.0);
    let g = hue_to_rgb(p, q, h);
    let b = hue_to_rgb(p, q, h - 1.0 / 3.0);

    (
        (r * 255.0).round() as u8,
        (g * 255.0).round() as u8,
        (b * 255.0).round() as u8,
    )
}

/// Helper for HSL-to-RGB conversion: compute a single channel from hue.
fn hue_to_rgb(p: f64, q: f64, mut t: f64) -> f64 {
    if t < 0.0 {
        t += 1.0;
    }
    if t > 1.0 {
        t -= 1.0;
    }

    if t < 1.0 / 6.0 {
        return ((q - p) * 6.0).mul_add(t, p);
    }
    if t < 1.0 / 2.0 {
        return q;
    }
    if t < 2.0 / 3.0 {
        return ((q - p) * (2.0 / 3.0 - t)).mul_add(6.0, p);
    }
    p
}

// =============================================================================
// Tests
// =============================================================================

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
    fn test_tint_lighten() {
        // 50% tint on black should give gray
        let result = apply_tint("#000000", 0.5);
        assert_eq!(result, "#808080");
    }

    #[test]
    fn test_tint_darken() {
        // 50% shade on white should give gray
        let result = apply_tint("#FFFFFF", -0.5);
        assert_eq!(result, "#808080");
    }
}

// =============================================================================
// Fill color resolution tests
// =============================================================================

#[cfg(test)]
#[allow(
    clippy::unwrap_used,
    clippy::expect_used,
    clippy::indexing_slicing,
    clippy::float_cmp,
    clippy::excessive_precision,
    clippy::panic
)]
mod fill_color_tests {
    use super::*;
    use crate::style::ThemeColor;

    fn default_theme_colors() -> Vec<String> {
        vec![
            "#000000".to_string(), // 0: dk1 (Dark 1)
            "#FFFFFF".to_string(), // 1: lt1 (Light 1)
            "#44546A".to_string(), // 2: dk2 (Dark 2)
            "#E7E6E6".to_string(), // 3: lt2 (Light 2)
            "#4472C4".to_string(), // 4: accent1
            "#ED7D31".to_string(), // 5: accent2
            "#A5A5A5".to_string(), // 6: accent3
            "#FFC000".to_string(), // 7: accent4
            "#5B9BD5".to_string(), // 8: accent5
            "#70AD47".to_string(), // 9: accent6
            "#0563C1".to_string(), // 10: hlink
            "#954F72".to_string(), // 11: folHlink
        ]
    }

    // =========================================================================
    // RGB color resolution tests
    // =========================================================================

    #[test]
    fn test_resolve_rgb_color_argb_format() {
        // ARGB format: first 2 chars are alpha (FF = opaque)
        let color = ColorReference::from_rgb("FFFFFF00"); // Yellow with alpha

        let resolved = resolve_color(&color, &default_theme_colors(), None);
        assert_eq!(resolved, Some("#FFFF00".to_string())); // Alpha stripped
    }

    #[test]
    fn test_resolve_rgb_color_rgb_format() {
        // Some files use 6-char RGB without alpha
        let color = ColorReference::from_rgb("FF0000"); // Red

        let resolved = resolve_color(&color, &default_theme_colors(), None);
        assert_eq!(resolved, Some("#FF0000".to_string()));
    }

    #[test]
    fn test_resolve_rgb_color_with_hash() {
        // Handle color with leading hash
        let color = ColorReference::from_rgb("#00FF00"); // Green with hash

        let resolved = resolve_color(&color, &default_theme_colors(), None);
        assert_eq!(resolved, Some("#00FF00".to_string()));
    }

    #[test]
    fn test_resolve_rgb_common_fill_colors() {
        // Test common fill colors
        let test_cases = [
            ("FFFF0000", "#FF0000"), // Red
            ("FF00FF00", "#00FF00"), // Green
            ("FF0000FF", "#0000FF"), // Blue
            ("FFFFFF00", "#FFFF00"), // Yellow
            ("FFFF00FF", "#FF00FF"), // Magenta
            ("FF00FFFF", "#00FFFF"), // Cyan
            ("FFFFFFFF", "#FFFFFF"), // White
            ("FF000000", "#000000"), // Black
            ("FFC0C0C0", "#C0C0C0"), // Silver
            ("FF808080", "#808080"), // Gray
        ];

        for (input, expected) in test_cases {
            let color = ColorReference::from_rgb(input);

            let resolved = resolve_color(&color, &default_theme_colors(), None);
            assert_eq!(
                resolved,
                Some(expected.to_string()),
                "Failed for input: {input}"
            );
        }
    }

    // =========================================================================
    // Theme color resolution tests
    // =========================================================================

    #[test]
    fn test_resolve_theme_color_dark1() {
        // Dark1 (index 0) = dk1 (Text 1 - usually black)
        let color = ColorReference::from_theme(ThemeColor::Dark1);

        let resolved = resolve_color(&color, &default_theme_colors(), None);
        assert_eq!(resolved, Some("#000000".to_string()));
    }

    #[test]
    fn test_resolve_theme_color_light1() {
        // Light1 (index 1) = lt1 (Background 1 - usually white)
        let color = ColorReference::from_theme(ThemeColor::Light1);

        let resolved = resolve_color(&color, &default_theme_colors(), None);
        assert_eq!(resolved, Some("#FFFFFF".to_string()));
    }

    #[test]
    fn test_resolve_theme_color_accent1() {
        // theme="4" is accent1 (blue in default Office theme)
        let color = ColorReference::from_theme(ThemeColor::Accent1);

        let resolved = resolve_color(&color, &default_theme_colors(), None);
        assert_eq!(resolved, Some("#4472C4".to_string()));
    }

    #[test]
    fn test_resolve_all_theme_colors() {
        let theme_colors = default_theme_colors();
        let expected: [(ThemeColor, &str); 12] = [
            (ThemeColor::Dark1, "#000000"),             // 0: dk1 (Dark 1)
            (ThemeColor::Light1, "#FFFFFF"),            // 1: lt1 (Light 1)
            (ThemeColor::Dark2, "#44546A"),             // 2: dk2 (Dark 2)
            (ThemeColor::Light2, "#E7E6E6"),            // 3: lt2 (Light 2)
            (ThemeColor::Accent1, "#4472C4"),           // 4: accent1
            (ThemeColor::Accent2, "#ED7D31"),           // 5: accent2
            (ThemeColor::Accent3, "#A5A5A5"),           // 6: accent3
            (ThemeColor::Accent4, "#FFC000"),           // 7: accent4
            (ThemeColor::Accent5, "#5B9BD5"),           // 8: accent5
            (ThemeColor::Accent6, "#70AD47"),           // 9: accent6
            (ThemeColor::Hyperlink, "#0563C1"),         // 10: hlink
            (ThemeColor::FollowedHyperlink, "#954F72"), // 11: folHlink
        ];

        for (theme, expected_color) in expected {
            let color = ColorReference::from_theme(theme);

            let resolved = resolve_color(&color, &theme_colors, None);
            assert_eq!(
                resolved,
                Some(expected_color.to_string()),
                "Failed for theme: {theme:?} (index {})",
                theme.index()
            );
        }
    }

    // =========================================================================
    // Theme color with tint tests
    // =========================================================================

    #[test]
    fn test_resolve_theme_color_with_positive_tint() {
        // Positive tint lightens the color
        let color = ColorReference::from_theme_with_tint(ThemeColor::Accent1, 0.5);

        let resolved = resolve_color(&color, &default_theme_colors(), None);
        assert!(resolved.is_some());
        let hex = resolved.expect("resolved color");

        // Lightened color should have higher RGB values
        // Original: #4472C4 (R=68, G=114, B=196)
        // After 50% tint should be lighter
        assert!(hex.starts_with('#'));
        assert_eq!(hex.len(), 7);
    }

    #[test]
    fn test_resolve_theme_color_with_negative_tint() {
        // Negative tint darkens the color (shade)
        let color = ColorReference::from_theme_with_tint(ThemeColor::Accent1, -0.5);

        let resolved = resolve_color(&color, &default_theme_colors(), None);
        assert!(resolved.is_some());
        let hex = resolved.expect("resolved color");

        // Darkened color should have lower RGB values
        assert!(hex.starts_with('#'));
        assert_eq!(hex.len(), 7);
    }

    #[test]
    fn test_resolve_theme_color_with_small_tint() {
        // Small tint like Excel often uses
        let color = ColorReference::from_theme_with_tint(ThemeColor::Accent1, 0.39997558519241921);

        let resolved = resolve_color(&color, &default_theme_colors(), None);
        assert!(resolved.is_some());
    }

    #[test]
    fn test_resolve_theme_color_with_small_shade() {
        // Small shade like Excel often uses
        let color = ColorReference::from_theme_with_tint(ThemeColor::Accent1, -0.249977111117893);

        let resolved = resolve_color(&color, &default_theme_colors(), None);
        assert!(resolved.is_some());
    }

    #[test]
    fn test_tint_on_black_produces_gray() {
        // 50% tint on Dark1 (#000000, black) should give 50% gray
        let color = ColorReference::from_theme_with_tint(ThemeColor::Dark1, 0.5);

        let resolved = resolve_color(&color, &default_theme_colors(), None);
        assert_eq!(resolved, Some("#808080".to_string()));
    }

    #[test]
    fn test_shade_on_white_produces_gray() {
        // 50% shade on Light1 (#FFFFFF, white) should give 50% gray
        let color = ColorReference::from_theme_with_tint(ThemeColor::Light1, -0.5);

        let resolved = resolve_color(&color, &default_theme_colors(), None);
        assert_eq!(resolved, Some("#808080".to_string()));
    }

    // =========================================================================
    // Indexed color resolution tests
    // =========================================================================

    #[test]
    fn test_resolve_indexed_color_black() {
        // indexed="0" is black
        let color = ColorReference::from_indexed(0);

        let resolved = resolve_color(&color, &default_theme_colors(), None);
        assert_eq!(resolved, Some("#000000".to_string()));
    }

    #[test]
    fn test_resolve_indexed_color_white() {
        // indexed="1" is white
        let color = ColorReference::from_indexed(1);

        let resolved = resolve_color(&color, &default_theme_colors(), None);
        assert_eq!(resolved, Some("#FFFFFF".to_string()));
    }

    #[test]
    fn test_resolve_indexed_color_red() {
        // indexed="2" is red
        let color = ColorReference::from_indexed(2);

        let resolved = resolve_color(&color, &default_theme_colors(), None);
        assert_eq!(resolved, Some("#FF0000".to_string()));
    }

    #[test]
    fn test_resolve_indexed_color_yellow() {
        // indexed="5" is yellow
        let color = ColorReference::from_indexed(5);

        let resolved = resolve_color(&color, &default_theme_colors(), None);
        assert_eq!(resolved, Some("#FFFF00".to_string()));
    }

    #[test]
    fn test_resolve_indexed_color_8() {
        // indexed="8" is black (second occurrence in palette)
        let color = ColorReference::from_indexed(8);

        let resolved = resolve_color(&color, &default_theme_colors(), None);
        assert_eq!(resolved, Some("#000000".to_string()));
    }

    #[test]
    fn test_resolve_indexed_color_64_system_foreground() {
        // indexed="64" is special "system foreground" color
        let color = ColorReference::from_indexed(64);

        let resolved = resolve_color(&color, &default_theme_colors(), None);
        assert_eq!(resolved, Some("#000000".to_string()));
    }

    #[test]
    fn test_resolve_indexed_colors_common() {
        // Test common indexed colors used in fills
        let test_cases: [(u32, &str); 8] = [
            (0, "#000000"),  // Black
            (1, "#FFFFFF"),  // White
            (2, "#FF0000"),  // Red
            (3, "#00FF00"),  // Green
            (4, "#0000FF"),  // Blue
            (5, "#FFFF00"),  // Yellow
            (22, "#C0C0C0"), // Silver
            (23, "#808080"), // Gray
        ];

        for (indexed, expected) in test_cases {
            let color = ColorReference::from_indexed(indexed);

            let resolved = resolve_color(&color, &default_theme_colors(), None);
            assert_eq!(
                resolved,
                Some(expected.to_string()),
                "Failed for indexed: {indexed}"
            );
        }
    }

    // =========================================================================
    // Auto color tests
    // =========================================================================

    #[test]
    fn test_resolve_auto_color() {
        // auto="1" means use automatic color (defaults to black)
        let mut color = ColorReference::empty();
        color.set_auto(true);

        let resolved = resolve_color(&color, &default_theme_colors(), None);
        assert_eq!(resolved, Some("#000000".to_string()));
    }

    // =========================================================================
    // Color priority tests
    // =========================================================================

    #[test]
    fn test_color_priority_rgb_over_theme() {
        // RGB should take priority over theme if both are specified
        let mut color = ColorReference::from_rgb("FF0000");
        color.set_theme(ThemeColor::Accent1);

        let resolved = resolve_color(&color, &default_theme_colors(), None);
        assert_eq!(resolved, Some("#FF0000".to_string())); // RGB wins
    }

    #[test]
    fn test_color_priority_rgb_over_indexed() {
        // RGB should take priority over indexed
        let mut color = ColorReference::from_rgb("00FF00");
        color.set_indexed(2); // Red

        let resolved = resolve_color(&color, &default_theme_colors(), None);
        assert_eq!(resolved, Some("#00FF00".to_string())); // RGB wins
    }

    #[test]
    fn test_color_priority_theme_over_indexed() {
        // Theme should take priority over indexed
        let mut color = ColorReference::from_theme(ThemeColor::Accent1); // #4472C4
        color.set_indexed(2); // Red

        let resolved = resolve_color(&color, &default_theme_colors(), None);
        assert_eq!(resolved, Some("#4472C4".to_string())); // Theme wins
    }

    #[test]
    fn test_color_priority_indexed_over_auto() {
        // Indexed should take priority over auto
        let mut color = ColorReference::from_indexed(2); // Red
        color.set_auto(true);

        let resolved = resolve_color(&color, &default_theme_colors(), None);
        assert_eq!(resolved, Some("#FF0000".to_string())); // Indexed wins
    }

    // =========================================================================
    // Edge case tests
    // =========================================================================

    #[test]
    fn test_empty_color_reference() {
        // No color specified should return None
        let color = ColorReference::empty();

        let resolved = resolve_color(&color, &default_theme_colors(), None);
        assert!(resolved.is_none());
    }

    #[test]
    fn test_empty_theme_colors_uses_defaults() {
        // When theme_colors is empty, should use DEFAULT_THEME_COLORS
        let color = ColorReference::from_theme(ThemeColor::Accent1); // accent1

        let resolved = resolve_color(&color, &[], None);
        // Should fall back to DEFAULT_THEME_COLORS
        assert_eq!(resolved, Some("#4472C4".to_string()));
    }

    #[test]
    fn test_invalid_indexed_color() {
        // Indexed color beyond 64-color palette (but not 64)
        let color = ColorReference::from_indexed(100); // Invalid

        let resolved = resolve_color(&color, &default_theme_colors(), None);
        assert!(resolved.is_none());
    }

    // =========================================================================
    // Custom indexed color tests
    // =========================================================================

    #[test]
    fn test_custom_indexed_colors() {
        // Test with default palette (None)
        let color = ColorReference::from_indexed(2); // Red in default palette

        let resolved = resolve_color(&color, &[], None);
        assert_eq!(resolved, Some("#FF0000".to_string()));

        // Test with custom palette
        let custom_palette = vec![
            "#000000".to_string(),
            "#FFFFFF".to_string(),
            "#00FF00".to_string(), // Custom green instead of red at index 2
        ];

        let resolved_custom = resolve_color(&color, &[], Some(&custom_palette));
        assert_eq!(resolved_custom, Some("#00FF00".to_string()));
    }

    #[test]
    fn test_fallback_to_default_when_custom_too_short() {
        let color = ColorReference::from_indexed(10); // Index beyond custom palette

        // Custom palette with only 3 colors
        let custom_palette = vec![
            "#000000".to_string(),
            "#FFFFFF".to_string(),
            "#00FF00".to_string(),
        ];

        // Should fall back to default palette
        let resolved = resolve_color(&color, &[], Some(&custom_palette));
        assert!(resolved.is_some());
    }

    // =========================================================================
    // ColorReference::resolve method tests
    // =========================================================================

    #[test]
    fn test_color_reference_resolve_method() {
        let theme_colors = default_theme_colors();

        let color = ColorReference::from_rgb("FF0000");
        assert_eq!(
            color.resolve(&theme_colors, None),
            Some("#FF0000".to_string())
        );

        let color = ColorReference::from_theme(ThemeColor::Accent1);
        assert_eq!(
            color.resolve(&theme_colors, None),
            Some("#4472C4".to_string())
        );

        let color = ColorReference::empty();
        assert_eq!(color.resolve(&theme_colors, None), None);
    }
}
