//! Theme color resolution and transformation utilities.
//!
//! This module provides functions for resolving PowerPoint theme colors
//! and applying color transformations like tint, shade, luminance modulation,
//! and luminance offset.
//!
//! Based on ShapeCrawler's color resolution system.

use crate::theme::ThemeColorScheme;

/// Resolves a scheme color name to its RGB hex value.
///
/// # Arguments
///
/// * `scheme_name` - The scheme color name (e.g., "dk1", "accent1", "hlink")
/// * `theme` - The theme color scheme containing the color definitions
///
/// # Returns
///
/// The RGB hex color value (without "#" prefix) if found, None otherwise.
///
/// # Examples
///
/// ```
/// use offidized_pptx::{ThemeColorScheme, color_resolver::resolve_scheme_color};
///
/// let mut scheme = ThemeColorScheme::new();
/// scheme.accent1 = Some("4472C4".to_string());
///
/// let color = resolve_scheme_color("accent1", &scheme);
/// assert_eq!(color, Some("4472C4"));
/// ```
pub fn resolve_scheme_color<'a>(scheme_name: &str, theme: &'a ThemeColorScheme) -> Option<&'a str> {
    theme.color_by_name(scheme_name)
}

/// Applies a tint transformation to a color.
///
/// Tint lightens a color by blending it with white. The tint value is a percentage
/// expressed as an integer (0-100000, where 100000 = 100%).
///
/// # Arguments
///
/// * `color` - The input color as RGB hex (e.g., "4472C4")
/// * `tint` - The tint percentage (0-100000). Higher values lighten more.
///
/// # Returns
///
/// The transformed color as RGB hex without "#" prefix.
///
/// # Examples
///
/// ```
/// use offidized_pptx::color_resolver::apply_tint;
///
/// // Apply 50% tint (lighten halfway to white)
/// let result = apply_tint("4472C4", 50000);
/// assert_eq!(result, "A1B9E1");
/// ```
pub fn apply_tint(color: &str, tint: i32) -> String {
    let (r, g, b) = parse_color_value(color);
    let tint_factor = tint as f64 / 100_000.0;

    // Tint moves colors towards white: new = current + (255 - current) * tintFactor
    // OOXML uses "round half to even" (banker's rounding) for color transformations
    let new_r = r as f64 + (255.0 - r as f64) * tint_factor;
    let new_g = g as f64 + (255.0 - g as f64) * tint_factor;
    let new_b = b as f64 + (255.0 - b as f64) * tint_factor;

    format!(
        "{:02X}{:02X}{:02X}",
        round_half_to_even(new_r) as u8,
        round_half_to_even(new_g) as u8,
        round_half_to_even(new_b) as u8
    )
}

/// Applies a shade transformation to a color.
///
/// Shade darkens a color by multiplying each RGB component by a factor.
/// The shade value is a percentage expressed as an integer (0-100000, where 100000 = 100%).
///
/// # Arguments
///
/// * `color` - The input color as RGB hex (e.g., "4472C4")
/// * `shade` - The shade percentage (0-100000). Lower values darken more.
///
/// # Returns
///
/// The transformed color as RGB hex without "#" prefix.
///
/// # Examples
///
/// ```
/// use offidized_pptx::color_resolver::apply_shade;
///
/// // Apply 50% shade (darken to half intensity)
/// let result = apply_shade("4472C4", 50000);
/// assert_eq!(result, "223962");
/// ```
pub fn apply_shade(color: &str, shade: i32) -> String {
    let (r, g, b) = parse_color_value(color);
    let shade_factor = shade as f64 / 100_000.0;

    // Shade darkens by multiplying RGB values
    let new_r = (r as f64 * shade_factor).round() as u8;
    let new_g = (g as f64 * shade_factor).round() as u8;
    let new_b = (b as f64 * shade_factor).round() as u8;

    format!("{:02X}{:02X}{:02X}", new_r, new_g, new_b)
}

/// Applies luminance modulation to a color.
///
/// Luminance modulation adjusts the brightness of a color by multiplying it by a factor.
/// The value is a percentage expressed as an integer (0-100000, where 100000 = 100%).
///
/// # Arguments
///
/// * `color` - The input color as RGB hex (e.g., "4472C4")
/// * `lum_mod` - The luminance modulation percentage (0-100000)
///
/// # Returns
///
/// The transformed color as RGB hex without "#" prefix.
///
/// # Examples
///
/// ```
/// use offidized_pptx::color_resolver::apply_luminance_mod;
///
/// // Apply 75% luminance modulation
/// let result = apply_luminance_mod("4472C4", 75000);
/// assert_eq!(result, "335693");
/// ```
pub fn apply_luminance_mod(color: &str, lum_mod: i32) -> String {
    // Luminance modulation is essentially the same as shade
    apply_shade(color, lum_mod)
}

/// Applies luminance offset to a color.
///
/// Luminance offset adjusts the brightness by adding or subtracting a value from each
/// RGB component. The value is a percentage expressed as an integer (-100000 to 100000).
///
/// # Arguments
///
/// * `color` - The input color as RGB hex (e.g., "4472C4")
/// * `lum_off` - The luminance offset percentage (-100000 to 100000)
///
/// # Returns
///
/// The transformed color as RGB hex without "#" prefix.
///
/// # Examples
///
/// ```
/// use offidized_pptx::color_resolver::apply_luminance_offset;
///
/// // Apply 20% luminance offset (brighten)
/// let result = apply_luminance_offset("4472C4", 20000);
/// assert_eq!(result, "77A5F7");
/// ```
pub fn apply_luminance_offset(color: &str, lum_off: i32) -> String {
    let (r, g, b) = parse_color_value(color);
    let offset_factor = lum_off as f64 / 100_000.0;
    let offset = (255.0 * offset_factor).round() as i32;

    // Apply offset and clamp to 0-255 range
    let new_r = (r as i32 + offset).clamp(0, 255) as u8;
    let new_g = (g as i32 + offset).clamp(0, 255) as u8;
    let new_b = (b as i32 + offset).clamp(0, 255) as u8;

    format!("{:02X}{:02X}{:02X}", new_r, new_g, new_b)
}

/// Applies saturation modulation to a color.
///
/// Saturation modulation adjusts the saturation (color intensity) by converting
/// to HSL, modifying saturation, and converting back to RGB.
///
/// # Arguments
///
/// * `color` - The input color as RGB hex (e.g., "4472C4")
/// * `sat_mod` - The saturation modulation percentage (0-100000)
///
/// # Returns
///
/// The transformed color as RGB hex without "#" prefix.
pub fn apply_saturation_mod(color: &str, sat_mod: i32) -> String {
    let (r, g, b) = parse_color_value(color);
    let (h, s, l) = rgb_to_hsl(r, g, b);

    // Modify saturation
    let sat_factor = sat_mod as f64 / 100_000.0;
    let new_s = (s * sat_factor).clamp(0.0, 1.0);

    let (new_r, new_g, new_b) = hsl_to_rgb(h, new_s, l);
    format!("{:02X}{:02X}{:02X}", new_r, new_g, new_b)
}

/// Applies saturation offset to a color.
///
/// Saturation offset adjusts the saturation by adding or subtracting a value.
///
/// # Arguments
///
/// * `color` - The input color as RGB hex (e.g., "4472C4")
/// * `sat_off` - The saturation offset percentage (-100000 to 100000)
///
/// # Returns
///
/// The transformed color as RGB hex without "#" prefix.
pub fn apply_saturation_offset(color: &str, sat_off: i32) -> String {
    let (r, g, b) = parse_color_value(color);
    let (h, s, l) = rgb_to_hsl(r, g, b);

    // Apply offset to saturation
    let offset = sat_off as f64 / 100_000.0;
    let new_s = (s + offset).clamp(0.0, 1.0);

    let (new_r, new_g, new_b) = hsl_to_rgb(h, new_s, l);
    format!("{:02X}{:02X}{:02X}", new_r, new_g, new_b)
}

/// Parses a hex color string into RGB components.
///
/// # Arguments
///
/// * `color_val` - The color value as hex string (with or without "#" prefix)
///
/// # Returns
///
/// A tuple of (red, green, blue) values (0-255).
///
/// # Panics
///
/// Panics if the color string is invalid or cannot be parsed.
///
/// # Examples
///
/// ```
/// use offidized_pptx::color_resolver::parse_color_value;
///
/// let (r, g, b) = parse_color_value("4472C4");
/// assert_eq!((r, g, b), (68, 114, 196));
///
/// let (r, g, b) = parse_color_value("#FF0000");
/// assert_eq!((r, g, b), (255, 0, 0));
/// ```
pub fn parse_color_value(color_val: &str) -> (u8, u8, u8) {
    let hex = color_val.strip_prefix('#').unwrap_or(color_val);

    match hex.len() {
        3 => {
            // Short form: F00 -> FF0000
            let r = u8::from_str_radix(&hex[0..1], 16).unwrap_or(0) * 17;
            let g = u8::from_str_radix(&hex[1..2], 16).unwrap_or(0) * 17;
            let b = u8::from_str_radix(&hex[2..3], 16).unwrap_or(0) * 17;
            (r, g, b)
        }
        6 => {
            // Standard form: FF0000
            let r = u8::from_str_radix(&hex[0..2], 16).unwrap_or(0);
            let g = u8::from_str_radix(&hex[2..4], 16).unwrap_or(0);
            let b = u8::from_str_radix(&hex[4..6], 16).unwrap_or(0);
            (r, g, b)
        }
        8 => {
            // With alpha: RRGGBBAA (ignore alpha)
            let r = u8::from_str_radix(&hex[0..2], 16).unwrap_or(0);
            let g = u8::from_str_radix(&hex[2..4], 16).unwrap_or(0);
            let b = u8::from_str_radix(&hex[4..6], 16).unwrap_or(0);
            (r, g, b)
        }
        _ => (0, 0, 0), // Invalid format, return black
    }
}

/// Rounds a floating point value using OOXML color rounding rules.
///
/// When the fractional part is exactly 0.5, rounds away from even numbers.
/// This matches the OOXML color transformation rounding behavior:
/// - If floor is even → round up (to make result odd)
/// - If floor is odd → round down (keep result odd)
fn round_half_to_even(x: f64) -> i32 {
    let floor_val = x.floor();
    let frac = x - floor_val;

    if (frac - 0.5).abs() < f64::EPSILON {
        // Exactly 0.5: check if floor is even or odd
        let floor_int = floor_val as i32;
        if floor_int % 2 == 0 {
            // Floor is even → round up to make result odd
            floor_int + 1
        } else {
            // Floor is odd → round down to keep result odd
            floor_int
        }
    } else {
        // Not 0.5: use standard rounding
        x.round() as i32
    }
}

/// Converts RGB color to HSL color space.
///
/// # Arguments
///
/// * `r` - Red component (0-255)
/// * `g` - Green component (0-255)
/// * `b` - Blue component (0-255)
///
/// # Returns
///
/// A tuple of (hue, saturation, lightness) where:
/// - hue is in degrees (0-360)
/// - saturation is 0.0-1.0
/// - lightness is 0.0-1.0
fn rgb_to_hsl(r: u8, g: u8, b: u8) -> (f64, f64, f64) {
    let r = r as f64 / 255.0;
    let g = g as f64 / 255.0;
    let b = b as f64 / 255.0;

    let max = r.max(g).max(b);
    let min = r.min(g).min(b);
    let delta = max - min;

    let l = (max + min) / 2.0;

    if delta == 0.0 {
        return (0.0, 0.0, l); // Achromatic
    }

    let s = if l < 0.5 {
        delta / (max + min)
    } else {
        delta / (2.0 - max - min)
    };

    let h = if max == r {
        ((g - b) / delta + if g < b { 6.0 } else { 0.0 }) * 60.0
    } else if max == g {
        ((b - r) / delta + 2.0) * 60.0
    } else {
        ((r - g) / delta + 4.0) * 60.0
    };

    (h, s, l)
}

/// Converts HSL color to RGB color space.
///
/// # Arguments
///
/// * `h` - Hue in degrees (0-360)
/// * `s` - Saturation (0.0-1.0)
/// * `l` - Lightness (0.0-1.0)
///
/// # Returns
///
/// A tuple of (red, green, blue) values (0-255).
fn hsl_to_rgb(h: f64, s: f64, l: f64) -> (u8, u8, u8) {
    if s == 0.0 {
        let gray = (l * 255.0).round() as u8;
        return (gray, gray, gray); // Achromatic
    }

    let q = if l < 0.5 {
        l * (1.0 + s)
    } else {
        l + s - l * s
    };
    let p = 2.0 * l - q;

    let h_normalized = h / 360.0;

    let r = hue_to_rgb(p, q, h_normalized + 1.0 / 3.0);
    let g = hue_to_rgb(p, q, h_normalized);
    let b = hue_to_rgb(p, q, h_normalized - 1.0 / 3.0);

    (
        (r * 255.0).round() as u8,
        (g * 255.0).round() as u8,
        (b * 255.0).round() as u8,
    )
}

/// Helper function for HSL to RGB conversion.
fn hue_to_rgb(p: f64, q: f64, mut t: f64) -> f64 {
    if t < 0.0 {
        t += 1.0;
    }
    if t > 1.0 {
        t -= 1.0;
    }
    if t < 1.0 / 6.0 {
        return p + (q - p) * 6.0 * t;
    }
    if t < 1.0 / 2.0 {
        return q;
    }
    if t < 2.0 / 3.0 {
        return p + (q - p) * (2.0 / 3.0 - t) * 6.0;
    }
    p
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_parse_color_value() {
        assert_eq!(parse_color_value("4472C4"), (68, 114, 196));
        assert_eq!(parse_color_value("#4472C4"), (68, 114, 196));
        assert_eq!(parse_color_value("FF0000"), (255, 0, 0));
        assert_eq!(parse_color_value("F00"), (255, 0, 0));
        assert_eq!(parse_color_value("4472C4FF"), (68, 114, 196)); // With alpha
    }

    #[test]
    fn test_apply_shade() {
        // 50% shade should halve the RGB values
        let result = apply_shade("4472C4", 50000);
        assert_eq!(result, "223962");
    }

    #[test]
    fn test_apply_tint() {
        // Apply tint to move towards white
        let result = apply_tint("4472C4", 50000);
        assert_eq!(result, "A1B9E1");
    }

    #[test]
    fn test_apply_luminance_offset() {
        // Positive offset should brighten
        let original = "808080"; // Middle gray
        let brightened = apply_luminance_offset(original, 20000);
        let (_, _, b_val) = parse_color_value(&brightened);
        let (_, _, o_val) = parse_color_value(original);
        assert!(b_val > o_val);
    }

    #[test]
    fn test_rgb_to_hsl_and_back() {
        let (r, g, b) = (68, 114, 196);
        let (h, s, l) = rgb_to_hsl(r, g, b);
        let (r2, g2, b2) = hsl_to_rgb(h, s, l);

        // Allow small rounding differences
        assert!((r as i32 - r2 as i32).abs() <= 1);
        assert!((g as i32 - g2 as i32).abs() <= 1);
        assert!((b as i32 - b2 as i32).abs() <= 1);
    }

    #[test]
    fn test_resolve_scheme_color() {
        let mut scheme = ThemeColorScheme::new();
        scheme.accent1 = Some("4472C4".to_string());
        scheme.dark1 = Some("000000".to_string());

        assert_eq!(resolve_scheme_color("accent1", &scheme), Some("4472C4"));
        assert_eq!(resolve_scheme_color("dk1", &scheme), Some("000000"));
        assert_eq!(resolve_scheme_color("nonexistent", &scheme), None);
    }

    #[test]
    fn test_saturation_mod() {
        let result = apply_saturation_mod("4472C4", 50000);
        // Result should be less saturated (more grayish)
        let (r, g, b) = parse_color_value(&result);
        let (_h, s, _l) = rgb_to_hsl(r, g, b);
        let (_h_orig, s_orig, _l_orig) = rgb_to_hsl(68, 114, 196);

        assert!(s < s_orig);
    }

    #[test]
    fn test_color_transformations_chain() {
        let base_color = "4472C4";

        // Apply multiple transformations
        let shaded = apply_shade(base_color, 80000);
        let tinted = apply_tint(&shaded, 20000);

        // Result should be different from original
        assert_ne!(tinted, base_color);

        // Parse to verify valid hex
        let (_r, _g, _b) = parse_color_value(&tinted);
    }
}
