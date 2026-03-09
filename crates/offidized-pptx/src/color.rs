//! Color model for OOXML PresentationML.
//!
//! This module provides a unified color representation that handles both
//! direct sRGB colors (`a:srgbClr`) and theme/scheme color references
//! (`a:schemeClr`), along with color transforms like luminance modulation,
//! tints, shades, and alpha adjustments.

use std::io::Cursor;

use quick_xml::events::{BytesEnd, BytesStart, Event};
use quick_xml::Writer;

use crate::error::Result;

/// A color value in OOXML PresentationML.
///
/// Colors in OOXML can be specified either as direct sRGB hex values or as
/// references to theme/scheme colors. Both variants can have an optional alpha
/// value and a list of color transforms applied.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum ShapeColor {
    /// A direct sRGB color value (hex string like "FF0000").
    SrgbClr {
        /// The sRGB hex color value (e.g. "FF0000" for red).
        val: String,
        /// Alpha/opacity as percentage (0-100, where 100 = fully opaque).
        alpha: Option<u8>,
        /// Color transforms applied to this color.
        transforms: Vec<ColorTransform>,
    },
    /// A theme/scheme color reference.
    SchemeClr {
        /// The scheme color name (e.g. "accent1", "dk1", "lt1", "tx1").
        val: String,
        /// Alpha/opacity as percentage (0-100, where 100 = fully opaque).
        alpha: Option<u8>,
        /// Color transforms applied to this color.
        transforms: Vec<ColorTransform>,
    },
}

impl ShapeColor {
    /// Create a new sRGB color with no transforms.
    pub fn srgb(val: impl Into<String>) -> Self {
        Self::SrgbClr {
            val: val.into(),
            alpha: None,
            transforms: Vec::new(),
        }
    }

    /// Create a new sRGB color with alpha.
    pub fn srgb_with_alpha(val: impl Into<String>, alpha: u8) -> Self {
        Self::SrgbClr {
            val: val.into(),
            alpha: Some(alpha),
            transforms: Vec::new(),
        }
    }

    /// Create a new scheme color reference with no transforms.
    pub fn scheme(val: impl Into<String>) -> Self {
        Self::SchemeClr {
            val: val.into(),
            alpha: None,
            transforms: Vec::new(),
        }
    }

    /// Returns the sRGB hex value if this is an SrgbClr, None otherwise.
    pub fn srgb_value(&self) -> Option<&str> {
        match self {
            Self::SrgbClr { val, .. } => Some(val.as_str()),
            Self::SchemeClr { .. } => None,
        }
    }

    /// Returns the scheme color name if this is a SchemeClr, None otherwise.
    pub fn scheme_value(&self) -> Option<&str> {
        match self {
            Self::SrgbClr { .. } => None,
            Self::SchemeClr { val, .. } => Some(val.as_str()),
        }
    }

    /// Returns the alpha value regardless of color type.
    pub fn alpha(&self) -> Option<u8> {
        match self {
            Self::SrgbClr { alpha, .. } | Self::SchemeClr { alpha, .. } => *alpha,
        }
    }

    /// Returns the transforms regardless of color type.
    pub fn transforms(&self) -> &[ColorTransform] {
        match self {
            Self::SrgbClr { transforms, .. } | Self::SchemeClr { transforms, .. } => transforms,
        }
    }

    /// Returns true if this color is an sRGB color.
    pub fn is_srgb(&self) -> bool {
        matches!(self, Self::SrgbClr { .. })
    }

    /// Returns true if this color is a scheme color.
    pub fn is_scheme(&self) -> bool {
        matches!(self, Self::SchemeClr { .. })
    }
}

/// A color transformation applied to a base color.
///
/// These transforms modify the base color (either sRGB or scheme) and are
/// serialized as child elements of the color element in OOXML XML.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum ColorTransform {
    /// Luminance modulation (percentage in thousandths, e.g. 75000 = 75%).
    LumMod(i32),
    /// Luminance offset (percentage in thousandths).
    LumOff(i32),
    /// Tint (percentage in thousandths).
    Tint(i32),
    /// Shade (percentage in thousandths).
    Shade(i32),
    /// Saturation modulation (percentage in thousandths).
    SatMod(i32),
    /// Saturation offset (percentage in thousandths).
    SatOff(i32),
    /// Alpha/opacity (percentage in thousandths, e.g. 50000 = 50%).
    Alpha(i32),
    /// Hue offset (in 60,000ths of a degree).
    HueOff(i32),
    /// Hue modulation (percentage in thousandths).
    HueMod(i32),
    /// An unknown transform element preserved for roundtrip fidelity.
    Unknown(String, i32),
}

impl ColorTransform {
    /// Returns the XML element local name for this transform.
    pub fn xml_name(&self) -> &str {
        match self {
            Self::LumMod(_) => "lumMod",
            Self::LumOff(_) => "lumOff",
            Self::Tint(_) => "tint",
            Self::Shade(_) => "shade",
            Self::SatMod(_) => "satMod",
            Self::SatOff(_) => "satOff",
            Self::Alpha(_) => "alpha",
            Self::HueOff(_) => "hueOff",
            Self::HueMod(_) => "hueMod",
            Self::Unknown(name, _) => name.as_str(),
        }
    }

    /// Returns the value of this transform.
    pub fn value(&self) -> i32 {
        match self {
            Self::LumMod(v)
            | Self::LumOff(v)
            | Self::Tint(v)
            | Self::Shade(v)
            | Self::SatMod(v)
            | Self::SatOff(v)
            | Self::Alpha(v)
            | Self::HueOff(v)
            | Self::HueMod(v)
            | Self::Unknown(_, v) => *v,
        }
    }

    /// Parse a color transform from an XML local name and value.
    pub fn from_xml(local_name: &[u8], val: i32) -> Self {
        match local_name {
            b"lumMod" => Self::LumMod(val),
            b"lumOff" => Self::LumOff(val),
            b"tint" => Self::Tint(val),
            b"shade" => Self::Shade(val),
            b"satMod" => Self::SatMod(val),
            b"satOff" => Self::SatOff(val),
            b"alpha" => Self::Alpha(val),
            b"hueOff" => Self::HueOff(val),
            b"hueMod" => Self::HueMod(val),
            _ => Self::Unknown(String::from_utf8_lossy(local_name).into_owned(), val),
        }
    }
}

/// Write a `ShapeColor` to XML.
///
/// This writes the appropriate color element (`a:srgbClr` or `a:schemeClr`)
/// with its val attribute, alpha child (if present), and any color transforms.
pub fn write_color_xml<W: std::io::Write>(
    writer: &mut Writer<W>,
    color: &ShapeColor,
) -> Result<()> {
    let (tag_name, val, alpha, transforms) = match color {
        ShapeColor::SrgbClr {
            val,
            alpha,
            transforms,
        } => ("a:srgbClr", val.as_str(), alpha, transforms),
        ShapeColor::SchemeClr {
            val,
            alpha,
            transforms,
        } => ("a:schemeClr", val.as_str(), alpha, transforms),
    };

    let mut tag = BytesStart::new(tag_name);
    tag.push_attribute(("val", val));

    let has_children = alpha.is_some() || !transforms.is_empty();
    if !has_children {
        writer.write_event(Event::Empty(tag))?;
    } else {
        writer.write_event(Event::Start(tag))?;

        // Write alpha as a child element (not as a ColorTransform).
        if let Some(a) = alpha {
            let alpha_thousandths = (*a as u32) * 1000;
            let alpha_text = alpha_thousandths.to_string();
            let mut alpha_elem = BytesStart::new("a:alpha");
            alpha_elem.push_attribute(("val", alpha_text.as_str()));
            writer.write_event(Event::Empty(alpha_elem))?;
        }

        // Write transforms.
        for transform in transforms {
            // Skip Alpha transforms here -- they're handled above via the alpha field.
            if matches!(transform, ColorTransform::Alpha(_)) {
                continue;
            }
            let elem_name = format!("a:{}", transform.xml_name());
            let val_text = transform.value().to_string();
            let mut elem = BytesStart::new(elem_name.as_str());
            elem.push_attribute(("val", val_text.as_str()));
            writer.write_event(Event::Empty(elem))?;
        }

        writer.write_event(Event::End(BytesEnd::new(tag_name)))?;
    }

    Ok(())
}

/// Write a `ShapeColor` wrapped in a `<a:solidFill>` element.
pub fn write_solid_fill_color_xml<W: std::io::Write>(
    writer: &mut Writer<W>,
    color: &ShapeColor,
) -> Result<()> {
    writer.write_event(Event::Start(BytesStart::new("a:solidFill")))?;
    write_color_xml(writer, color)?;
    writer.write_event(Event::End(BytesEnd::new("a:solidFill")))?;
    Ok(())
}

/// Serialize a `ShapeColor` to a standalone XML string (for testing/debugging).
pub fn color_to_xml_string(color: &ShapeColor) -> Result<String> {
    let mut writer = Writer::new(Cursor::new(Vec::new()));
    write_color_xml(&mut writer, color)?;
    let bytes = writer.into_inner().into_inner();
    Ok(String::from_utf8_lossy(&bytes).into_owned())
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn srgb_color_simple() {
        let color = ShapeColor::srgb("FF0000");
        assert_eq!(color.srgb_value(), Some("FF0000"));
        assert_eq!(color.scheme_value(), None);
        assert_eq!(color.alpha(), None);
        assert!(color.transforms().is_empty());
        assert!(color.is_srgb());
        assert!(!color.is_scheme());
    }

    #[test]
    fn scheme_color_simple() {
        let color = ShapeColor::scheme("accent1");
        assert_eq!(color.srgb_value(), None);
        assert_eq!(color.scheme_value(), Some("accent1"));
        assert_eq!(color.alpha(), None);
        assert!(color.is_scheme());
    }

    #[test]
    fn srgb_color_with_alpha_serializes() {
        let color = ShapeColor::srgb_with_alpha("FF0000", 50);
        let xml = color_to_xml_string(&color).unwrap();
        assert!(xml.contains("a:srgbClr"));
        assert!(xml.contains(r#"val="FF0000""#));
        assert!(xml.contains(r#"val="50000""#));
    }

    #[test]
    fn scheme_color_with_transforms_serializes() {
        let color = ShapeColor::SchemeClr {
            val: "accent1".to_string(),
            alpha: None,
            transforms: vec![ColorTransform::LumMod(75000), ColorTransform::LumOff(25000)],
        };
        let xml = color_to_xml_string(&color).unwrap();
        assert!(xml.contains("a:schemeClr"));
        assert!(xml.contains(r#"val="accent1""#));
        assert!(xml.contains("a:lumMod"));
        assert!(xml.contains("a:lumOff"));
    }

    #[test]
    fn color_transform_roundtrip() {
        let transforms = [
            (b"lumMod" as &[u8], 75000),
            (b"lumOff", 25000),
            (b"tint", 50000),
            (b"shade", 80000),
            (b"satMod", 120000),
            (b"alpha", 50000),
        ];
        for (name, val) in transforms {
            let t = ColorTransform::from_xml(name, val);
            assert_eq!(t.value(), val);
        }
    }
}
