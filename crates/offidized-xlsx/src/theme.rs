//! Theme XML parsing for OOXML workbooks.
//!
//! Parses `/xl/theme/theme1.xml` to extract theme color palette and font names.
//! The theme defines 12 color slots (lt1, dk1, lt2, dk2, accent1-6, hlink, folHlink)
//! and major/minor font names used throughout the workbook.

use std::collections::HashMap;
use std::io::BufRead;

use quick_xml::events::Event;
use quick_xml::Reader;

/// Default Office theme colors (hex with `#` prefix) in canonical OOXML order:
/// lt1, dk1, lt2, dk2, accent1-6, hlink, folHlink.
///
/// These correspond to the standard Office 2007+ theme palette.
pub const DEFAULT_THEME_COLORS: [&str; 12] = [
    "#FFFFFF", // lt1 – Background 1 (Light 1)
    "#000000", // dk1 – Text 1 (Dark 1)
    "#EEECE1", // lt2 – Background 2 (Light 2)
    "#1F497D", // dk2 – Text 2 (Dark 2)
    "#4F81BD", // accent1
    "#C0504D", // accent2
    "#9BBB59", // accent3
    "#8064A2", // accent4
    "#4BACC6", // accent5
    "#F79646", // accent6
    "#0000FF", // hlink
    "#800080", // folHlink
];

/// Parsed theme data from `theme1.xml`.
#[derive(Debug, Clone)]
pub struct ParsedTheme {
    /// 12 hex color strings (with `#` prefix) in canonical OOXML order.
    pub colors: Vec<String>,
    /// Major (heading) font name from the theme.
    pub major_font: Option<String>,
    /// Minor (body) font name from the theme.
    pub minor_font: Option<String>,
}

impl Default for ParsedTheme {
    fn default() -> Self {
        Self {
            colors: DEFAULT_THEME_COLORS
                .iter()
                .map(|s| (*s).to_string())
                .collect(),
            major_font: None,
            minor_font: None,
        }
    }
}

/// Parse theme colors and fonts from theme XML bytes.
///
/// Returns a [`ParsedTheme`] with the 12 color slots and optional font names.
/// If the XML is malformed or missing expected elements, default colors are used.
pub(crate) fn parse_theme_xml(xml: &[u8]) -> ParsedTheme {
    let reader = std::io::BufReader::new(xml);
    let mut xml_reader = Reader::from_reader(reader);
    xml_reader.config_mut().trim_text(true);

    let mut colors = Vec::new();
    let mut major_font = None;
    let mut minor_font = None;
    let mut buf = Vec::new();

    loop {
        match xml_reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref e)) => {
                let name_bytes = e.local_name();
                match name_bytes.as_ref() {
                    b"clrScheme" => {
                        colors = parse_color_scheme(&mut xml_reader);
                    }
                    b"fontScheme" => {
                        let (major, minor) = parse_font_scheme(&mut xml_reader);
                        major_font = major;
                        minor_font = minor;
                    }
                    _ => {}
                }
            }
            Ok(Event::Eof) => break,
            Err(_) => break,
            _ => {}
        }
        buf.clear();
    }

    // Fall back to defaults if no colors were parsed
    if colors.is_empty() {
        colors = DEFAULT_THEME_COLORS
            .iter()
            .map(|s| (*s).to_string())
            .collect();
    }

    ParsedTheme {
        colors,
        major_font,
        minor_font,
    }
}

/// Color slot names in canonical OOXML order for theme index mapping.
///
/// Index 0 = lt1 (Light 1 / Background 1), index 1 = dk1 (Dark 1 / Text 1), etc.
const COLOR_ORDER: [&str; 12] = [
    "lt1", "dk1", "lt2", "dk2", "accent1", "accent2", "accent3", "accent4", "accent5", "accent6",
    "hlink", "folHlink",
];

/// Extract a color hex value from `sysClr` or `srgbClr` attributes on a `BytesStart` event.
fn extract_color_from_event(e: &quick_xml::events::BytesStart<'_>) -> Option<String> {
    let name_bytes = e.local_name();
    match name_bytes.as_ref() {
        b"sysClr" => {
            // System color — use `lastClr` attribute for the actual resolved color
            for attr in e.attributes().flatten() {
                let key = attr.key.local_name();
                if key.as_ref() == b"lastClr" {
                    let val = String::from_utf8_lossy(&attr.value).to_uppercase();
                    return Some(format!("#{val}"));
                }
            }
            None
        }
        b"srgbClr" => {
            // sRGB color — use `val` attribute
            for attr in e.attributes().flatten() {
                let key = attr.key.local_name();
                if key.as_ref() == b"val" {
                    let val = String::from_utf8_lossy(&attr.value).to_uppercase();
                    return Some(format!("#{val}"));
                }
            }
            None
        }
        _ => None,
    }
}

/// Parse theme colors from the `<a:clrScheme>` element.
///
/// Expects the reader to be positioned just after the `<clrScheme>` start event.
/// Reads child elements (dk1, lt1, dk2, lt2, accent1-6, hlink, folHlink) and extracts
/// their color values from nested `sysClr` or `srgbClr` elements.
fn parse_color_scheme<R: BufRead>(xml: &mut Reader<R>) -> Vec<String> {
    let mut color_map = HashMap::<String, String>::new();
    let mut buf = Vec::new();
    let mut current_color_name: Option<String> = None;
    let mut depth: u32 = 1; // Already inside clrScheme

    loop {
        match xml.read_event_into(&mut buf) {
            Ok(Event::Start(ref e)) => {
                depth = depth.saturating_add(1);
                let name_bytes = e.local_name();
                let name_str = String::from_utf8_lossy(name_bytes.as_ref()).to_string();

                if COLOR_ORDER.contains(&name_str.as_str()) {
                    current_color_name = Some(name_str);
                } else if let Some(ref color_name) = current_color_name {
                    if let Some(hex) = extract_color_from_event(e) {
                        color_map.insert(color_name.clone(), hex);
                    }
                }
            }
            Ok(Event::Empty(ref e)) => {
                let name_bytes = e.local_name();
                let name_str = String::from_utf8_lossy(name_bytes.as_ref()).to_string();

                if !COLOR_ORDER.contains(&name_str.as_str()) {
                    if let Some(ref color_name) = current_color_name {
                        if let Some(hex) = extract_color_from_event(e) {
                            color_map.insert(color_name.clone(), hex);
                        }
                    }
                }
            }
            Ok(Event::End(ref e)) => {
                depth = depth.saturating_sub(1);
                let name_bytes = e.local_name();
                let name_str = String::from_utf8_lossy(name_bytes.as_ref()).to_string();

                if COLOR_ORDER.contains(&name_str.as_str()) {
                    current_color_name = None;
                }

                if depth == 0 {
                    break;
                }
            }
            Ok(Event::Eof) => break,
            Err(_) => break,
            _ => {}
        }
        buf.clear();
    }

    // Build the color vector in the correct order, falling back to defaults
    COLOR_ORDER
        .iter()
        .enumerate()
        .map(|(idx, name)| {
            color_map.get(*name).cloned().unwrap_or_else(|| {
                DEFAULT_THEME_COLORS
                    .get(idx)
                    .map(|s| (*s).to_string())
                    .unwrap_or_else(|| "#000000".to_string())
            })
        })
        .collect()
}

/// Extract the `typeface` attribute from a `<a:latin>` element.
fn extract_typeface(e: &quick_xml::events::BytesStart<'_>) -> Option<String> {
    for attr in e.attributes().flatten() {
        let key = attr.key.local_name();
        if key.as_ref() == b"typeface" {
            let value = String::from_utf8_lossy(&attr.value).to_string();
            if !value.is_empty() {
                return Some(value);
            }
        }
    }
    None
}

/// Parse theme fonts from the `<a:fontScheme>` element.
///
/// Expects the reader to be positioned just after the `<fontScheme>` start event.
/// Returns `(major_font, minor_font)` extracted from `<a:latin typeface="..."/>`.
fn parse_font_scheme<R: BufRead>(xml: &mut Reader<R>) -> (Option<String>, Option<String>) {
    let mut major_font = None;
    let mut minor_font = None;
    let mut buf = Vec::new();
    let mut in_major_font = false;
    let mut in_minor_font = false;
    let mut depth: u32 = 1; // Already inside fontScheme

    loop {
        match xml.read_event_into(&mut buf) {
            Ok(Event::Start(ref e)) => {
                depth = depth.saturating_add(1);
                let name_bytes = e.local_name();

                match name_bytes.as_ref() {
                    b"majorFont" => {
                        in_major_font = true;
                        in_minor_font = false;
                    }
                    b"minorFont" => {
                        in_minor_font = true;
                        in_major_font = false;
                    }
                    b"latin" => {
                        if let Some(typeface) = extract_typeface(e) {
                            if in_major_font {
                                major_font = Some(typeface);
                            } else if in_minor_font {
                                minor_font = Some(typeface);
                            }
                        }
                    }
                    _ => {}
                }
            }
            Ok(Event::Empty(ref e)) => {
                let name_bytes = e.local_name();
                if name_bytes.as_ref() == b"latin" {
                    if let Some(typeface) = extract_typeface(e) {
                        if in_major_font {
                            major_font = Some(typeface);
                        } else if in_minor_font {
                            minor_font = Some(typeface);
                        }
                    }
                }
            }
            Ok(Event::End(ref e)) => {
                depth = depth.saturating_sub(1);
                let name_bytes = e.local_name();

                match name_bytes.as_ref() {
                    b"majorFont" => {
                        in_major_font = false;
                    }
                    b"minorFont" => {
                        in_minor_font = false;
                    }
                    _ => {}
                }

                if depth == 0 {
                    break;
                }
            }
            Ok(Event::Eof) => break,
            Err(_) => break,
            _ => {}
        }
        buf.clear();
    }

    (major_font, minor_font)
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
    use std::io::Cursor;

    #[test]
    fn parse_color_scheme_standard_office_theme() {
        let xml_content = r#"
        <a:clrScheme name="Office">
            <a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1>
            <a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1>
            <a:dk2><a:srgbClr val="1F497D"/></a:dk2>
            <a:lt2><a:srgbClr val="EEECE1"/></a:lt2>
            <a:accent1><a:srgbClr val="4F81BD"/></a:accent1>
            <a:accent2><a:srgbClr val="C0504D"/></a:accent2>
            <a:accent3><a:srgbClr val="9BBB59"/></a:accent3>
            <a:accent4><a:srgbClr val="8064A2"/></a:accent4>
            <a:accent5><a:srgbClr val="4BACC6"/></a:accent5>
            <a:accent6><a:srgbClr val="F79646"/></a:accent6>
            <a:hlink><a:srgbClr val="0000FF"/></a:hlink>
            <a:folHlink><a:srgbClr val="800080"/></a:folHlink>
        </a:clrScheme>
        "#;

        let cursor = Cursor::new(xml_content);
        let mut reader = Reader::from_reader(cursor);
        reader.config_mut().trim_text(true);

        // Skip to clrScheme start
        let mut buf = Vec::new();
        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"clrScheme" => break,
                Ok(Event::Eof) => panic!("Unexpected EOF"),
                _ => {}
            }
            buf.clear();
        }

        let colors = parse_color_scheme(&mut reader);

        assert_eq!(colors.len(), 12);
        assert_eq!(colors[0], "#FFFFFF"); // lt1 (Background 1)
        assert_eq!(colors[1], "#000000"); // dk1 (Text 1)
        assert_eq!(colors[2], "#EEECE1"); // lt2 (Background 2)
        assert_eq!(colors[3], "#1F497D"); // dk2 (Text 2)
        assert_eq!(colors[4], "#4F81BD"); // accent1
        assert_eq!(colors[5], "#C0504D"); // accent2
        assert_eq!(colors[6], "#9BBB59"); // accent3
        assert_eq!(colors[7], "#8064A2"); // accent4
        assert_eq!(colors[8], "#4BACC6"); // accent5
        assert_eq!(colors[9], "#F79646"); // accent6
        assert_eq!(colors[10], "#0000FF"); // hlink
        assert_eq!(colors[11], "#800080"); // folHlink
    }

    #[test]
    fn parse_font_scheme_standard_office() {
        let xml_content = r#"
        <a:fontScheme name="Office">
            <a:majorFont>
                <a:latin typeface="Cambria"/>
                <a:ea typeface=""/>
                <a:cs typeface=""/>
            </a:majorFont>
            <a:minorFont>
                <a:latin typeface="Calibri"/>
                <a:ea typeface=""/>
                <a:cs typeface=""/>
            </a:minorFont>
        </a:fontScheme>
        "#;

        let cursor = Cursor::new(xml_content);
        let mut reader = Reader::from_reader(cursor);
        reader.config_mut().trim_text(true);

        // Skip to fontScheme start
        let mut buf = Vec::new();
        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"fontScheme" => break,
                Ok(Event::Eof) => panic!("Unexpected EOF"),
                _ => {}
            }
            buf.clear();
        }

        let (major, minor) = parse_font_scheme(&mut reader);

        assert_eq!(major, Some("Cambria".to_string()));
        assert_eq!(minor, Some("Calibri".to_string()));
    }

    #[test]
    fn parse_font_scheme_custom_fonts() {
        let xml_content = r#"
        <a:fontScheme name="Custom">
            <a:majorFont>
                <a:latin typeface="Arial"/>
            </a:majorFont>
            <a:minorFont>
                <a:latin typeface="Times New Roman"/>
            </a:minorFont>
        </a:fontScheme>
        "#;

        let cursor = Cursor::new(xml_content);
        let mut reader = Reader::from_reader(cursor);
        reader.config_mut().trim_text(true);

        let mut buf = Vec::new();
        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"fontScheme" => break,
                Ok(Event::Eof) => panic!("Unexpected EOF"),
                _ => {}
            }
            buf.clear();
        }

        let (major, minor) = parse_font_scheme(&mut reader);

        assert_eq!(major, Some("Arial".to_string()));
        assert_eq!(minor, Some("Times New Roman".to_string()));
    }

    #[test]
    fn parse_full_theme_xml() {
        let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office Theme">
            <a:themeElements>
                <a:clrScheme name="Office">
                    <a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1>
                    <a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1>
                    <a:dk2><a:srgbClr val="44546A"/></a:dk2>
                    <a:lt2><a:srgbClr val="E7E6E6"/></a:lt2>
                    <a:accent1><a:srgbClr val="4472C4"/></a:accent1>
                    <a:accent2><a:srgbClr val="ED7D31"/></a:accent2>
                    <a:accent3><a:srgbClr val="A5A5A5"/></a:accent3>
                    <a:accent4><a:srgbClr val="FFC000"/></a:accent4>
                    <a:accent5><a:srgbClr val="5B9BD5"/></a:accent5>
                    <a:accent6><a:srgbClr val="70AD47"/></a:accent6>
                    <a:hlink><a:srgbClr val="0563C1"/></a:hlink>
                    <a:folHlink><a:srgbClr val="954F72"/></a:folHlink>
                </a:clrScheme>
                <a:fontScheme name="Office">
                    <a:majorFont>
                        <a:latin typeface="Calibri Light"/>
                    </a:majorFont>
                    <a:minorFont>
                        <a:latin typeface="Calibri"/>
                    </a:minorFont>
                </a:fontScheme>
            </a:themeElements>
        </a:theme>"#;

        let theme = parse_theme_xml(xml.as_bytes());

        assert_eq!(theme.colors.len(), 12);
        assert_eq!(theme.colors[0], "#FFFFFF"); // lt1
        assert_eq!(theme.colors[1], "#000000"); // dk1
        assert_eq!(theme.colors[2], "#E7E6E6"); // lt2
        assert_eq!(theme.colors[3], "#44546A"); // dk2
        assert_eq!(theme.colors[4], "#4472C4"); // accent1
        assert_eq!(theme.colors[5], "#ED7D31"); // accent2
        assert_eq!(theme.colors[6], "#A5A5A5"); // accent3
        assert_eq!(theme.colors[7], "#FFC000"); // accent4
        assert_eq!(theme.colors[8], "#5B9BD5"); // accent5
        assert_eq!(theme.colors[9], "#70AD47"); // accent6
        assert_eq!(theme.colors[10], "#0563C1"); // hlink
        assert_eq!(theme.colors[11], "#954F72"); // folHlink

        assert_eq!(theme.major_font, Some("Calibri Light".to_string()));
        assert_eq!(theme.minor_font, Some("Calibri".to_string()));
    }

    #[test]
    fn parse_theme_xml_defaults_on_empty() {
        let xml = r#"<?xml version="1.0" encoding="UTF-8"?>
        <a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Empty">
            <a:themeElements>
            </a:themeElements>
        </a:theme>"#;

        let theme = parse_theme_xml(xml.as_bytes());

        assert_eq!(theme.colors.len(), 12);
        // Should be all defaults
        for (i, color) in theme.colors.iter().enumerate() {
            assert_eq!(color, DEFAULT_THEME_COLORS[i]);
        }
        assert_eq!(theme.major_font, None);
        assert_eq!(theme.minor_font, None);
    }

    #[test]
    fn parse_theme_xml_defaults_on_malformed() {
        let xml = b"not valid xml at all <<>>";

        let theme = parse_theme_xml(xml);

        assert_eq!(theme.colors.len(), 12);
        for (i, color) in theme.colors.iter().enumerate() {
            assert_eq!(color, DEFAULT_THEME_COLORS[i]);
        }
    }

    #[test]
    fn default_parsed_theme_has_twelve_colors() {
        let theme = ParsedTheme::default();
        assert_eq!(theme.colors.len(), 12);
        assert_eq!(theme.major_font, None);
        assert_eq!(theme.minor_font, None);
    }
}
