//! SlideMaster XML parsing and writing.
//!
//! This module handles roundtrip-fidelity serialization of `<p:sldMaster>` elements,
//! the PresentationML slide master parts. A slide master defines the common geometry,
//! color scheme, fonts, background and shapes inherited by its associated slide layouts.

#![allow(dead_code)]

use std::io::Cursor;

use quick_xml::events::{BytesDecl, BytesEnd, BytesStart, Event};
use quick_xml::{Reader, Writer};

use crate::error::{PptxError, Result};
use crate::shape::{PlaceholderType, Shape};
use crate::slide::SlideBackground;
use crate::theme::ThemeColorScheme;
use offidized_opc::RawXmlNode;

const PRESENTATIONML_NS: &str = "http://schemas.openxmlformats.org/presentationml/2006/main";
const DRAWINGML_NS: &str = "http://schemas.openxmlformats.org/drawingml/2006/main";
const RELATIONSHIP_NS: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

/// Parsed metadata from a `<p:sldMaster>` element.
#[allow(dead_code)]
#[derive(Debug, Clone, PartialEq)]
pub(crate) struct ParsedSlideMasterData {
    /// The `preserve` attribute on `<p:sldMaster>`.
    pub preserve: Option<bool>,
    /// Layout relationship IDs from `<p:sldLayoutIdLst>`.
    pub layout_refs: Vec<ParsedLayoutRef>,
    /// Shapes from `<p:cSld><p:spTree>` (basic metadata extraction).
    pub shapes: Vec<Shape>,
    /// Theme color scheme parsed from an inline `<a:clrScheme>` (rare; usually in theme part).
    pub theme: Option<ThemeColorScheme>,
    /// Background from `<p:cSld><p:bg>`.
    pub background: Option<SlideBackground>,
    /// Color map attributes from `<p:clrMap>`.
    pub color_map: Vec<(String, String)>,
    /// The raw `<p:spTree>` element preserved for roundtrip fidelity.
    pub raw_sp_tree: Option<RawXmlNode>,
    /// Unknown direct children of `<p:sldMaster>` preserved for roundtrip fidelity.
    pub unknown_children: Vec<RawXmlNode>,
}

/// A layout reference from `<p:sldLayoutId>`.
#[derive(Debug, Clone, PartialEq, Eq)]
pub(crate) struct ParsedLayoutRef {
    /// The `id` attribute.
    pub id: Option<String>,
    /// The `r:id` relationship attribute.
    pub relationship_id: String,
}

/// Parse a `<p:sldMaster>` XML part into structured metadata.
///
/// This extracts:
/// - `preserve` attribute
/// - Layout references from `<p:sldLayoutIdLst>`
/// - Shapes from `<p:cSld><p:spTree>` (basic name/placeholder extraction)
/// - Background from `<p:cSld><p:bg>` (solid fill)
/// - Color map from `<p:clrMap>`
/// - Theme color scheme from `<a:clrScheme>` (if inline)
/// - Unknown children as `RawXmlNode` for roundtrip fidelity
/// - Raw shape tree for roundtrip fidelity
#[allow(dead_code)]
pub(crate) fn parse_slide_master_xml(xml: &[u8]) -> Result<ParsedSlideMasterData> {
    let mut reader = Reader::from_reader(Cursor::new(xml));
    reader.config_mut().trim_text(true);
    let mut buffer = Vec::new();

    let mut preserve = None;
    let mut layout_refs = Vec::new();
    let mut color_map = Vec::new();
    let mut unknown_children = Vec::new();
    let mut shapes = Vec::new();
    let mut background = None;
    let mut theme = None;
    let mut raw_sp_tree = None;

    // Depth tracking.
    let mut master_depth = 0_u32; // 0 = outside, 1 = inside <p:sldMaster>
    let mut c_sld_depth = 0_u32;
    let mut in_layout_id_lst = false;

    // Known direct children of <p:sldMaster> we handle explicitly.
    static KNOWN_MASTER_CHILDREN: &[&[u8]] =
        &[b"cSld", b"clrMap", b"sldLayoutIdLst", b"hf", b"txStyles"];

    loop {
        match reader.read_event_into(&mut buffer)? {
            Event::Start(ref event) => {
                let name = event.name();
                let local = local_name(name.as_ref());

                if master_depth == 0 && local == b"sldMaster" {
                    master_depth = 1;
                    preserve = get_attribute_value(event, b"preserve")
                        .as_deref()
                        .and_then(parse_xml_bool);
                } else if master_depth == 1 && local == b"cSld" {
                    c_sld_depth = 1;
                } else if c_sld_depth == 1 && local == b"spTree" {
                    // Capture the raw spTree for roundtrip and extract shapes.
                    let raw_node = RawXmlNode::read_element(&mut reader, event)?;
                    shapes = extract_shapes_from_raw_sp_tree(&raw_node);
                    raw_sp_tree = Some(raw_node);
                } else if c_sld_depth == 1 && local == b"bg" {
                    // Parse background.
                    background = parse_background_subtree(&mut reader)?;
                } else if master_depth == 1 && c_sld_depth == 0 && local == b"sldLayoutIdLst" {
                    in_layout_id_lst = true;
                } else if master_depth == 1 && c_sld_depth == 0 && local == b"clrMap" {
                    for attr in event.attributes().flatten() {
                        let key = String::from_utf8_lossy(attr.key.as_ref()).into_owned();
                        let val = String::from_utf8_lossy(&attr.value).into_owned();
                        color_map.push((key, val));
                    }
                    skip_to_end_tag(&mut reader, event.name().as_ref())?;
                } else if in_layout_id_lst && local == b"sldLayoutId" {
                    let id = get_attribute_value(event, b"id");
                    let relationship_id =
                        get_relationship_id_attribute_value(event).ok_or_else(|| {
                            PptxError::UnsupportedPackage(
                                "slide master layout id missing relationship `id` attribute"
                                    .to_string(),
                            )
                        })?;
                    layout_refs.push(ParsedLayoutRef {
                        id,
                        relationship_id,
                    });
                } else if master_depth == 1
                    && c_sld_depth == 0
                    && !in_layout_id_lst
                    && !KNOWN_MASTER_CHILDREN.contains(&local)
                {
                    // Check for inline theme color scheme.
                    if local == b"clrScheme" {
                        theme = parse_theme_color_scheme_subtree(&mut reader, event)?;
                    } else {
                        // Unknown direct child — capture for roundtrip.
                        unknown_children.push(RawXmlNode::read_element(&mut reader, event)?);
                    }
                }
            }
            Event::Empty(ref event) => {
                let name = event.name();
                let local = local_name(name.as_ref());

                if master_depth == 0 && local == b"sldMaster" {
                    break;
                } else if c_sld_depth == 1 && local == b"spTree" {
                    // Empty <p:spTree/> — no shapes.
                    raw_sp_tree = Some(RawXmlNode::from_empty_element(event));
                } else if in_layout_id_lst && local == b"sldLayoutId" {
                    let id = get_attribute_value(event, b"id");
                    let relationship_id =
                        get_relationship_id_attribute_value(event).ok_or_else(|| {
                            PptxError::UnsupportedPackage(
                                "slide master layout id missing relationship `id` attribute"
                                    .to_string(),
                            )
                        })?;
                    layout_refs.push(ParsedLayoutRef {
                        id,
                        relationship_id,
                    });
                } else if master_depth == 1 && c_sld_depth == 0 && local == b"clrMap" {
                    for attr in event.attributes().flatten() {
                        let key = String::from_utf8_lossy(attr.key.as_ref()).into_owned();
                        let val = String::from_utf8_lossy(&attr.value).into_owned();
                        color_map.push((key, val));
                    }
                } else if master_depth == 1
                    && c_sld_depth == 0
                    && !in_layout_id_lst
                    && !KNOWN_MASTER_CHILDREN.contains(&local)
                {
                    unknown_children.push(RawXmlNode::from_empty_element(event));
                }
            }
            Event::End(ref event) => {
                let name = event.name();
                let local = local_name(name.as_ref());

                if local == b"cSld" && c_sld_depth == 1 {
                    c_sld_depth = 0;
                } else if local == b"sldLayoutIdLst" {
                    in_layout_id_lst = false;
                } else if local == b"sldMaster" && master_depth == 1 {
                    break;
                }
            }
            Event::Eof => break,
            _ => {}
        }
        buffer.clear();
    }

    Ok(ParsedSlideMasterData {
        preserve,
        layout_refs,
        shapes,
        theme,
        background,
        color_map,
        raw_sp_tree,
        unknown_children,
    })
}

/// Write a `<p:sldMaster>` XML part from structured data.
///
/// This produces a full, namespace-declared `<p:sldMaster>` element with:
/// - `preserve` attribute (if set)
/// - `<p:cSld>` with background and shape tree
/// - `<p:clrMap>` color map
/// - `<p:sldLayoutIdLst>` with layout references
/// - Unknown children preserved from parsing
pub(crate) fn write_slide_master_xml(data: &WriteSlideMasterData<'_>) -> Result<Vec<u8>> {
    let mut writer = Writer::new_with_indent(Vec::new(), b' ', 2);
    writer.write_event(Event::Decl(BytesDecl::new(
        "1.0",
        Some("UTF-8"),
        Some("yes"),
    )))?;

    let mut sld_master = BytesStart::new("p:sldMaster");
    sld_master.push_attribute(("xmlns:p", PRESENTATIONML_NS));
    sld_master.push_attribute(("xmlns:a", DRAWINGML_NS));
    sld_master.push_attribute(("xmlns:r", RELATIONSHIP_NS));
    if data.preserve {
        sld_master.push_attribute(("preserve", "1"));
    }
    writer.write_event(Event::Start(sld_master))?;

    // <p:cSld>
    writer.write_event(Event::Start(BytesStart::new("p:cSld")))?;

    // Background.
    if let Some(background) = data.background {
        write_background_xml(&mut writer, background)?;
    }

    // Shape tree: prefer raw roundtrip data if available.
    if let Some(raw_tree) = data.raw_sp_tree {
        raw_tree.write_to(&mut writer)?;
    } else {
        write_minimal_sp_tree(&mut writer)?;
    }

    writer.write_event(Event::End(BytesEnd::new("p:cSld")))?;

    // Color map.
    write_color_map(&mut writer, &data.color_map)?;

    // Layout ID list.
    if !data.layout_refs.is_empty() {
        writer.write_event(Event::Start(BytesStart::new("p:sldLayoutIdLst")))?;
        for layout_ref in data.layout_refs {
            let mut sld_layout_id = BytesStart::new("p:sldLayoutId");
            if let Some(ref id) = layout_ref.id {
                sld_layout_id.push_attribute(("id", id.as_str()));
            }
            sld_layout_id.push_attribute(("r:id", layout_ref.relationship_id.as_str()));
            writer.write_event(Event::Empty(sld_layout_id))?;
        }
        writer.write_event(Event::End(BytesEnd::new("p:sldLayoutIdLst")))?;
    }

    // Unknown children (roundtrip fidelity).
    for node in data.unknown_children {
        node.write_to(&mut writer)?;
    }

    writer.write_event(Event::End(BytesEnd::new("p:sldMaster")))?;
    Ok(writer.into_inner())
}

/// Data required to write a slide master XML part.
#[derive(Debug)]
pub(crate) struct WriteSlideMasterData<'a> {
    /// Whether to set the `preserve` attribute.
    pub preserve: bool,
    /// Layout references to include in `<p:sldLayoutIdLst>`.
    pub layout_refs: &'a [ParsedLayoutRef],
    /// Background fill.
    pub background: Option<&'a SlideBackground>,
    /// Raw shape tree for roundtrip (preferred over building from scratch).
    pub raw_sp_tree: Option<&'a RawXmlNode>,
    /// Color map attribute pairs.
    pub color_map: Vec<(String, String)>,
    /// Unknown children for roundtrip fidelity.
    pub unknown_children: &'a [RawXmlNode],
}

// ── Internal helpers ────────────────────────────────────────────────────────

/// Extract local name from a potentially namespace-prefixed XML name.
fn local_name(name: &[u8]) -> &[u8] {
    name.rsplit(|byte| *byte == b':').next().unwrap_or(name)
}

/// Parse an XML boolean value.
fn parse_xml_bool(value: &str) -> Option<bool> {
    match value.trim().to_ascii_lowercase().as_str() {
        "1" | "true" | "on" => Some(true),
        "0" | "false" | "off" => Some(false),
        _ => None,
    }
}

/// Get an attribute value by its local name.
fn get_attribute_value(event: &BytesStart<'_>, expected_local_name: &[u8]) -> Option<String> {
    event.attributes().flatten().find_map(|attribute| {
        (local_name(attribute.key.as_ref()) == expected_local_name)
            .then(|| String::from_utf8_lossy(attribute.value.as_ref()).into_owned())
    })
}

/// Get the `r:id` relationship attribute value.
fn get_relationship_id_attribute_value(event: &BytesStart<'_>) -> Option<String> {
    if let Some(value) = get_exact_attribute_value(event, b"r:id") {
        return Some(value);
    }
    event.attributes().flatten().find_map(|attribute| {
        let key = attribute.key.as_ref();
        (local_name(key) == b"id" && key != b"id")
            .then(|| String::from_utf8_lossy(attribute.value.as_ref()).into_owned())
    })
}

/// Get an attribute value by its exact qualified name.
fn get_exact_attribute_value(event: &BytesStart<'_>, expected_key: &[u8]) -> Option<String> {
    event.attributes().flatten().find_map(|attribute| {
        (attribute.key.as_ref() == expected_key)
            .then(|| String::from_utf8_lossy(attribute.value.as_ref()).into_owned())
    })
}

/// Skip forward past the end tag matching the given fully-qualified name bytes.
fn skip_to_end_tag<R: std::io::BufRead>(reader: &mut Reader<R>, tag_name: &[u8]) -> Result<()> {
    let mut depth = 1_usize;
    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf)? {
            Event::Start(ref e) if e.name().as_ref() == tag_name => {
                depth += 1;
            }
            Event::End(ref e) if e.name().as_ref() == tag_name => {
                depth -= 1;
                if depth == 0 {
                    break;
                }
            }
            Event::Eof => break,
            _ => {}
        }
        buf.clear();
    }
    Ok(())
}

/// Extract basic `Shape` metadata from a captured `RawXmlNode` representing `<p:spTree>`.
///
/// This does a lightweight walk of the raw tree to find `<p:sp>` children and
/// extract their name, placeholder type, and placeholder index. This avoids
/// duplicating the full 1000-line shape parser from `presentation.rs` while
/// still providing useful domain model data for masters.
fn extract_shapes_from_raw_sp_tree(sp_tree: &RawXmlNode) -> Vec<Shape> {
    let RawXmlNode::Element { children, .. } = sp_tree else {
        return Vec::new();
    };

    let mut shapes = Vec::new();
    for child in children {
        let RawXmlNode::Element {
            name,
            children: sp_children,
            ..
        } = child
        else {
            continue;
        };
        if local_name(name.as_bytes()) != b"sp" {
            continue;
        }

        // Look for nvSpPr > cNvPr (name) and nvSpPr > nvPr > ph (placeholder).
        let mut shape_name = String::new();
        let mut placeholder_kind: Option<String> = None;
        let mut placeholder_idx: Option<u32> = None;

        for sp_child in sp_children {
            let RawXmlNode::Element {
                name: sp_child_name,
                children: nv_children,
                ..
            } = sp_child
            else {
                continue;
            };
            if local_name(sp_child_name.as_bytes()) != b"nvSpPr" {
                continue;
            }

            for nv_child in nv_children {
                let RawXmlNode::Element {
                    name: nv_child_name,
                    attributes: nv_attrs,
                    children: nv_grand_children,
                    ..
                } = nv_child
                else {
                    continue;
                };
                let nv_local = local_name(nv_child_name.as_bytes());
                if nv_local == b"cNvPr" {
                    for (key, val) in nv_attrs {
                        if key == "name" {
                            shape_name = val.clone();
                        }
                    }
                } else if nv_local == b"nvPr" {
                    for nv_grand in nv_grand_children {
                        let RawXmlNode::Element {
                            name: ph_name,
                            attributes: ph_attrs,
                            ..
                        } = nv_grand
                        else {
                            continue;
                        };
                        if local_name(ph_name.as_bytes()) == b"ph" {
                            for (key, val) in ph_attrs {
                                if key == "type" {
                                    placeholder_kind = Some(val.clone());
                                } else if key == "idx" {
                                    placeholder_idx = val.parse().ok();
                                }
                            }
                        }
                    }
                }
            }
        }

        let mut shape = Shape::new(shape_name);
        if let Some(kind) = &placeholder_kind {
            shape.set_placeholder_type(PlaceholderType::from_xml(kind));
        }
        if let Some(idx) = placeholder_idx {
            shape.set_placeholder_idx(idx);
        }
        // Extract text from txBody if present.
        extract_shape_text(sp_children, &mut shape);
        shapes.push(shape);
    }

    shapes
}

/// Extract text runs from `<p:txBody>` inside shape children into a `Shape`.
fn extract_shape_text(sp_children: &[RawXmlNode], shape: &mut Shape) {
    for child in sp_children {
        let RawXmlNode::Element {
            name,
            children: tx_children,
            ..
        } = child
        else {
            continue;
        };
        if local_name(name.as_bytes()) != b"txBody" {
            continue;
        }
        for tx_child in tx_children {
            let RawXmlNode::Element {
                name: para_name,
                children: para_children,
                ..
            } = tx_child
            else {
                continue;
            };
            if local_name(para_name.as_bytes()) != b"p" {
                continue;
            }
            let paragraph = shape.add_paragraph();
            for para_child in para_children {
                let RawXmlNode::Element {
                    name: run_name,
                    children: run_children,
                    ..
                } = para_child
                else {
                    continue;
                };
                if local_name(run_name.as_bytes()) != b"r" {
                    continue;
                }
                for run_child in run_children {
                    match run_child {
                        RawXmlNode::Element {
                            name: t_name,
                            children: t_children,
                            ..
                        } if local_name(t_name.as_bytes()) == b"t" => {
                            for t_child in t_children {
                                if let RawXmlNode::Text(text) = t_child {
                                    paragraph.add_run(text.as_str());
                                }
                            }
                        }
                        _ => {}
                    }
                }
            }
        }
    }
}

/// Parse a `<p:bg>` background subtree.
///
/// Handles the most common cases: solid fill from `<a:srgbClr>`.
/// More complex backgrounds (gradient, pattern, image) are preserved via
/// the raw spTree roundtrip mechanism.
fn parse_background_subtree<R: std::io::BufRead>(
    reader: &mut Reader<R>,
) -> Result<Option<SlideBackground>> {
    let mut buf = Vec::new();
    let mut in_bg_pr = false;
    let mut in_solid_fill = false;

    loop {
        match reader.read_event_into(&mut buf)? {
            Event::Start(ref event) => {
                let name = event.name();
                let local = local_name(name.as_ref());
                match local {
                    b"bgPr" => in_bg_pr = true,
                    b"solidFill" if in_bg_pr => in_solid_fill = true,
                    b"srgbClr" if in_solid_fill => {
                        if let Some(val) = get_attribute_value(event, b"val") {
                            // Skip to end of bg and return.
                            skip_past_bg_end(reader)?;
                            return Ok(Some(SlideBackground::Solid(val)));
                        }
                    }
                    _ => {}
                }
            }
            Event::Empty(ref event) => {
                let name = event.name();
                let local = local_name(name.as_ref());
                if local == b"srgbClr" && in_solid_fill {
                    if let Some(val) = get_attribute_value(event, b"val") {
                        skip_past_bg_end(reader)?;
                        return Ok(Some(SlideBackground::Solid(val)));
                    }
                }
            }
            Event::End(ref event) => {
                let name = event.name();
                let local = local_name(name.as_ref());
                if local == b"bg" {
                    return Ok(None);
                }
                if local == b"solidFill" {
                    in_solid_fill = false;
                }
                if local == b"bgPr" {
                    in_bg_pr = false;
                }
            }
            Event::Eof => break,
            _ => {}
        }
        buf.clear();
    }
    Ok(None)
}

/// Skip forward past the `</p:bg>` end tag.
fn skip_past_bg_end<R: std::io::BufRead>(reader: &mut Reader<R>) -> Result<()> {
    let mut buf = Vec::new();
    let mut depth = 1_usize; // We're inside <p:bg>, need to get past its end.
    loop {
        match reader.read_event_into(&mut buf)? {
            Event::Start(ref e) if local_name(e.name().as_ref()) == b"bg" => {
                depth += 1;
            }
            Event::End(ref e) if local_name(e.name().as_ref()) == b"bg" => {
                depth -= 1;
                if depth == 0 {
                    break;
                }
            }
            Event::Eof => break,
            _ => {}
        }
        buf.clear();
    }
    Ok(())
}

/// Parse an inline `<a:clrScheme>` subtree into a `ThemeColorScheme`.
fn parse_theme_color_scheme_subtree<R: std::io::BufRead>(
    reader: &mut Reader<R>,
    start_event: &BytesStart<'_>,
) -> Result<Option<ThemeColorScheme>> {
    let mut scheme = ThemeColorScheme::new();
    scheme.name = get_attribute_value(start_event, b"name");

    let mut buf = Vec::new();
    let mut current_color_name: Option<String> = None;

    loop {
        match reader.read_event_into(&mut buf)? {
            Event::Start(ref event) => {
                let name = event.name();
                let local = local_name(name.as_ref());
                let local_str = String::from_utf8_lossy(local).into_owned();
                match local_str.as_str() {
                    "dk1" | "lt1" | "dk2" | "lt2" | "accent1" | "accent2" | "accent3"
                    | "accent4" | "accent5" | "accent6" | "hlink" | "folHlink" => {
                        current_color_name = Some(local_str);
                    }
                    "srgbClr" | "sysClr" => {
                        if let Some(ref color_name) = current_color_name {
                            let color_val = if local_str == "sysClr" {
                                get_attribute_value(event, b"lastClr")
                                    .or_else(|| get_attribute_value(event, b"val"))
                            } else {
                                get_attribute_value(event, b"val")
                            };
                            if let Some(val) = color_val {
                                scheme.set_color_by_name(color_name, val);
                            }
                        }
                    }
                    _ => {}
                }
            }
            Event::Empty(ref event) => {
                let name = event.name();
                let local = local_name(name.as_ref());
                let local_str = String::from_utf8_lossy(local).into_owned();
                match local_str.as_str() {
                    "srgbClr" | "sysClr" => {
                        if let Some(ref color_name) = current_color_name {
                            let color_val = if local_str == "sysClr" {
                                get_attribute_value(event, b"lastClr")
                                    .or_else(|| get_attribute_value(event, b"val"))
                            } else {
                                get_attribute_value(event, b"val")
                            };
                            if let Some(val) = color_val {
                                scheme.set_color_by_name(color_name, val);
                            }
                        }
                    }
                    _ => {}
                }
            }
            Event::End(ref event) => {
                let name = event.name();
                let local = local_name(name.as_ref());
                if local == b"clrScheme" {
                    return Ok(Some(scheme));
                }
                let local_str = String::from_utf8_lossy(local).into_owned();
                match local_str.as_str() {
                    "dk1" | "lt1" | "dk2" | "lt2" | "accent1" | "accent2" | "accent3"
                    | "accent4" | "accent5" | "accent6" | "hlink" | "folHlink" => {
                        current_color_name = None;
                    }
                    _ => {}
                }
            }
            Event::Eof => break,
            _ => {}
        }
        buf.clear();
    }

    Ok(None)
}

/// Write a `<p:bg>` background element.
fn write_background_xml<W: std::io::Write>(
    writer: &mut Writer<W>,
    background: &SlideBackground,
) -> Result<()> {
    writer.write_event(Event::Start(BytesStart::new("p:bg")))?;
    writer.write_event(Event::Start(BytesStart::new("p:bgPr")))?;

    match background {
        SlideBackground::Solid(color) => {
            writer.write_event(Event::Start(BytesStart::new("a:solidFill")))?;
            let mut srgb_clr = BytesStart::new("a:srgbClr");
            srgb_clr.push_attribute(("val", color.as_str()));
            writer.write_event(Event::Empty(srgb_clr))?;
            writer.write_event(Event::End(BytesEnd::new("a:solidFill")))?;
        }
        SlideBackground::Gradient(gradient) => {
            // Write gradient fill: stops + linear angle.
            let mut grad_fill = BytesStart::new("a:gradFill");
            if gradient.fill_type == Some(crate::shape::GradientFillType::Path) {
                grad_fill.push_attribute(("flip", "none"));
                grad_fill.push_attribute(("rotWithShape", "1"));
            }
            writer.write_event(Event::Start(grad_fill))?;
            writer.write_event(Event::Start(BytesStart::new("a:gsLst")))?;
            for stop in &gradient.stops {
                let mut gs = BytesStart::new("a:gs");
                let pos_str = stop.position.to_string();
                gs.push_attribute(("pos", pos_str.as_str()));
                writer.write_event(Event::Start(gs))?;
                let mut srgb = BytesStart::new("a:srgbClr");
                srgb.push_attribute(("val", stop.color_srgb.as_str()));
                writer.write_event(Event::Empty(srgb))?;
                writer.write_event(Event::End(BytesEnd::new("a:gs")))?;
            }
            writer.write_event(Event::End(BytesEnd::new("a:gsLst")))?;
            if let Some(angle) = gradient.linear_angle {
                let mut lin = BytesStart::new("a:lin");
                let ang_str = angle.to_string();
                lin.push_attribute(("ang", ang_str.as_str()));
                lin.push_attribute(("scaled", "1"));
                writer.write_event(Event::Empty(lin))?;
            }
            writer.write_event(Event::End(BytesEnd::new("a:gradFill")))?;
        }
        SlideBackground::Pattern {
            pattern_type,
            foreground_color,
            background_color,
        } => {
            let mut patt_fill = BytesStart::new("a:pattFill");
            patt_fill.push_attribute(("prst", pattern_type.as_str()));
            writer.write_event(Event::Start(patt_fill))?;
            writer.write_event(Event::Start(BytesStart::new("a:fgClr")))?;
            let mut fg_clr = BytesStart::new("a:srgbClr");
            fg_clr.push_attribute(("val", foreground_color.as_str()));
            writer.write_event(Event::Empty(fg_clr))?;
            writer.write_event(Event::End(BytesEnd::new("a:fgClr")))?;
            writer.write_event(Event::Start(BytesStart::new("a:bgClr")))?;
            let mut bg_clr = BytesStart::new("a:srgbClr");
            bg_clr.push_attribute(("val", background_color.as_str()));
            writer.write_event(Event::Empty(bg_clr))?;
            writer.write_event(Event::End(BytesEnd::new("a:bgClr")))?;
            writer.write_event(Event::End(BytesEnd::new("a:pattFill")))?;
        }
        SlideBackground::Image { relationship_id } => {
            writer.write_event(Event::Start(BytesStart::new("a:blipFill")))?;
            let mut blip = BytesStart::new("a:blip");
            blip.push_attribute(("r:embed", relationship_id.as_str()));
            writer.write_event(Event::Empty(blip))?;
            writer.write_event(Event::Start(BytesStart::new("a:stretch")))?;
            writer.write_event(Event::Empty(BytesStart::new("a:fillRect")))?;
            writer.write_event(Event::End(BytesEnd::new("a:stretch")))?;
            writer.write_event(Event::End(BytesEnd::new("a:blipFill")))?;
        }
    }

    writer.write_event(Event::Empty(BytesStart::new("a:effectLst")))?;
    writer.write_event(Event::End(BytesEnd::new("p:bgPr")))?;
    writer.write_event(Event::End(BytesEnd::new("p:bg")))?;
    Ok(())
}

/// Write a minimal empty `<p:spTree/>`.
fn write_minimal_sp_tree<W: std::io::Write>(writer: &mut Writer<W>) -> Result<()> {
    writer.write_event(Event::Empty(BytesStart::new("p:spTree")))?;
    Ok(())
}

/// Write the `<p:clrMap>` element.
fn write_color_map<W: std::io::Write>(
    writer: &mut Writer<W>,
    attrs: &[(String, String)],
) -> Result<()> {
    let mut clr_map = BytesStart::new("p:clrMap");
    if attrs.is_empty() {
        // Default OOXML color map.
        clr_map.push_attribute(("bg1", "lt1"));
        clr_map.push_attribute(("tx1", "dk1"));
        clr_map.push_attribute(("bg2", "lt2"));
        clr_map.push_attribute(("tx2", "dk2"));
        clr_map.push_attribute(("accent1", "accent1"));
        clr_map.push_attribute(("accent2", "accent2"));
        clr_map.push_attribute(("accent3", "accent3"));
        clr_map.push_attribute(("accent4", "accent4"));
        clr_map.push_attribute(("accent5", "accent5"));
        clr_map.push_attribute(("accent6", "accent6"));
        clr_map.push_attribute(("hlink", "hlink"));
        clr_map.push_attribute(("folHlink", "folHlink"));
    } else {
        for (key, val) in attrs {
            clr_map.push_attribute((key.as_str(), val.as_str()));
        }
    }
    writer.write_event(Event::Empty(clr_map))?;
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn parse_minimal_slide_master() {
        let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldMaster xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
             xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
             xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <p:cSld>
    <p:spTree/>
  </p:cSld>
  <p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2"
            accent1="accent1" accent2="accent2" accent3="accent3"
            accent4="accent4" accent5="accent5" accent6="accent6"
            hlink="hlink" folHlink="folHlink"/>
  <p:sldLayoutIdLst>
    <p:sldLayoutId id="2147483649" r:id="rId1"/>
  </p:sldLayoutIdLst>
</p:sldMaster>"#;

        let parsed = parse_slide_master_xml(xml.as_bytes()).expect("should parse");
        assert_eq!(parsed.preserve, None);
        assert_eq!(parsed.layout_refs.len(), 1);
        assert_eq!(parsed.layout_refs[0].relationship_id, "rId1");
        assert_eq!(parsed.layout_refs[0].id, Some("2147483649".to_string()));
        assert_eq!(parsed.color_map.len(), 12);
        assert!(parsed.shapes.is_empty());
        assert!(parsed.background.is_none());
        assert!(parsed.unknown_children.is_empty());
        assert!(parsed.raw_sp_tree.is_some());
    }

    #[test]
    fn parse_slide_master_with_preserve() {
        let xml = r#"<p:sldMaster xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
             xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
             xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
             preserve="1">
  <p:cSld><p:spTree/></p:cSld>
  <p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2"
            accent1="accent1" accent2="accent2" accent3="accent3"
            accent4="accent4" accent5="accent5" accent6="accent6"
            hlink="hlink" folHlink="folHlink"/>
  <p:sldLayoutIdLst>
    <p:sldLayoutId id="100" r:id="rId2"/>
    <p:sldLayoutId id="101" r:id="rId3"/>
  </p:sldLayoutIdLst>
</p:sldMaster>"#;

        let parsed = parse_slide_master_xml(xml.as_bytes()).expect("should parse");
        assert_eq!(parsed.preserve, Some(true));
        assert_eq!(parsed.layout_refs.len(), 2);
        assert_eq!(parsed.layout_refs[0].relationship_id, "rId2");
        assert_eq!(parsed.layout_refs[1].relationship_id, "rId3");
    }

    #[test]
    fn parse_slide_master_with_shapes() {
        let xml = r#"<p:sldMaster xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
             xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
             xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <p:cSld>
    <p:spTree>
      <p:sp>
        <p:nvSpPr>
          <p:cNvPr id="2" name="Title Placeholder 1"/>
          <p:cNvSpPr/>
          <p:nvPr>
            <p:ph type="title"/>
          </p:nvPr>
        </p:nvSpPr>
        <p:txBody>
          <a:bodyPr/>
          <a:lstStyle/>
          <a:p><a:r><a:t>Click to edit title</a:t></a:r></a:p>
        </p:txBody>
      </p:sp>
    </p:spTree>
  </p:cSld>
  <p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2"
            accent1="accent1" accent2="accent2" accent3="accent3"
            accent4="accent4" accent5="accent5" accent6="accent6"
            hlink="hlink" folHlink="folHlink"/>
  <p:sldLayoutIdLst>
    <p:sldLayoutId id="2147483649" r:id="rId5"/>
  </p:sldLayoutIdLst>
</p:sldMaster>"#;

        let parsed = parse_slide_master_xml(xml.as_bytes()).expect("should parse");
        assert_eq!(parsed.shapes.len(), 1);
        assert_eq!(parsed.shapes[0].name(), "Title Placeholder 1");
        assert_eq!(parsed.shapes[0].placeholder_kind(), Some("title"));
        assert_eq!(
            parsed.shapes[0].paragraphs()[0].runs()[0].text(),
            "Click to edit title"
        );
    }

    #[test]
    fn parse_slide_master_with_background() {
        let xml = r#"<p:sldMaster xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
             xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
             xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <p:cSld>
    <p:bg>
      <p:bgPr>
        <a:solidFill>
          <a:srgbClr val="FF0000"/>
        </a:solidFill>
        <a:effectLst/>
      </p:bgPr>
    </p:bg>
    <p:spTree/>
  </p:cSld>
  <p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2"
            accent1="accent1" accent2="accent2" accent3="accent3"
            accent4="accent4" accent5="accent5" accent6="accent6"
            hlink="hlink" folHlink="folHlink"/>
  <p:sldLayoutIdLst>
    <p:sldLayoutId id="100" r:id="rId1"/>
  </p:sldLayoutIdLst>
</p:sldMaster>"#;

        let parsed = parse_slide_master_xml(xml.as_bytes()).expect("should parse");
        assert_eq!(
            parsed.background,
            Some(SlideBackground::Solid("FF0000".to_string()))
        );
    }

    #[test]
    fn parse_preserves_unknown_children() {
        let xml = r#"<p:sldMaster xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
             xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
             xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <p:cSld><p:spTree/></p:cSld>
  <p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2"
            accent1="accent1" accent2="accent2" accent3="accent3"
            accent4="accent4" accent5="accent5" accent6="accent6"
            hlink="hlink" folHlink="folHlink"/>
  <p:sldLayoutIdLst>
    <p:sldLayoutId id="100" r:id="rId1"/>
  </p:sldLayoutIdLst>
  <p:extLst>
    <p:ext uri="{some-unknown-extension}">
      <custom:data xmlns:custom="http://example.com">hello</custom:data>
    </p:ext>
  </p:extLst>
</p:sldMaster>"#;

        let parsed = parse_slide_master_xml(xml.as_bytes()).expect("should parse");
        assert_eq!(parsed.unknown_children.len(), 1);
        match &parsed.unknown_children[0] {
            RawXmlNode::Element { name, .. } => {
                assert_eq!(name, "p:extLst");
            }
            other => panic!("expected Element, got {other:?}"),
        }
    }

    #[test]
    fn parse_inline_theme_color_scheme() {
        let xml = r#"<p:sldMaster xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
             xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
             xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <p:cSld><p:spTree/></p:cSld>
  <p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2"
            accent1="accent1" accent2="accent2" accent3="accent3"
            accent4="accent4" accent5="accent5" accent6="accent6"
            hlink="hlink" folHlink="folHlink"/>
  <a:clrScheme name="Office">
    <a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1>
    <a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1>
    <a:accent1><a:srgbClr val="4472C4"/></a:accent1>
  </a:clrScheme>
  <p:sldLayoutIdLst>
    <p:sldLayoutId id="100" r:id="rId1"/>
  </p:sldLayoutIdLst>
</p:sldMaster>"#;

        let parsed = parse_slide_master_xml(xml.as_bytes()).expect("should parse");
        let theme = parsed.theme.expect("should have theme");
        assert_eq!(theme.name, Some("Office".to_string()));
        assert_eq!(theme.dark1.as_deref(), Some("000000"));
        assert_eq!(theme.light1.as_deref(), Some("FFFFFF"));
        assert_eq!(theme.accent1.as_deref(), Some("4472C4"));
    }

    #[test]
    fn write_minimal_slide_master() {
        let data = WriteSlideMasterData {
            preserve: false,
            layout_refs: &[ParsedLayoutRef {
                id: Some("2147483649".to_string()),
                relationship_id: "rId1".to_string(),
            }],
            background: None,
            raw_sp_tree: None,
            color_map: vec![],
            unknown_children: &[],
        };

        let bytes = write_slide_master_xml(&data).expect("should write");
        let xml = String::from_utf8(bytes).expect("valid UTF-8");

        assert!(xml.contains("p:sldMaster"));
        assert!(xml.contains("p:cSld"));
        assert!(xml.contains("p:spTree"));
        assert!(xml.contains("p:clrMap"));
        assert!(xml.contains("p:sldLayoutIdLst"));
        assert!(xml.contains("r:id=\"rId1\""));
        assert!(xml.contains("bg1=\"lt1\""));
        assert!(!xml.contains("preserve="));
    }

    #[test]
    fn write_slide_master_with_preserve() {
        let data = WriteSlideMasterData {
            preserve: true,
            layout_refs: &[],
            background: None,
            raw_sp_tree: None,
            color_map: vec![("bg1".to_string(), "dk1".to_string())],
            unknown_children: &[],
        };

        let bytes = write_slide_master_xml(&data).expect("should write");
        let xml = String::from_utf8(bytes).expect("valid UTF-8");

        assert!(xml.contains("preserve=\"1\""));
        assert!(xml.contains("bg1=\"dk1\""));
    }

    #[test]
    fn write_with_background() {
        let bg = SlideBackground::Solid("4472C4".to_string());
        let data = WriteSlideMasterData {
            preserve: false,
            layout_refs: &[ParsedLayoutRef {
                id: Some("100".to_string()),
                relationship_id: "rId1".to_string(),
            }],
            background: Some(&bg),
            raw_sp_tree: None,
            color_map: vec![],
            unknown_children: &[],
        };

        let bytes = write_slide_master_xml(&data).expect("should write");
        let xml = String::from_utf8(bytes).expect("valid UTF-8");

        assert!(xml.contains("p:bg"));
        assert!(xml.contains("a:solidFill"));
        assert!(xml.contains("4472C4"));
    }

    #[test]
    fn write_preserves_unknown_children() {
        let unknown = vec![RawXmlNode::Element {
            name: "p:extLst".to_string(),
            attributes: vec![],
            children: vec![RawXmlNode::Element {
                name: "p:ext".to_string(),
                attributes: vec![("uri".to_string(), "{test-guid}".to_string())],
                children: vec![],
            }],
        }];

        let data = WriteSlideMasterData {
            preserve: false,
            layout_refs: &[],
            background: None,
            raw_sp_tree: None,
            color_map: vec![],
            unknown_children: &unknown,
        };

        let bytes = write_slide_master_xml(&data).expect("should write");
        let xml = String::from_utf8(bytes).expect("valid UTF-8");

        assert!(xml.contains("p:extLst"));
        assert!(xml.contains("{test-guid}"));
    }

    #[test]
    fn write_preserves_raw_sp_tree() {
        let raw_tree = RawXmlNode::Element {
            name: "p:spTree".to_string(),
            attributes: vec![],
            children: vec![RawXmlNode::Element {
                name: "p:sp".to_string(),
                attributes: vec![],
                children: vec![RawXmlNode::Element {
                    name: "p:nvSpPr".to_string(),
                    attributes: vec![],
                    children: vec![RawXmlNode::Element {
                        name: "p:cNvPr".to_string(),
                        attributes: vec![
                            ("id".to_string(), "2".to_string()),
                            ("name".to_string(), "My Shape".to_string()),
                        ],
                        children: vec![],
                    }],
                }],
            }],
        };

        let data = WriteSlideMasterData {
            preserve: false,
            layout_refs: &[],
            background: None,
            raw_sp_tree: Some(&raw_tree),
            color_map: vec![],
            unknown_children: &[],
        };

        let bytes = write_slide_master_xml(&data).expect("should write");
        let xml = String::from_utf8(bytes).expect("valid UTF-8");

        assert!(xml.contains("My Shape"));
        assert!(xml.contains("p:spTree"));
        assert!(xml.contains("p:nvSpPr"));
    }

    #[test]
    fn parse_write_roundtrip_preserves_layout_refs() {
        let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldMaster xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
             xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
             xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <p:cSld>
    <p:spTree/>
  </p:cSld>
  <p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2"
            accent1="accent1" accent2="accent2" accent3="accent3"
            accent4="accent4" accent5="accent5" accent6="accent6"
            hlink="hlink" folHlink="folHlink"/>
  <p:sldLayoutIdLst>
    <p:sldLayoutId id="100" r:id="rId1"/>
    <p:sldLayoutId id="101" r:id="rId2"/>
    <p:sldLayoutId id="102" r:id="rId3"/>
  </p:sldLayoutIdLst>
</p:sldMaster>"#;

        let parsed = parse_slide_master_xml(xml.as_bytes()).expect("should parse");

        let data = WriteSlideMasterData {
            preserve: parsed.preserve.unwrap_or(false),
            layout_refs: &parsed.layout_refs,
            background: parsed.background.as_ref(),
            raw_sp_tree: parsed.raw_sp_tree.as_ref(),
            color_map: parsed.color_map.clone(),
            unknown_children: &parsed.unknown_children,
        };

        let written_bytes = write_slide_master_xml(&data).expect("should write");
        let reparsed = parse_slide_master_xml(&written_bytes).expect("should re-parse");

        assert_eq!(reparsed.layout_refs.len(), 3);
        assert_eq!(reparsed.layout_refs[0].relationship_id, "rId1");
        assert_eq!(reparsed.layout_refs[1].relationship_id, "rId2");
        assert_eq!(reparsed.layout_refs[2].relationship_id, "rId3");
        assert_eq!(reparsed.color_map.len(), parsed.color_map.len());
    }

    #[test]
    fn parse_write_roundtrip_with_shapes_and_background() {
        let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldMaster xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
             xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
             xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <p:cSld>
    <p:bg>
      <p:bgPr>
        <a:solidFill>
          <a:srgbClr val="0066CC"/>
        </a:solidFill>
        <a:effectLst/>
      </p:bgPr>
    </p:bg>
    <p:spTree>
      <p:sp>
        <p:nvSpPr>
          <p:cNvPr id="2" name="Master Title"/>
          <p:cNvSpPr/>
          <p:nvPr>
            <p:ph type="title"/>
          </p:nvPr>
        </p:nvSpPr>
        <p:txBody>
          <a:bodyPr/>
          <a:p><a:r><a:t>Hello</a:t></a:r></a:p>
        </p:txBody>
      </p:sp>
    </p:spTree>
  </p:cSld>
  <p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2"
            accent1="accent1" accent2="accent2" accent3="accent3"
            accent4="accent4" accent5="accent5" accent6="accent6"
            hlink="hlink" folHlink="folHlink"/>
  <p:sldLayoutIdLst>
    <p:sldLayoutId id="100" r:id="rId1"/>
  </p:sldLayoutIdLst>
</p:sldMaster>"#;

        let parsed = parse_slide_master_xml(xml.as_bytes()).expect("should parse");

        // Verify parsed data.
        assert_eq!(
            parsed.background,
            Some(SlideBackground::Solid("0066CC".to_string()))
        );
        assert_eq!(parsed.shapes.len(), 1);
        assert_eq!(parsed.shapes[0].name(), "Master Title");

        // Write and re-parse.
        let data = WriteSlideMasterData {
            preserve: false,
            layout_refs: &parsed.layout_refs,
            background: parsed.background.as_ref(),
            raw_sp_tree: parsed.raw_sp_tree.as_ref(),
            color_map: parsed.color_map.clone(),
            unknown_children: &parsed.unknown_children,
        };

        let written_bytes = write_slide_master_xml(&data).expect("should write");
        let reparsed = parse_slide_master_xml(&written_bytes).expect("should re-parse");

        // The raw spTree roundtrip preserves shape structure.
        assert_eq!(reparsed.shapes.len(), 1);
        assert_eq!(reparsed.shapes[0].name(), "Master Title");
        assert_eq!(reparsed.shapes[0].placeholder_kind(), Some("title"));
    }

    #[test]
    fn parse_multiple_shapes_with_placeholders() {
        let xml = r#"<p:sldMaster xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
             xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
             xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <p:cSld>
    <p:spTree>
      <p:sp>
        <p:nvSpPr>
          <p:cNvPr id="2" name="Title 1"/>
          <p:cNvSpPr/>
          <p:nvPr><p:ph type="title"/></p:nvPr>
        </p:nvSpPr>
        <p:txBody><a:bodyPr/><a:p><a:r><a:t>Title text</a:t></a:r></a:p></p:txBody>
      </p:sp>
      <p:sp>
        <p:nvSpPr>
          <p:cNvPr id="3" name="Body 1"/>
          <p:cNvSpPr/>
          <p:nvPr><p:ph type="body" idx="1"/></p:nvPr>
        </p:nvSpPr>
        <p:txBody><a:bodyPr/><a:p><a:r><a:t>Body text</a:t></a:r></a:p></p:txBody>
      </p:sp>
      <p:sp>
        <p:nvSpPr>
          <p:cNvPr id="4" name="Footer 1"/>
          <p:cNvSpPr/>
          <p:nvPr><p:ph type="ftr" idx="10"/></p:nvPr>
        </p:nvSpPr>
        <p:txBody><a:bodyPr/><a:p><a:r><a:t>Footer</a:t></a:r></a:p></p:txBody>
      </p:sp>
    </p:spTree>
  </p:cSld>
  <p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2"
            accent1="accent1" accent2="accent2" accent3="accent3"
            accent4="accent4" accent5="accent5" accent6="accent6"
            hlink="hlink" folHlink="folHlink"/>
  <p:sldLayoutIdLst>
    <p:sldLayoutId id="100" r:id="rId1"/>
  </p:sldLayoutIdLst>
</p:sldMaster>"#;

        let parsed = parse_slide_master_xml(xml.as_bytes()).expect("should parse");
        assert_eq!(parsed.shapes.len(), 3);

        assert_eq!(parsed.shapes[0].name(), "Title 1");
        assert_eq!(parsed.shapes[0].placeholder_kind(), Some("title"));

        assert_eq!(parsed.shapes[1].name(), "Body 1");
        assert_eq!(parsed.shapes[1].placeholder_kind(), Some("body"));
        assert_eq!(parsed.shapes[1].placeholder_idx(), Some(1));

        assert_eq!(parsed.shapes[2].name(), "Footer 1");
        assert_eq!(parsed.shapes[2].placeholder_kind(), Some("ftr"));
        assert_eq!(parsed.shapes[2].placeholder_idx(), Some(10));

        // Verify text extraction.
        assert_eq!(
            parsed.shapes[0].paragraphs()[0].runs()[0].text(),
            "Title text"
        );
        assert_eq!(
            parsed.shapes[1].paragraphs()[0].runs()[0].text(),
            "Body text"
        );
        assert_eq!(parsed.shapes[2].paragraphs()[0].runs()[0].text(), "Footer");
    }
}
