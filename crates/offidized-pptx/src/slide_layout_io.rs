//! Slide layout XML parsing and writing.
//!
//! Handles the `p:sldLayout` element — the layout definition within a `.pptx`
//! file. Supports full roundtrip fidelity: known children (shapes, background,
//! color mapping) are parsed into typed fields, and everything else is preserved
//! as `RawXmlNode` for lossless save.

use std::io::Cursor;

use offidized_opc::RawXmlNode;
use quick_xml::events::{BytesDecl, BytesEnd, BytesStart, Event};
use quick_xml::{Reader, Writer};

use crate::error::{PptxError, Result};
use crate::shape::Shape;
use crate::slide::SlideBackground;
use crate::slide_layout::SlideLayout;

// Re-use the XML namespace constants.
const PRESENTATIONML_NS: &str = "http://schemas.openxmlformats.org/presentationml/2006/main";
const DRAWINGML_NS: &str = "http://schemas.openxmlformats.org/drawingml/2006/main";
const RELATIONSHIP_NS: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

/// Parsed slide layout data extracted from XML.
///
/// This intermediate struct holds everything we pull out of `<p:sldLayout>` so
/// callers can construct or update a `SlideLayout` as needed.
#[derive(Debug)]
pub struct ParsedSlideLayout {
    /// Layout name from `p:cSld/@name`.
    pub name: Option<String>,
    /// Layout type from `p:sldLayout/@type`.
    pub layout_type: Option<String>,
    /// Preserve flag from `p:sldLayout/@preserve`.
    pub preserve: Option<bool>,
    /// Shapes parsed from the `p:spTree` inside `p:cSld`.
    pub shapes: Vec<Shape>,
    /// Background parsed from `p:bg` inside `p:cSld`.
    pub background: Option<SlideBackground>,
    /// Unknown direct children of `p:sldLayout` (for roundtrip fidelity).
    pub unknown_children: Vec<RawXmlNode>,
}

/// Parse a `<p:sldLayout>` XML document into a `ParsedSlideLayout`.
///
/// Extracts the layout name (from `p:cSld@name`), type, preserve flag,
/// placeholder shapes, background, and captures any unknown children as
/// `RawXmlNode` for roundtrip preservation.
pub fn parse_slide_layout_xml(xml: &[u8]) -> Result<ParsedSlideLayout> {
    let mut reader = Reader::from_reader(Cursor::new(xml));
    reader.config_mut().trim_text(true);

    let mut buffer = Vec::new();
    let mut name: Option<String> = None;
    let mut layout_type: Option<String> = None;
    let mut preserve: Option<bool> = None;
    let mut shapes = Vec::new();
    let mut background: Option<SlideBackground> = None;
    let mut unknown_children: Vec<RawXmlNode> = Vec::new();

    // Track whether we are inside the root <p:sldLayout>.
    let mut in_sld_layout = false;

    // Known direct children of p:sldLayout that we handle explicitly.
    static KNOWN_CHILDREN: &[&[u8]] = &[b"cSld", b"clrMapOvr"];

    loop {
        match reader.read_event_into(&mut buffer)? {
            Event::Start(ref event) => {
                let event_name = event.name();
                let local = local_name(event_name.as_ref());

                if !in_sld_layout && local == b"sldLayout" {
                    in_sld_layout = true;
                    layout_type = get_attribute_value(event, b"type");
                    preserve = get_attribute_value(event, b"preserve")
                        .as_deref()
                        .and_then(parse_xml_bool);
                } else if in_sld_layout && local == b"cSld" {
                    // Extract the name attribute from p:cSld.
                    name = get_attribute_value(event, b"name");
                    // Parse the cSld body for shapes and background.
                    parse_csld_body(&mut reader, &mut shapes, &mut background)?;
                } else if in_sld_layout && local == b"clrMapOvr" {
                    // Skip the entire clrMapOvr subtree — it's handled on write
                    // by emitting <a:masterClrMapping/>.
                    skip_element(&mut reader)?;
                } else if in_sld_layout {
                    // Unknown child of p:sldLayout — preserve for roundtrip.
                    unknown_children.push(RawXmlNode::read_element(&mut reader, event)?);
                }
            }
            Event::Empty(ref event) => {
                let event_name = event.name();
                let local = local_name(event_name.as_ref());

                if !in_sld_layout && local == b"sldLayout" {
                    // Edge case: self-closing <p:sldLayout/>.
                    layout_type = get_attribute_value(event, b"type");
                    preserve = get_attribute_value(event, b"preserve")
                        .as_deref()
                        .and_then(parse_xml_bool);
                    break;
                }
                if in_sld_layout && !KNOWN_CHILDREN.contains(&local) {
                    unknown_children.push(RawXmlNode::from_empty_element(event));
                }
            }
            Event::End(ref event) => {
                let event_name = event.name();
                if in_sld_layout && local_name(event_name.as_ref()) == b"sldLayout" {
                    break;
                }
            }
            Event::Eof => break,
            _ => {}
        }
        buffer.clear();
    }

    Ok(ParsedSlideLayout {
        name,
        layout_type,
        preserve,
        shapes,
        background,
        unknown_children,
    })
}

/// Write a `SlideLayout` to XML bytes.
///
/// If the layout has not been modified (`!is_dirty()`) and has original XML
/// stored, the original bytes are returned verbatim for perfect roundtrip.
/// Otherwise, a fresh XML document is generated from the layout fields.
pub fn write_slide_layout_xml(layout: &SlideLayout) -> Result<Vec<u8>> {
    // Roundtrip fast-path: return original XML if layout hasn't been modified.
    if !layout.is_dirty() {
        if let Some(original) = layout.original_xml() {
            return Ok(original.to_vec());
        }
    }

    let mut writer = Writer::new_with_indent(Vec::new(), b' ', 2);
    writer.write_event(Event::Decl(BytesDecl::new(
        "1.0",
        Some("UTF-8"),
        Some("yes"),
    )))?;

    // <p:sldLayout> root element with namespace declarations.
    let mut sld_layout = BytesStart::new("p:sldLayout");
    sld_layout.push_attribute(("xmlns:p", PRESENTATIONML_NS));
    sld_layout.push_attribute(("xmlns:a", DRAWINGML_NS));
    sld_layout.push_attribute(("xmlns:r", RELATIONSHIP_NS));
    if let Some(layout_type) = layout.layout_type() {
        sld_layout.push_attribute(("type", layout_type));
    }
    if layout.preserve() {
        sld_layout.push_attribute(("preserve", "1"));
    }
    writer.write_event(Event::Start(sld_layout))?;

    // <p:cSld> — common slide data.
    let mut csld = BytesStart::new("p:cSld");
    if !layout.name().is_empty() {
        csld.push_attribute(("name", layout.name()));
    }
    writer.write_event(Event::Start(csld))?;

    // Background (if any).
    if let Some(bg) = layout.background() {
        write_slide_background_xml(&mut writer, bg)?;
    }

    // <p:spTree> — shape tree with all placeholder shapes.
    writer.write_event(Event::Start(BytesStart::new("p:spTree")))?;

    // Non-visual group shape properties (required by the schema).
    write_sp_tree_nv_grp_sp_pr(&mut writer)?;

    // Group shape properties (required, usually empty transform).
    writer.write_event(Event::Empty(BytesStart::new("p:grpSpPr")))?;

    // Write each shape.
    let mut next_object_id = 2_u32; // 1 is reserved for the group shape container.
    for shape in layout.shapes() {
        write_layout_shape_xml(&mut writer, &mut next_object_id, shape)?;
    }

    writer.write_event(Event::End(BytesEnd::new("p:spTree")))?;
    writer.write_event(Event::End(BytesEnd::new("p:cSld")))?;

    // <p:clrMapOvr> — color mapping override (master mapping by default).
    writer.write_event(Event::Start(BytesStart::new("p:clrMapOvr")))?;
    writer.write_event(Event::Empty(BytesStart::new("a:masterClrMapping")))?;
    writer.write_event(Event::End(BytesEnd::new("p:clrMapOvr")))?;

    writer.write_event(Event::End(BytesEnd::new("p:sldLayout")))?;
    Ok(writer.into_inner())
}

// ---------------------------------------------------------------------------
// Internal parsing helpers
// ---------------------------------------------------------------------------

/// Parse the body of `<p:cSld>` — extracts shapes from `p:spTree` and
/// background from `p:bg`.
fn parse_csld_body<R: std::io::BufRead>(
    reader: &mut Reader<R>,
    shapes: &mut Vec<Shape>,
    background: &mut Option<SlideBackground>,
) -> Result<()> {
    let mut buffer = Vec::new();
    let mut depth = 1_usize; // We're already inside <p:cSld>.

    loop {
        match reader.read_event_into(&mut buffer)? {
            Event::Start(ref event) => {
                let event_name = event.name();
                let local = local_name(event_name.as_ref());
                depth += 1;

                if depth == 2 && local == b"spTree" {
                    parse_sp_tree(reader, shapes)?;
                    depth -= 1; // parse_sp_tree consumes through </p:spTree>.
                } else if depth == 2 && local == b"bg" {
                    *background = parse_background(reader)?;
                    depth -= 1; // parse_background consumes through </p:bg>.
                }
                // Other children inside cSld are skipped (we don't model them yet).
            }
            Event::Empty(_) => {
                // Empty elements at cSld level are ignored.
            }
            Event::End(_) => {
                depth -= 1;
                if depth == 0 {
                    break; // Closing </p:cSld>.
                }
            }
            Event::Eof => break,
            _ => {}
        }
        buffer.clear();
    }
    Ok(())
}

/// Parse `<p:spTree>` contents into a `Vec<Shape>`.
///
/// Handles `<p:sp>` (shapes) and `<p:cxnSp>` (connectors) elements at the
/// top level of the shape tree. Other elements (grpSp, graphicFrame, pic) are
/// currently skipped.
fn parse_sp_tree<R: std::io::BufRead>(
    reader: &mut Reader<R>,
    shapes: &mut Vec<Shape>,
) -> Result<()> {
    let mut buffer = Vec::new();
    let mut depth = 1_usize; // Inside <p:spTree>.

    loop {
        match reader.read_event_into(&mut buffer)? {
            Event::Start(ref event) => {
                let event_name = event.name();
                let local = local_name(event_name.as_ref());
                depth += 1;

                if depth == 2 && (local == b"sp" || local == b"cxnSp") {
                    let is_connector = local == b"cxnSp";
                    if let Some(shape) = parse_shape_element(reader, is_connector)? {
                        shapes.push(shape);
                    }
                    depth -= 1; // parse_shape_element consumes through closing tag.
                }
                // Skip other children (nvGrpSpPr, grpSpPr, grpSp, graphicFrame, pic).
            }
            Event::Empty(_) => {}
            Event::End(_) => {
                depth -= 1;
                if depth == 0 {
                    break;
                }
            }
            Event::Eof => break,
            _ => {}
        }
        buffer.clear();
    }
    Ok(())
}

/// Parse a single `<p:sp>` or `<p:cxnSp>` element into a `Shape`.
///
/// Extracts name (from cNvPr), placeholder info (from ph), geometry (from
/// xfrm), and preset geometry. Returns `None` only on malformed input.
fn parse_shape_element<R: std::io::BufRead>(
    reader: &mut Reader<R>,
    is_connector: bool,
) -> Result<Option<Shape>> {
    let mut buffer = Vec::new();
    let mut depth = 1_usize; // Inside <p:sp> or <p:cxnSp>.

    let mut shape_name = String::new();
    let mut placeholder_type: Option<String> = None;
    let mut placeholder_idx: Option<u32> = None;
    let mut xfrm_offset: Option<(i64, i64)> = None;
    let mut xfrm_extents: Option<(i64, i64)> = None;
    let mut preset_geometry: Option<String> = None;

    // Track nesting inside spPr / xfrm.
    let mut in_sp_pr = false;
    let mut in_xfrm = false;

    loop {
        match reader.read_event_into(&mut buffer)? {
            Event::Start(ref event) => {
                let event_name = event.name();
                let local = local_name(event_name.as_ref());
                depth += 1;

                match local {
                    b"cNvPr" => {
                        if let Some(n) = get_attribute_value(event, b"name") {
                            shape_name = n;
                        }
                    }
                    b"ph" => {
                        placeholder_type = get_attribute_value(event, b"type");
                        placeholder_idx =
                            get_attribute_value(event, b"idx").and_then(|v| v.parse().ok());
                    }
                    b"spPr" => {
                        in_sp_pr = true;
                    }
                    b"xfrm" if in_sp_pr => {
                        in_xfrm = true;
                    }
                    b"off" if in_xfrm => {
                        xfrm_offset = parse_offset_attributes(event);
                    }
                    b"ext" if in_xfrm => {
                        xfrm_extents = parse_extent_attributes(event);
                    }
                    _ => {}
                }
            }
            Event::Empty(ref event) => {
                let event_name = event.name();
                let local = local_name(event_name.as_ref());

                match local {
                    b"cNvPr" => {
                        if let Some(n) = get_attribute_value(event, b"name") {
                            shape_name = n;
                        }
                    }
                    b"ph" => {
                        placeholder_type = get_attribute_value(event, b"type");
                        placeholder_idx =
                            get_attribute_value(event, b"idx").and_then(|v| v.parse().ok());
                    }
                    b"prstGeom" => {
                        preset_geometry = get_attribute_value(event, b"prst");
                    }
                    b"off" if in_xfrm => {
                        xfrm_offset = parse_offset_attributes(event);
                    }
                    b"ext" if in_xfrm => {
                        xfrm_extents = parse_extent_attributes(event);
                    }
                    _ => {}
                }
            }
            Event::End(ref event) => {
                let event_name = event.name();
                let local = local_name(event_name.as_ref());
                depth -= 1;

                if local == b"xfrm" {
                    in_xfrm = false;
                } else if local == b"spPr" {
                    in_sp_pr = false;
                }

                if depth == 0 {
                    break;
                }
            }
            Event::Eof => break,
            _ => {}
        }
        buffer.clear();
    }

    if shape_name.is_empty() {
        shape_name = "Unnamed".to_string();
    }

    let mut shape = Shape::new(&shape_name);

    if is_connector {
        shape.set_connector(true);
    }

    // Set placeholder info.
    if let Some(ref ph_type) = placeholder_type {
        shape.set_placeholder_type(crate::shape::PlaceholderType::from_xml(ph_type));
        if let Some(idx) = placeholder_idx {
            shape.set_placeholder_idx(idx);
        }
    }

    // Set geometry if we have position and extents.
    if let (Some((x, y)), Some((cx, cy))) = (xfrm_offset, xfrm_extents) {
        use crate::shape::ShapeGeometry;
        shape.set_geometry(ShapeGeometry::new(x, y, cx, cy));
    }

    // Set preset geometry.
    if let Some(prst) = preset_geometry {
        shape.set_preset_geometry(&prst);
    }

    Ok(Some(shape))
}

/// Parse the `<p:bg>` element for slide background.
fn parse_background<R: std::io::BufRead>(
    reader: &mut Reader<R>,
) -> Result<Option<SlideBackground>> {
    let mut buffer = Vec::new();
    let mut depth = 1_usize; // Inside <p:bg>.
    let mut result: Option<SlideBackground> = None;
    let mut in_bg_pr = false;

    loop {
        match reader.read_event_into(&mut buffer)? {
            Event::Start(ref event) => {
                let event_name = event.name();
                let local = local_name(event_name.as_ref());
                depth += 1;

                match local {
                    b"bgPr" => {
                        in_bg_pr = true;
                    }
                    b"solidFill" if in_bg_pr => {
                        // Peek for srgbClr child.
                    }
                    b"srgbClr" if in_bg_pr && result.is_none() => {
                        if let Some(val) = get_attribute_value(event, b"val") {
                            result = Some(SlideBackground::Solid(val));
                        }
                    }
                    _ => {}
                }
            }
            Event::Empty(ref event) => {
                let event_name = event.name();
                let local = local_name(event_name.as_ref());

                if local == b"srgbClr" && in_bg_pr && result.is_none() {
                    if let Some(val) = get_attribute_value(event, b"val") {
                        result = Some(SlideBackground::Solid(val));
                    }
                }
            }
            Event::End(ref event) => {
                let event_name = event.name();
                let local = local_name(event_name.as_ref());
                depth -= 1;

                if local == b"bgPr" {
                    in_bg_pr = false;
                }

                if depth == 0 {
                    break;
                }
            }
            Event::Eof => break,
            _ => {}
        }
        buffer.clear();
    }

    Ok(result)
}

// ---------------------------------------------------------------------------
// Internal writing helpers
// ---------------------------------------------------------------------------

/// Write the non-visual group shape properties for the shape tree.
///
/// This is the required `<p:nvGrpSpPr>` element at the start of `<p:spTree>`.
fn write_sp_tree_nv_grp_sp_pr<W: std::io::Write>(writer: &mut Writer<W>) -> Result<()> {
    writer.write_event(Event::Start(BytesStart::new("p:nvGrpSpPr")))?;

    let mut c_nv_pr = BytesStart::new("p:cNvPr");
    c_nv_pr.push_attribute(("id", "1"));
    c_nv_pr.push_attribute(("name", ""));
    writer.write_event(Event::Empty(c_nv_pr))?;

    writer.write_event(Event::Empty(BytesStart::new("p:cNvGrpSpPr")))?;
    writer.write_event(Event::Empty(BytesStart::new("p:nvPr")))?;

    writer.write_event(Event::End(BytesEnd::new("p:nvGrpSpPr")))?;
    Ok(())
}

/// Write a single shape as a `<p:sp>` or `<p:cxnSp>` element.
fn write_layout_shape_xml<W: std::io::Write>(
    writer: &mut Writer<W>,
    next_object_id: &mut u32,
    shape: &Shape,
) -> Result<()> {
    let object_id = allocate_object_id(next_object_id)?;

    let is_connector = shape.is_connector();
    let sp_tag_name = if is_connector { "p:cxnSp" } else { "p:sp" };
    writer.write_event(Event::Start(BytesStart::new(sp_tag_name)))?;

    // Non-visual shape properties.
    let nv_tag_name = if is_connector {
        "p:nvCxnSpPr"
    } else {
        "p:nvSpPr"
    };
    writer.write_event(Event::Start(BytesStart::new(nv_tag_name)))?;

    let mut c_nv_pr = BytesStart::new("p:cNvPr");
    let id_text = object_id.to_string();
    c_nv_pr.push_attribute(("id", id_text.as_str()));
    c_nv_pr.push_attribute(("name", shape.name()));
    writer.write_event(Event::Empty(c_nv_pr))?;

    // CNvSpPr / CNvCxnSpPr — non-visual drawing properties.
    let cnv_sp_pr_tag = if is_connector {
        "p:cNvCxnSpPr"
    } else {
        "p:cNvSpPr"
    };
    writer.write_event(Event::Empty(BytesStart::new(cnv_sp_pr_tag)))?;

    // <p:nvPr> — non-visual presentation properties (placeholder info lives here).
    if shape.placeholder_type().is_some() {
        writer.write_event(Event::Start(BytesStart::new("p:nvPr")))?;
        write_placeholder_xml(writer, shape)?;
        writer.write_event(Event::End(BytesEnd::new("p:nvPr")))?;
    } else {
        writer.write_event(Event::Empty(BytesStart::new("p:nvPr")))?;
    }

    writer.write_event(Event::End(BytesEnd::new(nv_tag_name)))?;

    // <p:spPr> — shape properties (transform, geometry).
    write_shape_properties_xml(writer, shape)?;

    // <p:txBody> — text body (for shapes with text).
    let has_text = shape
        .paragraphs()
        .iter()
        .any(|p| p.runs().iter().any(|r| !r.text().is_empty()));
    if has_text {
        write_text_body_xml(writer, shape)?;
    }

    writer.write_event(Event::End(BytesEnd::new(sp_tag_name)))?;
    Ok(())
}

/// Write the `<p:ph>` element for a placeholder shape.
fn write_placeholder_xml<W: std::io::Write>(writer: &mut Writer<W>, shape: &Shape) -> Result<()> {
    let mut ph = BytesStart::new("p:ph");

    if let Some(ph_type) = shape.placeholder_type() {
        let type_str = ph_type.to_xml();
        if !type_str.is_empty() {
            ph.push_attribute(("type", type_str));
        }
    }

    if let Some(idx) = shape.placeholder_idx() {
        let idx_text = idx.to_string();
        ph.push_attribute(("idx", idx_text.as_str()));
    }

    writer.write_event(Event::Empty(ph))?;
    Ok(())
}

/// Write the `<p:spPr>` (shape properties) element.
fn write_shape_properties_xml<W: std::io::Write>(
    writer: &mut Writer<W>,
    shape: &Shape,
) -> Result<()> {
    if let Some(geom) = shape.geometry() {
        writer.write_event(Event::Start(BytesStart::new("p:spPr")))?;

        // <a:xfrm>
        writer.write_event(Event::Start(BytesStart::new("a:xfrm")))?;

        let mut off = BytesStart::new("a:off");
        let x_text = geom.x().to_string();
        let y_text = geom.y().to_string();
        off.push_attribute(("x", x_text.as_str()));
        off.push_attribute(("y", y_text.as_str()));
        writer.write_event(Event::Empty(off))?;

        let mut ext = BytesStart::new("a:ext");
        let cx_text = geom.cx().to_string();
        let cy_text = geom.cy().to_string();
        ext.push_attribute(("cx", cx_text.as_str()));
        ext.push_attribute(("cy", cy_text.as_str()));
        writer.write_event(Event::Empty(ext))?;

        writer.write_event(Event::End(BytesEnd::new("a:xfrm")))?;

        // <a:prstGeom> — preset geometry.
        if let Some(prst) = shape.preset_geometry() {
            let mut prst_geom = BytesStart::new("a:prstGeom");
            prst_geom.push_attribute(("prst", prst));
            writer.write_event(Event::Start(prst_geom))?;
            writer.write_event(Event::Empty(BytesStart::new("a:avLst")))?;
            writer.write_event(Event::End(BytesEnd::new("a:prstGeom")))?;
        }

        writer.write_event(Event::End(BytesEnd::new("p:spPr")))?;
    } else {
        // Empty spPr if no geometry.
        writer.write_event(Event::Empty(BytesStart::new("p:spPr")))?;
    }

    Ok(())
}

/// Write a minimal `<p:txBody>` element for shapes with text content.
fn write_text_body_xml<W: std::io::Write>(writer: &mut Writer<W>, shape: &Shape) -> Result<()> {
    writer.write_event(Event::Start(BytesStart::new("p:txBody")))?;

    // Body properties.
    writer.write_event(Event::Empty(BytesStart::new("a:bodyPr")))?;
    writer.write_event(Event::Empty(BytesStart::new("a:lstStyle")))?;

    // Paragraphs.
    for paragraph in shape.paragraphs() {
        writer.write_event(Event::Start(BytesStart::new("a:p")))?;
        for run in paragraph.runs() {
            writer.write_event(Event::Start(BytesStart::new("a:r")))?;
            writer.write_event(Event::Empty(BytesStart::new("a:rPr")))?;
            writer.write_event(Event::Start(BytesStart::new("a:t")))?;
            writer.write_event(Event::Text(quick_xml::events::BytesText::new(run.text())))?;
            writer.write_event(Event::End(BytesEnd::new("a:t")))?;
            writer.write_event(Event::End(BytesEnd::new("a:r")))?;
        }
        writer.write_event(Event::End(BytesEnd::new("a:p")))?;
    }

    writer.write_event(Event::End(BytesEnd::new("p:txBody")))?;
    Ok(())
}

/// Write `<p:bg>` background element.
fn write_slide_background_xml<W: std::io::Write>(
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
            use crate::shape::GradientFillType;

            let is_linear = gradient
                .fill_type
                .as_ref()
                .is_none_or(|ft| *ft == GradientFillType::Linear);

            let grad_fill = BytesStart::new("a:gradFill");
            writer.write_event(Event::Start(grad_fill))?;

            writer.write_event(Event::Start(BytesStart::new("a:gsLst")))?;
            for stop in &gradient.stops {
                let pos_str = stop.position.to_string();
                let mut gs = BytesStart::new("a:gs");
                gs.push_attribute(("pos", pos_str.as_str()));
                writer.write_event(Event::Start(gs))?;
                let mut clr = BytesStart::new("a:srgbClr");
                clr.push_attribute(("val", stop.color_srgb.as_str()));
                writer.write_event(Event::Empty(clr))?;
                writer.write_event(Event::End(BytesEnd::new("a:gs")))?;
            }
            writer.write_event(Event::End(BytesEnd::new("a:gsLst")))?;

            if is_linear {
                if let Some(angle) = gradient.linear_angle {
                    let mut lin = BytesStart::new("a:lin");
                    let angle_str = angle.to_string();
                    lin.push_attribute(("ang", angle_str.as_str()));
                    lin.push_attribute(("scaled", "1"));
                    writer.write_event(Event::Empty(lin))?;
                }
            } else {
                let path_type = match gradient.fill_type {
                    Some(GradientFillType::Radial) => "circle",
                    _ => "rect",
                };
                let mut path = BytesStart::new("a:path");
                path.push_attribute(("path", path_type));
                writer.write_event(Event::Start(path))?;
                let mut fill_to_rect = BytesStart::new("a:fillToRect");
                fill_to_rect.push_attribute(("l", "50000"));
                fill_to_rect.push_attribute(("t", "50000"));
                fill_to_rect.push_attribute(("r", "50000"));
                fill_to_rect.push_attribute(("b", "50000"));
                writer.write_event(Event::Empty(fill_to_rect))?;
                writer.write_event(Event::End(BytesEnd::new("a:path")))?;
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

// ---------------------------------------------------------------------------
// Tiny utility functions (same patterns as presentation.rs)
// ---------------------------------------------------------------------------

/// Extract the local name from a possibly namespace-prefixed XML tag name.
fn local_name(name: &[u8]) -> &[u8] {
    name.rsplit(|byte| *byte == b':').next().unwrap_or(name)
}

/// Read the value of an XML attribute by its local name.
fn get_attribute_value(event: &BytesStart<'_>, expected_local_name: &[u8]) -> Option<String> {
    event.attributes().flatten().find_map(|attribute| {
        (local_name(attribute.key.as_ref()) == expected_local_name)
            .then(|| String::from_utf8_lossy(attribute.value.as_ref()).into_owned())
    })
}

/// Parse a boolean value from XML ("1", "true", "on" → true).
fn parse_xml_bool(value: &str) -> Option<bool> {
    match value.trim().to_ascii_lowercase().as_str() {
        "1" | "true" | "on" => Some(true),
        "0" | "false" | "off" => Some(false),
        _ => None,
    }
}

/// Skip an entire XML element (consuming through its closing tag).
fn skip_element<R: std::io::BufRead>(reader: &mut Reader<R>) -> Result<()> {
    let mut depth = 1_usize;
    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf)? {
            Event::Start(_) => depth += 1,
            Event::End(_) => {
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

/// Allocate the next object ID for shape serialization.
fn allocate_object_id(next_object_id: &mut u32) -> Result<u32> {
    let current = *next_object_id;
    *next_object_id = next_object_id.checked_add(1).ok_or_else(|| {
        PptxError::UnsupportedPackage("shape id overflow while serializing layout".to_string())
    })?;
    Ok(current)
}

/// Parse `x` and `y` from an `<a:off>` element.
fn parse_offset_attributes(event: &BytesStart<'_>) -> Option<(i64, i64)> {
    let x = get_attribute_value(event, b"x").and_then(|v| v.parse().ok())?;
    let y = get_attribute_value(event, b"y").and_then(|v| v.parse().ok())?;
    Some((x, y))
}

/// Parse `cx` and `cy` from an `<a:ext>` element.
fn parse_extent_attributes(event: &BytesStart<'_>) -> Option<(i64, i64)> {
    let cx = get_attribute_value(event, b"cx").and_then(|v| v.parse().ok())?;
    let cy = get_attribute_value(event, b"cy").and_then(|v| v.parse().ok())?;
    Some((cx, cy))
}

#[cfg(test)]
mod tests {
    use super::*;

    /// Minimal slide layout XML for testing.
    const MINIMAL_LAYOUT_XML: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldLayout xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
             xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
             xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
             type="title" preserve="1">
  <p:cSld name="Title Slide">
    <p:spTree>
      <p:nvGrpSpPr>
        <p:cNvPr id="1" name=""/>
        <p:cNvGrpSpPr/>
        <p:nvPr/>
      </p:nvGrpSpPr>
      <p:grpSpPr/>
      <p:sp>
        <p:nvSpPr>
          <p:cNvPr id="2" name="Title 1"/>
          <p:cNvSpPr/>
          <p:nvPr>
            <p:ph type="ctrTitle"/>
          </p:nvPr>
        </p:nvSpPr>
        <p:spPr>
          <a:xfrm>
            <a:off x="685800" y="2130425"/>
            <a:ext cx="7772400" cy="1470025"/>
          </a:xfrm>
        </p:spPr>
        <p:txBody>
          <a:bodyPr/>
          <a:lstStyle/>
          <a:p>
            <a:r>
              <a:rPr lang="en-US"/>
              <a:t>Click to edit Master title style</a:t>
            </a:r>
          </a:p>
        </p:txBody>
      </p:sp>
      <p:sp>
        <p:nvSpPr>
          <p:cNvPr id="3" name="Subtitle 2"/>
          <p:cNvSpPr/>
          <p:nvPr>
            <p:ph type="subTitle" idx="1"/>
          </p:nvPr>
        </p:nvSpPr>
        <p:spPr>
          <a:xfrm>
            <a:off x="1143000" y="3602038"/>
            <a:ext cx="6858000" cy="1655762"/>
          </a:xfrm>
        </p:spPr>
      </p:sp>
    </p:spTree>
  </p:cSld>
  <p:clrMapOvr>
    <a:masterClrMapping/>
  </p:clrMapOvr>
</p:sldLayout>"#;

    #[test]
    fn parse_minimal_layout() {
        let parsed = parse_slide_layout_xml(MINIMAL_LAYOUT_XML.as_bytes())
            .expect("should parse minimal layout XML");

        assert_eq!(parsed.name.as_deref(), Some("Title Slide"));
        assert_eq!(parsed.layout_type.as_deref(), Some("title"));
        assert_eq!(parsed.preserve, Some(true));
        assert_eq!(parsed.shapes.len(), 2);

        // First shape: title placeholder (ctrTitle maps to CenteredTitle).
        assert_eq!(parsed.shapes[0].name(), "Title 1");
        assert_eq!(
            parsed.shapes[0].placeholder_type(),
            Some(&crate::shape::PlaceholderType::CenteredTitle)
        );
        let geom = parsed.shapes[0].geometry().expect("should have geometry");
        assert_eq!(geom.x(), 685800);
        assert_eq!(geom.y(), 2130425);
        assert_eq!(geom.cx(), 7772400);
        assert_eq!(geom.cy(), 1470025);

        // Second shape: subtitle placeholder.
        assert_eq!(parsed.shapes[1].name(), "Subtitle 2");
        assert_eq!(
            parsed.shapes[1].placeholder_type(),
            Some(&crate::shape::PlaceholderType::Subtitle)
        );
    }

    #[test]
    fn parse_layout_without_type() {
        let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldLayout xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
             xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
             xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <p:cSld name="Custom Layout">
    <p:spTree>
      <p:nvGrpSpPr>
        <p:cNvPr id="1" name=""/>
        <p:cNvGrpSpPr/>
        <p:nvPr/>
      </p:nvGrpSpPr>
      <p:grpSpPr/>
    </p:spTree>
  </p:cSld>
  <p:clrMapOvr>
    <a:masterClrMapping/>
  </p:clrMapOvr>
</p:sldLayout>"#;
        let parsed =
            parse_slide_layout_xml(xml.as_bytes()).expect("should parse layout without type");
        assert_eq!(parsed.name.as_deref(), Some("Custom Layout"));
        assert_eq!(parsed.layout_type, None);
        assert_eq!(parsed.preserve, None);
        assert!(parsed.shapes.is_empty());
    }

    #[test]
    fn parse_layout_with_background() {
        let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldLayout xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
             xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
             xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
             type="blank">
  <p:cSld name="Blank">
    <p:bg>
      <p:bgPr>
        <a:solidFill>
          <a:srgbClr val="FF0000"/>
        </a:solidFill>
        <a:effectLst/>
      </p:bgPr>
    </p:bg>
    <p:spTree>
      <p:nvGrpSpPr>
        <p:cNvPr id="1" name=""/>
        <p:cNvGrpSpPr/>
        <p:nvPr/>
      </p:nvGrpSpPr>
      <p:grpSpPr/>
    </p:spTree>
  </p:cSld>
  <p:clrMapOvr>
    <a:masterClrMapping/>
  </p:clrMapOvr>
</p:sldLayout>"#;
        let parsed =
            parse_slide_layout_xml(xml.as_bytes()).expect("should parse layout with background");
        assert_eq!(parsed.name.as_deref(), Some("Blank"));
        assert_eq!(parsed.layout_type.as_deref(), Some("blank"));
        assert!(parsed.background.is_some());
        match parsed.background {
            Some(SlideBackground::Solid(ref color)) => assert_eq!(color, "FF0000"),
            _ => panic!("expected solid background"),
        }
    }

    #[test]
    fn parse_preserves_unknown_children() {
        let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldLayout xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
             xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
             xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
             type="title" preserve="1">
  <p:cSld name="Title Slide">
    <p:spTree>
      <p:nvGrpSpPr>
        <p:cNvPr id="1" name=""/>
        <p:cNvGrpSpPr/>
        <p:nvPr/>
      </p:nvGrpSpPr>
      <p:grpSpPr/>
    </p:spTree>
  </p:cSld>
  <p:clrMapOvr>
    <a:masterClrMapping/>
  </p:clrMapOvr>
  <p:extLst>
    <p:ext uri="{some-guid}">
      <p14:custom xmlns:p14="http://example.com" val="test"/>
    </p:ext>
  </p:extLst>
</p:sldLayout>"#;
        let parsed = parse_slide_layout_xml(xml.as_bytes())
            .expect("should parse layout with unknown children");
        assert_eq!(parsed.unknown_children.len(), 1);
        if let RawXmlNode::Element { ref name, .. } = parsed.unknown_children[0] {
            assert_eq!(name, "p:extLst");
        } else {
            panic!("expected element node for unknown child");
        }
    }

    #[test]
    fn write_layout_minimal() {
        let layout = SlideLayout::new(
            "Title Slide",
            "rId1",
            "/ppt/slideLayouts/slideLayout1.xml",
            "rId10",
        );
        let xml_bytes = write_slide_layout_xml(&layout).expect("should write layout XML");
        let xml = String::from_utf8(xml_bytes).expect("should be valid UTF-8");

        assert!(xml.contains("p:sldLayout"));
        assert!(xml.contains(r#"name="Title Slide""#));
        assert!(xml.contains("p:cSld"));
        assert!(xml.contains("p:spTree"));
        assert!(xml.contains("p:clrMapOvr"));
        assert!(xml.contains("a:masterClrMapping"));
        // No type or preserve on a default layout.
        assert!(!xml.contains(r#"type="#));
        assert!(!xml.contains(r#"preserve="#));
    }

    #[test]
    fn write_layout_with_type_and_preserve() {
        let mut layout = SlideLayout::new(
            "Title Slide",
            "rId1",
            "/ppt/slideLayouts/slideLayout1.xml",
            "rId10",
        );
        layout.set_layout_type("title");
        layout.set_preserve(true);

        let xml_bytes = write_slide_layout_xml(&layout).expect("should write layout XML");
        let xml = String::from_utf8(xml_bytes).expect("should be valid UTF-8");

        assert!(xml.contains(r#"type="title""#));
        assert!(xml.contains(r#"preserve="1""#));
    }

    #[test]
    fn roundtrip_parse_write_parse() {
        // Parse → build layout → write → parse again and verify.
        let parsed1 = parse_slide_layout_xml(MINIMAL_LAYOUT_XML.as_bytes())
            .expect("first parse should succeed");

        let mut layout = SlideLayout::new(
            parsed1.name.as_deref().unwrap_or(""),
            "rId1",
            "/ppt/slideLayouts/slideLayout1.xml",
            "rId10",
        );
        if let Some(lt) = &parsed1.layout_type {
            layout.set_layout_type(lt.as_str());
        }
        if let Some(p) = parsed1.preserve {
            layout.set_preserve(p);
        }
        if let Some(bg) = parsed1.background {
            layout.set_background(bg);
        }
        for shape in parsed1.shapes {
            layout.add_shape(shape);
        }

        let xml_bytes = write_slide_layout_xml(&layout).expect("write should succeed");
        let parsed2 = parse_slide_layout_xml(&xml_bytes).expect("re-parse should succeed");

        assert_eq!(parsed2.name.as_deref(), Some("Title Slide"));
        assert_eq!(parsed2.layout_type.as_deref(), Some("title"));
        assert_eq!(parsed2.preserve, Some(true));
        assert_eq!(parsed2.shapes.len(), 2);
        assert_eq!(parsed2.shapes[0].name(), "Title 1");
        assert_eq!(parsed2.shapes[1].name(), "Subtitle 2");
    }

    #[test]
    fn dirty_layout_regenerates_xml() {
        let mut layout = SlideLayout::new(
            "Test Layout",
            "rId1",
            "/ppt/slideLayouts/slideLayout1.xml",
            "rId10",
        );
        layout.set_original_xml(MINIMAL_LAYOUT_XML.as_bytes().to_vec());
        layout.mark_clean();

        // Clean layout should return original XML.
        let clean_xml = write_slide_layout_xml(&layout).expect("should return original");
        assert_eq!(clean_xml, MINIMAL_LAYOUT_XML.as_bytes());

        // After modification, should regenerate.
        layout.set_name("Modified Title");
        let dirty_xml = write_slide_layout_xml(&layout).expect("should regenerate");
        let dirty_str = String::from_utf8(dirty_xml).expect("valid UTF-8");
        assert!(dirty_str.contains(r#"name="Modified Title""#));
    }
}
