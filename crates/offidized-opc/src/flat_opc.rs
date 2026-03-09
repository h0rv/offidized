//! FlatOPC (Flat OPC XML Document) support.
//!
//! FlatOPC is an XML representation of an entire OPC package in a single file.
//! Instead of a ZIP archive, all parts are embedded as XML elements with their
//! data either inline (for XML parts) or base64-encoded (for binary parts).
//!
//! The format uses the namespace `http://schemas.microsoft.com/office/2006/xmlPackage`.
//!
//! ## Example FlatOPC structure:
//! ```xml
//! <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
//! <pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage">
//!   <pkg:part pkg:name="/word/document.xml"
//!             pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml">
//!     <pkg:xmlData>
//!       <w:document xmlns:w="...">...</w:document>
//!     </pkg:xmlData>
//!   </pkg:part>
//!   <pkg:part pkg:name="/word/media/image1.png" pkg:contentType="image/png">
//!     <pkg:binaryData>iVBORw0KGgo...</pkg:binaryData>
//!   </pkg:part>
//! </pkg:package>
//! ```

use crate::content_types::ContentTypes;
use crate::error::{OpcError, Result};
use crate::package::Package;
use crate::part::Part;
use crate::relationship::Relationships;
use crate::uri::PartUri;

use quick_xml::events::{BytesDecl, BytesEnd, BytesRef, BytesStart, BytesText, Event};
use quick_xml::{Reader, Writer};
use std::io::{BufRead, Write};

const PKG_NS: &str = "http://schemas.microsoft.com/office/2006/xmlPackage";

/// Serialize a Package to FlatOPC XML format.
pub fn to_flat_opc<W: Write>(package: &Package, writer: W) -> Result<()> {
    let mut xml = Writer::new_with_indent(writer, b' ', 2);

    xml.write_event(Event::Decl(BytesDecl::new(
        "1.0",
        Some("UTF-8"),
        Some("yes"),
    )))?;

    let mut root = BytesStart::new("pkg:package");
    root.push_attribute(("xmlns:pkg", PKG_NS));
    xml.write_event(Event::Start(root))?;

    // Write [Content_Types].xml as a part
    {
        let mut ct_bytes = Vec::new();
        package.content_types().to_xml(&mut ct_bytes)?;

        let mut part_elem = BytesStart::new("pkg:part");
        part_elem.push_attribute(("pkg:name", "/[Content_Types].xml"));
        part_elem.push_attribute((
            "pkg:contentType",
            "application/vnd.openxmlformats-package.content-types+xml",
        ));
        xml.write_event(Event::Start(part_elem))?;

        xml.write_event(Event::Start(BytesStart::new("pkg:xmlData")))?;
        // Write the content types XML inline (raw)
        xml.get_mut().write_all(&ct_bytes)?;
        xml.write_event(Event::End(BytesEnd::new("pkg:xmlData")))?;

        xml.write_event(Event::End(BytesEnd::new("pkg:part")))?;
    }

    // Write package-level relationships
    if package.relationships().should_write_xml() {
        let mut rels_bytes = Vec::new();
        package.relationships().to_xml(&mut rels_bytes)?;

        let mut part_elem = BytesStart::new("pkg:part");
        part_elem.push_attribute(("pkg:name", "/_rels/.rels"));
        part_elem.push_attribute((
            "pkg:contentType",
            "application/vnd.openxmlformats-package.relationships+xml",
        ));
        xml.write_event(Event::Start(part_elem))?;

        xml.write_event(Event::Start(BytesStart::new("pkg:xmlData")))?;
        xml.get_mut().write_all(&rels_bytes)?;
        xml.write_event(Event::End(BytesEnd::new("pkg:xmlData")))?;

        xml.write_event(Event::End(BytesEnd::new("pkg:part")))?;
    }

    // Write all parts (sorted for deterministic output)
    let mut part_uris: Vec<&str> = package.part_uris();
    part_uris.sort();

    for uri in part_uris {
        let part = match package.get_part(uri) {
            Some(p) => p,
            None => continue,
        };

        let mut part_elem = BytesStart::new("pkg:part");
        part_elem.push_attribute(("pkg:name", uri));
        if let Some(ct) = &part.content_type {
            part_elem.push_attribute(("pkg:contentType", ct.as_str()));
        }
        xml.write_event(Event::Start(part_elem))?;

        if part.is_xml() {
            xml.write_event(Event::Start(BytesStart::new("pkg:xmlData")))?;
            xml.get_mut().write_all(part.data.as_bytes())?;
            xml.write_event(Event::End(BytesEnd::new("pkg:xmlData")))?;
        } else {
            xml.write_event(Event::Start(BytesStart::new("pkg:binaryData")))?;
            let mut base64_writer = Base64Writer::new(xml.get_mut());
            base64_writer.write_all(part.data.as_bytes())?;
            base64_writer.finish()?;
            xml.write_event(Event::End(BytesEnd::new("pkg:binaryData")))?;
        }

        // Write part relationships
        if part.relationships.should_write_xml() {
            let rels_uri = match part.uri.relationship_uri() {
                Ok(u) => u.as_str().to_string(),
                Err(_) => continue,
            };

            let mut rels_bytes = Vec::new();
            part.relationships.to_xml(&mut rels_bytes)?;

            let mut rels_elem = BytesStart::new("pkg:part");
            rels_elem.push_attribute(("pkg:name", rels_uri.as_str()));
            rels_elem.push_attribute((
                "pkg:contentType",
                "application/vnd.openxmlformats-package.relationships+xml",
            ));
            xml.write_event(Event::Start(rels_elem))?;

            xml.write_event(Event::Start(BytesStart::new("pkg:xmlData")))?;
            xml.get_mut().write_all(&rels_bytes)?;
            xml.write_event(Event::End(BytesEnd::new("pkg:xmlData")))?;

            xml.write_event(Event::End(BytesEnd::new("pkg:part")))?;
        }

        xml.write_event(Event::End(BytesEnd::new("pkg:part")))?;
    }

    xml.write_event(Event::End(BytesEnd::new("pkg:package")))?;
    Ok(())
}

/// Deserialize a FlatOPC XML document into a Package.
pub fn from_flat_opc<R: BufRead>(reader: R) -> Result<Package> {
    let mut xml = Reader::from_reader(reader);
    xml.config_mut().trim_text(true);

    let mut package = Package::new();
    let mut buf = Vec::new();
    let mut content_types_xml: Option<Vec<u8>> = None;
    let mut package_rels_xml: Option<Vec<u8>> = None;

    // State for current part being parsed
    let mut current_part_name: Option<String> = None;
    let mut current_content_type: Option<String> = None;
    let mut in_binary_data = false;
    let mut xml_data_buf = Vec::new();
    let mut binary_text_buf = String::new();

    // Track relationship parts to associate later
    let mut rels_parts: Vec<(String, Vec<u8>)> = Vec::new();

    loop {
        match xml.read_event_into(&mut buf) {
            Ok(Event::Start(ref e)) => {
                let name = e.name();
                let local = local_name(name.as_ref());
                match local {
                    b"part" => {
                        current_part_name = None;
                        current_content_type = None;
                        for attr in e.attributes().flatten() {
                            let key = local_name(attr.key.as_ref());
                            match key {
                                b"name" => {
                                    current_part_name = Some(
                                        attr.decode_and_unescape_value(xml.decoder())?.into_owned(),
                                    );
                                }
                                b"contentType" => {
                                    current_content_type = Some(
                                        attr.decode_and_unescape_value(xml.decoder())?.into_owned(),
                                    );
                                }
                                _ => {}
                            }
                        }
                    }
                    b"xmlData" => {
                        xml_data_buf.clear();
                        // Read inner XML events and re-serialize them to capture
                        // the XML content inside <pkg:xmlData>...</pkg:xmlData>.
                        let mut depth: u32 = 0;
                        let mut inner_buf = Vec::new();
                        loop {
                            match xml.read_event_into(&mut inner_buf) {
                                Ok(Event::Start(ref inner_e)) => {
                                    let inner_name = inner_e.name();
                                    let inner_local = local_name(inner_name.as_ref());
                                    if depth == 0 && inner_local == b"xmlData" {
                                        // Shouldn't happen, but guard
                                    }
                                    depth += 1;
                                    // Re-serialize the start tag
                                    xml_data_buf.push(b'<');
                                    xml_data_buf.extend_from_slice(inner_e.name().as_ref());
                                    for attr in inner_e.attributes().flatten() {
                                        xml_data_buf.push(b' ');
                                        xml_data_buf.extend_from_slice(attr.key.as_ref());
                                        xml_data_buf.extend_from_slice(b"=\"");
                                        xml_data_buf.extend_from_slice(&attr.value);
                                        xml_data_buf.push(b'"');
                                    }
                                    xml_data_buf.push(b'>');
                                }
                                Ok(Event::End(ref inner_e)) => {
                                    let inner_name = inner_e.name();
                                    let inner_local = local_name(inner_name.as_ref());
                                    if inner_local == b"xmlData" && depth == 0 {
                                        break; // End of xmlData
                                    }
                                    depth = depth.saturating_sub(1);
                                    xml_data_buf.extend_from_slice(b"</");
                                    xml_data_buf.extend_from_slice(inner_e.name().as_ref());
                                    xml_data_buf.push(b'>');
                                }
                                Ok(Event::Empty(ref inner_e)) => {
                                    xml_data_buf.push(b'<');
                                    xml_data_buf.extend_from_slice(inner_e.name().as_ref());
                                    for attr in inner_e.attributes().flatten() {
                                        xml_data_buf.push(b' ');
                                        xml_data_buf.extend_from_slice(attr.key.as_ref());
                                        xml_data_buf.extend_from_slice(b"=\"");
                                        xml_data_buf.extend_from_slice(&attr.value);
                                        xml_data_buf.push(b'"');
                                    }
                                    xml_data_buf.extend_from_slice(b"/>");
                                }
                                Ok(Event::Text(ref inner_e)) => {
                                    xml_data_buf
                                        .extend_from_slice(decode_text_event(inner_e)?.as_bytes());
                                }
                                Ok(Event::GeneralRef(ref inner_e)) => {
                                    append_general_ref(&mut xml_data_buf, inner_e)?;
                                }
                                Ok(Event::CData(ref inner_e)) => {
                                    xml_data_buf.extend_from_slice(b"<![CDATA[");
                                    xml_data_buf.extend_from_slice(inner_e.as_ref());
                                    xml_data_buf.extend_from_slice(b"]]>");
                                }
                                Ok(Event::Decl(ref inner_e)) => {
                                    xml_data_buf.extend_from_slice(b"<?xml");
                                    xml_data_buf.extend_from_slice(inner_e.as_ref());
                                    xml_data_buf.extend_from_slice(b"?>");
                                }
                                Ok(Event::Eof) => break,
                                Err(err) => return Err(err.into()),
                                _ => {}
                            }
                            inner_buf.clear();
                        }
                        // Process the collected XML data for the current part
                        if let Some(ref name) = current_part_name {
                            process_flat_opc_part(
                                name,
                                &current_content_type,
                                &xml_data_buf,
                                true,
                                &mut package,
                                &mut content_types_xml,
                                &mut package_rels_xml,
                                &mut rels_parts,
                            )?;
                        }
                    }
                    b"binaryData" => {
                        in_binary_data = true;
                        binary_text_buf.clear();
                    }
                    _ => {}
                }
            }
            Ok(Event::Text(ref e)) if in_binary_data => {
                binary_text_buf.push_str(&decode_text_event(e)?);
            }
            Ok(Event::End(ref e)) => {
                let end_name = e.name();
                let local = local_name(end_name.as_ref());
                if local == b"binaryData" && in_binary_data {
                    in_binary_data = false;
                    // Decode base64
                    let clean: String = binary_text_buf
                        .chars()
                        .filter(|c| !c.is_whitespace())
                        .collect();
                    let data = base64_decode(&clean)?;

                    if let Some(ref name) = current_part_name {
                        process_flat_opc_part(
                            name,
                            &current_content_type,
                            &data,
                            false,
                            &mut package,
                            &mut content_types_xml,
                            &mut package_rels_xml,
                            &mut rels_parts,
                        )?;
                    }
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => return Err(e.into()),
            _ => {}
        }
        buf.clear();
    }

    // Apply content types if found
    if let Some(ct_xml) = content_types_xml {
        let ct = ContentTypes::from_xml_bytes(ct_xml)?;
        *package.content_types_mut() = ct;
    }

    // Apply package relationships if found
    if let Some(rels_xml) = package_rels_xml {
        let rels = Relationships::from_xml_bytes(rels_xml)?;
        *package.relationships_mut() = rels;
    }

    // Apply part relationships
    for (rels_path, rels_data) in rels_parts {
        // Convert relationship path back to source part URI
        // e.g., "/word/_rels/document.xml.rels" -> "/word/document.xml"
        if let Some(source_uri) = rels_path_to_source_uri(&rels_path) {
            if let Some(part) = package.get_part_mut(&source_uri) {
                part.relationships = Relationships::from_xml_bytes(rels_data)?;
            }
        }
    }

    Ok(package)
}

fn process_flat_opc_part(
    name: &str,
    content_type: &Option<String>,
    data: &[u8],
    is_xml: bool,
    package: &mut Package,
    content_types_xml: &mut Option<Vec<u8>>,
    package_rels_xml: &mut Option<Vec<u8>>,
    rels_parts: &mut Vec<(String, Vec<u8>)>,
) -> Result<()> {
    // Special handling for [Content_Types].xml
    if name == "/[Content_Types].xml" {
        *content_types_xml = Some(data.to_vec());
        return Ok(());
    }

    // Special handling for package rels
    if name == "/_rels/.rels" {
        *package_rels_xml = Some(data.to_vec());
        return Ok(());
    }

    // Part-level relationship files
    if name.contains("/_rels/") && name.ends_with(".rels") {
        rels_parts.push((name.to_string(), data.to_vec()));
        return Ok(());
    }

    // Regular part
    let uri = PartUri::new(name)?;
    let mut part = if is_xml {
        Part::new_xml(uri, data.to_vec())
    } else {
        Part::new(uri, data.to_vec())
    };
    part.content_type = content_type.clone();
    package.set_part(part);

    Ok(())
}

/// Convert a relationship part path to its source part URI.
/// "/word/_rels/document.xml.rels" -> "/word/document.xml"
fn rels_path_to_source_uri(rels_path: &str) -> Option<String> {
    let path = rels_path.strip_suffix(".rels")?;
    // Find _rels/ and extract directory + filename
    let rels_idx = path.rfind("/_rels/")?;
    let dir = &path[..rels_idx];
    let filename = &path[rels_idx + 7..]; // skip "/_rels/"
    Some(format!("{}/{}", dir, filename))
}

fn local_name(name: &[u8]) -> &[u8] {
    name.rsplit(|byte| *byte == b':').next().unwrap_or(name)
}

fn decode_text_event(event: &BytesText<'_>) -> Result<String> {
    event
        .xml_content()
        .map(|text| text.into_owned())
        .map_err(quick_xml::Error::from)
        .map_err(OpcError::from)
}

fn append_general_ref(buf: &mut Vec<u8>, event: &BytesRef<'_>) -> Result<()> {
    let reference = event
        .decode()
        .map_err(quick_xml::Error::from)
        .map_err(OpcError::from)?;
    buf.push(b'&');
    buf.extend_from_slice(reference.as_bytes());
    buf.push(b';');
    Ok(())
}

// Simple base64 encoder/decoder (no external dependency)

const BASE64_CHARS: &[u8; 64] = b"ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/";

fn base64_encode(data: &[u8]) -> String {
    let mut result = String::with_capacity(data.len().div_ceil(3) * 4);
    for chunk in data.chunks(3) {
        let b0 = chunk[0] as u32;
        let b1 = if chunk.len() > 1 { chunk[1] as u32 } else { 0 };
        let b2 = if chunk.len() > 2 { chunk[2] as u32 } else { 0 };
        let triple = (b0 << 16) | (b1 << 8) | b2;

        result.push(BASE64_CHARS[((triple >> 18) & 0x3F) as usize] as char);
        result.push(BASE64_CHARS[((triple >> 12) & 0x3F) as usize] as char);
        if chunk.len() > 1 {
            result.push(BASE64_CHARS[((triple >> 6) & 0x3F) as usize] as char);
        } else {
            result.push('=');
        }
        if chunk.len() > 2 {
            result.push(BASE64_CHARS[(triple & 0x3F) as usize] as char);
        } else {
            result.push('=');
        }
    }
    result
}

fn base64_decode(input: &str) -> Result<Vec<u8>> {
    fn decode_char(c: u8) -> std::result::Result<u8, OpcError> {
        match c {
            b'A'..=b'Z' => Ok(c - b'A'),
            b'a'..=b'z' => Ok(c - b'a' + 26),
            b'0'..=b'9' => Ok(c - b'0' + 52),
            b'+' => Ok(62),
            b'/' => Ok(63),
            b'=' => Ok(0),
            _ => Err(OpcError::MalformedPackage(format!(
                "invalid base64 character: {}",
                c as char
            ))),
        }
    }

    let bytes = input.as_bytes();
    let mut result = Vec::with_capacity(bytes.len() * 3 / 4);

    for chunk in bytes.chunks(4) {
        if chunk.len() < 4 {
            break;
        }
        let b0 = decode_char(chunk[0])?;
        let b1 = decode_char(chunk[1])?;
        let b2 = decode_char(chunk[2])?;
        let b3 = decode_char(chunk[3])?;

        let triple = ((b0 as u32) << 18) | ((b1 as u32) << 12) | ((b2 as u32) << 6) | (b3 as u32);

        result.push(((triple >> 16) & 0xFF) as u8);
        if chunk[2] != b'=' {
            result.push(((triple >> 8) & 0xFF) as u8);
        }
        if chunk[3] != b'=' {
            result.push((triple & 0xFF) as u8);
        }
    }

    Ok(result)
}

/// A simple base64-encoding writer wrapper.
struct Base64Writer<'a, W: Write> {
    inner: &'a mut W,
    buffer: Vec<u8>,
}

impl<'a, W: Write> Base64Writer<'a, W> {
    fn new(inner: &'a mut W) -> Self {
        Self {
            inner,
            buffer: Vec::new(),
        }
    }

    fn finish(self) -> std::io::Result<()> {
        if !self.buffer.is_empty() {
            let encoded = base64_encode(&self.buffer);
            self.inner.write_all(encoded.as_bytes())?;
        }
        Ok(())
    }
}

impl<'a, W: Write> Write for Base64Writer<'a, W> {
    fn write(&mut self, buf: &[u8]) -> std::io::Result<usize> {
        self.buffer.extend_from_slice(buf);
        Ok(buf.len())
    }

    fn flush(&mut self) -> std::io::Result<()> {
        Ok(())
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::relationship::TargetMode;

    #[test]
    fn base64_roundtrip() {
        let data = b"Hello, World!";
        let encoded = base64_encode(data);
        let decoded = base64_decode(&encoded).unwrap();
        assert_eq!(decoded, data);
    }

    #[test]
    fn base64_empty() {
        assert_eq!(base64_encode(b""), "");
        assert_eq!(base64_decode("").unwrap(), Vec::<u8>::new());
    }

    #[test]
    fn base64_padding() {
        assert_eq!(base64_encode(b"a"), "YQ==");
        assert_eq!(base64_encode(b"ab"), "YWI=");
        assert_eq!(base64_encode(b"abc"), "YWJj");
    }

    #[test]
    fn rels_path_conversion() {
        assert_eq!(
            rels_path_to_source_uri("/word/_rels/document.xml.rels"),
            Some("/word/document.xml".to_string())
        );
        assert_eq!(
            rels_path_to_source_uri("/xl/_rels/workbook.xml.rels"),
            Some("/xl/workbook.xml".to_string())
        );
        assert_eq!(rels_path_to_source_uri("/invalid"), None);
    }

    #[test]
    fn flat_opc_roundtrip_with_xml_and_binary_parts() {
        let mut package = Package::new();

        // Add an XML part
        let mut xml_part = Part::new_xml(
            PartUri::new("/doc/main.xml").unwrap(),
            b"<document>Hello</document>".to_vec(),
        );
        xml_part.content_type = Some("application/xml".to_string());
        package.set_part(xml_part);

        // Add a binary part
        let mut bin_part = Part::new(
            PartUri::new("/media/image.png").unwrap(),
            vec![0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A],
        );
        bin_part.content_type = Some("image/png".to_string());
        package.set_part(bin_part);

        // Add a relationship
        package.relationships_mut().add_new(
            "http://example.com/doc".to_string(),
            "/doc/main.xml".to_string(),
            TargetMode::Internal,
        );

        // Serialize to FlatOPC
        let mut flat_xml = Vec::new();
        to_flat_opc(&package, &mut flat_xml).unwrap();

        let flat_str = String::from_utf8_lossy(&flat_xml);
        assert!(flat_str.contains("pkg:package"));
        assert!(flat_str.contains("pkg:xmlData"));
        assert!(flat_str.contains("pkg:binaryData"));
        assert!(flat_str.contains("/doc/main.xml"));
        assert!(flat_str.contains("/media/image.png"));
    }

    #[test]
    fn package_to_flat_opc_bytes() {
        let mut package = Package::new();
        let mut part = Part::new_xml(PartUri::new("/test.xml").unwrap(), b"<test/>".to_vec());
        part.content_type = Some("application/xml".to_string());
        package.set_part(part);

        let mut output = Vec::new();
        to_flat_opc(&package, &mut output).unwrap();
        assert!(!output.is_empty());

        let s = String::from_utf8_lossy(&output);
        assert!(s.contains("<test/>"));
    }
}
