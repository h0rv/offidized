//! Raw XML node preservation for roundtrip fidelity.
//!
//! This is the secret sauce. When we parse an XML element and encounter children
//! or attributes we don't have typed fields for, we capture them as `RawXmlNode`
//! and write them back verbatim on save. This means:
//!
//! - A file created in Excel 365 with bleeding-edge features can be opened,
//!   have a cell value changed, and saved — with all unknown features intact.
//! - We never need to implement 100% of the spec to be safe for production use.

use quick_xml::events::{BytesEnd, BytesStart, BytesText, Event};
use quick_xml::{Reader, Writer};
use std::io::{BufRead, Write};

/// A raw XML node preserved for roundtrip fidelity.
///
/// This captures any XML content (elements, text, CDATA) that our typed
/// structs don't have explicit fields for. It's stored as raw bytes and
/// written back unchanged.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum RawXmlNode {
    /// An element with its full content (opening tag through closing tag).
    Element {
        /// The full tag name (with namespace prefix if present).
        name: String,
        /// Attributes as raw key-value pairs.
        attributes: Vec<(String, String)>,
        /// Child content (nested elements, text nodes).
        children: Vec<RawXmlNode>,
    },

    /// A text node.
    Text(String),

    /// A CDATA section.
    CData(String),

    /// Raw bytes we couldn't parse — absolute last resort for preservation.
    RawBytes(Vec<u8>),
}

impl RawXmlNode {
    /// Read a complete element (and all its children) from the XML reader.
    ///
    /// Call this when you encounter a `Start` event for an element you don't
    /// recognize. It will consume everything through the matching `End` event.
    pub fn read_element<R: BufRead>(
        reader: &mut Reader<R>,
        start: &BytesStart<'_>,
    ) -> Result<Self, quick_xml::Error> {
        let name = String::from_utf8_lossy(start.name().as_ref()).into_owned();
        let attributes: Vec<(String, String)> = start
            .attributes()
            .flatten()
            .map(|a| {
                let key = String::from_utf8_lossy(a.key.as_ref()).into_owned();
                let val = String::from_utf8_lossy(&a.value).into_owned();
                (key, val)
            })
            .collect();

        let mut children = Vec::new();
        let mut buf = Vec::new();

        loop {
            match reader.read_event_into(&mut buf)? {
                Event::Start(ref e) => {
                    // Recursively capture nested elements
                    children.push(RawXmlNode::read_element(reader, e)?);
                }
                Event::Empty(ref e) => {
                    let child_name = String::from_utf8_lossy(e.name().as_ref()).into_owned();
                    let child_attrs: Vec<(String, String)> = e
                        .attributes()
                        .flatten()
                        .map(|a| {
                            let key = String::from_utf8_lossy(a.key.as_ref()).into_owned();
                            let val = String::from_utf8_lossy(&a.value).into_owned();
                            (key, val)
                        })
                        .collect();
                    children.push(RawXmlNode::Element {
                        name: child_name,
                        attributes: child_attrs,
                        children: Vec::new(),
                    });
                }
                Event::Text(ref e) => {
                    let text = e.unescape()?.into_owned();
                    if !text.trim().is_empty() {
                        children.push(RawXmlNode::Text(text));
                    }
                }
                Event::CData(ref e) => {
                    let text = String::from_utf8_lossy(e.as_ref()).into_owned();
                    children.push(RawXmlNode::CData(text));
                }
                Event::End(ref e) if e.name().as_ref() == start.name().as_ref() => {
                    break;
                }
                Event::Eof => break,
                _ => {} // Comments, PI, etc. — could preserve these too
            }
            buf.clear();
        }

        Ok(RawXmlNode::Element {
            name,
            attributes,
            children,
        })
    }

    /// Read a self-closing (empty) element as a `RawXmlNode`.
    pub fn from_empty_element(e: &BytesStart<'_>) -> Self {
        let name = String::from_utf8_lossy(e.name().as_ref()).into_owned();
        let attributes: Vec<(String, String)> = e
            .attributes()
            .flatten()
            .map(|a| {
                let key = String::from_utf8_lossy(a.key.as_ref()).into_owned();
                let val = String::from_utf8_lossy(&a.value).into_owned();
                (key, val)
            })
            .collect();
        RawXmlNode::Element {
            name,
            attributes,
            children: Vec::new(),
        }
    }

    /// Write this node to an XML writer.
    pub fn write_to<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), quick_xml::Error> {
        match self {
            RawXmlNode::Element {
                name,
                attributes,
                children,
            } => {
                let mut start = BytesStart::new(name.as_str());
                for (key, val) in attributes {
                    start.push_attribute((key.as_str(), val.as_str()));
                }

                if children.is_empty() {
                    writer.write_event(Event::Empty(start))?;
                } else {
                    writer.write_event(Event::Start(start))?;
                    for child in children {
                        child.write_to(writer)?;
                    }
                    writer.write_event(Event::End(BytesEnd::new(name.as_str())))?;
                }
            }
            RawXmlNode::Text(text) => {
                writer.write_event(Event::Text(BytesText::new(text)))?;
            }
            RawXmlNode::CData(text) => {
                writer.write_event(Event::CData(quick_xml::events::BytesCData::new(text)))?;
            }
            RawXmlNode::RawBytes(bytes) => {
                writer.get_mut().write_all(bytes)?;
            }
        }
        Ok(())
    }
}
