//! Content type resolution per OPC spec.
//!
//! `[Content_Types].xml` maps part URIs to their MIME-like content types,
//! either by exact path (Override) or by file extension (Default).

use quick_xml::events::{BytesDecl, BytesEnd, BytesStart, Event};
use quick_xml::{Reader, Writer};
use std::collections::HashMap;
use std::io::{BufRead, Write};

use crate::error::Result;

/// Content type registry for an OPC package.
#[derive(Debug, Clone, Default)]
pub struct ContentTypes {
    /// Extension → content type (e.g., "xml" → "application/xml")
    defaults: HashMap<String, String>,
    /// Absolute part URI → content type
    overrides: HashMap<String, String>,
    /// Original `[Content_Types].xml` bytes as loaded from the package.
    raw_xml: Option<Vec<u8>>,
    /// Whether defaults/overrides changed since parse or construction.
    dirty: bool,
}

impl ContentTypes {
    pub fn new() -> Self {
        let mut ct = Self {
            defaults: HashMap::new(),
            overrides: HashMap::new(),
            raw_xml: None,
            dirty: true,
        };

        for (extension, content_type) in [
            // Standard defaults
            (
                "rels",
                "application/vnd.openxmlformats-package.relationships+xml",
            ),
            ("xml", "application/xml"),
            // Common image media defaults
            ("bmp", "image/bmp"),
            ("gif", "image/gif"),
            ("jpeg", "image/jpeg"),
            ("jpg", "image/jpeg"),
            ("png", "image/png"),
            ("svg", "image/svg+xml"),
            ("tif", "image/tiff"),
            ("tiff", "image/tiff"),
            // Common audio/video media defaults
            ("mp3", "audio/mpeg"),
            ("wav", "audio/wav"),
            ("mp4", "video/mp4"),
            ("mpeg", "video/mpeg"),
            ("ogg", "audio/ogg"),
            ("wmv", "video/x-ms-wmv"),
        ] {
            ct.defaults
                .insert(extension.to_string(), content_type.to_string());
        }

        ct
    }

    /// Parse from raw `[Content_Types].xml` bytes while preserving exact input for no-op writes.
    pub fn from_xml_bytes(bytes: Vec<u8>) -> Result<Self> {
        let mut parsed = Self::from_xml(std::io::Cursor::new(bytes.as_slice()))?;
        parsed.raw_xml = Some(bytes);
        parsed.dirty = false;
        Ok(parsed)
    }

    /// Parse from `[Content_Types].xml`.
    pub fn from_xml<R: BufRead>(reader: R) -> Result<Self> {
        let mut xml = Reader::from_reader(reader);
        xml.config_mut().trim_text(true);

        let mut ct = ContentTypes {
            defaults: HashMap::new(),
            overrides: HashMap::new(),
            raw_xml: None,
            dirty: false,
        };
        let mut buf = Vec::new();

        loop {
            match xml.read_event_into(&mut buf) {
                Ok(Event::Empty(ref e)) => match e.name().as_ref() {
                    name if local_name(name) == b"Default" => {
                        let mut ext = String::new();
                        let mut content_type = String::new();
                        for attr in e.attributes().flatten() {
                            match attr.key.as_ref() {
                                b"Extension" => {
                                    ext =
                                        attr.decode_and_unescape_value(xml.decoder())?.into_owned()
                                }
                                b"ContentType" => {
                                    content_type =
                                        attr.decode_and_unescape_value(xml.decoder())?.into_owned()
                                }
                                _ => {}
                            }
                        }
                        ct.defaults.insert(ext, content_type);
                    }
                    name if local_name(name) == b"Override" => {
                        let mut part_name = String::new();
                        let mut content_type = String::new();
                        for attr in e.attributes().flatten() {
                            match attr.key.as_ref() {
                                b"PartName" => {
                                    part_name =
                                        attr.decode_and_unescape_value(xml.decoder())?.into_owned()
                                }
                                b"ContentType" => {
                                    content_type =
                                        attr.decode_and_unescape_value(xml.decoder())?.into_owned()
                                }
                                _ => {}
                            }
                        }
                        ct.overrides.insert(part_name, content_type);
                    }
                    _ => {}
                },
                Ok(Event::Eof) => break,
                Err(e) => return Err(e.into()),
                _ => {}
            }
            buf.clear();
        }

        Ok(ct)
    }

    /// Write to `[Content_Types].xml`.
    pub fn to_xml<W: Write>(&self, writer: W) -> Result<()> {
        if !self.dirty {
            if let Some(raw_xml) = &self.raw_xml {
                let mut writer = writer;
                writer.write_all(raw_xml)?;
                return Ok(());
            }
        }

        let mut xml = Writer::new_with_indent(writer, b' ', 2);

        xml.write_event(Event::Decl(BytesDecl::new(
            "1.0",
            Some("UTF-8"),
            Some("yes"),
        )))?;

        let mut root = BytesStart::new("Types");
        root.push_attribute((
            "xmlns",
            "http://schemas.openxmlformats.org/package/2006/content-types",
        ));
        xml.write_event(Event::Start(root))?;

        // Write defaults (sorted for deterministic output)
        let mut defaults: Vec<_> = self.defaults.iter().collect();
        defaults.sort_by(|(a, _), (b, _)| a.cmp(b));
        for (ext, content_type) in defaults {
            let mut elem = BytesStart::new("Default");
            elem.push_attribute(("Extension", ext.as_str()));
            elem.push_attribute(("ContentType", content_type.as_str()));
            xml.write_event(Event::Empty(elem))?;
        }

        // Write overrides (sorted for deterministic output)
        let mut overrides: Vec<_> = self.overrides.iter().collect();
        overrides.sort_by(|(a, _), (b, _)| a.cmp(b));
        for (part_name, content_type) in overrides {
            let mut elem = BytesStart::new("Override");
            elem.push_attribute(("PartName", part_name.as_str()));
            elem.push_attribute(("ContentType", content_type.as_str()));
            xml.write_event(Event::Empty(elem))?;
        }

        xml.write_event(Event::End(BytesEnd::new("Types")))?;

        Ok(())
    }

    /// Resolve the content type for a given part URI.
    pub fn get(&self, part_uri: &str) -> Option<&str> {
        // Check overrides first (exact match)
        if let Some(ct) = self.overrides.get(part_uri) {
            return Some(ct);
        }

        // Fall back to extension-based default
        let ext = part_uri.rsplit('.').next()?;
        self.defaults.get(ext).map(|s| s.as_str())
    }

    /// Whether a default exists for a file extension.
    pub fn has_default(&self, extension: &str) -> bool {
        self.defaults.contains_key(extension)
    }

    /// Get a default content type by file extension.
    pub fn get_default(&self, extension: &str) -> Option<&str> {
        self.defaults.get(extension).map(String::as_str)
    }

    /// Whether an override exists for an exact part URI.
    pub fn has_override(&self, part_uri: &str) -> bool {
        self.overrides.contains_key(part_uri)
    }

    /// Get an override content type by exact part URI.
    pub fn get_override(&self, part_uri: &str) -> Option<&str> {
        self.overrides.get(part_uri).map(String::as_str)
    }

    /// Add an extension-based default.
    pub fn add_default(&mut self, extension: impl Into<String>, content_type: impl Into<String>) {
        let extension = extension.into();
        let content_type = content_type.into();
        if self.defaults.get(&extension) != Some(&content_type) {
            self.dirty = true;
            self.defaults.insert(extension, content_type);
        }
    }

    /// Add a part-specific override.
    pub fn add_override(&mut self, part_uri: impl Into<String>, content_type: impl Into<String>) {
        let part_uri = part_uri.into();
        let content_type = content_type.into();
        if self.overrides.get(&part_uri) != Some(&content_type) {
            self.dirty = true;
            self.overrides.insert(part_uri, content_type);
        }
    }

    /// Remove a part-specific override, returning the content type if it existed.
    pub fn remove_override(&mut self, part_uri: &str) -> Option<String> {
        let removed = self.overrides.remove(part_uri);
        if removed.is_some() {
            self.dirty = true;
        }
        removed
    }

    /// Remove an extension-based default, returning the content type if it existed.
    pub fn remove_default(&mut self, extension: &str) -> Option<String> {
        let removed = self.defaults.remove(extension);
        if removed.is_some() {
            self.dirty = true;
        }
        removed
    }

    /// Iterate all extension-based defaults.
    pub fn defaults_iter(&self) -> impl Iterator<Item = (&str, &str)> {
        self.defaults.iter().map(|(k, v)| (k.as_str(), v.as_str()))
    }

    /// Iterate all part-specific overrides.
    pub fn overrides_iter(&self) -> impl Iterator<Item = (&str, &str)> {
        self.overrides.iter().map(|(k, v)| (k.as_str(), v.as_str()))
    }

    /// Whether this content type map has been modified.
    pub fn is_dirty(&self) -> bool {
        self.dirty
    }

    /// Mark content types as clean (used when loading malformed packages that
    /// omit `[Content_Types].xml` and should be preserved as-is on no-op save).
    pub fn mark_clean(&mut self) {
        self.dirty = false;
    }

    /// Number of extension-based defaults.
    pub fn default_count(&self) -> usize {
        self.defaults.len()
    }

    /// Number of part-specific overrides.
    pub fn override_count(&self) -> usize {
        self.overrides.len()
    }

    /// Validate that a content type string has a valid MIME-like format.
    ///
    /// A valid content type has the form `type/subtype` (possibly with parameters
    /// like `; charset=utf-8`). Both type and subtype must be non-empty.
    pub fn is_valid_content_type(content_type: &str) -> bool {
        // Strip parameters (everything after ;)
        let base = content_type.split(';').next().unwrap_or("");
        let base = base.trim();
        if base.is_empty() {
            return false;
        }
        match base.split_once('/') {
            Some((media_type, subtype)) => {
                let media_type = media_type.trim();
                let subtype = subtype.trim();
                !media_type.is_empty() && !subtype.is_empty()
            }
            None => false,
        }
    }
}

fn local_name(name: &[u8]) -> &[u8] {
    name.rsplit(|byte| *byte == b':').next().unwrap_or(name)
}

/// Well-known content types for OOXML parts.
pub struct ContentTypeValue;

impl ContentTypeValue {
    // SpreadsheetML
    pub const WORKBOOK: &str =
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml";
    pub const WORKSHEET: &str =
        "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml";
    pub const SHARED_STRINGS: &str =
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml";
    pub const SPREADSHEET_STYLES: &str =
        "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml";

    // WordprocessingML
    pub const WORD_DOCUMENT: &str =
        "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml";
    pub const WORD_STYLES: &str =
        "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml";
    pub const WORD_SETTINGS: &str =
        "application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml";
    pub const WORD_NUMBERING: &str =
        "application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml";

    // PresentationML
    pub const PRESENTATION: &str =
        "application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml";
    pub const SLIDE: &str =
        "application/vnd.openxmlformats-officedocument.presentationml.slide+xml";
    pub const SLIDE_LAYOUT: &str =
        "application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml";
    pub const SLIDE_MASTER: &str =
        "application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml";

    // Shared
    pub const THEME: &str = "application/vnd.openxmlformats-officedocument.theme+xml";
    pub const CORE_PROPERTIES: &str = "application/vnd.openxmlformats-package.core-properties+xml";

    // SpreadsheetML (additional)
    pub const CHARTSHEET: &str =
        "application/vnd.openxmlformats-officedocument.spreadsheetml.chartsheet+xml";
    pub const SPREADSHEET_COMMENTS: &str =
        "application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml";
    pub const SPREADSHEET_TABLE: &str =
        "application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml";
    pub const PIVOT_CACHE_DEFINITION: &str =
        "application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheDefinition+xml";
    pub const PIVOT_CACHE_RECORDS: &str =
        "application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheRecords+xml";
    pub const EXTERNAL_LINK: &str =
        "application/vnd.openxmlformats-officedocument.spreadsheetml.externalLink+xml";
    pub const CALC_CHAIN: &str =
        "application/vnd.openxmlformats-officedocument.spreadsheetml.calcChain+xml";
    pub const MACRO_ENABLED_WORKBOOK: &str = "application/vnd.ms-excel.sheet.macroEnabled.main+xml";
    pub const SPREADSHEET_TEMPLATE: &str =
        "application/vnd.openxmlformats-officedocument.spreadsheetml.template.main+xml";

    // WordprocessingML (additional)
    pub const WORD_HEADER: &str =
        "application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml";
    pub const WORD_FOOTER: &str =
        "application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml";
    pub const WORD_FOOTNOTES: &str =
        "application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml";
    pub const WORD_ENDNOTES: &str =
        "application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml";
    pub const WORD_COMMENTS: &str =
        "application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml";
    pub const WORD_FONT_TABLE: &str =
        "application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml";
    pub const WORD_WEB_SETTINGS: &str =
        "application/vnd.openxmlformats-officedocument.wordprocessingml.webSettings+xml";
    pub const MACRO_ENABLED_DOCUMENT: &str =
        "application/vnd.ms-word.document.macroEnabled.main+xml";
    pub const WORD_TEMPLATE: &str =
        "application/vnd.openxmlformats-officedocument.wordprocessingml.template.main+xml";

    // PresentationML (additional)
    pub const NOTES_SLIDE: &str =
        "application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml";
    pub const NOTES_MASTER: &str =
        "application/vnd.openxmlformats-officedocument.presentationml.notesMaster+xml";
    pub const HANDOUT_MASTER: &str =
        "application/vnd.openxmlformats-officedocument.presentationml.handoutMaster+xml";
    pub const PRESENTATION_COMMENTS: &str =
        "application/vnd.openxmlformats-officedocument.presentationml.comments+xml";
    pub const SLIDE_SHOW: &str =
        "application/vnd.openxmlformats-officedocument.presentationml.slideshow.main+xml";
    pub const MACRO_ENABLED_PRESENTATION: &str =
        "application/vnd.ms-powerpoint.presentation.macroEnabled.main+xml";
    pub const PRESENTATION_TEMPLATE: &str =
        "application/vnd.openxmlformats-officedocument.presentationml.template.main+xml";
    pub const TAGS: &str = "application/vnd.openxmlformats-officedocument.presentationml.tags+xml";

    // DrawingML
    pub const DRAWING: &str = "application/vnd.openxmlformats-officedocument.drawing+xml";
    pub const CHART: &str = "application/vnd.openxmlformats-officedocument.drawingml.chart+xml";
    pub const DIAGRAM_COLORS: &str =
        "application/vnd.openxmlformats-officedocument.drawingml.diagramColors+xml";
    pub const DIAGRAM_DATA: &str =
        "application/vnd.openxmlformats-officedocument.drawingml.diagramData+xml";
    pub const DIAGRAM_LAYOUT: &str =
        "application/vnd.openxmlformats-officedocument.drawingml.diagramLayout+xml";
    pub const DIAGRAM_STYLE: &str =
        "application/vnd.openxmlformats-officedocument.drawingml.diagramStyle+xml";
    pub const THEME_OVERRIDE: &str =
        "application/vnd.openxmlformats-officedocument.themeOverride+xml";

    // Shared
    pub const CUSTOM_PROPERTIES: &str =
        "application/vnd.openxmlformats-officedocument.custom-properties+xml";
    pub const CUSTOM_XML_PROPERTIES: &str =
        "application/vnd.openxmlformats-officedocument.customXmlProperties+xml";
    pub const EXTENDED_PROPERTIES: &str =
        "application/vnd.openxmlformats-officedocument.extended-properties+xml";
    pub const VML_DRAWING: &str = "application/vnd.openxmlformats-officedocument.vmlDrawing";
}

#[cfg(test)]
mod tests {
    use super::ContentTypes;
    use quick_xml::events::Event;
    use quick_xml::Reader;

    #[test]
    fn new_resolves_standard_and_common_media_defaults() {
        let ct = ContentTypes::new();

        assert_eq!(
            ct.get("/_rels/.rels"),
            Some("application/vnd.openxmlformats-package.relationships+xml")
        );
        assert_eq!(ct.get("/docProps/core.xml"), Some("application/xml"));

        for (part_uri, expected) in [
            ("/word/media/image.bmp", "image/bmp"),
            ("/word/media/image.gif", "image/gif"),
            ("/word/media/image.jpg", "image/jpeg"),
            ("/word/media/image.jpeg", "image/jpeg"),
            ("/word/media/image.png", "image/png"),
            ("/word/media/image.svg", "image/svg+xml"),
            ("/word/media/image.tif", "image/tiff"),
            ("/word/media/image.tiff", "image/tiff"),
            ("/word/media/audio.mp3", "audio/mpeg"),
            ("/word/media/audio.wav", "audio/wav"),
            ("/word/media/video.mp4", "video/mp4"),
            ("/word/media/video.mpeg", "video/mpeg"),
            ("/word/media/audio.ogg", "audio/ogg"),
            ("/word/media/video.wmv", "video/x-ms-wmv"),
        ] {
            assert_eq!(ct.get(part_uri), Some(expected), "part_uri={part_uri}");
        }
    }

    #[test]
    fn override_takes_precedence_over_default() {
        let mut ct = ContentTypes::new();
        ct.add_override("/word/media/image.png", "application/custom-image");

        assert_eq!(
            ct.get("/word/media/image.png"),
            Some("application/custom-image")
        );
        assert_eq!(ct.get("/word/media/other.png"), Some("image/png"));
    }

    #[test]
    fn helper_methods_report_default_and_override_presence() {
        let mut ct = ContentTypes::new();

        assert!(ct.has_default("png"));
        assert_eq!(ct.get_default("png"), Some("image/png"));
        assert!(!ct.has_default("unknown"));
        assert_eq!(ct.get_default("unknown"), None);

        let part = "/ppt/slides/slide1.xml";
        assert!(!ct.has_override(part));
        assert_eq!(ct.get_override(part), None);

        ct.add_override(part, "application/custom-slide+xml");
        assert!(ct.has_override(part));
        assert_eq!(ct.get_override(part), Some("application/custom-slide+xml"));
    }

    #[test]
    fn from_xml_bytes_preserves_original_bytes_on_noop_write() {
        let original = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
</Types>"#;

        let ct =
            ContentTypes::from_xml_bytes(original.to_vec()).expect("content types should parse");
        let mut serialized = Vec::new();
        ct.to_xml(&mut serialized)
            .expect("no-op write should succeed");

        assert_eq!(
            serialized, original,
            "no-op serialization should preserve exact original [Content_Types].xml bytes"
        );
    }

    #[test]
    fn mutation_marks_content_types_dirty_and_emits_updated_xml() {
        let original = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="xml" ContentType="application/xml"/>
</Types>"#;
        let mut ct =
            ContentTypes::from_xml_bytes(original.to_vec()).expect("content types should parse");
        ct.add_override(
            "/word/document.xml",
            "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml",
        );

        let mut serialized = Vec::new();
        ct.to_xml(&mut serialized)
            .expect("dirty write should succeed");
        assert_ne!(serialized, original, "dirty write should regenerate xml");

        let mut reader = Reader::from_reader(serialized.as_slice());
        let mut saw_override = false;
        let mut buf = Vec::new();
        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Empty(event)) if event.name().as_ref() == b"Override" => {
                    saw_override = true;
                }
                Ok(Event::Eof) => break,
                Ok(_) => {}
                Err(error) => panic!("serialized XML should parse: {error}"),
            }
            buf.clear();
        }
        assert!(saw_override, "serialized xml should include new override");
    }
}
