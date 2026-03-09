//! Relationship handling per OPC spec.
//!
//! Every part can have associated relationships stored in a `.rels` file.
//! Relationships define how parts connect — e.g., a workbook has relationships
//! to its worksheets, shared strings, styles, etc.

use quick_xml::events::{BytesDecl, BytesEnd, BytesStart, Event};
use quick_xml::{Reader, Writer};
use serde::{Deserialize, Serialize};
use std::collections::HashMap;
use std::io::{BufRead, Write};

use crate::error::Result;

/// A single relationship within an OPC package.
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct Relationship {
    /// Unique ID within the .rels file (e.g., "rId1").
    pub id: String,

    /// The relationship type URI (defines the semantic meaning).
    pub rel_type: String,

    /// Target URI (relative to the source part, or absolute).
    pub target: String,

    /// Whether the target is external (URL) vs internal (part).
    pub target_mode: TargetMode,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize, Default)]
pub enum TargetMode {
    #[default]
    Internal,
    External,
}

/// A collection of relationships from a single .rels file.
#[derive(Debug, Clone)]
pub struct Relationships {
    rels: Vec<Relationship>,
    by_id: HashMap<String, usize>,
    by_type: HashMap<String, Vec<usize>>,
    next_id: u32,
    /// Original relationship XML bytes as loaded from a package.
    raw_xml: Option<Vec<u8>>,
    /// Whether relationships have changed since parse/construction.
    dirty: bool,
}

impl Default for Relationships {
    fn default() -> Self {
        Self {
            rels: Vec::new(),
            by_id: HashMap::new(),
            by_type: HashMap::new(),
            next_id: 1,
            raw_xml: None,
            dirty: false,
        }
    }
}

impl Relationships {
    pub fn new() -> Self {
        Self::default()
    }

    /// Parse relationships from raw `.rels` bytes while preserving exact input for no-op writes.
    pub fn from_xml_bytes(bytes: Vec<u8>) -> Result<Self> {
        let mut parsed = Self::from_xml(std::io::Cursor::new(bytes.as_slice()))?;
        parsed.raw_xml = Some(bytes);
        parsed.dirty = false;
        Ok(parsed)
    }

    /// Parse relationships from a .rels XML stream.
    pub fn from_xml<R: BufRead>(reader: R) -> Result<Self> {
        let mut xml = Reader::from_reader(reader);
        xml.config_mut().trim_text(true);

        let mut rels = Relationships::new();
        let mut buf = Vec::new();

        loop {
            match xml.read_event_into(&mut buf) {
                Ok(Event::Empty(ref e)) | Ok(Event::Start(ref e))
                    if local_name(e.name().as_ref()) == b"Relationship" =>
                {
                    let mut id = String::new();
                    let mut rel_type = String::new();
                    let mut target = String::new();
                    let mut target_mode = TargetMode::Internal;

                    for attr in e.attributes().flatten() {
                        match attr.key.as_ref() {
                            b"Id" => {
                                id = attr.decode_and_unescape_value(xml.decoder())?.into_owned()
                            }
                            b"Type" => {
                                rel_type =
                                    attr.decode_and_unescape_value(xml.decoder())?.into_owned()
                            }
                            b"Target" => {
                                target = attr.decode_and_unescape_value(xml.decoder())?.into_owned()
                            }
                            b"TargetMode" => {
                                if attr.decode_and_unescape_value(xml.decoder())?.as_ref()
                                    == "External"
                                {
                                    target_mode = TargetMode::External;
                                }
                            }
                            _ => {}
                        }
                    }

                    rels.add_internal(
                        Relationship {
                            id,
                            rel_type,
                            target,
                            target_mode,
                        },
                        false,
                    );
                }
                Ok(Event::Eof) => break,
                Err(e) => return Err(e.into()),
                _ => {}
            }
            buf.clear();
        }

        rels.dirty = false;
        Ok(rels)
    }

    /// Write relationships to XML.
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

        let mut root = BytesStart::new("Relationships");
        root.push_attribute((
            "xmlns",
            "http://schemas.openxmlformats.org/package/2006/relationships",
        ));
        xml.write_event(Event::Start(root))?;

        for rel in &self.rels {
            let mut elem = BytesStart::new("Relationship");
            elem.push_attribute(("Id", rel.id.as_str()));
            elem.push_attribute(("Type", rel.rel_type.as_str()));
            elem.push_attribute(("Target", rel.target.as_str()));
            if rel.target_mode == TargetMode::External {
                elem.push_attribute(("TargetMode", "External"));
            }
            xml.write_event(Event::Empty(elem))?;
        }

        xml.write_event(Event::End(BytesEnd::new("Relationships")))?;

        Ok(())
    }

    /// Add a relationship.
    pub fn add(&mut self, rel: Relationship) {
        self.add_internal(rel, true);
    }

    fn add_internal(&mut self, rel: Relationship, mark_dirty: bool) {
        let idx = self.rels.len();

        // Track the highest rId number for generating new IDs
        if let Some(num) = rel
            .id
            .strip_prefix("rId")
            .and_then(|s| s.parse::<u32>().ok())
        {
            self.next_id = self.next_id.max(num + 1);
        }

        self.by_id.insert(rel.id.clone(), idx);
        self.by_type
            .entry(rel.rel_type.clone())
            .or_default()
            .push(idx);
        self.rels.push(rel);
        if mark_dirty {
            self.dirty = true;
        }
    }

    /// Create a new relationship with an auto-generated ID.
    pub fn add_new(
        &mut self,
        rel_type: String,
        target: String,
        target_mode: TargetMode,
    ) -> &Relationship {
        let idx = self.rels.len();
        let id = format!("rId{}", self.next_id);
        self.add_internal(
            Relationship {
                id,
                rel_type,
                target,
                target_mode,
            },
            true,
        );
        &self.rels[idx]
    }

    /// Get a relationship by ID.
    pub fn get_by_id(&self, id: &str) -> Option<&Relationship> {
        self.by_id.get(id).map(|&idx| &self.rels[idx])
    }

    /// Check whether a relationship ID exists.
    pub fn contains_id(&self, id: &str) -> bool {
        self.by_id.contains_key(id)
    }

    /// Get all relationships of a given type.
    pub fn get_by_type(&self, rel_type: &str) -> Vec<&Relationship> {
        self.by_type
            .get(rel_type)
            .map(|indices| indices.iter().map(|&idx| &self.rels[idx]).collect())
            .unwrap_or_default()
    }

    /// Get the first relationship of a given type (common case).
    pub fn get_first_by_type(&self, rel_type: &str) -> Option<&Relationship> {
        self.get_by_type(rel_type).into_iter().next()
    }

    /// Iterate all relationships.
    pub fn iter(&self) -> impl Iterator<Item = &Relationship> {
        self.rels.iter()
    }

    pub fn len(&self) -> usize {
        self.rels.len()
    }

    pub fn is_empty(&self) -> bool {
        self.rels.is_empty()
    }

    /// Whether this relationship part should be written to the package.
    ///
    /// Empty `.rels` files are valid and should be preserved if they existed
    /// in the source package.
    pub fn should_write_xml(&self) -> bool {
        !self.rels.is_empty() || self.raw_xml.is_some()
    }

    /// Remove all relationships, resetting the collection to empty.
    pub fn clear(&mut self) {
        if !self.rels.is_empty() {
            self.dirty = true;
        }
        self.rels.clear();
        self.by_id.clear();
        self.by_type.clear();
        self.next_id = 1;
    }

    /// Remove a relationship by its ID. Returns the removed relationship, or `None` if not found.
    pub fn remove_by_id(&mut self, id: &str) -> Option<Relationship> {
        let &idx = self.by_id.get(id)?;

        // Remove from the rels vec
        let removed = self.rels.remove(idx);
        self.dirty = true;

        // Rebuild indexes since removal shifts indices
        self.rebuild_indexes();

        Some(removed)
    }

    /// Remove all relationships of a given type. Returns the removed relationships.
    pub fn remove_by_type(&mut self, rel_type: &str) -> Vec<Relationship> {
        let indices: Vec<usize> = self.by_type.get(rel_type).cloned().unwrap_or_default();

        if indices.is_empty() {
            return Vec::new();
        }

        self.dirty = true;

        // Remove in reverse order to preserve indices
        let mut removed = Vec::with_capacity(indices.len());
        let mut sorted_indices = indices;
        sorted_indices.sort_unstable();
        for &idx in sorted_indices.iter().rev() {
            removed.push(self.rels.remove(idx));
        }
        removed.reverse();

        // Rebuild indexes
        self.rebuild_indexes();

        removed
    }

    /// Rebuild the by_id and by_type indexes from the rels vector.
    fn rebuild_indexes(&mut self) {
        self.by_id.clear();
        self.by_type.clear();
        for (idx, rel) in self.rels.iter().enumerate() {
            self.by_id.insert(rel.id.clone(), idx);
            self.by_type
                .entry(rel.rel_type.clone())
                .or_default()
                .push(idx);
        }
    }
}

fn local_name(name: &[u8]) -> &[u8] {
    name.rsplit(|byte| *byte == b':').next().unwrap_or(name)
}

/// Well-known relationship type URIs.
///
/// All 122 relationship types from Open-XML-SDK, covering ECMA-376 standard
/// types and Microsoft Office extension types through Office 2022.
pub struct RelationshipType;

impl RelationshipType {
    // ── Package-level (OPC core) ──────────────────────────────────────
    pub const CORE_PROPERTIES: &str =
        "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties";
    pub const EXTENDED_PROPERTIES: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties";
    pub const THUMBNAIL: &str =
        "http://schemas.openxmlformats.org/package/2006/relationships/metadata/thumbnail";

    // ── Core office document ──────────────────────────────────────────
    pub const OFFICE_DOCUMENT: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument";
    /// Alias: the main workbook relationship uses officeDocument.
    pub const WORKBOOK: &str = Self::OFFICE_DOCUMENT;
    /// Alias: the main Word document relationship uses officeDocument.
    pub const WORD_DOCUMENT: &str = Self::OFFICE_DOCUMENT;
    /// Alias: the main presentation relationship uses officeDocument.
    pub const PRESENTATION: &str = Self::OFFICE_DOCUMENT;
    pub const CUSTOM_PROPERTIES: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties";

    // ── SpreadsheetML ─────────────────────────────────────────────────
    pub const WORKSHEET: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet";
    pub const SHARED_STRINGS: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings";
    pub const STYLES: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles";
    pub const THEME: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme";
    pub const CHARTSHEET: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chartsheet";
    pub const DIALOG_SHEET: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/dialogsheet";
    pub const PIVOT_TABLE: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotTable";
    pub const PIVOT_CACHE_DEFINITION: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheDefinition";
    pub const PIVOT_CACHE_RECORDS: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheRecords";
    pub const CALC_CHAIN: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/calcChain";
    pub const QUERY_TABLE: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/queryTable";
    pub const TABLE_SINGLE_CELLS: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/tableSingleCells";
    pub const SHEET_METADATA: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sheetMetadata";
    pub const EXTERNAL_LINK: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLink";
    pub const CONNECTIONS: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/connections";
    pub const VOLATILE_DEPENDENCIES: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/volatileDependencies";

    // ── WordprocessingML ──────────────────────────────────────────────
    pub const WORD_SETTINGS: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings";
    pub const WORD_NUMBERING: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering";
    pub const WORD_FOOTNOTES: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes";
    pub const WORD_ENDNOTES: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes";
    pub const WORD_HEADER: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/header";
    pub const WORD_FOOTER: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer";
    pub const WORD_WEB_SETTINGS: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/webSettings";
    pub const FONT_TABLE: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable";
    pub const GLOSSARY_DOCUMENT: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/glossaryDocument";

    // ── PresentationML ────────────────────────────────────────────────
    pub const SLIDE: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide";
    pub const SLIDE_LAYOUT: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout";
    pub const SLIDE_MASTER: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster";
    pub const NOTES_SLIDE: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide";
    pub const NOTES_MASTER: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesMaster";
    pub const HANDOUT_MASTER: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/handoutMaster";
    pub const PRES_PROPS: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/presProps";
    pub const VIEW_PROPS: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/viewProps";
    pub const TAGS: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/tags";
    pub const SLIDE_UPDATE_INFO: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideUpdateInfo";
    pub const COMMENT_AUTHORS: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/commentAuthors";

    // ── Shared media & external content ───────────────────────────────
    pub const IMAGE: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";
    pub const AUDIO: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/audio";
    pub const VIDEO: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/video";
    pub const HYPERLINK: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink";
    pub const OLE_OBJECT: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/oleObject";
    pub const DRAWING: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing";
    pub const PACKAGE: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/package";
    pub const CONTROL: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/control";
    pub const CTRL_PROP: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/ctrlProp";
    pub const PRINTER_SETTINGS: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/printerSettings";

    // ── Shared chart & drawing ────────────────────────────────────────
    pub const CHART: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart";
    pub const CHART_USER_SHAPES: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chartUserShapes";
    pub const TABLE: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/table";
    pub const TABLE_STYLES: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/tableStyles";
    pub const COMMENTS: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments";
    pub const VML_DRAWING: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing";
    pub const THEME_OVERRIDE: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/themeOverride";

    // ── Custom XML & user data ────────────────────────────────────────
    pub const CUSTOM_XML: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml";
    pub const CUSTOM_XML_PROPERTIES: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXmlProps";
    pub const XML_MAPS: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/xmlMaps";
    pub const CUSTOM_PROPERTY: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/customProperty";
    pub const AF_CHUNK: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/aFChunk";
    pub const USERNAMES: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/usernames";
    pub const REVISION_HEADERS: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/revisionHeaders";
    pub const REVISION_LOG: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/revisionLog";
    pub const RECIPIENT_DATA: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/recipientData";

    // ── Diagram relationships ─────────────────────────────────────────
    pub const DIAGRAM_COLORS: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramColors";
    pub const DIAGRAM_DATA: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramData";
    pub const DIAGRAM_LAYOUT: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramLayoutDefinition";
    pub const DIAGRAM_STYLE: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramStyle";
    pub const DIAGRAM_QUICK_STYLE: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramQuickStyle";

    // ── Digital signature (OPC) ───────────────────────────────────────
    pub const DIGITAL_SIGNATURE: &str =
        "http://schemas.openxmlformats.org/package/2006/relationships/digital-signature/signature";
    pub const DIGITAL_SIGNATURE_ORIGIN: &str =
        "http://schemas.openxmlformats.org/package/2006/relationships/digital-signature/origin";
    pub const DIGITAL_SIGNATURE_CERTIFICATE: &str =
        "http://schemas.openxmlformats.org/package/2006/relationships/digital-signature/certificate";

    // ── Microsoft extensions (2006) ───────────────────────────────────
    pub const MACRO: &str = "http://schemas.microsoft.com/office/2006/relationships/vbaProject";
    pub const WORD_VBA_DATA: &str =
        "http://schemas.microsoft.com/office/2006/relationships/wordVbaData";
    pub const ACTIVE_X_CONTROL_BINARY: &str =
        "http://schemas.microsoft.com/office/2006/relationships/activeXControlBinary";
    pub const ATTACHED_TOOLBARS: &str =
        "http://schemas.microsoft.com/office/2006/relationships/attachedToolbars";
    pub const UI_EXTENSIBILITY: &str =
        "http://schemas.microsoft.com/office/2006/relationships/ui/extensibility";
    pub const USER_CUSTOMIZATION: &str =
        "http://schemas.microsoft.com/office/2006/relationships/ui/userCustomization";
    pub const KEY_MAP_CUSTOMIZATIONS: &str =
        "http://schemas.microsoft.com/office/2006/relationships/keyMapCustomizations";
    pub const LEGACY_DIAGRAM_TEXT: &str =
        "http://schemas.microsoft.com/office/2006/relationships/legacyDiagramText";
    pub const LEGACY_DOC_TEXT_INFO: &str =
        "http://schemas.microsoft.com/office/2006/relationships/legacyDocTextInfo";
    pub const XL_MACROSHEET: &str =
        "http://schemas.microsoft.com/office/2006/relationships/xlMacrosheet";
    pub const XL_INTL_MACROSHEET: &str =
        "http://schemas.microsoft.com/office/2006/relationships/xlIntlMacrosheet";
    pub const WS_SORT_MAP: &str =
        "http://schemas.microsoft.com/office/2006/relationships/wsSortMap";

    // ── Microsoft extensions (2007) ───────────────────────────────────
    pub const MEDIA: &str = "http://schemas.microsoft.com/office/2007/relationships/media";
    pub const DIAGRAM_DRAWING: &str =
        "http://schemas.microsoft.com/office/2007/relationships/diagramDrawing";
    pub const CUSTOM_DATA: &str =
        "http://schemas.microsoft.com/office/2007/relationships/customData";
    pub const CUSTOM_DATA_PROPS: &str =
        "http://schemas.microsoft.com/office/2007/relationships/customDataProps";
    pub const UI_EXTENSIBILITY_2007: &str =
        "http://schemas.microsoft.com/office/2007/relationships/ui/extensibility";
    pub const SLICER: &str = "http://schemas.microsoft.com/office/2007/relationships/slicer";
    pub const SLICER_CACHE: &str =
        "http://schemas.microsoft.com/office/2007/relationships/slicerCache";
    pub const STYLES_WITH_EFFECTS: &str =
        "http://schemas.microsoft.com/office/2007/relationships/stylesWithEffects";

    // ── Microsoft extensions (2011 — Office 2013) ─────────────────────
    pub const WEB_EXTENSION: &str =
        "http://schemas.microsoft.com/office/2011/relationships/webextension";
    pub const WEB_EXTENSION_TASK_PANES: &str =
        "http://schemas.microsoft.com/office/2011/relationships/webextensiontaskpanes";
    pub const PEOPLE: &str = "http://schemas.microsoft.com/office/2011/relationships/people";
    pub const CHART_STYLE: &str =
        "http://schemas.microsoft.com/office/2011/relationships/chartStyle";
    pub const CHART_COLOR_STYLE: &str =
        "http://schemas.microsoft.com/office/2011/relationships/chartColorStyle";
    pub const TIMELINE: &str = "http://schemas.microsoft.com/office/2011/relationships/timeline";
    pub const TIMELINE_CACHE: &str =
        "http://schemas.microsoft.com/office/2011/relationships/timelineCache";
    pub const COMMENTS_EXTENDED: &str =
        "http://schemas.microsoft.com/office/2011/relationships/commentsExtended";

    // ── Microsoft extensions (2014+) ──────────────────────────────────
    pub const CHART_EX: &str = "http://schemas.microsoft.com/office/2014/relationships/chartEx";
    pub const COMMENTS_IDS: &str =
        "http://schemas.microsoft.com/office/2016/09/relationships/commentsIds";
    pub const MODEL_3D: &str = "http://schemas.microsoft.com/office/2017/06/relationships/model3d";
    pub const RD_RICH_VALUE: &str =
        "http://schemas.microsoft.com/office/2017/06/relationships/rdRichValue";
    pub const RD_RICH_VALUE_STRUCTURE: &str =
        "http://schemas.microsoft.com/office/2017/06/relationships/rdRichValueStructure";
    pub const RD_RICH_VALUE_TYPES: &str =
        "http://schemas.microsoft.com/office/2017/06/relationships/rdRichValueTypes";
    pub const RD_ARRAY: &str = "http://schemas.microsoft.com/office/2017/06/relationships/rdArray";
    pub const RD_SUPPORTING_PROPERTY_BAG: &str =
        "http://schemas.microsoft.com/office/2017/06/relationships/rdSupportingPropertyBag";
    pub const RD_SUPPORTING_PROPERTY_BAG_STRUCTURE: &str =
        "http://schemas.microsoft.com/office/2017/06/relationships/rdSupportingPropertyBagStructure";
    pub const RICH_STYLES: &str =
        "http://schemas.microsoft.com/office/2017/06/relationships/richStyles";
    pub const PERSON: &str = "http://schemas.microsoft.com/office/2017/10/relationships/person";
    pub const THREADED_COMMENT: &str =
        "http://schemas.microsoft.com/office/2017/10/relationships/threadedComment";
    pub const COMMENTS_EXTENSIBLE: &str =
        "http://schemas.microsoft.com/office/2018/08/relationships/commentsExtensible";
    pub const COMMENTS_2018: &str =
        "http://schemas.microsoft.com/office/2018/10/relationships/comments";
    pub const AUTHORS_2018: &str =
        "http://schemas.microsoft.com/office/2018/10/relationships/authors";
    pub const NAMED_SHEET_VIEW: &str =
        "http://schemas.microsoft.com/office/2019/04/relationships/namedSheetView";
    pub const DOCUMENT_TASKS: &str =
        "http://schemas.microsoft.com/office/2019/05/relationships/documenttasks";
    pub const CLASSIFICATION_LABELS: &str =
        "http://schemas.microsoft.com/office/2020/02/relationships/classificationlabels";
    pub const RD_RICH_VALUE_WEB_IMAGE: &str =
        "http://schemas.microsoft.com/office/2020/07/relationships/rdRichValueWebImage";
    pub const FEATURE_PROPERTY_BAG: &str =
        "http://schemas.microsoft.com/office/2022/11/relationships/FeaturePropertyBag";
}

#[cfg(test)]
mod tests {
    use super::*;

    fn mk_rel(id: &str, rel_type: &str, target: &str) -> Relationship {
        Relationship {
            id: id.to_string(),
            rel_type: rel_type.to_string(),
            target: target.to_string(),
            target_mode: TargetMode::Internal,
        }
    }

    #[test]
    fn add_new_generates_ids_in_sequence() {
        let mut rels = Relationships::new();

        let first_id = rels
            .add_new(
                "type/a".to_string(),
                "a.xml".to_string(),
                TargetMode::Internal,
            )
            .id
            .clone();
        assert_eq!(first_id, "rId1");

        rels.add(mk_rel("rId7", "type/b", "b.xml"));

        let next_id = rels
            .add_new(
                "type/c".to_string(),
                "c.xml".to_string(),
                TargetMode::Internal,
            )
            .id
            .clone();
        assert_eq!(next_id, "rId8");

        let next_id_2 = rels
            .add_new(
                "type/d".to_string(),
                "d.xml".to_string(),
                TargetMode::Internal,
            )
            .id
            .clone();
        assert_eq!(next_id_2, "rId9");
    }

    #[test]
    fn lookup_by_id_works_for_existing_and_missing_ids() {
        let mut rels = Relationships::new();
        rels.add_new(
            "type/a".to_string(),
            "a.xml".to_string(),
            TargetMode::Internal,
        );

        let rel = rels.get_by_id("rId1").expect("relationship should exist");
        assert_eq!(rel.target, "a.xml");
        assert!(rels.contains_id("rId1"));
        assert!(!rels.contains_id("rId404"));
        assert!(rels.get_by_id("rId404").is_none());
    }

    #[test]
    fn xml_roundtrip_decodes_and_reescapes_relationship_target() {
        let input = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.com/?a=1&amp;b=2" TargetMode="External"/>
</Relationships>"#;

        let rels = Relationships::from_xml(std::io::Cursor::new(input.as_bytes()))
            .expect("relationships xml should parse");
        let rel = rels.get_by_id("rId1").expect("relationship should exist");
        assert_eq!(rel.target, "https://example.com/?a=1&b=2");
        assert_eq!(rel.target_mode, TargetMode::External);

        let mut output = Vec::new();
        rels.to_xml(&mut output)
            .expect("relationships xml should write");
        let output =
            String::from_utf8(output).expect("writer should produce UTF-8 relationship xml");

        assert!(output.contains(r#"Target="https://example.com/?a=1&amp;b=2""#));
        assert!(!output.contains("&amp;amp;"));

        let reparsed = Relationships::from_xml(std::io::Cursor::new(output.as_bytes()))
            .expect("roundtripped xml should parse");
        let reparsed_rel = reparsed
            .get_by_id("rId1")
            .expect("roundtripped relationship should exist");
        assert_eq!(reparsed_rel.target, "https://example.com/?a=1&b=2");
    }

    #[test]
    fn from_xml_bytes_preserves_original_bytes_on_noop_write() {
        let input = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://example.com/type" Target="a.xml"/>
</Relationships>"#;

        let rels =
            Relationships::from_xml_bytes(input.to_vec()).expect("relationships should parse");
        let mut output = Vec::new();
        rels.to_xml(&mut output)
            .expect("no-op relationship write should succeed");
        assert_eq!(
            output, input,
            "no-op relationship serialization should preserve exact original bytes"
        );
    }

    #[test]
    fn dirty_relationships_emit_regenerated_xml() {
        let input = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId2" Type="http://example.com/type" Target="a.xml"/>
</Relationships>"#;

        let mut rels =
            Relationships::from_xml_bytes(input.to_vec()).expect("relationships should parse");
        rels.add_new(
            "http://example.com/other".to_string(),
            "b.xml".to_string(),
            TargetMode::Internal,
        );

        let mut output = Vec::new();
        rels.to_xml(&mut output)
            .expect("dirty relationship write should succeed");
        assert_ne!(
            output, input,
            "dirty write should regenerate relationship xml"
        );
        let output_text =
            String::from_utf8(output).expect("serialized relationships should be UTF-8");
        assert!(output_text.contains("rId3"));
    }
}
