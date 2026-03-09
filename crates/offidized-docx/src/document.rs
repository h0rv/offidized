use std::collections::{BTreeMap, HashMap, HashSet};
use std::io::{BufRead, Cursor, Write};
use std::path::Path;

use offidized_opc::content_types::ContentTypeValue;
use offidized_opc::relationship::{RelationshipType, Relationships, TargetMode};
use offidized_opc::uri::PartUri;
use offidized_opc::{Package, Part, RawXmlNode};
use quick_xml::events::{BytesDecl, BytesEnd, BytesStart, BytesText, Event};
use quick_xml::{Reader, Writer};

use crate::bookmark::Bookmark;
use crate::comment::Comment;
use crate::content_control::ContentControl;
use crate::error::{DocxError, Result};
use crate::footnote::{Endnote, Footnote};
use crate::image::{FloatingImage, Image, InlineImage};
use crate::numbering::{
    NumberingDefinition, NumberingInstance, NumberingLevel, NumberingLevelOverride,
};
use crate::paragraph::{
    LineSpacingRule, Paragraph, ParagraphAlignment, ParagraphBorder, ParagraphBorders, TabStop,
    TabStopAlignment, TabStopLeader,
};
use crate::properties::DocumentProperties;
use crate::run::{FieldCode, Run, UnderlineType};
use crate::section::{
    HeaderFooter, LineNumberRestart, PageMargins, PageOrientation, Section, SectionBreakType,
    SectionVerticalAlignment,
};
use crate::style::{Style, StyleKind, StyleRegistry};
use crate::table::{
    CellBorders, CellMargins, Table, TableAlignment, TableBorder, TableBorders, TableLayout,
    TableRowProperties, TableWidthType, VerticalAlignment, VerticalMerge,
};

const WORD_DOCUMENT_URI: &str = "/word/document.xml";
const WORD_STYLES_URI: &str = "/word/styles.xml";
const WORD_MAIN_NS: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const WORD_REL_NS: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const DRAWINGML_NS: &str = "http://schemas.openxmlformats.org/drawingml/2006/main";
const WORDPROCESSING_DRAWING_NS: &str =
    "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
const DRAWINGML_PICTURE_NS: &str = "http://schemas.openxmlformats.org/drawingml/2006/picture";
const DRAWINGML_PICTURE_URI: &str = "http://schemas.openxmlformats.org/drawingml/2006/picture";
const OCTET_STREAM_CONTENT_TYPE: &str = "application/octet-stream";
const WORD_HEADER_CONTENT_TYPE: &str =
    "application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml";
const WORD_FOOTER_CONTENT_TYPE: &str =
    "application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml";
const WORD_HEADER_REL_TYPE: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/header";
const WORD_FOOTER_REL_TYPE: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer";
const WORD_FOOTNOTES_REL_TYPE: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes";
const WORD_ENDNOTES_REL_TYPE: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes";
const WORD_COMMENTS_REL_TYPE: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments";
const WORD_NUMBERING_REL_TYPE: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering";
const _WORD_SETTINGS_REL_TYPE: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings";
const _WORD_FOOTNOTES_CONTENT_TYPE: &str =
    "application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml";
const _WORD_ENDNOTES_CONTENT_TYPE: &str =
    "application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml";
const _WORD_COMMENTS_CONTENT_TYPE: &str =
    "application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml";
const _WORD_NUMBERING_CONTENT_TYPE: &str =
    "application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml";
const _WORD_SETTINGS_CONTENT_TYPE: &str =
    "application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml";
const CORE_PROPERTIES_URI: &str = "/docProps/core.xml";
const _CORE_PROPERTIES_CONTENT_TYPE: &str =
    "application/vnd.openxmlformats-package.core-properties+xml";
const _DC_NS: &str = "http://purl.org/dc/elements/1.1/";
const _CP_NS: &str = "http://schemas.openxmlformats.org/package/2006/metadata/core-properties";
const _DCTERMS_NS: &str = "http://purl.org/dc/terms/";

/// An ordered body item in a Word document.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum BodyItem<'a> {
    /// Paragraph item.
    Paragraph(&'a Paragraph),
    /// Table item.
    Table(&'a Table),
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum BodyItemRef {
    Paragraph(usize),
    Table(usize),
    Unknown(usize),
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum ParsedBodyItemKind {
    Paragraph,
    Table,
    Unknown,
}

#[derive(Debug, Clone, PartialEq, Eq, Default)]
struct SectionRelationshipIds {
    header_relationship_id: Option<String>,
    footer_relationship_id: Option<String>,
    first_page_header_relationship_id: Option<String>,
    first_page_footer_relationship_id: Option<String>,
    even_page_header_relationship_id: Option<String>,
    even_page_footer_relationship_id: Option<String>,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum DrawingKind {
    Inline,
    Anchor,
}

type PartRelationshipMaps = (
    Relationships,
    HashMap<String, String>,
    HashMap<usize, String>,
);

type DocumentRelationshipBuildResult = (
    Relationships,
    HashMap<String, String>,
    HashMap<usize, String>,
    SectionRelationshipIds,
);

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum ImageRelationshipInclusion {
    ReferencedInRuns,
    AllImages,
}

/// Iterator over document body items in XML order.
#[derive(Debug)]
pub struct BodyItems<'a> {
    document: &'a Document,
    cursor: usize,
}

impl<'a> Iterator for BodyItems<'a> {
    type Item = BodyItem<'a>;

    fn next(&mut self) -> Option<Self::Item> {
        while let Some(body_item_ref) = self.document.body.get(self.cursor).copied() {
            self.cursor = self.cursor.saturating_add(1);
            match body_item_ref {
                BodyItemRef::Paragraph(index) => {
                    if let Some(paragraph) = self.document.paragraphs.get(index) {
                        return Some(BodyItem::Paragraph(paragraph));
                    }
                }
                BodyItemRef::Table(index) => {
                    if let Some(table) = self.document.tables.get(index) {
                        return Some(BodyItem::Table(table));
                    }
                }
                BodyItemRef::Unknown(_) => {
                    // Unknown elements are preserved for roundtrip but not
                    // exposed through the high-level iterator.
                }
            }
        }

        None
    }
}

/// High-level Word document wrapper.
#[derive(Debug)]
pub struct Document {
    package: Package,
    paragraphs: Vec<Paragraph>,
    tables: Vec<Table>,
    images: Vec<Image>,
    section: Section,
    styles: StyleRegistry,
    body: Vec<BodyItemRef>,
    dirty: bool,
    /// Unknown XML children at the `<w:body>` level, preserved for roundtrip fidelity.
    unknown_body_children: Vec<RawXmlNode>,
    /// Footnotes parsed from `footnotes.xml`.
    footnotes: Vec<Footnote>,
    /// Endnotes parsed from `endnotes.xml`.
    endnotes: Vec<Endnote>,
    /// Bookmarks parsed from `w:bookmarkStart`/`w:bookmarkEnd` elements.
    bookmarks: Vec<Bookmark>,
    /// Comments parsed from `comments.xml`.
    comments: Vec<Comment>,
    /// Core document properties from `docProps/core.xml`.
    document_properties: DocumentProperties,
    /// Content controls (SDT) at the block level.
    content_controls: Vec<ContentControl>,
    /// Numbering definitions from `numbering.xml`.
    numbering_definitions: Vec<NumberingDefinition>,
    /// Numbering instances from `numbering.xml` (`w:num`).
    numbering_instances: Vec<NumberingInstance>,
    /// Document protection settings.
    protection: Option<DocumentProtection>,
    /// Extra namespace declarations from the original `<w:document>` element,
    /// preserved so that unknown elements/attributes using those prefixes remain
    /// valid XML on dirty save.
    extra_namespace_declarations: Vec<(String, String)>,
}

/// Document protection settings controlling editing restrictions.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct DocumentProtection {
    /// The editing restriction type (e.g. "readOnly", "comments", "trackedChanges", "forms").
    pub edit: String,
    /// Whether the protection is enforced.
    pub enforcement: bool,
}

impl DocumentProtection {
    /// Create a new document protection with the given editing restriction.
    ///
    /// Common values for `edit`: `"readOnly"`, `"comments"`, `"trackedChanges"`, `"forms"`.
    pub fn new(edit: impl Into<String>, enforcement: bool) -> Self {
        Self {
            edit: edit.into(),
            enforcement,
        }
    }

    /// Create a read-only protection.
    pub fn read_only() -> Self {
        Self::new("readOnly", true)
    }

    /// Create a comments-only protection.
    pub fn comments_only() -> Self {
        Self::new("comments", true)
    }

    /// Create a tracked-changes protection.
    pub fn tracked_changes() -> Self {
        Self::new("trackedChanges", true)
    }

    /// Create a forms-only protection.
    pub fn forms_only() -> Self {
        Self::new("forms", true)
    }
}

impl Document {
    /// Create a new in-memory document scaffold.
    pub fn new() -> Self {
        Self {
            package: Package::new(),
            paragraphs: Vec::new(),
            tables: Vec::new(),
            images: Vec::new(),
            section: Section::new(),
            styles: StyleRegistry::new(),
            body: Vec::new(),
            dirty: true,
            unknown_body_children: Vec::new(),
            footnotes: Vec::new(),
            endnotes: Vec::new(),
            bookmarks: Vec::new(),
            comments: Vec::new(),
            document_properties: DocumentProperties::new(),
            content_controls: Vec::new(),
            numbering_definitions: Vec::new(),
            numbering_instances: Vec::new(),
            protection: None,
            extra_namespace_declarations: Vec::new(),
        }
    }

    /// Open an existing `.docx` package from a file path.
    pub fn open(path: impl AsRef<Path>) -> Result<Self> {
        let package = Package::open(path)?;
        Self::from_package(package)
    }

    /// Open an existing `.docx` package from in-memory bytes.
    pub fn from_bytes(bytes: &[u8]) -> Result<Self> {
        let package = Package::from_bytes(bytes)?;
        Self::from_package(package)
    }

    /// Build a `Document` from an already-opened OPC package.
    fn from_package(package: Package) -> Result<Self> {
        let document_part_uri = resolve_word_document_part_uri(&package)?;
        let document_part_uri = PartUri::new(document_part_uri)?;
        let document_part = package
            .get_part(document_part_uri.as_str())
            .ok_or_else(|| {
                DocxError::UnsupportedPackage(format!(
                    "missing Word document part `{document_part_uri}`"
                ))
            })?;
        let hyperlink_targets = parse_hyperlink_targets_by_relationship_id(document_part);
        let (images, image_relationship_ids) =
            load_document_images(&package, &document_part_uri, document_part)?;
        let paragraphs = parse_paragraphs(
            document_part.data.as_bytes(),
            &hyperlink_targets,
            &image_relationship_ids,
        )?;
        let tables = parse_tables(document_part.data.as_bytes())?;
        let section = parse_section(
            &package,
            &document_part_uri,
            document_part,
            document_part.data.as_bytes(),
        )?;
        let styles = load_document_styles(&package, &document_part_uri, document_part)?;
        let (body_item_kinds, unknown_body_children) =
            parse_body_item_kinds(document_part.data.as_bytes())?;
        let body = bind_body_item_refs(
            &body_item_kinds,
            paragraphs.len(),
            tables.len(),
            unknown_body_children.len(),
        );
        let footnotes = load_footnotes(&package, &document_part_uri, document_part)?;
        let endnotes = load_endnotes(&package, &document_part_uri, document_part)?;
        let bookmarks = parse_bookmarks(document_part.data.as_bytes())?;
        let comments = load_comments(&package, &document_part_uri, document_part)?;
        let document_properties = load_document_properties(&package)?;
        let content_controls = parse_content_controls(document_part.data.as_bytes())?;
        let (numbering_definitions, numbering_instances) =
            load_numbering_definitions(&package, &document_part_uri, document_part)?;
        let extra_namespace_declarations = parse_root_element_namespace_declarations(
            document_part.data.as_bytes(),
            b"document",
            &["xmlns:w", "xmlns:r", "xmlns:wp", "xmlns:a", "xmlns:pic"],
        );

        Ok(Self {
            package,
            paragraphs,
            tables,
            images,
            section,
            styles,
            body,
            dirty: false,
            unknown_body_children,
            footnotes,
            endnotes,
            bookmarks,
            comments,
            document_properties,
            content_controls,
            numbering_definitions,
            numbering_instances,
            protection: None,
            extra_namespace_declarations,
        })
    }

    /// Save the current package to disk.
    pub fn save(&self, path: impl AsRef<Path>) -> Result<()> {
        let path = path.as_ref();
        if !self.dirty {
            self.package.save(path)?;
            return Ok(());
        }

        let package = self.build_save_package()?;
        package.save(path)?;
        Ok(())
    }

    /// Serialize the document to in-memory `.docx` bytes.
    ///
    /// Produces the same output as [`save`](Self::save) without touching the
    /// filesystem.
    pub fn to_bytes(&self) -> Result<Vec<u8>> {
        if !self.dirty {
            return Ok(self.package.to_bytes()?);
        }

        let package = self.build_save_package()?;
        Ok(package.to_bytes()?)
    }

    /// Build the fully-populated OPC package ready for serialization.
    ///
    /// Shared by [`save`](Self::save) and [`to_bytes`](Self::to_bytes).
    fn build_save_package(&self) -> Result<Package> {
        let mut package = self.package.clone();
        let mut document_passthrough_relationships = BTreeMap::<String, usize>::new();
        // Preserve existing media/header/footer topology for pass-through fidelity.
        // We only replace parts this serializer owns fully.
        let _ = package.remove_part("/word/document.xml");
        let _ = package.remove_part("/word/styles.xml");
        let document_part_uri = PartUri::new(WORD_DOCUMENT_URI)?;
        let (
            mut relationships,
            hyperlink_relationship_ids,
            image_relationship_ids,
            section_relationship_ids,
        ) = build_document_relationships_and_media_parts(
            &mut package,
            &document_part_uri,
            &self.paragraphs,
            &self.images,
            &self.section,
            &self.body,
        )?;
        if let Some(original_document_part) = self.package.get_part(document_part_uri.as_str()) {
            for relationship in original_document_part.relationships.iter() {
                if !is_rebuilt_document_relationship_type(relationship.rel_type.as_str()) {
                    record_passthrough_relationship(
                        &mut document_passthrough_relationships,
                        relationship.rel_type.as_str(),
                    );
                    relationships.add(relationship.clone());
                }
            }
        }
        let styles_to_write = build_effective_styles_registry(
            &self.styles,
            &self.paragraphs,
            &self.tables,
            &self.section,
        );
        if !styles_to_write.is_empty() {
            let styles_part_uri = PartUri::new(WORD_STYLES_URI)?;
            let styles_xml = serialize_styles_xml(&styles_to_write)?;
            let mut styles_part = Part::new_xml(styles_part_uri.clone(), styles_xml);
            styles_part.content_type = Some(ContentTypeValue::WORD_STYLES.to_string());
            package.set_part(styles_part);

            relationships.add_new(
                RelationshipType::STYLES.to_string(),
                relative_path_from_part(&document_part_uri, &styles_part_uri),
                TargetMode::Internal,
            );
        }
        let xml = serialize_document_xml(
            &self.paragraphs,
            &self.tables,
            &self.section,
            &self.body,
            &hyperlink_relationship_ids,
            &image_relationship_ids,
            &section_relationship_ids,
            &self.unknown_body_children,
            &self.extra_namespace_declarations,
        )?;
        let mut document_part = Part::new_xml(document_part_uri, xml);
        document_part.content_type = Some(ContentTypeValue::WORD_DOCUMENT.to_string());
        document_part.relationships = relationships;
        package.set_part(document_part);
        if package
            .relationships()
            .get_first_by_type(RelationshipType::WORD_DOCUMENT)
            .is_none()
        {
            package.relationships_mut().add_new(
                RelationshipType::WORD_DOCUMENT.to_string(),
                WORD_DOCUMENT_URI.trim_start_matches('/').to_string(),
                TargetMode::Internal,
            );
        }
        emit_passthrough_relationship_warnings("document", &document_passthrough_relationships);
        Ok(package)
    }

    /// Add a regular paragraph.
    pub fn add_paragraph(&mut self, text: impl Into<String>) -> &mut Paragraph {
        self.dirty = true;
        self.paragraphs.push(Paragraph::from_text(text));
        let idx = self.paragraphs.len().saturating_sub(1);
        self.body.push(BodyItemRef::Paragraph(idx));
        &mut self.paragraphs[idx]
    }

    /// Insert a paragraph at a specific body position.
    ///
    /// The position is an index into the `body` array (which interleaves
    /// paragraphs, tables, and unknown elements). The new paragraph is
    /// inserted before the element at `body_position`. If `body_position`
    /// is >= body.len(), appends at the end.
    pub fn insert_paragraph_at(&mut self, body_position: usize, paragraph: Paragraph) {
        self.dirty = true;
        let idx = self.paragraphs.len();
        self.paragraphs.push(paragraph);
        let pos = body_position.min(self.body.len());
        self.body.insert(pos, BodyItemRef::Paragraph(idx));
    }

    /// Find the body position of the paragraph with the given paragraph index.
    ///
    /// Returns the position in the `body` array where `Paragraph(paragraph_index)`
    /// is located, or `None` if no such entry exists. This is useful for
    /// determining insertion points relative to existing paragraphs.
    pub fn body_position_of_paragraph(&self, paragraph_index: usize) -> Option<usize> {
        self.body
            .iter()
            .position(|item| *item == BodyItemRef::Paragraph(paragraph_index))
    }

    /// Add a regular paragraph with an explicit paragraph style identifier.
    pub fn add_paragraph_with_style(
        &mut self,
        text: impl Into<String>,
        style_id: impl Into<String>,
    ) -> &mut Paragraph {
        self.dirty = true;
        let paragraph = self.add_paragraph(text);
        paragraph.set_style_id(style_id);
        paragraph
    }

    /// Add a heading paragraph.
    pub fn add_heading(&mut self, text: impl Into<String>, level: u8) -> &mut Paragraph {
        self.dirty = true;
        self.paragraphs.push(Paragraph::heading(text, level));
        let idx = self.paragraphs.len().saturating_sub(1);
        self.body.push(BodyItemRef::Paragraph(idx));
        &mut self.paragraphs[idx]
    }

    /// Add a bulleted paragraph.
    ///
    /// Ensures a bullet numbering definition exists and creates the paragraph with
    /// list numbering at level 0.
    pub fn add_bulleted_paragraph(&mut self, text: impl Into<String>) -> &mut Paragraph {
        let num_id = self.ensure_bullet_numbering_definition();
        let paragraph = self.add_paragraph(text);
        paragraph.set_numbering(num_id, 0);
        paragraph
    }

    /// Add a numbered list paragraph.
    ///
    /// Ensures a decimal numbering definition exists and creates the paragraph with
    /// list numbering at level 0.
    pub fn add_numbered_paragraph(&mut self, text: impl Into<String>) -> &mut Paragraph {
        let num_id = self.ensure_numbered_list_definition();
        let paragraph = self.add_paragraph(text);
        paragraph.set_numbering(num_id, 0);
        paragraph
    }

    /// Ensures a bullet numbering definition and instance exist.
    ///
    /// Returns the `num_id` of the bullet numbering instance. If one already exists,
    /// it is reused; otherwise a new definition and instance are created.
    pub fn ensure_bullet_numbering_definition(&mut self) -> u32 {
        // Look for an existing bullet definition
        for inst in &self.numbering_instances {
            if let Some(def) = self
                .numbering_definitions
                .iter()
                .find(|d| d.abstract_num_id() == inst.abstract_num_id())
            {
                if def
                    .level(0)
                    .map(|l| l.format() == "bullet")
                    .unwrap_or(false)
                {
                    return inst.num_id();
                }
            }
        }

        // Create a new one
        let abstract_id = self
            .numbering_definitions
            .iter()
            .map(|d| d.abstract_num_id())
            .max()
            .map(|m| m + 1)
            .unwrap_or(0);
        let num_id = self
            .numbering_instances
            .iter()
            .map(|i| i.num_id())
            .max()
            .map(|m| m + 1)
            .unwrap_or(1);

        self.numbering_definitions
            .push(NumberingDefinition::create_bullet(abstract_id));
        self.numbering_instances
            .push(NumberingInstance::new(num_id, abstract_id));
        self.dirty = true;
        num_id
    }

    /// Ensures a decimal numbered list definition and instance exist.
    ///
    /// Returns the `num_id` of the numbered list instance. If one already exists,
    /// it is reused; otherwise a new definition and instance are created.
    pub fn ensure_numbered_list_definition(&mut self) -> u32 {
        // Look for an existing decimal definition
        for inst in &self.numbering_instances {
            if let Some(def) = self
                .numbering_definitions
                .iter()
                .find(|d| d.abstract_num_id() == inst.abstract_num_id())
            {
                if def
                    .level(0)
                    .map(|l| l.format() == "decimal")
                    .unwrap_or(false)
                {
                    return inst.num_id();
                }
            }
        }

        // Create a new one
        let abstract_id = self
            .numbering_definitions
            .iter()
            .map(|d| d.abstract_num_id())
            .max()
            .map(|m| m + 1)
            .unwrap_or(0);
        let num_id = self
            .numbering_instances
            .iter()
            .map(|i| i.num_id())
            .max()
            .map(|m| m + 1)
            .unwrap_or(1);

        self.numbering_definitions
            .push(NumberingDefinition::create_numbered(abstract_id));
        self.numbering_instances
            .push(NumberingInstance::new(num_id, abstract_id));
        self.dirty = true;
        num_id
    }

    /// Read-only paragraph list.
    pub fn paragraphs(&self) -> &[Paragraph] {
        &self.paragraphs
    }

    /// Mutable paragraph list.
    pub fn paragraphs_mut(&mut self) -> &mut [Paragraph] {
        self.dirty = true;
        &mut self.paragraphs
    }

    /// Add a binary image payload and return its index for inline-image runs.
    pub fn add_image(
        &mut self,
        bytes: impl Into<Vec<u8>>,
        content_type: impl Into<String>,
    ) -> usize {
        self.dirty = true;
        self.images.push(Image::new(bytes, content_type));
        self.images.len().saturating_sub(1)
    }

    /// Read-only image list.
    pub fn images(&self) -> &[Image] {
        &self.images
    }

    /// Mutable image list.
    pub fn images_mut(&mut self) -> &mut [Image] {
        self.dirty = true;
        &mut self.images
    }

    /// Add a table scaffold.
    pub fn add_table(&mut self, rows: usize, columns: usize) -> &mut Table {
        self.dirty = true;
        self.tables.push(Table::new(rows, columns));
        let idx = self.tables.len().saturating_sub(1);
        self.body.push(BodyItemRef::Table(idx));
        &mut self.tables[idx]
    }

    /// Insert a table at a specific body position.
    ///
    /// The position is an index into the `body` array (which interleaves
    /// paragraphs, tables, and unknown elements). The new table is inserted
    /// before the element at `body_position`. If `body_position` is >=
    /// `body.len()`, appends at the end.
    pub fn insert_table_at(&mut self, body_position: usize, table: Table) {
        self.dirty = true;
        let idx = self.tables.len();
        self.tables.push(table);
        let pos = body_position.min(self.body.len());
        self.body.insert(pos, BodyItemRef::Table(idx));
    }

    /// Add a table scaffold with an explicit table style identifier.
    pub fn add_table_with_style(
        &mut self,
        rows: usize,
        columns: usize,
        style_id: impl Into<String>,
    ) -> &mut Table {
        self.dirty = true;
        let table = self.add_table(rows, columns);
        table.set_style_id(style_id);
        table
    }

    /// Read-only table list.
    pub fn tables(&self) -> &[Table] {
        &self.tables
    }

    /// Find the body position of the table with the given table index.
    ///
    /// Returns the position in the `body` array where `Table(table_index)` is
    /// located, or `None` if no such entry exists.
    pub fn body_position_of_table(&self, table_index: usize) -> Option<usize> {
        self.body
            .iter()
            .position(|item| *item == BodyItemRef::Table(table_index))
    }

    /// Mutable table list.
    pub fn tables_mut(&mut self) -> &mut [Table] {
        self.dirty = true;
        &mut self.tables
    }

    /// Document section layout settings (`w:sectPr`).
    pub fn section(&self) -> &Section {
        &self.section
    }

    /// Mutable section layout settings (`w:sectPr`).
    pub fn section_mut(&mut self) -> &mut Section {
        self.dirty = true;
        &mut self.section
    }

    /// Style registry (`/word/styles.xml`).
    pub fn styles(&self) -> &StyleRegistry {
        &self.styles
    }

    /// Mutable style registry (`/word/styles.xml`).
    pub fn styles_mut(&mut self) -> &mut StyleRegistry {
        self.dirty = true;
        &mut self.styles
    }

    /// Lookup a paragraph style by style identifier.
    pub fn paragraph_style(&self, style_id: &str) -> Option<&Style> {
        self.styles.paragraph_style(style_id)
    }

    /// Lookup a character style by style identifier.
    pub fn character_style(&self, style_id: &str) -> Option<&Style> {
        self.styles.character_style(style_id)
    }

    /// Lookup a table style by style identifier.
    pub fn table_style(&self, style_id: &str) -> Option<&Style> {
        self.styles.table_style(style_id)
    }

    /// Ordered body view preserving paragraph/table interleaving.
    pub fn body_items(&self) -> BodyItems<'_> {
        BodyItems {
            document: self,
            cursor: 0,
        }
    }

    /// Access the backing OPC package.
    pub fn package(&self) -> &Package {
        &self.package
    }

    /// Footnotes parsed from `footnotes.xml`.
    pub fn footnotes(&self) -> &[Footnote] {
        &self.footnotes
    }

    /// Mutable footnotes.
    pub fn footnotes_mut(&mut self) -> &mut Vec<Footnote> {
        self.dirty = true;
        &mut self.footnotes
    }

    /// Add a footnote.
    pub fn add_footnote(&mut self, footnote: Footnote) -> &mut Footnote {
        self.dirty = true;
        self.footnotes.push(footnote);
        let index = self.footnotes.len().saturating_sub(1);
        &mut self.footnotes[index]
    }

    /// Endnotes parsed from `endnotes.xml`.
    pub fn endnotes(&self) -> &[Endnote] {
        &self.endnotes
    }

    /// Mutable endnotes.
    pub fn endnotes_mut(&mut self) -> &mut Vec<Endnote> {
        self.dirty = true;
        &mut self.endnotes
    }

    /// Add an endnote.
    pub fn add_endnote(&mut self, endnote: Endnote) -> &mut Endnote {
        self.dirty = true;
        self.endnotes.push(endnote);
        let index = self.endnotes.len().saturating_sub(1);
        &mut self.endnotes[index]
    }

    /// Bookmarks in the document.
    pub fn bookmarks(&self) -> &[Bookmark] {
        &self.bookmarks
    }

    /// Mutable bookmarks.
    pub fn bookmarks_mut(&mut self) -> &mut Vec<Bookmark> {
        self.dirty = true;
        &mut self.bookmarks
    }

    /// Add a bookmark.
    pub fn add_bookmark(&mut self, bookmark: Bookmark) -> &mut Bookmark {
        self.dirty = true;
        self.bookmarks.push(bookmark);
        let index = self.bookmarks.len().saturating_sub(1);
        &mut self.bookmarks[index]
    }

    /// Comments parsed from `comments.xml`.
    pub fn comments(&self) -> &[Comment] {
        &self.comments
    }

    /// Mutable comments.
    pub fn comments_mut(&mut self) -> &mut Vec<Comment> {
        self.dirty = true;
        &mut self.comments
    }

    /// Add a comment.
    pub fn add_comment(&mut self, comment: Comment) -> &mut Comment {
        self.dirty = true;
        self.comments.push(comment);
        let index = self.comments.len().saturating_sub(1);
        &mut self.comments[index]
    }

    /// Core document properties from `docProps/core.xml`.
    pub fn document_properties(&self) -> &DocumentProperties {
        &self.document_properties
    }

    /// Mutable document properties.
    pub fn document_properties_mut(&mut self) -> &mut DocumentProperties {
        self.dirty = true;
        &mut self.document_properties
    }

    /// Replace document properties.
    pub fn set_document_properties(&mut self, properties: DocumentProperties) {
        self.dirty = true;
        self.document_properties = properties;
    }

    /// Content controls (SDT) at the block level.
    pub fn content_controls(&self) -> &[ContentControl] {
        &self.content_controls
    }

    /// Mutable content controls.
    pub fn content_controls_mut(&mut self) -> &mut Vec<ContentControl> {
        self.dirty = true;
        &mut self.content_controls
    }

    /// Add a content control.
    pub fn add_content_control(&mut self, sdt: ContentControl) -> &mut ContentControl {
        self.dirty = true;
        self.content_controls.push(sdt);
        let index = self.content_controls.len().saturating_sub(1);
        &mut self.content_controls[index]
    }

    /// Numbering definitions from `numbering.xml`.
    pub fn numbering_definitions(&self) -> &[NumberingDefinition] {
        &self.numbering_definitions
    }

    /// Mutable numbering definitions.
    pub fn numbering_definitions_mut(&mut self) -> &mut Vec<NumberingDefinition> {
        self.dirty = true;
        &mut self.numbering_definitions
    }

    /// Add a numbering definition.
    pub fn add_numbering_definition(
        &mut self,
        definition: NumberingDefinition,
    ) -> &mut NumberingDefinition {
        self.dirty = true;
        self.numbering_definitions.push(definition);
        let index = self.numbering_definitions.len().saturating_sub(1);
        &mut self.numbering_definitions[index]
    }

    /// Numbering instances from `numbering.xml`.
    pub fn numbering_instances(&self) -> &[NumberingInstance] {
        &self.numbering_instances
    }

    /// Mutable numbering instances.
    pub fn numbering_instances_mut(&mut self) -> &mut Vec<NumberingInstance> {
        self.dirty = true;
        &mut self.numbering_instances
    }

    /// Returns the document protection settings, if set.
    pub fn protection(&self) -> Option<&DocumentProtection> {
        self.protection.as_ref()
    }

    /// Set document protection.
    pub fn set_protection(&mut self, protection: DocumentProtection) -> &mut Self {
        self.protection = Some(protection);
        self.dirty = true;
        self
    }

    /// Remove document protection.
    pub fn clear_protection(&mut self) -> &mut Self {
        self.protection = None;
        self.dirty = true;
        self
    }
}

impl Default for Document {
    fn default() -> Self {
        Self::new()
    }
}

#[allow(clippy::too_many_arguments)]
fn serialize_document_xml(
    paragraphs: &[Paragraph],
    tables: &[Table],
    section: &Section,
    body: &[BodyItemRef],
    hyperlink_relationship_ids: &HashMap<String, String>,
    image_relationship_ids: &HashMap<usize, String>,
    section_relationship_ids: &SectionRelationshipIds,
    unknown_body_children: &[RawXmlNode],
    extra_namespace_declarations: &[(String, String)],
) -> Result<Vec<u8>> {
    let mut writer = Writer::new_with_indent(Vec::new(), b' ', 2);
    writer.write_event(Event::Decl(BytesDecl::new(
        "1.0",
        Some("UTF-8"),
        Some("yes"),
    )))?;

    let mut document = BytesStart::new("w:document");
    document.push_attribute(("xmlns:w", WORD_MAIN_NS));
    if !hyperlink_relationship_ids.is_empty()
        || !image_relationship_ids.is_empty()
        || section_relationship_ids.header_relationship_id.is_some()
        || section_relationship_ids.footer_relationship_id.is_some()
        || section_relationship_ids
            .first_page_header_relationship_id
            .is_some()
        || section_relationship_ids
            .first_page_footer_relationship_id
            .is_some()
        || section_relationship_ids
            .even_page_header_relationship_id
            .is_some()
        || section_relationship_ids
            .even_page_footer_relationship_id
            .is_some()
    {
        document.push_attribute(("xmlns:r", WORD_REL_NS));
    }
    if !image_relationship_ids.is_empty() {
        document.push_attribute(("xmlns:wp", WORDPROCESSING_DRAWING_NS));
        document.push_attribute(("xmlns:a", DRAWINGML_NS));
        document.push_attribute(("xmlns:pic", DRAWINGML_PICTURE_NS));
    }
    // Replay extra namespace declarations from the original XML so that unknown
    // elements/attributes using those prefixes remain valid on dirty save.
    for (prefix, uri) in extra_namespace_declarations {
        document.push_attribute((prefix.as_str(), uri.as_str()));
    }
    writer.write_event(Event::Start(document))?;
    writer.write_event(Event::Start(BytesStart::new("w:body")))?;

    let mut next_drawing_id = 1_u32;

    if body.is_empty() {
        for paragraph in paragraphs {
            write_paragraph_xml(
                &mut writer,
                paragraph,
                hyperlink_relationship_ids,
                image_relationship_ids,
                &mut next_drawing_id,
            )?;
        }
        for table in tables {
            write_table_xml(&mut writer, table)?;
        }
    } else {
        for body_item in body {
            match body_item {
                BodyItemRef::Paragraph(index) => {
                    if let Some(paragraph) = paragraphs.get(*index) {
                        write_paragraph_xml(
                            &mut writer,
                            paragraph,
                            hyperlink_relationship_ids,
                            image_relationship_ids,
                            &mut next_drawing_id,
                        )?;
                    }
                }
                BodyItemRef::Table(index) => {
                    if let Some(table) = tables.get(*index) {
                        write_table_xml(&mut writer, table)?;
                    }
                }
                BodyItemRef::Unknown(index) => {
                    if let Some(node) = unknown_body_children.get(*index) {
                        node.write_to(&mut writer)?;
                    }
                }
            }
        }
    }

    if section.has_properties() {
        write_section_xml(&mut writer, section, section_relationship_ids)?;
    }

    writer.write_event(Event::End(BytesEnd::new("w:body")))?;
    writer.write_event(Event::End(BytesEnd::new("w:document")))?;

    Ok(writer.into_inner())
}

fn write_paragraph_xml<W: Write>(
    writer: &mut Writer<W>,
    paragraph: &Paragraph,
    hyperlink_relationship_ids: &HashMap<String, String>,
    image_relationship_ids: &HashMap<usize, String>,
    next_drawing_id: &mut u32,
) -> Result<()> {
    writer.write_event(Event::Start(BytesStart::new("w:p")))?;

    if paragraph.has_properties() {
        writer.write_event(Event::Start(BytesStart::new("w:pPr")))?;

        if let Some(style_id) = paragraph.style_id() {
            let mut style = BytesStart::new("w:pStyle");
            style.push_attribute(("w:val", style_id));
            writer.write_event(Event::Empty(style))?;
        } else if let Some(level) = paragraph.heading_level() {
            let mut style = BytesStart::new("w:pStyle");
            let style_name = format!("Heading{level}");
            style.push_attribute(("w:val", style_name.as_str()));
            writer.write_event(Event::Empty(style))?;
        }
        if paragraph.numbering_num_id().is_some() || paragraph.numbering_ilvl().is_some() {
            writer.write_event(Event::Start(BytesStart::new("w:numPr")))?;

            if let Some(ilvl) = paragraph.numbering_ilvl() {
                let mut ilvl_node = BytesStart::new("w:ilvl");
                let ilvl_value = ilvl.to_string();
                ilvl_node.push_attribute(("w:val", ilvl_value.as_str()));
                writer.write_event(Event::Empty(ilvl_node))?;
            }
            if let Some(num_id) = paragraph.numbering_num_id() {
                let mut num_id_node = BytesStart::new("w:numId");
                let num_id_value = num_id.to_string();
                num_id_node.push_attribute(("w:val", num_id_value.as_str()));
                writer.write_event(Event::Empty(num_id_node))?;
            }

            writer.write_event(Event::End(BytesEnd::new("w:numPr")))?;
        }
        if let Some(alignment) = paragraph.alignment() {
            let mut jc = BytesStart::new("w:jc");
            jc.push_attribute(("w:val", alignment.to_xml_value()));
            writer.write_event(Event::Empty(jc))?;
        }
        if paragraph.spacing_before_twips().is_some()
            || paragraph.spacing_after_twips().is_some()
            || paragraph.line_spacing_twips().is_some()
            || paragraph.line_spacing_rule().is_some()
            || paragraph.before_autospacing().is_some()
            || paragraph.after_autospacing().is_some()
        {
            let mut spacing = BytesStart::new("w:spacing");
            let before_value = paragraph
                .spacing_before_twips()
                .map(|value| value.to_string());
            let after_value = paragraph
                .spacing_after_twips()
                .map(|value| value.to_string());
            let line_value = paragraph
                .line_spacing_twips()
                .map(|value| value.to_string());

            if let Some(value) = before_value.as_deref() {
                spacing.push_attribute(("w:before", value));
            }
            if let Some(auto) = paragraph.before_autospacing() {
                spacing.push_attribute(("w:beforeAutospacing", if auto { "1" } else { "0" }));
            }
            if let Some(value) = after_value.as_deref() {
                spacing.push_attribute(("w:after", value));
            }
            if let Some(auto) = paragraph.after_autospacing() {
                spacing.push_attribute(("w:afterAutospacing", if auto { "1" } else { "0" }));
            }
            if let Some(value) = line_value.as_deref() {
                spacing.push_attribute(("w:line", value));
            }
            if let Some(rule) = paragraph.line_spacing_rule() {
                spacing.push_attribute(("w:lineRule", rule.to_xml_value()));
            }
            writer.write_event(Event::Empty(spacing))?;
        }
        if paragraph.indent_left_twips().is_some()
            || paragraph.indent_right_twips().is_some()
            || paragraph.indent_first_line_twips().is_some()
            || paragraph.indent_hanging_twips().is_some()
        {
            let mut indentation = BytesStart::new("w:ind");
            let left_value = paragraph.indent_left_twips().map(|value| value.to_string());
            let right_value = paragraph
                .indent_right_twips()
                .map(|value| value.to_string());
            let first_line_value = paragraph
                .indent_first_line_twips()
                .map(|value| value.to_string());
            let hanging_value = paragraph
                .indent_hanging_twips()
                .map(|value| value.to_string());

            if let Some(value) = left_value.as_deref() {
                indentation.push_attribute(("w:left", value));
            }
            if let Some(value) = right_value.as_deref() {
                indentation.push_attribute(("w:right", value));
            }
            if let Some(value) = first_line_value.as_deref() {
                indentation.push_attribute(("w:firstLine", value));
            }
            if let Some(value) = hanging_value.as_deref() {
                indentation.push_attribute(("w:hanging", value));
            }
            writer.write_event(Event::Empty(indentation))?;
        }
        if paragraph.keep_next() {
            writer.write_event(Event::Empty(BytesStart::new("w:keepNext")))?;
        }
        if paragraph.keep_lines() {
            writer.write_event(Event::Empty(BytesStart::new("w:keepLines")))?;
        }
        if let Some(widow_control) = paragraph.widow_control() {
            if widow_control {
                writer.write_event(Event::Empty(BytesStart::new("w:widowControl")))?;
            } else {
                let mut wc = BytesStart::new("w:widowControl");
                wc.push_attribute(("w:val", "0"));
                writer.write_event(Event::Empty(wc))?;
            }
        }
        if !paragraph.tab_stops().is_empty() {
            write_tab_stops_xml(writer, paragraph.tab_stops())?;
        }
        if !paragraph.borders().is_empty() {
            write_paragraph_borders_xml(writer, paragraph.borders())?;
        }
        if paragraph.shading_color().is_some() {
            let mut shd = BytesStart::new("w:shd");
            if let Some(pattern) = paragraph.shading_pattern() {
                shd.push_attribute(("w:val", pattern));
            } else {
                shd.push_attribute(("w:val", "clear"));
            }
            if let Some(color_attr) = paragraph.shading_color_attribute() {
                shd.push_attribute(("w:color", color_attr));
            } else {
                shd.push_attribute(("w:color", "auto"));
            }
            if let Some(fill) = paragraph.shading_color() {
                shd.push_attribute(("w:fill", fill));
            }
            writer.write_event(Event::Empty(shd))?;
        }
        if paragraph.is_bidi() {
            writer.write_event(Event::Empty(BytesStart::new("w:bidi")))?;
        }
        if paragraph.page_break_before() {
            writer.write_event(Event::Empty(BytesStart::new("w:pageBreakBefore")))?;
        }
        if paragraph.contextual_spacing() {
            writer.write_event(Event::Empty(BytesStart::new("w:contextualSpacing")))?;
        }
        if let Some(level) = paragraph.outline_level() {
            let mut elem = BytesStart::new("w:outlineLvl");
            let val = level.to_string();
            elem.push_attribute(("w:val", val.as_str()));
            writer.write_event(Event::Empty(elem))?;
        }
        if let Some(section) = paragraph.section_properties() {
            let empty_section_rel_ids = SectionRelationshipIds::default();
            write_section_xml(writer, section, &empty_section_rel_ids)?;
        }

        for node in paragraph.unknown_property_children() {
            node.write_to(writer)?;
        }

        writer.write_event(Event::End(BytesEnd::new("w:pPr")))?;
    }

    // Write comment range starts before runs
    for comment_id in paragraph.comment_range_start_ids() {
        let mut elem = BytesStart::new("w:commentRangeStart");
        let id_value = comment_id.to_string();
        elem.push_attribute(("w:id", id_value.as_str()));
        writer.write_event(Event::Empty(elem))?;
    }

    let runs = paragraph.runs();
    let mut index = 0_usize;
    while let Some(run) = runs.get(index) {
        if let Some(hyperlink_target) = run.hyperlink() {
            if let Some(relationship_id) = hyperlink_relationship_ids.get(hyperlink_target) {
                let mut hyperlink = BytesStart::new("w:hyperlink");
                hyperlink.push_attribute(("r:id", relationship_id.as_str()));
                if let Some(tooltip) = run.hyperlink_tooltip() {
                    hyperlink.push_attribute(("w:tooltip", tooltip));
                }
                if let Some(anchor) = run.hyperlink_anchor() {
                    hyperlink.push_attribute(("w:anchor", anchor));
                }
                writer.write_event(Event::Start(hyperlink))?;

                while let Some(grouped_run) = runs.get(index) {
                    if grouped_run.hyperlink() != Some(hyperlink_target) {
                        break;
                    }
                    write_run_xml(writer, grouped_run, image_relationship_ids, next_drawing_id)?;
                    index = index.saturating_add(1);
                }

                writer.write_event(Event::End(BytesEnd::new("w:hyperlink")))?;
                continue;
            }
        }

        write_run_xml(writer, run, image_relationship_ids, next_drawing_id)?;
        index = index.saturating_add(1);
    }

    // Write comment range ends after runs
    for comment_id in paragraph.comment_range_end_ids() {
        let mut elem = BytesStart::new("w:commentRangeEnd");
        let id_value = comment_id.to_string();
        elem.push_attribute(("w:id", id_value.as_str()));
        writer.write_event(Event::Empty(elem))?;
    }

    for node in paragraph.unknown_children() {
        node.write_to(writer)?;
    }

    writer.write_event(Event::End(BytesEnd::new("w:p")))?;

    Ok(())
}

fn write_run_xml<W: Write>(
    writer: &mut Writer<W>,
    run: &Run,
    image_relationship_ids: &HashMap<usize, String>,
    next_drawing_id: &mut u32,
) -> Result<()> {
    writer.write_event(Event::Start(BytesStart::new("w:r")))?;

    if run.has_properties() {
        writer.write_event(Event::Start(BytesStart::new("w:rPr")))?;

        if let Some(style_id) = run.style_id() {
            let mut style = BytesStart::new("w:rStyle");
            style.push_attribute(("w:val", style_id));
            writer.write_event(Event::Empty(style))?;
        }
        if run.is_bold() {
            writer.write_event(Event::Empty(BytesStart::new("w:b")))?;
        }
        if run.is_italic() {
            writer.write_event(Event::Empty(BytesStart::new("w:i")))?;
        }
        if let Some(ut) = run.underline_type() {
            let mut underline = BytesStart::new("w:u");
            let val = match ut {
                UnderlineType::Single => "single",
                UnderlineType::Double => "double",
                UnderlineType::Thick => "thick",
                UnderlineType::Dotted => "dotted",
                UnderlineType::DottedHeavy => "dottedHeavy",
                UnderlineType::Dash => "dash",
                UnderlineType::DashedHeavy => "dashedHeavy",
                UnderlineType::DashLong => "dashLong",
                UnderlineType::DashLongHeavy => "dashLongHeavy",
                UnderlineType::DashDot => "dotDash",
                UnderlineType::DashDotHeavy => "dashDotHeavy",
                UnderlineType::DashDotDot => "dotDotDash",
                UnderlineType::DashDotDotHeavy => "dashDotDotHeavy",
                UnderlineType::Wavy => "wave",
                UnderlineType::WavyHeavy => "wavyHeavy",
                UnderlineType::WavyDouble => "wavyDouble",
                UnderlineType::Words => "words",
            };
            underline.push_attribute(("w:val", val));
            writer.write_event(Event::Empty(underline))?;
        }
        if run.is_strikethrough() {
            writer.write_event(Event::Empty(BytesStart::new("w:strike")))?;
        }
        if run.is_double_strikethrough() {
            writer.write_event(Event::Empty(BytesStart::new("w:dstrike")))?;
        }
        if run.is_subscript() {
            let mut vert_align = BytesStart::new("w:vertAlign");
            vert_align.push_attribute(("w:val", "subscript"));
            writer.write_event(Event::Empty(vert_align))?;
        } else if run.is_superscript() {
            let mut vert_align = BytesStart::new("w:vertAlign");
            vert_align.push_attribute(("w:val", "superscript"));
            writer.write_event(Event::Empty(vert_align))?;
        }
        if run.is_small_caps() {
            writer.write_event(Event::Empty(BytesStart::new("w:smallCaps")))?;
        }
        if run.is_all_caps() {
            writer.write_event(Event::Empty(BytesStart::new("w:caps")))?;
        }
        if run.is_hidden() {
            writer.write_event(Event::Empty(BytesStart::new("w:vanish")))?;
        }
        if run.is_emboss() {
            writer.write_event(Event::Empty(BytesStart::new("w:emboss")))?;
        }
        if run.is_imprint() {
            writer.write_event(Event::Empty(BytesStart::new("w:imprint")))?;
        }
        if run.is_shadow() {
            writer.write_event(Event::Empty(BytesStart::new("w:shadow")))?;
        }
        if run.is_outline() {
            writer.write_event(Event::Empty(BytesStart::new("w:outline")))?;
        }
        if let Some(spacing) = run.character_spacing_twips() {
            let mut elem = BytesStart::new("w:spacing");
            let val = spacing.to_string();
            elem.push_attribute(("w:val", val.as_str()));
            writer.write_event(Event::Empty(elem))?;
        }
        if let Some(highlight_color) = run.highlight_color() {
            let mut highlight = BytesStart::new("w:highlight");
            highlight.push_attribute(("w:val", highlight_color));
            writer.write_event(Event::Empty(highlight))?;
        }
        {
            let has_font = run.font_family_ascii().is_some()
                || run.font_family_h_ansi().is_some()
                || run.font_family_cs().is_some()
                || run.font_family_east_asia().is_some()
                || run.font_family().is_some();
            if has_font {
                let mut fonts = BytesStart::new("w:rFonts");
                let fallback = run.font_family();
                if let Some(v) = run.font_family_ascii().or(fallback) {
                    fonts.push_attribute(("w:ascii", v));
                }
                if let Some(v) = run.font_family_h_ansi().or(fallback) {
                    fonts.push_attribute(("w:hAnsi", v));
                }
                if let Some(v) = run.font_family_cs().or(fallback) {
                    fonts.push_attribute(("w:cs", v));
                }
                if let Some(v) = run.font_family_east_asia() {
                    fonts.push_attribute(("w:eastAsia", v));
                }
                writer.write_event(Event::Empty(fonts))?;
            }
        }
        if let Some(font_size_half_points) = run.font_size_half_points() {
            let mut size = BytesStart::new("w:sz");
            let size_value = font_size_half_points.to_string();
            size.push_attribute(("w:val", size_value.as_str()));
            writer.write_event(Event::Empty(size))?;
        }
        if run.color().is_some() || run.theme_color().is_some() {
            let mut color_node = BytesStart::new("w:color");
            if let Some(color) = run.color() {
                color_node.push_attribute(("w:val", color));
            }
            if let Some(theme_color) = run.theme_color() {
                color_node.push_attribute(("w:themeColor", theme_color));
            }
            if let Some(theme_shade) = run.theme_shade() {
                color_node.push_attribute(("w:themeShade", theme_shade));
            }
            if let Some(theme_tint) = run.theme_tint() {
                color_node.push_attribute(("w:themeTint", theme_tint));
            }
            writer.write_event(Event::Empty(color_node))?;
        }
        if run.is_rtl() {
            writer.write_event(Event::Empty(BytesStart::new("w:rtl")))?;
        }

        for node in run.unknown_property_children() {
            node.write_to(writer)?;
        }

        writer.write_event(Event::End(BytesEnd::new("w:rPr")))?;
    }

    // Write complex field code if present
    if let Some(field_code) = run.field_code() {
        // begin
        let mut fld_char_begin = BytesStart::new("w:fldChar");
        fld_char_begin.push_attribute(("w:fldCharType", "begin"));
        writer.write_event(Event::Empty(fld_char_begin))?;

        // instruction text
        writer.write_event(Event::Start(BytesStart::new("w:instrText")))?;
        writer.write_event(Event::Text(BytesText::new(field_code.instruction())))?;
        writer.write_event(Event::End(BytesEnd::new("w:instrText")))?;

        // separate
        let mut fld_char_separate = BytesStart::new("w:fldChar");
        fld_char_separate.push_attribute(("w:fldCharType", "separate"));
        writer.write_event(Event::Empty(fld_char_separate))?;

        // result text
        if !field_code.result().is_empty() {
            writer.write_event(Event::Start(BytesStart::new("w:t")))?;
            writer.write_event(Event::Text(BytesText::new(field_code.result())))?;
            writer.write_event(Event::End(BytesEnd::new("w:t")))?;
        }

        // end
        let mut fld_char_end = BytesStart::new("w:fldChar");
        fld_char_end.push_attribute(("w:fldCharType", "end"));
        writer.write_event(Event::Empty(fld_char_end))?;
    }

    let has_drawing = run.inline_image().is_some() || run.floating_image().is_some();
    let has_special_content = run.footnote_reference_id().is_some()
        || run.endnote_reference_id().is_some()
        || run.has_tab()
        || run.has_break();
    if !run.text().is_empty() || (!has_drawing && !has_special_content) {
        let mut text = BytesStart::new("w:t");
        if run.text().starts_with(' ') || run.text().ends_with(' ') {
            text.push_attribute(("xml:space", "preserve"));
        }
        writer.write_event(Event::Start(text))?;
        writer.write_event(Event::Text(BytesText::new(run.text())))?;
        writer.write_event(Event::End(BytesEnd::new("w:t")))?;
    }
    if run.has_tab() {
        writer.write_event(Event::Empty(BytesStart::new("w:tab")))?;
    }
    if run.has_break() {
        writer.write_event(Event::Empty(BytesStart::new("w:br")))?;
    }
    if let Some(footnote_id) = run.footnote_reference_id() {
        let mut footnote_ref = BytesStart::new("w:footnoteReference");
        let id_value = footnote_id.to_string();
        footnote_ref.push_attribute(("w:id", id_value.as_str()));
        writer.write_event(Event::Empty(footnote_ref))?;
    }
    if let Some(endnote_id) = run.endnote_reference_id() {
        let mut endnote_ref = BytesStart::new("w:endnoteReference");
        let id_value = endnote_id.to_string();
        endnote_ref.push_attribute(("w:id", id_value.as_str()));
        writer.write_event(Event::Empty(endnote_ref))?;
    }
    if let Some(inline_image) = run.inline_image() {
        write_inline_drawing_xml(
            writer,
            inline_image,
            image_relationship_ids,
            next_drawing_id,
        )?;
    } else if let Some(floating_image) = run.floating_image() {
        write_floating_drawing_xml(
            writer,
            floating_image,
            image_relationship_ids,
            next_drawing_id,
        )?;
    }
    for node in run.unknown_children() {
        node.write_to(writer)?;
    }
    writer.write_event(Event::End(BytesEnd::new("w:r")))?;

    Ok(())
}

fn write_inline_drawing_xml<W: Write>(
    writer: &mut Writer<W>,
    inline_image: &InlineImage,
    image_relationship_ids: &HashMap<usize, String>,
    next_drawing_id: &mut u32,
) -> Result<()> {
    let Some(relationship_id) = image_relationship_ids.get(&inline_image.image_index()) else {
        return Ok(());
    };

    let drawing_id = *next_drawing_id;
    *next_drawing_id = next_drawing_id.checked_add(1).ok_or_else(|| {
        DocxError::UnsupportedPackage("drawing id overflow while serializing".to_string())
    })?;

    let width = inline_image.width_emu();
    let height = inline_image.height_emu();

    writer.write_event(Event::Start(BytesStart::new("w:drawing")))?;

    let mut inline = BytesStart::new("wp:inline");
    inline.push_attribute(("distT", "0"));
    inline.push_attribute(("distB", "0"));
    inline.push_attribute(("distL", "0"));
    inline.push_attribute(("distR", "0"));
    writer.write_event(Event::Start(inline))?;

    let mut extent = BytesStart::new("wp:extent");
    let width_text = width.to_string();
    let height_text = height.to_string();
    extent.push_attribute(("cx", width_text.as_str()));
    extent.push_attribute(("cy", height_text.as_str()));
    writer.write_event(Event::Empty(extent))?;

    let mut doc_pr = BytesStart::new("wp:docPr");
    let drawing_id_text = drawing_id.to_string();
    let default_name = format!("Picture {drawing_id}");
    doc_pr.push_attribute(("id", drawing_id_text.as_str()));
    doc_pr.push_attribute(("name", inline_image.name().unwrap_or(default_name.as_str())));
    if let Some(description) = inline_image.description() {
        doc_pr.push_attribute(("descr", description));
    }
    writer.write_event(Event::Empty(doc_pr))?;

    writer.write_event(Event::Start(BytesStart::new("a:graphic")))?;
    let mut graphic_data = BytesStart::new("a:graphicData");
    graphic_data.push_attribute(("uri", DRAWINGML_PICTURE_URI));
    writer.write_event(Event::Start(graphic_data))?;

    writer.write_event(Event::Start(BytesStart::new("pic:pic")))?;

    writer.write_event(Event::Start(BytesStart::new("pic:nvPicPr")))?;
    let mut c_nv_pr = BytesStart::new("pic:cNvPr");
    c_nv_pr.push_attribute(("id", drawing_id_text.as_str()));
    c_nv_pr.push_attribute(("name", inline_image.name().unwrap_or(default_name.as_str())));
    if let Some(description) = inline_image.description() {
        c_nv_pr.push_attribute(("descr", description));
    }
    writer.write_event(Event::Empty(c_nv_pr))?;
    writer.write_event(Event::Empty(BytesStart::new("pic:cNvPicPr")))?;
    writer.write_event(Event::End(BytesEnd::new("pic:nvPicPr")))?;

    writer.write_event(Event::Start(BytesStart::new("pic:blipFill")))?;
    let mut blip = BytesStart::new("a:blip");
    blip.push_attribute(("r:embed", relationship_id.as_str()));
    writer.write_event(Event::Empty(blip))?;
    writer.write_event(Event::Start(BytesStart::new("a:stretch")))?;
    writer.write_event(Event::Empty(BytesStart::new("a:fillRect")))?;
    writer.write_event(Event::End(BytesEnd::new("a:stretch")))?;
    writer.write_event(Event::End(BytesEnd::new("pic:blipFill")))?;

    writer.write_event(Event::Start(BytesStart::new("pic:spPr")))?;
    writer.write_event(Event::Start(BytesStart::new("a:xfrm")))?;
    let mut offset = BytesStart::new("a:off");
    offset.push_attribute(("x", "0"));
    offset.push_attribute(("y", "0"));
    writer.write_event(Event::Empty(offset))?;
    let mut ext = BytesStart::new("a:ext");
    ext.push_attribute(("cx", width_text.as_str()));
    ext.push_attribute(("cy", height_text.as_str()));
    writer.write_event(Event::Empty(ext))?;
    writer.write_event(Event::End(BytesEnd::new("a:xfrm")))?;
    let mut geometry = BytesStart::new("a:prstGeom");
    geometry.push_attribute(("prst", "rect"));
    writer.write_event(Event::Start(geometry))?;
    writer.write_event(Event::Empty(BytesStart::new("a:avLst")))?;
    writer.write_event(Event::End(BytesEnd::new("a:prstGeom")))?;
    writer.write_event(Event::End(BytesEnd::new("pic:spPr")))?;

    writer.write_event(Event::End(BytesEnd::new("pic:pic")))?;
    writer.write_event(Event::End(BytesEnd::new("a:graphicData")))?;
    writer.write_event(Event::End(BytesEnd::new("a:graphic")))?;

    writer.write_event(Event::End(BytesEnd::new("wp:inline")))?;
    writer.write_event(Event::End(BytesEnd::new("w:drawing")))?;

    Ok(())
}

fn write_floating_drawing_xml<W: Write>(
    writer: &mut Writer<W>,
    floating_image: &FloatingImage,
    image_relationship_ids: &HashMap<usize, String>,
    next_drawing_id: &mut u32,
) -> Result<()> {
    let Some(relationship_id) = image_relationship_ids.get(&floating_image.image_index()) else {
        return Ok(());
    };

    let drawing_id = *next_drawing_id;
    *next_drawing_id = next_drawing_id.checked_add(1).ok_or_else(|| {
        DocxError::UnsupportedPackage("drawing id overflow while serializing".to_string())
    })?;

    let width = floating_image.width_emu();
    let height = floating_image.height_emu();
    let offset_x = floating_image.offset_x_emu();
    let offset_y = floating_image.offset_y_emu();

    writer.write_event(Event::Start(BytesStart::new("w:drawing")))?;

    let mut anchor = BytesStart::new("wp:anchor");
    anchor.push_attribute(("distT", "0"));
    anchor.push_attribute(("distB", "0"));
    anchor.push_attribute(("distL", "0"));
    anchor.push_attribute(("distR", "0"));
    anchor.push_attribute(("simplePos", "0"));
    anchor.push_attribute(("relativeHeight", "0"));
    anchor.push_attribute(("behindDoc", "0"));
    anchor.push_attribute(("locked", "0"));
    anchor.push_attribute(("layoutInCell", "1"));
    anchor.push_attribute(("allowOverlap", "1"));
    writer.write_event(Event::Start(anchor))?;

    let mut simple_pos = BytesStart::new("wp:simplePos");
    let offset_x_text = offset_x.to_string();
    let offset_y_text = offset_y.to_string();
    simple_pos.push_attribute(("x", offset_x_text.as_str()));
    simple_pos.push_attribute(("y", offset_y_text.as_str()));
    writer.write_event(Event::Empty(simple_pos))?;

    let mut position_h = BytesStart::new("wp:positionH");
    position_h.push_attribute(("relativeFrom", "page"));
    writer.write_event(Event::Start(position_h))?;
    writer.write_event(Event::Start(BytesStart::new("wp:posOffset")))?;
    writer.write_event(Event::Text(BytesText::new(offset_x_text.as_str())))?;
    writer.write_event(Event::End(BytesEnd::new("wp:posOffset")))?;
    writer.write_event(Event::End(BytesEnd::new("wp:positionH")))?;

    let mut position_v = BytesStart::new("wp:positionV");
    position_v.push_attribute(("relativeFrom", "page"));
    writer.write_event(Event::Start(position_v))?;
    writer.write_event(Event::Start(BytesStart::new("wp:posOffset")))?;
    writer.write_event(Event::Text(BytesText::new(offset_y_text.as_str())))?;
    writer.write_event(Event::End(BytesEnd::new("wp:posOffset")))?;
    writer.write_event(Event::End(BytesEnd::new("wp:positionV")))?;

    let mut extent = BytesStart::new("wp:extent");
    let width_text = width.to_string();
    let height_text = height.to_string();
    extent.push_attribute(("cx", width_text.as_str()));
    extent.push_attribute(("cy", height_text.as_str()));
    writer.write_event(Event::Empty(extent))?;

    writer.write_event(Event::Empty(BytesStart::new("wp:wrapNone")))?;

    let mut doc_pr = BytesStart::new("wp:docPr");
    let drawing_id_text = drawing_id.to_string();
    let default_name = format!("Picture {drawing_id}");
    doc_pr.push_attribute(("id", drawing_id_text.as_str()));
    doc_pr.push_attribute((
        "name",
        floating_image.name().unwrap_or(default_name.as_str()),
    ));
    if let Some(description) = floating_image.description() {
        doc_pr.push_attribute(("descr", description));
    }
    writer.write_event(Event::Empty(doc_pr))?;

    writer.write_event(Event::Start(BytesStart::new("a:graphic")))?;
    let mut graphic_data = BytesStart::new("a:graphicData");
    graphic_data.push_attribute(("uri", DRAWINGML_PICTURE_URI));
    writer.write_event(Event::Start(graphic_data))?;

    writer.write_event(Event::Start(BytesStart::new("pic:pic")))?;

    writer.write_event(Event::Start(BytesStart::new("pic:nvPicPr")))?;
    let mut c_nv_pr = BytesStart::new("pic:cNvPr");
    c_nv_pr.push_attribute(("id", drawing_id_text.as_str()));
    c_nv_pr.push_attribute((
        "name",
        floating_image.name().unwrap_or(default_name.as_str()),
    ));
    if let Some(description) = floating_image.description() {
        c_nv_pr.push_attribute(("descr", description));
    }
    writer.write_event(Event::Empty(c_nv_pr))?;
    writer.write_event(Event::Empty(BytesStart::new("pic:cNvPicPr")))?;
    writer.write_event(Event::End(BytesEnd::new("pic:nvPicPr")))?;

    writer.write_event(Event::Start(BytesStart::new("pic:blipFill")))?;
    let mut blip = BytesStart::new("a:blip");
    blip.push_attribute(("r:embed", relationship_id.as_str()));
    writer.write_event(Event::Empty(blip))?;
    writer.write_event(Event::Start(BytesStart::new("a:stretch")))?;
    writer.write_event(Event::Empty(BytesStart::new("a:fillRect")))?;
    writer.write_event(Event::End(BytesEnd::new("a:stretch")))?;
    writer.write_event(Event::End(BytesEnd::new("pic:blipFill")))?;

    writer.write_event(Event::Start(BytesStart::new("pic:spPr")))?;
    writer.write_event(Event::Start(BytesStart::new("a:xfrm")))?;
    let mut offset = BytesStart::new("a:off");
    offset.push_attribute(("x", "0"));
    offset.push_attribute(("y", "0"));
    writer.write_event(Event::Empty(offset))?;
    let mut ext = BytesStart::new("a:ext");
    ext.push_attribute(("cx", width_text.as_str()));
    ext.push_attribute(("cy", height_text.as_str()));
    writer.write_event(Event::Empty(ext))?;
    writer.write_event(Event::End(BytesEnd::new("a:xfrm")))?;
    let mut geometry = BytesStart::new("a:prstGeom");
    geometry.push_attribute(("prst", "rect"));
    writer.write_event(Event::Start(geometry))?;
    writer.write_event(Event::Empty(BytesStart::new("a:avLst")))?;
    writer.write_event(Event::End(BytesEnd::new("a:prstGeom")))?;
    writer.write_event(Event::End(BytesEnd::new("pic:spPr")))?;

    writer.write_event(Event::End(BytesEnd::new("pic:pic")))?;
    writer.write_event(Event::End(BytesEnd::new("a:graphicData")))?;
    writer.write_event(Event::End(BytesEnd::new("a:graphic")))?;

    writer.write_event(Event::End(BytesEnd::new("wp:anchor")))?;
    writer.write_event(Event::End(BytesEnd::new("w:drawing")))?;

    Ok(())
}

fn write_table_xml<W: Write>(writer: &mut Writer<W>, table: &Table) -> Result<()> {
    writer.write_event(Event::Start(BytesStart::new("w:tbl")))?;

    if table.style_id().is_some()
        || !table.borders().is_empty()
        || table.alignment().is_some()
        || table.width_twips().is_some()
        || table.width_type().is_some()
        || table.layout().is_some()
        || table.first_row()
        || table.last_row()
        || table.first_column()
        || table.last_column()
        || table.no_h_band()
        || table.no_v_band()
        || !table.unknown_property_children().is_empty()
    {
        writer.write_event(Event::Start(BytesStart::new("w:tblPr")))?;
        if let Some(style_id) = table.style_id() {
            let mut style = BytesStart::new("w:tblStyle");
            style.push_attribute(("w:val", style_id));
            writer.write_event(Event::Empty(style))?;
        }
        if table.width_twips().is_some() || table.width_type().is_some() {
            let mut tbl_w = BytesStart::new("w:tblW");
            let width_value = table.width_twips().unwrap_or(0).to_string();
            tbl_w.push_attribute(("w:w", width_value.as_str()));
            let type_str = match table.width_type() {
                Some(TableWidthType::Dxa) => "dxa",
                Some(TableWidthType::Pct) => "pct",
                Some(TableWidthType::Auto) => "auto",
                None => "dxa",
            };
            tbl_w.push_attribute(("w:type", type_str));
            writer.write_event(Event::Empty(tbl_w))?;
        }
        if let Some(alignment) = table.alignment() {
            let mut jc = BytesStart::new("w:jc");
            let val = match alignment {
                TableAlignment::Left => "start",
                TableAlignment::Center => "center",
                TableAlignment::Right => "end",
            };
            jc.push_attribute(("w:val", val));
            writer.write_event(Event::Empty(jc))?;
        }
        if !table.borders().is_empty() {
            write_table_borders_xml(writer, table.borders())?;
        }
        if let Some(layout) = table.layout() {
            let mut tbl_layout = BytesStart::new("w:tblLayout");
            let val = match layout {
                TableLayout::Fixed => "fixed",
                TableLayout::AutoFit => "autofit",
            };
            tbl_layout.push_attribute(("w:type", val));
            writer.write_event(Event::Empty(tbl_layout))?;
        }
        if table.first_row()
            || table.last_row()
            || table.first_column()
            || table.last_column()
            || table.no_h_band()
            || table.no_v_band()
        {
            let mut tbl_look = BytesStart::new("w:tblLook");
            tbl_look.push_attribute(("w:firstRow", if table.first_row() { "1" } else { "0" }));
            tbl_look.push_attribute(("w:lastRow", if table.last_row() { "1" } else { "0" }));
            tbl_look.push_attribute((
                "w:firstColumn",
                if table.first_column() { "1" } else { "0" },
            ));
            tbl_look.push_attribute(("w:lastColumn", if table.last_column() { "1" } else { "0" }));
            tbl_look.push_attribute(("w:noHBand", if table.no_h_band() { "1" } else { "0" }));
            tbl_look.push_attribute(("w:noVBand", if table.no_v_band() { "1" } else { "0" }));
            writer.write_event(Event::Empty(tbl_look))?;
        }
        for node in table.unknown_property_children() {
            node.write_to(writer)?;
        }
        writer.write_event(Event::End(BytesEnd::new("w:tblPr")))?;
    }

    if !table.column_widths_twips().is_empty() {
        writer.write_event(Event::Start(BytesStart::new("w:tblGrid")))?;
        for width in table.column_widths_twips() {
            let mut grid_col = BytesStart::new("w:gridCol");
            let width_value = width.to_string();
            grid_col.push_attribute(("w:w", width_value.as_str()));
            writer.write_event(Event::Empty(grid_col))?;
        }
        writer.write_event(Event::End(BytesEnd::new("w:tblGrid")))?;
    }

    for row in 0..table.rows() {
        writer.write_event(Event::Start(BytesStart::new("w:tr")))?;
        if let Some(row_props) = table.row_properties(row) {
            if !row_props.is_empty() {
                write_table_row_properties_xml(writer, row_props)?;
            }
        }
        let mut column = 0_usize;
        while column < table.columns() {
            let Some(cell) = table.cell(row, column) else {
                column = column.saturating_add(1);
                continue;
            };
            if cell.is_horizontal_merge_continuation() {
                column = column.saturating_add(1);
                continue;
            }

            writer.write_event(Event::Start(BytesStart::new("w:tc")))?;
            let span = cell.horizontal_span().max(1);
            let has_cell_props = span > 1
                || cell.vertical_merge().is_some()
                || cell.shading_color().is_some()
                || cell.vertical_alignment().is_some()
                || cell.cell_width_twips().is_some()
                || !cell.borders().is_empty()
                || !cell.margins().is_empty()
                || !cell.unknown_property_children().is_empty();
            if has_cell_props {
                writer.write_event(Event::Start(BytesStart::new("w:tcPr")))?;
                if let Some(width) = cell.cell_width_twips() {
                    let mut tc_w = BytesStart::new("w:tcW");
                    let width_value = width.to_string();
                    tc_w.push_attribute(("w:w", width_value.as_str()));
                    tc_w.push_attribute(("w:type", "dxa"));
                    writer.write_event(Event::Empty(tc_w))?;
                }
                if span > 1 {
                    let mut grid_span = BytesStart::new("w:gridSpan");
                    let span_value = span.to_string();
                    grid_span.push_attribute(("w:val", span_value.as_str()));
                    writer.write_event(Event::Empty(grid_span))?;
                }
                if let Some(vertical_merge) = cell.vertical_merge() {
                    let mut v_merge = BytesStart::new("w:vMerge");
                    if matches!(vertical_merge, VerticalMerge::Restart) {
                        v_merge.push_attribute(("w:val", "restart"));
                    }
                    writer.write_event(Event::Empty(v_merge))?;
                }
                if !cell.borders().is_empty() {
                    write_cell_borders_xml(writer, cell.borders())?;
                }
                if let Some(shading_color) = cell.shading_color() {
                    let mut shd = BytesStart::new("w:shd");
                    if let Some(pattern) = cell.shading_pattern() {
                        shd.push_attribute(("w:val", pattern));
                    } else {
                        shd.push_attribute(("w:val", "clear"));
                    }
                    if let Some(color_attr) = cell.shading_color_attribute() {
                        shd.push_attribute(("w:color", color_attr));
                    } else {
                        shd.push_attribute(("w:color", "auto"));
                    }
                    shd.push_attribute(("w:fill", shading_color));
                    writer.write_event(Event::Empty(shd))?;
                }
                if !cell.margins().is_empty() {
                    write_cell_margins_xml(writer, cell.margins())?;
                }
                if let Some(vertical_alignment) = cell.vertical_alignment() {
                    let mut v_align = BytesStart::new("w:vAlign");
                    let val = match vertical_alignment {
                        VerticalAlignment::Top => "top",
                        VerticalAlignment::Center => "center",
                        VerticalAlignment::Bottom => "bottom",
                    };
                    v_align.push_attribute(("w:val", val));
                    writer.write_event(Event::Empty(v_align))?;
                }
                for node in cell.unknown_property_children() {
                    node.write_to(writer)?;
                }
                writer.write_event(Event::End(BytesEnd::new("w:tcPr")))?;
            }
            writer.write_event(Event::Start(BytesStart::new("w:p")))?;
            writer.write_event(Event::Start(BytesStart::new("w:r")))?;

            let cell_text = cell.text();
            let mut text = BytesStart::new("w:t");
            if cell_text.starts_with(' ') || cell_text.ends_with(' ') {
                text.push_attribute(("xml:space", "preserve"));
            }
            writer.write_event(Event::Start(text))?;
            writer.write_event(Event::Text(BytesText::new(cell_text)))?;
            writer.write_event(Event::End(BytesEnd::new("w:t")))?;

            writer.write_event(Event::End(BytesEnd::new("w:r")))?;
            writer.write_event(Event::End(BytesEnd::new("w:p")))?;
            writer.write_event(Event::End(BytesEnd::new("w:tc")))?;

            column = column.saturating_add(span);
        }
        writer.write_event(Event::End(BytesEnd::new("w:tr")))?;
    }

    writer.write_event(Event::End(BytesEnd::new("w:tbl")))?;
    Ok(())
}

fn write_table_borders_xml<W: Write>(writer: &mut Writer<W>, borders: &TableBorders) -> Result<()> {
    writer.write_event(Event::Start(BytesStart::new("w:tblBorders")))?;
    write_table_border_edge_xml(writer, "w:top", borders.top())?;
    write_table_border_edge_xml(writer, "w:left", borders.left())?;
    write_table_border_edge_xml(writer, "w:bottom", borders.bottom())?;
    write_table_border_edge_xml(writer, "w:right", borders.right())?;
    write_table_border_edge_xml(writer, "w:insideH", borders.inside_horizontal())?;
    write_table_border_edge_xml(writer, "w:insideV", borders.inside_vertical())?;
    writer.write_event(Event::End(BytesEnd::new("w:tblBorders")))?;
    Ok(())
}

fn write_table_border_edge_xml<W: Write>(
    writer: &mut Writer<W>,
    element_name: &str,
    border: Option<&TableBorder>,
) -> Result<()> {
    let Some(border) = border else {
        return Ok(());
    };

    let mut edge = BytesStart::new(element_name);
    if let Some(line_type) = border.line_type() {
        edge.push_attribute(("w:val", line_type));
    }
    if let Some(size) = border.size_eighth_points() {
        let size_value = size.to_string();
        edge.push_attribute(("w:sz", size_value.as_str()));
    }
    if let Some(color) = border.color() {
        edge.push_attribute(("w:color", color));
    }
    if let Some(space) = border.space_eighth_points() {
        let space_val = space.to_string();
        edge.push_attribute(("w:space", space_val.as_str()));
    }
    writer.write_event(Event::Empty(edge))?;
    Ok(())
}

fn write_tab_stops_xml<W: Write>(writer: &mut Writer<W>, tab_stops: &[TabStop]) -> Result<()> {
    writer.write_event(Event::Start(BytesStart::new("w:tabs")))?;
    for tab_stop in tab_stops {
        let mut tab = BytesStart::new("w:tab");
        tab.push_attribute(("w:val", tab_stop.alignment().to_xml_value()));
        let pos_value = tab_stop.position_twips().to_string();
        tab.push_attribute(("w:pos", pos_value.as_str()));
        if let Some(leader) = tab_stop.leader() {
            tab.push_attribute(("w:leader", leader.to_xml_value()));
        }
        if tab_stop.num_tab() {
            tab.push_attribute(("w:numTab", "1"));
        }
        writer.write_event(Event::Empty(tab))?;
    }
    writer.write_event(Event::End(BytesEnd::new("w:tabs")))?;
    Ok(())
}

fn write_paragraph_borders_xml<W: Write>(
    writer: &mut Writer<W>,
    borders: &ParagraphBorders,
) -> Result<()> {
    writer.write_event(Event::Start(BytesStart::new("w:pBdr")))?;
    write_paragraph_border_edge_xml(writer, "w:top", borders.top())?;
    write_paragraph_border_edge_xml(writer, "w:left", borders.left())?;
    write_paragraph_border_edge_xml(writer, "w:bottom", borders.bottom())?;
    write_paragraph_border_edge_xml(writer, "w:right", borders.right())?;
    write_paragraph_border_edge_xml(writer, "w:between", borders.between())?;
    writer.write_event(Event::End(BytesEnd::new("w:pBdr")))?;
    Ok(())
}

fn write_paragraph_border_edge_xml<W: Write>(
    writer: &mut Writer<W>,
    element_name: &str,
    border: Option<&ParagraphBorder>,
) -> Result<()> {
    let Some(border) = border else {
        return Ok(());
    };

    let mut edge = BytesStart::new(element_name);
    if let Some(line_type) = border.line_type() {
        edge.push_attribute(("w:val", line_type));
    }
    if let Some(size) = border.size_eighth_points() {
        let size_value = size.to_string();
        edge.push_attribute(("w:sz", size_value.as_str()));
    }
    if let Some(color) = border.color() {
        edge.push_attribute(("w:color", color));
    }
    if let Some(space) = border.space_points() {
        let space_value = space.to_string();
        edge.push_attribute(("w:space", space_value.as_str()));
    }
    writer.write_event(Event::Empty(edge))?;
    Ok(())
}

fn write_table_row_properties_xml<W: Write>(
    writer: &mut Writer<W>,
    row_props: &TableRowProperties,
) -> Result<()> {
    writer.write_event(Event::Start(BytesStart::new("w:trPr")))?;
    if row_props.repeat_header() {
        writer.write_event(Event::Empty(BytesStart::new("w:tblHeader")))?;
    }
    if let Some(height) = row_props.height_twips() {
        let mut tr_height = BytesStart::new("w:trHeight");
        let height_value = height.to_string();
        tr_height.push_attribute(("w:val", height_value.as_str()));
        if let Some(rule) = row_props.height_rule() {
            tr_height.push_attribute(("w:hRule", rule));
        }
        writer.write_event(Event::Empty(tr_height))?;
    }
    writer.write_event(Event::End(BytesEnd::new("w:trPr")))?;
    Ok(())
}

fn write_cell_borders_xml<W: Write>(writer: &mut Writer<W>, borders: &CellBorders) -> Result<()> {
    writer.write_event(Event::Start(BytesStart::new("w:tcBorders")))?;
    write_table_border_edge_xml(writer, "w:top", borders.top())?;
    write_table_border_edge_xml(writer, "w:left", borders.left())?;
    write_table_border_edge_xml(writer, "w:bottom", borders.bottom())?;
    write_table_border_edge_xml(writer, "w:right", borders.right())?;
    writer.write_event(Event::End(BytesEnd::new("w:tcBorders")))?;
    Ok(())
}

fn write_cell_margins_xml<W: Write>(writer: &mut Writer<W>, margins: &CellMargins) -> Result<()> {
    writer.write_event(Event::Start(BytesStart::new("w:tcMar")))?;
    if let Some(top) = margins.top_twips() {
        let mut elem = BytesStart::new("w:top");
        let value = top.to_string();
        elem.push_attribute(("w:w", value.as_str()));
        elem.push_attribute(("w:type", "dxa"));
        writer.write_event(Event::Empty(elem))?;
    }
    if let Some(left) = margins.left_twips() {
        let mut elem = BytesStart::new("w:left");
        let value = left.to_string();
        elem.push_attribute(("w:w", value.as_str()));
        elem.push_attribute(("w:type", "dxa"));
        writer.write_event(Event::Empty(elem))?;
    }
    if let Some(bottom) = margins.bottom_twips() {
        let mut elem = BytesStart::new("w:bottom");
        let value = bottom.to_string();
        elem.push_attribute(("w:w", value.as_str()));
        elem.push_attribute(("w:type", "dxa"));
        writer.write_event(Event::Empty(elem))?;
    }
    if let Some(right) = margins.right_twips() {
        let mut elem = BytesStart::new("w:right");
        let value = right.to_string();
        elem.push_attribute(("w:w", value.as_str()));
        elem.push_attribute(("w:type", "dxa"));
        writer.write_event(Event::Empty(elem))?;
    }
    writer.write_event(Event::End(BytesEnd::new("w:tcMar")))?;
    Ok(())
}

fn write_section_xml<W: Write>(
    writer: &mut Writer<W>,
    section: &Section,
    section_relationship_ids: &SectionRelationshipIds,
) -> Result<()> {
    writer.write_event(Event::Start(BytesStart::new("w:sectPr")))?;

    if let Some(header_relationship_id) = section_relationship_ids.header_relationship_id.as_deref()
    {
        let mut reference = BytesStart::new("w:headerReference");
        reference.push_attribute(("w:type", "default"));
        reference.push_attribute(("r:id", header_relationship_id));
        writer.write_event(Event::Empty(reference))?;
    }
    if let Some(footer_relationship_id) = section_relationship_ids.footer_relationship_id.as_deref()
    {
        let mut reference = BytesStart::new("w:footerReference");
        reference.push_attribute(("w:type", "default"));
        reference.push_attribute(("r:id", footer_relationship_id));
        writer.write_event(Event::Empty(reference))?;
    }
    if let Some(first_header_id) = section_relationship_ids
        .first_page_header_relationship_id
        .as_deref()
    {
        let mut reference = BytesStart::new("w:headerReference");
        reference.push_attribute(("w:type", "first"));
        reference.push_attribute(("r:id", first_header_id));
        writer.write_event(Event::Empty(reference))?;
    }
    if let Some(first_footer_id) = section_relationship_ids
        .first_page_footer_relationship_id
        .as_deref()
    {
        let mut reference = BytesStart::new("w:footerReference");
        reference.push_attribute(("w:type", "first"));
        reference.push_attribute(("r:id", first_footer_id));
        writer.write_event(Event::Empty(reference))?;
    }
    if let Some(even_header_id) = section_relationship_ids
        .even_page_header_relationship_id
        .as_deref()
    {
        let mut reference = BytesStart::new("w:headerReference");
        reference.push_attribute(("w:type", "even"));
        reference.push_attribute(("r:id", even_header_id));
        writer.write_event(Event::Empty(reference))?;
    }
    if let Some(even_footer_id) = section_relationship_ids
        .even_page_footer_relationship_id
        .as_deref()
    {
        let mut reference = BytesStart::new("w:footerReference");
        reference.push_attribute(("w:type", "even"));
        reference.push_attribute(("r:id", even_footer_id));
        writer.write_event(Event::Empty(reference))?;
    }
    if let Some(break_type) = section.break_type() {
        let mut type_elem = BytesStart::new("w:type");
        type_elem.push_attribute(("w:val", break_type.to_xml_value()));
        writer.write_event(Event::Empty(type_elem))?;
    }

    if section.page_width_twips().is_some()
        || section.page_height_twips().is_some()
        || section.page_orientation().is_some()
    {
        let mut page_size = BytesStart::new("w:pgSz");
        let width = section.page_width_twips().map(|value| value.to_string());
        let height = section.page_height_twips().map(|value| value.to_string());
        if let Some(value) = width.as_deref() {
            page_size.push_attribute(("w:w", value));
        }
        if let Some(value) = height.as_deref() {
            page_size.push_attribute(("w:h", value));
        }
        if let Some(orientation) = section.page_orientation() {
            page_size.push_attribute(("w:orient", orientation.to_xml_value()));
        }
        writer.write_event(Event::Empty(page_size))?;
    }

    let margins = section.page_margins();
    if !margins.is_empty() {
        let mut page_margins = BytesStart::new("w:pgMar");
        let top = margins.top_twips().map(|value| value.to_string());
        let right = margins.right_twips().map(|value| value.to_string());
        let bottom = margins.bottom_twips().map(|value| value.to_string());
        let left = margins.left_twips().map(|value| value.to_string());
        let header = margins.header_twips().map(|value| value.to_string());
        let footer = margins.footer_twips().map(|value| value.to_string());
        let gutter = margins.gutter_twips().map(|value| value.to_string());

        if let Some(value) = top.as_deref() {
            page_margins.push_attribute(("w:top", value));
        }
        if let Some(value) = right.as_deref() {
            page_margins.push_attribute(("w:right", value));
        }
        if let Some(value) = bottom.as_deref() {
            page_margins.push_attribute(("w:bottom", value));
        }
        if let Some(value) = left.as_deref() {
            page_margins.push_attribute(("w:left", value));
        }
        if let Some(value) = header.as_deref() {
            page_margins.push_attribute(("w:header", value));
        }
        if let Some(value) = footer.as_deref() {
            page_margins.push_attribute(("w:footer", value));
        }
        if let Some(value) = gutter.as_deref() {
            page_margins.push_attribute(("w:gutter", value));
        }
        writer.write_event(Event::Empty(page_margins))?;
    }

    if section.page_number_start().is_some() || section.page_number_format().is_some() {
        let mut pg_num_type = BytesStart::new("w:pgNumType");
        if let Some(format) = section.page_number_format() {
            pg_num_type.push_attribute(("w:fmt", format));
        }
        let start_value = section.page_number_start().map(|v| v.to_string());
        if let Some(value) = start_value.as_deref() {
            pg_num_type.push_attribute(("w:start", value));
        }
        writer.write_event(Event::Empty(pg_num_type))?;
    }

    // Multi-column layout
    if section.column_count().is_some()
        || section.column_space_twips().is_some()
        || section.column_separator()
    {
        let mut cols = BytesStart::new("w:cols");
        let num_str = section.column_count().map(|v| v.to_string());
        if let Some(ref num) = num_str {
            cols.push_attribute(("w:num", num.as_str()));
        }
        let space_str = section.column_space_twips().map(|v| v.to_string());
        if let Some(ref space) = space_str {
            cols.push_attribute(("w:space", space.as_str()));
        }
        if section.column_separator() {
            cols.push_attribute(("w:sep", "1"));
        }
        writer.write_event(Event::Empty(cols))?;
    }

    // Vertical alignment
    if let Some(v_align) = section.vertical_alignment() {
        let mut elem = BytesStart::new("w:vAlign");
        elem.push_attribute(("w:val", v_align.to_xml_value()));
        writer.write_event(Event::Empty(elem))?;
    }

    // Line numbering
    if section.line_numbering_start().is_some()
        || section.line_numbering_count_by().is_some()
        || section.line_numbering_restart().is_some()
        || section.line_numbering_distance_twips().is_some()
    {
        let mut ln_num = BytesStart::new("w:lnNumType");
        let start_str = section.line_numbering_start().map(|v| v.to_string());
        if let Some(ref start) = start_str {
            ln_num.push_attribute(("w:start", start.as_str()));
        }
        let count_str = section.line_numbering_count_by().map(|v| v.to_string());
        if let Some(ref count) = count_str {
            ln_num.push_attribute(("w:countBy", count.as_str()));
        }
        if let Some(restart) = section.line_numbering_restart() {
            ln_num.push_attribute(("w:restart", restart.to_xml_value()));
        }
        let dist_str = section
            .line_numbering_distance_twips()
            .map(|v| v.to_string());
        if let Some(ref dist) = dist_str {
            ln_num.push_attribute(("w:distance", dist.as_str()));
        }
        writer.write_event(Event::Empty(ln_num))?;
    }

    if section.title_page() {
        writer.write_event(Event::Empty(BytesStart::new("w:titlePg")))?;
    }

    writer.write_event(Event::End(BytesEnd::new("w:sectPr")))?;
    Ok(())
}

fn parse_paragraphs(
    xml: &[u8],
    hyperlink_targets: &HashMap<String, String>,
    image_indexes_by_relationship_id: &HashMap<String, usize>,
) -> Result<Vec<Paragraph>> {
    let mut reader = Reader::from_reader(Cursor::new(xml));
    reader.config_mut().trim_text(false);

    let mut paragraphs = Vec::new();
    let mut in_body = false;
    let mut table_depth = 0_usize;
    let mut hyperlink_depth = 0_usize;
    let mut in_text = false;
    let mut current_paragraph: Option<Paragraph> = None;
    let mut current_run_text: Option<String> = None;
    let mut current_run_properties = CurrentRunProperties::default();
    let mut current_hyperlink_target: Option<String> = None;
    let mut current_hyperlink_tooltip: Option<String> = None;
    let mut current_hyperlink_anchor: Option<String> = None;
    let mut position_h_depth = 0_usize;
    let mut position_v_depth = 0_usize;
    let mut in_position_h_offset = false;
    let mut in_position_v_offset = false;
    let mut paragraph_properties_depth = 0_usize;
    let mut run_properties_depth = 0_usize;
    let mut drawing_depth = 0_usize;
    let mut in_instr_text = false;
    let mut buffer = Vec::new();

    const KNOWN_PARAGRAPH_CHILDREN: &[&[u8]] = &[
        b"pPr",
        b"r",
        b"hyperlink",
        b"bookmarkStart",
        b"bookmarkEnd",
        b"commentRangeStart",
        b"commentRangeEnd",
    ];
    const KNOWN_PARAGRAPH_PROPERTY_CHILDREN: &[&[u8]] = &[
        b"pStyle",
        b"jc",
        b"spacing",
        b"ind",
        b"numPr",
        b"numId",
        b"ilvl",
        b"tabs",
        b"tab",
        b"pBdr",
        b"shd",
        b"keepNext",
        b"keepLines",
        b"widowControl",
        b"bidi",
        b"pageBreakBefore",
        b"contextualSpacing",
        b"outlineLvl",
        b"sectPr",
    ];
    const KNOWN_RUN_CHILDREN: &[&[u8]] = &[
        b"rPr",
        b"t",
        b"drawing",
        b"tab",
        b"br",
        b"footnoteReference",
        b"fldChar",
        b"instrText",
        b"endnoteReference",
    ];
    const KNOWN_RUN_PROPERTY_CHILDREN: &[&[u8]] = &[
        b"b",
        b"i",
        b"u",
        b"strike",
        b"dstrike",
        b"vertAlign",
        b"highlight",
        b"rStyle",
        b"rFonts",
        b"sz",
        b"color",
        b"rtl",
        b"smallCaps",
        b"caps",
        b"vanish",
        b"spacing",
        b"emboss",
        b"imprint",
        b"shadow",
        b"outline",
    ];

    loop {
        match reader.read_event_into(&mut buffer)? {
            Event::Start(ref event) => {
                if drawing_depth > 0 {
                    drawing_depth = drawing_depth.saturating_add(1);
                }
                if matches_local_name(event.name().as_ref(), b"body")
                    || matches_local_name(event.name().as_ref(), b"hdr")
                    || matches_local_name(event.name().as_ref(), b"ftr")
                {
                    in_body = true;
                } else if in_body && matches_local_name(event.name().as_ref(), b"tbl") {
                    table_depth = table_depth.saturating_add(1);
                } else if in_body
                    && table_depth == 0
                    && matches_local_name(event.name().as_ref(), b"p")
                {
                    current_paragraph = Some(Paragraph::new());
                } else if current_paragraph.is_some()
                    && table_depth == 0
                    && matches_local_name(event.name().as_ref(), b"hyperlink")
                {
                    hyperlink_depth = hyperlink_depth.saturating_add(1);
                    if hyperlink_depth == 1 {
                        current_hyperlink_target =
                            resolve_hyperlink_target(hyperlink_targets, event);
                        current_hyperlink_tooltip = parse_attribute_value(event, b"tooltip");
                        current_hyperlink_anchor = parse_attribute_value(event, b"anchor");
                    }
                } else if current_paragraph.is_some()
                    && table_depth == 0
                    && matches_local_name(event.name().as_ref(), b"commentRangeStart")
                {
                    if let Some(id) = parse_u32_attribute_value(event, b"id") {
                        if let Some(paragraph) = current_paragraph.as_mut() {
                            paragraph.add_comment_range_start(id);
                        }
                    }
                } else if current_paragraph.is_some()
                    && table_depth == 0
                    && matches_local_name(event.name().as_ref(), b"commentRangeEnd")
                {
                    if let Some(id) = parse_u32_attribute_value(event, b"id") {
                        if let Some(paragraph) = current_paragraph.as_mut() {
                            paragraph.add_comment_range_end(id);
                        }
                    }
                } else if current_paragraph.is_some()
                    && matches_local_name(event.name().as_ref(), b"r")
                {
                    current_run_text = Some(String::new());
                    current_run_properties.reset(
                        current_hyperlink_target.clone(),
                        current_hyperlink_tooltip.clone(),
                        current_hyperlink_anchor.clone(),
                    );
                } else if current_paragraph.is_some()
                    && matches_local_name(event.name().as_ref(), b"pPr")
                {
                    paragraph_properties_depth = 1;
                } else if current_run_text.is_some()
                    && matches_local_name(event.name().as_ref(), b"rPr")
                {
                    run_properties_depth = 1;
                } else if current_run_text.is_some()
                    && matches_local_name(event.name().as_ref(), b"drawing")
                {
                    drawing_depth = drawing_depth.saturating_add(1);
                } else if current_paragraph.is_some()
                    && matches_local_name(event.name().as_ref(), b"t")
                {
                    in_text = true;
                } else if current_paragraph.is_some()
                    && matches_local_name(event.name().as_ref(), b"pStyle")
                {
                    maybe_apply_paragraph_style(&mut current_paragraph, event);
                } else if current_paragraph.is_some()
                    && matches_local_name(event.name().as_ref(), b"jc")
                {
                    maybe_apply_paragraph_alignment(&mut current_paragraph, event);
                } else if current_paragraph.is_some()
                    && matches_local_name(event.name().as_ref(), b"spacing")
                {
                    maybe_apply_paragraph_spacing(&mut current_paragraph, event);
                } else if current_paragraph.is_some()
                    && matches_local_name(event.name().as_ref(), b"ind")
                {
                    maybe_apply_paragraph_indentation(&mut current_paragraph, event);
                } else if current_paragraph.is_some()
                    && matches_local_name(event.name().as_ref(), b"numId")
                {
                    maybe_apply_paragraph_numbering_num_id(&mut current_paragraph, event);
                } else if current_paragraph.is_some()
                    && matches_local_name(event.name().as_ref(), b"ilvl")
                {
                    maybe_apply_paragraph_numbering_ilvl(&mut current_paragraph, event);
                } else if current_paragraph.is_some()
                    && paragraph_properties_depth > 0
                    && matches_local_name(event.name().as_ref(), b"keepNext")
                {
                    if let Some(paragraph) = current_paragraph.as_mut() {
                        paragraph.set_keep_next(parse_on_off_property(event, true));
                    }
                } else if current_paragraph.is_some()
                    && paragraph_properties_depth > 0
                    && matches_local_name(event.name().as_ref(), b"keepLines")
                {
                    if let Some(paragraph) = current_paragraph.as_mut() {
                        paragraph.set_keep_lines(parse_on_off_property(event, true));
                    }
                } else if current_paragraph.is_some()
                    && paragraph_properties_depth > 0
                    && matches_local_name(event.name().as_ref(), b"widowControl")
                {
                    if let Some(paragraph) = current_paragraph.as_mut() {
                        paragraph.set_widow_control(parse_on_off_property(event, true));
                    }
                } else if current_paragraph.is_some()
                    && paragraph_properties_depth > 0
                    && matches_local_name(event.name().as_ref(), b"bidi")
                {
                    if let Some(paragraph) = current_paragraph.as_mut() {
                        paragraph.set_bidi(parse_on_off_property(event, true));
                    }
                } else if current_paragraph.is_some()
                    && paragraph_properties_depth > 0
                    && matches_local_name(event.name().as_ref(), b"pageBreakBefore")
                {
                    if let Some(paragraph) = current_paragraph.as_mut() {
                        paragraph.set_page_break_before(parse_on_off_property(event, true));
                    }
                } else if current_paragraph.is_some()
                    && paragraph_properties_depth > 0
                    && matches_local_name(event.name().as_ref(), b"contextualSpacing")
                {
                    if let Some(paragraph) = current_paragraph.as_mut() {
                        paragraph.set_contextual_spacing(parse_on_off_property(event, true));
                    }
                } else if current_paragraph.is_some()
                    && paragraph_properties_depth > 0
                    && matches_local_name(event.name().as_ref(), b"outlineLvl")
                {
                    if let Some(paragraph) = current_paragraph.as_mut() {
                        if let Some(level) = parse_u32_attribute_value(event, b"val") {
                            paragraph.set_outline_level(level as u8);
                        }
                    }
                } else if current_paragraph.is_some()
                    && paragraph_properties_depth > 0
                    && matches_local_name(event.name().as_ref(), b"tabs")
                {
                    // tabs element is a container; individual tab children are parsed below
                } else if current_paragraph.is_some()
                    && paragraph_properties_depth > 0
                    && matches_local_name(event.name().as_ref(), b"tab")
                {
                    maybe_apply_paragraph_tab_stop(&mut current_paragraph, event);
                } else if current_paragraph.is_some()
                    && paragraph_properties_depth > 0
                    && matches_local_name(event.name().as_ref(), b"pBdr")
                {
                    // pBdr is a container for border children; parsed by sub-elements
                } else if current_paragraph.is_some()
                    && paragraph_properties_depth > 0
                    && (matches_local_name(event.name().as_ref(), b"top")
                        || matches_local_name(event.name().as_ref(), b"left")
                        || matches_local_name(event.name().as_ref(), b"bottom")
                        || matches_local_name(event.name().as_ref(), b"right")
                        || matches_local_name(event.name().as_ref(), b"between"))
                {
                    maybe_apply_paragraph_border_edge(&mut current_paragraph, event);
                } else if current_paragraph.is_some()
                    && paragraph_properties_depth > 0
                    && matches_local_name(event.name().as_ref(), b"shd")
                {
                    maybe_apply_paragraph_shading(&mut current_paragraph, event);
                } else if current_paragraph.is_some()
                    && paragraph_properties_depth > 0
                    && matches_local_name(event.name().as_ref(), b"sectPr")
                {
                    // sectPr inside pPr: consume as subtree to avoid child
                    // elements (pgSz, pgMar, etc.) leaking into unknown
                    // property children. We parse basic section info from it.
                    let snippet = capture_xml_subtree(&mut reader, event.to_owned())?;
                    if let Some(para) = current_paragraph.as_mut() {
                        let inline_section = parse_inline_section_xml(snippet.as_bytes())?;
                        para.set_section_properties(inline_section);
                    }
                } else if current_run_text.is_some()
                    && matches_local_name(event.name().as_ref(), b"footnoteReference")
                {
                    if let Some(id) = parse_u32_attribute_value(event, b"id") {
                        current_run_properties.footnote_reference_id = Some(id);
                    }
                } else if current_run_text.is_some()
                    && matches_local_name(event.name().as_ref(), b"endnoteReference")
                {
                    if let Some(id) = parse_u32_attribute_value(event, b"id") {
                        current_run_properties.endnote_reference_id = Some(id);
                    }
                } else if current_run_text.is_some()
                    && matches_local_name(event.name().as_ref(), b"tab")
                    && run_properties_depth == 0
                {
                    current_run_properties.has_tab = true;
                } else if current_run_text.is_some()
                    && matches_local_name(event.name().as_ref(), b"br")
                    && run_properties_depth == 0
                {
                    current_run_properties.has_break = true;
                } else if current_run_text.is_some()
                    && matches_local_name(event.name().as_ref(), b"b")
                {
                    current_run_properties.bold = parse_on_off_property(event, true);
                } else if current_run_text.is_some()
                    && matches_local_name(event.name().as_ref(), b"i")
                {
                    current_run_properties.italic = parse_on_off_property(event, true);
                } else if current_run_text.is_some()
                    && matches_local_name(event.name().as_ref(), b"u")
                {
                    current_run_properties.underline_type = parse_underline_property(event);
                } else if current_run_text.is_some()
                    && matches_local_name(event.name().as_ref(), b"rStyle")
                {
                    current_run_properties.style_id = parse_attribute_value(event, b"val");
                } else if current_run_text.is_some()
                    && matches_local_name(event.name().as_ref(), b"rFonts")
                {
                    current_run_properties.font_family_ascii =
                        parse_attribute_value(event, b"ascii");
                    current_run_properties.font_family_h_ansi =
                        parse_attribute_value(event, b"hAnsi");
                    current_run_properties.font_family_cs = parse_attribute_value(event, b"cs");
                    current_run_properties.font_family_east_asia =
                        parse_attribute_value(event, b"eastAsia");
                    if let Some(font_family) = parse_run_font_family(event) {
                        current_run_properties.font_family = Some(font_family);
                    }
                } else if current_run_text.is_some()
                    && matches_local_name(event.name().as_ref(), b"sz")
                {
                    current_run_properties.font_size_half_points =
                        parse_run_size_half_points(event);
                } else if current_run_text.is_some()
                    && matches_local_name(event.name().as_ref(), b"color")
                {
                    current_run_properties.color = parse_run_color(event);
                    current_run_properties.theme_color =
                        parse_attribute_value(event, b"themeColor");
                    current_run_properties.theme_shade =
                        parse_attribute_value(event, b"themeShade");
                    current_run_properties.theme_tint = parse_attribute_value(event, b"themeTint");
                } else if current_run_text.is_some()
                    && matches_local_name(event.name().as_ref(), b"strike")
                {
                    current_run_properties.strikethrough = parse_on_off_property(event, true);
                } else if current_run_text.is_some()
                    && matches_local_name(event.name().as_ref(), b"dstrike")
                {
                    current_run_properties.double_strikethrough =
                        parse_on_off_property(event, true);
                } else if current_run_text.is_some()
                    && matches_local_name(event.name().as_ref(), b"vertAlign")
                {
                    parse_vert_align_property(event, &mut current_run_properties);
                } else if current_run_text.is_some()
                    && matches_local_name(event.name().as_ref(), b"highlight")
                {
                    current_run_properties.highlight_color = parse_attribute_value(event, b"val");
                } else if current_run_text.is_some()
                    && run_properties_depth > 0
                    && matches_local_name(event.name().as_ref(), b"rtl")
                {
                    current_run_properties.rtl = parse_on_off_property(event, true);
                } else if current_run_text.is_some()
                    && matches_local_name(event.name().as_ref(), b"smallCaps")
                {
                    current_run_properties.small_caps = parse_on_off_property(event, true);
                } else if current_run_text.is_some()
                    && matches_local_name(event.name().as_ref(), b"caps")
                {
                    current_run_properties.all_caps = parse_on_off_property(event, true);
                } else if current_run_text.is_some()
                    && matches_local_name(event.name().as_ref(), b"vanish")
                {
                    current_run_properties.hidden = parse_on_off_property(event, true);
                } else if current_run_text.is_some()
                    && matches_local_name(event.name().as_ref(), b"emboss")
                {
                    current_run_properties.emboss = parse_on_off_property(event, true);
                } else if current_run_text.is_some()
                    && matches_local_name(event.name().as_ref(), b"imprint")
                {
                    current_run_properties.imprint = parse_on_off_property(event, true);
                } else if current_run_text.is_some()
                    && matches_local_name(event.name().as_ref(), b"shadow")
                {
                    current_run_properties.shadow = parse_on_off_property(event, true);
                } else if current_run_text.is_some()
                    && matches_local_name(event.name().as_ref(), b"outline")
                {
                    current_run_properties.outline = parse_on_off_property(event, true);
                } else if current_run_text.is_some()
                    && run_properties_depth > 0
                    && matches_local_name(event.name().as_ref(), b"spacing")
                {
                    current_run_properties.character_spacing_twips =
                        parse_i32_attribute_value(event, b"val");
                } else if current_run_text.is_some()
                    && run_properties_depth == 0
                    && matches_local_name(event.name().as_ref(), b"fldChar")
                {
                    if let Some(fld_type) = parse_attribute_value(event, b"fldCharType") {
                        match fld_type.as_str() {
                            "begin" => {
                                current_run_properties.in_field = true;
                                current_run_properties.field_separated = false;
                                current_run_properties.field_instruction = Some(String::new());
                                current_run_properties.field_result = Some(String::new());
                            }
                            "separate" => {
                                current_run_properties.field_separated = true;
                            }
                            "end" => {
                                current_run_properties.in_field = false;
                            }
                            _ => {}
                        }
                    }
                } else if current_run_text.is_some()
                    && matches_local_name(event.name().as_ref(), b"instrText")
                {
                    in_instr_text = true;
                } else if current_run_text.is_some()
                    && matches_local_name(event.name().as_ref(), b"blip")
                {
                    maybe_apply_run_image_relationship(&mut current_run_properties, event);
                } else if current_run_text.is_some()
                    && matches_local_name(event.name().as_ref(), b"extent")
                {
                    maybe_apply_run_image_extent(&mut current_run_properties, event);
                } else if current_run_text.is_some()
                    && (matches_local_name(event.name().as_ref(), b"docPr")
                        || matches_local_name(event.name().as_ref(), b"cNvPr"))
                {
                    maybe_apply_run_image_doc_properties(&mut current_run_properties, event);
                } else if current_run_text.is_some()
                    && matches_local_name(event.name().as_ref(), b"inline")
                {
                    current_run_properties.drawing_kind = Some(DrawingKind::Inline);
                } else if current_run_text.is_some()
                    && matches_local_name(event.name().as_ref(), b"anchor")
                {
                    current_run_properties.drawing_kind = Some(DrawingKind::Anchor);
                } else if current_run_text.is_some()
                    && matches_local_name(event.name().as_ref(), b"simplePos")
                {
                    maybe_apply_run_floating_image_simple_position(
                        &mut current_run_properties,
                        event,
                    );
                } else if current_run_text.is_some()
                    && matches_local_name(event.name().as_ref(), b"positionH")
                {
                    position_h_depth = position_h_depth.saturating_add(1);
                } else if current_run_text.is_some()
                    && matches_local_name(event.name().as_ref(), b"positionV")
                {
                    position_v_depth = position_v_depth.saturating_add(1);
                } else if current_run_text.is_some()
                    && matches_local_name(event.name().as_ref(), b"posOffset")
                {
                    in_position_h_offset = position_h_depth > 0;
                    in_position_v_offset = position_v_depth > 0 && !in_position_h_offset;
                } else if paragraph_properties_depth > 0 {
                    let name_bytes = event.name();
                    let local = local_name(name_bytes.as_ref());
                    if !KNOWN_PARAGRAPH_PROPERTY_CHILDREN.contains(&local) {
                        if let Some(paragraph) = current_paragraph.as_mut() {
                            paragraph.push_unknown_property_child(RawXmlNode::read_element(
                                &mut reader,
                                event,
                            )?);
                        }
                    }
                } else if run_properties_depth > 0 {
                    let name_bytes = event.name();
                    let local = local_name(name_bytes.as_ref());
                    if !KNOWN_RUN_PROPERTY_CHILDREN.contains(&local) {
                        current_run_properties
                            .unknown_property_children
                            .push(RawXmlNode::read_element(&mut reader, event)?);
                    }
                } else if current_paragraph.is_some()
                    && table_depth == 0
                    && current_run_text.is_none()
                    && hyperlink_depth == 0
                {
                    let name_bytes = event.name();
                    let local = local_name(name_bytes.as_ref());
                    if !KNOWN_PARAGRAPH_CHILDREN.contains(&local) {
                        if let Some(paragraph) = current_paragraph.as_mut() {
                            paragraph
                                .push_unknown_child(RawXmlNode::read_element(&mut reader, event)?);
                        }
                    }
                } else if current_run_text.is_some()
                    && run_properties_depth == 0
                    && drawing_depth == 0
                {
                    let name_bytes = event.name();
                    let local = local_name(name_bytes.as_ref());
                    if !KNOWN_RUN_CHILDREN.contains(&local) {
                        current_run_properties
                            .unknown_children
                            .push(RawXmlNode::read_element(&mut reader, event)?);
                    }
                }
            }
            Event::Empty(ref event) => {
                if in_body && table_depth == 0 && matches_local_name(event.name().as_ref(), b"p") {
                    paragraphs.push(Paragraph::new());
                } else if current_paragraph.is_some()
                    && table_depth == 0
                    && matches_local_name(event.name().as_ref(), b"commentRangeStart")
                {
                    if let Some(id) = parse_u32_attribute_value(event, b"id") {
                        if let Some(paragraph) = current_paragraph.as_mut() {
                            paragraph.add_comment_range_start(id);
                        }
                    }
                } else if current_paragraph.is_some()
                    && table_depth == 0
                    && matches_local_name(event.name().as_ref(), b"commentRangeEnd")
                {
                    if let Some(id) = parse_u32_attribute_value(event, b"id") {
                        if let Some(paragraph) = current_paragraph.as_mut() {
                            paragraph.add_comment_range_end(id);
                        }
                    }
                } else if current_paragraph.is_some()
                    && matches_local_name(event.name().as_ref(), b"pStyle")
                {
                    maybe_apply_paragraph_style(&mut current_paragraph, event);
                } else if current_paragraph.is_some()
                    && matches_local_name(event.name().as_ref(), b"jc")
                {
                    maybe_apply_paragraph_alignment(&mut current_paragraph, event);
                } else if current_paragraph.is_some()
                    && matches_local_name(event.name().as_ref(), b"spacing")
                {
                    maybe_apply_paragraph_spacing(&mut current_paragraph, event);
                } else if current_paragraph.is_some()
                    && matches_local_name(event.name().as_ref(), b"ind")
                {
                    maybe_apply_paragraph_indentation(&mut current_paragraph, event);
                } else if current_paragraph.is_some()
                    && matches_local_name(event.name().as_ref(), b"numId")
                {
                    maybe_apply_paragraph_numbering_num_id(&mut current_paragraph, event);
                } else if current_paragraph.is_some()
                    && matches_local_name(event.name().as_ref(), b"ilvl")
                {
                    maybe_apply_paragraph_numbering_ilvl(&mut current_paragraph, event);
                } else if current_paragraph.is_some()
                    && paragraph_properties_depth > 0
                    && matches_local_name(event.name().as_ref(), b"keepNext")
                {
                    if let Some(paragraph) = current_paragraph.as_mut() {
                        paragraph.set_keep_next(parse_on_off_property(event, true));
                    }
                } else if current_paragraph.is_some()
                    && paragraph_properties_depth > 0
                    && matches_local_name(event.name().as_ref(), b"keepLines")
                {
                    if let Some(paragraph) = current_paragraph.as_mut() {
                        paragraph.set_keep_lines(parse_on_off_property(event, true));
                    }
                } else if current_paragraph.is_some()
                    && paragraph_properties_depth > 0
                    && matches_local_name(event.name().as_ref(), b"widowControl")
                {
                    if let Some(paragraph) = current_paragraph.as_mut() {
                        paragraph.set_widow_control(parse_on_off_property(event, true));
                    }
                } else if current_paragraph.is_some()
                    && paragraph_properties_depth > 0
                    && matches_local_name(event.name().as_ref(), b"bidi")
                {
                    if let Some(paragraph) = current_paragraph.as_mut() {
                        paragraph.set_bidi(parse_on_off_property(event, true));
                    }
                } else if current_paragraph.is_some()
                    && paragraph_properties_depth > 0
                    && matches_local_name(event.name().as_ref(), b"pageBreakBefore")
                {
                    if let Some(paragraph) = current_paragraph.as_mut() {
                        paragraph.set_page_break_before(parse_on_off_property(event, true));
                    }
                } else if current_paragraph.is_some()
                    && paragraph_properties_depth > 0
                    && matches_local_name(event.name().as_ref(), b"contextualSpacing")
                {
                    if let Some(paragraph) = current_paragraph.as_mut() {
                        paragraph.set_contextual_spacing(parse_on_off_property(event, true));
                    }
                } else if current_paragraph.is_some()
                    && paragraph_properties_depth > 0
                    && matches_local_name(event.name().as_ref(), b"outlineLvl")
                {
                    if let Some(paragraph) = current_paragraph.as_mut() {
                        if let Some(level) = parse_u32_attribute_value(event, b"val") {
                            paragraph.set_outline_level(level as u8);
                        }
                    }
                } else if current_paragraph.is_some()
                    && paragraph_properties_depth > 0
                    && matches_local_name(event.name().as_ref(), b"tab")
                {
                    maybe_apply_paragraph_tab_stop(&mut current_paragraph, event);
                } else if current_paragraph.is_some()
                    && paragraph_properties_depth > 0
                    && (matches_local_name(event.name().as_ref(), b"top")
                        || matches_local_name(event.name().as_ref(), b"left")
                        || matches_local_name(event.name().as_ref(), b"bottom")
                        || matches_local_name(event.name().as_ref(), b"right")
                        || matches_local_name(event.name().as_ref(), b"between"))
                {
                    maybe_apply_paragraph_border_edge(&mut current_paragraph, event);
                } else if current_paragraph.is_some()
                    && paragraph_properties_depth > 0
                    && matches_local_name(event.name().as_ref(), b"shd")
                    && current_run_text.is_none()
                {
                    maybe_apply_paragraph_shading(&mut current_paragraph, event);
                } else if current_run_text.is_some()
                    && matches_local_name(event.name().as_ref(), b"footnoteReference")
                {
                    if let Some(id) = parse_u32_attribute_value(event, b"id") {
                        current_run_properties.footnote_reference_id = Some(id);
                    }
                } else if current_run_text.is_some()
                    && matches_local_name(event.name().as_ref(), b"endnoteReference")
                {
                    if let Some(id) = parse_u32_attribute_value(event, b"id") {
                        current_run_properties.endnote_reference_id = Some(id);
                    }
                } else if current_run_text.is_some()
                    && matches_local_name(event.name().as_ref(), b"tab")
                    && run_properties_depth == 0
                {
                    current_run_properties.has_tab = true;
                } else if current_run_text.is_some()
                    && matches_local_name(event.name().as_ref(), b"br")
                    && run_properties_depth == 0
                {
                    current_run_properties.has_break = true;
                } else if current_paragraph.is_some()
                    && matches_local_name(event.name().as_ref(), b"r")
                {
                    if current_run_text.is_none() {
                        current_run_text = Some(String::new());
                        current_run_properties.reset(
                            current_hyperlink_target.clone(),
                            current_hyperlink_tooltip.clone(),
                            current_hyperlink_anchor.clone(),
                        );
                    }
                    finalize_current_run(
                        &mut current_paragraph,
                        &mut current_run_text,
                        &mut current_run_properties,
                        image_indexes_by_relationship_id,
                    );
                    current_run_properties.reset(
                        current_hyperlink_target.clone(),
                        current_hyperlink_tooltip.clone(),
                        current_hyperlink_anchor.clone(),
                    );
                } else if current_paragraph.is_some()
                    && matches_local_name(event.name().as_ref(), b"pPr")
                {
                    paragraph_properties_depth = 0;
                } else if current_run_text.is_some()
                    && matches_local_name(event.name().as_ref(), b"rPr")
                {
                    run_properties_depth = 0;
                } else if current_run_text.is_some()
                    && matches_local_name(event.name().as_ref(), b"drawing")
                {
                    drawing_depth = drawing_depth.saturating_sub(1);
                } else if current_run_text.is_some()
                    && matches_local_name(event.name().as_ref(), b"b")
                {
                    current_run_properties.bold = parse_on_off_property(event, true);
                } else if current_run_text.is_some()
                    && matches_local_name(event.name().as_ref(), b"i")
                {
                    current_run_properties.italic = parse_on_off_property(event, true);
                } else if current_run_text.is_some()
                    && matches_local_name(event.name().as_ref(), b"u")
                {
                    current_run_properties.underline_type = parse_underline_property(event);
                } else if current_run_text.is_some()
                    && matches_local_name(event.name().as_ref(), b"rStyle")
                {
                    current_run_properties.style_id = parse_attribute_value(event, b"val");
                } else if current_run_text.is_some()
                    && matches_local_name(event.name().as_ref(), b"rFonts")
                {
                    current_run_properties.font_family_ascii =
                        parse_attribute_value(event, b"ascii");
                    current_run_properties.font_family_h_ansi =
                        parse_attribute_value(event, b"hAnsi");
                    current_run_properties.font_family_cs = parse_attribute_value(event, b"cs");
                    current_run_properties.font_family_east_asia =
                        parse_attribute_value(event, b"eastAsia");
                    if let Some(font_family) = parse_run_font_family(event) {
                        current_run_properties.font_family = Some(font_family);
                    }
                } else if current_run_text.is_some()
                    && matches_local_name(event.name().as_ref(), b"sz")
                {
                    current_run_properties.font_size_half_points =
                        parse_run_size_half_points(event);
                } else if current_run_text.is_some()
                    && matches_local_name(event.name().as_ref(), b"color")
                {
                    current_run_properties.color = parse_run_color(event);
                    current_run_properties.theme_color =
                        parse_attribute_value(event, b"themeColor");
                    current_run_properties.theme_shade =
                        parse_attribute_value(event, b"themeShade");
                    current_run_properties.theme_tint = parse_attribute_value(event, b"themeTint");
                } else if current_run_text.is_some()
                    && matches_local_name(event.name().as_ref(), b"strike")
                {
                    current_run_properties.strikethrough = parse_on_off_property(event, true);
                } else if current_run_text.is_some()
                    && matches_local_name(event.name().as_ref(), b"dstrike")
                {
                    current_run_properties.double_strikethrough =
                        parse_on_off_property(event, true);
                } else if current_run_text.is_some()
                    && matches_local_name(event.name().as_ref(), b"vertAlign")
                {
                    parse_vert_align_property(event, &mut current_run_properties);
                } else if current_run_text.is_some()
                    && matches_local_name(event.name().as_ref(), b"highlight")
                {
                    current_run_properties.highlight_color = parse_attribute_value(event, b"val");
                } else if current_run_text.is_some()
                    && run_properties_depth > 0
                    && matches_local_name(event.name().as_ref(), b"rtl")
                {
                    current_run_properties.rtl = parse_on_off_property(event, true);
                } else if current_run_text.is_some()
                    && matches_local_name(event.name().as_ref(), b"smallCaps")
                {
                    current_run_properties.small_caps = parse_on_off_property(event, true);
                } else if current_run_text.is_some()
                    && matches_local_name(event.name().as_ref(), b"caps")
                {
                    current_run_properties.all_caps = parse_on_off_property(event, true);
                } else if current_run_text.is_some()
                    && matches_local_name(event.name().as_ref(), b"vanish")
                {
                    current_run_properties.hidden = parse_on_off_property(event, true);
                } else if current_run_text.is_some()
                    && matches_local_name(event.name().as_ref(), b"emboss")
                {
                    current_run_properties.emboss = parse_on_off_property(event, true);
                } else if current_run_text.is_some()
                    && matches_local_name(event.name().as_ref(), b"imprint")
                {
                    current_run_properties.imprint = parse_on_off_property(event, true);
                } else if current_run_text.is_some()
                    && matches_local_name(event.name().as_ref(), b"shadow")
                {
                    current_run_properties.shadow = parse_on_off_property(event, true);
                } else if current_run_text.is_some()
                    && matches_local_name(event.name().as_ref(), b"outline")
                {
                    current_run_properties.outline = parse_on_off_property(event, true);
                } else if current_run_text.is_some()
                    && run_properties_depth > 0
                    && matches_local_name(event.name().as_ref(), b"spacing")
                {
                    current_run_properties.character_spacing_twips =
                        parse_i32_attribute_value(event, b"val");
                } else if current_run_text.is_some()
                    && run_properties_depth == 0
                    && matches_local_name(event.name().as_ref(), b"fldChar")
                {
                    if let Some(fld_type) = parse_attribute_value(event, b"fldCharType") {
                        match fld_type.as_str() {
                            "begin" => {
                                current_run_properties.in_field = true;
                                current_run_properties.field_separated = false;
                                current_run_properties.field_instruction = Some(String::new());
                                current_run_properties.field_result = Some(String::new());
                            }
                            "separate" => {
                                current_run_properties.field_separated = true;
                            }
                            "end" => {
                                current_run_properties.in_field = false;
                            }
                            _ => {}
                        }
                    }
                } else if current_run_text.is_some()
                    && matches_local_name(event.name().as_ref(), b"blip")
                {
                    maybe_apply_run_image_relationship(&mut current_run_properties, event);
                } else if current_run_text.is_some()
                    && matches_local_name(event.name().as_ref(), b"extent")
                {
                    maybe_apply_run_image_extent(&mut current_run_properties, event);
                } else if current_run_text.is_some()
                    && (matches_local_name(event.name().as_ref(), b"docPr")
                        || matches_local_name(event.name().as_ref(), b"cNvPr"))
                {
                    maybe_apply_run_image_doc_properties(&mut current_run_properties, event);
                } else if current_run_text.is_some()
                    && matches_local_name(event.name().as_ref(), b"inline")
                {
                    current_run_properties.drawing_kind = Some(DrawingKind::Inline);
                } else if current_run_text.is_some()
                    && matches_local_name(event.name().as_ref(), b"anchor")
                {
                    current_run_properties.drawing_kind = Some(DrawingKind::Anchor);
                } else if current_run_text.is_some()
                    && matches_local_name(event.name().as_ref(), b"simplePos")
                {
                    maybe_apply_run_floating_image_simple_position(
                        &mut current_run_properties,
                        event,
                    );
                } else if paragraph_properties_depth > 0 {
                    let name_bytes = event.name();
                    let local = local_name(name_bytes.as_ref());
                    if !KNOWN_PARAGRAPH_PROPERTY_CHILDREN.contains(&local) {
                        if let Some(paragraph) = current_paragraph.as_mut() {
                            paragraph
                                .push_unknown_property_child(RawXmlNode::from_empty_element(event));
                        }
                    }
                } else if run_properties_depth > 0 {
                    let name_bytes = event.name();
                    let local = local_name(name_bytes.as_ref());
                    if !KNOWN_RUN_PROPERTY_CHILDREN.contains(&local) {
                        current_run_properties
                            .unknown_property_children
                            .push(RawXmlNode::from_empty_element(event));
                    }
                } else if current_paragraph.is_some()
                    && table_depth == 0
                    && current_run_text.is_none()
                    && hyperlink_depth == 0
                {
                    let name_bytes = event.name();
                    let local = local_name(name_bytes.as_ref());
                    if !KNOWN_PARAGRAPH_CHILDREN.contains(&local) {
                        if let Some(paragraph) = current_paragraph.as_mut() {
                            paragraph.push_unknown_child(RawXmlNode::from_empty_element(event));
                        }
                    }
                } else if current_run_text.is_some()
                    && run_properties_depth == 0
                    && drawing_depth == 0
                {
                    let name_bytes = event.name();
                    let local = local_name(name_bytes.as_ref());
                    if !KNOWN_RUN_CHILDREN.contains(&local) {
                        current_run_properties
                            .unknown_children
                            .push(RawXmlNode::from_empty_element(event));
                    }
                }
            }
            Event::Text(ref event) => {
                if in_instr_text && current_run_text.is_some() {
                    let text = event
                        .xml_content()
                        .map_err(quick_xml::Error::from)?
                        .into_owned();
                    if let Some(instr) = current_run_properties.field_instruction.as_mut() {
                        instr.push_str(text.trim());
                    }
                } else if in_text && current_run_text.is_some() {
                    let text = event
                        .xml_content()
                        .map_err(quick_xml::Error::from)?
                        .into_owned();
                    if current_run_properties.in_field && current_run_properties.field_separated {
                        // Text after "separate" fldChar is the field result
                        if let Some(result) = current_run_properties.field_result.as_mut() {
                            result.push_str(text.as_str());
                        }
                    } else if let Some(run_text) = current_run_text.as_mut() {
                        run_text.push_str(text.as_str());
                    }
                } else if in_position_h_offset && current_run_text.is_some() {
                    let text = event
                        .xml_content()
                        .map_err(quick_xml::Error::from)?
                        .into_owned();
                    maybe_apply_run_floating_image_position_offset(
                        &mut current_run_properties,
                        text.as_str(),
                        true,
                    );
                } else if in_position_v_offset && current_run_text.is_some() {
                    let text = event
                        .xml_content()
                        .map_err(quick_xml::Error::from)?
                        .into_owned();
                    maybe_apply_run_floating_image_position_offset(
                        &mut current_run_properties,
                        text.as_str(),
                        false,
                    );
                }
            }
            Event::CData(ref event) => {
                if in_text && current_run_text.is_some() {
                    let text = String::from_utf8_lossy(event.as_ref()).into_owned();
                    if let Some(run_text) = current_run_text.as_mut() {
                        run_text.push_str(text.as_str());
                    }
                } else if in_position_h_offset && current_run_text.is_some() {
                    let text = String::from_utf8_lossy(event.as_ref()).into_owned();
                    maybe_apply_run_floating_image_position_offset(
                        &mut current_run_properties,
                        text.as_str(),
                        true,
                    );
                } else if in_position_v_offset && current_run_text.is_some() {
                    let text = String::from_utf8_lossy(event.as_ref()).into_owned();
                    maybe_apply_run_floating_image_position_offset(
                        &mut current_run_properties,
                        text.as_str(),
                        false,
                    );
                }
            }
            Event::End(ref event) => {
                if matches_local_name(event.name().as_ref(), b"body")
                    || matches_local_name(event.name().as_ref(), b"hdr")
                    || matches_local_name(event.name().as_ref(), b"ftr")
                {
                    in_body = false;
                    table_depth = 0;
                    hyperlink_depth = 0;
                    current_hyperlink_target = None;
                } else if matches_local_name(event.name().as_ref(), b"t") {
                    in_text = false;
                } else if matches_local_name(event.name().as_ref(), b"instrText") {
                    in_instr_text = false;
                } else if matches_local_name(event.name().as_ref(), b"r") {
                    finalize_current_run(
                        &mut current_paragraph,
                        &mut current_run_text,
                        &mut current_run_properties,
                        image_indexes_by_relationship_id,
                    );
                    current_run_properties.reset(
                        current_hyperlink_target.clone(),
                        current_hyperlink_tooltip.clone(),
                        current_hyperlink_anchor.clone(),
                    );
                    position_h_depth = 0;
                    position_v_depth = 0;
                    in_position_h_offset = false;
                    in_position_v_offset = false;
                } else if matches_local_name(event.name().as_ref(), b"p") {
                    in_text = false;
                    current_run_text = None;
                    if let Some(paragraph) = current_paragraph.take() {
                        paragraphs.push(paragraph);
                    }
                    position_h_depth = 0;
                    position_v_depth = 0;
                    in_position_h_offset = false;
                    in_position_v_offset = false;
                } else if in_body && matches_local_name(event.name().as_ref(), b"hyperlink") {
                    hyperlink_depth = hyperlink_depth.saturating_sub(1);
                    if hyperlink_depth == 0 {
                        current_hyperlink_target = None;
                        current_hyperlink_tooltip = None;
                        current_hyperlink_anchor = None;
                    }
                } else if in_body && matches_local_name(event.name().as_ref(), b"tbl") {
                    table_depth = table_depth.saturating_sub(1);
                } else if matches_local_name(event.name().as_ref(), b"positionH") {
                    position_h_depth = position_h_depth.saturating_sub(1);
                    in_position_h_offset = false;
                } else if matches_local_name(event.name().as_ref(), b"positionV") {
                    position_v_depth = position_v_depth.saturating_sub(1);
                    in_position_v_offset = false;
                } else if matches_local_name(event.name().as_ref(), b"posOffset") {
                    in_position_h_offset = false;
                    in_position_v_offset = false;
                }

                if matches_local_name(event.name().as_ref(), b"pPr") {
                    paragraph_properties_depth = 0;
                }
                if matches_local_name(event.name().as_ref(), b"rPr") {
                    run_properties_depth = 0;
                }
                if drawing_depth > 0 {
                    drawing_depth = drawing_depth.saturating_sub(1);
                }
            }
            Event::Eof => break,
            _ => {}
        }
        buffer.clear();
    }

    Ok(paragraphs)
}

#[derive(Debug, Clone, PartialEq, Eq)]
struct ParsedTableCell {
    text: String,
    horizontal_span: usize,
    vertical_merge: Option<VerticalMerge>,
    shading_color: Option<String>,
    shading_color_attribute: Option<String>,
    shading_pattern: Option<String>,
    vertical_alignment: Option<VerticalAlignment>,
    cell_width_twips: Option<u32>,
    borders: CellBorders,
    margins: CellMargins,
    unknown_property_children: Vec<RawXmlNode>,
}

#[derive(Debug, Clone, PartialEq, Eq, Default)]
struct ParsedRowProperties {
    height_twips: Option<u32>,
    height_rule: Option<String>,
    repeat_header: bool,
}

fn parse_tables(xml: &[u8]) -> Result<Vec<Table>> {
    let mut reader = Reader::from_reader(Cursor::new(xml));
    reader.config_mut().trim_text(false);

    let mut tables = Vec::new();
    let mut in_body = false;
    let mut table_depth = 0_usize;
    let mut in_table_borders = false;
    let mut in_cell_borders = false;
    let mut in_cell_margins = false;
    let mut in_table_grid = false;
    let mut in_row_properties = false;
    let mut in_text = false;
    let mut rows: Vec<Vec<ParsedTableCell>> = Vec::new();
    let mut current_row: Option<Vec<ParsedTableCell>> = None;
    let mut current_cell_text: Option<String> = None;
    let mut current_cell_span = 1_usize;
    let mut current_cell_vertical_merge: Option<VerticalMerge> = None;
    let mut current_cell_shading_color: Option<String> = None;
    let mut current_cell_shading_color_attribute: Option<String> = None;
    let mut current_cell_shading_pattern: Option<String> = None;
    let mut current_cell_vertical_alignment: Option<VerticalAlignment> = None;
    let mut current_cell_width_twips: Option<u32> = None;
    let mut current_cell_borders = CellBorders::new();
    let mut current_cell_margins = CellMargins::new();
    let mut current_table_style_id: Option<String> = None;
    let mut current_table_borders = TableBorders::new();
    let mut current_table_column_widths: Vec<u32> = Vec::new();
    let mut current_row_properties: Vec<ParsedRowProperties> = Vec::new();
    let mut current_parsed_row_props = ParsedRowProperties::default();
    let mut current_table_alignment: Option<TableAlignment> = None;
    let mut current_table_width_twips: Option<u32> = None;
    let mut current_table_width_type: Option<TableWidthType> = None;
    let mut current_table_layout: Option<TableLayout> = None;
    let mut current_table_first_row = false;
    let mut current_table_last_row = false;
    let mut current_table_first_column = false;
    let mut current_table_last_column = false;
    let mut current_table_no_h_band = false;
    let mut current_table_no_v_band = false;
    let mut in_table_properties = false;
    let mut in_cell_properties = false;
    let mut current_cell_unknown_property_children: Vec<RawXmlNode> = Vec::new();
    let mut current_table_unknown_property_children: Vec<RawXmlNode> = Vec::new();
    let mut buffer = Vec::new();

    const KNOWN_TABLE_PROPERTY_CHILDREN: &[&[u8]] = &[
        b"tblStyle",
        b"tblW",
        b"jc",
        b"tblBorders",
        b"tblLayout",
        b"tblLook",
    ];
    const KNOWN_CELL_PROPERTY_CHILDREN: &[&[u8]] = &[
        b"tcW",
        b"gridSpan",
        b"vMerge",
        b"shd",
        b"vAlign",
        b"tcBorders",
        b"tcMar",
    ];

    loop {
        match reader.read_event_into(&mut buffer)? {
            Event::Start(ref event) => {
                if matches_local_name(event.name().as_ref(), b"body") {
                    in_body = true;
                } else if in_body && matches_local_name(event.name().as_ref(), b"tbl") {
                    table_depth = table_depth.saturating_add(1);
                    if table_depth == 1 {
                        rows.clear();
                        current_row = None;
                        current_cell_text = None;
                        current_cell_span = 1;
                        current_table_style_id = None;
                        current_table_borders.clear();
                        current_table_column_widths.clear();
                        current_row_properties.clear();
                        current_table_alignment = None;
                        current_table_width_twips = None;
                        current_table_width_type = None;
                        current_table_layout = None;
                        current_table_first_row = false;
                        current_table_last_row = false;
                        current_table_first_column = false;
                        current_table_last_column = false;
                        current_table_no_h_band = false;
                        current_table_no_v_band = false;
                        in_table_borders = false;
                        in_cell_borders = false;
                        in_cell_margins = false;
                        in_table_grid = false;
                        in_row_properties = false;
                        in_table_properties = false;
                        in_cell_properties = false;
                        in_text = false;
                        current_table_unknown_property_children.clear();
                    }
                } else if table_depth == 1 && matches_local_name(event.name().as_ref(), b"tr") {
                    current_row = Some(Vec::new());
                    current_parsed_row_props = ParsedRowProperties::default();
                    in_row_properties = false;
                } else if table_depth == 1 && matches_local_name(event.name().as_ref(), b"tc") {
                    current_cell_text = Some(String::new());
                    current_cell_span = 1;
                    current_cell_vertical_merge = None;
                    current_cell_shading_color = None;
                    current_cell_shading_color_attribute = None;
                    current_cell_shading_pattern = None;
                    current_cell_vertical_alignment = None;
                    current_cell_width_twips = None;
                    current_cell_borders = CellBorders::new();
                    current_cell_margins = CellMargins::new();
                    current_cell_unknown_property_children.clear();
                    in_cell_borders = false;
                    in_cell_margins = false;
                    in_cell_properties = false;
                } else if table_depth == 1
                    && current_row.is_some()
                    && current_cell_text.is_some()
                    && matches_local_name(event.name().as_ref(), b"tcPr")
                {
                    in_cell_properties = true;
                } else if table_depth == 1
                    && current_row.is_none()
                    && matches_local_name(event.name().as_ref(), b"tblPr")
                {
                    in_table_properties = true;
                } else if table_depth == 1 && matches_local_name(event.name().as_ref(), b"tblGrid")
                {
                    in_table_grid = true;
                } else if table_depth == 1
                    && in_table_grid
                    && matches_local_name(event.name().as_ref(), b"gridCol")
                {
                    if let Some(width) = parse_u32_attribute_value(event, b"w") {
                        current_table_column_widths.push(width);
                    }
                } else if table_depth == 1
                    && current_row.is_some()
                    && current_cell_text.is_none()
                    && matches_local_name(event.name().as_ref(), b"trPr")
                {
                    in_row_properties = true;
                } else if table_depth == 1
                    && in_row_properties
                    && matches_local_name(event.name().as_ref(), b"tblHeader")
                {
                    current_parsed_row_props.repeat_header = parse_on_off_property(event, true);
                } else if table_depth == 1
                    && in_row_properties
                    && matches_local_name(event.name().as_ref(), b"trHeight")
                {
                    current_parsed_row_props.height_twips =
                        parse_u32_attribute_value(event, b"val");
                    current_parsed_row_props.height_rule = parse_attribute_value(event, b"hRule");
                } else if table_depth == 1
                    && current_cell_text.is_some()
                    && matches_local_name(event.name().as_ref(), b"tcBorders")
                {
                    in_cell_borders = true;
                } else if table_depth == 1 && in_cell_borders {
                    maybe_apply_cell_border_edge(&mut current_cell_borders, event);
                } else if table_depth == 1
                    && current_cell_text.is_some()
                    && matches_local_name(event.name().as_ref(), b"tcMar")
                {
                    in_cell_margins = true;
                } else if table_depth == 1 && in_cell_margins {
                    maybe_apply_cell_margin_edge(&mut current_cell_margins, event);
                } else if table_depth == 1
                    && current_cell_text.is_some()
                    && matches_local_name(event.name().as_ref(), b"t")
                {
                    in_text = true;
                } else if table_depth == 1
                    && matches_local_name(event.name().as_ref(), b"tblBorders")
                {
                    in_table_borders = true;
                } else if table_depth == 1 && matches_local_name(event.name().as_ref(), b"tblStyle")
                {
                    maybe_apply_table_style_id(&mut current_table_style_id, event);
                } else if table_depth == 1
                    && current_cell_text.is_none()
                    && matches_local_name(event.name().as_ref(), b"jc")
                {
                    if let Some(val) = parse_attribute_value(event, b"val") {
                        current_table_alignment = match val.as_str() {
                            "start" | "left" => Some(TableAlignment::Left),
                            "center" => Some(TableAlignment::Center),
                            "end" | "right" => Some(TableAlignment::Right),
                            _ => None,
                        };
                    }
                } else if table_depth == 1
                    && current_cell_text.is_none()
                    && matches_local_name(event.name().as_ref(), b"tblW")
                {
                    current_table_width_twips = parse_u32_attribute_value(event, b"w");
                    if let Some(type_val) = parse_attribute_value(event, b"type") {
                        current_table_width_type = match type_val.as_str() {
                            "dxa" => Some(TableWidthType::Dxa),
                            "pct" => Some(TableWidthType::Pct),
                            "auto" => Some(TableWidthType::Auto),
                            _ => None,
                        };
                    }
                } else if table_depth == 1
                    && current_cell_text.is_none()
                    && matches_local_name(event.name().as_ref(), b"tblLayout")
                {
                    if let Some(type_val) = parse_attribute_value(event, b"type") {
                        current_table_layout = match type_val.as_str() {
                            "fixed" => Some(TableLayout::Fixed),
                            "autofit" => Some(TableLayout::AutoFit),
                            _ => None,
                        };
                    }
                } else if table_depth == 1
                    && current_cell_text.is_none()
                    && matches_local_name(event.name().as_ref(), b"tblLook")
                {
                    current_table_first_row =
                        parse_attribute_value(event, b"firstRow").as_deref() == Some("1");
                    current_table_last_row =
                        parse_attribute_value(event, b"lastRow").as_deref() == Some("1");
                    current_table_first_column =
                        parse_attribute_value(event, b"firstColumn").as_deref() == Some("1");
                    current_table_last_column =
                        parse_attribute_value(event, b"lastColumn").as_deref() == Some("1");
                    current_table_no_h_band =
                        parse_attribute_value(event, b"noHBand").as_deref() == Some("1");
                    current_table_no_v_band =
                        parse_attribute_value(event, b"noVBand").as_deref() == Some("1");
                } else if table_depth == 1 && in_table_borders {
                    maybe_apply_table_border_edge(&mut current_table_borders, event);
                } else if table_depth == 1
                    && current_cell_text.is_some()
                    && matches_local_name(event.name().as_ref(), b"gridSpan")
                {
                    current_cell_span = parse_u32_attribute_value(event, b"val")
                        .and_then(|value| usize::try_from(value).ok())
                        .filter(|value| *value > 0)
                        .unwrap_or(1);
                } else if table_depth == 1
                    && current_cell_text.is_some()
                    && matches_local_name(event.name().as_ref(), b"vMerge")
                {
                    current_cell_vertical_merge = Some(parse_vertical_merge(event));
                } else if table_depth == 1
                    && current_cell_text.is_some()
                    && matches_local_name(event.name().as_ref(), b"shd")
                {
                    current_cell_shading_color = parse_shading_fill_color(event);
                    current_cell_shading_color_attribute = parse_shading_color_attribute(event);
                    current_cell_shading_pattern = parse_shading_pattern(event);
                } else if table_depth == 1
                    && current_cell_text.is_some()
                    && matches_local_name(event.name().as_ref(), b"vAlign")
                {
                    current_cell_vertical_alignment = parse_vertical_alignment(event);
                } else if table_depth == 1
                    && current_cell_text.is_some()
                    && matches_local_name(event.name().as_ref(), b"tcW")
                {
                    current_cell_width_twips = parse_u32_attribute_value(event, b"w");
                } else if table_depth == 1 && in_cell_properties {
                    let name_bytes = event.name();
                    let local = local_name(name_bytes.as_ref());
                    if !KNOWN_CELL_PROPERTY_CHILDREN.contains(&local) {
                        current_cell_unknown_property_children
                            .push(RawXmlNode::read_element(&mut reader, event)?);
                    }
                } else if table_depth == 1 && in_table_properties {
                    let name_bytes = event.name();
                    let local = local_name(name_bytes.as_ref());
                    if !KNOWN_TABLE_PROPERTY_CHILDREN.contains(&local) {
                        current_table_unknown_property_children
                            .push(RawXmlNode::read_element(&mut reader, event)?);
                    }
                }
            }
            Event::Empty(ref event) => {
                if table_depth == 1 && matches_local_name(event.name().as_ref(), b"tc") {
                    if let Some(row) = current_row.as_mut() {
                        row.push(ParsedTableCell {
                            text: String::new(),
                            horizontal_span: 1,
                            vertical_merge: None,
                            shading_color: None,
                            shading_color_attribute: None,
                            shading_pattern: None,
                            vertical_alignment: None,
                            cell_width_twips: None,
                            borders: CellBorders::new(),
                            margins: CellMargins::new(),
                            unknown_property_children: Vec::new(),
                        });
                    }
                } else if table_depth == 1 && matches_local_name(event.name().as_ref(), b"tr") {
                    current_row_properties.push(current_parsed_row_props.clone());
                    rows.push(Vec::new());
                } else if table_depth == 1
                    && in_table_grid
                    && matches_local_name(event.name().as_ref(), b"gridCol")
                {
                    if let Some(width) = parse_u32_attribute_value(event, b"w") {
                        current_table_column_widths.push(width);
                    }
                } else if table_depth == 1
                    && in_row_properties
                    && matches_local_name(event.name().as_ref(), b"tblHeader")
                {
                    current_parsed_row_props.repeat_header = parse_on_off_property(event, true);
                } else if table_depth == 1
                    && in_row_properties
                    && matches_local_name(event.name().as_ref(), b"trHeight")
                {
                    current_parsed_row_props.height_twips =
                        parse_u32_attribute_value(event, b"val");
                    current_parsed_row_props.height_rule = parse_attribute_value(event, b"hRule");
                } else if table_depth == 1 && in_cell_borders {
                    maybe_apply_cell_border_edge(&mut current_cell_borders, event);
                } else if table_depth == 1 && in_cell_margins {
                    maybe_apply_cell_margin_edge(&mut current_cell_margins, event);
                } else if table_depth == 1
                    && current_cell_text.is_some()
                    && matches_local_name(event.name().as_ref(), b"gridSpan")
                {
                    current_cell_span = parse_u32_attribute_value(event, b"val")
                        .and_then(|value| usize::try_from(value).ok())
                        .filter(|value| *value > 0)
                        .unwrap_or(1);
                } else if table_depth == 1
                    && current_cell_text.is_some()
                    && matches_local_name(event.name().as_ref(), b"vMerge")
                {
                    current_cell_vertical_merge = Some(parse_vertical_merge(event));
                } else if table_depth == 1
                    && current_cell_text.is_some()
                    && matches_local_name(event.name().as_ref(), b"shd")
                {
                    current_cell_shading_color = parse_shading_fill_color(event);
                    current_cell_shading_color_attribute = parse_shading_color_attribute(event);
                    current_cell_shading_pattern = parse_shading_pattern(event);
                } else if table_depth == 1
                    && current_cell_text.is_some()
                    && matches_local_name(event.name().as_ref(), b"vAlign")
                {
                    current_cell_vertical_alignment = parse_vertical_alignment(event);
                } else if table_depth == 1
                    && current_cell_text.is_some()
                    && matches_local_name(event.name().as_ref(), b"tcW")
                {
                    current_cell_width_twips = parse_u32_attribute_value(event, b"w");
                } else if table_depth == 1 && matches_local_name(event.name().as_ref(), b"tblStyle")
                {
                    maybe_apply_table_style_id(&mut current_table_style_id, event);
                } else if table_depth == 1
                    && current_cell_text.is_none()
                    && matches_local_name(event.name().as_ref(), b"jc")
                {
                    if let Some(val) = parse_attribute_value(event, b"val") {
                        current_table_alignment = match val.as_str() {
                            "start" | "left" => Some(TableAlignment::Left),
                            "center" => Some(TableAlignment::Center),
                            "end" | "right" => Some(TableAlignment::Right),
                            _ => None,
                        };
                    }
                } else if table_depth == 1
                    && current_cell_text.is_none()
                    && matches_local_name(event.name().as_ref(), b"tblW")
                {
                    current_table_width_twips = parse_u32_attribute_value(event, b"w");
                    if let Some(type_val) = parse_attribute_value(event, b"type") {
                        current_table_width_type = match type_val.as_str() {
                            "dxa" => Some(TableWidthType::Dxa),
                            "pct" => Some(TableWidthType::Pct),
                            "auto" => Some(TableWidthType::Auto),
                            _ => None,
                        };
                    }
                } else if table_depth == 1
                    && current_cell_text.is_none()
                    && matches_local_name(event.name().as_ref(), b"tblLayout")
                {
                    if let Some(type_val) = parse_attribute_value(event, b"type") {
                        current_table_layout = match type_val.as_str() {
                            "fixed" => Some(TableLayout::Fixed),
                            "autofit" => Some(TableLayout::AutoFit),
                            _ => None,
                        };
                    }
                } else if table_depth == 1
                    && current_cell_text.is_none()
                    && matches_local_name(event.name().as_ref(), b"tblLook")
                {
                    current_table_first_row =
                        parse_attribute_value(event, b"firstRow").as_deref() == Some("1");
                    current_table_last_row =
                        parse_attribute_value(event, b"lastRow").as_deref() == Some("1");
                    current_table_first_column =
                        parse_attribute_value(event, b"firstColumn").as_deref() == Some("1");
                    current_table_last_column =
                        parse_attribute_value(event, b"lastColumn").as_deref() == Some("1");
                    current_table_no_h_band =
                        parse_attribute_value(event, b"noHBand").as_deref() == Some("1");
                    current_table_no_v_band =
                        parse_attribute_value(event, b"noVBand").as_deref() == Some("1");
                } else if table_depth == 1 && in_table_borders {
                    maybe_apply_table_border_edge(&mut current_table_borders, event);
                } else if table_depth == 1 && in_cell_properties {
                    let name_bytes = event.name();
                    let local = local_name(name_bytes.as_ref());
                    if !KNOWN_CELL_PROPERTY_CHILDREN.contains(&local) {
                        current_cell_unknown_property_children
                            .push(RawXmlNode::from_empty_element(event));
                    }
                } else if table_depth == 1 && in_table_properties {
                    let name_bytes = event.name();
                    let local = local_name(name_bytes.as_ref());
                    if !KNOWN_TABLE_PROPERTY_CHILDREN.contains(&local) {
                        current_table_unknown_property_children
                            .push(RawXmlNode::from_empty_element(event));
                    }
                }
            }
            Event::Text(ref event) => {
                if in_text && current_cell_text.is_some() {
                    let text = event
                        .xml_content()
                        .map_err(quick_xml::Error::from)?
                        .into_owned();
                    if let Some(cell_text) = current_cell_text.as_mut() {
                        cell_text.push_str(text.as_str());
                    }
                }
            }
            Event::CData(ref event) => {
                if in_text && current_cell_text.is_some() {
                    let text = String::from_utf8_lossy(event.as_ref()).into_owned();
                    if let Some(cell_text) = current_cell_text.as_mut() {
                        cell_text.push_str(text.as_str());
                    }
                }
            }
            Event::End(ref event) => {
                if matches_local_name(event.name().as_ref(), b"body") {
                    in_body = false;
                } else if table_depth == 1 && matches_local_name(event.name().as_ref(), b"t") {
                    in_text = false;
                } else if table_depth == 1 && matches_local_name(event.name().as_ref(), b"tblGrid")
                {
                    in_table_grid = false;
                } else if table_depth == 1 && matches_local_name(event.name().as_ref(), b"trPr") {
                    in_row_properties = false;
                } else if table_depth == 1
                    && matches_local_name(event.name().as_ref(), b"tcBorders")
                {
                    in_cell_borders = false;
                } else if table_depth == 1 && matches_local_name(event.name().as_ref(), b"tcMar") {
                    in_cell_margins = false;
                } else if table_depth == 1 && matches_local_name(event.name().as_ref(), b"tcPr") {
                    in_cell_properties = false;
                } else if table_depth == 1 && matches_local_name(event.name().as_ref(), b"tblPr") {
                    in_table_properties = false;
                } else if table_depth == 1 && matches_local_name(event.name().as_ref(), b"tc") {
                    in_text = false;
                    in_cell_borders = false;
                    in_cell_margins = false;
                    in_cell_properties = false;
                    if let Some(cell_text) = current_cell_text.take() {
                        if let Some(row) = current_row.as_mut() {
                            row.push(ParsedTableCell {
                                text: cell_text,
                                horizontal_span: current_cell_span.max(1),
                                vertical_merge: current_cell_vertical_merge.take(),
                                shading_color: current_cell_shading_color.take(),
                                shading_color_attribute: current_cell_shading_color_attribute
                                    .take(),
                                shading_pattern: current_cell_shading_pattern.take(),
                                vertical_alignment: current_cell_vertical_alignment.take(),
                                cell_width_twips: current_cell_width_twips.take(),
                                borders: std::mem::take(&mut current_cell_borders),
                                margins: std::mem::take(&mut current_cell_margins),
                                unknown_property_children: std::mem::take(
                                    &mut current_cell_unknown_property_children,
                                ),
                            });
                        }
                    }
                    current_cell_span = 1;
                    current_cell_vertical_merge = None;
                    current_cell_shading_color = None;
                    current_cell_shading_color_attribute = None;
                    current_cell_shading_pattern = None;
                    current_cell_vertical_alignment = None;
                    current_cell_width_twips = None;
                    current_cell_borders = CellBorders::new();
                    current_cell_margins = CellMargins::new();
                    current_cell_unknown_property_children.clear();
                } else if table_depth == 1 && matches_local_name(event.name().as_ref(), b"tr") {
                    in_cell_borders = false;
                    in_cell_margins = false;
                    in_cell_properties = false;
                    if let Some(cell_text) = current_cell_text.take() {
                        if let Some(row) = current_row.as_mut() {
                            row.push(ParsedTableCell {
                                text: cell_text,
                                horizontal_span: current_cell_span.max(1),
                                vertical_merge: current_cell_vertical_merge.take(),
                                shading_color: current_cell_shading_color.take(),
                                shading_color_attribute: current_cell_shading_color_attribute
                                    .take(),
                                shading_pattern: current_cell_shading_pattern.take(),
                                vertical_alignment: current_cell_vertical_alignment.take(),
                                cell_width_twips: current_cell_width_twips.take(),
                                borders: std::mem::take(&mut current_cell_borders),
                                margins: std::mem::take(&mut current_cell_margins),
                                unknown_property_children: std::mem::take(
                                    &mut current_cell_unknown_property_children,
                                ),
                            });
                        }
                    }
                    if let Some(row) = current_row.take() {
                        rows.push(row);
                    }
                    current_row_properties.push(std::mem::take(&mut current_parsed_row_props));
                    current_cell_span = 1;
                    current_cell_vertical_merge = None;
                    current_cell_shading_color = None;
                    current_cell_shading_color_attribute = None;
                    current_cell_shading_pattern = None;
                    current_cell_vertical_alignment = None;
                    current_cell_width_twips = None;
                    current_cell_borders = CellBorders::new();
                    current_cell_margins = CellMargins::new();
                    current_cell_unknown_property_children.clear();
                    in_row_properties = false;
                } else if table_depth == 1
                    && matches_local_name(event.name().as_ref(), b"tblBorders")
                {
                    in_table_borders = false;
                } else if in_body && matches_local_name(event.name().as_ref(), b"tbl") {
                    if table_depth == 1 {
                        // in_cell_borders and in_cell_margins are reset by the
                        // outer scope that follows; assigning here is redundant.
                        if let Some(cell_text) = current_cell_text.take() {
                            if let Some(row) = current_row.as_mut() {
                                row.push(ParsedTableCell {
                                    text: cell_text,
                                    horizontal_span: current_cell_span.max(1),
                                    vertical_merge: current_cell_vertical_merge.take(),
                                    shading_color: current_cell_shading_color.take(),
                                    shading_color_attribute: current_cell_shading_color_attribute
                                        .take(),
                                    shading_pattern: current_cell_shading_pattern.take(),
                                    vertical_alignment: current_cell_vertical_alignment.take(),
                                    cell_width_twips: current_cell_width_twips.take(),
                                    borders: std::mem::take(&mut current_cell_borders),
                                    margins: std::mem::take(&mut current_cell_margins),
                                    unknown_property_children: std::mem::take(
                                        &mut current_cell_unknown_property_children,
                                    ),
                                });
                            }
                        }
                        if let Some(row) = current_row.take() {
                            rows.push(row);
                        }
                        tables.push(build_table_from_rows(
                            &rows,
                            current_table_style_id.as_deref(),
                            &current_table_borders,
                            &current_table_column_widths,
                            &current_row_properties,
                            current_table_alignment,
                            current_table_width_twips,
                            current_table_width_type,
                            current_table_layout,
                            current_table_first_row,
                            current_table_last_row,
                            current_table_first_column,
                            current_table_last_column,
                            current_table_no_h_band,
                            current_table_no_v_band,
                            &current_table_unknown_property_children,
                        ));
                    }
                    table_depth = table_depth.saturating_sub(1);
                    current_table_style_id = None;
                    current_table_alignment = None;
                    current_table_width_twips = None;
                    current_table_width_type = None;
                    current_table_layout = None;
                    current_table_first_row = false;
                    current_table_last_row = false;
                    current_table_first_column = false;
                    current_table_last_column = false;
                    current_table_no_h_band = false;
                    current_table_no_v_band = false;
                    in_table_borders = false;
                    in_table_grid = false;
                    in_row_properties = false;
                    in_cell_borders = false;
                    in_cell_margins = false;
                    in_text = false;
                    current_cell_span = 1;
                    current_cell_vertical_merge = None;
                    current_cell_shading_color = None;
                    current_cell_shading_color_attribute = None;
                    current_cell_shading_pattern = None;
                    current_cell_vertical_alignment = None;
                    current_cell_width_twips = None;
                    current_cell_borders = CellBorders::new();
                    current_cell_margins = CellMargins::new();
                }
            }
            Event::Eof => break,
            _ => {}
        }
        buffer.clear();
    }

    Ok(tables)
}

fn parse_section(
    package: &Package,
    document_part_uri: &PartUri,
    document_part: &Part,
    xml: &[u8],
) -> Result<Section> {
    let mut reader = Reader::from_reader(Cursor::new(xml));
    reader.config_mut().trim_text(false);

    let mut section = Section::new();
    let mut in_body = false;
    let mut table_depth = 0_usize;
    let mut section_depth = 0_usize;
    let mut section_relationship_ids = SectionRelationshipIds::default();
    let mut buffer = Vec::new();

    loop {
        match reader.read_event_into(&mut buffer)? {
            Event::Start(ref event) => {
                if matches_local_name(event.name().as_ref(), b"body") {
                    in_body = true;
                } else if in_body && matches_local_name(event.name().as_ref(), b"tbl") {
                    table_depth = table_depth.saturating_add(1);
                } else if in_body
                    && table_depth == 0
                    && matches_local_name(event.name().as_ref(), b"sectPr")
                {
                    section_depth = section_depth.saturating_add(1);
                } else if section_depth > 0 && matches_local_name(event.name().as_ref(), b"pgSz") {
                    maybe_apply_section_page_size(&mut section, event);
                } else if section_depth > 0 && matches_local_name(event.name().as_ref(), b"pgMar") {
                    maybe_apply_section_page_margins(&mut section, event);
                } else if section_depth > 0
                    && matches_local_name(event.name().as_ref(), b"headerReference")
                {
                    maybe_apply_section_reference(
                        &mut section_relationship_ids,
                        &mut section,
                        event,
                        true,
                    );
                } else if section_depth > 0
                    && matches_local_name(event.name().as_ref(), b"footerReference")
                {
                    maybe_apply_section_reference(
                        &mut section_relationship_ids,
                        &mut section,
                        event,
                        false,
                    );
                } else if section_depth > 0 && matches_local_name(event.name().as_ref(), b"titlePg")
                {
                    section.set_title_page(parse_on_off_property(event, true));
                } else if section_depth > 0 && matches_local_name(event.name().as_ref(), b"type") {
                    if let Some(value) = parse_attribute_value(event, b"val") {
                        section.set_break_type_option(SectionBreakType::from_xml_value(
                            value.as_str(),
                        ));
                    }
                } else if section_depth > 0
                    && matches_local_name(event.name().as_ref(), b"pgNumType")
                {
                    maybe_apply_section_page_num_type(&mut section, event);
                } else if section_depth > 0 && matches_local_name(event.name().as_ref(), b"cols") {
                    maybe_apply_section_columns(&mut section, event);
                } else if section_depth > 0 && matches_local_name(event.name().as_ref(), b"vAlign")
                {
                    if let Some(value) = parse_attribute_value(event, b"val") {
                        section.set_vertical_alignment(
                            SectionVerticalAlignment::from_xml_value(value.as_str())
                                .unwrap_or(SectionVerticalAlignment::Top),
                        );
                    }
                } else if section_depth > 0
                    && matches_local_name(event.name().as_ref(), b"lnNumType")
                {
                    maybe_apply_section_line_numbering(&mut section, event);
                }
            }
            Event::Empty(ref event) => {
                if in_body
                    && table_depth == 0
                    && matches_local_name(event.name().as_ref(), b"sectPr")
                {
                    section = Section::new();
                } else if section_depth > 0 && matches_local_name(event.name().as_ref(), b"pgSz") {
                    maybe_apply_section_page_size(&mut section, event);
                } else if section_depth > 0 && matches_local_name(event.name().as_ref(), b"pgMar") {
                    maybe_apply_section_page_margins(&mut section, event);
                } else if section_depth > 0
                    && matches_local_name(event.name().as_ref(), b"headerReference")
                {
                    maybe_apply_section_reference(
                        &mut section_relationship_ids,
                        &mut section,
                        event,
                        true,
                    );
                } else if section_depth > 0
                    && matches_local_name(event.name().as_ref(), b"footerReference")
                {
                    maybe_apply_section_reference(
                        &mut section_relationship_ids,
                        &mut section,
                        event,
                        false,
                    );
                } else if section_depth > 0 && matches_local_name(event.name().as_ref(), b"titlePg")
                {
                    section.set_title_page(parse_on_off_property(event, true));
                } else if section_depth > 0
                    && matches_local_name(event.name().as_ref(), b"pgNumType")
                {
                    maybe_apply_section_page_num_type(&mut section, event);
                } else if section_depth > 0 && matches_local_name(event.name().as_ref(), b"type") {
                    if let Some(value) = parse_attribute_value(event, b"val") {
                        section.set_break_type_option(SectionBreakType::from_xml_value(
                            value.as_str(),
                        ));
                    }
                } else if section_depth > 0 && matches_local_name(event.name().as_ref(), b"cols") {
                    maybe_apply_section_columns(&mut section, event);
                } else if section_depth > 0 && matches_local_name(event.name().as_ref(), b"vAlign")
                {
                    if let Some(value) = parse_attribute_value(event, b"val") {
                        section.set_vertical_alignment(
                            SectionVerticalAlignment::from_xml_value(value.as_str())
                                .unwrap_or(SectionVerticalAlignment::Top),
                        );
                    }
                } else if section_depth > 0
                    && matches_local_name(event.name().as_ref(), b"lnNumType")
                {
                    maybe_apply_section_line_numbering(&mut section, event);
                }
            }
            Event::End(ref event) => {
                if matches_local_name(event.name().as_ref(), b"body") {
                    in_body = false;
                    table_depth = 0;
                    section_depth = 0;
                } else if in_body && matches_local_name(event.name().as_ref(), b"tbl") {
                    table_depth = table_depth.saturating_sub(1);
                } else if matches_local_name(event.name().as_ref(), b"sectPr") {
                    section_depth = section_depth.saturating_sub(1);
                }
            }
            Event::Eof => break,
            _ => {}
        }
        buffer.clear();
    }

    section.set_header_option(load_header_footer_part(
        package,
        document_part_uri,
        document_part,
        section_relationship_ids.header_relationship_id.as_deref(),
        WORD_HEADER_REL_TYPE,
    )?);
    section.set_footer_option(load_header_footer_part(
        package,
        document_part_uri,
        document_part,
        section_relationship_ids.footer_relationship_id.as_deref(),
        WORD_FOOTER_REL_TYPE,
    )?);
    section.set_first_page_header_option(load_header_footer_part(
        package,
        document_part_uri,
        document_part,
        section_relationship_ids
            .first_page_header_relationship_id
            .as_deref(),
        WORD_HEADER_REL_TYPE,
    )?);
    section.set_first_page_footer_option(load_header_footer_part(
        package,
        document_part_uri,
        document_part,
        section_relationship_ids
            .first_page_footer_relationship_id
            .as_deref(),
        WORD_FOOTER_REL_TYPE,
    )?);
    section.set_even_page_header_option(load_header_footer_part(
        package,
        document_part_uri,
        document_part,
        section_relationship_ids
            .even_page_header_relationship_id
            .as_deref(),
        WORD_HEADER_REL_TYPE,
    )?);
    section.set_even_page_footer_option(load_header_footer_part(
        package,
        document_part_uri,
        document_part,
        section_relationship_ids
            .even_page_footer_relationship_id
            .as_deref(),
        WORD_FOOTER_REL_TYPE,
    )?);

    Ok(section)
}

fn parse_body_item_kinds(xml: &[u8]) -> Result<(Vec<ParsedBodyItemKind>, Vec<RawXmlNode>)> {
    let mut reader = Reader::from_reader(Cursor::new(xml));
    reader.config_mut().trim_text(false);

    let mut body_item_kinds = Vec::new();
    let mut unknown_children = Vec::new();
    let mut in_body = false;
    let mut table_depth = 0_usize;
    let mut buffer = Vec::new();

    /// Known top-level body element local names that we handle.
    fn is_known_body_element(name: &[u8]) -> bool {
        matches_local_name(name, b"p")
            || matches_local_name(name, b"tbl")
            || matches_local_name(name, b"sectPr")
    }

    loop {
        match reader.read_event_into(&mut buffer)? {
            Event::Start(ref event) => {
                if matches_local_name(event.name().as_ref(), b"body") {
                    in_body = true;
                } else if in_body && matches_local_name(event.name().as_ref(), b"tbl") {
                    if table_depth == 0 {
                        body_item_kinds.push(ParsedBodyItemKind::Table);
                    }
                    table_depth = table_depth.saturating_add(1);
                } else if in_body
                    && table_depth == 0
                    && matches_local_name(event.name().as_ref(), b"p")
                {
                    body_item_kinds.push(ParsedBodyItemKind::Paragraph);
                } else if in_body
                    && table_depth == 0
                    && !is_known_body_element(event.name().as_ref())
                {
                    // Skip Start-based unknown elements — they may contain
                    // `<w:p>` / `<w:tbl>` children that are already parsed
                    // by parse_paragraphs / parse_tables. Consume to avoid
                    // advancing the body_item_kinds sequence incorrectly.
                    let _ = RawXmlNode::read_element(&mut reader, event);
                }
            }
            Event::Empty(ref event) => {
                if in_body && table_depth == 0 && matches_local_name(event.name().as_ref(), b"p") {
                    body_item_kinds.push(ParsedBodyItemKind::Paragraph);
                } else if in_body
                    && table_depth == 0
                    && !is_known_body_element(event.name().as_ref())
                {
                    body_item_kinds.push(ParsedBodyItemKind::Unknown);
                    unknown_children.push(RawXmlNode::from_empty_element(event));
                }
            }
            Event::End(ref event) => {
                if matches_local_name(event.name().as_ref(), b"body") {
                    in_body = false;
                    table_depth = 0;
                } else if in_body && matches_local_name(event.name().as_ref(), b"tbl") {
                    table_depth = table_depth.saturating_sub(1);
                }
            }
            Event::Eof => break,
            _ => {}
        }
        buffer.clear();
    }

    Ok((body_item_kinds, unknown_children))
}

fn bind_body_item_refs(
    body_item_kinds: &[ParsedBodyItemKind],
    paragraph_count: usize,
    table_count: usize,
    unknown_count: usize,
) -> Vec<BodyItemRef> {
    let mut paragraph_index = 0_usize;
    let mut table_index = 0_usize;
    let mut unknown_index = 0_usize;
    let mut body_item_refs = Vec::with_capacity(
        paragraph_count
            .saturating_add(table_count)
            .saturating_add(unknown_count),
    );

    for item_kind in body_item_kinds {
        match item_kind {
            ParsedBodyItemKind::Paragraph if paragraph_index < paragraph_count => {
                body_item_refs.push(BodyItemRef::Paragraph(paragraph_index));
                paragraph_index = paragraph_index.saturating_add(1);
            }
            ParsedBodyItemKind::Table if table_index < table_count => {
                body_item_refs.push(BodyItemRef::Table(table_index));
                table_index = table_index.saturating_add(1);
            }
            ParsedBodyItemKind::Unknown if unknown_index < unknown_count => {
                body_item_refs.push(BodyItemRef::Unknown(unknown_index));
                unknown_index = unknown_index.saturating_add(1);
            }
            _ => {}
        }
    }

    while paragraph_index < paragraph_count {
        body_item_refs.push(BodyItemRef::Paragraph(paragraph_index));
        paragraph_index = paragraph_index.saturating_add(1);
    }
    while table_index < table_count {
        body_item_refs.push(BodyItemRef::Table(table_index));
        table_index = table_index.saturating_add(1);
    }

    body_item_refs
}

#[allow(clippy::too_many_arguments)]
fn build_table_from_rows(
    rows: &[Vec<ParsedTableCell>],
    style_id: Option<&str>,
    borders: &TableBorders,
    column_widths: &[u32],
    row_properties: &[ParsedRowProperties],
    alignment: Option<TableAlignment>,
    width_twips: Option<u32>,
    width_type: Option<TableWidthType>,
    layout: Option<TableLayout>,
    first_row: bool,
    last_row: bool,
    first_column: bool,
    last_column: bool,
    no_h_band: bool,
    no_v_band: bool,
    unknown_table_property_children: &[RawXmlNode],
) -> Table {
    let row_count = rows.len();
    let column_count = rows
        .iter()
        .map(|row| {
            row.iter()
                .map(|cell| cell.horizontal_span.max(1))
                .sum::<usize>()
        })
        .max()
        .unwrap_or(0);
    let mut table = Table::new(row_count, column_count);

    for (row_idx, row) in rows.iter().enumerate() {
        let mut column_idx = 0_usize;
        for cell in row {
            if column_idx >= column_count {
                break;
            }
            let _ = table.set_cell_text(row_idx, column_idx, cell.text.clone());
            let span = cell.horizontal_span.max(1).min(column_count - column_idx);
            if span > 1 {
                let _ = table.set_horizontal_span(row_idx, column_idx, span);
            }
            if let Some(table_cell) = table.cell_mut(row_idx, column_idx) {
                if let Some(vertical_merge) = cell.vertical_merge {
                    table_cell.set_vertical_merge(vertical_merge);
                }
                if let Some(ref shading_color) = cell.shading_color {
                    table_cell.set_shading_color(shading_color.clone());
                }
                if let Some(ref shading_color_attr) = cell.shading_color_attribute {
                    table_cell.set_shading_color_attribute(shading_color_attr.clone());
                }
                table_cell.set_shading_pattern_option(cell.shading_pattern.clone());
                if let Some(vertical_alignment) = cell.vertical_alignment {
                    table_cell.set_vertical_alignment(vertical_alignment);
                }
                if let Some(cell_width_twips) = cell.cell_width_twips {
                    table_cell.set_cell_width_twips(cell_width_twips);
                }
                if !cell.borders.is_empty() {
                    table_cell.set_borders(cell.borders.clone());
                }
                if !cell.margins.is_empty() {
                    table_cell.set_margins(cell.margins);
                }
                for node in &cell.unknown_property_children {
                    table_cell.push_unknown_property_child(node.clone());
                }
            }
            column_idx = column_idx.saturating_add(span);
        }
    }

    if !borders.is_empty() {
        table.set_borders(borders.clone());
    }
    table.set_style_id_option(style_id.map(str::to_string));
    if !column_widths.is_empty() {
        table.set_column_widths_twips(column_widths.to_vec());
    }
    for (row_idx, parsed_props) in row_properties.iter().enumerate() {
        if let Some(row_props) = table.row_properties_mut(row_idx) {
            if parsed_props.repeat_header {
                row_props.set_repeat_header(true);
            }
            if let Some(height) = parsed_props.height_twips {
                row_props.set_height_twips(height);
            }
            if let Some(ref rule) = parsed_props.height_rule {
                row_props.set_height_rule(rule.clone());
            }
        }
    }
    if let Some(alignment) = alignment {
        table.set_alignment(alignment);
    }
    if let Some(width) = width_twips {
        table.set_width_twips(width);
    }
    if let Some(wt) = width_type {
        table.set_width_type(wt);
    }
    if let Some(layout) = layout {
        table.set_layout(layout);
    }
    table.set_first_row(first_row);
    table.set_last_row(last_row);
    table.set_first_column(first_column);
    table.set_last_column(last_column);
    table.set_no_h_band(no_h_band);
    table.set_no_v_band(no_v_band);
    for node in unknown_table_property_children {
        table.push_unknown_property_child(node.clone());
    }

    table
}

fn build_document_relationships_and_media_parts(
    package: &mut Package,
    document_part_uri: &PartUri,
    paragraphs: &[Paragraph],
    images: &[Image],
    section: &Section,
    body: &[BodyItemRef],
) -> Result<DocumentRelationshipBuildResult> {
    let mut next_media_index = 1_u32;
    let paragraph_indices = paragraph_indices_in_body_order(paragraphs.len(), body);
    let ordered_paragraphs: Vec<&Paragraph> = paragraph_indices
        .iter()
        .filter_map(|paragraph_index| paragraphs.get(*paragraph_index))
        .collect();
    let (mut relationships, hyperlink_relationship_ids, image_relationship_ids) =
        build_part_relationships_and_media_parts(
            package,
            document_part_uri,
            &ordered_paragraphs,
            images,
            &mut next_media_index,
            ImageRelationshipInclusion::AllImages,
        )?;

    let mut section_relationship_ids = SectionRelationshipIds::default();
    if let Some(header) = section.header() {
        let header_part_uri = PartUri::new("/word/header1.xml")?;
        let header_paragraphs: Vec<&Paragraph> = header.paragraphs().iter().collect();
        let (
            header_relationships,
            header_hyperlink_relationship_ids,
            header_image_relationship_ids,
        ) = build_part_relationships_and_media_parts(
            package,
            &header_part_uri,
            &header_paragraphs,
            images,
            &mut next_media_index,
            ImageRelationshipInclusion::ReferencedInRuns,
        )?;
        let header_xml = serialize_header_footer_xml(
            "w:hdr",
            header.paragraphs(),
            header.tables(),
            &header_hyperlink_relationship_ids,
            &header_image_relationship_ids,
        )?;
        let mut header_part = Part::new_xml(header_part_uri.clone(), header_xml);
        header_part.content_type = Some(WORD_HEADER_CONTENT_TYPE.to_string());
        header_part.relationships = header_relationships;
        package.set_part(header_part);
        let header_relationship = relationships.add_new(
            WORD_HEADER_REL_TYPE.to_string(),
            relative_path_from_part(document_part_uri, &header_part_uri),
            TargetMode::Internal,
        );
        section_relationship_ids.header_relationship_id = Some(header_relationship.id.clone());
    }
    if let Some(footer) = section.footer() {
        let footer_part_uri = PartUri::new("/word/footer1.xml")?;
        let footer_paragraphs: Vec<&Paragraph> = footer.paragraphs().iter().collect();
        let (
            footer_relationships,
            footer_hyperlink_relationship_ids,
            footer_image_relationship_ids,
        ) = build_part_relationships_and_media_parts(
            package,
            &footer_part_uri,
            &footer_paragraphs,
            images,
            &mut next_media_index,
            ImageRelationshipInclusion::ReferencedInRuns,
        )?;
        let footer_xml = serialize_header_footer_xml(
            "w:ftr",
            footer.paragraphs(),
            footer.tables(),
            &footer_hyperlink_relationship_ids,
            &footer_image_relationship_ids,
        )?;
        let mut footer_part = Part::new_xml(footer_part_uri.clone(), footer_xml);
        footer_part.content_type = Some(WORD_FOOTER_CONTENT_TYPE.to_string());
        footer_part.relationships = footer_relationships;
        package.set_part(footer_part);
        let footer_relationship = relationships.add_new(
            WORD_FOOTER_REL_TYPE.to_string(),
            relative_path_from_part(document_part_uri, &footer_part_uri),
            TargetMode::Internal,
        );
        section_relationship_ids.footer_relationship_id = Some(footer_relationship.id.clone());
    }
    if let Some(first_header) = section.first_page_header() {
        let first_header_part_uri = PartUri::new("/word/header2.xml")?;
        let first_header_paragraphs: Vec<&Paragraph> = first_header.paragraphs().iter().collect();
        let (
            first_header_relationships,
            first_header_hyperlink_relationship_ids,
            first_header_image_relationship_ids,
        ) = build_part_relationships_and_media_parts(
            package,
            &first_header_part_uri,
            &first_header_paragraphs,
            images,
            &mut next_media_index,
            ImageRelationshipInclusion::ReferencedInRuns,
        )?;
        let first_header_xml = serialize_header_footer_xml(
            "w:hdr",
            first_header.paragraphs(),
            first_header.tables(),
            &first_header_hyperlink_relationship_ids,
            &first_header_image_relationship_ids,
        )?;
        let mut first_header_part = Part::new_xml(first_header_part_uri.clone(), first_header_xml);
        first_header_part.content_type = Some(WORD_HEADER_CONTENT_TYPE.to_string());
        first_header_part.relationships = first_header_relationships;
        package.set_part(first_header_part);
        let first_header_relationship = relationships.add_new(
            WORD_HEADER_REL_TYPE.to_string(),
            relative_path_from_part(document_part_uri, &first_header_part_uri),
            TargetMode::Internal,
        );
        section_relationship_ids.first_page_header_relationship_id =
            Some(first_header_relationship.id.clone());
    }
    if let Some(first_footer) = section.first_page_footer() {
        let first_footer_part_uri = PartUri::new("/word/footer2.xml")?;
        let first_footer_paragraphs: Vec<&Paragraph> = first_footer.paragraphs().iter().collect();
        let (
            first_footer_relationships,
            first_footer_hyperlink_relationship_ids,
            first_footer_image_relationship_ids,
        ) = build_part_relationships_and_media_parts(
            package,
            &first_footer_part_uri,
            &first_footer_paragraphs,
            images,
            &mut next_media_index,
            ImageRelationshipInclusion::ReferencedInRuns,
        )?;
        let first_footer_xml = serialize_header_footer_xml(
            "w:ftr",
            first_footer.paragraphs(),
            first_footer.tables(),
            &first_footer_hyperlink_relationship_ids,
            &first_footer_image_relationship_ids,
        )?;
        let mut first_footer_part = Part::new_xml(first_footer_part_uri.clone(), first_footer_xml);
        first_footer_part.content_type = Some(WORD_FOOTER_CONTENT_TYPE.to_string());
        first_footer_part.relationships = first_footer_relationships;
        package.set_part(first_footer_part);
        let first_footer_relationship = relationships.add_new(
            WORD_FOOTER_REL_TYPE.to_string(),
            relative_path_from_part(document_part_uri, &first_footer_part_uri),
            TargetMode::Internal,
        );
        section_relationship_ids.first_page_footer_relationship_id =
            Some(first_footer_relationship.id.clone());
    }
    if let Some(even_header) = section.even_page_header() {
        let even_header_part_uri = PartUri::new("/word/header3.xml")?;
        let even_header_paragraphs: Vec<&Paragraph> = even_header.paragraphs().iter().collect();
        let (
            even_header_relationships,
            even_header_hyperlink_relationship_ids,
            even_header_image_relationship_ids,
        ) = build_part_relationships_and_media_parts(
            package,
            &even_header_part_uri,
            &even_header_paragraphs,
            images,
            &mut next_media_index,
            ImageRelationshipInclusion::ReferencedInRuns,
        )?;
        let even_header_xml = serialize_header_footer_xml(
            "w:hdr",
            even_header.paragraphs(),
            even_header.tables(),
            &even_header_hyperlink_relationship_ids,
            &even_header_image_relationship_ids,
        )?;
        let mut even_header_part = Part::new_xml(even_header_part_uri.clone(), even_header_xml);
        even_header_part.content_type = Some(WORD_HEADER_CONTENT_TYPE.to_string());
        even_header_part.relationships = even_header_relationships;
        package.set_part(even_header_part);
        let even_header_relationship = relationships.add_new(
            WORD_HEADER_REL_TYPE.to_string(),
            relative_path_from_part(document_part_uri, &even_header_part_uri),
            TargetMode::Internal,
        );
        section_relationship_ids.even_page_header_relationship_id =
            Some(even_header_relationship.id.clone());
    }
    if let Some(even_footer) = section.even_page_footer() {
        let even_footer_part_uri = PartUri::new("/word/footer3.xml")?;
        let even_footer_paragraphs: Vec<&Paragraph> = even_footer.paragraphs().iter().collect();
        let (
            even_footer_relationships,
            even_footer_hyperlink_relationship_ids,
            even_footer_image_relationship_ids,
        ) = build_part_relationships_and_media_parts(
            package,
            &even_footer_part_uri,
            &even_footer_paragraphs,
            images,
            &mut next_media_index,
            ImageRelationshipInclusion::ReferencedInRuns,
        )?;
        let even_footer_xml = serialize_header_footer_xml(
            "w:ftr",
            even_footer.paragraphs(),
            even_footer.tables(),
            &even_footer_hyperlink_relationship_ids,
            &even_footer_image_relationship_ids,
        )?;
        let mut even_footer_part = Part::new_xml(even_footer_part_uri.clone(), even_footer_xml);
        even_footer_part.content_type = Some(WORD_FOOTER_CONTENT_TYPE.to_string());
        even_footer_part.relationships = even_footer_relationships;
        package.set_part(even_footer_part);
        let even_footer_relationship = relationships.add_new(
            WORD_FOOTER_REL_TYPE.to_string(),
            relative_path_from_part(document_part_uri, &even_footer_part_uri),
            TargetMode::Internal,
        );
        section_relationship_ids.even_page_footer_relationship_id =
            Some(even_footer_relationship.id.clone());
    }

    Ok((
        relationships,
        hyperlink_relationship_ids,
        image_relationship_ids,
        section_relationship_ids,
    ))
}

fn paragraph_indices_in_body_order(paragraph_count: usize, body: &[BodyItemRef]) -> Vec<usize> {
    if body.is_empty() {
        return (0..paragraph_count).collect();
    }

    let mut indices = Vec::new();
    for body_item in body {
        if let BodyItemRef::Paragraph(index) = body_item {
            if *index < paragraph_count {
                indices.push(*index);
            }
        }
    }

    indices
}

fn collect_run_hyperlink_relationship(
    run: &Run,
    relationships: &mut Relationships,
    relationship_ids: &mut HashMap<String, String>,
    seen_targets: &mut HashSet<String>,
) {
    let Some(target) = run.hyperlink() else {
        return;
    };
    let target = target.trim();
    if target.is_empty() {
        return;
    }

    let target_string = target.to_string();
    if !seen_targets.insert(target_string.clone()) {
        return;
    }

    let relationship_id = relationships
        .add_new(
            RelationshipType::HYPERLINK.to_string(),
            target_string.clone(),
            TargetMode::External,
        )
        .id
        .clone();
    relationship_ids.insert(target_string, relationship_id);
}

fn build_part_relationships_and_media_parts(
    package: &mut Package,
    part_uri: &PartUri,
    paragraphs: &[&Paragraph],
    images: &[Image],
    next_media_index: &mut u32,
    image_relationship_inclusion: ImageRelationshipInclusion,
) -> Result<PartRelationshipMaps> {
    let mut relationships = Relationships::new();
    let mut hyperlink_relationship_ids = HashMap::new();
    let mut image_relationship_ids = HashMap::new();
    let mut seen_hyperlink_targets = HashSet::new();
    let mut seen_image_indexes = HashSet::new();

    for paragraph in paragraphs {
        for run in paragraph.runs() {
            collect_run_hyperlink_relationship(
                run,
                &mut relationships,
                &mut hyperlink_relationship_ids,
                &mut seen_hyperlink_targets,
            );
            if let Some(image_index) = run_image_index(run) {
                if image_index < images.len() {
                    seen_image_indexes.insert(image_index);
                }
            }
        }
    }

    let image_indexes = ordered_image_indexes_for_part(
        seen_image_indexes,
        images.len(),
        image_relationship_inclusion,
    );
    for image_index in image_indexes {
        collect_image_relationship_for_index(
            package,
            part_uri,
            images,
            image_index,
            &mut relationships,
            &mut image_relationship_ids,
            next_media_index,
        )?;
    }

    Ok((
        relationships,
        hyperlink_relationship_ids,
        image_relationship_ids,
    ))
}

fn ordered_image_indexes_for_part(
    seen_image_indexes: HashSet<usize>,
    image_count: usize,
    image_relationship_inclusion: ImageRelationshipInclusion,
) -> Vec<usize> {
    match image_relationship_inclusion {
        ImageRelationshipInclusion::AllImages => (0..image_count).collect(),
        ImageRelationshipInclusion::ReferencedInRuns => {
            let mut indexes = seen_image_indexes.into_iter().collect::<Vec<_>>();
            indexes.sort_unstable();
            indexes
        }
    }
}

fn serialize_header_footer_xml(
    root_name: &str,
    paragraphs: &[Paragraph],
    tables: &[Table],
    hyperlink_relationship_ids: &HashMap<String, String>,
    image_relationship_ids: &HashMap<usize, String>,
) -> Result<Vec<u8>> {
    let mut writer = Writer::new_with_indent(Vec::new(), b' ', 2);
    writer.write_event(Event::Decl(BytesDecl::new(
        "1.0",
        Some("UTF-8"),
        Some("yes"),
    )))?;

    let mut root = BytesStart::new(root_name);
    root.push_attribute(("xmlns:w", WORD_MAIN_NS));
    if !hyperlink_relationship_ids.is_empty() || !image_relationship_ids.is_empty() {
        root.push_attribute(("xmlns:r", WORD_REL_NS));
    }
    if !image_relationship_ids.is_empty() {
        root.push_attribute(("xmlns:wp", WORDPROCESSING_DRAWING_NS));
        root.push_attribute(("xmlns:a", DRAWINGML_NS));
        root.push_attribute(("xmlns:pic", DRAWINGML_PICTURE_NS));
    }
    writer.write_event(Event::Start(root))?;

    let mut next_drawing_id = 1_u32;
    if paragraphs.is_empty() && tables.is_empty() {
        writer.write_event(Event::Empty(BytesStart::new("w:p")))?;
    } else {
        for paragraph in paragraphs {
            write_paragraph_xml(
                &mut writer,
                paragraph,
                hyperlink_relationship_ids,
                image_relationship_ids,
                &mut next_drawing_id,
            )?;
        }
        for table in tables {
            write_table_xml(&mut writer, table)?;
        }
    }

    writer.write_event(Event::End(BytesEnd::new(root_name)))?;
    Ok(writer.into_inner())
}

fn collect_image_relationship_for_index(
    package: &mut Package,
    part_uri: &PartUri,
    images: &[Image],
    image_index: usize,
    relationships: &mut Relationships,
    relationship_ids: &mut HashMap<usize, String>,
    next_media_index: &mut u32,
) -> Result<()> {
    if relationship_ids.contains_key(&image_index) || image_index >= images.len() {
        return Ok(());
    }

    let image = &images[image_index];
    let media_index = *next_media_index;
    *next_media_index = next_media_index.checked_add(1).ok_or_else(|| {
        DocxError::UnsupportedPackage("media index overflow while serializing".to_string())
    })?;

    let extension = extension_for_content_type(image.content_type());
    let media_part_uri = PartUri::new(format!("/word/media/image{media_index}.{extension}"))?;
    let relationship_target = relative_path_from_part(part_uri, &media_part_uri);
    let relationship = relationships.add_new(
        RelationshipType::IMAGE.to_string(),
        relationship_target,
        TargetMode::Internal,
    );

    let mut media_part = Part::new(media_part_uri, image.bytes().to_vec());
    media_part.content_type = Some(image.content_type().to_string());
    package.set_part(media_part);

    relationship_ids.insert(image_index, relationship.id.clone());
    Ok(())
}

fn run_image_index(run: &Run) -> Option<usize> {
    run.inline_image()
        .map(InlineImage::image_index)
        .or_else(|| run.floating_image().map(FloatingImage::image_index))
}

fn load_document_images(
    package: &Package,
    document_part_uri: &PartUri,
    document_part: &Part,
) -> Result<(Vec<Image>, HashMap<String, usize>)> {
    let mut images = Vec::new();
    let mut image_indexes_by_relationship_id = HashMap::new();
    let mut image_indexes_by_part_uri = HashMap::new();

    for relationship in document_part.relationships.iter() {
        if relationship.rel_type != RelationshipType::IMAGE {
            continue;
        }
        if relationship.target_mode != TargetMode::Internal {
            continue;
        }

        let image_uri = document_part_uri.resolve_relative(relationship.target.as_str())?;
        let image_index =
            if let Some(image_index) = image_indexes_by_part_uri.get(image_uri.as_str()) {
                *image_index
            } else {
                let Some(image_part) = package.get_part(image_uri.as_str()) else {
                    tracing::warn!(
                        relationship_id = relationship.id.as_str(),
                        part_uri = image_uri.as_str(),
                        "missing image part for relationship; skipping image reference"
                    );
                    continue;
                };

                let content_type = image_part.content_type.clone().unwrap_or_else(|| {
                    fallback_content_type_for_extension(image_uri.extension()).to_string()
                });
                let image = Image::new(image_part.data.as_bytes().to_vec(), content_type);
                let image_index = images.len();
                images.push(image);
                image_indexes_by_part_uri.insert(image_uri.as_str().to_string(), image_index);
                image_index
            };

        image_indexes_by_relationship_id.insert(relationship.id.clone(), image_index);
    }

    Ok((images, image_indexes_by_relationship_id))
}

fn load_document_styles(
    package: &Package,
    document_part_uri: &PartUri,
    document_part: &Part,
) -> Result<StyleRegistry> {
    let Some(styles_relationship) = document_part.relationships.iter().find(|relationship| {
        relationship.rel_type == RelationshipType::STYLES
            && relationship.target_mode == TargetMode::Internal
    }) else {
        return Ok(StyleRegistry::new());
    };

    let styles_uri = document_part_uri.resolve_relative(styles_relationship.target.as_str())?;
    let Some(styles_part) = package.get_part(styles_uri.as_str()) else {
        return Ok(StyleRegistry::new());
    };

    parse_styles_xml(styles_part.data.as_bytes())
}

fn parse_styles_xml(xml: &[u8]) -> Result<StyleRegistry> {
    let mut reader = Reader::from_reader(Cursor::new(xml));
    reader.config_mut().trim_text(false);

    let mut registry = StyleRegistry::new();
    let mut current_style: Option<Style> = None;
    let mut buffer = Vec::new();

    loop {
        match reader.read_event_into(&mut buffer)? {
            Event::Start(ref event) => {
                if matches_local_name(event.name().as_ref(), b"style") {
                    current_style = parse_style_from_element(event);
                } else if let Some(style) = current_style.as_mut() {
                    if matches_local_name(event.name().as_ref(), b"name") {
                        if let Some(style_name) = parse_attribute_value(event, b"val") {
                            style.set_name(style_name);
                        }
                    } else if matches_local_name(event.name().as_ref(), b"basedOn") {
                        if let Some(based_on) = parse_attribute_value(event, b"val") {
                            style.set_based_on(based_on);
                        }
                    } else if matches_local_name(event.name().as_ref(), b"next") {
                        if let Some(next_style) = parse_attribute_value(event, b"val") {
                            style.set_next_style(next_style);
                        }
                    } else if matches_local_name(event.name().as_ref(), b"pPr") {
                        let snippet = capture_xml_subtree(&mut reader, event.to_owned())?;
                        style.set_paragraph_properties_xml(snippet);
                    } else if matches_local_name(event.name().as_ref(), b"rPr") {
                        let snippet = capture_xml_subtree(&mut reader, event.to_owned())?;
                        style.set_run_properties_xml(snippet);
                    } else if matches_local_name(event.name().as_ref(), b"tblPr") {
                        let snippet = capture_xml_subtree(&mut reader, event.to_owned())?;
                        style.set_table_properties_xml(snippet);
                    } else if matches_local_name(event.name().as_ref(), b"tblStylePr") {
                        let snippet = capture_xml_subtree(&mut reader, event.to_owned())?;
                        style.add_table_style_properties_xml(snippet);
                    }
                }
            }
            Event::Empty(ref event) => {
                if matches_local_name(event.name().as_ref(), b"style") {
                    if let Some(style) = parse_style_from_element(event) {
                        let _ = registry.add_style(style);
                    }
                } else if let Some(style) = current_style.as_mut() {
                    if matches_local_name(event.name().as_ref(), b"name") {
                        if let Some(style_name) = parse_attribute_value(event, b"val") {
                            style.set_name(style_name);
                        }
                    } else if matches_local_name(event.name().as_ref(), b"basedOn") {
                        if let Some(based_on) = parse_attribute_value(event, b"val") {
                            style.set_based_on(based_on);
                        }
                    } else if matches_local_name(event.name().as_ref(), b"next") {
                        if let Some(next_style) = parse_attribute_value(event, b"val") {
                            style.set_next_style(next_style);
                        }
                    } else if matches_local_name(event.name().as_ref(), b"pPr") {
                        let snippet =
                            serialize_xml_event_to_fragment(Event::Empty(event.to_owned()))?;
                        style.set_paragraph_properties_xml(snippet);
                    } else if matches_local_name(event.name().as_ref(), b"rPr") {
                        let snippet =
                            serialize_xml_event_to_fragment(Event::Empty(event.to_owned()))?;
                        style.set_run_properties_xml(snippet);
                    } else if matches_local_name(event.name().as_ref(), b"tblPr") {
                        let snippet =
                            serialize_xml_event_to_fragment(Event::Empty(event.to_owned()))?;
                        style.set_table_properties_xml(snippet);
                    } else if matches_local_name(event.name().as_ref(), b"tblStylePr") {
                        let snippet =
                            serialize_xml_event_to_fragment(Event::Empty(event.to_owned()))?;
                        style.add_table_style_properties_xml(snippet);
                    }
                }
            }
            Event::End(ref event) => {
                if matches_local_name(event.name().as_ref(), b"style") {
                    if let Some(style) = current_style.take() {
                        let _ = registry.add_style(style);
                    }
                }
            }
            Event::Eof => break,
            _ => {}
        }
        buffer.clear();
    }

    Ok(registry)
}

fn serialize_styles_xml(styles: &StyleRegistry) -> Result<Vec<u8>> {
    let mut writer = Writer::new_with_indent(Vec::new(), b' ', 2);
    writer.write_event(Event::Decl(BytesDecl::new(
        "1.0",
        Some("UTF-8"),
        Some("yes"),
    )))?;

    let mut styles_element = BytesStart::new("w:styles");
    styles_element.push_attribute(("xmlns:w", WORD_MAIN_NS));
    writer.write_event(Event::Start(styles_element))?;

    for style in styles.styles() {
        let mut style_element = BytesStart::new("w:style");
        style_element.push_attribute(("w:type", style.kind().to_xml_value()));
        style_element.push_attribute(("w:styleId", style.style_id()));
        writer.write_event(Event::Start(style_element))?;

        if let Some(name) = style.name() {
            let mut name_element = BytesStart::new("w:name");
            name_element.push_attribute(("w:val", name));
            writer.write_event(Event::Empty(name_element))?;
        }
        if let Some(based_on) = style.based_on() {
            let mut based_on_element = BytesStart::new("w:basedOn");
            based_on_element.push_attribute(("w:val", based_on));
            writer.write_event(Event::Empty(based_on_element))?;
        }
        if let Some(next_style) = style.next_style() {
            let mut next_element = BytesStart::new("w:next");
            next_element.push_attribute(("w:val", next_style));
            writer.write_event(Event::Empty(next_element))?;
        }

        if let Some(snippet) = style.paragraph_properties_xml() {
            write_xml_fragment(&mut writer, snippet)?;
        }
        if let Some(snippet) = style.run_properties_xml() {
            write_xml_fragment(&mut writer, snippet)?;
        }
        if let Some(snippet) = style.table_properties_xml() {
            write_xml_fragment(&mut writer, snippet)?;
        }
        for snippet in style.table_style_properties_xml() {
            write_xml_fragment(&mut writer, snippet)?;
        }

        writer.write_event(Event::End(BytesEnd::new("w:style")))?;
    }

    writer.write_event(Event::End(BytesEnd::new("w:styles")))?;
    Ok(writer.into_inner())
}

fn parse_style_from_element(event: &BytesStart<'_>) -> Option<Style> {
    let kind = parse_attribute_value(event, b"type")
        .and_then(|value| StyleKind::from_xml_value(value.to_ascii_lowercase().as_str()))?;
    let style_id = parse_attribute_value(event, b"styleId")?;
    let style_id = style_id.trim();
    if style_id.is_empty() {
        return None;
    }

    Some(Style::new(kind, style_id.to_string()))
}

fn capture_xml_subtree<R: BufRead>(
    reader: &mut Reader<R>,
    start: BytesStart<'_>,
) -> Result<String> {
    let mut writer = Writer::new(Vec::new());
    writer.write_event(Event::Start(start.into_owned()))?;
    let mut depth = 1_usize;
    let mut buffer = Vec::new();

    loop {
        let event = reader.read_event_into(&mut buffer)?;
        if matches!(event, Event::Eof) {
            return Err(DocxError::UnsupportedPackage(
                "unexpected EOF while parsing style subtree".to_string(),
            ));
        }
        if matches!(event, Event::Start(_)) {
            depth = depth.saturating_add(1);
        } else if matches!(event, Event::End(_)) {
            depth = depth.saturating_sub(1);
        }
        writer.write_event(event.into_owned())?;
        if depth == 0 {
            break;
        }
        buffer.clear();
    }

    Ok(String::from_utf8_lossy(writer.into_inner().as_slice()).into_owned())
}

fn serialize_xml_event_to_fragment(event: Event<'_>) -> Result<String> {
    let mut writer = Writer::new(Vec::new());
    writer.write_event(event.into_owned())?;
    Ok(String::from_utf8_lossy(writer.into_inner().as_slice()).into_owned())
}

fn write_xml_fragment(writer: &mut Writer<Vec<u8>>, snippet: &str) -> Result<()> {
    writer.get_mut().write_all(snippet.as_bytes())?;
    Ok(())
}

fn build_effective_styles_registry(
    styles: &StyleRegistry,
    paragraphs: &[Paragraph],
    tables: &[Table],
    section: &Section,
) -> StyleRegistry {
    let mut merged = styles.clone();

    // Always include the "Normal" paragraph style — Word expects styles.xml
    // to exist and contain at least this default style.
    let _ = merged.ensure_style(StyleKind::Paragraph, "Normal");

    for paragraph in paragraphs {
        if let Some(style_id) = paragraph.style_id() {
            let _ = merged.ensure_style(StyleKind::Paragraph, style_id);
        }
        for run in paragraph.runs() {
            if let Some(style_id) = run.style_id() {
                let _ = merged.ensure_style(StyleKind::Character, style_id);
            }
        }
    }
    for table in tables {
        if let Some(style_id) = table.style_id() {
            let _ = merged.ensure_style(StyleKind::Table, style_id);
        }
    }
    if let Some(header) = section.header() {
        for paragraph in header.paragraphs() {
            if let Some(style_id) = paragraph.style_id() {
                let _ = merged.ensure_style(StyleKind::Paragraph, style_id);
            }
            for run in paragraph.runs() {
                if let Some(style_id) = run.style_id() {
                    let _ = merged.ensure_style(StyleKind::Character, style_id);
                }
            }
        }
    }
    if let Some(footer) = section.footer() {
        for paragraph in footer.paragraphs() {
            if let Some(style_id) = paragraph.style_id() {
                let _ = merged.ensure_style(StyleKind::Paragraph, style_id);
            }
            for run in paragraph.runs() {
                if let Some(style_id) = run.style_id() {
                    let _ = merged.ensure_style(StyleKind::Character, style_id);
                }
            }
        }
    }
    if let Some(first_page_header) = section.first_page_header() {
        for paragraph in first_page_header.paragraphs() {
            if let Some(style_id) = paragraph.style_id() {
                let _ = merged.ensure_style(StyleKind::Paragraph, style_id);
            }
            for run in paragraph.runs() {
                if let Some(style_id) = run.style_id() {
                    let _ = merged.ensure_style(StyleKind::Character, style_id);
                }
            }
        }
    }
    if let Some(first_page_footer) = section.first_page_footer() {
        for paragraph in first_page_footer.paragraphs() {
            if let Some(style_id) = paragraph.style_id() {
                let _ = merged.ensure_style(StyleKind::Paragraph, style_id);
            }
            for run in paragraph.runs() {
                if let Some(style_id) = run.style_id() {
                    let _ = merged.ensure_style(StyleKind::Character, style_id);
                }
            }
        }
    }

    merged
}

fn extension_for_content_type(content_type: &str) -> &'static str {
    let normalized = content_type
        .split(';')
        .next()
        .map(str::trim)
        .unwrap_or_default()
        .to_ascii_lowercase();

    match normalized.as_str() {
        "image/png" => "png",
        "image/jpeg" => "jpeg",
        "image/jpg" => "jpg",
        "image/gif" => "gif",
        "image/bmp" => "bmp",
        "image/tiff" => "tiff",
        "image/svg+xml" => "svg",
        _ => "bin",
    }
}

fn fallback_content_type_for_extension(extension: Option<&str>) -> &'static str {
    match extension.unwrap_or_default().to_ascii_lowercase().as_str() {
        "png" => "image/png",
        "jpeg" | "jpg" => "image/jpeg",
        "gif" => "image/gif",
        "bmp" => "image/bmp",
        "tif" | "tiff" => "image/tiff",
        "svg" => "image/svg+xml",
        _ => OCTET_STREAM_CONTENT_TYPE,
    }
}

fn relative_path_from_part(from_part_uri: &PartUri, target_part_uri: &PartUri) -> String {
    let from_segments: Vec<&str> = from_part_uri
        .directory()
        .trim_start_matches('/')
        .split('/')
        .filter(|segment| !segment.is_empty())
        .collect();
    let target_segments: Vec<&str> = target_part_uri
        .as_str()
        .trim_start_matches('/')
        .split('/')
        .filter(|segment| !segment.is_empty())
        .collect();

    let mut common_length = 0_usize;
    while common_length < from_segments.len()
        && common_length < target_segments.len()
        && from_segments[common_length] == target_segments[common_length]
    {
        common_length = common_length.saturating_add(1);
    }

    let mut relative_segments = Vec::new();
    for _ in common_length..from_segments.len() {
        relative_segments.push("..".to_string());
    }
    for segment in target_segments.iter().skip(common_length) {
        relative_segments.push((*segment).to_string());
    }

    if relative_segments.is_empty() {
        ".".to_string()
    } else {
        relative_segments.join("/")
    }
}

fn parse_hyperlink_targets_by_relationship_id(document_part: &Part) -> HashMap<String, String> {
    let mut hyperlink_targets = HashMap::new();

    for relationship in document_part.relationships.iter() {
        if relationship.rel_type == RelationshipType::HYPERLINK
            && relationship.target_mode == TargetMode::External
        {
            hyperlink_targets.insert(relationship.id.clone(), relationship.target.clone());
        }
    }

    hyperlink_targets
}

fn is_rebuilt_document_relationship_type(rel_type: &str) -> bool {
    matches!(
        rel_type,
        RelationshipType::HYPERLINK
            | RelationshipType::IMAGE
            | RelationshipType::STYLES
            | WORD_HEADER_REL_TYPE
            | WORD_FOOTER_REL_TYPE
    )
}

fn record_passthrough_relationship(counts: &mut BTreeMap<String, usize>, rel_type: &str) {
    *counts.entry(rel_type.to_string()).or_default() += 1;
}

fn emit_passthrough_relationship_warnings(scope: &str, counts: &BTreeMap<String, usize>) {
    for (rel_type, count) in counts {
        tracing::warn!(
            scope = scope,
            relationship_type = rel_type.as_str(),
            count = *count,
            "pass-through preserving unsupported relationship type; editing not implemented yet"
        );
    }
}

#[derive(Debug, Default)]
struct CurrentRunProperties {
    bold: bool,
    italic: bool,
    underline_type: Option<UnderlineType>,
    strikethrough: bool,
    double_strikethrough: bool,
    subscript: bool,
    superscript: bool,
    small_caps: bool,
    all_caps: bool,
    hidden: bool,
    emboss: bool,
    imprint: bool,
    shadow: bool,
    outline: bool,
    character_spacing_twips: Option<i32>,
    highlight_color: Option<String>,
    style_id: Option<String>,
    font_family: Option<String>,
    font_family_ascii: Option<String>,
    font_family_h_ansi: Option<String>,
    font_family_cs: Option<String>,
    font_family_east_asia: Option<String>,
    font_size_half_points: Option<u16>,
    color: Option<String>,
    theme_color: Option<String>,
    theme_shade: Option<String>,
    theme_tint: Option<String>,
    hyperlink: Option<String>,
    hyperlink_tooltip: Option<String>,
    hyperlink_anchor: Option<String>,
    footnote_reference_id: Option<u32>,
    endnote_reference_id: Option<u32>,
    has_tab: bool,
    has_break: bool,
    image_relationship_id: Option<String>,
    image_width_emu: Option<u32>,
    image_height_emu: Option<u32>,
    image_name: Option<String>,
    image_description: Option<String>,
    drawing_kind: Option<DrawingKind>,
    rtl: bool,
    field_instruction: Option<String>,
    field_result: Option<String>,
    in_field: bool,
    field_separated: bool,
    floating_image_offset_x_emu: Option<i32>,
    floating_image_offset_y_emu: Option<i32>,
    unknown_children: Vec<RawXmlNode>,
    unknown_property_children: Vec<RawXmlNode>,
}

impl CurrentRunProperties {
    fn reset(
        &mut self,
        hyperlink: Option<String>,
        hyperlink_tooltip: Option<String>,
        hyperlink_anchor: Option<String>,
    ) {
        self.bold = false;
        self.italic = false;
        self.underline_type = None;
        self.strikethrough = false;
        self.double_strikethrough = false;
        self.subscript = false;
        self.superscript = false;
        self.small_caps = false;
        self.all_caps = false;
        self.hidden = false;
        self.emboss = false;
        self.imprint = false;
        self.shadow = false;
        self.outline = false;
        self.character_spacing_twips = None;
        self.highlight_color = None;
        self.style_id = None;
        self.font_family = None;
        self.font_family_ascii = None;
        self.font_family_h_ansi = None;
        self.font_family_cs = None;
        self.font_family_east_asia = None;
        self.font_size_half_points = None;
        self.color = None;
        self.theme_color = None;
        self.theme_shade = None;
        self.theme_tint = None;
        self.hyperlink = hyperlink;
        self.hyperlink_tooltip = hyperlink_tooltip;
        self.hyperlink_anchor = hyperlink_anchor;
        self.footnote_reference_id = None;
        self.endnote_reference_id = None;
        self.has_tab = false;
        self.has_break = false;
        self.rtl = false;
        self.field_instruction = None;
        self.field_result = None;
        self.in_field = false;
        self.field_separated = false;
        self.image_relationship_id = None;
        self.image_width_emu = None;
        self.image_height_emu = None;
        self.image_name = None;
        self.image_description = None;
        self.drawing_kind = None;
        self.floating_image_offset_x_emu = None;
        self.floating_image_offset_y_emu = None;
        self.unknown_children.clear();
        self.unknown_property_children.clear();
    }
}

fn finalize_current_run(
    paragraph: &mut Option<Paragraph>,
    current_run_text: &mut Option<String>,
    run_properties: &mut CurrentRunProperties,
    image_indexes_by_relationship_id: &HashMap<String, usize>,
) {
    let Some(run_text) = current_run_text.take() else {
        return;
    };
    let Some(paragraph) = paragraph.as_mut() else {
        return;
    };

    let run = paragraph.add_run(run_text);
    run.set_bold(run_properties.bold);
    run.set_italic(run_properties.italic);
    if let Some(ut) = run_properties.underline_type.take() {
        run.set_underline_type(ut);
    }
    run.set_strikethrough(run_properties.strikethrough);
    run.set_double_strikethrough(run_properties.double_strikethrough);
    run.set_subscript(run_properties.subscript);
    run.set_superscript(run_properties.superscript);
    run.set_small_caps(run_properties.small_caps);
    run.set_all_caps(run_properties.all_caps);
    run.set_hidden(run_properties.hidden);
    run.set_emboss(run_properties.emboss);
    run.set_imprint(run_properties.imprint);
    run.set_shadow(run_properties.shadow);
    run.set_outline(run_properties.outline);
    if let Some(spacing) = run_properties.character_spacing_twips.take() {
        run.set_character_spacing_twips(spacing);
    }
    if let Some(highlight_color) = run_properties.highlight_color.take() {
        run.set_highlight_color(highlight_color);
    }
    run.set_style_id_option(run_properties.style_id.take());
    if let Some(font_family) = run_properties.font_family.take() {
        run.set_font_family(font_family);
    }
    if let Some(ascii) = run_properties.font_family_ascii.take() {
        run.set_font_family_ascii(ascii);
    }
    if let Some(h_ansi) = run_properties.font_family_h_ansi.take() {
        run.set_font_family_h_ansi(h_ansi);
    }
    if let Some(cs) = run_properties.font_family_cs.take() {
        run.set_font_family_cs(cs);
    }
    if let Some(east_asia) = run_properties.font_family_east_asia.take() {
        run.set_font_family_east_asia(east_asia);
    }
    if let Some(font_size_half_points) = run_properties.font_size_half_points.take() {
        run.set_font_size_half_points(font_size_half_points);
    }
    if let Some(color) = run_properties.color.take() {
        run.set_color(color);
    }
    if let Some(theme_color) = run_properties.theme_color.take() {
        run.set_theme_color(theme_color);
    }
    if let Some(theme_shade) = run_properties.theme_shade.take() {
        run.set_theme_shade(theme_shade);
    }
    if let Some(theme_tint) = run_properties.theme_tint.take() {
        run.set_theme_tint(theme_tint);
    }
    if let Some(hyperlink) = run_properties.hyperlink.take() {
        run.set_hyperlink(hyperlink);
    }
    if let Some(tooltip) = run_properties.hyperlink_tooltip.take() {
        run.set_hyperlink_tooltip(tooltip);
    }
    if let Some(anchor) = run_properties.hyperlink_anchor.take() {
        run.set_hyperlink_anchor(anchor);
    }
    if let Some(footnote_id) = run_properties.footnote_reference_id.take() {
        run.set_footnote_reference_id(footnote_id);
    }
    if let Some(endnote_id) = run_properties.endnote_reference_id.take() {
        run.set_endnote_reference_id(endnote_id);
    }
    run.set_rtl(run_properties.rtl);
    if let Some(instruction) = run_properties.field_instruction.take() {
        let result = run_properties.field_result.take().unwrap_or_default();
        run.set_field_code(FieldCode::new(instruction, result));
    }
    if run_properties.has_tab {
        run.set_has_tab(true);
    }
    if run_properties.has_break {
        run.set_has_break(true);
    }
    for node in run_properties.unknown_property_children.drain(..) {
        run.push_unknown_property_child(node);
    }
    if let Some(relationship_id) = run_properties.image_relationship_id.take() {
        if let Some(image_index) = image_indexes_by_relationship_id.get(relationship_id.as_str()) {
            let image_width = run_properties.image_width_emu.unwrap_or_default();
            let image_height = run_properties.image_height_emu.unwrap_or_default();
            let image_name = run_properties.image_name.take();
            let image_description = run_properties.image_description.take();
            match run_properties.drawing_kind.unwrap_or(DrawingKind::Inline) {
                DrawingKind::Inline => {
                    let mut inline_image =
                        InlineImage::new(*image_index, image_width, image_height);
                    if let Some(name) = image_name {
                        inline_image.set_name(name);
                    }
                    if let Some(description) = image_description {
                        inline_image.set_description(description);
                    }
                    run.set_inline_image(inline_image);
                }
                DrawingKind::Anchor => {
                    let mut floating_image =
                        FloatingImage::new(*image_index, image_width, image_height);
                    floating_image.set_offsets_emu(
                        run_properties
                            .floating_image_offset_x_emu
                            .unwrap_or_default(),
                        run_properties
                            .floating_image_offset_y_emu
                            .unwrap_or_default(),
                    );
                    if let Some(name) = image_name {
                        floating_image.set_name(name);
                    }
                    if let Some(description) = image_description {
                        floating_image.set_description(description);
                    }
                    run.set_floating_image(floating_image);
                }
            }
        }
    }
    for node in run_properties.unknown_children.drain(..) {
        run.push_unknown_child(node);
    }
}

fn maybe_apply_paragraph_alignment(paragraph: &mut Option<Paragraph>, event: &BytesStart<'_>) {
    let Some(paragraph) = paragraph.as_mut() else {
        return;
    };

    if let Some(value) = parse_attribute_value(event, b"val") {
        let alignment_value = value.to_ascii_lowercase();
        if let Some(alignment) = ParagraphAlignment::from_xml_value(alignment_value.as_str()) {
            paragraph.set_alignment(alignment);
        }
    }
}

fn maybe_apply_paragraph_spacing(paragraph: &mut Option<Paragraph>, event: &BytesStart<'_>) {
    let Some(paragraph) = paragraph.as_mut() else {
        return;
    };

    if let Some(value) = parse_u32_attribute_value(event, b"before") {
        paragraph.set_spacing_before_twips(value);
    }
    if let Some(value) = parse_u32_attribute_value(event, b"after") {
        paragraph.set_spacing_after_twips(value);
    }
    if let Some(value) = parse_u32_attribute_value(event, b"line") {
        paragraph.set_line_spacing_twips(value);
    }
    if let Some(value) = parse_attribute_value(event, b"lineRule") {
        if let Some(rule) = LineSpacingRule::from_xml_value(value.as_str()) {
            paragraph.set_line_spacing_rule(rule);
        }
    }
    if let Some(value) = parse_attribute_value(event, b"beforeAutospacing") {
        let auto = !matches!(value.as_str(), "0" | "false" | "off" | "no");
        paragraph.set_before_autospacing(auto);
    }
    if let Some(value) = parse_attribute_value(event, b"afterAutospacing") {
        let auto = !matches!(value.as_str(), "0" | "false" | "off" | "no");
        paragraph.set_after_autospacing(auto);
    }
}

fn maybe_apply_paragraph_indentation(paragraph: &mut Option<Paragraph>, event: &BytesStart<'_>) {
    let Some(paragraph) = paragraph.as_mut() else {
        return;
    };

    if let Some(value) = parse_u32_attribute_value(event, b"left") {
        paragraph.set_indent_left_twips(value);
    }
    if let Some(value) = parse_u32_attribute_value(event, b"right") {
        paragraph.set_indent_right_twips(value);
    }
    if let Some(value) = parse_u32_attribute_value(event, b"firstLine") {
        paragraph.set_indent_first_line_twips(value);
    }
    if let Some(value) = parse_u32_attribute_value(event, b"hanging") {
        paragraph.set_indent_hanging_twips(value);
    }
}

fn maybe_apply_paragraph_numbering_num_id(
    paragraph: &mut Option<Paragraph>,
    event: &BytesStart<'_>,
) {
    let Some(paragraph) = paragraph.as_mut() else {
        return;
    };

    paragraph.set_numbering_num_id(parse_u32_attribute_value(event, b"val"));
}

fn maybe_apply_paragraph_numbering_ilvl(paragraph: &mut Option<Paragraph>, event: &BytesStart<'_>) {
    let Some(paragraph) = paragraph.as_mut() else {
        return;
    };

    let ilvl = parse_u32_attribute_value(event, b"val").and_then(|value| u8::try_from(value).ok());
    paragraph.set_numbering_ilvl(ilvl);
}

fn maybe_apply_paragraph_tab_stop(paragraph: &mut Option<Paragraph>, event: &BytesStart<'_>) {
    let Some(paragraph) = paragraph.as_mut() else {
        return;
    };

    let alignment = parse_attribute_value(event, b"val")
        .and_then(|value| TabStopAlignment::from_xml_value(value.to_ascii_lowercase().as_str()))
        .unwrap_or(TabStopAlignment::Left);
    let position = parse_u32_attribute_value(event, b"pos").unwrap_or(0);
    let leader = parse_attribute_value(event, b"leader")
        .and_then(|value| TabStopLeader::from_xml_value(value.to_ascii_lowercase().as_str()));

    let mut tab_stop = TabStop::new(position, alignment);
    if let Some(leader) = leader {
        tab_stop.set_leader(leader);
    }
    // Check for numTab attribute (non-standard extension, some Word documents use this)
    if let Some(val) = parse_attribute_value(event, b"numTab") {
        if val == "1" || val.eq_ignore_ascii_case("true") || val.eq_ignore_ascii_case("on") {
            tab_stop.set_num_tab(true);
        }
    }
    paragraph.add_tab_stop(tab_stop);
}

fn maybe_apply_paragraph_border_edge(paragraph: &mut Option<Paragraph>, event: &BytesStart<'_>) {
    let Some(paragraph) = paragraph.as_mut() else {
        return;
    };

    let mut border = ParagraphBorder::default();
    border.set_line_type_option(parse_attribute_value(event, b"val"));
    if let Some(size) =
        parse_u32_attribute_value(event, b"sz").and_then(|value| u16::try_from(value).ok())
    {
        border.set_size_eighth_points(size);
    }
    if let Some(color) = parse_attribute_value(event, b"color") {
        border.set_color(color);
    }
    if let Some(space) = parse_u32_attribute_value(event, b"space") {
        border.set_space_points(space);
    }

    if border.line_type().is_none()
        && border.size_eighth_points().is_none()
        && border.color().is_none()
        && border.space_points().is_none()
    {
        return;
    }

    if matches_local_name(event.name().as_ref(), b"top") {
        paragraph.borders_mut().set_top(border);
    } else if matches_local_name(event.name().as_ref(), b"left") {
        paragraph.borders_mut().set_left(border);
    } else if matches_local_name(event.name().as_ref(), b"bottom") {
        paragraph.borders_mut().set_bottom(border);
    } else if matches_local_name(event.name().as_ref(), b"right") {
        paragraph.borders_mut().set_right(border);
    } else if matches_local_name(event.name().as_ref(), b"between") {
        paragraph.borders_mut().set_between(border);
    }
}

fn maybe_apply_paragraph_shading(paragraph: &mut Option<Paragraph>, event: &BytesStart<'_>) {
    let Some(paragraph) = paragraph.as_mut() else {
        return;
    };

    let fill = parse_attribute_value(event, b"fill").and_then(|value| {
        let trimmed = value.trim().trim_start_matches('#').to_ascii_uppercase();
        if trimmed.is_empty() || trimmed == "AUTO" {
            None
        } else {
            Some(trimmed)
        }
    });
    let pattern = parse_attribute_value(event, b"val").and_then(|value| {
        let trimmed = value.trim();
        if trimmed.is_empty() {
            None
        } else {
            Some(trimmed.to_string())
        }
    });
    paragraph.set_shading_color_option(fill);
    paragraph.set_shading_pattern_option(pattern);

    let color_attr = parse_attribute_value(event, b"color").and_then(|value| {
        let trimmed = value.trim();
        if trimmed.is_empty() {
            None
        } else {
            Some(trimmed.to_string())
        }
    });
    paragraph.set_shading_color_attribute_option(color_attr);
}

fn parse_inline_section_xml(xml: &[u8]) -> Result<Section> {
    let mut reader = Reader::from_reader(Cursor::new(xml));
    reader.config_mut().trim_text(false);

    let mut section = Section::new();
    let mut buffer = Vec::new();

    loop {
        match reader.read_event_into(&mut buffer)? {
            Event::Start(ref event) => {
                if matches_local_name(event.name().as_ref(), b"pgSz") {
                    maybe_apply_section_page_size(&mut section, event);
                } else if matches_local_name(event.name().as_ref(), b"pgMar") {
                    maybe_apply_section_page_margins(&mut section, event);
                } else if matches_local_name(event.name().as_ref(), b"type") {
                    if let Some(value) = parse_attribute_value(event, b"val") {
                        section.set_break_type_option(SectionBreakType::from_xml_value(
                            value.as_str(),
                        ));
                    }
                } else if matches_local_name(event.name().as_ref(), b"titlePg") {
                    section.set_title_page(parse_on_off_property(event, true));
                } else if matches_local_name(event.name().as_ref(), b"cols") {
                    maybe_apply_section_columns(&mut section, event);
                } else if matches_local_name(event.name().as_ref(), b"vAlign") {
                    if let Some(value) = parse_attribute_value(event, b"val") {
                        section.set_vertical_alignment(
                            SectionVerticalAlignment::from_xml_value(value.as_str())
                                .unwrap_or(SectionVerticalAlignment::Top),
                        );
                    }
                } else if matches_local_name(event.name().as_ref(), b"lnNumType") {
                    maybe_apply_section_line_numbering(&mut section, event);
                }
            }
            Event::Empty(ref event) => {
                if matches_local_name(event.name().as_ref(), b"pgSz") {
                    maybe_apply_section_page_size(&mut section, event);
                } else if matches_local_name(event.name().as_ref(), b"pgMar") {
                    maybe_apply_section_page_margins(&mut section, event);
                } else if matches_local_name(event.name().as_ref(), b"type") {
                    if let Some(value) = parse_attribute_value(event, b"val") {
                        section.set_break_type_option(SectionBreakType::from_xml_value(
                            value.as_str(),
                        ));
                    }
                } else if matches_local_name(event.name().as_ref(), b"titlePg") {
                    section.set_title_page(parse_on_off_property(event, true));
                } else if matches_local_name(event.name().as_ref(), b"cols") {
                    maybe_apply_section_columns(&mut section, event);
                } else if matches_local_name(event.name().as_ref(), b"vAlign") {
                    if let Some(value) = parse_attribute_value(event, b"val") {
                        section.set_vertical_alignment(
                            SectionVerticalAlignment::from_xml_value(value.as_str())
                                .unwrap_or(SectionVerticalAlignment::Top),
                        );
                    }
                } else if matches_local_name(event.name().as_ref(), b"lnNumType") {
                    maybe_apply_section_line_numbering(&mut section, event);
                }
            }
            Event::Eof => break,
            _ => {}
        }
        buffer.clear();
    }

    Ok(section)
}

fn maybe_apply_run_image_relationship(
    run_properties: &mut CurrentRunProperties,
    event: &BytesStart<'_>,
) {
    if let Some(relationship_id) = parse_attribute_value(event, b"embed") {
        run_properties.image_relationship_id = Some(relationship_id);
    }
}

fn maybe_apply_run_image_extent(run_properties: &mut CurrentRunProperties, event: &BytesStart<'_>) {
    if let Some(width_emu) = parse_u32_attribute_value(event, b"cx") {
        run_properties.image_width_emu = Some(width_emu);
    }
    if let Some(height_emu) = parse_u32_attribute_value(event, b"cy") {
        run_properties.image_height_emu = Some(height_emu);
    }
}

fn maybe_apply_run_image_doc_properties(
    run_properties: &mut CurrentRunProperties,
    event: &BytesStart<'_>,
) {
    if let Some(name) = parse_attribute_value(event, b"name") {
        if !name.trim().is_empty() {
            run_properties.image_name = Some(name);
        }
    }
    if let Some(description) = parse_attribute_value(event, b"descr") {
        if !description.trim().is_empty() {
            run_properties.image_description = Some(description);
        }
    }
}

fn maybe_apply_run_floating_image_simple_position(
    run_properties: &mut CurrentRunProperties,
    event: &BytesStart<'_>,
) {
    if let Some(offset_x_emu) = parse_i32_attribute_value(event, b"x") {
        run_properties.floating_image_offset_x_emu = Some(offset_x_emu);
    }
    if let Some(offset_y_emu) = parse_i32_attribute_value(event, b"y") {
        run_properties.floating_image_offset_y_emu = Some(offset_y_emu);
    }
}

fn maybe_apply_run_floating_image_position_offset(
    run_properties: &mut CurrentRunProperties,
    text: &str,
    horizontal: bool,
) {
    let parsed = text.trim().parse::<i32>().ok();
    if horizontal {
        if let Some(offset_x_emu) = parsed {
            run_properties.floating_image_offset_x_emu = Some(offset_x_emu);
        }
    } else if let Some(offset_y_emu) = parsed {
        run_properties.floating_image_offset_y_emu = Some(offset_y_emu);
    }
}

fn maybe_apply_table_style_id(current_table_style_id: &mut Option<String>, event: &BytesStart<'_>) {
    let Some(style_id) = parse_attribute_value(event, b"val") else {
        return;
    };
    let style_id = style_id.trim();
    if style_id.is_empty() {
        return;
    }

    *current_table_style_id = Some(style_id.to_string());
}

fn maybe_apply_table_border_edge(borders: &mut TableBorders, event: &BytesStart<'_>) {
    let Some(border) = parse_table_border(event) else {
        return;
    };

    if matches_local_name(event.name().as_ref(), b"top") {
        borders.set_top(border);
    } else if matches_local_name(event.name().as_ref(), b"left") {
        borders.set_left(border);
    } else if matches_local_name(event.name().as_ref(), b"bottom") {
        borders.set_bottom(border);
    } else if matches_local_name(event.name().as_ref(), b"right") {
        borders.set_right(border);
    } else if matches_local_name(event.name().as_ref(), b"insideH") {
        borders.set_inside_horizontal(border);
    } else if matches_local_name(event.name().as_ref(), b"insideV") {
        borders.set_inside_vertical(border);
    }
}

fn maybe_apply_cell_border_edge(borders: &mut CellBorders, event: &BytesStart<'_>) {
    let Some(border) = parse_table_border(event) else {
        return;
    };

    if matches_local_name(event.name().as_ref(), b"top") {
        borders.set_top(border);
    } else if matches_local_name(event.name().as_ref(), b"left") {
        borders.set_left(border);
    } else if matches_local_name(event.name().as_ref(), b"bottom") {
        borders.set_bottom(border);
    } else if matches_local_name(event.name().as_ref(), b"right") {
        borders.set_right(border);
    }
}

fn maybe_apply_cell_margin_edge(margins: &mut CellMargins, event: &BytesStart<'_>) {
    let Some(width) = parse_u32_attribute_value(event, b"w") else {
        return;
    };

    if matches_local_name(event.name().as_ref(), b"top") {
        margins.set_top_twips(width);
    } else if matches_local_name(event.name().as_ref(), b"left")
        || matches_local_name(event.name().as_ref(), b"start")
    {
        margins.set_left_twips(width);
    } else if matches_local_name(event.name().as_ref(), b"bottom") {
        margins.set_bottom_twips(width);
    } else if matches_local_name(event.name().as_ref(), b"right")
        || matches_local_name(event.name().as_ref(), b"end")
    {
        margins.set_right_twips(width);
    }
}

fn parse_table_border(event: &BytesStart<'_>) -> Option<TableBorder> {
    let mut border = TableBorder::default();
    border.set_line_type_option(parse_attribute_value(event, b"val"));
    if let Some(size) =
        parse_u32_attribute_value(event, b"sz").and_then(|value| u16::try_from(value).ok())
    {
        border.set_size_eighth_points(size);
    }
    if let Some(color) = parse_attribute_value(event, b"color") {
        border.set_color(color);
    }
    if let Some(space) =
        parse_u32_attribute_value(event, b"space").and_then(|value| u16::try_from(value).ok())
    {
        border.set_space_eighth_points(space);
    }

    if border.line_type().is_none()
        && border.size_eighth_points().is_none()
        && border.color().is_none()
        && border.space_eighth_points().is_none()
    {
        None
    } else {
        Some(border)
    }
}

fn parse_vertical_merge(event: &BytesStart<'_>) -> VerticalMerge {
    match parse_attribute_value(event, b"val") {
        Some(value) if value.eq_ignore_ascii_case("restart") => VerticalMerge::Restart,
        _ => VerticalMerge::Continue,
    }
}

fn parse_shading_fill_color(event: &BytesStart<'_>) -> Option<String> {
    let fill = parse_attribute_value(event, b"fill")?;
    let trimmed = fill.trim().trim_start_matches('#').to_ascii_uppercase();
    if trimmed.is_empty() || trimmed == "AUTO" {
        None
    } else {
        Some(trimmed)
    }
}

fn parse_shading_color_attribute(event: &BytesStart<'_>) -> Option<String> {
    let color = parse_attribute_value(event, b"color")?;
    let trimmed = color.trim();
    if trimmed.is_empty() {
        None
    } else {
        Some(trimmed.to_string())
    }
}

fn parse_shading_pattern(event: &BytesStart<'_>) -> Option<String> {
    let val = parse_attribute_value(event, b"val")?;
    let trimmed = val.trim();
    if trimmed.is_empty() {
        None
    } else {
        Some(trimmed.to_string())
    }
}

fn parse_vertical_alignment(event: &BytesStart<'_>) -> Option<VerticalAlignment> {
    let value = parse_attribute_value(event, b"val")?;
    match value.to_ascii_lowercase().as_str() {
        "top" => Some(VerticalAlignment::Top),
        "center" => Some(VerticalAlignment::Center),
        "bottom" => Some(VerticalAlignment::Bottom),
        _ => None,
    }
}

fn maybe_apply_section_page_size(section: &mut Section, event: &BytesStart<'_>) {
    section.set_page_width_twips(parse_u32_attribute_value(event, b"w"));
    section.set_page_height_twips(parse_u32_attribute_value(event, b"h"));
    let orientation = parse_attribute_value(event, b"orient")
        .and_then(|value| PageOrientation::from_xml_value(value.to_ascii_lowercase().as_str()));
    section.set_page_orientation_option(orientation);
}

fn maybe_apply_section_page_margins(section: &mut Section, event: &BytesStart<'_>) {
    let mut margins = PageMargins::new();

    if let Some(top) = parse_u32_attribute_value(event, b"top") {
        margins.set_top_twips(top);
    }
    if let Some(right) = parse_u32_attribute_value(event, b"right") {
        margins.set_right_twips(right);
    }
    if let Some(bottom) = parse_u32_attribute_value(event, b"bottom") {
        margins.set_bottom_twips(bottom);
    }
    if let Some(left) = parse_u32_attribute_value(event, b"left") {
        margins.set_left_twips(left);
    }
    if let Some(header) = parse_u32_attribute_value(event, b"header") {
        margins.set_header_twips(header);
    }
    if let Some(footer) = parse_u32_attribute_value(event, b"footer") {
        margins.set_footer_twips(footer);
    }
    if let Some(gutter) = parse_u32_attribute_value(event, b"gutter") {
        margins.set_gutter_twips(gutter);
    }

    section.set_page_margins(margins);
}

fn maybe_apply_section_reference(
    section_relationship_ids: &mut SectionRelationshipIds,
    _section: &mut Section,
    event: &BytesStart<'_>,
    is_header: bool,
) {
    let Some(relationship_id) = parse_attribute_value(event, b"id") else {
        return;
    };
    let reference_type = parse_attribute_value(event, b"type")
        .map(|value| value.to_ascii_lowercase())
        .unwrap_or_else(|| "default".to_string());

    match reference_type.as_str() {
        "default" => {
            if is_header {
                section_relationship_ids.header_relationship_id = Some(relationship_id);
            } else {
                section_relationship_ids.footer_relationship_id = Some(relationship_id);
            }
        }
        "first" => {
            if is_header {
                section_relationship_ids.first_page_header_relationship_id = Some(relationship_id);
            } else {
                section_relationship_ids.first_page_footer_relationship_id = Some(relationship_id);
            }
        }
        "even" => {
            if is_header {
                section_relationship_ids.even_page_header_relationship_id = Some(relationship_id);
            } else {
                section_relationship_ids.even_page_footer_relationship_id = Some(relationship_id);
            }
        }
        _ => {
            // Unknown reference type — OOXML only defines "default", "first", "even".
        }
    }
}

fn load_header_footer_part(
    package: &Package,
    document_part_uri: &PartUri,
    document_part: &Part,
    relationship_id: Option<&str>,
    expected_relationship_type: &str,
) -> Result<Option<HeaderFooter>> {
    let Some(relationship_id) = relationship_id else {
        return Ok(None);
    };

    let Some(relationship) = document_part
        .relationships
        .iter()
        .find(|relationship| relationship.id == relationship_id)
    else {
        return Ok(None);
    };
    if relationship.rel_type != expected_relationship_type
        || relationship.target_mode != TargetMode::Internal
    {
        return Ok(None);
    }

    let part_uri = document_part_uri.resolve_relative(relationship.target.as_str())?;
    let Some(part) = package.get_part(part_uri.as_str()) else {
        return Ok(None);
    };

    let hyperlink_targets = parse_hyperlink_targets_by_relationship_id(part);
    let image_indexes_by_relationship_id = HashMap::new();
    let paragraphs = parse_paragraphs(
        part.data.as_bytes(),
        &hyperlink_targets,
        &image_indexes_by_relationship_id,
    )?;
    let tables = parse_tables(part.data.as_bytes())?;

    let mut header_footer = HeaderFooter::new();
    header_footer.set_paragraphs(paragraphs);
    header_footer.set_tables(tables);
    Ok(Some(header_footer))
}

fn resolve_hyperlink_target(
    hyperlink_targets: &HashMap<String, String>,
    event: &BytesStart<'_>,
) -> Option<String> {
    let relationship_id = parse_attribute_value(event, b"id")?;

    hyperlink_targets.get(relationship_id.as_str()).cloned()
}

fn parse_run_font_family(event: &BytesStart<'_>) -> Option<String> {
    parse_attribute_value(event, b"ascii")
        .or_else(|| parse_attribute_value(event, b"hAnsi"))
        .or_else(|| parse_attribute_value(event, b"cs"))
        .or_else(|| parse_attribute_value(event, b"eastAsia"))
        .and_then(|value| if value.is_empty() { None } else { Some(value) })
}

fn parse_run_size_half_points(event: &BytesStart<'_>) -> Option<u16> {
    parse_attribute_value(event, b"val")?.parse::<u16>().ok()
}

fn parse_run_color(event: &BytesStart<'_>) -> Option<String> {
    let value = parse_attribute_value(event, b"val")?;
    normalize_color_value(value.as_str())
}

fn normalize_color_value(value: &str) -> Option<String> {
    let trimmed = value.trim();
    if trimmed.is_empty() {
        return None;
    }

    Some(trimmed.trim_start_matches('#').to_ascii_uppercase())
}

fn parse_u32_attribute_value(event: &BytesStart<'_>, attribute_name: &[u8]) -> Option<u32> {
    parse_attribute_value(event, attribute_name)?
        .parse::<u32>()
        .ok()
}

fn parse_i32_attribute_value(event: &BytesStart<'_>, attribute_name: &[u8]) -> Option<i32> {
    parse_attribute_value(event, attribute_name)?
        .parse::<i32>()
        .ok()
}

fn parse_attribute_value(event: &BytesStart<'_>, attribute_name: &[u8]) -> Option<String> {
    for attribute in event.attributes().flatten() {
        if matches_local_name(attribute.key.as_ref(), attribute_name) {
            if let Ok(value) = attribute.unescape_value() {
                return Some(value.into_owned());
            }
            let fallback = String::from_utf8_lossy(attribute.value.as_ref()).into_owned();
            return Some(fallback);
        }
    }

    None
}

fn parse_on_off_property(event: &BytesStart<'_>, default: bool) -> bool {
    for attribute in event.attributes().flatten() {
        if matches_local_name(attribute.key.as_ref(), b"val") {
            let value = String::from_utf8_lossy(attribute.value.as_ref()).to_ascii_lowercase();
            return !matches!(value.as_str(), "0" | "false" | "off" | "no");
        }
    }

    default
}

fn parse_underline_property(event: &BytesStart<'_>) -> Option<UnderlineType> {
    for attribute in event.attributes().flatten() {
        if matches_local_name(attribute.key.as_ref(), b"val") {
            let value = String::from_utf8_lossy(attribute.value.as_ref());
            return match value.as_ref() {
                "none" | "0" | "false" | "off" | "no" => None,
                "single" => Some(UnderlineType::Single),
                "double" => Some(UnderlineType::Double),
                "thick" => Some(UnderlineType::Thick),
                "dotted" => Some(UnderlineType::Dotted),
                "dottedHeavy" => Some(UnderlineType::DottedHeavy),
                "dash" => Some(UnderlineType::Dash),
                "dashedHeavy" => Some(UnderlineType::DashedHeavy),
                "dashLong" => Some(UnderlineType::DashLong),
                "dashLongHeavy" => Some(UnderlineType::DashLongHeavy),
                "dotDash" => Some(UnderlineType::DashDot),
                "dashDotHeavy" => Some(UnderlineType::DashDotHeavy),
                "dotDotDash" => Some(UnderlineType::DashDotDot),
                "dashDotDotHeavy" => Some(UnderlineType::DashDotDotHeavy),
                "wave" => Some(UnderlineType::Wavy),
                "wavyHeavy" => Some(UnderlineType::WavyHeavy),
                "wavyDouble" => Some(UnderlineType::WavyDouble),
                "words" => Some(UnderlineType::Words),
                _ => Some(UnderlineType::Single), // unknown type defaults to single
            };
        }
    }
    Some(UnderlineType::Single) // w:u present with no val means single
}

fn parse_vert_align_property(event: &BytesStart<'_>, props: &mut CurrentRunProperties) {
    if let Some(value) = parse_attribute_value(event, b"val") {
        let lower = value.to_ascii_lowercase();
        match lower.as_str() {
            "subscript" => {
                props.subscript = true;
                props.superscript = false;
            }
            "superscript" => {
                props.superscript = true;
                props.subscript = false;
            }
            _ => {}
        }
    }
}

fn maybe_apply_paragraph_style(paragraph: &mut Option<Paragraph>, event: &BytesStart<'_>) {
    let Some(paragraph) = paragraph.as_mut() else {
        return;
    };

    let Some(style_name) = parse_attribute_value(event, b"val") else {
        return;
    };
    let style_name = style_name.trim();
    if style_name.is_empty() {
        return;
    }

    paragraph.set_style_id_option(Some(style_name.to_string()));
    if let Some(level) = heading_level_from_style_name(style_name) {
        paragraph.set_heading_level(Some(level));
    }
}

fn heading_level_from_style_name(style_name: &str) -> Option<u8> {
    let level_text = style_name
        .strip_prefix("Heading")
        .or_else(|| style_name.strip_prefix("heading"))?;
    let level = level_text.parse::<u8>().ok()?;

    (1..=9).contains(&level).then_some(level)
}

fn local_name(name: &[u8]) -> &[u8] {
    name.rsplit(|byte| *byte == b':').next().unwrap_or(name)
}

fn matches_local_name(name: &[u8], local_name: &[u8]) -> bool {
    if name == local_name {
        return true;
    }

    name.rsplit(|byte| *byte == b':')
        .next()
        .is_some_and(|local| local == local_name)
}

fn resolve_word_document_part_uri(package: &Package) -> Result<String> {
    for relationship in package
        .relationships()
        .get_by_type(RelationshipType::WORD_DOCUMENT)
    {
        if relationship.target_mode != TargetMode::Internal {
            continue;
        }

        let part_uri = normalize_relationship_target(relationship.target.as_str())?;
        if package.get_part(part_uri.as_str()).is_some() {
            return Ok(part_uri);
        }
    }

    if package.get_part(WORD_DOCUMENT_URI).is_some() {
        return Ok(WORD_DOCUMENT_URI.to_string());
    }

    Err(DocxError::UnsupportedPackage(format!(
        "missing officeDocument relationship to `{WORD_DOCUMENT_URI}`"
    )))
}

fn normalize_relationship_target(target: &str) -> Result<String> {
    let mut normalized = target.trim().replace('\\', "/");
    while let Some(stripped) = normalized.strip_prefix("./") {
        normalized = stripped.to_string();
    }
    if !normalized.starts_with('/') {
        normalized.insert(0, '/');
    }

    Ok(PartUri::new(normalized)?.to_string())
}

fn maybe_apply_section_page_num_type(section: &mut Section, event: &BytesStart<'_>) {
    if let Some(start) = parse_u32_attribute_value(event, b"start") {
        section.set_page_number_start_option(Some(start));
    }
    if let Some(format) = parse_attribute_value(event, b"fmt") {
        section.set_page_number_format_option(Some(format));
    }
}

fn maybe_apply_section_columns(section: &mut Section, event: &BytesStart<'_>) {
    if let Some(num) = parse_u32_attribute_value(event, b"num").and_then(|v| u16::try_from(v).ok())
    {
        section.set_column_count(num);
    }
    if let Some(space) = parse_u32_attribute_value(event, b"space") {
        section.set_column_space_twips(space);
    }
    if parse_attribute_value(event, b"sep").is_some_and(|v| v == "1" || v == "true") {
        section.set_column_separator(true);
    }
}

fn maybe_apply_section_line_numbering(section: &mut Section, event: &BytesStart<'_>) {
    if let Some(start) = parse_u32_attribute_value(event, b"start") {
        section.set_line_numbering_start(start);
    }
    if let Some(count_by) = parse_u32_attribute_value(event, b"countBy") {
        section.set_line_numbering_count_by(count_by);
    }
    if let Some(restart) = parse_attribute_value(event, b"restart") {
        if let Some(restart_val) = LineNumberRestart::from_xml_value(restart.as_str()) {
            section.set_line_numbering_restart(restart_val);
        }
    }
    if let Some(distance) = parse_u32_attribute_value(event, b"distance") {
        section.set_line_numbering_distance_twips(distance);
    }
}

fn load_footnotes(
    package: &Package,
    document_part_uri: &PartUri,
    document_part: &Part,
) -> Result<Vec<Footnote>> {
    let Some(rel) = document_part
        .relationships
        .get_first_by_type(WORD_FOOTNOTES_REL_TYPE)
    else {
        return Ok(Vec::new());
    };

    let part_uri = document_part_uri.resolve_relative(rel.target.as_str())?;
    let Some(part) = package.get_part(part_uri.as_str()) else {
        return Ok(Vec::new());
    };

    parse_footnotes_xml(part.data.as_bytes())
}

fn parse_footnotes_xml(xml: &[u8]) -> Result<Vec<Footnote>> {
    let mut reader = Reader::from_reader(Cursor::new(xml));
    reader.config_mut().trim_text(false);

    let mut footnotes = Vec::new();
    let mut current_id: Option<u32> = None;
    let mut current_text = String::new();
    let mut in_text = false;
    let mut buffer = Vec::new();

    loop {
        match reader.read_event_into(&mut buffer)? {
            Event::Start(ref event) => {
                if matches_local_name(event.name().as_ref(), b"footnote") {
                    current_id = parse_u32_attribute_value(event, b"id");
                    current_text.clear();
                } else if current_id.is_some() && matches_local_name(event.name().as_ref(), b"t") {
                    in_text = true;
                }
            }
            Event::Text(ref event) => {
                if in_text {
                    if let Ok(text) = event.xml_content() {
                        current_text.push_str(text.as_ref());
                    }
                }
            }
            Event::End(ref event) => {
                if matches_local_name(event.name().as_ref(), b"t") {
                    in_text = false;
                } else if matches_local_name(event.name().as_ref(), b"footnote") {
                    if let Some(id) = current_id.take() {
                        // Separator and continuation separator footnotes (type="separator" etc.)
                        // typically have ids 0 and 1; we keep all for fidelity.
                        let footnote = if current_text.is_empty() {
                            Footnote::new(id)
                        } else {
                            Footnote::from_text(id, current_text.clone())
                        };
                        footnotes.push(footnote);
                    }
                    current_text.clear();
                }
            }
            Event::Eof => break,
            _ => {}
        }
        buffer.clear();
    }

    Ok(footnotes)
}

fn load_endnotes(
    package: &Package,
    document_part_uri: &PartUri,
    document_part: &Part,
) -> Result<Vec<Endnote>> {
    let Some(rel) = document_part
        .relationships
        .get_first_by_type(WORD_ENDNOTES_REL_TYPE)
    else {
        return Ok(Vec::new());
    };

    let part_uri = document_part_uri.resolve_relative(rel.target.as_str())?;
    let Some(part) = package.get_part(part_uri.as_str()) else {
        return Ok(Vec::new());
    };

    parse_endnotes_xml(part.data.as_bytes())
}

fn parse_endnotes_xml(xml: &[u8]) -> Result<Vec<Endnote>> {
    let mut reader = Reader::from_reader(Cursor::new(xml));
    reader.config_mut().trim_text(false);

    let mut endnotes = Vec::new();
    let mut current_id: Option<u32> = None;
    let mut current_text = String::new();
    let mut in_text = false;
    let mut buffer = Vec::new();

    loop {
        match reader.read_event_into(&mut buffer)? {
            Event::Start(ref event) => {
                if matches_local_name(event.name().as_ref(), b"endnote") {
                    current_id = parse_u32_attribute_value(event, b"id");
                    current_text.clear();
                } else if current_id.is_some() && matches_local_name(event.name().as_ref(), b"t") {
                    in_text = true;
                }
            }
            Event::Text(ref event) => {
                if in_text {
                    if let Ok(text) = event.xml_content() {
                        current_text.push_str(text.as_ref());
                    }
                }
            }
            Event::End(ref event) => {
                if matches_local_name(event.name().as_ref(), b"t") {
                    in_text = false;
                } else if matches_local_name(event.name().as_ref(), b"endnote") {
                    if let Some(id) = current_id.take() {
                        let endnote = if current_text.is_empty() {
                            Endnote::new(id)
                        } else {
                            Endnote::from_text(id, current_text.clone())
                        };
                        endnotes.push(endnote);
                    }
                    current_text.clear();
                }
            }
            Event::Eof => break,
            _ => {}
        }
        buffer.clear();
    }

    Ok(endnotes)
}

/// Parse the first occurrence of `root_element_name` in `xml` and capture
/// extra `xmlns:*` namespace declarations that are not in `always_emitted`.
fn parse_root_element_namespace_declarations(
    xml: &[u8],
    root_element_name: &[u8],
    always_emitted: &[&str],
) -> Vec<(String, String)> {
    let mut reader = Reader::from_reader(Cursor::new(xml));
    reader.config_mut().trim_text(false);
    let mut buffer = Vec::new();
    loop {
        match reader.read_event_into(&mut buffer) {
            Ok(Event::Start(ref e)) | Ok(Event::Empty(ref e)) => {
                if local_name(e.name().as_ref()) == root_element_name {
                    return offidized_opc::xml_util::capture_extra_namespace_declarations(
                        e,
                        always_emitted,
                    );
                }
            }
            Ok(Event::Eof) | Err(_) => return Vec::new(),
            _ => {}
        }
        buffer.clear();
    }
}

fn parse_bookmarks(xml: &[u8]) -> Result<Vec<Bookmark>> {
    let mut reader = Reader::from_reader(Cursor::new(xml));
    reader.config_mut().trim_text(false);

    let mut bookmarks = Vec::new();
    let mut bookmark_starts: HashMap<u32, (String, usize)> = HashMap::new();
    let mut in_body = false;
    let mut table_depth = 0_usize;
    let mut paragraph_index = 0_usize;
    let mut buffer = Vec::new();

    loop {
        match reader.read_event_into(&mut buffer)? {
            Event::Start(ref event) => {
                if matches_local_name(event.name().as_ref(), b"body") {
                    in_body = true;
                } else if in_body && matches_local_name(event.name().as_ref(), b"tbl") {
                    table_depth = table_depth.saturating_add(1);
                } else if in_body
                    && table_depth == 0
                    && matches_local_name(event.name().as_ref(), b"p")
                {
                    // paragraph_index tracks the current paragraph being parsed
                } else if in_body && matches_local_name(event.name().as_ref(), b"bookmarkStart") {
                    if let Some(id) = parse_u32_attribute_value(event, b"id") {
                        let name = parse_attribute_value(event, b"name").unwrap_or_default();
                        bookmark_starts.insert(id, (name, paragraph_index));
                    }
                } else if in_body && matches_local_name(event.name().as_ref(), b"bookmarkEnd") {
                    if let Some(id) = parse_u32_attribute_value(event, b"id") {
                        if let Some((name, start_para)) = bookmark_starts.remove(&id) {
                            bookmarks.push(Bookmark::new(id, name, start_para, paragraph_index));
                        }
                    }
                }
            }
            Event::Empty(ref event) => {
                if in_body && matches_local_name(event.name().as_ref(), b"bookmarkStart") {
                    if let Some(id) = parse_u32_attribute_value(event, b"id") {
                        let name = parse_attribute_value(event, b"name").unwrap_or_default();
                        bookmark_starts.insert(id, (name, paragraph_index));
                    }
                } else if in_body && matches_local_name(event.name().as_ref(), b"bookmarkEnd") {
                    if let Some(id) = parse_u32_attribute_value(event, b"id") {
                        if let Some((name, start_para)) = bookmark_starts.remove(&id) {
                            bookmarks.push(Bookmark::new(id, name, start_para, paragraph_index));
                        }
                    }
                } else if in_body
                    && table_depth == 0
                    && matches_local_name(event.name().as_ref(), b"p")
                {
                    paragraph_index = paragraph_index.saturating_add(1);
                }
            }
            Event::End(ref event) => {
                if matches_local_name(event.name().as_ref(), b"body") {
                    in_body = false;
                } else if in_body && matches_local_name(event.name().as_ref(), b"tbl") {
                    table_depth = table_depth.saturating_sub(1);
                } else if in_body
                    && table_depth == 0
                    && matches_local_name(event.name().as_ref(), b"p")
                {
                    paragraph_index = paragraph_index.saturating_add(1);
                }
            }
            Event::Eof => break,
            _ => {}
        }
        buffer.clear();
    }

    // Any bookmarkStart without a matching bookmarkEnd: treat as single-paragraph bookmark.
    for (id, (name, start_para)) in bookmark_starts {
        bookmarks.push(Bookmark::new(id, name, start_para, start_para));
    }

    Ok(bookmarks)
}

fn load_comments(
    package: &Package,
    document_part_uri: &PartUri,
    document_part: &Part,
) -> Result<Vec<Comment>> {
    let Some(rel) = document_part
        .relationships
        .get_first_by_type(WORD_COMMENTS_REL_TYPE)
    else {
        return Ok(Vec::new());
    };

    let part_uri = document_part_uri.resolve_relative(rel.target.as_str())?;
    let Some(part) = package.get_part(part_uri.as_str()) else {
        return Ok(Vec::new());
    };

    parse_comments_xml(part.data.as_bytes())
}

fn parse_comments_xml(xml: &[u8]) -> Result<Vec<Comment>> {
    let mut reader = Reader::from_reader(Cursor::new(xml));
    reader.config_mut().trim_text(false);

    let mut comments = Vec::new();
    let mut current_id: Option<u32> = None;
    let mut current_author: Option<String> = None;
    let mut current_date: Option<String> = None;
    let mut current_text = String::new();
    let mut in_text = false;
    let mut buffer = Vec::new();

    loop {
        match reader.read_event_into(&mut buffer)? {
            Event::Start(ref event) => {
                if matches_local_name(event.name().as_ref(), b"comment") {
                    current_id = parse_u32_attribute_value(event, b"id");
                    current_author = parse_attribute_value(event, b"author");
                    current_date = parse_attribute_value(event, b"date");
                    current_text.clear();
                } else if current_id.is_some() && matches_local_name(event.name().as_ref(), b"t") {
                    in_text = true;
                }
            }
            Event::Text(ref event) => {
                if in_text {
                    if let Ok(text) = event.xml_content() {
                        current_text.push_str(text.as_ref());
                    }
                }
            }
            Event::End(ref event) => {
                if matches_local_name(event.name().as_ref(), b"t") {
                    in_text = false;
                } else if matches_local_name(event.name().as_ref(), b"comment") {
                    if let Some(id) = current_id.take() {
                        let author = current_author.take().unwrap_or_default();
                        let mut comment = if current_text.is_empty() {
                            Comment::new(id, author)
                        } else {
                            Comment::from_text(id, author, current_text.clone())
                        };
                        if let Some(date) = current_date.take() {
                            comment.set_date(date);
                        }
                        comments.push(comment);
                    }
                    current_text.clear();
                    current_author = None;
                    current_date = None;
                }
            }
            Event::Eof => break,
            _ => {}
        }
        buffer.clear();
    }

    Ok(comments)
}

fn load_document_properties(package: &Package) -> Result<DocumentProperties> {
    let Some(part) = package.get_part(CORE_PROPERTIES_URI) else {
        return Ok(DocumentProperties::new());
    };

    parse_core_properties_xml(part.data.as_bytes())
}

fn parse_core_properties_xml(xml: &[u8]) -> Result<DocumentProperties> {
    let mut reader = Reader::from_reader(Cursor::new(xml));
    reader.config_mut().trim_text(false);

    let mut props = DocumentProperties::new();
    let mut current_element: Option<String> = None;
    let mut current_text = String::new();
    let mut buffer = Vec::new();

    loop {
        match reader.read_event_into(&mut buffer)? {
            Event::Start(ref event) => {
                let name_bytes = event.name().as_ref().to_vec();
                let local = local_name(&name_bytes);
                match local {
                    b"title" | b"subject" | b"creator" | b"description" | b"keywords"
                    | b"lastModifiedBy" | b"created" | b"modified" => {
                        current_element = Some(String::from_utf8_lossy(local).into_owned());
                        current_text.clear();
                    }
                    _ => {}
                }
            }
            Event::Text(ref event) => {
                if current_element.is_some() {
                    if let Ok(text) = event.xml_content() {
                        current_text.push_str(text.as_ref());
                    }
                }
            }
            Event::End(ref event) => {
                let name_bytes = event.name().as_ref().to_vec();
                let local = local_name(&name_bytes);
                if let Some(ref elem) = current_element {
                    let local_str = std::str::from_utf8(local).unwrap_or("");
                    if local_str == elem.as_str() {
                        match elem.as_str() {
                            "title" => props.set_title(current_text.clone()),
                            "subject" => props.set_subject(current_text.clone()),
                            "creator" => props.set_creator(current_text.clone()),
                            "description" => props.set_description(current_text.clone()),
                            "keywords" => props.set_keywords(current_text.clone()),
                            "lastModifiedBy" => props.set_last_modified_by(current_text.clone()),
                            "created" => props.set_created(current_text.clone()),
                            "modified" => props.set_modified(current_text.clone()),
                            _ => {}
                        }
                        current_element = None;
                        current_text.clear();
                    }
                }
            }
            Event::Eof => break,
            _ => {}
        }
        buffer.clear();
    }

    Ok(props)
}

fn parse_content_controls(xml: &[u8]) -> Result<Vec<ContentControl>> {
    let mut reader = Reader::from_reader(Cursor::new(xml));
    reader.config_mut().trim_text(false);

    let mut content_controls = Vec::new();
    let mut in_body = false;
    let mut table_depth = 0_usize;
    let mut sdt_depth = 0_usize;
    let mut in_sdt_pr = false;
    let mut current_tag: Option<String> = None;
    let mut current_alias: Option<String> = None;
    let mut in_sdt_content = false;
    let mut sdt_paragraphs: Vec<String> = Vec::new();
    let mut current_para_text = String::new();
    let mut in_text = false;
    let mut current_unknown_sdt_pr_children: Vec<RawXmlNode> = Vec::new();
    let mut buffer = Vec::new();

    const KNOWN_SDT_PR_CHILDREN: &[&[u8]] = &[
        b"tag",
        b"alias",
        b"text",
        b"picture",
        b"comboBox",
        b"dropDownList",
        b"date",
        b"docPartList",
        b"lock",
        b"rPr",
        b"placeholder",
        b"showingPlcHdr",
        b"dataBinding",
        b"temporary",
        b"id",
        b"checkbox",
        b"repeatingSectionItem",
    ];

    loop {
        match reader.read_event_into(&mut buffer)? {
            Event::Start(ref event) => {
                if matches_local_name(event.name().as_ref(), b"body") {
                    in_body = true;
                } else if in_body && matches_local_name(event.name().as_ref(), b"tbl") {
                    table_depth = table_depth.saturating_add(1);
                } else if in_body
                    && table_depth == 0
                    && matches_local_name(event.name().as_ref(), b"sdt")
                {
                    sdt_depth = sdt_depth.saturating_add(1);
                    if sdt_depth == 1 {
                        current_tag = None;
                        current_alias = None;
                        sdt_paragraphs.clear();
                        current_unknown_sdt_pr_children.clear();
                        in_sdt_pr = false;
                        in_sdt_content = false;
                    }
                } else if sdt_depth == 1 && matches_local_name(event.name().as_ref(), b"sdtPr") {
                    in_sdt_pr = true;
                } else if sdt_depth == 1
                    && in_sdt_pr
                    && matches_local_name(event.name().as_ref(), b"tag")
                {
                    current_tag = parse_attribute_value(event, b"val");
                } else if sdt_depth == 1
                    && in_sdt_pr
                    && matches_local_name(event.name().as_ref(), b"alias")
                {
                    current_alias = parse_attribute_value(event, b"val");
                } else if sdt_depth == 1 && in_sdt_pr {
                    let name_bytes = event.name();
                    let local = local_name(name_bytes.as_ref());
                    if !KNOWN_SDT_PR_CHILDREN.contains(&local) {
                        current_unknown_sdt_pr_children
                            .push(RawXmlNode::read_element(&mut reader, event)?);
                    }
                } else if sdt_depth == 1 && matches_local_name(event.name().as_ref(), b"sdtContent")
                {
                    in_sdt_content = true;
                } else if sdt_depth == 1
                    && in_sdt_content
                    && matches_local_name(event.name().as_ref(), b"p")
                {
                    current_para_text.clear();
                } else if sdt_depth == 1
                    && in_sdt_content
                    && matches_local_name(event.name().as_ref(), b"t")
                {
                    in_text = true;
                }
            }
            Event::Empty(ref event) => {
                if sdt_depth == 1 && in_sdt_pr && matches_local_name(event.name().as_ref(), b"tag")
                {
                    current_tag = parse_attribute_value(event, b"val");
                } else if sdt_depth == 1
                    && in_sdt_pr
                    && matches_local_name(event.name().as_ref(), b"alias")
                {
                    current_alias = parse_attribute_value(event, b"val");
                } else if sdt_depth == 1 && in_sdt_pr {
                    let name_bytes = event.name();
                    let local = local_name(name_bytes.as_ref());
                    if !KNOWN_SDT_PR_CHILDREN.contains(&local) {
                        current_unknown_sdt_pr_children.push(RawXmlNode::from_empty_element(event));
                    }
                }
            }
            Event::Text(ref event) => {
                if in_text && sdt_depth == 1 && in_sdt_content {
                    if let Ok(text) = event.xml_content() {
                        current_para_text.push_str(text.as_ref());
                    }
                }
            }
            Event::End(ref event) => {
                if matches_local_name(event.name().as_ref(), b"body") {
                    in_body = false;
                } else if in_body && matches_local_name(event.name().as_ref(), b"tbl") {
                    table_depth = table_depth.saturating_sub(1);
                } else if matches_local_name(event.name().as_ref(), b"t") {
                    in_text = false;
                } else if sdt_depth == 1
                    && in_sdt_content
                    && matches_local_name(event.name().as_ref(), b"p")
                {
                    sdt_paragraphs.push(current_para_text.clone());
                    current_para_text.clear();
                } else if matches_local_name(event.name().as_ref(), b"sdtPr") {
                    in_sdt_pr = false;
                } else if matches_local_name(event.name().as_ref(), b"sdtContent") {
                    in_sdt_content = false;
                } else if matches_local_name(event.name().as_ref(), b"sdt") {
                    if sdt_depth == 1 {
                        let mut sdt = ContentControl::new();
                        if let Some(tag) = current_tag.take() {
                            sdt.set_tag(tag);
                        }
                        if let Some(alias) = current_alias.take() {
                            sdt.set_alias(alias);
                        }
                        for para_text in sdt_paragraphs.drain(..) {
                            sdt.add_paragraph(para_text);
                        }
                        for node in current_unknown_sdt_pr_children.drain(..) {
                            sdt.push_unknown_sdt_pr_child(node);
                        }
                        content_controls.push(sdt);
                    }
                    sdt_depth = sdt_depth.saturating_sub(1);
                }
            }
            Event::Eof => break,
            _ => {}
        }
        buffer.clear();
    }

    Ok(content_controls)
}

fn load_numbering_definitions(
    package: &Package,
    document_part_uri: &PartUri,
    document_part: &Part,
) -> Result<(Vec<NumberingDefinition>, Vec<NumberingInstance>)> {
    let Some(rel) = document_part
        .relationships
        .get_first_by_type(WORD_NUMBERING_REL_TYPE)
    else {
        return Ok((Vec::new(), Vec::new()));
    };

    let part_uri = document_part_uri.resolve_relative(rel.target.as_str())?;
    let Some(part) = package.get_part(part_uri.as_str()) else {
        return Ok((Vec::new(), Vec::new()));
    };

    parse_numbering_xml(part.data.as_bytes())
}

fn parse_numbering_xml(xml: &[u8]) -> Result<(Vec<NumberingDefinition>, Vec<NumberingInstance>)> {
    let mut reader = Reader::from_reader(Cursor::new(xml));
    reader.config_mut().trim_text(false);

    let mut definitions = Vec::new();
    let mut instances = Vec::new();
    let mut current_abstract_num_id: Option<u32> = None;
    let mut current_levels: Vec<NumberingLevel> = Vec::new();
    let mut in_lvl = false;
    let mut current_level_ilvl: Option<u8> = None;
    let mut current_level_start: Option<u32> = None;
    let mut current_level_format: Option<String> = None;
    let mut current_level_text: Option<String> = None;
    let mut current_level_alignment: Option<String> = None;

    // Level paragraph/run property tracking
    let mut in_lvl_ppr = false;
    let mut in_lvl_rpr = false;
    let mut in_lvl_tabs = false;
    let mut current_level_indent_left: Option<u32> = None;
    let mut current_level_indent_hanging: Option<u32> = None;
    let mut current_level_tab_stop: Option<u32> = None;
    let mut current_level_suffix: Option<String> = None;
    let mut current_level_font_family: Option<String> = None;
    let mut current_level_font_size: Option<u16> = None;
    let mut current_level_bold: Option<bool> = None;
    let mut current_level_italic: Option<bool> = None;
    let mut current_level_color: Option<String> = None;

    // w:num parsing state
    let mut in_num = false;
    let mut current_num_id: Option<u32> = None;
    let mut current_num_abstract_id: Option<u32> = None;
    let mut current_num_overrides: Vec<NumberingLevelOverride> = Vec::new();
    let mut in_lvl_override = false;
    let mut current_override_ilvl: Option<u8> = None;

    let mut buffer = Vec::new();

    loop {
        match reader.read_event_into(&mut buffer)? {
            Event::Start(ref event) => {
                let name_bytes = event.name();
                let local = local_name(name_bytes.as_ref());

                if local == b"abstractNum" {
                    current_abstract_num_id = parse_u32_attribute_value(event, b"abstractNumId");
                    current_levels.clear();
                } else if current_abstract_num_id.is_some() && local == b"lvl" {
                    in_lvl = true;
                    current_level_ilvl = parse_u32_attribute_value(event, b"ilvl")
                        .and_then(|v| u8::try_from(v).ok());
                    current_level_start = None;
                    current_level_format = None;
                    current_level_text = None;
                    current_level_alignment = None;
                    current_level_indent_left = None;
                    current_level_indent_hanging = None;
                    current_level_tab_stop = None;
                    current_level_suffix = None;
                    current_level_font_family = None;
                    current_level_font_size = None;
                    current_level_bold = None;
                    current_level_italic = None;
                    current_level_color = None;
                    in_lvl_ppr = false;
                    in_lvl_rpr = false;
                    in_lvl_tabs = false;
                } else if in_lvl && local == b"start" && !in_num {
                    current_level_start = parse_u32_attribute_value(event, b"val");
                } else if in_lvl && local == b"numFmt" {
                    current_level_format = parse_attribute_value(event, b"val");
                } else if in_lvl && local == b"lvlText" {
                    current_level_text = parse_attribute_value(event, b"val");
                } else if in_lvl && local == b"lvlJc" {
                    current_level_alignment = parse_attribute_value(event, b"val");
                } else if in_lvl && local == b"suff" {
                    current_level_suffix = parse_attribute_value(event, b"val");
                } else if in_lvl && local == b"pPr" {
                    in_lvl_ppr = true;
                } else if in_lvl && !in_lvl_ppr && local == b"rPr" {
                    in_lvl_rpr = true;
                } else if in_lvl_ppr && local == b"tabs" {
                    in_lvl_tabs = true;
                } else if in_lvl_ppr && local == b"ind" {
                    current_level_indent_left = parse_u32_attribute_value(event, b"left");
                    current_level_indent_hanging = parse_u32_attribute_value(event, b"hanging");
                } else if in_lvl_tabs && local == b"tab" {
                    if current_level_tab_stop.is_none() {
                        current_level_tab_stop = parse_u32_attribute_value(event, b"pos");
                    }
                } else if in_lvl_rpr && local == b"rFonts" {
                    current_level_font_family = parse_attribute_value(event, b"ascii");
                } else if in_lvl_rpr && local == b"sz" {
                    current_level_font_size = parse_u32_attribute_value(event, b"val")
                        .and_then(|v| u16::try_from(v).ok());
                } else if in_lvl_rpr && local == b"b" {
                    let val = parse_attribute_value(event, b"val");
                    current_level_bold = Some(!matches!(val.as_deref(), Some("0") | Some("false")));
                } else if in_lvl_rpr && local == b"i" {
                    let val = parse_attribute_value(event, b"val");
                    current_level_italic =
                        Some(!matches!(val.as_deref(), Some("0") | Some("false")));
                } else if in_lvl_rpr && local == b"color" {
                    current_level_color = parse_attribute_value(event, b"val");
                } else if local == b"num" && !in_lvl {
                    in_num = true;
                    current_num_id = parse_u32_attribute_value(event, b"numId");
                    current_num_abstract_id = None;
                    current_num_overrides.clear();
                } else if in_num && local == b"abstractNumId" {
                    current_num_abstract_id = parse_u32_attribute_value(event, b"val");
                } else if in_num && local == b"lvlOverride" {
                    in_lvl_override = true;
                    current_override_ilvl = parse_u32_attribute_value(event, b"ilvl")
                        .and_then(|v| u8::try_from(v).ok());
                } else if in_lvl_override && local == b"startOverride" {
                    if let (Some(ilvl), Some(start)) = (
                        current_override_ilvl,
                        parse_u32_attribute_value(event, b"val"),
                    ) {
                        let mut override_val = NumberingLevelOverride::new(ilvl);
                        override_val.set_start_override(start);
                        current_num_overrides.push(override_val);
                    }
                }
            }
            Event::Empty(ref event) => {
                let name_bytes = event.name();
                let local = local_name(name_bytes.as_ref());

                if current_abstract_num_id.is_some() && local == b"lvl" {
                    // Empty <w:lvl/> is unusual but handle gracefully
                    let ilvl = parse_u32_attribute_value(event, b"ilvl")
                        .and_then(|v| u8::try_from(v).ok())
                        .unwrap_or(0);
                    current_levels.push(NumberingLevel::new(ilvl, 1, "", ""));
                } else if in_lvl && local == b"start" && !in_num {
                    current_level_start = parse_u32_attribute_value(event, b"val");
                } else if in_lvl && local == b"numFmt" {
                    current_level_format = parse_attribute_value(event, b"val");
                } else if in_lvl && local == b"lvlText" {
                    current_level_text = parse_attribute_value(event, b"val");
                } else if in_lvl && local == b"lvlJc" {
                    current_level_alignment = parse_attribute_value(event, b"val");
                } else if in_lvl && local == b"suff" {
                    current_level_suffix = parse_attribute_value(event, b"val");
                } else if in_lvl_ppr && local == b"ind" {
                    current_level_indent_left = parse_u32_attribute_value(event, b"left");
                    current_level_indent_hanging = parse_u32_attribute_value(event, b"hanging");
                } else if in_lvl_tabs && local == b"tab" {
                    if current_level_tab_stop.is_none() {
                        current_level_tab_stop = parse_u32_attribute_value(event, b"pos");
                    }
                } else if in_lvl_rpr && local == b"rFonts" {
                    current_level_font_family = parse_attribute_value(event, b"ascii");
                } else if in_lvl_rpr && local == b"sz" {
                    current_level_font_size = parse_u32_attribute_value(event, b"val")
                        .and_then(|v| u16::try_from(v).ok());
                } else if in_lvl_rpr && local == b"b" {
                    let val = parse_attribute_value(event, b"val");
                    current_level_bold = Some(!matches!(val.as_deref(), Some("0") | Some("false")));
                } else if in_lvl_rpr && local == b"i" {
                    let val = parse_attribute_value(event, b"val");
                    current_level_italic =
                        Some(!matches!(val.as_deref(), Some("0") | Some("false")));
                } else if in_lvl_rpr && local == b"color" {
                    current_level_color = parse_attribute_value(event, b"val");
                } else if in_num && local == b"abstractNumId" {
                    current_num_abstract_id = parse_u32_attribute_value(event, b"val");
                } else if in_num && local == b"lvlOverride" {
                    // Empty lvlOverride without startOverride - just record the level
                    let ilvl = parse_u32_attribute_value(event, b"ilvl")
                        .and_then(|v| u8::try_from(v).ok());
                    if let Some(ilvl) = ilvl {
                        current_num_overrides.push(NumberingLevelOverride::new(ilvl));
                    }
                } else if in_lvl_override && local == b"startOverride" {
                    if let (Some(ilvl), Some(start)) = (
                        current_override_ilvl,
                        parse_u32_attribute_value(event, b"val"),
                    ) {
                        let mut override_val = NumberingLevelOverride::new(ilvl);
                        override_val.set_start_override(start);
                        current_num_overrides.push(override_val);
                    }
                }
            }
            Event::End(ref event) => {
                let name_bytes = event.name();
                let local = local_name(name_bytes.as_ref());

                if local == b"lvl" {
                    if in_lvl {
                        let ilvl = current_level_ilvl.unwrap_or(0);
                        let start = current_level_start.unwrap_or(1);
                        let format = current_level_format.take().unwrap_or_default();
                        let text = current_level_text.take().unwrap_or_default();
                        let mut level = NumberingLevel::new(ilvl, start, format, text);
                        if let Some(alignment) = current_level_alignment.take() {
                            level.set_alignment(alignment);
                        }
                        if let Some(indent_left) = current_level_indent_left.take() {
                            level.set_indent_left_twips(indent_left);
                        }
                        if let Some(indent_hanging) = current_level_indent_hanging.take() {
                            level.set_indent_hanging_twips(indent_hanging);
                        }
                        if let Some(tab_stop) = current_level_tab_stop.take() {
                            level.set_tab_stop_twips(tab_stop);
                        }
                        if let Some(suffix) = current_level_suffix.take() {
                            level.set_suffix(suffix);
                        }
                        if let Some(font_family) = current_level_font_family.take() {
                            level.set_font_family(font_family);
                        }
                        if let Some(font_size) = current_level_font_size.take() {
                            level.set_font_size_half_points(font_size);
                        }
                        if let Some(bold) = current_level_bold.take() {
                            level.set_bold(bold);
                        }
                        if let Some(italic) = current_level_italic.take() {
                            level.set_italic(italic);
                        }
                        if let Some(color) = current_level_color.take() {
                            level.set_color(color);
                        }
                        current_levels.push(level);
                        in_lvl = false;
                        in_lvl_ppr = false;
                        in_lvl_rpr = false;
                        in_lvl_tabs = false;
                    }
                } else if local == b"abstractNum" {
                    if let Some(id) = current_abstract_num_id.take() {
                        let mut def = NumberingDefinition::new(id);
                        def.set_levels(current_levels.clone());
                        definitions.push(def);
                    }
                    current_levels.clear();
                } else if in_lvl && local == b"pPr" {
                    in_lvl_ppr = false;
                    in_lvl_tabs = false;
                } else if in_lvl && local == b"rPr" {
                    in_lvl_rpr = false;
                } else if local == b"lvlOverride" {
                    in_lvl_override = false;
                    current_override_ilvl = None;
                } else if local == b"num" {
                    if let (Some(num_id), Some(abstract_id)) =
                        (current_num_id.take(), current_num_abstract_id.take())
                    {
                        let mut instance = NumberingInstance::new(num_id, abstract_id);
                        instance.set_level_overrides(std::mem::take(&mut current_num_overrides));
                        instances.push(instance);
                    }
                    in_num = false;
                }
            }
            Event::Eof => break,
            _ => {}
        }
        buffer.clear();
    }

    Ok((definitions, instances))
}

#[cfg(test)]
mod tests {
    use std::ffi::OsStr;
    use std::fs;
    use std::path::{Path, PathBuf};

    use tempfile::tempdir;

    use crate::document::{BodyItem, Document, DocumentProtection, ParsedBodyItemKind};
    use crate::image::{FloatingImage, InlineImage};
    use crate::paragraph::ParagraphAlignment;
    use crate::section::{HeaderFooter, PageOrientation};
    use crate::style::Style;
    use crate::table::{Table, TableBorder, VerticalAlignment, VerticalMerge};
    use offidized_opc::content_types::ContentTypeValue;
    use offidized_opc::relationship::{RelationshipType, TargetMode};

    const DOCX_REFERENCE_ASSETS_RELATIVE_ROOT: &str =
        "Open-XML-SDK/test/DocumentFormat.OpenXml.Tests.Assets/assets";
    const CURATED_REFERENCE_DOCX_FIXTURES: &[&str] = &[
        "TestFiles/HelloWorld.docx",
        "TestFiles/Plain.docx",
        "TestFiles/Hyperlink.docx",
        "TestFiles/mailmerge.docx",
        "TestFiles/Comments.docx",
        "TestFiles/DocProps.docx",
        "TestFiles/HelloO14.docx",
        "TestFiles/UnknownElement.docx",
        "TestFiles/simpleSdt.docx",
        "TestFiles/mcdoc.docx",
        "TestFiles/mcinleaf.docx",
        "TestFiles/Of16-09-UnknownElement.docx",
        "TestFiles/Data-Bound-Content-Controls.docx",
    ];
    const LARGE_SWEEP_REFERENCE_DOCX_DIRS: &[&str] = &[
        "TestFiles",
        "TestDataStorage/O14ISOStrict/Word",
        "TestDataStorage/O15Conformance/WD/CommentExTest",
    ];
    const LARGE_SWEEP_SUCCESS_THRESHOLD: f64 = 0.80;
    const MAX_REPORTED_SWEEP_FAILURES: usize = 64;

    type MarginsFingerprint = (
        Option<u32>,
        Option<u32>,
        Option<u32>,
        Option<u32>,
        Option<u32>,
        Option<u32>,
        Option<u32>,
    );

    #[derive(Debug, Clone, PartialEq, Eq)]
    struct RunFingerprint {
        text: String,
        style_id: Option<String>,
        hyperlink: Option<String>,
        inline_image: Option<InlineImageFingerprint>,
        floating_image: Option<FloatingImageFingerprint>,
    }

    #[derive(Debug, Clone, PartialEq, Eq)]
    struct InlineImageFingerprint {
        image_index: usize,
        width_emu: u32,
        height_emu: u32,
        name: Option<String>,
        description: Option<String>,
    }

    #[derive(Debug, Clone, PartialEq, Eq)]
    struct FloatingImageFingerprint {
        image_index: usize,
        width_emu: u32,
        height_emu: u32,
        offset_x_emu: i32,
        offset_y_emu: i32,
        name: Option<String>,
        description: Option<String>,
    }

    #[derive(Debug, Clone, PartialEq, Eq)]
    struct ParagraphFingerprint {
        text: String,
        heading_level: Option<u8>,
        style_id: Option<String>,
        alignment: Option<ParagraphAlignment>,
        numbering_num_id: Option<u32>,
        numbering_ilvl: Option<u8>,
        runs: Vec<RunFingerprint>,
    }

    #[derive(Debug, Clone, PartialEq, Eq)]
    struct TableCellFingerprint {
        text: String,
        horizontal_span: usize,
        horizontal_merge_continuation: bool,
        vertical_merge: Option<VerticalMerge>,
        shading_color: Option<String>,
        vertical_alignment: Option<VerticalAlignment>,
        cell_width_twips: Option<u32>,
    }

    #[derive(Debug, Clone, PartialEq, Eq)]
    struct TableFingerprint {
        rows: usize,
        columns: usize,
        style_id: Option<String>,
        cells: Vec<TableCellFingerprint>,
    }

    #[derive(Debug, Clone, PartialEq, Eq)]
    struct ImageFingerprint {
        content_type: String,
        bytes_len: usize,
        bytes_checksum: u64,
    }

    #[derive(Debug, Clone, PartialEq, Eq)]
    struct SectionFingerprint {
        page_width_twips: Option<u32>,
        page_height_twips: Option<u32>,
        page_orientation: Option<PageOrientation>,
        margins: MarginsFingerprint,
        header_paragraphs: Vec<String>,
        footer_paragraphs: Vec<String>,
    }

    #[derive(Debug, Clone, PartialEq, Eq)]
    struct DocumentFingerprint {
        body_item_kinds: Vec<ParsedBodyItemKind>,
        paragraphs: Vec<ParagraphFingerprint>,
        tables: Vec<TableFingerprint>,
        images: Vec<ImageFingerprint>,
        section: SectionFingerprint,
    }

    fn workspace_root() -> PathBuf {
        let manifest_dir = Path::new(env!("CARGO_MANIFEST_DIR"));
        manifest_dir
            .parent()
            .and_then(Path::parent)
            .expect("crate should live under workspace_root/crates/<crate_name>")
            .to_path_buf()
    }

    fn reference_assets_root() -> PathBuf {
        workspace_root().join(DOCX_REFERENCE_ASSETS_RELATIVE_ROOT)
    }

    fn reference_fixture_path(relative_fixture_path: &str) -> PathBuf {
        reference_assets_root().join(relative_fixture_path)
    }

    fn fixture_display_path(path: &Path) -> String {
        let root = reference_assets_root();
        path.strip_prefix(&root)
            .unwrap_or(path)
            .to_string_lossy()
            .replace('\\', "/")
    }

    fn bytes_checksum(bytes: &[u8]) -> u64 {
        bytes.iter().fold(0_u64, |checksum, byte| {
            checksum
                .wrapping_mul(16_777_619)
                .wrapping_add(u64::from(*byte) + 1)
        })
    }

    fn document_fingerprint(document: &Document) -> DocumentFingerprint {
        let body_item_kinds = document
            .body_items()
            .map(|item| match item {
                BodyItem::Paragraph(_) => ParsedBodyItemKind::Paragraph,
                BodyItem::Table(_) => ParsedBodyItemKind::Table,
            })
            .collect::<Vec<_>>();

        let paragraphs = document
            .paragraphs()
            .iter()
            .map(|paragraph| ParagraphFingerprint {
                text: paragraph.text(),
                heading_level: paragraph.heading_level(),
                style_id: paragraph.style_id().map(str::to_owned),
                alignment: paragraph.alignment(),
                numbering_num_id: paragraph.numbering_num_id(),
                numbering_ilvl: paragraph.numbering_ilvl(),
                runs: paragraph
                    .runs()
                    .iter()
                    .map(|run| RunFingerprint {
                        text: run.text().to_string(),
                        style_id: run.style_id().map(str::to_owned),
                        hyperlink: run.hyperlink().map(str::to_owned),
                        inline_image: run.inline_image().map(|inline| InlineImageFingerprint {
                            image_index: inline.image_index(),
                            width_emu: inline.width_emu(),
                            height_emu: inline.height_emu(),
                            name: inline.name().map(str::to_owned),
                            description: inline.description().map(str::to_owned),
                        }),
                        floating_image: run.floating_image().map(|floating| {
                            FloatingImageFingerprint {
                                image_index: floating.image_index(),
                                width_emu: floating.width_emu(),
                                height_emu: floating.height_emu(),
                                offset_x_emu: floating.offset_x_emu(),
                                offset_y_emu: floating.offset_y_emu(),
                                name: floating.name().map(str::to_owned),
                                description: floating.description().map(str::to_owned),
                            }
                        }),
                    })
                    .collect(),
            })
            .collect::<Vec<_>>();

        let tables = document
            .tables()
            .iter()
            .map(|table| {
                let cells = (0..table.rows())
                    .flat_map(|row| (0..table.columns()).map(move |column| (row, column)))
                    .map(|(row, column)| {
                        let cell = table
                            .cell(row, column)
                            .expect("cell iteration should stay in bounds");
                        TableCellFingerprint {
                            text: cell.text().to_string(),
                            horizontal_span: cell.horizontal_span(),
                            horizontal_merge_continuation: cell.is_horizontal_merge_continuation(),
                            vertical_merge: cell.vertical_merge(),
                            shading_color: cell.shading_color().map(str::to_owned),
                            vertical_alignment: cell.vertical_alignment(),
                            cell_width_twips: cell.cell_width_twips(),
                        }
                    })
                    .collect::<Vec<_>>();

                TableFingerprint {
                    rows: table.rows(),
                    columns: table.columns(),
                    style_id: table.style_id().map(str::to_owned),
                    cells,
                }
            })
            .collect::<Vec<_>>();

        let images = document
            .images()
            .iter()
            .map(|image| ImageFingerprint {
                content_type: image.content_type().to_string(),
                bytes_len: image.bytes().len(),
                bytes_checksum: bytes_checksum(image.bytes()),
            })
            .collect::<Vec<_>>();

        let section = SectionFingerprint {
            page_width_twips: document.section().page_width_twips(),
            page_height_twips: document.section().page_height_twips(),
            page_orientation: document.section().page_orientation(),
            margins: (
                document.section().page_margins().top_twips(),
                document.section().page_margins().right_twips(),
                document.section().page_margins().bottom_twips(),
                document.section().page_margins().left_twips(),
                document.section().page_margins().header_twips(),
                document.section().page_margins().footer_twips(),
                document.section().page_margins().gutter_twips(),
            ),
            header_paragraphs: document
                .section()
                .header()
                .map(|header| header.paragraphs().iter().map(|p| p.text()).collect())
                .unwrap_or_default(),
            footer_paragraphs: document
                .section()
                .footer()
                .map(|footer| footer.paragraphs().iter().map(|p| p.text()).collect())
                .unwrap_or_default(),
        };

        DocumentFingerprint {
            body_item_kinds,
            paragraphs,
            tables,
            images,
            section,
        }
    }

    fn roundtrip_fixture_and_compare_fingerprint(path: &Path) -> Result<(), String> {
        let opened = Document::open(path).map_err(|error| format!("open failed: {error}"))?;
        let before = document_fingerprint(&opened);

        let temp = tempdir().map_err(|error| format!("create temp dir failed: {error}"))?;
        let rewritten_path = temp.path().join("rewritten.docx");
        opened
            .save(&rewritten_path)
            .map_err(|error| format!("save failed: {error}"))?;

        let reopened = Document::open(&rewritten_path)
            .map_err(|error| format!("reopen rewritten file failed: {error}"))?;
        let after = document_fingerprint(&reopened);

        if before != after {
            return Err("fingerprint changed after roundtrip".to_string());
        }

        Ok(())
    }

    fn diff_document_fingerprint(
        before: &DocumentFingerprint,
        after: &DocumentFingerprint,
    ) -> Vec<String> {
        let mut diffs = Vec::new();

        if before.body_item_kinds != after.body_item_kinds {
            diffs.push(format!(
                "body_item_kinds changed: {:?} -> {:?}",
                before.body_item_kinds, after.body_item_kinds
            ));
        }

        if before.paragraphs.len() != after.paragraphs.len() {
            diffs.push(format!(
                "paragraph count changed: {} -> {}",
                before.paragraphs.len(),
                after.paragraphs.len()
            ));
        }
        if let Some(index) = before
            .paragraphs
            .iter()
            .zip(after.paragraphs.iter())
            .position(|(left, right)| left != right)
        {
            diffs.push(format!(
                "paragraph[{index}] changed: {:?} -> {:?}",
                before.paragraphs[index], after.paragraphs[index]
            ));
        }

        if before.tables.len() != after.tables.len() {
            diffs.push(format!(
                "table count changed: {} -> {}",
                before.tables.len(),
                after.tables.len()
            ));
        }
        if let Some(index) = before
            .tables
            .iter()
            .zip(after.tables.iter())
            .position(|(left, right)| left != right)
        {
            diffs.push(format!(
                "table[{index}] changed: {:?} -> {:?}",
                before.tables[index], after.tables[index]
            ));
        }

        if before.images != after.images {
            diffs.push(format!(
                "images changed: {:?} -> {:?}",
                before.images, after.images
            ));
        }

        if before.section != after.section {
            diffs.push(format!(
                "section changed: {:?} -> {:?}",
                before.section, after.section
            ));
        }

        diffs
    }

    fn collect_docx_fixtures_from_relative_dir(relative_dir: &str) -> Vec<PathBuf> {
        let start = reference_fixture_path(relative_dir);
        if !start.exists() {
            return Vec::new();
        }

        let mut collected = Vec::new();
        let mut stack = vec![start];
        while let Some(path) = stack.pop() {
            let entries =
                fs::read_dir(&path).unwrap_or_else(|error| panic!("read_dir failed: {error}"));
            for entry in entries {
                let entry = entry.unwrap_or_else(|error| panic!("read_dir entry failed: {error}"));
                let entry_path = entry.path();
                if entry_path.is_dir() {
                    stack.push(entry_path);
                    continue;
                }

                let is_docx = entry_path
                    .extension()
                    .and_then(OsStr::to_str)
                    .is_some_and(|extension| extension.eq_ignore_ascii_case("docx"));
                if is_docx {
                    collected.push(entry_path);
                }
            }
        }

        collected.sort();
        collected
    }

    fn large_sweep_success_threshold() -> f64 {
        std::env::var("OFFIDIZED_DOCX_SWEEP_SUCCESS_THRESHOLD")
            .ok()
            .and_then(|raw| raw.parse::<f64>().ok())
            .filter(|threshold| (0.0..=1.0).contains(threshold))
            .unwrap_or(LARGE_SWEEP_SUCCESS_THRESHOLD)
    }

    #[test]
    fn save_writes_minimal_wordprocessing_parts() {
        let dir = tempdir().expect("create temp dir");
        let path = dir.path().join("minimal.docx");

        let mut document = Document::new();
        document.add_heading("Heading", 1);
        document.add_paragraph("Body");

        document.save(&path).expect("save docx");

        let package = offidized_opc::Package::open(&path).expect("open package");
        let office_document_relationship = package
            .relationships()
            .get_first_by_type(RelationshipType::WORD_DOCUMENT)
            .expect("word officeDocument relationship");
        assert_eq!(office_document_relationship.target, "word/document.xml");
        assert!(package.get_part("/word/document.xml").is_some());
        assert_eq!(
            package.content_types().get_override("/word/document.xml"),
            Some(ContentTypeValue::WORD_DOCUMENT)
        );
    }

    #[test]
    fn open_parses_roundtrip_paragraph_runs_and_heading_style() {
        let dir = tempdir().expect("create temp dir");
        let path = dir.path().join("roundtrip.docx");

        let mut document = Document::new();
        document.add_heading("Heading", 1);
        let paragraph = document.add_paragraph("Hello");
        let run = paragraph.add_run(" world");
        run.set_bold(true);
        run.set_italic(true);
        run.set_underline(true);
        run.set_font_family("Calibri");
        run.set_font_size_half_points(28);
        run.set_color("#3a5fcd");

        document.save(&path).expect("save roundtrip docx");

        let reopened = Document::open(&path).expect("open roundtrip docx");
        assert_eq!(reopened.paragraphs().len(), 2);
        assert_eq!(reopened.paragraphs()[0].text(), "Heading");
        assert_eq!(reopened.paragraphs()[0].heading_level(), Some(1));
        assert_eq!(reopened.paragraphs()[0].style_id(), Some("Heading1"));
        assert_eq!(reopened.paragraphs()[1].text(), "Hello world");
        assert_eq!(reopened.paragraphs()[1].runs().len(), 2);
        assert!(!reopened.paragraphs()[1].runs()[0].is_bold());
        assert!(reopened.paragraphs()[1].runs()[1].is_bold());
        assert!(reopened.paragraphs()[1].runs()[1].is_italic());
        assert!(reopened.paragraphs()[1].runs()[1].is_underline());
        assert_eq!(
            reopened.paragraphs()[1].runs()[1].font_family(),
            Some("Calibri")
        );
        assert_eq!(
            reopened.paragraphs()[1].runs()[1].font_size_half_points(),
            Some(28)
        );
        assert_eq!(reopened.paragraphs()[1].runs()[1].color(), Some("3A5FCD"));
    }

    #[test]
    fn open_save_roundtrips_paragraph_formatting() {
        let dir = tempdir().expect("create temp dir");
        let path = dir.path().join("paragraph-formatting-roundtrip.docx");

        let mut document = Document::new();
        let paragraph = document.add_paragraph("Formatted paragraph");
        paragraph.set_alignment(ParagraphAlignment::Center);
        paragraph.set_spacing_before_twips(120);
        paragraph.set_spacing_after_twips(80);
        paragraph.set_line_spacing_twips(360);
        paragraph.set_indent_left_twips(300);
        paragraph.set_indent_right_twips(120);
        paragraph.set_indent_first_line_twips(240);
        paragraph.set_indent_hanging_twips(90);

        document
            .save(&path)
            .expect("save paragraph formatting docx");

        let reopened = Document::open(&path).expect("open paragraph formatting docx");
        assert_eq!(reopened.paragraphs().len(), 1);
        let formatted = &reopened.paragraphs()[0];
        assert_eq!(formatted.text(), "Formatted paragraph");
        assert_eq!(formatted.alignment(), Some(ParagraphAlignment::Center));
        assert_eq!(formatted.spacing_before_twips(), Some(120));
        assert_eq!(formatted.spacing_after_twips(), Some(80));
        assert_eq!(formatted.line_spacing_twips(), Some(360));
        assert_eq!(formatted.indent_left_twips(), Some(300));
        assert_eq!(formatted.indent_right_twips(), Some(120));
        assert_eq!(formatted.indent_first_line_twips(), Some(240));
        assert_eq!(formatted.indent_hanging_twips(), Some(90));
    }

    #[test]
    fn open_save_roundtrips_paragraph_numbering() {
        let dir = tempdir().expect("create temp dir");
        let path = dir.path().join("paragraph-numbering-roundtrip.docx");

        let mut document = Document::new();
        let first = document.add_paragraph("First item");
        first.set_numbering(7, 0);
        let second = document.add_paragraph("Nested item");
        second.set_numbering(7, 1);
        let plain = document.add_paragraph("Plain paragraph");
        plain.clear_numbering();

        document.save(&path).expect("save paragraph numbering docx");

        let reopened = Document::open(&path).expect("open paragraph numbering docx");
        assert_eq!(reopened.paragraphs().len(), 3);
        assert_eq!(reopened.paragraphs()[0].numbering_num_id(), Some(7));
        assert_eq!(reopened.paragraphs()[0].numbering_ilvl(), Some(0));
        assert_eq!(reopened.paragraphs()[1].numbering_num_id(), Some(7));
        assert_eq!(reopened.paragraphs()[1].numbering_ilvl(), Some(1));
        assert_eq!(reopened.paragraphs()[2].numbering_num_id(), None);
        assert_eq!(reopened.paragraphs()[2].numbering_ilvl(), None);
    }

    #[test]
    fn open_save_roundtrips_hyperlinks_with_deterministic_relationships() {
        let dir = tempdir().expect("create temp dir");
        let path = dir.path().join("hyperlink-roundtrip.docx");

        let mut document = Document::new();
        let paragraph = document.add_paragraph("Visit ");
        paragraph.add_hyperlink("example", "https://example.com");
        paragraph.add_hyperlink(" docs", "https://example.com");
        paragraph.add_run(" and ");
        paragraph.add_hyperlink("Rust", "https://www.rust-lang.org");

        document.save(&path).expect("save hyperlink docx");

        let package = offidized_opc::Package::open(&path).expect("open hyperlink package");
        let document_part = package
            .get_part("/word/document.xml")
            .expect("word document part");

        let hyperlink_relationships = document_part
            .relationships
            .get_by_type(RelationshipType::HYPERLINK);
        assert_eq!(hyperlink_relationships.len(), 2);
        assert_eq!(hyperlink_relationships[0].id, "rId1");
        assert_eq!(hyperlink_relationships[0].target, "https://example.com");
        assert_eq!(hyperlink_relationships[0].target_mode, TargetMode::External);
        assert_eq!(hyperlink_relationships[1].id, "rId2");
        assert_eq!(
            hyperlink_relationships[1].target,
            "https://www.rust-lang.org"
        );
        assert_eq!(hyperlink_relationships[1].target_mode, TargetMode::External);

        let reopened = Document::open(&path).expect("open hyperlink docx");
        let runs = reopened.paragraphs()[0].runs();
        assert_eq!(runs.len(), 5);
        assert_eq!(runs[0].text(), "Visit ");
        assert_eq!(runs[1].hyperlink(), Some("https://example.com"));
        assert_eq!(runs[2].hyperlink(), Some("https://example.com"));
        assert_eq!(runs[3].hyperlink(), None);
        assert_eq!(runs[4].hyperlink(), Some("https://www.rust-lang.org"));
    }

    #[test]
    fn open_save_roundtrips_inline_images_with_deterministic_media_relationships() {
        let dir = tempdir().expect("create temp dir");
        let path = dir.path().join("inline-image-roundtrip.docx");

        let mut document = Document::new();
        let first_image_index = document.add_image(vec![0_u8, 1, 2, 3], "image/png");
        let second_image_index = document.add_image(vec![4_u8, 5, 6], "image/jpeg");

        let paragraph = document.add_paragraph("Images: ");
        let first_run = paragraph.add_inline_image(first_image_index, 990_000, 792_000);
        let mut first_inline = InlineImage::new(first_image_index, 990_000, 792_000);
        first_inline.set_name("Picture 1");
        first_inline.set_description("first image");
        first_run.set_inline_image(first_inline);
        paragraph.add_run(" and ");
        paragraph.add_inline_image(second_image_index, 720_000, 540_000);

        document.save(&path).expect("save inline image docx");

        let package = offidized_opc::Package::open(&path).expect("open inline image package");
        assert!(package.get_part("/word/media/image1.png").is_some());
        assert!(package.get_part("/word/media/image2.jpeg").is_some());
        assert_eq!(
            package
                .content_types()
                .get_override("/word/media/image1.png"),
            Some("image/png")
        );
        assert_eq!(
            package
                .content_types()
                .get_override("/word/media/image2.jpeg"),
            Some("image/jpeg")
        );

        let document_part = package
            .get_part("/word/document.xml")
            .expect("word document part");
        let image_relationships = document_part
            .relationships
            .get_by_type(RelationshipType::IMAGE);
        assert_eq!(image_relationships.len(), 2);
        assert_eq!(image_relationships[0].id, "rId1");
        assert_eq!(image_relationships[0].target, "media/image1.png");
        assert_eq!(image_relationships[0].target_mode, TargetMode::Internal);
        assert_eq!(image_relationships[1].id, "rId2");
        assert_eq!(image_relationships[1].target, "media/image2.jpeg");
        assert_eq!(image_relationships[1].target_mode, TargetMode::Internal);

        let reopened = Document::open(&path).expect("open inline image docx");
        assert_eq!(reopened.images().len(), 2);
        assert_eq!(reopened.images()[0].content_type(), "image/png");
        assert_eq!(reopened.images()[1].content_type(), "image/jpeg");

        let runs = reopened.paragraphs()[0].runs();
        assert_eq!(runs.len(), 4);
        assert_eq!(runs[0].text(), "Images: ");
        assert_eq!(runs[2].text(), " and ");
        assert_eq!(
            runs[1].inline_image().map(InlineImage::width_emu),
            Some(990_000)
        );
        assert_eq!(
            runs[1].inline_image().and_then(InlineImage::name),
            Some("Picture 1")
        );
        assert_eq!(
            runs[1].inline_image().and_then(InlineImage::description),
            Some("first image")
        );
        assert_eq!(
            runs[1].inline_image().map(InlineImage::image_index),
            Some(0)
        );
        assert_eq!(
            runs[3].inline_image().map(InlineImage::height_emu),
            Some(540_000)
        );
        assert_eq!(
            runs[3].inline_image().map(InlineImage::image_index),
            Some(1)
        );
    }

    #[test]
    fn open_save_roundtrips_images_with_deterministic_indexes_under_nontrivial_run_order() {
        let dir = tempdir().expect("create temp dir");
        let path = dir.path().join("nontrivial-image-order-roundtrip.docx");

        let mut document = Document::new();
        let first_image_index = document.add_image(vec![1_u8, 2, 3], "image/png");
        let second_image_index = document.add_image(vec![4_u8, 5, 6], "image/jpeg");
        let third_image_index = document.add_image(vec![7_u8, 8, 9], "image/gif");

        let first_paragraph = document.add_paragraph("First ");
        first_paragraph.add_inline_image(third_image_index, 420_000, 360_000);
        first_paragraph.add_run(" then ");
        first_paragraph.add_inline_image(first_image_index, 510_000, 480_000);

        let second_paragraph = document.add_paragraph("Second ");
        let floating_run = second_paragraph.add_floating_image(third_image_index, 300_000, 240_000);
        let mut floating_image = FloatingImage::new(third_image_index, 300_000, 240_000);
        floating_image.set_offsets_emu(111_000, -222_000);
        floating_run.set_floating_image(floating_image);
        second_paragraph.add_run(" and ");
        second_paragraph.add_inline_image(second_image_index, 330_000, 270_000);

        document
            .save(&path)
            .expect("save nontrivial image order docx");

        let package = offidized_opc::Package::open(&path).expect("open nontrivial image package");
        let document_part = package
            .get_part("/word/document.xml")
            .expect("word document part");
        let image_relationships = document_part
            .relationships
            .get_by_type(RelationshipType::IMAGE);
        assert_eq!(image_relationships.len(), 3);
        assert_eq!(image_relationships[0].target, "media/image1.png");
        assert_eq!(image_relationships[1].target, "media/image2.jpeg");
        assert_eq!(image_relationships[2].target, "media/image3.gif");

        let reopened = Document::open(&path).expect("reopen nontrivial image order docx");
        assert_eq!(reopened.images().len(), 3);
        assert_eq!(reopened.images()[0].content_type(), "image/png");
        assert_eq!(reopened.images()[1].content_type(), "image/jpeg");
        assert_eq!(reopened.images()[2].content_type(), "image/gif");

        let first_runs = reopened.paragraphs()[0].runs();
        assert_eq!(
            first_runs[1].inline_image().map(InlineImage::image_index),
            Some(2)
        );
        assert_eq!(
            first_runs[3].inline_image().map(InlineImage::image_index),
            Some(0)
        );

        let second_runs = reopened.paragraphs()[1].runs();
        assert_eq!(
            second_runs[1]
                .floating_image()
                .map(FloatingImage::image_index),
            Some(2)
        );
        assert_eq!(
            second_runs[3].inline_image().map(InlineImage::image_index),
            Some(1)
        );
    }

    #[test]
    fn open_save_open_preserves_unused_document_images_without_bloating_header_footer() {
        let dir = tempdir().expect("create temp dir");
        let path = dir.path().join("unused-images-roundtrip.docx");
        let rewritten_path = dir.path().join("unused-images-roundtrip-rewritten.docx");

        let mut document = Document::new();
        let used_image_index = document.add_image(vec![10_u8, 11, 12, 13], "image/png");
        let _unused_jpeg_index = document.add_image(vec![14_u8, 15, 16], "image/jpeg");
        let _unused_svg_index = document.add_image(vec![17_u8, 18, 19, 20], "image/svg+xml");

        document
            .add_paragraph("Body image ")
            .add_inline_image(used_image_index, 250_000, 200_000);

        document
            .section_mut()
            .set_header(HeaderFooter::from_text("Header text"));
        document
            .section_mut()
            .set_footer(HeaderFooter::from_text("Footer text"));

        document.save(&path).expect("save unused image docx");

        let opened = Document::open(&path).expect("open unused image docx");
        let opened_image_fingerprint = opened
            .images()
            .iter()
            .map(|image| {
                (
                    image.content_type().to_string(),
                    image.bytes().len(),
                    bytes_checksum(image.bytes()),
                )
            })
            .collect::<Vec<_>>();
        assert_eq!(opened_image_fingerprint.len(), 3);

        opened
            .save(&rewritten_path)
            .expect("resave unused image docx");

        let reopened = Document::open(&rewritten_path).expect("reopen unused image docx");
        let reopened_image_fingerprint = reopened
            .images()
            .iter()
            .map(|image| {
                (
                    image.content_type().to_string(),
                    image.bytes().len(),
                    bytes_checksum(image.bytes()),
                )
            })
            .collect::<Vec<_>>();
        assert_eq!(reopened_image_fingerprint, opened_image_fingerprint);

        let package =
            offidized_opc::Package::open(&rewritten_path).expect("open rewritten unused package");
        let document_part = package
            .get_part("/word/document.xml")
            .expect("word document part");
        assert_eq!(
            document_part
                .relationships
                .get_by_type(RelationshipType::IMAGE)
                .len(),
            3
        );

        let header_part = package
            .get_part("/word/header1.xml")
            .expect("word header part");
        assert_eq!(
            header_part
                .relationships
                .get_by_type(RelationshipType::IMAGE)
                .len(),
            0
        );

        let footer_part = package
            .get_part("/word/footer1.xml")
            .expect("word footer part");
        assert_eq!(
            footer_part
                .relationships
                .get_by_type(RelationshipType::IMAGE)
                .len(),
            0
        );
    }

    #[test]
    fn open_save_roundtrip_officeimo_document_with_images_preserves_image_counts_and_content_types()
    {
        let fixture_path = workspace_root()
            .join("references/OfficeIMO/OfficeIMO.Tests/Documents/DocumentWithImages.docx");
        if !fixture_path.is_file() {
            eprintln!(
                "skipping test: OfficeIMO image fixture not found at `{}`",
                fixture_path.display()
            );
            return;
        }

        let expected_content_types_sorted = vec![
            "image/jpeg".to_string(),
            "image/jpeg".to_string(),
            "image/jpeg".to_string(),
            "image/jpeg".to_string(),
            "image/jpeg".to_string(),
            "image/png".to_string(),
            "image/png".to_string(),
        ];

        let opened = Document::open(&fixture_path).expect("open OfficeIMO image fixture");
        let mut opened_content_types = opened
            .images()
            .iter()
            .map(|image| image.content_type().to_string())
            .collect::<Vec<_>>();
        assert_eq!(opened.images().len(), 7);
        opened_content_types.sort();
        assert_eq!(opened_content_types, expected_content_types_sorted);

        let dir = tempdir().expect("create temp dir");
        let rewritten_path = dir.path().join("officeimo-images-roundtrip.docx");
        opened
            .save(&rewritten_path)
            .expect("save OfficeIMO image roundtrip docx");

        let reopened = Document::open(&rewritten_path).expect("open rewritten OfficeIMO docx");
        let mut reopened_content_types = reopened
            .images()
            .iter()
            .map(|image| image.content_type().to_string())
            .collect::<Vec<_>>();
        assert_eq!(reopened.images().len(), 7);
        reopened_content_types.sort();
        assert_eq!(reopened_content_types, expected_content_types_sorted);
    }

    #[test]
    fn open_save_roundtrips_floating_images_with_anchor_mode() {
        let dir = tempdir().expect("create temp dir");
        let path = dir.path().join("floating-image-roundtrip.docx");

        let mut document = Document::new();
        let floating_image_index = document.add_image(vec![1_u8, 2, 3, 4], "image/png");
        let inline_image_index = document.add_image(vec![5_u8, 6, 7], "image/jpeg");

        let paragraph = document.add_paragraph("Float ");
        let floating_run = paragraph.add_floating_image(floating_image_index, 990_000, 792_000);
        let mut floating_image = FloatingImage::new(floating_image_index, 990_000, 792_000);
        floating_image.set_offsets_emu(222_000, -111_000);
        floating_image.set_name("Anchored image");
        floating_image.set_description("floating image");
        floating_run.set_floating_image(floating_image);
        paragraph.add_run(" + inline ");
        paragraph.add_inline_image(inline_image_index, 720_000, 540_000);

        document.save(&path).expect("save floating image docx");

        let package = offidized_opc::Package::open(&path).expect("open floating image package");
        let document_part = package
            .get_part("/word/document.xml")
            .expect("word document part");
        let document_xml = String::from_utf8_lossy(document_part.data.as_bytes());
        assert!(document_xml.contains("<wp:anchor"));
        assert!(document_xml.contains("<wp:simplePos"));
        assert!(document_xml.contains("<wp:positionH"));
        assert!(document_xml.contains("<wp:positionV"));
        let image_relationships = document_part
            .relationships
            .get_by_type(RelationshipType::IMAGE);
        assert_eq!(image_relationships.len(), 2);

        let reopened = Document::open(&path).expect("open floating image docx");
        let runs = reopened.paragraphs()[0].runs();
        assert_eq!(runs.len(), 4);
        assert_eq!(runs[0].text(), "Float ");
        assert_eq!(runs[2].text(), " + inline ");
        assert_eq!(runs[1].inline_image(), None);
        assert_eq!(
            runs[1].floating_image().map(FloatingImage::width_emu),
            Some(990_000)
        );
        assert_eq!(
            runs[1].floating_image().map(FloatingImage::height_emu),
            Some(792_000)
        );
        assert_eq!(
            runs[1].floating_image().map(FloatingImage::offset_x_emu),
            Some(222_000)
        );
        assert_eq!(
            runs[1].floating_image().map(FloatingImage::offset_y_emu),
            Some(-111_000)
        );
        assert_eq!(
            runs[1].floating_image().and_then(FloatingImage::name),
            Some("Anchored image")
        );
        assert_eq!(
            runs[1]
                .floating_image()
                .and_then(FloatingImage::description),
            Some("floating image")
        );
        assert_eq!(
            runs[1].floating_image().map(FloatingImage::image_index),
            Some(0)
        );
        assert_eq!(
            runs[3].inline_image().map(InlineImage::image_index),
            Some(1)
        );
    }

    #[test]
    fn open_save_roundtrips_table_text() {
        let dir = tempdir().expect("create temp dir");
        let path = dir.path().join("table-roundtrip.docx");

        let mut document = Document::new();
        let table = document.add_table(2, 2);
        assert!(table.set_cell_text(0, 0, "A1"));
        assert!(table.set_cell_text(0, 1, "A2"));
        assert!(table.set_cell_text(1, 0, "B1"));
        assert!(!table.set_cell_text(2, 0, "out-of-bounds"));

        document.save(&path).expect("save table roundtrip docx");

        let reopened = Document::open(&path).expect("open table roundtrip docx");
        assert_eq!(reopened.tables().len(), 1);
        assert_eq!(reopened.tables()[0].rows(), 2);
        assert_eq!(reopened.tables()[0].columns(), 2);
        assert_eq!(reopened.tables()[0].cell_text(0, 0), Some("A1"));
        assert_eq!(reopened.tables()[0].cell_text(0, 1), Some("A2"));
        assert_eq!(reopened.tables()[0].cell_text(1, 0), Some("B1"));
        assert_eq!(reopened.tables()[0].cell_text(1, 1), Some(""));
        assert_eq!(reopened.tables()[0].cell_text(3, 3), None);
    }

    #[test]
    fn insert_table_at_preserves_body_order() {
        let mut document = Document::new();
        document.add_paragraph("Before");
        document.add_paragraph("After");

        let mut table = Table::new(1, 1);
        assert!(table.set_cell_text(0, 0, "Cell"));
        document.insert_table_at(1, table);

        assert_eq!(document.body_position_of_paragraph(0), Some(0));
        assert_eq!(document.body_position_of_table(0), Some(1));
        assert_eq!(document.body_position_of_paragraph(1), Some(2));
        assert_eq!(document.tables()[0].cell_text(0, 0), Some("Cell"));
    }

    #[test]
    fn open_save_roundtrips_section_properties() {
        let dir = tempdir().expect("create temp dir");
        let path = dir.path().join("section-roundtrip.docx");

        let mut document = Document::new();
        document.add_paragraph("Sectioned");
        let section = document.section_mut();
        section.set_page_size_twips(15_840, 12_240);
        section.set_page_orientation(PageOrientation::Landscape);
        let margins = section.page_margins_mut();
        margins.set_top_twips(1_080);
        margins.set_right_twips(720);
        margins.set_bottom_twips(1_080);
        margins.set_left_twips(720);
        margins.set_header_twips(360);
        margins.set_footer_twips(360);
        margins.set_gutter_twips(100);

        document.save(&path).expect("save section roundtrip docx");

        let package = offidized_opc::Package::open(&path).expect("open section package");
        let document_part = package
            .get_part("/word/document.xml")
            .expect("word document part");
        let document_xml = String::from_utf8_lossy(document_part.data.as_bytes());
        assert!(document_xml.contains("<w:sectPr>"));
        assert!(document_xml.contains("<w:pgSz"));
        assert!(document_xml.contains("w:orient=\"landscape\""));
        assert!(document_xml.contains("<w:pgMar"));

        let reopened = Document::open(&path).expect("open section roundtrip docx");
        assert_eq!(reopened.section().page_width_twips(), Some(15_840));
        assert_eq!(reopened.section().page_height_twips(), Some(12_240));
        assert_eq!(
            reopened.section().page_orientation(),
            Some(PageOrientation::Landscape)
        );
        assert_eq!(reopened.section().page_margins().top_twips(), Some(1_080));
        assert_eq!(reopened.section().page_margins().right_twips(), Some(720));
        assert_eq!(
            reopened.section().page_margins().bottom_twips(),
            Some(1_080)
        );
        assert_eq!(reopened.section().page_margins().left_twips(), Some(720));
        assert_eq!(reopened.section().page_margins().header_twips(), Some(360));
        assert_eq!(reopened.section().page_margins().footer_twips(), Some(360));
        assert_eq!(reopened.section().page_margins().gutter_twips(), Some(100));
    }

    #[test]
    fn open_save_roundtrips_section_headers_and_footers_with_references() {
        let dir = tempdir().expect("create temp dir");
        let path = dir.path().join("section-header-footer-roundtrip.docx");
        let rewritten_path = dir
            .path()
            .join("section-header-footer-roundtrip-rewritten.docx");

        let mut document = Document::new();
        document.add_paragraph("Body paragraph");
        let mut header = HeaderFooter::new();
        header.add_paragraph("Header");
        header.add_paragraph("Page title");
        let mut footer = HeaderFooter::from_text("Footer");
        footer
            .paragraphs_mut()
            .first_mut()
            .expect("footer paragraph")
            .add_run(" text");
        document.section_mut().set_header(header);
        document.section_mut().set_footer(footer);

        document
            .save(&path)
            .expect("save section header/footer docx");

        let package =
            offidized_opc::Package::open(&path).expect("open section header/footer package");
        assert!(package.get_part("/word/header1.xml").is_some());
        assert!(package.get_part("/word/footer1.xml").is_some());
        assert_eq!(
            package.content_types().get_override("/word/header1.xml"),
            Some(super::WORD_HEADER_CONTENT_TYPE)
        );
        assert_eq!(
            package.content_types().get_override("/word/footer1.xml"),
            Some(super::WORD_FOOTER_CONTENT_TYPE)
        );
        let document_part = package
            .get_part("/word/document.xml")
            .expect("word document part");
        let header_relationships = document_part
            .relationships
            .iter()
            .filter(|relationship| relationship.rel_type == super::WORD_HEADER_REL_TYPE)
            .collect::<Vec<_>>();
        assert_eq!(header_relationships.len(), 1);
        assert_eq!(header_relationships[0].target, "header1.xml");
        let footer_relationships = document_part
            .relationships
            .iter()
            .filter(|relationship| relationship.rel_type == super::WORD_FOOTER_REL_TYPE)
            .collect::<Vec<_>>();
        assert_eq!(footer_relationships.len(), 1);
        assert_eq!(footer_relationships[0].target, "footer1.xml");
        let document_xml = String::from_utf8_lossy(document_part.data.as_bytes());
        assert!(document_xml.contains("<w:headerReference"));
        assert!(document_xml.contains("<w:footerReference"));

        let reopened = Document::open(&path).expect("open section header/footer docx");
        assert_eq!(
            reopened
                .section()
                .header()
                .map(|header| header.paragraphs().len()),
            Some(2)
        );
        assert_eq!(
            reopened
                .section()
                .header()
                .and_then(|header| header.paragraphs().first())
                .map(|paragraph| paragraph.text()),
            Some("Header".to_string())
        );
        assert_eq!(
            reopened
                .section()
                .footer()
                .and_then(|footer| footer.paragraphs().first())
                .map(|paragraph| paragraph.text()),
            Some("Footer text".to_string())
        );

        reopened
            .save(&rewritten_path)
            .expect("resave section header/footer docx");
        let rewritten_package =
            offidized_opc::Package::open(&rewritten_path).expect("open rewritten package");
        assert!(rewritten_package.get_part("/word/header1.xml").is_some());
        assert!(rewritten_package.get_part("/word/footer1.xml").is_some());
    }

    #[test]
    fn open_save_roundtrips_style_ids_and_styles_part() {
        let dir = tempdir().expect("create temp dir");
        let path = dir.path().join("styles-roundtrip.docx");
        let rewritten_path = dir.path().join("styles-roundtrip-rewritten.docx");

        let mut document = Document::new();
        let body_style = document.styles_mut().add_paragraph_style("BodyText");
        body_style.set_name("Body Text");
        body_style.set_paragraph_properties_xml(
            "<w:pPr><w:spacing w:before=\"120\" w:after=\"80\"/></w:pPr>",
        );
        body_style.set_run_properties_xml("<w:rPr><w:color w:val=\"3A5FCD\"/></w:rPr>");
        let emphasis_style = document.styles_mut().add_character_style("Emphasis");
        emphasis_style.set_name("Emphasis");
        emphasis_style.set_run_properties_xml("<w:rPr><w:b/><w:i/></w:rPr>");
        let table_style = document.styles_mut().add_table_style("TableGrid");
        table_style.set_name("Table Grid");
        table_style.set_table_properties_xml(
            "<w:tblPr><w:tblBorders><w:insideH w:val=\"single\" w:sz=\"8\"/></w:tblBorders></w:tblPr>",
        );
        table_style.add_table_style_properties_xml(
            "<w:tblStylePr w:type=\"firstRow\"><w:rPr><w:b/></w:rPr></w:tblStylePr>",
        );
        table_style.add_table_style_properties_xml(
            "<w:tblStylePr w:type=\"lastRow\"><w:rPr><w:i/></w:rPr></w:tblStylePr>",
        );

        let paragraph = document.add_paragraph_with_style("Hello", "BodyText");
        paragraph.add_run_with_style(" world", "Emphasis");
        let table = document.add_table_with_style(1, 2, "TableGrid");
        assert!(table.set_cell_text(0, 0, "A"));
        assert!(table.set_cell_text(0, 1, "B"));

        document.save(&path).expect("save styles roundtrip docx");

        let package = offidized_opc::Package::open(&path).expect("open styles package");
        assert!(package.get_part("/word/styles.xml").is_some());
        assert_eq!(
            package.content_types().get_override("/word/styles.xml"),
            Some(ContentTypeValue::WORD_STYLES)
        );
        let styles_part = package.get_part("/word/styles.xml").expect("styles part");
        let styles_xml = String::from_utf8_lossy(styles_part.data.as_bytes());
        assert!(styles_xml.contains("<w:pPr><w:spacing w:before=\"120\" w:after=\"80\"/></w:pPr>"));
        assert!(styles_xml.contains("<w:rPr><w:color w:val=\"3A5FCD\"/></w:rPr>"));
        assert!(styles_xml.contains(
            "<w:tblPr><w:tblBorders><w:insideH w:val=\"single\" w:sz=\"8\"/></w:tblBorders></w:tblPr>"
        ));
        assert!(styles_xml.contains("w:tblStylePr w:type=\"firstRow\""));
        assert!(styles_xml.contains("w:tblStylePr w:type=\"lastRow\""));
        let document_part = package
            .get_part("/word/document.xml")
            .expect("word document part");
        assert_eq!(
            document_part
                .relationships
                .get_by_type(RelationshipType::STYLES)
                .len(),
            1
        );
        let document_xml = String::from_utf8_lossy(document_part.data.as_bytes());
        assert!(document_xml.contains("<w:tblStyle w:val=\"TableGrid\""));

        let reopened = Document::open(&path).expect("open styles roundtrip docx");
        assert_eq!(reopened.paragraphs()[0].style_id(), Some("BodyText"));
        assert_eq!(
            reopened.paragraphs()[0].runs()[1].style_id(),
            Some("Emphasis")
        );
        assert_eq!(reopened.tables()[0].style_id(), Some("TableGrid"));
        assert_eq!(
            reopened.paragraph_style("BodyText").and_then(Style::name),
            Some("Body Text")
        );
        assert_eq!(
            reopened.character_style("Emphasis").and_then(Style::name),
            Some("Emphasis")
        );
        assert_eq!(
            reopened.table_style("TableGrid").and_then(Style::name),
            Some("Table Grid")
        );
        assert_eq!(
            reopened
                .paragraph_style("BodyText")
                .and_then(Style::paragraph_properties_xml),
            Some("<w:pPr><w:spacing w:before=\"120\" w:after=\"80\"/></w:pPr>")
        );
        assert_eq!(
            reopened
                .paragraph_style("BodyText")
                .and_then(Style::run_properties_xml),
            Some("<w:rPr><w:color w:val=\"3A5FCD\"/></w:rPr>")
        );
        assert_eq!(
            reopened
                .character_style("Emphasis")
                .and_then(Style::run_properties_xml),
            Some("<w:rPr><w:b/><w:i/></w:rPr>")
        );
        assert_eq!(
            reopened
                .table_style("TableGrid")
                .and_then(Style::table_properties_xml),
            Some(
                "<w:tblPr><w:tblBorders><w:insideH w:val=\"single\" w:sz=\"8\"/></w:tblBorders></w:tblPr>"
            )
        );
        assert_eq!(
            reopened
                .table_style("TableGrid")
                .map(|style| style.table_style_properties_xml().len()),
            Some(2)
        );
        assert!(reopened
            .table_style("TableGrid")
            .map(|style| style.table_style_properties_xml()[0].contains("w:type=\"firstRow\""))
            .unwrap_or(false));
        assert!(reopened
            .table_style("TableGrid")
            .map(|style| style.table_style_properties_xml()[1].contains("w:type=\"lastRow\""))
            .unwrap_or(false));

        reopened
            .save(&rewritten_path)
            .expect("resave styles roundtrip docx");
        let rewritten = Document::open(&rewritten_path).expect("open rewritten styles docx");
        assert_eq!(
            rewritten.table_style("TableGrid").and_then(Style::name),
            Some("Table Grid")
        );
        assert_eq!(
            rewritten
                .paragraph_style("BodyText")
                .and_then(Style::paragraph_properties_xml),
            Some("<w:pPr><w:spacing w:before=\"120\" w:after=\"80\"/></w:pPr>")
        );
        assert_eq!(
            rewritten
                .table_style("TableGrid")
                .map(|style| style.table_style_properties_xml().len()),
            Some(2)
        );
        assert_eq!(rewritten.tables()[0].style_id(), Some("TableGrid"));
    }

    #[test]
    fn parse_and_serialize_styles_preserve_style_property_blocks() {
        let styles_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:style w:type="paragraph" w:styleId="BodyText">
    <w:name w:val="Body Text"/>
    <w:pPr><w:spacing w:before="120" w:after="80"/></w:pPr>
    <w:rPr><w:color w:val="3A5FCD"/></w:rPr>
  </w:style>
  <w:style w:type="character" w:styleId="Emphasis">
    <w:name w:val="Emphasis"/>
    <w:rPr><w:b/><w:i/></w:rPr>
  </w:style>
  <w:style w:type="table" w:styleId="TableGrid">
    <w:name w:val="Table Grid"/>
    <w:tblPr><w:tblBorders><w:insideH w:val="single" w:sz="8"/></w:tblBorders></w:tblPr>
    <w:tblStylePr w:type="firstRow"><w:rPr><w:b/></w:rPr></w:tblStylePr>
    <w:tblStylePr w:type="lastRow"><w:rPr><w:i/></w:rPr></w:tblStylePr>
  </w:style>
</w:styles>"#;

        let parsed = super::parse_styles_xml(styles_xml.as_bytes()).expect("parse styles xml");
        let body = parsed
            .paragraph_style("BodyText")
            .expect("parsed paragraph style");
        assert_eq!(
            body.paragraph_properties_xml(),
            Some("<w:pPr><w:spacing w:before=\"120\" w:after=\"80\"/></w:pPr>")
        );
        assert_eq!(
            body.run_properties_xml(),
            Some("<w:rPr><w:color w:val=\"3A5FCD\"/></w:rPr>")
        );
        let table = parsed.table_style("TableGrid").expect("parsed table style");
        assert_eq!(
            table.table_properties_xml(),
            Some(
                "<w:tblPr><w:tblBorders><w:insideH w:val=\"single\" w:sz=\"8\"/></w:tblBorders></w:tblPr>"
            )
        );
        assert_eq!(table.table_style_properties_xml().len(), 2);

        let serialized = super::serialize_styles_xml(&parsed).expect("serialize styles xml");
        let serialized = String::from_utf8_lossy(serialized.as_slice());
        assert!(serialized.contains("<w:pPr><w:spacing w:before=\"120\" w:after=\"80\"/></w:pPr>"));
        assert!(serialized.contains("<w:rPr><w:b/><w:i/></w:rPr>"));
        assert!(serialized.contains("w:tblStylePr w:type=\"firstRow\""));
        assert!(serialized.contains("w:tblStylePr w:type=\"lastRow\""));

        let reparsed =
            super::parse_styles_xml(serialized.as_bytes()).expect("reparse serialized styles xml");
        assert_eq!(
            reparsed
                .paragraph_style("BodyText")
                .and_then(Style::paragraph_properties_xml),
            Some("<w:pPr><w:spacing w:before=\"120\" w:after=\"80\"/></w:pPr>")
        );
        assert_eq!(
            reparsed
                .table_style("TableGrid")
                .map(|style| style.table_style_properties_xml().len()),
            Some(2)
        );
    }

    #[test]
    fn open_save_roundtrips_table_horizontal_merges_and_borders() {
        let dir = tempdir().expect("create temp dir");
        let path = dir.path().join("table-fidelity-roundtrip.docx");

        let mut document = Document::new();
        let table = document.add_table(2, 4);
        assert!(table.set_cell_text(0, 0, "Merged A"));
        assert!(table.set_cell_text(0, 2, "C"));
        assert!(table.set_cell_text(0, 3, "D"));
        assert!(table.set_cell_text(1, 0, "1"));
        assert!(table.set_cell_text(1, 1, "2"));
        assert!(table.set_cell_text(1, 2, "3"));
        assert!(table.set_cell_text(1, 3, "4"));
        assert!(table.merge_cells_horizontally(0, 0, 2));

        let mut top = TableBorder::new("single");
        top.set_size_eighth_points(12);
        top.set_color("00AA11");
        table.borders_mut().set_top(top);
        let mut inside_v = TableBorder::new("dashed");
        inside_v.set_size_eighth_points(8);
        inside_v.set_color("AA5500");
        table.borders_mut().set_inside_vertical(inside_v);

        document.save(&path).expect("save table fidelity docx");

        let package = offidized_opc::Package::open(&path).expect("open table fidelity package");
        let document_part = package
            .get_part("/word/document.xml")
            .expect("word document part");
        let document_xml = String::from_utf8_lossy(document_part.data.as_bytes());
        assert!(document_xml.contains("<w:gridSpan w:val=\"2\""));
        assert!(document_xml.contains("<w:tblBorders>"));
        assert!(document_xml.contains("<w:top"));
        assert!(document_xml.contains("<w:insideV"));

        let reopened = Document::open(&path).expect("open table fidelity docx");
        let table = &reopened.tables()[0];
        assert_eq!(table.rows(), 2);
        assert_eq!(table.columns(), 4);
        assert_eq!(table.cell_text(0, 0), Some("Merged A"));
        assert_eq!(table.cell_text(0, 2), Some("C"));
        assert_eq!(table.cell_text(0, 3), Some("D"));
        assert_eq!(table.cell(0, 0).map(|cell| cell.horizontal_span()), Some(2));
        assert_eq!(
            table
                .cell(0, 1)
                .map(|cell| cell.is_horizontal_merge_continuation()),
            Some(true)
        );
        assert_eq!(
            table.borders().top().and_then(|border| border.line_type()),
            Some("single")
        );
        assert_eq!(
            table
                .borders()
                .top()
                .and_then(|border| border.size_eighth_points()),
            Some(12)
        );
        assert_eq!(
            table.borders().top().and_then(|border| border.color()),
            Some("00AA11")
        );
        assert_eq!(
            table
                .borders()
                .inside_vertical()
                .and_then(|border| border.line_type()),
            Some("dashed")
        );
    }

    #[test]
    fn open_save_roundtrips_table_cell_vertical_merge() {
        let dir = tempdir().expect("create temp dir");
        let path = dir.path().join("table-vmerge-roundtrip.docx");

        let mut document = Document::new();
        let table = document.add_table(3, 2);
        assert!(table.set_cell_text(0, 0, "Top"));
        assert!(table.set_cell_text(1, 0, ""));
        assert!(table.set_cell_text(2, 0, ""));
        assert!(table.set_cell_text(0, 1, "A"));
        assert!(table.set_cell_text(1, 1, "B"));
        assert!(table.set_cell_text(2, 1, "C"));

        table
            .cell_mut(0, 0)
            .expect("cell")
            .set_vertical_merge(VerticalMerge::Restart);
        table
            .cell_mut(1, 0)
            .expect("cell")
            .set_vertical_merge(VerticalMerge::Continue);
        table
            .cell_mut(2, 0)
            .expect("cell")
            .set_vertical_merge(VerticalMerge::Continue);

        document.save(&path).expect("save vmerge docx");

        let package = offidized_opc::Package::open(&path).expect("open vmerge package");
        let document_part = package
            .get_part("/word/document.xml")
            .expect("word document part");
        let document_xml = String::from_utf8_lossy(document_part.data.as_bytes());
        assert!(
            document_xml.contains("w:val=\"restart\""),
            "expected vMerge restart in XML"
        );
        assert!(
            document_xml.contains("<w:vMerge/>"),
            "expected empty vMerge (continue) in XML"
        );

        let reopened = Document::open(&path).expect("open vmerge docx");
        let table = &reopened.tables()[0];
        assert_eq!(
            table.cell(0, 0).and_then(|c| c.vertical_merge()),
            Some(VerticalMerge::Restart)
        );
        assert_eq!(
            table.cell(1, 0).and_then(|c| c.vertical_merge()),
            Some(VerticalMerge::Continue)
        );
        assert_eq!(
            table.cell(2, 0).and_then(|c| c.vertical_merge()),
            Some(VerticalMerge::Continue)
        );
        assert_eq!(table.cell(0, 1).and_then(|c| c.vertical_merge()), None);
    }

    #[test]
    fn open_save_roundtrips_table_cell_shading_color() {
        let dir = tempdir().expect("create temp dir");
        let path = dir.path().join("table-shading-roundtrip.docx");

        let mut document = Document::new();
        let table = document.add_table(1, 2);
        assert!(table.set_cell_text(0, 0, "Shaded"));
        assert!(table.set_cell_text(0, 1, "Plain"));

        table
            .cell_mut(0, 0)
            .expect("cell")
            .set_shading_color("#FF0000");

        document.save(&path).expect("save shading docx");

        let package = offidized_opc::Package::open(&path).expect("open shading package");
        let document_part = package
            .get_part("/word/document.xml")
            .expect("word document part");
        let document_xml = String::from_utf8_lossy(document_part.data.as_bytes());
        assert!(
            document_xml.contains("w:fill=\"FF0000\""),
            "expected shading fill in XML"
        );

        let reopened = Document::open(&path).expect("open shading docx");
        let table = &reopened.tables()[0];
        assert_eq!(
            table.cell(0, 0).and_then(|c| c.shading_color()),
            Some("FF0000")
        );
        assert_eq!(table.cell(0, 1).and_then(|c| c.shading_color()), None);
    }

    #[test]
    fn open_save_roundtrips_table_cell_vertical_alignment() {
        let dir = tempdir().expect("create temp dir");
        let path = dir.path().join("table-valign-roundtrip.docx");

        let mut document = Document::new();
        let table = document.add_table(1, 3);
        assert!(table.set_cell_text(0, 0, "Top"));
        assert!(table.set_cell_text(0, 1, "Center"));
        assert!(table.set_cell_text(0, 2, "Bottom"));

        table
            .cell_mut(0, 0)
            .expect("cell")
            .set_vertical_alignment(VerticalAlignment::Top);
        table
            .cell_mut(0, 1)
            .expect("cell")
            .set_vertical_alignment(VerticalAlignment::Center);
        table
            .cell_mut(0, 2)
            .expect("cell")
            .set_vertical_alignment(VerticalAlignment::Bottom);

        document.save(&path).expect("save valign docx");

        let package = offidized_opc::Package::open(&path).expect("open valign package");
        let document_part = package
            .get_part("/word/document.xml")
            .expect("word document part");
        let document_xml = String::from_utf8_lossy(document_part.data.as_bytes());
        assert!(
            document_xml.contains("w:val=\"center\""),
            "expected vAlign center in XML"
        );
        assert!(
            document_xml.contains("w:val=\"bottom\""),
            "expected vAlign bottom in XML"
        );

        let reopened = Document::open(&path).expect("open valign docx");
        let table = &reopened.tables()[0];
        assert_eq!(
            table.cell(0, 0).and_then(|c| c.vertical_alignment()),
            Some(VerticalAlignment::Top)
        );
        assert_eq!(
            table.cell(0, 1).and_then(|c| c.vertical_alignment()),
            Some(VerticalAlignment::Center)
        );
        assert_eq!(
            table.cell(0, 2).and_then(|c| c.vertical_alignment()),
            Some(VerticalAlignment::Bottom)
        );
    }

    #[test]
    fn open_save_roundtrips_table_cell_width_twips() {
        let dir = tempdir().expect("create temp dir");
        let path = dir.path().join("table-cell-width-roundtrip.docx");

        let mut document = Document::new();
        let table = document.add_table(1, 2);
        assert!(table.set_cell_text(0, 0, "Wide"));
        assert!(table.set_cell_text(0, 1, "Narrow"));

        table
            .cell_mut(0, 0)
            .expect("cell")
            .set_cell_width_twips(4800);
        table
            .cell_mut(0, 1)
            .expect("cell")
            .set_cell_width_twips(2400);

        document.save(&path).expect("save cell width docx");

        let package = offidized_opc::Package::open(&path).expect("open cell width package");
        let document_part = package
            .get_part("/word/document.xml")
            .expect("word document part");
        let document_xml = String::from_utf8_lossy(document_part.data.as_bytes());
        assert!(
            document_xml.contains("w:w=\"4800\""),
            "expected tcW width 4800 in XML"
        );
        assert!(
            document_xml.contains("w:type=\"dxa\""),
            "expected tcW type dxa in XML"
        );

        let reopened = Document::open(&path).expect("open cell width docx");
        let table = &reopened.tables()[0];
        assert_eq!(
            table.cell(0, 0).and_then(|c| c.cell_width_twips()),
            Some(4800)
        );
        assert_eq!(
            table.cell(0, 1).and_then(|c| c.cell_width_twips()),
            Some(2400)
        );
    }

    #[test]
    fn open_save_roundtrips_combined_table_cell_properties() {
        let dir = tempdir().expect("create temp dir");
        let path = dir.path().join("table-combined-cell-props-roundtrip.docx");

        let mut document = Document::new();
        let table = document.add_table(2, 2);
        assert!(table.set_cell_text(0, 0, "Styled"));
        assert!(table.set_cell_text(0, 1, "Plain"));
        assert!(table.set_cell_text(1, 0, "Continued"));
        assert!(table.set_cell_text(1, 1, "Also plain"));

        {
            let cell = table.cell_mut(0, 0).expect("cell");
            cell.set_vertical_merge(VerticalMerge::Restart);
            cell.set_shading_color("AABB00");
            cell.set_vertical_alignment(VerticalAlignment::Center);
            cell.set_cell_width_twips(3600);
        }
        {
            let cell = table.cell_mut(1, 0).expect("cell");
            cell.set_vertical_merge(VerticalMerge::Continue);
            cell.set_shading_color("AABB00");
            cell.set_vertical_alignment(VerticalAlignment::Center);
            cell.set_cell_width_twips(3600);
        }

        document.save(&path).expect("save combined cell props docx");

        let reopened = Document::open(&path).expect("open combined cell props docx");
        let table = &reopened.tables()[0];

        let cell_0_0 = table.cell(0, 0).expect("cell 0,0");
        assert_eq!(cell_0_0.vertical_merge(), Some(VerticalMerge::Restart));
        assert_eq!(cell_0_0.shading_color(), Some("AABB00"));
        assert_eq!(
            cell_0_0.vertical_alignment(),
            Some(VerticalAlignment::Center)
        );
        assert_eq!(cell_0_0.cell_width_twips(), Some(3600));

        let cell_1_0 = table.cell(1, 0).expect("cell 1,0");
        assert_eq!(cell_1_0.vertical_merge(), Some(VerticalMerge::Continue));
        assert_eq!(cell_1_0.shading_color(), Some("AABB00"));
        assert_eq!(
            cell_1_0.vertical_alignment(),
            Some(VerticalAlignment::Center)
        );
        assert_eq!(cell_1_0.cell_width_twips(), Some(3600));

        let cell_0_1 = table.cell(0, 1).expect("cell 0,1");
        assert_eq!(cell_0_1.vertical_merge(), None);
        assert_eq!(cell_0_1.shading_color(), None);
        assert_eq!(cell_0_1.vertical_alignment(), None);
        assert_eq!(cell_0_1.cell_width_twips(), None);
    }

    #[test]
    fn open_save_roundtrips_ordered_body_content() {
        let dir = tempdir().expect("create temp dir");
        let path = dir.path().join("ordered-body-roundtrip.docx");
        let rewritten_path = dir.path().join("ordered-body-roundtrip-rewritten.docx");

        let mut document = Document::new();
        document.add_paragraph("First paragraph");
        let table = document.add_table(1, 1);
        assert!(table.set_cell_text(0, 0, "Cell"));
        document.add_paragraph("Second paragraph");

        document.save(&path).expect("save ordered body docx");

        let reopened = Document::open(&path).expect("open ordered body docx");
        assert_eq!(reopened.paragraphs().len(), 2);
        assert_eq!(reopened.tables().len(), 1);
        assert_eq!(reopened.paragraphs()[0].text(), "First paragraph");
        assert_eq!(reopened.paragraphs()[1].text(), "Second paragraph");
        assert_eq!(reopened.tables()[0].cell_text(0, 0), Some("Cell"));

        let mut body_items = reopened.body_items();
        let Some(BodyItem::Paragraph(first_paragraph)) = body_items.next() else {
            panic!("expected first body item to be a paragraph");
        };
        assert_eq!(first_paragraph.text(), "First paragraph");

        let Some(BodyItem::Table(first_table)) = body_items.next() else {
            panic!("expected second body item to be a table");
        };
        assert_eq!(first_table.cell_text(0, 0), Some("Cell"));

        let Some(BodyItem::Paragraph(second_paragraph)) = body_items.next() else {
            panic!("expected third body item to be a paragraph");
        };
        assert_eq!(second_paragraph.text(), "Second paragraph");
        assert!(body_items.next().is_none());

        reopened
            .save(&rewritten_path)
            .expect("resave ordered body docx");

        let package = offidized_opc::Package::open(&rewritten_path).expect("open rewritten docx");
        let document_part_uri =
            super::resolve_word_document_part_uri(&package).expect("resolve document part uri");
        let document_part = package
            .get_part(document_part_uri.as_str())
            .expect("document part must exist");

        let (body_kinds, _unknown) =
            super::parse_body_item_kinds(document_part.data.as_bytes()).expect("parse body kinds");
        assert_eq!(
            body_kinds,
            vec![
                ParsedBodyItemKind::Paragraph,
                ParsedBodyItemKind::Table,
                ParsedBodyItemKind::Paragraph
            ]
        );
    }

    #[test]
    fn reference_corpus_curated_golden_roundtrip_docx() {
        let mut failures = Vec::new();
        for fixture in CURATED_REFERENCE_DOCX_FIXTURES {
            let fixture_path = reference_fixture_path(fixture);
            if !fixture_path.is_file() {
                failures.push(format!("{fixture}: missing fixture"));
                continue;
            }

            if let Err(error) = roundtrip_fixture_and_compare_fingerprint(&fixture_path) {
                failures.push(format!("{fixture}: {error}"));
            }
        }

        assert!(
            failures.is_empty(),
            "curated reference corpus failures ({}):\n{}",
            failures.len(),
            failures.join("\n")
        );
    }

    #[test]
    #[ignore = "large corpus sweep for local verification"]
    fn reference_corpus_large_sweep_meets_success_threshold_docx() {
        let fixtures = LARGE_SWEEP_REFERENCE_DOCX_DIRS
            .iter()
            .flat_map(|dir| collect_docx_fixtures_from_relative_dir(dir))
            .collect::<Vec<_>>();
        assert!(
            !fixtures.is_empty(),
            "no docx fixtures found under {:?}",
            LARGE_SWEEP_REFERENCE_DOCX_DIRS
        );

        let mut passed = 0_usize;
        let mut failures = Vec::new();
        for fixture in &fixtures {
            match roundtrip_fixture_and_compare_fingerprint(fixture) {
                Ok(()) => passed = passed.saturating_add(1),
                Err(error) => {
                    failures.push(format!("{}: {error}", fixture_display_path(fixture)));
                }
            }
        }

        let total = fixtures.len();
        let threshold = large_sweep_success_threshold();
        let success_rate = passed as f64 / total as f64;
        println!(
            "docx reference sweep: {passed}/{total} passed ({:.2}%), threshold {:.2}%",
            success_rate * 100.0,
            threshold * 100.0
        );
        let print_failures = std::env::var("OFFIDIZED_DOCX_SWEEP_PRINT_FAILURES")
            .ok()
            .as_deref()
            .is_some_and(|value| value == "1" || value.eq_ignore_ascii_case("true"));
        if print_failures && !failures.is_empty() {
            println!("docx reference sweep failures ({} total):", failures.len());
            for failure in &failures {
                println!("{failure}");
            }
        }
        let failure_details = if failures.is_empty() {
            String::new()
        } else {
            let reported = failures
                .iter()
                .take(MAX_REPORTED_SWEEP_FAILURES)
                .cloned()
                .collect::<Vec<_>>();
            let mut details = reported.join("\n");
            if failures.len() > MAX_REPORTED_SWEEP_FAILURES {
                details.push_str(&format!(
                    "\n... and {} more failures",
                    failures.len() - MAX_REPORTED_SWEEP_FAILURES
                ));
            }
            details
        };

        assert!(
            success_rate >= threshold,
            "reference corpus sweep below threshold: {}/{} ({:.2}% < {:.2}%)\n{}",
            passed,
            total,
            success_rate * 100.0,
            threshold * 100.0,
            failure_details
        );
    }

    #[test]
    fn reference_corpus_known_roundtrip_regression_fixtures_docx() {
        let fixtures = [
            "TestFiles/Complex01.docx",
            "TestFiles/Document.docx",
            "TestFiles/May_12_04.docx",
            "TestFiles/Notes.docx",
            "TestFiles/complex0.docx",
            "TestFiles/complex2010.docx",
            "TestFiles/svg.docx",
            "TestDataStorage/O15Conformance/WD/CommentExTest/Comments-Sample-15-12-01/Comment049.docx",
        ];

        let mut missing = Vec::new();
        let mut failures = Vec::new();

        for fixture in fixtures {
            let path = reference_fixture_path(fixture);
            if !path.is_file() {
                missing.push(fixture.to_string());
                continue;
            }

            if let Err(error) = roundtrip_fixture_and_compare_fingerprint(&path) {
                failures.push(format!("{fixture}: {error}"));
            }
        }

        if !missing.is_empty() && failures.is_empty() {
            eprintln!(
                "skipping known DOCX regression fixture test because fixtures are missing:\n{}",
                missing.join("\n")
            );
            return;
        }

        assert!(
            missing.is_empty(),
            "known DOCX regression fixtures missing:\n{}",
            missing.join("\n")
        );
        assert!(
            failures.is_empty(),
            "known DOCX regression fixture failures:\n{}",
            failures.join("\n")
        );
    }

    #[test]
    #[ignore = "debug helper for fingerprint divergence details"]
    fn debug_reference_corpus_known_failure_fingerprint_diffs_docx() {
        let fixtures = [
            "TestFiles/Complex01.docx",
            "TestFiles/Document.docx",
            "TestFiles/May_12_04.docx",
            "TestFiles/Notes.docx",
            "TestFiles/complex0.docx",
            "TestFiles/complex2010.docx",
            "TestFiles/svg.docx",
            "TestDataStorage/O15Conformance/WD/CommentExTest/Comments-Sample-15-12-01/Comment049.docx",
        ];

        for fixture in fixtures {
            let path = reference_fixture_path(fixture);
            assert!(path.is_file(), "fixture missing: {fixture}");

            let opened = Document::open(&path).expect("open fixture");
            let before = document_fingerprint(&opened);
            let dir = tempdir().expect("create temp dir");
            let rewritten_path = dir.path().join("rewritten.docx");
            opened.save(&rewritten_path).expect("save rewritten");
            let reopened = Document::open(&rewritten_path).expect("reopen rewritten");
            let after = document_fingerprint(&reopened);

            if before != after {
                println!("fixture `{fixture}` fingerprint diffs:");
                for diff in diff_document_fingerprint(&before, &after) {
                    println!("  - {diff}");
                }
            } else {
                println!("fixture `{fixture}` now matches fingerprint");
            }
        }
    }

    #[test]
    fn tables_mut_allows_safe_updates() {
        let mut document = Document::new();
        document.add_table(1, 1);

        if let Some(table) = document.tables_mut().first_mut() {
            assert!(table.set_cell_text(0, 0, "Updated"));
        }

        assert_eq!(document.tables()[0].cell_text(0, 0), Some("Updated"));
    }

    #[test]
    fn new_from_scratch_save_and_reopen_works() {
        let mut document = Document::new();
        document.add_paragraph("Hello");
        document.add_heading("Title", 1);
        document.add_table(2, 2);

        let dir = tempdir().expect("create temp dir");
        let path = dir.path().join("scratch.docx");
        document.save(&path).expect("save new document");

        let reopened = Document::open(&path).expect("reopen saved document");
        assert_eq!(reopened.paragraphs().len(), 2);
        assert_eq!(reopened.paragraphs()[0].text(), "Hello");
        assert_eq!(reopened.paragraphs()[1].text(), "Title");
        assert_eq!(reopened.tables().len(), 1);
        assert_eq!(reopened.tables()[0].rows(), 2);
        assert_eq!(reopened.tables()[0].columns(), 2);
    }

    #[test]
    fn dirty_flag_tracks_new_open_default_and_mutating_apis() {
        let dir = tempdir().expect("create temp dir");
        let seed_path = dir.path().join("seed.docx");

        let mut seed = Document::new();
        seed.add_paragraph("seed");
        seed.save(&seed_path).expect("save seed docx");

        assert!(Document::new().dirty, "new document should start dirty");
        assert!(
            Document::default().dirty,
            "default document should start dirty"
        );

        let opened = Document::open(&seed_path).expect("open seed docx");
        assert!(
            !opened.dirty,
            "opened document should be pristine before any mutation"
        );

        let mut document = Document::open(&seed_path).expect("open for add_paragraph");
        document.add_paragraph("paragraph");
        assert!(document.dirty, "add_paragraph should mark dirty");

        let mut document = Document::open(&seed_path).expect("open for add_paragraph_with_style");
        document.add_paragraph_with_style("styled paragraph", "Normal");
        assert!(document.dirty, "add_paragraph_with_style should mark dirty");

        let mut document = Document::open(&seed_path).expect("open for add_heading");
        document.add_heading("Heading", 1);
        assert!(document.dirty, "add_heading should mark dirty");

        let mut document = Document::open(&seed_path).expect("open for add_image");
        document.add_image(b"image-bytes".to_vec(), "image/png");
        assert!(document.dirty, "add_image should mark dirty");

        let mut document = Document::open(&seed_path).expect("open for images_mut");
        let _ = document.images_mut();
        assert!(document.dirty, "images_mut should mark dirty");

        let mut document = Document::open(&seed_path).expect("open for add_table");
        document.add_table(1, 1);
        assert!(document.dirty, "add_table should mark dirty");

        let mut document = Document::open(&seed_path).expect("open for add_table_with_style");
        document.add_table_with_style(1, 1, "TableGrid");
        assert!(document.dirty, "add_table_with_style should mark dirty");

        let mut document = Document::open(&seed_path).expect("open for tables_mut");
        let _ = document.tables_mut();
        assert!(document.dirty, "tables_mut should mark dirty");

        let mut document = Document::open(&seed_path).expect("open for section_mut");
        let _ = document.section_mut();
        assert!(document.dirty, "section_mut should mark dirty");

        let mut document = Document::open(&seed_path).expect("open for styles_mut");
        let _ = document.styles_mut();
        assert!(document.dirty, "styles_mut should mark dirty");
    }

    #[test]
    fn pristine_open_save_preserves_original_package_without_reserializing() {
        let dir = tempdir().expect("create temp dir");
        let original_path = dir.path().join("original.docx");
        let rewritten_path = dir.path().join("rewritten.docx");
        let orphan_part_uri = "/word/header-orphan.xml";
        let orphan_part_data = b"<orphan>preserve me</orphan>";

        {
            let mut document = Document::new();
            document.add_paragraph("seed");
            document.save(&original_path).expect("save initial docx");
        }

        {
            let mut package = offidized_opc::Package::open(&original_path).expect("open package");
            let part_uri =
                offidized_opc::uri::PartUri::new(orphan_part_uri).expect("parse orphan part uri");
            let part = offidized_opc::Part::new_xml(part_uri, orphan_part_data.to_vec());
            package.set_part(part);
            package
                .save(&original_path)
                .expect("save package with orphan");
        }

        let opened = Document::open(&original_path).expect("open document");
        assert!(!opened.dirty, "opened document should be pristine");
        opened
            .save(&rewritten_path)
            .expect("save pristine document without edits");

        let rewritten_package =
            offidized_opc::Package::open(&rewritten_path).expect("open rewritten package");
        let preserved_part = rewritten_package
            .get_part(orphan_part_uri)
            .expect("orphan word part should be preserved on pristine save");
        assert_eq!(preserved_part.data.as_bytes(), orphan_part_data);
    }

    #[test]
    fn roundtrip_preserves_unknown_package_parts() {
        let dir = tempdir().expect("create temp dir");
        let original_path = dir.path().join("original.docx");
        let custom_relationship_type = "https://offidized.dev/relationships/custom-docx-data";

        // Create a document, save it, then inject an extra part.
        {
            let mut document = Document::new();
            document.add_paragraph("test content");
            document.save(&original_path).expect("save initial");
        }

        // Re-open the package at OPC level, inject a custom part.
        let custom_part_uri = "/customXml/item1.xml";
        let custom_part_data = b"<root><data>preserved</data></root>";
        {
            let mut package = offidized_opc::Package::open(&original_path).expect("open package");
            let custom_uri = offidized_opc::uri::PartUri::new(custom_part_uri).expect("parse uri");
            let part = offidized_opc::Part::new_xml(custom_uri, custom_part_data.to_vec());
            package.set_part(part);
            package
                .get_part_mut(super::WORD_DOCUMENT_URI)
                .expect("document part should exist")
                .relationships
                .add_new(
                    custom_relationship_type.to_string(),
                    "../customXml/item1.xml".to_string(),
                    TargetMode::Internal,
                );
            package.save(&original_path).expect("save with custom part");
        }

        // Open as Document, add a paragraph, save again.
        let roundtripped_path = dir.path().join("roundtripped.docx");
        {
            let mut document = Document::open(&original_path).expect("open with custom part");
            document.add_paragraph("extra paragraph");
            document
                .save(&roundtripped_path)
                .expect("save roundtripped");
        }

        // Verify the custom part survived the roundtrip.
        let package =
            offidized_opc::Package::open(&roundtripped_path).expect("open roundtripped package");
        let custom_part = package.get_part(custom_part_uri);
        assert!(
            custom_part.is_some(),
            "custom part should survive roundtrip"
        );
        assert_eq!(
            custom_part.unwrap().data.as_bytes(),
            custom_part_data,
            "custom part data should be preserved"
        );
        let document_part = package
            .get_part(super::WORD_DOCUMENT_URI)
            .expect("document part should exist");
        let custom_relationships = document_part
            .relationships
            .get_by_type(custom_relationship_type);
        assert_eq!(
            custom_relationships.len(),
            1,
            "unmanaged document relationship should survive dirty save as pass-through",
        );

        // Also verify the document content is intact.
        let document = Document::open(&roundtripped_path).expect("reopen roundtripped");
        assert_eq!(document.paragraphs().len(), 2);
        assert_eq!(document.paragraphs()[0].text(), "test content");
        assert_eq!(document.paragraphs()[1].text(), "extra paragraph");
    }

    #[test]
    fn dirty_roundtrip_preserves_unknown_paragraph_and_run_children() {
        let dir = tempdir().expect("create temp dir");
        let original_path = dir.path().join("original.docx");
        let mutated_path = dir.path().join("mutated.docx");
        let roundtripped_path = dir.path().join("roundtripped.docx");

        {
            let mut document = Document::new();
            document.add_paragraph("seed");
            document.save(&original_path).expect("save initial docx");
        }

        {
            let mut package = offidized_opc::Package::open(&original_path).expect("open package");
            let document_part = package
                .get_part_mut(super::WORD_DOCUMENT_URI)
                .expect("document part should exist");
            let document_xml = String::from_utf8_lossy(document_part.data.as_bytes()).into_owned();
            let injected_xml = document_xml
                .replacen(
                    "<w:p>",
                    "<w:p><w:pPr><w:customPPr w:val=\"yes\"/></w:pPr>",
                    1,
                )
                .replacen(
                    "<w:r>",
                    "<w:r><w:rPr><w:customRPr w:val=\"yes\"/></w:rPr>",
                    1,
                )
                .replacen("</w:r>", "<w:customRunChild w:val=\"yes\"/></w:r>", 1)
                .replacen("</w:p>", "<w:customParagraphChild w:val=\"yes\"/></w:p>", 1);
            assert!(injected_xml.contains("<w:customPPr"));
            assert!(injected_xml.contains("<w:customRPr"));
            assert!(injected_xml.contains("<w:customRunChild"));
            assert!(injected_xml.contains("<w:customParagraphChild"));
            document_part.data = offidized_opc::PartData::Xml(injected_xml.into_bytes());
            package
                .save(&mutated_path)
                .expect("save package with unknown paragraph/run children");
        }

        let mut opened = Document::open(&mutated_path).expect("open mutated docx");
        opened.add_paragraph("extra");
        opened
            .save(&roundtripped_path)
            .expect("save roundtripped docx");

        let final_package =
            offidized_opc::Package::open(&roundtripped_path).expect("open roundtripped package");
        let final_document_xml = String::from_utf8_lossy(
            final_package
                .get_part(super::WORD_DOCUMENT_URI)
                .expect("document part should exist")
                .data
                .as_bytes(),
        )
        .into_owned();

        assert!(final_document_xml.contains("<w:customPPr"));
        assert!(final_document_xml.contains("<w:customRPr"));
        assert!(final_document_xml.contains("<w:customRunChild"));
        assert!(final_document_xml.contains("<w:customParagraphChild"));
    }

    #[test]
    fn open_save_roundtrips_strikethrough_properties() {
        let dir = tempdir().expect("create temp dir");
        let path = dir.path().join("strikethrough-roundtrip.docx");

        let mut document = Document::new();
        let paragraph = document.add_paragraph("plain");
        let strike_run = paragraph.add_run(" struck");
        strike_run.set_strikethrough(true);
        let dstrike_run = paragraph.add_run(" double");
        dstrike_run.set_double_strikethrough(true);

        document.save(&path).expect("save strikethrough docx");

        let reopened = Document::open(&path).expect("open strikethrough docx");
        let runs = reopened.paragraphs()[0].runs();
        assert_eq!(runs.len(), 3);
        assert!(!runs[0].is_strikethrough());
        assert!(!runs[0].is_double_strikethrough());
        assert!(runs[1].is_strikethrough());
        assert!(!runs[1].is_double_strikethrough());
        assert!(!runs[2].is_strikethrough());
        assert!(runs[2].is_double_strikethrough());
    }

    #[test]
    fn open_save_roundtrips_subscript_superscript_properties() {
        let dir = tempdir().expect("create temp dir");
        let path = dir.path().join("vert-align-roundtrip.docx");

        let mut document = Document::new();
        let paragraph = document.add_paragraph("H");
        let sub_run = paragraph.add_run("2");
        sub_run.set_subscript(true);
        let mid_run = paragraph.add_run("O is water. E=mc");
        let _ = mid_run;
        let sup_run = paragraph.add_run("2");
        sup_run.set_superscript(true);

        document.save(&path).expect("save vert-align docx");

        let reopened = Document::open(&path).expect("open vert-align docx");
        let runs = reopened.paragraphs()[0].runs();
        assert_eq!(runs.len(), 4);
        assert!(!runs[0].is_subscript());
        assert!(!runs[0].is_superscript());
        assert!(runs[1].is_subscript());
        assert!(!runs[1].is_superscript());
        assert!(!runs[2].is_subscript());
        assert!(!runs[2].is_superscript());
        assert!(!runs[3].is_subscript());
        assert!(runs[3].is_superscript());
    }

    #[test]
    fn open_save_roundtrips_highlight_color() {
        let dir = tempdir().expect("create temp dir");
        let path = dir.path().join("highlight-roundtrip.docx");

        let mut document = Document::new();
        let paragraph = document.add_paragraph("normal");
        let highlighted_run = paragraph.add_run(" highlighted");
        highlighted_run.set_highlight_color("yellow");
        let green_run = paragraph.add_run(" green");
        green_run.set_highlight_color("green");

        document.save(&path).expect("save highlight docx");

        let reopened = Document::open(&path).expect("open highlight docx");
        let runs = reopened.paragraphs()[0].runs();
        assert_eq!(runs.len(), 3);
        assert_eq!(runs[0].highlight_color(), None);
        assert_eq!(runs[1].highlight_color(), Some("yellow"));
        assert_eq!(runs[2].highlight_color(), Some("green"));
    }

    #[test]
    fn serialized_xml_contains_new_run_property_elements() {
        let dir = tempdir().expect("create temp dir");
        let path = dir.path().join("xml-check.docx");

        let mut document = Document::new();
        let paragraph = document.add_paragraph("");
        let run = paragraph.add_run("all props");
        run.set_strikethrough(true);
        run.set_double_strikethrough(true);
        run.set_subscript(true);
        run.set_highlight_color("cyan");

        document.save(&path).expect("save xml-check docx");

        let package = offidized_opc::Package::open(&path).expect("open package");
        let document_xml = String::from_utf8_lossy(
            package
                .get_part(super::WORD_DOCUMENT_URI)
                .expect("document part")
                .data
                .as_bytes(),
        )
        .into_owned();

        assert!(
            document_xml.contains("<w:strike/>"),
            "should contain w:strike element"
        );
        assert!(
            document_xml.contains("<w:dstrike/>"),
            "should contain w:dstrike element"
        );
        assert!(
            document_xml.contains("w:vertAlign"),
            "should contain w:vertAlign element"
        );
        assert!(
            document_xml.contains("\"subscript\""),
            "should contain subscript value"
        );
        assert!(
            document_xml.contains("w:highlight"),
            "should contain w:highlight element"
        );
        assert!(
            document_xml.contains("\"cyan\""),
            "should contain cyan highlight value"
        );
    }

    #[test]
    fn superscript_serialized_xml_uses_correct_vert_align_value() {
        let dir = tempdir().expect("create temp dir");
        let path = dir.path().join("superscript-xml-check.docx");

        let mut document = Document::new();
        let paragraph = document.add_paragraph("");
        let run = paragraph.add_run("sup");
        run.set_superscript(true);

        document.save(&path).expect("save superscript docx");

        let package = offidized_opc::Package::open(&path).expect("open package");
        let document_xml = String::from_utf8_lossy(
            package
                .get_part(super::WORD_DOCUMENT_URI)
                .expect("document part")
                .data
                .as_bytes(),
        )
        .into_owned();

        assert!(
            document_xml.contains("\"superscript\""),
            "should contain superscript value in vertAlign"
        );
    }

    #[test]
    fn tab_stops_roundtrip_through_serialize_and_parse() {
        use crate::paragraph::{TabStop, TabStopAlignment, TabStopLeader};

        let dir = tempdir().expect("create temp dir");
        let path = dir.path().join("tab-stops-roundtrip.docx");

        let mut document = Document::new();
        let paragraph = document.add_paragraph("Tab test");
        paragraph.add_tab_stop(TabStop::new(2880, TabStopAlignment::Center));
        let mut right_tab = TabStop::new(5760, TabStopAlignment::Right);
        right_tab.set_leader(TabStopLeader::Dot);
        paragraph.add_tab_stop(right_tab);
        document.save(&path).expect("save tab stops docx");

        let loaded = Document::open(&path).expect("open tab stops docx");
        assert!(!loaded.paragraphs().is_empty());
        let para = &loaded.paragraphs()[0];
        assert_eq!(para.tab_stops().len(), 2, "should have 2 tab stops");
        assert_eq!(para.tab_stops()[0].alignment(), TabStopAlignment::Center);
        assert_eq!(para.tab_stops()[0].position_twips(), 2880);
        assert_eq!(para.tab_stops()[1].alignment(), TabStopAlignment::Right);
        assert_eq!(para.tab_stops()[1].position_twips(), 5760);
        assert_eq!(para.tab_stops()[1].leader(), Some(TabStopLeader::Dot));
    }

    #[test]
    fn paragraph_borders_roundtrip_through_serialize_and_parse() {
        use crate::paragraph::ParagraphBorder;

        let dir = tempdir().expect("create temp dir");
        let path = dir.path().join("para-borders-roundtrip.docx");

        let mut document = Document::new();
        let paragraph = document.add_paragraph("Bordered");
        let mut top_border = ParagraphBorder::default();
        top_border.set_line_type("single");
        top_border.set_size_eighth_points(6);
        top_border.set_color("FF0000");
        paragraph.borders_mut().set_top(top_border);
        document.save(&path).expect("save borders docx");

        let loaded = Document::open(&path).expect("open borders docx");
        let para = &loaded.paragraphs()[0];
        let top = para.borders().top().expect("should have top border");
        assert_eq!(top.line_type(), Some("single"));
        assert_eq!(top.size_eighth_points(), Some(6));
        assert_eq!(top.color(), Some("FF0000"));
    }

    #[test]
    fn paragraph_shading_roundtrip_through_serialize_and_parse() {
        let dir = tempdir().expect("create temp dir");
        let path = dir.path().join("para-shading-roundtrip.docx");

        let mut document = Document::new();
        let paragraph = document.add_paragraph("Shaded");
        paragraph.set_shading_color("FFFF00");
        document.save(&path).expect("save shading docx");

        let loaded = Document::open(&path).expect("open shading docx");
        let para = &loaded.paragraphs()[0];
        assert_eq!(para.shading_color(), Some("FFFF00"));
    }

    #[test]
    fn keep_next_keep_lines_widow_control_roundtrip() {
        let dir = tempdir().expect("create temp dir");
        let path = dir.path().join("keep-props-roundtrip.docx");

        let mut document = Document::new();
        let paragraph = document.add_paragraph("Flow control");
        paragraph.set_keep_next(true);
        paragraph.set_keep_lines(true);
        paragraph.set_widow_control(false);
        document.save(&path).expect("save flow props docx");

        let loaded = Document::open(&path).expect("open flow props docx");
        let para = &loaded.paragraphs()[0];
        assert!(para.keep_next());
        assert!(para.keep_lines());
        assert_eq!(para.widow_control(), Some(false));
    }

    #[test]
    fn run_tab_and_break_roundtrip() {
        let dir = tempdir().expect("create temp dir");
        let path = dir.path().join("run-special-roundtrip.docx");

        let mut document = Document::new();
        let paragraph = document.add_paragraph("");
        let tab_run = paragraph.add_run("");
        tab_run.set_has_tab(true);
        let br_run = paragraph.add_run("");
        br_run.set_has_break(true);
        document.save(&path).expect("save special content docx");

        let loaded = Document::open(&path).expect("open special content docx");
        let para = &loaded.paragraphs()[0];
        let runs = para.runs();
        // runs[0] is the initial empty run from add_paragraph(""),
        // runs[1] is the tab run, runs[2] is the break run.
        assert!(runs.len() >= 3, "should have at least 3 runs");
        assert!(runs[1].has_tab(), "second run should have tab");
        assert!(runs[2].has_break(), "third run should have break");
    }

    #[test]
    fn run_footnote_endnote_reference_roundtrip() {
        let dir = tempdir().expect("create temp dir");
        let path = dir.path().join("footnote-endnote-roundtrip.docx");

        let mut document = Document::new();
        let paragraph = document.add_paragraph("");
        let fn_run = paragraph.add_run("");
        fn_run.set_footnote_reference_id(1);
        let en_run = paragraph.add_run("");
        en_run.set_endnote_reference_id(2);
        document.save(&path).expect("save references docx");

        let loaded = Document::open(&path).expect("open references docx");
        let para = &loaded.paragraphs()[0];
        let runs = para.runs();
        // runs[0] is the initial empty run from add_paragraph(""),
        // runs[1] is the footnote run, runs[2] is the endnote run.
        assert!(runs.len() >= 3);
        assert_eq!(runs[1].footnote_reference_id(), Some(1));
        assert_eq!(runs[2].endnote_reference_id(), Some(2));
    }

    #[test]
    fn table_column_widths_roundtrip() {
        let dir = tempdir().expect("create temp dir");
        let path = dir.path().join("table-col-widths-roundtrip.docx");

        let mut document = Document::new();
        let table = document.add_table(2, 3);
        table.set_column_widths_twips(vec![2400, 3600, 2400]);
        document.save(&path).expect("save col widths docx");

        let loaded = Document::open(&path).expect("open col widths docx");
        assert!(!loaded.tables().is_empty());
        assert_eq!(
            loaded.tables()[0].column_widths_twips(),
            &[2400, 3600, 2400]
        );
    }

    #[test]
    fn table_row_properties_roundtrip() {
        let dir = tempdir().expect("create temp dir");
        let path = dir.path().join("table-row-props-roundtrip.docx");

        let mut document = Document::new();
        let table = document.add_table(2, 2);
        if let Some(row_props) = table.row_properties_mut(0) {
            row_props.set_repeat_header(true);
            row_props.set_height_twips(720);
            row_props.set_height_rule("exact");
        }
        document.save(&path).expect("save row props docx");

        let loaded = Document::open(&path).expect("open row props docx");
        assert!(!loaded.tables().is_empty());
        let row0_props = loaded.tables()[0].row_properties(0).expect("row 0 props");
        assert!(row0_props.repeat_header(), "row 0 should be header");
        assert_eq!(row0_props.height_twips(), Some(720));
        assert_eq!(row0_props.height_rule(), Some("exact"));
    }

    #[test]
    fn cell_borders_and_margins_roundtrip() {
        use crate::table::{CellBorders, CellMargins, TableBorder};

        let dir = tempdir().expect("create temp dir");
        let path = dir.path().join("cell-borders-margins-roundtrip.docx");

        let mut document = Document::new();
        let table = document.add_table(1, 1);
        if let Some(cell) = table.cell_mut(0, 0) {
            cell.set_text("test");
            let mut borders = CellBorders::new();
            borders.set_top(TableBorder::new("single"));
            borders.set_bottom(TableBorder::new("double"));
            cell.set_borders(borders);
            let mut margins = CellMargins::new();
            margins.set_top_twips(100);
            margins.set_left_twips(200);
            cell.set_margins(margins);
        }
        document.save(&path).expect("save cell borders docx");

        let loaded = Document::open(&path).expect("open cell borders docx");
        assert!(!loaded.tables().is_empty());
        let cell = loaded.tables()[0].cell(0, 0).expect("cell 0,0");
        let top_border = cell.borders().top().expect("should have top border");
        assert_eq!(top_border.line_type(), Some("single"));
        let bottom_border = cell.borders().bottom().expect("should have bottom border");
        assert_eq!(bottom_border.line_type(), Some("double"));
        assert_eq!(cell.margins().top_twips(), Some(100));
        assert_eq!(cell.margins().left_twips(), Some(200));
    }

    #[test]
    fn style_based_on_and_next_roundtrip() {
        let dir = tempdir().expect("create temp dir");
        let path = dir.path().join("style-inheritance-roundtrip.docx");

        let mut document = Document::new();
        document.add_paragraph_with_style("Test", "BodyText");
        let styles = document.styles_mut();
        let style = styles.ensure_paragraph_style("BodyText");
        style.set_name("Body Text");
        style.set_based_on("Normal");
        style.set_next_style("Normal");
        styles.ensure_paragraph_style("Normal").set_name("Normal");
        document.save(&path).expect("save style inheritance docx");

        let loaded = Document::open(&path).expect("open style inheritance docx");
        let styles = loaded.styles();
        let body_text = styles
            .paragraph_style("BodyText")
            .expect("should have BodyText style");
        assert_eq!(body_text.based_on(), Some("Normal"));
        assert_eq!(body_text.next_style(), Some("Normal"));
    }

    #[test]
    fn section_break_type_roundtrip() {
        use crate::section::SectionBreakType;

        let dir = tempdir().expect("create temp dir");
        let path = dir.path().join("section-break-roundtrip.docx");

        let mut document = Document::new();
        document.add_paragraph("Page content");
        document
            .section_mut()
            .set_break_type(SectionBreakType::Continuous);
        document.save(&path).expect("save section break docx");

        let loaded = Document::open(&path).expect("open section break docx");
        assert_eq!(
            loaded.section().break_type(),
            Some(SectionBreakType::Continuous)
        );
    }

    #[test]
    fn title_page_and_first_page_header_footer_roundtrip() {
        let dir = tempdir().expect("create temp dir");
        let path = dir.path().join("first-page-hf-roundtrip.docx");

        let mut document = Document::new();
        document.add_paragraph("Document body");

        let section = document.section_mut();
        section.set_title_page(true);

        let mut default_header = HeaderFooter::new();
        default_header.add_paragraph("Default Header");
        section.set_header(default_header);

        let mut default_footer = HeaderFooter::new();
        default_footer.add_paragraph("Default Footer");
        section.set_footer(default_footer);

        let mut first_header = HeaderFooter::new();
        first_header.add_paragraph("First Page Header");
        section.set_first_page_header(first_header);

        let mut first_footer = HeaderFooter::new();
        first_footer.add_paragraph("First Page Footer");
        section.set_first_page_footer(first_footer);

        document
            .save(&path)
            .expect("save first page header/footer docx");

        let loaded = Document::open(&path).expect("open first page header/footer docx");
        assert!(loaded.section().title_page());

        let default_header = loaded
            .section()
            .header()
            .expect("should have default header");
        assert!(
            default_header
                .paragraphs()
                .iter()
                .any(|p| p.text().contains("Default Header")),
            "default header text"
        );

        let default_footer = loaded
            .section()
            .footer()
            .expect("should have default footer");
        assert!(
            default_footer
                .paragraphs()
                .iter()
                .any(|p| p.text().contains("Default Footer")),
            "default footer text"
        );

        let first_header = loaded
            .section()
            .first_page_header()
            .expect("should have first page header");
        assert!(
            first_header
                .paragraphs()
                .iter()
                .any(|p| p.text().contains("First Page Header")),
            "first page header text"
        );

        let first_footer = loaded
            .section()
            .first_page_footer()
            .expect("should have first page footer");
        assert!(
            first_footer
                .paragraphs()
                .iter()
                .any(|p| p.text().contains("First Page Footer")),
            "first page footer text"
        );
    }

    #[test]
    fn paragraph_shading_pattern_roundtrip() {
        let dir = tempdir().expect("create temp dir");
        let path = dir.path().join("para-shading-pattern-roundtrip.docx");

        let mut document = Document::new();
        let paragraph = document.add_paragraph("Patterned shading");
        paragraph.set_shading_color("00FF00");
        paragraph.set_shading_pattern("diagStripe");
        document.save(&path).expect("save shading pattern docx");

        let loaded = Document::open(&path).expect("open shading pattern docx");
        let para = &loaded.paragraphs()[0];
        assert_eq!(para.shading_color(), Some("00FF00"));
        assert_eq!(para.shading_pattern(), Some("diagStripe"));
    }

    #[test]
    fn inline_section_properties_roundtrip() {
        use crate::section::{Section, SectionBreakType};

        let dir = tempdir().expect("create temp dir");
        let path = dir.path().join("inline-sectpr-roundtrip.docx");

        let mut document = Document::new();
        let paragraph = document.add_paragraph("Section break here");
        let mut inline_section = Section::new();
        inline_section.set_break_type(SectionBreakType::NextPage);
        inline_section.set_page_size_twips(12240, 15840);
        paragraph.set_section_properties(inline_section);
        document.add_paragraph("After break");
        document.save(&path).expect("save inline section docx");

        let loaded = Document::open(&path).expect("open inline section docx");
        assert!(loaded.paragraphs().len() >= 2);
        let first_para = &loaded.paragraphs()[0];
        let inline_sect = first_para
            .section_properties()
            .expect("should have inline section properties");
        assert_eq!(
            inline_sect.break_type(),
            Some(SectionBreakType::NextPage),
            "inline section break type should be NextPage"
        );
        assert_eq!(inline_sect.page_width_twips(), Some(12240));
        assert_eq!(inline_sect.page_height_twips(), Some(15840));
    }

    // ======== Feature: Even page headers/footers ========

    #[test]
    fn even_page_header_footer_roundtrip() {
        let dir = tempdir().expect("create temp dir");
        let path = dir.path().join("even-hf-roundtrip.docx");

        let mut document = Document::new();
        document.add_paragraph("Body content");

        let section = document.section_mut();
        let mut default_header = HeaderFooter::new();
        default_header.add_paragraph("Odd Header");
        section.set_header(default_header);

        let mut even_header = HeaderFooter::new();
        even_header.add_paragraph("Even Header");
        section.set_even_page_header(even_header);

        let mut even_footer = HeaderFooter::new();
        even_footer.add_paragraph("Even Footer");
        section.set_even_page_footer(even_footer);

        document.save(&path).expect("save even hf docx");

        let loaded = Document::open(&path).expect("open even hf docx");

        let even_header = loaded
            .section()
            .even_page_header()
            .expect("should have even page header");
        assert!(
            even_header
                .paragraphs()
                .iter()
                .any(|p| p.text().contains("Even Header")),
            "even page header text"
        );

        let even_footer = loaded
            .section()
            .even_page_footer()
            .expect("should have even page footer");
        assert!(
            even_footer
                .paragraphs()
                .iter()
                .any(|p| p.text().contains("Even Footer")),
            "even page footer text"
        );
    }

    #[test]
    fn even_page_header_is_none_by_default() {
        let doc = Document::new();
        assert!(doc.section().even_page_header().is_none());
        assert!(doc.section().even_page_footer().is_none());
    }

    // ======== Feature: Page numbering format/start ========

    #[test]
    fn page_numbering_roundtrip() {
        let dir = tempdir().expect("create temp dir");
        let path = dir.path().join("pg-num-roundtrip.docx");

        let mut document = Document::new();
        document.add_paragraph("Page content");

        let section = document.section_mut();
        section.set_page_number_start(5);
        section.set_page_number_format("lowerRoman");

        document.save(&path).expect("save page num docx");

        let loaded = Document::open(&path).expect("open page num docx");
        assert_eq!(loaded.section().page_number_start(), Some(5));
        assert_eq!(loaded.section().page_number_format(), Some("lowerRoman"));
    }

    #[test]
    fn page_numbering_is_none_by_default() {
        let doc = Document::new();
        assert_eq!(doc.section().page_number_start(), None);
        assert_eq!(doc.section().page_number_format(), None);
    }

    // ======== Feature: Footnotes/Endnotes ========

    #[test]
    fn footnotes_api_basic() {
        use crate::footnote::Footnote;

        let mut doc = Document::new();
        assert!(doc.footnotes().is_empty());

        doc.add_footnote(Footnote::from_text(1, "First footnote"));
        doc.add_footnote(Footnote::from_text(2, "Second footnote"));

        assert_eq!(doc.footnotes().len(), 2);
        assert_eq!(doc.footnotes()[0].text(), "First footnote");
        assert_eq!(doc.footnotes()[1].id(), 2);
    }

    #[test]
    fn endnotes_api_basic() {
        use crate::footnote::Endnote;

        let mut doc = Document::new();
        assert!(doc.endnotes().is_empty());

        doc.add_endnote(Endnote::from_text(1, "Endnote text"));

        assert_eq!(doc.endnotes().len(), 1);
        assert_eq!(doc.endnotes()[0].text(), "Endnote text");
    }

    // ======== Feature: Complex field support ========

    #[test]
    fn field_code_roundtrip() {
        use crate::run::FieldCode;

        let dir = tempdir().expect("create temp dir");
        let path = dir.path().join("field-code-roundtrip.docx");

        let mut document = Document::new();
        let para = document.add_paragraph("");
        let run = para.add_run("");
        run.set_field_code(FieldCode::new("PAGE", "3"));
        document.save(&path).expect("save field code docx");

        let loaded = Document::open(&path).expect("open field code docx");
        let runs: Vec<_> = loaded
            .paragraphs()
            .iter()
            .flat_map(|p| p.runs().iter())
            .collect();
        let field_run = runs
            .iter()
            .find(|r| r.field_code().is_some())
            .expect("should find a run with a field code");
        let field_code = field_run.field_code().unwrap();
        assert_eq!(field_code.instruction(), "PAGE");
        assert_eq!(field_code.result(), "3");
    }

    #[test]
    fn field_code_api() {
        use crate::run::FieldCode;

        let mut doc = Document::new();
        let para = doc.add_paragraph("");
        let run = para.add_run("");
        assert!(run.field_code().is_none());

        run.set_field_code(FieldCode::new("DATE", "2024-01-01"));
        assert_eq!(run.field_code().unwrap().instruction(), "DATE");

        run.clear_field_code();
        assert!(run.field_code().is_none());
    }

    // ======== Feature: Bookmarks ========

    #[test]
    fn bookmarks_api_basic() {
        use crate::bookmark::Bookmark;

        let mut doc = Document::new();
        assert!(doc.bookmarks().is_empty());

        doc.add_bookmark(Bookmark::new(0, "intro", 0, 2));
        doc.add_bookmark(Bookmark::new(1, "_GoBack", 5, 5));

        assert_eq!(doc.bookmarks().len(), 2);
        assert_eq!(doc.bookmarks()[0].name(), "intro");
        assert_eq!(doc.bookmarks()[0].start_paragraph_index(), 0);
        assert_eq!(doc.bookmarks()[0].end_paragraph_index(), 2);
        assert!(!doc.bookmarks()[0].is_single_paragraph());
        assert!(doc.bookmarks()[1].is_single_paragraph());
    }

    #[test]
    fn bookmarks_mutable() {
        use crate::bookmark::Bookmark;

        let mut doc = Document::new();
        doc.add_bookmark(Bookmark::new(0, "mark1", 0, 0));
        assert_eq!(doc.bookmarks().len(), 1);

        doc.bookmarks_mut().push(Bookmark::new(1, "mark2", 1, 2));
        assert_eq!(doc.bookmarks().len(), 2);
    }

    // ======== Feature: Comments ========

    #[test]
    fn comments_api_basic() {
        use crate::comment::Comment;

        let mut doc = Document::new();
        assert!(doc.comments().is_empty());

        let mut c = Comment::from_text(1, "Reviewer", "Good point");
        c.set_date("2024-01-15T10:00:00Z");
        doc.add_comment(c);

        assert_eq!(doc.comments().len(), 1);
        assert_eq!(doc.comments()[0].text(), "Good point");
        assert_eq!(doc.comments()[0].author(), "Reviewer");
        assert_eq!(doc.comments()[0].date(), Some("2024-01-15T10:00:00Z"));
    }

    #[test]
    fn comment_range_on_paragraph() {
        let mut doc = Document::new();
        let para = doc.add_paragraph("Annotated text");
        para.add_comment_range_start(1);
        para.add_comment_range_end(1);

        assert_eq!(para.comment_range_start_ids(), &[1]);
        assert_eq!(para.comment_range_end_ids(), &[1]);
    }

    // ======== Feature: Document properties ========

    #[test]
    fn document_properties_api_basic() {
        let mut doc = Document::new();
        assert!(doc.document_properties().is_empty());

        doc.document_properties_mut().set_title("My Report");
        doc.document_properties_mut().set_creator("Author");

        assert_eq!(doc.document_properties().title(), Some("My Report"));
        assert_eq!(doc.document_properties().creator(), Some("Author"));
    }

    #[test]
    fn document_properties_can_be_replaced() {
        use crate::properties::DocumentProperties;

        let mut doc = Document::new();
        let mut props = DocumentProperties::new();
        props.set_title("Replaced Title");
        doc.set_document_properties(props);

        assert_eq!(doc.document_properties().title(), Some("Replaced Title"));
    }

    // ======== Feature: RTL text direction ========

    #[test]
    fn bidi_paragraph_roundtrip() {
        let dir = tempdir().expect("create temp dir");
        let path = dir.path().join("bidi-roundtrip.docx");

        let mut document = Document::new();
        let para = document.add_paragraph("RTL paragraph");
        para.set_bidi(true);
        document.save(&path).expect("save bidi docx");

        let loaded = Document::open(&path).expect("open bidi docx");
        assert!(loaded.paragraphs()[0].is_bidi());
    }

    #[test]
    fn rtl_run_roundtrip() {
        let dir = tempdir().expect("create temp dir");
        let path = dir.path().join("rtl-run-roundtrip.docx");

        let mut document = Document::new();
        let para = document.add_paragraph("");
        let run = para.add_run("Hebrew text");
        run.set_rtl(true);
        document.save(&path).expect("save rtl run docx");

        let loaded = Document::open(&path).expect("open rtl run docx");
        let runs = loaded.paragraphs()[0].runs();
        // First run is the empty-text run from add_paragraph(""), second is the RTL run
        let rtl_run = runs
            .iter()
            .find(|r| !r.text().is_empty())
            .expect("find non-empty run");
        assert!(rtl_run.is_rtl());
        assert_eq!(rtl_run.text(), "Hebrew text");
    }

    // ======== Feature: Content controls (SDT) ========

    #[test]
    fn content_controls_api_basic() {
        use crate::content_control::ContentControl;

        let mut doc = Document::new();
        assert!(doc.content_controls().is_empty());

        let mut sdt = ContentControl::new();
        sdt.set_tag("myTag");
        sdt.set_alias("My Control");
        sdt.add_paragraph("Content 1");
        doc.add_content_control(sdt);

        assert_eq!(doc.content_controls().len(), 1);
        assert_eq!(doc.content_controls()[0].tag(), Some("myTag"));
        assert_eq!(doc.content_controls()[0].text(), "Content 1");
    }

    #[test]
    fn content_controls_mutable() {
        use crate::content_control::ContentControl;

        let mut doc = Document::new();
        let sdt = ContentControl::new();
        doc.add_content_control(sdt);
        assert_eq!(doc.content_controls().len(), 1);

        doc.content_controls_mut().clear();
        assert!(doc.content_controls().is_empty());
    }

    // ======== Feature: Numbering definitions ========

    #[test]
    fn numbering_definitions_api_basic() {
        use crate::numbering::{NumberingDefinition, NumberingLevel};

        let mut doc = Document::new();
        assert!(doc.numbering_definitions().is_empty());

        let mut def = NumberingDefinition::new(0);
        def.add_level(NumberingLevel::new(0, 1, "decimal", "%1."));
        def.add_level(NumberingLevel::new(1, 1, "lowerLetter", "%2)"));
        doc.add_numbering_definition(def);

        assert_eq!(doc.numbering_definitions().len(), 1);
        assert_eq!(doc.numbering_definitions()[0].levels().len(), 2);
    }

    #[test]
    fn numbering_definitions_mutable() {
        use crate::numbering::{NumberingDefinition, NumberingLevel};

        let mut doc = Document::new();
        let mut def = NumberingDefinition::new(1);
        def.add_level(NumberingLevel::new(0, 1, "bullet", ""));
        doc.add_numbering_definition(def);

        assert_eq!(doc.numbering_definitions().len(), 1);
        doc.numbering_definitions_mut().clear();
        assert!(doc.numbering_definitions().is_empty());
    }

    // ======== Feature: Parse footnotes XML ========

    #[test]
    fn parse_footnotes_from_xml() {
        use super::parse_footnotes_xml;

        let xml = br#"<?xml version="1.0" encoding="UTF-8"?>
        <w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
          <w:footnote w:id="0" w:type="separator">
            <w:p><w:r><w:t></w:t></w:r></w:p>
          </w:footnote>
          <w:footnote w:id="1">
            <w:p><w:r><w:t>First footnote text</w:t></w:r></w:p>
          </w:footnote>
          <w:footnote w:id="2">
            <w:p><w:r><w:t>Second footnote</w:t></w:r></w:p>
          </w:footnote>
        </w:footnotes>"#;

        let footnotes = parse_footnotes_xml(xml).expect("parse footnotes");
        assert_eq!(footnotes.len(), 3);
        assert_eq!(footnotes[1].id(), 1);
        assert_eq!(footnotes[1].text(), "First footnote text");
        assert_eq!(footnotes[2].id(), 2);
        assert_eq!(footnotes[2].text(), "Second footnote");
    }

    // ======== Feature: Parse endnotes XML ========

    #[test]
    fn parse_endnotes_from_xml() {
        use super::parse_endnotes_xml;

        let xml = br#"<?xml version="1.0" encoding="UTF-8"?>
        <w:endnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
          <w:endnote w:id="0" w:type="separator">
            <w:p><w:r><w:t></w:t></w:r></w:p>
          </w:endnote>
          <w:endnote w:id="1">
            <w:p><w:r><w:t>Endnote content here</w:t></w:r></w:p>
          </w:endnote>
        </w:endnotes>"#;

        let endnotes = parse_endnotes_xml(xml).expect("parse endnotes");
        assert_eq!(endnotes.len(), 2);
        assert_eq!(endnotes[1].id(), 1);
        assert_eq!(endnotes[1].text(), "Endnote content here");
    }

    // ======== Feature: Parse bookmarks ========

    #[test]
    fn parse_bookmarks_from_document_xml() {
        use super::parse_bookmarks;

        let xml = br#"<?xml version="1.0" encoding="UTF-8"?>
        <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
          <w:body>
            <w:bookmarkStart w:id="0" w:name="intro"/>
            <w:p><w:r><w:t>First</w:t></w:r></w:p>
            <w:p><w:r><w:t>Second</w:t></w:r></w:p>
            <w:bookmarkEnd w:id="0"/>
            <w:p><w:r><w:t>Third</w:t></w:r></w:p>
          </w:body>
        </w:document>"#;

        let bookmarks = parse_bookmarks(xml).expect("parse bookmarks");
        assert_eq!(bookmarks.len(), 1);
        assert_eq!(bookmarks[0].name(), "intro");
        assert_eq!(bookmarks[0].id(), 0);
    }

    // ======== Feature: Parse comments XML ========

    #[test]
    fn parse_comments_from_xml() {
        use super::parse_comments_xml;

        let xml = br#"<?xml version="1.0" encoding="UTF-8"?>
        <w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
          <w:comment w:id="1" w:author="Reviewer" w:date="2024-01-15T10:00:00Z">
            <w:p><w:r><w:t>This needs revision.</w:t></w:r></w:p>
          </w:comment>
          <w:comment w:id="2" w:author="Editor">
            <w:p><w:r><w:t>Looks good.</w:t></w:r></w:p>
          </w:comment>
        </w:comments>"#;

        let comments = parse_comments_xml(xml).expect("parse comments");
        assert_eq!(comments.len(), 2);
        assert_eq!(comments[0].id(), 1);
        assert_eq!(comments[0].author(), "Reviewer");
        assert_eq!(comments[0].date(), Some("2024-01-15T10:00:00Z"));
        assert_eq!(comments[0].text(), "This needs revision.");
        assert_eq!(comments[1].id(), 2);
        assert_eq!(comments[1].author(), "Editor");
        assert_eq!(comments[1].text(), "Looks good.");
    }

    // ======== Feature: Parse core properties XML ========

    #[test]
    fn parse_core_properties_from_xml() {
        use super::parse_core_properties_xml;

        let xml = br#"<?xml version="1.0" encoding="UTF-8"?>
        <cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
            xmlns:dc="http://purl.org/dc/elements/1.1/"
            xmlns:dcterms="http://purl.org/dc/terms/">
          <dc:title>Quarterly Report</dc:title>
          <dc:creator>John Doe</dc:creator>
          <dc:subject>Finance</dc:subject>
          <dc:description>Q1 revenue analysis</dc:description>
          <cp:keywords>finance, quarterly</cp:keywords>
          <cp:lastModifiedBy>Jane Smith</cp:lastModifiedBy>
          <dcterms:created>2024-01-15T10:00:00Z</dcterms:created>
          <dcterms:modified>2024-03-20T14:30:00Z</dcterms:modified>
        </cp:coreProperties>"#;

        let props = parse_core_properties_xml(xml).expect("parse core properties");
        assert_eq!(props.title(), Some("Quarterly Report"));
        assert_eq!(props.creator(), Some("John Doe"));
        assert_eq!(props.subject(), Some("Finance"));
        assert_eq!(props.description(), Some("Q1 revenue analysis"));
        assert_eq!(props.keywords(), Some("finance, quarterly"));
        assert_eq!(props.last_modified_by(), Some("Jane Smith"));
        assert_eq!(props.created(), Some("2024-01-15T10:00:00Z"));
        assert_eq!(props.modified(), Some("2024-03-20T14:30:00Z"));
    }

    // ======== Feature: Parse content controls XML ========

    #[test]
    fn parse_content_controls_from_xml() {
        use super::parse_content_controls;

        let xml = br#"<?xml version="1.0" encoding="UTF-8"?>
        <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
          <w:body>
            <w:sdt>
              <w:sdtPr>
                <w:tag w:val="myTag"/>
                <w:alias w:val="My Control"/>
              </w:sdtPr>
              <w:sdtContent>
                <w:p><w:r><w:t>Content paragraph 1</w:t></w:r></w:p>
                <w:p><w:r><w:t>Content paragraph 2</w:t></w:r></w:p>
              </w:sdtContent>
            </w:sdt>
            <w:p><w:r><w:t>Normal paragraph</w:t></w:r></w:p>
          </w:body>
        </w:document>"#;

        let controls = parse_content_controls(xml).expect("parse content controls");
        assert_eq!(controls.len(), 1);
        assert_eq!(controls[0].tag(), Some("myTag"));
        assert_eq!(controls[0].alias(), Some("My Control"));
        assert_eq!(controls[0].content().len(), 2);
        assert_eq!(
            controls[0].text(),
            "Content paragraph 1\nContent paragraph 2"
        );
    }

    // ======== Feature: Parse numbering XML ========

    #[test]
    fn parse_numbering_definitions_from_xml() {
        use super::parse_numbering_xml;

        let xml = br#"<?xml version="1.0" encoding="UTF-8"?>
        <w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
          <w:abstractNum w:abstractNumId="0">
            <w:lvl w:ilvl="0">
              <w:start w:val="1"/>
              <w:numFmt w:val="decimal"/>
              <w:lvlText w:val="%1."/>
              <w:lvlJc w:val="left"/>
            </w:lvl>
            <w:lvl w:ilvl="1">
              <w:start w:val="1"/>
              <w:numFmt w:val="lowerLetter"/>
              <w:lvlText w:val="%2)"/>
            </w:lvl>
          </w:abstractNum>
          <w:abstractNum w:abstractNumId="1">
            <w:lvl w:ilvl="0">
              <w:start w:val="1"/>
              <w:numFmt w:val="bullet"/>
              <w:lvlText w:val=""/>
            </w:lvl>
          </w:abstractNum>
        </w:numbering>"#;

        let (defs, instances) = parse_numbering_xml(xml).expect("parse numbering");
        assert_eq!(defs.len(), 2);
        assert_eq!(defs[0].abstract_num_id(), 0);
        assert_eq!(defs[0].levels().len(), 2);
        assert_eq!(defs[0].levels()[0].format(), "decimal");
        assert_eq!(defs[0].levels()[0].text(), "%1.");
        assert_eq!(defs[0].levels()[0].alignment(), Some("left"));
        assert_eq!(defs[0].levels()[1].format(), "lowerLetter");
        assert_eq!(defs[0].levels()[1].text(), "%2)");
        assert_eq!(defs[0].levels()[1].alignment(), None);

        assert_eq!(defs[1].abstract_num_id(), 1);
        assert_eq!(defs[1].levels().len(), 1);
        assert_eq!(defs[1].levels()[0].format(), "bullet");
        assert!(instances.is_empty());
    }

    // ======== Feature: Parse comment range on paragraphs ========

    #[test]
    fn comment_range_parsed_in_paragraphs() {
        use super::parse_paragraphs;

        let xml = br#"<?xml version="1.0" encoding="UTF-8"?>
        <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
          <w:body>
            <w:p>
              <w:commentRangeStart w:id="1"/>
              <w:r><w:t>Annotated text</w:t></w:r>
              <w:commentRangeEnd w:id="1"/>
            </w:p>
          </w:body>
        </w:document>"#;

        let hyperlinks = std::collections::HashMap::new();
        let images = std::collections::HashMap::new();
        let paragraphs = parse_paragraphs(xml, &hyperlinks, &images).expect("parse paragraphs");
        assert_eq!(paragraphs.len(), 1);
        assert_eq!(paragraphs[0].comment_range_start_ids(), &[1]);
        assert_eq!(paragraphs[0].comment_range_end_ids(), &[1]);
    }

    // ======== Feature: Parse bidi paragraph ========

    #[test]
    fn bidi_parsed_in_paragraphs() {
        use super::parse_paragraphs;

        let xml = br#"<?xml version="1.0" encoding="UTF-8"?>
        <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
          <w:body>
            <w:p>
              <w:pPr><w:bidi/></w:pPr>
              <w:r><w:t>RTL text</w:t></w:r>
            </w:p>
          </w:body>
        </w:document>"#;

        let hyperlinks = std::collections::HashMap::new();
        let images = std::collections::HashMap::new();
        let paragraphs = parse_paragraphs(xml, &hyperlinks, &images).expect("parse paragraphs");
        assert_eq!(paragraphs.len(), 1);
        assert!(paragraphs[0].is_bidi());
    }

    // ======== Feature: Parse RTL run ========

    #[test]
    fn rtl_parsed_in_runs() {
        use super::parse_paragraphs;

        let xml = br#"<?xml version="1.0" encoding="UTF-8"?>
        <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
          <w:body>
            <w:p>
              <w:r>
                <w:rPr><w:rtl/></w:rPr>
                <w:t>Hebrew</w:t>
              </w:r>
            </w:p>
          </w:body>
        </w:document>"#;

        let hyperlinks = std::collections::HashMap::new();
        let images = std::collections::HashMap::new();
        let paragraphs = parse_paragraphs(xml, &hyperlinks, &images).expect("parse paragraphs");
        assert_eq!(paragraphs.len(), 1);
        assert_eq!(paragraphs[0].runs().len(), 1);
        assert!(paragraphs[0].runs()[0].is_rtl());
    }

    // ======== Feature: Parse field codes ========

    #[test]
    fn field_code_parsed_in_runs() {
        use super::parse_paragraphs;

        let xml = br#"<?xml version="1.0" encoding="UTF-8"?>
        <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
          <w:body>
            <w:p>
              <w:r>
                <w:fldChar w:fldCharType="begin"/>
                <w:instrText> PAGE </w:instrText>
                <w:fldChar w:fldCharType="separate"/>
                <w:t>7</w:t>
                <w:fldChar w:fldCharType="end"/>
              </w:r>
            </w:p>
          </w:body>
        </w:document>"#;

        let hyperlinks = std::collections::HashMap::new();
        let images = std::collections::HashMap::new();
        let paragraphs = parse_paragraphs(xml, &hyperlinks, &images).expect("parse paragraphs");
        assert_eq!(paragraphs.len(), 1);
        assert_eq!(paragraphs[0].runs().len(), 1);
        let field_code = paragraphs[0].runs()[0]
            .field_code()
            .expect("should have field code");
        assert_eq!(field_code.instruction(), "PAGE");
        assert_eq!(field_code.result(), "7");
    }

    // ======== Feature: Parse pgNumType ========

    #[test]
    fn page_num_type_parsed_in_section() {
        use super::parse_section;
        use offidized_opc::{Package, Part};

        let xml = br#"<?xml version="1.0" encoding="UTF-8"?>
        <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
          <w:body>
            <w:p><w:r><w:t>hello</w:t></w:r></w:p>
            <w:sectPr>
              <w:pgNumType w:fmt="lowerRoman" w:start="3"/>
            </w:sectPr>
          </w:body>
        </w:document>"#;

        let package = Package::new();
        let part_uri = offidized_opc::uri::PartUri::new("/word/document.xml").expect("valid uri");
        let part = Part::new_xml(part_uri.clone(), xml.to_vec());

        let section = parse_section(&package, &part_uri, &part, xml).expect("parse section");
        assert_eq!(section.page_number_start(), Some(3));
        assert_eq!(section.page_number_format(), Some("lowerRoman"));
    }

    // ======== Phase 3A: Bullet/Numbering API tests ========

    #[test]
    fn add_bulleted_paragraph_creates_numbering() {
        let mut doc = Document::new();
        doc.add_bulleted_paragraph("Bullet item 1");
        doc.add_bulleted_paragraph("Bullet item 2");

        assert_eq!(doc.paragraphs().len(), 2);
        assert_eq!(doc.paragraphs()[0].text(), "Bullet item 1");
        assert_eq!(doc.paragraphs()[1].text(), "Bullet item 2");

        // Both should share the same numbering num_id
        let num_id_0 = doc.paragraphs()[0].numbering_num_id();
        let num_id_1 = doc.paragraphs()[1].numbering_num_id();
        assert!(num_id_0.is_some());
        assert_eq!(num_id_0, num_id_1);

        // A bullet numbering definition should exist
        assert!(!doc.numbering_definitions().is_empty());
    }

    #[test]
    fn add_numbered_paragraph_creates_numbering() {
        let mut doc = Document::new();
        doc.add_numbered_paragraph("Item 1");
        doc.add_numbered_paragraph("Item 2");

        assert_eq!(doc.paragraphs().len(), 2);
        let num_id_0 = doc.paragraphs()[0].numbering_num_id();
        assert!(num_id_0.is_some());
    }

    #[test]
    fn ensure_bullet_reuses_existing_definition() {
        let mut doc = Document::new();
        let id1 = doc.ensure_bullet_numbering_definition();
        let id2 = doc.ensure_bullet_numbering_definition();
        assert_eq!(id1, id2);
    }

    #[test]
    fn ensure_numbered_reuses_existing_definition() {
        let mut doc = Document::new();
        let id1 = doc.ensure_numbered_list_definition();
        let id2 = doc.ensure_numbered_list_definition();
        assert_eq!(id1, id2);
    }

    #[test]
    fn bullet_and_numbered_get_different_ids() {
        let mut doc = Document::new();
        let bullet_id = doc.ensure_bullet_numbering_definition();
        let numbered_id = doc.ensure_numbered_list_definition();
        assert_ne!(bullet_id, numbered_id);
    }

    // ======== Phase 3B: Document Protection tests ========

    #[test]
    fn document_protection_defaults() {
        let doc = Document::new();
        assert!(doc.protection().is_none());
    }

    #[test]
    fn document_protection_set_and_clear() {
        let mut doc = Document::new();
        doc.set_protection(DocumentProtection::read_only());
        assert!(doc.protection().is_some());
        assert_eq!(doc.protection().unwrap().edit, "readOnly");
        assert!(doc.protection().unwrap().enforcement);

        doc.clear_protection();
        assert!(doc.protection().is_none());
    }

    #[test]
    fn document_protection_constructors() {
        let ro = DocumentProtection::read_only();
        assert_eq!(ro.edit, "readOnly");
        assert!(ro.enforcement);

        let co = DocumentProtection::comments_only();
        assert_eq!(co.edit, "comments");

        let tc = DocumentProtection::tracked_changes();
        assert_eq!(tc.edit, "trackedChanges");

        let fo = DocumentProtection::forms_only();
        assert_eq!(fo.edit, "forms");
    }

    #[test]
    fn document_protection_custom() {
        let prot = DocumentProtection::new("readOnly", false);
        assert_eq!(prot.edit, "readOnly");
        assert!(!prot.enforcement);
    }

    /// Regression: extra namespace declarations on `<w:document>` must survive
    /// a dirty save so that unknown elements/attributes using those prefixes
    /// remain valid XML.
    #[test]
    fn dirty_save_preserves_extra_namespace_declarations() {
        use offidized_opc::uri::PartUri;
        use offidized_opc::{Package, Part};

        // Minimal document.xml with extra xmlns:mc and xmlns:w14 declarations.
        let doc_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
            xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
            xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml">
  <w:body>
    <w:p><w:r><w:t>Hello</w:t></w:r></w:p>
    <w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>
  </w:body>
</w:document>"#;

        let mut package = Package::new();

        let doc_uri = PartUri::new("/word/document.xml").unwrap();
        let mut doc_part = Part::new_xml(doc_uri, doc_xml.to_vec());
        doc_part.content_type = Some(ContentTypeValue::WORD_DOCUMENT.to_string());
        package.set_part(doc_part);

        package.relationships_mut().add_new(
            RelationshipType::WORD_DOCUMENT.to_string(),
            "word/document.xml".to_string(),
            TargetMode::Internal,
        );

        let tmpdir = tempdir().unwrap();
        let pkg_path = tmpdir.path().join("test.docx");
        package.save(&pkg_path).unwrap();

        // Open, modify (marks dirty), and save.
        let mut doc = Document::open(&pkg_path).unwrap();
        doc.add_paragraph("Added paragraph");
        let out_path = tmpdir.path().join("out.docx");
        doc.save(&out_path).unwrap();

        // Extract the document XML and verify extra namespace declarations survived.
        let out_package = Package::open(&out_path).unwrap();
        let doc_part = out_package
            .get_part("/word/document.xml")
            .expect("document part missing");
        let doc_xml_out = String::from_utf8_lossy(doc_part.data.as_bytes());

        assert!(
            doc_xml_out.contains("xmlns:mc"),
            "mc namespace declaration missing from document XML:\n{doc_xml_out}"
        );
        assert!(
            doc_xml_out.contains("xmlns:w14"),
            "w14 namespace declaration missing from document XML:\n{doc_xml_out}"
        );
    }
}
