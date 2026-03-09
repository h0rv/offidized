use std::collections::{BTreeMap, HashMap};
use std::io::Cursor;
use std::path::Path;

use offidized_opc::content_types::ContentTypeValue;
use offidized_opc::relationship::{RelationshipType, TargetMode};
use offidized_opc::uri::PartUri;
use offidized_opc::{Package, Part, PartData, RawXmlNode};
use quick_xml::events::{BytesDecl, BytesEnd, BytesStart, BytesText, Event};
use quick_xml::{Reader, Writer};

use crate::chart::{
    Chart, ChartDataLabel, ChartSeries, ChartType, LegendPosition, SeriesBorder, SeriesFill,
    SeriesMarker,
};
use crate::color::{write_color_xml, write_solid_fill_color_xml, ColorTransform, ShapeColor};
use crate::comment::SlideComment;
use crate::error::{PptxError, Result};
use crate::image::Image;
use crate::shape::{
    ArrowSize, ArrowType, AutoFitType, BulletStyle, ConnectionInfo, GradientFill, GradientFillType,
    GradientStop, LineArrow, LineCompoundStyle, LineDashStyle, LineSpacing, LineSpacingUnit,
    MediaType, ParagraphProperties, PatternFill, PatternFillType, PictureFill, PlaceholderType,
    Shape, ShapeGeometry, ShapeGlow, ShapeOutline, ShapeParagraph, ShapeReflection, ShapeShadow,
    ShapeType, SpacingUnit, SpacingValue, TextAlignment, TextAnchor,
};
use crate::slide::{ShapeGroup, Slide, SlideBackground, SlideHeaderFooter};
use crate::slide_layout::SlideLayout;
use crate::slide_layout_io;
use crate::slide_master::SlideMaster;
use crate::slide_master_io;
use crate::table::{CellBorder, CellBorders, CellTextAnchor, Table, TableCell, TextDirection};
use crate::text::RunProperties;
use crate::theme::ThemeColorScheme;
use crate::timing::{SlideAnimationNode, SlideTiming};
use crate::transition::{SlideTransition, SlideTransitionKind, TransitionSound, TransitionSpeed};

const PRESENTATION_PART_URI: &str = "/ppt/presentation.xml";
const PRESENTATIONML_NS: &str = "http://schemas.openxmlformats.org/presentationml/2006/main";
const DRAWINGML_NS: &str = "http://schemas.openxmlformats.org/drawingml/2006/main";
const CHART_NS: &str = "http://schemas.openxmlformats.org/drawingml/2006/chart";
const RELATIONSHIP_NS: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const TABLE_GRAPHIC_DATA_URI: &str = "http://schemas.openxmlformats.org/drawingml/2006/table";
const CHART_GRAPHIC_DATA_URI: &str = "http://schemas.openxmlformats.org/drawingml/2006/chart";
const OCTET_STREAM_CONTENT_TYPE: &str = "application/octet-stream";
const CHART_CONTENT_TYPE: &str =
    "application/vnd.openxmlformats-officedocument.drawingml.chart+xml";
const NOTES_SLIDE_RELATIONSHIP_TYPE: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide";
const NOTES_SLIDE_CONTENT_TYPE: &str =
    "application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml";
const COMMENTS_RELATIONSHIP_TYPE: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments";
const COMMENT_AUTHORS_RELATIONSHIP_TYPE: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/commentAuthors";
const COMMENTS_CONTENT_TYPE: &str =
    "application/vnd.openxmlformats-officedocument.presentationml.comments+xml";
const COMMENT_AUTHORS_CONTENT_TYPE: &str =
    "application/vnd.openxmlformats-officedocument.presentationml.commentAuthors+xml";
const COMMENT_AUTHORS_PART_URI: &str = "/ppt/commentAuthors.xml";
const DEFAULT_SLIDE_MASTER_PART_URI: &str = "/ppt/slideMasters/slideMaster1.xml";
const DEFAULT_SLIDE_LAYOUT_PART_URI: &str = "/ppt/slideLayouts/slideLayout1.xml";
const DEFAULT_SLIDE_MASTER_ID: u32 = 2_147_483_648;
const DEFAULT_SLIDE_LAYOUT_ID: u32 = 2_147_483_649;
const THEME_RELATIONSHIP_TYPE: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme";
#[allow(dead_code)]
const HYPERLINK_RELATIONSHIP_TYPE: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink";

#[derive(Debug, Clone)]
struct SlideRef {
    slide_id: u32,
    relationship_id: String,
    part_uri: PartUri,
}

#[derive(Debug, Clone)]
struct SlideMasterRef {
    master_id: u32,
    relationship_id: String,
}

#[derive(Debug, Clone)]
struct ParsedPresentationRefs {
    slide_refs: Vec<ParsedSlideRef>,
    slide_master_refs: Vec<ParsedSlideMasterRef>,
}

#[derive(Debug, Clone)]
struct ParsedSlideRef {
    relationship_id: String,
}

#[derive(Debug, Clone)]
struct ParsedSlideMasterRef {
    relationship_id: String,
}

#[derive(Debug, Clone, PartialEq, Eq)]
struct ParsedSlideLayoutRef {
    relationship_id: String,
}

#[derive(Debug, Clone, PartialEq)]
struct ParsedSlideMasterMetadata {
    layout_refs: Vec<ParsedSlideLayoutRef>,
    shapes: Vec<Shape>,
}

#[derive(Debug, Clone, PartialEq)]
struct ParsedSlideLayoutMetadata {
    name: Option<String>,
    layout_type: Option<String>,
    preserve: Option<bool>,
    shapes: Vec<Shape>,
}

#[derive(Debug, Clone, PartialEq)]
pub struct SlideLayoutMetadata {
    relationship_id: String,
    part_uri: String,
    name: Option<String>,
    layout_type: Option<String>,
    preserve: Option<bool>,
    shapes: Vec<Shape>,
}

impl SlideLayoutMetadata {
    pub fn relationship_id(&self) -> &str {
        self.relationship_id.as_str()
    }

    pub fn part_uri(&self) -> &str {
        self.part_uri.as_str()
    }

    pub fn name(&self) -> Option<&str> {
        self.name.as_deref()
    }

    pub fn layout_type(&self) -> Option<&str> {
        self.layout_type.as_deref()
    }

    pub fn r#type(&self) -> Option<&str> {
        self.layout_type()
    }

    pub fn preserve(&self) -> Option<bool> {
        self.preserve
    }

    pub fn shapes(&self) -> &[Shape] {
        self.shapes.as_slice()
    }
}

#[derive(Debug, Clone, PartialEq)]
pub struct SlideMasterMetadata {
    relationship_id: String,
    part_uri: String,
    layouts: Vec<SlideLayoutMetadata>,
    shapes: Vec<Shape>,
}

impl SlideMasterMetadata {
    pub fn relationship_id(&self) -> &str {
        self.relationship_id.as_str()
    }

    pub fn part_uri(&self) -> &str {
        self.part_uri.as_str()
    }

    pub fn layouts(&self) -> &[SlideLayoutMetadata] {
        self.layouts.as_slice()
    }

    pub fn shapes(&self) -> &[Shape] {
        self.shapes.as_slice()
    }
}

#[derive(Debug, Clone)]
struct ParsedImageRef {
    relationship_id: String,
    name: Option<String>,
}

#[derive(Debug, Clone)]
struct SerializedImageRef {
    relationship_id: String,
    name: String,
}

#[derive(Debug, Clone)]
struct ParsedChartRef {
    relationship_id: String,
}

#[derive(Debug, Clone)]
struct SerializedChartRef {
    relationship_id: String,
    name: String,
}

#[derive(Debug, Clone)]
struct ParsedSlideComment {
    author_id: u32,
    text: String,
}

#[derive(Debug, Clone)]
struct SerializedSlideComment {
    author_id: u32,
    comment_index: u32,
    text: String,
}

#[derive(Debug, Clone)]
struct SerializedCommentAuthor {
    id: u32,
    name: String,
    last_comment_index: u32,
}

#[derive(Debug)]
pub struct Presentation {
    package: Package,
    slides: Vec<Slide>,
    /// Legacy read-only slide master metadata (backward compatibility).
    slide_masters: Vec<SlideMasterMetadata>,
    /// Mutable slide masters (write API).
    slide_masters_v2: Vec<SlideMaster>,
    /// Slide width in EMUs (Feature #9).
    slide_width_emu: Option<i64>,
    /// Slide height in EMUs (Feature #9).
    slide_height_emu: Option<i64>,
    /// Theme color scheme (Feature #14).
    theme_color_scheme: Option<ThemeColorScheme>,
    /// Theme font scheme.
    theme_font_scheme: Option<crate::theme::ThemeFontScheme>,
    /// Presentation sections.
    sections: Vec<crate::slide::PresentationSection>,
    /// First slide number (`firstSlideNum` attr on `<p:presentation>`).
    first_slide_number: Option<u32>,
    /// Show special placeholders on title slide
    /// (`showSpecialPlsOnTitleSld` attr on `<p:presentation>`).
    show_special_pls_on_title_sld: Option<bool>,
    /// Right-to-left presentation (`rtl` attr on `<p:presentation>`).
    right_to_left: Option<bool>,
    /// Custom shows (named subsets of slides).
    custom_shows: Vec<crate::custom_show::CustomShow>,
    /// Slide show settings.
    slide_show_settings: Option<crate::custom_show::SlideShowSettings>,
    /// Presentation properties (metadata from docProps/core.xml and docProps/app.xml).
    presentation_properties: crate::presentation_properties::PresentationProperties,
    dirty: bool,
}

impl Presentation {
    pub fn new() -> Self {
        Self {
            package: Package::new(),
            slides: Vec::new(),
            slide_masters: Vec::new(),
            slide_masters_v2: Vec::new(),
            slide_width_emu: None,
            slide_height_emu: None,
            theme_color_scheme: None,
            theme_font_scheme: None,
            sections: Vec::new(),
            first_slide_number: None,
            show_special_pls_on_title_sld: None,
            right_to_left: None,
            custom_shows: Vec::new(),
            slide_show_settings: None,
            presentation_properties: crate::presentation_properties::PresentationProperties::new(),
            dirty: true,
        }
    }

    pub fn open(path: impl AsRef<Path>) -> Result<Self> {
        let package = Package::open(path)?;
        Self::from_package(package)
    }

    /// Open an existing `.pptx` package from in-memory bytes.
    pub fn from_bytes(bytes: &[u8]) -> Result<Self> {
        let package = Package::from_bytes(bytes)?;
        Self::from_package(package)
    }

    /// Build a `Presentation` from an already-opened OPC package.
    fn from_package(package: Package) -> Result<Self> {
        let presentation_uri = resolve_presentation_part_uri(&package)?;
        let presentation_part = package.get_part(presentation_uri.as_str()).ok_or_else(|| {
            PptxError::UnsupportedPackage("missing presentation part".to_string())
        })?;
        let parsed_refs = parse_presentation_xml(presentation_part.data.as_bytes())?;
        let (slide_width_emu, slide_height_emu) =
            parse_slide_size(presentation_part.data.as_bytes())?;
        let pres_attrs = parse_presentation_attrs(presentation_part.data.as_bytes())?;
        let theme_color_scheme =
            parse_theme_from_package(&package, &presentation_uri, presentation_part)?;
        let theme_font_scheme =
            parse_theme_font_scheme_from_package(&package, &presentation_uri, presentation_part)?;
        let sections = parse_presentation_sections(presentation_part.data.as_bytes())?;
        let slide_masters = resolve_slide_masters(
            &package,
            &presentation_uri,
            presentation_part,
            &parsed_refs.slide_master_refs,
        )?;

        let mut slides = Vec::with_capacity(parsed_refs.slide_refs.len());
        for parsed_ref in parsed_refs.slide_refs {
            let relationship = presentation_part
                .relationships
                .get_by_id(parsed_ref.relationship_id.as_str())
                .ok_or_else(|| {
                    PptxError::UnsupportedPackage(format!(
                        "missing slide relationship `{}`",
                        parsed_ref.relationship_id
                    ))
                })?;
            if relationship.target_mode != TargetMode::Internal {
                return Err(PptxError::UnsupportedPackage(format!(
                    "slide relationship `{}` is external",
                    relationship.id
                )));
            }

            let slide_uri = presentation_uri.resolve_relative(relationship.target.as_str())?;
            let slide_part = package.get_part(slide_uri.as_str()).ok_or_else(|| {
                PptxError::UnsupportedPackage(format!(
                    "missing slide part `{}` for relationship `{}`",
                    slide_uri.as_str(),
                    relationship.id
                ))
            })?;
            let _ = resolve_slide_layout_part_uri(&package, &slide_uri, slide_part)?;

            let slide_xml = slide_part.data.as_bytes();
            let mut parsed_shapes = parse_slide_shapes(slide_xml)?;
            resolve_hyperlink_urls(&mut parsed_shapes, slide_part);
            let parsed_tables = parse_slide_tables(slide_xml)?;
            let parsed_image_refs = parse_slide_pictures(slide_xml)?;
            let parsed_chart_refs = parse_slide_charts(slide_xml)?;
            let parsed_transition = parse_slide_transition(slide_xml)?;
            let parsed_timing = parse_slide_timing(slide_xml)?;
            let parsed_unknown_children = parse_slide_unknown_children(slide_xml)?;
            let parsed_notes_text = parse_slide_notes_text(&package, &slide_uri, slide_part)?;
            let parsed_comments = parse_slide_comments(&package, &slide_uri, slide_part)?;
            let parsed_background = parse_slide_background(slide_xml)?;
            let parsed_hidden = parse_slide_hidden(slide_xml)?;
            let parsed_header_footer = parse_slide_header_footer(slide_xml)?;
            let parsed_grouped_shapes = parse_slide_grouped_shapes(slide_xml)?;

            let legacy_runs = extract_legacy_text_runs(parsed_shapes.first());
            let title = legacy_runs.first().cloned().unwrap_or_default();
            let mut slide = Slide::new(title);
            for text in legacy_runs.into_iter().skip(1) {
                slide.add_text_run(text);
            }

            for shape in parsed_shapes.into_iter().skip(1) {
                append_shape(&mut slide, shape);
            }

            for table in parsed_tables {
                append_table(&mut slide, table);
            }

            for parsed_image in parsed_image_refs {
                if let Some(image) =
                    resolve_slide_image(&package, &slide_uri, slide_part, parsed_image)?
                {
                    slide.push_image(image);
                }
            }

            for parsed_chart in parsed_chart_refs {
                if let Some(chart) =
                    resolve_slide_chart(&package, &slide_uri, slide_part, parsed_chart)?
                {
                    slide.push_chart(chart);
                }
            }

            for comment in parsed_comments {
                slide.push_comment(comment);
            }

            if let Some(transition) = parsed_transition {
                slide.set_transition(transition);
            }

            if let Some(timing) = parsed_timing {
                slide.set_timing(timing);
            }

            if let Some(notes_text) = parsed_notes_text {
                slide.set_notes_text(notes_text);
            }

            if let Some(background) = parsed_background {
                slide.set_background(background);
            }
            if parsed_hidden {
                slide.set_hidden(true);
            }
            if let Some(mut header_footer) = parsed_header_footer {
                // Feature #12: Extract footer_text and date_time_text from placeholder shapes.
                for shape in slide.shapes() {
                    if let Some(pk) = shape.placeholder_kind() {
                        match pk {
                            "ftr" => {
                                if header_footer.footer_text.is_none() {
                                    let text: String = shape
                                        .paragraphs()
                                        .iter()
                                        .flat_map(|p| p.runs().iter().map(|r| r.text().to_string()))
                                        .collect::<Vec<_>>()
                                        .join("");
                                    if !text.is_empty() {
                                        header_footer.footer_text = Some(text);
                                    }
                                }
                            }
                            "dt" => {
                                if header_footer.date_time_text.is_none() {
                                    let text: String = shape
                                        .paragraphs()
                                        .iter()
                                        .flat_map(|p| p.runs().iter().map(|r| r.text().to_string()))
                                        .collect::<Vec<_>>()
                                        .join("");
                                    if !text.is_empty() {
                                        header_footer.date_time_text = Some(text);
                                    }
                                }
                            }
                            _ => {}
                        }
                    }
                }
                slide.set_header_footer(header_footer);
            }
            for group in parsed_grouped_shapes {
                slide.push_grouped_shape(group);
            }

            if !parsed_unknown_children.is_empty() {
                slide.set_unknown_children(parsed_unknown_children);
            }

            // Capture extra namespace declarations from the original <p:sld> element
            // so that unknown elements/attributes using those prefixes remain valid
            // on dirty save.
            let extra_ns = parse_root_element_namespace_declarations(
                slide_part.data.as_bytes(),
                b"sld",
                &["xmlns:p", "xmlns:a", "xmlns:c", "xmlns:r"],
            );
            if !extra_ns.is_empty() {
                slide.set_extra_namespace_declarations(extra_ns);
            }

            slide.set_original_part_bytes(
                slide_uri.as_str().to_string(),
                slide_part.data.as_bytes().to_vec(),
            );

            slides.push(slide);
        }

        // Convert slide_masters metadata to SlideMaster for write API
        let slide_masters_v2 = slide_masters
            .iter()
            .map(|metadata| {
                let layouts = metadata
                    .layouts()
                    .iter()
                    .map(|layout_meta| {
                        let mut layout = SlideLayout::new(
                            layout_meta.name().unwrap_or("Untitled Layout"),
                            metadata.relationship_id(),
                            layout_meta.part_uri(),
                            layout_meta.relationship_id(),
                        );
                        if let Some(layout_type) = layout_meta.layout_type() {
                            layout.set_layout_type(layout_type);
                        }
                        layout.set_preserve(layout_meta.preserve().unwrap_or(false));
                        for shape in layout_meta.shapes() {
                            layout.add_shape(shape.clone());
                        }
                        // Store original XML if available
                        if let Some(part) = package.get_part(layout_meta.part_uri()) {
                            layout.set_original_xml(part.data.as_bytes().to_vec());
                        }
                        layout.mark_clean();
                        layout
                    })
                    .collect();
                SlideMaster::from_metadata(
                    metadata.relationship_id().to_string(),
                    metadata.part_uri().to_string(),
                    layouts,
                    metadata.shapes().to_vec(),
                )
            })
            .collect();

        // Parse presentation properties from docProps/core.xml and docProps/app.xml
        let mut presentation_properties =
            crate::presentation_properties::PresentationProperties::new();

        // Try to load core properties
        if let Some(core_part) = package.get_part("/docProps/core.xml") {
            let _ = presentation_properties.parse_core_xml(core_part.data.as_bytes());
        }

        // Try to load app properties
        if let Some(app_part) = package.get_part("/docProps/app.xml") {
            let _ = presentation_properties.parse_app_xml(app_part.data.as_bytes());
        }

        Ok(Self {
            package,
            slides,
            slide_masters,
            slide_masters_v2,
            slide_width_emu,
            slide_height_emu,
            theme_color_scheme,
            theme_font_scheme,
            sections,
            first_slide_number: pres_attrs.first_slide_number,
            show_special_pls_on_title_sld: pres_attrs.show_special_pls_on_title_sld,
            right_to_left: pres_attrs.right_to_left,
            custom_shows: Vec::new(),
            slide_show_settings: None,
            presentation_properties,
            dirty: false,
        })
    }

    pub fn save(&mut self, path: impl AsRef<Path>) -> Result<()> {
        let path = path.as_ref();
        if !self.dirty {
            self.package.save(path)?;
            return Ok(());
        }

        let mut package = self.package.clone();
        let mut presentation_passthrough_relationships = BTreeMap::<String, usize>::new();
        let mut slide_passthrough_relationships = BTreeMap::<String, usize>::new();
        // Preserve existing slide/master/layout/media/chart/notes/comment topology for
        // pass-through fidelity. We only replace parts this serializer owns fully.
        let _ = package.remove_part(PRESENTATION_PART_URI);

        let presentation_uri = PartUri::new(PRESENTATION_PART_URI)?;
        let mut presentation_part = Part::new_xml(presentation_uri.clone(), Vec::new());
        presentation_part.content_type = Some(ContentTypeValue::PRESENTATION.to_string());
        if let Some(original_presentation_part) = self.package.get_part(presentation_uri.as_str()) {
            for relationship in original_presentation_part.relationships.iter() {
                if !is_rebuilt_presentation_relationship_type(relationship.rel_type.as_str()) {
                    record_passthrough_relationship(
                        &mut presentation_passthrough_relationships,
                        relationship.rel_type.as_str(),
                    );
                    presentation_part.relationships.add(relationship.clone());
                }
            }
        }

        // Serialize slide masters and layouts.
        // If we have mutable masters (v2 API), use those. Otherwise create a default master.
        let slide_master_refs = if !self.slide_masters_v2.is_empty() {
            serialize_all_slide_masters_and_layouts(
                &mut package,
                &presentation_uri,
                &mut presentation_part,
                &self.slide_masters_v2,
            )?
        } else {
            // Fallback: create a single default master and layout for backward compatibility.
            let slide_master_part_uri = PartUri::new(DEFAULT_SLIDE_MASTER_PART_URI)?;
            let slide_layout_part_uri = PartUri::new(DEFAULT_SLIDE_LAYOUT_PART_URI)?;
            let slide_master_relationship = presentation_part.relationships.add_new(
                RelationshipType::SLIDE_MASTER.to_string(),
                relative_path_from_part(&presentation_uri, &slide_master_part_uri),
                TargetMode::Internal,
            );
            let slide_master_ref = SlideMasterRef {
                master_id: DEFAULT_SLIDE_MASTER_ID,
                relationship_id: slide_master_relationship.id.clone(),
            };

            let slide_master_metadata = attach_slide_master_and_layout_parts(
                &mut package,
                &slide_master_part_uri,
                &slide_layout_part_uri,
                slide_master_ref.relationship_id.as_str(),
            )?;
            self.slide_masters = vec![slide_master_metadata];

            vec![slide_master_ref]
        };

        let mut slide_refs = Vec::with_capacity(self.slides.len());
        for (index, _) in self.slides.iter().enumerate() {
            let slide_number = u32::try_from(index + 1).map_err(|_| {
                PptxError::UnsupportedPackage("too many slides to serialize".to_string())
            })?;
            let slide_id = 255_u32.checked_add(slide_number).ok_or_else(|| {
                PptxError::UnsupportedPackage("slide id overflow while serializing".to_string())
            })?;

            let target = format!("slides/slide{slide_number}.xml");
            let relationship = presentation_part.relationships.add_new(
                RelationshipType::SLIDE.to_string(),
                target,
                TargetMode::Internal,
            );
            let part_uri = presentation_uri.resolve_relative(relationship.target.as_str())?;

            slide_refs.push(SlideRef {
                slide_id,
                relationship_id: relationship.id.clone(),
                part_uri,
            });
        }

        presentation_part.data = PartData::Xml(serialize_presentation_xml(
            &slide_refs,
            &slide_master_refs,
            self.slide_width_emu,
            self.slide_height_emu,
            self.first_slide_number,
            self.show_special_pls_on_title_sld,
            self.right_to_left,
        )?);
        package.set_part(presentation_part);

        // Only add the package-level presentation relationship if not already present
        // (the cloned package may already have it from the original file).
        if package
            .relationships()
            .get_first_by_type(RelationshipType::WORKBOOK)
            .is_none()
        {
            package.relationships_mut().add_new(
                RelationshipType::WORKBOOK.to_string(),
                PRESENTATION_PART_URI.to_string(),
                TargetMode::Internal,
            );
        }

        let mut next_media_index = 1_u32;
        let mut next_chart_index = 1_u32;
        let mut serialized_comment_authors = collect_comment_authors(&self.slides)?;
        let has_comments = !serialized_comment_authors.is_empty();
        let comment_authors_part_uri = if has_comments {
            Some(PartUri::new(COMMENT_AUTHORS_PART_URI)?)
        } else {
            None
        };
        for (index, (slide, slide_ref)) in self.slides.iter().zip(slide_refs.iter()).enumerate() {
            let mut slide_part = Part::new_xml(slide_ref.part_uri.clone(), Vec::new());
            slide_part.content_type = Some(ContentTypeValue::SLIDE.to_string());
            let original_slide_part = self.package.get_part(slide_ref.part_uri.as_str());
            let can_passthrough_slide = !slide.dirty()
                && slide
                    .original_part_bytes()
                    .is_some_and(|(part_uri, _)| part_uri == slide_ref.part_uri.as_str());
            if can_passthrough_slide {
                if let Some(original_slide_part) = original_slide_part {
                    for relationship in original_slide_part.relationships.iter() {
                        slide_part.relationships.add(relationship.clone());
                    }
                }
                if let Some((_, original_bytes)) = slide.original_part_bytes() {
                    slide_part.data = PartData::Xml(original_bytes.to_vec());
                    package.set_part(slide_part);
                    continue;
                }
            }

            if let Some(original_slide_part) = original_slide_part {
                for relationship in original_slide_part.relationships.iter() {
                    if !is_rebuilt_slide_relationship_type(relationship.rel_type.as_str()) {
                        record_passthrough_relationship(
                            &mut slide_passthrough_relationships,
                            relationship.rel_type.as_str(),
                        );
                        slide_part.relationships.add(relationship.clone());
                    }
                }
            }

            let image_refs = attach_slide_image_parts(
                &mut package,
                &slide_ref.part_uri,
                &mut slide_part,
                slide.images(),
                &mut next_media_index,
            )?;
            let chart_refs = attach_slide_chart_parts(
                &mut package,
                &slide_ref.part_uri,
                &mut slide_part,
                slide.charts(),
                &mut next_chart_index,
            )?;
            if !slide.comments().is_empty() {
                let slide_number = u32::try_from(index + 1).map_err(|_| {
                    PptxError::UnsupportedPackage(
                        "too many comment slides to serialize".to_string(),
                    )
                })?;
                attach_slide_comments_part(
                    &mut package,
                    &slide_ref.part_uri,
                    &mut slide_part,
                    slide_number,
                    slide.comments(),
                    &mut serialized_comment_authors,
                    comment_authors_part_uri.as_ref().ok_or_else(|| {
                        PptxError::UnsupportedPackage(
                            "missing comment authors part while serializing comments".to_string(),
                        )
                    })?,
                )?;
            }
            if let Some(notes_text) = slide.notes_text() {
                let slide_number = u32::try_from(index + 1).map_err(|_| {
                    PptxError::UnsupportedPackage("too many notes slides to serialize".to_string())
                })?;
                attach_slide_notes_part(
                    &mut package,
                    &slide_ref.part_uri,
                    &mut slide_part,
                    slide_number,
                    notes_text,
                )?;
            }
            // Link slide to its layout.
            // Use slide's layout reference if set, otherwise use first layout of first master.
            let layout_part_uri_str =
                if let Some((master_idx, layout_idx)) = slide.layout_reference() {
                    // Slide has explicit layout reference
                    self.slide_masters_v2
                        .get(master_idx)
                        .and_then(|master| master.layouts().get(layout_idx))
                        .map(|layout| layout.part_uri())
                        .unwrap_or(DEFAULT_SLIDE_LAYOUT_PART_URI)
                } else if !self.slide_masters_v2.is_empty() {
                    // Use first layout of first master by default
                    self.slide_masters_v2[0]
                        .layouts()
                        .first()
                        .map(|l| l.part_uri())
                        .unwrap_or(DEFAULT_SLIDE_LAYOUT_PART_URI)
                } else {
                    // Fallback to default layout
                    DEFAULT_SLIDE_LAYOUT_PART_URI
                };
            let layout_part_uri = PartUri::new(layout_part_uri_str)?;

            slide_part.relationships.add_new(
                RelationshipType::SLIDE_LAYOUT.to_string(),
                relative_path_from_part(&slide_ref.part_uri, &layout_part_uri),
                TargetMode::Internal,
            );
            slide_part.data = PartData::Xml(serialize_slide_xml(slide, &image_refs, &chart_refs)?);

            package.set_part(slide_part);
        }

        if let Some(comment_authors_part_uri) = comment_authors_part_uri {
            let mut comment_authors_part = Part::new_xml(
                comment_authors_part_uri,
                serialize_comment_authors_xml(&serialized_comment_authors)?,
            );
            comment_authors_part.content_type = Some(COMMENT_AUTHORS_CONTENT_TYPE.to_string());
            package.set_part(comment_authors_part);
        }

        emit_passthrough_relationship_warnings(
            "presentation",
            &presentation_passthrough_relationships,
        );
        emit_passthrough_relationship_warnings("slide", &slide_passthrough_relationships);

        package.save(path)?;
        self.package = package;
        self.dirty = false;
        Ok(())
    }

    pub fn add_slide(&mut self) -> &mut Slide {
        self.add_slide_with_title("")
    }

    pub fn add_slide_with_title(&mut self, title: impl Into<String>) -> &mut Slide {
        self.dirty = true;
        let index = self.slides.len();
        self.slides.push(Slide::new(title));
        &mut self.slides[index]
    }

    pub fn slides(&self) -> &[Slide] {
        &self.slides
    }

    /// Gets a mutable reference to all slides.
    pub fn slides_mut(&mut self) -> &mut Vec<Slide> {
        self.dirty = true;
        &mut self.slides
    }

    pub fn slide(&self, index: usize) -> Option<&Slide> {
        self.slides.get(index)
    }

    pub fn slide_mut(&mut self, index: usize) -> Option<&mut Slide> {
        self.dirty = true;
        self.slides.get_mut(index)
    }

    /// Search all slides for text occurrences.
    ///
    /// Returns tuples of `(slide_index, shape_index, paragraph_index, run_index)`.
    pub fn find_text(&self, needle: &str) -> Vec<(usize, usize, usize, usize)> {
        let mut results = Vec::new();
        for (slide_idx, slide) in self.slides.iter().enumerate() {
            for (si, pi, ri) in slide.find_text(needle) {
                results.push((slide_idx, si, pi, ri));
            }
        }
        results
    }

    /// Replace all occurrences of `old` with `new` across all slides.
    ///
    /// Returns the total number of replacements made.
    pub fn replace_text(&mut self, old: &str, new: &str) -> usize {
        let mut count = 0;
        for slide in &mut self.slides {
            count += slide.replace_text(old, new);
        }
        if count > 0 {
            self.dirty = true;
        }
        count
    }

    pub fn remove_slide(&mut self, index: usize) -> Option<Slide> {
        if index >= self.slides.len() {
            return None;
        }

        self.dirty = true;
        Some(self.slides.remove(index))
    }

    /// Clones a slide at the given index and appends the copy to the end.
    ///
    /// Returns the index of the new slide, or `None` if `index` is out of range.
    pub fn clone_slide(&mut self, index: usize) -> Option<usize> {
        if index >= self.slides.len() {
            return None;
        }
        let cloned = self.slides[index].clone();
        self.slides.push(cloned);
        self.dirty = true;
        Some(self.slides.len() - 1)
    }

    /// Moves a slide from `from_index` to `to_index`.
    pub fn move_slide(&mut self, from_index: usize, to_index: usize) -> bool {
        if from_index >= self.slides.len() || to_index >= self.slides.len() {
            return false;
        }
        if from_index == to_index {
            return true;
        }
        let slide = self.slides.remove(from_index);
        self.slides.insert(to_index, slide);
        self.dirty = true;
        true
    }

    pub fn slide_count(&self) -> usize {
        self.slides.len()
    }

    /// Returns read-only slide master metadata (legacy API).
    pub fn slide_masters(&self) -> &[SlideMasterMetadata] {
        self.slide_masters.as_slice()
    }

    // ── Slide Master Write API ──

    /// Returns a reference to mutable slide masters.
    pub fn slide_masters_v2(&self) -> &[SlideMaster] {
        &self.slide_masters_v2
    }

    /// Returns a mutable reference to mutable slide masters.
    pub fn slide_masters_mut(&mut self) -> &mut Vec<SlideMaster> {
        self.dirty = true;
        &mut self.slide_masters_v2
    }

    /// Returns a reference to a slide master by index.
    pub fn slide_master(&self, index: usize) -> Option<&SlideMaster> {
        self.slide_masters_v2.get(index)
    }

    /// Returns a mutable reference to a slide master by index.
    pub fn slide_master_mut(&mut self, index: usize) -> Option<&mut SlideMaster> {
        self.dirty = true;
        self.slide_masters_v2.get_mut(index)
    }

    /// Adds a new slide master to the presentation.
    pub fn add_slide_master(&mut self, master: SlideMaster) {
        self.slide_masters_v2.push(master);
        self.dirty = true;
    }

    /// Removes a slide master by index.
    ///
    /// Returns `Some(master)` if a master was removed, or `None` if the index was out of bounds.
    pub fn remove_slide_master(&mut self, index: usize) -> Option<SlideMaster> {
        if index < self.slide_masters_v2.len() {
            self.dirty = true;
            Some(self.slide_masters_v2.remove(index))
        } else {
            None
        }
    }

    // ── Slide Layout Write API ──

    /// Returns a reference to all slide layouts across all masters.
    ///
    /// This provides a flat view of all layouts in the presentation.
    /// To access layouts for a specific master, use `slide_master(index).layouts()`.
    pub fn layouts(&self) -> Vec<&SlideLayout> {
        self.slide_masters_v2
            .iter()
            .flat_map(|master| master.layouts())
            .collect()
    }

    /// Returns a mutable reference to all slide layouts via their masters.
    ///
    /// Note: This returns mutable references to masters, as layouts are owned by masters.
    /// To modify layouts, access them through `slide_master_mut(index).layouts_mut()`.
    pub fn layouts_mut(&mut self) -> &mut Vec<SlideMaster> {
        self.dirty = true;
        &mut self.slide_masters_v2
    }

    /// Returns a reference to a specific layout by master index and layout index.
    ///
    /// # Arguments
    ///
    /// * `master_index` - The index of the slide master
    /// * `layout_index` - The index of the layout within that master
    pub fn layout(&self, master_index: usize, layout_index: usize) -> Option<&SlideLayout> {
        self.slide_masters_v2
            .get(master_index)
            .and_then(|master| master.layout(layout_index))
    }

    /// Returns a mutable reference to a specific layout by master index and layout index.
    ///
    /// # Arguments
    ///
    /// * `master_index` - The index of the slide master
    /// * `layout_index` - The index of the layout within that master
    pub fn layout_mut(
        &mut self,
        master_index: usize,
        layout_index: usize,
    ) -> Option<&mut SlideLayout> {
        self.dirty = true;
        self.slide_masters_v2
            .get_mut(master_index)
            .and_then(|master| master.layout_mut(layout_index))
    }

    /// Adds a layout to a specific slide master.
    ///
    /// # Arguments
    ///
    /// * `master_index` - The index of the slide master to add the layout to
    /// * `layout` - The layout to add
    ///
    /// Returns `Ok(())` if successful, or `Err` if the master index is out of bounds.
    pub fn add_layout(
        &mut self,
        master_index: usize,
        layout: SlideLayout,
    ) -> std::result::Result<(), &'static str> {
        if let Some(master) = self.slide_masters_v2.get_mut(master_index) {
            master.add_layout(layout);
            self.dirty = true;
            Ok(())
        } else {
            Err("Master index out of bounds")
        }
    }

    /// Removes a layout from a specific slide master.
    ///
    /// # Arguments
    ///
    /// * `master_index` - The index of the slide master
    /// * `layout_index` - The index of the layout within that master
    ///
    /// Returns `Some(layout)` if a layout was removed, or `None` if indices were out of bounds.
    pub fn remove_layout(
        &mut self,
        master_index: usize,
        layout_index: usize,
    ) -> Option<SlideLayout> {
        self.slide_masters_v2
            .get_mut(master_index)
            .and_then(|master| {
                self.dirty = true;
                master.remove_layout(layout_index)
            })
    }

    // ── Feature #9: Slide size ──

    /// Slide width in EMUs. Standard 10" = 9144000.
    pub fn slide_width_emu(&self) -> Option<i64> {
        self.slide_width_emu
    }

    /// Set slide width in EMUs.
    pub fn set_slide_width_emu(&mut self, width: i64) {
        self.slide_width_emu = Some(width);
        self.dirty = true;
    }

    /// Slide height in EMUs. Standard 7.5" = 6858000.
    pub fn slide_height_emu(&self) -> Option<i64> {
        self.slide_height_emu
    }

    /// Set slide height in EMUs.
    pub fn set_slide_height_emu(&mut self, height: i64) {
        self.slide_height_emu = Some(height);
        self.dirty = true;
    }

    // ── Feature #14: Theme colors ──

    /// Theme color scheme.
    pub fn theme_color_scheme(&self) -> Option<&ThemeColorScheme> {
        self.theme_color_scheme.as_ref()
    }

    /// Set theme color scheme.
    pub fn set_theme_color_scheme(&mut self, scheme: ThemeColorScheme) {
        self.theme_color_scheme = Some(scheme);
        self.dirty = true;
    }

    // ── Theme fonts ──

    /// Theme font scheme.
    pub fn theme_font_scheme(&self) -> Option<&crate::theme::ThemeFontScheme> {
        self.theme_font_scheme.as_ref()
    }

    /// Set theme font scheme.
    pub fn set_theme_font_scheme(&mut self, scheme: crate::theme::ThemeFontScheme) {
        self.theme_font_scheme = Some(scheme);
        self.dirty = true;
    }

    // ── Sections ──

    /// Presentation sections.
    pub fn sections(&self) -> &[crate::slide::PresentationSection] {
        &self.sections
    }

    /// Mutable access to sections.
    pub fn sections_mut(&mut self) -> &mut Vec<crate::slide::PresentationSection> {
        self.dirty = true;
        &mut self.sections
    }

    /// Add a presentation section.
    pub fn add_section(
        &mut self,
        name: impl Into<String>,
    ) -> &mut crate::slide::PresentationSection {
        self.dirty = true;
        self.sections
            .push(crate::slide::PresentationSection::new(name));
        let index = self.sections.len().saturating_sub(1);
        &mut self.sections[index]
    }

    // ── Presentation properties ──

    /// First slide number (`firstSlideNum` attr on `<p:presentation>`).
    pub fn first_slide_number(&self) -> Option<u32> {
        self.first_slide_number
    }

    /// Set first slide number.
    pub fn set_first_slide_number(&mut self, number: u32) {
        self.first_slide_number = Some(number);
        self.dirty = true;
    }

    /// Clear first slide number.
    pub fn clear_first_slide_number(&mut self) {
        self.first_slide_number = None;
        self.dirty = true;
    }

    /// Whether special placeholders are shown on the title slide.
    pub fn show_special_pls_on_title_sld(&self) -> Option<bool> {
        self.show_special_pls_on_title_sld
    }

    /// Set whether special placeholders are shown on the title slide.
    pub fn set_show_special_pls_on_title_sld(&mut self, show: bool) {
        self.show_special_pls_on_title_sld = Some(show);
        self.dirty = true;
    }

    /// Clear show special placeholders on title slide.
    pub fn clear_show_special_pls_on_title_sld(&mut self) {
        self.show_special_pls_on_title_sld = None;
        self.dirty = true;
    }

    /// Right-to-left presentation (`rtl` attr on `<p:presentation>`).
    pub fn right_to_left(&self) -> Option<bool> {
        self.right_to_left
    }

    /// Set right-to-left.
    pub fn set_right_to_left(&mut self, rtl: bool) {
        self.right_to_left = Some(rtl);
        self.dirty = true;
    }

    /// Clear right-to-left.
    pub fn clear_right_to_left(&mut self) {
        self.right_to_left = None;
        self.dirty = true;
    }

    // ── Custom shows ──

    /// Gets the list of custom shows.
    pub fn custom_shows(&self) -> &[crate::custom_show::CustomShow] {
        &self.custom_shows
    }

    /// Gets a mutable reference to the custom shows list.
    pub fn custom_shows_mut(&mut self) -> &mut Vec<crate::custom_show::CustomShow> {
        self.dirty = true;
        &mut self.custom_shows
    }

    /// Adds a custom show.
    pub fn add_custom_show(&mut self, show: crate::custom_show::CustomShow) {
        self.custom_shows.push(show);
        self.dirty = true;
    }

    /// Removes a custom show by index.
    pub fn remove_custom_show(&mut self, index: usize) -> Option<crate::custom_show::CustomShow> {
        if index < self.custom_shows.len() {
            self.dirty = true;
            Some(self.custom_shows.remove(index))
        } else {
            None
        }
    }

    /// Finds a custom show by name.
    pub fn custom_show_by_name(&self, name: &str) -> Option<&crate::custom_show::CustomShow> {
        self.custom_shows.iter().find(|s| s.name() == name)
    }

    // ── Presentation properties ──

    /// Gets the presentation properties (metadata).
    pub fn properties(&self) -> &crate::presentation_properties::PresentationProperties {
        &self.presentation_properties
    }

    /// Gets a mutable reference to the presentation properties.
    pub fn properties_mut(
        &mut self,
    ) -> &mut crate::presentation_properties::PresentationProperties {
        self.dirty = true;
        &mut self.presentation_properties
    }

    // ── Slide show settings ──

    /// Gets the slide show settings.
    pub fn slide_show_settings(&self) -> Option<&crate::custom_show::SlideShowSettings> {
        self.slide_show_settings.as_ref()
    }

    /// Sets the slide show settings.
    pub fn set_slide_show_settings(&mut self, settings: crate::custom_show::SlideShowSettings) {
        self.slide_show_settings = Some(settings);
        self.dirty = true;
    }

    /// Clears the slide show settings.
    pub fn clear_slide_show_settings(&mut self) {
        self.slide_show_settings = None;
        self.dirty = true;
    }
}

impl Default for Presentation {
    fn default() -> Self {
        Self::new()
    }
}

fn resolve_presentation_part_uri(package: &Package) -> Result<PartUri> {
    for relationship in package
        .relationships()
        .get_by_type(RelationshipType::WORKBOOK)
    {
        if relationship.target_mode != TargetMode::Internal {
            continue;
        }
        let part_uri = normalize_relationship_target(relationship.target.as_str())?;
        if package.get_part(part_uri.as_str()).is_some() {
            return Ok(part_uri);
        }
    }

    let fallback = PartUri::new(PRESENTATION_PART_URI)?;
    if package.get_part(fallback.as_str()).is_some() {
        return Ok(fallback);
    }

    Err(PptxError::UnsupportedPackage(
        "presentation part not found".to_string(),
    ))
}

fn normalize_relationship_target(target: &str) -> Result<PartUri> {
    let mut normalized = target.trim().replace('\\', "/");
    while let Some(stripped) = normalized.strip_prefix("./") {
        normalized = stripped.to_string();
    }

    if !normalized.starts_with('/') {
        normalized.insert(0, '/');
    }

    PartUri::new(normalized).map_err(Into::into)
}

fn is_rebuilt_presentation_relationship_type(rel_type: &str) -> bool {
    rel_type == RelationshipType::SLIDE || rel_type == RelationshipType::SLIDE_MASTER
}

fn is_rebuilt_slide_relationship_type(rel_type: &str) -> bool {
    matches!(
        rel_type,
        RelationshipType::IMAGE
            | RelationshipType::CHART
            | RelationshipType::SLIDE_LAYOUT
            | NOTES_SLIDE_RELATIONSHIP_TYPE
            | COMMENTS_RELATIONSHIP_TYPE
            | COMMENT_AUTHORS_RELATIONSHIP_TYPE
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

fn parse_presentation_xml(xml: &[u8]) -> Result<ParsedPresentationRefs> {
    let mut reader = Reader::from_reader(Cursor::new(xml));
    reader.config_mut().trim_text(true);
    let mut buffer = Vec::new();
    let mut slide_refs = Vec::new();
    let mut slide_master_refs = Vec::new();

    loop {
        match reader.read_event_into(&mut buffer)? {
            Event::Empty(ref event) | Event::Start(ref event)
                if local_name(event.name().as_ref()) == b"sldId"
                    && has_presentation_prefix(event.name().as_ref()) =>
            {
                let relationship_id =
                    get_relationship_id_attribute_value(event).ok_or_else(|| {
                        PptxError::UnsupportedPackage(
                            "presentation slide id missing relationship `id` attribute".to_string(),
                        )
                    })?;
                slide_refs.push(ParsedSlideRef { relationship_id });
            }
            Event::Empty(ref event) | Event::Start(ref event)
                if local_name(event.name().as_ref()) == b"sldMasterId"
                    && has_presentation_prefix(event.name().as_ref()) =>
            {
                let relationship_id =
                    get_relationship_id_attribute_value(event).ok_or_else(|| {
                        PptxError::UnsupportedPackage(
                            "presentation slide master id missing relationship `id` attribute"
                                .to_string(),
                        )
                    })?;
                slide_master_refs.push(ParsedSlideMasterRef { relationship_id });
            }
            Event::Eof => break,
            _ => {}
        }
        buffer.clear();
    }

    Ok(ParsedPresentationRefs {
        slide_refs,
        slide_master_refs,
    })
}

fn parse_slide_master_xml(xml: &[u8]) -> Result<ParsedSlideMasterMetadata> {
    let mut reader = Reader::from_reader(Cursor::new(xml));
    reader.config_mut().trim_text(true);

    let mut buffer = Vec::new();
    let mut layout_refs = Vec::new();

    loop {
        match reader.read_event_into(&mut buffer)? {
            Event::Empty(ref event) | Event::Start(ref event)
                if local_name(event.name().as_ref()) == b"sldLayoutId"
                    && has_presentation_prefix(event.name().as_ref()) =>
            {
                let relationship_id =
                    get_relationship_id_attribute_value(event).ok_or_else(|| {
                        PptxError::UnsupportedPackage(
                            "slide master layout id missing relationship `id` attribute"
                                .to_string(),
                        )
                    })?;
                layout_refs.push(ParsedSlideLayoutRef { relationship_id });
            }
            Event::Eof => break,
            _ => {}
        }
        buffer.clear();
    }

    let shapes = parse_slide_shapes(xml)?;
    Ok(ParsedSlideMasterMetadata {
        layout_refs,
        shapes,
    })
}

fn parse_slide_layout_xml_metadata(xml: &[u8]) -> Result<ParsedSlideLayoutMetadata> {
    let mut reader = Reader::from_reader(Cursor::new(xml));
    reader.config_mut().trim_text(true);

    let mut buffer = Vec::new();
    let mut name = None;
    let mut layout_type = None;
    let mut preserve = None;

    loop {
        match reader.read_event_into(&mut buffer)? {
            Event::Empty(ref event) | Event::Start(ref event)
                if local_name(event.name().as_ref()) == b"sldLayout" =>
            {
                name = get_attribute_value(event, b"name");
                layout_type = get_attribute_value(event, b"type");
                preserve = get_attribute_value(event, b"preserve")
                    .as_deref()
                    .and_then(parse_xml_bool);
                break;
            }
            Event::Eof => break,
            _ => {}
        }
        buffer.clear();
    }

    let shapes = parse_slide_shapes(xml)?;
    Ok(ParsedSlideLayoutMetadata {
        name,
        layout_type,
        preserve,
        shapes,
    })
}

fn resolve_slide_masters(
    package: &Package,
    presentation_uri: &PartUri,
    presentation_part: &Part,
    parsed_slide_master_refs: &[ParsedSlideMasterRef],
) -> Result<Vec<SlideMasterMetadata>> {
    let mut slide_masters = Vec::with_capacity(parsed_slide_master_refs.len());

    for parsed_slide_master_ref in parsed_slide_master_refs {
        let relationship = presentation_part
            .relationships
            .get_by_id(parsed_slide_master_ref.relationship_id.as_str())
            .ok_or_else(|| {
                PptxError::UnsupportedPackage(format!(
                    "missing slide master relationship `{}`",
                    parsed_slide_master_ref.relationship_id
                ))
            })?;
        if relationship.target_mode != TargetMode::Internal {
            return Err(PptxError::UnsupportedPackage(format!(
                "slide master relationship `{}` is external",
                relationship.id
            )));
        }

        let slide_master_uri = presentation_uri.resolve_relative(relationship.target.as_str())?;
        let slide_master_part = package.get_part(slide_master_uri.as_str()).ok_or_else(|| {
            PptxError::UnsupportedPackage(format!(
                "missing slide master part `{}` for relationship `{}`",
                slide_master_uri.as_str(),
                relationship.id
            ))
        })?;
        let parsed_master_metadata = parse_slide_master_xml(slide_master_part.data.as_bytes())?;
        let ParsedSlideMasterMetadata {
            layout_refs: parsed_layout_refs,
            shapes: parsed_master_shapes,
        } = parsed_master_metadata;
        let mut layouts = Vec::with_capacity(parsed_layout_refs.len());
        for parsed_layout_ref in parsed_layout_refs {
            let layout_relationship = slide_master_part
                .relationships
                .get_by_id(parsed_layout_ref.relationship_id.as_str())
                .ok_or_else(|| {
                    PptxError::UnsupportedPackage(format!(
                        "missing slide layout relationship `{}` from slide master `{}`",
                        parsed_layout_ref.relationship_id,
                        slide_master_uri.as_str()
                    ))
                })?;
            if layout_relationship.target_mode != TargetMode::Internal {
                return Err(PptxError::UnsupportedPackage(format!(
                    "slide layout relationship `{}` is external",
                    layout_relationship.id
                )));
            }
            let slide_layout_uri =
                slide_master_uri.resolve_relative(layout_relationship.target.as_str())?;
            let slide_layout_part =
                package.get_part(slide_layout_uri.as_str()).ok_or_else(|| {
                    PptxError::UnsupportedPackage(format!(
                        "missing slide layout part `{}` for relationship `{}`",
                        slide_layout_uri.as_str(),
                        layout_relationship.id
                    ))
                })?;
            let _ = resolve_slide_layout_master_part_uri(
                package,
                &slide_layout_uri,
                slide_layout_part,
            )?;
            let parsed_layout_metadata =
                parse_slide_layout_xml_metadata(slide_layout_part.data.as_bytes())?;
            let ParsedSlideLayoutMetadata {
                name,
                layout_type,
                preserve,
                shapes,
            } = parsed_layout_metadata;

            layouts.push(SlideLayoutMetadata {
                relationship_id: layout_relationship.id.clone(),
                part_uri: slide_layout_uri.as_str().to_string(),
                name,
                layout_type,
                preserve,
                shapes,
            });
        }

        slide_masters.push(SlideMasterMetadata {
            relationship_id: relationship.id.clone(),
            part_uri: slide_master_uri.as_str().to_string(),
            layouts,
            shapes: parsed_master_shapes,
        });
    }

    Ok(slide_masters)
}

fn resolve_slide_layout_part_uri(
    package: &Package,
    slide_uri: &PartUri,
    slide_part: &Part,
) -> Result<Option<PartUri>> {
    let Some(layout_relationship) = slide_part
        .relationships
        .get_first_by_type(RelationshipType::SLIDE_LAYOUT)
    else {
        return Ok(None);
    };
    if layout_relationship.target_mode != TargetMode::Internal {
        return Err(PptxError::UnsupportedPackage(format!(
            "slide layout relationship `{}` is external",
            layout_relationship.id
        )));
    }

    let slide_layout_uri = slide_uri.resolve_relative(layout_relationship.target.as_str())?;
    let slide_layout_part = package.get_part(slide_layout_uri.as_str()).ok_or_else(|| {
        PptxError::UnsupportedPackage(format!(
            "missing slide layout part `{}` for relationship `{}`",
            slide_layout_uri.as_str(),
            layout_relationship.id
        ))
    })?;
    let _ = resolve_slide_layout_master_part_uri(package, &slide_layout_uri, slide_layout_part)?;

    Ok(Some(slide_layout_uri))
}

fn resolve_slide_layout_master_part_uri(
    package: &Package,
    slide_layout_uri: &PartUri,
    slide_layout_part: &Part,
) -> Result<Option<PartUri>> {
    let Some(master_relationship) = slide_layout_part
        .relationships
        .get_first_by_type(RelationshipType::SLIDE_MASTER)
    else {
        return Ok(None);
    };
    if master_relationship.target_mode != TargetMode::Internal {
        return Err(PptxError::UnsupportedPackage(format!(
            "slide master relationship `{}` from layout part is external",
            master_relationship.id
        )));
    }

    let slide_master_uri =
        slide_layout_uri.resolve_relative(master_relationship.target.as_str())?;
    if package.get_part(slide_master_uri.as_str()).is_none() {
        return Err(PptxError::UnsupportedPackage(format!(
            "missing slide master part `{}` for relationship `{}`",
            slide_master_uri.as_str(),
            master_relationship.id
        )));
    }

    Ok(Some(slide_master_uri))
}

fn parse_slide_shapes(xml: &[u8]) -> Result<Vec<Shape>> {
    let mut reader = Reader::from_reader(Cursor::new(xml));
    reader.config_mut().trim_text(false);

    let mut buffer = Vec::new();
    let mut shapes = Vec::new();

    let mut in_shape = false;
    let mut shape_depth = 0_usize;
    let mut current_shape_name = String::new();
    let mut current_shape_paragraphs: Vec<ParsedParagraphData> = Vec::new();
    let mut current_shape_type = ShapeType::AutoShape;
    let mut current_placeholder_kind: Option<String> = None;
    let mut current_placeholder_idx: Option<u32> = None;
    let mut current_shape_xfrm_offset: Option<(i64, i64)> = None;
    let mut current_shape_xfrm_extents: Option<(i64, i64)> = None;
    let mut current_shape_preset_geometry: Option<String> = None;
    let mut current_shape_solid_fill_srgb: Option<String> = None;
    let mut current_shape_solid_fill_alpha: Option<u8> = None;
    let mut current_shape_solid_fill_color: Option<ShapeColor> = None;
    let mut current_shape_outline = ShapeOutline::default();
    // Tracks the ShapeColor currently being built from srgbClr/schemeClr parsing.
    // This is a shared variable used across all color contexts. The context flags
    // (in_shape_solid_fill, in_outline_solid_fill, etc.) determine where the
    // finished color gets stored.
    let mut building_color: Option<ShapeColor> = None;
    let mut current_shape_gradient_fill: Option<GradientFill> = None;
    let mut current_shape_pattern_fill: Option<PatternFill> = None;
    let mut current_shape_picture_fill: Option<PictureFill> = None;
    let mut current_shape_no_fill = false;
    let mut current_shape_rotation: Option<i32> = None;
    let mut current_shape_flip_h = false;
    let mut current_shape_flip_v = false;
    let mut current_shape_custom_geometry_raw: Option<RawXmlNode> = None;
    let mut current_shape_preset_geometry_adjustments: Option<RawXmlNode> = None;
    let mut current_shape_hidden = false;
    let mut current_shape_alt_text: Option<String> = None;
    let mut current_shape_alt_text_title: Option<String> = None;
    let mut current_shape_is_smartart = false;
    let mut current_shape_is_connector = false;
    let mut current_shape_start_connection: Option<ConnectionInfo> = None;
    let mut current_shape_end_connection: Option<ConnectionInfo> = None;
    let mut current_shape_media: Option<(MediaType, String)> = None;
    let mut current_shape_shadow: Option<ShapeShadow> = None;
    let mut current_shape_glow: Option<ShapeGlow> = None;
    let mut current_shape_reflection: Option<ShapeReflection> = None;
    let mut current_shape_text_anchor: Option<TextAnchor> = None;
    let mut current_shape_auto_fit: Option<AutoFitType> = None;
    let mut current_shape_text_direction: Option<TextDirection> = None;
    let mut current_shape_text_columns: Option<u32> = None;
    let mut current_shape_text_column_spacing: Option<i64> = None;
    let mut current_shape_text_inset_left: Option<i64> = None;
    let mut current_shape_text_inset_right: Option<i64> = None;
    let mut current_shape_text_inset_top: Option<i64> = None;
    let mut current_shape_text_inset_bottom: Option<i64> = None;
    let mut current_shape_word_wrap: Option<bool> = None;
    let mut current_shape_body_pr_rot: Option<i32> = None;
    let mut current_shape_body_pr_rtl_col: Option<bool> = None;
    let mut current_shape_body_pr_from_word_art: Option<bool> = None;
    let mut current_shape_body_pr_force_aa: Option<bool> = None;
    let mut current_shape_body_pr_compat_ln_spc: Option<bool> = None;
    let mut current_shape_unknown_attrs: Vec<(String, String)> = Vec::new();
    let mut current_shape_unknown_children: Vec<RawXmlNode> = Vec::new();
    let mut current_lst_style_raw: Option<RawXmlNode> = None;
    let mut current_body_pr_unknown_children: Vec<RawXmlNode> = Vec::new();
    let mut current_body_pr_unknown_attrs: Vec<(String, String)> = Vec::new();
    let mut current_paragraph_runs: Option<Vec<ParsedRunData>> = None;
    let mut current_paragraph_properties = ParagraphProperties::default();
    let mut current_run_properties = RunProperties::default();
    let mut in_rpr = false;
    let mut in_rpr_solid_fill = false;
    // Tracks which spacing parent we're inside: "lnSpc", "spcBef", or "spcAft".
    let mut in_spacing_parent: Option<&'static str> = None;
    let mut in_tx_body = false;
    let mut in_text = false;
    let mut current_text = String::new();
    let mut in_sp_pr = false;
    let mut sp_pr_depth = 0_usize;
    let mut in_xfrm = false;
    let mut in_shape_solid_fill = false;
    let mut in_outline = false;
    let mut in_outline_solid_fill = false;
    let mut in_grad_fill = false;
    let mut in_grad_gs = false;
    let mut grad_gs_pos: u32 = 0;
    let mut in_patt_fill = false;
    let mut in_patt_fg = false;
    let mut in_patt_bg = false;
    let mut in_blip_fill = false;
    let mut in_effect_lst = false;
    let mut in_outer_shdw = false;
    let mut in_glow = false;
    let mut in_body_pr = false;
    let mut in_ppr = false;
    let mut in_bu_clr = false;

    loop {
        match reader.read_event_into(&mut buffer)? {
            Event::Start(ref event) => {
                let event_name = event.name();
                let local = local_name(event_name.as_ref());

                if local == b"sp" || local == b"cxnSp" {
                    if !in_shape {
                        in_shape = true;
                        shape_depth = 1;
                        current_shape_name = format!("Shape {}", shapes.len() + 1);
                        current_shape_paragraphs.clear();
                        current_shape_type = ShapeType::AutoShape;
                        current_placeholder_kind = None;
                        current_placeholder_idx = None;
                        current_shape_xfrm_offset = None;
                        current_shape_xfrm_extents = None;
                        current_shape_preset_geometry = None;
                        current_shape_solid_fill_srgb = None;
                        current_shape_solid_fill_alpha = None;
                        current_shape_solid_fill_color = None;
                        current_shape_outline = ShapeOutline::default();
                        current_shape_gradient_fill = None;
                        current_shape_pattern_fill = None;
                        building_color = None;
                        current_shape_picture_fill = None;
                        current_shape_no_fill = false;
                        current_shape_rotation = None;
                        current_shape_flip_h = false;
                        current_shape_flip_v = false;
                        current_shape_custom_geometry_raw = None;
                        current_shape_preset_geometry_adjustments = None;
                        current_shape_hidden = false;
                        current_shape_alt_text = None;
                        current_shape_alt_text_title = None;
                        current_shape_is_smartart = false;
                        // Feature #9: connector shapes.
                        current_shape_is_connector = local == b"cxnSp";
                        current_shape_start_connection = None;
                        current_shape_end_connection = None;
                        current_shape_media = None;
                        current_shape_shadow = None;
                        current_shape_glow = None;
                        current_shape_reflection = None;
                        current_shape_text_anchor = None;
                        current_shape_auto_fit = None;
                        current_shape_text_direction = None;
                        current_shape_text_columns = None;
                        current_shape_text_column_spacing = None;
                        current_shape_text_inset_left = None;
                        current_shape_text_inset_right = None;
                        current_shape_text_inset_top = None;
                        current_shape_text_inset_bottom = None;
                        current_shape_word_wrap = None;
                        current_shape_body_pr_rot = None;
                        current_shape_body_pr_rtl_col = None;
                        current_shape_body_pr_from_word_art = None;
                        current_shape_body_pr_force_aa = None;
                        current_shape_body_pr_compat_ln_spc = None;
                        current_shape_unknown_attrs.clear();
                        current_shape_unknown_children.clear();
                        for attribute in event.attributes().flatten() {
                            current_shape_unknown_attrs.push((
                                String::from_utf8_lossy(attribute.key.as_ref()).into_owned(),
                                String::from_utf8_lossy(attribute.value.as_ref()).into_owned(),
                            ));
                        }
                        current_paragraph_runs = None;
                        in_tx_body = false;
                        in_text = false;
                        current_text.clear();
                        in_sp_pr = false;
                        sp_pr_depth = 0;
                        in_xfrm = false;
                        in_shape_solid_fill = false;
                        in_outline = false;
                        in_outline_solid_fill = false;
                        in_grad_fill = false;
                        in_grad_gs = false;
                        in_patt_fill = false;
                        in_patt_fg = false;
                        in_patt_bg = false;
                        in_blip_fill = false;
                        in_effect_lst = false;
                        in_outer_shdw = false;
                        in_glow = false;
                        in_body_pr = false;
                        in_ppr = false;
                        in_bu_clr = false;
                    } else {
                        shape_depth = shape_depth.saturating_add(1);
                    }
                    buffer.clear();
                    continue;
                }

                if !in_shape {
                    buffer.clear();
                    continue;
                }

                shape_depth = shape_depth.saturating_add(1);

                match local {
                    b"cNvPr" => {
                        if let Some(name) = get_attribute_value(event, b"name") {
                            current_shape_name = name;
                        }
                        // Feature #4: hidden attribute.
                        if get_attribute_value(event, b"hidden")
                            .as_deref()
                            .and_then(parse_xml_bool)
                            .unwrap_or(false)
                        {
                            current_shape_hidden = true;
                        }
                        // Feature #10: alt text.
                        current_shape_alt_text = get_attribute_value(event, b"descr");
                        current_shape_alt_text_title = get_attribute_value(event, b"title");
                    }
                    b"cNvSpPr" | b"cNvCxnSpPr" => {
                        current_shape_type = parse_shape_type(event);
                    }
                    // Feature #9: connector start/end connection points.
                    b"stCxn" if current_shape_is_connector => {
                        let shape_id =
                            get_attribute_value(event, b"id").and_then(|v| v.parse().ok());
                        let idx = get_attribute_value(event, b"idx").and_then(|v| v.parse().ok());
                        if let (Some(id), Some(idx)) = (shape_id, idx) {
                            current_shape_start_connection = Some(ConnectionInfo {
                                shape_id: id,
                                connection_point_index: idx,
                            });
                        }
                    }
                    b"endCxn" if current_shape_is_connector => {
                        let shape_id =
                            get_attribute_value(event, b"id").and_then(|v| v.parse().ok());
                        let idx = get_attribute_value(event, b"idx").and_then(|v| v.parse().ok());
                        if let (Some(id), Some(idx)) = (shape_id, idx) {
                            current_shape_end_connection = Some(ConnectionInfo {
                                shape_id: id,
                                connection_point_index: idx,
                            });
                        }
                    }
                    b"ph" => {
                        current_placeholder_kind = get_attribute_value(event, b"type");
                        current_placeholder_idx =
                            get_attribute_value(event, b"idx").and_then(|idx| idx.parse().ok());
                    }
                    b"spPr" => {
                        in_sp_pr = true;
                        sp_pr_depth = 0;
                        in_xfrm = false;
                        in_shape_solid_fill = false;
                        in_outline = false;
                        in_grad_fill = false;
                        in_patt_fill = false;
                    }
                    b"xfrm" if in_sp_pr && sp_pr_depth == 0 => {
                        in_xfrm = true;
                        // Feature #3: rotation.
                        current_shape_rotation =
                            get_attribute_value(event, b"rot").and_then(|v| v.parse::<i32>().ok());
                        // Flip attributes.
                        current_shape_flip_h = get_attribute_value(event, b"flipH")
                            .as_deref()
                            .and_then(parse_xml_bool)
                            .unwrap_or(false);
                        current_shape_flip_v = get_attribute_value(event, b"flipV")
                            .as_deref()
                            .and_then(parse_xml_bool)
                            .unwrap_or(false);
                    }
                    b"off" if in_xfrm => {
                        if let (Some(x), Some(y)) = (
                            parse_i64_attribute_value(event, b"x"),
                            parse_i64_attribute_value(event, b"y"),
                        ) {
                            current_shape_xfrm_offset = Some((x, y));
                        }
                    }
                    b"ext" if in_xfrm => {
                        if let (Some(cx), Some(cy)) = (
                            parse_i64_attribute_value(event, b"cx"),
                            parse_i64_attribute_value(event, b"cy"),
                        ) {
                            current_shape_xfrm_extents = Some((cx, cy));
                        }
                    }
                    b"prstGeom" if in_sp_pr && sp_pr_depth == 0 => {
                        current_shape_preset_geometry = get_attribute_value(event, b"prst");
                        // Capture avLst children as raw XML for roundtrip.
                        current_shape_preset_geometry_adjustments = None;
                    }
                    b"avLst" if in_sp_pr && current_shape_preset_geometry.is_some() => {
                        current_shape_preset_geometry_adjustments =
                            Some(RawXmlNode::read_element(&mut reader, event)?);
                        shape_depth = shape_depth.saturating_sub(1);
                    }
                    b"custGeom" if in_sp_pr && sp_pr_depth == 0 => {
                        current_shape_custom_geometry_raw =
                            Some(RawXmlNode::read_element(&mut reader, event)?);
                        shape_depth = shape_depth.saturating_sub(1);
                    }
                    b"solidFill" if in_sp_pr && sp_pr_depth == 0 && !in_outline => {
                        in_shape_solid_fill = true;
                    }
                    b"solidFill" if in_outline => {
                        in_outline_solid_fill = true;
                    }
                    b"srgbClr" if in_outline_solid_fill => {
                        let val = get_attribute_value(event, b"val");
                        current_shape_outline.color_srgb = val.clone();
                        building_color = val.map(ShapeColor::srgb);
                    }
                    b"schemeClr" if in_outline_solid_fill => {
                        if let Some(val) = get_attribute_value(event, b"val") {
                            building_color = Some(ShapeColor::scheme(val));
                        }
                    }
                    b"srgbClr" if in_shape_solid_fill => {
                        let val = get_attribute_value(event, b"val");
                        current_shape_solid_fill_srgb = val.clone();
                        building_color = val.map(ShapeColor::srgb);
                    }
                    b"schemeClr" if in_shape_solid_fill => {
                        if let Some(val) = get_attribute_value(event, b"val") {
                            building_color = Some(ShapeColor::scheme(val));
                        }
                    }
                    // Transparency/opacity: alpha child of srgbClr/schemeClr.
                    b"alpha" if in_outline_solid_fill => {
                        if let Some(val) =
                            get_attribute_value(event, b"val").and_then(|v| v.parse::<u32>().ok())
                        {
                            // val is in 1/1000ths of percent, so 50000 = 50%.
                            let alpha_pct = (val / 1000) as u8;
                            current_shape_outline.alpha = Some(alpha_pct);
                            if let Some(ref mut bc) = building_color {
                                match bc {
                                    ShapeColor::SrgbClr { alpha, .. }
                                    | ShapeColor::SchemeClr { alpha, .. } => {
                                        *alpha = Some(alpha_pct);
                                    }
                                }
                            }
                        }
                    }
                    b"alpha" if in_shape_solid_fill => {
                        if let Some(val) =
                            get_attribute_value(event, b"val").and_then(|v| v.parse::<u32>().ok())
                        {
                            let alpha_pct = (val / 1000) as u8;
                            current_shape_solid_fill_alpha = Some(alpha_pct);
                            if let Some(ref mut bc) = building_color {
                                match bc {
                                    ShapeColor::SrgbClr { alpha, .. }
                                    | ShapeColor::SchemeClr { alpha, .. } => {
                                        *alpha = Some(alpha_pct);
                                    }
                                }
                            }
                        }
                    }
                    // Color transforms inside srgbClr/schemeClr (lumMod, lumOff, tint, shade, etc.).
                    b"lumMod" | b"lumOff" | b"tint" | b"shade" | b"satMod" | b"satOff"
                    | b"hueOff" | b"hueMod"
                        if building_color.is_some() =>
                    {
                        if let Some(val) =
                            get_attribute_value(event, b"val").and_then(|v| v.parse::<i32>().ok())
                        {
                            let name_bytes = event.name();
                            let transform =
                                ColorTransform::from_xml(local_name(name_bytes.as_ref()), val);
                            if let Some(ref mut bc) = building_color {
                                match bc {
                                    ShapeColor::SrgbClr { transforms, .. }
                                    | ShapeColor::SchemeClr { transforms, .. } => {
                                        transforms.push(transform);
                                    }
                                }
                            }
                        }
                    }
                    // Feature #1: Outline.
                    b"ln" if in_sp_pr && sp_pr_depth == 0 => {
                        in_outline = true;
                        current_shape_outline = ShapeOutline::default();
                        current_shape_outline.width_emu =
                            get_attribute_value(event, b"w").and_then(|v| v.parse().ok());
                        current_shape_outline.compound_style = get_attribute_value(event, b"cmpd")
                            .and_then(|v| LineCompoundStyle::from_xml(&v));
                    }
                    b"prstDash" if in_outline => {
                        current_shape_outline.dash_style = get_attribute_value(event, b"val")
                            .and_then(|v| LineDashStyle::from_xml(&v));
                    }
                    // Line arrows (headEnd/tailEnd inside a:ln).
                    b"headEnd" if in_outline => {
                        let arrow_type = get_attribute_value(event, b"type")
                            .and_then(|v| ArrowType::from_xml(&v))
                            .unwrap_or(ArrowType::None);
                        let width = get_attribute_value(event, b"w")
                            .and_then(|v| ArrowSize::from_xml(&v))
                            .unwrap_or(ArrowSize::Medium);
                        let length = get_attribute_value(event, b"len")
                            .and_then(|v| ArrowSize::from_xml(&v))
                            .unwrap_or(ArrowSize::Medium);
                        current_shape_outline.head_arrow = Some(LineArrow {
                            arrow_type,
                            width,
                            length,
                        });
                    }
                    b"tailEnd" if in_outline => {
                        let arrow_type = get_attribute_value(event, b"type")
                            .and_then(|v| ArrowType::from_xml(&v))
                            .unwrap_or(ArrowType::None);
                        let width = get_attribute_value(event, b"w")
                            .and_then(|v| ArrowSize::from_xml(&v))
                            .unwrap_or(ArrowSize::Medium);
                        let length = get_attribute_value(event, b"len")
                            .and_then(|v| ArrowSize::from_xml(&v))
                            .unwrap_or(ArrowSize::Medium);
                        current_shape_outline.tail_arrow = Some(LineArrow {
                            arrow_type,
                            width,
                            length,
                        });
                    }
                    // Feature #2: Gradient fill.
                    b"gradFill" if in_sp_pr && sp_pr_depth == 0 => {
                        in_grad_fill = true;
                        current_shape_gradient_fill = Some(GradientFill::new());
                    }
                    b"gsLst" if in_grad_fill => {}
                    b"gs" if in_grad_fill => {
                        in_grad_gs = true;
                        grad_gs_pos = get_attribute_value(event, b"pos")
                            .and_then(|v| v.parse().ok())
                            .unwrap_or(0);
                    }
                    b"srgbClr" if in_grad_gs => {
                        if let (Some(ref mut grad), Some(val)) = (
                            &mut current_shape_gradient_fill,
                            get_attribute_value(event, b"val"),
                        ) {
                            building_color = Some(ShapeColor::srgb(&val));
                            grad.stops.push(GradientStop {
                                position: grad_gs_pos,
                                color_srgb: val,
                                color: None,
                            });
                        }
                    }
                    b"schemeClr" if in_grad_gs => {
                        if let Some(val) = get_attribute_value(event, b"val") {
                            building_color = Some(ShapeColor::scheme(&val));
                            if let Some(ref mut grad) = current_shape_gradient_fill {
                                grad.stops.push(GradientStop {
                                    position: grad_gs_pos,
                                    color_srgb: String::new(),
                                    color: None,
                                });
                            }
                        }
                    }
                    b"lin" if in_grad_fill => {
                        if let Some(ref mut grad) = current_shape_gradient_fill {
                            grad.fill_type = Some(GradientFillType::Linear);
                            grad.linear_angle =
                                get_attribute_value(event, b"ang").and_then(|v| v.parse().ok());
                        }
                    }
                    b"path" if in_grad_fill => {
                        if let Some(ref mut grad) = current_shape_gradient_fill {
                            let path_type = get_attribute_value(event, b"path");
                            grad.fill_type = path_type
                                .as_deref()
                                .and_then(GradientFillType::from_xml)
                                .or(Some(GradientFillType::Path));
                        }
                    }
                    // Feature #2: Pattern fill.
                    b"pattFill" if in_sp_pr && sp_pr_depth == 0 => {
                        in_patt_fill = true;
                        let pattern_type = get_attribute_value(event, b"prst")
                            .map(|v| PatternFillType::from_xml(&v))
                            .unwrap_or(PatternFillType::Other("unknown".to_string()));
                        current_shape_pattern_fill = Some(PatternFill::new(pattern_type));
                    }
                    b"fgClr" if in_patt_fill => {
                        in_patt_fg = true;
                    }
                    b"bgClr" if in_patt_fill => {
                        in_patt_bg = true;
                    }
                    b"srgbClr" if in_patt_fg => {
                        if let Some(ref mut pattern) = current_shape_pattern_fill {
                            let val = get_attribute_value(event, b"val");
                            building_color = val.as_deref().map(ShapeColor::srgb);
                            pattern.foreground_srgb = val;
                        }
                    }
                    b"schemeClr" if in_patt_fg => {
                        if let Some(val) = get_attribute_value(event, b"val") {
                            building_color = Some(ShapeColor::scheme(&val));
                        }
                    }
                    b"srgbClr" if in_patt_bg => {
                        if let Some(ref mut pattern) = current_shape_pattern_fill {
                            let val = get_attribute_value(event, b"val");
                            building_color = val.as_deref().map(ShapeColor::srgb);
                            pattern.background_srgb = val;
                        }
                    }
                    b"schemeClr" if in_patt_bg => {
                        if let Some(val) = get_attribute_value(event, b"val") {
                            building_color = Some(ShapeColor::scheme(&val));
                        }
                    }
                    // Feature #2: No fill.
                    b"noFill" if in_sp_pr && sp_pr_depth == 0 => {
                        current_shape_no_fill = true;
                    }
                    // Feature #1 (picture fill): blipFill in spPr.
                    b"blipFill" if in_sp_pr && sp_pr_depth == 0 => {
                        in_blip_fill = true;
                    }
                    b"blip" if in_blip_fill => {
                        if let Some(rid) = get_attribute_value(event, b"r:embed") {
                            current_shape_picture_fill = Some(PictureFill {
                                relationship_id: rid,
                                stretch: false,
                                crop: None,
                            });
                        }
                    }
                    b"srcRect" if in_blip_fill => {
                        // Parse image crop rectangle (l, t, r, b attributes in thousandths of a percent)
                        if let Some(ref mut pf) = current_shape_picture_fill {
                            let l = get_attribute_value(event, b"l")
                                .and_then(|s| s.parse::<i32>().ok())
                                .unwrap_or(0);
                            let t = get_attribute_value(event, b"t")
                                .and_then(|s| s.parse::<i32>().ok())
                                .unwrap_or(0);
                            let r = get_attribute_value(event, b"r")
                                .and_then(|s| s.parse::<i32>().ok())
                                .unwrap_or(0);
                            let b = get_attribute_value(event, b"b")
                                .and_then(|s| s.parse::<i32>().ok())
                                .unwrap_or(0);
                            if l != 0 || t != 0 || r != 0 || b != 0 {
                                pf.crop =
                                    Some(crate::image::ImageCrop::from_pptx_format(l, t, r, b));
                            }
                        }
                    }
                    b"stretch" if in_blip_fill => {
                        if let Some(ref mut pf) = current_shape_picture_fill {
                            pf.stretch = true;
                        }
                    }
                    // Shape effects: effectLst inside spPr.
                    b"effectLst" if in_sp_pr && sp_pr_depth == 0 => {
                        in_effect_lst = true;
                    }
                    b"outerShdw" if in_effect_lst => {
                        in_outer_shdw = true;
                        let dist = get_attribute_value(event, b"dist")
                            .and_then(|v| v.parse::<i64>().ok())
                            .unwrap_or(0);
                        let dir = get_attribute_value(event, b"dir")
                            .and_then(|v| v.parse::<i64>().ok())
                            .unwrap_or(0);
                        let blur = get_attribute_value(event, b"blurRad")
                            .and_then(|v| v.parse::<i64>().ok())
                            .unwrap_or(0);
                        // Convert polar (dist, dir) to cartesian (dx, dy).
                        // dir is in 60000ths of a degree; 0 = right, 5400000 = down.
                        let dir_rad = (dir as f64) / 60000.0_f64 * std::f64::consts::PI / 180.0;
                        let dx = (dist as f64 * dir_rad.cos()) as i64;
                        let dy = (dist as f64 * dir_rad.sin()) as i64;
                        current_shape_shadow = Some(ShapeShadow {
                            offset_x: dx,
                            offset_y: dy,
                            blur_radius: blur,
                            color: String::new(),
                            alpha: None,
                            color_full: None,
                        });
                    }
                    b"srgbClr" if in_outer_shdw => {
                        if let Some(ref mut shadow) = current_shape_shadow {
                            if let Some(val) = get_attribute_value(event, b"val") {
                                building_color = Some(ShapeColor::srgb(&val));
                                shadow.color = val;
                            }
                        }
                    }
                    b"schemeClr" if in_outer_shdw => {
                        if let Some(val) = get_attribute_value(event, b"val") {
                            building_color = Some(ShapeColor::scheme(&val));
                        }
                    }
                    b"alpha" if in_outer_shdw => {
                        if let Some(ref mut shadow) = current_shape_shadow {
                            // Alpha in OOXML is in thousandths of a percent (e.g. 50000 = 50%).
                            let alpha_val = get_attribute_value(event, b"val")
                                .and_then(|v| v.parse::<u32>().ok())
                                .map(|v| (v / 1000) as u8);
                            shadow.alpha = alpha_val;
                            if let Some(alpha_pct) = alpha_val {
                                if let Some(ref mut bc) = building_color {
                                    match bc {
                                        ShapeColor::SrgbClr { alpha, .. }
                                        | ShapeColor::SchemeClr { alpha, .. } => {
                                            *alpha = Some(alpha_pct);
                                        }
                                    }
                                }
                            }
                        }
                    }
                    b"glow" if in_effect_lst => {
                        in_glow = true;
                        let rad = get_attribute_value(event, b"rad")
                            .and_then(|v| v.parse::<i64>().ok())
                            .unwrap_or(0);
                        current_shape_glow = Some(ShapeGlow {
                            radius: rad,
                            color: String::new(),
                            alpha: None,
                            color_full: None,
                        });
                    }
                    b"srgbClr" if in_glow => {
                        if let Some(ref mut glow) = current_shape_glow {
                            if let Some(val) = get_attribute_value(event, b"val") {
                                building_color = Some(ShapeColor::srgb(&val));
                                glow.color = val;
                            }
                        }
                    }
                    b"schemeClr" if in_glow => {
                        if let Some(val) = get_attribute_value(event, b"val") {
                            building_color = Some(ShapeColor::scheme(&val));
                        }
                    }
                    b"alpha" if in_glow => {
                        if let Some(ref mut glow) = current_shape_glow {
                            let alpha_val = get_attribute_value(event, b"val")
                                .and_then(|v| v.parse::<u32>().ok())
                                .map(|v| (v / 1000) as u8);
                            glow.alpha = alpha_val;
                            if let Some(alpha_pct) = alpha_val {
                                if let Some(ref mut bc) = building_color {
                                    match bc {
                                        ShapeColor::SrgbClr { alpha, .. }
                                        | ShapeColor::SchemeClr { alpha, .. } => {
                                            *alpha = Some(alpha_pct);
                                        }
                                    }
                                }
                            }
                        }
                    }
                    b"reflection" if in_effect_lst => {
                        let blur = get_attribute_value(event, b"blurRad")
                            .and_then(|v| v.parse::<i64>().ok())
                            .unwrap_or(0);
                        let dist = get_attribute_value(event, b"dist")
                            .and_then(|v| v.parse::<i64>().ok())
                            .unwrap_or(0);
                        let dir =
                            get_attribute_value(event, b"dir").and_then(|v| v.parse::<i64>().ok());
                        let start_a = get_attribute_value(event, b"stA")
                            .and_then(|v| v.parse::<u32>().ok())
                            .map(|v| (v / 1000) as u8);
                        let end_a = get_attribute_value(event, b"endA")
                            .and_then(|v| v.parse::<u32>().ok())
                            .map(|v| (v / 1000) as u8);
                        current_shape_reflection = Some(ShapeReflection {
                            blur_radius: blur,
                            start_alpha: start_a,
                            end_alpha: end_a,
                            distance: dist,
                            direction: dir,
                        });
                    }
                    // Feature #7: SmartArt detection (dgm:relIds inside graphicData).
                    b"relIds" if in_shape => {
                        // The presence of dgm:relIds indicates SmartArt.
                        current_shape_is_smartart = true;
                    }
                    // Feature #8: Audio/video media.
                    b"audioFile" if in_shape => {
                        if let Some(rid) = get_attribute_value(event, b"r:link") {
                            current_shape_media = Some((MediaType::Audio, rid));
                        }
                    }
                    b"videoFile" if in_shape => {
                        if let Some(rid) = get_attribute_value(event, b"r:link") {
                            current_shape_media = Some((MediaType::Video, rid));
                        }
                    }
                    b"txBody" => {
                        in_tx_body = true;
                    }
                    b"bodyPr" if in_tx_body => {
                        in_body_pr = true;
                        // Parse text anchor from anchor/anchorCtr attributes.
                        let anchor_val = get_attribute_value(event, b"anchor");
                        let anchor_ctr = get_attribute_value(event, b"anchorCtr")
                            .as_deref()
                            .and_then(parse_xml_bool)
                            .unwrap_or(false);
                        if let Some(ref anchor) = anchor_val {
                            current_shape_text_anchor = TextAnchor::from_xml(anchor, anchor_ctr);
                        }
                        // Text direction.
                        current_shape_text_direction = get_attribute_value(event, b"vert")
                            .and_then(|v| TextDirection::from_xml(&v));
                        // Text columns.
                        current_shape_text_columns =
                            get_attribute_value(event, b"numCol").and_then(|v| v.parse().ok());
                        current_shape_text_column_spacing =
                            get_attribute_value(event, b"spcCol").and_then(|v| v.parse().ok());
                        // Text insets/margins.
                        current_shape_text_inset_left =
                            get_attribute_value(event, b"lIns").and_then(|v| v.parse().ok());
                        current_shape_text_inset_right =
                            get_attribute_value(event, b"rIns").and_then(|v| v.parse().ok());
                        current_shape_text_inset_top =
                            get_attribute_value(event, b"tIns").and_then(|v| v.parse().ok());
                        current_shape_text_inset_bottom =
                            get_attribute_value(event, b"bIns").and_then(|v| v.parse().ok());
                        // Word wrap.
                        current_shape_word_wrap =
                            get_attribute_value(event, b"wrap").map(|v| v == "square");
                        // Extended bodyPr attributes.
                        current_shape_body_pr_rot =
                            get_attribute_value(event, b"rot").and_then(|v| v.parse().ok());
                        current_shape_body_pr_rtl_col = get_attribute_value(event, b"rtlCol")
                            .as_deref()
                            .and_then(parse_xml_bool);
                        current_shape_body_pr_from_word_art =
                            get_attribute_value(event, b"fromWordArt")
                                .as_deref()
                                .and_then(parse_xml_bool);
                        current_shape_body_pr_force_aa = get_attribute_value(event, b"forceAA")
                            .as_deref()
                            .and_then(parse_xml_bool);
                        current_shape_body_pr_compat_ln_spc =
                            get_attribute_value(event, b"compatLnSpc")
                                .as_deref()
                                .and_then(parse_xml_bool);
                        // Capture unknown bodyPr attributes for roundtrip.
                        let known_attrs: &[&[u8]] = &[
                            b"anchor",
                            b"anchorCtr",
                            b"vert",
                            b"numCol",
                            b"spcCol",
                            b"lIns",
                            b"rIns",
                            b"tIns",
                            b"bIns",
                            b"wrap",
                            b"rot",
                            b"rtlCol",
                            b"fromWordArt",
                            b"forceAA",
                            b"compatLnSpc",
                        ];
                        for attr in event.attributes().flatten() {
                            let key = attr.key.local_name();
                            if !known_attrs.iter().any(|k| key.as_ref() == *k) {
                                let k = String::from_utf8_lossy(key.as_ref()).into_owned();
                                let v = String::from_utf8_lossy(&attr.value).into_owned();
                                current_body_pr_unknown_attrs.push((k, v));
                            }
                        }
                    }
                    // Auto-fit elements inside bodyPr.
                    b"noAutofit" if in_body_pr => {
                        current_shape_auto_fit = Some(AutoFitType::None);
                    }
                    b"normAutofit" if in_body_pr => {
                        current_shape_auto_fit = Some(AutoFitType::Normal);
                    }
                    b"spAutoFit" if in_body_pr => {
                        current_shape_auto_fit = Some(AutoFitType::ShrinkOnOverflow);
                    }
                    // Capture lstStyle as a raw node for roundtrip fidelity.
                    b"lstStyle" if in_tx_body => {
                        current_lst_style_raw = Some(RawXmlNode::read_element(&mut reader, event)?);
                        shape_depth = shape_depth.saturating_sub(1);
                    }
                    b"p" if in_tx_body => {
                        current_paragraph_runs = Some(Vec::new());
                        current_paragraph_properties = ParagraphProperties::default();
                    }
                    b"pPr" if current_paragraph_runs.is_some() && !in_rpr => {
                        in_ppr = true;
                        current_paragraph_properties.alignment =
                            get_attribute_value(event, b"algn")
                                .and_then(|v| TextAlignment::from_xml(&v));
                        current_paragraph_properties.level =
                            get_attribute_value(event, b"lvl").and_then(|v| v.parse().ok());
                        current_paragraph_properties.margin_left_emu =
                            get_attribute_value(event, b"marL").and_then(|v| v.parse().ok());
                        current_paragraph_properties.margin_right_emu =
                            get_attribute_value(event, b"marR").and_then(|v| v.parse().ok());
                        current_paragraph_properties.indent_emu =
                            get_attribute_value(event, b"indent").and_then(|v| v.parse().ok());
                    }
                    // Feature #11: Bullets.
                    b"buNone" if in_ppr => {
                        current_paragraph_properties.bullet.style = Some(BulletStyle::None);
                    }
                    b"buChar" if in_ppr => {
                        if let Some(ch) = get_attribute_value(event, b"char") {
                            current_paragraph_properties.bullet.style = Some(BulletStyle::Char(ch));
                        }
                    }
                    b"buAutoNum" if in_ppr => {
                        if let Some(t) = get_attribute_value(event, b"type") {
                            current_paragraph_properties.bullet.style =
                                Some(BulletStyle::AutoNum(t));
                        }
                    }
                    b"buFont" if in_ppr => {
                        current_paragraph_properties.bullet.font_name =
                            get_attribute_value(event, b"typeface");
                    }
                    b"buSzPct" if in_ppr => {
                        current_paragraph_properties.bullet.size_percent =
                            get_attribute_value(event, b"val").and_then(|v| v.parse().ok());
                    }
                    b"buClr" if in_ppr => {
                        in_bu_clr = true;
                    }
                    b"srgbClr" if in_bu_clr => {
                        let val = get_attribute_value(event, b"val");
                        building_color = val.as_deref().map(ShapeColor::srgb);
                        current_paragraph_properties.bullet.color_srgb = val;
                    }
                    b"schemeClr" if in_bu_clr => {
                        if let Some(val) = get_attribute_value(event, b"val") {
                            building_color = Some(ShapeColor::scheme(&val));
                        }
                    }
                    b"lnSpc" if current_paragraph_runs.is_some() => {
                        in_spacing_parent = Some("lnSpc");
                    }
                    b"spcBef" if current_paragraph_runs.is_some() => {
                        in_spacing_parent = Some("spcBef");
                    }
                    b"spcAft" if current_paragraph_runs.is_some() => {
                        in_spacing_parent = Some("spcAft");
                    }
                    b"rPr" if current_paragraph_runs.is_some() => {
                        in_rpr = true;
                        in_rpr_solid_fill = false;
                        current_run_properties = RunProperties::default();
                        current_run_properties.bold = get_attribute_value(event, b"b")
                            .as_deref()
                            .and_then(parse_xml_bool);
                        current_run_properties.italic = get_attribute_value(event, b"i")
                            .as_deref()
                            .and_then(parse_xml_bool);
                        current_run_properties.underline = get_attribute_value(event, b"u")
                            .and_then(|v| crate::text::UnderlineStyle::from_xml(&v));
                        current_run_properties.strikethrough =
                            get_attribute_value(event, b"strike")
                                .and_then(|v| crate::text::StrikethroughStyle::from_xml(&v));
                        current_run_properties.font_size =
                            get_attribute_value(event, b"sz").and_then(|v| v.parse().ok());
                        current_run_properties.language = get_attribute_value(event, b"lang");
                        // Character spacing.
                        current_run_properties.character_spacing =
                            get_attribute_value(event, b"spc").and_then(|v| v.parse().ok());
                        // Kerning.
                        current_run_properties.kerning =
                            get_attribute_value(event, b"kern").and_then(|v| v.parse().ok());
                        // Subscript/superscript baseline.
                        current_run_properties.baseline =
                            get_attribute_value(event, b"baseline").and_then(|v| v.parse().ok());
                    }
                    b"solidFill" if in_rpr => {
                        in_rpr_solid_fill = true;
                    }
                    b"srgbClr" if in_rpr_solid_fill => {
                        let val = get_attribute_value(event, b"val");
                        building_color = val.as_deref().map(ShapeColor::srgb);
                        current_run_properties.font_color_srgb = val;
                    }
                    b"schemeClr" if in_rpr_solid_fill => {
                        if let Some(val) = get_attribute_value(event, b"val") {
                            building_color = Some(ShapeColor::scheme(&val));
                        }
                    }
                    b"latin" if in_rpr => {
                        current_run_properties.font_name = get_attribute_value(event, b"typeface");
                    }
                    b"ea" if in_rpr => {
                        current_run_properties.font_name_east_asian =
                            get_attribute_value(event, b"typeface");
                    }
                    b"cs" if in_rpr => {
                        current_run_properties.font_name_complex_script =
                            get_attribute_value(event, b"typeface");
                    }
                    // Feature #10: Hyperlinks in run properties.
                    b"hlinkClick" if in_rpr => {
                        current_run_properties.hyperlink_click_rid =
                            get_exact_attribute_value(event, b"r:id");
                        current_run_properties.hyperlink_tooltip =
                            get_attribute_value(event, b"tooltip");
                    }
                    b"t" if current_paragraph_runs.is_some() => {
                        in_text = true;
                        current_text.clear();
                    }
                    _ => {
                        if in_body_pr
                            && !matches!(local, b"noAutofit" | b"normAutofit" | b"spAutoFit")
                        {
                            // Unknown bodyPr child — capture for roundtrip.
                            current_body_pr_unknown_children
                                .push(RawXmlNode::read_element(&mut reader, event)?);
                            shape_depth = shape_depth.saturating_sub(1);
                        } else if shape_depth == 2
                            && !matches!(local, b"nvSpPr" | b"spPr" | b"txBody")
                        {
                            current_shape_unknown_children
                                .push(RawXmlNode::read_element(&mut reader, event)?);
                            shape_depth = shape_depth.saturating_sub(1);
                        }
                    }
                }
                if in_sp_pr && local != b"spPr" {
                    sp_pr_depth = sp_pr_depth.saturating_add(1);
                }
            }
            Event::Empty(ref event) => {
                if !in_shape {
                    buffer.clear();
                    continue;
                }

                let event_name = event.name();
                let local = local_name(event_name.as_ref());
                match local {
                    b"cNvPr" => {
                        if let Some(name) = get_attribute_value(event, b"name") {
                            current_shape_name = name;
                        }
                        if get_attribute_value(event, b"hidden")
                            .as_deref()
                            .and_then(parse_xml_bool)
                            .unwrap_or(false)
                        {
                            current_shape_hidden = true;
                        }
                        // Feature #10: alt text.
                        current_shape_alt_text = get_attribute_value(event, b"descr");
                        current_shape_alt_text_title = get_attribute_value(event, b"title");
                    }
                    b"cNvSpPr" | b"cNvCxnSpPr" => {
                        current_shape_type = parse_shape_type(event);
                    }
                    // Feature #9: connector start/end connection points (self-closing).
                    b"stCxn" if current_shape_is_connector => {
                        let shape_id =
                            get_attribute_value(event, b"id").and_then(|v| v.parse().ok());
                        let idx = get_attribute_value(event, b"idx").and_then(|v| v.parse().ok());
                        if let (Some(id), Some(idx)) = (shape_id, idx) {
                            current_shape_start_connection = Some(ConnectionInfo {
                                shape_id: id,
                                connection_point_index: idx,
                            });
                        }
                    }
                    b"endCxn" if current_shape_is_connector => {
                        let shape_id =
                            get_attribute_value(event, b"id").and_then(|v| v.parse().ok());
                        let idx = get_attribute_value(event, b"idx").and_then(|v| v.parse().ok());
                        if let (Some(id), Some(idx)) = (shape_id, idx) {
                            current_shape_end_connection = Some(ConnectionInfo {
                                shape_id: id,
                                connection_point_index: idx,
                            });
                        }
                    }
                    b"ph" => {
                        current_placeholder_kind = get_attribute_value(event, b"type");
                        current_placeholder_idx =
                            get_attribute_value(event, b"idx").and_then(|idx| idx.parse().ok());
                    }
                    // Feature #1: self-closing blip in blipFill.
                    b"blip" if in_blip_fill => {
                        if let Some(rid) = get_attribute_value(event, b"r:embed") {
                            current_shape_picture_fill = Some(PictureFill {
                                relationship_id: rid,
                                stretch: false,
                                crop: None,
                            });
                        }
                    }
                    b"srcRect" if in_blip_fill => {
                        // Parse image crop rectangle (l, t, r, b attributes in thousandths of a percent)
                        if let Some(ref mut pf) = current_shape_picture_fill {
                            let l = get_attribute_value(event, b"l")
                                .and_then(|s| s.parse::<i32>().ok())
                                .unwrap_or(0);
                            let t = get_attribute_value(event, b"t")
                                .and_then(|s| s.parse::<i32>().ok())
                                .unwrap_or(0);
                            let r = get_attribute_value(event, b"r")
                                .and_then(|s| s.parse::<i32>().ok())
                                .unwrap_or(0);
                            let b = get_attribute_value(event, b"b")
                                .and_then(|s| s.parse::<i32>().ok())
                                .unwrap_or(0);
                            if l != 0 || t != 0 || r != 0 || b != 0 {
                                pf.crop =
                                    Some(crate::image::ImageCrop::from_pptx_format(l, t, r, b));
                            }
                        }
                    }
                    b"stretch" if in_blip_fill => {
                        if let Some(ref mut pf) = current_shape_picture_fill {
                            pf.stretch = true;
                        }
                    }
                    // Feature #7: SmartArt detection (self-closing dgm:relIds).
                    b"relIds" if in_shape => {
                        current_shape_is_smartart = true;
                    }
                    // Shape effects (self-closing) in Empty handler.
                    b"effectLst" if in_sp_pr && sp_pr_depth == 0 => {}
                    b"outerShdw" if in_effect_lst => {
                        let dist = get_attribute_value(event, b"dist")
                            .and_then(|v| v.parse::<i64>().ok())
                            .unwrap_or(0);
                        let dir = get_attribute_value(event, b"dir")
                            .and_then(|v| v.parse::<i64>().ok())
                            .unwrap_or(0);
                        let blur = get_attribute_value(event, b"blurRad")
                            .and_then(|v| v.parse::<i64>().ok())
                            .unwrap_or(0);
                        let dir_rad = (dir as f64) / 60000.0_f64 * std::f64::consts::PI / 180.0;
                        current_shape_shadow = Some(ShapeShadow {
                            offset_x: (dist as f64 * dir_rad.cos()) as i64,
                            offset_y: (dist as f64 * dir_rad.sin()) as i64,
                            blur_radius: blur,
                            color: String::new(),
                            alpha: None,
                            color_full: None,
                        });
                    }
                    b"srgbClr" if in_outer_shdw => {
                        if let Some(ref mut shadow) = current_shape_shadow {
                            if let Some(val) = get_attribute_value(event, b"val") {
                                building_color = Some(ShapeColor::srgb(&val));
                                shadow.color = val;
                            }
                        }
                    }
                    b"schemeClr" if in_outer_shdw => {
                        if let Some(val) = get_attribute_value(event, b"val") {
                            building_color = Some(ShapeColor::scheme(&val));
                        }
                    }
                    b"alpha" if in_outer_shdw => {
                        if let Some(ref mut shadow) = current_shape_shadow {
                            let alpha_val = get_attribute_value(event, b"val")
                                .and_then(|v| v.parse::<u32>().ok())
                                .map(|v| (v / 1000) as u8);
                            shadow.alpha = alpha_val;
                            if let Some(alpha_pct) = alpha_val {
                                if let Some(ref mut bc) = building_color {
                                    match bc {
                                        ShapeColor::SrgbClr { alpha, .. }
                                        | ShapeColor::SchemeClr { alpha, .. } => {
                                            *alpha = Some(alpha_pct);
                                        }
                                    }
                                }
                            }
                        }
                    }
                    b"glow" if in_effect_lst => {
                        let rad = get_attribute_value(event, b"rad")
                            .and_then(|v| v.parse::<i64>().ok())
                            .unwrap_or(0);
                        current_shape_glow = Some(ShapeGlow {
                            radius: rad,
                            color: String::new(),
                            alpha: None,
                            color_full: None,
                        });
                    }
                    b"srgbClr" if in_glow => {
                        if let Some(ref mut glow) = current_shape_glow {
                            if let Some(val) = get_attribute_value(event, b"val") {
                                building_color = Some(ShapeColor::srgb(&val));
                                glow.color = val;
                            }
                        }
                    }
                    b"schemeClr" if in_glow => {
                        if let Some(val) = get_attribute_value(event, b"val") {
                            building_color = Some(ShapeColor::scheme(&val));
                        }
                    }
                    b"alpha" if in_glow => {
                        if let Some(ref mut glow) = current_shape_glow {
                            let alpha_val = get_attribute_value(event, b"val")
                                .and_then(|v| v.parse::<u32>().ok())
                                .map(|v| (v / 1000) as u8);
                            glow.alpha = alpha_val;
                            if let Some(alpha_pct) = alpha_val {
                                if let Some(ref mut bc) = building_color {
                                    match bc {
                                        ShapeColor::SrgbClr { alpha, .. }
                                        | ShapeColor::SchemeClr { alpha, .. } => {
                                            *alpha = Some(alpha_pct);
                                        }
                                    }
                                }
                            }
                        }
                    }
                    b"reflection" if in_effect_lst => {
                        let blur = get_attribute_value(event, b"blurRad")
                            .and_then(|v| v.parse::<i64>().ok())
                            .unwrap_or(0);
                        let dist = get_attribute_value(event, b"dist")
                            .and_then(|v| v.parse::<i64>().ok())
                            .unwrap_or(0);
                        let dir =
                            get_attribute_value(event, b"dir").and_then(|v| v.parse::<i64>().ok());
                        let start_a = get_attribute_value(event, b"stA")
                            .and_then(|v| v.parse::<u32>().ok())
                            .map(|v| (v / 1000) as u8);
                        let end_a = get_attribute_value(event, b"endA")
                            .and_then(|v| v.parse::<u32>().ok())
                            .map(|v| (v / 1000) as u8);
                        current_shape_reflection = Some(ShapeReflection {
                            blur_radius: blur,
                            start_alpha: start_a,
                            end_alpha: end_a,
                            distance: dist,
                            direction: dir,
                        });
                    }
                    b"bodyPr" if in_tx_body => {
                        let anchor_val = get_attribute_value(event, b"anchor");
                        let anchor_ctr = get_attribute_value(event, b"anchorCtr")
                            .as_deref()
                            .and_then(parse_xml_bool)
                            .unwrap_or(false);
                        if let Some(ref anchor) = anchor_val {
                            current_shape_text_anchor = TextAnchor::from_xml(anchor, anchor_ctr);
                        }
                        // Text direction.
                        current_shape_text_direction = get_attribute_value(event, b"vert")
                            .and_then(|v| TextDirection::from_xml(&v));
                        // Text columns.
                        current_shape_text_columns =
                            get_attribute_value(event, b"numCol").and_then(|v| v.parse().ok());
                        current_shape_text_column_spacing =
                            get_attribute_value(event, b"spcCol").and_then(|v| v.parse().ok());
                        // Text insets/margins.
                        current_shape_text_inset_left =
                            get_attribute_value(event, b"lIns").and_then(|v| v.parse().ok());
                        current_shape_text_inset_right =
                            get_attribute_value(event, b"rIns").and_then(|v| v.parse().ok());
                        current_shape_text_inset_top =
                            get_attribute_value(event, b"tIns").and_then(|v| v.parse().ok());
                        current_shape_text_inset_bottom =
                            get_attribute_value(event, b"bIns").and_then(|v| v.parse().ok());
                        // Word wrap.
                        current_shape_word_wrap =
                            get_attribute_value(event, b"wrap").map(|v| v == "square");
                        // Extended bodyPr attributes.
                        current_shape_body_pr_rot =
                            get_attribute_value(event, b"rot").and_then(|v| v.parse().ok());
                        current_shape_body_pr_rtl_col = get_attribute_value(event, b"rtlCol")
                            .as_deref()
                            .and_then(parse_xml_bool);
                        current_shape_body_pr_from_word_art =
                            get_attribute_value(event, b"fromWordArt")
                                .as_deref()
                                .and_then(parse_xml_bool);
                        current_shape_body_pr_force_aa = get_attribute_value(event, b"forceAA")
                            .as_deref()
                            .and_then(parse_xml_bool);
                        current_shape_body_pr_compat_ln_spc =
                            get_attribute_value(event, b"compatLnSpc")
                                .as_deref()
                                .and_then(parse_xml_bool);
                        // Capture unknown bodyPr attributes for roundtrip.
                        let known_attrs: &[&[u8]] = &[
                            b"anchor",
                            b"anchorCtr",
                            b"vert",
                            b"numCol",
                            b"spcCol",
                            b"lIns",
                            b"rIns",
                            b"tIns",
                            b"bIns",
                            b"wrap",
                            b"rot",
                            b"rtlCol",
                            b"fromWordArt",
                            b"forceAA",
                            b"compatLnSpc",
                        ];
                        for attr in event.attributes().flatten() {
                            let key = attr.key.local_name();
                            if !known_attrs.iter().any(|k| key.as_ref() == *k) {
                                let k = String::from_utf8_lossy(key.as_ref()).into_owned();
                                let v = String::from_utf8_lossy(&attr.value).into_owned();
                                current_body_pr_unknown_attrs.push((k, v));
                            }
                        }
                    }
                    b"noAutofit" if in_body_pr => {
                        current_shape_auto_fit = Some(AutoFitType::None);
                    }
                    b"normAutofit" if in_body_pr => {
                        current_shape_auto_fit = Some(AutoFitType::Normal);
                    }
                    b"spAutoFit" if in_body_pr => {
                        current_shape_auto_fit = Some(AutoFitType::ShrinkOnOverflow);
                    }
                    // Feature #8: Audio/video media (self-closing).
                    b"audioFile" if in_shape => {
                        if let Some(rid) = get_attribute_value(event, b"r:link") {
                            current_shape_media = Some((MediaType::Audio, rid));
                        }
                    }
                    b"videoFile" if in_shape => {
                        if let Some(rid) = get_attribute_value(event, b"r:link") {
                            current_shape_media = Some((MediaType::Video, rid));
                        }
                    }
                    b"off" if in_xfrm => {
                        if let (Some(x), Some(y)) = (
                            parse_i64_attribute_value(event, b"x"),
                            parse_i64_attribute_value(event, b"y"),
                        ) {
                            current_shape_xfrm_offset = Some((x, y));
                        }
                    }
                    b"ext" if in_xfrm => {
                        if let (Some(cx), Some(cy)) = (
                            parse_i64_attribute_value(event, b"cx"),
                            parse_i64_attribute_value(event, b"cy"),
                        ) {
                            current_shape_xfrm_extents = Some((cx, cy));
                        }
                    }
                    b"prstGeom" if in_sp_pr && sp_pr_depth == 0 => {
                        current_shape_preset_geometry = get_attribute_value(event, b"prst");
                    }
                    b"srgbClr" if in_outline_solid_fill => {
                        let val = get_attribute_value(event, b"val");
                        building_color = val.as_deref().map(ShapeColor::srgb);
                        current_shape_outline.color_srgb = val;
                    }
                    b"schemeClr" if in_outline_solid_fill => {
                        if let Some(val) = get_attribute_value(event, b"val") {
                            building_color = Some(ShapeColor::scheme(&val));
                        }
                    }
                    b"srgbClr" if in_shape_solid_fill => {
                        let val = get_attribute_value(event, b"val");
                        building_color = val.as_deref().map(ShapeColor::srgb);
                        current_shape_solid_fill_srgb = val;
                    }
                    b"schemeClr" if in_shape_solid_fill => {
                        if let Some(val) = get_attribute_value(event, b"val") {
                            building_color = Some(ShapeColor::scheme(&val));
                        }
                    }
                    b"srgbClr" if in_rpr_solid_fill => {
                        let val = get_attribute_value(event, b"val");
                        building_color = val.as_deref().map(ShapeColor::srgb);
                        current_run_properties.font_color_srgb = val;
                    }
                    b"schemeClr" if in_rpr_solid_fill => {
                        if let Some(val) = get_attribute_value(event, b"val") {
                            building_color = Some(ShapeColor::scheme(&val));
                        }
                    }
                    b"srgbClr" if in_grad_gs => {
                        if let (Some(ref mut grad), Some(val)) = (
                            &mut current_shape_gradient_fill,
                            get_attribute_value(event, b"val"),
                        ) {
                            building_color = Some(ShapeColor::srgb(&val));
                            grad.stops.push(GradientStop {
                                position: grad_gs_pos,
                                color_srgb: val,
                                color: None,
                            });
                        }
                    }
                    b"schemeClr" if in_grad_gs => {
                        if let Some(val) = get_attribute_value(event, b"val") {
                            building_color = Some(ShapeColor::scheme(&val));
                            if let Some(ref mut grad) = current_shape_gradient_fill {
                                grad.stops.push(GradientStop {
                                    position: grad_gs_pos,
                                    color_srgb: String::new(),
                                    color: None,
                                });
                            }
                        }
                    }
                    b"srgbClr" if in_patt_fg => {
                        if let Some(ref mut pattern) = current_shape_pattern_fill {
                            let val = get_attribute_value(event, b"val");
                            building_color = val.as_deref().map(ShapeColor::srgb);
                            pattern.foreground_srgb = val;
                        }
                    }
                    b"schemeClr" if in_patt_fg => {
                        if let Some(val) = get_attribute_value(event, b"val") {
                            building_color = Some(ShapeColor::scheme(&val));
                        }
                    }
                    b"srgbClr" if in_patt_bg => {
                        if let Some(ref mut pattern) = current_shape_pattern_fill {
                            let val = get_attribute_value(event, b"val");
                            building_color = val.as_deref().map(ShapeColor::srgb);
                            pattern.background_srgb = val;
                        }
                    }
                    b"schemeClr" if in_patt_bg => {
                        if let Some(val) = get_attribute_value(event, b"val") {
                            building_color = Some(ShapeColor::scheme(&val));
                        }
                    }
                    b"srgbClr" if in_bu_clr => {
                        let val = get_attribute_value(event, b"val");
                        building_color = val.as_deref().map(ShapeColor::srgb);
                        current_paragraph_properties.bullet.color_srgb = val;
                    }
                    b"schemeClr" if in_bu_clr => {
                        if let Some(val) = get_attribute_value(event, b"val") {
                            building_color = Some(ShapeColor::scheme(&val));
                        }
                    }
                    b"noFill" if in_sp_pr && sp_pr_depth == 0 => {
                        current_shape_no_fill = true;
                    }
                    b"ln" if in_sp_pr && sp_pr_depth == 0 => {
                        // Self-closing <a:ln/>.
                        let outline = ShapeOutline {
                            width_emu: get_attribute_value(event, b"w")
                                .and_then(|v| v.parse().ok()),
                            compound_style: get_attribute_value(event, b"cmpd")
                                .and_then(|v| LineCompoundStyle::from_xml(&v)),
                            ..Default::default()
                        };
                        if outline.is_set() {
                            current_shape_outline = outline;
                        }
                    }
                    b"prstDash" if in_outline => {
                        current_shape_outline.dash_style = get_attribute_value(event, b"val")
                            .and_then(|v| LineDashStyle::from_xml(&v));
                    }
                    b"lin" if in_grad_fill => {
                        if let Some(ref mut grad) = current_shape_gradient_fill {
                            grad.fill_type = Some(GradientFillType::Linear);
                            grad.linear_angle =
                                get_attribute_value(event, b"ang").and_then(|v| v.parse().ok());
                        }
                    }
                    // Feature #11: Bullets in empty elements.
                    b"buNone" if in_ppr => {
                        current_paragraph_properties.bullet.style = Some(BulletStyle::None);
                    }
                    b"buChar" if in_ppr => {
                        if let Some(ch) = get_attribute_value(event, b"char") {
                            current_paragraph_properties.bullet.style = Some(BulletStyle::Char(ch));
                        }
                    }
                    b"buAutoNum" if in_ppr => {
                        if let Some(t) = get_attribute_value(event, b"type") {
                            current_paragraph_properties.bullet.style =
                                Some(BulletStyle::AutoNum(t));
                        }
                    }
                    b"buFont" if in_ppr => {
                        current_paragraph_properties.bullet.font_name =
                            get_attribute_value(event, b"typeface");
                    }
                    b"buSzPct" if in_ppr => {
                        current_paragraph_properties.bullet.size_percent =
                            get_attribute_value(event, b"val").and_then(|v| v.parse().ok());
                    }
                    // Feature #10: Hyperlinks.
                    b"hlinkClick" if in_rpr => {
                        current_run_properties.hyperlink_click_rid =
                            get_exact_attribute_value(event, b"r:id");
                        current_run_properties.hyperlink_tooltip =
                            get_attribute_value(event, b"tooltip");
                    }
                    // Transparency: alpha inside srgbClr/schemeClr (Empty).
                    b"alpha" if in_outline_solid_fill => {
                        if let Some(val) =
                            get_attribute_value(event, b"val").and_then(|v| v.parse::<u32>().ok())
                        {
                            let alpha_pct = (val / 1000) as u8;
                            current_shape_outline.alpha = Some(alpha_pct);
                            if let Some(ref mut bc) = building_color {
                                match bc {
                                    ShapeColor::SrgbClr { alpha, .. }
                                    | ShapeColor::SchemeClr { alpha, .. } => {
                                        *alpha = Some(alpha_pct);
                                    }
                                }
                            }
                        }
                    }
                    b"alpha" if in_shape_solid_fill => {
                        if let Some(val) =
                            get_attribute_value(event, b"val").and_then(|v| v.parse::<u32>().ok())
                        {
                            let alpha_pct = (val / 1000) as u8;
                            current_shape_solid_fill_alpha = Some(alpha_pct);
                            if let Some(ref mut bc) = building_color {
                                match bc {
                                    ShapeColor::SrgbClr { alpha, .. }
                                    | ShapeColor::SchemeClr { alpha, .. } => {
                                        *alpha = Some(alpha_pct);
                                    }
                                }
                            }
                        }
                    }
                    // Color transforms (Empty handler).
                    b"lumMod" | b"lumOff" | b"tint" | b"shade" | b"satMod" | b"satOff"
                    | b"hueOff" | b"hueMod"
                        if building_color.is_some() =>
                    {
                        if let Some(val) =
                            get_attribute_value(event, b"val").and_then(|v| v.parse::<i32>().ok())
                        {
                            let name_bytes = event.name();
                            let transform =
                                ColorTransform::from_xml(local_name(name_bytes.as_ref()), val);
                            if let Some(ref mut bc) = building_color {
                                match bc {
                                    ShapeColor::SrgbClr { transforms, .. }
                                    | ShapeColor::SchemeClr { transforms, .. } => {
                                        transforms.push(transform);
                                    }
                                }
                            }
                        }
                    }
                    // Line arrows (self-closing headEnd/tailEnd).
                    b"headEnd" if in_outline => {
                        let arrow_type = get_attribute_value(event, b"type")
                            .and_then(|v| ArrowType::from_xml(&v))
                            .unwrap_or(ArrowType::None);
                        let width = get_attribute_value(event, b"w")
                            .and_then(|v| ArrowSize::from_xml(&v))
                            .unwrap_or(ArrowSize::Medium);
                        let length = get_attribute_value(event, b"len")
                            .and_then(|v| ArrowSize::from_xml(&v))
                            .unwrap_or(ArrowSize::Medium);
                        current_shape_outline.head_arrow = Some(LineArrow {
                            arrow_type,
                            width,
                            length,
                        });
                    }
                    b"tailEnd" if in_outline => {
                        let arrow_type = get_attribute_value(event, b"type")
                            .and_then(|v| ArrowType::from_xml(&v))
                            .unwrap_or(ArrowType::None);
                        let width = get_attribute_value(event, b"w")
                            .and_then(|v| ArrowSize::from_xml(&v))
                            .unwrap_or(ArrowSize::Medium);
                        let length = get_attribute_value(event, b"len")
                            .and_then(|v| ArrowSize::from_xml(&v))
                            .unwrap_or(ArrowSize::Medium);
                        current_shape_outline.tail_arrow = Some(LineArrow {
                            arrow_type,
                            width,
                            length,
                        });
                    }
                    b"rPr" if current_paragraph_runs.is_some() => {
                        // Self-closing <a:rPr .../> with attributes only.
                        current_run_properties = RunProperties::default();
                        current_run_properties.bold = get_attribute_value(event, b"b")
                            .as_deref()
                            .and_then(parse_xml_bool);
                        current_run_properties.italic = get_attribute_value(event, b"i")
                            .as_deref()
                            .and_then(parse_xml_bool);
                        current_run_properties.underline = get_attribute_value(event, b"u")
                            .and_then(|v| crate::text::UnderlineStyle::from_xml(&v));
                        current_run_properties.strikethrough =
                            get_attribute_value(event, b"strike")
                                .and_then(|v| crate::text::StrikethroughStyle::from_xml(&v));
                        current_run_properties.font_size =
                            get_attribute_value(event, b"sz").and_then(|v| v.parse().ok());
                        current_run_properties.language = get_attribute_value(event, b"lang");
                        // Character spacing.
                        current_run_properties.character_spacing =
                            get_attribute_value(event, b"spc").and_then(|v| v.parse().ok());
                        // Kerning.
                        current_run_properties.kerning =
                            get_attribute_value(event, b"kern").and_then(|v| v.parse().ok());
                        // Subscript/superscript baseline.
                        current_run_properties.baseline =
                            get_attribute_value(event, b"baseline").and_then(|v| v.parse().ok());
                    }
                    b"pPr" if current_paragraph_runs.is_some() => {
                        // Self-closing <a:pPr .../>.
                        current_paragraph_properties.alignment =
                            get_attribute_value(event, b"algn")
                                .and_then(|v| TextAlignment::from_xml(&v));
                        current_paragraph_properties.level =
                            get_attribute_value(event, b"lvl").and_then(|v| v.parse().ok());
                        current_paragraph_properties.margin_left_emu =
                            get_attribute_value(event, b"marL").and_then(|v| v.parse().ok());
                        current_paragraph_properties.margin_right_emu =
                            get_attribute_value(event, b"marR").and_then(|v| v.parse().ok());
                        current_paragraph_properties.indent_emu =
                            get_attribute_value(event, b"indent").and_then(|v| v.parse().ok());
                        // No children for self-closing pPr, so no bullets parsed here.
                    }
                    b"latin" if in_rpr => {
                        current_run_properties.font_name = get_attribute_value(event, b"typeface");
                    }
                    b"ea" if in_rpr => {
                        current_run_properties.font_name_east_asian =
                            get_attribute_value(event, b"typeface");
                    }
                    b"cs" if in_rpr => {
                        current_run_properties.font_name_complex_script =
                            get_attribute_value(event, b"typeface");
                    }
                    b"spcPct" if current_paragraph_runs.is_some() => {
                        if let Some(val) =
                            get_attribute_value(event, b"val").and_then(|v| v.parse::<i32>().ok())
                        {
                            match in_spacing_parent {
                                Some("lnSpc") => {
                                    current_paragraph_properties.line_spacing_pct =
                                        Some(val as u32);
                                    current_paragraph_properties.line_spacing =
                                        Some(LineSpacing::percent(val));
                                }
                                Some("spcBef") => {
                                    current_paragraph_properties.space_before =
                                        Some(SpacingValue::percent(val));
                                }
                                Some("spcAft") => {
                                    current_paragraph_properties.space_after =
                                        Some(SpacingValue::percent(val));
                                }
                                _ => {}
                            }
                        }
                    }
                    b"spcPts" if current_paragraph_runs.is_some() => {
                        if let Some(val) =
                            get_attribute_value(event, b"val").and_then(|v| v.parse::<i32>().ok())
                        {
                            match in_spacing_parent {
                                Some("lnSpc") => {
                                    current_paragraph_properties.line_spacing_pts =
                                        Some(val as u32);
                                    current_paragraph_properties.line_spacing =
                                        Some(LineSpacing::points(val));
                                }
                                Some("spcBef") => {
                                    current_paragraph_properties.space_before_pts =
                                        Some(val as u32);
                                    current_paragraph_properties.space_before =
                                        Some(SpacingValue::points(val));
                                }
                                Some("spcAft") => {
                                    current_paragraph_properties.space_after_pts = Some(val as u32);
                                    current_paragraph_properties.space_after =
                                        Some(SpacingValue::points(val));
                                }
                                _ => {}
                            }
                        }
                    }
                    b"t" => {
                        if let Some(paragraph_runs) = current_paragraph_runs.as_mut() {
                            paragraph_runs.push(ParsedRunData {
                                text: String::new(),
                                properties: std::mem::take(&mut current_run_properties),
                            });
                        }
                    }
                    _ => {
                        if shape_depth == 1 && !matches!(local, b"nvSpPr" | b"spPr" | b"txBody") {
                            current_shape_unknown_children
                                .push(RawXmlNode::from_empty_element(event));
                        }
                    }
                }
            }
            Event::Text(ref event) if in_shape && in_text => {
                let text = event
                    .xml_content()
                    .map_err(quick_xml::Error::from)?
                    .into_owned();
                current_text.push_str(text.as_str());
            }
            Event::CData(ref event) if in_shape && in_text => {
                let text = String::from_utf8_lossy(event.as_ref());
                current_text.push_str(text.as_ref());
            }
            Event::End(ref event) => {
                if !in_shape {
                    buffer.clear();
                    continue;
                }

                let event_name = event.name();
                let local = local_name(event_name.as_ref());
                match local {
                    b"t" if in_text => {
                        if let Some(paragraph_runs) = current_paragraph_runs.as_mut() {
                            paragraph_runs.push(ParsedRunData {
                                text: std::mem::take(&mut current_text),
                                properties: std::mem::take(&mut current_run_properties),
                            });
                        }
                        in_text = false;
                    }
                    b"rPr" if in_rpr => {
                        in_rpr = false;
                        in_rpr_solid_fill = false;
                    }
                    b"solidFill" if in_rpr_solid_fill => {
                        // Finalize building_color -> run font color.
                        if let Some(color) = building_color.take() {
                            current_run_properties.font_color = Some(color);
                        }
                        in_rpr_solid_fill = false;
                    }
                    b"lnSpc" | b"spcBef" | b"spcAft" if in_spacing_parent.is_some() => {
                        in_spacing_parent = None;
                    }
                    b"pPr" if in_ppr => {
                        in_ppr = false;
                        in_bu_clr = false;
                    }
                    b"buClr" if in_bu_clr => {
                        // Finalize building_color -> bullet color.
                        if let Some(color) = building_color.take() {
                            current_paragraph_properties.bullet.color = Some(color);
                        }
                        in_bu_clr = false;
                    }
                    b"ln" if in_outline => {
                        // Finalize building_color -> outline color.
                        if let Some(color) = building_color.take() {
                            current_shape_outline.color = Some(color);
                        }
                        in_outline = false;
                        in_outline_solid_fill = false;
                    }
                    b"solidFill" if in_outline_solid_fill => {
                        // Finalize building_color -> outline color.
                        if let Some(color) = building_color.take() {
                            current_shape_outline.color = Some(color);
                        }
                        in_outline_solid_fill = false;
                    }
                    b"gradFill" if in_grad_fill => {
                        in_grad_fill = false;
                    }
                    b"gs" if in_grad_gs => {
                        // Finalize building_color -> gradient stop color.
                        if let Some(color) = building_color.take() {
                            if let Some(ref mut grad) = current_shape_gradient_fill {
                                if let Some(last) = grad.stops.last_mut() {
                                    last.color = Some(color);
                                }
                            }
                        }
                        in_grad_gs = false;
                    }
                    b"pattFill" if in_patt_fill => {
                        in_patt_fill = false;
                        in_patt_fg = false;
                        in_patt_bg = false;
                    }
                    b"fgClr" if in_patt_fg => {
                        // Finalize building_color -> pattern foreground color.
                        if let Some(color) = building_color.take() {
                            if let Some(ref mut pattern) = current_shape_pattern_fill {
                                pattern.foreground_color = Some(color);
                            }
                        }
                        in_patt_fg = false;
                    }
                    b"bgClr" if in_patt_bg => {
                        // Finalize building_color -> pattern background color.
                        if let Some(color) = building_color.take() {
                            if let Some(ref mut pattern) = current_shape_pattern_fill {
                                pattern.background_color = Some(color);
                            }
                        }
                        in_patt_bg = false;
                    }
                    b"blipFill" if in_blip_fill => {
                        in_blip_fill = false;
                    }
                    b"effectLst" if in_effect_lst => {
                        in_effect_lst = false;
                        in_outer_shdw = false;
                        in_glow = false;
                    }
                    b"outerShdw" if in_outer_shdw => {
                        // Finalize building_color -> shadow color.
                        if let Some(color) = building_color.take() {
                            if let Some(ref mut shadow) = current_shape_shadow {
                                shadow.color_full = Some(color);
                            }
                        }
                        in_outer_shdw = false;
                    }
                    b"glow" if in_glow => {
                        // Finalize building_color -> glow color.
                        if let Some(color) = building_color.take() {
                            if let Some(ref mut glow) = current_shape_glow {
                                glow.color_full = Some(color);
                            }
                        }
                        in_glow = false;
                    }
                    b"bodyPr" if in_body_pr => {
                        in_body_pr = false;
                    }
                    b"p" if in_tx_body => {
                        if let Some(paragraph_runs) = current_paragraph_runs.take() {
                            current_shape_paragraphs.push(ParsedParagraphData {
                                runs: paragraph_runs,
                                properties: std::mem::take(&mut current_paragraph_properties),
                            });
                        }
                    }
                    b"txBody" => {
                        in_tx_body = false;
                        in_body_pr = false;
                    }
                    b"xfrm" if in_xfrm => {
                        in_xfrm = false;
                    }
                    b"solidFill" if in_shape_solid_fill => {
                        // Finalize building_color -> shape solid fill color.
                        if let Some(color) = building_color.take() {
                            current_shape_solid_fill_color = Some(color);
                        }
                        in_shape_solid_fill = false;
                    }
                    b"spPr" => {
                        in_sp_pr = false;
                        sp_pr_depth = 0;
                        in_xfrm = false;
                        in_shape_solid_fill = false;
                        in_outline = false;
                        in_outline_solid_fill = false;
                        in_grad_fill = false;
                        in_patt_fill = false;
                        in_blip_fill = false;
                        in_effect_lst = false;
                        in_outer_shdw = false;
                        in_glow = false;
                    }
                    _ => {}
                }

                if in_sp_pr && local != b"spPr" {
                    sp_pr_depth = sp_pr_depth.saturating_sub(1);
                }

                shape_depth = shape_depth.saturating_sub(1);
                if shape_depth == 0 {
                    let outline_to_set = if current_shape_outline.is_set() {
                        Some(std::mem::take(&mut current_shape_outline))
                    } else {
                        None
                    };
                    let shape = build_shape(ParsedShapeData {
                        name: std::mem::take(&mut current_shape_name),
                        paragraphs: std::mem::take(&mut current_shape_paragraphs),
                        placeholder_kind: std::mem::take(&mut current_placeholder_kind),
                        placeholder_idx: current_placeholder_idx.take(),
                        shape_type: current_shape_type,
                        geometry: shape_geometry_from_parts(
                            current_shape_xfrm_offset.take(),
                            current_shape_xfrm_extents.take(),
                        ),
                        preset_geometry: std::mem::take(&mut current_shape_preset_geometry),
                        solid_fill_srgb: std::mem::take(&mut current_shape_solid_fill_srgb),
                        solid_fill_alpha: current_shape_solid_fill_alpha.take(),
                        solid_fill_color: current_shape_solid_fill_color.take(),
                        outline: outline_to_set,
                        gradient_fill: current_shape_gradient_fill.take(),
                        pattern_fill: current_shape_pattern_fill.take(),
                        picture_fill: current_shape_picture_fill.take(),
                        no_fill: current_shape_no_fill,
                        rotation: current_shape_rotation.take(),
                        flip_h: current_shape_flip_h,
                        flip_v: current_shape_flip_v,
                        hidden: current_shape_hidden,
                        alt_text: current_shape_alt_text.take(),
                        alt_text_title: current_shape_alt_text_title.take(),
                        is_smartart: current_shape_is_smartart,
                        is_connector: current_shape_is_connector,
                        start_connection: current_shape_start_connection.take(),
                        end_connection: current_shape_end_connection.take(),
                        media: current_shape_media.take(),
                        shadow: current_shape_shadow.take(),
                        glow: current_shape_glow.take(),
                        reflection: current_shape_reflection.take(),
                        text_anchor: current_shape_text_anchor.take(),
                        auto_fit: current_shape_auto_fit.take(),
                        text_direction: current_shape_text_direction.take(),
                        text_columns: current_shape_text_columns.take(),
                        text_column_spacing: current_shape_text_column_spacing.take(),
                        text_inset_left: current_shape_text_inset_left.take(),
                        text_inset_right: current_shape_text_inset_right.take(),
                        text_inset_top: current_shape_text_inset_top.take(),
                        text_inset_bottom: current_shape_text_inset_bottom.take(),
                        word_wrap: current_shape_word_wrap.take(),
                        body_pr_rot: current_shape_body_pr_rot.take(),
                        body_pr_rtl_col: current_shape_body_pr_rtl_col.take(),
                        body_pr_from_word_art: current_shape_body_pr_from_word_art.take(),
                        body_pr_force_aa: current_shape_body_pr_force_aa.take(),
                        body_pr_compat_ln_spc: current_shape_body_pr_compat_ln_spc.take(),
                        custom_geometry_raw: current_shape_custom_geometry_raw.take(),
                        preset_geometry_adjustments: current_shape_preset_geometry_adjustments
                            .take(),
                        unknown_attrs: std::mem::take(&mut current_shape_unknown_attrs),
                        unknown_children: std::mem::take(&mut current_shape_unknown_children),
                        lst_style_raw: current_lst_style_raw.take(),
                        body_pr_unknown_children: std::mem::take(
                            &mut current_body_pr_unknown_children,
                        ),
                        body_pr_unknown_attrs: std::mem::take(&mut current_body_pr_unknown_attrs),
                    });
                    shapes.push(shape);

                    in_shape = false;
                    current_shape_type = ShapeType::AutoShape;
                    current_placeholder_kind = None;
                    current_placeholder_idx = None;
                    current_shape_xfrm_offset = None;
                    current_shape_xfrm_extents = None;
                    current_shape_preset_geometry = None;
                    current_shape_solid_fill_srgb = None;
                    current_shape_outline = ShapeOutline::default();
                    current_shape_gradient_fill = None;
                    current_shape_pattern_fill = None;
                    current_shape_picture_fill = None;
                    current_shape_no_fill = false;
                    current_shape_rotation = None;
                    current_shape_flip_h = false;
                    current_shape_flip_v = false;
                    current_shape_custom_geometry_raw = None;
                    current_shape_preset_geometry_adjustments = None;
                    current_shape_hidden = false;
                    current_shape_alt_text = None;
                    current_shape_alt_text_title = None;
                    current_shape_is_smartart = false;
                    current_shape_is_connector = false;
                    current_shape_start_connection = None;
                    current_shape_end_connection = None;
                    current_shape_media = None;
                    current_shape_shadow = None;
                    current_shape_glow = None;
                    current_shape_reflection = None;
                    current_shape_text_anchor = None;
                    current_shape_auto_fit = None;
                    current_shape_text_direction = None;
                    current_shape_text_columns = None;
                    current_shape_text_column_spacing = None;
                    current_shape_text_inset_left = None;
                    current_shape_text_inset_right = None;
                    current_shape_text_inset_top = None;
                    current_shape_text_inset_bottom = None;
                    current_shape_word_wrap = None;
                    current_shape_body_pr_rot = None;
                    current_shape_body_pr_rtl_col = None;
                    current_shape_body_pr_from_word_art = None;
                    current_shape_body_pr_force_aa = None;
                    current_shape_body_pr_compat_ln_spc = None;
                    current_shape_unknown_attrs.clear();
                    current_shape_unknown_children.clear();
                    current_paragraph_runs = None;
                    in_tx_body = false;
                    in_text = false;
                    current_text.clear();
                    in_sp_pr = false;
                    sp_pr_depth = 0;
                    in_xfrm = false;
                    in_shape_solid_fill = false;
                    in_outline = false;
                    in_outline_solid_fill = false;
                    in_grad_fill = false;
                    in_patt_fill = false;
                    in_effect_lst = false;
                    in_outer_shdw = false;
                    in_glow = false;
                    in_body_pr = false;
                    in_ppr = false;
                    in_bu_clr = false;
                }
            }
            Event::Eof => break,
            _ => {}
        }
        buffer.clear();
    }

    Ok(shapes)
}

fn parse_slide_tables(xml: &[u8]) -> Result<Vec<Table>> {
    let mut reader = Reader::from_reader(Cursor::new(xml));
    reader.config_mut().trim_text(false);

    let mut buffer = Vec::new();
    let mut tables = Vec::new();

    let mut in_table = false;
    let mut table_depth = 0_usize;
    let mut current_cell_rows: Vec<Vec<ParsedTableCellData>> = Vec::new();
    let mut current_row: Option<Vec<ParsedTableCellData>> = None;
    let mut current_cell_data: Option<ParsedTableCellData> = None;
    let mut in_text = false;
    let mut current_text = String::new();

    // Feature #6: Column widths and row heights.
    let mut column_widths: Vec<i64> = Vec::new();
    let mut row_heights: Vec<i64> = Vec::new();

    // Feature #5: Cell formatting state.
    let mut in_tc_pr = false;
    let mut in_tc_pr_solid_fill = false;
    let mut in_tc_pr_border: Option<String> = None; // "lnL", "lnR", "lnT", "lnB"
    let mut in_tc_pr_border_solid_fill = false;
    let mut current_border: Option<CellBorder> = None;

    // Run-level formatting in table cells.
    let mut in_rpr = false;
    let mut cell_rpr_bold: Option<bool> = None;
    let mut cell_rpr_italic: Option<bool> = None;
    let mut cell_rpr_font_size: Option<u32> = None;
    let mut in_rpr_solid_fill = false;
    let mut cell_rpr_font_color: Option<String> = None;

    // Shared color accumulator for the table parser (mirrors main shape parser pattern).
    let mut building_color: Option<ShapeColor> = None;

    loop {
        match reader.read_event_into(&mut buffer)? {
            Event::Start(ref event) => {
                let event_name = event.name();
                let local = local_name(event_name.as_ref());

                if local == b"tbl" {
                    if !in_table {
                        in_table = true;
                        table_depth = 1;
                        current_cell_rows.clear();
                        current_row = None;
                        current_cell_data = None;
                        in_text = false;
                        current_text.clear();
                        column_widths.clear();
                        row_heights.clear();
                    } else {
                        table_depth = table_depth.saturating_add(1);
                    }
                    buffer.clear();
                    continue;
                }

                if !in_table {
                    buffer.clear();
                    continue;
                }

                table_depth = table_depth.saturating_add(1);
                match local {
                    b"tr" => {
                        let height = parse_i64_attribute_value(event, b"h").unwrap_or(0);
                        row_heights.push(height);
                        current_row = Some(Vec::new());
                    }
                    b"tc" => {
                        let grid_span =
                            get_attribute_value(event, b"gridSpan").and_then(|v| v.parse().ok());
                        let row_span =
                            get_attribute_value(event, b"rowSpan").and_then(|v| v.parse().ok());
                        let v_merge = get_attribute_value(event, b"vMerge")
                            .as_deref()
                            .and_then(parse_xml_bool)
                            .unwrap_or(false);
                        current_cell_data = Some(ParsedTableCellData {
                            text: String::new(),
                            fill_color_srgb: None,
                            fill_color: None,
                            borders: CellBorders::default(),
                            bold: None,
                            italic: None,
                            font_size: None,
                            font_color_srgb: None,
                            font_color: None,
                            grid_span,
                            row_span,
                            v_merge,
                            vertical_alignment: None,
                            margin_left: None,
                            margin_right: None,
                            margin_top: None,
                            margin_bottom: None,
                            text_direction: None,
                        });
                        in_rpr = false;
                        cell_rpr_bold = None;
                        cell_rpr_italic = None;
                        cell_rpr_font_size = None;
                        cell_rpr_font_color = None;
                    }
                    b"t" if current_cell_data.is_some() => {
                        in_text = true;
                        current_text.clear();
                    }
                    b"tcPr" if current_cell_data.is_some() => {
                        in_tc_pr = true;
                        // Parse tcPr attributes for cell properties.
                        if let Some(ref mut cell_data) = current_cell_data {
                            cell_data.vertical_alignment = get_attribute_value(event, b"anchor")
                                .and_then(|v| CellTextAnchor::from_xml(&v));
                            cell_data.margin_left =
                                get_attribute_value(event, b"marL").and_then(|v| v.parse().ok());
                            cell_data.margin_right =
                                get_attribute_value(event, b"marR").and_then(|v| v.parse().ok());
                            cell_data.margin_top =
                                get_attribute_value(event, b"marT").and_then(|v| v.parse().ok());
                            cell_data.margin_bottom =
                                get_attribute_value(event, b"marB").and_then(|v| v.parse().ok());
                            cell_data.text_direction = get_attribute_value(event, b"vert")
                                .and_then(|v| TextDirection::from_xml(&v));
                        }
                    }
                    b"solidFill" if in_tc_pr && in_tc_pr_border.is_some() => {
                        in_tc_pr_border_solid_fill = true;
                    }
                    b"solidFill" if in_tc_pr && !in_tc_pr_border_solid_fill => {
                        in_tc_pr_solid_fill = true;
                    }
                    b"solidFill" if in_rpr => {
                        in_rpr_solid_fill = true;
                    }
                    // srgbClr/schemeClr as Start events (when they have child transforms).
                    b"srgbClr" if in_tc_pr_border_solid_fill => {
                        if let Some(val) = get_attribute_value(event, b"val") {
                            if let Some(ref mut border) = current_border {
                                border.color_srgb = Some(val.clone());
                            }
                            building_color = Some(ShapeColor::srgb(val));
                        }
                    }
                    b"schemeClr" if in_tc_pr_border_solid_fill => {
                        if let Some(val) = get_attribute_value(event, b"val") {
                            building_color = Some(ShapeColor::scheme(val));
                        }
                    }
                    b"srgbClr" if in_tc_pr_solid_fill => {
                        if let Some(val) = get_attribute_value(event, b"val") {
                            if let Some(ref mut cell_data) = current_cell_data {
                                cell_data.fill_color_srgb = Some(val.clone());
                            }
                            building_color = Some(ShapeColor::srgb(val));
                        }
                    }
                    b"schemeClr" if in_tc_pr_solid_fill => {
                        if let Some(val) = get_attribute_value(event, b"val") {
                            building_color = Some(ShapeColor::scheme(val));
                        }
                    }
                    b"srgbClr" if in_rpr_solid_fill => {
                        if let Some(val) = get_attribute_value(event, b"val") {
                            cell_rpr_font_color = Some(val.clone());
                            building_color = Some(ShapeColor::srgb(val));
                        }
                    }
                    b"schemeClr" if in_rpr_solid_fill => {
                        if let Some(val) = get_attribute_value(event, b"val") {
                            building_color = Some(ShapeColor::scheme(val));
                        }
                    }
                    b"lnL" | b"lnR" | b"lnT" | b"lnB" | b"lnTlToBr" | b"lnBlToTr" if in_tc_pr => {
                        let border_name = String::from_utf8_lossy(local).into_owned();
                        let width = parse_i64_attribute_value(event, b"w");
                        current_border = Some(CellBorder {
                            width_emu: width,
                            color_srgb: None,
                            color: None,
                        });
                        in_tc_pr_border = Some(border_name);
                    }
                    b"rPr" if current_cell_data.is_some() && !in_tc_pr => {
                        in_rpr = true;
                        cell_rpr_bold = get_attribute_value(event, b"b")
                            .as_deref()
                            .and_then(parse_xml_bool);
                        cell_rpr_italic = get_attribute_value(event, b"i")
                            .as_deref()
                            .and_then(parse_xml_bool);
                        cell_rpr_font_size =
                            get_attribute_value(event, b"sz").and_then(|v| v.parse().ok());
                    }
                    _ => {}
                }
            }
            Event::Empty(ref event) => {
                if !in_table {
                    buffer.clear();
                    continue;
                }

                let event_name = event.name();
                let local = local_name(event_name.as_ref());
                match local {
                    b"gridCol" => {
                        let width = parse_i64_attribute_value(event, b"w").unwrap_or(0);
                        column_widths.push(width);
                    }
                    b"srgbClr" if in_tc_pr_border_solid_fill => {
                        if let Some(val) = get_attribute_value(event, b"val") {
                            if let Some(ref mut border) = current_border {
                                border.color_srgb = Some(val.clone());
                            }
                            building_color = Some(ShapeColor::srgb(val));
                        }
                    }
                    b"schemeClr" if in_tc_pr_border_solid_fill => {
                        if let Some(val) = get_attribute_value(event, b"val") {
                            building_color = Some(ShapeColor::scheme(val));
                        }
                    }
                    b"srgbClr" if in_tc_pr_solid_fill => {
                        if let Some(val) = get_attribute_value(event, b"val") {
                            if let Some(ref mut cell_data) = current_cell_data {
                                cell_data.fill_color_srgb = Some(val.clone());
                            }
                            building_color = Some(ShapeColor::srgb(val));
                        }
                    }
                    b"schemeClr" if in_tc_pr_solid_fill => {
                        if let Some(val) = get_attribute_value(event, b"val") {
                            building_color = Some(ShapeColor::scheme(val));
                        }
                    }
                    b"srgbClr" if in_rpr_solid_fill => {
                        if let Some(val) = get_attribute_value(event, b"val") {
                            cell_rpr_font_color = Some(val.clone());
                            building_color = Some(ShapeColor::srgb(val));
                        }
                    }
                    b"schemeClr" if in_rpr_solid_fill => {
                        if let Some(val) = get_attribute_value(event, b"val") {
                            building_color = Some(ShapeColor::scheme(val));
                        }
                    }
                    // Color transforms and alpha as Empty elements in table context.
                    b"alpha" | b"lumMod" | b"lumOff" | b"tint" | b"shade" | b"satMod"
                    | b"satOff" | b"hueOff" | b"hueMod"
                        if building_color.is_some() =>
                    {
                        if let Some(val_str) = get_attribute_value(event, b"val") {
                            if let Ok(val) = val_str.parse::<i32>() {
                                if let Some(ref mut color) = building_color {
                                    if local == b"alpha" {
                                        // Set alpha as percentage (thousandths -> 0-100).
                                        let pct = (val / 1000).clamp(0, 100) as u8;
                                        match color {
                                            ShapeColor::SrgbClr { alpha, .. }
                                            | ShapeColor::SchemeClr { alpha, .. } => {
                                                *alpha = Some(pct);
                                            }
                                        }
                                    } else {
                                        let transform = ColorTransform::from_xml(local, val);
                                        match color {
                                            ShapeColor::SrgbClr { transforms, .. }
                                            | ShapeColor::SchemeClr { transforms, .. } => {
                                                transforms.push(transform);
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                    b"rPr" if current_cell_data.is_some() && !in_tc_pr => {
                        // Self-closing rPr.
                        cell_rpr_bold = get_attribute_value(event, b"b")
                            .as_deref()
                            .and_then(parse_xml_bool);
                        cell_rpr_italic = get_attribute_value(event, b"i")
                            .as_deref()
                            .and_then(parse_xml_bool);
                        cell_rpr_font_size =
                            get_attribute_value(event, b"sz").and_then(|v| v.parse().ok());
                    }
                    b"tcPr" if current_cell_data.is_some() => {
                        // Self-closing <a:tcPr/> with attributes.
                        if let Some(ref mut cell_data) = current_cell_data {
                            cell_data.vertical_alignment = get_attribute_value(event, b"anchor")
                                .and_then(|v| CellTextAnchor::from_xml(&v));
                            cell_data.margin_left =
                                get_attribute_value(event, b"marL").and_then(|v| v.parse().ok());
                            cell_data.margin_right =
                                get_attribute_value(event, b"marR").and_then(|v| v.parse().ok());
                            cell_data.margin_top =
                                get_attribute_value(event, b"marT").and_then(|v| v.parse().ok());
                            cell_data.margin_bottom =
                                get_attribute_value(event, b"marB").and_then(|v| v.parse().ok());
                            cell_data.text_direction = get_attribute_value(event, b"vert")
                                .and_then(|v| TextDirection::from_xml(&v));
                        }
                    }
                    b"lnL" | b"lnR" | b"lnT" | b"lnB" | b"lnTlToBr" | b"lnBlToTr" if in_tc_pr => {
                        // Self-closing border (no fill child).
                        let width = parse_i64_attribute_value(event, b"w");
                        let border = CellBorder {
                            width_emu: width,
                            color_srgb: None,
                            color: None,
                        };
                        if let Some(ref mut cell_data) = current_cell_data {
                            match local {
                                b"lnL" => cell_data.borders.left = Some(border),
                                b"lnR" => cell_data.borders.right = Some(border),
                                b"lnT" => cell_data.borders.top = Some(border),
                                b"lnB" => cell_data.borders.bottom = Some(border),
                                b"lnTlToBr" => cell_data.borders.diagonal_down = Some(border),
                                b"lnBlToTr" => cell_data.borders.diagonal_up = Some(border),
                                _ => {}
                            }
                        }
                    }
                    _ => {}
                }
            }
            Event::Text(ref event) if in_table && in_text => {
                let text = event
                    .xml_content()
                    .map_err(quick_xml::Error::from)?
                    .into_owned();
                current_text.push_str(text.as_str());
            }
            Event::CData(ref event) if in_table && in_text => {
                let text = String::from_utf8_lossy(event.as_ref());
                current_text.push_str(text.as_ref());
            }
            Event::End(ref event) => {
                if !in_table {
                    buffer.clear();
                    continue;
                }

                let event_name = event.name();
                let local = local_name(event_name.as_ref());
                match local {
                    b"t" if in_text => {
                        if let Some(ref mut cell_data) = current_cell_data {
                            cell_data.text.push_str(current_text.as_str());
                        }
                        current_text.clear();
                        in_text = false;
                    }
                    b"srgbClr" | b"schemeClr"
                        if in_tc_pr_border_solid_fill
                            || in_tc_pr_solid_fill
                            || in_rpr_solid_fill =>
                    {
                        // End of a color element with children (transforms).
                        // building_color stays — it will be finalized at solidFill/border End.
                    }
                    b"solidFill" if in_tc_pr_border_solid_fill => {
                        // Border fill color finalized at the border End event.
                        in_tc_pr_border_solid_fill = false;
                    }
                    b"solidFill" if in_tc_pr_solid_fill => {
                        if let Some(color) = building_color.take() {
                            if let Some(ref mut cell_data) = current_cell_data {
                                cell_data.fill_color = Some(color);
                            }
                        }
                        in_tc_pr_solid_fill = false;
                    }
                    b"solidFill" if in_rpr_solid_fill => {
                        if let Some(color) = building_color.take() {
                            if let Some(ref mut cell_data) = current_cell_data {
                                cell_data.font_color = Some(color);
                            }
                        }
                        in_rpr_solid_fill = false;
                    }
                    b"lnL" | b"lnR" | b"lnT" | b"lnB" | b"lnTlToBr" | b"lnBlToTr"
                        if in_tc_pr_border.is_some() =>
                    {
                        if let Some(mut border) = current_border.take() {
                            if let Some(color) = building_color.take() {
                                border.color = Some(color);
                            }
                            if let Some(ref mut cell_data) = current_cell_data {
                                let border_name = in_tc_pr_border.as_deref().unwrap_or("");
                                match border_name {
                                    "lnL" => cell_data.borders.left = Some(border),
                                    "lnR" => cell_data.borders.right = Some(border),
                                    "lnT" => cell_data.borders.top = Some(border),
                                    "lnB" => cell_data.borders.bottom = Some(border),
                                    "lnTlToBr" => cell_data.borders.diagonal_down = Some(border),
                                    "lnBlToTr" => cell_data.borders.diagonal_up = Some(border),
                                    _ => {}
                                }
                            }
                        }
                        in_tc_pr_border = None;
                    }
                    b"tcPr" if in_tc_pr => {
                        in_tc_pr = false;
                        in_tc_pr_solid_fill = false;
                    }
                    b"rPr" if in_rpr => {
                        in_rpr = false;
                        in_rpr_solid_fill = false;
                    }
                    b"tc" => {
                        // Apply accumulated run properties to cell data.
                        if let Some(ref mut cell_data) = current_cell_data {
                            if cell_rpr_bold.is_some() {
                                cell_data.bold = cell_rpr_bold;
                            }
                            if cell_rpr_italic.is_some() {
                                cell_data.italic = cell_rpr_italic;
                            }
                            if cell_rpr_font_size.is_some() {
                                cell_data.font_size = cell_rpr_font_size;
                            }
                            if cell_rpr_font_color.is_some() {
                                cell_data.font_color_srgb = cell_rpr_font_color.take();
                            }
                        }
                        if let Some(row) = current_row.as_mut() {
                            row.push(current_cell_data.take().unwrap_or(ParsedTableCellData {
                                text: String::new(),
                                fill_color_srgb: None,
                                fill_color: None,
                                borders: CellBorders::default(),
                                bold: None,
                                italic: None,
                                font_size: None,
                                font_color_srgb: None,
                                font_color: None,
                                grid_span: None,
                                row_span: None,
                                v_merge: false,
                                vertical_alignment: None,
                                margin_left: None,
                                margin_right: None,
                                margin_top: None,
                                margin_bottom: None,
                                text_direction: None,
                            }));
                        }
                        in_tc_pr = false;
                        in_rpr = false;
                    }
                    b"tr" => {
                        if let Some(row) = current_row.take() {
                            current_cell_rows.push(row);
                        }
                    }
                    _ => {}
                }

                table_depth = table_depth.saturating_sub(1);
                if table_depth == 0 {
                    let cell_rows = std::mem::take(&mut current_cell_rows);
                    let table = build_table_from_parsed_cells(
                        cell_rows,
                        std::mem::take(&mut column_widths),
                        std::mem::take(&mut row_heights),
                    );
                    tables.push(table);

                    in_table = false;
                    current_row = None;
                    current_cell_data = None;
                    in_text = false;
                    current_text.clear();
                    in_tc_pr = false;
                    in_tc_pr_solid_fill = false;
                    in_tc_pr_border = None;
                    in_tc_pr_border_solid_fill = false;
                    current_border = None;
                    in_rpr = false;
                    in_rpr_solid_fill = false;
                }
            }
            Event::Eof => break,
            _ => {}
        }

        buffer.clear();
    }

    Ok(tables)
}

fn parse_slide_pictures(xml: &[u8]) -> Result<Vec<ParsedImageRef>> {
    let mut reader = Reader::from_reader(Cursor::new(xml));
    reader.config_mut().trim_text(true);

    let mut buffer = Vec::new();
    let mut image_refs = Vec::new();

    let mut in_picture = false;
    let mut picture_depth = 0_usize;
    let mut current_name: Option<String> = None;
    let mut current_relationship_id: Option<String> = None;

    loop {
        match reader.read_event_into(&mut buffer)? {
            Event::Start(ref event) => {
                let event_name = event.name();
                let local = local_name(event_name.as_ref());

                if local == b"pic" {
                    if !in_picture {
                        in_picture = true;
                        picture_depth = 1;
                        current_name = None;
                        current_relationship_id = None;
                    } else {
                        picture_depth = picture_depth.saturating_add(1);
                    }
                    buffer.clear();
                    continue;
                }

                if !in_picture {
                    buffer.clear();
                    continue;
                }

                picture_depth = picture_depth.saturating_add(1);
                if local == b"cNvPr" {
                    current_name = get_attribute_value(event, b"name");
                } else if local == b"blip" {
                    current_relationship_id = get_attribute_value(event, b"embed");
                }
            }
            Event::Empty(ref event) => {
                if !in_picture {
                    buffer.clear();
                    continue;
                }

                let event_name = event.name();
                let local = local_name(event_name.as_ref());
                if local == b"cNvPr" {
                    current_name = get_attribute_value(event, b"name");
                } else if local == b"blip" {
                    current_relationship_id = get_attribute_value(event, b"embed");
                }
            }
            Event::End(ref event) => {
                if !in_picture {
                    buffer.clear();
                    continue;
                }

                picture_depth = picture_depth.saturating_sub(1);
                if picture_depth == 0 {
                    if let Some(relationship_id) = current_relationship_id.take() {
                        image_refs.push(ParsedImageRef {
                            relationship_id,
                            name: current_name.take(),
                        });
                    }
                    in_picture = false;
                } else if local_name(event.name().as_ref()) == b"pic" {
                    in_picture = false;
                }
            }
            Event::Eof => break,
            _ => {}
        }

        buffer.clear();
    }

    Ok(image_refs)
}

fn parse_slide_charts(xml: &[u8]) -> Result<Vec<ParsedChartRef>> {
    let mut reader = Reader::from_reader(Cursor::new(xml));
    reader.config_mut().trim_text(true);

    let mut buffer = Vec::new();
    let mut chart_refs = Vec::new();

    loop {
        match reader.read_event_into(&mut buffer)? {
            Event::Start(ref event) | Event::Empty(ref event)
                if local_name(event.name().as_ref()) == b"chart" =>
            {
                if let Some(relationship_id) = get_exact_attribute_value(event, b"r:id") {
                    chart_refs.push(ParsedChartRef { relationship_id });
                }
            }
            Event::Eof => break,
            _ => {}
        }
        buffer.clear();
    }

    Ok(chart_refs)
}

fn parse_slide_transition(xml: &[u8]) -> Result<Option<SlideTransition>> {
    let mut reader = Reader::from_reader(Cursor::new(xml));
    reader.config_mut().trim_text(true);

    let mut buffer = Vec::new();
    let mut in_transition = false;
    let mut transition_depth = 0_usize;
    let mut transition_kind = SlideTransitionKind::Unspecified;
    let mut advance_on_click: Option<bool> = None;
    let mut advance_after_ms: Option<u32> = None;
    let mut speed: Option<TransitionSpeed> = None;
    let mut sound: Option<TransitionSound> = None;
    let mut in_snd_ac = false;

    loop {
        match reader.read_event_into(&mut buffer)? {
            Event::Start(ref event) => {
                let event_name = event.name();
                let local = local_name(event_name.as_ref());
                if local == b"transition" {
                    in_transition = true;
                    transition_depth = 1;
                    transition_kind = SlideTransitionKind::Unspecified;
                    advance_on_click = get_attribute_value(event, b"advClick")
                        .and_then(|value| parse_xml_bool(value.as_str()));
                    advance_after_ms = get_attribute_value(event, b"advTm")
                        .and_then(|value| value.parse::<u32>().ok());
                    speed = get_attribute_value(event, b"spd")
                        .and_then(|v| TransitionSpeed::from_xml(&v));
                    buffer.clear();
                    continue;
                }

                if in_transition {
                    transition_depth = transition_depth.saturating_add(1);
                    if transition_depth == 2 {
                        update_transition_kind_from_xml(&mut transition_kind, local);
                    }
                    if local == b"sndAc" {
                        in_snd_ac = true;
                    }
                    if in_snd_ac && local == b"snd" {
                        let snd_name = get_attribute_value(event, b"name").unwrap_or_default();
                        let snd_rid =
                            get_exact_attribute_value(event, b"r:embed").unwrap_or_default();
                        sound = Some(TransitionSound::new(snd_name, snd_rid));
                    }
                }
            }
            Event::Empty(ref event) => {
                let event_name = event.name();
                let local = local_name(event_name.as_ref());
                if local == b"transition" {
                    let mut transition = SlideTransition::new(SlideTransitionKind::Unspecified);
                    transition.set_advance_on_click(
                        get_attribute_value(event, b"advClick")
                            .and_then(|value| parse_xml_bool(value.as_str())),
                    );
                    transition.set_advance_after_ms(
                        get_attribute_value(event, b"advTm")
                            .and_then(|value| value.parse::<u32>().ok()),
                    );
                    transition.set_speed(
                        get_attribute_value(event, b"spd")
                            .and_then(|v| TransitionSpeed::from_xml(&v)),
                    );
                    return Ok(Some(transition));
                }

                if in_transition {
                    if transition_depth == 1 {
                        update_transition_kind_from_xml(&mut transition_kind, local);
                    }
                    if in_snd_ac && local == b"snd" {
                        let snd_name = get_attribute_value(event, b"name").unwrap_or_default();
                        let snd_rid =
                            get_exact_attribute_value(event, b"r:embed").unwrap_or_default();
                        sound = Some(TransitionSound::new(snd_name, snd_rid));
                    }
                }
            }
            Event::End(ref event) if in_transition => {
                let event_name = event.name();
                let local_bytes = local_name(event_name.as_ref());
                transition_depth = transition_depth.saturating_sub(1);
                if local_bytes == b"sndAc" {
                    in_snd_ac = false;
                }
                if transition_depth == 0 && local_bytes == b"transition" {
                    let mut transition = SlideTransition::new(transition_kind);
                    transition.set_advance_on_click(advance_on_click);
                    transition.set_advance_after_ms(advance_after_ms);
                    transition.set_speed(speed);
                    transition.set_sound(sound);
                    return Ok(Some(transition));
                }
            }
            Event::Eof => break,
            _ => {}
        }
        buffer.clear();
    }

    Ok(None)
}

fn parse_slide_timing(xml: &[u8]) -> Result<Option<SlideTiming>> {
    let mut reader = Reader::from_reader(Cursor::new(xml));
    reader.config_mut().trim_text(false);

    let mut buffer = Vec::new();
    let mut timing_depth = 0_usize;
    let mut timing_inner_writer: Option<Writer<Vec<u8>>> = None;

    loop {
        match reader.read_event_into(&mut buffer)? {
            Event::Start(ref event) => {
                let is_timing = local_name(event.name().as_ref()) == b"timing";
                if timing_depth == 0 && is_timing {
                    timing_depth = 1;
                    timing_inner_writer = Some(Writer::new(Vec::new()));
                } else if timing_depth > 0 {
                    timing_depth = timing_depth.saturating_add(1);
                    if let Some(writer) = timing_inner_writer.as_mut() {
                        writer.write_event(Event::Start(event.to_owned()))?;
                    }
                }
            }
            Event::Empty(ref event) => {
                let is_timing = local_name(event.name().as_ref()) == b"timing";
                if timing_depth == 0 && is_timing {
                    return Ok(Some(SlideTiming::new("")));
                }

                if timing_depth > 0 {
                    if let Some(writer) = timing_inner_writer.as_mut() {
                        writer.write_event(Event::Empty(event.to_owned()))?;
                    }
                }
            }
            Event::End(ref event) => {
                if timing_depth == 0 {
                    buffer.clear();
                    continue;
                }

                timing_depth = timing_depth.saturating_sub(1);
                if timing_depth == 0 && local_name(event.name().as_ref()) == b"timing" {
                    let Some(writer) = timing_inner_writer.take() else {
                        return Ok(Some(SlideTiming::new("")));
                    };
                    let raw_inner_xml =
                        String::from_utf8(writer.into_inner()).map_err(|error| {
                            PptxError::UnsupportedPackage(format!(
                                "slide timing inner xml is not UTF-8: {error}"
                            ))
                        })?;
                    return Ok(Some(SlideTiming::new(raw_inner_xml)));
                }

                if let Some(writer) = timing_inner_writer.as_mut() {
                    writer.write_event(Event::End(event.to_owned()))?;
                }
            }
            Event::Text(ref event) if timing_depth > 0 => {
                if let Some(writer) = timing_inner_writer.as_mut() {
                    writer.write_event(Event::Text(event.to_owned()))?;
                }
            }
            Event::CData(ref event) if timing_depth > 0 => {
                if let Some(writer) = timing_inner_writer.as_mut() {
                    writer.write_event(Event::CData(event.to_owned()))?;
                }
            }
            Event::Comment(ref event) if timing_depth > 0 => {
                if let Some(writer) = timing_inner_writer.as_mut() {
                    writer.write_event(Event::Comment(event.to_owned()))?;
                }
            }
            Event::PI(ref event) if timing_depth > 0 => {
                if let Some(writer) = timing_inner_writer.as_mut() {
                    writer.write_event(Event::PI(event.to_owned()))?;
                }
            }
            Event::Eof => break,
            _ => {}
        }
        buffer.clear();
    }

    Ok(None)
}

/// Parse unknown top-level children of `p:sld` for roundtrip fidelity.
///
/// Known children (`cSld`, `transition`, `timing`) are skipped. Everything
/// else is captured as `RawXmlNode` so it survives a save round-trip.
fn parse_slide_unknown_children(xml: &[u8]) -> Result<Vec<RawXmlNode>> {
    let mut reader = Reader::from_reader(Cursor::new(xml));
    reader.config_mut().trim_text(true);
    let mut buffer = Vec::new();
    let mut unknown = Vec::new();
    let mut sld_depth = 0_usize;
    static KNOWN_LOCALS: &[&[u8]] = &[b"cSld", b"transition", b"timing", b"hf"];

    loop {
        match reader.read_event_into(&mut buffer)? {
            Event::Start(ref event) => {
                let event_name = event.name();
                let local = local_name(event_name.as_ref());
                if sld_depth == 0 && local == b"sld" {
                    sld_depth = 1;
                } else if sld_depth == 1 {
                    if KNOWN_LOCALS.contains(&local) {
                        // Skip the entire known subtree.
                        let mut depth = 1_usize;
                        let mut skip_buf = Vec::new();
                        loop {
                            match reader.read_event_into(&mut skip_buf)? {
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
                            skip_buf.clear();
                        }
                    } else {
                        unknown.push(RawXmlNode::read_element(&mut reader, event)?);
                    }
                }
            }
            Event::Empty(ref event) => {
                let event_name = event.name();
                let local = local_name(event_name.as_ref());
                if sld_depth == 0 && local == b"sld" {
                    // Empty <p:sld/> — unlikely but handle gracefully.
                    break;
                }
                if sld_depth == 1 && !KNOWN_LOCALS.contains(&local) {
                    unknown.push(RawXmlNode::from_empty_element(event));
                }
            }
            Event::End(ref event) => {
                let event_name = event.name();
                if sld_depth == 1 && local_name(event_name.as_ref()) == b"sld" {
                    break;
                }
            }
            Event::Eof => break,
            _ => {}
        }
        buffer.clear();
    }

    Ok(unknown)
}

fn parse_slide_notes_text(
    package: &Package,
    slide_uri: &PartUri,
    slide_part: &Part,
) -> Result<Option<String>> {
    let Some(relationship) = slide_part
        .relationships
        .get_first_by_type(NOTES_SLIDE_RELATIONSHIP_TYPE)
    else {
        return Ok(None);
    };

    if relationship.target_mode != TargetMode::Internal {
        return Err(PptxError::UnsupportedPackage(format!(
            "notes relationship `{}` is external",
            relationship.id
        )));
    }

    let notes_uri = slide_uri.resolve_relative(relationship.target.as_str())?;
    let notes_part = package.get_part(notes_uri.as_str()).ok_or_else(|| {
        PptxError::UnsupportedPackage(format!(
            "missing notes part `{}` for relationship `{}`",
            notes_uri.as_str(),
            relationship.id
        ))
    })?;
    Ok(Some(parse_notes_slide_text(notes_part.data.as_bytes())?))
}

fn parse_slide_comments(
    package: &Package,
    slide_uri: &PartUri,
    slide_part: &Part,
) -> Result<Vec<SlideComment>> {
    let Some(comments_relationship) = slide_part
        .relationships
        .get_first_by_type(COMMENTS_RELATIONSHIP_TYPE)
    else {
        return Ok(Vec::new());
    };

    if comments_relationship.target_mode != TargetMode::Internal {
        return Err(PptxError::UnsupportedPackage(format!(
            "comments relationship `{}` is external",
            comments_relationship.id
        )));
    }

    let comments_uri = slide_uri.resolve_relative(comments_relationship.target.as_str())?;
    let comments_part = package.get_part(comments_uri.as_str()).ok_or_else(|| {
        PptxError::UnsupportedPackage(format!(
            "missing comments part `{}` for relationship `{}`",
            comments_uri.as_str(),
            comments_relationship.id
        ))
    })?;

    let authors = if let Some(authors_relationship) = slide_part
        .relationships
        .get_first_by_type(COMMENT_AUTHORS_RELATIONSHIP_TYPE)
    {
        if authors_relationship.target_mode != TargetMode::Internal {
            HashMap::new()
        } else {
            let authors_uri = slide_uri.resolve_relative(authors_relationship.target.as_str())?;
            package
                .get_part(authors_uri.as_str())
                .map(|authors_part| parse_comment_authors_xml(authors_part.data.as_bytes()))
                .transpose()?
                .unwrap_or_default()
        }
    } else {
        HashMap::new()
    };
    let parsed_comments = parse_comments_xml(comments_part.data.as_bytes())?;

    let mut comments = Vec::with_capacity(parsed_comments.len());
    for parsed_comment in parsed_comments {
        let author = authors
            .get(&parsed_comment.author_id)
            .cloned()
            .unwrap_or_else(|| format!("Author {}", parsed_comment.author_id));
        comments.push(SlideComment::new(author, parsed_comment.text));
    }

    Ok(comments)
}

fn parse_notes_slide_text(xml: &[u8]) -> Result<String> {
    let mut reader = Reader::from_reader(Cursor::new(xml));
    reader.config_mut().trim_text(false);

    let mut buffer = Vec::new();
    let mut paragraphs = Vec::new();
    let mut current_paragraph = String::new();
    let mut paragraph_depth = 0_usize;
    let mut in_text = false;
    let mut current_text = String::new();

    loop {
        match reader.read_event_into(&mut buffer)? {
            Event::Start(ref event) => {
                let event_name = event.name();
                let local = local_name(event_name.as_ref());
                if local == b"p" {
                    paragraph_depth = paragraph_depth.saturating_add(1);
                    if paragraph_depth == 1 {
                        current_paragraph.clear();
                    }
                } else if local == b"t" && paragraph_depth > 0 {
                    in_text = true;
                    current_text.clear();
                }
            }
            Event::Text(ref event) if in_text => {
                let text = event
                    .xml_content()
                    .map_err(quick_xml::Error::from)?
                    .into_owned();
                current_text.push_str(text.as_str());
            }
            Event::CData(ref event) if in_text => {
                let text = String::from_utf8_lossy(event.as_ref());
                current_text.push_str(text.as_ref());
            }
            Event::End(ref event) => {
                let event_name = event.name();
                let local = local_name(event_name.as_ref());
                if local == b"t" && in_text {
                    current_paragraph.push_str(current_text.as_str());
                    current_text.clear();
                    in_text = false;
                } else if local == b"p" && paragraph_depth > 0 {
                    paragraph_depth = paragraph_depth.saturating_sub(1);
                    if paragraph_depth == 0 {
                        paragraphs.push(std::mem::take(&mut current_paragraph));
                    }
                }
            }
            Event::Eof => break,
            _ => {}
        }
        buffer.clear();
    }

    Ok(paragraphs.join("\n"))
}

fn parse_comment_authors_xml(xml: &[u8]) -> Result<HashMap<u32, String>> {
    let mut reader = Reader::from_reader(Cursor::new(xml));
    reader.config_mut().trim_text(true);

    let mut buffer = Vec::new();
    let mut authors = HashMap::new();

    loop {
        match reader.read_event_into(&mut buffer)? {
            Event::Start(ref event) | Event::Empty(ref event)
                if local_name(event.name().as_ref()) == b"cmAuthor" =>
            {
                let author_id = get_attribute_value(event, b"id")
                    .and_then(|id| id.parse::<u32>().ok())
                    .ok_or_else(|| {
                        PptxError::UnsupportedPackage(
                            "comment author entry missing valid `id` attribute".to_string(),
                        )
                    })?;
                let author_name = get_attribute_value(event, b"name").unwrap_or_default();
                authors.insert(author_id, author_name);
            }
            Event::Eof => break,
            _ => {}
        }
        buffer.clear();
    }

    Ok(authors)
}

fn parse_comments_xml(xml: &[u8]) -> Result<Vec<ParsedSlideComment>> {
    let mut reader = Reader::from_reader(Cursor::new(xml));
    reader.config_mut().trim_text(false);

    let mut buffer = Vec::new();
    let mut comments = Vec::new();

    let mut in_comment = false;
    let mut comment_depth = 0_usize;
    let mut in_text = false;
    let mut current_text = String::new();
    let mut current_author_id = 0_u32;

    loop {
        match reader.read_event_into(&mut buffer)? {
            Event::Start(ref event) => {
                let event_name = event.name();
                let local = local_name(event_name.as_ref());

                if local == b"cm" {
                    if !in_comment {
                        in_comment = true;
                        comment_depth = 1;
                        current_text.clear();
                        current_author_id = get_attribute_value(event, b"authorId")
                            .and_then(|value| value.parse::<u32>().ok())
                            .ok_or_else(|| {
                                PptxError::UnsupportedPackage(
                                    "comment entry missing valid `authorId` attribute".to_string(),
                                )
                            })?;
                    } else {
                        comment_depth = comment_depth.saturating_add(1);
                    }
                    buffer.clear();
                    continue;
                }

                if !in_comment {
                    buffer.clear();
                    continue;
                }

                comment_depth = comment_depth.saturating_add(1);
                if local == b"text" {
                    in_text = true;
                    current_text.clear();
                }
            }
            Event::Empty(ref event)
                if in_comment && local_name(event.name().as_ref()) == b"text" =>
            {
                in_text = false;
                current_text.clear();
            }
            Event::Text(ref event) if in_comment && in_text => {
                let text = event
                    .xml_content()
                    .map_err(quick_xml::Error::from)?
                    .into_owned();
                current_text.push_str(text.as_str());
            }
            Event::CData(ref event) if in_comment && in_text => {
                let text = String::from_utf8_lossy(event.as_ref());
                current_text.push_str(text.as_ref());
            }
            Event::End(ref event) => {
                if !in_comment {
                    buffer.clear();
                    continue;
                }

                let event_name = event.name();
                let local = local_name(event_name.as_ref());
                if local == b"text" && in_text {
                    in_text = false;
                }

                comment_depth = comment_depth.saturating_sub(1);
                if comment_depth == 0 {
                    comments.push(ParsedSlideComment {
                        author_id: current_author_id,
                        text: std::mem::take(&mut current_text),
                    });
                    in_comment = false;
                    in_text = false;
                }
            }
            Event::Eof => break,
            _ => {}
        }
        buffer.clear();
    }

    Ok(comments)
}

// ── Feature #9: Parse slide size from p:sldSz ──

fn parse_slide_size(xml: &[u8]) -> Result<(Option<i64>, Option<i64>)> {
    let mut reader = Reader::from_reader(Cursor::new(xml));
    reader.config_mut().trim_text(true);
    let mut buffer = Vec::new();

    loop {
        match reader.read_event_into(&mut buffer)? {
            Event::Empty(ref event) | Event::Start(ref event)
                if local_name(event.name().as_ref()) == b"sldSz" =>
            {
                let cx = get_attribute_value(event, b"cx").and_then(|v| v.parse::<i64>().ok());
                let cy = get_attribute_value(event, b"cy").and_then(|v| v.parse::<i64>().ok());
                return Ok((cx, cy));
            }
            Event::Eof => break,
            _ => {}
        }
        buffer.clear();
    }

    Ok((None, None))
}

/// Parsed presentation-level attributes from `<p:presentation>`.
#[derive(Debug, Clone, Default)]
struct ParsedPresentationAttrs {
    first_slide_number: Option<u32>,
    show_special_pls_on_title_sld: Option<bool>,
    right_to_left: Option<bool>,
}

fn parse_presentation_attrs(xml: &[u8]) -> Result<ParsedPresentationAttrs> {
    let mut reader = Reader::from_reader(Cursor::new(xml));
    reader.config_mut().trim_text(true);
    let mut buffer = Vec::new();

    loop {
        match reader.read_event_into(&mut buffer)? {
            Event::Start(ref event) | Event::Empty(ref event)
                if local_name(event.name().as_ref()) == b"presentation" =>
            {
                let first_slide_number =
                    get_attribute_value(event, b"firstSlideNum").and_then(|v| v.parse().ok());
                let show_special_pls_on_title_sld =
                    get_attribute_value(event, b"showSpecialPlsOnTitleSld")
                        .as_deref()
                        .and_then(parse_xml_bool);
                let right_to_left = get_attribute_value(event, b"rtl")
                    .as_deref()
                    .and_then(parse_xml_bool);
                return Ok(ParsedPresentationAttrs {
                    first_slide_number,
                    show_special_pls_on_title_sld,
                    right_to_left,
                });
            }
            Event::Eof => break,
            _ => {}
        }
        buffer.clear();
    }

    Ok(ParsedPresentationAttrs::default())
}

// ── Feature #7: Parse slide background ──

fn parse_slide_background(xml: &[u8]) -> Result<Option<SlideBackground>> {
    let mut reader = Reader::from_reader(Cursor::new(xml));
    reader.config_mut().trim_text(true);
    let mut buffer = Vec::new();

    let mut in_bg = false;
    let mut in_bg_pr = false;
    let mut in_solid_fill = false;
    let mut in_grad_fill = false;
    let mut grad_stops = Vec::new();
    let mut grad_lin_angle: Option<i32> = None;
    let mut in_gs = false;
    let mut current_gs_pos: u32 = 0;
    // Feature #5: pattern and image backgrounds.
    let mut in_patt_fill = false;
    let mut patt_type = String::new();
    let mut patt_fg_color = String::new();
    let mut patt_bg_color = String::new();
    let mut in_patt_fg = false;
    let mut in_patt_bg = false;
    let mut in_blip_fill = false;
    let mut blip_rid: Option<String> = None;

    loop {
        match reader.read_event_into(&mut buffer)? {
            Event::Start(ref event) => {
                let name = event.name();
                let local = local_name(name.as_ref());
                match local {
                    b"bg" if !in_bg => {
                        in_bg = true;
                    }
                    b"bgPr" if in_bg => {
                        in_bg_pr = true;
                    }
                    b"solidFill" if in_bg_pr => {
                        in_solid_fill = true;
                    }
                    b"gradFill" if in_bg_pr => {
                        in_grad_fill = true;
                    }
                    b"gs" if in_grad_fill => {
                        in_gs = true;
                        current_gs_pos = get_attribute_value(event, b"pos")
                            .and_then(|v| v.parse().ok())
                            .unwrap_or(0);
                    }
                    b"srgbClr" if in_solid_fill => {
                        if let Some(val) = get_attribute_value(event, b"val") {
                            return Ok(Some(SlideBackground::Solid(val)));
                        }
                    }
                    b"srgbClr" if in_gs => {
                        if let Some(val) = get_attribute_value(event, b"val") {
                            grad_stops.push(GradientStop {
                                position: current_gs_pos,
                                color_srgb: val,
                                color: None,
                            });
                        }
                    }
                    b"lin" if in_grad_fill => {
                        grad_lin_angle =
                            get_attribute_value(event, b"ang").and_then(|v| v.parse().ok());
                    }
                    // Feature #5: pattern fill background.
                    b"pattFill" if in_bg_pr => {
                        in_patt_fill = true;
                        patt_type = get_attribute_value(event, b"prst").unwrap_or_default();
                    }
                    b"fgClr" if in_patt_fill => {
                        in_patt_fg = true;
                    }
                    b"bgClr" if in_patt_fill => {
                        in_patt_bg = true;
                    }
                    b"srgbClr" if in_patt_fg => {
                        if let Some(val) = get_attribute_value(event, b"val") {
                            patt_fg_color = val;
                        }
                    }
                    b"srgbClr" if in_patt_bg => {
                        if let Some(val) = get_attribute_value(event, b"val") {
                            patt_bg_color = val;
                        }
                    }
                    // Feature #5: image background (blipFill).
                    b"blipFill" if in_bg_pr => {
                        in_blip_fill = true;
                    }
                    b"blip" if in_blip_fill => {
                        blip_rid = get_attribute_value(event, b"r:embed");
                    }
                    _ => {}
                }
            }
            Event::Empty(ref event) => {
                let name = event.name();
                let local = local_name(name.as_ref());
                match local {
                    b"srgbClr" if in_solid_fill => {
                        if let Some(val) = get_attribute_value(event, b"val") {
                            return Ok(Some(SlideBackground::Solid(val)));
                        }
                    }
                    b"srgbClr" if in_gs => {
                        if let Some(val) = get_attribute_value(event, b"val") {
                            grad_stops.push(GradientStop {
                                position: current_gs_pos,
                                color_srgb: val,
                                color: None,
                            });
                        }
                    }
                    b"lin" if in_grad_fill => {
                        grad_lin_angle =
                            get_attribute_value(event, b"ang").and_then(|v| v.parse().ok());
                    }
                    b"srgbClr" if in_patt_fg => {
                        if let Some(val) = get_attribute_value(event, b"val") {
                            patt_fg_color = val;
                        }
                    }
                    b"srgbClr" if in_patt_bg => {
                        if let Some(val) = get_attribute_value(event, b"val") {
                            patt_bg_color = val;
                        }
                    }
                    b"blip" if in_blip_fill => {
                        blip_rid = get_attribute_value(event, b"r:embed");
                    }
                    _ => {}
                }
            }
            Event::End(ref event) => {
                let name = event.name();
                let local = local_name(name.as_ref());
                match local {
                    b"bg" => {
                        if in_grad_fill && !grad_stops.is_empty() {
                            let mut grad = GradientFill::new();
                            grad.fill_type = Some(GradientFillType::Linear);
                            grad.linear_angle = grad_lin_angle;
                            grad.stops = grad_stops;
                            return Ok(Some(SlideBackground::Gradient(grad)));
                        }
                        if in_patt_fill {
                            return Ok(Some(SlideBackground::Pattern {
                                pattern_type: patt_type,
                                foreground_color: patt_fg_color,
                                background_color: patt_bg_color,
                            }));
                        }
                        if let Some(rid) = blip_rid {
                            return Ok(Some(SlideBackground::Image {
                                relationship_id: rid,
                            }));
                        }
                        return Ok(None);
                    }
                    b"solidFill" => in_solid_fill = false,
                    b"gradFill" => {
                        if !grad_stops.is_empty() {
                            let mut grad = GradientFill::new();
                            grad.fill_type = Some(GradientFillType::Linear);
                            grad.linear_angle = grad_lin_angle;
                            grad.stops = grad_stops;
                            return Ok(Some(SlideBackground::Gradient(grad)));
                        }
                        in_grad_fill = false;
                    }
                    b"pattFill" => in_patt_fill = false,
                    b"fgClr" => in_patt_fg = false,
                    b"bgClr" => in_patt_bg = false,
                    b"blipFill" => in_blip_fill = false,
                    b"gs" => in_gs = false,
                    b"bgPr" => in_bg_pr = false,
                    _ => {}
                }
            }
            Event::Eof => break,
            _ => {}
        }
        buffer.clear();
    }

    Ok(None)
}

// ── Feature #8: Parse slide hidden state ──

fn parse_slide_hidden(xml: &[u8]) -> Result<bool> {
    let mut reader = Reader::from_reader(Cursor::new(xml));
    reader.config_mut().trim_text(true);
    let mut buffer = Vec::new();

    loop {
        match reader.read_event_into(&mut buffer)? {
            Event::Start(ref event) | Event::Empty(ref event)
                if local_name(event.name().as_ref()) == b"sld" =>
            {
                let show = get_attribute_value(event, b"show");
                if let Some(show_value) = show {
                    return Ok(parse_xml_bool(show_value.as_str()) == Some(false));
                }
                return Ok(false);
            }
            Event::Eof => break,
            _ => {}
        }
        buffer.clear();
    }

    Ok(false)
}

// ── Feature #12: Parse grouped shapes ──

fn parse_slide_grouped_shapes(xml: &[u8]) -> Result<Vec<ShapeGroup>> {
    let mut reader = Reader::from_reader(Cursor::new(xml));
    reader.config_mut().trim_text(false);
    let mut buffer = Vec::new();
    let mut groups = Vec::new();

    let mut in_grp_sp = false;
    let mut grp_sp_depth = 0_usize;
    let mut current_group_name = String::new();
    let mut in_child_sp = false;
    let mut child_sp_depth = 0_usize;
    let mut child_shape_name = String::new();
    let mut child_shape_paragraphs: Vec<ParsedParagraphData> = Vec::new();
    let mut child_shape_type = ShapeType::AutoShape;
    let mut child_in_tx_body = false;
    let mut child_paragraph_runs: Option<Vec<ParsedRunData>> = None;
    let mut child_paragraph_properties = ParagraphProperties::default();
    let mut child_run_properties = RunProperties::default();
    let mut child_in_rpr = false;
    let mut child_in_text = false;
    let mut child_current_text = String::new();
    let mut child_shapes: Vec<Shape> = Vec::new();

    loop {
        match reader.read_event_into(&mut buffer)? {
            Event::Start(ref event) => {
                let name = event.name();
                let local = local_name(name.as_ref());

                if local == b"grpSp" && !in_grp_sp && !in_child_sp {
                    in_grp_sp = true;
                    grp_sp_depth = 1;
                    current_group_name.clear();
                    child_shapes.clear();
                    buffer.clear();
                    continue;
                }

                if !in_grp_sp {
                    buffer.clear();
                    continue;
                }

                if local == b"sp" && !in_child_sp {
                    in_child_sp = true;
                    child_sp_depth = 1;
                    child_shape_name.clear();
                    child_shape_paragraphs.clear();
                    child_shape_type = ShapeType::AutoShape;
                    child_in_tx_body = false;
                    child_paragraph_runs = None;
                    child_in_rpr = false;
                    child_in_text = false;
                    child_current_text.clear();
                    buffer.clear();
                    continue;
                }

                if in_child_sp {
                    child_sp_depth = child_sp_depth.saturating_add(1);
                    match local {
                        b"cNvPr" => {
                            child_shape_name =
                                get_attribute_value(event, b"name").unwrap_or_default();
                        }
                        b"cNvSpPr" => {
                            child_shape_type = parse_shape_type(event);
                        }
                        b"txBody" => {
                            child_in_tx_body = true;
                        }
                        b"p" if child_in_tx_body => {
                            child_paragraph_runs = Some(Vec::new());
                            child_paragraph_properties = ParagraphProperties::default();
                        }
                        b"rPr" if child_paragraph_runs.is_some() => {
                            child_in_rpr = true;
                            child_run_properties = RunProperties::default();
                            child_run_properties.bold = get_attribute_value(event, b"b")
                                .as_deref()
                                .and_then(parse_xml_bool);
                            child_run_properties.italic = get_attribute_value(event, b"i")
                                .as_deref()
                                .and_then(parse_xml_bool);
                            child_run_properties.font_size =
                                get_attribute_value(event, b"sz").and_then(|v| v.parse().ok());
                        }
                        b"t" if child_paragraph_runs.is_some() => {
                            child_in_text = true;
                            child_current_text.clear();
                        }
                        _ => {}
                    }
                } else {
                    grp_sp_depth = grp_sp_depth.saturating_add(1);
                    if local == b"cNvPr" && grp_sp_depth == 3 {
                        current_group_name = get_attribute_value(event, b"name")
                            .unwrap_or_else(|| format!("Group {}", groups.len() + 1));
                    }
                }
            }
            Event::Empty(ref event) if in_grp_sp => {
                let name = event.name();
                let local = local_name(name.as_ref());
                if in_child_sp {
                    if local == b"cNvPr" {
                        child_shape_name = get_attribute_value(event, b"name").unwrap_or_default();
                    } else if local == b"cNvSpPr" {
                        child_shape_type = parse_shape_type(event);
                    }
                } else if local == b"cNvPr" {
                    current_group_name = get_attribute_value(event, b"name")
                        .unwrap_or_else(|| format!("Group {}", groups.len() + 1));
                }
            }
            Event::Text(ref event) if in_child_sp && child_in_text => {
                let text = event
                    .xml_content()
                    .map_err(quick_xml::Error::from)?
                    .into_owned();
                child_current_text.push_str(text.as_str());
            }
            Event::End(ref event) => {
                if !in_grp_sp {
                    buffer.clear();
                    continue;
                }

                let name = event.name();
                let local = local_name(name.as_ref());

                if in_child_sp {
                    match local {
                        b"t" if child_in_text => {
                            if let Some(runs) = child_paragraph_runs.as_mut() {
                                runs.push(ParsedRunData {
                                    text: std::mem::take(&mut child_current_text),
                                    properties: std::mem::take(&mut child_run_properties),
                                });
                            }
                            child_in_text = false;
                        }
                        b"rPr" if child_in_rpr => {
                            child_in_rpr = false;
                        }
                        b"p" if child_in_tx_body => {
                            if let Some(runs) = child_paragraph_runs.take() {
                                child_shape_paragraphs.push(ParsedParagraphData {
                                    runs,
                                    properties: std::mem::take(&mut child_paragraph_properties),
                                });
                            }
                        }
                        b"txBody" => {
                            child_in_tx_body = false;
                        }
                        _ => {}
                    }

                    child_sp_depth = child_sp_depth.saturating_sub(1);
                    if child_sp_depth == 0 {
                        let shape = build_shape(ParsedShapeData {
                            name: std::mem::take(&mut child_shape_name),
                            paragraphs: std::mem::take(&mut child_shape_paragraphs),
                            placeholder_kind: None,
                            placeholder_idx: None,
                            shape_type: child_shape_type,
                            geometry: None,
                            preset_geometry: None,
                            solid_fill_srgb: None,
                            solid_fill_alpha: None,
                            solid_fill_color: None,
                            outline: None,
                            gradient_fill: None,
                            pattern_fill: None,
                            picture_fill: None,
                            no_fill: false,
                            rotation: None,
                            flip_h: false,
                            flip_v: false,
                            hidden: false,
                            alt_text: None,
                            alt_text_title: None,
                            is_smartart: false,
                            is_connector: false,
                            start_connection: None,
                            end_connection: None,
                            media: None,
                            shadow: None,
                            glow: None,
                            reflection: None,
                            text_anchor: None,
                            auto_fit: None,
                            text_direction: None,
                            text_columns: None,
                            text_column_spacing: None,
                            text_inset_left: None,
                            text_inset_right: None,
                            text_inset_top: None,
                            text_inset_bottom: None,
                            word_wrap: None,
                            body_pr_rot: None,
                            body_pr_rtl_col: None,
                            body_pr_from_word_art: None,
                            body_pr_force_aa: None,
                            body_pr_compat_ln_spc: None,
                            custom_geometry_raw: None,
                            preset_geometry_adjustments: None,
                            unknown_attrs: Vec::new(),
                            unknown_children: Vec::new(),
                            lst_style_raw: None,
                            body_pr_unknown_children: Vec::new(),
                            body_pr_unknown_attrs: Vec::new(),
                        });
                        child_shapes.push(shape);
                        in_child_sp = false;
                    }
                } else {
                    grp_sp_depth = grp_sp_depth.saturating_sub(1);
                    if grp_sp_depth == 0 && local == b"grpSp" {
                        let mut group = ShapeGroup::new(std::mem::take(&mut current_group_name));
                        for shape in std::mem::take(&mut child_shapes) {
                            group.add_shape(shape);
                        }
                        groups.push(group);
                        in_grp_sp = false;
                    }
                }
            }
            Event::Eof => break,
            _ => {}
        }
        buffer.clear();
    }

    Ok(groups)
}

// ── Feature #14: Parse theme from package ──

fn parse_theme_from_package(
    package: &Package,
    presentation_uri: &PartUri,
    presentation_part: &Part,
) -> Result<Option<ThemeColorScheme>> {
    let Some(theme_relationship) = presentation_part
        .relationships
        .get_first_by_type(THEME_RELATIONSHIP_TYPE)
    else {
        return Ok(None);
    };

    if theme_relationship.target_mode != TargetMode::Internal {
        return Ok(None);
    }

    let theme_uri = presentation_uri.resolve_relative(theme_relationship.target.as_str())?;
    let Some(theme_part) = package.get_part(theme_uri.as_str()) else {
        return Ok(None);
    };

    parse_theme_color_scheme(theme_part.data.as_bytes())
}

fn parse_theme_color_scheme(xml: &[u8]) -> Result<Option<ThemeColorScheme>> {
    let mut reader = Reader::from_reader(Cursor::new(xml));
    reader.config_mut().trim_text(true);
    let mut buffer = Vec::new();
    let mut scheme = ThemeColorScheme::new();
    let mut in_clr_scheme = false;
    let mut current_color_name: Option<String> = None;

    loop {
        match reader.read_event_into(&mut buffer)? {
            Event::Start(ref event) => {
                let name = event.name();
                let local = local_name(name.as_ref());
                if local == b"clrScheme" {
                    in_clr_scheme = true;
                    scheme.name = get_attribute_value(event, b"name");
                } else if in_clr_scheme {
                    let local_str = String::from_utf8_lossy(local).into_owned();
                    match local_str.as_str() {
                        "dk1" | "lt1" | "dk2" | "lt2" | "accent1" | "accent2" | "accent3"
                        | "accent4" | "accent5" | "accent6" | "hlink" | "folHlink" => {
                            current_color_name = Some(local_str);
                        }
                        "srgbClr" | "sysClr" => {
                            if let Some(ref color_name) = current_color_name {
                                // For srgbClr, use val; for sysClr, use lastClr first, then val.
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
            }
            Event::Empty(ref event) if in_clr_scheme => {
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
        buffer.clear();
    }

    Ok(None)
}

// ── Feature #6: Parse theme font scheme ──

fn parse_theme_font_scheme_from_package(
    package: &Package,
    presentation_uri: &PartUri,
    presentation_part: &Part,
) -> Result<Option<crate::theme::ThemeFontScheme>> {
    let Some(theme_relationship) = presentation_part
        .relationships
        .get_first_by_type(THEME_RELATIONSHIP_TYPE)
    else {
        return Ok(None);
    };

    if theme_relationship.target_mode != TargetMode::Internal {
        return Ok(None);
    }

    let theme_uri = presentation_uri.resolve_relative(theme_relationship.target.as_str())?;
    let Some(theme_part) = package.get_part(theme_uri.as_str()) else {
        return Ok(None);
    };

    parse_theme_font_scheme(theme_part.data.as_bytes())
}

fn parse_theme_font_scheme(xml: &[u8]) -> Result<Option<crate::theme::ThemeFontScheme>> {
    let mut reader = Reader::from_reader(Cursor::new(xml));
    reader.config_mut().trim_text(true);
    let mut buffer = Vec::new();

    let mut in_font_scheme = false;
    let mut in_major_font = false;
    let mut in_minor_font = false;
    let mut major_latin = String::new();
    let mut minor_latin = String::new();
    let mut major_east_asian: Option<String> = None;
    let mut minor_east_asian: Option<String> = None;

    loop {
        match reader.read_event_into(&mut buffer)? {
            Event::Start(ref event) => {
                let name = event.name();
                let local = local_name(name.as_ref());
                match local {
                    b"fontScheme" => {
                        in_font_scheme = true;
                    }
                    b"majorFont" if in_font_scheme => {
                        in_major_font = true;
                    }
                    b"minorFont" if in_font_scheme => {
                        in_minor_font = true;
                    }
                    b"latin" if in_major_font => {
                        if let Some(typeface) = get_attribute_value(event, b"typeface") {
                            major_latin = typeface;
                        }
                    }
                    b"latin" if in_minor_font => {
                        if let Some(typeface) = get_attribute_value(event, b"typeface") {
                            minor_latin = typeface;
                        }
                    }
                    b"ea" if in_major_font => {
                        major_east_asian = get_attribute_value(event, b"typeface");
                    }
                    b"ea" if in_minor_font => {
                        minor_east_asian = get_attribute_value(event, b"typeface");
                    }
                    _ => {}
                }
            }
            Event::Empty(ref event) => {
                let name = event.name();
                let local = local_name(name.as_ref());
                match local {
                    b"latin" if in_major_font => {
                        if let Some(typeface) = get_attribute_value(event, b"typeface") {
                            major_latin = typeface;
                        }
                    }
                    b"latin" if in_minor_font => {
                        if let Some(typeface) = get_attribute_value(event, b"typeface") {
                            minor_latin = typeface;
                        }
                    }
                    b"ea" if in_major_font => {
                        major_east_asian = get_attribute_value(event, b"typeface");
                    }
                    b"ea" if in_minor_font => {
                        minor_east_asian = get_attribute_value(event, b"typeface");
                    }
                    _ => {}
                }
            }
            Event::End(ref event) => {
                let name = event.name();
                let local = local_name(name.as_ref());
                match local {
                    b"fontScheme" => {
                        if !major_latin.is_empty() || !minor_latin.is_empty() {
                            let mut scheme =
                                crate::theme::ThemeFontScheme::new(major_latin, minor_latin);
                            scheme.major_east_asian = major_east_asian;
                            scheme.minor_east_asian = minor_east_asian;
                            return Ok(Some(scheme));
                        }
                        return Ok(None);
                    }
                    b"majorFont" => in_major_font = false,
                    b"minorFont" => in_minor_font = false,
                    _ => {}
                }
            }
            Event::Eof => break,
            _ => {}
        }
        buffer.clear();
    }

    Ok(None)
}

// ── Feature #13: Parse presentation sections ──

fn parse_presentation_sections(xml: &[u8]) -> Result<Vec<crate::slide::PresentationSection>> {
    use crate::slide::PresentationSection;

    let mut reader = Reader::from_reader(Cursor::new(xml));
    reader.config_mut().trim_text(true);
    let mut buffer = Vec::new();

    let mut sections = Vec::new();
    let mut in_section_lst = false;
    let mut in_section = false;
    let mut current_section_name = String::new();
    let mut current_section_slide_ids: Vec<u32> = Vec::new();

    loop {
        match reader.read_event_into(&mut buffer)? {
            Event::Start(ref event) => {
                let name = event.name();
                let local = local_name(name.as_ref());
                match local {
                    b"sectionLst" => {
                        in_section_lst = true;
                    }
                    b"section" if in_section_lst => {
                        in_section = true;
                        current_section_name =
                            get_attribute_value(event, b"name").unwrap_or_default();
                        current_section_slide_ids.clear();
                    }
                    b"sldId" if in_section => {
                        if let Some(id) =
                            get_attribute_value(event, b"id").and_then(|v| v.parse().ok())
                        {
                            current_section_slide_ids.push(id);
                        }
                    }
                    _ => {}
                }
            }
            Event::Empty(ref event) => {
                let name = event.name();
                let local = local_name(name.as_ref());
                match local {
                    b"sldId" if in_section => {
                        if let Some(id) =
                            get_attribute_value(event, b"id").and_then(|v| v.parse().ok())
                        {
                            current_section_slide_ids.push(id);
                        }
                    }
                    _ => {}
                }
            }
            Event::End(ref event) => {
                let name = event.name();
                let local = local_name(name.as_ref());
                match local {
                    b"section" if in_section => {
                        let mut section =
                            PresentationSection::new(std::mem::take(&mut current_section_name));
                        section.set_slide_ids(std::mem::take(&mut current_section_slide_ids));
                        sections.push(section);
                        in_section = false;
                    }
                    b"sectionLst" => {
                        in_section_lst = false;
                    }
                    _ => {}
                }
            }
            Event::Eof => break,
            _ => {}
        }
        buffer.clear();
    }

    Ok(sections)
}

// ── Feature #15: Parse header/footer ──

fn parse_slide_header_footer(xml: &[u8]) -> Result<Option<SlideHeaderFooter>> {
    let mut reader = Reader::from_reader(Cursor::new(xml));
    reader.config_mut().trim_text(true);
    let mut buffer = Vec::new();

    loop {
        match reader.read_event_into(&mut buffer)? {
            Event::Empty(ref event) | Event::Start(ref event)
                if local_name(event.name().as_ref()) == b"hf" =>
            {
                let hf = SlideHeaderFooter {
                    show_slide_number: get_attribute_value(event, b"sldNum")
                        .as_deref()
                        .and_then(parse_xml_bool),
                    show_date_time: get_attribute_value(event, b"dt")
                        .as_deref()
                        .and_then(parse_xml_bool),
                    show_header: get_attribute_value(event, b"hdr")
                        .as_deref()
                        .and_then(parse_xml_bool),
                    show_footer: get_attribute_value(event, b"ftr")
                        .as_deref()
                        .and_then(parse_xml_bool),
                    footer_text: None,
                    date_time_text: None,
                };
                return Ok(Some(hf));
            }
            Event::Eof => break,
            _ => {}
        }
        buffer.clear();
    }

    Ok(None)
}

fn transition_kind_from_xml(local_name: &[u8]) -> SlideTransitionKind {
    match local_name {
        b"cut" => SlideTransitionKind::Cut,
        b"fade" => SlideTransitionKind::Fade,
        b"push" => SlideTransitionKind::Push,
        b"wipe" => SlideTransitionKind::Wipe,
        _ => SlideTransitionKind::Other(String::from_utf8_lossy(local_name).into_owned()),
    }
}

fn update_transition_kind_from_xml(current: &mut SlideTransitionKind, local_name: &[u8]) {
    let candidate = transition_kind_from_xml(local_name);
    if is_known_transition_kind(&candidate) || matches!(current, SlideTransitionKind::Unspecified) {
        *current = candidate;
    }
}

fn is_known_transition_kind(kind: &SlideTransitionKind) -> bool {
    matches!(
        kind,
        SlideTransitionKind::Cut
            | SlideTransitionKind::Fade
            | SlideTransitionKind::Push
            | SlideTransitionKind::Wipe
    )
}

/// Resolve hyperlink relationship IDs in shapes to actual URLs.
/// Walks through all shapes and their text runs, and when a run has a
/// `hyperlink_click_rid`, looks up the relationship target URL in the slide's
/// relationships and stores it on the run properties.
fn resolve_hyperlink_urls(shapes: &mut [Shape], slide_part: &Part) {
    for shape in shapes.iter_mut() {
        for paragraph in shape.paragraphs_mut() {
            for run in paragraph.runs_mut() {
                if let Some(rid) = run.properties().hyperlink_click_rid.clone() {
                    if let Some(relationship) = slide_part.relationships.get_by_id(&rid) {
                        run.properties_mut().hyperlink_url = Some(relationship.target.clone());
                    }
                }
            }
        }
    }
}

struct ParsedRunData {
    text: String,
    properties: RunProperties,
}

struct ParsedParagraphData {
    runs: Vec<ParsedRunData>,
    properties: ParagraphProperties,
}

struct ParsedShapeData {
    name: String,
    paragraphs: Vec<ParsedParagraphData>,
    placeholder_kind: Option<String>,
    placeholder_idx: Option<u32>,
    shape_type: ShapeType,
    geometry: Option<ShapeGeometry>,
    preset_geometry: Option<String>,
    solid_fill_srgb: Option<String>,
    solid_fill_alpha: Option<u8>,
    solid_fill_color: Option<ShapeColor>,
    outline: Option<ShapeOutline>,
    gradient_fill: Option<GradientFill>,
    pattern_fill: Option<PatternFill>,
    picture_fill: Option<PictureFill>,
    no_fill: bool,
    rotation: Option<i32>,
    flip_h: bool,
    flip_v: bool,
    hidden: bool,
    alt_text: Option<String>,
    alt_text_title: Option<String>,
    is_smartart: bool,
    is_connector: bool,
    start_connection: Option<ConnectionInfo>,
    end_connection: Option<ConnectionInfo>,
    media: Option<(MediaType, String)>,
    shadow: Option<ShapeShadow>,
    glow: Option<ShapeGlow>,
    reflection: Option<ShapeReflection>,
    text_anchor: Option<TextAnchor>,
    auto_fit: Option<AutoFitType>,
    text_direction: Option<TextDirection>,
    text_columns: Option<u32>,
    text_column_spacing: Option<i64>,
    text_inset_left: Option<i64>,
    text_inset_right: Option<i64>,
    text_inset_top: Option<i64>,
    text_inset_bottom: Option<i64>,
    word_wrap: Option<bool>,
    body_pr_rot: Option<i32>,
    body_pr_rtl_col: Option<bool>,
    body_pr_from_word_art: Option<bool>,
    body_pr_force_aa: Option<bool>,
    body_pr_compat_ln_spc: Option<bool>,
    custom_geometry_raw: Option<RawXmlNode>,
    preset_geometry_adjustments: Option<RawXmlNode>,
    unknown_attrs: Vec<(String, String)>,
    unknown_children: Vec<RawXmlNode>,
    lst_style_raw: Option<RawXmlNode>,
    body_pr_unknown_children: Vec<RawXmlNode>,
    body_pr_unknown_attrs: Vec<(String, String)>,
}

fn build_shape(parsed: ParsedShapeData) -> Shape {
    let mut shape = Shape::new(parsed.name);
    shape.set_shape_type(parsed.shape_type);

    if let Some(ref placeholder_kind) = parsed.placeholder_kind {
        shape.set_placeholder_kind(placeholder_kind.clone());
        shape.set_placeholder_type(PlaceholderType::from_xml(placeholder_kind));
    }
    if let Some(placeholder_idx) = parsed.placeholder_idx {
        shape.set_placeholder_idx(placeholder_idx);
    }
    if let Some(geometry) = parsed.geometry {
        shape.set_geometry(geometry);
    }
    if let Some(preset_geometry) = parsed.preset_geometry {
        shape.set_preset_geometry(preset_geometry);
    }
    if let Some(solid_fill_srgb) = parsed.solid_fill_srgb {
        shape.set_solid_fill_srgb(solid_fill_srgb);
    }
    if let Some(alpha) = parsed.solid_fill_alpha {
        shape.set_solid_fill_alpha(alpha);
    }
    if let Some(color) = parsed.solid_fill_color {
        shape.set_solid_fill_color(color);
    }
    if let Some(outline) = parsed.outline {
        shape.set_outline(outline);
    }
    if let Some(gradient_fill) = parsed.gradient_fill {
        shape.set_gradient_fill(gradient_fill);
    }
    if let Some(pattern_fill) = parsed.pattern_fill {
        shape.set_pattern_fill(pattern_fill);
    }
    if let Some(picture_fill) = parsed.picture_fill {
        shape.set_picture_fill(picture_fill);
    }
    if parsed.no_fill {
        shape.set_no_fill(true);
    }
    if let Some(rotation) = parsed.rotation {
        shape.set_rotation(rotation);
    }
    if parsed.flip_h {
        shape.set_flip_h(true);
    }
    if parsed.flip_v {
        shape.set_flip_v(true);
    }
    if parsed.hidden {
        shape.set_hidden(true);
    }
    if let Some(alt_text) = parsed.alt_text {
        shape.set_alt_text(alt_text);
    }
    if let Some(alt_text_title) = parsed.alt_text_title {
        shape.set_alt_text_title(alt_text_title);
    }
    if parsed.is_smartart {
        shape.set_smartart(true);
    }
    if parsed.is_connector {
        shape.set_connector(true);
    }
    if let Some(start_conn) = parsed.start_connection {
        shape.set_start_connection(start_conn);
    }
    if let Some(end_conn) = parsed.end_connection {
        shape.set_end_connection(end_conn);
    }
    if let Some((media_type, rid)) = parsed.media {
        shape.set_media(media_type, rid);
    }
    if let Some(shadow) = parsed.shadow {
        shape.set_shadow(shadow);
    }
    if let Some(glow) = parsed.glow {
        shape.set_glow(glow);
    }
    if let Some(reflection) = parsed.reflection {
        shape.set_reflection(reflection);
    }
    if let Some(text_anchor) = parsed.text_anchor {
        shape.set_text_anchor(text_anchor);
    }
    if let Some(auto_fit) = parsed.auto_fit {
        shape.set_auto_fit(auto_fit);
    }
    if let Some(text_direction) = parsed.text_direction {
        shape.set_text_direction(text_direction);
    }
    if let Some(text_columns) = parsed.text_columns {
        shape.set_text_columns(text_columns);
    }
    if let Some(text_column_spacing) = parsed.text_column_spacing {
        shape.set_text_column_spacing(text_column_spacing);
    }
    if let Some(inset) = parsed.text_inset_left {
        shape.set_text_inset_left(inset);
    }
    if let Some(inset) = parsed.text_inset_right {
        shape.set_text_inset_right(inset);
    }
    if let Some(inset) = parsed.text_inset_top {
        shape.set_text_inset_top(inset);
    }
    if let Some(inset) = parsed.text_inset_bottom {
        shape.set_text_inset_bottom(inset);
    }
    if let Some(word_wrap) = parsed.word_wrap {
        shape.set_word_wrap(word_wrap);
    }
    if let Some(rot) = parsed.body_pr_rot {
        shape.set_body_pr_rot(rot);
    }
    if let Some(rtl) = parsed.body_pr_rtl_col {
        shape.set_body_pr_rtl_col(rtl);
    }
    if let Some(fwa) = parsed.body_pr_from_word_art {
        shape.set_body_pr_from_word_art(fwa);
    }
    if let Some(faa) = parsed.body_pr_force_aa {
        shape.set_body_pr_force_aa(faa);
    }
    if let Some(cls) = parsed.body_pr_compat_ln_spc {
        shape.set_body_pr_compat_ln_spc(cls);
    }
    if let Some(custom_geom) = parsed.custom_geometry_raw {
        shape.set_custom_geometry_raw(custom_geom);
    }
    if let Some(prst_adj) = parsed.preset_geometry_adjustments {
        shape.set_preset_geometry_adjustments(prst_adj);
    }
    if !parsed.unknown_attrs.is_empty() {
        shape.set_unknown_attrs(parsed.unknown_attrs);
    }
    for node in parsed.unknown_children {
        shape.push_unknown_child(node);
    }
    if let Some(lst_style) = parsed.lst_style_raw {
        shape.set_lst_style_raw(lst_style);
    }
    for node in parsed.body_pr_unknown_children {
        shape.push_body_pr_unknown_child(node);
    }
    if !parsed.body_pr_unknown_attrs.is_empty() {
        shape.set_body_pr_unknown_attrs(parsed.body_pr_unknown_attrs);
    }

    for parsed_paragraph in parsed.paragraphs {
        let paragraph = shape.add_paragraph();
        paragraph.set_properties(parsed_paragraph.properties);
        for parsed_run in parsed_paragraph.runs {
            let run = paragraph.add_run(parsed_run.text);
            run.set_properties(parsed_run.properties);
        }
    }

    shape
}

#[allow(dead_code)]
fn build_table_from_rows(rows: Vec<Vec<String>>) -> Table {
    let row_count = rows.len();
    let col_count = rows.iter().map(Vec::len).max().unwrap_or(0);
    let mut table = Table::new(row_count, col_count);

    for (row_index, row) in rows.iter().enumerate() {
        for (col_index, cell_text) in row.iter().enumerate() {
            let _ = table.set_cell_text(row_index, col_index, cell_text.clone());
        }
    }

    table
}

/// Parsed cell data from table XML. Defined here because `parse_slide_tables`
/// defines a local struct with the same layout; this version is used by the builder.
struct ParsedTableCellData {
    text: String,
    fill_color_srgb: Option<String>,
    fill_color: Option<ShapeColor>,
    borders: CellBorders,
    bold: Option<bool>,
    italic: Option<bool>,
    font_size: Option<u32>,
    font_color_srgb: Option<String>,
    font_color: Option<ShapeColor>,
    grid_span: Option<u32>,
    row_span: Option<u32>,
    v_merge: bool,
    vertical_alignment: Option<CellTextAnchor>,
    margin_left: Option<i64>,
    margin_right: Option<i64>,
    margin_top: Option<i64>,
    margin_bottom: Option<i64>,
    text_direction: Option<TextDirection>,
}

fn build_table_from_parsed_cells(
    rows: Vec<Vec<ParsedTableCellData>>,
    column_widths: Vec<i64>,
    row_heights: Vec<i64>,
) -> Table {
    let row_count = rows.len();
    let col_count = rows.iter().map(Vec::len).max().unwrap_or(0);
    let mut table = Table::new(row_count, col_count);

    for (row_index, row) in rows.into_iter().enumerate() {
        for (col_index, cell_data) in row.into_iter().enumerate() {
            if let Some(cell) = table.cell_mut(row_index, col_index) {
                cell.set_text(cell_data.text);
                if let Some(color) = cell_data.fill_color_srgb {
                    cell.set_fill_color_srgb(color);
                }
                if let Some(color) = cell_data.fill_color {
                    cell.set_fill_color(color);
                }
                if cell_data.borders.is_set() {
                    *cell.borders_mut() = cell_data.borders;
                }
                if let Some(bold) = cell_data.bold {
                    cell.set_bold(bold);
                }
                if let Some(italic) = cell_data.italic {
                    cell.set_italic(italic);
                }
                if let Some(size) = cell_data.font_size {
                    cell.set_font_size(size);
                }
                if let Some(color) = cell_data.font_color_srgb {
                    cell.set_font_color_srgb(color);
                }
                if let Some(color) = cell_data.font_color {
                    cell.set_font_color(color);
                }
                if let Some(gs) = cell_data.grid_span {
                    cell.set_grid_span(gs);
                }
                if let Some(rs) = cell_data.row_span {
                    cell.set_row_span(rs);
                }
                if cell_data.v_merge {
                    cell.set_v_merge(true);
                }
                if let Some(va) = cell_data.vertical_alignment {
                    cell.set_vertical_alignment(va);
                }
                if let Some(m) = cell_data.margin_left {
                    cell.set_margin_left(m);
                }
                if let Some(m) = cell_data.margin_right {
                    cell.set_margin_right(m);
                }
                if let Some(m) = cell_data.margin_top {
                    cell.set_margin_top(m);
                }
                if let Some(m) = cell_data.margin_bottom {
                    cell.set_margin_bottom(m);
                }
                if let Some(td) = cell_data.text_direction {
                    cell.set_text_direction(td);
                }
            }
        }
    }

    if !column_widths.is_empty() {
        table.set_column_widths(column_widths);
    }
    if !row_heights.is_empty() {
        table.set_row_heights(row_heights);
    }

    table
}

fn extract_legacy_text_runs(shape: Option<&Shape>) -> Vec<String> {
    let Some(shape) = shape else {
        return Vec::new();
    };
    let Some(paragraph) = shape.paragraphs().first() else {
        return Vec::new();
    };

    paragraph
        .runs()
        .iter()
        .map(|run| run.text().to_string())
        .collect()
}

fn append_shape(slide: &mut Slide, shape: Shape) {
    let target_shape = slide.add_shape(shape.name().to_string());
    target_shape.set_shape_type(shape.shape_type());
    if let Some(placeholder_kind) = shape.placeholder_kind() {
        target_shape.set_placeholder_kind(placeholder_kind.to_string());
    }
    if let Some(placeholder_type) = shape.placeholder_type() {
        target_shape.set_placeholder_type(placeholder_type.clone());
    }
    if let Some(placeholder_idx) = shape.placeholder_idx() {
        target_shape.set_placeholder_idx(placeholder_idx);
    }
    if let Some(geometry) = shape.geometry() {
        target_shape.set_geometry(geometry);
    }
    if let Some(preset_geometry) = shape.preset_geometry() {
        target_shape.set_preset_geometry(preset_geometry.to_string());
    }
    if let Some(solid_fill_srgb) = shape.solid_fill_srgb() {
        target_shape.set_solid_fill_srgb(solid_fill_srgb.to_string());
    }
    if let Some(alpha) = shape.solid_fill_alpha() {
        target_shape.set_solid_fill_alpha(alpha);
    }
    if let Some(outline) = shape.outline() {
        target_shape.set_outline(outline.clone());
    }
    if let Some(gradient_fill) = shape.gradient_fill() {
        target_shape.set_gradient_fill(gradient_fill.clone());
    }
    if let Some(pattern_fill) = shape.pattern_fill() {
        target_shape.set_pattern_fill(pattern_fill.clone());
    }
    if let Some(picture_fill) = shape.picture_fill() {
        target_shape.set_picture_fill(picture_fill.clone());
    }
    if shape.is_no_fill() {
        target_shape.set_no_fill(true);
    }
    if let Some(rotation) = shape.rotation() {
        target_shape.set_rotation(rotation);
    }
    if shape.is_hidden() {
        target_shape.set_hidden(true);
    }
    if let Some(alt_text) = shape.alt_text() {
        target_shape.set_alt_text(alt_text.to_string());
    }
    if let Some(alt_text_title) = shape.alt_text_title() {
        target_shape.set_alt_text_title(alt_text_title.to_string());
    }
    if shape.is_smartart() {
        target_shape.set_smartart(true);
    }
    if shape.is_connector() {
        target_shape.set_connector(true);
    }
    if let Some(conn) = shape.start_connection() {
        target_shape.set_start_connection(conn.clone());
    }
    if let Some(conn) = shape.end_connection() {
        target_shape.set_end_connection(conn.clone());
    }
    if let Some((media_type, rid)) = shape.media() {
        target_shape.set_media(*media_type, rid.to_string());
    }
    target_shape.set_unknown_attrs(shape.unknown_attrs().to_vec());
    for node in shape.unknown_children() {
        target_shape.push_unknown_child(node.clone());
    }

    for paragraph in shape.paragraphs() {
        let target_paragraph = target_shape.add_paragraph();
        target_paragraph.set_properties(paragraph.properties().clone());
        for run in paragraph.runs() {
            let target_run = target_paragraph.add_run(run.text().to_string());
            target_run.set_properties(run.properties().clone());
        }
    }
}

fn append_table(slide: &mut Slide, table: Table) {
    let target_table = slide.add_table(table.rows(), table.cols());

    // Feature #6: Copy column widths and row heights.
    for (col_idx, &width) in table.column_widths_emu().iter().enumerate() {
        target_table.set_column_width_emu(col_idx, width);
    }
    for (row_idx, &height) in table.row_heights_emu().iter().enumerate() {
        target_table.set_row_height_emu(row_idx, height);
    }

    for row_index in 0..table.rows() {
        for col_index in 0..table.cols() {
            if let Some(source_cell) = table.cell(row_index, col_index) {
                if let Some(target_cell) = target_table.cell_mut(row_index, col_index) {
                    target_cell.set_text(source_cell.text().to_string());

                    // Feature #5: Copy cell formatting.
                    if let Some(color) = source_cell.fill_color_srgb() {
                        target_cell.set_fill_color_srgb(color.to_string());
                    }
                    if let Some(bold) = source_cell.bold() {
                        target_cell.set_bold(bold);
                    }
                    if let Some(italic) = source_cell.italic() {
                        target_cell.set_italic(italic);
                    }
                    if let Some(size) = source_cell.font_size() {
                        target_cell.set_font_size(size);
                    }
                    if let Some(color) = source_cell.font_color_srgb() {
                        target_cell.set_font_color_srgb(color.to_string());
                    }
                    if source_cell.borders().is_set() {
                        *target_cell.borders_mut() = source_cell.borders().clone();
                    }
                    // Feature #3 (table merged cells): copy merge data.
                    if let Some(gs) = source_cell.grid_span() {
                        target_cell.set_grid_span(gs);
                    }
                    if let Some(rs) = source_cell.row_span() {
                        target_cell.set_row_span(rs);
                    }
                    if source_cell.is_v_merge() {
                        target_cell.set_v_merge(true);
                    }
                }
            }
        }
    }
}

fn resolve_slide_image(
    package: &Package,
    slide_uri: &PartUri,
    slide_part: &Part,
    parsed_image: ParsedImageRef,
) -> Result<Option<Image>> {
    let Some(relationship) = slide_part
        .relationships
        .get_by_id(parsed_image.relationship_id.as_str())
    else {
        tracing::warn!(
            relationship_id = parsed_image.relationship_id.as_str(),
            "missing image relationship; skipping image reference"
        );
        return Ok(None);
    };

    if relationship.target_mode != TargetMode::Internal {
        tracing::warn!(
            relationship_id = relationship.id.as_str(),
            "image relationship is external; skipping image reference"
        );
        return Ok(None);
    }

    let image_uri = slide_uri.resolve_relative(relationship.target.as_str())?;
    let Some(image_part) = package.get_part(image_uri.as_str()) else {
        tracing::warn!(
            relationship_id = relationship.id.as_str(),
            part_uri = image_uri.as_str(),
            "missing image part for relationship; skipping image reference"
        );
        return Ok(None);
    };

    let content_type = image_part
        .content_type
        .clone()
        .unwrap_or_else(|| fallback_content_type_for_extension(image_uri.extension()).to_string());

    let mut image = Image::new(image_part.data.as_bytes().to_vec(), content_type);
    image.set_name(parsed_image.name);
    image.set_relationship_id(Some(parsed_image.relationship_id));
    Ok(Some(image))
}

fn resolve_slide_chart(
    package: &Package,
    slide_uri: &PartUri,
    slide_part: &Part,
    parsed_chart: ParsedChartRef,
) -> Result<Option<Chart>> {
    let Some(relationship) = slide_part
        .relationships
        .get_by_id(parsed_chart.relationship_id.as_str())
    else {
        tracing::warn!(
            relationship_id = parsed_chart.relationship_id.as_str(),
            "missing chart relationship; skipping chart reference"
        );
        return Ok(None);
    };

    if relationship.target_mode != TargetMode::Internal {
        tracing::warn!(
            relationship_id = relationship.id.as_str(),
            "chart relationship is external; skipping chart reference"
        );
        return Ok(None);
    }

    let chart_uri = slide_uri.resolve_relative(relationship.target.as_str())?;
    let Some(chart_part) = package.get_part(chart_uri.as_str()) else {
        tracing::warn!(
            relationship_id = relationship.id.as_str(),
            part_uri = chart_uri.as_str(),
            "missing chart part for relationship; skipping chart reference"
        );
        return Ok(None);
    };

    Ok(Some(parse_chart_xml(chart_part.data.as_bytes())?))
}

fn parse_chart_xml(xml: &[u8]) -> Result<Chart> {
    #[derive(Clone, Copy, Debug, PartialEq, Eq)]
    enum ChartTextTarget {
        None,
        Title,
        Category,
        Value,
        SeriesName,
    }

    let mut reader = Reader::from_reader(Cursor::new(xml));
    reader.config_mut().trim_text(false);

    let mut buffer = Vec::new();
    let mut chart = Chart::new("");

    let mut in_title = false;
    let mut title_depth = 0_usize;
    let mut in_category_container = false;
    let mut category_depth = 0_usize;
    let mut in_value_container = false;
    let mut value_depth = 0_usize;
    let mut in_series_tx = false;
    let mut series_tx_depth = 0_usize;

    let mut text_target = ChartTextTarget::None;
    let mut current_text = String::new();
    let mut title_chunks: Vec<String> = Vec::new();
    let mut categories: Vec<String> = Vec::new();
    let mut values: Vec<f64> = Vec::new();

    // Feature #13: Chart type detection.
    let mut detected_chart_type: Option<ChartType> = None;
    let mut bar_dir_horizontal = false;
    let mut in_plot_area = false;

    // Feature #13: Legend.
    let mut in_legend = false;
    let mut legend_pos: Option<LegendPosition> = None;
    let mut has_legend = false;

    // Feature #13: Multi-series tracking.
    let mut in_series = false;
    let mut _series_index = 0_u32;
    let mut current_series_name = String::new();
    let mut current_series_values: Vec<f64> = Vec::new();
    let mut all_series: Vec<(String, Vec<f64>)> = Vec::new();

    // Chart-type-specific properties.
    let mut bar_gap_width: Option<u32> = None;
    let mut bar_overlap: Option<i32> = None;
    let mut pie_first_slice_angle: Option<u32> = None;
    let mut pie_hole_size: Option<u32> = None;
    let mut scatter_style: Option<crate::chart::ScatterStyle> = None;
    let mut bubble_scale: Option<u32> = None;

    // Feature #4 (chart axes): axis parsing state.
    let mut in_cat_ax = false;
    let mut in_val_ax = false;
    let mut cat_ax_title: Option<String> = None;
    let mut val_ax_title: Option<String> = None;
    let mut cat_ax_has_major_gridlines = false;
    let mut val_ax_has_major_gridlines = false;
    let mut val_ax_min: Option<f64> = None;
    let mut val_ax_max: Option<f64> = None;
    let mut val_ax_major_unit: Option<f64> = None;
    let mut in_axis_title = false;
    let mut axis_title_depth = 0_usize;
    let mut in_axis_title_text = false;
    let mut axis_title_text = String::new();

    loop {
        match reader.read_event_into(&mut buffer)? {
            Event::Start(ref event) => {
                let event_name = event.name();
                let local = local_name(event_name.as_ref());

                if in_title {
                    title_depth = title_depth.saturating_add(1);
                }
                if in_category_container {
                    category_depth = category_depth.saturating_add(1);
                }
                if in_value_container {
                    value_depth = value_depth.saturating_add(1);
                }
                if in_series_tx {
                    series_tx_depth = series_tx_depth.saturating_add(1);
                }

                let local_str = std::str::from_utf8(local).unwrap_or("");
                match local {
                    b"plotArea" => {
                        in_plot_area = true;
                    }
                    // Detect chart type from plot area children.
                    b"barChart" | b"bar3DChart" | b"lineChart" | b"line3DChart" | b"pieChart"
                    | b"pie3DChart" | b"areaChart" | b"area3DChart" | b"scatterChart"
                    | b"doughnutChart" | b"radarChart"
                        if in_plot_area && detected_chart_type.is_none() =>
                    {
                        detected_chart_type = ChartType::from_xml_element(local_str);
                    }
                    b"title" if !in_title && !in_series => {
                        in_title = true;
                        title_depth = 1;
                    }
                    b"ser" if in_plot_area => {
                        in_series = true;
                        current_series_name.clear();
                        current_series_values.clear();
                    }
                    b"tx" if in_series && !in_title => {
                        in_series_tx = true;
                        series_tx_depth = 1;
                    }
                    b"cat" if !in_category_container => {
                        in_category_container = true;
                        category_depth = 1;
                    }
                    b"val" if in_series && !in_value_container => {
                        in_value_container = true;
                        value_depth = 1;
                    }
                    b"legend" => {
                        in_legend = true;
                        has_legend = true;
                    }
                    // Feature #4 (chart axes): detect axis elements.
                    b"catAx" if in_plot_area => {
                        in_cat_ax = true;
                    }
                    b"valAx" if in_plot_area => {
                        in_val_ax = true;
                    }
                    b"title" if (in_cat_ax || in_val_ax) && !in_axis_title => {
                        in_axis_title = true;
                        axis_title_depth = 1;
                        axis_title_text.clear();
                    }
                    b"t" if in_axis_title => {
                        in_axis_title_text = true;
                        current_text.clear();
                    }
                    b"majorGridlines" if in_cat_ax => {
                        cat_ax_has_major_gridlines = true;
                    }
                    b"majorGridlines" if in_val_ax => {
                        val_ax_has_major_gridlines = true;
                    }
                    b"t" if in_title => {
                        text_target = ChartTextTarget::Title;
                        current_text.clear();
                    }
                    b"v" if in_series_tx => {
                        text_target = ChartTextTarget::SeriesName;
                        current_text.clear();
                    }
                    b"v" if in_category_container => {
                        text_target = ChartTextTarget::Category;
                        current_text.clear();
                    }
                    b"v" if in_value_container => {
                        text_target = ChartTextTarget::Value;
                        current_text.clear();
                    }
                    _ => {}
                }
            }
            Event::Empty(ref event) => {
                let event_name = event.name();
                let local = local_name(event_name.as_ref());
                match local {
                    b"t" if in_title => {
                        title_chunks.push(String::new());
                    }
                    b"v" if in_category_container => {
                        categories.push(String::new());
                    }
                    b"v" if in_value_container => {
                        if in_series {
                            current_series_values.push(0.0);
                        } else {
                            values.push(0.0);
                        }
                    }
                    b"barDir" => {
                        let dir = get_attribute_value(event, b"val");
                        if dir.as_deref() == Some("bar") {
                            bar_dir_horizontal = true;
                        }
                    }
                    b"legendPos" if in_legend => {
                        legend_pos = get_attribute_value(event, b"val")
                            .and_then(|v| LegendPosition::from_xml(v.as_str()));
                    }
                    // Feature #4 (chart axes): scaling properties.
                    b"min" if in_val_ax => {
                        val_ax_min =
                            get_attribute_value(event, b"val").and_then(|v| v.parse().ok());
                    }
                    b"max" if in_val_ax => {
                        val_ax_max =
                            get_attribute_value(event, b"val").and_then(|v| v.parse().ok());
                    }
                    b"majorUnit" if in_val_ax => {
                        val_ax_major_unit =
                            get_attribute_value(event, b"val").and_then(|v| v.parse().ok());
                    }
                    b"majorGridlines" if in_cat_ax => {
                        cat_ax_has_major_gridlines = true;
                    }
                    b"majorGridlines" if in_val_ax => {
                        val_ax_has_major_gridlines = true;
                    }
                    // Chart-type-specific properties.
                    b"gapWidth" if in_plot_area => {
                        bar_gap_width =
                            get_attribute_value(event, b"val").and_then(|v| v.parse().ok());
                    }
                    b"overlap" if in_plot_area => {
                        bar_overlap =
                            get_attribute_value(event, b"val").and_then(|v| v.parse().ok());
                    }
                    b"firstSliceAng" if in_plot_area => {
                        pie_first_slice_angle =
                            get_attribute_value(event, b"val").and_then(|v| v.parse().ok());
                    }
                    b"holeSize" if in_plot_area => {
                        pie_hole_size =
                            get_attribute_value(event, b"val").and_then(|v| v.parse().ok());
                    }
                    b"scatterStyle" if in_plot_area => {
                        scatter_style = get_attribute_value(event, b"val")
                            .and_then(|v| crate::chart::ScatterStyle::from_xml(&v));
                    }
                    b"bubbleScale" if in_plot_area => {
                        bubble_scale =
                            get_attribute_value(event, b"val").and_then(|v| v.parse().ok());
                    }
                    _ => {}
                }
            }
            Event::Text(ref event) if text_target != ChartTextTarget::None => {
                let text = event
                    .xml_content()
                    .map_err(quick_xml::Error::from)?
                    .into_owned();
                current_text.push_str(text.as_str());
            }
            Event::CData(ref event) if text_target != ChartTextTarget::None => {
                let text = String::from_utf8_lossy(event.as_ref());
                current_text.push_str(text.as_ref());
            }
            Event::End(ref event) => {
                let event_name = event.name();
                let local = local_name(event_name.as_ref());

                match (local, text_target) {
                    (b"t", ChartTextTarget::Title) => {
                        title_chunks.push(std::mem::take(&mut current_text));
                        text_target = ChartTextTarget::None;
                    }
                    (b"v", ChartTextTarget::SeriesName) => {
                        current_series_name = std::mem::take(&mut current_text);
                        text_target = ChartTextTarget::None;
                    }
                    (b"v", ChartTextTarget::Category) => {
                        categories.push(std::mem::take(&mut current_text));
                        text_target = ChartTextTarget::None;
                    }
                    (b"v", ChartTextTarget::Value) => {
                        let numeric_value = current_text.parse::<f64>().unwrap_or(0.0);
                        if in_series {
                            current_series_values.push(numeric_value);
                        } else {
                            values.push(numeric_value);
                        }
                        current_text.clear();
                        text_target = ChartTextTarget::None;
                    }
                    _ => {}
                }

                if in_title {
                    title_depth = title_depth.saturating_sub(1);
                    if title_depth == 0 && local == b"title" {
                        in_title = false;
                    }
                }

                if in_series_tx {
                    series_tx_depth = series_tx_depth.saturating_sub(1);
                    if series_tx_depth == 0 && local == b"tx" {
                        in_series_tx = false;
                    }
                }

                if in_category_container {
                    category_depth = category_depth.saturating_sub(1);
                    if category_depth == 0 && local == b"cat" {
                        in_category_container = false;
                    }
                }

                if in_value_container {
                    value_depth = value_depth.saturating_sub(1);
                    if value_depth == 0 && local == b"val" {
                        in_value_container = false;
                    }
                }

                // Feature #4 (chart axes): handle axis title text end.
                if in_axis_title {
                    if local == b"t" && in_axis_title_text {
                        axis_title_text.push_str(&current_text);
                        current_text.clear();
                        in_axis_title_text = false;
                    }
                    axis_title_depth = axis_title_depth.saturating_sub(1);
                    if axis_title_depth == 0 && local == b"title" {
                        in_axis_title = false;
                        if !axis_title_text.is_empty() {
                            if in_cat_ax {
                                cat_ax_title = Some(std::mem::take(&mut axis_title_text));
                            } else if in_val_ax {
                                val_ax_title = Some(std::mem::take(&mut axis_title_text));
                            }
                        }
                    }
                }

                match local {
                    b"ser" if in_series => {
                        all_series.push((
                            std::mem::take(&mut current_series_name),
                            std::mem::take(&mut current_series_values),
                        ));
                        _series_index += 1;
                        in_series = false;
                    }
                    b"legend" if in_legend => {
                        in_legend = false;
                    }
                    b"plotArea" if in_plot_area => {
                        in_plot_area = false;
                    }
                    b"catAx" if in_cat_ax => {
                        in_cat_ax = false;
                    }
                    b"valAx" if in_val_ax => {
                        in_val_ax = false;
                    }
                    _ => {}
                }
            }
            Event::Eof => break,
            _ => {}
        }
        buffer.clear();
    }

    chart.set_title(title_chunks.join(""));

    // Feature #13: Chart type.
    if let Some(ct) = detected_chart_type {
        // If bar direction is horizontal, override to Bar type.
        if bar_dir_horizontal {
            chart.set_chart_type(ChartType::Bar);
            chart.set_bar_direction_horizontal(true);
        } else {
            chart.set_chart_type(ct);
        }
    }

    // Set categories/values from the first series. Additional series become
    // `additional_series` on the Chart.
    if !all_series.is_empty() {
        let (_first_name, first_values) = all_series.remove(0);
        chart.set_data_points(categories, first_values);

        for (name, vals) in all_series {
            let mut series = ChartSeries::new(name);
            series.set_values(vals);
            chart.add_series(series);
        }
    } else {
        chart.set_data_points(categories, values);
    }

    // Feature #13: Legend.
    if has_legend {
        chart.set_show_legend(true);
        if let Some(pos) = legend_pos {
            chart.set_legend_position(pos);
        }
    }

    // Feature #4: Chart axes.
    use crate::chart::ChartAxis;
    if cat_ax_title.is_some() || cat_ax_has_major_gridlines {
        let mut cat_axis = ChartAxis::new();
        cat_axis.title = cat_ax_title;
        cat_axis.has_major_gridlines = cat_ax_has_major_gridlines;
        chart.set_category_axis(cat_axis);
    }
    if val_ax_title.is_some()
        || val_ax_has_major_gridlines
        || val_ax_min.is_some()
        || val_ax_max.is_some()
        || val_ax_major_unit.is_some()
    {
        let mut val_axis = ChartAxis::new();
        val_axis.title = val_ax_title;
        val_axis.has_major_gridlines = val_ax_has_major_gridlines;
        val_axis.min_value = val_ax_min;
        val_axis.max_value = val_ax_max;
        val_axis.major_unit = val_ax_major_unit;
        chart.set_value_axis(val_axis);
    }

    // Chart-type-specific properties.
    if let Some(gw) = bar_gap_width {
        chart.set_bar_gap_width(gw);
    }
    if let Some(ov) = bar_overlap {
        chart.set_bar_overlap(ov);
    }
    if let Some(angle) = pie_first_slice_angle {
        chart.set_pie_first_slice_angle(angle);
    }
    if let Some(hs) = pie_hole_size {
        chart.set_pie_hole_size(hs);
    }
    if let Some(ss) = scatter_style {
        chart.set_scatter_style(ss);
    }
    if let Some(bs) = bubble_scale {
        chart.set_bubble_scale(bs);
    }

    Ok(chart)
}

/// Serialize all slide masters and their layouts from the v2 API.
///
/// This function:
/// 1. Creates part URIs for each master and layout
/// 2. Serializes each layout's XML using the slide_layout_io module
/// 3. Serializes each master's XML with references to its layouts
/// 4. Adds all parts and relationships to the package
/// 5. Returns a vector of SlideMasterRef for the presentation XML
fn serialize_all_slide_masters_and_layouts(
    package: &mut Package,
    presentation_uri: &PartUri,
    presentation_part: &mut Part,
    slide_masters_v2: &[SlideMaster],
) -> Result<Vec<SlideMasterRef>> {
    let mut slide_master_refs = Vec::with_capacity(slide_masters_v2.len());
    let mut next_master_id = DEFAULT_SLIDE_MASTER_ID;
    let mut next_layout_id = DEFAULT_SLIDE_LAYOUT_ID;

    for (master_index, master) in slide_masters_v2.iter().enumerate() {
        let master_number = u32::try_from(master_index + 1).map_err(|_| {
            PptxError::UnsupportedPackage("too many slide masters to serialize".to_string())
        })?;

        // Create master part URI (e.g., "/ppt/slideMasters/slideMaster1.xml")
        let master_part_uri_str = format!("/ppt/slideMasters/slideMaster{master_number}.xml");
        let master_part_uri = PartUri::new(&master_part_uri_str)?;

        // Add relationship from presentation to master
        let master_relationship = presentation_part.relationships.add_new(
            RelationshipType::SLIDE_MASTER.to_string(),
            relative_path_from_part(presentation_uri, &master_part_uri),
            TargetMode::Internal,
        );

        let master_id = next_master_id;
        next_master_id = next_master_id.checked_add(1).ok_or_else(|| {
            PptxError::UnsupportedPackage("master id overflow while serializing".to_string())
        })?;

        slide_master_refs.push(SlideMasterRef {
            master_id,
            relationship_id: master_relationship.id.clone(),
        });

        // Serialize layouts for this master
        let mut master_part = Part::new_xml(master_part_uri.clone(), Vec::new());
        master_part.content_type = Some(ContentTypeValue::SLIDE_MASTER.to_string());

        let mut layout_refs = Vec::with_capacity(master.layouts().len());

        for (layout_index, layout) in master.layouts().iter().enumerate() {
            let layout_number =
                u32::try_from(layout_index + 1 + master_index * 10).map_err(|_| {
                    PptxError::UnsupportedPackage("too many slide layouts to serialize".to_string())
                })?;

            // Create layout part URI (e.g., "/ppt/slideLayouts/slideLayout1.xml")
            let layout_part_uri_str = format!("/ppt/slideLayouts/slideLayout{layout_number}.xml");
            let layout_part_uri = PartUri::new(&layout_part_uri_str)?;

            // Add relationship from master to layout
            let layout_relationship = master_part.relationships.add_new(
                RelationshipType::SLIDE_LAYOUT.to_string(),
                relative_path_from_part(&master_part_uri, &layout_part_uri),
                TargetMode::Internal,
            );

            let layout_id = next_layout_id;
            next_layout_id = next_layout_id.checked_add(1).ok_or_else(|| {
                PptxError::UnsupportedPackage("layout id overflow while serializing".to_string())
            })?;

            layout_refs.push(slide_master_io::ParsedLayoutRef {
                id: Some(layout_id.to_string()),
                relationship_id: layout_relationship.id.clone(),
            });

            // Create layout part with back-reference to master
            let mut layout_part = Part::new_xml(layout_part_uri.clone(), Vec::new());
            layout_part.content_type = Some(ContentTypeValue::SLIDE_LAYOUT.to_string());
            layout_part.relationships.add_new(
                RelationshipType::SLIDE_MASTER.to_string(),
                relative_path_from_part(&layout_part_uri, &master_part_uri),
                TargetMode::Internal,
            );

            // Serialize layout XML
            layout_part.data = PartData::Xml(slide_layout_io::write_slide_layout_xml(layout)?);
            package.set_part(layout_part);
        }

        // Serialize master XML with references to all its layouts
        let master_write_data = slide_master_io::WriteSlideMasterData {
            preserve: master.preserve(),
            layout_refs: &layout_refs,
            background: master.background(),
            raw_sp_tree: None, // TODO: preserve raw shape tree if available
            color_map: vec![], // TODO: preserve color map if available
            unknown_children: &[],
        };

        master_part.data =
            PartData::Xml(slide_master_io::write_slide_master_xml(&master_write_data)?);
        package.set_part(master_part);
    }

    Ok(slide_master_refs)
}

fn attach_slide_master_and_layout_parts(
    package: &mut Package,
    slide_master_part_uri: &PartUri,
    slide_layout_part_uri: &PartUri,
    presentation_relationship_id: &str,
) -> Result<SlideMasterMetadata> {
    let mut slide_master_part = Part::new_xml(slide_master_part_uri.clone(), Vec::new());
    slide_master_part.content_type = Some(ContentTypeValue::SLIDE_MASTER.to_string());
    let slide_master_layout_relationship = slide_master_part.relationships.add_new(
        RelationshipType::SLIDE_LAYOUT.to_string(),
        relative_path_from_part(slide_master_part_uri, slide_layout_part_uri),
        TargetMode::Internal,
    );
    let slide_master_layout_relationship_id = slide_master_layout_relationship.id.clone();
    slide_master_part.data = PartData::Xml(serialize_slide_master_xml(
        DEFAULT_SLIDE_LAYOUT_ID,
        slide_master_layout_relationship_id.as_str(),
    )?);

    let mut slide_layout_part = Part::new_xml(slide_layout_part_uri.clone(), Vec::new());
    slide_layout_part.content_type = Some(ContentTypeValue::SLIDE_LAYOUT.to_string());
    slide_layout_part.relationships.add_new(
        RelationshipType::SLIDE_MASTER.to_string(),
        relative_path_from_part(slide_layout_part_uri, slide_master_part_uri),
        TargetMode::Internal,
    );
    slide_layout_part.data = PartData::Xml(serialize_slide_layout_xml()?);

    package.set_part(slide_master_part);
    package.set_part(slide_layout_part);

    Ok(SlideMasterMetadata {
        relationship_id: presentation_relationship_id.to_string(),
        part_uri: slide_master_part_uri.as_str().to_string(),
        layouts: vec![SlideLayoutMetadata {
            relationship_id: slide_master_layout_relationship_id,
            part_uri: slide_layout_part_uri.as_str().to_string(),
            name: None,
            layout_type: Some("title".to_string()),
            preserve: Some(true),
            shapes: Vec::new(),
        }],
        shapes: Vec::new(),
    })
}

fn attach_slide_image_parts(
    package: &mut Package,
    slide_part_uri: &PartUri,
    slide_part: &mut Part,
    images: &[Image],
    next_media_index: &mut u32,
) -> Result<Vec<SerializedImageRef>> {
    let mut image_refs = Vec::with_capacity(images.len());

    for (image_index, image) in images.iter().enumerate() {
        let media_index = *next_media_index;
        *next_media_index = next_media_index.checked_add(1).ok_or_else(|| {
            PptxError::UnsupportedPackage("media index overflow while serializing".to_string())
        })?;

        let extension = extension_for_content_type(image.content_type());
        let media_part_uri = PartUri::new(format!("/ppt/media/image{media_index}.{extension}"))?;
        let relationship_target = relative_path_from_part(slide_part_uri, &media_part_uri);
        let relationship = slide_part.relationships.add_new(
            RelationshipType::IMAGE.to_string(),
            relationship_target,
            TargetMode::Internal,
        );

        let mut media_part = Part::new(media_part_uri, image.bytes().to_vec());
        media_part.content_type = Some(image.content_type().to_string());
        package.set_part(media_part);

        let picture_name = image
            .name()
            .map(str::to_string)
            .unwrap_or_else(|| format!("Picture {}", image_index + 1));

        image_refs.push(SerializedImageRef {
            relationship_id: relationship.id.clone(),
            name: picture_name,
        });
    }

    Ok(image_refs)
}

fn attach_slide_chart_parts(
    package: &mut Package,
    slide_part_uri: &PartUri,
    slide_part: &mut Part,
    charts: &[Chart],
    next_chart_index: &mut u32,
) -> Result<Vec<SerializedChartRef>> {
    let mut chart_refs = Vec::with_capacity(charts.len());

    for chart in charts {
        let chart_index = *next_chart_index;
        *next_chart_index = next_chart_index.checked_add(1).ok_or_else(|| {
            PptxError::UnsupportedPackage("chart index overflow while serializing".to_string())
        })?;

        let chart_part_uri = PartUri::new(format!("/ppt/charts/chart{chart_index}.xml"))?;
        let relationship_target = relative_path_from_part(slide_part_uri, &chart_part_uri);
        let relationship = slide_part.relationships.add_new(
            RelationshipType::CHART.to_string(),
            relationship_target,
            TargetMode::Internal,
        );

        let mut chart_part = Part::new_xml(chart_part_uri, serialize_chart_xml(chart)?);
        chart_part.content_type = Some(CHART_CONTENT_TYPE.to_string());
        package.set_part(chart_part);

        chart_refs.push(SerializedChartRef {
            relationship_id: relationship.id.clone(),
            name: format!("Chart {chart_index}"),
        });
    }

    Ok(chart_refs)
}

fn collect_comment_authors(slides: &[Slide]) -> Result<Vec<SerializedCommentAuthor>> {
    let mut authors = Vec::new();
    let mut author_id_by_name = HashMap::new();

    for slide in slides {
        for comment in slide.comments() {
            if author_id_by_name.contains_key(comment.author()) {
                continue;
            }

            let author_id = u32::try_from(authors.len()).map_err(|_| {
                PptxError::UnsupportedPackage("too many comment authors to serialize".to_string())
            })?;
            author_id_by_name.insert(comment.author().to_string(), author_id);
            authors.push(SerializedCommentAuthor {
                id: author_id,
                name: comment.author().to_string(),
                last_comment_index: 0,
            });
        }
    }

    Ok(authors)
}

fn attach_slide_comments_part(
    package: &mut Package,
    slide_part_uri: &PartUri,
    slide_part: &mut Part,
    slide_number: u32,
    comments: &[SlideComment],
    serialized_comment_authors: &mut [SerializedCommentAuthor],
    comment_authors_part_uri: &PartUri,
) -> Result<()> {
    let comments_part_uri = PartUri::new(format!("/ppt/comments/comment{slide_number}.xml"))?;
    let comments_relationship_target = relative_path_from_part(slide_part_uri, &comments_part_uri);
    slide_part.relationships.add_new(
        COMMENTS_RELATIONSHIP_TYPE.to_string(),
        comments_relationship_target,
        TargetMode::Internal,
    );
    let comment_authors_relationship_target =
        relative_path_from_part(slide_part_uri, comment_authors_part_uri);
    slide_part.relationships.add_new(
        COMMENT_AUTHORS_RELATIONSHIP_TYPE.to_string(),
        comment_authors_relationship_target,
        TargetMode::Internal,
    );

    let author_id_by_name: HashMap<String, u32> = serialized_comment_authors
        .iter()
        .map(|author| (author.name.clone(), author.id))
        .collect();
    let mut serialized_comments = Vec::with_capacity(comments.len());
    for (comment_index, comment) in comments.iter().enumerate() {
        let author_id = *author_id_by_name.get(comment.author()).ok_or_else(|| {
            PptxError::UnsupportedPackage(format!(
                "missing serialized comment author `{}`",
                comment.author()
            ))
        })?;
        let comment_index_u32 = u32::try_from(comment_index).map_err(|_| {
            PptxError::UnsupportedPackage("too many comments on a single slide".to_string())
        })?;

        let author_index = usize::try_from(author_id).map_err(|_| {
            PptxError::UnsupportedPackage("comment author id conversion overflow".to_string())
        })?;
        if let Some(author) = serialized_comment_authors.get_mut(author_index) {
            author.last_comment_index = author.last_comment_index.max(comment_index_u32);
        }

        serialized_comments.push(SerializedSlideComment {
            author_id,
            comment_index: comment_index_u32,
            text: comment.text().to_string(),
        });
    }

    let mut comments_part = Part::new_xml(
        comments_part_uri,
        serialize_comments_xml(&serialized_comments)?,
    );
    comments_part.content_type = Some(COMMENTS_CONTENT_TYPE.to_string());
    package.set_part(comments_part);

    Ok(())
}

fn attach_slide_notes_part(
    package: &mut Package,
    slide_part_uri: &PartUri,
    slide_part: &mut Part,
    slide_number: u32,
    notes_text: &str,
) -> Result<()> {
    let notes_part_uri = PartUri::new(format!("/ppt/notesSlides/notesSlide{slide_number}.xml"))?;
    let relationship_target = relative_path_from_part(slide_part_uri, &notes_part_uri);
    slide_part.relationships.add_new(
        NOTES_SLIDE_RELATIONSHIP_TYPE.to_string(),
        relationship_target,
        TargetMode::Internal,
    );

    let mut notes_part = Part::new_xml(
        notes_part_uri.clone(),
        serialize_notes_slide_xml(notes_text)?,
    );
    notes_part.content_type = Some(NOTES_SLIDE_CONTENT_TYPE.to_string());
    notes_part.relationships.add_new(
        RelationshipType::SLIDE.to_string(),
        relative_path_from_part(&notes_part_uri, slide_part_uri),
        TargetMode::Internal,
    );
    package.set_part(notes_part);

    Ok(())
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

fn serialize_presentation_xml(
    slide_refs: &[SlideRef],
    slide_master_refs: &[SlideMasterRef],
    slide_width_emu: Option<i64>,
    slide_height_emu: Option<i64>,
    first_slide_number: Option<u32>,
    show_special_pls_on_title_sld: Option<bool>,
    right_to_left: Option<bool>,
) -> Result<Vec<u8>> {
    let mut writer = Writer::new_with_indent(Vec::new(), b' ', 2);
    writer.write_event(Event::Decl(BytesDecl::new(
        "1.0",
        Some("UTF-8"),
        Some("yes"),
    )))?;

    let mut presentation = BytesStart::new("p:presentation");
    presentation.push_attribute(("xmlns:p", PRESENTATIONML_NS));
    presentation.push_attribute(("xmlns:a", DRAWINGML_NS));
    presentation.push_attribute(("xmlns:r", RELATIONSHIP_NS));
    let first_slide_num_text = first_slide_number.map(|v| v.to_string());
    if let Some(ref fsn) = first_slide_num_text {
        presentation.push_attribute(("firstSlideNum", fsn.as_str()));
    }
    if let Some(show) = show_special_pls_on_title_sld {
        presentation.push_attribute(("showSpecialPlsOnTitleSld", xml_bool_value(show)));
    }
    if let Some(rtl) = right_to_left {
        presentation.push_attribute(("rtl", xml_bool_value(rtl)));
    }
    writer.write_event(Event::Start(presentation))?;

    writer.write_event(Event::Start(BytesStart::new("p:sldMasterIdLst")))?;
    for slide_master_ref in slide_master_refs {
        let mut slide_master_id = BytesStart::new("p:sldMasterId");
        let slide_master_id_text = slide_master_ref.master_id.to_string();
        slide_master_id.push_attribute(("id", slide_master_id_text.as_str()));
        slide_master_id.push_attribute(("r:id", slide_master_ref.relationship_id.as_str()));
        writer.write_event(Event::Empty(slide_master_id))?;
    }
    writer.write_event(Event::End(BytesEnd::new("p:sldMasterIdLst")))?;

    writer.write_event(Event::Start(BytesStart::new("p:sldIdLst")))?;
    for slide_ref in slide_refs {
        let mut slide_id = BytesStart::new("p:sldId");
        let slide_id_text = slide_ref.slide_id.to_string();
        slide_id.push_attribute(("id", slide_id_text.as_str()));
        slide_id.push_attribute(("r:id", slide_ref.relationship_id.as_str()));
        writer.write_event(Event::Empty(slide_id))?;
    }
    writer.write_event(Event::End(BytesEnd::new("p:sldIdLst")))?;

    // Feature #9: Slide size.
    if slide_width_emu.is_some() || slide_height_emu.is_some() {
        let mut sld_sz = BytesStart::new("p:sldSz");
        let cx_text = slide_width_emu.unwrap_or(9_144_000).to_string();
        let cy_text = slide_height_emu.unwrap_or(6_858_000).to_string();
        sld_sz.push_attribute(("cx", cx_text.as_str()));
        sld_sz.push_attribute(("cy", cy_text.as_str()));
        writer.write_event(Event::Empty(sld_sz))?;
    }

    writer.write_event(Event::End(BytesEnd::new("p:presentation")))?;

    Ok(writer.into_inner())
}

fn serialize_slide_master_xml(layout_id: u32, layout_relationship_id: &str) -> Result<Vec<u8>> {
    // Create minimal data using the new I/O module.
    let layout_ref = slide_master_io::ParsedLayoutRef {
        id: Some(layout_id.to_string()),
        relationship_id: layout_relationship_id.to_string(),
    };
    let data = slide_master_io::WriteSlideMasterData {
        preserve: false,
        layout_refs: &[layout_ref],
        background: None,
        raw_sp_tree: None,
        color_map: vec![],
        unknown_children: &[],
    };

    slide_master_io::write_slide_master_xml(&data)
}

fn serialize_slide_layout_xml() -> Result<Vec<u8>> {
    // Create a minimal layout using the new I/O module.
    let mut layout = SlideLayout::new("", "rId1", DEFAULT_SLIDE_LAYOUT_PART_URI, "rId2");
    layout.set_layout_type("title");
    layout.set_preserve(true);

    slide_layout_io::write_slide_layout_xml(&layout)
}

fn serialize_slide_xml(
    slide: &Slide,
    image_refs: &[SerializedImageRef],
    chart_refs: &[SerializedChartRef],
) -> Result<Vec<u8>> {
    let mut writer = Writer::new_with_indent(Vec::new(), b' ', 2);
    writer.write_event(Event::Decl(BytesDecl::new(
        "1.0",
        Some("UTF-8"),
        Some("yes"),
    )))?;

    let mut sld = BytesStart::new("p:sld");
    sld.push_attribute(("xmlns:p", PRESENTATIONML_NS));
    sld.push_attribute(("xmlns:a", DRAWINGML_NS));
    sld.push_attribute(("xmlns:c", CHART_NS));
    sld.push_attribute(("xmlns:r", RELATIONSHIP_NS));
    // Replay extra namespace declarations from the original XML.
    for (prefix, uri) in slide.extra_namespace_declarations() {
        sld.push_attribute((prefix.as_str(), uri.as_str()));
    }
    // Feature #8: Slide hidden state.
    if slide.is_hidden() {
        sld.push_attribute(("show", "0"));
    }
    writer.write_event(Event::Start(sld))?;

    writer.write_event(Event::Start(BytesStart::new("p:cSld")))?;

    // Feature #7: Slide background.
    if let Some(background) = slide.background() {
        write_slide_background_xml(&mut writer, background)?;
    }

    writer.write_event(Event::Start(BytesStart::new("p:spTree")))?;

    // TODO(roundtrip): spTree children (shapes, tables, images, charts, groups) are currently
    // written in separate batches rather than in their original document order. This means
    // the z-order of mixed element types (e.g., a shape between two images) is not preserved.
    // To fix this, the Slide struct would need a single ordered list (e.g., `Vec<SlideChild>`
    // enum with Shape/Table/Image/Chart/Group/Unknown variants) that preserves the original
    // child ordering from parsing. This is a significant refactor.
    let mut next_object_id = 1_u32;

    // Build a temporary Shape for the legacy title + text_runs fields.
    let mut title_shape = Shape::new("Title 1");
    {
        let para = title_shape.add_paragraph();
        para.add_run(slide.title());
        for run in slide.text_runs() {
            para.add_run(run.text());
        }
    }
    write_shape_xml(
        &mut writer,
        allocate_object_id(&mut next_object_id)?,
        &title_shape,
    )?;

    for shape in slide.shapes() {
        write_shape_xml(&mut writer, allocate_object_id(&mut next_object_id)?, shape)?;
    }

    for (table_index, table) in slide.tables().iter().enumerate() {
        write_table_xml(
            &mut writer,
            allocate_object_id(&mut next_object_id)?,
            table_index + 1,
            table,
        )?;
    }

    for image_ref in image_refs {
        write_picture_xml(
            &mut writer,
            allocate_object_id(&mut next_object_id)?,
            image_ref.name.as_str(),
            image_ref.relationship_id.as_str(),
        )?;
    }

    for chart_ref in chart_refs {
        write_chart_graphic_frame_xml(
            &mut writer,
            allocate_object_id(&mut next_object_id)?,
            chart_ref.name.as_str(),
            chart_ref.relationship_id.as_str(),
        )?;
    }

    // Feature #12: Grouped shapes.
    for group in slide.grouped_shapes() {
        write_shape_group_xml(&mut writer, &mut next_object_id, group)?;
    }

    writer.write_event(Event::End(BytesEnd::new("p:spTree")))?;
    writer.write_event(Event::End(BytesEnd::new("p:cSld")))?;

    // Feature #15: Header/footer configuration.
    if let Some(hf) = slide.header_footer() {
        write_slide_header_footer_xml(&mut writer, hf)?;
    }

    if let Some(transition) = slide.transition() {
        write_slide_transition_xml(&mut writer, transition)?;
    }
    if let Some(timing) = slide.timing() {
        write_slide_timing_xml(&mut writer, timing)?;
    }
    for unknown in slide.unknown_children() {
        unknown.write_to(&mut writer)?;
    }
    writer.write_event(Event::End(BytesEnd::new("p:sld")))?;

    Ok(writer.into_inner())
}

fn write_slide_transition_xml<W: std::io::Write>(
    writer: &mut Writer<W>,
    transition: &SlideTransition,
) -> Result<()> {
    let mut transition_element = BytesStart::new("p:transition");

    let advance_on_click = transition.advance_on_click().map(xml_bool_value);
    if let Some(advance_on_click) = advance_on_click {
        transition_element.push_attribute(("advClick", advance_on_click));
    }

    let advance_after_ms = transition.advance_after_ms().map(|value| value.to_string());
    if let Some(ref advance_after_ms) = advance_after_ms {
        transition_element.push_attribute(("advTm", advance_after_ms.as_str()));
    }

    if let Some(speed) = transition.speed() {
        transition_element.push_attribute(("spd", speed.to_xml()));
    }

    let has_kind = !matches!(transition.kind(), SlideTransitionKind::Unspecified);
    let has_sound = transition.sound().is_some();

    if has_kind || has_sound {
        writer.write_event(Event::Start(transition_element))?;

        if has_kind {
            let transition_child_name =
                format!("p:{}", transition_kind_to_xml_name(transition.kind()));
            writer.write_event(Event::Empty(BytesStart::new(
                transition_child_name.as_str(),
            )))?;
        }

        // Sound action.
        if let Some(snd) = transition.sound() {
            writer.write_event(Event::Start(BytesStart::new("p:sndAc")))?;
            writer.write_event(Event::Start(BytesStart::new("p:stSnd")))?;
            let mut snd_elem = BytesStart::new("p:snd");
            snd_elem.push_attribute(("r:embed", snd.relationship_id.as_str()));
            snd_elem.push_attribute(("name", snd.name.as_str()));
            writer.write_event(Event::Empty(snd_elem))?;
            writer.write_event(Event::End(BytesEnd::new("p:stSnd")))?;
            writer.write_event(Event::End(BytesEnd::new("p:sndAc")))?;
        }

        writer.write_event(Event::End(BytesEnd::new("p:transition")))?;
    } else {
        writer.write_event(Event::Empty(transition_element))?;
    }

    Ok(())
}

fn write_slide_timing_xml<W: std::io::Write>(
    writer: &mut Writer<W>,
    timing: &SlideTiming,
) -> Result<()> {
    if timing.raw_inner_xml().is_empty() && timing.animations().is_empty() {
        writer.write_event(Event::Empty(BytesStart::new("p:timing")))?;
        return Ok(());
    }

    if !timing.raw_inner_xml().is_empty() {
        ensure_valid_timing_inner_xml(timing.raw_inner_xml())?;
        writer.get_mut().write_all(b"<p:timing>")?;
        writer
            .get_mut()
            .write_all(timing.raw_inner_xml().as_bytes())?;
        writer.get_mut().write_all(b"</p:timing>")?;
        return Ok(());
    }

    writer.write_event(Event::Start(BytesStart::new("p:timing")))?;
    write_typed_timing_inner_xml(writer, timing.animations())?;
    writer.write_event(Event::End(BytesEnd::new("p:timing")))?;
    Ok(())
}

fn write_typed_timing_inner_xml<W: std::io::Write>(
    writer: &mut Writer<W>,
    animations: &[SlideAnimationNode],
) -> Result<()> {
    writer.write_event(Event::Start(BytesStart::new("p:tnLst")))?;

    for animation in animations {
        writer.write_event(Event::Start(BytesStart::new("p:par")))?;

        let id = animation.id().to_string();
        let duration_ms = animation.duration_ms().map(|value| value.to_string());
        let mut c_tn = BytesStart::new("p:cTn");
        c_tn.push_attribute(("id", id.as_str()));
        if let Some(ref duration_ms) = duration_ms {
            c_tn.push_attribute(("dur", duration_ms.as_str()));
        }
        if let Some(trigger) = animation.trigger() {
            c_tn.push_attribute(("nodeType", trigger));
        }
        if let Some(event) = animation.event() {
            c_tn.push_attribute(("evtFilter", event));
        }

        writer.write_event(Event::Empty(c_tn))?;
        writer.write_event(Event::End(BytesEnd::new("p:par")))?;
    }

    writer.write_event(Event::End(BytesEnd::new("p:tnLst")))?;
    Ok(())
}

fn ensure_valid_timing_inner_xml(inner_xml: &str) -> Result<()> {
    let mut reader = Reader::from_reader(Cursor::new(inner_xml.as_bytes()));
    reader.config_mut().trim_text(false);
    let mut buffer = Vec::new();

    loop {
        if reader.read_event_into(&mut buffer)? == Event::Eof {
            break;
        }
        buffer.clear();
    }

    Ok(())
}

fn serialize_notes_slide_xml(notes_text: &str) -> Result<Vec<u8>> {
    let mut writer = Writer::new_with_indent(Vec::new(), b' ', 2);
    writer.write_event(Event::Decl(BytesDecl::new(
        "1.0",
        Some("UTF-8"),
        Some("yes"),
    )))?;

    let mut notes = BytesStart::new("p:notes");
    notes.push_attribute(("xmlns:p", PRESENTATIONML_NS));
    notes.push_attribute(("xmlns:a", DRAWINGML_NS));
    notes.push_attribute(("xmlns:r", RELATIONSHIP_NS));
    writer.write_event(Event::Start(notes))?;

    writer.write_event(Event::Start(BytesStart::new("p:cSld")))?;
    writer.write_event(Event::Start(BytesStart::new("p:spTree")))?;

    writer.write_event(Event::Start(BytesStart::new("p:sp")))?;
    writer.write_event(Event::Start(BytesStart::new("p:nvSpPr")))?;

    let mut c_nv_pr = BytesStart::new("p:cNvPr");
    c_nv_pr.push_attribute(("id", "1"));
    c_nv_pr.push_attribute(("name", "Notes Placeholder 1"));
    writer.write_event(Event::Empty(c_nv_pr))?;

    writer.write_event(Event::Empty(BytesStart::new("p:cNvSpPr")))?;
    writer.write_event(Event::Empty(BytesStart::new("p:nvPr")))?;
    writer.write_event(Event::End(BytesEnd::new("p:nvSpPr")))?;

    writer.write_event(Event::Start(BytesStart::new("p:txBody")))?;
    writer.write_event(Event::Empty(BytesStart::new("a:bodyPr")))?;
    writer.write_event(Event::Empty(BytesStart::new("a:lstStyle")))?;

    for paragraph in notes_text.split('\n') {
        writer.write_event(Event::Start(BytesStart::new("a:p")))?;
        if !paragraph.is_empty() {
            write_plain_text_run(&mut writer, paragraph)?;
        }
        writer.write_event(Event::End(BytesEnd::new("a:p")))?;
    }

    writer.write_event(Event::End(BytesEnd::new("p:txBody")))?;
    writer.write_event(Event::End(BytesEnd::new("p:sp")))?;

    writer.write_event(Event::End(BytesEnd::new("p:spTree")))?;
    writer.write_event(Event::End(BytesEnd::new("p:cSld")))?;
    writer.write_event(Event::End(BytesEnd::new("p:notes")))?;

    Ok(writer.into_inner())
}

fn serialize_chart_xml(chart: &Chart) -> Result<Vec<u8>> {
    let mut writer = Writer::new_with_indent(Vec::new(), b' ', 2);
    writer.write_event(Event::Decl(BytesDecl::new(
        "1.0",
        Some("UTF-8"),
        Some("yes"),
    )))?;

    let mut chart_space = BytesStart::new("c:chartSpace");
    chart_space.push_attribute(("xmlns:c", CHART_NS));
    chart_space.push_attribute(("xmlns:a", DRAWINGML_NS));
    chart_space.push_attribute(("xmlns:r", RELATIONSHIP_NS));
    writer.write_event(Event::Start(chart_space))?;

    writer.write_event(Event::Start(BytesStart::new("c:chart")))?;
    writer.write_event(Event::Start(BytesStart::new("c:title")))?;
    writer.write_event(Event::Start(BytesStart::new("c:tx")))?;
    writer.write_event(Event::Start(BytesStart::new("c:rich")))?;
    writer.write_event(Event::Empty(BytesStart::new("a:bodyPr")))?;
    writer.write_event(Event::Empty(BytesStart::new("a:lstStyle")))?;
    writer.write_event(Event::Start(BytesStart::new("a:p")))?;
    write_plain_text_run(&mut writer, chart.title())?;
    writer.write_event(Event::End(BytesEnd::new("a:p")))?;
    writer.write_event(Event::End(BytesEnd::new("c:rich")))?;
    writer.write_event(Event::End(BytesEnd::new("c:tx")))?;
    writer.write_event(Event::End(BytesEnd::new("c:title")))?;

    writer.write_event(Event::Start(BytesStart::new("c:plotArea")))?;
    writer.write_event(Event::Empty(BytesStart::new("c:layout")))?;

    // Feature #13: Use the chart's chart_type for the element name.
    let chart_element_name = chart.chart_type().to_xml_element();
    writer.write_event(Event::Start(BytesStart::new(chart_element_name)))?;

    // Bar/Column charts need barDir and grouping.
    let is_bar_type = matches!(chart.chart_type(), ChartType::Bar | ChartType::Column);
    if is_bar_type {
        let mut bar_dir = BytesStart::new("c:barDir");
        let bar_dir_val = if chart.is_bar_direction_horizontal() {
            "bar"
        } else {
            "col"
        };
        bar_dir.push_attribute(("val", bar_dir_val));
        writer.write_event(Event::Empty(bar_dir))?;
        let mut grouping = BytesStart::new("c:grouping");
        grouping.push_attribute(("val", "clustered"));
        writer.write_event(Event::Empty(grouping))?;
    }

    // Scatter chart style.
    if chart.chart_type() == ChartType::Scatter {
        if let Some(ss) = chart.scatter_style() {
            let mut elem = BytesStart::new("c:scatterStyle");
            elem.push_attribute(("val", ss.to_xml()));
            writer.write_event(Event::Empty(elem))?;
        }
    }

    // Write first series (series index 0).
    write_chart_series_xml(
        &mut writer,
        0,
        chart.title(),
        chart.categories(),
        chart.values(),
        None,
        None,
        None,
    )?;

    // Feature #13: Write additional series.
    for (i, series) in chart.additional_series().iter().enumerate() {
        let series_idx = u32::try_from(i + 1).unwrap_or(1);
        write_chart_series_xml(
            &mut writer,
            series_idx,
            series.name(),
            chart.categories(),
            series.values(),
            series.fill(),
            series.border(),
            series.marker(),
        )?;
    }

    // Data labels.
    if let Some(data_labels) = chart.data_labels() {
        write_chart_data_labels(&mut writer, data_labels)?;
    }

    // Chart-type-specific properties (inside the chart type element).
    if is_bar_type {
        if let Some(gw) = chart.bar_gap_width() {
            let mut elem = BytesStart::new("c:gapWidth");
            let gw_str = gw.to_string();
            elem.push_attribute(("val", gw_str.as_str()));
            writer.write_event(Event::Empty(elem))?;
        }
        if let Some(ov) = chart.bar_overlap() {
            let mut elem = BytesStart::new("c:overlap");
            let ov_str = ov.to_string();
            elem.push_attribute(("val", ov_str.as_str()));
            writer.write_event(Event::Empty(elem))?;
        }
    }

    // Pie chart properties.
    if matches!(chart.chart_type(), ChartType::Pie | ChartType::Doughnut) {
        if let Some(angle) = chart.pie_first_slice_angle() {
            let mut elem = BytesStart::new("c:firstSliceAng");
            let angle_str = angle.to_string();
            elem.push_attribute(("val", angle_str.as_str()));
            writer.write_event(Event::Empty(elem))?;
        }
    }

    // Doughnut hole size.
    if chart.chart_type() == ChartType::Doughnut {
        if let Some(hs) = chart.pie_hole_size() {
            let mut elem = BytesStart::new("c:holeSize");
            let hs_str = hs.to_string();
            elem.push_attribute(("val", hs_str.as_str()));
            writer.write_event(Event::Empty(elem))?;
        }
    }

    // Bubble chart scale.
    if chart.chart_type() == ChartType::Scatter {
        if let Some(bs) = chart.bubble_scale() {
            let mut elem = BytesStart::new("c:bubbleScale");
            let bs_str = bs.to_string();
            elem.push_attribute(("val", bs_str.as_str()));
            writer.write_event(Event::Empty(elem))?;
        }
    }

    // Axes (only for chart types that have them).
    let needs_axes = !matches!(chart.chart_type(), ChartType::Pie | ChartType::Doughnut);
    if needs_axes {
        let mut first_axis_id = BytesStart::new("c:axId");
        first_axis_id.push_attribute(("val", "1"));
        writer.write_event(Event::Empty(first_axis_id))?;
        let mut second_axis_id = BytesStart::new("c:axId");
        second_axis_id.push_attribute(("val", "2"));
        writer.write_event(Event::Empty(second_axis_id))?;
    }

    writer.write_event(Event::End(BytesEnd::new(chart_element_name)))?;

    if needs_axes {
        writer.write_event(Event::Start(BytesStart::new("c:catAx")))?;
        let mut cat_axis_id = BytesStart::new("c:axId");
        cat_axis_id.push_attribute(("val", "1"));
        writer.write_event(Event::Empty(cat_axis_id))?;
        // Feature #4 (chart axes): category axis properties.
        if let Some(cat_axis) = chart.category_axis() {
            if let Some(ref title) = cat_axis.title {
                write_chart_axis_title(&mut writer, title)?;
            }
            if cat_axis.has_major_gridlines {
                writer.write_event(Event::Empty(BytesStart::new("c:majorGridlines")))?;
            }
        }
        writer.write_event(Event::End(BytesEnd::new("c:catAx")))?;

        writer.write_event(Event::Start(BytesStart::new("c:valAx")))?;
        let mut val_axis_id = BytesStart::new("c:axId");
        val_axis_id.push_attribute(("val", "2"));
        writer.write_event(Event::Empty(val_axis_id))?;
        // Feature #4 (chart axes): value axis properties.
        if let Some(val_axis) = chart.value_axis() {
            if let Some(ref title) = val_axis.title {
                write_chart_axis_title(&mut writer, title)?;
            }
            if val_axis.has_major_gridlines {
                writer.write_event(Event::Empty(BytesStart::new("c:majorGridlines")))?;
            }
            // Scaling properties.
            let has_scaling = val_axis.min_value.is_some() || val_axis.max_value.is_some();
            if has_scaling {
                writer.write_event(Event::Start(BytesStart::new("c:scaling")))?;
                if let Some(min_val) = val_axis.min_value {
                    let mut min_elem = BytesStart::new("c:min");
                    let min_text = min_val.to_string();
                    min_elem.push_attribute(("val", min_text.as_str()));
                    writer.write_event(Event::Empty(min_elem))?;
                }
                if let Some(max_val) = val_axis.max_value {
                    let mut max_elem = BytesStart::new("c:max");
                    let max_text = max_val.to_string();
                    max_elem.push_attribute(("val", max_text.as_str()));
                    writer.write_event(Event::Empty(max_elem))?;
                }
                writer.write_event(Event::End(BytesEnd::new("c:scaling")))?;
            }
            if let Some(major_unit) = val_axis.major_unit {
                let mut mu = BytesStart::new("c:majorUnit");
                let mu_text = major_unit.to_string();
                mu.push_attribute(("val", mu_text.as_str()));
                writer.write_event(Event::Empty(mu))?;
            }
        }
        writer.write_event(Event::End(BytesEnd::new("c:valAx")))?;
    }

    writer.write_event(Event::End(BytesEnd::new("c:plotArea")))?;

    // Feature #13: Legend.
    if chart.show_legend() {
        writer.write_event(Event::Start(BytesStart::new("c:legend")))?;
        if let Some(pos) = chart.legend_position() {
            let mut legend_pos = BytesStart::new("c:legendPos");
            legend_pos.push_attribute(("val", pos.to_xml()));
            writer.write_event(Event::Empty(legend_pos))?;
        }
        writer.write_event(Event::End(BytesEnd::new("c:legend")))?;
    }

    writer.write_event(Event::End(BytesEnd::new("c:chart")))?;
    writer.write_event(Event::End(BytesEnd::new("c:chartSpace")))?;

    Ok(writer.into_inner())
}

/// Write a single chart data series (`<c:ser>`) element.
fn write_chart_series_xml<W: std::io::Write>(
    writer: &mut Writer<W>,
    index: u32,
    name: &str,
    categories: &[String],
    values: &[f64],
    fill: Option<&SeriesFill>,
    border: Option<&SeriesBorder>,
    marker: Option<&SeriesMarker>,
) -> Result<()> {
    writer.write_event(Event::Start(BytesStart::new("c:ser")))?;

    let idx_text = index.to_string();
    let mut series_idx = BytesStart::new("c:idx");
    series_idx.push_attribute(("val", idx_text.as_str()));
    writer.write_event(Event::Empty(series_idx))?;
    let mut series_order = BytesStart::new("c:order");
    series_order.push_attribute(("val", idx_text.as_str()));
    writer.write_event(Event::Empty(series_order))?;

    // Series name.
    writer.write_event(Event::Start(BytesStart::new("c:tx")))?;
    writer.write_event(Event::Start(BytesStart::new("c:v")))?;
    writer.write_event(Event::Text(BytesText::new(name)))?;
    writer.write_event(Event::End(BytesEnd::new("c:v")))?;
    writer.write_event(Event::End(BytesEnd::new("c:tx")))?;

    // Series shape properties (fill and border).
    let has_fill = matches!(fill, Some(f) if !matches!(f, SeriesFill::None));
    let has_border = border.is_some();
    if has_fill || has_border {
        writer.write_event(Event::Start(BytesStart::new("c:spPr")))?;
        if let Some(series_fill) = fill {
            write_chart_series_fill(writer, series_fill)?;
        }
        if let Some(series_border) = border {
            write_chart_series_border(writer, series_border)?;
        }
        writer.write_event(Event::End(BytesEnd::new("c:spPr")))?;
    }

    // Marker styling.
    if let Some(m) = marker {
        write_chart_series_marker(writer, m)?;
    }

    // Categories.
    let point_count = categories.len().min(values.len());
    if !categories.is_empty() {
        writer.write_event(Event::Start(BytesStart::new("c:cat")))?;
        writer.write_event(Event::Start(BytesStart::new("c:strLit")))?;
        let cat_count_text = point_count.to_string();
        let mut cat_count = BytesStart::new("c:ptCount");
        cat_count.push_attribute(("val", cat_count_text.as_str()));
        writer.write_event(Event::Empty(cat_count))?;
        for (i, category) in categories.iter().take(point_count).enumerate() {
            let mut point = BytesStart::new("c:pt");
            let i_text = i.to_string();
            point.push_attribute(("idx", i_text.as_str()));
            writer.write_event(Event::Start(point))?;
            writer.write_event(Event::Start(BytesStart::new("c:v")))?;
            writer.write_event(Event::Text(BytesText::new(category.as_str())))?;
            writer.write_event(Event::End(BytesEnd::new("c:v")))?;
            writer.write_event(Event::End(BytesEnd::new("c:pt")))?;
        }
        writer.write_event(Event::End(BytesEnd::new("c:strLit")))?;
        writer.write_event(Event::End(BytesEnd::new("c:cat")))?;
    }

    // Values.
    writer.write_event(Event::Start(BytesStart::new("c:val")))?;
    writer.write_event(Event::Start(BytesStart::new("c:numLit")))?;
    let val_count_text = values.len().to_string();
    let mut val_count = BytesStart::new("c:ptCount");
    val_count.push_attribute(("val", val_count_text.as_str()));
    writer.write_event(Event::Empty(val_count))?;
    for (i, value) in values.iter().enumerate() {
        let mut point = BytesStart::new("c:pt");
        let i_text = i.to_string();
        point.push_attribute(("idx", i_text.as_str()));
        writer.write_event(Event::Start(point))?;
        writer.write_event(Event::Start(BytesStart::new("c:v")))?;
        let value_text = value.to_string();
        writer.write_event(Event::Text(BytesText::new(value_text.as_str())))?;
        writer.write_event(Event::End(BytesEnd::new("c:v")))?;
        writer.write_event(Event::End(BytesEnd::new("c:pt")))?;
    }
    writer.write_event(Event::End(BytesEnd::new("c:numLit")))?;
    writer.write_event(Event::End(BytesEnd::new("c:val")))?;

    writer.write_event(Event::End(BytesEnd::new("c:ser")))?;
    Ok(())
}

/// Write chart series fill element inside `c:spPr`.
fn write_chart_series_fill<W: std::io::Write>(
    writer: &mut Writer<W>,
    fill: &SeriesFill,
) -> Result<()> {
    match fill {
        SeriesFill::Solid(color) => {
            writer.write_event(Event::Start(BytesStart::new("a:solidFill")))?;
            let mut srgb = BytesStart::new("a:srgbClr");
            srgb.push_attribute(("val", color.as_str()));
            writer.write_event(Event::Empty(srgb))?;
            writer.write_event(Event::End(BytesEnd::new("a:solidFill")))?;
        }
        SeriesFill::Gradient(start, end) => {
            writer.write_event(Event::Start(BytesStart::new("a:gradFill")))?;
            writer.write_event(Event::Start(BytesStart::new("a:gsLst")))?;
            // Start stop at 0%.
            let mut gs0 = BytesStart::new("a:gs");
            gs0.push_attribute(("pos", "0"));
            writer.write_event(Event::Start(gs0))?;
            let mut srgb_start = BytesStart::new("a:srgbClr");
            srgb_start.push_attribute(("val", start.as_str()));
            writer.write_event(Event::Empty(srgb_start))?;
            writer.write_event(Event::End(BytesEnd::new("a:gs")))?;
            // End stop at 100%.
            let mut gs1 = BytesStart::new("a:gs");
            gs1.push_attribute(("pos", "100000"));
            writer.write_event(Event::Start(gs1))?;
            let mut srgb_end = BytesStart::new("a:srgbClr");
            srgb_end.push_attribute(("val", end.as_str()));
            writer.write_event(Event::Empty(srgb_end))?;
            writer.write_event(Event::End(BytesEnd::new("a:gs")))?;
            writer.write_event(Event::End(BytesEnd::new("a:gsLst")))?;
            writer.write_event(Event::End(BytesEnd::new("a:gradFill")))?;
        }
        SeriesFill::Pattern(pattern_type, fg, bg) => {
            let mut patt = BytesStart::new("a:pattFill");
            patt.push_attribute(("prst", pattern_type.as_str()));
            writer.write_event(Event::Start(patt))?;
            writer.write_event(Event::Start(BytesStart::new("a:fgClr")))?;
            let mut fg_clr = BytesStart::new("a:srgbClr");
            fg_clr.push_attribute(("val", fg.as_str()));
            writer.write_event(Event::Empty(fg_clr))?;
            writer.write_event(Event::End(BytesEnd::new("a:fgClr")))?;
            writer.write_event(Event::Start(BytesStart::new("a:bgClr")))?;
            let mut bg_clr = BytesStart::new("a:srgbClr");
            bg_clr.push_attribute(("val", bg.as_str()));
            writer.write_event(Event::Empty(bg_clr))?;
            writer.write_event(Event::End(BytesEnd::new("a:bgClr")))?;
            writer.write_event(Event::End(BytesEnd::new("a:pattFill")))?;
        }
        SeriesFill::None | SeriesFill::Picture(_) => {
            // NoFill for None; Picture fill requires relationship IDs (not supported here).
            writer.write_event(Event::Empty(BytesStart::new("a:noFill")))?;
        }
    }
    Ok(())
}

/// Write chart series border/outline inside `c:spPr`.
fn write_chart_series_border<W: std::io::Write>(
    writer: &mut Writer<W>,
    border: &SeriesBorder,
) -> Result<()> {
    let width_emu = (border.width * 12700.0) as i64;
    let mut ln = BytesStart::new("a:ln");
    let w_text = width_emu.to_string();
    ln.push_attribute(("w", w_text.as_str()));
    writer.write_event(Event::Start(ln))?;
    writer.write_event(Event::Start(BytesStart::new("a:solidFill")))?;
    let mut srgb = BytesStart::new("a:srgbClr");
    srgb.push_attribute(("val", border.color.as_str()));
    writer.write_event(Event::Empty(srgb))?;
    writer.write_event(Event::End(BytesEnd::new("a:solidFill")))?;
    if border.dash_style != "solid" {
        let mut dash = BytesStart::new("a:prstDash");
        dash.push_attribute(("val", border.dash_style.as_str()));
        writer.write_event(Event::Empty(dash))?;
    }
    writer.write_event(Event::End(BytesEnd::new("a:ln")))?;
    Ok(())
}

/// Write chart series marker element (`c:marker`).
fn write_chart_series_marker<W: std::io::Write>(
    writer: &mut Writer<W>,
    marker: &SeriesMarker,
) -> Result<()> {
    writer.write_event(Event::Start(BytesStart::new("c:marker")))?;
    let mut symbol = BytesStart::new("c:symbol");
    symbol.push_attribute(("val", marker.shape.to_xml()));
    writer.write_event(Event::Empty(symbol))?;
    let size_text = marker.size.to_string();
    let mut size_elem = BytesStart::new("c:size");
    size_elem.push_attribute(("val", size_text.as_str()));
    writer.write_event(Event::Empty(size_elem))?;
    // Marker fill/border via spPr.
    let has_marker_fill = matches!(&marker.fill, Some(f) if !matches!(f, SeriesFill::None));
    let has_marker_border = marker.border.is_some();
    if has_marker_fill || has_marker_border {
        writer.write_event(Event::Start(BytesStart::new("c:spPr")))?;
        if let Some(ref mfill) = marker.fill {
            write_chart_series_fill(writer, mfill)?;
        }
        if let Some(ref mborder) = marker.border {
            write_chart_series_border(writer, mborder)?;
        }
        writer.write_event(Event::End(BytesEnd::new("c:spPr")))?;
    }
    writer.write_event(Event::End(BytesEnd::new("c:marker")))?;
    Ok(())
}

/// Write chart data labels element (`c:dLbls`).
fn write_chart_data_labels<W: std::io::Write>(
    writer: &mut Writer<W>,
    labels: &ChartDataLabel,
) -> Result<()> {
    writer.write_event(Event::Start(BytesStart::new("c:dLbls")))?;
    if let Some(ref pos) = labels.position {
        let mut elem = BytesStart::new("c:dLblPos");
        elem.push_attribute(("val", pos.as_str()));
        writer.write_event(Event::Empty(elem))?;
    }
    let show_val = if labels.show_value { "1" } else { "0" };
    let mut sv = BytesStart::new("c:showVal");
    sv.push_attribute(("val", show_val));
    writer.write_event(Event::Empty(sv))?;
    let show_cat = if labels.show_category_name { "1" } else { "0" };
    let mut sc = BytesStart::new("c:showCatName");
    sc.push_attribute(("val", show_cat));
    writer.write_event(Event::Empty(sc))?;
    let show_ser = if labels.show_series_name { "1" } else { "0" };
    let mut ss = BytesStart::new("c:showSerName");
    ss.push_attribute(("val", show_ser));
    writer.write_event(Event::Empty(ss))?;
    let show_pct = if labels.show_percentage { "1" } else { "0" };
    let mut sp = BytesStart::new("c:showPercent");
    sp.push_attribute(("val", show_pct));
    writer.write_event(Event::Empty(sp))?;
    let show_lk = if labels.show_legend_key { "1" } else { "0" };
    let mut slk = BytesStart::new("c:showLegendKey");
    slk.push_attribute(("val", show_lk));
    writer.write_event(Event::Empty(slk))?;
    if let Some(ref sep) = labels.separator {
        writer.write_event(Event::Start(BytesStart::new("c:separator")))?;
        writer.write_event(Event::Text(BytesText::new(sep.as_str())))?;
        writer.write_event(Event::End(BytesEnd::new("c:separator")))?;
    }
    writer.write_event(Event::End(BytesEnd::new("c:dLbls")))?;
    Ok(())
}

/// Write a chart axis title element.
fn write_chart_axis_title<W: std::io::Write>(writer: &mut Writer<W>, title: &str) -> Result<()> {
    writer.write_event(Event::Start(BytesStart::new("c:title")))?;
    writer.write_event(Event::Start(BytesStart::new("c:tx")))?;
    writer.write_event(Event::Start(BytesStart::new("c:rich")))?;
    writer.write_event(Event::Empty(BytesStart::new("a:bodyPr")))?;
    writer.write_event(Event::Empty(BytesStart::new("a:lstStyle")))?;
    writer.write_event(Event::Start(BytesStart::new("a:p")))?;
    write_plain_text_run(writer, title)?;
    writer.write_event(Event::End(BytesEnd::new("a:p")))?;
    writer.write_event(Event::End(BytesEnd::new("c:rich")))?;
    writer.write_event(Event::End(BytesEnd::new("c:tx")))?;
    writer.write_event(Event::End(BytesEnd::new("c:title")))?;
    Ok(())
}

fn serialize_comment_authors_xml(authors: &[SerializedCommentAuthor]) -> Result<Vec<u8>> {
    let mut writer = Writer::new_with_indent(Vec::new(), b' ', 2);
    writer.write_event(Event::Decl(BytesDecl::new(
        "1.0",
        Some("UTF-8"),
        Some("yes"),
    )))?;

    let mut cm_authors = BytesStart::new("p:cmAuthorLst");
    cm_authors.push_attribute(("xmlns:p", PRESENTATIONML_NS));
    cm_authors.push_attribute(("xmlns:a", DRAWINGML_NS));
    writer.write_event(Event::Start(cm_authors))?;

    for author in authors {
        let mut cm_author = BytesStart::new("p:cmAuthor");
        let id_text = author.id.to_string();
        let last_idx_text = author.last_comment_index.to_string();
        cm_author.push_attribute(("id", id_text.as_str()));
        cm_author.push_attribute(("name", author.name.as_str()));
        cm_author.push_attribute(("initials", ""));
        cm_author.push_attribute(("lastIdx", last_idx_text.as_str()));
        writer.write_event(Event::Empty(cm_author))?;
    }

    writer.write_event(Event::End(BytesEnd::new("p:cmAuthorLst")))?;
    Ok(writer.into_inner())
}

fn serialize_comments_xml(comments: &[SerializedSlideComment]) -> Result<Vec<u8>> {
    let mut writer = Writer::new_with_indent(Vec::new(), b' ', 2);
    writer.write_event(Event::Decl(BytesDecl::new(
        "1.0",
        Some("UTF-8"),
        Some("yes"),
    )))?;

    let mut cm_lst = BytesStart::new("p:cmLst");
    cm_lst.push_attribute(("xmlns:p", PRESENTATIONML_NS));
    cm_lst.push_attribute(("xmlns:a", DRAWINGML_NS));
    writer.write_event(Event::Start(cm_lst))?;

    for comment in comments {
        let mut cm = BytesStart::new("p:cm");
        let author_id_text = comment.author_id.to_string();
        let idx_text = comment.comment_index.to_string();
        cm.push_attribute(("authorId", author_id_text.as_str()));
        cm.push_attribute(("idx", idx_text.as_str()));
        cm.push_attribute(("dt", "2024-01-01T00:00:00Z"));
        writer.write_event(Event::Start(cm))?;

        let mut pos = BytesStart::new("p:pos");
        pos.push_attribute(("x", "0"));
        pos.push_attribute(("y", "0"));
        writer.write_event(Event::Empty(pos))?;

        writer.write_event(Event::Start(BytesStart::new("p:text")))?;
        writer.write_event(Event::Text(BytesText::new(comment.text.as_str())))?;
        writer.write_event(Event::End(BytesEnd::new("p:text")))?;

        writer.write_event(Event::End(BytesEnd::new("p:cm")))?;
    }

    writer.write_event(Event::End(BytesEnd::new("p:cmLst")))?;
    Ok(writer.into_inner())
}

fn allocate_object_id(next_object_id: &mut u32) -> Result<u32> {
    let current = *next_object_id;
    *next_object_id = next_object_id.checked_add(1).ok_or_else(|| {
        PptxError::UnsupportedPackage("shape id overflow while serializing".to_string())
    })?;
    Ok(current)
}

fn write_shape_xml<W: std::io::Write>(
    writer: &mut Writer<W>,
    object_id: u32,
    shape: &Shape,
) -> Result<()> {
    let is_connector = shape.is_connector();
    let sp_tag_name = if is_connector { "p:cxnSp" } else { "p:sp" };
    let mut shape_tag = BytesStart::new(sp_tag_name);
    for (key, value) in shape.unknown_attrs() {
        shape_tag.push_attribute((key.as_str(), value.as_str()));
    }
    writer.write_event(Event::Start(shape_tag))?;

    let nv_tag_name = if is_connector {
        "p:nvCxnSpPr"
    } else {
        "p:nvSpPr"
    };
    writer.write_event(Event::Start(BytesStart::new(nv_tag_name)))?;
    let mut c_nv_pr = BytesStart::new("p:cNvPr");
    let object_id_text = object_id.to_string();
    c_nv_pr.push_attribute(("id", object_id_text.as_str()));
    c_nv_pr.push_attribute(("name", shape.name()));
    // Feature #4: Hidden.
    if shape.is_hidden() {
        c_nv_pr.push_attribute(("hidden", "1"));
    }
    // Feature #10: Alt text.
    if let Some(descr) = shape.alt_text() {
        c_nv_pr.push_attribute(("descr", descr));
    }
    if let Some(title) = shape.alt_text_title() {
        c_nv_pr.push_attribute(("title", title));
    }
    writer.write_event(Event::Empty(c_nv_pr))?;

    if is_connector {
        // Feature #9: Connector shape non-visual properties.
        let c_nv_cxn_sp_pr = BytesStart::new("p:cNvCxnSpPr");
        writer.write_event(Event::Start(c_nv_cxn_sp_pr))?;
        if let Some(start_conn) = shape.start_connection() {
            let mut st_cxn = BytesStart::new("a:stCxn");
            let id_text = start_conn.shape_id.to_string();
            let idx_text = start_conn.connection_point_index.to_string();
            st_cxn.push_attribute(("id", id_text.as_str()));
            st_cxn.push_attribute(("idx", idx_text.as_str()));
            writer.write_event(Event::Empty(st_cxn))?;
        }
        if let Some(end_conn) = shape.end_connection() {
            let mut end_cxn = BytesStart::new("a:endCxn");
            let id_text = end_conn.shape_id.to_string();
            let idx_text = end_conn.connection_point_index.to_string();
            end_cxn.push_attribute(("id", id_text.as_str()));
            end_cxn.push_attribute(("idx", idx_text.as_str()));
            writer.write_event(Event::Empty(end_cxn))?;
        }
        writer.write_event(Event::End(BytesEnd::new("p:cNvCxnSpPr")))?;
    } else {
        let mut c_nv_sp_pr = BytesStart::new("p:cNvSpPr");
        if matches!(shape.shape_type(), ShapeType::TextBox) {
            c_nv_sp_pr.push_attribute(("txBox", "1"));
        }
        writer.write_event(Event::Empty(c_nv_sp_pr))?;
    }

    if shape.placeholder_kind().is_some() || shape.placeholder_idx().is_some() {
        writer.write_event(Event::Start(BytesStart::new("p:nvPr")))?;

        let mut ph = BytesStart::new("p:ph");
        if let Some(placeholder_kind) = shape.placeholder_kind() {
            ph.push_attribute(("type", placeholder_kind));
        }
        let placeholder_idx_text = shape.placeholder_idx().map(|idx| idx.to_string());
        if let Some(ref placeholder_idx_text) = placeholder_idx_text {
            ph.push_attribute(("idx", placeholder_idx_text.as_str()));
        }
        writer.write_event(Event::Empty(ph))?;

        writer.write_event(Event::End(BytesEnd::new("p:nvPr")))?;
    } else {
        writer.write_event(Event::Empty(BytesStart::new("p:nvPr")))?;
    }

    writer.write_event(Event::End(BytesEnd::new(nv_tag_name)))?;

    let has_sp_pr = shape.geometry().is_some()
        || shape.preset_geometry().is_some()
        || shape.custom_geometry_raw().is_some()
        || shape.solid_fill_srgb().is_some()
        || shape.outline().is_some()
        || shape.gradient_fill().is_some()
        || shape.pattern_fill().is_some()
        || shape.picture_fill().is_some()
        || shape.is_no_fill()
        || shape.rotation().is_some()
        || shape.flip_h()
        || shape.flip_v()
        || shape.shadow().is_some()
        || shape.glow().is_some()
        || shape.reflection().is_some();

    if has_sp_pr {
        writer.write_event(Event::Start(BytesStart::new("p:spPr")))?;

        // Helper: build xfrm attributes (rotation, flipH, flipV).
        let needs_xfrm = shape.geometry().is_some()
            || shape.rotation().is_some()
            || shape.flip_h()
            || shape.flip_v();
        if needs_xfrm {
            let mut xfrm = BytesStart::new("a:xfrm");
            let rot_text = shape.rotation().map(|r| r.to_string());
            if let Some(ref rot_text) = rot_text {
                xfrm.push_attribute(("rot", rot_text.as_str()));
            }
            if shape.flip_h() {
                xfrm.push_attribute(("flipH", "1"));
            }
            if shape.flip_v() {
                xfrm.push_attribute(("flipV", "1"));
            }
            if let Some(geometry) = shape.geometry() {
                writer.write_event(Event::Start(xfrm))?;

                let mut off = BytesStart::new("a:off");
                let off_x_text = geometry.x().to_string();
                let off_y_text = geometry.y().to_string();
                off.push_attribute(("x", off_x_text.as_str()));
                off.push_attribute(("y", off_y_text.as_str()));
                writer.write_event(Event::Empty(off))?;

                let mut ext = BytesStart::new("a:ext");
                let ext_cx_text = geometry.cx().to_string();
                let ext_cy_text = geometry.cy().to_string();
                ext.push_attribute(("cx", ext_cx_text.as_str()));
                ext.push_attribute(("cy", ext_cy_text.as_str()));
                writer.write_event(Event::Empty(ext))?;

                writer.write_event(Event::End(BytesEnd::new("a:xfrm")))?;
            } else {
                // Rotation/flip without geometry: emit xfrm with just attributes.
                writer.write_event(Event::Empty(xfrm))?;
            }
        }

        // Geometry: custom geometry takes precedence over preset.
        if let Some(custom_geom) = shape.custom_geometry_raw() {
            custom_geom.write_to(writer)?;
        } else if let Some(preset_geometry) = shape.preset_geometry() {
            let mut prst_geom = BytesStart::new("a:prstGeom");
            prst_geom.push_attribute(("prst", preset_geometry));
            writer.write_event(Event::Start(prst_geom))?;
            // Replay avLst from original, or emit empty avLst.
            if let Some(avlst) = shape.preset_geometry_adjustments() {
                avlst.write_to(writer)?;
            } else {
                writer.write_event(Event::Empty(BytesStart::new("a:avLst")))?;
            }
            writer.write_event(Event::End(BytesEnd::new("a:prstGeom")))?;
        }

        // Fill (mutually exclusive).
        if shape.is_no_fill() {
            writer.write_event(Event::Empty(BytesStart::new("a:noFill")))?;
        } else if let Some(color) = shape.solid_fill_color() {
            // Prefer the full ShapeColor (supports both srgbClr and schemeClr with transforms).
            write_solid_fill_color_xml(writer, color)?;
        } else if let Some(solid_fill_srgb) = shape.solid_fill_srgb() {
            writer.write_event(Event::Start(BytesStart::new("a:solidFill")))?;
            let mut srgb_clr = BytesStart::new("a:srgbClr");
            srgb_clr.push_attribute(("val", solid_fill_srgb));
            // Alpha/opacity on solid fill color.
            if let Some(alpha) = shape.solid_fill_alpha() {
                let alpha_val = (alpha as u32) * 1000;
                let alpha_text = alpha_val.to_string();
                writer.write_event(Event::Start(srgb_clr))?;
                let mut alpha_elem = BytesStart::new("a:alpha");
                alpha_elem.push_attribute(("val", alpha_text.as_str()));
                writer.write_event(Event::Empty(alpha_elem))?;
                writer.write_event(Event::End(BytesEnd::new("a:srgbClr")))?;
            } else {
                writer.write_event(Event::Empty(srgb_clr))?;
            }
            writer.write_event(Event::End(BytesEnd::new("a:solidFill")))?;
        } else if let Some(gradient) = shape.gradient_fill() {
            // Feature #2: Gradient fill.
            write_gradient_fill_xml(writer, gradient)?;
        } else if let Some(pattern) = shape.pattern_fill() {
            // Feature #2: Pattern fill.
            write_pattern_fill_xml(writer, pattern)?;
        } else if let Some(picture_fill) = shape.picture_fill() {
            // Feature #1 (picture fill): blipFill in spPr.
            writer.write_event(Event::Start(BytesStart::new("a:blipFill")))?;
            let mut blip = BytesStart::new("a:blip");
            blip.push_attribute(("r:embed", picture_fill.relationship_id.as_str()));
            writer.write_event(Event::Empty(blip))?;
            // Write srcRect if crop is set
            if let Some(ref crop) = picture_fill.crop {
                let (l, t, r, b) = crop.to_pptx_format();
                if l != 0 || t != 0 || r != 0 || b != 0 {
                    let mut src_rect = BytesStart::new("a:srcRect");
                    if l != 0 {
                        src_rect.push_attribute(("l", l.to_string().as_str()));
                    }
                    if t != 0 {
                        src_rect.push_attribute(("t", t.to_string().as_str()));
                    }
                    if r != 0 {
                        src_rect.push_attribute(("r", r.to_string().as_str()));
                    }
                    if b != 0 {
                        src_rect.push_attribute(("b", b.to_string().as_str()));
                    }
                    writer.write_event(Event::Empty(src_rect))?;
                }
            }
            if picture_fill.stretch {
                writer.write_event(Event::Start(BytesStart::new("a:stretch")))?;
                writer.write_event(Event::Empty(BytesStart::new("a:fillRect")))?;
                writer.write_event(Event::End(BytesEnd::new("a:stretch")))?;
            }
            writer.write_event(Event::End(BytesEnd::new("a:blipFill")))?;
        }

        // Feature #1: Outline.
        if let Some(outline) = shape.outline() {
            write_outline_xml(writer, outline)?;
        }

        // Shape effects (effectLst).
        let has_effects =
            shape.shadow().is_some() || shape.glow().is_some() || shape.reflection().is_some();
        if has_effects {
            write_effect_list_xml(writer, shape)?;
        }

        writer.write_event(Event::End(BytesEnd::new("p:spPr")))?;
    }

    writer.write_event(Event::Start(BytesStart::new("p:txBody")))?;

    // bodyPr with text anchor, auto-fit, text direction, columns, insets, word wrap.
    let has_body_pr_attrs = shape.text_anchor().is_some()
        || shape.text_direction().is_some()
        || shape.text_columns().is_some()
        || shape.text_column_spacing().is_some()
        || shape.text_inset_left().is_some()
        || shape.text_inset_right().is_some()
        || shape.text_inset_top().is_some()
        || shape.text_inset_bottom().is_some()
        || shape.word_wrap().is_some()
        || shape.body_pr_rot().is_some()
        || shape.body_pr_rtl_col().is_some()
        || shape.body_pr_from_word_art().is_some()
        || shape.body_pr_force_aa().is_some()
        || shape.body_pr_compat_ln_spc().is_some()
        || !shape.body_pr_unknown_attrs().is_empty();
    let has_body_pr_children =
        shape.auto_fit().is_some() || !shape.body_pr_unknown_children().is_empty();
    if has_body_pr_attrs || has_body_pr_children {
        let mut body_pr = BytesStart::new("a:bodyPr");
        if let Some(anchor) = shape.text_anchor() {
            body_pr.push_attribute(("anchor", anchor.to_xml_anchor()));
            if anchor.is_centered() {
                body_pr.push_attribute(("anchorCtr", "1"));
            }
        }
        if let Some(direction) = shape.text_direction() {
            body_pr.push_attribute(("vert", direction.to_xml()));
        }
        if let Some(columns) = shape.text_columns() {
            let val = columns.to_string();
            body_pr.push_attribute(("numCol", val.as_str()));
        }
        if let Some(spacing) = shape.text_column_spacing() {
            let val = spacing.to_string();
            body_pr.push_attribute(("spcCol", val.as_str()));
        }
        if let Some(inset) = shape.text_inset_left() {
            let val = inset.to_string();
            body_pr.push_attribute(("lIns", val.as_str()));
        }
        if let Some(inset) = shape.text_inset_right() {
            let val = inset.to_string();
            body_pr.push_attribute(("rIns", val.as_str()));
        }
        if let Some(inset) = shape.text_inset_top() {
            let val = inset.to_string();
            body_pr.push_attribute(("tIns", val.as_str()));
        }
        if let Some(inset) = shape.text_inset_bottom() {
            let val = inset.to_string();
            body_pr.push_attribute(("bIns", val.as_str()));
        }
        if let Some(wrap) = shape.word_wrap() {
            body_pr.push_attribute(("wrap", if wrap { "square" } else { "none" }));
        }
        let body_pr_rot_text = shape.body_pr_rot().map(|r| r.to_string());
        if let Some(ref rot_text) = body_pr_rot_text {
            body_pr.push_attribute(("rot", rot_text.as_str()));
        }
        if let Some(rtl) = shape.body_pr_rtl_col() {
            body_pr.push_attribute(("rtlCol", if rtl { "1" } else { "0" }));
        }
        if let Some(fwa) = shape.body_pr_from_word_art() {
            body_pr.push_attribute(("fromWordArt", if fwa { "1" } else { "0" }));
        }
        if let Some(faa) = shape.body_pr_force_aa() {
            body_pr.push_attribute(("forceAA", if faa { "1" } else { "0" }));
        }
        if let Some(cls) = shape.body_pr_compat_ln_spc() {
            body_pr.push_attribute(("compatLnSpc", if cls { "1" } else { "0" }));
        }
        // Replay unknown bodyPr attributes.
        for (key, value) in shape.body_pr_unknown_attrs() {
            body_pr.push_attribute((key.as_str(), value.as_str()));
        }
        if has_body_pr_children {
            writer.write_event(Event::Start(body_pr))?;
            if let Some(auto_fit) = shape.auto_fit() {
                writer.write_event(Event::Empty(BytesStart::new(auto_fit.to_xml_tag())))?;
            }
            // Replay unknown bodyPr children.
            for node in shape.body_pr_unknown_children() {
                node.write_to(writer)?;
            }
            writer.write_event(Event::End(BytesEnd::new("a:bodyPr")))?;
        } else {
            writer.write_event(Event::Empty(body_pr))?;
        }
    } else {
        writer.write_event(Event::Empty(BytesStart::new("a:bodyPr")))?;
    }
    // lstStyle: prefer captured raw node, fall back to empty element.
    if let Some(lst_style) = shape.lst_style_raw() {
        lst_style.write_to(writer)?;
    } else {
        writer.write_event(Event::Empty(BytesStart::new("a:lstStyle")))?;
    }

    let paragraphs = shape.paragraphs();
    if paragraphs.is_empty() {
        writer.write_event(Event::Start(BytesStart::new("a:p")))?;
        writer.write_event(Event::End(BytesEnd::new("a:p")))?;
    } else {
        for paragraph in paragraphs {
            write_paragraph_xml(writer, paragraph)?;
        }
    }

    writer.write_event(Event::End(BytesEnd::new("p:txBody")))?;
    for node in shape.unknown_children() {
        node.write_to(writer)?;
    }
    writer.write_event(Event::End(BytesEnd::new(sp_tag_name)))?;

    Ok(())
}

/// Write an outline (`<a:ln>`) element.
/// Write shape effects (`<a:effectLst>`) element.
fn write_effect_list_xml<W: std::io::Write>(writer: &mut Writer<W>, shape: &Shape) -> Result<()> {
    writer.write_event(Event::Start(BytesStart::new("a:effectLst")))?;

    // Glow effect.
    if let Some(glow) = shape.glow() {
        let mut glow_el = BytesStart::new("a:glow");
        let rad_text = glow.radius.to_string();
        glow_el.push_attribute(("rad", rad_text.as_str()));
        let has_glow_color =
            glow.color_full.is_some() || !glow.color.is_empty() || glow.alpha.is_some();
        if !has_glow_color {
            writer.write_event(Event::Empty(glow_el))?;
        } else {
            writer.write_event(Event::Start(glow_el))?;
            if let Some(ref color) = glow.color_full {
                write_color_xml(writer, color)?;
            } else if !glow.color.is_empty() {
                let mut srgb_clr = BytesStart::new("a:srgbClr");
                srgb_clr.push_attribute(("val", glow.color.as_str()));
                if let Some(alpha) = glow.alpha {
                    writer.write_event(Event::Start(srgb_clr))?;
                    let mut alpha_el = BytesStart::new("a:alpha");
                    let alpha_val = ((alpha as u32) * 1000).to_string();
                    alpha_el.push_attribute(("val", alpha_val.as_str()));
                    writer.write_event(Event::Empty(alpha_el))?;
                    writer.write_event(Event::End(BytesEnd::new("a:srgbClr")))?;
                } else {
                    writer.write_event(Event::Empty(srgb_clr))?;
                }
            }
            writer.write_event(Event::End(BytesEnd::new("a:glow")))?;
        }
    }

    // Outer shadow effect.
    if let Some(shadow) = shape.shadow() {
        let mut outer_shdw = BytesStart::new("a:outerShdw");
        // Convert cartesian (dx, dy) back to polar (dist, dir).
        let dx = shadow.offset_x as f64;
        let dy = shadow.offset_y as f64;
        let dist = (dx * dx + dy * dy).sqrt() as i64;
        let dir = if dist > 0 {
            let rad = dy.atan2(dx);
            (rad * 180.0 / std::f64::consts::PI * 60000.0) as i64
        } else {
            0
        };
        let dist_text = dist.to_string();
        let dir_text = dir.to_string();
        let blur_text = shadow.blur_radius.to_string();
        outer_shdw.push_attribute(("blurRad", blur_text.as_str()));
        outer_shdw.push_attribute(("dist", dist_text.as_str()));
        outer_shdw.push_attribute(("dir", dir_text.as_str()));

        let has_shadow_color =
            shadow.color_full.is_some() || !shadow.color.is_empty() || shadow.alpha.is_some();
        if !has_shadow_color {
            writer.write_event(Event::Empty(outer_shdw))?;
        } else {
            writer.write_event(Event::Start(outer_shdw))?;
            if let Some(ref color) = shadow.color_full {
                write_color_xml(writer, color)?;
            } else if !shadow.color.is_empty() {
                let mut srgb_clr = BytesStart::new("a:srgbClr");
                srgb_clr.push_attribute(("val", shadow.color.as_str()));
                if let Some(alpha) = shadow.alpha {
                    writer.write_event(Event::Start(srgb_clr))?;
                    let mut alpha_el = BytesStart::new("a:alpha");
                    let alpha_val = ((alpha as u32) * 1000).to_string();
                    alpha_el.push_attribute(("val", alpha_val.as_str()));
                    writer.write_event(Event::Empty(alpha_el))?;
                    writer.write_event(Event::End(BytesEnd::new("a:srgbClr")))?;
                } else {
                    writer.write_event(Event::Empty(srgb_clr))?;
                }
            }
            writer.write_event(Event::End(BytesEnd::new("a:outerShdw")))?;
        }
    }

    // Reflection effect.
    if let Some(reflection) = shape.reflection() {
        let mut refl = BytesStart::new("a:reflection");
        let blur_text = reflection.blur_radius.to_string();
        refl.push_attribute(("blurRad", blur_text.as_str()));
        let dist_text = reflection.distance.to_string();
        refl.push_attribute(("dist", dist_text.as_str()));
        if let Some(dir) = reflection.direction {
            let dir_text = dir.to_string();
            refl.push_attribute(("dir", dir_text.as_str()));
        }
        if let Some(start_a) = reflection.start_alpha {
            let sta_text = ((start_a as u32) * 1000).to_string();
            refl.push_attribute(("stA", sta_text.as_str()));
        }
        if let Some(end_a) = reflection.end_alpha {
            let end_a_text = ((end_a as u32) * 1000).to_string();
            refl.push_attribute(("endA", end_a_text.as_str()));
        }
        writer.write_event(Event::Empty(refl))?;
    }

    writer.write_event(Event::End(BytesEnd::new("a:effectLst")))?;
    Ok(())
}

fn write_outline_xml<W: std::io::Write>(
    writer: &mut Writer<W>,
    outline: &ShapeOutline,
) -> Result<()> {
    let mut ln = BytesStart::new("a:ln");
    let width_text = outline.width_emu.map(|w| w.to_string());
    if let Some(ref width_text) = width_text {
        ln.push_attribute(("w", width_text.as_str()));
    }
    if let Some(ref compound_style) = outline.compound_style {
        ln.push_attribute(("cmpd", compound_style.to_xml()));
    }

    let has_children = outline.color_srgb.is_some()
        || outline.color.is_some()
        || outline.dash_style.is_some()
        || outline.head_arrow.is_some()
        || outline.tail_arrow.is_some();
    if has_children {
        writer.write_event(Event::Start(ln))?;

        if let Some(ref color) = outline.color {
            // Prefer full ShapeColor (supports schemeClr + transforms).
            write_solid_fill_color_xml(writer, color)?;
        } else if let Some(ref color) = outline.color_srgb {
            writer.write_event(Event::Start(BytesStart::new("a:solidFill")))?;
            let mut srgb_clr = BytesStart::new("a:srgbClr");
            srgb_clr.push_attribute(("val", color.as_str()));
            // Alpha/opacity on outline color.
            if let Some(alpha) = outline.alpha {
                let alpha_val = (alpha as u32) * 1000;
                let alpha_text = alpha_val.to_string();
                writer.write_event(Event::Start(srgb_clr))?;
                let mut alpha_elem = BytesStart::new("a:alpha");
                alpha_elem.push_attribute(("val", alpha_text.as_str()));
                writer.write_event(Event::Empty(alpha_elem))?;
                writer.write_event(Event::End(BytesEnd::new("a:srgbClr")))?;
            } else {
                writer.write_event(Event::Empty(srgb_clr))?;
            }
            writer.write_event(Event::End(BytesEnd::new("a:solidFill")))?;
        }

        if let Some(dash_style) = outline.dash_style {
            let mut prst_dash = BytesStart::new("a:prstDash");
            prst_dash.push_attribute(("val", dash_style.to_xml()));
            writer.write_event(Event::Empty(prst_dash))?;
        }

        // Line arrows.
        if let Some(ref head_arrow) = outline.head_arrow {
            let mut head_end = BytesStart::new("a:headEnd");
            head_end.push_attribute(("type", head_arrow.arrow_type.to_xml()));
            head_end.push_attribute(("w", head_arrow.width.to_xml()));
            head_end.push_attribute(("len", head_arrow.length.to_xml()));
            writer.write_event(Event::Empty(head_end))?;
        }
        if let Some(ref tail_arrow) = outline.tail_arrow {
            let mut tail_end = BytesStart::new("a:tailEnd");
            tail_end.push_attribute(("type", tail_arrow.arrow_type.to_xml()));
            tail_end.push_attribute(("w", tail_arrow.width.to_xml()));
            tail_end.push_attribute(("len", tail_arrow.length.to_xml()));
            writer.write_event(Event::Empty(tail_end))?;
        }

        writer.write_event(Event::End(BytesEnd::new("a:ln")))?;
    } else {
        writer.write_event(Event::Empty(ln))?;
    }

    Ok(())
}

/// Write a gradient fill (`<a:gradFill>`) element.
fn write_gradient_fill_xml<W: std::io::Write>(
    writer: &mut Writer<W>,
    gradient: &GradientFill,
) -> Result<()> {
    writer.write_event(Event::Start(BytesStart::new("a:gradFill")))?;

    writer.write_event(Event::Start(BytesStart::new("a:gsLst")))?;
    for stop in &gradient.stops {
        let mut gs = BytesStart::new("a:gs");
        let pos_text = stop.position.to_string();
        gs.push_attribute(("pos", pos_text.as_str()));
        writer.write_event(Event::Start(gs))?;

        if let Some(ref color) = stop.color {
            write_color_xml(writer, color)?;
        } else {
            let mut srgb_clr = BytesStart::new("a:srgbClr");
            srgb_clr.push_attribute(("val", stop.color_srgb.as_str()));
            writer.write_event(Event::Empty(srgb_clr))?;
        }

        writer.write_event(Event::End(BytesEnd::new("a:gs")))?;
    }
    writer.write_event(Event::End(BytesEnd::new("a:gsLst")))?;

    if let Some(GradientFillType::Linear) = gradient.fill_type {
        let mut lin = BytesStart::new("a:lin");
        let angle_text = gradient.linear_angle.unwrap_or(0).to_string();
        lin.push_attribute(("ang", angle_text.as_str()));
        lin.push_attribute(("scaled", "0"));
        writer.write_event(Event::Empty(lin))?;
    } else if let Some(ref fill_type) = gradient.fill_type {
        let mut path = BytesStart::new("a:path");
        path.push_attribute(("path", fill_type.to_xml()));
        writer.write_event(Event::Empty(path))?;
    }

    writer.write_event(Event::End(BytesEnd::new("a:gradFill")))?;
    Ok(())
}

/// Write a pattern fill (`<a:pattFill>`) element.
fn write_pattern_fill_xml<W: std::io::Write>(
    writer: &mut Writer<W>,
    pattern: &PatternFill,
) -> Result<()> {
    let mut patt_fill = BytesStart::new("a:pattFill");
    patt_fill.push_attribute(("prst", pattern.pattern_type.to_xml()));
    writer.write_event(Event::Start(patt_fill))?;

    if let Some(ref fg_color) = pattern.foreground_color {
        writer.write_event(Event::Start(BytesStart::new("a:fgClr")))?;
        write_color_xml(writer, fg_color)?;
        writer.write_event(Event::End(BytesEnd::new("a:fgClr")))?;
    } else if let Some(ref fg_color) = pattern.foreground_srgb {
        writer.write_event(Event::Start(BytesStart::new("a:fgClr")))?;
        let mut srgb_clr = BytesStart::new("a:srgbClr");
        srgb_clr.push_attribute(("val", fg_color.as_str()));
        writer.write_event(Event::Empty(srgb_clr))?;
        writer.write_event(Event::End(BytesEnd::new("a:fgClr")))?;
    }

    if let Some(ref bg_color) = pattern.background_color {
        writer.write_event(Event::Start(BytesStart::new("a:bgClr")))?;
        write_color_xml(writer, bg_color)?;
        writer.write_event(Event::End(BytesEnd::new("a:bgClr")))?;
    } else if let Some(ref bg_color) = pattern.background_srgb {
        writer.write_event(Event::Start(BytesStart::new("a:bgClr")))?;
        let mut srgb_clr = BytesStart::new("a:srgbClr");
        srgb_clr.push_attribute(("val", bg_color.as_str()));
        writer.write_event(Event::Empty(srgb_clr))?;
        writer.write_event(Event::End(BytesEnd::new("a:bgClr")))?;
    }

    writer.write_event(Event::End(BytesEnd::new("a:pattFill")))?;
    Ok(())
}

/// Write slide background element.
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
            write_gradient_fill_xml(writer, gradient)?;
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

/// Write header/footer element.
fn write_slide_header_footer_xml<W: std::io::Write>(
    writer: &mut Writer<W>,
    hf: &SlideHeaderFooter,
) -> Result<()> {
    let mut element = BytesStart::new("p:hf");
    if let Some(show_sld_num) = hf.show_slide_number {
        element.push_attribute(("sldNum", xml_bool_value(show_sld_num)));
    }
    if let Some(show_dt) = hf.show_date_time {
        element.push_attribute(("dt", xml_bool_value(show_dt)));
    }
    if let Some(show_hdr) = hf.show_header {
        element.push_attribute(("hdr", xml_bool_value(show_hdr)));
    }
    if let Some(show_ftr) = hf.show_footer {
        element.push_attribute(("ftr", xml_bool_value(show_ftr)));
    }
    writer.write_event(Event::Empty(element))?;
    Ok(())
}

/// Write a grouped shapes element (`<p:grpSp>`).
fn write_shape_group_xml<W: std::io::Write>(
    writer: &mut Writer<W>,
    next_object_id: &mut u32,
    group: &ShapeGroup,
) -> Result<()> {
    writer.write_event(Event::Start(BytesStart::new("p:grpSp")))?;

    writer.write_event(Event::Start(BytesStart::new("p:nvGrpSpPr")))?;
    let mut c_nv_pr = BytesStart::new("p:cNvPr");
    let object_id_text = allocate_object_id(next_object_id)?.to_string();
    c_nv_pr.push_attribute(("id", object_id_text.as_str()));
    c_nv_pr.push_attribute(("name", group.name()));
    writer.write_event(Event::Empty(c_nv_pr))?;
    writer.write_event(Event::Empty(BytesStart::new("p:cNvGrpSpPr")))?;
    writer.write_event(Event::Empty(BytesStart::new("p:nvPr")))?;
    writer.write_event(Event::End(BytesEnd::new("p:nvGrpSpPr")))?;

    writer.write_event(Event::Empty(BytesStart::new("p:grpSpPr")))?;

    for shape in group.shapes() {
        write_shape_xml(writer, allocate_object_id(next_object_id)?, shape)?;
    }

    writer.write_event(Event::End(BytesEnd::new("p:grpSp")))?;
    Ok(())
}

fn write_paragraph_xml<W: std::io::Write>(
    writer: &mut Writer<W>,
    paragraph: &ShapeParagraph,
) -> Result<()> {
    writer.write_event(Event::Start(BytesStart::new("a:p")))?;

    // Write paragraph properties if any are set.
    let props = paragraph.properties();
    let has_ppr = props.alignment.is_some()
        || props.level.is_some()
        || props.margin_left_emu.is_some()
        || props.margin_right_emu.is_some()
        || props.indent_emu.is_some();
    let has_old_spacing = props.line_spacing_pct.is_some()
        || props.line_spacing_pts.is_some()
        || props.space_before_pts.is_some()
        || props.space_after_pts.is_some();
    let has_typed_spacing =
        props.line_spacing.is_some() || props.space_before.is_some() || props.space_after.is_some();
    let has_spacing = has_old_spacing || has_typed_spacing;
    let has_bullet = props.bullet.style.is_some()
        || props.bullet.font_name.is_some()
        || props.bullet.size_percent.is_some()
        || props.bullet.color_srgb.is_some();
    let has_ppr_children = has_spacing || has_bullet;

    if has_ppr || has_ppr_children {
        let mut ppr = BytesStart::new("a:pPr");
        if let Some(alignment) = props.alignment {
            ppr.push_attribute(("algn", alignment.to_xml()));
        }
        let level_text = props.level.map(|v| v.to_string());
        if let Some(ref level_text) = level_text {
            ppr.push_attribute(("lvl", level_text.as_str()));
        }
        let mar_l_text = props.margin_left_emu.map(|v| v.to_string());
        if let Some(ref mar_l_text) = mar_l_text {
            ppr.push_attribute(("marL", mar_l_text.as_str()));
        }
        let mar_r_text = props.margin_right_emu.map(|v| v.to_string());
        if let Some(ref mar_r_text) = mar_r_text {
            ppr.push_attribute(("marR", mar_r_text.as_str()));
        }
        let indent_text = props.indent_emu.map(|v| v.to_string());
        if let Some(ref indent_text) = indent_text {
            ppr.push_attribute(("indent", indent_text.as_str()));
        }

        if has_ppr_children {
            writer.write_event(Event::Start(ppr))?;

            // Line spacing: prefer typed field, fall back to old fields.
            if let Some(ref ls) = props.line_spacing {
                writer.write_event(Event::Start(BytesStart::new("a:lnSpc")))?;
                let val_text = ls.value.to_string();
                match ls.unit {
                    LineSpacingUnit::Percent => {
                        let mut spc_pct = BytesStart::new("a:spcPct");
                        spc_pct.push_attribute(("val", val_text.as_str()));
                        writer.write_event(Event::Empty(spc_pct))?;
                    }
                    LineSpacingUnit::Points => {
                        let mut spc_pts = BytesStart::new("a:spcPts");
                        spc_pts.push_attribute(("val", val_text.as_str()));
                        writer.write_event(Event::Empty(spc_pts))?;
                    }
                }
                writer.write_event(Event::End(BytesEnd::new("a:lnSpc")))?;
            } else if props.line_spacing_pct.is_some() || props.line_spacing_pts.is_some() {
                writer.write_event(Event::Start(BytesStart::new("a:lnSpc")))?;
                if let Some(pct) = props.line_spacing_pct {
                    let mut spc_pct = BytesStart::new("a:spcPct");
                    let pct_text = pct.to_string();
                    spc_pct.push_attribute(("val", pct_text.as_str()));
                    writer.write_event(Event::Empty(spc_pct))?;
                } else if let Some(pts) = props.line_spacing_pts {
                    let mut spc_pts = BytesStart::new("a:spcPts");
                    let pts_text = pts.to_string();
                    spc_pts.push_attribute(("val", pts_text.as_str()));
                    writer.write_event(Event::Empty(spc_pts))?;
                }
                writer.write_event(Event::End(BytesEnd::new("a:lnSpc")))?;
            }

            // Space before: prefer typed field, fall back to old field.
            if let Some(ref sb) = props.space_before {
                writer.write_event(Event::Start(BytesStart::new("a:spcBef")))?;
                let val_text = sb.value.to_string();
                match sb.unit {
                    SpacingUnit::Percent => {
                        let mut spc_pct = BytesStart::new("a:spcPct");
                        spc_pct.push_attribute(("val", val_text.as_str()));
                        writer.write_event(Event::Empty(spc_pct))?;
                    }
                    SpacingUnit::Points => {
                        let mut spc_pts = BytesStart::new("a:spcPts");
                        spc_pts.push_attribute(("val", val_text.as_str()));
                        writer.write_event(Event::Empty(spc_pts))?;
                    }
                }
                writer.write_event(Event::End(BytesEnd::new("a:spcBef")))?;
            } else if let Some(pts) = props.space_before_pts {
                writer.write_event(Event::Start(BytesStart::new("a:spcBef")))?;
                let mut spc_pts = BytesStart::new("a:spcPts");
                let pts_text = pts.to_string();
                spc_pts.push_attribute(("val", pts_text.as_str()));
                writer.write_event(Event::Empty(spc_pts))?;
                writer.write_event(Event::End(BytesEnd::new("a:spcBef")))?;
            }

            // Space after: prefer typed field, fall back to old field.
            if let Some(ref sa) = props.space_after {
                writer.write_event(Event::Start(BytesStart::new("a:spcAft")))?;
                let val_text = sa.value.to_string();
                match sa.unit {
                    SpacingUnit::Percent => {
                        let mut spc_pct = BytesStart::new("a:spcPct");
                        spc_pct.push_attribute(("val", val_text.as_str()));
                        writer.write_event(Event::Empty(spc_pct))?;
                    }
                    SpacingUnit::Points => {
                        let mut spc_pts = BytesStart::new("a:spcPts");
                        spc_pts.push_attribute(("val", val_text.as_str()));
                        writer.write_event(Event::Empty(spc_pts))?;
                    }
                }
                writer.write_event(Event::End(BytesEnd::new("a:spcAft")))?;
            } else if let Some(pts) = props.space_after_pts {
                writer.write_event(Event::Start(BytesStart::new("a:spcAft")))?;
                let mut spc_pts = BytesStart::new("a:spcPts");
                let pts_text = pts.to_string();
                spc_pts.push_attribute(("val", pts_text.as_str()));
                writer.write_event(Event::Empty(spc_pts))?;
                writer.write_event(Event::End(BytesEnd::new("a:spcAft")))?;
            }

            // Feature #11: Bullet properties.
            if has_bullet {
                // Bullet font.
                if let Some(ref font_name) = props.bullet.font_name {
                    let mut bu_font = BytesStart::new("a:buFont");
                    bu_font.push_attribute(("typeface", font_name.as_str()));
                    writer.write_event(Event::Empty(bu_font))?;
                }

                // Bullet size as percentage of text size.
                if let Some(size_pct) = props.bullet.size_percent {
                    let mut bu_sz_pct = BytesStart::new("a:buSzPct");
                    let pct_text = size_pct.to_string();
                    bu_sz_pct.push_attribute(("val", pct_text.as_str()));
                    writer.write_event(Event::Empty(bu_sz_pct))?;
                }

                // Bullet color.
                if let Some(ref color) = props.bullet.color {
                    writer.write_event(Event::Start(BytesStart::new("a:buClr")))?;
                    write_color_xml(writer, color)?;
                    writer.write_event(Event::End(BytesEnd::new("a:buClr")))?;
                } else if let Some(ref color) = props.bullet.color_srgb {
                    writer.write_event(Event::Start(BytesStart::new("a:buClr")))?;
                    let mut srgb_clr = BytesStart::new("a:srgbClr");
                    srgb_clr.push_attribute(("val", color.as_str()));
                    writer.write_event(Event::Empty(srgb_clr))?;
                    writer.write_event(Event::End(BytesEnd::new("a:buClr")))?;
                }

                // Bullet style (buNone, buChar, buAutoNum).
                match &props.bullet.style {
                    Some(BulletStyle::None) => {
                        writer.write_event(Event::Empty(BytesStart::new("a:buNone")))?;
                    }
                    Some(BulletStyle::Char(ch)) => {
                        let mut bu_char = BytesStart::new("a:buChar");
                        bu_char.push_attribute(("char", ch.as_str()));
                        writer.write_event(Event::Empty(bu_char))?;
                    }
                    Some(BulletStyle::AutoNum(num_type)) => {
                        let mut bu_auto_num = BytesStart::new("a:buAutoNum");
                        bu_auto_num.push_attribute(("type", num_type.as_str()));
                        writer.write_event(Event::Empty(bu_auto_num))?;
                    }
                    None => {}
                }
            }

            writer.write_event(Event::End(BytesEnd::new("a:pPr")))?;
        } else {
            writer.write_event(Event::Empty(ppr))?;
        }
    }

    for run in paragraph.runs() {
        write_text_run_xml(writer, run)?;
    }

    writer.write_event(Event::End(BytesEnd::new("a:p")))?;
    Ok(())
}

fn write_table_xml<W: std::io::Write>(
    writer: &mut Writer<W>,
    object_id: u32,
    table_index: usize,
    table: &Table,
) -> Result<()> {
    writer.write_event(Event::Start(BytesStart::new("p:graphicFrame")))?;

    writer.write_event(Event::Start(BytesStart::new("p:nvGraphicFramePr")))?;
    let mut c_nv_pr = BytesStart::new("p:cNvPr");
    let object_id_text = object_id.to_string();
    let table_name = format!("Table {table_index}");
    c_nv_pr.push_attribute(("id", object_id_text.as_str()));
    c_nv_pr.push_attribute(("name", table_name.as_str()));
    writer.write_event(Event::Empty(c_nv_pr))?;
    writer.write_event(Event::Empty(BytesStart::new("p:cNvGraphicFramePr")))?;
    writer.write_event(Event::Empty(BytesStart::new("p:nvPr")))?;
    writer.write_event(Event::End(BytesEnd::new("p:nvGraphicFramePr")))?;

    // Write table position/size if set.
    if let Some((x, y, cx, cy)) = table.geometry() {
        writer.write_event(Event::Start(BytesStart::new("p:xfrm")))?;
        let mut off = BytesStart::new("a:off");
        off.push_attribute(("x", x.to_string().as_str()));
        off.push_attribute(("y", y.to_string().as_str()));
        writer.write_event(Event::Empty(off))?;
        let mut ext = BytesStart::new("a:ext");
        ext.push_attribute(("cx", cx.to_string().as_str()));
        ext.push_attribute(("cy", cy.to_string().as_str()));
        writer.write_event(Event::Empty(ext))?;
        writer.write_event(Event::End(BytesEnd::new("p:xfrm")))?;
    }

    writer.write_event(Event::Start(BytesStart::new("a:graphic")))?;
    let mut graphic_data = BytesStart::new("a:graphicData");
    graphic_data.push_attribute(("uri", TABLE_GRAPHIC_DATA_URI));
    writer.write_event(Event::Start(graphic_data))?;

    writer.write_event(Event::Start(BytesStart::new("a:tbl")))?;

    let mut tbl_pr = BytesStart::new("a:tblPr");
    tbl_pr.push_attribute(("firstRow", "1"));
    tbl_pr.push_attribute(("bandRow", "1"));
    writer.write_event(Event::Empty(tbl_pr))?;

    // Feature #6: Column widths from table data.
    writer.write_event(Event::Start(BytesStart::new("a:tblGrid")))?;
    let col_widths = table.column_widths_emu();
    for col_index in 0..table.cols() {
        let mut grid_col = BytesStart::new("a:gridCol");
        let w_text = col_widths.get(col_index).copied().unwrap_or(0).to_string();
        grid_col.push_attribute(("w", w_text.as_str()));
        writer.write_event(Event::Empty(grid_col))?;
    }
    writer.write_event(Event::End(BytesEnd::new("a:tblGrid")))?;

    let row_heights = table.row_heights_emu();
    for row_index in 0..table.rows() {
        let mut row = BytesStart::new("a:tr");
        // Feature #6: Row heights from table data.
        let h_text = row_heights.get(row_index).copied().unwrap_or(0).to_string();
        row.push_attribute(("h", h_text.as_str()));
        writer.write_event(Event::Start(row))?;

        for col_index in 0..table.cols() {
            let cell = table.cell(row_index, col_index);
            let mut tc = BytesStart::new("a:tc");
            // Feature #3 (table merged cells): write gridSpan, rowSpan, vMerge.
            let grid_span_text = cell.and_then(|c| c.grid_span()).map(|v| v.to_string());
            let row_span_text = cell.and_then(|c| c.row_span()).map(|v| v.to_string());
            if let Some(ref gs) = grid_span_text {
                tc.push_attribute(("gridSpan", gs.as_str()));
            }
            if let Some(ref rs) = row_span_text {
                tc.push_attribute(("rowSpan", rs.as_str()));
            }
            if cell.is_some_and(|c| c.is_v_merge()) {
                tc.push_attribute(("vMerge", "1"));
            }
            writer.write_event(Event::Start(tc))?;

            // Feature #5: Cell text with formatting.
            writer.write_event(Event::Start(BytesStart::new("a:txBody")))?;
            writer.write_event(Event::Empty(BytesStart::new("a:bodyPr")))?;
            writer.write_event(Event::Empty(BytesStart::new("a:lstStyle")))?;
            writer.write_event(Event::Start(BytesStart::new("a:p")))?;
            if let Some(cell) = cell {
                let cell_text = cell.text();
                if !cell_text.is_empty() {
                    // Check if cell has text formatting overrides.
                    let has_fmt = cell.bold().is_some()
                        || cell.italic().is_some()
                        || cell.font_size().is_some()
                        || cell.font_color().is_some()
                        || cell.font_color_srgb().is_some();
                    if has_fmt {
                        write_formatted_table_cell_run(writer, cell)?;
                    } else {
                        write_plain_text_run(writer, cell_text)?;
                    }
                }
            }
            writer.write_event(Event::End(BytesEnd::new("a:p")))?;
            writer.write_event(Event::End(BytesEnd::new("a:txBody")))?;

            // Feature #5: Cell properties (fill, borders).
            write_table_cell_properties(writer, cell)?;

            writer.write_event(Event::End(BytesEnd::new("a:tc")))?;
        }

        writer.write_event(Event::End(BytesEnd::new("a:tr")))?;
    }

    writer.write_event(Event::End(BytesEnd::new("a:tbl")))?;
    writer.write_event(Event::End(BytesEnd::new("a:graphicData")))?;
    writer.write_event(Event::End(BytesEnd::new("a:graphic")))?;
    writer.write_event(Event::End(BytesEnd::new("p:graphicFrame")))?;

    Ok(())
}

fn write_chart_graphic_frame_xml<W: std::io::Write>(
    writer: &mut Writer<W>,
    object_id: u32,
    chart_name: &str,
    relationship_id: &str,
) -> Result<()> {
    writer.write_event(Event::Start(BytesStart::new("p:graphicFrame")))?;

    writer.write_event(Event::Start(BytesStart::new("p:nvGraphicFramePr")))?;
    let mut c_nv_pr = BytesStart::new("p:cNvPr");
    let object_id_text = object_id.to_string();
    c_nv_pr.push_attribute(("id", object_id_text.as_str()));
    c_nv_pr.push_attribute(("name", chart_name));
    writer.write_event(Event::Empty(c_nv_pr))?;
    writer.write_event(Event::Empty(BytesStart::new("p:cNvGraphicFramePr")))?;
    writer.write_event(Event::Empty(BytesStart::new("p:nvPr")))?;
    writer.write_event(Event::End(BytesEnd::new("p:nvGraphicFramePr")))?;

    writer.write_event(Event::Start(BytesStart::new("a:graphic")))?;
    let mut graphic_data = BytesStart::new("a:graphicData");
    graphic_data.push_attribute(("uri", CHART_GRAPHIC_DATA_URI));
    writer.write_event(Event::Start(graphic_data))?;

    let mut chart = BytesStart::new("c:chart");
    chart.push_attribute(("r:id", relationship_id));
    writer.write_event(Event::Empty(chart))?;

    writer.write_event(Event::End(BytesEnd::new("a:graphicData")))?;
    writer.write_event(Event::End(BytesEnd::new("a:graphic")))?;
    writer.write_event(Event::End(BytesEnd::new("p:graphicFrame")))?;

    Ok(())
}

fn write_picture_xml<W: std::io::Write>(
    writer: &mut Writer<W>,
    object_id: u32,
    name: &str,
    relationship_id: &str,
) -> Result<()> {
    writer.write_event(Event::Start(BytesStart::new("p:pic")))?;

    writer.write_event(Event::Start(BytesStart::new("p:nvPicPr")))?;
    let mut c_nv_pr = BytesStart::new("p:cNvPr");
    let object_id_text = object_id.to_string();
    c_nv_pr.push_attribute(("id", object_id_text.as_str()));
    c_nv_pr.push_attribute(("name", name));
    writer.write_event(Event::Empty(c_nv_pr))?;
    writer.write_event(Event::Empty(BytesStart::new("p:cNvPicPr")))?;
    writer.write_event(Event::Empty(BytesStart::new("p:nvPr")))?;
    writer.write_event(Event::End(BytesEnd::new("p:nvPicPr")))?;

    writer.write_event(Event::Start(BytesStart::new("p:blipFill")))?;
    let mut blip = BytesStart::new("a:blip");
    blip.push_attribute(("r:embed", relationship_id));
    writer.write_event(Event::Empty(blip))?;
    writer.write_event(Event::Start(BytesStart::new("a:stretch")))?;
    writer.write_event(Event::Empty(BytesStart::new("a:fillRect")))?;
    writer.write_event(Event::End(BytesEnd::new("a:stretch")))?;
    writer.write_event(Event::End(BytesEnd::new("p:blipFill")))?;

    writer.write_event(Event::Start(BytesStart::new("p:spPr")))?;
    writer.write_event(Event::End(BytesEnd::new("p:spPr")))?;

    writer.write_event(Event::End(BytesEnd::new("p:pic")))?;

    Ok(())
}

fn write_text_run_xml<W: std::io::Write>(
    writer: &mut Writer<W>,
    run: &crate::text::TextRun,
) -> Result<()> {
    writer.write_event(Event::Start(BytesStart::new("a:r")))?;

    // Write run properties (<a:rPr>) if any are set.
    let props = run.properties();
    let has_attrs = props.bold.is_some()
        || props.italic.is_some()
        || props.underline.is_some()
        || props.strikethrough.is_some()
        || props.font_size.is_some()
        || props.language.is_some()
        || props.character_spacing.is_some()
        || props.kerning.is_some()
        || props.baseline.is_some();
    let has_children = props.font_color.is_some()
        || props.font_color_srgb.is_some()
        || props.font_name.is_some()
        || props.font_name_east_asian.is_some()
        || props.font_name_complex_script.is_some()
        || props.hyperlink_click_rid.is_some();

    if has_attrs || has_children {
        let mut rpr = BytesStart::new("a:rPr");
        if let Some(bold) = props.bold {
            rpr.push_attribute(("b", xml_bool_value(bold)));
        }
        if let Some(italic) = props.italic {
            rpr.push_attribute(("i", xml_bool_value(italic)));
        }
        if let Some(underline) = props.underline {
            rpr.push_attribute(("u", underline.to_xml()));
        }
        if let Some(strikethrough) = props.strikethrough {
            rpr.push_attribute(("strike", strikethrough.to_xml()));
        }
        let sz_text = props.font_size.map(|v| v.to_string());
        if let Some(ref sz_text) = sz_text {
            rpr.push_attribute(("sz", sz_text.as_str()));
        }
        if let Some(ref lang) = props.language {
            rpr.push_attribute(("lang", lang.as_str()));
        }
        // Character spacing.
        let spc_text = props.character_spacing.map(|v| v.to_string());
        if let Some(ref spc_text) = spc_text {
            rpr.push_attribute(("spc", spc_text.as_str()));
        }
        // Kerning.
        let kern_text = props.kerning.map(|v| v.to_string());
        if let Some(ref kern_text) = kern_text {
            rpr.push_attribute(("kern", kern_text.as_str()));
        }
        // Subscript/superscript baseline.
        let baseline_text = props.baseline.map(|v| v.to_string());
        if let Some(ref baseline_text) = baseline_text {
            rpr.push_attribute(("baseline", baseline_text.as_str()));
        }

        if has_children {
            writer.write_event(Event::Start(rpr))?;

            if let Some(ref color) = props.font_color {
                write_solid_fill_color_xml(writer, color)?;
            } else if let Some(ref color) = props.font_color_srgb {
                writer.write_event(Event::Start(BytesStart::new("a:solidFill")))?;
                let mut srgb_clr = BytesStart::new("a:srgbClr");
                srgb_clr.push_attribute(("val", color.as_str()));
                writer.write_event(Event::Empty(srgb_clr))?;
                writer.write_event(Event::End(BytesEnd::new("a:solidFill")))?;
            }

            if let Some(ref font_name) = props.font_name {
                let mut latin = BytesStart::new("a:latin");
                latin.push_attribute(("typeface", font_name.as_str()));
                writer.write_event(Event::Empty(latin))?;
            }

            if let Some(ref font_name) = props.font_name_east_asian {
                let mut ea = BytesStart::new("a:ea");
                ea.push_attribute(("typeface", font_name.as_str()));
                writer.write_event(Event::Empty(ea))?;
            }

            if let Some(ref font_name) = props.font_name_complex_script {
                let mut cs = BytesStart::new("a:cs");
                cs.push_attribute(("typeface", font_name.as_str()));
                writer.write_event(Event::Empty(cs))?;
            }

            // Feature #10: Hyperlink click with optional tooltip.
            if let Some(ref rid) = props.hyperlink_click_rid {
                let mut hlink_click = BytesStart::new("a:hlinkClick");
                hlink_click.push_attribute(("r:id", rid.as_str()));
                if let Some(ref tooltip) = props.hyperlink_tooltip {
                    hlink_click.push_attribute(("tooltip", tooltip.as_str()));
                }
                writer.write_event(Event::Empty(hlink_click))?;
            }

            writer.write_event(Event::End(BytesEnd::new("a:rPr")))?;
        } else {
            writer.write_event(Event::Empty(rpr))?;
        }
    }

    let text_value = run.text();
    let mut text = BytesStart::new("a:t");
    if text_value.starts_with(' ') || text_value.ends_with(' ') {
        text.push_attribute(("xml:space", "preserve"));
    }
    writer.write_event(Event::Start(text))?;
    writer.write_event(Event::Text(BytesText::new(text_value)))?;
    writer.write_event(Event::End(BytesEnd::new("a:t")))?;

    writer.write_event(Event::End(BytesEnd::new("a:r")))?;
    Ok(())
}

/// Write a plain-text run without formatting (used for table cells).
fn write_plain_text_run<W: std::io::Write>(writer: &mut Writer<W>, text_value: &str) -> Result<()> {
    writer.write_event(Event::Start(BytesStart::new("a:r")))?;

    let mut text = BytesStart::new("a:t");
    if text_value.starts_with(' ') || text_value.ends_with(' ') {
        text.push_attribute(("xml:space", "preserve"));
    }
    writer.write_event(Event::Start(text))?;
    writer.write_event(Event::Text(BytesText::new(text_value)))?;
    writer.write_event(Event::End(BytesEnd::new("a:t")))?;

    writer.write_event(Event::End(BytesEnd::new("a:r")))?;
    Ok(())
}

/// Write a text run with table cell formatting overrides.
fn write_formatted_table_cell_run<W: std::io::Write>(
    writer: &mut Writer<W>,
    cell: &TableCell,
) -> Result<()> {
    writer.write_event(Event::Start(BytesStart::new("a:r")))?;

    let has_attrs = cell.bold().is_some() || cell.italic().is_some() || cell.font_size().is_some();
    let has_children = cell.font_color().is_some() || cell.font_color_srgb().is_some();

    if has_attrs || has_children {
        let mut rpr = BytesStart::new("a:rPr");
        if let Some(bold) = cell.bold() {
            rpr.push_attribute(("b", xml_bool_value(bold)));
        }
        if let Some(italic) = cell.italic() {
            rpr.push_attribute(("i", xml_bool_value(italic)));
        }
        let sz_text = cell.font_size().map(|v| v.to_string());
        if let Some(ref sz_text) = sz_text {
            rpr.push_attribute(("sz", sz_text.as_str()));
        }

        if has_children {
            writer.write_event(Event::Start(rpr))?;
            if let Some(color) = cell.font_color() {
                write_solid_fill_color_xml(writer, color)?;
            } else if let Some(color) = cell.font_color_srgb() {
                writer.write_event(Event::Start(BytesStart::new("a:solidFill")))?;
                let mut srgb_clr = BytesStart::new("a:srgbClr");
                srgb_clr.push_attribute(("val", color));
                writer.write_event(Event::Empty(srgb_clr))?;
                writer.write_event(Event::End(BytesEnd::new("a:solidFill")))?;
            }
            writer.write_event(Event::End(BytesEnd::new("a:rPr")))?;
        } else {
            writer.write_event(Event::Empty(rpr))?;
        }
    }

    let text_value = cell.text();
    let mut text = BytesStart::new("a:t");
    if text_value.starts_with(' ') || text_value.ends_with(' ') {
        text.push_attribute(("xml:space", "preserve"));
    }
    writer.write_event(Event::Start(text))?;
    writer.write_event(Event::Text(BytesText::new(text_value)))?;
    writer.write_event(Event::End(BytesEnd::new("a:t")))?;

    writer.write_event(Event::End(BytesEnd::new("a:r")))?;
    Ok(())
}

/// Write a table cell border element.
fn write_cell_border_xml<W: std::io::Write>(
    writer: &mut Writer<W>,
    element_name: &str,
    border: &CellBorder,
) -> Result<()> {
    let mut ln = BytesStart::new(element_name);
    let width_text = border.width_emu.map(|w| w.to_string());
    if let Some(ref width_text) = width_text {
        ln.push_attribute(("w", width_text.as_str()));
    }

    if let Some(ref color) = border.color {
        writer.write_event(Event::Start(ln))?;
        write_solid_fill_color_xml(writer, color)?;
        writer.write_event(Event::End(BytesEnd::new(element_name)))?;
    } else if let Some(ref color) = border.color_srgb {
        writer.write_event(Event::Start(ln))?;
        writer.write_event(Event::Start(BytesStart::new("a:solidFill")))?;
        let mut srgb_clr = BytesStart::new("a:srgbClr");
        srgb_clr.push_attribute(("val", color.as_str()));
        writer.write_event(Event::Empty(srgb_clr))?;
        writer.write_event(Event::End(BytesEnd::new("a:solidFill")))?;
        writer.write_event(Event::End(BytesEnd::new(element_name)))?;
    } else {
        writer.write_event(Event::Empty(ln))?;
    }

    Ok(())
}

/// Write table cell properties (`<a:tcPr>`).
fn write_table_cell_properties<W: std::io::Write>(
    writer: &mut Writer<W>,
    cell: Option<&TableCell>,
) -> Result<()> {
    let Some(cell) = cell else {
        writer.write_event(Event::Empty(BytesStart::new("a:tcPr")))?;
        return Ok(());
    };

    let has_fill = cell.fill_color().is_some() || cell.fill_color_srgb().is_some();
    let has_borders = cell.borders().is_set();
    let has_attrs = cell.vertical_alignment().is_some()
        || cell.margin_left().is_some()
        || cell.margin_right().is_some()
        || cell.margin_top().is_some()
        || cell.margin_bottom().is_some()
        || cell.text_direction().is_some();

    if has_fill || has_borders || has_attrs {
        let mut tc_pr = BytesStart::new("a:tcPr");

        // Attributes on <a:tcPr>.
        if let Some(va) = cell.vertical_alignment() {
            tc_pr.push_attribute(("anchor", va.to_xml()));
        }
        let mar_l_text = cell.margin_left().map(|v| v.to_string());
        if let Some(ref ml) = mar_l_text {
            tc_pr.push_attribute(("marL", ml.as_str()));
        }
        let mar_r_text = cell.margin_right().map(|v| v.to_string());
        if let Some(ref mr) = mar_r_text {
            tc_pr.push_attribute(("marR", mr.as_str()));
        }
        let mar_t_text = cell.margin_top().map(|v| v.to_string());
        if let Some(ref mt) = mar_t_text {
            tc_pr.push_attribute(("marT", mt.as_str()));
        }
        let mar_b_text = cell.margin_bottom().map(|v| v.to_string());
        if let Some(ref mb) = mar_b_text {
            tc_pr.push_attribute(("marB", mb.as_str()));
        }
        if let Some(td) = cell.text_direction() {
            tc_pr.push_attribute(("vert", td.to_xml()));
        }

        if has_fill || has_borders {
            writer.write_event(Event::Start(tc_pr))?;

            // Borders (must come before fill in tcPr children per schema).
            if let Some(ref border) = cell.borders().left {
                if border.is_set() {
                    write_cell_border_xml(writer, "a:lnL", border)?;
                }
            }
            if let Some(ref border) = cell.borders().right {
                if border.is_set() {
                    write_cell_border_xml(writer, "a:lnR", border)?;
                }
            }
            if let Some(ref border) = cell.borders().top {
                if border.is_set() {
                    write_cell_border_xml(writer, "a:lnT", border)?;
                }
            }
            if let Some(ref border) = cell.borders().bottom {
                if border.is_set() {
                    write_cell_border_xml(writer, "a:lnB", border)?;
                }
            }
            if let Some(ref border) = cell.borders().diagonal_down {
                if border.is_set() {
                    write_cell_border_xml(writer, "a:lnTlToBr", border)?;
                }
            }
            if let Some(ref border) = cell.borders().diagonal_up {
                if border.is_set() {
                    write_cell_border_xml(writer, "a:lnBlToTr", border)?;
                }
            }

            // Cell fill.
            if let Some(color) = cell.fill_color() {
                write_solid_fill_color_xml(writer, color)?;
            } else if let Some(color) = cell.fill_color_srgb() {
                writer.write_event(Event::Start(BytesStart::new("a:solidFill")))?;
                let mut srgb_clr = BytesStart::new("a:srgbClr");
                srgb_clr.push_attribute(("val", color));
                writer.write_event(Event::Empty(srgb_clr))?;
                writer.write_event(Event::End(BytesEnd::new("a:solidFill")))?;
            }

            writer.write_event(Event::End(BytesEnd::new("a:tcPr")))?;
        } else {
            writer.write_event(Event::Empty(tc_pr))?;
        }
    } else {
        writer.write_event(Event::Empty(BytesStart::new("a:tcPr")))?;
    }

    Ok(())
}

fn transition_kind_to_xml_name(kind: &SlideTransitionKind) -> &str {
    match kind {
        SlideTransitionKind::Unspecified => "",
        SlideTransitionKind::Cut => "cut",
        SlideTransitionKind::Fade => "fade",
        SlideTransitionKind::Push => "push",
        SlideTransitionKind::Wipe => "wipe",
        SlideTransitionKind::Other(name) => name.as_str(),
    }
}

fn xml_bool_value(value: bool) -> &'static str {
    if value {
        "1"
    } else {
        "0"
    }
}

fn parse_shape_type(event: &BytesStart<'_>) -> ShapeType {
    if get_attribute_value(event, b"txBox")
        .as_deref()
        .and_then(parse_xml_bool)
        .unwrap_or(false)
    {
        ShapeType::TextBox
    } else {
        ShapeType::AutoShape
    }
}

fn shape_geometry_from_parts(
    offset: Option<(i64, i64)>,
    extents: Option<(i64, i64)>,
) -> Option<ShapeGeometry> {
    match (offset, extents) {
        (Some((x, y)), Some((cx, cy))) => Some(ShapeGeometry::new(x, y, cx, cy)),
        _ => None,
    }
}

fn parse_i64_attribute_value(event: &BytesStart<'_>, expected_local_name: &[u8]) -> Option<i64> {
    get_attribute_value(event, expected_local_name).and_then(|value| value.parse().ok())
}

fn parse_xml_bool(value: &str) -> Option<bool> {
    match value.trim().to_ascii_lowercase().as_str() {
        "1" | "true" | "on" => Some(true),
        "0" | "false" | "off" => Some(false),
        _ => None,
    }
}

#[cfg(test)]
fn parse_slide_text_runs(xml: &[u8]) -> Result<Vec<String>> {
    let shapes = parse_slide_shapes(xml)?;
    Ok(extract_legacy_text_runs(shapes.first()))
}

fn get_attribute_value(event: &BytesStart<'_>, expected_local_name: &[u8]) -> Option<String> {
    event.attributes().flatten().find_map(|attribute| {
        (local_name(attribute.key.as_ref()) == expected_local_name)
            .then(|| String::from_utf8_lossy(attribute.value.as_ref()).into_owned())
    })
}

fn get_exact_attribute_value(event: &BytesStart<'_>, expected_key: &[u8]) -> Option<String> {
    event.attributes().flatten().find_map(|attribute| {
        (attribute.key.as_ref() == expected_key)
            .then(|| String::from_utf8_lossy(attribute.value.as_ref()).into_owned())
    })
}

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

fn local_name(name: &[u8]) -> &[u8] {
    name.rsplit(|byte| *byte == b':').next().unwrap_or(name)
}

fn has_presentation_prefix(name: &[u8]) -> bool {
    name.starts_with(b"p:")
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

#[cfg(test)]
mod tests {
    use std::fs;
    use std::path::{Path, PathBuf};
    use std::result::Result as StdResult;

    use super::Presentation;
    use crate::shape::{ShapeGeometry, ShapeType};
    use crate::slide::Slide;
    use crate::timing::{SlideAnimationNode, SlideTiming};
    use crate::transition::{SlideTransition, SlideTransitionKind};
    use offidized_opc::content_types::ContentTypeValue;
    use offidized_opc::relationship::RelationshipType;

    #[test]
    fn add_slide_with_title_adds_slide() {
        let mut presentation = Presentation::new();
        presentation.add_slide_with_title("Q4 Results");

        assert_eq!(presentation.slides().len(), 1);
        assert_eq!(presentation.slides()[0].title(), "Q4 Results");
        assert_eq!(presentation.slide_count(), 1);
    }

    #[test]
    fn slide_helpers_support_mutation_and_removal() {
        let mut presentation = Presentation::new();
        let slide = presentation.add_slide();
        slide.set_title("Intro");
        slide.add_text_run("Agenda");
        presentation.add_slide_with_title("Summary");

        assert_eq!(presentation.slide_count(), 2);
        assert_eq!(presentation.slide(0).map(Slide::title), Some("Intro"));

        if let Some(slide_mut) = presentation.slide_mut(1) {
            slide_mut.set_title("Wrap up");
        }
        assert_eq!(presentation.slide(1).map(Slide::title), Some("Wrap up"));

        let removed = presentation.remove_slide(0);
        assert!(removed.is_some());
        assert_eq!(presentation.slide_count(), 1);
        assert!(presentation.remove_slide(8).is_none());
    }

    #[test]
    fn open_save_smoke_roundtrips_slides() {
        let tempdir = tempfile::tempdir().expect("tempdir should be created");
        let path = tempdir.path().join("smoke.pptx");

        let mut presentation = Presentation::new();
        let slide = presentation.add_slide_with_title("Intro");
        slide.add_text_run("Agenda");
        presentation.add_slide_with_title("Summary");
        presentation.save(&path).expect("save smoke pptx");

        let package = offidized_opc::Package::open(&path).expect("open saved package");
        assert!(package.get_part("/ppt/presentation.xml").is_some());
        assert!(package.get_part("/ppt/slides/slide1.xml").is_some());
        assert!(package.get_part("/ppt/slides/slide2.xml").is_some());
        assert!(package
            .relationships()
            .get_first_by_type(RelationshipType::WORKBOOK)
            .is_some());
        assert_eq!(
            package
                .content_types()
                .get_override("/ppt/presentation.xml"),
            Some(ContentTypeValue::PRESENTATION)
        );
        assert_eq!(
            package
                .content_types()
                .get_override("/ppt/slides/slide1.xml"),
            Some(ContentTypeValue::SLIDE)
        );

        let reopened = Presentation::open(&path).expect("open smoke pptx");
        assert_eq!(reopened.slides().len(), 2);
        assert_eq!(reopened.slides()[0].title(), "Intro");
        assert_eq!(reopened.slides()[0].text_runs().len(), 1);
        assert_eq!(reopened.slides()[0].text_runs()[0].text(), "Agenda");
        assert_eq!(reopened.slides()[1].title(), "Summary");
    }

    #[test]
    fn roundtrip_slide_master_layout_parts_and_relationships() {
        let tempdir = tempfile::tempdir().expect("tempdir should be created");
        let path = tempdir.path().join("master-layout.pptx");

        let mut presentation = Presentation::new();
        presentation.add_slide_with_title("Intro");
        presentation.add_slide_with_title("Summary");
        presentation.save(&path).expect("save master/layout pptx");

        let package = offidized_opc::Package::open(&path).expect("open master/layout package");
        let presentation_part = package
            .get_part(super::PRESENTATION_PART_URI)
            .expect("presentation part should exist");

        let presentation_to_master = presentation_part
            .relationships
            .get_first_by_type(RelationshipType::SLIDE_MASTER)
            .expect("presentation should reference slide master");
        assert_eq!(
            presentation_to_master.target,
            "slideMasters/slideMaster1.xml"
        );

        let slide_master_part = package
            .get_part(super::DEFAULT_SLIDE_MASTER_PART_URI)
            .expect("slide master part should exist");
        assert_eq!(
            slide_master_part.content_type.as_deref(),
            Some(ContentTypeValue::SLIDE_MASTER)
        );
        let master_to_layout = slide_master_part
            .relationships
            .get_first_by_type(RelationshipType::SLIDE_LAYOUT)
            .expect("slide master should reference slide layout");
        assert_eq!(master_to_layout.target, "../slideLayouts/slideLayout1.xml");

        let slide_layout_part = package
            .get_part(super::DEFAULT_SLIDE_LAYOUT_PART_URI)
            .expect("slide layout part should exist");
        assert_eq!(
            slide_layout_part.content_type.as_deref(),
            Some(ContentTypeValue::SLIDE_LAYOUT)
        );
        let layout_to_master = slide_layout_part
            .relationships
            .get_first_by_type(RelationshipType::SLIDE_MASTER)
            .expect("slide layout should reference slide master");
        assert_eq!(layout_to_master.target, "../slideMasters/slideMaster1.xml");

        for slide_part_path in ["/ppt/slides/slide1.xml", "/ppt/slides/slide2.xml"] {
            let slide_part = package
                .get_part(slide_part_path)
                .expect("slide part should exist");
            let slide_to_layout = slide_part
                .relationships
                .get_first_by_type(RelationshipType::SLIDE_LAYOUT)
                .expect("slide should reference slide layout");
            assert_eq!(slide_to_layout.target, "../slideLayouts/slideLayout1.xml");
        }

        let presentation_xml = String::from_utf8(presentation_part.data.as_bytes().to_vec())
            .expect("presentation xml utf8");
        assert!(presentation_xml.contains("<p:sldMasterIdLst>"));
        assert!(presentation_xml.contains("<p:sldMasterId"));

        let reopened = Presentation::open(&path).expect("open master/layout pptx");
        let slide_masters = reopened.slide_masters();
        assert_eq!(slide_masters.len(), 1);
        assert_eq!(
            slide_masters[0].part_uri(),
            super::DEFAULT_SLIDE_MASTER_PART_URI
        );
        assert!(slide_masters[0].relationship_id().starts_with("rId"));
        assert!(slide_masters[0].shapes().is_empty());
        assert_eq!(slide_masters[0].layouts().len(), 1);
        assert_eq!(
            slide_masters[0].layouts()[0].part_uri(),
            super::DEFAULT_SLIDE_LAYOUT_PART_URI
        );
        assert_eq!(slide_masters[0].layouts()[0].name(), None);
        assert_eq!(slide_masters[0].layouts()[0].layout_type(), Some("title"));
        assert_eq!(slide_masters[0].layouts()[0].r#type(), Some("title"));
        assert_eq!(slide_masters[0].layouts()[0].preserve(), Some(true));
        assert!(slide_masters[0].layouts()[0].shapes().is_empty());
    }

    #[test]
    fn parse_slide_master_and_layout_xml_extracts_metadata() {
        let master_xml = r#"
            <p:sldMaster xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
                         xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                         xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
              <p:cSld>
                <p:spTree>
                  <p:sp>
                    <p:nvSpPr>
                      <p:cNvPr id="2" name="Master Title Placeholder"/>
                      <p:cNvSpPr/>
                      <p:nvPr>
                        <p:ph type="title"/>
                      </p:nvPr>
                    </p:nvSpPr>
                    <p:txBody>
                      <a:bodyPr/>
                      <a:lstStyle/>
                      <a:p><a:r><a:t>Master heading</a:t></a:r></a:p>
                    </p:txBody>
                  </p:sp>
                </p:spTree>
              </p:cSld>
              <p:sldLayoutIdLst>
                <p:sldLayoutId id="2147483649" r:id="rId5"/>
              </p:sldLayoutIdLst>
            </p:sldMaster>
        "#;
        let parsed_master =
            super::parse_slide_master_xml(master_xml.as_bytes()).expect("parse master metadata");
        assert_eq!(parsed_master.layout_refs.len(), 1);
        assert_eq!(parsed_master.layout_refs[0].relationship_id, "rId5");
        assert_eq!(parsed_master.shapes.len(), 1);
        assert_eq!(parsed_master.shapes[0].name(), "Master Title Placeholder");
        assert_eq!(parsed_master.shapes[0].placeholder_kind(), Some("title"));
        assert_eq!(
            parsed_master.shapes[0].paragraphs()[0].runs()[0].text(),
            "Master heading"
        );

        let layout_xml = r#"
            <p:sldLayout xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
                         xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                         xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
                         name="Title and Content"
                         type="titleAndContent"
                         preserve="0">
              <p:cSld>
                <p:spTree>
                  <p:sp>
                    <p:nvSpPr>
                      <p:cNvPr id="2" name="Layout Body Placeholder"/>
                      <p:cNvSpPr/>
                      <p:nvPr>
                        <p:ph type="body" idx="1"/>
                      </p:nvPr>
                    </p:nvSpPr>
                    <p:txBody>
                      <a:bodyPr/>
                      <a:lstStyle/>
                      <a:p><a:r><a:t>Layout body</a:t></a:r></a:p>
                    </p:txBody>
                  </p:sp>
                </p:spTree>
              </p:cSld>
            </p:sldLayout>
        "#;
        let parsed_layout = super::parse_slide_layout_xml_metadata(layout_xml.as_bytes())
            .expect("parse layout metadata");
        assert_eq!(parsed_layout.name, Some("Title and Content".to_string()));
        assert_eq!(
            parsed_layout.layout_type,
            Some("titleAndContent".to_string())
        );
        assert_eq!(parsed_layout.preserve, Some(false));
        assert_eq!(parsed_layout.shapes.len(), 1);
        assert_eq!(parsed_layout.shapes[0].name(), "Layout Body Placeholder");
        assert_eq!(parsed_layout.shapes[0].placeholder_kind(), Some("body"));
        assert_eq!(parsed_layout.shapes[0].placeholder_idx(), Some(1));
        assert_eq!(
            parsed_layout.shapes[0].paragraphs()[0].runs()[0].text(),
            "Layout body"
        );
    }

    #[test]
    fn open_save_roundtrips_multiple_text_runs_in_order() {
        let tempdir = tempfile::tempdir().expect("tempdir should be created");
        let path = tempdir.path().join("multi-run.pptx");

        let mut presentation = Presentation::new();
        let slide = presentation.add_slide_with_title("Intro");
        slide.add_text_run("First bullet");
        slide.add_text_run("Second bullet");
        slide.add_text_run("Third bullet");
        presentation.save(&path).expect("save multi-run pptx");

        let reopened = Presentation::open(&path).expect("open multi-run pptx");
        assert_eq!(reopened.slide_count(), 1);

        let reopened_slide = reopened.slide(0).expect("slide 0 should exist");
        assert_eq!(reopened_slide.title(), "Intro");
        let run_texts: Vec<&str> = reopened_slide
            .text_runs()
            .iter()
            .map(|text_run| text_run.text())
            .collect();
        assert_eq!(
            run_texts,
            vec!["First bullet", "Second bullet", "Third bullet"]
        );
    }

    #[test]
    fn parse_slide_text_runs_preserves_empty_run_order() {
        let xml = r#"
            <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
                   xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
              <p:cSld>
                <p:spTree>
                  <p:sp>
                    <p:txBody>
                      <a:bodyPr />
                      <a:lstStyle />
                      <a:p>
                        <a:r><a:t>Title</a:t></a:r>
                        <a:r><a:t/></a:r>
                        <a:r><a:t>Body</a:t></a:r>
                      </a:p>
                    </p:txBody>
                  </p:sp>
                </p:spTree>
              </p:cSld>
            </p:sld>
        "#;

        let runs = super::parse_slide_text_runs(xml.as_bytes()).expect("parse text runs");
        assert_eq!(runs, vec!["Title", "", "Body"]);
    }

    #[test]
    fn parse_presentation_xml_ignores_non_p_namespace_slide_id_entries() {
        let xml = r#"
            <p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
                            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
                            xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main">
              <p:sldMasterIdLst>
                <p:sldMasterId id="2147483648" r:id="rId1"/>
              </p:sldMasterIdLst>
              <p:sldIdLst>
                <p:sldId id="256" r:id="rId2"/>
              </p:sldIdLst>
              <p:extLst>
                <p:ext uri="{521415D9-36F7-43E2-AB2F-B90AF26B5E84}">
                  <p14:sectionLst>
                    <p14:section name="Section 1" id="{A}">
                      <p14:sldIdLst>
                        <p14:sldId id="256"/>
                      </p14:sldIdLst>
                    </p14:section>
                  </p14:sectionLst>
                </p:ext>
              </p:extLst>
            </p:presentation>
        "#;

        let refs = super::parse_presentation_xml(xml.as_bytes()).expect("parse presentation refs");
        assert_eq!(refs.slide_refs.len(), 1);
        assert_eq!(refs.slide_refs[0].relationship_id, "rId2");
        assert_eq!(refs.slide_master_refs.len(), 1);
        assert_eq!(refs.slide_master_refs[0].relationship_id, "rId1");
    }

    #[test]
    fn parse_slide_shapes_extracts_geometry_and_fill_metadata() {
        let xml = r#"
            <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
                   xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
              <p:cSld>
                <p:spTree>
                  <p:sp>
                    <p:nvSpPr>
                      <p:cNvPr id="2" name="Rounded Box"/>
                      <p:cNvSpPr/>
                      <p:nvPr/>
                    </p:nvSpPr>
                    <p:spPr>
                      <a:xfrm>
                        <a:off x="1524000" y="1397000"/>
                        <a:ext cx="5486400" cy="3708400"/>
                      </a:xfrm>
                      <a:prstGeom prst="roundRect">
                        <a:avLst/>
                      </a:prstGeom>
                      <a:solidFill>
                        <a:srgbClr val="4472C4"/>
                      </a:solidFill>
                    </p:spPr>
                    <p:txBody>
                      <a:bodyPr/>
                      <a:lstStyle/>
                      <a:p><a:r><a:t>Hello</a:t></a:r></a:p>
                    </p:txBody>
                  </p:sp>
                </p:spTree>
              </p:cSld>
            </p:sld>
        "#;

        let shapes = super::parse_slide_shapes(xml.as_bytes()).expect("parse shapes");
        assert_eq!(shapes.len(), 1);
        let shape = &shapes[0];
        assert_eq!(shape.name(), "Rounded Box");
        assert_eq!(
            shape.geometry(),
            Some(ShapeGeometry::new(
                1_524_000, 1_397_000, 5_486_400, 3_708_400
            ))
        );
        assert_eq!(shape.preset_geometry(), Some("roundRect"));
        assert_eq!(shape.solid_fill_srgb(), Some("4472C4"));
        assert_eq!(shape.paragraphs()[0].runs()[0].text(), "Hello");
    }

    #[test]
    fn parse_modify_serialize_roundtrip_shape_geometry_and_fill_metadata() {
        let xml = r#"
            <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
                   xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
              <p:cSld>
                <p:spTree>
                  <p:sp>
                    <p:nvSpPr>
                      <p:cNvPr id="2" name="Source Shape"/>
                      <p:cNvSpPr/>
                      <p:nvPr/>
                    </p:nvSpPr>
                    <p:spPr>
                      <a:xfrm>
                        <a:off x="100" y="200"/>
                        <a:ext cx="300" cy="400"/>
                      </a:xfrm>
                      <a:prstGeom prst="rect">
                        <a:avLst/>
                      </a:prstGeom>
                      <a:solidFill>
                        <a:srgbClr val="112233"/>
                      </a:solidFill>
                    </p:spPr>
                    <p:txBody>
                      <a:bodyPr/>
                      <a:lstStyle/>
                      <a:p><a:r><a:t>Seed</a:t></a:r></a:p>
                    </p:txBody>
                  </p:sp>
                </p:spTree>
              </p:cSld>
            </p:sld>
        "#;

        let mut parsed_shapes = super::parse_slide_shapes(xml.as_bytes()).expect("parse shapes");
        assert_eq!(parsed_shapes.len(), 1);
        let mut parsed_shape = parsed_shapes.pop().expect("shape should exist");

        parsed_shape.set_geometry(ShapeGeometry::new(7_000, 8_000, 9_000, 10_000));
        parsed_shape.set_preset_geometry("ellipse");
        parsed_shape.set_solid_fill_srgb("00FF7F");
        parsed_shape.add_paragraph_with_text("Tail");

        let tempdir = tempfile::tempdir().expect("tempdir should be created");
        let path = tempdir.path().join("shape-geometry-roundtrip.pptx");
        let mut presentation = Presentation::new();
        let slide = presentation.add_slide_with_title("Intro");
        super::append_shape(slide, parsed_shape);
        presentation
            .save(&path)
            .expect("save geometry roundtrip pptx");

        let package = offidized_opc::Package::open(&path).expect("open geometry package");
        let slide_part = package
            .get_part("/ppt/slides/slide1.xml")
            .expect("slide part should exist");
        let slide_xml =
            String::from_utf8(slide_part.data.as_bytes().to_vec()).expect("slide xml utf8");
        assert!(slide_xml.contains(r#"<a:off x="7000" y="8000"/>"#));
        assert!(slide_xml.contains(r#"<a:ext cx="9000" cy="10000"/>"#));
        assert!(slide_xml.contains(r#"<a:prstGeom prst="ellipse">"#));
        assert!(slide_xml.contains(r#"<a:srgbClr val="00FF7F"/>"#));

        let reopened = Presentation::open(&path).expect("open geometry roundtrip pptx");
        let reopened_slide = reopened.slide(0).expect("slide 0 should exist");
        assert_eq!(reopened_slide.shape_count(), 1);

        let reopened_shape = &reopened_slide.shapes()[0];
        assert_eq!(reopened_shape.name(), "Source Shape");
        assert_eq!(
            reopened_shape.geometry(),
            Some(ShapeGeometry::new(7_000, 8_000, 9_000, 10_000))
        );
        assert_eq!(reopened_shape.preset_geometry(), Some("ellipse"));
        assert_eq!(reopened_shape.solid_fill_srgb(), Some("00FF7F"));
        assert_eq!(reopened_shape.paragraph_count(), 2);
        assert_eq!(reopened_shape.paragraphs()[0].runs()[0].text(), "Seed");
        assert_eq!(reopened_shape.paragraphs()[1].runs()[0].text(), "Tail");
    }

    #[test]
    fn roundtrip_shape_text_frames() {
        let tempdir = tempfile::tempdir().expect("tempdir should be created");
        let path = tempdir.path().join("shapes.pptx");

        let mut presentation = Presentation::new();
        let slide = presentation.add_slide_with_title("Intro");
        let shape = slide.add_shape("Summary box");
        shape.add_paragraph_with_text("First paragraph");
        let paragraph = shape.add_paragraph();
        paragraph.add_run("Second");
        paragraph.add_run(" line");

        presentation.save(&path).expect("save shape pptx");
        let reopened = Presentation::open(&path).expect("open shape pptx");
        let reopened_slide = reopened.slide(0).expect("slide 0 should exist");

        assert_eq!(reopened_slide.title(), "Intro");
        assert_eq!(reopened_slide.shape_count(), 1);
        let reopened_shape = &reopened_slide.shapes()[0];
        assert_eq!(reopened_shape.name(), "Summary box");
        assert_eq!(reopened_shape.paragraph_count(), 2);
        assert_eq!(
            reopened_shape.paragraphs()[0].runs()[0].text(),
            "First paragraph"
        );
        assert_eq!(reopened_shape.paragraphs()[1].runs()[0].text(), "Second");
        assert_eq!(reopened_shape.paragraphs()[1].runs()[1].text(), " line");
    }

    #[test]
    fn roundtrip_text_run_formatting() {
        use crate::shape::TextAlignment;
        use crate::text::{StrikethroughStyle, UnderlineStyle};

        let tempdir = tempfile::tempdir().expect("tempdir should be created");
        let path = tempdir.path().join("text-formatting.pptx");

        let mut presentation = Presentation::new();
        let slide = presentation.add_slide_with_title("Formatting Test");
        let shape = slide.add_shape("Styled Text");

        // Paragraph 1: bold red text with alignment.
        let para = shape.add_paragraph();
        para.set_alignment(TextAlignment::Center);
        para.set_level(1);
        let run = para.add_run("Bold Red");
        run.set_bold(true)
            .set_italic(false)
            .set_font_size(2400)
            .set_font_color("FF0000")
            .set_font_name("Arial")
            .set_language("en-US")
            .set_underline(UnderlineStyle::Single)
            .set_strikethrough(StrikethroughStyle::Double);

        // Paragraph 2: plain text, no formatting.
        shape.add_paragraph_with_text("Plain text");

        presentation.save(&path).expect("save formatted pptx");
        let reopened = Presentation::open(&path).expect("open formatted pptx");
        let slide = reopened.slide(0).expect("slide 0");
        assert_eq!(slide.shape_count(), 1);

        let shape = &slide.shapes()[0];
        assert_eq!(shape.paragraph_count(), 2);

        // Verify paragraph 1 formatting.
        let para = &shape.paragraphs()[0];
        assert_eq!(para.alignment(), Some(TextAlignment::Center));
        assert_eq!(para.level(), Some(1));
        assert_eq!(para.run_count(), 1);

        let run = &para.runs()[0];
        assert_eq!(run.text(), "Bold Red");
        assert!(run.is_bold());
        assert!(!run.is_italic());
        assert_eq!(run.font_size(), Some(2400));
        assert_eq!(run.font_color(), Some("FF0000"));
        assert_eq!(run.font_name(), Some("Arial"));
        assert_eq!(run.language(), Some("en-US"));
        assert_eq!(run.underline(), Some(UnderlineStyle::Single));
        assert_eq!(run.strikethrough(), Some(StrikethroughStyle::Double));

        // Verify paragraph 2 is plain.
        let para2 = &shape.paragraphs()[1];
        assert_eq!(para2.alignment(), None);
        assert_eq!(para2.level(), None);
        assert_eq!(para2.run_count(), 1);
        let run2 = &para2.runs()[0];
        assert_eq!(run2.text(), "Plain text");
        assert!(!run2.is_bold());
        assert!(!run2.is_italic());
        assert_eq!(run2.font_size(), None);
        assert_eq!(run2.font_color(), None);
    }

    #[test]
    fn roundtrip_paragraph_spacing_properties() {
        use crate::shape::TextAlignment;

        let tempdir = tempfile::tempdir().expect("tempdir should be created");
        let path = tempdir.path().join("para-spacing.pptx");

        let mut presentation = Presentation::new();
        let slide = presentation.add_slide_with_title("Spacing Test");
        let shape = slide.add_shape("Spaced Text");

        let para = shape.add_paragraph();
        {
            let props = para.properties_mut();
            props.alignment = Some(TextAlignment::Right);
            props.margin_left_emu = Some(914400);
            props.indent_emu = Some(-457200);
            props.line_spacing_pct = Some(150000);
            props.space_before_pts = Some(600);
            props.space_after_pts = Some(300);
        }
        para.add_run("Spaced content");

        presentation.save(&path).expect("save spacing pptx");
        let reopened = Presentation::open(&path).expect("open spacing pptx");
        let slide = reopened.slide(0).expect("slide 0");
        let shape = &slide.shapes()[0];
        let para = &shape.paragraphs()[0];
        let props = para.properties();

        assert_eq!(props.alignment, Some(TextAlignment::Right));
        assert_eq!(props.margin_left_emu, Some(914400));
        assert_eq!(props.indent_emu, Some(-457200));
        assert_eq!(props.line_spacing_pct, Some(150000));
        assert_eq!(props.space_before_pts, Some(600));
        assert_eq!(props.space_after_pts, Some(300));
        assert_eq!(para.runs()[0].text(), "Spaced content");
    }

    #[test]
    fn roundtrip_shape_placeholder_metadata() {
        let tempdir = tempfile::tempdir().expect("tempdir should be created");
        let path = tempdir.path().join("shape-placeholder-metadata.pptx");

        let mut presentation = Presentation::new();
        let slide = presentation.add_slide_with_title("Intro");
        let shape = slide.add_shape("Body Placeholder");
        shape.set_placeholder_kind("body");
        shape.set_placeholder_idx(7);
        shape.add_paragraph_with_text("Agenda");

        presentation
            .save(&path)
            .expect("save placeholder shape pptx");
        let package = offidized_opc::Package::open(&path).expect("open placeholder package");
        let slide_part = package
            .get_part("/ppt/slides/slide1.xml")
            .expect("slide part should exist");
        let slide_xml =
            String::from_utf8(slide_part.data.as_bytes().to_vec()).expect("slide xml utf8");
        assert!(slide_xml.contains(r#"<p:ph type="body" idx="7"/>"#));

        let reopened = Presentation::open(&path).expect("open placeholder shape pptx");
        let reopened_slide = reopened.slide(0).expect("slide 0 should exist");
        assert_eq!(reopened_slide.shape_count(), 1);

        let reopened_shape = &reopened_slide.shapes()[0];
        assert_eq!(reopened_shape.name(), "Body Placeholder");
        assert_eq!(reopened_shape.shape_type(), ShapeType::AutoShape);
        assert_eq!(reopened_shape.placeholder_kind(), Some("body"));
        assert_eq!(reopened_shape.placeholder_idx(), Some(7));
        assert_eq!(reopened_shape.paragraphs()[0].runs()[0].text(), "Agenda");
    }

    #[test]
    fn roundtrip_shape_textbox_metadata() {
        let tempdir = tempfile::tempdir().expect("tempdir should be created");
        let path = tempdir.path().join("shape-textbox-metadata.pptx");

        let mut presentation = Presentation::new();
        let slide = presentation.add_slide_with_title("Intro");
        let shape = slide.add_shape("Textbox");
        shape.set_shape_type(ShapeType::TextBox);
        shape.add_paragraph_with_text("Free text");

        presentation.save(&path).expect("save textbox shape pptx");
        let package = offidized_opc::Package::open(&path).expect("open textbox package");
        let slide_part = package
            .get_part("/ppt/slides/slide1.xml")
            .expect("slide part should exist");
        let slide_xml =
            String::from_utf8(slide_part.data.as_bytes().to_vec()).expect("slide xml utf8");
        assert!(slide_xml.contains(r#"<p:cNvSpPr txBox="1"/>"#));

        let reopened = Presentation::open(&path).expect("open textbox shape pptx");
        let reopened_slide = reopened.slide(0).expect("slide 0 should exist");
        assert_eq!(reopened_slide.shape_count(), 1);

        let reopened_shape = &reopened_slide.shapes()[0];
        assert_eq!(reopened_shape.name(), "Textbox");
        assert_eq!(reopened_shape.shape_type(), ShapeType::TextBox);
        assert_eq!(reopened_shape.placeholder_kind(), None);
        assert_eq!(reopened_shape.placeholder_idx(), None);
        assert_eq!(reopened_shape.paragraphs()[0].runs()[0].text(), "Free text");
    }

    #[test]
    fn roundtrip_tables_with_cell_text_grid() {
        let tempdir = tempfile::tempdir().expect("tempdir should be created");
        let path = tempdir.path().join("tables.pptx");

        let mut presentation = Presentation::new();
        let slide = presentation.add_slide_with_title("Table slide");
        let table = slide.add_table(2, 3);
        assert!(table.set_cell_text(0, 0, "Q1"));
        assert!(table.set_cell_text(0, 1, "Q2"));
        assert!(table.set_cell_text(1, 0, "100"));
        assert!(table.set_cell_text(1, 1, "120"));

        presentation.save(&path).expect("save table pptx");
        let reopened = Presentation::open(&path).expect("open table pptx");
        let reopened_slide = reopened.slide(0).expect("slide 0 should exist");

        assert_eq!(reopened_slide.table_count(), 1);
        let reopened_table = &reopened_slide.tables()[0];
        assert_eq!(reopened_table.rows(), 2);
        assert_eq!(reopened_table.cols(), 3);
        assert_eq!(reopened_table.cell_text(0, 0), Some("Q1"));
        assert_eq!(reopened_table.cell_text(0, 1), Some("Q2"));
        assert_eq!(reopened_table.cell_text(1, 0), Some("100"));
        assert_eq!(reopened_table.cell_text(1, 1), Some("120"));
        assert_eq!(reopened_table.cell_text(0, 2), Some(""));
    }

    #[test]
    fn roundtrip_slide_transition_metadata() {
        let tempdir = tempfile::tempdir().expect("tempdir should be created");
        let path = tempdir.path().join("transition.pptx");

        let mut presentation = Presentation::new();
        let slide = presentation.add_slide_with_title("Transition slide");
        let mut transition = SlideTransition::new(SlideTransitionKind::Fade);
        transition.set_advance_on_click(Some(false));
        transition.set_advance_after_ms(Some(1_500));
        slide.set_transition(transition.clone());

        presentation.save(&path).expect("save transition pptx");

        let package = offidized_opc::Package::open(&path).expect("open transition package");
        let slide_part = package
            .get_part("/ppt/slides/slide1.xml")
            .expect("slide part should exist");
        let slide_xml =
            String::from_utf8(slide_part.data.as_bytes().to_vec()).expect("slide xml utf8");
        assert!(slide_xml.contains(r#"<p:transition advClick="0" advTm="1500">"#));
        assert!(slide_xml.contains("<p:fade/>"));

        let reopened = Presentation::open(&path).expect("open transition pptx");
        let reopened_slide = reopened.slide(0).expect("slide 0 should exist");
        assert_eq!(reopened_slide.transition(), Some(&transition));
    }

    #[test]
    fn roundtrip_slide_timing_metadata() {
        let tempdir = tempfile::tempdir().expect("tempdir should be created");
        let path = tempdir.path().join("timing.pptx");

        let mut presentation = Presentation::new();
        let slide = presentation.add_slide_with_title("Timing slide");
        let timing_inner_xml = concat!(
            r#"<p:tnLst><p:par><p:cTn id="1" dur="indefinite" restart="never" nodeType="tmRoot" "#,
            r#"xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2007/7/12/main"/>"#,
            r#"</p:par></p:tnLst><p:bldLst/>"#,
        );
        let timing = SlideTiming::new(timing_inner_xml);
        slide.set_timing(timing.clone());

        presentation.save(&path).expect("save timing pptx");

        let package = offidized_opc::Package::open(&path).expect("open timing package");
        let slide_part = package
            .get_part("/ppt/slides/slide1.xml")
            .expect("slide part should exist");
        let slide_xml =
            String::from_utf8(slide_part.data.as_bytes().to_vec()).expect("slide xml utf8");
        assert!(slide_xml.contains("<p:timing>"));
        assert!(slide_xml.contains(r#"<p:cTn id="1" dur="indefinite" restart="never" nodeType="tmRoot" xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2007/7/12/main"/>"#));
        assert!(slide_xml.contains("<p:bldLst/>"));

        let reopened = Presentation::open(&path).expect("open timing pptx");
        let reopened_slide = reopened.slide(0).expect("slide 0 should exist");
        assert_eq!(reopened_slide.timing(), Some(&timing));
    }

    #[test]
    fn parse_slide_timing_extracts_typed_animation_nodes() {
        let xml = r#"
            <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
              <p:cSld/>
              <p:timing>
                <p:tnLst>
                  <p:par>
                    <p:cTn id="1" dur="indefinite" nodeType="tmRoot"/>
                  </p:par>
                  <p:par>
                    <p:cTn id="2" dur="450" nodeType="clickEffect" evtFilter="cancelBubble"/>
                  </p:par>
                </p:tnLst>
              </p:timing>
            </p:sld>
        "#;

        let timing = super::parse_slide_timing(xml.as_bytes())
            .expect("parse timing")
            .expect("timing should exist");
        assert_eq!(timing.animations().len(), 2);
        assert_eq!(timing.animations()[0].id(), 1);
        assert_eq!(timing.animations()[0].duration_ms(), None);
        assert_eq!(timing.animations()[0].trigger(), Some("tmRoot"));
        assert_eq!(timing.animations()[0].event(), None);
        assert_eq!(timing.animations()[1].id(), 2);
        assert_eq!(timing.animations()[1].duration_ms(), Some(450));
        assert_eq!(timing.animations()[1].trigger(), Some("clickEffect"));
        assert_eq!(timing.animations()[1].event(), Some("cancelBubble"));
    }

    #[test]
    fn mutate_typed_timeline_serializes_when_raw_inner_xml_empty() {
        let tempdir = tempfile::tempdir().expect("tempdir should be created");
        let path = tempdir.path().join("typed-timing.pptx");

        let mut presentation = Presentation::new();
        let slide = presentation.add_slide_with_title("Typed timing");
        let mut timing = SlideTiming::new("");

        let mut root = SlideAnimationNode::new(1);
        root.set_trigger(Some("tmRoot"));
        timing.animations_mut().push(root);

        let mut click_effect = SlideAnimationNode::new(2);
        click_effect.set_duration_ms(Some(750));
        click_effect.set_trigger(Some("clickEffect"));
        click_effect.set_event(Some("cancelBubble"));
        timing.animations_mut().push(click_effect);
        slide.set_timing(timing);

        presentation.save(&path).expect("save typed timing pptx");

        let package = offidized_opc::Package::open(&path).expect("open typed timing package");
        let slide_part = package
            .get_part("/ppt/slides/slide1.xml")
            .expect("slide part should exist");
        let slide_xml =
            String::from_utf8(slide_part.data.as_bytes().to_vec()).expect("slide xml utf8");
        assert!(slide_xml.contains("<p:timing>"));
        assert!(slide_xml.contains("<p:tnLst>"));
        assert!(slide_xml.contains(r#"<p:cTn id="1" nodeType="tmRoot"/>"#));
        assert!(slide_xml.contains(
            r#"<p:cTn id="2" dur="750" nodeType="clickEffect" evtFilter="cancelBubble"/>"#
        ));

        let reopened = Presentation::open(&path).expect("open typed timing pptx");
        let reopened_timing = reopened
            .slide(0)
            .and_then(|slide| slide.timing())
            .expect("slide timing should exist");
        assert_eq!(reopened_timing.animations().len(), 2);
        assert_eq!(reopened_timing.animations()[0].id(), 1);
        assert_eq!(reopened_timing.animations()[1].id(), 2);
        assert_eq!(reopened_timing.animations()[1].duration_ms(), Some(750));
        assert_eq!(
            reopened_timing.animations()[1].event(),
            Some("cancelBubble")
        );
    }

    #[test]
    fn roundtrip_timing_preserves_unknown_inner_xml() {
        let tempdir = tempfile::tempdir().expect("tempdir should be created");
        let path = tempdir.path().join("timing-unknown.pptx");
        let timing_inner_xml = concat!(
            r#"<p:tnLst><p:par><p:cTn id="1" dur="indefinite" nodeType="tmRoot"/></p:par></p:tnLst>"#,
            r#"<p:unknownTag foo="bar"><x:custom xmlns:x="urn:offidized:test">value</x:custom></p:unknownTag>"#,
        );

        let mut presentation = Presentation::new();
        let slide = presentation.add_slide_with_title("Timing unknown");
        slide.set_timing(SlideTiming::new(timing_inner_xml));
        presentation.save(&path).expect("save unknown timing pptx");

        let package = offidized_opc::Package::open(&path).expect("open unknown timing package");
        let slide_part = package
            .get_part("/ppt/slides/slide1.xml")
            .expect("slide part should exist");
        let slide_xml =
            String::from_utf8(slide_part.data.as_bytes().to_vec()).expect("slide xml utf8");
        assert!(slide_xml.contains(r#"<p:unknownTag foo="bar">"#));
        assert!(slide_xml.contains(r#"<x:custom xmlns:x="urn:offidized:test">value</x:custom>"#));

        let reopened = Presentation::open(&path).expect("open unknown timing pptx");
        let reopened_timing = reopened
            .slide(0)
            .and_then(|slide| slide.timing())
            .expect("slide timing should exist");
        assert_eq!(reopened_timing.raw_inner_xml(), timing_inner_xml);
        assert_eq!(reopened_timing.animations().len(), 1);
        assert_eq!(reopened_timing.animations()[0].id(), 1);
    }

    #[test]
    fn open_save_roundtrip_preserves_fixture_timing_inner_xml() {
        let workspace_root = workspace_root();
        let fixture_path = workspace_root.join(
            "references/Open-XML-SDK/test/DocumentFormat.OpenXml.Tests.Assets/assets/TestFiles/animation.pptx",
        );
        assert!(fixture_path.is_file(), "animation fixture should exist");

        let mut presentation = Presentation::open(&fixture_path).expect("open animation fixture");
        let original_timing_inner_xml = presentation
            .slide(0)
            .and_then(|slide| slide.timing())
            .map(|timing| timing.raw_inner_xml().to_string())
            .expect("fixture slide should have timing");

        let tempdir = tempfile::tempdir().expect("tempdir should be created");
        let path = tempdir.path().join("animation-roundtrip.pptx");
        presentation
            .save(&path)
            .expect("save animation fixture roundtrip");

        let package =
            offidized_opc::Package::open(&path).expect("open animation roundtrip package");
        let slide_part = package
            .get_part("/ppt/slides/slide1.xml")
            .expect("slide part should exist");
        let slide_xml =
            String::from_utf8(slide_part.data.as_bytes().to_vec()).expect("slide xml utf8");
        assert!(slide_xml.contains("<p:timing>"));

        let reopened = Presentation::open(&path).expect("open animation roundtrip pptx");
        let reopened_timing_inner_xml = reopened
            .slide(0)
            .and_then(|slide| slide.timing())
            .map(|timing| timing.raw_inner_xml().to_string())
            .expect("roundtripped slide should have timing");
        assert_eq!(reopened_timing_inner_xml, original_timing_inner_xml);
    }

    #[test]
    fn open_save_roundtrip_preserves_shapecrawler_chart_fixture_payload() {
        let workspace_root = workspace_root();
        let fixture_path = workspace_root.join(
            "references/ShapeCrawler/tests/ShapeCrawler.DevTests/Assets/charts/001 bar chart.pptx",
        );
        if !fixture_path.is_file() {
            eprintln!(
                "skipping test: ShapeCrawler chart fixture not found at `{}`",
                fixture_path.display()
            );
            return;
        }

        let expected_categories = [
            "Category 1",
            "Category 2",
            "Category 3",
            "Category 4",
            "Category 1",
            "Category 2",
            "Category 3",
            "Category 4",
            "Category 1",
            "Category 2",
            "Category 3",
            "Category 4",
        ];
        // chart.values() returns only the first series; additional series are
        // accessible via chart.series().
        let expected_values = [4.3, 2.5, 3.5, 4.5];

        let mut presentation =
            Presentation::open(&fixture_path).expect("open ShapeCrawler chart fixture");
        assert_eq!(presentation.slide_count(), 1);
        let opened_chart = presentation
            .slide(0)
            .and_then(|slide| slide.charts().first())
            .expect("fixture should contain one chart");
        assert_eq!(opened_chart.title(), "Bar Chart");
        assert_eq!(
            opened_chart
                .categories()
                .iter()
                .map(String::as_str)
                .collect::<Vec<_>>(),
            expected_categories
        );
        assert_eq!(opened_chart.values(), expected_values);

        let tempdir = tempfile::tempdir().expect("tempdir should be created");
        let path = tempdir.path().join("shapecrawler-chart-roundtrip.pptx");
        presentation
            .save(&path)
            .expect("save ShapeCrawler chart roundtrip");

        let reopened = Presentation::open(&path).expect("open ShapeCrawler chart roundtrip");
        let reopened_chart = reopened
            .slide(0)
            .and_then(|slide| slide.charts().first())
            .expect("roundtripped presentation should contain one chart");
        assert_eq!(reopened_chart.title(), "Bar Chart");
        assert_eq!(
            reopened_chart
                .categories()
                .iter()
                .map(String::as_str)
                .collect::<Vec<_>>(),
            expected_categories
        );
        assert_eq!(reopened_chart.values(), expected_values);
    }

    #[test]
    fn open_save_roundtrip_supports_namespaced_slide_relationship_ids_in_presentation_xml() {
        let workspace_root = workspace_root();
        let fixture_path = workspace_root
            .join("references/ShapeCrawler/tests/ShapeCrawler.DevTests/Assets/008.pptx");
        if !fixture_path.is_file() {
            eprintln!(
                "skipping test: ShapeCrawler fixture not found at `{}`",
                fixture_path.display()
            );
            return;
        }

        let mut presentation =
            Presentation::open(&fixture_path).expect("open namespaced slide id fixture");
        let expected_slide_count = presentation.slide_count();
        assert!(expected_slide_count > 0);

        let tempdir = tempfile::tempdir().expect("tempdir should be created");
        let path = tempdir.path().join("namespaced-slide-id-roundtrip.pptx");
        presentation
            .save(&path)
            .expect("save namespaced slide id roundtrip");

        let reopened =
            Presentation::open(&path).expect("reopen namespaced slide id roundtrip fixture");
        assert_eq!(reopened.slide_count(), expected_slide_count);
    }

    #[test]
    fn open_save_roundtrip_supports_comments_without_comment_authors_relationship() {
        let workspace_root = workspace_root();
        let fixture_path = workspace_root.join(
            "references/ShapeCrawler/tests/ShapeCrawler.DevTests/Assets/autoshapes/autoshape-case008_text-frame.pptx",
        );
        if !fixture_path.is_file() {
            eprintln!(
                "skipping test: ShapeCrawler comment fixture not found at `{}`",
                fixture_path.display()
            );
            return;
        }

        let mut presentation =
            Presentation::open(&fixture_path).expect("open comment-authorless fixture");
        let expected_slide_count = presentation.slide_count();
        assert!(expected_slide_count > 0);

        let tempdir = tempfile::tempdir().expect("tempdir should be created");
        let path = tempdir.path().join("comment-authorless-roundtrip.pptx");
        presentation
            .save(&path)
            .expect("save comment-authorless roundtrip");

        let reopened = Presentation::open(&path).expect("reopen comment-authorless roundtrip");
        assert_eq!(reopened.slide_count(), expected_slide_count);
    }

    #[test]
    fn roundtrip_notes_parts_relationships_and_text() {
        let tempdir = tempfile::tempdir().expect("tempdir should be created");
        let path = tempdir.path().join("notes.pptx");

        let mut presentation = Presentation::new();
        let slide = presentation.add_slide_with_title("Notes slide");
        slide.set_notes_text("Presenter note line 1\nPresenter note line 2");

        presentation.save(&path).expect("save notes pptx");

        let package = offidized_opc::Package::open(&path).expect("open notes package");
        let slide_part = package
            .get_part("/ppt/slides/slide1.xml")
            .expect("slide part should exist");
        let notes_relationship = slide_part
            .relationships
            .get_first_by_type(super::NOTES_SLIDE_RELATIONSHIP_TYPE)
            .expect("notes relationship should exist");
        assert_eq!(notes_relationship.target, "../notesSlides/notesSlide1.xml");

        let notes_part = package
            .get_part("/ppt/notesSlides/notesSlide1.xml")
            .expect("notes part should exist");
        assert_eq!(
            notes_part.content_type.as_deref(),
            Some(super::NOTES_SLIDE_CONTENT_TYPE)
        );
        let notes_xml =
            String::from_utf8(notes_part.data.as_bytes().to_vec()).expect("notes xml utf8");
        assert!(notes_xml.contains("Presenter note line 1"));
        assert!(notes_xml.contains("Presenter note line 2"));

        let reopened = Presentation::open(&path).expect("open notes pptx");
        let reopened_slide = reopened.slide(0).expect("slide 0 should exist");
        assert_eq!(
            reopened_slide.notes_text(),
            Some("Presenter note line 1\nPresenter note line 2")
        );
    }

    #[test]
    fn roundtrip_images_with_media_parts_and_relationships() {
        let tempdir = tempfile::tempdir().expect("tempdir should be created");
        let path = tempdir.path().join("images.pptx");

        let mut presentation = Presentation::new();
        let slide = presentation.add_slide_with_title("Image slide");
        slide.add_image(vec![0_u8, 1, 2, 3], "image/png");
        slide.add_image(vec![4_u8, 5, 6], "image/jpeg");

        presentation.save(&path).expect("save image pptx");

        let package = offidized_opc::Package::open(&path).expect("open image package");
        assert!(package.get_part("/ppt/media/image1.png").is_some());
        assert!(package.get_part("/ppt/media/image2.jpeg").is_some());
        assert_eq!(
            package
                .content_types()
                .get_override("/ppt/media/image1.png"),
            Some("image/png")
        );
        assert_eq!(
            package
                .content_types()
                .get_override("/ppt/media/image2.jpeg"),
            Some("image/jpeg")
        );

        let slide_part = package
            .get_part("/ppt/slides/slide1.xml")
            .expect("slide part should exist");
        let image_relationships = slide_part
            .relationships
            .get_by_type(RelationshipType::IMAGE);
        assert_eq!(image_relationships.len(), 2);
        assert_eq!(image_relationships[0].target, "../media/image1.png");
        assert_eq!(image_relationships[1].target, "../media/image2.jpeg");

        let reopened = Presentation::open(&path).expect("open image pptx");
        let reopened_slide = reopened.slide(0).expect("slide 0 should exist");
        assert_eq!(reopened_slide.image_count(), 2);

        let first = &reopened_slide.images()[0];
        assert_eq!(first.content_type(), "image/png");
        assert_eq!(first.bytes(), [0_u8, 1, 2, 3]);
        assert_eq!(first.relationship_id(), Some("rId1"));
        assert_eq!(first.name(), Some("Picture 1"));

        let second = &reopened_slide.images()[1];
        assert_eq!(second.content_type(), "image/jpeg");
        assert_eq!(second.bytes(), [4_u8, 5, 6]);
        assert_eq!(second.relationship_id(), Some("rId2"));
        assert_eq!(second.name(), Some("Picture 2"));
    }

    #[test]
    fn roundtrip_charts_with_chart_parts_relationships_and_data() {
        let tempdir = tempfile::tempdir().expect("tempdir should be created");
        let path = tempdir.path().join("charts.pptx");

        let mut presentation = Presentation::new();
        let slide_one = presentation.add_slide_with_title("Charts 1");
        let chart_one = slide_one.add_chart("Revenue");
        chart_one.add_data_point("Q1", 10.0);
        chart_one.add_data_point("Q2", 12.5);
        let chart_two = slide_one.add_chart("Profit");
        chart_two.add_data_point("Q1", 3.0);

        let slide_two = presentation.add_slide_with_title("Charts 2");
        let chart_three = slide_two.add_chart("Growth");
        chart_three.add_data_point("2024", 8.0);

        presentation.save(&path).expect("save chart pptx");

        let package = offidized_opc::Package::open(&path).expect("open chart package");
        for chart_part in [
            "/ppt/charts/chart1.xml",
            "/ppt/charts/chart2.xml",
            "/ppt/charts/chart3.xml",
        ] {
            assert!(package.get_part(chart_part).is_some());
            assert_eq!(
                package.content_types().get_override(chart_part),
                Some(super::CHART_CONTENT_TYPE)
            );
        }

        let slide1_part = package
            .get_part("/ppt/slides/slide1.xml")
            .expect("slide1 should exist");
        let slide1_chart_relationships = slide1_part
            .relationships
            .get_by_type(RelationshipType::CHART);
        assert_eq!(slide1_chart_relationships.len(), 2);
        assert_eq!(slide1_chart_relationships[0].target, "../charts/chart1.xml");
        assert_eq!(slide1_chart_relationships[1].target, "../charts/chart2.xml");
        let slide1_xml =
            String::from_utf8(slide1_part.data.as_bytes().to_vec()).expect("slide1 xml utf8");
        assert!(slide1_xml.contains(r#"<c:chart r:id="rId1"/>"#));
        assert!(slide1_xml.contains(r#"<c:chart r:id="rId2"/>"#));

        let slide2_part = package
            .get_part("/ppt/slides/slide2.xml")
            .expect("slide2 should exist");
        let slide2_chart_relationships = slide2_part
            .relationships
            .get_by_type(RelationshipType::CHART);
        assert_eq!(slide2_chart_relationships.len(), 1);
        assert_eq!(slide2_chart_relationships[0].target, "../charts/chart3.xml");

        let reopened = Presentation::open(&path).expect("open chart pptx");
        let reopened_slide_one = reopened.slide(0).expect("slide 0 should exist");
        assert_eq!(reopened_slide_one.chart_count(), 2);
        assert_eq!(reopened_slide_one.charts()[0].title(), "Revenue");
        assert_eq!(reopened_slide_one.charts()[0].categories(), ["Q1", "Q2"]);
        assert_eq!(reopened_slide_one.charts()[0].values(), [10.0, 12.5]);
        assert_eq!(reopened_slide_one.charts()[1].title(), "Profit");
        assert_eq!(reopened_slide_one.charts()[1].categories(), ["Q1"]);
        assert_eq!(reopened_slide_one.charts()[1].values(), [3.0]);

        let reopened_slide_two = reopened.slide(1).expect("slide 1 should exist");
        assert_eq!(reopened_slide_two.chart_count(), 1);
        assert_eq!(reopened_slide_two.charts()[0].title(), "Growth");
        assert_eq!(reopened_slide_two.charts()[0].categories(), ["2024"]);
        assert_eq!(reopened_slide_two.charts()[0].values(), [8.0]);
    }

    #[test]
    fn roundtrip_comments_with_comment_parts_and_authors() {
        let tempdir = tempfile::tempdir().expect("tempdir should be created");
        let path = tempdir.path().join("comments.pptx");

        let mut presentation = Presentation::new();
        let slide_one = presentation.add_slide_with_title("Comments 1");
        slide_one.add_comment("Alice", "Needs revision");
        slide_one.add_comment("Bob", "Looks good");

        let slide_two = presentation.add_slide_with_title("Comments 2");
        slide_two.add_comment("Alice", "Follow up next week");

        presentation.save(&path).expect("save comments pptx");

        let package = offidized_opc::Package::open(&path).expect("open comments package");
        let comment_authors_part = package
            .get_part(super::COMMENT_AUTHORS_PART_URI)
            .expect("comment authors part should exist");
        assert_eq!(
            comment_authors_part.content_type.as_deref(),
            Some(super::COMMENT_AUTHORS_CONTENT_TYPE)
        );
        let comment_authors_xml = String::from_utf8(comment_authors_part.data.as_bytes().to_vec())
            .expect("comment authors xml utf8");
        assert!(comment_authors_xml.contains(r#"name="Alice""#));
        assert!(comment_authors_xml.contains(r#"name="Bob""#));

        let comment1_part = package
            .get_part("/ppt/comments/comment1.xml")
            .expect("comment1 part should exist");
        let comment2_part = package
            .get_part("/ppt/comments/comment2.xml")
            .expect("comment2 part should exist");
        assert_eq!(
            comment1_part.content_type.as_deref(),
            Some(super::COMMENTS_CONTENT_TYPE)
        );
        assert_eq!(
            comment2_part.content_type.as_deref(),
            Some(super::COMMENTS_CONTENT_TYPE)
        );

        let slide1_part = package
            .get_part("/ppt/slides/slide1.xml")
            .expect("slide1 part should exist");
        let slide1_comment_relationships = slide1_part
            .relationships
            .get_by_type(super::COMMENTS_RELATIONSHIP_TYPE);
        assert_eq!(slide1_comment_relationships.len(), 1);
        assert_eq!(
            slide1_comment_relationships[0].target,
            "../comments/comment1.xml"
        );
        let slide1_author_relationships = slide1_part
            .relationships
            .get_by_type(super::COMMENT_AUTHORS_RELATIONSHIP_TYPE);
        assert_eq!(slide1_author_relationships.len(), 1);
        assert_eq!(
            slide1_author_relationships[0].target,
            "../commentAuthors.xml"
        );

        let slide2_part = package
            .get_part("/ppt/slides/slide2.xml")
            .expect("slide2 part should exist");
        let slide2_comment_relationships = slide2_part
            .relationships
            .get_by_type(super::COMMENTS_RELATIONSHIP_TYPE);
        assert_eq!(slide2_comment_relationships.len(), 1);
        assert_eq!(
            slide2_comment_relationships[0].target,
            "../comments/comment2.xml"
        );
        let slide2_author_relationships = slide2_part
            .relationships
            .get_by_type(super::COMMENT_AUTHORS_RELATIONSHIP_TYPE);
        assert_eq!(slide2_author_relationships.len(), 1);
        assert_eq!(
            slide2_author_relationships[0].target,
            "../commentAuthors.xml"
        );

        let reopened = Presentation::open(&path).expect("open comments pptx");
        let reopened_slide_one = reopened.slide(0).expect("slide 0 should exist");
        assert_eq!(reopened_slide_one.comment_count(), 2);
        assert_eq!(reopened_slide_one.comments()[0].author(), "Alice");
        assert_eq!(reopened_slide_one.comments()[0].text(), "Needs revision");
        assert_eq!(reopened_slide_one.comments()[1].author(), "Bob");
        assert_eq!(reopened_slide_one.comments()[1].text(), "Looks good");

        let reopened_slide_two = reopened.slide(1).expect("slide 1 should exist");
        assert_eq!(reopened_slide_two.comment_count(), 1);
        assert_eq!(reopened_slide_two.comments()[0].author(), "Alice");
        assert_eq!(
            reopened_slide_two.comments()[0].text(),
            "Follow up next week"
        );
    }

    const REFERENCE_CORPUS_PPTX_FIXTURES: &[&str] = &[
        "references/ShapeCrawler/tests/ShapeCrawler.DevTests/Assets/007_2 slides.pptx",
        "references/ShapeCrawler/tests/ShapeCrawler.DevTests/Assets/056_slide-notes.pptx",
        "references/ShapeCrawler/tests/ShapeCrawler.DevTests/Assets/065 table.pptx",
        "references/ShapeCrawler/tests/ShapeCrawler.DevTests/Assets/078 textbox.pptx",
        "references/ShapeCrawler/tests/ShapeCrawler.DevTests/Assets/charts/001 bar chart.pptx",
        "references/ShapeCrawler/tests/ShapeCrawler.DevTests/Assets/pictures/pictures-case001.pptx",
        "references/ShapeCrawler/tests/ShapeCrawler.DevTests/Assets/autoshapes/autoshape-case005_text-frame.pptx",
        "references/Open-XML-SDK/test/DocumentFormat.OpenXml.Tests.Assets/assets/TestFiles/animation.pptx",
    ];
    const BROADER_REFERENCE_CORPUS_ROOTS: &[&str] = &[
        "references/ShapeCrawler/tests/ShapeCrawler.DevTests/Assets",
        "references/Open-XML-SDK/test/DocumentFormat.OpenXml.Tests.Assets/assets/TestFiles",
    ];
    const BROADER_REFERENCE_CORPUS_MIN_SUCCESS_PERCENT: usize = 65;
    const BROADER_REFERENCE_CORPUS_FAILURE_LOG_LIMIT: usize = 25;
    const KNOWN_UNSUPPORTED_BROADER_SWEEP_FIXTURES: &[&str] = &["encrypted_pptx.pptx"];

    #[derive(Debug, Clone, PartialEq, Eq)]
    struct RelationshipSnapshot {
        id: String,
        rel_type: String,
        target: String,
        target_mode_is_external: bool,
    }

    #[derive(Debug, Clone, PartialEq, Eq)]
    struct PartSnapshot {
        uri: String,
        content_type: Option<String>,
        relationships: Vec<RelationshipSnapshot>,
        data: Vec<u8>,
    }

    #[derive(Debug, Clone, PartialEq, Eq)]
    struct PackageSnapshot {
        package_relationships: Vec<RelationshipSnapshot>,
        parts: Vec<PartSnapshot>,
    }

    type TestResult<T> = StdResult<T, String>;

    #[test]
    fn golden_roundtrip_reference_corpus_is_deterministic() {
        let workspace_root = workspace_root();
        let mut failures = Vec::new();
        let mut passed = 0usize;

        for relative_fixture_path in REFERENCE_CORPUS_PPTX_FIXTURES {
            let fixture_path = workspace_root.join(relative_fixture_path);
            match assert_fixture_roundtrip_is_deterministic(&fixture_path) {
                Ok(()) => passed += 1,
                Err(error) => failures.push(format!(
                    "fixture `{}` failed: {error}",
                    fixture_path.display()
                )),
            }
        }

        if !failures.is_empty() {
            let failure_report = failures.join("\n");
            panic!(
                "reference corpus roundtrip regressions:\n{failure_report}\n{} passed / {} total",
                passed,
                REFERENCE_CORPUS_PPTX_FIXTURES.len()
            );
        }
    }

    #[test]
    #[ignore = "broader corpus sweep is expensive; run with --ignored --nocapture"]
    fn broader_reference_corpus_roundtrip_sweep_meets_success_threshold() {
        let workspace_root = workspace_root();
        let mut fixture_paths = Vec::new();
        for relative_root in BROADER_REFERENCE_CORPUS_ROOTS {
            let fixture_root = workspace_root.join(relative_root);
            let mut discovered =
                discover_pptx_fixtures_recursive(&fixture_root).unwrap_or_else(|error| {
                    panic!(
                        "discover fixtures under `{}`: {error}",
                        fixture_root.display()
                    )
                });
            fixture_paths.append(&mut discovered);
        }

        fixture_paths.sort();
        fixture_paths.dedup();
        assert!(
            !fixture_paths.is_empty(),
            "broader corpus sweep discovered zero fixtures"
        );

        let total = fixture_paths.len();
        let required_successes =
            (total * BROADER_REFERENCE_CORPUS_MIN_SUCCESS_PERCENT).div_ceil(100usize);
        let mut passed = 0usize;
        let mut failures = Vec::new();

        for fixture_path in &fixture_paths {
            match assert_fixture_roundtrip_is_deterministic(fixture_path) {
                Ok(()) => passed += 1,
                Err(error) => failures.push(format!("{}: {error}", fixture_path.display())),
            }
        }

        let failed = failures.len();
        println!(
            "broader corpus sweep summary: passed={passed} failed={failed} total={total} threshold={}%(required_successes={required_successes})",
            BROADER_REFERENCE_CORPUS_MIN_SUCCESS_PERCENT
        );
        if !failures.is_empty() {
            println!(
                "broader corpus sweep failures (showing up to {}):",
                BROADER_REFERENCE_CORPUS_FAILURE_LOG_LIMIT
            );
            for failure in failures
                .iter()
                .take(BROADER_REFERENCE_CORPUS_FAILURE_LOG_LIMIT)
            {
                println!(" - {failure}");
            }
            if failures.len() > BROADER_REFERENCE_CORPUS_FAILURE_LOG_LIMIT {
                println!(
                    " - ... {} additional failures omitted",
                    failures.len() - BROADER_REFERENCE_CORPUS_FAILURE_LOG_LIMIT
                );
            }
        }

        assert!(
            passed >= required_successes,
            "broader corpus sweep success threshold not met: passed={passed} failed={failed} total={total} threshold={}%(required_successes={required_successes})",
            BROADER_REFERENCE_CORPUS_MIN_SUCCESS_PERCENT
        );
    }

    fn assert_fixture_roundtrip_is_deterministic(fixture_path: &Path) -> TestResult<()> {
        if !fixture_path.is_file() {
            return Err("fixture file is missing".to_string());
        }

        let tempdir = tempfile::tempdir().map_err(|error| format!("create tempdir: {error}"))?;
        let first_roundtrip_path = tempdir.path().join("first-roundtrip.pptx");
        let second_roundtrip_path = tempdir.path().join("second-roundtrip.pptx");

        let mut original_presentation =
            Presentation::open(fixture_path).map_err(|error| format!("open fixture: {error}"))?;
        if original_presentation.slide_count() == 0 {
            return Err("fixture opened with zero slides".to_string());
        }

        original_presentation
            .save(&first_roundtrip_path)
            .map_err(|error| format!("save first roundtrip: {error}"))?;

        let mut first_roundtrip_presentation = Presentation::open(&first_roundtrip_path)
            .map_err(|error| format!("open first roundtrip: {error}"))?;
        if first_roundtrip_presentation.slide_count() != original_presentation.slide_count() {
            return Err(format!(
                "slide count changed after first roundtrip: {} -> {}",
                original_presentation.slide_count(),
                first_roundtrip_presentation.slide_count()
            ));
        }

        first_roundtrip_presentation
            .save(&second_roundtrip_path)
            .map_err(|error| format!("save second roundtrip: {error}"))?;

        let second_roundtrip_presentation = Presentation::open(&second_roundtrip_path)
            .map_err(|error| format!("open second roundtrip: {error}"))?;
        if second_roundtrip_presentation.slide_count() != first_roundtrip_presentation.slide_count()
        {
            return Err(format!(
                "slide count changed after second roundtrip: {} -> {}",
                first_roundtrip_presentation.slide_count(),
                second_roundtrip_presentation.slide_count()
            ));
        }

        let first_snapshot = package_snapshot(&first_roundtrip_path)?;
        let second_snapshot = package_snapshot(&second_roundtrip_path)?;
        if first_snapshot != second_snapshot {
            return Err(format!(
                "package snapshot mismatch after second roundtrip (first parts: {}, second parts: {})",
                first_snapshot.parts.len(),
                second_snapshot.parts.len()
            ));
        }

        Ok(())
    }

    fn package_snapshot(path: &Path) -> TestResult<PackageSnapshot> {
        let package = offidized_opc::Package::open(path)
            .map_err(|error| format!("open package snapshot `{}`: {error}", path.display()))?;

        let package_relationships = package
            .relationships()
            .iter()
            .map(relationship_snapshot)
            .collect::<Vec<_>>();

        let mut parts = package
            .parts()
            .map(|part| PartSnapshot {
                uri: part.uri.as_str().to_string(),
                content_type: part.content_type.clone(),
                relationships: part
                    .relationships
                    .iter()
                    .map(relationship_snapshot)
                    .collect::<Vec<_>>(),
                data: part.data.as_bytes().to_vec(),
            })
            .collect::<Vec<_>>();
        parts.sort_by(|left, right| left.uri.cmp(&right.uri));

        Ok(PackageSnapshot {
            package_relationships,
            parts,
        })
    }

    fn discover_pptx_fixtures_recursive(root: &Path) -> TestResult<Vec<PathBuf>> {
        if !root.is_dir() {
            return Err(format!("fixture root is missing: `{}`", root.display()));
        }

        let mut fixtures = Vec::new();
        let mut stack = vec![root.to_path_buf()];
        while let Some(directory) = stack.pop() {
            let entries = fs::read_dir(&directory).map_err(|error| {
                format!("read fixture directory `{}`: {error}", directory.display())
            })?;
            for entry in entries {
                let entry = entry.map_err(|error| {
                    format!(
                        "read fixture directory entry in `{}`: {error}",
                        directory.display()
                    )
                })?;
                let entry_path = entry.path();
                let file_type = entry.file_type().map_err(|error| {
                    format!(
                        "read file type for fixture path `{}`: {error}",
                        entry_path.display()
                    )
                })?;
                if file_type.is_dir() {
                    stack.push(entry_path);
                    continue;
                }
                if file_type.is_file()
                    && entry_path
                        .extension()
                        .and_then(|extension| extension.to_str())
                        .is_some_and(|extension| extension.eq_ignore_ascii_case("pptx"))
                    && !entry_path
                        .file_name()
                        .and_then(|file_name| file_name.to_str())
                        .is_some_and(|file_name| {
                            KNOWN_UNSUPPORTED_BROADER_SWEEP_FIXTURES
                                .iter()
                                .any(|known| known.eq_ignore_ascii_case(file_name))
                        })
                {
                    fixtures.push(entry_path);
                }
            }
        }
        fixtures.sort();
        Ok(fixtures)
    }

    fn relationship_snapshot(relationship: &offidized_opc::Relationship) -> RelationshipSnapshot {
        RelationshipSnapshot {
            id: relationship.id.clone(),
            rel_type: relationship.rel_type.clone(),
            target: relationship.target.clone(),
            target_mode_is_external: relationship.target_mode
                == offidized_opc::relationship::TargetMode::External,
        }
    }

    fn workspace_root() -> PathBuf {
        Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("../..")
            .canonicalize()
            .expect("workspace root path should resolve")
    }

    #[test]
    fn new_from_scratch_add_slides_save_roundtrip() {
        let tempdir = tempfile::tempdir().expect("tempdir");
        let path = tempdir.path().join("from_scratch.pptx");

        let mut pres = Presentation::new();
        pres.add_slide_with_title("First");
        pres.add_slide_with_title("Second");
        pres.save(&path).expect("save from-scratch");

        let reopened = Presentation::open(&path).expect("reopen from-scratch");
        assert_eq!(reopened.slide_count(), 2);
        assert_eq!(reopened.slides()[0].title(), "First");
        assert_eq!(reopened.slides()[1].title(), "Second");
    }

    #[test]
    fn roundtrip_preserves_unknown_package_parts() {
        // Create a presentation, save, then inject a custom part into the
        // ZIP that is *not* in the prefixes we strip. After open+save the
        // custom part must still be present.
        let tempdir = tempfile::tempdir().expect("tempdir");
        let original_path = tempdir.path().join("with_extra.pptx");
        let roundtripped_path = tempdir.path().join("roundtripped.pptx");

        // Step 1: create a minimal file.
        let mut pres = Presentation::new();
        pres.add_slide_with_title("Slide A");
        pres.save(&original_path).expect("initial save");

        // Step 2: inject a custom part via the OPC layer.
        {
            let mut package = offidized_opc::Package::open(&original_path).expect("open package");
            let custom_uri =
                offidized_opc::uri::PartUri::new("/customXml/item1.xml").expect("custom uri");
            let custom_part =
                offidized_opc::Part::new_xml(custom_uri, b"<hello>world</hello>".to_vec());
            package.set_part(custom_part);
            package.save(&original_path).expect("save with custom part");
        }

        // Step 3: open through Presentation and save again.
        let mut pres2 = Presentation::open(&original_path).expect("open with custom part");
        assert_eq!(pres2.slide_count(), 1);
        pres2.save(&roundtripped_path).expect("roundtrip save");

        // Step 4: verify the custom part survived.
        let roundtripped_package =
            offidized_opc::Package::open(&roundtripped_path).expect("open roundtripped");
        let custom_part = roundtripped_package.get_part("/customXml/item1.xml");
        assert!(
            custom_part.is_some(),
            "custom part /customXml/item1.xml should survive roundtrip"
        );
        assert_eq!(
            custom_part.unwrap().data.as_bytes(),
            b"<hello>world</hello>"
        );
    }

    #[test]
    fn dirty_save_preserves_unmanaged_presentation_and_slide_relationships() {
        const CUSTOM_PRESENTATION_RELATIONSHIP_TYPE: &str =
            "https://offidized.dev/relationships/presentation-metadata";
        const CUSTOM_SLIDE_RELATIONSHIP_TYPE: &str =
            "https://offidized.dev/relationships/slide-metadata";

        let tempdir = tempfile::tempdir().expect("tempdir");
        let baseline_path = tempdir.path().join("baseline.pptx");
        let injected_path = tempdir.path().join("injected.pptx");
        let roundtripped_path = tempdir.path().join("roundtripped.pptx");

        let mut presentation = Presentation::new();
        presentation.add_slide_with_title("Before");
        presentation
            .save(&baseline_path)
            .expect("save baseline presentation");

        let mut package = offidized_opc::Package::open(&baseline_path).expect("open baseline");
        let presentation_custom_uri =
            offidized_opc::uri::PartUri::new("/ppt/custom/pres-meta1.xml")
                .expect("valid presentation custom URI");
        let mut presentation_custom_part = offidized_opc::Part::new_xml(
            presentation_custom_uri.clone(),
            b"<custom scope=\"presentation\"/>".to_vec(),
        );
        presentation_custom_part.content_type = Some("application/xml".to_string());
        package.set_part(presentation_custom_part);
        package
            .get_part_mut("/ppt/presentation.xml")
            .expect("presentation part should exist")
            .relationships
            .add_new(
                CUSTOM_PRESENTATION_RELATIONSHIP_TYPE.to_string(),
                "custom/pres-meta1.xml".to_string(),
                offidized_opc::relationship::TargetMode::Internal,
            );

        let slide_custom_uri = offidized_opc::uri::PartUri::new("/ppt/custom/slide-meta1.xml")
            .expect("valid slide custom URI");
        let mut slide_custom_part = offidized_opc::Part::new_xml(
            slide_custom_uri.clone(),
            b"<custom scope=\"slide\"/>".to_vec(),
        );
        slide_custom_part.content_type = Some("application/xml".to_string());
        package.set_part(slide_custom_part);
        package
            .get_part_mut("/ppt/slides/slide1.xml")
            .expect("slide1 part should exist")
            .relationships
            .add_new(
                CUSTOM_SLIDE_RELATIONSHIP_TYPE.to_string(),
                "../custom/slide-meta1.xml".to_string(),
                offidized_opc::relationship::TargetMode::Internal,
            );
        package.save(&injected_path).expect("save injected package");

        let mut dirty_presentation =
            Presentation::open(&injected_path).expect("open injected presentation");
        dirty_presentation
            .slide_mut(0)
            .expect("slide 0 should exist")
            .set_title("After");
        dirty_presentation
            .save(&roundtripped_path)
            .expect("save dirty roundtrip");

        let roundtripped_package =
            offidized_opc::Package::open(&roundtripped_path).expect("open roundtripped package");
        let persisted_presentation_custom_part = roundtripped_package
            .get_part(presentation_custom_uri.as_str())
            .expect("presentation custom part must survive dirty save");
        assert_eq!(
            persisted_presentation_custom_part.data.as_bytes(),
            b"<custom scope=\"presentation\"/>"
        );
        let persisted_slide_custom_part = roundtripped_package
            .get_part(slide_custom_uri.as_str())
            .expect("slide custom part must survive dirty save");
        assert_eq!(
            persisted_slide_custom_part.data.as_bytes(),
            b"<custom scope=\"slide\"/>"
        );

        let persisted_presentation_part = roundtripped_package
            .get_part("/ppt/presentation.xml")
            .expect("presentation part should exist");
        let presentation_custom_relationships = persisted_presentation_part
            .relationships
            .get_by_type(CUSTOM_PRESENTATION_RELATIONSHIP_TYPE);
        assert_eq!(
            presentation_custom_relationships.len(),
            1,
            "custom presentation relationship must survive dirty save"
        );
        let presentation_custom_target = offidized_opc::uri::PartUri::new("/ppt/presentation.xml")
            .expect("valid presentation URI")
            .resolve_relative(presentation_custom_relationships[0].target.as_str())
            .expect("resolve custom presentation target");
        assert_eq!(presentation_custom_target, presentation_custom_uri);

        let persisted_slide_part = roundtripped_package
            .get_part("/ppt/slides/slide1.xml")
            .expect("slide1 part should exist");
        let slide_custom_relationships = persisted_slide_part
            .relationships
            .get_by_type(CUSTOM_SLIDE_RELATIONSHIP_TYPE);
        assert_eq!(
            slide_custom_relationships.len(),
            1,
            "custom slide relationship must survive dirty save"
        );
        let slide_custom_target = offidized_opc::uri::PartUri::new("/ppt/slides/slide1.xml")
            .expect("valid slide URI")
            .resolve_relative(slide_custom_relationships[0].target.as_str())
            .expect("resolve custom slide target");
        assert_eq!(slide_custom_target, slide_custom_uri);

        let reopened = Presentation::open(&roundtripped_path).expect("open final presentation");
        assert_eq!(
            reopened.slide(0).expect("slide 0 should exist").title(),
            "After"
        );
    }

    #[test]
    fn parse_and_serialize_shape_preserves_unknown_shape_children() {
        let slide_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
       xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
       xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <p:cSld>
    <p:spTree>
      <p:sp>
        <p:nvSpPr><p:cNvPr id="2" name="Body"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
        <p:spPr/>
        <p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>Hello</a:t></a:r></a:p></p:txBody>
        <p:extLst><p:ext uri="{shape-unknown}"/></p:extLst>
      </p:sp>
    </p:spTree>
  </p:cSld>
</p:sld>"#;

        let parsed_shapes = super::parse_slide_shapes(slide_xml).expect("parse shapes");
        assert_eq!(parsed_shapes.len(), 1);
        assert_eq!(parsed_shapes[0].unknown_children().len(), 1);

        let mut slide = Slide::new("Title");
        super::append_shape(&mut slide, parsed_shapes[0].clone());
        let xml = super::serialize_slide_xml(&slide, &[], &[]).expect("serialize slide");
        let serialized = String::from_utf8(xml).expect("utf8");
        assert!(serialized.contains("<p:extLst>"));
        assert!(serialized.contains("shape-unknown"));
    }

    #[test]
    fn dirty_save_passthroughs_clean_slide_bytes() {
        let tempdir = tempfile::tempdir().expect("tempdir");
        let baseline_path = tempdir.path().join("baseline.pptx");
        let injected_path = tempdir.path().join("injected.pptx");
        let roundtripped_path = tempdir.path().join("roundtripped.pptx");

        let mut presentation = Presentation::new();
        presentation.add_slide_with_title("Keep");
        presentation.add_slide_with_title("Change");
        presentation
            .save(&baseline_path)
            .expect("save baseline presentation");

        let mut injected_package =
            offidized_opc::Package::open(&baseline_path).expect("open baseline");
        let keep_slide_xml = String::from_utf8_lossy(
            injected_package
                .get_part("/ppt/slides/slide1.xml")
                .expect("slide1 should exist")
                .data
                .as_bytes(),
        )
        .into_owned();
        let keep_slide_xml = keep_slide_xml.replacen(
            "<p:sld xmlns:p=\"http://schemas.openxmlformats.org/presentationml/2006/main\" xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" xmlns:c=\"http://schemas.openxmlformats.org/drawingml/2006/chart\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">",
            "<p:sld xmlns:p=\"http://schemas.openxmlformats.org/presentationml/2006/main\" xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" xmlns:c=\"http://schemas.openxmlformats.org/drawingml/2006/chart\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" xmlns:foo=\"urn:offidized:test\" foo:keep=\"1\">",
            1,
        );
        injected_package
            .get_part_mut("/ppt/slides/slide1.xml")
            .expect("slide1 should exist")
            .data = offidized_opc::PartData::Xml(keep_slide_xml.into_bytes());
        injected_package
            .save(&injected_path)
            .expect("save injected package");

        let injected_slide_bytes = offidized_opc::Package::open(&injected_path)
            .expect("open injected package")
            .get_part("/ppt/slides/slide1.xml")
            .expect("slide1 should exist")
            .data
            .as_bytes()
            .to_vec();

        let mut opened = Presentation::open(&injected_path).expect("open injected presentation");
        opened
            .slide_mut(1)
            .expect("slide2 should exist")
            .set_title("After");
        opened
            .save(&roundtripped_path)
            .expect("save dirty presentation");

        let final_slide_bytes = offidized_opc::Package::open(&roundtripped_path)
            .expect("open roundtripped package")
            .get_part("/ppt/slides/slide1.xml")
            .expect("slide1 should exist")
            .data
            .as_bytes()
            .to_vec();
        assert_eq!(
            final_slide_bytes, injected_slide_bytes,
            "clean slide bytes should be passed through unchanged when another slide is dirty",
        );
    }

    #[test]
    fn roundtrip_preserves_unknown_slide_xml_children() {
        // Verify that unknown children of <p:sld> survive a roundtrip.
        let slide_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
       xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
       xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <p:cSld>
    <p:spTree>
      <p:sp>
        <p:nvSpPr><p:cNvPr id="1" name="Title 1"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
        <p:spPr/>
        <p:txBody>
          <a:bodyPr/>
          <a:p><a:r><a:t>Hello</a:t></a:r></a:p>
        </p:txBody>
      </p:sp>
    </p:spTree>
  </p:cSld>
  <p:extLst>
    <p:ext uri="{BB962C8B-B14F-4D97-AF65-F5344CB8AC3E}">
      <p14:creationId xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main" val="12345"/>
    </p:ext>
  </p:extLst>
</p:sld>"#;

        let unknown = super::parse_slide_unknown_children(slide_xml).expect("parse unknown");
        assert_eq!(unknown.len(), 1, "should capture <p:extLst>");
        match &unknown[0] {
            offidized_opc::RawXmlNode::Element { name, .. } => {
                assert_eq!(name, "p:extLst");
            }
            other => panic!("expected Element, got {other:?}"),
        }
    }

    // ── Feature #1: Shape outline roundtrip ──

    #[test]
    fn roundtrip_shape_outline() {
        use crate::shape::{LineDashStyle, ShapeOutline};

        let tempdir = tempfile::tempdir().expect("tempdir");
        let path = tempdir.path().join("outline.pptx");

        let mut prs = Presentation::new();
        let slide = prs.add_slide_with_title("Outline Test");
        let shape = slide.add_shape("Box");
        shape.set_outline(ShapeOutline {
            width_emu: Some(25400),
            color_srgb: Some("FF0000".to_string()),
            dash_style: Some(LineDashStyle::Dash),
            compound_style: None,
            head_arrow: None,
            tail_arrow: None,
            alpha: None,
            color: None,
        });
        shape.add_paragraph_with_text("Outlined");
        prs.save(&path).expect("save outline pptx");

        let pkg = offidized_opc::Package::open(&path).expect("open outline pkg");
        let slide_xml = String::from_utf8(
            pkg.get_part("/ppt/slides/slide1.xml")
                .expect("slide1")
                .data
                .as_bytes()
                .to_vec(),
        )
        .unwrap();
        assert!(slide_xml.contains(r#"<a:ln w="25400">"#));
        assert!(slide_xml.contains(r#"<a:srgbClr val="FF0000"/>"#));
        assert!(slide_xml.contains(r#"<a:prstDash val="dash"/>"#));

        let reopened = Presentation::open(&path).expect("open outline pptx");
        let shape = &reopened.slide(0).unwrap().shapes()[0];
        let outline = shape.outline().expect("should have outline");
        assert_eq!(outline.width_emu, Some(25400));
        assert_eq!(outline.color_srgb.as_deref(), Some("FF0000"));
        assert_eq!(outline.dash_style, Some(LineDashStyle::Dash));
    }

    // ── Feature #2: Gradient fill roundtrip ──

    #[test]
    fn roundtrip_shape_gradient_fill() {
        use crate::shape::{GradientFill, GradientFillType, GradientStop};

        let tempdir = tempfile::tempdir().expect("tempdir");
        let path = tempdir.path().join("gradient.pptx");

        let mut prs = Presentation::new();
        let slide = prs.add_slide_with_title("Gradient");
        let shape = slide.add_shape("GradBox");
        shape.set_gradient_fill(GradientFill {
            stops: vec![
                GradientStop {
                    position: 0,
                    color_srgb: "FF0000".to_string(),
                    color: None,
                },
                GradientStop {
                    position: 100000,
                    color_srgb: "0000FF".to_string(),
                    color: None,
                },
            ],
            fill_type: Some(GradientFillType::Linear),
            linear_angle: Some(5400000),
        });
        shape.add_paragraph_with_text("Gradient shape");
        prs.save(&path).expect("save gradient pptx");

        let pkg = offidized_opc::Package::open(&path).expect("open gradient pkg");
        let slide_xml = String::from_utf8(
            pkg.get_part("/ppt/slides/slide1.xml")
                .expect("slide1")
                .data
                .as_bytes()
                .to_vec(),
        )
        .unwrap();
        assert!(slide_xml.contains("<a:gradFill>"));
        assert!(slide_xml.contains(r#"<a:gs pos="0">"#));
        assert!(slide_xml.contains(r#"<a:gs pos="100000">"#));
        assert!(slide_xml.contains(r#"<a:lin ang="5400000" scaled="0"/>"#));

        let reopened = Presentation::open(&path).expect("open gradient pptx");
        let shape = &reopened.slide(0).unwrap().shapes()[0];
        let gradient = shape.gradient_fill().expect("should have gradient");
        assert_eq!(gradient.stops.len(), 2);
        assert_eq!(gradient.stops[0].color_srgb, "FF0000");
        assert_eq!(gradient.stops[1].color_srgb, "0000FF");
    }

    // ── Feature #3: Shape rotation roundtrip ──

    #[test]
    fn roundtrip_shape_rotation() {
        let tempdir = tempfile::tempdir().expect("tempdir");
        let path = tempdir.path().join("rotation.pptx");

        let mut prs = Presentation::new();
        let slide = prs.add_slide_with_title("Rotation");
        let shape = slide.add_shape("Rotated");
        shape.set_geometry(ShapeGeometry::new(100, 200, 300, 400));
        shape.set_rotation(5400000); // 90 degrees
        shape.add_paragraph_with_text("Rotated text");
        prs.save(&path).expect("save rotation pptx");

        let pkg = offidized_opc::Package::open(&path).expect("open rotation pkg");
        let slide_xml = String::from_utf8(
            pkg.get_part("/ppt/slides/slide1.xml")
                .expect("slide1")
                .data
                .as_bytes()
                .to_vec(),
        )
        .unwrap();
        assert!(slide_xml.contains(r#"rot="5400000">"#));

        let reopened = Presentation::open(&path).expect("open rotation pptx");
        let shape = &reopened.slide(0).unwrap().shapes()[0];
        assert_eq!(shape.rotation(), Some(5400000));
    }

    // ── Feature #4: Shape hidden state roundtrip ──

    #[test]
    fn roundtrip_shape_hidden() {
        let tempdir = tempfile::tempdir().expect("tempdir");
        let path = tempdir.path().join("hidden-shape.pptx");

        let mut prs = Presentation::new();
        let slide = prs.add_slide_with_title("Hidden");
        let shape = slide.add_shape("Invisible");
        shape.set_hidden(true);
        shape.add_paragraph_with_text("Hidden text");
        prs.save(&path).expect("save hidden shape pptx");

        let pkg = offidized_opc::Package::open(&path).expect("open hidden shape pkg");
        let slide_xml = String::from_utf8(
            pkg.get_part("/ppt/slides/slide1.xml")
                .expect("slide1")
                .data
                .as_bytes()
                .to_vec(),
        )
        .unwrap();
        assert!(slide_xml.contains(r#"hidden="1""#));

        let reopened = Presentation::open(&path).expect("open hidden shape pptx");
        let shape = &reopened.slide(0).unwrap().shapes()[0];
        assert!(shape.is_hidden());
    }

    // ── Feature #5 & #6: Table cell formatting and dimensions ──

    #[test]
    fn roundtrip_table_cell_formatting_and_dimensions() {
        use crate::table::CellBorder;

        let tempdir = tempfile::tempdir().expect("tempdir");
        let path = tempdir.path().join("table-fmt.pptx");

        let mut prs = Presentation::new();
        let slide = prs.add_slide_with_title("Table Fmt");
        let table = slide.add_table(2, 2);
        table.set_cell_text(0, 0, "Header");
        table.set_column_width_emu(0, 914400);
        table.set_column_width_emu(1, 1828800);
        table.set_row_height_emu(0, 457200);
        table.set_row_height_emu(1, 228600);

        if let Some(cell) = table.cell_mut(0, 0) {
            cell.set_fill_color_srgb("AABBCC");
            cell.set_bold(true);
            cell.set_font_size(2400);
            cell.set_font_color_srgb("FF0000");
            cell.borders_mut().top = Some(CellBorder {
                width_emu: Some(12700),
                color_srgb: Some("000000".to_string()),
                color: None,
            });
        }

        prs.save(&path).expect("save table fmt pptx");

        let pkg = offidized_opc::Package::open(&path).expect("open table fmt pkg");
        let slide_xml = String::from_utf8(
            pkg.get_part("/ppt/slides/slide1.xml")
                .expect("slide1")
                .data
                .as_bytes()
                .to_vec(),
        )
        .unwrap();
        assert!(slide_xml.contains(r#"<a:gridCol w="914400"/>"#));
        assert!(slide_xml.contains(r#"<a:gridCol w="1828800"/>"#));
        assert!(slide_xml.contains(r#"<a:tr h="457200">"#));
        assert!(slide_xml.contains(r#"<a:tr h="228600">"#));
        assert!(slide_xml.contains(r#"b="1""#));
        assert!(slide_xml.contains(r#"<a:srgbClr val="AABBCC"/>"#));

        let reopened = Presentation::open(&path).expect("open table fmt pptx");
        let table = &reopened.slide(0).unwrap().tables()[0];
        assert_eq!(table.column_widths_emu()[0], 914400);
        assert_eq!(table.column_widths_emu()[1], 1828800);
        assert_eq!(table.row_heights_emu()[0], 457200);
        assert_eq!(table.row_heights_emu()[1], 228600);

        let cell = table.cell(0, 0).unwrap();
        assert_eq!(cell.fill_color_srgb(), Some("AABBCC"));
        assert_eq!(cell.bold(), Some(true));
        assert_eq!(cell.font_size(), Some(2400));
        assert_eq!(cell.font_color_srgb(), Some("FF0000"));
        assert!(cell.borders().top.is_some());
        assert_eq!(cell.borders().top.as_ref().unwrap().width_emu, Some(12700));
    }

    // ── Feature #7: Slide background roundtrip ──

    #[test]
    fn roundtrip_slide_background() {
        use crate::slide::SlideBackground;

        let tempdir = tempfile::tempdir().expect("tempdir");
        let path = tempdir.path().join("background.pptx");

        let mut prs = Presentation::new();
        let slide = prs.add_slide_with_title("BG Test");
        slide.set_background(SlideBackground::Solid("FF6600".to_string()));
        prs.save(&path).expect("save bg pptx");

        let pkg = offidized_opc::Package::open(&path).expect("open bg pkg");
        let slide_xml = String::from_utf8(
            pkg.get_part("/ppt/slides/slide1.xml")
                .expect("slide1")
                .data
                .as_bytes()
                .to_vec(),
        )
        .unwrap();
        assert!(slide_xml.contains("<p:bg>"));
        assert!(slide_xml.contains(r#"<a:srgbClr val="FF6600"/>"#));

        let reopened = Presentation::open(&path).expect("open bg pptx");
        let slide = reopened.slide(0).unwrap();
        match slide.background() {
            Some(SlideBackground::Solid(color)) => assert_eq!(color.as_str(), "FF6600"),
            other => panic!("expected Solid background, got {other:?}"),
        }
    }

    // ── Feature #8: Slide hidden state roundtrip ──

    #[test]
    fn roundtrip_slide_hidden() {
        let tempdir = tempfile::tempdir().expect("tempdir");
        let path = tempdir.path().join("hidden-slide.pptx");

        let mut prs = Presentation::new();
        let slide = prs.add_slide_with_title("Hidden Slide");
        slide.set_hidden(true);
        prs.save(&path).expect("save hidden slide pptx");

        let pkg = offidized_opc::Package::open(&path).expect("open hidden slide pkg");
        let slide_xml = String::from_utf8(
            pkg.get_part("/ppt/slides/slide1.xml")
                .expect("slide1")
                .data
                .as_bytes()
                .to_vec(),
        )
        .unwrap();
        assert!(slide_xml.contains(r#"show="0""#));

        let reopened = Presentation::open(&path).expect("open hidden slide pptx");
        let slide = reopened.slide(0).unwrap();
        assert!(slide.is_hidden());
    }

    // ── Feature #9: Slide size roundtrip ──

    #[test]
    fn roundtrip_slide_size() {
        let tempdir = tempfile::tempdir().expect("tempdir");
        let path = tempdir.path().join("slide-size.pptx");

        let mut prs = Presentation::new();
        prs.set_slide_width_emu(12_192_000); // widescreen
        prs.set_slide_height_emu(6_858_000);
        prs.add_slide_with_title("Widescreen");
        prs.save(&path).expect("save slide size pptx");

        let pkg = offidized_opc::Package::open(&path).expect("open slide size pkg");
        let pres_xml = String::from_utf8(
            pkg.get_part(super::PRESENTATION_PART_URI)
                .expect("pres part")
                .data
                .as_bytes()
                .to_vec(),
        )
        .unwrap();
        assert!(pres_xml.contains(r#"cx="12192000""#));
        assert!(pres_xml.contains(r#"cy="6858000""#));

        let reopened = Presentation::open(&path).expect("open slide size pptx");
        assert_eq!(reopened.slide_width_emu(), Some(12_192_000));
        assert_eq!(reopened.slide_height_emu(), Some(6_858_000));
    }

    // ── Feature #11: Bullet properties roundtrip ──

    #[test]
    fn roundtrip_bullet_properties() {
        use crate::shape::{BulletProperties, BulletStyle};

        let tempdir = tempfile::tempdir().expect("tempdir");
        let path = tempdir.path().join("bullets.pptx");

        let mut prs = Presentation::new();
        let slide = prs.add_slide_with_title("Bullets");
        let shape = slide.add_shape("List");
        let para = shape.add_paragraph();
        {
            let props = para.properties_mut();
            props.bullet = BulletProperties {
                style: Some(BulletStyle::Char("\u{2022}".to_string())),
                font_name: Some("Symbol".to_string()),
                size_percent: Some(100000),
                color_srgb: Some("FF0000".to_string()),
                color: None,
            };
        }
        para.add_run("Bullet item");

        prs.save(&path).expect("save bullets pptx");

        let pkg = offidized_opc::Package::open(&path).expect("open bullets pkg");
        let slide_xml = String::from_utf8(
            pkg.get_part("/ppt/slides/slide1.xml")
                .expect("slide1")
                .data
                .as_bytes()
                .to_vec(),
        )
        .unwrap();
        assert!(slide_xml.contains(r#"<a:buFont typeface="Symbol"/>"#));
        assert!(slide_xml.contains(r#"<a:buSzPct val="100000"/>"#));
        assert!(slide_xml.contains(r#"<a:buClr>"#));
        assert!(slide_xml.contains(r#"<a:buChar char="#));

        let reopened = Presentation::open(&path).expect("open bullets pptx");
        let shape = &reopened.slide(0).unwrap().shapes()[0];
        let para = &shape.paragraphs()[0];
        let bullet = &para.properties().bullet;
        assert_eq!(
            bullet.style,
            Some(BulletStyle::Char("\u{2022}".to_string()))
        );
        assert_eq!(bullet.font_name.as_deref(), Some("Symbol"));
        assert_eq!(bullet.size_percent, Some(100000));
        assert_eq!(bullet.color_srgb.as_deref(), Some("FF0000"));
    }

    // ── Feature #13: Chart type and legend roundtrip ──

    #[test]
    fn roundtrip_chart_type_and_legend() {
        use crate::chart::{ChartSeries, ChartType, LegendPosition};

        let tempdir = tempfile::tempdir().expect("tempdir");
        let path = tempdir.path().join("chart-type.pptx");

        let mut prs = Presentation::new();
        let slide = prs.add_slide_with_title("Chart Type");
        let chart = slide.add_chart("Revenue");
        chart.add_data_point("Q1", 10.0);
        chart.add_data_point("Q2", 14.5);
        chart.set_chart_type(ChartType::Line);
        chart.set_legend_position(LegendPosition::Bottom);

        let mut series2 = ChartSeries::new("Expenses");
        series2.set_values(vec![8.0, 12.0]);
        chart.add_series(series2);

        prs.save(&path).expect("save chart type pptx");

        let pkg = offidized_opc::Package::open(&path).expect("open chart type pkg");
        // Find the chart part.
        let chart_part = pkg.get_part("/ppt/charts/chart1.xml").expect("chart1 part");
        let chart_xml = String::from_utf8(chart_part.data.as_bytes().to_vec()).expect("chart xml");
        assert!(chart_xml.contains("<c:lineChart>"));
        assert!(chart_xml.contains("<c:legend>"));
        assert!(chart_xml.contains(r#"<c:legendPos val="b"/>"#));
        // Should have 2 series.
        let ser_count = chart_xml.matches("<c:ser>").count();
        assert_eq!(ser_count, 2, "should have 2 series");

        let reopened = Presentation::open(&path).expect("open chart type pptx");
        let chart = &reopened.slide(0).unwrap().charts()[0];
        assert_eq!(chart.chart_type(), ChartType::Line);
        assert!(chart.show_legend());
        assert_eq!(chart.legend_position(), Some(LegendPosition::Bottom));
        assert_eq!(chart.additional_series().len(), 1);
        assert_eq!(chart.additional_series()[0].values(), &[8.0, 12.0]);
    }

    // ── Feature #15: Header/footer roundtrip ──

    #[test]
    fn roundtrip_slide_header_footer() {
        use crate::slide::SlideHeaderFooter;

        let tempdir = tempfile::tempdir().expect("tempdir");
        let path = tempdir.path().join("header-footer.pptx");

        let mut prs = Presentation::new();
        let slide = prs.add_slide_with_title("HF Test");
        slide.set_header_footer(SlideHeaderFooter {
            show_slide_number: Some(true),
            show_date_time: Some(false),
            show_header: None,
            show_footer: Some(true),
            footer_text: None,
            date_time_text: None,
        });
        prs.save(&path).expect("save hf pptx");

        let pkg = offidized_opc::Package::open(&path).expect("open hf pkg");
        let slide_xml = String::from_utf8(
            pkg.get_part("/ppt/slides/slide1.xml")
                .expect("slide1")
                .data
                .as_bytes()
                .to_vec(),
        )
        .unwrap();
        assert!(slide_xml.contains(r#"<p:hf "#));
        assert!(slide_xml.contains(r#"sldNum="1""#));
        assert!(slide_xml.contains(r#"dt="0""#));
        assert!(slide_xml.contains(r#"ftr="1""#));

        let reopened = Presentation::open(&path).expect("open hf pptx");
        let slide = reopened.slide(0).unwrap();
        let hf = slide.header_footer().expect("should have hf");
        assert_eq!(hf.show_slide_number, Some(true));
        assert_eq!(hf.show_date_time, Some(false));
        assert_eq!(hf.show_footer, Some(true));
    }

    // ── Feature #14: Parse theme color scheme ──

    #[test]
    fn parse_theme_color_scheme_extracts_colors() {
        let theme_xml = r#"
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
              </a:themeElements>
            </a:theme>
        "#;

        let scheme = super::parse_theme_color_scheme(theme_xml.as_bytes())
            .expect("parse theme scheme")
            .expect("should have color scheme");
        assert_eq!(scheme.dark1.as_deref(), Some("000000"));
        assert_eq!(scheme.light1.as_deref(), Some("FFFFFF"));
        assert_eq!(scheme.dark2.as_deref(), Some("44546A"));
        assert_eq!(scheme.light2.as_deref(), Some("E7E6E6"));
        assert_eq!(scheme.accent1.as_deref(), Some("4472C4"));
        assert_eq!(scheme.accent2.as_deref(), Some("ED7D31"));
        assert_eq!(scheme.hyperlink.as_deref(), Some("0563C1"));
        assert_eq!(scheme.followed_hyperlink.as_deref(), Some("954F72"));
        assert_eq!(scheme.name.as_deref(), Some("Office"));
    }

    // ── Feature #8: Parse slide hidden attribute ──

    #[test]
    fn parse_slide_hidden_attribute() {
        let xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
       xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
       show="0">
  <p:cSld>
    <p:spTree>
      <p:sp>
        <p:nvSpPr><p:cNvPr id="1" name="Title"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
        <p:txBody><a:bodyPr/><a:p><a:r><a:t>Hidden</a:t></a:r></a:p></p:txBody>
      </p:sp>
    </p:spTree>
  </p:cSld>
</p:sld>"#;

        let hidden = super::parse_slide_hidden(xml).expect("parse hidden");
        assert!(hidden);
    }

    // ── Feature #9: Parse slide size ──

    #[test]
    fn parse_slide_size_from_presentation_xml() {
        let xml = r#"
            <p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
                            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
              <p:sldMasterIdLst>
                <p:sldMasterId id="2147483648" r:id="rId1"/>
              </p:sldMasterIdLst>
              <p:sldIdLst>
                <p:sldId id="256" r:id="rId2"/>
              </p:sldIdLst>
              <p:sldSz cx="12192000" cy="6858000"/>
            </p:presentation>
        "#;

        let (width, height) = super::parse_slide_size(xml.as_bytes()).expect("parse size");
        assert_eq!(width, Some(12_192_000));
        assert_eq!(height, Some(6_858_000));
    }

    // ── Feature #7: Parse slide background ──

    #[test]
    fn parse_slide_background_solid() {
        let xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
       xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <p:cSld>
    <p:bg>
      <p:bgPr>
        <a:solidFill>
          <a:srgbClr val="FF6600"/>
        </a:solidFill>
        <a:effectLst/>
      </p:bgPr>
    </p:bg>
    <p:spTree/>
  </p:cSld>
</p:sld>"#;

        let bg = super::parse_slide_background(xml).expect("parse bg");
        match bg {
            Some(crate::slide::SlideBackground::Solid(color)) => {
                assert_eq!(color, "FF6600");
            }
            other => panic!("expected Solid background, got {other:?}"),
        }
    }

    // ── Feature #13: Parse chart XML with chart type and legend ──

    #[test]
    fn parse_chart_xml_detects_chart_type_and_legend() {
        use crate::chart::{ChartType, LegendPosition};

        let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
              xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <c:chart>
    <c:title>
      <c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>Revenue</a:t></a:r></a:p></c:rich></c:tx>
    </c:title>
    <c:plotArea>
      <c:layout/>
      <c:lineChart>
        <c:ser>
          <c:idx val="0"/>
          <c:order val="0"/>
          <c:tx><c:v>Series1</c:v></c:tx>
          <c:cat>
            <c:strLit>
              <c:ptCount val="2"/>
              <c:pt idx="0"><c:v>Q1</c:v></c:pt>
              <c:pt idx="1"><c:v>Q2</c:v></c:pt>
            </c:strLit>
          </c:cat>
          <c:val>
            <c:numLit>
              <c:ptCount val="2"/>
              <c:pt idx="0"><c:v>10</c:v></c:pt>
              <c:pt idx="1"><c:v>20</c:v></c:pt>
            </c:numLit>
          </c:val>
        </c:ser>
        <c:ser>
          <c:idx val="1"/>
          <c:order val="1"/>
          <c:tx><c:v>Series2</c:v></c:tx>
          <c:val>
            <c:numLit>
              <c:ptCount val="2"/>
              <c:pt idx="0"><c:v>5</c:v></c:pt>
              <c:pt idx="1"><c:v>15</c:v></c:pt>
            </c:numLit>
          </c:val>
        </c:ser>
      </c:lineChart>
    </c:plotArea>
    <c:legend>
      <c:legendPos val="b"/>
    </c:legend>
  </c:chart>
</c:chartSpace>"#;

        let chart = super::parse_chart_xml(xml.as_bytes()).expect("parse chart");
        assert_eq!(chart.title(), "Revenue");
        assert_eq!(chart.chart_type(), ChartType::Line);
        assert!(chart.show_legend());
        assert_eq!(chart.legend_position(), Some(LegendPosition::Bottom));
        assert_eq!(chart.values(), &[10.0, 20.0]);
        assert_eq!(chart.categories(), &["Q1", "Q2"]);
        assert_eq!(chart.additional_series().len(), 1);
        assert_eq!(chart.additional_series()[0].values(), &[5.0, 15.0]);
    }

    // ── Feature #5 & #6: Parse table XML with cell formatting ──

    #[test]
    fn parse_table_xml_with_cell_formatting() {
        let xml = r#"
<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
       xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <p:cSld>
    <p:spTree>
      <p:graphicFrame>
        <a:graphic>
          <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table">
            <a:tbl>
              <a:tblPr/>
              <a:tblGrid>
                <a:gridCol w="914400"/>
                <a:gridCol w="1828800"/>
              </a:tblGrid>
              <a:tr h="457200">
                <a:tc>
                  <a:txBody>
                    <a:bodyPr/>
                    <a:p>
                      <a:r>
                        <a:rPr b="1" sz="2400">
                          <a:solidFill><a:srgbClr val="FF0000"/></a:solidFill>
                        </a:rPr>
                        <a:t>Header</a:t>
                      </a:r>
                    </a:p>
                  </a:txBody>
                  <a:tcPr>
                    <a:lnT w="12700">
                      <a:solidFill><a:srgbClr val="000000"/></a:solidFill>
                    </a:lnT>
                    <a:solidFill><a:srgbClr val="AABBCC"/></a:solidFill>
                  </a:tcPr>
                </a:tc>
                <a:tc>
                  <a:txBody><a:bodyPr/><a:p><a:r><a:t>Value</a:t></a:r></a:p></a:txBody>
                  <a:tcPr/>
                </a:tc>
              </a:tr>
            </a:tbl>
          </a:graphicData>
        </a:graphic>
      </p:graphicFrame>
    </p:spTree>
  </p:cSld>
</p:sld>"#;

        let tables = super::parse_slide_tables(xml.as_bytes()).expect("parse tables");
        assert_eq!(tables.len(), 1);
        let table = &tables[0];
        assert_eq!(table.rows(), 1);
        assert_eq!(table.cols(), 2);
        assert_eq!(table.column_widths_emu()[0], 914400);
        assert_eq!(table.column_widths_emu()[1], 1828800);
        assert_eq!(table.row_heights_emu()[0], 457200);

        let cell = table.cell(0, 0).unwrap();
        assert_eq!(cell.text(), "Header");
        assert_eq!(cell.fill_color_srgb(), Some("AABBCC"));
        assert_eq!(cell.bold(), Some(true));
        assert_eq!(cell.font_size(), Some(2400));
        assert_eq!(cell.font_color_srgb(), Some("FF0000"));
        assert!(cell.borders().top.is_some());
        assert_eq!(cell.borders().top.as_ref().unwrap().width_emu, Some(12700));
        assert_eq!(
            cell.borders().top.as_ref().unwrap().color_srgb.as_deref(),
            Some("000000")
        );

        let cell2 = table.cell(0, 1).unwrap();
        assert_eq!(cell2.text(), "Value");
        assert_eq!(cell2.fill_color_srgb(), None);
    }

    /// Regression: extra namespace declarations on `<p:sld>` must survive
    /// a dirty save so that unknown elements/attributes using those prefixes
    /// remain valid XML.
    #[test]
    fn dirty_save_preserves_extra_slide_namespace_declarations() {
        use offidized_opc::relationship::TargetMode;
        use offidized_opc::uri::PartUri;
        use offidized_opc::{Package, Part};

        // Minimal presentation.xml with one slide reference.
        let pres_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
                xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
                xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <p:sldIdLst>
    <p:sldId id="256" r:id="rId1"/>
  </p:sldIdLst>
</p:presentation>"#;

        // Slide XML with extra xmlns:mc and xmlns:p14 declarations.
        let slide_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
       xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
       xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
       xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
       xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main">
  <p:cSld>
    <p:spTree>
      <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
      <p:grpSpPr/>
      <p:sp>
        <p:nvSpPr><p:cNvPr id="2" name="Title 1"/><p:cNvSpPr/><p:nvPr><p:ph type="title"/></p:nvPr></p:nvSpPr>
        <p:spPr/>
        <p:txBody>
          <a:bodyPr/>
          <a:p><a:r><a:t>Hello Slide</a:t></a:r></a:p>
        </p:txBody>
      </p:sp>
    </p:spTree>
  </p:cSld>
</p:sld>"#;

        let mut package = Package::new();

        let pres_uri = PartUri::new("/ppt/presentation.xml").unwrap();
        let mut pres_part = Part::new_xml(pres_uri, pres_xml.to_vec());
        pres_part.content_type = Some(ContentTypeValue::PRESENTATION.to_string());
        pres_part.relationships.add_new(
            RelationshipType::SLIDE.to_string(),
            "slides/slide1.xml".to_string(),
            TargetMode::Internal,
        );
        package.set_part(pres_part);

        let slide_uri = PartUri::new("/ppt/slides/slide1.xml").unwrap();
        let mut slide_part = Part::new_xml(slide_uri, slide_xml.to_vec());
        slide_part.content_type = Some(ContentTypeValue::SLIDE.to_string());
        package.set_part(slide_part);

        package.relationships_mut().add_new(
            RelationshipType::PRESENTATION.to_string(),
            "ppt/presentation.xml".to_string(),
            TargetMode::Internal,
        );

        let tmpdir = tempfile::tempdir().unwrap();
        let pkg_path = tmpdir.path().join("test.pptx");
        package.save(&pkg_path).unwrap();

        // Open, modify a slide (marks dirty), and save.
        let mut pres = Presentation::open(&pkg_path).unwrap();
        assert_eq!(pres.slides()[0].title(), "Hello Slide");
        pres.slides_mut()[0].set_title("Modified Title");
        let out_path = tmpdir.path().join("out.pptx");
        pres.save(&out_path).unwrap();

        // Extract the slide XML and verify extra namespace declarations survived.
        let out_package = Package::open(&out_path).unwrap();
        let slide_part = out_package
            .get_part("/ppt/slides/slide1.xml")
            .expect("slide part missing");
        let slide_xml_out = String::from_utf8_lossy(slide_part.data.as_bytes());

        assert!(
            slide_xml_out.contains("xmlns:mc"),
            "mc namespace declaration missing from slide XML:\n{slide_xml_out}"
        );
        assert!(
            slide_xml_out.contains("xmlns:p14"),
            "p14 namespace declaration missing from slide XML:\n{slide_xml_out}"
        );
    }
}
