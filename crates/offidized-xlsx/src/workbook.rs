use std::collections::BTreeMap;
use std::io::Cursor;
use std::path::Path;

use offidized_opc::content_types::ContentTypeValue;
use offidized_opc::relationship::{RelationshipType, TargetMode};
use offidized_opc::uri::PartUri;
use offidized_opc::{Package, Part, PartData, RawXmlNode};
use quick_xml::events::{BytesDecl, BytesEnd, BytesStart, BytesText, Event};
use quick_xml::{Reader, Writer};

use crate::auto_filter::{AutoFilter, CustomFilter, CustomFilterOperator, FilterType};
use crate::cell::RichTextRun;
use crate::cell::{normalize_cell_reference, Cell, CellValue};
use crate::chart::Chart;
use crate::column::Column;
use crate::error::{Result, XlsxError};
use crate::named_style::{CellStyleXf, NamedStyle};
use crate::print_settings::{PageBreak, PageBreaks, PrintArea, PrintHeaderFooter};
use crate::range::CellRange;
use crate::row::Row;
use crate::shared_strings::{SharedStringEntry, SharedStrings};
use crate::sparkline::{
    Sparkline, SparklineAxisType, SparklineColors, SparklineEmptyCells, SparklineGroup,
    SparklineType,
};
use crate::style::{
    Alignment, Border, BorderSide, CellProtection, ColorReference, Fill, Font, FontScheme,
    FontVerticalAlign, GradientFill, GradientFillType, HorizontalAlignment, PatternFill,
    PatternFillType, Style, StyleTable, ThemeColor, VerticalAlignment,
};
use crate::worksheet::{
    CellAnchor, CfValueObject, CfValueObjectType, ColorScaleStop, Comment, ConditionalFormatting,
    ConditionalFormattingOperator, ConditionalFormattingRuleType, DataValidation,
    DataValidationErrorStyle, DataValidationType, FreezePane, Hyperlink, ImageAnchorType,
    PageMargins, PageOrientation, PageSetup, SheetProtection, SheetViewOptions, SheetVisibility,
    TableColumn, TotalFunction, Worksheet, WorksheetImage, WorksheetImageExt, WorksheetTable,
};

const WORKBOOK_PART_URI: &str = "/xl/workbook.xml";
const STYLES_PART_URI: &str = "/xl/styles.xml";
const SPREADSHEETML_NS: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const RELATIONSHIP_NS: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const TABLE_RELATIONSHIP_TYPE: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/table";
const TABLE_PART_CONTENT_TYPE: &str =
    "application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml";
const DRAWING_RELATIONSHIP_TYPE: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing";
const DRAWING_PART_CONTENT_TYPE: &str = "application/vnd.openxmlformats-officedocument.drawing+xml";
const DRAWINGML_NS: &str = "http://schemas.openxmlformats.org/drawingml/2006/main";
const DRAWINGML_SPREADSHEET_NS: &str =
    "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing";
const DEFAULT_IMAGE_EXTENT_CX: u64 = 952_500;
const DEFAULT_IMAGE_EXTENT_CY: u64 = 952_500;
const HYPERLINK_RELATIONSHIP_TYPE: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink";
const COMMENTS_RELATIONSHIP_TYPE: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments";
const COMMENTS_PART_CONTENT_TYPE: &str =
    "application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml";
const CHART_RELATIONSHIP_TYPE: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart";
const CHART_PART_CONTENT_TYPE: &str =
    "application/vnd.openxmlformats-officedocument.drawingml.chart+xml";
const CHART_NS: &str = "http://schemas.openxmlformats.org/drawingml/2006/chart";
const PIVOT_TABLE_RELATIONSHIP_TYPE: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotTable";
const PIVOT_TABLE_PART_CONTENT_TYPE: &str =
    "application/vnd.openxmlformats-officedocument.spreadsheetml.pivotTable+xml";
const PIVOT_CACHE_DEFINITION_RELATIONSHIP_TYPE: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheDefinition";
const PIVOT_CACHE_DEFINITION_PART_CONTENT_TYPE: &str =
    "application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheDefinition+xml";
const PIVOT_CACHE_RECORDS_RELATIONSHIP_TYPE: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheRecords";
const PIVOT_CACHE_RECORDS_PART_CONTENT_TYPE: &str =
    "application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheRecords+xml";
const SPARKLINE_EXT_URI: &str = "{05C60535-1F16-4FD2-B633-F4F36F0B64E0}";
const X14_NS: &str = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main";
const XM_NS: &str = "http://schemas.microsoft.com/office/excel/2006/main";
#[derive(Debug, Clone)]
struct WorksheetTablePartRef {
    relationship_id: String,
}

#[derive(Debug, Clone)]
struct WorksheetPivotTablePartRef {
    #[allow(dead_code)]
    relationship_id: String,
}

#[derive(Debug, Clone)]
struct WorksheetDrawingPartRef {
    relationship_id: String,
}

#[derive(Debug, Clone)]
struct WorksheetImagePartRef {
    relationship_id: String,
}

/// Maps a hyperlink index to its external relationship id (if external).
#[derive(Debug, Clone)]
struct WorksheetHyperlinkRef {
    relationship_id: String,
}

#[derive(Debug, Clone, Default)]
struct WorksheetRelationshipIds {
    table_relationship_ids: Vec<String>,
    drawing_relationship_ids: Vec<String>,
    hyperlink_relationship_ids: Vec<String>,
}

#[derive(Debug, Clone)]
struct ParsedDrawingImageRef {
    relationship_id: String,
    anchor_cell: String,
    ext: Option<WorksheetImageExt>,
    anchor_type: ImageAnchorType,
    from_anchor: Option<CellAnchor>,
    to_anchor: Option<CellAnchor>,
    extent_cx: Option<i64>,
    extent_cy: Option<i64>,
    position_x: Option<i64>,
    position_y: Option<i64>,
    crop_left: Option<f64>,
    crop_right: Option<f64>,
    crop_top: Option<f64>,
    crop_bottom: Option<f64>,
}

#[derive(Debug, Clone)]
struct WorkbookSheetRef {
    sheet_id: u32,
    name: String,
    relationship_id: String,
    part_uri: PartUri,
    visibility: SheetVisibility,
}

#[derive(Debug, Clone)]
struct ParsedSheetRef {
    #[allow(dead_code)]
    sheet_id: u32,
    name: String,
    relationship_id: String,
    visibility: SheetVisibility,
}

#[derive(Debug, Clone, Default)]
struct ParsedWorkbook {
    sheets: Vec<ParsedSheetRef>,
    defined_names: Vec<DefinedName>,
    workbook_protection: Option<WorkbookProtection>,
    calc_settings: Option<CalculationSettings>,
    /// Raw attributes from `<fileVersion>` for roundtrip preservation.
    file_version_attrs: Vec<(String, String)>,
    /// Raw attributes from `<workbookPr>` for roundtrip preservation.
    workbook_pr_attrs: Vec<(String, String)>,
    /// Raw attributes for each `<workbookView>` inside `<bookViews>`.
    book_views: Vec<Vec<(String, String)>>,
    /// Pivot cache entries: (cacheId, r:id) for roundtrip.
    pivot_caches: Vec<(u32, String)>,
    /// Raw `<customWorkbookViews>` element preserved for roundtrip fidelity.
    custom_workbook_views: Vec<RawXmlNode>,
}

/// Workbook-level defined name (named range/formula).
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct DefinedName {
    name: String,
    reference: String,
    local_sheet_id: Option<u32>,
    /// Unknown attributes preserved for roundtrip fidelity (hidden, comment, etc.).
    unknown_attrs: Vec<(String, String)>,
}

impl DefinedName {
    /// Creates a new defined name with the given name and reference.
    pub fn new(name: impl Into<String>, reference: impl Into<String>) -> Self {
        Self {
            name: name.into().trim().to_string(),
            reference: reference.into().trim().to_string(),
            local_sheet_id: None,
            unknown_attrs: Vec::new(),
        }
    }

    pub fn name(&self) -> &str {
        self.name.as_str()
    }

    pub fn reference(&self) -> &str {
        self.reference.as_str()
    }

    pub fn local_sheet_id(&self) -> Option<u32> {
        self.local_sheet_id
    }

    pub fn set_reference(&mut self, reference: impl Into<String>) -> &mut Self {
        self.reference = reference.into().trim().to_string();
        self
    }

    pub fn set_local_sheet_id(&mut self, local_sheet_id: u32) -> &mut Self {
        self.local_sheet_id = Some(local_sheet_id);
        self
    }

    pub fn clear_local_sheet_id(&mut self) -> &mut Self {
        self.local_sheet_id = None;
        self
    }
}

/// Workbook-level protection settings.
#[derive(Debug, Clone, PartialEq, Eq, Default)]
pub struct WorkbookProtection {
    lock_structure: bool,
    lock_windows: bool,
    password_hash: Option<String>,
}

impl WorkbookProtection {
    /// Creates new workbook protection with structure locked.
    pub fn new() -> Self {
        Self {
            lock_structure: true,
            ..Self::default()
        }
    }

    /// Returns whether the workbook structure is locked.
    pub fn lock_structure(&self) -> bool {
        self.lock_structure
    }

    /// Sets whether the workbook structure is locked.
    pub fn set_lock_structure(&mut self, value: bool) -> &mut Self {
        self.lock_structure = value;
        self
    }

    /// Returns whether windows are locked.
    pub fn lock_windows(&self) -> bool {
        self.lock_windows
    }

    /// Sets whether windows are locked.
    pub fn set_lock_windows(&mut self, value: bool) -> &mut Self {
        self.lock_windows = value;
        self
    }

    /// Returns the password hash.
    pub fn password_hash(&self) -> Option<&str> {
        self.password_hash.as_deref()
    }

    /// Sets the password hash.
    pub fn set_password_hash(&mut self, hash: impl Into<String>) -> &mut Self {
        let hash = hash.into();
        let hash = hash.trim().to_string();
        self.password_hash = if hash.is_empty() { None } else { Some(hash) };
        self
    }

    /// Clears the password hash.
    pub fn clear_password_hash(&mut self) -> &mut Self {
        self.password_hash = None;
        self
    }

    pub(crate) fn has_metadata(&self) -> bool {
        self.lock_structure || self.lock_windows || self.password_hash.is_some()
    }
}

/// Workbook calculation settings.
#[derive(Debug, Clone, PartialEq, Eq, Default)]
pub struct CalculationSettings {
    calc_mode: Option<String>,
    calc_id: Option<u32>,
    full_calc_on_load: Option<bool>,
}

impl CalculationSettings {
    pub fn new() -> Self {
        Self::default()
    }

    /// Returns the calculation mode ("manual", "auto", "autoNoTable").
    pub fn calc_mode(&self) -> Option<&str> {
        self.calc_mode.as_deref()
    }

    /// Sets the calculation mode.
    pub fn set_calc_mode(&mut self, mode: impl Into<String>) -> &mut Self {
        self.calc_mode = Some(mode.into());
        self
    }

    /// Clears the calculation mode.
    pub fn clear_calc_mode(&mut self) -> &mut Self {
        self.calc_mode = None;
        self
    }

    /// Returns the calculation ID.
    pub fn calc_id(&self) -> Option<u32> {
        self.calc_id
    }

    /// Sets the calculation ID.
    pub fn set_calc_id(&mut self, id: u32) -> &mut Self {
        self.calc_id = Some(id);
        self
    }

    /// Clears the calculation ID.
    pub fn clear_calc_id(&mut self) -> &mut Self {
        self.calc_id = None;
        self
    }

    /// Returns whether a full calculation is performed on load.
    pub fn full_calc_on_load(&self) -> Option<bool> {
        self.full_calc_on_load
    }

    /// Sets whether a full calculation is performed on load.
    pub fn set_full_calc_on_load(&mut self, value: bool) -> &mut Self {
        self.full_calc_on_load = Some(value);
        self
    }

    /// Clears the full calc on load setting.
    pub fn clear_full_calc_on_load(&mut self) -> &mut Self {
        self.full_calc_on_load = None;
        self
    }

    pub(crate) fn has_metadata(&self) -> bool {
        self.calc_mode.is_some() || self.calc_id.is_some() || self.full_calc_on_load.is_some()
    }
}

/// An in-memory workbook.
#[derive(Debug, Clone)]
pub struct Workbook {
    sheets: Vec<Worksheet>,
    defined_names: Vec<DefinedName>,
    styles: StyleTable,
    named_styles: Vec<NamedStyle>,
    cell_style_xfs: Vec<CellStyleXf>,
    workbook_protection: Option<WorkbookProtection>,
    calc_settings: Option<CalculationSettings>,
    /// Raw attributes from `<fileVersion>` for roundtrip preservation.
    file_version_attrs: Vec<(String, String)>,
    /// Raw attributes from `<workbookPr>` for roundtrip preservation.
    workbook_pr_attrs: Vec<(String, String)>,
    /// Raw attributes for each `<workbookView>` inside `<bookViews>`.
    book_views: Vec<Vec<(String, String)>>,
    /// Pivot cache entries: (cacheId, r:id) for roundtrip.
    pivot_caches: Vec<(u32, String)>,
    /// Raw `<customWorkbookViews>` children preserved for roundtrip fidelity.
    custom_workbook_views: Vec<RawXmlNode>,
    /// 12 theme color hex strings (with `#` prefix) in canonical OOXML order.
    theme_colors: Vec<String>,
    /// Major (heading) font name from the theme.
    major_font: Option<String>,
    /// Minor (body) font name from the theme.
    minor_font: Option<String>,
    /// Custom indexed color palette override from styles.xml, if present.
    indexed_colors: Option<Vec<String>>,
    /// Original package preserved for roundtrip fidelity.
    source_package: Option<Package>,
    dirty: bool,
}

impl Default for Workbook {
    fn default() -> Self {
        Self {
            sheets: Vec::new(),
            defined_names: Vec::new(),
            styles: StyleTable::new(),
            named_styles: Vec::new(),
            cell_style_xfs: Vec::new(),
            workbook_protection: None,
            calc_settings: None,
            file_version_attrs: Vec::new(),
            workbook_pr_attrs: Vec::new(),
            book_views: Vec::new(),
            pivot_caches: Vec::new(),
            custom_workbook_views: Vec::new(),
            theme_colors: crate::theme::DEFAULT_THEME_COLORS
                .iter()
                .map(|s| (*s).to_string())
                .collect(),
            major_font: None,
            minor_font: None,
            indexed_colors: None,
            source_package: None,
            dirty: true,
        }
    }
}

impl Workbook {
    pub fn new() -> Self {
        Self::default()
    }

    pub fn open(path: impl AsRef<Path>) -> Result<Self> {
        let package = Package::open(path)?;
        Self::from_package(package)
    }

    pub fn from_bytes(bytes: &[u8]) -> Result<Self> {
        let package = Package::from_bytes(bytes)?;
        Self::from_package(package)
    }

    fn from_package(package: Package) -> Result<Self> {
        let workbook_uri = resolve_workbook_part_uri(&package)?;
        let workbook_part = package.get_part(workbook_uri.as_str()).ok_or_else(|| {
            XlsxError::UnsupportedPackage(format!(
                "missing workbook part `{}`",
                workbook_uri.as_str()
            ))
        })?;

        let parsed_workbook = parse_workbook_xml(workbook_part.data.as_bytes())?;
        let shared_strings = parse_shared_strings_part(&package, &workbook_uri, workbook_part)?;
        let (styles, parsed_cell_style_xfs, parsed_named_styles, parsed_indexed_colors) =
            parse_styles_part(&package, &workbook_uri, workbook_part)?
                .unwrap_or_else(|| (StyleTable::new(), Vec::new(), Vec::new(), None));

        // Parse theme (colors + fonts) from theme1.xml
        let parsed_theme = parse_theme_part(&package, &workbook_uri, workbook_part);

        let mut sheets = Vec::with_capacity(parsed_workbook.sheets.len());

        for sheet_ref in parsed_workbook.sheets {
            let Some(relationship) = workbook_part
                .relationships
                .get_by_id(&sheet_ref.relationship_id)
            else {
                tracing::warn!(
                    relationship_id = sheet_ref.relationship_id.as_str(),
                    sheet_name = sheet_ref.name.as_str(),
                    "missing workbook relationship for sheet; skipping sheet"
                );
                continue;
            };

            if relationship.target_mode != TargetMode::Internal {
                tracing::warn!(
                    relationship_id = relationship.id.as_str(),
                    "sheet relationship is external; skipping sheet"
                );
                continue;
            }

            let sheet_uri = workbook_uri.resolve_relative(relationship.target.as_str())?;
            let sheet_part = package.get_part(sheet_uri.as_str()).ok_or_else(|| {
                XlsxError::UnsupportedPackage(format!(
                    "missing worksheet part `{}` for relationship `{}`",
                    sheet_uri.as_str(),
                    relationship.id
                ))
            })?;

            let mut worksheet = parse_worksheet_xml(
                sheet_ref.name.as_str(),
                sheet_part.data.as_bytes(),
                shared_strings.as_deref(),
            )?;
            let worksheet_relationship_ids =
                parse_worksheet_relationship_ids(sheet_part.data.as_bytes())?;
            load_worksheet_tables(
                &mut worksheet,
                &package,
                &sheet_uri,
                sheet_part,
                &worksheet_relationship_ids.table_relationship_ids,
            )?;
            load_worksheet_images(
                &mut worksheet,
                &package,
                &sheet_uri,
                sheet_part,
                &worksheet_relationship_ids.drawing_relationship_ids,
            )?;
            crate::chart_io::load_worksheet_charts(
                &mut worksheet,
                &package,
                &sheet_uri,
                sheet_part,
                &worksheet_relationship_ids.drawing_relationship_ids,
            )?;
            crate::pivot_table_io::load_worksheet_pivot_tables(
                &mut worksheet,
                &package,
                &sheet_uri,
                sheet_part,
            )?;
            resolve_worksheet_hyperlink_urls(&mut worksheet, sheet_part);
            load_worksheet_comments(&mut worksheet, &package, &sheet_uri, sheet_part);
            extract_sparklines_from_ext_lst(&mut worksheet);
            worksheet.set_parsed_visibility(sheet_ref.visibility);
            worksheet.set_original_part_bytes(
                sheet_uri.as_str().to_string(),
                sheet_part.data.as_bytes().to_vec(),
            );
            sheets.push(worksheet);
        }

        let mut workbook = Self {
            sheets,
            defined_names: parsed_workbook.defined_names,
            styles,
            named_styles: parsed_named_styles,
            cell_style_xfs: parsed_cell_style_xfs,
            workbook_protection: parsed_workbook.workbook_protection,
            calc_settings: parsed_workbook.calc_settings,
            file_version_attrs: parsed_workbook.file_version_attrs,
            workbook_pr_attrs: parsed_workbook.workbook_pr_attrs,
            book_views: parsed_workbook.book_views,
            pivot_caches: parsed_workbook.pivot_caches,
            custom_workbook_views: parsed_workbook.custom_workbook_views,
            theme_colors: parsed_theme.colors,
            major_font: parsed_theme.major_font,
            minor_font: parsed_theme.minor_font,
            indexed_colors: parsed_indexed_colors,
            source_package: Some(package),
            dirty: false,
        };
        // Wire up print areas from defined names
        for defined_name in &workbook.defined_names {
            if defined_name.name() == "_xlnm.Print_Area" {
                if let Some(sheet_id) = defined_name.local_sheet_id() {
                    if let Some(sheet) = workbook.sheets.get_mut(sheet_id as usize) {
                        // Strip sheet name prefix (e.g. "Sheet1!$A$1:$D$10" -> "$A$1:$D$10")
                        let reference = defined_name.reference();
                        let range = reference
                            .find('!')
                            .map(|pos| &reference[pos + 1..])
                            .unwrap_or(reference);
                        sheet.set_parsed_print_area(PrintArea::new(range));
                    }
                }
            }
        }

        workbook.ensure_style_capacity_for_cell_ids()?;
        Ok(workbook)
    }

    pub fn save(&self, path: impl AsRef<Path>) -> Result<()> {
        if !self.dirty {
            if let Some(package) = &self.source_package {
                package.save(path)?;
                return Ok(());
            }
        }

        let package = self.build_package()?;
        package.save(path)?;
        Ok(())
    }

    pub fn to_bytes(&self) -> Result<Vec<u8>> {
        if !self.dirty {
            if let Some(package) = &self.source_package {
                return Ok(package.to_bytes()?);
            }
        }

        let package = self.build_package()?;
        Ok(package.to_bytes()?)
    }

    fn build_package(&self) -> Result<Package> {
        let mut package = match &self.source_package {
            Some(original) => {
                let mut pkg = original.clone();
                let _ = pkg.remove_part("/xl/sharedStrings.xml");
                let _ = pkg.remove_part("/xl/styles.xml");
                let _ = pkg.remove_part("/xl/workbook.xml");
                pkg
            }
            None => Package::new(),
        };
        let workbook_uri = PartUri::new(WORKBOOK_PART_URI)?;
        let mut workbook_part = Part::new_xml(workbook_uri.clone(), Vec::new());
        workbook_part.content_type = Some(ContentTypeValue::WORKBOOK.to_string());
        let mut workbook_passthrough_relationships = BTreeMap::<String, usize>::new();
        let mut worksheet_passthrough_relationships = BTreeMap::<String, usize>::new();
        let mut drawing_passthrough_relationships = BTreeMap::<String, usize>::new();
        if let Some(original) = &self.source_package {
            if let Some(original_workbook_part) = original.get_part(workbook_uri.as_str()) {
                for relationship in original_workbook_part.relationships.iter() {
                    if !is_rebuilt_workbook_relationship_type(relationship.rel_type.as_str()) {
                        record_passthrough_relationship(
                            &mut workbook_passthrough_relationships,
                            relationship.rel_type.as_str(),
                        );
                        workbook_part.relationships.add(relationship.clone());
                    }
                }
            }
        }

        let mut sheet_refs = Vec::with_capacity(self.sheets.len());
        let mut shared_strings = SharedStrings::new();
        if let Some(original) = &self.source_package {
            if let Ok(original_workbook_uri) = resolve_workbook_part_uri(original) {
                if let Some(original_workbook_part) =
                    original.get_part(original_workbook_uri.as_str())
                {
                    match parse_shared_strings_part(
                        original,
                        &original_workbook_uri,
                        original_workbook_part,
                    ) {
                        Ok(Some(existing_shared_strings)) => {
                            // Use push_raw (not intern) to preserve original indices.
                            // Passthrough sheets keep their raw XML which references
                            // the original shared string indices. Deduplicating here
                            // would shift those indices and corrupt unchanged sheets.
                            for entry in existing_shared_strings {
                                match entry {
                                    SharedStringEntry::Plain(value) => {
                                        shared_strings.push_raw(value);
                                    }
                                    SharedStringEntry::RichText {
                                        runs,
                                        phonetic_runs,
                                        phonetic_pr,
                                    } => {
                                        shared_strings.push_raw_rich_text(
                                            runs,
                                            phonetic_runs,
                                            phonetic_pr,
                                        );
                                    }
                                }
                            }
                        }
                        Ok(None) => {}
                        Err(error) => {
                            tracing::warn!(
                                error = %error,
                                "failed to seed shared strings from source workbook"
                            );
                        }
                    }
                }
            }
        }
        let mut next_table_id = 1_u32;
        let mut next_pivot_table_id = 1_u32;
        let mut next_drawing_id = 1_u32;
        let mut next_image_id = 1_u32;
        let mut next_chart_id = 1_u32;

        for (index, sheet) in self.sheets.iter().enumerate() {
            let sheet_id = u32::try_from(index + 1).map_err(|_| {
                XlsxError::InvalidWorkbookState("too many worksheets to serialize".to_string())
            })?;

            let sheet_target = format!("worksheets/sheet{sheet_id}.xml");
            let relationship = workbook_part.relationships.add_new(
                RelationshipType::WORKSHEET.to_string(),
                sheet_target,
                TargetMode::Internal,
            );

            let sheet_uri = workbook_uri.resolve_relative(relationship.target.as_str())?;
            sheet_refs.push(WorkbookSheetRef {
                sheet_id,
                name: sheet.name().to_string(),
                relationship_id: relationship.id.clone(),
                part_uri: sheet_uri,
                visibility: sheet.visibility(),
            });
        }

        for (sheet, sheet_ref) in self.sheets.iter().zip(sheet_refs.iter()) {
            let mut table_refs = Vec::with_capacity(sheet.tables().len());
            let mut sheet_part = Part::new_xml(sheet_ref.part_uri.clone(), Vec::new());
            sheet_part.content_type = Some(ContentTypeValue::WORKSHEET.to_string());
            let original_sheet_part = self
                .source_package
                .as_ref()
                .and_then(|original| original.get_part(sheet_ref.part_uri.as_str()));

            let can_passthrough_sheet = !sheet.dirty()
                && sheet
                    .original_part_bytes()
                    .is_some_and(|(part_uri, _)| part_uri == sheet_ref.part_uri.as_str());
            if can_passthrough_sheet {
                if let Some(original_sheet_part) = original_sheet_part {
                    for relationship in original_sheet_part.relationships.iter() {
                        sheet_part.relationships.add(relationship.clone());
                    }
                }
                if let Some((_, original_bytes)) = sheet.original_part_bytes() {
                    sheet_part.data = PartData::Xml(original_bytes.to_vec());
                    package.set_part(sheet_part);
                    continue;
                }
            }

            if let Some(original) = &self.source_package {
                if let Some(original_sheet_part) = original.get_part(sheet_ref.part_uri.as_str()) {
                    for relationship in original_sheet_part.relationships.iter() {
                        if !is_rebuilt_worksheet_relationship_type(relationship.rel_type.as_str()) {
                            record_passthrough_relationship(
                                &mut worksheet_passthrough_relationships,
                                relationship.rel_type.as_str(),
                            );
                            sheet_part.relationships.add(relationship.clone());
                        }
                    }
                }
            }
            let mut drawing_ref = None;

            // If the sheet has no images but the original had a drawing relationship,
            // preserve it for charts and other non-image drawings.
            if sheet.images().is_empty() {
                if let Some(original) = &self.source_package {
                    if let Some(original_sheet_part) =
                        original.get_part(sheet_ref.part_uri.as_str())
                    {
                        for relationship in original_sheet_part.relationships.iter() {
                            if relationship.rel_type.as_str() == DRAWING_RELATIONSHIP_TYPE {
                                sheet_part.relationships.add(relationship.clone());
                                drawing_ref = Some(WorksheetDrawingPartRef {
                                    relationship_id: relationship.id.clone(),
                                });
                                break;
                            }
                        }
                    }
                }
            }

            for table in sheet.tables() {
                let table_id = next_table_id;
                next_table_id = next_table_id.checked_add(1).ok_or_else(|| {
                    XlsxError::InvalidWorkbookState(
                        "too many tables to serialize in workbook".to_string(),
                    )
                })?;

                let relationship = sheet_part.relationships.add_new(
                    TABLE_RELATIONSHIP_TYPE.to_string(),
                    format!("../tables/table{table_id}.xml"),
                    TargetMode::Internal,
                );
                table_refs.push(WorksheetTablePartRef {
                    relationship_id: relationship.id.clone(),
                });

                let table_uri = PartUri::new(format!("/xl/tables/table{table_id}.xml").as_str())?;
                let mut table_part =
                    Part::new_xml(table_uri, serialize_table_xml(table, table_id)?);
                table_part.content_type = Some(TABLE_PART_CONTENT_TYPE.to_string());
                package.set_part(table_part);
            }

            // Serialize pivot tables if the sheet has any.
            for pivot_table in sheet.pivot_tables() {
                let pivot_id = next_pivot_table_id;
                next_pivot_table_id = next_pivot_table_id.checked_add(1).ok_or_else(|| {
                    XlsxError::InvalidWorkbookState(
                        "too many pivot tables to serialize in workbook".to_string(),
                    )
                })?;

                // Create pivot cache records part with actual data
                let cache_records_uri = PartUri::new(
                    format!("/xl/pivotCache/pivotCacheRecords{pivot_id}.xml").as_str(),
                )?;
                let mut cache_records_part = Part::new_xml(
                    cache_records_uri.clone(),
                    crate::pivot_table_io::serialize_pivot_cache_records_xml(pivot_table, self)?,
                );
                cache_records_part.content_type =
                    Some(PIVOT_CACHE_RECORDS_PART_CONTENT_TYPE.to_string());
                package.set_part(cache_records_part);

                // Create pivot cache definition part
                let cache_def_uri = PartUri::new(
                    format!("/xl/pivotCache/pivotCacheDefinition{pivot_id}.xml").as_str(),
                )?;
                let mut cache_def_part = Part::new_xml(
                    cache_def_uri.clone(),
                    crate::pivot_table_io::serialize_pivot_cache_definition_xml(pivot_table, self)?,
                );
                cache_def_part.content_type =
                    Some(PIVOT_CACHE_DEFINITION_PART_CONTENT_TYPE.to_string());

                // Add relationship from cache definition to cache records
                cache_def_part.relationships.add_new(
                    PIVOT_CACHE_RECORDS_RELATIONSHIP_TYPE.to_string(),
                    format!("/xl/pivotCache/pivotCacheRecords{pivot_id}.xml"),
                    TargetMode::Internal,
                );
                package.set_part(cache_def_part);

                // Add relationship from workbook to pivot cache definition (CRITICAL!)
                workbook_part.relationships.add_new(
                    PIVOT_CACHE_DEFINITION_RELATIONSHIP_TYPE.to_string(),
                    format!("/xl/pivotCache/pivotCacheDefinition{pivot_id}.xml"),
                    TargetMode::Internal,
                );

                // Create pivot table part
                // Match ClosedXML naming: first table is "pivotTable.xml", subsequent are "pivotTable2.xml", etc.
                let pivot_table_name = if pivot_id == 1 {
                    "/xl/pivotTables/pivotTable.xml".to_string()
                } else {
                    format!("/xl/pivotTables/pivotTable{pivot_id}.xml")
                };
                let pivot_uri = PartUri::new(pivot_table_name.as_str())?;
                let mut pivot_part = Part::new_xml(
                    pivot_uri.clone(),
                    crate::pivot_table_io::serialize_pivot_table_xml(pivot_table, self)?,
                );
                pivot_part.content_type = Some(PIVOT_TABLE_PART_CONTENT_TYPE.to_string());

                // Add relationship from pivot table to cache definition
                pivot_part.relationships.add_new(
                    PIVOT_CACHE_DEFINITION_RELATIONSHIP_TYPE.to_string(),
                    format!("/xl/pivotCache/pivotCacheDefinition{pivot_id}.xml"),
                    TargetMode::Internal,
                );
                package.set_part(pivot_part);

                // Add relationship from worksheet to pivot table
                let relationship = sheet_part.relationships.add_new(
                    PIVOT_TABLE_RELATIONSHIP_TYPE.to_string(),
                    pivot_table_name.clone(),
                    TargetMode::Internal,
                );
                let _ = WorksheetPivotTablePartRef {
                    relationship_id: relationship.id.clone(),
                };
            }

            // ---- Images + Charts: collect refs first, write combined drawing at end ----
            let mut image_refs = Vec::new();
            let mut chart_rel_ids = Vec::new();
            let has_images = !sheet.images().is_empty();
            let has_charts = !sheet.charts().is_empty();

            // Ensure we have a drawing part if we need one (images, charts, or both).
            let mut drawing_part_and_uri: Option<(Part, PartUri)> = None;
            if has_images || has_charts {
                let drawing_id = next_drawing_id;
                next_drawing_id = next_drawing_id.checked_add(1).ok_or_else(|| {
                    XlsxError::InvalidWorkbookState(
                        "too many drawings to serialize in workbook".to_string(),
                    )
                })?;

                // Check if the sheet already has a drawing relationship (from source).
                let (drawing_uri, reused) = if let Some(dr) = drawing_ref.as_ref() {
                    let dr_rel = sheet_part
                        .relationships
                        .get_by_id(dr.relationship_id.as_str());
                    if let Some(rel) = dr_rel {
                        (
                            sheet_ref.part_uri.resolve_relative(rel.target.as_str())?,
                            true,
                        )
                    } else {
                        let uri =
                            PartUri::new(format!("/xl/drawings/drawing{drawing_id}.xml").as_str())?;
                        (uri, false)
                    }
                } else {
                    let uri =
                        PartUri::new(format!("/xl/drawings/drawing{drawing_id}.xml").as_str())?;
                    (uri, false)
                };

                if !reused {
                    let drawing_relationship = sheet_part.relationships.add_new(
                        DRAWING_RELATIONSHIP_TYPE.to_string(),
                        format!("../drawings/drawing{drawing_id}.xml"),
                        TargetMode::Internal,
                    );
                    drawing_ref = Some(WorksheetDrawingPartRef {
                        relationship_id: drawing_relationship.id.clone(),
                    });
                }

                let mut drawing_part =
                    if let Some(existing) = package.get_part(drawing_uri.as_str()) {
                        existing.clone()
                    } else {
                        let mut dp = Part::new_xml(drawing_uri.clone(), Vec::new());
                        dp.content_type = Some(DRAWING_PART_CONTENT_TYPE.to_string());
                        dp
                    };

                // Preserve non-image relationships from original drawing.
                if let Some(original) = &self.source_package {
                    if let Some(original_drawing_part) = original.get_part(drawing_uri.as_str()) {
                        for relationship in original_drawing_part.relationships.iter() {
                            if relationship.rel_type.as_str() != RelationshipType::IMAGE
                                && relationship.rel_type.as_str() != CHART_RELATIONSHIP_TYPE
                            {
                                record_passthrough_relationship(
                                    &mut drawing_passthrough_relationships,
                                    relationship.rel_type.as_str(),
                                );
                                drawing_part.relationships.add(relationship.clone());
                            }
                        }
                    }
                }

                // Add image media parts + relationships.
                if has_images {
                    for image in sheet.images() {
                        let image_id = next_image_id;
                        next_image_id = next_image_id.checked_add(1).ok_or_else(|| {
                            XlsxError::InvalidWorkbookState(
                                "too many images to serialize in workbook".to_string(),
                            )
                        })?;

                        let extension = media_extension_from_content_type(image.content_type());
                        let image_relationship = drawing_part.relationships.add_new(
                            RelationshipType::IMAGE.to_string(),
                            format!("../media/image{image_id}.{extension}"),
                            TargetMode::Internal,
                        );
                        image_refs.push(WorksheetImagePartRef {
                            relationship_id: image_relationship.id.clone(),
                        });

                        let image_uri = PartUri::new(
                            format!("/xl/media/image{image_id}.{extension}").as_str(),
                        )?;
                        let mut image_part = Part::new(image_uri, image.bytes().to_vec());
                        image_part.content_type = Some(image.content_type().to_string());
                        package.set_part(image_part);
                    }
                }

                // Add chart parts + relationships.
                if has_charts {
                    for chart in sheet.charts() {
                        let chart_id = next_chart_id;
                        next_chart_id = next_chart_id.checked_add(1).ok_or_else(|| {
                            XlsxError::InvalidWorkbookState(
                                "too many charts to serialize in workbook".to_string(),
                            )
                        })?;

                        let chart_relationship = drawing_part.relationships.add_new(
                            CHART_RELATIONSHIP_TYPE.to_string(),
                            format!("../charts/chart{chart_id}.xml"),
                            TargetMode::Internal,
                        );
                        chart_rel_ids.push(chart_relationship.id.clone());

                        let chart_uri =
                            PartUri::new(format!("/xl/charts/chart{chart_id}.xml").as_str())?;
                        let mut chart_part =
                            Part::new_xml(chart_uri, crate::chart_io::serialize_chart_xml(chart)?);
                        chart_part.content_type = Some(CHART_PART_CONTENT_TYPE.to_string());
                        package.set_part(chart_part);
                    }
                }

                drawing_part_and_uri = Some((drawing_part, drawing_uri));
            }

            // Write combined drawing XML with both image and chart anchors.
            if let Some((mut drawing_part, drawing_uri)) = drawing_part_and_uri {
                drawing_part.data = PartData::Xml(serialize_drawing_xml(
                    sheet.images(),
                    &image_refs,
                    sheet.charts(),
                    &chart_rel_ids,
                )?);
                package.set_part(drawing_part);
                let _ = drawing_uri;
            }

            // Create external hyperlink relationships.
            let mut hyperlink_refs = Vec::new();
            for hyperlink in sheet.hyperlinks() {
                if let Some(url) = hyperlink.url() {
                    let relationship = sheet_part.relationships.add_new(
                        HYPERLINK_RELATIONSHIP_TYPE.to_string(),
                        url.to_string(),
                        TargetMode::External,
                    );
                    hyperlink_refs.push(WorksheetHyperlinkRef {
                        relationship_id: relationship.id.clone(),
                    });
                }
            }

            // Create comments part if there are comments.
            if !sheet.comments().is_empty() {
                let comments_relationship = sheet_part.relationships.add_new(
                    COMMENTS_RELATIONSHIP_TYPE.to_string(),
                    format!("../comments{}.xml", sheet_ref.sheet_id),
                    TargetMode::Internal,
                );
                let _ = comments_relationship;

                let comments_uri =
                    PartUri::new(format!("/xl/comments{}.xml", sheet_ref.sheet_id).as_str())?;
                let mut comments_part =
                    Part::new_xml(comments_uri, serialize_comments_xml(sheet.comments())?);
                comments_part.content_type = Some(COMMENTS_PART_CONTENT_TYPE.to_string());
                package.set_part(comments_part);
            }

            sheet_part.data = PartData::Xml(serialize_worksheet_xml(
                sheet,
                &mut shared_strings,
                &table_refs,
                drawing_ref.as_ref(),
                &hyperlink_refs,
            )?);
            package.set_part(sheet_part);
        }

        workbook_part.relationships.add_new(
            RelationshipType::STYLES.to_string(),
            "styles.xml".to_string(),
            TargetMode::Internal,
        );
        let mut styles_part = Part::new_xml(
            PartUri::new(STYLES_PART_URI)?,
            serialize_styles_xml(
                &self.resolved_styles_for_write()?,
                &self.cell_style_xfs,
                &self.named_styles,
            )?,
        );
        styles_part.content_type = Some(ContentTypeValue::SPREADSHEET_STYLES.to_string());
        package.set_part(styles_part);

        if !shared_strings.is_empty() {
            let shared_strings_target = {
                let relationship = workbook_part.relationships.add_new(
                    RelationshipType::SHARED_STRINGS.to_string(),
                    "sharedStrings.xml".to_string(),
                    TargetMode::Internal,
                );
                relationship.target.clone()
            };
            let shared_strings_uri =
                workbook_uri.resolve_relative(shared_strings_target.as_str())?;
            let mut shared_strings_part = Part::new_xml(
                shared_strings_uri,
                serialize_shared_strings_xml(&shared_strings)?,
            );
            shared_strings_part.content_type = Some(ContentTypeValue::SHARED_STRINGS.to_string());
            package.set_part(shared_strings_part);
        }

        // Build pivot cache list: start with parsed caches, then add newly
        // created ones. New pivot tables created during save already add
        // relationships; we need to record them in the workbook XML too.
        let mut pivot_caches_to_write = self.pivot_caches.clone();

        // Collect relationship IDs for newly-created pivot cache definitions.
        // The save loop above creates workbook-level relationships for each
        // pivot table's cache definition. We scan the relationships we've
        // built so far and add any pivotCacheDefinition relationships that
        // aren't already recorded.
        for rel in workbook_part.relationships.iter() {
            if rel.rel_type == PIVOT_CACHE_DEFINITION_RELATIONSHIP_TYPE {
                let already_recorded = pivot_caches_to_write.iter().any(|(_, rid)| rid == &rel.id);
                if !already_recorded {
                    // Assign a new cacheId (max existing + 1)
                    let next_cache_id = pivot_caches_to_write
                        .iter()
                        .map(|(cid, _)| *cid)
                        .max()
                        .unwrap_or(0)
                        .saturating_add(1);
                    pivot_caches_to_write.push((next_cache_id, rel.id.clone()));
                }
            }
        }

        workbook_part.data = PartData::Xml(serialize_workbook_xml(
            &sheet_refs,
            &self.defined_names,
            self.workbook_protection.as_ref(),
            self.calc_settings.as_ref(),
            &self.file_version_attrs,
            &self.workbook_pr_attrs,
            &self.book_views,
            &pivot_caches_to_write,
            &self.custom_workbook_views,
        )?);
        package.set_part(workbook_part);
        if self.source_package.is_none() {
            package.relationships_mut().add_new(
                RelationshipType::WORKBOOK.to_string(),
                WORKBOOK_PART_URI.to_string(),
                TargetMode::Internal,
            );
        }

        emit_passthrough_relationship_warnings("workbook", &workbook_passthrough_relationships);
        emit_passthrough_relationship_warnings("worksheet", &worksheet_passthrough_relationships);
        emit_passthrough_relationship_warnings("drawing", &drawing_passthrough_relationships);

        Ok(package)
    }

    pub fn add_sheet(&mut self, name: impl Into<String>) -> &mut Worksheet {
        self.dirty = true;
        let name = name.into();

        if let Some(index) = self
            .sheets
            .iter()
            .position(|sheet| sheet.name() == name.as_str())
        {
            return &mut self.sheets[index];
        }

        self.sheets.push(Worksheet::new(name));
        let last_index = self.sheets.len() - 1;
        &mut self.sheets[last_index]
    }

    pub fn sheet(&self, name: &str) -> Option<&Worksheet> {
        self.sheets.iter().find(|sheet| sheet.name() == name)
    }

    pub fn sheet_mut(&mut self, name: &str) -> Option<&mut Worksheet> {
        let index = self.sheets.iter().position(|sheet| sheet.name() == name)?;
        self.dirty = true;
        self.sheets.get_mut(index)
    }

    pub fn sheet_names(&self) -> Vec<&str> {
        self.sheets.iter().map(Worksheet::name).collect()
    }

    pub fn worksheets(&self) -> &[Worksheet] {
        &self.sheets
    }

    pub fn worksheets_mut(&mut self) -> &mut [Worksheet] {
        self.dirty = true;
        &mut self.sheets
    }

    pub fn remove_sheet(&mut self, name: &str) -> Option<Worksheet> {
        let index = self.sheets.iter().position(|sheet| sheet.name() == name)?;
        self.dirty = true;
        Some(self.sheets.remove(index))
    }

    pub fn contains_sheet(&self, name: &str) -> bool {
        self.sheets.iter().any(|sheet| sheet.name() == name)
    }

    pub fn defined_names(&self) -> &[DefinedName] {
        self.defined_names.as_slice()
    }

    pub fn defined_name(&self, name: &str) -> Option<&DefinedName> {
        self.defined_names
            .iter()
            .find(|defined_name| defined_name.name() == name)
    }

    pub fn defined_name_mut(&mut self, name: &str) -> Option<&mut DefinedName> {
        let index = self
            .defined_names
            .iter()
            .position(|defined_name| defined_name.name() == name)?;
        self.dirty = true;
        self.defined_names.get_mut(index)
    }

    pub fn add_defined_name(
        &mut self,
        name: impl Into<String>,
        reference: impl Into<String>,
    ) -> &mut DefinedName {
        self.dirty = true;
        let candidate = DefinedName::new(name, reference);
        let name = candidate.name().to_string();

        if let Some(existing_index) = self
            .defined_names
            .iter()
            .position(|defined_name| defined_name.name() == name)
        {
            let existing = &mut self.defined_names[existing_index];
            existing.set_reference(candidate.reference());
            existing.clear_local_sheet_id();
            return existing;
        }

        self.defined_names.push(candidate);
        let index = self.defined_names.len() - 1;
        &mut self.defined_names[index]
    }

    pub fn remove_defined_name(&mut self, name: &str) -> Option<DefinedName> {
        let index = self
            .defined_names
            .iter()
            .position(|defined_name| defined_name.name() == name)?;
        self.dirty = true;
        Some(self.defined_names.remove(index))
    }

    /// Evaluates a formula string in the context of a specific sheet and cell.
    ///
    /// The formula may optionally start with `=`. Cell references are resolved
    /// against this workbook's sheets.
    ///
    /// ```ignore
    /// let mut wb = Workbook::new();
    /// let ws = wb.add_sheet("Sheet1");
    /// ws.cell_mut("A1")?.set_value(10);
    /// ws.cell_mut("A2")?.set_value(20);
    /// let result = wb.evaluate_formula("=SUM(A1:A2)", "Sheet1", 3, 1);
    /// assert_eq!(result, CellValue::Number(30.0));
    /// ```
    pub fn evaluate_formula(
        &self,
        formula: &str,
        sheet_name: &str,
        row: u32,
        col: u32,
    ) -> CellValue {
        let provider = crate::formula_bridge::WorkbookProvider::new(self);
        let ctx =
            offidized_formula::EvalContext::new(&provider, Some(sheet_name.to_string()), row, col);
        let result = offidized_formula::evaluate(formula, &ctx);
        crate::formula_bridge::scalar_to_cell_value(result.as_scalar())
    }

    pub fn styles(&self) -> &StyleTable {
        &self.styles
    }

    pub fn styles_mut(&mut self) -> &mut StyleTable {
        self.dirty = true;
        &mut self.styles
    }

    pub fn style(&self, style_id: u32) -> Option<&Style> {
        self.styles.style(style_id)
    }

    pub fn style_mut(&mut self, style_id: u32) -> Option<&mut Style> {
        if self.styles.style(style_id).is_some() {
            self.dirty = true;
        }
        self.styles.style_mut(style_id)
    }

    pub fn add_style(&mut self, style: Style) -> Result<u32> {
        self.dirty = true;
        self.styles.add_style(style)
    }

    pub fn clear_styles(&mut self) -> &mut Self {
        self.dirty = true;
        self.styles.clear_custom_styles();
        self
    }

    /// Returns workbook protection settings, if configured.
    pub fn workbook_protection(&self) -> Option<&WorkbookProtection> {
        self.workbook_protection.as_ref()
    }

    /// Sets workbook protection.
    pub fn set_workbook_protection(&mut self, protection: WorkbookProtection) -> &mut Self {
        self.workbook_protection = Some(protection);
        self.dirty = true;
        self
    }

    /// Clears workbook protection.
    pub fn clear_workbook_protection(&mut self) -> &mut Self {
        self.workbook_protection = None;
        self.dirty = true;
        self
    }

    /// Returns calculation settings, if configured.
    pub fn calc_settings(&self) -> Option<&CalculationSettings> {
        self.calc_settings.as_ref()
    }

    /// Sets calculation settings.
    pub fn set_calc_settings(&mut self, settings: CalculationSettings) -> &mut Self {
        self.calc_settings = Some(settings);
        self.dirty = true;
        self
    }

    /// Clears calculation settings.
    pub fn clear_calc_settings(&mut self) -> &mut Self {
        self.calc_settings = None;
        self.dirty = true;
        self
    }

    /// Returns whether the workbook uses the 1904 date system.
    ///
    /// When `true`, serial date 0 corresponds to 1904-01-01 (Mac epoch).
    /// When `false` (the default), serial date 1 corresponds to 1900-01-01 (Windows epoch).
    ///
    /// This is stored as the `date1904` attribute on the `<workbookPr>` element
    /// in `workbook.xml`.
    pub fn date1904(&self) -> bool {
        self.workbook_pr_attrs
            .iter()
            .any(|(k, v)| k == "date1904" && (v == "1" || v.eq_ignore_ascii_case("true")))
    }

    /// Returns the named cell styles parsed from styles.xml.
    pub fn named_styles(&self) -> &[NamedStyle] {
        &self.named_styles
    }

    /// Returns a mutable reference to the named styles list.
    pub fn named_styles_mut(&mut self) -> &mut Vec<NamedStyle> {
        self.dirty = true;
        &mut self.named_styles
    }

    /// Returns the cell style XF records parsed from styles.xml.
    pub fn cell_style_xfs(&self) -> &[CellStyleXf] {
        &self.cell_style_xfs
    }

    /// Returns a mutable reference to the cell style XF records.
    pub fn cell_style_xfs_mut(&mut self) -> &mut Vec<CellStyleXf> {
        self.dirty = true;
        &mut self.cell_style_xfs
    }

    /// Returns the 12 theme color hex strings (with `#` prefix) in canonical OOXML order.
    ///
    /// The order is: lt1, dk1, lt2, dk2, accent1-6, hlink, folHlink.
    /// If no theme was parsed, defaults to the standard Office theme palette.
    pub fn theme_colors(&self) -> &[String] {
        &self.theme_colors
    }

    /// Returns the major (heading) font name from the workbook theme, if present.
    pub fn major_font(&self) -> Option<&str> {
        self.major_font.as_deref()
    }

    /// Returns the minor (body) font name from the workbook theme, if present.
    pub fn minor_font(&self) -> Option<&str> {
        self.minor_font.as_deref()
    }

    /// Returns the custom indexed color palette from styles.xml, if present.
    ///
    /// When a workbook defines a custom indexed color palette, this returns the
    /// list of hex color strings (with `#` prefix). When absent, the standard
    /// legacy indexed color palette should be used.
    pub fn indexed_colors(&self) -> Option<&[String]> {
        self.indexed_colors.as_deref()
    }

    fn resolved_styles_for_write(&self) -> Result<Vec<Style>> {
        let mut styles = self.styles.styles().to_vec();
        if styles.is_empty() {
            styles.push(Style::new());
        }

        if let Some(max_style_id) = self.max_cell_style_id() {
            let max_style_index = usize::try_from(max_style_id).map_err(|_| {
                XlsxError::InvalidWorkbookState(format!(
                    "cell style_id `{max_style_id}` cannot be represented on this platform"
                ))
            })?;
            let required_len = max_style_index.checked_add(1).ok_or_else(|| {
                XlsxError::InvalidWorkbookState("style table length overflow".to_string())
            })?;
            if styles.len() < required_len {
                styles.resize(required_len, Style::new());
            }
        }

        Ok(styles)
    }

    fn ensure_style_capacity_for_cell_ids(&mut self) -> Result<()> {
        if let Some(max_style_id) = self.max_cell_style_id() {
            let max_style_index = usize::try_from(max_style_id).map_err(|_| {
                XlsxError::InvalidWorkbookState(format!(
                    "cell style_id `{max_style_id}` cannot be represented on this platform"
                ))
            })?;
            let required_len = max_style_index.checked_add(1).ok_or_else(|| {
                XlsxError::InvalidWorkbookState("style table length overflow".to_string())
            })?;
            self.styles.ensure_len(required_len);
        }
        Ok(())
    }

    fn max_cell_style_id(&self) -> Option<u32> {
        self.sheets
            .iter()
            .flat_map(|sheet| sheet.cells())
            .filter_map(|(_, cell)| cell.style_id())
            .max()
    }
}

fn resolve_workbook_part_uri(package: &Package) -> Result<PartUri> {
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

    let fallback = PartUri::new(WORKBOOK_PART_URI)?;
    if package.get_part(fallback.as_str()).is_some() {
        return Ok(fallback);
    }

    Err(XlsxError::UnsupportedPackage(
        "workbook part not found".to_string(),
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

fn is_rebuilt_workbook_relationship_type(rel_type: &str) -> bool {
    matches!(
        rel_type,
        RelationshipType::WORKSHEET | RelationshipType::STYLES | RelationshipType::SHARED_STRINGS
    )
}

fn is_rebuilt_worksheet_relationship_type(rel_type: &str) -> bool {
    rel_type == TABLE_RELATIONSHIP_TYPE
        || rel_type == DRAWING_RELATIONSHIP_TYPE
        || rel_type == HYPERLINK_RELATIONSHIP_TYPE
        || rel_type == COMMENTS_RELATIONSHIP_TYPE
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

fn parse_workbook_xml(xml: &[u8]) -> Result<ParsedWorkbook> {
    let mut reader = Reader::from_reader(Cursor::new(xml));
    reader.config_mut().trim_text(false);
    let mut buffer = Vec::new();
    let mut parsed_workbook = ParsedWorkbook::default();

    #[derive(Debug, Default)]
    struct DefinedNameState {
        name: Option<String>,
        local_sheet_id: Option<u32>,
        reference: Option<String>,
        unknown_attrs: Vec<(String, String)>,
    }

    let mut current_defined_name: Option<DefinedNameState> = None;
    let mut in_defined_name = false;
    let mut in_book_views = false;

    loop {
        match reader.read_event_into(&mut buffer)? {
            Event::Empty(ref event) | Event::Start(ref event)
                if local_name(event.name().as_ref()) == b"sheet" =>
            {
                let mut sheet_name = None;
                let mut relationship_id = None;
                let mut sheet_id = None;
                let mut visibility = SheetVisibility::Visible;

                for attribute in event.attributes().flatten() {
                    let key = local_name(attribute.key.as_ref());
                    // Use unescape_value to decode XML entities (e.g. &amp; → &)
                    let value = attribute
                        .unescape_value()
                        .unwrap_or_else(|_| String::from_utf8_lossy(attribute.value.as_ref()))
                        .into_owned();

                    match key {
                        b"name" => sheet_name = Some(value),
                        b"id" => {
                            let trimmed = value.trim();
                            if !trimmed.is_empty() {
                                relationship_id = Some(trimmed.to_string());
                            }
                        }
                        b"sheetId" => {
                            sheet_id = value.parse::<u32>().ok();
                        }
                        b"state" => {
                            visibility = SheetVisibility::from_xml_value(value.as_str());
                        }
                        _ => {}
                    }
                }

                let name = sheet_name.ok_or_else(|| {
                    XlsxError::UnsupportedPackage(
                        "workbook sheet missing `name` attribute".to_string(),
                    )
                })?;
                let Some(relationship_id) = relationship_id else {
                    tracing::warn!(
                        sheet_name = name.as_str(),
                        "workbook sheet missing `r:id` attribute; skipping sheet"
                    );
                    buffer.clear();
                    continue;
                };

                parsed_workbook.sheets.push(ParsedSheetRef {
                    sheet_id: sheet_id.unwrap_or(0),
                    name,
                    relationship_id,
                    visibility,
                });
            }
            Event::Empty(ref event) | Event::Start(ref event)
                if local_name(event.name().as_ref()) == b"workbookProtection" =>
            {
                let mut protection = WorkbookProtection::default();
                for attribute in event.attributes().flatten() {
                    let key = local_name(attribute.key.as_ref());
                    let value = String::from_utf8_lossy(attribute.value.as_ref()).into_owned();
                    match key {
                        b"lockStructure" => {
                            protection.lock_structure =
                                parse_xml_bool(value.as_str()).unwrap_or(false);
                        }
                        b"lockWindows" => {
                            protection.lock_windows =
                                parse_xml_bool(value.as_str()).unwrap_or(false);
                        }
                        b"workbookPassword" | b"workbookHashValue" => {
                            let trimmed = value.trim().to_string();
                            if !trimmed.is_empty() {
                                protection.password_hash = Some(trimmed);
                            }
                        }
                        _ => {}
                    }
                }
                if protection.has_metadata() {
                    parsed_workbook.workbook_protection = Some(protection);
                }
            }
            Event::Empty(ref event) | Event::Start(ref event)
                if local_name(event.name().as_ref()) == b"fileVersion" =>
            {
                parsed_workbook.file_version_attrs = event
                    .attributes()
                    .flatten()
                    .map(|a| {
                        (
                            String::from_utf8_lossy(a.key.as_ref()).into_owned(),
                            String::from_utf8_lossy(a.value.as_ref()).into_owned(),
                        )
                    })
                    .collect();
            }
            Event::Empty(ref event) | Event::Start(ref event)
                if local_name(event.name().as_ref()) == b"workbookPr" =>
            {
                parsed_workbook.workbook_pr_attrs = event
                    .attributes()
                    .flatten()
                    .map(|a| {
                        (
                            String::from_utf8_lossy(a.key.as_ref()).into_owned(),
                            String::from_utf8_lossy(a.value.as_ref()).into_owned(),
                        )
                    })
                    .collect();
            }
            Event::Start(ref event) if local_name(event.name().as_ref()) == b"bookViews" => {
                in_book_views = true;
            }
            Event::End(ref event) if local_name(event.name().as_ref()) == b"bookViews" => {
                in_book_views = false;
            }
            Event::Empty(ref event) | Event::Start(ref event)
                if in_book_views && local_name(event.name().as_ref()) == b"workbookView" =>
            {
                parsed_workbook.book_views.push(
                    event
                        .attributes()
                        .flatten()
                        .map(|a| {
                            (
                                String::from_utf8_lossy(a.key.as_ref()).into_owned(),
                                String::from_utf8_lossy(a.value.as_ref()).into_owned(),
                            )
                        })
                        .collect(),
                );
            }
            Event::Empty(ref event) | Event::Start(ref event)
                if local_name(event.name().as_ref()) == b"calcPr" =>
            {
                let mut settings = CalculationSettings::default();
                for attribute in event.attributes().flatten() {
                    let key = local_name(attribute.key.as_ref());
                    let value = String::from_utf8_lossy(attribute.value.as_ref()).into_owned();
                    match key {
                        b"calcMode" => {
                            let trimmed = value.trim().to_string();
                            if !trimmed.is_empty() {
                                settings.calc_mode = Some(trimmed);
                            }
                        }
                        b"calcId" => {
                            settings.calc_id = value.trim().parse::<u32>().ok();
                        }
                        b"fullCalcOnLoad" => {
                            settings.full_calc_on_load = parse_xml_bool(value.as_str());
                        }
                        _ => {}
                    }
                }
                if settings.has_metadata() {
                    parsed_workbook.calc_settings = Some(settings);
                }
            }
            Event::Start(ref event) if local_name(event.name().as_ref()) == b"definedName" => {
                let mut state = DefinedNameState::default();
                for attribute in event.attributes().flatten() {
                    let key = local_name(attribute.key.as_ref());
                    let value = String::from_utf8_lossy(attribute.value.as_ref()).into_owned();
                    let key_str = String::from_utf8_lossy(attribute.key.as_ref()).into_owned();
                    match key {
                        b"name" => state.name = Some(value.trim().to_string()),
                        b"localSheetId" => state.local_sheet_id = value.parse::<u32>().ok(),
                        _ => {
                            state.unknown_attrs.push((key_str, value));
                        }
                    }
                }
                current_defined_name = Some(state);
                in_defined_name = true;
            }
            Event::Empty(ref event) if local_name(event.name().as_ref()) == b"definedName" => {
                let mut name = None;
                let mut local_sheet_id = None;
                let mut unknown_attrs = Vec::new();

                for attribute in event.attributes().flatten() {
                    let key = local_name(attribute.key.as_ref());
                    let value = String::from_utf8_lossy(attribute.value.as_ref()).into_owned();
                    let key_str = String::from_utf8_lossy(attribute.key.as_ref()).into_owned();
                    match key {
                        b"name" => name = Some(value.trim().to_string()),
                        b"localSheetId" => local_sheet_id = value.parse::<u32>().ok(),
                        _ => {
                            unknown_attrs.push((key_str, value));
                        }
                    }
                }

                if let Some(name) = name.filter(|value| !value.is_empty()) {
                    let mut defined_name = DefinedName::new(name, "");
                    if let Some(local_sheet_id) = local_sheet_id {
                        defined_name.set_local_sheet_id(local_sheet_id);
                    }
                    defined_name.unknown_attrs = unknown_attrs;
                    parsed_workbook.defined_names.push(defined_name);
                }
            }
            Event::Text(ref event) if in_defined_name => {
                let text = event
                    .xml_content()
                    .map_err(quick_xml::Error::from)?
                    .into_owned();
                if let Some(state) = current_defined_name.as_mut() {
                    state
                        .reference
                        .get_or_insert_with(String::new)
                        .push_str(text.as_str());
                }
            }
            Event::CData(ref event) if in_defined_name => {
                let text = String::from_utf8_lossy(event.as_ref()).into_owned();
                if let Some(state) = current_defined_name.as_mut() {
                    state
                        .reference
                        .get_or_insert_with(String::new)
                        .push_str(text.as_str());
                }
            }
            Event::End(ref event) if local_name(event.name().as_ref()) == b"definedName" => {
                in_defined_name = false;
                if let Some(state) = current_defined_name.take() {
                    let Some(name) = state.name.filter(|value| !value.is_empty()) else {
                        buffer.clear();
                        continue;
                    };

                    let mut defined_name = DefinedName::new(
                        name,
                        state.reference.unwrap_or_default().trim().to_string(),
                    );
                    if let Some(local_sheet_id) = state.local_sheet_id {
                        defined_name.set_local_sheet_id(local_sheet_id);
                    }
                    defined_name.unknown_attrs = state.unknown_attrs;
                    parsed_workbook.defined_names.push(defined_name);
                }
            }
            // Parse <pivotCaches> — contains <pivotCache cacheId="..." r:id="..." />
            Event::Start(ref event) if local_name(event.name().as_ref()) == b"pivotCaches" => {
                // Read children until </pivotCaches>
                let mut pc_buf = Vec::new();
                loop {
                    match reader.read_event_into(&mut pc_buf)? {
                        Event::Empty(ref e) | Event::Start(ref e)
                            if local_name(e.name().as_ref()) == b"pivotCache" =>
                        {
                            let mut cache_id: Option<u32> = None;
                            let mut r_id: Option<String> = None;
                            for attr in e.attributes().flatten() {
                                let key = local_name(attr.key.as_ref());
                                let val = String::from_utf8_lossy(attr.value.as_ref()).into_owned();
                                match key {
                                    b"cacheId" => cache_id = val.parse::<u32>().ok(),
                                    b"id" => r_id = Some(val),
                                    _ => {}
                                }
                            }
                            if let (Some(cid), Some(rid)) = (cache_id, r_id) {
                                parsed_workbook.pivot_caches.push((cid, rid));
                            }
                        }
                        Event::End(ref e) if local_name(e.name().as_ref()) == b"pivotCaches" => {
                            break;
                        }
                        Event::Eof => break,
                        _ => {}
                    }
                    pc_buf.clear();
                }
            }
            // Capture <customWorkbookViews> as raw nodes for roundtrip
            Event::Start(ref event)
                if local_name(event.name().as_ref()) == b"customWorkbookViews" =>
            {
                let mut cwv_buf = Vec::new();
                loop {
                    match reader.read_event_into(&mut cwv_buf)? {
                        Event::Start(ref e) => {
                            parsed_workbook
                                .custom_workbook_views
                                .push(RawXmlNode::read_element(&mut reader, e)?);
                        }
                        Event::Empty(ref e) => {
                            parsed_workbook
                                .custom_workbook_views
                                .push(RawXmlNode::from_empty_element(e));
                        }
                        Event::End(ref e)
                            if local_name(e.name().as_ref()) == b"customWorkbookViews" =>
                        {
                            break;
                        }
                        Event::Eof => break,
                        _ => {}
                    }
                    cwv_buf.clear();
                }
            }
            Event::Eof => break,
            _ => {}
        }
        buffer.clear();
    }

    Ok(parsed_workbook)
}

fn parse_shared_strings_part(
    package: &Package,
    workbook_uri: &PartUri,
    workbook_part: &Part,
) -> Result<Option<Vec<SharedStringEntry>>> {
    let Some(relationship) = workbook_part
        .relationships
        .get_first_by_type(RelationshipType::SHARED_STRINGS)
    else {
        return Ok(None);
    };

    if relationship.target_mode != TargetMode::Internal {
        return Err(XlsxError::UnsupportedPackage(format!(
            "shared strings relationship `{}` is external",
            relationship.id
        )));
    }

    let shared_strings_uri = workbook_uri.resolve_relative(relationship.target.as_str())?;
    let shared_strings_part = package
        .get_part(shared_strings_uri.as_str())
        .ok_or_else(|| {
            XlsxError::UnsupportedPackage(format!(
                "missing shared strings part `{}` for relationship `{}`",
                shared_strings_uri.as_str(),
                relationship.id
            ))
        })?;

    Ok(Some(parse_shared_strings_xml(
        shared_strings_part.data.as_bytes(),
    )?))
}

/// Parse theme data from the theme part referenced by the workbook.
///
/// Resolves the theme relationship, reads the XML, and returns parsed theme data.
/// Returns defaults if no theme relationship exists or the part is missing.
fn parse_theme_part(
    package: &Package,
    workbook_uri: &PartUri,
    workbook_part: &Part,
) -> crate::theme::ParsedTheme {
    let Some(relationship) = workbook_part
        .relationships
        .get_first_by_type(RelationshipType::THEME)
    else {
        return crate::theme::ParsedTheme::default();
    };

    if relationship.target_mode != TargetMode::Internal {
        return crate::theme::ParsedTheme::default();
    }

    let Ok(theme_uri) = workbook_uri.resolve_relative(relationship.target.as_str()) else {
        return crate::theme::ParsedTheme::default();
    };

    let Some(theme_part) = package.get_part(theme_uri.as_str()) else {
        return crate::theme::ParsedTheme::default();
    };

    crate::theme::parse_theme_xml(theme_part.data.as_bytes())
}

#[allow(clippy::type_complexity)]
fn parse_styles_part(
    package: &Package,
    workbook_uri: &PartUri,
    workbook_part: &Part,
) -> Result<
    Option<(
        StyleTable,
        Vec<CellStyleXf>,
        Vec<NamedStyle>,
        Option<Vec<String>>,
    )>,
> {
    let Some(relationship) = workbook_part
        .relationships
        .get_first_by_type(RelationshipType::STYLES)
    else {
        return Ok(None);
    };

    if relationship.target_mode != TargetMode::Internal {
        return Err(XlsxError::UnsupportedPackage(format!(
            "styles relationship `{}` is external",
            relationship.id
        )));
    }

    let styles_uri = workbook_uri.resolve_relative(relationship.target.as_str())?;
    let styles_part = package.get_part(styles_uri.as_str()).ok_or_else(|| {
        XlsxError::UnsupportedPackage(format!(
            "missing styles part `{}` for relationship `{}`",
            styles_uri.as_str(),
            relationship.id
        ))
    })?;

    Ok(Some(parse_styles_xml(styles_part.data.as_bytes())?))
}

#[allow(clippy::type_complexity)]
fn parse_styles_xml(
    xml: &[u8],
) -> Result<(
    StyleTable,
    Vec<CellStyleXf>,
    Vec<NamedStyle>,
    Option<Vec<String>>,
)> {
    #[derive(Debug, Clone, Copy, PartialEq, Eq)]
    enum BorderSideKind {
        Left,
        Right,
        Top,
        Bottom,
        Diagonal,
    }

    impl BorderSideKind {
        fn from_xml_name(name: &[u8]) -> Option<Self> {
            match local_name(name) {
                b"left" => Some(Self::Left),
                b"right" => Some(Self::Right),
                b"top" => Some(Self::Top),
                b"bottom" => Some(Self::Bottom),
                b"diagonal" => Some(Self::Diagonal),
                _ => None,
            }
        }

        fn apply_to_border(self, border: &mut Border, side: BorderSide) {
            match self {
                Self::Left => {
                    border.set_left(side);
                }
                Self::Right => {
                    border.set_right(side);
                }
                Self::Top => {
                    border.set_top(side);
                }
                Self::Bottom => {
                    border.set_bottom(side);
                }
                Self::Diagonal => {
                    border.set_diagonal(side);
                }
            }
        }
    }

    #[derive(Debug, Default)]
    struct XfState {
        num_fmt_id: u32,
        font_id: u32,
        fill_id: u32,
        border_id: u32,
        alignment: Option<Alignment>,
        protection: Option<CellProtection>,
    }

    impl XfState {
        fn from_xml_xf(event: &BytesStart<'_>) -> Self {
            let mut state = Self::default();
            for attribute in event.attributes().flatten() {
                let key = local_name(attribute.key.as_ref());
                let value = String::from_utf8_lossy(attribute.value.as_ref()).into_owned();
                match key {
                    b"numFmtId" => state.num_fmt_id = value.trim().parse::<u32>().unwrap_or(0),
                    b"fontId" => state.font_id = value.trim().parse::<u32>().unwrap_or(0),
                    b"fillId" => state.fill_id = value.trim().parse::<u32>().unwrap_or(0),
                    b"borderId" => state.border_id = value.trim().parse::<u32>().unwrap_or(0),
                    _ => {}
                }
            }
            state
        }

        fn into_style(
            self,
            number_formats: &BTreeMap<u32, String>,
            fonts: &[Font],
            fills: &[Fill],
            gradient_fills: &[Option<GradientFill>],
            borders: &[Border],
        ) -> Style {
            let mut style = Style::new();
            if let Some(number_format) = number_formats.get(&self.num_fmt_id) {
                style.set_number_format(number_format.to_string());
            }
            if let Some(alignment) = self.alignment.filter(Alignment::has_metadata) {
                style.set_alignment(alignment);
            }
            if let Some(protection) = self.protection.filter(|p| p.has_metadata()) {
                style.set_protection(protection);
            }
            if let Some(font) = usize::try_from(self.font_id)
                .ok()
                .and_then(|index| fonts.get(index))
                .filter(|font| font.has_metadata())
            {
                style.set_font(font.clone());
            }
            if let Some(fill) = usize::try_from(self.fill_id)
                .ok()
                .and_then(|index| fills.get(index))
                .filter(|fill| fill.has_metadata())
            {
                style.set_fill(fill.clone());
            }
            // Apply gradient fill if present at the same fill index
            if let Some(Some(gf)) = usize::try_from(self.fill_id)
                .ok()
                .and_then(|index| gradient_fills.get(index))
            {
                style.set_gradient_fill(gf.clone());
            }
            if let Some(border) = usize::try_from(self.border_id)
                .ok()
                .and_then(|index| borders.get(index))
                .filter(|border| border.has_metadata())
            {
                style.set_border(border.clone());
            }
            style
        }
    }

    let mut reader = Reader::from_reader(Cursor::new(xml));
    reader.config_mut().trim_text(false);
    let mut buffer = Vec::new();
    let mut number_formats = BTreeMap::<u32, String>::new();
    let mut fonts = Vec::<Font>::new();
    let mut fills = Vec::<Fill>::new();
    let mut gradient_fills = Vec::<Option<GradientFill>>::new();
    let mut borders = Vec::<Border>::new();
    let mut styles = Vec::new();
    let mut cell_style_xfs = Vec::<CellStyleXf>::new();
    let mut named_styles = Vec::<NamedStyle>::new();
    let mut in_cell_xfs = false;
    let mut in_cell_style_xfs = false;
    let mut current_xf: Option<XfState> = None;
    let mut current_font: Option<Font> = None;
    let mut current_fill: Option<Fill> = None;
    let mut current_fill_fg_ref: Option<ColorReference> = None;
    let mut current_fill_bg_ref: Option<ColorReference> = None;
    let mut current_gradient_fill: Option<GradientFill> = None;
    let mut current_border: Option<Border> = None;
    let mut current_border_side: Option<(BorderSideKind, BorderSide)> = None;
    let mut indexed_colors: Option<Vec<String>> = None;
    let mut in_indexed_colors = false;

    loop {
        match reader.read_event_into(&mut buffer)? {
            Event::Start(ref event) | Event::Empty(ref event)
                if local_name(event.name().as_ref()) == b"numFmt" =>
            {
                let mut num_fmt_id = None;
                let mut format_code = None;
                for attribute in event.attributes().flatten() {
                    let key = local_name(attribute.key.as_ref());
                    let value = String::from_utf8_lossy(attribute.value.as_ref()).into_owned();
                    match key {
                        b"numFmtId" => num_fmt_id = value.trim().parse::<u32>().ok(),
                        b"formatCode" => format_code = Some(value),
                        _ => {}
                    }
                }

                if let (Some(num_fmt_id), Some(format_code)) = (num_fmt_id, format_code) {
                    number_formats.insert(num_fmt_id, format_code);
                }
            }
            Event::Start(ref event) if local_name(event.name().as_ref()) == b"font" => {
                current_font = Some(Font::new());
            }
            Event::Empty(ref event) if local_name(event.name().as_ref()) == b"font" => {
                fonts.push(Font::new());
            }
            Event::Start(ref event)
                if current_font.is_some() && local_name(event.name().as_ref()) == b"name" =>
            {
                if let (Some(font), Some(name)) =
                    (current_font.as_mut(), parse_xml_attr(event, b"val"))
                {
                    font.set_name(name);
                }
            }
            Event::Empty(ref event)
                if current_font.is_some() && local_name(event.name().as_ref()) == b"name" =>
            {
                if let (Some(font), Some(name)) =
                    (current_font.as_mut(), parse_xml_attr(event, b"val"))
                {
                    font.set_name(name);
                }
            }
            Event::Start(ref event)
                if current_font.is_some() && local_name(event.name().as_ref()) == b"sz" =>
            {
                if let (Some(font), Some(size)) =
                    (current_font.as_mut(), parse_xml_attr(event, b"val"))
                {
                    font.set_size(size);
                }
            }
            Event::Empty(ref event)
                if current_font.is_some() && local_name(event.name().as_ref()) == b"sz" =>
            {
                if let (Some(font), Some(size)) =
                    (current_font.as_mut(), parse_xml_attr(event, b"val"))
                {
                    font.set_size(size);
                }
            }
            Event::Start(ref event)
                if current_font.is_some() && local_name(event.name().as_ref()) == b"b" =>
            {
                if let Some(font) = current_font.as_mut() {
                    font.set_bold(parse_font_toggle(event));
                }
            }
            Event::Empty(ref event)
                if current_font.is_some() && local_name(event.name().as_ref()) == b"b" =>
            {
                if let Some(font) = current_font.as_mut() {
                    font.set_bold(parse_font_toggle(event));
                }
            }
            Event::Start(ref event)
                if current_font.is_some() && local_name(event.name().as_ref()) == b"i" =>
            {
                if let Some(font) = current_font.as_mut() {
                    font.set_italic(parse_font_toggle(event));
                }
            }
            Event::Empty(ref event)
                if current_font.is_some() && local_name(event.name().as_ref()) == b"i" =>
            {
                if let Some(font) = current_font.as_mut() {
                    font.set_italic(parse_font_toggle(event));
                }
            }
            Event::Start(ref event)
                if current_font.is_some() && local_name(event.name().as_ref()) == b"u" =>
            {
                if let Some(font) = current_font.as_mut() {
                    font.set_underline(parse_font_toggle(event));
                }
            }
            Event::Empty(ref event)
                if current_font.is_some() && local_name(event.name().as_ref()) == b"u" =>
            {
                if let Some(font) = current_font.as_mut() {
                    font.set_underline(parse_font_toggle(event));
                }
            }
            Event::Start(ref event) | Event::Empty(ref event)
                if current_font.is_some() && local_name(event.name().as_ref()) == b"strike" =>
            {
                if let Some(font) = current_font.as_mut() {
                    font.set_strikethrough(parse_font_toggle(event));
                }
            }
            Event::Start(ref event) | Event::Empty(ref event)
                if current_font.is_some() && local_name(event.name().as_ref()) == b"dStrike" =>
            {
                if let Some(font) = current_font.as_mut() {
                    font.set_double_strikethrough(parse_font_toggle(event));
                }
            }
            Event::Start(ref event) | Event::Empty(ref event)
                if current_font.is_some() && local_name(event.name().as_ref()) == b"shadow" =>
            {
                if let Some(font) = current_font.as_mut() {
                    font.set_shadow(parse_font_toggle(event));
                }
            }
            Event::Start(ref event) | Event::Empty(ref event)
                if current_font.is_some() && local_name(event.name().as_ref()) == b"outline" =>
            {
                if let Some(font) = current_font.as_mut() {
                    font.set_outline(parse_font_toggle(event));
                }
            }
            Event::Start(ref event) | Event::Empty(ref event)
                if current_font.is_some() && local_name(event.name().as_ref()) == b"vertAlign" =>
            {
                if let Some(font) = current_font.as_mut() {
                    if let Some(val) = parse_xml_attr(event, b"val") {
                        match val.as_str() {
                            "subscript" => {
                                font.set_subscript(true);
                            }
                            "superscript" => {
                                font.set_superscript(true);
                            }
                            _ => {}
                        }
                        if let Some(va) = FontVerticalAlign::from_xml_value(val.as_str()) {
                            font.set_vertical_align(va);
                        }
                    }
                }
            }
            Event::Start(ref event) | Event::Empty(ref event)
                if current_font.is_some() && local_name(event.name().as_ref()) == b"scheme" =>
            {
                if let Some(font) = current_font.as_mut() {
                    if let Some(val) = parse_xml_attr(event, b"val") {
                        if let Some(scheme) = FontScheme::from_xml_value(val.as_str()) {
                            font.set_font_scheme(scheme);
                        }
                    }
                }
            }
            Event::Start(ref event)
                if current_font.is_some() && local_name(event.name().as_ref()) == b"color" =>
            {
                if let (Some(font), Some(color_ref)) =
                    (current_font.as_mut(), parse_color_ref(event))
                {
                    if let Some(rgb) = color_ref.rgb() {
                        font.set_color(rgb.to_string());
                    }
                    font.set_color_ref(color_ref);
                }
            }
            Event::Empty(ref event)
                if current_font.is_some() && local_name(event.name().as_ref()) == b"color" =>
            {
                if let (Some(font), Some(color_ref)) =
                    (current_font.as_mut(), parse_color_ref(event))
                {
                    if let Some(rgb) = color_ref.rgb() {
                        font.set_color(rgb.to_string());
                    }
                    font.set_color_ref(color_ref);
                }
            }
            Event::End(ref event) if local_name(event.name().as_ref()) == b"font" => {
                if let Some(font) = current_font.take() {
                    fonts.push(font);
                }
            }
            Event::Start(ref event) if local_name(event.name().as_ref()) == b"fill" => {
                current_fill = Some(Fill::new());
            }
            Event::Empty(ref event) if local_name(event.name().as_ref()) == b"fill" => {
                fills.push(Fill::new());
                gradient_fills.push(None);
            }
            Event::Start(ref event)
                if current_fill.is_some()
                    && local_name(event.name().as_ref()) == b"patternFill" =>
            {
                if let (Some(fill), Some(pattern)) =
                    (current_fill.as_mut(), parse_xml_attr(event, b"patternType"))
                {
                    fill.set_pattern(pattern);
                }
            }
            Event::Empty(ref event)
                if current_fill.is_some()
                    && local_name(event.name().as_ref()) == b"patternFill" =>
            {
                if let (Some(fill), Some(pattern)) =
                    (current_fill.as_mut(), parse_xml_attr(event, b"patternType"))
                {
                    fill.set_pattern(pattern);
                }
            }
            Event::Start(ref event)
                if current_fill.is_some() && local_name(event.name().as_ref()) == b"fgColor" =>
            {
                if let Some(color_ref) = parse_color_ref(event) {
                    if let Some(fill) = current_fill.as_mut() {
                        if let Some(rgb) = color_ref.rgb() {
                            fill.set_foreground_color(rgb.to_string());
                        }
                    }
                    current_fill_fg_ref = Some(color_ref);
                }
            }
            Event::Empty(ref event)
                if current_fill.is_some() && local_name(event.name().as_ref()) == b"fgColor" =>
            {
                if let Some(color_ref) = parse_color_ref(event) {
                    if let Some(fill) = current_fill.as_mut() {
                        if let Some(rgb) = color_ref.rgb() {
                            fill.set_foreground_color(rgb.to_string());
                        }
                    }
                    current_fill_fg_ref = Some(color_ref);
                }
            }
            Event::Start(ref event)
                if current_fill.is_some() && local_name(event.name().as_ref()) == b"bgColor" =>
            {
                if let Some(color_ref) = parse_color_ref(event) {
                    if let Some(fill) = current_fill.as_mut() {
                        if let Some(rgb) = color_ref.rgb() {
                            fill.set_background_color(rgb.to_string());
                        }
                    }
                    current_fill_bg_ref = Some(color_ref);
                }
            }
            Event::Empty(ref event)
                if current_fill.is_some() && local_name(event.name().as_ref()) == b"bgColor" =>
            {
                if let Some(color_ref) = parse_color_ref(event) {
                    if let Some(fill) = current_fill.as_mut() {
                        if let Some(rgb) = color_ref.rgb() {
                            fill.set_background_color(rgb.to_string());
                        }
                    }
                    current_fill_bg_ref = Some(color_ref);
                }
            }
            // Parse <gradientFill> inside <fill>
            Event::Start(ref event)
                if current_fill.is_some()
                    && local_name(event.name().as_ref()) == b"gradientFill" =>
            {
                // Read attributes: type, degree, left, right, top, bottom
                let mut gf_type = GradientFillType::Linear;
                let mut gf_degree: Option<f64> = None;
                let mut gf_left: Option<f64> = None;
                let mut gf_right: Option<f64> = None;
                let mut gf_top: Option<f64> = None;
                let mut gf_bottom: Option<f64> = None;
                for attr in event.attributes().flatten() {
                    let key = local_name(attr.key.as_ref());
                    let v = String::from_utf8_lossy(attr.value.as_ref()).into_owned();
                    match key {
                        b"type" => {
                            gf_type = GradientFillType::from_xml_value(v.trim())
                                .unwrap_or(GradientFillType::Linear);
                        }
                        b"degree" => gf_degree = v.trim().parse::<f64>().ok(),
                        b"left" => gf_left = v.trim().parse::<f64>().ok(),
                        b"right" => gf_right = v.trim().parse::<f64>().ok(),
                        b"top" => gf_top = v.trim().parse::<f64>().ok(),
                        b"bottom" => gf_bottom = v.trim().parse::<f64>().ok(),
                        _ => {}
                    }
                }
                let mut gf = match gf_type {
                    GradientFillType::Linear => {
                        let mut g = GradientFill::linear(gf_degree.unwrap_or(0.0));
                        if gf_degree.is_none() {
                            g.clear_degree();
                        }
                        g
                    }
                    GradientFillType::Path => {
                        let mut p = GradientFill::path();
                        if let Some(d) = gf_degree {
                            p.set_degree(d);
                        }
                        p
                    }
                };
                if let Some(l) = gf_left {
                    gf.set_left(l);
                }
                if let Some(r) = gf_right {
                    gf.set_right(r);
                }
                if let Some(t) = gf_top {
                    gf.set_top(t);
                }
                if let Some(b) = gf_bottom {
                    gf.set_bottom(b);
                }

                // Read child <stop> elements until </gradientFill>
                let mut gf_depth = 0_u32;
                let mut gf_stop_pos: Option<f64> = None;
                loop {
                    let gf_event = reader.read_event_into(&mut buffer)?;
                    match &gf_event {
                        Event::Start(_) => gf_depth += 1,
                        Event::End(_) => {
                            if gf_depth == 0 {
                                break;
                            }
                            gf_depth = gf_depth.saturating_sub(1);
                            // Leaving a <stop> element — clear position
                            gf_stop_pos = None;
                        }
                        _ => {}
                    }
                    match gf_event {
                        Event::Start(ref e) | Event::Empty(ref e) => {
                            let name_bytes = e.name();
                            let local = local_name(name_bytes.as_ref());
                            if local == b"stop" {
                                // Read position attribute
                                for attr in e.attributes().flatten() {
                                    if local_name(attr.key.as_ref()) == b"position" {
                                        let v = String::from_utf8_lossy(attr.value.as_ref());
                                        gf_stop_pos = v.trim().parse::<f64>().ok();
                                    }
                                }
                            } else if local == b"color" && gf_stop_pos.is_some() {
                                // Parse color within a <stop>
                                if let Some(pos) = gf_stop_pos {
                                    if let Some(color_ref) = parse_color_ref(e) {
                                        gf.add_stop_with_color_ref(pos, color_ref);
                                    }
                                }
                            }
                        }
                        Event::Eof => break,
                        _ => {}
                    }
                    buffer.clear();
                }
                current_gradient_fill = Some(gf);
            }
            Event::End(ref event) if local_name(event.name().as_ref()) == b"fill" => {
                if let Some(mut fill) = current_fill.take() {
                    // Build a structured PatternFill from the string-based fields if a pattern is set.
                    if let Some(pattern_str) = fill.pattern() {
                        if let Some(pattern_type) = PatternFillType::from_xml_value(pattern_str) {
                            let mut pf = PatternFill::new(pattern_type);
                            if let Some(fg_ref) = current_fill_fg_ref.take() {
                                pf.set_fg_color(fg_ref);
                            } else if let Some(fg) = fill.foreground_color() {
                                pf.set_fg_color(ColorReference::from_rgb(fg));
                            }
                            if let Some(bg_ref) = current_fill_bg_ref.take() {
                                pf.set_bg_color(bg_ref);
                            } else if let Some(bg) = fill.background_color() {
                                pf.set_bg_color(ColorReference::from_rgb(bg));
                            }
                            fill.set_pattern_fill(pf);
                        }
                    }
                    // Clear any remaining refs in case they weren't consumed
                    current_fill_fg_ref = None;
                    current_fill_bg_ref = None;
                    gradient_fills.push(current_gradient_fill.take());
                    fills.push(fill);
                }
            }
            Event::Start(ref event) if local_name(event.name().as_ref()) == b"border" => {
                let mut border = Border::new();
                for attribute in event.attributes().flatten() {
                    let key = local_name(attribute.key.as_ref());
                    let value = String::from_utf8_lossy(attribute.value.as_ref()).into_owned();
                    match key {
                        b"diagonalUp" => {
                            if let Some(v) = parse_xml_bool(value.as_str()) {
                                border.set_diagonal_up(v);
                            }
                        }
                        b"diagonalDown" => {
                            if let Some(v) = parse_xml_bool(value.as_str()) {
                                border.set_diagonal_down(v);
                            }
                        }
                        _ => {}
                    }
                }
                current_border = Some(border);
                current_border_side = None;
            }
            Event::Empty(ref event) if local_name(event.name().as_ref()) == b"border" => {
                borders.push(Border::new());
            }
            Event::Start(ref event)
                if current_border.is_some()
                    && BorderSideKind::from_xml_name(event.name().as_ref()).is_some() =>
            {
                let mut side = BorderSide::new();
                if let Some(style) = parse_xml_attr(event, b"style") {
                    side.set_style(style);
                }
                if let Some(side_kind) = BorderSideKind::from_xml_name(event.name().as_ref()) {
                    current_border_side = Some((side_kind, side));
                }
            }
            Event::Empty(ref event)
                if current_border.is_some()
                    && BorderSideKind::from_xml_name(event.name().as_ref()).is_some() =>
            {
                let mut side = BorderSide::new();
                if let Some(style) = parse_xml_attr(event, b"style") {
                    side.set_style(style);
                }
                if let Some(border) = current_border.as_mut() {
                    if side.has_metadata() {
                        if let Some(side_kind) =
                            BorderSideKind::from_xml_name(event.name().as_ref())
                        {
                            side_kind.apply_to_border(border, side);
                        }
                    }
                }
            }
            Event::Start(ref event)
                if current_border_side.is_some()
                    && local_name(event.name().as_ref()) == b"color" =>
            {
                if let Some((_, side)) = current_border_side.as_mut() {
                    if let Some(color_ref) = parse_color_ref(event) {
                        if let Some(rgb) = color_ref.rgb() {
                            side.set_color(rgb.to_string());
                        }
                        side.set_color_ref(color_ref);
                    }
                }
            }
            Event::Empty(ref event)
                if current_border_side.is_some()
                    && local_name(event.name().as_ref()) == b"color" =>
            {
                if let Some((_, side)) = current_border_side.as_mut() {
                    if let Some(color_ref) = parse_color_ref(event) {
                        if let Some(rgb) = color_ref.rgb() {
                            side.set_color(rgb.to_string());
                        }
                        side.set_color_ref(color_ref);
                    }
                }
            }
            Event::End(ref event)
                if current_border.is_some()
                    && BorderSideKind::from_xml_name(event.name().as_ref()).is_some() =>
            {
                if let (Some(border), Some((side_kind, side))) =
                    (current_border.as_mut(), current_border_side.take())
                {
                    if side.has_metadata() {
                        side_kind.apply_to_border(border, side);
                    }
                }
            }
            Event::End(ref event) if local_name(event.name().as_ref()) == b"border" => {
                if let Some(border) = current_border.take() {
                    borders.push(border);
                }
                current_border_side = None;
            }
            Event::Start(ref event) if local_name(event.name().as_ref()) == b"cellStyleXfs" => {
                in_cell_style_xfs = true;
            }
            Event::End(ref event) if local_name(event.name().as_ref()) == b"cellStyleXfs" => {
                in_cell_style_xfs = false;
            }
            Event::Start(ref event)
                if in_cell_style_xfs && local_name(event.name().as_ref()) == b"xf" =>
            {
                let xf = XfState::from_xml_xf(event);
                let mut csxf = CellStyleXf::new();
                csxf.set_num_fmt_id(xf.num_fmt_id);
                csxf.set_font_id(xf.font_id);
                csxf.set_fill_id(xf.fill_id);
                csxf.set_border_id(xf.border_id);
                cell_style_xfs.push(csxf);
            }
            Event::Empty(ref event)
                if in_cell_style_xfs && local_name(event.name().as_ref()) == b"xf" =>
            {
                let xf = XfState::from_xml_xf(event);
                let mut csxf = CellStyleXf::new();
                csxf.set_num_fmt_id(xf.num_fmt_id);
                csxf.set_font_id(xf.font_id);
                csxf.set_fill_id(xf.fill_id);
                csxf.set_border_id(xf.border_id);
                cell_style_xfs.push(csxf);
            }
            Event::Start(ref event) if local_name(event.name().as_ref()) == b"cellXfs" => {
                in_cell_xfs = true;
            }
            Event::End(ref event) if local_name(event.name().as_ref()) == b"cellXfs" => {
                in_cell_xfs = false;
                current_xf = None;
            }
            Event::Start(ref event)
                if in_cell_xfs && local_name(event.name().as_ref()) == b"xf" =>
            {
                current_xf = Some(XfState::from_xml_xf(event));
            }
            Event::Empty(ref event)
                if in_cell_xfs && local_name(event.name().as_ref()) == b"xf" =>
            {
                let xf = XfState::from_xml_xf(event);
                styles.push(xf.into_style(
                    &number_formats,
                    &fonts,
                    &fills,
                    &gradient_fills,
                    &borders,
                ));
            }
            Event::Start(ref event) | Event::Empty(ref event)
                if local_name(event.name().as_ref()) == b"cellStyle" =>
            {
                let mut name = String::new();
                let mut xf_id = 0_u32;
                let mut builtin_id: Option<u32> = None;
                for attribute in event.attributes().flatten() {
                    let key = local_name(attribute.key.as_ref());
                    let value = String::from_utf8_lossy(attribute.value.as_ref()).into_owned();
                    match key {
                        b"name" => name = value,
                        b"xfId" => xf_id = value.trim().parse::<u32>().unwrap_or(0),
                        b"builtinId" => builtin_id = value.trim().parse::<u32>().ok(),
                        _ => {}
                    }
                }
                let mut ns = NamedStyle::new(name, xf_id);
                if let Some(bid) = builtin_id {
                    ns.set_builtin_id(bid);
                }
                named_styles.push(ns);
            }
            Event::Start(ref event)
                if current_xf.is_some() && local_name(event.name().as_ref()) == b"alignment" =>
            {
                if let Some(xf) = current_xf.as_mut() {
                    xf.alignment = parse_style_alignment(event);
                }
            }
            Event::Empty(ref event)
                if current_xf.is_some() && local_name(event.name().as_ref()) == b"alignment" =>
            {
                if let Some(xf) = current_xf.as_mut() {
                    xf.alignment = parse_style_alignment(event);
                }
            }
            Event::Start(ref event) | Event::Empty(ref event)
                if current_xf.is_some() && local_name(event.name().as_ref()) == b"protection" =>
            {
                if let Some(xf) = current_xf.as_mut() {
                    xf.protection = parse_cell_protection(event);
                }
            }
            Event::End(ref event) if local_name(event.name().as_ref()) == b"xf" => {
                if let Some(xf) = current_xf.take() {
                    styles.push(xf.into_style(
                        &number_formats,
                        &fonts,
                        &fills,
                        &gradient_fills,
                        &borders,
                    ));
                }
            }
            Event::Start(ref event) if local_name(event.name().as_ref()) == b"indexedColors" => {
                in_indexed_colors = true;
                indexed_colors = Some(Vec::new());
            }
            Event::End(ref event) if local_name(event.name().as_ref()) == b"indexedColors" => {
                in_indexed_colors = false;
            }
            Event::Start(ref event) | Event::Empty(ref event)
                if in_indexed_colors && local_name(event.name().as_ref()) == b"rgbColor" =>
            {
                if let Some(ref mut colors) = indexed_colors {
                    let rgb = parse_xml_attr(event, b"rgb")
                        .map(|v| {
                            // ARGB format: strip alpha prefix if 8 chars
                            let hex = v.to_uppercase();
                            if hex.len() == 8 {
                                format!("#{}", &hex[2..])
                            } else {
                                format!("#{hex}")
                            }
                        })
                        .unwrap_or_else(|| "#000000".to_string());
                    colors.push(rgb);
                }
            }
            Event::Eof => break,
            _ => {}
        }
        buffer.clear();
    }

    Ok((
        StyleTable::from_styles(styles),
        cell_style_xfs,
        named_styles,
        indexed_colors,
    ))
}

fn parse_xml_attr(event: &BytesStart<'_>, attribute_name: &[u8]) -> Option<String> {
    for attribute in event.attributes().flatten() {
        if local_name(attribute.key.as_ref()) != attribute_name {
            continue;
        }
        let value = String::from_utf8_lossy(attribute.value.as_ref()).into_owned();
        let value = value.trim();
        if value.is_empty() {
            return None;
        }
        return Some(value.to_string());
    }

    None
}

fn parse_font_toggle(event: &BytesStart<'_>) -> bool {
    parse_xml_attr(event, b"val")
        .as_deref()
        .and_then(parse_xml_bool)
        .unwrap_or(true)
}

fn parse_color_ref(event: &BytesStart<'_>) -> Option<ColorReference> {
    let mut color = ColorReference::empty();
    let mut has_any = false;
    for attribute in event.attributes().flatten() {
        let key = local_name(attribute.key.as_ref());
        let value = String::from_utf8_lossy(attribute.value.as_ref()).into_owned();
        match key {
            b"rgb" => {
                color.set_rgb(value);
                has_any = true;
            }
            b"theme" => {
                if let Ok(idx) = value.trim().parse::<u32>() {
                    if let Some(theme) = ThemeColor::from_index(idx) {
                        color.set_theme(theme);
                        has_any = true;
                    }
                }
            }
            b"tint" => {
                if let Ok(t) = value.trim().parse::<f64>() {
                    color.set_tint(t);
                    has_any = true;
                }
            }
            b"indexed" => {
                if let Ok(idx) = value.trim().parse::<u32>() {
                    color.set_indexed(idx);
                    has_any = true;
                }
            }
            b"auto" => {
                if value.trim() == "1" || value.trim() == "true" {
                    color.set_auto(true);
                    has_any = true;
                }
            }
            _ => {}
        }
    }
    if has_any {
        Some(color)
    } else {
        None
    }
}

fn parse_style_alignment(event: &BytesStart<'_>) -> Option<Alignment> {
    let mut alignment = Alignment::new();

    for attribute in event.attributes().flatten() {
        let key = local_name(attribute.key.as_ref());
        let value = String::from_utf8_lossy(attribute.value.as_ref()).into_owned();
        match key {
            b"horizontal" => {
                if let Some(horizontal) = HorizontalAlignment::from_xml_value(value.trim()) {
                    alignment.set_horizontal(horizontal);
                }
            }
            b"vertical" => {
                if let Some(vertical) = VerticalAlignment::from_xml_value(value.trim()) {
                    alignment.set_vertical(vertical);
                }
            }
            b"wrapText" => {
                if let Some(wrap_text) = parse_xml_bool(value.as_str()) {
                    alignment.set_wrap_text(wrap_text);
                }
            }
            b"indent" => {
                if let Ok(indent) = value.trim().parse::<u32>() {
                    if indent > 0 {
                        alignment.set_indent(indent);
                    }
                }
            }
            b"textRotation" => {
                if let Ok(rotation) = value.trim().parse::<u32>() {
                    alignment.set_text_rotation(rotation);
                }
            }
            b"shrinkToFit" => {
                if let Some(shrink) = parse_xml_bool(value.as_str()) {
                    alignment.set_shrink_to_fit(shrink);
                }
            }
            b"readingOrder" => {
                if let Ok(order) = value.trim().parse::<u32>() {
                    alignment.set_reading_order(order);
                }
            }
            _ => {}
        }
    }

    alignment.has_metadata().then_some(alignment)
}

fn parse_cell_protection(event: &BytesStart<'_>) -> Option<CellProtection> {
    let mut protection = CellProtection::new();

    for attribute in event.attributes().flatten() {
        let key = local_name(attribute.key.as_ref());
        let value = String::from_utf8_lossy(attribute.value.as_ref()).into_owned();
        match key {
            b"locked" => {
                if let Some(locked) = parse_xml_bool(value.as_str()) {
                    protection.set_locked(locked);
                }
            }
            b"hidden" => {
                if let Some(hidden) = parse_xml_bool(value.as_str()) {
                    protection.set_hidden(hidden);
                }
            }
            _ => {}
        }
    }

    protection.has_metadata().then_some(protection)
}

fn parse_xml_bool(raw: &str) -> Option<bool> {
    match raw.trim() {
        "1" | "true" => Some(true),
        "0" | "false" => Some(false),
        _ => None,
    }
}

fn parse_worksheet_xml(
    name: &str,
    xml: &[u8],
    shared_strings: Option<&[SharedStringEntry]>,
) -> Result<Worksheet> {
    let mut reader = Reader::from_reader(Cursor::new(xml));
    reader.config_mut().trim_text(false);
    let mut buffer = Vec::new();

    #[derive(Debug, Default)]
    struct CellState {
        reference: Option<String>,
        fallback_reference: Option<String>,
        cell_type: Option<String>,
        style_id: Option<u32>,
        formula: Option<String>,
        value: Option<String>,
        inline_text: Option<String>,
        /// Whether the `<f>` element has `t="array"`.
        is_array_formula: bool,
        /// The `ref` attribute on the `<f>` element (used by array and shared formulas).
        formula_ref: Option<String>,
        /// The `si` attribute on `<f t="shared">`.
        shared_formula_index: Option<u32>,
        /// Whether the `<f>` element has `t="shared"`.
        is_shared_formula: bool,
        unknown_attrs: Vec<(String, String)>,
        unknown_children: Vec<RawXmlNode>,
    }

    impl CellState {
        fn into_cell(self, shared_strings: Option<&[SharedStringEntry]>) -> Result<(String, Cell)> {
            let reference = self.reference.or(self.fallback_reference).ok_or_else(|| {
                XlsxError::UnsupportedPackage("cell missing `r` attribute".to_string())
            })?;

            let mut cell = Cell::new();

            if let Some(style_id) = self.style_id {
                cell.set_style_id(style_id);
            }

            cell.set_unknown_attrs(self.unknown_attrs);
            for unknown_child in self.unknown_children {
                cell.push_unknown_child(unknown_child);
            }

            let has_formula = self.formula.is_some();

            if let Some(formula) = self.formula {
                cell.set_formula(formula);
            }

            // Array formula attributes
            if self.is_array_formula {
                cell.set_array_formula(true);
                if let Some(ref_range) = &self.formula_ref {
                    cell.set_array_range(ref_range.clone());
                }
            }

            // Shared formula attributes — the master cell's `ref` range is
            // stored in `array_range` so serialization can write it back.
            // Use `set_formula_ref` to avoid setting `is_array_formula = true`.
            if let Some(si) = self.shared_formula_index {
                cell.set_shared_formula_index(si);
                if !self.is_array_formula {
                    if let Some(ref_range) = &self.formula_ref {
                        cell.set_formula_ref(ref_range.clone());
                    }
                }
            }

            match self.cell_type.as_deref() {
                Some("inlineStr") => {
                    if let Some(text) = self.inline_text {
                        cell.set_value(text);
                    } else {
                        cell.set_value(CellValue::Blank);
                    }
                }
                Some("b") => {
                    if let Some(raw) = self.value {
                        let bool_value = match raw.as_str() {
                            "1" | "true" => true,
                            "0" | "false" => false,
                            _ => {
                                tracing::warn!(
                                    value = raw.as_str(),
                                    "invalid bool cell value; preserving as string fallback"
                                );
                                cell.set_value(raw);
                                return Ok((reference, cell));
                            }
                        };
                        cell.set_value(bool_value);
                    }
                }
                Some("s") => {
                    let Some(raw) = self.value else {
                        tracing::warn!(
                            "shared string cell missing `<v>` value; using blank fallback"
                        );
                        cell.set_value(CellValue::Blank);
                        return Ok((reference, cell));
                    };

                    let Some(shared_strings) = shared_strings else {
                        tracing::warn!(
                            "shared string cell encountered without shared strings table; using blank fallback"
                        );
                        cell.set_value(CellValue::Blank);
                        return Ok((reference, cell));
                    };

                    let Ok(index) = raw.trim().parse::<usize>() else {
                        tracing::warn!(
                            value = raw.as_str(),
                            "invalid shared string index; using blank fallback"
                        );
                        cell.set_value(CellValue::Blank);
                        return Ok((reference, cell));
                    };

                    let Some(entry) = shared_strings.get(index) else {
                        tracing::warn!(
                            index,
                            "shared string index out of bounds; using blank fallback"
                        );
                        cell.set_value(CellValue::Blank);
                        return Ok((reference, cell));
                    };

                    match entry {
                        SharedStringEntry::Plain(value) => {
                            cell.set_value(value.clone());
                        }
                        SharedStringEntry::RichText { runs, .. } => {
                            cell.set_value(CellValue::RichText(runs.clone()));
                        }
                    }
                }
                Some("n") | None => {
                    if let Some(raw) = self.value {
                        if raw.is_empty() {
                            cell.set_value(CellValue::Blank);
                        } else {
                            match raw.parse::<f64>() {
                                Ok(number) => {
                                    cell.set_value(number);
                                }
                                Err(_) => {
                                    tracing::warn!(
                                        value = raw.as_str(),
                                        "invalid numeric cell value; preserving as string fallback"
                                    );
                                    cell.set_value(raw);
                                }
                            }
                        }
                    }
                }
                Some("str") => {
                    if let Some(raw) = self.value {
                        cell.set_value(raw);
                    }
                }
                Some("d") => {
                    if let Some(raw) = self.value {
                        cell.set_value(CellValue::date(raw));
                    }
                }
                Some("e") => {
                    if let Some(raw) = self.value {
                        cell.set_value(CellValue::error(raw));
                    }
                }
                Some(other) => {
                    tracing::warn!(
                        cell_type = other,
                        "unsupported cell type; preserving cell value as string/blank fallback"
                    );
                    if let Some(raw) = self.value {
                        cell.set_value(raw);
                    } else {
                        cell.set_value(CellValue::Blank);
                    }
                }
            }

            // When a cell has both a formula and a value, the value from `<v>`
            // is the cached formula result — move it to `cached_value` so the
            // cell's primary `value` stays `None` (formula cells derive their
            // display value from recalculation).
            if has_formula || self.is_shared_formula {
                if let Some(parsed_value) = cell.value().cloned() {
                    if !matches!(parsed_value, CellValue::Blank) {
                        cell.set_cached_value(parsed_value);
                        cell.clear_value();
                    }
                }
            }

            Ok((reference, cell))
        }
    }

    #[derive(Debug)]
    struct RowState {
        row: Row,
    }

    impl RowState {
        fn new(index: u32) -> Self {
            Self {
                row: Row::new(index),
            }
        }
    }

    #[derive(Debug, Default)]
    struct DataValidationState {
        validation_type: Option<DataValidationType>,
        sqref: Vec<CellRange>,
        formula1: Option<String>,
        formula2: Option<String>,
        error_style: Option<DataValidationErrorStyle>,
        error_title: Option<String>,
        error_message: Option<String>,
        prompt_title: Option<String>,
        prompt_message: Option<String>,
        show_input_message: Option<bool>,
        show_error_message: Option<bool>,
    }

    impl DataValidationState {
        fn into_data_validation(self) -> Option<DataValidation> {
            let validation_type = self.validation_type?;
            let mut dv = DataValidation::from_parsed_parts(
                validation_type,
                self.sqref,
                self.formula1.unwrap_or_default(),
                self.formula2,
            )
            .ok()?;
            if let Some(style) = self.error_style {
                dv.set_error_style(style);
            }
            if let Some(title) = self.error_title {
                dv.set_error_title(title);
            }
            if let Some(msg) = self.error_message {
                dv.set_error_message(msg);
            }
            if let Some(title) = self.prompt_title {
                dv.set_prompt_title(title);
            }
            if let Some(msg) = self.prompt_message {
                dv.set_prompt_message(msg);
            }
            if let Some(v) = self.show_input_message {
                dv.set_show_input_message(v);
            }
            if let Some(v) = self.show_error_message {
                dv.set_show_error_message(v);
            }
            Some(dv)
        }
    }

    #[derive(Debug, Default)]
    struct ConditionalFormattingState {
        sqref: Vec<CellRange>,
        raw_sqref: Option<String>,
    }

    #[derive(Debug, Default)]
    struct ConditionalFormattingRuleState {
        rule_type: Option<ConditionalFormattingRuleType>,
        formulas: Vec<String>,
        operator: Option<ConditionalFormattingOperator>,
        dxf_id: Option<u32>,
        priority: Option<u32>,
        stop_if_true: Option<bool>,
        text: Option<String>,
        time_period: Option<String>,
        rank: Option<u32>,
        percent: Option<bool>,
        bottom: Option<bool>,
        above_average: Option<bool>,
        equal_average: Option<bool>,
        std_dev: Option<u32>,
        color_scale_stops: Vec<ColorScaleStop>,
        data_bar_min: Option<CfValueObject>,
        data_bar_max: Option<CfValueObject>,
        data_bar_color: Option<String>,
        data_bar_show_value: Option<bool>,
        data_bar_min_length: Option<u32>,
        data_bar_max_length: Option<u32>,
        icon_set_name: Option<String>,
        icon_set_values: Vec<CfValueObject>,
        icon_set_show_value: Option<bool>,
        icon_set_reverse: Option<bool>,
    }

    impl ConditionalFormattingRuleState {
        fn into_conditional_formatting(
            self,
            sqref: &[CellRange],
            raw_sqref: Option<&str>,
        ) -> Option<ConditionalFormatting> {
            let rule_type = self.rule_type?;
            let mut cf = ConditionalFormatting::from_parsed_parts_raw(
                rule_type,
                sqref.to_vec(),
                raw_sqref.map(String::from),
                self.formulas,
            )
            .ok()?;
            if let Some(op) = self.operator {
                cf.set_operator(op);
            }
            if let Some(id) = self.dxf_id {
                cf.set_dxf_id(id);
            }
            if let Some(p) = self.priority {
                cf.set_priority(p);
            }
            if let Some(v) = self.stop_if_true {
                cf.set_stop_if_true(v);
            }
            if let Some(t) = self.text {
                cf.set_text(t);
            }
            if let Some(tp) = self.time_period {
                cf.set_time_period(tp);
            }
            if let Some(r) = self.rank {
                cf.set_rank(r);
            }
            if let Some(v) = self.percent {
                cf.set_cf_percent(v);
            }
            if let Some(v) = self.bottom {
                cf.set_cf_bottom(v);
            }
            if let Some(v) = self.above_average {
                cf.set_above_average(v);
            }
            if let Some(v) = self.equal_average {
                cf.set_equal_average(v);
            }
            if let Some(v) = self.std_dev {
                cf.set_std_dev(v);
            }
            if !self.color_scale_stops.is_empty() {
                cf.set_color_scale_stops(self.color_scale_stops);
            }
            if let Some(min) = self.data_bar_min {
                cf.set_data_bar_min(min);
            }
            if let Some(max) = self.data_bar_max {
                cf.set_data_bar_max(max);
            }
            if let Some(color) = self.data_bar_color {
                cf.set_data_bar_color(color);
            }
            if let Some(v) = self.data_bar_show_value {
                cf.set_data_bar_show_value(v);
            }
            if let Some(v) = self.data_bar_min_length {
                cf.set_data_bar_min_length(v);
            }
            if let Some(v) = self.data_bar_max_length {
                cf.set_data_bar_max_length(v);
            }
            if let Some(name) = self.icon_set_name {
                cf.set_icon_set_name(name);
            }
            if !self.icon_set_values.is_empty() {
                cf.set_icon_set_values(self.icon_set_values);
            }
            if let Some(v) = self.icon_set_show_value {
                cf.set_icon_set_show_value(v);
            }
            if let Some(v) = self.icon_set_reverse {
                cf.set_icon_set_reverse(v);
            }
            Some(cf)
        }

        fn parse_attributes(state: &mut ConditionalFormattingRuleState, event: &BytesStart<'_>) {
            for attribute in event.attributes().flatten() {
                let key = local_name(attribute.key.as_ref());
                let value = String::from_utf8_lossy(attribute.value.as_ref()).into_owned();
                match key {
                    b"type" => {
                        state.rule_type =
                            ConditionalFormattingRuleType::from_xml_type(value.trim());
                    }
                    b"operator" => {
                        state.operator =
                            ConditionalFormattingOperator::from_xml_value(value.trim());
                    }
                    b"dxfId" => {
                        state.dxf_id = value.trim().parse::<u32>().ok();
                    }
                    b"priority" => {
                        state.priority = value.trim().parse::<u32>().ok();
                    }
                    b"stopIfTrue" => {
                        state.stop_if_true = Some(value.trim() == "1" || value.trim() == "true");
                    }
                    b"text" => {
                        state.text = Some(value);
                    }
                    b"timePeriod" => {
                        state.time_period = Some(value);
                    }
                    b"rank" => {
                        state.rank = value.trim().parse::<u32>().ok();
                    }
                    b"percent" => {
                        state.percent = Some(value.trim() == "1" || value.trim() == "true");
                    }
                    b"bottom" => {
                        state.bottom = Some(value.trim() == "1" || value.trim() == "true");
                    }
                    b"aboveAverage" => {
                        // Default is true in the spec; only set to false when "0" or "false"
                        state.above_average = Some(value.trim() != "0" && value.trim() != "false");
                    }
                    b"equalAverage" => {
                        state.equal_average = Some(value.trim() == "1" || value.trim() == "true");
                    }
                    b"stdDev" => {
                        state.std_dev = value.trim().parse::<u32>().ok();
                    }
                    _ => {}
                }
            }
        }
    }

    /// Parses a `<color>` element and returns a color string (from rgb, theme, or indexed attrs).
    fn parse_cf_color_element(event: &BytesStart<'_>) -> String {
        let mut color = String::new();
        for attr in event.attributes().flatten() {
            let key = local_name(attr.key.as_ref());
            let v = String::from_utf8_lossy(attr.value.as_ref()).into_owned();
            match key {
                b"rgb" => {
                    color = v;
                    break;
                }
                b"theme" => {
                    if color.is_empty() {
                        color = format!("theme:{v}");
                    }
                }
                b"indexed" => {
                    if color.is_empty() {
                        color = format!("indexed:{v}");
                    }
                }
                _ => {}
            }
        }
        color
    }

    // Known top-level elements inside <worksheet> that the parser handles.
    const KNOWN_WORKSHEET_CHILDREN: &[&[u8]] = &[
        b"sheetData",
        b"sheetViews",
        b"sheetPr",
        b"tabColor",
        b"mergeCells",
        b"mergeCell",
        b"autoFilter",
        b"filterColumn",
        b"filters",
        b"filter",
        b"customFilters",
        b"customFilter",
        b"top10",
        b"dynamicFilter",
        b"colorFilter",
        b"iconFilter",
        b"conditionalFormatting",
        b"cfRule",
        b"colorScale",
        b"dataBar",
        b"iconSet",
        b"dataValidations",
        b"dataValidation",
        b"tableParts",
        b"tablePart",
        b"drawing",
        b"sheetFormatPr",
        b"cols",
        b"dimension",
        b"sheetProtection",
        b"pageMargins",
        b"pageSetup",
        b"headerFooter",
        b"printOptions",
        b"hyperlinks",
        b"hyperlink",
        b"rowBreaks",
        b"colBreaks",
    ];

    let mut worksheet = Worksheet::new(name.to_string());
    let mut current_cell: Option<CellState> = None;
    let mut current_row_state: Option<RowState> = None;
    let mut current_data_validation: Option<DataValidationState> = None;
    let mut current_conditional_formatting: Option<ConditionalFormattingState> = None;
    let mut current_conditional_formatting_rule: Option<ConditionalFormattingRuleState> = None;
    let mut in_formula = false;
    let mut in_value = false;
    let mut in_inline_text = false;
    let mut in_data_validation_formula1 = false;
    let mut in_data_validation_formula2 = false;
    let mut in_conditional_formatting_formula = false;
    let mut in_sheet_pr = false;
    let mut depth: u32 = 0;
    let mut current_row_index: Option<u32> = None;
    let mut next_row_index: u32 = 1;
    let mut next_column_index_in_row: u32 = 1;

    loop {
        let event = reader.read_event_into(&mut buffer)?;

        // Track nesting depth for unknown element detection.
        match &event {
            Event::Start(_) => depth += 1,
            Event::End(_) => depth = depth.saturating_sub(1),
            _ => {}
        }

        match event {
            // Capture extra namespace declarations from <worksheet> for dirty-save roundtrip.
            Event::Start(ref event)
                if depth == 1 && local_name(event.name().as_ref()) == b"worksheet" =>
            {
                let extra_ns = offidized_opc::xml_util::capture_extra_namespace_declarations(
                    event,
                    &["xmlns", "xmlns:r"],
                );
                if !extra_ns.is_empty() {
                    worksheet.set_extra_namespace_declarations(extra_ns);
                }
            }
            Event::Start(ref event) if local_name(event.name().as_ref()) == b"c" => {
                let mut state = CellState::default();
                for attribute in event.attributes().flatten() {
                    let key = local_name(attribute.key.as_ref());
                    let raw_key = String::from_utf8_lossy(attribute.key.as_ref()).into_owned();
                    let value = String::from_utf8_lossy(attribute.value.as_ref()).into_owned();
                    match key {
                        b"r" => state.reference = Some(value),
                        b"t" => state.cell_type = Some(value),
                        b"s" => state.style_id = value.parse::<u32>().ok(),
                        _ => state.unknown_attrs.push((raw_key, value)),
                    }
                }
                if state.reference.is_none() {
                    let row_index = current_row_index.unwrap_or(next_row_index);
                    let column_index = next_column_index_in_row.max(1);
                    state.fallback_reference = Some(build_cell_reference(column_index, row_index)?);
                    next_column_index_in_row = column_index.saturating_add(1);
                } else if let Some(reference) = state.reference.as_ref() {
                    if let Ok((column_index, row_index)) =
                        cell_reference_to_column_row(reference.as_str())
                    {
                        if current_row_index.is_none() {
                            current_row_index = Some(row_index);
                        }
                        next_row_index = next_row_index.max(row_index.saturating_add(1));
                        next_column_index_in_row = column_index.saturating_add(1);
                    }
                }
                current_cell = Some(state);
            }
            Event::Empty(ref event) if local_name(event.name().as_ref()) == b"c" => {
                let mut state = CellState::default();
                for attribute in event.attributes().flatten() {
                    let key = local_name(attribute.key.as_ref());
                    let raw_key = String::from_utf8_lossy(attribute.key.as_ref()).into_owned();
                    let value = String::from_utf8_lossy(attribute.value.as_ref()).into_owned();
                    match key {
                        b"r" => state.reference = Some(value),
                        b"t" => state.cell_type = Some(value),
                        b"s" => state.style_id = value.parse::<u32>().ok(),
                        _ => state.unknown_attrs.push((raw_key, value)),
                    }
                }
                if state.reference.is_none() {
                    let row_index = current_row_index.unwrap_or(next_row_index);
                    let column_index = next_column_index_in_row.max(1);
                    state.fallback_reference = Some(build_cell_reference(column_index, row_index)?);
                    next_column_index_in_row = column_index.saturating_add(1);
                } else if let Some(reference) = state.reference.as_ref() {
                    if let Ok((column_index, row_index)) =
                        cell_reference_to_column_row(reference.as_str())
                    {
                        if current_row_index.is_none() {
                            current_row_index = Some(row_index);
                        }
                        next_row_index = next_row_index.max(row_index.saturating_add(1));
                        next_column_index_in_row = column_index.saturating_add(1);
                    }
                }

                let (reference, cell) = state.into_cell(shared_strings)?;
                if let Err(error) = worksheet.insert_cell(reference.as_str(), cell) {
                    tracing::warn!(
                        reference = reference.as_str(),
                        error = %error,
                        "invalid worksheet cell reference; dropping parsed cell"
                    );
                }
            }
            Event::Start(ref event) if local_name(event.name().as_ref()) == b"row" => {
                let parsed_row_index = parse_xml_attr(event, b"r")
                    .and_then(|value| value.trim().parse::<u32>().ok())
                    .filter(|value| *value > 0);
                let row_index = parsed_row_index.unwrap_or(next_row_index);
                let mut row_state = RowState::new(row_index);
                let mut unknown_attrs = Vec::new();
                for attribute in event.attributes().flatten() {
                    let raw_key = String::from_utf8_lossy(attribute.key.as_ref()).into_owned();
                    let local = local_name(attribute.key.as_ref());
                    let value = String::from_utf8_lossy(attribute.value.as_ref()).into_owned();
                    match local {
                        b"r" => {}
                        b"customHeight" => {
                            if value.trim() == "1" || value.trim() == "true" {
                                row_state.row.set_custom_height(true);
                            }
                        }
                        b"ht" => {
                            if let Ok(height) = value.trim().parse::<f64>() {
                                row_state.row.set_height(height);
                            }
                        }
                        b"hidden" => {
                            if value.trim() == "1" || value.trim() == "true" {
                                row_state.row.set_hidden(true);
                            }
                        }
                        b"outlineLevel" => {
                            if let Ok(level) = value.trim().parse::<u8>() {
                                row_state.row.set_outline_level(level);
                            }
                        }
                        b"collapsed" => {
                            if value.trim() == "1" || value.trim() == "true" {
                                row_state.row.set_collapsed(true);
                            }
                        }
                        _ => unknown_attrs.push((raw_key, value)),
                    }
                }
                row_state.row.set_unknown_attrs(unknown_attrs);
                current_row_state = Some(row_state);
                current_row_index = Some(row_index);
                next_row_index = row_index.saturating_add(1);
                next_column_index_in_row = 1;
            }
            Event::Empty(ref event) if local_name(event.name().as_ref()) == b"row" => {
                let parsed_row_index = parse_xml_attr(event, b"r")
                    .and_then(|value| value.trim().parse::<u32>().ok())
                    .filter(|value| *value > 0);
                let row_index = parsed_row_index.unwrap_or(next_row_index);
                let mut row = Row::new(row_index);
                let mut unknown_attrs = Vec::new();
                for attribute in event.attributes().flatten() {
                    let raw_key = String::from_utf8_lossy(attribute.key.as_ref()).into_owned();
                    let local = local_name(attribute.key.as_ref());
                    let value = String::from_utf8_lossy(attribute.value.as_ref()).into_owned();
                    match local {
                        b"r" => {}
                        b"customHeight" => {
                            if value.trim() == "1" || value.trim() == "true" {
                                row.set_custom_height(true);
                            }
                        }
                        b"ht" => {
                            if let Ok(height) = value.trim().parse::<f64>() {
                                row.set_height(height);
                            }
                        }
                        b"hidden" => {
                            if value.trim() == "1" || value.trim() == "true" {
                                row.set_hidden(true);
                            }
                        }
                        b"outlineLevel" => {
                            if let Ok(level) = value.trim().parse::<u8>() {
                                row.set_outline_level(level);
                            }
                        }
                        b"collapsed" => {
                            if value.trim() == "1" || value.trim() == "true" {
                                row.set_collapsed(true);
                            }
                        }
                        _ => unknown_attrs.push((raw_key, value)),
                    }
                }
                row.set_unknown_attrs(unknown_attrs);
                worksheet.insert_row(row);
                next_row_index = row_index.saturating_add(1);
                current_row_index = None;
                next_column_index_in_row = 1;
            }
            Event::Empty(ref event) | Event::Start(ref event)
                if local_name(event.name().as_ref()) == b"mergeCell" =>
            {
                for attribute in event.attributes().flatten() {
                    if local_name(attribute.key.as_ref()) == b"ref" {
                        let range = String::from_utf8_lossy(attribute.value.as_ref()).into_owned();
                        if let Ok(range) = CellRange::parse(range.as_str()) {
                            worksheet.push_merged_range(range);
                        }
                    }
                }
            }
            Event::Empty(ref event) | Event::Start(ref event)
                if local_name(event.name().as_ref()) == b"col" =>
            {
                let mut min_col: Option<u32> = None;
                let mut max_col: Option<u32> = None;
                let mut col_width: Option<f64> = None;
                let mut col_hidden = false;
                let mut col_outline_level: u8 = 0;
                let mut col_collapsed = false;
                let mut col_style_index: Option<u32> = None;
                let mut col_best_fit = false;
                let mut col_custom_width = false;
                for attribute in event.attributes().flatten() {
                    let local = local_name(attribute.key.as_ref());
                    let value = String::from_utf8_lossy(attribute.value.as_ref()).into_owned();
                    match local {
                        b"min" => min_col = value.trim().parse::<u32>().ok(),
                        b"max" => max_col = value.trim().parse::<u32>().ok(),
                        b"width" => col_width = value.trim().parse::<f64>().ok(),
                        b"hidden" => {
                            col_hidden = value.trim() == "1" || value.trim() == "true";
                        }
                        b"outlineLevel" => {
                            col_outline_level = value.trim().parse::<u8>().unwrap_or(0);
                        }
                        b"collapsed" => {
                            col_collapsed = value.trim() == "1" || value.trim() == "true";
                        }
                        b"style" => {
                            col_style_index = value.trim().parse::<u32>().ok();
                        }
                        b"bestFit" => {
                            col_best_fit = value.trim() == "1" || value.trim() == "true";
                        }
                        b"customWidth" => {
                            col_custom_width = value.trim() == "1" || value.trim() == "true";
                        }
                        _ => {}
                    }
                }
                if let (Some(min), Some(max)) = (min_col, max_col) {
                    for col_index in min..=max {
                        if col_index == 0 {
                            continue;
                        }
                        let mut column = Column::new(col_index);
                        if let Some(width) = col_width {
                            column.set_width(width);
                        }
                        column.set_hidden(col_hidden);
                        if col_outline_level > 0 {
                            column.set_outline_level(col_outline_level);
                        }
                        column.set_collapsed(col_collapsed);
                        if let Some(style_index) = col_style_index {
                            column.set_style_index(style_index);
                        }
                        column.set_best_fit(col_best_fit);
                        column.set_custom_width(col_custom_width);
                        worksheet.insert_column(column);
                    }
                }
            }
            Event::Empty(ref event) | Event::Start(ref event)
                if local_name(event.name().as_ref()) == b"hyperlink" =>
            {
                let mut cell_ref: Option<String> = None;
                let mut r_id: Option<String> = None;
                let mut location: Option<String> = None;
                let mut tooltip: Option<String> = None;
                for attribute in event.attributes().flatten() {
                    let local = local_name(attribute.key.as_ref());
                    let value = String::from_utf8_lossy(attribute.value.as_ref()).into_owned();
                    match local {
                        b"ref" => cell_ref = Some(value),
                        b"id" => r_id = Some(value),
                        b"location" => location = Some(value),
                        b"tooltip" => tooltip = Some(value),
                        _ => {}
                    }
                }
                if let Some(cell_ref) = cell_ref {
                    // r:id references an external relationship containing the URL.
                    // location is an internal reference (e.g., "Sheet2!A1").
                    // We store the r:id string temporarily; the caller resolves it
                    // to the actual URL using the worksheet relationships.
                    // For now, store r:id as the url placeholder so the save path
                    // can resolve it back.
                    let url = r_id;
                    if let Ok(hyperlink) =
                        Hyperlink::from_parsed_parts(cell_ref, url, location, tooltip, None)
                    {
                        worksheet.push_hyperlink(hyperlink);
                    }
                }
            }
            Event::Empty(ref event) | Event::Start(ref event)
                if local_name(event.name().as_ref()) == b"sheetView" =>
            {
                let options = parse_sheet_view_options(event);
                if options.has_metadata() {
                    worksheet.set_parsed_sheet_view_options(options);
                }
            }
            Event::Empty(ref event) | Event::Start(ref event)
                if local_name(event.name().as_ref()) == b"pane" =>
            {
                if let Some(freeze_pane) = parse_freeze_pane(event) {
                    worksheet.set_parsed_freeze_pane(freeze_pane);
                }
            }
            Event::Empty(ref event) | Event::Start(ref event)
                if local_name(event.name().as_ref()) == b"autoFilter" =>
            {
                let mut af = AutoFilter::new();
                for attribute in event.attributes().flatten() {
                    if local_name(attribute.key.as_ref()) == b"ref" {
                        let range = String::from_utf8_lossy(attribute.value.as_ref()).into_owned();
                        let _ = af.set_range(range.as_str());
                    }
                }
                worksheet.set_parsed_auto_filter(af);
            }
            Event::Start(ref event) | Event::Empty(ref event)
                if local_name(event.name().as_ref()) == b"filterColumn" =>
            {
                if let Some(af) = worksheet.auto_filter_internal_mut() {
                    let mut col_id: Option<u32> = None;
                    let mut show_button: Option<bool> = None;
                    for attribute in event.attributes().flatten() {
                        let local = local_name(attribute.key.as_ref());
                        let value = String::from_utf8_lossy(attribute.value.as_ref()).into_owned();
                        match local {
                            b"colId" => col_id = value.trim().parse::<u32>().ok(),
                            b"showButton" => {
                                show_button = Some(value.trim() != "0");
                            }
                            b"hiddenButton" => {
                                // hiddenButton="1" means button is hidden (inverse of showButton)
                                show_button = Some(value.trim() == "0");
                            }
                            _ => {}
                        }
                    }
                    if let Some(col_id) = col_id {
                        let fc = af.filter_column_mut(col_id);
                        if let Some(show_button) = show_button {
                            fc.set_show_button(show_button);
                        }
                    }
                }
            }
            Event::Empty(ref event) | Event::Start(ref event)
                if local_name(event.name().as_ref()) == b"filters" =>
            {
                if let Some(af) = worksheet.auto_filter_internal_mut() {
                    if let Some(fc) = af.filter_columns_mut().last_mut() {
                        for attribute in event.attributes().flatten() {
                            if local_name(attribute.key.as_ref()) == b"blank" {
                                let value =
                                    String::from_utf8_lossy(attribute.value.as_ref()).into_owned();
                                if value.trim() == "1" || value.eq_ignore_ascii_case("true") {
                                    fc.set_blank(true);
                                }
                            }
                        }
                    }
                }
            }
            Event::Empty(ref event) | Event::Start(ref event)
                if local_name(event.name().as_ref()) == b"filter" =>
            {
                if let Some(af) = worksheet.auto_filter_internal_mut() {
                    for attribute in event.attributes().flatten() {
                        if local_name(attribute.key.as_ref()) == b"val" {
                            let value =
                                String::from_utf8_lossy(attribute.value.as_ref()).into_owned();
                            if let Some(fc) = af.filter_columns_mut().last_mut() {
                                fc.add_value(value);
                            }
                        }
                    }
                }
            }
            Event::Empty(ref event) | Event::Start(ref event)
                if local_name(event.name().as_ref()) == b"customFilters" =>
            {
                if let Some(af) = worksheet.auto_filter_internal_mut() {
                    if let Some(fc) = af.filter_columns_mut().last_mut() {
                        for attribute in event.attributes().flatten() {
                            if local_name(attribute.key.as_ref()) == b"and" {
                                let value =
                                    String::from_utf8_lossy(attribute.value.as_ref()).into_owned();
                                fc.set_custom_filters_and(
                                    value.trim() == "1" || value.trim() == "true",
                                );
                            }
                        }
                    }
                }
            }
            Event::Empty(ref event) | Event::Start(ref event)
                if local_name(event.name().as_ref()) == b"customFilter" =>
            {
                if let Some(af) = worksheet.auto_filter_internal_mut() {
                    if let Some(fc) = af.filter_columns_mut().last_mut() {
                        let mut operator = CustomFilterOperator::Equal;
                        let mut val = String::new();
                        for attribute in event.attributes().flatten() {
                            let local = local_name(attribute.key.as_ref());
                            let value =
                                String::from_utf8_lossy(attribute.value.as_ref()).into_owned();
                            match local {
                                b"operator" => {
                                    operator = CustomFilterOperator::from_xml_value(value.trim())
                                }
                                b"val" => val = value,
                                _ => {}
                            }
                        }
                        fc.add_custom_filter(CustomFilter::new(operator, val));
                    }
                }
            }
            Event::Empty(ref event) | Event::Start(ref event)
                if local_name(event.name().as_ref()) == b"top10" =>
            {
                if let Some(af) = worksheet.auto_filter_internal_mut() {
                    if let Some(fc) = af.filter_columns_mut().last_mut() {
                        for attribute in event.attributes().flatten() {
                            let local = local_name(attribute.key.as_ref());
                            let value =
                                String::from_utf8_lossy(attribute.value.as_ref()).into_owned();
                            match local {
                                b"top" => {
                                    fc.set_top(value.trim() != "0");
                                }
                                b"percent" => {
                                    fc.set_percent(value.trim() == "1" || value.trim() == "true");
                                }
                                b"val" => {
                                    if let Ok(v) = value.trim().parse::<f64>() {
                                        fc.set_top10_val(v);
                                    }
                                }
                                _ => {}
                            }
                        }
                    }
                }
            }
            Event::Empty(ref event) | Event::Start(ref event)
                if local_name(event.name().as_ref()) == b"dynamicFilter" =>
            {
                if let Some(af) = worksheet.auto_filter_internal_mut() {
                    if let Some(fc) = af.filter_columns_mut().last_mut() {
                        for attribute in event.attributes().flatten() {
                            if local_name(attribute.key.as_ref()) == b"type" {
                                let value =
                                    String::from_utf8_lossy(attribute.value.as_ref()).into_owned();
                                fc.set_dynamic_type(value);
                            }
                        }
                    }
                }
            }
            Event::Empty(ref event) | Event::Start(ref event)
                if local_name(event.name().as_ref()) == b"colorFilter" =>
            {
                if let Some(af) = worksheet.auto_filter_internal_mut() {
                    if let Some(fc) = af.filter_columns_mut().last_mut() {
                        // Always mark as Color filter when this element is present
                        fc.set_filter_type(FilterType::Color);
                        for attribute in event.attributes().flatten() {
                            let local = local_name(attribute.key.as_ref());
                            let value =
                                String::from_utf8_lossy(attribute.value.as_ref()).into_owned();
                            match local {
                                b"dxfId" => {
                                    if let Ok(id) = value.trim().parse::<u32>() {
                                        fc.set_dxf_id(id);
                                    }
                                }
                                b"cellColor" => {
                                    fc.set_cell_color(value.trim() != "0");
                                }
                                _ => {}
                            }
                        }
                    }
                }
            }
            Event::Empty(ref event) | Event::Start(ref event)
                if local_name(event.name().as_ref()) == b"iconFilter" =>
            {
                if let Some(af) = worksheet.auto_filter_internal_mut() {
                    if let Some(fc) = af.filter_columns_mut().last_mut() {
                        for attribute in event.attributes().flatten() {
                            let local = local_name(attribute.key.as_ref());
                            let value =
                                String::from_utf8_lossy(attribute.value.as_ref()).into_owned();
                            match local {
                                b"iconSet" => {
                                    fc.set_icon_set(value);
                                }
                                b"iconId" => {
                                    if let Ok(id) = value.trim().parse::<u32>() {
                                        fc.set_icon_id(id);
                                    }
                                }
                                _ => {}
                            }
                        }
                    }
                }
            }
            // Parse <sheetPr> (container for tabColor and other properties)
            Event::Start(ref event) if local_name(event.name().as_ref()) == b"sheetPr" => {
                in_sheet_pr = true;
            }
            Event::End(ref event) if local_name(event.name().as_ref()) == b"sheetPr" => {
                in_sheet_pr = false;
            }
            // Parse <tabColor> inside <sheetPr>
            Event::Empty(ref event) | Event::Start(ref event)
                if in_sheet_pr && local_name(event.name().as_ref()) == b"tabColor" =>
            {
                if let Some(rgb) = parse_xml_attr(event, b"rgb") {
                    // Strip leading "FF" alpha channel if present (ARGB -> RGB)
                    let color = if rgb.len() == 8 && rgb.starts_with("FF") {
                        rgb[2..].to_string()
                    } else {
                        rgb
                    };
                    worksheet.set_parsed_tab_color(color);
                }
            }
            // Parse <sheetFormatPr>
            Event::Empty(ref event) | Event::Start(ref event)
                if local_name(event.name().as_ref()) == b"sheetFormatPr" =>
            {
                let mut raw_attrs = Vec::new();
                for attribute in event.attributes().flatten() {
                    let key = local_name(attribute.key.as_ref());
                    let raw_key = String::from_utf8_lossy(attribute.key.as_ref()).into_owned();
                    let value = String::from_utf8_lossy(attribute.value.as_ref()).into_owned();
                    match key {
                        b"defaultRowHeight" => {
                            if let Ok(height) = value.trim().parse::<f64>() {
                                worksheet.set_parsed_default_row_height(height);
                            }
                        }
                        b"defaultColWidth" => {
                            if let Ok(width) = value.trim().parse::<f64>() {
                                worksheet.set_parsed_default_column_width(width);
                            }
                        }
                        b"customHeight" => {
                            if let Some(v) = parse_xml_bool(value.as_str()) {
                                worksheet.set_parsed_custom_height(v);
                            }
                        }
                        _ => {}
                    }
                    raw_attrs.push((raw_key, value));
                }
                worksheet.set_raw_sheet_format_pr_attrs(raw_attrs);
            }
            Event::Empty(ref event) | Event::Start(ref event)
                if local_name(event.name().as_ref()) == b"sheetProtection" =>
            {
                let protection = parse_sheet_protection(event);
                if protection.has_metadata() {
                    worksheet.set_parsed_protection(protection);
                }
            }
            Event::Empty(ref event) | Event::Start(ref event)
                if local_name(event.name().as_ref()) == b"pageSetup" =>
            {
                let page_setup = parse_page_setup(event);
                if page_setup.has_metadata() {
                    worksheet.set_parsed_page_setup(page_setup);
                }
            }
            Event::Empty(ref event) | Event::Start(ref event)
                if local_name(event.name().as_ref()) == b"pageMargins" =>
            {
                let margins = parse_page_margins(event);
                if margins.has_metadata() {
                    worksheet.set_parsed_page_margins(margins);
                }
            }
            Event::Empty(ref event) | Event::Start(ref event)
                if local_name(event.name().as_ref()) == b"printOptions" =>
            {
                let attrs: Vec<(String, String)> = event
                    .attributes()
                    .flatten()
                    .map(|a| {
                        (
                            String::from_utf8_lossy(a.key.as_ref()).into_owned(),
                            String::from_utf8_lossy(a.value.as_ref()).into_owned(),
                        )
                    })
                    .collect();
                if !attrs.is_empty() {
                    worksheet.set_raw_print_options_attrs(attrs);
                }
            }
            Event::Empty(ref event) | Event::Start(ref event)
                if local_name(event.name().as_ref()) == b"dimension" =>
            {
                if let Some(dim_ref) = parse_xml_attr(event, b"ref") {
                    worksheet.set_raw_dimension_ref(dim_ref);
                }
            }
            Event::Start(ref event)
                if local_name(event.name().as_ref()) == b"conditionalFormatting" =>
            {
                let mut state = ConditionalFormattingState::default();
                for attribute in event.attributes().flatten() {
                    if local_name(attribute.key.as_ref()) == b"sqref" {
                        let value = String::from_utf8_lossy(attribute.value.as_ref()).into_owned();
                        state.sqref = parse_sqref_ranges(value.as_str());
                        // Preserve raw sqref for ranges CellRange can't represent (e.g. "A:A")
                        if state.sqref.is_empty() {
                            state.raw_sqref = Some(value);
                        }
                    }
                }
                current_conditional_formatting = Some(state);
            }
            Event::Start(ref event) if local_name(event.name().as_ref()) == b"cfRule" => {
                let mut state = ConditionalFormattingRuleState::default();
                ConditionalFormattingRuleState::parse_attributes(&mut state, event);
                current_conditional_formatting_rule = Some(state);
            }
            Event::Empty(ref event) if local_name(event.name().as_ref()) == b"cfRule" => {
                let mut state = ConditionalFormattingRuleState::default();
                ConditionalFormattingRuleState::parse_attributes(&mut state, event);

                if let Some(cf_state) = current_conditional_formatting.as_ref() {
                    let raw = cf_state.raw_sqref.as_deref();
                    if let Some(rule) =
                        state.into_conditional_formatting(cf_state.sqref.as_slice(), raw)
                    {
                        worksheet.push_conditional_formatting(rule);
                    }
                }
            }
            // Parse <colorScale> child of <cfRule>
            Event::Start(ref event)
                if local_name(event.name().as_ref()) == b"colorScale"
                    && current_conditional_formatting_rule.is_some() =>
            {
                // Read all children within <colorScale>: <cfvo> and <color> elements
                let mut cfvos: Vec<CfValueObject> = Vec::new();
                let mut colors: Vec<String> = Vec::new();
                let mut cs_depth = 0_u32;
                loop {
                    let cs_event = reader.read_event_into(&mut buffer)?;
                    match &cs_event {
                        Event::Start(_) => cs_depth += 1,
                        Event::End(_) => {
                            if cs_depth == 0 {
                                break;
                            }
                            cs_depth = cs_depth.saturating_sub(1);
                        }
                        _ => {}
                    }
                    match cs_event {
                        Event::Empty(ref e) | Event::Start(ref e) => {
                            let name_bytes = e.name();
                            let local = local_name(name_bytes.as_ref());
                            if local == b"cfvo" {
                                let mut vtype: Option<CfValueObjectType> = None;
                                let mut val: Option<String> = None;
                                for attr in e.attributes().flatten() {
                                    let key = local_name(attr.key.as_ref());
                                    let v =
                                        String::from_utf8_lossy(attr.value.as_ref()).into_owned();
                                    match key {
                                        b"type" => {
                                            vtype = CfValueObjectType::from_xml_value(v.trim())
                                        }
                                        b"val" => val = Some(v),
                                        _ => {}
                                    }
                                }
                                if let Some(vt) = vtype {
                                    cfvos.push(CfValueObject {
                                        value_type: vt,
                                        value: val,
                                    });
                                }
                            } else if local == b"color" {
                                let color = parse_cf_color_element(e);
                                colors.push(color);
                            }
                        }
                        Event::Eof => break,
                        _ => {}
                    }
                    buffer.clear();
                }
                // Pair cfvo and color into ColorScaleStops
                if let Some(state) = current_conditional_formatting_rule.as_mut() {
                    for (cfvo, color) in cfvos.into_iter().zip(colors.into_iter()) {
                        state.color_scale_stops.push(ColorScaleStop { cfvo, color });
                    }
                }
                depth = depth.saturating_sub(1);
            }
            // Parse <dataBar> child of <cfRule>
            Event::Start(ref event)
                if local_name(event.name().as_ref()) == b"dataBar"
                    && current_conditional_formatting_rule.is_some() =>
            {
                let mut cfvos: Vec<CfValueObject> = Vec::new();
                let mut bar_color: Option<String> = None;
                let mut db_show_value: Option<bool> = None;
                let mut db_min_length: Option<u32> = None;
                let mut db_max_length: Option<u32> = None;
                // Read attributes on the <dataBar> element itself
                for attr in event.attributes().flatten() {
                    let key = local_name(attr.key.as_ref());
                    let v = String::from_utf8_lossy(attr.value.as_ref()).into_owned();
                    match key {
                        b"showValue" => db_show_value = Some(v.trim() != "0"),
                        b"minLength" => db_min_length = v.trim().parse::<u32>().ok(),
                        b"maxLength" => db_max_length = v.trim().parse::<u32>().ok(),
                        _ => {}
                    }
                }
                let mut db_depth = 0_u32;
                loop {
                    let db_event = reader.read_event_into(&mut buffer)?;
                    match &db_event {
                        Event::Start(_) => db_depth += 1,
                        Event::End(_) => {
                            if db_depth == 0 {
                                break;
                            }
                            db_depth = db_depth.saturating_sub(1);
                        }
                        _ => {}
                    }
                    match db_event {
                        Event::Empty(ref e) | Event::Start(ref e) => {
                            let name_bytes = e.name();
                            let local = local_name(name_bytes.as_ref());
                            if local == b"cfvo" {
                                let mut vtype: Option<CfValueObjectType> = None;
                                let mut val: Option<String> = None;
                                for attr in e.attributes().flatten() {
                                    let key = local_name(attr.key.as_ref());
                                    let v =
                                        String::from_utf8_lossy(attr.value.as_ref()).into_owned();
                                    match key {
                                        b"type" => {
                                            vtype = CfValueObjectType::from_xml_value(v.trim())
                                        }
                                        b"val" => val = Some(v),
                                        _ => {}
                                    }
                                }
                                if let Some(vt) = vtype {
                                    cfvos.push(CfValueObject {
                                        value_type: vt,
                                        value: val,
                                    });
                                }
                            } else if local == b"color" {
                                bar_color = Some(parse_cf_color_element(e));
                            }
                        }
                        Event::Eof => break,
                        _ => {}
                    }
                    buffer.clear();
                }
                if let Some(state) = current_conditional_formatting_rule.as_mut() {
                    if cfvos.len() >= 2 {
                        state.data_bar_min = Some(cfvos.remove(0));
                        state.data_bar_max = Some(cfvos.remove(0));
                    } else if let Some(cfvo) = cfvos.into_iter().next() {
                        state.data_bar_min = Some(cfvo);
                    }
                    state.data_bar_color = bar_color;
                    state.data_bar_show_value = db_show_value;
                    state.data_bar_min_length = db_min_length;
                    state.data_bar_max_length = db_max_length;
                }
                depth = depth.saturating_sub(1);
            }
            // Parse <iconSet> child of <cfRule>
            Event::Start(ref event)
                if local_name(event.name().as_ref()) == b"iconSet"
                    && current_conditional_formatting_rule.is_some() =>
            {
                // Read attributes on the <iconSet> element itself
                let mut is_name: Option<String> = None;
                let mut is_show_value: Option<bool> = None;
                let mut is_reverse: Option<bool> = None;
                for attr in event.attributes().flatten() {
                    let key = local_name(attr.key.as_ref());
                    let v = String::from_utf8_lossy(attr.value.as_ref()).into_owned();
                    match key {
                        b"iconSet" => is_name = Some(v),
                        b"showValue" => {
                            is_show_value = Some(v.trim() != "0" && v.trim() != "false");
                        }
                        b"reverse" => {
                            is_reverse = Some(v.trim() == "1" || v.trim() == "true");
                        }
                        _ => {}
                    }
                }
                let mut cfvos: Vec<CfValueObject> = Vec::new();
                let mut is_depth = 0_u32;
                loop {
                    let is_event = reader.read_event_into(&mut buffer)?;
                    match &is_event {
                        Event::Start(_) => is_depth += 1,
                        Event::End(_) => {
                            if is_depth == 0 {
                                break;
                            }
                            is_depth = is_depth.saturating_sub(1);
                        }
                        _ => {}
                    }
                    match is_event {
                        Event::Empty(ref e) | Event::Start(ref e) => {
                            let name_bytes = e.name();
                            let local = local_name(name_bytes.as_ref());
                            if local == b"cfvo" {
                                let mut vtype: Option<CfValueObjectType> = None;
                                let mut val: Option<String> = None;
                                for attr in e.attributes().flatten() {
                                    let key = local_name(attr.key.as_ref());
                                    let v =
                                        String::from_utf8_lossy(attr.value.as_ref()).into_owned();
                                    match key {
                                        b"type" => {
                                            vtype = CfValueObjectType::from_xml_value(v.trim())
                                        }
                                        b"val" => val = Some(v),
                                        _ => {}
                                    }
                                }
                                if let Some(vt) = vtype {
                                    cfvos.push(CfValueObject {
                                        value_type: vt,
                                        value: val,
                                    });
                                }
                            }
                        }
                        Event::Eof => break,
                        _ => {}
                    }
                    buffer.clear();
                }
                if let Some(state) = current_conditional_formatting_rule.as_mut() {
                    state.icon_set_name = is_name;
                    state.icon_set_values = cfvos;
                    state.icon_set_show_value = is_show_value;
                    state.icon_set_reverse = is_reverse;
                }
                depth = depth.saturating_sub(1);
            }
            Event::Start(ref event) if local_name(event.name().as_ref()) == b"headerFooter" => {
                let mut hf = PrintHeaderFooter::new();
                for attribute in event.attributes().flatten() {
                    let local = local_name(attribute.key.as_ref());
                    let value = String::from_utf8_lossy(attribute.value.as_ref()).into_owned();
                    match local {
                        b"differentOddEven" => {
                            hf.set_different_odd_even(
                                value.trim() == "1" || value.trim() == "true",
                            );
                        }
                        b"differentFirst" => {
                            hf.set_different_first(value.trim() == "1" || value.trim() == "true");
                        }
                        _ => {}
                    }
                }
                // Read child elements for header/footer text
                let mut current_hf_element: Option<String> = None;
                let mut hf_depth = 0_u32;
                loop {
                    let hf_event = reader.read_event_into(&mut buffer)?;
                    match &hf_event {
                        Event::Start(_) => hf_depth += 1,
                        Event::End(_) => {
                            if hf_depth == 0 {
                                // End of headerFooter
                                break;
                            }
                            hf_depth = hf_depth.saturating_sub(1);
                        }
                        _ => {}
                    }
                    match hf_event {
                        Event::Start(ref e) => {
                            let name_bytes = e.name();
                            let local = local_name(name_bytes.as_ref());
                            match local {
                                b"oddHeader" | b"oddFooter" | b"evenHeader" | b"evenFooter"
                                | b"firstHeader" | b"firstFooter" => {
                                    current_hf_element =
                                        Some(String::from_utf8_lossy(local).into_owned());
                                }
                                _ => {}
                            }
                        }
                        Event::Text(ref e) => {
                            if let Some(ref elem_name) = current_hf_element {
                                let text = e.xml_content().unwrap_or_default().into_owned();
                                match elem_name.as_str() {
                                    "oddHeader" => {
                                        hf.set_odd_header(text);
                                    }
                                    "oddFooter" => {
                                        hf.set_odd_footer(text);
                                    }
                                    "evenHeader" => {
                                        hf.set_even_header(text);
                                    }
                                    "evenFooter" => {
                                        hf.set_even_footer(text);
                                    }
                                    "firstHeader" => {
                                        hf.set_first_header(text);
                                    }
                                    "firstFooter" => {
                                        hf.set_first_footer(text);
                                    }
                                    _ => {}
                                }
                            }
                        }
                        Event::End(ref e) => {
                            let name_bytes = e.name();
                            let local = local_name(name_bytes.as_ref());
                            if matches!(
                                local,
                                b"oddHeader"
                                    | b"oddFooter"
                                    | b"evenHeader"
                                    | b"evenFooter"
                                    | b"firstHeader"
                                    | b"firstFooter"
                            ) {
                                current_hf_element = None;
                            }
                        }
                        Event::Eof => break,
                        _ => {}
                    }
                    buffer.clear();
                }
                worksheet.set_parsed_header_footer(hf);
                // Consumed the headerFooter End event above, so adjust depth
                depth = depth.saturating_sub(1);
            }
            Event::Start(ref event)
                if local_name(event.name().as_ref()) == b"rowBreaks"
                    || local_name(event.name().as_ref()) == b"colBreaks" =>
            {
                let is_row_breaks = local_name(event.name().as_ref()) == b"rowBreaks";
                let mut breaks = worksheet
                    .page_breaks()
                    .cloned()
                    .unwrap_or_else(PageBreaks::new);
                let mut brk_depth = 0_u32;
                loop {
                    let brk_event = reader.read_event_into(&mut buffer)?;
                    match &brk_event {
                        Event::Start(_) => brk_depth += 1,
                        Event::End(_) => {
                            if brk_depth == 0 {
                                break;
                            }
                            brk_depth = brk_depth.saturating_sub(1);
                        }
                        _ => {}
                    }
                    match brk_event {
                        Event::Empty(ref e) | Event::Start(ref e)
                            if local_name(e.name().as_ref()) == b"brk" =>
                        {
                            let mut id: Option<u32> = None;
                            let mut manual = false;
                            for attribute in e.attributes().flatten() {
                                let local = local_name(attribute.key.as_ref());
                                let value =
                                    String::from_utf8_lossy(attribute.value.as_ref()).into_owned();
                                match local {
                                    b"id" => id = value.trim().parse::<u32>().ok(),
                                    b"man" => {
                                        manual = value.trim() == "1" || value.trim() == "true"
                                    }
                                    _ => {}
                                }
                            }
                            if let Some(id) = id {
                                let pb = PageBreak::with_manual(id, manual);
                                if is_row_breaks {
                                    breaks.row_breaks_mut().push(pb);
                                } else {
                                    breaks.col_breaks_mut().push(pb);
                                }
                            }
                        }
                        Event::Eof => break,
                        _ => {}
                    }
                    buffer.clear();
                }
                worksheet.set_parsed_page_breaks(breaks);
                depth = depth.saturating_sub(1);
            }
            Event::Start(ref event) if local_name(event.name().as_ref()) == b"dataValidation" => {
                let mut state = DataValidationState::default();

                for attribute in event.attributes().flatten() {
                    let key = local_name(attribute.key.as_ref());
                    let value = String::from_utf8_lossy(attribute.value.as_ref()).into_owned();
                    match key {
                        b"type" => {
                            state.validation_type = DataValidationType::from_xml_type(value.trim());
                        }
                        b"sqref" => {
                            state.sqref = parse_sqref_ranges(value.as_str());
                        }
                        b"errorStyle" => {
                            state.error_style =
                                DataValidationErrorStyle::from_xml_value(value.trim());
                        }
                        b"errorTitle" => {
                            state.error_title = Some(value);
                        }
                        b"error" => {
                            state.error_message = Some(value);
                        }
                        b"promptTitle" => {
                            state.prompt_title = Some(value);
                        }
                        b"prompt" => {
                            state.prompt_message = Some(value);
                        }
                        b"showInputMessage" => {
                            state.show_input_message = parse_xml_bool(value.as_str());
                        }
                        b"showErrorMessage" => {
                            state.show_error_message = parse_xml_bool(value.as_str());
                        }
                        _ => {}
                    }
                }

                current_data_validation = Some(state);
            }
            Event::Empty(ref event) if local_name(event.name().as_ref()) == b"dataValidation" => {
                let mut state = DataValidationState::default();

                for attribute in event.attributes().flatten() {
                    let key = local_name(attribute.key.as_ref());
                    let value = String::from_utf8_lossy(attribute.value.as_ref()).into_owned();
                    match key {
                        b"type" => {
                            state.validation_type = DataValidationType::from_xml_type(value.trim());
                        }
                        b"sqref" => {
                            state.sqref = parse_sqref_ranges(value.as_str());
                        }
                        b"errorStyle" => {
                            state.error_style =
                                DataValidationErrorStyle::from_xml_value(value.trim());
                        }
                        b"errorTitle" => {
                            state.error_title = Some(value);
                        }
                        b"error" => {
                            state.error_message = Some(value);
                        }
                        b"promptTitle" => {
                            state.prompt_title = Some(value);
                        }
                        b"prompt" => {
                            state.prompt_message = Some(value);
                        }
                        b"showInputMessage" => {
                            state.show_input_message = parse_xml_bool(value.as_str());
                        }
                        b"showErrorMessage" => {
                            state.show_error_message = parse_xml_bool(value.as_str());
                        }
                        _ => {}
                    }
                }

                if let Some(data_validation) = state.into_data_validation() {
                    worksheet.push_data_validation(data_validation);
                }
            }
            Event::Start(ref event) if local_name(event.name().as_ref()) == b"f" => {
                in_formula = true;
                // Parse <f> attributes: t="array"|"shared", ref="...", si="..."
                if let Some(state) = current_cell.as_mut() {
                    for attribute in event.attributes().flatten() {
                        let key = local_name(attribute.key.as_ref());
                        let value = String::from_utf8_lossy(attribute.value.as_ref()).into_owned();
                        match key {
                            b"t" if value == "array" => {
                                state.is_array_formula = true;
                            }
                            b"t" if value == "shared" => {
                                state.is_shared_formula = true;
                            }
                            b"ref" => {
                                state.formula_ref = Some(value);
                            }
                            b"si" => {
                                state.shared_formula_index = value.parse::<u32>().ok();
                            }
                            _ => {}
                        }
                    }
                }
            }
            Event::Start(ref event) if local_name(event.name().as_ref()) == b"v" => {
                in_value = true;
            }
            Event::Start(ref event)
                if local_name(event.name().as_ref()) == b"t"
                    && current_cell
                        .as_ref()
                        .and_then(|state| state.cell_type.as_ref())
                        .is_some_and(|cell_type| cell_type == "inlineStr") =>
            {
                in_inline_text = true;
            }
            Event::Start(ref event)
                if local_name(event.name().as_ref()) == b"formula1"
                    && current_data_validation.is_some() =>
            {
                in_data_validation_formula1 = true;
            }
            Event::Start(ref event)
                if local_name(event.name().as_ref()) == b"formula2"
                    && current_data_validation.is_some() =>
            {
                in_data_validation_formula2 = true;
            }
            Event::Start(ref event)
                if local_name(event.name().as_ref()) == b"formula"
                    && current_conditional_formatting_rule.is_some() =>
            {
                in_conditional_formatting_formula = true;
                if let Some(state) = current_conditional_formatting_rule.as_mut() {
                    state.formulas.push(String::new());
                }
            }
            Event::Start(ref event) if current_cell.is_some() => {
                let name_bytes = event.name();
                let local = local_name(name_bytes.as_ref());
                // Keep known cell children handled by typed parsing; preserve everything else.
                if !matches!(local, b"f" | b"v" | b"is" | b"t" | b"r" | b"rPr") {
                    if let Some(state) = current_cell.as_mut() {
                        state
                            .unknown_children
                            .push(RawXmlNode::read_element(&mut reader, event)?);
                        depth = depth.saturating_sub(1);
                    }
                }
            }
            Event::Start(ref event) if current_row_state.is_some() && current_cell.is_none() => {
                let name_bytes = event.name();
                let local = local_name(name_bytes.as_ref());
                if local != b"c" {
                    if let Some(state) = current_row_state.as_mut() {
                        state
                            .row
                            .push_unknown_child(RawXmlNode::read_element(&mut reader, event)?);
                        depth = depth.saturating_sub(1);
                    }
                }
            }
            // Handle self-closing <f t="shared" si="0"/> elements (shared formula dependents).
            Event::Empty(ref event)
                if current_cell.is_some() && local_name(event.name().as_ref()) == b"f" =>
            {
                if let Some(state) = current_cell.as_mut() {
                    for attribute in event.attributes().flatten() {
                        let key = local_name(attribute.key.as_ref());
                        let value = String::from_utf8_lossy(attribute.value.as_ref()).into_owned();
                        match key {
                            b"t" if value == "array" => {
                                state.is_array_formula = true;
                            }
                            b"t" if value == "shared" => {
                                state.is_shared_formula = true;
                            }
                            b"ref" => {
                                state.formula_ref = Some(value);
                            }
                            b"si" => {
                                state.shared_formula_index = value.parse::<u32>().ok();
                            }
                            _ => {}
                        }
                    }
                }
            }
            Event::Empty(ref event) if current_cell.is_some() => {
                let name_bytes = event.name();
                let local = local_name(name_bytes.as_ref());
                if !matches!(local, b"f" | b"v" | b"is" | b"t" | b"r" | b"rPr") {
                    if let Some(state) = current_cell.as_mut() {
                        state
                            .unknown_children
                            .push(RawXmlNode::from_empty_element(event));
                    }
                }
            }
            Event::Empty(ref event) if current_row_state.is_some() && current_cell.is_none() => {
                let name_bytes = event.name();
                let local = local_name(name_bytes.as_ref());
                if local != b"c" {
                    if let Some(state) = current_row_state.as_mut() {
                        state
                            .row
                            .push_unknown_child(RawXmlNode::from_empty_element(event));
                    }
                }
            }
            Event::Text(ref event) => {
                let text = event
                    .xml_content()
                    .map_err(quick_xml::Error::from)?
                    .into_owned();
                if in_formula {
                    if let Some(state) = current_cell.as_mut() {
                        state
                            .formula
                            .get_or_insert_with(String::new)
                            .push_str(text.as_str());
                    }
                } else if in_value {
                    if let Some(state) = current_cell.as_mut() {
                        state
                            .value
                            .get_or_insert_with(String::new)
                            .push_str(text.as_str());
                    }
                } else if in_inline_text {
                    if let Some(state) = current_cell.as_mut() {
                        state
                            .inline_text
                            .get_or_insert_with(String::new)
                            .push_str(text.as_str());
                    }
                } else if in_data_validation_formula1 {
                    if let Some(state) = current_data_validation.as_mut() {
                        state
                            .formula1
                            .get_or_insert_with(String::new)
                            .push_str(text.as_str());
                    }
                } else if in_data_validation_formula2 {
                    if let Some(state) = current_data_validation.as_mut() {
                        state
                            .formula2
                            .get_or_insert_with(String::new)
                            .push_str(text.as_str());
                    }
                } else if in_conditional_formatting_formula {
                    if let Some(state) = current_conditional_formatting_rule.as_mut() {
                        if let Some(formula) = state.formulas.last_mut() {
                            formula.push_str(text.as_str());
                        }
                    }
                }
            }
            Event::CData(ref event) => {
                let text = String::from_utf8_lossy(event.as_ref()).into_owned();
                if in_formula {
                    if let Some(state) = current_cell.as_mut() {
                        state
                            .formula
                            .get_or_insert_with(String::new)
                            .push_str(text.as_str());
                    }
                } else if in_value {
                    if let Some(state) = current_cell.as_mut() {
                        state
                            .value
                            .get_or_insert_with(String::new)
                            .push_str(text.as_str());
                    }
                } else if in_inline_text {
                    if let Some(state) = current_cell.as_mut() {
                        state
                            .inline_text
                            .get_or_insert_with(String::new)
                            .push_str(text.as_str());
                    }
                } else if in_data_validation_formula1 {
                    if let Some(state) = current_data_validation.as_mut() {
                        state
                            .formula1
                            .get_or_insert_with(String::new)
                            .push_str(text.as_str());
                    }
                } else if in_data_validation_formula2 {
                    if let Some(state) = current_data_validation.as_mut() {
                        state
                            .formula2
                            .get_or_insert_with(String::new)
                            .push_str(text.as_str());
                    }
                } else if in_conditional_formatting_formula {
                    if let Some(state) = current_conditional_formatting_rule.as_mut() {
                        if let Some(formula) = state.formulas.last_mut() {
                            formula.push_str(text.as_str());
                        }
                    }
                }
            }
            Event::End(ref event) if local_name(event.name().as_ref()) == b"f" => {
                in_formula = false;
            }
            Event::End(ref event) if local_name(event.name().as_ref()) == b"v" => {
                in_value = false;
            }
            Event::End(ref event) if local_name(event.name().as_ref()) == b"t" => {
                in_inline_text = false;
            }
            Event::End(ref event) if local_name(event.name().as_ref()) == b"formula1" => {
                in_data_validation_formula1 = false;
            }
            Event::End(ref event) if local_name(event.name().as_ref()) == b"formula2" => {
                in_data_validation_formula2 = false;
            }
            Event::End(ref event) if local_name(event.name().as_ref()) == b"formula" => {
                in_conditional_formatting_formula = false;
            }
            Event::End(ref event) if local_name(event.name().as_ref()) == b"cfRule" => {
                in_conditional_formatting_formula = false;
                if let (Some(state), Some(cf_state)) = (
                    current_conditional_formatting_rule.take(),
                    current_conditional_formatting.as_ref(),
                ) {
                    let raw = cf_state.raw_sqref.as_deref();
                    if let Some(rule) =
                        state.into_conditional_formatting(cf_state.sqref.as_slice(), raw)
                    {
                        worksheet.push_conditional_formatting(rule);
                    }
                }
            }
            Event::End(ref event)
                if local_name(event.name().as_ref()) == b"conditionalFormatting" =>
            {
                current_conditional_formatting_rule = None;
                current_conditional_formatting = None;
                in_conditional_formatting_formula = false;
            }
            Event::End(ref event) if local_name(event.name().as_ref()) == b"dataValidation" => {
                in_data_validation_formula1 = false;
                in_data_validation_formula2 = false;
                if let Some(state) = current_data_validation.take() {
                    if let Some(data_validation) = state.into_data_validation() {
                        worksheet.push_data_validation(data_validation);
                    }
                }
            }
            Event::End(ref event) if local_name(event.name().as_ref()) == b"c" => {
                if let Some(state) = current_cell.take() {
                    let (reference, cell) = state.into_cell(shared_strings)?;
                    if let Err(error) = worksheet.insert_cell(reference.as_str(), cell) {
                        tracing::warn!(
                            reference = reference.as_str(),
                            error = %error,
                            "invalid worksheet cell reference; dropping parsed cell"
                        );
                    }
                }
                in_formula = false;
                in_value = false;
                in_inline_text = false;
                in_conditional_formatting_formula = false;
            }
            Event::End(ref event) if local_name(event.name().as_ref()) == b"row" => {
                if let Some(row_state) = current_row_state.take() {
                    worksheet.insert_row(row_state.row);
                }
                current_row_index = None;
                next_column_index_in_row = 1;
            }
            Event::Eof => break,
            // Capture unknown elements at the worksheet level for roundtrip fidelity.
            // depth == 2 for Start means the element is a direct child of <worksheet> (depth 1).
            Event::Start(ref e) if depth == 2 => {
                let name_bytes = e.name();
                let local = local_name(name_bytes.as_ref());
                if !KNOWN_WORKSHEET_CHILDREN.contains(&local) {
                    let node = RawXmlNode::read_element(&mut reader, e)?;
                    worksheet.push_unknown_child(node);
                    // read_element consumed through the matching End event,
                    // so restore depth to worksheet level.
                    depth = 1;
                }
            }
            // depth == 1 for Empty means the element is a direct child of <worksheet>.
            Event::Empty(ref e) if depth == 1 => {
                let name_bytes = e.name();
                let local = local_name(name_bytes.as_ref());
                if !KNOWN_WORKSHEET_CHILDREN.contains(&local) {
                    worksheet.push_unknown_child(RawXmlNode::from_empty_element(e));
                }
            }
            _ => {}
        }
        buffer.clear();
    }

    Ok(worksheet)
}

fn parse_worksheet_relationship_ids(xml: &[u8]) -> Result<WorksheetRelationshipIds> {
    let mut reader = Reader::from_reader(Cursor::new(xml));
    reader.config_mut().trim_text(true);
    let mut buffer = Vec::new();
    let mut relationship_ids = WorksheetRelationshipIds::default();

    loop {
        match reader.read_event_into(&mut buffer)? {
            Event::Empty(ref event) | Event::Start(ref event)
                if local_name(event.name().as_ref()) == b"tablePart" =>
            {
                for attribute in event.attributes().flatten() {
                    if local_name(attribute.key.as_ref()) == b"id" {
                        let value = String::from_utf8_lossy(attribute.value.as_ref()).into_owned();
                        if !value.is_empty()
                            && !relationship_ids.table_relationship_ids.contains(&value)
                        {
                            relationship_ids.table_relationship_ids.push(value);
                        }
                    }
                }
            }
            Event::Empty(ref event) | Event::Start(ref event)
                if local_name(event.name().as_ref()) == b"drawing" =>
            {
                for attribute in event.attributes().flatten() {
                    if local_name(attribute.key.as_ref()) == b"id" {
                        let value = String::from_utf8_lossy(attribute.value.as_ref()).into_owned();
                        if !value.is_empty()
                            && !relationship_ids.drawing_relationship_ids.contains(&value)
                        {
                            relationship_ids.drawing_relationship_ids.push(value);
                        }
                    }
                }
            }
            Event::Empty(ref event) | Event::Start(ref event)
                if local_name(event.name().as_ref()) == b"hyperlink" =>
            {
                for attribute in event.attributes().flatten() {
                    if local_name(attribute.key.as_ref()) == b"id" {
                        let value = String::from_utf8_lossy(attribute.value.as_ref()).into_owned();
                        if !value.is_empty()
                            && !relationship_ids.hyperlink_relationship_ids.contains(&value)
                        {
                            relationship_ids.hyperlink_relationship_ids.push(value);
                        }
                    }
                }
            }
            Event::Eof => break,
            _ => {}
        }
        buffer.clear();
    }

    Ok(relationship_ids)
}

fn load_worksheet_tables(
    worksheet: &mut Worksheet,
    package: &Package,
    worksheet_uri: &PartUri,
    worksheet_part: &Part,
    table_relationship_ids: &[String],
) -> Result<()> {
    for relationship_id in table_relationship_ids {
        let Some(relationship) = worksheet_part
            .relationships
            .get_by_id(relationship_id.as_str())
        else {
            continue;
        };
        if relationship.target_mode != TargetMode::Internal
            || relationship.rel_type != TABLE_RELATIONSHIP_TYPE
        {
            continue;
        }

        let table_uri = worksheet_uri.resolve_relative(relationship.target.as_str())?;
        let Some(table_part) = package.get_part(table_uri.as_str()) else {
            continue;
        };

        if let Some(table) = parse_table_xml(table_part.data.as_bytes())? {
            worksheet.push_table(table);
        }
    }

    Ok(())
}

fn load_worksheet_images(
    worksheet: &mut Worksheet,
    package: &Package,
    worksheet_uri: &PartUri,
    worksheet_part: &Part,
    drawing_relationship_ids: &[String],
) -> Result<()> {
    for drawing_relationship_id in drawing_relationship_ids {
        let Some(drawing_relationship) = worksheet_part
            .relationships
            .get_by_id(drawing_relationship_id.as_str())
        else {
            continue;
        };

        if drawing_relationship.target_mode != TargetMode::Internal
            || drawing_relationship.rel_type != DRAWING_RELATIONSHIP_TYPE
        {
            continue;
        }

        let drawing_uri = worksheet_uri.resolve_relative(drawing_relationship.target.as_str())?;
        let Some(drawing_part) = package.get_part(drawing_uri.as_str()) else {
            continue;
        };

        for drawing_image in parse_drawing_xml(drawing_part.data.as_bytes())? {
            let Some(image_relationship) = drawing_part
                .relationships
                .get_by_id(drawing_image.relationship_id.as_str())
            else {
                continue;
            };

            if image_relationship.target_mode != TargetMode::Internal
                || image_relationship.rel_type != RelationshipType::IMAGE
            {
                continue;
            }

            let image_uri = drawing_uri.resolve_relative(image_relationship.target.as_str())?;
            let Some(image_part) = package.get_part(image_uri.as_str()) else {
                continue;
            };

            let content_type = image_part
                .content_type
                .as_deref()
                .or_else(|| package.content_types().get(image_uri.as_str()))
                .map(str::to_string);

            let Some(content_type) = content_type else {
                continue;
            };

            let image = WorksheetImage::from_parsed_parts(
                image_part.data.as_bytes().to_vec(),
                content_type,
                drawing_image.anchor_cell,
                drawing_image.ext,
            );
            if let Ok(mut image) = image {
                image.set_anchor_type(drawing_image.anchor_type);
                if let Some(from) = drawing_image.from_anchor {
                    image.set_from_anchor(from);
                }
                if let Some(to) = drawing_image.to_anchor {
                    image.set_to_anchor(to);
                }
                if let Some(cx) = drawing_image.extent_cx {
                    image.set_extent_cx(cx);
                }
                if let Some(cy) = drawing_image.extent_cy {
                    image.set_extent_cy(cy);
                }
                if let Some(x) = drawing_image.position_x {
                    image.set_position_x(x);
                }
                if let Some(y) = drawing_image.position_y {
                    image.set_position_y(y);
                }
                if let Some(v) = drawing_image.crop_left {
                    image.set_crop_left(v);
                }
                if let Some(v) = drawing_image.crop_right {
                    image.set_crop_right(v);
                }
                if let Some(v) = drawing_image.crop_top {
                    image.set_crop_top(v);
                }
                if let Some(v) = drawing_image.crop_bottom {
                    image.set_crop_bottom(v);
                }
                worksheet.push_image(image);
            }
        }
    }

    Ok(())
}

/// Resolves hyperlink r:id references to actual URLs using the worksheet part relationships.
/// Hyperlinks parsed from XML store the r:id value in the `url` field. This function
/// replaces those r:id references with the actual target URL from the relationships.
fn resolve_worksheet_hyperlink_urls(worksheet: &mut Worksheet, sheet_part: &Part) {
    // We need to rebuild the hyperlinks with resolved URLs.
    let old_hyperlinks: Vec<Hyperlink> = worksheet.hyperlinks().to_vec();
    worksheet.clear_hyperlinks();
    // Undo the dirty flag from clear since this is part of loading.
    for hyperlink in old_hyperlinks {
        let resolved_url = hyperlink.url().and_then(|r_id| {
            sheet_part
                .relationships
                .get_by_id(r_id)
                .map(|rel| rel.target.clone())
        });
        // If the url was a relationship ID that resolved to a target, use that.
        // Otherwise, keep whatever was stored (for internal-only hyperlinks, url is None).
        let url = if resolved_url.is_some() {
            resolved_url
        } else {
            // If it didn't resolve, it wasn't a relationship ID; keep as-is.
            hyperlink.url().map(str::to_string)
        };
        if let Ok(resolved) = Hyperlink::from_parsed_parts(
            hyperlink.cell_ref().to_string(),
            url,
            hyperlink.location().map(str::to_string),
            hyperlink.tooltip().map(str::to_string),
            hyperlink.display().map(str::to_string),
        ) {
            worksheet.push_hyperlink(resolved);
        }
    }
}

fn parse_drawing_xml(xml: &[u8]) -> Result<Vec<ParsedDrawingImageRef>> {
    let mut reader = Reader::from_reader(Cursor::new(xml));
    reader.config_mut().trim_text(true);
    let mut buffer = Vec::new();
    let mut images = Vec::new();
    let mut current_anchor: Option<ParsedDrawingImageRef> = None;
    let mut from_col: Option<u32> = None;
    let mut from_row: Option<u32> = None;
    let mut from_col_off: Option<i64> = None;
    let mut from_row_off: Option<i64> = None;
    let mut to_col: Option<u32> = None;
    let mut to_row: Option<u32> = None;
    let mut to_col_off: Option<i64> = None;
    let mut to_row_off: Option<i64> = None;
    let mut in_from = false;
    let mut in_to = false;
    let mut in_col = false;
    let mut in_row = false;
    let mut in_col_off = false;
    let mut in_row_off = false;

    fn new_parsed_ref(anchor_type: ImageAnchorType) -> ParsedDrawingImageRef {
        ParsedDrawingImageRef {
            relationship_id: String::new(),
            anchor_cell: String::new(),
            ext: None,
            anchor_type,
            from_anchor: None,
            to_anchor: None,
            extent_cx: None,
            extent_cy: None,
            position_x: None,
            position_y: None,
            crop_left: None,
            crop_right: None,
            crop_top: None,
            crop_bottom: None,
        }
    }

    #[allow(clippy::too_many_arguments)]
    fn reset_state(
        from_col: &mut Option<u32>,
        from_row: &mut Option<u32>,
        from_col_off: &mut Option<i64>,
        from_row_off: &mut Option<i64>,
        to_col: &mut Option<u32>,
        to_row: &mut Option<u32>,
        to_col_off: &mut Option<i64>,
        to_row_off: &mut Option<i64>,
        in_from: &mut bool,
        in_to: &mut bool,
        in_col: &mut bool,
        in_row: &mut bool,
        in_col_off: &mut bool,
        in_row_off: &mut bool,
    ) {
        *from_col = None;
        *from_row = None;
        *from_col_off = None;
        *from_row_off = None;
        *to_col = None;
        *to_row = None;
        *to_col_off = None;
        *to_row_off = None;
        *in_from = false;
        *in_to = false;
        *in_col = false;
        *in_row = false;
        *in_col_off = false;
        *in_row_off = false;
    }

    loop {
        match reader.read_event_into(&mut buffer)? {
            Event::Start(ref event) => {
                let name_bytes = event.name();
                let local = local_name(name_bytes.as_ref());
                match local {
                    b"oneCellAnchor" => {
                        current_anchor = Some(new_parsed_ref(ImageAnchorType::OneCell));
                        reset_state(
                            &mut from_col,
                            &mut from_row,
                            &mut from_col_off,
                            &mut from_row_off,
                            &mut to_col,
                            &mut to_row,
                            &mut to_col_off,
                            &mut to_row_off,
                            &mut in_from,
                            &mut in_to,
                            &mut in_col,
                            &mut in_row,
                            &mut in_col_off,
                            &mut in_row_off,
                        );
                    }
                    b"twoCellAnchor" => {
                        current_anchor = Some(new_parsed_ref(ImageAnchorType::TwoCell));
                        reset_state(
                            &mut from_col,
                            &mut from_row,
                            &mut from_col_off,
                            &mut from_row_off,
                            &mut to_col,
                            &mut to_row,
                            &mut to_col_off,
                            &mut to_row_off,
                            &mut in_from,
                            &mut in_to,
                            &mut in_col,
                            &mut in_row,
                            &mut in_col_off,
                            &mut in_row_off,
                        );
                    }
                    b"absoluteAnchor" => {
                        current_anchor = Some(new_parsed_ref(ImageAnchorType::Absolute));
                        reset_state(
                            &mut from_col,
                            &mut from_row,
                            &mut from_col_off,
                            &mut from_row_off,
                            &mut to_col,
                            &mut to_row,
                            &mut to_col_off,
                            &mut to_row_off,
                            &mut in_from,
                            &mut in_to,
                            &mut in_col,
                            &mut in_row,
                            &mut in_col_off,
                            &mut in_row_off,
                        );
                    }
                    b"from" if current_anchor.is_some() => {
                        in_from = true;
                    }
                    b"to" if current_anchor.is_some() => {
                        in_to = true;
                    }
                    b"col" if current_anchor.is_some() && (in_from || in_to) => {
                        in_col = true;
                    }
                    b"colOff" if current_anchor.is_some() && (in_from || in_to) => {
                        in_col_off = true;
                    }
                    b"row" if current_anchor.is_some() && (in_from || in_to) => {
                        in_row = true;
                    }
                    b"rowOff" if current_anchor.is_some() && (in_from || in_to) => {
                        in_row_off = true;
                    }
                    b"ext" if current_anchor.is_some() => {
                        if let Some(anchor) = current_anchor.as_mut() {
                            parse_drawing_ext_attributes(event, anchor);
                        }
                    }
                    b"pos" if current_anchor.is_some() => {
                        if let Some(anchor) = current_anchor.as_mut() {
                            parse_drawing_pos_attributes(event, anchor);
                        }
                    }
                    b"blip" if current_anchor.is_some() => {
                        if let Some(anchor) = current_anchor.as_mut() {
                            for attribute in event.attributes().flatten() {
                                if local_name(attribute.key.as_ref()) == b"embed" {
                                    anchor.relationship_id =
                                        String::from_utf8_lossy(attribute.value.as_ref())
                                            .into_owned();
                                }
                            }
                        }
                    }
                    b"srcRect" if current_anchor.is_some() => {
                        if let Some(anchor) = current_anchor.as_mut() {
                            parse_src_rect_attributes(event, anchor);
                        }
                    }
                    _ => {}
                }
            }
            Event::Empty(ref event) => {
                let name_bytes = event.name();
                let local = local_name(name_bytes.as_ref());
                match local {
                    b"ext" if current_anchor.is_some() => {
                        if let Some(anchor) = current_anchor.as_mut() {
                            parse_drawing_ext_attributes(event, anchor);
                        }
                    }
                    b"pos" if current_anchor.is_some() => {
                        if let Some(anchor) = current_anchor.as_mut() {
                            parse_drawing_pos_attributes(event, anchor);
                        }
                    }
                    b"blip" if current_anchor.is_some() => {
                        if let Some(anchor) = current_anchor.as_mut() {
                            for attribute in event.attributes().flatten() {
                                if local_name(attribute.key.as_ref()) == b"embed" {
                                    anchor.relationship_id =
                                        String::from_utf8_lossy(attribute.value.as_ref())
                                            .into_owned();
                                }
                            }
                        }
                    }
                    b"srcRect" if current_anchor.is_some() => {
                        if let Some(anchor) = current_anchor.as_mut() {
                            parse_src_rect_attributes(event, anchor);
                        }
                    }
                    _ => {}
                }
            }
            Event::Text(ref event) => {
                if current_anchor.is_some() {
                    let text = event
                        .xml_content()
                        .map_err(quick_xml::Error::from)?
                        .into_owned();
                    if in_col {
                        let val = text.trim().parse::<u32>().ok();
                        if in_from {
                            from_col = val;
                        } else if in_to {
                            to_col = val;
                        }
                    } else if in_row {
                        let val = text.trim().parse::<u32>().ok();
                        if in_from {
                            from_row = val;
                        } else if in_to {
                            to_row = val;
                        }
                    } else if in_col_off {
                        let val = text.trim().parse::<i64>().ok();
                        if in_from {
                            from_col_off = val;
                        } else if in_to {
                            to_col_off = val;
                        }
                    } else if in_row_off {
                        let val = text.trim().parse::<i64>().ok();
                        if in_from {
                            from_row_off = val;
                        } else if in_to {
                            to_row_off = val;
                        }
                    }
                }
            }
            Event::End(ref event) => {
                let name_bytes = event.name();
                let local = local_name(name_bytes.as_ref());
                match local {
                    b"col" => in_col = false,
                    b"colOff" => in_col_off = false,
                    b"row" => in_row = false,
                    b"rowOff" => in_row_off = false,
                    b"from" => {
                        if let Some(anchor) = current_anchor.as_mut() {
                            if let (Some(col), Some(row)) = (from_col, from_row) {
                                anchor.from_anchor = Some(CellAnchor::new(
                                    col,
                                    row,
                                    from_col_off.unwrap_or(0),
                                    from_row_off.unwrap_or(0),
                                ));
                            }
                        }
                        in_from = false;
                        in_col = false;
                        in_row = false;
                        in_col_off = false;
                        in_row_off = false;
                    }
                    b"to" => {
                        if let Some(anchor) = current_anchor.as_mut() {
                            if let (Some(col), Some(row)) = (to_col, to_row) {
                                anchor.to_anchor = Some(CellAnchor::new(
                                    col,
                                    row,
                                    to_col_off.unwrap_or(0),
                                    to_row_off.unwrap_or(0),
                                ));
                            }
                        }
                        in_to = false;
                        in_col = false;
                        in_row = false;
                        in_col_off = false;
                        in_row_off = false;
                    }
                    b"oneCellAnchor" | b"twoCellAnchor" | b"absoluteAnchor" => {
                        if let Some(mut anchor) = current_anchor.take() {
                            if anchor.relationship_id.is_empty() {
                                buffer.clear();
                                continue;
                            }
                            if let Some(ref from) = anchor.from_anchor {
                                let anchor_cell = build_cell_reference(
                                    from.col().checked_add(1).ok_or_else(|| {
                                        XlsxError::UnsupportedPackage(
                                            "drawing anchor column overflow".to_string(),
                                        )
                                    })?,
                                    from.row().checked_add(1).ok_or_else(|| {
                                        XlsxError::UnsupportedPackage(
                                            "drawing anchor row overflow".to_string(),
                                        )
                                    })?,
                                )?;
                                anchor.anchor_cell = anchor_cell;
                            } else if anchor.anchor_type == ImageAnchorType::Absolute {
                                anchor.anchor_cell = "A1".to_string();
                            } else {
                                buffer.clear();
                                continue;
                            }
                            images.push(anchor);
                        }
                        reset_state(
                            &mut from_col,
                            &mut from_row,
                            &mut from_col_off,
                            &mut from_row_off,
                            &mut to_col,
                            &mut to_row,
                            &mut to_col_off,
                            &mut to_row_off,
                            &mut in_from,
                            &mut in_to,
                            &mut in_col,
                            &mut in_row,
                            &mut in_col_off,
                            &mut in_row_off,
                        );
                    }
                    _ => {}
                }
            }
            Event::Eof => break,
            _ => {}
        }
        buffer.clear();
    }

    Ok(images)
}

/// Parses extent attributes (`cx`, `cy`) from a drawing `<xdr:ext>` element.
fn parse_drawing_ext_attributes(event: &BytesStart<'_>, anchor: &mut ParsedDrawingImageRef) {
    let mut cx = None;
    let mut cy = None;
    for attribute in event.attributes().flatten() {
        let key = local_name(attribute.key.as_ref());
        let value = String::from_utf8_lossy(attribute.value.as_ref()).into_owned();
        match key {
            b"cx" => cx = value.trim().parse::<u64>().ok(),
            b"cy" => cy = value.trim().parse::<u64>().ok(),
            _ => {}
        }
    }
    if let (Some(cx), Some(cy)) = (cx, cy) {
        anchor.ext = WorksheetImageExt::new(cx, cy).ok();
        anchor.extent_cx = Some(cx as i64);
        anchor.extent_cy = Some(cy as i64);
    }
}

/// Parses position attributes (`x`, `y`) from a drawing `<xdr:pos>` element.
fn parse_drawing_pos_attributes(event: &BytesStart<'_>, anchor: &mut ParsedDrawingImageRef) {
    for attribute in event.attributes().flatten() {
        let key = local_name(attribute.key.as_ref());
        let value = String::from_utf8_lossy(attribute.value.as_ref()).into_owned();
        match key {
            b"x" => anchor.position_x = value.trim().parse::<i64>().ok(),
            b"y" => anchor.position_y = value.trim().parse::<i64>().ok(),
            _ => {}
        }
    }
}

/// Parses `<a:srcRect>` crop attributes. Values in OOXML are thousandths of a percent.
fn parse_src_rect_attributes(event: &BytesStart<'_>, anchor: &mut ParsedDrawingImageRef) {
    for attribute in event.attributes().flatten() {
        let key = local_name(attribute.key.as_ref());
        let value = String::from_utf8_lossy(attribute.value.as_ref()).into_owned();
        match key {
            b"l" => {
                if let Ok(v) = value.trim().parse::<f64>() {
                    anchor.crop_left = Some(v / 100_000.0);
                }
            }
            b"r" => {
                if let Ok(v) = value.trim().parse::<f64>() {
                    anchor.crop_right = Some(v / 100_000.0);
                }
            }
            b"t" => {
                if let Ok(v) = value.trim().parse::<f64>() {
                    anchor.crop_top = Some(v / 100_000.0);
                }
            }
            b"b" => {
                if let Ok(v) = value.trim().parse::<f64>() {
                    anchor.crop_bottom = Some(v / 100_000.0);
                }
            }
            _ => {}
        }
    }
}

fn parse_table_xml(xml: &[u8]) -> Result<Option<WorksheetTable>> {
    let mut reader = Reader::from_reader(Cursor::new(xml));
    reader.config_mut().trim_text(true);
    let mut buffer = Vec::new();

    let mut table_name: Option<String> = None;
    let mut table_range: Option<CellRange> = None;
    let mut has_header_row = true;
    let mut totals_row_shown: Option<bool> = None;
    let mut style_name: Option<String> = None;
    let mut show_first_column: Option<bool> = None;
    let mut show_last_column: Option<bool> = None;
    let mut show_row_stripes: Option<bool> = None;
    let mut show_column_stripes: Option<bool> = None;
    let mut columns: Vec<TableColumn> = Vec::new();
    let mut current_column: Option<TableColumn> = None;
    let mut in_totals_row_formula = false;
    let mut totals_row_formula_text = String::new();
    let mut unknown_table_attrs: Vec<(String, String)> = Vec::new();
    let mut unknown_style_attrs: Vec<(String, String)> = Vec::new();

    loop {
        match reader.read_event_into(&mut buffer)? {
            Event::Start(ref event) | Event::Empty(ref event)
                if local_name(event.name().as_ref()) == b"table" =>
            {
                let mut name = None;
                let mut display_name = None;
                for attribute in event.attributes().flatten() {
                    let key = local_name(attribute.key.as_ref());
                    let raw_key = String::from_utf8_lossy(attribute.key.as_ref()).into_owned();
                    let value = String::from_utf8_lossy(attribute.value.as_ref()).into_owned();
                    match key {
                        b"name" => {
                            if !value.trim().is_empty() {
                                name = Some(value.clone());
                            }
                        }
                        b"displayName" => {
                            if !value.trim().is_empty() {
                                display_name = Some(value.clone());
                            }
                        }
                        b"ref" => {
                            if let Ok(range) = CellRange::parse(value.as_str()) {
                                table_range = Some(range);
                            }
                        }
                        b"headerRowCount" => {
                            if let Ok(header_row_count) = value.trim().parse::<u32>() {
                                has_header_row = header_row_count > 0;
                            }
                        }
                        b"totalsRowShown" => {
                            totals_row_shown = parse_xml_bool(value.as_str());
                        }
                        _ => {
                            unknown_table_attrs.push((raw_key, value));
                            continue;
                        }
                    }
                }
                table_name = display_name.or(name);
            }
            Event::Start(ref event) | Event::Empty(ref event)
                if local_name(event.name().as_ref()) == b"tableStyleInfo" =>
            {
                for attribute in event.attributes().flatten() {
                    let key = local_name(attribute.key.as_ref());
                    let raw_key = String::from_utf8_lossy(attribute.key.as_ref()).into_owned();
                    let value = String::from_utf8_lossy(attribute.value.as_ref()).into_owned();
                    match key {
                        b"name" => {
                            if !value.trim().is_empty() {
                                style_name = Some(value);
                            }
                        }
                        b"showFirstColumn" => {
                            show_first_column = parse_xml_bool(value.as_str());
                        }
                        b"showLastColumn" => {
                            show_last_column = parse_xml_bool(value.as_str());
                        }
                        b"showRowStripes" => {
                            show_row_stripes = parse_xml_bool(value.as_str());
                        }
                        b"showColumnStripes" => {
                            show_column_stripes = parse_xml_bool(value.as_str());
                        }
                        _ => {
                            unknown_style_attrs.push((raw_key, value));
                        }
                    }
                }
            }
            Event::Empty(ref event) if local_name(event.name().as_ref()) == b"tableColumn" => {
                let tc = parse_table_column_attributes(event);
                columns.push(tc);
            }
            Event::Start(ref event) if local_name(event.name().as_ref()) == b"tableColumn" => {
                let tc = parse_table_column_attributes(event);
                current_column = Some(tc);
            }
            Event::Start(ref event) if local_name(event.name().as_ref()) == b"totalsRowFormula" => {
                if current_column.is_some() {
                    in_totals_row_formula = true;
                    totals_row_formula_text.clear();
                }
            }
            Event::End(ref event) if local_name(event.name().as_ref()) == b"totalsRowFormula" => {
                if let Some(col) = current_column.as_mut() {
                    if in_totals_row_formula && !totals_row_formula_text.trim().is_empty() {
                        col.set_totals_row_formula(totals_row_formula_text.trim().to_string());
                    }
                }
                in_totals_row_formula = false;
            }
            Event::End(ref event) if local_name(event.name().as_ref()) == b"tableColumn" => {
                if let Some(col) = current_column.take() {
                    columns.push(col);
                }
            }
            Event::Text(ref event) if in_totals_row_formula => {
                if let Ok(text) = event.xml_content() {
                    totals_row_formula_text.push_str(text.as_ref());
                }
            }
            Event::Eof => break,
            _ => {}
        }
        buffer.clear();
    }

    let Some(name) = table_name else {
        return Ok(None);
    };
    let Some(range) = table_range else {
        return Ok(None);
    };

    let mut table = match WorksheetTable::from_parsed_parts(name, range, has_header_row) {
        Ok(table) => table,
        Err(_) => return Ok(None),
    };

    if let Some(v) = totals_row_shown {
        table.set_totals_row_shown(v);
    }
    if let Some(v) = style_name {
        table.set_style_name(v);
    }
    if let Some(v) = show_first_column {
        table.set_show_first_column(v);
    }
    if let Some(v) = show_last_column {
        table.set_show_last_column(v);
    }
    if let Some(v) = show_row_stripes {
        table.set_show_row_stripes(v);
    }
    if let Some(v) = show_column_stripes {
        table.set_show_column_stripes(v);
    }
    for col in columns {
        table.push_column(col);
    }
    if !unknown_table_attrs.is_empty() {
        table.set_unknown_table_attrs(unknown_table_attrs);
    }
    if !unknown_style_attrs.is_empty() {
        table.set_unknown_style_attrs(unknown_style_attrs);
    }

    Ok(Some(table))
}

/// Extracts `TableColumn` fields from a `<tableColumn>` element's attributes.
fn parse_table_column_attributes(event: &BytesStart<'_>) -> TableColumn {
    let mut col_name = String::new();
    let mut col_id: u32 = 0;
    let mut col_totals_label: Option<String> = None;
    let mut col_totals_function: Option<TotalFunction> = None;
    let mut col_unknown_attrs: Vec<(String, String)> = Vec::new();
    for attribute in event.attributes().flatten() {
        let key = local_name(attribute.key.as_ref());
        let raw_key = String::from_utf8_lossy(attribute.key.as_ref()).into_owned();
        let value = String::from_utf8_lossy(attribute.value.as_ref()).into_owned();
        match key {
            b"id" => {
                col_id = value.trim().parse::<u32>().unwrap_or(0);
            }
            b"name" => {
                col_name = value;
            }
            b"totalsRowLabel" => {
                if !value.trim().is_empty() {
                    col_totals_label = Some(value);
                }
            }
            b"totalsRowFunction" => {
                col_totals_function = TotalFunction::from_xml_value(value.as_str());
            }
            _ => {
                col_unknown_attrs.push((raw_key, value));
            }
        }
    }
    let mut tc = TableColumn::new(col_name, col_id);
    if let Some(label) = col_totals_label {
        tc.set_totals_row_label(label);
    }
    if let Some(func) = col_totals_function {
        tc.set_totals_row_function(func);
    }
    if !col_unknown_attrs.is_empty() {
        tc.set_unknown_attrs(col_unknown_attrs);
    }
    tc
}

fn parse_shared_strings_xml(xml: &[u8]) -> Result<Vec<SharedStringEntry>> {
    let mut reader = Reader::from_reader(Cursor::new(xml));
    reader.config_mut().trim_text(false);
    let mut buffer = Vec::new();

    let mut entries = Vec::new();

    // State tracking for the current <si> element
    let mut in_si = false;
    let mut in_text = false;
    let mut in_run = false;
    let mut in_rpr = false;
    let mut current_plain_text: Option<String> = None;
    let mut current_runs: Vec<RichTextRun> = Vec::new();
    let mut current_run_text = String::new();
    let mut current_run_bold: Option<bool> = None;
    let mut current_run_italic: Option<bool> = None;
    let mut current_run_font_name: Option<String> = None;
    let mut current_run_font_size: Option<String> = None;
    let mut current_run_color: Option<String> = None;
    let mut current_run_unknown_rpr: Vec<RawXmlNode> = Vec::new();
    let mut has_runs = false;
    let mut phonetic_runs: Vec<RawXmlNode> = Vec::new();
    let mut phonetic_pr: Option<RawXmlNode> = None;

    loop {
        match reader.read_event_into(&mut buffer)? {
            Event::Start(ref event) => {
                let name_bytes = event.name();
                let local = local_name(name_bytes.as_ref());
                match local {
                    b"si" => {
                        in_si = true;
                        current_plain_text = None;
                        current_runs.clear();
                        has_runs = false;
                        phonetic_runs.clear();
                        phonetic_pr = None;
                    }
                    b"t" if in_si && !in_rpr => {
                        in_text = true;
                        if !in_run && current_plain_text.is_none() {
                            current_plain_text = Some(String::new());
                        }
                    }
                    b"r" if in_si => {
                        in_run = true;
                        has_runs = true;
                        current_run_text.clear();
                        current_run_bold = None;
                        current_run_italic = None;
                        current_run_font_name = None;
                        current_run_font_size = None;
                        current_run_color = None;
                        current_run_unknown_rpr.clear();
                    }
                    b"rPr" if in_run => {
                        in_rpr = true;
                    }
                    b"b" if in_rpr => {
                        // <b/> means bold=true, <b val="0"/> means false.
                        // A Start event for <b> (non-empty) is unusual but handle it.
                        current_run_bold = Some(true);
                    }
                    b"i" if in_rpr => {
                        current_run_italic = Some(true);
                    }
                    b"rFont" if in_rpr => {
                        for attr in event.attributes().flatten() {
                            let key = local_name(attr.key.as_ref());
                            if key == b"val" {
                                current_run_font_name =
                                    Some(String::from_utf8_lossy(attr.value.as_ref()).into_owned());
                            }
                        }
                    }
                    b"sz" if in_rpr => {
                        for attr in event.attributes().flatten() {
                            let key = local_name(attr.key.as_ref());
                            if key == b"val" {
                                current_run_font_size =
                                    Some(String::from_utf8_lossy(attr.value.as_ref()).into_owned());
                            }
                        }
                    }
                    b"color" if in_rpr => {
                        for attr in event.attributes().flatten() {
                            let key = local_name(attr.key.as_ref());
                            if key == b"rgb" {
                                current_run_color =
                                    Some(String::from_utf8_lossy(attr.value.as_ref()).into_owned());
                            }
                        }
                        // Capture as unknown if it has non-rgb attributes
                        // (theme, indexed, etc.) for roundtrip
                        let has_rgb = event
                            .attributes()
                            .flatten()
                            .any(|a| local_name(a.key.as_ref()) == b"rgb");
                        if !has_rgb {
                            current_run_unknown_rpr
                                .push(RawXmlNode::read_element(&mut reader, event)?);
                            buffer.clear();
                            continue;
                        }
                    }
                    b"rPh" if in_si => {
                        phonetic_runs.push(RawXmlNode::read_element(&mut reader, event)?);
                        buffer.clear();
                        continue;
                    }
                    b"phoneticPr" if in_si => {
                        phonetic_pr = Some(RawXmlNode::read_element(&mut reader, event)?);
                        buffer.clear();
                        continue;
                    }
                    _ if in_rpr => {
                        // Unknown rPr child element — capture for roundtrip
                        current_run_unknown_rpr.push(RawXmlNode::read_element(&mut reader, event)?);
                        buffer.clear();
                        continue;
                    }
                    _ => {}
                }
            }
            Event::Empty(ref event) => {
                let name_bytes = event.name();
                let local = local_name(name_bytes.as_ref());
                match local {
                    b"si" => {
                        entries.push(SharedStringEntry::Plain(String::new()));
                    }
                    b"t" if in_si && !in_rpr => {
                        // Empty <t/> — just an empty text segment
                        if !in_run && current_plain_text.is_none() {
                            current_plain_text = Some(String::new());
                        }
                    }
                    b"b" if in_rpr => {
                        // Check for val="0"
                        let mut bold_val = true;
                        for attr in event.attributes().flatten() {
                            let key = local_name(attr.key.as_ref());
                            if key == b"val" {
                                let val = String::from_utf8_lossy(attr.value.as_ref()).into_owned();
                                if val == "0" || val == "false" {
                                    bold_val = false;
                                }
                            }
                        }
                        current_run_bold = Some(bold_val);
                    }
                    b"i" if in_rpr => {
                        let mut italic_val = true;
                        for attr in event.attributes().flatten() {
                            let key = local_name(attr.key.as_ref());
                            if key == b"val" {
                                let val = String::from_utf8_lossy(attr.value.as_ref()).into_owned();
                                if val == "0" || val == "false" {
                                    italic_val = false;
                                }
                            }
                        }
                        current_run_italic = Some(italic_val);
                    }
                    b"rFont" if in_rpr => {
                        for attr in event.attributes().flatten() {
                            let key = local_name(attr.key.as_ref());
                            if key == b"val" {
                                current_run_font_name =
                                    Some(String::from_utf8_lossy(attr.value.as_ref()).into_owned());
                            }
                        }
                    }
                    b"sz" if in_rpr => {
                        for attr in event.attributes().flatten() {
                            let key = local_name(attr.key.as_ref());
                            if key == b"val" {
                                current_run_font_size =
                                    Some(String::from_utf8_lossy(attr.value.as_ref()).into_owned());
                            }
                        }
                    }
                    b"color" if in_rpr => {
                        for attr in event.attributes().flatten() {
                            let key = local_name(attr.key.as_ref());
                            if key == b"rgb" {
                                current_run_color =
                                    Some(String::from_utf8_lossy(attr.value.as_ref()).into_owned());
                            }
                        }
                        // Capture as unknown if it has non-rgb attributes
                        let has_rgb = event
                            .attributes()
                            .flatten()
                            .any(|a| local_name(a.key.as_ref()) == b"rgb");
                        if !has_rgb {
                            current_run_unknown_rpr.push(RawXmlNode::from_empty_element(event));
                        }
                    }
                    b"rPh" if in_si => {
                        phonetic_runs.push(RawXmlNode::from_empty_element(event));
                    }
                    b"phoneticPr" if in_si => {
                        phonetic_pr = Some(RawXmlNode::from_empty_element(event));
                    }
                    _ if in_rpr => {
                        // Unknown rPr child — capture for roundtrip
                        current_run_unknown_rpr.push(RawXmlNode::from_empty_element(event));
                    }
                    _ => {}
                }
            }
            Event::Text(ref event) if in_text => {
                let text = event
                    .xml_content()
                    .map_err(quick_xml::Error::from)?
                    .into_owned();
                if in_run {
                    current_run_text.push_str(text.as_str());
                } else if let Some(current) = current_plain_text.as_mut() {
                    current.push_str(text.as_str());
                }
            }
            Event::CData(ref event) if in_text => {
                let text = String::from_utf8_lossy(event.as_ref()).into_owned();
                if in_run {
                    current_run_text.push_str(text.as_str());
                } else if let Some(current) = current_plain_text.as_mut() {
                    current.push_str(text.as_str());
                }
            }
            Event::End(ref event) => {
                let name_bytes = event.name();
                let local = local_name(name_bytes.as_ref());
                match local {
                    b"t" => {
                        in_text = false;
                    }
                    b"rPr" => {
                        in_rpr = false;
                    }
                    b"r" => {
                        let mut run = RichTextRun::new(std::mem::take(&mut current_run_text));
                        if let Some(bold) = current_run_bold.take() {
                            run.set_bold(bold);
                        }
                        if let Some(italic) = current_run_italic.take() {
                            run.set_italic(italic);
                        }
                        if let Some(name) = current_run_font_name.take() {
                            run.set_font_name(name);
                        }
                        if let Some(size) = current_run_font_size.take() {
                            run.set_font_size(size);
                        }
                        if let Some(color) = current_run_color.take() {
                            run.set_color(color);
                        }
                        if !current_run_unknown_rpr.is_empty() {
                            run.set_unknown_rpr_children(std::mem::take(
                                &mut current_run_unknown_rpr,
                            ));
                        }
                        current_runs.push(run);
                        in_run = false;
                    }
                    b"si" => {
                        if has_runs {
                            entries.push(SharedStringEntry::RichText {
                                runs: std::mem::take(&mut current_runs),
                                phonetic_runs: std::mem::take(&mut phonetic_runs),
                                phonetic_pr: phonetic_pr.take(),
                            });
                        } else {
                            let mut plain = current_plain_text.take().unwrap_or_default();
                            // If there are phonetic elements on a plain text entry,
                            // promote to rich text to preserve them.
                            if !phonetic_runs.is_empty() || phonetic_pr.is_some() {
                                let run = RichTextRun::new(std::mem::take(&mut plain));
                                entries.push(SharedStringEntry::RichText {
                                    runs: vec![run],
                                    phonetic_runs: std::mem::take(&mut phonetic_runs),
                                    phonetic_pr: phonetic_pr.take(),
                                });
                            } else {
                                entries.push(SharedStringEntry::Plain(plain));
                            }
                        }
                        in_si = false;
                        in_text = false;
                        in_run = false;
                        in_rpr = false;
                    }
                    _ => {}
                }
            }
            Event::Eof => break,
            _ => {}
        }
        buffer.clear();
    }

    Ok(entries)
}

#[allow(clippy::too_many_arguments)]
fn serialize_workbook_xml(
    sheet_refs: &[WorkbookSheetRef],
    defined_names: &[DefinedName],
    workbook_protection: Option<&WorkbookProtection>,
    calc_settings: Option<&CalculationSettings>,
    file_version_attrs: &[(String, String)],
    workbook_pr_attrs: &[(String, String)],
    book_views: &[Vec<(String, String)>],
    pivot_caches: &[(u32, String)],
    custom_workbook_views: &[RawXmlNode],
) -> Result<Vec<u8>> {
    let mut writer = Writer::new_with_indent(Vec::new(), b' ', 2);
    writer.write_event(Event::Decl(BytesDecl::new(
        "1.0",
        Some("UTF-8"),
        Some("yes"),
    )))?;

    let mut workbook = BytesStart::new("workbook");
    workbook.push_attribute(("xmlns", SPREADSHEETML_NS));
    workbook.push_attribute(("xmlns:r", RELATIONSHIP_NS));
    writer.write_event(Event::Start(workbook))?;

    // Serialize fileVersion (roundtrip preservation)
    if !file_version_attrs.is_empty() {
        let mut tag = BytesStart::new("fileVersion");
        for (key, value) in file_version_attrs {
            tag.push_attribute((key.as_str(), value.as_str()));
        }
        writer.write_event(Event::Empty(tag))?;
    }

    // Serialize workbookPr (roundtrip preservation)
    if !workbook_pr_attrs.is_empty() {
        let mut tag = BytesStart::new("workbookPr");
        for (key, value) in workbook_pr_attrs {
            tag.push_attribute((key.as_str(), value.as_str()));
        }
        writer.write_event(Event::Empty(tag))?;
    }

    // Serialize bookViews (roundtrip preservation)
    if !book_views.is_empty() {
        writer.write_event(Event::Start(BytesStart::new("bookViews")))?;
        for view_attrs in book_views {
            let mut tag = BytesStart::new("workbookView");
            for (key, value) in view_attrs {
                tag.push_attribute((key.as_str(), value.as_str()));
            }
            writer.write_event(Event::Empty(tag))?;
        }
        writer.write_event(Event::End(BytesEnd::new("bookViews")))?;
    }

    // Serialize workbookProtection
    if let Some(protection) = workbook_protection.filter(|p| p.has_metadata()) {
        let mut tag = BytesStart::new("workbookProtection");
        if protection.lock_structure {
            tag.push_attribute(("lockStructure", "1"));
        }
        if protection.lock_windows {
            tag.push_attribute(("lockWindows", "1"));
        }
        if let Some(hash) = protection.password_hash() {
            tag.push_attribute(("workbookPassword", hash));
        }
        writer.write_event(Event::Empty(tag))?;
    }

    writer.write_event(Event::Start(BytesStart::new("sheets")))?;
    for sheet_ref in sheet_refs {
        let mut sheet = BytesStart::new("sheet");
        let sheet_id = sheet_ref.sheet_id.to_string();
        sheet.push_attribute(("name", sheet_ref.name.as_str()));
        sheet.push_attribute(("sheetId", sheet_id.as_str()));
        if sheet_ref.visibility != SheetVisibility::Visible {
            sheet.push_attribute(("state", sheet_ref.visibility.as_xml_value()));
        }
        sheet.push_attribute(("r:id", sheet_ref.relationship_id.as_str()));
        writer.write_event(Event::Empty(sheet))?;
    }
    writer.write_event(Event::End(BytesEnd::new("sheets")))?;

    if !defined_names.is_empty() {
        writer.write_event(Event::Start(BytesStart::new("definedNames")))?;
        for defined_name in defined_names {
            if defined_name.name().is_empty() {
                continue;
            }

            let mut defined_name_tag = BytesStart::new("definedName");
            defined_name_tag.push_attribute(("name", defined_name.name()));

            let local_sheet_id = defined_name.local_sheet_id().map(|value| value.to_string());
            if let Some(local_sheet_id) = local_sheet_id.as_deref() {
                defined_name_tag.push_attribute(("localSheetId", local_sheet_id));
            }

            // Replay unknown attributes (hidden, comment, description, etc.)
            for (key, value) in &defined_name.unknown_attrs {
                defined_name_tag.push_attribute((key.as_str(), value.as_str()));
            }

            if defined_name.reference().is_empty() {
                writer.write_event(Event::Empty(defined_name_tag))?;
                continue;
            }

            writer.write_event(Event::Start(defined_name_tag))?;
            writer.write_event(Event::Text(BytesText::new(defined_name.reference())))?;
            writer.write_event(Event::End(BytesEnd::new("definedName")))?;
        }
        writer.write_event(Event::End(BytesEnd::new("definedNames")))?;
    }

    // Serialize calcPr
    if let Some(settings) = calc_settings.filter(|s| s.has_metadata()) {
        let mut tag = BytesStart::new("calcPr");
        if let Some(mode) = settings.calc_mode() {
            tag.push_attribute(("calcMode", mode));
        }
        if let Some(id) = settings.calc_id() {
            let id_text = id.to_string();
            tag.push_attribute(("calcId", id_text.as_str()));
        }
        if let Some(full_calc) = settings.full_calc_on_load() {
            tag.push_attribute(("fullCalcOnLoad", if full_calc { "1" } else { "0" }));
        }
        writer.write_event(Event::Empty(tag))?;
    }

    // Serialize pivotCaches
    if !pivot_caches.is_empty() {
        writer.write_event(Event::Start(BytesStart::new("pivotCaches")))?;
        for (cache_id, r_id) in pivot_caches {
            let mut tag = BytesStart::new("pivotCache");
            let cache_id_text = cache_id.to_string();
            tag.push_attribute(("cacheId", cache_id_text.as_str()));
            tag.push_attribute(("r:id", r_id.as_str()));
            writer.write_event(Event::Empty(tag))?;
        }
        writer.write_event(Event::End(BytesEnd::new("pivotCaches")))?;
    }

    // Serialize customWorkbookViews (roundtrip preservation)
    if !custom_workbook_views.is_empty() {
        writer.write_event(Event::Start(BytesStart::new("customWorkbookViews")))?;
        for node in custom_workbook_views {
            node.write_to(&mut writer)?;
        }
        writer.write_event(Event::End(BytesEnd::new("customWorkbookViews")))?;
    }

    writer.write_event(Event::End(BytesEnd::new("workbook")))?;

    Ok(writer.into_inner())
}

/// Extracts sparkline groups from `<extLst>` in a worksheet's unknown children.
fn extract_sparklines_from_ext_lst(worksheet: &mut Worksheet) {
    let unknown_children = worksheet.unknown_children_mut();
    let mut sparkline_groups: Vec<SparklineGroup> = Vec::new();

    for ext_lst_node in unknown_children.iter_mut() {
        let RawXmlNode::Element { name, children, .. } = ext_lst_node else {
            continue;
        };

        let local = strip_ns_prefix(name.as_str());
        if local != "extLst" {
            continue;
        }

        children.retain(|ext_child| {
            let RawXmlNode::Element {
                name: ext_name,
                attributes: ext_attrs,
                children: ext_children,
            } = ext_child
            else {
                return true;
            };

            let ext_local = strip_ns_prefix(ext_name.as_str());
            if ext_local != "ext" {
                return true;
            }

            let has_sparkline_uri = ext_attrs
                .iter()
                .any(|(key, val)| key == "uri" && val.eq_ignore_ascii_case(SPARKLINE_EXT_URI));
            if !has_sparkline_uri {
                return true;
            }

            for sparkline_groups_node in ext_children {
                let RawXmlNode::Element {
                    name: sg_name,
                    children: sg_children,
                    ..
                } = sparkline_groups_node
                else {
                    continue;
                };

                let sg_local = strip_ns_prefix(sg_name.as_str());
                if sg_local != "sparklineGroups" {
                    continue;
                }

                for group_node in sg_children {
                    if let Some(group) = parse_sparkline_group_node(group_node) {
                        sparkline_groups.push(group);
                    }
                }
            }

            false
        });
    }

    unknown_children.retain(|node| {
        if let RawXmlNode::Element { name, children, .. } = node {
            let local = strip_ns_prefix(name.as_str());
            if local == "extLst" && children.is_empty() {
                return false;
            }
        }
        true
    });

    if !sparkline_groups.is_empty() {
        worksheet.set_parsed_sparkline_groups(sparkline_groups);
    }
}

fn parse_sparkline_group_node(node: &RawXmlNode) -> Option<SparklineGroup> {
    let RawXmlNode::Element {
        name,
        attributes,
        children,
    } = node
    else {
        return None;
    };

    let local = strip_ns_prefix(name.as_str());
    if local != "sparklineGroup" {
        return None;
    }

    let mut group = SparklineGroup::new();

    for (key, val) in attributes {
        match key.as_str() {
            "type" => {
                group.set_sparkline_type(SparklineType::from_xml_value(val));
            }
            "displayEmptyCellsAs" => {
                group.set_display_empty_cells_as(SparklineEmptyCells::from_xml_value(val));
            }
            "markers" => {
                group.set_markers(val == "1" || val == "true");
            }
            "high" => {
                group.set_high_point(val == "1" || val == "true");
            }
            "low" => {
                group.set_low_point(val == "1" || val == "true");
            }
            "first" => {
                group.set_first_point(val == "1" || val == "true");
            }
            "last" => {
                group.set_last_point(val == "1" || val == "true");
            }
            "negative" => {
                group.set_negative_points(val == "1" || val == "true");
            }
            "displayXAxis" => {
                group.set_display_x_axis(val == "1" || val == "true");
            }
            "displayHidden" => {
                group.set_display_hidden(val == "1" || val == "true");
            }
            "minAxisType" => {
                group.set_min_axis_type(SparklineAxisType::from_xml_value(val));
            }
            "maxAxisType" => {
                group.set_max_axis_type(SparklineAxisType::from_xml_value(val));
            }
            "rightToLeft" => {
                group.set_right_to_left(val == "1" || val == "true");
            }
            "manualMin" => {
                if let Ok(v) = val.parse::<f64>() {
                    group.set_manual_min(v);
                }
            }
            "manualMax" => {
                if let Ok(v) = val.parse::<f64>() {
                    group.set_manual_max(v);
                }
            }
            "lineWeight" => {
                if let Ok(v) = val.parse::<f64>() {
                    group.set_line_weight(v);
                }
            }
            "dateAxis" => {
                group.set_date_axis(val == "1" || val == "true");
            }
            _ => {}
        }
    }

    let mut colors = SparklineColors::new();
    for child in children {
        let RawXmlNode::Element {
            name: child_name,
            attributes: child_attrs,
            children: child_children,
        } = child
        else {
            continue;
        };

        let child_local = strip_ns_prefix(child_name.as_str());
        match child_local {
            "colorSeries" => {
                colors.series = extract_color_rgb(child_attrs);
            }
            "colorNegative" => {
                colors.negative = extract_color_rgb(child_attrs);
            }
            "colorAxis" => {
                colors.axis = extract_color_rgb(child_attrs);
            }
            "colorMarkers" => {
                colors.markers = extract_color_rgb(child_attrs);
            }
            "colorFirst" => {
                colors.first = extract_color_rgb(child_attrs);
            }
            "colorLast" => {
                colors.last = extract_color_rgb(child_attrs);
            }
            "colorHigh" => {
                colors.high = extract_color_rgb(child_attrs);
            }
            "colorLow" => {
                colors.low = extract_color_rgb(child_attrs);
            }
            "sparklines" => {
                for sparkline_child in child_children {
                    if let Some(sparkline) = parse_sparkline_node(sparkline_child) {
                        group.add_sparkline(sparkline);
                    }
                }
            }
            _ => {}
        }
    }

    if colors.has_any() {
        group.set_colors(colors);
    }

    Some(group)
}

fn extract_color_rgb(attrs: &[(String, String)]) -> Option<String> {
    for (key, val) in attrs {
        if key == "rgb" {
            return Some(val.clone());
        }
    }
    for (key, val) in attrs {
        if key == "theme" {
            return Some(format!("theme:{val}"));
        }
        if key == "indexed" {
            return Some(format!("indexed:{val}"));
        }
    }
    None
}

fn parse_sparkline_node(node: &RawXmlNode) -> Option<Sparkline> {
    let RawXmlNode::Element { name, children, .. } = node else {
        return None;
    };

    let local = strip_ns_prefix(name.as_str());
    if local != "sparkline" {
        return None;
    }

    let mut formula: Option<String> = None;
    let mut sqref: Option<String> = None;

    for child in children {
        if let RawXmlNode::Element {
            name: child_name,
            children: child_children,
            ..
        } = child
        {
            let child_local = strip_ns_prefix(child_name.as_str());
            match child_local {
                "f" => {
                    formula = extract_text_content(child_children);
                }
                "sqref" => {
                    sqref = extract_text_content(child_children);
                }
                _ => {}
            }
        }
    }

    let location = sqref?;
    let data_range = formula.unwrap_or_default();
    Some(Sparkline::new(location, data_range))
}

fn extract_text_content(children: &[RawXmlNode]) -> Option<String> {
    let mut text = String::new();
    for child in children {
        if let RawXmlNode::Text(t) = child {
            text.push_str(t);
        }
    }
    if text.is_empty() {
        None
    } else {
        Some(text)
    }
}

fn strip_ns_prefix(name: &str) -> &str {
    name.rsplit(':').next().unwrap_or(name)
}

fn serialize_sparklines_ext_lst<W: std::io::Write>(
    writer: &mut Writer<W>,
    sparkline_groups: &[SparklineGroup],
    has_existing_ext_lst: bool,
) -> std::result::Result<(), quick_xml::Error> {
    if sparkline_groups.is_empty() {
        return Ok(());
    }

    if !has_existing_ext_lst {
        writer.write_event(Event::Start(BytesStart::new("extLst")))?;
    }

    let mut ext_tag = BytesStart::new("ext");
    ext_tag.push_attribute(("uri", SPARKLINE_EXT_URI));
    ext_tag.push_attribute(("xmlns:x14", X14_NS));
    writer.write_event(Event::Start(ext_tag))?;

    let mut sg_root = BytesStart::new("x14:sparklineGroups");
    sg_root.push_attribute(("xmlns:xm", XM_NS));
    writer.write_event(Event::Start(sg_root))?;

    for group in sparkline_groups {
        serialize_sparkline_group(writer, group)?;
    }

    writer.write_event(Event::End(BytesEnd::new("x14:sparklineGroups")))?;
    writer.write_event(Event::End(BytesEnd::new("ext")))?;

    if !has_existing_ext_lst {
        writer.write_event(Event::End(BytesEnd::new("extLst")))?;
    }

    Ok(())
}

fn serialize_sparkline_group<W: std::io::Write>(
    writer: &mut Writer<W>,
    group: &SparklineGroup,
) -> std::result::Result<(), quick_xml::Error> {
    let mut tag = BytesStart::new("x14:sparklineGroup");

    if group.sparkline_type() != SparklineType::Line {
        tag.push_attribute(("type", group.sparkline_type().as_str()));
    }
    if group.display_empty_cells_as() != SparklineEmptyCells::Gap {
        tag.push_attribute((
            "displayEmptyCellsAs",
            group.display_empty_cells_as().as_str(),
        ));
    }
    if group.markers() {
        tag.push_attribute(("markers", "1"));
    }
    if group.high_point() {
        tag.push_attribute(("high", "1"));
    }
    if group.low_point() {
        tag.push_attribute(("low", "1"));
    }
    if group.first_point() {
        tag.push_attribute(("first", "1"));
    }
    if group.last_point() {
        tag.push_attribute(("last", "1"));
    }
    if group.negative_points() {
        tag.push_attribute(("negative", "1"));
    }
    if group.display_x_axis() {
        tag.push_attribute(("displayXAxis", "1"));
    }
    if group.display_hidden() {
        tag.push_attribute(("displayHidden", "1"));
    }
    if group.min_axis_type() != SparklineAxisType::Individual {
        tag.push_attribute(("minAxisType", group.min_axis_type().as_str()));
    }
    if group.max_axis_type() != SparklineAxisType::Individual {
        tag.push_attribute(("maxAxisType", group.max_axis_type().as_str()));
    }
    if group.right_to_left() {
        tag.push_attribute(("rightToLeft", "1"));
    }
    if let Some(weight) = group.line_weight() {
        let weight_text = weight.to_string();
        tag.push_attribute(("lineWeight", weight_text.as_str()));
    }
    if group.date_axis() {
        tag.push_attribute(("dateAxis", "1"));
    }
    if let Some(manual_min) = group.manual_min() {
        let min_text = manual_min.to_string();
        tag.push_attribute(("manualMin", min_text.as_str()));
    }
    if let Some(manual_max) = group.manual_max() {
        let max_text = manual_max.to_string();
        tag.push_attribute(("manualMax", max_text.as_str()));
    }

    writer.write_event(Event::Start(tag))?;

    let colors = group.colors();
    write_sparkline_color(writer, "x14:colorSeries", colors.series.as_deref())?;
    write_sparkline_color(writer, "x14:colorNegative", colors.negative.as_deref())?;
    write_sparkline_color(writer, "x14:colorAxis", colors.axis.as_deref())?;
    write_sparkline_color(writer, "x14:colorMarkers", colors.markers.as_deref())?;
    write_sparkline_color(writer, "x14:colorFirst", colors.first.as_deref())?;
    write_sparkline_color(writer, "x14:colorLast", colors.last.as_deref())?;
    write_sparkline_color(writer, "x14:colorHigh", colors.high.as_deref())?;
    write_sparkline_color(writer, "x14:colorLow", colors.low.as_deref())?;

    if !group.sparklines().is_empty() {
        writer.write_event(Event::Start(BytesStart::new("x14:sparklines")))?;
        for sparkline in group.sparklines() {
            writer.write_event(Event::Start(BytesStart::new("x14:sparkline")))?;

            writer.write_event(Event::Start(BytesStart::new("xm:f")))?;
            writer.write_event(Event::Text(BytesText::new(sparkline.data_range())))?;
            writer.write_event(Event::End(BytesEnd::new("xm:f")))?;

            writer.write_event(Event::Start(BytesStart::new("xm:sqref")))?;
            writer.write_event(Event::Text(BytesText::new(sparkline.location())))?;
            writer.write_event(Event::End(BytesEnd::new("xm:sqref")))?;

            writer.write_event(Event::End(BytesEnd::new("x14:sparkline")))?;
        }
        writer.write_event(Event::End(BytesEnd::new("x14:sparklines")))?;
    }

    writer.write_event(Event::End(BytesEnd::new("x14:sparklineGroup")))?;
    Ok(())
}

fn write_sparkline_color<W: std::io::Write>(
    writer: &mut Writer<W>,
    tag_name: &str,
    color: Option<&str>,
) -> std::result::Result<(), quick_xml::Error> {
    if let Some(color_val) = color {
        let mut tag = BytesStart::new(tag_name);
        if let Some(stripped) = color_val.strip_prefix("theme:") {
            tag.push_attribute(("theme", stripped));
        } else if let Some(stripped) = color_val.strip_prefix("indexed:") {
            tag.push_attribute(("indexed", stripped));
        } else {
            tag.push_attribute(("rgb", color_val));
        }
        writer.write_event(Event::Empty(tag))?;
    }
    Ok(())
}

/// Returns `true` if two columns have identical metadata (excluding index),
/// so they can be compacted into a single `<col min="..." max="...">` range.
fn columns_equal(a: &Column, b: &Column) -> bool {
    a.width() == b.width()
        && a.is_hidden() == b.is_hidden()
        && a.outline_level() == b.outline_level()
        && a.is_collapsed() == b.is_collapsed()
        && a.style_index() == b.style_index()
        && a.is_best_fit() == b.is_best_fit()
        && a.custom_width() == b.custom_width()
}

fn serialize_worksheet_xml(
    worksheet: &Worksheet,
    shared_strings: &mut SharedStrings,
    table_refs: &[WorksheetTablePartRef],
    drawing_ref: Option<&WorksheetDrawingPartRef>,
    hyperlink_refs: &[WorksheetHyperlinkRef],
) -> Result<Vec<u8>> {
    let mut writer = Writer::new_with_indent(Vec::new(), b' ', 2);
    writer.write_event(Event::Decl(BytesDecl::new(
        "1.0",
        Some("UTF-8"),
        Some("yes"),
    )))?;

    let mut worksheet_tag = BytesStart::new("worksheet");
    worksheet_tag.push_attribute(("xmlns", SPREADSHEETML_NS));
    if !table_refs.is_empty() || drawing_ref.is_some() || !hyperlink_refs.is_empty() {
        worksheet_tag.push_attribute(("xmlns:r", RELATIONSHIP_NS));
    }
    // Replay extra namespace declarations from the original XML so that unknown
    // attributes using those prefixes (e.g. x14ac:dyDescent) remain valid.
    for (prefix, uri) in worksheet.extra_namespace_declarations() {
        worksheet_tag.push_attribute((prefix.as_str(), uri.as_str()));
    }
    writer.write_event(Event::Start(worksheet_tag))?;

    // Write <sheetPr> if we have tab color.
    if let Some(tab_color) = worksheet.tab_color() {
        writer.write_event(Event::Start(BytesStart::new("sheetPr")))?;
        let mut tab_color_tag = BytesStart::new("tabColor");
        let argb = format!("FF{tab_color}");
        tab_color_tag.push_attribute(("rgb", argb.as_str()));
        writer.write_event(Event::Empty(tab_color_tag))?;
        writer.write_event(Event::End(BytesEnd::new("sheetPr")))?;
    }

    // Write <dimension> if we have a saved dimension ref.
    if let Some(dim_ref) = worksheet.raw_dimension_ref() {
        let mut tag = BytesStart::new("dimension");
        tag.push_attribute(("ref", dim_ref));
        writer.write_event(Event::Empty(tag))?;
    }

    // Write <sheetFormatPr> if we have raw attrs from parsing or typed values.
    {
        let raw_attrs = worksheet.raw_sheet_format_pr_attrs();
        if !raw_attrs.is_empty() {
            // Use raw attributes to preserve all original values (roundtrip fidelity).
            // Override known typed fields if the user has programmatically changed them.
            let mut fmt_pr = BytesStart::new("sheetFormatPr");
            let drh_override = worksheet.default_row_height().map(|v| v.to_string());
            let dcw_override = worksheet.default_column_width().map(|v| v.to_string());
            let ch_override = worksheet.custom_height();
            for (key, value) in raw_attrs {
                let local = local_name(key.as_bytes());
                match local {
                    b"defaultRowHeight" => {
                        if let Some(drh) = drh_override.as_deref() {
                            fmt_pr.push_attribute((key.as_str(), drh));
                        } else {
                            fmt_pr.push_attribute((key.as_str(), value.as_str()));
                        }
                    }
                    b"defaultColWidth" => {
                        if let Some(dcw) = dcw_override.as_deref() {
                            fmt_pr.push_attribute((key.as_str(), dcw));
                        } else {
                            fmt_pr.push_attribute((key.as_str(), value.as_str()));
                        }
                    }
                    b"customHeight" => {
                        if let Some(ch) = ch_override {
                            fmt_pr.push_attribute((key.as_str(), if ch { "1" } else { "0" }));
                        } else {
                            fmt_pr.push_attribute((key.as_str(), value.as_str()));
                        }
                    }
                    _ => {
                        fmt_pr.push_attribute((key.as_str(), value.as_str()));
                    }
                }
            }
            writer.write_event(Event::Empty(fmt_pr))?;
        } else {
            // Fallback for programmatically created worksheets (no raw attrs).
            let has_format_pr = worksheet.default_row_height().is_some()
                || worksheet.default_column_width().is_some()
                || worksheet.custom_height().is_some();
            if has_format_pr {
                let mut fmt_pr = BytesStart::new("sheetFormatPr");
                let drh_text = worksheet.default_row_height().map(|v| v.to_string());
                if let Some(drh) = drh_text.as_deref() {
                    fmt_pr.push_attribute(("defaultRowHeight", drh));
                }
                let dcw_text = worksheet.default_column_width().map(|v| v.to_string());
                if let Some(dcw) = dcw_text.as_deref() {
                    fmt_pr.push_attribute(("defaultColWidth", dcw));
                }
                if let Some(ch) = worksheet.custom_height() {
                    fmt_pr.push_attribute(("customHeight", if ch { "1" } else { "0" }));
                }
                writer.write_event(Event::Empty(fmt_pr))?;
            }
        }
    }

    // Write <sheetViews> if we have freeze panes or view options.
    let has_view_options = worksheet
        .sheet_view_options()
        .is_some_and(|opts| opts.has_metadata());
    if worksheet.freeze_pane().is_some() || has_view_options {
        writer.write_event(Event::Start(BytesStart::new("sheetViews")))?;

        let mut sheet_view = BytesStart::new("sheetView");
        sheet_view.push_attribute(("workbookViewId", "0"));

        // Apply view options
        if let Some(options) = worksheet.sheet_view_options() {
            if let Some(show_gridlines) = options.show_gridlines() {
                sheet_view
                    .push_attribute(("showGridLines", if show_gridlines { "1" } else { "0" }));
            }
            if let Some(show_headers) = options.show_row_col_headers() {
                sheet_view
                    .push_attribute(("showRowColHeaders", if show_headers { "1" } else { "0" }));
            }
            if let Some(show_formulas) = options.show_formulas() {
                sheet_view.push_attribute(("showFormulas", if show_formulas { "1" } else { "0" }));
            }
            let zoom_text = options.zoom_scale().map(|v| v.to_string());
            if let Some(zoom_text) = zoom_text.as_deref() {
                sheet_view.push_attribute(("zoomScale", zoom_text));
            }
            let zoom_normal_text = options.zoom_scale_normal().map(|v| v.to_string());
            if let Some(zoom_normal_text) = zoom_normal_text.as_deref() {
                sheet_view.push_attribute(("zoomScaleNormal", zoom_normal_text));
            }
            if let Some(rtl) = options.right_to_left() {
                sheet_view.push_attribute(("rightToLeft", if rtl { "1" } else { "0" }));
            }
            if let Some(tab_selected) = options.tab_selected() {
                sheet_view.push_attribute(("tabSelected", if tab_selected { "1" } else { "0" }));
            }
            if let Some(view) = options.view() {
                sheet_view.push_attribute(("view", view));
            }
        }

        writer.write_event(Event::Start(sheet_view))?;

        if let Some(freeze_pane) = worksheet.freeze_pane() {
            let mut pane = BytesStart::new("pane");
            let x_split_text = freeze_pane.x_split().to_string();
            let y_split_text = freeze_pane.y_split().to_string();
            if freeze_pane.x_split() > 0 {
                pane.push_attribute(("xSplit", x_split_text.as_str()));
            }
            if freeze_pane.y_split() > 0 {
                pane.push_attribute(("ySplit", y_split_text.as_str()));
            }
            pane.push_attribute(("topLeftCell", freeze_pane.top_left_cell()));
            pane.push_attribute(("state", "frozen"));

            let active_pane = match (freeze_pane.x_split(), freeze_pane.y_split()) {
                (x, y) if x > 0 && y > 0 => Some("bottomRight"),
                (x, 0) if x > 0 => Some("topRight"),
                (0, y) if y > 0 => Some("bottomLeft"),
                _ => None,
            };
            if let Some(active_pane) = active_pane {
                pane.push_attribute(("activePane", active_pane));
            }

            writer.write_event(Event::Empty(pane))?;
        }

        writer.write_event(Event::End(BytesEnd::new("sheetView")))?;
        writer.write_event(Event::End(BytesEnd::new("sheetViews")))?;
    }

    // Write <cols> section for column metadata (width, hidden, outline, style, etc.).
    {
        let cols_to_write: Vec<&Column> = worksheet
            .columns()
            .filter(|col| {
                col.width().is_some()
                    || col.is_hidden()
                    || col.outline_level() > 0
                    || col.is_collapsed()
                    || col.style_index().is_some()
                    || col.is_best_fit()
            })
            .collect();
        if !cols_to_write.is_empty() {
            // Compact adjacent columns with identical metadata into ranges.
            let mut col_ranges: Vec<(u32, u32, &Column)> = Vec::new();
            for col in &cols_to_write {
                if let Some(last) = col_ranges.last_mut() {
                    let (_, ref mut max, representative) = *last;
                    if col.index() == *max + 1 && columns_equal(col, representative) {
                        *max = col.index();
                        continue;
                    }
                }
                col_ranges.push((col.index(), col.index(), col));
            }
            writer.write_event(Event::Start(BytesStart::new("cols")))?;
            for (min, max, col) in &col_ranges {
                let mut col_tag = BytesStart::new("col");
                let min_text = min.to_string();
                let max_text = max.to_string();
                col_tag.push_attribute(("min", min_text.as_str()));
                col_tag.push_attribute(("max", max_text.as_str()));
                if let Some(width) = col.width() {
                    let width_text = width.to_string();
                    col_tag.push_attribute(("width", width_text.as_str()));
                }
                if let Some(style_index) = col.style_index() {
                    if style_index > 0 {
                        let style_text = style_index.to_string();
                        col_tag.push_attribute(("style", style_text.as_str()));
                    }
                }
                if col.is_hidden() {
                    col_tag.push_attribute(("hidden", "1"));
                }
                if col.is_best_fit() {
                    col_tag.push_attribute(("bestFit", "1"));
                }
                if col.custom_width() {
                    col_tag.push_attribute(("customWidth", "1"));
                }
                if col.outline_level() > 0 {
                    let level_text = col.outline_level().to_string();
                    col_tag.push_attribute(("outlineLevel", level_text.as_str()));
                }
                if col.is_collapsed() {
                    col_tag.push_attribute(("collapsed", "1"));
                }
                writer.write_event(Event::Empty(col_tag))?;
            }
            writer.write_event(Event::End(BytesEnd::new("cols")))?;
        }
    }

    writer.write_event(Event::Start(BytesStart::new("sheetData")))?;

    let mut row_cells: BTreeMap<u32, BTreeMap<String, &Cell>> = BTreeMap::new();
    for (reference, cell) in worksheet.cells() {
        let row_index = row_index_from_reference(reference)?;
        row_cells
            .entry(row_index)
            .or_default()
            .insert(reference.to_string(), cell);
    }

    let mut row_order: BTreeMap<u32, Option<&Row>> = BTreeMap::new();
    for row in worksheet.rows() {
        row_order.insert(row.index(), Some(row));
    }
    for row_index in row_cells.keys() {
        row_order.entry(*row_index).or_insert(None);
    }

    for (row_index, row_metadata) in row_order {
        let mut row_cells = row_cells.remove(&row_index).unwrap_or_default();
        let mut row = BytesStart::new("row");
        let mut row_known_keys: Vec<&str> = vec!["r"];
        let row_index_text = row_index.to_string();
        row.push_attribute(("r", row_index_text.as_str()));
        let row_height_text = row_metadata
            .and_then(Row::height)
            .map(|value| value.to_string());
        if let Some(row_height_text) = row_height_text.as_deref() {
            row.push_attribute(("ht", row_height_text));
            if row_metadata.is_some_and(|m| m.custom_height()) {
                row.push_attribute(("customHeight", "1"));
            }
            row_known_keys.extend(["ht", "customHeight"]);
        }
        if row_metadata.is_some_and(|m| m.is_hidden()) {
            row.push_attribute(("hidden", "1"));
            row_known_keys.push("hidden");
        }
        if let Some(outline_level) = row_metadata
            .map(|m| m.outline_level())
            .filter(|level| *level > 0)
        {
            let level_text = outline_level.to_string();
            row.push_attribute(("outlineLevel", level_text.as_str()));
            row_known_keys.push("outlineLevel");
        }
        if row_metadata.is_some_and(|m| m.is_collapsed()) {
            row.push_attribute(("collapsed", "1"));
            row_known_keys.push("collapsed");
        }
        if let Some(row_metadata) = row_metadata {
            offidized_opc::xml_util::push_unknown_attrs_deduped(
                &mut row,
                row_metadata.unknown_attrs(),
                &row_known_keys,
            );
        }

        let has_row_children = !row_cells.is_empty()
            || row_metadata.is_some_and(|metadata| !metadata.unknown_children().is_empty());
        if !has_row_children {
            writer.write_event(Event::Empty(row))?;
            continue;
        }

        writer.write_event(Event::Start(row))?;

        for (reference, cell) in std::mem::take(&mut row_cells) {
            let mut cell_tag = BytesStart::new("c");
            let mut cell_known_keys: Vec<&str> = vec!["r"];
            cell_tag.push_attribute(("r", reference.as_str()));

            if let Some(style_id) = cell.style_id() {
                let style_text = style_id.to_string();
                cell_tag.push_attribute(("s", style_text.as_str()));
                cell_known_keys.push("s");
            }

            // Unknown attrs are written AFTER the type attribute below, so we
            // must also track `t` as a known key to prevent duplication.
            cell_known_keys.push("t");

            offidized_opc::xml_util::push_unknown_attrs_deduped(
                &mut cell_tag,
                cell.unknown_attrs(),
                &cell_known_keys,
            );

            // Determine the cell type attribute from the primary value or, for
            // formula cells, from the cached value.
            let type_source = cell.value().or(cell.cached_value());
            if let Some(value) = type_source {
                match value {
                    CellValue::String(_) | CellValue::RichText(_) => {
                        // For formula cells with a cached string result, use
                        // t="str" (computed string) rather than t="s" (shared
                        // string index) because the cached value is the literal
                        // text, not a shared string index.
                        if cell.formula().is_some() || cell.shared_formula_index().is_some() {
                            cell_tag.push_attribute(("t", "str"));
                        } else {
                            cell_tag.push_attribute(("t", "s"));
                        }
                    }
                    CellValue::Bool(_) => cell_tag.push_attribute(("t", "b")),
                    CellValue::Date(_) => cell_tag.push_attribute(("t", "d")),
                    CellValue::Error(_) => cell_tag.push_attribute(("t", "e")),
                    CellValue::Blank | CellValue::Number(_) | CellValue::DateTime(_) => {}
                }
            }

            let has_formula = cell.formula().is_some();
            let has_shared_formula = cell.shared_formula_index().is_some();
            let has_value = cell
                .value()
                .is_some_and(|value| !matches!(value, CellValue::Blank));
            let has_cached_value = cell
                .cached_value()
                .is_some_and(|value| !matches!(value, CellValue::Blank));
            let has_unknown_children = !cell.unknown_children().is_empty();
            let has_array_formula = cell.is_array_formula();

            if !has_formula
                && !has_shared_formula
                && !has_value
                && !has_cached_value
                && !has_unknown_children
            {
                writer.write_event(Event::Empty(cell_tag))?;
                continue;
            }

            writer.write_event(Event::Start(cell_tag))?;

            if let Some(formula) = cell.formula() {
                // Master cell with formula text
                if has_array_formula {
                    let mut f_tag = BytesStart::new("f");
                    f_tag.push_attribute(("t", "array"));
                    if let Some(array_range) = cell.array_range() {
                        f_tag.push_attribute(("ref", array_range));
                    }
                    writer.write_event(Event::Start(f_tag))?;
                } else if let Some(si) = cell.shared_formula_index() {
                    let mut f_tag = BytesStart::new("f");
                    f_tag.push_attribute(("t", "shared"));
                    let si_text = si.to_string();
                    f_tag.push_attribute(("si", si_text.as_str()));
                    // For shared formula masters, include the ref range if set
                    if let Some(array_range) = cell.array_range() {
                        f_tag.push_attribute(("ref", array_range));
                    }
                    writer.write_event(Event::Start(f_tag))?;
                } else {
                    writer.write_event(Event::Start(BytesStart::new("f")))?;
                }
                writer.write_event(Event::Text(BytesText::new(formula)))?;
                writer.write_event(Event::End(BytesEnd::new("f")))?;
            } else if let Some(si) = cell.shared_formula_index() {
                // Dependent cell: shared formula index but no formula text
                // Write as self-closing <f t="shared" si="N"/>
                let mut f_tag = BytesStart::new("f");
                f_tag.push_attribute(("t", "shared"));
                let si_text = si.to_string();
                f_tag.push_attribute(("si", si_text.as_str()));
                writer.write_event(Event::Empty(f_tag))?;
            }

            if let Some(value) = cell.value() {
                match value {
                    CellValue::String(text) => {
                        let shared_string_index = shared_strings.intern(text.as_str());
                        let shared_string_index_text = shared_string_index.to_string();
                        writer.write_event(Event::Start(BytesStart::new("v")))?;
                        writer.write_event(Event::Text(BytesText::new(
                            shared_string_index_text.as_str(),
                        )))?;
                        writer.write_event(Event::End(BytesEnd::new("v")))?;
                    }
                    CellValue::RichText(runs) => {
                        // Rich text is stored in shared strings with formatting.
                        let shared_string_index =
                            shared_strings.intern_rich_text(runs.clone(), Vec::new(), None);
                        let shared_string_index_text = shared_string_index.to_string();
                        writer.write_event(Event::Start(BytesStart::new("v")))?;
                        writer.write_event(Event::Text(BytesText::new(
                            shared_string_index_text.as_str(),
                        )))?;
                        writer.write_event(Event::End(BytesEnd::new("v")))?;
                    }
                    CellValue::Number(number) => {
                        let number_text = number.to_string();
                        writer.write_event(Event::Start(BytesStart::new("v")))?;
                        writer.write_event(Event::Text(BytesText::new(number_text.as_str())))?;
                        writer.write_event(Event::End(BytesEnd::new("v")))?;
                    }
                    CellValue::DateTime(serial) => {
                        // DateTime is stored as a numeric value (serial date number)
                        let serial_text = serial.to_string();
                        writer.write_event(Event::Start(BytesStart::new("v")))?;
                        writer.write_event(Event::Text(BytesText::new(serial_text.as_str())))?;
                        writer.write_event(Event::End(BytesEnd::new("v")))?;
                    }
                    CellValue::Bool(boolean) => {
                        writer.write_event(Event::Start(BytesStart::new("v")))?;
                        writer.write_event(Event::Text(BytesText::new(if *boolean {
                            "1"
                        } else {
                            "0"
                        })))?;
                        writer.write_event(Event::End(BytesEnd::new("v")))?;
                    }
                    CellValue::Date(date) => {
                        writer.write_event(Event::Start(BytesStart::new("v")))?;
                        writer.write_event(Event::Text(BytesText::new(date.as_str())))?;
                        writer.write_event(Event::End(BytesEnd::new("v")))?;
                    }
                    CellValue::Error(error) => {
                        writer.write_event(Event::Start(BytesStart::new("v")))?;
                        writer.write_event(Event::Text(BytesText::new(error.as_str())))?;
                        writer.write_event(Event::End(BytesEnd::new("v")))?;
                    }
                    CellValue::Blank => {}
                }
            }

            // Write the cached formula result as <v> when the cell has a
            // formula but no primary value. Cached values are written as
            // literal text (not shared string indices).
            if cell.value().is_none() || matches!(cell.value(), Some(CellValue::Blank)) {
                if let Some(cached) = cell.cached_value() {
                    match cached {
                        CellValue::String(text)
                        | CellValue::Error(text)
                        | CellValue::Date(text) => {
                            writer.write_event(Event::Start(BytesStart::new("v")))?;
                            writer.write_event(Event::Text(BytesText::new(text.as_str())))?;
                            writer.write_event(Event::End(BytesEnd::new("v")))?;
                        }
                        CellValue::Number(number) => {
                            let number_text = number.to_string();
                            writer.write_event(Event::Start(BytesStart::new("v")))?;
                            writer
                                .write_event(Event::Text(BytesText::new(number_text.as_str())))?;
                            writer.write_event(Event::End(BytesEnd::new("v")))?;
                        }
                        CellValue::DateTime(serial) => {
                            let serial_text = serial.to_string();
                            writer.write_event(Event::Start(BytesStart::new("v")))?;
                            writer
                                .write_event(Event::Text(BytesText::new(serial_text.as_str())))?;
                            writer.write_event(Event::End(BytesEnd::new("v")))?;
                        }
                        CellValue::Bool(boolean) => {
                            writer.write_event(Event::Start(BytesStart::new("v")))?;
                            writer.write_event(Event::Text(BytesText::new(if *boolean {
                                "1"
                            } else {
                                "0"
                            })))?;
                            writer.write_event(Event::End(BytesEnd::new("v")))?;
                        }
                        CellValue::RichText(runs) => {
                            let plain_text: String = runs.iter().map(|r| r.text()).collect();
                            writer.write_event(Event::Start(BytesStart::new("v")))?;
                            writer.write_event(Event::Text(BytesText::new(plain_text.as_str())))?;
                            writer.write_event(Event::End(BytesEnd::new("v")))?;
                        }
                        CellValue::Blank => {}
                    }
                }
            }

            for node in cell.unknown_children() {
                node.write_to(&mut writer)?;
            }

            writer.write_event(Event::End(BytesEnd::new("c")))?;
        }

        if let Some(row_metadata) = row_metadata {
            for node in row_metadata.unknown_children() {
                node.write_to(&mut writer)?;
            }
        }

        writer.write_event(Event::End(BytesEnd::new("row")))?;
    }

    writer.write_event(Event::End(BytesEnd::new("sheetData")))?;

    // Write sheetProtection
    if let Some(protection) = worksheet.protection().filter(|p| p.has_metadata()) {
        let mut tag = BytesStart::new("sheetProtection");
        if protection.sheet() {
            tag.push_attribute(("sheet", "1"));
        }
        if protection.objects() {
            tag.push_attribute(("objects", "1"));
        }
        if protection.scenarios() {
            tag.push_attribute(("scenarios", "1"));
        }
        if protection.format_cells() {
            tag.push_attribute(("formatCells", "1"));
        }
        if protection.format_columns() {
            tag.push_attribute(("formatColumns", "1"));
        }
        if protection.format_rows() {
            tag.push_attribute(("formatRows", "1"));
        }
        if protection.insert_columns() {
            tag.push_attribute(("insertColumns", "1"));
        }
        if protection.insert_rows() {
            tag.push_attribute(("insertRows", "1"));
        }
        if protection.insert_hyperlinks() {
            tag.push_attribute(("insertHyperlinks", "1"));
        }
        if protection.delete_columns() {
            tag.push_attribute(("deleteColumns", "1"));
        }
        if protection.delete_rows() {
            tag.push_attribute(("deleteRows", "1"));
        }
        if protection.select_locked_cells() {
            tag.push_attribute(("selectLockedCells", "1"));
        }
        if protection.sort() {
            tag.push_attribute(("sort", "1"));
        }
        if protection.auto_filter() {
            tag.push_attribute(("autoFilter", "1"));
        }
        if protection.pivot_tables() {
            tag.push_attribute(("pivotTables", "1"));
        }
        if protection.select_unlocked_cells() {
            tag.push_attribute(("selectUnlockedCells", "1"));
        }
        if let Some(hash) = protection.password_hash() {
            tag.push_attribute(("password", hash));
        }
        writer.write_event(Event::Empty(tag))?;
    }

    if let Some(auto_filter) = worksheet.auto_filter() {
        if let Some(range) = auto_filter.range() {
            let has_columns = !auto_filter.filter_columns().is_empty();
            let mut auto_filter_tag = BytesStart::new("autoFilter");
            let range_str = format_cell_range(range);
            auto_filter_tag.push_attribute(("ref", range_str.as_str()));

            if !has_columns {
                writer.write_event(Event::Empty(auto_filter_tag))?;
            } else {
                writer.write_event(Event::Start(auto_filter_tag))?;
                for fc in auto_filter.filter_columns() {
                    let mut fc_tag = BytesStart::new("filterColumn");
                    let col_id_text = fc.col_id().to_string();
                    fc_tag.push_attribute(("colId", col_id_text.as_str()));
                    if fc.show_button() == Some(false) {
                        fc_tag.push_attribute(("showButton", "0"));
                    }

                    match fc.filter_type() {
                        FilterType::Values => {
                            writer.write_event(Event::Start(fc_tag))?;
                            let mut filters_tag = BytesStart::new("filters");
                            if fc.blank() == Some(true) {
                                filters_tag.push_attribute(("blank", "1"));
                            }
                            writer.write_event(Event::Start(filters_tag))?;
                            for val in fc.values() {
                                let mut f_tag = BytesStart::new("filter");
                                f_tag.push_attribute(("val", val.as_str()));
                                writer.write_event(Event::Empty(f_tag))?;
                            }
                            writer.write_event(Event::End(BytesEnd::new("filters")))?;
                            writer.write_event(Event::End(BytesEnd::new("filterColumn")))?;
                        }
                        FilterType::Custom => {
                            writer.write_event(Event::Start(fc_tag))?;
                            let mut cf_tag = BytesStart::new("customFilters");
                            if fc.custom_filters_and() == Some(true) {
                                cf_tag.push_attribute(("and", "1"));
                            }
                            writer.write_event(Event::Start(cf_tag))?;
                            for cf in fc.custom_filters() {
                                let mut cf_elem = BytesStart::new("customFilter");
                                if cf.operator() != CustomFilterOperator::Equal {
                                    cf_elem.push_attribute(("operator", cf.operator().as_str()));
                                }
                                cf_elem.push_attribute(("val", cf.val()));
                                writer.write_event(Event::Empty(cf_elem))?;
                            }
                            writer.write_event(Event::End(BytesEnd::new("customFilters")))?;
                            writer.write_event(Event::End(BytesEnd::new("filterColumn")))?;
                        }
                        FilterType::Top10 => {
                            writer.write_event(Event::Start(fc_tag))?;
                            let mut t10 = BytesStart::new("top10");
                            if let Some(top) = fc.top() {
                                t10.push_attribute(("top", if top { "1" } else { "0" }));
                            }
                            if let Some(percent) = fc.percent() {
                                t10.push_attribute(("percent", if percent { "1" } else { "0" }));
                            }
                            let val_text = fc.top10_val().map(|v| v.to_string());
                            if let Some(val_text) = val_text.as_deref() {
                                t10.push_attribute(("val", val_text));
                            }
                            writer.write_event(Event::Empty(t10))?;
                            writer.write_event(Event::End(BytesEnd::new("filterColumn")))?;
                        }
                        FilterType::Dynamic => {
                            writer.write_event(Event::Start(fc_tag))?;
                            let mut df = BytesStart::new("dynamicFilter");
                            if let Some(dt) = fc.dynamic_type() {
                                df.push_attribute(("type", dt));
                            }
                            writer.write_event(Event::Empty(df))?;
                            writer.write_event(Event::End(BytesEnd::new("filterColumn")))?;
                        }
                        FilterType::Color => {
                            writer.write_event(Event::Start(fc_tag))?;
                            let mut cf = BytesStart::new("colorFilter");
                            let dxf_id_text = fc.dxf_id().map(|v| v.to_string());
                            if let Some(dxf_id) = dxf_id_text.as_deref() {
                                cf.push_attribute(("dxfId", dxf_id));
                            }
                            if let Some(cell_color) = fc.cell_color() {
                                cf.push_attribute((
                                    "cellColor",
                                    if cell_color { "1" } else { "0" },
                                ));
                            }
                            writer.write_event(Event::Empty(cf))?;
                            writer.write_event(Event::End(BytesEnd::new("filterColumn")))?;
                        }
                        FilterType::Icon => {
                            writer.write_event(Event::Start(fc_tag))?;
                            let mut icon = BytesStart::new("iconFilter");
                            if let Some(icon_set) = fc.icon_set() {
                                icon.push_attribute(("iconSet", icon_set));
                            }
                            let icon_id_text = fc.icon_id().map(|v| v.to_string());
                            if let Some(icon_id) = icon_id_text.as_deref() {
                                icon.push_attribute(("iconId", icon_id));
                            }
                            writer.write_event(Event::Empty(icon))?;
                            writer.write_event(Event::End(BytesEnd::new("filterColumn")))?;
                        }
                        FilterType::None => {
                            writer.write_event(Event::Empty(fc_tag))?;
                        }
                    }
                }
                writer.write_event(Event::End(BytesEnd::new("autoFilter")))?;
            }
        }
    }

    if !worksheet.merged_ranges().is_empty() {
        let mut merge_cells = BytesStart::new("mergeCells");
        let count_text = worksheet.merged_ranges().len().to_string();
        merge_cells.push_attribute(("count", count_text.as_str()));
        writer.write_event(Event::Start(merge_cells))?;

        for range in worksheet.merged_ranges() {
            let mut merge_cell = BytesStart::new("mergeCell");
            let reference = format_cell_range(range);
            merge_cell.push_attribute(("ref", reference.as_str()));
            writer.write_event(Event::Empty(merge_cell))?;
        }

        writer.write_event(Event::End(BytesEnd::new("mergeCells")))?;
    }

    if !worksheet.conditional_formattings().is_empty() {
        for (index, conditional_formatting) in
            worksheet.conditional_formattings().iter().enumerate()
        {
            let mut conditional_formatting_tag = BytesStart::new("conditionalFormatting");
            let sqref = conditional_formatting.sqref_xml();
            conditional_formatting_tag.push_attribute(("sqref", sqref.as_str()));
            writer.write_event(Event::Start(conditional_formatting_tag))?;

            let mut cf_rule = BytesStart::new("cfRule");
            cf_rule.push_attribute(("type", conditional_formatting.rule_type().as_xml_type()));

            // Use parsed priority if available, otherwise fall back to index-based priority.
            let priority_str = conditional_formatting
                .priority()
                .unwrap_or((index + 1) as u32)
                .to_string();
            cf_rule.push_attribute(("priority", priority_str.as_str()));

            if let Some(op) = conditional_formatting.operator() {
                cf_rule.push_attribute(("operator", op.as_xml_value()));
            }
            if let Some(dxf_id) = conditional_formatting.dxf_id() {
                let dxf_id_str = dxf_id.to_string();
                cf_rule.push_attribute(("dxfId", dxf_id_str.as_str()));
            }
            if let Some(true) = conditional_formatting.stop_if_true() {
                cf_rule.push_attribute(("stopIfTrue", "1"));
            }
            if let Some(text) = conditional_formatting.text() {
                cf_rule.push_attribute(("text", text));
            }
            if let Some(tp) = conditional_formatting.time_period() {
                cf_rule.push_attribute(("timePeriod", tp));
            }
            if let Some(rank) = conditional_formatting.rank() {
                let rank_str = rank.to_string();
                cf_rule.push_attribute(("rank", rank_str.as_str()));
            }
            if let Some(true) = conditional_formatting.cf_percent() {
                cf_rule.push_attribute(("percent", "1"));
            }
            if let Some(true) = conditional_formatting.cf_bottom() {
                cf_rule.push_attribute(("bottom", "1"));
            }
            if let Some(false) = conditional_formatting.above_average() {
                cf_rule.push_attribute(("aboveAverage", "0"));
            }
            if let Some(true) = conditional_formatting.equal_average() {
                cf_rule.push_attribute(("equalAverage", "1"));
            }
            if let Some(sd) = conditional_formatting.std_dev() {
                let sd_str = sd.to_string();
                cf_rule.push_attribute(("stdDev", sd_str.as_str()));
            }

            // Determine if cfRule has child elements
            let has_formulas = !conditional_formatting.formulas().is_empty();
            let has_color_scale = !conditional_formatting.color_scale_stops().is_empty();
            let has_data_bar = conditional_formatting.data_bar_min().is_some()
                || conditional_formatting.data_bar_max().is_some();
            let has_icon_set = conditional_formatting.icon_set_name().is_some()
                || !conditional_formatting.icon_set_values().is_empty();
            let has_children = has_formulas || has_color_scale || has_data_bar || has_icon_set;

            if has_children {
                writer.write_event(Event::Start(cf_rule))?;

                // Write <colorScale> child element
                if has_color_scale {
                    writer.write_event(Event::Start(BytesStart::new("colorScale")))?;
                    // Write all <cfvo> elements first, then all <color> elements
                    for stop in conditional_formatting.color_scale_stops() {
                        let mut cfvo_tag = BytesStart::new("cfvo");
                        cfvo_tag.push_attribute(("type", stop.cfvo.value_type.as_xml_value()));
                        if let Some(ref val) = stop.cfvo.value {
                            cfvo_tag.push_attribute(("val", val.as_str()));
                        }
                        writer.write_event(Event::Empty(cfvo_tag))?;
                    }
                    for stop in conditional_formatting.color_scale_stops() {
                        let mut color_tag = BytesStart::new("color");
                        write_color_attribute(&mut color_tag, &stop.color);
                        writer.write_event(Event::Empty(color_tag))?;
                    }
                    writer.write_event(Event::End(BytesEnd::new("colorScale")))?;
                }

                // Write <dataBar> child element
                if has_data_bar {
                    let mut db_tag = BytesStart::new("dataBar");
                    if conditional_formatting.data_bar_show_value() == Some(false) {
                        db_tag.push_attribute(("showValue", "0"));
                    }
                    let min_len_text = conditional_formatting
                        .data_bar_min_length()
                        .map(|v| v.to_string());
                    if let Some(ref v) = min_len_text {
                        db_tag.push_attribute(("minLength", v.as_str()));
                    }
                    let max_len_text = conditional_formatting
                        .data_bar_max_length()
                        .map(|v| v.to_string());
                    if let Some(ref v) = max_len_text {
                        db_tag.push_attribute(("maxLength", v.as_str()));
                    }
                    writer.write_event(Event::Start(db_tag))?;
                    if let Some(min) = conditional_formatting.data_bar_min() {
                        let mut cfvo_tag = BytesStart::new("cfvo");
                        cfvo_tag.push_attribute(("type", min.value_type.as_xml_value()));
                        if let Some(ref val) = min.value {
                            cfvo_tag.push_attribute(("val", val.as_str()));
                        }
                        writer.write_event(Event::Empty(cfvo_tag))?;
                    }
                    if let Some(max) = conditional_formatting.data_bar_max() {
                        let mut cfvo_tag = BytesStart::new("cfvo");
                        cfvo_tag.push_attribute(("type", max.value_type.as_xml_value()));
                        if let Some(ref val) = max.value {
                            cfvo_tag.push_attribute(("val", val.as_str()));
                        }
                        writer.write_event(Event::Empty(cfvo_tag))?;
                    }
                    if let Some(color) = conditional_formatting.data_bar_color() {
                        let mut color_tag = BytesStart::new("color");
                        write_color_attribute(&mut color_tag, color);
                        writer.write_event(Event::Empty(color_tag))?;
                    }
                    writer.write_event(Event::End(BytesEnd::new("dataBar")))?;
                }

                // Write <iconSet> child element
                if has_icon_set {
                    let mut icon_set_tag = BytesStart::new("iconSet");
                    if let Some(name) = conditional_formatting.icon_set_name() {
                        icon_set_tag.push_attribute(("iconSet", name));
                    }
                    if let Some(false) = conditional_formatting.icon_set_show_value() {
                        icon_set_tag.push_attribute(("showValue", "0"));
                    }
                    if let Some(true) = conditional_formatting.icon_set_reverse() {
                        icon_set_tag.push_attribute(("reverse", "1"));
                    }
                    if conditional_formatting.icon_set_values().is_empty() {
                        writer.write_event(Event::Empty(icon_set_tag))?;
                    } else {
                        writer.write_event(Event::Start(icon_set_tag))?;
                        for cfvo in conditional_formatting.icon_set_values() {
                            let mut cfvo_tag = BytesStart::new("cfvo");
                            cfvo_tag.push_attribute(("type", cfvo.value_type.as_xml_value()));
                            if let Some(ref val) = cfvo.value {
                                cfvo_tag.push_attribute(("val", val.as_str()));
                            }
                            writer.write_event(Event::Empty(cfvo_tag))?;
                        }
                        writer.write_event(Event::End(BytesEnd::new("iconSet")))?;
                    }
                }

                // Write <formula> child elements
                for formula in conditional_formatting.formulas() {
                    writer.write_event(Event::Start(BytesStart::new("formula")))?;
                    writer.write_event(Event::Text(BytesText::new(formula.as_str())))?;
                    writer.write_event(Event::End(BytesEnd::new("formula")))?;
                }

                writer.write_event(Event::End(BytesEnd::new("cfRule")))?;
            } else {
                writer.write_event(Event::Empty(cf_rule))?;
            }

            writer.write_event(Event::End(BytesEnd::new("conditionalFormatting")))?;
        }
    }

    if !worksheet.data_validations().is_empty() {
        let mut data_validations = BytesStart::new("dataValidations");
        let count_text = worksheet.data_validations().len().to_string();
        data_validations.push_attribute(("count", count_text.as_str()));
        writer.write_event(Event::Start(data_validations))?;

        for data_validation in worksheet.data_validations() {
            let mut data_validation_tag = BytesStart::new("dataValidation");
            data_validation_tag
                .push_attribute(("type", data_validation.validation_type().as_xml_type()));
            let sqref = data_validation.sqref_xml();
            data_validation_tag.push_attribute(("sqref", sqref.as_str()));

            if let Some(error_style) = data_validation.error_style() {
                data_validation_tag.push_attribute(("errorStyle", error_style.as_xml_value()));
            }
            if let Some(error_title) = data_validation.error_title() {
                data_validation_tag.push_attribute(("errorTitle", error_title));
            }
            if let Some(error_message) = data_validation.error_message() {
                data_validation_tag.push_attribute(("error", error_message));
            }
            if let Some(prompt_title) = data_validation.prompt_title() {
                data_validation_tag.push_attribute(("promptTitle", prompt_title));
            }
            if let Some(prompt_message) = data_validation.prompt_message() {
                data_validation_tag.push_attribute(("prompt", prompt_message));
            }
            if let Some(show_input) = data_validation.show_input_message() {
                data_validation_tag
                    .push_attribute(("showInputMessage", if show_input { "1" } else { "0" }));
            }
            if let Some(show_error) = data_validation.show_error_message() {
                data_validation_tag
                    .push_attribute(("showErrorMessage", if show_error { "1" } else { "0" }));
            }

            writer.write_event(Event::Start(data_validation_tag))?;

            writer.write_event(Event::Start(BytesStart::new("formula1")))?;
            writer.write_event(Event::Text(BytesText::new(data_validation.formula1())))?;
            writer.write_event(Event::End(BytesEnd::new("formula1")))?;

            if let Some(formula2) = data_validation.formula2() {
                writer.write_event(Event::Start(BytesStart::new("formula2")))?;
                writer.write_event(Event::Text(BytesText::new(formula2)))?;
                writer.write_event(Event::End(BytesEnd::new("formula2")))?;
            }

            writer.write_event(Event::End(BytesEnd::new("dataValidation")))?;
        }

        writer.write_event(Event::End(BytesEnd::new("dataValidations")))?;
    }

    // Write <hyperlinks> section.
    if !worksheet.hyperlinks().is_empty() {
        // hyperlink_refs contains only external hyperlinks (those with a URL),
        // in the same order they appear in the filtered iterator.
        let mut external_ref_index = 0_usize;
        writer.write_event(Event::Start(BytesStart::new("hyperlinks")))?;
        for hyperlink in worksheet.hyperlinks() {
            let mut hyperlink_tag = BytesStart::new("hyperlink");
            hyperlink_tag.push_attribute(("ref", hyperlink.cell_ref()));
            if hyperlink.url().is_some() {
                if let Some(href) = hyperlink_refs.get(external_ref_index) {
                    hyperlink_tag.push_attribute(("r:id", href.relationship_id.as_str()));
                    external_ref_index += 1;
                }
            }
            if let Some(location) = hyperlink.location() {
                hyperlink_tag.push_attribute(("location", location));
            }
            if let Some(tooltip) = hyperlink.tooltip() {
                hyperlink_tag.push_attribute(("tooltip", tooltip));
            }
            writer.write_event(Event::Empty(hyperlink_tag))?;
        }
        writer.write_event(Event::End(BytesEnd::new("hyperlinks")))?;
    }

    // Write printOptions
    {
        let print_options = worksheet.raw_print_options_attrs();
        if !print_options.is_empty() {
            let mut tag = BytesStart::new("printOptions");
            for (key, value) in print_options {
                tag.push_attribute((key.as_str(), value.as_str()));
            }
            writer.write_event(Event::Empty(tag))?;
        }
    }

    // Write pageMargins
    if let Some(margins) = worksheet.page_margins().filter(|m| m.has_metadata()) {
        let mut tag = BytesStart::new("pageMargins");
        let left_text = margins.left().map(|v| v.to_string());
        let right_text = margins.right().map(|v| v.to_string());
        let top_text = margins.top().map(|v| v.to_string());
        let bottom_text = margins.bottom().map(|v| v.to_string());
        let header_text = margins.header().map(|v| v.to_string());
        let footer_text = margins.footer().map(|v| v.to_string());
        if let Some(left) = left_text.as_deref() {
            tag.push_attribute(("left", left));
        }
        if let Some(right) = right_text.as_deref() {
            tag.push_attribute(("right", right));
        }
        if let Some(top) = top_text.as_deref() {
            tag.push_attribute(("top", top));
        }
        if let Some(bottom) = bottom_text.as_deref() {
            tag.push_attribute(("bottom", bottom));
        }
        if let Some(header) = header_text.as_deref() {
            tag.push_attribute(("header", header));
        }
        if let Some(footer) = footer_text.as_deref() {
            tag.push_attribute(("footer", footer));
        }
        writer.write_event(Event::Empty(tag))?;
    }

    // Write pageSetup
    if let Some(page_setup) = worksheet.page_setup().filter(|p| p.has_metadata()) {
        let mut tag = BytesStart::new("pageSetup");
        if let Some(orientation) = page_setup.orientation() {
            tag.push_attribute(("orientation", orientation.as_xml_value()));
        }
        let paper_size_text = page_setup.paper_size().map(|v| v.to_string());
        if let Some(paper_size) = paper_size_text.as_deref() {
            tag.push_attribute(("paperSize", paper_size));
        }
        let scale_text = page_setup.scale().map(|v| v.to_string());
        if let Some(scale) = scale_text.as_deref() {
            tag.push_attribute(("scale", scale));
        }
        let fit_to_width_text = page_setup.fit_to_width().map(|v| v.to_string());
        if let Some(fit_to_width) = fit_to_width_text.as_deref() {
            tag.push_attribute(("fitToWidth", fit_to_width));
        }
        let fit_to_height_text = page_setup.fit_to_height().map(|v| v.to_string());
        if let Some(fit_to_height) = fit_to_height_text.as_deref() {
            tag.push_attribute(("fitToHeight", fit_to_height));
        }
        let first_page_text = page_setup.first_page_number().map(|v| v.to_string());
        if let Some(first_page) = first_page_text.as_deref() {
            tag.push_attribute(("firstPageNumber", first_page));
        }
        writer.write_event(Event::Empty(tag))?;
    }

    // Write headerFooter
    if let Some(hf) = worksheet.header_footer().filter(|h| h.has_metadata()) {
        let mut hf_tag = BytesStart::new("headerFooter");
        if hf.different_odd_even() {
            hf_tag.push_attribute(("differentOddEven", "1"));
        }
        if hf.different_first() {
            hf_tag.push_attribute(("differentFirst", "1"));
        }
        writer.write_event(Event::Start(hf_tag))?;
        if let Some(text) = hf.odd_header() {
            writer.write_event(Event::Start(BytesStart::new("oddHeader")))?;
            writer.write_event(Event::Text(BytesText::new(text)))?;
            writer.write_event(Event::End(BytesEnd::new("oddHeader")))?;
        }
        if let Some(text) = hf.odd_footer() {
            writer.write_event(Event::Start(BytesStart::new("oddFooter")))?;
            writer.write_event(Event::Text(BytesText::new(text)))?;
            writer.write_event(Event::End(BytesEnd::new("oddFooter")))?;
        }
        if let Some(text) = hf.even_header() {
            writer.write_event(Event::Start(BytesStart::new("evenHeader")))?;
            writer.write_event(Event::Text(BytesText::new(text)))?;
            writer.write_event(Event::End(BytesEnd::new("evenHeader")))?;
        }
        if let Some(text) = hf.even_footer() {
            writer.write_event(Event::Start(BytesStart::new("evenFooter")))?;
            writer.write_event(Event::Text(BytesText::new(text)))?;
            writer.write_event(Event::End(BytesEnd::new("evenFooter")))?;
        }
        if let Some(text) = hf.first_header() {
            writer.write_event(Event::Start(BytesStart::new("firstHeader")))?;
            writer.write_event(Event::Text(BytesText::new(text)))?;
            writer.write_event(Event::End(BytesEnd::new("firstHeader")))?;
        }
        if let Some(text) = hf.first_footer() {
            writer.write_event(Event::Start(BytesStart::new("firstFooter")))?;
            writer.write_event(Event::Text(BytesText::new(text)))?;
            writer.write_event(Event::End(BytesEnd::new("firstFooter")))?;
        }
        writer.write_event(Event::End(BytesEnd::new("headerFooter")))?;
    }

    // Write rowBreaks / colBreaks
    if let Some(breaks) = worksheet.page_breaks() {
        if !breaks.row_breaks().is_empty() {
            let mut rb_tag = BytesStart::new("rowBreaks");
            let count_text = breaks.row_breaks().len().to_string();
            rb_tag.push_attribute(("count", count_text.as_str()));
            rb_tag.push_attribute(("manualBreakCount", count_text.as_str()));
            writer.write_event(Event::Start(rb_tag))?;
            for brk in breaks.row_breaks() {
                let mut brk_tag = BytesStart::new("brk");
                let id_text = brk.id().to_string();
                brk_tag.push_attribute(("id", id_text.as_str()));
                if brk.manual() {
                    brk_tag.push_attribute(("man", "1"));
                }
                writer.write_event(Event::Empty(brk_tag))?;
            }
            writer.write_event(Event::End(BytesEnd::new("rowBreaks")))?;
        }
        if !breaks.col_breaks().is_empty() {
            let mut cb_tag = BytesStart::new("colBreaks");
            let count_text = breaks.col_breaks().len().to_string();
            cb_tag.push_attribute(("count", count_text.as_str()));
            cb_tag.push_attribute(("manualBreakCount", count_text.as_str()));
            writer.write_event(Event::Start(cb_tag))?;
            for brk in breaks.col_breaks() {
                let mut brk_tag = BytesStart::new("brk");
                let id_text = brk.id().to_string();
                brk_tag.push_attribute(("id", id_text.as_str()));
                if brk.manual() {
                    brk_tag.push_attribute(("man", "1"));
                }
                writer.write_event(Event::Empty(brk_tag))?;
            }
            writer.write_event(Event::End(BytesEnd::new("colBreaks")))?;
        }
    }

    if let Some(drawing_ref) = drawing_ref {
        let mut drawing_tag = BytesStart::new("drawing");
        drawing_tag.push_attribute(("r:id", drawing_ref.relationship_id.as_str()));
        writer.write_event(Event::Empty(drawing_tag))?;
    }

    if !table_refs.is_empty() {
        let mut table_parts = BytesStart::new("tableParts");
        let count = table_refs.len().to_string();
        table_parts.push_attribute(("count", count.as_str()));
        writer.write_event(Event::Start(table_parts))?;

        for table_ref in table_refs {
            let mut table_part = BytesStart::new("tablePart");
            table_part.push_attribute(("r:id", table_ref.relationship_id.as_str()));
            writer.write_event(Event::Empty(table_part))?;
        }

        writer.write_event(Event::End(BytesEnd::new("tableParts")))?;
    }

    // Emit unknown children preserved from parsing for roundtrip fidelity.
    // We need special handling for <extLst>: if sparklines exist we must merge
    // them into the extLst rather than creating a duplicate element.
    let has_sparklines = !worksheet.sparkline_groups().is_empty();
    let mut found_ext_lst = false;

    for node in worksheet.unknown_children() {
        if has_sparklines {
            if let RawXmlNode::Element {
                name,
                attributes,
                children,
            } = node
            {
                let local = strip_ns_prefix(name.as_str());
                if local == "extLst" {
                    // Write the extLst open tag manually so we can inject sparkline ext.
                    let mut start = BytesStart::new(name.as_str());
                    for (key, val) in attributes {
                        start.push_attribute((key.as_str(), val.as_str()));
                    }
                    writer.write_event(Event::Start(start))?;

                    // Write existing ext children.
                    for child in children {
                        child.write_to(&mut writer)?;
                    }

                    // Inject sparkline ext inside this extLst.
                    serialize_sparklines_ext_lst(&mut writer, worksheet.sparkline_groups(), true)?;

                    writer.write_event(Event::End(BytesEnd::new(name.as_str())))?;
                    found_ext_lst = true;
                    continue;
                }
            }
        }
        node.write_to(&mut writer)?;
    }

    // If sparklines exist but there was no extLst in unknown_children, create a new one.
    if has_sparklines && !found_ext_lst {
        serialize_sparklines_ext_lst(&mut writer, worksheet.sparkline_groups(), false)?;
    }

    writer.write_event(Event::End(BytesEnd::new("worksheet")))?;

    Ok(writer.into_inner())
}

fn serialize_drawing_xml(
    images: &[WorksheetImage],
    image_refs: &[WorksheetImagePartRef],
    charts: &[Chart],
    chart_rel_ids: &[String],
) -> Result<Vec<u8>> {
    if images.len() != image_refs.len() {
        return Err(XlsxError::InvalidWorkbookState(
            "drawing image references do not match worksheet images".to_string(),
        ));
    }
    if charts.len() != chart_rel_ids.len() {
        return Err(XlsxError::InvalidWorkbookState(
            "chart relationship IDs do not match worksheet charts".to_string(),
        ));
    }

    let mut writer = Writer::new_with_indent(Vec::new(), b' ', 2);
    writer.write_event(Event::Decl(BytesDecl::new(
        "1.0",
        Some("UTF-8"),
        Some("yes"),
    )))?;

    let mut root = BytesStart::new("xdr:wsDr");
    root.push_attribute(("xmlns:xdr", DRAWINGML_SPREADSHEET_NS));
    root.push_attribute(("xmlns:a", DRAWINGML_NS));
    root.push_attribute(("xmlns:r", RELATIONSHIP_NS));
    if !charts.is_empty() {
        root.push_attribute(("xmlns:c", CHART_NS));
    }
    writer.write_event(Event::Start(root))?;

    for (index, (image, image_ref)) in images.iter().zip(image_refs.iter()).enumerate() {
        let (column_index, row_index) = cell_reference_to_column_row(image.anchor_cell())?;
        let zero_based_column = column_index.checked_sub(1).ok_or_else(|| {
            XlsxError::InvalidWorkbookState("worksheet image anchor column underflow".to_string())
        })?;
        let zero_based_row = row_index.checked_sub(1).ok_or_else(|| {
            XlsxError::InvalidWorkbookState("worksheet image anchor row underflow".to_string())
        })?;
        let ext = match image.ext() {
            Some(ext) => ext,
            None => WorksheetImageExt::new(DEFAULT_IMAGE_EXTENT_CX, DEFAULT_IMAGE_EXTENT_CY)?,
        };

        let anchor_tag = match image.anchor_type() {
            ImageAnchorType::TwoCell => "xdr:twoCellAnchor",
            ImageAnchorType::OneCell => "xdr:oneCellAnchor",
            ImageAnchorType::Absolute => "xdr:absoluteAnchor",
        };
        writer.write_event(Event::Start(BytesStart::new(anchor_tag)))?;

        match image.anchor_type() {
            ImageAnchorType::Absolute => {
                let mut pos_tag = BytesStart::new("xdr:pos");
                let x_text = image.position_x().unwrap_or(0).to_string();
                let y_text = image.position_y().unwrap_or(0).to_string();
                pos_tag.push_attribute(("x", x_text.as_str()));
                pos_tag.push_attribute(("y", y_text.as_str()));
                writer.write_event(Event::Empty(pos_tag))?;
                let mut ext_tag = BytesStart::new("xdr:ext");
                let cx_text = image.extent_cx().unwrap_or(ext.cx() as i64).to_string();
                let cy_text = image.extent_cy().unwrap_or(ext.cy() as i64).to_string();
                ext_tag.push_attribute(("cx", cx_text.as_str()));
                ext_tag.push_attribute(("cy", cy_text.as_str()));
                writer.write_event(Event::Empty(ext_tag))?;
            }
            ImageAnchorType::OneCell | ImageAnchorType::TwoCell => {
                serialize_cell_anchor_element(
                    &mut writer,
                    "xdr:from",
                    image
                        .from_anchor()
                        .map(|a| (a.col(), a.row(), a.col_offset(), a.row_offset())),
                    zero_based_column,
                    zero_based_row,
                )?;
                if image.anchor_type() == ImageAnchorType::TwoCell {
                    if let Some(to) = image.to_anchor() {
                        serialize_cell_anchor_element(
                            &mut writer,
                            "xdr:to",
                            Some((to.col(), to.row(), to.col_offset(), to.row_offset())),
                            0,
                            0,
                        )?;
                    } else {
                        serialize_cell_anchor_element(
                            &mut writer,
                            "xdr:to",
                            None,
                            zero_based_column + 1,
                            zero_based_row + 1,
                        )?;
                    }
                } else {
                    let mut ext_tag = BytesStart::new("xdr:ext");
                    let cx_text = image.extent_cx().unwrap_or(ext.cx() as i64).to_string();
                    let cy_text = image.extent_cy().unwrap_or(ext.cy() as i64).to_string();
                    ext_tag.push_attribute(("cx", cx_text.as_str()));
                    ext_tag.push_attribute(("cy", cy_text.as_str()));
                    writer.write_event(Event::Empty(ext_tag))?;
                }
            }
        }

        writer.write_event(Event::Start(BytesStart::new("xdr:pic")))?;
        writer.write_event(Event::Start(BytesStart::new("xdr:nvPicPr")))?;

        let mut c_nv_pr = BytesStart::new("xdr:cNvPr");
        let image_id = u32::try_from(index + 1).map_err(|_| {
            XlsxError::InvalidWorkbookState(
                "too many worksheet images to serialize drawing metadata".to_string(),
            )
        })?;
        let image_id_text = image_id.to_string();
        let image_name = format!("Image {}", index + 1);
        c_nv_pr.push_attribute(("id", image_id_text.as_str()));
        c_nv_pr.push_attribute(("name", image_name.as_str()));
        writer.write_event(Event::Empty(c_nv_pr))?;
        writer.write_event(Event::Empty(BytesStart::new("xdr:cNvPicPr")))?;
        writer.write_event(Event::End(BytesEnd::new("xdr:nvPicPr")))?;

        writer.write_event(Event::Start(BytesStart::new("xdr:blipFill")))?;
        let mut blip = BytesStart::new("a:blip");
        blip.push_attribute(("r:embed", image_ref.relationship_id.as_str()));
        writer.write_event(Event::Empty(blip))?;
        let has_crop = image.crop_left().is_some()
            || image.crop_right().is_some()
            || image.crop_top().is_some()
            || image.crop_bottom().is_some();
        if has_crop {
            let mut src_rect = BytesStart::new("a:srcRect");
            if let Some(l) = image.crop_left() {
                let l_text = ((l * 100_000.0) as i64).to_string();
                src_rect.push_attribute(("l", l_text.as_str()));
            }
            if let Some(t) = image.crop_top() {
                let t_text = ((t * 100_000.0) as i64).to_string();
                src_rect.push_attribute(("t", t_text.as_str()));
            }
            if let Some(r) = image.crop_right() {
                let r_text = ((r * 100_000.0) as i64).to_string();
                src_rect.push_attribute(("r", r_text.as_str()));
            }
            if let Some(b) = image.crop_bottom() {
                let b_text = ((b * 100_000.0) as i64).to_string();
                src_rect.push_attribute(("b", b_text.as_str()));
            }
            writer.write_event(Event::Empty(src_rect))?;
        }
        writer.write_event(Event::Start(BytesStart::new("a:stretch")))?;
        writer.write_event(Event::Empty(BytesStart::new("a:fillRect")))?;
        writer.write_event(Event::End(BytesEnd::new("a:stretch")))?;
        writer.write_event(Event::End(BytesEnd::new("xdr:blipFill")))?;

        writer.write_event(Event::Start(BytesStart::new("xdr:spPr")))?;
        let mut preset_geometry = BytesStart::new("a:prstGeom");
        preset_geometry.push_attribute(("prst", "rect"));
        writer.write_event(Event::Start(preset_geometry))?;
        writer.write_event(Event::Empty(BytesStart::new("a:avLst")))?;
        writer.write_event(Event::End(BytesEnd::new("a:prstGeom")))?;
        writer.write_event(Event::End(BytesEnd::new("xdr:spPr")))?;

        writer.write_event(Event::End(BytesEnd::new("xdr:pic")))?;
        writer.write_event(Event::Empty(BytesStart::new("xdr:clientData")))?;
        writer.write_event(Event::End(BytesEnd::new(anchor_tag)))?;
    }

    // Chart anchors (after image anchors so NV IDs don't collide).
    let nv_id_offset = images.len();
    for (index, (chart, rel_id)) in charts.iter().zip(chart_rel_ids.iter()).enumerate() {
        writer.write_event(Event::Start(BytesStart::new("xdr:twoCellAnchor")))?;

        // From anchor
        let from_col = chart.from_col().to_string();
        let from_row = chart.from_row().to_string();
        writer.write_event(Event::Start(BytesStart::new("xdr:from")))?;
        writer.write_event(Event::Start(BytesStart::new("xdr:col")))?;
        writer.write_event(Event::Text(BytesText::new(from_col.as_str())))?;
        writer.write_event(Event::End(BytesEnd::new("xdr:col")))?;
        writer.write_event(Event::Start(BytesStart::new("xdr:colOff")))?;
        writer.write_event(Event::Text(BytesText::new("0")))?;
        writer.write_event(Event::End(BytesEnd::new("xdr:colOff")))?;
        writer.write_event(Event::Start(BytesStart::new("xdr:row")))?;
        writer.write_event(Event::Text(BytesText::new(from_row.as_str())))?;
        writer.write_event(Event::End(BytesEnd::new("xdr:row")))?;
        writer.write_event(Event::Start(BytesStart::new("xdr:rowOff")))?;
        writer.write_event(Event::Text(BytesText::new("0")))?;
        writer.write_event(Event::End(BytesEnd::new("xdr:rowOff")))?;
        writer.write_event(Event::End(BytesEnd::new("xdr:from")))?;

        // To anchor
        let to_col = chart.to_col().to_string();
        let to_row = chart.to_row().to_string();
        writer.write_event(Event::Start(BytesStart::new("xdr:to")))?;
        writer.write_event(Event::Start(BytesStart::new("xdr:col")))?;
        writer.write_event(Event::Text(BytesText::new(to_col.as_str())))?;
        writer.write_event(Event::End(BytesEnd::new("xdr:col")))?;
        writer.write_event(Event::Start(BytesStart::new("xdr:colOff")))?;
        writer.write_event(Event::Text(BytesText::new("0")))?;
        writer.write_event(Event::End(BytesEnd::new("xdr:colOff")))?;
        writer.write_event(Event::Start(BytesStart::new("xdr:row")))?;
        writer.write_event(Event::Text(BytesText::new(to_row.as_str())))?;
        writer.write_event(Event::End(BytesEnd::new("xdr:row")))?;
        writer.write_event(Event::Start(BytesStart::new("xdr:rowOff")))?;
        writer.write_event(Event::Text(BytesText::new("0")))?;
        writer.write_event(Event::End(BytesEnd::new("xdr:rowOff")))?;
        writer.write_event(Event::End(BytesEnd::new("xdr:to")))?;

        // Graphic frame
        writer.write_event(Event::Start(BytesStart::new("xdr:graphicFrame")))?;

        // Non-visual properties
        writer.write_event(Event::Start(BytesStart::new("xdr:nvGraphicFramePr")))?;
        let mut c_nv_pr = BytesStart::new("xdr:cNvPr");
        let chart_nv_id = u32::try_from(nv_id_offset + index + 1).map_err(|_| {
            XlsxError::InvalidWorkbookState("too many drawing objects to serialize".to_string())
        })?;
        let chart_nv_id_text = chart_nv_id.to_string();
        let chart_nv_name = chart
            .name()
            .map(|n| n.to_string())
            .unwrap_or_else(|| format!("Chart {}", index + 1));
        c_nv_pr.push_attribute(("id", chart_nv_id_text.as_str()));
        c_nv_pr.push_attribute(("name", chart_nv_name.as_str()));
        writer.write_event(Event::Empty(c_nv_pr))?;
        writer.write_event(Event::Empty(BytesStart::new("xdr:cNvGraphicFramePr")))?;
        writer.write_event(Event::End(BytesEnd::new("xdr:nvGraphicFramePr")))?;

        // Transform
        writer.write_event(Event::Start(BytesStart::new("xdr:xfrm")))?;
        let mut off = BytesStart::new("a:off");
        off.push_attribute(("x", "0"));
        off.push_attribute(("y", "0"));
        writer.write_event(Event::Empty(off))?;
        let mut ext = BytesStart::new("a:ext");
        ext.push_attribute(("cx", "0"));
        ext.push_attribute(("cy", "0"));
        writer.write_event(Event::Empty(ext))?;
        writer.write_event(Event::End(BytesEnd::new("xdr:xfrm")))?;

        // Graphic
        writer.write_event(Event::Start(BytesStart::new("a:graphic")))?;
        let mut graphic_data = BytesStart::new("a:graphicData");
        graphic_data.push_attribute(("uri", CHART_NS));
        writer.write_event(Event::Start(graphic_data))?;
        let mut chart_ref_tag = BytesStart::new("c:chart");
        chart_ref_tag.push_attribute(("r:id", rel_id.as_str()));
        writer.write_event(Event::Empty(chart_ref_tag))?;
        writer.write_event(Event::End(BytesEnd::new("a:graphicData")))?;
        writer.write_event(Event::End(BytesEnd::new("a:graphic")))?;

        writer.write_event(Event::End(BytesEnd::new("xdr:graphicFrame")))?;
        writer.write_event(Event::Empty(BytesStart::new("xdr:clientData")))?;
        writer.write_event(Event::End(BytesEnd::new("xdr:twoCellAnchor")))?;
    }

    writer.write_event(Event::End(BytesEnd::new("xdr:wsDr")))?;
    Ok(writer.into_inner())
}

/// Writes a cell anchor element (`<xdr:from>` or `<xdr:to>`).
fn serialize_cell_anchor_element(
    writer: &mut Writer<Vec<u8>>,
    tag_name: &str,
    anchor_data: Option<(u32, u32, i64, i64)>,
    default_col: u32,
    default_row: u32,
) -> Result<()> {
    let (col, row, col_off, row_off) = anchor_data.unwrap_or((default_col, default_row, 0, 0));
    writer.write_event(Event::Start(BytesStart::new(tag_name)))?;
    writer.write_event(Event::Start(BytesStart::new("xdr:col")))?;
    let col_text = col.to_string();
    writer.write_event(Event::Text(BytesText::new(col_text.as_str())))?;
    writer.write_event(Event::End(BytesEnd::new("xdr:col")))?;
    writer.write_event(Event::Start(BytesStart::new("xdr:colOff")))?;
    let col_off_text = col_off.to_string();
    writer.write_event(Event::Text(BytesText::new(col_off_text.as_str())))?;
    writer.write_event(Event::End(BytesEnd::new("xdr:colOff")))?;
    writer.write_event(Event::Start(BytesStart::new("xdr:row")))?;
    let row_text = row.to_string();
    writer.write_event(Event::Text(BytesText::new(row_text.as_str())))?;
    writer.write_event(Event::End(BytesEnd::new("xdr:row")))?;
    writer.write_event(Event::Start(BytesStart::new("xdr:rowOff")))?;
    let row_off_text = row_off.to_string();
    writer.write_event(Event::Text(BytesText::new(row_off_text.as_str())))?;
    writer.write_event(Event::End(BytesEnd::new("xdr:rowOff")))?;
    writer.write_event(Event::End(BytesEnd::new(tag_name)))?;
    Ok(())
}

fn serialize_table_xml(table: &WorksheetTable, table_id: u32) -> Result<Vec<u8>> {
    let mut writer = Writer::new_with_indent(Vec::new(), b' ', 2);
    writer.write_event(Event::Decl(BytesDecl::new(
        "1.0",
        Some("UTF-8"),
        Some("yes"),
    )))?;

    let mut table_tag = BytesStart::new("table");
    table_tag.push_attribute(("xmlns", SPREADSHEETML_NS));
    let table_id_text = table_id.to_string();
    table_tag.push_attribute(("id", table_id_text.as_str()));
    table_tag.push_attribute(("name", table.name()));
    table_tag.push_attribute(("displayName", table.name()));
    let table_range = table.range_xml();
    table_tag.push_attribute(("ref", table_range.as_str()));
    table_tag.push_attribute((
        "headerRowCount",
        if table.has_header_row() { "1" } else { "0" },
    ));
    if let Some(totals) = table.totals_row_shown() {
        table_tag.push_attribute(("totalsRowShown", if totals { "1" } else { "0" }));
    }
    // Replay unknown attributes from the original <table> element.
    for (key, val) in table.unknown_table_attrs() {
        table_tag.push_attribute((key.as_str(), val.as_str()));
    }
    writer.write_event(Event::Start(table_tag))?;

    if table.has_header_row() {
        let mut auto_filter_tag = BytesStart::new("autoFilter");
        auto_filter_tag.push_attribute(("ref", table_range.as_str()));
        writer.write_event(Event::Empty(auto_filter_tag))?;
    }

    let mut table_columns_tag = BytesStart::new("tableColumns");
    if table.columns().is_empty() {
        let column_count = table.range().width().to_string();
        table_columns_tag.push_attribute(("count", column_count.as_str()));
        writer.write_event(Event::Start(table_columns_tag))?;
        for column_index in 1..=table.range().width() {
            let mut table_column = BytesStart::new("tableColumn");
            let column_id = column_index.to_string();
            let column_name = format!("Column{column_index}");
            table_column.push_attribute(("id", column_id.as_str()));
            table_column.push_attribute(("name", column_name.as_str()));
            writer.write_event(Event::Empty(table_column))?;
        }
    } else {
        let column_count = table.columns().len().to_string();
        table_columns_tag.push_attribute(("count", column_count.as_str()));
        writer.write_event(Event::Start(table_columns_tag))?;
        for col in table.columns() {
            let mut tc_tag = BytesStart::new("tableColumn");
            let column_id = col.id().to_string();
            tc_tag.push_attribute(("id", column_id.as_str()));
            tc_tag.push_attribute(("name", col.name()));
            if let Some(label) = col.totals_row_label() {
                tc_tag.push_attribute(("totalsRowLabel", label));
            }
            if let Some(func) = col.totals_row_function() {
                tc_tag.push_attribute(("totalsRowFunction", func.as_xml_value()));
            }
            // Replay unknown attributes from the original <tableColumn> element.
            for (key, val) in col.unknown_attrs() {
                tc_tag.push_attribute((key.as_str(), val.as_str()));
            }
            if let Some(formula) = col.totals_row_formula() {
                writer.write_event(Event::Start(tc_tag))?;
                writer.write_event(Event::Start(BytesStart::new("totalsRowFormula")))?;
                writer.write_event(Event::Text(BytesText::new(formula)))?;
                writer.write_event(Event::End(BytesEnd::new("totalsRowFormula")))?;
                writer.write_event(Event::End(BytesEnd::new("tableColumn")))?;
            } else {
                writer.write_event(Event::Empty(tc_tag))?;
            }
        }
    }
    writer.write_event(Event::End(BytesEnd::new("tableColumns")))?;

    let mut table_style_info = BytesStart::new("tableStyleInfo");
    let style_name = table.style_name().unwrap_or("TableStyleMedium2");
    table_style_info.push_attribute(("name", style_name));
    let show_first = table.show_first_column().unwrap_or(false);
    table_style_info.push_attribute(("showFirstColumn", if show_first { "1" } else { "0" }));
    let show_last = table.show_last_column().unwrap_or(false);
    table_style_info.push_attribute(("showLastColumn", if show_last { "1" } else { "0" }));
    let show_row_stripes = table.show_row_stripes().unwrap_or(true);
    table_style_info.push_attribute(("showRowStripes", if show_row_stripes { "1" } else { "0" }));
    let show_col_stripes = table.show_column_stripes().unwrap_or(false);
    table_style_info.push_attribute((
        "showColumnStripes",
        if show_col_stripes { "1" } else { "0" },
    ));
    // Replay unknown attributes from the original <tableStyleInfo> element.
    for (key, val) in table.unknown_style_attrs() {
        table_style_info.push_attribute((key.as_str(), val.as_str()));
    }
    writer.write_event(Event::Empty(table_style_info))?;

    writer.write_event(Event::End(BytesEnd::new("table")))?;
    Ok(writer.into_inner())
}

fn serialize_shared_strings_xml(shared_strings: &SharedStrings) -> Result<Vec<u8>> {
    let mut writer = Writer::new_with_indent(Vec::new(), b' ', 2);
    writer.write_event(Event::Decl(BytesDecl::new(
        "1.0",
        Some("UTF-8"),
        Some("yes"),
    )))?;

    let mut sst = BytesStart::new("sst");
    sst.push_attribute(("xmlns", SPREADSHEETML_NS));
    let unique_count = shared_strings.len().to_string();
    sst.push_attribute(("count", unique_count.as_str()));
    sst.push_attribute(("uniqueCount", unique_count.as_str()));
    writer.write_event(Event::Start(sst))?;

    for index in 0..shared_strings.len() {
        let entry = shared_strings.get_entry(index).ok_or_else(|| {
            XlsxError::InvalidWorkbookState(format!(
                "missing shared string entry at index `{index}`"
            ))
        })?;

        writer.write_event(Event::Start(BytesStart::new("si")))?;

        match entry {
            SharedStringEntry::Plain(value) => {
                let mut t_tag = BytesStart::new("t");
                // Preserve leading/trailing whitespace with xml:space="preserve"
                if value.starts_with(|c: char| c.is_ascii_whitespace())
                    || value.ends_with(|c: char| c.is_ascii_whitespace())
                {
                    t_tag.push_attribute(("xml:space", "preserve"));
                }
                writer.write_event(Event::Start(t_tag))?;
                writer.write_event(Event::Text(BytesText::new(value)))?;
                writer.write_event(Event::End(BytesEnd::new("t")))?;
            }
            SharedStringEntry::RichText {
                runs,
                phonetic_runs,
                phonetic_pr,
            } => {
                for run in runs {
                    writer.write_event(Event::Start(BytesStart::new("r")))?;

                    // Write <rPr> formatting properties if any exist
                    if run.has_formatting() {
                        writer.write_event(Event::Start(BytesStart::new("rPr")))?;

                        if let Some(bold) = run.bold() {
                            if bold {
                                writer.write_event(Event::Empty(BytesStart::new("b")))?;
                            } else {
                                let mut tag = BytesStart::new("b");
                                tag.push_attribute(("val", "0"));
                                writer.write_event(Event::Empty(tag))?;
                            }
                        }

                        if let Some(italic) = run.italic() {
                            if italic {
                                writer.write_event(Event::Empty(BytesStart::new("i")))?;
                            } else {
                                let mut tag = BytesStart::new("i");
                                tag.push_attribute(("val", "0"));
                                writer.write_event(Event::Empty(tag))?;
                            }
                        }

                        // Write unknown rPr children (underline, strikethrough, charset, etc.)
                        for node in run.unknown_rpr_children() {
                            node.write_to(&mut writer)?;
                        }

                        if let Some(size) = run.font_size() {
                            let mut tag = BytesStart::new("sz");
                            tag.push_attribute(("val", size));
                            writer.write_event(Event::Empty(tag))?;
                        }

                        if let Some(color) = run.color() {
                            let mut tag = BytesStart::new("color");
                            tag.push_attribute(("rgb", color));
                            writer.write_event(Event::Empty(tag))?;
                        }

                        if let Some(name) = run.font_name() {
                            let mut tag = BytesStart::new("rFont");
                            tag.push_attribute(("val", name));
                            writer.write_event(Event::Empty(tag))?;
                        }

                        writer.write_event(Event::End(BytesEnd::new("rPr")))?;
                    }

                    // Write <t> element with text
                    let text = run.text();
                    let mut t_tag = BytesStart::new("t");
                    if text.starts_with(|c: char| c.is_ascii_whitespace())
                        || text.ends_with(|c: char| c.is_ascii_whitespace())
                    {
                        t_tag.push_attribute(("xml:space", "preserve"));
                    }
                    writer.write_event(Event::Start(t_tag))?;
                    writer.write_event(Event::Text(BytesText::new(text)))?;
                    writer.write_event(Event::End(BytesEnd::new("t")))?;

                    writer.write_event(Event::End(BytesEnd::new("r")))?;
                }

                // Write phonetic runs (<rPh>) for CJK roundtrip
                for ph_run in phonetic_runs {
                    ph_run.write_to(&mut writer)?;
                }

                // Write phonetic properties (<phoneticPr>)
                if let Some(ph_pr) = phonetic_pr {
                    ph_pr.write_to(&mut writer)?;
                }
            }
        }

        writer.write_event(Event::End(BytesEnd::new("si")))?;
    }

    writer.write_event(Event::End(BytesEnd::new("sst")))?;
    Ok(writer.into_inner())
}

fn serialize_styles_xml(
    styles: &[Style],
    cell_style_xfs: &[CellStyleXf],
    named_styles: &[NamedStyle],
) -> Result<Vec<u8>> {
    #[derive(Debug, Clone, Copy)]
    struct SerializedXf {
        num_fmt_id: u32,
        font_id: u32,
        fill_id: u32,
        border_id: u32,
    }

    fn style_table_id_for<T>(table: &mut Vec<T>, value: &T) -> Result<u32>
    where
        T: Clone + PartialEq,
    {
        if let Some(index) = table.iter().position(|entry| entry == value) {
            return u32::try_from(index).map_err(|_| {
                XlsxError::InvalidWorkbookState(
                    "style table index cannot be represented as u32".to_string(),
                )
            });
        }

        table.push(value.clone());
        u32::try_from(table.len() - 1).map_err(|_| {
            XlsxError::InvalidWorkbookState(
                "style table index cannot be represented as u32".to_string(),
            )
        })
    }

    let mut writer = Writer::new_with_indent(Vec::new(), b' ', 2);
    writer.write_event(Event::Decl(BytesDecl::new(
        "1.0",
        Some("UTF-8"),
        Some("yes"),
    )))?;

    let mut stylesheet = BytesStart::new("styleSheet");
    stylesheet.push_attribute(("xmlns", SPREADSHEETML_NS));
    writer.write_event(Event::Start(stylesheet))?;

    let mut number_formats = Vec::<(u32, String)>::new();
    let mut number_format_ids = BTreeMap::<String, u32>::new();
    let mut fonts = vec![Font::new()];
    let mut fills = vec![Fill::new()];
    let mut borders = vec![Border::new()];
    let mut serialized_xfs = Vec::<SerializedXf>::with_capacity(styles.len());

    for style in styles {
        let num_fmt_id = if let Some(number_format) = style.number_format() {
            if let Some(existing_num_fmt_id) = number_format_ids.get(number_format).copied() {
                existing_num_fmt_id
            } else {
                let next_num_fmt_id = 164_u32
                    .checked_add(u32::try_from(number_formats.len()).map_err(|_| {
                        XlsxError::InvalidWorkbookState(
                            "too many custom number formats to serialize".to_string(),
                        )
                    })?)
                    .ok_or_else(|| {
                        XlsxError::InvalidWorkbookState(
                            "custom number format id overflow".to_string(),
                        )
                    })?;
                number_formats.push((next_num_fmt_id, number_format.to_string()));
                number_format_ids.insert(number_format.to_string(), next_num_fmt_id);
                next_num_fmt_id
            }
        } else {
            0
        };

        let font_id = match style.font().filter(|font| font.has_metadata()) {
            Some(font) => style_table_id_for(&mut fonts, font)?,
            None => 0,
        };

        let fill_id = match style.fill().filter(|fill| fill.has_metadata()) {
            Some(fill) => style_table_id_for(&mut fills, fill)?,
            None => 0,
        };

        let border_id = match style.border().filter(|border| border.has_metadata()) {
            Some(border) => style_table_id_for(&mut borders, border)?,
            None => 0,
        };

        serialized_xfs.push(SerializedXf {
            num_fmt_id,
            font_id,
            fill_id,
            border_id,
        });
    }

    if !number_formats.is_empty() {
        let mut num_fmts = BytesStart::new("numFmts");
        let count = number_formats.len().to_string();
        num_fmts.push_attribute(("count", count.as_str()));
        writer.write_event(Event::Start(num_fmts))?;
        for (num_fmt_id, format_code) in &number_formats {
            let mut num_fmt = BytesStart::new("numFmt");
            let num_fmt_id_text = num_fmt_id.to_string();
            num_fmt.push_attribute(("numFmtId", num_fmt_id_text.as_str()));
            num_fmt.push_attribute(("formatCode", format_code.as_str()));
            writer.write_event(Event::Empty(num_fmt))?;
        }
        writer.write_event(Event::End(BytesEnd::new("numFmts")))?;
    }

    writer.write_event(Event::Start({
        let mut fonts_tag = BytesStart::new("fonts");
        let count = fonts.len().to_string();
        fonts_tag.push_attribute(("count", count.as_str()));
        fonts_tag
    }))?;
    for font in &fonts {
        if !font.has_metadata() {
            writer.write_event(Event::Empty(BytesStart::new("font")))?;
            continue;
        }

        writer.write_event(Event::Start(BytesStart::new("font")))?;
        if let Some(name) = font.name() {
            let mut name_tag = BytesStart::new("name");
            name_tag.push_attribute(("val", name));
            writer.write_event(Event::Empty(name_tag))?;
        }
        if let Some(size) = font.size() {
            let mut size_tag = BytesStart::new("sz");
            size_tag.push_attribute(("val", size));
            writer.write_event(Event::Empty(size_tag))?;
        }
        if let Some(bold) = font.bold() {
            let mut bold_tag = BytesStart::new("b");
            if !bold {
                bold_tag.push_attribute(("val", "0"));
            }
            writer.write_event(Event::Empty(bold_tag))?;
        }
        if let Some(italic) = font.italic() {
            let mut italic_tag = BytesStart::new("i");
            if !italic {
                italic_tag.push_attribute(("val", "0"));
            }
            writer.write_event(Event::Empty(italic_tag))?;
        }
        if let Some(underline) = font.underline() {
            let mut underline_tag = BytesStart::new("u");
            if !underline {
                underline_tag.push_attribute(("val", "0"));
            }
            writer.write_event(Event::Empty(underline_tag))?;
        }
        if let Some(strike) = font.strikethrough() {
            let mut strike_tag = BytesStart::new("strike");
            if !strike {
                strike_tag.push_attribute(("val", "0"));
            }
            writer.write_event(Event::Empty(strike_tag))?;
        }
        if let Some(dstrike) = font.double_strikethrough() {
            let mut dstrike_tag = BytesStart::new("dStrike");
            if !dstrike {
                dstrike_tag.push_attribute(("val", "0"));
            }
            writer.write_event(Event::Empty(dstrike_tag))?;
        }
        if let Some(shadow) = font.shadow() {
            let mut shadow_tag = BytesStart::new("shadow");
            if !shadow {
                shadow_tag.push_attribute(("val", "0"));
            }
            writer.write_event(Event::Empty(shadow_tag))?;
        }
        if let Some(outline) = font.outline() {
            let mut outline_tag = BytesStart::new("outline");
            if !outline {
                outline_tag.push_attribute(("val", "0"));
            }
            writer.write_event(Event::Empty(outline_tag))?;
        }
        if let Some(va) = font.vertical_align() {
            let mut vert_align_tag = BytesStart::new("vertAlign");
            vert_align_tag.push_attribute(("val", va.as_xml_value()));
            writer.write_event(Event::Empty(vert_align_tag))?;
        } else {
            // Fallback: write legacy subscript/superscript fields if vertical_align is not set.
            if font.subscript() == Some(true) {
                let mut vert_align_tag = BytesStart::new("vertAlign");
                vert_align_tag.push_attribute(("val", "subscript"));
                writer.write_event(Event::Empty(vert_align_tag))?;
            }
            if font.superscript() == Some(true) {
                let mut vert_align_tag = BytesStart::new("vertAlign");
                vert_align_tag.push_attribute(("val", "superscript"));
                writer.write_event(Event::Empty(vert_align_tag))?;
            }
        }
        if let Some(scheme) = font.font_scheme() {
            let mut scheme_tag = BytesStart::new("scheme");
            scheme_tag.push_attribute(("val", scheme.as_xml_value()));
            writer.write_event(Event::Empty(scheme_tag))?;
        }
        if let Some(color_ref) = font.color_ref() {
            let mut color_tag = BytesStart::new("color");
            if let Some(rgb) = color_ref.rgb() {
                color_tag.push_attribute(("rgb", rgb));
            }
            if let Some(theme) = color_ref.theme() {
                let idx = theme.index().to_string();
                color_tag.push_attribute(("theme", idx.as_str()));
            }
            if let Some(tint) = color_ref.tint() {
                let tint_str = tint.to_string();
                color_tag.push_attribute(("tint", tint_str.as_str()));
            }
            if let Some(indexed) = color_ref.indexed() {
                let idx_str = indexed.to_string();
                color_tag.push_attribute(("indexed", idx_str.as_str()));
            }
            if color_ref.auto() == Some(true) {
                color_tag.push_attribute(("auto", "1"));
            }
            writer.write_event(Event::Empty(color_tag))?;
        } else if let Some(color) = font.color() {
            let mut color_tag = BytesStart::new("color");
            color_tag.push_attribute(("rgb", color));
            writer.write_event(Event::Empty(color_tag))?;
        }
        writer.write_event(Event::End(BytesEnd::new("font")))?;
    }
    writer.write_event(Event::End(BytesEnd::new("fonts")))?;

    writer.write_event(Event::Start({
        let mut fills_tag = BytesStart::new("fills");
        let count = fills.len().to_string();
        fills_tag.push_attribute(("count", count.as_str()));
        fills_tag
    }))?;
    for fill in &fills {
        if !fill.has_metadata() {
            writer.write_event(Event::Empty(BytesStart::new("fill")))?;
            continue;
        }

        writer.write_event(Event::Start(BytesStart::new("fill")))?;
        let mut pattern_fill_tag = BytesStart::new("patternFill");

        // Use the string-based pattern if available, otherwise fall back to PatternFill.
        if let Some(pattern) = fill.pattern() {
            pattern_fill_tag.push_attribute(("patternType", pattern));
        } else if let Some(pf) = fill.pattern_fill() {
            pattern_fill_tag.push_attribute(("patternType", pf.pattern_type().as_xml_value()));
        }

        writer.write_event(Event::Start(pattern_fill_tag))?;

        // Serialize foreground color: prefer PatternFill (has full ColorReference),
        // fall back to string-based.
        if let Some(fg) = fill.pattern_fill().and_then(|pf| pf.fg_color()) {
            let mut fg_tag = BytesStart::new("fgColor");
            if let Some(rgb) = fg.rgb() {
                fg_tag.push_attribute(("rgb", rgb));
            }
            if let Some(theme) = fg.theme() {
                let idx = theme.index().to_string();
                fg_tag.push_attribute(("theme", idx.as_str()));
            }
            if let Some(tint) = fg.tint() {
                let tint_str = tint.to_string();
                fg_tag.push_attribute(("tint", tint_str.as_str()));
            }
            if let Some(indexed) = fg.indexed() {
                let idx_str = indexed.to_string();
                fg_tag.push_attribute(("indexed", idx_str.as_str()));
            }
            if fg.auto() == Some(true) {
                fg_tag.push_attribute(("auto", "1"));
            }
            writer.write_event(Event::Empty(fg_tag))?;
        } else if let Some(color) = fill.foreground_color() {
            let mut fg_color = BytesStart::new("fgColor");
            fg_color.push_attribute(("rgb", color));
            writer.write_event(Event::Empty(fg_color))?;
        }

        // Serialize background color: prefer PatternFill (has full ColorReference),
        // fall back to string-based.
        if let Some(bg) = fill.pattern_fill().and_then(|pf| pf.bg_color()) {
            let mut bg_tag = BytesStart::new("bgColor");
            if let Some(rgb) = bg.rgb() {
                bg_tag.push_attribute(("rgb", rgb));
            }
            if let Some(theme) = bg.theme() {
                let idx = theme.index().to_string();
                bg_tag.push_attribute(("theme", idx.as_str()));
            }
            if let Some(tint) = bg.tint() {
                let tint_str = tint.to_string();
                bg_tag.push_attribute(("tint", tint_str.as_str()));
            }
            if let Some(indexed) = bg.indexed() {
                let idx_str = indexed.to_string();
                bg_tag.push_attribute(("indexed", idx_str.as_str()));
            }
            if bg.auto() == Some(true) {
                bg_tag.push_attribute(("auto", "1"));
            }
            writer.write_event(Event::Empty(bg_tag))?;
        } else if let Some(color) = fill.background_color() {
            let mut bg_color = BytesStart::new("bgColor");
            bg_color.push_attribute(("rgb", color));
            writer.write_event(Event::Empty(bg_color))?;
        }

        writer.write_event(Event::End(BytesEnd::new("patternFill")))?;
        writer.write_event(Event::End(BytesEnd::new("fill")))?;
    }
    writer.write_event(Event::End(BytesEnd::new("fills")))?;

    writer.write_event(Event::Start({
        let mut borders_tag = BytesStart::new("borders");
        let count = borders.len().to_string();
        borders_tag.push_attribute(("count", count.as_str()));
        borders_tag
    }))?;
    for border in &borders {
        if !border.has_metadata() {
            writer.write_event(Event::Empty(BytesStart::new("border")))?;
            continue;
        }

        let mut border_tag = BytesStart::new("border");
        if let Some(diag_up) = border.diagonal_up() {
            border_tag.push_attribute(("diagonalUp", if diag_up { "1" } else { "0" }));
        }
        if let Some(diag_down) = border.diagonal_down() {
            border_tag.push_attribute(("diagonalDown", if diag_down { "1" } else { "0" }));
        }
        writer.write_event(Event::Start(border_tag))?;
        for (name, side) in [
            ("left", border.left()),
            ("right", border.right()),
            ("top", border.top()),
            ("bottom", border.bottom()),
            ("diagonal", border.diagonal()),
        ] {
            match side {
                Some(side) if side.has_metadata() => {
                    let mut side_tag = BytesStart::new(name);
                    if let Some(style) = side.style() {
                        side_tag.push_attribute(("style", style));
                    }

                    if let Some(color_ref) = side.color_ref() {
                        writer.write_event(Event::Start(side_tag))?;
                        let mut color_tag = BytesStart::new("color");
                        if let Some(rgb) = color_ref.rgb() {
                            color_tag.push_attribute(("rgb", rgb));
                        }
                        if let Some(theme) = color_ref.theme() {
                            let idx = theme.index().to_string();
                            color_tag.push_attribute(("theme", idx.as_str()));
                        }
                        if let Some(tint) = color_ref.tint() {
                            let tint_str = tint.to_string();
                            color_tag.push_attribute(("tint", tint_str.as_str()));
                        }
                        if let Some(indexed) = color_ref.indexed() {
                            let idx_str = indexed.to_string();
                            color_tag.push_attribute(("indexed", idx_str.as_str()));
                        }
                        if color_ref.auto() == Some(true) {
                            color_tag.push_attribute(("auto", "1"));
                        }
                        writer.write_event(Event::Empty(color_tag))?;
                        writer.write_event(Event::End(BytesEnd::new(name)))?;
                    } else if let Some(color) = side.color() {
                        writer.write_event(Event::Start(side_tag))?;
                        let mut color_tag = BytesStart::new("color");
                        color_tag.push_attribute(("rgb", color));
                        writer.write_event(Event::Empty(color_tag))?;
                        writer.write_event(Event::End(BytesEnd::new(name)))?;
                    } else {
                        writer.write_event(Event::Empty(side_tag))?;
                    }
                }
                _ => {
                    writer.write_event(Event::Empty(BytesStart::new(name)))?;
                }
            }
        }
        writer.write_event(Event::End(BytesEnd::new("border")))?;
    }
    writer.write_event(Event::End(BytesEnd::new("borders")))?;

    {
        let xf_entries: Vec<&CellStyleXf> = if cell_style_xfs.is_empty() {
            // Provide a default entry when none exist
            Vec::new()
        } else {
            cell_style_xfs.iter().collect()
        };
        let count = if xf_entries.is_empty() {
            1
        } else {
            xf_entries.len()
        };
        let count_text = count.to_string();
        let mut csxfs_tag = BytesStart::new("cellStyleXfs");
        csxfs_tag.push_attribute(("count", count_text.as_str()));
        writer.write_event(Event::Start(csxfs_tag))?;
        if xf_entries.is_empty() {
            writer.write_event(Event::Empty({
                let mut xf = BytesStart::new("xf");
                xf.push_attribute(("numFmtId", "0"));
                xf.push_attribute(("fontId", "0"));
                xf.push_attribute(("fillId", "0"));
                xf.push_attribute(("borderId", "0"));
                xf
            }))?;
        } else {
            for csxf in &xf_entries {
                let mut xf = BytesStart::new("xf");
                let nfid = csxf.num_fmt_id().unwrap_or(0).to_string();
                let fid = csxf.font_id().unwrap_or(0).to_string();
                let flid = csxf.fill_id().unwrap_or(0).to_string();
                let bid = csxf.border_id().unwrap_or(0).to_string();
                xf.push_attribute(("numFmtId", nfid.as_str()));
                xf.push_attribute(("fontId", fid.as_str()));
                xf.push_attribute(("fillId", flid.as_str()));
                xf.push_attribute(("borderId", bid.as_str()));
                writer.write_event(Event::Empty(xf))?;
            }
        }
        writer.write_event(Event::End(BytesEnd::new("cellStyleXfs")))?;
    }

    let cell_xfs_count = styles.len().to_string();
    let mut cell_xfs = BytesStart::new("cellXfs");
    cell_xfs.push_attribute(("count", cell_xfs_count.as_str()));
    writer.write_event(Event::Start(cell_xfs))?;
    for (style, serialized_xf) in styles.iter().zip(serialized_xfs.iter()) {
        let num_fmt_id_text = serialized_xf.num_fmt_id.to_string();
        let font_id_text = serialized_xf.font_id.to_string();
        let fill_id_text = serialized_xf.fill_id.to_string();
        let border_id_text = serialized_xf.border_id.to_string();

        let mut xf = BytesStart::new("xf");
        xf.push_attribute(("numFmtId", num_fmt_id_text.as_str()));
        xf.push_attribute(("fontId", font_id_text.as_str()));
        xf.push_attribute(("fillId", fill_id_text.as_str()));
        xf.push_attribute(("borderId", border_id_text.as_str()));
        xf.push_attribute(("xfId", "0"));

        if style.number_format().is_some() {
            xf.push_attribute(("applyNumberFormat", "1"));
        }
        if serialized_xf.font_id > 0 {
            xf.push_attribute(("applyFont", "1"));
        }
        if serialized_xf.fill_id > 0 {
            xf.push_attribute(("applyFill", "1"));
        }
        if serialized_xf.border_id > 0 {
            xf.push_attribute(("applyBorder", "1"));
        }

        let has_alignment = style.alignment().is_some_and(|a| a.has_metadata());
        let has_protection = style.protection().is_some_and(|p| p.has_metadata());

        if has_alignment {
            xf.push_attribute(("applyAlignment", "1"));
        }
        if has_protection {
            xf.push_attribute(("applyProtection", "1"));
        }

        if has_alignment || has_protection {
            writer.write_event(Event::Start(xf))?;

            if let Some(alignment) = style.alignment().filter(|a| a.has_metadata()) {
                let mut alignment_tag = BytesStart::new("alignment");
                let horizontal = alignment.horizontal().map(|value| value.as_xml_value());
                let vertical = alignment.vertical().map(|value| value.as_xml_value());
                let wrap_text = alignment.wrap_text();
                if let Some(horizontal) = horizontal {
                    alignment_tag.push_attribute(("horizontal", horizontal));
                }
                if let Some(vertical) = vertical {
                    alignment_tag.push_attribute(("vertical", vertical));
                }
                if let Some(wrap_text) = wrap_text {
                    alignment_tag.push_attribute(("wrapText", if wrap_text { "1" } else { "0" }));
                }
                let indent_text = alignment.indent().map(|v| v.to_string());
                if let Some(indent) = indent_text.as_deref() {
                    alignment_tag.push_attribute(("indent", indent));
                }
                let rotation_text = alignment.text_rotation().map(|v| v.to_string());
                if let Some(rotation) = rotation_text.as_deref() {
                    alignment_tag.push_attribute(("textRotation", rotation));
                }
                if let Some(shrink) = alignment.shrink_to_fit() {
                    alignment_tag.push_attribute(("shrinkToFit", if shrink { "1" } else { "0" }));
                }
                let reading_order_text = alignment.reading_order().map(|v| v.to_string());
                if let Some(reading_order) = reading_order_text.as_deref() {
                    alignment_tag.push_attribute(("readingOrder", reading_order));
                }
                writer.write_event(Event::Empty(alignment_tag))?;
            }

            if let Some(protection) = style.protection().filter(|p| p.has_metadata()) {
                let mut protection_tag = BytesStart::new("protection");
                if let Some(locked) = protection.locked() {
                    protection_tag.push_attribute(("locked", if locked { "1" } else { "0" }));
                }
                if let Some(hidden) = protection.hidden() {
                    protection_tag.push_attribute(("hidden", if hidden { "1" } else { "0" }));
                }
                writer.write_event(Event::Empty(protection_tag))?;
            }

            writer.write_event(Event::End(BytesEnd::new("xf")))?;
        } else {
            writer.write_event(Event::Empty(xf))?;
        }
    }
    writer.write_event(Event::End(BytesEnd::new("cellXfs")))?;

    {
        let ns_entries: &[NamedStyle] = if named_styles.is_empty() {
            &[]
        } else {
            named_styles
        };
        let count = if ns_entries.is_empty() {
            1
        } else {
            ns_entries.len()
        };
        let count_text = count.to_string();
        let mut cs_tag = BytesStart::new("cellStyles");
        cs_tag.push_attribute(("count", count_text.as_str()));
        writer.write_event(Event::Start(cs_tag))?;
        if ns_entries.is_empty() {
            writer.write_event(Event::Empty({
                let mut cell_style = BytesStart::new("cellStyle");
                cell_style.push_attribute(("name", "Normal"));
                cell_style.push_attribute(("xfId", "0"));
                cell_style.push_attribute(("builtinId", "0"));
                cell_style
            }))?;
        } else {
            for ns in ns_entries {
                let mut cell_style = BytesStart::new("cellStyle");
                cell_style.push_attribute(("name", ns.name()));
                let xf_id_text = ns.xf_id().to_string();
                cell_style.push_attribute(("xfId", xf_id_text.as_str()));
                if let Some(bid) = ns.builtin_id() {
                    let bid_text = bid.to_string();
                    cell_style.push_attribute(("builtinId", bid_text.as_str()));
                }
                writer.write_event(Event::Empty(cell_style))?;
            }
        }
        writer.write_event(Event::End(BytesEnd::new("cellStyles")))?;
    }

    writer.write_event(Event::End(BytesEnd::new("styleSheet")))?;
    Ok(writer.into_inner())
}

fn parse_freeze_pane(event: &BytesStart<'_>) -> Option<FreezePane> {
    let mut x_split = 0_u32;
    let mut y_split = 0_u32;
    let mut top_left_cell = None;
    let mut state = None;

    for attribute in event.attributes().flatten() {
        let key = local_name(attribute.key.as_ref());
        let value = String::from_utf8_lossy(attribute.value.as_ref()).into_owned();
        match key {
            b"xSplit" => x_split = parse_pane_split(value.as_str()).unwrap_or(0),
            b"ySplit" => y_split = parse_pane_split(value.as_str()).unwrap_or(0),
            b"topLeftCell" => top_left_cell = Some(value),
            b"state" => state = Some(value),
            _ => {}
        }
    }

    let state = state.as_deref().unwrap_or("");
    if !matches!(state, "frozen" | "frozenSplit") {
        return None;
    }

    if x_split == 0 && y_split == 0 {
        return None;
    }

    top_left_cell
        .as_deref()
        .and_then(|top_left_cell| {
            FreezePane::with_top_left_cell(x_split, y_split, top_left_cell).ok()
        })
        .or_else(|| FreezePane::new(x_split, y_split).ok())
}

fn parse_pane_split(raw: &str) -> Option<u32> {
    let trimmed = raw.trim();
    if trimmed.is_empty() {
        return None;
    }

    if let Ok(value) = trimmed.parse::<u32>() {
        return Some(value);
    }

    let parsed = trimmed.parse::<f64>().ok()?;
    if !parsed.is_finite() || parsed < 0.0 || parsed.fract() != 0.0 {
        return None;
    }

    // `as` is safe here after finite, non-negative, integral check.
    Some(parsed as u32)
}

fn parse_sheet_protection(event: &BytesStart<'_>) -> SheetProtection {
    let mut protection = SheetProtection::default();
    for attribute in event.attributes().flatten() {
        let key = local_name(attribute.key.as_ref());
        let value = String::from_utf8_lossy(attribute.value.as_ref()).into_owned();
        match key {
            b"password" | b"hashValue" => {
                let trimmed = value.trim().to_string();
                if !trimmed.is_empty() {
                    protection.password_hash = Some(trimmed);
                }
            }
            b"sheet" => protection.sheet = parse_xml_bool(value.as_str()).unwrap_or(false),
            b"objects" => protection.objects = parse_xml_bool(value.as_str()).unwrap_or(false),
            b"scenarios" => protection.scenarios = parse_xml_bool(value.as_str()).unwrap_or(false),
            b"formatCells" => {
                protection.format_cells = parse_xml_bool(value.as_str()).unwrap_or(false)
            }
            b"formatColumns" => {
                protection.format_columns = parse_xml_bool(value.as_str()).unwrap_or(false)
            }
            b"formatRows" => {
                protection.format_rows = parse_xml_bool(value.as_str()).unwrap_or(false)
            }
            b"insertColumns" => {
                protection.insert_columns = parse_xml_bool(value.as_str()).unwrap_or(false)
            }
            b"insertRows" => {
                protection.insert_rows = parse_xml_bool(value.as_str()).unwrap_or(false)
            }
            b"insertHyperlinks" => {
                protection.insert_hyperlinks = parse_xml_bool(value.as_str()).unwrap_or(false)
            }
            b"deleteColumns" => {
                protection.delete_columns = parse_xml_bool(value.as_str()).unwrap_or(false)
            }
            b"deleteRows" => {
                protection.delete_rows = parse_xml_bool(value.as_str()).unwrap_or(false)
            }
            b"selectLockedCells" => {
                protection.select_locked_cells = parse_xml_bool(value.as_str()).unwrap_or(false)
            }
            b"sort" => protection.sort = parse_xml_bool(value.as_str()).unwrap_or(false),
            b"autoFilter" => {
                protection.auto_filter = parse_xml_bool(value.as_str()).unwrap_or(false)
            }
            b"pivotTables" => {
                protection.pivot_tables = parse_xml_bool(value.as_str()).unwrap_or(false)
            }
            b"selectUnlockedCells" => {
                protection.select_unlocked_cells = parse_xml_bool(value.as_str()).unwrap_or(false)
            }
            _ => {}
        }
    }
    protection
}

fn parse_page_setup(event: &BytesStart<'_>) -> PageSetup {
    let mut page_setup = PageSetup::default();
    for attribute in event.attributes().flatten() {
        let key = local_name(attribute.key.as_ref());
        let value = String::from_utf8_lossy(attribute.value.as_ref()).into_owned();
        match key {
            b"orientation" => {
                page_setup.orientation = Some(PageOrientation::from_xml_value(value.as_str()));
            }
            b"paperSize" => {
                page_setup.paper_size = value.trim().parse::<u32>().ok();
            }
            b"scale" => {
                page_setup.scale = value.trim().parse::<u32>().ok();
            }
            b"fitToWidth" => {
                page_setup.fit_to_width = value.trim().parse::<u32>().ok();
            }
            b"fitToHeight" => {
                page_setup.fit_to_height = value.trim().parse::<u32>().ok();
            }
            b"firstPageNumber" => {
                page_setup.first_page_number = value.trim().parse::<u32>().ok();
            }
            _ => {}
        }
    }
    page_setup
}

/// Writes a color string as an attribute on a `<color>` tag.
/// Handles formats like "FFRRGGBB" (rgb), "theme:N", "indexed:N".
fn write_color_attribute(tag: &mut BytesStart<'_>, color: &str) {
    if let Some(theme_val) = color.strip_prefix("theme:") {
        tag.push_attribute(("theme", theme_val));
    } else if let Some(indexed_val) = color.strip_prefix("indexed:") {
        tag.push_attribute(("indexed", indexed_val));
    } else {
        tag.push_attribute(("rgb", color));
    }
}

fn parse_page_margins(event: &BytesStart<'_>) -> PageMargins {
    let mut margins = PageMargins::default();
    for attribute in event.attributes().flatten() {
        let key = local_name(attribute.key.as_ref());
        let value = String::from_utf8_lossy(attribute.value.as_ref()).into_owned();
        match key {
            b"left" => margins.left = value.trim().parse::<f64>().ok(),
            b"right" => margins.right = value.trim().parse::<f64>().ok(),
            b"top" => margins.top = value.trim().parse::<f64>().ok(),
            b"bottom" => margins.bottom = value.trim().parse::<f64>().ok(),
            b"header" => margins.header = value.trim().parse::<f64>().ok(),
            b"footer" => margins.footer = value.trim().parse::<f64>().ok(),
            _ => {}
        }
    }
    margins
}

fn parse_sheet_view_options(event: &BytesStart<'_>) -> SheetViewOptions {
    let mut options = SheetViewOptions::default();
    for attribute in event.attributes().flatten() {
        let key = local_name(attribute.key.as_ref());
        let value = String::from_utf8_lossy(attribute.value.as_ref()).into_owned();
        match key {
            b"showGridLines" => {
                options.show_gridlines = parse_xml_bool(value.as_str());
            }
            b"showRowColHeaders" => {
                options.show_row_col_headers = parse_xml_bool(value.as_str());
            }
            b"showFormulas" => {
                options.show_formulas = parse_xml_bool(value.as_str());
            }
            b"zoomScale" => {
                options.zoom_scale = value.trim().parse::<u32>().ok();
            }
            b"zoomScaleNormal" => {
                options.zoom_scale_normal = value.trim().parse::<u32>().ok();
            }
            b"rightToLeft" => {
                options.right_to_left = parse_xml_bool(value.as_str());
            }
            b"tabSelected" => {
                options.tab_selected = parse_xml_bool(value.as_str());
            }
            b"view" => {
                let trimmed = value.trim().to_string();
                if !trimmed.is_empty() {
                    options.view = Some(trimmed);
                }
            }
            _ => {}
        }
    }
    options
}

/// Load comments from the comments XML part associated with a worksheet.
fn load_worksheet_comments(
    worksheet: &mut Worksheet,
    package: &Package,
    worksheet_uri: &PartUri,
    worksheet_part: &Part,
) {
    let Some(relationship) = worksheet_part
        .relationships
        .get_first_by_type(COMMENTS_RELATIONSHIP_TYPE)
    else {
        return;
    };

    if relationship.target_mode != TargetMode::Internal {
        return;
    }

    let Ok(comments_uri) = worksheet_uri.resolve_relative(relationship.target.as_str()) else {
        return;
    };

    let Some(comments_part) = package.get_part(comments_uri.as_str()) else {
        return;
    };

    if let Ok(comments) = parse_comments_xml(comments_part.data.as_bytes()) {
        for comment in comments {
            worksheet.push_comment(comment);
        }
    }
}

fn parse_comments_xml(xml: &[u8]) -> Result<Vec<Comment>> {
    let mut reader = Reader::from_reader(Cursor::new(xml));
    reader.config_mut().trim_text(false);
    let mut buffer = Vec::new();
    let mut comments = Vec::new();
    let mut authors = Vec::<String>::new();
    let mut in_author = false;
    let mut current_author_text = String::new();
    let mut current_comment_ref: Option<String> = None;
    let mut current_comment_author_id: Option<usize> = None;
    let mut current_comment_text = String::new();
    let mut in_comment_text = false;

    loop {
        match reader.read_event_into(&mut buffer)? {
            Event::Start(ref event) if local_name(event.name().as_ref()) == b"author" => {
                in_author = true;
                current_author_text.clear();
            }
            Event::Text(ref event) if in_author => {
                let text = event
                    .xml_content()
                    .map_err(quick_xml::Error::from)?
                    .into_owned();
                current_author_text.push_str(text.as_str());
            }
            Event::End(ref event) if local_name(event.name().as_ref()) == b"author" => {
                in_author = false;
                authors.push(std::mem::take(&mut current_author_text));
            }
            Event::Start(ref event) if local_name(event.name().as_ref()) == b"comment" => {
                for attribute in event.attributes().flatten() {
                    let key = local_name(attribute.key.as_ref());
                    let value = String::from_utf8_lossy(attribute.value.as_ref()).into_owned();
                    match key {
                        b"ref" => current_comment_ref = Some(value),
                        b"authorId" => {
                            current_comment_author_id = value.trim().parse::<usize>().ok()
                        }
                        _ => {}
                    }
                }
                current_comment_text.clear();
            }
            Event::Start(ref event)
                if local_name(event.name().as_ref()) == b"t" && current_comment_ref.is_some() =>
            {
                in_comment_text = true;
            }
            Event::Text(ref event) if in_comment_text => {
                let text = event
                    .xml_content()
                    .map_err(quick_xml::Error::from)?
                    .into_owned();
                current_comment_text.push_str(text.as_str());
            }
            Event::End(ref event) if local_name(event.name().as_ref()) == b"t" => {
                in_comment_text = false;
            }
            Event::End(ref event) if local_name(event.name().as_ref()) == b"comment" => {
                in_comment_text = false;
                if let Some(cell_ref) = current_comment_ref.take() {
                    let author = current_comment_author_id
                        .and_then(|id| authors.get(id))
                        .cloned()
                        .unwrap_or_default();
                    if let Ok(comment) = Comment::from_parsed_parts(
                        cell_ref,
                        author,
                        std::mem::take(&mut current_comment_text),
                    ) {
                        comments.push(comment);
                    }
                }
                current_comment_author_id = None;
            }
            Event::Eof => break,
            _ => {}
        }
        buffer.clear();
    }

    Ok(comments)
}

fn serialize_comments_xml(comments: &[Comment]) -> Result<Vec<u8>> {
    let mut writer = Writer::new_with_indent(Vec::new(), b' ', 2);
    writer.write_event(Event::Decl(BytesDecl::new(
        "1.0",
        Some("UTF-8"),
        Some("yes"),
    )))?;

    let mut root = BytesStart::new("comments");
    root.push_attribute(("xmlns", SPREADSHEETML_NS));
    writer.write_event(Event::Start(root))?;

    // Collect unique authors
    let mut authors = Vec::<String>::new();
    for comment in comments {
        if !authors.contains(&comment.author().to_string()) {
            authors.push(comment.author().to_string());
        }
    }

    writer.write_event(Event::Start(BytesStart::new("authors")))?;
    for author in &authors {
        writer.write_event(Event::Start(BytesStart::new("author")))?;
        writer.write_event(Event::Text(BytesText::new(author.as_str())))?;
        writer.write_event(Event::End(BytesEnd::new("author")))?;
    }
    writer.write_event(Event::End(BytesEnd::new("authors")))?;

    writer.write_event(Event::Start(BytesStart::new("commentList")))?;
    for comment in comments {
        let author_id = authors
            .iter()
            .position(|a| a == comment.author())
            .unwrap_or(0);
        let mut comment_tag = BytesStart::new("comment");
        comment_tag.push_attribute(("ref", comment.cell_ref()));
        let author_id_text = author_id.to_string();
        comment_tag.push_attribute(("authorId", author_id_text.as_str()));
        writer.write_event(Event::Start(comment_tag))?;

        writer.write_event(Event::Start(BytesStart::new("text")))?;
        writer.write_event(Event::Start(BytesStart::new("t")))?;
        writer.write_event(Event::Text(BytesText::new(comment.text())))?;
        writer.write_event(Event::End(BytesEnd::new("t")))?;
        writer.write_event(Event::End(BytesEnd::new("text")))?;

        writer.write_event(Event::End(BytesEnd::new("comment")))?;
    }
    writer.write_event(Event::End(BytesEnd::new("commentList")))?;

    writer.write_event(Event::End(BytesEnd::new("comments")))?;
    Ok(writer.into_inner())
}

fn parse_sqref_ranges(sqref: &str) -> Vec<CellRange> {
    let mut ranges = Vec::new();
    for token in sqref.split_whitespace() {
        if let Ok(range) = CellRange::parse(token) {
            if !ranges.contains(&range) {
                ranges.push(range);
            }
        }
    }
    ranges
}

fn format_cell_range(range: &CellRange) -> String {
    if range.start() == range.end() {
        range.start().to_string()
    } else {
        format!("{}:{}", range.start(), range.end())
    }
}

fn media_extension_from_content_type(content_type: &str) -> &'static str {
    match content_type.trim().to_ascii_lowercase().as_str() {
        "image/png" => "png",
        "image/jpeg" => "jpeg",
        "image/jpg" => "jpg",
        "image/gif" => "gif",
        "image/bmp" => "bmp",
        "image/tiff" => "tiff",
        "image/x-icon" => "ico",
        "image/svg+xml" => "svg",
        _ => "bin",
    }
}

fn cell_reference_to_column_row(reference: &str) -> Result<(u32, u32)> {
    let normalized = normalize_cell_reference(reference)?;
    let split_index = normalized
        .char_indices()
        .find_map(|(index, ch)| ch.is_ascii_digit().then_some(index))
        .ok_or_else(|| XlsxError::InvalidCellReference(reference.to_string()))?;
    let (column_name, row_text) = normalized.split_at(split_index);
    let column_index = column_name
        .bytes()
        .try_fold(0_u32, |acc, byte| {
            acc.checked_mul(26)
                .and_then(|value| value.checked_add(u32::from(byte - b'A' + 1)))
        })
        .ok_or_else(|| XlsxError::InvalidCellReference(reference.to_string()))?;
    let row_index = row_text
        .parse::<u32>()
        .map_err(|_| XlsxError::InvalidCellReference(reference.to_string()))?;
    Ok((column_index, row_index))
}

fn build_cell_reference(column_index: u32, row_index: u32) -> Result<String> {
    let column_name = column_index_to_name(column_index)?;
    if row_index == 0 {
        return Err(XlsxError::InvalidCellReference(format!(
            "{column_name}{row_index}"
        )));
    }
    Ok(format!("{column_name}{row_index}"))
}

fn column_index_to_name(mut column_index: u32) -> Result<String> {
    if column_index == 0 {
        return Err(XlsxError::InvalidCellReference("0".to_string()));
    }

    let mut letters = Vec::new();
    while column_index > 0 {
        let remainder = (column_index - 1) % 26;
        letters.push((b'A' + remainder as u8) as char);
        column_index = (column_index - 1) / 26;
    }
    letters.reverse();
    Ok(letters.into_iter().collect())
}

fn row_index_from_reference(reference: &str) -> Result<u32> {
    let index = reference
        .char_indices()
        .find_map(|(idx, ch)| if ch.is_ascii_digit() { Some(idx) } else { None })
        .ok_or_else(|| XlsxError::InvalidCellReference(reference.to_string()))?;
    let row_text = &reference[index..];
    row_text
        .parse::<u32>()
        .map_err(|_| XlsxError::InvalidCellReference(reference.to_string()))
}

fn local_name(name: &[u8]) -> &[u8] {
    name.rsplit(|byte| *byte == b':').next().unwrap_or(name)
}

#[cfg(test)]
mod tests {
    use tempfile::NamedTempFile;

    use super::*;
    use crate::style::{
        Alignment, Border, BorderSide, Fill, Font, HorizontalAlignment, Style, VerticalAlignment,
    };

    const SHARED_STRINGS_PART_URI: &str = "/xl/sharedStrings.xml";
    const STYLES_PART_URI: &str = "/xl/styles.xml";
    const DRAWING_PART_URI: &str = "/xl/drawings/drawing1.xml";
    const CUSTOM_WORKBOOK_RELATIONSHIP_TYPE: &str =
        "https://offidized.dev/relationships/custom-data";
    const CUSTOM_WORKSHEET_RELATIONSHIP_TYPE: &str =
        "https://offidized.dev/relationships/custom-sheet-data";

    #[test]
    fn basic_sheet_and_cell_set_get() {
        let mut workbook = Workbook::new();
        let sheet = workbook.add_sheet("Sales");

        sheet
            .cell_mut("A1")
            .expect("valid cell A1")
            .set_value("Product");
        sheet
            .cell_mut("B2")
            .expect("valid cell B2")
            .set_value(42_000)
            .set_formula("SUM(B2:B3)")
            .set_style_id(3);

        let sheet = workbook.sheet("Sales").expect("sheet Sales exists");

        assert_eq!(
            sheet.cell("A1").and_then(|cell| cell.value()),
            Some(&CellValue::String("Product".to_string()))
        );
        assert_eq!(
            sheet.cell("B2").and_then(|cell| cell.value()),
            Some(&CellValue::Number(42_000.0))
        );
        assert_eq!(
            sheet.cell("B2").and_then(|cell| cell.formula()),
            Some("SUM(B2:B3)")
        );
        assert_eq!(sheet.cell("B2").and_then(|cell| cell.style_id()), Some(3));
    }

    #[test]
    fn open_save_smoke_test() {
        let mut workbook = Workbook::new();
        let sheet = workbook.add_sheet("Data");

        sheet
            .cell_mut("A1")
            .expect("valid cell A1")
            .set_value("Hello");
        sheet
            .cell_mut("B1")
            .expect("valid cell B1")
            .set_value(123_i32);
        sheet.cell_mut("C1").expect("valid cell C1").set_value(true);

        let temp_file = NamedTempFile::new().expect("temp file should be created");
        workbook.save(temp_file.path()).expect("save workbook");

        let package = Package::open(temp_file.path()).expect("open saved workbook package");
        assert!(package.get_part("/xl/workbook.xml").is_some());
        assert!(package.get_part("/xl/worksheets/sheet1.xml").is_some());
        assert!(package
            .relationships()
            .get_first_by_type(RelationshipType::WORKBOOK)
            .is_some());
        assert_eq!(
            package.content_types().get_override("/xl/workbook.xml"),
            Some(ContentTypeValue::WORKBOOK)
        );
        assert_eq!(
            package
                .content_types()
                .get_override("/xl/worksheets/sheet1.xml"),
            Some(ContentTypeValue::WORKSHEET)
        );
        assert!(package.get_part(STYLES_PART_URI).is_some());
        assert_eq!(
            package.content_types().get_override(STYLES_PART_URI),
            Some(ContentTypeValue::SPREADSHEET_STYLES)
        );

        let loaded = Workbook::open(temp_file.path()).expect("open workbook");
        let loaded_sheet = loaded.sheet("Data").expect("sheet Data exists");

        assert_eq!(
            loaded_sheet.cell("A1").and_then(|cell| cell.value()),
            Some(&CellValue::String("Hello".to_string()))
        );
        assert_eq!(
            loaded_sheet.cell("B1").and_then(|cell| cell.value()),
            Some(&CellValue::Number(123.0))
        );
        assert_eq!(
            loaded_sheet.cell("C1").and_then(|cell| cell.value()),
            Some(&CellValue::Bool(true))
        );
    }

    #[test]
    fn date_and_error_cells_roundtrip() {
        let mut workbook = Workbook::new();
        let sheet = workbook.add_sheet("Data");
        sheet
            .cell_mut("A1")
            .expect("cell should be valid")
            .set_value(CellValue::date("2026-01-31T12:00:00"));
        sheet
            .cell_mut("B1")
            .expect("cell should be valid")
            .set_value(CellValue::error("#DIV/0!"));

        let temp_file = NamedTempFile::new().expect("temp file should be created");
        workbook.save(temp_file.path()).expect("save workbook");

        let package = Package::open(temp_file.path()).expect("open saved workbook package");
        let worksheet_part = package
            .get_part("/xl/worksheets/sheet1.xml")
            .expect("worksheet part should exist");
        let worksheet_xml = String::from_utf8_lossy(worksheet_part.data.as_bytes());
        assert!(worksheet_xml.contains("t=\"d\""));
        assert!(worksheet_xml.contains("<v>2026-01-31T12:00:00</v>"));
        assert!(worksheet_xml.contains("t=\"e\""));
        assert!(worksheet_xml.contains("<v>#DIV/0!</v>"));

        let loaded = Workbook::open(temp_file.path()).expect("open workbook");
        let loaded_sheet = loaded.sheet("Data").expect("sheet should exist");
        assert_eq!(
            loaded_sheet.cell("A1").and_then(Cell::value),
            Some(&CellValue::Date("2026-01-31T12:00:00".to_string()))
        );
        assert_eq!(
            loaded_sheet.cell("B1").and_then(Cell::value),
            Some(&CellValue::Error("#DIV/0!".to_string()))
        );
    }

    #[test]
    fn styles_roundtrip_with_number_formats_alignment_and_primitives() {
        let mut workbook = Workbook::new();

        let mut date_alignment = Alignment::new();
        date_alignment
            .set_horizontal(HorizontalAlignment::Center)
            .set_vertical(VerticalAlignment::Top)
            .set_wrap_text(true);
        let mut date_font = Font::new();
        date_font
            .set_name("Aptos")
            .set_size("11")
            .set_bold(true)
            .set_italic(true)
            .set_underline(true)
            .set_color("FF112233")
            .set_color_ref(ColorReference::from_rgb("FF112233"));
        let mut date_fill = Fill::new();
        date_fill
            .set_pattern("solid")
            .set_foreground_color("FF445566")
            .set_background_color("FF778899");
        // The parser auto-builds a PatternFill from the string-based fields on load,
        // so we must set it here too for the roundtrip assertion to match.
        let mut pf = PatternFill::new(PatternFillType::Solid);
        pf.set_fg_color(ColorReference::from_rgb("FF445566"));
        pf.set_bg_color(ColorReference::from_rgb("FF778899"));
        date_fill.set_pattern_fill(pf);
        let mut left_border_side = BorderSide::new();
        left_border_side
            .set_style("thin")
            .set_color("FFABCDEF")
            .set_color_ref(ColorReference::from_rgb("FFABCDEF"));
        let mut bottom_border_side = BorderSide::new();
        bottom_border_side
            .set_style("double")
            .set_color("FFA1B2C3")
            .set_color_ref(ColorReference::from_rgb("FFA1B2C3"));
        let mut date_border = Border::new();
        date_border
            .set_left(left_border_side)
            .set_bottom(bottom_border_side);

        let mut date_style = Style::new();
        date_style
            .set_number_format("yyyy-mm-dd")
            .set_alignment(date_alignment.clone())
            .set_font(date_font.clone())
            .set_fill(date_fill.clone())
            .set_border(date_border.clone());
        let date_style_id = workbook
            .add_style(date_style)
            .expect("style id should fit within u32");

        let mut numeric_style = Style::new();
        numeric_style.set_number_format("0.00");
        let numeric_style_id = workbook
            .add_style(numeric_style)
            .expect("style id should fit within u32");

        let sheet = workbook.add_sheet("Data");
        sheet
            .cell_mut("A1")
            .expect("cell should be valid")
            .set_value(CellValue::date("2026-02-12T00:00:00"))
            .set_style_id(date_style_id);
        sheet
            .cell_mut("B1")
            .expect("cell should be valid")
            .set_value(123.45_f64)
            .set_style_id(numeric_style_id);

        let temp_file = NamedTempFile::new().expect("temp file should be created");
        workbook.save(temp_file.path()).expect("save workbook");

        let package = Package::open(temp_file.path()).expect("open saved workbook package");
        let workbook_part = package
            .get_part(WORKBOOK_PART_URI)
            .expect("workbook part should exist");
        assert!(workbook_part
            .relationships
            .get_first_by_type(RelationshipType::STYLES)
            .is_some());
        assert!(package.get_part(STYLES_PART_URI).is_some());
        assert_eq!(
            package.content_types().get_override(STYLES_PART_URI),
            Some(ContentTypeValue::SPREADSHEET_STYLES)
        );

        let styles_part = package
            .get_part(STYLES_PART_URI)
            .expect("styles part should exist");
        let styles_xml = String::from_utf8_lossy(styles_part.data.as_bytes());
        assert!(styles_xml.contains("<numFmts count=\"2\">"));
        assert!(styles_xml.contains("formatCode=\"yyyy-mm-dd\""));
        assert!(styles_xml.contains("formatCode=\"0.00\""));
        assert!(styles_xml.contains("<fonts count=\"2\">"));
        assert!(styles_xml.contains("<name val=\"Aptos\""));
        assert!(styles_xml.contains("<sz val=\"11\""));
        assert!(styles_xml.contains("<b"));
        assert!(styles_xml.contains("<i"));
        assert!(styles_xml.contains("<u"));
        assert!(styles_xml.contains("rgb=\"FF112233\""));
        assert!(styles_xml.contains("<fills count=\"2\">"));
        assert!(styles_xml.contains("patternType=\"solid\""));
        assert!(styles_xml.contains("rgb=\"FF445566\""));
        assert!(styles_xml.contains("rgb=\"FF778899\""));
        assert!(styles_xml.contains("<borders count=\"2\">"));
        assert!(styles_xml.contains("<left style=\"thin\">"));
        assert!(styles_xml.contains("<bottom style=\"double\">"));
        assert!(styles_xml.contains("rgb=\"FFABCDEF\""));
        assert!(styles_xml.contains("rgb=\"FFA1B2C3\""));
        assert!(styles_xml.contains("<cellXfs count=\"3\">"));
        assert!(styles_xml.contains("applyFont=\"1\""));
        assert!(styles_xml.contains("applyFill=\"1\""));
        assert!(styles_xml.contains("applyBorder=\"1\""));
        assert!(styles_xml.contains("applyAlignment=\"1\""));
        assert!(styles_xml.contains("horizontal=\"center\""));
        assert!(styles_xml.contains("vertical=\"top\""));
        assert!(styles_xml.contains("wrapText=\"1\""));

        let loaded = Workbook::open(temp_file.path()).expect("open workbook");
        assert_eq!(loaded.styles().len(), 3);

        let loaded_date_style = loaded
            .style(date_style_id)
            .expect("date style should be loaded");
        assert_eq!(loaded_date_style.number_format(), Some("yyyy-mm-dd"));
        let loaded_alignment = loaded_date_style
            .alignment()
            .expect("date style alignment should be loaded");
        assert_eq!(
            loaded_alignment.horizontal(),
            Some(HorizontalAlignment::Center)
        );
        assert_eq!(loaded_alignment.vertical(), Some(VerticalAlignment::Top));
        assert_eq!(loaded_alignment.wrap_text(), Some(true));
        assert_eq!(loaded_date_style.font(), Some(&date_font));
        assert_eq!(loaded_date_style.fill(), Some(&date_fill));
        assert_eq!(loaded_date_style.border(), Some(&date_border));

        let loaded_numeric_style = loaded
            .style(numeric_style_id)
            .expect("numeric style should be loaded");
        assert_eq!(loaded_numeric_style.number_format(), Some("0.00"));
        assert_eq!(loaded_numeric_style.font(), None);
        assert_eq!(loaded_numeric_style.fill(), None);
        assert_eq!(loaded_numeric_style.border(), None);

        let loaded_sheet = loaded.sheet("Data").expect("sheet should be loaded");
        assert_eq!(
            loaded_sheet.cell("A1").and_then(Cell::style_id),
            Some(date_style_id)
        );
        assert_eq!(
            loaded_sheet.cell("B1").and_then(Cell::style_id),
            Some(numeric_style_id)
        );
    }

    #[test]
    fn workbook_sheet_helpers_work() {
        let mut workbook = Workbook::new();
        workbook.add_sheet("Summary");
        workbook.add_sheet("Data");
        workbook.add_sheet("Summary");

        assert!(workbook.contains_sheet("Summary"));
        assert!(workbook.contains_sheet("Data"));
        assert!(!workbook.contains_sheet("Missing"));

        assert_eq!(workbook.sheet_names(), vec!["Summary", "Data"]);
        assert_eq!(
            workbook
                .worksheets()
                .iter()
                .map(Worksheet::name)
                .collect::<Vec<_>>(),
            vec!["Summary", "Data"]
        );

        for worksheet in workbook.worksheets_mut().iter_mut() {
            let sheet_name = worksheet.name().to_string();
            worksheet
                .cell_mut("A1")
                .expect("cell should be valid")
                .set_value(format!("{sheet_name} sheet"));
        }

        assert_eq!(
            workbook
                .sheet("Summary")
                .and_then(|sheet| sheet.cell("A1"))
                .and_then(Cell::value),
            Some(&CellValue::String("Summary sheet".to_string()))
        );

        let removed = workbook
            .remove_sheet("Data")
            .expect("sheet should be removed");
        assert_eq!(removed.name(), "Data");
        assert!(!workbook.contains_sheet("Data"));
        assert!(workbook.remove_sheet("Missing").is_none());
    }

    #[test]
    fn open_save_roundtrip_regression_multiple_sheets() {
        let mut workbook = Workbook::new();
        workbook
            .add_sheet("Summary")
            .cell_mut("A1")
            .expect("cell should be valid")
            .set_value("Report");
        workbook
            .sheet_mut("Summary")
            .expect("sheet should exist")
            .cell_mut("B2")
            .expect("cell should be valid")
            .set_value(9.5_f64)
            .set_formula("SUM(B2:B4)")
            .set_style_id(11);

        workbook
            .add_sheet("Data")
            .cell_mut("A1")
            .expect("cell should be valid")
            .set_value(true);
        workbook
            .sheet_mut("Data")
            .expect("sheet should exist")
            .cell_mut("A2")
            .expect("cell should be valid")
            .set_value("ok");

        let temp_file_one = NamedTempFile::new().expect("temp file should be created");
        workbook.save(temp_file_one.path()).expect("save workbook");

        let reloaded = Workbook::open(temp_file_one.path()).expect("open workbook");
        assert_eq!(reloaded.sheet_names(), vec!["Summary", "Data"]);
        assert_eq!(
            reloaded
                .sheet("Summary")
                .and_then(|sheet| sheet.cell("A1"))
                .and_then(Cell::value),
            Some(&CellValue::String("Report".to_string()))
        );
        assert_eq!(
            reloaded
                .sheet("Summary")
                .and_then(|sheet| sheet.cell("B2"))
                .and_then(Cell::formula),
            Some("SUM(B2:B4)")
        );
        assert_eq!(
            reloaded
                .sheet("Summary")
                .and_then(|sheet| sheet.cell("B2"))
                .and_then(Cell::style_id),
            Some(11)
        );
        assert_eq!(
            reloaded
                .sheet("Data")
                .and_then(|sheet| sheet.cell("A1"))
                .and_then(Cell::value),
            Some(&CellValue::Bool(true))
        );

        let temp_file_two = NamedTempFile::new().expect("temp file should be created");
        reloaded.save(temp_file_two.path()).expect("save workbook");
        let reopened = Workbook::open(temp_file_two.path()).expect("open workbook");
        assert_eq!(reopened.sheet_names(), vec!["Summary", "Data"]);
        assert_eq!(
            reopened
                .sheet("Data")
                .and_then(|sheet| sheet.cell("A2"))
                .and_then(Cell::value),
            Some(&CellValue::String("ok".to_string()))
        );
    }

    #[test]
    fn save_writes_shared_strings_part_and_roundtrips_values() {
        let mut workbook = Workbook::new();
        let sheet = workbook.add_sheet("Data");
        sheet
            .cell_mut("A1")
            .expect("cell should be valid")
            .set_value("Hello");
        sheet
            .cell_mut("A2")
            .expect("cell should be valid")
            .set_value("Hello");
        sheet
            .cell_mut("B1")
            .expect("cell should be valid")
            .set_value("World");

        let temp_file = NamedTempFile::new().expect("temp file should be created");
        workbook.save(temp_file.path()).expect("save workbook");

        let package = Package::open(temp_file.path()).expect("open saved workbook package");
        let workbook_part = package
            .get_part(WORKBOOK_PART_URI)
            .expect("workbook part should exist");
        let shared_strings_relationship = workbook_part
            .relationships
            .get_first_by_type(RelationshipType::SHARED_STRINGS)
            .expect("shared strings relationship should exist");
        assert_eq!(
            shared_strings_relationship.target_mode,
            TargetMode::Internal
        );

        let resolved_shared_strings_uri = PartUri::new(WORKBOOK_PART_URI)
            .expect("workbook uri should be valid")
            .resolve_relative(shared_strings_relationship.target.as_str())
            .expect("shared strings target should resolve");
        assert_eq!(
            resolved_shared_strings_uri.as_str(),
            SHARED_STRINGS_PART_URI
        );
        assert!(package.get_part(SHARED_STRINGS_PART_URI).is_some());
        assert_eq!(
            package
                .content_types()
                .get_override(SHARED_STRINGS_PART_URI),
            Some(ContentTypeValue::SHARED_STRINGS)
        );

        let shared_strings_part = package
            .get_part(SHARED_STRINGS_PART_URI)
            .expect("shared strings part should exist");
        let shared_strings = parse_shared_strings_xml(shared_strings_part.data.as_bytes())
            .expect("parse shared strings xml");
        let plain_texts: Vec<String> = shared_strings.iter().map(|e| e.plain_text()).collect();
        assert_eq!(plain_texts, vec!["Hello".to_string(), "World".to_string()]);

        let worksheet_part = package
            .get_part("/xl/worksheets/sheet1.xml")
            .expect("worksheet part should exist");
        let worksheet_xml = String::from_utf8_lossy(worksheet_part.data.as_bytes());
        assert!(worksheet_xml.contains("t=\"s\""));
        assert!(!worksheet_xml.contains("inlineStr"));
        assert_eq!(worksheet_xml.matches("<v>0</v>").count(), 2);
        assert_eq!(worksheet_xml.matches("<v>1</v>").count(), 1);

        let loaded = Workbook::open(temp_file.path()).expect("open workbook");
        let loaded_sheet = loaded.sheet("Data").expect("sheet should exist");
        assert_eq!(
            loaded_sheet.cell("A1").and_then(Cell::value),
            Some(&CellValue::String("Hello".to_string()))
        );
        assert_eq!(
            loaded_sheet.cell("A2").and_then(Cell::value),
            Some(&CellValue::String("Hello".to_string()))
        );
        assert_eq!(
            loaded_sheet.cell("B1").and_then(Cell::value),
            Some(&CellValue::String("World".to_string()))
        );
    }

    #[test]
    fn open_supports_legacy_inline_string_cells() {
        let workbook_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Data" sheetId="1" r:id="rId1"/>
  </sheets>
</workbook>"#;
        let worksheet_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1" t="inlineStr">
        <is>
          <t>Legacy inline</t>
        </is>
      </c>
    </row>
  </sheetData>
</worksheet>"#;

        let mut package = Package::new();
        let workbook_uri = PartUri::new(WORKBOOK_PART_URI).expect("workbook uri should be valid");
        let mut workbook_part = Part::new_xml(workbook_uri, workbook_xml.as_bytes().to_vec());
        workbook_part.content_type = Some(ContentTypeValue::WORKBOOK.to_string());
        workbook_part.relationships.add_new(
            RelationshipType::WORKSHEET.to_string(),
            "worksheets/sheet1.xml".to_string(),
            TargetMode::Internal,
        );
        package.set_part(workbook_part);

        let mut worksheet_part = Part::new_xml(
            PartUri::new("/xl/worksheets/sheet1.xml").expect("worksheet uri should be valid"),
            worksheet_xml.as_bytes().to_vec(),
        );
        worksheet_part.content_type = Some(ContentTypeValue::WORKSHEET.to_string());
        package.set_part(worksheet_part);

        package.relationships_mut().add_new(
            RelationshipType::WORKBOOK.to_string(),
            WORKBOOK_PART_URI.to_string(),
            TargetMode::Internal,
        );

        let temp_file = NamedTempFile::new().expect("temp file should be created");
        package.save(temp_file.path()).expect("save package");

        let loaded = Workbook::open(temp_file.path()).expect("open workbook");
        assert_eq!(
            loaded
                .sheet("Data")
                .and_then(|sheet| sheet.cell("A1"))
                .and_then(Cell::value),
            Some(&CellValue::String("Legacy inline".to_string()))
        );
    }

    #[test]
    fn open_supports_rows_and_cells_without_explicit_references() {
        let workbook_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Data" sheetId="1" r:id="rId1"/>
  </sheets>
</workbook>"#;
        let worksheet_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row>
      <c t="inlineStr">
        <is><t>first</t></is>
      </c>
      <c>
        <v>42</v>
      </c>
    </row>
  </sheetData>
</worksheet>"#;

        let mut package = Package::new();
        let workbook_uri = PartUri::new(WORKBOOK_PART_URI).expect("workbook uri should be valid");
        let mut workbook_part = Part::new_xml(workbook_uri, workbook_xml.as_bytes().to_vec());
        workbook_part.content_type = Some(ContentTypeValue::WORKBOOK.to_string());
        workbook_part.relationships.add_new(
            RelationshipType::WORKSHEET.to_string(),
            "worksheets/sheet1.xml".to_string(),
            TargetMode::Internal,
        );
        package.set_part(workbook_part);

        let mut worksheet_part = Part::new_xml(
            PartUri::new("/xl/worksheets/sheet1.xml").expect("worksheet uri should be valid"),
            worksheet_xml.as_bytes().to_vec(),
        );
        worksheet_part.content_type = Some(ContentTypeValue::WORKSHEET.to_string());
        package.set_part(worksheet_part);

        package.relationships_mut().add_new(
            RelationshipType::WORKBOOK.to_string(),
            WORKBOOK_PART_URI.to_string(),
            TargetMode::Internal,
        );

        let temp_file = NamedTempFile::new().expect("temp file should be created");
        package.save(temp_file.path()).expect("save package");

        let loaded = Workbook::open(temp_file.path()).expect("open workbook");
        let sheet = loaded.sheet("Data").expect("sheet should exist");
        assert_eq!(
            sheet.cell("A1").and_then(Cell::value),
            Some(&CellValue::String("first".to_string()))
        );
        assert_eq!(
            sheet.cell("B1").and_then(Cell::value),
            Some(&CellValue::Number(42.0))
        );
    }

    #[test]
    fn open_treats_shared_string_cells_missing_value_as_blank() {
        let workbook_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Data" sheetId="1" r:id="rId1"/>
  </sheets>
</workbook>"#;
        let worksheet_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1" t="s"/>
      <c r="B1" t="s"><v>0</v></c>
    </row>
  </sheetData>
</worksheet>"#;
        let shared_strings_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="1" uniqueCount="1">
  <si><t>hello</t></si>
</sst>"#;

        let mut package = Package::new();
        let workbook_uri = PartUri::new(WORKBOOK_PART_URI).expect("workbook uri should be valid");
        let mut workbook_part = Part::new_xml(workbook_uri, workbook_xml.as_bytes().to_vec());
        workbook_part.content_type = Some(ContentTypeValue::WORKBOOK.to_string());
        workbook_part.relationships.add_new(
            RelationshipType::WORKSHEET.to_string(),
            "worksheets/sheet1.xml".to_string(),
            TargetMode::Internal,
        );
        workbook_part.relationships.add_new(
            RelationshipType::SHARED_STRINGS.to_string(),
            "sharedStrings.xml".to_string(),
            TargetMode::Internal,
        );
        package.set_part(workbook_part);

        let mut worksheet_part = Part::new_xml(
            PartUri::new("/xl/worksheets/sheet1.xml").expect("worksheet uri should be valid"),
            worksheet_xml.as_bytes().to_vec(),
        );
        worksheet_part.content_type = Some(ContentTypeValue::WORKSHEET.to_string());
        package.set_part(worksheet_part);

        let mut shared_strings_part = Part::new_xml(
            PartUri::new(SHARED_STRINGS_PART_URI).expect("shared strings uri should be valid"),
            shared_strings_xml.as_bytes().to_vec(),
        );
        shared_strings_part.content_type = Some(ContentTypeValue::SHARED_STRINGS.to_string());
        package.set_part(shared_strings_part);

        package.relationships_mut().add_new(
            RelationshipType::WORKBOOK.to_string(),
            WORKBOOK_PART_URI.to_string(),
            TargetMode::Internal,
        );

        let temp_file = NamedTempFile::new().expect("temp file should be created");
        package.save(temp_file.path()).expect("save package");

        let loaded = Workbook::open(temp_file.path()).expect("open workbook");
        let sheet = loaded.sheet("Data").expect("sheet should exist");
        assert_eq!(
            sheet.cell("A1").and_then(Cell::value),
            Some(&CellValue::Blank)
        );
        assert_eq!(
            sheet.cell("B1").and_then(Cell::value),
            Some(&CellValue::String("hello".to_string()))
        );
    }

    #[test]
    fn defined_names_roundtrip_through_workbook_xml() {
        let mut workbook = Workbook::new();
        workbook.add_sheet("Data");
        workbook.add_defined_name("GlobalRange", "Data!$A$1:$A$5");
        workbook
            .add_defined_name("LocalRange", "Data!$B$1:$B$3")
            .set_local_sheet_id(0);

        let temp_file = NamedTempFile::new().expect("temp file should be created");
        workbook.save(temp_file.path()).expect("save workbook");

        let package = Package::open(temp_file.path()).expect("open saved workbook package");
        let workbook_part = package
            .get_part(WORKBOOK_PART_URI)
            .expect("workbook part should exist");
        let workbook_xml = String::from_utf8_lossy(workbook_part.data.as_bytes());
        assert!(workbook_xml.contains("<definedNames>"));
        assert!(workbook_xml.contains("name=\"GlobalRange\""));
        assert!(workbook_xml.contains(">Data!$A$1:$A$5</definedName>"));
        assert!(workbook_xml.contains("name=\"LocalRange\""));
        assert!(workbook_xml.contains("localSheetId=\"0\""));
        assert!(workbook_xml.contains(">Data!$B$1:$B$3</definedName>"));

        let loaded = Workbook::open(temp_file.path()).expect("open workbook");
        assert_eq!(loaded.defined_names().len(), 2);
        assert_eq!(
            loaded
                .defined_name("GlobalRange")
                .expect("global defined name should exist")
                .reference(),
            "Data!$A$1:$A$5"
        );
        assert_eq!(
            loaded
                .defined_name("GlobalRange")
                .expect("global defined name should exist")
                .local_sheet_id(),
            None
        );
        assert_eq!(
            loaded
                .defined_name("LocalRange")
                .expect("local defined name should exist")
                .reference(),
            "Data!$B$1:$B$3"
        );
        assert_eq!(
            loaded
                .defined_name("LocalRange")
                .expect("local defined name should exist")
                .local_sheet_id(),
            Some(0)
        );
    }

    #[test]
    fn merge_cells_roundtrip_per_worksheet() {
        let mut workbook = Workbook::new();
        let sheet = workbook.add_sheet("Data");
        sheet
            .add_merged_range("B2:A1")
            .expect("merged range should parse");
        sheet
            .add_merged_range("A1:B2")
            .expect("duplicate merged range should parse");
        sheet
            .add_merged_range("D4")
            .expect("single-cell merged range should parse");

        let temp_file = NamedTempFile::new().expect("temp file should be created");
        workbook.save(temp_file.path()).expect("save workbook");

        let package = Package::open(temp_file.path()).expect("open saved workbook package");
        let worksheet_part = package
            .get_part("/xl/worksheets/sheet1.xml")
            .expect("worksheet part should exist");
        let worksheet_xml = String::from_utf8_lossy(worksheet_part.data.as_bytes());
        assert!(worksheet_xml.contains("<mergeCells count=\"2\">"));
        assert!(worksheet_xml.contains("<mergeCell ref=\"A1:B2\"/>"));
        assert!(worksheet_xml.contains("<mergeCell ref=\"D4\"/>"));

        let loaded = Workbook::open(temp_file.path()).expect("open workbook");
        let loaded_sheet = loaded.sheet("Data").expect("sheet should exist");
        assert_eq!(loaded_sheet.merged_ranges().len(), 2);
        assert_eq!(loaded_sheet.merged_ranges()[0].start(), "A1");
        assert_eq!(loaded_sheet.merged_ranges()[0].end(), "B2");
        assert_eq!(loaded_sheet.merged_ranges()[1].start(), "D4");
        assert_eq!(loaded_sheet.merged_ranges()[1].end(), "D4");
    }

    #[test]
    fn freeze_panes_roundtrip_per_worksheet() {
        let mut workbook = Workbook::new();
        workbook
            .add_sheet("Data")
            .set_freeze_panes_with_top_left_cell(1, 2, "c5")
            .expect("freeze pane should be valid");

        let temp_file = NamedTempFile::new().expect("temp file should be created");
        workbook.save(temp_file.path()).expect("save workbook");

        let package = Package::open(temp_file.path()).expect("open saved workbook package");
        let worksheet_part = package
            .get_part("/xl/worksheets/sheet1.xml")
            .expect("worksheet part should exist");
        let worksheet_xml = String::from_utf8_lossy(worksheet_part.data.as_bytes());
        assert!(worksheet_xml.contains("<sheetViews>"));
        assert!(worksheet_xml.contains("<pane"));
        assert!(worksheet_xml.contains("xSplit=\"1\""));
        assert!(worksheet_xml.contains("ySplit=\"2\""));
        assert!(worksheet_xml.contains("topLeftCell=\"C5\""));
        assert!(worksheet_xml.contains("state=\"frozen\""));
        assert!(worksheet_xml.contains("activePane=\"bottomRight\""));

        let loaded = Workbook::open(temp_file.path()).expect("open workbook");
        let freeze_pane = loaded
            .sheet("Data")
            .and_then(Worksheet::freeze_pane)
            .expect("freeze pane should be loaded");
        assert_eq!(freeze_pane.x_split(), 1);
        assert_eq!(freeze_pane.y_split(), 2);
        assert_eq!(freeze_pane.top_left_cell(), "C5");
    }

    #[test]
    fn auto_filter_roundtrip_per_worksheet() {
        let mut workbook = Workbook::new();
        workbook
            .add_sheet("Data")
            .set_auto_filter("C5:A1")
            .expect("auto filter range should be valid");

        let temp_file = NamedTempFile::new().expect("temp file should be created");
        workbook.save(temp_file.path()).expect("save workbook");

        let package = Package::open(temp_file.path()).expect("open saved workbook package");
        let worksheet_part = package
            .get_part("/xl/worksheets/sheet1.xml")
            .expect("worksheet part should exist");
        let worksheet_xml = String::from_utf8_lossy(worksheet_part.data.as_bytes());
        assert!(worksheet_xml.contains("<autoFilter ref=\"A1:C5\"/>"));

        let loaded = Workbook::open(temp_file.path()).expect("open workbook");
        let auto_filter = loaded
            .sheet("Data")
            .and_then(Worksheet::auto_filter)
            .expect("auto filter should be loaded");
        let range = auto_filter
            .range()
            .expect("auto filter range should be set");
        assert_eq!(range.start(), "A1");
        assert_eq!(range.end(), "C5");
    }

    #[test]
    fn worksheet_tables_roundtrip_with_parts_and_relationships() {
        let mut workbook = Workbook::new();
        let sheet = workbook.add_sheet("Data");
        sheet.add_table(
            WorksheetTable::with_header_row("SalesTable", "A1:C4", true)
                .expect("table should be valid"),
        );
        sheet.add_table(
            WorksheetTable::with_header_row("ArchiveTable", "E2:F5", false)
                .expect("table should be valid"),
        );

        let temp_file = NamedTempFile::new().expect("temp file should be created");
        workbook.save(temp_file.path()).expect("save workbook");

        let package = Package::open(temp_file.path()).expect("open saved workbook package");
        let worksheet_part = package
            .get_part("/xl/worksheets/sheet1.xml")
            .expect("worksheet part should exist");
        let worksheet_xml = String::from_utf8_lossy(worksheet_part.data.as_bytes());
        assert!(worksheet_xml.contains("<tableParts count=\"2\">"));
        assert_eq!(worksheet_xml.matches("<tablePart r:id=").count(), 2);

        let table_relationships = worksheet_part
            .relationships
            .get_by_type(TABLE_RELATIONSHIP_TYPE);
        assert_eq!(table_relationships.len(), 2);
        for relationship in table_relationships {
            assert_eq!(relationship.target_mode, TargetMode::Internal);
            let table_uri = PartUri::new("/xl/worksheets/sheet1.xml")
                .expect("worksheet uri should be valid")
                .resolve_relative(relationship.target.as_str())
                .expect("table target should resolve");
            assert!(
                package.get_part(table_uri.as_str()).is_some(),
                "missing table part {}",
                table_uri.as_str()
            );
            assert_eq!(
                package.content_types().get_override(table_uri.as_str()),
                Some(TABLE_PART_CONTENT_TYPE)
            );
        }

        let table1_part = package
            .get_part("/xl/tables/table1.xml")
            .expect("first table part should exist");
        let table1_xml = String::from_utf8_lossy(table1_part.data.as_bytes());
        assert!(table1_xml.contains("name=\"SalesTable\""));
        assert!(table1_xml.contains("displayName=\"SalesTable\""));
        assert!(table1_xml.contains("ref=\"A1:C4\""));
        assert!(table1_xml.contains("headerRowCount=\"1\""));

        let table2_part = package
            .get_part("/xl/tables/table2.xml")
            .expect("second table part should exist");
        let table2_xml = String::from_utf8_lossy(table2_part.data.as_bytes());
        assert!(table2_xml.contains("name=\"ArchiveTable\""));
        assert!(table2_xml.contains("displayName=\"ArchiveTable\""));
        assert!(table2_xml.contains("ref=\"E2:F5\""));
        assert!(table2_xml.contains("headerRowCount=\"0\""));

        let loaded = Workbook::open(temp_file.path()).expect("open workbook");
        let loaded_sheet = loaded.sheet("Data").expect("sheet should exist");
        assert_eq!(loaded_sheet.tables().len(), 2);
        assert_eq!(loaded_sheet.tables()[0].name(), "SalesTable");
        assert_eq!(loaded_sheet.tables()[0].range().start(), "A1");
        assert_eq!(loaded_sheet.tables()[0].range().end(), "C4");
        assert!(loaded_sheet.tables()[0].has_header_row());
        assert_eq!(loaded_sheet.tables()[1].name(), "ArchiveTable");
        assert_eq!(loaded_sheet.tables()[1].range().start(), "E2");
        assert_eq!(loaded_sheet.tables()[1].range().end(), "F5");
        assert!(!loaded_sheet.tables()[1].has_header_row());
    }

    #[test]
    fn worksheet_images_roundtrip_with_parts_and_relationships() {
        let mut workbook = Workbook::new();
        let sheet = workbook.add_sheet("Data");
        let custom_ext = WorksheetImageExt::new(222_222, 111_111).expect("ext should be valid");
        sheet
            .add_image(
                vec![137_u8, 80_u8, 78_u8],
                "image/png",
                "B2",
                Some(custom_ext),
            )
            .expect("png image should be valid");
        sheet
            .add_image(vec![255_u8, 216_u8, 255_u8], "image/jpeg", "D5", None)
            .expect("jpeg image should be valid");

        let temp_file = NamedTempFile::new().expect("temp file should be created");
        workbook.save(temp_file.path()).expect("save workbook");

        let package = Package::open(temp_file.path()).expect("open saved workbook package");
        let worksheet_part = package
            .get_part("/xl/worksheets/sheet1.xml")
            .expect("worksheet part should exist");
        let worksheet_xml = String::from_utf8_lossy(worksheet_part.data.as_bytes());
        assert!(worksheet_xml.contains(
            "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\""
        ));
        assert!(worksheet_xml.contains("<drawing r:id="));

        let drawing_relationships = worksheet_part
            .relationships
            .get_by_type(DRAWING_RELATIONSHIP_TYPE);
        assert_eq!(drawing_relationships.len(), 1);
        let drawing_relationship = drawing_relationships[0];
        assert_eq!(drawing_relationship.target_mode, TargetMode::Internal);
        let drawing_uri = PartUri::new("/xl/worksheets/sheet1.xml")
            .expect("worksheet uri should be valid")
            .resolve_relative(drawing_relationship.target.as_str())
            .expect("drawing target should resolve");
        assert_eq!(drawing_uri.as_str(), DRAWING_PART_URI);
        assert_eq!(
            package.content_types().get_override(DRAWING_PART_URI),
            Some(DRAWING_PART_CONTENT_TYPE)
        );

        let drawing_part = package
            .get_part(DRAWING_PART_URI)
            .expect("drawing part should exist");
        let drawing_xml = String::from_utf8_lossy(drawing_part.data.as_bytes());
        assert_eq!(drawing_xml.matches("<xdr:oneCellAnchor>").count(), 2);
        assert!(drawing_xml.contains("<xdr:col>1</xdr:col>"));
        assert!(drawing_xml.contains("<xdr:row>1</xdr:row>"));
        assert!(drawing_xml.contains("cx=\"222222\""));
        assert!(drawing_xml.contains("cy=\"111111\""));
        assert!(drawing_xml.contains("<xdr:col>3</xdr:col>"));
        assert!(drawing_xml.contains("<xdr:row>4</xdr:row>"));
        assert!(drawing_xml.contains("cx=\"952500\""));
        assert!(drawing_xml.contains("cy=\"952500\""));

        let image_relationships = drawing_part
            .relationships
            .get_by_type(RelationshipType::IMAGE);
        assert_eq!(image_relationships.len(), 2);
        for relationship in image_relationships {
            assert_eq!(relationship.target_mode, TargetMode::Internal);
            let image_uri = PartUri::new(DRAWING_PART_URI)
                .expect("drawing uri should be valid")
                .resolve_relative(relationship.target.as_str())
                .expect("image target should resolve");
            assert!(
                package.get_part(image_uri.as_str()).is_some(),
                "missing image part {}",
                image_uri.as_str()
            );
        }

        let image1_part = package
            .get_part("/xl/media/image1.png")
            .expect("first image part should exist");
        assert_eq!(image1_part.data.as_bytes(), &[137_u8, 80_u8, 78_u8]);
        assert_eq!(
            package.content_types().get_override("/xl/media/image1.png"),
            Some("image/png")
        );

        let image2_part = package
            .get_part("/xl/media/image2.jpeg")
            .expect("second image part should exist");
        assert_eq!(image2_part.data.as_bytes(), &[255_u8, 216_u8, 255_u8]);
        assert_eq!(
            package
                .content_types()
                .get_override("/xl/media/image2.jpeg"),
            Some("image/jpeg")
        );

        let loaded = Workbook::open(temp_file.path()).expect("open workbook");
        let loaded_sheet = loaded.sheet("Data").expect("sheet should exist");
        assert_eq!(loaded_sheet.images().len(), 2);
        assert_eq!(loaded_sheet.images()[0].bytes(), &[137_u8, 80_u8, 78_u8]);
        assert_eq!(loaded_sheet.images()[0].content_type(), "image/png");
        assert_eq!(loaded_sheet.images()[0].anchor_cell(), "B2");
        assert_eq!(loaded_sheet.images()[0].ext(), Some(custom_ext));
        assert_eq!(loaded_sheet.images()[1].bytes(), &[255_u8, 216_u8, 255_u8]);
        assert_eq!(loaded_sheet.images()[1].content_type(), "image/jpeg");
        assert_eq!(loaded_sheet.images()[1].anchor_cell(), "D5");
        assert_eq!(
            loaded_sheet.images()[1].ext(),
            Some(
                WorksheetImageExt::new(DEFAULT_IMAGE_EXTENT_CX, DEFAULT_IMAGE_EXTENT_CY)
                    .expect("default ext should be valid")
            )
        );
    }

    #[test]
    fn conditional_formattings_roundtrip_for_cell_is_and_expression() {
        let mut workbook = Workbook::new();
        let sheet = workbook.add_sheet("Data");
        sheet.add_conditional_formatting(
            ConditionalFormatting::cell_is(["A1:A10", "C1"], ["5", "10"])
                .expect("cellIs conditional formatting should be valid"),
        );
        sheet.add_conditional_formatting(
            ConditionalFormatting::expression(["D1:D10"], ["MOD(ROW(),2)=0"])
                .expect("expression conditional formatting should be valid"),
        );

        let temp_file = NamedTempFile::new().expect("temp file should be created");
        workbook.save(temp_file.path()).expect("save workbook");

        let package = Package::open(temp_file.path()).expect("open saved workbook package");
        let worksheet_part = package
            .get_part("/xl/worksheets/sheet1.xml")
            .expect("worksheet part should exist");
        let worksheet_xml = String::from_utf8_lossy(worksheet_part.data.as_bytes());
        assert!(worksheet_xml.contains("<conditionalFormatting sqref=\"A1:A10 C1\">"));
        assert!(worksheet_xml.contains("<cfRule type=\"cellIs\""));
        assert!(worksheet_xml.contains("<formula>5</formula>"));
        assert!(worksheet_xml.contains("<formula>10</formula>"));
        assert!(worksheet_xml.contains("<conditionalFormatting sqref=\"D1:D10\">"));
        assert!(worksheet_xml.contains("<cfRule type=\"expression\""));
        assert!(worksheet_xml.contains("<formula>MOD(ROW(),2)=0</formula>"));

        let loaded = Workbook::open(temp_file.path()).expect("open workbook");
        let loaded_sheet = loaded.sheet("Data").expect("sheet should exist");
        let rules = loaded_sheet.conditional_formattings();
        assert_eq!(rules.len(), 2);
        assert_eq!(rules[0].rule_type(), ConditionalFormattingRuleType::CellIs);
        assert_eq!(rules[0].sqref().len(), 2);
        assert_eq!(rules[0].formulas(), &["5".to_string(), "10".to_string()]);
        assert_eq!(
            rules[1].rule_type(),
            ConditionalFormattingRuleType::Expression
        );
        assert_eq!(rules[1].sqref().len(), 1);
        assert_eq!(rules[1].formulas(), &["MOD(ROW(),2)=0".to_string()]);
    }

    #[test]
    fn data_validations_roundtrip_for_basic_rule_types() {
        let mut workbook = Workbook::new();
        let sheet = workbook.add_sheet("Data");
        sheet.add_data_validation(
            DataValidation::list(["A1:A3", "C1"], "\"Yes,No\"")
                .expect("list validation should be valid"),
        );

        let mut whole =
            DataValidation::whole(["B1:B10"], "1").expect("whole validation should be valid");
        whole.set_formula2("10");
        sheet.add_data_validation(whole);

        let mut decimal =
            DataValidation::decimal(["D1:D5"], "0.5").expect("decimal validation should be valid");
        decimal.set_formula2("9.5");
        sheet.add_data_validation(decimal);

        let mut date = DataValidation::date(["E1:E2"], "DATE(2024,1,1)")
            .expect("date validation should be valid");
        date.set_formula2("DATE(2024,12,31)");
        sheet.add_data_validation(date);

        let mut text_length = DataValidation::text_length(["F1:F4", "H1:H2"], "3")
            .expect("text length validation should be valid");
        text_length.set_formula2("15");
        sheet.add_data_validation(text_length);

        let temp_file = NamedTempFile::new().expect("temp file should be created");
        workbook.save(temp_file.path()).expect("save workbook");

        let package = Package::open(temp_file.path()).expect("open saved workbook package");
        let worksheet_part = package
            .get_part("/xl/worksheets/sheet1.xml")
            .expect("worksheet part should exist");
        let worksheet_xml = String::from_utf8_lossy(worksheet_part.data.as_bytes());
        assert!(worksheet_xml.contains("<dataValidations count=\"5\">"));
        assert!(worksheet_xml.contains("type=\"list\""));
        assert!(worksheet_xml.contains("type=\"whole\""));
        assert!(worksheet_xml.contains("type=\"decimal\""));
        assert!(worksheet_xml.contains("type=\"date\""));
        assert!(worksheet_xml.contains("type=\"textLength\""));
        assert!(worksheet_xml.contains("sqref=\"A1:A3 C1\""));
        assert!(worksheet_xml.contains("<formula1>&quot;Yes,No&quot;</formula1>"));
        assert!(worksheet_xml.contains("<formula1>1</formula1>"));
        assert!(worksheet_xml.contains("<formula2>10</formula2>"));
        assert!(worksheet_xml.contains("<formula1>0.5</formula1>"));
        assert!(worksheet_xml.contains("<formula2>9.5</formula2>"));
        assert!(worksheet_xml.contains("<formula1>DATE(2024,1,1)</formula1>"));
        assert!(worksheet_xml.contains("<formula2>DATE(2024,12,31)</formula2>"));
        assert!(worksheet_xml.contains("sqref=\"F1:F4 H1:H2\""));

        let loaded = Workbook::open(temp_file.path()).expect("open workbook");
        let loaded_sheet = loaded.sheet("Data").expect("sheet should exist");
        let validations = loaded_sheet.data_validations();
        assert_eq!(validations.len(), 5);

        assert_eq!(validations[0].validation_type(), DataValidationType::List);
        assert_eq!(validations[0].sqref().len(), 2);
        assert_eq!(validations[0].formula1(), "\"Yes,No\"");
        assert_eq!(validations[0].formula2(), None);

        assert_eq!(validations[1].validation_type(), DataValidationType::Whole);
        assert_eq!(validations[1].formula1(), "1");
        assert_eq!(validations[1].formula2(), Some("10"));

        assert_eq!(
            validations[2].validation_type(),
            DataValidationType::Decimal
        );
        assert_eq!(validations[2].formula1(), "0.5");
        assert_eq!(validations[2].formula2(), Some("9.5"));

        assert_eq!(validations[3].validation_type(), DataValidationType::Date);
        assert_eq!(validations[3].formula1(), "DATE(2024,1,1)");
        assert_eq!(validations[3].formula2(), Some("DATE(2024,12,31)"));

        assert_eq!(
            validations[4].validation_type(),
            DataValidationType::TextLength
        );
        assert_eq!(validations[4].sqref().len(), 2);
        assert_eq!(validations[4].formula1(), "3");
        assert_eq!(validations[4].formula2(), Some("15"));
    }

    #[test]
    fn new_from_scratch_workbook_still_works() {
        let mut workbook = Workbook::new();
        let sheet = workbook.add_sheet("Sheet1");
        sheet.cell_mut("A1").expect("valid cell").set_value("Hello");
        sheet.cell_mut("B1").expect("valid cell").set_value(42_i32);

        let temp_file = NamedTempFile::new().expect("create temp file");
        workbook.save(temp_file.path()).expect("save new workbook");

        let package = Package::open(temp_file.path()).expect("reopen package");
        assert!(
            package
                .relationships()
                .get_first_by_type(RelationshipType::WORKBOOK)
                .is_some(),
            "package-level workbook relationship must exist for new workbooks"
        );

        let loaded = Workbook::open(temp_file.path()).expect("reopen workbook");
        let sheet = loaded.sheet("Sheet1").expect("Sheet1 exists");
        assert_eq!(
            sheet.cell("A1").and_then(|c| c.value()),
            Some(&CellValue::String("Hello".to_string()))
        );
        assert_eq!(
            sheet.cell("B1").and_then(|c| c.value()),
            Some(&CellValue::Number(42.0))
        );
    }

    #[test]
    fn mutating_opened_workbook_marks_it_dirty_and_persists_changes() {
        let mut workbook = Workbook::new();
        workbook
            .add_sheet("Data")
            .cell_mut("A1")
            .expect("valid cell")
            .set_value("before");

        let temp1 = NamedTempFile::new().expect("create temp file");
        workbook.save(temp1.path()).expect("save initial workbook");

        let mut loaded = Workbook::open(temp1.path()).expect("open workbook");
        assert!(
            !loaded.dirty,
            "open() should start from a pristine workbook state"
        );
        loaded
            .sheet_mut("Data")
            .expect("Data sheet should exist")
            .cell_mut("A1")
            .expect("valid cell")
            .set_value("after");

        let temp2 = NamedTempFile::new().expect("create temp file");
        loaded.save(temp2.path()).expect("save modified workbook");

        let reopened = Workbook::open(temp2.path()).expect("reopen workbook");
        assert_eq!(
            reopened
                .sheet("Data")
                .and_then(|sheet| sheet.cell("A1"))
                .and_then(Cell::value),
            Some(&CellValue::String("after".to_string()))
        );
    }

    #[test]
    fn roundtrip_preserves_unknown_parts_and_worksheet_children() {
        // Build a package with an extra part that the xlsx layer doesn't know about
        // (e.g. /xl/theme/theme1.xml) and an unknown worksheet child element.
        let mut workbook = Workbook::new();
        let sheet = workbook.add_sheet("Data");
        sheet.cell_mut("A1").expect("valid cell").set_value("test");

        let temp1 = NamedTempFile::new().expect("create temp file");
        workbook.save(temp1.path()).expect("save initial workbook");

        // Inject an extra part into the package and an unknown element into the worksheet XML.
        let mut package = Package::open(temp1.path()).expect("open package");
        let theme_uri = PartUri::new("/xl/theme/theme1.xml").expect("valid URI");
        let theme_part = Part::new_xml(
            theme_uri,
            b"<a:theme xmlns:a=\"mock\">contents</a:theme>".to_vec(),
        );
        package.set_part(theme_part);
        package
            .get_part_mut(WORKBOOK_PART_URI)
            .expect("workbook part")
            .relationships
            .add_new(
                CUSTOM_WORKBOOK_RELATIONSHIP_TYPE.to_string(),
                "theme/theme1.xml".to_string(),
                TargetMode::Internal,
            );

        // Modify the worksheet XML to include an unknown element.
        let ws_part = package
            .get_part_mut("/xl/worksheets/sheet1.xml")
            .expect("worksheet part");
        let original_xml = String::from_utf8_lossy(ws_part.data.as_bytes()).into_owned();
        let modified_xml = original_xml.replace(
            "</worksheet>",
            "<extLst><ext uri=\"{test-roundtrip}\">payload</ext></extLst></worksheet>",
        );
        ws_part.data = PartData::Xml(modified_xml.into_bytes());

        let temp2 = NamedTempFile::new().expect("create temp file");
        package.save(temp2.path()).expect("save modified package");

        // Open with offidized-xlsx and re-save.
        let loaded = Workbook::open(temp2.path()).expect("open modified workbook");
        let temp3 = NamedTempFile::new().expect("create temp file");
        loaded.save(temp3.path()).expect("roundtrip save");

        // Verify the extra part survived.
        let final_package = Package::open(temp3.path()).expect("open final package");
        assert!(
            final_package.get_part("/xl/theme/theme1.xml").is_some(),
            "unknown part /xl/theme/theme1.xml must survive roundtrip"
        );
        let final_workbook_part = final_package
            .get_part(WORKBOOK_PART_URI)
            .expect("workbook part should exist");
        let custom_relationships = final_workbook_part
            .relationships
            .get_by_type(CUSTOM_WORKBOOK_RELATIONSHIP_TYPE);
        assert_eq!(
            custom_relationships.len(),
            1,
            "custom workbook relationship must survive pristine open/save"
        );
        let theme_target_uri = PartUri::new(WORKBOOK_PART_URI)
            .expect("workbook uri should be valid")
            .resolve_relative(custom_relationships[0].target.as_str())
            .expect("theme target should resolve");
        assert_eq!(theme_target_uri.as_str(), "/xl/theme/theme1.xml");

        // Verify the unknown worksheet child element survived.
        let ws_part = final_package
            .get_part("/xl/worksheets/sheet1.xml")
            .expect("worksheet part");
        let ws_xml = String::from_utf8_lossy(ws_part.data.as_bytes());
        assert!(
            ws_xml.contains("extLst"),
            "unknown worksheet child <extLst> must survive roundtrip, got: {ws_xml}"
        );
        assert!(
            ws_xml.contains("{test-roundtrip}"),
            "unknown element content must survive roundtrip"
        );

        // Verify the workbook still opens correctly.
        let final_workbook = Workbook::open(temp3.path()).expect("reopen final workbook");
        let final_sheet = final_workbook.sheet("Data").expect("Data sheet exists");
        assert_eq!(
            final_sheet.cell("A1").and_then(|c| c.value()),
            Some(&CellValue::String("test".to_string()))
        );
    }

    #[test]
    fn dirty_save_preserves_unmanaged_workbook_and_worksheet_relationships() {
        let mut workbook = Workbook::new();
        let sheet = workbook.add_sheet("Data");
        sheet
            .cell_mut("A1")
            .expect("valid cell")
            .set_value("before");

        let baseline = NamedTempFile::new().expect("create baseline temp file");
        workbook
            .save(baseline.path())
            .expect("save baseline workbook");

        let mut package = Package::open(baseline.path()).expect("open baseline package");
        let workbook_custom_uri =
            PartUri::new("/xl/custom/workbook-meta1.xml").expect("valid workbook custom part URI");
        let mut workbook_custom_part = Part::new_xml(
            workbook_custom_uri.clone(),
            b"<custom workbook=\"meta\"/>".to_vec(),
        );
        workbook_custom_part.content_type = Some("application/xml".to_string());
        package.set_part(workbook_custom_part);
        package
            .get_part_mut(WORKBOOK_PART_URI)
            .expect("workbook part should exist")
            .relationships
            .add_new(
                CUSTOM_WORKBOOK_RELATIONSHIP_TYPE.to_string(),
                "custom/workbook-meta1.xml".to_string(),
                TargetMode::Internal,
            );

        let sheet_custom_uri =
            PartUri::new("/xl/custom/sheet-meta1.xml").expect("valid worksheet custom part URI");
        let mut sheet_custom_part = Part::new_xml(
            sheet_custom_uri.clone(),
            b"<custom sheet=\"meta\"/>".to_vec(),
        );
        sheet_custom_part.content_type = Some("application/xml".to_string());
        package.set_part(sheet_custom_part);
        package
            .get_part_mut("/xl/worksheets/sheet1.xml")
            .expect("worksheet part should exist")
            .relationships
            .add_new(
                CUSTOM_WORKSHEET_RELATIONSHIP_TYPE.to_string(),
                "../custom/sheet-meta1.xml".to_string(),
                TargetMode::Internal,
            );

        let injected = NamedTempFile::new().expect("create injected temp file");
        package
            .save(injected.path())
            .expect("save package with unmanaged relationships");

        let mut opened = Workbook::open(injected.path()).expect("open injected workbook");
        opened
            .sheet_mut("Data")
            .expect("Data sheet should exist")
            .cell_mut("A1")
            .expect("valid cell")
            .set_value("after");

        let roundtripped = NamedTempFile::new().expect("create roundtripped temp file");
        opened
            .save(roundtripped.path())
            .expect("save dirty roundtripped workbook");

        let final_package = Package::open(roundtripped.path()).expect("open final package");

        let workbook_part = final_package
            .get_part(WORKBOOK_PART_URI)
            .expect("workbook part should exist");
        let workbook_custom_relationships = workbook_part
            .relationships
            .get_by_type(CUSTOM_WORKBOOK_RELATIONSHIP_TYPE);
        assert_eq!(
            workbook_custom_relationships.len(),
            1,
            "unmanaged workbook relationship should survive dirty save as pass-through",
        );
        let workbook_custom_target = PartUri::new(WORKBOOK_PART_URI)
            .expect("workbook URI should parse")
            .resolve_relative(workbook_custom_relationships[0].target.as_str())
            .expect("workbook custom relationship target should resolve");
        assert_eq!(
            workbook_custom_target.as_str(),
            workbook_custom_uri.as_str()
        );
        assert!(
            final_package
                .get_part(workbook_custom_uri.as_str())
                .is_some(),
            "unmanaged workbook target part should survive dirty save",
        );

        let worksheet_part = final_package
            .get_part("/xl/worksheets/sheet1.xml")
            .expect("worksheet part should exist");
        let worksheet_custom_relationships = worksheet_part
            .relationships
            .get_by_type(CUSTOM_WORKSHEET_RELATIONSHIP_TYPE);
        assert_eq!(
            worksheet_custom_relationships.len(),
            1,
            "unmanaged worksheet relationship should survive dirty save as pass-through",
        );
        let worksheet_custom_target = PartUri::new("/xl/worksheets/sheet1.xml")
            .expect("worksheet URI should parse")
            .resolve_relative(worksheet_custom_relationships[0].target.as_str())
            .expect("worksheet custom relationship target should resolve");
        assert_eq!(worksheet_custom_target.as_str(), sheet_custom_uri.as_str());
        assert!(
            final_package.get_part(sheet_custom_uri.as_str()).is_some(),
            "unmanaged worksheet target part should survive dirty save",
        );

        let reopened = Workbook::open(roundtripped.path()).expect("reopen roundtripped workbook");
        assert_eq!(
            reopened
                .sheet("Data")
                .and_then(|s| s.cell("A1"))
                .and_then(Cell::value),
            Some(&CellValue::String("after".to_string())),
            "managed workbook mutation should still be persisted",
        );
    }

    #[test]
    fn roundtrip_preserves_unknown_row_and_cell_attrs_and_children() {
        let workbook_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Data" sheetId="1" r:id="rId1"/>
  </sheets>
</workbook>"#;
        let worksheet_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
           xmlns:foo="urn:offidized:test">
  <sheetData>
    <row r="1" ht="15" foo:rowAttr="keep">
      <c r="A1" t="str" foo:cellAttr="retain">
        <v>hello</v>
        <foo:cellExtra marker="yes"/>
      </c>
      <foo:rowExtra marker="yes"/>
    </row>
  </sheetData>
</worksheet>"#;

        let mut package = Package::new();
        let workbook_uri = PartUri::new(WORKBOOK_PART_URI).expect("workbook uri should be valid");
        let mut workbook_part = Part::new_xml(workbook_uri, workbook_xml.as_bytes().to_vec());
        workbook_part.content_type = Some(ContentTypeValue::WORKBOOK.to_string());
        workbook_part.relationships.add_new(
            RelationshipType::WORKSHEET.to_string(),
            "worksheets/sheet1.xml".to_string(),
            TargetMode::Internal,
        );
        package.set_part(workbook_part);

        let mut worksheet_part = Part::new_xml(
            PartUri::new("/xl/worksheets/sheet1.xml").expect("worksheet uri should be valid"),
            worksheet_xml.as_bytes().to_vec(),
        );
        worksheet_part.content_type = Some(ContentTypeValue::WORKSHEET.to_string());
        package.set_part(worksheet_part);

        package.relationships_mut().add_new(
            RelationshipType::WORKBOOK.to_string(),
            WORKBOOK_PART_URI.to_string(),
            TargetMode::Internal,
        );

        let input = NamedTempFile::new().expect("temp input should be created");
        package.save(input.path()).expect("save input package");

        let loaded = Workbook::open(input.path()).expect("open workbook");
        let output = NamedTempFile::new().expect("temp output should be created");
        loaded.save(output.path()).expect("save workbook");

        let out_package = Package::open(output.path()).expect("open output package");
        let out_sheet = out_package
            .get_part("/xl/worksheets/sheet1.xml")
            .expect("sheet1 should exist");
        let out_xml = String::from_utf8_lossy(out_sheet.data.as_bytes());
        assert!(out_xml.contains("foo:rowAttr=\"keep\""));
        assert!(out_xml.contains("foo:cellAttr=\"retain\""));
        assert!(out_xml.contains("<foo:rowExtra"));
        assert!(out_xml.contains("<foo:cellExtra"));
    }

    #[test]
    fn dirty_save_passthroughs_clean_worksheet_bytes() {
        let mut workbook = Workbook::new();
        workbook
            .add_sheet("Keep")
            .cell_mut("A1")
            .expect("valid cell")
            .set_value("keep");
        workbook
            .add_sheet("Change")
            .cell_mut("A1")
            .expect("valid cell")
            .set_value("before");

        let baseline = NamedTempFile::new().expect("temp baseline should be created");
        workbook
            .save(baseline.path())
            .expect("save baseline workbook");

        let mut injected_package =
            Package::open(baseline.path()).expect("open baseline package for injection");
        let keep_sheet_part = injected_package
            .get_part("/xl/worksheets/sheet1.xml")
            .expect("keep sheet should exist");
        let keep_sheet_xml = String::from_utf8_lossy(keep_sheet_part.data.as_bytes()).into_owned();
        let keep_sheet_xml = keep_sheet_xml.replacen(
            "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">",
            "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:foo=\"urn:offidized:test\" foo:keep=\"1\">",
            1,
        );
        injected_package
            .get_part_mut("/xl/worksheets/sheet1.xml")
            .expect("keep sheet should exist")
            .data = PartData::Xml(keep_sheet_xml.into_bytes());

        let injected = NamedTempFile::new().expect("temp injected should be created");
        injected_package
            .save(injected.path())
            .expect("save injected package");

        let injected_keep_bytes = Package::open(injected.path())
            .expect("reopen injected package")
            .get_part("/xl/worksheets/sheet1.xml")
            .expect("keep sheet should exist")
            .data
            .as_bytes()
            .to_vec();

        let mut opened = Workbook::open(injected.path()).expect("open injected workbook");
        opened
            .sheet_mut("Change")
            .expect("change sheet should exist")
            .cell_mut("A1")
            .expect("valid cell")
            .set_value("after");

        let roundtripped = NamedTempFile::new().expect("temp roundtripped should be created");
        opened
            .save(roundtripped.path())
            .expect("save roundtripped workbook");

        let final_keep_bytes = Package::open(roundtripped.path())
            .expect("open roundtripped package")
            .get_part("/xl/worksheets/sheet1.xml")
            .expect("keep sheet should exist")
            .data
            .as_bytes()
            .to_vec();
        assert_eq!(
            final_keep_bytes, injected_keep_bytes,
            "clean worksheet part bytes should be passed through unchanged on dirty workbook saves",
        );
    }

    #[test]
    fn row_hidden_roundtrips_through_worksheet_xml() {
        let mut workbook = Workbook::new();
        let sheet = workbook.add_sheet("Data");

        sheet
            .cell_mut("A1")
            .expect("cell should be valid")
            .set_value("visible");
        sheet
            .cell_mut("A2")
            .expect("cell should be valid")
            .set_value("hidden");
        sheet
            .row_mut(2)
            .expect("row should be valid")
            .set_hidden(true);
        sheet
            .cell_mut("A3")
            .expect("cell should be valid")
            .set_value("also visible");

        let temp_file = NamedTempFile::new().expect("temp file should be created");
        workbook.save(temp_file.path()).expect("save workbook");

        let package = Package::open(temp_file.path()).expect("open saved package");
        let worksheet_xml = String::from_utf8_lossy(
            package
                .get_part("/xl/worksheets/sheet1.xml")
                .expect("worksheet part should exist")
                .data
                .as_bytes(),
        )
        .into_owned();
        assert!(
            worksheet_xml.contains("hidden=\"1\""),
            "hidden attribute should be present in worksheet XML"
        );

        let loaded = Workbook::open(temp_file.path()).expect("open workbook");
        let loaded_sheet = loaded.sheet("Data").expect("sheet should exist");

        assert!(
            !loaded_sheet.row(1).map(|r| r.is_hidden()).unwrap_or(false),
            "row 1 should not be hidden"
        );
        assert!(
            loaded_sheet.row(2).expect("row 2 should exist").is_hidden(),
            "row 2 should be hidden"
        );
        assert!(
            !loaded_sheet.row(3).map(|r| r.is_hidden()).unwrap_or(false),
            "row 3 should not be hidden"
        );
    }

    #[test]
    fn column_hidden_roundtrips_through_worksheet_xml() {
        let mut workbook = Workbook::new();
        let sheet = workbook.add_sheet("Data");

        sheet
            .cell_mut("A1")
            .expect("cell should be valid")
            .set_value("col A");
        sheet
            .cell_mut("B1")
            .expect("cell should be valid")
            .set_value("col B");
        sheet
            .column_mut(2)
            .expect("column should be valid")
            .set_hidden(true);

        let temp_file = NamedTempFile::new().expect("temp file should be created");
        workbook.save(temp_file.path()).expect("save workbook");

        let package = Package::open(temp_file.path()).expect("open saved package");
        let worksheet_xml = String::from_utf8_lossy(
            package
                .get_part("/xl/worksheets/sheet1.xml")
                .expect("worksheet part should exist")
                .data
                .as_bytes(),
        )
        .into_owned();
        assert!(
            worksheet_xml.contains("<cols>"),
            "cols section should be present"
        );
        assert!(
            worksheet_xml.contains("hidden=\"1\""),
            "hidden attribute should be present for column"
        );

        let loaded = Workbook::open(temp_file.path()).expect("open workbook");
        let loaded_sheet = loaded.sheet("Data").expect("sheet should exist");

        assert!(
            !loaded_sheet
                .column(1)
                .map(|c| c.is_hidden())
                .unwrap_or(false),
            "column 1 should not be hidden"
        );
        assert!(
            loaded_sheet
                .column(2)
                .expect("column 2 should exist")
                .is_hidden(),
            "column 2 should be hidden"
        );
    }

    #[test]
    fn column_width_roundtrips_through_cols_xml() {
        let mut workbook = Workbook::new();
        let sheet = workbook.add_sheet("Data");

        sheet
            .cell_mut("A1")
            .expect("cell should be valid")
            .set_value("hello");
        sheet
            .column_mut(1)
            .expect("column should be valid")
            .set_width(25.5);

        let temp_file = NamedTempFile::new().expect("temp file should be created");
        workbook.save(temp_file.path()).expect("save workbook");

        let loaded = Workbook::open(temp_file.path()).expect("open workbook");
        let loaded_sheet = loaded.sheet("Data").expect("sheet should exist");

        assert_eq!(
            loaded_sheet
                .column(1)
                .expect("column 1 should exist")
                .width(),
            Some(25.5),
            "column width should roundtrip"
        );
    }

    #[test]
    fn external_hyperlink_roundtrips_through_worksheet_xml() {
        use crate::worksheet::Hyperlink;

        let mut workbook = Workbook::new();
        let sheet = workbook.add_sheet("Data");

        sheet
            .cell_mut("A1")
            .expect("cell should be valid")
            .set_value("Click me");

        let mut hl =
            Hyperlink::external("A1", "https://example.com").expect("hyperlink should be valid");
        hl.set_tooltip("Example");
        sheet.add_hyperlink(hl);

        let temp_file = NamedTempFile::new().expect("temp file should be created");
        workbook.save(temp_file.path()).expect("save workbook");

        let package = Package::open(temp_file.path()).expect("open saved package");
        let worksheet_xml = String::from_utf8_lossy(
            package
                .get_part("/xl/worksheets/sheet1.xml")
                .expect("worksheet part should exist")
                .data
                .as_bytes(),
        )
        .into_owned();
        assert!(
            worksheet_xml.contains("<hyperlinks>"),
            "hyperlinks section should be present"
        );
        assert!(
            worksheet_xml.contains("r:id="),
            "hyperlink should have r:id for external link"
        );
        assert!(
            worksheet_xml.contains("tooltip=\"Example\""),
            "hyperlink should have tooltip"
        );

        // Verify the relationship points to the URL.
        let sheet_part = package
            .get_part("/xl/worksheets/sheet1.xml")
            .expect("worksheet part should exist");
        let has_hyperlink_rel = sheet_part.relationships.iter().any(|rel| {
            rel.rel_type == HYPERLINK_RELATIONSHIP_TYPE
                && rel.target == "https://example.com"
                && rel.target_mode == TargetMode::External
        });
        assert!(
            has_hyperlink_rel,
            "worksheet should have an external hyperlink relationship"
        );

        let loaded = Workbook::open(temp_file.path()).expect("open workbook");
        let loaded_sheet = loaded.sheet("Data").expect("sheet should exist");

        assert_eq!(loaded_sheet.hyperlinks().len(), 1);
        assert_eq!(loaded_sheet.hyperlinks()[0].cell_ref(), "A1");
        assert_eq!(
            loaded_sheet.hyperlinks()[0].url(),
            Some("https://example.com")
        );
        assert_eq!(loaded_sheet.hyperlinks()[0].tooltip(), Some("Example"));
        assert!(loaded_sheet.hyperlinks()[0].location().is_none());
    }

    #[test]
    fn internal_hyperlink_roundtrips_through_worksheet_xml() {
        use crate::worksheet::Hyperlink;

        let mut workbook = Workbook::new();
        let sheet = workbook.add_sheet("Data");

        sheet
            .cell_mut("B2")
            .expect("cell should be valid")
            .set_value("Go to Sheet2");
        let hl = Hyperlink::internal("B2", "Sheet2!A1").expect("hyperlink should be valid");
        sheet.add_hyperlink(hl);

        let temp_file = NamedTempFile::new().expect("temp file should be created");
        workbook.save(temp_file.path()).expect("save workbook");

        let package = Package::open(temp_file.path()).expect("open saved package");
        let worksheet_xml = String::from_utf8_lossy(
            package
                .get_part("/xl/worksheets/sheet1.xml")
                .expect("worksheet part should exist")
                .data
                .as_bytes(),
        )
        .into_owned();
        assert!(
            worksheet_xml.contains("location=\"Sheet2!A1\""),
            "hyperlink should have location attribute"
        );
        // Internal hyperlink should NOT have r:id.
        assert!(
            !worksheet_xml.contains("r:id="),
            "internal hyperlink should not have r:id"
        );

        let loaded = Workbook::open(temp_file.path()).expect("open workbook");
        let loaded_sheet = loaded.sheet("Data").expect("sheet should exist");

        assert_eq!(loaded_sheet.hyperlinks().len(), 1);
        assert_eq!(loaded_sheet.hyperlinks()[0].cell_ref(), "B2");
        assert!(loaded_sheet.hyperlinks()[0].url().is_none());
        assert_eq!(loaded_sheet.hyperlinks()[0].location(), Some("Sheet2!A1"));
    }

    #[test]
    fn mixed_hyperlinks_roundtrip_through_worksheet_xml() {
        use crate::worksheet::Hyperlink;

        let mut workbook = Workbook::new();
        let sheet = workbook.add_sheet("Data");

        sheet
            .cell_mut("A1")
            .expect("cell should be valid")
            .set_value("External");
        sheet
            .cell_mut("A2")
            .expect("cell should be valid")
            .set_value("Internal");
        sheet
            .cell_mut("A3")
            .expect("cell should be valid")
            .set_value("External 2");

        sheet.add_hyperlink(
            Hyperlink::external("A1", "https://example.com/first")
                .expect("hyperlink should be valid"),
        );
        sheet.add_hyperlink(
            Hyperlink::internal("A2", "Sheet2!B5").expect("hyperlink should be valid"),
        );
        sheet.add_hyperlink(
            Hyperlink::external("A3", "https://example.com/second")
                .expect("hyperlink should be valid"),
        );

        let temp_file = NamedTempFile::new().expect("temp file should be created");
        workbook.save(temp_file.path()).expect("save workbook");

        let loaded = Workbook::open(temp_file.path()).expect("open workbook");
        let loaded_sheet = loaded.sheet("Data").expect("sheet should exist");

        assert_eq!(loaded_sheet.hyperlinks().len(), 3);
        assert_eq!(
            loaded_sheet.hyperlinks()[0].url(),
            Some("https://example.com/first")
        );
        assert!(loaded_sheet.hyperlinks()[0].location().is_none());
        assert!(loaded_sheet.hyperlinks()[1].url().is_none());
        assert_eq!(loaded_sheet.hyperlinks()[1].location(), Some("Sheet2!B5"));
        assert_eq!(
            loaded_sheet.hyperlinks()[2].url(),
            Some("https://example.com/second")
        );
    }

    // ===== Feature 1: Sheet visibility roundtrip =====

    #[test]
    fn sheet_visibility_roundtrips_through_workbook_xml() {
        use crate::worksheet::SheetVisibility;

        let mut workbook = Workbook::new();
        let sheet1 = workbook.add_sheet("Visible");
        sheet1
            .cell_mut("A1")
            .expect("cell should be valid")
            .set_value("v");

        let sheet2 = workbook.add_sheet("Hidden");
        sheet2
            .cell_mut("A1")
            .expect("cell should be valid")
            .set_value("h");
        sheet2.set_visibility(SheetVisibility::Hidden);

        let sheet3 = workbook.add_sheet("VeryHidden");
        sheet3
            .cell_mut("A1")
            .expect("cell should be valid")
            .set_value("vh");
        sheet3.set_visibility(SheetVisibility::VeryHidden);

        let temp_file = NamedTempFile::new().expect("temp file should be created");
        workbook.save(temp_file.path()).expect("save workbook");

        let package = Package::open(temp_file.path()).expect("open saved package");
        let workbook_xml = String::from_utf8_lossy(
            package
                .get_part(WORKBOOK_PART_URI)
                .expect("workbook part should exist")
                .data
                .as_bytes(),
        )
        .into_owned();
        assert!(
            workbook_xml.contains("state=\"hidden\""),
            "hidden sheet should have state attribute in workbook XML"
        );
        assert!(
            workbook_xml.contains("state=\"veryHidden\""),
            "very hidden sheet should have state attribute in workbook XML"
        );

        let loaded = Workbook::open(temp_file.path()).expect("open workbook");
        assert_eq!(
            loaded.sheet("Visible").unwrap().visibility(),
            SheetVisibility::Visible
        );
        assert_eq!(
            loaded.sheet("Hidden").unwrap().visibility(),
            SheetVisibility::Hidden
        );
        assert_eq!(
            loaded.sheet("VeryHidden").unwrap().visibility(),
            SheetVisibility::VeryHidden
        );
    }

    // ===== Feature 2: Sheet protection roundtrip =====

    #[test]
    fn sheet_protection_roundtrips_through_worksheet_xml() {
        use crate::worksheet::SheetProtection;

        let mut workbook = Workbook::new();
        let sheet = workbook.add_sheet("Data");
        sheet
            .cell_mut("A1")
            .expect("cell should be valid")
            .set_value("protected");

        let mut protection = SheetProtection::new();
        protection
            .set_objects(true)
            .set_sort(true)
            .set_auto_filter(true)
            .set_password_hash("ABCD");
        sheet.set_protection(protection);

        let temp_file = NamedTempFile::new().expect("temp file should be created");
        workbook.save(temp_file.path()).expect("save workbook");

        let package = Package::open(temp_file.path()).expect("open saved package");
        let worksheet_xml = String::from_utf8_lossy(
            package
                .get_part("/xl/worksheets/sheet1.xml")
                .expect("worksheet part should exist")
                .data
                .as_bytes(),
        )
        .into_owned();
        assert!(
            worksheet_xml.contains("<sheetProtection"),
            "sheetProtection element should be present"
        );
        assert!(
            worksheet_xml.contains("sheet=\"1\""),
            "sheet attribute should be present"
        );
        assert!(
            worksheet_xml.contains("objects=\"1\""),
            "objects attribute should be present"
        );
        assert!(
            worksheet_xml.contains("sort=\"1\""),
            "sort attribute should be present"
        );
        assert!(
            worksheet_xml.contains("autoFilter=\"1\""),
            "autoFilter attribute should be present"
        );
        assert!(
            worksheet_xml.contains("password=\"ABCD\""),
            "password attribute should be present"
        );

        let loaded = Workbook::open(temp_file.path()).expect("open workbook");
        let loaded_sheet = loaded.sheet("Data").expect("sheet should exist");
        let loaded_protection = loaded_sheet
            .protection()
            .expect("protection should be loaded");
        assert!(loaded_protection.sheet());
        assert!(loaded_protection.objects());
        assert!(loaded_protection.sort());
        assert!(loaded_protection.auto_filter());
        assert_eq!(loaded_protection.password_hash(), Some("ABCD"));
    }

    // ===== Feature 3: Page setup and margins roundtrip =====

    #[test]
    fn page_setup_roundtrips_through_worksheet_xml() {
        use crate::worksheet::{PageOrientation, PageSetup};

        let mut workbook = Workbook::new();
        let sheet = workbook.add_sheet("Data");
        sheet
            .cell_mut("A1")
            .expect("cell should be valid")
            .set_value("test");

        let mut setup = PageSetup::new();
        setup
            .set_orientation(PageOrientation::Landscape)
            .set_paper_size(9)
            .set_scale(85)
            .set_fit_to_width(1)
            .set_fit_to_height(0)
            .set_first_page_number(3);
        sheet.set_page_setup(setup);

        let temp_file = NamedTempFile::new().expect("temp file should be created");
        workbook.save(temp_file.path()).expect("save workbook");

        let package = Package::open(temp_file.path()).expect("open saved package");
        let worksheet_xml = String::from_utf8_lossy(
            package
                .get_part("/xl/worksheets/sheet1.xml")
                .expect("worksheet part should exist")
                .data
                .as_bytes(),
        )
        .into_owned();
        assert!(
            worksheet_xml.contains("<pageSetup"),
            "pageSetup element should be present"
        );
        assert!(
            worksheet_xml.contains("orientation=\"landscape\""),
            "orientation attribute should be present"
        );
        assert!(
            worksheet_xml.contains("paperSize=\"9\""),
            "paperSize attribute should be present"
        );
        assert!(
            worksheet_xml.contains("scale=\"85\""),
            "scale attribute should be present"
        );

        let loaded = Workbook::open(temp_file.path()).expect("open workbook");
        let loaded_sheet = loaded.sheet("Data").expect("sheet should exist");
        let loaded_setup = loaded_sheet
            .page_setup()
            .expect("page setup should be loaded");
        assert_eq!(loaded_setup.orientation(), Some(PageOrientation::Landscape));
        assert_eq!(loaded_setup.paper_size(), Some(9));
        assert_eq!(loaded_setup.scale(), Some(85));
        assert_eq!(loaded_setup.fit_to_width(), Some(1));
        assert_eq!(loaded_setup.fit_to_height(), Some(0));
        assert_eq!(loaded_setup.first_page_number(), Some(3));
    }

    #[test]
    fn page_margins_roundtrip_through_worksheet_xml() {
        use crate::worksheet::PageMargins;

        let mut workbook = Workbook::new();
        let sheet = workbook.add_sheet("Data");
        sheet
            .cell_mut("A1")
            .expect("cell should be valid")
            .set_value("test");

        let mut margins = PageMargins::new();
        margins
            .set_left(0.7)
            .set_right(0.7)
            .set_top(0.75)
            .set_bottom(0.75)
            .set_header(0.3)
            .set_footer(0.3);
        sheet.set_page_margins(margins);

        let temp_file = NamedTempFile::new().expect("temp file should be created");
        workbook.save(temp_file.path()).expect("save workbook");

        let package = Package::open(temp_file.path()).expect("open saved package");
        let worksheet_xml = String::from_utf8_lossy(
            package
                .get_part("/xl/worksheets/sheet1.xml")
                .expect("worksheet part should exist")
                .data
                .as_bytes(),
        )
        .into_owned();
        assert!(
            worksheet_xml.contains("<pageMargins"),
            "pageMargins element should be present"
        );
        assert!(
            worksheet_xml.contains("left=\"0.7\""),
            "left margin should be present"
        );
        assert!(
            worksheet_xml.contains("right=\"0.7\""),
            "right margin should be present"
        );
        assert!(
            worksheet_xml.contains("top=\"0.75\""),
            "top margin should be present"
        );

        let loaded = Workbook::open(temp_file.path()).expect("open workbook");
        let loaded_sheet = loaded.sheet("Data").expect("sheet should exist");
        let loaded_margins = loaded_sheet
            .page_margins()
            .expect("page margins should be loaded");
        assert_eq!(loaded_margins.left(), Some(0.7));
        assert_eq!(loaded_margins.right(), Some(0.7));
        assert_eq!(loaded_margins.top(), Some(0.75));
        assert_eq!(loaded_margins.bottom(), Some(0.75));
        assert_eq!(loaded_margins.header(), Some(0.3));
        assert_eq!(loaded_margins.footer(), Some(0.3));
    }

    // ===== Feature 7: Row/column grouping outline levels roundtrip =====

    #[test]
    fn row_outline_level_roundtrips_through_worksheet_xml() {
        let mut workbook = Workbook::new();
        let sheet = workbook.add_sheet("Data");

        sheet
            .cell_mut("A1")
            .expect("cell should be valid")
            .set_value("level0");
        sheet
            .cell_mut("A2")
            .expect("cell should be valid")
            .set_value("level2");
        sheet
            .row_mut(2)
            .expect("row should be valid")
            .set_outline_level(2)
            .set_collapsed(true);

        let temp_file = NamedTempFile::new().expect("temp file should be created");
        workbook.save(temp_file.path()).expect("save workbook");

        let package = Package::open(temp_file.path()).expect("open saved package");
        let worksheet_xml = String::from_utf8_lossy(
            package
                .get_part("/xl/worksheets/sheet1.xml")
                .expect("worksheet part should exist")
                .data
                .as_bytes(),
        )
        .into_owned();
        assert!(
            worksheet_xml.contains("outlineLevel=\"2\""),
            "outlineLevel attribute should be present for row"
        );
        assert!(
            worksheet_xml.contains("collapsed=\"1\""),
            "collapsed attribute should be present for row"
        );

        let loaded = Workbook::open(temp_file.path()).expect("open workbook");
        let loaded_sheet = loaded.sheet("Data").expect("sheet should exist");
        let row2 = loaded_sheet.row(2).expect("row 2 should exist");
        assert_eq!(row2.outline_level(), 2);
        assert!(row2.is_collapsed());
    }

    #[test]
    fn column_outline_level_roundtrips_through_worksheet_xml() {
        let mut workbook = Workbook::new();
        let sheet = workbook.add_sheet("Data");

        sheet
            .cell_mut("A1")
            .expect("cell should be valid")
            .set_value("test");
        sheet
            .column_mut(2)
            .expect("column should be valid")
            .set_outline_level(3)
            .set_collapsed(true);

        let temp_file = NamedTempFile::new().expect("temp file should be created");
        workbook.save(temp_file.path()).expect("save workbook");

        let package = Package::open(temp_file.path()).expect("open saved package");
        let worksheet_xml = String::from_utf8_lossy(
            package
                .get_part("/xl/worksheets/sheet1.xml")
                .expect("worksheet part should exist")
                .data
                .as_bytes(),
        )
        .into_owned();
        assert!(
            worksheet_xml.contains("outlineLevel=\"3\""),
            "outlineLevel attribute should be present for column"
        );
        assert!(
            worksheet_xml.contains("collapsed=\"1\""),
            "collapsed attribute should be present for column"
        );

        let loaded = Workbook::open(temp_file.path()).expect("open workbook");
        let loaded_sheet = loaded.sheet("Data").expect("sheet should exist");
        let col2 = loaded_sheet.column(2).expect("column 2 should exist");
        assert_eq!(col2.outline_level(), 3);
        assert!(col2.is_collapsed());
    }

    // ===== Feature 8: Sheet view options roundtrip =====

    #[test]
    fn sheet_view_options_roundtrip_through_worksheet_xml() {
        use crate::worksheet::SheetViewOptions;

        let mut workbook = Workbook::new();
        let sheet = workbook.add_sheet("Data");
        sheet
            .cell_mut("A1")
            .expect("cell should be valid")
            .set_value("test");

        let mut options = SheetViewOptions::new();
        options
            .set_show_gridlines(false)
            .set_show_row_col_headers(false)
            .set_show_formulas(true)
            .set_zoom_scale(150)
            .set_zoom_scale_normal(100)
            .set_right_to_left(true)
            .set_tab_selected(true);
        sheet.set_sheet_view_options(options);

        let temp_file = NamedTempFile::new().expect("temp file should be created");
        workbook.save(temp_file.path()).expect("save workbook");

        let package = Package::open(temp_file.path()).expect("open saved package");
        let worksheet_xml = String::from_utf8_lossy(
            package
                .get_part("/xl/worksheets/sheet1.xml")
                .expect("worksheet part should exist")
                .data
                .as_bytes(),
        )
        .into_owned();
        assert!(
            worksheet_xml.contains("showGridLines=\"0\""),
            "showGridLines attribute should be present"
        );
        assert!(
            worksheet_xml.contains("showRowColHeaders=\"0\""),
            "showRowColHeaders attribute should be present"
        );
        assert!(
            worksheet_xml.contains("showFormulas=\"1\""),
            "showFormulas attribute should be present"
        );
        assert!(
            worksheet_xml.contains("zoomScale=\"150\""),
            "zoomScale attribute should be present"
        );
        assert!(
            worksheet_xml.contains("rightToLeft=\"1\""),
            "rightToLeft attribute should be present"
        );
        assert!(
            worksheet_xml.contains("tabSelected=\"1\""),
            "tabSelected attribute should be present"
        );

        let loaded = Workbook::open(temp_file.path()).expect("open workbook");
        let loaded_sheet = loaded.sheet("Data").expect("sheet should exist");
        let loaded_options = loaded_sheet
            .sheet_view_options()
            .expect("sheet view options should be loaded");
        assert_eq!(loaded_options.show_gridlines(), Some(false));
        assert_eq!(loaded_options.show_row_col_headers(), Some(false));
        assert_eq!(loaded_options.show_formulas(), Some(true));
        assert_eq!(loaded_options.zoom_scale(), Some(150));
        assert_eq!(loaded_options.zoom_scale_normal(), Some(100));
        assert_eq!(loaded_options.right_to_left(), Some(true));
        assert_eq!(loaded_options.tab_selected(), Some(true));
    }

    // ===== Feature 10: Alignment indent and text rotation roundtrip =====

    #[test]
    fn alignment_indent_and_text_rotation_roundtrip_through_styles() {
        let mut workbook = Workbook::new();

        let mut alignment = Alignment::new();
        alignment
            .set_horizontal(HorizontalAlignment::Left)
            .set_indent(3)
            .set_text_rotation(45)
            .set_shrink_to_fit(true)
            .set_reading_order(1);

        let mut style = Style::new();
        style.set_alignment(alignment);
        let style_id = workbook
            .add_style(style)
            .expect("style id should fit within u32");

        let sheet = workbook.add_sheet("Data");
        sheet
            .cell_mut("A1")
            .expect("cell should be valid")
            .set_value("styled")
            .set_style_id(style_id);

        let temp_file = NamedTempFile::new().expect("temp file should be created");
        workbook.save(temp_file.path()).expect("save workbook");

        let package = Package::open(temp_file.path()).expect("open saved package");
        let styles_xml = String::from_utf8_lossy(
            package
                .get_part(STYLES_PART_URI)
                .expect("styles part should exist")
                .data
                .as_bytes(),
        )
        .into_owned();
        assert!(
            styles_xml.contains("indent=\"3\""),
            "indent attribute should be present in styles XML"
        );
        assert!(
            styles_xml.contains("textRotation=\"45\""),
            "textRotation attribute should be present in styles XML"
        );
        assert!(
            styles_xml.contains("shrinkToFit=\"1\""),
            "shrinkToFit attribute should be present in styles XML"
        );
        assert!(
            styles_xml.contains("readingOrder=\"1\""),
            "readingOrder attribute should be present in styles XML"
        );

        let loaded = Workbook::open(temp_file.path()).expect("open workbook");
        let loaded_style = loaded.style(style_id).expect("style should be loaded");
        let loaded_alignment = loaded_style
            .alignment()
            .expect("alignment should be loaded");
        assert_eq!(loaded_alignment.indent(), Some(3));
        assert_eq!(loaded_alignment.text_rotation(), Some(45));
        assert_eq!(loaded_alignment.shrink_to_fit(), Some(true));
        assert_eq!(loaded_alignment.reading_order(), Some(1));
        assert_eq!(
            loaded_alignment.horizontal(),
            Some(HorizontalAlignment::Left)
        );
    }

    // ===== Feature 11: Border diagonal roundtrip =====

    #[test]
    fn border_diagonal_roundtrips_through_styles() {
        let mut workbook = Workbook::new();

        let mut diag_side = BorderSide::new();
        diag_side.set_style("thin").set_color("FFFF0000");

        let mut border = Border::new();
        border
            .set_diagonal(diag_side)
            .set_diagonal_up(true)
            .set_diagonal_down(false);

        let mut style = Style::new();
        style.set_border(border);
        let style_id = workbook
            .add_style(style)
            .expect("style id should fit within u32");

        let sheet = workbook.add_sheet("Data");
        sheet
            .cell_mut("A1")
            .expect("cell should be valid")
            .set_value("diagonal")
            .set_style_id(style_id);

        let temp_file = NamedTempFile::new().expect("temp file should be created");
        workbook.save(temp_file.path()).expect("save workbook");

        let package = Package::open(temp_file.path()).expect("open saved package");
        let styles_xml = String::from_utf8_lossy(
            package
                .get_part(STYLES_PART_URI)
                .expect("styles part should exist")
                .data
                .as_bytes(),
        )
        .into_owned();
        assert!(
            styles_xml.contains("diagonalUp=\"1\""),
            "diagonalUp attribute should be present in styles XML"
        );
        assert!(
            styles_xml.contains("diagonalDown=\"0\""),
            "diagonalDown attribute should be present in styles XML"
        );
        assert!(
            styles_xml.contains("<diagonal style=\"thin\">"),
            "diagonal element should be present in styles XML"
        );
        assert!(
            styles_xml.contains("rgb=\"FFFF0000\""),
            "diagonal color should be present in styles XML"
        );

        let loaded = Workbook::open(temp_file.path()).expect("open workbook");
        let loaded_style = loaded.style(style_id).expect("style should be loaded");
        let loaded_border = loaded_style.border().expect("border should be loaded");
        assert_eq!(loaded_border.diagonal_up(), Some(true));
        assert_eq!(loaded_border.diagonal_down(), Some(false));
        let loaded_diag = loaded_border
            .diagonal()
            .expect("diagonal side should be loaded");
        assert_eq!(loaded_diag.style(), Some("thin"));
        assert_eq!(loaded_diag.color(), Some("FFFF0000"));
    }

    // ===== Feature 12: Workbook protection roundtrip =====

    #[test]
    fn workbook_protection_accessors_work() {
        let mut workbook = Workbook::new();
        assert!(workbook.workbook_protection().is_none());

        let mut prot = WorkbookProtection::new();
        assert!(prot.lock_structure());
        assert!(!prot.lock_windows());
        assert!(prot.password_hash().is_none());

        prot.set_lock_windows(true).set_password_hash("HASH123");
        assert!(prot.lock_windows());
        assert_eq!(prot.password_hash(), Some("HASH123"));

        workbook.set_workbook_protection(prot);
        assert!(workbook.workbook_protection().is_some());

        workbook.clear_workbook_protection();
        assert!(workbook.workbook_protection().is_none());
    }

    #[test]
    fn workbook_protection_roundtrips_through_workbook_xml() {
        let mut workbook = Workbook::new();
        workbook.add_sheet("Data");

        let mut prot = WorkbookProtection::new();
        prot.set_lock_windows(true).set_password_hash("ABCDEF");
        workbook.set_workbook_protection(prot);

        let temp_file = NamedTempFile::new().expect("temp file should be created");
        workbook.save(temp_file.path()).expect("save workbook");

        let package = Package::open(temp_file.path()).expect("open saved package");
        let workbook_xml = String::from_utf8_lossy(
            package
                .get_part(WORKBOOK_PART_URI)
                .expect("workbook part should exist")
                .data
                .as_bytes(),
        )
        .into_owned();
        assert!(
            workbook_xml.contains("<workbookProtection"),
            "workbookProtection element should be present"
        );
        assert!(
            workbook_xml.contains("lockStructure=\"1\""),
            "lockStructure attribute should be present"
        );
        assert!(
            workbook_xml.contains("lockWindows=\"1\""),
            "lockWindows attribute should be present"
        );
        assert!(
            workbook_xml.contains("workbookPassword=\"ABCDEF\""),
            "workbookPassword attribute should be present"
        );

        let loaded = Workbook::open(temp_file.path()).expect("open workbook");
        let loaded_prot = loaded
            .workbook_protection()
            .expect("workbook protection should be loaded");
        assert!(loaded_prot.lock_structure());
        assert!(loaded_prot.lock_windows());
        assert_eq!(loaded_prot.password_hash(), Some("ABCDEF"));
    }

    // ===== Feature 13: Calculation settings roundtrip =====

    #[test]
    fn calculation_settings_accessors_work() {
        let mut workbook = Workbook::new();
        assert!(workbook.calc_settings().is_none());

        let mut settings = CalculationSettings::new();
        assert!(!settings.has_metadata());

        settings
            .set_calc_mode("manual")
            .set_calc_id(191029)
            .set_full_calc_on_load(true);

        assert!(settings.has_metadata());
        assert_eq!(settings.calc_mode(), Some("manual"));
        assert_eq!(settings.calc_id(), Some(191029));
        assert_eq!(settings.full_calc_on_load(), Some(true));

        workbook.set_calc_settings(settings);
        assert!(workbook.calc_settings().is_some());

        workbook.clear_calc_settings();
        assert!(workbook.calc_settings().is_none());
    }

    #[test]
    fn calculation_settings_roundtrips_through_workbook_xml() {
        let mut workbook = Workbook::new();
        workbook.add_sheet("Data");

        let mut settings = CalculationSettings::new();
        settings
            .set_calc_mode("manual")
            .set_calc_id(191029)
            .set_full_calc_on_load(true);
        workbook.set_calc_settings(settings);

        let temp_file = NamedTempFile::new().expect("temp file should be created");
        workbook.save(temp_file.path()).expect("save workbook");

        let package = Package::open(temp_file.path()).expect("open saved package");
        let workbook_xml = String::from_utf8_lossy(
            package
                .get_part(WORKBOOK_PART_URI)
                .expect("workbook part should exist")
                .data
                .as_bytes(),
        )
        .into_owned();
        assert!(
            workbook_xml.contains("<calcPr"),
            "calcPr element should be present"
        );
        assert!(
            workbook_xml.contains("calcMode=\"manual\""),
            "calcMode attribute should be present"
        );
        assert!(
            workbook_xml.contains("calcId=\"191029\""),
            "calcId attribute should be present"
        );
        assert!(
            workbook_xml.contains("fullCalcOnLoad=\"1\""),
            "fullCalcOnLoad attribute should be present"
        );

        let loaded = Workbook::open(temp_file.path()).expect("open workbook");
        let loaded_settings = loaded
            .calc_settings()
            .expect("calculation settings should be loaded");
        assert_eq!(loaded_settings.calc_mode(), Some("manual"));
        assert_eq!(loaded_settings.calc_id(), Some(191029));
        assert_eq!(loaded_settings.full_calc_on_load(), Some(true));
    }

    // ===== Feature 6: Comments roundtrip =====

    #[test]
    fn comments_roundtrip_through_comments_xml_part() {
        use crate::worksheet::Comment;

        let mut workbook = Workbook::new();
        let sheet = workbook.add_sheet("Data");
        sheet
            .cell_mut("A1")
            .expect("cell should be valid")
            .set_value("commented");

        let comment1 =
            Comment::new("A1", "Author One", "First comment").expect("comment should be valid");
        let comment2 =
            Comment::new("B2", "Author Two", "Second comment").expect("comment should be valid");
        sheet.add_comment(comment1);
        sheet.add_comment(comment2);

        let temp_file = NamedTempFile::new().expect("temp file should be created");
        workbook.save(temp_file.path()).expect("save workbook");

        // Verify comments part exists
        let package = Package::open(temp_file.path()).expect("open saved package");
        let worksheet_part = package
            .get_part("/xl/worksheets/sheet1.xml")
            .expect("worksheet part should exist");
        let comments_relationships = worksheet_part
            .relationships
            .get_by_type(COMMENTS_RELATIONSHIP_TYPE);
        assert_eq!(
            comments_relationships.len(),
            1,
            "worksheet should have a comments relationship"
        );

        let comments_uri = PartUri::new("/xl/worksheets/sheet1.xml")
            .expect("worksheet uri should be valid")
            .resolve_relative(comments_relationships[0].target.as_str())
            .expect("comments target should resolve");
        let comments_part = package
            .get_part(comments_uri.as_str())
            .expect("comments part should exist");
        let comments_xml = String::from_utf8_lossy(comments_part.data.as_bytes());
        assert!(
            comments_xml.contains("<author>Author One</author>"),
            "first author should be in comments XML"
        );
        assert!(
            comments_xml.contains("<author>Author Two</author>"),
            "second author should be in comments XML"
        );
        assert!(
            comments_xml.contains("ref=\"A1\""),
            "first comment cell ref should be in comments XML"
        );
        assert!(
            comments_xml.contains("ref=\"B2\""),
            "second comment cell ref should be in comments XML"
        );
        assert!(
            comments_xml.contains("First comment"),
            "first comment text should be in comments XML"
        );
        assert!(
            comments_xml.contains("Second comment"),
            "second comment text should be in comments XML"
        );

        // Verify roundtrip
        let loaded = Workbook::open(temp_file.path()).expect("open workbook");
        let loaded_sheet = loaded.sheet("Data").expect("sheet should exist");
        assert_eq!(loaded_sheet.comments().len(), 2);
        assert_eq!(loaded_sheet.comments()[0].cell_ref(), "A1");
        assert_eq!(loaded_sheet.comments()[0].author(), "Author One");
        assert_eq!(loaded_sheet.comments()[0].text(), "First comment");
        assert_eq!(loaded_sheet.comments()[1].cell_ref(), "B2");
        assert_eq!(loaded_sheet.comments()[1].author(), "Author Two");
        assert_eq!(loaded_sheet.comments()[1].text(), "Second comment");
    }

    // ===== Feature 4 & 5: Insert/delete rows/columns roundtrip =====

    #[test]
    fn insert_rows_then_save_preserves_shifted_cells() {
        let mut workbook = Workbook::new();
        let sheet = workbook.add_sheet("Data");
        sheet
            .cell_mut("A1")
            .expect("cell should be valid")
            .set_value("row1");
        sheet
            .cell_mut("A2")
            .expect("cell should be valid")
            .set_value("row2");

        sheet.insert_rows(2, 1).expect("insert should succeed");

        let temp_file = NamedTempFile::new().expect("temp file should be created");
        workbook.save(temp_file.path()).expect("save workbook");

        let loaded = Workbook::open(temp_file.path()).expect("open workbook");
        let loaded_sheet = loaded.sheet("Data").expect("sheet should exist");
        assert_eq!(
            loaded_sheet.cell("A1").and_then(Cell::value),
            Some(&CellValue::String("row1".to_string()))
        );
        assert!(loaded_sheet.cell("A2").and_then(Cell::value).is_none());
        assert_eq!(
            loaded_sheet.cell("A3").and_then(Cell::value),
            Some(&CellValue::String("row2".to_string()))
        );
    }

    #[test]
    fn delete_columns_then_save_preserves_shifted_cells() {
        let mut workbook = Workbook::new();
        let sheet = workbook.add_sheet("Data");
        sheet
            .cell_mut("A1")
            .expect("cell should be valid")
            .set_value("colA");
        sheet
            .cell_mut("B1")
            .expect("cell should be valid")
            .set_value("colB");
        sheet
            .cell_mut("C1")
            .expect("cell should be valid")
            .set_value("colC");

        sheet.delete_columns(2, 1).expect("delete should succeed");

        let temp_file = NamedTempFile::new().expect("temp file should be created");
        workbook.save(temp_file.path()).expect("save workbook");

        let loaded = Workbook::open(temp_file.path()).expect("open workbook");
        let loaded_sheet = loaded.sheet("Data").expect("sheet should exist");
        assert_eq!(
            loaded_sheet.cell("A1").and_then(Cell::value),
            Some(&CellValue::String("colA".to_string()))
        );
        assert_eq!(
            loaded_sheet.cell("B1").and_then(Cell::value),
            Some(&CellValue::String("colC".to_string()))
        );
        assert!(loaded_sheet.cell("C1").and_then(Cell::value).is_none());
    }

    // ===== Font effects roundtrip =====

    #[test]
    fn font_effects_roundtrip_through_styles_xml() {
        let mut font = Font::new();
        font.set_name("Calibri")
            .set_size("11")
            .set_strikethrough(true)
            .set_shadow(true)
            .set_outline(true)
            .set_superscript(true);

        let mut style = Style::new();
        style.set_font(font);

        let styles = vec![style];
        let xml = serialize_styles_xml(&styles, &[], &[]).expect("serialize should succeed");
        let xml_str = String::from_utf8_lossy(&xml);

        assert!(xml_str.contains("<strike"), "should contain strike element");
        assert!(xml_str.contains("<shadow"), "should contain shadow element");
        assert!(
            xml_str.contains("<outline"),
            "should contain outline element"
        );
        assert!(
            xml_str.contains("vertAlign") && xml_str.contains("superscript"),
            "should contain vertAlign superscript"
        );

        let (parsed_styles, _, _, _) = parse_styles_xml(&xml).expect("parse should succeed");
        let parsed_font = parsed_styles.style(0).and_then(Style::font);
        assert!(parsed_font.is_some());
        let parsed_font = parsed_font.unwrap();
        assert_eq!(parsed_font.strikethrough(), Some(true));
        assert_eq!(parsed_font.shadow(), Some(true));
        assert_eq!(parsed_font.outline(), Some(true));
        assert_eq!(parsed_font.superscript(), Some(true));
    }

    #[test]
    fn double_strikethrough_and_subscript_roundtrip() {
        let mut font = Font::new();
        font.set_name("Arial")
            .set_double_strikethrough(true)
            .set_subscript(true);

        let mut style = Style::new();
        style.set_font(font);

        let styles = vec![style];
        let xml = serialize_styles_xml(&styles, &[], &[]).expect("serialize should succeed");
        let xml_str = String::from_utf8_lossy(&xml);

        assert!(
            xml_str.contains("<dStrike"),
            "should contain dStrike element"
        );
        assert!(
            xml_str.contains("vertAlign") && xml_str.contains("subscript"),
            "should contain vertAlign subscript"
        );

        let (parsed_styles, _, _, _) = parse_styles_xml(&xml).expect("parse should succeed");
        let parsed_font = parsed_styles.style(0).and_then(Style::font).unwrap();
        assert_eq!(parsed_font.double_strikethrough(), Some(true));
        assert_eq!(parsed_font.subscript(), Some(true));
    }

    // ===== Cell-level protection roundtrip =====

    #[test]
    fn cell_protection_roundtrip_through_styles_xml() {
        let mut protection = CellProtection::new();
        protection.set_locked(false).set_hidden(true);

        let mut style = Style::new();
        style.set_protection(protection);

        let styles = vec![style];
        let xml = serialize_styles_xml(&styles, &[], &[]).expect("serialize should succeed");
        let xml_str = String::from_utf8_lossy(&xml);

        assert!(
            xml_str.contains("applyProtection=\"1\""),
            "should have applyProtection"
        );
        assert!(
            xml_str.contains("<protection"),
            "should contain protection element"
        );
        assert!(xml_str.contains("locked=\"0\""), "should have locked=0");
        assert!(xml_str.contains("hidden=\"1\""), "should have hidden=1");

        let (parsed_styles, _, _, _) = parse_styles_xml(&xml).expect("parse should succeed");
        let parsed_protection = parsed_styles.style(0).and_then(Style::protection);
        assert!(parsed_protection.is_some());
        let parsed_protection = parsed_protection.unwrap();
        assert_eq!(parsed_protection.locked(), Some(false));
        assert_eq!(parsed_protection.hidden(), Some(true));
    }

    // ===== Tab color roundtrip =====

    #[test]
    fn tab_color_roundtrip() {
        let mut workbook = Workbook::new();
        let sheet = workbook.add_sheet("Colored");
        sheet.set_tab_color("FF0000");

        let temp_file = NamedTempFile::new().expect("temp file should be created");
        workbook.save(temp_file.path()).expect("save workbook");

        let loaded = Workbook::open(temp_file.path()).expect("open workbook");
        let loaded_sheet = loaded.sheet("Colored").expect("sheet should exist");
        assert_eq!(loaded_sheet.tab_color(), Some("FF0000"));
    }

    // ===== Default row height / column width roundtrip =====

    #[test]
    fn default_row_height_roundtrip() {
        let mut workbook = Workbook::new();
        let sheet = workbook.add_sheet("Formatted");
        sheet.set_default_row_height(20.0);
        sheet.set_default_column_width(12.5);
        sheet.set_custom_height(true);

        let temp_file = NamedTempFile::new().expect("temp file should be created");
        workbook.save(temp_file.path()).expect("save workbook");

        let loaded = Workbook::open(temp_file.path()).expect("open workbook");
        let loaded_sheet = loaded.sheet("Formatted").expect("sheet should exist");
        assert_eq!(loaded_sheet.default_row_height(), Some(20.0));
        assert_eq!(loaded_sheet.default_column_width(), Some(12.5));
        assert_eq!(loaded_sheet.custom_height(), Some(true));
    }

    // ===== Table enhancements roundtrip =====

    #[test]
    fn table_style_info_roundtrip() {
        let mut workbook = Workbook::new();
        let sheet = workbook.add_sheet("Tables");
        sheet
            .cell_mut("A1")
            .expect("valid cell")
            .set_value("Header1");
        sheet
            .cell_mut("B1")
            .expect("valid cell")
            .set_value("Header2");
        sheet.cell_mut("A2").expect("valid cell").set_value("Data1");
        sheet.cell_mut("B2").expect("valid cell").set_value("Data2");

        let mut table = WorksheetTable::new("MyTable", "A1:B2").expect("valid table");
        table.set_totals_row_shown(true);
        table.set_style_name("TableStyleMedium9");
        table.set_show_first_column(true);
        table.set_show_last_column(false);
        table.set_show_row_stripes(true);
        table.set_show_column_stripes(true);

        sheet.add_table(table);

        let temp_file = NamedTempFile::new().expect("temp file should be created");
        workbook.save(temp_file.path()).expect("save workbook");

        let loaded = Workbook::open(temp_file.path()).expect("open workbook");
        let loaded_sheet = loaded.sheet("Tables").expect("sheet should exist");
        assert_eq!(loaded_sheet.tables().len(), 1);

        let loaded_table = &loaded_sheet.tables()[0];
        assert_eq!(loaded_table.totals_row_shown(), Some(true));
        assert_eq!(loaded_table.style_name(), Some("TableStyleMedium9"));
        assert_eq!(loaded_table.show_first_column(), Some(true));
        assert_eq!(loaded_table.show_last_column(), Some(false));
        assert_eq!(loaded_table.show_row_stripes(), Some(true));
        assert_eq!(loaded_table.show_column_stripes(), Some(true));
    }

    // ===== Feature 12: Cached formula values =====

    #[test]
    fn cached_value_roundtrip_through_xml() {
        let mut workbook = Workbook::new();
        let sheet = workbook.add_sheet("Formulas");
        // Set up a cell with formula + cached value (simulating what Excel does)
        sheet
            .cell_mut("A1")
            .expect("valid cell")
            .set_formula("SUM(B1:B10)")
            .set_cached_value(CellValue::Number(42.0));

        let temp_file = NamedTempFile::new().expect("temp file should be created");
        workbook.save(temp_file.path()).expect("save workbook");

        // Check the raw XML contains both <f> and <v>
        let package = Package::open(temp_file.path()).expect("open package");
        let worksheet_xml = String::from_utf8_lossy(
            package
                .get_part("/xl/worksheets/sheet1.xml")
                .expect("sheet1 should exist")
                .data
                .as_bytes(),
        )
        .into_owned();
        assert!(
            worksheet_xml.contains("<f>SUM(B1:B10)</f>"),
            "formula element should be present: {worksheet_xml}"
        );
        assert!(
            worksheet_xml.contains("<v>42</v>"),
            "cached value element should be present: {worksheet_xml}"
        );

        // Reload and verify cached value is preserved
        let loaded = Workbook::open(temp_file.path()).expect("open workbook");
        let loaded_sheet = loaded.sheet("Formulas").expect("sheet should exist");
        let cell = loaded_sheet.cell("A1").expect("cell should exist");
        assert_eq!(cell.formula(), Some("SUM(B1:B10)"));
        assert_eq!(cell.cached_value(), Some(&CellValue::Number(42.0)));
        // Primary value should be None for formula cells
        assert!(cell.value().is_none());
    }

    #[test]
    fn cached_string_value_roundtrip() {
        let mut workbook = Workbook::new();
        let sheet = workbook.add_sheet("Formulas");
        sheet
            .cell_mut("A1")
            .expect("valid cell")
            .set_formula("CONCATENATE(B1,C1)")
            .set_cached_value(CellValue::String("HelloWorld".to_string()));

        let temp_file = NamedTempFile::new().expect("temp file should be created");
        workbook.save(temp_file.path()).expect("save workbook");

        // Check the raw XML - should have t="str" for a formula cell with string result
        let package = Package::open(temp_file.path()).expect("open package");
        let worksheet_xml = String::from_utf8_lossy(
            package
                .get_part("/xl/worksheets/sheet1.xml")
                .expect("sheet1 should exist")
                .data
                .as_bytes(),
        )
        .into_owned();
        assert!(
            worksheet_xml.contains("t=\"str\""),
            "cell type should be str for formula with string result: {worksheet_xml}"
        );
        assert!(
            worksheet_xml.contains("<v>HelloWorld</v>"),
            "cached string value should be present: {worksheet_xml}"
        );

        let loaded = Workbook::open(temp_file.path()).expect("open workbook");
        let cell = loaded
            .sheet("Formulas")
            .and_then(|s| s.cell("A1"))
            .expect("cell should exist");
        assert_eq!(cell.formula(), Some("CONCATENATE(B1,C1)"));
        assert_eq!(
            cell.cached_value(),
            Some(&CellValue::String("HelloWorld".to_string()))
        );
    }

    // ===== Feature 12: Shared formulas =====

    #[test]
    fn shared_formula_master_and_dependent_roundtrip() {
        let mut workbook = Workbook::new();
        let sheet = workbook.add_sheet("Shared");

        // Master cell: has formula text, shared index, and a ref range.
        // Use set_formula_ref (not set_array_range) so is_array_formula stays false.
        let master = sheet.cell_mut("A1").expect("valid cell");
        master.set_formula("SUM(B1:C1)").set_shared_formula_index(0);
        master.set_formula_ref("A1:A5");
        // Also set a cached value for the master
        sheet
            .cell_mut("A1")
            .expect("valid cell")
            .set_cached_value(CellValue::Number(10.0));

        // Dependent cell: only shared index, no formula text
        sheet
            .cell_mut("A2")
            .expect("valid cell")
            .set_shared_formula_index(0);
        sheet
            .cell_mut("A2")
            .expect("valid cell")
            .set_cached_value(CellValue::Number(20.0));

        let temp_file = NamedTempFile::new().expect("temp file should be created");
        workbook.save(temp_file.path()).expect("save workbook");

        // Check the raw XML
        let package = Package::open(temp_file.path()).expect("open package");
        let worksheet_xml = String::from_utf8_lossy(
            package
                .get_part("/xl/worksheets/sheet1.xml")
                .expect("sheet1 should exist")
                .data
                .as_bytes(),
        )
        .into_owned();

        // Master cell should have <f t="shared" si="0" ref="A1:A5">SUM(B1:C1)</f>
        assert!(
            worksheet_xml.contains("t=\"shared\""),
            "shared formula type should be present: {worksheet_xml}"
        );
        assert!(
            worksheet_xml.contains("si=\"0\""),
            "shared formula index should be present: {worksheet_xml}"
        );
        assert!(
            worksheet_xml.contains("SUM(B1:C1)"),
            "formula text should be present: {worksheet_xml}"
        );

        // Reload and verify
        let loaded = Workbook::open(temp_file.path()).expect("open workbook");
        let loaded_sheet = loaded.sheet("Shared").expect("sheet should exist");

        let master = loaded_sheet.cell("A1").expect("master cell should exist");
        assert_eq!(master.formula(), Some("SUM(B1:C1)"));
        assert_eq!(master.shared_formula_index(), Some(0));
        assert_eq!(master.cached_value(), Some(&CellValue::Number(10.0)));

        let dependent = loaded_sheet
            .cell("A2")
            .expect("dependent cell should exist");
        assert!(
            dependent.formula().is_none(),
            "dependent should have no formula text"
        );
        assert_eq!(dependent.shared_formula_index(), Some(0));
        assert_eq!(dependent.cached_value(), Some(&CellValue::Number(20.0)));
    }

    #[test]
    fn parse_worksheet_with_shared_formula_xml() {
        // Simulate XML that Excel would produce with shared formulas
        let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<sheetData>
<row r="1">
  <c r="A1"><f t="shared" ref="A1:A3" si="0">B1+C1</f><v>10</v></c>
</row>
<row r="2">
  <c r="A2"><f t="shared" si="0"/><v>20</v></c>
</row>
<row r="3">
  <c r="A3"><f t="shared" si="0"/><v>30</v></c>
</row>
</sheetData>
</worksheet>"#;

        let worksheet =
            parse_worksheet_xml("Sheet1", xml.as_bytes(), None).expect("parsing should succeed");

        let a1 = worksheet.cell("A1").expect("A1 should exist");
        assert_eq!(a1.formula(), Some("B1+C1"));
        assert_eq!(a1.shared_formula_index(), Some(0));
        assert_eq!(a1.cached_value(), Some(&CellValue::Number(10.0)));
        assert!(a1.value().is_none());

        let a2 = worksheet.cell("A2").expect("A2 should exist");
        assert!(a2.formula().is_none());
        assert_eq!(a2.shared_formula_index(), Some(0));
        assert_eq!(a2.cached_value(), Some(&CellValue::Number(20.0)));

        let a3 = worksheet.cell("A3").expect("A3 should exist");
        assert!(a3.formula().is_none());
        assert_eq!(a3.shared_formula_index(), Some(0));
        assert_eq!(a3.cached_value(), Some(&CellValue::Number(30.0)));
    }

    #[test]
    fn parse_worksheet_with_array_formula_attributes() {
        // Verify that array formula attributes are now parsed from <f> element
        let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<sheetData>
<row r="1">
  <c r="A1"><f t="array" ref="A1:A3">SUM(B1:B3*C1:C3)</f><v>100</v></c>
</row>
</sheetData>
</worksheet>"#;

        let worksheet =
            parse_worksheet_xml("Sheet1", xml.as_bytes(), None).expect("parsing should succeed");

        let a1 = worksheet.cell("A1").expect("A1 should exist");
        assert_eq!(a1.formula(), Some("SUM(B1:B3*C1:C3)"));
        assert!(a1.is_array_formula());
        assert_eq!(a1.array_range(), Some("A1:A3"));
        assert_eq!(a1.cached_value(), Some(&CellValue::Number(100.0)));
    }

    // ===== Sparkline parsing and serialization tests =====

    #[test]
    fn sparkline_parse_from_ext_lst() {
        let mut worksheet = Worksheet::new("Sheet1");

        // Build a fake extLst RawXmlNode with sparkline content.
        let sparkline_ext = RawXmlNode::Element {
            name: "ext".to_string(),
            attributes: vec![
                ("uri".to_string(), SPARKLINE_EXT_URI.to_string()),
                (
                    "xmlns:x14".to_string(),
                    "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main".to_string(),
                ),
            ],
            children: vec![RawXmlNode::Element {
                name: "x14:sparklineGroups".to_string(),
                attributes: vec![(
                    "xmlns:xm".to_string(),
                    "http://schemas.microsoft.com/office/excel/2006/main".to_string(),
                )],
                children: vec![RawXmlNode::Element {
                    name: "x14:sparklineGroup".to_string(),
                    attributes: vec![
                        ("type".to_string(), "column".to_string()),
                        ("displayEmptyCellsAs".to_string(), "zero".to_string()),
                        ("markers".to_string(), "1".to_string()),
                        ("high".to_string(), "1".to_string()),
                        ("low".to_string(), "1".to_string()),
                        ("first".to_string(), "1".to_string()),
                        ("last".to_string(), "1".to_string()),
                        ("negative".to_string(), "1".to_string()),
                        ("displayXAxis".to_string(), "1".to_string()),
                        ("minAxisType".to_string(), "group".to_string()),
                        ("maxAxisType".to_string(), "group".to_string()),
                    ],
                    children: vec![
                        RawXmlNode::Element {
                            name: "x14:colorSeries".to_string(),
                            attributes: vec![("rgb".to_string(), "FF376092".to_string())],
                            children: vec![],
                        },
                        RawXmlNode::Element {
                            name: "x14:colorNegative".to_string(),
                            attributes: vec![("rgb".to_string(), "FFD00000".to_string())],
                            children: vec![],
                        },
                        RawXmlNode::Element {
                            name: "x14:colorAxis".to_string(),
                            attributes: vec![("rgb".to_string(), "FF000000".to_string())],
                            children: vec![],
                        },
                        RawXmlNode::Element {
                            name: "x14:sparklines".to_string(),
                            attributes: vec![],
                            children: vec![
                                RawXmlNode::Element {
                                    name: "x14:sparkline".to_string(),
                                    attributes: vec![],
                                    children: vec![
                                        RawXmlNode::Element {
                                            name: "xm:f".to_string(),
                                            attributes: vec![],
                                            children: vec![RawXmlNode::Text(
                                                "Sheet1!B1:B10".to_string(),
                                            )],
                                        },
                                        RawXmlNode::Element {
                                            name: "xm:sqref".to_string(),
                                            attributes: vec![],
                                            children: vec![RawXmlNode::Text("A1".to_string())],
                                        },
                                    ],
                                },
                                RawXmlNode::Element {
                                    name: "x14:sparkline".to_string(),
                                    attributes: vec![],
                                    children: vec![
                                        RawXmlNode::Element {
                                            name: "xm:f".to_string(),
                                            attributes: vec![],
                                            children: vec![RawXmlNode::Text(
                                                "Sheet1!C1:C10".to_string(),
                                            )],
                                        },
                                        RawXmlNode::Element {
                                            name: "xm:sqref".to_string(),
                                            attributes: vec![],
                                            children: vec![RawXmlNode::Text("A2".to_string())],
                                        },
                                    ],
                                },
                            ],
                        },
                    ],
                }],
            }],
        };

        let ext_lst = RawXmlNode::Element {
            name: "extLst".to_string(),
            attributes: vec![],
            children: vec![sparkline_ext],
        };

        worksheet.push_unknown_child(ext_lst);
        extract_sparklines_from_ext_lst(&mut worksheet);

        // The sparkline ext was extracted and the extLst removed (it was the only ext).
        assert!(worksheet.unknown_children().is_empty());

        let groups = worksheet.sparkline_groups();
        assert_eq!(groups.len(), 1);

        let group = &groups[0];
        assert_eq!(group.sparkline_type(), SparklineType::Column);
        assert_eq!(group.display_empty_cells_as(), SparklineEmptyCells::Zero);
        assert!(group.markers());
        assert!(group.high_point());
        assert!(group.low_point());
        assert!(group.first_point());
        assert!(group.last_point());
        assert!(group.negative_points());
        assert!(group.display_x_axis());
        assert_eq!(group.min_axis_type(), SparklineAxisType::Group);
        assert_eq!(group.max_axis_type(), SparklineAxisType::Group);

        assert_eq!(group.colors().series.as_deref(), Some("FF376092"));
        assert_eq!(group.colors().negative.as_deref(), Some("FFD00000"));
        assert_eq!(group.colors().axis.as_deref(), Some("FF000000"));

        assert_eq!(group.sparklines().len(), 2);
        assert_eq!(group.sparklines()[0].location(), "A1");
        assert_eq!(group.sparklines()[0].data_range(), "Sheet1!B1:B10");
        assert_eq!(group.sparklines()[1].location(), "A2");
        assert_eq!(group.sparklines()[1].data_range(), "Sheet1!C1:C10");
    }

    #[test]
    fn sparkline_ext_lst_preserves_other_ext_elements() {
        let mut worksheet = Worksheet::new("Sheet1");

        let other_ext = RawXmlNode::Element {
            name: "ext".to_string(),
            attributes: vec![("uri".to_string(), "{OTHER-URI-NOT-SPARKLINES}".to_string())],
            children: vec![RawXmlNode::Text("some data".to_string())],
        };

        let sparkline_ext = RawXmlNode::Element {
            name: "ext".to_string(),
            attributes: vec![("uri".to_string(), SPARKLINE_EXT_URI.to_string())],
            children: vec![RawXmlNode::Element {
                name: "x14:sparklineGroups".to_string(),
                attributes: vec![],
                children: vec![RawXmlNode::Element {
                    name: "x14:sparklineGroup".to_string(),
                    attributes: vec![],
                    children: vec![RawXmlNode::Element {
                        name: "x14:sparklines".to_string(),
                        attributes: vec![],
                        children: vec![RawXmlNode::Element {
                            name: "x14:sparkline".to_string(),
                            attributes: vec![],
                            children: vec![
                                RawXmlNode::Element {
                                    name: "xm:f".to_string(),
                                    attributes: vec![],
                                    children: vec![RawXmlNode::Text("Sheet1!B1:B5".to_string())],
                                },
                                RawXmlNode::Element {
                                    name: "xm:sqref".to_string(),
                                    attributes: vec![],
                                    children: vec![RawXmlNode::Text("A1".to_string())],
                                },
                            ],
                        }],
                    }],
                }],
            }],
        };

        let ext_lst = RawXmlNode::Element {
            name: "extLst".to_string(),
            attributes: vec![],
            children: vec![other_ext, sparkline_ext],
        };

        worksheet.push_unknown_child(ext_lst);
        extract_sparklines_from_ext_lst(&mut worksheet);

        // The sparkline ext should be removed, but the other ext should remain.
        assert_eq!(worksheet.unknown_children().len(), 1);
        if let RawXmlNode::Element { name, children, .. } = &worksheet.unknown_children()[0] {
            assert_eq!(name, "extLst");
            assert_eq!(children.len(), 1);
            if let RawXmlNode::Element { attributes, .. } = &children[0] {
                assert_eq!(
                    attributes[0],
                    ("uri".to_string(), "{OTHER-URI-NOT-SPARKLINES}".to_string())
                );
            }
        }

        // Sparklines should still be parsed.
        assert_eq!(worksheet.sparkline_groups().len(), 1);
        assert_eq!(worksheet.sparkline_groups()[0].sparklines().len(), 1);
    }

    #[test]
    fn sparkline_serialize_roundtrip() {
        let mut group = SparklineGroup::new();
        group
            .set_sparkline_type(SparklineType::Column)
            .set_markers(true)
            .set_high_point(true)
            .add_sparkline(Sparkline::new("A1", "Sheet1!B1:B10"))
            .add_sparkline(Sparkline::new("A2", "Sheet1!C1:C10"));
        group.colors_mut().series = Some("FF376092".to_string());
        group.colors_mut().negative = Some("FFD00000".to_string());

        let mut worksheet = Worksheet::new("Sheet1");
        worksheet.add_sparkline_group(group);

        // Serialize the worksheet.
        let mut shared_strings = SharedStrings::new();
        let xml_bytes = serialize_worksheet_xml(&worksheet, &mut shared_strings, &[], None, &[])
            .expect("serialization should succeed");

        let xml_str = String::from_utf8(xml_bytes).expect("valid UTF-8");

        // Verify the XML contains sparkline elements.
        assert!(xml_str.contains("<extLst>"), "should contain extLst");
        assert!(
            xml_str.contains(SPARKLINE_EXT_URI),
            "should contain sparkline URI"
        );
        assert!(
            xml_str.contains("x14:sparklineGroups"),
            "should contain sparklineGroups"
        );
        assert!(
            xml_str.contains("x14:sparklineGroup"),
            "should contain sparklineGroup"
        );
        assert!(
            xml_str.contains("type=\"column\""),
            "should have column type"
        );
        assert!(xml_str.contains("markers=\"1\""), "should have markers");
        assert!(xml_str.contains("high=\"1\""), "should have high point");
        assert!(
            xml_str.contains("x14:sparkline"),
            "should contain sparkline elements"
        );
        assert!(
            xml_str.contains("Sheet1!B1:B10"),
            "should contain data range"
        );
        assert!(
            xml_str.contains("rgb=\"FF376092\""),
            "should contain color series"
        );
        assert!(
            xml_str.contains("rgb=\"FFD00000\""),
            "should contain color negative"
        );

        // Now parse the XML back through the parser and verify roundtrip.
        let mut parsed = parse_worksheet_xml("Sheet1", xml_str.as_bytes(), None)
            .expect("parsing should succeed");
        extract_sparklines_from_ext_lst(&mut parsed);

        let groups = parsed.sparkline_groups();
        assert_eq!(groups.len(), 1, "should have 1 sparkline group");

        let g = &groups[0];
        assert_eq!(g.sparkline_type(), SparklineType::Column);
        assert!(g.markers());
        assert!(g.high_point());
        assert_eq!(g.sparklines().len(), 2);
        assert_eq!(g.sparklines()[0].location(), "A1");
        assert_eq!(g.sparklines()[0].data_range(), "Sheet1!B1:B10");
        assert_eq!(g.sparklines()[1].location(), "A2");
        assert_eq!(g.sparklines()[1].data_range(), "Sheet1!C1:C10");
        assert_eq!(g.colors().series.as_deref(), Some("FF376092"));
        assert_eq!(g.colors().negative.as_deref(), Some("FFD00000"));
    }

    /// Regression: modifying a cell on a worksheet whose rows have attributes like
    /// `customHeight`, `spans`, or `x14ac:dyDescent` should NOT produce duplicate
    /// XML attributes, and namespace declarations must survive the dirty save.
    #[test]
    fn dirty_save_no_duplicate_row_attrs_and_preserves_namespaces() {
        // Build a worksheet XML with `customHeight`, `spans`, and `x14ac:dyDescent`
        // on a row element, plus an extra namespace on the <worksheet>.
        let ws_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
           xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
           xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac"
           xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">
  <sheetData>
    <row r="1" ht="15.75" customHeight="1" spans="1:5" x14ac:dyDescent="0.25">
      <c r="A1" t="s"><v>0</v></c>
      <c r="B1" t="s"><v>1</v></c>
    </row>
    <row r="2" ht="20" customHeight="1" spans="1:3" thickBot="1" x14ac:dyDescent="0.3">
      <c r="A2"><v>42</v></c>
    </row>
  </sheetData>
</worksheet>"#;

        let sst_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="2" uniqueCount="2">
  <si><t>Hello</t></si>
  <si><t>World</t></si>
</sst>"#;

        let wb_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Sheet1" sheetId="1" r:id="rId1"/>
  </sheets>
</workbook>"#;

        let mut package = Package::new();

        let wb_uri = PartUri::new("/xl/workbook.xml").unwrap();
        let mut wb_part = Part::new_xml(wb_uri, wb_xml.to_vec());
        wb_part.content_type = Some(ContentTypeValue::WORKBOOK.to_string());
        wb_part.relationships.add_new(
            RelationshipType::WORKSHEET.to_string(),
            "worksheets/sheet1.xml".to_string(),
            TargetMode::Internal,
        );
        wb_part.relationships.add_new(
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings"
                .to_string(),
            "sharedStrings.xml".to_string(),
            TargetMode::Internal,
        );
        package.set_part(wb_part);

        let ws_uri = PartUri::new("/xl/worksheets/sheet1.xml").unwrap();
        let mut ws_part = Part::new_xml(ws_uri, ws_xml.to_vec());
        ws_part.content_type = Some(ContentTypeValue::WORKSHEET.to_string());
        package.set_part(ws_part);

        let sst_uri = PartUri::new("/xl/sharedStrings.xml").unwrap();
        let mut sst_part = Part::new_xml(sst_uri, sst_xml.to_vec());
        sst_part.content_type = Some(
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"
                .to_string(),
        );
        package.set_part(sst_part);

        // Add required OPC relationship
        package.relationships_mut().add_new(
            RelationshipType::OFFICE_DOCUMENT.to_string(),
            "xl/workbook.xml".to_string(),
            TargetMode::Internal,
        );

        let tmpdir = tempfile::tempdir().unwrap();
        let pkg_path = tmpdir.path().join("test.xlsx");
        package.save(&pkg_path).unwrap();

        // Open, modify a cell (marks sheet dirty), and save.
        let mut wb = Workbook::open(&pkg_path).unwrap();
        wb.worksheets_mut()[0]
            .cell_mut("A1")
            .unwrap()
            .set_value("Modified");
        let out_path = tmpdir.path().join("out.xlsx");
        wb.save(&out_path).unwrap();

        // Extract the worksheet XML and verify.
        let out_package = Package::open(&out_path).unwrap();
        let ws_part = out_package
            .get_part("/xl/worksheets/sheet1.xml")
            .expect("worksheet part missing");
        let ws_xml_out = String::from_utf8_lossy(ws_part.data.as_bytes());

        // No duplicate customHeight attributes.
        for line in ws_xml_out.lines() {
            let count = line.matches("customHeight").count();
            assert!(count <= 1, "duplicate customHeight on line: {line}");
        }

        // The x14ac namespace should be declared on <worksheet>.
        assert!(
            ws_xml_out.contains("xmlns:x14ac"),
            "x14ac namespace declaration missing from worksheet XML:\n{ws_xml_out}"
        );

        // The mc namespace should also be preserved.
        assert!(
            ws_xml_out.contains("xmlns:mc"),
            "mc namespace declaration missing from worksheet XML:\n{ws_xml_out}"
        );

        // x14ac:dyDescent should be present on rows (roundtripped as unknown attr).
        assert!(
            ws_xml_out.contains("x14ac:dyDescent"),
            "x14ac:dyDescent attr missing from worksheet XML:\n{ws_xml_out}"
        );

        // All original cell data should survive (not just the modified cell).
        let wb2 = Workbook::open(&out_path).unwrap();
        let ws = &wb2.worksheets()[0];
        assert!(
            matches!(ws.cell("A1").unwrap().value(), Some(CellValue::String(s)) if s == "Modified"),
            "A1 should be 'Modified', got: {:?}",
            ws.cell("A1").unwrap().value()
        );
        assert!(
            matches!(ws.cell("B1").unwrap().value(), Some(CellValue::String(s)) if s == "World"),
            "B1 should be 'World', got: {:?}",
            ws.cell("B1").unwrap().value()
        );
        assert!(
            matches!(ws.cell("A2").unwrap().value(), Some(CellValue::Number(n)) if (*n - 42.0).abs() < f64::EPSILON),
            "A2 should be 42, got: {:?}",
            ws.cell("A2").unwrap().value()
        );
    }
}
