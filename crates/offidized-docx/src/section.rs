use crate::paragraph::Paragraph;
use crate::table::Table;

/// Page orientation for a section.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum PageOrientation {
    Portrait,
    Landscape,
}

impl PageOrientation {
    pub(crate) fn from_xml_value(value: &str) -> Option<Self> {
        match value {
            "portrait" => Some(Self::Portrait),
            "landscape" => Some(Self::Landscape),
            _ => None,
        }
    }

    pub(crate) fn to_xml_value(self) -> &'static str {
        match self {
            Self::Portrait => "portrait",
            Self::Landscape => "landscape",
        }
    }
}

/// Section break type for non-final sections.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum SectionBreakType {
    /// Next page section break (`w:type w:val="nextPage"`).
    NextPage,
    /// Continuous section break (`w:type w:val="continuous"`).
    Continuous,
    /// Even page section break (`w:type w:val="evenPage"`).
    EvenPage,
    /// Odd page section break (`w:type w:val="oddPage"`).
    OddPage,
}

impl SectionBreakType {
    pub(crate) fn from_xml_value(value: &str) -> Option<Self> {
        match value {
            "nextPage" => Some(Self::NextPage),
            "continuous" => Some(Self::Continuous),
            "evenPage" => Some(Self::EvenPage),
            "oddPage" => Some(Self::OddPage),
            _ => None,
        }
    }

    pub(crate) fn to_xml_value(self) -> &'static str {
        match self {
            Self::NextPage => "nextPage",
            Self::Continuous => "continuous",
            Self::EvenPage => "evenPage",
            Self::OddPage => "oddPage",
        }
    }
}

/// Vertical alignment of text within a section.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum SectionVerticalAlignment {
    /// Text aligned to the top of the page (`w:vAlign w:val="top"`).
    Top,
    /// Text centered vertically on the page (`w:vAlign w:val="center"`).
    Center,
    /// Text justified vertically across the page (`w:vAlign w:val="both"`).
    Justify,
    /// Text aligned to the bottom of the page (`w:vAlign w:val="bottom"`).
    Bottom,
}

impl SectionVerticalAlignment {
    pub(crate) fn from_xml_value(value: &str) -> Option<Self> {
        match value {
            "top" => Some(Self::Top),
            "center" => Some(Self::Center),
            "both" => Some(Self::Justify),
            "bottom" => Some(Self::Bottom),
            _ => None,
        }
    }

    pub(crate) fn to_xml_value(self) -> &'static str {
        match self {
            Self::Top => "top",
            Self::Center => "center",
            Self::Justify => "both",
            Self::Bottom => "bottom",
        }
    }
}

/// Controls when line numbering restarts.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum LineNumberRestart {
    /// Restart line numbering on each new page.
    NewPage,
    /// Restart line numbering on each new section.
    NewSection,
    /// Continuous line numbering across sections.
    Continuous,
}

impl LineNumberRestart {
    pub(crate) fn from_xml_value(value: &str) -> Option<Self> {
        match value {
            "newPage" => Some(Self::NewPage),
            "newSection" => Some(Self::NewSection),
            "continuous" => Some(Self::Continuous),
            _ => None,
        }
    }

    pub(crate) fn to_xml_value(self) -> &'static str {
        match self {
            Self::NewPage => "newPage",
            Self::NewSection => "newSection",
            Self::Continuous => "continuous",
        }
    }
}

/// Page margins for a section, in twips.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub struct PageMargins {
    top_twips: Option<u32>,
    right_twips: Option<u32>,
    bottom_twips: Option<u32>,
    left_twips: Option<u32>,
    header_twips: Option<u32>,
    footer_twips: Option<u32>,
    gutter_twips: Option<u32>,
}

impl PageMargins {
    pub fn new() -> Self {
        Self::default()
    }

    pub fn top_twips(&self) -> Option<u32> {
        self.top_twips
    }

    pub fn set_top_twips(&mut self, value: u32) {
        self.top_twips = Some(value);
    }

    pub fn right_twips(&self) -> Option<u32> {
        self.right_twips
    }

    pub fn set_right_twips(&mut self, value: u32) {
        self.right_twips = Some(value);
    }

    pub fn bottom_twips(&self) -> Option<u32> {
        self.bottom_twips
    }

    pub fn set_bottom_twips(&mut self, value: u32) {
        self.bottom_twips = Some(value);
    }

    pub fn left_twips(&self) -> Option<u32> {
        self.left_twips
    }

    pub fn set_left_twips(&mut self, value: u32) {
        self.left_twips = Some(value);
    }

    pub fn header_twips(&self) -> Option<u32> {
        self.header_twips
    }

    pub fn set_header_twips(&mut self, value: u32) {
        self.header_twips = Some(value);
    }

    pub fn footer_twips(&self) -> Option<u32> {
        self.footer_twips
    }

    pub fn set_footer_twips(&mut self, value: u32) {
        self.footer_twips = Some(value);
    }

    pub fn gutter_twips(&self) -> Option<u32> {
        self.gutter_twips
    }

    pub fn set_gutter_twips(&mut self, value: u32) {
        self.gutter_twips = Some(value);
    }

    pub fn clear(&mut self) {
        *self = Self::default();
    }

    pub(crate) fn is_empty(&self) -> bool {
        self.top_twips.is_none()
            && self.right_twips.is_none()
            && self.bottom_twips.is_none()
            && self.left_twips.is_none()
            && self.header_twips.is_none()
            && self.footer_twips.is_none()
            && self.gutter_twips.is_none()
    }
}

/// Header or footer content scoped to a section.
#[derive(Debug, Clone, PartialEq, Eq, Default)]
pub struct HeaderFooter {
    paragraphs: Vec<Paragraph>,
    tables: Vec<Table>,
}

impl HeaderFooter {
    pub fn new() -> Self {
        Self::default()
    }

    pub fn from_text(text: impl Into<String>) -> Self {
        let mut header_footer = Self::new();
        header_footer.add_paragraph(text);
        header_footer
    }

    pub fn add_paragraph(&mut self, text: impl Into<String>) -> &mut Paragraph {
        self.paragraphs.push(Paragraph::from_text(text));
        let index = self.paragraphs.len().saturating_sub(1);
        &mut self.paragraphs[index]
    }

    pub fn paragraphs(&self) -> &[Paragraph] {
        &self.paragraphs
    }

    pub fn paragraphs_mut(&mut self) -> &mut [Paragraph] {
        &mut self.paragraphs
    }

    pub fn set_paragraphs(&mut self, paragraphs: Vec<Paragraph>) {
        self.paragraphs = paragraphs;
    }

    /// Add a table to this header/footer.
    pub fn add_table(&mut self, rows: usize, columns: usize) -> &mut Table {
        self.tables.push(Table::new(rows, columns));
        let index = self.tables.len().saturating_sub(1);
        &mut self.tables[index]
    }

    /// Tables in this header/footer.
    pub fn tables(&self) -> &[Table] {
        &self.tables
    }

    /// Mutable tables.
    pub fn tables_mut(&mut self) -> &mut [Table] {
        &mut self.tables
    }

    /// Set tables.
    pub fn set_tables(&mut self, tables: Vec<Table>) {
        self.tables = tables;
    }

    pub fn clear(&mut self) {
        self.paragraphs.clear();
        self.tables.clear();
    }
}

/// Section settings for document-level layout.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Section {
    title: Option<String>,
    page_width_twips: Option<u32>,
    page_height_twips: Option<u32>,
    page_orientation: Option<PageOrientation>,
    page_margins: PageMargins,
    header: Option<HeaderFooter>,
    footer: Option<HeaderFooter>,
    first_page_header: Option<HeaderFooter>,
    first_page_footer: Option<HeaderFooter>,
    even_page_header: Option<HeaderFooter>,
    even_page_footer: Option<HeaderFooter>,
    title_page: bool,
    break_type: Option<SectionBreakType>,
    page_number_start: Option<u32>,
    page_number_format: Option<String>,
    // Multi-column layout
    column_count: Option<u16>,
    column_space_twips: Option<u32>,
    column_separator: bool,
    // Vertical alignment
    vertical_alignment: Option<SectionVerticalAlignment>,
    // Line numbering
    line_numbering_start: Option<u32>,
    line_numbering_count_by: Option<u32>,
    line_numbering_restart: Option<LineNumberRestart>,
    line_numbering_distance_twips: Option<u32>,
    // Link to previous (headers/footers) — defaults to true per OOXML spec
    header_link_to_previous: bool,
    footer_link_to_previous: bool,
    first_header_link_to_previous: bool,
    first_footer_link_to_previous: bool,
    even_header_link_to_previous: bool,
    even_footer_link_to_previous: bool,
}

impl Default for Section {
    fn default() -> Self {
        Self {
            title: None,
            page_width_twips: None,
            page_height_twips: None,
            page_orientation: None,
            page_margins: PageMargins::default(),
            header: None,
            footer: None,
            first_page_header: None,
            first_page_footer: None,
            even_page_header: None,
            even_page_footer: None,
            title_page: false,
            break_type: None,
            page_number_start: None,
            page_number_format: None,
            column_count: None,
            column_space_twips: None,
            column_separator: false,
            vertical_alignment: None,
            line_numbering_start: None,
            line_numbering_count_by: None,
            line_numbering_restart: None,
            line_numbering_distance_twips: None,
            header_link_to_previous: true,
            footer_link_to_previous: true,
            first_header_link_to_previous: true,
            first_footer_link_to_previous: true,
            even_header_link_to_previous: true,
            even_footer_link_to_previous: true,
        }
    }
}

impl Section {
    pub fn new() -> Self {
        Self::default()
    }

    pub fn with_title(title: impl Into<String>) -> Self {
        Self {
            title: Some(title.into()),
            ..Self::default()
        }
    }

    pub fn title(&self) -> Option<&str> {
        self.title.as_deref()
    }

    /// Page width in twips.
    pub fn page_width_twips(&self) -> Option<u32> {
        self.page_width_twips
    }

    /// Page height in twips.
    pub fn page_height_twips(&self) -> Option<u32> {
        self.page_height_twips
    }

    /// Set page size in twips.
    pub fn set_page_size_twips(&mut self, width_twips: u32, height_twips: u32) {
        self.page_width_twips = Some(width_twips);
        self.page_height_twips = Some(height_twips);
    }

    /// Clear explicit page size.
    pub fn clear_page_size_twips(&mut self) {
        self.page_width_twips = None;
        self.page_height_twips = None;
    }

    /// Page orientation.
    pub fn page_orientation(&self) -> Option<PageOrientation> {
        self.page_orientation
    }

    /// Set page orientation.
    pub fn set_page_orientation(&mut self, orientation: PageOrientation) {
        self.page_orientation = Some(orientation);
    }

    /// Clear explicit page orientation.
    pub fn clear_page_orientation(&mut self) {
        self.page_orientation = None;
    }

    /// Page margins.
    pub fn page_margins(&self) -> &PageMargins {
        &self.page_margins
    }

    /// Mutable page margins.
    pub fn page_margins_mut(&mut self) -> &mut PageMargins {
        &mut self.page_margins
    }

    /// Replace page margins.
    pub fn set_page_margins(&mut self, margins: PageMargins) {
        self.page_margins = margins;
    }

    /// Clear page margins.
    pub fn clear_page_margins(&mut self) {
        self.page_margins.clear();
    }

    /// Default header content for this section.
    pub fn header(&self) -> Option<&HeaderFooter> {
        self.header.as_ref()
    }

    /// Mutable default header content for this section.
    pub fn header_mut(&mut self) -> Option<&mut HeaderFooter> {
        self.header.as_mut()
    }

    /// Create header content when missing and return it.
    pub fn ensure_header(&mut self) -> &mut HeaderFooter {
        self.header.get_or_insert_with(HeaderFooter::new)
    }

    /// Replace default header content.
    pub fn set_header(&mut self, header: HeaderFooter) {
        self.header = Some(header);
    }

    /// Remove default header content.
    pub fn clear_header(&mut self) {
        self.header = None;
    }

    /// Default footer content for this section.
    pub fn footer(&self) -> Option<&HeaderFooter> {
        self.footer.as_ref()
    }

    /// Mutable default footer content for this section.
    pub fn footer_mut(&mut self) -> Option<&mut HeaderFooter> {
        self.footer.as_mut()
    }

    /// Create footer content when missing and return it.
    pub fn ensure_footer(&mut self) -> &mut HeaderFooter {
        self.footer.get_or_insert_with(HeaderFooter::new)
    }

    /// Replace default footer content.
    pub fn set_footer(&mut self, footer: HeaderFooter) {
        self.footer = Some(footer);
    }

    /// Remove default footer content.
    pub fn clear_footer(&mut self) {
        self.footer = None;
    }

    /// First page header content for this section (requires `title_page` enabled).
    pub fn first_page_header(&self) -> Option<&HeaderFooter> {
        self.first_page_header.as_ref()
    }

    /// Mutable first page header content.
    pub fn first_page_header_mut(&mut self) -> Option<&mut HeaderFooter> {
        self.first_page_header.as_mut()
    }

    /// Create first page header content when missing and return it.
    /// Automatically enables `title_page`.
    pub fn ensure_first_page_header(&mut self) -> &mut HeaderFooter {
        self.title_page = true;
        self.first_page_header.get_or_insert_with(HeaderFooter::new)
    }

    /// Replace first page header content. Automatically enables `title_page`.
    pub fn set_first_page_header(&mut self, header: HeaderFooter) {
        self.title_page = true;
        self.first_page_header = Some(header);
    }

    /// Remove first page header content.
    pub fn clear_first_page_header(&mut self) {
        self.first_page_header = None;
    }

    /// First page footer content for this section (requires `title_page` enabled).
    pub fn first_page_footer(&self) -> Option<&HeaderFooter> {
        self.first_page_footer.as_ref()
    }

    /// Mutable first page footer content.
    pub fn first_page_footer_mut(&mut self) -> Option<&mut HeaderFooter> {
        self.first_page_footer.as_mut()
    }

    /// Create first page footer content when missing and return it.
    /// Automatically enables `title_page`.
    pub fn ensure_first_page_footer(&mut self) -> &mut HeaderFooter {
        self.title_page = true;
        self.first_page_footer.get_or_insert_with(HeaderFooter::new)
    }

    /// Replace first page footer content. Automatically enables `title_page`.
    pub fn set_first_page_footer(&mut self, footer: HeaderFooter) {
        self.title_page = true;
        self.first_page_footer = Some(footer);
    }

    /// Remove first page footer content.
    pub fn clear_first_page_footer(&mut self) {
        self.first_page_footer = None;
    }

    /// Even page header content for this section.
    pub fn even_page_header(&self) -> Option<&HeaderFooter> {
        self.even_page_header.as_ref()
    }

    /// Mutable even page header content.
    pub fn even_page_header_mut(&mut self) -> Option<&mut HeaderFooter> {
        self.even_page_header.as_mut()
    }

    /// Create even page header content when missing and return it.
    pub fn ensure_even_page_header(&mut self) -> &mut HeaderFooter {
        self.even_page_header.get_or_insert_with(HeaderFooter::new)
    }

    /// Replace even page header content.
    pub fn set_even_page_header(&mut self, header: HeaderFooter) {
        self.even_page_header = Some(header);
    }

    /// Remove even page header content.
    pub fn clear_even_page_header(&mut self) {
        self.even_page_header = None;
    }

    /// Even page footer content for this section.
    pub fn even_page_footer(&self) -> Option<&HeaderFooter> {
        self.even_page_footer.as_ref()
    }

    /// Mutable even page footer content.
    pub fn even_page_footer_mut(&mut self) -> Option<&mut HeaderFooter> {
        self.even_page_footer.as_mut()
    }

    /// Create even page footer content when missing and return it.
    pub fn ensure_even_page_footer(&mut self) -> &mut HeaderFooter {
        self.even_page_footer.get_or_insert_with(HeaderFooter::new)
    }

    /// Replace even page footer content.
    pub fn set_even_page_footer(&mut self, footer: HeaderFooter) {
        self.even_page_footer = Some(footer);
    }

    /// Remove even page footer content.
    pub fn clear_even_page_footer(&mut self) {
        self.even_page_footer = None;
    }

    /// Whether this section has a different first page header/footer (`w:titlePg`).
    pub fn title_page(&self) -> bool {
        self.title_page
    }

    /// Enable or disable different first page header/footer (`w:titlePg`).
    pub fn set_title_page(&mut self, title_page: bool) {
        self.title_page = title_page;
    }

    /// Section break type for non-final sections.
    pub fn break_type(&self) -> Option<SectionBreakType> {
        self.break_type
    }

    /// Set section break type.
    pub fn set_break_type(&mut self, break_type: SectionBreakType) {
        self.break_type = Some(break_type);
    }

    /// Clear section break type.
    pub fn clear_break_type(&mut self) {
        self.break_type = None;
    }

    /// Starting page number for this section (`w:pgNumType w:start`).
    pub fn page_number_start(&self) -> Option<u32> {
        self.page_number_start
    }

    /// Set starting page number for this section.
    pub fn set_page_number_start(&mut self, start: u32) {
        self.page_number_start = Some(start);
    }

    /// Clear starting page number.
    pub fn clear_page_number_start(&mut self) {
        self.page_number_start = None;
    }

    /// Page number format for this section (`w:pgNumType w:fmt`), e.g. `"decimal"`, `"lowerRoman"`.
    pub fn page_number_format(&self) -> Option<&str> {
        self.page_number_format.as_deref()
    }

    /// Set page number format for this section.
    pub fn set_page_number_format(&mut self, format: impl Into<String>) {
        let format = format.into();
        self.page_number_format = if format.trim().is_empty() {
            None
        } else {
            Some(format)
        };
    }

    /// Clear page number format.
    pub fn clear_page_number_format(&mut self) {
        self.page_number_format = None;
    }

    // ---- Multi-column layout ----

    /// Number of columns in this section (default 1).
    pub fn column_count(&self) -> Option<u16> {
        self.column_count
    }

    /// Set the number of columns for this section.
    pub fn set_column_count(&mut self, count: u16) {
        self.column_count = Some(count);
    }

    /// Clear the explicit column count.
    pub fn clear_column_count(&mut self) {
        self.column_count = None;
    }

    /// Space between columns in twips.
    pub fn column_space_twips(&self) -> Option<u32> {
        self.column_space_twips
    }

    /// Set space between columns in twips.
    pub fn set_column_space_twips(&mut self, space: u32) {
        self.column_space_twips = Some(space);
    }

    /// Clear the explicit column spacing.
    pub fn clear_column_space_twips(&mut self) {
        self.column_space_twips = None;
    }

    /// Whether a separator line is drawn between columns (`w:cols w:sep`).
    pub fn column_separator(&self) -> bool {
        self.column_separator
    }

    /// Enable or disable separator line between columns.
    pub fn set_column_separator(&mut self, separator: bool) {
        self.column_separator = separator;
    }

    // ---- Vertical alignment ----

    /// Vertical alignment of text within this section (`w:vAlign`).
    pub fn vertical_alignment(&self) -> Option<SectionVerticalAlignment> {
        self.vertical_alignment
    }

    /// Set the vertical alignment for this section.
    pub fn set_vertical_alignment(&mut self, alignment: SectionVerticalAlignment) {
        self.vertical_alignment = Some(alignment);
    }

    /// Clear the explicit vertical alignment.
    pub fn clear_vertical_alignment(&mut self) {
        self.vertical_alignment = None;
    }

    // ---- Line numbering ----

    /// Starting line number for line numbering (`w:lnNumType w:start`).
    pub fn line_numbering_start(&self) -> Option<u32> {
        self.line_numbering_start
    }

    /// Set the starting line number.
    pub fn set_line_numbering_start(&mut self, start: u32) {
        self.line_numbering_start = Some(start);
    }

    /// Clear the starting line number.
    pub fn clear_line_numbering_start(&mut self) {
        self.line_numbering_start = None;
    }

    /// Line numbering interval — display every Nth line number (`w:lnNumType w:countBy`).
    pub fn line_numbering_count_by(&self) -> Option<u32> {
        self.line_numbering_count_by
    }

    /// Set the line numbering interval.
    pub fn set_line_numbering_count_by(&mut self, count_by: u32) {
        self.line_numbering_count_by = Some(count_by);
    }

    /// Clear the line numbering interval.
    pub fn clear_line_numbering_count_by(&mut self) {
        self.line_numbering_count_by = None;
    }

    /// When line numbering restarts (`w:lnNumType w:restart`).
    pub fn line_numbering_restart(&self) -> Option<LineNumberRestart> {
        self.line_numbering_restart
    }

    /// Set when line numbering restarts.
    pub fn set_line_numbering_restart(&mut self, restart: LineNumberRestart) {
        self.line_numbering_restart = Some(restart);
    }

    /// Clear the line numbering restart setting.
    pub fn clear_line_numbering_restart(&mut self) {
        self.line_numbering_restart = None;
    }

    /// Distance of line numbers from the text edge in twips (`w:lnNumType w:distance`).
    pub fn line_numbering_distance_twips(&self) -> Option<u32> {
        self.line_numbering_distance_twips
    }

    /// Set the distance of line numbers from the text edge in twips.
    pub fn set_line_numbering_distance_twips(&mut self, distance: u32) {
        self.line_numbering_distance_twips = Some(distance);
    }

    /// Clear the line numbering distance.
    pub fn clear_line_numbering_distance_twips(&mut self) {
        self.line_numbering_distance_twips = None;
    }

    // ---- Link to previous (headers/footers) ----

    /// Whether the default header links to the previous section (defaults to `true` per spec).
    pub fn header_link_to_previous(&self) -> bool {
        self.header_link_to_previous
    }

    /// Set whether the default header links to the previous section.
    pub fn set_header_link_to_previous(&mut self, link: bool) {
        self.header_link_to_previous = link;
    }

    /// Whether the default footer links to the previous section (defaults to `true` per spec).
    pub fn footer_link_to_previous(&self) -> bool {
        self.footer_link_to_previous
    }

    /// Set whether the default footer links to the previous section.
    pub fn set_footer_link_to_previous(&mut self, link: bool) {
        self.footer_link_to_previous = link;
    }

    /// Whether the first page header links to the previous section (defaults to `true` per spec).
    pub fn first_header_link_to_previous(&self) -> bool {
        self.first_header_link_to_previous
    }

    /// Set whether the first page header links to the previous section.
    pub fn set_first_header_link_to_previous(&mut self, link: bool) {
        self.first_header_link_to_previous = link;
    }

    /// Whether the first page footer links to the previous section (defaults to `true` per spec).
    pub fn first_footer_link_to_previous(&self) -> bool {
        self.first_footer_link_to_previous
    }

    /// Set whether the first page footer links to the previous section.
    pub fn set_first_footer_link_to_previous(&mut self, link: bool) {
        self.first_footer_link_to_previous = link;
    }

    /// Whether the even page header links to the previous section (defaults to `true` per spec).
    pub fn even_header_link_to_previous(&self) -> bool {
        self.even_header_link_to_previous
    }

    /// Set whether the even page header links to the previous section.
    pub fn set_even_header_link_to_previous(&mut self, link: bool) {
        self.even_header_link_to_previous = link;
    }

    /// Whether the even page footer links to the previous section (defaults to `true` per spec).
    pub fn even_footer_link_to_previous(&self) -> bool {
        self.even_footer_link_to_previous
    }

    /// Set whether the even page footer links to the previous section.
    pub fn set_even_footer_link_to_previous(&mut self, link: bool) {
        self.even_footer_link_to_previous = link;
    }

    pub(crate) fn set_header_option(&mut self, header: Option<HeaderFooter>) {
        self.header = header;
    }

    pub(crate) fn set_footer_option(&mut self, footer: Option<HeaderFooter>) {
        self.footer = footer;
    }

    pub(crate) fn set_first_page_header_option(&mut self, header: Option<HeaderFooter>) {
        self.first_page_header = header;
    }

    pub(crate) fn set_first_page_footer_option(&mut self, footer: Option<HeaderFooter>) {
        self.first_page_footer = footer;
    }

    pub(crate) fn set_even_page_header_option(&mut self, header: Option<HeaderFooter>) {
        self.even_page_header = header;
    }

    pub(crate) fn set_even_page_footer_option(&mut self, footer: Option<HeaderFooter>) {
        self.even_page_footer = footer;
    }

    pub(crate) fn set_page_number_start_option(&mut self, value: Option<u32>) {
        self.page_number_start = value;
    }

    pub(crate) fn set_page_number_format_option(&mut self, value: Option<String>) {
        self.page_number_format = value;
    }

    pub(crate) fn set_page_width_twips(&mut self, value: Option<u32>) {
        self.page_width_twips = value;
    }

    pub(crate) fn set_page_height_twips(&mut self, value: Option<u32>) {
        self.page_height_twips = value;
    }

    pub(crate) fn set_page_orientation_option(&mut self, value: Option<PageOrientation>) {
        self.page_orientation = value;
    }

    pub(crate) fn set_break_type_option(&mut self, value: Option<SectionBreakType>) {
        self.break_type = value;
    }

    pub(crate) fn has_properties(&self) -> bool {
        self.page_width_twips.is_some()
            || self.page_height_twips.is_some()
            || self.page_orientation.is_some()
            || !self.page_margins.is_empty()
            || self.header.is_some()
            || self.footer.is_some()
            || self.first_page_header.is_some()
            || self.first_page_footer.is_some()
            || self.even_page_header.is_some()
            || self.even_page_footer.is_some()
            || self.title_page
            || self.break_type.is_some()
            || self.page_number_start.is_some()
            || self.page_number_format.is_some()
            || self.column_count.is_some()
            || self.column_space_twips.is_some()
            || self.column_separator
            || self.vertical_alignment.is_some()
            || self.line_numbering_start.is_some()
            || self.line_numbering_count_by.is_some()
            || self.line_numbering_restart.is_some()
            || self.line_numbering_distance_twips.is_some()
            || !self.header_link_to_previous
            || !self.footer_link_to_previous
            || !self.first_header_link_to_previous
            || !self.first_footer_link_to_previous
            || !self.even_header_link_to_previous
            || !self.even_footer_link_to_previous
    }
}

#[cfg(test)]
mod tests {
    use super::{
        HeaderFooter, LineNumberRestart, PageMargins, PageOrientation, Section, SectionBreakType,
        SectionVerticalAlignment,
    };

    #[test]
    fn section_tracks_page_size_orientation_and_margins() {
        let mut section = Section::new();
        section.set_page_size_twips(12_240, 15_840);
        section.set_page_orientation(PageOrientation::Landscape);
        let margins = section.page_margins_mut();
        margins.set_top_twips(1_440);
        margins.set_right_twips(720);
        margins.set_bottom_twips(1_440);
        margins.set_left_twips(720);

        assert_eq!(section.page_width_twips(), Some(12_240));
        assert_eq!(section.page_height_twips(), Some(15_840));
        assert_eq!(section.page_orientation(), Some(PageOrientation::Landscape));
        assert_eq!(section.page_margins().top_twips(), Some(1_440));
        assert_eq!(section.page_margins().right_twips(), Some(720));
        assert_eq!(section.page_margins().bottom_twips(), Some(1_440));
        assert_eq!(section.page_margins().left_twips(), Some(720));
    }

    #[test]
    fn margins_clear_resets_all_values() {
        let mut margins = PageMargins::new();
        margins.set_top_twips(1);
        margins.set_right_twips(2);
        margins.set_bottom_twips(3);
        margins.set_left_twips(4);
        margins.set_header_twips(5);
        margins.set_footer_twips(6);
        margins.set_gutter_twips(7);

        margins.clear();

        assert_eq!(margins.top_twips(), None);
        assert_eq!(margins.right_twips(), None);
        assert_eq!(margins.bottom_twips(), None);
        assert_eq!(margins.left_twips(), None);
        assert_eq!(margins.header_twips(), None);
        assert_eq!(margins.footer_twips(), None);
        assert_eq!(margins.gutter_twips(), None);
    }

    #[test]
    fn section_tracks_header_and_footer_content() {
        let mut section = Section::new();
        section
            .ensure_header()
            .add_paragraph("Header paragraph")
            .add_run("!");
        section.set_footer(HeaderFooter::from_text("Footer paragraph"));

        assert_eq!(
            section
                .header()
                .and_then(|header| header.paragraphs().first())
                .map(|paragraph| paragraph.text()),
            Some("Header paragraph!".to_string())
        );
        assert_eq!(
            section
                .footer()
                .and_then(|footer| footer.paragraphs().first())
                .map(|paragraph| paragraph.text()),
            Some("Footer paragraph".to_string())
        );

        section.clear_header();
        section.clear_footer();

        assert_eq!(section.header(), None);
        assert_eq!(section.footer(), None);
    }

    #[test]
    fn first_page_header_footer_and_title_page() {
        let mut section = Section::new();
        assert!(!section.title_page());
        assert_eq!(section.first_page_header(), None);
        assert_eq!(section.first_page_footer(), None);

        section
            .ensure_first_page_header()
            .add_paragraph("First page header");
        assert!(section.title_page());
        assert_eq!(
            section
                .first_page_header()
                .and_then(|hf| hf.paragraphs().first())
                .map(|p| p.text()),
            Some("First page header".to_string())
        );

        section.set_first_page_footer(HeaderFooter::from_text("First page footer"));
        assert_eq!(
            section
                .first_page_footer()
                .and_then(|hf| hf.paragraphs().first())
                .map(|p| p.text()),
            Some("First page footer".to_string())
        );

        section.clear_first_page_header();
        section.clear_first_page_footer();
        assert_eq!(section.first_page_header(), None);
        assert_eq!(section.first_page_footer(), None);
    }

    #[test]
    fn section_break_type_can_be_set_and_cleared() {
        let mut section = Section::new();
        assert_eq!(section.break_type(), None);

        section.set_break_type(SectionBreakType::Continuous);
        assert_eq!(section.break_type(), Some(SectionBreakType::Continuous));

        section.set_break_type(SectionBreakType::EvenPage);
        assert_eq!(section.break_type(), Some(SectionBreakType::EvenPage));

        section.clear_break_type();
        assert_eq!(section.break_type(), None);
    }

    #[test]
    fn section_break_type_xml_roundtrip() {
        assert_eq!(
            SectionBreakType::from_xml_value("nextPage"),
            Some(SectionBreakType::NextPage)
        );
        assert_eq!(
            SectionBreakType::from_xml_value("continuous"),
            Some(SectionBreakType::Continuous)
        );
        assert_eq!(
            SectionBreakType::from_xml_value("evenPage"),
            Some(SectionBreakType::EvenPage)
        );
        assert_eq!(
            SectionBreakType::from_xml_value("oddPage"),
            Some(SectionBreakType::OddPage)
        );
        assert_eq!(SectionBreakType::from_xml_value("invalid"), None);

        assert_eq!(SectionBreakType::NextPage.to_xml_value(), "nextPage");
        assert_eq!(SectionBreakType::Continuous.to_xml_value(), "continuous");
    }

    #[test]
    fn title_page_affects_has_properties() {
        let mut section = Section::new();
        assert!(!section.has_properties());

        section.set_title_page(true);
        assert!(section.has_properties());

        section.set_title_page(false);
        assert!(!section.has_properties());
    }

    #[test]
    fn even_page_header_and_footer_can_be_set_and_cleared() {
        let mut section = Section::new();
        assert_eq!(section.even_page_header(), None);
        assert_eq!(section.even_page_footer(), None);

        section
            .ensure_even_page_header()
            .add_paragraph("Even page header");
        assert_eq!(
            section
                .even_page_header()
                .and_then(|hf| hf.paragraphs().first())
                .map(|p| p.text()),
            Some("Even page header".to_string())
        );

        section.set_even_page_footer(HeaderFooter::from_text("Even page footer"));
        assert_eq!(
            section
                .even_page_footer()
                .and_then(|hf| hf.paragraphs().first())
                .map(|p| p.text()),
            Some("Even page footer".to_string())
        );

        section.clear_even_page_header();
        section.clear_even_page_footer();
        assert_eq!(section.even_page_header(), None);
        assert_eq!(section.even_page_footer(), None);
    }

    #[test]
    fn even_page_header_footer_affects_has_properties() {
        let mut section = Section::new();
        assert!(!section.has_properties());

        section.set_even_page_header(HeaderFooter::from_text("Even header"));
        assert!(section.has_properties());

        section.clear_even_page_header();
        assert!(!section.has_properties());

        section.set_even_page_footer(HeaderFooter::from_text("Even footer"));
        assert!(section.has_properties());

        section.clear_even_page_footer();
        assert!(!section.has_properties());
    }

    #[test]
    fn page_number_start_can_be_set_and_cleared() {
        let mut section = Section::new();
        assert_eq!(section.page_number_start(), None);

        section.set_page_number_start(5);
        assert_eq!(section.page_number_start(), Some(5));
        assert!(section.has_properties());

        section.clear_page_number_start();
        assert_eq!(section.page_number_start(), None);
    }

    #[test]
    fn page_number_format_can_be_set_and_cleared() {
        let mut section = Section::new();
        assert_eq!(section.page_number_format(), None);

        section.set_page_number_format("lowerRoman");
        assert_eq!(section.page_number_format(), Some("lowerRoman"));
        assert!(section.has_properties());

        section.clear_page_number_format();
        assert_eq!(section.page_number_format(), None);
    }

    // ---- Multi-column layout tests ----

    #[test]
    fn column_count_can_be_set_and_cleared() {
        let mut section = Section::new();
        assert_eq!(section.column_count(), None);

        section.set_column_count(3);
        assert_eq!(section.column_count(), Some(3));
        assert!(section.has_properties());

        section.clear_column_count();
        assert_eq!(section.column_count(), None);
    }

    #[test]
    fn column_space_twips_can_be_set_and_cleared() {
        let mut section = Section::new();
        assert_eq!(section.column_space_twips(), None);

        section.set_column_space_twips(720);
        assert_eq!(section.column_space_twips(), Some(720));
        assert!(section.has_properties());

        section.clear_column_space_twips();
        assert_eq!(section.column_space_twips(), None);
    }

    #[test]
    fn column_separator_can_be_toggled() {
        let mut section = Section::new();
        assert!(!section.column_separator());
        assert!(!section.has_properties());

        section.set_column_separator(true);
        assert!(section.column_separator());
        assert!(section.has_properties());

        section.set_column_separator(false);
        assert!(!section.column_separator());
    }

    // ---- Vertical alignment tests ----

    #[test]
    fn vertical_alignment_can_be_set_and_cleared() {
        let mut section = Section::new();
        assert_eq!(section.vertical_alignment(), None);

        section.set_vertical_alignment(SectionVerticalAlignment::Center);
        assert_eq!(
            section.vertical_alignment(),
            Some(SectionVerticalAlignment::Center)
        );
        assert!(section.has_properties());

        section.set_vertical_alignment(SectionVerticalAlignment::Bottom);
        assert_eq!(
            section.vertical_alignment(),
            Some(SectionVerticalAlignment::Bottom)
        );

        section.clear_vertical_alignment();
        assert_eq!(section.vertical_alignment(), None);
    }

    #[test]
    fn vertical_alignment_xml_roundtrip() {
        assert_eq!(
            SectionVerticalAlignment::from_xml_value("top"),
            Some(SectionVerticalAlignment::Top)
        );
        assert_eq!(
            SectionVerticalAlignment::from_xml_value("center"),
            Some(SectionVerticalAlignment::Center)
        );
        assert_eq!(
            SectionVerticalAlignment::from_xml_value("both"),
            Some(SectionVerticalAlignment::Justify)
        );
        assert_eq!(
            SectionVerticalAlignment::from_xml_value("bottom"),
            Some(SectionVerticalAlignment::Bottom)
        );
        assert_eq!(SectionVerticalAlignment::from_xml_value("invalid"), None);

        assert_eq!(SectionVerticalAlignment::Top.to_xml_value(), "top");
        assert_eq!(SectionVerticalAlignment::Center.to_xml_value(), "center");
        assert_eq!(SectionVerticalAlignment::Justify.to_xml_value(), "both");
        assert_eq!(SectionVerticalAlignment::Bottom.to_xml_value(), "bottom");
    }

    // ---- Line numbering tests ----

    #[test]
    fn line_numbering_start_can_be_set_and_cleared() {
        let mut section = Section::new();
        assert_eq!(section.line_numbering_start(), None);

        section.set_line_numbering_start(1);
        assert_eq!(section.line_numbering_start(), Some(1));
        assert!(section.has_properties());

        section.clear_line_numbering_start();
        assert_eq!(section.line_numbering_start(), None);
    }

    #[test]
    fn line_numbering_count_by_can_be_set_and_cleared() {
        let mut section = Section::new();
        assert_eq!(section.line_numbering_count_by(), None);

        section.set_line_numbering_count_by(5);
        assert_eq!(section.line_numbering_count_by(), Some(5));
        assert!(section.has_properties());

        section.clear_line_numbering_count_by();
        assert_eq!(section.line_numbering_count_by(), None);
    }

    #[test]
    fn line_numbering_restart_can_be_set_and_cleared() {
        let mut section = Section::new();
        assert_eq!(section.line_numbering_restart(), None);

        section.set_line_numbering_restart(LineNumberRestart::NewPage);
        assert_eq!(
            section.line_numbering_restart(),
            Some(LineNumberRestart::NewPage)
        );
        assert!(section.has_properties());

        section.set_line_numbering_restart(LineNumberRestart::Continuous);
        assert_eq!(
            section.line_numbering_restart(),
            Some(LineNumberRestart::Continuous)
        );

        section.clear_line_numbering_restart();
        assert_eq!(section.line_numbering_restart(), None);
    }

    #[test]
    fn line_numbering_distance_twips_can_be_set_and_cleared() {
        let mut section = Section::new();
        assert_eq!(section.line_numbering_distance_twips(), None);

        section.set_line_numbering_distance_twips(360);
        assert_eq!(section.line_numbering_distance_twips(), Some(360));
        assert!(section.has_properties());

        section.clear_line_numbering_distance_twips();
        assert_eq!(section.line_numbering_distance_twips(), None);
    }

    #[test]
    fn line_number_restart_xml_roundtrip() {
        assert_eq!(
            LineNumberRestart::from_xml_value("newPage"),
            Some(LineNumberRestart::NewPage)
        );
        assert_eq!(
            LineNumberRestart::from_xml_value("newSection"),
            Some(LineNumberRestart::NewSection)
        );
        assert_eq!(
            LineNumberRestart::from_xml_value("continuous"),
            Some(LineNumberRestart::Continuous)
        );
        assert_eq!(LineNumberRestart::from_xml_value("invalid"), None);

        assert_eq!(LineNumberRestart::NewPage.to_xml_value(), "newPage");
        assert_eq!(LineNumberRestart::NewSection.to_xml_value(), "newSection");
        assert_eq!(LineNumberRestart::Continuous.to_xml_value(), "continuous");
    }

    // ---- Link to previous tests ----

    #[test]
    fn link_to_previous_defaults_to_true() {
        let section = Section::new();
        assert!(section.header_link_to_previous());
        assert!(section.footer_link_to_previous());
        assert!(section.first_header_link_to_previous());
        assert!(section.first_footer_link_to_previous());
        assert!(section.even_header_link_to_previous());
        assert!(section.even_footer_link_to_previous());
    }

    #[test]
    fn link_to_previous_does_not_affect_has_properties_when_true() {
        let section = Section::new();
        // All link-to-previous fields default to true, so they should not trigger has_properties
        assert!(!section.has_properties());
    }

    #[test]
    fn header_link_to_previous_can_be_set() {
        let mut section = Section::new();

        section.set_header_link_to_previous(false);
        assert!(!section.header_link_to_previous());
        assert!(section.has_properties());

        section.set_header_link_to_previous(true);
        assert!(section.header_link_to_previous());
    }

    #[test]
    fn footer_link_to_previous_can_be_set() {
        let mut section = Section::new();

        section.set_footer_link_to_previous(false);
        assert!(!section.footer_link_to_previous());
        assert!(section.has_properties());

        section.set_footer_link_to_previous(true);
        assert!(section.footer_link_to_previous());
    }

    #[test]
    fn first_header_link_to_previous_can_be_set() {
        let mut section = Section::new();

        section.set_first_header_link_to_previous(false);
        assert!(!section.first_header_link_to_previous());
        assert!(section.has_properties());

        section.set_first_header_link_to_previous(true);
        assert!(section.first_header_link_to_previous());
    }

    #[test]
    fn first_footer_link_to_previous_can_be_set() {
        let mut section = Section::new();

        section.set_first_footer_link_to_previous(false);
        assert!(!section.first_footer_link_to_previous());
        assert!(section.has_properties());

        section.set_first_footer_link_to_previous(true);
        assert!(section.first_footer_link_to_previous());
    }

    #[test]
    fn even_header_link_to_previous_can_be_set() {
        let mut section = Section::new();

        section.set_even_header_link_to_previous(false);
        assert!(!section.even_header_link_to_previous());
        assert!(section.has_properties());

        section.set_even_header_link_to_previous(true);
        assert!(section.even_header_link_to_previous());
    }

    #[test]
    fn even_footer_link_to_previous_can_be_set() {
        let mut section = Section::new();

        section.set_even_footer_link_to_previous(false);
        assert!(!section.even_footer_link_to_previous());
        assert!(section.has_properties());

        section.set_even_footer_link_to_previous(true);
        assert!(section.even_footer_link_to_previous());
    }
}
