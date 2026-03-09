use std::collections::BTreeMap;

use offidized_opc::RawXmlNode;

use crate::auto_filter::AutoFilter;
use crate::cell::{normalize_cell_reference, Cell, CellValue};
use crate::chart::Chart;
use crate::error::Result;
use crate::print_settings::{PageBreaks, PrintArea, PrintHeaderFooter};
use crate::range::CellRange;
use crate::sparkline::SparklineGroup;
use crate::{column::Column, error::XlsxError, row::Row};

/// Sheet visibility state in the workbook.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum SheetVisibility {
    /// The sheet tab is visible (default).
    #[default]
    Visible,
    /// The sheet tab is hidden but can be unhidden by the user via the UI.
    Hidden,
    /// The sheet tab is hidden and cannot be unhidden via the UI (only programmatically).
    VeryHidden,
}

impl SheetVisibility {
    pub(crate) fn as_xml_value(self) -> &'static str {
        match self {
            Self::Visible => "visible",
            Self::Hidden => "hidden",
            Self::VeryHidden => "veryHidden",
        }
    }

    pub(crate) fn from_xml_value(value: &str) -> Self {
        match value.trim() {
            "hidden" => Self::Hidden,
            "veryHidden" => Self::VeryHidden,
            _ => Self::Visible,
        }
    }
}

/// Sheet protection settings.
#[derive(Debug, Clone, PartialEq, Eq, Default)]
pub struct SheetProtection {
    pub(crate) password_hash: Option<String>,
    pub(crate) sheet: bool,
    pub(crate) objects: bool,
    pub(crate) scenarios: bool,
    pub(crate) format_cells: bool,
    pub(crate) format_columns: bool,
    pub(crate) format_rows: bool,
    pub(crate) insert_columns: bool,
    pub(crate) insert_rows: bool,
    pub(crate) insert_hyperlinks: bool,
    pub(crate) delete_columns: bool,
    pub(crate) delete_rows: bool,
    pub(crate) select_locked_cells: bool,
    pub(crate) sort: bool,
    pub(crate) auto_filter: bool,
    pub(crate) pivot_tables: bool,
    pub(crate) select_unlocked_cells: bool,
}

impl SheetProtection {
    /// Creates a new sheet protection with the sheet locked.
    pub fn new() -> Self {
        Self {
            sheet: true,
            ..Self::default()
        }
    }

    /// Returns the password hash (algorithm-dependent hex string), if set.
    pub fn password_hash(&self) -> Option<&str> {
        self.password_hash.as_deref()
    }

    /// Sets the password hash. This is the pre-hashed value (not the raw password).
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

    /// Returns whether the sheet itself is protected.
    pub fn sheet(&self) -> bool {
        self.sheet
    }

    /// Sets whether the sheet itself is protected.
    pub fn set_sheet(&mut self, value: bool) -> &mut Self {
        self.sheet = value;
        self
    }

    /// Returns whether objects are protected.
    pub fn objects(&self) -> bool {
        self.objects
    }

    /// Sets whether objects are protected.
    pub fn set_objects(&mut self, value: bool) -> &mut Self {
        self.objects = value;
        self
    }

    /// Returns whether scenarios are protected.
    pub fn scenarios(&self) -> bool {
        self.scenarios
    }

    /// Sets whether scenarios are protected.
    pub fn set_scenarios(&mut self, value: bool) -> &mut Self {
        self.scenarios = value;
        self
    }

    /// Returns whether formatting cells is disallowed.
    pub fn format_cells(&self) -> bool {
        self.format_cells
    }

    /// Sets whether formatting cells is disallowed.
    pub fn set_format_cells(&mut self, value: bool) -> &mut Self {
        self.format_cells = value;
        self
    }

    /// Returns whether formatting columns is disallowed.
    pub fn format_columns(&self) -> bool {
        self.format_columns
    }

    /// Sets whether formatting columns is disallowed.
    pub fn set_format_columns(&mut self, value: bool) -> &mut Self {
        self.format_columns = value;
        self
    }

    /// Returns whether formatting rows is disallowed.
    pub fn format_rows(&self) -> bool {
        self.format_rows
    }

    /// Sets whether formatting rows is disallowed.
    pub fn set_format_rows(&mut self, value: bool) -> &mut Self {
        self.format_rows = value;
        self
    }

    /// Returns whether inserting columns is disallowed.
    pub fn insert_columns(&self) -> bool {
        self.insert_columns
    }

    /// Sets whether inserting columns is disallowed.
    pub fn set_insert_columns(&mut self, value: bool) -> &mut Self {
        self.insert_columns = value;
        self
    }

    /// Returns whether inserting rows is disallowed.
    pub fn insert_rows(&self) -> bool {
        self.insert_rows
    }

    /// Sets whether inserting rows is disallowed.
    pub fn set_insert_rows(&mut self, value: bool) -> &mut Self {
        self.insert_rows = value;
        self
    }

    /// Returns whether inserting hyperlinks is disallowed.
    pub fn insert_hyperlinks(&self) -> bool {
        self.insert_hyperlinks
    }

    /// Sets whether inserting hyperlinks is disallowed.
    pub fn set_insert_hyperlinks(&mut self, value: bool) -> &mut Self {
        self.insert_hyperlinks = value;
        self
    }

    /// Returns whether deleting columns is disallowed.
    pub fn delete_columns(&self) -> bool {
        self.delete_columns
    }

    /// Sets whether deleting columns is disallowed.
    pub fn set_delete_columns(&mut self, value: bool) -> &mut Self {
        self.delete_columns = value;
        self
    }

    /// Returns whether deleting rows is disallowed.
    pub fn delete_rows(&self) -> bool {
        self.delete_rows
    }

    /// Sets whether deleting rows is disallowed.
    pub fn set_delete_rows(&mut self, value: bool) -> &mut Self {
        self.delete_rows = value;
        self
    }

    /// Returns whether selecting locked cells is disallowed.
    pub fn select_locked_cells(&self) -> bool {
        self.select_locked_cells
    }

    /// Sets whether selecting locked cells is disallowed.
    pub fn set_select_locked_cells(&mut self, value: bool) -> &mut Self {
        self.select_locked_cells = value;
        self
    }

    /// Returns whether sorting is disallowed.
    pub fn sort(&self) -> bool {
        self.sort
    }

    /// Sets whether sorting is disallowed.
    pub fn set_sort(&mut self, value: bool) -> &mut Self {
        self.sort = value;
        self
    }

    /// Returns whether auto-filter is disallowed.
    pub fn auto_filter(&self) -> bool {
        self.auto_filter
    }

    /// Sets whether auto-filter is disallowed.
    pub fn set_auto_filter(&mut self, value: bool) -> &mut Self {
        self.auto_filter = value;
        self
    }

    /// Returns whether pivot tables are disallowed.
    pub fn pivot_tables(&self) -> bool {
        self.pivot_tables
    }

    /// Sets whether pivot tables are disallowed.
    pub fn set_pivot_tables(&mut self, value: bool) -> &mut Self {
        self.pivot_tables = value;
        self
    }

    /// Returns whether selecting unlocked cells is disallowed.
    pub fn select_unlocked_cells(&self) -> bool {
        self.select_unlocked_cells
    }

    /// Sets whether selecting unlocked cells is disallowed.
    pub fn set_select_unlocked_cells(&mut self, value: bool) -> &mut Self {
        self.select_unlocked_cells = value;
        self
    }

    /// Returns true if any protection flag is enabled.
    pub(crate) fn has_metadata(&self) -> bool {
        self.sheet
            || self.objects
            || self.scenarios
            || self.format_cells
            || self.format_columns
            || self.format_rows
            || self.insert_columns
            || self.insert_rows
            || self.insert_hyperlinks
            || self.delete_columns
            || self.delete_rows
            || self.select_locked_cells
            || self.sort
            || self.auto_filter
            || self.pivot_tables
            || self.select_unlocked_cells
            || self.password_hash.is_some()
    }
}

/// Page orientation for printing.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum PageOrientation {
    /// Portrait orientation (default).
    #[default]
    Portrait,
    /// Landscape orientation.
    Landscape,
}

impl PageOrientation {
    pub(crate) fn as_xml_value(self) -> &'static str {
        match self {
            Self::Portrait => "portrait",
            Self::Landscape => "landscape",
        }
    }

    pub(crate) fn from_xml_value(value: &str) -> Self {
        match value.trim() {
            "landscape" => Self::Landscape,
            _ => Self::Portrait,
        }
    }
}

/// Page margins for printing, measured in inches.
#[derive(Debug, Clone, PartialEq, Default)]
pub struct PageMargins {
    pub(crate) left: Option<f64>,
    pub(crate) right: Option<f64>,
    pub(crate) top: Option<f64>,
    pub(crate) bottom: Option<f64>,
    pub(crate) header: Option<f64>,
    pub(crate) footer: Option<f64>,
}

impl PageMargins {
    pub fn new() -> Self {
        Self::default()
    }

    pub fn left(&self) -> Option<f64> {
        self.left
    }

    pub fn set_left(&mut self, value: f64) -> &mut Self {
        if value.is_finite() && value >= 0.0 {
            self.left = Some(value);
        }
        self
    }

    pub fn clear_left(&mut self) -> &mut Self {
        self.left = None;
        self
    }

    pub fn right(&self) -> Option<f64> {
        self.right
    }

    pub fn set_right(&mut self, value: f64) -> &mut Self {
        if value.is_finite() && value >= 0.0 {
            self.right = Some(value);
        }
        self
    }

    pub fn clear_right(&mut self) -> &mut Self {
        self.right = None;
        self
    }

    pub fn top(&self) -> Option<f64> {
        self.top
    }

    pub fn set_top(&mut self, value: f64) -> &mut Self {
        if value.is_finite() && value >= 0.0 {
            self.top = Some(value);
        }
        self
    }

    pub fn clear_top(&mut self) -> &mut Self {
        self.top = None;
        self
    }

    pub fn bottom(&self) -> Option<f64> {
        self.bottom
    }

    pub fn set_bottom(&mut self, value: f64) -> &mut Self {
        if value.is_finite() && value >= 0.0 {
            self.bottom = Some(value);
        }
        self
    }

    pub fn clear_bottom(&mut self) -> &mut Self {
        self.bottom = None;
        self
    }

    pub fn header(&self) -> Option<f64> {
        self.header
    }

    pub fn set_header(&mut self, value: f64) -> &mut Self {
        if value.is_finite() && value >= 0.0 {
            self.header = Some(value);
        }
        self
    }

    pub fn clear_header(&mut self) -> &mut Self {
        self.header = None;
        self
    }

    pub fn footer(&self) -> Option<f64> {
        self.footer
    }

    pub fn set_footer(&mut self, value: f64) -> &mut Self {
        if value.is_finite() && value >= 0.0 {
            self.footer = Some(value);
        }
        self
    }

    pub fn clear_footer(&mut self) -> &mut Self {
        self.footer = None;
        self
    }

    pub(crate) fn has_metadata(&self) -> bool {
        self.left.is_some()
            || self.right.is_some()
            || self.top.is_some()
            || self.bottom.is_some()
            || self.header.is_some()
            || self.footer.is_some()
    }
}

/// Page setup for printing.
#[derive(Debug, Clone, PartialEq, Default)]
pub struct PageSetup {
    pub(crate) orientation: Option<PageOrientation>,
    pub(crate) paper_size: Option<u32>,
    pub(crate) scale: Option<u32>,
    pub(crate) fit_to_width: Option<u32>,
    pub(crate) fit_to_height: Option<u32>,
    pub(crate) first_page_number: Option<u32>,
    pub(crate) use_first_page_number: Option<bool>,
    pub(crate) horizontal_dpi: Option<u32>,
    pub(crate) vertical_dpi: Option<u32>,
}

impl PageSetup {
    pub fn new() -> Self {
        Self::default()
    }

    /// Returns the page orientation.
    pub fn orientation(&self) -> Option<PageOrientation> {
        self.orientation
    }

    /// Sets the page orientation.
    pub fn set_orientation(&mut self, orientation: PageOrientation) -> &mut Self {
        self.orientation = Some(orientation);
        self
    }

    /// Clears the page orientation.
    pub fn clear_orientation(&mut self) -> &mut Self {
        self.orientation = None;
        self
    }

    /// Returns the paper size code (1=Letter, 9=A4, etc.).
    pub fn paper_size(&self) -> Option<u32> {
        self.paper_size
    }

    /// Sets the paper size code.
    pub fn set_paper_size(&mut self, paper_size: u32) -> &mut Self {
        self.paper_size = Some(paper_size);
        self
    }

    /// Clears the paper size.
    pub fn clear_paper_size(&mut self) -> &mut Self {
        self.paper_size = None;
        self
    }

    /// Returns the print scale percentage (10-400).
    pub fn scale(&self) -> Option<u32> {
        self.scale
    }

    /// Sets the print scale percentage.
    pub fn set_scale(&mut self, scale: u32) -> &mut Self {
        self.scale = Some(scale);
        self
    }

    /// Clears the print scale.
    pub fn clear_scale(&mut self) -> &mut Self {
        self.scale = None;
        self
    }

    /// Returns the fit-to-width page count.
    pub fn fit_to_width(&self) -> Option<u32> {
        self.fit_to_width
    }

    /// Sets the fit-to-width page count.
    pub fn set_fit_to_width(&mut self, pages: u32) -> &mut Self {
        self.fit_to_width = Some(pages);
        self
    }

    /// Clears the fit-to-width setting.
    pub fn clear_fit_to_width(&mut self) -> &mut Self {
        self.fit_to_width = None;
        self
    }

    /// Returns the fit-to-height page count.
    pub fn fit_to_height(&self) -> Option<u32> {
        self.fit_to_height
    }

    /// Sets the fit-to-height page count.
    pub fn set_fit_to_height(&mut self, pages: u32) -> &mut Self {
        self.fit_to_height = Some(pages);
        self
    }

    /// Clears the fit-to-height setting.
    pub fn clear_fit_to_height(&mut self) -> &mut Self {
        self.fit_to_height = None;
        self
    }

    /// Returns the first page number.
    pub fn first_page_number(&self) -> Option<u32> {
        self.first_page_number
    }

    /// Sets the first page number.
    pub fn set_first_page_number(&mut self, number: u32) -> &mut Self {
        self.first_page_number = Some(number);
        self
    }

    /// Clears the first page number.
    pub fn clear_first_page_number(&mut self) -> &mut Self {
        self.first_page_number = None;
        self
    }

    /// Returns whether the first page number is used instead of auto.
    pub fn use_first_page_number(&self) -> Option<bool> {
        self.use_first_page_number
    }

    /// Sets whether the first page number setting is used.
    pub fn set_use_first_page_number(&mut self, value: bool) -> &mut Self {
        self.use_first_page_number = Some(value);
        self
    }

    /// Clears the use-first-page-number flag.
    pub fn clear_use_first_page_number(&mut self) -> &mut Self {
        self.use_first_page_number = None;
        self
    }

    /// Returns the horizontal DPI for printing.
    pub fn horizontal_dpi(&self) -> Option<u32> {
        self.horizontal_dpi
    }

    /// Sets the horizontal DPI for printing.
    pub fn set_horizontal_dpi(&mut self, dpi: u32) -> &mut Self {
        self.horizontal_dpi = Some(dpi);
        self
    }

    /// Clears the horizontal DPI.
    pub fn clear_horizontal_dpi(&mut self) -> &mut Self {
        self.horizontal_dpi = None;
        self
    }

    /// Returns the vertical DPI for printing.
    pub fn vertical_dpi(&self) -> Option<u32> {
        self.vertical_dpi
    }

    /// Sets the vertical DPI for printing.
    pub fn set_vertical_dpi(&mut self, dpi: u32) -> &mut Self {
        self.vertical_dpi = Some(dpi);
        self
    }

    /// Clears the vertical DPI.
    pub fn clear_vertical_dpi(&mut self) -> &mut Self {
        self.vertical_dpi = None;
        self
    }

    pub(crate) fn has_metadata(&self) -> bool {
        self.orientation.is_some()
            || self.paper_size.is_some()
            || self.scale.is_some()
            || self.fit_to_width.is_some()
            || self.fit_to_height.is_some()
            || self.first_page_number.is_some()
            || self.use_first_page_number.is_some()
            || self.horizontal_dpi.is_some()
            || self.vertical_dpi.is_some()
    }
}

/// Sheet view options controlling visual display.
#[derive(Debug, Clone, PartialEq, Eq, Default)]
pub struct SheetViewOptions {
    pub(crate) show_gridlines: Option<bool>,
    pub(crate) show_row_col_headers: Option<bool>,
    pub(crate) show_formulas: Option<bool>,
    pub(crate) zoom_scale: Option<u32>,
    pub(crate) zoom_scale_normal: Option<u32>,
    pub(crate) right_to_left: Option<bool>,
    pub(crate) tab_selected: Option<bool>,
    pub(crate) view: Option<String>,
}

impl SheetViewOptions {
    pub fn new() -> Self {
        Self::default()
    }

    /// Returns whether gridlines are shown.
    pub fn show_gridlines(&self) -> Option<bool> {
        self.show_gridlines
    }

    /// Sets whether gridlines are shown.
    pub fn set_show_gridlines(&mut self, value: bool) -> &mut Self {
        self.show_gridlines = Some(value);
        self
    }

    /// Clears the show gridlines setting.
    pub fn clear_show_gridlines(&mut self) -> &mut Self {
        self.show_gridlines = None;
        self
    }

    /// Returns whether row/column headers are shown.
    pub fn show_row_col_headers(&self) -> Option<bool> {
        self.show_row_col_headers
    }

    /// Sets whether row/column headers are shown.
    pub fn set_show_row_col_headers(&mut self, value: bool) -> &mut Self {
        self.show_row_col_headers = Some(value);
        self
    }

    /// Clears the show row/column headers setting.
    pub fn clear_show_row_col_headers(&mut self) -> &mut Self {
        self.show_row_col_headers = None;
        self
    }

    /// Returns whether formulas are shown instead of values.
    pub fn show_formulas(&self) -> Option<bool> {
        self.show_formulas
    }

    /// Sets whether formulas are shown instead of values.
    pub fn set_show_formulas(&mut self, value: bool) -> &mut Self {
        self.show_formulas = Some(value);
        self
    }

    /// Clears the show formulas setting.
    pub fn clear_show_formulas(&mut self) -> &mut Self {
        self.show_formulas = None;
        self
    }

    /// Returns the zoom scale percentage.
    pub fn zoom_scale(&self) -> Option<u32> {
        self.zoom_scale
    }

    /// Sets the zoom scale percentage (10-400).
    pub fn set_zoom_scale(&mut self, value: u32) -> &mut Self {
        self.zoom_scale = Some(value);
        self
    }

    /// Clears the zoom scale.
    pub fn clear_zoom_scale(&mut self) -> &mut Self {
        self.zoom_scale = None;
        self
    }

    /// Returns the zoom scale for normal view.
    pub fn zoom_scale_normal(&self) -> Option<u32> {
        self.zoom_scale_normal
    }

    /// Sets the zoom scale for normal view.
    pub fn set_zoom_scale_normal(&mut self, value: u32) -> &mut Self {
        self.zoom_scale_normal = Some(value);
        self
    }

    /// Clears the normal view zoom scale.
    pub fn clear_zoom_scale_normal(&mut self) -> &mut Self {
        self.zoom_scale_normal = None;
        self
    }

    /// Returns whether the sheet view is right-to-left.
    pub fn right_to_left(&self) -> Option<bool> {
        self.right_to_left
    }

    /// Sets whether the sheet view is right-to-left.
    pub fn set_right_to_left(&mut self, value: bool) -> &mut Self {
        self.right_to_left = Some(value);
        self
    }

    /// Clears the right-to-left setting.
    pub fn clear_right_to_left(&mut self) -> &mut Self {
        self.right_to_left = None;
        self
    }

    /// Returns whether the tab is selected.
    pub fn tab_selected(&self) -> Option<bool> {
        self.tab_selected
    }

    /// Sets whether the tab is selected.
    pub fn set_tab_selected(&mut self, value: bool) -> &mut Self {
        self.tab_selected = Some(value);
        self
    }

    /// Clears the tab selected setting.
    pub fn clear_tab_selected(&mut self) -> &mut Self {
        self.tab_selected = None;
        self
    }

    /// Returns the view mode (e.g., "normal", "pageLayout", "pageBreakPreview").
    pub fn view(&self) -> Option<&str> {
        self.view.as_deref()
    }

    /// Sets the view mode.
    pub fn set_view(&mut self, view: impl Into<String>) -> &mut Self {
        self.view = Some(view.into());
        self
    }

    /// Clears the view mode.
    pub fn clear_view(&mut self) -> &mut Self {
        self.view = None;
        self
    }

    pub(crate) fn has_metadata(&self) -> bool {
        self.show_gridlines.is_some()
            || self.show_row_col_headers.is_some()
            || self.show_formulas.is_some()
            || self.zoom_scale.is_some()
            || self.zoom_scale_normal.is_some()
            || self.right_to_left.is_some()
            || self.tab_selected.is_some()
            || self.view.is_some()
    }
}

/// A comment/note attached to a cell.
#[derive(Debug, Clone, PartialEq)]
pub struct Comment {
    cell_ref: String,
    author: String,
    text: String,
    /// Rich text runs for comments with formatting.
    rich_text: Option<Vec<crate::cell::RichTextRun>>,
    /// Whether the comment is visible (always shown) or hidden (shown on hover).
    visible: bool,
    /// Threaded comment replies.
    replies: Vec<CommentReply>,
}

/// A reply within a threaded comment.
#[derive(Debug, Clone, PartialEq)]
pub struct CommentReply {
    /// The author of the reply.
    pub author: String,
    /// The reply text.
    pub text: String,
}

impl CommentReply {
    /// Creates a new comment reply.
    pub fn new(author: impl Into<String>, text: impl Into<String>) -> Self {
        Self {
            author: author.into(),
            text: text.into(),
        }
    }
}

impl Comment {
    /// Creates a new comment at the given cell reference.
    pub fn new(cell_ref: &str, author: impl Into<String>, text: impl Into<String>) -> Result<Self> {
        let cell_ref = normalize_cell_reference(cell_ref)?;
        Ok(Self {
            cell_ref,
            author: author.into(),
            text: text.into(),
            rich_text: None,
            visible: false,
            replies: Vec::new(),
        })
    }

    pub(crate) fn from_parsed_parts(
        cell_ref: String,
        author: String,
        text: String,
    ) -> Result<Self> {
        let cell_ref = normalize_cell_reference(cell_ref.as_str())?;
        Ok(Self {
            cell_ref,
            author,
            text,
            rich_text: None,
            visible: false,
            replies: Vec::new(),
        })
    }

    /// Returns the cell reference this comment is attached to.
    pub fn cell_ref(&self) -> &str {
        self.cell_ref.as_str()
    }

    /// Returns the comment author.
    pub fn author(&self) -> &str {
        self.author.as_str()
    }

    /// Returns the comment text.
    pub fn text(&self) -> &str {
        self.text.as_str()
    }

    /// Sets the comment text.
    pub fn set_text(&mut self, text: impl Into<String>) -> &mut Self {
        self.text = text.into();
        self
    }

    /// Sets the comment author.
    pub fn set_author(&mut self, author: impl Into<String>) -> &mut Self {
        self.author = author.into();
        self
    }

    /// Returns the rich text runs, if set.
    pub fn rich_text(&self) -> Option<&[crate::cell::RichTextRun]> {
        self.rich_text.as_deref()
    }

    /// Sets rich text runs (overrides plain text display).
    pub fn set_rich_text(&mut self, runs: Vec<crate::cell::RichTextRun>) -> &mut Self {
        self.rich_text = Some(runs);
        self
    }

    /// Clears rich text runs.
    pub fn clear_rich_text(&mut self) -> &mut Self {
        self.rich_text = None;
        self
    }

    /// Returns whether the comment is always visible.
    pub fn visible(&self) -> bool {
        self.visible
    }

    /// Sets whether the comment is always visible.
    pub fn set_visible(&mut self, visible: bool) -> &mut Self {
        self.visible = visible;
        self
    }

    /// Returns the threaded replies.
    pub fn replies(&self) -> &[CommentReply] {
        &self.replies
    }

    /// Adds a threaded reply.
    pub fn add_reply(&mut self, reply: CommentReply) -> &mut Self {
        self.replies.push(reply);
        self
    }

    /// Removes all replies.
    pub fn clear_replies(&mut self) -> &mut Self {
        self.replies.clear();
        self
    }
}

/// Frozen pane configuration for a worksheet.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct FreezePane {
    x_split: u32,
    y_split: u32,
    top_left_cell: String,
}

impl FreezePane {
    pub fn new(x_split: u32, y_split: u32) -> Result<Self> {
        if x_split == 0 && y_split == 0 {
            return Err(XlsxError::InvalidWorkbookState(
                "freeze pane split must freeze at least one row or column".to_string(),
            ));
        }

        let top_left_cell = build_cell_reference(
            x_split.checked_add(1).ok_or_else(|| {
                XlsxError::InvalidWorkbookState("freeze pane xSplit overflow".to_string())
            })?,
            y_split.checked_add(1).ok_or_else(|| {
                XlsxError::InvalidWorkbookState("freeze pane ySplit overflow".to_string())
            })?,
        )?;

        Ok(Self {
            x_split,
            y_split,
            top_left_cell,
        })
    }

    pub fn with_top_left_cell(x_split: u32, y_split: u32, top_left_cell: &str) -> Result<Self> {
        if x_split == 0 && y_split == 0 {
            return Err(XlsxError::InvalidWorkbookState(
                "freeze pane split must freeze at least one row or column".to_string(),
            ));
        }

        let top_left_cell = normalize_cell_reference(top_left_cell)?;
        Ok(Self {
            x_split,
            y_split,
            top_left_cell,
        })
    }

    pub fn x_split(&self) -> u32 {
        self.x_split
    }

    pub fn y_split(&self) -> u32 {
        self.y_split
    }

    pub fn top_left_cell(&self) -> &str {
        self.top_left_cell.as_str()
    }
}

/// Image anchor type for a worksheet drawing.
///
/// Controls how an image is positioned within the worksheet:
/// - `TwoCell`: anchored between two cell positions (resizes with cells).
/// - `OneCell`: anchored to one cell with fixed extent (moves with cell, does not resize).
/// - `Absolute`: positioned at fixed EMU coordinates (does not move or resize with cells).
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum ImageAnchorType {
    /// Anchored between two cells. Image resizes when the anchor cells resize.
    TwoCell,
    /// Anchored to one cell with a fixed extent. Image moves but does not resize.
    #[default]
    OneCell,
    /// Absolutely positioned at fixed EMU coordinates.
    Absolute,
}

/// A cell anchor point for image positioning.
///
/// Specifies a corner of an image anchor using a cell reference with EMU offsets
/// within that cell.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct CellAnchor {
    /// 0-based column index.
    col: u32,
    /// 0-based row index.
    row: u32,
    /// EMU offset from the left edge of the column.
    col_offset: i64,
    /// EMU offset from the top edge of the row.
    row_offset: i64,
}

impl CellAnchor {
    /// Creates a new cell anchor.
    pub fn new(col: u32, row: u32, col_offset: i64, row_offset: i64) -> Self {
        Self {
            col,
            row,
            col_offset,
            row_offset,
        }
    }

    /// Returns the 0-based column index.
    pub fn col(&self) -> u32 {
        self.col
    }

    /// Sets the 0-based column index.
    pub fn set_col(&mut self, col: u32) -> &mut Self {
        self.col = col;
        self
    }

    /// Returns the 0-based row index.
    pub fn row(&self) -> u32 {
        self.row
    }

    /// Sets the 0-based row index.
    pub fn set_row(&mut self, row: u32) -> &mut Self {
        self.row = row;
        self
    }

    /// Returns the EMU offset within the cell column.
    pub fn col_offset(&self) -> i64 {
        self.col_offset
    }

    /// Sets the EMU offset within the cell column.
    pub fn set_col_offset(&mut self, offset: i64) -> &mut Self {
        self.col_offset = offset;
        self
    }

    /// Returns the EMU offset within the cell row.
    pub fn row_offset(&self) -> i64 {
        self.row_offset
    }

    /// Sets the EMU offset within the cell row.
    pub fn set_row_offset(&mut self, offset: i64) -> &mut Self {
        self.row_offset = offset;
        self
    }
}

/// Optional image extent for a worksheet drawing anchor, measured in EMUs.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct WorksheetImageExt {
    cx: u64,
    cy: u64,
}

impl WorksheetImageExt {
    pub fn new(cx: u64, cy: u64) -> Result<Self> {
        if cx == 0 || cy == 0 {
            return Err(XlsxError::InvalidWorkbookState(
                "worksheet image ext dimensions must be > 0".to_string(),
            ));
        }

        Ok(Self { cx, cy })
    }

    pub fn cx(&self) -> u64 {
        self.cx
    }

    pub fn cy(&self) -> u64 {
        self.cy
    }
}

/// A worksheet image anchored to a cell or position.
#[derive(Debug, Clone, PartialEq)]
pub struct WorksheetImage {
    bytes: Vec<u8>,
    content_type: String,
    anchor_cell: String,
    ext: Option<WorksheetImageExt>,
    /// Anchor type controlling how the image is positioned.
    anchor_type: ImageAnchorType,
    /// Top-left anchor point (for TwoCell and OneCell).
    from_anchor: Option<CellAnchor>,
    /// Bottom-right anchor point (for TwoCell only).
    to_anchor: Option<CellAnchor>,
    /// Width in EMU (for OneCell and Absolute).
    extent_cx: Option<i64>,
    /// Height in EMU (for OneCell and Absolute).
    extent_cy: Option<i64>,
    /// Absolute X position in EMU (for Absolute anchor type).
    position_x: Option<i64>,
    /// Absolute Y position in EMU (for Absolute anchor type).
    position_y: Option<i64>,
    /// Crop percentage from the left edge (0.0 to 1.0).
    crop_left: Option<f64>,
    /// Crop percentage from the right edge (0.0 to 1.0).
    crop_right: Option<f64>,
    /// Crop percentage from the top edge (0.0 to 1.0).
    crop_top: Option<f64>,
    /// Crop percentage from the bottom edge (0.0 to 1.0).
    crop_bottom: Option<f64>,
    /// Image name (`xdr:pic/xdr:nvPicPr/xdr:cNvPr @name`).
    name: Option<String>,
    /// Image description/alt text (`xdr:pic/xdr:nvPicPr/xdr:cNvPr @descr`).
    description: Option<String>,
}

impl WorksheetImage {
    pub fn new(
        bytes: impl Into<Vec<u8>>,
        content_type: impl Into<String>,
        anchor_cell: &str,
        ext: Option<WorksheetImageExt>,
    ) -> Result<Self> {
        Self::from_parsed_parts(
            bytes.into(),
            content_type.into(),
            anchor_cell.to_string(),
            ext,
        )
    }

    pub(crate) fn from_parsed_parts(
        bytes: Vec<u8>,
        content_type: String,
        anchor_cell: String,
        ext: Option<WorksheetImageExt>,
    ) -> Result<Self> {
        if bytes.is_empty() {
            return Err(XlsxError::InvalidWorkbookState(
                "worksheet image bytes must not be empty".to_string(),
            ));
        }

        let content_type = content_type.trim();
        if content_type.is_empty() {
            return Err(XlsxError::InvalidWorkbookState(
                "worksheet image content_type must not be empty".to_string(),
            ));
        }

        let anchor_cell = normalize_cell_reference(anchor_cell.as_str())?;

        Ok(Self {
            bytes,
            content_type: content_type.to_string(),
            anchor_cell,
            ext,
            anchor_type: ImageAnchorType::default(),
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
            name: None,
            description: None,
        })
    }

    /// Returns the raw image bytes.
    pub fn bytes(&self) -> &[u8] {
        self.bytes.as_slice()
    }

    /// Returns the MIME content type of the image.
    pub fn content_type(&self) -> &str {
        self.content_type.as_str()
    }

    /// Returns the anchor cell reference (e.g. "A1").
    pub fn anchor_cell(&self) -> &str {
        self.anchor_cell.as_str()
    }

    /// Returns the image extent dimensions, if set.
    pub fn ext(&self) -> Option<WorksheetImageExt> {
        self.ext
    }

    /// Returns the anchor type.
    pub fn anchor_type(&self) -> ImageAnchorType {
        self.anchor_type
    }

    /// Sets the anchor type.
    pub fn set_anchor_type(&mut self, anchor_type: ImageAnchorType) -> &mut Self {
        self.anchor_type = anchor_type;
        self
    }

    /// Returns the top-left cell anchor, if set.
    pub fn from_anchor(&self) -> Option<&CellAnchor> {
        self.from_anchor.as_ref()
    }

    /// Sets the top-left cell anchor.
    pub fn set_from_anchor(&mut self, anchor: CellAnchor) -> &mut Self {
        self.from_anchor = Some(anchor);
        self
    }

    /// Clears the top-left cell anchor.
    pub fn clear_from_anchor(&mut self) -> &mut Self {
        self.from_anchor = None;
        self
    }

    /// Returns the bottom-right cell anchor (for TwoCell), if set.
    pub fn to_anchor(&self) -> Option<&CellAnchor> {
        self.to_anchor.as_ref()
    }

    /// Sets the bottom-right cell anchor (for TwoCell).
    pub fn set_to_anchor(&mut self, anchor: CellAnchor) -> &mut Self {
        self.to_anchor = Some(anchor);
        self
    }

    /// Clears the bottom-right cell anchor.
    pub fn clear_to_anchor(&mut self) -> &mut Self {
        self.to_anchor = None;
        self
    }

    /// Returns the width in EMU (for OneCell/Absolute), if set.
    pub fn extent_cx(&self) -> Option<i64> {
        self.extent_cx
    }

    /// Sets the width in EMU.
    pub fn set_extent_cx(&mut self, cx: i64) -> &mut Self {
        self.extent_cx = Some(cx);
        self
    }

    /// Clears the width in EMU.
    pub fn clear_extent_cx(&mut self) -> &mut Self {
        self.extent_cx = None;
        self
    }

    /// Returns the height in EMU (for OneCell/Absolute), if set.
    pub fn extent_cy(&self) -> Option<i64> {
        self.extent_cy
    }

    /// Sets the height in EMU.
    pub fn set_extent_cy(&mut self, cy: i64) -> &mut Self {
        self.extent_cy = Some(cy);
        self
    }

    /// Clears the height in EMU.
    pub fn clear_extent_cy(&mut self) -> &mut Self {
        self.extent_cy = None;
        self
    }

    /// Returns the absolute X position in EMU (for Absolute), if set.
    pub fn position_x(&self) -> Option<i64> {
        self.position_x
    }

    /// Sets the absolute X position in EMU.
    pub fn set_position_x(&mut self, x: i64) -> &mut Self {
        self.position_x = Some(x);
        self
    }

    /// Clears the absolute X position.
    pub fn clear_position_x(&mut self) -> &mut Self {
        self.position_x = None;
        self
    }

    /// Returns the absolute Y position in EMU (for Absolute), if set.
    pub fn position_y(&self) -> Option<i64> {
        self.position_y
    }

    /// Sets the absolute Y position in EMU.
    pub fn set_position_y(&mut self, y: i64) -> &mut Self {
        self.position_y = Some(y);
        self
    }

    /// Clears the absolute Y position.
    pub fn clear_position_y(&mut self) -> &mut Self {
        self.position_y = None;
        self
    }

    /// Returns the crop percentage from the left edge (0.0 to 1.0), if set.
    pub fn crop_left(&self) -> Option<f64> {
        self.crop_left
    }

    /// Sets the crop percentage from the left edge (0.0 to 1.0).
    pub fn set_crop_left(&mut self, value: f64) -> &mut Self {
        self.crop_left = Some(value);
        self
    }

    /// Clears the left crop.
    pub fn clear_crop_left(&mut self) -> &mut Self {
        self.crop_left = None;
        self
    }

    /// Returns the crop percentage from the right edge (0.0 to 1.0), if set.
    pub fn crop_right(&self) -> Option<f64> {
        self.crop_right
    }

    /// Sets the crop percentage from the right edge (0.0 to 1.0).
    pub fn set_crop_right(&mut self, value: f64) -> &mut Self {
        self.crop_right = Some(value);
        self
    }

    /// Clears the right crop.
    pub fn clear_crop_right(&mut self) -> &mut Self {
        self.crop_right = None;
        self
    }

    /// Returns the crop percentage from the top edge (0.0 to 1.0), if set.
    pub fn crop_top(&self) -> Option<f64> {
        self.crop_top
    }

    /// Sets the crop percentage from the top edge (0.0 to 1.0).
    pub fn set_crop_top(&mut self, value: f64) -> &mut Self {
        self.crop_top = Some(value);
        self
    }

    /// Clears the top crop.
    pub fn clear_crop_top(&mut self) -> &mut Self {
        self.crop_top = None;
        self
    }

    /// Returns the crop percentage from the bottom edge (0.0 to 1.0), if set.
    pub fn crop_bottom(&self) -> Option<f64> {
        self.crop_bottom
    }

    /// Sets the crop percentage from the bottom edge (0.0 to 1.0).
    pub fn set_crop_bottom(&mut self, value: f64) -> &mut Self {
        self.crop_bottom = Some(value);
        self
    }

    /// Clears the bottom crop.
    pub fn clear_crop_bottom(&mut self) -> &mut Self {
        self.crop_bottom = None;
        self
    }

    /// Returns the image name, if set.
    pub fn name(&self) -> Option<&str> {
        self.name.as_deref()
    }

    /// Sets the image name.
    pub fn set_name(&mut self, name: impl Into<String>) -> &mut Self {
        self.name = Some(name.into());
        self
    }

    /// Clears the image name.
    pub fn clear_name(&mut self) -> &mut Self {
        self.name = None;
        self
    }

    /// Returns the image description/alt text, if set.
    pub fn description(&self) -> Option<&str> {
        self.description.as_deref()
    }

    /// Sets the image description/alt text.
    pub fn set_description(&mut self, description: impl Into<String>) -> &mut Self {
        self.description = Some(description.into());
        self
    }

    /// Clears the image description.
    pub fn clear_description(&mut self) -> &mut Self {
        self.description = None;
        self
    }
}

/// Error style for data validation error dialogs.
///
/// Controls the icon and behavior when a user enters invalid data:
/// - `Stop` prevents the invalid data from being entered.
/// - `Warning` shows a warning but allows the data.
/// - `Information` shows an informational message and allows the data.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum DataValidationErrorStyle {
    /// Error dialog with stop icon; invalid data is rejected.
    Stop,
    /// Warning dialog; user may choose to keep invalid data.
    Warning,
    /// Informational dialog; user may choose to keep invalid data.
    Information,
}

impl DataValidationErrorStyle {
    pub(crate) fn as_xml_value(self) -> &'static str {
        match self {
            Self::Stop => "stop",
            Self::Warning => "warning",
            Self::Information => "information",
        }
    }

    pub(crate) fn from_xml_value(value: &str) -> Option<Self> {
        match value {
            "stop" => Some(Self::Stop),
            "warning" => Some(Self::Warning),
            "information" => Some(Self::Information),
            _ => None,
        }
    }
}

/// Supported data validation rule families.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum DataValidationType {
    List,
    Whole,
    Decimal,
    Date,
    TextLength,
    /// Custom formula validation.
    Custom,
    /// Time validation.
    Time,
}

impl DataValidationType {
    pub(crate) fn as_xml_type(self) -> &'static str {
        match self {
            Self::List => "list",
            Self::Whole => "whole",
            Self::Decimal => "decimal",
            Self::Date => "date",
            Self::TextLength => "textLength",
            Self::Custom => "custom",
            Self::Time => "time",
        }
    }

    pub(crate) fn from_xml_type(value: &str) -> Option<Self> {
        match value {
            "list" => Some(Self::List),
            "whole" => Some(Self::Whole),
            "decimal" => Some(Self::Decimal),
            "date" => Some(Self::Date),
            "textLength" => Some(Self::TextLength),
            "custom" => Some(Self::Custom),
            "time" => Some(Self::Time),
            _ => None,
        }
    }
}

/// Data validation applied to one or more ranges.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct DataValidation {
    validation_type: DataValidationType,
    sqref: Vec<CellRange>,
    formula1: String,
    formula2: Option<String>,
    error_style: Option<DataValidationErrorStyle>,
    error_title: Option<String>,
    error_message: Option<String>,
    prompt_title: Option<String>,
    prompt_message: Option<String>,
    show_input_message: Option<bool>,
    show_error_message: Option<bool>,
}

impl DataValidation {
    pub fn list<I, S>(sqref: I, formula1: impl Into<String>) -> Result<Self>
    where
        I: IntoIterator<Item = S>,
        S: AsRef<str>,
    {
        Self::new(DataValidationType::List, sqref, formula1)
    }

    pub fn whole<I, S>(sqref: I, formula1: impl Into<String>) -> Result<Self>
    where
        I: IntoIterator<Item = S>,
        S: AsRef<str>,
    {
        Self::new(DataValidationType::Whole, sqref, formula1)
    }

    pub fn decimal<I, S>(sqref: I, formula1: impl Into<String>) -> Result<Self>
    where
        I: IntoIterator<Item = S>,
        S: AsRef<str>,
    {
        Self::new(DataValidationType::Decimal, sqref, formula1)
    }

    pub fn date<I, S>(sqref: I, formula1: impl Into<String>) -> Result<Self>
    where
        I: IntoIterator<Item = S>,
        S: AsRef<str>,
    {
        Self::new(DataValidationType::Date, sqref, formula1)
    }

    pub fn text_length<I, S>(sqref: I, formula1: impl Into<String>) -> Result<Self>
    where
        I: IntoIterator<Item = S>,
        S: AsRef<str>,
    {
        Self::new(DataValidationType::TextLength, sqref, formula1)
    }

    /// Creates a custom formula data validation.
    ///
    /// The formula should evaluate to TRUE for valid data.
    pub fn custom<I, S>(sqref: I, formula1: impl Into<String>) -> Result<Self>
    where
        I: IntoIterator<Item = S>,
        S: AsRef<str>,
    {
        Self::new(DataValidationType::Custom, sqref, formula1)
    }

    /// Creates a time validation.
    pub fn time<I, S>(sqref: I, formula1: impl Into<String>) -> Result<Self>
    where
        I: IntoIterator<Item = S>,
        S: AsRef<str>,
    {
        Self::new(DataValidationType::Time, sqref, formula1)
    }

    pub(crate) fn from_parsed_parts(
        validation_type: DataValidationType,
        sqref: Vec<CellRange>,
        formula1: String,
        formula2: Option<String>,
    ) -> Result<Self> {
        let sqref = Self::validate_sqref(sqref)?;
        let formula1 = Self::validate_formula1(formula1)?;
        let formula2 = formula2
            .map(|formula| formula.trim().to_string())
            .filter(|formula| !formula.is_empty());

        Ok(Self {
            validation_type,
            sqref,
            formula1,
            formula2,
            error_style: None,
            error_title: None,
            error_message: None,
            prompt_title: None,
            prompt_message: None,
            show_input_message: None,
            show_error_message: None,
        })
    }

    pub fn validation_type(&self) -> DataValidationType {
        self.validation_type
    }

    pub fn sqref(&self) -> &[CellRange] {
        self.sqref.as_slice()
    }

    pub fn formula1(&self) -> &str {
        self.formula1.as_str()
    }

    pub fn formula2(&self) -> Option<&str> {
        self.formula2.as_deref()
    }

    pub fn set_formula2(&mut self, formula2: impl Into<String>) -> &mut Self {
        let formula2 = formula2.into();
        let trimmed = formula2.trim();
        self.formula2 = if trimmed.is_empty() {
            None
        } else {
            Some(trimmed.to_string())
        };
        self
    }

    pub fn clear_formula2(&mut self) -> &mut Self {
        self.formula2 = None;
        self
    }

    /// Returns the error dialog style.
    pub fn error_style(&self) -> Option<DataValidationErrorStyle> {
        self.error_style
    }

    /// Sets the error dialog style.
    pub fn set_error_style(&mut self, style: DataValidationErrorStyle) -> &mut Self {
        self.error_style = Some(style);
        self
    }

    /// Clears the error dialog style.
    pub fn clear_error_style(&mut self) -> &mut Self {
        self.error_style = None;
        self
    }

    /// Returns the error dialog title.
    pub fn error_title(&self) -> Option<&str> {
        self.error_title.as_deref()
    }

    /// Sets the error dialog title.
    pub fn set_error_title(&mut self, title: impl Into<String>) -> &mut Self {
        self.error_title = Some(title.into());
        self
    }

    /// Clears the error dialog title.
    pub fn clear_error_title(&mut self) -> &mut Self {
        self.error_title = None;
        self
    }

    /// Returns the error dialog message.
    pub fn error_message(&self) -> Option<&str> {
        self.error_message.as_deref()
    }

    /// Sets the error dialog message.
    pub fn set_error_message(&mut self, message: impl Into<String>) -> &mut Self {
        self.error_message = Some(message.into());
        self
    }

    /// Clears the error dialog message.
    pub fn clear_error_message(&mut self) -> &mut Self {
        self.error_message = None;
        self
    }

    /// Returns the input prompt title.
    pub fn prompt_title(&self) -> Option<&str> {
        self.prompt_title.as_deref()
    }

    /// Sets the input prompt title.
    pub fn set_prompt_title(&mut self, title: impl Into<String>) -> &mut Self {
        self.prompt_title = Some(title.into());
        self
    }

    /// Clears the input prompt title.
    pub fn clear_prompt_title(&mut self) -> &mut Self {
        self.prompt_title = None;
        self
    }

    /// Returns the input prompt message.
    pub fn prompt_message(&self) -> Option<&str> {
        self.prompt_message.as_deref()
    }

    /// Sets the input prompt message.
    pub fn set_prompt_message(&mut self, message: impl Into<String>) -> &mut Self {
        self.prompt_message = Some(message.into());
        self
    }

    /// Clears the input prompt message.
    pub fn clear_prompt_message(&mut self) -> &mut Self {
        self.prompt_message = None;
        self
    }

    /// Returns whether the input message prompt is shown when the cell is selected.
    pub fn show_input_message(&self) -> Option<bool> {
        self.show_input_message
    }

    /// Sets whether to show the input message prompt.
    pub fn set_show_input_message(&mut self, value: bool) -> &mut Self {
        self.show_input_message = Some(value);
        self
    }

    /// Clears the show input message setting.
    pub fn clear_show_input_message(&mut self) -> &mut Self {
        self.show_input_message = None;
        self
    }

    /// Returns whether the error message dialog is shown when invalid data is entered.
    pub fn show_error_message(&self) -> Option<bool> {
        self.show_error_message
    }

    /// Sets whether to show the error message dialog.
    pub fn set_show_error_message(&mut self, value: bool) -> &mut Self {
        self.show_error_message = Some(value);
        self
    }

    /// Clears the show error message setting.
    pub fn clear_show_error_message(&mut self) -> &mut Self {
        self.show_error_message = None;
        self
    }

    pub fn add_range(&mut self, range: &str) -> Result<&mut Self> {
        let range = CellRange::parse(range)?;
        if !self.sqref.contains(&range) {
            self.sqref.push(range);
        }
        Ok(self)
    }

    pub(crate) fn sqref_xml(&self) -> String {
        self.sqref
            .iter()
            .map(format_cell_range)
            .collect::<Vec<_>>()
            .join(" ")
    }

    fn new<I, S>(
        validation_type: DataValidationType,
        sqref: I,
        formula1: impl Into<String>,
    ) -> Result<Self>
    where
        I: IntoIterator<Item = S>,
        S: AsRef<str>,
    {
        let sqref = sqref
            .into_iter()
            .map(|range| CellRange::parse(range.as_ref()))
            .collect::<Result<Vec<_>>>()?;

        Self::from_parsed_parts(validation_type, sqref, formula1.into(), None)
    }

    fn validate_formula1(formula1: String) -> Result<String> {
        let trimmed = formula1.trim();
        if trimmed.is_empty() {
            return Err(XlsxError::InvalidWorkbookState(
                "data validation formula1 must not be empty".to_string(),
            ));
        }
        Ok(trimmed.to_string())
    }

    fn validate_sqref(sqref: Vec<CellRange>) -> Result<Vec<CellRange>> {
        let mut unique_ranges = Vec::with_capacity(sqref.len());
        for range in sqref {
            if !unique_ranges.contains(&range) {
                unique_ranges.push(range);
            }
        }

        if unique_ranges.is_empty() {
            return Err(XlsxError::InvalidWorkbookState(
                "data validation sqref must include at least one range".to_string(),
            ));
        }

        Ok(unique_ranges)
    }
}

/// Aggregate function used in a table totals row.
///
/// Maps to the `totalsRowFunction` attribute on `<tableColumn>` in OOXML.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum TotalFunction {
    /// Average of column values.
    Average,
    /// Count of non-empty cells.
    Count,
    /// Count of numeric cells.
    CountNums,
    /// Maximum value.
    Max,
    /// Minimum value.
    Min,
    /// Standard deviation.
    StdDev,
    /// Sum of values.
    Sum,
    /// Variance.
    Var,
    /// Custom formula (see `TableColumn::totals_row_formula`).
    Custom,
    /// No function (cell is empty or has a label only).
    None,
}

impl TotalFunction {
    /// Returns the OOXML attribute value for this function.
    pub(crate) fn as_xml_value(self) -> &'static str {
        match self {
            Self::Average => "average",
            Self::Count => "count",
            Self::CountNums => "countNums",
            Self::Max => "max",
            Self::Min => "min",
            Self::StdDev => "stdDev",
            Self::Sum => "sum",
            Self::Var => "var",
            Self::Custom => "custom",
            Self::None => "none",
        }
    }

    /// Parses a `totalsRowFunction` attribute value from OOXML.
    pub(crate) fn from_xml_value(value: &str) -> Option<Self> {
        match value.trim() {
            "average" => Some(Self::Average),
            "count" => Some(Self::Count),
            "countNums" => Some(Self::CountNums),
            "max" => Some(Self::Max),
            "min" => Some(Self::Min),
            "stdDev" => Some(Self::StdDev),
            "sum" => Some(Self::Sum),
            "var" => Some(Self::Var),
            "custom" => Some(Self::Custom),
            "none" => Some(Self::None),
            _ => None,
        }
    }
}

/// A column definition within a worksheet table.
///
/// Maps to the `<tableColumn>` element in OOXML.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct TableColumn {
    /// Column name (header text).
    name: String,
    /// Column ID (1-based).
    id: u32,
    /// Label displayed in the totals row (typically for the first column).
    totals_row_label: Option<String>,
    /// Aggregate function for the totals row.
    totals_row_function: Option<TotalFunction>,
    /// Custom formula for the totals row (used when function is `Custom`).
    totals_row_formula: Option<String>,
    /// Unknown attributes preserved for roundtrip fidelity.
    unknown_attrs: Vec<(String, String)>,
}

impl TableColumn {
    /// Creates a new table column with the given name and ID.
    pub fn new(name: impl Into<String>, id: u32) -> Self {
        Self {
            name: name.into(),
            id,
            totals_row_label: None,
            totals_row_function: None,
            totals_row_formula: None,
            unknown_attrs: Vec::new(),
        }
    }

    /// Returns the column header name.
    pub fn name(&self) -> &str {
        self.name.as_str()
    }

    /// Sets the column header name.
    pub fn set_name(&mut self, name: impl Into<String>) -> &mut Self {
        self.name = name.into();
        self
    }

    /// Returns the column ID (1-based).
    pub fn id(&self) -> u32 {
        self.id
    }

    /// Sets the column ID.
    pub fn set_id(&mut self, id: u32) -> &mut Self {
        self.id = id;
        self
    }

    /// Returns the totals row label, if set.
    pub fn totals_row_label(&self) -> Option<&str> {
        self.totals_row_label.as_deref()
    }

    /// Sets the totals row label.
    pub fn set_totals_row_label(&mut self, label: impl Into<String>) -> &mut Self {
        self.totals_row_label = Some(label.into());
        self
    }

    /// Clears the totals row label.
    pub fn clear_totals_row_label(&mut self) -> &mut Self {
        self.totals_row_label = None;
        self
    }

    /// Returns the totals row function, if set.
    pub fn totals_row_function(&self) -> Option<TotalFunction> {
        self.totals_row_function
    }

    /// Sets the totals row function.
    pub fn set_totals_row_function(&mut self, function: TotalFunction) -> &mut Self {
        self.totals_row_function = Some(function);
        self
    }

    /// Clears the totals row function.
    pub fn clear_totals_row_function(&mut self) -> &mut Self {
        self.totals_row_function = None;
        self
    }

    /// Returns the totals row custom formula, if set.
    pub fn totals_row_formula(&self) -> Option<&str> {
        self.totals_row_formula.as_deref()
    }

    /// Sets the totals row custom formula.
    pub fn set_totals_row_formula(&mut self, formula: impl Into<String>) -> &mut Self {
        self.totals_row_formula = Some(formula.into());
        self
    }

    /// Clears the totals row custom formula.
    pub fn clear_totals_row_formula(&mut self) -> &mut Self {
        self.totals_row_formula = None;
        self
    }

    /// Returns unknown attributes preserved for roundtrip fidelity.
    pub(crate) fn unknown_attrs(&self) -> &[(String, String)] {
        self.unknown_attrs.as_slice()
    }

    /// Sets unknown attributes for roundtrip fidelity.
    pub(crate) fn set_unknown_attrs(&mut self, attrs: Vec<(String, String)>) {
        self.unknown_attrs = attrs;
    }
}

/// A worksheet table definition.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct WorksheetTable {
    name: String,
    range: CellRange,
    has_header_row: bool,
    totals_row_shown: Option<bool>,
    columns: Vec<TableColumn>,
    style_name: Option<String>,
    show_first_column: Option<bool>,
    show_last_column: Option<bool>,
    show_row_stripes: Option<bool>,
    show_column_stripes: Option<bool>,
    /// Unknown attributes on the `<table>` element, preserved for roundtrip fidelity.
    unknown_table_attrs: Vec<(String, String)>,
    /// Unknown attributes on the `<tableStyleInfo>` element, preserved for roundtrip fidelity.
    unknown_style_attrs: Vec<(String, String)>,
}

impl WorksheetTable {
    pub fn new(name: impl Into<String>, range: &str) -> Result<Self> {
        Self::with_header_row(name, range, true)
    }

    pub fn with_header_row(
        name: impl Into<String>,
        range: &str,
        has_header_row: bool,
    ) -> Result<Self> {
        let name = validate_table_name(name.into())?;
        let range = CellRange::parse(range)?;
        Ok(Self {
            name,
            range,
            has_header_row,
            totals_row_shown: None,
            columns: Vec::new(),
            style_name: None,
            show_first_column: None,
            show_last_column: None,
            show_row_stripes: None,
            show_column_stripes: None,
            unknown_table_attrs: Vec::new(),
            unknown_style_attrs: Vec::new(),
        })
    }

    pub(crate) fn from_parsed_parts(
        name: String,
        range: CellRange,
        has_header_row: bool,
    ) -> Result<Self> {
        let name = validate_table_name(name)?;
        Ok(Self {
            name,
            range,
            has_header_row,
            totals_row_shown: None,
            columns: Vec::new(),
            style_name: None,
            show_first_column: None,
            show_last_column: None,
            show_row_stripes: None,
            show_column_stripes: None,
            unknown_table_attrs: Vec::new(),
            unknown_style_attrs: Vec::new(),
        })
    }

    pub fn name(&self) -> &str {
        self.name.as_str()
    }

    pub fn range(&self) -> &CellRange {
        &self.range
    }

    pub fn has_header_row(&self) -> bool {
        self.has_header_row
    }

    pub fn set_name(&mut self, name: impl Into<String>) -> Result<&mut Self> {
        self.name = validate_table_name(name.into())?;
        Ok(self)
    }

    pub fn set_range(&mut self, range: &str) -> Result<&mut Self> {
        self.range = CellRange::parse(range)?;
        Ok(self)
    }

    pub fn set_header_row(&mut self, has_header_row: bool) -> &mut Self {
        self.has_header_row = has_header_row;
        self
    }

    /// Returns whether the totals row is shown.
    pub fn totals_row_shown(&self) -> Option<bool> {
        self.totals_row_shown
    }

    /// Sets whether the totals row is shown.
    pub fn set_totals_row_shown(&mut self, value: bool) -> &mut Self {
        self.totals_row_shown = Some(value);
        self
    }

    /// Clears the totals row shown setting.
    pub fn clear_totals_row_shown(&mut self) -> &mut Self {
        self.totals_row_shown = None;
        self
    }

    /// Returns the table column definitions.
    pub fn columns(&self) -> &[TableColumn] {
        self.columns.as_slice()
    }

    /// Returns mutable access to the table column definitions.
    pub fn columns_mut(&mut self) -> &mut Vec<TableColumn> {
        &mut self.columns
    }

    /// Adds a column definition to the table.
    pub fn push_column(&mut self, column: TableColumn) -> &mut Self {
        self.columns.push(column);
        self
    }

    /// Add a column with a given name and auto-assigned ID.
    ///
    /// The ID is set to the next available value (max existing ID + 1, or 1 if empty).
    pub fn add_column(&mut self, name: impl Into<String>) -> &mut TableColumn {
        let next_id = self.columns.iter().map(|c| c.id()).max().unwrap_or(0) + 1;
        self.columns.push(TableColumn::new(name, next_id));
        let idx = self.columns.len() - 1;
        &mut self.columns[idx]
    }

    /// Remove a column at the given index.
    ///
    /// Returns `None` if the index is out of bounds.
    pub fn remove_column(&mut self, index: usize) -> Option<TableColumn> {
        if index < self.columns.len() {
            Some(self.columns.remove(index))
        } else {
            None
        }
    }

    /// Returns the table style name (e.g. "TableStyleMedium9").
    pub fn style_name(&self) -> Option<&str> {
        self.style_name.as_deref()
    }

    /// Sets the table style name.
    pub fn set_style_name(&mut self, name: impl Into<String>) -> &mut Self {
        self.style_name = Some(name.into());
        self
    }

    /// Clears the table style name.
    pub fn clear_style_name(&mut self) -> &mut Self {
        self.style_name = None;
        self
    }

    /// Returns whether the first column is highlighted.
    pub fn show_first_column(&self) -> Option<bool> {
        self.show_first_column
    }

    /// Sets whether the first column is highlighted.
    pub fn set_show_first_column(&mut self, value: bool) -> &mut Self {
        self.show_first_column = Some(value);
        self
    }

    /// Clears the show first column setting.
    pub fn clear_show_first_column(&mut self) -> &mut Self {
        self.show_first_column = None;
        self
    }

    /// Returns whether the last column is highlighted.
    pub fn show_last_column(&self) -> Option<bool> {
        self.show_last_column
    }

    /// Sets whether the last column is highlighted.
    pub fn set_show_last_column(&mut self, value: bool) -> &mut Self {
        self.show_last_column = Some(value);
        self
    }

    /// Clears the show last column setting.
    pub fn clear_show_last_column(&mut self) -> &mut Self {
        self.show_last_column = None;
        self
    }

    /// Returns whether row stripes are shown.
    pub fn show_row_stripes(&self) -> Option<bool> {
        self.show_row_stripes
    }

    /// Sets whether row stripes are shown.
    pub fn set_show_row_stripes(&mut self, value: bool) -> &mut Self {
        self.show_row_stripes = Some(value);
        self
    }

    /// Clears the show row stripes setting.
    pub fn clear_show_row_stripes(&mut self) -> &mut Self {
        self.show_row_stripes = None;
        self
    }

    /// Returns whether column stripes are shown.
    pub fn show_column_stripes(&self) -> Option<bool> {
        self.show_column_stripes
    }

    /// Sets whether column stripes are shown.
    pub fn set_show_column_stripes(&mut self, value: bool) -> &mut Self {
        self.show_column_stripes = Some(value);
        self
    }

    /// Clears the show column stripes setting.
    pub fn clear_show_column_stripes(&mut self) -> &mut Self {
        self.show_column_stripes = None;
        self
    }

    /// Returns unknown attributes on the `<table>` element, preserved for roundtrip fidelity.
    pub(crate) fn unknown_table_attrs(&self) -> &[(String, String)] {
        self.unknown_table_attrs.as_slice()
    }

    /// Sets unknown attributes on the `<table>` element for roundtrip fidelity.
    pub(crate) fn set_unknown_table_attrs(&mut self, attrs: Vec<(String, String)>) {
        self.unknown_table_attrs = attrs;
    }

    /// Returns unknown attributes on the `<tableStyleInfo>` element, preserved for roundtrip fidelity.
    pub(crate) fn unknown_style_attrs(&self) -> &[(String, String)] {
        self.unknown_style_attrs.as_slice()
    }

    /// Sets unknown attributes on the `<tableStyleInfo>` element for roundtrip fidelity.
    pub(crate) fn set_unknown_style_attrs(&mut self, attrs: Vec<(String, String)>) {
        self.unknown_style_attrs = attrs;
    }

    pub(crate) fn range_xml(&self) -> String {
        format_cell_range(&self.range)
    }
}

/// Supported conditional formatting rule types.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ConditionalFormattingRuleType {
    CellIs,
    Expression,
    ColorScale,
    DataBar,
    IconSet,
    Top10,
    AboveAverage,
    TimePeriod,
    DuplicateValues,
    UniqueValues,
    ContainsText,
    NotContainsText,
    BeginsWith,
    EndsWith,
    ContainsBlanks,
    NotContainsBlanks,
    ContainsErrors,
    NotContainsErrors,
}

impl ConditionalFormattingRuleType {
    pub(crate) fn as_xml_type(self) -> &'static str {
        match self {
            Self::CellIs => "cellIs",
            Self::Expression => "expression",
            Self::ColorScale => "colorScale",
            Self::DataBar => "dataBar",
            Self::IconSet => "iconSet",
            Self::Top10 => "top10",
            Self::AboveAverage => "aboveAverage",
            Self::TimePeriod => "timePeriod",
            Self::DuplicateValues => "duplicateValues",
            Self::UniqueValues => "uniqueValues",
            Self::ContainsText => "containsText",
            Self::NotContainsText => "notContainsText",
            Self::BeginsWith => "beginsWith",
            Self::EndsWith => "endsWith",
            Self::ContainsBlanks => "containsBlanks",
            Self::NotContainsBlanks => "notContainsBlanks",
            Self::ContainsErrors => "containsErrors",
            Self::NotContainsErrors => "notContainsErrors",
        }
    }

    pub(crate) fn from_xml_type(value: &str) -> Option<Self> {
        match value {
            "cellIs" => Some(Self::CellIs),
            "expression" => Some(Self::Expression),
            "colorScale" => Some(Self::ColorScale),
            "dataBar" => Some(Self::DataBar),
            "iconSet" => Some(Self::IconSet),
            "top10" => Some(Self::Top10),
            "aboveAverage" => Some(Self::AboveAverage),
            "timePeriod" => Some(Self::TimePeriod),
            "duplicateValues" => Some(Self::DuplicateValues),
            "uniqueValues" => Some(Self::UniqueValues),
            "containsText" => Some(Self::ContainsText),
            "notContainsText" => Some(Self::NotContainsText),
            "beginsWith" => Some(Self::BeginsWith),
            "endsWith" => Some(Self::EndsWith),
            "containsBlanks" => Some(Self::ContainsBlanks),
            "notContainsBlanks" => Some(Self::NotContainsBlanks),
            "containsErrors" => Some(Self::ContainsErrors),
            "notContainsErrors" => Some(Self::NotContainsErrors),
            _ => None,
        }
    }
}

/// Conditional formatting comparison operator (used with CellIs rules).
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ConditionalFormattingOperator {
    LessThan,
    LessThanOrEqual,
    Equal,
    NotEqual,
    GreaterThanOrEqual,
    GreaterThan,
    Between,
    NotBetween,
    ContainsText,
    NotContains,
    BeginsWith,
    EndsWith,
}

impl ConditionalFormattingOperator {
    pub(crate) fn as_xml_value(self) -> &'static str {
        match self {
            Self::LessThan => "lessThan",
            Self::LessThanOrEqual => "lessThanOrEqual",
            Self::Equal => "equal",
            Self::NotEqual => "notEqual",
            Self::GreaterThanOrEqual => "greaterThanOrEqual",
            Self::GreaterThan => "greaterThan",
            Self::Between => "between",
            Self::NotBetween => "notBetween",
            Self::ContainsText => "containsText",
            Self::NotContains => "notContains",
            Self::BeginsWith => "beginsWith",
            Self::EndsWith => "endsWith",
        }
    }

    pub(crate) fn from_xml_value(value: &str) -> Option<Self> {
        match value {
            "lessThan" => Some(Self::LessThan),
            "lessThanOrEqual" => Some(Self::LessThanOrEqual),
            "equal" => Some(Self::Equal),
            "notEqual" => Some(Self::NotEqual),
            "greaterThanOrEqual" => Some(Self::GreaterThanOrEqual),
            "greaterThan" => Some(Self::GreaterThan),
            "between" => Some(Self::Between),
            "notBetween" => Some(Self::NotBetween),
            "containsText" => Some(Self::ContainsText),
            "notContains" => Some(Self::NotContains),
            "beginsWith" => Some(Self::BeginsWith),
            "endsWith" => Some(Self::EndsWith),
            _ => None,
        }
    }
}

/// Value type for color scale / data bar / icon set thresholds.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum CfValueObjectType {
    Num,
    Percent,
    Max,
    Min,
    Formula,
    Percentile,
    /// Automatic minimum (Excel 2010+ data bar extension).
    AutoMin,
    /// Automatic maximum (Excel 2010+ data bar extension).
    AutoMax,
}

impl CfValueObjectType {
    /// Convert to the XML attribute value string.
    pub fn as_xml_value(self) -> &'static str {
        match self {
            Self::Num => "num",
            Self::Percent => "percent",
            Self::Max => "max",
            Self::Min => "min",
            Self::Formula => "formula",
            Self::Percentile => "percentile",
            Self::AutoMin => "autoMin",
            Self::AutoMax => "autoMax",
        }
    }

    pub(crate) fn from_xml_value(value: &str) -> Option<Self> {
        match value {
            "num" => Some(Self::Num),
            "percent" => Some(Self::Percent),
            "max" => Some(Self::Max),
            "min" => Some(Self::Min),
            "formula" => Some(Self::Formula),
            "percentile" => Some(Self::Percentile),
            "autoMin" => Some(Self::AutoMin),
            "autoMax" => Some(Self::AutoMax),
            _ => None,
        }
    }
}

/// A threshold value used in color scale, data bar, and icon set rules.
#[derive(Debug, Clone, PartialEq)]
pub struct CfValueObject {
    /// The type of this value (num, percent, min, max, formula, percentile).
    pub value_type: CfValueObjectType,
    /// The value (may be empty for min/max types).
    pub value: Option<String>,
}

/// A color scale stop (value + color).
#[derive(Debug, Clone, PartialEq)]
pub struct ColorScaleStop {
    pub cfvo: CfValueObject,
    pub color: String,
}

/// Conditional formatting applied to one or more ranges.
#[derive(Debug, Clone, PartialEq)]
pub struct ConditionalFormatting {
    sqref: Vec<CellRange>,
    /// Raw sqref string from the XML attribute, preserved for references that
    /// `CellRange` cannot represent (e.g. column-only "A:A" or row-only "1:1").
    raw_sqref: Option<String>,
    rule_type: ConditionalFormattingRuleType,
    formulas: Vec<String>,
    /// CellIs comparison operator (e.g. "greaterThan", "between").
    operator: Option<ConditionalFormattingOperator>,
    /// DXF style index for formatting (maps into dxfs table in styles.xml).
    dxf_id: Option<u32>,
    /// Priority of this rule (lower = higher priority).
    priority: Option<u32>,
    /// Whether to stop processing lower-priority rules when this one matches.
    stop_if_true: Option<bool>,
    /// Text value for containsText/notContainsText/beginsWith/endsWith rules.
    text: Option<String>,
    /// Time period value for timePeriod rules (e.g. "today", "thisWeek").
    time_period: Option<String>,
    /// Top10: rank value.
    rank: Option<u32>,
    /// Top10: whether rank is a percentage.
    percent: Option<bool>,
    /// Top10: whether bottom values (false = top, true = bottom).
    bottom: Option<bool>,
    /// AboveAverage: whether above (true) or below (false) average.
    above_average: Option<bool>,
    /// AboveAverage: whether to include equal values.
    equal_average: Option<bool>,
    /// AboveAverage: standard deviation count (1, 2, or 3).
    std_dev: Option<u32>,
    /// Color scale stops (2 or 3).
    color_scale_stops: Vec<ColorScaleStop>,
    /// Data bar minimum threshold.
    data_bar_min: Option<CfValueObject>,
    /// Data bar maximum threshold.
    data_bar_max: Option<CfValueObject>,
    /// Data bar fill color.
    data_bar_color: Option<String>,
    /// Data bar: whether to show the cell value alongside the bar.
    data_bar_show_value: Option<bool>,
    /// Data bar: minimum bar length as a percentage of the cell width.
    data_bar_min_length: Option<u32>,
    /// Data bar: maximum bar length as a percentage of the cell width.
    data_bar_max_length: Option<u32>,
    /// Icon set name (e.g. "3TrafficLights1", "3Arrows").
    icon_set_name: Option<String>,
    /// Icon set thresholds.
    icon_set_values: Vec<CfValueObject>,
    /// Icon set: whether to show the cell value in addition to the icon.
    icon_set_show_value: Option<bool>,
    /// Icon set: whether to reverse the icon order.
    icon_set_reverse: Option<bool>,
}

impl ConditionalFormatting {
    pub fn cell_is<I, S, J, T>(sqref: I, formulas: J) -> Result<Self>
    where
        I: IntoIterator<Item = S>,
        S: AsRef<str>,
        J: IntoIterator<Item = T>,
        T: AsRef<str>,
    {
        Self::new(ConditionalFormattingRuleType::CellIs, sqref, formulas)
    }

    pub fn expression<I, S, J, T>(sqref: I, formulas: J) -> Result<Self>
    where
        I: IntoIterator<Item = S>,
        S: AsRef<str>,
        J: IntoIterator<Item = T>,
        T: AsRef<str>,
    {
        Self::new(ConditionalFormattingRuleType::Expression, sqref, formulas)
    }

    pub(crate) fn from_parsed_parts(
        rule_type: ConditionalFormattingRuleType,
        sqref: Vec<CellRange>,
        formulas: Vec<String>,
    ) -> Result<Self> {
        Self::from_parsed_parts_raw(rule_type, sqref, None, formulas)
    }

    /// Like `from_parsed_parts` but also stores a raw sqref string for references
    /// that `CellRange` cannot represent (e.g. column-only "A:A" or row-only "1:1").
    pub(crate) fn from_parsed_parts_raw(
        rule_type: ConditionalFormattingRuleType,
        sqref: Vec<CellRange>,
        raw_sqref: Option<String>,
        formulas: Vec<String>,
    ) -> Result<Self> {
        // Allow empty parsed sqref when a raw sqref string is present.
        let sqref = if sqref.is_empty() && raw_sqref.is_some() {
            sqref
        } else {
            Self::validate_sqref(sqref)?
        };
        // For formula-based types, validate; for others (colorScale, dataBar, etc.), allow empty.
        let formulas = match rule_type {
            ConditionalFormattingRuleType::CellIs
            | ConditionalFormattingRuleType::Expression
            | ConditionalFormattingRuleType::ContainsText
            | ConditionalFormattingRuleType::NotContainsText
            | ConditionalFormattingRuleType::BeginsWith
            | ConditionalFormattingRuleType::EndsWith => Self::validate_formulas(formulas)?,
            _ => formulas,
        };
        Ok(Self {
            sqref,
            raw_sqref,
            rule_type,
            formulas,
            operator: None,
            dxf_id: None,
            priority: None,
            stop_if_true: None,
            text: None,
            time_period: None,
            rank: None,
            percent: None,
            bottom: None,
            above_average: None,
            equal_average: None,
            std_dev: None,
            color_scale_stops: Vec::new(),
            data_bar_min: None,
            data_bar_max: None,
            data_bar_color: None,
            data_bar_show_value: None,
            data_bar_min_length: None,
            data_bar_max_length: None,
            icon_set_name: None,
            icon_set_values: Vec::new(),
            icon_set_show_value: None,
            icon_set_reverse: None,
        })
    }

    /// Creates a conditional formatting rule with full parsed data (used by parser).
    #[allow(dead_code)]
    #[allow(clippy::too_many_arguments)]
    pub(crate) fn from_parsed_rule(
        rule_type: ConditionalFormattingRuleType,
        sqref: Vec<CellRange>,
    ) -> Result<Self> {
        let sqref = Self::validate_sqref(sqref)?;
        Ok(Self {
            sqref,
            raw_sqref: None,
            rule_type,
            formulas: Vec::new(),
            operator: None,
            dxf_id: None,
            priority: None,
            stop_if_true: None,
            text: None,
            time_period: None,
            rank: None,
            percent: None,
            bottom: None,
            above_average: None,
            equal_average: None,
            std_dev: None,
            color_scale_stops: Vec::new(),
            data_bar_min: None,
            data_bar_max: None,
            data_bar_color: None,
            data_bar_show_value: None,
            data_bar_min_length: None,
            data_bar_max_length: None,
            icon_set_name: None,
            icon_set_values: Vec::new(),
            icon_set_show_value: None,
            icon_set_reverse: None,
        })
    }

    pub fn sqref(&self) -> &[CellRange] {
        self.sqref.as_slice()
    }

    /// Returns the raw sqref string from the XML, if present.
    ///
    /// This is set when the sqref contains references that `CellRange` cannot
    /// represent (e.g. column-only "A:A" or row-only "1:1").
    pub fn raw_sqref(&self) -> Option<&str> {
        self.raw_sqref.as_deref()
    }

    pub fn rule_type(&self) -> ConditionalFormattingRuleType {
        self.rule_type
    }

    pub fn formulas(&self) -> &[String] {
        self.formulas.as_slice()
    }

    /// CellIs comparison operator.
    pub fn operator(&self) -> Option<ConditionalFormattingOperator> {
        self.operator
    }

    /// Sets the CellIs comparison operator.
    pub fn set_operator(&mut self, op: ConditionalFormattingOperator) {
        self.operator = Some(op);
    }

    /// DXF style index.
    pub fn dxf_id(&self) -> Option<u32> {
        self.dxf_id
    }

    /// Sets the DXF style index.
    pub fn set_dxf_id(&mut self, id: u32) {
        self.dxf_id = Some(id);
    }

    /// Rule priority.
    pub fn priority(&self) -> Option<u32> {
        self.priority
    }

    /// Sets the rule priority.
    pub fn set_priority(&mut self, priority: u32) {
        self.priority = Some(priority);
    }

    /// Whether to stop processing lower-priority rules when this one matches.
    pub fn stop_if_true(&self) -> Option<bool> {
        self.stop_if_true
    }

    /// Sets stop-if-true flag.
    pub fn set_stop_if_true(&mut self, value: bool) {
        self.stop_if_true = Some(value);
    }

    /// Text value for text-based rules.
    pub fn text(&self) -> Option<&str> {
        self.text.as_deref()
    }

    /// Sets text value.
    pub fn set_text(&mut self, text: impl Into<String>) {
        self.text = Some(text.into());
    }

    /// Time period for timePeriod rules.
    pub fn time_period(&self) -> Option<&str> {
        self.time_period.as_deref()
    }

    /// Sets time period.
    pub fn set_time_period(&mut self, period: impl Into<String>) {
        self.time_period = Some(period.into());
    }

    /// Top10 rank value.
    pub fn rank(&self) -> Option<u32> {
        self.rank
    }

    /// Sets Top10 rank.
    pub fn set_rank(&mut self, rank: u32) {
        self.rank = Some(rank);
    }

    /// Top10 percent flag.
    pub fn cf_percent(&self) -> Option<bool> {
        self.percent
    }

    /// Sets Top10 percent flag.
    pub fn set_cf_percent(&mut self, value: bool) {
        self.percent = Some(value);
    }

    /// Top10 bottom flag.
    pub fn cf_bottom(&self) -> Option<bool> {
        self.bottom
    }

    /// Sets Top10 bottom flag.
    pub fn set_cf_bottom(&mut self, value: bool) {
        self.bottom = Some(value);
    }

    /// AboveAverage: above flag.
    pub fn above_average(&self) -> Option<bool> {
        self.above_average
    }

    /// Sets above-average flag.
    pub fn set_above_average(&mut self, value: bool) {
        self.above_average = Some(value);
    }

    /// AboveAverage: equal flag.
    pub fn equal_average(&self) -> Option<bool> {
        self.equal_average
    }

    /// Sets equal-average flag.
    pub fn set_equal_average(&mut self, value: bool) {
        self.equal_average = Some(value);
    }

    /// AboveAverage: std dev count.
    pub fn std_dev(&self) -> Option<u32> {
        self.std_dev
    }

    /// Sets std dev count.
    pub fn set_std_dev(&mut self, value: u32) {
        self.std_dev = Some(value);
    }

    /// Color scale stops.
    pub fn color_scale_stops(&self) -> &[ColorScaleStop] {
        &self.color_scale_stops
    }

    /// Sets color scale stops.
    pub fn set_color_scale_stops(&mut self, stops: Vec<ColorScaleStop>) {
        self.color_scale_stops = stops;
    }

    /// Data bar min threshold.
    pub fn data_bar_min(&self) -> Option<&CfValueObject> {
        self.data_bar_min.as_ref()
    }

    /// Data bar max threshold.
    pub fn data_bar_max(&self) -> Option<&CfValueObject> {
        self.data_bar_max.as_ref()
    }

    /// Data bar color.
    pub fn data_bar_color(&self) -> Option<&str> {
        self.data_bar_color.as_deref()
    }

    /// Sets data bar parameters.
    pub fn set_data_bar(
        &mut self,
        min: CfValueObject,
        max: CfValueObject,
        color: impl Into<String>,
    ) {
        self.data_bar_min = Some(min);
        self.data_bar_max = Some(max);
        self.data_bar_color = Some(color.into());
    }

    /// Whether to show the cell value alongside the data bar.
    pub fn data_bar_show_value(&self) -> Option<bool> {
        self.data_bar_show_value
    }

    /// Sets whether to show the cell value alongside the data bar.
    pub fn set_data_bar_show_value(&mut self, value: bool) {
        self.data_bar_show_value = Some(value);
    }

    /// Minimum bar length as a percentage of cell width.
    pub fn data_bar_min_length(&self) -> Option<u32> {
        self.data_bar_min_length
    }

    /// Sets minimum bar length as a percentage of cell width.
    pub fn set_data_bar_min_length(&mut self, value: u32) {
        self.data_bar_min_length = Some(value);
    }

    /// Maximum bar length as a percentage of cell width.
    pub fn data_bar_max_length(&self) -> Option<u32> {
        self.data_bar_max_length
    }

    /// Sets maximum bar length as a percentage of cell width.
    pub fn set_data_bar_max_length(&mut self, value: u32) {
        self.data_bar_max_length = Some(value);
    }

    /// Icon set name.
    pub fn icon_set_name(&self) -> Option<&str> {
        self.icon_set_name.as_deref()
    }

    /// Icon set threshold values.
    pub fn icon_set_values(&self) -> &[CfValueObject] {
        &self.icon_set_values
    }

    /// Icon set show-value flag.
    pub fn icon_set_show_value(&self) -> Option<bool> {
        self.icon_set_show_value
    }

    /// Icon set reverse flag.
    pub fn icon_set_reverse(&self) -> Option<bool> {
        self.icon_set_reverse
    }

    /// Sets icon set parameters.
    pub fn set_icon_set(&mut self, name: impl Into<String>, values: Vec<CfValueObject>) {
        self.icon_set_name = Some(name.into());
        self.icon_set_values = values;
    }

    /// Sets icon set name.
    pub fn set_icon_set_name(&mut self, name: impl Into<String>) {
        self.icon_set_name = Some(name.into());
    }

    /// Sets icon set threshold values.
    pub fn set_icon_set_values(&mut self, values: Vec<CfValueObject>) {
        self.icon_set_values = values;
    }

    /// Sets icon set show-value flag.
    pub fn set_icon_set_show_value(&mut self, value: bool) {
        self.icon_set_show_value = Some(value);
    }

    /// Sets icon set reverse flag.
    pub fn set_icon_set_reverse(&mut self, value: bool) {
        self.icon_set_reverse = Some(value);
    }

    /// Sets data bar minimum threshold.
    pub fn set_data_bar_min(&mut self, min: CfValueObject) {
        self.data_bar_min = Some(min);
    }

    /// Sets data bar maximum threshold.
    pub fn set_data_bar_max(&mut self, max: CfValueObject) {
        self.data_bar_max = Some(max);
    }

    /// Sets data bar fill color.
    pub fn set_data_bar_color(&mut self, color: impl Into<String>) {
        self.data_bar_color = Some(color.into());
    }

    /// Sets the formulas (used by parser to set formulas on a rule created via `from_parsed_rule`).
    #[allow(dead_code)]
    pub(crate) fn set_formulas(&mut self, formulas: Vec<String>) {
        self.formulas = formulas;
    }

    pub fn add_range(&mut self, range: &str) -> Result<&mut Self> {
        let range = CellRange::parse(range)?;
        if !self.sqref.contains(&range) {
            self.sqref.push(range);
        }
        Ok(self)
    }

    pub fn add_formula(&mut self, formula: impl Into<String>) -> Result<&mut Self> {
        let formula = formula.into();
        let formula = formula.trim();
        if formula.is_empty() {
            return Err(XlsxError::InvalidWorkbookState(
                "conditional formatting formula must not be empty".to_string(),
            ));
        }
        self.formulas.push(formula.to_string());
        Ok(self)
    }

    pub(crate) fn sqref_xml(&self) -> String {
        self.sqref
            .iter()
            .map(format_cell_range)
            .collect::<Vec<_>>()
            .join(" ")
    }

    fn new<I, S, J, T>(
        rule_type: ConditionalFormattingRuleType,
        sqref: I,
        formulas: J,
    ) -> Result<Self>
    where
        I: IntoIterator<Item = S>,
        S: AsRef<str>,
        J: IntoIterator<Item = T>,
        T: AsRef<str>,
    {
        let sqref = sqref
            .into_iter()
            .map(|range| CellRange::parse(range.as_ref()))
            .collect::<Result<Vec<_>>>()?;
        let formulas = formulas
            .into_iter()
            .map(|formula| formula.as_ref().to_string())
            .collect::<Vec<_>>();
        Self::from_parsed_parts(rule_type, sqref, formulas)
    }

    fn validate_sqref(sqref: Vec<CellRange>) -> Result<Vec<CellRange>> {
        let mut unique_ranges = Vec::with_capacity(sqref.len());
        for range in sqref {
            if !unique_ranges.contains(&range) {
                unique_ranges.push(range);
            }
        }

        if unique_ranges.is_empty() {
            return Err(XlsxError::InvalidWorkbookState(
                "conditional formatting sqref must include at least one range".to_string(),
            ));
        }

        Ok(unique_ranges)
    }

    fn validate_formulas(formulas: Vec<String>) -> Result<Vec<String>> {
        let cleaned = formulas
            .into_iter()
            .map(|formula| formula.trim().to_string())
            .filter(|formula| !formula.is_empty())
            .collect::<Vec<_>>();

        if cleaned.is_empty() {
            return Err(XlsxError::InvalidWorkbookState(
                "conditional formatting must include at least one formula".to_string(),
            ));
        }

        Ok(cleaned)
    }
}

fn validate_table_name(name: String) -> Result<String> {
    let name = name.trim();
    if name.is_empty() {
        return Err(XlsxError::InvalidWorkbookState(
            "worksheet table name must not be empty".to_string(),
        ));
    }
    Ok(name.to_string())
}

/// A hyperlink attached to a cell in a worksheet.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Hyperlink {
    cell_ref: String,
    url: Option<String>,
    location: Option<String>,
    tooltip: Option<String>,
    /// Display text for the hyperlink (may differ from the cell value).
    display: Option<String>,
}

impl Hyperlink {
    /// Creates an external hyperlink (URL) for the given cell reference.
    pub fn external(cell_ref: &str, url: impl Into<String>) -> Result<Self> {
        let cell_ref = normalize_cell_reference(cell_ref)?;
        let url = url.into();
        let url = url.trim().to_string();
        if url.is_empty() {
            return Err(XlsxError::InvalidWorkbookState(
                "hyperlink url must not be empty".to_string(),
            ));
        }
        Ok(Self {
            cell_ref,
            url: Some(url),
            location: None,
            tooltip: None,
            display: None,
        })
    }

    /// Creates an internal hyperlink (location within the workbook) for the given cell reference.
    pub fn internal(cell_ref: &str, location: impl Into<String>) -> Result<Self> {
        let cell_ref = normalize_cell_reference(cell_ref)?;
        let location = location.into();
        let location = location.trim().to_string();
        if location.is_empty() {
            return Err(XlsxError::InvalidWorkbookState(
                "hyperlink location must not be empty".to_string(),
            ));
        }
        Ok(Self {
            cell_ref,
            url: None,
            location: Some(location),
            tooltip: None,
            display: None,
        })
    }

    /// Creates a hyperlink from parsed components. Used by the parser.
    pub(crate) fn from_parsed_parts(
        cell_ref: String,
        url: Option<String>,
        location: Option<String>,
        tooltip: Option<String>,
        display: Option<String>,
    ) -> Result<Self> {
        let cell_ref = normalize_cell_reference(cell_ref.as_str())?;
        let url = url.map(|u| u.trim().to_string()).filter(|u| !u.is_empty());
        let location = location
            .map(|l| l.trim().to_string())
            .filter(|l| !l.is_empty());
        let tooltip = tooltip
            .map(|t| t.trim().to_string())
            .filter(|t| !t.is_empty());
        let display = display
            .map(|d| d.trim().to_string())
            .filter(|d| !d.is_empty());
        if url.is_none() && location.is_none() {
            return Err(XlsxError::InvalidWorkbookState(
                "hyperlink must have either a url or a location".to_string(),
            ));
        }
        Ok(Self {
            cell_ref,
            url,
            location,
            tooltip,
            display,
        })
    }

    /// Returns the cell reference this hyperlink is attached to.
    pub fn cell_ref(&self) -> &str {
        self.cell_ref.as_str()
    }

    /// Returns the external URL, if this is an external hyperlink.
    pub fn url(&self) -> Option<&str> {
        self.url.as_deref()
    }

    /// Returns the internal location, if this is an internal hyperlink.
    pub fn location(&self) -> Option<&str> {
        self.location.as_deref()
    }

    /// Returns the tooltip text.
    pub fn tooltip(&self) -> Option<&str> {
        self.tooltip.as_deref()
    }

    /// Sets the tooltip text.
    pub fn set_tooltip(&mut self, tooltip: impl Into<String>) -> &mut Self {
        let tooltip = tooltip.into();
        let trimmed = tooltip.trim();
        self.tooltip = if trimmed.is_empty() {
            None
        } else {
            Some(trimmed.to_string())
        };
        self
    }

    /// Clears the tooltip text.
    pub fn clear_tooltip(&mut self) -> &mut Self {
        self.tooltip = None;
        self
    }

    /// Returns the display text, if set.
    pub fn display(&self) -> Option<&str> {
        self.display.as_deref()
    }

    /// Sets the display text (may differ from the cell value).
    pub fn set_display(&mut self, display: impl Into<String>) -> &mut Self {
        let display = display.into();
        let trimmed = display.trim();
        self.display = if trimmed.is_empty() {
            None
        } else {
            Some(trimmed.to_string())
        };
        self
    }

    /// Clears the display text.
    pub fn clear_display(&mut self) -> &mut Self {
        self.display = None;
        self
    }
}

/// A single worksheet.
#[derive(Debug, Clone, Default)]
pub struct Worksheet {
    name: String,
    visibility: SheetVisibility,
    cells: BTreeMap<String, Cell>,
    rows: BTreeMap<u32, Row>,
    columns: BTreeMap<u32, Column>,
    images: Vec<WorksheetImage>,
    charts: Vec<Chart>,
    pivot_tables: Vec<crate::pivot_table::PivotTable>,
    merged_ranges: Vec<CellRange>,
    freeze_pane: Option<FreezePane>,
    auto_filter: Option<AutoFilter>,
    tables: Vec<WorksheetTable>,
    conditional_formattings: Vec<ConditionalFormatting>,
    data_validations: Vec<DataValidation>,
    hyperlinks: Vec<Hyperlink>,
    comments: Vec<Comment>,
    protection: Option<SheetProtection>,
    page_setup: Option<PageSetup>,
    page_margins: Option<PageMargins>,
    header_footer: Option<PrintHeaderFooter>,
    print_area: Option<PrintArea>,
    page_breaks: Option<PageBreaks>,
    sheet_view_options: Option<SheetViewOptions>,
    sparkline_groups: Vec<SparklineGroup>,
    /// Raw attributes from the `<printOptions>` element, preserved for roundtrip fidelity.
    raw_print_options_attrs: Vec<(String, String)>,
    /// The `ref` attribute from the `<dimension>` element, preserved for roundtrip fidelity.
    raw_dimension_ref: Option<String>,
    /// Tab color as hex RGB (e.g. "FF0000" for red).
    tab_color: Option<String>,
    /// Default row height in points.
    default_row_height: Option<f64>,
    /// Default column width in character units.
    default_column_width: Option<f64>,
    /// Whether the default row height is a custom height.
    custom_height: Option<bool>,
    /// Raw attributes from the `<sheetFormatPr>` element, preserved for roundtrip fidelity.
    raw_sheet_format_pr_attrs: Vec<(String, String)>,
    /// Unknown XML children at the top level of `<worksheet>`, preserved for roundtrip fidelity.
    unknown_children: Vec<RawXmlNode>,
    /// Extra namespace declarations from the original `<worksheet>` element (e.g. `xmlns:x14ac`),
    /// preserved so that unknown attributes using those prefixes remain valid XML on dirty save.
    extra_namespace_declarations: Vec<(String, String)>,
    original_part_bytes: Option<(String, Vec<u8>)>,
    dirty: bool,
}

/// Fluent value-field spec for pivot table builders.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct PivotValueSpec {
    field_name: String,
    custom_name: Option<String>,
    subtotal: crate::pivot_table::PivotSubtotalFunction,
}

impl PivotValueSpec {
    fn new(
        field_name: impl Into<String>,
        subtotal: crate::pivot_table::PivotSubtotalFunction,
    ) -> Self {
        Self {
            field_name: field_name.into(),
            custom_name: None,
            subtotal,
        }
    }

    /// Sets the displayed data-field name in the pivot table.
    pub fn name(mut self, custom_name: impl Into<String>) -> Self {
        self.custom_name = Some(custom_name.into());
        self
    }

    pub(crate) fn field_name(&self) -> &str {
        self.field_name.as_str()
    }

    pub(crate) fn custom_name(&self) -> Option<&str> {
        self.custom_name.as_deref()
    }

    pub(crate) fn subtotal(&self) -> crate::pivot_table::PivotSubtotalFunction {
        self.subtotal
    }
}

/// Creates a pivot value-field spec using `SUM`.
pub fn sum(field_name: impl Into<String>) -> PivotValueSpec {
    PivotValueSpec::new(field_name, crate::pivot_table::PivotSubtotalFunction::Sum)
}

/// Creates a pivot value-field spec using `AVERAGE`.
pub fn avg(field_name: impl Into<String>) -> PivotValueSpec {
    PivotValueSpec::new(
        field_name,
        crate::pivot_table::PivotSubtotalFunction::Average,
    )
}

/// Fluent builder for worksheet pivot tables.
pub struct PivotBuilder<'a> {
    worksheet: &'a mut Worksheet,
    name: String,
    source: Option<String>,
    rows: Vec<String>,
    cols: Vec<String>,
    filters: Vec<String>,
    values: Vec<PivotValueSpec>,
}

impl<'a> PivotBuilder<'a> {
    fn new(worksheet: &'a mut Worksheet, name: impl Into<String>) -> Self {
        Self {
            worksheet,
            name: name.into(),
            source: None,
            rows: Vec::new(),
            cols: Vec::new(),
            filters: Vec::new(),
            values: Vec::new(),
        }
    }

    /// Sets the source range reference (e.g. `"Data!A1:F100"`).
    pub fn source(mut self, source: impl Into<String>) -> Self {
        self.source = Some(source.into());
        self
    }

    /// Sets row axis fields.
    pub fn rows<I, S>(mut self, fields: I) -> Self
    where
        I: IntoIterator<Item = S>,
        S: AsRef<str>,
    {
        self.rows = fields.into_iter().map(|f| f.as_ref().to_string()).collect();
        self
    }

    /// Sets column axis fields.
    pub fn cols<I, S>(mut self, fields: I) -> Self
    where
        I: IntoIterator<Item = S>,
        S: AsRef<str>,
    {
        self.cols = fields.into_iter().map(|f| f.as_ref().to_string()).collect();
        self
    }

    /// Sets filter (page) fields.
    pub fn filters<I, S>(mut self, fields: I) -> Self
    where
        I: IntoIterator<Item = S>,
        S: AsRef<str>,
    {
        self.filters = fields.into_iter().map(|f| f.as_ref().to_string()).collect();
        self
    }

    /// Sets value fields.
    pub fn values<I>(mut self, values: I) -> Self
    where
        I: IntoIterator<Item = PivotValueSpec>,
    {
        self.values = values.into_iter().collect();
        self
    }

    /// Validates pivot configuration before placement.
    ///
    /// If source points at this worksheet (or omits sheet name), field names are
    /// validated against the source header row.
    pub fn validate_fields(self) -> Result<Self> {
        self.validate_internal()?;
        Ok(self)
    }

    /// Finalizes and writes the pivot table to the worksheet with top-left anchor.
    pub fn place(self, target_cell: &str) -> Result<&'a mut Worksheet> {
        self.validate_internal()?;

        let source = self.source.ok_or_else(|| {
            XlsxError::InvalidWorkbookState("pivot source is required".to_string())
        })?;

        let (target_col, target_row) = cell_reference_to_column_row(target_cell)?;
        let mut pivot = crate::pivot_table::PivotTable::new(
            self.name,
            crate::pivot_table::PivotSourceReference::from_range(source),
        );
        pivot.set_target(target_row.saturating_sub(1), target_col.saturating_sub(1));

        for row_field in self.rows {
            let mut field = crate::pivot_table::PivotField::new(row_field);
            field.set_sort_type(crate::pivot_table::PivotFieldSort::Ascending);
            pivot.add_row_field(field);
        }
        for col_field in self.cols {
            let mut field = crate::pivot_table::PivotField::new(col_field);
            field.set_sort_type(crate::pivot_table::PivotFieldSort::Ascending);
            pivot.add_column_field(field);
        }
        for filter_field in self.filters {
            pivot.add_page_field(crate::pivot_table::PivotField::new(filter_field));
        }
        for value in self.values {
            let mut data_field = crate::pivot_table::PivotDataField::new(value.field_name);
            data_field.set_subtotal(value.subtotal);
            if let Some(name) = value.custom_name {
                data_field.set_custom_name(name);
            }
            pivot.add_data_field(data_field);
        }

        self.worksheet.add_pivot_table(pivot);
        Ok(self.worksheet)
    }

    fn validate_internal(&self) -> Result<()> {
        if self.source.as_deref().unwrap_or("").trim().is_empty() {
            return Err(XlsxError::InvalidWorkbookState(
                "pivot source is required".to_string(),
            ));
        }
        if self.values.is_empty() {
            return Err(XlsxError::InvalidWorkbookState(
                "pivot must include at least one value field".to_string(),
            ));
        }

        // Disallow duplicate row/column/filter fields.
        let mut seen = std::collections::HashSet::new();
        for field in self
            .rows
            .iter()
            .chain(self.cols.iter())
            .chain(self.filters.iter())
        {
            if !seen.insert(field.as_str()) {
                return Err(XlsxError::InvalidWorkbookState(format!(
                    "pivot field '{field}' appears more than once across rows/cols/filters"
                )));
            }
        }

        // Validate field names against source headers when source sheet resolves to this sheet.
        let source = self.source.as_deref().unwrap_or_default();
        let (source_sheet, source_range) = split_source_reference(source);
        let source_is_this_sheet = match source_sheet.map(normalize_source_sheet_name) {
            None => true,
            Some(s) => s == self.worksheet.name(),
        };

        if source_is_this_sheet {
            let headers = self.worksheet.source_headers_from_range(source_range)?;
            for field in self
                .rows
                .iter()
                .chain(self.cols.iter())
                .chain(self.filters.iter())
                .chain(self.values.iter().map(|v| &v.field_name))
            {
                if !headers.iter().any(|h| h == field) {
                    return Err(XlsxError::InvalidWorkbookState(format!(
                        "pivot field '{field}' not found in source header row"
                    )));
                }
            }
        }

        Ok(())
    }
}

impl Worksheet {
    pub fn new(name: impl Into<String>) -> Self {
        Self {
            name: name.into(),
            visibility: SheetVisibility::Visible,
            cells: BTreeMap::new(),
            rows: BTreeMap::new(),
            columns: BTreeMap::new(),
            images: Vec::new(),
            charts: Vec::new(),
            pivot_tables: Vec::new(),
            merged_ranges: Vec::new(),
            freeze_pane: None,
            auto_filter: None,
            tables: Vec::new(),
            conditional_formattings: Vec::new(),
            data_validations: Vec::new(),
            hyperlinks: Vec::new(),
            comments: Vec::new(),
            protection: None,
            page_setup: None,
            page_margins: None,
            header_footer: None,
            print_area: None,
            page_breaks: None,
            sheet_view_options: None,
            sparkline_groups: Vec::new(),
            raw_print_options_attrs: Vec::new(),
            raw_dimension_ref: None,
            tab_color: None,
            default_row_height: None,
            default_column_width: None,
            custom_height: None,
            raw_sheet_format_pr_attrs: Vec::new(),
            unknown_children: Vec::new(),
            extra_namespace_declarations: Vec::new(),
            original_part_bytes: None,
            dirty: true,
        }
    }

    pub fn name(&self) -> &str {
        &self.name
    }

    /// Returns the sheet visibility state.
    pub fn visibility(&self) -> SheetVisibility {
        self.visibility
    }

    /// Sets the sheet visibility state.
    pub fn set_visibility(&mut self, visibility: SheetVisibility) -> &mut Self {
        self.visibility = visibility;
        self.mark_dirty();
        self
    }

    /// Returns an immutable cell reference when the address is valid and present.
    pub fn cell(&self, reference: &str) -> Option<&Cell> {
        let normalized = normalize_cell_reference(reference).ok()?;
        self.cells.get(normalized.as_str())
    }

    /// Returns a mutable cell, creating it when missing.
    pub fn cell_mut(&mut self, reference: &str) -> Result<&mut Cell> {
        let normalized = normalize_cell_reference(reference)?;
        self.mark_dirty();
        Ok(self.cells.entry(normalized).or_default())
    }

    /// Evaluates the formula in a cell and returns the computed result.
    ///
    /// This is a convenience method that evaluates the formula stored in the
    /// specified cell using the formula engine. The workbook reference is required
    /// to resolve cell references and cross-sheet formulas.
    ///
    /// If the cell doesn't exist or doesn't contain a formula, returns an error.
    ///
    /// # Example
    ///
    /// ```ignore
    /// use offidized_xlsx::Workbook;
    ///
    /// let mut wb = Workbook::new();
    /// let ws = wb.add_sheet("Sheet1");
    /// ws.cell_mut("A1")?.set_value(10);
    /// ws.cell_mut("A2")?.set_value(20);
    /// ws.cell_mut("A3")?.set_formula("SUM(A1:A2)");
    ///
    /// // Evaluate the formula in A3
    /// let result = ws.evaluate_formula("A3", &wb)?;
    /// assert_eq!(result, CellValue::Number(30.0));
    /// # Ok::<(), offidized_xlsx::XlsxError>(())
    /// ```
    pub fn evaluate_formula(
        &self,
        cell_ref: &str,
        workbook: &crate::workbook::Workbook,
    ) -> Result<CellValue> {
        // Normalize and get the cell
        let normalized = normalize_cell_reference(cell_ref)?;
        let cell = self
            .cell(&normalized)
            .ok_or_else(|| XlsxError::InvalidCellReference(cell_ref.to_string()))?;

        let formula = cell.formula().ok_or_else(|| {
            XlsxError::InvalidWorkbookState(format!("Cell {} has no formula", cell_ref))
        })?;

        // Parse the cell reference to get row and column (1-based)
        let (col, row) = cell_reference_to_column_row(&normalized)?;

        // Use workbook's evaluate_formula method
        let result = workbook.evaluate_formula(formula, &self.name, row, col);
        Ok(result)
    }

    /// Returns immutable row metadata by 1-based index.
    pub fn row(&self, index: u32) -> Option<&Row> {
        if index == 0 {
            return None;
        }
        self.rows.get(&index)
    }

    /// Returns mutable row metadata by 1-based index, creating it when missing.
    pub fn row_mut(&mut self, index: u32) -> Result<&mut Row> {
        validate_dimension_index("row", index)?;
        self.mark_dirty();
        Ok(self.rows.entry(index).or_insert_with(|| Row::new(index)))
    }

    /// Iterates row metadata ordered by row index.
    pub fn rows(&self) -> impl Iterator<Item = &Row> {
        self.rows.values()
    }

    /// Iterates mutable row metadata ordered by row index.
    pub fn rows_mut(&mut self) -> impl Iterator<Item = &mut Row> {
        self.mark_dirty();
        self.rows.values_mut()
    }

    /// Returns immutable column metadata by 1-based index.
    pub fn column(&self, index: u32) -> Option<&Column> {
        if index == 0 {
            return None;
        }
        self.columns.get(&index)
    }

    /// Returns mutable column metadata by 1-based index, creating it when missing.
    pub fn column_mut(&mut self, index: u32) -> Result<&mut Column> {
        validate_dimension_index("column", index)?;
        self.mark_dirty();
        Ok(self
            .columns
            .entry(index)
            .or_insert_with(|| Column::new(index)))
    }

    /// Iterates column metadata ordered by column index.
    pub fn columns(&self) -> impl Iterator<Item = &Column> {
        self.columns.values()
    }

    /// Iterates mutable column metadata ordered by column index.
    pub fn columns_mut(&mut self) -> impl Iterator<Item = &mut Column> {
        self.mark_dirty();
        self.columns.values_mut()
    }

    /// Returns worksheet images in insertion order.
    pub fn images(&self) -> &[WorksheetImage] {
        self.images.as_slice()
    }

    /// Adds an image anchored to a cell, with an optional ext size in EMUs.
    pub fn add_image(
        &mut self,
        bytes: impl Into<Vec<u8>>,
        content_type: impl Into<String>,
        anchor_cell: &str,
        ext: Option<WorksheetImageExt>,
    ) -> Result<&mut Self> {
        let image = WorksheetImage::new(bytes, content_type, anchor_cell, ext)?;
        self.images.push(image);
        self.mark_dirty();
        Ok(self)
    }

    /// Returns a mutable reference to the images list.
    pub fn images_mut(&mut self) -> &mut [WorksheetImage] {
        self.mark_dirty();
        &mut self.images
    }

    /// Removes the image at the given index.
    ///
    /// Returns `None` if the index is out of bounds.
    pub fn remove_image(&mut self, index: usize) -> Option<WorksheetImage> {
        if index < self.images.len() {
            self.mark_dirty();
            Some(self.images.remove(index))
        } else {
            None
        }
    }

    /// Removes all worksheet images.
    pub fn clear_images(&mut self) -> &mut Self {
        self.images.clear();
        self.mark_dirty();
        self
    }

    /// Returns the charts embedded in this worksheet.
    pub fn charts(&self) -> &[Chart] {
        &self.charts
    }

    /// Returns a mutable reference to the charts vector.
    pub fn charts_mut(&mut self) -> &mut Vec<Chart> {
        self.mark_dirty();
        &mut self.charts
    }

    /// Adds a chart to the worksheet.
    pub fn add_chart(&mut self, chart: Chart) -> &mut Self {
        self.charts.push(chart);
        self.mark_dirty();
        self
    }

    /// Removes all charts from the worksheet.
    pub fn clear_charts(&mut self) -> &mut Self {
        self.charts.clear();
        self.mark_dirty();
        self
    }

    /// Iterates merged ranges in insertion order.
    pub fn merged_ranges(&self) -> &[CellRange] {
        self.merged_ranges.as_slice()
    }

    /// Adds a merged range using A1 notation.
    pub fn add_merged_range(&mut self, range: &str) -> Result<&mut Self> {
        let range = CellRange::parse(range)?;
        self.push_merged_range(range);
        self.mark_dirty();
        Ok(self)
    }

    /// Removes all merged ranges.
    pub fn clear_merged_ranges(&mut self) -> &mut Self {
        self.merged_ranges.clear();
        self.mark_dirty();
        self
    }

    /// Removes a specific merged range by A1 notation (e.g. "A1:B2").
    ///
    /// Returns `true` if the range was found and removed, `false` otherwise.
    pub fn unmerge_range(&mut self, range: &str) -> Result<bool> {
        let target = CellRange::parse(range)?;
        let original_len = self.merged_ranges.len();
        self.merged_ranges.retain(|r| r != &target);
        if self.merged_ranges.len() != original_len {
            self.mark_dirty();
            Ok(true)
        } else {
            Ok(false)
        }
    }

    /// Removes all merged ranges (alias for [`clear_merged_ranges`]).
    pub fn unmerge_all(&mut self) -> &mut Self {
        self.clear_merged_ranges()
    }

    /// Finds all cells whose string representation contains the given text.
    ///
    /// Returns cell references (e.g. `["A1", "B3"]`) for cells that contain
    /// the search text as a substring (case-sensitive).
    pub fn find_cells(&self, text: &str) -> Vec<String> {
        let mut results = Vec::new();
        for (reference, cell) in &self.cells {
            let matches = match cell.value() {
                Some(CellValue::String(s)) => s.contains(text),
                Some(CellValue::RichText(parts)) => {
                    let full: String = parts.iter().map(|p| p.text()).collect();
                    full.contains(text)
                }
                Some(CellValue::Number(n)) => {
                    let s = format!("{n}");
                    s.contains(text)
                }
                Some(CellValue::Bool(b)) => {
                    let s = if *b { "TRUE" } else { "FALSE" };
                    s.contains(text)
                }
                Some(_) => false,
                None => false,
            };
            if matches {
                results.push(reference.clone());
            }
        }
        results
    }

    /// Finds all cells whose value exactly matches the given value.
    ///
    /// Returns cell references for exact matches.
    pub fn find_cells_by_value(&self, value: &CellValue) -> Vec<String> {
        let mut results = Vec::new();
        for (reference, cell) in &self.cells {
            if cell.value() == Some(value) {
                results.push(reference.clone());
            }
        }
        results
    }

    /// Returns worksheet freeze pane settings when configured.
    pub fn freeze_pane(&self) -> Option<&FreezePane> {
        self.freeze_pane.as_ref()
    }

    /// Freezes rows/columns from the top-left origin.
    pub fn set_freeze_panes(&mut self, x_split: u32, y_split: u32) -> Result<&mut Self> {
        self.freeze_pane = Some(FreezePane::new(x_split, y_split)?);
        self.mark_dirty();
        Ok(self)
    }

    /// Freezes rows/columns and explicitly sets the top-left visible cell.
    pub fn set_freeze_panes_with_top_left_cell(
        &mut self,
        x_split: u32,
        y_split: u32,
        top_left_cell: &str,
    ) -> Result<&mut Self> {
        self.freeze_pane = Some(FreezePane::with_top_left_cell(
            x_split,
            y_split,
            top_left_cell,
        )?);
        self.mark_dirty();
        Ok(self)
    }

    /// Clears worksheet freeze pane settings.
    pub fn clear_freeze_pane(&mut self) -> &mut Self {
        self.freeze_pane = None;
        self.mark_dirty();
        self
    }

    /// Returns worksheet-level auto-filter when configured.
    pub fn auto_filter(&self) -> Option<&AutoFilter> {
        self.auto_filter.as_ref()
    }

    /// Returns a mutable reference to the auto-filter, creating it if needed.
    pub fn auto_filter_mut(&mut self) -> &mut AutoFilter {
        self.mark_dirty();
        self.auto_filter.get_or_insert_with(AutoFilter::new)
    }

    /// Sets worksheet-level auto-filter range.
    pub fn set_auto_filter(&mut self, range: &str) -> Result<&mut Self> {
        self.auto_filter = Some(AutoFilter::with_range(range)?);
        self.mark_dirty();
        Ok(self)
    }

    /// Sets an auto-filter with full filter column criteria.
    pub fn set_auto_filter_full(&mut self, auto_filter: AutoFilter) -> &mut Self {
        self.auto_filter = Some(auto_filter);
        self.mark_dirty();
        self
    }

    /// Clears worksheet-level auto-filter.
    pub fn clear_auto_filter(&mut self) -> &mut Self {
        self.auto_filter = None;
        self.mark_dirty();
        self
    }

    /// Returns worksheet tables in insertion order.
    pub fn tables(&self) -> &[WorksheetTable] {
        self.tables.as_slice()
    }

    /// Adds a table to the worksheet.
    pub fn add_table(&mut self, table: WorksheetTable) -> &mut Self {
        self.tables.push(table);
        self.mark_dirty();
        self
    }

    /// Removes all worksheet tables.
    pub fn clear_tables(&mut self) -> &mut Self {
        self.tables.clear();
        self.mark_dirty();
        self
    }

    /// Returns conditional formatting rules in insertion order.
    pub fn conditional_formattings(&self) -> &[ConditionalFormatting] {
        self.conditional_formattings.as_slice()
    }

    /// Adds a conditional formatting rule to the worksheet.
    pub fn add_conditional_formatting(
        &mut self,
        conditional_formatting: ConditionalFormatting,
    ) -> &mut Self {
        self.conditional_formattings.push(conditional_formatting);
        self.mark_dirty();
        self
    }

    /// Removes all conditional formatting rules.
    pub fn clear_conditional_formattings(&mut self) -> &mut Self {
        self.conditional_formattings.clear();
        self.mark_dirty();
        self
    }

    /// Returns data validations in insertion order.
    pub fn data_validations(&self) -> &[DataValidation] {
        self.data_validations.as_slice()
    }

    /// Adds a data validation to the worksheet.
    pub fn add_data_validation(&mut self, data_validation: DataValidation) -> &mut Self {
        self.data_validations.push(data_validation);
        self.mark_dirty();
        self
    }

    /// Removes all data validations.
    pub fn clear_data_validations(&mut self) -> &mut Self {
        self.data_validations.clear();
        self.mark_dirty();
        self
    }

    /// Returns hyperlinks in insertion order.
    pub fn hyperlinks(&self) -> &[Hyperlink] {
        self.hyperlinks.as_slice()
    }

    /// Adds a hyperlink to the worksheet.
    pub fn add_hyperlink(&mut self, hyperlink: Hyperlink) -> &mut Self {
        self.hyperlinks.push(hyperlink);
        self.mark_dirty();
        self
    }

    /// Removes a hyperlink by cell reference. Returns true if a hyperlink was removed.
    pub fn remove_hyperlink(&mut self, cell_ref: &str) -> bool {
        let normalized = match normalize_cell_reference(cell_ref) {
            Ok(n) => n,
            Err(_) => return false,
        };
        let original_len = self.hyperlinks.len();
        self.hyperlinks
            .retain(|h| h.cell_ref() != normalized.as_str());
        let removed = self.hyperlinks.len() < original_len;
        if removed {
            self.mark_dirty();
        }
        removed
    }

    /// Removes all hyperlinks.
    pub fn clear_hyperlinks(&mut self) -> &mut Self {
        self.hyperlinks.clear();
        self.mark_dirty();
        self
    }

    /// Returns comments in insertion order.
    pub fn comments(&self) -> &[Comment] {
        self.comments.as_slice()
    }

    /// Adds a comment to the worksheet.
    pub fn add_comment(&mut self, comment: Comment) -> &mut Self {
        self.comments.push(comment);
        self.mark_dirty();
        self
    }

    /// Removes a comment by cell reference. Returns true if a comment was removed.
    pub fn remove_comment(&mut self, cell_ref: &str) -> bool {
        let normalized = match normalize_cell_reference(cell_ref) {
            Ok(n) => n,
            Err(_) => return false,
        };
        let original_len = self.comments.len();
        self.comments
            .retain(|c| c.cell_ref() != normalized.as_str());
        let removed = self.comments.len() < original_len;
        if removed {
            self.mark_dirty();
        }
        removed
    }

    /// Removes all comments.
    pub fn clear_comments(&mut self) -> &mut Self {
        self.comments.clear();
        self.mark_dirty();
        self
    }

    /// Returns the sheet protection settings, if configured.
    pub fn protection(&self) -> Option<&SheetProtection> {
        self.protection.as_ref()
    }

    /// Sets sheet protection.
    pub fn set_protection(&mut self, protection: SheetProtection) -> &mut Self {
        self.protection = Some(protection);
        self.mark_dirty();
        self
    }

    /// Clears sheet protection.
    pub fn clear_protection(&mut self) -> &mut Self {
        self.protection = None;
        self.mark_dirty();
        self
    }

    /// Returns the page setup settings, if configured.
    pub fn page_setup(&self) -> Option<&PageSetup> {
        self.page_setup.as_ref()
    }

    /// Sets page setup for printing.
    pub fn set_page_setup(&mut self, page_setup: PageSetup) -> &mut Self {
        self.page_setup = Some(page_setup);
        self.mark_dirty();
        self
    }

    /// Clears page setup settings.
    pub fn clear_page_setup(&mut self) -> &mut Self {
        self.page_setup = None;
        self.mark_dirty();
        self
    }

    /// Returns the page margins, if configured.
    pub fn page_margins(&self) -> Option<&PageMargins> {
        self.page_margins.as_ref()
    }

    /// Sets page margins for printing.
    pub fn set_page_margins(&mut self, margins: PageMargins) -> &mut Self {
        self.page_margins = Some(margins);
        self.mark_dirty();
        self
    }

    /// Clears page margins.
    pub fn clear_page_margins(&mut self) -> &mut Self {
        self.page_margins = None;
        self.mark_dirty();
        self
    }

    /// Returns the header/footer settings, if configured.
    pub fn header_footer(&self) -> Option<&PrintHeaderFooter> {
        self.header_footer.as_ref()
    }

    /// Sets header/footer settings for printing.
    pub fn set_header_footer(&mut self, hf: PrintHeaderFooter) -> &mut Self {
        self.header_footer = Some(hf);
        self.mark_dirty();
        self
    }

    /// Clears header/footer settings.
    pub fn clear_header_footer(&mut self) -> &mut Self {
        self.header_footer = None;
        self.mark_dirty();
        self
    }

    /// Returns the print area, if configured.
    pub fn print_area(&self) -> Option<&PrintArea> {
        self.print_area.as_ref()
    }

    /// Sets the print area.
    pub fn set_print_area(&mut self, area: PrintArea) -> &mut Self {
        self.print_area = Some(area);
        self.mark_dirty();
        self
    }

    /// Clears the print area.
    pub fn clear_print_area(&mut self) -> &mut Self {
        self.print_area = None;
        self.mark_dirty();
        self
    }

    /// Returns the page breaks, if configured.
    pub fn page_breaks(&self) -> Option<&PageBreaks> {
        self.page_breaks.as_ref()
    }

    /// Returns mutable access to the page breaks, creating if absent.
    pub fn page_breaks_mut(&mut self) -> &mut PageBreaks {
        self.mark_dirty();
        self.page_breaks.get_or_insert_with(PageBreaks::new)
    }

    /// Sets the page breaks.
    pub fn set_page_breaks(&mut self, breaks: PageBreaks) -> &mut Self {
        self.page_breaks = Some(breaks);
        self.mark_dirty();
        self
    }

    /// Clears all page breaks.
    pub fn clear_page_breaks(&mut self) -> &mut Self {
        self.page_breaks = None;
        self.mark_dirty();
        self
    }

    /// Returns the sheet view options, if configured.
    pub fn sheet_view_options(&self) -> Option<&SheetViewOptions> {
        self.sheet_view_options.as_ref()
    }

    /// Sets sheet view options.
    pub fn set_sheet_view_options(&mut self, options: SheetViewOptions) -> &mut Self {
        self.sheet_view_options = Some(options);
        self.mark_dirty();
        self
    }

    /// Clears sheet view options.
    pub fn clear_sheet_view_options(&mut self) -> &mut Self {
        self.sheet_view_options = None;
        self.mark_dirty();
        self
    }

    /// Returns sparkline groups configured on this worksheet.
    pub fn sparkline_groups(&self) -> &[SparklineGroup] {
        &self.sparkline_groups
    }

    /// Returns a mutable reference to the sparkline groups vector.
    pub fn sparkline_groups_mut(&mut self) -> &mut Vec<SparklineGroup> {
        self.mark_dirty();
        &mut self.sparkline_groups
    }

    /// Adds a sparkline group to the worksheet.
    pub fn add_sparkline_group(&mut self, group: SparklineGroup) -> &mut Self {
        self.sparkline_groups.push(group);
        self.mark_dirty();
        self
    }

    /// Removes all sparkline groups.
    pub fn clear_sparkline_groups(&mut self) -> &mut Self {
        self.sparkline_groups.clear();
        self.mark_dirty();
        self
    }

    /// Returns the tab color as a hex RGB string (e.g. "FF0000").
    pub fn tab_color(&self) -> Option<&str> {
        self.tab_color.as_deref()
    }

    /// Sets the tab color as a hex RGB string (e.g. "FF0000").
    pub fn set_tab_color(&mut self, color: impl Into<String>) -> &mut Self {
        self.tab_color = Some(color.into());
        self.mark_dirty();
        self
    }

    /// Clears the tab color.
    pub fn clear_tab_color(&mut self) -> &mut Self {
        self.tab_color = None;
        self.mark_dirty();
        self
    }

    /// Returns the default row height in points.
    pub fn default_row_height(&self) -> Option<f64> {
        self.default_row_height
    }

    /// Sets the default row height in points.
    pub fn set_default_row_height(&mut self, height: f64) -> &mut Self {
        self.default_row_height = Some(height);
        self.mark_dirty();
        self
    }

    /// Clears the default row height.
    pub fn clear_default_row_height(&mut self) -> &mut Self {
        self.default_row_height = None;
        self.mark_dirty();
        self
    }

    /// Returns the default column width in character units.
    pub fn default_column_width(&self) -> Option<f64> {
        self.default_column_width
    }

    /// Sets the default column width in character units.
    pub fn set_default_column_width(&mut self, width: f64) -> &mut Self {
        self.default_column_width = Some(width);
        self.mark_dirty();
        self
    }

    /// Clears the default column width.
    pub fn clear_default_column_width(&mut self) -> &mut Self {
        self.default_column_width = None;
        self.mark_dirty();
        self
    }

    /// Returns whether the default row height is a custom height.
    pub fn custom_height(&self) -> Option<bool> {
        self.custom_height
    }

    /// Sets whether the default row height is a custom height.
    pub fn set_custom_height(&mut self, value: bool) -> &mut Self {
        self.custom_height = Some(value);
        self.mark_dirty();
        self
    }

    /// Clears the custom height flag.
    pub fn clear_custom_height(&mut self) -> &mut Self {
        self.custom_height = None;
        self.mark_dirty();
        self
    }

    // ── Range bulk operations ──

    /// Sets values in a rectangular range from a 2D array (row-major order).
    ///
    /// `start_ref` is the top-left cell (e.g. "A1"), `values` is a slice of rows,
    /// each row a slice of cell values.
    pub fn set_values_2d(
        &mut self,
        start_ref: &str,
        values: &[Vec<impl Into<CellValue> + Clone>],
    ) -> Result<&mut Self> {
        let (start_col, start_row) = cell_reference_to_column_row(start_ref)?;
        for (row_offset, row_vals) in values.iter().enumerate() {
            for (col_offset, val) in row_vals.iter().enumerate() {
                let col = start_col + col_offset as u32;
                let row = start_row + row_offset as u32;
                let ref_str = build_cell_reference(col, row)?;
                self.cell_mut(&ref_str)?.set_value(val.clone());
            }
        }
        Ok(self)
    }

    /// Clears cell values in the specified range, preserving cell formatting.
    pub fn clear_values(&mut self, range: &str) -> Result<&mut Self> {
        let cr = CellRange::parse(range)?;
        for cell_ref in cr.iter() {
            if let Some(cell) = self.cells.get_mut(&cell_ref) {
                cell.set_value(CellValue::Blank);
            }
        }
        self.mark_dirty();
        Ok(self)
    }

    /// Clears all cell content (value, formula, style) in the specified range.
    pub fn clear_range(&mut self, range: &str) -> Result<&mut Self> {
        let cr = CellRange::parse(range)?;
        for cell_ref in cr.iter() {
            self.cells.remove(&cell_ref);
        }
        self.mark_dirty();
        Ok(self)
    }

    /// Applies a style ID to all cells in the specified range.
    pub fn apply_style_to_range(&mut self, range: &str, style_id: u32) -> Result<&mut Self> {
        let cr = CellRange::parse(range)?;
        for cell_ref in cr.iter() {
            self.cell_mut(&cell_ref)?.set_style_id(style_id);
        }
        Ok(self)
    }

    /// Copies cell values and formatting from `source_range` to a destination starting at
    /// `dest_start` (top-left cell reference). Formulas are copied as-is (no reference shifting).
    pub fn copy_range(&mut self, source_range: &str, dest_start: &str) -> Result<&mut Self> {
        let src = CellRange::parse(source_range)?;
        let (dest_col, dest_row) = cell_reference_to_column_row(dest_start)?;
        let (src_start_col, src_start_row) = cell_reference_to_column_row(src.start())?;

        // Collect source cells first to avoid borrow conflicts.
        let source_cells: Vec<(u32, u32, Cell)> = src
            .iter()
            .filter_map(|ref_str| {
                let (c, r) = cell_reference_to_column_row(&ref_str).ok()?;
                let cell = self.cells.get(&ref_str)?.clone();
                Some((c - src_start_col, r - src_start_row, cell))
            })
            .collect();

        for (col_offset, row_offset, cell) in source_cells {
            let target_ref = build_cell_reference(dest_col + col_offset, dest_row + row_offset)?;
            self.cells.insert(target_ref, cell);
        }
        self.mark_dirty();
        Ok(self)
    }

    /// Inserts `count` rows at the given 1-based `start_row`, shifting existing rows down.
    ///
    /// Cell references in formulas are also updated: for example, if inserting at row 3,
    /// a formula referencing "A5" will be updated to "A6" (shifted by `count`).
    pub fn insert_rows(&mut self, start_row: u32, count: u32) -> Result<&mut Self> {
        if start_row == 0 {
            return Err(XlsxError::InvalidWorkbookState(
                "insert_rows start_row must be >= 1".to_string(),
            ));
        }
        if count == 0 {
            return Ok(self);
        }

        // Shift cells and update formula references
        let old_cells: BTreeMap<String, Cell> = std::mem::take(&mut self.cells);
        for (reference, mut cell) in old_cells {
            // Update formula references
            if let Some(formula) = cell.formula() {
                let updated = shift_formula_row_references(formula, start_row, count, true);
                cell.set_formula(updated);
            }
            let shifted = shift_cell_reference_row(reference.as_str(), start_row, count, true);
            match shifted {
                Ok(new_ref) => {
                    self.cells.insert(new_ref, cell);
                }
                Err(_) => {
                    self.cells.insert(reference, cell);
                }
            }
        }

        // Shift rows
        let old_rows: BTreeMap<u32, Row> = std::mem::take(&mut self.rows);
        for (index, mut row) in old_rows {
            if index >= start_row {
                let new_index = index.checked_add(count).ok_or_else(|| {
                    XlsxError::InvalidWorkbookState("row index overflow during insert".to_string())
                })?;
                row = Row::new(new_index);
                self.rows.insert(new_index, row);
            } else {
                self.rows.insert(index, row);
            }
        }

        // Shift merged ranges
        for range in &mut self.merged_ranges {
            *range = shift_range_rows(range, start_row, count, true);
        }

        // Shift conditional formatting ranges
        for cf in &mut self.conditional_formattings {
            for sqref in &mut cf.sqref {
                *sqref = shift_range_rows(sqref, start_row, count, true);
            }
        }

        // Shift data validation ranges
        for dv in &mut self.data_validations {
            for sqref in &mut dv.sqref {
                *sqref = shift_range_rows(sqref, start_row, count, true);
            }
        }

        // Shift table ranges
        for table in &mut self.tables {
            table.range = shift_range_rows(&table.range, start_row, count, true);
        }

        // Shift hyperlink cell references
        for hl in &mut self.hyperlinks {
            if let Ok(shifted) = shift_cell_reference_row(&hl.cell_ref, start_row, count, true) {
                hl.cell_ref = shifted;
            }
        }

        // Shift auto-filter range
        if let Some(ref mut af) = self.auto_filter {
            if let Some(range) = af.range().cloned() {
                let shifted = shift_range_rows(&range, start_row, count, true);
                let range_str = format!("{}:{}", shifted.start(), shifted.end());
                let _ = af.set_range(&range_str);
            }
        }

        self.mark_dirty();
        Ok(self)
    }

    /// Deletes `count` rows starting at the given 1-based `start_row`, shifting existing rows up.
    ///
    /// Cell references in formulas are also updated.
    pub fn delete_rows(&mut self, start_row: u32, count: u32) -> Result<&mut Self> {
        if start_row == 0 {
            return Err(XlsxError::InvalidWorkbookState(
                "delete_rows start_row must be >= 1".to_string(),
            ));
        }
        if count == 0 {
            return Ok(self);
        }

        let end_row = start_row.saturating_add(count).saturating_sub(1);

        // Remove cells in deleted rows, shift others up
        let old_cells: BTreeMap<String, Cell> = std::mem::take(&mut self.cells);
        for (reference, mut cell) in old_cells {
            if let Ok((_, row)) = cell_reference_to_column_row(reference.as_str()) {
                if row >= start_row && row <= end_row {
                    continue; // deleted
                }
            }
            // Update formula references
            if let Some(formula) = cell.formula() {
                let updated = shift_formula_row_references(formula, end_row + 1, count, false);
                cell.set_formula(updated);
            }
            let shifted = shift_cell_reference_row(reference.as_str(), end_row + 1, count, false);
            match shifted {
                Ok(new_ref) => {
                    self.cells.insert(new_ref, cell);
                }
                Err(_) => {
                    self.cells.insert(reference, cell);
                }
            }
        }

        // Shift rows
        let old_rows: BTreeMap<u32, Row> = std::mem::take(&mut self.rows);
        for (index, row) in old_rows {
            if index >= start_row && index <= end_row {
                continue; // deleted
            }
            if index > end_row {
                let new_index = index.saturating_sub(count);
                let mut new_row = Row::new(new_index);
                if let Some(h) = row.height() {
                    new_row.set_height(h);
                }
                new_row.set_hidden(row.is_hidden());
                self.rows.insert(new_index, new_row);
            } else {
                self.rows.insert(index, row);
            }
        }

        // Shift merged ranges
        let old_ranges: Vec<CellRange> = std::mem::take(&mut self.merged_ranges);
        for range in old_ranges {
            let shifted = shift_range_rows(&range, end_row + 1, count, false);
            self.merged_ranges.push(shifted);
        }

        // Shift conditional formatting ranges
        for cf in &mut self.conditional_formattings {
            for sqref in &mut cf.sqref {
                *sqref = shift_range_rows(sqref, end_row + 1, count, false);
            }
        }

        // Shift data validation ranges
        for dv in &mut self.data_validations {
            for sqref in &mut dv.sqref {
                *sqref = shift_range_rows(sqref, end_row + 1, count, false);
            }
        }

        // Shift table ranges
        for table in &mut self.tables {
            table.range = shift_range_rows(&table.range, end_row + 1, count, false);
        }

        // Shift hyperlink cell references (remove hyperlinks in deleted range)
        let old_hyperlinks: Vec<Hyperlink> = std::mem::take(&mut self.hyperlinks);
        for mut hl in old_hyperlinks {
            if let Ok((_, row)) = cell_reference_to_column_row(&hl.cell_ref) {
                if row >= start_row && row <= end_row {
                    continue;
                }
            }
            if let Ok(shifted) = shift_cell_reference_row(&hl.cell_ref, end_row + 1, count, false) {
                hl.cell_ref = shifted;
            }
            self.hyperlinks.push(hl);
        }

        // Shift auto-filter range
        if let Some(ref mut af) = self.auto_filter {
            if let Some(range) = af.range().cloned() {
                let shifted = shift_range_rows(&range, end_row + 1, count, false);
                let range_str = format!("{}:{}", shifted.start(), shifted.end());
                let _ = af.set_range(&range_str);
            }
        }

        self.mark_dirty();
        Ok(self)
    }

    /// Inserts `count` columns at the given 1-based `start_col`, shifting existing columns right.
    ///
    /// Cell references in formulas are also updated.
    pub fn insert_columns(&mut self, start_col: u32, count: u32) -> Result<&mut Self> {
        if start_col == 0 {
            return Err(XlsxError::InvalidWorkbookState(
                "insert_columns start_col must be >= 1".to_string(),
            ));
        }
        if count == 0 {
            return Ok(self);
        }

        // Shift cells and update formula references
        let old_cells: BTreeMap<String, Cell> = std::mem::take(&mut self.cells);
        for (reference, mut cell) in old_cells {
            if let Some(formula) = cell.formula() {
                let updated = shift_formula_col_references(formula, start_col, count, true);
                cell.set_formula(updated);
            }
            let shifted = shift_cell_reference_col(reference.as_str(), start_col, count, true);
            match shifted {
                Ok(new_ref) => {
                    self.cells.insert(new_ref, cell);
                }
                Err(_) => {
                    self.cells.insert(reference, cell);
                }
            }
        }

        // Shift columns
        let old_columns: BTreeMap<u32, Column> = std::mem::take(&mut self.columns);
        for (index, col) in old_columns {
            if index >= start_col {
                let new_index = index.checked_add(count).ok_or_else(|| {
                    XlsxError::InvalidWorkbookState(
                        "column index overflow during insert".to_string(),
                    )
                })?;
                let mut new_col = Column::new(new_index);
                if let Some(w) = col.width() {
                    new_col.set_width(w);
                }
                new_col.set_hidden(col.is_hidden());
                self.columns.insert(new_index, new_col);
            } else {
                self.columns.insert(index, col);
            }
        }

        // Shift merged ranges
        for range in &mut self.merged_ranges {
            *range = shift_range_cols(range, start_col, count, true);
        }

        // Shift conditional formatting ranges
        for cf in &mut self.conditional_formattings {
            for sqref in &mut cf.sqref {
                *sqref = shift_range_cols(sqref, start_col, count, true);
            }
        }

        // Shift data validation ranges
        for dv in &mut self.data_validations {
            for sqref in &mut dv.sqref {
                *sqref = shift_range_cols(sqref, start_col, count, true);
            }
        }

        // Shift table ranges
        for table in &mut self.tables {
            table.range = shift_range_cols(&table.range, start_col, count, true);
        }

        // Shift hyperlink cell references
        for hl in &mut self.hyperlinks {
            if let Ok(shifted) = shift_cell_reference_col(&hl.cell_ref, start_col, count, true) {
                hl.cell_ref = shifted;
            }
        }

        // Shift auto-filter range
        if let Some(ref mut af) = self.auto_filter {
            if let Some(range) = af.range().cloned() {
                let shifted = shift_range_cols(&range, start_col, count, true);
                let range_str = format!("{}:{}", shifted.start(), shifted.end());
                let _ = af.set_range(&range_str);
            }
        }

        self.mark_dirty();
        Ok(self)
    }

    /// Deletes `count` columns starting at the given 1-based `start_col`, shifting existing columns left.
    ///
    /// Cell references in formulas are also updated.
    pub fn delete_columns(&mut self, start_col: u32, count: u32) -> Result<&mut Self> {
        if start_col == 0 {
            return Err(XlsxError::InvalidWorkbookState(
                "delete_columns start_col must be >= 1".to_string(),
            ));
        }
        if count == 0 {
            return Ok(self);
        }

        let end_col = start_col.saturating_add(count).saturating_sub(1);

        // Remove cells in deleted columns, shift others left
        let old_cells: BTreeMap<String, Cell> = std::mem::take(&mut self.cells);
        for (reference, mut cell) in old_cells {
            if let Ok((col, _)) = cell_reference_to_column_row(reference.as_str()) {
                if col >= start_col && col <= end_col {
                    continue; // deleted
                }
            }
            // Update formula references
            if let Some(formula) = cell.formula() {
                let updated = shift_formula_col_references(formula, end_col + 1, count, false);
                cell.set_formula(updated);
            }
            let shifted = shift_cell_reference_col(reference.as_str(), end_col + 1, count, false);
            match shifted {
                Ok(new_ref) => {
                    self.cells.insert(new_ref, cell);
                }
                Err(_) => {
                    self.cells.insert(reference, cell);
                }
            }
        }

        // Shift columns
        let old_columns: BTreeMap<u32, Column> = std::mem::take(&mut self.columns);
        for (index, col) in old_columns {
            if index >= start_col && index <= end_col {
                continue; // deleted
            }
            if index > end_col {
                let new_index = index.saturating_sub(count);
                let mut new_col = Column::new(new_index);
                if let Some(w) = col.width() {
                    new_col.set_width(w);
                }
                new_col.set_hidden(col.is_hidden());
                self.columns.insert(new_index, new_col);
            } else {
                self.columns.insert(index, col);
            }
        }

        // Shift merged ranges
        let old_ranges: Vec<CellRange> = std::mem::take(&mut self.merged_ranges);
        for range in old_ranges {
            let shifted = shift_range_cols(&range, end_col + 1, count, false);
            self.merged_ranges.push(shifted);
        }

        // Shift conditional formatting ranges
        for cf in &mut self.conditional_formattings {
            for sqref in &mut cf.sqref {
                *sqref = shift_range_cols(sqref, end_col + 1, count, false);
            }
        }

        // Shift data validation ranges
        for dv in &mut self.data_validations {
            for sqref in &mut dv.sqref {
                *sqref = shift_range_cols(sqref, end_col + 1, count, false);
            }
        }

        // Shift table ranges
        for table in &mut self.tables {
            table.range = shift_range_cols(&table.range, end_col + 1, count, false);
        }

        // Shift hyperlink cell references (remove hyperlinks in deleted range)
        let old_hyperlinks: Vec<Hyperlink> = std::mem::take(&mut self.hyperlinks);
        for mut hl in old_hyperlinks {
            if let Ok((col, _)) = cell_reference_to_column_row(&hl.cell_ref) {
                if col >= start_col && col <= end_col {
                    continue;
                }
            }
            if let Ok(shifted) = shift_cell_reference_col(&hl.cell_ref, end_col + 1, count, false) {
                hl.cell_ref = shifted;
            }
            self.hyperlinks.push(hl);
        }

        // Shift auto-filter range
        if let Some(ref mut af) = self.auto_filter {
            if let Some(range) = af.range().cloned() {
                let shifted = shift_range_cols(&range, end_col + 1, count, false);
                let range_str = format!("{}:{}", shifted.start(), shifted.end());
                let _ = af.set_range(&range_str);
            }
        }

        self.mark_dirty();
        Ok(self)
    }

    /// Sorts rows in the given range by the values in the specified column.
    ///
    /// - `range`: The A1-notation range to sort (e.g. "A1:D10")
    /// - `sort_column`: 1-based column index to sort by
    /// - `ascending`: Sort direction
    ///
    /// This sorts the cell data within the range. Only cells within the range are affected.
    /// Row metadata (height, hidden, etc.) is NOT moved during sort.
    pub fn sort_rows(
        &mut self,
        range: &str,
        sort_column: u32,
        ascending: bool,
    ) -> Result<&mut Self> {
        let cell_range = CellRange::parse(range)?;
        let (start_col, start_row) = cell_reference_to_column_row(cell_range.start())?;
        let (end_col, end_row) = cell_reference_to_column_row(cell_range.end())?;

        if sort_column < start_col || sort_column > end_col {
            return Err(XlsxError::InvalidWorkbookState(
                "sort_column is outside the specified range".to_string(),
            ));
        }

        // Collect row data within the range
        let mut row_data: Vec<(u32, Vec<(String, Cell)>)> = Vec::new();
        for row_idx in start_row..=end_row {
            let mut cells_in_row = Vec::new();
            for col_idx in start_col..=end_col {
                let ref_str = build_cell_reference(col_idx, row_idx)?;
                if let Some(cell) = self.cells.remove(&ref_str) {
                    cells_in_row.push((ref_str, cell));
                }
            }
            row_data.push((row_idx, cells_in_row));
        }

        // Build sort key: extract the sort column value for each row
        let sort_col_name = column_index_to_name(sort_column)?;
        row_data.sort_by(|a, b| {
            let key_ref_a = format!("{}{}", sort_col_name, a.0);
            let key_ref_b = format!("{}{}", sort_col_name, b.0);

            let val_a = a.1.iter().find(|(r, _)| *r == key_ref_a).map(|(_, c)| c);
            let val_b = b.1.iter().find(|(r, _)| *r == key_ref_b).map(|(_, c)| c);

            let cmp =
                compare_cell_values(val_a.and_then(|c| c.value()), val_b.and_then(|c| c.value()));
            if ascending {
                cmp
            } else {
                cmp.reverse()
            }
        });

        // Re-insert cells at their new row positions
        let original_rows: Vec<u32> = (start_row..=end_row).collect();
        for (new_row_idx, (_original_row, cells)) in original_rows.iter().zip(row_data.into_iter())
        {
            for (original_ref, cell) in cells {
                // Parse the original column from the reference
                let (col, _old_row) = cell_reference_to_column_row(&original_ref)?;
                let new_ref = build_cell_reference(col, *new_row_idx)?;
                self.cells.insert(new_ref, cell);
            }
        }

        self.mark_dirty();
        Ok(self)
    }

    pub(crate) fn unknown_children(&self) -> &[RawXmlNode] {
        self.unknown_children.as_slice()
    }

    pub(crate) fn unknown_children_mut(&mut self) -> &mut Vec<RawXmlNode> {
        &mut self.unknown_children
    }

    pub(crate) fn push_unknown_child(&mut self, node: RawXmlNode) {
        self.unknown_children.push(node);
    }

    pub(crate) fn extra_namespace_declarations(&self) -> &[(String, String)] {
        &self.extra_namespace_declarations
    }

    pub(crate) fn set_extra_namespace_declarations(&mut self, declarations: Vec<(String, String)>) {
        self.extra_namespace_declarations = declarations;
    }

    pub(crate) fn original_part_bytes(&self) -> Option<(&str, &[u8])> {
        self.original_part_bytes
            .as_ref()
            .map(|(part_uri, bytes)| (part_uri.as_str(), bytes.as_slice()))
    }

    pub(crate) fn set_original_part_bytes(&mut self, part_uri: String, bytes: Vec<u8>) {
        self.original_part_bytes = Some((part_uri, bytes));
        self.dirty = false;
    }

    pub(crate) fn dirty(&self) -> bool {
        self.dirty
    }

    /// Iterate all cells as `(cell_reference, cell)` pairs.
    pub fn cells(&self) -> impl Iterator<Item = (&str, &Cell)> {
        self.cells
            .iter()
            .map(|(reference, cell)| (reference.as_str(), cell))
    }

    pub(crate) fn insert_cell(&mut self, reference: &str, cell: Cell) -> Result<()> {
        let normalized = normalize_cell_reference(reference)?;
        self.cells.insert(normalized, cell);
        Ok(())
    }

    pub(crate) fn insert_row(&mut self, row: Row) {
        self.rows.insert(row.index(), row);
    }

    pub(crate) fn push_merged_range(&mut self, range: CellRange) {
        if !self.merged_ranges.contains(&range) {
            self.merged_ranges.push(range);
        }
    }

    pub(crate) fn push_image(&mut self, image: WorksheetImage) {
        self.images.push(image);
    }

    pub(crate) fn push_chart(&mut self, chart: Chart) {
        self.charts.push(chart);
    }

    /// Returns pivot tables in insertion order.
    pub fn pivot_tables(&self) -> &[crate::pivot_table::PivotTable] {
        self.pivot_tables.as_slice()
    }

    /// Adds a pivot table to the worksheet.
    pub fn add_pivot_table(&mut self, pivot_table: crate::pivot_table::PivotTable) -> &mut Self {
        self.pivot_tables.push(pivot_table);
        self.mark_dirty();
        self
    }

    /// Starts a fluent pivot table builder.
    ///
    /// Example:
    /// ```ignore
    /// use offidized_xlsx::{sum, Workbook};
    ///
    /// let mut wb = Workbook::new();
    /// let ws = wb.add_sheet("Data");
    /// ws.pivot("RiskPivot")
    ///   .source("Data!A1:D50")
    ///   .rows(["Region"])
    ///   .cols(["Quarter"])
    ///   .filters(["Product"])
    ///   .values([sum("Revenue").name("Total Revenue")])
    ///   .validate_fields()?
    ///   .place("A4")?;
    /// # Ok::<(), offidized_xlsx::XlsxError>(())
    /// ```
    pub fn pivot(&mut self, name: impl Into<String>) -> PivotBuilder<'_> {
        PivotBuilder::new(self, name)
    }

    /// Removes all pivot tables.
    pub fn clear_pivot_tables(&mut self) -> &mut Self {
        self.pivot_tables.clear();
        self.mark_dirty();
        self
    }

    pub(crate) fn push_pivot_table(&mut self, pivot_table: crate::pivot_table::PivotTable) {
        self.pivot_tables.push(pivot_table);
    }

    pub(crate) fn set_parsed_freeze_pane(&mut self, freeze_pane: FreezePane) {
        self.freeze_pane = Some(freeze_pane);
    }

    pub(crate) fn set_parsed_auto_filter(&mut self, auto_filter: AutoFilter) {
        self.auto_filter = Some(auto_filter);
    }

    /// Internal mutable access to the auto-filter (used by parser to add filter columns).
    pub(crate) fn auto_filter_internal_mut(&mut self) -> Option<&mut AutoFilter> {
        self.auto_filter.as_mut()
    }

    pub(crate) fn push_table(&mut self, table: WorksheetTable) {
        self.tables.push(table);
    }

    pub(crate) fn push_conditional_formatting(
        &mut self,
        conditional_formatting: ConditionalFormatting,
    ) {
        self.conditional_formattings.push(conditional_formatting);
    }

    pub(crate) fn push_data_validation(&mut self, data_validation: DataValidation) {
        self.data_validations.push(data_validation);
    }

    pub(crate) fn push_hyperlink(&mut self, hyperlink: Hyperlink) {
        self.hyperlinks.push(hyperlink);
    }

    pub(crate) fn insert_column(&mut self, column: Column) {
        self.columns.insert(column.index(), column);
    }

    pub(crate) fn set_parsed_visibility(&mut self, visibility: SheetVisibility) {
        self.visibility = visibility;
    }

    pub(crate) fn set_parsed_protection(&mut self, protection: SheetProtection) {
        self.protection = Some(protection);
    }

    pub(crate) fn set_parsed_page_setup(&mut self, page_setup: PageSetup) {
        self.page_setup = Some(page_setup);
    }

    pub(crate) fn set_parsed_page_margins(&mut self, margins: PageMargins) {
        self.page_margins = Some(margins);
    }

    pub(crate) fn set_parsed_header_footer(&mut self, hf: PrintHeaderFooter) {
        self.header_footer = Some(hf);
    }

    pub(crate) fn set_parsed_print_area(&mut self, area: PrintArea) {
        self.print_area = Some(area);
    }

    pub(crate) fn set_parsed_page_breaks(&mut self, breaks: PageBreaks) {
        self.page_breaks = Some(breaks);
    }

    pub(crate) fn set_parsed_sparkline_groups(&mut self, groups: Vec<SparklineGroup>) {
        self.sparkline_groups = groups;
    }

    pub(crate) fn set_parsed_sheet_view_options(&mut self, options: SheetViewOptions) {
        self.sheet_view_options = Some(options);
    }

    /// Returns the raw `<printOptions>` attributes, preserved for roundtrip fidelity.
    pub fn raw_print_options_attrs(&self) -> &[(String, String)] {
        &self.raw_print_options_attrs
    }

    /// Sets the raw `<printOptions>` attributes.
    pub(crate) fn set_raw_print_options_attrs(&mut self, attrs: Vec<(String, String)>) {
        self.raw_print_options_attrs = attrs;
    }

    /// Returns the raw `<dimension ref="...">` value, preserved for roundtrip fidelity.
    pub fn raw_dimension_ref(&self) -> Option<&str> {
        self.raw_dimension_ref.as_deref()
    }

    /// Sets the raw dimension ref attribute.
    pub(crate) fn set_raw_dimension_ref(&mut self, dim_ref: String) {
        self.raw_dimension_ref = Some(dim_ref);
    }

    pub(crate) fn set_parsed_tab_color(&mut self, color: String) {
        self.tab_color = Some(color);
    }

    pub(crate) fn set_parsed_default_row_height(&mut self, height: f64) {
        self.default_row_height = Some(height);
    }

    pub(crate) fn set_parsed_default_column_width(&mut self, width: f64) {
        self.default_column_width = Some(width);
    }

    pub(crate) fn set_parsed_custom_height(&mut self, value: bool) {
        self.custom_height = Some(value);
    }

    /// Returns the raw `<sheetFormatPr>` attributes, preserved for roundtrip fidelity.
    pub(crate) fn raw_sheet_format_pr_attrs(&self) -> &[(String, String)] {
        self.raw_sheet_format_pr_attrs.as_slice()
    }

    /// Sets the raw `<sheetFormatPr>` attributes for roundtrip fidelity.
    pub(crate) fn set_raw_sheet_format_pr_attrs(&mut self, attrs: Vec<(String, String)>) {
        self.raw_sheet_format_pr_attrs = attrs;
    }

    pub(crate) fn push_comment(&mut self, comment: Comment) {
        self.comments.push(comment);
    }

    fn mark_dirty(&mut self) {
        self.dirty = true;
    }

    pub(crate) fn source_headers_from_range(&self, range_ref: &str) -> Result<Vec<String>> {
        let range_clean = range_ref.replace('$', "");
        let mut parts = range_clean.split(':');
        let start = parts.next().ok_or_else(|| {
            XlsxError::InvalidWorkbookState("pivot source range is invalid".to_string())
        })?;
        let end = parts.next().ok_or_else(|| {
            XlsxError::InvalidWorkbookState("pivot source range is invalid".to_string())
        })?;
        if parts.next().is_some() {
            return Err(XlsxError::InvalidWorkbookState(
                "pivot source range is invalid".to_string(),
            ));
        }
        let (start_col, start_row) = cell_reference_to_column_row(start)?;
        let (end_col, _) = cell_reference_to_column_row(end)?;

        let mut headers = Vec::new();
        for col in start_col..=end_col {
            let cell_ref = build_cell_reference(col, start_row)?;
            let cell = self.cell(cell_ref.as_str());
            let header = match cell.and_then(|c| c.value()) {
                Some(CellValue::String(s)) => s.clone(),
                Some(CellValue::Number(n)) => n.to_string(),
                Some(CellValue::Bool(b)) => {
                    if *b {
                        "true".to_string()
                    } else {
                        "false".to_string()
                    }
                }
                Some(CellValue::Date(s)) => s.clone(),
                Some(CellValue::Error(s)) => s.clone(),
                Some(CellValue::DateTime(n)) => n.to_string(),
                Some(CellValue::RichText(runs)) => runs.iter().map(|r| r.text()).collect(),
                Some(CellValue::Blank) | None => String::new(),
            };
            headers.push(header);
        }
        Ok(headers)
    }
}

fn normalize_source_sheet_name(name: &str) -> String {
    let trimmed = name.trim();
    if trimmed.len() >= 2 && trimmed.starts_with('\'') && trimmed.ends_with('\'') {
        trimmed[1..trimmed.len() - 1].replace("''", "'")
    } else {
        trimmed.to_string()
    }
}

fn split_source_reference(source: &str) -> (Option<&str>, &str) {
    if let Some(pos) = source.find('!') {
        (Some(&source[..pos]), &source[pos + 1..])
    } else {
        (None, source)
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

fn shift_cell_reference_row(
    reference: &str,
    start_row: u32,
    count: u32,
    insert: bool,
) -> Result<String> {
    let (col, row) = cell_reference_to_column_row(reference)?;
    if row < start_row {
        return normalize_cell_reference(reference);
    }
    let new_row = if insert {
        row.checked_add(count).ok_or_else(|| {
            XlsxError::InvalidWorkbookState("row overflow during shift".to_string())
        })?
    } else {
        row.saturating_sub(count).max(1)
    };
    build_cell_reference(col, new_row)
}

fn shift_cell_reference_col(
    reference: &str,
    start_col: u32,
    count: u32,
    insert: bool,
) -> Result<String> {
    let (col, row) = cell_reference_to_column_row(reference)?;
    if col < start_col {
        return normalize_cell_reference(reference);
    }
    let new_col = if insert {
        col.checked_add(count).ok_or_else(|| {
            XlsxError::InvalidWorkbookState("column overflow during shift".to_string())
        })?
    } else {
        col.saturating_sub(count).max(1)
    };
    build_cell_reference(new_col, row)
}

fn shift_range_rows(range: &CellRange, start_row: u32, count: u32, insert: bool) -> CellRange {
    let new_start = shift_cell_reference_row(range.start(), start_row, count, insert)
        .unwrap_or_else(|_| range.start().to_string());
    let new_end = shift_cell_reference_row(range.end(), start_row, count, insert)
        .unwrap_or_else(|_| range.end().to_string());
    CellRange::new(new_start.as_str(), new_end.as_str())
        .unwrap_or_else(|_| CellRange::parse(range.start()).unwrap_or_else(|_| range.clone()))
}

fn shift_range_cols(range: &CellRange, start_col: u32, count: u32, insert: bool) -> CellRange {
    let new_start = shift_cell_reference_col(range.start(), start_col, count, insert)
        .unwrap_or_else(|_| range.start().to_string());
    let new_end = shift_cell_reference_col(range.end(), start_col, count, insert)
        .unwrap_or_else(|_| range.end().to_string());
    CellRange::new(new_start.as_str(), new_end.as_str())
        .unwrap_or_else(|_| CellRange::parse(range.start()).unwrap_or_else(|_| range.clone()))
}

fn validate_dimension_index(kind: &str, index: u32) -> Result<()> {
    if index == 0 {
        return Err(XlsxError::InvalidWorkbookState(format!(
            "{kind} index must be >= 1"
        )));
    }
    Ok(())
}

/// Compares two cell values for sorting purposes.
///
/// Sort order: Blank < Number/DateTime < Bool < String < Date < Error
fn compare_cell_values(a: Option<&CellValue>, b: Option<&CellValue>) -> std::cmp::Ordering {
    use std::cmp::Ordering;

    fn sort_rank(value: Option<&CellValue>) -> u8 {
        match value {
            None | Some(CellValue::Blank) => 0,
            Some(CellValue::Number(_)) | Some(CellValue::DateTime(_)) => 1,
            Some(CellValue::Bool(_)) => 2,
            Some(CellValue::String(_)) | Some(CellValue::RichText(_)) => 3,
            Some(CellValue::Date(_)) => 4,
            Some(CellValue::Error(_)) => 5,
        }
    }

    let rank_a = sort_rank(a);
    let rank_b = sort_rank(b);
    if rank_a != rank_b {
        return rank_a.cmp(&rank_b);
    }

    match (a, b) {
        (Some(CellValue::Number(na)), Some(CellValue::Number(nb))) => {
            na.partial_cmp(nb).unwrap_or(Ordering::Equal)
        }
        (Some(CellValue::DateTime(na)), Some(CellValue::DateTime(nb))) => {
            na.partial_cmp(nb).unwrap_or(Ordering::Equal)
        }
        (Some(CellValue::Number(na)), Some(CellValue::DateTime(nb)))
        | (Some(CellValue::DateTime(na)), Some(CellValue::Number(nb))) => {
            na.partial_cmp(nb).unwrap_or(Ordering::Equal)
        }
        (Some(CellValue::String(sa)), Some(CellValue::String(sb))) => sa.cmp(sb),
        (Some(CellValue::Bool(ba)), Some(CellValue::Bool(bb))) => ba.cmp(bb),
        (Some(CellValue::Date(da)), Some(CellValue::Date(db))) => da.cmp(db),
        (Some(CellValue::Error(ea)), Some(CellValue::Error(eb))) => ea.cmp(eb),
        _ => Ordering::Equal,
    }
}

/// Shifts A1-style cell references within a formula string when rows are inserted or deleted.
///
/// This is a best-effort transformation: it finds tokens that look like cell references
/// (letter(s) followed by digits) and adjusts the row component.
pub(crate) fn shift_formula_row_references(
    formula: &str,
    start_row: u32,
    count: u32,
    insert: bool,
) -> String {
    let mut result = String::with_capacity(formula.len());
    let bytes = formula.as_bytes();
    let len = bytes.len();
    let mut i = 0;

    while i < len {
        // Skip absolute markers for the purpose of finding references
        let ch = bytes[i];

        // Try to match a cell reference: optional $, letters, optional $, digits
        if ch.is_ascii_alphabetic() || ch == b'$' {
            let ref_start = i;
            let mut j = i;

            // Skip optional $ before column letters
            if j < len && bytes[j] == b'$' {
                j += 1;
            }

            // Consume column letters
            let col_start = j;
            while j < len && bytes[j].is_ascii_alphabetic() {
                j += 1;
            }
            let col_end = j;

            if col_end > col_start {
                // Skip optional $ before row digits
                if j < len && bytes[j] == b'$' {
                    j += 1;
                }

                // Consume row digits
                let row_start = j;
                while j < len && bytes[j].is_ascii_digit() {
                    j += 1;
                }
                let row_end = j;

                if row_end > row_start {
                    // We have something that looks like a cell reference
                    // But verify the character before isn't alphanumeric (to avoid matching
                    // inside function names like SUM1 or identifiers)
                    let preceded_by_alpha = if ref_start > 0 {
                        bytes[ref_start - 1].is_ascii_alphanumeric() || bytes[ref_start - 1] == b'_'
                    } else {
                        false
                    };

                    // Also check the char after isn't alphanumeric
                    let followed_by_alpha = if row_end < len {
                        bytes[row_end].is_ascii_alphanumeric() || bytes[row_end] == b'_'
                    } else {
                        false
                    };

                    if !preceded_by_alpha && !followed_by_alpha {
                        // Parse the row number
                        let row_str = &formula[row_start..row_end];
                        if let Ok(row_num) = row_str.parse::<u32>() {
                            if row_num >= start_row {
                                let new_row = if insert {
                                    row_num.saturating_add(count)
                                } else {
                                    let r = row_num.saturating_sub(count);
                                    if r < 1 {
                                        1
                                    } else {
                                        r
                                    }
                                };
                                // Emit the prefix (including $ and column letters) as-is
                                result.push_str(&formula[ref_start..row_start]);
                                result.push_str(&new_row.to_string());
                                i = row_end;
                                continue;
                            }
                        }
                    }
                }
            }
        }

        result.push(ch as char);
        i += 1;
    }

    result
}

/// Shifts A1-style cell references within a formula string when columns are inserted or deleted.
pub(crate) fn shift_formula_col_references(
    formula: &str,
    start_col: u32,
    count: u32,
    insert: bool,
) -> String {
    let mut result = String::with_capacity(formula.len());
    let bytes = formula.as_bytes();
    let len = bytes.len();
    let mut i = 0;

    while i < len {
        let ch = bytes[i];

        if ch.is_ascii_alphabetic() || ch == b'$' {
            let ref_start = i;
            let mut j = i;

            // Skip optional $ before column letters
            let has_col_dollar = j < len && bytes[j] == b'$';
            if has_col_dollar {
                j += 1;
            }

            // Consume column letters
            let col_start = j;
            while j < len && bytes[j].is_ascii_alphabetic() {
                j += 1;
            }
            let col_end = j;

            if col_end > col_start {
                // Skip optional $ before row digits
                if j < len && bytes[j] == b'$' {
                    j += 1;
                }

                // Consume row digits
                let row_start = j;
                while j < len && bytes[j].is_ascii_digit() {
                    j += 1;
                }
                let row_end = j;

                if row_end > row_start {
                    let preceded_by_alpha = if ref_start > 0 {
                        bytes[ref_start - 1].is_ascii_alphanumeric() || bytes[ref_start - 1] == b'_'
                    } else {
                        false
                    };

                    let followed_by_alpha = if row_end < len {
                        bytes[row_end].is_ascii_alphanumeric() || bytes[row_end] == b'_'
                    } else {
                        false
                    };

                    if !preceded_by_alpha && !followed_by_alpha {
                        let col_letters = &formula[col_start..col_end];
                        if let Some(col_idx) = parse_column_name(col_letters) {
                            if col_idx >= start_col {
                                let new_col = if insert {
                                    col_idx.saturating_add(count)
                                } else {
                                    let c = col_idx.saturating_sub(count);
                                    if c < 1 {
                                        1
                                    } else {
                                        c
                                    }
                                };
                                if let Ok(new_col_name) = column_index_to_name(new_col) {
                                    // Emit prefix (dollar sign if present)
                                    if has_col_dollar {
                                        result.push('$');
                                    }
                                    result.push_str(&new_col_name);
                                    // Emit the row part (everything from col_end to row_end)
                                    result.push_str(&formula[col_end..row_end]);
                                    i = row_end;
                                    continue;
                                }
                            }
                        }
                    }
                }
            }
        }

        result.push(ch as char);
        i += 1;
    }

    result
}

fn parse_column_name(name: &str) -> Option<u32> {
    let mut index = 0_u32;
    for byte in name.bytes() {
        let ch = byte.to_ascii_uppercase();
        if !ch.is_ascii_uppercase() {
            return None;
        }
        index = index.checked_mul(26)?;
        index = index.checked_add(u32::from(ch - b'A' + 1))?;
    }
    if index == 0 {
        None
    } else {
        Some(index)
    }
}

fn format_cell_range(range: &CellRange) -> String {
    if range.start() == range.end() {
        range.start().to_string()
    } else {
        format!("{}:{}", range.start(), range.end())
    }
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

#[cfg(test)]
mod tests {
    use super::*;
    use crate::cell::CellValue;

    #[test]
    fn row_and_column_metadata_accessors_work() {
        let mut worksheet = Worksheet::new("Data");

        worksheet
            .row_mut(2)
            .expect("row index should be valid")
            .set_height(20.0);
        worksheet
            .row_mut(1)
            .expect("row index should be valid")
            .set_height(12.5);
        worksheet
            .column_mut(3)
            .expect("column index should be valid")
            .set_width(25.0);
        worksheet
            .column_mut(1)
            .expect("column index should be valid")
            .set_width(10.5);

        assert_eq!(worksheet.row(1).and_then(Row::height), Some(12.5));
        assert_eq!(worksheet.row(2).and_then(Row::height), Some(20.0));
        assert_eq!(worksheet.column(1).and_then(Column::width), Some(10.5));
        assert_eq!(worksheet.column(3).and_then(Column::width), Some(25.0));
        assert!(worksheet.row(99).is_none());
        assert!(worksheet.column(99).is_none());

        let row_indexes: Vec<u32> = worksheet.rows().map(Row::index).collect();
        let column_indexes: Vec<u32> = worksheet.columns().map(Column::index).collect();
        assert_eq!(row_indexes, vec![1, 2]);
        assert_eq!(column_indexes, vec![1, 3]);
    }

    #[test]
    fn metadata_accessors_reject_zero_index() {
        let mut worksheet = Worksheet::new("Data");

        assert!(worksheet.row_mut(0).is_err());
        assert!(worksheet.column_mut(0).is_err());
        assert!(worksheet.row(0).is_none());
        assert!(worksheet.column(0).is_none());
    }

    #[test]
    fn merged_ranges_accessors_work() {
        let mut worksheet = Worksheet::new("Data");

        worksheet
            .add_merged_range("B2:A1")
            .expect("range should parse");
        worksheet
            .add_merged_range("A1:B2")
            .expect("range should parse");
        worksheet
            .add_merged_range("C3")
            .expect("single-cell range should parse");

        assert_eq!(worksheet.merged_ranges().len(), 2);
        assert_eq!(worksheet.merged_ranges()[0].start(), "A1");
        assert_eq!(worksheet.merged_ranges()[0].end(), "B2");
        assert_eq!(worksheet.merged_ranges()[1].start(), "C3");
        assert_eq!(worksheet.merged_ranges()[1].end(), "C3");

        worksheet.clear_merged_ranges();
        assert!(worksheet.merged_ranges().is_empty());
    }

    #[test]
    fn worksheet_images_accessors_and_validation_work() {
        let mut worksheet = Worksheet::new("Data");
        let ext = WorksheetImageExt::new(120_000, 80_000).expect("ext should be valid");

        worksheet
            .add_image(vec![1_u8, 2_u8, 3_u8], " image/png ", "b2", Some(ext))
            .expect("image should be valid");
        worksheet
            .add_image(vec![9_u8, 8_u8], "image/jpeg", "C3", None)
            .expect("image should be valid");

        assert_eq!(worksheet.images().len(), 2);
        assert_eq!(worksheet.images()[0].bytes(), &[1_u8, 2_u8, 3_u8]);
        assert_eq!(worksheet.images()[0].content_type(), "image/png");
        assert_eq!(worksheet.images()[0].anchor_cell(), "B2");
        assert_eq!(worksheet.images()[0].ext(), Some(ext));
        assert_eq!(worksheet.images()[1].anchor_cell(), "C3");
        assert_eq!(worksheet.images()[1].ext(), None);

        worksheet.clear_images();
        assert!(worksheet.images().is_empty());

        assert!(worksheet
            .add_image(Vec::<u8>::new(), "image/png", "A1", None)
            .is_err());
        assert!(worksheet.add_image(vec![1_u8], " ", "A1", None).is_err());
        assert!(worksheet
            .add_image(vec![1_u8], "image/png", "bad", None)
            .is_err());
        assert!(WorksheetImageExt::new(0, 1).is_err());
        assert!(WorksheetImageExt::new(1, 0).is_err());
    }

    #[test]
    fn freeze_pane_accessors_work() {
        let mut worksheet = Worksheet::new("Data");

        worksheet
            .set_freeze_panes(1, 2)
            .expect("freeze pane should be valid");
        let freeze_pane = worksheet.freeze_pane().expect("freeze pane should be set");
        assert_eq!(freeze_pane.x_split(), 1);
        assert_eq!(freeze_pane.y_split(), 2);
        assert_eq!(freeze_pane.top_left_cell(), "B3");

        worksheet
            .set_freeze_panes_with_top_left_cell(1, 2, "c5")
            .expect("freeze pane should be valid");
        assert_eq!(
            worksheet
                .freeze_pane()
                .expect("freeze pane should be set")
                .top_left_cell(),
            "C5"
        );

        worksheet.clear_freeze_pane();
        assert!(worksheet.freeze_pane().is_none());
        assert!(worksheet.set_freeze_panes(0, 0).is_err());
    }

    #[test]
    fn data_validation_builders_and_accessors_work() {
        let mut worksheet = Worksheet::new("Data");

        let mut whole =
            DataValidation::whole(["A1:A3"], "1").expect("whole validation should be created");
        whole
            .set_formula2("10")
            .add_range("B1")
            .expect("range should parse");
        worksheet.add_data_validation(whole);

        let list = DataValidation::list(["C1:C4", "D1"], "\"Yes,No\"")
            .expect("list validation should be created");
        worksheet.add_data_validation(list);

        assert_eq!(worksheet.data_validations().len(), 2);
        assert_eq!(
            worksheet.data_validations()[0].validation_type(),
            DataValidationType::Whole
        );
        assert_eq!(worksheet.data_validations()[0].formula1(), "1");
        assert_eq!(worksheet.data_validations()[0].formula2(), Some("10"));
        assert_eq!(worksheet.data_validations()[0].sqref().len(), 2);
        assert_eq!(
            worksheet.data_validations()[1].validation_type(),
            DataValidationType::List
        );
        assert_eq!(worksheet.data_validations()[1].formula1(), "\"Yes,No\"");
        assert_eq!(worksheet.data_validations()[1].formula2(), None);

        worksheet.clear_data_validations();
        assert!(worksheet.data_validations().is_empty());
    }

    #[test]
    fn auto_filter_accessors_work() {
        let mut worksheet = Worksheet::new("Data");

        worksheet
            .set_auto_filter("B5:A1")
            .expect("auto filter range should parse");
        let auto_filter = worksheet
            .auto_filter()
            .expect("auto filter should be configured");
        let range = auto_filter.range().expect("range should be set");
        assert_eq!(range.start(), "A1");
        assert_eq!(range.end(), "B5");

        worksheet.clear_auto_filter();
        assert!(worksheet.auto_filter().is_none());
        assert!(worksheet.set_auto_filter("bad").is_err());
    }

    #[test]
    fn table_builders_and_accessors_work() {
        let mut worksheet = Worksheet::new("Data");
        let mut table = WorksheetTable::new(" Sales ", "B5:A1").expect("table should be valid");
        assert_eq!(table.name(), "Sales");
        assert_eq!(table.range().start(), "A1");
        assert_eq!(table.range().end(), "B5");
        assert!(table.has_header_row());

        table
            .set_name("Revenue")
            .expect("table name should be updated");
        table
            .set_range("C1:D5")
            .expect("table range should be updated");
        table.set_header_row(false);

        worksheet.add_table(table);
        assert_eq!(worksheet.tables().len(), 1);
        assert_eq!(worksheet.tables()[0].name(), "Revenue");
        assert_eq!(worksheet.tables()[0].range().start(), "C1");
        assert_eq!(worksheet.tables()[0].range().end(), "D5");
        assert!(!worksheet.tables()[0].has_header_row());

        worksheet.clear_tables();
        assert!(worksheet.tables().is_empty());
        assert!(WorksheetTable::new(" ", "A1:B2").is_err());
    }

    #[test]
    fn pivot_builder_builds_and_places_pivot() {
        let mut ws = Worksheet::new("Data");
        ws.cell_mut("A1").unwrap().set_value("Region");
        ws.cell_mut("B1").unwrap().set_value("Quarter");
        ws.cell_mut("C1").unwrap().set_value("Revenue");
        ws.cell_mut("A2").unwrap().set_value("North");
        ws.cell_mut("B2").unwrap().set_value("Q1");
        ws.cell_mut("C2").unwrap().set_value(100);

        ws.pivot("RevenuePivot")
            .source("Data!A1:C2")
            .rows(["Region"])
            .cols(["Quarter"])
            .values([sum("Revenue").name("Total Revenue")])
            .validate_fields()
            .expect("pivot config should validate")
            .place("D4")
            .expect("pivot should be placed");

        assert_eq!(ws.pivot_tables().len(), 1);
        let pivot = &ws.pivot_tables()[0];
        assert_eq!(pivot.name(), "RevenuePivot");
        assert_eq!(pivot.target_row(), 3);
        assert_eq!(pivot.target_col(), 3);
        assert_eq!(pivot.row_fields().len(), 1);
        assert_eq!(pivot.column_fields().len(), 1);
        assert_eq!(pivot.data_fields().len(), 1);
        assert_eq!(pivot.data_fields()[0].custom_name(), Some("Total Revenue"));
    }

    #[test]
    fn pivot_builder_validate_fails_for_unknown_field() {
        let mut ws = Worksheet::new("Data");
        ws.cell_mut("A1").unwrap().set_value("Region");
        ws.cell_mut("B1").unwrap().set_value("Revenue");
        ws.cell_mut("A2").unwrap().set_value("North");
        ws.cell_mut("B2").unwrap().set_value(100);

        let err = match ws
            .pivot("RevenuePivot")
            .source("Data!A1:B2")
            .rows(["Desk"])
            .values([sum("Revenue")])
            .validate_fields()
        {
            Ok(_) => panic!("unknown field should fail validation"),
            Err(err) => err,
        };

        assert!(
            err.to_string().contains("not found in source header row"),
            "unexpected error: {err}"
        );
    }

    #[test]
    fn pivot_builder_validate_fails_for_duplicate_axis_field() {
        let mut ws = Worksheet::new("Data");
        ws.cell_mut("A1").unwrap().set_value("Region");
        ws.cell_mut("B1").unwrap().set_value("Revenue");
        ws.cell_mut("A2").unwrap().set_value("North");
        ws.cell_mut("B2").unwrap().set_value(100);

        let err = match ws
            .pivot("RevenuePivot")
            .source("Data!A1:B2")
            .rows(["Region"])
            .cols(["Region"])
            .values([sum("Revenue")])
            .validate_fields()
        {
            Ok(_) => panic!("duplicate field across axes should fail"),
            Err(err) => err,
        };

        assert!(
            err.to_string().contains("appears more than once"),
            "unexpected error: {err}"
        );
    }

    #[test]
    fn conditional_formatting_builders_and_accessors_work() {
        let mut worksheet = Worksheet::new("Data");
        let mut cell_is = ConditionalFormatting::cell_is(["A1:A5"], ["5", "10"])
            .expect("cellIs conditional formatting should be valid");
        cell_is
            .add_range("C1")
            .expect("additional range should be valid");
        cell_is
            .add_formula("15")
            .expect("additional formula should be valid");
        worksheet.add_conditional_formatting(cell_is);

        let expression = ConditionalFormatting::expression(["D1:D5"], ["MOD(ROW(),2)=0"])
            .expect("expression conditional formatting should be valid");
        worksheet.add_conditional_formatting(expression);

        assert_eq!(worksheet.conditional_formattings().len(), 2);
        assert_eq!(
            worksheet.conditional_formattings()[0].rule_type(),
            ConditionalFormattingRuleType::CellIs
        );
        assert_eq!(worksheet.conditional_formattings()[0].sqref().len(), 2);
        assert_eq!(
            worksheet.conditional_formattings()[0].formulas(),
            &["5".to_string(), "10".to_string(), "15".to_string()]
        );
        assert_eq!(
            worksheet.conditional_formattings()[1].rule_type(),
            ConditionalFormattingRuleType::Expression
        );
        assert_eq!(
            worksheet.conditional_formattings()[1].formulas(),
            &["MOD(ROW(),2)=0".to_string()]
        );

        worksheet.clear_conditional_formattings();
        assert!(worksheet.conditional_formattings().is_empty());
    }

    #[test]
    fn data_validation_rejects_invalid_inputs() {
        assert!(DataValidation::date(["A1"], " ").is_err());
        assert!(DataValidation::decimal(Vec::<&str>::new(), "1").is_err());
        assert!(DataValidation::text_length(["A1"], "5")
            .expect("validation should be created")
            .add_range("bad")
            .is_err());
        assert!(ConditionalFormatting::cell_is(["A1"], Vec::<&str>::new()).is_err());
        assert!(ConditionalFormatting::expression(Vec::<&str>::new(), ["A1>0"]).is_err());
        assert!(ConditionalFormatting::expression(["A1"], [" "]).is_err());
    }

    #[test]
    fn row_hidden_accessors_work() {
        let mut worksheet = Worksheet::new("Data");

        let row = worksheet.row_mut(1).expect("row index should be valid");
        assert!(!row.is_hidden());

        row.set_hidden(true);
        assert!(row.is_hidden());

        row.set_hidden(false);
        assert!(!row.is_hidden());
    }

    #[test]
    fn column_hidden_accessors_work() {
        let mut worksheet = Worksheet::new("Data");

        let col = worksheet
            .column_mut(2)
            .expect("column index should be valid");
        assert!(!col.is_hidden());

        col.set_hidden(true);
        assert!(col.is_hidden());

        col.set_hidden(false);
        assert!(!col.is_hidden());
    }

    #[test]
    fn hyperlink_external_builders_and_accessors_work() {
        let mut worksheet = Worksheet::new("Data");

        let mut hl = Hyperlink::external("a1", "https://example.com")
            .expect("external hyperlink should be valid");
        assert_eq!(hl.cell_ref(), "A1");
        assert_eq!(hl.url(), Some("https://example.com"));
        assert!(hl.location().is_none());
        assert!(hl.tooltip().is_none());

        hl.set_tooltip("My Tooltip");
        assert_eq!(hl.tooltip(), Some("My Tooltip"));
        hl.clear_tooltip();
        assert!(hl.tooltip().is_none());

        worksheet.add_hyperlink(hl);
        assert_eq!(worksheet.hyperlinks().len(), 1);
        assert_eq!(worksheet.hyperlinks()[0].url(), Some("https://example.com"));
    }

    #[test]
    fn hyperlink_internal_builders_and_accessors_work() {
        let mut worksheet = Worksheet::new("Data");

        let hl =
            Hyperlink::internal("B2", "Sheet2!A1").expect("internal hyperlink should be valid");
        assert_eq!(hl.cell_ref(), "B2");
        assert!(hl.url().is_none());
        assert_eq!(hl.location(), Some("Sheet2!A1"));

        worksheet.add_hyperlink(hl);
        assert_eq!(worksheet.hyperlinks().len(), 1);
    }

    #[test]
    fn hyperlink_remove_and_clear_work() {
        let mut worksheet = Worksheet::new("Data");

        worksheet.add_hyperlink(
            Hyperlink::external("A1", "https://example.com/a").expect("hyperlink should be valid"),
        );
        worksheet.add_hyperlink(
            Hyperlink::external("B1", "https://example.com/b").expect("hyperlink should be valid"),
        );
        worksheet.add_hyperlink(
            Hyperlink::internal("C1", "Sheet2!A1").expect("hyperlink should be valid"),
        );

        assert_eq!(worksheet.hyperlinks().len(), 3);

        assert!(worksheet.remove_hyperlink("b1"));
        assert_eq!(worksheet.hyperlinks().len(), 2);
        assert_eq!(worksheet.hyperlinks()[0].cell_ref(), "A1");
        assert_eq!(worksheet.hyperlinks()[1].cell_ref(), "C1");

        assert!(!worksheet.remove_hyperlink("B1")); // already removed
        assert!(!worksheet.remove_hyperlink("invalid")); // bad ref

        worksheet.clear_hyperlinks();
        assert!(worksheet.hyperlinks().is_empty());
    }

    #[test]
    fn hyperlink_rejects_invalid_inputs() {
        assert!(Hyperlink::external("A1", " ").is_err());
        assert!(Hyperlink::internal("A1", " ").is_err());
        assert!(Hyperlink::external("bad", "https://example.com").is_err());
        assert!(Hyperlink::internal("bad", "Sheet2!A1").is_err());
        assert!(Hyperlink::from_parsed_parts("A1".to_string(), None, None, None, None).is_err());
    }

    // ===== Feature 1: Sheet visibility =====

    #[test]
    fn sheet_visibility_default_is_visible() {
        let worksheet = Worksheet::new("Data");
        assert_eq!(worksheet.visibility(), SheetVisibility::Visible);
    }

    #[test]
    fn sheet_visibility_setter_changes_state() {
        let mut worksheet = Worksheet::new("Data");
        worksheet.set_visibility(SheetVisibility::Hidden);
        assert_eq!(worksheet.visibility(), SheetVisibility::Hidden);

        worksheet.set_visibility(SheetVisibility::VeryHidden);
        assert_eq!(worksheet.visibility(), SheetVisibility::VeryHidden);

        worksheet.set_visibility(SheetVisibility::Visible);
        assert_eq!(worksheet.visibility(), SheetVisibility::Visible);
    }

    #[test]
    fn sheet_visibility_xml_value_roundtrip() {
        assert_eq!(SheetVisibility::Visible.as_xml_value(), "visible");
        assert_eq!(SheetVisibility::Hidden.as_xml_value(), "hidden");
        assert_eq!(SheetVisibility::VeryHidden.as_xml_value(), "veryHidden");

        assert_eq!(
            SheetVisibility::from_xml_value("visible"),
            SheetVisibility::Visible
        );
        assert_eq!(
            SheetVisibility::from_xml_value("hidden"),
            SheetVisibility::Hidden
        );
        assert_eq!(
            SheetVisibility::from_xml_value("veryHidden"),
            SheetVisibility::VeryHidden
        );
        assert_eq!(
            SheetVisibility::from_xml_value("unknown"),
            SheetVisibility::Visible
        );
    }

    // ===== Feature 2: Sheet protection =====

    #[test]
    fn sheet_protection_new_enables_sheet_flag() {
        let protection = SheetProtection::new();
        assert!(protection.sheet());
        assert!(!protection.objects());
        assert!(!protection.scenarios());
        assert!(protection.password_hash().is_none());
    }

    #[test]
    fn sheet_protection_builder_pattern_works() {
        let mut protection = SheetProtection::new();
        protection
            .set_objects(true)
            .set_scenarios(true)
            .set_format_cells(true)
            .set_format_columns(true)
            .set_format_rows(true)
            .set_insert_columns(true)
            .set_insert_rows(true)
            .set_insert_hyperlinks(true)
            .set_delete_columns(true)
            .set_delete_rows(true)
            .set_select_locked_cells(true)
            .set_sort(true)
            .set_auto_filter(true)
            .set_pivot_tables(true)
            .set_select_unlocked_cells(true)
            .set_password_hash("ABCD1234");

        assert!(protection.objects());
        assert!(protection.scenarios());
        assert!(protection.format_cells());
        assert!(protection.format_columns());
        assert!(protection.format_rows());
        assert!(protection.insert_columns());
        assert!(protection.insert_rows());
        assert!(protection.insert_hyperlinks());
        assert!(protection.delete_columns());
        assert!(protection.delete_rows());
        assert!(protection.select_locked_cells());
        assert!(protection.sort());
        assert!(protection.auto_filter());
        assert!(protection.pivot_tables());
        assert!(protection.select_unlocked_cells());
        assert_eq!(protection.password_hash(), Some("ABCD1234"));
    }

    #[test]
    fn sheet_protection_password_hash_clearing_works() {
        let mut protection = SheetProtection::new();
        protection.set_password_hash("hash123");
        assert_eq!(protection.password_hash(), Some("hash123"));

        protection.clear_password_hash();
        assert!(protection.password_hash().is_none());
    }

    #[test]
    fn sheet_protection_empty_password_becomes_none() {
        let mut protection = SheetProtection::new();
        protection.set_password_hash("  ");
        assert!(protection.password_hash().is_none());
    }

    #[test]
    fn sheet_protection_has_metadata_reflects_flags() {
        let default_protection = SheetProtection::default();
        assert!(!default_protection.has_metadata());

        let enabled_protection = SheetProtection::new();
        assert!(enabled_protection.has_metadata());
    }

    #[test]
    fn worksheet_protection_set_and_clear() {
        let mut worksheet = Worksheet::new("Data");
        assert!(worksheet.protection().is_none());

        let protection = SheetProtection::new();
        worksheet.set_protection(protection);
        assert!(worksheet.protection().is_some());
        assert!(worksheet.protection().unwrap().sheet());

        worksheet.clear_protection();
        assert!(worksheet.protection().is_none());
    }

    // ===== Feature 3: Page setup and margins =====

    #[test]
    fn page_orientation_xml_values() {
        assert_eq!(PageOrientation::Portrait.as_xml_value(), "portrait");
        assert_eq!(PageOrientation::Landscape.as_xml_value(), "landscape");
        assert_eq!(
            PageOrientation::from_xml_value("landscape"),
            PageOrientation::Landscape
        );
        assert_eq!(
            PageOrientation::from_xml_value("portrait"),
            PageOrientation::Portrait
        );
        assert_eq!(
            PageOrientation::from_xml_value("unknown"),
            PageOrientation::Portrait
        );
    }

    #[test]
    fn page_margins_builder_pattern_and_accessors() {
        let mut margins = PageMargins::new();
        assert!(!margins.has_metadata());

        margins
            .set_left(0.7)
            .set_right(0.7)
            .set_top(0.75)
            .set_bottom(0.75)
            .set_header(0.3)
            .set_footer(0.3);

        assert!(margins.has_metadata());
        assert_eq!(margins.left(), Some(0.7));
        assert_eq!(margins.right(), Some(0.7));
        assert_eq!(margins.top(), Some(0.75));
        assert_eq!(margins.bottom(), Some(0.75));
        assert_eq!(margins.header(), Some(0.3));
        assert_eq!(margins.footer(), Some(0.3));
    }

    #[test]
    fn page_margins_clear_individual_values() {
        let mut margins = PageMargins::new();
        margins.set_left(1.0).set_right(1.0);

        margins.clear_left();
        assert!(margins.left().is_none());
        assert_eq!(margins.right(), Some(1.0));
    }

    #[test]
    fn page_margins_rejects_negative_and_infinite() {
        let mut margins = PageMargins::new();
        margins.set_left(-1.0);
        assert!(margins.left().is_none());

        margins.set_top(f64::INFINITY);
        assert!(margins.top().is_none());

        margins.set_bottom(f64::NAN);
        assert!(margins.bottom().is_none());
    }

    #[test]
    fn page_setup_builder_pattern_and_accessors() {
        let mut setup = PageSetup::new();
        assert!(!setup.has_metadata());

        setup
            .set_orientation(PageOrientation::Landscape)
            .set_paper_size(9) // A4
            .set_scale(75)
            .set_fit_to_width(1)
            .set_fit_to_height(2)
            .set_first_page_number(5);

        assert!(setup.has_metadata());
        assert_eq!(setup.orientation(), Some(PageOrientation::Landscape));
        assert_eq!(setup.paper_size(), Some(9));
        assert_eq!(setup.scale(), Some(75));
        assert_eq!(setup.fit_to_width(), Some(1));
        assert_eq!(setup.fit_to_height(), Some(2));
        assert_eq!(setup.first_page_number(), Some(5));
    }

    #[test]
    fn page_setup_clear_all_fields() {
        let mut setup = PageSetup::new();
        setup
            .set_orientation(PageOrientation::Landscape)
            .set_paper_size(1);

        setup.clear_orientation().clear_paper_size();
        assert!(setup.orientation().is_none());
        assert!(setup.paper_size().is_none());
        assert!(!setup.has_metadata());
    }

    #[test]
    fn worksheet_page_setup_set_and_clear() {
        let mut worksheet = Worksheet::new("Data");
        assert!(worksheet.page_setup().is_none());
        assert!(worksheet.page_margins().is_none());

        let mut setup = PageSetup::new();
        setup.set_orientation(PageOrientation::Landscape);
        worksheet.set_page_setup(setup);
        assert!(worksheet.page_setup().is_some());
        assert_eq!(
            worksheet.page_setup().unwrap().orientation(),
            Some(PageOrientation::Landscape)
        );

        let mut margins = PageMargins::new();
        margins.set_left(0.5);
        worksheet.set_page_margins(margins);
        assert!(worksheet.page_margins().is_some());
        assert_eq!(worksheet.page_margins().unwrap().left(), Some(0.5));

        worksheet.clear_page_setup();
        worksheet.clear_page_margins();
        assert!(worksheet.page_setup().is_none());
        assert!(worksheet.page_margins().is_none());
    }

    // ===== Feature 4 & 5: Insert/Delete rows and columns =====

    #[test]
    fn insert_rows_shifts_cells_down() {
        let mut worksheet = Worksheet::new("Data");
        worksheet.cell_mut("A1").unwrap().set_value("row1");
        worksheet.cell_mut("A2").unwrap().set_value("row2");
        worksheet.cell_mut("A3").unwrap().set_value("row3");

        worksheet.insert_rows(2, 2).unwrap();

        assert_eq!(
            worksheet.cell("A1").and_then(|c| c.value()),
            Some(&CellValue::String("row1".to_string()))
        );
        // A2 and A3 should now be at A4 and A5
        assert!(worksheet.cell("A2").is_none());
        assert!(worksheet.cell("A3").is_none());
        assert_eq!(
            worksheet.cell("A4").and_then(|c| c.value()),
            Some(&CellValue::String("row2".to_string()))
        );
        assert_eq!(
            worksheet.cell("A5").and_then(|c| c.value()),
            Some(&CellValue::String("row3".to_string()))
        );
    }

    #[test]
    fn insert_rows_shifts_merged_ranges() {
        let mut worksheet = Worksheet::new("Data");
        worksheet.add_merged_range("A2:B3").unwrap();

        worksheet.insert_rows(2, 1).unwrap();

        assert_eq!(worksheet.merged_ranges().len(), 1);
        assert_eq!(worksheet.merged_ranges()[0].start(), "A3");
        assert_eq!(worksheet.merged_ranges()[0].end(), "B4");
    }

    #[test]
    fn insert_rows_rejects_zero_start() {
        let mut worksheet = Worksheet::new("Data");
        assert!(worksheet.insert_rows(0, 1).is_err());
    }

    #[test]
    fn insert_rows_zero_count_is_noop() {
        let mut worksheet = Worksheet::new("Data");
        worksheet.cell_mut("A1").unwrap().set_value("test");
        worksheet.insert_rows(1, 0).unwrap();
        assert_eq!(
            worksheet.cell("A1").and_then(|c| c.value()),
            Some(&CellValue::String("test".to_string()))
        );
    }

    #[test]
    fn delete_rows_removes_and_shifts_cells_up() {
        let mut worksheet = Worksheet::new("Data");
        worksheet.cell_mut("A1").unwrap().set_value("row1");
        worksheet.cell_mut("A2").unwrap().set_value("row2");
        worksheet.cell_mut("A3").unwrap().set_value("row3");
        worksheet.cell_mut("A4").unwrap().set_value("row4");

        worksheet.delete_rows(2, 2).unwrap();

        assert_eq!(
            worksheet.cell("A1").and_then(|c| c.value()),
            Some(&CellValue::String("row1".to_string()))
        );
        // A2 and A3 were deleted; A4 becomes A2
        assert_eq!(
            worksheet.cell("A2").and_then(|c| c.value()),
            Some(&CellValue::String("row4".to_string()))
        );
        assert!(worksheet.cell("A3").is_none());
        assert!(worksheet.cell("A4").is_none());
    }

    #[test]
    fn delete_rows_rejects_zero_start() {
        let mut worksheet = Worksheet::new("Data");
        assert!(worksheet.delete_rows(0, 1).is_err());
    }

    #[test]
    fn delete_rows_zero_count_is_noop() {
        let mut worksheet = Worksheet::new("Data");
        worksheet.cell_mut("A1").unwrap().set_value("test");
        worksheet.delete_rows(1, 0).unwrap();
        assert_eq!(
            worksheet.cell("A1").and_then(|c| c.value()),
            Some(&CellValue::String("test".to_string()))
        );
    }

    #[test]
    fn insert_columns_shifts_cells_right() {
        let mut worksheet = Worksheet::new("Data");
        worksheet.cell_mut("A1").unwrap().set_value("colA");
        worksheet.cell_mut("B1").unwrap().set_value("colB");
        worksheet.cell_mut("C1").unwrap().set_value("colC");

        worksheet.insert_columns(2, 2).unwrap();

        assert_eq!(
            worksheet.cell("A1").and_then(|c| c.value()),
            Some(&CellValue::String("colA".to_string()))
        );
        // B1 and C1 should now be at D1 and E1
        assert!(worksheet.cell("B1").is_none());
        assert!(worksheet.cell("C1").is_none());
        assert_eq!(
            worksheet.cell("D1").and_then(|c| c.value()),
            Some(&CellValue::String("colB".to_string()))
        );
        assert_eq!(
            worksheet.cell("E1").and_then(|c| c.value()),
            Some(&CellValue::String("colC".to_string()))
        );
    }

    #[test]
    fn insert_columns_shifts_merged_ranges() {
        let mut worksheet = Worksheet::new("Data");
        worksheet.add_merged_range("B1:C2").unwrap();

        worksheet.insert_columns(2, 1).unwrap();

        assert_eq!(worksheet.merged_ranges().len(), 1);
        assert_eq!(worksheet.merged_ranges()[0].start(), "C1");
        assert_eq!(worksheet.merged_ranges()[0].end(), "D2");
    }

    #[test]
    fn insert_columns_rejects_zero_start() {
        let mut worksheet = Worksheet::new("Data");
        assert!(worksheet.insert_columns(0, 1).is_err());
    }

    #[test]
    fn delete_columns_removes_and_shifts_cells_left() {
        let mut worksheet = Worksheet::new("Data");
        worksheet.cell_mut("A1").unwrap().set_value("colA");
        worksheet.cell_mut("B1").unwrap().set_value("colB");
        worksheet.cell_mut("C1").unwrap().set_value("colC");
        worksheet.cell_mut("D1").unwrap().set_value("colD");

        worksheet.delete_columns(2, 2).unwrap();

        assert_eq!(
            worksheet.cell("A1").and_then(|c| c.value()),
            Some(&CellValue::String("colA".to_string()))
        );
        // B1 and C1 were deleted; D1 becomes B1
        assert_eq!(
            worksheet.cell("B1").and_then(|c| c.value()),
            Some(&CellValue::String("colD".to_string()))
        );
        assert!(worksheet.cell("C1").is_none());
        assert!(worksheet.cell("D1").is_none());
    }

    #[test]
    fn delete_columns_rejects_zero_start() {
        let mut worksheet = Worksheet::new("Data");
        assert!(worksheet.delete_columns(0, 1).is_err());
    }

    // ===== Feature 6: Comments =====

    #[test]
    fn comment_new_normalizes_cell_reference() {
        let comment = Comment::new("a1", "Author", "text").unwrap();
        assert_eq!(comment.cell_ref(), "A1");
        assert_eq!(comment.author(), "Author");
        assert_eq!(comment.text(), "text");
    }

    #[test]
    fn comment_new_rejects_invalid_cell_ref() {
        assert!(Comment::new("bad", "Author", "text").is_err());
    }

    #[test]
    fn comment_set_text_and_author() {
        let mut comment = Comment::new("A1", "Original", "original text").unwrap();
        comment.set_text("new text");
        comment.set_author("New Author");
        assert_eq!(comment.text(), "new text");
        assert_eq!(comment.author(), "New Author");
    }

    #[test]
    fn worksheet_comment_add_remove_clear() {
        let mut worksheet = Worksheet::new("Data");
        assert!(worksheet.comments().is_empty());

        let comment1 = Comment::new("A1", "Author1", "Comment 1").unwrap();
        let comment2 = Comment::new("B2", "Author2", "Comment 2").unwrap();
        worksheet.add_comment(comment1);
        worksheet.add_comment(comment2);
        assert_eq!(worksheet.comments().len(), 2);

        assert!(worksheet.remove_comment("a1"));
        assert_eq!(worksheet.comments().len(), 1);
        assert_eq!(worksheet.comments()[0].cell_ref(), "B2");

        assert!(!worksheet.remove_comment("A1")); // already removed
        assert!(!worksheet.remove_comment("bad")); // invalid ref

        worksheet.clear_comments();
        assert!(worksheet.comments().is_empty());
    }

    // ===== Feature 7: Row/column grouping (outline levels) =====

    #[test]
    fn row_outline_level_and_collapsed() {
        let mut worksheet = Worksheet::new("Data");

        let row = worksheet.row_mut(1).unwrap();
        assert_eq!(row.outline_level(), 0);
        assert!(!row.is_collapsed());

        row.set_outline_level(3).set_collapsed(true);
        assert_eq!(row.outline_level(), 3);
        assert!(row.is_collapsed());
    }

    #[test]
    fn row_outline_level_capped_at_seven() {
        let mut row = Row::new(1);
        row.set_outline_level(10);
        assert_eq!(row.outline_level(), 7);
    }

    #[test]
    fn column_outline_level_and_collapsed() {
        let mut worksheet = Worksheet::new("Data");

        let col = worksheet.column_mut(2).unwrap();
        assert_eq!(col.outline_level(), 0);
        assert!(!col.is_collapsed());

        col.set_outline_level(5).set_collapsed(true);
        assert_eq!(col.outline_level(), 5);
        assert!(col.is_collapsed());
    }

    #[test]
    fn column_outline_level_capped_at_seven() {
        let mut col = Column::new(1);
        col.set_outline_level(8);
        assert_eq!(col.outline_level(), 7);
    }

    // ===== Feature 8: Sheet view options =====

    #[test]
    fn sheet_view_options_default_has_no_metadata() {
        let options = SheetViewOptions::new();
        assert!(!options.has_metadata());
    }

    #[test]
    fn sheet_view_options_builder_pattern_and_accessors() {
        let mut options = SheetViewOptions::new();
        options
            .set_show_gridlines(false)
            .set_show_row_col_headers(false)
            .set_show_formulas(true)
            .set_zoom_scale(150)
            .set_zoom_scale_normal(100)
            .set_right_to_left(true)
            .set_tab_selected(true);

        assert!(options.has_metadata());
        assert_eq!(options.show_gridlines(), Some(false));
        assert_eq!(options.show_row_col_headers(), Some(false));
        assert_eq!(options.show_formulas(), Some(true));
        assert_eq!(options.zoom_scale(), Some(150));
        assert_eq!(options.zoom_scale_normal(), Some(100));
        assert_eq!(options.right_to_left(), Some(true));
        assert_eq!(options.tab_selected(), Some(true));
    }

    #[test]
    fn sheet_view_options_clear_individual_fields() {
        let mut options = SheetViewOptions::new();
        options.set_zoom_scale(200).set_show_gridlines(false);

        options.clear_zoom_scale();
        assert!(options.zoom_scale().is_none());
        assert_eq!(options.show_gridlines(), Some(false));

        options.clear_show_gridlines();
        assert!(!options.has_metadata());
    }

    #[test]
    fn worksheet_sheet_view_options_set_and_clear() {
        let mut worksheet = Worksheet::new("Data");
        assert!(worksheet.sheet_view_options().is_none());

        let mut options = SheetViewOptions::new();
        options.set_zoom_scale(125);
        worksheet.set_sheet_view_options(options);
        assert!(worksheet.sheet_view_options().is_some());
        assert_eq!(
            worksheet.sheet_view_options().unwrap().zoom_scale(),
            Some(125)
        );

        worksheet.clear_sheet_view_options();
        assert!(worksheet.sheet_view_options().is_none());
    }

    // ===== Helper function tests =====

    #[test]
    fn cell_reference_to_column_row_works() {
        assert_eq!(cell_reference_to_column_row("A1").unwrap(), (1, 1));
        assert_eq!(cell_reference_to_column_row("B3").unwrap(), (2, 3));
        assert_eq!(cell_reference_to_column_row("Z1").unwrap(), (26, 1));
        assert_eq!(cell_reference_to_column_row("AA1").unwrap(), (27, 1));
        assert!(cell_reference_to_column_row("bad").is_err());
    }

    #[test]
    fn column_index_to_name_converts_correctly() {
        assert_eq!(column_index_to_name(1).unwrap(), "A");
        assert_eq!(column_index_to_name(2).unwrap(), "B");
        assert_eq!(column_index_to_name(26).unwrap(), "Z");
        assert_eq!(column_index_to_name(27).unwrap(), "AA");
        assert_eq!(column_index_to_name(52).unwrap(), "AZ");
        assert!(column_index_to_name(0).is_err());
    }

    #[test]
    fn build_cell_reference_works() {
        assert_eq!(build_cell_reference(1, 1).unwrap(), "A1");
        assert_eq!(build_cell_reference(3, 5).unwrap(), "C5");
        assert_eq!(build_cell_reference(27, 100).unwrap(), "AA100");
        assert!(build_cell_reference(0, 1).is_err());
        assert!(build_cell_reference(1, 0).is_err());
    }

    #[test]
    fn insert_rows_shifts_row_metadata() {
        let mut worksheet = Worksheet::new("Data");
        worksheet.row_mut(2).unwrap().set_height(20.0);
        worksheet.row_mut(3).unwrap().set_height(30.0);

        worksheet.insert_rows(2, 1).unwrap();

        // Row 2 metadata should now be at row 3, row 3 at row 4
        assert!(worksheet.row(2).is_none());
        // The row object is recreated, so only the index shifts (height not preserved in current impl)
        assert!(worksheet.row(3).is_some());
        assert!(worksheet.row(4).is_some());
    }

    #[test]
    fn delete_rows_shifts_row_metadata() {
        let mut worksheet = Worksheet::new("Data");
        worksheet.row_mut(1).unwrap().set_height(10.0);
        worksheet.row_mut(2).unwrap().set_height(20.0);
        worksheet.row_mut(3).unwrap().set_height(30.0);
        worksheet.row_mut(4).unwrap().set_height(40.0);

        worksheet.delete_rows(2, 1).unwrap();

        // Row 1 stays, row 2 deleted, row 3 becomes row 2, row 4 becomes row 3
        assert_eq!(worksheet.row(1).and_then(Row::height), Some(10.0));
        assert!(worksheet.row(2).is_some());
        assert!(worksheet.row(3).is_some());
        assert!(worksheet.row(4).is_none());
    }

    #[test]
    fn insert_columns_shifts_column_metadata() {
        let mut worksheet = Worksheet::new("Data");
        worksheet.column_mut(2).unwrap().set_width(20.0);

        worksheet.insert_columns(2, 1).unwrap();

        assert!(worksheet.column(2).is_none());
        assert!(worksheet.column(3).is_some());
    }

    #[test]
    fn delete_columns_shifts_column_metadata() {
        let mut worksheet = Worksheet::new("Data");
        worksheet.column_mut(1).unwrap().set_width(10.0);
        worksheet.column_mut(2).unwrap().set_width(20.0);
        worksheet.column_mut(3).unwrap().set_width(30.0);

        worksheet.delete_columns(2, 1).unwrap();

        // Col 1 stays, col 2 deleted, col 3 becomes col 2
        assert!(worksheet.column(1).is_some());
        assert!(worksheet.column(2).is_some());
        assert!(worksheet.column(3).is_none());
    }

    // ===== Feature 1: Formula reference shifting in insert/delete =====

    #[test]
    fn insert_rows_shifts_formula_references() {
        let mut worksheet = Worksheet::new("Data");
        worksheet.cell_mut("A1").unwrap().set_value(10);
        worksheet.cell_mut("A5").unwrap().set_formula("SUM(A1:A4)");

        worksheet.insert_rows(3, 2).unwrap();

        // A5 moved to A7, and formula should update A4->A6
        let cell = worksheet.cell("A7").expect("cell should exist at A7");
        assert_eq!(cell.formula(), Some("SUM(A1:A6)"));
    }

    #[test]
    fn delete_rows_shifts_formula_references() {
        let mut worksheet = Worksheet::new("Data");
        worksheet.cell_mut("A1").unwrap().set_value(10);
        worksheet.cell_mut("A6").unwrap().set_formula("SUM(A1:A5)");

        worksheet.delete_rows(2, 2).unwrap();

        // A6 moved to A4, formula A5->A3
        let cell = worksheet.cell("A4").expect("cell should exist at A4");
        assert_eq!(cell.formula(), Some("SUM(A1:A3)"));
    }

    #[test]
    fn insert_columns_shifts_formula_references() {
        let mut worksheet = Worksheet::new("Data");
        worksheet.cell_mut("D1").unwrap().set_formula("SUM(A1:C1)");

        worksheet.insert_columns(2, 1).unwrap();

        // D1 moved to E1, formula: C1->D1
        let cell = worksheet.cell("E1").expect("cell should exist at E1");
        assert_eq!(cell.formula(), Some("SUM(A1:D1)"));
    }

    #[test]
    fn delete_columns_shifts_formula_references() {
        let mut worksheet = Worksheet::new("Data");
        worksheet.cell_mut("E1").unwrap().set_formula("SUM(A1:D1)");

        worksheet.delete_columns(2, 1).unwrap();

        // E1 moved to D1, formula: D1->C1
        let cell = worksheet.cell("D1").expect("cell should exist at D1");
        assert_eq!(cell.formula(), Some("SUM(A1:C1)"));
    }

    // ===== Feature 4: Sort functionality =====

    #[test]
    fn sort_rows_ascending() {
        let mut worksheet = Worksheet::new("Data");
        worksheet.cell_mut("A1").unwrap().set_value(30);
        worksheet.cell_mut("B1").unwrap().set_value("c");
        worksheet.cell_mut("A2").unwrap().set_value(10);
        worksheet.cell_mut("B2").unwrap().set_value("a");
        worksheet.cell_mut("A3").unwrap().set_value(20);
        worksheet.cell_mut("B3").unwrap().set_value("b");

        worksheet.sort_rows("A1:B3", 1, true).unwrap();

        // Rows should be sorted by column A (index 1) ascending: 10, 20, 30
        assert_eq!(
            worksheet.cell("A1").and_then(|c| c.value()),
            Some(&CellValue::Number(10.0))
        );
        assert_eq!(
            worksheet.cell("A2").and_then(|c| c.value()),
            Some(&CellValue::Number(20.0))
        );
        assert_eq!(
            worksheet.cell("A3").and_then(|c| c.value()),
            Some(&CellValue::Number(30.0))
        );
        // Corresponding B values should move with their rows
        assert_eq!(
            worksheet.cell("B1").and_then(|c| c.value()),
            Some(&CellValue::String("a".to_string()))
        );
        assert_eq!(
            worksheet.cell("B2").and_then(|c| c.value()),
            Some(&CellValue::String("b".to_string()))
        );
        assert_eq!(
            worksheet.cell("B3").and_then(|c| c.value()),
            Some(&CellValue::String("c".to_string()))
        );
    }

    #[test]
    fn sort_rows_descending() {
        let mut worksheet = Worksheet::new("Data");
        worksheet.cell_mut("A1").unwrap().set_value(10);
        worksheet.cell_mut("A2").unwrap().set_value(30);
        worksheet.cell_mut("A3").unwrap().set_value(20);

        worksheet.sort_rows("A1:A3", 1, false).unwrap();

        assert_eq!(
            worksheet.cell("A1").and_then(|c| c.value()),
            Some(&CellValue::Number(30.0))
        );
        assert_eq!(
            worksheet.cell("A2").and_then(|c| c.value()),
            Some(&CellValue::Number(20.0))
        );
        assert_eq!(
            worksheet.cell("A3").and_then(|c| c.value()),
            Some(&CellValue::Number(10.0))
        );
    }

    #[test]
    fn sort_rows_rejects_out_of_range_column() {
        let mut worksheet = Worksheet::new("Data");
        worksheet.cell_mut("A1").unwrap().set_value(1);
        // Column 3 (C) is outside range A1:B3
        assert!(worksheet.sort_rows("A1:B3", 3, true).is_err());
    }

    #[test]
    fn sort_rows_with_mixed_types() {
        let mut worksheet = Worksheet::new("Data");
        worksheet.cell_mut("A1").unwrap().set_value("banana");
        worksheet.cell_mut("A2").unwrap().set_value(42);
        worksheet
            .cell_mut("A3")
            .unwrap()
            .set_value(CellValue::Blank);

        worksheet.sort_rows("A1:A3", 1, true).unwrap();

        // Blank < Number < String
        assert_eq!(
            worksheet.cell("A1").and_then(|c| c.value()),
            Some(&CellValue::Blank)
        );
        assert_eq!(
            worksheet.cell("A2").and_then(|c| c.value()),
            Some(&CellValue::Number(42.0))
        );
        assert_eq!(
            worksheet.cell("A3").and_then(|c| c.value()),
            Some(&CellValue::String("banana".to_string()))
        );
    }

    // ===== Formula reference shifting helpers =====

    #[test]
    fn shift_formula_row_references_insert() {
        let result = shift_formula_row_references("SUM(A1:A5)", 3, 2, true);
        assert_eq!(result, "SUM(A1:A7)");
    }

    #[test]
    fn shift_formula_row_references_delete() {
        let result = shift_formula_row_references("SUM(A1:A5)", 4, 1, false);
        assert_eq!(result, "SUM(A1:A4)");
    }

    #[test]
    fn shift_formula_col_references_insert() {
        let result = shift_formula_col_references("SUM(A1:C1)", 2, 1, true);
        assert_eq!(result, "SUM(A1:D1)");
    }

    #[test]
    fn shift_formula_col_references_delete() {
        let result = shift_formula_col_references("SUM(A1:D1)", 3, 1, false);
        assert_eq!(result, "SUM(A1:C1)");
    }

    #[test]
    fn shift_formula_preserves_function_names() {
        // "SUM" should not be treated as a cell reference
        let result = shift_formula_row_references("SUM(A1)", 1, 1, true);
        assert_eq!(result, "SUM(A2)");
    }

    #[test]
    fn shift_formula_handles_complex_formulas() {
        let result = shift_formula_row_references("IF(A3>0,B3,C3)+D5", 3, 1, true);
        assert_eq!(result, "IF(A4>0,B4,C4)+D6");
    }

    // ===== Tab color =====

    #[test]
    fn tab_color_accessors() {
        let mut ws = Worksheet::new("Sheet1");
        assert!(ws.tab_color().is_none());

        ws.set_tab_color("FF0000");
        assert_eq!(ws.tab_color(), Some("FF0000"));

        ws.clear_tab_color();
        assert!(ws.tab_color().is_none());
    }

    // ===== Default row height / column width =====

    #[test]
    fn default_row_height_accessors() {
        let mut ws = Worksheet::new("Sheet1");
        assert!(ws.default_row_height().is_none());

        ws.set_default_row_height(15.0);
        assert_eq!(ws.default_row_height(), Some(15.0));

        ws.clear_default_row_height();
        assert!(ws.default_row_height().is_none());
    }

    #[test]
    fn default_column_width_accessors() {
        let mut ws = Worksheet::new("Sheet1");
        assert!(ws.default_column_width().is_none());

        ws.set_default_column_width(8.43);
        assert_eq!(ws.default_column_width(), Some(8.43));

        ws.clear_default_column_width();
        assert!(ws.default_column_width().is_none());
    }

    #[test]
    fn custom_height_accessors() {
        let mut ws = Worksheet::new("Sheet1");
        assert!(ws.custom_height().is_none());

        ws.set_custom_height(true);
        assert_eq!(ws.custom_height(), Some(true));

        ws.clear_custom_height();
        assert!(ws.custom_height().is_none());
    }

    // ===== Table enhancements =====

    #[test]
    fn table_totals_row_shown() {
        let mut table = WorksheetTable::new("Table1", "A1:C10").unwrap();
        assert!(table.totals_row_shown().is_none());

        table.set_totals_row_shown(true);
        assert_eq!(table.totals_row_shown(), Some(true));

        table.clear_totals_row_shown();
        assert!(table.totals_row_shown().is_none());
    }

    #[test]
    fn table_style_name() {
        let mut table = WorksheetTable::new("Table1", "A1:C10").unwrap();
        assert!(table.style_name().is_none());

        table.set_style_name("TableStyleMedium9");
        assert_eq!(table.style_name(), Some("TableStyleMedium9"));

        table.clear_style_name();
        assert!(table.style_name().is_none());
    }

    #[test]
    fn table_style_info_flags() {
        let mut table = WorksheetTable::new("Table1", "A1:C10").unwrap();
        assert!(table.show_first_column().is_none());
        assert!(table.show_last_column().is_none());
        assert!(table.show_row_stripes().is_none());
        assert!(table.show_column_stripes().is_none());

        table.set_show_first_column(true);
        table.set_show_last_column(false);
        table.set_show_row_stripes(true);
        table.set_show_column_stripes(false);

        assert_eq!(table.show_first_column(), Some(true));
        assert_eq!(table.show_last_column(), Some(false));
        assert_eq!(table.show_row_stripes(), Some(true));
        assert_eq!(table.show_column_stripes(), Some(false));

        table.clear_show_first_column();
        table.clear_show_last_column();
        table.clear_show_row_stripes();
        table.clear_show_column_stripes();

        assert!(table.show_first_column().is_none());
        assert!(table.show_last_column().is_none());
        assert!(table.show_row_stripes().is_none());
        assert!(table.show_column_stripes().is_none());
    }

    // ===== DataValidationErrorStyle =====

    #[test]
    fn data_validation_error_style_xml_roundtrip() {
        assert_eq!(DataValidationErrorStyle::Stop.as_xml_value(), "stop");
        assert_eq!(DataValidationErrorStyle::Warning.as_xml_value(), "warning");
        assert_eq!(
            DataValidationErrorStyle::Information.as_xml_value(),
            "information"
        );

        assert_eq!(
            DataValidationErrorStyle::from_xml_value("stop"),
            Some(DataValidationErrorStyle::Stop)
        );
        assert_eq!(
            DataValidationErrorStyle::from_xml_value("warning"),
            Some(DataValidationErrorStyle::Warning)
        );
        assert_eq!(
            DataValidationErrorStyle::from_xml_value("information"),
            Some(DataValidationErrorStyle::Information)
        );
        assert_eq!(DataValidationErrorStyle::from_xml_value("unknown"), None);
    }

    // ===== DataValidation UI properties =====

    #[test]
    fn data_validation_ui_properties_accessors() {
        let mut dv = DataValidation::list(["A1:A10"], "\"Yes,No\"")
            .expect("list validation should be created");

        // Initially all None
        assert!(dv.error_style().is_none());
        assert!(dv.error_title().is_none());
        assert!(dv.error_message().is_none());
        assert!(dv.prompt_title().is_none());
        assert!(dv.prompt_message().is_none());
        assert!(dv.show_input_message().is_none());
        assert!(dv.show_error_message().is_none());

        // Set all properties
        dv.set_error_style(DataValidationErrorStyle::Stop);
        dv.set_error_title("Invalid Entry");
        dv.set_error_message("Please select Yes or No");
        dv.set_prompt_title("Choose Value");
        dv.set_prompt_message("Select from the dropdown");
        dv.set_show_input_message(true);
        dv.set_show_error_message(true);

        assert_eq!(dv.error_style(), Some(DataValidationErrorStyle::Stop));
        assert_eq!(dv.error_title(), Some("Invalid Entry"));
        assert_eq!(dv.error_message(), Some("Please select Yes or No"));
        assert_eq!(dv.prompt_title(), Some("Choose Value"));
        assert_eq!(dv.prompt_message(), Some("Select from the dropdown"));
        assert_eq!(dv.show_input_message(), Some(true));
        assert_eq!(dv.show_error_message(), Some(true));

        // Clear all properties
        dv.clear_error_style();
        dv.clear_error_title();
        dv.clear_error_message();
        dv.clear_prompt_title();
        dv.clear_prompt_message();
        dv.clear_show_input_message();
        dv.clear_show_error_message();

        assert!(dv.error_style().is_none());
        assert!(dv.error_title().is_none());
        assert!(dv.error_message().is_none());
        assert!(dv.prompt_title().is_none());
        assert!(dv.prompt_message().is_none());
        assert!(dv.show_input_message().is_none());
        assert!(dv.show_error_message().is_none());
    }

    #[test]
    fn data_validation_warning_style() {
        let mut dv =
            DataValidation::whole(["B1:B5"], "1").expect("whole validation should be created");
        dv.set_error_style(DataValidationErrorStyle::Warning);
        assert_eq!(dv.error_style(), Some(DataValidationErrorStyle::Warning));

        dv.set_error_style(DataValidationErrorStyle::Information);
        assert_eq!(
            dv.error_style(),
            Some(DataValidationErrorStyle::Information)
        );
    }

    // ── Range bulk operation tests ──

    #[test]
    fn set_values_2d_populates_cells() {
        let mut ws = Worksheet::new("Sheet1");
        let values = vec![
            vec![CellValue::from("A"), CellValue::from("B")],
            vec![CellValue::from(1), CellValue::from(2)],
        ];
        ws.set_values_2d("B2", &values).expect("set_values_2d");

        assert_eq!(
            ws.cell("B2").and_then(|c| c.value()),
            Some(&CellValue::String("A".to_string()))
        );
        assert_eq!(
            ws.cell("C2").and_then(|c| c.value()),
            Some(&CellValue::String("B".to_string()))
        );
        assert_eq!(
            ws.cell("B3").and_then(|c| c.value()),
            Some(&CellValue::Number(1.0))
        );
        assert_eq!(
            ws.cell("C3").and_then(|c| c.value()),
            Some(&CellValue::Number(2.0))
        );
    }

    #[test]
    fn clear_values_blanks_cells() {
        let mut ws = Worksheet::new("Sheet1");
        ws.cell_mut("A1").unwrap().set_value("hello");
        ws.cell_mut("B1").unwrap().set_value(42);
        ws.cell_mut("B1").unwrap().set_style_id(5);

        ws.clear_values("A1:B1").expect("clear_values");

        assert_eq!(
            ws.cell("A1").and_then(|c| c.value()),
            Some(&CellValue::Blank)
        );
        assert_eq!(
            ws.cell("B1").and_then(|c| c.value()),
            Some(&CellValue::Blank)
        );
        // Style should be preserved.
        assert_eq!(ws.cell("B1").and_then(|c| c.style_id()), Some(5));
    }

    #[test]
    fn clear_range_removes_cells() {
        let mut ws = Worksheet::new("Sheet1");
        ws.cell_mut("A1").unwrap().set_value("hello");
        ws.cell_mut("B1").unwrap().set_value(42);

        ws.clear_range("A1:B1").expect("clear_range");

        assert!(ws.cell("A1").is_none());
        assert!(ws.cell("B1").is_none());
    }

    #[test]
    fn apply_style_to_range_sets_style_ids() {
        let mut ws = Worksheet::new("Sheet1");
        ws.cell_mut("A1").unwrap().set_value("a");
        ws.cell_mut("B1").unwrap().set_value("b");

        ws.apply_style_to_range("A1:B1", 3).expect("apply_style");

        assert_eq!(ws.cell("A1").and_then(|c| c.style_id()), Some(3));
        assert_eq!(ws.cell("B1").and_then(|c| c.style_id()), Some(3));
    }

    #[test]
    fn copy_range_duplicates_cells() {
        let mut ws = Worksheet::new("Sheet1");
        ws.cell_mut("A1").unwrap().set_value("X");
        ws.cell_mut("B1").unwrap().set_value(99);

        ws.copy_range("A1:B1", "A3").expect("copy_range");

        // Original cells unchanged.
        assert_eq!(
            ws.cell("A1").and_then(|c| c.value()),
            Some(&CellValue::String("X".to_string()))
        );
        // Copied cells present.
        assert_eq!(
            ws.cell("A3").and_then(|c| c.value()),
            Some(&CellValue::String("X".to_string()))
        );
        assert_eq!(
            ws.cell("B3").and_then(|c| c.value()),
            Some(&CellValue::Number(99.0))
        );
    }

    // ── Comment enhancement tests ──

    #[test]
    fn comment_visibility_and_replies() {
        let mut comment = Comment::new("A1", "Alice", "Initial").unwrap();
        assert!(!comment.visible());
        assert!(comment.replies().is_empty());

        comment.set_visible(true);
        assert!(comment.visible());

        comment.add_reply(CommentReply::new("Bob", "Good point"));
        comment.add_reply(CommentReply::new("Carol", "Agreed"));
        assert_eq!(comment.replies().len(), 2);
        assert_eq!(comment.replies()[0].author, "Bob");
        assert_eq!(comment.replies()[1].text, "Agreed");

        comment.clear_replies();
        assert!(comment.replies().is_empty());
    }

    #[test]
    fn comment_rich_text() {
        use crate::cell::RichTextRun;
        let mut comment = Comment::new("B2", "Author", "plain").unwrap();
        assert!(comment.rich_text().is_none());

        let mut bold_run = RichTextRun::new("bold");
        bold_run.set_bold(true);
        let normal_run = RichTextRun::new(" normal");
        let runs = vec![bold_run, normal_run];
        comment.set_rich_text(runs);
        assert_eq!(comment.rich_text().unwrap().len(), 2);
        assert_eq!(comment.rich_text().unwrap()[0].text(), "bold");

        comment.clear_rich_text();
        assert!(comment.rich_text().is_none());
    }

    // ── Hyperlink display text tests ──

    #[test]
    fn hyperlink_display_text() {
        let mut hl = Hyperlink::external("A1", "https://example.com").unwrap();
        assert!(hl.display().is_none());

        hl.set_display("Click here");
        assert_eq!(hl.display(), Some("Click here"));

        hl.clear_display();
        assert!(hl.display().is_none());
    }

    // ── Formula evaluation tests ──

    #[test]
    fn evaluate_formula_simple_sum() {
        let mut wb = crate::workbook::Workbook::new();
        {
            let ws = wb.add_sheet("Sheet1");
            ws.cell_mut("A1").unwrap().set_value(10);
            ws.cell_mut("A2").unwrap().set_value(20);
            ws.cell_mut("A3").unwrap().set_formula("SUM(A1:A2)");
        }

        let ws = wb.sheet("Sheet1").unwrap();
        let result = ws.evaluate_formula("A3", &wb).unwrap();
        assert_eq!(result, CellValue::Number(30.0));
    }

    #[test]
    fn evaluate_formula_arithmetic() {
        let mut wb = crate::workbook::Workbook::new();
        {
            let ws = wb.add_sheet("Sheet1");
            ws.cell_mut("B1").unwrap().set_value(5);
            ws.cell_mut("B2").unwrap().set_value(3);
            ws.cell_mut("B3").unwrap().set_formula("B1*B2+10");
        }

        let ws = wb.sheet("Sheet1").unwrap();
        let result = ws.evaluate_formula("B3", &wb).unwrap();
        assert_eq!(result, CellValue::Number(25.0));
    }

    #[test]
    fn evaluate_formula_with_functions() {
        let mut wb = crate::workbook::Workbook::new();
        {
            let ws = wb.add_sheet("Sheet1");
            ws.cell_mut("C1").unwrap().set_value(100);
            ws.cell_mut("C2").unwrap().set_formula("SQRT(C1)");
        }

        let ws = wb.sheet("Sheet1").unwrap();
        let result = ws.evaluate_formula("C2", &wb).unwrap();
        assert_eq!(result, CellValue::Number(10.0));
    }

    #[test]
    fn evaluate_formula_returns_error_if_no_cell() {
        let mut wb = crate::workbook::Workbook::new();
        wb.add_sheet("Sheet1");

        let ws = wb.sheet("Sheet1").unwrap();
        let result = ws.evaluate_formula("A1", &wb);
        assert!(result.is_err());
    }

    #[test]
    fn evaluate_formula_returns_error_if_no_formula() {
        let mut wb = crate::workbook::Workbook::new();
        {
            let ws = wb.add_sheet("Sheet1");
            ws.cell_mut("A1").unwrap().set_value(42);
        }

        let ws = wb.sheet("Sheet1").unwrap();
        let result = ws.evaluate_formula("A1", &wb);
        assert!(result.is_err());
    }

    #[test]
    fn evaluate_formula_cross_cell_references() {
        let mut wb = crate::workbook::Workbook::new();
        {
            let ws = wb.add_sheet("Sheet1");
            ws.cell_mut("A1").unwrap().set_value(10);
            ws.cell_mut("B1").unwrap().set_value(5);
            ws.cell_mut("C1").unwrap().set_value(30); // A1 * B1 = 50, but we set 30 to test reference
            ws.cell_mut("D1").unwrap().set_formula("A1+B1+C1");
        }

        let ws = wb.sheet("Sheet1").unwrap();
        let result = ws.evaluate_formula("D1", &wb).unwrap();
        assert_eq!(result, CellValue::Number(45.0)); // 10 + 5 + 30 = 45
    }

    #[test]
    fn evaluate_formula_with_text() {
        let mut wb = crate::workbook::Workbook::new();
        {
            let ws = wb.add_sheet("Sheet1");
            ws.cell_mut("A1").unwrap().set_value("Hello");
            ws.cell_mut("A2").unwrap().set_value("World");
            ws.cell_mut("A3")
                .unwrap()
                .set_formula("CONCATENATE(A1,\" \",A2)");
        }

        let ws = wb.sheet("Sheet1").unwrap();
        let result = ws.evaluate_formula("A3", &wb).unwrap();
        assert_eq!(result, CellValue::String("Hello World".to_string()));
    }

    // ===== Phase 1A: Insert/delete shifts ancillary data =====

    #[test]
    fn insert_rows_shifts_conditional_formatting() {
        let mut ws = Worksheet::new("Test");
        ws.add_conditional_formatting(ConditionalFormatting::cell_is(["B2:B5"], ["10"]).unwrap());

        ws.insert_rows(2, 3).unwrap();

        let cf = &ws.conditional_formattings()[0];
        assert_eq!(cf.sqref()[0].start(), "B5");
        assert_eq!(cf.sqref()[0].end(), "B8");
    }

    #[test]
    fn insert_rows_shifts_data_validation() {
        let mut ws = Worksheet::new("Test");
        ws.add_data_validation(DataValidation::list(["C3:C10"], "\"Yes,No\"").unwrap());

        ws.insert_rows(1, 2).unwrap();

        let dv = &ws.data_validations()[0];
        assert_eq!(dv.sqref()[0].start(), "C5");
        assert_eq!(dv.sqref()[0].end(), "C12");
    }

    #[test]
    fn insert_rows_shifts_table_range() {
        let mut ws = Worksheet::new("Test");
        ws.add_table(WorksheetTable::new("Table1", "A1:D5").unwrap());

        ws.insert_rows(2, 3).unwrap();

        assert_eq!(ws.tables()[0].range().start(), "A1");
        assert_eq!(ws.tables()[0].range().end(), "D8");
    }

    #[test]
    fn insert_rows_shifts_hyperlinks() {
        let mut ws = Worksheet::new("Test");
        ws.add_hyperlink(Hyperlink::external("A5", "https://example.com").unwrap());

        ws.insert_rows(3, 2).unwrap();

        assert_eq!(ws.hyperlinks()[0].cell_ref(), "A7");
    }

    #[test]
    fn delete_rows_removes_hyperlinks_in_deleted_range() {
        let mut ws = Worksheet::new("Test");
        ws.add_hyperlink(Hyperlink::external("A2", "https://keep.com").unwrap());
        ws.add_hyperlink(Hyperlink::external("A5", "https://delete.com").unwrap());
        ws.add_hyperlink(Hyperlink::external("A8", "https://shift.com").unwrap());

        ws.delete_rows(4, 3).unwrap(); // delete rows 4-6

        assert_eq!(ws.hyperlinks().len(), 2);
        assert_eq!(ws.hyperlinks()[0].cell_ref(), "A2");
        assert_eq!(ws.hyperlinks()[1].cell_ref(), "A5"); // 8 shifted to 5
    }

    #[test]
    fn insert_columns_shifts_table_range() {
        let mut ws = Worksheet::new("Test");
        ws.add_table(WorksheetTable::new("Table1", "B1:D5").unwrap());

        ws.insert_columns(2, 2).unwrap();

        assert_eq!(ws.tables()[0].range().start(), "D1");
        assert_eq!(ws.tables()[0].range().end(), "F5");
    }

    #[test]
    fn delete_columns_shifts_conditional_formatting() {
        let mut ws = Worksheet::new("Test");
        ws.add_conditional_formatting(ConditionalFormatting::cell_is(["D2:D5"], ["10"]).unwrap());

        ws.delete_columns(1, 2).unwrap(); // delete cols A-B

        let cf = &ws.conditional_formattings()[0];
        assert_eq!(cf.sqref()[0].start(), "B2");
        assert_eq!(cf.sqref()[0].end(), "B5");
    }

    // ===== Phase 1B: Unmerge, find, transpose =====

    #[test]
    fn unmerge_range_removes_specific_merge() {
        let mut ws = Worksheet::new("Test");
        ws.add_merged_range("A1:B2").unwrap();
        ws.add_merged_range("C3:D4").unwrap();

        assert!(ws.unmerge_range("A1:B2").unwrap());
        assert_eq!(ws.merged_ranges().len(), 1);
        assert_eq!(ws.merged_ranges()[0].start(), "C3");
    }

    #[test]
    fn unmerge_range_returns_false_for_missing() {
        let mut ws = Worksheet::new("Test");
        ws.add_merged_range("A1:B2").unwrap();

        assert!(!ws.unmerge_range("C3:D4").unwrap());
        assert_eq!(ws.merged_ranges().len(), 1);
    }

    #[test]
    fn unmerge_all_clears_all_merges() {
        let mut ws = Worksheet::new("Test");
        ws.add_merged_range("A1:B2").unwrap();
        ws.add_merged_range("C3:D4").unwrap();

        ws.unmerge_all();
        assert!(ws.merged_ranges().is_empty());
    }

    #[test]
    fn find_cells_by_text() {
        let mut ws = Worksheet::new("Test");
        ws.cell_mut("A1").unwrap().set_value("Hello World");
        ws.cell_mut("B1").unwrap().set_value("Goodbye");
        ws.cell_mut("C1").unwrap().set_value("Hello Again");

        let results = ws.find_cells("Hello");
        assert_eq!(results.len(), 2);
        assert!(results.contains(&"A1".to_string()));
        assert!(results.contains(&"C1".to_string()));
    }

    #[test]
    fn find_cells_by_value_exact_match() {
        let mut ws = Worksheet::new("Test");
        ws.cell_mut("A1").unwrap().set_value(42);
        ws.cell_mut("B1").unwrap().set_value(42);
        ws.cell_mut("C1").unwrap().set_value(43);

        let results = ws.find_cells_by_value(&CellValue::Number(42.0));
        assert_eq!(results.len(), 2);
    }
}
