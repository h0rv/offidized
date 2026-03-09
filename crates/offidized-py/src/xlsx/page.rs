//! Python bindings for worksheet page setup, margins, header/footer, print area,
//! page breaks, and sheet view options from `offidized_xlsx`.

use super::lock_wb;
use crate::error::value_error;
use offidized_xlsx::{
    PageBreaks as CorePageBreaks, PageMargins as CorePageMargins,
    PageOrientation as CorePageOrientation, PageSetup as CorePageSetup, PrintArea as CorePrintArea,
    PrintHeaderFooter as CorePrintHeaderFooter, SheetViewOptions as CoreSheetViewOptions,
    Workbook as CoreWorkbook,
};
use pyo3::prelude::*;
use std::sync::{Arc, Mutex};

// =============================================================================
// XlsxPageSetup
// =============================================================================

/// Page setup settings for worksheet printing.
#[pyclass(module = "offidized._native", name = "XlsxPageSetup")]
#[derive(Clone, Default)]
pub struct XlsxPageSetup {
    inner: CorePageSetup,
}

impl XlsxPageSetup {
    pub(super) fn from_core(core: CorePageSetup) -> Self {
        Self { inner: core }
    }

    pub(super) fn into_core(self) -> CorePageSetup {
        self.inner
    }
}

#[pymethods]
impl XlsxPageSetup {
    #[new]
    pub fn new() -> Self {
        Self::default()
    }

    /// Return the page orientation as "portrait" or "landscape", or None.
    #[getter]
    pub fn orientation(&self) -> Option<String> {
        self.inner.orientation().map(|o| match o {
            CorePageOrientation::Portrait => "portrait".to_string(),
            CorePageOrientation::Landscape => "landscape".to_string(),
        })
    }

    /// Set the page orientation ("portrait" or "landscape").
    #[setter]
    pub fn set_orientation(&mut self, value: Option<String>) -> PyResult<()> {
        match value.as_deref() {
            None => {
                self.inner.clear_orientation();
            }
            Some("portrait") => {
                self.inner.set_orientation(CorePageOrientation::Portrait);
            }
            Some("landscape") => {
                self.inner.set_orientation(CorePageOrientation::Landscape);
            }
            Some(s) => {
                return Err(value_error(format!(
                    "Unknown orientation '{s}': expected 'portrait' or 'landscape'"
                )));
            }
        }
        Ok(())
    }

    /// Return the paper size code (1=Letter, 9=A4, etc.), or None.
    #[getter]
    pub fn paper_size(&self) -> Option<u32> {
        self.inner.paper_size()
    }

    /// Set the paper size code.
    #[setter]
    pub fn set_paper_size(&mut self, value: Option<u32>) {
        if let Some(v) = value {
            self.inner.set_paper_size(v);
        } else {
            self.inner.clear_paper_size();
        }
    }

    /// Return the print scale percentage (10-400), or None.
    #[getter]
    pub fn scale(&self) -> Option<u32> {
        self.inner.scale()
    }

    /// Set the print scale percentage.
    #[setter]
    pub fn set_scale(&mut self, value: Option<u32>) {
        if let Some(v) = value {
            self.inner.set_scale(v);
        } else {
            self.inner.clear_scale();
        }
    }

    /// Return the fit-to-width page count, or None.
    #[getter]
    pub fn fit_to_width(&self) -> Option<u32> {
        self.inner.fit_to_width()
    }

    /// Set the fit-to-width page count.
    #[setter]
    pub fn set_fit_to_width(&mut self, value: Option<u32>) {
        if let Some(v) = value {
            self.inner.set_fit_to_width(v);
        } else {
            self.inner.clear_fit_to_width();
        }
    }

    /// Return the fit-to-height page count, or None.
    #[getter]
    pub fn fit_to_height(&self) -> Option<u32> {
        self.inner.fit_to_height()
    }

    /// Set the fit-to-height page count.
    #[setter]
    pub fn set_fit_to_height(&mut self, value: Option<u32>) {
        if let Some(v) = value {
            self.inner.set_fit_to_height(v);
        } else {
            self.inner.clear_fit_to_height();
        }
    }

    /// Return the horizontal DPI for printing, or None.
    #[getter]
    pub fn horizontal_dpi(&self) -> Option<u32> {
        self.inner.horizontal_dpi()
    }

    /// Set the horizontal DPI for printing.
    #[setter]
    pub fn set_horizontal_dpi(&mut self, value: Option<u32>) {
        if let Some(v) = value {
            self.inner.set_horizontal_dpi(v);
        } else {
            self.inner.clear_horizontal_dpi();
        }
    }

    /// Return the vertical DPI for printing, or None.
    #[getter]
    pub fn vertical_dpi(&self) -> Option<u32> {
        self.inner.vertical_dpi()
    }

    /// Set the vertical DPI for printing.
    #[setter]
    pub fn set_vertical_dpi(&mut self, value: Option<u32>) {
        if let Some(v) = value {
            self.inner.set_vertical_dpi(v);
        } else {
            self.inner.clear_vertical_dpi();
        }
    }
}

// =============================================================================
// XlsxPageMargins
// =============================================================================

/// Page margins for worksheet printing, measured in inches.
#[pyclass(module = "offidized._native", name = "XlsxPageMargins")]
#[derive(Clone, Default)]
pub struct XlsxPageMargins {
    inner: CorePageMargins,
}

impl XlsxPageMargins {
    pub(super) fn from_core(core: CorePageMargins) -> Self {
        Self { inner: core }
    }

    pub(super) fn into_core(self) -> CorePageMargins {
        self.inner
    }
}

#[pymethods]
impl XlsxPageMargins {
    #[new]
    pub fn new() -> Self {
        Self::default()
    }

    /// Return the left margin in inches, or None.
    #[getter]
    pub fn left(&self) -> Option<f64> {
        self.inner.left()
    }

    /// Set the left margin in inches.
    #[setter]
    pub fn set_left(&mut self, value: Option<f64>) {
        if let Some(v) = value {
            self.inner.set_left(v);
        } else {
            self.inner.clear_left();
        }
    }

    /// Return the right margin in inches, or None.
    #[getter]
    pub fn right(&self) -> Option<f64> {
        self.inner.right()
    }

    /// Set the right margin in inches.
    #[setter]
    pub fn set_right(&mut self, value: Option<f64>) {
        if let Some(v) = value {
            self.inner.set_right(v);
        } else {
            self.inner.clear_right();
        }
    }

    /// Return the top margin in inches, or None.
    #[getter]
    pub fn top(&self) -> Option<f64> {
        self.inner.top()
    }

    /// Set the top margin in inches.
    #[setter]
    pub fn set_top(&mut self, value: Option<f64>) {
        if let Some(v) = value {
            self.inner.set_top(v);
        } else {
            self.inner.clear_top();
        }
    }

    /// Return the bottom margin in inches, or None.
    #[getter]
    pub fn bottom(&self) -> Option<f64> {
        self.inner.bottom()
    }

    /// Set the bottom margin in inches.
    #[setter]
    pub fn set_bottom(&mut self, value: Option<f64>) {
        if let Some(v) = value {
            self.inner.set_bottom(v);
        } else {
            self.inner.clear_bottom();
        }
    }

    /// Return the header margin in inches, or None.
    #[getter]
    pub fn header(&self) -> Option<f64> {
        self.inner.header()
    }

    /// Set the header margin in inches.
    #[setter]
    pub fn set_header(&mut self, value: Option<f64>) {
        if let Some(v) = value {
            self.inner.set_header(v);
        } else {
            self.inner.clear_header();
        }
    }

    /// Return the footer margin in inches, or None.
    #[getter]
    pub fn footer(&self) -> Option<f64> {
        self.inner.footer()
    }

    /// Set the footer margin in inches.
    #[setter]
    pub fn set_footer(&mut self, value: Option<f64>) {
        if let Some(v) = value {
            self.inner.set_footer(v);
        } else {
            self.inner.clear_footer();
        }
    }
}

// =============================================================================
// XlsxPrintHeaderFooter
// =============================================================================

/// Header and footer content for printed worksheet pages.
///
/// Strings use OOXML formatting codes such as `&L`, `&C`, `&R` for
/// left/center/right sections, `&P` for page number, and `&D` for date.
#[pyclass(module = "offidized._native", name = "XlsxPrintHeaderFooter")]
#[derive(Clone, Default)]
pub struct XlsxPrintHeaderFooter {
    inner: CorePrintHeaderFooter,
}

impl XlsxPrintHeaderFooter {
    pub(super) fn from_core(core: CorePrintHeaderFooter) -> Self {
        Self { inner: core }
    }

    pub(super) fn into_core(self) -> CorePrintHeaderFooter {
        self.inner
    }
}

#[pymethods]
impl XlsxPrintHeaderFooter {
    #[new]
    pub fn new() -> Self {
        Self::default()
    }

    /// Return the odd-page header string, or None.
    #[getter]
    pub fn odd_header(&self) -> Option<String> {
        self.inner.odd_header().map(|s| s.to_string())
    }

    /// Set the odd-page header string.
    #[setter]
    pub fn set_odd_header(&mut self, value: Option<String>) {
        if let Some(v) = value {
            self.inner.set_odd_header(v);
        } else {
            self.inner.clear_odd_header();
        }
    }

    /// Return the odd-page footer string, or None.
    #[getter]
    pub fn odd_footer(&self) -> Option<String> {
        self.inner.odd_footer().map(|s| s.to_string())
    }

    /// Set the odd-page footer string.
    #[setter]
    pub fn set_odd_footer(&mut self, value: Option<String>) {
        if let Some(v) = value {
            self.inner.set_odd_footer(v);
        } else {
            self.inner.clear_odd_footer();
        }
    }

    /// Return the even-page header string, or None.
    #[getter]
    pub fn even_header(&self) -> Option<String> {
        self.inner.even_header().map(|s| s.to_string())
    }

    /// Set the even-page header string.
    #[setter]
    pub fn set_even_header(&mut self, value: Option<String>) {
        if let Some(v) = value {
            self.inner.set_even_header(v);
        } else {
            self.inner.clear_even_header();
        }
    }

    /// Return the even-page footer string, or None.
    #[getter]
    pub fn even_footer(&self) -> Option<String> {
        self.inner.even_footer().map(|s| s.to_string())
    }

    /// Set the even-page footer string.
    #[setter]
    pub fn set_even_footer(&mut self, value: Option<String>) {
        if let Some(v) = value {
            self.inner.set_even_footer(v);
        } else {
            self.inner.clear_even_footer();
        }
    }

    /// Return the first-page header string, or None.
    #[getter]
    pub fn first_header(&self) -> Option<String> {
        self.inner.first_header().map(|s| s.to_string())
    }

    /// Set the first-page header string.
    #[setter]
    pub fn set_first_header(&mut self, value: Option<String>) {
        if let Some(v) = value {
            self.inner.set_first_header(v);
        } else {
            self.inner.clear_first_header();
        }
    }

    /// Return the first-page footer string, or None.
    #[getter]
    pub fn first_footer(&self) -> Option<String> {
        self.inner.first_footer().map(|s| s.to_string())
    }

    /// Set the first-page footer string.
    #[setter]
    pub fn set_first_footer(&mut self, value: Option<String>) {
        if let Some(v) = value {
            self.inner.set_first_footer(v);
        } else {
            self.inner.clear_first_footer();
        }
    }

    /// Return whether different odd/even headers/footers are enabled.
    #[getter]
    pub fn different_odd_even(&self) -> bool {
        self.inner.different_odd_even()
    }

    /// Set whether different odd/even headers/footers are enabled.
    #[setter]
    pub fn set_different_odd_even(&mut self, value: bool) {
        self.inner.set_different_odd_even(value);
    }

    /// Return whether a different first-page header/footer is enabled.
    #[getter]
    pub fn different_first(&self) -> bool {
        self.inner.different_first()
    }

    /// Set whether a different first-page header/footer is enabled.
    #[setter]
    pub fn set_different_first(&mut self, value: bool) {
        self.inner.set_different_first(value);
    }
}

// =============================================================================
// XlsxPageBreaks
// =============================================================================

/// Collection of row and column page breaks for a worksheet.
#[pyclass(module = "offidized._native", name = "XlsxPageBreaks")]
#[derive(Clone, Default)]
pub struct XlsxPageBreaks {
    inner: CorePageBreaks,
}

impl XlsxPageBreaks {
    pub(super) fn from_core(core: CorePageBreaks) -> Self {
        Self { inner: core }
    }

    pub(super) fn into_core(self) -> CorePageBreaks {
        self.inner
    }
}

#[pymethods]
impl XlsxPageBreaks {
    #[new]
    pub fn new() -> Self {
        Self::default()
    }

    /// Return all row break indices (1-based).
    pub fn row_breaks(&self) -> Vec<u32> {
        self.inner.row_breaks().iter().map(|b| b.id()).collect()
    }

    /// Add a row break at the given 1-based row index.
    pub fn add_row_break(&mut self, row: u32) {
        self.inner.add_row_break(row);
    }

    /// Clear all row breaks.
    pub fn clear_row_breaks(&mut self) {
        self.inner.clear_row_breaks();
    }

    /// Return all column break indices (1-based).
    pub fn col_breaks(&self) -> Vec<u32> {
        self.inner.col_breaks().iter().map(|b| b.id()).collect()
    }

    /// Add a column break at the given 1-based column index.
    pub fn add_col_break(&mut self, col: u32) {
        self.inner.add_col_break(col);
    }

    /// Clear all column breaks.
    pub fn clear_col_breaks(&mut self) {
        self.inner.clear_col_breaks();
    }

    /// Clear all row and column breaks.
    pub fn clear_all(&mut self) {
        self.inner.clear_all();
    }

    /// Return True if there are any row or column breaks.
    pub fn has_breaks(&self) -> bool {
        self.inner.has_breaks()
    }
}

// =============================================================================
// XlsxSheetViewOptions
// =============================================================================

/// Sheet view options controlling visual display.
#[pyclass(module = "offidized._native", name = "XlsxSheetViewOptions")]
#[derive(Clone, Default)]
pub struct XlsxSheetViewOptions {
    inner: CoreSheetViewOptions,
}

impl XlsxSheetViewOptions {
    pub(super) fn from_core(core: CoreSheetViewOptions) -> Self {
        Self { inner: core }
    }

    pub(super) fn into_core(self) -> CoreSheetViewOptions {
        self.inner
    }
}

#[pymethods]
impl XlsxSheetViewOptions {
    #[new]
    pub fn new() -> Self {
        Self::default()
    }

    /// Return whether gridlines are shown, or None if unset.
    #[getter]
    pub fn show_gridlines(&self) -> Option<bool> {
        self.inner.show_gridlines()
    }

    /// Set whether gridlines are shown.
    #[setter]
    pub fn set_show_gridlines(&mut self, value: Option<bool>) {
        if let Some(v) = value {
            self.inner.set_show_gridlines(v);
        } else {
            self.inner.clear_show_gridlines();
        }
    }

    /// Return whether row/column headers are shown, or None if unset.
    #[getter]
    pub fn show_row_col_headers(&self) -> Option<bool> {
        self.inner.show_row_col_headers()
    }

    /// Set whether row/column headers are shown.
    #[setter]
    pub fn set_show_row_col_headers(&mut self, value: Option<bool>) {
        if let Some(v) = value {
            self.inner.set_show_row_col_headers(v);
        } else {
            self.inner.clear_show_row_col_headers();
        }
    }

    /// Return whether formulas are shown instead of values, or None if unset.
    #[getter]
    pub fn show_formulas(&self) -> Option<bool> {
        self.inner.show_formulas()
    }

    /// Set whether formulas are shown instead of values.
    #[setter]
    pub fn set_show_formulas(&mut self, value: Option<bool>) {
        if let Some(v) = value {
            self.inner.set_show_formulas(v);
        } else {
            self.inner.clear_show_formulas();
        }
    }

    /// Return the zoom scale percentage, or None if unset.
    #[getter]
    pub fn zoom_scale(&self) -> Option<u32> {
        self.inner.zoom_scale()
    }

    /// Set the zoom scale percentage (10-400).
    #[setter]
    pub fn set_zoom_scale(&mut self, value: Option<u32>) {
        if let Some(v) = value {
            self.inner.set_zoom_scale(v);
        } else {
            self.inner.clear_zoom_scale();
        }
    }

    /// Return whether the sheet view is right-to-left, or None if unset.
    #[getter]
    pub fn right_to_left(&self) -> Option<bool> {
        self.inner.right_to_left()
    }

    /// Set whether the sheet view is right-to-left.
    #[setter]
    pub fn set_right_to_left(&mut self, value: Option<bool>) {
        if let Some(v) = value {
            self.inner.set_right_to_left(v);
        } else {
            self.inner.clear_right_to_left();
        }
    }

    /// Return whether the tab is selected, or None if unset.
    #[getter]
    pub fn tab_selected(&self) -> Option<bool> {
        self.inner.tab_selected()
    }

    /// Set whether the tab is selected.
    #[setter]
    pub fn set_tab_selected(&mut self, value: Option<bool>) {
        if let Some(v) = value {
            self.inner.set_tab_selected(v);
        } else {
            self.inner.clear_tab_selected();
        }
    }

    /// Return the view mode (e.g. "normal", "pageLayout", "pageBreakPreview"), or None.
    #[getter]
    pub fn view(&self) -> Option<String> {
        self.inner.view().map(|s| s.to_string())
    }

    /// Set the view mode.
    #[setter]
    pub fn set_view(&mut self, value: Option<String>) {
        if let Some(v) = value {
            self.inner.set_view(v);
        } else {
            self.inner.clear_view();
        }
    }
}

// =============================================================================
// Worksheet helper functions
// =============================================================================

pub(super) fn ws_page_setup(
    wb: &Arc<Mutex<CoreWorkbook>>,
    sheet_name: &str,
) -> PyResult<Option<XlsxPageSetup>> {
    let wb = lock_wb(wb)?;
    let ws = wb
        .sheet(sheet_name)
        .ok_or_else(|| value_error(format!("worksheet '{sheet_name}' not found")))?;
    Ok(ws.page_setup().cloned().map(XlsxPageSetup::from_core))
}

pub(super) fn ws_set_page_setup(
    wb: &Arc<Mutex<CoreWorkbook>>,
    sheet_name: &str,
    setup: XlsxPageSetup,
) -> PyResult<()> {
    let mut wb = lock_wb(wb)?;
    let ws = wb
        .sheet_mut(sheet_name)
        .ok_or_else(|| value_error(format!("worksheet '{sheet_name}' not found")))?;
    ws.set_page_setup(setup.into_core());
    Ok(())
}

pub(super) fn ws_page_margins(
    wb: &Arc<Mutex<CoreWorkbook>>,
    sheet_name: &str,
) -> PyResult<Option<XlsxPageMargins>> {
    let wb = lock_wb(wb)?;
    let ws = wb
        .sheet(sheet_name)
        .ok_or_else(|| value_error(format!("worksheet '{sheet_name}' not found")))?;
    Ok(ws.page_margins().cloned().map(XlsxPageMargins::from_core))
}

pub(super) fn ws_set_page_margins(
    wb: &Arc<Mutex<CoreWorkbook>>,
    sheet_name: &str,
    margins: XlsxPageMargins,
) -> PyResult<()> {
    let mut wb = lock_wb(wb)?;
    let ws = wb
        .sheet_mut(sheet_name)
        .ok_or_else(|| value_error(format!("worksheet '{sheet_name}' not found")))?;
    ws.set_page_margins(margins.into_core());
    Ok(())
}

pub(super) fn ws_header_footer(
    wb: &Arc<Mutex<CoreWorkbook>>,
    sheet_name: &str,
) -> PyResult<Option<XlsxPrintHeaderFooter>> {
    let wb = lock_wb(wb)?;
    let ws = wb
        .sheet(sheet_name)
        .ok_or_else(|| value_error(format!("worksheet '{sheet_name}' not found")))?;
    Ok(ws
        .header_footer()
        .cloned()
        .map(XlsxPrintHeaderFooter::from_core))
}

pub(super) fn ws_set_header_footer(
    wb: &Arc<Mutex<CoreWorkbook>>,
    sheet_name: &str,
    hf: XlsxPrintHeaderFooter,
) -> PyResult<()> {
    let mut wb = lock_wb(wb)?;
    let ws = wb
        .sheet_mut(sheet_name)
        .ok_or_else(|| value_error(format!("worksheet '{sheet_name}' not found")))?;
    ws.set_header_footer(hf.into_core());
    Ok(())
}

pub(super) fn ws_print_area(
    wb: &Arc<Mutex<CoreWorkbook>>,
    sheet_name: &str,
) -> PyResult<Option<String>> {
    let wb = lock_wb(wb)?;
    let ws = wb
        .sheet(sheet_name)
        .ok_or_else(|| value_error(format!("worksheet '{sheet_name}' not found")))?;
    Ok(ws.print_area().map(|a| a.range().to_string()))
}

pub(super) fn ws_set_print_area(
    wb: &Arc<Mutex<CoreWorkbook>>,
    sheet_name: &str,
    range: &str,
) -> PyResult<()> {
    let mut wb = lock_wb(wb)?;
    let ws = wb
        .sheet_mut(sheet_name)
        .ok_or_else(|| value_error(format!("worksheet '{sheet_name}' not found")))?;
    ws.set_print_area(CorePrintArea::new(range));
    Ok(())
}

pub(super) fn ws_clear_print_area(wb: &Arc<Mutex<CoreWorkbook>>, sheet_name: &str) -> PyResult<()> {
    let mut wb = lock_wb(wb)?;
    let ws = wb
        .sheet_mut(sheet_name)
        .ok_or_else(|| value_error(format!("worksheet '{sheet_name}' not found")))?;
    ws.clear_print_area();
    Ok(())
}

pub(super) fn ws_page_breaks(
    wb: &Arc<Mutex<CoreWorkbook>>,
    sheet_name: &str,
) -> PyResult<Option<XlsxPageBreaks>> {
    let wb = lock_wb(wb)?;
    let ws = wb
        .sheet(sheet_name)
        .ok_or_else(|| value_error(format!("worksheet '{sheet_name}' not found")))?;
    Ok(ws.page_breaks().cloned().map(XlsxPageBreaks::from_core))
}

pub(super) fn ws_set_page_breaks(
    wb: &Arc<Mutex<CoreWorkbook>>,
    sheet_name: &str,
    breaks: XlsxPageBreaks,
) -> PyResult<()> {
    let mut wb = lock_wb(wb)?;
    let ws = wb
        .sheet_mut(sheet_name)
        .ok_or_else(|| value_error(format!("worksheet '{sheet_name}' not found")))?;
    ws.set_page_breaks(breaks.into_core());
    Ok(())
}

pub(super) fn ws_sheet_view_options(
    wb: &Arc<Mutex<CoreWorkbook>>,
    sheet_name: &str,
) -> PyResult<Option<XlsxSheetViewOptions>> {
    let wb = lock_wb(wb)?;
    let ws = wb
        .sheet(sheet_name)
        .ok_or_else(|| value_error(format!("worksheet '{sheet_name}' not found")))?;
    Ok(ws
        .sheet_view_options()
        .cloned()
        .map(XlsxSheetViewOptions::from_core))
}

pub(super) fn ws_set_sheet_view_options(
    wb: &Arc<Mutex<CoreWorkbook>>,
    sheet_name: &str,
    options: XlsxSheetViewOptions,
) -> PyResult<()> {
    let mut wb = lock_wb(wb)?;
    let ws = wb
        .sheet_mut(sheet_name)
        .ok_or_else(|| value_error(format!("worksheet '{sheet_name}' not found")))?;
    ws.set_sheet_view_options(options.into_core());
    Ok(())
}

// =============================================================================
// Registration
// =============================================================================

pub(super) fn register(module: &Bound<'_, PyModule>) -> PyResult<()> {
    module.add_class::<XlsxPageSetup>()?;
    module.add_class::<XlsxPageMargins>()?;
    module.add_class::<XlsxPrintHeaderFooter>()?;
    module.add_class::<XlsxPageBreaks>()?;
    module.add_class::<XlsxSheetViewOptions>()?;
    Ok(())
}
