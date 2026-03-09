//! Python bindings for the offidized xlsx (Excel) API.

// Sub-module types expose `into_core()` for API completeness even when not yet
// called from this module. Suppress dead_code warnings on those items.
#![allow(dead_code)]

mod chart;
mod comment;
mod conditional;
mod data_validation;
mod finance;
mod image;
mod lint;
mod page;
mod pivot;
mod protection;
mod row_col;
mod sparkline;
mod table;

use crate::error::{value_error, xlsx_error_to_py};
use offidized_xlsx::{
    Alignment as CoreAlignment, Border as CoreBorder, BorderSide as CoreBorderSide, CellValue,
    Fill as CoreFill, Font as CoreFont, HorizontalAlignment, Hyperlink, RichTextRun,
    SheetProtection, SheetVisibility, Style as CoreStyle, VerticalAlignment,
    Workbook as CoreWorkbook,
};
use pyo3::prelude::*;
use std::sync::{Arc, Mutex, MutexGuard};

// =============================================================================
// Helpers
// =============================================================================

pub(super) fn lock_wb(inner: &Arc<Mutex<CoreWorkbook>>) -> PyResult<MutexGuard<'_, CoreWorkbook>> {
    inner
        .lock()
        .map_err(|e| value_error(format!("Failed to lock workbook: {e}")))
}

fn cell_value_to_py(py: Python<'_>, value: Option<&CellValue>) -> PyResult<PyObject> {
    match value {
        None | Some(CellValue::Blank) => Ok(py.None()),
        Some(CellValue::String(s)) => Ok(s.into_pyobject(py)?.into_any().unbind()),
        Some(CellValue::Number(n)) => Ok(n.into_pyobject(py)?.into_any().unbind()),
        Some(CellValue::Bool(b)) => Ok(b.into_pyobject(py)?.to_owned().into_any().unbind()),
        Some(CellValue::Date(s)) => Ok(s.into_pyobject(py)?.into_any().unbind()),
        Some(CellValue::Error(s)) => Ok(s.into_pyobject(py)?.into_any().unbind()),
        Some(CellValue::DateTime(n)) => Ok(n.into_pyobject(py)?.into_any().unbind()),
        Some(CellValue::RichText(runs)) => {
            let text: String = runs.iter().map(|r| r.text()).collect::<Vec<_>>().join("");
            Ok(text.into_pyobject(py)?.into_any().unbind())
        }
    }
}

fn py_to_cell_value(value: &Bound<'_, PyAny>) -> PyResult<CellValue> {
    // Bool MUST come before int — Python bool is a subclass of int.
    if let Ok(b) = value.extract::<bool>() {
        return Ok(CellValue::Bool(b));
    }
    if let Ok(i) = value.extract::<i64>() {
        return Ok(CellValue::Number(i as f64));
    }
    if let Ok(f) = value.extract::<f64>() {
        return Ok(CellValue::Number(f));
    }
    if let Ok(s) = value.extract::<String>() {
        return Ok(CellValue::String(s));
    }
    Err(value_error("Value must be str, int, float, or bool"))
}

fn parse_horizontal_alignment(s: &str) -> PyResult<HorizontalAlignment> {
    match s.to_lowercase().as_str() {
        "general" => Ok(HorizontalAlignment::General),
        "left" => Ok(HorizontalAlignment::Left),
        "center" => Ok(HorizontalAlignment::Center),
        "right" => Ok(HorizontalAlignment::Right),
        "fill" => Ok(HorizontalAlignment::Fill),
        "justify" => Ok(HorizontalAlignment::Justify),
        "centercontinuous" | "center_continuous" => Ok(HorizontalAlignment::CenterContinuous),
        "distributed" => Ok(HorizontalAlignment::Distributed),
        _ => Err(value_error(format!("Unknown horizontal alignment: {s}"))),
    }
}

fn parse_vertical_alignment(s: &str) -> PyResult<VerticalAlignment> {
    match s.to_lowercase().as_str() {
        "top" => Ok(VerticalAlignment::Top),
        "center" => Ok(VerticalAlignment::Center),
        "bottom" => Ok(VerticalAlignment::Bottom),
        "justify" => Ok(VerticalAlignment::Justify),
        "distributed" => Ok(VerticalAlignment::Distributed),
        _ => Err(value_error(format!("Unknown vertical alignment: {s}"))),
    }
}

// =============================================================================
// Workbook
// =============================================================================

/// Python wrapper for `offidized_xlsx::Workbook`.
#[pyclass(module = "offidized._native", name = "Workbook")]
pub struct Workbook {
    inner: Arc<Mutex<CoreWorkbook>>,
}

impl Default for Workbook {
    fn default() -> Self {
        Self::new()
    }
}

#[pymethods]
impl Workbook {
    #[new]
    pub fn new() -> Self {
        Self {
            inner: Arc::new(Mutex::new(CoreWorkbook::new())),
        }
    }

    /// Open an existing workbook from a file path.
    #[staticmethod]
    pub fn open(path: &str) -> PyResult<Self> {
        let workbook = CoreWorkbook::open(path).map_err(xlsx_error_to_py)?;
        Ok(Self {
            inner: Arc::new(Mutex::new(workbook)),
        })
    }

    /// Load a workbook from bytes.
    #[staticmethod]
    pub fn from_bytes(bytes: &[u8]) -> PyResult<Self> {
        let workbook = CoreWorkbook::from_bytes(bytes).map_err(xlsx_error_to_py)?;
        Ok(Self {
            inner: Arc::new(Mutex::new(workbook)),
        })
    }

    /// Serialize the workbook to bytes.
    pub fn to_bytes(&self) -> PyResult<Vec<u8>> {
        let wb = lock_wb(&self.inner)?;
        wb.to_bytes().map_err(xlsx_error_to_py)
    }

    /// Save workbook to a `.xlsx` path.
    pub fn save(&self, path: &str) -> PyResult<()> {
        let wb = lock_wb(&self.inner)?;
        wb.save(path).map_err(xlsx_error_to_py)
    }

    /// Return the names of all worksheets.
    pub fn sheet_names(&self) -> PyResult<Vec<String>> {
        let wb = lock_wb(&self.inner)?;
        Ok(wb.sheet_names().iter().map(|s| s.to_string()).collect())
    }

    /// Get a worksheet by name.
    pub fn sheet(&self, name: &str) -> PyResult<Worksheet> {
        let wb = lock_wb(&self.inner)?;
        if wb.sheet(name).is_none() {
            return Err(value_error(format!("worksheet '{name}' not found")));
        }
        Ok(Worksheet {
            workbook: Arc::clone(&self.inner),
            name_key: name.to_string(),
        })
    }

    /// Add a worksheet by name and return a wrapper for it.
    pub fn add_sheet(&mut self, name: &str) -> PyResult<Worksheet> {
        let mut wb = lock_wb(&self.inner)?;
        wb.add_sheet(name);
        Ok(Worksheet {
            workbook: Arc::clone(&self.inner),
            name_key: name.to_string(),
        })
    }

    /// Remove a worksheet by name. Returns True if it existed.
    pub fn remove_sheet(&mut self, name: &str) -> PyResult<bool> {
        let mut wb = lock_wb(&self.inner)?;
        Ok(wb.remove_sheet(name).is_some())
    }

    /// Return the number of worksheets.
    pub fn sheet_count(&self) -> PyResult<usize> {
        let wb = lock_wb(&self.inner)?;
        Ok(wb.worksheets().len())
    }

    /// Return whether a worksheet with the given name exists.
    pub fn contains_sheet(&self, name: &str) -> PyResult<bool> {
        let wb = lock_wb(&self.inner)?;
        Ok(wb.contains_sheet(name))
    }

    /// Return all defined names as a list of (name, reference) tuples.
    pub fn defined_names(&self) -> PyResult<Vec<(String, String)>> {
        let wb = lock_wb(&self.inner)?;
        Ok(wb
            .defined_names()
            .iter()
            .map(|dn| (dn.name().to_string(), dn.reference().to_string()))
            .collect())
    }

    /// Add a defined name.
    pub fn add_defined_name(&mut self, name: &str, reference: &str) -> PyResult<()> {
        let mut wb = lock_wb(&self.inner)?;
        wb.add_defined_name(name, reference);
        Ok(())
    }

    /// Remove a defined name. Returns True if it existed.
    pub fn remove_defined_name(&mut self, name: &str) -> PyResult<bool> {
        let mut wb = lock_wb(&self.inner)?;
        Ok(wb.remove_defined_name(name).is_some())
    }

    /// Add a style and return its style_id.
    pub fn add_style(&mut self, style: &XlsxStyle) -> PyResult<u32> {
        let mut wb = lock_wb(&self.inner)?;
        wb.add_style(style.to_core()).map_err(xlsx_error_to_py)
    }

    // -- Sub-module delegating methods -----------------------------------------

    // protection
    /// Return workbook-level protection, or None.
    pub fn workbook_protection(&self) -> PyResult<Option<protection::XlsxWorkbookProtection>> {
        protection::wb_workbook_protection(&self.inner)
    }

    /// Set workbook-level protection.
    pub fn set_workbook_protection(
        &mut self,
        prot: protection::XlsxWorkbookProtection,
    ) -> PyResult<()> {
        protection::wb_set_workbook_protection(&self.inner, prot)
    }

    /// Clear workbook-level protection.
    pub fn clear_workbook_protection(&mut self) -> PyResult<()> {
        protection::wb_clear_workbook_protection(&self.inner)
    }

    // lint
    /// Run workbook lint checks. Pass a list of check names or an empty list for all.
    pub fn lint(&self, checks: Vec<String>) -> PyResult<lint::XlsxLintReport> {
        lint::wb_lint(&self.inner, checks)
    }
}

// =============================================================================
// Worksheet
// =============================================================================

/// Python wrapper referencing a worksheet within a workbook by name.
#[pyclass(module = "offidized._native", name = "Worksheet")]
pub struct Worksheet {
    workbook: Arc<Mutex<CoreWorkbook>>,
    name_key: String,
}

#[pymethods]
impl Worksheet {
    /// Return the worksheet name.
    pub fn name(&self) -> PyResult<String> {
        Ok(self.name_key.clone())
    }

    /// Get a cell wrapper for the given reference (e.g. "A1").
    pub fn cell(&self, reference: &str) -> PyResult<XlsxCell> {
        Ok(XlsxCell {
            workbook: Arc::clone(&self.workbook),
            sheet_name: self.name_key.clone(),
            cell_ref: reference.to_string(),
        })
    }

    /// Get the cell value as a Python object (str/int/float/bool/None).
    pub fn cell_value(&self, reference: &str, py: Python<'_>) -> PyResult<PyObject> {
        let wb = lock_wb(&self.workbook)?;
        let ws = wb
            .sheet(&self.name_key)
            .ok_or_else(|| value_error(format!("worksheet '{}' not found", self.name_key)))?;
        let value = ws.cell(reference).and_then(|c| c.value());
        cell_value_to_py(py, value)
    }

    /// Set a cell value from a Python object (str/int/float/bool).
    pub fn set_cell_value(&mut self, reference: &str, value: &Bound<'_, PyAny>) -> PyResult<()> {
        let cv = py_to_cell_value(value)?;
        let mut wb = lock_wb(&self.workbook)?;
        let ws = wb
            .sheet_mut(&self.name_key)
            .ok_or_else(|| value_error(format!("worksheet '{}' not found", self.name_key)))?;
        ws.cell_mut(reference)
            .map_err(xlsx_error_to_py)?
            .set_value(cv);
        Ok(())
    }

    /// Get the formula string for a cell, or None.
    pub fn cell_formula(&self, reference: &str) -> PyResult<Option<String>> {
        let wb = lock_wb(&self.workbook)?;
        let ws = wb
            .sheet(&self.name_key)
            .ok_or_else(|| value_error(format!("worksheet '{}' not found", self.name_key)))?;
        Ok(ws
            .cell(reference)
            .and_then(|c| c.formula())
            .map(|s| s.to_string()))
    }

    /// Set a formula on a cell.
    pub fn set_cell_formula(&mut self, reference: &str, formula: &str) -> PyResult<()> {
        let mut wb = lock_wb(&self.workbook)?;
        let ws = wb
            .sheet_mut(&self.name_key)
            .ok_or_else(|| value_error(format!("worksheet '{}' not found", self.name_key)))?;
        ws.cell_mut(reference)
            .map_err(xlsx_error_to_py)?
            .set_formula(formula);
        Ok(())
    }

    /// Return the list of merged range strings (e.g. ["A1:B2"]).
    pub fn merged_ranges(&self) -> PyResult<Vec<String>> {
        let wb = lock_wb(&self.workbook)?;
        let ws = wb
            .sheet(&self.name_key)
            .ok_or_else(|| value_error(format!("worksheet '{}' not found", self.name_key)))?;
        Ok(ws
            .merged_ranges()
            .iter()
            .map(|r| format!("{}:{}", r.start(), r.end()))
            .collect())
    }

    /// Add a merged range (e.g. "A1:B2").
    pub fn add_merged_range(&mut self, range: &str) -> PyResult<()> {
        let mut wb = lock_wb(&self.workbook)?;
        let ws = wb
            .sheet_mut(&self.name_key)
            .ok_or_else(|| value_error(format!("worksheet '{}' not found", self.name_key)))?;
        ws.add_merged_range(range).map_err(xlsx_error_to_py)?;
        Ok(())
    }

    /// Unmerge a specific range. Returns True if the range was found and removed.
    pub fn unmerge_range(&mut self, range: &str) -> PyResult<bool> {
        let mut wb = lock_wb(&self.workbook)?;
        let ws = wb
            .sheet_mut(&self.name_key)
            .ok_or_else(|| value_error(format!("worksheet '{}' not found", self.name_key)))?;
        ws.unmerge_range(range).map_err(xlsx_error_to_py)
    }

    /// Set freeze panes at the given column/row split.
    pub fn freeze_panes(&mut self, x_split: u32, y_split: u32) -> PyResult<()> {
        let mut wb = lock_wb(&self.workbook)?;
        let ws = wb
            .sheet_mut(&self.name_key)
            .ok_or_else(|| value_error(format!("worksheet '{}' not found", self.name_key)))?;
        ws.set_freeze_panes(x_split, y_split)
            .map_err(xlsx_error_to_py)?;
        Ok(())
    }

    /// Clear freeze panes.
    pub fn clear_freeze_panes(&mut self) -> PyResult<()> {
        let mut wb = lock_wb(&self.workbook)?;
        let ws = wb
            .sheet_mut(&self.name_key)
            .ok_or_else(|| value_error(format!("worksheet '{}' not found", self.name_key)))?;
        ws.clear_freeze_pane();
        Ok(())
    }

    /// Set an auto-filter on the given range (e.g. "A1:D10").
    pub fn set_auto_filter(&mut self, range: &str) -> PyResult<()> {
        let mut wb = lock_wb(&self.workbook)?;
        let ws = wb
            .sheet_mut(&self.name_key)
            .ok_or_else(|| value_error(format!("worksheet '{}' not found", self.name_key)))?;
        ws.set_auto_filter(range).map_err(xlsx_error_to_py)?;
        Ok(())
    }

    /// Clear the auto-filter.
    pub fn clear_auto_filter(&mut self) -> PyResult<()> {
        let mut wb = lock_wb(&self.workbook)?;
        let ws = wb
            .sheet_mut(&self.name_key)
            .ok_or_else(|| value_error(format!("worksheet '{}' not found", self.name_key)))?;
        ws.clear_auto_filter();
        Ok(())
    }

    /// Return the list of table names in this worksheet.
    pub fn tables(&self) -> PyResult<Vec<String>> {
        let wb = lock_wb(&self.workbook)?;
        let ws = wb
            .sheet(&self.name_key)
            .ok_or_else(|| value_error(format!("worksheet '{}' not found", self.name_key)))?;
        Ok(ws.tables().iter().map(|t| t.name().to_string()).collect())
    }

    /// Return hyperlinks as a list of (cell_ref, url) tuples.
    pub fn hyperlinks(&self) -> PyResult<Vec<(String, String)>> {
        let wb = lock_wb(&self.workbook)?;
        let ws = wb
            .sheet(&self.name_key)
            .ok_or_else(|| value_error(format!("worksheet '{}' not found", self.name_key)))?;
        Ok(ws
            .hyperlinks()
            .iter()
            .filter_map(|h| {
                h.url()
                    .map(|url| (h.cell_ref().to_string(), url.to_string()))
            })
            .collect())
    }

    /// Add an external hyperlink to a cell.
    pub fn add_hyperlink(&mut self, cell_ref: &str, url: &str) -> PyResult<()> {
        let mut wb = lock_wb(&self.workbook)?;
        let ws = wb
            .sheet_mut(&self.name_key)
            .ok_or_else(|| value_error(format!("worksheet '{}' not found", self.name_key)))?;
        let hyperlink = Hyperlink::external(cell_ref, url).map_err(xlsx_error_to_py)?;
        ws.add_hyperlink(hyperlink);
        Ok(())
    }

    /// Find all cell references containing the given text.
    pub fn find_cells(&self, text: &str) -> PyResult<Vec<String>> {
        let wb = lock_wb(&self.workbook)?;
        let ws = wb
            .sheet(&self.name_key)
            .ok_or_else(|| value_error(format!("worksheet '{}' not found", self.name_key)))?;
        Ok(ws.find_cells(text))
    }

    /// Return the tab color as a hex RGB string, or None.
    pub fn tab_color(&self) -> PyResult<Option<String>> {
        let wb = lock_wb(&self.workbook)?;
        let ws = wb
            .sheet(&self.name_key)
            .ok_or_else(|| value_error(format!("worksheet '{}' not found", self.name_key)))?;
        Ok(ws.tab_color().map(|s| s.to_string()))
    }

    /// Set the tab color as a hex RGB string (e.g. "FF0000").
    pub fn set_tab_color(&mut self, color: &str) -> PyResult<()> {
        let mut wb = lock_wb(&self.workbook)?;
        let ws = wb
            .sheet_mut(&self.name_key)
            .ok_or_else(|| value_error(format!("worksheet '{}' not found", self.name_key)))?;
        ws.set_tab_color(color);
        Ok(())
    }

    /// Return whether the sheet is protected.
    pub fn protection(&self) -> PyResult<bool> {
        let wb = lock_wb(&self.workbook)?;
        let ws = wb
            .sheet(&self.name_key)
            .ok_or_else(|| value_error(format!("worksheet '{}' not found", self.name_key)))?;
        Ok(ws.protection().is_some())
    }

    /// Set or clear sheet protection.
    pub fn set_protection(&mut self, protected: bool) -> PyResult<()> {
        let mut wb = lock_wb(&self.workbook)?;
        let ws = wb
            .sheet_mut(&self.name_key)
            .ok_or_else(|| value_error(format!("worksheet '{}' not found", self.name_key)))?;
        if protected {
            ws.set_protection(SheetProtection::default());
        } else {
            ws.clear_protection();
        }
        Ok(())
    }

    /// Get sheet visibility: "visible", "hidden", or "veryHidden".
    pub fn visibility(&self) -> PyResult<String> {
        let wb = lock_wb(&self.workbook)?;
        let ws = wb
            .sheet(&self.name_key)
            .ok_or_else(|| value_error(format!("worksheet '{}' not found", self.name_key)))?;
        Ok(match ws.visibility() {
            SheetVisibility::Visible => "visible",
            SheetVisibility::Hidden => "hidden",
            SheetVisibility::VeryHidden => "veryHidden",
        }
        .to_owned())
    }

    /// Set sheet visibility: "visible", "hidden", or "veryHidden".
    pub fn set_visibility(&self, visibility: &str) -> PyResult<()> {
        let vis = match visibility {
            "visible" => SheetVisibility::Visible,
            "hidden" => SheetVisibility::Hidden,
            "veryHidden" | "very_hidden" => SheetVisibility::VeryHidden,
            _ => return Err(value_error(format!("Unknown visibility: {visibility}"))),
        };
        let mut wb = lock_wb(&self.workbook)?;
        let ws = wb
            .sheet_mut(&self.name_key)
            .ok_or_else(|| value_error(format!("worksheet '{}' not found", self.name_key)))?;
        ws.set_visibility(vis);
        Ok(())
    }

    // -- Sub-module delegating methods -----------------------------------------

    // row_col
    /// Return a row wrapper for the given 1-based row index.
    pub fn row(&self, index: u32) -> row_col::XlsxRow {
        row_col::XlsxRow::new(Arc::clone(&self.workbook), self.name_key.clone(), index)
    }

    /// Return a column wrapper for the given 1-based column index.
    pub fn column(&self, index: u32) -> row_col::XlsxColumn {
        row_col::XlsxColumn::new(Arc::clone(&self.workbook), self.name_key.clone(), index)
    }

    // comment
    /// Return all comments as a list of XlsxComment objects.
    pub fn comments(&self) -> PyResult<Vec<comment::XlsxComment>> {
        comment::ws_comments(&self.workbook, &self.name_key)
    }

    /// Add a comment to a cell.
    pub fn add_comment(&mut self, cmt: comment::XlsxComment) -> PyResult<()> {
        comment::ws_add_comment(&self.workbook, &self.name_key, &cmt)
    }

    /// Remove a comment from a cell. Returns True if it existed.
    pub fn remove_comment(&mut self, cell_ref: &str) -> PyResult<bool> {
        comment::ws_remove_comment(&self.workbook, &self.name_key, cell_ref)
    }

    /// Clear all comments from this worksheet.
    pub fn clear_comments(&mut self) -> PyResult<()> {
        comment::ws_clear_comments(&self.workbook, &self.name_key)
    }

    // protection (detailed)
    /// Return the detailed sheet protection settings, or None.
    pub fn protection_detail(&self) -> PyResult<Option<protection::XlsxSheetProtection>> {
        protection::ws_protection_detail(&self.workbook, &self.name_key)
    }

    /// Set detailed sheet protection settings.
    pub fn set_protection_detail(&mut self, prot: protection::XlsxSheetProtection) -> PyResult<()> {
        protection::ws_set_protection_detail(&self.workbook, &self.name_key, prot)
    }

    // page setup
    /// Return the page setup, or None.
    pub fn page_setup(&self) -> PyResult<Option<page::XlsxPageSetup>> {
        page::ws_page_setup(&self.workbook, &self.name_key)
    }

    /// Set page setup.
    pub fn set_page_setup(&mut self, setup: page::XlsxPageSetup) -> PyResult<()> {
        page::ws_set_page_setup(&self.workbook, &self.name_key, setup)
    }

    /// Return the page margins, or None.
    pub fn page_margins(&self) -> PyResult<Option<page::XlsxPageMargins>> {
        page::ws_page_margins(&self.workbook, &self.name_key)
    }

    /// Set page margins.
    pub fn set_page_margins(&mut self, margins: page::XlsxPageMargins) -> PyResult<()> {
        page::ws_set_page_margins(&self.workbook, &self.name_key, margins)
    }

    /// Return the header/footer settings, or None.
    pub fn header_footer(&self) -> PyResult<Option<page::XlsxPrintHeaderFooter>> {
        page::ws_header_footer(&self.workbook, &self.name_key)
    }

    /// Set the header/footer settings.
    pub fn set_header_footer(&mut self, hf: page::XlsxPrintHeaderFooter) -> PyResult<()> {
        page::ws_set_header_footer(&self.workbook, &self.name_key, hf)
    }

    /// Return the print area string, or None.
    pub fn print_area(&self) -> PyResult<Option<String>> {
        page::ws_print_area(&self.workbook, &self.name_key)
    }

    /// Set the print area (e.g. "A1:D10").
    pub fn set_print_area(&mut self, range: &str) -> PyResult<()> {
        page::ws_set_print_area(&self.workbook, &self.name_key, range)
    }

    /// Clear the print area.
    pub fn clear_print_area(&mut self) -> PyResult<()> {
        page::ws_clear_print_area(&self.workbook, &self.name_key)
    }

    /// Return the page breaks, or None.
    pub fn page_breaks(&self) -> PyResult<Option<page::XlsxPageBreaks>> {
        page::ws_page_breaks(&self.workbook, &self.name_key)
    }

    /// Set page breaks.
    pub fn set_page_breaks(&mut self, breaks: page::XlsxPageBreaks) -> PyResult<()> {
        page::ws_set_page_breaks(&self.workbook, &self.name_key, breaks)
    }

    /// Return the sheet view options, or None.
    pub fn sheet_view_options(&self) -> PyResult<Option<page::XlsxSheetViewOptions>> {
        page::ws_sheet_view_options(&self.workbook, &self.name_key)
    }

    /// Set sheet view options.
    pub fn set_sheet_view_options(&mut self, options: page::XlsxSheetViewOptions) -> PyResult<()> {
        page::ws_set_sheet_view_options(&self.workbook, &self.name_key, options)
    }

    // images
    /// Return all images as a list of XlsxWorksheetImage objects.
    pub fn images(&self) -> PyResult<Vec<image::XlsxWorksheetImage>> {
        image::ws_images(&self.workbook, &self.name_key)
    }

    /// Add an image to this worksheet.
    pub fn add_image(&mut self, img: image::XlsxWorksheetImage) -> PyResult<()> {
        image::ws_add_image(&self.workbook, &self.name_key, img)
    }

    /// Clear all images from this worksheet.
    pub fn clear_images(&mut self) -> PyResult<()> {
        image::ws_clear_images(&self.workbook, &self.name_key)
    }

    // data validation
    /// Return all data validations.
    pub fn data_validations(&self) -> PyResult<Vec<data_validation::XlsxDataValidation>> {
        data_validation::ws_data_validations(&self.workbook, &self.name_key)
    }

    /// Add a data validation.
    pub fn add_data_validation(&mut self, dv: data_validation::XlsxDataValidation) -> PyResult<()> {
        data_validation::ws_add_data_validation(&self.workbook, &self.name_key, dv)
    }

    /// Clear all data validations.
    pub fn clear_data_validations(&mut self) -> PyResult<()> {
        data_validation::ws_clear_data_validations(&self.workbook, &self.name_key)
    }

    // tables (rich)
    /// Return all worksheet tables as XlsxWorksheetTable objects.
    pub fn worksheet_tables(&self) -> PyResult<Vec<table::XlsxWorksheetTable>> {
        table::ws_tables(&self.workbook, &self.name_key)
    }

    /// Add a worksheet table.
    pub fn add_table(&mut self, tbl: table::XlsxWorksheetTable) -> PyResult<()> {
        table::ws_add_table(&self.workbook, &self.name_key, tbl)
    }

    // sparklines
    /// Return all sparkline groups.
    pub fn sparkline_groups(&self) -> PyResult<Vec<sparkline::XlsxSparklineGroup>> {
        sparkline::ws_sparkline_groups(&self.workbook, &self.name_key)
    }

    /// Add a sparkline group.
    pub fn add_sparkline_group(&mut self, group: sparkline::XlsxSparklineGroup) -> PyResult<()> {
        sparkline::ws_add_sparkline_group(&self.workbook, &self.name_key, group)
    }

    /// Clear all sparkline groups.
    pub fn clear_sparkline_groups(&mut self) -> PyResult<()> {
        sparkline::ws_clear_sparkline_groups(&self.workbook, &self.name_key)
    }

    // conditional formatting
    /// Return all conditional formattings.
    pub fn conditional_formattings(&self) -> PyResult<Vec<conditional::XlsxConditionalFormatting>> {
        conditional::ws_conditional_formattings(&self.workbook, &self.name_key)
    }

    /// Add a conditional formatting.
    pub fn add_conditional_formatting(
        &mut self,
        cf: conditional::XlsxConditionalFormatting,
    ) -> PyResult<()> {
        conditional::ws_add_conditional_formatting(&self.workbook, &self.name_key, cf)
    }

    /// Clear all conditional formattings.
    pub fn clear_conditional_formattings(&mut self) -> PyResult<()> {
        conditional::ws_clear_conditional_formattings(&self.workbook, &self.name_key)
    }

    // charts
    /// Return all charts.
    pub fn charts(&self) -> PyResult<Vec<chart::XlsxChart>> {
        chart::ws_charts(&self.workbook, &self.name_key)
    }

    /// Add a chart.
    pub fn add_chart(&mut self, c: chart::XlsxChart) -> PyResult<()> {
        chart::ws_add_chart(&self.workbook, &self.name_key, c)
    }

    /// Clear all charts.
    pub fn clear_charts(&mut self) -> PyResult<()> {
        chart::ws_clear_charts(&self.workbook, &self.name_key)
    }

    // pivot tables
    /// Return all pivot tables.
    pub fn pivot_tables(&self) -> PyResult<Vec<pivot::XlsxPivotTable>> {
        pivot::ws_pivot_tables(&self.workbook, &self.name_key)
    }

    /// Add a pivot table from a pre-built object with an A1 target reference.
    #[pyo3(signature = (pt, target = None))]
    pub fn add_pivot_table(
        &mut self,
        pt: pivot::XlsxPivotTable,
        target: Option<&str>,
    ) -> PyResult<()> {
        pivot::ws_add_pivot_table_obj(&self.workbook, &self.name_key, pt, target)
    }

    /// Clear all pivot tables.
    pub fn clear_pivot_tables(&mut self) -> PyResult<()> {
        pivot::ws_clear_pivot_tables(&self.workbook, &self.name_key)
    }
}

// =============================================================================
// XlsxCell
// =============================================================================

/// Python wrapper referencing a single cell within a workbook.
#[pyclass(module = "offidized._native", name = "XlsxCell")]
pub struct XlsxCell {
    workbook: Arc<Mutex<CoreWorkbook>>,
    sheet_name: String,
    cell_ref: String,
}

#[pymethods]
impl XlsxCell {
    /// Get the cell value as a Python object (str/int/float/bool/None).
    pub fn value(&self, py: Python<'_>) -> PyResult<PyObject> {
        let wb = lock_wb(&self.workbook)?;
        let ws = wb
            .sheet(&self.sheet_name)
            .ok_or_else(|| value_error(format!("worksheet '{}' not found", self.sheet_name)))?;
        let value = ws.cell(&self.cell_ref).and_then(|c| c.value());
        cell_value_to_py(py, value)
    }

    /// Set the cell value from a Python object (str/int/float/bool).
    pub fn set_value(&mut self, value: &Bound<'_, PyAny>) -> PyResult<()> {
        let cv = py_to_cell_value(value)?;
        let mut wb = lock_wb(&self.workbook)?;
        let ws = wb
            .sheet_mut(&self.sheet_name)
            .ok_or_else(|| value_error(format!("worksheet '{}' not found", self.sheet_name)))?;
        ws.cell_mut(&self.cell_ref)
            .map_err(xlsx_error_to_py)?
            .set_value(cv);
        Ok(())
    }

    /// Get the formula string, or None.
    pub fn formula(&self) -> PyResult<Option<String>> {
        let wb = lock_wb(&self.workbook)?;
        let ws = wb
            .sheet(&self.sheet_name)
            .ok_or_else(|| value_error(format!("worksheet '{}' not found", self.sheet_name)))?;
        Ok(ws
            .cell(&self.cell_ref)
            .and_then(|c| c.formula())
            .map(|s| s.to_string()))
    }

    /// Set a formula on this cell.
    pub fn set_formula(&mut self, formula: &str) -> PyResult<()> {
        let mut wb = lock_wb(&self.workbook)?;
        let ws = wb
            .sheet_mut(&self.sheet_name)
            .ok_or_else(|| value_error(format!("worksheet '{}' not found", self.sheet_name)))?;
        ws.cell_mut(&self.cell_ref)
            .map_err(xlsx_error_to_py)?
            .set_formula(formula);
        Ok(())
    }

    /// Clear both value and formula.
    pub fn clear(&mut self) -> PyResult<()> {
        let mut wb = lock_wb(&self.workbook)?;
        let ws = wb
            .sheet_mut(&self.sheet_name)
            .ok_or_else(|| value_error(format!("worksheet '{}' not found", self.sheet_name)))?;
        ws.cell_mut(&self.cell_ref)
            .map_err(xlsx_error_to_py)?
            .clear();
        Ok(())
    }

    /// Get the style ID, or None.
    pub fn style_id(&self) -> PyResult<Option<u32>> {
        let wb = lock_wb(&self.workbook)?;
        let ws = wb
            .sheet(&self.sheet_name)
            .ok_or_else(|| value_error(format!("worksheet '{}' not found", self.sheet_name)))?;
        Ok(ws.cell(&self.cell_ref).and_then(|c| c.style_id()))
    }

    /// Set the style ID on this cell.
    pub fn set_style_id(&mut self, style_id: u32) -> PyResult<()> {
        let mut wb = lock_wb(&self.workbook)?;
        let ws = wb
            .sheet_mut(&self.sheet_name)
            .ok_or_else(|| value_error(format!("worksheet '{}' not found", self.sheet_name)))?;
        ws.cell_mut(&self.cell_ref)
            .map_err(xlsx_error_to_py)?
            .set_style_id(style_id);
        Ok(())
    }

    /// Return the cell reference string (e.g. "A1").
    pub fn reference(&self) -> String {
        self.cell_ref.clone()
    }

    /// Return whether the cell has a formula.
    pub fn has_formula(&self) -> PyResult<bool> {
        let wb = lock_wb(&self.workbook)?;
        let ws = wb
            .sheet(&self.sheet_name)
            .ok_or_else(|| value_error(format!("worksheet '{}' not found", self.sheet_name)))?;
        Ok(ws.cell(&self.cell_ref).and_then(|c| c.formula()).is_some())
    }

    /// Get rich text runs, or None if the cell doesn't contain rich text.
    pub fn rich_text(&self) -> PyResult<Option<Vec<XlsxRichTextRun>>> {
        let wb = lock_wb(&self.workbook)?;
        let ws = wb
            .sheet(&self.sheet_name)
            .ok_or_else(|| value_error(format!("worksheet '{}' not found", self.sheet_name)))?;
        let cell = ws.cell(&self.cell_ref);
        Ok(cell.and_then(|c| {
            c.rich_text().map(|runs| {
                runs.iter()
                    .map(|r| XlsxRichTextRun { inner: r.clone() })
                    .collect()
            })
        }))
    }

    /// Set cell value to rich text.
    pub fn set_rich_text(&self, runs: Vec<XlsxRichTextRun>) -> PyResult<()> {
        let mut wb = lock_wb(&self.workbook)?;
        let ws = wb
            .sheet_mut(&self.sheet_name)
            .ok_or_else(|| value_error(format!("worksheet '{}' not found", self.sheet_name)))?;
        let core_runs: Vec<RichTextRun> = runs.into_iter().map(|r| r.inner).collect();
        ws.cell_mut(&self.cell_ref)
            .map_err(xlsx_error_to_py)?
            .set_rich_text(core_runs);
        Ok(())
    }
}

// =============================================================================
// XlsxRichTextRun
// =============================================================================

/// Python wrapper for a rich text run within a cell.
#[pyclass(module = "offidized._native", name = "XlsxRichTextRun")]
#[derive(Clone)]
pub struct XlsxRichTextRun {
    inner: RichTextRun,
}

#[pymethods]
impl XlsxRichTextRun {
    /// Create a new rich text run with the given text and optional formatting.
    #[new]
    #[pyo3(signature = (text, bold=None, italic=None, font_name=None, font_size=None, color=None))]
    pub fn new(
        text: &str,
        bold: Option<bool>,
        italic: Option<bool>,
        font_name: Option<&str>,
        font_size: Option<&str>,
        color: Option<&str>,
    ) -> Self {
        let mut run = RichTextRun::new(text);
        if let Some(b) = bold {
            run.set_bold(b);
        }
        if let Some(i) = italic {
            run.set_italic(i);
        }
        if let Some(f) = font_name {
            run.set_font_name(f);
        }
        if let Some(s) = font_size {
            run.set_font_size(s);
        }
        if let Some(c) = color {
            run.set_color(c);
        }
        Self { inner: run }
    }

    /// The text content of this run.
    #[getter]
    pub fn text(&self) -> &str {
        self.inner.text()
    }

    /// Set the text content.
    #[setter]
    pub fn set_text(&mut self, text: &str) {
        self.inner.set_text(text);
    }

    /// Whether this run is bold (None if unset).
    #[getter]
    pub fn bold(&self) -> Option<bool> {
        self.inner.bold()
    }

    /// Set bold formatting.
    #[setter]
    pub fn set_bold(&mut self, bold: bool) {
        self.inner.set_bold(bold);
    }

    /// Whether this run is italic (None if unset).
    #[getter]
    pub fn italic(&self) -> Option<bool> {
        self.inner.italic()
    }

    /// Set italic formatting.
    #[setter]
    pub fn set_italic(&mut self, italic: bool) {
        self.inner.set_italic(italic);
    }

    /// The font name (None if unset).
    #[getter]
    pub fn font_name(&self) -> Option<&str> {
        self.inner.font_name()
    }

    /// Set the font name.
    #[setter]
    pub fn set_font_name(&mut self, name: &str) {
        self.inner.set_font_name(name);
    }

    /// The font size (None if unset).
    #[getter]
    pub fn font_size(&self) -> Option<&str> {
        self.inner.font_size()
    }

    /// Set the font size.
    #[setter]
    pub fn set_font_size(&mut self, size: &str) {
        self.inner.set_font_size(size);
    }

    /// The color (None if unset).
    #[getter]
    pub fn color(&self) -> Option<&str> {
        self.inner.color()
    }

    /// Set the color.
    #[setter]
    pub fn set_color(&mut self, color: &str) {
        self.inner.set_color(color);
    }
}

// =============================================================================
// XlsxStyle
// =============================================================================

/// Python wrapper for a cell style (value type, not a reference).
#[pyclass(module = "offidized._native", name = "XlsxStyle")]
#[derive(Clone, Default)]
pub struct XlsxStyle {
    number_format: Option<String>,
    font: Option<XlsxFont>,
    fill: Option<XlsxFill>,
    border: Option<XlsxBorder>,
    alignment: Option<XlsxAlignment>,
}

impl XlsxStyle {
    fn to_core(&self) -> CoreStyle {
        let mut style = CoreStyle::new();
        if let Some(ref fmt) = self.number_format {
            style.set_number_format(fmt.as_str());
        }
        if let Some(ref font) = self.font {
            style.set_font(font.to_core());
        }
        if let Some(ref fill) = self.fill {
            style.set_fill(fill.to_core());
        }
        if let Some(ref border) = self.border {
            style.set_border(border.to_core());
        }
        if let Some(ref alignment) = self.alignment {
            style.set_alignment(alignment.to_core());
        }
        style
    }
}

#[allow(clippy::new_without_default)]
#[pymethods]
impl XlsxStyle {
    #[new]
    pub fn new() -> Self {
        Self::default()
    }

    /// Get the number format string, or None.
    pub fn number_format(&self) -> Option<String> {
        self.number_format.clone()
    }

    /// Set the number format string.
    pub fn set_number_format(&mut self, fmt: &str) {
        self.number_format = Some(fmt.to_string());
    }

    /// Get the font, or a default XlsxFont if unset.
    pub fn font(&self) -> XlsxFont {
        self.font.clone().unwrap_or_default()
    }

    /// Set the font.
    pub fn set_font(&mut self, font: &XlsxFont) {
        self.font = Some(font.clone());
    }

    /// Get the fill, or a default XlsxFill if unset.
    pub fn fill(&self) -> XlsxFill {
        self.fill.clone().unwrap_or_default()
    }

    /// Set the fill.
    pub fn set_fill(&mut self, fill: &XlsxFill) {
        self.fill = Some(fill.clone());
    }

    /// Get the border, or a default XlsxBorder if unset.
    pub fn border(&self) -> XlsxBorder {
        self.border.clone().unwrap_or_default()
    }

    /// Set the border.
    pub fn set_border(&mut self, border: &XlsxBorder) {
        self.border = Some(border.clone());
    }

    /// Get the alignment, or a default XlsxAlignment if unset.
    pub fn alignment(&self) -> XlsxAlignment {
        self.alignment.clone().unwrap_or_default()
    }

    /// Set the alignment.
    pub fn set_alignment(&mut self, alignment: &XlsxAlignment) {
        self.alignment = Some(alignment.clone());
    }
}

// =============================================================================
// XlsxFont
// =============================================================================

/// Python wrapper for font styling.
#[pyclass(module = "offidized._native", name = "XlsxFont")]
#[derive(Clone, Default)]
pub struct XlsxFont {
    name: Option<String>,
    size: Option<String>,
    bold: Option<bool>,
    italic: Option<bool>,
    underline: Option<bool>,
    color: Option<String>,
    strikethrough: Option<bool>,
}

impl XlsxFont {
    fn to_core(&self) -> CoreFont {
        let mut font = CoreFont::new();
        if let Some(ref name) = self.name {
            font.set_name(name.as_str());
        }
        if let Some(ref size) = self.size {
            font.set_size(size.as_str());
        }
        if let Some(bold) = self.bold {
            font.set_bold(bold);
        }
        if let Some(italic) = self.italic {
            font.set_italic(italic);
        }
        if let Some(underline) = self.underline {
            font.set_underline(underline);
        }
        if let Some(ref color) = self.color {
            font.set_color(color.as_str());
        }
        if let Some(strikethrough) = self.strikethrough {
            font.set_strikethrough(strikethrough);
        }
        font
    }
}

#[allow(clippy::new_without_default)]
#[pymethods]
impl XlsxFont {
    #[new]
    pub fn new() -> Self {
        Self::default()
    }

    /// Get the font name, or None.
    #[getter]
    pub fn name(&self) -> Option<String> {
        self.name.clone()
    }

    /// Set the font name.
    #[setter]
    pub fn set_name(&mut self, name: Option<String>) {
        self.name = name;
    }

    /// Get the font size as a string (e.g. "11"), or None.
    #[getter]
    pub fn size(&self) -> Option<String> {
        self.size.clone()
    }

    /// Set the font size as a string (e.g. "11").
    #[setter]
    pub fn set_size(&mut self, size: Option<String>) {
        self.size = size;
    }

    /// Get bold flag, or None.
    #[getter]
    pub fn bold(&self) -> Option<bool> {
        self.bold
    }

    /// Set bold flag.
    #[setter]
    pub fn set_bold(&mut self, bold: Option<bool>) {
        self.bold = bold;
    }

    /// Get italic flag, or None.
    #[getter]
    pub fn italic(&self) -> Option<bool> {
        self.italic
    }

    /// Set italic flag.
    #[setter]
    pub fn set_italic(&mut self, italic: Option<bool>) {
        self.italic = italic;
    }

    /// Get underline flag, or None.
    #[getter]
    pub fn underline(&self) -> Option<bool> {
        self.underline
    }

    /// Set underline flag.
    #[setter]
    pub fn set_underline(&mut self, underline: Option<bool>) {
        self.underline = underline;
    }

    /// Get the font color as an ARGB hex string, or None.
    #[getter]
    pub fn color(&self) -> Option<String> {
        self.color.clone()
    }

    /// Set the font color as an ARGB hex string (e.g. "FFFF0000").
    #[setter]
    pub fn set_color(&mut self, color: Option<String>) {
        self.color = color;
    }

    /// Get strikethrough flag, or None.
    #[getter]
    pub fn strikethrough(&self) -> Option<bool> {
        self.strikethrough
    }

    /// Set strikethrough flag.
    #[setter]
    pub fn set_strikethrough(&mut self, strikethrough: Option<bool>) {
        self.strikethrough = strikethrough;
    }
}

// =============================================================================
// XlsxFill
// =============================================================================

/// Python wrapper for fill styling.
#[pyclass(module = "offidized._native", name = "XlsxFill")]
#[derive(Clone, Default)]
pub struct XlsxFill {
    pattern: Option<String>,
    foreground_color: Option<String>,
    background_color: Option<String>,
}

impl XlsxFill {
    fn to_core(&self) -> CoreFill {
        let mut fill = CoreFill::new();
        if let Some(ref pattern) = self.pattern {
            fill.set_pattern(pattern.as_str());
        }
        if let Some(ref fg) = self.foreground_color {
            fill.set_foreground_color(fg.as_str());
        }
        if let Some(ref bg) = self.background_color {
            fill.set_background_color(bg.as_str());
        }
        fill
    }
}

#[allow(clippy::new_without_default)]
#[pymethods]
impl XlsxFill {
    #[new]
    pub fn new() -> Self {
        Self::default()
    }

    /// Get the fill pattern (e.g. "solid"), or None.
    #[getter]
    pub fn pattern(&self) -> Option<String> {
        self.pattern.clone()
    }

    /// Set the fill pattern.
    #[setter]
    pub fn set_pattern(&mut self, pattern: Option<String>) {
        self.pattern = pattern;
    }

    /// Get the foreground color as ARGB hex, or None.
    #[getter]
    pub fn foreground_color(&self) -> Option<String> {
        self.foreground_color.clone()
    }

    /// Set the foreground color as ARGB hex.
    #[setter]
    pub fn set_foreground_color(&mut self, color: Option<String>) {
        self.foreground_color = color;
    }

    /// Get the background color as ARGB hex, or None.
    #[getter]
    pub fn background_color(&self) -> Option<String> {
        self.background_color.clone()
    }

    /// Set the background color as ARGB hex.
    #[setter]
    pub fn set_background_color(&mut self, color: Option<String>) {
        self.background_color = color;
    }
}

// =============================================================================
// XlsxBorder
// =============================================================================

/// Python wrapper for border styling.
#[pyclass(module = "offidized._native", name = "XlsxBorder")]
#[derive(Clone, Default)]
pub struct XlsxBorder {
    left_style: Option<String>,
    left_color: Option<String>,
    right_style: Option<String>,
    right_color: Option<String>,
    top_style: Option<String>,
    top_color: Option<String>,
    bottom_style: Option<String>,
    bottom_color: Option<String>,
}

impl XlsxBorder {
    fn to_core(&self) -> CoreBorder {
        let mut border = CoreBorder::new();
        if self.left_style.is_some() || self.left_color.is_some() {
            let mut side = CoreBorderSide::new();
            if let Some(ref s) = self.left_style {
                side.set_style(s.as_str());
            }
            if let Some(ref c) = self.left_color {
                side.set_color(c.as_str());
            }
            border.set_left(side);
        }
        if self.right_style.is_some() || self.right_color.is_some() {
            let mut side = CoreBorderSide::new();
            if let Some(ref s) = self.right_style {
                side.set_style(s.as_str());
            }
            if let Some(ref c) = self.right_color {
                side.set_color(c.as_str());
            }
            border.set_right(side);
        }
        if self.top_style.is_some() || self.top_color.is_some() {
            let mut side = CoreBorderSide::new();
            if let Some(ref s) = self.top_style {
                side.set_style(s.as_str());
            }
            if let Some(ref c) = self.top_color {
                side.set_color(c.as_str());
            }
            border.set_top(side);
        }
        if self.bottom_style.is_some() || self.bottom_color.is_some() {
            let mut side = CoreBorderSide::new();
            if let Some(ref s) = self.bottom_style {
                side.set_style(s.as_str());
            }
            if let Some(ref c) = self.bottom_color {
                side.set_color(c.as_str());
            }
            border.set_bottom(side);
        }
        border
    }
}

#[allow(clippy::new_without_default)]
#[pymethods]
impl XlsxBorder {
    #[new]
    pub fn new() -> Self {
        Self::default()
    }

    #[getter]
    pub fn left_style(&self) -> Option<String> {
        self.left_style.clone()
    }

    #[setter]
    pub fn set_left_style(&mut self, style: Option<String>) {
        self.left_style = style;
    }

    #[getter]
    pub fn left_color(&self) -> Option<String> {
        self.left_color.clone()
    }

    #[setter]
    pub fn set_left_color(&mut self, color: Option<String>) {
        self.left_color = color;
    }

    #[getter]
    pub fn right_style(&self) -> Option<String> {
        self.right_style.clone()
    }

    #[setter]
    pub fn set_right_style(&mut self, style: Option<String>) {
        self.right_style = style;
    }

    #[getter]
    pub fn right_color(&self) -> Option<String> {
        self.right_color.clone()
    }

    #[setter]
    pub fn set_right_color(&mut self, color: Option<String>) {
        self.right_color = color;
    }

    #[getter]
    pub fn top_style(&self) -> Option<String> {
        self.top_style.clone()
    }

    #[setter]
    pub fn set_top_style(&mut self, style: Option<String>) {
        self.top_style = style;
    }

    #[getter]
    pub fn top_color(&self) -> Option<String> {
        self.top_color.clone()
    }

    #[setter]
    pub fn set_top_color(&mut self, color: Option<String>) {
        self.top_color = color;
    }

    #[getter]
    pub fn bottom_style(&self) -> Option<String> {
        self.bottom_style.clone()
    }

    #[setter]
    pub fn set_bottom_style(&mut self, style: Option<String>) {
        self.bottom_style = style;
    }

    #[getter]
    pub fn bottom_color(&self) -> Option<String> {
        self.bottom_color.clone()
    }

    #[setter]
    pub fn set_bottom_color(&mut self, color: Option<String>) {
        self.bottom_color = color;
    }
}

// =============================================================================
// XlsxAlignment
// =============================================================================

/// Python wrapper for cell alignment.
#[pyclass(module = "offidized._native", name = "XlsxAlignment")]
#[derive(Clone, Default)]
pub struct XlsxAlignment {
    horizontal: Option<String>,
    vertical: Option<String>,
    wrap_text: Option<bool>,
    text_rotation: Option<u32>,
}

impl XlsxAlignment {
    fn to_core(&self) -> CoreAlignment {
        let mut alignment = CoreAlignment::new();
        if let Some(ref h) = self.horizontal {
            // Best-effort: if parsing fails, silently skip
            if let Ok(ha) = parse_horizontal_alignment(h) {
                alignment.set_horizontal(ha);
            }
        }
        if let Some(ref v) = self.vertical {
            if let Ok(va) = parse_vertical_alignment(v) {
                alignment.set_vertical(va);
            }
        }
        if let Some(wrap) = self.wrap_text {
            alignment.set_wrap_text(wrap);
        }
        if let Some(rotation) = self.text_rotation {
            alignment.set_text_rotation(rotation);
        }
        alignment
    }
}

#[allow(clippy::new_without_default)]
#[pymethods]
impl XlsxAlignment {
    #[new]
    pub fn new() -> Self {
        Self::default()
    }

    /// Get horizontal alignment as a string (e.g. "left", "center"), or None.
    #[getter]
    pub fn horizontal(&self) -> Option<String> {
        self.horizontal.clone()
    }

    /// Set horizontal alignment (e.g. "left", "center", "right", "fill", "justify", "distributed").
    #[setter]
    pub fn set_horizontal(&mut self, value: Option<String>) {
        self.horizontal = value;
    }

    /// Get vertical alignment as a string (e.g. "top", "center"), or None.
    #[getter]
    pub fn vertical(&self) -> Option<String> {
        self.vertical.clone()
    }

    /// Set vertical alignment (e.g. "top", "center", "bottom", "justify", "distributed").
    #[setter]
    pub fn set_vertical(&mut self, value: Option<String>) {
        self.vertical = value;
    }

    /// Get wrap text flag, or None.
    #[getter]
    pub fn wrap_text(&self) -> Option<bool> {
        self.wrap_text
    }

    /// Set wrap text flag.
    #[setter]
    pub fn set_wrap_text(&mut self, value: Option<bool>) {
        self.wrap_text = value;
    }

    /// Get text rotation in degrees (0-180, 255=vertical), or None.
    #[getter]
    pub fn text_rotation(&self) -> Option<u32> {
        self.text_rotation
    }

    /// Set text rotation in degrees (0-180, 255=vertical).
    #[setter]
    pub fn set_text_rotation(&mut self, value: Option<u32>) {
        self.text_rotation = value;
    }
}

// =============================================================================
// Registration
// =============================================================================

pub(crate) fn register(module: &Bound<'_, PyModule>) -> PyResult<()> {
    // Core types
    module.add_class::<Workbook>()?;
    module.add_class::<Worksheet>()?;
    module.add_class::<XlsxCell>()?;
    module.add_class::<XlsxRichTextRun>()?;
    module.add_class::<XlsxStyle>()?;
    module.add_class::<XlsxFont>()?;
    module.add_class::<XlsxFill>()?;
    module.add_class::<XlsxBorder>()?;
    module.add_class::<XlsxAlignment>()?;

    // Sub-module types
    row_col::register(module)?;
    comment::register(module)?;
    protection::register(module)?;
    page::register(module)?;
    image::register(module)?;
    data_validation::register(module)?;
    table::register(module)?;
    sparkline::register(module)?;
    conditional::register(module)?;
    chart::register(module)?;
    pivot::register(module)?;
    lint::register(module)?;
    finance::register(module)?;

    Ok(())
}
