//! Python bindings for worksheet image types from `offidized_xlsx`.
//!
//! Wraps [`WorksheetImage`] and [`WorksheetImageExt`] with a PyO3 class.
//! Worksheet helper functions (`ws_*`) are called from the parent `Worksheet`
//! `#[pymethods]` block.

use super::lock_wb;
use crate::error::{value_error, xlsx_error_to_py};
use offidized_xlsx::{
    Workbook as CoreWorkbook, WorksheetImage as CoreWorksheetImage,
    WorksheetImageExt as CoreWorksheetImageExt,
};
use pyo3::prelude::*;
use pyo3::types::PyBytes;
use std::sync::{Arc, Mutex};

// =============================================================================
// XlsxWorksheetImage
// =============================================================================

/// A worksheet image anchored to a cell (value type).
#[pyclass(
    module = "offidized._native",
    name = "XlsxWorksheetImage",
    from_py_object
)]
#[derive(Clone)]
pub struct XlsxWorksheetImage {
    inner: CoreWorksheetImage,
}

impl XlsxWorksheetImage {
    pub(super) fn from_core(core: CoreWorksheetImage) -> Self {
        Self { inner: core }
    }

    pub(super) fn into_core(self) -> CoreWorksheetImage {
        self.inner
    }
}

#[pymethods]
impl XlsxWorksheetImage {
    /// Create a new worksheet image.
    ///
    /// `bytes` must be non-empty. `content_type` is the MIME type (e.g.
    /// "image/png"). `anchor_cell` is an A1 reference (e.g. "B2"). `ext_cx`
    /// and `ext_cy` are optional width/height in EMUs (both must be > 0 if
    /// provided).
    #[new]
    #[pyo3(signature = (bytes, content_type, anchor_cell, ext_cx=None, ext_cy=None))]
    pub fn new(
        bytes: Vec<u8>,
        content_type: &str,
        anchor_cell: &str,
        ext_cx: Option<u64>,
        ext_cy: Option<u64>,
    ) -> PyResult<Self> {
        let ext = match (ext_cx, ext_cy) {
            (Some(cx), Some(cy)) => {
                Some(CoreWorksheetImageExt::new(cx, cy).map_err(xlsx_error_to_py)?)
            }
            (None, None) => None,
            _ => {
                return Err(value_error(
                    "ext_cx and ext_cy must both be provided or both be None",
                ));
            }
        };
        let core = CoreWorksheetImage::new(bytes, content_type, anchor_cell, ext)
            .map_err(xlsx_error_to_py)?;
        Ok(Self { inner: core })
    }

    /// Return the raw image bytes.
    pub fn bytes<'py>(&self, py: Python<'py>) -> Bound<'py, PyBytes> {
        PyBytes::new(py, self.inner.bytes())
    }

    /// Return the MIME content type (e.g. "image/png").
    #[getter]
    pub fn content_type(&self) -> &str {
        self.inner.content_type()
    }

    /// Return the anchor cell reference (e.g. "B2").
    #[getter]
    pub fn anchor_cell(&self) -> &str {
        self.inner.anchor_cell()
    }

    /// Return the width in EMUs from the ext, or None.
    #[getter]
    pub fn ext_cx(&self) -> Option<u64> {
        self.inner.ext().map(|e| e.cx())
    }

    /// Return the height in EMUs from the ext, or None.
    #[getter]
    pub fn ext_cy(&self) -> Option<u64> {
        self.inner.ext().map(|e| e.cy())
    }

    /// Return the left crop fraction (0.0 to 1.0), or None.
    #[getter]
    pub fn crop_left(&self) -> Option<f64> {
        self.inner.crop_left()
    }

    /// Set the left crop fraction (0.0 to 1.0).
    #[setter]
    pub fn set_crop_left(&mut self, value: Option<f64>) {
        if let Some(v) = value {
            self.inner.set_crop_left(v);
        } else {
            self.inner.clear_crop_left();
        }
    }

    /// Return the right crop fraction (0.0 to 1.0), or None.
    #[getter]
    pub fn crop_right(&self) -> Option<f64> {
        self.inner.crop_right()
    }

    /// Set the right crop fraction (0.0 to 1.0).
    #[setter]
    pub fn set_crop_right(&mut self, value: Option<f64>) {
        if let Some(v) = value {
            self.inner.set_crop_right(v);
        } else {
            self.inner.clear_crop_right();
        }
    }

    /// Return the top crop fraction (0.0 to 1.0), or None.
    #[getter]
    pub fn crop_top(&self) -> Option<f64> {
        self.inner.crop_top()
    }

    /// Set the top crop fraction (0.0 to 1.0).
    #[setter]
    pub fn set_crop_top(&mut self, value: Option<f64>) {
        if let Some(v) = value {
            self.inner.set_crop_top(v);
        } else {
            self.inner.clear_crop_top();
        }
    }

    /// Return the bottom crop fraction (0.0 to 1.0), or None.
    #[getter]
    pub fn crop_bottom(&self) -> Option<f64> {
        self.inner.crop_bottom()
    }

    /// Set the bottom crop fraction (0.0 to 1.0).
    #[setter]
    pub fn set_crop_bottom(&mut self, value: Option<f64>) {
        if let Some(v) = value {
            self.inner.set_crop_bottom(v);
        } else {
            self.inner.clear_crop_bottom();
        }
    }

    /// Return the image name, or None.
    #[getter]
    pub fn name(&self) -> Option<&str> {
        self.inner.name()
    }

    /// Set the image name.
    #[setter]
    pub fn set_name(&mut self, value: Option<String>) {
        if let Some(v) = value {
            self.inner.set_name(v);
        } else {
            self.inner.clear_name();
        }
    }

    /// Return the image description/alt text, or None.
    #[getter]
    pub fn description(&self) -> Option<&str> {
        self.inner.description()
    }

    /// Set the image description/alt text.
    #[setter]
    pub fn set_description(&mut self, value: Option<String>) {
        if let Some(v) = value {
            self.inner.set_description(v);
        } else {
            self.inner.clear_description();
        }
    }
}

// =============================================================================
// Worksheet helper functions
// =============================================================================

pub(super) fn ws_images(
    wb: &Arc<Mutex<CoreWorkbook>>,
    sheet_name: &str,
) -> PyResult<Vec<XlsxWorksheetImage>> {
    let wb = lock_wb(wb)?;
    let ws = wb
        .sheet(sheet_name)
        .ok_or_else(|| value_error(format!("worksheet '{sheet_name}' not found")))?;
    Ok(ws
        .images()
        .iter()
        .cloned()
        .map(XlsxWorksheetImage::from_core)
        .collect())
}

pub(super) fn ws_add_image(
    wb: &Arc<Mutex<CoreWorkbook>>,
    sheet_name: &str,
    image: XlsxWorksheetImage,
) -> PyResult<()> {
    let mut wb = lock_wb(wb)?;
    let ws = wb
        .sheet_mut(sheet_name)
        .ok_or_else(|| value_error(format!("worksheet '{sheet_name}' not found")))?;
    let core = image.into_core();
    ws.add_image(
        core.bytes().to_vec(),
        core.content_type(),
        core.anchor_cell(),
        core.ext(),
    )
    .map_err(xlsx_error_to_py)?;
    // Re-apply crop and metadata on the last inserted image
    if let Some(img) = ws.images_mut().last_mut() {
        if let Some(v) = core.crop_left() {
            img.set_crop_left(v);
        }
        if let Some(v) = core.crop_right() {
            img.set_crop_right(v);
        }
        if let Some(v) = core.crop_top() {
            img.set_crop_top(v);
        }
        if let Some(v) = core.crop_bottom() {
            img.set_crop_bottom(v);
        }
        if let Some(n) = core.name() {
            img.set_name(n);
        }
        if let Some(d) = core.description() {
            img.set_description(d);
        }
    }
    Ok(())
}

pub(super) fn ws_clear_images(wb: &Arc<Mutex<CoreWorkbook>>, sheet_name: &str) -> PyResult<()> {
    let mut wb = lock_wb(wb)?;
    let ws = wb
        .sheet_mut(sheet_name)
        .ok_or_else(|| value_error(format!("worksheet '{sheet_name}' not found")))?;
    ws.clear_images();
    Ok(())
}

// =============================================================================
// Registration
// =============================================================================

pub(super) fn register(module: &Bound<'_, PyModule>) -> PyResult<()> {
    module.add_class::<XlsxWorksheetImage>()?;
    Ok(())
}
