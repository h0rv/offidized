use crate::error::{pptx_error_to_py, value_error};
use offidized_pptx::{
    BulletStyle as CoreBulletStyle, ChartSeries, CustomShow as CoreCustomShow, PatternFill,
    PatternFillType, Presentation as CorePresentation, ShapeGlow,
    ShapeParagraph as CoreShapeParagraph, ShapeReflection, ShapeShadow, SlideBackground,
    SlideShowSettings as CoreSlideShowSettings, SlideTransition as CoreSlideTransition,
    SlideTransitionKind as CoreSlideTransitionKind, StrikethroughStyle as CoreStrikethroughStyle,
    TextAlignment as CoreTextAlignment, TextRun as CoreTextRun,
    UnderlineStyle as CoreUnderlineStyle,
};
use pyo3::prelude::*;
use pyo3::types::PyList;
use std::sync::{Arc, Mutex};

// =============================================================================
// PPTX Bindings - Core Types
// =============================================================================

/// Python wrapper for `offidized_pptx::Presentation`.
#[pyclass(module = "offidized._native", name = "Presentation")]
pub struct Presentation {
    inner: Arc<Mutex<CorePresentation>>,
}

impl Default for Presentation {
    fn default() -> Self {
        Self::new()
    }
}

#[pymethods]
impl Presentation {
    #[new]
    pub fn new() -> Self {
        Self {
            inner: Arc::new(Mutex::new(CorePresentation::new())),
        }
    }

    /// Open an existing presentation from a file path.
    #[staticmethod]
    pub fn open(path: &str) -> PyResult<Self> {
        let presentation = CorePresentation::open(path).map_err(pptx_error_to_py)?;
        Ok(Self {
            inner: Arc::new(Mutex::new(presentation)),
        })
    }

    /// Load from bytes.
    #[staticmethod]
    pub fn from_bytes(bytes: &[u8]) -> PyResult<Self> {
        let presentation = CorePresentation::from_bytes(bytes).map_err(pptx_error_to_py)?;
        Ok(Self {
            inner: Arc::new(Mutex::new(presentation)),
        })
    }

    /// Get the number of slides in the presentation.
    pub fn slide_count(&self) -> PyResult<usize> {
        let pres = self
            .inner
            .lock()
            .map_err(|e| value_error(format!("Failed to lock presentation: {}", e)))?;
        Ok(pres.slides().len())
    }

    /// Get slide width in EMUs (English Metric Units).
    pub fn slide_width(&self) -> PyResult<Option<i64>> {
        let pres = self
            .inner
            .lock()
            .map_err(|e| value_error(format!("Failed to lock presentation: {}", e)))?;
        Ok(pres.slide_width_emu())
    }

    /// Get slide height in EMUs.
    pub fn slide_height(&self) -> PyResult<Option<i64>> {
        let pres = self
            .inner
            .lock()
            .map_err(|e| value_error(format!("Failed to lock presentation: {}", e)))?;
        Ok(pres.slide_height_emu())
    }

    /// Set slide dimensions in EMUs.
    pub fn set_slide_size(&mut self, width: i64, height: i64) -> PyResult<()> {
        let mut pres = self
            .inner
            .lock()
            .map_err(|e| value_error(format!("Failed to lock presentation: {}", e)))?;
        pres.set_slide_width_emu(width);
        pres.set_slide_height_emu(height);
        Ok(())
    }

    /// Add a new blank slide and return its index.
    pub fn add_slide(&mut self) -> PyResult<usize> {
        let mut pres = self
            .inner
            .lock()
            .map_err(|e| value_error(format!("Failed to lock presentation: {}", e)))?;
        let index = pres.slides().len();
        pres.add_slide();
        Ok(index)
    }

    /// Add a new slide with a title and return its index.
    pub fn add_slide_with_title(&mut self, title: &str) -> PyResult<usize> {
        let mut pres = self
            .inner
            .lock()
            .map_err(|e| value_error(format!("Failed to lock presentation: {}", e)))?;
        let index = pres.slides().len();
        pres.add_slide_with_title(title);
        Ok(index)
    }

    /// Get a slide by index (0-based).
    pub fn get_slide(&self, index: usize) -> PyResult<Slide> {
        let pres = self
            .inner
            .lock()
            .map_err(|e| value_error(format!("Failed to lock presentation: {}", e)))?;

        if index >= pres.slides().len() {
            return Err(value_error(format!(
                "Slide index {} out of range (0..{})",
                index,
                pres.slides().len()
            )));
        }

        Ok(Slide {
            presentation: Arc::clone(&self.inner),
            index,
        })
    }

    /// Get all slide indices as a list.
    pub fn slides(&self, py: Python<'_>) -> PyResult<Py<PyList>> {
        let pres = self
            .inner
            .lock()
            .map_err(|e| value_error(format!("Failed to lock presentation: {}", e)))?;

        let indices: Vec<usize> = (0..pres.slides().len()).collect();

        Ok(PyList::new(py, indices)?.into())
    }

    /// Get presentation properties for metadata.
    pub fn properties(&self) -> PyResult<PresentationProperties> {
        Ok(PresentationProperties {
            presentation: Arc::clone(&self.inner),
        })
    }

    /// Get slide show settings, or None if not set.
    pub fn slide_show_settings(&self) -> PyResult<Option<PySlideShowSettings>> {
        let pres = self
            .inner
            .lock()
            .map_err(|e| value_error(format!("Failed to lock presentation: {}", e)))?;
        Ok(pres
            .slide_show_settings()
            .map(|s| PySlideShowSettings { inner: s.clone() }))
    }

    /// Set slide show settings.
    pub fn set_slide_show_settings(&mut self, settings: &PySlideShowSettings) -> PyResult<()> {
        let mut pres = self
            .inner
            .lock()
            .map_err(|e| value_error(format!("Failed to lock presentation: {}", e)))?;
        pres.set_slide_show_settings(settings.inner.clone());
        Ok(())
    }

    /// Get all custom shows.
    pub fn custom_shows(&self, _py: Python<'_>) -> PyResult<Vec<PyCustomShow>> {
        let pres = self
            .inner
            .lock()
            .map_err(|e| value_error(format!("Failed to lock presentation: {}", e)))?;
        Ok(pres
            .custom_shows()
            .iter()
            .map(|cs| PyCustomShow { inner: cs.clone() })
            .collect())
    }

    /// Add a custom show to the presentation.
    pub fn add_custom_show(&mut self, show: &PyCustomShow) -> PyResult<()> {
        let mut pres = self
            .inner
            .lock()
            .map_err(|e| value_error(format!("Failed to lock presentation: {}", e)))?;
        pres.custom_shows_mut().push(show.inner.clone());
        Ok(())
    }

    /// Save presentation to a `.pptx` path.
    pub fn save(&mut self, path: &str) -> PyResult<()> {
        let mut pres = self
            .inner
            .lock()
            .map_err(|e| value_error(format!("Failed to lock presentation: {}", e)))?;
        pres.save(path).map_err(pptx_error_to_py)
    }

    /// Remove a slide by index.
    pub fn remove_slide(&mut self, index: usize) -> PyResult<bool> {
        let mut pres = self
            .inner
            .lock()
            .map_err(|e| value_error(format!("Failed to lock presentation: {}", e)))?;
        Ok(pres.remove_slide(index).is_some())
    }

    /// Clone (duplicate) a slide by index. Returns the new slide index.
    pub fn clone_slide(&mut self, index: usize) -> PyResult<Option<usize>> {
        let mut pres = self
            .inner
            .lock()
            .map_err(|e| value_error(format!("Failed to lock presentation: {}", e)))?;
        Ok(pres.clone_slide(index))
    }

    /// Move a slide from one position to another.
    pub fn move_slide(&mut self, from_index: usize, to_index: usize) -> PyResult<bool> {
        let mut pres = self
            .inner
            .lock()
            .map_err(|e| value_error(format!("Failed to lock presentation: {}", e)))?;
        Ok(pres.move_slide(from_index, to_index))
    }

    /// Find all occurrences of text. Returns list of (slide_idx, shape_idx, para_idx, run_idx).
    pub fn find_text(&self, needle: &str) -> PyResult<Vec<(usize, usize, usize, usize)>> {
        let pres = self
            .inner
            .lock()
            .map_err(|e| value_error(format!("Failed to lock presentation: {}", e)))?;
        Ok(pres.find_text(needle))
    }

    /// Replace all occurrences of old_text with new_text. Returns count of replacements.
    pub fn replace_text(&mut self, old_text: &str, new_text: &str) -> PyResult<usize> {
        let mut pres = self
            .inner
            .lock()
            .map_err(|e| value_error(format!("Failed to lock presentation: {}", e)))?;
        Ok(pres.replace_text(old_text, new_text))
    }

    /// Serialize presentation to bytes.
    #[allow(clippy::wrong_self_convention)]
    pub fn to_bytes(&mut self) -> PyResult<Vec<u8>> {
        let mut pres = self
            .inner
            .lock()
            .map_err(|e| value_error(format!("Failed to lock presentation: {}", e)))?;
        let dir = std::env::temp_dir();
        let path = dir.join(format!("offidized-pptx-{}.pptx", std::process::id()));
        pres.save(&path).map_err(pptx_error_to_py)?;
        drop(pres);
        let bytes = std::fs::read(&path)
            .map_err(|e| value_error(format!("Failed to read temp file: {}", e)))?;
        let _ = std::fs::remove_file(&path);
        Ok(bytes)
    }

    /// Get theme color by name (e.g. "dk1", "lt1", "accent1"), or None.
    pub fn theme_color(&self, name: &str) -> PyResult<Option<String>> {
        let pres = self
            .inner
            .lock()
            .map_err(|e| value_error(format!("Failed to lock presentation: {}", e)))?;
        Ok(pres
            .theme_color_scheme()
            .and_then(|s| s.color_by_name(name).map(String::from)))
    }

    /// Set a theme color by name.
    pub fn set_theme_color(&mut self, name: &str, value: &str) -> PyResult<()> {
        let mut pres = self
            .inner
            .lock()
            .map_err(|e| value_error(format!("Failed to lock presentation: {}", e)))?;
        let mut scheme = pres.theme_color_scheme().cloned().unwrap_or_default();
        if !scheme.set_color_by_name(name, value) {
            return Err(value_error(format!("Unknown theme color name: {}", name)));
        }
        pres.set_theme_color_scheme(scheme);
        Ok(())
    }

    /// Get theme font names as (major_latin, minor_latin), or None.
    pub fn theme_fonts(&self) -> PyResult<Option<(String, String)>> {
        let pres = self
            .inner
            .lock()
            .map_err(|e| value_error(format!("Failed to lock presentation: {}", e)))?;
        Ok(pres
            .theme_font_scheme()
            .map(|f| (f.major_latin.clone(), f.minor_latin.clone())))
    }

    /// Get number of slide masters.
    pub fn slide_master_count(&self) -> PyResult<usize> {
        let pres = self
            .inner
            .lock()
            .map_err(|e| value_error(format!("Failed to lock presentation: {}", e)))?;
        Ok(pres.slide_masters_v2().len())
    }

    /// Get total number of layouts across all masters.
    pub fn layout_count(&self) -> PyResult<usize> {
        let pres = self
            .inner
            .lock()
            .map_err(|e| value_error(format!("Failed to lock presentation: {}", e)))?;
        Ok(pres.layouts().len())
    }
}

/// Python wrapper for presentation properties (metadata).
#[pyclass(module = "offidized._native", name = "PresentationProperties")]
#[derive(Clone)]
pub struct PresentationProperties {
    presentation: Arc<Mutex<CorePresentation>>,
}

#[pymethods]
impl PresentationProperties {
    /// Get the author.
    pub fn author(&self) -> PyResult<Option<String>> {
        let pres = self
            .presentation
            .lock()
            .map_err(|e| value_error(format!("Failed to lock presentation: {}", e)))?;
        Ok(pres.properties().author().map(String::from))
    }

    /// Set the author.
    #[pyo3(signature = (author=None))]
    pub fn set_author(&mut self, author: Option<String>) -> PyResult<()> {
        let mut pres = self
            .presentation
            .lock()
            .map_err(|e| value_error(format!("Failed to lock presentation: {}", e)))?;
        if let Some(author) = author {
            pres.properties_mut().set_author(author);
        } else {
            pres.properties_mut().clear_author();
        }
        Ok(())
    }

    /// Get the title.
    pub fn title(&self) -> PyResult<Option<String>> {
        let pres = self
            .presentation
            .lock()
            .map_err(|e| value_error(format!("Failed to lock presentation: {}", e)))?;
        Ok(pres.properties().title().map(String::from))
    }

    /// Set the title.
    #[pyo3(signature = (title=None))]
    pub fn set_title(&mut self, title: Option<String>) -> PyResult<()> {
        let mut pres = self
            .presentation
            .lock()
            .map_err(|e| value_error(format!("Failed to lock presentation: {}", e)))?;
        if let Some(title) = title {
            pres.properties_mut().set_title(title);
        } else {
            pres.properties_mut().clear_title();
        }
        Ok(())
    }

    /// Get the subject.
    pub fn subject(&self) -> PyResult<Option<String>> {
        let pres = self
            .presentation
            .lock()
            .map_err(|e| value_error(format!("Failed to lock presentation: {}", e)))?;
        Ok(pres.properties().subject().map(String::from))
    }

    /// Set the subject.
    #[pyo3(signature = (subject=None))]
    pub fn set_subject(&mut self, subject: Option<String>) -> PyResult<()> {
        let mut pres = self
            .presentation
            .lock()
            .map_err(|e| value_error(format!("Failed to lock presentation: {}", e)))?;
        if let Some(subject) = subject {
            pres.properties_mut().set_subject(subject);
        } else {
            pres.properties_mut().clear_subject();
        }
        Ok(())
    }

    /// Get keywords.
    pub fn keywords(&self) -> PyResult<Option<String>> {
        let pres = self
            .presentation
            .lock()
            .map_err(|e| value_error(format!("Failed to lock presentation: {}", e)))?;
        Ok(pres.properties().keywords().map(String::from))
    }

    /// Set keywords.
    #[pyo3(signature = (keywords=None))]
    pub fn set_keywords(&mut self, keywords: Option<String>) -> PyResult<()> {
        let mut pres = self
            .presentation
            .lock()
            .map_err(|e| value_error(format!("Failed to lock presentation: {}", e)))?;
        if let Some(keywords) = keywords {
            pres.properties_mut().set_keywords(keywords);
        } else {
            pres.properties_mut().clear_keywords();
        }
        Ok(())
    }

    /// Get category.
    pub fn category(&self) -> PyResult<Option<String>> {
        let pres = self
            .presentation
            .lock()
            .map_err(|e| value_error(format!("Failed to lock presentation: {}", e)))?;
        Ok(pres.properties().category().map(String::from))
    }

    /// Set category.
    #[pyo3(signature = (category=None))]
    pub fn set_category(&mut self, category: Option<String>) -> PyResult<()> {
        let mut pres = self
            .presentation
            .lock()
            .map_err(|e| value_error(format!("Failed to lock presentation: {}", e)))?;
        if let Some(category) = category {
            pres.properties_mut().set_category(category);
        } else {
            pres.properties_mut().clear_category();
        }
        Ok(())
    }

    /// Get company.
    pub fn company(&self) -> PyResult<Option<String>> {
        let pres = self
            .presentation
            .lock()
            .map_err(|e| value_error(format!("Failed to lock presentation: {}", e)))?;
        Ok(pres.properties().company().map(String::from))
    }

    /// Set company.
    #[pyo3(signature = (company=None))]
    pub fn set_company(&mut self, company: Option<String>) -> PyResult<()> {
        let mut pres = self
            .presentation
            .lock()
            .map_err(|e| value_error(format!("Failed to lock presentation: {}", e)))?;
        if let Some(company) = company {
            pres.properties_mut().set_company(company);
        } else {
            pres.properties_mut().clear_company();
        }
        Ok(())
    }
}

// =============================================================================
// PPTX Bindings - Slide
// =============================================================================

/// Python wrapper for `offidized_pptx::Slide`.
#[pyclass(module = "offidized._native", name = "Slide")]
#[derive(Clone)]
pub struct Slide {
    presentation: Arc<Mutex<CorePresentation>>,
    index: usize,
}

#[pymethods]
impl Slide {
    /// Get the number of shapes on this slide.
    pub fn shape_count(&self) -> PyResult<usize> {
        let pres = self
            .presentation
            .lock()
            .map_err(|e| value_error(format!("Failed to lock presentation: {}", e)))?;
        let slide = pres
            .slides()
            .get(self.index)
            .ok_or_else(|| value_error("Slide no longer exists"))?;
        Ok(slide.shapes().len())
    }

    /// Add a shape to the slide.
    pub fn add_shape(&mut self, name: &str) -> PyResult<usize> {
        let mut pres = self
            .presentation
            .lock()
            .map_err(|e| value_error(format!("Failed to lock presentation: {}", e)))?;
        let slide = pres
            .slides_mut()
            .get_mut(self.index)
            .ok_or_else(|| value_error("Slide no longer exists"))?;

        let index = slide.shapes().len();
        slide.add_shape(name);
        Ok(index)
    }

    /// Add a table to the slide.
    pub fn add_table(&mut self, rows: usize, cols: usize) -> PyResult<usize> {
        let mut pres = self
            .presentation
            .lock()
            .map_err(|e| value_error(format!("Failed to lock presentation: {}", e)))?;
        let slide = pres
            .slides_mut()
            .get_mut(self.index)
            .ok_or_else(|| value_error("Slide no longer exists"))?;

        let index = slide.tables().len();
        slide.add_table(rows, cols);
        Ok(index)
    }

    /// Add an image to the slide.
    pub fn add_image(&mut self, path: &str) -> PyResult<usize> {
        let mut pres = self
            .presentation
            .lock()
            .map_err(|e| value_error(format!("Failed to lock presentation: {}", e)))?;
        let slide = pres
            .slides_mut()
            .get_mut(self.index)
            .ok_or_else(|| value_error("Slide no longer exists"))?;

        let image_bytes = std::fs::read(path)
            .map_err(|e| value_error(format!("Failed to read image file: {}", e)))?;

        let content_type = if path.to_lowercase().ends_with(".png") {
            "image/png"
        } else if path.to_lowercase().ends_with(".jpg") || path.to_lowercase().ends_with(".jpeg") {
            "image/jpeg"
        } else if path.to_lowercase().ends_with(".gif") {
            "image/gif"
        } else {
            "image/png"
        };

        let index = slide.images().len();
        slide.add_image(image_bytes, content_type);
        Ok(index)
    }

    /// Add a chart to the slide.
    pub fn add_chart(&mut self, chart_type: &str) -> PyResult<usize> {
        let chart_title = match chart_type.to_lowercase().as_str() {
            "bar" => "Bar Chart",
            "column" => "Column Chart",
            "line" => "Line Chart",
            "pie" => "Pie Chart",
            "scatter" => "Scatter Chart",
            "area" => "Area Chart",
            "doughnut" => "Doughnut Chart",
            "radar" => "Radar Chart",
            _ => "Chart",
        };

        let mut pres = self
            .presentation
            .lock()
            .map_err(|e| value_error(format!("Failed to lock presentation: {}", e)))?;
        let slide = pres
            .slides_mut()
            .get_mut(self.index)
            .ok_or_else(|| value_error("Slide no longer exists"))?;

        let index = slide.charts().len();
        slide.add_chart(chart_title);
        Ok(index)
    }

    /// Get slide notes text.
    pub fn notes(&self) -> PyResult<Option<String>> {
        let pres = self
            .presentation
            .lock()
            .map_err(|e| value_error(format!("Failed to lock presentation: {}", e)))?;
        let slide = pres
            .slides()
            .get(self.index)
            .ok_or_else(|| value_error("Slide no longer exists"))?;
        Ok(slide.notes_text().map(String::from))
    }

    /// Set slide notes text.
    #[pyo3(signature = (notes=None))]
    pub fn set_notes(&mut self, notes: Option<String>) -> PyResult<()> {
        let mut pres = self
            .presentation
            .lock()
            .map_err(|e| value_error(format!("Failed to lock presentation: {}", e)))?;
        let slide = pres
            .slides_mut()
            .get_mut(self.index)
            .ok_or_else(|| value_error("Slide no longer exists"))?;
        if let Some(notes) = notes {
            slide.set_notes_text(notes);
        } else {
            slide.set_notes_text("");
        }
        Ok(())
    }

    /// Get slide transition, or None if no transition is set.
    pub fn transition(&self) -> PyResult<Option<PySlideTransition>> {
        let pres = self
            .presentation
            .lock()
            .map_err(|e| value_error(format!("Failed to lock presentation: {}", e)))?;
        let slide = pres
            .slides()
            .get(self.index)
            .ok_or_else(|| value_error("Slide no longer exists"))?;
        Ok(slide
            .transition()
            .map(|t| PySlideTransition { inner: t.clone() }))
    }

    /// Set slide transition.
    pub fn set_transition(&mut self, transition: &PySlideTransition) -> PyResult<()> {
        let mut pres = self
            .presentation
            .lock()
            .map_err(|e| value_error(format!("Failed to lock presentation: {}", e)))?;
        let slide = pres
            .slides_mut()
            .get_mut(self.index)
            .ok_or_else(|| value_error("Slide no longer exists"))?;
        slide.set_transition(transition.inner.clone());
        Ok(())
    }

    /// Get a shape by index (0-based), providing access to text paragraphs and runs.
    pub fn get_shape(&self, index: usize) -> PyResult<PyShape> {
        let pres = self
            .presentation
            .lock()
            .map_err(|e| value_error(format!("Failed to lock presentation: {}", e)))?;
        let slide = pres
            .slides()
            .get(self.index)
            .ok_or_else(|| value_error("Slide no longer exists"))?;
        if index >= slide.shapes().len() {
            return Err(value_error(format!(
                "Shape index {} out of range (0..{})",
                index,
                slide.shapes().len()
            )));
        }
        Ok(PyShape {
            presentation: Arc::clone(&self.presentation),
            slide_index: self.index,
            shape_index: index,
        })
    }

    /// Get a table by index.
    pub fn get_table(&self, index: usize) -> PyResult<Table> {
        let pres = self
            .presentation
            .lock()
            .map_err(|e| value_error(format!("Failed to lock presentation: {}", e)))?;
        let slide = pres
            .slides()
            .get(self.index)
            .ok_or_else(|| value_error("Slide no longer exists"))?;
        if index >= slide.tables().len() {
            return Err(value_error(format!("Table index {} out of range", index)));
        }
        Ok(Table {
            presentation: Arc::clone(&self.presentation),
            slide_index: self.index,
            table_index: index,
        })
    }

    /// Get number of tables on this slide.
    pub fn table_count(&self) -> PyResult<usize> {
        let pres = self
            .presentation
            .lock()
            .map_err(|e| value_error(format!("Failed to lock presentation: {}", e)))?;
        let slide = pres
            .slides()
            .get(self.index)
            .ok_or_else(|| value_error("Slide no longer exists"))?;
        Ok(slide.tables().len())
    }

    /// Get a chart by index.
    pub fn get_chart(&self, index: usize) -> PyResult<Chart> {
        let pres = self
            .presentation
            .lock()
            .map_err(|e| value_error(format!("Failed to lock presentation: {}", e)))?;
        let slide = pres
            .slides()
            .get(self.index)
            .ok_or_else(|| value_error("Slide no longer exists"))?;
        if index >= slide.charts().len() {
            return Err(value_error(format!("Chart index {} out of range", index)));
        }
        Ok(Chart {
            presentation: Arc::clone(&self.presentation),
            slide_index: self.index,
            chart_index: index,
        })
    }

    /// Get number of charts on this slide.
    pub fn chart_count(&self) -> PyResult<usize> {
        let pres = self
            .presentation
            .lock()
            .map_err(|e| value_error(format!("Failed to lock presentation: {}", e)))?;
        let slide = pres
            .slides()
            .get(self.index)
            .ok_or_else(|| value_error("Slide no longer exists"))?;
        Ok(slide.charts().len())
    }

    /// Get an image by index.
    pub fn get_image(&self, index: usize) -> PyResult<Image> {
        let pres = self
            .presentation
            .lock()
            .map_err(|e| value_error(format!("Failed to lock presentation: {}", e)))?;
        let slide = pres
            .slides()
            .get(self.index)
            .ok_or_else(|| value_error("Slide no longer exists"))?;
        if index >= slide.images().len() {
            return Err(value_error(format!("Image index {} out of range", index)));
        }
        Ok(Image {
            presentation: Arc::clone(&self.presentation),
            slide_index: self.index,
            image_index: index,
        })
    }

    /// Get number of images on this slide.
    pub fn image_count(&self) -> PyResult<usize> {
        let pres = self
            .presentation
            .lock()
            .map_err(|e| value_error(format!("Failed to lock presentation: {}", e)))?;
        let slide = pres
            .slides()
            .get(self.index)
            .ok_or_else(|| value_error("Slide no longer exists"))?;
        Ok(slide.images().len())
    }

    /// Get slide background type: "solid", "gradient", "pattern", "image", or None.
    pub fn background_type(&self) -> PyResult<Option<String>> {
        let pres = self
            .presentation
            .lock()
            .map_err(|e| value_error(format!("Failed to lock presentation: {}", e)))?;
        let slide = pres
            .slides()
            .get(self.index)
            .ok_or_else(|| value_error("Slide no longer exists"))?;
        Ok(slide.background().map(|bg| {
            match bg {
                SlideBackground::Solid(_) => "solid",
                SlideBackground::Gradient(_) => "gradient",
                SlideBackground::Pattern { .. } => "pattern",
                SlideBackground::Image { .. } => "image",
            }
            .to_owned()
        }))
    }

    /// Get solid background color, or None.
    pub fn background_solid_color(&self) -> PyResult<Option<String>> {
        let pres = self
            .presentation
            .lock()
            .map_err(|e| value_error(format!("Failed to lock presentation: {}", e)))?;
        let slide = pres
            .slides()
            .get(self.index)
            .ok_or_else(|| value_error("Slide no longer exists"))?;
        Ok(match slide.background() {
            Some(SlideBackground::Solid(c)) => Some(c.clone()),
            _ => None,
        })
    }

    /// Set solid background color.
    pub fn set_background_solid(&mut self, color: &str) -> PyResult<()> {
        let mut pres = self
            .presentation
            .lock()
            .map_err(|e| value_error(format!("Failed to lock presentation: {}", e)))?;
        let slide = pres
            .slides_mut()
            .get_mut(self.index)
            .ok_or_else(|| value_error("Slide no longer exists"))?;
        slide.set_background(SlideBackground::Solid(color.to_string()));
        Ok(())
    }

    /// Clear background.
    pub fn clear_background(&mut self) -> PyResult<()> {
        let mut pres = self
            .presentation
            .lock()
            .map_err(|e| value_error(format!("Failed to lock presentation: {}", e)))?;
        let slide = pres
            .slides_mut()
            .get_mut(self.index)
            .ok_or_else(|| value_error("Slide no longer exists"))?;
        slide.clear_background();
        Ok(())
    }
}

// =============================================================================
// PPTX Bindings - Table Wrapper
// =============================================================================

/// Python wrapper for `offidized_pptx::Table`.
#[pyclass(module = "offidized._native", name = "Table")]
#[derive(Clone)]
pub struct Table {
    presentation: Arc<Mutex<CorePresentation>>,
    slide_index: usize,
    table_index: usize,
}

#[pymethods]
impl Table {
    /// Set table position and size in EMUs.
    pub fn set_geometry(&mut self, x: i64, y: i64, width: i64, height: i64) -> PyResult<()> {
        let mut pres = self
            .presentation
            .lock()
            .map_err(|e| value_error(format!("Failed to lock presentation: {}", e)))?;
        let slide = pres
            .slides_mut()
            .get_mut(self.slide_index)
            .ok_or_else(|| value_error("Slide no longer exists"))?;
        let table = slide
            .tables_mut()
            .get_mut(self.table_index)
            .ok_or_else(|| value_error("Table no longer exists"))?;
        table.set_geometry(x, y, width, height);
        Ok(())
    }

    /// Get number of rows.
    pub fn rows(&self) -> PyResult<usize> {
        let pres = self
            .presentation
            .lock()
            .map_err(|e| value_error(format!("Failed to lock presentation: {}", e)))?;
        let slide = pres
            .slides()
            .get(self.slide_index)
            .ok_or_else(|| value_error("Slide no longer exists"))?;
        let table = slide
            .tables()
            .get(self.table_index)
            .ok_or_else(|| value_error("Table no longer exists"))?;
        Ok(table.rows())
    }

    /// Get number of columns.
    pub fn columns(&self) -> PyResult<usize> {
        let pres = self
            .presentation
            .lock()
            .map_err(|e| value_error(format!("Failed to lock presentation: {}", e)))?;
        let slide = pres
            .slides()
            .get(self.slide_index)
            .ok_or_else(|| value_error("Slide no longer exists"))?;
        let table = slide
            .tables()
            .get(self.table_index)
            .ok_or_else(|| value_error("Table no longer exists"))?;
        Ok(table.cols())
    }

    /// Get cell text.
    pub fn cell_text(&self, row: usize, col: usize) -> PyResult<String> {
        let pres = self
            .presentation
            .lock()
            .map_err(|e| value_error(format!("Failed to lock presentation: {}", e)))?;
        let slide = pres
            .slides()
            .get(self.slide_index)
            .ok_or_else(|| value_error("Slide no longer exists"))?;
        let table = slide
            .tables()
            .get(self.table_index)
            .ok_or_else(|| value_error("Table no longer exists"))?;
        match table.cell_text(row, col) {
            Some(text) => Ok(text.to_string()),
            None => Err(value_error(format!(
                "Cell ({}, {}) out of bounds",
                row, col
            ))),
        }
    }

    /// Set cell text.
    pub fn set_cell_text(&mut self, row: usize, col: usize, text: &str) -> PyResult<()> {
        let mut pres = self
            .presentation
            .lock()
            .map_err(|e| value_error(format!("Failed to lock presentation: {}", e)))?;
        let slide = pres
            .slides_mut()
            .get_mut(self.slide_index)
            .ok_or_else(|| value_error("Slide no longer exists"))?;
        let table = slide
            .tables_mut()
            .get_mut(self.table_index)
            .ok_or_else(|| value_error("Table no longer exists"))?;
        if table.set_cell_text(row, col, text) {
            Ok(())
        } else {
            Err(value_error(format!(
                "Cell ({}, {}) out of bounds",
                row, col
            )))
        }
    }

    /// Set column width in EMUs.
    pub fn set_column_width(&mut self, col: usize, width: i64) -> PyResult<()> {
        let mut pres = self
            .presentation
            .lock()
            .map_err(|e| value_error(format!("Failed to lock presentation: {}", e)))?;
        let slide = pres
            .slides_mut()
            .get_mut(self.slide_index)
            .ok_or_else(|| value_error("Slide no longer exists"))?;
        let table = slide
            .tables_mut()
            .get_mut(self.table_index)
            .ok_or_else(|| value_error("Table no longer exists"))?;
        if table.set_column_width_emu(col, width) {
            Ok(())
        } else {
            Err(value_error(format!("Column {} out of bounds", col)))
        }
    }

    /// Set row height in EMUs.
    pub fn set_row_height(&mut self, row: usize, height: i64) -> PyResult<()> {
        let mut pres = self
            .presentation
            .lock()
            .map_err(|e| value_error(format!("Failed to lock presentation: {}", e)))?;
        let slide = pres
            .slides_mut()
            .get_mut(self.slide_index)
            .ok_or_else(|| value_error("Slide no longer exists"))?;
        let table = slide
            .tables_mut()
            .get_mut(self.table_index)
            .ok_or_else(|| value_error("Table no longer exists"))?;
        if table.set_row_height_emu(row, height) {
            Ok(())
        } else {
            Err(value_error(format!("Row {} out of bounds", row)))
        }
    }

    /// Set cell background color.
    pub fn set_cell_fill(&mut self, row: usize, col: usize, color: &str) -> PyResult<()> {
        let mut pres = self
            .presentation
            .lock()
            .map_err(|e| value_error(format!("Failed to lock presentation: {}", e)))?;
        let slide = pres
            .slides_mut()
            .get_mut(self.slide_index)
            .ok_or_else(|| value_error("Slide no longer exists"))?;
        let table = slide
            .tables_mut()
            .get_mut(self.table_index)
            .ok_or_else(|| value_error("Table no longer exists"))?;
        let cell = table
            .cell_mut(row, col)
            .ok_or_else(|| value_error(format!("Cell ({}, {}) out of bounds", row, col)))?;
        cell.set_fill_color_srgb(color);
        Ok(())
    }

    /// Set cell text bold.
    pub fn set_cell_bold(&mut self, row: usize, col: usize, bold: bool) -> PyResult<()> {
        let mut pres = self
            .presentation
            .lock()
            .map_err(|e| value_error(format!("Failed to lock presentation: {}", e)))?;
        let slide = pres
            .slides_mut()
            .get_mut(self.slide_index)
            .ok_or_else(|| value_error("Slide no longer exists"))?;
        let table = slide
            .tables_mut()
            .get_mut(self.table_index)
            .ok_or_else(|| value_error("Table no longer exists"))?;
        let cell = table
            .cell_mut(row, col)
            .ok_or_else(|| value_error(format!("Cell ({}, {}) out of bounds", row, col)))?;
        cell.set_bold(bold);
        Ok(())
    }

    /// Set cell text italic.
    pub fn set_cell_italic(&mut self, row: usize, col: usize, italic: bool) -> PyResult<()> {
        let mut pres = self
            .presentation
            .lock()
            .map_err(|e| value_error(format!("Failed to lock presentation: {}", e)))?;
        let slide = pres
            .slides_mut()
            .get_mut(self.slide_index)
            .ok_or_else(|| value_error("Slide no longer exists"))?;
        let table = slide
            .tables_mut()
            .get_mut(self.table_index)
            .ok_or_else(|| value_error("Table no longer exists"))?;
        let cell = table
            .cell_mut(row, col)
            .ok_or_else(|| value_error(format!("Cell ({}, {}) out of bounds", row, col)))?;
        cell.set_italic(italic);
        Ok(())
    }

    /// Set cell text font size in hundredths of a point (e.g. 1200 = 12pt).
    pub fn set_cell_font_size(&mut self, row: usize, col: usize, size: u32) -> PyResult<()> {
        let mut pres = self
            .presentation
            .lock()
            .map_err(|e| value_error(format!("Failed to lock presentation: {}", e)))?;
        let slide = pres
            .slides_mut()
            .get_mut(self.slide_index)
            .ok_or_else(|| value_error("Slide no longer exists"))?;
        let table = slide
            .tables_mut()
            .get_mut(self.table_index)
            .ok_or_else(|| value_error("Table no longer exists"))?;
        let cell = table
            .cell_mut(row, col)
            .ok_or_else(|| value_error(format!("Cell ({}, {}) out of bounds", row, col)))?;
        cell.set_font_size(size);
        Ok(())
    }

    /// Set cell text color as sRGB hex (e.g. "FF0000").
    pub fn set_cell_font_color(&mut self, row: usize, col: usize, color: &str) -> PyResult<()> {
        let mut pres = self
            .presentation
            .lock()
            .map_err(|e| value_error(format!("Failed to lock presentation: {}", e)))?;
        let slide = pres
            .slides_mut()
            .get_mut(self.slide_index)
            .ok_or_else(|| value_error("Slide no longer exists"))?;
        let table = slide
            .tables_mut()
            .get_mut(self.table_index)
            .ok_or_else(|| value_error("Table no longer exists"))?;
        let cell = table
            .cell_mut(row, col)
            .ok_or_else(|| value_error(format!("Cell ({}, {}) out of bounds", row, col)))?;
        cell.set_font_color_srgb(color);
        Ok(())
    }

    /// Insert a row at the given index with height in EMUs.
    pub fn insert_row(&mut self, at: usize, height_emu: i64) -> PyResult<bool> {
        let mut pres = self
            .presentation
            .lock()
            .map_err(|e| value_error(format!("Failed to lock presentation: {}", e)))?;
        let slide = pres
            .slides_mut()
            .get_mut(self.slide_index)
            .ok_or_else(|| value_error("Slide no longer exists"))?;
        let table = slide
            .tables_mut()
            .get_mut(self.table_index)
            .ok_or_else(|| value_error("Table no longer exists"))?;
        Ok(table.insert_row(at, height_emu))
    }

    /// Remove a row at the given index.
    pub fn remove_row(&mut self, at: usize) -> PyResult<bool> {
        let mut pres = self
            .presentation
            .lock()
            .map_err(|e| value_error(format!("Failed to lock presentation: {}", e)))?;
        let slide = pres
            .slides_mut()
            .get_mut(self.slide_index)
            .ok_or_else(|| value_error("Slide no longer exists"))?;
        let table = slide
            .tables_mut()
            .get_mut(self.table_index)
            .ok_or_else(|| value_error("Table no longer exists"))?;
        Ok(table.remove_row(at))
    }

    /// Merge cells from (start_row, start_col) to (end_row, end_col).
    pub fn merge_cells(
        &mut self,
        start_row: usize,
        start_col: usize,
        end_row: usize,
        end_col: usize,
    ) -> PyResult<bool> {
        let mut pres = self
            .presentation
            .lock()
            .map_err(|e| value_error(format!("Failed to lock presentation: {}", e)))?;
        let slide = pres
            .slides_mut()
            .get_mut(self.slide_index)
            .ok_or_else(|| value_error("Slide no longer exists"))?;
        let table = slide
            .tables_mut()
            .get_mut(self.table_index)
            .ok_or_else(|| value_error("Table no longer exists"))?;
        Ok(table.merge_cells(start_row, start_col, end_row, end_col))
    }
}

// =============================================================================
// PPTX Bindings - Chart Wrapper
// =============================================================================

/// Python wrapper for `offidized_pptx::Chart`.
#[pyclass(module = "offidized._native", name = "Chart")]
#[derive(Clone)]
pub struct Chart {
    presentation: Arc<Mutex<CorePresentation>>,
    slide_index: usize,
    chart_index: usize,
}

#[pymethods]
impl Chart {
    /// Get chart title.
    pub fn title(&self) -> PyResult<String> {
        let pres = self
            .presentation
            .lock()
            .map_err(|e| value_error(format!("Failed to lock presentation: {}", e)))?;
        let slide = pres
            .slides()
            .get(self.slide_index)
            .ok_or_else(|| value_error("Slide no longer exists"))?;
        let chart = slide
            .charts()
            .get(self.chart_index)
            .ok_or_else(|| value_error("Chart no longer exists"))?;
        Ok(chart.title().to_string())
    }

    /// Set chart title.
    pub fn set_title(&mut self, title: &str) -> PyResult<()> {
        let mut pres = self
            .presentation
            .lock()
            .map_err(|e| value_error(format!("Failed to lock presentation: {}", e)))?;
        let slide = pres
            .slides_mut()
            .get_mut(self.slide_index)
            .ok_or_else(|| value_error("Slide no longer exists"))?;
        let chart = slide
            .charts_mut()
            .get_mut(self.chart_index)
            .ok_or_else(|| value_error("Chart no longer exists"))?;
        chart.set_title(title);
        Ok(())
    }

    /// Add a data point.
    pub fn add_data_point(&mut self, category: &str, value: f64) -> PyResult<()> {
        let mut pres = self
            .presentation
            .lock()
            .map_err(|e| value_error(format!("Failed to lock presentation: {}", e)))?;
        let slide = pres
            .slides_mut()
            .get_mut(self.slide_index)
            .ok_or_else(|| value_error("Slide no longer exists"))?;
        let chart = slide
            .charts_mut()
            .get_mut(self.chart_index)
            .ok_or_else(|| value_error("Chart no longer exists"))?;
        chart.add_data_point(category, value);
        Ok(())
    }

    /// Get number of data points.
    pub fn point_count(&self) -> PyResult<usize> {
        let pres = self
            .presentation
            .lock()
            .map_err(|e| value_error(format!("Failed to lock presentation: {}", e)))?;
        let slide = pres
            .slides()
            .get(self.slide_index)
            .ok_or_else(|| value_error("Slide no longer exists"))?;
        let chart = slide
            .charts()
            .get(self.chart_index)
            .ok_or_else(|| value_error("Chart no longer exists"))?;
        Ok(chart.point_count())
    }

    /// Show/hide legend.
    pub fn set_show_legend(&mut self, show: bool) -> PyResult<()> {
        let mut pres = self
            .presentation
            .lock()
            .map_err(|e| value_error(format!("Failed to lock presentation: {}", e)))?;
        let slide = pres
            .slides_mut()
            .get_mut(self.slide_index)
            .ok_or_else(|| value_error("Slide no longer exists"))?;
        let chart = slide
            .charts_mut()
            .get_mut(self.chart_index)
            .ok_or_else(|| value_error("Chart no longer exists"))?;
        chart.set_show_legend(show);
        Ok(())
    }

    /// Check if legend is shown.
    pub fn show_legend(&self) -> PyResult<bool> {
        let pres = self
            .presentation
            .lock()
            .map_err(|e| value_error(format!("Failed to lock presentation: {}", e)))?;
        let slide = pres
            .slides()
            .get(self.slide_index)
            .ok_or_else(|| value_error("Slide no longer exists"))?;
        let chart = slide
            .charts()
            .get(self.chart_index)
            .ok_or_else(|| value_error("Chart no longer exists"))?;
        Ok(chart.show_legend())
    }

    /// Get number of additional series.
    pub fn series_count(&self) -> PyResult<usize> {
        let pres = self
            .presentation
            .lock()
            .map_err(|e| value_error(format!("Failed to lock presentation: {}", e)))?;
        let slide = pres
            .slides()
            .get(self.slide_index)
            .ok_or_else(|| value_error("Slide no longer exists"))?;
        let chart = slide
            .charts()
            .get(self.chart_index)
            .ok_or_else(|| value_error("Chart no longer exists"))?;
        Ok(chart.additional_series().len())
    }

    /// Add a series with name and values.
    pub fn add_series(&mut self, name: &str, values: Vec<f64>) -> PyResult<()> {
        let mut pres = self
            .presentation
            .lock()
            .map_err(|e| value_error(format!("Failed to lock presentation: {}", e)))?;
        let slide = pres
            .slides_mut()
            .get_mut(self.slide_index)
            .ok_or_else(|| value_error("Slide no longer exists"))?;
        let chart = slide
            .charts_mut()
            .get_mut(self.chart_index)
            .ok_or_else(|| value_error("Chart no longer exists"))?;
        let mut series = ChartSeries::new(name);
        series.set_values(values);
        chart.add_series(series);
        Ok(())
    }

    /// Remove a series by index.
    pub fn remove_series(&mut self, index: usize) -> PyResult<bool> {
        let mut pres = self
            .presentation
            .lock()
            .map_err(|e| value_error(format!("Failed to lock presentation: {}", e)))?;
        let slide = pres
            .slides_mut()
            .get_mut(self.slide_index)
            .ok_or_else(|| value_error("Slide no longer exists"))?;
        let chart = slide
            .charts_mut()
            .get_mut(self.chart_index)
            .ok_or_else(|| value_error("Chart no longer exists"))?;
        Ok(chart.remove_series(index).is_some())
    }

    /// Set category axis title.
    pub fn set_category_axis_title(&mut self, title: &str) -> PyResult<()> {
        let mut pres = self
            .presentation
            .lock()
            .map_err(|e| value_error(format!("Failed to lock presentation: {}", e)))?;
        let slide = pres
            .slides_mut()
            .get_mut(self.slide_index)
            .ok_or_else(|| value_error("Slide no longer exists"))?;
        let chart = slide
            .charts_mut()
            .get_mut(self.chart_index)
            .ok_or_else(|| value_error("Chart no longer exists"))?;
        chart.set_category_axis_title(title);
        Ok(())
    }

    /// Set value axis title.
    pub fn set_value_axis_title(&mut self, title: &str) -> PyResult<()> {
        let mut pres = self
            .presentation
            .lock()
            .map_err(|e| value_error(format!("Failed to lock presentation: {}", e)))?;
        let slide = pres
            .slides_mut()
            .get_mut(self.slide_index)
            .ok_or_else(|| value_error("Slide no longer exists"))?;
        let chart = slide
            .charts_mut()
            .get_mut(self.chart_index)
            .ok_or_else(|| value_error("Chart no longer exists"))?;
        chart.set_value_axis_title(title);
        Ok(())
    }
}

// =============================================================================
// PPTX Bindings - Image Wrapper
// =============================================================================

/// Python wrapper for `offidized_pptx::Image`.
#[pyclass(module = "offidized._native", name = "Image")]
#[derive(Clone)]
pub struct Image {
    presentation: Arc<Mutex<CorePresentation>>,
    slide_index: usize,
    image_index: usize,
}

#[pymethods]
impl Image {
    /// Get image name.
    pub fn name(&self) -> PyResult<Option<String>> {
        let pres = self
            .presentation
            .lock()
            .map_err(|e| value_error(format!("Failed to lock presentation: {}", e)))?;
        let slide = pres
            .slides()
            .get(self.slide_index)
            .ok_or_else(|| value_error("Slide no longer exists"))?;
        let image = slide
            .images()
            .get(self.image_index)
            .ok_or_else(|| value_error("Image no longer exists"))?;
        Ok(image.name().map(String::from))
    }

    /// Get content type.
    pub fn content_type(&self) -> PyResult<String> {
        let pres = self
            .presentation
            .lock()
            .map_err(|e| value_error(format!("Failed to lock presentation: {}", e)))?;
        let slide = pres
            .slides()
            .get(self.slide_index)
            .ok_or_else(|| value_error("Slide no longer exists"))?;
        let image = slide
            .images()
            .get(self.image_index)
            .ok_or_else(|| value_error("Image no longer exists"))?;
        Ok(image.content_type().to_string())
    }

    /// Get image bytes.
    pub fn bytes(&self) -> PyResult<Vec<u8>> {
        let pres = self
            .presentation
            .lock()
            .map_err(|e| value_error(format!("Failed to lock presentation: {}", e)))?;
        let slide = pres
            .slides()
            .get(self.slide_index)
            .ok_or_else(|| value_error("Slide no longer exists"))?;
        let image = slide
            .images()
            .get(self.image_index)
            .ok_or_else(|| value_error("Image no longer exists"))?;
        Ok(image.bytes().to_vec())
    }

    /// Get crop as (left, top, right, bottom) percentages, or None.
    pub fn crop(&self) -> PyResult<Option<(f64, f64, f64, f64)>> {
        let pres = self
            .presentation
            .lock()
            .map_err(|e| value_error(format!("Failed to lock presentation: {}", e)))?;
        let slide = pres
            .slides()
            .get(self.slide_index)
            .ok_or_else(|| value_error("Slide no longer exists"))?;
        let image = slide
            .images()
            .get(self.image_index)
            .ok_or_else(|| value_error("Image no longer exists"))?;
        Ok(image.crop().map(|c| (c.left, c.top, c.right, c.bottom)))
    }

    /// Get transparency (0.0 = opaque, 1.0 = fully transparent), or None.
    pub fn transparency(&self) -> PyResult<Option<f64>> {
        let pres = self
            .presentation
            .lock()
            .map_err(|e| value_error(format!("Failed to lock presentation: {}", e)))?;
        let slide = pres
            .slides()
            .get(self.slide_index)
            .ok_or_else(|| value_error("Slide no longer exists"))?;
        let image = slide
            .images()
            .get(self.image_index)
            .ok_or_else(|| value_error("Image no longer exists"))?;
        Ok(image.transparency())
    }

    /// Set transparency.
    pub fn set_transparency(&mut self, alpha: f64) -> PyResult<()> {
        let mut pres = self
            .presentation
            .lock()
            .map_err(|e| value_error(format!("Failed to lock presentation: {}", e)))?;
        let slide = pres
            .slides_mut()
            .get_mut(self.slide_index)
            .ok_or_else(|| value_error("Slide no longer exists"))?;
        let image = slide
            .images_mut()
            .get_mut(self.image_index)
            .ok_or_else(|| value_error("Image no longer exists"))?;
        image.set_transparency(alpha);
        Ok(())
    }
}

// =============================================================================
// PPTX Bindings - Text Formatting (Shape, ShapeParagraph, TextRun)
// =============================================================================

/// Python wrapper for a shape on a slide, providing access to paragraphs and text.
#[pyclass(module = "offidized._native", name = "Shape")]
#[derive(Clone)]
pub struct PyShape {
    presentation: Arc<Mutex<CorePresentation>>,
    slide_index: usize,
    shape_index: usize,
}

#[pymethods]
impl PyShape {
    /// Get the shape name.
    pub fn name(&self) -> PyResult<String> {
        let pres = self
            .presentation
            .lock()
            .map_err(|e| value_error(format!("Failed to lock presentation: {}", e)))?;
        let slide = pres
            .slides()
            .get(self.slide_index)
            .ok_or_else(|| value_error("Slide no longer exists"))?;
        let shape = slide
            .shapes()
            .get(self.shape_index)
            .ok_or_else(|| value_error("Shape no longer exists"))?;
        Ok(shape.name().to_string())
    }

    /// Get number of paragraphs in the shape's text frame.
    pub fn paragraph_count(&self) -> PyResult<usize> {
        let pres = self
            .presentation
            .lock()
            .map_err(|e| value_error(format!("Failed to lock presentation: {}", e)))?;
        let slide = pres
            .slides()
            .get(self.slide_index)
            .ok_or_else(|| value_error("Slide no longer exists"))?;
        let shape = slide
            .shapes()
            .get(self.shape_index)
            .ok_or_else(|| value_error("Shape no longer exists"))?;
        Ok(shape.paragraph_count())
    }

    /// Get a paragraph by index (0-based).
    pub fn get_paragraph(&self, index: usize) -> PyResult<ShapeParagraph> {
        let pres = self
            .presentation
            .lock()
            .map_err(|e| value_error(format!("Failed to lock presentation: {}", e)))?;
        let slide = pres
            .slides()
            .get(self.slide_index)
            .ok_or_else(|| value_error("Slide no longer exists"))?;
        let shape = slide
            .shapes()
            .get(self.shape_index)
            .ok_or_else(|| value_error("Shape no longer exists"))?;
        if index >= shape.paragraph_count() {
            return Err(value_error(format!(
                "Paragraph index {} out of range (0..{})",
                index,
                shape.paragraph_count()
            )));
        }
        Ok(ShapeParagraph {
            presentation: Arc::clone(&self.presentation),
            slide_index: self.slide_index,
            shape_index: self.shape_index,
            paragraph_index: index,
        })
    }

    /// Add a new empty paragraph and return it.
    pub fn add_paragraph(&mut self) -> PyResult<ShapeParagraph> {
        let mut pres = self
            .presentation
            .lock()
            .map_err(|e| value_error(format!("Failed to lock presentation: {}", e)))?;
        let slide = pres
            .slides_mut()
            .get_mut(self.slide_index)
            .ok_or_else(|| value_error("Slide no longer exists"))?;
        let shape = slide
            .shapes_mut()
            .get_mut(self.shape_index)
            .ok_or_else(|| value_error("Shape no longer exists"))?;
        let _ = shape.add_paragraph();
        let paragraph_index = shape.paragraph_count().saturating_sub(1);
        Ok(ShapeParagraph {
            presentation: Arc::clone(&self.presentation),
            slide_index: self.slide_index,
            shape_index: self.shape_index,
            paragraph_index,
        })
    }

    /// Add a paragraph with text and return it.
    pub fn add_paragraph_with_text(&mut self, text: &str) -> PyResult<ShapeParagraph> {
        let mut pres = self
            .presentation
            .lock()
            .map_err(|e| value_error(format!("Failed to lock presentation: {}", e)))?;
        let slide = pres
            .slides_mut()
            .get_mut(self.slide_index)
            .ok_or_else(|| value_error("Slide no longer exists"))?;
        let shape = slide
            .shapes_mut()
            .get_mut(self.shape_index)
            .ok_or_else(|| value_error("Shape no longer exists"))?;
        let _ = shape.add_paragraph_with_text(text);
        let paragraph_index = shape.paragraph_count().saturating_sub(1);
        Ok(ShapeParagraph {
            presentation: Arc::clone(&self.presentation),
            slide_index: self.slide_index,
            shape_index: self.shape_index,
            paragraph_index,
        })
    }

    // ── New shape methods ──

    /// Get solid fill color as sRGB hex, or None.
    pub fn solid_fill_srgb(&self) -> PyResult<Option<String>> {
        let pres = self
            .presentation
            .lock()
            .map_err(|e| value_error(format!("Failed to lock presentation: {}", e)))?;
        let slide = pres
            .slides()
            .get(self.slide_index)
            .ok_or_else(|| value_error("Slide no longer exists"))?;
        let shape = slide
            .shapes()
            .get(self.shape_index)
            .ok_or_else(|| value_error("Shape no longer exists"))?;
        Ok(shape.solid_fill_srgb().map(String::from))
    }

    /// Set solid fill color as sRGB hex.
    pub fn set_solid_fill_srgb(&mut self, color: &str) -> PyResult<()> {
        let mut pres = self
            .presentation
            .lock()
            .map_err(|e| value_error(format!("Failed to lock presentation: {}", e)))?;
        let slide = pres
            .slides_mut()
            .get_mut(self.slide_index)
            .ok_or_else(|| value_error("Slide no longer exists"))?;
        let shape = slide
            .shapes_mut()
            .get_mut(self.shape_index)
            .ok_or_else(|| value_error("Shape no longer exists"))?;
        shape.set_solid_fill_srgb(color);
        Ok(())
    }

    /// Clear solid fill.
    pub fn clear_solid_fill(&mut self) -> PyResult<()> {
        let mut pres = self
            .presentation
            .lock()
            .map_err(|e| value_error(format!("Failed to lock presentation: {}", e)))?;
        let slide = pres
            .slides_mut()
            .get_mut(self.slide_index)
            .ok_or_else(|| value_error("Slide no longer exists"))?;
        let shape = slide
            .shapes_mut()
            .get_mut(self.shape_index)
            .ok_or_else(|| value_error("Shape no longer exists"))?;
        shape.clear_solid_fill_srgb();
        Ok(())
    }

    /// Get alt text.
    pub fn alt_text(&self) -> PyResult<Option<String>> {
        let pres = self
            .presentation
            .lock()
            .map_err(|e| value_error(format!("Failed to lock presentation: {}", e)))?;
        let slide = pres
            .slides()
            .get(self.slide_index)
            .ok_or_else(|| value_error("Slide no longer exists"))?;
        let shape = slide
            .shapes()
            .get(self.shape_index)
            .ok_or_else(|| value_error("Shape no longer exists"))?;
        Ok(shape.alt_text().map(String::from))
    }

    /// Set alt text.
    pub fn set_alt_text(&mut self, text: &str) -> PyResult<()> {
        let mut pres = self
            .presentation
            .lock()
            .map_err(|e| value_error(format!("Failed to lock presentation: {}", e)))?;
        let slide = pres
            .slides_mut()
            .get_mut(self.slide_index)
            .ok_or_else(|| value_error("Slide no longer exists"))?;
        let shape = slide
            .shapes_mut()
            .get_mut(self.shape_index)
            .ok_or_else(|| value_error("Shape no longer exists"))?;
        shape.set_alt_text(text);
        Ok(())
    }

    /// Get shape position and size as (x, y, width, height) in EMUs.
    pub fn geometry(&self) -> PyResult<Option<(i64, i64, i64, i64)>> {
        let pres = self
            .presentation
            .lock()
            .map_err(|e| value_error(format!("Failed to lock presentation: {}", e)))?;
        let slide = pres
            .slides()
            .get(self.slide_index)
            .ok_or_else(|| value_error("Slide no longer exists"))?;
        let shape = slide
            .shapes()
            .get(self.shape_index)
            .ok_or_else(|| value_error("Shape no longer exists"))?;
        Ok(shape.geometry().map(|g| (g.x(), g.y(), g.cx(), g.cy())))
    }

    /// Set shape position and size in EMUs (English Metric Units).
    ///
    /// Parameters:
    ///   x: Left edge position in EMUs (1 inch = 914400 EMUs)
    ///   y: Top edge position in EMUs
    ///   width: Shape width in EMUs
    ///   height: Shape height in EMUs
    pub fn set_geometry(&mut self, x: i64, y: i64, width: i64, height: i64) -> PyResult<()> {
        use offidized_pptx::ShapeGeometry;
        let mut pres = self
            .presentation
            .lock()
            .map_err(|e| value_error(format!("Failed to lock presentation: {}", e)))?;
        let slide = pres
            .slides_mut()
            .get_mut(self.slide_index)
            .ok_or_else(|| value_error("Slide no longer exists"))?;
        let shape = slide
            .shapes_mut()
            .get_mut(self.shape_index)
            .ok_or_else(|| value_error("Shape no longer exists"))?;
        shape.set_geometry(ShapeGeometry::new(x, y, width, height));
        Ok(())
    }

    /// Get preset geometry name.
    pub fn preset_geometry(&self) -> PyResult<Option<String>> {
        let pres = self
            .presentation
            .lock()
            .map_err(|e| value_error(format!("Failed to lock presentation: {}", e)))?;
        let slide = pres
            .slides()
            .get(self.slide_index)
            .ok_or_else(|| value_error("Slide no longer exists"))?;
        let shape = slide
            .shapes()
            .get(self.shape_index)
            .ok_or_else(|| value_error("Shape no longer exists"))?;
        Ok(shape.preset_geometry().map(String::from))
    }

    /// Set preset geometry.
    pub fn set_preset_geometry(&mut self, geometry: &str) -> PyResult<()> {
        let mut pres = self
            .presentation
            .lock()
            .map_err(|e| value_error(format!("Failed to lock presentation: {}", e)))?;
        let slide = pres
            .slides_mut()
            .get_mut(self.slide_index)
            .ok_or_else(|| value_error("Slide no longer exists"))?;
        let shape = slide
            .shapes_mut()
            .get_mut(self.shape_index)
            .ok_or_else(|| value_error("Shape no longer exists"))?;
        shape.set_preset_geometry(geometry);
        Ok(())
    }

    /// Check if shape is SmartArt.
    pub fn is_smartart(&self) -> PyResult<bool> {
        let pres = self
            .presentation
            .lock()
            .map_err(|e| value_error(format!("Failed to lock presentation: {}", e)))?;
        let slide = pres
            .slides()
            .get(self.slide_index)
            .ok_or_else(|| value_error("Slide no longer exists"))?;
        let shape = slide
            .shapes()
            .get(self.shape_index)
            .ok_or_else(|| value_error("Shape no longer exists"))?;
        Ok(shape.is_smartart())
    }

    /// Get word wrap setting.
    pub fn word_wrap(&self) -> PyResult<Option<bool>> {
        let pres = self
            .presentation
            .lock()
            .map_err(|e| value_error(format!("Failed to lock presentation: {}", e)))?;
        let slide = pres
            .slides()
            .get(self.slide_index)
            .ok_or_else(|| value_error("Slide no longer exists"))?;
        let shape = slide
            .shapes()
            .get(self.shape_index)
            .ok_or_else(|| value_error("Shape no longer exists"))?;
        Ok(shape.word_wrap())
    }

    /// Set word wrap.
    pub fn set_word_wrap(&mut self, wrap: bool) -> PyResult<()> {
        let mut pres = self
            .presentation
            .lock()
            .map_err(|e| value_error(format!("Failed to lock presentation: {}", e)))?;
        let slide = pres
            .slides_mut()
            .get_mut(self.slide_index)
            .ok_or_else(|| value_error("Slide no longer exists"))?;
        let shape = slide
            .shapes_mut()
            .get_mut(self.shape_index)
            .ok_or_else(|| value_error("Shape no longer exists"))?;
        shape.set_word_wrap(wrap);
        Ok(())
    }

    // ── Shape outline ──

    /// Set shape outline/border with color and width.
    ///
    /// Parameters:
    ///   color: sRGB hex color (e.g. "FF0000")
    ///   width_pt: Line width in points (default 1.0)
    ///   dash: Dash style - "solid", "dot", "dash", "lgDash", "dashDot",
    ///         "lgDashDot", "lgDashDotDot", "sysDash", "sysDot",
    ///         "sysDashDot", "sysDashDotDot" (default "solid")
    #[pyo3(signature = (color, width_pt=1.0, dash="solid"))]
    pub fn set_outline(&mut self, color: &str, width_pt: f64, dash: &str) -> PyResult<()> {
        use offidized_pptx::{LineDashStyle, ShapeOutline};
        let mut pres = self
            .presentation
            .lock()
            .map_err(|e| value_error(format!("Failed to lock presentation: {}", e)))?;
        let slide = pres
            .slides_mut()
            .get_mut(self.slide_index)
            .ok_or_else(|| value_error("Slide no longer exists"))?;
        let shape = slide
            .shapes_mut()
            .get_mut(self.shape_index)
            .ok_or_else(|| value_error("Shape no longer exists"))?;
        let dash_style = LineDashStyle::from_xml(dash);
        let mut outline = ShapeOutline::new();
        outline.width_emu = Some((width_pt * 12700.0) as i64);
        outline.color_srgb = Some(color.to_string());
        outline.dash_style = dash_style;
        shape.set_outline(outline);
        Ok(())
    }

    /// Clear the shape outline.
    pub fn clear_outline(&mut self) -> PyResult<()> {
        let mut pres = self
            .presentation
            .lock()
            .map_err(|e| value_error(format!("Failed to lock presentation: {}", e)))?;
        let slide = pres
            .slides_mut()
            .get_mut(self.slide_index)
            .ok_or_else(|| value_error("Slide no longer exists"))?;
        let shape = slide
            .shapes_mut()
            .get_mut(self.shape_index)
            .ok_or_else(|| value_error("Shape no longer exists"))?;
        shape.clear_outline();
        Ok(())
    }

    // ── Gradient fill ──

    /// Set a linear gradient fill on the shape.
    ///
    /// Parameters:
    ///   stops: List of (position_pct, srgb_hex) tuples. Position is 0-100.
    ///   angle: Gradient angle in degrees (0 = left-to-right, 90 = top-to-bottom).
    #[pyo3(signature = (stops, angle=0.0))]
    pub fn set_gradient_fill(&mut self, stops: Vec<(f64, String)>, angle: f64) -> PyResult<()> {
        use offidized_pptx::{GradientFill, GradientFillType, GradientStop};
        let mut pres = self
            .presentation
            .lock()
            .map_err(|e| value_error(format!("Failed to lock presentation: {}", e)))?;
        let slide = pres
            .slides_mut()
            .get_mut(self.slide_index)
            .ok_or_else(|| value_error("Slide no longer exists"))?;
        let shape = slide
            .shapes_mut()
            .get_mut(self.shape_index)
            .ok_or_else(|| value_error("Shape no longer exists"))?;
        let mut fill = GradientFill::new();
        fill.fill_type = Some(GradientFillType::Linear);
        fill.linear_angle = Some((angle * 60000.0) as i32);
        fill.stops = stops
            .into_iter()
            .map(|(pos, color)| GradientStop {
                position: (pos * 1000.0) as u32,
                color_srgb: color,
                color: None,
            })
            .collect();
        shape.set_gradient_fill(fill);
        Ok(())
    }

    // ── Shape rotation ──

    /// Get shape rotation in degrees.
    pub fn rotation(&self) -> PyResult<Option<f64>> {
        let pres = self
            .presentation
            .lock()
            .map_err(|e| value_error(format!("Failed to lock presentation: {}", e)))?;
        let slide = pres
            .slides()
            .get(self.slide_index)
            .ok_or_else(|| value_error("Slide no longer exists"))?;
        let shape = slide
            .shapes()
            .get(self.shape_index)
            .ok_or_else(|| value_error("Shape no longer exists"))?;
        Ok(shape.rotation().map(|r| r as f64 / 60000.0))
    }

    /// Set shape rotation in degrees (0-360).
    pub fn set_rotation(&mut self, degrees: f64) -> PyResult<()> {
        let mut pres = self
            .presentation
            .lock()
            .map_err(|e| value_error(format!("Failed to lock presentation: {}", e)))?;
        let slide = pres
            .slides_mut()
            .get_mut(self.slide_index)
            .ok_or_else(|| value_error("Slide no longer exists"))?;
        let shape = slide
            .shapes_mut()
            .get_mut(self.shape_index)
            .ok_or_else(|| value_error("Shape no longer exists"))?;
        shape.set_rotation((degrees * 60000.0) as i32);
        Ok(())
    }

    // ── Flip ──

    /// Get horizontal flip state.
    pub fn flip_h(&self) -> PyResult<bool> {
        let pres = self
            .presentation
            .lock()
            .map_err(|e| value_error(format!("Failed to lock presentation: {}", e)))?;
        let slide = pres
            .slides()
            .get(self.slide_index)
            .ok_or_else(|| value_error("Slide no longer exists"))?;
        let shape = slide
            .shapes()
            .get(self.shape_index)
            .ok_or_else(|| value_error("Shape no longer exists"))?;
        Ok(shape.flip_h())
    }

    /// Set horizontal flip.
    pub fn set_flip_h(&mut self, flip: bool) -> PyResult<()> {
        let mut pres = self
            .presentation
            .lock()
            .map_err(|e| value_error(format!("Failed to lock presentation: {}", e)))?;
        let slide = pres
            .slides_mut()
            .get_mut(self.slide_index)
            .ok_or_else(|| value_error("Slide no longer exists"))?;
        let shape = slide
            .shapes_mut()
            .get_mut(self.shape_index)
            .ok_or_else(|| value_error("Shape no longer exists"))?;
        shape.set_flip_h(flip);
        Ok(())
    }

    /// Get vertical flip state.
    pub fn flip_v(&self) -> PyResult<bool> {
        let pres = self
            .presentation
            .lock()
            .map_err(|e| value_error(format!("Failed to lock presentation: {}", e)))?;
        let slide = pres
            .slides()
            .get(self.slide_index)
            .ok_or_else(|| value_error("Slide no longer exists"))?;
        let shape = slide
            .shapes()
            .get(self.shape_index)
            .ok_or_else(|| value_error("Shape no longer exists"))?;
        Ok(shape.flip_v())
    }

    /// Set vertical flip.
    pub fn set_flip_v(&mut self, flip: bool) -> PyResult<()> {
        let mut pres = self
            .presentation
            .lock()
            .map_err(|e| value_error(format!("Failed to lock presentation: {}", e)))?;
        let slide = pres
            .slides_mut()
            .get_mut(self.slide_index)
            .ok_or_else(|| value_error("Slide no longer exists"))?;
        let shape = slide
            .shapes_mut()
            .get_mut(self.shape_index)
            .ok_or_else(|| value_error("Shape no longer exists"))?;
        shape.set_flip_v(flip);
        Ok(())
    }

    // ── Pattern fill ──

    /// Get pattern fill info as (pattern_type, foreground_srgb, background_srgb), or None.
    pub fn pattern_fill(&self) -> PyResult<Option<(String, Option<String>, Option<String>)>> {
        let pres = self
            .presentation
            .lock()
            .map_err(|e| value_error(format!("Failed to lock presentation: {}", e)))?;
        let slide = pres
            .slides()
            .get(self.slide_index)
            .ok_or_else(|| value_error("Slide no longer exists"))?;
        let shape = slide
            .shapes()
            .get(self.shape_index)
            .ok_or_else(|| value_error("Shape no longer exists"))?;
        Ok(shape.pattern_fill().map(|pf| {
            (
                pf.pattern_type.to_xml().to_string(),
                pf.foreground_srgb.clone(),
                pf.background_srgb.clone(),
            )
        }))
    }

    /// Set pattern fill.
    #[pyo3(signature = (pattern_type, foreground=None, background=None))]
    pub fn set_pattern_fill(
        &mut self,
        pattern_type: &str,
        foreground: Option<&str>,
        background: Option<&str>,
    ) -> PyResult<()> {
        let mut pres = self
            .presentation
            .lock()
            .map_err(|e| value_error(format!("Failed to lock presentation: {}", e)))?;
        let slide = pres
            .slides_mut()
            .get_mut(self.slide_index)
            .ok_or_else(|| value_error("Slide no longer exists"))?;
        let shape = slide
            .shapes_mut()
            .get_mut(self.shape_index)
            .ok_or_else(|| value_error("Shape no longer exists"))?;
        let mut fill = PatternFill::new(PatternFillType::from_xml(pattern_type));
        if let Some(fg) = foreground {
            fill.foreground_srgb = Some(fg.to_string());
        }
        if let Some(bg) = background {
            fill.background_srgb = Some(bg.to_string());
        }
        shape.set_pattern_fill(fill);
        Ok(())
    }

    /// Clear pattern fill.
    pub fn clear_pattern_fill(&mut self) -> PyResult<()> {
        let mut pres = self
            .presentation
            .lock()
            .map_err(|e| value_error(format!("Failed to lock presentation: {}", e)))?;
        let slide = pres
            .slides_mut()
            .get_mut(self.slide_index)
            .ok_or_else(|| value_error("Slide no longer exists"))?;
        let shape = slide
            .shapes_mut()
            .get_mut(self.shape_index)
            .ok_or_else(|| value_error("Shape no longer exists"))?;
        shape.clear_pattern_fill();
        Ok(())
    }

    // ── Shadow ──

    /// Get shadow as (offset_x, offset_y, blur_radius, color, alpha), or None.
    #[allow(clippy::type_complexity)]
    pub fn shadow(&self) -> PyResult<Option<(i64, i64, i64, String, Option<u8>)>> {
        let pres = self
            .presentation
            .lock()
            .map_err(|e| value_error(format!("Failed to lock presentation: {}", e)))?;
        let slide = pres
            .slides()
            .get(self.slide_index)
            .ok_or_else(|| value_error("Slide no longer exists"))?;
        let shape = slide
            .shapes()
            .get(self.shape_index)
            .ok_or_else(|| value_error("Shape no longer exists"))?;
        Ok(shape.shadow().map(|s| {
            (
                s.offset_x,
                s.offset_y,
                s.blur_radius,
                s.color.clone(),
                s.alpha,
            )
        }))
    }

    /// Set shadow.
    pub fn set_shadow(
        &mut self,
        offset_x: i64,
        offset_y: i64,
        blur_radius: i64,
        color: &str,
    ) -> PyResult<()> {
        let mut pres = self
            .presentation
            .lock()
            .map_err(|e| value_error(format!("Failed to lock presentation: {}", e)))?;
        let slide = pres
            .slides_mut()
            .get_mut(self.slide_index)
            .ok_or_else(|| value_error("Slide no longer exists"))?;
        let shape = slide
            .shapes_mut()
            .get_mut(self.shape_index)
            .ok_or_else(|| value_error("Shape no longer exists"))?;
        shape.set_shadow(ShapeShadow::new(offset_x, offset_y, blur_radius, color));
        Ok(())
    }

    /// Clear shadow.
    pub fn clear_shadow(&mut self) -> PyResult<()> {
        let mut pres = self
            .presentation
            .lock()
            .map_err(|e| value_error(format!("Failed to lock presentation: {}", e)))?;
        let slide = pres
            .slides_mut()
            .get_mut(self.slide_index)
            .ok_or_else(|| value_error("Slide no longer exists"))?;
        let shape = slide
            .shapes_mut()
            .get_mut(self.shape_index)
            .ok_or_else(|| value_error("Shape no longer exists"))?;
        shape.clear_shadow();
        Ok(())
    }

    // ── Glow ──

    /// Get glow as (radius, color, alpha), or None.
    pub fn glow(&self) -> PyResult<Option<(i64, String, Option<u8>)>> {
        let pres = self
            .presentation
            .lock()
            .map_err(|e| value_error(format!("Failed to lock presentation: {}", e)))?;
        let slide = pres
            .slides()
            .get(self.slide_index)
            .ok_or_else(|| value_error("Slide no longer exists"))?;
        let shape = slide
            .shapes()
            .get(self.shape_index)
            .ok_or_else(|| value_error("Shape no longer exists"))?;
        Ok(shape.glow().map(|g| (g.radius, g.color.clone(), g.alpha)))
    }

    /// Set glow.
    pub fn set_glow(&mut self, radius: i64, color: &str) -> PyResult<()> {
        let mut pres = self
            .presentation
            .lock()
            .map_err(|e| value_error(format!("Failed to lock presentation: {}", e)))?;
        let slide = pres
            .slides_mut()
            .get_mut(self.slide_index)
            .ok_or_else(|| value_error("Slide no longer exists"))?;
        let shape = slide
            .shapes_mut()
            .get_mut(self.shape_index)
            .ok_or_else(|| value_error("Shape no longer exists"))?;
        shape.set_glow(ShapeGlow::new(radius, color));
        Ok(())
    }

    /// Clear glow.
    pub fn clear_glow(&mut self) -> PyResult<()> {
        let mut pres = self
            .presentation
            .lock()
            .map_err(|e| value_error(format!("Failed to lock presentation: {}", e)))?;
        let slide = pres
            .slides_mut()
            .get_mut(self.slide_index)
            .ok_or_else(|| value_error("Slide no longer exists"))?;
        let shape = slide
            .shapes_mut()
            .get_mut(self.shape_index)
            .ok_or_else(|| value_error("Shape no longer exists"))?;
        shape.clear_glow();
        Ok(())
    }

    // ── Reflection ──

    /// Get reflection as (blur_radius, distance, start_alpha, end_alpha), or None.
    #[allow(clippy::type_complexity)]
    pub fn reflection(&self) -> PyResult<Option<(i64, i64, Option<u8>, Option<u8>)>> {
        let pres = self
            .presentation
            .lock()
            .map_err(|e| value_error(format!("Failed to lock presentation: {}", e)))?;
        let slide = pres
            .slides()
            .get(self.slide_index)
            .ok_or_else(|| value_error("Slide no longer exists"))?;
        let shape = slide
            .shapes()
            .get(self.shape_index)
            .ok_or_else(|| value_error("Shape no longer exists"))?;
        Ok(shape
            .reflection()
            .map(|r| (r.blur_radius, r.distance, r.start_alpha, r.end_alpha)))
    }

    /// Set reflection.
    pub fn set_reflection(&mut self, blur_radius: i64, distance: i64) -> PyResult<()> {
        let mut pres = self
            .presentation
            .lock()
            .map_err(|e| value_error(format!("Failed to lock presentation: {}", e)))?;
        let slide = pres
            .slides_mut()
            .get_mut(self.slide_index)
            .ok_or_else(|| value_error("Slide no longer exists"))?;
        let shape = slide
            .shapes_mut()
            .get_mut(self.shape_index)
            .ok_or_else(|| value_error("Shape no longer exists"))?;
        shape.set_reflection(ShapeReflection::new(blur_radius, distance));
        Ok(())
    }

    /// Clear reflection.
    pub fn clear_reflection(&mut self) -> PyResult<()> {
        let mut pres = self
            .presentation
            .lock()
            .map_err(|e| value_error(format!("Failed to lock presentation: {}", e)))?;
        let slide = pres
            .slides_mut()
            .get_mut(self.slide_index)
            .ok_or_else(|| value_error("Slide no longer exists"))?;
        let shape = slide
            .shapes_mut()
            .get_mut(self.shape_index)
            .ok_or_else(|| value_error("Shape no longer exists"))?;
        shape.clear_reflection();
        Ok(())
    }

    // ── Placeholder ──

    /// Get placeholder kind (e.g. "title", "body"), or None.
    pub fn placeholder_kind(&self) -> PyResult<Option<String>> {
        let pres = self
            .presentation
            .lock()
            .map_err(|e| value_error(format!("Failed to lock presentation: {}", e)))?;
        let slide = pres
            .slides()
            .get(self.slide_index)
            .ok_or_else(|| value_error("Slide no longer exists"))?;
        let shape = slide
            .shapes()
            .get(self.shape_index)
            .ok_or_else(|| value_error("Shape no longer exists"))?;
        Ok(shape.placeholder_kind().map(String::from))
    }

    /// Set placeholder kind.
    pub fn set_placeholder_kind(&mut self, kind: &str) -> PyResult<()> {
        let mut pres = self
            .presentation
            .lock()
            .map_err(|e| value_error(format!("Failed to lock presentation: {}", e)))?;
        let slide = pres
            .slides_mut()
            .get_mut(self.slide_index)
            .ok_or_else(|| value_error("Slide no longer exists"))?;
        let shape = slide
            .shapes_mut()
            .get_mut(self.shape_index)
            .ok_or_else(|| value_error("Shape no longer exists"))?;
        shape.set_placeholder_kind(kind);
        Ok(())
    }

    /// Clear placeholder kind.
    pub fn clear_placeholder_kind(&mut self) -> PyResult<()> {
        let mut pres = self
            .presentation
            .lock()
            .map_err(|e| value_error(format!("Failed to lock presentation: {}", e)))?;
        let slide = pres
            .slides_mut()
            .get_mut(self.slide_index)
            .ok_or_else(|| value_error("Slide no longer exists"))?;
        let shape = slide
            .shapes_mut()
            .get_mut(self.shape_index)
            .ok_or_else(|| value_error("Shape no longer exists"))?;
        shape.clear_placeholder_kind();
        Ok(())
    }

    /// Get placeholder index, or None.
    pub fn placeholder_idx(&self) -> PyResult<Option<u32>> {
        let pres = self
            .presentation
            .lock()
            .map_err(|e| value_error(format!("Failed to lock presentation: {}", e)))?;
        let slide = pres
            .slides()
            .get(self.slide_index)
            .ok_or_else(|| value_error("Slide no longer exists"))?;
        let shape = slide
            .shapes()
            .get(self.shape_index)
            .ok_or_else(|| value_error("Shape no longer exists"))?;
        Ok(shape.placeholder_idx())
    }

    /// Set placeholder index.
    pub fn set_placeholder_idx(&mut self, idx: u32) -> PyResult<()> {
        let mut pres = self
            .presentation
            .lock()
            .map_err(|e| value_error(format!("Failed to lock presentation: {}", e)))?;
        let slide = pres
            .slides_mut()
            .get_mut(self.slide_index)
            .ok_or_else(|| value_error("Slide no longer exists"))?;
        let shape = slide
            .shapes_mut()
            .get_mut(self.shape_index)
            .ok_or_else(|| value_error("Shape no longer exists"))?;
        shape.set_placeholder_idx(idx);
        Ok(())
    }
}

/// Python wrapper for `offidized_pptx::ShapeParagraph`.
#[pyclass(module = "offidized._native", name = "ShapeParagraph")]
#[derive(Clone)]
pub struct ShapeParagraph {
    presentation: Arc<Mutex<CorePresentation>>,
    slide_index: usize,
    shape_index: usize,
    paragraph_index: usize,
}

#[pymethods]
impl ShapeParagraph {
    /// Get number of text runs in this paragraph.
    pub fn run_count(&self) -> PyResult<usize> {
        let pres = self
            .presentation
            .lock()
            .map_err(|e| value_error(format!("Failed to lock presentation: {}", e)))?;
        let para = self.get_paragraph_ref(&pres)?;
        Ok(para.run_count())
    }

    /// Get a text run by index (0-based).
    pub fn get_run(&self, index: usize) -> PyResult<TextRun> {
        let pres = self
            .presentation
            .lock()
            .map_err(|e| value_error(format!("Failed to lock presentation: {}", e)))?;
        let para = self.get_paragraph_ref(&pres)?;
        if index >= para.run_count() {
            return Err(value_error(format!(
                "Run index {} out of range (0..{})",
                index,
                para.run_count()
            )));
        }
        Ok(TextRun {
            presentation: Arc::clone(&self.presentation),
            slide_index: self.slide_index,
            shape_index: self.shape_index,
            paragraph_index: self.paragraph_index,
            run_index: index,
        })
    }

    /// Add a text run with the given text and return it.
    pub fn add_run(&mut self, text: &str) -> PyResult<TextRun> {
        let mut pres = self
            .presentation
            .lock()
            .map_err(|e| value_error(format!("Failed to lock presentation: {}", e)))?;
        let para = self.get_paragraph_mut(&mut pres)?;
        let _ = para.add_run(text);
        let run_index = para.run_count().saturating_sub(1);
        Ok(TextRun {
            presentation: Arc::clone(&self.presentation),
            slide_index: self.slide_index,
            shape_index: self.shape_index,
            paragraph_index: self.paragraph_index,
            run_index,
        })
    }

    /// Get all text in this paragraph (concatenating all runs).
    pub fn text(&self) -> PyResult<String> {
        let pres = self
            .presentation
            .lock()
            .map_err(|e| value_error(format!("Failed to lock presentation: {}", e)))?;
        let para = self.get_paragraph_ref(&pres)?;
        Ok(para
            .runs()
            .iter()
            .map(|r| r.text())
            .collect::<Vec<_>>()
            .join(""))
    }

    /// Get text alignment ("left", "center", "right", "justified", "distributed", or None).
    pub fn alignment(&self) -> PyResult<Option<String>> {
        let pres = self
            .presentation
            .lock()
            .map_err(|e| value_error(format!("Failed to lock presentation: {}", e)))?;
        let para = self.get_paragraph_ref(&pres)?;
        Ok(para.alignment().map(|a| a.to_xml().to_string()))
    }

    /// Set text alignment ("left", "center", "right", "justified", "distributed").
    pub fn set_alignment(&mut self, alignment: &str) -> PyResult<()> {
        let align = CoreTextAlignment::from_xml(alignment)
            .ok_or_else(|| value_error(format!("Unknown alignment: {}", alignment)))?;
        let mut pres = self
            .presentation
            .lock()
            .map_err(|e| value_error(format!("Failed to lock presentation: {}", e)))?;
        let para = self.get_paragraph_mut(&mut pres)?;
        para.set_alignment(align);
        Ok(())
    }

    /// Get indentation level (0-8).
    pub fn level(&self) -> PyResult<Option<u32>> {
        let pres = self
            .presentation
            .lock()
            .map_err(|e| value_error(format!("Failed to lock presentation: {}", e)))?;
        let para = self.get_paragraph_ref(&pres)?;
        Ok(para.level())
    }

    /// Set indentation level (0-8).
    pub fn set_level(&mut self, level: u32) -> PyResult<()> {
        let mut pres = self
            .presentation
            .lock()
            .map_err(|e| value_error(format!("Failed to lock presentation: {}", e)))?;
        let para = self.get_paragraph_mut(&mut pres)?;
        para.set_level(level);
        Ok(())
    }

    /// Get line spacing as percentage (100000 = single, 150000 = 1.5x).
    pub fn line_spacing_pct(&self) -> PyResult<Option<u32>> {
        let pres = self
            .presentation
            .lock()
            .map_err(|e| value_error(format!("Failed to lock presentation: {}", e)))?;
        let para = self.get_paragraph_ref(&pres)?;
        Ok(para.properties().line_spacing_pct)
    }

    /// Set line spacing as percentage (100000 = single, 150000 = 1.5x).
    pub fn set_line_spacing_pct(&mut self, value: u32) -> PyResult<()> {
        let mut pres = self
            .presentation
            .lock()
            .map_err(|e| value_error(format!("Failed to lock presentation: {}", e)))?;
        let para = self.get_paragraph_mut(&mut pres)?;
        para.properties_mut().line_spacing_pct = Some(value);
        Ok(())
    }

    /// Get space before paragraph in hundredths of a point.
    pub fn space_before_pts(&self) -> PyResult<Option<u32>> {
        let pres = self
            .presentation
            .lock()
            .map_err(|e| value_error(format!("Failed to lock presentation: {}", e)))?;
        let para = self.get_paragraph_ref(&pres)?;
        Ok(para.properties().space_before_pts)
    }

    /// Set space before paragraph in hundredths of a point.
    pub fn set_space_before_pts(&mut self, value: u32) -> PyResult<()> {
        let mut pres = self
            .presentation
            .lock()
            .map_err(|e| value_error(format!("Failed to lock presentation: {}", e)))?;
        let para = self.get_paragraph_mut(&mut pres)?;
        para.properties_mut().space_before_pts = Some(value);
        Ok(())
    }

    /// Get space after paragraph in hundredths of a point.
    pub fn space_after_pts(&self) -> PyResult<Option<u32>> {
        let pres = self
            .presentation
            .lock()
            .map_err(|e| value_error(format!("Failed to lock presentation: {}", e)))?;
        let para = self.get_paragraph_ref(&pres)?;
        Ok(para.properties().space_after_pts)
    }

    /// Set space after paragraph in hundredths of a point.
    pub fn set_space_after_pts(&mut self, value: u32) -> PyResult<()> {
        let mut pres = self
            .presentation
            .lock()
            .map_err(|e| value_error(format!("Failed to lock presentation: {}", e)))?;
        let para = self.get_paragraph_mut(&mut pres)?;
        para.properties_mut().space_after_pts = Some(value);
        Ok(())
    }

    /// Get left margin in EMUs.
    pub fn margin_left(&self) -> PyResult<Option<i64>> {
        let pres = self
            .presentation
            .lock()
            .map_err(|e| value_error(format!("Failed to lock presentation: {}", e)))?;
        let para = self.get_paragraph_ref(&pres)?;
        Ok(para.properties().margin_left_emu)
    }

    /// Set left margin in EMUs.
    pub fn set_margin_left(&mut self, value: i64) -> PyResult<()> {
        let mut pres = self
            .presentation
            .lock()
            .map_err(|e| value_error(format!("Failed to lock presentation: {}", e)))?;
        let para = self.get_paragraph_mut(&mut pres)?;
        para.properties_mut().margin_left_emu = Some(value);
        Ok(())
    }

    /// Get first-line indent in EMUs (negative = hanging indent).
    pub fn indent(&self) -> PyResult<Option<i64>> {
        let pres = self
            .presentation
            .lock()
            .map_err(|e| value_error(format!("Failed to lock presentation: {}", e)))?;
        let para = self.get_paragraph_ref(&pres)?;
        Ok(para.properties().indent_emu)
    }

    /// Set first-line indent in EMUs (negative = hanging indent).
    pub fn set_indent(&mut self, value: i64) -> PyResult<()> {
        let mut pres = self
            .presentation
            .lock()
            .map_err(|e| value_error(format!("Failed to lock presentation: {}", e)))?;
        let para = self.get_paragraph_mut(&mut pres)?;
        para.properties_mut().indent_emu = Some(value);
        Ok(())
    }

    /// Get bullet style: "none", "char:<char>", "autonum:<type>", or None.
    pub fn bullet_style(&self) -> PyResult<Option<String>> {
        let pres = self
            .presentation
            .lock()
            .map_err(|e| value_error(format!("Failed to lock presentation: {}", e)))?;
        let para = self.get_paragraph_ref(&pres)?;
        Ok(para.properties().bullet.style.as_ref().map(|s| match s {
            CoreBulletStyle::None => "none".to_string(),
            CoreBulletStyle::Char(c) => format!("char:{c}"),
            CoreBulletStyle::AutoNum(t) => format!("autonum:{t}"),
        }))
    }

    /// Set bullet to a character bullet (e.g., "\u{2022}" for bullet point).
    pub fn set_bullet_char(&mut self, character: &str) -> PyResult<()> {
        let mut pres = self
            .presentation
            .lock()
            .map_err(|e| value_error(format!("Failed to lock presentation: {}", e)))?;
        let para = self.get_paragraph_mut(&mut pres)?;
        para.properties_mut().bullet.style = Some(CoreBulletStyle::Char(character.to_string()));
        Ok(())
    }

    /// Set bullet to auto-numbered (e.g., "arabicPeriod", "alphaLcParenR").
    pub fn set_bullet_autonum(&mut self, autonum_type: &str) -> PyResult<()> {
        let mut pres = self
            .presentation
            .lock()
            .map_err(|e| value_error(format!("Failed to lock presentation: {}", e)))?;
        let para = self.get_paragraph_mut(&mut pres)?;
        para.properties_mut().bullet.style =
            Some(CoreBulletStyle::AutoNum(autonum_type.to_string()));
        Ok(())
    }

    /// Remove bullet from this paragraph.
    pub fn clear_bullet(&mut self) -> PyResult<()> {
        let mut pres = self
            .presentation
            .lock()
            .map_err(|e| value_error(format!("Failed to lock presentation: {}", e)))?;
        let para = self.get_paragraph_mut(&mut pres)?;
        para.properties_mut().bullet.style = Some(CoreBulletStyle::None);
        Ok(())
    }

    /// Get bullet font name.
    pub fn bullet_font_name(&self) -> PyResult<Option<String>> {
        let pres = self
            .presentation
            .lock()
            .map_err(|e| value_error(format!("Failed to lock presentation: {}", e)))?;
        let para = self.get_paragraph_ref(&pres)?;
        Ok(para.properties().bullet.font_name.clone())
    }

    /// Set bullet font name.
    pub fn set_bullet_font_name(&mut self, name: &str) -> PyResult<()> {
        let mut pres = self
            .presentation
            .lock()
            .map_err(|e| value_error(format!("Failed to lock presentation: {}", e)))?;
        let para = self.get_paragraph_mut(&mut pres)?;
        para.properties_mut().bullet.font_name = Some(name.to_string());
        Ok(())
    }

    /// Get bullet color as sRGB hex string (e.g., "FF0000").
    pub fn bullet_color(&self) -> PyResult<Option<String>> {
        let pres = self
            .presentation
            .lock()
            .map_err(|e| value_error(format!("Failed to lock presentation: {}", e)))?;
        let para = self.get_paragraph_ref(&pres)?;
        Ok(para.properties().bullet.color_srgb.clone())
    }

    /// Set bullet color as sRGB hex string (e.g., "FF0000").
    pub fn set_bullet_color(&mut self, srgb_hex: &str) -> PyResult<()> {
        let mut pres = self
            .presentation
            .lock()
            .map_err(|e| value_error(format!("Failed to lock presentation: {}", e)))?;
        let para = self.get_paragraph_mut(&mut pres)?;
        para.properties_mut().bullet.color_srgb = Some(srgb_hex.to_string());
        Ok(())
    }

    /// Get bullet size as percentage of text size (thousandths of a percent).
    pub fn bullet_size_percent(&self) -> PyResult<Option<u32>> {
        let pres = self
            .presentation
            .lock()
            .map_err(|e| value_error(format!("Failed to lock presentation: {}", e)))?;
        let para = self.get_paragraph_ref(&pres)?;
        Ok(para.properties().bullet.size_percent)
    }

    /// Set bullet size as percentage of text size (thousandths of a percent).
    pub fn set_bullet_size_percent(&mut self, value: u32) -> PyResult<()> {
        let mut pres = self
            .presentation
            .lock()
            .map_err(|e| value_error(format!("Failed to lock presentation: {}", e)))?;
        let para = self.get_paragraph_mut(&mut pres)?;
        para.properties_mut().bullet.size_percent = Some(value);
        Ok(())
    }
}

impl ShapeParagraph {
    fn get_paragraph_ref<'a>(
        &self,
        pres: &'a std::sync::MutexGuard<'_, CorePresentation>,
    ) -> PyResult<&'a CoreShapeParagraph> {
        let slide = pres
            .slides()
            .get(self.slide_index)
            .ok_or_else(|| value_error("Slide no longer exists"))?;
        let shape = slide
            .shapes()
            .get(self.shape_index)
            .ok_or_else(|| value_error("Shape no longer exists"))?;
        shape
            .paragraphs()
            .get(self.paragraph_index)
            .ok_or_else(|| value_error("Paragraph no longer exists"))
    }

    fn get_paragraph_mut<'a>(
        &self,
        pres: &'a mut std::sync::MutexGuard<'_, CorePresentation>,
    ) -> PyResult<&'a mut CoreShapeParagraph> {
        let slide = pres
            .slides_mut()
            .get_mut(self.slide_index)
            .ok_or_else(|| value_error("Slide no longer exists"))?;
        let shape = slide
            .shapes_mut()
            .get_mut(self.shape_index)
            .ok_or_else(|| value_error("Shape no longer exists"))?;
        shape
            .paragraphs_mut()
            .get_mut(self.paragraph_index)
            .ok_or_else(|| value_error("Paragraph no longer exists"))
    }
}

/// Python wrapper for `offidized_pptx::TextRun` with formatting properties.
#[pyclass(module = "offidized._native", name = "TextRun")]
#[derive(Clone)]
pub struct TextRun {
    presentation: Arc<Mutex<CorePresentation>>,
    slide_index: usize,
    shape_index: usize,
    paragraph_index: usize,
    run_index: usize,
}

#[pymethods]
impl TextRun {
    /// Get the text content.
    pub fn text(&self) -> PyResult<String> {
        let pres = self
            .presentation
            .lock()
            .map_err(|e| value_error(format!("Failed to lock presentation: {}", e)))?;
        let run = self.get_run_ref(&pres)?;
        Ok(run.text().to_string())
    }

    /// Set the text content.
    pub fn set_text(&mut self, text: &str) -> PyResult<()> {
        let mut pres = self
            .presentation
            .lock()
            .map_err(|e| value_error(format!("Failed to lock presentation: {}", e)))?;
        let run = self.get_run_mut(&mut pres)?;
        run.set_text(text);
        Ok(())
    }

    /// Get bold state.
    pub fn is_bold(&self) -> PyResult<bool> {
        let pres = self
            .presentation
            .lock()
            .map_err(|e| value_error(format!("Failed to lock presentation: {}", e)))?;
        let run = self.get_run_ref(&pres)?;
        Ok(run.is_bold())
    }

    /// Set bold state.
    pub fn set_bold(&mut self, bold: bool) -> PyResult<()> {
        let mut pres = self
            .presentation
            .lock()
            .map_err(|e| value_error(format!("Failed to lock presentation: {}", e)))?;
        let run = self.get_run_mut(&mut pres)?;
        run.set_bold(bold);
        Ok(())
    }

    /// Get italic state.
    pub fn is_italic(&self) -> PyResult<bool> {
        let pres = self
            .presentation
            .lock()
            .map_err(|e| value_error(format!("Failed to lock presentation: {}", e)))?;
        let run = self.get_run_ref(&pres)?;
        Ok(run.is_italic())
    }

    /// Set italic state.
    pub fn set_italic(&mut self, italic: bool) -> PyResult<()> {
        let mut pres = self
            .presentation
            .lock()
            .map_err(|e| value_error(format!("Failed to lock presentation: {}", e)))?;
        let run = self.get_run_mut(&mut pres)?;
        run.set_italic(italic);
        Ok(())
    }

    /// Get underline style (e.g., "sng", "dbl", "heavy", or None).
    pub fn underline(&self) -> PyResult<Option<String>> {
        let pres = self
            .presentation
            .lock()
            .map_err(|e| value_error(format!("Failed to lock presentation: {}", e)))?;
        let run = self.get_run_ref(&pres)?;
        Ok(run.underline().map(|u| u.to_xml().to_string()))
    }

    /// Set underline style ("sng", "dbl", "heavy", "dotted", "dash", "wavy", etc.).
    pub fn set_underline(&mut self, style: &str) -> PyResult<()> {
        let underline = CoreUnderlineStyle::from_xml(style)
            .ok_or_else(|| value_error(format!("Unknown underline style: {}", style)))?;
        let mut pres = self
            .presentation
            .lock()
            .map_err(|e| value_error(format!("Failed to lock presentation: {}", e)))?;
        let run = self.get_run_mut(&mut pres)?;
        run.set_underline(underline);
        Ok(())
    }

    /// Remove underline.
    pub fn clear_underline(&mut self) -> PyResult<()> {
        let mut pres = self
            .presentation
            .lock()
            .map_err(|e| value_error(format!("Failed to lock presentation: {}", e)))?;
        let run = self.get_run_mut(&mut pres)?;
        run.clear_underline();
        Ok(())
    }

    /// Get strikethrough style ("sngStrike", "dblStrike", or None).
    pub fn strikethrough(&self) -> PyResult<Option<String>> {
        let pres = self
            .presentation
            .lock()
            .map_err(|e| value_error(format!("Failed to lock presentation: {}", e)))?;
        let run = self.get_run_ref(&pres)?;
        Ok(run.strikethrough().map(|s| s.to_xml().to_string()))
    }

    /// Set strikethrough style ("sngStrike" or "dblStrike").
    pub fn set_strikethrough(&mut self, style: &str) -> PyResult<()> {
        let strike = CoreStrikethroughStyle::from_xml(style)
            .ok_or_else(|| value_error(format!("Unknown strikethrough style: {}", style)))?;
        let mut pres = self
            .presentation
            .lock()
            .map_err(|e| value_error(format!("Failed to lock presentation: {}", e)))?;
        let run = self.get_run_mut(&mut pres)?;
        run.set_strikethrough(strike);
        Ok(())
    }

    /// Remove strikethrough.
    pub fn clear_strikethrough(&mut self) -> PyResult<()> {
        let mut pres = self
            .presentation
            .lock()
            .map_err(|e| value_error(format!("Failed to lock presentation: {}", e)))?;
        let run = self.get_run_mut(&mut pres)?;
        run.clear_strikethrough();
        Ok(())
    }

    /// Get font size in hundredths of a point (2400 = 24pt).
    pub fn font_size(&self) -> PyResult<Option<u32>> {
        let pres = self
            .presentation
            .lock()
            .map_err(|e| value_error(format!("Failed to lock presentation: {}", e)))?;
        let run = self.get_run_ref(&pres)?;
        Ok(run.font_size())
    }

    /// Set font size in hundredths of a point (2400 = 24pt).
    pub fn set_font_size(&mut self, hundredths_of_point: u32) -> PyResult<()> {
        let mut pres = self
            .presentation
            .lock()
            .map_err(|e| value_error(format!("Failed to lock presentation: {}", e)))?;
        let run = self.get_run_mut(&mut pres)?;
        run.set_font_size(hundredths_of_point);
        Ok(())
    }

    /// Get font color as sRGB hex (e.g., "FF0000" for red).
    pub fn font_color(&self) -> PyResult<Option<String>> {
        let pres = self
            .presentation
            .lock()
            .map_err(|e| value_error(format!("Failed to lock presentation: {}", e)))?;
        let run = self.get_run_ref(&pres)?;
        Ok(run.font_color().map(String::from))
    }

    /// Set font color as sRGB hex (e.g., "FF0000" for red).
    pub fn set_font_color(&mut self, srgb_hex: &str) -> PyResult<()> {
        let mut pres = self
            .presentation
            .lock()
            .map_err(|e| value_error(format!("Failed to lock presentation: {}", e)))?;
        let run = self.get_run_mut(&mut pres)?;
        run.set_font_color(srgb_hex);
        Ok(())
    }

    /// Get Latin font name (e.g., "Arial", "Calibri").
    pub fn font_name(&self) -> PyResult<Option<String>> {
        let pres = self
            .presentation
            .lock()
            .map_err(|e| value_error(format!("Failed to lock presentation: {}", e)))?;
        let run = self.get_run_ref(&pres)?;
        Ok(run.font_name().map(String::from))
    }

    /// Set Latin font name.
    pub fn set_font_name(&mut self, name: &str) -> PyResult<()> {
        let mut pres = self
            .presentation
            .lock()
            .map_err(|e| value_error(format!("Failed to lock presentation: {}", e)))?;
        let run = self.get_run_mut(&mut pres)?;
        run.set_font_name(name);
        Ok(())
    }

    /// Get language tag (e.g., "en-US").
    pub fn language(&self) -> PyResult<Option<String>> {
        let pres = self
            .presentation
            .lock()
            .map_err(|e| value_error(format!("Failed to lock presentation: {}", e)))?;
        let run = self.get_run_ref(&pres)?;
        Ok(run.language().map(String::from))
    }

    /// Set language tag.
    pub fn set_language(&mut self, lang: &str) -> PyResult<()> {
        let mut pres = self
            .presentation
            .lock()
            .map_err(|e| value_error(format!("Failed to lock presentation: {}", e)))?;
        let run = self.get_run_mut(&mut pres)?;
        run.set_language(lang);
        Ok(())
    }

    /// Get hyperlink URL.
    pub fn hyperlink_url(&self) -> PyResult<Option<String>> {
        let pres = self
            .presentation
            .lock()
            .map_err(|e| value_error(format!("Failed to lock presentation: {}", e)))?;
        let run = self.get_run_ref(&pres)?;
        Ok(run.hyperlink_url().map(String::from))
    }

    /// Set hyperlink URL.
    pub fn set_hyperlink_url(&mut self, url: &str) -> PyResult<()> {
        let mut pres = self
            .presentation
            .lock()
            .map_err(|e| value_error(format!("Failed to lock presentation: {}", e)))?;
        let run = self.get_run_mut(&mut pres)?;
        run.set_hyperlink_url(url);
        Ok(())
    }

    /// Get hyperlink tooltip.
    pub fn hyperlink_tooltip(&self) -> PyResult<Option<String>> {
        let pres = self
            .presentation
            .lock()
            .map_err(|e| value_error(format!("Failed to lock presentation: {}", e)))?;
        let run = self.get_run_ref(&pres)?;
        Ok(run.hyperlink_tooltip().map(String::from))
    }

    /// Set hyperlink tooltip.
    pub fn set_hyperlink_tooltip(&mut self, tooltip: &str) -> PyResult<()> {
        let mut pres = self
            .presentation
            .lock()
            .map_err(|e| value_error(format!("Failed to lock presentation: {}", e)))?;
        let run = self.get_run_mut(&mut pres)?;
        run.set_hyperlink_tooltip(tooltip);
        Ok(())
    }

    /// Get character spacing in hundredths of a point.
    pub fn character_spacing(&self) -> PyResult<Option<i32>> {
        let pres = self
            .presentation
            .lock()
            .map_err(|e| value_error(format!("Failed to lock presentation: {}", e)))?;
        let run = self.get_run_ref(&pres)?;
        Ok(run.character_spacing())
    }

    /// Set character spacing in hundredths of a point. Negative = condensed.
    pub fn set_character_spacing(&mut self, hundredths_of_point: i32) -> PyResult<()> {
        let mut pres = self
            .presentation
            .lock()
            .map_err(|e| value_error(format!("Failed to lock presentation: {}", e)))?;
        let run = self.get_run_mut(&mut pres)?;
        run.set_character_spacing(hundredths_of_point);
        Ok(())
    }

    /// Get baseline offset (positive = superscript, negative = subscript).
    pub fn baseline(&self) -> PyResult<Option<i32>> {
        let pres = self
            .presentation
            .lock()
            .map_err(|e| value_error(format!("Failed to lock presentation: {}", e)))?;
        let run = self.get_run_ref(&pres)?;
        Ok(run.baseline())
    }

    /// Set baseline offset (e.g., 30000 for superscript, -25000 for subscript).
    pub fn set_baseline(&mut self, baseline: i32) -> PyResult<()> {
        let mut pres = self
            .presentation
            .lock()
            .map_err(|e| value_error(format!("Failed to lock presentation: {}", e)))?;
        let run = self.get_run_mut(&mut pres)?;
        run.set_baseline(baseline);
        Ok(())
    }

    /// Get kerning threshold in hundredths of a point.
    pub fn kerning(&self) -> PyResult<Option<i32>> {
        let pres = self
            .presentation
            .lock()
            .map_err(|e| value_error(format!("Failed to lock presentation: {}", e)))?;
        let run = self.get_run_ref(&pres)?;
        Ok(run.kerning())
    }

    /// Set kerning threshold in hundredths of a point (0 to disable).
    pub fn set_kerning(&mut self, hundredths_of_point: i32) -> PyResult<()> {
        let mut pres = self
            .presentation
            .lock()
            .map_err(|e| value_error(format!("Failed to lock presentation: {}", e)))?;
        let run = self.get_run_mut(&mut pres)?;
        run.set_kerning(hundredths_of_point);
        Ok(())
    }
}

impl TextRun {
    fn get_run_ref<'a>(
        &self,
        pres: &'a std::sync::MutexGuard<'_, CorePresentation>,
    ) -> PyResult<&'a CoreTextRun> {
        let slide = pres
            .slides()
            .get(self.slide_index)
            .ok_or_else(|| value_error("Slide no longer exists"))?;
        let shape = slide
            .shapes()
            .get(self.shape_index)
            .ok_or_else(|| value_error("Shape no longer exists"))?;
        let para = shape
            .paragraphs()
            .get(self.paragraph_index)
            .ok_or_else(|| value_error("Paragraph no longer exists"))?;
        para.runs()
            .get(self.run_index)
            .ok_or_else(|| value_error("Run no longer exists"))
    }

    fn get_run_mut<'a>(
        &self,
        pres: &'a mut std::sync::MutexGuard<'_, CorePresentation>,
    ) -> PyResult<&'a mut CoreTextRun> {
        let slide = pres
            .slides_mut()
            .get_mut(self.slide_index)
            .ok_or_else(|| value_error("Slide no longer exists"))?;
        let shape = slide
            .shapes_mut()
            .get_mut(self.shape_index)
            .ok_or_else(|| value_error("Shape no longer exists"))?;
        let para = shape
            .paragraphs_mut()
            .get_mut(self.paragraph_index)
            .ok_or_else(|| value_error("Paragraph no longer exists"))?;
        para.runs_mut()
            .get_mut(self.run_index)
            .ok_or_else(|| value_error("Run no longer exists"))
    }
}

// =============================================================================
// PPTX Bindings - Slide Show Settings, Custom Shows, Transitions
// =============================================================================

/// Python wrapper for `offidized_pptx::SlideShowSettings`.
#[pyclass(module = "offidized._native", name = "SlideShowSettings")]
#[derive(Clone)]
pub struct PySlideShowSettings {
    inner: CoreSlideShowSettings,
}

#[pymethods]
impl PySlideShowSettings {
    /// Create new slide show settings.
    #[new]
    fn new() -> Self {
        Self {
            inner: CoreSlideShowSettings::new(),
        }
    }
}

/// Python wrapper for `offidized_pptx::CustomShow`.
#[pyclass(module = "offidized._native", name = "CustomShow")]
#[derive(Clone)]
pub struct PyCustomShow {
    inner: CoreCustomShow,
}

#[pymethods]
impl PyCustomShow {
    /// Create new custom show with name.
    #[new]
    #[pyo3(signature = (name, id=None))]
    fn new(name: &str, id: Option<u32>) -> Self {
        let show_id = id.unwrap_or_else(|| {
            use std::collections::hash_map::DefaultHasher;
            use std::hash::{Hash, Hasher};
            let mut hasher = DefaultHasher::new();
            name.hash(&mut hasher);
            (hasher.finish() & 0xFFFFFFFF) as u32
        });
        Self {
            inner: CoreCustomShow::new(name, show_id),
        }
    }

    /// Get the custom show name.
    fn name(&self) -> String {
        self.inner.name().to_string()
    }

    /// Set the custom show name.
    fn set_name(&mut self, name: &str) {
        self.inner.set_name(name);
    }
}

/// Python wrapper for `offidized_pptx::SlideTransition`.
#[pyclass(module = "offidized._native", name = "SlideTransition")]
#[derive(Clone)]
pub struct PySlideTransition {
    inner: CoreSlideTransition,
}

#[pymethods]
impl PySlideTransition {
    /// Create a new slide transition with the given kind.
    #[new]
    fn new(kind: &str) -> PyResult<Self> {
        let transition_kind = match kind.to_lowercase().as_str() {
            "unspecified" => CoreSlideTransitionKind::Unspecified,
            "cut" => CoreSlideTransitionKind::Cut,
            "fade" => CoreSlideTransitionKind::Fade,
            "push" => CoreSlideTransitionKind::Push,
            "wipe" => CoreSlideTransitionKind::Wipe,
            _ => CoreSlideTransitionKind::Other(kind.to_string()),
        };
        Ok(Self {
            inner: CoreSlideTransition::new(transition_kind),
        })
    }
}

// =============================================================================
// Module Registration
// =============================================================================

pub(crate) fn register(module: &Bound<'_, PyModule>) -> PyResult<()> {
    // Core
    module.add_class::<Presentation>()?;
    module.add_class::<PresentationProperties>()?;
    module.add_class::<Slide>()?;
    // Content types
    module.add_class::<Table>()?;
    module.add_class::<Chart>()?;
    module.add_class::<Image>()?;
    // Text formatting
    module.add_class::<PyShape>()?;
    module.add_class::<ShapeParagraph>()?;
    module.add_class::<TextRun>()?;
    // Settings & transitions
    module.add_class::<PySlideShowSettings>()?;
    module.add_class::<PyCustomShow>()?;
    module.add_class::<PySlideTransition>()?;
    Ok(())
}
