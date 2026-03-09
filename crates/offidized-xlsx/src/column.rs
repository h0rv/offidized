/// Minimal column metadata.
#[derive(Debug, Clone, PartialEq, Default)]
pub struct Column {
    index: u32,
    width: Option<f64>,
    hidden: bool,
    outline_level: u8,
    collapsed: bool,
    style_index: Option<u32>,
    best_fit: bool,
    custom_width: bool,
}

impl Column {
    pub fn new(index: u32) -> Self {
        Self {
            index,
            width: None,
            hidden: false,
            outline_level: 0,
            collapsed: false,
            style_index: None,
            best_fit: false,
            custom_width: false,
        }
    }

    pub fn index(&self) -> u32 {
        self.index
    }

    pub fn width(&self) -> Option<f64> {
        self.width
    }

    pub fn set_width(&mut self, width: f64) -> &mut Self {
        if width.is_finite() && width >= 0.0 {
            self.width = Some(width);
        }
        self
    }

    pub fn clear_width(&mut self) -> &mut Self {
        self.width = None;
        self
    }

    /// Returns whether the column is hidden.
    pub fn is_hidden(&self) -> bool {
        self.hidden
    }

    /// Sets the column hidden state.
    pub fn set_hidden(&mut self, hidden: bool) -> &mut Self {
        self.hidden = hidden;
        self
    }

    /// Returns the outline (grouping) level for this column (0-7).
    pub fn outline_level(&self) -> u8 {
        self.outline_level
    }

    /// Sets the outline (grouping) level for this column (0-7).
    pub fn set_outline_level(&mut self, level: u8) -> &mut Self {
        self.outline_level = level.min(7);
        self
    }

    /// Returns whether the column group is collapsed.
    pub fn is_collapsed(&self) -> bool {
        self.collapsed
    }

    /// Sets whether the column group is collapsed.
    pub fn set_collapsed(&mut self, collapsed: bool) -> &mut Self {
        self.collapsed = collapsed;
        self
    }

    /// Returns the style index applied to this column, if any.
    pub fn style_index(&self) -> Option<u32> {
        self.style_index
    }

    /// Sets the style index for this column.
    pub fn set_style_index(&mut self, index: u32) -> &mut Self {
        self.style_index = Some(index);
        self
    }

    /// Clears the style index.
    pub fn clear_style_index(&mut self) -> &mut Self {
        self.style_index = None;
        self
    }

    /// Returns whether the column width is a best-fit auto width.
    pub fn is_best_fit(&self) -> bool {
        self.best_fit
    }

    /// Sets whether the column width is a best-fit auto width.
    pub fn set_best_fit(&mut self, best_fit: bool) -> &mut Self {
        self.best_fit = best_fit;
        self
    }

    /// Returns whether the column has a custom (non-default) width.
    pub fn custom_width(&self) -> bool {
        self.custom_width
    }

    /// Sets whether the column has a custom (non-default) width.
    pub fn set_custom_width(&mut self, custom_width: bool) -> &mut Self {
        self.custom_width = custom_width;
        self
    }
}
