use offidized_opc::RawXmlNode;

/// Minimal row metadata.
#[derive(Debug, Clone, PartialEq, Default)]
pub struct Row {
    index: u32,
    height: Option<f64>,
    hidden: bool,
    custom_height: bool,
    outline_level: u8,
    collapsed: bool,
    unknown_attrs: Vec<(String, String)>,
    unknown_children: Vec<RawXmlNode>,
}

impl Row {
    pub fn new(index: u32) -> Self {
        Self {
            index,
            height: None,
            hidden: false,
            custom_height: false,
            outline_level: 0,
            collapsed: false,
            unknown_attrs: Vec::new(),
            unknown_children: Vec::new(),
        }
    }

    pub fn index(&self) -> u32 {
        self.index
    }

    pub fn height(&self) -> Option<f64> {
        self.height
    }

    pub fn set_height(&mut self, height: f64) -> &mut Self {
        if height.is_finite() && height >= 0.0 {
            self.height = Some(height);
        }
        self
    }

    pub fn clear_height(&mut self) -> &mut Self {
        self.height = None;
        self
    }

    /// Returns whether the row has a custom height.
    pub fn custom_height(&self) -> bool {
        self.custom_height
    }

    /// Sets whether the row has a custom height.
    pub fn set_custom_height(&mut self, value: bool) -> &mut Self {
        self.custom_height = value;
        self
    }

    /// Returns whether the row is hidden.
    pub fn is_hidden(&self) -> bool {
        self.hidden
    }

    /// Sets the row hidden state.
    pub fn set_hidden(&mut self, hidden: bool) -> &mut Self {
        self.hidden = hidden;
        self
    }

    /// Returns the outline (grouping) level for this row (0-7).
    pub fn outline_level(&self) -> u8 {
        self.outline_level
    }

    /// Sets the outline (grouping) level for this row (0-7).
    pub fn set_outline_level(&mut self, level: u8) -> &mut Self {
        self.outline_level = level.min(7);
        self
    }

    /// Returns whether the row group is collapsed.
    pub fn is_collapsed(&self) -> bool {
        self.collapsed
    }

    /// Sets whether the row group is collapsed.
    pub fn set_collapsed(&mut self, collapsed: bool) -> &mut Self {
        self.collapsed = collapsed;
        self
    }

    pub(crate) fn unknown_attrs(&self) -> &[(String, String)] {
        self.unknown_attrs.as_slice()
    }

    pub(crate) fn set_unknown_attrs(&mut self, attrs: Vec<(String, String)>) {
        self.unknown_attrs = attrs;
    }

    pub(crate) fn unknown_children(&self) -> &[RawXmlNode] {
        self.unknown_children.as_slice()
    }

    pub(crate) fn push_unknown_child(&mut self, node: RawXmlNode) {
        self.unknown_children.push(node);
    }
}
