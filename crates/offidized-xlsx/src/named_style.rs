/// Named cell styles (ECMA-376 18.8.7 cellStyle, 18.8.10 cellStyleXfs).
///
/// A named style is a user-visible style entry (e.g. "Normal", "Heading 1") that maps
/// to a `cellStyleXf` record via `xf_id`. The optional `builtin_id` links to one of
/// the 63 predefined styles defined by the OOXML spec.
/// A named cell style visible to the user in the Styles gallery.
///
/// Each `NamedStyle` maps to a `CellStyleXf` entry via `xf_id`.
/// If the style corresponds to a built-in OOXML style, `builtin_id` is set.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct NamedStyle {
    name: String,
    xf_id: u32,
    builtin_id: Option<u32>,
}

impl NamedStyle {
    /// Creates a new named style with the given display name and xf_id reference.
    pub fn new(name: impl Into<String>, xf_id: u32) -> Self {
        Self {
            name: name.into(),
            xf_id,
            builtin_id: None,
        }
    }

    /// Returns the display name of this style.
    pub fn name(&self) -> &str {
        &self.name
    }

    /// Sets the display name of this style.
    pub fn set_name(&mut self, name: impl Into<String>) -> &mut Self {
        self.name = name.into();
        self
    }

    /// Returns the `xf_id` index into the `cellStyleXfs` table.
    pub fn xf_id(&self) -> u32 {
        self.xf_id
    }

    /// Sets the `xf_id` index.
    pub fn set_xf_id(&mut self, xf_id: u32) -> &mut Self {
        self.xf_id = xf_id;
        self
    }

    /// Returns the built-in style ID, if this is a predefined style.
    pub fn builtin_id(&self) -> Option<u32> {
        self.builtin_id
    }

    /// Sets the built-in style ID.
    pub fn set_builtin_id(&mut self, id: u32) -> &mut Self {
        self.builtin_id = Some(id);
        self
    }

    /// Clears the built-in style ID.
    pub fn clear_builtin_id(&mut self) -> &mut Self {
        self.builtin_id = None;
        self
    }
}

/// Formatting attributes for a cell style XF record (`cellStyleXfs`).
///
/// Each field is an index into the corresponding style table component
/// (number formats, fonts, fills, borders). `None` means "not specified"
/// and inherits from the default.
#[derive(Debug, Clone, PartialEq, Eq, Default)]
pub struct CellStyleXf {
    num_fmt_id: Option<u32>,
    font_id: Option<u32>,
    fill_id: Option<u32>,
    border_id: Option<u32>,
}

impl CellStyleXf {
    /// Creates a new empty `CellStyleXf`.
    pub fn new() -> Self {
        Self::default()
    }

    /// Returns the number format ID.
    pub fn num_fmt_id(&self) -> Option<u32> {
        self.num_fmt_id
    }

    /// Sets the number format ID.
    pub fn set_num_fmt_id(&mut self, id: u32) -> &mut Self {
        self.num_fmt_id = Some(id);
        self
    }

    /// Clears the number format ID.
    pub fn clear_num_fmt_id(&mut self) -> &mut Self {
        self.num_fmt_id = None;
        self
    }

    /// Returns the font ID.
    pub fn font_id(&self) -> Option<u32> {
        self.font_id
    }

    /// Sets the font ID.
    pub fn set_font_id(&mut self, id: u32) -> &mut Self {
        self.font_id = Some(id);
        self
    }

    /// Clears the font ID.
    pub fn clear_font_id(&mut self) -> &mut Self {
        self.font_id = None;
        self
    }

    /// Returns the fill ID.
    pub fn fill_id(&self) -> Option<u32> {
        self.fill_id
    }

    /// Sets the fill ID.
    pub fn set_fill_id(&mut self, id: u32) -> &mut Self {
        self.fill_id = Some(id);
        self
    }

    /// Clears the fill ID.
    pub fn clear_fill_id(&mut self) -> &mut Self {
        self.fill_id = None;
        self
    }

    /// Returns the border ID.
    pub fn border_id(&self) -> Option<u32> {
        self.border_id
    }

    /// Sets the border ID.
    pub fn set_border_id(&mut self, id: u32) -> &mut Self {
        self.border_id = Some(id);
        self
    }

    /// Clears the border ID.
    pub fn clear_border_id(&mut self) -> &mut Self {
        self.border_id = None;
        self
    }
}

/// Returns the human-readable name for a built-in cell style ID (ECMA-376 18.8.7).
///
/// The OOXML spec defines 63 built-in style IDs. This function maps each ID
/// to its canonical English name as it appears in Excel's Styles gallery.
pub fn builtin_style_name(builtin_id: u32) -> Option<&'static str> {
    match builtin_id {
        0 => Some("Normal"),
        1 => Some("RowLevel_1"),
        2 => Some("ColLevel_1"),
        3 => Some("Comma"),
        4 => Some("Currency"),
        5 => Some("Percent"),
        6 => Some("Comma [0]"),
        7 => Some("Currency [0]"),
        8 => Some("Hyperlink"),
        9 => Some("Followed Hyperlink"),
        10 => Some("Note"),
        11 => Some("Warning Text"),
        15 => Some("Title"),
        16 => Some("Heading 1"),
        17 => Some("Heading 2"),
        18 => Some("Heading 3"),
        19 => Some("Heading 4"),
        20 => Some("Input"),
        21 => Some("Output"),
        22 => Some("Calculation"),
        23 => Some("Check Cell"),
        24 => Some("Linked Cell"),
        25 => Some("Total"),
        26 => Some("Good"),
        27 => Some("Bad"),
        28 => Some("Neutral"),
        29 => Some("Accent1"),
        30 => Some("20% - Accent1"),
        31 => Some("40% - Accent1"),
        32 => Some("60% - Accent1"),
        33 => Some("Accent2"),
        34 => Some("20% - Accent2"),
        35 => Some("40% - Accent2"),
        36 => Some("60% - Accent2"),
        37 => Some("Accent3"),
        38 => Some("20% - Accent3"),
        39 => Some("40% - Accent3"),
        40 => Some("60% - Accent3"),
        41 => Some("Accent4"),
        42 => Some("20% - Accent4"),
        43 => Some("40% - Accent4"),
        44 => Some("60% - Accent4"),
        45 => Some("Accent5"),
        46 => Some("20% - Accent5"),
        47 => Some("40% - Accent5"),
        48 => Some("60% - Accent5"),
        49 => Some("Accent6"),
        50 => Some("20% - Accent6"),
        51 => Some("40% - Accent6"),
        52 => Some("60% - Accent6"),
        53 => Some("Explanatory Text"),
        54 => Some("RowLevel_2"),
        55 => Some("RowLevel_3"),
        56 => Some("RowLevel_4"),
        57 => Some("RowLevel_5"),
        58 => Some("RowLevel_6"),
        59 => Some("RowLevel_7"),
        60 => Some("ColLevel_2"),
        61 => Some("ColLevel_3"),
        62 => Some("ColLevel_4"),
        63 => Some("ColLevel_5"),
        64 => Some("ColLevel_6"),
        65 => Some("ColLevel_7"),
        _ => None,
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn named_style_construction_and_accessors() {
        let mut style = NamedStyle::new("Normal", 0);
        assert_eq!(style.name(), "Normal");
        assert_eq!(style.xf_id(), 0);
        assert_eq!(style.builtin_id(), None);

        style.set_builtin_id(0);
        assert_eq!(style.builtin_id(), Some(0));

        style.set_name("Custom Style");
        assert_eq!(style.name(), "Custom Style");

        style.set_xf_id(5);
        assert_eq!(style.xf_id(), 5);

        style.clear_builtin_id();
        assert_eq!(style.builtin_id(), None);
    }

    #[test]
    fn cell_style_xf_fields() {
        let mut xf = CellStyleXf::new();
        assert_eq!(xf.num_fmt_id(), None);
        assert_eq!(xf.font_id(), None);
        assert_eq!(xf.fill_id(), None);
        assert_eq!(xf.border_id(), None);

        xf.set_num_fmt_id(164)
            .set_font_id(1)
            .set_fill_id(2)
            .set_border_id(3);
        assert_eq!(xf.num_fmt_id(), Some(164));
        assert_eq!(xf.font_id(), Some(1));
        assert_eq!(xf.fill_id(), Some(2));
        assert_eq!(xf.border_id(), Some(3));

        xf.clear_num_fmt_id()
            .clear_font_id()
            .clear_fill_id()
            .clear_border_id();
        assert_eq!(xf.num_fmt_id(), None);
        assert_eq!(xf.font_id(), None);
        assert_eq!(xf.fill_id(), None);
        assert_eq!(xf.border_id(), None);
    }

    #[test]
    fn builtin_style_name_lookups() {
        assert_eq!(builtin_style_name(0), Some("Normal"));
        assert_eq!(builtin_style_name(15), Some("Title"));
        assert_eq!(builtin_style_name(16), Some("Heading 1"));
        assert_eq!(builtin_style_name(17), Some("Heading 2"));
        assert_eq!(builtin_style_name(18), Some("Heading 3"));
        assert_eq!(builtin_style_name(19), Some("Heading 4"));
        assert_eq!(builtin_style_name(26), Some("Good"));
        assert_eq!(builtin_style_name(27), Some("Bad"));
        assert_eq!(builtin_style_name(28), Some("Neutral"));
        assert_eq!(builtin_style_name(29), Some("Accent1"));
        assert_eq!(builtin_style_name(30), Some("20% - Accent1"));
        assert_eq!(builtin_style_name(31), Some("40% - Accent1"));
        assert_eq!(builtin_style_name(32), Some("60% - Accent1"));
        assert_eq!(builtin_style_name(49), Some("Accent6"));
        assert_eq!(builtin_style_name(50), Some("20% - Accent6"));
        assert_eq!(builtin_style_name(51), Some("40% - Accent6"));
        assert_eq!(builtin_style_name(52), Some("60% - Accent6"));
        assert_eq!(builtin_style_name(53), Some("Explanatory Text"));
        assert_eq!(builtin_style_name(100), None);
        assert_eq!(builtin_style_name(u32::MAX), None);
    }

    #[test]
    fn builtin_ids_12_through_14_are_not_defined() {
        // IDs 12, 13, 14 are not defined in the spec
        assert_eq!(builtin_style_name(12), None);
        assert_eq!(builtin_style_name(13), None);
        assert_eq!(builtin_style_name(14), None);
    }
}
