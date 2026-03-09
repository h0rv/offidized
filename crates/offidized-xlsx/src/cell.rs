use crate::error::{Result, XlsxError};
use crate::style::Font;
use offidized_opc::RawXmlNode;

/// Supported cell value types for the minimal scaffold.
#[derive(Debug, Clone, PartialEq)]
pub enum CellValue {
    Blank,
    String(String),
    Number(f64),
    Bool(bool),
    Date(String),
    Error(String),
    /// A numeric date/time value stored as an Excel serial number.
    ///
    /// Excel stores dates as floating-point numbers where the integer part represents
    /// the number of days since the epoch (1899-12-30 in the 1900 date system) and
    /// the fractional part represents the time of day.  Built-in number format IDs
    /// 14-22 are date/time formats.
    DateTime(f64),
    /// Rich text: multiple formatted runs within a single cell.
    RichText(Vec<RichTextRun>),
}

impl CellValue {
    pub fn date(value: impl Into<String>) -> Self {
        Self::Date(value.into())
    }

    pub fn error(value: impl Into<String>) -> Self {
        Self::Error(value.into())
    }

    /// Creates a `DateTime` value from an Excel serial date number.
    pub fn date_time(serial: f64) -> Self {
        Self::DateTime(serial)
    }

    /// Creates a `RichText` value from a vector of runs.
    pub fn rich_text(runs: Vec<RichTextRun>) -> Self {
        Self::RichText(runs)
    }

    /// Returns true if this value represents a date/time serial number.
    pub fn is_date_time(&self) -> bool {
        matches!(self, Self::DateTime(_))
    }

    /// Returns the DateTime serial number if this is a DateTime variant.
    pub fn as_date_time(&self) -> Option<f64> {
        match self {
            Self::DateTime(serial) => Some(*serial),
            _ => None,
        }
    }
}

impl From<&str> for CellValue {
    fn from(value: &str) -> Self {
        Self::String(value.to_string())
    }
}

impl From<String> for CellValue {
    fn from(value: String) -> Self {
        Self::String(value)
    }
}

impl From<bool> for CellValue {
    fn from(value: bool) -> Self {
        Self::Bool(value)
    }
}

impl From<f32> for CellValue {
    fn from(value: f32) -> Self {
        Self::Number(f64::from(value))
    }
}

impl From<f64> for CellValue {
    fn from(value: f64) -> Self {
        Self::Number(value)
    }
}

macro_rules! impl_number_to_cell_value {
    ($($ty:ty),* $(,)?) => {
        $(
            impl From<$ty> for CellValue {
                fn from(value: $ty) -> Self {
                    Self::Number(value as f64)
                }
            }
        )*
    };
}

impl_number_to_cell_value!(i8, i16, i32, i64, isize, u8, u16, u32, u64, usize,);

/// A single formatted run within a rich text cell value.
///
/// Rich text cells contain multiple runs, each with its own text and optional
/// formatting (bold, italic, font name, size, color). This maps to the `<r>`
/// elements inside `<si>` (shared strings) or `<is>` (inline strings).
#[derive(Debug, Clone, PartialEq)]
pub struct RichTextRun {
    text: String,
    bold: Option<bool>,
    italic: Option<bool>,
    font_name: Option<String>,
    font_size: Option<String>,
    color: Option<String>,
    /// Unknown `<rPr>` children preserved for roundtrip fidelity.
    unknown_rpr_children: Vec<RawXmlNode>,
}

impl RichTextRun {
    /// Creates a new rich text run with the given text and no formatting.
    pub fn new(text: impl Into<String>) -> Self {
        Self {
            text: text.into(),
            bold: None,
            italic: None,
            font_name: None,
            font_size: None,
            color: None,
            unknown_rpr_children: Vec::new(),
        }
    }

    /// Returns the text content of this run.
    pub fn text(&self) -> &str {
        self.text.as_str()
    }

    /// Sets the text content.
    pub fn set_text(&mut self, text: impl Into<String>) -> &mut Self {
        self.text = text.into();
        self
    }

    /// Returns whether this run is bold.
    pub fn bold(&self) -> Option<bool> {
        self.bold
    }

    /// Sets bold formatting.
    pub fn set_bold(&mut self, bold: bool) -> &mut Self {
        self.bold = Some(bold);
        self
    }

    /// Clears bold formatting.
    pub fn clear_bold(&mut self) -> &mut Self {
        self.bold = None;
        self
    }

    /// Returns whether this run is italic.
    pub fn italic(&self) -> Option<bool> {
        self.italic
    }

    /// Sets italic formatting.
    pub fn set_italic(&mut self, italic: bool) -> &mut Self {
        self.italic = Some(italic);
        self
    }

    /// Clears italic formatting.
    pub fn clear_italic(&mut self) -> &mut Self {
        self.italic = None;
        self
    }

    /// Returns the font name for this run.
    pub fn font_name(&self) -> Option<&str> {
        self.font_name.as_deref()
    }

    /// Sets the font name.
    pub fn set_font_name(&mut self, name: impl Into<String>) -> &mut Self {
        self.font_name = Some(name.into());
        self
    }

    /// Clears the font name.
    pub fn clear_font_name(&mut self) -> &mut Self {
        self.font_name = None;
        self
    }

    /// Returns the font size for this run.
    pub fn font_size(&self) -> Option<&str> {
        self.font_size.as_deref()
    }

    /// Sets the font size.
    pub fn set_font_size(&mut self, size: impl Into<String>) -> &mut Self {
        self.font_size = Some(size.into());
        self
    }

    /// Clears the font size.
    pub fn clear_font_size(&mut self) -> &mut Self {
        self.font_size = None;
        self
    }

    /// Returns the color for this run (ARGB hex string).
    pub fn color(&self) -> Option<&str> {
        self.color.as_deref()
    }

    /// Sets the color (ARGB hex string).
    pub fn set_color(&mut self, color: impl Into<String>) -> &mut Self {
        self.color = Some(color.into());
        self
    }

    /// Clears the color.
    pub fn clear_color(&mut self) -> &mut Self {
        self.color = None;
        self
    }

    /// Returns true if any formatting properties are set on this run.
    pub fn has_formatting(&self) -> bool {
        self.bold.is_some()
            || self.italic.is_some()
            || self.font_name.is_some()
            || self.font_size.is_some()
            || self.color.is_some()
            || !self.unknown_rpr_children.is_empty()
    }

    /// Returns the unknown `<rPr>` children preserved for roundtrip fidelity.
    pub fn unknown_rpr_children(&self) -> &[RawXmlNode] {
        &self.unknown_rpr_children
    }

    /// Sets the unknown `<rPr>` children for roundtrip preservation.
    pub fn set_unknown_rpr_children(&mut self, children: Vec<RawXmlNode>) -> &mut Self {
        self.unknown_rpr_children = children;
        self
    }

    /// Converts a `Font` into the formatting properties of a new run with the given text.
    pub fn from_font(text: impl Into<String>, font: &Font) -> Self {
        let mut run = Self::new(text);
        if let Some(b) = font.bold() {
            run.set_bold(b);
        }
        if let Some(i) = font.italic() {
            run.set_italic(i);
        }
        if let Some(name) = font.name() {
            run.set_font_name(name);
        }
        if let Some(size) = font.size() {
            run.set_font_size(size);
        }
        if let Some(color) = font.color() {
            run.set_color(color);
        }
        run
    }
}

/// A comment or note attached to a cell.
///
/// This struct represents cell-level comment data. Comments are stored on the
/// `Cell` itself and also maintained in the worksheet-level comments list for
/// serialization into the `<comments>` XML part.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct CellComment {
    author: String,
    text: String,
}

impl CellComment {
    /// Creates a new cell comment.
    pub fn new(author: impl Into<String>, text: impl Into<String>) -> Self {
        Self {
            author: author.into(),
            text: text.into(),
        }
    }

    /// Returns the comment author.
    pub fn author(&self) -> &str {
        self.author.as_str()
    }

    /// Sets the comment author.
    pub fn set_author(&mut self, author: impl Into<String>) -> &mut Self {
        self.author = author.into();
        self
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
}

/// A single worksheet cell.
#[derive(Debug, Clone, Default, PartialEq)]
pub struct Cell {
    value: Option<CellValue>,
    formula: Option<String>,
    /// The cached result that Excel computed and stored alongside a formula.
    ///
    /// When a cell has both a `<f>` (formula) and a `<v>` (value) element, the
    /// `<v>` contains the last-computed result of the formula. This field stores
    /// that cached result so it can be round-tripped without loss.
    cached_value: Option<CellValue>,
    style_id: Option<u32>,
    comment: Option<CellComment>,
    /// Whether this cell contains an array (CSE) formula.
    is_array_formula: bool,
    /// The range that the array formula spans (e.g. "A1:C3").
    array_range: Option<String>,
    /// The shared formula index (`si` attribute on `<f t="shared">`).
    ///
    /// When multiple cells share the same formula pattern, Excel writes a
    /// "master" cell with `<f t="shared" ref="A1:A10" si="0">formula</f>` and
    /// dependent cells with `<f t="shared" si="0"/>`.  This field stores the
    /// `si` value for both cases.
    shared_formula_index: Option<u32>,
    unknown_attrs: Vec<(String, String)>,
    unknown_children: Vec<RawXmlNode>,
}

impl Cell {
    pub fn new() -> Self {
        Self::default()
    }

    pub fn value(&self) -> Option<&CellValue> {
        self.value.as_ref()
    }

    pub fn set_value(&mut self, value: impl Into<CellValue>) -> &mut Self {
        self.value = Some(value.into());
        self
    }

    pub fn clear_value(&mut self) -> &mut Self {
        self.value = None;
        self
    }

    pub fn formula(&self) -> Option<&str> {
        self.formula.as_deref()
    }

    pub fn set_formula(&mut self, formula: impl Into<String>) -> &mut Self {
        let formula_str = formula.into();
        // Strip leading '=' if present for better UX (users often include it)
        let formula_normalized = formula_str.strip_prefix('=').unwrap_or(&formula_str);
        self.formula = Some(formula_normalized.to_string());
        self
    }

    pub fn clear_formula(&mut self) -> &mut Self {
        self.formula = None;
        self
    }

    /// Returns the cached formula result, if present.
    ///
    /// When a cell has a formula, Excel stores the last-computed result in the
    /// `<v>` element alongside the `<f>` element. This method returns that
    /// cached result.
    pub fn cached_value(&self) -> Option<&CellValue> {
        self.cached_value.as_ref()
    }

    /// Sets the cached formula result.
    ///
    /// This is the value that Excel will display without recalculating the
    /// formula.
    pub fn set_cached_value(&mut self, value: CellValue) -> &mut Self {
        self.cached_value = Some(value);
        self
    }

    /// Clears the cached formula result.
    pub fn clear_cached_value(&mut self) -> &mut Self {
        self.cached_value = None;
        self
    }

    pub fn style_id(&self) -> Option<u32> {
        self.style_id
    }

    pub fn set_style_id(&mut self, style_id: u32) -> &mut Self {
        self.style_id = Some(style_id);
        self
    }

    pub fn clear_style_id(&mut self) -> &mut Self {
        self.style_id = None;
        self
    }

    // ---- Comment (Feature 2) ----

    /// Returns the cell comment, if set.
    pub fn comment(&self) -> Option<&CellComment> {
        self.comment.as_ref()
    }

    /// Sets a comment on this cell.
    pub fn set_comment(&mut self, comment: CellComment) -> &mut Self {
        self.comment = Some(comment);
        self
    }

    /// Clears the cell comment.
    pub fn clear_comment(&mut self) -> &mut Self {
        self.comment = None;
        self
    }

    // ---- Array/CSE formulas (Feature 11) ----

    /// Returns whether this cell contains an array (CSE) formula.
    pub fn is_array_formula(&self) -> bool {
        self.is_array_formula
    }

    /// Sets whether this cell contains an array formula.
    pub fn set_array_formula(&mut self, is_array: bool) -> &mut Self {
        self.is_array_formula = is_array;
        self
    }

    /// Returns the range that the array formula spans, if set.
    pub fn array_range(&self) -> Option<&str> {
        self.array_range.as_deref()
    }

    /// Sets the array formula range (e.g. "A1:C3").
    pub fn set_array_range(&mut self, range: impl Into<String>) -> &mut Self {
        self.array_range = Some(range.into());
        self.is_array_formula = true;
        self
    }

    /// Clears the array formula range and flag.
    pub fn clear_array_range(&mut self) -> &mut Self {
        self.array_range = None;
        self.is_array_formula = false;
        self
    }

    // ---- Shared formulas (Feature 12) ----

    /// Returns the shared formula index (`si` attribute), if set.
    ///
    /// Both master cells and dependent cells in a shared formula group have
    /// the same `si` value. The master cell also has formula text; dependent
    /// cells typically have only the `si` index with no formula text.
    pub fn shared_formula_index(&self) -> Option<u32> {
        self.shared_formula_index
    }

    /// Sets the shared formula index.
    pub fn set_shared_formula_index(&mut self, index: u32) -> &mut Self {
        self.shared_formula_index = Some(index);
        self
    }

    /// Clears the shared formula index.
    pub fn clear_shared_formula_index(&mut self) -> &mut Self {
        self.shared_formula_index = None;
        self
    }

    /// Clears all cell data (value, formula, style, comment, etc.).
    ///
    /// Resets the cell to its default empty state.
    pub fn clear(&mut self) -> &mut Self {
        self.value = None;
        self.formula = None;
        self.cached_value = None;
        self.style_id = None;
        self.comment = None;
        self.is_array_formula = false;
        self.array_range = None;
        self.shared_formula_index = None;
        self.unknown_attrs.clear();
        self.unknown_children.clear();
        self
    }

    // ---- Rich text helpers ----

    /// Returns the rich text runs if the cell value is `RichText`.
    pub fn rich_text(&self) -> Option<&[RichTextRun]> {
        match &self.value {
            Some(CellValue::RichText(runs)) => Some(runs.as_slice()),
            _ => None,
        }
    }

    /// Sets this cell's value to rich text.
    pub fn set_rich_text(&mut self, runs: Vec<RichTextRun>) -> &mut Self {
        self.value = Some(CellValue::RichText(runs));
        self
    }

    /// Sets the formula reference range without toggling the array formula flag.
    ///
    /// Used internally by the parser to store the `ref` attribute of `<f>` for
    /// shared formulas (where `is_array_formula` must remain `false`).
    pub(crate) fn set_formula_ref(&mut self, range: impl Into<String>) {
        self.array_range = Some(range.into());
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

/// Built-in number format IDs that represent date/time formats.
/// If a cell's style references one of these format IDs, the numeric value
/// should be interpreted as a DateTime serial number.
pub const BUILTIN_DATE_FORMAT_IDS: &[u32] = &[
    14, 15, 16, 17, 18, 19, 20, 21, 22,
    // These additional IDs are also date formats in many locales:
    27, 28, 29, 30, 31, 32, 33, 34, 35, 36, 45, 46, 47, 50, 51, 52, 53, 54, 55, 56, 57, 58,
];

/// Returns true if the given number format ID is a built-in date format.
pub fn is_builtin_date_format(num_fmt_id: u32) -> bool {
    BUILTIN_DATE_FORMAT_IDS.contains(&num_fmt_id)
}

pub(crate) fn normalize_cell_reference(reference: &str) -> Result<String> {
    let trimmed = reference.trim();
    if trimmed.is_empty() {
        return Err(XlsxError::InvalidCellReference(reference.to_string()));
    }

    let split_idx = trimmed
        .char_indices()
        .find_map(|(idx, ch)| if ch.is_ascii_digit() { Some(idx) } else { None })
        .ok_or_else(|| XlsxError::InvalidCellReference(trimmed.to_string()))?;

    let (column, row) = trimmed.split_at(split_idx);

    if column.is_empty()
        || row.is_empty()
        || !column.chars().all(|ch| ch.is_ascii_alphabetic())
        || !row.chars().all(|ch| ch.is_ascii_digit())
        || row.starts_with('0')
    {
        return Err(XlsxError::InvalidCellReference(trimmed.to_string()));
    }

    let row_number: u32 = row
        .parse()
        .map_err(|_| XlsxError::InvalidCellReference(trimmed.to_string()))?;

    if row_number == 0 {
        return Err(XlsxError::InvalidCellReference(trimmed.to_string()));
    }

    let normalized_column: String = column.chars().map(|ch| ch.to_ascii_uppercase()).collect();

    Ok(format!("{normalized_column}{row_number}"))
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn cell_setters_and_getters_work() {
        let mut cell = Cell::new();

        assert!(cell.value().is_none());
        assert!(cell.formula().is_none());
        assert!(cell.style_id().is_none());

        cell.set_value("hello")
            .set_formula("SUM(A1:A5)")
            .set_style_id(7);

        assert_eq!(cell.value(), Some(&CellValue::String("hello".to_string())));
        assert_eq!(cell.formula(), Some("SUM(A1:A5)"));
        assert_eq!(cell.style_id(), Some(7));

        cell.clear_value().clear_formula().clear_style_id();

        assert!(cell.value().is_none());
        assert!(cell.formula().is_none());
        assert!(cell.style_id().is_none());
    }

    #[test]
    fn date_and_error_value_variants_are_accessible() {
        let mut cell = Cell::new();
        cell.set_value(CellValue::date("2025-01-31"));
        assert_eq!(
            cell.value(),
            Some(&CellValue::Date("2025-01-31".to_string()))
        );

        cell.set_value(CellValue::error("#DIV/0!"));
        assert_eq!(cell.value(), Some(&CellValue::Error("#DIV/0!".to_string())));
    }

    #[test]
    fn normalize_cell_reference_validates_input() {
        assert_eq!(normalize_cell_reference("a1").unwrap(), "A1");
        assert_eq!(normalize_cell_reference("  bc12 ").unwrap(), "BC12");

        assert!(normalize_cell_reference("").is_err());
        assert!(normalize_cell_reference("A0").is_err());
        assert!(normalize_cell_reference("12").is_err());
        assert!(normalize_cell_reference("A").is_err());
        assert!(normalize_cell_reference("A-1").is_err());
    }

    // ===== Feature 2: CellComment =====

    #[test]
    fn cell_comment_accessors_work() {
        let mut comment = CellComment::new("Alice", "This is a note");
        assert_eq!(comment.author(), "Alice");
        assert_eq!(comment.text(), "This is a note");

        comment.set_author("Bob").set_text("Updated note");
        assert_eq!(comment.author(), "Bob");
        assert_eq!(comment.text(), "Updated note");
    }

    #[test]
    fn cell_comment_set_and_clear_on_cell() {
        let mut cell = Cell::new();
        assert!(cell.comment().is_none());

        cell.set_comment(CellComment::new("Author", "Hello"));
        assert!(cell.comment().is_some());
        assert_eq!(cell.comment().unwrap().author(), "Author");
        assert_eq!(cell.comment().unwrap().text(), "Hello");

        cell.clear_comment();
        assert!(cell.comment().is_none());
    }

    // ===== Feature 3: Rich text =====

    #[test]
    fn rich_text_run_formatting() {
        let mut run = RichTextRun::new("Hello");
        assert_eq!(run.text(), "Hello");
        assert!(!run.has_formatting());

        run.set_bold(true)
            .set_italic(true)
            .set_font_name("Arial")
            .set_font_size("12")
            .set_color("FFFF0000");
        assert!(run.has_formatting());
        assert_eq!(run.bold(), Some(true));
        assert_eq!(run.italic(), Some(true));
        assert_eq!(run.font_name(), Some("Arial"));
        assert_eq!(run.font_size(), Some("12"));
        assert_eq!(run.color(), Some("FFFF0000"));

        run.clear_bold()
            .clear_italic()
            .clear_font_name()
            .clear_font_size()
            .clear_color();
        assert!(!run.has_formatting());
    }

    #[test]
    fn cell_rich_text_value() {
        let mut cell = Cell::new();
        assert!(cell.rich_text().is_none());

        let mut run1 = RichTextRun::new("Bold ");
        run1.set_bold(true);
        let run2 = RichTextRun::new("normal");

        cell.set_rich_text(vec![run1.clone(), run2]);
        assert!(cell.rich_text().is_some());
        assert_eq!(cell.rich_text().unwrap().len(), 2);
        assert_eq!(cell.rich_text().unwrap()[0].text(), "Bold ");
        assert_eq!(cell.rich_text().unwrap()[0].bold(), Some(true));
        assert_eq!(cell.rich_text().unwrap()[1].text(), "normal");
    }

    #[test]
    fn rich_text_run_from_font() {
        let mut font = Font::new();
        font.set_bold(true)
            .set_name("Calibri")
            .set_size("11")
            .set_color("FF000000");
        let run = RichTextRun::from_font("styled", &font);
        assert_eq!(run.text(), "styled");
        assert_eq!(run.bold(), Some(true));
        assert_eq!(run.font_name(), Some("Calibri"));
        assert_eq!(run.font_size(), Some("11"));
        assert_eq!(run.color(), Some("FF000000"));
    }

    // ===== Feature 5: DateTime type =====

    #[test]
    fn datetime_cell_value() {
        let dt = CellValue::date_time(44927.5); // 2023-01-01 12:00
        assert!(dt.is_date_time());
        assert_eq!(dt.as_date_time(), Some(44927.5));

        let num = CellValue::Number(44927.5);
        assert!(!num.is_date_time());
        assert!(num.as_date_time().is_none());
    }

    #[test]
    fn datetime_roundtrip_on_cell() {
        let mut cell = Cell::new();
        cell.set_value(CellValue::date_time(44927.0));
        assert_eq!(cell.value(), Some(&CellValue::DateTime(44927.0)));
        assert!(cell.value().unwrap().is_date_time());
    }

    #[test]
    fn builtin_date_format_detection() {
        assert!(is_builtin_date_format(14));
        assert!(is_builtin_date_format(22));
        assert!(!is_builtin_date_format(0));
        assert!(!is_builtin_date_format(1));
        assert!(!is_builtin_date_format(13));
        assert!(!is_builtin_date_format(23));
    }

    // ===== Feature 11: Array/CSE formulas =====

    #[test]
    fn array_formula_flag() {
        let mut cell = Cell::new();
        assert!(!cell.is_array_formula());
        assert!(cell.array_range().is_none());

        cell.set_array_formula(true);
        assert!(cell.is_array_formula());

        cell.set_array_formula(false);
        assert!(!cell.is_array_formula());
    }

    #[test]
    fn array_formula_range() {
        let mut cell = Cell::new();
        cell.set_formula("{SUM(A1:A3*B1:B3)}")
            .set_array_range("C1:C3");

        assert!(cell.is_array_formula());
        assert_eq!(cell.array_range(), Some("C1:C3"));
        assert_eq!(cell.formula(), Some("{SUM(A1:A3*B1:B3)}"));

        cell.clear_array_range();
        assert!(!cell.is_array_formula());
        assert!(cell.array_range().is_none());
    }

    // ===== Feature 12: Cached values =====

    #[test]
    fn cached_value_accessors() {
        let mut cell = Cell::new();
        assert!(cell.cached_value().is_none());

        cell.set_formula("SUM(A1:A5)")
            .set_cached_value(CellValue::Number(42.0));

        assert_eq!(cell.formula(), Some("SUM(A1:A5)"));
        assert_eq!(cell.cached_value(), Some(&CellValue::Number(42.0)));

        cell.clear_cached_value();
        assert!(cell.cached_value().is_none());
    }

    #[test]
    fn cached_value_with_string_formula_result() {
        let mut cell = Cell::new();
        cell.set_formula("CONCATENATE(A1,B1)")
            .set_cached_value(CellValue::String("HelloWorld".to_string()));

        assert_eq!(cell.formula(), Some("CONCATENATE(A1,B1)"));
        assert_eq!(
            cell.cached_value(),
            Some(&CellValue::String("HelloWorld".to_string()))
        );
    }

    // ===== Feature 12: Shared formulas =====

    #[test]
    fn shared_formula_index_accessors() {
        let mut cell = Cell::new();
        assert!(cell.shared_formula_index().is_none());

        cell.set_shared_formula_index(0);
        assert_eq!(cell.shared_formula_index(), Some(0));

        cell.set_shared_formula_index(5);
        assert_eq!(cell.shared_formula_index(), Some(5));

        cell.clear_shared_formula_index();
        assert!(cell.shared_formula_index().is_none());
    }

    #[test]
    fn shared_formula_master_cell() {
        let mut cell = Cell::new();
        cell.set_formula("SUM(B1:C1)").set_shared_formula_index(0);
        // Use set_formula_ref (not set_array_range) so is_array_formula stays false
        cell.set_formula_ref("A1:A10");

        assert_eq!(cell.formula(), Some("SUM(B1:C1)"));
        assert_eq!(cell.shared_formula_index(), Some(0));
        assert_eq!(cell.array_range(), Some("A1:A10"));
        assert!(!cell.is_array_formula());
    }

    #[test]
    fn shared_formula_dependent_cell() {
        // Dependent cells have shared_formula_index but no formula text
        let mut cell = Cell::new();
        cell.set_shared_formula_index(0);

        assert!(cell.formula().is_none());
        assert_eq!(cell.shared_formula_index(), Some(0));
    }
}
