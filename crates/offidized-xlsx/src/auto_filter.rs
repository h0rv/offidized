//! Auto-filter criteria types for worksheet and table auto-filters.
//!
//! An [`AutoFilter`] captures both the range and the per-column filter criteria
//! that Excel stores inside `<autoFilter>` elements.

use crate::error::Result;
use crate::range::CellRange;

/// The kind of filter applied to a column.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum FilterType {
    /// No filter applied (default).
    #[default]
    None,
    /// A discrete value list filter (`<filters>`).
    Values,
    /// One or two custom comparison filters (`<customFilters>`).
    Custom,
    /// Top-N / bottom-N filter (`<top10>`).
    Top10,
    /// Dynamic date/value filter (`<dynamicFilter>`).
    Dynamic,
    /// Filter by cell or font colour (`<colorFilter>`).
    Color,
    /// Filter by conditional-formatting icon (`<iconFilter>`).
    Icon,
}

impl FilterType {
    /// Returns the OOXML element name for the filter type, if applicable.
    pub fn as_xml_element(&self) -> Option<&'static str> {
        match self {
            Self::None => None,
            Self::Values => Some("filters"),
            Self::Custom => Some("customFilters"),
            Self::Top10 => Some("top10"),
            Self::Dynamic => Some("dynamicFilter"),
            Self::Color => Some("colorFilter"),
            Self::Icon => Some("iconFilter"),
        }
    }
}

/// Comparison operator used in a [`CustomFilter`].
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum CustomFilterOperator {
    /// Equals (`=`).
    #[default]
    Equal,
    /// Not equal (`!=`).
    NotEqual,
    /// Greater than (`>`).
    GreaterThan,
    /// Greater than or equal (`>=`).
    GreaterThanOrEqual,
    /// Less than (`<`).
    LessThan,
    /// Less than or equal (`<=`).
    LessThanOrEqual,
}

impl CustomFilterOperator {
    /// Returns the OOXML attribute value for this operator.
    pub fn as_str(&self) -> &'static str {
        match self {
            Self::Equal => "equal",
            Self::NotEqual => "notEqual",
            Self::GreaterThan => "greaterThan",
            Self::GreaterThanOrEqual => "greaterThanOrEqual",
            Self::LessThan => "lessThan",
            Self::LessThanOrEqual => "lessThanOrEqual",
        }
    }

    /// Parses an OOXML attribute value into a [`CustomFilterOperator`].
    pub fn from_xml_value(value: &str) -> Self {
        match value {
            "equal" => Self::Equal,
            "notEqual" => Self::NotEqual,
            "greaterThan" => Self::GreaterThan,
            "greaterThanOrEqual" => Self::GreaterThanOrEqual,
            "lessThan" => Self::LessThan,
            "lessThanOrEqual" => Self::LessThanOrEqual,
            _ => Self::Equal,
        }
    }
}

/// A single custom comparison filter criterion.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct CustomFilter {
    operator: CustomFilterOperator,
    val: String,
}

impl CustomFilter {
    /// Creates a new custom filter with the given operator and value.
    pub fn new(operator: CustomFilterOperator, val: impl Into<String>) -> Self {
        Self {
            operator,
            val: val.into(),
        }
    }

    /// Returns the comparison operator.
    pub fn operator(&self) -> CustomFilterOperator {
        self.operator
    }

    /// Sets the comparison operator.
    pub fn set_operator(&mut self, operator: CustomFilterOperator) -> &mut Self {
        self.operator = operator;
        self
    }

    /// Returns the comparison value.
    pub fn val(&self) -> &str {
        &self.val
    }

    /// Sets the comparison value.
    pub fn set_val(&mut self, val: impl Into<String>) -> &mut Self {
        self.val = val.into();
        self
    }
}

/// Per-column filter settings within an [`AutoFilter`].
#[derive(Debug, Clone, PartialEq, Default)]
pub struct FilterColumn {
    /// Zero-based column id relative to the auto-filter range.
    col_id: u32,
    /// Whether any filter criterion is active on this column.
    has_filter: bool,
    /// The kind of filter applied.
    filter_type: FilterType,
    /// Whether the column drop-down button is visible (default `true` when absent in XML).
    show_button: Option<bool>,

    // -- Values filter --
    /// Whether blank cells are included in a values filter (`<filters blank="1"/>`).
    blank: Option<bool>,
    /// Discrete accepted values (`<filter val="..."/>`).
    values: Vec<String>,

    // -- Custom filter --
    /// Custom comparison filters (1 or 2).
    custom_filters: Vec<CustomFilter>,
    /// When two custom filters are present, `true` means AND logic, `false` means OR.
    custom_filters_and: Option<bool>,

    // -- Top10 filter --
    /// Whether to select the top values (`true`) or bottom values (`false`).
    top: Option<bool>,
    /// Whether the value is a percentage rather than an item count.
    percent: Option<bool>,
    /// The top/bottom N value.
    top10_val: Option<f64>,

    // -- Dynamic filter --
    /// The `type` attribute on `<dynamicFilter>`, e.g. "aboveAverage", "today".
    dynamic_type: Option<String>,

    // -- Color filter --
    /// DXF (differential formatting) id for colour-based filtering.
    dxf_id: Option<u32>,
    /// Whether the filter targets the cell background colour (`true`) or font colour (`false`).
    cell_color: Option<bool>,

    // -- Icon filter --
    /// Icon set name, e.g. "3Arrows".
    icon_set: Option<String>,
    /// Zero-based icon index within the set.
    icon_id: Option<u32>,
}

impl FilterColumn {
    /// Creates a new filter column for the given zero-based column id.
    pub fn new(col_id: u32) -> Self {
        Self {
            col_id,
            ..Self::default()
        }
    }

    // -- Core accessors --

    /// Returns the zero-based column id.
    pub fn col_id(&self) -> u32 {
        self.col_id
    }

    /// Sets the zero-based column id.
    pub fn set_col_id(&mut self, col_id: u32) -> &mut Self {
        self.col_id = col_id;
        self
    }

    /// Returns whether any filter criterion is active.
    pub fn has_filter(&self) -> bool {
        self.has_filter
    }

    /// Returns the filter type.
    pub fn filter_type(&self) -> FilterType {
        self.filter_type
    }

    /// Sets the filter type.
    pub fn set_filter_type(&mut self, filter_type: FilterType) -> &mut Self {
        self.filter_type = filter_type;
        self.has_filter = filter_type != FilterType::None;
        self
    }

    /// Returns whether the column drop-down button is visible.
    pub fn show_button(&self) -> Option<bool> {
        self.show_button
    }

    /// Sets whether the column drop-down button is visible.
    pub fn set_show_button(&mut self, show_button: bool) -> &mut Self {
        self.show_button = Some(show_button);
        self
    }

    // -- Values filter accessors --

    /// Returns whether blank cells are included in the values filter.
    pub fn blank(&self) -> Option<bool> {
        self.blank
    }

    /// Sets whether blank cells are included in the values filter.
    pub fn set_blank(&mut self, value: bool) -> &mut Self {
        self.blank = Some(value);
        self.filter_type = FilterType::Values;
        self.has_filter = true;
        self
    }

    /// Returns the discrete accepted values.
    pub fn values(&self) -> &[String] {
        &self.values
    }

    /// Adds a discrete accepted value.
    pub fn add_value(&mut self, value: impl Into<String>) -> &mut Self {
        self.values.push(value.into());
        self.filter_type = FilterType::Values;
        self.has_filter = true;
        self
    }

    /// Sets the full list of discrete accepted values.
    pub fn set_values(&mut self, values: Vec<String>) -> &mut Self {
        self.has_filter = !values.is_empty();
        if self.has_filter {
            self.filter_type = FilterType::Values;
        }
        self.values = values;
        self
    }

    // -- Custom filter accessors --

    /// Returns the custom comparison filters.
    pub fn custom_filters(&self) -> &[CustomFilter] {
        &self.custom_filters
    }

    /// Adds a custom comparison filter.
    pub fn add_custom_filter(&mut self, filter: CustomFilter) -> &mut Self {
        self.custom_filters.push(filter);
        self.filter_type = FilterType::Custom;
        self.has_filter = true;
        self
    }

    /// Sets the full list of custom comparison filters.
    pub fn set_custom_filters(&mut self, filters: Vec<CustomFilter>) -> &mut Self {
        self.has_filter = !filters.is_empty();
        if self.has_filter {
            self.filter_type = FilterType::Custom;
        }
        self.custom_filters = filters;
        self
    }

    /// Returns whether multiple custom filters use AND logic.
    pub fn custom_filters_and(&self) -> Option<bool> {
        self.custom_filters_and
    }

    /// Sets whether multiple custom filters use AND logic.
    pub fn set_custom_filters_and(&mut self, value: bool) -> &mut Self {
        self.custom_filters_and = Some(value);
        self
    }

    // -- Top10 filter accessors --

    /// Returns whether top values are selected (`true`) or bottom values (`false`).
    pub fn top(&self) -> Option<bool> {
        self.top
    }

    /// Sets whether top values are selected.
    pub fn set_top(&mut self, value: bool) -> &mut Self {
        self.top = Some(value);
        self.filter_type = FilterType::Top10;
        self.has_filter = true;
        self
    }

    /// Returns whether the top10 value is a percentage.
    pub fn percent(&self) -> Option<bool> {
        self.percent
    }

    /// Sets whether the top10 value is a percentage.
    pub fn set_percent(&mut self, value: bool) -> &mut Self {
        self.percent = Some(value);
        self
    }

    /// Returns the top/bottom N value.
    pub fn top10_val(&self) -> Option<f64> {
        self.top10_val
    }

    /// Sets the top/bottom N value.
    pub fn set_top10_val(&mut self, value: f64) -> &mut Self {
        self.top10_val = Some(value);
        self.filter_type = FilterType::Top10;
        self.has_filter = true;
        self
    }

    // -- Dynamic filter accessors --

    /// Returns the dynamic filter type string.
    pub fn dynamic_type(&self) -> Option<&str> {
        self.dynamic_type.as_deref()
    }

    /// Sets the dynamic filter type string.
    pub fn set_dynamic_type(&mut self, value: impl Into<String>) -> &mut Self {
        self.dynamic_type = Some(value.into());
        self.filter_type = FilterType::Dynamic;
        self.has_filter = true;
        self
    }

    // -- Color filter accessors --

    /// Returns the DXF id for colour-based filtering.
    pub fn dxf_id(&self) -> Option<u32> {
        self.dxf_id
    }

    /// Sets the DXF id for colour-based filtering.
    pub fn set_dxf_id(&mut self, value: u32) -> &mut Self {
        self.dxf_id = Some(value);
        self
    }

    /// Returns whether the filter targets cell background colour.
    pub fn cell_color(&self) -> Option<bool> {
        self.cell_color
    }

    /// Sets whether the filter targets cell background colour (`true`) or font colour (`false`).
    pub fn set_cell_color(&mut self, value: bool) -> &mut Self {
        self.cell_color = Some(value);
        self.filter_type = FilterType::Color;
        self.has_filter = true;
        self
    }

    // -- Icon filter accessors --

    /// Returns the icon set name.
    pub fn icon_set(&self) -> Option<&str> {
        self.icon_set.as_deref()
    }

    /// Sets the icon set name.
    pub fn set_icon_set(&mut self, value: impl Into<String>) -> &mut Self {
        self.icon_set = Some(value.into());
        self.filter_type = FilterType::Icon;
        self.has_filter = true;
        self
    }

    /// Returns the icon index within the set.
    pub fn icon_id(&self) -> Option<u32> {
        self.icon_id
    }

    /// Sets the icon index within the set.
    pub fn set_icon_id(&mut self, value: u32) -> &mut Self {
        self.icon_id = Some(value);
        self.filter_type = FilterType::Icon;
        self.has_filter = true;
        self
    }
}

/// Worksheet auto-filter definition including range and per-column criteria.
#[derive(Debug, Clone, PartialEq, Default)]
pub struct AutoFilter {
    /// The cell range covered by the auto-filter (e.g. "A1:D10").
    range: Option<CellRange>,
    /// Per-column filter criteria, indexed by zero-based column id relative to the range.
    filter_columns: Vec<FilterColumn>,
}

impl AutoFilter {
    /// Creates a new empty auto-filter.
    pub fn new() -> Self {
        Self::default()
    }

    /// Creates a new auto-filter for the given range.
    pub fn with_range(range: &str) -> Result<Self> {
        Ok(Self {
            range: Some(CellRange::parse(range)?),
            filter_columns: Vec::new(),
        })
    }

    /// Returns the auto-filter range.
    pub fn range(&self) -> Option<&CellRange> {
        self.range.as_ref()
    }

    /// Sets the auto-filter range.
    pub fn set_range(&mut self, range: &str) -> Result<&mut Self> {
        self.range = Some(CellRange::parse(range)?);
        Ok(self)
    }

    /// Clears the auto-filter range.
    pub fn clear_range(&mut self) -> &mut Self {
        self.range = None;
        self
    }

    /// Returns the per-column filter definitions.
    pub fn filter_columns(&self) -> &[FilterColumn] {
        &self.filter_columns
    }

    /// Returns a mutable reference to the per-column filter definitions.
    pub fn filter_columns_mut(&mut self) -> &mut Vec<FilterColumn> {
        &mut self.filter_columns
    }

    /// Adds a column filter definition. Returns a mutable reference to `self` for chaining.
    pub fn add_column_filter(&mut self, column: FilterColumn) -> &mut Self {
        self.filter_columns.push(column);
        self
    }

    /// Removes all column filter definitions.
    pub fn clear_filters(&mut self) -> &mut Self {
        self.filter_columns.clear();
        self
    }

    /// Returns `true` when neither a range nor any column filter is set.
    pub fn is_empty(&self) -> bool {
        self.range.is_none() && self.filter_columns.is_empty()
    }

    /// Returns the filter column for the given zero-based column id, if present.
    pub fn filter_column(&self, col_id: u32) -> Option<&FilterColumn> {
        self.filter_columns.iter().find(|c| c.col_id == col_id)
    }

    /// Returns a mutable reference to the filter column for the given zero-based column id.
    /// Creates a new entry if not present.
    pub fn filter_column_mut(&mut self, col_id: u32) -> &mut FilterColumn {
        if let Some(index) = self.filter_columns.iter().position(|c| c.col_id == col_id) {
            &mut self.filter_columns[index]
        } else {
            self.filter_columns.push(FilterColumn::new(col_id));
            let len = self.filter_columns.len();
            &mut self.filter_columns[len - 1]
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn auto_filter_with_range() {
        let af = AutoFilter::with_range("A1:D10").expect("valid range");
        let range = af.range().expect("range should be set");
        assert_eq!(range.start(), "A1");
        assert_eq!(range.end(), "D10");
        assert!(af.filter_columns().is_empty());
    }

    #[test]
    fn auto_filter_invalid_range() {
        assert!(AutoFilter::with_range("").is_err());
        assert!(AutoFilter::with_range("bad").is_err());
    }

    #[test]
    fn filter_column_values() {
        let mut col = FilterColumn::new(0);
        col.add_value("Alpha").add_value("Beta");
        assert_eq!(col.filter_type(), FilterType::Values);
        assert!(col.has_filter());
        assert_eq!(col.values().len(), 2);
        assert_eq!(col.values()[0], "Alpha");
        assert_eq!(col.values()[1], "Beta");
    }

    #[test]
    fn filter_column_custom() {
        let mut col = FilterColumn::new(1);
        col.add_custom_filter(CustomFilter::new(CustomFilterOperator::GreaterThan, "100"))
            .add_custom_filter(CustomFilter::new(CustomFilterOperator::LessThan, "500"))
            .set_custom_filters_and(true);

        assert_eq!(col.filter_type(), FilterType::Custom);
        assert_eq!(col.custom_filters().len(), 2);
        assert_eq!(
            col.custom_filters()[0].operator(),
            CustomFilterOperator::GreaterThan
        );
        assert_eq!(col.custom_filters()[0].val(), "100");
        assert_eq!(
            col.custom_filters()[1].operator(),
            CustomFilterOperator::LessThan
        );
        assert_eq!(col.custom_filters()[1].val(), "500");
        assert_eq!(col.custom_filters_and(), Some(true));
    }

    #[test]
    fn filter_column_top10() {
        let mut col = FilterColumn::new(2);
        col.set_top(true).set_top10_val(10.0).set_percent(false);
        assert_eq!(col.filter_type(), FilterType::Top10);
        assert_eq!(col.top(), Some(true));
        assert_eq!(col.top10_val(), Some(10.0));
        assert_eq!(col.percent(), Some(false));
    }

    #[test]
    fn filter_column_dynamic() {
        let mut col = FilterColumn::new(0);
        col.set_dynamic_type("aboveAverage");
        assert_eq!(col.filter_type(), FilterType::Dynamic);
        assert_eq!(col.dynamic_type(), Some("aboveAverage"));
    }

    #[test]
    fn filter_column_color() {
        let mut col = FilterColumn::new(0);
        col.set_cell_color(true).set_dxf_id(3);
        assert_eq!(col.filter_type(), FilterType::Color);
        assert_eq!(col.cell_color(), Some(true));
        assert_eq!(col.dxf_id(), Some(3));
    }

    #[test]
    fn filter_column_icon() {
        let mut col = FilterColumn::new(0);
        col.set_icon_set("3Arrows").set_icon_id(1);
        assert_eq!(col.filter_type(), FilterType::Icon);
        assert_eq!(col.icon_set(), Some("3Arrows"));
        assert_eq!(col.icon_id(), Some(1));
    }

    #[test]
    fn auto_filter_add_and_clear_columns() {
        let mut af = AutoFilter::with_range("A1:C10").unwrap();
        let mut col = FilterColumn::new(0);
        col.add_value("X");
        af.add_column_filter(col);
        assert_eq!(af.filter_columns().len(), 1);

        af.clear_filters();
        assert!(af.filter_columns().is_empty());
    }

    #[test]
    fn auto_filter_filter_column_mut_creates_entry() {
        let mut af = AutoFilter::new();
        af.filter_column_mut(2).add_value("Created");
        assert_eq!(af.filter_columns().len(), 1);
        assert_eq!(af.filter_column(2).unwrap().values()[0], "Created");
    }

    #[test]
    fn auto_filter_filter_column_mut_returns_existing() {
        let mut af = AutoFilter::new();
        af.filter_column_mut(0).add_value("First");
        af.filter_column_mut(0).add_value("Second");
        assert_eq!(af.filter_columns().len(), 1);
        assert_eq!(af.filter_column(0).unwrap().values().len(), 2);
    }

    #[test]
    fn custom_filter_operator_roundtrip() {
        let operators = [
            CustomFilterOperator::Equal,
            CustomFilterOperator::NotEqual,
            CustomFilterOperator::GreaterThan,
            CustomFilterOperator::GreaterThanOrEqual,
            CustomFilterOperator::LessThan,
            CustomFilterOperator::LessThanOrEqual,
        ];
        for op in operators {
            let parsed = CustomFilterOperator::from_xml_value(op.as_str());
            assert_eq!(parsed, op, "roundtrip failed for {}", op.as_str());
        }
    }

    #[test]
    fn custom_filter_operator_from_str_unknown_defaults_to_equal() {
        assert_eq!(
            CustomFilterOperator::from_xml_value("bogus"),
            CustomFilterOperator::Equal
        );
    }

    #[test]
    fn show_button_accessor() {
        let mut col = FilterColumn::new(0);
        assert_eq!(col.show_button(), None);
        col.set_show_button(false);
        assert_eq!(col.show_button(), Some(false));
    }

    #[test]
    fn auto_filter_is_empty() {
        let af = AutoFilter::new();
        assert!(af.is_empty());

        let af2 = AutoFilter::with_range("A1:B2").unwrap();
        assert!(!af2.is_empty());
    }

    #[test]
    fn custom_filter_setters() {
        let mut cf = CustomFilter::new(CustomFilterOperator::Equal, "hello");
        assert_eq!(cf.operator(), CustomFilterOperator::Equal);
        assert_eq!(cf.val(), "hello");

        cf.set_operator(CustomFilterOperator::NotEqual)
            .set_val("world");
        assert_eq!(cf.operator(), CustomFilterOperator::NotEqual);
        assert_eq!(cf.val(), "world");
    }
}
