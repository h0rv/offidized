//! Pivot table domain objects and API.
//!
//! This module provides high-level types for working with Excel pivot tables,
//! including table definitions, fields (row/column/page/data), and pivot cache.

/// A pivot table in a worksheet.
///
/// Pivot tables provide interactive data summarization and analysis. They reference
/// a data source (typically a worksheet range or table) and allow users to organize
/// data by rows, columns, filters (page fields), and aggregate values.
#[derive(Debug, Clone, PartialEq)]
pub struct PivotTable {
    name: String,
    source_reference: PivotSourceReference,
    row_fields: Vec<PivotField>,
    column_fields: Vec<PivotField>,
    page_fields: Vec<PivotField>,
    data_fields: Vec<PivotDataField>,
    target_row: u32,
    target_col: u32,
    show_row_grand_totals: bool,
    show_column_grand_totals: bool,
    row_header_caption: Option<String>,
    column_header_caption: Option<String>,
    preserve_formatting: bool,
    use_auto_formatting: bool,
    page_wrap: u32,
    page_over_then_down: bool,
    subtotal_hidden_items: bool,
    row_grand_totals_caption: Option<String>,
    column_grand_totals_caption: Option<String>,
    field_print_titles: bool,
    item_print_titles: bool,
    merge_item: bool,
    indent: u32,
    outline_data: bool,
    compact: bool,
    compact_data: bool,
}

impl PivotTable {
    /// Creates a new pivot table with the given name and source reference.
    pub fn new(name: impl Into<String>, source_reference: PivotSourceReference) -> Self {
        Self {
            name: name.into(),
            source_reference,
            row_fields: Vec::new(),
            column_fields: Vec::new(),
            page_fields: Vec::new(),
            data_fields: Vec::new(),
            target_row: 0,
            target_col: 0,
            show_row_grand_totals: true,
            show_column_grand_totals: true,
            row_header_caption: None,
            column_header_caption: None,
            preserve_formatting: true,
            use_auto_formatting: false,
            page_wrap: 0,
            page_over_then_down: true,
            subtotal_hidden_items: false,
            row_grand_totals_caption: None,
            column_grand_totals_caption: None,
            field_print_titles: false,
            item_print_titles: false,
            merge_item: false,
            indent: 1,
            outline_data: false,
            compact: true,
            compact_data: true,
        }
    }

    /// Returns the pivot table name.
    pub fn name(&self) -> &str {
        &self.name
    }

    /// Sets the pivot table name.
    pub fn set_name(&mut self, name: impl Into<String>) -> &mut Self {
        self.name = name.into();
        self
    }

    /// Returns the source data reference.
    pub fn source_reference(&self) -> &PivotSourceReference {
        &self.source_reference
    }

    /// Sets the source data reference.
    pub fn set_source_reference(&mut self, source_reference: PivotSourceReference) -> &mut Self {
        self.source_reference = source_reference;
        self
    }

    /// Returns the target cell row (0-indexed).
    pub fn target_row(&self) -> u32 {
        self.target_row
    }

    /// Returns the target cell column (0-indexed).
    pub fn target_col(&self) -> u32 {
        self.target_col
    }

    /// Sets the target cell position (0-indexed row and column).
    pub fn set_target(&mut self, row: u32, col: u32) -> &mut Self {
        self.target_row = row;
        self.target_col = col;
        self
    }

    /// Returns the row fields (fields displayed as row headers).
    pub fn row_fields(&self) -> &[PivotField] {
        &self.row_fields
    }

    /// Returns mutable row fields.
    pub fn row_fields_mut(&mut self) -> &mut Vec<PivotField> {
        &mut self.row_fields
    }

    /// Adds a field to the row axis.
    pub fn add_row_field(&mut self, field: PivotField) -> &mut Self {
        self.row_fields.push(field);
        self
    }

    /// Returns the column fields (fields displayed as column headers).
    pub fn column_fields(&self) -> &[PivotField] {
        &self.column_fields
    }

    /// Returns mutable column fields.
    pub fn column_fields_mut(&mut self) -> &mut Vec<PivotField> {
        &mut self.column_fields
    }

    /// Adds a field to the column axis.
    pub fn add_column_field(&mut self, field: PivotField) -> &mut Self {
        self.column_fields.push(field);
        self
    }

    /// Returns the page fields (filter fields displayed above the pivot table).
    pub fn page_fields(&self) -> &[PivotField] {
        &self.page_fields
    }

    /// Returns mutable page fields.
    pub fn page_fields_mut(&mut self) -> &mut Vec<PivotField> {
        &mut self.page_fields
    }

    /// Adds a field to the page (filter) area.
    pub fn add_page_field(&mut self, field: PivotField) -> &mut Self {
        self.page_fields.push(field);
        self
    }

    /// Returns the data fields (aggregated values).
    pub fn data_fields(&self) -> &[PivotDataField] {
        &self.data_fields
    }

    /// Returns mutable data fields.
    pub fn data_fields_mut(&mut self) -> &mut Vec<PivotDataField> {
        &mut self.data_fields
    }

    /// Adds a data field (aggregated value column).
    pub fn add_data_field(&mut self, field: PivotDataField) -> &mut Self {
        self.data_fields.push(field);
        self
    }

    /// Returns whether row grand totals are shown.
    pub fn show_row_grand_totals(&self) -> bool {
        self.show_row_grand_totals
    }

    /// Sets whether to show row grand totals.
    pub fn set_show_row_grand_totals(&mut self, value: bool) -> &mut Self {
        self.show_row_grand_totals = value;
        self
    }

    /// Returns whether column grand totals are shown.
    pub fn show_column_grand_totals(&self) -> bool {
        self.show_column_grand_totals
    }

    /// Sets whether to show column grand totals.
    pub fn set_show_column_grand_totals(&mut self, value: bool) -> &mut Self {
        self.show_column_grand_totals = value;
        self
    }

    /// Returns the row header caption.
    pub fn row_header_caption(&self) -> Option<&str> {
        self.row_header_caption.as_deref()
    }

    /// Sets the row header caption.
    pub fn set_row_header_caption(&mut self, caption: impl Into<String>) -> &mut Self {
        self.row_header_caption = Some(caption.into());
        self
    }

    /// Clears the row header caption.
    pub fn clear_row_header_caption(&mut self) -> &mut Self {
        self.row_header_caption = None;
        self
    }

    /// Returns the column header caption.
    pub fn column_header_caption(&self) -> Option<&str> {
        self.column_header_caption.as_deref()
    }

    /// Sets the column header caption.
    pub fn set_column_header_caption(&mut self, caption: impl Into<String>) -> &mut Self {
        self.column_header_caption = Some(caption.into());
        self
    }

    /// Clears the column header caption.
    pub fn clear_column_header_caption(&mut self) -> &mut Self {
        self.column_header_caption = None;
        self
    }

    /// Returns whether cell formatting is preserved when the pivot table is refreshed.
    pub fn preserve_formatting(&self) -> bool {
        self.preserve_formatting
    }

    /// Sets whether to preserve cell formatting on refresh.
    pub fn set_preserve_formatting(&mut self, value: bool) -> &mut Self {
        self.preserve_formatting = value;
        self
    }

    /// Returns whether auto-formatting is enabled.
    pub fn use_auto_formatting(&self) -> bool {
        self.use_auto_formatting
    }

    /// Sets whether to use auto-formatting.
    pub fn set_use_auto_formatting(&mut self, value: bool) -> &mut Self {
        self.use_auto_formatting = value;
        self
    }

    /// Returns the page wrap value (number of page fields per row/column).
    pub fn page_wrap(&self) -> u32 {
        self.page_wrap
    }

    /// Sets the page wrap value. 0 means no limit.
    pub fn set_page_wrap(&mut self, value: u32) -> &mut Self {
        self.page_wrap = value;
        self
    }

    /// Returns whether page fields flow over then down (true) or down then over (false).
    pub fn page_over_then_down(&self) -> bool {
        self.page_over_then_down
    }

    /// Sets the page field layout order.
    pub fn set_page_over_then_down(&mut self, value: bool) -> &mut Self {
        self.page_over_then_down = value;
        self
    }

    /// Returns whether to include filtered items in subtotals.
    pub fn subtotal_hidden_items(&self) -> bool {
        self.subtotal_hidden_items
    }

    /// Sets whether to include filtered items in subtotals.
    pub fn set_subtotal_hidden_items(&mut self, value: bool) -> &mut Self {
        self.subtotal_hidden_items = value;
        self
    }
}

/// Reference to the pivot table data source.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum PivotSourceReference {
    /// A worksheet range reference (e.g., "Sheet1!$A$1:$D$100").
    WorksheetRange(String),
    /// A named table reference (e.g., "Table1").
    NamedTable(String),
}

impl PivotSourceReference {
    /// Creates a worksheet range source reference.
    pub fn from_range(range: impl Into<String>) -> Self {
        Self::WorksheetRange(range.into())
    }

    /// Creates a named table source reference.
    pub fn from_table(table_name: impl Into<String>) -> Self {
        Self::NamedTable(table_name.into())
    }

    /// Returns the source reference as a string.
    pub fn as_str(&self) -> &str {
        match self {
            Self::WorksheetRange(r) => r,
            Self::NamedTable(t) => t,
        }
    }
}

/// A field in a pivot table (row, column, or page field).
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct PivotField {
    /// The field name (typically a column header from the source data).
    name: String,
    /// Custom label to display (if different from the field name).
    custom_label: Option<String>,
    /// Whether to show subtotals for this field.
    show_all_subtotals: bool,
    /// Whether to insert blank lines after each item.
    insert_blank_rows: bool,
    /// Whether to show items with no data.
    show_empty_items: bool,
    /// Sort order.
    sort_type: PivotFieldSort,
    /// Whether to insert page breaks after each item.
    insert_page_breaks: bool,
}

impl PivotField {
    /// Creates a new pivot field with the given name.
    pub fn new(name: impl Into<String>) -> Self {
        Self {
            name: name.into(),
            custom_label: None,
            show_all_subtotals: false,
            insert_blank_rows: false,
            show_empty_items: false,
            sort_type: PivotFieldSort::Ascending,
            insert_page_breaks: false,
        }
    }

    /// Returns the field name.
    pub fn name(&self) -> &str {
        &self.name
    }

    /// Sets the field name.
    pub fn set_name(&mut self, name: impl Into<String>) -> &mut Self {
        self.name = name.into();
        self
    }

    /// Returns the custom label, if set.
    pub fn custom_label(&self) -> Option<&str> {
        self.custom_label.as_deref()
    }

    /// Sets a custom label to display instead of the field name.
    pub fn set_custom_label(&mut self, label: impl Into<String>) -> &mut Self {
        self.custom_label = Some(label.into());
        self
    }

    /// Clears the custom label.
    pub fn clear_custom_label(&mut self) -> &mut Self {
        self.custom_label = None;
        self
    }

    /// Returns whether all subtotals are shown for this field.
    pub fn show_all_subtotals(&self) -> bool {
        self.show_all_subtotals
    }

    /// Sets whether to show subtotals.
    pub fn set_show_all_subtotals(&mut self, value: bool) -> &mut Self {
        self.show_all_subtotals = value;
        self
    }

    /// Returns whether blank rows are inserted after each item.
    pub fn insert_blank_rows(&self) -> bool {
        self.insert_blank_rows
    }

    /// Sets whether to insert blank rows.
    pub fn set_insert_blank_rows(&mut self, value: bool) -> &mut Self {
        self.insert_blank_rows = value;
        self
    }

    /// Returns whether items with no data are shown.
    pub fn show_empty_items(&self) -> bool {
        self.show_empty_items
    }

    /// Sets whether to show empty items.
    pub fn set_show_empty_items(&mut self, value: bool) -> &mut Self {
        self.show_empty_items = value;
        self
    }

    /// Returns the sort type for this field.
    pub fn sort_type(&self) -> PivotFieldSort {
        self.sort_type
    }

    /// Sets the sort type.
    pub fn set_sort_type(&mut self, sort_type: PivotFieldSort) -> &mut Self {
        self.sort_type = sort_type;
        self
    }

    /// Returns whether page breaks are inserted after each item.
    pub fn insert_page_breaks(&self) -> bool {
        self.insert_page_breaks
    }

    /// Sets whether to insert page breaks.
    pub fn set_insert_page_breaks(&mut self, value: bool) -> &mut Self {
        self.insert_page_breaks = value;
        self
    }
}

/// Sort order for a pivot field.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum PivotFieldSort {
    /// Manual sorting (user-defined order).
    Manual,
    /// Ascending sort (A-Z, smallest to largest).
    #[default]
    Ascending,
    /// Descending sort (Z-A, largest to smallest).
    Descending,
}

impl PivotFieldSort {
    #[allow(dead_code)]
    pub(crate) fn as_xml_value(self) -> &'static str {
        match self {
            Self::Manual => "manual",
            Self::Ascending => "ascending",
            Self::Descending => "descending",
        }
    }

    #[allow(dead_code)]
    pub(crate) fn from_xml_value(value: &str) -> Self {
        match value.trim() {
            "ascending" => Self::Ascending,
            "descending" => Self::Descending,
            _ => Self::Manual,
        }
    }
}

/// A data field in a pivot table (an aggregated value column).
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct PivotDataField {
    /// The source field name from the data.
    field_name: String,
    /// Custom name to display (e.g., "Sum of Sales").
    custom_name: Option<String>,
    /// The aggregation function to apply.
    subtotal: PivotSubtotalFunction,
    /// Number format for the values (optional).
    number_format: Option<String>,
}

impl PivotDataField {
    /// Creates a new data field with the given field name.
    pub fn new(field_name: impl Into<String>) -> Self {
        Self {
            field_name: field_name.into(),
            custom_name: None,
            subtotal: PivotSubtotalFunction::Sum,
            number_format: None,
        }
    }

    /// Returns the source field name.
    pub fn field_name(&self) -> &str {
        &self.field_name
    }

    /// Sets the source field name.
    pub fn set_field_name(&mut self, name: impl Into<String>) -> &mut Self {
        self.field_name = name.into();
        self
    }

    /// Returns the custom display name.
    pub fn custom_name(&self) -> Option<&str> {
        self.custom_name.as_deref()
    }

    /// Sets the custom display name.
    pub fn set_custom_name(&mut self, name: impl Into<String>) -> &mut Self {
        self.custom_name = Some(name.into());
        self
    }

    /// Clears the custom name.
    pub fn clear_custom_name(&mut self) -> &mut Self {
        self.custom_name = None;
        self
    }

    /// Returns the subtotal function.
    pub fn subtotal(&self) -> PivotSubtotalFunction {
        self.subtotal
    }

    /// Sets the subtotal function.
    pub fn set_subtotal(&mut self, subtotal: PivotSubtotalFunction) -> &mut Self {
        self.subtotal = subtotal;
        self
    }

    /// Returns the number format.
    pub fn number_format(&self) -> Option<&str> {
        self.number_format.as_deref()
    }

    /// Sets the number format.
    pub fn set_number_format(&mut self, format: impl Into<String>) -> &mut Self {
        self.number_format = Some(format.into());
        self
    }

    /// Clears the number format.
    pub fn clear_number_format(&mut self) -> &mut Self {
        self.number_format = None;
        self
    }
}

/// Aggregation function for a data field.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum PivotSubtotalFunction {
    /// Average of values.
    Average,
    /// Count of values.
    Count,
    /// Count of numeric values.
    CountNums,
    /// Maximum value.
    Max,
    /// Minimum value.
    Min,
    /// Product of values.
    Product,
    /// Standard deviation (sample).
    StdDev,
    /// Standard deviation (population).
    StdDevP,
    /// Sum of values (default).
    #[default]
    Sum,
    /// Variance (sample).
    Var,
    /// Variance (population).
    VarP,
}

impl PivotSubtotalFunction {
    #[allow(dead_code)]
    pub(crate) fn as_xml_value(self) -> &'static str {
        match self {
            Self::Average => "average",
            Self::Count => "count",
            Self::CountNums => "countNums",
            Self::Max => "max",
            Self::Min => "min",
            Self::Product => "product",
            Self::StdDev => "stdDev",
            Self::StdDevP => "stdDevP",
            Self::Sum => "sum",
            Self::Var => "var",
            Self::VarP => "varP",
        }
    }

    pub(crate) fn from_xml_value(value: &str) -> Self {
        match value.trim() {
            "average" => Self::Average,
            "count" => Self::Count,
            "countNums" => Self::CountNums,
            "max" => Self::Max,
            "min" => Self::Min,
            "product" => Self::Product,
            "stdDev" => Self::StdDev,
            "stdDevP" => Self::StdDevP,
            "sum" => Self::Sum,
            "var" => Self::Var,
            "varP" => Self::VarP,
            _ => Self::Sum,
        }
    }
}
