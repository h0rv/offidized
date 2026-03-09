//! Chart domain model for SpreadsheetML charts.
//!
//! This module provides the high-level types for creating and manipulating
//! charts embedded in Excel worksheets. Charts are stored as separate XML
//! parts within the OOXML package, referenced via drawing relationships.
//!
//! # OOXML mapping
//!
//! Each chart type corresponds to an element under `<c:plotArea>` in the
//! chart XML part (e.g. `<c:barChart>`, `<c:lineChart>`, `<c:pieChart>`).
//! Series data is represented by `<c:ser>` elements, with category and
//! value references stored as `<c:numRef>` or `<c:strRef>` children.

// ===== ChartType =====

/// The type of chart to render.
///
/// Each variant maps to an OOXML chart element name under `<c:plotArea>`.
/// For combo charts, the top-level `ChartType` defines the primary type
/// and individual series can override with `ChartSeries::series_type`.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ChartType {
    /// Bar or column chart (`<c:barChart>`).
    Bar,
    /// Line chart (`<c:lineChart>`).
    Line,
    /// Pie chart (`<c:pieChart>`).
    Pie,
    /// Area chart (`<c:areaChart>`).
    Area,
    /// Scatter (XY) chart (`<c:scatterChart>`).
    Scatter,
    /// Doughnut chart (`<c:doughnutChart>`).
    Doughnut,
    /// Radar chart (`<c:radarChart>`).
    Radar,
    /// Bubble chart (`<c:bubbleChart>`).
    Bubble,
    /// Stock chart (`<c:stockChart>`).
    Stock,
    /// Surface chart (`<c:surfaceChart>`).
    Surface,
    /// Combo chart — uses per-series `series_type` to mix chart types.
    Combo,
}

impl ChartType {
    /// Returns the OOXML element name for this chart type.
    ///
    /// Combo charts do not have a single element name and return `"comboChart"`.
    pub fn as_str(self) -> &'static str {
        match self {
            Self::Bar => "barChart",
            Self::Line => "lineChart",
            Self::Pie => "pieChart",
            Self::Area => "areaChart",
            Self::Scatter => "scatterChart",
            Self::Doughnut => "doughnutChart",
            Self::Radar => "radarChart",
            Self::Bubble => "bubbleChart",
            Self::Stock => "stockChart",
            Self::Surface => "surfaceChart",
            Self::Combo => "comboChart",
        }
    }

    /// Parses a `ChartType` from an OOXML element name.
    ///
    /// Returns `None` if the string does not match a known chart element.
    pub fn from_xml_value(s: &str) -> Option<Self> {
        match s {
            "barChart" | "bar3DChart" => Some(Self::Bar),
            "lineChart" | "line3DChart" => Some(Self::Line),
            "pieChart" | "pie3DChart" => Some(Self::Pie),
            "areaChart" | "area3DChart" => Some(Self::Area),
            "scatterChart" => Some(Self::Scatter),
            "doughnutChart" => Some(Self::Doughnut),
            "radarChart" => Some(Self::Radar),
            "bubbleChart" => Some(Self::Bubble),
            "stockChart" => Some(Self::Stock),
            "surfaceChart" | "surface3DChart" => Some(Self::Surface),
            "comboChart" => Some(Self::Combo),
            _ => None,
        }
    }
}

// ===== BarDirection =====

/// The direction of a bar chart.
///
/// In OOXML, `<c:barDir val="col"/>` renders vertical columns and
/// `<c:barDir val="bar"/>` renders horizontal bars.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum BarDirection {
    /// Vertical columns (`val="col"`).
    Column,
    /// Horizontal bars (`val="bar"`).
    Bar,
}

impl BarDirection {
    /// Returns the OOXML attribute value for this direction.
    pub fn as_str(self) -> &'static str {
        match self {
            Self::Column => "col",
            Self::Bar => "bar",
        }
    }

    /// Parses a `BarDirection` from an OOXML attribute value.
    pub fn from_xml_value(s: &str) -> Option<Self> {
        match s {
            "col" => Some(Self::Column),
            "bar" => Some(Self::Bar),
            _ => None,
        }
    }
}

// ===== ChartGrouping =====

/// How series within a chart group are arranged.
///
/// Maps to `<c:grouping val="..."/>` in the chart XML.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ChartGrouping {
    /// Each series plotted independently (`"standard"`).
    Standard,
    /// Series stacked on top of each other (`"stacked"`).
    Stacked,
    /// Series stacked as percentage of total (`"percentStacked"`).
    PercentStacked,
    /// Series side-by-side within a category (`"clustered"`).
    Clustered,
}

impl ChartGrouping {
    /// Returns the OOXML attribute value for this grouping.
    pub fn as_str(self) -> &'static str {
        match self {
            Self::Standard => "standard",
            Self::Stacked => "stacked",
            Self::PercentStacked => "percentStacked",
            Self::Clustered => "clustered",
        }
    }

    /// Parses a `ChartGrouping` from an OOXML attribute value.
    pub fn from_xml_value(s: &str) -> Option<Self> {
        match s {
            "standard" => Some(Self::Standard),
            "stacked" => Some(Self::Stacked),
            "percentStacked" => Some(Self::PercentStacked),
            "clustered" => Some(Self::Clustered),
            _ => None,
        }
    }
}

// ===== ChartDataRef =====

/// A reference to chart data, combining an optional cell range formula
/// with cached values.
///
/// In OOXML, data references appear as `<c:numRef>` or `<c:strRef>`
/// elements containing a `<c:f>` formula and a `<c:numCache>` or
/// `<c:strCache>` with pre-computed values. The cache allows charts
/// to render without recalculating from the worksheet.
#[derive(Debug, Clone, PartialEq)]
pub struct ChartDataRef {
    /// Cell range formula (e.g. `"Sheet1!$B$2:$B$5"`).
    formula: Option<String>,
    /// Cached numeric values. `None` entries represent missing/blank cells.
    num_values: Vec<Option<f64>>,
    /// Cached string values.
    str_values: Vec<String>,
}

impl ChartDataRef {
    /// Creates an empty data reference with no formula or cached values.
    pub fn new() -> Self {
        Self {
            formula: None,
            num_values: Vec::new(),
            str_values: Vec::new(),
        }
    }

    /// Creates a data reference from a cell range formula.
    ///
    /// The formula should be in the standard OOXML format,
    /// e.g. `"Sheet1!$A$1:$A$10"`.
    pub fn from_formula(formula: impl Into<String>) -> Self {
        Self {
            formula: Some(formula.into()),
            num_values: Vec::new(),
            str_values: Vec::new(),
        }
    }

    /// Returns the cell range formula, if set.
    pub fn formula(&self) -> Option<&str> {
        self.formula.as_deref()
    }

    /// Sets the cell range formula.
    pub fn set_formula(&mut self, formula: impl Into<String>) -> &mut Self {
        self.formula = Some(formula.into());
        self
    }

    /// Clears the cell range formula.
    pub fn clear_formula(&mut self) -> &mut Self {
        self.formula = None;
        self
    }

    /// Returns the cached numeric values.
    pub fn num_values(&self) -> &[Option<f64>] {
        &self.num_values
    }

    /// Sets the cached numeric values.
    pub fn set_num_values(&mut self, values: Vec<Option<f64>>) -> &mut Self {
        self.num_values = values;
        self
    }

    /// Clears the cached numeric values.
    pub fn clear_num_values(&mut self) -> &mut Self {
        self.num_values.clear();
        self
    }

    /// Returns the cached string values.
    pub fn str_values(&self) -> &[String] {
        &self.str_values
    }

    /// Sets the cached string values.
    pub fn set_str_values(&mut self, values: Vec<String>) -> &mut Self {
        self.str_values = values;
        self
    }

    /// Clears the cached string values.
    pub fn clear_str_values(&mut self) -> &mut Self {
        self.str_values.clear();
        self
    }
}

impl Default for ChartDataRef {
    fn default() -> Self {
        Self::new()
    }
}

// ===== ChartSeries =====

/// A single data series within a chart.
///
/// Each series maps to a `<c:ser>` element in the chart XML. It contains
/// an index, order, optional name, and references to category/value data.
/// For scatter and bubble charts, `x_values` and `bubble_sizes` provide
/// additional data dimensions.
#[derive(Debug, Clone, PartialEq)]
pub struct ChartSeries {
    /// Zero-based index of the series (`<c:idx val="...">`).
    idx: u32,
    /// Plot order of the series (`<c:order val="...">`).
    order: u32,
    /// Display name for the series (literal text).
    name: Option<String>,
    /// Cell reference for the series name (`<c:tx><c:strRef><c:f>`).
    name_ref: Option<String>,
    /// Category (X-axis label) data reference.
    categories: Option<ChartDataRef>,
    /// Value (Y-axis) data reference.
    values: Option<ChartDataRef>,
    /// X-values for scatter and bubble charts.
    x_values: Option<ChartDataRef>,
    /// Bubble sizes for bubble charts.
    bubble_sizes: Option<ChartDataRef>,
    /// Fill color as a hex RGB string (e.g. `"FF4472C4"`).
    fill_color: Option<String>,
    /// Line/border color as a hex RGB string.
    line_color: Option<String>,
    /// Per-series chart type override for combo charts.
    series_type: Option<ChartType>,
}

impl ChartSeries {
    /// Creates a new series with the given index and order.
    pub fn new(idx: u32, order: u32) -> Self {
        Self {
            idx,
            order,
            name: None,
            name_ref: None,
            categories: None,
            values: None,
            x_values: None,
            bubble_sizes: None,
            fill_color: None,
            line_color: None,
            series_type: None,
        }
    }

    /// Returns the zero-based series index.
    pub fn idx(&self) -> u32 {
        self.idx
    }

    /// Sets the zero-based series index.
    pub fn set_idx(&mut self, idx: u32) -> &mut Self {
        self.idx = idx;
        self
    }

    /// Returns the plot order.
    pub fn order(&self) -> u32 {
        self.order
    }

    /// Sets the plot order.
    pub fn set_order(&mut self, order: u32) -> &mut Self {
        self.order = order;
        self
    }

    /// Returns the display name for this series.
    pub fn name(&self) -> Option<&str> {
        self.name.as_deref()
    }

    /// Sets the display name.
    pub fn set_name(&mut self, name: impl Into<String>) -> &mut Self {
        self.name = Some(name.into());
        self
    }

    /// Clears the display name.
    pub fn clear_name(&mut self) -> &mut Self {
        self.name = None;
        self
    }

    /// Returns the cell reference for the series name.
    pub fn name_ref(&self) -> Option<&str> {
        self.name_ref.as_deref()
    }

    /// Sets the cell reference for the series name.
    pub fn set_name_ref(&mut self, name_ref: impl Into<String>) -> &mut Self {
        self.name_ref = Some(name_ref.into());
        self
    }

    /// Clears the name reference.
    pub fn clear_name_ref(&mut self) -> &mut Self {
        self.name_ref = None;
        self
    }

    /// Returns the category data reference.
    pub fn categories(&self) -> Option<&ChartDataRef> {
        self.categories.as_ref()
    }

    /// Returns a mutable reference to the category data.
    pub fn categories_mut(&mut self) -> Option<&mut ChartDataRef> {
        self.categories.as_mut()
    }

    /// Sets the category data reference.
    pub fn set_categories(&mut self, data: ChartDataRef) -> &mut Self {
        self.categories = Some(data);
        self
    }

    /// Clears the category data reference.
    pub fn clear_categories(&mut self) -> &mut Self {
        self.categories = None;
        self
    }

    /// Returns the value data reference.
    pub fn values(&self) -> Option<&ChartDataRef> {
        self.values.as_ref()
    }

    /// Returns a mutable reference to the value data.
    pub fn values_mut(&mut self) -> Option<&mut ChartDataRef> {
        self.values.as_mut()
    }

    /// Sets the value data reference.
    pub fn set_values(&mut self, data: ChartDataRef) -> &mut Self {
        self.values = Some(data);
        self
    }

    /// Clears the value data reference.
    pub fn clear_values(&mut self) -> &mut Self {
        self.values = None;
        self
    }

    /// Returns the X-values reference (for scatter/bubble charts).
    pub fn x_values(&self) -> Option<&ChartDataRef> {
        self.x_values.as_ref()
    }

    /// Returns a mutable reference to the X-values data.
    pub fn x_values_mut(&mut self) -> Option<&mut ChartDataRef> {
        self.x_values.as_mut()
    }

    /// Sets the X-values reference.
    pub fn set_x_values(&mut self, data: ChartDataRef) -> &mut Self {
        self.x_values = Some(data);
        self
    }

    /// Clears the X-values reference.
    pub fn clear_x_values(&mut self) -> &mut Self {
        self.x_values = None;
        self
    }

    /// Returns the bubble sizes reference (for bubble charts).
    pub fn bubble_sizes(&self) -> Option<&ChartDataRef> {
        self.bubble_sizes.as_ref()
    }

    /// Returns a mutable reference to the bubble sizes data.
    pub fn bubble_sizes_mut(&mut self) -> Option<&mut ChartDataRef> {
        self.bubble_sizes.as_mut()
    }

    /// Sets the bubble sizes reference.
    pub fn set_bubble_sizes(&mut self, data: ChartDataRef) -> &mut Self {
        self.bubble_sizes = Some(data);
        self
    }

    /// Clears the bubble sizes reference.
    pub fn clear_bubble_sizes(&mut self) -> &mut Self {
        self.bubble_sizes = None;
        self
    }

    /// Returns the fill color (hex RGB string).
    pub fn fill_color(&self) -> Option<&str> {
        self.fill_color.as_deref()
    }

    /// Sets the fill color (hex RGB string, e.g. `"FF4472C4"`).
    pub fn set_fill_color(&mut self, color: impl Into<String>) -> &mut Self {
        self.fill_color = Some(color.into());
        self
    }

    /// Clears the fill color.
    pub fn clear_fill_color(&mut self) -> &mut Self {
        self.fill_color = None;
        self
    }

    /// Returns the line/border color (hex RGB string).
    pub fn line_color(&self) -> Option<&str> {
        self.line_color.as_deref()
    }

    /// Sets the line/border color (hex RGB string).
    pub fn set_line_color(&mut self, color: impl Into<String>) -> &mut Self {
        self.line_color = Some(color.into());
        self
    }

    /// Clears the line/border color.
    pub fn clear_line_color(&mut self) -> &mut Self {
        self.line_color = None;
        self
    }

    /// Returns the per-series chart type override (for combo charts).
    pub fn series_type(&self) -> Option<ChartType> {
        self.series_type
    }

    /// Sets the per-series chart type override.
    pub fn set_series_type(&mut self, chart_type: ChartType) -> &mut Self {
        self.series_type = Some(chart_type);
        self
    }

    /// Clears the per-series chart type override.
    pub fn clear_series_type(&mut self) -> &mut Self {
        self.series_type = None;
        self
    }
}

// ===== ChartAxis =====

/// An axis on a chart (category, value, date, or series axis).
///
/// Maps to `<c:catAx>`, `<c:valAx>`, `<c:dateAx>`, or `<c:serAx>`
/// elements in the chart XML.
#[derive(Debug, Clone, PartialEq)]
pub struct ChartAxis {
    /// Unique axis identifier (`<c:axId val="...">`).
    id: u32,
    /// Axis type: `"catAx"`, `"valAx"`, `"dateAx"`, or `"serAx"`.
    axis_type: String,
    /// Axis position: `"b"` (bottom), `"l"` (left), `"t"` (top), `"r"` (right).
    position: String,
    /// Axis title text.
    title: Option<String>,
    /// Minimum scale value.
    min: Option<f64>,
    /// Maximum scale value.
    max: Option<f64>,
    /// Major tick/gridline interval.
    major_unit: Option<f64>,
    /// Minor tick/gridline interval.
    minor_unit: Option<f64>,
    /// Whether major gridlines are displayed.
    major_gridlines: bool,
    /// Whether minor gridlines are displayed.
    minor_gridlines: bool,
    /// ID of the axis this axis crosses (`<c:crossAx val="...">`).
    crosses_ax: Option<u32>,
    /// Number format string for axis labels.
    num_fmt: Option<String>,
    /// Whether the axis is hidden (`<c:delete val="1">`).
    deleted: bool,
}

impl ChartAxis {
    /// Creates a new category axis with default settings.
    ///
    /// The axis is placed at the bottom with ID 1 and crosses value axis 2.
    pub fn new_category() -> Self {
        Self {
            id: 1,
            axis_type: "catAx".to_string(),
            position: "b".to_string(),
            title: None,
            min: None,
            max: None,
            major_unit: None,
            minor_unit: None,
            major_gridlines: false,
            minor_gridlines: false,
            crosses_ax: Some(2),
            num_fmt: None,
            deleted: false,
        }
    }

    /// Creates a new value axis with default settings.
    ///
    /// The axis is placed on the left with ID 2 and crosses category axis 1.
    pub fn new_value() -> Self {
        Self {
            id: 2,
            axis_type: "valAx".to_string(),
            position: "l".to_string(),
            title: None,
            min: None,
            max: None,
            major_unit: None,
            minor_unit: None,
            major_gridlines: true,
            minor_gridlines: false,
            crosses_ax: Some(1),
            num_fmt: None,
            deleted: false,
        }
    }

    /// Creates a value axis positioned at the bottom (for scatter / bubble charts
    /// which use two value axes instead of category + value).
    ///
    /// The axis has ID 1 and crosses value axis 2.
    pub fn new_value_bottom() -> Self {
        Self {
            id: 1,
            axis_type: "valAx".to_string(),
            position: "b".to_string(),
            title: None,
            min: None,
            max: None,
            major_unit: None,
            minor_unit: None,
            major_gridlines: false,
            minor_gridlines: false,
            crosses_ax: Some(2),
            num_fmt: None,
            deleted: false,
        }
    }

    /// Returns the axis ID.
    pub fn id(&self) -> u32 {
        self.id
    }

    /// Sets the axis ID.
    pub fn set_id(&mut self, id: u32) -> &mut Self {
        self.id = id;
        self
    }

    /// Returns the axis type (e.g. `"catAx"`, `"valAx"`).
    pub fn axis_type(&self) -> &str {
        &self.axis_type
    }

    /// Sets the axis type.
    pub fn set_axis_type(&mut self, axis_type: impl Into<String>) -> &mut Self {
        self.axis_type = axis_type.into();
        self
    }

    /// Returns the axis position (`"b"`, `"l"`, `"t"`, or `"r"`).
    pub fn position(&self) -> &str {
        &self.position
    }

    /// Sets the axis position.
    pub fn set_position(&mut self, position: impl Into<String>) -> &mut Self {
        self.position = position.into();
        self
    }

    /// Returns the axis title.
    pub fn title(&self) -> Option<&str> {
        self.title.as_deref()
    }

    /// Sets the axis title.
    pub fn set_title(&mut self, title: impl Into<String>) -> &mut Self {
        self.title = Some(title.into());
        self
    }

    /// Clears the axis title.
    pub fn clear_title(&mut self) -> &mut Self {
        self.title = None;
        self
    }

    /// Returns the minimum scale value.
    pub fn min(&self) -> Option<f64> {
        self.min
    }

    /// Sets the minimum scale value.
    pub fn set_min(&mut self, min: f64) -> &mut Self {
        self.min = Some(min);
        self
    }

    /// Clears the minimum scale value (auto-scaling).
    pub fn clear_min(&mut self) -> &mut Self {
        self.min = None;
        self
    }

    /// Returns the maximum scale value.
    pub fn max(&self) -> Option<f64> {
        self.max
    }

    /// Sets the maximum scale value.
    pub fn set_max(&mut self, max: f64) -> &mut Self {
        self.max = Some(max);
        self
    }

    /// Clears the maximum scale value (auto-scaling).
    pub fn clear_max(&mut self) -> &mut Self {
        self.max = None;
        self
    }

    /// Returns the major tick/gridline interval.
    pub fn major_unit(&self) -> Option<f64> {
        self.major_unit
    }

    /// Sets the major tick/gridline interval.
    pub fn set_major_unit(&mut self, unit: f64) -> &mut Self {
        self.major_unit = Some(unit);
        self
    }

    /// Clears the major unit (auto interval).
    pub fn clear_major_unit(&mut self) -> &mut Self {
        self.major_unit = None;
        self
    }

    /// Returns the minor tick/gridline interval.
    pub fn minor_unit(&self) -> Option<f64> {
        self.minor_unit
    }

    /// Sets the minor tick/gridline interval.
    pub fn set_minor_unit(&mut self, unit: f64) -> &mut Self {
        self.minor_unit = Some(unit);
        self
    }

    /// Clears the minor unit (auto interval).
    pub fn clear_minor_unit(&mut self) -> &mut Self {
        self.minor_unit = None;
        self
    }

    /// Returns whether major gridlines are displayed.
    pub fn major_gridlines(&self) -> bool {
        self.major_gridlines
    }

    /// Sets whether major gridlines are displayed.
    pub fn set_major_gridlines(&mut self, show: bool) -> &mut Self {
        self.major_gridlines = show;
        self
    }

    /// Returns whether minor gridlines are displayed.
    pub fn minor_gridlines(&self) -> bool {
        self.minor_gridlines
    }

    /// Sets whether minor gridlines are displayed.
    pub fn set_minor_gridlines(&mut self, show: bool) -> &mut Self {
        self.minor_gridlines = show;
        self
    }

    /// Returns the ID of the crossing axis.
    pub fn crosses_ax(&self) -> Option<u32> {
        self.crosses_ax
    }

    /// Sets the ID of the crossing axis.
    pub fn set_crosses_ax(&mut self, ax_id: u32) -> &mut Self {
        self.crosses_ax = Some(ax_id);
        self
    }

    /// Clears the crossing axis reference.
    pub fn clear_crosses_ax(&mut self) -> &mut Self {
        self.crosses_ax = None;
        self
    }

    /// Returns the number format string for axis labels.
    pub fn num_fmt(&self) -> Option<&str> {
        self.num_fmt.as_deref()
    }

    /// Sets the number format string.
    pub fn set_num_fmt(&mut self, fmt: impl Into<String>) -> &mut Self {
        self.num_fmt = Some(fmt.into());
        self
    }

    /// Clears the number format.
    pub fn clear_num_fmt(&mut self) -> &mut Self {
        self.num_fmt = None;
        self
    }

    /// Returns whether the axis is hidden.
    pub fn deleted(&self) -> bool {
        self.deleted
    }

    /// Sets whether the axis is hidden.
    pub fn set_deleted(&mut self, deleted: bool) -> &mut Self {
        self.deleted = deleted;
        self
    }
}

// ===== ChartLegend =====

/// Chart legend settings.
///
/// Maps to the `<c:legend>` element in chart XML. The position is one
/// of `"b"` (bottom), `"t"` (top), `"l"` (left), `"r"` (right), or
/// `"tr"` (top-right).
#[derive(Debug, Clone, PartialEq)]
pub struct ChartLegend {
    /// Legend position: `"b"`, `"t"`, `"l"`, `"r"`, or `"tr"`.
    position: String,
    /// Whether the legend overlaps the plot area.
    overlay: bool,
}

impl ChartLegend {
    /// Creates a new legend at the bottom of the chart.
    pub fn new() -> Self {
        Self {
            position: "b".to_string(),
            overlay: false,
        }
    }

    /// Returns the legend position.
    pub fn position(&self) -> &str {
        &self.position
    }

    /// Sets the legend position (`"b"`, `"t"`, `"l"`, `"r"`, or `"tr"`).
    pub fn set_position(&mut self, position: impl Into<String>) -> &mut Self {
        self.position = position.into();
        self
    }

    /// Returns whether the legend overlaps the plot area.
    pub fn overlay(&self) -> bool {
        self.overlay
    }

    /// Sets whether the legend overlaps the plot area.
    pub fn set_overlay(&mut self, overlay: bool) -> &mut Self {
        self.overlay = overlay;
        self
    }
}

impl Default for ChartLegend {
    fn default() -> Self {
        Self::new()
    }
}

// ===== Chart =====

/// A chart embedded in a worksheet.
///
/// This is the top-level chart domain object, combining the chart type,
/// series data, axes, legend, and anchor positioning. It maps to the
/// `<c:chartSpace>` root element in a chart XML part.
///
/// # Example
///
/// ```
/// use offidized_xlsx::chart::*;
///
/// // Bar charts come with default axes, grouping (clustered), and
/// // bar direction (column) — just set a title and add series:
/// let mut chart = Chart::new(ChartType::Bar)
///     .with_title("Monthly Sales");
///
/// let mut series = ChartSeries::new(0, 0);
/// series
///     .set_name("Revenue")
///     .set_categories(ChartDataRef::from_formula("Sheet1!$A$2:$A$5"))
///     .set_values(ChartDataRef::from_formula("Sheet1!$B$2:$B$5"));
///
/// chart.add_series(series);
/// assert_eq!(chart.axes().len(), 2); // cat + val axes auto-created
/// ```
#[derive(Debug, Clone, PartialEq)]
pub struct Chart {
    /// The primary chart type.
    chart_type: ChartType,
    /// Bar/column direction (only meaningful for bar charts).
    bar_direction: Option<BarDirection>,
    /// How series are grouped (stacked, clustered, etc.).
    grouping: Option<ChartGrouping>,
    /// Chart title text.
    title: Option<String>,
    /// Data series in this chart.
    series: Vec<ChartSeries>,
    /// Chart axes.
    axes: Vec<ChartAxis>,
    /// Legend settings.
    legend: Option<ChartLegend>,
    /// Whether each data point gets a different color.
    vary_colors: bool,
    /// Anchor: starting column (zero-based).
    from_col: u32,
    /// Anchor: starting row (zero-based).
    from_row: u32,
    /// Anchor: ending column (zero-based).
    to_col: u32,
    /// Anchor: ending row (zero-based).
    to_row: u32,
    /// Anchor: starting column offset in EMUs.
    from_col_off: i64,
    /// Anchor: starting row offset in EMUs.
    from_row_off: i64,
    /// Anchor: ending column offset in EMUs.
    to_col_off: i64,
    /// Anchor: ending row offset in EMUs.
    to_row_off: i64,
    /// Display name of the chart.
    name: Option<String>,
    /// Extent width in EMUs (from `<ext cx="..."/>` in one-cell anchors).
    extent_cx: Option<i64>,
    /// Extent height in EMUs (from `<ext cy="..."/>` in one-cell anchors).
    extent_cy: Option<i64>,
}

impl Chart {
    /// Creates a new chart of the specified type with default settings.
    ///
    /// The chart starts with no series, no axes, and a default anchor
    /// spanning columns 0-9 and rows 0-14.
    pub fn new(chart_type: ChartType) -> Self {
        let (grouping, bar_direction, vary_colors, axes) = Self::defaults_for_type(chart_type);
        Self {
            chart_type,
            bar_direction,
            grouping,
            title: None,
            series: Vec::new(),
            axes,
            legend: None,
            vary_colors,
            from_col: 0,
            from_row: 0,
            to_col: 9,
            to_row: 14,
            from_col_off: 0,
            from_row_off: 0,
            to_col_off: 0,
            to_row_off: 0,
            name: None,
            extent_cx: None,
            extent_cy: None,
        }
    }

    /// Returns sensible defaults for the given chart type:
    /// `(grouping, bar_direction, vary_colors, axes)`.
    fn defaults_for_type(
        chart_type: ChartType,
    ) -> (
        Option<ChartGrouping>,
        Option<BarDirection>,
        bool,
        Vec<ChartAxis>,
    ) {
        match chart_type {
            // Pie and doughnut: no grouping, no axes, vary colors on
            ChartType::Pie | ChartType::Doughnut => (None, None, true, Vec::new()),
            // Bar: clustered, default column direction, cat+val axes
            ChartType::Bar => (
                Some(ChartGrouping::Clustered),
                Some(BarDirection::Column),
                false,
                vec![ChartAxis::new_category(), ChartAxis::new_value()],
            ),
            // Scatter and bubble: standard, no bar dir, val+val axes
            ChartType::Scatter | ChartType::Bubble => (
                Some(ChartGrouping::Standard),
                None,
                false,
                vec![ChartAxis::new_value_bottom(), ChartAxis::new_value()],
            ),
            // Stock: standard, no bar dir, cat+val axes
            // Surface: standard, no bar dir, cat+val axes
            // Combo: standard, no bar dir, cat+val axes
            // Line, Area, Radar: standard, no bar dir, cat+val axes
            _ => (
                Some(ChartGrouping::Standard),
                None,
                false,
                vec![ChartAxis::new_category(), ChartAxis::new_value()],
            ),
        }
    }

    /// Returns the chart type.
    pub fn chart_type(&self) -> ChartType {
        self.chart_type
    }

    /// Sets the chart type, updating grouping and bar direction defaults to
    /// match the new type.  Existing axes are preserved — call
    /// [`clear_axes`](Self::clear_axes) first if you need to reset them.
    pub fn set_chart_type(&mut self, chart_type: ChartType) -> &mut Self {
        let (grouping, bar_direction, vary_colors, _axes) = Self::defaults_for_type(chart_type);
        self.chart_type = chart_type;
        self.grouping = grouping;
        self.bar_direction = bar_direction;
        self.vary_colors = vary_colors;
        self
    }

    /// Builder method: sets the chart title and returns self.
    pub fn with_title(mut self, title: impl Into<String>) -> Self {
        self.title = Some(title.into());
        self
    }

    /// Builder method: sets the bar direction and returns self.
    pub fn with_bar_direction(mut self, direction: BarDirection) -> Self {
        self.bar_direction = Some(direction);
        self
    }

    /// Builder method: sets the grouping and returns self.
    pub fn with_grouping(mut self, grouping: ChartGrouping) -> Self {
        self.grouping = Some(grouping);
        self
    }

    /// Returns the bar direction, if set.
    pub fn bar_direction(&self) -> Option<BarDirection> {
        self.bar_direction
    }

    /// Sets the bar direction.
    pub fn set_bar_direction(&mut self, direction: BarDirection) -> &mut Self {
        self.bar_direction = Some(direction);
        self
    }

    /// Clears the bar direction.
    pub fn clear_bar_direction(&mut self) -> &mut Self {
        self.bar_direction = None;
        self
    }

    /// Returns the chart grouping, if set.
    pub fn grouping(&self) -> Option<ChartGrouping> {
        self.grouping
    }

    /// Sets the chart grouping.
    pub fn set_grouping(&mut self, grouping: ChartGrouping) -> &mut Self {
        self.grouping = Some(grouping);
        self
    }

    /// Clears the chart grouping.
    pub fn clear_grouping(&mut self) -> &mut Self {
        self.grouping = None;
        self
    }

    /// Returns the chart title.
    pub fn title(&self) -> Option<&str> {
        self.title.as_deref()
    }

    /// Sets the chart title.
    pub fn set_title(&mut self, title: impl Into<String>) -> &mut Self {
        self.title = Some(title.into());
        self
    }

    /// Clears the chart title.
    pub fn clear_title(&mut self) -> &mut Self {
        self.title = None;
        self
    }

    /// Returns the data series in this chart.
    pub fn series(&self) -> &[ChartSeries] {
        &self.series
    }

    /// Returns a mutable reference to the data series.
    pub fn series_mut(&mut self) -> &mut Vec<ChartSeries> {
        &mut self.series
    }

    /// Adds a data series to the chart.
    pub fn add_series(&mut self, series: ChartSeries) -> &mut Self {
        self.series.push(series);
        self
    }

    /// Returns the chart axes.
    pub fn axes(&self) -> &[ChartAxis] {
        &self.axes
    }

    /// Returns a mutable reference to the chart axes.
    pub fn axes_mut(&mut self) -> &mut Vec<ChartAxis> {
        &mut self.axes
    }

    /// Adds an axis to the chart.
    pub fn add_axis(&mut self, axis: ChartAxis) -> &mut Self {
        self.axes.push(axis);
        self
    }

    /// Removes all axes from this chart.
    pub fn clear_axes(&mut self) -> &mut Self {
        self.axes.clear();
        self
    }

    /// Returns the legend settings, if present.
    pub fn legend(&self) -> Option<&ChartLegend> {
        self.legend.as_ref()
    }

    /// Returns a mutable reference to the legend.
    pub fn legend_mut(&mut self) -> Option<&mut ChartLegend> {
        self.legend.as_mut()
    }

    /// Sets the legend.
    pub fn set_legend(&mut self, legend: ChartLegend) -> &mut Self {
        self.legend = Some(legend);
        self
    }

    /// Clears the legend.
    pub fn clear_legend(&mut self) -> &mut Self {
        self.legend = None;
        self
    }

    /// Returns whether each data point gets a different color.
    pub fn vary_colors(&self) -> bool {
        self.vary_colors
    }

    /// Sets whether each data point gets a different color.
    pub fn set_vary_colors(&mut self, vary: bool) -> &mut Self {
        self.vary_colors = vary;
        self
    }

    /// Returns the anchor starting column (zero-based).
    pub fn from_col(&self) -> u32 {
        self.from_col
    }

    /// Sets the anchor starting column.
    pub fn set_from_col(&mut self, col: u32) -> &mut Self {
        self.from_col = col;
        self
    }

    /// Returns the anchor starting row (zero-based).
    pub fn from_row(&self) -> u32 {
        self.from_row
    }

    /// Sets the anchor starting row.
    pub fn set_from_row(&mut self, row: u32) -> &mut Self {
        self.from_row = row;
        self
    }

    /// Returns the anchor ending column (zero-based).
    pub fn to_col(&self) -> u32 {
        self.to_col
    }

    /// Sets the anchor ending column.
    pub fn set_to_col(&mut self, col: u32) -> &mut Self {
        self.to_col = col;
        self
    }

    /// Returns the anchor ending row (zero-based).
    pub fn to_row(&self) -> u32 {
        self.to_row
    }

    /// Sets the anchor ending row.
    pub fn set_to_row(&mut self, row: u32) -> &mut Self {
        self.to_row = row;
        self
    }

    /// Returns the anchor starting column offset in EMUs.
    pub fn from_col_off(&self) -> i64 {
        self.from_col_off
    }

    /// Sets the anchor starting column offset in EMUs.
    pub fn set_from_col_off(&mut self, off: i64) -> &mut Self {
        self.from_col_off = off;
        self
    }

    /// Returns the anchor starting row offset in EMUs.
    pub fn from_row_off(&self) -> i64 {
        self.from_row_off
    }

    /// Sets the anchor starting row offset in EMUs.
    pub fn set_from_row_off(&mut self, off: i64) -> &mut Self {
        self.from_row_off = off;
        self
    }

    /// Returns the anchor ending column offset in EMUs.
    pub fn to_col_off(&self) -> i64 {
        self.to_col_off
    }

    /// Sets the anchor ending column offset in EMUs.
    pub fn set_to_col_off(&mut self, off: i64) -> &mut Self {
        self.to_col_off = off;
        self
    }

    /// Returns the anchor ending row offset in EMUs.
    pub fn to_row_off(&self) -> i64 {
        self.to_row_off
    }

    /// Sets the anchor ending row offset in EMUs.
    pub fn set_to_row_off(&mut self, off: i64) -> &mut Self {
        self.to_row_off = off;
        self
    }

    /// Sets the full two-cell anchor in one call.
    pub fn set_anchor(
        &mut self,
        from_col: u32,
        from_row: u32,
        to_col: u32,
        to_row: u32,
    ) -> &mut Self {
        self.from_col = from_col;
        self.from_row = from_row;
        self.to_col = to_col;
        self.to_row = to_row;
        self
    }

    /// Returns the display name of the chart.
    pub fn name(&self) -> Option<&str> {
        self.name.as_deref()
    }

    /// Sets the display name.
    pub fn set_name(&mut self, name: impl Into<String>) -> &mut Self {
        self.name = Some(name.into());
        self
    }

    /// Clears the display name.
    pub fn clear_name(&mut self) -> &mut Self {
        self.name = None;
        self
    }

    /// Returns the extent width in EMUs, if set (from one-cell anchor `<ext cx="..."/>`).
    pub fn extent_cx(&self) -> Option<i64> {
        self.extent_cx
    }

    /// Sets the extent width in EMUs.
    pub fn set_extent_cx(&mut self, cx: i64) -> &mut Self {
        self.extent_cx = Some(cx);
        self
    }

    /// Clears the extent width.
    pub fn clear_extent_cx(&mut self) -> &mut Self {
        self.extent_cx = None;
        self
    }

    /// Returns the extent height in EMUs, if set (from one-cell anchor `<ext cy="..."/>`).
    pub fn extent_cy(&self) -> Option<i64> {
        self.extent_cy
    }

    /// Sets the extent height in EMUs.
    pub fn set_extent_cy(&mut self, cy: i64) -> &mut Self {
        self.extent_cy = Some(cy);
        self
    }

    /// Clears the extent height.
    pub fn clear_extent_cy(&mut self) -> &mut Self {
        self.extent_cy = None;
        self
    }
}

// ===== Tests =====

#[cfg(test)]
mod tests {
    use super::*;

    // ---- ChartType ----

    #[test]
    fn chart_type_as_str_roundtrips() {
        let types = [
            (ChartType::Bar, "barChart"),
            (ChartType::Line, "lineChart"),
            (ChartType::Pie, "pieChart"),
            (ChartType::Area, "areaChart"),
            (ChartType::Scatter, "scatterChart"),
            (ChartType::Doughnut, "doughnutChart"),
            (ChartType::Radar, "radarChart"),
            (ChartType::Bubble, "bubbleChart"),
            (ChartType::Stock, "stockChart"),
            (ChartType::Surface, "surfaceChart"),
            (ChartType::Combo, "comboChart"),
        ];

        for (variant, expected) in &types {
            assert_eq!(variant.as_str(), *expected);
            assert_eq!(ChartType::from_xml_value(expected), Some(*variant));
        }
    }

    #[test]
    fn chart_type_from_str_handles_3d_variants() {
        assert_eq!(
            ChartType::from_xml_value("bar3DChart"),
            Some(ChartType::Bar)
        );
        assert_eq!(
            ChartType::from_xml_value("line3DChart"),
            Some(ChartType::Line)
        );
        assert_eq!(
            ChartType::from_xml_value("pie3DChart"),
            Some(ChartType::Pie)
        );
        assert_eq!(
            ChartType::from_xml_value("area3DChart"),
            Some(ChartType::Area)
        );
        assert_eq!(
            ChartType::from_xml_value("surface3DChart"),
            Some(ChartType::Surface)
        );
    }

    #[test]
    fn chart_type_from_str_returns_none_for_unknown() {
        assert_eq!(ChartType::from_xml_value("unknownChart"), None);
        assert_eq!(ChartType::from_xml_value(""), None);
    }

    // ---- BarDirection ----

    #[test]
    fn bar_direction_roundtrips() {
        assert_eq!(BarDirection::Column.as_str(), "col");
        assert_eq!(BarDirection::Bar.as_str(), "bar");
        assert_eq!(
            BarDirection::from_xml_value("col"),
            Some(BarDirection::Column)
        );
        assert_eq!(BarDirection::from_xml_value("bar"), Some(BarDirection::Bar));
        assert_eq!(BarDirection::from_xml_value("invalid"), None);
    }

    // ---- ChartGrouping ----

    #[test]
    fn chart_grouping_roundtrips() {
        let groupings = [
            (ChartGrouping::Standard, "standard"),
            (ChartGrouping::Stacked, "stacked"),
            (ChartGrouping::PercentStacked, "percentStacked"),
            (ChartGrouping::Clustered, "clustered"),
        ];

        for (variant, expected) in &groupings {
            assert_eq!(variant.as_str(), *expected);
            assert_eq!(ChartGrouping::from_xml_value(expected), Some(*variant));
        }
        assert_eq!(ChartGrouping::from_xml_value("invalid"), None);
    }

    // ---- ChartDataRef ----

    #[test]
    fn chart_data_ref_new_is_empty() {
        let data = ChartDataRef::new();
        assert!(data.formula().is_none());
        assert!(data.num_values().is_empty());
        assert!(data.str_values().is_empty());
    }

    #[test]
    fn chart_data_ref_from_formula() {
        let data = ChartDataRef::from_formula("Sheet1!$B$2:$B$5");
        assert_eq!(data.formula(), Some("Sheet1!$B$2:$B$5"));
        assert!(data.num_values().is_empty());
    }

    #[test]
    fn chart_data_ref_setters() {
        let mut data = ChartDataRef::new();
        data.set_formula("Sheet1!$A$1:$A$3")
            .set_num_values(vec![Some(1.0), Some(2.0), None])
            .set_str_values(vec!["a".to_string(), "b".to_string()]);

        assert_eq!(data.formula(), Some("Sheet1!$A$1:$A$3"));
        assert_eq!(data.num_values(), &[Some(1.0), Some(2.0), None]);
        assert_eq!(data.str_values(), &["a", "b"]);

        data.clear_formula().clear_num_values().clear_str_values();
        assert!(data.formula().is_none());
        assert!(data.num_values().is_empty());
        assert!(data.str_values().is_empty());
    }

    #[test]
    fn chart_data_ref_default() {
        let data = ChartDataRef::default();
        assert!(data.formula().is_none());
        assert!(data.num_values().is_empty());
    }

    // ---- ChartSeries ----

    #[test]
    fn chart_series_constructor() {
        let series = ChartSeries::new(0, 0);
        assert_eq!(series.idx(), 0);
        assert_eq!(series.order(), 0);
        assert!(series.name().is_none());
        assert!(series.categories().is_none());
        assert!(series.values().is_none());
        assert!(series.x_values().is_none());
        assert!(series.bubble_sizes().is_none());
        assert!(series.fill_color().is_none());
        assert!(series.line_color().is_none());
        assert!(series.series_type().is_none());
    }

    #[test]
    fn chart_series_setters() {
        let mut series = ChartSeries::new(1, 2);
        series
            .set_name("Revenue")
            .set_name_ref("Sheet1!$B$1")
            .set_categories(ChartDataRef::from_formula("Sheet1!$A$2:$A$5"))
            .set_values(ChartDataRef::from_formula("Sheet1!$B$2:$B$5"))
            .set_fill_color("FF4472C4")
            .set_line_color("FF2F5597")
            .set_series_type(ChartType::Line);

        assert_eq!(series.idx(), 1);
        assert_eq!(series.order(), 2);
        assert_eq!(series.name(), Some("Revenue"));
        assert_eq!(series.name_ref(), Some("Sheet1!$B$1"));
        assert_eq!(
            series.categories().unwrap().formula(),
            Some("Sheet1!$A$2:$A$5")
        );
        assert_eq!(series.values().unwrap().formula(), Some("Sheet1!$B$2:$B$5"));
        assert_eq!(series.fill_color(), Some("FF4472C4"));
        assert_eq!(series.line_color(), Some("FF2F5597"));
        assert_eq!(series.series_type(), Some(ChartType::Line));
    }

    #[test]
    fn chart_series_clear_methods() {
        let mut series = ChartSeries::new(0, 0);
        series
            .set_name("Test")
            .set_name_ref("Sheet1!$A$1")
            .set_categories(ChartDataRef::new())
            .set_values(ChartDataRef::new())
            .set_x_values(ChartDataRef::new())
            .set_bubble_sizes(ChartDataRef::new())
            .set_fill_color("red")
            .set_line_color("blue")
            .set_series_type(ChartType::Bar);

        series
            .clear_name()
            .clear_name_ref()
            .clear_categories()
            .clear_values()
            .clear_x_values()
            .clear_bubble_sizes()
            .clear_fill_color()
            .clear_line_color()
            .clear_series_type();

        assert!(series.name().is_none());
        assert!(series.name_ref().is_none());
        assert!(series.categories().is_none());
        assert!(series.values().is_none());
        assert!(series.x_values().is_none());
        assert!(series.bubble_sizes().is_none());
        assert!(series.fill_color().is_none());
        assert!(series.line_color().is_none());
        assert!(series.series_type().is_none());
    }

    #[test]
    fn chart_series_scatter_data() {
        let mut series = ChartSeries::new(0, 0);
        series
            .set_x_values(ChartDataRef::from_formula("Sheet1!$A$2:$A$10"))
            .set_values(ChartDataRef::from_formula("Sheet1!$B$2:$B$10"))
            .set_bubble_sizes(ChartDataRef::from_formula("Sheet1!$C$2:$C$10"));

        assert_eq!(
            series.x_values().unwrap().formula(),
            Some("Sheet1!$A$2:$A$10")
        );
        assert_eq!(
            series.bubble_sizes().unwrap().formula(),
            Some("Sheet1!$C$2:$C$10")
        );
    }

    // ---- ChartAxis ----

    #[test]
    fn chart_axis_category_defaults() {
        let axis = ChartAxis::new_category();
        assert_eq!(axis.id(), 1);
        assert_eq!(axis.axis_type(), "catAx");
        assert_eq!(axis.position(), "b");
        assert!(axis.title().is_none());
        assert!(axis.min().is_none());
        assert!(axis.max().is_none());
        assert!(!axis.major_gridlines());
        assert!(!axis.minor_gridlines());
        assert_eq!(axis.crosses_ax(), Some(2));
        assert!(!axis.deleted());
    }

    #[test]
    fn chart_axis_value_defaults() {
        let axis = ChartAxis::new_value();
        assert_eq!(axis.id(), 2);
        assert_eq!(axis.axis_type(), "valAx");
        assert_eq!(axis.position(), "l");
        assert!(axis.major_gridlines());
        assert_eq!(axis.crosses_ax(), Some(1));
    }

    #[test]
    fn chart_axis_setters() {
        let mut axis = ChartAxis::new_value();
        axis.set_id(10)
            .set_axis_type("dateAx")
            .set_position("r")
            .set_title("Amount ($)")
            .set_min(0.0)
            .set_max(1000.0)
            .set_major_unit(100.0)
            .set_minor_unit(25.0)
            .set_major_gridlines(true)
            .set_minor_gridlines(true)
            .set_crosses_ax(5)
            .set_num_fmt("#,##0")
            .set_deleted(true);

        assert_eq!(axis.id(), 10);
        assert_eq!(axis.axis_type(), "dateAx");
        assert_eq!(axis.position(), "r");
        assert_eq!(axis.title(), Some("Amount ($)"));
        assert_eq!(axis.min(), Some(0.0));
        assert_eq!(axis.max(), Some(1000.0));
        assert_eq!(axis.major_unit(), Some(100.0));
        assert_eq!(axis.minor_unit(), Some(25.0));
        assert!(axis.major_gridlines());
        assert!(axis.minor_gridlines());
        assert_eq!(axis.crosses_ax(), Some(5));
        assert_eq!(axis.num_fmt(), Some("#,##0"));
        assert!(axis.deleted());
    }

    #[test]
    fn chart_axis_clear_methods() {
        let mut axis = ChartAxis::new_value();
        axis.set_title("T")
            .set_min(0.0)
            .set_max(100.0)
            .set_major_unit(10.0)
            .set_minor_unit(5.0)
            .set_crosses_ax(3)
            .set_num_fmt("0.00");

        axis.clear_title()
            .clear_min()
            .clear_max()
            .clear_major_unit()
            .clear_minor_unit()
            .clear_crosses_ax()
            .clear_num_fmt();

        assert!(axis.title().is_none());
        assert!(axis.min().is_none());
        assert!(axis.max().is_none());
        assert!(axis.major_unit().is_none());
        assert!(axis.minor_unit().is_none());
        assert!(axis.crosses_ax().is_none());
        assert!(axis.num_fmt().is_none());
    }

    // ---- ChartLegend ----

    #[test]
    fn chart_legend_defaults() {
        let legend = ChartLegend::new();
        assert_eq!(legend.position(), "b");
        assert!(!legend.overlay());
    }

    #[test]
    fn chart_legend_setters() {
        let mut legend = ChartLegend::new();
        legend.set_position("r").set_overlay(true);
        assert_eq!(legend.position(), "r");
        assert!(legend.overlay());
    }

    #[test]
    fn chart_legend_default_trait() {
        let legend = ChartLegend::default();
        assert_eq!(legend.position(), "b");
    }

    // ---- Chart ----

    #[test]
    fn chart_constructor() {
        let chart = Chart::new(ChartType::Bar);
        assert_eq!(chart.chart_type(), ChartType::Bar);
        assert!(chart.title().is_none());
        assert!(chart.series().is_empty());
        // Bar charts get default cat+val axes
        assert_eq!(chart.axes().len(), 2);
        assert_eq!(chart.axes()[0].axis_type(), "catAx");
        assert_eq!(chart.axes()[1].axis_type(), "valAx");
        // Bar charts get default grouping and bar direction
        assert_eq!(chart.grouping(), Some(ChartGrouping::Clustered));
        assert_eq!(chart.bar_direction(), Some(BarDirection::Column));
        assert!(chart.legend().is_none());
        assert!(!chart.vary_colors());
        assert_eq!(chart.from_col(), 0);
        assert_eq!(chart.from_row(), 0);
        assert_eq!(chart.to_col(), 9);
        assert_eq!(chart.to_row(), 14);
        assert!(chart.name().is_none());

        // Pie charts get no axes, no grouping, vary_colors on
        let pie = Chart::new(ChartType::Pie);
        assert!(pie.axes().is_empty());
        assert!(pie.grouping().is_none());
        assert!(pie.vary_colors());

        // Line charts get standard grouping + axes
        let line = Chart::new(ChartType::Line);
        assert_eq!(line.grouping(), Some(ChartGrouping::Standard));
        assert_eq!(line.axes().len(), 2);

        // Scatter charts get two value axes
        let scatter = Chart::new(ChartType::Scatter);
        assert_eq!(scatter.axes().len(), 2);
        assert_eq!(scatter.axes()[0].axis_type(), "valAx");
        assert_eq!(scatter.axes()[1].axis_type(), "valAx");
    }

    #[test]
    fn chart_builder_methods() {
        let chart = Chart::new(ChartType::Bar)
            .with_title("Sales")
            .with_bar_direction(BarDirection::Column)
            .with_grouping(ChartGrouping::Clustered);

        assert_eq!(chart.title(), Some("Sales"));
        assert_eq!(chart.bar_direction(), Some(BarDirection::Column));
        assert_eq!(chart.grouping(), Some(ChartGrouping::Clustered));
    }

    #[test]
    fn chart_add_series() {
        let mut chart = Chart::new(ChartType::Line);

        let mut s1 = ChartSeries::new(0, 0);
        s1.set_name("Series A");
        let mut s2 = ChartSeries::new(1, 1);
        s2.set_name("Series B");

        chart.add_series(s1).add_series(s2);

        assert_eq!(chart.series().len(), 2);
        assert_eq!(chart.series()[0].name(), Some("Series A"));
        assert_eq!(chart.series()[1].name(), Some("Series B"));
    }

    #[test]
    fn chart_add_axes() {
        // Start with pie (no default axes) and manually add
        let mut chart = Chart::new(ChartType::Pie);
        assert!(chart.axes().is_empty());
        chart
            .add_axis(ChartAxis::new_category())
            .add_axis(ChartAxis::new_value());

        assert_eq!(chart.axes().len(), 2);
        assert_eq!(chart.axes()[0].axis_type(), "catAx");
        assert_eq!(chart.axes()[1].axis_type(), "valAx");

        // clear_axes works
        chart.clear_axes();
        assert!(chart.axes().is_empty());
    }

    #[test]
    fn chart_setters() {
        let mut chart = Chart::new(ChartType::Pie);
        chart
            .set_chart_type(ChartType::Doughnut)
            .set_title("Donut")
            .set_vary_colors(true)
            .set_legend(ChartLegend::new())
            .set_name("Chart 1")
            .set_anchor(2, 3, 12, 18);

        assert_eq!(chart.chart_type(), ChartType::Doughnut);
        assert_eq!(chart.title(), Some("Donut"));
        assert!(chart.vary_colors());
        assert!(chart.legend().is_some());
        assert_eq!(chart.name(), Some("Chart 1"));
        assert_eq!(chart.from_col(), 2);
        assert_eq!(chart.from_row(), 3);
        assert_eq!(chart.to_col(), 12);
        assert_eq!(chart.to_row(), 18);
    }

    #[test]
    fn chart_clear_methods() {
        let mut chart = Chart::new(ChartType::Bar)
            .with_title("T")
            .with_bar_direction(BarDirection::Bar)
            .with_grouping(ChartGrouping::Stacked);

        chart.set_legend(ChartLegend::new()).set_name("N");

        chart
            .clear_title()
            .clear_bar_direction()
            .clear_grouping()
            .clear_legend()
            .clear_name();

        assert!(chart.title().is_none());
        assert!(chart.bar_direction().is_none());
        assert!(chart.grouping().is_none());
        assert!(chart.legend().is_none());
        assert!(chart.name().is_none());
    }

    #[test]
    fn chart_mutable_access() {
        let mut chart = Chart::new(ChartType::Line);
        chart.add_series(ChartSeries::new(0, 0));
        chart.add_axis(ChartAxis::new_value());
        chart.set_legend(ChartLegend::new());

        // Mutate through mutable accessors
        chart.series_mut()[0].set_name("Updated");
        chart.axes_mut()[0].set_title("Y Axis");
        chart.legend_mut().unwrap().set_position("t");

        assert_eq!(chart.series()[0].name(), Some("Updated"));
        assert_eq!(chart.axes()[0].title(), Some("Y Axis"));
        assert_eq!(chart.legend().unwrap().position(), "t");
    }

    #[test]
    fn chart_full_construction() {
        // End-to-end: build a complete bar chart
        // Bar charts now come with default axes, grouping, and bar direction
        let mut chart = Chart::new(ChartType::Bar).with_title("Monthly Revenue");

        // Defaults should already be set
        assert_eq!(chart.bar_direction(), Some(BarDirection::Column));
        assert_eq!(chart.grouping(), Some(ChartGrouping::Clustered));
        assert_eq!(chart.axes().len(), 2);

        let mut s = ChartSeries::new(0, 0);
        s.set_name("Revenue")
            .set_categories(ChartDataRef::from_formula("Sheet1!$A$2:$A$13"))
            .set_values(ChartDataRef::from_formula("Sheet1!$B$2:$B$13"))
            .set_fill_color("FF4472C4");

        chart.add_series(s);
        chart.set_legend(ChartLegend::new());
        chart.set_anchor(3, 1, 15, 20);
        chart.set_name("Chart 1");

        assert_eq!(chart.chart_type(), ChartType::Bar);
        assert_eq!(chart.title(), Some("Monthly Revenue"));
        assert_eq!(chart.series().len(), 1);
        assert_eq!(chart.axes().len(), 2);
        assert!(chart.legend().is_some());
        assert_eq!(chart.from_col(), 3);
        assert_eq!(chart.to_row(), 20);
        assert_eq!(chart.name(), Some("Chart 1"));
    }
}
