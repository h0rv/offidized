//! Sparkline types for Excel in-cell micro-charts.
//!
//! A [`SparklineGroup`] describes a set of related sparklines that share
//! formatting, type, and display options.  Each individual [`Sparkline`] maps
//! a data range to a single cell location.

/// The visual style of a sparkline.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum SparklineType {
    /// A line sparkline (default).
    #[default]
    Line,
    /// A column (bar) sparkline.
    Column,
    /// A win/loss (stacked) sparkline.
    Stacked,
}

impl SparklineType {
    /// Returns the OOXML attribute value for this sparkline type.
    pub fn as_str(&self) -> &'static str {
        match self {
            Self::Line => "line",
            Self::Column => "column",
            Self::Stacked => "stacked",
        }
    }

    /// Parses an OOXML attribute value into a [`SparklineType`].
    pub fn from_xml_value(value: &str) -> Self {
        match value {
            "column" => Self::Column,
            "stacked" => Self::Stacked,
            _ => Self::Line,
        }
    }
}

/// How a sparkline renders cells with no data.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum SparklineEmptyCells {
    /// Leave a gap in the sparkline (default).
    #[default]
    Gap,
    /// Treat empty cells as zero.
    Zero,
    /// Connect adjacent data points across the gap.
    Connect,
}

impl SparklineEmptyCells {
    /// Returns the OOXML attribute value for this setting.
    pub fn as_str(&self) -> &'static str {
        match self {
            Self::Gap => "gap",
            Self::Zero => "zero",
            Self::Connect => "connect",
        }
    }

    /// Parses an OOXML attribute value into a [`SparklineEmptyCells`].
    pub fn from_xml_value(value: &str) -> Self {
        match value {
            "zero" => Self::Zero,
            "connect" => Self::Connect,
            _ => Self::Gap,
        }
    }
}

/// Colour settings for a sparkline group.
///
/// All colours are stored as resolved hex strings (e.g. `"FF0000"` for red).
/// Fields that are `None` inherit their colour from the workbook theme.
#[derive(Debug, Clone, PartialEq, Eq, Default)]
pub struct SparklineColors {
    /// Main series line/fill colour.
    pub series: Option<String>,
    /// Colour for negative data points.
    pub negative: Option<String>,
    /// Colour of the sparkline axis line.
    pub axis: Option<String>,
    /// Colour for data-point markers (line sparklines only).
    pub markers: Option<String>,
    /// Colour for the first data point.
    pub first: Option<String>,
    /// Colour for the last data point.
    pub last: Option<String>,
    /// Colour for the highest data point.
    pub high: Option<String>,
    /// Colour for the lowest data point.
    pub low: Option<String>,
}

impl SparklineColors {
    /// Creates a new colour set with all values unset (theme defaults).
    pub fn new() -> Self {
        Self::default()
    }

    /// Returns `true` when at least one colour has been explicitly set.
    pub fn has_any(&self) -> bool {
        self.series.is_some()
            || self.negative.is_some()
            || self.axis.is_some()
            || self.markers.is_some()
            || self.first.is_some()
            || self.last.is_some()
            || self.high.is_some()
            || self.low.is_some()
    }
}

/// A single sparkline: one cell location mapped to one data range.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Sparkline {
    /// The cell where the sparkline is rendered (e.g. `"Sheet1!A1"`).
    location: String,
    /// The data range the sparkline visualises (e.g. `"Sheet1!B1:B10"`).
    data_range: String,
}

impl Sparkline {
    /// Creates a new sparkline mapping a data range to a cell location.
    pub fn new(location: impl Into<String>, data_range: impl Into<String>) -> Self {
        Self {
            location: location.into(),
            data_range: data_range.into(),
        }
    }

    /// Returns the cell location where the sparkline is rendered.
    pub fn location(&self) -> &str {
        &self.location
    }

    /// Sets the cell location.
    pub fn set_location(&mut self, location: impl Into<String>) -> &mut Self {
        self.location = location.into();
        self
    }

    /// Returns the data range the sparkline visualises.
    pub fn data_range(&self) -> &str {
        &self.data_range
    }

    /// Sets the data range.
    pub fn set_data_range(&mut self, data_range: impl Into<String>) -> &mut Self {
        self.data_range = data_range.into();
        self
    }
}

/// Axis type for sparkline min/max axis scaling.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum SparklineAxisType {
    /// Automatic axis scaling per sparkline (default).
    #[default]
    Individual,
    /// All sparklines in the group share the same axis scale.
    Group,
    /// A manually specified axis value is used.
    Custom,
}

impl SparklineAxisType {
    /// Returns the OOXML attribute value for this axis type.
    pub fn as_str(&self) -> &'static str {
        match self {
            Self::Individual => "individual",
            Self::Group => "group",
            Self::Custom => "custom",
        }
    }

    /// Parses an OOXML attribute value into a [`SparklineAxisType`].
    pub fn from_xml_value(value: &str) -> Self {
        match value {
            "group" => Self::Group,
            "custom" => Self::Custom,
            _ => Self::Individual,
        }
    }
}

/// A group of sparklines that share formatting and display options.
#[derive(Debug, Clone, PartialEq, Default)]
pub struct SparklineGroup {
    /// The visual type of sparklines in this group.
    sparkline_type: SparklineType,
    /// Individual sparkline entries.
    sparklines: Vec<Sparkline>,
    /// Colour configuration for the group.
    colors: SparklineColors,
    /// How empty cells are rendered.
    display_empty_cells_as: SparklineEmptyCells,

    // -- Display options --
    /// Show data-point markers (line sparklines only).
    markers: bool,
    /// Highlight the highest data point.
    high_point: bool,
    /// Highlight the lowest data point.
    low_point: bool,
    /// Highlight the first data point.
    first_point: bool,
    /// Highlight the last data point.
    last_point: bool,
    /// Highlight negative data points.
    negative_points: bool,
    /// Show the horizontal axis line.
    display_x_axis: bool,
    /// Whether to display data in hidden rows/columns.
    display_hidden: bool,
    /// Whether the sparkline reads data right-to-left.
    right_to_left: bool,

    /// Line weight in points (default 0.75).
    line_weight: Option<f64>,

    // -- Axis scaling --
    /// Minimum axis type.
    min_axis_type: SparklineAxisType,
    /// Maximum axis type.
    max_axis_type: SparklineAxisType,
    /// Manual minimum axis value (when `min_axis_type` is `Custom`).
    manual_min: Option<f64>,
    /// Manual maximum axis value (when `max_axis_type` is `Custom`).
    manual_max: Option<f64>,

    /// Whether this sparkline group uses a date axis.
    date_axis: bool,
    /// The date axis data range, if `date_axis` is `true`.
    date_axis_range: Option<String>,
}

impl SparklineGroup {
    /// Creates a new sparkline group with default settings.
    pub fn new() -> Self {
        Self::default()
    }

    // -- Type --

    /// Returns the sparkline type.
    pub fn sparkline_type(&self) -> SparklineType {
        self.sparkline_type
    }

    /// Sets the sparkline type.
    pub fn set_sparkline_type(&mut self, sparkline_type: SparklineType) -> &mut Self {
        self.sparkline_type = sparkline_type;
        self
    }

    // -- Sparklines --

    /// Returns the individual sparklines in this group.
    pub fn sparklines(&self) -> &[Sparkline] {
        &self.sparklines
    }

    /// Returns a mutable reference to the sparklines vector.
    pub fn sparklines_mut(&mut self) -> &mut Vec<Sparkline> {
        &mut self.sparklines
    }

    /// Adds a sparkline to this group.
    pub fn add_sparkline(&mut self, sparkline: Sparkline) -> &mut Self {
        self.sparklines.push(sparkline);
        self
    }

    /// Removes all sparklines from this group.
    pub fn clear_sparklines(&mut self) -> &mut Self {
        self.sparklines.clear();
        self
    }

    // -- Colors --

    /// Returns the colour configuration.
    pub fn colors(&self) -> &SparklineColors {
        &self.colors
    }

    /// Returns a mutable reference to the colour configuration.
    pub fn colors_mut(&mut self) -> &mut SparklineColors {
        &mut self.colors
    }

    /// Sets the colour configuration.
    pub fn set_colors(&mut self, colors: SparklineColors) -> &mut Self {
        self.colors = colors;
        self
    }

    // -- Empty cells display --

    /// Returns how empty cells are rendered.
    pub fn display_empty_cells_as(&self) -> SparklineEmptyCells {
        self.display_empty_cells_as
    }

    /// Sets how empty cells are rendered.
    pub fn set_display_empty_cells_as(&mut self, value: SparklineEmptyCells) -> &mut Self {
        self.display_empty_cells_as = value;
        self
    }

    // -- Display option accessors --

    /// Returns whether data-point markers are shown.
    pub fn markers(&self) -> bool {
        self.markers
    }

    /// Sets whether data-point markers are shown.
    pub fn set_markers(&mut self, value: bool) -> &mut Self {
        self.markers = value;
        self
    }

    /// Returns whether the highest data point is highlighted.
    pub fn high_point(&self) -> bool {
        self.high_point
    }

    /// Sets whether the highest data point is highlighted.
    pub fn set_high_point(&mut self, value: bool) -> &mut Self {
        self.high_point = value;
        self
    }

    /// Returns whether the lowest data point is highlighted.
    pub fn low_point(&self) -> bool {
        self.low_point
    }

    /// Sets whether the lowest data point is highlighted.
    pub fn set_low_point(&mut self, value: bool) -> &mut Self {
        self.low_point = value;
        self
    }

    /// Returns whether the first data point is highlighted.
    pub fn first_point(&self) -> bool {
        self.first_point
    }

    /// Sets whether the first data point is highlighted.
    pub fn set_first_point(&mut self, value: bool) -> &mut Self {
        self.first_point = value;
        self
    }

    /// Returns whether the last data point is highlighted.
    pub fn last_point(&self) -> bool {
        self.last_point
    }

    /// Sets whether the last data point is highlighted.
    pub fn set_last_point(&mut self, value: bool) -> &mut Self {
        self.last_point = value;
        self
    }

    /// Returns whether negative data points are highlighted.
    pub fn negative_points(&self) -> bool {
        self.negative_points
    }

    /// Sets whether negative data points are highlighted.
    pub fn set_negative_points(&mut self, value: bool) -> &mut Self {
        self.negative_points = value;
        self
    }

    /// Returns whether the horizontal axis line is shown.
    pub fn display_x_axis(&self) -> bool {
        self.display_x_axis
    }

    /// Sets whether the horizontal axis line is shown.
    pub fn set_display_x_axis(&mut self, value: bool) -> &mut Self {
        self.display_x_axis = value;
        self
    }

    /// Returns whether data in hidden rows/columns is displayed.
    pub fn display_hidden(&self) -> bool {
        self.display_hidden
    }

    /// Sets whether data in hidden rows/columns is displayed.
    pub fn set_display_hidden(&mut self, value: bool) -> &mut Self {
        self.display_hidden = value;
        self
    }

    /// Returns whether the sparkline reads data right-to-left.
    pub fn right_to_left(&self) -> bool {
        self.right_to_left
    }

    /// Sets whether the sparkline reads data right-to-left.
    pub fn set_right_to_left(&mut self, value: bool) -> &mut Self {
        self.right_to_left = value;
        self
    }

    // -- Line weight --

    /// Returns the line weight in points.
    pub fn line_weight(&self) -> Option<f64> {
        self.line_weight
    }

    /// Sets the line weight in points.
    pub fn set_line_weight(&mut self, value: f64) -> &mut Self {
        self.line_weight = Some(value);
        self
    }

    // -- Axis scaling --

    /// Returns the minimum axis type.
    pub fn min_axis_type(&self) -> SparklineAxisType {
        self.min_axis_type
    }

    /// Sets the minimum axis type.
    pub fn set_min_axis_type(&mut self, value: SparklineAxisType) -> &mut Self {
        self.min_axis_type = value;
        self
    }

    /// Returns the maximum axis type.
    pub fn max_axis_type(&self) -> SparklineAxisType {
        self.max_axis_type
    }

    /// Sets the maximum axis type.
    pub fn set_max_axis_type(&mut self, value: SparklineAxisType) -> &mut Self {
        self.max_axis_type = value;
        self
    }

    /// Returns the manual minimum axis value.
    pub fn manual_min(&self) -> Option<f64> {
        self.manual_min
    }

    /// Sets the manual minimum axis value.
    pub fn set_manual_min(&mut self, value: f64) -> &mut Self {
        self.manual_min = Some(value);
        self
    }

    /// Returns the manual maximum axis value.
    pub fn manual_max(&self) -> Option<f64> {
        self.manual_max
    }

    /// Sets the manual maximum axis value.
    pub fn set_manual_max(&mut self, value: f64) -> &mut Self {
        self.manual_max = Some(value);
        self
    }

    // -- Date axis --

    /// Returns whether this group uses a date axis.
    pub fn date_axis(&self) -> bool {
        self.date_axis
    }

    /// Sets whether this group uses a date axis.
    pub fn set_date_axis(&mut self, value: bool) -> &mut Self {
        self.date_axis = value;
        self
    }

    /// Returns the date axis data range.
    pub fn date_axis_range(&self) -> Option<&str> {
        self.date_axis_range.as_deref()
    }

    /// Sets the date axis data range.
    pub fn set_date_axis_range(&mut self, range: impl Into<String>) -> &mut Self {
        self.date_axis_range = Some(range.into());
        self
    }

    /// Returns `true` when the group has no sparklines.
    pub fn is_empty(&self) -> bool {
        self.sparklines.is_empty()
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn sparkline_type_roundtrip() {
        let types = [
            SparklineType::Line,
            SparklineType::Column,
            SparklineType::Stacked,
        ];
        for t in types {
            assert_eq!(SparklineType::from_xml_value(t.as_str()), t);
        }
    }

    #[test]
    fn sparkline_type_unknown_defaults_to_line() {
        assert_eq!(SparklineType::from_xml_value("bogus"), SparklineType::Line);
    }

    #[test]
    fn sparkline_empty_cells_roundtrip() {
        let values = [
            SparklineEmptyCells::Gap,
            SparklineEmptyCells::Zero,
            SparklineEmptyCells::Connect,
        ];
        for v in values {
            assert_eq!(SparklineEmptyCells::from_xml_value(v.as_str()), v);
        }
    }

    #[test]
    fn sparkline_empty_cells_unknown_defaults_to_gap() {
        assert_eq!(
            SparklineEmptyCells::from_xml_value("bogus"),
            SparklineEmptyCells::Gap
        );
    }

    #[test]
    fn sparkline_axis_type_roundtrip() {
        let values = [
            SparklineAxisType::Individual,
            SparklineAxisType::Group,
            SparklineAxisType::Custom,
        ];
        for v in values {
            assert_eq!(SparklineAxisType::from_xml_value(v.as_str()), v);
        }
    }

    #[test]
    fn sparkline_accessors() {
        let mut sl = Sparkline::new("Sheet1!A1", "Sheet1!B1:B10");
        assert_eq!(sl.location(), "Sheet1!A1");
        assert_eq!(sl.data_range(), "Sheet1!B1:B10");

        sl.set_location("Sheet1!C1").set_data_range("Sheet1!D1:D10");
        assert_eq!(sl.location(), "Sheet1!C1");
        assert_eq!(sl.data_range(), "Sheet1!D1:D10");
    }

    #[test]
    fn sparkline_colors_has_any() {
        let empty = SparklineColors::new();
        assert!(!empty.has_any());

        let mut with_series = SparklineColors::new();
        with_series.series = Some("FF0000".to_string());
        assert!(with_series.has_any());
    }

    #[test]
    fn sparkline_group_defaults() {
        let group = SparklineGroup::new();
        assert_eq!(group.sparkline_type(), SparklineType::Line);
        assert!(group.sparklines().is_empty());
        assert!(group.is_empty());
        assert!(!group.markers());
        assert!(!group.high_point());
        assert!(!group.low_point());
        assert!(!group.first_point());
        assert!(!group.last_point());
        assert!(!group.negative_points());
        assert!(!group.display_x_axis());
        assert!(!group.display_hidden());
        assert!(!group.right_to_left());
        assert!(!group.date_axis());
        assert_eq!(group.display_empty_cells_as(), SparklineEmptyCells::Gap);
        assert_eq!(group.min_axis_type(), SparklineAxisType::Individual);
        assert_eq!(group.max_axis_type(), SparklineAxisType::Individual);
        assert!(group.line_weight().is_none());
        assert!(group.manual_min().is_none());
        assert!(group.manual_max().is_none());
    }

    #[test]
    fn sparkline_group_add_and_clear() {
        let mut group = SparklineGroup::new();
        group
            .set_sparkline_type(SparklineType::Column)
            .add_sparkline(Sparkline::new("Sheet1!A1", "Sheet1!B1:B5"))
            .add_sparkline(Sparkline::new("Sheet1!A2", "Sheet1!B6:B10"));

        assert_eq!(group.sparkline_type(), SparklineType::Column);
        assert_eq!(group.sparklines().len(), 2);
        assert!(!group.is_empty());

        group.clear_sparklines();
        assert!(group.is_empty());
    }

    #[test]
    fn sparkline_group_display_options() {
        let mut group = SparklineGroup::new();
        group
            .set_markers(true)
            .set_high_point(true)
            .set_low_point(true)
            .set_first_point(true)
            .set_last_point(true)
            .set_negative_points(true)
            .set_display_x_axis(true)
            .set_display_hidden(true)
            .set_right_to_left(true);

        assert!(group.markers());
        assert!(group.high_point());
        assert!(group.low_point());
        assert!(group.first_point());
        assert!(group.last_point());
        assert!(group.negative_points());
        assert!(group.display_x_axis());
        assert!(group.display_hidden());
        assert!(group.right_to_left());
    }

    #[test]
    fn sparkline_group_line_weight() {
        let mut group = SparklineGroup::new();
        assert!(group.line_weight().is_none());
        group.set_line_weight(1.5);
        assert_eq!(group.line_weight(), Some(1.5));
    }

    #[test]
    fn sparkline_group_axis_scaling() {
        let mut group = SparklineGroup::new();
        group
            .set_min_axis_type(SparklineAxisType::Custom)
            .set_manual_min(-10.0)
            .set_max_axis_type(SparklineAxisType::Group)
            .set_manual_max(100.0);

        assert_eq!(group.min_axis_type(), SparklineAxisType::Custom);
        assert_eq!(group.manual_min(), Some(-10.0));
        assert_eq!(group.max_axis_type(), SparklineAxisType::Group);
        assert_eq!(group.manual_max(), Some(100.0));
    }

    #[test]
    fn sparkline_group_date_axis() {
        let mut group = SparklineGroup::new();
        group
            .set_date_axis(true)
            .set_date_axis_range("Sheet1!A1:A10");

        assert!(group.date_axis());
        assert_eq!(group.date_axis_range(), Some("Sheet1!A1:A10"));
    }

    #[test]
    fn sparkline_group_colors() {
        let mut group = SparklineGroup::new();
        group.colors_mut().series = Some("0000FF".to_string());
        group.colors_mut().negative = Some("FF0000".to_string());
        assert!(group.colors().has_any());
        assert_eq!(group.colors().series.as_deref(), Some("0000FF"));
        assert_eq!(group.colors().negative.as_deref(), Some("FF0000"));
    }

    #[test]
    fn sparkline_group_empty_cells() {
        let mut group = SparklineGroup::new();
        group.set_display_empty_cells_as(SparklineEmptyCells::Zero);
        assert_eq!(group.display_empty_cells_as(), SparklineEmptyCells::Zero);
    }
}
