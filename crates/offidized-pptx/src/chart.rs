// ── Feature #13: Chart improvements ──

/// Chart type enum covering the most common chart types.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum ChartType {
    Bar,
    #[default]
    Column,
    Line,
    Pie,
    Area,
    Scatter,
    Doughnut,
    Radar,
    Other,
}

impl ChartType {
    /// Parse from the XML element local name inside `c:plotArea`.
    pub fn from_xml_element(local_name: &str) -> Option<Self> {
        match local_name {
            "barChart" => Some(Self::Column), // barDir determines Bar vs Column
            "bar3DChart" => Some(Self::Column),
            "lineChart" | "line3DChart" => Some(Self::Line),
            "pieChart" | "pie3DChart" => Some(Self::Pie),
            "areaChart" | "area3DChart" => Some(Self::Area),
            "scatterChart" => Some(Self::Scatter),
            "doughnutChart" => Some(Self::Doughnut),
            "radarChart" => Some(Self::Radar),
            _ => None,
        }
    }

    /// The main XML element name for serialization.
    pub fn to_xml_element(self) -> &'static str {
        match self {
            Self::Bar => "c:barChart",
            Self::Column => "c:barChart",
            Self::Line => "c:lineChart",
            Self::Pie => "c:pieChart",
            Self::Area => "c:areaChart",
            Self::Scatter => "c:scatterChart",
            Self::Doughnut => "c:doughnutChart",
            Self::Radar => "c:radarChart",
            Self::Other => "c:barChart",
        }
    }
}

// ── Series styling ──

/// Fill type for chart series.
#[derive(Debug, Clone, PartialEq)]
pub enum SeriesFill {
    /// Solid color fill (hex color, e.g., "FF0000" for red).
    Solid(String),
    /// Gradient fill (start color, end color).
    Gradient(String, String),
    /// Pattern fill (pattern type, foreground color, background color).
    Pattern(String, String, String),
    /// Picture fill (image data as bytes).
    Picture(Vec<u8>),
    /// No fill.
    None,
}

impl SeriesFill {
    /// Creates a solid fill with the given hex color.
    pub fn solid(color: impl Into<String>) -> Self {
        let mut color = color.into();
        if color.starts_with('#') {
            color = color[1..].to_string();
        }
        if color.len() == 8 {
            color = color[..6].to_string();
        }
        Self::Solid(color)
    }

    /// Creates a gradient fill with start and end colors.
    pub fn gradient(start: impl Into<String>, end: impl Into<String>) -> Self {
        let mut start = start.into();
        let mut end = end.into();
        if start.starts_with('#') {
            start = start[1..].to_string();
        }
        if end.starts_with('#') {
            end = end[1..].to_string();
        }
        if start.len() == 8 {
            start = start[..6].to_string();
        }
        if end.len() == 8 {
            end = end[..6].to_string();
        }
        Self::Gradient(start, end)
    }

    /// Creates a pattern fill.
    pub fn pattern(
        pattern_type: impl Into<String>,
        fg_color: impl Into<String>,
        bg_color: impl Into<String>,
    ) -> Self {
        Self::Pattern(pattern_type.into(), fg_color.into(), bg_color.into())
    }

    /// Creates a picture fill from image bytes.
    pub fn picture(image_data: Vec<u8>) -> Self {
        Self::Picture(image_data)
    }
}

/// Border styling for chart series.
#[derive(Debug, Clone, PartialEq)]
pub struct SeriesBorder {
    /// Border color as hex (e.g., "000000" for black).
    pub color: String,
    /// Border width in points.
    pub width: f64,
    /// Dash style (e.g., "solid", "dash", "dot", "dashDot").
    pub dash_style: String,
}

impl SeriesBorder {
    /// Creates a new series border with default values.
    pub fn new() -> Self {
        Self {
            color: "000000".to_string(),
            width: 1.0,
            dash_style: "solid".to_string(),
        }
    }

    /// Sets the border color (hex string with or without #).
    pub fn with_color(mut self, color: impl Into<String>) -> Self {
        let mut color = color.into();
        if color.starts_with('#') {
            color = color[1..].to_string();
        }
        if color.len() == 8 {
            color = color[..6].to_string();
        }
        self.color = color;
        self
    }

    /// Sets the border width in points.
    pub fn with_width(mut self, width: f64) -> Self {
        self.width = width.max(0.0);
        self
    }

    /// Sets the dash style.
    pub fn with_dash_style(mut self, style: impl Into<String>) -> Self {
        self.dash_style = style.into();
        self
    }
}

impl Default for SeriesBorder {
    fn default() -> Self {
        Self::new()
    }
}

/// Marker shape for line/scatter chart series.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum MarkerShape {
    Circle,
    Square,
    Diamond,
    Triangle,
    X,
    Star,
    Dot,
    Dash,
    None,
}

impl MarkerShape {
    /// Convert to XML marker symbol value.
    pub fn to_xml(self) -> &'static str {
        match self {
            Self::Circle => "circle",
            Self::Square => "square",
            Self::Diamond => "diamond",
            Self::Triangle => "triangle",
            Self::X => "x",
            Self::Star => "star",
            Self::Dot => "dot",
            Self::Dash => "dash",
            Self::None => "none",
        }
    }

    /// Parse from XML marker symbol value.
    pub fn from_xml(value: &str) -> Option<Self> {
        match value {
            "circle" => Some(Self::Circle),
            "square" => Some(Self::Square),
            "diamond" => Some(Self::Diamond),
            "triangle" => Some(Self::Triangle),
            "x" => Some(Self::X),
            "star" => Some(Self::Star),
            "dot" => Some(Self::Dot),
            "dash" => Some(Self::Dash),
            "none" => Some(Self::None),
            _ => None,
        }
    }
}

/// Line style for line/scatter chart series.
#[derive(Debug, Clone, PartialEq)]
pub struct LineStyle {
    /// Line color as hex (e.g., "FF0000").
    pub color: String,
    /// Line width in points.
    pub width: f64,
    /// Dash style (e.g., "solid", "dash", "dot", "dashDot", "lgDash", "sysDot").
    pub dash_style: String,
    /// Whether the line is smooth (Bezier curves between data points).
    pub smooth: bool,
}

impl LineStyle {
    /// Creates a new line style with default values.
    pub fn new() -> Self {
        Self {
            color: "000000".to_string(),
            width: 2.0,
            dash_style: "solid".to_string(),
            smooth: false,
        }
    }

    /// Sets the line color (hex string with or without #).
    pub fn with_color(mut self, color: impl Into<String>) -> Self {
        let mut color = color.into();
        if color.starts_with('#') {
            color = color[1..].to_string();
        }
        if color.len() == 8 {
            color = color[..6].to_string();
        }
        self.color = color;
        self
    }

    /// Sets the line width in points.
    pub fn with_width(mut self, width: f64) -> Self {
        self.width = width.max(0.0);
        self
    }

    /// Sets the dash style.
    pub fn with_dash_style(mut self, style: impl Into<String>) -> Self {
        self.dash_style = style.into();
        self
    }

    /// Sets whether the line is smooth.
    pub fn with_smooth(mut self, smooth: bool) -> Self {
        self.smooth = smooth;
        self
    }
}

impl Default for LineStyle {
    fn default() -> Self {
        Self::new()
    }
}

/// Scatter chart style (c:scatterStyle).
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ScatterStyle {
    /// Lines with markers.
    LineMarker,
    /// Lines only (no markers).
    Line,
    /// Markers only (no lines).
    Marker,
    /// Smooth lines with markers.
    SmoothMarker,
    /// Smooth lines only.
    Smooth,
}

impl ScatterStyle {
    /// Convert to XML scatterStyle value.
    pub fn to_xml(self) -> &'static str {
        match self {
            Self::LineMarker => "lineMarker",
            Self::Line => "line",
            Self::Marker => "marker",
            Self::SmoothMarker => "smoothMarker",
            Self::Smooth => "smooth",
        }
    }

    /// Parse from XML scatterStyle value.
    pub fn from_xml(value: &str) -> Option<Self> {
        match value {
            "lineMarker" => Some(Self::LineMarker),
            "line" => Some(Self::Line),
            "marker" => Some(Self::Marker),
            "smoothMarker" => Some(Self::SmoothMarker),
            "smooth" => Some(Self::Smooth),
            _ => None,
        }
    }
}

/// Marker styling for line/scatter chart series.
#[derive(Debug, Clone, PartialEq)]
pub struct SeriesMarker {
    /// Marker shape.
    pub shape: MarkerShape,
    /// Marker size (2-72 points).
    pub size: u32,
    /// Marker fill color.
    pub fill: Option<SeriesFill>,
    /// Marker border.
    pub border: Option<SeriesBorder>,
}

impl SeriesMarker {
    /// Creates a new marker with default values.
    pub fn new(shape: MarkerShape) -> Self {
        Self {
            shape,
            size: 5,
            fill: None,
            border: None,
        }
    }

    /// Sets the marker size (clamped to 2-72 points).
    pub fn with_size(mut self, size: u32) -> Self {
        self.size = size.clamp(2, 72);
        self
    }

    /// Sets the marker fill.
    pub fn with_fill(mut self, fill: SeriesFill) -> Self {
        self.fill = Some(fill);
        self
    }

    /// Sets the marker border.
    pub fn with_border(mut self, border: SeriesBorder) -> Self {
        self.border = Some(border);
        self
    }
}

/// A single data series in a chart.
#[derive(Debug, Clone)]
pub struct ChartSeries {
    /// Series name/label.
    name: String,
    /// Data values.
    values: Vec<f64>,
    /// Series fill color/pattern.
    fill: Option<SeriesFill>,
    /// Series border/outline.
    border: Option<SeriesBorder>,
    /// Marker styling for line/scatter charts.
    marker: Option<SeriesMarker>,
    /// Line styling for line/scatter chart series.
    line_style: Option<LineStyle>,
    /// Whether to smooth lines (for line charts).
    smooth: bool,
    /// Explosion percentage for pie charts (0-100%).
    explosion: Option<u32>,
    /// X-axis values for scatter/bubble charts (c:xVal).
    x_values: Vec<f64>,
    /// Bubble sizes for bubble charts (c:bubbleSize).
    bubble_sizes: Vec<f64>,
}

impl ChartSeries {
    pub fn new(name: impl Into<String>) -> Self {
        Self {
            name: name.into(),
            values: Vec::new(),
            fill: None,
            border: None,
            marker: None,
            line_style: None,
            smooth: false,
            explosion: None,
            x_values: Vec::new(),
            bubble_sizes: Vec::new(),
        }
    }

    pub fn name(&self) -> &str {
        &self.name
    }

    pub fn set_name(&mut self, name: impl Into<String>) {
        self.name = name.into();
    }

    pub fn values(&self) -> &[f64] {
        &self.values
    }

    pub fn add_value(&mut self, value: f64) {
        self.values.push(value);
    }

    pub fn set_values(&mut self, values: Vec<f64>) {
        self.values = values;
    }

    // ── Series styling ──

    /// Gets the series fill.
    pub fn fill(&self) -> Option<&SeriesFill> {
        self.fill.as_ref()
    }

    /// Sets the series fill.
    pub fn set_fill(&mut self, fill: SeriesFill) {
        self.fill = Some(fill);
    }

    /// Clears the series fill.
    pub fn clear_fill(&mut self) {
        self.fill = None;
    }

    /// Builder method to set the fill.
    pub fn with_fill(mut self, fill: SeriesFill) -> Self {
        self.fill = Some(fill);
        self
    }

    /// Gets the series border.
    pub fn border(&self) -> Option<&SeriesBorder> {
        self.border.as_ref()
    }

    /// Sets the series border.
    pub fn set_border(&mut self, border: SeriesBorder) {
        self.border = Some(border);
    }

    /// Clears the series border.
    pub fn clear_border(&mut self) {
        self.border = None;
    }

    /// Builder method to set the border.
    pub fn with_border(mut self, border: SeriesBorder) -> Self {
        self.border = Some(border);
        self
    }

    /// Gets the series marker.
    pub fn marker(&self) -> Option<&SeriesMarker> {
        self.marker.as_ref()
    }

    /// Sets the series marker (for line/scatter charts).
    pub fn set_marker(&mut self, marker: SeriesMarker) {
        self.marker = Some(marker);
    }

    /// Clears the series marker.
    pub fn clear_marker(&mut self) {
        self.marker = None;
    }

    /// Builder method to set the marker.
    pub fn with_marker(mut self, marker: SeriesMarker) -> Self {
        self.marker = Some(marker);
        self
    }

    // ── Line styling (for line/scatter charts) ──

    /// Gets the series line style.
    pub fn line_style(&self) -> Option<&LineStyle> {
        self.line_style.as_ref()
    }

    /// Sets the series line style.
    pub fn set_line_style(&mut self, style: LineStyle) {
        self.line_style = Some(style);
    }

    /// Clears the series line style.
    pub fn clear_line_style(&mut self) {
        self.line_style = None;
    }

    /// Builder method to set the line style.
    pub fn with_line_style(mut self, style: LineStyle) -> Self {
        self.line_style = Some(style);
        self
    }

    // ── X-values for scatter/bubble charts ──

    /// Gets the x-axis values (for scatter/bubble charts).
    pub fn x_values(&self) -> &[f64] {
        &self.x_values
    }

    /// Sets the x-axis values (for scatter/bubble charts).
    pub fn set_x_values(&mut self, values: Vec<f64>) {
        self.x_values = values;
    }

    /// Adds an x-axis value.
    pub fn add_x_value(&mut self, value: f64) {
        self.x_values.push(value);
    }

    /// Builder method to set x-axis values.
    pub fn with_x_values(mut self, values: Vec<f64>) -> Self {
        self.x_values = values;
        self
    }

    // ── Bubble sizes for bubble charts ──

    /// Gets the bubble sizes (for bubble charts).
    pub fn bubble_sizes(&self) -> &[f64] {
        &self.bubble_sizes
    }

    /// Sets the bubble sizes (for bubble charts).
    pub fn set_bubble_sizes(&mut self, sizes: Vec<f64>) {
        self.bubble_sizes = sizes;
    }

    /// Adds a bubble size value.
    pub fn add_bubble_size(&mut self, size: f64) {
        self.bubble_sizes.push(size);
    }

    /// Builder method to set bubble sizes.
    pub fn with_bubble_sizes(mut self, sizes: Vec<f64>) -> Self {
        self.bubble_sizes = sizes;
        self
    }

    /// Gets whether lines are smoothed.
    pub fn smooth(&self) -> bool {
        self.smooth
    }

    /// Sets whether to smooth lines (for line charts).
    pub fn set_smooth(&mut self, smooth: bool) {
        self.smooth = smooth;
    }

    /// Builder method to set line smoothing.
    pub fn with_smooth(mut self, smooth: bool) -> Self {
        self.smooth = smooth;
        self
    }

    /// Gets the explosion percentage (for pie charts).
    pub fn explosion(&self) -> Option<u32> {
        self.explosion
    }

    /// Sets the explosion percentage for pie chart slices (0-100%).
    pub fn set_explosion(&mut self, percent: u32) {
        self.explosion = Some(percent.min(100));
    }

    /// Clears the explosion.
    pub fn clear_explosion(&mut self) {
        self.explosion = None;
    }

    /// Builder method to set explosion percentage.
    pub fn with_explosion(mut self, percent: u32) -> Self {
        self.explosion = Some(percent.min(100));
        self
    }
}

impl PartialEq for ChartSeries {
    fn eq(&self, other: &Self) -> bool {
        self.name == other.name
            && self.values.len() == other.values.len()
            && self
                .values
                .iter()
                .zip(other.values.iter())
                .all(|(a, b)| a.to_bits() == b.to_bits())
            && self.fill == other.fill
            && self.border == other.border
            && self.marker == other.marker
            && self.line_style == other.line_style
            && self.smooth == other.smooth
            && self.explosion == other.explosion
            && self.x_values.len() == other.x_values.len()
            && self
                .x_values
                .iter()
                .zip(other.x_values.iter())
                .all(|(a, b)| a.to_bits() == b.to_bits())
            && self.bubble_sizes.len() == other.bubble_sizes.len()
            && self
                .bubble_sizes
                .iter()
                .zip(other.bubble_sizes.iter())
                .all(|(a, b)| a.to_bits() == b.to_bits())
    }
}

impl Eq for ChartSeries {}

/// Chart legend position.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum LegendPosition {
    Bottom,
    Top,
    Left,
    Right,
    TopRight,
}

impl LegendPosition {
    pub fn from_xml(value: &str) -> Option<Self> {
        match value {
            "b" => Some(Self::Bottom),
            "t" => Some(Self::Top),
            "l" => Some(Self::Left),
            "r" => Some(Self::Right),
            "tr" => Some(Self::TopRight),
            _ => None,
        }
    }

    pub fn to_xml(self) -> &'static str {
        match self {
            Self::Bottom => "b",
            Self::Top => "t",
            Self::Left => "l",
            Self::Right => "r",
            Self::TopRight => "tr",
        }
    }
}

/// Chart axis properties.
#[derive(Debug, Clone)]
pub struct ChartAxis {
    /// Axis title text.
    pub title: Option<String>,
    /// Minimum value for the axis scale.
    pub min_value: Option<f64>,
    /// Maximum value for the axis scale.
    pub max_value: Option<f64>,
    /// Major unit interval for tick marks and gridlines.
    pub major_unit: Option<f64>,
    /// Minor unit interval for minor tick marks.
    pub minor_unit: Option<f64>,
    /// Whether major gridlines are shown.
    pub has_major_gridlines: bool,
    /// Whether minor gridlines are shown.
    pub has_minor_gridlines: bool,
    /// Major tick mark type ("in", "out", "cross", "none").
    pub major_tick_mark: Option<String>,
    /// Minor tick mark type ("in", "out", "cross", "none").
    pub minor_tick_mark: Option<String>,
    /// Label position ("nextTo", "low", "high", "none").
    pub label_position: Option<String>,
    /// Number format for axis labels (e.g., "0.00", "#,##0").
    pub number_format: Option<String>,
    /// Font size for axis labels in points.
    pub label_font_size: Option<u32>,
    /// Whether axis is visible.
    pub visible: bool,
    /// Whether axis scale is logarithmic.
    pub logarithmic: bool,
    /// Logarithmic base (default 10).
    pub log_base: Option<f64>,
}

impl ChartAxis {
    pub fn new() -> Self {
        Self {
            title: None,
            min_value: None,
            max_value: None,
            major_unit: None,
            minor_unit: None,
            has_major_gridlines: false,
            has_minor_gridlines: false,
            major_tick_mark: None,
            minor_tick_mark: None,
            label_position: None,
            number_format: None,
            label_font_size: None,
            visible: true,
            logarithmic: false,
            log_base: None,
        }
    }

    /// Builder method to set title.
    pub fn with_title(mut self, title: impl Into<String>) -> Self {
        self.title = Some(title.into());
        self
    }

    /// Builder method to set min/max values.
    pub fn with_bounds(mut self, min: f64, max: f64) -> Self {
        self.min_value = Some(min);
        self.max_value = Some(max);
        self
    }

    /// Builder method to set major unit.
    pub fn with_major_unit(mut self, unit: f64) -> Self {
        self.major_unit = Some(unit);
        self
    }

    /// Builder method to set minor unit.
    pub fn with_minor_unit(mut self, unit: f64) -> Self {
        self.minor_unit = Some(unit);
        self
    }

    /// Builder method to show/hide major gridlines.
    pub fn with_major_gridlines(mut self, show: bool) -> Self {
        self.has_major_gridlines = show;
        self
    }

    /// Builder method to show/hide minor gridlines.
    pub fn with_minor_gridlines(mut self, show: bool) -> Self {
        self.has_minor_gridlines = show;
        self
    }

    /// Builder method to set tick marks.
    pub fn with_tick_marks(mut self, major: impl Into<String>, minor: impl Into<String>) -> Self {
        self.major_tick_mark = Some(major.into());
        self.minor_tick_mark = Some(minor.into());
        self
    }

    /// Builder method to set label position.
    pub fn with_label_position(mut self, position: impl Into<String>) -> Self {
        self.label_position = Some(position.into());
        self
    }

    /// Builder method to set number format.
    pub fn with_number_format(mut self, format: impl Into<String>) -> Self {
        self.number_format = Some(format.into());
        self
    }

    /// Builder method to set label font size.
    pub fn with_label_font_size(mut self, size: u32) -> Self {
        self.label_font_size = Some(size);
        self
    }

    /// Builder method to set visibility.
    pub fn with_visibility(mut self, visible: bool) -> Self {
        self.visible = visible;
        self
    }

    /// Builder method to set logarithmic scale.
    pub fn with_logarithmic(mut self, enabled: bool, base: Option<f64>) -> Self {
        self.logarithmic = enabled;
        self.log_base = base;
        self
    }
}

impl Default for ChartAxis {
    fn default() -> Self {
        Self::new()
    }
}

impl PartialEq for ChartAxis {
    fn eq(&self, other: &Self) -> bool {
        self.title == other.title
            && float_option_eq(self.min_value, other.min_value)
            && float_option_eq(self.max_value, other.max_value)
            && float_option_eq(self.major_unit, other.major_unit)
            && float_option_eq(self.minor_unit, other.minor_unit)
            && self.has_major_gridlines == other.has_major_gridlines
            && self.has_minor_gridlines == other.has_minor_gridlines
            && self.major_tick_mark == other.major_tick_mark
            && self.minor_tick_mark == other.minor_tick_mark
            && self.label_position == other.label_position
            && self.number_format == other.number_format
            && self.label_font_size == other.label_font_size
            && self.visible == other.visible
            && self.logarithmic == other.logarithmic
            && float_option_eq(self.log_base, other.log_base)
    }
}

impl Eq for ChartAxis {}

fn float_option_eq(a: Option<f64>, b: Option<f64>) -> bool {
    match (a, b) {
        (Some(va), Some(vb)) => va.to_bits() == vb.to_bits(),
        (None, None) => true,
        _ => false,
    }
}

/// Data label configuration for chart series or chart-wide.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ChartDataLabel {
    /// Show category name in labels.
    pub show_category_name: bool,
    /// Show series name in labels.
    pub show_series_name: bool,
    /// Show value in labels.
    pub show_value: bool,
    /// Show percentage (for pie charts).
    pub show_percentage: bool,
    /// Show legend key (colored marker) next to label.
    pub show_legend_key: bool,
    /// Label position ("bestFit", "center", "inEnd", "inBase", "outEnd").
    pub position: Option<String>,
    /// Font size for labels in points.
    pub font_size: Option<u32>,
    /// Font color as hex (e.g., "000000" for black).
    pub font_color: Option<String>,
    /// Separator character between label components.
    pub separator: Option<String>,
}

impl ChartDataLabel {
    /// Creates default data labels (show value only).
    pub fn new() -> Self {
        Self {
            show_category_name: false,
            show_series_name: false,
            show_value: true,
            show_percentage: false,
            show_legend_key: false,
            position: None,
            font_size: None,
            font_color: None,
            separator: None,
        }
    }

    /// Builder method to show category name.
    pub fn with_category_name(mut self, show: bool) -> Self {
        self.show_category_name = show;
        self
    }

    /// Builder method to show series name.
    pub fn with_series_name(mut self, show: bool) -> Self {
        self.show_series_name = show;
        self
    }

    /// Builder method to show value.
    pub fn with_value(mut self, show: bool) -> Self {
        self.show_value = show;
        self
    }

    /// Builder method to show percentage (for pie charts).
    pub fn with_percentage(mut self, show: bool) -> Self {
        self.show_percentage = show;
        self
    }

    /// Builder method to show legend key.
    pub fn with_legend_key(mut self, show: bool) -> Self {
        self.show_legend_key = show;
        self
    }

    /// Builder method to set label position.
    pub fn with_position(mut self, position: impl Into<String>) -> Self {
        self.position = Some(position.into());
        self
    }

    /// Builder method to set font size.
    pub fn with_font_size(mut self, size: u32) -> Self {
        self.font_size = Some(size);
        self
    }

    /// Builder method to set font color.
    pub fn with_font_color(mut self, color: impl Into<String>) -> Self {
        let mut color = color.into();
        if color.starts_with('#') {
            color = color[1..].to_string();
        }
        self.font_color = Some(color);
        self
    }

    /// Builder method to set separator.
    pub fn with_separator(mut self, separator: impl Into<String>) -> Self {
        self.separator = Some(separator.into());
        self
    }
}

impl Default for ChartDataLabel {
    fn default() -> Self {
        Self::new()
    }
}

#[derive(Debug, Clone)]
pub struct Chart {
    title: String,
    categories: Vec<String>,
    values: Vec<f64>,
    /// Chart type (Feature #13).
    chart_type: ChartType,
    /// Multiple series (Feature #13). The `values` field above is the first series for
    /// backward compatibility.
    additional_series: Vec<ChartSeries>,
    /// Whether to show legend (Feature #13).
    show_legend: bool,
    /// Legend position (Feature #13).
    legend_position: Option<LegendPosition>,
    /// Whether barDir="bar" (horizontal) vs "col" (vertical) for bar/column charts.
    bar_direction_horizontal: bool,
    /// Category (horizontal) axis properties.
    category_axis: Option<ChartAxis>,
    /// Value (vertical) axis properties.
    value_axis: Option<ChartAxis>,
    /// Data label configuration for all series.
    data_labels: Option<ChartDataLabel>,

    // ── Bar/Column chart properties (c:barChart) ──
    /// Gap width between bar clusters as percentage (0-500, default 150).
    /// Maps to `c:gapWidth` in OOXML.
    bar_gap_width: Option<u32>,
    /// Overlap of bars within a cluster as percentage (-100 to 100).
    /// Positive values overlap, negative values add gaps. Maps to `c:overlap`.
    bar_overlap: Option<i32>,

    // ── Pie/Doughnut chart properties ──
    /// First slice angle in degrees (0-360). Maps to `c:firstSliceAng`.
    pie_first_slice_angle: Option<u32>,
    /// Hole size for doughnut charts as percentage (1-90). Maps to `c:holeSize`.
    pie_hole_size: Option<u32>,

    // ── Scatter/Bubble chart properties ──
    /// Scatter chart visual style. Maps to `c:scatterStyle`.
    scatter_style: Option<ScatterStyle>,
    /// Bubble scale percentage for bubble charts (0-300, default 100).
    /// Maps to `c:bubbleScale`.
    bubble_scale: Option<u32>,
}

impl Chart {
    pub fn new(title: impl Into<String>) -> Self {
        Self {
            title: title.into(),
            categories: Vec::new(),
            values: Vec::new(),
            chart_type: ChartType::Column,
            additional_series: Vec::new(),
            show_legend: false,
            legend_position: None,
            bar_direction_horizontal: false,
            category_axis: None,
            value_axis: None,
            data_labels: None,
            bar_gap_width: None,
            bar_overlap: None,
            pie_first_slice_angle: None,
            pie_hole_size: None,
            scatter_style: None,
            bubble_scale: None,
        }
    }

    pub fn title(&self) -> &str {
        &self.title
    }

    pub fn set_title(&mut self, title: impl Into<String>) {
        self.title = title.into();
    }

    pub fn categories(&self) -> &[String] {
        &self.categories
    }

    pub fn values(&self) -> &[f64] {
        &self.values
    }

    pub fn add_data_point(&mut self, category: impl Into<String>, value: f64) {
        self.categories.push(category.into());
        self.values.push(value);
    }

    pub fn clear_data_points(&mut self) {
        self.categories.clear();
        self.values.clear();
    }

    pub fn point_count(&self) -> usize {
        self.categories.len().min(self.values.len())
    }

    pub(crate) fn set_data_points(&mut self, categories: Vec<String>, values: Vec<f64>) {
        self.categories = categories;
        self.values = values;
    }

    // ── Feature #13: Chart improvements ──

    /// Chart type.
    pub fn chart_type(&self) -> ChartType {
        self.chart_type
    }

    /// Set chart type.
    pub fn set_chart_type(&mut self, chart_type: ChartType) {
        self.chart_type = chart_type;
    }

    /// Additional data series beyond the first.
    pub fn additional_series(&self) -> &[ChartSeries] {
        &self.additional_series
    }

    /// Add a data series.
    pub fn add_series(&mut self, series: ChartSeries) {
        self.additional_series.push(series);
    }

    /// Remove a data series by index.
    ///
    /// Returns `None` if the index is out of bounds.
    pub fn remove_series(&mut self, index: usize) -> Option<ChartSeries> {
        if index < self.additional_series.len() {
            Some(self.additional_series.remove(index))
        } else {
            None
        }
    }

    /// Clear all additional series.
    pub fn clear_additional_series(&mut self) {
        self.additional_series.clear();
    }

    /// Whether to show legend.
    pub fn show_legend(&self) -> bool {
        self.show_legend
    }

    /// Set whether to show legend.
    pub fn set_show_legend(&mut self, show: bool) {
        self.show_legend = show;
    }

    /// Legend position.
    pub fn legend_position(&self) -> Option<LegendPosition> {
        self.legend_position
    }

    /// Set legend position.
    pub fn set_legend_position(&mut self, position: LegendPosition) {
        self.legend_position = Some(position);
        self.show_legend = true;
    }

    /// Whether bar direction is horizontal (Bar chart) vs vertical (Column chart).
    pub fn is_bar_direction_horizontal(&self) -> bool {
        self.bar_direction_horizontal
    }

    /// Set bar direction.
    pub fn set_bar_direction_horizontal(&mut self, horizontal: bool) {
        self.bar_direction_horizontal = horizontal;
        if horizontal {
            self.chart_type = ChartType::Bar;
        }
    }

    // ── Chart axes ──

    /// Category (horizontal) axis properties.
    pub fn category_axis(&self) -> Option<&ChartAxis> {
        self.category_axis.as_ref()
    }

    /// Set category axis properties.
    pub fn set_category_axis(&mut self, axis: ChartAxis) {
        self.category_axis = Some(axis);
    }

    /// Clear category axis.
    pub fn clear_category_axis(&mut self) {
        self.category_axis = None;
    }

    /// Value (vertical) axis properties.
    pub fn value_axis(&self) -> Option<&ChartAxis> {
        self.value_axis.as_ref()
    }

    /// Set value axis properties.
    pub fn set_value_axis(&mut self, axis: ChartAxis) {
        self.value_axis = Some(axis);
    }

    /// Clear value axis.
    pub fn clear_value_axis(&mut self) {
        self.value_axis = None;
    }

    /// Set the category axis title.
    ///
    /// Creates a default axis if one doesn't exist.
    pub fn set_category_axis_title(&mut self, title: impl Into<String>) {
        if let Some(axis) = &mut self.category_axis {
            axis.title = Some(title.into());
        } else {
            self.category_axis = Some(ChartAxis::new().with_title(title));
        }
    }

    /// Set the value axis title.
    ///
    /// Creates a default axis if one doesn't exist.
    pub fn set_value_axis_title(&mut self, title: impl Into<String>) {
        if let Some(axis) = &mut self.value_axis {
            axis.title = Some(title.into());
        } else {
            self.value_axis = Some(ChartAxis::new().with_title(title));
        }
    }

    // ── Data labels ──

    /// Data label configuration.
    pub fn data_labels(&self) -> Option<&ChartDataLabel> {
        self.data_labels.as_ref()
    }

    /// Set data label configuration for all series.
    pub fn set_data_labels(&mut self, labels: ChartDataLabel) {
        self.data_labels = Some(labels);
    }

    /// Clear data labels.
    pub fn clear_data_labels(&mut self) {
        self.data_labels = None;
    }

    // ── Bar/Column chart properties ──

    /// Gap width between bar clusters as percentage (0-500).
    /// Default in OOXML is 150%. Maps to `c:gapWidth`.
    pub fn bar_gap_width(&self) -> Option<u32> {
        self.bar_gap_width
    }

    /// Sets the gap width between bar clusters (0-500%).
    pub fn set_bar_gap_width(&mut self, width: u32) {
        self.bar_gap_width = Some(width.min(500));
    }

    /// Clears the bar gap width (reverts to default).
    pub fn clear_bar_gap_width(&mut self) {
        self.bar_gap_width = None;
    }

    /// Overlap of bars within a cluster as percentage (-100 to 100).
    /// Positive values overlap bars, negative values add gaps. Maps to `c:overlap`.
    pub fn bar_overlap(&self) -> Option<i32> {
        self.bar_overlap
    }

    /// Sets the bar overlap percentage (-100 to 100).
    pub fn set_bar_overlap(&mut self, overlap: i32) {
        self.bar_overlap = Some(overlap.clamp(-100, 100));
    }

    /// Clears the bar overlap (reverts to default).
    pub fn clear_bar_overlap(&mut self) {
        self.bar_overlap = None;
    }

    // ── Pie/Doughnut chart properties ──

    /// First slice angle in degrees (0-360). Maps to `c:firstSliceAng`.
    pub fn pie_first_slice_angle(&self) -> Option<u32> {
        self.pie_first_slice_angle
    }

    /// Sets the first slice angle for pie charts (0-360 degrees).
    pub fn set_pie_first_slice_angle(&mut self, angle: u32) {
        self.pie_first_slice_angle = Some(angle.min(360));
    }

    /// Clears the first slice angle (reverts to default 0).
    pub fn clear_pie_first_slice_angle(&mut self) {
        self.pie_first_slice_angle = None;
    }

    /// Hole size for doughnut charts as percentage (1-90). Maps to `c:holeSize`.
    pub fn pie_hole_size(&self) -> Option<u32> {
        self.pie_hole_size
    }

    /// Sets the hole size for doughnut charts (1-90%).
    pub fn set_pie_hole_size(&mut self, size: u32) {
        self.pie_hole_size = Some(size.clamp(1, 90));
    }

    /// Clears the hole size (reverts to default).
    pub fn clear_pie_hole_size(&mut self) {
        self.pie_hole_size = None;
    }

    // ── Scatter/Bubble chart properties ──

    /// Scatter chart visual style. Maps to `c:scatterStyle`.
    pub fn scatter_style(&self) -> Option<ScatterStyle> {
        self.scatter_style
    }

    /// Sets the scatter chart style.
    pub fn set_scatter_style(&mut self, style: ScatterStyle) {
        self.scatter_style = Some(style);
    }

    /// Clears the scatter style (reverts to default).
    pub fn clear_scatter_style(&mut self) {
        self.scatter_style = None;
    }

    /// Bubble scale percentage for bubble charts (0-300). Maps to `c:bubbleScale`.
    pub fn bubble_scale(&self) -> Option<u32> {
        self.bubble_scale
    }

    /// Sets the bubble scale percentage (0-300%).
    pub fn set_bubble_scale(&mut self, scale: u32) {
        self.bubble_scale = Some(scale.min(300));
    }

    /// Clears the bubble scale (reverts to default 100%).
    pub fn clear_bubble_scale(&mut self) {
        self.bubble_scale = None;
    }
}

impl PartialEq for Chart {
    fn eq(&self, other: &Self) -> bool {
        self.title == other.title
            && self.categories == other.categories
            && self.values.len() == other.values.len()
            && self
                .values
                .iter()
                .zip(other.values.iter())
                .all(|(lhs, rhs)| lhs.to_bits() == rhs.to_bits())
            && self.chart_type == other.chart_type
            && self.additional_series == other.additional_series
            && self.show_legend == other.show_legend
            && self.legend_position == other.legend_position
            && self.category_axis == other.category_axis
            && self.value_axis == other.value_axis
            && self.data_labels == other.data_labels
            && self.bar_gap_width == other.bar_gap_width
            && self.bar_overlap == other.bar_overlap
            && self.pie_first_slice_angle == other.pie_first_slice_angle
            && self.pie_hole_size == other.pie_hole_size
            && self.scatter_style == other.scatter_style
            && self.bubble_scale == other.bubble_scale
    }
}

impl Eq for Chart {}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn stores_title_and_series_data() {
        let mut chart = Chart::new("Revenue");
        chart.add_data_point("Q1", 10.0);
        chart.add_data_point("Q2", 14.5);

        assert_eq!(chart.title(), "Revenue");
        assert_eq!(chart.categories(), ["Q1", "Q2"]);
        assert_eq!(chart.values(), [10.0, 14.5]);
        assert_eq!(chart.point_count(), 2);
    }

    #[test]
    fn clears_data_points() {
        let mut chart = Chart::new("Revenue");
        chart.add_data_point("Q1", 10.0);
        chart.clear_data_points();

        assert!(chart.categories().is_empty());
        assert!(chart.values().is_empty());
    }

    #[test]
    fn chart_type_defaults_to_column() {
        let chart = Chart::new("Revenue");
        assert_eq!(chart.chart_type(), ChartType::Column);
    }

    #[test]
    fn chart_type_xml_roundtrip() {
        assert_eq!(
            ChartType::from_xml_element("barChart"),
            Some(ChartType::Column)
        );
        assert_eq!(
            ChartType::from_xml_element("lineChart"),
            Some(ChartType::Line)
        );
        assert_eq!(
            ChartType::from_xml_element("pieChart"),
            Some(ChartType::Pie)
        );
        assert_eq!(
            ChartType::from_xml_element("doughnutChart"),
            Some(ChartType::Doughnut)
        );
    }

    #[test]
    fn multiple_series() {
        let mut chart = Chart::new("Revenue");
        chart.add_data_point("Q1", 10.0);
        chart.add_data_point("Q2", 14.5);

        let mut series2 = ChartSeries::new("Expenses");
        series2.add_value(8.0);
        series2.add_value(12.0);
        chart.add_series(series2);

        assert_eq!(chart.additional_series().len(), 1);
        assert_eq!(chart.additional_series()[0].name(), "Expenses");
        assert_eq!(chart.additional_series()[0].values(), [8.0, 12.0]);
    }

    #[test]
    fn legend_configuration() {
        let mut chart = Chart::new("Revenue");
        assert!(!chart.show_legend());
        assert!(chart.legend_position().is_none());

        chart.set_legend_position(LegendPosition::Bottom);
        assert!(chart.show_legend());
        assert_eq!(chart.legend_position(), Some(LegendPosition::Bottom));
    }

    #[test]
    fn bar_direction() {
        let mut chart = Chart::new("Revenue");
        assert!(!chart.is_bar_direction_horizontal());

        chart.set_bar_direction_horizontal(true);
        assert!(chart.is_bar_direction_horizontal());
        assert_eq!(chart.chart_type(), ChartType::Bar);
    }

    // ── Chart axes tests ──

    #[test]
    fn chart_axis_defaults() {
        let chart = Chart::new("Revenue");
        assert!(chart.category_axis().is_none());
        assert!(chart.value_axis().is_none());
    }

    #[test]
    fn chart_axis_roundtrip() {
        let mut chart = Chart::new("Revenue");

        let mut cat_axis = ChartAxis::new();
        cat_axis.title = Some("Quarter".to_string());
        cat_axis.has_major_gridlines = true;
        chart.set_category_axis(cat_axis);

        let mut val_axis = ChartAxis::new();
        val_axis.title = Some("Revenue ($M)".to_string());
        val_axis.min_value = Some(0.0);
        val_axis.max_value = Some(100.0);
        val_axis.major_unit = Some(25.0);
        val_axis.has_major_gridlines = true;
        chart.set_value_axis(val_axis);

        let cat = chart.category_axis().unwrap();
        assert_eq!(cat.title.as_deref(), Some("Quarter"));
        assert!(cat.has_major_gridlines);
        assert!(cat.min_value.is_none());

        let val = chart.value_axis().unwrap();
        assert_eq!(val.title.as_deref(), Some("Revenue ($M)"));
        assert_eq!(val.min_value, Some(0.0));
        assert_eq!(val.max_value, Some(100.0));
        assert_eq!(val.major_unit, Some(25.0));
        assert!(val.has_major_gridlines);

        chart.clear_category_axis();
        chart.clear_value_axis();
        assert!(chart.category_axis().is_none());
        assert!(chart.value_axis().is_none());
    }

    #[test]
    fn chart_axis_equality() {
        let mut a = ChartAxis::new();
        a.title = Some("X".to_string());
        a.min_value = Some(0.0);

        let mut b = ChartAxis::new();
        b.title = Some("X".to_string());
        b.min_value = Some(0.0);

        assert_eq!(a, b);

        b.min_value = Some(1.0);
        assert_ne!(a, b);
    }

    // ── Series styling tests ──

    #[test]
    fn series_fill_solid() {
        let fill = SeriesFill::solid("#FF0000");
        assert_eq!(fill, SeriesFill::Solid("FF0000".to_string()));

        let fill2 = SeriesFill::solid("00FF00");
        assert_eq!(fill2, SeriesFill::Solid("00FF00".to_string()));
    }

    #[test]
    fn series_border_builder() {
        let border = SeriesBorder::new()
            .with_color("#FF0000")
            .with_width(2.5)
            .with_dash_style("dash");

        assert_eq!(border.color, "FF0000");
        assert_eq!(border.width, 2.5);
        assert_eq!(border.dash_style, "dash");
    }

    #[test]
    fn series_marker_builder() {
        let marker = SeriesMarker::new(MarkerShape::Circle)
            .with_size(10)
            .with_fill(SeriesFill::solid("0000FF"));

        assert_eq!(marker.shape, MarkerShape::Circle);
        assert_eq!(marker.size, 10);
        assert!(marker.fill.is_some());
    }

    #[test]
    fn chart_series_with_styling() {
        let series = ChartSeries::new("Test")
            .with_fill(SeriesFill::solid("FF0000"))
            .with_border(SeriesBorder::new().with_color("000000"))
            .with_smooth(true);

        assert!(series.fill().is_some());
        assert!(series.border().is_some());
        assert!(series.smooth());
    }

    #[test]
    fn remove_series_by_index() {
        let mut chart = Chart::new("Test");
        chart.add_series(ChartSeries::new("S1"));
        chart.add_series(ChartSeries::new("S2"));

        let removed = chart.remove_series(0);
        assert!(removed.is_some());
        assert_eq!(removed.unwrap().name(), "S1");
        assert_eq!(chart.additional_series().len(), 1);
    }

    #[test]
    fn remove_series_out_of_bounds() {
        let mut chart = Chart::new("Test");
        assert!(chart.remove_series(0).is_none());
    }

    #[test]
    fn set_axis_titles_convenience() {
        let mut chart = Chart::new("Test");
        assert!(chart.category_axis().is_none());
        assert!(chart.value_axis().is_none());

        chart.set_category_axis_title("Months");
        chart.set_value_axis_title("Revenue");

        assert_eq!(
            chart.category_axis().unwrap().title.as_deref(),
            Some("Months")
        );
        assert_eq!(
            chart.value_axis().unwrap().title.as_deref(),
            Some("Revenue")
        );

        // Calling again on existing axis should update title
        chart.set_category_axis_title("Quarters");
        assert_eq!(
            chart.category_axis().unwrap().title.as_deref(),
            Some("Quarters")
        );
    }
}
