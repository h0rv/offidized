//! Additional chart types beyond the basic ones in chart.rs.
//!
//! This module provides specialized chart types including bubble charts, stock charts,
//! surface charts, 3D variants, and combination charts.

use std::io::{Read as IoRead, Write as IoWrite};

/// Bubble chart properties.
///
/// Bubble charts display three dimensions of data. Each data point has an X position, Y position,
/// and bubble size. Maps to `c:bubbleChart` in OOXML.
#[derive(Debug, Clone, PartialEq)]
pub struct BubbleChart {
    /// Series data (each series has x, y, and bubble size values).
    pub series: Vec<BubbleSeries>,
    /// Bubble scale percentage (0-300, default 100). Maps to `c:bubbleScale`.
    pub bubble_scale: Option<u32>,
    /// Whether to show negative bubbles. Maps to `c:showNegBubbles`.
    pub show_negative_bubbles: bool,
    /// Bubble size represents area (true) or width (false). Maps to `c:sizeRepresents`.
    pub size_represents_area: bool,
}

impl BubbleChart {
    /// Creates a new bubble chart.
    pub fn new() -> Self {
        Self {
            series: Vec::new(),
            bubble_scale: None,
            show_negative_bubbles: false,
            size_represents_area: true,
        }
    }

    /// Sets the bubble scale percentage (0-300%).
    pub fn set_bubble_scale(&mut self, scale: u32) {
        self.bubble_scale = Some(scale.min(300));
    }

    /// Adds a bubble series.
    pub fn add_series(&mut self, series: BubbleSeries) {
        self.series.push(series);
    }
}

impl Default for BubbleChart {
    fn default() -> Self {
        Self::new()
    }
}

/// A single series in a bubble chart.
#[derive(Debug, Clone, PartialEq)]
pub struct BubbleSeries {
    /// Series name.
    pub name: String,
    /// X-axis values.
    pub x_values: Vec<f64>,
    /// Y-axis values.
    pub y_values: Vec<f64>,
    /// Bubble sizes.
    pub bubble_sizes: Vec<f64>,
}

impl BubbleSeries {
    /// Creates a new bubble series.
    pub fn new(name: impl Into<String>) -> Self {
        Self {
            name: name.into(),
            x_values: Vec::new(),
            y_values: Vec::new(),
            bubble_sizes: Vec::new(),
        }
    }

    /// Adds a data point with x, y, and bubble size.
    pub fn add_point(&mut self, x: f64, y: f64, size: f64) {
        self.x_values.push(x);
        self.y_values.push(y);
        self.bubble_sizes.push(size);
    }
}

/// Stock chart style variants.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum StockChartStyle {
    /// High-Low-Close (3 series).
    HighLowClose,
    /// Open-High-Low-Close (4 series).
    OpenHighLowClose,
    /// Volume-High-Low-Close (4 series).
    VolumeHighLowClose,
    /// Volume-Open-High-Low-Close (5 series).
    VolumeOpenHighLowClose,
}

impl StockChartStyle {
    /// Returns the number of required series for this style.
    pub fn series_count(self) -> usize {
        match self {
            Self::HighLowClose => 3,
            Self::OpenHighLowClose | Self::VolumeHighLowClose => 4,
            Self::VolumeOpenHighLowClose => 5,
        }
    }
}

/// Stock chart properties.
///
/// Stock charts display stock market data with Open, High, Low, Close, and optionally Volume.
/// Maps to `c:stockChart` in OOXML.
#[derive(Debug, Clone, PartialEq)]
pub struct StockChart {
    /// Chart style determining which series are shown.
    pub style: StockChartStyle,
    /// Series data (order matters: Volume, Open, High, Low, Close).
    pub series: Vec<StockSeries>,
    /// Up/down bars showing difference between open and close.
    pub show_up_down_bars: bool,
    /// High-Low lines connecting high and low values.
    pub show_high_low_lines: bool,
}

impl StockChart {
    /// Creates a new stock chart.
    pub fn new(style: StockChartStyle) -> Self {
        Self {
            style,
            series: Vec::new(),
            show_up_down_bars: false,
            show_high_low_lines: true,
        }
    }

    /// Adds a stock series.
    pub fn add_series(&mut self, series: StockSeries) {
        self.series.push(series);
    }
}

/// A single series in a stock chart.
#[derive(Debug, Clone, PartialEq)]
pub struct StockSeries {
    /// Series name (e.g., "High", "Low", "Close", "Volume").
    pub name: String,
    /// Data values.
    pub values: Vec<f64>,
}

impl StockSeries {
    /// Creates a new stock series.
    pub fn new(name: impl Into<String>) -> Self {
        Self {
            name: name.into(),
            values: Vec::new(),
        }
    }

    /// Adds a value to the series.
    pub fn add_value(&mut self, value: f64) {
        self.values.push(value);
    }
}

/// Surface chart properties.
///
/// Surface charts display a 3D surface over a two-dimensional grid.
/// Maps to `c:surfaceChart` or `c:surface3DChart` in OOXML.
#[derive(Debug, Clone, PartialEq)]
pub struct SurfaceChart {
    /// Series data.
    pub series: Vec<SurfaceSeries>,
    /// Whether to display as wireframe (lines only).
    pub wireframe: bool,
    /// Whether to display as contour (2D top-down view).
    pub contour: bool,
}

impl SurfaceChart {
    /// Creates a new surface chart.
    pub fn new() -> Self {
        Self {
            series: Vec::new(),
            wireframe: false,
            contour: false,
        }
    }

    /// Adds a surface series.
    pub fn add_series(&mut self, series: SurfaceSeries) {
        self.series.push(series);
    }
}

impl Default for SurfaceChart {
    fn default() -> Self {
        Self::new()
    }
}

/// A single series in a surface chart.
#[derive(Debug, Clone, PartialEq)]
pub struct SurfaceSeries {
    /// Series name.
    pub name: String,
    /// Data values (organized as grid).
    pub values: Vec<Vec<f64>>,
}

impl SurfaceSeries {
    /// Creates a new surface series.
    pub fn new(name: impl Into<String>) -> Self {
        Self {
            name: name.into(),
            values: Vec::new(),
        }
    }

    /// Adds a row of values.
    pub fn add_row(&mut self, row: Vec<f64>) {
        self.values.push(row);
    }
}

/// 3D bar shape styles for 3D bar/column charts.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum BarShape3D {
    /// Rectangular box.
    Box,
    /// Cylinder.
    Cylinder,
    /// Cone.
    Cone,
    /// Pyramid.
    Pyramid,
}

impl BarShape3D {
    /// Convert to XML shape value.
    pub fn to_xml(self) -> &'static str {
        match self {
            Self::Box => "box",
            Self::Cylinder => "cylinder",
            Self::Cone => "cone",
            Self::Pyramid => "pyramid",
        }
    }

    /// Parse from XML shape value.
    pub fn from_xml(value: &str) -> Option<Self> {
        match value {
            "box" => Some(Self::Box),
            "cylinder" => Some(Self::Cylinder),
            "cone" => Some(Self::Cone),
            "pyramid" => Some(Self::Pyramid),
            _ => None,
        }
    }
}

/// 3D bar chart properties.
///
/// Maps to `c:bar3DChart` in OOXML.
#[derive(Debug, Clone, PartialEq)]
pub struct Bar3DChart {
    /// Series data.
    pub series: Vec<Bar3DSeries>,
    /// Gap depth between series (0-500%, default 150).
    pub gap_depth: Option<u32>,
    /// Bar shape (box, cylinder, cone, pyramid).
    pub shape: BarShape3D,
    /// Bar direction: true for horizontal (bar), false for vertical (column).
    pub horizontal: bool,
}

impl Bar3DChart {
    /// Creates a new 3D bar chart.
    pub fn new(horizontal: bool) -> Self {
        Self {
            series: Vec::new(),
            gap_depth: None,
            shape: BarShape3D::Box,
            horizontal,
        }
    }

    /// Sets the gap depth (0-500%).
    pub fn set_gap_depth(&mut self, depth: u32) {
        self.gap_depth = Some(depth.min(500));
    }

    /// Adds a series.
    pub fn add_series(&mut self, series: Bar3DSeries) {
        self.series.push(series);
    }
}

/// A single series in a 3D bar chart.
#[derive(Debug, Clone, PartialEq)]
pub struct Bar3DSeries {
    /// Series name.
    pub name: String,
    /// Data values.
    pub values: Vec<f64>,
}

impl Bar3DSeries {
    /// Creates a new 3D bar series.
    pub fn new(name: impl Into<String>) -> Self {
        Self {
            name: name.into(),
            values: Vec::new(),
        }
    }

    /// Adds a value to the series.
    pub fn add_value(&mut self, value: f64) {
        self.values.push(value);
    }
}

/// 3D column chart (alias for Bar3DChart with vertical bars).
pub type Column3DChart = Bar3DChart;

/// 3D pie chart properties.
///
/// Maps to `c:pie3DChart` in OOXML.
#[derive(Debug, Clone, PartialEq)]
pub struct Pie3DChart {
    /// Series data (typically one series).
    pub series: Vec<Pie3DSeries>,
    /// Rotation angle in degrees (0-360).
    pub rotation: Option<u32>,
    /// Elevation angle in degrees (0-90).
    pub elevation: Option<u32>,
}

impl Pie3DChart {
    /// Creates a new 3D pie chart.
    pub fn new() -> Self {
        Self {
            series: Vec::new(),
            rotation: None,
            elevation: None,
        }
    }

    /// Sets the rotation angle (0-360 degrees).
    pub fn set_rotation(&mut self, angle: u32) {
        self.rotation = Some(angle.min(360));
    }

    /// Sets the elevation angle (0-90 degrees).
    pub fn set_elevation(&mut self, angle: u32) {
        self.elevation = Some(angle.min(90));
    }

    /// Adds a series.
    pub fn add_series(&mut self, series: Pie3DSeries) {
        self.series.push(series);
    }
}

impl Default for Pie3DChart {
    fn default() -> Self {
        Self::new()
    }
}

/// A single series in a 3D pie chart.
#[derive(Debug, Clone, PartialEq)]
pub struct Pie3DSeries {
    /// Series name.
    pub name: String,
    /// Data values.
    pub values: Vec<f64>,
    /// Category names for each slice.
    pub categories: Vec<String>,
}

impl Pie3DSeries {
    /// Creates a new 3D pie series.
    pub fn new(name: impl Into<String>) -> Self {
        Self {
            name: name.into(),
            values: Vec::new(),
            categories: Vec::new(),
        }
    }

    /// Adds a data point with category and value.
    pub fn add_point(&mut self, category: impl Into<String>, value: f64) {
        self.categories.push(category.into());
        self.values.push(value);
    }
}

/// 3D line chart properties.
///
/// Maps to `c:line3DChart` in OOXML.
#[derive(Debug, Clone, PartialEq)]
pub struct Line3DChart {
    /// Series data.
    pub series: Vec<Line3DSeries>,
    /// Gap depth between series (0-500%, default 150).
    pub gap_depth: Option<u32>,
}

impl Line3DChart {
    /// Creates a new 3D line chart.
    pub fn new() -> Self {
        Self {
            series: Vec::new(),
            gap_depth: None,
        }
    }

    /// Sets the gap depth (0-500%).
    pub fn set_gap_depth(&mut self, depth: u32) {
        self.gap_depth = Some(depth.min(500));
    }

    /// Adds a series.
    pub fn add_series(&mut self, series: Line3DSeries) {
        self.series.push(series);
    }
}

impl Default for Line3DChart {
    fn default() -> Self {
        Self::new()
    }
}

/// A single series in a 3D line chart.
#[derive(Debug, Clone, PartialEq)]
pub struct Line3DSeries {
    /// Series name.
    pub name: String,
    /// Data values.
    pub values: Vec<f64>,
}

impl Line3DSeries {
    /// Creates a new 3D line series.
    pub fn new(name: impl Into<String>) -> Self {
        Self {
            name: name.into(),
            values: Vec::new(),
        }
    }

    /// Adds a value to the series.
    pub fn add_value(&mut self, value: f64) {
        self.values.push(value);
    }
}

/// 3D area chart properties.
///
/// Maps to `c:area3DChart` in OOXML.
#[derive(Debug, Clone, PartialEq)]
pub struct Area3DChart {
    /// Series data.
    pub series: Vec<Area3DSeries>,
    /// Gap depth between series (0-500%, default 150).
    pub gap_depth: Option<u32>,
}

impl Area3DChart {
    /// Creates a new 3D area chart.
    pub fn new() -> Self {
        Self {
            series: Vec::new(),
            gap_depth: None,
        }
    }

    /// Sets the gap depth (0-500%).
    pub fn set_gap_depth(&mut self, depth: u32) {
        self.gap_depth = Some(depth.min(500));
    }

    /// Adds a series.
    pub fn add_series(&mut self, series: Area3DSeries) {
        self.series.push(series);
    }
}

impl Default for Area3DChart {
    fn default() -> Self {
        Self::new()
    }
}

/// A single series in a 3D area chart.
#[derive(Debug, Clone, PartialEq)]
pub struct Area3DSeries {
    /// Series name.
    pub name: String,
    /// Data values.
    pub values: Vec<f64>,
}

impl Area3DSeries {
    /// Creates a new 3D area series.
    pub fn new(name: impl Into<String>) -> Self {
        Self {
            name: name.into(),
            values: Vec::new(),
        }
    }

    /// Adds a value to the series.
    pub fn add_value(&mut self, value: f64) {
        self.values.push(value);
    }
}

/// Chart type identifier for combination charts.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum CombinationChartType {
    /// Bar chart.
    Bar,
    /// Column chart.
    Column,
    /// Line chart.
    Line,
    /// Area chart.
    Area,
    /// Scatter chart.
    Scatter,
}

impl CombinationChartType {
    /// Convert to XML element name.
    pub fn to_xml_element(self) -> &'static str {
        match self {
            Self::Bar => "c:barChart",
            Self::Column => "c:barChart",
            Self::Line => "c:lineChart",
            Self::Area => "c:areaChart",
            Self::Scatter => "c:scatterChart",
        }
    }
}

/// Combination chart properties.
///
/// Combination charts display multiple chart types in a single chart area.
/// Each series can have a different chart type.
#[derive(Debug, Clone, PartialEq)]
pub struct CombinationChart {
    /// Series with their associated chart types.
    pub series: Vec<CombinationSeries>,
}

impl CombinationChart {
    /// Creates a new combination chart.
    pub fn new() -> Self {
        Self { series: Vec::new() }
    }

    /// Adds a series with its chart type.
    pub fn add_series(&mut self, series: CombinationSeries) {
        self.series.push(series);
    }
}

impl Default for CombinationChart {
    fn default() -> Self {
        Self::new()
    }
}

/// A single series in a combination chart with its chart type.
#[derive(Debug, Clone, PartialEq)]
pub struct CombinationSeries {
    /// Series name.
    pub name: String,
    /// Data values.
    pub values: Vec<f64>,
    /// Chart type for this series.
    pub chart_type: CombinationChartType,
}

impl CombinationSeries {
    /// Creates a new combination series.
    pub fn new(name: impl Into<String>, chart_type: CombinationChartType) -> Self {
        Self {
            name: name.into(),
            values: Vec::new(),
            chart_type,
        }
    }

    /// Adds a value to the series.
    pub fn add_value(&mut self, value: f64) {
        self.values.push(value);
    }
}

// ── XML Parsing Functions ──

/// Parse bubble chart from XML reader.
pub fn parse_bubble_chart<R: IoRead>(
    _reader: &mut quick_xml::Reader<R>,
) -> Result<BubbleChart, Box<dyn std::error::Error>> {
    // Placeholder: Full parsing would iterate over c:bubbleChart element
    // and extract c:ser (series), c:bubbleScale, c:showNegBubbles, etc.
    Ok(BubbleChart::new())
}

/// Write bubble chart to XML writer.
pub fn write_bubble_chart<W: IoWrite>(
    _writer: &mut quick_xml::Writer<W>,
    _chart: &BubbleChart,
) -> Result<(), Box<dyn std::error::Error>> {
    // Placeholder: Full writing would generate c:bubbleChart element
    // with c:ser, c:bubbleScale, c:showNegBubbles, etc.
    Ok(())
}

/// Parse stock chart from XML reader.
pub fn parse_stock_chart<R: IoRead>(
    _reader: &mut quick_xml::Reader<R>,
) -> Result<StockChart, Box<dyn std::error::Error>> {
    // Placeholder: Full parsing would iterate over c:stockChart element
    // and extract c:ser (series), c:hiLowLines, c:upDownBars, etc.
    Ok(StockChart::new(StockChartStyle::HighLowClose))
}

/// Write stock chart to XML writer.
pub fn write_stock_chart<W: IoWrite>(
    _writer: &mut quick_xml::Writer<W>,
    _chart: &StockChart,
) -> Result<(), Box<dyn std::error::Error>> {
    // Placeholder: Full writing would generate c:stockChart element
    // with c:ser, c:hiLowLines, c:upDownBars, etc.
    Ok(())
}

/// Parse surface chart from XML reader.
pub fn parse_surface_chart<R: IoRead>(
    _reader: &mut quick_xml::Reader<R>,
) -> Result<SurfaceChart, Box<dyn std::error::Error>> {
    // Placeholder: Full parsing would iterate over c:surfaceChart element
    // and extract c:ser (series), c:wireframe, etc.
    Ok(SurfaceChart::new())
}

/// Write surface chart to XML writer.
pub fn write_surface_chart<W: IoWrite>(
    _writer: &mut quick_xml::Writer<W>,
    _chart: &SurfaceChart,
) -> Result<(), Box<dyn std::error::Error>> {
    // Placeholder: Full writing would generate c:surfaceChart element
    // with c:ser, c:wireframe, etc.
    Ok(())
}

/// Parse 3D bar/column chart from XML reader.
pub fn parse_bar3d_chart<R: IoRead>(
    _reader: &mut quick_xml::Reader<R>,
) -> Result<Bar3DChart, Box<dyn std::error::Error>> {
    // Placeholder: Full parsing would iterate over c:bar3DChart element
    // and extract c:ser (series), c:gapDepth, c:shape, c:barDir, etc.
    Ok(Bar3DChart::new(false))
}

/// Write 3D bar/column chart to XML writer.
pub fn write_bar3d_chart<W: IoWrite>(
    _writer: &mut quick_xml::Writer<W>,
    _chart: &Bar3DChart,
) -> Result<(), Box<dyn std::error::Error>> {
    // Placeholder: Full writing would generate c:bar3DChart element
    // with c:ser, c:gapDepth, c:shape, c:barDir, etc.
    Ok(())
}

/// Parse 3D pie chart from XML reader.
pub fn parse_pie3d_chart<R: IoRead>(
    _reader: &mut quick_xml::Reader<R>,
) -> Result<Pie3DChart, Box<dyn std::error::Error>> {
    // Placeholder: Full parsing would iterate over c:pie3DChart element
    // and extract c:ser (series), c:view3D with rotX/rotY, etc.
    Ok(Pie3DChart::new())
}

/// Write 3D pie chart to XML writer.
pub fn write_pie3d_chart<W: IoWrite>(
    _writer: &mut quick_xml::Writer<W>,
    _chart: &Pie3DChart,
) -> Result<(), Box<dyn std::error::Error>> {
    // Placeholder: Full writing would generate c:pie3DChart element
    // with c:ser, c:view3D with rotX/rotY, etc.
    Ok(())
}

/// Parse 3D line chart from XML reader.
pub fn parse_line3d_chart<R: IoRead>(
    _reader: &mut quick_xml::Reader<R>,
) -> Result<Line3DChart, Box<dyn std::error::Error>> {
    // Placeholder: Full parsing would iterate over c:line3DChart element
    // and extract c:ser (series), c:gapDepth, etc.
    Ok(Line3DChart::new())
}

/// Write 3D line chart to XML writer.
pub fn write_line3d_chart<W: IoWrite>(
    _writer: &mut quick_xml::Writer<W>,
    _chart: &Line3DChart,
) -> Result<(), Box<dyn std::error::Error>> {
    // Placeholder: Full writing would generate c:line3DChart element
    // with c:ser, c:gapDepth, etc.
    Ok(())
}

/// Parse 3D area chart from XML reader.
pub fn parse_area3d_chart<R: IoRead>(
    _reader: &mut quick_xml::Reader<R>,
) -> Result<Area3DChart, Box<dyn std::error::Error>> {
    // Placeholder: Full parsing would iterate over c:area3DChart element
    // and extract c:ser (series), c:gapDepth, etc.
    Ok(Area3DChart::new())
}

/// Write 3D area chart to XML writer.
pub fn write_area3d_chart<W: IoWrite>(
    _writer: &mut quick_xml::Writer<W>,
    _chart: &Area3DChart,
) -> Result<(), Box<dyn std::error::Error>> {
    // Placeholder: Full writing would generate c:area3DChart element
    // with c:ser, c:gapDepth, etc.
    Ok(())
}

/// Parse combination chart from XML reader.
pub fn parse_combination_chart<R: IoRead>(
    _reader: &mut quick_xml::Reader<R>,
) -> Result<CombinationChart, Box<dyn std::error::Error>> {
    // Placeholder: Full parsing would iterate over multiple chart type elements
    // in c:plotArea and extract series from each.
    Ok(CombinationChart::new())
}

/// Write combination chart to XML writer.
pub fn write_combination_chart<W: IoWrite>(
    _writer: &mut quick_xml::Writer<W>,
    _chart: &CombinationChart,
) -> Result<(), Box<dyn std::error::Error>> {
    // Placeholder: Full writing would generate multiple chart type elements
    // in c:plotArea, one for each series type.
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn bubble_chart_creation() {
        let mut chart = BubbleChart::new();
        assert!(chart.series.is_empty());
        assert!(!chart.show_negative_bubbles);

        chart.set_bubble_scale(150);
        assert_eq!(chart.bubble_scale, Some(150));

        let mut series = BubbleSeries::new("Test");
        series.add_point(1.0, 2.0, 3.0);
        chart.add_series(series);
        assert_eq!(chart.series.len(), 1);
    }

    #[test]
    fn stock_chart_creation() {
        let mut chart = StockChart::new(StockChartStyle::OpenHighLowClose);
        assert_eq!(chart.style, StockChartStyle::OpenHighLowClose);
        assert_eq!(chart.style.series_count(), 4);

        let mut series = StockSeries::new("High");
        series.add_value(100.0);
        chart.add_series(series);
        assert_eq!(chart.series.len(), 1);
    }

    #[test]
    fn surface_chart_creation() {
        let mut chart = SurfaceChart::new();
        assert!(chart.series.is_empty());
        assert!(!chart.wireframe);
        assert!(!chart.contour);

        let mut series = SurfaceSeries::new("Test");
        series.add_row(vec![1.0, 2.0, 3.0]);
        chart.add_series(series);
        assert_eq!(chart.series.len(), 1);
    }

    #[test]
    fn bar3d_chart_creation() {
        let mut chart = Bar3DChart::new(false);
        assert!(!chart.horizontal);
        assert_eq!(chart.shape, BarShape3D::Box);

        chart.set_gap_depth(200);
        assert_eq!(chart.gap_depth, Some(200));

        let mut series = Bar3DSeries::new("Test");
        series.add_value(42.0);
        chart.add_series(series);
        assert_eq!(chart.series.len(), 1);
    }

    #[test]
    fn pie3d_chart_creation() {
        let mut chart = Pie3DChart::new();
        assert!(chart.rotation.is_none());
        assert!(chart.elevation.is_none());

        chart.set_rotation(45);
        chart.set_elevation(30);
        assert_eq!(chart.rotation, Some(45));
        assert_eq!(chart.elevation, Some(30));

        let mut series = Pie3DSeries::new("Test");
        series.add_point("Q1", 10.0);
        chart.add_series(series);
        assert_eq!(chart.series.len(), 1);
    }

    #[test]
    fn line3d_chart_creation() {
        let mut chart = Line3DChart::new();
        assert!(chart.gap_depth.is_none());

        chart.set_gap_depth(100);
        assert_eq!(chart.gap_depth, Some(100));

        let mut series = Line3DSeries::new("Test");
        series.add_value(42.0);
        chart.add_series(series);
        assert_eq!(chart.series.len(), 1);
    }

    #[test]
    fn area3d_chart_creation() {
        let mut chart = Area3DChart::new();
        assert!(chart.gap_depth.is_none());

        chart.set_gap_depth(150);
        assert_eq!(chart.gap_depth, Some(150));

        let mut series = Area3DSeries::new("Test");
        series.add_value(42.0);
        chart.add_series(series);
        assert_eq!(chart.series.len(), 1);
    }

    #[test]
    fn combination_chart_creation() {
        let mut chart = CombinationChart::new();
        assert!(chart.series.is_empty());

        let mut series1 = CombinationSeries::new("Revenue", CombinationChartType::Column);
        series1.add_value(100.0);
        chart.add_series(series1);

        let mut series2 = CombinationSeries::new("Trend", CombinationChartType::Line);
        series2.add_value(50.0);
        chart.add_series(series2);

        assert_eq!(chart.series.len(), 2);
        assert_eq!(chart.series[0].chart_type, CombinationChartType::Column);
        assert_eq!(chart.series[1].chart_type, CombinationChartType::Line);
    }

    #[test]
    fn bar_shape_3d_xml_roundtrip() {
        assert_eq!(BarShape3D::Box.to_xml(), "box");
        assert_eq!(BarShape3D::from_xml("box"), Some(BarShape3D::Box));
        assert_eq!(BarShape3D::Cylinder.to_xml(), "cylinder");
        assert_eq!(BarShape3D::from_xml("cylinder"), Some(BarShape3D::Cylinder));
        assert_eq!(BarShape3D::Cone.to_xml(), "cone");
        assert_eq!(BarShape3D::from_xml("cone"), Some(BarShape3D::Cone));
        assert_eq!(BarShape3D::Pyramid.to_xml(), "pyramid");
        assert_eq!(BarShape3D::from_xml("pyramid"), Some(BarShape3D::Pyramid));
        assert_eq!(BarShape3D::from_xml("invalid"), None);
    }

    #[test]
    fn stock_chart_style_series_count() {
        assert_eq!(StockChartStyle::HighLowClose.series_count(), 3);
        assert_eq!(StockChartStyle::OpenHighLowClose.series_count(), 4);
        assert_eq!(StockChartStyle::VolumeHighLowClose.series_count(), 4);
        assert_eq!(StockChartStyle::VolumeOpenHighLowClose.series_count(), 5);
    }
}
