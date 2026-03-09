//! Chart templates and themes for applying consistent styling to charts.
//!
//! This module provides a template system for charts with preset color schemes,
//! styles, and configurations. Templates can be applied to charts to ensure
//! consistent visual appearance across presentations.

use crate::chart::{
    Chart, ChartDataLabel, ChartSeries, LineStyle, MarkerShape, SeriesBorder, SeriesFill,
    SeriesMarker,
};

/// Visual style theme for charts.
///
/// Defines the overall aesthetic approach for chart rendering, including
/// color schemes, line weights, and visual effects.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum ChartStyle {
    /// Modern flat design with bright, saturated colors.
    /// Uses bold colors and minimal borders for a clean look.
    #[default]
    Modern,
    /// Traditional office style with muted colors and prominent borders.
    /// More conservative color palette suitable for professional documents.
    Classic,
    /// Vibrant multi-color palette with high contrast.
    /// Uses a wide variety of colors for maximum visual distinction.
    Colorful,
    /// Grayscale theme with varying shades of gray.
    /// Suitable for black-and-white printing or accessibility.
    Monochrome,
}

impl ChartStyle {
    /// Returns the color palette for this style.
    ///
    /// Each style has a predefined set of colors used for series.
    /// Colors are returned as hex strings without the '#' prefix.
    pub fn color_palette(self) -> &'static [&'static str] {
        match self {
            Self::Modern => &[
                "4472C4", // Blue
                "ED7D31", // Orange
                "A5A5A5", // Gray
                "FFC000", // Yellow
                "5B9BD5", // Light Blue
                "70AD47", // Green
            ],
            Self::Classic => &[
                "17375E", // Dark Blue
                "BE4B48", // Dark Red
                "76933C", // Olive Green
                "8E5D9F", // Purple
                "C46100", // Dark Orange
                "5F7186", // Slate Gray
            ],
            Self::Colorful => &[
                "FF6B6B", // Red
                "4ECDC4", // Turquoise
                "45B7D1", // Blue
                "FFA07A", // Light Salmon
                "98D8C8", // Mint
                "F7DC6F", // Yellow
                "BB8FCE", // Lavender
                "F8B739", // Gold
            ],
            Self::Monochrome => &[
                "2C2C2C", // Very Dark Gray
                "545454", // Dark Gray
                "7C7C7C", // Medium Gray
                "A4A4A4", // Light Gray
                "CCCCCC", // Very Light Gray
                "E0E0E0", // Near White Gray
            ],
        }
    }

    /// Returns the default border width for series in this style.
    pub fn border_width(self) -> f64 {
        match self {
            Self::Modern => 0.5,
            Self::Classic => 1.5,
            Self::Colorful => 1.0,
            Self::Monochrome => 1.0,
        }
    }

    /// Returns the default line width for line charts in this style.
    pub fn line_width(self) -> f64 {
        match self {
            Self::Modern => 2.5,
            Self::Classic => 2.0,
            Self::Colorful => 2.5,
            Self::Monochrome => 2.0,
        }
    }

    /// Returns whether this style uses smooth lines by default.
    pub fn use_smooth_lines(self) -> bool {
        match self {
            Self::Modern => true,
            Self::Classic => false,
            Self::Colorful => true,
            Self::Monochrome => false,
        }
    }

    /// Returns the default marker shape for this style.
    pub fn marker_shape(self) -> MarkerShape {
        match self {
            Self::Modern => MarkerShape::Circle,
            Self::Classic => MarkerShape::Square,
            Self::Colorful => MarkerShape::Circle,
            Self::Monochrome => MarkerShape::Diamond,
        }
    }

    /// Returns the default marker size for this style.
    pub fn marker_size(self) -> u32 {
        match self {
            Self::Modern => 6,
            Self::Classic => 5,
            Self::Colorful => 7,
            Self::Monochrome => 5,
        }
    }
}

/// Configuration template for chart styling.
///
/// Provides a complete set of preset configurations that can be applied
/// to a chart to achieve a consistent visual style.
#[derive(Debug, Clone, PartialEq)]
pub struct ChartTemplate {
    /// Name of this template.
    pub name: String,
    /// Visual style for this template.
    pub style: ChartStyle,
    /// Whether to show legend by default.
    pub show_legend: bool,
    /// Whether to show gridlines by default.
    pub show_gridlines: bool,
    /// Whether to apply data labels by default.
    pub show_data_labels: bool,
    /// Custom color palette (if specified, overrides style default).
    pub custom_colors: Option<Vec<String>>,
}

impl ChartTemplate {
    /// Creates a new template with the given name and style.
    pub fn new(name: impl Into<String>, style: ChartStyle) -> Self {
        Self {
            name: name.into(),
            style,
            show_legend: true,
            show_gridlines: true,
            show_data_labels: false,
            custom_colors: None,
        }
    }

    /// Sets whether to show legend.
    pub fn with_legend(mut self, show: bool) -> Self {
        self.show_legend = show;
        self
    }

    /// Sets whether to show gridlines.
    pub fn with_gridlines(mut self, show: bool) -> Self {
        self.show_gridlines = show;
        self
    }

    /// Sets whether to show data labels.
    pub fn with_data_labels(mut self, show: bool) -> Self {
        self.show_data_labels = show;
        self
    }

    /// Sets a custom color palette for this template.
    pub fn with_custom_colors(mut self, colors: Vec<String>) -> Self {
        self.custom_colors = Some(colors);
        self
    }

    /// Gets the color palette for this template.
    ///
    /// Returns custom colors if set, otherwise returns the style's default palette.
    pub fn colors(&self) -> Vec<String> {
        if let Some(ref custom) = self.custom_colors {
            custom.clone()
        } else {
            self.style
                .color_palette()
                .iter()
                .map(|s| s.to_string())
                .collect()
        }
    }

    /// Applies this template to a chart.
    ///
    /// This configures the chart's visual style according to the template's
    /// settings, including colors, borders, markers, and axes.
    pub fn apply_to_chart(&self, chart: &mut Chart) {
        // Apply legend settings
        chart.set_show_legend(self.show_legend);

        // Apply gridlines to axes
        if self.show_gridlines {
            let mut cat_axis = chart.category_axis().cloned().unwrap_or_default();
            cat_axis.has_major_gridlines = true;
            chart.set_category_axis(cat_axis);

            let mut val_axis = chart.value_axis().cloned().unwrap_or_default();
            val_axis.has_major_gridlines = true;
            chart.set_value_axis(val_axis);
        }

        // Apply data labels
        if self.show_data_labels {
            let labels = ChartDataLabel::new()
                .with_value(true)
                .with_position("bestFit".to_string());
            chart.set_data_labels(labels);
        }

        // Apply styling to existing series
        let colors = self.colors();
        for (i, series) in chart.additional_series().iter().enumerate() {
            let mut styled_series = series.clone();
            self.apply_to_series(&mut styled_series, i, &colors);
        }
    }

    /// Applies this template's styling to a chart series.
    ///
    /// # Arguments
    /// * `series` - The series to style
    /// * `index` - The series index (used for color selection)
    /// * `colors` - Color palette to use
    pub fn apply_to_series(&self, series: &mut ChartSeries, index: usize, colors: &[String]) {
        let color = &colors[index % colors.len()];

        // Apply fill color
        series.set_fill(SeriesFill::solid(color));

        // Apply border
        let border = SeriesBorder::new()
            .with_color(color)
            .with_width(self.style.border_width());
        series.set_border(border);

        // Apply line style for line/scatter charts
        let line = LineStyle::new()
            .with_color(color)
            .with_width(self.style.line_width())
            .with_smooth(self.style.use_smooth_lines());
        series.set_line_style(line);

        // Apply marker for line/scatter charts
        let marker = SeriesMarker::new(self.style.marker_shape())
            .with_size(self.style.marker_size())
            .with_fill(SeriesFill::solid(color));
        series.set_marker(marker);
    }

    /// Creates a new chart series with this template's styling.
    ///
    /// # Arguments
    /// * `name` - Name for the series
    /// * `index` - Series index (used for color selection)
    pub fn create_series(&self, name: impl Into<String>, index: usize) -> ChartSeries {
        let mut series = ChartSeries::new(name);
        let colors = self.colors();
        self.apply_to_series(&mut series, index, &colors);
        series
    }
}

impl Default for ChartTemplate {
    fn default() -> Self {
        Self::new("Default", ChartStyle::default())
    }
}

/// Registry of chart templates.
///
/// Provides a collection of named templates that can be retrieved and
/// applied to charts. Includes built-in templates for common styles.
#[derive(Debug, Clone, Default)]
pub struct ChartTemplateRegistry {
    templates: Vec<ChartTemplate>,
}

impl ChartTemplateRegistry {
    /// Creates a new empty registry.
    pub fn new() -> Self {
        Self {
            templates: Vec::new(),
        }
    }

    /// Creates a registry with built-in templates.
    ///
    /// Includes templates for Modern, Classic, Colorful, and Monochrome styles.
    pub fn with_builtin_templates() -> Self {
        let mut registry = Self::new();

        registry.register(
            ChartTemplate::new("Modern", ChartStyle::Modern)
                .with_legend(true)
                .with_gridlines(true),
        );

        registry.register(
            ChartTemplate::new("Classic", ChartStyle::Classic)
                .with_legend(true)
                .with_gridlines(true),
        );

        registry.register(
            ChartTemplate::new("Colorful", ChartStyle::Colorful)
                .with_legend(true)
                .with_gridlines(false),
        );

        registry.register(
            ChartTemplate::new("Monochrome", ChartStyle::Monochrome)
                .with_legend(true)
                .with_gridlines(true),
        );

        registry.register(
            ChartTemplate::new("Minimal", ChartStyle::Modern)
                .with_legend(false)
                .with_gridlines(false)
                .with_data_labels(true),
        );

        registry
    }

    /// Registers a new template in the registry.
    pub fn register(&mut self, template: ChartTemplate) {
        self.templates.push(template);
    }

    /// Retrieves a template by name.
    ///
    /// Returns `None` if no template with the given name exists.
    pub fn get(&self, name: &str) -> Option<&ChartTemplate> {
        self.templates.iter().find(|t| t.name == name)
    }

    /// Returns all registered template names.
    pub fn template_names(&self) -> Vec<&str> {
        self.templates.iter().map(|t| t.name.as_str()).collect()
    }

    /// Returns the number of registered templates.
    pub fn len(&self) -> usize {
        self.templates.len()
    }

    /// Returns whether the registry is empty.
    pub fn is_empty(&self) -> bool {
        self.templates.is_empty()
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn chart_style_color_palettes() {
        assert_eq!(ChartStyle::Modern.color_palette().len(), 6);
        assert_eq!(ChartStyle::Classic.color_palette().len(), 6);
        assert_eq!(ChartStyle::Colorful.color_palette().len(), 8);
        assert_eq!(ChartStyle::Monochrome.color_palette().len(), 6);
    }

    #[test]
    fn chart_style_properties() {
        assert_eq!(ChartStyle::Modern.border_width(), 0.5);
        assert_eq!(ChartStyle::Classic.border_width(), 1.5);
        assert!(ChartStyle::Modern.use_smooth_lines());
        assert!(!ChartStyle::Classic.use_smooth_lines());
    }

    #[test]
    fn template_creation() {
        let template = ChartTemplate::new("Test", ChartStyle::Modern)
            .with_legend(false)
            .with_gridlines(true)
            .with_data_labels(true);

        assert_eq!(template.name, "Test");
        assert_eq!(template.style, ChartStyle::Modern);
        assert!(!template.show_legend);
        assert!(template.show_gridlines);
        assert!(template.show_data_labels);
    }

    #[test]
    fn template_custom_colors() {
        let custom = vec!["FF0000".to_string(), "00FF00".to_string()];
        let template =
            ChartTemplate::new("Custom", ChartStyle::Modern).with_custom_colors(custom.clone());

        assert_eq!(template.colors(), custom);
    }

    #[test]
    fn template_default_colors() {
        let template = ChartTemplate::new("Test", ChartStyle::Colorful);
        let colors = template.colors();

        assert_eq!(colors.len(), 8);
        assert_eq!(colors[0], "FF6B6B");
    }

    #[test]
    fn template_apply_to_chart() {
        let mut chart = Chart::new("Test Chart");
        let template = ChartTemplate::new("Test", ChartStyle::Modern)
            .with_legend(true)
            .with_gridlines(true);

        template.apply_to_chart(&mut chart);

        assert!(chart.show_legend());
        assert!(chart.category_axis().unwrap().has_major_gridlines);
        assert!(chart.value_axis().unwrap().has_major_gridlines);
    }

    #[test]
    fn template_create_series() {
        let template = ChartTemplate::new("Test", ChartStyle::Modern);
        let series = template.create_series("Series 1", 0);

        assert_eq!(series.name(), "Series 1");
        assert!(series.fill().is_some());
        assert!(series.border().is_some());
        assert!(series.line_style().is_some());
        assert!(series.marker().is_some());
    }

    #[test]
    fn registry_builtin_templates() {
        let registry = ChartTemplateRegistry::with_builtin_templates();

        assert_eq!(registry.len(), 5);
        assert!(registry.get("Modern").is_some());
        assert!(registry.get("Classic").is_some());
        assert!(registry.get("Colorful").is_some());
        assert!(registry.get("Monochrome").is_some());
        assert!(registry.get("Minimal").is_some());
    }

    #[test]
    fn registry_register_and_retrieve() {
        let mut registry = ChartTemplateRegistry::new();
        let template = ChartTemplate::new("Custom", ChartStyle::Modern);

        registry.register(template);

        assert_eq!(registry.len(), 1);
        assert!(registry.get("Custom").is_some());
        assert_eq!(registry.get("Custom").unwrap().name, "Custom");
    }

    #[test]
    fn registry_template_names() {
        let registry = ChartTemplateRegistry::with_builtin_templates();
        let names = registry.template_names();

        assert_eq!(names.len(), 5);
        assert!(names.contains(&"Modern"));
        assert!(names.contains(&"Classic"));
        assert!(names.contains(&"Colorful"));
        assert!(names.contains(&"Monochrome"));
        assert!(names.contains(&"Minimal"));
    }

    #[test]
    fn registry_is_empty() {
        let registry = ChartTemplateRegistry::new();
        assert!(registry.is_empty());

        let registry = ChartTemplateRegistry::with_builtin_templates();
        assert!(!registry.is_empty());
    }
}
