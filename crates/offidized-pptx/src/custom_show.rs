/// Custom show (named subset of slides for a specific audience).
///
/// Maps to `p:custShow` in OOXML. A custom show allows a presenter
/// to define a subset and ordering of slides to show, separate from
/// the default presentation order.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct CustomShow {
    /// Display name of the custom show.
    name: String,
    /// Unique identifier for this custom show.
    id: u32,
    /// Ordered list of slide relationship IDs (`rId` values) included in this show.
    slide_ids: Vec<String>,
}

impl CustomShow {
    /// Creates a new custom show with the given name and ID.
    pub fn new(name: impl Into<String>, id: u32) -> Self {
        Self {
            name: name.into(),
            id,
            slide_ids: Vec::new(),
        }
    }

    /// Gets the display name.
    pub fn name(&self) -> &str {
        &self.name
    }

    /// Sets the display name.
    pub fn set_name(&mut self, name: impl Into<String>) {
        self.name = name.into();
    }

    /// Gets the unique ID.
    pub fn id(&self) -> u32 {
        self.id
    }

    /// Gets the ordered list of slide relationship IDs.
    pub fn slide_ids(&self) -> &[String] {
        &self.slide_ids
    }

    /// Adds a slide relationship ID to the show.
    pub fn add_slide_id(&mut self, rid: impl Into<String>) {
        self.slide_ids.push(rid.into());
    }

    /// Removes all occurrences of a slide relationship ID from the show.
    /// Used when a slide is deleted from the presentation.
    pub fn remove_slide_id(&mut self, rid: &str) {
        self.slide_ids.retain(|id| id != rid);
    }

    /// Clears all slide IDs from the show.
    pub fn clear_slide_ids(&mut self) {
        self.slide_ids.clear();
    }

    /// Returns the number of slides in this custom show.
    pub fn slide_count(&self) -> usize {
        self.slide_ids.len()
    }

    /// Sets the ordered list of slide relationship IDs.
    pub fn set_slide_ids(&mut self, ids: Vec<String>) {
        self.slide_ids = ids;
    }
}

/// Slide show type (how the presentation is displayed).
///
/// Maps to the child element of `p:showPr` in OOXML.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum SlideShowType {
    /// Standard presentation shown by a speaker (default).
    /// Maps to `p:present`.
    #[default]
    Present,
    /// Browsed by an individual in a window.
    /// Maps to `p:browse`.
    Browse,
    /// Displayed at a kiosk (full screen, loops automatically).
    /// Maps to `p:kiosk`.
    Kiosk,
}

impl SlideShowType {
    /// Parse from OOXML element name.
    pub fn from_xml(element_name: &str) -> Option<Self> {
        match element_name {
            "present" => Some(Self::Present),
            "browse" => Some(Self::Browse),
            "kiosk" => Some(Self::Kiosk),
            _ => None,
        }
    }

    /// Convert to OOXML element name.
    pub fn to_xml(self) -> &'static str {
        match self {
            Self::Present => "present",
            Self::Browse => "browse",
            Self::Kiosk => "kiosk",
        }
    }
}

/// Slide show settings for the presentation.
///
/// Maps to `p:showPr` (show properties) in OOXML.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct SlideShowSettings {
    /// How the slide show is displayed (present/browse/kiosk).
    show_type: SlideShowType,
    /// Whether to loop continuously until Esc is pressed.
    /// Maps to `loop` attribute.
    loop_continuously: bool,
    /// Whether to show narration during the slide show.
    /// Maps to `showNarration` attribute.
    show_narration: bool,
    /// Whether to show animations during the slide show.
    /// Maps to `showAnimation` attribute.
    show_animation: bool,
    /// Whether to use pen color.
    /// Maps to `usePenClr` attribute (the color value is stored separately).
    use_pen_color: bool,
    /// Pen color as sRGB hex (e.g., "FF0000" for red).
    pen_color: Option<String>,
    /// Name of the custom show to use, if any.
    /// When set, only the slides in the named custom show are displayed.
    custom_show_name: Option<String>,
}

impl SlideShowSettings {
    /// Creates default slide show settings.
    pub fn new() -> Self {
        Self {
            show_type: SlideShowType::Present,
            loop_continuously: false,
            show_narration: true,
            show_animation: true,
            use_pen_color: false,
            pen_color: None,
            custom_show_name: None,
        }
    }

    /// Gets the slide show type.
    pub fn show_type(&self) -> SlideShowType {
        self.show_type
    }

    /// Sets the slide show type.
    pub fn set_show_type(&mut self, show_type: SlideShowType) {
        self.show_type = show_type;
    }

    /// Whether the slide show loops continuously.
    pub fn loop_continuously(&self) -> bool {
        self.loop_continuously
    }

    /// Sets whether to loop continuously.
    pub fn set_loop_continuously(&mut self, loop_val: bool) {
        self.loop_continuously = loop_val;
    }

    /// Whether narration is shown.
    pub fn show_narration(&self) -> bool {
        self.show_narration
    }

    /// Sets whether to show narration.
    pub fn set_show_narration(&mut self, show: bool) {
        self.show_narration = show;
    }

    /// Whether animations are shown.
    pub fn show_animation(&self) -> bool {
        self.show_animation
    }

    /// Sets whether to show animations.
    pub fn set_show_animation(&mut self, show: bool) {
        self.show_animation = show;
    }

    /// Whether pen color is used.
    pub fn use_pen_color(&self) -> bool {
        self.use_pen_color
    }

    /// Gets the pen color (sRGB hex).
    pub fn pen_color(&self) -> Option<&str> {
        self.pen_color.as_deref()
    }

    /// Sets the pen color (sRGB hex, with or without #).
    pub fn set_pen_color(&mut self, color: impl Into<String>) {
        let mut color = color.into();
        if color.starts_with('#') {
            color = color[1..].to_string();
        }
        self.pen_color = Some(color);
        self.use_pen_color = true;
    }

    /// Clears the pen color.
    pub fn clear_pen_color(&mut self) {
        self.pen_color = None;
        self.use_pen_color = false;
    }

    /// Gets the custom show name to use for this slide show.
    pub fn custom_show_name(&self) -> Option<&str> {
        self.custom_show_name.as_deref()
    }

    /// Sets which custom show to use.
    pub fn set_custom_show_name(&mut self, name: impl Into<String>) {
        self.custom_show_name = Some(name.into());
    }

    /// Clears the custom show selection (uses all slides).
    pub fn clear_custom_show_name(&mut self) {
        self.custom_show_name = None;
    }
}

impl Default for SlideShowSettings {
    fn default() -> Self {
        Self::new()
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn custom_show_basic_operations() {
        let mut show = CustomShow::new("Sales Team", 1);
        assert_eq!(show.name(), "Sales Team");
        assert_eq!(show.id(), 1);
        assert_eq!(show.slide_count(), 0);

        show.add_slide_id("rId2");
        show.add_slide_id("rId5");
        show.add_slide_id("rId3");
        assert_eq!(show.slide_count(), 3);
        assert_eq!(show.slide_ids(), ["rId2", "rId5", "rId3"]);

        show.remove_slide_id("rId5");
        assert_eq!(show.slide_count(), 2);
        assert_eq!(show.slide_ids(), ["rId2", "rId3"]);
    }

    #[test]
    fn slide_show_type_xml_roundtrip() {
        assert_eq!(
            SlideShowType::from_xml("present"),
            Some(SlideShowType::Present)
        );
        assert_eq!(
            SlideShowType::from_xml("browse"),
            Some(SlideShowType::Browse)
        );
        assert_eq!(SlideShowType::from_xml("kiosk"), Some(SlideShowType::Kiosk));
        assert_eq!(SlideShowType::from_xml("unknown"), None);

        assert_eq!(SlideShowType::Present.to_xml(), "present");
        assert_eq!(SlideShowType::Browse.to_xml(), "browse");
        assert_eq!(SlideShowType::Kiosk.to_xml(), "kiosk");
    }

    #[test]
    fn slide_show_settings_defaults() {
        let settings = SlideShowSettings::new();
        assert_eq!(settings.show_type(), SlideShowType::Present);
        assert!(!settings.loop_continuously());
        assert!(settings.show_narration());
        assert!(settings.show_animation());
        assert!(!settings.use_pen_color());
        assert!(settings.pen_color().is_none());
        assert!(settings.custom_show_name().is_none());
    }

    #[test]
    fn slide_show_settings_pen_color() {
        let mut settings = SlideShowSettings::new();
        settings.set_pen_color("#FF0000");
        assert!(settings.use_pen_color());
        assert_eq!(settings.pen_color(), Some("FF0000"));

        settings.clear_pen_color();
        assert!(!settings.use_pen_color());
        assert!(settings.pen_color().is_none());
    }

    #[test]
    fn slide_show_settings_kiosk_mode() {
        let mut settings = SlideShowSettings::new();
        settings.set_show_type(SlideShowType::Kiosk);
        settings.set_loop_continuously(true);
        settings.set_show_narration(false);

        assert_eq!(settings.show_type(), SlideShowType::Kiosk);
        assert!(settings.loop_continuously());
        assert!(!settings.show_narration());
    }
}
