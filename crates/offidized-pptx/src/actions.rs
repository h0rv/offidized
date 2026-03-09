/// Action type for shape click/hover events in PowerPoint.
///
/// Maps to `<a:hlinkClick>` or `<a:hlinkHover>` on shape non-visual properties.
/// PowerPoint uses `ppaction://` URIs for internal navigation actions.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum ActionType {
    /// Navigate to an external URL (e.g., "https://example.com").
    Hyperlink(String),
    /// Jump to the next slide (`ppaction://hlinkshowjump?jump=nextslide`).
    NextSlide,
    /// Jump to the previous slide (`ppaction://hlinkshowjump?jump=previousslide`).
    PreviousSlide,
    /// Jump to the first slide (`ppaction://hlinkshowjump?jump=firstslide`).
    FirstSlide,
    /// Jump to the last slide (`ppaction://hlinkshowjump?jump=lastslide`).
    LastSlide,
    /// Jump to a specific slide by relationship ID (`ppaction://hlinksldjump`).
    SlideJump(String),
    /// End the slide show (`ppaction://hlinkshowjump?jump=endshow`).
    EndShow,
    /// Play a sound (relationship ID to embedded sound).
    PlaySound(String),
    /// Run a program/macro (`ppaction://program` or `ppaction://macro`).
    RunProgram(String),
    /// Run an OLE action (`ppaction://ole`).
    OleAction(String),
    /// Custom show (`ppaction://customshow`).
    CustomShow(String),
}

impl ActionType {
    /// Parse from OOXML `action` attribute value and optional relationship target.
    pub fn from_xml(action: &str, r_id_target: Option<&str>) -> Self {
        match action {
            "ppaction://hlinkshowjump?jump=nextslide" => Self::NextSlide,
            "ppaction://hlinkshowjump?jump=previousslide" => Self::PreviousSlide,
            "ppaction://hlinkshowjump?jump=firstslide" => Self::FirstSlide,
            "ppaction://hlinkshowjump?jump=lastslide" => Self::LastSlide,
            "ppaction://hlinkshowjump?jump=endshow" => Self::EndShow,
            s if s.starts_with("ppaction://hlinksldjump") => {
                Self::SlideJump(r_id_target.unwrap_or("").to_string())
            }
            s if s.starts_with("ppaction://program") => {
                Self::RunProgram(r_id_target.unwrap_or("").to_string())
            }
            s if s.starts_with("ppaction://macro") => {
                Self::RunProgram(r_id_target.unwrap_or("").to_string())
            }
            s if s.starts_with("ppaction://ole") => {
                Self::OleAction(r_id_target.unwrap_or("").to_string())
            }
            s if s.starts_with("ppaction://customshow") => {
                Self::CustomShow(r_id_target.unwrap_or("").to_string())
            }
            _ => Self::Hyperlink(r_id_target.unwrap_or(action).to_string()),
        }
    }

    /// Convert to OOXML action attribute value.
    pub fn to_xml_action(&self) -> &str {
        match self {
            Self::Hyperlink(_) => "",
            Self::NextSlide => "ppaction://hlinkshowjump?jump=nextslide",
            Self::PreviousSlide => "ppaction://hlinkshowjump?jump=previousslide",
            Self::FirstSlide => "ppaction://hlinkshowjump?jump=firstslide",
            Self::LastSlide => "ppaction://hlinkshowjump?jump=lastslide",
            Self::EndShow => "ppaction://hlinkshowjump?jump=endshow",
            Self::SlideJump(_) => "ppaction://hlinksldjump",
            Self::PlaySound(_) => "",
            Self::RunProgram(_) => "ppaction://program",
            Self::OleAction(_) => "ppaction://ole",
            Self::CustomShow(_) => "ppaction://customshow",
        }
    }

    /// Whether this action needs a relationship ID (`r:id` attribute).
    pub fn needs_relationship(&self) -> bool {
        matches!(
            self,
            Self::Hyperlink(_)
                | Self::SlideJump(_)
                | Self::PlaySound(_)
                | Self::RunProgram(_)
                | Self::OleAction(_)
                | Self::CustomShow(_)
        )
    }
}

/// Shape action settings for click and hover events.
///
/// Maps to `<a:hlinkClick>` and `<a:hlinkHover>` in shape non-visual properties.
#[derive(Debug, Clone, PartialEq, Eq, Default)]
pub struct ShapeAction {
    /// Action triggered by clicking the shape.
    click_action: Option<ActionType>,
    /// Tooltip shown on hover.
    click_tooltip: Option<String>,
    /// Relationship ID for click action.
    click_rid: Option<String>,
    /// Action triggered by hovering over the shape.
    hover_action: Option<ActionType>,
    /// Tooltip shown on hover.
    hover_tooltip: Option<String>,
    /// Relationship ID for hover action.
    hover_rid: Option<String>,
}

impl ShapeAction {
    /// Creates empty shape action settings.
    pub fn new() -> Self {
        Self::default()
    }

    /// Whether any action is configured.
    pub fn has_actions(&self) -> bool {
        self.click_action.is_some() || self.hover_action.is_some()
    }

    // ── Click action ──

    /// Gets the click action.
    pub fn click_action(&self) -> Option<&ActionType> {
        self.click_action.as_ref()
    }

    /// Sets the click action.
    pub fn set_click_action(&mut self, action: ActionType) {
        self.click_action = Some(action);
    }

    /// Clears the click action.
    pub fn clear_click_action(&mut self) {
        self.click_action = None;
        self.click_tooltip = None;
        self.click_rid = None;
    }

    /// Gets the click tooltip.
    pub fn click_tooltip(&self) -> Option<&str> {
        self.click_tooltip.as_deref()
    }

    /// Sets the click tooltip.
    pub fn set_click_tooltip(&mut self, tooltip: impl Into<String>) {
        self.click_tooltip = Some(tooltip.into());
    }

    /// Gets the click relationship ID.
    pub fn click_rid(&self) -> Option<&str> {
        self.click_rid.as_deref()
    }

    /// Sets the click relationship ID.
    pub fn set_click_rid(&mut self, rid: impl Into<String>) {
        self.click_rid = Some(rid.into());
    }

    // ── Hover action ──

    /// Gets the hover action.
    pub fn hover_action(&self) -> Option<&ActionType> {
        self.hover_action.as_ref()
    }

    /// Sets the hover action.
    pub fn set_hover_action(&mut self, action: ActionType) {
        self.hover_action = Some(action);
    }

    /// Clears the hover action.
    pub fn clear_hover_action(&mut self) {
        self.hover_action = None;
        self.hover_tooltip = None;
        self.hover_rid = None;
    }

    /// Gets the hover tooltip.
    pub fn hover_tooltip(&self) -> Option<&str> {
        self.hover_tooltip.as_deref()
    }

    /// Sets the hover tooltip.
    pub fn set_hover_tooltip(&mut self, tooltip: impl Into<String>) {
        self.hover_tooltip = Some(tooltip.into());
    }

    /// Gets the hover relationship ID.
    pub fn hover_rid(&self) -> Option<&str> {
        self.hover_rid.as_deref()
    }

    /// Sets the hover relationship ID.
    pub fn set_hover_rid(&mut self, rid: impl Into<String>) {
        self.hover_rid = Some(rid.into());
    }

    // ── Convenience builders ──

    /// Creates a click-hyperlink action.
    pub fn with_click_hyperlink(mut self, url: impl Into<String>) -> Self {
        self.click_action = Some(ActionType::Hyperlink(url.into()));
        self
    }

    /// Creates a click-next-slide action.
    pub fn with_click_next_slide(mut self) -> Self {
        self.click_action = Some(ActionType::NextSlide);
        self
    }

    /// Creates a click-previous-slide action.
    pub fn with_click_previous_slide(mut self) -> Self {
        self.click_action = Some(ActionType::PreviousSlide);
        self
    }

    /// Creates a click-end-show action.
    pub fn with_click_end_show(mut self) -> Self {
        self.click_action = Some(ActionType::EndShow);
        self
    }

    /// Sets the click tooltip via builder.
    pub fn with_click_tooltip(mut self, tooltip: impl Into<String>) -> Self {
        self.click_tooltip = Some(tooltip.into());
        self
    }
}

/// Action button preset type.
///
/// Standard action button shapes defined in the OOXML spec.
/// These are preset shapes with built-in action behaviors.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ActionButtonType {
    /// Back or Previous button.
    BackOrPrevious,
    /// Forward or Next button.
    ForwardOrNext,
    /// Beginning/Home button.
    Beginning,
    /// End button.
    End,
    /// Home button.
    Home,
    /// Information button.
    Information,
    /// Return button.
    Return,
    /// Movie/video button.
    Movie,
    /// Document button.
    Document,
    /// Sound button.
    Sound,
    /// Help button.
    Help,
    /// Blank (custom) action button.
    Blank,
}

impl ActionButtonType {
    /// Parse from OOXML preset shape name.
    pub fn from_xml(name: &str) -> Option<Self> {
        match name {
            "actionButtonBackPrevious" => Some(Self::BackOrPrevious),
            "actionButtonForwardNext" => Some(Self::ForwardOrNext),
            "actionButtonBeginning" => Some(Self::Beginning),
            "actionButtonEnd" => Some(Self::End),
            "actionButtonHome" => Some(Self::Home),
            "actionButtonInformation" => Some(Self::Information),
            "actionButtonReturn" => Some(Self::Return),
            "actionButtonMovie" => Some(Self::Movie),
            "actionButtonDocument" => Some(Self::Document),
            "actionButtonSound" => Some(Self::Sound),
            "actionButtonHelp" => Some(Self::Help),
            "actionButtonBlank" => Some(Self::Blank),
            _ => None,
        }
    }

    /// Convert to OOXML preset shape name.
    pub fn to_xml(self) -> &'static str {
        match self {
            Self::BackOrPrevious => "actionButtonBackPrevious",
            Self::ForwardOrNext => "actionButtonForwardNext",
            Self::Beginning => "actionButtonBeginning",
            Self::End => "actionButtonEnd",
            Self::Home => "actionButtonHome",
            Self::Information => "actionButtonInformation",
            Self::Return => "actionButtonReturn",
            Self::Movie => "actionButtonMovie",
            Self::Document => "actionButtonDocument",
            Self::Sound => "actionButtonSound",
            Self::Help => "actionButtonHelp",
            Self::Blank => "actionButtonBlank",
        }
    }
}

/// Embedded OLE object in a slide.
///
/// Maps to `<p:oleObj>` in OOXML. Represents an embedded or linked
/// external object (Excel spreadsheet, Word document, PDF, etc.).
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct EmbeddedObject {
    /// Display name for the embedded object.
    name: String,
    /// Program ID (e.g., "Excel.Sheet.12", "Word.Document.12", "AcroExch.Document").
    prog_id: Option<String>,
    /// Relationship ID to the embedded object data.
    relationship_id: Option<String>,
    /// Whether the object is linked (vs. embedded).
    is_linked: bool,
    /// Icon representation (vs. inline rendering).
    show_as_icon: bool,
    /// Image relationship ID for the object's visual representation.
    image_rid: Option<String>,
}

impl EmbeddedObject {
    /// Creates a new embedded object.
    pub fn new(name: impl Into<String>) -> Self {
        Self {
            name: name.into(),
            prog_id: None,
            relationship_id: None,
            is_linked: false,
            show_as_icon: false,
            image_rid: None,
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

    /// Gets the program ID (e.g., "Excel.Sheet.12").
    pub fn prog_id(&self) -> Option<&str> {
        self.prog_id.as_deref()
    }

    /// Sets the program ID.
    pub fn set_prog_id(&mut self, prog_id: impl Into<String>) {
        self.prog_id = Some(prog_id.into());
    }

    /// Gets the relationship ID to the embedded data.
    pub fn relationship_id(&self) -> Option<&str> {
        self.relationship_id.as_deref()
    }

    /// Sets the relationship ID.
    pub fn set_relationship_id(&mut self, rid: impl Into<String>) {
        self.relationship_id = Some(rid.into());
    }

    /// Whether the object is linked (vs. embedded).
    pub fn is_linked(&self) -> bool {
        self.is_linked
    }

    /// Sets whether the object is linked.
    pub fn set_linked(&mut self, linked: bool) {
        self.is_linked = linked;
    }

    /// Whether the object shows as an icon.
    pub fn show_as_icon(&self) -> bool {
        self.show_as_icon
    }

    /// Sets whether to show as icon.
    pub fn set_show_as_icon(&mut self, icon: bool) {
        self.show_as_icon = icon;
    }

    /// Gets the image relationship ID for the visual representation.
    pub fn image_rid(&self) -> Option<&str> {
        self.image_rid.as_deref()
    }

    /// Sets the image relationship ID.
    pub fn set_image_rid(&mut self, rid: impl Into<String>) {
        self.image_rid = Some(rid.into());
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn action_type_navigation_roundtrip() {
        let cases = [
            (
                "ppaction://hlinkshowjump?jump=nextslide",
                ActionType::NextSlide,
            ),
            (
                "ppaction://hlinkshowjump?jump=previousslide",
                ActionType::PreviousSlide,
            ),
            (
                "ppaction://hlinkshowjump?jump=firstslide",
                ActionType::FirstSlide,
            ),
            (
                "ppaction://hlinkshowjump?jump=lastslide",
                ActionType::LastSlide,
            ),
            ("ppaction://hlinkshowjump?jump=endshow", ActionType::EndShow),
        ];

        for (xml, expected) in &cases {
            let parsed = ActionType::from_xml(xml, None);
            assert_eq!(&parsed, expected);
            assert_eq!(parsed.to_xml_action(), *xml);
        }
    }

    #[test]
    fn action_type_slide_jump() {
        let action = ActionType::from_xml("ppaction://hlinksldjump", Some("slide3.xml"));
        assert_eq!(action, ActionType::SlideJump("slide3.xml".to_string()));
        assert_eq!(action.to_xml_action(), "ppaction://hlinksldjump");
        assert!(action.needs_relationship());
    }

    #[test]
    fn action_type_hyperlink() {
        let action = ActionType::from_xml("", Some("https://example.com"));
        assert_eq!(
            action,
            ActionType::Hyperlink("https://example.com".to_string())
        );
        assert!(action.needs_relationship());
    }

    #[test]
    fn shape_action_builder() {
        let action = ShapeAction::new()
            .with_click_next_slide()
            .with_click_tooltip("Go to next");

        assert!(action.has_actions());
        assert_eq!(action.click_action(), Some(&ActionType::NextSlide));
        assert_eq!(action.click_tooltip(), Some("Go to next"));
        assert!(action.hover_action().is_none());
    }

    #[test]
    fn shape_action_clear() {
        let mut action = ShapeAction::new().with_click_hyperlink("https://example.com");
        assert!(action.has_actions());

        action.clear_click_action();
        assert!(!action.has_actions());
    }

    #[test]
    fn action_button_type_roundtrip() {
        let types = [
            ("actionButtonBackPrevious", ActionButtonType::BackOrPrevious),
            ("actionButtonForwardNext", ActionButtonType::ForwardOrNext),
            ("actionButtonHome", ActionButtonType::Home),
            ("actionButtonEnd", ActionButtonType::End),
            ("actionButtonBlank", ActionButtonType::Blank),
        ];

        for (xml, expected) in &types {
            assert_eq!(ActionButtonType::from_xml(xml), Some(*expected));
            assert_eq!(expected.to_xml(), *xml);
        }

        assert_eq!(ActionButtonType::from_xml("notAnActionButton"), None);
    }

    #[test]
    fn embedded_object_basic() {
        let mut obj = EmbeddedObject::new("Budget.xlsx");
        assert_eq!(obj.name(), "Budget.xlsx");
        assert!(!obj.is_linked());
        assert!(!obj.show_as_icon());

        obj.set_prog_id("Excel.Sheet.12");
        obj.set_relationship_id("rId5");
        obj.set_linked(true);
        obj.set_show_as_icon(true);

        assert_eq!(obj.prog_id(), Some("Excel.Sheet.12"));
        assert_eq!(obj.relationship_id(), Some("rId5"));
        assert!(obj.is_linked());
        assert!(obj.show_as_icon());
    }
}
