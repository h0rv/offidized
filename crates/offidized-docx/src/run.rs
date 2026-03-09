use crate::image::{FloatingImage, InlineImage};
use offidized_opc::RawXmlNode;

/// Underline style variants for a text run (`w:u w:val`).
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum UnderlineType {
    /// Single underline (`single`).
    Single,
    /// Double underline (`double`).
    Double,
    /// Thick underline (`thick`).
    Thick,
    /// Dotted underline (`dotted`).
    Dotted,
    /// Heavy dotted underline (`dottedHeavy`).
    DottedHeavy,
    /// Dashed underline (`dash`).
    Dash,
    /// Heavy dashed underline (`dashedHeavy`).
    DashedHeavy,
    /// Long-dash underline (`dashLong`).
    DashLong,
    /// Heavy long-dash underline (`dashLongHeavy`).
    DashLongHeavy,
    /// Dash-dot underline (`dotDash`).
    DashDot,
    /// Heavy dash-dot underline (`dashDotHeavy`).
    DashDotHeavy,
    /// Dash-dot-dot underline (`dotDotDash`).
    DashDotDot,
    /// Heavy dash-dot-dot underline (`dashDotDotHeavy`).
    DashDotDotHeavy,
    /// Wavy underline (`wave`).
    Wavy,
    /// Heavy wavy underline (`wavyHeavy`).
    WavyHeavy,
    /// Double wavy underline (`wavyDouble`).
    WavyDouble,
    /// Underline words only, not spaces (`words`).
    Words,
}

/// A text run inside a paragraph.
#[derive(Debug, Clone, PartialEq, Eq, Default)]
pub struct Run {
    text: String,
    style_id: Option<String>,
    bold: bool,
    italic: bool,
    /// Underline style for this run. `None` means no underline.
    underline: Option<UnderlineType>,
    strikethrough: bool,
    double_strikethrough: bool,
    subscript: bool,
    superscript: bool,
    /// Whether this run uses small capitals (`w:smallCaps`).
    small_caps: bool,
    /// Whether this run uses all capitals (`w:caps`).
    all_caps: bool,
    /// Whether this run is hidden text (`w:vanish`).
    hidden: bool,
    /// Whether this run has emboss effect (`w:emboss`).
    emboss: bool,
    /// Whether this run has imprint/engrave effect (`w:imprint`).
    imprint: bool,
    /// Whether this run has shadow effect (`w:shadow`).
    shadow: bool,
    /// Whether this run has outline effect (`w:outline`).
    outline: bool,
    /// Character spacing adjustment in twips (`w:spacing w:val`).
    character_spacing_twips: Option<i32>,
    highlight_color: Option<String>,
    font_family: Option<String>,
    font_family_ascii: Option<String>,
    font_family_h_ansi: Option<String>,
    font_family_cs: Option<String>,
    font_family_east_asia: Option<String>,
    font_size_half_points: Option<u16>,
    color: Option<String>,
    /// Theme color reference (e.g. "accent1", "dark1", "light1").
    theme_color: Option<String>,
    /// Theme shade applied to the theme color (0-255, hex like "BF").
    theme_shade: Option<String>,
    /// Theme tint applied to the theme color (0-255, hex like "80").
    theme_tint: Option<String>,
    hyperlink: Option<String>,
    /// Tooltip text for a hyperlink (`w:hyperlink w:tooltip`).
    hyperlink_tooltip: Option<String>,
    /// Internal bookmark anchor for a hyperlink (`w:hyperlink w:anchor`).
    hyperlink_anchor: Option<String>,
    inline_image: Option<InlineImage>,
    floating_image: Option<FloatingImage>,
    /// Footnote reference id when this run is a footnote reference marker.
    footnote_reference_id: Option<u32>,
    /// Endnote reference id when this run is an endnote reference marker.
    endnote_reference_id: Option<u32>,
    /// Simple field instruction text (`w:fldSimple w:instr`).
    field_simple: Option<String>,
    /// Complex field code parsed from `w:fldChar` sequences.
    field_code: Option<FieldCode>,
    /// Whether this run contains a tab character (`w:tab`).
    has_tab: bool,
    /// Whether this run contains a line break (`w:br`).
    has_break: bool,
    /// Whether this run has right-to-left text direction (`w:rtl`).
    rtl: bool,
    unknown_children: Vec<RawXmlNode>,
    unknown_property_children: Vec<RawXmlNode>,
}

/// A complex field code parsed from `w:fldChar` begin/separate/end sequences.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct FieldCode {
    instruction: String,
    result: String,
}

impl FieldCode {
    /// Create a new field code with instruction and result.
    pub fn new(instruction: impl Into<String>, result: impl Into<String>) -> Self {
        Self {
            instruction: instruction.into(),
            result: result.into(),
        }
    }

    /// The field instruction text (e.g., `"PAGE"`, `"DATE \\@ \"M/d/yyyy\""`).
    pub fn instruction(&self) -> &str {
        &self.instruction
    }

    /// Set the field instruction text.
    pub fn set_instruction(&mut self, instruction: impl Into<String>) {
        self.instruction = instruction.into();
    }

    /// The field result text (the displayed value).
    pub fn result(&self) -> &str {
        &self.result
    }

    /// Set the field result text.
    pub fn set_result(&mut self, result: impl Into<String>) {
        self.result = result.into();
    }
}

impl Run {
    /// Create a new text run.
    pub fn new(text: impl Into<String>) -> Self {
        Self {
            text: text.into(),
            style_id: None,
            bold: false,
            italic: false,
            underline: None,
            strikethrough: false,
            double_strikethrough: false,
            subscript: false,
            superscript: false,
            small_caps: false,
            all_caps: false,
            hidden: false,
            emboss: false,
            imprint: false,
            shadow: false,
            outline: false,
            character_spacing_twips: None,
            highlight_color: None,
            font_family: None,
            font_family_ascii: None,
            font_family_h_ansi: None,
            font_family_cs: None,
            font_family_east_asia: None,
            font_size_half_points: None,
            color: None,
            theme_color: None,
            theme_shade: None,
            theme_tint: None,
            hyperlink: None,
            hyperlink_tooltip: None,
            hyperlink_anchor: None,
            inline_image: None,
            floating_image: None,
            footnote_reference_id: None,
            endnote_reference_id: None,
            field_simple: None,
            field_code: None,
            has_tab: false,
            has_break: false,
            rtl: false,
            unknown_children: Vec::new(),
            unknown_property_children: Vec::new(),
        }
    }

    /// Get this run's plain text.
    pub fn text(&self) -> &str {
        &self.text
    }

    /// Replace this run's plain text.
    pub fn set_text(&mut self, text: impl Into<String>) {
        self.text = text.into();
    }

    /// Whether this run is bold.
    pub fn style_id(&self) -> Option<&str> {
        self.style_id.as_deref()
    }

    /// Set run style identifier (`w:rStyle`).
    pub fn set_style_id(&mut self, style_id: impl Into<String>) {
        let style_id = style_id.into();
        let trimmed = style_id.trim();
        self.style_id = if trimmed.is_empty() {
            None
        } else {
            Some(trimmed.to_string())
        };
    }

    /// Clear run style identifier.
    pub fn clear_style_id(&mut self) {
        self.style_id = None;
    }

    pub(crate) fn set_style_id_option(&mut self, style_id: Option<String>) {
        self.style_id = style_id.and_then(|value| {
            let trimmed = value.trim();
            if trimmed.is_empty() {
                None
            } else {
                Some(trimmed.to_string())
            }
        });
    }

    /// Whether this run is bold.
    pub fn is_bold(&self) -> bool {
        self.bold
    }

    /// Set bold formatting for this run.
    pub fn set_bold(&mut self, bold: bool) {
        self.bold = bold;
    }

    /// Whether this run is italic.
    pub fn is_italic(&self) -> bool {
        self.italic
    }

    /// Set italic formatting for this run.
    pub fn set_italic(&mut self, italic: bool) {
        self.italic = italic;
    }

    /// Whether this run is underlined (any underline type).
    pub fn is_underline(&self) -> bool {
        self.underline.is_some()
    }

    /// Set underline formatting for this run.
    ///
    /// When `underline` is `true`, sets [`UnderlineType::Single`]; when `false`, removes
    /// the underline. For other underline styles use [`set_underline_type`](Self::set_underline_type).
    pub fn set_underline(&mut self, underline: bool) {
        self.underline = if underline {
            Some(UnderlineType::Single)
        } else {
            None
        };
    }

    /// The specific underline type, or `None` if the run is not underlined.
    pub fn underline_type(&self) -> Option<UnderlineType> {
        self.underline
    }

    /// Set a specific underline type for this run.
    pub fn set_underline_type(&mut self, ut: UnderlineType) {
        self.underline = Some(ut);
    }

    /// Remove underline formatting from this run.
    pub fn clear_underline(&mut self) {
        self.underline = None;
    }

    /// Whether this run has single strikethrough.
    pub fn is_strikethrough(&self) -> bool {
        self.strikethrough
    }

    /// Set single strikethrough formatting for this run.
    pub fn set_strikethrough(&mut self, strikethrough: bool) {
        self.strikethrough = strikethrough;
    }

    /// Whether this run has double strikethrough.
    pub fn is_double_strikethrough(&self) -> bool {
        self.double_strikethrough
    }

    /// Set double strikethrough formatting for this run.
    pub fn set_double_strikethrough(&mut self, double_strikethrough: bool) {
        self.double_strikethrough = double_strikethrough;
    }

    /// Whether this run is subscript.
    pub fn is_subscript(&self) -> bool {
        self.subscript
    }

    /// Set subscript formatting for this run.
    pub fn set_subscript(&mut self, subscript: bool) {
        self.subscript = subscript;
        if subscript {
            self.superscript = false;
        }
    }

    /// Whether this run is superscript.
    pub fn is_superscript(&self) -> bool {
        self.superscript
    }

    /// Set superscript formatting for this run.
    pub fn set_superscript(&mut self, superscript: bool) {
        self.superscript = superscript;
        if superscript {
            self.subscript = false;
        }
    }

    /// Whether this run uses small capitals (`w:smallCaps`).
    pub fn is_small_caps(&self) -> bool {
        self.small_caps
    }

    /// Set small capitals formatting for this run.
    pub fn set_small_caps(&mut self, small_caps: bool) {
        self.small_caps = small_caps;
    }

    /// Whether this run uses all capitals (`w:caps`).
    pub fn is_all_caps(&self) -> bool {
        self.all_caps
    }

    /// Set all capitals formatting for this run.
    pub fn set_all_caps(&mut self, all_caps: bool) {
        self.all_caps = all_caps;
    }

    /// Whether this run is hidden text (`w:vanish`).
    pub fn is_hidden(&self) -> bool {
        self.hidden
    }

    /// Set hidden text formatting for this run.
    pub fn set_hidden(&mut self, hidden: bool) {
        self.hidden = hidden;
    }

    /// Whether this run has emboss effect (`w:emboss`).
    pub fn is_emboss(&self) -> bool {
        self.emboss
    }

    /// Set emboss effect for this run.
    pub fn set_emboss(&mut self, emboss: bool) {
        self.emboss = emboss;
    }

    /// Whether this run has imprint/engrave effect (`w:imprint`).
    pub fn is_imprint(&self) -> bool {
        self.imprint
    }

    /// Set imprint/engrave effect for this run.
    pub fn set_imprint(&mut self, imprint: bool) {
        self.imprint = imprint;
    }

    /// Whether this run has shadow effect (`w:shadow`).
    pub fn is_shadow(&self) -> bool {
        self.shadow
    }

    /// Set shadow effect for this run.
    pub fn set_shadow(&mut self, shadow: bool) {
        self.shadow = shadow;
    }

    /// Whether this run has outline effect (`w:outline`).
    pub fn is_outline(&self) -> bool {
        self.outline
    }

    /// Set outline effect for this run.
    pub fn set_outline(&mut self, outline: bool) {
        self.outline = outline;
    }

    /// Character spacing adjustment in twips (`w:spacing w:val`).
    pub fn character_spacing_twips(&self) -> Option<i32> {
        self.character_spacing_twips
    }

    /// Set character spacing adjustment in twips.
    pub fn set_character_spacing_twips(&mut self, twips: i32) {
        self.character_spacing_twips = Some(twips);
    }

    /// Remove explicit character spacing adjustment.
    pub fn clear_character_spacing_twips(&mut self) {
        self.character_spacing_twips = None;
    }

    /// Highlight color for this run (e.g., `"yellow"`, `"green"`).
    pub fn highlight_color(&self) -> Option<&str> {
        self.highlight_color.as_deref()
    }

    /// Set highlight color for this run (e.g., `"yellow"`, `"green"`).
    pub fn set_highlight_color(&mut self, color: impl Into<String>) {
        let color = color.into();
        self.highlight_color = if color.is_empty() { None } else { Some(color) };
    }

    /// Remove highlight color from this run.
    pub fn clear_highlight_color(&mut self) {
        self.highlight_color = None;
    }

    /// Font family for this run.
    pub fn font_family(&self) -> Option<&str> {
        self.font_family.as_deref()
    }

    /// Set the font family for this run.
    pub fn set_font_family(&mut self, font_family: impl Into<String>) {
        let font_family = font_family.into();
        self.font_family = if font_family.is_empty() {
            None
        } else {
            Some(font_family)
        };
    }

    /// Remove font family formatting.
    pub fn clear_font_family(&mut self) {
        self.font_family = None;
    }

    /// ASCII font family (`w:rFonts w:ascii`).
    pub fn font_family_ascii(&self) -> Option<&str> {
        self.font_family_ascii.as_deref()
    }

    /// Set ASCII font family.
    pub fn set_font_family_ascii(&mut self, font: impl Into<String>) {
        self.font_family_ascii = Some(font.into());
    }

    /// Clear ASCII font family.
    pub fn clear_font_family_ascii(&mut self) {
        self.font_family_ascii = None;
    }

    /// High ANSI font family (`w:rFonts w:hAnsi`).
    pub fn font_family_h_ansi(&self) -> Option<&str> {
        self.font_family_h_ansi.as_deref()
    }

    /// Set High ANSI font family.
    pub fn set_font_family_h_ansi(&mut self, font: impl Into<String>) {
        self.font_family_h_ansi = Some(font.into());
    }

    /// Clear High ANSI font family.
    pub fn clear_font_family_h_ansi(&mut self) {
        self.font_family_h_ansi = None;
    }

    /// Complex script font family (`w:rFonts w:cs`).
    pub fn font_family_cs(&self) -> Option<&str> {
        self.font_family_cs.as_deref()
    }

    /// Set complex script font family.
    pub fn set_font_family_cs(&mut self, font: impl Into<String>) {
        self.font_family_cs = Some(font.into());
    }

    /// Clear complex script font family.
    pub fn clear_font_family_cs(&mut self) {
        self.font_family_cs = None;
    }

    /// East Asian font family (`w:rFonts w:eastAsia`).
    pub fn font_family_east_asia(&self) -> Option<&str> {
        self.font_family_east_asia.as_deref()
    }

    /// Set East Asian font family.
    pub fn set_font_family_east_asia(&mut self, font: impl Into<String>) {
        self.font_family_east_asia = Some(font.into());
    }

    /// Clear East Asian font family.
    pub fn clear_font_family_east_asia(&mut self) {
        self.font_family_east_asia = None;
    }

    /// Font size in half-points (e.g., 24 = 12pt).
    pub fn font_size_half_points(&self) -> Option<u16> {
        self.font_size_half_points
    }

    /// Set font size in half-points.
    pub fn set_font_size_half_points(&mut self, size: u16) {
        self.font_size_half_points = Some(size);
    }

    /// Remove explicit font size.
    pub fn clear_font_size_half_points(&mut self) {
        self.font_size_half_points = None;
    }

    /// Text color value (normalized, typically RGB hex without `#`).
    pub fn color(&self) -> Option<&str> {
        self.color.as_deref()
    }

    /// Set text color (e.g., `FF0000` or `#ff0000`).
    pub fn set_color(&mut self, color: impl Into<String>) {
        let color = color.into();
        self.color = normalize_color_value(color.as_str());
    }

    /// Remove explicit text color.
    pub fn clear_color(&mut self) {
        self.color = None;
    }

    /// Theme color reference (e.g. `"accent1"`, `"dark1"`, `"light1"`).
    pub fn theme_color(&self) -> Option<&str> {
        self.theme_color.as_deref()
    }

    /// Set the theme color reference.
    ///
    /// Common values: `"dark1"`, `"dark2"`, `"light1"`, `"light2"`,
    /// `"accent1"` through `"accent6"`, `"hyperlink"`, `"followedHyperlink"`.
    pub fn set_theme_color(&mut self, theme_color: impl Into<String>) {
        self.theme_color = Some(theme_color.into());
    }

    /// Clear the theme color reference.
    pub fn clear_theme_color(&mut self) {
        self.theme_color = None;
        self.theme_shade = None;
        self.theme_tint = None;
    }

    /// Theme shade value (0-255, stored as hex like `"BF"`).
    pub fn theme_shade(&self) -> Option<&str> {
        self.theme_shade.as_deref()
    }

    /// Set the theme shade value.
    pub fn set_theme_shade(&mut self, shade: impl Into<String>) {
        self.theme_shade = Some(shade.into());
    }

    /// Theme tint value (0-255, stored as hex like `"80"`).
    pub fn theme_tint(&self) -> Option<&str> {
        self.theme_tint.as_deref()
    }

    /// Set the theme tint value.
    pub fn set_theme_tint(&mut self, tint: impl Into<String>) {
        self.theme_tint = Some(tint.into());
    }

    /// Hyperlink URI associated with this run.
    pub fn hyperlink(&self) -> Option<&str> {
        self.hyperlink.as_deref()
    }

    /// Set hyperlink URI for this run.
    pub fn set_hyperlink(&mut self, hyperlink: impl Into<String>) {
        let hyperlink = hyperlink.into();
        self.hyperlink = if hyperlink.is_empty() {
            None
        } else {
            Some(hyperlink)
        };
    }

    /// Remove hyperlink association from this run.
    pub fn clear_hyperlink(&mut self) {
        self.hyperlink = None;
    }

    /// Tooltip text for a hyperlink (`w:hyperlink w:tooltip`).
    pub fn hyperlink_tooltip(&self) -> Option<&str> {
        self.hyperlink_tooltip.as_deref()
    }

    /// Set tooltip text for this run's hyperlink.
    pub fn set_hyperlink_tooltip(&mut self, tooltip: impl Into<String>) {
        let tooltip = tooltip.into();
        self.hyperlink_tooltip = if tooltip.is_empty() {
            None
        } else {
            Some(tooltip)
        };
    }

    /// Remove hyperlink tooltip from this run.
    pub fn clear_hyperlink_tooltip(&mut self) {
        self.hyperlink_tooltip = None;
    }

    /// Internal bookmark anchor for a hyperlink (`w:hyperlink w:anchor`).
    ///
    /// This is used for internal document links that jump to a bookmark
    /// within the same document, as opposed to external URL hyperlinks.
    pub fn hyperlink_anchor(&self) -> Option<&str> {
        self.hyperlink_anchor.as_deref()
    }

    /// Set internal bookmark anchor for this run's hyperlink.
    pub fn set_hyperlink_anchor(&mut self, anchor: impl Into<String>) {
        let anchor = anchor.into();
        self.hyperlink_anchor = if anchor.is_empty() {
            None
        } else {
            Some(anchor)
        };
    }

    /// Remove hyperlink anchor from this run.
    pub fn clear_hyperlink_anchor(&mut self) {
        self.hyperlink_anchor = None;
    }

    /// Inline image payload associated with this run.
    pub fn inline_image(&self) -> Option<&InlineImage> {
        self.inline_image.as_ref()
    }

    /// Set inline image payload for this run.
    pub fn set_inline_image(&mut self, inline_image: InlineImage) {
        self.inline_image = Some(inline_image);
        self.floating_image = None;
    }

    /// Remove inline image payload from this run.
    pub fn clear_inline_image(&mut self) {
        self.inline_image = None;
    }

    /// Floating image payload associated with this run.
    pub fn floating_image(&self) -> Option<&FloatingImage> {
        self.floating_image.as_ref()
    }

    /// Set floating image payload for this run.
    pub fn set_floating_image(&mut self, floating_image: FloatingImage) {
        self.floating_image = Some(floating_image);
        self.inline_image = None;
    }

    /// Remove floating image payload from this run.
    pub fn clear_floating_image(&mut self) {
        self.floating_image = None;
    }

    /// Footnote reference id (`w:footnoteReference w:id`).
    pub fn footnote_reference_id(&self) -> Option<u32> {
        self.footnote_reference_id
    }

    /// Set footnote reference id.
    pub fn set_footnote_reference_id(&mut self, id: u32) {
        self.footnote_reference_id = Some(id);
    }

    /// Clear footnote reference.
    pub fn clear_footnote_reference_id(&mut self) {
        self.footnote_reference_id = None;
    }

    /// Endnote reference id (`w:endnoteReference w:id`).
    pub fn endnote_reference_id(&self) -> Option<u32> {
        self.endnote_reference_id
    }

    /// Set endnote reference id.
    pub fn set_endnote_reference_id(&mut self, id: u32) {
        self.endnote_reference_id = Some(id);
    }

    /// Clear endnote reference.
    pub fn clear_endnote_reference_id(&mut self) {
        self.endnote_reference_id = None;
    }

    /// Simple field instruction text (`w:fldSimple w:instr`).
    pub fn field_simple(&self) -> Option<&str> {
        self.field_simple.as_deref()
    }

    /// Set simple field instruction (e.g., `"PAGE"`, `"DATE"`).
    pub fn set_field_simple(&mut self, instruction: impl Into<String>) {
        let instruction = instruction.into();
        self.field_simple = if instruction.trim().is_empty() {
            None
        } else {
            Some(instruction)
        };
    }

    /// Clear simple field instruction.
    pub fn clear_field_simple(&mut self) {
        self.field_simple = None;
    }

    /// Complex field code parsed from `w:fldChar` sequences.
    pub fn field_code(&self) -> Option<&FieldCode> {
        self.field_code.as_ref()
    }

    /// Set a complex field code on this run.
    pub fn set_field_code(&mut self, field_code: FieldCode) {
        self.field_code = Some(field_code);
    }

    /// Clear complex field code.
    pub fn clear_field_code(&mut self) {
        self.field_code = None;
    }

    /// Whether this run contains a tab character (`w:tab`).
    pub fn has_tab(&self) -> bool {
        self.has_tab
    }

    /// Set whether this run contains a tab character.
    pub fn set_has_tab(&mut self, has_tab: bool) {
        self.has_tab = has_tab;
    }

    /// Whether this run contains a line break (`w:br`).
    pub fn has_break(&self) -> bool {
        self.has_break
    }

    /// Set whether this run contains a line break.
    pub fn set_has_break(&mut self, has_break: bool) {
        self.has_break = has_break;
    }

    /// Whether this run has right-to-left text direction (`w:rtl`).
    pub fn is_rtl(&self) -> bool {
        self.rtl
    }

    /// Set right-to-left text direction for this run.
    pub fn set_rtl(&mut self, rtl: bool) {
        self.rtl = rtl;
    }

    pub(crate) fn has_properties(&self) -> bool {
        self.bold
            || self.italic
            || self.underline.is_some()
            || self.strikethrough
            || self.double_strikethrough
            || self.subscript
            || self.superscript
            || self.small_caps
            || self.all_caps
            || self.hidden
            || self.emboss
            || self.imprint
            || self.shadow
            || self.outline
            || self.character_spacing_twips.is_some()
            || self.highlight_color.is_some()
            || self.style_id.is_some()
            || self.font_family.is_some()
            || self.font_family_ascii.is_some()
            || self.font_family_h_ansi.is_some()
            || self.font_family_cs.is_some()
            || self.font_family_east_asia.is_some()
            || self.font_size_half_points.is_some()
            || self.color.is_some()
            || self.rtl
            || !self.unknown_property_children.is_empty()
    }

    pub(crate) fn unknown_children(&self) -> &[RawXmlNode] {
        self.unknown_children.as_slice()
    }

    pub(crate) fn push_unknown_child(&mut self, node: RawXmlNode) {
        self.unknown_children.push(node);
    }

    pub(crate) fn unknown_property_children(&self) -> &[RawXmlNode] {
        self.unknown_property_children.as_slice()
    }

    pub(crate) fn push_unknown_property_child(&mut self, node: RawXmlNode) {
        self.unknown_property_children.push(node);
    }
}

fn normalize_color_value(value: &str) -> Option<String> {
    let trimmed = value.trim();
    let normalized = trimmed.trim_start_matches('#').to_ascii_uppercase();

    if normalized.is_empty() {
        None
    } else {
        Some(normalized)
    }
}

#[cfg(test)]
mod tests {
    use super::{FieldCode, Run};
    use crate::image::{FloatingImage, InlineImage};

    #[test]
    fn formatting_flags_default_to_false() {
        let run = Run::new("hello");

        assert!(!run.is_bold());
        assert!(!run.is_italic());
        assert!(!run.is_underline());
        assert_eq!(run.underline_type(), None);
        assert!(!run.is_strikethrough());
        assert!(!run.is_double_strikethrough());
        assert!(!run.is_subscript());
        assert!(!run.is_superscript());
        assert!(!run.is_small_caps());
        assert!(!run.is_all_caps());
        assert!(!run.is_hidden());
        assert_eq!(run.character_spacing_twips(), None);
        assert_eq!(run.highlight_color(), None);
        assert_eq!(run.style_id(), None);
        assert_eq!(run.font_family(), None);
        assert_eq!(run.font_size_half_points(), None);
        assert_eq!(run.color(), None);
        assert_eq!(run.hyperlink(), None);
        assert_eq!(run.hyperlink_tooltip(), None);
        assert_eq!(run.hyperlink_anchor(), None);
        assert_eq!(run.inline_image(), None);
        assert_eq!(run.floating_image(), None);
    }

    #[test]
    fn formatting_flags_can_be_set() {
        let mut run = Run::new("hello");
        run.set_style_id("Emphasis");
        run.set_bold(true);
        run.set_italic(true);
        run.set_underline(true);
        run.set_font_family("Calibri");
        run.set_font_size_half_points(28);
        run.set_color("#3a5fcd");
        run.set_hyperlink("https://example.com");
        run.set_floating_image(FloatingImage::new(1, 990_000, 792_000));

        assert!(run.is_bold());
        assert!(run.is_italic());
        assert!(run.is_underline());
        assert_eq!(run.style_id(), Some("Emphasis"));
        assert_eq!(run.font_family(), Some("Calibri"));
        assert_eq!(run.font_size_half_points(), Some(28));
        assert_eq!(run.color(), Some("3A5FCD"));
        assert_eq!(run.hyperlink(), Some("https://example.com"));
        assert_eq!(
            run.floating_image().map(FloatingImage::image_index),
            Some(1)
        );
    }

    #[test]
    fn optional_character_formatting_can_be_cleared() {
        let mut run = Run::new("hello");
        run.set_style_id("Strong");
        run.set_font_family("Cambria");
        run.set_font_size_half_points(24);
        run.set_color("FFAA00");
        run.set_hyperlink("https://example.com");
        run.set_inline_image(InlineImage::new(2, 720_000, 540_000));
        run.set_floating_image(FloatingImage::new(3, 810_000, 640_000));

        run.clear_style_id();
        run.clear_font_family();
        run.clear_font_size_half_points();
        run.clear_color();
        run.clear_hyperlink();
        run.clear_inline_image();
        run.clear_floating_image();

        assert_eq!(run.style_id(), None);
        assert_eq!(run.font_family(), None);
        assert_eq!(run.font_size_half_points(), None);
        assert_eq!(run.color(), None);
        assert_eq!(run.hyperlink(), None);
        assert_eq!(run.inline_image(), None);
        assert_eq!(run.floating_image(), None);
    }

    #[test]
    fn setting_one_image_mode_clears_the_other() {
        let mut run = Run::new("");
        run.set_inline_image(InlineImage::new(2, 720_000, 540_000));
        assert_eq!(run.inline_image().map(InlineImage::image_index), Some(2));
        assert_eq!(run.floating_image(), None);

        run.set_floating_image(FloatingImage::new(3, 810_000, 640_000));
        assert_eq!(run.inline_image(), None);
        assert_eq!(
            run.floating_image().map(FloatingImage::image_index),
            Some(3)
        );

        run.set_inline_image(InlineImage::new(4, 640_000, 480_000));
        assert_eq!(run.inline_image().map(InlineImage::image_index), Some(4));
        assert_eq!(run.floating_image(), None);
    }

    #[test]
    fn strikethrough_can_be_set_and_read() {
        let mut run = Run::new("struck");
        assert!(!run.is_strikethrough());
        run.set_strikethrough(true);
        assert!(run.is_strikethrough());
        run.set_strikethrough(false);
        assert!(!run.is_strikethrough());
    }

    #[test]
    fn double_strikethrough_can_be_set_and_read() {
        let mut run = Run::new("double struck");
        assert!(!run.is_double_strikethrough());
        run.set_double_strikethrough(true);
        assert!(run.is_double_strikethrough());
        run.set_double_strikethrough(false);
        assert!(!run.is_double_strikethrough());
    }

    #[test]
    fn subscript_can_be_set_and_read() {
        let mut run = Run::new("sub");
        assert!(!run.is_subscript());
        run.set_subscript(true);
        assert!(run.is_subscript());
        assert!(!run.is_superscript());
    }

    #[test]
    fn superscript_can_be_set_and_read() {
        let mut run = Run::new("sup");
        assert!(!run.is_superscript());
        run.set_superscript(true);
        assert!(run.is_superscript());
        assert!(!run.is_subscript());
    }

    #[test]
    fn setting_subscript_clears_superscript_and_vice_versa() {
        let mut run = Run::new("toggle");
        run.set_superscript(true);
        assert!(run.is_superscript());
        run.set_subscript(true);
        assert!(run.is_subscript());
        assert!(!run.is_superscript());
        run.set_superscript(true);
        assert!(run.is_superscript());
        assert!(!run.is_subscript());
    }

    #[test]
    fn highlight_color_can_be_set_and_cleared() {
        let mut run = Run::new("highlight");
        assert_eq!(run.highlight_color(), None);
        run.set_highlight_color("yellow");
        assert_eq!(run.highlight_color(), Some("yellow"));
        run.set_highlight_color("green");
        assert_eq!(run.highlight_color(), Some("green"));
        run.clear_highlight_color();
        assert_eq!(run.highlight_color(), None);
    }

    #[test]
    fn new_properties_affect_has_properties() {
        let mut run = Run::new("test");
        assert!(!run.has_properties());

        run.set_strikethrough(true);
        assert!(run.has_properties());
        run.set_strikethrough(false);

        run.set_double_strikethrough(true);
        assert!(run.has_properties());
        run.set_double_strikethrough(false);

        run.set_subscript(true);
        assert!(run.has_properties());
        run.set_subscript(false);

        run.set_superscript(true);
        assert!(run.has_properties());
        run.set_superscript(false);

        run.set_highlight_color("yellow");
        assert!(run.has_properties());
        run.clear_highlight_color();

        assert!(!run.has_properties());
    }

    #[test]
    fn field_code_can_be_set_and_cleared() {
        let mut run = Run::new("");
        assert_eq!(run.field_code(), None);

        let field = FieldCode::new("PAGE", "3");
        run.set_field_code(field);
        assert_eq!(run.field_code().map(FieldCode::instruction), Some("PAGE"));
        assert_eq!(run.field_code().map(FieldCode::result), Some("3"));

        run.clear_field_code();
        assert_eq!(run.field_code(), None);
    }

    #[test]
    fn field_code_instruction_and_result_can_be_modified() {
        let mut field = FieldCode::new("DATE", "2024-01-01");
        assert_eq!(field.instruction(), "DATE");
        assert_eq!(field.result(), "2024-01-01");

        field.set_instruction("TIME");
        field.set_result("12:00");
        assert_eq!(field.instruction(), "TIME");
        assert_eq!(field.result(), "12:00");
    }

    #[test]
    fn rtl_can_be_set_and_read() {
        let mut run = Run::new("text");
        assert!(!run.is_rtl());

        run.set_rtl(true);
        assert!(run.is_rtl());
        assert!(run.has_properties());

        run.set_rtl(false);
        assert!(!run.is_rtl());
        assert!(!run.has_properties());
    }

    #[test]
    fn rtl_affects_has_properties() {
        let mut run = Run::new("test");
        assert!(!run.has_properties());

        run.set_rtl(true);
        assert!(run.has_properties());

        run.set_rtl(false);
        assert!(!run.has_properties());
    }

    // ── Theme color tests ──

    #[test]
    fn theme_color_defaults_none() {
        let run = Run::new("test");
        assert!(run.theme_color().is_none());
        assert!(run.theme_shade().is_none());
        assert!(run.theme_tint().is_none());
    }

    #[test]
    fn set_theme_color_and_modifiers() {
        let mut run = Run::new("test");
        run.set_theme_color("accent1");
        run.set_theme_shade("BF");
        run.set_theme_tint("80");

        assert_eq!(run.theme_color(), Some("accent1"));
        assert_eq!(run.theme_shade(), Some("BF"));
        assert_eq!(run.theme_tint(), Some("80"));
    }

    #[test]
    fn clear_theme_color_clears_all() {
        let mut run = Run::new("test");
        run.set_theme_color("dark1");
        run.set_theme_shade("BF");
        run.set_theme_tint("80");

        run.clear_theme_color();
        assert!(run.theme_color().is_none());
        assert!(run.theme_shade().is_none());
        assert!(run.theme_tint().is_none());
    }
}
