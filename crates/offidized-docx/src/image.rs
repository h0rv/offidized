const OCTET_STREAM_CONTENT_TYPE: &str = "application/octet-stream";

/// Minimal image payload scaffold.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Image {
    bytes: Vec<u8>,
    content_type: String,
}

impl Image {
    pub fn new(bytes: impl Into<Vec<u8>>, content_type: impl Into<String>) -> Self {
        Self {
            bytes: bytes.into(),
            content_type: normalize_content_type(content_type.into()),
        }
    }

    pub fn bytes(&self) -> &[u8] {
        &self.bytes
    }

    pub fn content_type(&self) -> &str {
        self.content_type.as_str()
    }
}

/// Minimal inline drawing metadata bound to a document image index.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct InlineImage {
    image_index: usize,
    width_emu: u32,
    height_emu: u32,
    name: Option<String>,
    description: Option<String>,
}

impl InlineImage {
    pub fn new(image_index: usize, width_emu: u32, height_emu: u32) -> Self {
        Self {
            image_index,
            width_emu,
            height_emu,
            name: None,
            description: None,
        }
    }

    pub fn image_index(&self) -> usize {
        self.image_index
    }

    pub fn width_emu(&self) -> u32 {
        self.width_emu
    }

    pub fn height_emu(&self) -> u32 {
        self.height_emu
    }

    pub fn set_width_emu(&mut self, width_emu: u32) {
        self.width_emu = width_emu;
    }

    pub fn set_height_emu(&mut self, height_emu: u32) {
        self.height_emu = height_emu;
    }

    pub fn name(&self) -> Option<&str> {
        self.name.as_deref()
    }

    pub fn set_name(&mut self, name: impl Into<String>) {
        self.name = normalize_optional_text(name.into());
    }

    pub fn clear_name(&mut self) {
        self.name = None;
    }

    pub fn description(&self) -> Option<&str> {
        self.description.as_deref()
    }

    pub fn set_description(&mut self, description: impl Into<String>) {
        self.description = normalize_optional_text(description.into());
    }

    pub fn clear_description(&mut self) {
        self.description = None;
    }
}

/// Text wrapping type for floating images.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum WrapType {
    /// Square wrapping — text wraps around a rectangular boundary.
    Square,
    /// Tight wrapping — text wraps close to the image contour.
    Tight,
    /// Behind text — image is placed behind the text layer.
    BehindText,
    /// In front of text — image is placed above the text layer.
    InFrontOfText,
    /// Top and bottom — text appears only above and below the image.
    TopAndBottom,
}

impl WrapType {
    /// Parse from XML element local name.
    pub fn from_xml(local: &[u8]) -> Option<Self> {
        match local {
            b"wrapSquare" => Some(Self::Square),
            b"wrapTight" => Some(Self::Tight),
            b"wrapNone" => None, // wrapNone alone doesn't indicate behind/front — need behindDoc
            b"wrapTopAndBottom" => Some(Self::TopAndBottom),
            _ => None,
        }
    }
}

/// Minimal floating drawing metadata bound to a document image index.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct FloatingImage {
    image_index: usize,
    width_emu: u32,
    height_emu: u32,
    offset_x_emu: i32,
    offset_y_emu: i32,
    name: Option<String>,
    description: Option<String>,
    /// Text wrapping mode.
    wrap_type: Option<WrapType>,
}

impl FloatingImage {
    pub fn new(image_index: usize, width_emu: u32, height_emu: u32) -> Self {
        Self {
            image_index,
            width_emu,
            height_emu,
            offset_x_emu: 0,
            offset_y_emu: 0,
            name: None,
            description: None,
            wrap_type: None,
        }
    }

    pub fn image_index(&self) -> usize {
        self.image_index
    }

    pub fn width_emu(&self) -> u32 {
        self.width_emu
    }

    pub fn height_emu(&self) -> u32 {
        self.height_emu
    }

    pub fn set_width_emu(&mut self, width_emu: u32) {
        self.width_emu = width_emu;
    }

    pub fn set_height_emu(&mut self, height_emu: u32) {
        self.height_emu = height_emu;
    }

    pub fn offset_x_emu(&self) -> i32 {
        self.offset_x_emu
    }

    pub fn offset_y_emu(&self) -> i32 {
        self.offset_y_emu
    }

    pub fn set_offset_x_emu(&mut self, offset_x_emu: i32) {
        self.offset_x_emu = offset_x_emu;
    }

    pub fn set_offset_y_emu(&mut self, offset_y_emu: i32) {
        self.offset_y_emu = offset_y_emu;
    }

    pub fn set_offsets_emu(&mut self, offset_x_emu: i32, offset_y_emu: i32) {
        self.offset_x_emu = offset_x_emu;
        self.offset_y_emu = offset_y_emu;
    }

    pub fn name(&self) -> Option<&str> {
        self.name.as_deref()
    }

    pub fn set_name(&mut self, name: impl Into<String>) {
        self.name = normalize_optional_text(name.into());
    }

    pub fn clear_name(&mut self) {
        self.name = None;
    }

    pub fn description(&self) -> Option<&str> {
        self.description.as_deref()
    }

    pub fn set_description(&mut self, description: impl Into<String>) {
        self.description = normalize_optional_text(description.into());
    }

    pub fn clear_description(&mut self) {
        self.description = None;
    }

    /// Returns the text wrapping type, if set.
    pub fn wrap_type(&self) -> Option<WrapType> {
        self.wrap_type
    }

    /// Sets the text wrapping type.
    pub fn set_wrap_type(&mut self, wrap_type: WrapType) {
        self.wrap_type = Some(wrap_type);
    }

    /// Clears the text wrapping type.
    pub fn clear_wrap_type(&mut self) {
        self.wrap_type = None;
    }
}

fn normalize_content_type(content_type: String) -> String {
    let normalized = content_type
        .split(';')
        .next()
        .map(str::trim)
        .unwrap_or_default();

    if normalized.is_empty() {
        OCTET_STREAM_CONTENT_TYPE.to_string()
    } else {
        normalized.to_ascii_lowercase()
    }
}

fn normalize_optional_text(value: String) -> Option<String> {
    let trimmed = value.trim();
    if trimmed.is_empty() {
        None
    } else {
        Some(trimmed.to_string())
    }
}

#[cfg(test)]
mod tests {
    use super::{FloatingImage, Image, InlineImage, WrapType};

    #[test]
    fn stores_binary_data_and_content_type() {
        let image = Image::new(vec![1_u8, 2, 3], "image/png; charset=utf-8");

        assert_eq!(image.bytes(), [1_u8, 2, 3]);
        assert_eq!(image.content_type(), "image/png");
    }

    #[test]
    fn inline_image_tracks_metadata() {
        let mut inline_image = InlineImage::new(3, 990_000, 792_000);
        inline_image.set_name("Picture 1");
        inline_image.set_description("example");

        assert_eq!(inline_image.image_index(), 3);
        assert_eq!(inline_image.width_emu(), 990_000);
        assert_eq!(inline_image.height_emu(), 792_000);
        assert_eq!(inline_image.name(), Some("Picture 1"));
        assert_eq!(inline_image.description(), Some("example"));

        inline_image.clear_name();
        inline_image.clear_description();
        assert_eq!(inline_image.name(), None);
        assert_eq!(inline_image.description(), None);
    }

    #[test]
    fn floating_image_tracks_metadata() {
        let mut floating_image = FloatingImage::new(2, 1_080_000, 864_000);
        floating_image.set_offsets_emu(12_345, -6_789);
        floating_image.set_name("Anchor image");
        floating_image.set_description("floating");

        assert_eq!(floating_image.image_index(), 2);
        assert_eq!(floating_image.width_emu(), 1_080_000);
        assert_eq!(floating_image.height_emu(), 864_000);
        assert_eq!(floating_image.offset_x_emu(), 12_345);
        assert_eq!(floating_image.offset_y_emu(), -6_789);
        assert_eq!(floating_image.name(), Some("Anchor image"));
        assert_eq!(floating_image.description(), Some("floating"));

        floating_image.clear_name();
        floating_image.clear_description();
        assert_eq!(floating_image.name(), None);
        assert_eq!(floating_image.description(), None);
    }

    #[test]
    fn wrap_type_defaults_none() {
        let img = FloatingImage::new(0, 100_000, 100_000);
        assert!(img.wrap_type().is_none());
    }

    #[test]
    fn wrap_type_set_and_clear() {
        let mut img = FloatingImage::new(0, 100_000, 100_000);
        img.set_wrap_type(WrapType::Square);
        assert_eq!(img.wrap_type(), Some(WrapType::Square));

        img.set_wrap_type(WrapType::BehindText);
        assert_eq!(img.wrap_type(), Some(WrapType::BehindText));

        img.clear_wrap_type();
        assert!(img.wrap_type().is_none());
    }

    #[test]
    fn wrap_type_all_variants() {
        let mut img = FloatingImage::new(0, 100_000, 100_000);
        for wt in [
            WrapType::Square,
            WrapType::Tight,
            WrapType::BehindText,
            WrapType::InFrontOfText,
            WrapType::TopAndBottom,
        ] {
            img.set_wrap_type(wt);
            assert_eq!(img.wrap_type(), Some(wt));
        }
    }
}
