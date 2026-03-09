/// Represents image cropping information.
///
/// Crop values are stored as percentages (0.0 to 100.0) representing
/// how much to crop from each edge. For example, left=10.0 means crop
/// 10% from the left edge.
#[derive(Debug, Clone, PartialEq)]
pub struct ImageCrop {
    /// Percentage to crop from the left edge (0.0 to 100.0)
    pub left: f64,
    /// Percentage to crop from the top edge (0.0 to 100.0)
    pub top: f64,
    /// Percentage to crop from the right edge (0.0 to 100.0)
    pub right: f64,
    /// Percentage to crop from the bottom edge (0.0 to 100.0)
    pub bottom: f64,
}

impl ImageCrop {
    /// Creates a new ImageCrop with all edges set to 0 (no cropping).
    pub fn none() -> Self {
        Self {
            left: 0.0,
            top: 0.0,
            right: 0.0,
            bottom: 0.0,
        }
    }

    /// Creates a new ImageCrop with specified values for each edge.
    pub fn new(left: f64, top: f64, right: f64, bottom: f64) -> Self {
        Self {
            left,
            top,
            right,
            bottom,
        }
    }

    /// Returns true if no cropping is applied (all values are 0).
    pub fn is_none(&self) -> bool {
        self.left == 0.0 && self.top == 0.0 && self.right == 0.0 && self.bottom == 0.0
    }

    /// Converts crop percentage to PowerPoint's internal format (0-100000).
    ///
    /// PowerPoint uses an integer scale where 100000 = 100%.
    pub fn to_pptx_format(&self) -> (i32, i32, i32, i32) {
        (
            (self.left * 1000.0) as i32,
            (self.top * 1000.0) as i32,
            (self.right * 1000.0) as i32,
            (self.bottom * 1000.0) as i32,
        )
    }

    /// Creates ImageCrop from PowerPoint's internal format (0-100000).
    pub fn from_pptx_format(left: i32, top: i32, right: i32, bottom: i32) -> Self {
        Self {
            left: left as f64 / 1000.0,
            top: top as f64 / 1000.0,
            right: right as f64 / 1000.0,
            bottom: bottom as f64 / 1000.0,
        }
    }
}

#[derive(Debug, Clone, PartialEq)]
pub struct Image {
    bytes: Vec<u8>,
    content_type: String,
    name: Option<String>,
    relationship_id: Option<String>,
    crop: Option<ImageCrop>,
    /// Transparency level (0.0 = fully opaque, 1.0 = fully transparent).
    transparency: Option<f64>,
}

impl Image {
    pub fn new(bytes: impl Into<Vec<u8>>, content_type: impl Into<String>) -> Self {
        Self {
            bytes: bytes.into(),
            content_type: content_type.into(),
            name: None,
            relationship_id: None,
            crop: None,
            transparency: None,
        }
    }

    pub fn bytes(&self) -> &[u8] {
        &self.bytes
    }

    pub fn content_type(&self) -> &str {
        &self.content_type
    }

    pub fn name(&self) -> Option<&str> {
        self.name.as_deref()
    }

    pub fn relationship_id(&self) -> Option<&str> {
        self.relationship_id.as_deref()
    }

    /// Returns the crop settings for this image, if any.
    pub fn crop(&self) -> Option<&ImageCrop> {
        self.crop.as_ref()
    }

    /// Sets the crop settings for this image.
    ///
    /// # Arguments
    /// * `crop` - The crop settings to apply, or None to remove cropping
    ///
    /// # Example
    /// ```no_run
    /// # use offidized_pptx::{Image, ImageCrop};
    /// let mut image = Image::new(vec![1, 2, 3], "image/png");
    /// // Crop 10% from left, 20% from top, 5% from right, 15% from bottom
    /// image.set_crop(Some(ImageCrop::new(10.0, 20.0, 5.0, 15.0)));
    /// ```
    pub fn set_crop(&mut self, crop: Option<ImageCrop>) {
        self.crop = crop;
    }

    /// Convenience method to crop the image from all sides equally.
    ///
    /// # Arguments
    /// * `percent` - Percentage to crop from all sides (0.0 to 100.0)
    pub fn crop_uniform(&mut self, percent: f64) {
        self.crop = Some(ImageCrop::new(percent, percent, percent, percent));
    }

    /// Returns the transparency level (0.0 = opaque, 1.0 = fully transparent), if set.
    pub fn transparency(&self) -> Option<f64> {
        self.transparency
    }

    /// Sets the transparency level (0.0 = opaque, 1.0 = fully transparent).
    pub fn set_transparency(&mut self, alpha: f64) {
        self.transparency = Some(alpha.clamp(0.0, 1.0));
    }

    /// Clears the transparency setting.
    pub fn clear_transparency(&mut self) {
        self.transparency = None;
    }

    pub(crate) fn set_name(&mut self, name: Option<String>) {
        self.name = name;
    }

    pub(crate) fn set_relationship_id(&mut self, relationship_id: Option<String>) {
        self.relationship_id = relationship_id;
    }
}

#[cfg(test)]
mod tests {
    use super::{Image, ImageCrop};

    #[test]
    fn stores_binary_data_and_content_type() {
        let image = Image::new(vec![1_u8, 2, 3], "image/png");
        assert_eq!(image.bytes(), [1_u8, 2, 3]);
        assert_eq!(image.content_type(), "image/png");
        assert!(image.crop().is_none());
    }

    #[test]
    fn image_crop_none() {
        let crop = ImageCrop::none();
        assert!(crop.is_none());
        assert_eq!(crop.left, 0.0);
        assert_eq!(crop.top, 0.0);
        assert_eq!(crop.right, 0.0);
        assert_eq!(crop.bottom, 0.0);
    }

    #[test]
    fn image_crop_new() {
        let crop = ImageCrop::new(10.0, 20.0, 5.0, 15.0);
        assert!(!crop.is_none());
        assert_eq!(crop.left, 10.0);
        assert_eq!(crop.top, 20.0);
        assert_eq!(crop.right, 5.0);
        assert_eq!(crop.bottom, 15.0);
    }

    #[test]
    fn image_crop_pptx_format_conversion() {
        let crop = ImageCrop::new(10.0, 20.0, 5.0, 15.0);
        let (l, t, r, b) = crop.to_pptx_format();
        assert_eq!(l, 10000);
        assert_eq!(t, 20000);
        assert_eq!(r, 5000);
        assert_eq!(b, 15000);

        let crop2 = ImageCrop::from_pptx_format(10000, 20000, 5000, 15000);
        assert_eq!(crop2.left, 10.0);
        assert_eq!(crop2.top, 20.0);
        assert_eq!(crop2.right, 5.0);
        assert_eq!(crop2.bottom, 15.0);
    }

    #[test]
    fn set_crop_on_image() {
        let mut image = Image::new(vec![1, 2, 3], "image/png");
        assert!(image.crop().is_none());

        image.set_crop(Some(ImageCrop::new(10.0, 20.0, 5.0, 15.0)));
        assert!(image.crop().is_some());
        let crop = image.crop().unwrap();
        assert_eq!(crop.left, 10.0);
        assert_eq!(crop.top, 20.0);

        image.set_crop(None);
        assert!(image.crop().is_none());
    }

    #[test]
    fn crop_uniform() {
        let mut image = Image::new(vec![1, 2, 3], "image/png");
        image.crop_uniform(15.0);

        let crop = image.crop().unwrap();
        assert_eq!(crop.left, 15.0);
        assert_eq!(crop.top, 15.0);
        assert_eq!(crop.right, 15.0);
        assert_eq!(crop.bottom, 15.0);
    }
}
