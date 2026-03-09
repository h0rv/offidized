//! Typed binary content helpers for OPC packages.
//!
//! In the Open XML SDK, `DataPart` and `MediaDataPart` provide typed access
//! to binary parts (images, audio, video, etc.) within a package. This module
//! provides equivalent helpers.

use crate::part::Part;
use crate::uri::PartUri;

/// Known media content type categories.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum MediaCategory {
    Image,
    Audio,
    Video,
    Other,
}

impl MediaCategory {
    /// Classify a content type string into a media category.
    pub fn from_content_type(content_type: &str) -> Self {
        let ct = content_type.to_lowercase();
        if ct.starts_with("image/") {
            Self::Image
        } else if ct.starts_with("audio/") {
            Self::Audio
        } else if ct.starts_with("video/") {
            Self::Video
        } else {
            Self::Other
        }
    }
}

/// A typed wrapper around a binary data part in a package.
///
/// Provides convenience methods for working with binary parts like images,
/// audio, and video files.
#[derive(Debug, Clone)]
pub struct DataPart {
    /// The part URI within the package.
    pub uri: String,
    /// The content type (e.g., "image/png").
    pub content_type: Option<String>,
    /// Raw bytes of the data.
    pub data: Vec<u8>,
}

impl DataPart {
    /// Create a new data part from bytes.
    pub fn new(uri: impl Into<String>, data: Vec<u8>) -> Self {
        Self {
            uri: uri.into(),
            content_type: None,
            data,
        }
    }

    /// Create a new data part with a content type.
    pub fn with_content_type(
        uri: impl Into<String>,
        content_type: impl Into<String>,
        data: Vec<u8>,
    ) -> Self {
        Self {
            uri: uri.into(),
            content_type: Some(content_type.into()),
            data,
        }
    }

    /// Create from an existing Part (if it's binary data).
    pub fn from_part(part: &Part) -> Option<Self> {
        if part.is_binary() {
            Some(Self {
                uri: part.uri.as_str().to_string(),
                content_type: part.content_type.clone(),
                data: part.data.as_bytes().to_vec(),
            })
        } else {
            None
        }
    }

    /// Get the media category of this part.
    pub fn media_category(&self) -> MediaCategory {
        self.content_type
            .as_deref()
            .map(MediaCategory::from_content_type)
            .unwrap_or(MediaCategory::Other)
    }

    /// Get the file extension from the URI.
    pub fn extension(&self) -> Option<&str> {
        self.uri.rsplit('.').next()
    }

    /// Get the size in bytes.
    pub fn size(&self) -> usize {
        self.data.len()
    }

    /// Convert to an OPC Part for adding to a package.
    pub fn to_part(&self) -> crate::error::Result<Part> {
        let uri = PartUri::new(&self.uri)?;
        let mut part = Part::new(uri, self.data.clone());
        part.content_type = self.content_type.clone();
        Ok(part)
    }
}

/// A typed wrapper for media (image/audio/video) data parts.
///
/// Extends `DataPart` with media-specific convenience methods.
#[derive(Debug, Clone)]
pub struct MediaDataPart {
    inner: DataPart,
    category: MediaCategory,
}

impl MediaDataPart {
    /// Create a new media data part.
    pub fn new(uri: impl Into<String>, content_type: impl Into<String>, data: Vec<u8>) -> Self {
        let content_type = content_type.into();
        let category = MediaCategory::from_content_type(&content_type);
        Self {
            inner: DataPart::with_content_type(uri, content_type, data),
            category,
        }
    }

    /// Create from an existing Part, returning None if not a media type.
    pub fn from_part(part: &Part) -> Option<Self> {
        let content_type = part.content_type.as_deref()?;
        let category = MediaCategory::from_content_type(content_type);
        match category {
            MediaCategory::Other => None,
            _ => Some(Self {
                inner: DataPart {
                    uri: part.uri.as_str().to_string(),
                    content_type: part.content_type.clone(),
                    data: part.data.as_bytes().to_vec(),
                },
                category,
            }),
        }
    }

    /// Get the media category.
    pub fn category(&self) -> MediaCategory {
        self.category
    }

    /// Whether this is an image.
    pub fn is_image(&self) -> bool {
        self.category == MediaCategory::Image
    }

    /// Whether this is audio.
    pub fn is_audio(&self) -> bool {
        self.category == MediaCategory::Audio
    }

    /// Whether this is video.
    pub fn is_video(&self) -> bool {
        self.category == MediaCategory::Video
    }

    /// Get the content type.
    pub fn content_type(&self) -> &str {
        self.inner
            .content_type
            .as_deref()
            .unwrap_or("application/octet-stream")
    }

    /// Get the part URI.
    pub fn uri(&self) -> &str {
        &self.inner.uri
    }

    /// Get the raw data bytes.
    pub fn data(&self) -> &[u8] {
        &self.inner.data
    }

    /// Get the size in bytes.
    pub fn size(&self) -> usize {
        self.inner.size()
    }

    /// Get the file extension.
    pub fn extension(&self) -> Option<&str> {
        self.inner.extension()
    }

    /// Convert to an OPC Part.
    pub fn to_part(&self) -> crate::error::Result<Part> {
        self.inner.to_part()
    }

    /// Access the inner DataPart.
    pub fn as_data_part(&self) -> &DataPart {
        &self.inner
    }
}

/// Helper to detect common image content types from file extension.
pub fn content_type_from_extension(extension: &str) -> Option<&'static str> {
    match extension.to_lowercase().as_str() {
        "png" => Some("image/png"),
        "jpg" | "jpeg" => Some("image/jpeg"),
        "gif" => Some("image/gif"),
        "bmp" => Some("image/bmp"),
        "svg" => Some("image/svg+xml"),
        "tif" | "tiff" => Some("image/tiff"),
        "ico" => Some("image/x-icon"),
        "webp" => Some("image/webp"),
        "mp3" => Some("audio/mpeg"),
        "wav" => Some("audio/wav"),
        "ogg" => Some("audio/ogg"),
        "mp4" => Some("video/mp4"),
        "mpeg" => Some("video/mpeg"),
        "wmv" => Some("video/x-ms-wmv"),
        "avi" => Some("video/x-msvideo"),
        "webm" => Some("video/webm"),
        _ => None,
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn media_category_classifies_content_types() {
        assert_eq!(
            MediaCategory::from_content_type("image/png"),
            MediaCategory::Image
        );
        assert_eq!(
            MediaCategory::from_content_type("audio/mpeg"),
            MediaCategory::Audio
        );
        assert_eq!(
            MediaCategory::from_content_type("video/mp4"),
            MediaCategory::Video
        );
        assert_eq!(
            MediaCategory::from_content_type("application/xml"),
            MediaCategory::Other
        );
    }

    #[test]
    fn data_part_basic_operations() {
        let part = DataPart::with_content_type(
            "/media/image1.png",
            "image/png",
            vec![0x89, 0x50, 0x4E, 0x47],
        );

        assert_eq!(part.uri, "/media/image1.png");
        assert_eq!(part.content_type.as_deref(), Some("image/png"));
        assert_eq!(part.size(), 4);
        assert_eq!(part.extension(), Some("png"));
        assert_eq!(part.media_category(), MediaCategory::Image);
    }

    #[test]
    fn media_data_part_type_checks() {
        let image = MediaDataPart::new("/media/img.png", "image/png", vec![1, 2, 3]);
        assert!(image.is_image());
        assert!(!image.is_audio());
        assert!(!image.is_video());
        assert_eq!(image.category(), MediaCategory::Image);

        let audio = MediaDataPart::new("/media/sound.mp3", "audio/mpeg", vec![4, 5, 6]);
        assert!(audio.is_audio());
        assert!(!audio.is_image());

        let video = MediaDataPart::new("/media/clip.mp4", "video/mp4", vec![7, 8, 9]);
        assert!(video.is_video());
        assert!(!audio.is_image());
    }

    #[test]
    fn media_data_part_from_part_filters_non_media() {
        let xml_part = Part::new_xml(PartUri::new("/doc.xml").unwrap(), b"<doc/>".to_vec());
        assert!(MediaDataPart::from_part(&xml_part).is_none());

        let mut img_part = Part::new(PartUri::new("/media/img.png").unwrap(), vec![1, 2, 3]);
        img_part.content_type = Some("image/png".to_string());
        let media = MediaDataPart::from_part(&img_part).unwrap();
        assert!(media.is_image());
        assert_eq!(media.data(), &[1, 2, 3]);
    }

    #[test]
    fn content_type_from_extension_covers_common_types() {
        assert_eq!(content_type_from_extension("png"), Some("image/png"));
        assert_eq!(content_type_from_extension("JPG"), Some("image/jpeg"));
        assert_eq!(content_type_from_extension("mp3"), Some("audio/mpeg"));
        assert_eq!(content_type_from_extension("mp4"), Some("video/mp4"));
        assert_eq!(content_type_from_extension("xyz"), None);
    }

    #[test]
    fn data_part_to_opc_part() {
        let data_part = DataPart::with_content_type(
            "/media/image1.png",
            "image/png",
            vec![0x89, 0x50, 0x4E, 0x47],
        );

        let opc_part = data_part.to_part().unwrap();
        assert_eq!(opc_part.uri.as_str(), "/media/image1.png");
        assert_eq!(opc_part.content_type.as_deref(), Some("image/png"));
        assert!(opc_part.is_binary());
        assert_eq!(opc_part.data.as_bytes(), &[0x89, 0x50, 0x4E, 0x47]);
    }
}
