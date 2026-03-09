use serde::{Deserialize, Serialize};

use crate::types::Hyperlink;

/// A drawing object (image, chart, or shape) in a worksheet
#[derive(Debug, Serialize, Deserialize, Clone)]
#[serde(rename_all = "camelCase")]
pub struct Drawing {
    /// Type of anchor: "twoCellAnchor", "oneCellAnchor", or "absoluteAnchor"
    pub anchor_type: String,
    /// Type of drawing: "picture", "chart", "shape", or "textbox"
    pub drawing_type: String,
    /// Name of the drawing object
    #[serde(skip_serializing_if = "Option::is_none")]
    pub name: Option<String>,
    /// Description/alt text for the drawing
    #[serde(skip_serializing_if = "Option::is_none")]
    pub description: Option<String>,
    /// Title for the drawing
    #[serde(skip_serializing_if = "Option::is_none")]
    pub title: Option<String>,
    /// Starting column (0-indexed)
    #[serde(skip_serializing_if = "Option::is_none")]
    pub from_col: Option<u32>,
    /// Starting row (0-indexed)
    #[serde(skip_serializing_if = "Option::is_none")]
    pub from_row: Option<u32>,
    /// Column offset in EMUs (English Metric Units)
    #[serde(skip_serializing_if = "Option::is_none")]
    pub from_col_off: Option<i64>,
    /// Row offset in EMUs
    #[serde(skip_serializing_if = "Option::is_none")]
    pub from_row_off: Option<i64>,
    /// Ending column (0-indexed, for twoCellAnchor)
    #[serde(skip_serializing_if = "Option::is_none")]
    pub to_col: Option<u32>,
    /// Ending row (0-indexed, for twoCellAnchor)
    #[serde(skip_serializing_if = "Option::is_none")]
    pub to_row: Option<u32>,
    /// Column offset in EMUs for ending position
    #[serde(skip_serializing_if = "Option::is_none")]
    pub to_col_off: Option<i64>,
    /// Row offset in EMUs for ending position
    #[serde(skip_serializing_if = "Option::is_none")]
    pub to_row_off: Option<i64>,
    /// Absolute X position in EMUs (for absoluteAnchor)
    #[serde(skip_serializing_if = "Option::is_none")]
    pub pos_x: Option<i64>,
    /// Absolute Y position in EMUs (for absoluteAnchor)
    #[serde(skip_serializing_if = "Option::is_none")]
    pub pos_y: Option<i64>,
    /// Width in EMUs (for oneCellAnchor and absoluteAnchor)
    #[serde(skip_serializing_if = "Option::is_none")]
    pub extent_cx: Option<i64>,
    /// Height in EMUs (for oneCellAnchor and absoluteAnchor)
    #[serde(skip_serializing_if = "Option::is_none")]
    pub extent_cy: Option<i64>,
    /// How the drawing behaves when cells are moved/resized
    #[serde(skip_serializing_if = "Option::is_none")]
    pub edit_as: Option<String>,
    /// Relationship ID for the image (rId)
    #[serde(skip_serializing_if = "Option::is_none")]
    pub image_id: Option<String>,
    /// Relationship ID for the chart (rId)
    #[serde(skip_serializing_if = "Option::is_none")]
    pub chart_id: Option<String>,
    /// Shape type (for shapes)
    #[serde(skip_serializing_if = "Option::is_none")]
    pub shape_type: Option<String>,
    /// Fill color for shapes
    #[serde(skip_serializing_if = "Option::is_none")]
    pub fill_color: Option<String>,
    /// Line/stroke color for shapes
    #[serde(skip_serializing_if = "Option::is_none")]
    pub line_color: Option<String>,
    /// Text content for shapes
    #[serde(skip_serializing_if = "Option::is_none")]
    pub text_content: Option<String>,
    /// Rotation angle in 1/60000th of a degree (60000 units = 1 degree)
    #[serde(skip_serializing_if = "Option::is_none")]
    pub rotation: Option<i64>,
    /// Horizontal flip
    #[serde(skip_serializing_if = "Option::is_none")]
    pub flip_h: Option<bool>,
    /// Vertical flip
    #[serde(skip_serializing_if = "Option::is_none")]
    pub flip_v: Option<bool>,
    /// Hyperlink associated with this drawing
    #[serde(skip_serializing_if = "Option::is_none")]
    pub hyperlink: Option<Hyperlink>,
    /// Transform offset X in EMUs (from a:xfrm/a:off) - precise position from Excel
    #[serde(skip_serializing_if = "Option::is_none")]
    pub xfrm_x: Option<i64>,
    /// Transform offset Y in EMUs (from a:xfrm/a:off) - precise position from Excel
    #[serde(skip_serializing_if = "Option::is_none")]
    pub xfrm_y: Option<i64>,
    /// Transform extent width in EMUs (from a:xfrm/a:ext) - precise size from Excel
    #[serde(skip_serializing_if = "Option::is_none")]
    pub xfrm_cx: Option<i64>,
    /// Transform extent height in EMUs (from a:xfrm/a:ext) - precise size from Excel
    #[serde(skip_serializing_if = "Option::is_none")]
    pub xfrm_cy: Option<i64>,
}

/// An embedded image extracted from the XLSX file
#[derive(Debug, Serialize, Deserialize, Clone)]
#[serde(rename_all = "camelCase")]
pub struct EmbeddedImage {
    /// Unique identifier for the image (usually the path within the archive)
    pub id: String,
    /// MIME type of the image (e.g., "image/png", "image/jpeg")
    pub mime_type: String,
    /// Base64-encoded image data
    pub data: String,
    /// Original filename within the XLSX archive
    #[serde(skip_serializing_if = "Option::is_none")]
    pub filename: Option<String>,
    /// Width in pixels (if known from image metadata)
    #[serde(skip_serializing_if = "Option::is_none")]
    pub width: Option<u32>,
    /// Height in pixels (if known from image metadata)
    #[serde(skip_serializing_if = "Option::is_none")]
    pub height: Option<u32>,
}

/// Image format/MIME type detection
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ImageFormat {
    Png,
    Jpeg,
    Gif,
    Bmp,
    Tiff,
    Webp,
    Emf,
    Wmf,
    Unknown,
}

impl ImageFormat {
    /// Detect image format from file extension
    #[must_use]
    pub fn from_extension(ext: &str) -> Self {
        match ext.to_lowercase().as_str() {
            "png" => Self::Png,
            "jpg" | "jpeg" => Self::Jpeg,
            "gif" => Self::Gif,
            "bmp" => Self::Bmp,
            "tif" | "tiff" => Self::Tiff,
            "webp" => Self::Webp,
            "emf" => Self::Emf,
            "wmf" => Self::Wmf,
            _ => Self::Unknown,
        }
    }

    /// Detect image format from magic bytes
    #[must_use]
    pub fn from_magic_bytes(data: &[u8]) -> Self {
        if data.len() < 4 {
            return Self::Unknown;
        }

        // PNG: 89 50 4E 47
        if data.starts_with(&[0x89, 0x50, 0x4E, 0x47]) {
            return Self::Png;
        }

        // JPEG: FF D8 FF
        if data.starts_with(&[0xFF, 0xD8, 0xFF]) {
            return Self::Jpeg;
        }

        // GIF: GIF87a or GIF89a
        if data.starts_with(b"GIF87a") || data.starts_with(b"GIF89a") {
            return Self::Gif;
        }

        // BMP: BM
        if data.starts_with(b"BM") {
            return Self::Bmp;
        }

        // TIFF: II or MM
        if data.starts_with(&[0x49, 0x49, 0x2A, 0x00])
            || data.starts_with(&[0x4D, 0x4D, 0x00, 0x2A])
        {
            return Self::Tiff;
        }

        // WebP: RIFF....WEBP
        if data.len() >= 12 && data.starts_with(b"RIFF") && data.get(8..12) == Some(b"WEBP") {
            return Self::Webp;
        }

        // EMF: 01 00 00 00
        if data.starts_with(&[0x01, 0x00, 0x00, 0x00]) && data.len() >= 40 {
            return Self::Emf;
        }

        // WMF: D7 CD C6 9A
        if data.starts_with(&[0xD7, 0xCD, 0xC6, 0x9A]) {
            return Self::Wmf;
        }

        Self::Unknown
    }

    /// Get MIME type for this image format
    #[must_use]
    pub fn mime_type(&self) -> &'static str {
        match self {
            Self::Png => "image/png",
            Self::Jpeg => "image/jpeg",
            Self::Gif => "image/gif",
            Self::Bmp => "image/bmp",
            Self::Tiff => "image/tiff",
            Self::Webp => "image/webp",
            Self::Emf => "image/x-emf",
            Self::Wmf => "image/x-wmf",
            Self::Unknown => "application/octet-stream",
        }
    }
}
