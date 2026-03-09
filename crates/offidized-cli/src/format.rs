use std::path::Path;

use anyhow::{bail, Result};

/// Detected file format based on extension.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum FileFormat {
    Xlsx,
    Docx,
    Pptx,
}

impl FileFormat {
    /// Detect format from file extension.
    pub fn detect(path: &Path) -> Result<Self> {
        match path
            .extension()
            .and_then(|ext| ext.to_str())
            .map(|ext| ext.to_ascii_lowercase())
            .as_deref()
        {
            Some("xlsx") => Ok(Self::Xlsx),
            Some("docx") => Ok(Self::Docx),
            Some("pptx") => Ok(Self::Pptx),
            Some(other) => bail!("unsupported file extension: .{other}"),
            None => bail!("cannot detect format: file has no extension"),
        }
    }
}
