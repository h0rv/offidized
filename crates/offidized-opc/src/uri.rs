//! Part URI handling per OPC spec (ECMA-376 Part 2).
//!
//! Part URIs are absolute paths within the package, always starting with `/`.
//! Relationship targets can be relative, resolved against the source part's URI.

use crate::error::Result;

/// A normalized, absolute part URI within an OPC package.
///
/// Examples: `/xl/workbook.xml`, `/word/document.xml`, `/_rels/.rels`
#[derive(Debug, Clone, PartialEq, Eq, Hash, PartialOrd, Ord)]
pub struct PartUri(String);

impl PartUri {
    /// Create a PartUri from a string, normalizing it.
    pub fn new(uri: impl Into<String>) -> Result<Self> {
        let mut uri = uri.into();

        // Normalize: ensure leading slash
        if !uri.starts_with('/') {
            uri = format!("/{uri}");
        }

        // Normalize: remove trailing slash (except root)
        if uri.len() > 1 && uri.ends_with('/') {
            uri.pop();
        }

        // Normalize: collapse double slashes
        while uri.contains("//") {
            uri = uri.replace("//", "/");
        }

        Ok(Self(uri))
    }

    /// Create from a ZIP entry path (no leading slash in ZIP).
    pub fn from_zip_path(path: &str) -> Result<Self> {
        Self::new(format!("/{path}"))
    }

    /// Get the URI as a ZIP entry path (no leading slash).
    pub fn to_zip_path(&self) -> &str {
        &self.0[1..] // strip leading /
    }

    /// Get the directory containing this part.
    pub fn directory(&self) -> &str {
        match self.0.rfind('/') {
            Some(pos) if pos > 0 => &self.0[..pos],
            _ => "/",
        }
    }

    /// Get the filename component.
    pub fn filename(&self) -> &str {
        match self.0.rfind('/') {
            Some(pos) => &self.0[pos + 1..],
            None => &self.0,
        }
    }

    /// Get the file extension (without dot), if any.
    pub fn extension(&self) -> Option<&str> {
        self.filename().rsplit_once('.').map(|(_, ext)| ext)
    }

    /// Get the filename stem (without extension).
    ///
    /// For `/xl/worksheets/sheet1.xml`, returns `"sheet1"`.
    pub fn stem(&self) -> &str {
        let filename = self.filename();
        match filename.rsplit_once('.') {
            Some((stem, _)) => stem,
            None => filename,
        }
    }

    /// Resolve a relative URI against this part's directory.
    ///
    /// For example, resolving `../media/image1.png` against `/xl/worksheets/sheet1.xml`
    /// yields `/xl/media/image1.png`.
    pub fn resolve_relative(&self, relative: &str) -> Result<PartUri> {
        if relative.starts_with('/') {
            return PartUri::new(relative);
        }

        let base_dir = self.directory();
        let mut segments: Vec<&str> = base_dir.split('/').filter(|s| !s.is_empty()).collect();

        for component in relative.split('/') {
            match component {
                "." | "" => {}
                ".." => {
                    segments.pop();
                }
                other => segments.push(other),
            }
        }

        let resolved = format!("/{}", segments.join("/"));
        PartUri::new(resolved)
    }

    /// Get the relationship part URI for this part.
    ///
    /// `/xl/workbook.xml` → `/xl/_rels/workbook.xml.rels`
    pub fn relationship_uri(&self) -> Result<PartUri> {
        PartUri::from_zip_path(&self.relationship_zip_path())
    }

    /// Get the relationship part path as a ZIP entry path.
    ///
    /// `/workbook.xml` → `_rels/workbook.xml.rels`
    /// `/xl/workbook.xml` → `xl/_rels/workbook.xml.rels`
    pub fn relationship_zip_path(&self) -> String {
        let filename = self.filename();

        if self.directory() == "/" {
            format!("_rels/{filename}.rels")
        } else {
            format!(
                "{}/_rels/{filename}.rels",
                self.directory().trim_start_matches('/')
            )
        }
    }

    /// Returns true if a ZIP entry path is a part-level relationships file.
    ///
    /// Accepts both `_rels/*.rels` (root-level parts) and `*/_rels/*.rels` (nested parts),
    /// while excluding package-level relationships (`_rels/.rels`).
    pub fn is_part_relationship_zip_path(path: &str) -> bool {
        if !path.ends_with(".rels") || path == "_rels/.rels" {
            return false;
        }

        path.starts_with("_rels/") || path.contains("/_rels/")
    }

    pub fn as_str(&self) -> &str {
        &self.0
    }
}

impl std::fmt::Display for PartUri {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        write!(f, "{}", self.0)
    }
}

impl AsRef<str> for PartUri {
    fn as_ref(&self) -> &str {
        &self.0
    }
}

/// The well-known URI for the package-level relationships.
pub const PACKAGE_RELS_URI: &str = "/_rels/.rels";

/// The well-known URI for content types.
pub const CONTENT_TYPES_URI: &str = "/[Content_Types].xml";

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_from_zip_path() {
        let uri = PartUri::from_zip_path("xl/workbook.xml").unwrap();
        assert_eq!(uri.as_str(), "/xl/workbook.xml");
        assert_eq!(uri.to_zip_path(), "xl/workbook.xml");
    }

    #[test]
    fn test_directory_and_filename() {
        let uri = PartUri::new("/xl/worksheets/sheet1.xml").unwrap();
        assert_eq!(uri.directory(), "/xl/worksheets");
        assert_eq!(uri.filename(), "sheet1.xml");
        assert_eq!(uri.extension(), Some("xml"));
    }

    #[test]
    fn test_resolve_relative() {
        let uri = PartUri::new("/xl/worksheets/sheet1.xml").unwrap();
        let resolved = uri.resolve_relative("../media/image1.png").unwrap();
        assert_eq!(resolved.as_str(), "/xl/media/image1.png");
    }

    #[test]
    fn test_relationship_uri() {
        let uri = PartUri::new("/xl/workbook.xml").unwrap();
        let rels = uri.relationship_uri().unwrap();
        assert_eq!(rels.as_str(), "/xl/_rels/workbook.xml.rels");
        assert_eq!(uri.relationship_zip_path(), "xl/_rels/workbook.xml.rels");
    }

    #[test]
    fn test_relationship_uri_root_level() {
        let uri = PartUri::new("/workbook.xml").unwrap();
        let rels = uri.relationship_uri().unwrap();
        assert_eq!(rels.as_str(), "/_rels/workbook.xml.rels");
        assert_eq!(uri.relationship_zip_path(), "_rels/workbook.xml.rels");
    }

    #[test]
    fn test_stem() {
        let uri = PartUri::new("/xl/worksheets/sheet1.xml").unwrap();
        assert_eq!(uri.stem(), "sheet1");

        let uri = PartUri::new("/media/image").unwrap();
        assert_eq!(uri.stem(), "image");
    }

    #[test]
    fn test_is_part_relationship_zip_path() {
        assert!(PartUri::is_part_relationship_zip_path(
            "_rels/workbook.xml.rels"
        ));
        assert!(PartUri::is_part_relationship_zip_path(
            "xl/_rels/workbook.xml.rels"
        ));
        assert!(!PartUri::is_part_relationship_zip_path("_rels/.rels"));
        assert!(!PartUri::is_part_relationship_zip_path("xl/workbook.xml"));
    }
}
