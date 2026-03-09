//! Part extension mapping — maps content types to file extensions.
//!
//! This mirrors `PartExtensionProvider` from the Open XML SDK. Given a content
//! type, it returns the standard file extension used in OPC packages.

use std::collections::HashMap;

/// Maps content types to their standard file extensions.
///
/// OOXML defines standard extensions for each content type. For example,
/// worksheet parts use `.xml`, images use `.png`/`.jpg`, etc.
#[derive(Debug, Clone)]
pub struct PartExtensionMap {
    map: HashMap<String, String>,
}

impl Default for PartExtensionMap {
    fn default() -> Self {
        Self::new()
    }
}

impl PartExtensionMap {
    /// Create a new map pre-populated with standard OOXML extensions.
    pub fn new() -> Self {
        let mut map = HashMap::new();

        // All standard XML-based content types default to .xml
        for ct in [
            // SpreadsheetML
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml",
            "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml",
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml",
            "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml",
            "application/vnd.openxmlformats-officedocument.spreadsheetml.chartsheet+xml",
            "application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml",
            "application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml",
            "application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheDefinition+xml",
            "application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheRecords+xml",
            "application/vnd.openxmlformats-officedocument.spreadsheetml.externalLink+xml",
            "application/vnd.openxmlformats-officedocument.spreadsheetml.calcChain+xml",
            // WordprocessingML
            "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml",
            "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml",
            "application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml",
            "application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml",
            "application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml",
            "application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml",
            "application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml",
            "application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml",
            "application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml",
            "application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml",
            "application/vnd.openxmlformats-officedocument.wordprocessingml.webSettings+xml",
            // PresentationML
            "application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml",
            "application/vnd.openxmlformats-officedocument.presentationml.slide+xml",
            "application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml",
            "application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml",
            "application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml",
            "application/vnd.openxmlformats-officedocument.presentationml.notesMaster+xml",
            "application/vnd.openxmlformats-officedocument.presentationml.handoutMaster+xml",
            "application/vnd.openxmlformats-officedocument.presentationml.comments+xml",
            "application/vnd.openxmlformats-officedocument.presentationml.tags+xml",
            // DrawingML
            "application/vnd.openxmlformats-officedocument.drawing+xml",
            "application/vnd.openxmlformats-officedocument.drawingml.chart+xml",
            "application/vnd.openxmlformats-officedocument.drawingml.diagramColors+xml",
            "application/vnd.openxmlformats-officedocument.drawingml.diagramData+xml",
            "application/vnd.openxmlformats-officedocument.drawingml.diagramLayout+xml",
            "application/vnd.openxmlformats-officedocument.drawingml.diagramStyle+xml",
            "application/vnd.openxmlformats-officedocument.theme+xml",
            "application/vnd.openxmlformats-officedocument.themeOverride+xml",
            // Shared
            "application/vnd.openxmlformats-package.core-properties+xml",
            "application/vnd.openxmlformats-officedocument.extended-properties+xml",
            "application/vnd.openxmlformats-officedocument.custom-properties+xml",
            "application/vnd.openxmlformats-officedocument.customXmlProperties+xml",
            // Generic
            "application/xml",
        ] {
            map.insert(ct.to_string(), ".xml".to_string());
        }

        // Relationship files
        map.insert(
            "application/vnd.openxmlformats-package.relationships+xml".to_string(),
            ".rels".to_string(),
        );

        // VML Drawing
        map.insert(
            "application/vnd.openxmlformats-officedocument.vmlDrawing".to_string(),
            ".vml".to_string(),
        );

        // Binary/macro
        map.insert(
            "application/vnd.ms-office.vbaProject".to_string(),
            ".bin".to_string(),
        );

        // Printer settings
        for ct in [
            "application/vnd.openxmlformats-officedocument.spreadsheetml.printerSettings",
            "application/vnd.openxmlformats-officedocument.wordprocessingml.printerSettings",
            "application/vnd.openxmlformats-officedocument.presentationml.printerSettings",
        ] {
            map.insert(ct.to_string(), ".bin".to_string());
        }

        Self { map }
    }

    /// Get the file extension for a content type.
    pub fn get_extension(&self, content_type: &str) -> Option<&str> {
        self.map.get(content_type).map(String::as_str)
    }

    /// Register a custom content type → extension mapping.
    pub fn register(&mut self, content_type: impl Into<String>, extension: impl Into<String>) {
        self.map.insert(content_type.into(), extension.into());
    }

    /// Check if a content type has a registered extension.
    pub fn contains(&self, content_type: &str) -> bool {
        self.map.contains_key(content_type)
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn standard_xml_content_types_map_to_xml_extension() {
        let map = PartExtensionMap::new();
        assert_eq!(
            map.get_extension(
                "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"
            ),
            Some(".xml")
        );
        assert_eq!(
            map.get_extension(
                "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"
            ),
            Some(".xml")
        );
        assert_eq!(
            map.get_extension("application/vnd.openxmlformats-officedocument.theme+xml"),
            Some(".xml")
        );
    }

    #[test]
    fn binary_content_types_map_to_bin_extension() {
        let map = PartExtensionMap::new();
        assert_eq!(
            map.get_extension("application/vnd.ms-office.vbaProject"),
            Some(".bin")
        );
        assert_eq!(
            map.get_extension(
                "application/vnd.openxmlformats-officedocument.spreadsheetml.printerSettings"
            ),
            Some(".bin")
        );
    }

    #[test]
    fn custom_mapping_overrides_default() {
        let mut map = PartExtensionMap::new();
        map.register("application/custom", ".custom");
        assert_eq!(map.get_extension("application/custom"), Some(".custom"));
    }

    #[test]
    fn unknown_content_type_returns_none() {
        let map = PartExtensionMap::new();
        assert_eq!(map.get_extension("application/totally-unknown"), None);
    }
}
