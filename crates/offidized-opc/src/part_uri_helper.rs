//! Helper for creating deterministic and unique OPC part URIs.
//!
//! This mirrors the behavior of Open-XML-SDK's `PartUriHelper`, adapted to
//! `PartUri` and Rust error handling.

use std::collections::{BTreeMap, BTreeSet};

use crate::error::{OpcError, Result};
use crate::uri::PartUri;

// Original constants used in tests.
#[cfg(test)]
const WORD_FOOTER_CONTENT_TYPE: &str =
    "application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml";
#[cfg(test)]
const WORD_HEADER_CONTENT_TYPE: &str =
    "application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml";
#[cfg(test)]
const SPREADSHEET_WORKSHEET_CONTENT_TYPE: &str =
    "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml";

/// All content types from the Open XML SDK that use numbered part URIs
/// (i.e. the first generated URI gets suffix `1` instead of no suffix).
const NUMBERED_CONTENT_TYPES: &[&str] = &[
    // WordprocessingML
    "application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml",
    "application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml",
    // SpreadsheetML
    "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml",
    "application/vnd.openxmlformats-officedocument.spreadsheetml.chartsheet+xml",
    "application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml",
    "application/vnd.openxmlformats-officedocument.spreadsheetml.dialogsheet+xml",
    "application/vnd.openxmlformats-officedocument.drawing+xml",
    "application/vnd.openxmlformats-officedocument.spreadsheetml.externalLink+xml",
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheetMetadata+xml",
    "application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheDefinition+xml",
    "application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheRecords+xml",
    "application/vnd.openxmlformats-officedocument.spreadsheetml.queryTable+xml",
    "application/vnd.openxmlformats-officedocument.spreadsheetml.revisionLog+xml",
    "application/vnd.openxmlformats-officedocument.spreadsheetml.tableSingleCells+xml",
    "application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml",
    // PresentationML
    "application/vnd.openxmlformats-officedocument.presentationml.comments+xml",
    "application/vnd.openxmlformats-officedocument.presentationml.handoutMaster+xml",
    "application/vnd.openxmlformats-officedocument.presentationml.notesMaster+xml",
    "application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml",
    "application/vnd.openxmlformats-officedocument.presentationml.slide+xml",
    "application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml",
    "application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml",
    "application/vnd.openxmlformats-officedocument.presentationml.slideUpdateInfo+xml",
    "application/vnd.openxmlformats-officedocument.presentationml.tags+xml",
    // DrawingML
    "application/vnd.openxmlformats-officedocument.drawingml.chart+xml",
    "application/vnd.openxmlformats-officedocument.drawingml.chartshapes+xml",
    "application/vnd.openxmlformats-officedocument.drawingml.diagramColors+xml",
    "application/vnd.openxmlformats-officedocument.drawingml.diagramData+xml",
    "application/vnd.openxmlformats-officedocument.drawingml.diagramLayout+xml",
    "application/vnd.openxmlformats-officedocument.drawingml.diagramStyle+xml",
    "application/vnd.openxmlformats-officedocument.theme+xml",
    "application/vnd.openxmlformats-officedocument.themeOverride+xml",
    // Shared
    "application/vnd.openxmlformats-officedocument.customXmlProperties+xml",
    "application/vnd.openxmlformats-officedocument.spreadsheetml.printerSettings",
    "application/vnd.openxmlformats-officedocument.wordprocessingml.printerSettings",
    "application/vnd.openxmlformats-officedocument.presentationml.printerSettings",
];

/// Tracks sequence numbers per content type and reserved part URIs.
///
/// Sequence behavior follows Open-XML-SDK:
/// - Most content types produce no suffix on the first unique attempt.
/// - Numbered content types start at suffix `1` on the first attempt.
#[derive(Debug, Default, Clone)]
pub struct PartUriHelper {
    sequence_numbers: BTreeMap<String, u64>,
    reserved_uris: BTreeSet<PartUri>,
}

impl PartUriHelper {
    /// Create a new helper with no reserved URIs.
    pub fn new() -> Self {
        Self::default()
    }

    /// Reserve an existing URI and advance sequence tracking for the content type.
    pub fn reserve_uri(&mut self, content_type: &str, part_uri: &PartUri) {
        let key = Self::normalized_content_type(content_type);
        match self.sequence_numbers.get_mut(&key) {
            Some(value) => {
                *value = value.saturating_add(1);
            }
            None => {
                self.sequence_numbers.insert(key, 1);
            }
        }

        self.reserved_uris.insert(part_uri.clone());
    }

    /// Create a part URI from a parent URI and target components.
    ///
    /// When `force_unique` is true, sequence suffixes are appended according to
    /// content type rules until an unreserved URI is found.
    pub fn create_part_uri(
        &mut self,
        content_type: &str,
        parent: &PartUri,
        target_path: &str,
        target_name: &str,
        target_ext: &str,
        force_unique: bool,
    ) -> Result<PartUri> {
        let normalized_ext = Self::normalize_extension(target_ext);

        if force_unique {
            loop {
                let sequence = self.next_sequence_suffix(content_type)?;
                let file_name = format!("{target_name}{sequence}{normalized_ext}");
                let relative = Self::combine_target_path(target_path, &file_name);
                let candidate = parent.resolve_relative(&relative)?;

                if !self.reserved_uris.contains(&candidate) {
                    self.reserved_uris.insert(candidate.clone());
                    return Ok(candidate);
                }
            }
        }

        let file_name = format!("{target_name}{normalized_ext}");
        let relative = Self::combine_target_path(target_path, &file_name);
        let candidate = parent.resolve_relative(&relative)?;

        if self.reserved_uris.contains(&candidate) {
            return Err(OpcError::InvalidUri(format!(
                "Part URI is already reserved: {}",
                candidate.as_str()
            )));
        }

        self.reserved_uris.insert(candidate.clone());
        Ok(candidate)
    }

    /// Ensure a unique URI for an existing target URI relative to `parent`.
    ///
    /// This follows the Open-XML-SDK flow:
    /// 1. Resolve `target_uri` against `parent`.
    /// 2. Rebuild the URI using `.` plus filename stem/ext with `force_unique=true`.
    pub fn ensure_unique_part_uri(
        &mut self,
        content_type: &str,
        parent: &PartUri,
        target_uri: &str,
    ) -> Result<PartUri> {
        let resolved_target = parent.resolve_relative(target_uri)?;
        let (target_name, target_ext) = Self::split_target_filename(target_uri)?;

        self.create_part_uri(
            content_type,
            &resolved_target,
            ".",
            &target_name,
            &target_ext,
            true,
        )
    }

    fn next_sequence_suffix(&mut self, content_type: &str) -> Result<String> {
        let key = Self::normalized_content_type(content_type);

        if let Some(value) = self.sequence_numbers.get_mut(&key) {
            let next = value.checked_add(1).ok_or_else(|| {
                OpcError::InvalidUri(format!(
                    "Sequence number overflow for content type: {content_type}"
                ))
            })?;
            *value = next;
            return Ok(next.to_string());
        }

        self.sequence_numbers.insert(key.clone(), 1);
        if Self::is_numbered_content_type(&key) {
            Ok("1".to_string())
        } else {
            Ok(String::new())
        }
    }

    fn normalized_content_type(content_type: &str) -> String {
        content_type.trim().to_ascii_lowercase()
    }

    fn is_numbered_content_type(content_type: &str) -> bool {
        NUMBERED_CONTENT_TYPES.contains(&content_type)
    }

    fn normalize_extension(extension: &str) -> String {
        if extension.is_empty() {
            String::new()
        } else if extension.starts_with('.') {
            extension.to_string()
        } else {
            format!(".{extension}")
        }
    }

    fn combine_target_path(target_path: &str, file_name: &str) -> String {
        let path = target_path.replace('\\', "/");
        if path.is_empty() || path == "." {
            return file_name.to_string();
        }

        let trimmed = path.trim_end_matches('/');
        if trimmed.is_empty() {
            format!("/{file_name}")
        } else if trimmed == "." {
            file_name.to_string()
        } else {
            format!("{trimmed}/{file_name}")
        }
    }

    fn split_target_filename(target_uri: &str) -> Result<(String, String)> {
        if target_uri.ends_with('/') || target_uri.ends_with('\\') {
            return Err(OpcError::InvalidUri(format!(
                "Target URI must include a filename: {target_uri}"
            )));
        }

        let normalized = target_uri.replace('\\', "/");
        let file_name = normalized
            .rsplit('/')
            .next()
            .filter(|segment| !segment.is_empty())
            .ok_or_else(|| {
                OpcError::InvalidUri(format!("Target URI must include a filename: {target_uri}"))
            })?;

        if file_name == "." || file_name == ".." {
            return Err(OpcError::InvalidUri(format!(
                "Target URI must include a real filename: {target_uri}"
            )));
        }

        if let Some((stem, ext)) = file_name.rsplit_once('.') {
            if !stem.is_empty() {
                return Ok((stem.to_string(), format!(".{ext}")));
            }
        }

        Ok((file_name.to_string(), String::new()))
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn unknown_content_type_force_unique_matches_openxml_behavior() {
        let mut helper = PartUriHelper::new();
        let parent = PartUri::new("/").expect("valid parent URI");

        let first = helper
            .ensure_unique_part_uri("unknown", &parent, "/target")
            .expect("first unique URI");
        assert_eq!(first.as_str(), "/target");

        let second = helper
            .ensure_unique_part_uri("unknown", &parent, "/target")
            .expect("second unique URI");
        assert_eq!(second.as_str(), "/target2");
    }

    #[test]
    fn numbered_word_header_starts_at_one() {
        let mut helper = PartUriHelper::new();
        let parent = PartUri::new("/word/document.xml").expect("valid parent URI");

        let first = helper
            .ensure_unique_part_uri(WORD_HEADER_CONTENT_TYPE, &parent, "header.xml")
            .expect("first unique header URI");
        assert_eq!(first.as_str(), "/word/header1.xml");

        let second = helper
            .ensure_unique_part_uri(WORD_HEADER_CONTENT_TYPE, &parent, "header.xml")
            .expect("second unique header URI");
        assert_eq!(second.as_str(), "/word/header2.xml");
    }

    #[test]
    fn numbered_word_footer_starts_at_one() {
        let mut helper = PartUriHelper::new();
        let parent = PartUri::new("/word/document.xml").expect("valid parent URI");

        let first = helper
            .ensure_unique_part_uri(WORD_FOOTER_CONTENT_TYPE, &parent, "footer.xml")
            .expect("first unique footer URI");
        assert_eq!(first.as_str(), "/word/footer1.xml");

        let second = helper
            .ensure_unique_part_uri(WORD_FOOTER_CONTENT_TYPE, &parent, "footer.xml")
            .expect("second unique footer URI");
        assert_eq!(second.as_str(), "/word/footer2.xml");
    }

    #[test]
    fn numbered_spreadsheet_worksheet_starts_at_one() {
        let mut helper = PartUriHelper::new();
        let parent = PartUri::new("/xl/workbook.xml").expect("valid parent URI");

        let first = helper
            .create_part_uri(
                SPREADSHEET_WORKSHEET_CONTENT_TYPE,
                &parent,
                "worksheets",
                "sheet",
                ".xml",
                true,
            )
            .expect("first unique worksheet URI");
        assert_eq!(first.as_str(), "/xl/worksheets/sheet1.xml");

        let second = helper
            .create_part_uri(
                SPREADSHEET_WORKSHEET_CONTENT_TYPE,
                &parent,
                "worksheets",
                "sheet",
                ".xml",
                true,
            )
            .expect("second unique worksheet URI");
        assert_eq!(second.as_str(), "/xl/worksheets/sheet2.xml");
    }

    #[test]
    fn reserve_uri_advances_sequence_and_prevents_reuse() {
        let mut helper = PartUriHelper::new();
        let reserved = PartUri::new("/target").expect("valid reserved URI");

        helper.reserve_uri("unknown", &reserved);

        let unique = helper
            .create_part_uri("unknown", &reserved, ".", "target", "", true)
            .expect("unique URI after reservation");
        assert_eq!(unique.as_str(), "/target2");
    }

    #[test]
    fn numbered_content_types_array_has_expected_entries() {
        // 3 original (footer, header, worksheet) + 33 additional from Open XML SDK = 36 total
        assert_eq!(NUMBERED_CONTENT_TYPES.len(), 36);
    }

    #[test]
    fn presentation_slide_is_numbered_and_starts_at_one() {
        let slide_ct = "application/vnd.openxmlformats-officedocument.presentationml.slide+xml";
        let mut helper = PartUriHelper::new();
        let parent = PartUri::new("/ppt/presentation.xml").expect("valid parent URI");

        let first = helper
            .create_part_uri(slide_ct, &parent, "slides", "slide", ".xml", true)
            .expect("first unique slide URI");
        assert_eq!(first.as_str(), "/ppt/slides/slide1.xml");

        let second = helper
            .create_part_uri(slide_ct, &parent, "slides", "slide", ".xml", true)
            .expect("second unique slide URI");
        assert_eq!(second.as_str(), "/ppt/slides/slide2.xml");
    }

    #[test]
    fn unknown_content_type_starts_with_no_suffix() {
        let mut helper = PartUriHelper::new();
        let parent = PartUri::new("/").expect("valid parent URI");

        let first = helper
            .create_part_uri("application/x-custom", &parent, ".", "thing", ".bin", true)
            .expect("first unique URI for unknown type");
        assert_eq!(first.as_str(), "/thing.bin");

        let second = helper
            .create_part_uri("application/x-custom", &parent, ".", "thing", ".bin", true)
            .expect("second unique URI for unknown type");
        assert_eq!(second.as_str(), "/thing2.bin");
    }

    #[test]
    fn all_numbered_content_types_are_recognized() {
        for ct in NUMBERED_CONTENT_TYPES {
            assert!(
                PartUriHelper::is_numbered_content_type(ct),
                "Content type should be numbered: {ct}"
            );
        }
    }

    #[test]
    fn non_numbered_content_type_is_not_recognized() {
        assert!(!PartUriHelper::is_numbered_content_type("application/xml"));
        assert!(!PartUriHelper::is_numbered_content_type(
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"
        ));
    }
}
