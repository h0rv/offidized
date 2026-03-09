//! Configuration for opening OPC packages.
//!
//! Maps to `OpenSettings` in the Open XML SDK. Controls behavior when
//! reading packages, such as maximum part size limits.

/// Office compatibility level for OOXML packages.
///
/// Maps to `MarkupCompatibilityProcessSettings.TargetFileFormatVersions`
/// in the Open XML SDK. Determines which features and namespaces are expected.
#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord, Hash, Default)]
pub enum CompatibilityLevel {
    /// ECMA-376 1st edition (Office 2007).
    Office2007,
    /// ECMA-376 2nd edition (Office 2010).
    Office2010,
    /// ECMA-376 4th edition (Office 2013).
    #[default]
    Office2013,
    /// Office 2016 extensions.
    Office2016,
    /// Office 2019 extensions.
    Office2019,
    /// Office 2021 extensions.
    Office2021,
}

impl CompatibilityLevel {
    /// Year representation of this compatibility level.
    pub fn year(self) -> u16 {
        match self {
            Self::Office2007 => 2007,
            Self::Office2010 => 2010,
            Self::Office2013 => 2013,
            Self::Office2016 => 2016,
            Self::Office2019 => 2019,
            Self::Office2021 => 2021,
        }
    }
}

/// How the markup compatibility preprocessor handles unrecognized elements.
///
/// Maps to `MarkupCompatibilityProcessMode` in the Open XML SDK.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum MarkupCompatibilityProcessMode {
    /// No markup compatibility processing — preserve everything as-is.
    #[default]
    NoProcess,
    /// Process markup compatibility elements according to the target version.
    ProcessAllParts,
    /// Process only loaded parts (skip parts not explicitly accessed).
    ProcessLoadedPartsOnly,
}

/// Configuration settings for opening OPC packages.
///
/// Mirrors `OpenSettings` from the Open XML SDK. Use this to control
/// behavior when reading packages, such as imposing part size limits.
#[derive(Debug, Clone)]
pub struct OpenSettings {
    /// Maximum number of characters allowed in a single XML part.
    ///
    /// If a part exceeds this limit, an error is returned during loading.
    /// Set to `0` (default) to disable the limit.
    pub max_characters_in_part: u64,

    /// Whether to automatically save changes when closing.
    ///
    /// Not currently used — provided for API compatibility with Open XML SDK.
    pub auto_save: bool,

    /// Whether to mark the package as read-only after opening.
    ///
    /// When `true`, mutations to the package will be rejected.
    /// Default: `false`.
    pub read_only: bool,

    /// Whether to strictly validate relationships on open.
    ///
    /// When `true`, invalid relationship targets cause an error.
    /// When `false` (default), invalid targets are silently ignored.
    pub strict_relationship_validation: bool,

    /// Target Office version for markup compatibility processing.
    ///
    /// Determines which features and namespaces are expected when processing
    /// markup compatibility elements.
    pub compatibility_level: CompatibilityLevel,

    /// How to handle markup compatibility elements.
    ///
    /// Controls whether `mc:AlternateContent`, `mc:Choice`, and `mc:Fallback`
    /// elements are processed or preserved as-is.
    pub markup_compatibility_process_mode: MarkupCompatibilityProcessMode,
}

impl Default for OpenSettings {
    fn default() -> Self {
        Self {
            max_characters_in_part: 0,
            auto_save: true,
            read_only: false,
            strict_relationship_validation: false,
            compatibility_level: CompatibilityLevel::default(),
            markup_compatibility_process_mode: MarkupCompatibilityProcessMode::default(),
        }
    }
}

impl OpenSettings {
    /// Create default settings.
    pub fn new() -> Self {
        Self::default()
    }

    /// Create settings with a character limit on parts.
    pub fn with_max_characters(max: u64) -> Self {
        Self {
            max_characters_in_part: max,
            ..Default::default()
        }
    }

    /// Create read-only settings.
    pub fn read_only() -> Self {
        Self {
            read_only: true,
            ..Default::default()
        }
    }

    /// Create settings targeting a specific Office version.
    pub fn with_compatibility_level(level: CompatibilityLevel) -> Self {
        Self {
            compatibility_level: level,
            ..Default::default()
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn compatibility_level_ordering() {
        assert!(CompatibilityLevel::Office2007 < CompatibilityLevel::Office2013);
        assert!(CompatibilityLevel::Office2013 < CompatibilityLevel::Office2021);
    }

    #[test]
    fn compatibility_level_years() {
        assert_eq!(CompatibilityLevel::Office2007.year(), 2007);
        assert_eq!(CompatibilityLevel::Office2021.year(), 2021);
    }

    #[test]
    fn default_settings() {
        let settings = OpenSettings::default();
        assert_eq!(settings.compatibility_level, CompatibilityLevel::Office2013);
        assert_eq!(
            settings.markup_compatibility_process_mode,
            MarkupCompatibilityProcessMode::NoProcess
        );
    }
}
