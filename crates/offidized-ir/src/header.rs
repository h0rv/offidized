//! IR header parsing and writing.
//!
//! Every IR file starts with a TOML front matter block using `+++` delimiters:
//!
//! ```text
//! +++
//! source = "report.xlsx"
//! format = "xlsx"
//! mode = "content"
//! version = 1
//! checksum = "sha256:a1b2c3d4..."
//! +++
//! ```

use crate::{Format, IrError, Mode, Result};

/// Metadata header for an IR file.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct IrHeader {
    /// Path to the original source file.
    pub source: String,
    /// Office file format.
    pub format: Format,
    /// IR mode (content, style, full).
    pub mode: Mode,
    /// IR format version (always 1 for now).
    pub version: u32,
    /// SHA-256 checksum of the source file at derive time.
    pub checksum: String,
}

impl IrHeader {
    /// Serialize the header as a TOML front matter block.
    pub fn write(&self) -> String {
        let mut s = String::with_capacity(256);
        s.push_str("+++\n");
        s.push_str("source = \"");
        s.push_str(&self.source);
        s.push_str("\"\n");
        s.push_str("format = \"");
        s.push_str(self.format.as_str());
        s.push_str("\"\n");
        s.push_str("mode = \"");
        s.push_str(self.mode.as_str());
        s.push_str("\"\n");
        s.push_str("version = ");
        s.push_str(&self.version.to_string());
        s.push('\n');
        s.push_str("checksum = \"");
        s.push_str(&self.checksum);
        s.push_str("\"\n");
        s.push_str("+++\n");
        s
    }

    /// Parse a TOML front matter block from the beginning of an IR string.
    ///
    /// Returns the parsed header and the remaining body text.
    pub fn parse(input: &str) -> Result<(Self, String)> {
        let trimmed = input.trim_start();

        // Find opening +++
        if !trimmed.starts_with("+++") {
            return Err(IrError::InvalidHeader("missing opening +++".into()));
        }
        let after_open = skip_line(trimmed);

        // Find closing +++ line
        let mut header_end = 0;
        let mut found_close = false;
        for line in after_open.lines() {
            if line.trim() == "+++" {
                found_close = true;
                break;
            }
            // +1 for the \n that lines() strips
            header_end += line.len() + 1;
        }

        if !found_close {
            return Err(IrError::InvalidHeader("missing closing +++".into()));
        }

        let header_text = &after_open[..header_end];

        // Body starts after the closing +++ line
        let after_close = &after_open[header_end..];
        let body = skip_line(after_close);

        // Parse key=value pairs
        let mut source = None;
        let mut format = None;
        let mut mode = None;
        let mut version = None;
        let mut checksum = None;

        for line in header_text.lines() {
            let line = line.trim();
            if line.is_empty() {
                continue;
            }

            if let Some((key, value)) = line.split_once('=') {
                let key = key.trim();
                let value = value.trim().trim_matches('"');
                match key {
                    "source" => source = Some(value.to_string()),
                    "format" => format = Some(value.to_string()),
                    "mode" => mode = Some(value.to_string()),
                    "version" => version = Some(value.to_string()),
                    "checksum" => checksum = Some(value.to_string()),
                    _ => {} // ignore unknown keys for forward compatibility
                }
            }
        }

        let format_str =
            format.ok_or_else(|| IrError::InvalidHeader("missing format field".into()))?;
        let mode_str = mode.ok_or_else(|| IrError::InvalidHeader("missing mode field".into()))?;
        let version_str =
            version.ok_or_else(|| IrError::InvalidHeader("missing version field".into()))?;

        let header = IrHeader {
            source: source.ok_or_else(|| IrError::InvalidHeader("missing source field".into()))?,
            format: Format::parse_str(&format_str)?,
            mode: Mode::parse_str(&mode_str)?,
            version: version_str.parse::<u32>().map_err(|_| {
                IrError::InvalidHeader(format!("invalid version number: {version_str}"))
            })?,
            checksum: checksum
                .ok_or_else(|| IrError::InvalidHeader("missing checksum field".into()))?,
        };

        Ok((header, body.to_string()))
    }
}

/// Skip the first line of a string (everything up to and including the first `\n`).
fn skip_line(s: &str) -> &str {
    match s.find('\n') {
        Some(i) => &s[i + 1..],
        None => "",
    }
}

#[cfg(test)]
mod tests {
    #![allow(clippy::panic_in_result_fn)]

    use super::*;

    type TestResult = std::result::Result<(), Box<dyn std::error::Error>>;

    #[test]
    fn header_roundtrip() -> TestResult {
        let header = IrHeader {
            source: "report.xlsx".to_string(),
            format: Format::Xlsx,
            mode: Mode::Content,
            version: 1,
            checksum: "sha256:abc123".to_string(),
        };

        let written = header.write();
        let (parsed, body) = IrHeader::parse(&written)?;

        assert_eq!(parsed.source, "report.xlsx");
        assert_eq!(parsed.format, Format::Xlsx);
        assert_eq!(parsed.mode, Mode::Content);
        assert_eq!(parsed.version, 1);
        assert_eq!(parsed.checksum, "sha256:abc123");
        assert!(body.is_empty());

        Ok(())
    }

    #[test]
    fn header_with_body() -> TestResult {
        let input = "+++\n\
                      source = \"test.xlsx\"\n\
                      format = \"xlsx\"\n\
                      mode = \"content\"\n\
                      version = 1\n\
                      checksum = \"sha256:def456\"\n\
                      +++\n\
                      === Sheet: Sheet1 ===\n\
                      A1: hello\n";

        let (header, body) = IrHeader::parse(input)?;
        assert_eq!(header.source, "test.xlsx");
        assert!(body.contains("=== Sheet: Sheet1 ==="));
        assert!(body.contains("A1: hello"));

        Ok(())
    }

    #[test]
    fn header_missing_delimiter() {
        let input = "source = \"test.xlsx\"\nformat = \"xlsx\"\n";
        let result = IrHeader::parse(input);
        assert!(result.is_err());
    }
}
