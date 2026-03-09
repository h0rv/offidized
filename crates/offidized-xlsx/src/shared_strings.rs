use crate::cell::RichTextRun;
use offidized_opc::RawXmlNode;

/// A single entry in the shared string table.
///
/// Each entry is either a plain text string or a rich text sequence of
/// formatted runs. Rich text entries preserve `<r><rPr>...</rPr><t>text</t></r>`
/// structures for roundtrip fidelity.
#[derive(Debug, Clone)]
pub enum SharedStringEntry {
    /// A plain text string (`<si><t>text</t></si>`).
    Plain(String),
    /// Rich text runs (`<si><r>...</r><r>...</r></si>`).
    RichText {
        /// The formatted runs.
        runs: Vec<RichTextRun>,
        /// Phonetic run elements (`<rPh>`) preserved for roundtrip.
        phonetic_runs: Vec<RawXmlNode>,
        /// Phonetic properties (`<phoneticPr>`) preserved for roundtrip.
        phonetic_pr: Option<RawXmlNode>,
    },
}

impl SharedStringEntry {
    /// Returns the plain text content of this entry (concatenated for rich text).
    pub fn plain_text(&self) -> String {
        match self {
            SharedStringEntry::Plain(s) => s.clone(),
            SharedStringEntry::RichText { runs, .. } => runs.iter().map(|r| r.text()).collect(),
        }
    }
}

/// Shared string table supporting both plain and rich text entries.
#[derive(Debug, Clone, Default)]
pub struct SharedStrings {
    entries: Vec<SharedStringEntry>,
}

impl SharedStrings {
    /// Creates a new empty shared string table.
    pub fn new() -> Self {
        Self::default()
    }

    /// Returns the number of entries.
    pub fn len(&self) -> usize {
        self.entries.len()
    }

    /// Returns true if the table has no entries.
    pub fn is_empty(&self) -> bool {
        self.entries.is_empty()
    }

    /// Intern a plain text string, returning its index.
    ///
    /// If an identical plain text string already exists, returns its index.
    /// Otherwise appends a new entry.
    pub fn intern(&mut self, value: impl Into<String>) -> usize {
        let value = value.into();

        if let Some(index) = self.entries.iter().position(|entry| match entry {
            SharedStringEntry::Plain(existing) => existing == &value,
            _ => false,
        }) {
            return index;
        }

        self.entries.push(SharedStringEntry::Plain(value));
        self.entries.len() - 1
    }

    /// Intern a rich text entry, returning its index.
    ///
    /// Rich text entries are always appended (no deduplication) because the
    /// formatting metadata makes equality checks complex and unreliable.
    pub fn intern_rich_text(
        &mut self,
        runs: Vec<RichTextRun>,
        phonetic_runs: Vec<RawXmlNode>,
        phonetic_pr: Option<RawXmlNode>,
    ) -> usize {
        self.entries.push(SharedStringEntry::RichText {
            runs,
            phonetic_runs,
            phonetic_pr,
        });
        self.entries.len() - 1
    }

    /// Append an entry without deduplication, preserving the original index.
    ///
    /// Use this when seeding from an existing shared string table to maintain
    /// index compatibility with passthrough (unchanged) sheets whose raw XML
    /// still references the original indices.
    pub fn push_raw(&mut self, value: impl Into<String>) -> usize {
        self.entries.push(SharedStringEntry::Plain(value.into()));
        self.entries.len() - 1
    }

    /// Append a rich text entry without deduplication, preserving the original index.
    pub fn push_raw_rich_text(
        &mut self,
        runs: Vec<RichTextRun>,
        phonetic_runs: Vec<RawXmlNode>,
        phonetic_pr: Option<RawXmlNode>,
    ) -> usize {
        self.entries.push(SharedStringEntry::RichText {
            runs,
            phonetic_runs,
            phonetic_pr,
        });
        self.entries.len() - 1
    }

    /// Returns the plain text value at the given index.
    pub fn get(&self, index: usize) -> Option<&str> {
        self.entries.get(index).map(|entry| match entry {
            SharedStringEntry::Plain(s) => s.as_str(),
            SharedStringEntry::RichText { runs, .. } => {
                // For plain text access, return the first run's text if available.
                // Callers needing the full concatenated text should use `get_entry`.
                if runs.len() == 1 {
                    runs[0].text()
                } else {
                    // Can't return a reference to a concatenated string, so
                    // return empty for multi-run entries. Callers should use
                    // `get_entry` instead.
                    ""
                }
            }
        })
    }

    /// Returns the entry at the given index.
    pub fn get_entry(&self, index: usize) -> Option<&SharedStringEntry> {
        self.entries.get(index)
    }
}
