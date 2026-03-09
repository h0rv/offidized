use serde::{Deserialize, Serialize};

/// A single run of text with optional styling
#[derive(Debug, Serialize, Deserialize, Clone)]
#[serde(rename_all = "camelCase")]
pub struct RichTextRun {
    pub text: String,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub style: Option<RunStyle>,
}

/// Style properties for a rich text run
#[derive(Debug, Serialize, Deserialize, Clone, Default)]
#[serde(rename_all = "camelCase")]
pub struct RunStyle {
    #[serde(skip_serializing_if = "Option::is_none")]
    pub font_family: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub font_size: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub font_color: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub bold: Option<bool>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub italic: Option<bool>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub underline: Option<bool>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub strikethrough: Option<bool>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub vert_align: Option<VerticalAlign>,
}

/// A single run of text with optional styling for rendering.
#[derive(Debug, Clone)]
pub struct TextRunData {
    pub text: String,
    pub bold: Option<bool>,
    pub italic: Option<bool>,
    pub font_size: Option<f32>,
    pub font_color: Option<String>,
    pub font_family: Option<String>,
    pub underline: Option<bool>,
    pub strikethrough: Option<bool>,
}

/// Vertical alignment for subscript/superscript in rich text runs
#[derive(Debug, Serialize, Deserialize, Clone, Copy)]
#[serde(rename_all = "lowercase")]
pub enum VerticalAlign {
    Baseline,
    Subscript,
    Superscript,
}

/// Shared string entry - can be plain text or rich text
#[derive(Debug, Clone)]
pub enum SharedString {
    Plain(String),
    Rich(Vec<RichTextRun>),
}

impl SharedString {
    /// Get the plain text representation (concatenated for rich text)
    pub fn plain_text(&self) -> String {
        match self {
            SharedString::Plain(s) => s.clone(),
            SharedString::Rich(runs) => {
                let total_len: usize = runs.iter().map(|r| r.text.len()).sum();
                let mut combined = String::with_capacity(total_len);
                for run in runs {
                    combined.push_str(&run.text);
                }
                combined
            }
        }
    }

    /// Get the rich text runs, if any
    pub fn rich_runs(&self) -> Option<&Vec<RichTextRun>> {
        match self {
            SharedString::Plain(_) => None,
            SharedString::Rich(runs) => Some(runs),
        }
    }
}
