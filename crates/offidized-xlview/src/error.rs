//! Error types for the Excel viewer.

use offidized_xlsx::XlsxError;

/// Errors that can occur in the Excel viewer.
#[derive(Debug, thiserror::Error)]
pub enum XlViewError {
    /// An error from the underlying xlsx library.
    #[error("xlsx error: {0}")]
    Xlsx(#[from] XlsxError),

    /// A JavaScript/WASM interop error.
    #[error("js error: {0}")]
    Js(String),

    /// Canvas rendering error.
    #[error("render error: {0}")]
    Render(String),

    /// The requested sheet index is out of range.
    #[error("sheet index {0} out of range")]
    SheetIndexOutOfRange(usize),

    /// A generic error with a descriptive message.
    #[error("{0}")]
    Other(String),

    /// An I/O error (e.g. from ZIP operations).
    #[error("io error: {0}")]
    Io(#[from] std::io::Error),

    /// A ZIP archive error.
    #[cfg(feature = "editing")]
    #[error("zip error: {0}")]
    Zip(#[from] zip::result::ZipError),
}

/// Result type alias for viewer operations.
pub type Result<T> = std::result::Result<T, XlViewError>;

impl From<wasm_bindgen::JsValue> for XlViewError {
    fn from(val: wasm_bindgen::JsValue) -> Self {
        Self::Js(format!("{val:?}"))
    }
}

impl From<&str> for XlViewError {
    fn from(s: &str) -> Self {
        Self::Render(s.to_string())
    }
}

impl From<js_sys::Object> for XlViewError {
    fn from(obj: js_sys::Object) -> Self {
        Self::Js(format!("{obj:?}"))
    }
}
