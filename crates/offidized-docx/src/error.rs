use thiserror::Error;

/// Result type used by the docx crate.
pub type Result<T> = std::result::Result<T, DocxError>;

/// Errors emitted by high-level docx APIs.
#[derive(Debug, Error)]
pub enum DocxError {
    #[error("I/O error: {0}")]
    Io(#[from] std::io::Error),

    #[error("OPC package error: {0}")]
    Opc(#[from] offidized_opc::OpcError),

    #[error("XML parse/write error: {0}")]
    Xml(#[from] quick_xml::Error),

    #[error("unsupported Word package: {0}")]
    UnsupportedPackage(String),
}
