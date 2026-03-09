use thiserror::Error;

pub type Result<T> = std::result::Result<T, PptxError>;

#[derive(Debug, Error)]
pub enum PptxError {
    #[error("I/O error: {0}")]
    Io(#[from] std::io::Error),

    #[error("OPC error: {0}")]
    Opc(#[from] offidized_opc::OpcError),

    #[error("XML error: {0}")]
    Xml(#[from] quick_xml::Error),

    #[error("unsupported presentation package: {0}")]
    UnsupportedPackage(String),

    #[error("invalid operation: {0}")]
    InvalidOperation(String),
}
