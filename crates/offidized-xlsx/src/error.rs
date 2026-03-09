use thiserror::Error;

/// Errors returned by `offidized-xlsx`.
#[derive(Debug, Error)]
pub enum XlsxError {
    #[error("I/O error: {0}")]
    Io(#[from] std::io::Error),

    #[error("OPC error: {0}")]
    Opc(#[from] offidized_opc::OpcError),

    #[error("UTF-8 decode error: {0}")]
    Utf8(#[from] std::str::Utf8Error),

    #[error("XML parse/write error: {0}")]
    Xml(#[from] quick_xml::Error),

    #[error("XML serialization error: {0}")]
    XmlSerialize(String),

    #[error("XML deserialization error: {0}")]
    XmlDeserialize(String),

    #[error("invalid cell reference: {0}")]
    InvalidCellReference(String),

    #[error("invalid formula: {0}")]
    InvalidFormula(String),

    #[error("unsupported workbook package: {0}")]
    UnsupportedPackage(String),

    #[error("invalid workbook state: {0}")]
    InvalidWorkbookState(String),
}

pub type Result<T> = std::result::Result<T, XlsxError>;
