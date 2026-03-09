use thiserror::Error;

#[derive(Error, Debug)]
pub enum OpcError {
    #[error("ZIP error: {0}")]
    Zip(#[from] zip::result::ZipError),

    #[error("XML parse error: {0}")]
    Xml(#[from] quick_xml::Error),

    #[error("IO error: {0}")]
    Io(#[from] std::io::Error),

    #[error("Part not found: {0}")]
    PartNotFound(String),

    #[error("Invalid relationship: {0}")]
    InvalidRelationship(String),

    #[error("Invalid content type: {0}")]
    InvalidContentType(String),

    #[error("Invalid part URI: {0}")]
    InvalidUri(String),

    #[error("Malformed package: {0}")]
    MalformedPackage(String),
}

pub type Result<T> = std::result::Result<T, OpcError>;
