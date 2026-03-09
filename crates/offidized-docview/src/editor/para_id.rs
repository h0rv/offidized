//! App-level paragraph identifier.
//!
//! Uses UUID v7 (timestamp-sortable) to provide stable, session-independent
//! paragraph identifiers decoupled from CRDT internal IDs.

use uuid::Uuid;

/// App-level paragraph identifier.
///
/// Opaque, stable across sessions. Generated on import or creation.
/// Uses UUID v7 (timestamp-sortable).
#[derive(Clone, Debug, PartialEq, Eq, Hash)]
pub struct ParaId([u8; 16]);

impl ParaId {
    /// Generate a new unique paragraph ID (UUID v7).
    pub fn new() -> Self {
        Self(Uuid::now_v7().into_bytes())
    }

    /// Create from raw bytes.
    pub fn from_bytes(bytes: [u8; 16]) -> Self {
        Self(bytes)
    }

    /// Get the raw bytes.
    pub fn as_bytes(&self) -> &[u8; 16] {
        &self.0
    }
}

impl Default for ParaId {
    fn default() -> Self {
        Self::new()
    }
}

impl std::fmt::Display for ParaId {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        write!(f, "{}", Uuid::from_bytes(self.0))
    }
}

impl std::str::FromStr for ParaId {
    type Err = uuid::Error;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        let uuid = Uuid::parse_str(s)?;
        Ok(Self(uuid.into_bytes()))
    }
}
