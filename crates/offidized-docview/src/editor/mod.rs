//! Collaborative editing module (requires `editing` feature).
//!
//! Provides CRDT-backed document editing via yrs (Y-CRDT).

pub mod bridge;
pub mod crdt_doc;
pub mod export;
pub mod import;
pub mod intent;
pub mod para_id;
pub mod tokens;
pub mod view;
