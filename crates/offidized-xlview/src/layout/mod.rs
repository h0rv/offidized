//! Layout engine for computing cell positions and viewport management.
//!
//! This module handles:
//! - Pre-computing cell positions from column widths and row heights
//! - Managing viewport state (scroll position, visible range)
//! - Binary search for efficient cell lookup at screen coordinates
//! - Merge range handling

mod sheet_layout;
mod viewport;

pub use sheet_layout::{CellRect, MergeInfo, SheetLayout};
pub use viewport::Viewport;
