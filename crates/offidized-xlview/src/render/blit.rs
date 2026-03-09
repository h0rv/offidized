//! Layout helpers for scrollable region calculations.

use crate::layout::{SheetLayout, Viewport};

#[derive(Debug, Clone, Copy, PartialEq)]
pub struct ScrollableRegion {
    pub local_x: f64,
    pub local_y: f64,
    pub width: f64,
    pub height: f64,
    pub abs_x: f64,
    pub abs_y: f64,
    pub content_width: f64,
    pub content_height: f64,
}

pub fn scrollable_region(
    layout: &SheetLayout,
    viewport: &Viewport,
    header_offset_x: f64,
    header_offset_y: f64,
    scrollbar_size: f64,
) -> ScrollableRegion {
    let content_width = (f64::from(viewport.width) - scrollbar_size - header_offset_x).max(0.0);
    let content_height = (f64::from(viewport.height) - scrollbar_size - header_offset_y).max(0.0);
    let frozen_width = f64::from(layout.frozen_cols_width());
    let frozen_height = f64::from(layout.frozen_rows_height());
    let scrollable_width = (content_width - frozen_width).max(0.0);
    let scrollable_height = (content_height - frozen_height).max(0.0);
    let local_x = frozen_width;
    let local_y = frozen_height;
    ScrollableRegion {
        local_x,
        local_y,
        width: scrollable_width,
        height: scrollable_height,
        abs_x: local_x + header_offset_x,
        abs_y: local_y + header_offset_y,
        content_width,
        content_height,
    }
}

#[cfg(test)]
#[allow(
    clippy::unwrap_used,
    clippy::expect_used,
    clippy::indexing_slicing,
    clippy::float_cmp,
    clippy::panic
)]
mod tests {
    use super::*;
    use crate::layout::SheetLayout;

    fn layout_with_frozen(rows: u32, cols: u32) -> SheetLayout {
        let col_widths = std::collections::HashMap::new();
        let row_heights = std::collections::HashMap::new();
        let hidden_cols = std::collections::HashSet::new();
        let hidden_rows = std::collections::HashSet::new();
        SheetLayout::new(
            10,
            10,
            &col_widths,
            &row_heights,
            &hidden_cols,
            &hidden_rows,
            &[],
            rows,
            cols,
        )
    }

    #[test]
    fn scrollable_region_includes_header_offsets() {
        let layout = layout_with_frozen(1, 2);
        let mut viewport = Viewport::new();
        viewport.width = 800.0;
        viewport.height = 600.0;

        let header_offset_x = 40.0;
        let header_offset_y = 20.0;
        let scrollbar_size = 14.0;
        let region = scrollable_region(
            &layout,
            &viewport,
            header_offset_x,
            header_offset_y,
            scrollbar_size,
        );

        let frozen_width = f64::from(layout.frozen_cols_width());
        let frozen_height = f64::from(layout.frozen_rows_height());
        let content_width = (f64::from(viewport.width) - scrollbar_size - header_offset_x).max(0.0);
        let content_height =
            (f64::from(viewport.height) - scrollbar_size - header_offset_y).max(0.0);

        assert_eq!(region.content_width, content_width);
        assert_eq!(region.content_height, content_height);
        assert_eq!(region.local_x, frozen_width);
        assert_eq!(region.local_y, frozen_height);
        assert_eq!(region.width, (content_width - frozen_width).max(0.0));
        assert_eq!(region.height, (content_height - frozen_height).max(0.0));
        assert_eq!(region.abs_x, frozen_width + header_offset_x);
        assert_eq!(region.abs_y, frozen_height + header_offset_y);
    }
}
