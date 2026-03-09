//! Mouse, click, and keyboard event handlers for `XlView`.
//!
//! All methods here are `pub(crate)` helpers called from the wasm-exported
//! public API that lives in `mod.rs`.

#[cfg(target_arch = "wasm32")]
use js_sys::Function;
#[cfg(target_arch = "wasm32")]
use std::cell::RefCell;
#[cfg(target_arch = "wasm32")]
use std::rc::Rc;
#[cfg(target_arch = "wasm32")]
use web_sys::HtmlElement;

#[cfg(target_arch = "wasm32")]
use wasm_bindgen::prelude::*;

#[cfg(target_arch = "wasm32")]
use super::{col_to_letter, HeaderDragMode, HitTarget, SharedState, XlView, RESIZE_HANDLE_SIZE};
#[cfg(target_arch = "wasm32")]
use crate::types::selection::Selection;

#[cfg(target_arch = "wasm32")]
impl XlView {
    // Internal event handlers that work with shared state
    pub(crate) fn internal_mouse_down(state: &Rc<RefCell<SharedState>>, x: f32, y: f32) {
        let callback = (|| {
            let mut s = state.borrow_mut();

            // Use hit_test to determine what was clicked
            let hit = Self::hit_test(&s, x, y);

            match hit {
                HitTarget::None => {
                    return None;
                }
                HitTarget::CornerHeader => {
                    // Select all - extract values first to avoid borrow conflict
                    let (max_row, max_col) = {
                        let layout = s.layouts.get(s.active_sheet)?;
                        (layout.max_row, layout.max_col)
                    };
                    s.selection = Some(Selection::all());
                    s.selection_start = Some((0, 0));
                    s.selection_end = Some((max_row, max_col));
                    s.is_selecting = false;
                    s.is_resizing = false;
                    s.header_drag_mode = None;
                    s.needs_overlay_render = true;
                    return s.render_callback.clone();
                }
                HitTarget::ColumnHeader(col) => {
                    // Select entire column - extract values first
                    let max_row = s.layouts.get(s.active_sheet)?.max_row;
                    s.selection = Some(Selection::column_range(col, col));
                    s.selection_start = Some((0, col));
                    s.selection_end = Some((max_row, col));
                    s.is_selecting = false;
                    s.is_resizing = false;
                    s.header_drag_mode = Some(HeaderDragMode::Column);
                    s.needs_overlay_render = true;
                    return s.render_callback.clone();
                }
                HitTarget::RowHeader(row) => {
                    // Select entire row - extract values first
                    let max_col = s.layouts.get(s.active_sheet)?.max_col;
                    s.selection = Some(Selection::row_range(row, row));
                    s.selection_start = Some((row, 0));
                    s.selection_end = Some((row, max_col));
                    s.is_selecting = false;
                    s.is_resizing = false;
                    s.header_drag_mode = Some(HeaderDragMode::Row);
                    s.needs_overlay_render = true;
                    return s.render_callback.clone();
                }
                HitTarget::Cell(row, col) => {
                    let layout = s.layouts.get(s.active_sheet)?;

                    // Check if clicking on resize handle (if there's an existing selection)
                    if let (Some(start), Some(end)) = (s.selection_start, s.selection_end) {
                        let min_row = start.0.min(end.0);
                        let max_row = start.0.max(end.0);
                        let min_col = start.1.min(end.1);
                        let max_col = start.1.max(end.1);

                        // Get bottom-right corner of selection in screen coordinates
                        let sel_x2 = layout
                            .col_positions
                            .get((max_col + 1) as usize)
                            .copied()
                            .unwrap_or(0.0);
                        let sel_y2 = layout
                            .row_positions
                            .get((max_row + 1) as usize)
                            .copied()
                            .unwrap_or(0.0);

                        // Convert to screen coordinates (accounting for headers)
                        let header_offset_x = if s.show_headers {
                            s.header_config.row_header_width
                        } else {
                            0.0
                        };
                        let header_offset_y = if s.show_headers {
                            s.header_config.col_header_height
                        } else {
                            0.0
                        };
                        let handle_screen_x = sel_x2 - s.viewport.scroll_x + header_offset_x;
                        let handle_screen_y = sel_y2 - s.viewport.scroll_y + header_offset_y;

                        // Check if click is within resize handle area
                        let hit_area = RESIZE_HANDLE_SIZE + 8.0;
                        if (x - handle_screen_x).abs() <= hit_area
                            && (y - handle_screen_y).abs() <= hit_area
                        {
                            s.selection_start = Some((min_row, min_col));
                            s.is_resizing = true;
                            s.is_selecting = false;
                            s.header_drag_mode = None;
                            s.needs_overlay_render = true;
                            return s.render_callback.clone();
                        }
                    }

                    // Normal cell selection
                    s.selection = Some(Selection::cell_range(row, col, row, col));
                    s.selection_start = Some((row, col));
                    s.selection_end = Some((row, col));
                    s.is_selecting = true;
                    s.is_resizing = false;
                    s.header_drag_mode = None;
                    s.needs_overlay_render = true;
                    return s.render_callback.clone();
                }
            }
        })();
        Self::invoke_render_callback(callback);
    }

    pub(crate) fn internal_mouse_move(state: &Rc<RefCell<SharedState>>, x: f32, y: f32) {
        let callback = (|| {
            let mut s = state.borrow_mut();

            // Handle header drag selection (for multi-row/column selection)
            if let Some(drag_mode) = s.header_drag_mode {
                // Extract header config values before borrowing layout
                let header_height = if s.show_headers {
                    s.header_config.col_header_height
                } else {
                    0.0
                };
                let header_width = if s.show_headers {
                    s.header_config.row_header_width
                } else {
                    0.0
                };
                let scroll_x = s.viewport.scroll_x;
                let scroll_y = s.viewport.scroll_y;
                let selection_start = s.selection_start;
                let selection_end = s.selection_end;

                // Extract layout values
                let Some(layout) = s.layouts.get(s.active_sheet) else {
                    return None;
                };
                let frozen_height = layout.frozen_rows_height();
                let frozen_width = layout.frozen_cols_width();
                let max_row = layout.max_row;
                let max_col = layout.max_col;

                match drag_mode {
                    HeaderDragMode::Row => {
                        // Dragging across row headers
                        let content_y = y - header_height;

                        let row = if content_y < frozen_height {
                            layout.row_at_y(content_y)
                        } else {
                            let sheet_y = content_y - frozen_height + scroll_y;
                            layout.row_at_y(sheet_y)
                        };

                        if let Some(row) = row {
                            let start_row = selection_start.map(|(r, _)| r).unwrap_or(0);
                            let new_end = Some((row, max_col));
                            if selection_end != new_end {
                                s.selection_end = new_end;
                                s.selection = Some(Selection::row_range(start_row, row));
                                s.needs_overlay_render = true;
                                return s.render_callback.clone();
                            }
                        }
                    }
                    HeaderDragMode::Column => {
                        // Dragging across column headers
                        let content_x = x - header_width;

                        let col = if content_x < frozen_width {
                            layout.col_at_x(content_x)
                        } else {
                            let sheet_x = content_x - frozen_width + scroll_x;
                            layout.col_at_x(sheet_x)
                        };

                        if let Some(col) = col {
                            let start_col = selection_start.map(|(_, c)| c).unwrap_or(0);
                            let new_end = Some((max_row, col));
                            if selection_end != new_end {
                                s.selection_end = new_end;
                                s.selection = Some(Selection::column_range(start_col, col));
                                s.needs_overlay_render = true;
                                return s.render_callback.clone();
                            }
                        }
                    }
                }
                return None;
            }

            // Handle regular cell selection or resize
            if !s.is_selecting && !s.is_resizing {
                return None;
            }

            // Extract values from s before borrowing layout
            let header_width = if s.show_headers {
                s.header_config.row_header_width
            } else {
                0.0
            };
            let header_height = if s.show_headers {
                s.header_config.col_header_height
            } else {
                0.0
            };
            let scroll_x = s.viewport.scroll_x;
            let scroll_y = s.viewport.scroll_y;
            let viewport_height = s.viewport.height;
            let selection_start = s.selection_start;
            let selection_end = s.selection_end;

            let Some(layout) = s.layouts.get(s.active_sheet) else {
                return None;
            };

            let content_x = (x - header_width).max(0.0);
            let content_y = (y - header_height)
                .min(viewport_height - header_height)
                .max(0.0);

            let frozen_width = layout.frozen_cols_width();
            let frozen_height = layout.frozen_rows_height();

            let col = if content_x < frozen_width {
                layout.col_at_x(content_x)
            } else {
                let sheet_x = content_x - frozen_width + scroll_x;
                layout.col_at_x(sheet_x)
            };

            let row = if content_y < frozen_height {
                layout.row_at_y(content_y)
            } else {
                let sheet_y = content_y - frozen_height + scroll_y;
                layout.row_at_y(sheet_y)
            };

            if let (Some(col), Some(row)) = (col, row) {
                let new_end = Some((row, col));
                if selection_end != new_end {
                    s.selection_end = new_end;
                    // Update the Selection object for cell range
                    if let Some((start_row, start_col)) = selection_start {
                        s.selection = Some(Selection::cell_range(start_row, start_col, row, col));
                    }
                    s.needs_overlay_render = true;
                    return s.render_callback.clone();
                }
            }
            None
        })();
        Self::invoke_render_callback(callback);
    }

    pub(crate) fn internal_mouse_up(state: &Rc<RefCell<SharedState>>) {
        let mut s = state.borrow_mut();
        s.is_selecting = false;
        s.is_resizing = false;
        s.header_drag_mode = None;
    }

    /// Determine what was clicked at the given position
    pub(crate) fn hit_test(s: &SharedState, x: f32, y: f32) -> HitTarget {
        // Tab bar is now DOM-based, not part of canvas hit testing

        // If headers are not shown, just check cells
        if !s.show_headers {
            let Some(layout) = s.layouts.get(s.active_sheet) else {
                return HitTarget::None;
            };
            let sheet_x = x + s.viewport.scroll_x;
            let sheet_y = y + s.viewport.scroll_y;
            let col = layout.col_at_x(sheet_x);
            let row = layout.row_at_y(sheet_y);
            return match (col, row) {
                (Some(col), Some(row)) => HitTarget::Cell(row, col),
                _ => HitTarget::None,
            };
        }

        let header_width = s.header_config.row_header_width;
        let header_height = s.header_config.col_header_height;

        // Check for corner header (select all)
        if x < header_width && y < header_height {
            return HitTarget::CornerHeader;
        }

        // Check for column header
        if y < header_height {
            let Some(layout) = s.layouts.get(s.active_sheet) else {
                return HitTarget::None;
            };
            let content_x = x - header_width;
            let frozen_width = layout.frozen_cols_width();

            let col = if content_x < frozen_width {
                layout.col_at_x(content_x)
            } else {
                let sheet_x = content_x - frozen_width + s.viewport.scroll_x;
                layout.col_at_x(sheet_x)
            };

            return match col {
                Some(col) => HitTarget::ColumnHeader(col),
                None => HitTarget::None,
            };
        }

        // Check for row header
        if x < header_width {
            let Some(layout) = s.layouts.get(s.active_sheet) else {
                return HitTarget::None;
            };
            let content_y = y - header_height;
            let frozen_height = layout.frozen_rows_height();

            let row = if content_y < frozen_height {
                layout.row_at_y(content_y)
            } else {
                let sheet_y = content_y - frozen_height + s.viewport.scroll_y;
                layout.row_at_y(sheet_y)
            };

            return match row {
                Some(row) => HitTarget::RowHeader(row),
                None => HitTarget::None,
            };
        }

        // Otherwise it's a cell
        let Some(layout) = s.layouts.get(s.active_sheet) else {
            return HitTarget::None;
        };

        let content_x = x - header_width;
        let content_y = y - header_height;
        let frozen_width = layout.frozen_cols_width();
        let frozen_height = layout.frozen_rows_height();

        let col = if content_x < frozen_width {
            layout.col_at_x(content_x)
        } else {
            let sheet_x = content_x - frozen_width + s.viewport.scroll_x;
            layout.col_at_x(sheet_x)
        };

        let row = if content_y < frozen_height {
            layout.row_at_y(content_y)
        } else {
            let sheet_y = content_y - frozen_height + s.viewport.scroll_y;
            layout.row_at_y(sheet_y)
        };

        match (col, row) {
            (Some(col), Some(row)) => HitTarget::Cell(row, col),
            _ => HitTarget::None,
        }
    }

    pub(crate) fn invoke_render_callback(callback: Option<Function>) {
        if let Some(callback) = callback {
            let _ = callback.call0(&JsValue::NULL);
        }
    }

    pub(crate) fn internal_key_down(
        state: &Rc<RefCell<SharedState>>,
        key: &str,
        ctrl: bool,
    ) -> bool {
        if ctrl && (key == "c" || key == "C") {
            let mut s = state.borrow_mut();
            if let Some(text) = Self::get_selected_values_from_state(&mut s) {
                drop(s);
                Self::copy_to_clipboard_internal(&text);
                return true;
            }
        }
        false
    }

    pub(crate) fn update_hover_ui(
        state: &Rc<RefCell<SharedState>>,
        element: &HtmlElement,
        tooltip: Option<&HtmlElement>,
        x: f32,
        y: f32,
        client_x: i32,
        client_y: i32,
    ) {
        let (is_selecting, is_resizing) = {
            let s = state.borrow();
            (s.is_selecting, s.is_resizing)
        };
        if is_selecting || is_resizing {
            let _ = element.style().set_property("cursor", "default");
            if let Some(tooltip) = tooltip {
                let _ = tooltip.style().set_property("display", "none");
            }
            return;
        }

        let (over_link, comment) = {
            let s = state.borrow();
            Self::hover_info(&s, x, y)
        };

        let cursor = if over_link { "pointer" } else { "default" };
        let _ = element.style().set_property("cursor", cursor);

        if let Some(tooltip) = tooltip {
            if let Some(comment) = comment {
                tooltip.set_text_content(Some(&comment));
                let _ = tooltip.style().set_property("display", "block");
                let _ = tooltip
                    .style()
                    .set_property("left", &format!("{}px", client_x + 12));
                let _ = tooltip
                    .style()
                    .set_property("top", &format!("{}px", client_y + 12));
            } else {
                let _ = tooltip.style().set_property("display", "none");
            }
        }
    }

    pub(crate) fn hover_info(s: &SharedState, x: f32, y: f32) -> (bool, Option<String>) {
        let tab_y = s.viewport.height;

        // Only check cells area, not tab bar
        if y >= tab_y {
            return (false, None);
        }

        let Some(workbook) = &s.workbook else {
            return (false, None);
        };
        let Some(layout) = s.layouts.get(s.active_sheet) else {
            return (false, None);
        };
        let Some(sheet) = workbook.sheets.get(s.active_sheet) else {
            return (false, None);
        };

        // Convert screen coordinates to sheet coordinates, handling headers
        // and frozen panes the same way hit_test() does.
        let (col, row) = if s.show_headers {
            let header_width = s.header_config.row_header_width;
            let header_height = s.header_config.col_header_height;

            // Click is inside header area — no cell
            if x < header_width || y < header_height {
                return (false, None);
            }

            let content_x = x - header_width;
            let content_y = y - header_height;
            let frozen_width = layout.frozen_cols_width();
            let frozen_height = layout.frozen_rows_height();

            let col = if content_x < frozen_width {
                layout.col_at_x(content_x)
            } else {
                layout.col_at_x(content_x - frozen_width + s.viewport.scroll_x)
            };
            let row = if content_y < frozen_height {
                layout.row_at_y(content_y)
            } else {
                layout.row_at_y(content_y - frozen_height + s.viewport.scroll_y)
            };
            (col, row)
        } else {
            let sheet_x = x + s.viewport.scroll_x;
            let sheet_y = y + s.viewport.scroll_y;
            (layout.col_at_x(sheet_x), layout.row_at_y(sheet_y))
        };

        let Some(col) = col else {
            return (false, None);
        };
        let Some(row) = row else {
            return (false, None);
        };
        let Some(cell_idx) = sheet.cell_index_at(row, col) else {
            return (false, None);
        };
        let cell_data = &sheet.cells[cell_idx];
        let over_link = cell_data
            .cell
            .hyperlink
            .as_ref()
            .is_some_and(|h| h.is_external);

        let comment = if cell_data.cell.has_comment == Some(true) {
            let cell_ref = format!("{}{}", col_to_letter(col), row + 1);
            if let Some(comment_idx) = sheet.comments_by_cell.get(&cell_ref) {
                if let Some(comment) = sheet.comments.get(*comment_idx) {
                    Some(if let Some(author) = &comment.author {
                        format!("{}: {}", author, comment.text)
                    } else {
                        comment.text.clone()
                    })
                } else {
                    None
                }
            } else {
                None
            }
        } else {
            None
        };

        (over_link, comment)
    }

    pub(crate) fn internal_click(state: &Rc<RefCell<SharedState>>, x: f32, y: f32) {
        let (_, url, callback) = {
            let mut s = state.borrow_mut();
            Self::handle_click_state(&mut s, x, y)
        };
        if let Some(url) = url {
            Self::open_url(&url);
        }
        Self::invoke_render_callback(callback);
    }

    pub(crate) fn handle_click_state(
        s: &mut SharedState,
        x: f32,
        y: f32,
    ) -> (Option<usize>, Option<String>, Option<Function>) {
        let tab_y = s.viewport.height;
        let viewport_width = s.viewport.width;

        if y < tab_y {
            let url = Self::hyperlink_url_at(s, x, y);
            return (None, url, None);
        }

        let Some(workbook) = &s.workbook else {
            return (None, None, None);
        };

        // Calculate total tab width to check if scrolling is needed
        let mut total_tab_width = 8.0f32;
        for sheet in &workbook.sheets {
            total_tab_width += sheet.name.len() as f32 * 8.0 + 24.0 + 4.0;
        }

        let needs_scroll = total_tab_width > viewport_width;
        let button_width = if needs_scroll { 24.0 } else { 0.0 };

        // Handle scroll button clicks
        if needs_scroll {
            let max_scroll = (total_tab_width - viewport_width + button_width * 2.0).max(0.0);

            // Left scroll button
            if x < button_width && s.viewport.tab_scroll_x > 0.0 {
                s.viewport.tab_scroll_x = (s.viewport.tab_scroll_x - 60.0).max(0.0);
                s.needs_overlay_render = true;
                return (None, None, s.render_callback.clone());
            }

            // Right scroll button
            if x >= viewport_width - button_width && s.viewport.tab_scroll_x < max_scroll {
                s.viewport.tab_scroll_x = (s.viewport.tab_scroll_x + 60.0).min(max_scroll);
                s.needs_overlay_render = true;
                return (None, None, s.render_callback.clone());
            }
        }

        // Check tab clicks (with scroll offset applied)
        let tab_scroll = s.viewport.tab_scroll_x;
        let mut tab_x = 8.0f32 + button_width - tab_scroll;
        for (i, sheet) in workbook.sheets.iter().enumerate() {
            let tab_width = sheet.name.len() as f32 * 8.0 + 24.0;
            if x >= tab_x
                && x < tab_x + tab_width
                && x >= button_width
                && x < viewport_width - button_width
            {
                if Self::set_active_sheet_state(s, i) {
                    s.needs_render = true;
                    return (Some(i), None, s.render_callback.clone());
                }
                return (None, None, None);
            }
            tab_x += tab_width + 4.0;
        }
        (None, None, None)
    }

    pub(crate) fn hyperlink_url_at(s: &SharedState, x: f32, y: f32) -> Option<String> {
        let Some(workbook) = &s.workbook else {
            return None;
        };
        let Some(layout) = s.layouts.get(s.active_sheet) else {
            return None;
        };
        let Some(sheet) = workbook.sheets.get(s.active_sheet) else {
            return None;
        };
        let sheet_x = x + s.viewport.scroll_x;
        let sheet_y = y + s.viewport.scroll_y;
        let Some(col) = layout.col_at_x(sheet_x) else {
            return None;
        };
        let Some(row) = layout.row_at_y(sheet_y) else {
            return None;
        };
        let Some(cell_idx) = sheet.cell_index_at(row, col) else {
            return None;
        };
        let cell_data = &sheet.cells[cell_idx];
        let hyperlink = cell_data.cell.hyperlink.as_ref()?;
        if hyperlink.is_external {
            Some(hyperlink.target.clone())
        } else {
            None
        }
    }

    pub(crate) fn set_active_sheet_state(s: &mut SharedState, index: usize) -> bool {
        let Some(workbook) = &s.workbook else {
            return false;
        };
        if index >= workbook.sheets.len() {
            return false;
        }
        Self::invalidate_visible_cell_cache(s, false);
        s.active_sheet = index;
        let (frozen_x, frozen_y) = s
            .layouts
            .get(index)
            .map(|layout| (layout.frozen_cols_width(), layout.frozen_rows_height()))
            .unwrap_or((0.0, 0.0));
        s.viewport.scroll_x = frozen_x;
        s.viewport.scroll_y = frozen_y;
        Self::begin_idle_prewarm(s);
        s.selection_start = None;
        s.selection_end = None;
        true
    }

    pub(crate) fn open_url(url: &str) {
        if let Some(window) = web_sys::window() {
            let _ = window.open_with_url_and_target(url, "_blank");
        }
    }
}
