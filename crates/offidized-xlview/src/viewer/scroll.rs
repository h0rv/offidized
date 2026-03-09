//! Scroll-related logic for `XlView`.
//!
//! Includes viewport scroll state management and scroll settle timeout scheduling.

#[cfg(target_arch = "wasm32")]
use std::cell::RefCell;
#[cfg(target_arch = "wasm32")]
use std::rc::Rc;
#[cfg(target_arch = "wasm32")]
use wasm_bindgen::closure::Closure;
#[cfg(target_arch = "wasm32")]
use wasm_bindgen::prelude::*;

#[cfg(target_arch = "wasm32")]
use super::{now_ms, SharedState, XlView};

/// Delay (ms) after scroll stops before triggering a settle render.
#[cfg(target_arch = "wasm32")]
const SCROLL_SETTLE_DELAY_MS: u32 = 100;

#[cfg(target_arch = "wasm32")]
impl XlView {
    pub(crate) fn scroll_state(
        s: &mut SharedState,
        delta_x: f32,
        delta_y: f32,
    ) -> Option<(f32, f32)> {
        let active = s.active_sheet;
        let layout_bounds = s.layouts.get(active).map(|l| {
            (
                l.total_width(),
                l.total_height(),
                l.frozen_cols_width(),
                l.frozen_rows_height(),
            )
        });
        if let Some((total_w, total_h, frozen_w, frozen_h)) = layout_bounds {
            // Minimum scroll is at the frozen boundary
            let min_x = frozen_w;
            let min_y = frozen_h;
            // Maximum scroll allows viewing the end of the content
            let scrollable_w = total_w - frozen_w;
            let scrollable_h = total_h - frozen_h;
            let viewport_content_w = s.viewport.width - frozen_w;
            let viewport_content_h = s.viewport.height - frozen_h;
            let max_x = frozen_w + (scrollable_w - viewport_content_w).max(0.0);
            let max_y = frozen_h + (scrollable_h - viewport_content_h).max(0.0);

            let new_x = (s.viewport.scroll_x + delta_x).clamp(min_x, max_x);
            let new_y = (s.viewport.scroll_y + delta_y).clamp(min_y, max_y);
            let dx = new_x - s.viewport.scroll_x;
            let dy = new_y - s.viewport.scroll_y;
            if dx.abs() > f32::EPSILON || dy.abs() > f32::EPSILON {
                s.viewport.scroll_x = new_x;
                s.viewport.scroll_y = new_y;
                return Some((dx, dy));
            }
        }
        None
    }

    pub(crate) fn schedule_scroll_settle_timeout(state: &Rc<RefCell<SharedState>>) {
        let Some(window) = web_sys::window() else {
            return;
        };
        let mut s = state.borrow_mut();
        // Cancel any existing timer
        if let Some(timer_id) = s.scroll_settle_timer.take() {
            window.clear_timeout_with_handle(timer_id);
        }
        if s.scroll_settle_closure.is_none() {
            let weak_state = Rc::downgrade(state);
            let closure = Closure::wrap(Box::new(move || {
                if let Some(state) = weak_state.upgrade() {
                    XlView::handle_scroll_settle(&state);
                }
            }) as Box<dyn FnMut()>);
            s.scroll_settle_closure = Some(closure);
        }
        let Some(callback) = s.scroll_settle_closure.as_ref() else {
            return;
        };
        match window.set_timeout_with_callback_and_timeout_and_arguments_0(
            callback.as_ref().unchecked_ref(),
            SCROLL_SETTLE_DELAY_MS as i32,
        ) {
            Ok(id) => s.scroll_settle_timer = Some(id),
            Err(_) => s.scroll_settle_timer = None,
        }
    }

    pub(crate) fn handle_scroll_settle(state: &Rc<RefCell<SharedState>>) {
        let callback = {
            let mut s = state.borrow_mut();
            s.scroll_settle_timer = None;
            // Check if scroll is still ongoing
            let elapsed = now_ms() - s.last_scroll_ms;
            if elapsed < f64::from(SCROLL_SETTLE_DELAY_MS) {
                // Still scrolling, reschedule
                drop(s);
                Self::schedule_scroll_settle_timeout(state);
                return;
            }
            s.needs_render = true;
            s.needs_overlay_render = true;
            s.render_callback.clone()
        };
        Self::invoke_render_callback(callback);
    }
}
