//! DOM input overlay for cell editing.
//!
//! Creates an `<input>` element positioned over the editing cell.
//! Keyboard handling (Enter/Escape/Tab) is done on the JS side via
//! the wrapper in `xl-edit.ts`.

use wasm_bindgen::JsCast;
use web_sys::{Document, HtmlElement, HtmlInputElement};

/// Input overlay for cell editing.
pub(crate) struct InputOverlay {
    input: Option<HtmlInputElement>,
}

impl InputOverlay {
    pub(crate) fn new() -> Self {
        InputOverlay { input: None }
    }

    /// Show the input overlay at the given rectangle.
    ///
    /// `rect` is `[x, y, w, h]` in logical (CSS) pixels relative to the viewport.
    /// `container` is the scroll container element to position relative to.
    pub(crate) fn show(
        &mut self,
        rect: &[f32],
        current_value: &str,
        container: Option<&HtmlElement>,
    ) {
        let Some(document) = web_sys::window().and_then(|w| w.document()) else {
            return;
        };

        let x = rect.first().copied().unwrap_or(0.0);
        let y = rect.get(1).copied().unwrap_or(0.0);
        let w = rect.get(2).copied().unwrap_or(100.0);
        let h = rect.get(3).copied().unwrap_or(24.0);

        let input = self.get_or_create_input(&document, container);
        let style = input.style();

        let _ = style.set_property("display", "block");
        let _ = style.set_property("left", &format!("{x}px"));
        let _ = style.set_property("top", &format!("{y}px"));
        let _ = style.set_property("width", &format!("{w}px"));
        let _ = style.set_property("height", &format!("{h}px"));

        input.set_value(current_value);

        // Focus and select all text
        let _ = input.focus();
        input.select();
    }

    /// Hide the input overlay.
    pub(crate) fn hide(&mut self) {
        if let Some(ref input) = self.input {
            let _ = input.style().set_property("display", "none");
            let _ = input.blur();
        }
    }

    /// Get current input value.
    pub(crate) fn value(&self) -> Option<String> {
        self.input.as_ref().map(|i| i.value())
    }

    /// Get or create the `<input>` element.
    fn get_or_create_input(
        &mut self,
        document: &Document,
        container: Option<&HtmlElement>,
    ) -> &HtmlInputElement {
        if self.input.is_none() {
            if let Ok(el) = document.create_element("input") {
                if let Ok(input) = el.dyn_into::<HtmlInputElement>() {
                    input.set_type("text");
                    let style = input.style();
                    let _ = style.set_property("position", "absolute");
                    let _ = style.set_property("z-index", "1000");
                    let _ = style.set_property("box-sizing", "border-box");
                    let _ = style.set_property("border", "2px solid #4285f4");
                    let _ = style.set_property("outline", "none");
                    let _ = style.set_property("padding", "0 4px");
                    let _ = style.set_property("font-family", "inherit");
                    let _ = style.set_property("font-size", "13px");
                    let _ = style.set_property("background", "#fff");
                    let _ = style.set_property("display", "none");

                    // Append to container or document body
                    if let Some(c) = container {
                        let _ = c.append_child(&input);
                    } else if let Some(body) = document.body() {
                        let _ = body.append_child(&input);
                    }

                    self.input = Some(input);
                }
            }
        }

        // Safe: we just created it above if it was None
        // If creation somehow failed, we'll get the previous one or this will be unreachable
        #[allow(clippy::expect_used)]
        self.input.as_ref().expect("input element must exist")
    }
}

impl Drop for InputOverlay {
    fn drop(&mut self) {
        if let Some(ref input) = self.input {
            if let Some(parent) = input.parent_node() {
                let _ = parent.remove_child(input);
            }
        }
    }
}
