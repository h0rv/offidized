//! `wasm_bindgen` bridge exposing [`PptView`] to JavaScript.

use wasm_bindgen::prelude::wasm_bindgen;
use wasm_bindgen::JsValue;

use crate::convert::convert_presentation;
use crate::model::PresentationViewModel;

/// WASM-exposed presentation viewer handle.
///
/// Parse a `.pptx` file and retrieve its view model as a JS object.
#[wasm_bindgen]
pub struct PptView {
    model: Option<PresentationViewModel>,
}

impl Default for PptView {
    fn default() -> Self {
        Self::new()
    }
}

#[wasm_bindgen]
impl PptView {
    /// Create a new, empty `PptView` instance.
    #[wasm_bindgen(constructor)]
    pub fn new() -> Self {
        #[cfg(target_arch = "wasm32")]
        console_error_panic_hook::set_once();
        Self { model: None }
    }

    /// Parse `.pptx` bytes and return the view model as a JS value.
    pub fn parse(&mut self, data: &[u8]) -> Result<JsValue, JsValue> {
        let pres = offidized_pptx::Presentation::from_bytes(data)
            .map_err(|e| JsValue::from_str(&e.to_string()))?;
        let model = convert_presentation(&pres).map_err(|e| JsValue::from_str(&e.to_string()))?;
        let js =
            serde_wasm_bindgen::to_value(&model).map_err(|e| JsValue::from_str(&e.to_string()))?;
        self.model = Some(model);
        Ok(js)
    }

    /// Get the base64 data URI for an image by its index.
    #[wasm_bindgen(js_name = imageDataUri)]
    pub fn image_data_uri(&self, index: usize) -> Option<String> {
        self.model
            .as_ref()?
            .images
            .get(index)
            .map(|img| img.data_uri.clone())
    }
}
