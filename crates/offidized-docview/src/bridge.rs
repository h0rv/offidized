//! `wasm_bindgen` bridge exposing [`DocView`] to JavaScript.

use wasm_bindgen::prelude::wasm_bindgen;
use wasm_bindgen::JsValue;

use crate::convert::convert_document;
use crate::model::DocViewModel;

/// WASM-exposed document viewer handle.
///
/// Parse a `.docx` file and retrieve its view model as a JS object.
#[wasm_bindgen]
pub struct DocView {
    model: Option<DocViewModel>,
}

impl Default for DocView {
    fn default() -> Self {
        Self::new()
    }
}

#[wasm_bindgen]
impl DocView {
    /// Create a new, empty `DocView` instance.
    #[wasm_bindgen(constructor)]
    pub fn new() -> Self {
        #[cfg(target_arch = "wasm32")]
        console_error_panic_hook::set_once();
        Self { model: None }
    }

    /// Parse `.docx` bytes and return the view model as a JS value.
    pub fn parse(&mut self, data: &[u8]) -> Result<JsValue, JsValue> {
        let doc = offidized_docx::Document::from_bytes(data)
            .map_err(|e| JsValue::from_str(&e.to_string()))?;
        let model = convert_document(&doc).map_err(|e| JsValue::from_str(&e.to_string()))?;
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
