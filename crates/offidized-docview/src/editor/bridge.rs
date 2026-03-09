//! WASM bridge for collaborative document editing.
//!
//! [`DocEdit`] wraps the CRDT document, original `.docx`, and undo manager.
//! TypeScript sends editing intents; WASM applies them as CRDT transactions
//! and returns view model updates.
//!
//! # Borrow safety
//!
//! All `#[wasm_bindgen]` methods take `&self` (shared borrow) and use
//! `Cell<Option<CrdtDoc>>` with a take/put pattern instead of `RefCell`.
//! This avoids borrow-counter poisoning: on wasm32, `wasm_bindgen`'s
//! `maybe_catch_unwind` can catch panics without running destructors,
//! permanently poisoning both `WasmRefCell` and `RefCell`. With `Cell`,
//! a caught panic simply leaves `None` — subsequent calls return a
//! clean error instead of a borrow panic.

use std::cell::Cell;
use std::collections::HashMap;
use std::sync::Arc;

use base64::Engine;
use uuid::Uuid;
use wasm_bindgen::prelude::wasm_bindgen;
use wasm_bindgen::JsValue;
use yrs::updates::decoder::Decode;
use yrs::updates::encoder::Encode;
use yrs::{
    Any, Array, GetString, Map, MapPrelim, Out, ReadTxn, StateVector, Text, TextPrelim, Transact,
    Update,
};

use super::crdt_doc::CrdtDoc;
use super::intent::{EditIntent, IntentAttrValue};
use super::para_id::ParaId;
use super::{export, import, tokens, view};

type TextFormattingMap = HashMap<Arc<str>, Any>;
type ParagraphFormattingMap = HashMap<String, Any>;
type DecodedDataUri = (String, Vec<u8>);

// ---------------------------------------------------------------------------
// Error types
// ---------------------------------------------------------------------------

/// Errors during intent processing.
#[derive(Debug, thiserror::Error)]
pub enum EditError {
    /// General editing error.
    #[error("edit error: {0}")]
    General(String),
    /// Position decoding failed.
    #[error("position error: {0}")]
    Position(String),
}

// ---------------------------------------------------------------------------
// Position encoding/decoding
// ---------------------------------------------------------------------------

/// Decode a position string into (body_array_index, char_offset).
///
/// For MVP, positions are base64-encoded `[body_index_u32_le, offset_u32_le]`.
/// Full CRDT-relative positions (StickyIndex) are future work.
fn decode_position(pos: &str) -> Result<(u32, u32), String> {
    use base64::Engine;
    let engine = base64::engine::general_purpose::STANDARD;
    let bytes = engine
        .decode(pos)
        .map_err(|e| format!("base64 decode: {e}"))?;
    if bytes.len() < 8 {
        return Err("position too short".to_string());
    }
    let body_index = u32::from_le_bytes([bytes[0], bytes[1], bytes[2], bytes[3]]);
    let char_offset = u32::from_le_bytes([bytes[4], bytes[5], bytes[6], bytes[7]]);
    Ok((body_index, char_offset))
}

/// Encode a position as a base64 string.
fn encode_position(body_index: u32, char_offset: u32) -> String {
    use base64::Engine;
    let engine = base64::engine::general_purpose::STANDARD;
    let mut bytes = Vec::with_capacity(8);
    bytes.extend_from_slice(&body_index.to_le_bytes());
    bytes.extend_from_slice(&char_offset.to_le_bytes());
    engine.encode(&bytes)
}

// ---------------------------------------------------------------------------
// DocEdit -- WASM-exported struct
// ---------------------------------------------------------------------------

/// WASM-exported document editor handle.
///
/// Wraps a CRDT document (the editing source of truth), the original
/// parsed `.docx` (for export and sections/images), and an undo manager.
///
/// Uses `Cell<Option<CrdtDoc>>` with a take/put pattern so that all
/// `#[wasm_bindgen]` methods can take `&self` without any borrow
/// tracking. This is critical because on wasm32, `wasm_bindgen`'s
/// `maybe_catch_unwind` can catch panics without running destructors,
/// permanently poisoning both `WasmRefCell` and `RefCell` borrow
/// counters. `Cell` has no borrow counter — if a panic is caught,
/// the `Cell` simply contains `None` and subsequent calls get a
/// clean error instead of a borrow panic.
#[wasm_bindgen]
pub struct DocEdit {
    crdt: Cell<Option<CrdtDoc>>,
    original_doc: offidized_docx::Document,
    #[allow(dead_code)]
    original_bytes: Vec<u8>,
    undo_mgr: Cell<Option<yrs::undo::UndoManager<()>>>,
}

#[wasm_bindgen]
impl DocEdit {
    /// Create a new `DocEdit` by loading `.docx` bytes.
    ///
    /// Parses the document, imports it into the CRDT, and prepares
    /// the editing session.
    #[wasm_bindgen(constructor)]
    pub fn new(data: &[u8]) -> Result<DocEdit, JsValue> {
        #[cfg(target_arch = "wasm32")]
        console_error_panic_hook::set_once();

        let doc = offidized_docx::Document::from_bytes(data)
            .map_err(|e| JsValue::from_str(&e.to_string()))?;
        let mut crdt = CrdtDoc::new();
        import::import_document(&doc, &mut crdt).map_err(|e| JsValue::from_str(&e.to_string()))?;

        let undo_mgr = Self::create_undo_manager(&crdt);

        Ok(DocEdit {
            crdt: Cell::new(Some(crdt)),
            original_doc: doc,
            original_bytes: data.to_vec(),
            undo_mgr: Cell::new(Some(undo_mgr)),
        })
    }

    /// Create a new blank document for editing from scratch.
    #[wasm_bindgen]
    pub fn blank() -> Result<DocEdit, JsValue> {
        #[cfg(target_arch = "wasm32")]
        console_error_panic_hook::set_once();

        let mut doc = offidized_docx::Document::new();
        doc.add_paragraph("");
        let bytes = doc
            .to_bytes()
            .map_err(|e| JsValue::from_str(&e.to_string()))?;
        let doc2 = offidized_docx::Document::from_bytes(&bytes)
            .map_err(|e| JsValue::from_str(&e.to_string()))?;
        let mut crdt = CrdtDoc::new();
        import::import_document(&doc2, &mut crdt).map_err(|e| JsValue::from_str(&e.to_string()))?;

        let undo_mgr = Self::create_undo_manager(&crdt);

        Ok(DocEdit {
            crdt: Cell::new(Some(crdt)),
            original_doc: doc2,
            original_bytes: bytes,
            undo_mgr: Cell::new(Some(undo_mgr)),
        })
    }

    /// Get the current view model as a JS value.
    ///
    /// This performs a full CRDT-to-view-model conversion.
    /// Call once after load, then use `apply_intent` for incremental updates.
    #[wasm_bindgen(js_name = viewModel)]
    pub fn view_model(&self) -> Result<JsValue, JsValue> {
        let crdt = self
            .crdt
            .take()
            .ok_or_else(|| JsValue::from_str("editor state lost after previous error"))?;
        let result = view::crdt_to_view_model(&crdt, &self.original_doc);
        self.crdt.set(Some(crdt));
        let model = result.map_err(|e| JsValue::from_str(&e.to_string()))?;
        serde_wasm_bindgen::to_value(&model).map_err(|e| JsValue::from_str(&e.to_string()))
    }

    /// Apply an editing intent.
    ///
    /// The intent is a JSON string matching the `EditIntent` enum.
    /// Call `viewModel()` separately after this returns to get the
    /// updated view model.
    #[wasm_bindgen(js_name = applyIntent)]
    pub fn apply_intent(&self, intent_json: &str) -> Result<(), JsValue> {
        let intent: EditIntent = serde_json::from_str(intent_json)
            .map_err(|e| JsValue::from_str(&format!("invalid intent JSON: {e}")))?;

        // Undo/Redo are handled here (not in process_intent_on) because
        // the UndoManager needs its own transaction, separate from the
        // CRDT processing transaction.
        match &intent {
            EditIntent::Undo => {
                if let Some(mut mgr) = self.undo_mgr.take() {
                    let _ = mgr.try_undo();
                    self.undo_mgr.set(Some(mgr));
                }
                // Mark all paragraphs dirty (undo may affect any paragraph).
                if let Some(mut crdt) = self.crdt.take() {
                    Self::mark_all_dirty(&mut crdt);
                    self.crdt.set(Some(crdt));
                }
                return Ok(());
            }
            EditIntent::Redo => {
                if let Some(mut mgr) = self.undo_mgr.take() {
                    let _ = mgr.try_redo();
                    self.undo_mgr.set(Some(mgr));
                }
                if let Some(mut crdt) = self.crdt.take() {
                    Self::mark_all_dirty(&mut crdt);
                    self.crdt.set(Some(crdt));
                }
                return Ok(());
            }
            _ => {}
        }

        if Self::should_reset_history_before(&intent) {
            self.reset_history_capture();
        }

        let mut crdt = self
            .crdt
            .take()
            .ok_or_else(|| JsValue::from_str("editor state lost after previous error"))?;
        let result = Self::process_intent_on(&mut crdt, &intent);

        // After paragraph-creating intents, expand the undo manager's
        // scope to include any newly created TextRefs.
        if matches!(
            intent,
            EditIntent::InsertParagraph { .. }
                | EditIntent::InsertTable { .. }
                | EditIntent::InsertTableRow { .. }
                | EditIntent::InsertTableColumn { .. }
                | EditIntent::InsertInlineImage { .. }
        ) {
            self.expand_undo_scope(&crdt);
        }

        self.crdt.set(Some(crdt));
        result.map_err(|e| JsValue::from_str(&e.to_string()))?;
        if Self::should_reset_history_after(&intent) {
            self.reset_history_capture();
        }
        Ok(())
    }

    /// Export the current document state as `.docx` bytes.
    #[wasm_bindgen]
    pub fn save(&self) -> Result<Vec<u8>, JsValue> {
        let mut crdt = self
            .crdt
            .take()
            .ok_or_else(|| JsValue::from_str("editor state lost after previous error"))?;
        let result = export::export_to_docx(&crdt, &self.original_doc);
        if result.is_ok() {
            crdt.clear_dirty();
        }
        self.crdt.set(Some(crdt));
        result.map_err(|e| JsValue::from_str(&e.to_string()))
    }

    /// Check if any paragraphs have been modified.
    #[wasm_bindgen(js_name = isDirty)]
    pub fn is_dirty(&self) -> bool {
        let crdt = match self.crdt.take() {
            Some(c) => c,
            None => return false,
        };
        let dirty = !crdt.dirty_paragraphs().is_empty();
        self.crdt.set(Some(crdt));
        dirty
    }

    /// Get the base64 data URI for an image by its index.
    #[wasm_bindgen(js_name = imageDataUri)]
    pub fn image_data_uri(&self, index: usize) -> Option<String> {
        let img = self.original_doc.images().get(index)?;
        let engine = base64::engine::general_purpose::STANDARD;
        let encoded = engine.encode(img.bytes());
        Some(format!("data:{};base64,{encoded}", img.content_type()))
    }

    /// Insert an inline image token at the current selection, replacing any
    /// selected range.
    #[wasm_bindgen(js_name = insertInlineImage)]
    pub fn insert_inline_image(
        &self,
        anchor: &str,
        focus: &str,
        data_uri: &str,
        width_pt: f64,
        height_pt: f64,
        name: Option<String>,
        description: Option<String>,
    ) -> Result<(), JsValue> {
        let intent = EditIntent::InsertInlineImage {
            anchor: anchor.to_string(),
            focus: focus.to_string(),
            data_uri: data_uri.to_string(),
            width_pt,
            height_pt,
            name,
            description,
        };
        let intent_json = serde_json::to_string(&intent)
            .map_err(|e| JsValue::from_str(&format!("serialize image intent: {e}")))?;
        self.apply_intent(&intent_json)
    }

    /// Query the formatting attributes at a given position.
    ///
    /// Returns a JSON object like `{"bold":true,"italic":true}` with the
    /// formatting attrs of the character to the left of the given offset.
    /// Used by TypeScript to highlight active format buttons in the toolbar.
    #[wasm_bindgen(js_name = formattingAt)]
    pub fn formatting_at(&self, position: &str) -> JsValue {
        let crdt = match self.crdt.take() {
            Some(c) => c,
            None => return JsValue::NULL,
        };
        let result = (|| -> Result<(TextFormattingMap, ParagraphFormattingMap), String> {
            let (body_idx, offset) = decode_position(position)?;
            let text_ref =
                Self::paragraph_text_ref_from(&crdt, body_idx).map_err(|e| e.to_string())?;
            let txn = crdt.doc().transact();
            let attrs = Self::query_formatting_at(&text_ref, &txn, offset);
            let paragraph_attrs = crdt
                .body()
                .get(&txn, body_idx)
                .and_then(|entry| entry.cast::<yrs::MapRef>().ok())
                .map(|map_ref| Self::collect_paragraph_formatting(&map_ref, &txn))
                .unwrap_or_default();
            Ok((attrs, paragraph_attrs))
        })();
        self.crdt.set(Some(crdt));

        match result {
            Ok((attrs, paragraph_attrs)) => {
                let obj = js_sys::Object::new();
                for (k, v) in &attrs {
                    match v {
                        Any::Bool(true) => {
                            let _ =
                                js_sys::Reflect::set(&obj, &JsValue::from_str(k), &JsValue::TRUE);
                        }
                        Any::String(s) => {
                            let _ = js_sys::Reflect::set(
                                &obj,
                                &JsValue::from_str(k),
                                &JsValue::from_str(s),
                            );
                        }
                        Any::Number(n) => {
                            let (out_key, out_value) = match k.as_ref() {
                                "fontSize" => ("fontSizePt", JsValue::from_f64(*n / 2.0)),
                                "width" => (
                                    "widthPt",
                                    JsValue::from_f64(crate::units::emu_to_pt(*n as u32)),
                                ),
                                "height" => (
                                    "heightPt",
                                    JsValue::from_f64(crate::units::emu_to_pt(*n as u32)),
                                ),
                                _ => (k.as_ref(), JsValue::from_f64(*n)),
                            };
                            let _ =
                                js_sys::Reflect::set(&obj, &JsValue::from_str(out_key), &out_value);
                        }
                        _ => {}
                    }
                }
                for (key, value) in paragraph_attrs {
                    match value {
                        Any::Bool(flag) => {
                            let _ = js_sys::Reflect::set(
                                &obj,
                                &JsValue::from_str(&key),
                                &JsValue::from_bool(flag),
                            );
                        }
                        Any::String(text) => {
                            let _ = js_sys::Reflect::set(
                                &obj,
                                &JsValue::from_str(&key),
                                &JsValue::from_str(text.as_ref()),
                            );
                        }
                        Any::Number(n) => {
                            let _ = js_sys::Reflect::set(
                                &obj,
                                &JsValue::from_str(&key),
                                &JsValue::from_f64(n),
                            );
                        }
                        _ => {}
                    }
                }
                obj.into()
            }
            Err(_) => JsValue::NULL,
        }
    }

    /// Encode a position as base64 for TypeScript to use.
    ///
    /// TypeScript calls this with the paragraph's body index and
    /// character offset to get an opaque position string for intents.
    #[wasm_bindgen(js_name = encodePosition)]
    pub fn encode_position_js(body_index: u32, char_offset: u32) -> String {
        encode_position(body_index, char_offset)
    }

    /// Get the number of paragraphs in the document body.
    #[wasm_bindgen(js_name = bodyLength)]
    pub fn body_length(&self) -> u32 {
        let crdt = match self.crdt.take() {
            Some(c) => c,
            None => return 0,
        };
        let txn = crdt.doc().transact();
        let len = crdt.body().len(&txn);
        drop(txn);
        self.crdt.set(Some(crdt));
        len
    }

    // -- CRDT sync surface --------------------------------------------------

    /// Return the document's state vector (for sync handshake).
    ///
    /// The returned bytes can be sent to a remote peer, which then calls
    /// [`Self::encode_diff`] to compute a minimal update.
    #[wasm_bindgen(js_name = encodeStateVector)]
    pub fn encode_state_vector(&self) -> Result<Vec<u8>, JsValue> {
        let crdt = self
            .crdt
            .take()
            .ok_or_else(|| JsValue::from_str("editor state lost after previous error"))?;
        let txn = crdt.doc().transact();
        let sv = txn.state_vector().encode_v1();
        drop(txn);
        self.crdt.set(Some(crdt));
        Ok(sv)
    }

    /// Encode the full document state as a binary update.
    ///
    /// Equivalent to encoding a diff against an empty state vector --
    /// the result contains everything needed to reconstruct the document.
    #[wasm_bindgen(js_name = encodeStateAsUpdate)]
    pub fn encode_state_as_update(&self) -> Result<Vec<u8>, JsValue> {
        let crdt = self
            .crdt
            .take()
            .ok_or_else(|| JsValue::from_str("editor state lost after previous error"))?;
        let txn = crdt.doc().transact();
        let update = txn.encode_state_as_update_v1(&StateVector::default());
        drop(txn);
        self.crdt.set(Some(crdt));
        Ok(update)
    }

    /// Encode only the changes since a remote state vector.
    ///
    /// `remote_sv` is a binary state vector obtained from the remote peer's
    /// [`Self::encode_state_vector`].  The returned update contains only
    /// the operations the remote is missing.
    #[wasm_bindgen(js_name = encodeDiff)]
    pub fn encode_diff(&self, remote_sv: &[u8]) -> Result<Vec<u8>, JsValue> {
        let decoded_sv = StateVector::decode_v1(remote_sv)
            .map_err(|e| JsValue::from_str(&format!("failed to decode state vector: {e}")))?;
        let crdt = self
            .crdt
            .take()
            .ok_or_else(|| JsValue::from_str("editor state lost after previous error"))?;
        let txn = crdt.doc().transact();
        let update = txn.encode_state_as_update_v1(&decoded_sv);
        drop(txn);
        self.crdt.set(Some(crdt));
        Ok(update)
    }

    /// Apply a remote update from another client.
    ///
    /// After applying, all paragraphs are marked dirty (the update may
    /// have touched any paragraph) and the undo scope is rebuilt to
    /// cover any newly created paragraphs.
    ///
    /// Remaining limit: with the current `yrs::UndoManager` integration,
    /// any real remote mutation still acts as a history boundary and resets
    /// local undo/redo state. We only preserve local history for redundant
    /// no-op updates that leave the CRDT state vector unchanged.
    #[wasm_bindgen(js_name = applyUpdate)]
    pub fn apply_update(&self, update: &[u8]) -> Result<(), JsValue> {
        let decoded = Update::decode_v1(update)
            .map_err(|e| JsValue::from_str(&format!("failed to decode update: {e}")))?;
        let mut crdt = self
            .crdt
            .take()
            .ok_or_else(|| JsValue::from_str("editor state lost after previous error"))?;
        let before_sv = {
            let txn = crdt.doc().transact();
            txn.state_vector().encode_v1()
        };
        {
            let mut txn = Self::transact_mut_remote(&crdt);
            txn.apply_update(decoded)
                .map_err(|e| JsValue::from_str(&format!("failed to apply update: {e}")))?;
        }
        let after_sv = {
            let txn = crdt.doc().transact();
            txn.state_vector().encode_v1()
        };
        if before_sv == after_sv {
            self.crdt.set(Some(crdt));
            return Ok(());
        }
        Self::mark_all_dirty(&mut crdt);
        self.expand_undo_scope(&crdt);
        self.crdt.set(Some(crdt));
        self.reset_history_capture();
        Ok(())
    }
}

// ---------------------------------------------------------------------------
// Private helpers (NOT wasm_bindgen)
// ---------------------------------------------------------------------------

impl DocEdit {
    const REMOTE_ORIGIN: &'static str = "docedit-remote";
    const PARAGRAPH_ATTR_KEYS_TO_CLONE_ON_SPLIT: &[&str] = &[
        "alignment",
        "headingLevel",
        "styleId",
        "spacingBeforeTwips",
        "spacingAfterTwips",
        "lineSpacingTwips",
        "lineSpacingRule",
        "indentLeftTwips",
        "indentRightTwips",
        "indentFirstLineTwips",
        "indentHangingTwips",
        "numberingKind",
        "numberingNumId",
        "numberingIlvl",
        "pageBreakBefore",
        "keepNext",
        "keepLines",
    ];

    fn intent_attr_value_to_any(value: &IntentAttrValue) -> Any {
        match value {
            IntentAttrValue::Null => Any::Null,
            IntentAttrValue::Bool(v) => Any::Bool(*v),
            IntentAttrValue::Number(v) => Any::Number(*v),
            IntentAttrValue::String(v) => Any::String(Arc::from(v.as_str())),
        }
    }

    fn normalize_color_value(value: &str) -> String {
        value.trim().trim_start_matches('#').to_uppercase()
    }

    fn normalize_optional_text(value: Option<&str>) -> Option<String> {
        value.and_then(|value| {
            let trimmed = value.trim();
            if trimmed.is_empty() {
                None
            } else {
                Some(trimmed.to_string())
            }
        })
    }

    fn is_token_metadata_attr_key(key: &str) -> bool {
        matches!(
            key,
            tokens::ATTR_TOKEN_TYPE
                | "breakType"
                | "fieldType"
                | "instr"
                | "presentation"
                | "fieldId"
                | "id"
                | "imageRef"
                | "width"
                | "height"
                | "name"
                | "description"
                | "opaqueId"
                | "xml"
        )
    }

    fn normalize_text_attr_key(key: &str) -> &str {
        match key {
            "fontSizePt" => "fontSize",
            "widthPt" => "width",
            "heightPt" => "height",
            _ => key,
        }
    }

    fn normalize_text_attr_value(key: &str, value: &IntentAttrValue) -> Any {
        match (key, value) {
            ("fontSizePt", IntentAttrValue::Number(v)) => Any::Number((v * 2.0).round()),
            ("widthPt" | "heightPt", IntentAttrValue::Number(v)) => {
                Any::Number((v.max(1.0) * 12_700.0).round())
            }
            ("color", IntentAttrValue::String(v)) => {
                Any::String(Arc::from(Self::normalize_color_value(v)))
            }
            _ => Self::intent_attr_value_to_any(value),
        }
    }

    fn collect_text_attrs(attrs: &HashMap<String, IntentAttrValue>) -> HashMap<Arc<str>, Any> {
        let mut normalized = HashMap::<Arc<str>, Any>::new();
        for (key, value) in attrs {
            normalized.insert(
                Arc::from(Self::normalize_text_attr_key(key)),
                Self::normalize_text_attr_value(key, value),
            );
        }
        normalized
    }

    fn normalize_paragraph_attr_value(key: &str, value: &IntentAttrValue) -> Any {
        match (key, value) {
            ("alignment", IntentAttrValue::String(v)) => {
                Any::String(Arc::from(v.trim().to_ascii_lowercase()))
            }
            (
                "spacingBeforePt" | "spacingAfterPt" | "indentLeftPt" | "indentFirstLinePt",
                IntentAttrValue::Number(v),
            ) => Any::Number((v.max(0.0) * 20.0).round()),
            ("lineSpacingMultiple", IntentAttrValue::Number(v)) => {
                Any::Number((v.max(0.5) * 240.0).round())
            }
            ("headingLevel", IntentAttrValue::Number(v)) => Any::Number(v.round().clamp(1.0, 9.0)),
            ("numberingNumId", IntentAttrValue::Number(v)) => Any::Number(v.round().max(1.0)),
            ("numberingIlvl", IntentAttrValue::Number(v)) => Any::Number(v.round().clamp(0.0, 8.0)),
            ("numberingKind", IntentAttrValue::String(v)) => {
                Any::String(Arc::from(v.trim().to_ascii_lowercase()))
            }
            _ => Self::intent_attr_value_to_any(value),
        }
    }

    fn collect_paragraph_attrs(attrs: &HashMap<String, IntentAttrValue>) -> HashMap<Arc<str>, Any> {
        let mut normalized = HashMap::<Arc<str>, Any>::new();
        for (key, value) in attrs {
            let normalized_key = match key.as_str() {
                "spacingBeforePt" => "spacingBeforeTwips",
                "spacingAfterPt" => "spacingAfterTwips",
                "indentLeftPt" => "indentLeftTwips",
                "indentFirstLinePt" => "indentFirstLineTwips",
                "lineSpacingMultiple" => "lineSpacingTwips",
                _ => key.as_str(),
            };
            normalized.insert(
                Arc::from(normalized_key),
                Self::normalize_paragraph_attr_value(key, value),
            );
            if key == "lineSpacingMultiple" {
                let rule_value = match value {
                    IntentAttrValue::Number(_) => Any::String(Arc::from("auto")),
                    IntentAttrValue::Null => Any::Null,
                    _ => continue,
                };
                normalized.insert(Arc::from("lineSpacingRule"), rule_value);
            }
        }
        normalized
    }

    fn collect_paragraph_formatting(
        map_ref: &yrs::MapRef,
        txn: &impl ReadTxn,
    ) -> HashMap<String, Any> {
        let mut result = HashMap::<String, Any>::new();
        for (stored_key, public_key) in [
            ("alignment", "alignment"),
            ("headingLevel", "headingLevel"),
            ("spacingBeforeTwips", "spacingBeforePt"),
            ("spacingAfterTwips", "spacingAfterPt"),
            ("indentLeftTwips", "indentLeftPt"),
            ("indentFirstLineTwips", "indentFirstLinePt"),
            ("lineSpacingTwips", "lineSpacingMultiple"),
            ("numberingKind", "numberingKind"),
            ("numberingNumId", "numberingNumId"),
            ("numberingIlvl", "numberingIlvl"),
        ] {
            if let Some(Out::Any(value)) = map_ref.get(txn, stored_key) {
                let public_value = match (stored_key, value) {
                    (
                        "spacingBeforeTwips"
                        | "spacingAfterTwips"
                        | "indentLeftTwips"
                        | "indentFirstLineTwips",
                        Any::Number(n),
                    ) => Any::Number(n / 20.0),
                    ("lineSpacingTwips", Any::Number(n)) => {
                        let rule = map_ref
                            .get(txn, "lineSpacingRule")
                            .and_then(|value| match value {
                                Out::Any(Any::String(rule)) => Some(rule.to_string()),
                                _ => None,
                            })
                            .unwrap_or_else(|| "auto".to_string());
                        if rule == "auto" {
                            Any::Number(n / 240.0)
                        } else {
                            continue;
                        }
                    }
                    (_, other) => other,
                };
                result.insert(public_key.to_string(), public_value);
            }
        }
        result
    }

    fn transact_mut_local(crdt: &CrdtDoc) -> yrs::TransactionMut<'_> {
        crdt.doc().transact_mut_with(crdt.doc().client_id())
    }

    fn transact_mut_remote(crdt: &CrdtDoc) -> yrs::TransactionMut<'_> {
        crdt.doc().transact_mut_with(Self::REMOTE_ORIGIN)
    }

    fn normalize_range(
        anchor_body: u32,
        anchor_offset: u32,
        focus_body: u32,
        focus_offset: u32,
    ) -> ((u32, u32), (u32, u32)) {
        if anchor_body < focus_body || (anchor_body == focus_body && anchor_offset <= focus_offset)
        {
            ((anchor_body, anchor_offset), (focus_body, focus_offset))
        } else {
            ((focus_body, focus_offset), (anchor_body, anchor_offset))
        }
    }

    fn decode_normalized_range(
        anchor: &str,
        focus: &str,
    ) -> Result<((u32, u32), (u32, u32)), EditError> {
        let (anchor_body_idx, anchor_offset) =
            decode_position(anchor).map_err(EditError::Position)?;
        let (focus_body_idx, focus_offset) = decode_position(focus).map_err(EditError::Position)?;
        Ok(Self::normalize_range(
            anchor_body_idx,
            anchor_offset,
            focus_body_idx,
            focus_offset,
        ))
    }

    fn utf16_offset_to_byte_index(text: &str, offset: u32) -> usize {
        if offset == 0 {
            return 0;
        }
        let mut consumed: u32 = 0;
        for (byte_idx, ch) in text.char_indices() {
            let next = consumed + ch.len_utf16() as u32;
            if next == offset {
                return byte_idx + ch.len_utf8();
            }
            if next > offset {
                // Invalid split inside a code point: snap to current boundary.
                return byte_idx;
            }
            consumed = next;
        }
        text.len()
    }

    fn decode_data_uri(data_uri: &str) -> Result<DecodedDataUri, EditError> {
        let Some((header, body)) = data_uri.split_once(',') else {
            return Err(EditError::General("image data URI missing comma".into()));
        };
        if !header.starts_with("data:") || !header.ends_with(";base64") {
            return Err(EditError::General(
                "image data URI must be base64-encoded".into(),
            ));
        }
        let content_type = header
            .trim_start_matches("data:")
            .trim_end_matches(";base64")
            .trim();
        if content_type.is_empty() {
            return Err(EditError::General(
                "image data URI missing content type".into(),
            ));
        }
        let bytes = base64::engine::general_purpose::STANDARD
            .decode(body)
            .map_err(|e| EditError::General(format!("image base64 decode failed: {e}")))?;
        Ok((content_type.to_string(), bytes))
    }

    fn token_at_utf16_offset(
        text_ref: &yrs::TextRef,
        txn: &impl ReadTxn,
        offset: u32,
    ) -> Option<tokens::TokenType> {
        let diffs = text_ref.diff(txn, yrs::types::text::YChange::identity);
        let mut pos_u16: u32 = 0;
        for diff in &diffs {
            let chunk_text = match &diff.insert {
                Out::Any(Any::String(s)) => s.to_string(),
                _ => String::new(),
            };
            if chunk_text.is_empty() {
                continue;
            }
            let chunk_len = chunk_text.encode_utf16().count() as u32;
            if pos_u16 + chunk_len > offset {
                let attrs = diff.attributes.as_ref()?.as_ref().clone();
                return tokens::attrs_to_token(&attrs);
            }
            pos_u16 += chunk_len;
        }
        None
    }

    fn token_deletable(token: &tokens::TokenType) -> bool {
        matches!(token, tokens::TokenType::InlineImage { .. })
    }

    fn contains_non_deletable_token_in_range(
        text_ref: &yrs::TextRef,
        txn: &impl ReadTxn,
        start: u32,
        end: u32,
    ) -> bool {
        if start >= end {
            return false;
        }
        let diffs = text_ref.diff(txn, yrs::types::text::YChange::identity);
        let mut pos_u16: u32 = 0;
        for diff in &diffs {
            let chunk_text = match &diff.insert {
                Out::Any(Any::String(s)) => s.to_string(),
                _ => String::new(),
            };
            if chunk_text.is_empty() {
                continue;
            }
            let chunk_len = chunk_text.encode_utf16().count() as u32;
            let chunk_end = pos_u16 + chunk_len;
            if chunk_end <= start {
                pos_u16 = chunk_end;
                continue;
            }
            if pos_u16 >= end {
                break;
            }
            if chunk_text.chars().any(tokens::is_sentinel) {
                let token = diff
                    .attributes
                    .as_ref()
                    .and_then(|attrs| tokens::attrs_to_token(attrs.as_ref()));
                if token
                    .as_ref()
                    .is_none_or(|token| !Self::token_deletable(token))
                {
                    return true;
                }
            }
            pos_u16 = chunk_end;
        }
        false
    }

    fn paragraph_text_ref_from_body_in_txn(
        body: &yrs::ArrayRef,
        txn: &yrs::TransactionMut<'_>,
        body_index: u32,
    ) -> Result<yrs::TextRef, EditError> {
        let entry = body
            .get(txn, body_index)
            .ok_or_else(|| EditError::General(format!("body[{body_index}] not found")))?;
        let map_ref = entry
            .cast::<yrs::MapRef>()
            .map_err(|_| EditError::General(format!("body[{body_index}] is not a map")))?;
        let text_out = map_ref
            .get(txn, "text")
            .ok_or_else(|| EditError::General(format!("body[{body_index}] has no text")))?;
        text_out
            .cast::<yrs::TextRef>()
            .map_err(|_| EditError::General("text is not a TextRef".into()))
    }

    fn trailing_chunks_after_offset(
        text_ref: &yrs::TextRef,
        txn: &yrs::TransactionMut<'_>,
        offset: u32,
    ) -> Vec<(String, HashMap<Arc<str>, Any>)> {
        let mut trailing: Vec<(String, HashMap<Arc<str>, Any>)> = Vec::new();
        let diffs = text_ref.diff(txn, yrs::types::text::YChange::identity);
        let mut pos_u16: u32 = 0;
        for diff in &diffs {
            let text = match &diff.insert {
                Out::Any(Any::String(s)) => s.to_string(),
                _ => String::new(),
            };
            if text.is_empty() {
                continue;
            }
            let attrs = diff
                .attributes
                .as_ref()
                .map(|b| b.as_ref().clone())
                .unwrap_or_default();
            let chunk_u16 = text.encode_utf16().count() as u32;
            let chunk_end = pos_u16 + chunk_u16;

            if chunk_end <= offset {
                pos_u16 = chunk_end;
                continue;
            }
            if pos_u16 >= offset {
                trailing.push((text, attrs));
                pos_u16 = chunk_end;
                continue;
            }

            let split_within = offset - pos_u16;
            let split_byte = Self::utf16_offset_to_byte_index(&text, split_within);
            let after = &text[split_byte..];
            if !after.is_empty() {
                trailing.push((after.to_string(), attrs));
            }
            pos_u16 = chunk_end;
        }
        trailing
    }

    fn split_paste_lines(text: &str) -> Vec<String> {
        text.replace("\r\n", "\n")
            .replace('\r', "\n")
            .split('\n')
            .map(ToString::to_string)
            .collect()
    }

    fn clone_paragraph_attrs_on_split(
        source_map_ref: &yrs::MapRef,
        target_map_ref: &yrs::MapRef,
        txn: &mut yrs::TransactionMut<'_>,
    ) {
        for key in Self::PARAGRAPH_ATTR_KEYS_TO_CLONE_ON_SPLIT {
            if let Some(Out::Any(any)) = source_map_ref.get(txn, key) {
                target_map_ref.insert(txn, *key, any);
            }
        }
    }

    fn insert_empty_cloned_paragraph_after(
        body: &yrs::ArrayRef,
        txn: &mut yrs::TransactionMut<'_>,
        body_idx: u32,
    ) -> Result<(ParaId, yrs::TextRef), EditError> {
        let new_id = ParaId::new();
        let prelim = MapPrelim::from([
            ("type".to_string(), Any::String(Arc::from("paragraph"))),
            ("id".to_string(), Any::String(Arc::from(new_id.to_string()))),
        ]);
        body.insert(txn, body_idx + 1, prelim);

        let new_map_value = body
            .get(txn, body_idx + 1)
            .ok_or_else(|| EditError::General("failed to get new paragraph".into()))?;
        let new_map_ref = new_map_value
            .cast::<yrs::MapRef>()
            .map_err(|_| EditError::General("new paragraph is not a map".into()))?;
        new_map_ref.insert(txn, "text", TextPrelim::new(""));

        if let Some(current_entry) = body.get(txn, body_idx) {
            if let Ok(current_map_ref) = current_entry.cast::<yrs::MapRef>() {
                Self::clone_paragraph_attrs_on_split(&current_map_ref, &new_map_ref, txn);
            }
        }

        let new_text_value = new_map_ref
            .get(txn, "text")
            .ok_or_else(|| EditError::General("failed to get new text ref".into()))?;
        let new_text_ref = new_text_value
            .cast::<yrs::TextRef>()
            .map_err(|_| EditError::General("text is not a TextRef".into()))?;
        Ok((new_id, new_text_ref))
    }

    fn delete_normalized_range_on(
        crdt: &mut CrdtDoc,
        start_body_idx: u32,
        start_offset: u32,
        end_body_idx: u32,
        end_offset: u32,
    ) -> Result<(u32, u32), EditError> {
        if start_body_idx == end_body_idx {
            let text_ref = Self::paragraph_text_ref_from(crdt, start_body_idx)?;
            let mut txn = Self::transact_mut_local(crdt);
            let len = text_ref.len(&txn);
            let lo = start_offset.min(len);
            let hi = end_offset.min(len);
            if lo < hi {
                if Self::contains_non_deletable_token_in_range(&text_ref, &txn, lo, hi) {
                    return Ok((start_body_idx, lo));
                }
                text_ref.remove_range(&mut txn, lo, hi - lo);
            }
            return Ok((start_body_idx, lo));
        }

        let start_text_ref = Self::paragraph_text_ref_from(crdt, start_body_idx)?;
        let end_text_ref = Self::paragraph_text_ref_from(crdt, end_body_idx)?;
        let body = crdt.body();
        let mut txn = Self::transact_mut_local(crdt);

        let start_len = start_text_ref.len(&txn);
        let end_len = end_text_ref.len(&txn);
        let start_offset = start_offset.min(start_len);
        let end_offset = end_offset.min(end_len);

        if Self::contains_non_deletable_token_in_range(
            &start_text_ref,
            &txn,
            start_offset,
            start_len,
        ) {
            return Ok((start_body_idx, start_offset));
        }

        for body_idx in (start_body_idx + 1)..end_body_idx {
            let text_ref = Self::paragraph_text_ref_from_body_in_txn(&body, &txn, body_idx)?;
            let text_len = text_ref.len(&txn);
            if Self::contains_non_deletable_token_in_range(&text_ref, &txn, 0, text_len) {
                return Ok((start_body_idx, start_offset));
            }
        }

        if Self::contains_non_deletable_token_in_range(&end_text_ref, &txn, 0, end_offset) {
            return Ok((start_body_idx, start_offset));
        }

        let trailing = Self::trailing_chunks_after_offset(&end_text_ref, &txn, end_offset);

        if start_offset < start_len {
            start_text_ref.remove_range(&mut txn, start_offset, start_len - start_offset);
        }

        let mut insert_pos = start_offset;
        for (text, attrs) in &trailing {
            if text.is_empty() {
                continue;
            }
            start_text_ref.insert_with_attributes(&mut txn, insert_pos, text, attrs.clone());
            insert_pos += text.encode_utf16().count() as u32;
        }

        for _ in (start_body_idx + 1)..=end_body_idx {
            body.remove(&mut txn, start_body_idx + 1);
        }

        Ok((start_body_idx, start_offset))
    }

    /// Clamp an offset to the text length within an existing transaction.
    ///
    /// Using a single transaction for both length-check and mutation
    /// eliminates any read/write state mismatch.
    fn clamp_offset(text_ref: &yrs::TextRef, txn: &yrs::TransactionMut<'_>, offset: u32) -> u32 {
        offset.min(text_ref.len(txn))
    }

    /// Process a single editing intent by translating it into CRDT transactions.
    ///
    /// All yrs operations use a single `TransactionMut` to ensure the length
    /// check and the actual mutation see the same CRDT state.
    fn process_intent_on(crdt: &mut CrdtDoc, intent: &EditIntent) -> Result<(), EditError> {
        match intent {
            EditIntent::InsertText {
                data,
                anchor,
                attrs: intent_attrs,
            }
            | EditIntent::InsertFromComposition {
                data,
                anchor,
                attrs: intent_attrs,
            } => {
                let (body_idx, offset) = decode_position(anchor).map_err(EditError::Position)?;
                let clean = tokens::strip_sentinels(data);
                if clean.is_empty() {
                    return Ok(());
                }

                let text_ref = Self::paragraph_text_ref_from(crdt, body_idx)?;
                let mut txn = Self::transact_mut_local(crdt);
                let offset = Self::clamp_offset(&text_ref, &txn, offset);

                // Build formatting attributes for the inserted text.
                // If the intent provides explicit attrs, use those.
                // Otherwise, inherit formatting from the character to the
                // left of the insertion point (skipping sentinel attrs).
                let insert_attrs = if let Some(provided) = intent_attrs {
                    Self::collect_text_attrs(provided)
                } else {
                    Self::inherit_attrs_at(&text_ref, &txn, offset)
                };

                text_ref.insert_with_attributes(&mut txn, offset, &clean, insert_attrs);
                drop(txn);

                Self::mark_dirty_on(crdt, body_idx)?;
                Ok(())
            }
            EditIntent::DeleteBackward { anchor, focus }
            | EditIntent::DeleteForward { anchor, focus }
            | EditIntent::DeleteByCut { anchor, focus } => {
                let ((start_body_idx, start_offset), (end_body_idx, end_offset)) =
                    Self::decode_normalized_range(anchor, focus)?;

                if start_body_idx != end_body_idx || start_offset != end_offset {
                    let (dirty_body_idx, _) = Self::delete_normalized_range_on(
                        crdt,
                        start_body_idx,
                        start_offset,
                        end_body_idx,
                        end_offset,
                    )?;
                    Self::mark_dirty_on(crdt, dirty_body_idx)?;
                    return Ok(());
                }

                let body_idx = start_body_idx;
                let lo = start_offset;
                let text_ref = Self::paragraph_text_ref_from(crdt, body_idx)?;
                let mut txn = Self::transact_mut_local(crdt);
                let len = text_ref.len(&txn);
                let lo = lo.min(len);
                // Single-char delete at a collapsed cursor.
                match intent {
                    EditIntent::DeleteBackward { .. } => {
                        if lo == 0 {
                            // Merge with previous paragraph if possible.
                            if body_idx > 0 {
                                drop(txn);
                                let prev_text_ref =
                                    Self::paragraph_text_ref_from(crdt, body_idx - 1)?;
                                let body = crdt.body();
                                let mut txn = Self::transact_mut_local(crdt);

                                // Get current paragraph content (with attrs).
                                let diffs =
                                    text_ref.diff(&txn, yrs::types::text::YChange::identity);
                                let prev_len = prev_text_ref.len(&txn);

                                // Append each chunk to the previous paragraph.
                                let mut insert_pos = prev_len;
                                for diff in &diffs {
                                    let text = match &diff.insert {
                                        Out::Any(Any::String(s)) => s.to_string(),
                                        _ => String::new(),
                                    };
                                    if text.is_empty() {
                                        continue;
                                    }
                                    let attrs = diff
                                        .attributes
                                        .as_ref()
                                        .map(|b| b.as_ref().clone())
                                        .unwrap_or_default();
                                    prev_text_ref
                                        .insert_with_attributes(&mut txn, insert_pos, &text, attrs);
                                    insert_pos += text.encode_utf16().count() as u32;
                                }

                                // Remove current paragraph from body array.
                                body.remove(&mut txn, body_idx);
                                drop(txn);

                                // Mark the merged paragraph as dirty.
                                Self::mark_dirty_on(crdt, body_idx - 1)?;
                                return Ok(());
                            }
                            return Ok(());
                        }
                        if let Some(token) = Self::token_at_utf16_offset(&text_ref, &txn, lo - 1) {
                            if !Self::token_deletable(&token) {
                                return Ok(());
                            }
                        }
                        text_ref.remove_range(&mut txn, lo - 1, 1);
                    }
                    EditIntent::DeleteForward { .. } => {
                        if lo >= len {
                            drop(txn);
                            let next_body_idx = body_idx + 1;
                            if Self::paragraph_text_ref_from(crdt, next_body_idx).is_err() {
                                return Ok(());
                            }
                            let (dirty_body_idx, _) = Self::delete_normalized_range_on(
                                crdt,
                                body_idx,
                                lo,
                                next_body_idx,
                                0,
                            )?;
                            Self::mark_dirty_on(crdt, dirty_body_idx)?;
                            return Ok(());
                        }
                        if let Some(token) = Self::token_at_utf16_offset(&text_ref, &txn, lo) {
                            if !Self::token_deletable(&token) {
                                return Ok(());
                            }
                        }
                        text_ref.remove_range(&mut txn, lo, 1);
                    }
                    _ => {
                        // DeleteByCut with collapsed selection is a no-op.
                        return Ok(());
                    }
                }
                drop(txn);

                Self::mark_dirty_on(crdt, body_idx)?;
                Ok(())
            }
            EditIntent::FormatBold { anchor, focus } => {
                Self::toggle_format_on(crdt, anchor, focus, "bold")
            }
            EditIntent::FormatItalic { anchor, focus } => {
                Self::toggle_format_on(crdt, anchor, focus, "italic")
            }
            EditIntent::FormatUnderline { anchor, focus } => {
                Self::toggle_format_on(crdt, anchor, focus, "underline")
            }
            EditIntent::FormatStrikethrough { anchor, focus } => {
                Self::toggle_format_on(crdt, anchor, focus, "strike")
            }
            EditIntent::SetTextAttrs {
                anchor,
                focus,
                attrs,
            } => Self::set_text_attrs_on(crdt, anchor, focus, attrs),
            EditIntent::SetParagraphAttrs {
                anchor,
                focus,
                attrs,
            } => Self::set_paragraph_attrs_on(crdt, anchor, focus, attrs),
            EditIntent::InsertFromPaste {
                data,
                anchor,
                focus,
                attrs: intent_attrs,
            } => {
                let ((start_body_idx, start_offset), (end_body_idx, end_offset)) =
                    Self::decode_normalized_range(anchor, focus)?;

                let clean = tokens::strip_sentinels(data);
                let (insert_body_idx, insert_offset) =
                    if start_body_idx != end_body_idx || start_offset != end_offset {
                        Self::delete_normalized_range_on(
                            crdt,
                            start_body_idx,
                            start_offset,
                            end_body_idx,
                            end_offset,
                        )?
                    } else {
                        (start_body_idx, start_offset)
                    };

                if !clean.is_empty() {
                    let lines = Self::split_paste_lines(&clean);
                    let text_ref = Self::paragraph_text_ref_from(crdt, insert_body_idx)?;
                    let body = crdt.body();
                    let mut new_paragraph_ids = Vec::new();
                    let mut txn = Self::transact_mut_local(crdt);
                    let insert_at = Self::clamp_offset(&text_ref, &txn, insert_offset);
                    let insert_attrs = if let Some(provided) = intent_attrs {
                        Self::collect_text_attrs(provided)
                    } else {
                        Self::inherit_attrs_at(&text_ref, &txn, insert_at)
                    };

                    if lines.len() <= 1 {
                        text_ref.insert_with_attributes(&mut txn, insert_at, &clean, insert_attrs);
                    } else {
                        let text_len = text_ref.len(&txn);
                        let trailing =
                            Self::trailing_chunks_after_offset(&text_ref, &txn, insert_at);
                        if insert_at < text_len {
                            text_ref.remove_range(&mut txn, insert_at, text_len - insert_at);
                        }

                        if let Some(first_line) = lines.first() {
                            if !first_line.is_empty() {
                                text_ref.insert_with_attributes(
                                    &mut txn,
                                    insert_at,
                                    first_line,
                                    insert_attrs.clone(),
                                );
                            }
                        }

                        let mut current_body_idx = insert_body_idx;
                        for line in lines.iter().skip(1) {
                            let (new_id, new_text_ref) = Self::insert_empty_cloned_paragraph_after(
                                &body,
                                &mut txn,
                                current_body_idx,
                            )?;
                            new_paragraph_ids.push(new_id);
                            current_body_idx += 1;
                            if !line.is_empty() {
                                new_text_ref.insert_with_attributes(
                                    &mut txn,
                                    0,
                                    line,
                                    insert_attrs.clone(),
                                );
                            }
                        }

                        let final_text_ref = Self::paragraph_text_ref_from_body_in_txn(
                            &body,
                            &txn,
                            current_body_idx,
                        )?;
                        let mut trailing_insert_at = final_text_ref.len(&txn);
                        for (text, attrs) in trailing {
                            final_text_ref.insert_with_attributes(
                                &mut txn,
                                trailing_insert_at,
                                &text,
                                attrs.clone(),
                            );
                            trailing_insert_at += text.encode_utf16().count() as u32;
                        }
                    }
                    drop(txn);

                    let last_dirty_idx = insert_body_idx + (lines.len() as u32).saturating_sub(1);
                    for body_idx in insert_body_idx..=last_dirty_idx {
                        Self::mark_dirty_on(crdt, body_idx)?;
                    }
                    for para_id in new_paragraph_ids {
                        crdt.register_new_paragraph(para_id);
                    }
                }

                Ok(())
            }
            EditIntent::InsertInlineImage {
                anchor,
                focus,
                data_uri,
                width_pt,
                height_pt,
                name,
                description,
            } => {
                let ((start_body_idx, start_offset), (end_body_idx, end_offset)) =
                    Self::decode_normalized_range(anchor, focus)?;
                let (insert_body_idx, insert_offset) =
                    if start_body_idx != end_body_idx || start_offset != end_offset {
                        Self::delete_normalized_range_on(
                            crdt,
                            start_body_idx,
                            start_offset,
                            end_body_idx,
                            end_offset,
                        )?
                    } else {
                        (start_body_idx, start_offset)
                    };

                let (content_type, bytes) = Self::decode_data_uri(data_uri)?;
                let image_ref = format!("img:local:{}", Uuid::now_v7());
                let width_emu = ((*width_pt).max(1.0) * 12_700.0).round() as i64;
                let height_emu = ((*height_pt).max(1.0) * 12_700.0).round() as i64;
                let name = Self::normalize_optional_text(name.as_deref());
                let description = Self::normalize_optional_text(description.as_deref());
                let token = tokens::TokenType::InlineImage {
                    image_ref: image_ref.clone(),
                    width: width_emu,
                    height: height_emu,
                };
                let mut attrs = tokens::token_to_attrs(&token);
                let sentinel = tokens::SENTINEL.to_string();
                let images = crdt.images_map();
                let text_ref = Self::paragraph_text_ref_from(crdt, insert_body_idx)?;

                let mut txn = Self::transact_mut_local(crdt);
                let insert_at = Self::clamp_offset(&text_ref, &txn, insert_offset);
                attrs.extend(Self::inherit_attrs_at(&text_ref, &txn, insert_at));
                if let Some(name) = name.as_deref() {
                    attrs.insert(Arc::from("name"), Any::String(Arc::from(name)));
                }
                if let Some(description) = description.as_deref() {
                    attrs.insert(
                        Arc::from("description"),
                        Any::String(Arc::from(description)),
                    );
                }
                images.insert(
                    &mut txn,
                    image_ref.as_str(),
                    MapPrelim::from([
                        (
                            "contentType".to_string(),
                            Any::String(Arc::from(content_type.as_str())),
                        ),
                        (
                            "dataUri".to_string(),
                            Any::String(Arc::from(data_uri.as_str())),
                        ),
                        ("width".to_string(), Any::Number(width_emu as f64)),
                        ("height".to_string(), Any::Number(height_emu as f64)),
                        (
                            "name".to_string(),
                            name.as_deref()
                                .map(|value| Any::String(Arc::from(value)))
                                .unwrap_or(Any::Null),
                        ),
                        (
                            "description".to_string(),
                            description
                                .as_deref()
                                .map(|value| Any::String(Arc::from(value)))
                                .unwrap_or(Any::Null),
                        ),
                    ]),
                );
                text_ref.insert_with_attributes(&mut txn, insert_at, &sentinel, attrs);
                drop(txn);

                crdt.image_blobs_mut().insert(image_ref, bytes);
                Self::mark_dirty_on(crdt, insert_body_idx)?;
                Ok(())
            }
            EditIntent::InsertParagraph { anchor } => {
                let (body_idx, offset) = decode_position(anchor).map_err(EditError::Position)?;
                let text_ref = Self::paragraph_text_ref_from(crdt, body_idx)?;
                let body = crdt.body();

                let mut txn = Self::transact_mut_local(crdt);
                let offset = Self::clamp_offset(&text_ref, &txn, offset);
                let text_len = text_ref.len(&txn);

                // 1. Collect content after the split point (with attributes).
                let trailing = Self::trailing_chunks_after_offset(&text_ref, &txn, offset);
                if offset < text_len {
                    // 2. Remove trailing content from current paragraph.
                    text_ref.remove_range(&mut txn, offset, text_len - offset);
                }

                // 3. Create new paragraph entry at body_idx + 1.
                let (new_id, new_text_ref) =
                    Self::insert_empty_cloned_paragraph_after(&body, &mut txn, body_idx)?;

                // 5. Write trailing content into new paragraph.
                let mut insert_pos: u32 = 0;
                for (text, attrs) in &trailing {
                    new_text_ref.insert_with_attributes(&mut txn, insert_pos, text, attrs.clone());
                    insert_pos += text.encode_utf16().count() as u32;
                }

                drop(txn);

                // Mark both paragraphs dirty.
                Self::mark_dirty_on(crdt, body_idx)?;
                crdt.register_new_paragraph(new_id);
                Ok(())
            }
            EditIntent::InsertTable {
                anchor,
                rows,
                columns,
            } => {
                let (body_idx, _) = decode_position(anchor).map_err(EditError::Position)?;
                let row_count = (*rows).max(1);
                let column_count = (*columns).max(1);
                let body = crdt.body();
                let mut txn = Self::transact_mut_local(crdt);
                let body_len = body.len(&txn);
                let insert_after = if body_len == 0 {
                    None
                } else {
                    Some(body_idx.min(body_len.saturating_sub(1)))
                };
                let table_id = Self::insert_empty_table_after(
                    &body,
                    &mut txn,
                    insert_after,
                    row_count,
                    column_count,
                )?;
                drop(txn);

                crdt.register_new_paragraph(table_id);
                Ok(())
            }
            EditIntent::SetTableCellText {
                body_index,
                row,
                col,
                text,
            } => {
                let table_map = Self::table_map_ref_from(crdt, *body_index)?;
                let key = Self::table_cell_key(*row, *col);
                let clean = tokens::strip_sentinels(text);
                let mut txn = Self::transact_mut_local(crdt);

                if table_map.get(&txn, key.as_str()).is_none() {
                    table_map.insert(&mut txn, key.as_str(), TextPrelim::new(""));
                }

                let text_out = table_map
                    .get(&txn, key.as_str())
                    .ok_or_else(|| EditError::General(format!("missing table cell {key}")))?;
                let text_ref = text_out
                    .cast::<yrs::TextRef>()
                    .map_err(|_| EditError::General(format!("table cell {key} is not text")))?;
                let old_len = text_ref.len(&txn);
                if old_len > 0 {
                    text_ref.remove_range(&mut txn, 0, old_len);
                }
                if !clean.is_empty() {
                    text_ref.insert(&mut txn, 0, &clean);
                }
                drop(txn);

                Self::mark_dirty_on(crdt, *body_index)?;
                Ok(())
            }
            EditIntent::InsertTableRow { body_index, row } => {
                Self::rewrite_table_on(crdt, *body_index, |cells, rows, columns| {
                    let insert_at = (*row).min(rows);
                    let mut next = vec![vec![String::new(); columns]; rows + 1];
                    for (src_row, row_cells) in cells.iter().enumerate().take(rows) {
                        let dst_row = if src_row < insert_at {
                            src_row
                        } else {
                            src_row + 1
                        };
                        next[dst_row] = row_cells.clone();
                    }
                    (next, rows + 1, columns)
                })?;
                Ok(())
            }
            EditIntent::RemoveTableRow { body_index, row } => {
                Self::rewrite_table_on(crdt, *body_index, |cells, rows, columns| {
                    if rows <= 1 {
                        return (cells, rows, columns);
                    }
                    let remove_at = (*row).min(rows - 1);
                    let mut next = Vec::with_capacity(rows - 1);
                    for (src_row, row_cells) in cells.into_iter().enumerate() {
                        if src_row == remove_at {
                            continue;
                        }
                        next.push(row_cells);
                    }
                    (next, rows - 1, columns)
                })?;
                Ok(())
            }
            EditIntent::InsertTableColumn { body_index, col } => {
                Self::rewrite_table_on(crdt, *body_index, |cells, rows, columns| {
                    let insert_at = (*col).min(columns);
                    let mut next = vec![vec![String::new(); columns + 1]; rows];
                    for (row_idx, row_cells) in cells.iter().enumerate().take(rows) {
                        for (src_col, text) in row_cells.iter().enumerate().take(columns) {
                            let dst_col = if src_col < insert_at {
                                src_col
                            } else {
                                src_col + 1
                            };
                            next[row_idx][dst_col] = text.clone();
                        }
                    }
                    (next, rows, columns + 1)
                })?;
                Ok(())
            }
            EditIntent::RemoveTableColumn { body_index, col } => {
                Self::rewrite_table_on(crdt, *body_index, |cells, rows, columns| {
                    if columns <= 1 {
                        return (cells, rows, columns);
                    }
                    let remove_at = (*col).min(columns - 1);
                    let mut next = vec![Vec::with_capacity(columns - 1); rows];
                    for row_idx in 0..rows {
                        for (src_col, text) in cells[row_idx].iter().enumerate() {
                            if src_col == remove_at {
                                continue;
                            }
                            next[row_idx].push(text.clone());
                        }
                    }
                    (next, rows, columns - 1)
                })?;
                Ok(())
            }
            EditIntent::InsertLineBreak { anchor } => {
                let (body_idx, offset) = decode_position(anchor).map_err(EditError::Position)?;
                let text_ref = Self::paragraph_text_ref_from(crdt, body_idx)?;
                let token = tokens::TokenType::LineBreak { break_type: None };
                let attrs = tokens::token_to_attrs(&token);
                let sentinel = tokens::SENTINEL.to_string();
                let mut txn = Self::transact_mut_local(crdt);
                let offset = Self::clamp_offset(&text_ref, &txn, offset);
                text_ref.insert_with_attributes(&mut txn, offset, &sentinel, attrs);
                drop(txn);
                Self::mark_dirty_on(crdt, body_idx)?;
                Ok(())
            }
            EditIntent::InsertTab { anchor } => {
                let (body_idx, offset) = decode_position(anchor).map_err(EditError::Position)?;
                let text_ref = Self::paragraph_text_ref_from(crdt, body_idx)?;
                let token = tokens::TokenType::Tab;
                let attrs = tokens::token_to_attrs(&token);
                let sentinel = tokens::SENTINEL.to_string();
                let mut txn = Self::transact_mut_local(crdt);
                let offset = Self::clamp_offset(&text_ref, &txn, offset);
                text_ref.insert_with_attributes(&mut txn, offset, &sentinel, attrs);
                drop(txn);
                Self::mark_dirty_on(crdt, body_idx)?;
                Ok(())
            }
            EditIntent::Undo | EditIntent::Redo => {
                // Handled in apply_intent() directly via the UndoManager.
                Ok(())
            }
        }
    }

    /// Get the body map for a body item at the given index.
    fn body_map_ref_from(crdt: &CrdtDoc, body_index: u32) -> Result<yrs::MapRef, EditError> {
        let txn = crdt.doc().transact();
        let body = crdt.body();
        let entry = body
            .get(&txn, body_index)
            .ok_or_else(|| EditError::General(format!("body[{body_index}] not found")))?;
        entry
            .cast::<yrs::MapRef>()
            .map_err(|_| EditError::General(format!("body[{body_index}] is not a map")))
    }

    /// Get the paragraph map for a body item at the given index.
    fn paragraph_map_ref_from(crdt: &CrdtDoc, body_index: u32) -> Result<yrs::MapRef, EditError> {
        let txn = crdt.doc().transact();
        let map_ref = Self::body_map_ref_from(crdt, body_index)?;
        let type_str = map_ref.get(&txn, "type").and_then(|value| match value {
            Out::Any(Any::String(value)) => Some(value.to_string()),
            _ => None,
        });
        if type_str.as_deref() != Some("paragraph") {
            return Err(EditError::General(format!(
                "body[{body_index}] is not a paragraph"
            )));
        }
        Ok(map_ref)
    }

    /// Get the `TextRef` for a paragraph at the given body array index.
    fn paragraph_text_ref_from(crdt: &CrdtDoc, body_index: u32) -> Result<yrs::TextRef, EditError> {
        let txn = crdt.doc().transact();
        let map_ref = Self::paragraph_map_ref_from(crdt, body_index)?;
        let text_out = map_ref
            .get(&txn, "text")
            .ok_or_else(|| EditError::General(format!("body[{body_index}] has no text")))?;
        text_out
            .cast::<yrs::TextRef>()
            .map_err(|_| EditError::General("text is not a TextRef".into()))
    }

    fn table_map_ref_from(crdt: &CrdtDoc, body_index: u32) -> Result<yrs::MapRef, EditError> {
        let txn = crdt.doc().transact();
        let map_ref = Self::body_map_ref_from(crdt, body_index)?;
        let type_str = map_ref.get(&txn, "type").and_then(|value| match value {
            Out::Any(Any::String(value)) => Some(value.to_string()),
            _ => None,
        });
        if type_str.as_deref() != Some("table") {
            return Err(EditError::General(format!(
                "body[{body_index}] is not a table"
            )));
        }
        Ok(map_ref)
    }

    fn table_cell_key(row: usize, col: usize) -> String {
        format!("cell:{row}:{col}")
    }

    fn insert_empty_table_after(
        body: &yrs::ArrayRef,
        txn: &mut yrs::TransactionMut<'_>,
        insert_after_body_idx: Option<u32>,
        rows: usize,
        columns: usize,
    ) -> Result<ParaId, EditError> {
        let new_id = ParaId::new();
        let insert_idx = insert_after_body_idx
            .map(|body_idx| body_idx.saturating_add(1))
            .unwrap_or(0);
        let prelim = MapPrelim::from([
            ("type".to_string(), Any::String(Arc::from("table"))),
            ("id".to_string(), Any::String(Arc::from(new_id.to_string()))),
            ("rows".to_string(), Any::Number(rows as f64)),
            ("columns".to_string(), Any::Number(columns as f64)),
        ]);
        body.insert(txn, insert_idx, prelim);

        let table_value = body
            .get(txn, insert_idx)
            .ok_or_else(|| EditError::General("failed to get new table".into()))?;
        let table_map = table_value
            .cast::<yrs::MapRef>()
            .map_err(|_| EditError::General("new table is not a map".into()))?;

        for row in 0..rows {
            for col in 0..columns {
                let key = Self::table_cell_key(row, col);
                table_map.insert(txn, key.as_str(), TextPrelim::new(""));
            }
        }

        Ok(new_id)
    }

    fn table_dimensions(
        table_map: &yrs::MapRef,
        txn: &impl ReadTxn,
    ) -> Result<(usize, usize), EditError> {
        let rows = table_map
            .get(txn, "rows")
            .and_then(|value| match value {
                Out::Any(Any::Number(value)) => Some(value.round().max(0.0) as usize),
                _ => None,
            })
            .ok_or_else(|| EditError::General("table is missing row count".into()))?;
        let columns = table_map
            .get(txn, "columns")
            .and_then(|value| match value {
                Out::Any(Any::Number(value)) => Some(value.round().max(0.0) as usize),
                _ => None,
            })
            .ok_or_else(|| EditError::General("table is missing column count".into()))?;
        Ok((rows, columns))
    }

    fn snapshot_table_cells(
        table_map: &yrs::MapRef,
        txn: &impl ReadTxn,
        rows: usize,
        columns: usize,
    ) -> Vec<Vec<String>> {
        let mut cells = vec![vec![String::new(); columns]; rows];
        for (row_idx, row_cells) in cells.iter_mut().enumerate().take(rows) {
            for (col_idx, cell_text) in row_cells.iter_mut().enumerate().take(columns) {
                let key = Self::table_cell_key(row_idx, col_idx);
                *cell_text = table_map
                    .get(txn, key.as_str())
                    .and_then(|value| value.cast::<yrs::TextRef>().ok())
                    .map(|text_ref| text_ref.get_string(txn))
                    .unwrap_or_default();
            }
        }
        cells
    }

    fn write_table_cells(
        table_map: &yrs::MapRef,
        txn: &mut yrs::TransactionMut<'_>,
        cells: &[Vec<String>],
        rows: usize,
        columns: usize,
    ) {
        table_map.insert(txn, "rows", Any::Number(rows as f64));
        table_map.insert(txn, "columns", Any::Number(columns as f64));

        let keys_to_remove: Vec<String> = table_map
            .iter(txn)
            .filter_map(|(key, _)| {
                let key_str: &str = key;
                key_str.starts_with("cell:").then(|| key_str.to_string())
            })
            .collect();
        for key in keys_to_remove {
            table_map.remove(txn, key.as_str());
        }

        for (row_idx, row_cells) in cells.iter().enumerate().take(rows) {
            for (col_idx, cell_text) in row_cells.iter().enumerate().take(columns) {
                let key = Self::table_cell_key(row_idx, col_idx);
                table_map.insert(txn, key.as_str(), TextPrelim::new(cell_text));
            }
        }
    }

    fn rewrite_table_on<F>(crdt: &mut CrdtDoc, body_index: u32, rewrite: F) -> Result<(), EditError>
    where
        F: FnOnce(Vec<Vec<String>>, usize, usize) -> (Vec<Vec<String>>, usize, usize),
    {
        let table_map = Self::table_map_ref_from(crdt, body_index)?;
        let mut txn = Self::transact_mut_local(crdt);
        let (rows, columns) = Self::table_dimensions(&table_map, &txn)?;
        let cells = Self::snapshot_table_cells(&table_map, &txn, rows, columns);
        let (next_cells, next_rows, next_columns) = rewrite(cells, rows, columns);
        Self::write_table_cells(&table_map, &mut txn, &next_cells, next_rows, next_columns);
        drop(txn);
        Self::mark_dirty_on(crdt, body_index)?;
        Ok(())
    }

    /// Mark the body item at the given body index as dirty.
    fn mark_dirty_on(crdt: &mut CrdtDoc, body_index: u32) -> Result<(), EditError> {
        let txn = crdt.doc().transact();
        let body = crdt.body();
        let entry = body
            .get(&txn, body_index)
            .ok_or_else(|| EditError::General(format!("body[{body_index}] not found")))?;
        let map_ref = entry
            .cast::<yrs::MapRef>()
            .map_err(|_| EditError::General("not a map".into()))?;
        let id_str = match map_ref.get(&txn, "id") {
            Some(Out::Any(Any::String(s))) => s.to_string(),
            _ => return Err(EditError::General("body item has no id".into())),
        };
        drop(txn);
        let para_id: ParaId = id_str
            .parse()
            .map_err(|_| EditError::General(format!("invalid ParaId: {id_str}")))?;
        crdt.mark_dirty(&para_id);
        Ok(())
    }

    /// Inherit formatting attributes from the character to the left of `offset`.
    ///
    /// Walks the diff chunks to find which one contains `offset - 1`, then
    /// copies its formatting attributes (bold, italic, etc.) — excluding
    /// sentinel token attrs like `_tokenType`.
    fn inherit_attrs_at(
        text_ref: &yrs::TextRef,
        txn: &yrs::TransactionMut<'_>,
        offset: u32,
    ) -> HashMap<Arc<str>, Any> {
        if offset == 0 {
            return HashMap::new();
        }
        let diffs = text_ref.diff(txn, yrs::types::text::YChange::identity);
        let target = offset - 1;
        let mut pos: u32 = 0;
        for diff in &diffs {
            let chunk_text = match &diff.insert {
                Out::Any(Any::String(s)) => s.to_string(),
                _ => String::new(),
            };
            let chunk_len = chunk_text.chars().count() as u32;
            if pos + chunk_len > target {
                // This chunk contains the character to the left of the cursor.
                if let Some(attrs) = &diff.attributes {
                    // Copy formatting attrs, skip sentinel/token attrs.
                    let mut result = HashMap::<Arc<str>, Any>::new();
                    for (k, v) in attrs.as_ref() {
                        let key: &str = k;
                        if key.starts_with('_') || Self::is_token_metadata_attr_key(key) {
                            continue;
                        }
                        match v {
                            Any::Bool(true) | Any::String(_) | Any::Number(_) => {
                                result.insert(Arc::<str>::clone(k), v.clone());
                            }
                            _ => {}
                        }
                    }
                    return result;
                }
                return HashMap::new();
            }
            pos += chunk_len;
        }
        HashMap::new()
    }

    /// Query formatting attributes at the character to the left of `offset`.
    ///
    /// Like `inherit_attrs_at` but takes a read-only transaction.
    fn query_formatting_at(
        text_ref: &yrs::TextRef,
        txn: &impl ReadTxn,
        offset: u32,
    ) -> HashMap<Arc<str>, Any> {
        if offset == 0 {
            return HashMap::new();
        }
        let diffs = text_ref.diff(txn, yrs::types::text::YChange::identity);
        let target = offset - 1;
        let mut pos: u32 = 0;
        for diff in &diffs {
            let chunk_text = match &diff.insert {
                Out::Any(Any::String(s)) => s.to_string(),
                _ => String::new(),
            };
            let chunk_len = chunk_text.chars().count() as u32;
            if pos + chunk_len > target {
                if let Some(attrs) = &diff.attributes {
                    let mut result = HashMap::<Arc<str>, Any>::new();
                    for (k, v) in attrs.as_ref() {
                        let key: &str = k;
                        if key.starts_with('_') || Self::is_token_metadata_attr_key(key) {
                            continue;
                        }
                        match v {
                            Any::Bool(true) | Any::String(_) | Any::Number(_) => {
                                result.insert(Arc::<str>::clone(k), v.clone());
                            }
                            _ => {}
                        }
                    }
                    return result;
                }
                return HashMap::new();
            }
            pos += chunk_len;
        }
        HashMap::new()
    }

    /// Build undo manager options with a platform-appropriate clock.
    fn undo_options() -> yrs::undo::Options<()> {
        use std::collections::HashSet;

        // On WASM, SystemClock is not available. Use js_sys::Date::now()
        // or a simple counter as the clock source.
        #[cfg(target_arch = "wasm32")]
        let timestamp: Arc<dyn yrs::sync::Clock> = Arc::new(|| js_sys::Date::now() as u64);

        #[cfg(not(target_arch = "wasm32"))]
        let timestamp: Arc<dyn yrs::sync::Clock> = Arc::new(|| {
            std::time::SystemTime::now()
                .duration_since(std::time::UNIX_EPOCH)
                .map(|d| d.as_millis() as u64)
                .unwrap_or(0)
        });

        yrs::undo::Options {
            capture_timeout_millis: 500,
            tracked_origins: HashSet::new(),
            capture_transaction: None,
            timestamp,
            init_undo_stack: Vec::new(),
            init_redo_stack: Vec::new(),
        }
    }

    /// Create an UndoManager scoped to all editable body TextRefs.
    fn create_undo_manager(crdt: &CrdtDoc) -> yrs::undo::UndoManager<()> {
        let body = crdt.body();
        let mut mgr =
            yrs::undo::UndoManager::with_scope_and_options(crdt.doc(), &body, Self::undo_options());
        mgr.include_origin(crdt.doc().client_id());
        let images = crdt.images_map();
        mgr.expand_scope(&images);
        let txn = crdt.doc().transact();
        for i in 0..body.len(&txn) {
            if let Some(item) = body.get(&txn, i) {
                if let Ok(map_ref) = item.cast::<yrs::MapRef>() {
                    Self::expand_undo_scope_for_body_map(&mut mgr, &map_ref, &txn);
                }
            }
        }
        mgr
    }

    /// Expand the UndoManager's scope to include all current editable body TextRefs.
    ///
    /// Called after body item insertion to ensure new text refs are tracked.
    fn expand_undo_scope(&self, crdt: &CrdtDoc) {
        let mut mgr = match self.undo_mgr.take() {
            Some(m) => m,
            None => return,
        };
        let images = crdt.images_map();
        mgr.expand_scope(&images);
        let txn = crdt.doc().transact();
        let body = crdt.body();
        for i in 0..body.len(&txn) {
            if let Some(item) = body.get(&txn, i) {
                if let Ok(map_ref) = item.cast::<yrs::MapRef>() {
                    Self::expand_undo_scope_for_body_map(&mut mgr, &map_ref, &txn);
                }
            }
        }
        self.undo_mgr.set(Some(mgr));
    }

    fn expand_undo_scope_for_body_map(
        mgr: &mut yrs::undo::UndoManager<()>,
        map_ref: &yrs::MapRef,
        txn: &impl ReadTxn,
    ) {
        for (key, value) in map_ref.iter(txn) {
            let key_str: &str = key;
            if key_str != "text" && !key_str.starts_with("cell:") {
                continue;
            }
            if let Ok(text_ref) = value.cast::<yrs::TextRef>() {
                mgr.expand_scope(&text_ref);
            }
        }
    }

    fn reset_history_capture(&self) {
        if let Some(mut mgr) = self.undo_mgr.take() {
            mgr.reset();
            self.undo_mgr.set(Some(mgr));
        }
    }

    fn should_reset_history_before(intent: &EditIntent) -> bool {
        match intent {
            EditIntent::InsertFromPaste { .. }
            | EditIntent::InsertFromComposition { .. }
            | EditIntent::InsertParagraph { .. }
            | EditIntent::InsertInlineImage { .. }
            | EditIntent::InsertTable { .. }
            | EditIntent::InsertLineBreak { .. }
            | EditIntent::InsertTab { .. }
            | EditIntent::SetTableCellText { .. }
            | EditIntent::FormatBold { .. }
            | EditIntent::FormatItalic { .. }
            | EditIntent::FormatUnderline { .. }
            | EditIntent::FormatStrikethrough { .. }
            | EditIntent::SetTextAttrs { .. }
            | EditIntent::SetParagraphAttrs { .. }
            | EditIntent::DeleteByCut { .. } => true,
            EditIntent::DeleteBackward { anchor, focus }
            | EditIntent::DeleteForward { anchor, focus } => anchor != focus,
            _ => false,
        }
    }

    fn should_reset_history_after(intent: &EditIntent) -> bool {
        Self::should_reset_history_before(intent)
    }

    /// Mark all paragraphs as dirty.
    ///
    /// Used after undo/redo, which may modify any paragraph.
    fn mark_all_dirty(crdt: &mut CrdtDoc) {
        let ids: Vec<_> = crdt.para_index_map().keys().cloned().collect();
        for id in ids {
            crdt.mark_dirty(&id);
        }
    }

    fn set_text_attrs_on(
        crdt: &mut CrdtDoc,
        anchor: &str,
        focus: &str,
        attrs: &HashMap<String, IntentAttrValue>,
    ) -> Result<(), EditError> {
        let ((start_body_idx, start_offset), (end_body_idx, end_offset)) =
            Self::decode_normalized_range(anchor, focus)?;
        if start_body_idx == end_body_idx && start_offset == end_offset {
            return Ok(());
        }

        let patch = Self::collect_text_attrs(attrs);
        if patch.is_empty() {
            return Ok(());
        }

        let mut dirty_body_indices = Vec::new();
        let body = crdt.body();
        let mut txn = Self::transact_mut_local(crdt);
        for body_idx in start_body_idx..=end_body_idx {
            let Ok(text_ref) = Self::paragraph_text_ref_from_body_in_txn(&body, &txn, body_idx)
            else {
                continue;
            };
            let len = text_ref.len(&txn);
            let lo = if body_idx == start_body_idx {
                start_offset.min(len)
            } else {
                0
            };
            let hi = if body_idx == end_body_idx {
                end_offset.min(len)
            } else {
                len
            };
            if lo >= hi {
                continue;
            }
            text_ref.format(&mut txn, lo, hi - lo, patch.clone());
            dirty_body_indices.push(body_idx);
        }
        drop(txn);

        for body_idx in dirty_body_indices {
            Self::mark_dirty_on(crdt, body_idx)?;
        }
        Ok(())
    }

    fn set_paragraph_attrs_on(
        crdt: &mut CrdtDoc,
        anchor: &str,
        focus: &str,
        attrs: &HashMap<String, IntentAttrValue>,
    ) -> Result<(), EditError> {
        let ((start_body_idx, _), (end_body_idx, _)) =
            Self::decode_normalized_range(anchor, focus)?;
        let patch = Self::collect_paragraph_attrs(attrs);
        if patch.is_empty() {
            return Ok(());
        }

        let mut dirty_body_indices = Vec::new();
        let body = crdt.body();
        let mut txn = Self::transact_mut_local(crdt);
        for body_idx in start_body_idx..=end_body_idx {
            let Some(entry) = body.get(&txn, body_idx) else {
                continue;
            };
            let Ok(map_ref) = entry.cast::<yrs::MapRef>() else {
                continue;
            };
            let is_paragraph = map_ref
                .get(&txn, "type")
                .and_then(|value| match value {
                    Out::Any(Any::String(s)) => Some(s),
                    _ => None,
                })
                .is_some_and(|value| value.as_ref() == "paragraph");
            if !is_paragraph {
                continue;
            }
            for (key, value) in &patch {
                if key.as_ref() == "headingLevel" {
                    match value {
                        Any::Number(level) => {
                            let heading_level = level.round().clamp(1.0, 9.0) as u8;
                            map_ref.insert(
                                &mut txn,
                                "headingLevel",
                                Any::Number(f64::from(heading_level)),
                            );
                            map_ref.insert(
                                &mut txn,
                                "styleId",
                                Any::String(Arc::from(format!("Heading{heading_level}"))),
                            );
                        }
                        Any::Null => {
                            let _ = map_ref.remove(&mut txn, "headingLevel");
                            let clear_heading_style = map_ref
                                .get(&txn, "styleId")
                                .and_then(|value| match value {
                                    Out::Any(Any::String(style_id)) => Some(style_id),
                                    _ => None,
                                })
                                .is_some_and(|style_id| {
                                    let lower = style_id.to_ascii_lowercase();
                                    lower.starts_with("heading")
                                });
                            if clear_heading_style {
                                let _ = map_ref.remove(&mut txn, "styleId");
                            }
                        }
                        _ => {}
                    }
                    continue;
                }
                match value {
                    Any::Null => {
                        let _ = map_ref.remove(&mut txn, key.as_ref());
                    }
                    _ => {
                        map_ref.insert(&mut txn, key.as_ref(), value.clone());
                    }
                }
            }
            dirty_body_indices.push(body_idx);
        }
        drop(txn);

        for body_idx in dirty_body_indices {
            Self::mark_dirty_on(crdt, body_idx)?;
        }
        Ok(())
    }

    /// Toggle a formatting attribute on a range.
    fn toggle_format_on(
        crdt: &mut CrdtDoc,
        anchor: &str,
        focus: &str,
        attr: &str,
    ) -> Result<(), EditError> {
        let (body_idx, start) = decode_position(anchor).map_err(EditError::Position)?;
        let (_, end) = decode_position(focus).map_err(EditError::Position)?;
        let (lo, hi) = if start <= end {
            (start, end)
        } else {
            (end, start)
        };
        if lo == hi {
            return Ok(());
        }

        let text_ref = Self::paragraph_text_ref_from(crdt, body_idx)?;
        let mut txn = Self::transact_mut_local(crdt);

        // Clamp offsets to text length within the same transaction.
        let len = text_ref.len(&txn);
        let lo = lo.min(len);
        let hi = hi.min(len);
        if lo == hi {
            return Ok(());
        }

        // Check current formatting to determine toggle direction.
        let diffs = text_ref.diff(&txn, yrs::types::text::YChange::identity);

        let is_formatted = diffs.iter().any(|d| {
            d.attributes
                .as_ref()
                .is_some_and(|attrs| matches!(attrs.get(attr as &str), Some(Any::Bool(true))))
        });

        let value = if is_formatted {
            Any::Null // Remove formatting
        } else {
            Any::Bool(true)
        };

        let mut attrs: HashMap<Arc<str>, Any> = HashMap::new();
        attrs.insert(Arc::from(attr), value);

        text_ref.format(&mut txn, lo, hi - lo, attrs);
        drop(txn);

        Self::mark_dirty_on(crdt, body_idx)?;
        Ok(())
    }
}

// ---------------------------------------------------------------------------
// Tests
// ---------------------------------------------------------------------------

#[cfg(test)]
#[allow(clippy::unwrap_used, clippy::expect_used, clippy::panic)]
mod tests {
    use super::*;
    use yrs::GetString;

    /// Construct a `DocEdit` directly (without WASM) from a single-paragraph doc.
    fn make_test_editor(text: &str) -> DocEdit {
        let mut doc = offidized_docx::Document::new();
        doc.add_paragraph(text);
        let bytes = doc.to_bytes().expect("to_bytes");
        let doc2 = offidized_docx::Document::from_bytes(&bytes).expect("from_bytes");
        let mut crdt = CrdtDoc::new();
        import::import_document(&doc2, &mut crdt).expect("import");
        let undo_mgr = DocEdit::create_undo_manager(&crdt);
        DocEdit {
            crdt: Cell::new(Some(crdt)),
            original_doc: doc2,
            original_bytes: bytes,
            undo_mgr: Cell::new(Some(undo_mgr)),
        }
    }

    fn make_test_editor_with_paragraphs(lines: &[&str]) -> DocEdit {
        let mut doc = offidized_docx::Document::new();
        if lines.is_empty() {
            doc.add_paragraph("");
        } else {
            for line in lines {
                doc.add_paragraph(*line);
            }
        }
        let bytes = doc.to_bytes().expect("to_bytes");
        let doc2 = offidized_docx::Document::from_bytes(&bytes).expect("from_bytes");
        let mut crdt = CrdtDoc::new();
        import::import_document(&doc2, &mut crdt).expect("import");
        let undo_mgr = DocEdit::create_undo_manager(&crdt);
        DocEdit {
            crdt: Cell::new(Some(crdt)),
            original_doc: doc2,
            original_bytes: bytes,
            undo_mgr: Cell::new(Some(undo_mgr)),
        }
    }

    /// Apply an intent to the editor via the inner CRDT.
    fn apply(editor: &DocEdit, intent: &EditIntent) {
        let mut crdt = editor.crdt.take().expect("crdt");
        DocEdit::process_intent_on(&mut crdt, intent).expect("intent");
        editor.crdt.set(Some(crdt));
    }

    fn apply_via_api(editor: &DocEdit, intent: &EditIntent) {
        let json = serde_json::to_string(intent).expect("serialize intent");
        editor.apply_intent(&json).expect("apply intent");
    }

    /// Get the text content of paragraph 0 from the CRDT.
    fn para_text(editor: &DocEdit) -> String {
        let crdt = editor.crdt.take().expect("crdt");
        let txn = crdt.doc().transact();
        let body = crdt.body();
        let entry = body.get(&txn, 0).expect("body[0]");
        let map_ref = entry.cast::<yrs::MapRef>().expect("is map");
        let text_out = map_ref.get(&txn, "text").expect("has text");
        let text_ref = text_out.cast::<yrs::TextRef>().expect("is TextRef");
        let text = text_ref.get_string(&txn);
        drop(txn);
        editor.crdt.set(Some(crdt));
        text
    }

    fn para_text_at(editor: &DocEdit, index: u32) -> String {
        let crdt = editor.crdt.take().expect("crdt");
        let txn = crdt.doc().transact();
        let body = crdt.body();
        let entry = body.get(&txn, index).expect("body[index]");
        let map_ref = entry.cast::<yrs::MapRef>().expect("is map");
        let text_out = map_ref.get(&txn, "text").expect("has text");
        let text_ref = text_out.cast::<yrs::TextRef>().expect("is TextRef");
        let text = text_ref.get_string(&txn);
        drop(txn);
        editor.crdt.set(Some(crdt));
        text
    }

    fn body_len(editor: &DocEdit) -> u32 {
        let crdt = editor.crdt.take().expect("crdt");
        let txn = crdt.doc().transact();
        let len = crdt.body().len(&txn);
        drop(txn);
        editor.crdt.set(Some(crdt));
        len
    }

    fn body_item_type(editor: &DocEdit, index: u32) -> Option<String> {
        let crdt = editor.crdt.take().expect("crdt");
        let txn = crdt.doc().transact();
        let value = crdt
            .body()
            .get(&txn, index)
            .and_then(|entry| entry.cast::<yrs::MapRef>().ok())
            .and_then(|map_ref| map_ref.get(&txn, "type"))
            .and_then(|value| match value {
                Out::Any(Any::String(value)) => Some(value.to_string()),
                _ => None,
            });
        drop(txn);
        editor.crdt.set(Some(crdt));
        value
    }

    fn table_dimensions(editor: &DocEdit, body_index: u32) -> Option<(usize, usize)> {
        let crdt = editor.crdt.take().expect("crdt");
        let txn = crdt.doc().transact();
        let value = crdt
            .body()
            .get(&txn, body_index)
            .and_then(|entry| entry.cast::<yrs::MapRef>().ok())
            .and_then(|map_ref| {
                let rows = map_ref.get(&txn, "rows").and_then(|value| match value {
                    Out::Any(Any::Number(value)) => Some(value as usize),
                    _ => None,
                })?;
                let columns = map_ref.get(&txn, "columns").and_then(|value| match value {
                    Out::Any(Any::Number(value)) => Some(value as usize),
                    _ => None,
                })?;
                Some((rows, columns))
            });
        drop(txn);
        editor.crdt.set(Some(crdt));
        value
    }

    fn table_cell_text(editor: &DocEdit, body_index: u32, row: usize, col: usize) -> String {
        let crdt = editor.crdt.take().expect("crdt");
        let txn = crdt.doc().transact();
        let key = format!("cell:{row}:{col}");
        let text = crdt
            .body()
            .get(&txn, body_index)
            .and_then(|entry| entry.cast::<yrs::MapRef>().ok())
            .and_then(|map_ref| map_ref.get(&txn, key.as_str()))
            .and_then(|value| value.cast::<yrs::TextRef>().ok())
            .map(|text_ref| text_ref.get_string(&txn))
            .unwrap_or_default();
        drop(txn);
        editor.crdt.set(Some(crdt));
        text
    }

    fn paragraph_heading_level(editor: &DocEdit, index: u32) -> Option<f64> {
        let crdt = editor.crdt.take().expect("crdt");
        let txn = crdt.doc().transact();
        let value = crdt
            .body()
            .get(&txn, index)
            .and_then(|entry| entry.cast::<yrs::MapRef>().ok())
            .and_then(|map_ref| map_ref.get(&txn, "headingLevel"))
            .and_then(|value| match value {
                Out::Any(Any::Number(n)) => Some(n),
                _ => None,
            });
        drop(txn);
        editor.crdt.set(Some(crdt));
        value
    }

    fn paragraph_alignment(editor: &DocEdit, index: u32) -> Option<String> {
        let crdt = editor.crdt.take().expect("crdt");
        let txn = crdt.doc().transact();
        let value = crdt
            .body()
            .get(&txn, index)
            .and_then(|entry| entry.cast::<yrs::MapRef>().ok())
            .and_then(|map_ref| map_ref.get(&txn, "alignment"))
            .and_then(|value| match value {
                Out::Any(Any::String(value)) => Some(value.to_string()),
                _ => None,
            });
        drop(txn);
        editor.crdt.set(Some(crdt));
        value
    }

    fn paragraph_numbering(
        editor: &DocEdit,
        index: u32,
    ) -> (Option<String>, Option<f64>, Option<f64>) {
        let crdt = editor.crdt.take().expect("crdt");
        let txn = crdt.doc().transact();
        let value = crdt
            .body()
            .get(&txn, index)
            .and_then(|entry| entry.cast::<yrs::MapRef>().ok())
            .map(|map_ref| {
                let kind = map_ref
                    .get(&txn, "numberingKind")
                    .and_then(|value| match value {
                        Out::Any(Any::String(v)) => Some(v.to_string()),
                        _ => None,
                    });
                let num_id = map_ref
                    .get(&txn, "numberingNumId")
                    .and_then(|value| match value {
                        Out::Any(Any::Number(v)) => Some(v),
                        _ => None,
                    });
                let ilvl = map_ref
                    .get(&txn, "numberingIlvl")
                    .and_then(|value| match value {
                        Out::Any(Any::Number(v)) => Some(v),
                        _ => None,
                    });
                (kind, num_id, ilvl)
            })
            .unwrap_or((None, None, None));
        drop(txn);
        editor.crdt.set(Some(crdt));
        value
    }

    fn formatting_map(editor: &DocEdit, body_index: u32, offset: u32) -> HashMap<String, Any> {
        let crdt = editor.crdt.take().expect("crdt");
        let text_ref = DocEdit::paragraph_text_ref_from(&crdt, body_index).expect("text ref");
        let paragraph_map = DocEdit::paragraph_map_ref_from(&crdt, body_index).expect("paragraph");
        let txn = crdt.doc().transact();
        let mut result = HashMap::<String, Any>::new();
        for (key, value) in DocEdit::query_formatting_at(&text_ref, &txn, offset) {
            result.insert(key.to_string(), value);
        }
        for (key, value) in DocEdit::collect_paragraph_formatting(&paragraph_map, &txn) {
            result.insert(key, value);
        }
        drop(txn);
        editor.crdt.set(Some(crdt));
        result
    }

    fn view_model(editor: &DocEdit) -> crate::model::DocViewModel {
        let crdt = editor.crdt.take().expect("crdt");
        let vm = view::crdt_to_view_model(&crdt, &editor.original_doc).expect("view model");
        editor.crdt.set(Some(crdt));
        vm
    }

    fn inline_image_ref_at(editor: &DocEdit, body_index: u32) -> Option<String> {
        let crdt = editor.crdt.take().expect("crdt");
        let text_ref = DocEdit::paragraph_text_ref_from(&crdt, body_index).expect("text ref");
        let txn = crdt.doc().transact();
        let diffs = text_ref.diff(&txn, yrs::types::text::YChange::identity);
        let image_ref = diffs.iter().find_map(|diff| {
            let attrs = diff.attributes.as_ref()?;
            attrs.get("imageRef" as &str).and_then(|value| match value {
                Any::String(value) => Some(value.to_string()),
                _ => None,
            })
        });
        drop(txn);
        editor.crdt.set(Some(crdt));
        image_ref
    }

    #[test]
    fn test_encode_decode_position() {
        let encoded = encode_position(3, 42);
        let (body_idx, offset) = decode_position(&encoded).expect("decode");
        assert_eq!(body_idx, 3);
        assert_eq!(offset, 42);
    }

    #[test]
    fn test_encode_decode_position_zero() {
        let encoded = encode_position(0, 0);
        let (body_idx, offset) = decode_position(&encoded).expect("decode");
        assert_eq!(body_idx, 0);
        assert_eq!(offset, 0);
    }

    #[test]
    fn test_encode_decode_position_max() {
        let encoded = encode_position(u32::MAX, u32::MAX);
        let (body_idx, offset) = decode_position(&encoded).expect("decode");
        assert_eq!(body_idx, u32::MAX);
        assert_eq!(offset, u32::MAX);
    }

    #[test]
    fn test_decode_position_too_short() {
        use base64::Engine;
        let engine = base64::engine::general_purpose::STANDARD;
        let short = engine.encode([0u8; 4]); // Only 4 bytes, need 8
        assert!(decode_position(&short).is_err());
    }

    #[test]
    fn test_new_and_view_model() {
        let editor = make_test_editor("Hello world");
        let crdt = editor.crdt.take().expect("crdt");
        let vm = view::crdt_to_view_model(&crdt, &editor.original_doc).expect("view model");
        editor.crdt.set(Some(crdt));
        assert_eq!(vm.body.len(), 1);
        match &vm.body[0] {
            crate::model::BodyItem::Paragraph(p) => {
                assert_eq!(p.runs.len(), 1);
                assert_eq!(p.runs[0].text, "Hello world");
            }
            _ => panic!("expected paragraph"),
        }
    }

    #[test]
    fn test_insert_text() {
        let editor = make_test_editor("Hello");
        let anchor = encode_position(0, 5);
        let intent = EditIntent::InsertText {
            data: " world".to_string(),
            anchor,
            attrs: None,
        };
        apply(&editor, &intent);
        assert_eq!(para_text(&editor), "Hello world");
        assert!(editor.is_dirty());
    }

    #[test]
    fn test_insert_text_utf16_offset_after_astral() {
        let editor = make_test_editor("😀b");
        // Browser DOM offsets are UTF-16 code units. Offset 3 is the end
        // of "😀b" (2 code units for 😀 + 1 for b).
        let anchor = encode_position(0, 3);
        let intent = EditIntent::InsertText {
            data: "!".to_string(),
            anchor,
            attrs: None,
        };
        apply(&editor, &intent);
        assert_eq!(para_text(&editor), "😀b!");
    }

    #[test]
    fn test_insert_text_strips_sentinels() {
        let editor = make_test_editor("Hello");
        let anchor = encode_position(0, 5);
        let data = format!(" w{}rld", tokens::SENTINEL);
        let intent = EditIntent::InsertText {
            data,
            anchor,
            attrs: None,
        };
        apply(&editor, &intent);
        assert_eq!(para_text(&editor), "Hello wrld");
    }

    #[test]
    fn test_delete_backward() {
        let editor = make_test_editor("Hello");
        let pos = encode_position(0, 5);
        let intent = EditIntent::DeleteBackward {
            anchor: pos.clone(),
            focus: pos,
        };
        apply(&editor, &intent);
        assert_eq!(para_text(&editor), "Hell");
    }

    #[test]
    fn test_delete_backward_at_start() {
        let editor = make_test_editor("Hello");
        let pos = encode_position(0, 0);
        let intent = EditIntent::DeleteBackward {
            anchor: pos.clone(),
            focus: pos,
        };
        apply(&editor, &intent);
        assert_eq!(para_text(&editor), "Hello");
    }

    #[test]
    fn test_delete_forward() {
        let editor = make_test_editor("Hello");
        let pos = encode_position(0, 0);
        let intent = EditIntent::DeleteForward {
            anchor: pos.clone(),
            focus: pos,
        };
        apply(&editor, &intent);
        assert_eq!(para_text(&editor), "ello");
    }

    #[test]
    fn test_range_delete() {
        let editor = make_test_editor("Hello world");
        let anchor = encode_position(0, 0);
        let focus = encode_position(0, 6);
        let intent = EditIntent::DeleteBackward { anchor, focus };
        apply(&editor, &intent);
        assert_eq!(para_text(&editor), "world");
    }

    #[test]
    fn test_multiline_range_delete_merges_paragraphs() {
        let editor = make_test_editor_with_paragraphs(&["alpha", "bravo", "charlie"]);
        let anchor = encode_position(0, 2);
        let focus = encode_position(2, 3);
        let intent = EditIntent::DeleteBackward { anchor, focus };
        apply(&editor, &intent);

        assert_eq!(body_len(&editor), 1);
        assert_eq!(para_text_at(&editor, 0), "alrlie");
    }

    #[test]
    fn test_multiline_delete_by_cut_merges_paragraphs() {
        let editor = make_test_editor_with_paragraphs(&["abcd", "efgh", "ijkl"]);
        let anchor = encode_position(0, 1);
        let focus = encode_position(2, 2);
        let intent = EditIntent::DeleteByCut { anchor, focus };
        apply(&editor, &intent);

        assert_eq!(body_len(&editor), 1);
        assert_eq!(para_text_at(&editor, 0), "akl");
    }

    #[test]
    fn test_multiline_paste_replaces_full_selection() {
        let editor = make_test_editor_with_paragraphs(&["abcd", "efgh", "ijkl"]);
        let anchor = encode_position(0, 1);
        let focus = encode_position(2, 2);
        let intent = EditIntent::InsertFromPaste {
            data: "ZZ".to_string(),
            anchor,
            focus,
            attrs: None,
        };
        apply(&editor, &intent);

        assert_eq!(body_len(&editor), 1);
        assert_eq!(para_text_at(&editor, 0), "aZZkl");
    }

    #[test]
    fn test_sentinel_protection() {
        let mut doc = offidized_docx::Document::new();
        let para = doc.add_paragraph("");
        if let Some(run) = para.runs_mut().first_mut() {
            run.set_has_tab(true);
        }
        let bytes = doc.to_bytes().expect("to_bytes");
        let doc2 = offidized_docx::Document::from_bytes(&bytes).expect("from_bytes");
        let mut crdt = CrdtDoc::new();
        import::import_document(&doc2, &mut crdt).expect("import");
        let editor = DocEdit {
            crdt: Cell::new(Some(crdt)),
            original_doc: doc2,
            original_bytes: bytes,
            undo_mgr: Cell::new(None),
        };

        let pos = encode_position(0, 1);
        let intent = EditIntent::DeleteBackward {
            anchor: pos.clone(),
            focus: pos,
        };
        apply(&editor, &intent);
        assert_eq!(para_text(&editor), tokens::SENTINEL.to_string());
    }

    #[test]
    fn test_sentinel_protection_forward() {
        let mut doc = offidized_docx::Document::new();
        let para = doc.add_paragraph("");
        if let Some(run) = para.runs_mut().first_mut() {
            run.set_has_tab(true);
        }
        let bytes = doc.to_bytes().expect("to_bytes");
        let doc2 = offidized_docx::Document::from_bytes(&bytes).expect("from_bytes");
        let mut crdt = CrdtDoc::new();
        import::import_document(&doc2, &mut crdt).expect("import");
        let editor = DocEdit {
            crdt: Cell::new(Some(crdt)),
            original_doc: doc2,
            original_bytes: bytes,
            undo_mgr: Cell::new(None),
        };

        let pos = encode_position(0, 0);
        let intent = EditIntent::DeleteForward {
            anchor: pos.clone(),
            focus: pos,
        };
        apply(&editor, &intent);
        assert_eq!(para_text(&editor), tokens::SENTINEL.to_string());
    }

    #[test]
    fn test_sentinel_range_protection() {
        let mut doc = offidized_docx::Document::new();
        let para = doc.add_paragraph("AB");
        if let Some(run) = para.runs_mut().first_mut() {
            run.set_has_tab(true);
        }
        let bytes = doc.to_bytes().expect("to_bytes");
        let doc2 = offidized_docx::Document::from_bytes(&bytes).expect("from_bytes");
        let mut crdt = CrdtDoc::new();
        import::import_document(&doc2, &mut crdt).expect("import");
        let editor = DocEdit {
            crdt: Cell::new(Some(crdt)),
            original_doc: doc2,
            original_bytes: bytes,
            undo_mgr: Cell::new(None),
        };

        let anchor = encode_position(0, 0);
        let focus = encode_position(0, 3);
        let intent = EditIntent::DeleteByCut { anchor, focus };
        apply(&editor, &intent);
        let content = para_text(&editor);
        assert!(content.contains(tokens::SENTINEL));
    }

    #[test]
    fn test_insert_inline_image_resolves_local_view_image() {
        let editor = make_test_editor("A");
        let anchor = encode_position(0, 1);
        editor
            .insert_inline_image(
                &anchor,
                &anchor,
                "data:image/png;base64,AQID",
                24.0,
                12.0,
                Some("Logo".to_string()),
                Some("Product logo".to_string()),
            )
            .expect("insert inline image");

        assert_eq!(para_text(&editor), format!("A{}", tokens::SENTINEL));
        let image_ref = inline_image_ref_at(&editor, 0).expect("image ref");
        assert!(image_ref.starts_with("img:local:"));

        let vm = view_model(&editor);
        assert_eq!(vm.images.len(), 1);
        let crate::model::BodyItem::Paragraph(paragraph) = &vm.body[0] else {
            panic!("expected paragraph");
        };
        let image_run = paragraph
            .runs
            .iter()
            .find_map(|run| run.inline_image.as_ref())
            .expect("inline image run");
        assert_eq!(image_run.image_index, 0);
        assert_eq!(image_run.width_pt, 24.0);
        assert_eq!(image_run.height_pt, 12.0);
        assert_eq!(image_run.name.as_deref(), Some("Logo"));
        assert_eq!(image_run.description.as_deref(), Some("Product logo"));
        assert_eq!(vm.images[0].content_type, "image/png");
        assert_eq!(vm.images[0].data_uri, "data:image/png;base64,AQID");
    }

    #[test]
    fn test_inline_image_delete_backward_removes_token() {
        let editor = make_test_editor("");
        let anchor = encode_position(0, 0);
        editor
            .insert_inline_image(
                &anchor,
                &anchor,
                "data:image/png;base64,AQID",
                20.0,
                10.0,
                None,
                None,
            )
            .expect("insert inline image");

        apply(
            &editor,
            &EditIntent::DeleteBackward {
                anchor: encode_position(0, 1),
                focus: encode_position(0, 1),
            },
        );

        assert_eq!(para_text(&editor), "");
    }

    #[test]
    fn test_inline_image_delete_forward_removes_token() {
        let editor = make_test_editor("");
        let anchor = encode_position(0, 0);
        editor
            .insert_inline_image(
                &anchor,
                &anchor,
                "data:image/png;base64,AQID",
                20.0,
                10.0,
                None,
                None,
            )
            .expect("insert inline image");

        apply(
            &editor,
            &EditIntent::DeleteForward {
                anchor: encode_position(0, 0),
                focus: encode_position(0, 0),
            },
        );

        assert_eq!(para_text(&editor), "");
    }

    #[test]
    fn test_inline_image_range_delete_removes_token() {
        let editor = make_test_editor("AB");
        let anchor = encode_position(0, 1);
        editor
            .insert_inline_image(
                &anchor,
                &anchor,
                "data:image/png;base64,AQID",
                20.0,
                10.0,
                None,
                None,
            )
            .expect("insert inline image");

        apply(
            &editor,
            &EditIntent::DeleteByCut {
                anchor: encode_position(0, 1),
                focus: encode_position(0, 2),
            },
        );

        assert_eq!(para_text(&editor), "AB");
    }

    #[test]
    fn test_inline_image_save_reload_roundtrip() {
        let editor = make_test_editor("A");
        let anchor = encode_position(0, 1);
        editor
            .insert_inline_image(
                &anchor,
                &anchor,
                "data:image/png;base64,AQID",
                24.0,
                12.0,
                Some("Logo".to_string()),
                Some("Product logo".to_string()),
            )
            .expect("insert inline image");

        let bytes = editor.save().expect("save");
        let reopened = DocEdit::new(&bytes).expect("reopen editor");
        let vm = view_model(&reopened);
        assert_eq!(vm.images.len(), 1);
        let crate::model::BodyItem::Paragraph(paragraph) = &vm.body[0] else {
            panic!("expected paragraph");
        };
        let image_run = paragraph
            .runs
            .iter()
            .find_map(|run| run.inline_image.as_ref())
            .expect("inline image run");
        assert_eq!(image_run.image_index, 0);
        assert_eq!(image_run.width_pt, 24.0);
        assert_eq!(image_run.height_pt, 12.0);
        assert_eq!(image_run.name.as_deref(), Some("Logo"));
        assert_eq!(image_run.description.as_deref(), Some("Product logo"));
        assert_eq!(vm.images[0].data_uri, "data:image/png;base64,AQID");
    }

    #[test]
    fn test_format_bold() {
        let editor = make_test_editor("Hello world");
        let anchor = encode_position(0, 0);
        let focus = encode_position(0, 5);
        let intent = EditIntent::FormatBold { anchor, focus };
        apply(&editor, &intent);

        let crdt = editor.crdt.take().expect("crdt");
        let txn = crdt.doc().transact();
        let body = crdt.body();
        let entry = body.get(&txn, 0).expect("body[0]");
        let map_ref = entry.cast::<yrs::MapRef>().expect("is map");
        let text_out = map_ref.get(&txn, "text").expect("has text");
        let text_ref = text_out.cast::<yrs::TextRef>().expect("is TextRef");
        let diffs = text_ref.diff(&txn, yrs::types::text::YChange::identity);

        assert!(
            diffs.len() >= 2,
            "expected multiple diff chunks, got {}",
            diffs.len()
        );

        let first_attrs = diffs[0].attributes.as_ref().expect("first has attrs");
        assert_eq!(first_attrs.get("bold" as &str), Some(&Any::Bool(true)));
        drop(txn);
        editor.crdt.set(Some(crdt));
        assert!(editor.is_dirty());
    }

    #[test]
    fn test_format_italic() {
        let editor = make_test_editor("Hello world");
        let anchor = encode_position(0, 6);
        let focus = encode_position(0, 11);
        let intent = EditIntent::FormatItalic { anchor, focus };
        apply(&editor, &intent);

        let crdt = editor.crdt.take().expect("crdt");
        let txn = crdt.doc().transact();
        let body = crdt.body();
        let entry = body.get(&txn, 0).expect("body[0]");
        let map_ref = entry.cast::<yrs::MapRef>().expect("is map");
        let text_out = map_ref.get(&txn, "text").expect("has text");
        let text_ref = text_out.cast::<yrs::TextRef>().expect("is TextRef");
        let diffs = text_ref.diff(&txn, yrs::types::text::YChange::identity);

        let world_chunk = diffs
            .iter()
            .find(|d| matches!(&d.insert, Out::Any(Any::String(s)) if s.as_ref() == "world"));
        assert!(world_chunk.is_some(), "should have a 'world' chunk");
        let attrs = world_chunk
            .and_then(|c| c.attributes.as_ref())
            .expect("world chunk has attrs");
        assert_eq!(attrs.get("italic" as &str), Some(&Any::Bool(true)));
        drop(txn);
        editor.crdt.set(Some(crdt));
    }

    #[test]
    fn test_format_empty_range_noop() {
        let editor = make_test_editor("Hello");
        let pos = encode_position(0, 3);
        let intent = EditIntent::FormatBold {
            anchor: pos.clone(),
            focus: pos,
        };
        apply(&editor, &intent);
        assert!(!editor.is_dirty());
    }

    #[test]
    fn test_insert_text_accepts_rich_attrs() {
        let editor = make_test_editor("");
        let intent = EditIntent::InsertText {
            data: "Hello".to_string(),
            anchor: encode_position(0, 0),
            attrs: Some(HashMap::from([
                (
                    "fontFamily".to_string(),
                    IntentAttrValue::String("Comic Sans MS".to_string()),
                ),
                ("fontSizePt".to_string(), IntentAttrValue::Number(14.0)),
                (
                    "color".to_string(),
                    IntentAttrValue::String("#FF5500".to_string()),
                ),
            ])),
        };
        apply(&editor, &intent);

        let crdt = editor.crdt.take().expect("crdt");
        let vm = view::crdt_to_view_model(&crdt, &editor.original_doc).expect("view model");
        editor.crdt.set(Some(crdt));
        let crate::model::BodyItem::Paragraph(paragraph) = &vm.body[0] else {
            panic!("expected paragraph");
        };
        assert_eq!(
            paragraph.runs[0].font_family.as_deref(),
            Some("Comic Sans MS")
        );
        assert_eq!(paragraph.runs[0].font_size_pt, Some(14.0));
        assert_eq!(paragraph.runs[0].color.as_deref(), Some("FF5500"));
    }

    #[test]
    fn test_set_text_attrs_updates_formatting_at_response() {
        let editor = make_test_editor("Hello");
        apply(
            &editor,
            &EditIntent::SetTextAttrs {
                anchor: encode_position(0, 0),
                focus: encode_position(0, 5),
                attrs: HashMap::from([
                    (
                        "fontFamily".to_string(),
                        IntentAttrValue::String("Courier New".to_string()),
                    ),
                    ("fontSizePt".to_string(), IntentAttrValue::Number(12.0)),
                    (
                        "color".to_string(),
                        IntentAttrValue::String("00AAFF".to_string()),
                    ),
                ]),
            },
        );

        let fmt = formatting_map(&editor, 0, 3);
        assert_eq!(
            fmt.get("fontFamily").and_then(|value| match value {
                Any::String(value) => Some(value.as_ref()),
                _ => None,
            }),
            Some("Courier New")
        );
        assert_eq!(
            fmt.get("fontSize").and_then(|value| match value {
                Any::Number(value) => Some(*value / 2.0),
                _ => None,
            }),
            Some(12.0)
        );
        assert_eq!(
            fmt.get("color").and_then(|value| match value {
                Any::String(value) => Some(value.as_ref()),
                _ => None,
            }),
            Some("00AAFF")
        );
    }

    #[test]
    fn test_set_text_attrs_null_removes_attribute() {
        let editor = make_test_editor("Hello");
        apply(
            &editor,
            &EditIntent::SetTextAttrs {
                anchor: encode_position(0, 0),
                focus: encode_position(0, 5),
                attrs: HashMap::from([(
                    "color".to_string(),
                    IntentAttrValue::String("112233".to_string()),
                )]),
            },
        );
        apply(
            &editor,
            &EditIntent::SetTextAttrs {
                anchor: encode_position(0, 0),
                focus: encode_position(0, 5),
                attrs: HashMap::from([("color".to_string(), IntentAttrValue::Null)]),
            },
        );

        let fmt = formatting_map(&editor, 0, 2);
        assert!(!fmt.contains_key("color"));
    }

    #[test]
    fn test_set_paragraph_attrs_updates_heading_level() {
        let editor = make_test_editor("Hello");
        apply(
            &editor,
            &EditIntent::SetParagraphAttrs {
                anchor: encode_position(0, 0),
                focus: encode_position(0, 0),
                attrs: HashMap::from([("headingLevel".to_string(), IntentAttrValue::Number(2.0))]),
            },
        );

        assert_eq!(paragraph_heading_level(&editor, 0), Some(2.0));
        let fmt = formatting_map(&editor, 0, 1);
        assert_eq!(
            fmt.get("headingLevel").and_then(|value| match value {
                Any::Number(value) => Some(*value),
                _ => None,
            }),
            Some(2.0)
        );

        apply(
            &editor,
            &EditIntent::SetParagraphAttrs {
                anchor: encode_position(0, 0),
                focus: encode_position(0, 0),
                attrs: HashMap::from([("headingLevel".to_string(), IntentAttrValue::Null)]),
            },
        );
        assert_eq!(paragraph_heading_level(&editor, 0), None);
    }

    #[test]
    fn test_set_paragraph_attrs_updates_alignment() {
        let editor = make_test_editor("Hello");
        apply(
            &editor,
            &EditIntent::SetParagraphAttrs {
                anchor: encode_position(0, 0),
                focus: encode_position(0, 0),
                attrs: HashMap::from([(
                    "alignment".to_string(),
                    IntentAttrValue::String("center".to_string()),
                )]),
            },
        );

        assert_eq!(paragraph_alignment(&editor, 0), Some("center".to_string()));
        let fmt = formatting_map(&editor, 0, 1);
        assert_eq!(
            fmt.get("alignment").and_then(|value| match value {
                Any::String(value) => Some(value.to_string()),
                _ => None,
            }),
            Some("center".to_string())
        );

        apply(
            &editor,
            &EditIntent::SetParagraphAttrs {
                anchor: encode_position(0, 0),
                focus: encode_position(0, 0),
                attrs: HashMap::from([("alignment".to_string(), IntentAttrValue::Null)]),
            },
        );
        assert_eq!(paragraph_alignment(&editor, 0), None);
    }

    #[test]
    fn test_sync_roundtrip_preserves_text_and_paragraph_formatting() {
        let editor1 = make_test_editor("Hello");
        apply(
            &editor1,
            &EditIntent::SetTextAttrs {
                anchor: encode_position(0, 0),
                focus: encode_position(0, 5),
                attrs: HashMap::from([
                    (
                        "fontFamily".to_string(),
                        IntentAttrValue::String("PT Serif".to_string()),
                    ),
                    (
                        "color".to_string(),
                        IntentAttrValue::String("00AAFF".to_string()),
                    ),
                ]),
            },
        );
        apply(
            &editor1,
            &EditIntent::SetParagraphAttrs {
                anchor: encode_position(0, 0),
                focus: encode_position(0, 5),
                attrs: HashMap::from([
                    ("headingLevel".to_string(), IntentAttrValue::Number(2.0)),
                    (
                        "alignment".to_string(),
                        IntentAttrValue::String("right".to_string()),
                    ),
                ]),
            },
        );

        let editor2 = make_test_editor("");
        let full_state = editor1.encode_state_as_update().expect("full state");
        editor2.apply_update(&full_state).expect("apply full state");

        let crdt = editor2.crdt.take().expect("crdt");
        let vm = view::crdt_to_view_model(&crdt, &editor2.original_doc).expect("view model");
        editor2.crdt.set(Some(crdt));
        let paragraph = vm
            .body
            .iter()
            .find_map(|item| match item {
                crate::model::BodyItem::Paragraph(paragraph)
                    if paragraph.runs.iter().any(|run| run.text == "Hello") =>
                {
                    Some(paragraph)
                }
                _ => None,
            })
            .expect("paragraph with Hello");
        assert_eq!(paragraph.heading_level, Some(2));
        assert_eq!(paragraph.alignment.as_deref(), Some("right"));
        let hello_run = paragraph
            .runs
            .iter()
            .find(|run| run.text == "Hello")
            .expect("hello run");
        assert_eq!(hello_run.font_family.as_deref(), Some("PT Serif"));
        assert_eq!(hello_run.color.as_deref(), Some("00AAFF"));
    }

    #[test]
    fn test_insert_paragraph_preserves_heading_level_on_new_paragraph() {
        let editor = make_test_editor("Heading");
        apply(
            &editor,
            &EditIntent::SetParagraphAttrs {
                anchor: encode_position(0, 0),
                focus: encode_position(0, 7),
                attrs: HashMap::from([("headingLevel".to_string(), IntentAttrValue::Number(2.0))]),
            },
        );
        apply(
            &editor,
            &EditIntent::InsertParagraph {
                anchor: encode_position(0, 7),
            },
        );

        assert_eq!(body_len(&editor), 2);
        assert_eq!(paragraph_heading_level(&editor, 0), Some(2.0));
        assert_eq!(paragraph_heading_level(&editor, 1), Some(2.0));
    }

    #[test]
    fn test_set_paragraph_attrs_updates_numbering_formatting_state() {
        let editor = make_test_editor("Item");
        apply(
            &editor,
            &EditIntent::SetParagraphAttrs {
                anchor: encode_position(0, 0),
                focus: encode_position(0, 4),
                attrs: HashMap::from([
                    (
                        "numberingKind".to_string(),
                        IntentAttrValue::String("bullet".to_string()),
                    ),
                    ("numberingNumId".to_string(), IntentAttrValue::Number(7.0)),
                    ("numberingIlvl".to_string(), IntentAttrValue::Number(0.0)),
                ]),
            },
        );

        assert_eq!(
            paragraph_numbering(&editor, 0),
            (Some("bullet".to_string()), Some(7.0), Some(0.0))
        );
        let fmt = formatting_map(&editor, 0, 1);
        assert_eq!(
            fmt.get("numberingKind").and_then(|value| match value {
                Any::String(value) => Some(value.to_string()),
                _ => None,
            }),
            Some("bullet".to_string())
        );
    }

    #[test]
    fn test_insert_paragraph_preserves_numbering_on_new_list_item() {
        let editor = make_test_editor("Item");
        apply(
            &editor,
            &EditIntent::SetParagraphAttrs {
                anchor: encode_position(0, 0),
                focus: encode_position(0, 4),
                attrs: HashMap::from([
                    (
                        "numberingKind".to_string(),
                        IntentAttrValue::String("decimal".to_string()),
                    ),
                    ("numberingNumId".to_string(), IntentAttrValue::Number(3.0)),
                    ("numberingIlvl".to_string(), IntentAttrValue::Number(0.0)),
                ]),
            },
        );
        apply(
            &editor,
            &EditIntent::InsertParagraph {
                anchor: encode_position(0, 4),
            },
        );

        assert_eq!(
            paragraph_numbering(&editor, 0),
            (Some("decimal".to_string()), Some(3.0), Some(0.0))
        );
        assert_eq!(
            paragraph_numbering(&editor, 1),
            (Some("decimal".to_string()), Some(3.0), Some(0.0))
        );
    }

    #[test]
    fn test_insert_line_break() {
        let editor = make_test_editor("Hello");
        let anchor = encode_position(0, 5);
        let intent = EditIntent::InsertLineBreak { anchor };
        apply(&editor, &intent);

        let content = para_text(&editor);
        assert_eq!(content.len(), "Hello".len() + tokens::SENTINEL.len_utf8());
        assert!(content.contains(tokens::SENTINEL));
    }

    #[test]
    fn test_paste_with_selection() {
        let editor = make_test_editor("Hello world");
        let anchor = encode_position(0, 6);
        let focus = encode_position(0, 11);
        let intent = EditIntent::InsertFromPaste {
            data: "there".to_string(),
            anchor,
            focus,
            attrs: None,
        };
        apply(&editor, &intent);
        assert_eq!(para_text(&editor), "Hello there");
    }

    #[test]
    fn test_insert_from_paste_accepts_attrs() {
        let editor = make_test_editor("Hello world");
        let anchor = encode_position(0, 6);
        let focus = encode_position(0, 11);
        let intent = EditIntent::InsertFromPaste {
            data: "there".to_string(),
            anchor,
            focus,
            attrs: Some(HashMap::from([
                (
                    "fontFamily".to_string(),
                    IntentAttrValue::String("Courier New".to_string()),
                ),
                ("fontSizePt".to_string(), IntentAttrValue::Number(13.0)),
            ])),
        };
        apply(&editor, &intent);

        let crdt = editor.crdt.take().expect("crdt");
        let vm = view::crdt_to_view_model(&crdt, &editor.original_doc).expect("view model");
        editor.crdt.set(Some(crdt));
        let crate::model::BodyItem::Paragraph(paragraph) = &vm.body[0] else {
            panic!("expected paragraph");
        };
        let replaced_run = paragraph
            .runs
            .iter()
            .find(|run| run.text == "there")
            .expect("replaced run");
        assert_eq!(replaced_run.font_family.as_deref(), Some("Courier New"));
        assert_eq!(replaced_run.font_size_pt, Some(13.0));
    }

    #[test]
    fn test_undo_redo_noop() {
        let editor = make_test_editor("Hello");
        apply(&editor, &EditIntent::Undo);
        apply(&editor, &EditIntent::Redo);
        assert_eq!(para_text(&editor), "Hello");
    }

    #[test]
    fn test_save_roundtrip() {
        let editor = make_test_editor("Hello");
        let anchor = encode_position(0, 5);
        let intent = EditIntent::InsertText {
            data: " world".to_string(),
            anchor,
            attrs: None,
        };
        apply(&editor, &intent);

        let mut crdt = editor.crdt.take().expect("crdt");
        let bytes = export::export_to_docx(&crdt, &editor.original_doc).expect("export");
        crdt.clear_dirty();
        editor.crdt.set(Some(crdt));

        let reopened = offidized_docx::Document::from_bytes(&bytes).expect("reopen");
        assert_eq!(reopened.paragraphs().len(), 1);
        assert_eq!(reopened.paragraphs()[0].text(), "Hello world");
    }

    #[test]
    fn test_is_dirty_after_edit() {
        let editor = make_test_editor("Hello");
        assert!(!editor.is_dirty());

        let anchor = encode_position(0, 5);
        let intent = EditIntent::InsertText {
            data: "!".to_string(),
            anchor,
            attrs: None,
        };
        apply(&editor, &intent);
        assert!(editor.is_dirty());
    }

    #[test]
    fn test_body_length() {
        let editor = make_test_editor("Hello");
        assert_eq!(editor.body_length(), 1);
    }

    #[test]
    fn test_encode_position_js() {
        let encoded = DocEdit::encode_position_js(7, 99);
        let (bi, off) = decode_position(&encoded).expect("decode");
        assert_eq!(bi, 7);
        assert_eq!(off, 99);
    }

    /// Regression: check that get_string().chars().count() matches the
    /// offset that yrs actually accepts for insert/insert_with_attributes.
    #[test]
    fn test_yrs_len_vs_chars_count() {
        let editor = make_test_editor("Hello");

        // After initial import, check consistency
        let crdt = editor.crdt.take().expect("crdt");
        let text_ref = DocEdit::paragraph_text_ref_from(&crdt, 0).expect("text_ref");
        let txn = crdt.doc().transact();
        let str_len = text_ref.get_string(&txn).chars().count() as u32;
        let yrs_len = text_ref.len(&txn);
        drop(txn);
        assert_eq!(
            str_len, yrs_len,
            "chars count {str_len} vs yrs len {yrs_len} after import"
        );
        editor.crdt.set(Some(crdt));

        // Insert a line break (sentinel with attributes) at offset 5
        apply(
            &editor,
            &EditIntent::InsertLineBreak {
                anchor: encode_position(0, 5),
            },
        );

        // Check again — does chars().count() still match yrs len?
        let crdt = editor.crdt.take().expect("crdt");
        let text_ref = DocEdit::paragraph_text_ref_from(&crdt, 0).expect("text_ref");
        let txn = crdt.doc().transact();
        let str_after = text_ref.get_string(&txn);
        let str_len = str_after.chars().count() as u32;
        let yrs_len = text_ref.len(&txn);
        drop(txn);
        eprintln!(
            "After line break: string={:?}, chars={str_len}, yrs_len={yrs_len}",
            str_after
        );
        assert_eq!(
            str_len, yrs_len,
            "chars count {str_len} vs yrs len {yrs_len} after line break"
        );
        editor.crdt.set(Some(crdt));
    }

    /// Regression: insert text → line break → more text at various offsets.
    /// This exercises the offset clamping for insert_with_attributes.
    #[test]
    fn test_insert_after_line_break() {
        let editor = make_test_editor("");

        // Type "Hello" into the blank paragraph
        apply(
            &editor,
            &EditIntent::InsertText {
                data: "Hello".to_string(),
                anchor: encode_position(0, 0),
                attrs: None,
            },
        );
        assert_eq!(para_text(&editor), "Hello");

        // Press Enter at the end
        apply(
            &editor,
            &EditIntent::InsertLineBreak {
                anchor: encode_position(0, 5),
            },
        );

        // Now type after the line break.
        // The text is "Hello" + SENTINEL, so 6 chars. Offset 6 = after sentinel.
        apply(
            &editor,
            &EditIntent::InsertText {
                data: "World".to_string(),
                anchor: encode_position(0, 6),
                attrs: None,
            },
        );

        let content = para_text(&editor);
        assert!(
            content.contains("Hello"),
            "should still contain Hello, got: {content:?}"
        );
        assert!(
            content.contains("World"),
            "should contain World, got: {content:?}"
        );
    }

    /// Regression: type "Hi", Enter, then type "X" after the sentinel.
    /// Mirrors the exact browser flow that was failing.
    #[test]
    fn test_type_enter_type_short() {
        let editor = make_test_editor("");

        // Type "Hi" char by char
        apply(
            &editor,
            &EditIntent::InsertText {
                data: "H".to_string(),
                anchor: encode_position(0, 0),
                attrs: None,
            },
        );
        apply(
            &editor,
            &EditIntent::InsertText {
                data: "i".to_string(),
                anchor: encode_position(0, 1),
                attrs: None,
            },
        );
        assert_eq!(para_text(&editor), "Hi");

        // Check yrs len
        let crdt = editor.crdt.take().expect("crdt");
        let text_ref = DocEdit::paragraph_text_ref_from(&crdt, 0).expect("text_ref");
        let txn = crdt.doc().transact();
        let len_before = text_ref.len(&txn);
        drop(txn);
        editor.crdt.set(Some(crdt));
        eprintln!(
            "Before Enter: text={:?}, yrs_len={len_before}",
            para_text(&editor)
        );
        assert_eq!(len_before, 2);

        // Press Enter at offset 2
        apply(
            &editor,
            &EditIntent::InsertLineBreak {
                anchor: encode_position(0, 2),
            },
        );

        // Check yrs len after
        let crdt = editor.crdt.take().expect("crdt");
        let text_ref = DocEdit::paragraph_text_ref_from(&crdt, 0).expect("text_ref");
        let txn = crdt.doc().transact();
        let raw = text_ref.get_string(&txn);
        let len_after = text_ref.len(&txn);
        drop(txn);
        editor.crdt.set(Some(crdt));
        eprintln!(
            "After Enter: raw={:?}, raw_chars={}, yrs_len={len_after}",
            raw,
            raw.chars().count()
        );
        assert_eq!(len_after, 3, "Hi(2) + sentinel(1) = 3");

        // Now type "X" at offset 3 — right after the sentinel
        apply(
            &editor,
            &EditIntent::InsertText {
                data: "X".to_string(),
                anchor: encode_position(0, 3),
                attrs: None,
            },
        );

        let content = para_text(&editor);
        eprintln!(
            "After X: raw={:?}, chars={}",
            content,
            content.chars().count()
        );
        assert!(
            content.contains('X'),
            "should contain 'X' after typing past sentinel, got: {content:?}"
        );

        // Check raw diff output from yrs
        let crdt = editor.crdt.take().expect("crdt");
        let text_ref = DocEdit::paragraph_text_ref_from(&crdt, 0).expect("text_ref");
        let txn = crdt.doc().transact();
        let diffs = text_ref.diff(&txn, yrs::types::text::YChange::identity);
        eprintln!("Diff chunks: {}", diffs.len());
        for (i, d) in diffs.iter().enumerate() {
            eprintln!(
                "  chunk[{i}]: insert={:?}, attrs={:?}",
                d.insert, d.attributes
            );
        }
        drop(txn);

        // Also check via view model
        let vm = view::crdt_to_view_model(&crdt, &editor.original_doc).expect("view model");
        editor.crdt.set(Some(crdt));
        let runs: Vec<_> = match &vm.body[0] {
            crate::model::BodyItem::Paragraph(p) => p
                .runs
                .iter()
                .map(|r| format!("text={:?} break={}", r.text, r.has_break))
                .collect(),
            _ => panic!("expected paragraph"),
        };
        eprintln!("View model runs: {runs:?}");
    }

    /// Regression: type into a blank paragraph at offset 0.
    /// Browser can send offset 1 for an empty paragraph (because of <br> placeholder).
    #[test]
    fn test_insert_text_into_empty_at_offset_1() {
        let editor = make_test_editor("");
        // Browser might send offset 1 for an empty paragraph
        apply(
            &editor,
            &EditIntent::InsertText {
                data: "A".to_string(),
                anchor: encode_position(0, 1),
                attrs: None,
            },
        );
        // Should succeed (offset clamped to 0) and text should be "A"
        assert_eq!(para_text(&editor), "A");
    }

    /// Regression: line break into a blank paragraph at offset 1.
    #[test]
    fn test_line_break_into_empty_at_offset_1() {
        let editor = make_test_editor("");
        apply(
            &editor,
            &EditIntent::InsertLineBreak {
                anchor: encode_position(0, 1),
            },
        );
        // Should succeed (offset clamped to 0) and contain sentinel
        let content = para_text(&editor);
        assert!(
            content.contains(tokens::SENTINEL),
            "should contain sentinel, got: {content:?}"
        );
    }

    /// Regression: line break at offset way beyond text length.
    #[test]
    fn test_line_break_at_huge_offset() {
        let editor = make_test_editor("Hi");
        apply(
            &editor,
            &EditIntent::InsertLineBreak {
                anchor: encode_position(0, 999),
            },
        );
        let content = para_text(&editor);
        assert!(
            content.contains(tokens::SENTINEL),
            "should contain sentinel, got: {content:?}"
        );
    }

    /// Regression: simulate exact browser flow — character-by-character typing.
    #[test]
    fn test_browser_char_by_char_typing() {
        let editor = make_test_editor("");
        for (i, ch) in "Hello".chars().enumerate() {
            apply(
                &editor,
                &EditIntent::InsertText {
                    data: ch.to_string(),
                    anchor: encode_position(0, i as u32),
                    attrs: None,
                },
            );
        }
        assert_eq!(para_text(&editor), "Hello");

        // Now press Enter at the end (offset 5)
        apply(
            &editor,
            &EditIntent::InsertLineBreak {
                anchor: encode_position(0, 5),
            },
        );
        let content = para_text(&editor);
        assert!(content.starts_with("Hello"), "got: {content:?}");
        assert!(content.contains(tokens::SENTINEL), "got: {content:?}");
    }

    /// Regression: multi-paragraph document editing.
    #[test]
    fn test_multi_paragraph_editing() {
        let mut doc = offidized_docx::Document::new();
        doc.add_paragraph("First paragraph");
        doc.add_paragraph("Second paragraph");
        doc.add_paragraph("Third paragraph");
        let bytes = doc.to_bytes().expect("to_bytes");
        let doc2 = offidized_docx::Document::from_bytes(&bytes).expect("from_bytes");
        let mut crdt = CrdtDoc::new();
        import::import_document(&doc2, &mut crdt).expect("import");
        let editor = DocEdit {
            crdt: Cell::new(Some(crdt)),
            original_doc: doc2,
            original_bytes: bytes,
            undo_mgr: Cell::new(None),
        };

        // Edit second paragraph
        apply(
            &editor,
            &EditIntent::InsertText {
                data: "!".to_string(),
                anchor: encode_position(1, 16), // end of "Second paragraph"
                attrs: None,
            },
        );

        // Insert line break in first paragraph
        apply(
            &editor,
            &EditIntent::InsertLineBreak {
                anchor: encode_position(0, 15), // end of "First paragraph"
            },
        );
    }

    /// Regression: document with formatting — insert after bold text.
    #[test]
    fn test_insert_after_formatted_text() {
        let mut doc = offidized_docx::Document::new();
        let para = doc.add_paragraph("");
        // Add a bold run then a normal run
        para.add_run("Bold text");
        if let Some(run) = para.runs_mut().last_mut() {
            run.set_bold(true);
        }
        para.add_run(" normal text");
        let bytes = doc.to_bytes().expect("to_bytes");
        let doc2 = offidized_docx::Document::from_bytes(&bytes).expect("from_bytes");
        let mut crdt = CrdtDoc::new();
        import::import_document(&doc2, &mut crdt).expect("import");
        let editor = DocEdit {
            crdt: Cell::new(Some(crdt)),
            original_doc: doc2,
            original_bytes: bytes,
            undo_mgr: Cell::new(None),
        };

        // Check the CRDT text
        let crdt = editor.crdt.take().expect("crdt");
        let text_ref = DocEdit::paragraph_text_ref_from(&crdt, 0).expect("text_ref");
        let txn = crdt.doc().transact();
        let text = text_ref.get_string(&txn);
        let chars = text.chars().count() as u32;
        let yrs_len = text_ref.len(&txn);
        eprintln!(
            "Formatted text: {:?}, chars={chars}, yrs_len={yrs_len}",
            text
        );
        drop(txn);
        editor.crdt.set(Some(crdt));

        // Try inserting at the end (which might be different from chars count)
        apply(
            &editor,
            &EditIntent::InsertText {
                data: "!".to_string(),
                anchor: encode_position(0, yrs_len),
                attrs: None,
            },
        );

        // Try line break at the end
        apply(
            &editor,
            &EditIntent::InsertLineBreak {
                anchor: encode_position(0, yrs_len + 1), // after the "!" we just added
            },
        );
    }

    /// Regression: document with tab sentinels — editing around sentinels.
    #[test]
    fn test_edit_around_tab_sentinels() {
        let mut doc = offidized_docx::Document::new();
        let para = doc.add_paragraph("Before");
        if let Some(run) = para.runs_mut().first_mut() {
            run.set_has_tab(true);
        }
        para.add_run("After");
        let bytes = doc.to_bytes().expect("to_bytes");
        let doc2 = offidized_docx::Document::from_bytes(&bytes).expect("from_bytes");
        let mut crdt = CrdtDoc::new();
        import::import_document(&doc2, &mut crdt).expect("import");
        let editor = DocEdit {
            crdt: Cell::new(Some(crdt)),
            original_doc: doc2,
            original_bytes: bytes,
            undo_mgr: Cell::new(None),
        };

        // Check CRDT state
        let crdt = editor.crdt.take().expect("crdt");
        let text_ref = DocEdit::paragraph_text_ref_from(&crdt, 0).expect("text_ref");
        let txn = crdt.doc().transact();
        let text = text_ref.get_string(&txn);
        let yrs_len = text_ref.len(&txn);
        eprintln!("Tab text: {:?}, yrs_len={yrs_len}", text);
        drop(txn);
        editor.crdt.set(Some(crdt));

        // Insert at end
        apply(
            &editor,
            &EditIntent::InsertText {
                data: "!".to_string(),
                anchor: encode_position(0, yrs_len),
                attrs: None,
            },
        );

        // Line break after everything
        apply(
            &editor,
            &EditIntent::InsertLineBreak {
                anchor: encode_position(0, yrs_len + 1),
            },
        );
    }

    /// Regression: OffsetKind::Utf16 vs chars().count() mismatch.
    ///
    /// The CRDT doc uses OffsetKind::Utf16 (for browser compatibility),
    /// but the import uses chars().count() for position tracking.
    /// For non-BMP characters (emoji), chars().count() gives code points
    /// but yrs expects UTF-16 code units (2 per non-BMP char).
    /// This mismatch corrupts the internal block structure, causing
    /// SplittableString::block_offset overflow on subsequent edits.
    #[test]
    fn test_emoji_offset_mismatch() {
        // Create a document with emoji (non-BMP character)
        let mut doc = offidized_docx::Document::new();
        doc.add_paragraph("Hi 😀 there");
        let bytes = doc.to_bytes().expect("to_bytes");
        let doc2 = offidized_docx::Document::from_bytes(&bytes).expect("from_bytes");
        let mut crdt = CrdtDoc::new();
        import::import_document(&doc2, &mut crdt).expect("import");
        let editor = DocEdit {
            crdt: Cell::new(Some(crdt)),
            original_doc: doc2,
            original_bytes: bytes,
            undo_mgr: Cell::new(None),
        };

        // Check text
        let crdt = editor.crdt.take().expect("crdt");
        let text_ref = DocEdit::paragraph_text_ref_from(&crdt, 0).expect("text_ref");
        let txn = crdt.doc().transact();
        let text = text_ref.get_string(&txn);
        let chars_count = text.chars().count() as u32;
        let utf16_count = text.encode_utf16().count() as u32;
        let yrs_len = text_ref.len(&txn);
        drop(txn);
        eprintln!(
            "Emoji text: {:?}, chars={chars_count}, utf16={utf16_count}, yrs_len={yrs_len}",
            text
        );
        editor.crdt.set(Some(crdt));

        // Try inserting at the end using yrs_len offset
        apply(
            &editor,
            &EditIntent::InsertText {
                data: "!".to_string(),
                anchor: encode_position(0, yrs_len),
                attrs: None,
            },
        );

        // Try line break at the end
        apply(
            &editor,
            &EditIntent::InsertLineBreak {
                anchor: encode_position(0, yrs_len + 1),
            },
        );

        // If we get here without panic, the offsets are consistent
    }

    /// Regression: aggressive editing after importing document with emoji.
    ///
    /// The import uses chars().count() for position tracking, but with
    /// OffsetKind::Utf16, non-BMP characters need 2 code units.
    /// This test inserts/deletes/formats at EVERY possible offset to
    /// find any internal block structure corruption.
    #[test]
    fn test_emoji_aggressive_editing() {
        let mut doc = offidized_docx::Document::new();
        doc.add_paragraph("A 😀 B 🎉 C");
        let bytes = doc.to_bytes().expect("to_bytes");
        let doc2 = offidized_docx::Document::from_bytes(&bytes).expect("from_bytes");
        let mut crdt = CrdtDoc::new();
        import::import_document(&doc2, &mut crdt).expect("import");
        let editor = DocEdit {
            crdt: Cell::new(Some(crdt)),
            original_doc: doc2,
            original_bytes: bytes,
            undo_mgr: Cell::new(None),
        };

        // Check initial state
        let crdt_ref = editor.crdt.take().expect("crdt");
        let text_ref = DocEdit::paragraph_text_ref_from(&crdt_ref, 0).expect("text_ref");
        let txn = crdt_ref.doc().transact();
        let initial_len = text_ref.len(&txn);
        drop(txn);
        editor.crdt.set(Some(crdt_ref));
        eprintln!("Initial yrs_len={initial_len}");

        // Try inserting "x" at every offset from 0 to initial_len + 5
        for offset in 0..=(initial_len + 5) {
            let editor2 = {
                // Clone the initial state for each test
                let mut doc = offidized_docx::Document::new();
                doc.add_paragraph("A 😀 B 🎉 C");
                let bytes = doc.to_bytes().expect("to_bytes");
                let doc2 = offidized_docx::Document::from_bytes(&bytes).expect("from_bytes");
                let mut crdt = CrdtDoc::new();
                import::import_document(&doc2, &mut crdt).expect("import");
                DocEdit {
                    crdt: Cell::new(Some(crdt)),
                    original_doc: doc2,
                    original_bytes: bytes,
                    undo_mgr: Cell::new(None),
                }
            };

            // Insert text
            apply(
                &editor2,
                &EditIntent::InsertText {
                    data: "x".to_string(),
                    anchor: encode_position(0, offset),
                    attrs: None,
                },
            );

            // Insert line break
            let crdt_ref = editor2.crdt.take().expect("crdt");
            let text_ref = DocEdit::paragraph_text_ref_from(&crdt_ref, 0).expect("text_ref");
            let txn = crdt_ref.doc().transact();
            let new_len = text_ref.len(&txn);
            drop(txn);
            editor2.crdt.set(Some(crdt_ref));

            apply(
                &editor2,
                &EditIntent::InsertLineBreak {
                    anchor: encode_position(0, new_len),
                },
            );
        }
    }

    /// Regression: rapid interleaved operations mimicking real editing.
    #[test]
    fn test_rapid_interleaved_edit_sequence() {
        let editor = make_test_editor("");

        // Type "Hello"
        for (i, ch) in "Hello".chars().enumerate() {
            apply(
                &editor,
                &EditIntent::InsertText {
                    data: ch.to_string(),
                    anchor: encode_position(0, i as u32),
                    attrs: None,
                },
            );
        }

        // Bold "Hello"
        apply(
            &editor,
            &EditIntent::FormatBold {
                anchor: encode_position(0, 0),
                focus: encode_position(0, 5),
            },
        );

        // Type " World" after bold
        for (i, ch) in " World".chars().enumerate() {
            apply(
                &editor,
                &EditIntent::InsertText {
                    data: ch.to_string(),
                    anchor: encode_position(0, 5 + i as u32),
                    attrs: None,
                },
            );
        }

        // Press Enter at end (offset 11)
        apply(
            &editor,
            &EditIntent::InsertLineBreak {
                anchor: encode_position(0, 11),
            },
        );

        // Type after the line break (offset 12)
        apply(
            &editor,
            &EditIntent::InsertText {
                data: "After".to_string(),
                anchor: encode_position(0, 12),
                attrs: None,
            },
        );

        let content = para_text(&editor);
        assert!(content.contains("Hello"), "got: {content:?}");
        assert!(content.contains("World"), "got: {content:?}");
        assert!(content.contains("After"), "got: {content:?}");
    }

    // -----------------------------------------------------------------------
    // CRDT sync tests
    // -----------------------------------------------------------------------

    #[test]
    fn test_encode_state_vector() {
        let editor = DocEdit::blank().expect("blank");
        let sv = editor.encode_state_vector().expect("encode_state_vector");
        assert!(!sv.is_empty(), "state vector should be non-empty");
    }

    #[test]
    fn test_encode_state_as_update() {
        let editor = DocEdit::blank().expect("blank");
        // Insert some text so the update is meaningful.
        apply(
            &editor,
            &EditIntent::InsertText {
                data: "Hello sync".to_string(),
                anchor: encode_position(0, 0),
                attrs: None,
            },
        );
        let update = editor
            .encode_state_as_update()
            .expect("encode_state_as_update");
        assert!(!update.is_empty(), "full update should be non-empty");
    }

    /// Helper: collect text from all paragraphs in the CRDT body.
    fn all_para_text(editor: &DocEdit) -> String {
        let crdt = editor.crdt.take().expect("crdt");
        let txn = crdt.doc().transact();
        let body = crdt.body();
        let mut texts = Vec::new();
        for i in 0..body.len(&txn) {
            if let Some(entry) = body.get(&txn, i) {
                if let Ok(map_ref) = entry.cast::<yrs::MapRef>() {
                    if let Some(text_out) = map_ref.get(&txn, "text") {
                        if let Ok(text_ref) = text_out.cast::<yrs::TextRef>() {
                            texts.push(text_ref.get_string(&txn));
                        }
                    }
                }
            }
        }
        drop(txn);
        editor.crdt.set(Some(crdt));
        texts.join("|")
    }

    fn find_paragraph_index_containing(editor: &DocEdit, needle: &str) -> u32 {
        let crdt = editor.crdt.take().expect("crdt");
        let txn = crdt.doc().transact();
        let body = crdt.body();
        let mut found = None;
        for i in 0..body.len(&txn) {
            let Some(entry) = body.get(&txn, i) else {
                continue;
            };
            let Ok(map_ref) = entry.cast::<yrs::MapRef>() else {
                continue;
            };
            let Some(text_out) = map_ref.get(&txn, "text") else {
                continue;
            };
            let Ok(text_ref) = text_out.cast::<yrs::TextRef>() else {
                continue;
            };
            if text_ref.get_string(&txn).contains(needle) {
                found = Some(i);
                break;
            }
        }
        drop(txn);
        editor.crdt.set(Some(crdt));
        found.unwrap_or_else(|| panic!("paragraph containing {needle:?} not found"))
    }

    #[test]
    fn test_apply_update_roundtrip() {
        let editor1 = DocEdit::blank().expect("blank editor1");
        let editor2 = DocEdit::blank().expect("blank editor2");

        // Insert text into editor1.
        apply(
            &editor1,
            &EditIntent::InsertText {
                data: "Synced!".to_string(),
                anchor: encode_position(0, 0),
                attrs: None,
            },
        );
        assert_eq!(para_text(&editor1), "Synced!");

        // Encode the full state from editor1 and apply to editor2.
        let update = editor1.encode_state_as_update().expect("encode");
        editor2.apply_update(&update).expect("apply");

        // editor2 should now have the text somewhere in its body
        // (CRDT merge may reorder paragraphs from different clients).
        let text2 = all_para_text(&editor2);
        assert!(
            text2.contains("Synced!"),
            "editor2 body should contain 'Synced!', got: {text2:?}"
        );
    }

    #[test]
    fn test_apply_update_noop_preserves_local_undo_history() {
        let editor1 = make_test_editor("Hello");
        let editor2 = make_test_editor("");

        let full_state = editor1.encode_state_as_update().expect("full state");
        editor2.apply_update(&full_state).expect("apply full state");

        let hello_index = find_paragraph_index_containing(&editor2, "Hello");
        apply_via_api(
            &editor2,
            &EditIntent::InsertText {
                data: " local".to_string(),
                anchor: encode_position(hello_index, 5),
                attrs: None,
            },
        );

        editor2
            .apply_update(&full_state)
            .expect("apply redundant full state");

        let merged = all_para_text(&editor2);
        assert!(
            merged.contains("local"),
            "merged text should still contain the local edit before undo, got: {merged:?}"
        );

        apply_via_api(&editor2, &EditIntent::Undo);

        let after_undo = all_para_text(&editor2);
        assert!(
            !after_undo.contains("local"),
            "redundant remote updates should not invalidate local undo history, got: {after_undo:?}"
        );
    }

    #[test]
    fn test_apply_update_expands_undo_scope_for_remote_paragraphs() {
        let editor1 = make_test_editor("Hello");
        let editor2 = make_test_editor("");

        let full_state = editor1.encode_state_as_update().expect("full state");
        editor2.apply_update(&full_state).expect("apply full state");

        apply(
            &editor1,
            &EditIntent::InsertParagraph {
                anchor: encode_position(0, 5),
            },
        );
        apply(
            &editor1,
            &EditIntent::InsertText {
                data: "Remote paragraph".to_string(),
                anchor: encode_position(1, 0),
                attrs: None,
            },
        );

        let sv2 = editor2.encode_state_vector().expect("state vector");
        let diff = editor1.encode_diff(&sv2).expect("diff");
        editor2.apply_update(&diff).expect("apply diff");

        let remote_index = find_paragraph_index_containing(&editor2, "Remote paragraph");
        apply_via_api(
            &editor2,
            &EditIntent::InsertText {
                data: " local".to_string(),
                anchor: encode_position(remote_index, 16),
                attrs: None,
            },
        );

        apply_via_api(&editor2, &EditIntent::Undo);

        let after_undo = all_para_text(&editor2);
        assert!(
            after_undo.contains("Remote paragraph"),
            "remote paragraph should remain after undo, got: {after_undo:?}"
        );
        assert!(
            !after_undo.contains("Remote paragraph local"),
            "undo should track local edits inside a remotely created paragraph, got: {after_undo:?}"
        );
    }

    #[test]
    fn test_apply_update_real_remote_change_preserves_local_undo_history() {
        let editor1 = make_test_editor("Hello");
        let editor2 = make_test_editor("");

        let full_state = editor1.encode_state_as_update().expect("full state");
        editor2.apply_update(&full_state).expect("apply full state");

        let hello_index = find_paragraph_index_containing(&editor2, "Hello");
        apply_via_api(
            &editor2,
            &EditIntent::InsertText {
                data: " local".to_string(),
                anchor: encode_position(hello_index, 5),
                attrs: None,
            },
        );

        apply(
            &editor1,
            &EditIntent::InsertText {
                data: " remote".to_string(),
                anchor: encode_position(0, 5),
                attrs: None,
            },
        );

        let sv2 = editor2.encode_state_vector().expect("state vector");
        let diff = editor1.encode_diff(&sv2).expect("diff");
        editor2.apply_update(&diff).expect("apply diff");
        apply_via_api(&editor2, &EditIntent::Undo);

        let after_undo = all_para_text(&editor2);
        assert!(
            !after_undo.contains(" local"),
            "undo should remove the local edit after a real remote mutation, got: {after_undo:?}"
        );
        assert!(
            after_undo.contains("remote"),
            "remote content should remain present after local undo, got: {after_undo:?}"
        );
    }

    #[test]
    fn test_apply_update_remote_paragraph_insert_keeps_prior_local_undo() {
        let editor1 = make_test_editor("Hello");
        let editor2 = make_test_editor("");

        let full_state = editor1.encode_state_as_update().expect("full state");
        editor2.apply_update(&full_state).expect("apply full state");

        let hello_index = find_paragraph_index_containing(&editor2, "Hello");
        apply_via_api(
            &editor2,
            &EditIntent::InsertText {
                data: " local".to_string(),
                anchor: encode_position(hello_index, 5),
                attrs: None,
            },
        );

        apply(
            &editor1,
            &EditIntent::InsertParagraph {
                anchor: encode_position(0, 5),
            },
        );
        apply(
            &editor1,
            &EditIntent::InsertText {
                data: "Remote paragraph".to_string(),
                anchor: encode_position(1, 0),
                attrs: None,
            },
        );

        let sv2 = editor2.encode_state_vector().expect("state vector");
        let diff = editor1.encode_diff(&sv2).expect("diff");
        editor2.apply_update(&diff).expect("apply diff");
        apply_via_api(&editor2, &EditIntent::Undo);

        let after_undo = all_para_text(&editor2);
        assert!(
            !after_undo.contains(" local"),
            "undo should still remove the prior local edit after a structural remote insert, got: {after_undo:?}"
        );
        assert!(
            after_undo.contains("Remote paragraph"),
            "remote paragraph should remain present after local undo, got: {after_undo:?}"
        );
    }

    #[test]
    fn test_insert_table_creates_table_body_item() {
        let editor = make_test_editor("Before");

        apply_via_api(
            &editor,
            &EditIntent::InsertTable {
                anchor: encode_position(0, 0),
                rows: 2,
                columns: 3,
            },
        );

        assert_eq!(body_len(&editor), 2);
        assert_eq!(body_item_type(&editor, 0).as_deref(), Some("paragraph"));
        assert_eq!(body_item_type(&editor, 1).as_deref(), Some("table"));
        assert_eq!(table_dimensions(&editor, 1), Some((2, 3)));
        assert_eq!(table_cell_text(&editor, 1, 0, 0), "");
        assert_eq!(table_cell_text(&editor, 1, 1, 2), "");
    }

    #[test]
    fn test_set_table_cell_text_updates_crdt_and_view() {
        let editor = make_test_editor("Before");

        apply_via_api(
            &editor,
            &EditIntent::InsertTable {
                anchor: encode_position(0, 0),
                rows: 1,
                columns: 2,
            },
        );
        apply_via_api(
            &editor,
            &EditIntent::SetTableCellText {
                body_index: 1,
                row: 0,
                col: 1,
                text: "Updated".to_string(),
            },
        );

        assert_eq!(table_cell_text(&editor, 1, 0, 1), "Updated");

        let crdt = editor.crdt.take().expect("crdt");
        let view_model = view::crdt_to_view_model(&crdt, &editor.original_doc).expect("view");
        editor.crdt.set(Some(crdt));

        let crate::model::BodyItem::Table(table) = &view_model.body[1] else {
            panic!("expected table body item");
        };
        assert_eq!(table.rows.len(), 1);
        assert_eq!(table.rows[0].cells.len(), 2);
        assert_eq!(table.rows[0].cells[1].text, "Updated");
    }

    #[test]
    fn test_insert_table_row_and_column_shift_existing_cells() {
        let editor = make_test_editor("Before");

        apply_via_api(
            &editor,
            &EditIntent::InsertTable {
                anchor: encode_position(0, 0),
                rows: 2,
                columns: 2,
            },
        );
        apply_via_api(
            &editor,
            &EditIntent::SetTableCellText {
                body_index: 1,
                row: 0,
                col: 0,
                text: "A1".to_string(),
            },
        );
        apply_via_api(
            &editor,
            &EditIntent::SetTableCellText {
                body_index: 1,
                row: 1,
                col: 1,
                text: "B2".to_string(),
            },
        );
        apply_via_api(
            &editor,
            &EditIntent::InsertTableRow {
                body_index: 1,
                row: 1,
            },
        );
        apply_via_api(
            &editor,
            &EditIntent::InsertTableColumn {
                body_index: 1,
                col: 1,
            },
        );

        assert_eq!(table_dimensions(&editor, 1), Some((3, 3)));
        assert_eq!(table_cell_text(&editor, 1, 0, 0), "A1");
        assert_eq!(table_cell_text(&editor, 1, 2, 2), "B2");
        assert_eq!(table_cell_text(&editor, 1, 1, 1), "");
    }

    #[test]
    fn test_remove_table_row_and_column_shift_remaining_cells() {
        let editor = make_test_editor("Before");

        apply_via_api(
            &editor,
            &EditIntent::InsertTable {
                anchor: encode_position(0, 0),
                rows: 2,
                columns: 2,
            },
        );
        apply_via_api(
            &editor,
            &EditIntent::SetTableCellText {
                body_index: 1,
                row: 1,
                col: 1,
                text: "B2".to_string(),
            },
        );
        apply_via_api(
            &editor,
            &EditIntent::RemoveTableRow {
                body_index: 1,
                row: 0,
            },
        );
        apply_via_api(
            &editor,
            &EditIntent::RemoveTableColumn {
                body_index: 1,
                col: 0,
            },
        );

        assert_eq!(table_dimensions(&editor, 1), Some((1, 1)));
        assert_eq!(table_cell_text(&editor, 1, 0, 0), "B2");
    }

    #[test]
    fn test_encode_diff() {
        let editor1 = DocEdit::blank().expect("blank editor1");
        let editor2 = DocEdit::blank().expect("blank editor2");

        // Insert text into editor1.
        apply(
            &editor1,
            &EditIntent::InsertText {
                data: "Diff test".to_string(),
                anchor: encode_position(0, 0),
                attrs: None,
            },
        );

        // Get editor2's state vector, compute diff from editor1.
        let sv2 = editor2.encode_state_vector().expect("sv2");
        let diff = editor1.encode_diff(&sv2).expect("encode_diff");
        assert!(!diff.is_empty(), "diff should be non-empty");

        // Apply the diff to editor2.
        editor2.apply_update(&diff).expect("apply diff");

        // editor2 should now contain the text somewhere in its body.
        let text2 = all_para_text(&editor2);
        assert!(
            text2.contains("Diff test"),
            "editor2 body should contain 'Diff test', got: {text2:?}"
        );
    }
}
