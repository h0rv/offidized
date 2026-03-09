//! Editing intents from the TypeScript editor shell.
//!
//! Intents represent user actions (keystrokes, formatting commands, clipboard
//! operations) that the WASM `apply_intent()` handler processes and translates
//! into CRDT transactions.

use serde::{Deserialize, Serialize};

/// Generic attribute value carried across the TS/WASM intent boundary.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
#[serde(untagged)]
pub enum IntentAttrValue {
    /// Remove the attribute when applied.
    Null,
    /// Boolean attribute value.
    Bool(bool),
    /// Numeric attribute value.
    Number(f64),
    /// String attribute value.
    String(String),
}

/// An editing intent from the TypeScript editor shell.
///
/// Intents represent user actions that modify the document.
/// The WASM `apply_intent()` handler processes these and translates
/// them into CRDT transactions.
#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(tag = "type", rename_all = "camelCase")]
pub enum EditIntent {
    /// Insert text at the cursor position.
    #[serde(rename = "insertText")]
    InsertText {
        /// The text to insert (U+FFFC will be stripped).
        data: String,
        /// CRDT-relative anchor position (opaque bytes, base64-encoded).
        anchor: String,
        /// Optional formatting attributes to apply to the inserted text.
        /// When absent, the Rust side inherits from the character to the left.
        #[serde(default, skip_serializing_if = "Option::is_none")]
        attrs: Option<std::collections::HashMap<String, IntentAttrValue>>,
    },
    /// Insert committed IME/composition text at the cursor position.
    #[serde(rename = "insertFromComposition")]
    InsertFromComposition {
        /// The composed text to insert (U+FFFC will be stripped).
        data: String,
        /// CRDT-relative anchor position (opaque bytes, base64-encoded).
        anchor: String,
        /// Optional formatting attributes to apply to the inserted text.
        /// When absent, the Rust side inherits from the character to the left.
        #[serde(default, skip_serializing_if = "Option::is_none")]
        attrs: Option<std::collections::HashMap<String, IntentAttrValue>>,
    },
    /// Delete content backward (Backspace).
    #[serde(rename = "deleteContentBackward")]
    DeleteBackward {
        /// CRDT-relative anchor position.
        anchor: String,
        /// CRDT-relative focus position (for range delete).
        focus: String,
    },
    /// Delete content forward (Delete key).
    #[serde(rename = "deleteContentForward")]
    DeleteForward {
        /// CRDT-relative anchor position.
        anchor: String,
        /// CRDT-relative focus position.
        focus: String,
    },
    /// Insert a paragraph break (Enter key).
    #[serde(rename = "insertParagraph")]
    InsertParagraph {
        /// Position where the split occurs.
        anchor: String,
    },
    /// Insert a line break (Shift+Enter).
    #[serde(rename = "insertLineBreak")]
    InsertLineBreak {
        /// Position for the line break.
        anchor: String,
    },
    /// Insert a tab character.
    #[serde(rename = "insertTab")]
    InsertTab {
        /// Position for the tab.
        anchor: String,
    },
    /// Toggle bold formatting on selection.
    #[serde(rename = "formatBold")]
    FormatBold {
        /// Selection anchor.
        anchor: String,
        /// Selection focus.
        focus: String,
    },
    /// Toggle italic formatting on selection.
    #[serde(rename = "formatItalic")]
    FormatItalic {
        /// Selection anchor.
        anchor: String,
        /// Selection focus.
        focus: String,
    },
    /// Toggle underline formatting on selection.
    #[serde(rename = "formatUnderline")]
    FormatUnderline {
        /// Selection anchor.
        anchor: String,
        /// Selection focus.
        focus: String,
    },
    /// Toggle strikethrough formatting on selection.
    #[serde(rename = "formatStrikethrough")]
    FormatStrikethrough {
        /// Selection anchor.
        anchor: String,
        /// Selection focus.
        focus: String,
    },
    /// Set explicit text attributes across a selection.
    #[serde(rename = "setTextAttrs")]
    SetTextAttrs {
        /// Selection anchor.
        anchor: String,
        /// Selection focus.
        focus: String,
        /// Attribute patch to apply. `null` removes an attribute.
        attrs: std::collections::HashMap<String, IntentAttrValue>,
    },
    /// Set paragraph attributes across one or more selected paragraphs.
    #[serde(rename = "setParagraphAttrs")]
    SetParagraphAttrs {
        /// Selection anchor.
        anchor: String,
        /// Selection focus.
        focus: String,
        /// Paragraph attribute patch to apply. `null` removes an attribute.
        attrs: std::collections::HashMap<String, IntentAttrValue>,
    },
    /// Insert content from paste.
    #[serde(rename = "insertFromPaste")]
    InsertFromPaste {
        /// Pasted text (U+FFFC will be stripped).
        data: String,
        /// Position to insert at.
        anchor: String,
        /// End of selection to replace (same as anchor if no selection).
        focus: String,
        /// Optional formatting attributes to apply to the inserted text.
        /// When absent, the Rust side inherits from the character to the left.
        #[serde(default, skip_serializing_if = "Option::is_none")]
        attrs: Option<std::collections::HashMap<String, IntentAttrValue>>,
    },
    /// Insert an inline image at the current selection.
    #[serde(rename = "insertInlineImage")]
    InsertInlineImage {
        /// Selection anchor / insertion position.
        anchor: String,
        /// End of selection to replace (same as anchor if no selection).
        focus: String,
        /// Browser data URI for the image payload.
        data_uri: String,
        /// Display width in points.
        width_pt: f64,
        /// Display height in points.
        height_pt: f64,
        /// Optional image name.
        #[serde(default, skip_serializing_if = "Option::is_none")]
        name: Option<String>,
        /// Optional image description / alt text.
        #[serde(default, skip_serializing_if = "Option::is_none")]
        description: Option<String>,
    },
    /// Insert a new table after the current body item.
    #[serde(rename = "insertTable")]
    InsertTable {
        /// Selection anchor used to choose the insertion body item.
        anchor: String,
        /// Number of rows in the inserted table.
        rows: usize,
        /// Number of columns in the inserted table.
        columns: usize,
    },
    /// Replace the plain-text contents of a table cell.
    #[serde(rename = "setTableCellText")]
    SetTableCellText {
        /// Body index of the table in the CRDT body array.
        #[serde(rename = "bodyIndex")]
        body_index: u32,
        /// Zero-based table row index.
        row: usize,
        /// Zero-based table column index.
        col: usize,
        /// New cell text.
        text: String,
    },
    /// Insert a blank row into a table.
    #[serde(rename = "insertTableRow")]
    InsertTableRow {
        /// Body index of the table in the CRDT body array.
        #[serde(rename = "bodyIndex")]
        body_index: u32,
        /// Zero-based row index at which to insert the blank row.
        row: usize,
    },
    /// Remove a row from a table.
    #[serde(rename = "removeTableRow")]
    RemoveTableRow {
        /// Body index of the table in the CRDT body array.
        #[serde(rename = "bodyIndex")]
        body_index: u32,
        /// Zero-based row index to remove.
        row: usize,
    },
    /// Insert a blank column into a table.
    #[serde(rename = "insertTableColumn")]
    InsertTableColumn {
        /// Body index of the table in the CRDT body array.
        #[serde(rename = "bodyIndex")]
        body_index: u32,
        /// Zero-based column index at which to insert the blank column.
        col: usize,
    },
    /// Remove a column from a table.
    #[serde(rename = "removeTableColumn")]
    RemoveTableColumn {
        /// Body index of the table in the CRDT body array.
        #[serde(rename = "bodyIndex")]
        body_index: u32,
        /// Zero-based column index to remove.
        col: usize,
    },
    /// Delete by cut.
    #[serde(rename = "deleteByCut")]
    DeleteByCut {
        /// Selection anchor.
        anchor: String,
        /// Selection focus.
        focus: String,
    },
    /// Undo last action.
    #[serde(rename = "historyUndo")]
    Undo,
    /// Redo last undone action.
    #[serde(rename = "historyRedo")]
    Redo,
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn deserialize_insert_text() {
        let json = r#"{"type":"insertText","data":"hello","anchor":"AAAA"}"#;
        let intent: EditIntent = serde_json::from_str(json).expect("deserialize insertText");
        match intent {
            EditIntent::InsertText { data, anchor, .. } => {
                assert_eq!(data, "hello");
                assert_eq!(anchor, "AAAA");
            }
            other => panic!("expected InsertText, got {other:?}"),
        }
    }

    #[test]
    fn deserialize_insert_text_with_string_attr() {
        let json = r#"{"type":"insertText","data":"hello","anchor":"AAAA","attrs":{"fontFamily":"Noto Sans"}}"#;
        let intent: EditIntent =
            serde_json::from_str(json).expect("deserialize insertText with attrs");
        match intent {
            EditIntent::InsertText { attrs, .. } => {
                let attrs = attrs.expect("attrs");
                assert_eq!(
                    attrs.get("fontFamily"),
                    Some(&IntentAttrValue::String("Noto Sans".to_string()))
                );
            }
            other => panic!("expected InsertText, got {other:?}"),
        }
    }

    #[test]
    fn deserialize_undo() {
        let json = r#"{"type":"historyUndo"}"#;
        let intent: EditIntent = serde_json::from_str(json).expect("deserialize historyUndo");
        assert!(matches!(intent, EditIntent::Undo));
    }

    #[test]
    fn deserialize_format_bold() {
        let json = r#"{"type":"formatBold","anchor":"AA","focus":"BB"}"#;
        let intent: EditIntent = serde_json::from_str(json).expect("deserialize formatBold");
        match intent {
            EditIntent::FormatBold { anchor, focus } => {
                assert_eq!(anchor, "AA");
                assert_eq!(focus, "BB");
            }
            other => panic!("expected FormatBold, got {other:?}"),
        }
    }

    #[test]
    fn deserialize_set_text_attrs() {
        let json = r#"{"type":"setTextAttrs","anchor":"AA","focus":"BB","attrs":{"fontSizePt":12,"color":"FF0000","bold":true,"fontFamily":null}}"#;
        let intent: EditIntent = serde_json::from_str(json).expect("deserialize setTextAttrs");
        match intent {
            EditIntent::SetTextAttrs {
                anchor,
                focus,
                attrs,
            } => {
                assert_eq!(anchor, "AA");
                assert_eq!(focus, "BB");
                assert_eq!(
                    attrs.get("fontSizePt"),
                    Some(&IntentAttrValue::Number(12.0))
                );
                assert_eq!(
                    attrs.get("color"),
                    Some(&IntentAttrValue::String("FF0000".to_string()))
                );
                assert_eq!(attrs.get("bold"), Some(&IntentAttrValue::Bool(true)));
                assert_eq!(attrs.get("fontFamily"), Some(&IntentAttrValue::Null));
            }
            other => panic!("expected SetTextAttrs, got {other:?}"),
        }
    }

    #[test]
    fn roundtrip_serialize() {
        let intent = EditIntent::InsertParagraph {
            anchor: "pos123".to_string(),
        };
        let json = serde_json::to_string(&intent).expect("serialize");
        let recovered: EditIntent = serde_json::from_str(&json).expect("deserialize");
        match recovered {
            EditIntent::InsertParagraph { anchor } => {
                assert_eq!(anchor, "pos123");
            }
            other => panic!("expected InsertParagraph, got {other:?}"),
        }
    }

    #[test]
    fn deserialize_insert_table() {
        let json = r#"{"type":"insertTable","anchor":"AAAA","rows":2,"columns":3}"#;
        let intent: EditIntent = serde_json::from_str(json).expect("deserialize insertTable");
        match intent {
            EditIntent::InsertTable {
                anchor,
                rows,
                columns,
            } => {
                assert_eq!(anchor, "AAAA");
                assert_eq!(rows, 2);
                assert_eq!(columns, 3);
            }
            other => panic!("expected InsertTable, got {other:?}"),
        }
    }

    #[test]
    fn deserialize_set_table_cell_text() {
        let json = r#"{"type":"setTableCellText","bodyIndex":4,"row":1,"col":2,"text":"Updated"}"#;
        let intent: EditIntent = serde_json::from_str(json).expect("deserialize setTableCellText");
        match intent {
            EditIntent::SetTableCellText {
                body_index,
                row,
                col,
                text,
            } => {
                assert_eq!(body_index, 4);
                assert_eq!(row, 1);
                assert_eq!(col, 2);
                assert_eq!(text, "Updated");
            }
            other => panic!("expected SetTableCellText, got {other:?}"),
        }
    }

    #[test]
    fn deserialize_insert_table_row() {
        let json = r#"{"type":"insertTableRow","bodyIndex":4,"row":1}"#;
        let intent: EditIntent = serde_json::from_str(json).expect("deserialize insertTableRow");
        match intent {
            EditIntent::InsertTableRow { body_index, row } => {
                assert_eq!(body_index, 4);
                assert_eq!(row, 1);
            }
            other => panic!("expected InsertTableRow, got {other:?}"),
        }
    }

    #[test]
    fn deserialize_remove_table_column() {
        let json = r#"{"type":"removeTableColumn","bodyIndex":4,"col":2}"#;
        let intent: EditIntent = serde_json::from_str(json).expect("deserialize removeTableColumn");
        match intent {
            EditIntent::RemoveTableColumn { body_index, col } => {
                assert_eq!(body_index, 4);
                assert_eq!(col, 2);
            }
            other => panic!("expected RemoveTableColumn, got {other:?}"),
        }
    }
}
