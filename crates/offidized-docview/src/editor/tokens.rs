//! Token types for inline non-text content in CRDT rich text.
//!
//! Defines what can appear as U+FFFC sentinel characters in the CRDT text,
//! and provides conversion between [`TokenType`] and yrs `Attrs`.

use std::collections::HashMap;
use std::sync::Arc;
use yrs::Any;

/// The sentinel character used to represent inline tokens in CRDT rich text.
///
/// U+FFFC (Object Replacement Character). This is the ONLY sentinel used.
/// No private-use-area characters are used.
pub const SENTINEL: char = '\u{FFFC}';

/// Attribute key for token type stored in yrs attrs.
pub const ATTR_TOKEN_TYPE: &str = "tokenType";

/// Check if a character is the sentinel.
pub fn is_sentinel(c: char) -> bool {
    c == SENTINEL
}

/// Strip all sentinel characters from text.
///
/// Used on import (stripping literal U+FFFC from `w:t` content) and on
/// user input (stripping sentinels from pasted/typed text).
pub fn strip_sentinels(text: &str) -> String {
    text.replace(SENTINEL, "")
}

/// Token types that can appear as sentinel attributes in CRDT rich text.
///
/// Each token type corresponds to an OOXML inline element. The token's
/// attributes are stored as yrs `Attrs` on the sentinel character.
#[derive(Debug, Clone, PartialEq)]
pub enum TokenType {
    /// Simple field code (`w:fldSimple`).
    FieldSimple {
        /// Field type (e.g., "PAGE", "DATE").
        field_type: String,
        /// Field instruction string.
        instr: String,
        /// Presentation text (visible result).
        presentation: String,
    },
    /// Complex field begin marker (`w:fldChar type="begin"`).
    FieldBegin {
        /// Shared field ID grouping begin/code/separate/end.
        field_id: String,
    },
    /// Complex field instruction text (`w:instrText`).
    FieldCode {
        /// Shared field ID.
        field_id: String,
        /// Instruction string.
        instr: String,
    },
    /// Complex field separate marker (`w:fldChar type="separate"`).
    FieldSeparate {
        /// Shared field ID.
        field_id: String,
    },
    /// Complex field end marker (`w:fldChar type="end"`).
    FieldEnd {
        /// Shared field ID.
        field_id: String,
    },
    /// Footnote reference (`w:footnoteReference`).
    FootnoteRef {
        /// Footnote ID.
        id: u32,
    },
    /// Endnote reference (`w:endnoteReference`).
    EndnoteRef {
        /// Endnote ID.
        id: u32,
    },
    /// Bookmark start marker (`w:bookmarkStart`).
    BookmarkStart {
        /// Bookmark ID.
        id: String,
        /// Bookmark name.
        name: String,
    },
    /// Bookmark end marker (`w:bookmarkEnd`).
    BookmarkEnd {
        /// Bookmark ID.
        id: String,
    },
    /// Comment range start (`w:commentRangeStart`).
    CommentStart {
        /// Comment ID.
        id: String,
    },
    /// Comment range end (`w:commentRangeEnd`).
    CommentEnd {
        /// Comment ID.
        id: String,
    },
    /// Tab character (`w:tab`).
    Tab,
    /// Line break (`w:br`).
    LineBreak {
        /// Break type (e.g., "page", "column"). None for regular line break.
        break_type: Option<String>,
    },
    /// Inline image (`w:drawing` with `wp:inline`).
    InlineImage {
        /// Content-hash reference into the blob store.
        image_ref: String,
        /// Width in EMUs.
        width: i64,
        /// Height in EMUs.
        height: i64,
    },
    /// Unknown run child preserved as raw XML.
    ///
    /// Has a stable `opaque_id` for immutability enforcement in collaboration.
    Opaque {
        /// Stable ID for this opaque token (UUID v7, assigned on import).
        opaque_id: String,
        /// Serialized XML string, re-emitted verbatim on export.
        xml: String,
    },
}

/// Convert a [`TokenType`] to yrs attributes for a sentinel character.
///
/// The returned map can be used with `Text::insert_with_attributes` to
/// store token metadata alongside the sentinel character.
pub fn token_to_attrs(token: &TokenType) -> HashMap<Arc<str>, Any> {
    let mut attrs = HashMap::new();
    match token {
        TokenType::Tab => {
            attrs.insert(Arc::from(ATTR_TOKEN_TYPE), Any::String(Arc::from("tab")));
        }
        TokenType::LineBreak { break_type } => {
            attrs.insert(
                Arc::from(ATTR_TOKEN_TYPE),
                Any::String(Arc::from("lineBreak")),
            );
            if let Some(bt) = break_type {
                attrs.insert(Arc::from("breakType"), Any::String(Arc::from(bt.as_str())));
            }
        }
        TokenType::FootnoteRef { id } => {
            attrs.insert(
                Arc::from(ATTR_TOKEN_TYPE),
                Any::String(Arc::from("footnoteRef")),
            );
            attrs.insert(Arc::from("id"), Any::Number(f64::from(*id)));
        }
        TokenType::EndnoteRef { id } => {
            attrs.insert(
                Arc::from(ATTR_TOKEN_TYPE),
                Any::String(Arc::from("endnoteRef")),
            );
            attrs.insert(Arc::from("id"), Any::Number(f64::from(*id)));
        }
        TokenType::FieldSimple {
            field_type,
            instr,
            presentation,
        } => {
            attrs.insert(
                Arc::from(ATTR_TOKEN_TYPE),
                Any::String(Arc::from("fieldSimple")),
            );
            attrs.insert(
                Arc::from("fieldType"),
                Any::String(Arc::from(field_type.as_str())),
            );
            attrs.insert(Arc::from("instr"), Any::String(Arc::from(instr.as_str())));
            attrs.insert(
                Arc::from("presentation"),
                Any::String(Arc::from(presentation.as_str())),
            );
        }
        TokenType::FieldBegin { field_id } => {
            attrs.insert(
                Arc::from(ATTR_TOKEN_TYPE),
                Any::String(Arc::from("fieldBegin")),
            );
            attrs.insert(
                Arc::from("fieldId"),
                Any::String(Arc::from(field_id.as_str())),
            );
        }
        TokenType::FieldCode { field_id, instr } => {
            attrs.insert(
                Arc::from(ATTR_TOKEN_TYPE),
                Any::String(Arc::from("fieldCode")),
            );
            attrs.insert(
                Arc::from("fieldId"),
                Any::String(Arc::from(field_id.as_str())),
            );
            attrs.insert(Arc::from("instr"), Any::String(Arc::from(instr.as_str())));
        }
        TokenType::FieldSeparate { field_id } => {
            attrs.insert(
                Arc::from(ATTR_TOKEN_TYPE),
                Any::String(Arc::from("fieldSeparate")),
            );
            attrs.insert(
                Arc::from("fieldId"),
                Any::String(Arc::from(field_id.as_str())),
            );
        }
        TokenType::FieldEnd { field_id } => {
            attrs.insert(
                Arc::from(ATTR_TOKEN_TYPE),
                Any::String(Arc::from("fieldEnd")),
            );
            attrs.insert(
                Arc::from("fieldId"),
                Any::String(Arc::from(field_id.as_str())),
            );
        }
        TokenType::BookmarkStart { id, name } => {
            attrs.insert(
                Arc::from(ATTR_TOKEN_TYPE),
                Any::String(Arc::from("bookmarkStart")),
            );
            attrs.insert(Arc::from("id"), Any::String(Arc::from(id.as_str())));
            attrs.insert(Arc::from("name"), Any::String(Arc::from(name.as_str())));
        }
        TokenType::BookmarkEnd { id } => {
            attrs.insert(
                Arc::from(ATTR_TOKEN_TYPE),
                Any::String(Arc::from("bookmarkEnd")),
            );
            attrs.insert(Arc::from("id"), Any::String(Arc::from(id.as_str())));
        }
        TokenType::CommentStart { id } => {
            attrs.insert(
                Arc::from(ATTR_TOKEN_TYPE),
                Any::String(Arc::from("commentStart")),
            );
            attrs.insert(Arc::from("id"), Any::String(Arc::from(id.as_str())));
        }
        TokenType::CommentEnd { id } => {
            attrs.insert(
                Arc::from(ATTR_TOKEN_TYPE),
                Any::String(Arc::from("commentEnd")),
            );
            attrs.insert(Arc::from("id"), Any::String(Arc::from(id.as_str())));
        }
        TokenType::InlineImage {
            image_ref,
            width,
            height,
        } => {
            attrs.insert(
                Arc::from(ATTR_TOKEN_TYPE),
                Any::String(Arc::from("inlineImage")),
            );
            attrs.insert(
                Arc::from("imageRef"),
                Any::String(Arc::from(image_ref.as_str())),
            );
            attrs.insert(Arc::from("width"), Any::Number(*width as f64));
            attrs.insert(Arc::from("height"), Any::Number(*height as f64));
        }
        TokenType::Opaque { opaque_id, xml } => {
            attrs.insert(Arc::from(ATTR_TOKEN_TYPE), Any::String(Arc::from("opaque")));
            attrs.insert(
                Arc::from("opaqueId"),
                Any::String(Arc::from(opaque_id.as_str())),
            );
            attrs.insert(Arc::from("xml"), Any::String(Arc::from(xml.as_str())));
        }
    }
    attrs
}

/// Extract a string value from a yrs [`Any`].
fn any_as_str(val: &Any) -> Option<&str> {
    match val {
        Any::String(s) => Some(s),
        _ => None,
    }
}

/// Extract a number value from a yrs [`Any`].
fn any_as_f64(val: &Any) -> Option<f64> {
    match val {
        Any::Number(n) => Some(*n),
        _ => None,
    }
}

/// Convert yrs attributes back to a [`TokenType`].
///
/// Returns `None` if the attrs do not contain a valid `tokenType` key
/// or if required fields for that token type are missing.
pub fn attrs_to_token(attrs: &HashMap<Arc<str>, Any>) -> Option<TokenType> {
    let token_type = any_as_str(attrs.get("tokenType" as &str)?)?;
    match token_type {
        "tab" => Some(TokenType::Tab),
        "lineBreak" => {
            let break_type = attrs
                .get("breakType" as &str)
                .and_then(any_as_str)
                .map(String::from);
            Some(TokenType::LineBreak { break_type })
        }
        "footnoteRef" => {
            let id = any_as_f64(attrs.get("id" as &str)?)? as u32;
            Some(TokenType::FootnoteRef { id })
        }
        "endnoteRef" => {
            let id = any_as_f64(attrs.get("id" as &str)?)? as u32;
            Some(TokenType::EndnoteRef { id })
        }
        "fieldSimple" => {
            let field_type = any_as_str(attrs.get("fieldType" as &str)?)?.to_string();
            let instr = any_as_str(attrs.get("instr" as &str)?)?.to_string();
            let presentation = any_as_str(attrs.get("presentation" as &str)?)?.to_string();
            Some(TokenType::FieldSimple {
                field_type,
                instr,
                presentation,
            })
        }
        "fieldBegin" => {
            let field_id = any_as_str(attrs.get("fieldId" as &str)?)?.to_string();
            Some(TokenType::FieldBegin { field_id })
        }
        "fieldCode" => {
            let field_id = any_as_str(attrs.get("fieldId" as &str)?)?.to_string();
            let instr = any_as_str(attrs.get("instr" as &str)?)?.to_string();
            Some(TokenType::FieldCode { field_id, instr })
        }
        "fieldSeparate" => {
            let field_id = any_as_str(attrs.get("fieldId" as &str)?)?.to_string();
            Some(TokenType::FieldSeparate { field_id })
        }
        "fieldEnd" => {
            let field_id = any_as_str(attrs.get("fieldId" as &str)?)?.to_string();
            Some(TokenType::FieldEnd { field_id })
        }
        "bookmarkStart" => {
            let id = any_as_str(attrs.get("id" as &str)?)?.to_string();
            let name = any_as_str(attrs.get("name" as &str)?)?.to_string();
            Some(TokenType::BookmarkStart { id, name })
        }
        "bookmarkEnd" => {
            let id = any_as_str(attrs.get("id" as &str)?)?.to_string();
            Some(TokenType::BookmarkEnd { id })
        }
        "commentStart" => {
            let id = any_as_str(attrs.get("id" as &str)?)?.to_string();
            Some(TokenType::CommentStart { id })
        }
        "commentEnd" => {
            let id = any_as_str(attrs.get("id" as &str)?)?.to_string();
            Some(TokenType::CommentEnd { id })
        }
        "inlineImage" => {
            let image_ref = any_as_str(attrs.get("imageRef" as &str)?)?.to_string();
            let width = any_as_f64(attrs.get("width" as &str)?)? as i64;
            let height = any_as_f64(attrs.get("height" as &str)?)? as i64;
            Some(TokenType::InlineImage {
                image_ref,
                width,
                height,
            })
        }
        "opaque" => {
            let opaque_id = any_as_str(attrs.get("opaqueId" as &str)?)?.to_string();
            let xml = any_as_str(attrs.get("xml" as &str)?)?.to_string();
            Some(TokenType::Opaque { opaque_id, xml })
        }
        _ => None,
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn sentinel_constant() {
        assert!(is_sentinel('\u{FFFC}'));
        assert!(!is_sentinel('a'));
    }

    #[test]
    fn strip_sentinels_removes_all() {
        let text = format!("hello{}world{}!", SENTINEL, SENTINEL);
        assert_eq!(strip_sentinels(&text), "helloworld!");
    }

    #[test]
    fn strip_sentinels_no_op_on_clean() {
        assert_eq!(strip_sentinels("hello world"), "hello world");
    }

    #[test]
    fn roundtrip_tab() {
        let token = TokenType::Tab;
        let attrs = token_to_attrs(&token);
        let recovered = attrs_to_token(&attrs);
        assert_eq!(recovered, Some(token));
    }

    #[test]
    fn roundtrip_line_break_with_type() {
        let token = TokenType::LineBreak {
            break_type: Some("page".to_string()),
        };
        let attrs = token_to_attrs(&token);
        let recovered = attrs_to_token(&attrs);
        assert_eq!(recovered, Some(token));
    }

    #[test]
    fn roundtrip_line_break_no_type() {
        let token = TokenType::LineBreak { break_type: None };
        let attrs = token_to_attrs(&token);
        let recovered = attrs_to_token(&attrs);
        assert_eq!(recovered, Some(token));
    }

    #[test]
    fn roundtrip_footnote_ref() {
        let token = TokenType::FootnoteRef { id: 42 };
        let attrs = token_to_attrs(&token);
        let recovered = attrs_to_token(&attrs);
        assert_eq!(recovered, Some(token));
    }

    #[test]
    fn roundtrip_field_simple() {
        let token = TokenType::FieldSimple {
            field_type: "PAGE".to_string(),
            instr: "PAGE \\* MERGEFORMAT".to_string(),
            presentation: "3".to_string(),
        };
        let attrs = token_to_attrs(&token);
        let recovered = attrs_to_token(&attrs);
        assert_eq!(recovered, Some(token));
    }

    #[test]
    fn roundtrip_inline_image() {
        let token = TokenType::InlineImage {
            image_ref: "sha256:abc123".to_string(),
            width: 914400,
            height: 457200,
        };
        let attrs = token_to_attrs(&token);
        let recovered = attrs_to_token(&attrs);
        assert_eq!(recovered, Some(token));
    }

    #[test]
    fn roundtrip_opaque() {
        let token = TokenType::Opaque {
            opaque_id: "some-uuid".to_string(),
            xml: "<w:sym w:font=\"Wingdings\" w:char=\"F0FC\"/>".to_string(),
        };
        let attrs = token_to_attrs(&token);
        let recovered = attrs_to_token(&attrs);
        assert_eq!(recovered, Some(token));
    }

    #[test]
    fn roundtrip_bookmark() {
        let token = TokenType::BookmarkStart {
            id: "0".to_string(),
            name: "_GoBack".to_string(),
        };
        let attrs = token_to_attrs(&token);
        let recovered = attrs_to_token(&attrs);
        assert_eq!(recovered, Some(token));
    }

    #[test]
    fn roundtrip_complex_field() {
        let tokens = vec![
            TokenType::FieldBegin {
                field_id: "f1".to_string(),
            },
            TokenType::FieldCode {
                field_id: "f1".to_string(),
                instr: "TOC \\o".to_string(),
            },
            TokenType::FieldSeparate {
                field_id: "f1".to_string(),
            },
            TokenType::FieldEnd {
                field_id: "f1".to_string(),
            },
        ];
        for token in tokens {
            let attrs = token_to_attrs(&token);
            let recovered = attrs_to_token(&attrs);
            assert_eq!(recovered, Some(token));
        }
    }

    #[test]
    fn unknown_token_type_returns_none() {
        let mut attrs = HashMap::new();
        attrs.insert(
            Arc::from(ATTR_TOKEN_TYPE),
            Any::String(Arc::from("unknown")),
        );
        assert_eq!(attrs_to_token(&attrs), None);
    }

    #[test]
    fn empty_attrs_returns_none() {
        let attrs = HashMap::new();
        assert_eq!(attrs_to_token(&attrs), None);
    }
}
