//! Shared XML serialization utilities used by all format crates.

use quick_xml::events::BytesStart;

/// Push unknown/preserved attributes onto a [`BytesStart`] tag, skipping any
/// whose key has already been written as a known attribute.
///
/// This provides a generic safety net against duplicate XML attributes. When a
/// format crate's serializer explicitly writes attributes (e.g. `ht`, `customHeight`,
/// `hidden`) and then replays preserved unknown attributes from the original XML,
/// any overlap would produce invalid XML (duplicate attribute keys). This function
/// filters out those overlaps.
///
/// # Arguments
///
/// * `tag` - The element tag being built.
/// * `unknown_attrs` - Preserved attributes from the original XML.
/// * `known_keys` - Keys of attributes already pushed onto `tag` by the caller.
///
/// # Example
///
/// ```ignore
/// let mut row = BytesStart::new("row");
/// let mut known = vec!["r"];
/// row.push_attribute(("r", "1"));
///
/// if let Some(ht) = height_text.as_deref() {
///     row.push_attribute(("ht", ht));
///     row.push_attribute(("customHeight", "1"));
///     known.extend(["ht", "customHeight"]);
/// }
///
/// push_unknown_attrs_deduped(&mut row, unknown_attrs, &known);
/// ```
pub fn push_unknown_attrs_deduped(
    tag: &mut BytesStart<'_>,
    unknown_attrs: &[(String, String)],
    known_keys: &[&str],
) {
    for (key, value) in unknown_attrs {
        if !known_keys.contains(&key.as_str()) {
            tag.push_attribute((key.as_str(), value.as_str()));
        }
    }
}

/// Capture extra namespace declarations (any `xmlns:*` attributes beyond those
/// that the serializer always emits) from a parsed XML start event.
///
/// `always_emitted` should list the namespace prefixes the serializer hardcodes
/// (e.g. `&["xmlns", "xmlns:r"]`). All other `xmlns:*` attributes are returned
/// for replay during dirty-save reconstruction.
pub fn capture_extra_namespace_declarations(
    event: &BytesStart<'_>,
    always_emitted: &[&str],
) -> Vec<(String, String)> {
    let mut extra = Vec::new();
    for attr in event.attributes().flatten() {
        let key = String::from_utf8_lossy(attr.key.as_ref()).into_owned();
        if key.starts_with("xmlns:") && !always_emitted.contains(&key.as_str()) {
            let val = String::from_utf8_lossy(attr.value.as_ref()).into_owned();
            extra.push((key, val));
        }
    }
    extra
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn dedup_filters_known_keys() {
        let mut tag = BytesStart::new("row");
        tag.push_attribute(("r", "1"));
        tag.push_attribute(("ht", "15"));
        tag.push_attribute(("customHeight", "1"));

        let unknown_attrs = vec![
            ("customHeight".to_string(), "1".to_string()),
            ("spans".to_string(), "1:5".to_string()),
            ("x14ac:dyDescent".to_string(), "0.25".to_string()),
        ];

        push_unknown_attrs_deduped(&mut tag, &unknown_attrs, &["r", "ht", "customHeight"]);

        // Serialize the tag to XML to verify
        let mut writer = quick_xml::Writer::new(Vec::new());
        writer
            .write_event(quick_xml::events::Event::Empty(tag))
            .unwrap();
        let xml = String::from_utf8(writer.into_inner()).unwrap();

        // customHeight should appear exactly once
        assert_eq!(xml.matches("customHeight").count(), 1, "xml was: {xml}");
        // spans and x14ac:dyDescent should still appear
        assert!(xml.contains("spans"), "xml was: {xml}");
        assert!(xml.contains("x14ac:dyDescent"), "xml was: {xml}");
    }

    #[test]
    fn dedup_passes_through_when_no_overlap() {
        let mut tag = BytesStart::new("c");
        tag.push_attribute(("r", "A1"));

        let unknown_attrs = vec![
            ("foo".to_string(), "bar".to_string()),
            ("baz".to_string(), "qux".to_string()),
        ];

        push_unknown_attrs_deduped(&mut tag, &unknown_attrs, &["r"]);

        let mut writer = quick_xml::Writer::new(Vec::new());
        writer
            .write_event(quick_xml::events::Event::Empty(tag))
            .unwrap();
        let xml = String::from_utf8(writer.into_inner()).unwrap();

        assert!(xml.contains("foo=\"bar\""), "xml was: {xml}");
        assert!(xml.contains("baz=\"qux\""), "xml was: {xml}");
    }

    #[test]
    fn capture_ns_skips_always_emitted() {
        let mut tag = BytesStart::new("worksheet");
        tag.push_attribute(("xmlns", "http://main"));
        tag.push_attribute(("xmlns:r", "http://rel"));
        tag.push_attribute(("xmlns:x14ac", "http://x14ac"));
        tag.push_attribute(("xmlns:mc", "http://mc"));

        let extra = capture_extra_namespace_declarations(&tag, &["xmlns", "xmlns:r"]);

        assert_eq!(extra.len(), 2);
        assert!(extra
            .iter()
            .any(|(k, v)| k == "xmlns:x14ac" && v == "http://x14ac"));
        assert!(extra
            .iter()
            .any(|(k, v)| k == "xmlns:mc" && v == "http://mc"));
    }
}
