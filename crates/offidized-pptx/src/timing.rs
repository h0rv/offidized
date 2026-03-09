use std::io::Cursor;

use quick_xml::events::{BytesStart, Event};
use quick_xml::Reader;

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct SlideAnimationNode {
    id: u32,
    duration_ms: Option<u64>,
    trigger: Option<String>,
    event: Option<String>,
}

impl SlideAnimationNode {
    pub fn new(id: u32) -> Self {
        Self {
            id,
            duration_ms: None,
            trigger: None,
            event: None,
        }
    }

    pub fn id(&self) -> u32 {
        self.id
    }

    pub fn set_id(&mut self, id: u32) {
        self.id = id;
    }

    pub fn duration_ms(&self) -> Option<u64> {
        self.duration_ms
    }

    pub fn set_duration_ms(&mut self, duration_ms: Option<u64>) {
        self.duration_ms = duration_ms;
    }

    pub fn trigger(&self) -> Option<&str> {
        self.trigger.as_deref()
    }

    pub fn set_trigger(&mut self, trigger: Option<impl Into<String>>) {
        self.trigger = trigger.map(Into::into);
    }

    pub fn event(&self) -> Option<&str> {
        self.event.as_deref()
    }

    pub fn set_event(&mut self, event: Option<impl Into<String>>) {
        self.event = event.map(Into::into);
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct SlideTiming {
    raw_inner_xml: String,
    animations: Vec<SlideAnimationNode>,
}

impl SlideTiming {
    pub fn new(raw_inner_xml: impl Into<String>) -> Self {
        let raw_inner_xml = raw_inner_xml.into();
        let animations = parse_animation_nodes(raw_inner_xml.as_str());
        Self {
            raw_inner_xml,
            animations,
        }
    }

    pub fn raw_inner_xml(&self) -> &str {
        self.raw_inner_xml.as_str()
    }

    pub fn set_raw_inner_xml(&mut self, raw_inner_xml: impl Into<String>) {
        self.raw_inner_xml = raw_inner_xml.into();
        self.animations = parse_animation_nodes(self.raw_inner_xml.as_str());
    }

    pub fn animations(&self) -> &[SlideAnimationNode] {
        self.animations.as_slice()
    }

    pub fn animations_mut(&mut self) -> &mut Vec<SlideAnimationNode> {
        self.raw_inner_xml.clear();
        &mut self.animations
    }
}

fn parse_animation_nodes(raw_inner_xml: &str) -> Vec<SlideAnimationNode> {
    if raw_inner_xml.is_empty() {
        return Vec::new();
    }

    let mut reader = Reader::from_reader(Cursor::new(raw_inner_xml.as_bytes()));
    reader.config_mut().trim_text(false);
    let mut buffer = Vec::new();
    let mut nodes = Vec::new();

    loop {
        match reader.read_event_into(&mut buffer) {
            Ok(Event::Start(ref event)) | Ok(Event::Empty(ref event)) => {
                if local_name(event.name().as_ref()) != b"cTn" {
                    buffer.clear();
                    continue;
                }

                if let Some(node) = parse_c_tn_node(event) {
                    nodes.push(node);
                }
            }
            Ok(Event::Eof) => break,
            Ok(_) => {}
            Err(_) => return Vec::new(),
        }
        buffer.clear();
    }

    nodes
}

fn parse_c_tn_node(event: &BytesStart<'_>) -> Option<SlideAnimationNode> {
    let id = get_attribute_value(event, b"id")?.parse::<u32>().ok()?;
    let mut node = SlideAnimationNode::new(id);
    node.set_duration_ms(
        get_attribute_value(event, b"dur").and_then(|value| parse_duration_ms(value.as_str())),
    );
    node.set_trigger(
        get_attribute_value(event, b"trigger").or_else(|| get_attribute_value(event, b"nodeType")),
    );
    node.set_event(
        get_attribute_value(event, b"evt").or_else(|| get_attribute_value(event, b"evtFilter")),
    );
    Some(node)
}

fn parse_duration_ms(value: &str) -> Option<u64> {
    value.parse::<u64>().ok()
}

fn get_attribute_value(event: &BytesStart<'_>, expected_local_name: &[u8]) -> Option<String> {
    event.attributes().flatten().find_map(|attribute| {
        (local_name(attribute.key.as_ref()) == expected_local_name)
            .then(|| String::from_utf8_lossy(attribute.value.as_ref()).into_owned())
    })
}

fn local_name(name: &[u8]) -> &[u8] {
    name.rsplit(|byte| *byte == b':').next().unwrap_or(name)
}

#[cfg(test)]
mod tests {
    use super::{SlideAnimationNode, SlideTiming};

    #[test]
    fn stores_raw_inner_xml_and_parses_typed_nodes() {
        let timing_inner_xml = concat!(
            r#"<p:tnLst>"#,
            r#"<p:par><p:cTn id="1" dur="indefinite" nodeType="tmRoot"/></p:par>"#,
            r#"<p:par><p:cTn id="2" dur="350" nodeType="clickEffect" evtFilter="cancelBubble"/></p:par>"#,
            r#"</p:tnLst>"#,
        );
        let timing = SlideTiming::new(timing_inner_xml);

        assert_eq!(timing.raw_inner_xml(), timing_inner_xml);
        assert_eq!(timing.animations().len(), 2);
        assert_eq!(timing.animations()[0].id(), 1);
        assert_eq!(timing.animations()[0].duration_ms(), None);
        assert_eq!(timing.animations()[0].trigger(), Some("tmRoot"));
        assert_eq!(timing.animations()[0].event(), None);
        assert_eq!(timing.animations()[1].id(), 2);
        assert_eq!(timing.animations()[1].duration_ms(), Some(350));
        assert_eq!(timing.animations()[1].trigger(), Some("clickEffect"));
        assert_eq!(timing.animations()[1].event(), Some("cancelBubble"));
    }

    #[test]
    fn set_raw_inner_xml_resyncs_typed_nodes() {
        let mut timing = SlideTiming::new("<p:tnLst><p:par><p:cTn id=\"1\"/></p:par></p:tnLst>");
        assert_eq!(timing.animations().len(), 1);

        timing
            .set_raw_inner_xml("<p:tnLst><p:par><p:cTn id=\"9\" dur=\"1200\"/></p:par></p:tnLst>");
        assert_eq!(
            timing.raw_inner_xml(),
            "<p:tnLst><p:par><p:cTn id=\"9\" dur=\"1200\"/></p:par></p:tnLst>"
        );
        assert_eq!(timing.animations().len(), 1);
        assert_eq!(timing.animations()[0].id(), 9);
        assert_eq!(timing.animations()[0].duration_ms(), Some(1_200));
    }

    #[test]
    fn animations_mut_clears_raw_inner_xml_and_uses_typed_nodes() {
        let mut timing = SlideTiming::new("<p:tnLst><p:par><p:cTn id=\"1\"/></p:par></p:tnLst>");
        assert!(!timing.raw_inner_xml().is_empty());

        let mut new_node = SlideAnimationNode::new(42);
        new_node.set_duration_ms(Some(500));
        new_node.set_trigger(Some("clickEffect"));
        new_node.set_event(Some("onClick"));
        timing.animations_mut().push(new_node.clone());

        assert_eq!(timing.raw_inner_xml(), "");
        assert_eq!(timing.animations().last(), Some(&new_node));
    }
}
