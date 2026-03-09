//! Advanced animation controls and sequence management.
//!
//! This module extends the basic `timing` module with advanced animation features:
//! - Animation triggers and event handlers
//! - Animation sequence management
//! - Timing properties (delay, duration, acceleration)
//! - Animation effects and presets
//!
//! Ported from ShapeCrawler's animation system concepts.

use std::io::Cursor;

use quick_xml::events::{BytesStart, Event};
use quick_xml::Reader;

/// Animation trigger type (when an animation should start).
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum AnimationTrigger {
    /// Start on click.
    OnClick,
    /// Start with previous animation.
    WithPrevious,
    /// Start after previous animation.
    AfterPrevious,
    /// Start on shape click.
    OnShapeClick,
    /// Start based on timing.
    Timing,
}

impl AnimationTrigger {
    pub fn from_xml(value: &str) -> Option<Self> {
        match value {
            "onClick" | "clickEffect" => Some(Self::OnClick),
            "withPrev" | "withPrevious" => Some(Self::WithPrevious),
            "afterPrev" | "afterPrevious" => Some(Self::AfterPrevious),
            "onShapeClick" | "shapeClick" => Some(Self::OnShapeClick),
            "tmRoot" | "timing" => Some(Self::Timing),
            _ => None,
        }
    }

    pub fn to_xml(self) -> &'static str {
        match self {
            Self::OnClick => "onClick",
            Self::WithPrevious => "withPrev",
            Self::AfterPrevious => "afterPrev",
            Self::OnShapeClick => "onShapeClick",
            Self::Timing => "tmRoot",
        }
    }
}

/// Animation effect type (the visual effect applied).
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum AnimationEffectType {
    /// Entrance effect (appear).
    Entrance,
    /// Emphasis effect (highlight).
    Emphasis,
    /// Exit effect (disappear).
    Exit,
    /// Motion path effect.
    MotionPath,
}

impl AnimationEffectType {
    pub fn from_xml(value: &str) -> Option<Self> {
        match value {
            "entrance" | "in" => Some(Self::Entrance),
            "emphasis" | "emph" => Some(Self::Emphasis),
            "exit" | "out" => Some(Self::Exit),
            "path" | "motionPath" => Some(Self::MotionPath),
            _ => None,
        }
    }

    pub fn to_xml(self) -> &'static str {
        match self {
            Self::Entrance => "entrance",
            Self::Emphasis => "emphasis",
            Self::Exit => "exit",
            Self::MotionPath => "path",
        }
    }
}

/// Animation restart behavior.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum AnimationRestart {
    /// Always restart.
    Always,
    /// Never restart.
    Never,
    /// Restart when not active.
    WhenNotActive,
}

impl AnimationRestart {
    pub fn from_xml(value: &str) -> Option<Self> {
        match value {
            "always" => Some(Self::Always),
            "never" => Some(Self::Never),
            "whenNotActive" => Some(Self::WhenNotActive),
            _ => None,
        }
    }

    pub fn to_xml(self) -> &'static str {
        match self {
            Self::Always => "always",
            Self::Never => "never",
            Self::WhenNotActive => "whenNotActive",
        }
    }
}

/// Animation timing properties (delay, duration, acceleration, etc.).
#[derive(Debug, Clone, PartialEq)]
pub struct AnimationTiming {
    /// Start delay in milliseconds.
    delay_ms: u64,
    /// Duration in milliseconds (None = indefinite).
    duration_ms: Option<u64>,
    /// Acceleration ratio (0.0 - 1.0).
    accel: f64,
    /// Deceleration ratio (0.0 - 1.0).
    decel: f64,
    /// Restart behavior.
    restart: AnimationRestart,
    /// Auto-reverse flag.
    auto_reverse: bool,
    /// Repeat count (None = 1, Some(0) = infinite).
    repeat_count: Option<u32>,
}

impl Default for AnimationTiming {
    fn default() -> Self {
        Self {
            delay_ms: 0,
            duration_ms: Some(500),
            accel: 0.0,
            decel: 0.0,
            restart: AnimationRestart::Never,
            auto_reverse: false,
            repeat_count: None,
        }
    }
}

impl AnimationTiming {
    pub fn new() -> Self {
        Self::default()
    }

    pub fn delay_ms(&self) -> u64 {
        self.delay_ms
    }

    pub fn set_delay_ms(&mut self, delay_ms: u64) {
        self.delay_ms = delay_ms;
    }

    pub fn duration_ms(&self) -> Option<u64> {
        self.duration_ms
    }

    pub fn set_duration_ms(&mut self, duration_ms: Option<u64>) {
        self.duration_ms = duration_ms;
    }

    pub fn accel(&self) -> f64 {
        self.accel
    }

    /// Set acceleration ratio (clamped to 0.0 - 1.0).
    pub fn set_accel(&mut self, accel: f64) {
        self.accel = accel.clamp(0.0, 1.0);
    }

    pub fn decel(&self) -> f64 {
        self.decel
    }

    /// Set deceleration ratio (clamped to 0.0 - 1.0).
    pub fn set_decel(&mut self, decel: f64) {
        self.decel = decel.clamp(0.0, 1.0);
    }

    pub fn restart(&self) -> AnimationRestart {
        self.restart
    }

    pub fn set_restart(&mut self, restart: AnimationRestart) {
        self.restart = restart;
    }

    pub fn auto_reverse(&self) -> bool {
        self.auto_reverse
    }

    pub fn set_auto_reverse(&mut self, auto_reverse: bool) {
        self.auto_reverse = auto_reverse;
    }

    pub fn repeat_count(&self) -> Option<u32> {
        self.repeat_count
    }

    pub fn set_repeat_count(&mut self, repeat_count: Option<u32>) {
        self.repeat_count = repeat_count;
    }
}

/// Animation effect instance with timing and trigger information.
///
/// Represents a single animation effect applied to a shape or element.
#[derive(Debug, Clone, PartialEq)]
pub struct AnimationEffect {
    /// Animation node ID.
    id: u32,
    /// Target shape ID.
    target_shape_id: Option<u32>,
    /// Animation trigger.
    trigger: AnimationTrigger,
    /// Effect type.
    effect_type: AnimationEffectType,
    /// Timing properties.
    timing: AnimationTiming,
    /// Effect preset name (e.g., "Fade", "Wipe", "Fly In").
    preset_name: Option<String>,
}

impl AnimationEffect {
    pub fn new(id: u32, trigger: AnimationTrigger, effect_type: AnimationEffectType) -> Self {
        Self {
            id,
            target_shape_id: None,
            trigger,
            effect_type,
            timing: AnimationTiming::default(),
            preset_name: None,
        }
    }

    pub fn id(&self) -> u32 {
        self.id
    }

    pub fn set_id(&mut self, id: u32) {
        self.id = id;
    }

    pub fn target_shape_id(&self) -> Option<u32> {
        self.target_shape_id
    }

    pub fn set_target_shape_id(&mut self, target_shape_id: Option<u32>) {
        self.target_shape_id = target_shape_id;
    }

    pub fn trigger(&self) -> AnimationTrigger {
        self.trigger
    }

    pub fn set_trigger(&mut self, trigger: AnimationTrigger) {
        self.trigger = trigger;
    }

    pub fn effect_type(&self) -> AnimationEffectType {
        self.effect_type
    }

    pub fn set_effect_type(&mut self, effect_type: AnimationEffectType) {
        self.effect_type = effect_type;
    }

    pub fn timing(&self) -> &AnimationTiming {
        &self.timing
    }

    pub fn timing_mut(&mut self) -> &mut AnimationTiming {
        &mut self.timing
    }

    pub fn set_timing(&mut self, timing: AnimationTiming) {
        self.timing = timing;
    }

    pub fn preset_name(&self) -> Option<&str> {
        self.preset_name.as_deref()
    }

    pub fn set_preset_name(&mut self, preset_name: Option<impl Into<String>>) {
        self.preset_name = preset_name.map(Into::into);
    }
}

/// Animation sequence managing multiple animation effects.
///
/// An animation sequence groups related animation effects and provides
/// ordering and synchronization control.
#[derive(Debug, Clone, PartialEq)]
pub struct AnimationSequence {
    /// Sequence ID.
    id: u32,
    /// Animation effects in this sequence.
    effects: Vec<AnimationEffect>,
    /// Whether this sequence runs concurrently with others.
    concurrent: bool,
}

impl AnimationSequence {
    pub fn new(id: u32) -> Self {
        Self {
            id,
            effects: Vec::new(),
            concurrent: false,
        }
    }

    pub fn id(&self) -> u32 {
        self.id
    }

    pub fn set_id(&mut self, id: u32) {
        self.id = id;
    }

    pub fn effects(&self) -> &[AnimationEffect] {
        &self.effects
    }

    pub fn effects_mut(&mut self) -> &mut Vec<AnimationEffect> {
        &mut self.effects
    }

    pub fn add_effect(&mut self, effect: AnimationEffect) {
        self.effects.push(effect);
    }

    pub fn remove_effect(&mut self, effect_id: u32) -> Option<AnimationEffect> {
        self.effects
            .iter()
            .position(|e| e.id() == effect_id)
            .map(|idx| self.effects.remove(idx))
    }

    pub fn concurrent(&self) -> bool {
        self.concurrent
    }

    pub fn set_concurrent(&mut self, concurrent: bool) {
        self.concurrent = concurrent;
    }
}

/// Parse animation effects from PresentationML timing XML.
///
/// Extracts animation nodes from `<p:timing>` content, building typed
/// `AnimationEffect` instances with timing and trigger information.
pub fn parse_animation_effects(raw_timing_xml: &str) -> Vec<AnimationEffect> {
    if raw_timing_xml.is_empty() {
        return Vec::new();
    }

    let mut reader = Reader::from_reader(Cursor::new(raw_timing_xml.as_bytes()));
    reader.config_mut().trim_text(false);
    let mut buffer = Vec::new();
    let mut effects = Vec::new();

    loop {
        match reader.read_event_into(&mut buffer) {
            Ok(Event::Start(ref event)) | Ok(Event::Empty(ref event)) => {
                let name_bytes = event.name();
                let local = local_name(name_bytes.as_ref());

                // Parse animation effect nodes
                if local == b"animEffect" || local == b"anim" || local == b"set" {
                    if let Some(effect) = parse_animation_effect_node(event) {
                        effects.push(effect);
                    }
                }
            }
            Ok(Event::Eof) => break,
            Ok(_) => {}
            Err(_) => break,
        }
        buffer.clear();
    }

    effects
}

fn parse_animation_effect_node(event: &BytesStart<'_>) -> Option<AnimationEffect> {
    // Extract ID from parent cTn or direct attributes
    let id = get_attribute_value(event, b"id")
        .or_else(|| get_attribute_value(event, b"nodeId"))
        .and_then(|v| v.parse::<u32>().ok())
        .unwrap_or(0);

    // Determine trigger type
    let trigger_str = get_attribute_value(event, b"trigger")
        .or_else(|| get_attribute_value(event, b"nodeType"))
        .unwrap_or_else(|| "onClick".to_string());
    let trigger = AnimationTrigger::from_xml(&trigger_str).unwrap_or(AnimationTrigger::OnClick);

    // Determine effect type
    let effect_str = get_attribute_value(event, b"presetClass")
        .or_else(|| get_attribute_value(event, b"effectType"))
        .unwrap_or_else(|| "entrance".to_string());
    let effect_type =
        AnimationEffectType::from_xml(&effect_str).unwrap_or(AnimationEffectType::Entrance);

    let mut effect = AnimationEffect::new(id, trigger, effect_type);

    // Parse timing properties
    if let Some(dur) = get_attribute_value(event, b"dur") {
        if let Ok(duration) = dur.parse::<u64>() {
            effect.timing_mut().set_duration_ms(Some(duration));
        } else if dur == "indefinite" {
            effect.timing_mut().set_duration_ms(None);
        }
    }

    if let Some(delay) = get_attribute_value(event, b"delay") {
        if let Ok(delay_val) = delay.parse::<u64>() {
            effect.timing_mut().set_delay_ms(delay_val);
        }
    }

    if let Some(accel) = get_attribute_value(event, b"accel") {
        if let Ok(accel_val) = accel.parse::<f64>() {
            effect.timing_mut().set_accel(accel_val);
        }
    }

    if let Some(decel) = get_attribute_value(event, b"decel") {
        if let Ok(decel_val) = decel.parse::<f64>() {
            effect.timing_mut().set_decel(decel_val);
        }
    }

    if let Some(restart) = get_attribute_value(event, b"restart") {
        if let Some(restart_val) = AnimationRestart::from_xml(&restart) {
            effect.timing_mut().set_restart(restart_val);
        }
    }

    if let Some(auto_rev) = get_attribute_value(event, b"autoRev") {
        effect
            .timing_mut()
            .set_auto_reverse(&auto_rev == "1" || &auto_rev == "true");
    }

    // Parse target shape ID
    if let Some(spid) = get_attribute_value(event, b"spid") {
        if let Ok(shape_id) = spid.parse::<u32>() {
            effect.set_target_shape_id(Some(shape_id));
        }
    }

    // Parse preset name
    if let Some(preset) = get_attribute_value(event, b"presetID") {
        effect.set_preset_name(Some(preset));
    }

    Some(effect)
}

// ── Helper functions ──

fn local_name(name: &[u8]) -> &[u8] {
    name.rsplit(|byte| *byte == b':').next().unwrap_or(name)
}

fn get_attribute_value(event: &BytesStart<'_>, expected_local_name: &[u8]) -> Option<String> {
    event.attributes().flatten().find_map(|attribute| {
        (local_name(attribute.key.as_ref()) == expected_local_name)
            .then(|| String::from_utf8_lossy(attribute.value.as_ref()).into_owned())
    })
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn animation_trigger_xml_roundtrip() {
        for (xml, expected) in [
            ("onClick", AnimationTrigger::OnClick),
            ("clickEffect", AnimationTrigger::OnClick),
            ("withPrev", AnimationTrigger::WithPrevious),
            ("afterPrev", AnimationTrigger::AfterPrevious),
            ("onShapeClick", AnimationTrigger::OnShapeClick),
            ("tmRoot", AnimationTrigger::Timing),
        ] {
            assert_eq!(AnimationTrigger::from_xml(xml), Some(expected));
        }
        assert_eq!(AnimationTrigger::from_xml("unknown"), None);
    }

    #[test]
    fn animation_effect_type_xml_roundtrip() {
        for (xml, expected) in [
            ("entrance", AnimationEffectType::Entrance),
            ("in", AnimationEffectType::Entrance),
            ("emphasis", AnimationEffectType::Emphasis),
            ("exit", AnimationEffectType::Exit),
            ("path", AnimationEffectType::MotionPath),
        ] {
            assert_eq!(AnimationEffectType::from_xml(xml), Some(expected));
        }
        assert_eq!(AnimationEffectType::from_xml("unknown"), None);
    }

    #[test]
    fn animation_timing_default() {
        let timing = AnimationTiming::default();
        assert_eq!(timing.delay_ms(), 0);
        assert_eq!(timing.duration_ms(), Some(500));
        assert_eq!(timing.accel(), 0.0);
        assert_eq!(timing.decel(), 0.0);
        assert_eq!(timing.restart(), AnimationRestart::Never);
        assert!(!timing.auto_reverse());
        assert_eq!(timing.repeat_count(), None);
    }

    #[test]
    fn animation_timing_setters() {
        let mut timing = AnimationTiming::new();

        timing.set_delay_ms(200);
        timing.set_duration_ms(Some(1000));
        timing.set_accel(0.3);
        timing.set_decel(0.4);
        timing.set_restart(AnimationRestart::Always);
        timing.set_auto_reverse(true);
        timing.set_repeat_count(Some(2));

        assert_eq!(timing.delay_ms(), 200);
        assert_eq!(timing.duration_ms(), Some(1000));
        assert_eq!(timing.accel(), 0.3);
        assert_eq!(timing.decel(), 0.4);
        assert_eq!(timing.restart(), AnimationRestart::Always);
        assert!(timing.auto_reverse());
        assert_eq!(timing.repeat_count(), Some(2));
    }

    #[test]
    fn animation_timing_accel_decel_clamp() {
        let mut timing = AnimationTiming::new();

        timing.set_accel(1.5);
        assert_eq!(timing.accel(), 1.0);

        timing.set_accel(-0.5);
        assert_eq!(timing.accel(), 0.0);

        timing.set_decel(2.0);
        assert_eq!(timing.decel(), 1.0);

        timing.set_decel(-1.0);
        assert_eq!(timing.decel(), 0.0);
    }

    #[test]
    fn animation_effect_roundtrip() {
        let mut effect =
            AnimationEffect::new(42, AnimationTrigger::OnClick, AnimationEffectType::Entrance);
        effect.set_target_shape_id(Some(100));
        effect.timing_mut().set_duration_ms(Some(750));
        effect.timing_mut().set_delay_ms(100);
        effect.set_preset_name(Some("Fade"));

        assert_eq!(effect.id(), 42);
        assert_eq!(effect.target_shape_id(), Some(100));
        assert_eq!(effect.trigger(), AnimationTrigger::OnClick);
        assert_eq!(effect.effect_type(), AnimationEffectType::Entrance);
        assert_eq!(effect.timing().duration_ms(), Some(750));
        assert_eq!(effect.timing().delay_ms(), 100);
        assert_eq!(effect.preset_name(), Some("Fade"));
    }

    #[test]
    fn animation_sequence_add_remove() {
        let mut seq = AnimationSequence::new(1);
        assert_eq!(seq.effects().len(), 0);

        let effect1 =
            AnimationEffect::new(10, AnimationTrigger::OnClick, AnimationEffectType::Entrance);
        let effect2 = AnimationEffect::new(
            20,
            AnimationTrigger::AfterPrevious,
            AnimationEffectType::Exit,
        );

        seq.add_effect(effect1);
        seq.add_effect(effect2);
        assert_eq!(seq.effects().len(), 2);

        let removed = seq.remove_effect(10);
        assert!(removed.is_some());
        assert_eq!(removed.unwrap().id(), 10);
        assert_eq!(seq.effects().len(), 1);
        assert_eq!(seq.effects()[0].id(), 20);
    }

    #[test]
    fn animation_sequence_concurrent() {
        let mut seq = AnimationSequence::new(5);
        assert!(!seq.concurrent());

        seq.set_concurrent(true);
        assert!(seq.concurrent());
    }

    #[test]
    fn parse_animation_effects_empty() {
        let effects = parse_animation_effects("");
        assert!(effects.is_empty());
    }

    #[test]
    fn parse_animation_effects_with_simple_node() {
        let xml = r#"<p:animEffect id="100" trigger="onClick" presetClass="entrance" dur="500"/>"#;
        let effects = parse_animation_effects(xml);

        assert_eq!(effects.len(), 1);
        assert_eq!(effects[0].id(), 100);
        assert_eq!(effects[0].trigger(), AnimationTrigger::OnClick);
        assert_eq!(effects[0].effect_type(), AnimationEffectType::Entrance);
        assert_eq!(effects[0].timing().duration_ms(), Some(500));
    }
}
