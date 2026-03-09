/// Transition speed (`spd` attribute on `<p:transition>`).
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum TransitionSpeed {
    Slow,
    Med,
    Fast,
}

impl TransitionSpeed {
    pub fn from_xml(value: &str) -> Option<Self> {
        match value {
            "slow" => Some(Self::Slow),
            "med" => Some(Self::Med),
            "fast" => Some(Self::Fast),
            _ => None,
        }
    }

    pub fn to_xml(self) -> &'static str {
        match self {
            Self::Slow => "slow",
            Self::Med => "med",
            Self::Fast => "fast",
        }
    }
}

/// Transition sound action (`<p:sndAc>`).
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct TransitionSound {
    /// Sound name.
    pub name: String,
    /// Relationship ID referencing the sound media part.
    pub relationship_id: String,
}

impl TransitionSound {
    pub fn new(name: impl Into<String>, relationship_id: impl Into<String>) -> Self {
        Self {
            name: name.into(),
            relationship_id: relationship_id.into(),
        }
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct SlideTransition {
    kind: SlideTransitionKind,
    advance_on_click: Option<bool>,
    advance_after_ms: Option<u32>,
    /// Transition speed (`spd` attribute).
    speed: Option<TransitionSpeed>,
    /// Transition sound action (`<p:sndAc>`).
    sound: Option<TransitionSound>,
}

impl SlideTransition {
    pub fn new(kind: SlideTransitionKind) -> Self {
        Self {
            kind,
            advance_on_click: None,
            advance_after_ms: None,
            speed: None,
            sound: None,
        }
    }

    pub fn kind(&self) -> &SlideTransitionKind {
        &self.kind
    }

    pub fn set_kind(&mut self, kind: SlideTransitionKind) {
        self.kind = kind;
    }

    pub fn advance_on_click(&self) -> Option<bool> {
        self.advance_on_click
    }

    pub fn set_advance_on_click(&mut self, advance_on_click: Option<bool>) {
        self.advance_on_click = advance_on_click;
    }

    pub fn advance_after_ms(&self) -> Option<u32> {
        self.advance_after_ms
    }

    pub fn set_advance_after_ms(&mut self, advance_after_ms: Option<u32>) {
        self.advance_after_ms = advance_after_ms;
    }

    /// Transition speed.
    pub fn speed(&self) -> Option<TransitionSpeed> {
        self.speed
    }

    /// Set transition speed.
    pub fn set_speed(&mut self, speed: Option<TransitionSpeed>) {
        self.speed = speed;
    }

    /// Transition sound.
    pub fn sound(&self) -> Option<&TransitionSound> {
        self.sound.as_ref()
    }

    /// Set transition sound.
    pub fn set_sound(&mut self, sound: Option<TransitionSound>) {
        self.sound = sound;
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub enum SlideTransitionKind {
    Unspecified,
    Cut,
    Fade,
    Push,
    Wipe,
    Other(String),
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn stores_transition_metadata() {
        let mut transition = SlideTransition::new(SlideTransitionKind::Fade);
        transition.set_advance_on_click(Some(false));
        transition.set_advance_after_ms(Some(2_000));

        assert_eq!(transition.kind(), &SlideTransitionKind::Fade);
        assert_eq!(transition.advance_on_click(), Some(false));
        assert_eq!(transition.advance_after_ms(), Some(2_000));
    }

    #[test]
    fn transition_speed_xml_roundtrip() {
        for (xml, expected) in [
            ("slow", TransitionSpeed::Slow),
            ("med", TransitionSpeed::Med),
            ("fast", TransitionSpeed::Fast),
        ] {
            assert_eq!(TransitionSpeed::from_xml(xml), Some(expected));
            assert_eq!(expected.to_xml(), xml);
        }
        assert_eq!(TransitionSpeed::from_xml("unknown"), None);
    }

    #[test]
    fn transition_speed_on_slide_transition() {
        let mut transition = SlideTransition::new(SlideTransitionKind::Fade);
        assert_eq!(transition.speed(), None);

        transition.set_speed(Some(TransitionSpeed::Slow));
        assert_eq!(transition.speed(), Some(TransitionSpeed::Slow));

        transition.set_speed(None);
        assert_eq!(transition.speed(), None);
    }

    #[test]
    fn transition_sound_roundtrip() {
        let mut transition = SlideTransition::new(SlideTransitionKind::Cut);
        assert!(transition.sound().is_none());

        let sound = TransitionSound::new("Chime", "rId5");
        transition.set_sound(Some(sound));

        let s = transition.sound().unwrap();
        assert_eq!(s.name, "Chime");
        assert_eq!(s.relationship_id, "rId5");

        transition.set_sound(None);
        assert!(transition.sound().is_none());
    }
}
