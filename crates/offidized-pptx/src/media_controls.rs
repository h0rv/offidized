//! Media playback controls for audio and video shapes.
//!
//! This module provides advanced media playback controls including:
//! - Audio/video start mode (click, auto, on click)
//! - Playback options (loop, volume, autoplay)
//! - Media file handling and MIME type detection
//!
//! Ported from ShapeCrawler's media content features.

use std::io::Cursor;

use quick_xml::events::{BytesStart, Event};
use quick_xml::Reader;

/// Audio start mode (how media should begin playback).
///
/// Maps to the `<p14:media>` element and related timing nodes in PresentationML.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum AudioStartMode {
    /// Start in click sequence (default PowerPoint behavior).
    InClickSequence,
    /// Start automatically when slide loads.
    Automatically,
    /// Start when the media shape is clicked.
    WhenClickedOn,
}

impl AudioStartMode {
    pub fn from_xml(value: &str) -> Option<Self> {
        match value {
            "inClickSeq" | "clickSequence" => Some(Self::InClickSequence),
            "auto" | "automatically" => Some(Self::Automatically),
            "whenClicked" | "onClick" => Some(Self::WhenClickedOn),
            _ => None,
        }
    }

    pub fn to_xml(self) -> &'static str {
        match self {
            Self::InClickSequence => "inClickSeq",
            Self::Automatically => "auto",
            Self::WhenClickedOn => "whenClicked",
        }
    }
}

/// Audio/video file type enumeration.
///
/// Common MIME types for embedded media in presentations.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum AudioType {
    /// MP3 audio (audio/mpeg).
    Mp3,
    /// WAV audio (audio/wav).
    Wave,
    /// Other audio format.
    OtherAudio,
}

impl AudioType {
    pub fn from_mime(mime: &str) -> Self {
        match mime {
            "audio/mpeg" | "audio/mp3" => Self::Mp3,
            "audio/wav" | "audio/wave" | "audio/x-wav" => Self::Wave,
            _ => Self::OtherAudio,
        }
    }

    pub fn to_mime(self) -> &'static str {
        match self {
            Self::Mp3 => "audio/mpeg",
            Self::Wave => "audio/wav",
            Self::OtherAudio => "application/octet-stream",
        }
    }
}

/// Video file type enumeration.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum VideoType {
    /// MP4 video (video/mp4).
    Mp4,
    /// AVI video (video/x-msvideo).
    Avi,
    /// WMV video (video/x-ms-wmv).
    Wmv,
    /// Other video format.
    OtherVideo,
}

impl VideoType {
    pub fn from_mime(mime: &str) -> Self {
        match mime {
            "video/mp4" => Self::Mp4,
            "video/x-msvideo" | "video/avi" => Self::Avi,
            "video/x-ms-wmv" => Self::Wmv,
            _ => Self::OtherVideo,
        }
    }

    pub fn to_mime(self) -> &'static str {
        match self {
            Self::Mp4 => "video/mp4",
            Self::Avi => "video/x-msvideo",
            Self::Wmv => "video/x-ms-wmv",
            Self::OtherVideo => "application/octet-stream",
        }
    }
}

/// Media playback controls for audio/video shapes.
///
/// Controls playback behavior like autoplay, loop, and volume.
/// These settings are stored in the `<p14:media>` element and related timing nodes.
#[derive(Debug, Clone, PartialEq)]
pub struct MediaPlaybackControls {
    /// Start mode (how playback begins).
    start_mode: AudioStartMode,
    /// Whether to loop playback.
    loop_playback: bool,
    /// Volume level (0-100, where 100 is full volume).
    volume: u8,
    /// Whether to hide media icon during presentation.
    hide_during_show: bool,
    /// Number of times to repeat (None = infinite when loop is true).
    repeat_count: Option<u32>,
}

impl Default for MediaPlaybackControls {
    fn default() -> Self {
        Self {
            start_mode: AudioStartMode::InClickSequence,
            loop_playback: false,
            volume: 100,
            hide_during_show: false,
            repeat_count: None,
        }
    }
}

impl MediaPlaybackControls {
    pub fn new() -> Self {
        Self::default()
    }

    pub fn start_mode(&self) -> AudioStartMode {
        self.start_mode
    }

    pub fn set_start_mode(&mut self, start_mode: AudioStartMode) {
        self.start_mode = start_mode;
    }

    pub fn loop_playback(&self) -> bool {
        self.loop_playback
    }

    pub fn set_loop_playback(&mut self, loop_playback: bool) {
        self.loop_playback = loop_playback;
    }

    pub fn volume(&self) -> u8 {
        self.volume
    }

    /// Set volume level (0-100).
    ///
    /// Values above 100 are clamped to 100.
    pub fn set_volume(&mut self, volume: u8) {
        self.volume = volume.min(100);
    }

    pub fn hide_during_show(&self) -> bool {
        self.hide_during_show
    }

    pub fn set_hide_during_show(&mut self, hide_during_show: bool) {
        self.hide_during_show = hide_during_show;
    }

    pub fn repeat_count(&self) -> Option<u32> {
        self.repeat_count
    }

    pub fn set_repeat_count(&mut self, repeat_count: Option<u32>) {
        self.repeat_count = repeat_count;
    }

    /// Parse media controls from `<p14:media>` XML fragment.
    pub fn from_xml(xml: &str) -> Self {
        let mut controls = Self::default();

        if xml.is_empty() {
            return controls;
        }

        let mut reader = Reader::from_reader(Cursor::new(xml.as_bytes()));
        reader.config_mut().trim_text(false);
        let mut buffer = Vec::new();

        loop {
            match reader.read_event_into(&mut buffer) {
                Ok(Event::Start(ref event)) | Ok(Event::Empty(ref event)) => {
                    let name_bytes = event.name();
                    let local = local_name(name_bytes.as_ref());

                    // Parse start mode from timing attributes
                    if local == b"cTn" {
                        if let Some(node_type) = get_attribute_value(event, b"nodeType")
                            .or_else(|| get_attribute_value(event, b"presetClass"))
                        {
                            if let Some(mode) = AudioStartMode::from_xml(&node_type) {
                                controls.set_start_mode(mode);
                            }
                        }
                    }

                    // Parse loop attribute
                    if local == b"cMediaNode" || local == b"video" || local == b"audio" {
                        if let Some(loop_val) = get_attribute_value(event, b"loop") {
                            controls.set_loop_playback(&loop_val == "1" || &loop_val == "true");
                        }

                        if let Some(vol_val) = get_attribute_value(event, b"vol") {
                            if let Ok(vol) = vol_val.parse::<u8>() {
                                controls.set_volume(vol);
                            }
                        }

                        if let Some(show_val) = get_attribute_value(event, b"showWhenStopped") {
                            controls.set_hide_during_show(&show_val == "0" || &show_val == "false");
                        }
                    }
                }
                Ok(Event::Eof) => break,
                Ok(_) => {}
                Err(_) => break,
            }
            buffer.clear();
        }

        controls
    }

    /// Serialize media controls to XML attributes.
    ///
    /// Returns attribute pairs suitable for `<p14:media>` or timing elements.
    pub fn to_xml_attributes(&self) -> Vec<(String, String)> {
        let mut attrs = Vec::new();

        attrs.push((
            "loop".to_string(),
            if self.loop_playback { "1" } else { "0" }.to_string(),
        ));
        attrs.push(("vol".to_string(), self.volume.to_string()));
        attrs.push((
            "showWhenStopped".to_string(),
            if self.hide_during_show { "0" } else { "1" }.to_string(),
        ));

        if let Some(count) = self.repeat_count {
            attrs.push(("repeatCount".to_string(), count.to_string()));
        }

        attrs
    }
}

/// Media content wrapper for audio/video embedded in presentations.
///
/// Stores the media relationship ID, MIME type, and playback controls.
#[derive(Debug, Clone, PartialEq)]
pub struct MediaContent {
    /// Relationship ID pointing to the media data part.
    relationship_id: String,
    /// MIME type of the media.
    mime_type: String,
    /// Playback controls.
    playback_controls: MediaPlaybackControls,
}

impl MediaContent {
    pub fn new(relationship_id: impl Into<String>, mime_type: impl Into<String>) -> Self {
        Self {
            relationship_id: relationship_id.into(),
            mime_type: mime_type.into(),
            playback_controls: MediaPlaybackControls::default(),
        }
    }

    pub fn relationship_id(&self) -> &str {
        &self.relationship_id
    }

    pub fn set_relationship_id(&mut self, relationship_id: impl Into<String>) {
        self.relationship_id = relationship_id.into();
    }

    pub fn mime_type(&self) -> &str {
        &self.mime_type
    }

    pub fn set_mime_type(&mut self, mime_type: impl Into<String>) {
        self.mime_type = mime_type.into();
    }

    pub fn playback_controls(&self) -> &MediaPlaybackControls {
        &self.playback_controls
    }

    pub fn playback_controls_mut(&mut self) -> &mut MediaPlaybackControls {
        &mut self.playback_controls
    }

    pub fn set_playback_controls(&mut self, playback_controls: MediaPlaybackControls) {
        self.playback_controls = playback_controls;
    }
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
    fn audio_start_mode_xml_roundtrip() {
        for (xml, expected) in [
            ("inClickSeq", AudioStartMode::InClickSequence),
            ("clickSequence", AudioStartMode::InClickSequence),
            ("auto", AudioStartMode::Automatically),
            ("automatically", AudioStartMode::Automatically),
            ("whenClicked", AudioStartMode::WhenClickedOn),
            ("onClick", AudioStartMode::WhenClickedOn),
        ] {
            assert_eq!(AudioStartMode::from_xml(xml), Some(expected));
        }
        assert_eq!(AudioStartMode::from_xml("unknown"), None);
    }

    #[test]
    fn audio_type_from_mime() {
        assert_eq!(AudioType::from_mime("audio/mpeg"), AudioType::Mp3);
        assert_eq!(AudioType::from_mime("audio/mp3"), AudioType::Mp3);
        assert_eq!(AudioType::from_mime("audio/wav"), AudioType::Wave);
        assert_eq!(AudioType::from_mime("audio/wave"), AudioType::Wave);
        assert_eq!(AudioType::from_mime("audio/ogg"), AudioType::OtherAudio);
    }

    #[test]
    fn video_type_from_mime() {
        assert_eq!(VideoType::from_mime("video/mp4"), VideoType::Mp4);
        assert_eq!(VideoType::from_mime("video/x-msvideo"), VideoType::Avi);
        assert_eq!(VideoType::from_mime("video/x-ms-wmv"), VideoType::Wmv);
        assert_eq!(VideoType::from_mime("video/webm"), VideoType::OtherVideo);
    }

    #[test]
    fn media_playback_controls_default() {
        let controls = MediaPlaybackControls::default();
        assert_eq!(controls.start_mode(), AudioStartMode::InClickSequence);
        assert!(!controls.loop_playback());
        assert_eq!(controls.volume(), 100);
        assert!(!controls.hide_during_show());
        assert_eq!(controls.repeat_count(), None);
    }

    #[test]
    fn media_playback_controls_setters() {
        let mut controls = MediaPlaybackControls::new();

        controls.set_start_mode(AudioStartMode::Automatically);
        controls.set_loop_playback(true);
        controls.set_volume(75);
        controls.set_hide_during_show(true);
        controls.set_repeat_count(Some(3));

        assert_eq!(controls.start_mode(), AudioStartMode::Automatically);
        assert!(controls.loop_playback());
        assert_eq!(controls.volume(), 75);
        assert!(controls.hide_during_show());
        assert_eq!(controls.repeat_count(), Some(3));
    }

    #[test]
    fn media_playback_controls_volume_clamp() {
        let mut controls = MediaPlaybackControls::new();
        controls.set_volume(150);
        assert_eq!(controls.volume(), 100);
    }

    #[test]
    fn media_playback_controls_to_xml_attributes() {
        let mut controls = MediaPlaybackControls::new();
        controls.set_loop_playback(true);
        controls.set_volume(80);
        controls.set_hide_during_show(false);
        controls.set_repeat_count(Some(5));

        let attrs = controls.to_xml_attributes();

        assert!(attrs.contains(&("loop".to_string(), "1".to_string())));
        assert!(attrs.contains(&("vol".to_string(), "80".to_string())));
        assert!(attrs.contains(&("showWhenStopped".to_string(), "1".to_string())));
        assert!(attrs.contains(&("repeatCount".to_string(), "5".to_string())));
    }

    #[test]
    fn media_content_roundtrip() {
        let mut media = MediaContent::new("rId5", "video/mp4");
        media.playback_controls_mut().set_loop_playback(true);
        media.playback_controls_mut().set_volume(90);

        assert_eq!(media.relationship_id(), "rId5");
        assert_eq!(media.mime_type(), "video/mp4");
        assert!(media.playback_controls().loop_playback());
        assert_eq!(media.playback_controls().volume(), 90);
    }

    #[test]
    fn media_playback_controls_from_xml_empty() {
        let controls = MediaPlaybackControls::from_xml("");
        assert_eq!(controls, MediaPlaybackControls::default());
    }

    #[test]
    fn media_playback_controls_from_xml_with_attributes() {
        let xml = r#"<p:cMediaNode loop="1" vol="50" showWhenStopped="0"/>"#;
        let controls = MediaPlaybackControls::from_xml(xml);

        assert!(controls.loop_playback());
        assert_eq!(controls.volume(), 50);
        assert!(controls.hide_during_show());
    }
}
