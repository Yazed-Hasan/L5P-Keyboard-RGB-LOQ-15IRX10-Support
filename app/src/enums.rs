use crate::manager::{custom_effect::CustomEffect, profile::Profile};
use serde::{Deserialize, Serialize};
use strum_macros::{Display, EnumIter, EnumString, IntoStaticStr};

#[derive(Clone, Copy, EnumString, Serialize, Deserialize, Display, EnumIter, Debug, IntoStaticStr, Default)]
pub enum Effects {
    #[default]
    Static,
    Breath,
    Smooth,
    Wave,
    Lightning,
    AmbientLight {
        fps: u8,
        saturation_boost: f32,
    },
    SmoothWave {
        mode: SwipeMode,
        clean_with_black: bool,
    },
    Swipe {
        mode: SwipeMode,
        clean_with_black: bool,
    },
    Disco,
    Christmas,
    Fade,
    Temperature,
    Ripple,
    #[strum(serialize = "Audio React")]
    AudioReact {
        sensitivity: f32,
        smoothness: f32,
        min_brightness: u8,
        bass: f32,
        mid: f32,
        treble: f32,
        presence: f32,
        #[serde(default = "default_audio_squelch")]
        squelch: f32,
        #[serde(default = "default_audio_punch")]
        punch: f32,
        color_mode: AudioColorMode,
        style: AudioStyle,
    },
}

fn default_audio_squelch() -> f32 {
    0.07
}

fn default_audio_punch() -> f32 {
    0.65
}

#[derive(Default, Debug, Clone, Copy, Serialize, Deserialize, EnumIter, EnumString, PartialEq, Eq)]
pub enum AudioColorMode {
    Profile,
    #[default]
    Spectrum,
    Rainbow,
}

#[derive(Default, Debug, Clone, Copy, Serialize, Deserialize, EnumIter, EnumString, PartialEq, Eq)]
pub enum AudioStyle {
    #[default]
    Levels,
    Pulse,
    Wave,
    Bloom,
    Center,
}

#[derive(Default, Debug, Clone, Copy, Serialize, Deserialize, EnumIter, EnumString, PartialEq)]
pub enum SwipeMode {
    #[default]
    Change,
    Fill,
}

impl PartialEq for Effects {
    fn eq(&self, other: &Self) -> bool {
        core::mem::discriminant(self) == core::mem::discriminant(other)
    }
}

#[allow(dead_code)]
impl Effects {
    pub fn takes_color_array(self) -> bool {
        matches!(
            self,
            Self::Static
                | Self::Breath
                | Self::Lightning
                | Self::Swipe { .. }
                | Self::Fade
                | Self::Ripple
                | Self::AudioReact {
                    color_mode: AudioColorMode::Profile,
                    ..
                }
        )
    }

    pub fn takes_direction(self) -> bool {
        matches!(self, Self::Wave | Self::SmoothWave { .. } | Self::Swipe { .. })
    }

    pub fn takes_speed(self) -> bool {
        matches!(
            self,
            Self::Breath
                | Self::Smooth
                | Self::Wave
                | Self::Lightning
                | Self::SmoothWave { .. }
                | Self::Swipe { .. }
                | Self::Disco
                | Self::Fade
                | Self::Ripple
        )
    }

    pub fn is_built_in(self) -> bool {
        matches!(self, Self::Static | Self::Breath | Self::Smooth | Self::Wave)
    }

    pub fn audio_react_default() -> Self {
        Self::AudioReact {
            sensitivity: 1.2,
            smoothness: 0.72,
            min_brightness: 6,
            bass: 1.0,
            mid: 1.0,
            treble: 1.0,
            presence: 1.0,
            squelch: 0.07,
            punch: 0.65,
            color_mode: AudioColorMode::Spectrum,
            style: AudioStyle::Levels,
        }
    }
}

#[derive(Clone, Copy, EnumString, Serialize, Deserialize, Debug, EnumIter, IntoStaticStr, PartialEq, Eq, Default)]
pub enum Direction {
    #[default]
    Left,
    Right,
}

#[derive(PartialEq, Eq, EnumIter, IntoStaticStr, Clone, Copy, Default, Serialize, Deserialize, Debug, Display, EnumString)]
pub enum Brightness {
    #[default]
    Low,
    High,
}

#[derive(Debug)]
pub enum Message {
    CustomEffect { effect: CustomEffect },
    Profile { profile: Profile },
    Exit,
}
