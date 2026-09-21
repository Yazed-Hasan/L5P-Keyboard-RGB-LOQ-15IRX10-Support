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
        #[serde(default)]
        idle_brightness: Option<u8>,
        bass: f32,
        mid: f32,
        treble: f32,
        presence: f32,
        #[serde(default = "default_audio_squelch")]
        squelch: f32,
        #[serde(default = "default_audio_punch")]
        punch: f32,
        #[serde(default = "default_audio_contrast")]
        contrast: f32,
        #[serde(default = "default_audio_spread")]
        spread: f32,
        #[serde(default = "default_audio_hue_shift")]
        hue_shift: f32,
        #[serde(default)]
        color_ramp: f32,
        #[serde(default = "default_audio_motion")]
        motion: f32,
        #[serde(default)]
        follow_system_volume: bool,
        #[serde(default)]
        ripple_color: bool,
        #[serde(default = "default_ripple_strength")]
        ripple_strength: f32,
        #[serde(default = "default_ripple_speed")]
        ripple_speed: f32,
        #[serde(default = "default_ripple_width")]
        ripple_width: f32,
        #[serde(default = "default_ripple_twist")]
        ripple_twist: f32,
        #[serde(default)]
        ripple_origin: RippleOrigin,
        #[serde(default)]
        ripple_tint: RippleTint,
        #[serde(default = "default_ripple_rgb")]
        ripple_rgb: [u8; 3],
        #[serde(default)]
        ripple_kind: RippleKind,
        #[serde(default)]
        ripple_trigger: RippleTrigger,
        #[serde(default)]
        ripple_shockwave: bool,
        #[serde(default = "default_ripple_shock_strength")]
        ripple_shock_strength: f32,
        #[serde(default = "default_ripple_shock_sensitivity")]
        ripple_shock_sensitivity: f32,
        color_mode: AudioColorMode,
        style: AudioStyle,
    },
    Stars {
        #[serde(default)]
        params: StarsParams,
    },
    Rain {
        #[serde(default)]
        params: RainParams,
    },
    Aurora {
        #[serde(default)]
        params: AuroraParams,
    },
    Scanner {
        #[serde(default)]
        params: ScannerParams,
    },
    Battery {
        #[serde(default)]
        params: BatteryParams,
    },
}

fn default_audio_squelch() -> f32 {
    0.07
}

fn default_audio_punch() -> f32 {
    0.65
}

fn default_audio_contrast() -> f32 {
    1.15
}

fn default_audio_spread() -> f32 {
    0.35
}

fn default_audio_hue_shift() -> f32 {
    0.25
}

fn default_audio_motion() -> f32 {
    0.7
}

fn default_ripple_strength() -> f32 {
    1.0
}

fn default_ripple_speed() -> f32 {
    1.0
}

fn default_ripple_width() -> f32 {
    0.45
}

fn default_ripple_twist() -> f32 {
    0.7
}

fn default_ripple_shock_strength() -> f32 {
    1.25
}

fn default_ripple_shock_sensitivity() -> f32 {
    0.55
}

fn default_ripple_rgb() -> [u8; 3] {
    [255, 48, 96]
}

#[derive(Default, Debug, Clone, Copy, Serialize, Deserialize, EnumIter, EnumString, PartialEq, Eq)]
pub enum RippleOrigin {
    #[default]
    Auto,
    Center,
    Left,
    Right,
}

#[derive(Default, Debug, Clone, Copy, Serialize, Deserialize, EnumIter, EnumString, PartialEq, Eq)]
pub enum RippleTint {
    #[default]
    ColorChange,
    Current,
    Custom,
    Rainbow,
}

#[derive(Default, Debug, Clone, Copy, Serialize, Deserialize, EnumIter, EnumString, PartialEq, Eq)]
pub enum RippleKind {
    #[default]
    Ring,
    Wave,
    Pulse,
    Double,
    Fill,
    Echo,
}

#[derive(Default, Debug, Clone, Copy, Serialize, Deserialize, EnumIter, EnumString, PartialEq, Eq)]
pub enum RippleTrigger {
    #[default]
    All,
    Bass,
    Kick,
    TripleKick,
}

#[derive(Default, Debug, Clone, Copy, Serialize, Deserialize, EnumIter, EnumString, PartialEq, Eq)]
pub enum AudioColorMode {
    #[default]
    #[serde(alias = "Profile")]
    #[strum(serialize = "Custom")]
    Custom,
    Spectrum,
    Rainbow,
    Heat,
    Ice,
    Sunset,
    Energy,
    Mono,
    Ramp,
}

#[derive(Default, Debug, Clone, Copy, Serialize, Deserialize, EnumIter, EnumString, PartialEq, Eq)]
pub enum AudioStyle {
    #[default]
    Levels,
    Pulse,
    Wave,
    Bloom,
    Center,
    Mirror,
    Fire,
    Strobe,
    Sparkle,
    Chase,
    Gradient,
    BeatGates,
    Vu,
    TempoPulse,
    Ripple,
}

#[derive(Default, Debug, Clone, Copy, Serialize, Deserialize, EnumIter, EnumString, PartialEq)]
pub enum SwipeMode {
    #[default]
    Change,
    Fill,
}

#[derive(Default, Debug, Clone, Copy, Serialize, Deserialize, EnumIter, EnumString, PartialEq, Eq)]
pub enum StarsPalette {
    #[default]
    Custom,
    White,
    Gold,
    Rainbow,
    Random,
}

#[derive(Debug, Clone, Copy, Serialize, Deserialize, PartialEq)]
#[serde(default)]
pub struct StarsParams {
    pub density: f32,
    pub twinkle: f32,
    pub size: f32,
    pub background: f32,
    pub shooting: f32,
    pub hue_drift: f32,
    pub palette: StarsPalette,
}

impl Default for StarsParams {
    fn default() -> Self {
        Self {
            density: 0.45,
            twinkle: 0.7,
            size: 0.35,
            background: 0.06,
            shooting: 0.08,
            hue_drift: 0.15,
            palette: StarsPalette::Custom,
        }
    }
}

impl StarsParams {
    pub fn normalized(self) -> Self {
        Self {
            density: self.density.clamp(0.05, 1.0),
            twinkle: self.twinkle.clamp(0.2, 2.0),
            size: self.size.clamp(0.05, 1.0),
            background: self.background.clamp(0.0, 0.4),
            shooting: self.shooting.clamp(0.0, 0.3),
            hue_drift: self.hue_drift.clamp(0.0, 1.0),
            palette: self.palette,
        }
    }
}

#[derive(Default, Debug, Clone, Copy, Serialize, Deserialize, EnumIter, EnumString, PartialEq, Eq)]
pub enum RainPalette {
    #[default]
    Ice,
    Custom,
    Neon,
    Rainbow,
}

#[derive(Debug, Clone, Copy, Serialize, Deserialize, PartialEq)]
#[serde(default)]
pub struct RainParams {
    pub density: f32,
    pub speed: f32,
    pub trail: f32,
    pub splash: f32,
    pub wind: f32,
    pub wet: f32,
    pub direction: Direction,
    pub palette: RainPalette,
}

impl Default for RainParams {
    fn default() -> Self {
        Self {
            density: 0.55,
            speed: 0.85,
            trail: 0.45,
            splash: 0.7,
            wind: 0.15,
            wet: 0.2,
            direction: Direction::Right,
            palette: RainPalette::Ice,
        }
    }
}

impl RainParams {
    pub fn normalized(self) -> Self {
        Self {
            density: self.density.clamp(0.05, 1.0),
            speed: self.speed.clamp(0.15, 2.5),
            trail: self.trail.clamp(0.05, 1.0),
            splash: self.splash.clamp(0.0, 1.0),
            wind: self.wind.clamp(0.0, 1.0),
            wet: self.wet.clamp(0.0, 1.0),
            direction: self.direction,
            palette: self.palette,
        }
    }
}

#[derive(Default, Debug, Clone, Copy, Serialize, Deserialize, EnumIter, EnumString, PartialEq, Eq)]
pub enum AuroraPalette {
    #[default]
    Borealis,
    Custom,
    Twilight,
    Rainbow,
}

#[derive(Debug, Clone, Copy, Serialize, Deserialize, PartialEq)]
#[serde(default)]
pub struct AuroraParams {
    pub layers: u8,
    pub speed: f32,
    pub wavelength: f32,
    pub contrast: f32,
    pub hue_drift: f32,
    pub softness: f32,
    pub brightness: f32,
    pub palette: AuroraPalette,
}

impl Default for AuroraParams {
    fn default() -> Self {
        Self {
            layers: 3,
            speed: 0.45,
            wavelength: 0.7,
            contrast: 1.1,
            hue_drift: 0.25,
            softness: 0.45,
            brightness: 0.85,
            palette: AuroraPalette::Borealis,
        }
    }
}

impl AuroraParams {
    pub fn normalized(self) -> Self {
        Self {
            layers: self.layers.clamp(1, 4),
            speed: self.speed.clamp(0.05, 2.0),
            wavelength: self.wavelength.clamp(0.15, 2.0),
            contrast: self.contrast.clamp(0.4, 2.2),
            hue_drift: self.hue_drift.clamp(0.0, 1.0),
            softness: self.softness.clamp(0.0, 1.0),
            brightness: self.brightness.clamp(0.15, 1.0),
            palette: self.palette,
        }
    }
}

#[derive(Default, Debug, Clone, Copy, Serialize, Deserialize, EnumIter, EnumString, PartialEq, Eq)]
pub enum ScannerPath {
    #[default]
    Bounce,
    Wrap,
}

#[derive(Default, Debug, Clone, Copy, Serialize, Deserialize, EnumIter, EnumString, PartialEq, Eq)]
pub enum ScannerPalette {
    #[default]
    Red,
    Custom,
    Ice,
    Rainbow,
}

#[derive(Debug, Clone, Copy, Serialize, Deserialize, PartialEq)]
#[serde(default)]
pub struct ScannerParams {
    pub width: f32,
    pub speed: f32,
    pub trail: f32,
    pub field: f32,
    pub dual: bool,
    pub path: ScannerPath,
    pub direction: Direction,
    pub palette: ScannerPalette,
}

impl Default for ScannerParams {
    fn default() -> Self {
        Self {
            width: 0.28,
            speed: 0.7,
            trail: 0.55,
            field: 0.08,
            dual: false,
            path: ScannerPath::Bounce,
            direction: Direction::Right,
            palette: ScannerPalette::Red,
        }
    }
}

impl ScannerParams {
    pub fn normalized(self) -> Self {
        Self {
            width: self.width.clamp(0.06, 0.8),
            speed: self.speed.clamp(0.1, 2.5),
            trail: self.trail.clamp(0.0, 1.0),
            field: self.field.clamp(0.0, 0.4),
            dual: self.dual,
            path: self.path,
            direction: self.direction,
            palette: self.palette,
        }
    }
}

#[derive(Default, Debug, Clone, Copy, Serialize, Deserialize, EnumIter, EnumString, PartialEq, Eq)]
pub enum BatteryPalette {
    #[default]
    Traffic,
    Custom,
    Ice,
    Heat,
}

#[derive(Debug, Clone, Copy, Serialize, Deserialize, PartialEq)]
#[serde(default)]
pub struct BatteryParams {
    pub low_pct: u8,
    pub mid_pct: u8,
    pub unused_dim: f32,
    pub pulse: f32,
    pub pulse_speed: f32,
    pub smoothing: f32,
    pub reverse: bool,
    pub palette: BatteryPalette,
}

impl Default for BatteryParams {
    fn default() -> Self {
        Self {
            low_pct: 15,
            mid_pct: 50,
            unused_dim: 0.06,
            pulse: 0.7,
            pulse_speed: 0.8,
            smoothing: 0.45,
            reverse: false,
            palette: BatteryPalette::Traffic,
        }
    }
}

impl BatteryParams {
    pub fn normalized(self) -> Self {
        let low = self.low_pct.clamp(5, 40);
        Self {
            low_pct: low,
            mid_pct: self.mid_pct.clamp(low.saturating_add(5), 90),
            unused_dim: self.unused_dim.clamp(0.0, 0.35),
            pulse: self.pulse.clamp(0.0, 1.5),
            pulse_speed: self.pulse_speed.clamp(0.15, 2.5),
            smoothing: self.smoothing.clamp(0.0, 0.95),
            reverse: self.reverse,
            palette: self.palette,
        }
    }
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
                    color_mode: AudioColorMode::Custom,
                    ..
                }
                | Self::AudioReact {
                    color_mode: AudioColorMode::Mono,
                    ..
                }
                | Self::Stars {
                    params: StarsParams {
                        palette: StarsPalette::Custom,
                        ..
                    },
                }
                | Self::Rain {
                    params: RainParams {
                        palette: RainPalette::Custom,
                        ..
                    },
                }
                | Self::Aurora {
                    params: AuroraParams {
                        palette: AuroraPalette::Custom,
                        ..
                    },
                }
                | Self::Scanner {
                    params: ScannerParams {
                        palette: ScannerPalette::Custom,
                        ..
                    },
                }
                | Self::Battery {
                    params: BatteryParams {
                        palette: BatteryPalette::Custom,
                        ..
                    },
                }
        )
    }

    pub fn takes_direction(self) -> bool {
        matches!(self, Self::Wave | Self::SmoothWave { .. } | Self::Swipe { .. })
    }

    pub fn is_scene(self) -> bool {
        matches!(
            self,
            Self::Stars { .. } | Self::Rain { .. } | Self::Aurora { .. } | Self::Scanner { .. } | Self::Battery { .. }
        )
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
            idle_brightness: Some(6),
            bass: 1.0,
            mid: 1.0,
            treble: 1.0,
            presence: 1.0,
            squelch: 0.07,
            punch: 0.65,
            contrast: 1.15,
            spread: 0.35,
            hue_shift: 0.25,
            color_ramp: 0.4,
            motion: 0.7,
            follow_system_volume: false,
            ripple_color: false,
            ripple_strength: 1.0,
            ripple_speed: 1.0,
            ripple_width: 0.45,
            ripple_twist: 0.7,
            ripple_origin: RippleOrigin::Auto,
            ripple_tint: RippleTint::ColorChange,
            ripple_rgb: [255, 48, 96],
            ripple_kind: RippleKind::Ring,
            ripple_trigger: RippleTrigger::All,
            ripple_shockwave: false,
            ripple_shock_strength: 1.25,
            ripple_shock_sensitivity: 0.55,
            color_mode: AudioColorMode::Custom,
            style: AudioStyle::Levels,
        }
    }

    pub fn ambient_default() -> Self {
        Self::AmbientLight {
            fps: 24,
            saturation_boost: 0.2,
        }
    }

    pub fn smooth_wave_default() -> Self {
        Self::SmoothWave {
            mode: SwipeMode::Change,
            clean_with_black: false,
        }
    }

    pub fn swipe_default() -> Self {
        Self::Swipe {
            mode: SwipeMode::Change,
            clean_with_black: false,
        }
    }

    pub fn stars_default() -> Self {
        Self::Stars {
            params: StarsParams::default(),
        }
    }

    pub fn rain_default() -> Self {
        Self::Rain {
            params: RainParams::default(),
        }
    }

    pub fn aurora_default() -> Self {
        Self::Aurora {
            params: AuroraParams::default(),
        }
    }

    pub fn scanner_default() -> Self {
        Self::Scanner {
            params: ScannerParams::default(),
        }
    }

    pub fn battery_default() -> Self {
        Self::Battery {
            params: BatteryParams::default(),
        }
    }

    pub fn factory_default(self) -> Self {
        match self {
            Self::AudioReact { .. } => Self::audio_react_default(),
            Self::AmbientLight { .. } => Self::ambient_default(),
            Self::SmoothWave { .. } => Self::smooth_wave_default(),
            Self::Swipe { .. } => Self::swipe_default(),
            Self::Stars { .. } => Self::stars_default(),
            Self::Rain { .. } => Self::rain_default(),
            Self::Aurora { .. } => Self::aurora_default(),
            Self::Scanner { .. } => Self::scanner_default(),
            Self::Battery { .. } => Self::battery_default(),
            other => other,
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
