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
        #[serde(default)]
        analysis: AudioAnalysis,
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
    #[strum(serialize = "Type Heat")]
    TypeHeat {
        #[serde(default)]
        params: TypeHeatParams,
    },
    Pacifica {
        #[serde(default)]
        params: PacificaParams,
    },
    #[strum(serialize = "Digital Rain")]
    DigitalRain {
        #[serde(default)]
        params: DigitalRainParams,
    },
    Fireworks {
        #[serde(default)]
        params: FireworksParams,
    },
    Nexus {
        #[serde(default)]
        params: NexusParams,
    },
    Comet {
        #[serde(default)]
        params: CometParams,
    },
    Juggle {
        #[serde(default)]
        params: JuggleParams,
    },
    #[strum(serialize = "Bouncing Balls")]
    Bounce {
        #[serde(default)]
        params: BounceParams,
    },
    Dissolve {
        #[serde(default)]
        params: DissolveParams,
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

fn default_comet_head() -> f32 {
    1.0
}

fn default_comet_glow() -> f32 {
    0.22
}

fn default_comet_fade() -> f32 {
    1.25
}

fn default_comet_wobble() -> f32 {
    0.35
}

fn default_comet_hue_speed() -> f32 {
    1.0
}

fn default_comet_sparkle() -> f32 {
    0.0
}

fn default_comet_saturation() -> f32 {
    1.0
}

fn default_comet_gap() -> f32 {
    0.46
}

fn default_comet_follow() -> f32 {
    0.62
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
    Oscilloscope,
    Spectrogram,
    Stereo,
    Pitch,
    Lissajous,
    Bubbles,
    KeyColor,
    MidSide,
    Eq24,
    PanNeedle,
    Collision,
    Snake,
    Ripple,
    Gravcenter,
    Melt,
    Wavelength,
}

#[derive(Default, Debug, Clone, Copy, Serialize, Deserialize, EnumIter, EnumString, PartialEq, Eq)]
pub enum AudioAnalysis {
    Auto,
    #[default]
    Classic,
    Accurate,
    Beats,
    Spectrum,
    Mel,
    Studio,
    Hpss,
    Complex,
    Tempo,
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

#[derive(Default, Debug, Clone, Copy, Serialize, Deserialize, EnumIter, EnumString, PartialEq, Eq)]
pub enum TypeHeatPalette {
    #[default]
    Heat,
    Ice,
    Custom,
}

#[derive(Debug, Clone, Copy, Serialize, Deserialize, PartialEq)]
#[serde(default)]
pub struct TypeHeatParams {
    pub heat: f32,
    pub cool: f32,
    pub hold_boost: f32,
    pub background: f32,
    pub palette: TypeHeatPalette,
}

impl Default for TypeHeatParams {
    fn default() -> Self {
        Self {
            heat: 0.28,
            cool: 0.55,
            hold_boost: 0.35,
            background: 0.04,
            palette: TypeHeatPalette::Heat,
        }
    }
}

impl TypeHeatParams {
    pub fn normalized(self) -> Self {
        Self {
            heat: self.heat.clamp(0.05, 0.8),
            cool: self.cool.clamp(0.1, 2.0),
            hold_boost: self.hold_boost.clamp(0.0, 1.2),
            background: self.background.clamp(0.0, 0.4),
            palette: self.palette,
        }
    }
}

#[derive(Default, Debug, Clone, Copy, Serialize, Deserialize, EnumIter, EnumString, PartialEq, Eq)]
pub enum PacificaPalette {
    #[default]
    Ocean,
    Ice,
    Custom,
}

#[derive(Debug, Clone, Copy, Serialize, Deserialize, PartialEq)]
#[serde(default)]
pub struct PacificaParams {
    pub speed: f32,
    pub intensity: f32,
    pub depth: f32,
    pub background: f32,
    pub palette: PacificaPalette,
}

impl Default for PacificaParams {
    fn default() -> Self {
        Self {
            speed: 0.55,
            intensity: 0.85,
            depth: 0.7,
            background: 0.08,
            palette: PacificaPalette::Ocean,
        }
    }
}

impl PacificaParams {
    pub fn normalized(self) -> Self {
        Self {
            speed: self.speed.clamp(0.08, 2.0),
            intensity: self.intensity.clamp(0.2, 1.0),
            depth: self.depth.clamp(0.15, 1.0),
            background: self.background.clamp(0.0, 0.4),
            palette: self.palette,
        }
    }
}

#[derive(Default, Debug, Clone, Copy, Serialize, Deserialize, EnumIter, EnumString, PartialEq, Eq)]
pub enum DigitalRainPalette {
    #[default]
    Matrix,
    Ice,
    Custom,
}

#[derive(Debug, Clone, Copy, Serialize, Deserialize, PartialEq)]
#[serde(default)]
pub struct DigitalRainParams {
    pub density: f32,
    pub speed: f32,
    pub trail: f32,
    pub background: f32,
    pub palette: DigitalRainPalette,
}

impl Default for DigitalRainParams {
    fn default() -> Self {
        Self {
            density: 0.55,
            speed: 0.9,
            trail: 0.55,
            background: 0.04,
            palette: DigitalRainPalette::Matrix,
        }
    }
}

impl DigitalRainParams {
    pub fn normalized(self) -> Self {
        Self {
            density: self.density.clamp(0.08, 1.0),
            speed: self.speed.clamp(0.15, 2.5),
            trail: self.trail.clamp(0.1, 1.0),
            background: self.background.clamp(0.0, 0.35),
            palette: self.palette,
        }
    }
}

#[derive(Default, Debug, Clone, Copy, Serialize, Deserialize, EnumIter, EnumString, PartialEq, Eq)]
pub enum FireworksPalette {
    #[default]
    Festival,
    Ice,
    Custom,
}

#[derive(Debug, Clone, Copy, Serialize, Deserialize, PartialEq)]
#[serde(default)]
pub struct FireworksParams {
    pub rate: f32,
    pub size: f32,
    pub trail: f32,
    pub background: f32,
    pub palette: FireworksPalette,
}

impl Default for FireworksParams {
    fn default() -> Self {
        Self {
            rate: 0.45,
            size: 0.4,
            trail: 0.55,
            background: 0.03,
            palette: FireworksPalette::Festival,
        }
    }
}

impl FireworksParams {
    pub fn normalized(self) -> Self {
        Self {
            rate: self.rate.clamp(0.05, 1.0),
            size: self.size.clamp(0.12, 1.0),
            trail: self.trail.clamp(0.1, 1.0),
            background: self.background.clamp(0.0, 0.3),
            palette: self.palette,
        }
    }
}

#[derive(Default, Debug, Clone, Copy, Serialize, Deserialize, EnumIter, EnumString, PartialEq, Eq)]
pub enum NexusPalette {
    #[default]
    Cyan,
    Heat,
    Ice,
    Custom,
}

#[derive(Debug, Clone, Copy, Serialize, Deserialize, PartialEq)]
#[serde(default)]
pub struct NexusParams {
    pub pulse: f32,
    pub fade: f32,
    pub cross: f32,
    pub background: f32,
    pub palette: NexusPalette,
}

impl Default for NexusParams {
    fn default() -> Self {
        Self {
            pulse: 1.0,
            fade: 0.7,
            cross: 0.75,
            background: 0.03,
            palette: NexusPalette::Cyan,
        }
    }
}

impl NexusParams {
    pub fn normalized(self) -> Self {
        Self {
            pulse: self.pulse.clamp(0.2, 1.5),
            fade: self.fade.clamp(0.15, 2.0),
            cross: self.cross.clamp(0.15, 1.0),
            background: self.background.clamp(0.0, 0.35),
            palette: self.palette,
        }
    }
}

#[derive(Default, Debug, Clone, Copy, Serialize, Deserialize, EnumIter, EnumString, PartialEq, Eq)]
pub enum CometPalette {
    #[default]
    Heat,
    Ice,
    Custom,
    Rainbow,
}

#[derive(Debug, Clone, Copy, Serialize, Deserialize, PartialEq)]
#[serde(default)]
pub struct CometParams {
    pub speed: f32,
    pub tail: f32,
    pub size: f32,
    pub background: f32,
    pub dual: bool,
    pub path: ScannerPath,
    pub direction: Direction,
    pub palette: CometPalette,
    #[serde(default = "default_comet_head")]
    pub head: f32,
    #[serde(default = "default_comet_glow")]
    pub glow: f32,
    #[serde(default = "default_comet_fade")]
    pub fade: f32,
    #[serde(default = "default_comet_wobble")]
    pub wobble: f32,
    #[serde(default = "default_comet_hue_speed")]
    pub hue_speed: f32,
    #[serde(default = "default_comet_sparkle")]
    pub sparkle: f32,
    #[serde(default = "default_comet_saturation")]
    pub saturation: f32,
    #[serde(default = "default_comet_gap")]
    pub gap: f32,
    #[serde(default = "default_comet_follow")]
    pub follow: f32,
    #[serde(default)]
    pub opposite: bool,
    #[serde(default)]
    pub triple: bool,
}

impl Default for CometParams {
    fn default() -> Self {
        Self {
            speed: 0.85,
            tail: 0.72,
            size: 0.45,
            background: 0.03,
            dual: false,
            path: ScannerPath::Wrap,
            direction: Direction::Right,
            palette: CometPalette::Heat,
            head: default_comet_head(),
            glow: default_comet_glow(),
            fade: default_comet_fade(),
            wobble: default_comet_wobble(),
            hue_speed: default_comet_hue_speed(),
            sparkle: default_comet_sparkle(),
            saturation: default_comet_saturation(),
            gap: default_comet_gap(),
            follow: default_comet_follow(),
            opposite: false,
            triple: false,
        }
    }
}

impl CometParams {
    pub fn normalized(self) -> Self {
        Self {
            speed: self.speed.clamp(0.08, 3.0),
            tail: self.tail.clamp(0.08, 1.0),
            size: self.size.clamp(0.08, 1.0),
            background: self.background.clamp(0.0, 0.45),
            dual: self.dual,
            path: self.path,
            direction: self.direction,
            palette: self.palette,
            head: self.head.clamp(0.25, 1.6),
            glow: self.glow.clamp(0.0, 1.0),
            fade: self.fade.clamp(0.45, 2.8),
            wobble: self.wobble.clamp(0.0, 1.0),
            hue_speed: self.hue_speed.clamp(0.0, 2.5),
            sparkle: self.sparkle.clamp(0.0, 1.0),
            saturation: self.saturation.clamp(0.15, 1.0),
            gap: self.gap.clamp(0.18, 0.72),
            follow: self.follow.clamp(0.2, 1.0),
            opposite: self.opposite,
            triple: self.triple && self.dual,
        }
    }
}

#[derive(Default, Debug, Clone, Copy, Serialize, Deserialize, EnumIter, EnumString, PartialEq, Eq)]
pub enum JugglePalette {
    #[default]
    Rainbow,
    Custom,
}

#[derive(Debug, Clone, Copy, Serialize, Deserialize, PartialEq)]
#[serde(default)]
pub struct JuggleParams {
    pub dots: f32,
    pub speed: f32,
    pub trail: f32,
    pub background: f32,
    pub palette: JugglePalette,
}

impl Default for JuggleParams {
    fn default() -> Self {
        Self {
            dots: 5.0,
            speed: 0.7,
            trail: 0.62,
            background: 0.03,
            palette: JugglePalette::Rainbow,
        }
    }
}

impl JuggleParams {
    pub fn normalized(self) -> Self {
        Self {
            dots: self.dots.clamp(2.0, 8.0),
            speed: self.speed.clamp(0.15, 2.4),
            trail: self.trail.clamp(0.1, 1.0),
            background: self.background.clamp(0.0, 0.4),
            palette: self.palette,
        }
    }
}

#[derive(Default, Debug, Clone, Copy, Serialize, Deserialize, EnumIter, EnumString, PartialEq, Eq)]
pub enum BouncePalette {
    #[default]
    Rainbow,
    Custom,
}

#[derive(Debug, Clone, Copy, Serialize, Deserialize, PartialEq)]
#[serde(default)]
pub struct BounceParams {
    pub count: f32,
    pub gravity: f32,
    pub size: f32,
    pub trail: f32,
    pub background: f32,
    pub palette: BouncePalette,
}

impl Default for BounceParams {
    fn default() -> Self {
        Self {
            count: 3.0,
            gravity: 0.7,
            size: 0.35,
            trail: 0.55,
            background: 0.03,
            palette: BouncePalette::Rainbow,
        }
    }
}

impl BounceParams {
    pub fn normalized(self) -> Self {
        Self {
            count: self.count.clamp(1.0, 8.0),
            gravity: self.gravity.clamp(0.15, 1.6),
            size: self.size.clamp(0.1, 1.0),
            trail: self.trail.clamp(0.08, 1.0),
            background: self.background.clamp(0.0, 0.4),
            palette: self.palette,
        }
    }
}

#[derive(Default, Debug, Clone, Copy, Serialize, Deserialize, EnumIter, EnumString, PartialEq, Eq)]
pub enum DissolvePalette {
    #[default]
    Custom,
    Rainbow,
}

#[derive(Debug, Clone, Copy, Serialize, Deserialize, PartialEq)]
#[serde(default)]
pub struct DissolveParams {
    pub speed: f32,
    pub background: f32,
    pub random_colors: bool,
    pub palette: DissolvePalette,
}

impl Default for DissolveParams {
    fn default() -> Self {
        Self {
            speed: 0.7,
            background: 0.04,
            random_colors: false,
            palette: DissolvePalette::Custom,
        }
    }
}

impl DissolveParams {
    pub fn normalized(self) -> Self {
        Self {
            speed: self.speed.clamp(0.15, 2.2),
            background: self.background.clamp(0.0, 0.4),
            random_colors: self.random_colors,
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
                | Self::TypeHeat {
                    params: TypeHeatParams {
                        palette: TypeHeatPalette::Custom,
                        ..
                    },
                }
                | Self::Pacifica {
                    params: PacificaParams {
                        palette: PacificaPalette::Custom,
                        ..
                    },
                }
                | Self::DigitalRain {
                    params: DigitalRainParams {
                        palette: DigitalRainPalette::Custom,
                        ..
                    },
                }
                | Self::Fireworks {
                    params: FireworksParams {
                        palette: FireworksPalette::Custom,
                        ..
                    },
                }
                | Self::Nexus {
                    params: NexusParams {
                        palette: NexusPalette::Custom,
                        ..
                    },
                }
                | Self::Comet {
                    params: CometParams {
                        palette: CometPalette::Custom,
                        ..
                    },
                }
                | Self::Juggle {
                    params: JuggleParams {
                        palette: JugglePalette::Custom,
                        ..
                    },
                }
                | Self::Bounce {
                    params: BounceParams {
                        palette: BouncePalette::Custom,
                        ..
                    },
                }
                | Self::Dissolve {
                    params: DissolveParams {
                        palette: DissolvePalette::Custom,
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
            Self::Stars { .. }
                | Self::Rain { .. }
                | Self::Aurora { .. }
                | Self::Scanner { .. }
                | Self::Battery { .. }
                | Self::TypeHeat { .. }
                | Self::Pacifica { .. }
                | Self::DigitalRain { .. }
                | Self::Fireworks { .. }
                | Self::Nexus { .. }
                | Self::Comet { .. }
                | Self::Juggle { .. }
                | Self::Bounce { .. }
                | Self::Dissolve { .. }
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
            analysis: AudioAnalysis::Auto,
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

    pub fn type_heat_default() -> Self {
        Self::TypeHeat {
            params: TypeHeatParams::default(),
        }
    }

    pub fn pacifica_default() -> Self {
        Self::Pacifica {
            params: PacificaParams::default(),
        }
    }

    pub fn digital_rain_default() -> Self {
        Self::DigitalRain {
            params: DigitalRainParams::default(),
        }
    }

    pub fn fireworks_default() -> Self {
        Self::Fireworks {
            params: FireworksParams::default(),
        }
    }

    pub fn nexus_default() -> Self {
        Self::Nexus {
            params: NexusParams::default(),
        }
    }

    pub fn comet_default() -> Self {
        Self::Comet {
            params: CometParams::default(),
        }
    }

    pub fn juggle_default() -> Self {
        Self::Juggle {
            params: JuggleParams::default(),
        }
    }

    pub fn bounce_default() -> Self {
        Self::Bounce {
            params: BounceParams::default(),
        }
    }

    pub fn dissolve_default() -> Self {
        Self::Dissolve {
            params: DissolveParams::default(),
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
            Self::TypeHeat { .. } => Self::type_heat_default(),
            Self::Pacifica { .. } => Self::pacifica_default(),
            Self::DigitalRain { .. } => Self::digital_rain_default(),
            Self::Fireworks { .. } => Self::fireworks_default(),
            Self::Nexus { .. } => Self::nexus_default(),
            Self::Comet { .. } => Self::comet_default(),
            Self::Juggle { .. } => Self::juggle_default(),
            Self::Bounce { .. } => Self::bounce_default(),
            Self::Dissolve { .. } => Self::dissolve_default(),
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
