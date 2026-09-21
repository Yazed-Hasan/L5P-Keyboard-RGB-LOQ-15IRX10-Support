use crate::enums::AudioColorMode;

use super::lamps;

pub struct ColorCtx {
    pub custom_rgb: [u8; 12],
    pub t: f32,
    pub energy: f32,
    pub phase: f32,
    pub ramp_phase: f32,
    pub motion: f32,
}

enum TintSource {
    Palette(&'static [u8; 12]),
    Custom,
    Mono,
    Rainbow,
    Energy,
    Ramp,
}

const SPECTRUM: [u8; 12] = [255, 24, 48, 255, 140, 16, 36, 220, 120, 72, 120, 255];
const HEAT: [u8; 12] = [160, 6, 0, 255, 88, 6, 255, 180, 20, 255, 230, 70];
const ICE: [u8; 12] = [0, 180, 220, 20, 120, 255, 80, 160, 255, 220, 240, 255];
const SUNSET: [u8; 12] = [80, 20, 120, 180, 40, 140, 255, 80, 100, 255, 160, 40];

fn source(mode: AudioColorMode) -> TintSource {
    match mode {
        AudioColorMode::Custom => TintSource::Custom,
        AudioColorMode::Mono => TintSource::Mono,
        AudioColorMode::Rainbow => TintSource::Rainbow,
        AudioColorMode::Energy => TintSource::Energy,
        AudioColorMode::Ramp => TintSource::Ramp,
        AudioColorMode::Spectrum => TintSource::Palette(&SPECTRUM),
        AudioColorMode::Heat => TintSource::Palette(&HEAT),
        AudioColorMode::Ice => TintSource::Palette(&ICE),
        AudioColorMode::Sunset => TintSource::Palette(&SUNSET),
    }
}

pub fn tint(mode: AudioColorMode, ctx: ColorCtx) -> [u8; 3] {
    match source(mode) {
        TintSource::Palette(stops) => lamps::sample_zones(stops, ctx.t),
        TintSource::Custom => lamps::sample_zones(&ctx.custom_rgb, ctx.t),
        TintSource::Mono => lamps::zone_rgb(&ctx.custom_rgb, 0),
        TintSource::Rainbow => {
            let hue = (ctx.t * 360.0 + ctx.phase * (18.0 + ctx.motion * 24.0)) % 360.0;
            let (r, g, b) = hsv_to_rgb(hue, 0.92, 1.0);
            [r, g, b]
        }
        TintSource::Energy => {
            let hue = (220.0 - ctx.energy.clamp(0.0, 1.0) * 220.0).clamp(0.0, 220.0);
            let (r, g, b) = hsv_to_rgb(hue, 0.92, 1.0);
            [r, g, b]
        }
        TintSource::Ramp => {
            let base = ctx.ramp_phase * (180.0 / std::f32::consts::PI);
            let hue = (base + ctx.t * 80.0).rem_euclid(360.0);
            let (r, g, b) = hsv_to_rgb(hue, 0.9, 1.0);
            [r, g, b]
        }
    }
}

pub fn name(mode: AudioColorMode) -> &'static str {
    match mode {
        AudioColorMode::Custom => "Custom",
        AudioColorMode::Spectrum => "Spectrum",
        AudioColorMode::Rainbow => "Rainbow",
        AudioColorMode::Heat => "Heat",
        AudioColorMode::Ice => "Ice",
        AudioColorMode::Sunset => "Sunset",
        AudioColorMode::Energy => "Energy",
        AudioColorMode::Mono => "Mono",
        AudioColorMode::Ramp => "Ramp",
    }
}

pub fn tip(mode: AudioColorMode) -> &'static str {
    match mode {
        AudioColorMode::Custom => "Use the four color boxes above. Brightness still follows the music.",
        AudioColorMode::Spectrum => "Bass is red/orange, mids are green, highs go blue/purple.",
        AudioColorMode::Rainbow => "Colors cycle across the keyboard as the music plays.",
        AudioColorMode::Heat => "Red to orange to yellow across the keyboard, like a fire bar.",
        AudioColorMode::Ice => "Cyan to blue to white across the keyboard.",
        AudioColorMode::Sunset => "Purple to pink to orange across the keyboard.",
        AudioColorMode::Energy => "Color follows loudness: cool when quiet, hot on peaks.",
        AudioColorMode::Mono => "Whole keyboard uses Color 1. Brightness still follows the music.",
        AudioColorMode::Ramp => "A wide color wash that slowly ramps across the keyboard. Motion sets the speed.",
    }
}

pub fn uses_swatches(mode: AudioColorMode) -> bool {
    matches!(mode, AudioColorMode::Custom | AudioColorMode::Mono)
}

pub fn hue_shift_ok(mode: AudioColorMode) -> bool {
    !matches!(
        mode,
        AudioColorMode::Rainbow | AudioColorMode::Energy | AudioColorMode::Ramp
    )
}

pub fn swatch_hint(mode: AudioColorMode) -> Option<(&'static str, &'static str)> {
    if !uses_swatches(mode) {
        return None;
    }
    match mode {
        AudioColorMode::Custom => Some((
            "Use the 4 color boxes above for each zone.",
            "Left to right matches the four keyboard zones. The wide box under them tints all four at once.",
        )),
        AudioColorMode::Mono => Some((
            "Uses Color 1 (left box) for the whole keyboard.",
            "The other three boxes are ignored in Mono. Brightness still follows the music.",
        )),
        _ => None,
    }
}

pub fn ease_rgb(prev: &mut [f32; 3], target: [u8; 3], dt: f32, ramp: f32) -> [u8; 3] {
    let ramp = ramp.clamp(0.0, 1.0);
    if ramp < 0.01 {
        *prev = [target[0] as f32, target[1] as f32, target[2] as f32];
        return target;
    }
    let tau = 0.035 + ramp * 0.55;
    let alpha = (1.0 - (-dt / tau.max(0.008)).exp()).clamp(0.0, 1.0);
    for i in 0..3 {
        prev[i] += (target[i] as f32 - prev[i]) * alpha;
    }
    [
        prev[0].round().clamp(0.0, 255.0) as u8,
        prev[1].round().clamp(0.0, 255.0) as u8,
        prev[2].round().clamp(0.0, 255.0) as u8,
    ]
}

pub fn rgb_to_hsv(r: u8, g: u8, b: u8) -> (f32, f32, f32) {
    let r = r as f32 / 255.0;
    let g = g as f32 / 255.0;
    let b = b as f32 / 255.0;
    let max = r.max(g).max(b);
    let min = r.min(g).min(b);
    let delta = max - min;
    let hue = if delta < 1e-6 {
        0.0
    } else if (max - r).abs() < 1e-6 {
        60.0 * (((g - b) / delta) % 6.0)
    } else if (max - g).abs() < 1e-6 {
        60.0 * (((b - r) / delta) + 2.0)
    } else {
        60.0 * (((r - g) / delta) + 4.0)
    };
    let sat = if max < 1e-6 { 0.0 } else { delta / max };
    ((hue + 360.0) % 360.0, sat, max)
}

pub fn hsv_to_rgb(h: f32, s: f32, v: f32) -> (u8, u8, u8) {
    let c = v * s;
    let x = c * (1.0 - ((h / 60.0) % 2.0 - 1.0).abs());
    let m = v - c;
    let (r, g, b) = match h {
        h if h < 60.0 => (c, x, 0.0),
        h if h < 120.0 => (x, c, 0.0),
        h if h < 180.0 => (0.0, c, x),
        h if h < 240.0 => (0.0, x, c),
        h if h < 300.0 => (x, 0.0, c),
        _ => (c, 0.0, x),
    };
    (((r + m) * 255.0) as u8, ((g + m) * 255.0) as u8, ((b + m) * 255.0) as u8)
}
