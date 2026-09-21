//! Auto picks among every analysis engine from the mix (Classic only on 4-zone).

use crate::enums::{AudioAnalysis, AudioStyle};

const MARGIN: f32 = 0.12;
const HOLD: f32 = 0.9;
const MIN_SWITCH: f32 = 1.5;

const POOL: [AudioAnalysis; 9] = [
    AudioAnalysis::Classic,
    AudioAnalysis::Accurate,
    AudioAnalysis::Beats,
    AudioAnalysis::Spectrum,
    AudioAnalysis::Mel,
    AudioAnalysis::Studio,
    AudioAnalysis::Hpss,
    AudioAnalysis::Complex,
    AudioAnalysis::Tempo,
];

#[derive(Clone, Copy, Debug)]
pub struct AutoFeatures {
    pub rms: f32,
    pub flux: f32,
    pub kick_share: f32,
    pub flatness: f32,
    pub centroid: f32,
    pub bass_share: f32,
    pub band_spread: f32,
    pub perc_share: f32,
    pub complex_kick: f32,
    pub superflux_kick: f32,
    pub tempo_locked: bool,
    pub tempo_strength: f32,
    pub lamp_n: usize,
    pub style: AudioStyle,
}

impl Default for AutoFeatures {
    fn default() -> Self {
        Self {
            rms: 0.05,
            flux: 0.05,
            kick_share: 0.25,
            flatness: 0.4,
            centroid: 0.5,
            bass_share: 0.2,
            band_spread: 0.0,
            perc_share: 0.0,
            complex_kick: 0.0,
            superflux_kick: 0.0,
            tempo_locked: false,
            tempo_strength: 0.0,
            lamp_n: 24,
            style: AudioStyle::Levels,
        }
    }
}

/// Hysteresis + debounce + silence freeze. Scores every engine except Auto.
pub struct AutoPicker {
    current: AudioAnalysis,
    pending: AudioAnalysis,
    hold: f32,
    since_switch: f32,
    flux_hist: [f32; 64],
    flux_n: usize,
    silent_for: f32,
    was_silent: bool,
}

impl AutoPicker {
    pub fn new() -> Self {
        Self {
            current: AudioAnalysis::Accurate,
            pending: AudioAnalysis::Accurate,
            hold: 0.0,
            since_switch: 3.0,
            flux_hist: [0.0; 64],
            flux_n: 0,
            silent_for: 0.0,
            was_silent: false,
        }
    }

    pub fn current(&self) -> AudioAnalysis {
        self.current
    }

    pub fn reset(&mut self) {
        *self = Self::new();
    }

    pub fn tick(&mut self, dt: f32, feat: &AutoFeatures, gated: bool) -> AudioAnalysis {
        if gated || feat.rms < 4e-5 {
            self.silent_for = (self.silent_for + dt).min(8.0);
            if self.silent_for >= 1.0 {
                self.was_silent = true;
            }
            return self.current;
        }

        if self.was_silent && self.silent_for >= 1.0 {
            self.current = AudioAnalysis::Accurate;
            self.pending = AudioAnalysis::Accurate;
            self.hold = 0.0;
            self.since_switch = 0.0;
            self.flux_n = 0;
        }
        self.was_silent = false;
        self.silent_for = 0.0;
        self.since_switch = (self.since_switch + dt).min(30.0);

        let flux = feat.flux.max(0.0);
        if self.flux_n < self.flux_hist.len() {
            self.flux_hist[self.flux_n] = flux;
            self.flux_n += 1;
        } else {
            self.flux_hist.rotate_left(1);
            self.flux_hist[self.flux_hist.len() - 1] = flux;
        }
        let med = median(&self.flux_hist[..self.flux_n.max(1)]);
        let target = pick(self.current, feat, med);
        if target != self.current {
            if target == self.pending {
                self.hold += dt;
            } else {
                self.pending = target;
                self.hold = 0.0;
            }
            if self.hold > HOLD && self.since_switch > MIN_SWITCH {
                self.current = target;
                self.since_switch = 0.0;
                self.hold = 0.0;
            }
        } else {
            self.hold = 0.0;
            self.pending = self.current;
        }
        self.current
    }
}

fn pick(current: AudioAnalysis, feat: &AutoFeatures, med: f32) -> AudioAnalysis {
    let mut best = current;
    let mut best_s = score(current, feat, med);
    for mode in POOL {
        let s = score(mode, feat, med);
        let need = if mode == current { best_s } else { best_s + MARGIN };
        if s > need + 1e-6 {
            best = mode;
            best_s = s;
        }
    }
    best
}

fn score(mode: AudioAnalysis, f: &AutoFeatures, med: f32) -> f32 {
    let med = med.max(0.008);
    let intense = f.flux > med * 1.3;
    let calm = f.flux < med * 0.7;
    let vocal = f.centroid > 0.12 && f.centroid < 0.55;
    let music = match mode {
        AudioAnalysis::Auto => 0.0,
        AudioAnalysis::Classic => {
            if f.lamp_n != 4 {
                0.0
            } else if f.kick_share < 0.28 && !intense && f.flatness < 0.42 {
                // Modest prior so Mel/Studio/HPSS/Beats can still win on 4-zone.
                0.80 + (0.28 - f.kick_share) * 0.15
            } else {
                0.06
            }
        }
        AudioAnalysis::Accurate => {
            let mixed = f.kick_share > 0.22 && f.kick_share < 0.42;
            0.40 + if mixed { 0.12 } else { 0.0 }
        }
        AudioAnalysis::Beats => {
            let mut s = f.kick_share * 1.35;
            if intense {
                s += 0.22;
            }
            if f.flatness > 0.4 {
                s += 0.12;
            }
            s
        }
        AudioAnalysis::Spectrum => {
            let mut s = (1.0 - f.kick_share) * 0.55 + (1.0 - f.flatness) * 0.45;
            if calm {
                s += 0.22;
            }
            s
        }
        AudioAnalysis::Mel => {
            let mut s = (1.0 - f.kick_share) * 0.28 + (1.0 - f.flatness) * 0.22 + f.band_spread * 0.7;
            if vocal {
                s += 0.22;
            }
            s
        }
        AudioAnalysis::Studio => {
            let mut s = f.bass_share * 0.85 + (0.22 - f.centroid).max(0.0) * 2.8;
            if f.bass_share > 0.55 && f.centroid < 0.16 {
                s += 0.28;
            }
            s
        }
        AudioAnalysis::Hpss => {
            let mixed = f.perc_share > 0.28 && f.perc_share < 0.78;
            if mixed {
                0.58 + f.kick_share * 0.35 + (0.5 - (f.perc_share - 0.5).abs()) * 0.4
            } else {
                f.perc_share * 0.15
            }
        }
        AudioAnalysis::Complex => {
            let ratio = f.complex_kick / (f.superflux_kick + 1e-4);
            let messy = f.flux > med * 1.05 && f.kick_share > 0.18;
            let mut s = f.complex_kick.min(0.3);
            if ratio > 1.08 {
                s += 0.5;
            }
            if ratio > 1.2 {
                s += 0.25;
            }
            if messy {
                s += 0.25;
            }
            s
        }
        AudioAnalysis::Tempo => {
            if f.tempo_locked {
                0.50 + (f.tempo_strength / 8.0).clamp(0.0, 0.45)
            } else {
                0.04
            }
        }
    };
    music + style_boost(f.style, mode)
}

fn style_boost(style: AudioStyle, mode: AudioAnalysis) -> f32 {
    let hit = match style {
        AudioStyle::Levels | AudioStyle::Vu => matches!(mode, AudioAnalysis::Spectrum | AudioAnalysis::Studio),
        AudioStyle::Pulse | AudioStyle::Fire | AudioStyle::Sparkle | AudioStyle::Bubbles => {
            matches!(mode, AudioAnalysis::Beats)
        }
        AudioStyle::Wave | AudioStyle::Bloom | AudioStyle::Gradient | AudioStyle::Mirror => {
            matches!(mode, AudioAnalysis::Mel | AudioAnalysis::Spectrum)
        }
        AudioStyle::Center
        | AudioStyle::Chase
        | AudioStyle::Snake
        | AudioStyle::Oscilloscope
        | AudioStyle::Stereo
        | AudioStyle::Lissajous
        | AudioStyle::MidSide
        | AudioStyle::PanNeedle => matches!(mode, AudioAnalysis::Accurate),
        AudioStyle::Strobe | AudioStyle::Collision | AudioStyle::Ripple => {
            matches!(mode, AudioAnalysis::Complex | AudioAnalysis::Hpss | AudioAnalysis::Beats)
        }
        AudioStyle::BeatGates | AudioStyle::TempoPulse => {
            matches!(mode, AudioAnalysis::Tempo | AudioAnalysis::Beats)
        }
        AudioStyle::Spectrogram
        | AudioStyle::Eq24
        | AudioStyle::Gravcenter
        | AudioStyle::Melt
        | AudioStyle::Wavelength
        | AudioStyle::Pitch
        | AudioStyle::KeyColor => matches!(mode, AudioAnalysis::Mel),
    };
    if hit {
        0.15
    } else {
        0.0
    }
}

fn median(xs: &[f32]) -> f32 {
    if xs.is_empty() {
        return 0.02;
    }
    let mut v = xs.to_vec();
    v.sort_by(|a, b| a.total_cmp(b));
    v[v.len() / 2]
}

#[cfg(test)]
mod tests {
    use super::*;

    fn run(feat: AutoFeatures) -> AudioAnalysis {
        let mut auto = AutoPicker::new();
        let mut last = AudioAnalysis::Accurate;
        for _ in 0..80 {
            last = auto.tick(0.05, &feat, false);
        }
        last
    }

    fn kick_feat() -> AutoFeatures {
        AutoFeatures {
            flux: 0.1,
            kick_share: 0.583,
            flatness: 0.55,
            superflux_kick: 0.14,
            ..AutoFeatures::default()
        }
    }

    fn four(mut feat: AutoFeatures) -> AutoFeatures {
        feat.lamp_n = 4;
        feat
    }

    #[test]
    fn kick_train_becomes_beats() {
        assert_eq!(run(kick_feat()), AudioAnalysis::Beats);
        assert_eq!(run(four(kick_feat())), AudioAnalysis::Beats);
    }

    #[test]
    fn slow_sine_becomes_spectrum() {
        let feat = AutoFeatures {
            flux: 0.02,
            kick_share: 0.167,
            flatness: 0.15,
            superflux_kick: 0.004,
            ..AutoFeatures::default()
        };
        assert_eq!(run(feat), AudioAnalysis::Spectrum);
        assert_eq!(run(four(feat)), AudioAnalysis::Spectrum);
    }

    #[test]
    fn silence_does_not_flip() {
        let mut auto = AutoPicker::new();
        for _ in 0..40 {
            auto.tick(0.05, &kick_feat(), false);
        }
        let locked = auto.current();
        let silent = AutoFeatures {
            rms: 0.0,
            flux: 0.0,
            kick_share: 0.0,
            flatness: 1.0,
            ..AutoFeatures::default()
        };
        for _ in 0..40 {
            auto.tick(0.05, &silent, true);
        }
        assert_eq!(auto.current(), locked);
    }

    #[test]
    fn mixed_kick_and_tone_becomes_hpss() {
        let feat = AutoFeatures {
            flux: 0.08,
            kick_share: 0.35,
            flatness: 0.4,
            perc_share: 0.5,
            superflux_kick: 0.08,
            ..AutoFeatures::default()
        };
        assert_eq!(run(feat), AudioAnalysis::Hpss);
        assert_eq!(run(four(feat)), AudioAnalysis::Hpss);
    }

    #[test]
    fn locked_davies_becomes_tempo() {
        let feat = AutoFeatures {
            flux: 0.08,
            kick_share: 0.28,
            flatness: 0.4,
            tempo_locked: true,
            tempo_strength: 4.0,
            style: AudioStyle::TempoPulse,
            ..AutoFeatures::default()
        };
        assert_eq!(run(feat), AudioAnalysis::Tempo);
        assert_eq!(run(four(feat)), AudioAnalysis::Tempo);
    }

    #[test]
    fn low_bass_pileup_becomes_studio() {
        let feat = AutoFeatures {
            flux: 0.05,
            kick_share: 0.2,
            flatness: 0.4,
            centroid: 0.08,
            bass_share: 0.72,
            style: AudioStyle::Vu,
            ..AutoFeatures::default()
        };
        assert_eq!(run(feat), AudioAnalysis::Studio);
        assert_eq!(run(four(feat)), AudioAnalysis::Studio);
    }

    #[test]
    fn spread_tonal_becomes_mel() {
        let feat = AutoFeatures {
            flux: 0.03,
            kick_share: 0.12,
            flatness: 0.18,
            centroid: 0.32,
            band_spread: 0.72,
            style: AudioStyle::Eq24,
            ..AutoFeatures::default()
        };
        assert_eq!(run(feat), AudioAnalysis::Mel);
        assert_eq!(run(four(feat)), AudioAnalysis::Mel);
    }

    #[test]
    fn complex_cleaner_than_flux_becomes_complex() {
        let feat = AutoFeatures {
            flux: 0.12,
            kick_share: 0.28,
            flatness: 0.38,
            complex_kick: 0.16,
            superflux_kick: 0.06,
            style: AudioStyle::Ripple,
            ..AutoFeatures::default()
        };
        assert_eq!(run(feat), AudioAnalysis::Complex);
        assert_eq!(run(four(feat)), AudioAnalysis::Complex);
    }

    #[test]
    fn twenty_four_lamp_never_classic() {
        let feat = AutoFeatures {
            flux: 0.02,
            kick_share: 0.15,
            flatness: 0.2,
            lamp_n: 24,
            ..AutoFeatures::default()
        };
        assert_ne!(run(feat), AudioAnalysis::Classic);
    }

    #[test]
    fn four_zone_simple_can_classic() {
        let feat = AutoFeatures {
            flux: 0.06,
            kick_share: 0.24,
            flatness: 0.38,
            lamp_n: 4,
            style: AudioStyle::Center,
            ..AutoFeatures::default()
        };
        assert_eq!(run(feat), AudioAnalysis::Classic);
    }
}
