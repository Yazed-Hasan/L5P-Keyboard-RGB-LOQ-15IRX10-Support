//! LedFx-style 24-band mel filterbank (matt_mel / HTK mel).

use super::audio_dsp::{triangle_mag, EQ_BANDS};

pub struct MelBank {
    agc: [f32; EQ_BANDS],
}

impl MelBank {
    pub fn new() -> Self {
        Self { agc: [1e-4; EQ_BANDS] }
    }

    pub fn tick(&mut self, mag: &[f32], sample_rate: u32, fft_size: usize) -> ([f32; EQ_BANDS], [f32; 4]) {
        let mut eq24 = [0.0f32; EQ_BANDS];
        let lo_hz = 30.0f32;
        let hi_hz = 12_000.0f32;
        let lo_mel = hz_to_mel(lo_hz);
        let hi_mel = hz_to_mel(hi_hz);
        for i in 0..EQ_BANDS {
            let t0 = i as f32 / EQ_BANDS as f32;
            let t1 = (i + 1) as f32 / EQ_BANDS as f32;
            let mid_t = (t0 + t1) * 0.5;
            let mid = mel_to_hz(lo_mel + (hi_mel - lo_mel) * mid_t);
            let lo = mel_to_hz(lo_mel + (hi_mel - lo_mel) * t0).max(20.0);
            let hi = mel_to_hz(lo_mel + (hi_mel - lo_mel) * t1).min(16_000.0);
            let tilt = (mid / 800.0).powf(0.18).clamp(0.6, 1.7);
            let raw = triangle_mag(mag, sample_rate, fft_size, lo, mid, hi) * tilt;
            self.agc[i] = (self.agc[i] * 0.997).max(raw).max(1e-4);
            eq24[i] = (raw / self.agc[i]).clamp(0.0, 1.0);
        }
        (eq24, bands4_from_eq24(&eq24))
    }
}

pub fn bands4_from_eq24(eq: &[f32; EQ_BANDS]) -> [f32; 4] {
    let mut out = [0.0f32; 4];
    let chunk = EQ_BANDS / 4;
    for b in 0..4 {
        let start = b * chunk;
        let end = if b == 3 { EQ_BANDS } else { start + chunk };
        let mut acc = 0.0;
        let mut n = 0.0;
        for slot in eq.iter().take(end).skip(start) {
            acc += *slot;
            n += 1.0;
        }
        out[b] = if n > 0.0 { (acc / n).clamp(0.0, 1.0) } else { 0.0 };
    }
    out
}

fn hz_to_mel(hz: f32) -> f32 {
    2595.0 * (1.0 + hz.max(0.0) / 700.0).log10()
}

fn mel_to_hz(mel: f32) -> f32 {
    700.0 * (10.0f32.powf(mel / 2595.0) - 1.0)
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::manager::effects::audio_dsp::Analyzer;

    fn tone_mag(hz: f32) -> (Vec<f32>, u32, usize) {
        let sr = 48_000u32;
        let n = 2048usize;
        let mut samples = vec![0.0f32; n];
        for (i, slot) in samples.iter_mut().enumerate() {
            *slot = (2.0 * std::f32::consts::PI * hz * i as f32 / sr as f32).sin() * 0.4;
        }
        let mut an = Analyzer::new(sr);
        for _ in 0..4 {
            an.process_window(&samples);
        }
        (an.mag().to_vec(), sr, an.fft_size())
    }

    #[test]
    fn mel_100hz_peaks_left_of_8k() {
        let mut mel = MelBank::new();
        let (low_mag, sr, n) = tone_mag(100.0);
        let (high_mag, _, _) = tone_mag(8000.0);
        let (low, _) = mel.tick(&low_mag, sr, n);
        let mut mel2 = MelBank::new();
        let (high, _) = mel2.tick(&high_mag, sr, n);
        let low_peak = low
            .iter()
            .enumerate()
            .max_by(|a, b| a.1.total_cmp(b.1))
            .map(|(i, _)| i)
            .unwrap();
        let high_peak = high
            .iter()
            .enumerate()
            .max_by(|a, b| a.1.total_cmp(b.1))
            .map(|(i, _)| i)
            .unwrap();
        assert!(low_peak < high_peak, "100Hz bin {low_peak} should be left of 8kHz bin {high_peak}");
    }
}
