//! Causal median HPSS (sevagh/Real-Time-HPSS). Flux comes from the percussive mask.

use std::collections::VecDeque;

pub struct Hpss {
    hist: VecDeque<Vec<f32>>,
    width: usize,
    freq_span: usize,
    prev_log: Vec<f32>,
    primed: bool,
    perc_share: f32,
}

impl Hpss {
    pub fn new() -> Self {
        Self {
            hist: VecDeque::new(),
            width: 17,
            freq_span: 17,
            prev_log: Vec::new(),
            primed: false,
            perc_share: 0.0,
        }
    }

    pub fn perc_share(&self) -> f32 {
        self.perc_share
    }

    /// Returns (p_flux, kick_flux) in the same ln1p units as SuperFlux.
    pub fn tick(&mut self, mag: &[f32], sample_rate: u32, fft_size: usize) -> (f32, f32) {
        if mag.is_empty() {
            return (0.0, 0.0);
        }
        self.hist.push_back(mag.to_vec());
        while self.hist.len() > self.width {
            self.hist.pop_front();
        }
        let bins = mag.len();
        let mut harm = vec![0.0f32; bins];
        let mut perc = vec![0.0f32; bins];
        let frames = self.hist.len();
        let mut col = vec![0.0f32; frames];
        for b in 0..bins {
            for (t, frame) in self.hist.iter().enumerate() {
                col[t] = frame.get(b).copied().unwrap_or(0.0);
            }
            harm[b] = median(&col);
            let lo = b.saturating_sub(self.freq_span / 2);
            let hi = (b + self.freq_span / 2 + 1).min(bins);
            perc[b] = median(&mag[lo..hi]);
        }
        let mut p_mag = vec![0.0f32; bins];
        let mut p_e = 0.0f32;
        let mut h_e = 0.0f32;
        for b in 1..bins {
            let h2 = harm[b] * harm[b];
            let p2 = perc[b] * perc[b];
            let den = h2 + p2 + 1e-9;
            p_mag[b] = mag[b] * p2 / den;
            p_e += p_mag[b];
            h_e += mag[b] * h2 / den;
        }
        self.perc_share = p_e / (p_e + h_e + 1e-9);
        if self.prev_log.len() != bins {
            self.prev_log = vec![0.0; bins];
            self.primed = false;
        }
        let mut logm = vec![0.0f32; bins];
        for b in 1..bins {
            logm[b] = p_mag[b].ln_1p();
        }
        if !self.primed {
            self.prev_log = logm;
            self.primed = true;
            return (0.0, 0.0);
        }
        let bin_hz = sample_rate as f32 / fft_size.max(1) as f32;
        let k_lo = (30.0 / bin_hz).floor().max(1.0) as usize;
        let k_hi = (120.0 / bin_hz).ceil().min((bins.saturating_sub(1)) as f32) as usize;
        let mut flux = 0.0f32;
        let mut kick_flux = 0.0f32;
        for i in 1..bins.saturating_sub(1).max(2) {
            let prev_max = self.prev_log[i.saturating_sub(1)]
                .max(self.prev_log[i])
                .max(self.prev_log[(i + 1).min(bins - 1)]);
            let d = (logm[i] - prev_max).max(0.0);
            flux += d;
            if i >= k_lo && i <= k_hi {
                kick_flux += d;
            }
        }
        self.prev_log = logm;
        let flux_div = (bins as f32 / 21.3).max(8.0);
        let kick_div = ((k_hi.saturating_sub(k_lo) + 1) as f32 / 7.5).max(4.0);
        (
            (flux / flux_div).ln_1p(),
            (kick_flux / kick_div).ln_1p(),
        )
    }
}

fn median(xs: &[f32]) -> f32 {
    if xs.is_empty() {
        return 0.0;
    }
    let mut v = xs.to_vec();
    v.sort_by(|a, b| a.total_cmp(b));
    v[v.len() / 2]
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::manager::effects::audio_dsp::Analyzer;

    fn mag_of(hz: f32, burst: bool) -> (Vec<f32>, u32, usize) {
        let sr = 48_000u32;
        let n = 2048usize;
        let mut samples = vec![0.0f32; n];
        for (i, slot) in samples.iter_mut().enumerate() {
            let t = i as f32 / sr as f32;
            let env = if burst { (-t * 16.0).exp() } else { 1.0 };
            *slot = (2.0 * std::f32::consts::PI * hz * t).sin() * 0.35 * env;
        }
        let mut an = Analyzer::new(sr);
        an.process_window(&samples);
        (an.mag().to_vec(), sr, an.fft_size())
    }

    fn silence_mag() -> (Vec<f32>, u32, usize) {
        let mut an = Analyzer::new(48_000);
        an.process_window(&vec![0.0; 2048]);
        (an.mag().to_vec(), 48_000u32, 2048usize)
    }

    #[test]
    fn kick_raises_percussive_flux_more_than_sine() {
        let (z, sr, n) = silence_mag();
        let (kick_mag, _, _) = mag_of(55.0, true);
        let (sine_mag, _, _) = mag_of(440.0, false);
        let mut hpss_k = Hpss::new();
        let mut hpss_s = Hpss::new();
        hpss_k.tick(&z, sr, n);
        hpss_s.tick(&z, sr, n);
        let kick_p = hpss_k.tick(&kick_mag, sr, n).0;
        let mut sine_p = 0.0;
        for _ in 0..6 {
            sine_p = hpss_s.tick(&sine_mag, sr, n).0;
        }
        assert!(kick_p > sine_p * 1.15, "kick P-flux {kick_p} vs sine {sine_p}");
    }
}
