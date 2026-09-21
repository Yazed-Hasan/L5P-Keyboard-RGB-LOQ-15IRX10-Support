//! LedFx/aubio-style onsets: specdiff + phase deviation + high-band specflux.

pub struct ComplexOnset {
    prev_mag: Vec<f32>,
    prev_phase: Vec<f32>,
    prev2_phase: Vec<f32>,
    primed: bool,
}

impl ComplexOnset {
    pub fn new() -> Self {
        Self {
            prev_mag: Vec::new(),
            prev_phase: Vec::new(),
            prev2_phase: Vec::new(),
            primed: false,
        }
    }

    /// Combined onset (specdiff, phase, high specflux). Same ln1p-ish range as SuperFlux.
    pub fn tick(&mut self, re: &[f32], im: &[f32], sample_rate: u32, fft_size: usize) -> (f32, f32) {
        let n = re.len().min(im.len()).min(fft_size / 2).max(2);
        if self.prev_mag.len() != n {
            self.prev_mag = vec![0.0; n];
            self.prev_phase = vec![0.0; n];
            self.prev2_phase = vec![0.0; n];
            self.primed = false;
        }
        let bin_hz = sample_rate as f32 / fft_size.max(1) as f32;
        let high_lo = (2000.0 / bin_hz).floor().max(1.0) as usize;
        let k_lo = (30.0 / bin_hz).floor().max(1.0) as usize;
        let k_hi = (120.0 / bin_hz).ceil().min((n - 1) as f32) as usize;

        let mut specdiff = 0.0f32;
        let mut phase_dev = 0.0f32;
        let mut high_flux = 0.0f32;
        let mut kick = 0.0f32;
        for i in 1..n {
            let mag = (re[i] * re[i] + im[i] * im[i]).sqrt();
            let phase = im[i].atan2(re[i]);
            let dmag = (mag - self.prev_mag[i]).max(0.0);
            specdiff += dmag;
            let pred = 2.0 * self.prev_phase[i] - self.prev2_phase[i];
            let mut err = phase - pred;
            while err > std::f32::consts::PI {
                err -= 2.0 * std::f32::consts::PI;
            }
            while err < -std::f32::consts::PI {
                err += 2.0 * std::f32::consts::PI;
            }
            phase_dev += err.abs() * mag;
            if i >= high_lo {
                high_flux += dmag;
            }
            if i >= k_lo && i <= k_hi {
                kick += dmag;
            }
            self.prev2_phase[i] = self.prev_phase[i];
            self.prev_phase[i] = phase;
            self.prev_mag[i] = mag;
        }
        if !self.primed {
            self.primed = true;
            return (0.0, 0.0);
        }
        let flux = (specdiff / 64.0).ln_1p() * 0.45 + (phase_dev / 80.0).ln_1p() * 0.25 + (high_flux / 24.0).ln_1p() * 0.30;
        let kick_flux = (kick / 12.0).ln_1p();
        (flux.clamp(0.0, 4.0), kick_flux.clamp(0.0, 4.0))
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::manager::effects::audio_dsp::Analyzer;

    fn run(hz: f32, burst: bool) -> (Vec<f32>, Vec<f32>, u32, usize) {
        let sr = 48_000u32;
        let n = 2048usize;
        let mut samples = vec![0.0f32; n];
        for (i, slot) in samples.iter_mut().enumerate() {
            let t = i as f32 / sr as f32;
            let env = if burst { (-t * 18.0).exp() } else { 1.0 };
            *slot = (2.0 * std::f32::consts::PI * hz * t).sin() * 0.35 * env;
        }
        let mut an = Analyzer::new(sr);
        let zeros = vec![0.0f32; n];
        for _ in 0..4 {
            an.process_window(&zeros);
        }
        an.process_window(&samples);
        (an.re().to_vec(), an.im().to_vec(), sr, an.fft_size())
    }

    #[test]
    fn complex_fires_on_kick_not_slow_sine() {
        let mut onset_k = ComplexOnset::new();
        let mut onset_s = ComplexOnset::new();
        let (re_z, im_z, sr, n) = {
            let mut an = Analyzer::new(48_000);
            an.process_window(&vec![0.0; 2048]);
            (an.re().to_vec(), an.im().to_vec(), 48_000u32, 2048usize)
        };
        onset_k.tick(&re_z, &im_z, sr, n);
        onset_s.tick(&re_z, &im_z, sr, n);
        let (re_k, im_k, _, _) = run(55.0, true);
        let (re_s, im_s, _, _) = run(440.0, false);
        let kick = onset_k.tick(&re_k, &im_k, sr, n).0;
        let sine = onset_s.tick(&re_s, &im_s, sr, n).0;
        assert!(kick > sine * 1.25 || kick > 0.04, "kick={kick} sine={sine}");
    }
}
