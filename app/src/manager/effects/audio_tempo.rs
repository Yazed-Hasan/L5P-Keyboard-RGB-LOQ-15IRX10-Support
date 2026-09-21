//! Causal Davies-style tempo: autocorrelation of the onset envelope.

use std::collections::VecDeque;

pub struct DaviesTempo {
    odf: VecDeque<f32>,
    hop_hz: f32,
    ibi: f32,
    locked: bool,
    strength: f32,
}

impl DaviesTempo {
    pub fn new(sample_rate: u32, hop: usize) -> Self {
        let hop_hz = sample_rate.max(1) as f32 / hop.max(1) as f32;
        Self {
            odf: VecDeque::new(),
            hop_hz,
            ibi: 0.5,
            locked: false,
            strength: 0.0,
        }
    }

    pub fn ibi(&self) -> f32 {
        self.ibi
    }

    pub fn locked(&self) -> bool {
        self.locked
    }

    pub fn strength(&self) -> f32 {
        self.strength
    }

    pub fn reset(&mut self) {
        self.odf.clear();
        self.locked = false;
        self.strength = 0.0;
    }

    pub fn tick(&mut self, flux: f32) {
        self.odf.push_back(flux.max(0.0));
        while self.odf.len() > 512 {
            self.odf.pop_front();
        }
        if self.odf.len() < 80 {
            self.locked = false;
            return;
        }
        let lag_min = (0.28 * self.hop_hz).round().clamp(8.0, 200.0) as usize;
        let lag_max = (0.85 * self.hop_hz).round().clamp(lag_min as f32 + 4.0, 400.0) as usize;
        let lag_max = lag_max.min(self.odf.len() / 2);
        if lag_max <= lag_min {
            self.locked = false;
            return;
        }
        let mut best_lag = lag_min;
        let mut best = 0.0f32;
        let mut scores = Vec::with_capacity(lag_max.saturating_sub(lag_min) + 1);
        for lag in lag_min..=lag_max {
            let mut acc = 0.0f32;
            let n = self.odf.len() - lag;
            for i in 0..n {
                acc += self.odf[i] * self.odf[i + lag];
            }
            let mut score = acc / n as f32;
            let lag2 = lag * 2;
            if lag2 <= lag_max {
                let mut acc2 = 0.0f32;
                let n2 = self.odf.len() - lag2;
                for i in 0..n2 {
                    acc2 += self.odf[i] * self.odf[i + lag2];
                }
                score += 0.5 * acc2 / n2 as f32;
            }
            scores.push(score);
            if score > best {
                best = score;
                best_lag = lag;
            }
        }
        let mut ranked = scores;
        ranked.sort_by(|a, b| a.total_cmp(b));
        let med = ranked[ranked.len() / 2].max(1e-12);
        self.strength = best / med;
        self.locked = self.strength > 2.4 && best > 1e-8;
        if self.locked {
            let period = best_lag as f32 / self.hop_hz;
            self.ibi = (self.ibi * 0.72 + period * 0.28).clamp(0.28, 0.85);
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn two_hz_pulse_locks_near_120_bpm() {
        let mut tempo = DaviesTempo::new(48_000, 256);
        let hop_hz = 48_000.0f32 / 256.0;
        let period = (0.5 * hop_hz).round() as usize;
        for i in 0..400 {
            let flux = if i % period == 0 { 1.0 } else { 0.02 };
            tempo.tick(flux);
        }
        let bpm = 60.0 / tempo.ibi();
        assert!(tempo.locked(), "should lock onto a 2 Hz pulse");
        assert!((bpm - 120.0).abs() < 18.0, "bpm={bpm} ibi={}", tempo.ibi());
    }
}
