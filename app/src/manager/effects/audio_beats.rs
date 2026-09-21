#[derive(Clone, Copy, Debug, Default)]
pub struct BeatHits {
    pub any: bool,
    pub bass: bool,
    pub kick: bool,
    pub triple: bool,
}

impl BeatHits {
    pub fn none() -> Self {
        Self::default()
    }
}

/// SuperFlux peak picker + IBI tempo lock (scheb/NeewerLite style).
pub struct SuperFluxTracker {
    slow: f32,
    prev: f32,
    bass_env: f32,
    kick_env: f32,
    since: f32,
    ibi: f32,
    cool: f32,
    kick_age: [f32; 3],
    kick_n: u8,
    got_beat: bool,
    hist: [f32; 48],
    hist_n: u8,
}

impl SuperFluxTracker {
    pub fn new() -> Self {
        Self {
            slow: 0.02,
            prev: 0.0,
            bass_env: 0.1,
            kick_env: 0.08,
            since: 1.0,
            ibi: 0.5,
            cool: 0.0,
            kick_age: [99.0; 3],
            kick_n: 0,
            got_beat: false,
            hist: [0.0; 48],
            hist_n: 0,
        }
    }

    pub fn ibi(&self) -> f32 {
        self.ibi
    }

    pub fn tempo_locked(&self) -> bool {
        self.got_beat && self.since < 1.2
    }

    pub fn reset(&mut self) {
        *self = Self::new();
    }

    pub fn tick(
        &mut self,
        dt: f32,
        gated: bool,
        flux: f32,
        kick_flux: f32,
        bass: f32,
        kick: f32,
        peak: f32,
        punch: f32,
        squelch: f32,
        beats_bias: bool,
    ) -> BeatHits {
        self.since = (self.since + dt).min(4.0);
        self.cool = (self.cool - dt).max(0.0);
        for age in &mut self.kick_age {
            *age = (*age + dt).min(8.0);
        }
        if gated {
            self.prev = 0.0;
            return BeatHits::none();
        }

        let onset = flux.max(kick_flux * if beats_bias { 1.35 } else { 1.0 });
        if (self.hist_n as usize) < self.hist.len() {
            self.hist[self.hist_n as usize] = onset;
            self.hist_n += 1;
        } else {
            self.hist.rotate_left(1);
            self.hist[self.hist.len() - 1] = onset;
        }
        let median = median_of(&self.hist[..self.hist_n.max(1) as usize]);
        self.slow = ema(self.slow, onset, dt, 0.28).max(0.006);
        let lambda = if beats_bias {
            1.35 - punch * 0.18
        } else {
            1.55 - punch * 0.2
        }
        .clamp(1.15, 1.9);
        let thresh = (median * lambda).max(self.slow * (lambda * 0.85)) + 0.01 + squelch * 0.035;
        self.bass_env = ema(self.bass_env, bass, dt, 0.05).max(0.04);
        self.kick_env = ema(self.kick_env, kick.max(kick_flux), dt, 0.06).max(0.03);
        let bass_share = bass / peak.max(1e-3);
        let bass_jump = (bass - self.bass_env).max(0.0);
        let kick_on = (kick.max(kick_flux) - self.kick_env).max(0.0);
        let kick_share = kick.max(kick_flux) / peak.max(1e-3);
        let crossed = onset > thresh && self.prev <= thresh;
        let strong = onset > thresh && onset > self.prev && (onset - self.prev) > thresh * 0.32;
        let min_gap = if beats_bias {
            (self.ibi * 0.48).clamp(0.16, 0.32)
        } else {
            (self.ibi * 0.56).clamp(0.2, 0.36)
        };
        let ready = self.cool <= 0.0 && self.since >= min_gap;
        let hit = ready && (crossed || strong) && onset > 0.012;
        self.prev = onset;

        let mut hits = BeatHits::none();
        if hit {
            hits.any = true;
            hits.bass = bass_share >= 0.34 && (bass > 0.1 || bass_jump > 0.035 || kick_flux > 0.04);
            hits.kick = kick_on > 0.014
                && (kick > 0.06 || kick_flux > 0.05)
                && kick_share >= 0.08
                && (bass_share >= 0.24 || kick_flux > 0.07);
            if hits.kick {
                hits.bass = true;
            }
            self.accept();
        }
        hits.triple = if hits.kick { self.note_kick() } else { false };
        hits
    }

    fn accept(&mut self) {
        if self.since > 0.18 && self.since < 1.2 {
            self.ibi = (self.ibi * 0.62 + self.since * 0.38).clamp(0.28, 0.85);
        }
        self.cool = (self.ibi * 0.5).clamp(0.16, 0.34);
        self.since = 0.0;
        self.got_beat = true;
    }

    fn note_kick(&mut self) -> bool {
        let mut kept = [99.0f32; 3];
        let mut n = 0usize;
        for i in 0..self.kick_n as usize {
            if self.kick_age[i] < 0.75 && n < 3 {
                kept[n] = self.kick_age[i];
                n += 1;
            }
        }
        if n == 3 {
            kept[0] = kept[1];
            kept[1] = kept[2];
            kept[2] = 0.0;
        } else {
            kept[n] = 0.0;
            n += 1;
        }
        self.kick_age = kept;
        self.kick_n = n as u8;
        if n >= 3 {
            self.kick_n = 0;
            self.kick_age = [99.0; 3];
            true
        } else {
            false
        }
    }
}

fn ema(current: f32, target: f32, dt: f32, tau: f32) -> f32 {
    let alpha = 1.0 - (-dt / tau.max(0.006)).exp();
    current + (target - current) * alpha.clamp(0.0, 1.0)
}

fn median_of(xs: &[f32]) -> f32 {
    if xs.is_empty() {
        return 0.02;
    }
    let mut v = xs.to_vec();
    v.sort_by(|a, b| a.total_cmp(b));
    v[v.len() / 2]
}
