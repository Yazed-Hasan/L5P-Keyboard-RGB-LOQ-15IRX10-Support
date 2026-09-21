use std::f32::consts::PI;

pub const DSP_SIZE: usize = 2048;
pub const HOP: usize = 256;
pub const STUDIO_SIZE: usize = 4096;
pub const STUDIO_HOP: usize = 512;
pub const EQ_BANDS: usize = 24;

#[derive(Clone, Copy, Debug)]
pub struct AnalysisFrame {
    pub bands4: [f32; 4],
    pub eq24: [f32; EQ_BANDS],
    pub rms: f32,
    pub flux: f32,
    pub kick_flux: f32,
    #[allow(dead_code)]
    pub centroid: f32,
    pub flatness: f32,
    pub kick: f32,
}

impl Default for AnalysisFrame {
    fn default() -> Self {
        Self {
            bands4: [0.0; 4],
            eq24: [0.0; EQ_BANDS],
            rms: 0.0,
            flux: 0.0,
            kick_flux: 0.0,
            centroid: 0.5,
            flatness: 1.0,
            kick: 0.0,
        }
    }
}

pub struct Analyzer {
    sample_rate: u32,
    fft_size: usize,
    hop: usize,
    max_hops: usize,
    buf: Vec<f32>,
    pending: Vec<f32>,
    window: Vec<f32>,
    prev_log: Vec<f32>,
    re: Vec<f32>,
    im: Vec<f32>,
    mag: Vec<f32>,
    agc4: [f32; 4],
    agc24: [f32; EQ_BANDS],
    last_seq: u64,
    last_hops: usize,
    hop_flux: Vec<f32>,
    last: AnalysisFrame,
}

impl Analyzer {
    pub fn new(sample_rate: u32) -> Self {
        Self::with_size(sample_rate, DSP_SIZE, HOP, 8)
    }

    pub fn studio(sample_rate: u32) -> Self {
        Self::with_size(sample_rate, STUDIO_SIZE, STUDIO_HOP, 4)
    }

    fn with_size(sample_rate: u32, fft_size: usize, hop: usize, max_hops: usize) -> Self {
        let sr = sample_rate.max(8_000);
        let fft_size = fft_size.max(64);
        let hop = hop.clamp(32, fft_size);
        let n = fft_size as f32;
        let mut window = vec![0.0f32; fft_size];
        for (i, slot) in window.iter_mut().enumerate() {
            let t = 2.0 * PI * i as f32 / n;
            *slot = 0.35875 - 0.48829 * t.cos() + 0.14128 * (2.0 * t).cos() - 0.01168 * (3.0 * t).cos();
        }
        Self {
            sample_rate: sr,
            fft_size,
            hop,
            max_hops: max_hops.max(1),
            buf: vec![0.0; fft_size],
            pending: Vec::new(),
            window,
            prev_log: vec![0.0; fft_size / 2],
            re: vec![0.0; fft_size],
            im: vec![0.0; fft_size],
            mag: vec![0.0; fft_size / 2],
            agc4: [1e-4; 4],
            agc24: [1e-4; EQ_BANDS],
            last_seq: 0,
            last_hops: 0,
            hop_flux: Vec::new(),
            last: AnalysisFrame::default(),
        }
    }

    pub fn sample_rate(&self) -> u32 {
        self.sample_rate
    }

    pub fn fft_size(&self) -> usize {
        self.fft_size
    }

    pub fn hop(&self) -> usize {
        self.hop
    }

    pub fn last_hops(&self) -> usize {
        self.last_hops
    }

    pub fn hop_fluxes(&self) -> &[f32] {
        &self.hop_flux
    }

    pub fn mag(&self) -> &[f32] {
        &self.mag
    }

    pub fn re(&self) -> &[f32] {
        &self.re
    }

    pub fn im(&self) -> &[f32] {
        &self.im
    }

    /// Consume newly written ring samples using a monotonic write counter.
    pub fn observe_seq(&mut self, ring: &[f32], seq: u64) -> AnalysisFrame {
        self.hop_flux.clear();
        if ring.is_empty() {
            self.last_hops = 0;
            return self.last;
        }
        let delta = seq.saturating_sub(self.last_seq) as usize;
        self.last_seq = seq;
        if delta == 0 {
            self.last_hops = 0;
            return self.last;
        }
        let take = delta.min(ring.len());
        self.pending.extend_from_slice(&ring[ring.len() - take..]);
        if self.pending.len() > self.fft_size * 4 {
            let extra = self.pending.len() - self.fft_size * 2;
            self.pending.drain(..extra);
        }
        let mut hops = 0usize;
        while self.pending.len() >= self.hop && hops < self.max_hops {
            let hop: Vec<f32> = self.pending.drain(..self.hop).collect();
            self.push_hop(&hop);
            self.hop_flux.push(self.last.flux);
            hops += 1;
        }
        self.last_hops = hops;
        self.last
    }

    /// One STFT of `samples` (padded/truncated to fft_size). Used by tests.
    pub fn process_window(&mut self, samples: &[f32]) -> AnalysisFrame {
        self.buf.fill(0.0);
        let n = samples.len().min(self.fft_size);
        if n > 0 {
            self.buf[self.fft_size - n..].copy_from_slice(&samples[samples.len() - n..]);
        }
        self.process_current();
        self.last
    }

    fn push_hop(&mut self, hop: &[f32]) {
        let h = hop.len().min(self.fft_size);
        self.buf.rotate_left(h);
        let n = self.buf.len();
        self.buf[n - h..].copy_from_slice(&hop[..h]);
        self.process_current();
    }

    fn process_current(&mut self) {
        let n = self.fft_size;
        self.re.fill(0.0);
        self.im.fill(0.0);
        for i in 0..n {
            self.re[i] = self.buf[i] * self.window[i];
        }
        fft(&mut self.re, &mut self.im);

        let nyquist = n / 2;
        if self.mag.len() != nyquist {
            self.mag.resize(nyquist, 0.0);
            self.prev_log.resize(nyquist, 0.0);
        }
        let mut logm = vec![0.0f32; nyquist];
        let mut acc = 0.0f32;
        for i in 1..nyquist {
            self.mag[i] = (self.re[i] * self.re[i] + self.im[i] * self.im[i]).sqrt();
            logm[i] = self.mag[i].ln_1p();
            acc += self.buf[i] * self.buf[i];
        }
        acc += self.buf[0] * self.buf[0];
        self.mag[0] = self.re[0].abs();
        let rms = (acc / n as f32).sqrt();

        let mut flux = 0.0f32;
        let mut kick_flux = 0.0f32;
        let bin_hz = self.sample_rate as f32 / n as f32;
        let k_lo = (30.0 / bin_hz).floor().max(1.0) as usize;
        let k_hi = (120.0 / bin_hz).ceil().min((nyquist - 2) as f32) as usize;
        for i in 1..nyquist - 1 {
            let prev_max = self.prev_log[i.saturating_sub(1)]
                .max(self.prev_log[i])
                .max(self.prev_log[(i + 1).min(nyquist - 1)]);
            let d = (logm[i] - prev_max).max(0.0);
            flux += d;
            if i >= k_lo && i <= k_hi {
                kick_flux += d;
            }
        }
        self.prev_log.copy_from_slice(&logm);

        let flux_div = (nyquist as f32 / 21.3).max(8.0);
        let kick_div = ((k_hi.saturating_sub(k_lo) + 1) as f32 / 7.5).max(4.0);
        let flux_n = (flux / flux_div).ln_1p();
        let kick_flux_n = (kick_flux / kick_div).ln_1p();

        let raw4 = [
            triangle_mag(&self.mag, self.sample_rate, n, 25.0, 70.0, 200.0),
            triangle_mag(&self.mag, self.sample_rate, n, 90.0, 320.0, 800.0),
            triangle_mag(&self.mag, self.sample_rate, n, 400.0, 1400.0, 3500.0),
            triangle_mag(&self.mag, self.sample_rate, n, 1800.0, 5000.0, 12000.0),
        ];
        let kick = triangle_mag(&self.mag, self.sample_rate, n, 30.0, 55.0, 105.0);
        let mut bands4 = [0.0f32; 4];
        for i in 0..4 {
            self.agc4[i] = (self.agc4[i] * 0.997).max(raw4[i]).max(1e-4);
            bands4[i] = (raw4[i] / self.agc4[i]).clamp(0.0, 1.0);
        }

        let lo_hz = 30.0f32;
        let hi_hz = 12_000.0f32;
        let ratio = hi_hz / lo_hz;
        let mut eq24 = [0.0f32; EQ_BANDS];
        for i in 0..EQ_BANDS {
            let t0 = i as f32 / EQ_BANDS as f32;
            let t1 = (i + 1) as f32 / EQ_BANDS as f32;
            let mid_t = (t0 + t1) * 0.5;
            let mid = lo_hz * ratio.powf(mid_t);
            let lo = (lo_hz * ratio.powf(t0) * 0.72).max(20.0);
            let hi = (lo_hz * ratio.powf(t1) * 1.38).min(16_000.0);
            let tilt = (mid / 1000.0).powf(0.22).clamp(0.55, 1.8);
            let raw = triangle_mag(&self.mag, self.sample_rate, n, lo, mid, hi) * tilt;
            self.agc24[i] = (self.agc24[i] * 0.997).max(raw).max(1e-4);
            eq24[i] = (raw / self.agc24[i]).clamp(0.0, 1.0);
        }
        blur_neighbors(&mut eq24);

        self.last = AnalysisFrame {
            bands4,
            eq24,
            rms,
            flux: flux_n,
            kick_flux: kick_flux_n,
            centroid: spectral_centroid(&self.mag, bin_hz),
            flatness: spectral_flatness(&self.mag),
            kick,
        };
    }
}

pub fn triangle_mag(mag: &[f32], sample_rate: u32, fft_size: usize, lo: f32, mid: f32, hi: f32) -> f32 {
    let bin_hz = sample_rate as f32 / fft_size.max(1) as f32;
    let mut acc = 0.0;
    let mut weight = 0.0;
    let start = (lo / bin_hz).floor().max(1.0) as usize;
    let end = (hi / bin_hz).ceil().min((mag.len().saturating_sub(1)) as f32) as usize;
    if start > end {
        return 0.0;
    }
    for i in start..=end {
        let freq = i as f32 * bin_hz;
        let tri = if freq <= mid {
            (freq - lo) / (mid - lo).max(1e-3)
        } else {
            (hi - freq) / (hi - mid).max(1e-3)
        }
        .clamp(0.0, 1.0);
        if tri <= 0.0 {
            continue;
        }
        acc += mag[i] * tri;
        weight += tri;
    }
    let mean = if weight > 0.0 { acc / weight } else { 0.0 };
    (mean / 48.0).ln_1p()
}

fn spectral_centroid(mag: &[f32], bin_hz: f32) -> f32 {
    let mut wsum = 0.0f32;
    let mut msum = 0.0f32;
    for (i, &m) in mag.iter().enumerate().skip(1) {
        wsum += m * i as f32 * bin_hz;
        msum += m;
    }
    if msum < 1e-9 {
        return 0.5;
    }
    ((wsum / msum) / 8_000.0).clamp(0.0, 1.0)
}

fn spectral_flatness(mag: &[f32]) -> f32 {
    let mut log_sum = 0.0f32;
    let mut sum = 0.0f32;
    let mut n = 0.0f32;
    for &m in mag.iter().skip(1) {
        if m > 0.001 {
            log_sum += m.max(1e-12).ln();
            sum += m;
            n += 1.0;
        }
    }
    if n < 8.0 {
        return 1.0;
    }
    let geo = (log_sum / n).exp();
    let arith = sum / n;
    (geo / arith.max(1e-12)).clamp(0.0, 1.0)
}

fn blur_neighbors(bands: &mut [f32]) {
    if bands.len() < 3 {
        return;
    }
    let orig = bands.to_vec();
    for i in 0..bands.len() {
        let left = orig[i.saturating_sub(1)];
        let right = orig[(i + 1).min(orig.len() - 1)];
        bands[i] = orig[i] * 0.62 + left * 0.19 + right * 0.19;
    }
}

pub(crate) fn fft(re: &mut [f32], im: &mut [f32]) {
    let n = re.len();
    let mut j = 0usize;
    for i in 1..n {
        let mut bit = n >> 1;
        while j & bit != 0 {
            j ^= bit;
            bit >>= 1;
        }
        j ^= bit;
        if i < j {
            re.swap(i, j);
            im.swap(i, j);
        }
    }
    let mut len = 2usize;
    while len <= n {
        let ang = -2.0 * PI / len as f32;
        let (wlen_re, wlen_im) = (ang.cos(), ang.sin());
        let mut i = 0usize;
        while i < n {
            let mut w_re = 1.0f32;
            let mut w_im = 0.0f32;
            for k in 0..len / 2 {
                let u_re = re[i + k];
                let u_im = im[i + k];
                let v_re = re[i + k + len / 2] * w_re - im[i + k + len / 2] * w_im;
                let v_im = re[i + k + len / 2] * w_im + im[i + k + len / 2] * w_re;
                re[i + k] = u_re + v_re;
                im[i + k] = u_im + v_im;
                re[i + k + len / 2] = u_re - v_re;
                im[i + k + len / 2] = u_im - v_im;
                let next_re = w_re * wlen_re - w_im * wlen_im;
                w_im = w_re * wlen_im + w_im * wlen_re;
                w_re = next_re;
            }
            i += len;
        }
        len <<= 1;
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn kick_burst(n: usize, sample_rate: u32) -> Vec<f32> {
        let mut out = vec![0.0f32; n];
        let f = 55.0;
        for i in 0..n {
            let t = i as f32 / sample_rate as f32;
            let env = (-t * 18.0).exp();
            out[i] = (2.0 * PI * f * t).sin() * env;
        }
        out
    }

    fn sine(n: usize, sample_rate: u32, hz: f32) -> Vec<f32> {
        (0..n)
            .map(|i| (2.0 * PI * hz * i as f32 / sample_rate as f32).sin() * 0.2)
            .collect()
    }

    #[test]
    fn superflux_fires_on_kick_not_slow_sine() {
        let mut kick_an = Analyzer::new(48_000);
        let mut sine_an = Analyzer::new(48_000);
        let zeros = vec![0.0f32; DSP_SIZE];
        let kick = kick_burst(DSP_SIZE, 48_000);
        let tone = sine(DSP_SIZE, 48_000, 440.0);
        for _ in 0..4 {
            kick_an.process_window(&zeros);
            sine_an.process_window(&zeros);
        }
        for _ in 0..6 {
            sine_an.process_window(&tone);
        }
        let sine_flux = sine_an.process_window(&tone).flux;
        let kick_flux = kick_an.process_window(&kick).flux;
        assert!(kick_flux > sine_flux * 1.8, "kick={kick_flux} sine={sine_flux}");
        assert!(kick_flux > 0.05, "kick flux too small: {kick_flux}");
    }

    #[test]
    fn bands_and_eq_are_finite() {
        let mut an = Analyzer::new(48_000);
        let tone = sine(DSP_SIZE, 48_000, 220.0);
        let frame = an.process_window(&tone);
        assert!(frame.bands4.iter().all(|v| v.is_finite()));
        assert!(frame.eq24.iter().all(|v| v.is_finite()));
        assert!(frame.flatness.is_finite());
    }

    #[test]
    fn seq_hops_do_not_drop_a_33ms_frame() {
        let mut an = Analyzer::new(48_000);
        let ring = vec![0.01f32; 8192];
        an.observe_seq(&ring, 256);
        assert_eq!(an.last_hops(), 1);
        // 33 ms at 48 kHz ≈ 1584 samples.
        an.observe_seq(&ring, 256 + 1584);
        assert_eq!(an.last_hops(), 6, "1584 samples should yield six 256-sample hops");
        assert!(an.last_hops() * an.hop() <= 1584);
        assert_eq!(an.hop_fluxes().len(), 6);
    }

    #[test]
    fn studio_separates_45_from_90_better_than_2048() {
        let sr = 48_000u32;
        let tone = sine(STUDIO_SIZE, sr, 45.0);
        let mut accurate = Analyzer::new(sr);
        let mut studio = Analyzer::studio(sr);
        for _ in 0..4 {
            accurate.process_window(&tone);
            studio.process_window(&tone);
        }
        let ratio = |an: &Analyzer| {
            let a = triangle_mag(an.mag(), an.sample_rate(), an.fft_size(), 30.0, 45.0, 62.0);
            let b = triangle_mag(an.mag(), an.sample_rate(), an.fft_size(), 75.0, 90.0, 110.0);
            a / b.max(1e-9)
        };
        let r2048 = ratio(&accurate);
        let r4096 = ratio(&studio);
        assert!(
            r4096 > r2048 * 1.05,
            "studio 45/90 {r4096} should beat 2048 {r2048}"
        );
    }
}
