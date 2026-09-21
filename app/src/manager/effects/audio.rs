use std::{
    f32::consts::PI,
    sync::{
        atomic::Ordering,
        Arc, Mutex,
    },
    thread,
    time::{Duration, Instant},
};

use rand::Rng;

use crate::{
    enums::{AudioColorMode, AudioStyle, Effects, RippleKind, RippleOrigin, RippleTint, RippleTrigger},
    manager::{
        effects::{audio_color, lamps},
        profile::Profile,
        Inner,
    },
};

const FFT_SIZE: usize = 1024;
const FAST_WIN: usize = 256;
const SAMPLE_RATE: u32 = 48_000;
pub const DEFAULT_AUDIO_RGB: [u8; 12] = [255, 24, 48, 255, 140, 16, 36, 220, 120, 72, 120, 255];

#[derive(Clone, Copy, Debug)]
pub struct AudioReactParams {
    pub sensitivity: f32,
    pub smoothness: f32,
    pub min_brightness: u8,
    pub idle_brightness: u8,
    pub bass: f32,
    pub mid: f32,
    pub treble: f32,
    pub presence: f32,
    pub squelch: f32,
    pub punch: f32,
    pub contrast: f32,
    pub spread: f32,
    pub hue_shift: f32,
    pub color_ramp: f32,
    pub motion: f32,
    pub follow_system_volume: bool,
    pub ripple_color: bool,
    pub ripple_strength: f32,
    pub ripple_speed: f32,
    pub ripple_width: f32,
    pub ripple_twist: f32,
    pub ripple_origin: RippleOrigin,
    pub ripple_tint: RippleTint,
    pub ripple_rgb: [u8; 3],
    pub ripple_kind: RippleKind,
    pub ripple_trigger: RippleTrigger,
    pub ripple_shockwave: bool,
    pub ripple_shock_strength: f32,
    pub ripple_shock_sensitivity: f32,
    pub color_mode: AudioColorMode,
    pub style: AudioStyle,
    pub custom_rgb: [u8; 12],
}

impl AudioReactParams {
    pub fn from_effect(effect: Effects) -> Self {
        match effect {
            Effects::AudioReact {
                sensitivity,
                smoothness,
                min_brightness,
                idle_brightness,
                bass,
                mid,
                treble,
                presence,
                squelch,
                punch,
                contrast,
                spread,
                hue_shift,
                color_ramp,
                motion,
                follow_system_volume,
                ripple_color,
                ripple_strength,
                ripple_speed,
                ripple_width,
                ripple_twist,
                ripple_origin,
                ripple_tint,
                ripple_rgb,
                ripple_kind,
                ripple_trigger,
                ripple_shockwave,
                ripple_shock_strength,
                ripple_shock_sensitivity,
                color_mode,
                style,
            } => Self {
                sensitivity,
                smoothness,
                min_brightness,
                idle_brightness: idle_brightness.unwrap_or(min_brightness),
                bass,
                mid,
                treble,
                presence,
                squelch,
                punch,
                contrast,
                spread,
                hue_shift,
                color_ramp,
                motion,
                follow_system_volume,
                ripple_color: ripple_color || matches!(style, AudioStyle::Ripple),
                ripple_strength,
                ripple_speed,
                ripple_width,
                ripple_twist,
                ripple_origin,
                ripple_tint,
                ripple_rgb,
                ripple_kind,
                ripple_trigger,
                ripple_shockwave,
                ripple_shock_strength,
                ripple_shock_sensitivity,
                color_mode,
                style: if matches!(style, AudioStyle::Ripple) {
                    AudioStyle::Levels
                } else {
                    style
                },
                custom_rgb: DEFAULT_AUDIO_RGB,
            }
            .normalized(),
            _ => Self::default(),
        }
    }

    pub fn with_rgb(mut self, rgb: [u8; 12]) -> Self {
        self.custom_rgb = if rgb.iter().all(|&c| c == 0) {
            DEFAULT_AUDIO_RGB
        } else {
            rgb
        };
        self
    }

    pub fn normalized(self) -> Self {
        let or_one = |v: f32| if v <= 0.001 { 1.0 } else { v };
        Self {
            sensitivity: if self.sensitivity <= 0.05 {
                1.2
            } else {
                self.sensitivity.clamp(0.2, 5.0)
            },
            smoothness: self.smoothness.clamp(0.0, 0.95),
            min_brightness: self.min_brightness.min(80),
            idle_brightness: self.idle_brightness.min(80),
            bass: or_one(self.bass).clamp(0.0, 2.5),
            mid: or_one(self.mid).clamp(0.0, 2.5),
            treble: or_one(self.treble).clamp(0.0, 2.5),
            presence: or_one(self.presence).clamp(0.0, 2.5),
            squelch: self.squelch.clamp(0.0, 0.35),
            punch: self.punch.clamp(0.0, 2.0),
            contrast: if self.contrast <= 0.05 { 1.15 } else { self.contrast.clamp(0.5, 2.2) },
            spread: self.spread.clamp(0.0, 1.0),
            hue_shift: self.hue_shift.clamp(0.0, 1.0),
            color_ramp: self.color_ramp.clamp(0.0, 1.0),
            motion: self.motion.clamp(0.0, 2.0),
            follow_system_volume: self.follow_system_volume,
            ripple_color: self.ripple_color,
            ripple_strength: self.ripple_strength.clamp(0.0, 2.0),
            ripple_speed: self.ripple_speed.clamp(0.2, 2.5),
            ripple_width: self.ripple_width.clamp(0.1, 1.0),
            ripple_twist: self.ripple_twist.clamp(0.0, 1.5),
            ripple_origin: self.ripple_origin,
            ripple_tint: self.ripple_tint,
            ripple_rgb: self.ripple_rgb,
            ripple_kind: self.ripple_kind,
            ripple_trigger: self.ripple_trigger,
            ripple_shockwave: self.ripple_shockwave,
            ripple_shock_strength: self.ripple_shock_strength.clamp(0.3, 2.0),
            ripple_shock_sensitivity: self.ripple_shock_sensitivity.clamp(0.0, 1.0),
            color_mode: self.color_mode,
            style: self.style,
            custom_rgb: if self.custom_rgb.iter().all(|&c| c == 0) {
                DEFAULT_AUDIO_RGB
            } else {
                self.custom_rgb
            },
        }
    }
}

impl Default for AudioReactParams {
    fn default() -> Self {
        Self {
            sensitivity: 1.2,
            smoothness: 0.72,
            min_brightness: 6,
            idle_brightness: 6,
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
            custom_rgb: DEFAULT_AUDIO_RGB,
        }
    }
}

pub fn play(manager: &mut Inner, profile: &Profile) {
    let params = manager.audio_params.clone();
    {
        let mut guard = params.lock().unwrap();
        *guard = AudioReactParams::from_effect(profile.effect).with_rgb(profile.rgb_array());
    }

    let samples = Arc::new(Mutex::new(vec![0.0f32; FFT_SIZE]));
    let capture_ok = Arc::new(std::sync::atomic::AtomicBool::new(false));
    let stop = manager.stop_signals.manager_stop_signal.clone();
    let capture_stop = stop.clone();
    let capture_samples = samples.clone();
    let capture_flag = capture_ok.clone();

    let capture = thread::spawn(move || {
        #[cfg(target_os = "windows")]
        {
            if let Err(err) = capture_loopback(capture_samples, capture_stop, capture_flag) {
                legion_rgb_driver::debug_log(&format!("AUDIO: loopback failed: {err}"));
            }
        }
        #[cfg(not(target_os = "windows"))]
        {
            let _ = (capture_samples, capture_stop, capture_flag);
            legion_rgb_driver::debug_log("AUDIO: loopback is only available on Windows");
        }
    });

    legion_rgb_driver::debug_log("EFFECT: audio react (system playback)");

    #[cfg(target_os = "windows")]
    {
        let _ = wasapi::initialize_mta();
    }

    let mut short_avg = [0.0f32; 4];
    let mut long_avg = [0.02f32; 4];
    let mut prev_short = [0.0f32; 4];
    let mut levels = [0.0f32; 4];
    let mut spec_short: Vec<f32> = Vec::new();
    let mut spec_long: Vec<f32> = Vec::new();
    let mut spec_prev: Vec<f32> = Vec::new();
    let mut spec_levels: Vec<f32> = Vec::new();
    let mut spec_beat_env: Vec<f32> = Vec::new();
    let mut spec_hold: Vec<f32> = Vec::new();
    let mut wave_phase = 0.0f32;
    let mut ramp_phase = 0.0f32;
    let mut tint_state: Vec<[f32; 3]> = Vec::new();
    let mut sparkle = [0.0f32; 4];
    let mut sparkle_lamps: Vec<f32> = Vec::new();
    let mut chase = 0.0f32;
    let mut strobe = 0.0f32;
    let mut vu_hold = 0.0f32;
    let mut vu_peak = 0.0f32;
    let mut tempo_phase = 0.0f32;
    let mut color_ripples: Vec<ColorRipple> = Vec::new();
    let mut prev_gate = [0.0f32; 4];
    let mut beat_env = [0.12f32; 4];
    let mut fast_env = 0.0f32;
    let mut fast_peak = 0.002f32;
    let mut prev_loud = 0.0f32;
    let mut drop_loud = 0.0f32;
    let mut drop_trough = 1.0f32;
    let mut shock_cool = 0.0f32;
    let mut beats = BeatTracker::new();
    let mut last_tick = Instant::now();
    let mut last_log = Instant::now();
    let mut last_vol_poll = Instant::now() - Duration::from_secs(1);
    let mut vol_gain = 1.0f32;
    let mut idle_mix = 1.0f32;
    let mut rng = rand::rng();

    while !manager.stop_signals.manager_stop_signal.load(Ordering::SeqCst) {
        let now = Instant::now();
        let dt = now.saturating_duration_since(last_tick).as_secs_f32().clamp(0.004, 0.04);
        last_tick = now;
        let p = params.lock().unwrap().normalized();
        if p.follow_system_volume {
            if last_vol_poll.elapsed() >= Duration::from_millis(80) {
                vol_gain = system_output_gain();
                last_vol_poll = Instant::now();
            }
        } else {
            vol_gain = 1.0;
        }
        let volume_dead = p.follow_system_volume && vol_gain <= 0.0;
        let beat_gates = matches!(p.style, AudioStyle::BeatGates);
        let n = manager.lamp_n();
        let (frame, spec_frame, kick_raw, fast_rms) = {
            let guard = samples.lock().unwrap();
            let rms = recent_rms(&guard);
            if n > 4 {
                let (spec, kick) = analyze_spectrum(&guard, SAMPLE_RATE, n);
                ([0.0f32; 4], spec, kick, rms)
            } else {
                let (bands, kick) = analyze_bands(&guard, SAMPLE_RATE);
                (bands, Vec::new(), kick, rms)
            }
        };

        // Quiet songs / low Windows volume still have a beat shape; only treat
        // near-digital-silence as empty. The old 0.008 peak floor made 5–20%
        // volume look like "nothing playing".
        let noise_floor = (2.5e-5 + p.squelch * 1.2e-4).max(1.5e-5);
        fast_env = ema_toward(fast_env, fast_rms, dt, 0.01);
        if fast_env > fast_peak {
            fast_peak += (fast_env - fast_peak) * 0.55;
        } else {
            let drop_tau = if fast_env < fast_peak * 0.35 { 0.12 } else { 0.55 };
            fast_peak = ema_toward(fast_peak, fast_env, dt, drop_tau).max(noise_floor);
        }
        let loud = ((fast_env / fast_peak.max(noise_floor)) * (0.55 + p.sensitivity * 0.5)).clamp(0.0, 1.0);
        let flux_fast = (loud - prev_loud).max(0.0);
        prev_loud = loud;

        let boosted = [
            frame[0] * p.bass,
            frame[1] * p.mid,
            frame[2] * p.treble,
            frame[3] * p.presence,
        ];
        let rel_gate = if beat_gates {
            (p.squelch * 1.15).max(0.01)
        } else {
            (p.squelch * 1.6).max(0.012)
        };
        let gated = volume_dead || fast_env < noise_floor * 1.6 || loud < rel_gate;
        let env_floor = (fast_env * 0.18 + noise_floor).max(noise_floor);
        let short_tau = if beat_gates {
            0.01
        } else {
            0.012 + p.smoothness * 0.014
        };
        let long_tau = 0.55;
        let attack_tau = 0.008 + p.smoothness * 0.012;
        let release_tau = 0.035 + p.smoothness * 0.14;

        if n > 4 {
            resize_spec(&mut spec_short, n, 0.0);
            resize_spec(&mut spec_long, n, 0.02);
            resize_spec(&mut spec_prev, n, 0.0);
            resize_spec(&mut spec_levels, n, 0.0);
            resize_spec(&mut spec_beat_env, n, 0.12);
            resize_spec(&mut spec_hold, n, 0.0);
            follow_spectrum(
                &spec_frame,
                &mut spec_short,
                &mut spec_long,
                &mut spec_prev,
                &mut spec_levels,
                &mut spec_beat_env,
                &mut spec_hold,
                &p,
                beat_gates,
                gated,
                dt,
                short_tau,
                long_tau,
                attack_tau,
                release_tau,
                env_floor,
                loud,
                flux_fast,
            );
        } else if beat_gates {
            let mut band = [0.0f32; 4];
            let mut band_flux = [0.0f32; 4];
            for i in 0..4 {
                let raw = boosted[i].max(0.0);
                short_avg[i] = ema_toward(short_avg[i], raw, dt, short_tau);
                long_avg[i] = ema_toward(long_avg[i], raw.max(env_floor), dt, long_tau).max(env_floor);
                let shape = short_avg[i] / long_avg[i];
                band[i] = (shape / (shape + 0.7)).clamp(0.0, 1.0);
                band_flux[i] = ((short_avg[i] - prev_short[i]).max(0.0) / long_avg[i]).clamp(0.0, 2.0);
                prev_short[i] = short_avg[i];
                beat_env[i] = ema_toward(beat_env[i], band[i], dt, 0.22).max(0.05);
            }
            let strongest = band
                .iter()
                .enumerate()
                .max_by(|a, b| a.1.total_cmp(b.1))
                .map(|(i, _)| i)
                .unwrap_or(0);
            let kick = flux_fast > (0.032 - p.punch * 0.01).max(0.016);
            let hold = (0.065 + p.smoothness * 0.05).clamp(0.05, 0.14);
            for i in 0..4 {
                let thresh = (beat_env[i] * 1.16 + 0.03 + p.squelch * 0.12).clamp(0.07, 0.5);
                let local_hit = !gated && band[i] > thresh && band_flux[i] > 0.012;
                let kick_hit = !gated && kick && i == strongest && band[i] > beat_env[i] * 0.72;
                let hit = local_hit || kick_hit;
                if hit && (sparkle[i] < 0.018 || band_flux[i] > 0.05 || kick_hit) {
                    sparkle[i] = hold;
                } else {
                    sparkle[i] = (sparkle[i] - dt).max(0.0);
                }
                levels[i] = if sparkle[i] > 0.0 { 1.0 } else { 0.0 };
            }
        } else {
            for i in 0..4 {
                let raw = boosted[i].max(0.0);
                short_avg[i] = ema_toward(short_avg[i], raw, dt, short_tau);
                long_avg[i] = ema_toward(long_avg[i], raw.max(env_floor), dt, long_tau).max(env_floor);
                let shape = short_avg[i] / long_avg[i];
                let band = (shape / (shape + 0.7)).clamp(0.0, 1.0);
                let flux = ((short_avg[i] - prev_short[i]).max(0.0) / long_avg[i]) + flux_fast;
                prev_short[i] = short_avg[i];
                let target = if gated {
                    0.0
                } else {
                    (loud * (0.28 + 0.72 * band) + flux * p.punch * 0.9).clamp(0.0, 1.0)
                };
                let tau = if target > levels[i] { attack_tau } else { release_tau };
                levels[i] = ema_toward(levels[i], target, dt, tau);
            }
        }

        let idle_tau = if gated {
            0.05 + p.smoothness * 0.18
        } else {
            0.02 + p.smoothness * 0.04
        };
        idle_mix = ema_toward(idle_mix, if gated { 1.0 } else { 0.0 }, dt, idle_tau);
        drop_loud = ema_toward(drop_loud, loud, dt, 0.2);
        if loud < drop_trough {
            drop_trough = loud;
        } else {
            drop_trough = ema_toward(drop_trough, drop_loud.min(loud), dt, 1.05);
        }
        shock_cool = (shock_cool - dt).max(0.0);
        let bass_now = if n > 4 {
            spec_levels.iter().take((n / 5).max(1)).copied().fold(0.0f32, f32::max)
        } else {
            levels[0]
        };
        let band_peak = if n > 4 {
            spec_levels.iter().copied().fold(0.0f32, f32::max)
        } else {
            levels.iter().copied().fold(0.0f32, f32::max)
        };
        let rise = (loud - drop_trough).max(0.0);
        let kick_now = (kick_raw * p.bass).clamp(0.0, 2.5);
        let hits = if p.ripple_color && !volume_dead {
            beats.tick(dt, gated, flux_fast, bass_now, kick_now, band_peak, p.punch, p.squelch)
        } else {
            beats.reset();
            BeatHits::none()
        };
        let drop_hit = if p.ripple_color && p.ripple_shockwave && (hits.kick || hits.bass) && shock_cool <= 0.0 {
            let sens = p.ripple_shock_sensitivity;
            let rise_need = (0.30 - sens * 0.14).clamp(0.16, 0.36);
            rise > rise_need && loud > 0.2
        } else {
            false
        };
        if drop_hit {
            shock_cool = 0.5;
            drop_trough = (drop_loud * 0.85 + loud * 0.15).min(loud);
        }
        let ring_hit = match p.ripple_trigger {
            RippleTrigger::All => hits.any,
            RippleTrigger::Bass => hits.bass,
            RippleTrigger::Kick => hits.kick,
            RippleTrigger::TripleKick => hits.triple,
        };
        ramp_phase = (ramp_phase + dt * (0.12 + p.motion * 0.55)) % (2.0 * std::f32::consts::PI);
        if tint_state.len() != n {
            tint_state.resize(n, [0.0; 3]);
        }

        let level_slice: &[f32] = if n > 4 { &spec_levels } else { &levels };
        let rgb = if volume_dead {
            sparkle = [0.0; 4];
            sparkle_lamps.fill(0.0);
            levels = [0.0; 4];
            spec_levels.fill(0.0);
            spec_hold.fill(0.0);
            strobe = 0.0;
            vu_hold = 0.0;
            vu_peak = 0.0;
            tempo_phase = 0.0;
            color_ripples.clear();
            idle_mix = 1.0;
            drop_trough = 1.0;
            drop_loud = 0.0;
            shock_cool = 0.0;
            beats.reset();
            vec![[0u8; 3]; n]
        } else {
            render_audio(
                p,
                level_slice,
                idle_mix,
                flux_fast,
                &mut wave_phase,
                ramp_phase,
                &mut tint_state,
                &mut sparkle,
                &mut sparkle_lamps,
                &mut chase,
                &mut strobe,
                &mut vu_hold,
                &mut vu_peak,
                &mut tempo_phase,
                &mut color_ripples,
                &mut prev_gate,
                dt,
                n,
                drop_hit,
                ring_hit,
                beats.ibi(),
                beats.tempo_locked(),
                &mut rng,
            )
        };
        manager.paint_lamps(&rgb);

        if last_log.elapsed() > Duration::from_secs(3) {
            let (b0, b1, b2, b3) = if n > 4 && spec_levels.len() == n {
                (
                    spec_levels[0],
                    spec_levels[n / 4],
                    spec_levels[n / 2],
                    spec_levels[n - 1],
                )
            } else {
                (levels[0], levels[1], levels[2], levels[3])
            };
            legion_rgb_driver::debug_log(&format!(
                "AUDIO: capture={} loud={:.2} bass={:.2} mid={:.2} treble={:.2} presence={:.2} vol={:.2} follow={} lamps={}",
                capture_ok.load(Ordering::Relaxed),
                loud,
                b0,
                b1,
                b2,
                b3,
                vol_gain,
                p.follow_system_volume,
                n
            ));
            last_log = Instant::now();
        }

        thread::sleep(Duration::from_millis(if n > 4 { 33 } else { 8 }));
    }

    let _ = capture.join();
}

struct BeatHits {
    any: bool,
    bass: bool,
    kick: bool,
    triple: bool,
}

impl BeatHits {
    fn none() -> Self {
        Self {
            any: false,
            bass: false,
            kick: false,
            triple: false,
        }
    }
}

struct BeatTracker {
    onset_slow: f32,
    onset_p: f32,
    bass_env: f32,
    kick_env: f32,
    since: f32,
    ibi: f32,
    cool: f32,
    kick_age: [f32; 3],
    kick_n: u8,
    got_beat: bool,
}

impl BeatTracker {
    fn new() -> Self {
        Self {
            onset_slow: 0.02,
            onset_p: 0.0,
            bass_env: 0.1,
            kick_env: 0.08,
            since: 1.0,
            ibi: 0.5,
            cool: 0.0,
            kick_age: [99.0; 3],
            kick_n: 0,
            got_beat: false,
        }
    }

    fn ibi(&self) -> f32 {
        self.ibi
    }

    fn tempo_locked(&self) -> bool {
        self.got_beat && self.since < 1.2
    }

    fn reset(&mut self) {
        *self = Self::new();
    }

    fn tick(
        &mut self,
        dt: f32,
        gated: bool,
        flux: f32,
        bass: f32,
        kick: f32,
        peak: f32,
        punch: f32,
        squelch: f32,
    ) -> BeatHits {
        self.since = (self.since + dt).min(4.0);
        self.cool = (self.cool - dt).max(0.0);
        for age in &mut self.kick_age {
            *age = (*age + dt).min(8.0);
        }
        if gated {
            self.onset_p = 0.0;
            return BeatHits::none();
        }

        let onset = flux.max(0.0);
        self.onset_slow = ema_toward(self.onset_slow, onset, dt, 0.32).max(0.008);
        let thresh = self.onset_slow * (1.7 - punch * 0.22).clamp(1.25, 1.95) + 0.014 + squelch * 0.04;
        self.bass_env = ema_toward(self.bass_env, bass, dt, 0.05).max(0.04);
        self.kick_env = ema_toward(self.kick_env, kick, dt, 0.07).max(0.03);
        let bass_share = bass / peak.max(1e-3);
        let bass_jump = (bass - self.bass_env).max(0.0);
        let kick_on = (kick - self.kick_env).max(0.0);
        let kick_share = kick / peak.max(1e-3);
        let crossed = onset > thresh && self.onset_p <= thresh;
        let strong = onset > thresh && onset > self.onset_p && (onset - self.onset_p) > thresh * 0.4;
        let min_gap = (self.ibi * 0.56).clamp(0.22, 0.36);
        let ready = self.cool <= 0.0 && self.since >= min_gap;
        let hit = ready && (crossed || strong) && onset > 0.016;
        self.onset_p = onset;

        let mut hits = BeatHits::none();
        if hit {
            hits.any = true;
            hits.bass = bass_share >= 0.38 && (bass > 0.11 || bass_jump > 0.04);
            hits.kick = kick_on > 0.018
                && kick > 0.07
                && kick_share >= 0.1
                && bass_share >= 0.28
                && kick_on >= bass_jump * 0.35;
            if hits.kick {
                hits.bass = true;
            }
            self.accept();
        }
        hits.triple = if hits.kick { self.note_kick() } else { false };
        hits
    }

    fn accept(&mut self) {
        if self.since > 0.2 && self.since < 1.15 {
            self.ibi = (self.ibi * 0.62 + self.since * 0.38).clamp(0.28, 0.85);
        }
        self.cool = (self.ibi * 0.56).clamp(0.22, 0.36);
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

fn ema_toward(current: f32, target: f32, dt: f32, tau: f32) -> f32 {
    let alpha = 1.0 - (-dt / tau.max(0.006)).exp();
    current + (target - current) * alpha.clamp(0.0, 1.0)
}

/// Windows master volume (0..=1). Mute or <=5% is 0 so lights can fully stop.
/// Loopback RMS is often pre-volume, so this must come from IAudioEndpointVolume.
fn system_output_gain() -> f32 {
    #[cfg(target_os = "windows")]
    {
        use windows::Win32::Media::Audio::Endpoints::IAudioEndpointVolume;
        use windows::Win32::Media::Audio::{eConsole, eRender, IMMDeviceEnumerator, MMDeviceEnumerator};
        use windows::Win32::System::Com::{CoCreateInstance, CLSCTX_ALL};

        const STOP_AT: f32 = 0.05;
        unsafe {
            let enumerator: IMMDeviceEnumerator = match CoCreateInstance(&MMDeviceEnumerator, None, CLSCTX_ALL) {
                Ok(v) => v,
                Err(_) => return 1.0,
            };
            let device = match enumerator.GetDefaultAudioEndpoint(eRender, eConsole) {
                Ok(v) => v,
                Err(_) => return 1.0,
            };
            let volume: IAudioEndpointVolume = match device.Activate(CLSCTX_ALL, None) {
                Ok(v) => v,
                Err(_) => return 1.0,
            };
            if volume.GetMute().map(|m| m.as_bool()).unwrap_or(false) {
                return 0.0;
            }
            match volume.GetMasterVolumeLevelScalar() {
                Ok(scalar) if scalar <= STOP_AT => 0.0,
                Ok(scalar) => scalar.clamp(0.0, 1.0),
                Err(_) => 1.0,
            }
        }
    }
    #[cfg(not(target_os = "windows"))]
    {
        1.0
    }
}

fn recent_rms(samples: &[f32]) -> f32 {
    if samples.is_empty() {
        return 0.0;
    }
    let n = FAST_WIN.min(samples.len());
    let start = samples.len() - n;
    let acc = samples[start..].iter().map(|s| s * s).sum::<f32>();
    (acc / n as f32).sqrt()
}

fn led_curve(energy: f32, contrast: f32) -> f32 {
    energy.clamp(0.0, 1.0).powf(contrast.clamp(0.5, 2.2))
}

fn resize_spec(v: &mut Vec<f32>, n: usize, fill: f32) {
    if v.len() != n {
        v.resize(n, fill);
    }
}

fn freq_gain(params: &AudioReactParams, t: f32) -> f32 {
    lamps::sample_bands(&[params.bass, params.mid, params.treble, params.presence], t).max(0.0)
}

fn follow_spectrum(
    frame: &[f32],
    short: &mut [f32],
    long: &mut [f32],
    prev: &mut [f32],
    levels: &mut [f32],
    beat_env: &mut [f32],
    hold: &mut [f32],
    p: &AudioReactParams,
    beat_gates: bool,
    gated: bool,
    dt: f32,
    short_tau: f32,
    long_tau: f32,
    attack_tau: f32,
    release_tau: f32,
    env_floor: f32,
    loud: f32,
    flux_fast: f32,
) {
    let n = frame.len().min(short.len()).min(levels.len());
    if n == 0 {
        return;
    }
    let mut band = vec![0.0f32; n];
    let mut band_flux = vec![0.0f32; n];
    for i in 0..n {
        let raw = frame[i].max(0.0) * freq_gain(p, lamps::pos(i, n));
        short[i] = ema_toward(short[i], raw, dt, short_tau);
        long[i] = ema_toward(long[i], raw.max(env_floor), dt, long_tau).max(env_floor);
        let shape = short[i] / long[i];
        band[i] = (shape / (shape + 0.7)).clamp(0.0, 1.0);
        band_flux[i] = ((short[i] - prev[i]).max(0.0) / long[i]).clamp(0.0, 2.0);
        prev[i] = short[i];
    }
    if beat_gates {
        for i in 0..n {
            beat_env[i] = ema_toward(beat_env[i], band[i], dt, 0.22).max(0.05);
        }
        let strongest = band
            .iter()
            .enumerate()
            .max_by(|a, b| a.1.total_cmp(b.1))
            .map(|(i, _)| i)
            .unwrap_or(0);
        let kick = flux_fast > (0.032 - p.punch * 0.01).max(0.016);
        let hold_t = (0.065 + p.smoothness * 0.05).clamp(0.05, 0.14);
        for i in 0..n {
            let thresh = (beat_env[i] * 1.16 + 0.03 + p.squelch * 0.12).clamp(0.07, 0.5);
            let local_hit = !gated && band[i] > thresh && band_flux[i] > 0.012;
            let kick_hit = !gated && kick && i == strongest && band[i] > beat_env[i] * 0.72;
            let hit = local_hit || kick_hit;
            if hit && (hold[i] < 0.018 || band_flux[i] > 0.05 || kick_hit) {
                hold[i] = hold_t;
            } else {
                hold[i] = (hold[i] - dt).max(0.0);
            }
            levels[i] = if hold[i] > 0.0 { 1.0 } else { 0.0 };
        }
    } else {
        for i in 0..n {
            let flux = band_flux[i] + flux_fast;
            let target = if gated {
                0.0
            } else {
                (loud * (0.28 + 0.72 * band[i]) + flux * p.punch * 0.9).clamp(0.0, 1.0)
            };
            let tau = if target > levels[i] { attack_tau } else { release_tau };
            levels[i] = ema_toward(levels[i], target, dt, tau);
        }
    }
}

fn render_audio(
    params: AudioReactParams,
    levels: &[f32],
    idle_mix: f32,
    flux: f32,
    wave_phase: &mut f32,
    ramp_phase: f32,
    tint_state: &mut [[f32; 3]],
    _sparkle: &mut [f32; 4],
    sparkle_lamps: &mut Vec<f32>,
    chase: &mut f32,
    strobe: &mut f32,
    vu_hold: &mut f32,
    vu_peak: &mut f32,
    tempo_phase: &mut f32,
    color_ripples: &mut Vec<ColorRipple>,
    _prev_gate: &mut [f32; 4],
    dt: f32,
    lamp_count: usize,
    drop_hit: bool,
    ring_hit: bool,
    tempo_ibi: f32,
    tempo_locked: bool,
    rng: &mut impl rand::Rng,
) -> Vec<[u8; 3]> {
    let n = lamp_count.max(1);
    if sparkle_lamps.len() != n {
        sparkle_lamps.resize(n, 0.0);
    }
    let beat_floor = params.min_brightness as f32 / 100.0;
    let idle = params.idle_brightness as f32 / 100.0;
    let rest = idle_mix.clamp(0.0, 1.0);
    let motion = 0.35 + params.motion * 1.4;
    let fine = n > 4 && levels.len() >= n;
    let four = to_four(levels);
    let drive = if fine {
        levels.iter().take((n / 4).max(1)).copied().fold(0.0f32, f32::max)
    } else {
        four[0].max(four[1])
    };
    let peak = levels.iter().copied().fold(0.0f32, f32::max);
    *wave_phase = (*wave_phase + dt * (0.35 + motion * (0.6 + drive * 2.4))) % (2.0 * PI);

    let mut energies = if fine {
        levels[..n].to_vec()
    } else if matches!(params.style, AudioStyle::BeatGates | AudioStyle::Vu | AudioStyle::TempoPulse) {
        four.to_vec()
    } else {
        spatial_blur(&four, params.spread).to_vec()
    };
    if fine && !matches!(params.style, AudioStyle::BeatGates | AudioStyle::Vu | AudioStyle::TempoPulse) {
        spatial_blur_n(&mut energies, params.spread);
    }

    match params.style {
        AudioStyle::Levels | AudioStyle::Wave | AudioStyle::Gradient | AudioStyle::Center => {}
        AudioStyle::Mirror => {
            if fine {
                let orig = energies.clone();
                for i in 0..n {
                    let src = i.min(n - 1 - i);
                    energies[i] = orig[src];
                }
            }
        }
        AudioStyle::Pulse => {
            let pulse = if fine {
                let bass = energies.iter().take((n / 4).max(1)).copied().fold(0.0f32, f32::max);
                (energies.iter().sum::<f32>() / n as f32).max(bass * 0.85)
            } else {
                (four.iter().sum::<f32>() / 4.0).max(four[0] * 0.85)
            };
            energies.fill(pulse);
        }
        AudioStyle::Bloom => {
            if fine {
                cascade_bloom(&mut energies);
            } else {
                energies.resize(4, 0.0);
                energies[0] = four[0];
                energies[1] = four[1].max(four[0] * 0.62);
                energies[2] = four[2].max(four[0] * 0.32 + four[1] * 0.45);
                energies[3] = four[3].max(four[1] * 0.28 + four[2] * 0.40);
            }
        }
        AudioStyle::Fire => {
            if fine {
                cascade_fire(&mut energies);
            } else {
                energies.resize(4, 0.0);
                energies[0] = four[0];
                energies[1] = (four[0] * 0.75 + four[1] * 0.55).min(1.0);
                energies[2] = (four[1] * 0.55 + four[2] * 0.7).min(1.0);
                energies[3] = (four[2] * 0.4 + four[3] * 0.85).min(1.0);
            }
        }
        AudioStyle::Strobe => {
            *strobe *= (1.0 - dt * (12.0 + motion * 6.0)).max(0.0);
            let hit = flux > (0.038 - params.punch * 0.012).max(0.016) && peak > 0.08;
            if hit {
                *strobe = (0.72 + peak * 0.28 + params.punch * 0.12).clamp(0.75, 1.0);
            }
            energies.fill(*strobe);
        }
        AudioStyle::Sparkle => {
            for spark in sparkle_lamps.iter_mut() {
                *spark *= (1.0 - dt * (5.5 + motion * 3.0)).max(0.0);
            }
            if (peak > 0.18 || flux > 0.03) && rng.random::<f32>() < (0.1 + peak * 0.4 + params.punch * 0.1) {
                let z = rng.random_range(0..n);
                sparkle_lamps[z] = sparkle_lamps[z].max(0.55 + peak * 0.45);
            }
        }
        AudioStyle::Chase => {
            *chase = (*chase + dt * (1.1 + motion * 3.8 + drive * 2.2) * n as f32 / 4.0) % n as f32;
        }
        AudioStyle::Vu => {
            if peak > *vu_hold {
                *vu_hold = peak;
            } else {
                *vu_hold = ema_toward(*vu_hold, peak, dt, 0.08);
            }
            if *vu_hold > *vu_peak {
                *vu_peak = *vu_hold;
            } else {
                *vu_peak = ema_toward(*vu_peak, *vu_hold, dt, 0.55);
            }
            let fill_x = vu_hold.clamp(0.0, 1.0) * n as f32;
            let mark = (vu_peak.clamp(0.0, 1.0) * (n as f32 - 1.0)).round();
            energies.resize(n, 0.0);
            for i in 0..n {
                let bar = (fill_x - i as f32).clamp(0.0, 1.0);
                let tip = if (i as f32 - mark).abs() < 0.51 { 0.92 } else { 0.0 };
                energies[i] = bar.max(tip);
            }
        }
        AudioStyle::TempoPulse => {
            if ring_hit {
                *tempo_phase = 0.0;
            } else if tempo_locked {
                *tempo_phase = (*tempo_phase + dt / tempo_ibi.max(0.28)) % 1.0;
            }
            let pulse = if tempo_locked {
                let shape = (0.5 + 0.5 * (*tempo_phase * 2.0 * PI).cos()).powf(2.2);
                (0.12 + 0.88 * shape).clamp(0.0, 1.0)
            } else if fine {
                let bass = energies.iter().take((n / 4).max(1)).copied().fold(0.0f32, f32::max);
                (energies.iter().sum::<f32>() / n as f32).max(bass * 0.85)
            } else {
                (four.iter().sum::<f32>() / 4.0).max(four[0] * 0.85)
            };
            energies.resize(n, 0.0);
            energies.fill(pulse);
        }
        AudioStyle::BeatGates | AudioStyle::Ripple => {}
    }

    if params.ripple_color {
        let tune = RippleTune::from_params(&params);
        tick_color_ripples(
            color_ripples,
            dt,
            flux,
            peak,
            drive,
            &four,
            levels,
            n,
            fine,
            params.punch,
            ramp_phase,
            drop_hit,
            ring_hit,
            &tune,
        );
    } else {
        color_ripples.clear();
    }

    let zone_w = (n as f32 / 4.0).max(1.0);
    let four_e = to_four(&energies);
    (0..n)
        .map(|i| {
            let t = lamps::pos(i, n);
            let local = if fine {
                energies[i]
            } else {
                lamps::sample_bands(&four_e, t)
            };
            let energy = match params.style {
                AudioStyle::Levels | AudioStyle::Ripple => {
                    if fine {
                        local
                    } else {
                        hard_band(&four_e, i, n)
                    }
                }
                AudioStyle::Pulse | AudioStyle::Strobe | AudioStyle::TempoPulse => {
                    energies.first().copied().unwrap_or(0.0)
                }
                AudioStyle::Vu => {
                    if fine {
                        local
                    } else {
                        hard_band(&four_e, i, n)
                    }
                }
                AudioStyle::Wave => {
                    let travel = 0.5 + 0.5 * ((*wave_phase) - t * 2.0 * PI).sin();
                    if fine {
                        (local * 0.28 + travel * drive.max(peak) * 0.78 * (0.4 + 0.6 * local)).clamp(0.0, 1.0)
                    } else {
                        (local * 0.28 + travel * drive.max(peak) * 0.78).clamp(0.0, 1.0)
                    }
                }
                AudioStyle::Bloom | AudioStyle::Fire => {
                    if fine {
                        local
                    } else {
                        lamps::sample_bands(&four_e, t)
                    }
                }
                AudioStyle::Center => {
                    let dist = (t - 0.5).abs() * 2.0;
                    let width = (0.4 + peak * 0.58).clamp(0.4, 0.98);
                    let geo = ((1.0 - dist / width).max(0.0).powf(1.25) * (0.15 + peak * 0.9)).clamp(0.0, 1.0);
                    if fine {
                        (geo * (0.35 + 0.65 * local)).clamp(0.0, 1.0)
                    } else {
                        geo
                    }
                }
                AudioStyle::Mirror => {
                    if fine {
                        local
                    } else {
                        let half = (n / 2).max(2);
                        let src = i.min(n - 1 - i);
                        let t_src = lamps::pos(src, half) * 0.5;
                        lamps::sample_bands(&four_e, t_src)
                    }
                }
                AudioStyle::Sparkle => (local * 0.28 + sparkle_lamps[i]).clamp(0.0, 1.0),
                AudioStyle::Chase => {
                    let dist = (i as f32 - *chase).abs().min(n as f32 - (i as f32 - *chase).abs());
                    let sigma = (zone_w * 0.55).max(0.5);
                    let spot = (-dist * dist / (2.0 * sigma * sigma)).exp();
                    let base = (spot * (0.2 + drive.max(peak) * 0.9)).clamp(0.0, 1.0);
                    if fine {
                        (base * (0.45 + 0.55 * local)).clamp(0.0, 1.0)
                    } else {
                        base
                    }
                }
                AudioStyle::Gradient => {
                    let g = if fine {
                        let mut acc = 0.0;
                        let mut w = 0.0;
                        for (idx, v) in levels.iter().take(n).enumerate() {
                            let wt = 1.0 - lamps::pos(idx, n) * 0.5;
                            acc += *v * wt;
                            w += wt;
                        }
                        (acc / w.max(1e-3)).clamp(0.0, 1.0)
                    } else {
                        (four[0] * 0.45 + four[1] * 0.3 + four[2] * 0.15 + four[3] * 0.1).clamp(0.0, 1.0)
                    };
                    let bump = (0.12 + (1.0 - (t - g).abs()) * (0.45 + peak * 0.55)).clamp(0.0, 1.0);
                    if fine {
                        (bump * (0.4 + 0.6 * local)).clamp(0.0, 1.0)
                    } else {
                        bump
                    }
                }
                AudioStyle::BeatGates => {
                    if fine {
                        local
                    } else {
                        hard_band(&four_e, i, n)
                    }
                }
            };
            let ring = if params.ripple_color {
                ripple_energy(color_ripples, i, t, n, &RippleTune::from_params(&params))
            } else {
                0.0
            };
            let energy = (energy + ring * params.ripple_strength).clamp(0.0, 1.0);
            let beat_amt = if matches!(params.style, AudioStyle::BeatGates | AudioStyle::Strobe) {
                energy.clamp(0.0, 1.0)
            } else {
                (beat_floor + (1.0 - beat_floor) * led_curve(energy, params.contrast)).clamp(0.0, 1.0)
            };
            let amount = (rest * idle + (1.0 - rest) * beat_amt).clamp(0.0, 1.0);
            let ctx = audio_color::ColorCtx {
                custom_rgb: params.custom_rgb,
                t,
                energy,
                phase: *wave_phase,
                ramp_phase,
                motion: params.motion,
            };
            let mut rgb = audio_color::tint(params.color_mode, ctx);
            if matches!(params.style, AudioStyle::Fire) {
                rgb = lamps::mix_rgb(fire_color(t), rgb, 0.32);
            }
            let (mut r, mut g, mut b) = (rgb[0], rgb[1], rgb[2]);
            if params.hue_shift > 0.01
                && audio_color::hue_shift_ok(params.color_mode)
                && !matches!(params.style, AudioStyle::Fire)
            {
                let hue_add = energy * params.hue_shift * 90.0 + *wave_phase * params.hue_shift * 12.0;
                let (h, s, v) = audio_color::rgb_to_hsv(r, g, b);
                (r, g, b) = audio_color::hsv_to_rgb((h + hue_add) % 360.0, s, v);
            }
            if params.ripple_color && ring > 0.02 {
                let tune = RippleTune::from_params(&params);
                let mix = (ring * (0.65 + 0.35 * params.ripple_strength.min(1.0))).clamp(0.0, 1.0);
                match params.ripple_tint {
                    RippleTint::Current => {}
                    RippleTint::ColorChange => {
                        let hue_add = ripple_hue(color_ripples, i, t, n, &tune) * (0.4 + 0.6 * ring);
                        let (h, s, v) = audio_color::rgb_to_hsv(r, g, b);
                        (r, g, b) = audio_color::hsv_to_rgb((h + hue_add).rem_euclid(360.0), s.max(0.5), v);
                    }
                    RippleTint::Custom => {
                        [r, g, b] = lamps::mix_rgb([r, g, b], params.ripple_rgb, mix);
                    }
                    RippleTint::Rainbow => {
                        let hue = ripple_hue(color_ripples, i, t, n, &tune);
                        let (rr, gg, bb) = audio_color::hsv_to_rgb(hue, 0.95, 1.0);
                        [r, g, b] = lamps::mix_rgb([r, g, b], [rr, gg, bb], mix);
                    }
                }
            }
            let rgb = audio_color::ease_rgb(&mut tint_state[i], [r, g, b], dt, params.color_ramp);
            lamps::scale_rgb(rgb, amount)
        })
        .collect()
}

fn to_four(levels: &[f32]) -> [f32; 4] {
    [
        levels.first().copied().unwrap_or(0.0),
        levels.get(1).copied().unwrap_or(0.0),
        levels.get(2).copied().unwrap_or(0.0),
        levels.get(3).copied().unwrap_or(0.0),
    ]
}

fn hard_band(levels: &[f32; 4], i: usize, n: usize) -> f32 {
    levels[(i * 4 / n.max(1)).min(3)]
}

struct ColorRipple {
    origin: f32,
    age: f32,
    hue: f32,
    shock: bool,
}

struct RippleTune {
    speed: f32,
    width: f32,
    twist: f32,
    origin: RippleOrigin,
    shock_strength: f32,
    kind: RippleKind,
    trigger: RippleTrigger,
}

impl RippleTune {
    fn from_params(params: &AudioReactParams) -> Self {
        Self {
            speed: params.ripple_speed,
            width: params.ripple_width,
            twist: params.ripple_twist,
            origin: params.ripple_origin,
            shock_strength: params.ripple_shock_strength,
            kind: params.ripple_kind,
            trigger: params.ripple_trigger,
        }
    }

    fn travel(&self) -> f32 {
        0.35 + self.speed * 1.85
    }

    fn ring_width(&self, n: usize) -> f32 {
        let w = self.width.clamp(0.1, 1.0);
        if n > 4 {
            0.03 + w * 0.14
        } else {
            0.12 + w * 0.28
        }
    }
}

fn ripple_origin_at(
    tune: &RippleTune,
    four: &[f32; 4],
    levels: &[f32],
    n: usize,
    fine: bool,
    shock: bool,
) -> f32 {
    match tune.origin {
        RippleOrigin::Center => 0.5,
        RippleOrigin::Left => {
            if n > 4 {
                0.0
            } else {
                0.08
            }
        }
        RippleOrigin::Right => {
            if n > 4 {
                1.0
            } else {
                0.92
            }
        }
        RippleOrigin::Auto if shock => 0.5,
        RippleOrigin::Auto if matches!(tune.trigger, RippleTrigger::Kick | RippleTrigger::TripleKick) => {
            if n > 4 {
                0.0
            } else {
                0.08
            }
        }
        RippleOrigin::Auto => {
            if fine && levels.len() >= n {
                let i = levels
                    .iter()
                    .take(n)
                    .enumerate()
                    .max_by(|a, b| a.1.total_cmp(b.1))
                    .map(|(i, _)| i)
                    .unwrap_or(n / 2);
                lamps::pos(i, n)
            } else {
                let i = four
                    .iter()
                    .enumerate()
                    .max_by(|a, b| a.1.total_cmp(b.1))
                    .map(|(i, _)| i)
                    .unwrap_or(0);
                [0.12, 0.38, 0.62, 0.88][i]
            }
        }
    }
}

fn tick_color_ripples(
    ripples: &mut Vec<ColorRipple>,
    dt: f32,
    _flux: f32,
    _peak: f32,
    drive: f32,
    four: &[f32; 4],
    levels: &[f32],
    n: usize,
    fine: bool,
    _punch: f32,
    ramp_phase: f32,
    drop_hit: bool,
    ring_hit: bool,
    tune: &RippleTune,
) {
    for ripple in ripples.iter_mut() {
        ripple.age += dt;
    }
    ripples.retain(|ripple| ripple.age < ripple_lifetime(ripple.shock, n, tune));
    let hue = (ramp_phase * (180.0 / PI) + drive * 40.0).rem_euclid(360.0);
    if drop_hit {
        ripples.push(ColorRipple {
            origin: ripple_origin_at(tune, four, levels, n, fine, true),
            age: 0.0,
            hue: (hue + 48.0).rem_euclid(360.0),
            shock: true,
        });
    } else {
        let cooling = ripples.iter().all(|r| r.age > 0.18);
        if ring_hit && cooling {
            ripples.push(ColorRipple {
                origin: ripple_origin_at(tune, four, levels, n, fine, false),
                age: 0.0,
                hue,
                shock: false,
            });
        }
    }
    if ripples.len() > 5 {
        ripples.remove(0);
    }
}

fn ripple_lifetime(shock: bool, n: usize, tune: &RippleTune) -> f32 {
    if n > 4 {
        let max_i = (n as f32 - 1.0).max(1.0);
        let speed = if shock { tune.travel() * 1.65 } else { tune.travel() };
        let strip_speed = (speed * max_i).max(4.0);
        let span = max_i / strip_speed;
        if shock {
            (span + 0.28).clamp(0.9, 1.8)
        } else {
            (span + 0.2).clamp(0.7, 1.4)
        }
    } else if shock {
        1.28
    } else {
        0.95
    }
}

fn origin_strip(origin: f32, n: usize) -> i32 {
    let max = (n as i32 - 1).max(0);
    (origin.clamp(0.0, 1.0) * max as f32).round() as i32
}

fn ripple_sample(ripples: &[ColorRipple], i: usize, t: f32, n: usize, tune: &RippleTune) -> (f32, f32) {
    if n > 4 {
        ripple_sample_strips(ripples, i, n, tune)
    } else {
        ripple_sample_zones(ripples, t, n, tune)
    }
}

fn gauss_ring(dist: f32, radius: f32, width: f32) -> f32 {
    let d = dist - radius;
    let w = width.max(0.02);
    (-(d * d) / (2.0 * w * w)).exp()
}

fn zone_shape(kind: RippleKind, dist: f32, radius: f32, width: f32, fade: f32, shock: bool, age: f32, power: f32) -> f32 {
    if shock {
        let ring = gauss_ring(dist, radius, width) * fade;
        let fill = if dist <= radius {
            (1.0 - dist / radius.max(0.04)).clamp(0.0, 1.0).powf(0.65) * fade * 0.38
        } else {
            0.0
        };
        return ((ring + fill) * power).clamp(0.0, 1.0);
    }
    let ring = gauss_ring(dist, radius, width) * fade;
    let center = gauss_ring(dist, 0.0, (width * 3.5).max(0.08)) * (-age / 0.14).exp() * 0.65;
    let amt = match kind {
        RippleKind::Ring => ring + center,
        RippleKind::Wave => {
            let wash = if dist <= radius {
                fade * 0.24 * (1.0 - dist / radius.max(0.04))
            } else {
                0.0
            };
            gauss_ring(dist, radius, width * 1.55) * fade + wash + center * 0.3
        }
        RippleKind::Pulse => {
            let sigma = (0.08 + age * (0.38 + width * 0.35)).max(0.06);
            (-dist * dist / (2.0 * sigma * sigma)).exp() * fade
        }
        RippleKind::Double => ring + gauss_ring(dist, radius * 0.5, width) * fade * 0.85 + center,
        RippleKind::Fill => {
            if dist <= radius {
                fade * 0.42 + gauss_ring(dist, radius, width) * fade
            } else {
                0.0
            }
        }
        RippleKind::Echo => {
            let gap = 0.18;
            let e1 = if radius > gap {
                gauss_ring(dist, radius - gap, width) * fade * 0.75
            } else {
                0.0
            };
            let e2 = if radius > gap * 2.0 {
                gauss_ring(dist, radius - gap * 2.0, width) * fade * 0.5
            } else {
                0.0
            };
            ring + e1 + e2 + center
        }
    };
    (amt * power).clamp(0.0, 1.0)
}

fn strip_on(dist: f32, radius: f32, half: f32) -> bool {
    radius >= 0.0 && (dist - radius).abs() <= half
}

fn strip_front(dist: f32, radius: f32, half: f32, fade: f32) -> f32 {
    if strip_on(dist, radius, half) {
        fade
    } else {
        0.0
    }
}

fn strip_shape(
    kind: RippleKind,
    dist: f32,
    radius: f32,
    half: f32,
    fade: f32,
    shock: bool,
    age: f32,
    power: f32,
    strip_speed: f32,
) -> f32 {
    if shock {
        let mut amt = strip_front(dist, radius, half, fade) * power;
        if radius > 0.0 && dist < radius && dist >= (radius - 2.0).max(0.0) {
            amt = amt.max((1.0 - (radius - dist) / 2.0).clamp(0.0, 1.0) * fade * 0.32 * power);
        }
        return amt.clamp(0.0, 1.0);
    }
    let flash = if dist < 0.5 && age < 0.08 {
        (1.0 - age / 0.08) * 0.9
    } else {
        0.0
    };
    let amt = match kind {
        RippleKind::Ring => strip_front(dist, radius, half, fade).max(flash),
        RippleKind::Wave => {
            let band = strip_front(dist, radius, (half + 1.0).min(2.0), fade);
            let trail = if dist < radius && dist >= (radius - 2.0).max(0.0) {
                (1.0 - (radius - dist) / 2.0).clamp(0.0, 1.0) * fade * 0.4
            } else {
                0.0
            };
            band.max(trail).max(flash * 0.6)
        }
        RippleKind::Pulse => {
            let sigma = (0.8 + age * strip_speed * 0.5).max(0.7);
            (-dist * dist / (2.0 * sigma * sigma)).exp() * fade
        }
        RippleKind::Double => {
            let inner = (radius * 0.5).floor();
            strip_front(dist, radius, half, fade)
                .max(if radius >= 2.0 {
                    strip_front(dist, inner, half, fade * 0.85)
                } else {
                    0.0
                })
                .max(flash)
        }
        RippleKind::Fill => {
            if dist <= radius {
                let edge = if (dist - radius).abs() <= half { 1.0 } else { 0.42 };
                fade * edge
            } else {
                flash
            }
        }
        RippleKind::Echo => {
            let mut amt = flash;
            for (gap, scale) in [(0.0, 1.0), (3.0, 0.75), (6.0, 0.5)] {
                let r = radius - gap;
                if r >= 0.0 {
                    amt = amt.max(strip_front(dist, r, half, fade * scale));
                }
            }
            amt
        }
    };
    (amt * power).clamp(0.0, 1.0)
}

fn ripple_sample_zones(ripples: &[ColorRipple], t: f32, n: usize, tune: &RippleTune) -> (f32, f32) {
    let mut energy = 0.0f32;
    let mut hue = 0.0f32;
    let mut weight = 0.0f32;
    for ripple in ripples {
        let (speed, width, lifetime, power) = if ripple.shock {
            (
                tune.travel() * 1.65,
                tune.ring_width(n) * (1.85 + tune.shock_strength * 0.55),
                ripple_lifetime(true, n, tune),
                (0.9 + tune.shock_strength * 0.55).clamp(0.7, 2.0),
            )
        } else {
            (tune.travel(), tune.ring_width(n), ripple_lifetime(false, n, tune), 1.0)
        };
        let radius = ripple.age * speed;
        let fade = (1.0 - ripple.age / lifetime)
            .clamp(0.0, 1.0)
            .powf(if ripple.shock { 0.85 } else { 1.2 });
        let dist = (t - ripple.origin).abs();
        let amt = zone_shape(tune.kind, dist, radius, width, fade, ripple.shock, ripple.age, power);
        if amt > 0.02 {
            let twist = if ripple.shock { tune.twist * 1.35 } else { tune.twist };
            energy = energy.max(amt);
            hue += (ripple.hue + (radius * 160.0 + ripple.age * 90.0) * twist) * amt;
            weight += amt;
        }
    }
    let hue = if weight > 1e-3 { hue / weight } else { 0.0 };
    (energy.clamp(0.0, 1.0), hue.rem_euclid(360.0))
}

fn ripple_sample_strips(ripples: &[ColorRipple], i: usize, n: usize, tune: &RippleTune) -> (f32, f32) {
    let mut energy = 0.0f32;
    let mut hue = 0.0f32;
    let mut weight = 0.0f32;
    let max_i = (n as i32 - 1).max(1);
    let i = i as i32;
    for ripple in ripples {
        let origin = origin_strip(ripple.origin, n);
        let shock = ripple.shock;
        let travel = if shock { tune.travel() * 1.65 } else { tune.travel() };
        let strip_speed = (travel * max_i as f32).max(4.0);
        let lifetime = ripple_lifetime(shock, n, tune);
        let power = if shock {
            (0.9 + tune.shock_strength * 0.55).clamp(0.7, 2.0)
        } else {
            1.0
        };
        let radius = (ripple.age * strip_speed).floor();
        let fade = (1.0 - ripple.age / lifetime)
            .clamp(0.0, 1.0)
            .powf(if shock { 0.85 } else { 1.15 });
        let dist = (i - origin).abs() as f32;
        let thick = if shock {
            (1.0 + (tune.width - 0.1) * 2.4 + (tune.shock_strength - 1.0).max(0.0) * 0.8).floor().clamp(1.0, 4.0)
        } else {
            (1.0 + (tune.width - 0.35).max(0.0) * 3.2).floor().clamp(1.0, 3.0)
        };
        let half = ((thick - 1.0) * 0.5).floor();
        let amt = strip_shape(tune.kind, dist, radius, half, fade, shock, ripple.age, power, strip_speed);
        if amt > 0.02 {
            let twist = if shock { tune.twist * 1.35 } else { tune.twist };
            energy = energy.max(amt);
            hue += (ripple.hue + (radius * 14.0 + ripple.age * 90.0) * twist) * amt;
            weight += amt;
        }
    }
    let hue = if weight > 1e-3 { hue / weight } else { 0.0 };
    (energy.clamp(0.0, 1.0), hue.rem_euclid(360.0))
}

fn ripple_energy(ripples: &[ColorRipple], i: usize, t: f32, n: usize, tune: &RippleTune) -> f32 {
    ripple_sample(ripples, i, t, n, tune).0
}

fn ripple_hue(ripples: &[ColorRipple], i: usize, t: f32, n: usize, tune: &RippleTune) -> f32 {
    ripple_sample(ripples, i, t, n, tune).1
}

fn cascade_bloom(e: &mut [f32]) {
    for i in 1..e.len() {
        e[i] = e[i].max(e[i - 1] * 0.62).min(1.0);
    }
}

fn cascade_fire(e: &mut [f32]) {
    for i in 1..e.len() {
        e[i] = (e[i] * 0.85 + e[i - 1] * 0.55).min(1.0);
    }
}

fn fire_color(t: f32) -> [u8; 3] {
    let t = t.clamp(0.0, 1.0);
    if t < 0.5 {
        lamps::mix_rgb([160, 6, 0], [255, 88, 6], t * 2.0)
    } else {
        lamps::mix_rgb([255, 88, 6], [255, 230, 70], (t - 0.5) * 2.0)
    }
}

fn spatial_blur(levels: &[f32; 4], spread: f32) -> [f32; 4] {
    let mix = spread.clamp(0.0, 1.0);
    let blur = [
        levels[0] * 0.78 + levels[1] * 0.22,
        levels[1] * 0.62 + levels[0] * 0.19 + levels[2] * 0.19,
        levels[2] * 0.62 + levels[1] * 0.19 + levels[3] * 0.19,
        levels[3] * 0.78 + levels[2] * 0.22,
    ];
    [
        levels[0] * (1.0 - mix) + blur[0] * mix,
        levels[1] * (1.0 - mix) + blur[1] * mix,
        levels[2] * (1.0 - mix) + blur[2] * mix,
        levels[3] * (1.0 - mix) + blur[3] * mix,
    ]
}

fn spatial_blur_n(levels: &mut [f32], spread: f32) {
    let mix = spread.clamp(0.0, 1.0);
    let n = levels.len();
    if mix < 0.001 || n < 2 {
        return;
    }
    let orig = levels.to_vec();
    for i in 0..n {
        let left = orig[i.saturating_sub(1)];
        let center = orig[i];
        let right = orig[(i + 1).min(n - 1)];
        let blur = if i == 0 {
            center * 0.78 + right * 0.22
        } else if i + 1 == n {
            center * 0.78 + left * 0.22
        } else {
            center * 0.62 + left * 0.19 + right * 0.19
        };
        levels[i] = center * (1.0 - mix) + blur * mix;
    }
}

fn compute_fft(samples: &[f32]) -> Option<([f32; FFT_SIZE], [f32; FFT_SIZE])> {
    if samples.len() < FFT_SIZE {
        return None;
    }
    let mut re = [0.0f32; FFT_SIZE];
    let mut im = [0.0f32; FFT_SIZE];
    let n = FFT_SIZE as f32;
    for i in 0..FFT_SIZE {
        let t = 2.0 * PI * i as f32 / n;
        let window = 0.35875 - 0.48829 * t.cos() + 0.14128 * (2.0 * t).cos() - 0.01168 * (3.0 * t).cos();
        re[i] = samples[i] * window;
    }
    fft(&mut re, &mut im);
    Some((re, im))
}

fn triangle_band(re: &[f32], im: &[f32], sample_rate: u32, lo: f32, mid: f32, hi: f32) -> f32 {
    let bin_hz = sample_rate as f32 / FFT_SIZE as f32;
    let mut acc = 0.0;
    let mut weight = 0.0;
    let start = (lo / bin_hz).floor().max(1.0) as usize;
    let end = (hi / bin_hz).ceil().min((FFT_SIZE / 2 - 1) as f32) as usize;
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
        let mag = (re[i] * re[i] + im[i] * im[i]).sqrt();
        let pink = (freq / 800.0).sqrt().clamp(0.35, 2.6);
        acc += mag * tri * pink;
        weight += tri;
    }
    let mean = if weight > 0.0 { acc / weight } else { 0.0 };
    (mean / 48.0).ln_1p()
}

fn kick_energy(re: &[f32], im: &[f32], sample_rate: u32) -> f32 {
    triangle_band(re, im, sample_rate, 30.0, 55.0, 105.0)
}

fn analyze_bands(samples: &[f32], sample_rate: u32) -> ([f32; 4], f32) {
    let Some((re, im)) = compute_fft(samples) else {
        return ([0.0; 4], 0.0);
    };
    let filters = [
        (25.0, 70.0, 200.0),
        (90.0, 320.0, 800.0),
        (400.0, 1400.0, 3500.0),
        (1800.0, 5000.0, 12000.0),
    ];
    let mut bands = [0.0f32; 4];
    for (band, (lo, mid, hi)) in filters.iter().enumerate() {
        bands[band] = triangle_band(&re, &im, sample_rate, *lo, *mid, *hi);
    }
    (bands, kick_energy(&re, &im, sample_rate))
}

fn analyze_spectrum(samples: &[f32], sample_rate: u32, bands: usize) -> (Vec<f32>, f32) {
    let bands = bands.max(1);
    let Some((re, im)) = compute_fft(samples) else {
        return (vec![0.0; bands], 0.0);
    };
    let log_lo = 25.0f32.ln();
    let log_hi = 12_000.0f32.ln();
    let points = bands + 2;
    let edge = |k: usize| {
        let t = k as f32 / (points - 1) as f32;
        (log_lo + (log_hi - log_lo) * t).exp()
    };
    let spec = (0..bands)
        .map(|b| triangle_band(&re, &im, sample_rate, edge(b), edge(b + 1), edge(b + 2)))
        .collect();
    (spec, kick_energy(&re, &im, sample_rate))
}

fn fft(re: &mut [f32], im: &mut [f32]) {
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

    let mut len = 2;
    while len <= n {
        let ang = -2.0 * PI / len as f32;
        let (wlen_im, wlen_re) = ang.sin_cos();
        let half = len / 2;
        for i in (0..n).step_by(len) {
            let mut wr = 1.0;
            let mut wi = 0.0;
            for k in 0..half {
                let u_re = re[i + k];
                let u_im = im[i + k];
                let v_re = re[i + k + half] * wr - im[i + k + half] * wi;
                let v_im = re[i + k + half] * wi + im[i + k + half] * wr;
                re[i + k] = u_re + v_re;
                im[i + k] = u_im + v_im;
                re[i + k + half] = u_re - v_re;
                im[i + k + half] = u_im - v_im;
                let next_wr = wr * wlen_re - wi * wlen_im;
                wi = wr * wlen_im + wi * wlen_re;
                wr = next_wr;
            }
        }
        len <<= 1;
    }
}

#[cfg(target_os = "windows")]
fn capture_loopback(
    samples: Arc<Mutex<Vec<f32>>>,
    stop: Arc<std::sync::atomic::AtomicBool>,
    capture_ok: Arc<std::sync::atomic::AtomicBool>,
) -> Result<(), String> {
    use std::{
        collections::{HashMap, HashSet},
        sync::atomic::AtomicUsize,
    };

    wasapi::initialize_mta()
        .ok()
        .map_err(|_| "COM init failed".to_string())?;

    let per_device: Arc<Mutex<HashMap<String, Vec<f32>>>> = Arc::new(Mutex::new(HashMap::new()));
    let wanted: Arc<Mutex<HashSet<String>>> = Arc::new(Mutex::new(HashSet::new()));
    let preferred: Arc<Mutex<HashSet<String>>> = Arc::new(Mutex::new(HashSet::new()));
    let live = Arc::new(AtomicUsize::new(0));
    let mut workers: HashMap<String, thread::JoinHandle<()>> = HashMap::new();
    let mut last_ids = HashSet::new();
    let mut last_pref = HashSet::new();

    while !stop.load(Ordering::SeqCst) {
        let targets = list_render_targets();
        let ids: HashSet<String> = targets.iter().map(|(id, _, _)| id.clone()).collect();
        let pref: HashSet<String> = targets
            .iter()
            .filter(|(_, _, is_pref)| *is_pref)
            .map(|(id, _, _)| id.clone())
            .collect();
        if let Ok(mut guard) = wanted.lock() {
            *guard = ids.clone();
        }
        if let Ok(mut guard) = preferred.lock() {
            *guard = pref.clone();
        }

        if ids != last_ids || pref != last_pref {
            let names: Vec<String> = targets
                .iter()
                .map(|(_, name, is_pref)| {
                    if *is_pref {
                        format!("{name}*")
                    } else {
                        name.clone()
                    }
                })
                .collect();
            legion_rgb_driver::debug_log(&format!(
                "AUDIO: follow outputs [{}]",
                names.join(" | ")
            ));
            last_ids = ids.clone();
            last_pref = pref;
        }

        for (id, name, _) in targets {
            if workers.contains_key(&id) {
                continue;
            }
            let samples = samples.clone();
            let per_device = per_device.clone();
            let wanted = wanted.clone();
            let preferred = preferred.clone();
            let stop = stop.clone();
            let live = live.clone();
            let capture_ok = capture_ok.clone();
            workers.insert(
                id.clone(),
                thread::spawn(move || {
                    if let Err(err) = capture_one_output(
                        &id,
                        &name,
                        samples,
                        per_device,
                        wanted,
                        preferred,
                        stop,
                        live,
                        capture_ok,
                    ) {
                        legion_rgb_driver::debug_log(&format!("AUDIO: {name} loopback ended: {err}"));
                    }
                }),
            );
        }

        let finished: Vec<String> = workers
            .iter()
            .filter(|(_, handle)| handle.is_finished())
            .map(|(id, _)| id.clone())
            .collect();
        for id in finished {
            if let Some(handle) = workers.remove(&id) {
                let _ = handle.join();
            }
        }
        capture_ok.store(live.load(Ordering::Relaxed) > 0, Ordering::Relaxed);
        thread::sleep(Duration::from_millis(350));
    }

    if let Ok(mut guard) = wanted.lock() {
        guard.clear();
    }
    for (_, handle) in workers {
        let _ = handle.join();
    }
    capture_ok.store(false, Ordering::Relaxed);
    Ok(())
}

#[cfg(target_os = "windows")]
fn list_render_targets() -> Vec<(String, String, bool)> {
    use std::collections::{BTreeMap, HashSet};
    use wasapi::{get_default_device_for_role, DeviceCollection, Direction, Role};

    let mut found: BTreeMap<String, String> = BTreeMap::new();
    if let Ok(collection) = DeviceCollection::new(&Direction::Render) {
        if let Ok(count) = collection.get_nbr_devices() {
            for idx in 0..count {
                if let Ok(device) = collection.get_device_at_index(idx) {
                    if let (Ok(id), Ok(name)) = (device.get_id(), device.get_friendlyname()) {
                        found.insert(id, name);
                    }
                }
            }
        }
    }
    let mut preferred = HashSet::new();
    for role in [Role::Console, Role::Multimedia, Role::Communications] {
        if let Ok(device) = get_default_device_for_role(&Direction::Render, &role) {
            if let (Ok(id), Ok(name)) = (device.get_id(), device.get_friendlyname()) {
                preferred.insert(id.clone());
                found.entry(id).or_insert(name);
            }
        }
    }
    found
        .into_iter()
        .map(|(id, name)| {
            let is_pref = preferred.contains(&id);
            (id, name, is_pref)
        })
        .collect()
}

#[cfg(target_os = "windows")]
fn publish_mix(
    per_device: &Mutex<std::collections::HashMap<String, Vec<f32>>>,
    samples: &Mutex<Vec<f32>>,
    preferred: &Mutex<std::collections::HashSet<String>>,
) {
    let mix = {
        let Ok(map) = per_device.lock() else {
            return;
        };
        if map.is_empty() {
            vec![0.0f32; FFT_SIZE]
        } else {
            let pref = preferred.lock().ok();
            let is_pref = |id: &str| pref.as_ref().map(|p| p.contains(id)).unwrap_or(false);
            let mut best_pref: Option<(f32, &Vec<f32>)> = None;
            let mut best_any: Option<(f32, &Vec<f32>)> = None;
            for (id, buf) in map.iter() {
                let rms = recent_rms(buf);
                if best_any.as_ref().map(|(e, _)| rms > *e).unwrap_or(true) {
                    best_any = Some((rms, buf));
                }
                if is_pref(id) && best_pref.as_ref().map(|(e, _)| rms > *e).unwrap_or(true) {
                    best_pref = Some((rms, buf));
                }
            }
            let chosen = match (best_pref, best_any) {
                (Some((pref_rms, pref_buf)), Some((any_rms, any_buf)))
                    if any_rms > pref_rms * 2.8 && any_rms > 0.012 =>
                {
                    any_buf
                }
                (Some((_, pref_buf)), _) => pref_buf,
                (_, Some((_, any_buf))) => any_buf,
                _ => return,
            };
            chosen.clone()
        }
    };
    if let Ok(mut guard) = samples.lock() {
        *guard = mix;
    }
}

#[cfg(target_os = "windows")]
fn capture_one_output(
    id: &str,
    name: &str,
    samples: Arc<Mutex<Vec<f32>>>,
    per_device: Arc<Mutex<std::collections::HashMap<String, Vec<f32>>>>,
    wanted: Arc<Mutex<std::collections::HashSet<String>>>,
    preferred: Arc<Mutex<std::collections::HashSet<String>>>,
    stop: Arc<std::sync::atomic::AtomicBool>,
    live: Arc<std::sync::atomic::AtomicUsize>,
    capture_ok: Arc<std::sync::atomic::AtomicBool>,
) -> Result<(), String> {
    use std::collections::VecDeque;
    use wasapi::{DeviceCollection, Direction, SampleType, StreamMode, WaveFormat};

    wasapi::initialize_mta()
        .ok()
        .map_err(|_| "COM init failed".to_string())?;

    let collection = DeviceCollection::new(&Direction::Render).map_err(|e| e.to_string())?;
    let count = collection.get_nbr_devices().map_err(|e| e.to_string())?;
    let mut device = None;
    for idx in 0..count {
        if let Ok(candidate) = collection.get_device_at_index(idx) {
            if candidate.get_id().ok().as_deref() == Some(id) {
                device = Some(candidate);
                break;
            }
        }
    }
    let device = device.ok_or_else(|| format!("output gone: {name}"))?;
    let mut audio_client = device.get_iaudioclient().map_err(|e| e.to_string())?;
    let format = WaveFormat::new(32, 32, &SampleType::Float, SAMPLE_RATE as usize, 2, None);
    let mode = StreamMode::EventsShared {
        autoconvert: true,
        buffer_duration_hns: 100_000,
    };
    audio_client
        .initialize_client(&format, &Direction::Capture, &mode)
        .map_err(|e| e.to_string())?;
    let event = audio_client.set_get_eventhandle().map_err(|e| e.to_string())?;
    let capture_client = audio_client.get_audiocaptureclient().map_err(|e| e.to_string())?;
    audio_client.start_stream().map_err(|e| e.to_string())?;
    live.fetch_add(1, Ordering::Relaxed);
    capture_ok.store(true, Ordering::Relaxed);
    legion_rgb_driver::debug_log(&format!("AUDIO: WASAPI loopback started on {name}"));

    let mut queue: VecDeque<u8> = VecDeque::new();
    let mut ring = vec![0.0f32; FFT_SIZE];
    let mut write_at = 0usize;
    let channels = 2usize;
    let bytes_per_sample = 4usize;
    let mut stale_reads = 0u8;

    while !stop.load(Ordering::SeqCst) {
        if let Ok(guard) = wanted.lock() {
            if !guard.contains(id) {
                break;
            }
        }
        if event.wait_for_event(20).is_err() {
            continue;
        }
        if capture_client.read_from_device_to_deque(&mut queue).is_err() {
            stale_reads = stale_reads.saturating_add(1);
            if stale_reads >= 8 {
                break;
            }
            continue;
        }
        while capture_client
            .get_next_packet_size()
            .ok()
            .flatten()
            .unwrap_or(0)
            > 0
        {
            if capture_client.read_from_device_to_deque(&mut queue).is_err() {
                break;
            }
        }
        stale_reads = 0;
        let frame_bytes = bytes_per_sample * channels;
        while queue.len() >= frame_bytes {
            let mut left = [0u8; 4];
            let mut right = [0u8; 4];
            for b in &mut left {
                *b = queue.pop_front().unwrap();
            }
            for b in &mut right {
                *b = queue.pop_front().unwrap();
            }
            let l = f32::from_le_bytes(left);
            let r = f32::from_le_bytes(right);
            ring[write_at] = (l + r) * 0.5;
            write_at = (write_at + 1) % FFT_SIZE;
        }
        let window = {
            let mut window = vec![0.0f32; FFT_SIZE];
            let (tail, head) = ring.split_at(write_at);
            let split = FFT_SIZE - write_at;
            window[split..].copy_from_slice(tail);
            window[..split].copy_from_slice(head);
            window
        };
        if let Ok(mut map) = per_device.lock() {
            map.insert(id.to_string(), window);
        }
        publish_mix(&per_device, &samples, &preferred);
    }

    let _ = audio_client.stop_stream();
    if let Ok(mut map) = per_device.lock() {
        map.remove(id);
    }
    publish_mix(&per_device, &samples, &preferred);
    live.fetch_sub(1, Ordering::Relaxed);
    capture_ok.store(live.load(Ordering::Relaxed) > 0, Ordering::Relaxed);
    Ok(())
}
