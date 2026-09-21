use std::{
    f32::consts::PI,
    sync::{
        atomic::{AtomicU64, Ordering},
        Arc, Mutex,
    },
    thread,
    time::{Duration, Instant},
};

use crate::{
    enums::{AudioAnalysis, AudioColorMode, AudioStyle, Effects, RippleKind, RippleOrigin, RippleTint, RippleTrigger},
    manager::{
        effects::{audio_auto, audio_beats, audio_color, audio_dsp, audio_flow, audio_hpss, audio_mel, audio_onset, audio_tempo, lamps},
        profile::Profile,
        Inner,
    },
};

const FFT_SIZE: usize = 1024;
const RING: usize = 8192;
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
    pub analysis: AudioAnalysis,
    pub auto_resolved: AudioAnalysis,
    pub bpm: u16,
    pub bpm_locked: bool,
    pub drop: bool,
    pub custom_rgb: [u8; 12],
}

#[derive(Clone, Copy, Debug)]
pub struct AudioHud {
    pub resolved: AudioAnalysis,
    pub bpm: u16,
    pub locked: bool,
    pub drop: bool,
}

impl Default for AudioHud {
    fn default() -> Self {
        Self {
            resolved: AudioAnalysis::Accurate,
            bpm: 0,
            locked: false,
            drop: false,
        }
    }
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
                analysis,
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
                analysis,
                auto_resolved: AudioAnalysis::Accurate,
                bpm: 0,
                bpm_locked: false,
                drop: false,
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

    pub fn hud(&self) -> AudioHud {
        AudioHud {
            resolved: if matches!(self.auto_resolved, AudioAnalysis::Auto) {
                AudioAnalysis::Accurate
            } else {
                self.auto_resolved
            },
            bpm: if self.bpm_locked { self.bpm } else { 0 },
            locked: self.bpm_locked && self.bpm > 0,
            drop: self.drop,
        }
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
            analysis: self.analysis,
            auto_resolved: if matches!(self.auto_resolved, AudioAnalysis::Classic | AudioAnalysis::Auto) {
                AudioAnalysis::Accurate
            } else {
                self.auto_resolved
            },
            bpm: self.bpm,
            bpm_locked: self.bpm_locked,
            drop: self.drop,
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
            analysis: AudioAnalysis::Auto,
            auto_resolved: AudioAnalysis::Accurate,
            bpm: 0,
            bpm_locked: false,
            drop: false,
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

    let samples = Arc::new(Mutex::new(vec![0.0f32; RING]));
    let stereo = Arc::new(Mutex::new([vec![0.0f32; RING], vec![0.0f32; RING]]));
    let capture_seq = Arc::new(AtomicU64::new(0));
    let capture_ok = Arc::new(std::sync::atomic::AtomicBool::new(false));
    let stop = manager.stop_signals.manager_stop_signal.clone();
    let capture_stop = stop.clone();
    let capture_samples = samples.clone();
    let capture_stereo = stereo.clone();
    let capture_flag = capture_ok.clone();
    let capture_seq_thread = capture_seq.clone();

    let capture = thread::spawn(move || {
        #[cfg(target_os = "windows")]
        {
            if let Err(err) = capture_loopback(capture_samples, capture_stereo, capture_seq_thread, capture_stop, capture_flag) {
                legion_rgb_driver::debug_log(&format!("AUDIO: loopback failed: {err}"));
            }
        }
        #[cfg(not(target_os = "windows"))]
        {
            let _ = (capture_samples, capture_stereo, capture_seq_thread, capture_stop, capture_flag);
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
    let mut scope_hold: Vec<f32> = Vec::new();
    let mut scope_peak = 0.002f32;
    let mut spec_hist_e: Vec<f32> = Vec::new();
    let mut spec_hist_h: Vec<f32> = Vec::new();
    let mut spec_hist_bands = [[0.0f32; 4]; 4];
    let mut stereo_l: Vec<f32> = Vec::new();
    let mut stereo_r: Vec<f32> = Vec::new();
    let mut pitch_hue = 210.0f32;
    let mut pitch_hz = 0.0f32;
    let mut key_hue = 200.0f32;
    let mut liss_width = 0.12f32;
    let mut bubbles: Vec<AudioBubble> = Vec::new();
    let mut color_ripples: Vec<ColorRipple> = Vec::new();
    let mut column = ColumnFx::new();
    let mut prev_gate = [0.0f32; 4];
    let mut beat_env = [0.12f32; 4];
    let mut fast_env = 0.0f32;
    let mut fast_peak = 0.002f32;
    let mut prev_loud = 0.0f32;
    let mut drop_loud = 0.0f32;
    let mut drop_trough = 1.0f32;
    let mut shock_cool = 0.0f32;
    let mut drop_flash_age = 1.0f32;
    let mut beats = BeatTracker::new();
    let mut flux_beats = audio_beats::SuperFluxTracker::new();
    let mut analyzer = audio_dsp::Analyzer::new(SAMPLE_RATE);
    let mut studio = audio_dsp::Analyzer::studio(SAMPLE_RATE);
    let mut mel_bank = audio_mel::MelBank::new();
    let mut hpss = audio_hpss::Hpss::new();
    let mut complex_on = audio_onset::ComplexOnset::new();
    let mut davies = audio_tempo::DaviesTempo::new(SAMPLE_RATE, audio_dsp::HOP);
    let mut auto_pick = audio_auto::AutoPicker::new();
    let mut last_tick = Instant::now();
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
        let (mut frame, mut kick_raw, fast_rms, wave_snap, left_snap, right_snap, ring_snap, seq) = {
            let guard = samples.lock().unwrap();
            let seq = capture_seq.load(Ordering::Relaxed);
            let rms = recent_rms(&guard);
            let classic_win = last_slice(&guard, FFT_SIZE);
            let (left, right) = {
                let lr = stereo.lock().unwrap();
                (last_slice(&lr[0], FFT_SIZE), last_slice(&lr[1], FFT_SIZE))
            };
            let (bands, kick) = analyze_bands(&classic_win, SAMPLE_RATE);
            (bands, kick, rms, classic_win, left, right, guard.clone(), seq)
        };

        let classic_bands = frame;
        let classic_kick = kick_raw;
        let is_auto = matches!(p.analysis, AudioAnalysis::Auto);
        let classic_forced = matches!(p.analysis, AudioAnalysis::Classic);
        let mut dsp = audio_dsp::AnalysisFrame::default();
        let mut mel_eq = [0.0f32; audio_dsp::EQ_BANDS];
        let mut mel_b4 = [0.0f32; 4];
        let mut hpss_fk = (0.0f32, 0.0f32);
        let mut complex_fk = (0.0f32, 0.0f32);
        let mut studio_frame = audio_dsp::AnalysisFrame::default();
        let mut ran_studio = false;

        if !classic_forced {
            if is_auto {
                dsp = analyzer.observe_seq(&ring_snap, seq);
                let (eq, b4) = mel_bank.tick(analyzer.mag(), analyzer.sample_rate(), analyzer.fft_size());
                mel_eq = eq;
                mel_b4 = b4;
                hpss_fk = hpss.tick(analyzer.mag(), analyzer.sample_rate(), analyzer.fft_size());
                complex_fk = complex_on.tick(
                    analyzer.re(),
                    analyzer.im(),
                    analyzer.sample_rate(),
                    analyzer.fft_size(),
                );
                for &f in analyzer.hop_fluxes() {
                    davies.tick(f);
                }
                let (bass_share, _) = eq24_stats(&dsp.eq24);
                let bass_need = dsp.centroid < 0.18 && bass_share > 0.45;
                if matches!(auto_pick.current(), AudioAnalysis::Studio) || bass_need {
                    studio_frame = studio.observe_seq(&ring_snap, seq);
                    ran_studio = true;
                }
            } else if matches!(p.analysis, AudioAnalysis::Studio) {
                dsp = studio.observe_seq(&ring_snap, seq);
                studio_frame = dsp;
                ran_studio = true;
            } else {
                dsp = analyzer.observe_seq(&ring_snap, seq);
                match p.analysis {
                    AudioAnalysis::Mel => {
                        let (eq, b4) = mel_bank.tick(analyzer.mag(), analyzer.sample_rate(), analyzer.fft_size());
                        dsp.eq24 = eq;
                        dsp.bands4 = b4;
                    }
                    AudioAnalysis::Hpss => {
                        let (p_flux, k_flux) = hpss.tick(analyzer.mag(), analyzer.sample_rate(), analyzer.fft_size());
                        dsp.flux = p_flux;
                        dsp.kick_flux = k_flux;
                    }
                    AudioAnalysis::Complex => {
                        let (flux, k_flux) = complex_on.tick(
                            analyzer.re(),
                            analyzer.im(),
                            analyzer.sample_rate(),
                            analyzer.fft_size(),
                        );
                        dsp.flux = flux;
                        dsp.kick_flux = k_flux;
                    }
                    AudioAnalysis::Tempo => {
                        for &f in analyzer.hop_fluxes() {
                            davies.tick(f);
                        }
                    }
                    _ => {}
                }
            }
            if !is_auto {
                frame = dsp.bands4;
                kick_raw = dsp.kick;
            }
        }

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
        let mut flux_fast = (loud - prev_loud).max(0.0);
        prev_loud = loud;

        let gated_scout = volume_dead || fast_env < noise_floor * 1.6;
        let resolved = if is_auto {
            let (bass_share, band_spread) = eq24_stats(&dsp.eq24);
            let kick_share = dsp.kick_flux / (dsp.flux + dsp.kick_flux + 1e-4);
            auto_pick.tick(
                dt,
                &audio_auto::AutoFeatures {
                    rms: fast_env,
                    flux: dsp.flux,
                    kick_share,
                    flatness: dsp.flatness,
                    centroid: dsp.centroid,
                    bass_share,
                    band_spread,
                    perc_share: hpss.perc_share(),
                    complex_kick: complex_fk.1,
                    superflux_kick: dsp.kick_flux,
                    tempo_locked: davies.locked(),
                    tempo_strength: davies.strength(),
                    lamp_n: n,
                    style: p.style,
                },
                gated_scout,
            )
        } else if classic_forced {
            AudioAnalysis::Classic
        } else {
            p.analysis
        };
        if is_auto {
            match resolved {
                AudioAnalysis::Classic if n == 4 => {
                    frame = classic_bands;
                    kick_raw = classic_kick;
                }
                AudioAnalysis::Mel => {
                    dsp.eq24 = mel_eq;
                    dsp.bands4 = mel_b4;
                    frame = dsp.bands4;
                    kick_raw = dsp.kick;
                }
                AudioAnalysis::Studio => {
                    if ran_studio {
                        dsp = studio_frame;
                    }
                    frame = dsp.bands4;
                    kick_raw = dsp.kick;
                }
                AudioAnalysis::Hpss => {
                    dsp.flux = hpss_fk.0;
                    dsp.kick_flux = hpss_fk.1;
                    frame = dsp.bands4;
                    kick_raw = dsp.kick;
                }
                AudioAnalysis::Complex => {
                    dsp.flux = complex_fk.0;
                    dsp.kick_flux = complex_fk.1;
                    frame = dsp.bands4;
                    kick_raw = dsp.kick;
                }
                _ => {
                    frame = dsp.bands4;
                    kick_raw = dsp.kick;
                }
            }
        }
        if let Ok(mut live) = params.lock() {
            live.auto_resolved = resolved;
        }
        let engine = if is_auto { resolved } else { p.analysis };
        let classic = classic_forced || (is_auto && matches!(engine, AudioAnalysis::Classic) && n == 4);
        if !classic {
            flux_fast = dsp.flux * (0.45 + p.punch * 0.25) + flux_fast * 0.35;
        }
        let beats_bias = matches!(
            engine,
            AudioAnalysis::Beats | AudioAnalysis::Complex | AudioAnalysis::Hpss
        );
        let spectrum_feel = matches!(
            engine,
            AudioAnalysis::Spectrum | AudioAnalysis::Mel | AudioAnalysis::Studio
        );

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
        let long_tau = if spectrum_feel { 0.85 } else { 0.55 };
        let attack_tau = if beats_bias {
            0.006 + p.smoothness * 0.008
        } else {
            0.008 + p.smoothness * 0.012
        };
        let release_tau = if spectrum_feel {
            0.06 + p.smoothness * 0.18
        } else {
            0.035 + p.smoothness * 0.14
        };

        if beat_gates {
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
                let flux = if spectrum_feel {
                    0.0
                } else {
                    ((short_avg[i] - prev_short[i]).max(0.0) / long_avg[i]) + flux_fast
                };
                prev_short[i] = short_avg[i];
                let punch_amt = if spectrum_feel { 0.0 } else { flux * p.punch * 0.9 };
                let target = if gated {
                    0.0
                } else {
                    (loud * (0.28 + 0.72 * band) + punch_amt).clamp(0.0, 1.0)
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
        let bass_now = levels[0];
        let band_peak = levels.iter().copied().fold(0.0f32, f32::max);
        let rise = (loud - drop_trough).max(0.0);
        let kick_now = (kick_raw * p.bass).clamp(0.0, 2.5);
        let hits = if !volume_dead {
            if classic {
                beats.tick(dt, gated, flux_fast, bass_now, kick_now, band_peak, p.punch, p.squelch)
            } else {
                flux_beats.tick(
                    dt,
                    gated,
                    dsp.flux,
                    dsp.kick_flux,
                    bass_now,
                    kick_now,
                    band_peak,
                    p.punch,
                    p.squelch,
                    beats_bias,
                )
            }
        } else {
            beats.reset();
            flux_beats.reset();
            davies.reset();
            audio_beats::BeatHits::none()
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
        let ring_hit = p.ripple_color
            && match p.ripple_trigger {
                RippleTrigger::All => hits.any,
                RippleTrigger::Bass => hits.bass,
                RippleTrigger::Kick => hits.kick,
                RippleTrigger::TripleKick => hits.triple,
            };
        drop_flash_age = if hits.kick || hits.triple || drop_hit {
            0.0
        } else {
            (drop_flash_age + dt).min(4.0)
        };
        let (ibi, locked) = if classic {
            (beats.ibi(), beats.tempo_locked())
        } else if matches!(engine, AudioAnalysis::Tempo) || davies.locked() {
            (davies.ibi(), davies.locked())
        } else {
            (flux_beats.ibi(), flux_beats.tempo_locked())
        };
        if let Ok(mut live) = params.lock() {
            live.bpm = audio_flow::bpm_from_ibi(ibi, locked);
            live.bpm_locked = locked;
            live.drop = !volume_dead && drop_flash_age < 0.25;
        }
        ramp_phase = (ramp_phase + dt * (0.12 + p.motion * 0.55)) % (2.0 * std::f32::consts::PI);
        if tint_state.len() != n {
            tint_state.resize(n, [0.0; 3]);
        }

        if matches!(p.style, AudioStyle::Stereo) {
            follow_stereo(
                &left_snap,
                &right_snap,
                n,
                dt,
                attack_tau,
                release_tau,
                p.contrast,
                &mut stereo_l,
                &mut stereo_r,
            );
        }
        if matches!(p.style, AudioStyle::Pitch) {
            follow_pitch(&wave_snap, dt, &mut pitch_hz, &mut pitch_hue);
        }
        if matches!(p.style, AudioStyle::KeyColor) {
            follow_key(&wave_snap, dt, &mut key_hue);
        }
        if matches!(p.style, AudioStyle::Lissajous) {
            let target = stereo_width(&left_snap, &right_snap);
            liss_width = ema_toward(liss_width, target, dt, 0.09);
        }
        if matches!(p.style, AudioStyle::MidSide) {
            follow_midside(&left_snap, &right_snap, dt, attack_tau, release_tau, &mut stereo_l);
        }
        if matches!(p.style, AudioStyle::Bubbles) {
            if sparkle_lamps.len() != n {
                sparkle_lamps.resize(n, 0.0);
            }
            tick_audio_bubbles(
                &mut bubbles,
                &mut sparkle_lamps,
                n,
                dt,
                flux_fast,
                loud,
                levels[0],
                p.punch,
                p.motion,
                &mut rng,
            );
        }

        if matches!(p.style, AudioStyle::Eq24 | AudioStyle::Wavelength | AudioStyle::Melt) {
            let mut raw = if classic {
                analyze_eq24(&wave_snap, SAMPLE_RATE)
            } else {
                dsp.eq24
            };
            if classic {
                blur_neighbors(&mut raw);
            }
            if column.eq.len() != n {
                column.eq = vec![0.0; n];
            }
            let tau = 0.045 + p.smoothness * 0.14;
            for i in 0..n {
                let t = lamps::pos(i, n);
                let src = t * (raw.len().saturating_sub(1) as f32);
                let lo = src.floor() as usize;
                let hi = (lo + 1).min(raw.len() - 1);
                let f = src - lo as f32;
                let target = raw[lo] * (1.0 - f) + raw[hi] * f;
                column.eq[i] = ema_toward(column.eq[i], target, dt, tau);
            }
            blur_neighbors(&mut column.eq);
        }
        if matches!(p.style, AudioStyle::PanNeedle) {
            let (l, r) = stereo_gains(&left_snap, &right_snap, p.contrast);
            let target = if l + r < 1e-4 { 0.5 } else { r / (l + r) };
            column.pan = ema_toward(column.pan, target.clamp(0.0, 1.0), dt, 0.07);
        }
        column.centroid = if classic {
            spectral_centroid(&levels)
        } else {
            dsp.centroid
        };

        let rgb = if volume_dead {
            sparkle = [0.0; 4];
            sparkle_lamps.fill(0.0);
            levels = [0.0; 4];
            strobe = 0.0;
            vu_hold = 0.0;
            vu_peak = 0.0;
            tempo_phase = 0.0;
            scope_hold.fill(0.0);
            spec_hist_e.fill(0.0);
            spec_hist_h.fill(0.0);
            spec_hist_bands = [[0.0; 4]; 4];
            stereo_l.fill(0.0);
            stereo_r.fill(0.0);
            bubbles.clear();
            sparkle_lamps.fill(0.0);
            color_ripples.clear();
            column.reset();
            idle_mix = 1.0;
            drop_trough = 1.0;
            drop_loud = 0.0;
            shock_cool = 0.0;
            drop_flash_age = 1.0;
            beats.reset();
            flux_beats.reset();
            davies.reset();
            auto_pick.reset();
            if let Ok(mut live) = params.lock() {
                live.bpm = 0;
                live.bpm_locked = false;
                live.drop = false;
            }
            vec![[0u8; 3]; n]
        } else {
            let (ibi, locked) = if classic {
                (beats.ibi(), beats.tempo_locked())
            } else if matches!(engine, AudioAnalysis::Tempo) || davies.locked() {
                (davies.ibi(), davies.locked())
            } else {
                (flux_beats.ibi(), flux_beats.tempo_locked())
            };
            render_audio(
                p,
                &levels,
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
                &wave_snap,
                &mut scope_hold,
                &mut scope_peak,
                &mut spec_hist_e,
                &mut spec_hist_h,
                &mut spec_hist_bands,
                &stereo_l,
                &stereo_r,
                &mut color_ripples,
                &mut prev_gate,
                dt,
                n,
                drop_hit,
                ring_hit,
                ibi,
                locked,
                if matches!(p.style, AudioStyle::KeyColor) { key_hue } else { pitch_hue },
                liss_width,
                &mut rng,
                &mut column,
            )
        };
        manager.paint_lamps(&rgb);

        thread::sleep(Duration::from_millis(if n > 4 { 33 } else { 8 }));
    }

    let _ = capture.join();
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
    ) -> audio_beats::BeatHits {
        self.since = (self.since + dt).min(4.0);
        self.cool = (self.cool - dt).max(0.0);
        for age in &mut self.kick_age {
            *age = (*age + dt).min(8.0);
        }
        if gated {
            self.onset_p = 0.0;
            return audio_beats::BeatHits::none();
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

        let mut hits = audio_beats::BeatHits::none();
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

fn eq24_stats(eq: &[f32; audio_dsp::EQ_BANDS]) -> (f32, f32) {
    let mut sum = 0.0f32;
    let mut low = 0.0f32;
    for (i, v) in eq.iter().enumerate() {
        sum += *v;
        if i < 3 {
            low += *v;
        }
    }
    let mean = sum / audio_dsp::EQ_BANDS as f32;
    let mut var = 0.0f32;
    for v in eq {
        let d = *v - mean;
        var += d * d;
    }
    let spread = (var / audio_dsp::EQ_BANDS as f32).sqrt();
    let bass = if sum > 1e-6 { (low / sum).clamp(0.0, 1.0) } else { 0.0 };
    (bass, spread.clamp(0.0, 1.0))
}

fn last_slice(samples: &[f32], n: usize) -> Vec<f32> {
    if n == 0 {
        return Vec::new();
    }
    if samples.len() >= n {
        samples[samples.len() - n..].to_vec()
    } else {
        let mut out = vec![0.0f32; n];
        let start = n - samples.len();
        out[start..].copy_from_slice(samples);
        out
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

fn skip_spatial_blur(style: AudioStyle) -> bool {
    matches!(
        style,
        AudioStyle::BeatGates
            | AudioStyle::Vu
            | AudioStyle::TempoPulse
            | AudioStyle::Oscilloscope
            | AudioStyle::Spectrogram
            | AudioStyle::Stereo
            | AudioStyle::Pitch
            | AudioStyle::Lissajous
            | AudioStyle::Bubbles
            | AudioStyle::KeyColor
            | AudioStyle::MidSide
            | AudioStyle::Eq24
            | AudioStyle::PanNeedle
            | AudioStyle::Collision
            | AudioStyle::Snake
            | AudioStyle::Gravcenter
            | AudioStyle::Melt
            | AudioStyle::Wavelength
    )
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
    wave: &[f32],
    scope_hold: &mut Vec<f32>,
    scope_peak: &mut f32,
    spec_hist_e: &mut Vec<f32>,
    spec_hist_h: &mut Vec<f32>,
    _spec_hist_bands: &mut [[f32; 4]; 4],
    left_levels: &[f32],
    right_levels: &[f32],
    color_ripples: &mut Vec<ColorRipple>,
    _prev_gate: &mut [f32; 4],
    dt: f32,
    lamp_count: usize,
    drop_hit: bool,
    ring_hit: bool,
    tempo_ibi: f32,
    tempo_locked: bool,
    pitch_hue: f32,
    liss_width: f32,
    rng: &mut impl rand::Rng,
    column: &mut ColumnFx,
) -> Vec<[u8; 3]> {
    let n = lamp_count.max(1);
    if sparkle_lamps.len() != n {
        sparkle_lamps.resize(n, 0.0);
    }
    let beat_floor = params.min_brightness as f32 / 100.0;
    let idle = params.idle_brightness as f32 / 100.0;
    let rest = idle_mix.clamp(0.0, 1.0);
    let motion = 0.35 + params.motion * 1.4;
    let fine = n > 4;
    let four = to_four(levels);
    let four_s = if skip_spatial_blur(params.style) {
        four
    } else {
        spatial_blur(&four, params.spread)
    };
    let drive = four_s[0].max(four_s[1]);
    let peak = four_s.iter().copied().fold(0.0f32, f32::max);
    *wave_phase = (*wave_phase + dt * (0.35 + motion * (0.6 + drive * 2.4))) % (2.0 * PI);

    let mut energies = expand_strips(&four_s, n);

    match params.style {
        AudioStyle::Levels => {
            energies = expand_meters(&four_s, n);
        }
        AudioStyle::Wave | AudioStyle::Gradient | AudioStyle::Center => {}
        AudioStyle::Mirror => {}
        AudioStyle::Pulse | AudioStyle::Pitch | AudioStyle::KeyColor => {
            let pulse = (four_s.iter().sum::<f32>() / 4.0).max(four_s[0] * 0.85);
            energies.fill(pulse);
        }
        AudioStyle::Bloom => {
            let mut col = [0.0f32; 4];
            col[0] = four[0];
            col[1] = four[1].max(four[0] * 0.62);
            col[2] = four[2].max(four[0] * 0.32 + four[1] * 0.45);
            col[3] = four[3].max(four[1] * 0.28 + four[2] * 0.40);
            energies = expand_meters(&col, n);
        }
        AudioStyle::Fire => {
            let mut col = [0.0f32; 4];
            col[0] = four[0];
            col[1] = (four[0] * 0.75 + four[1] * 0.55).min(1.0);
            col[2] = (four[1] * 0.55 + four[2] * 0.7).min(1.0);
            col[3] = (four[2] * 0.4 + four[3] * 0.85).min(1.0);
            energies = expand_meters(&col, n);
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
                let center = rng.random_range(0..n);
                let amt = 0.55 + peak * 0.45;
                for d in -1i32..=1 {
                    let i = center as i32 + d;
                    if i >= 0 && (i as usize) < n {
                        let fall = if d == 0 { 1.0 } else { 0.42 };
                        sparkle_lamps[i as usize] = sparkle_lamps[i as usize].max(amt * fall);
                    }
                }
            }
        }
        AudioStyle::Chase => {
            *chase = (*chase + dt * (1.1 + motion * 3.8 + drive * 2.2)) % n.max(1) as f32;
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
            let cols = n.max(1) as f32;
            let fill_x = vu_hold.clamp(0.0, 1.0) * cols;
            let mark = (vu_peak.clamp(0.0, 1.0) * (cols - 1.0).max(0.0)).round();
            energies.resize(n, 0.0);
            for i in 0..n {
                let col = i as f32;
                let bar = (fill_x - col).clamp(0.0, 1.0);
                let tip = if (col - mark).abs() < 0.51 { 0.92 } else { 0.0 };
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
            } else {
                (four_s.iter().sum::<f32>() / 4.0).max(four_s[0] * 0.85)
            };
            energies.resize(n, 0.0);
            energies.fill(pulse);
        }
        AudioStyle::Oscilloscope => {
            if scope_hold.len() != n {
                scope_hold.resize(n, 0.0);
            }
            let inst = wave.iter().fold(0.0f32, |a, s| a.max(s.abs()));
            if inst > *scope_peak {
                *scope_peak = inst;
            } else {
                *scope_peak = ema_toward(*scope_peak, inst, dt, 0.28).max(1.5e-4);
            }
            let gain = (0.7 + params.sensitivity * 0.55) / scope_peak.max(1.5e-4);
            let len = wave.len().max(1);
            energies.resize(n, 0.0);
            for c in 0..n {
                let idx = if n <= 1 { 0 } else { c * (len - 1) / (n - 1) };
                let v = (wave.get(idx).copied().unwrap_or(0.0).abs() * gain).clamp(0.0, 1.0);
                if v > scope_hold[c] {
                    scope_hold[c] = v;
                } else {
                    scope_hold[c] = ema_toward(scope_hold[c], v, dt, 0.05);
                }
                energies[c] = scope_hold[c];
            }
        }
        AudioStyle::Spectrogram => {
            if spec_hist_e.len() != n {
                spec_hist_e.resize(n, 0.0);
                spec_hist_h.resize(n, 0.5);
            }
            if n > 0 {
                spec_hist_e.copy_within(0..n.saturating_sub(1), 1);
                spec_hist_h.copy_within(0..n.saturating_sub(1), 1);
                spec_hist_e[0] = peak.max(drive);
                spec_hist_h[0] = spectral_centroid(&four_s);
            }
            energies = spec_hist_e.clone();
        }
        AudioStyle::Stereo => {
            energies.resize(n, 0.0);
            layout_stereo(&mut energies, left_levels, right_levels, levels);
        }
        AudioStyle::MidSide => {
            energies.resize(n, 0.0);
            layout_midside(&mut energies, left_levels);
        }
        AudioStyle::Eq24 => {
            if column.eq.len() == n {
                energies = column.eq.clone();
            } else {
                energies = expand_strips(&four_s, n);
            }
        }
        AudioStyle::PanNeedle => {
            energies = vec![0.0; n];
            let sigma = if n > 4 { 0.9 } else { 0.55 };
            stamp_gauss(&mut energies, column.pan, sigma);
            let loud = peak.max(drive);
            for slot in energies.iter_mut() {
                *slot *= 0.12 + 0.88 * loud;
            }
        }
        AudioStyle::Collision => {
            tick_collision(column, n, dt, flux, peak, drive, params.motion, sparkle_lamps);
            energies = vec![0.0; n];
            if column.clash_on {
                let sigma = if n > 4 { 0.7 } else { 0.45 };
                stamp_gauss(&mut energies, column.clash_l, sigma);
                stamp_gauss(&mut energies, column.clash_r, sigma);
            }
            if column.clash_flash > 0.02 {
                stamp_gauss(&mut energies, column.clash_x, if n > 4 { 1.4 } else { 0.7 });
                for slot in energies.iter_mut() {
                    *slot = (*slot + column.clash_flash * 0.55).min(1.0);
                }
            }
            for i in 0..n {
                energies[i] = (energies[i] + sparkle_lamps.get(i).copied().unwrap_or(0.0)).clamp(0.0, 1.0);
            }
        }
        AudioStyle::Snake => {
            energies = vec![0.0; n];
            tick_snake(column, &mut energies, n, dt, flux, peak, drive, params.motion, ring_hit);
        }
        AudioStyle::Gravcenter => {
            audio_flow::tick_grav(&mut column.grav_h, peak.max(drive), dt, params.smoothness);
            energies = audio_flow::grav_map(n, column.grav_h, *wave_phase);
        }
        AudioStyle::Melt => {
            let inject = if column.eq.len() == n {
                column.eq.clone()
            } else {
                expand_strips(&four_s, n)
            };
            audio_flow::tick_melt(&mut column.melt, &inject, dt, params.smoothness);
            energies = column.melt.clone();
        }
        AudioStyle::Wavelength => {
            audio_flow::tick_wavelength(&mut column.wave_scroll, dt, params.motion);
            let bands: Vec<f32> = if column.eq.len() == n {
                column.eq.clone()
            } else {
                expand_strips(&four_s, n)
            };
            energies = audio_flow::wavelength_energy(n, &bands, column.wave_scroll);
        }
        AudioStyle::Bubbles => {}
        AudioStyle::Lissajous => {
            let pulse = (four_s.iter().sum::<f32>() / 4.0).max(peak);
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
                AudioStyle::Pulse | AudioStyle::Strobe | AudioStyle::TempoPulse | AudioStyle::Pitch | AudioStyle::KeyColor => {
                    energies.first().copied().unwrap_or(0.0)
                }
                AudioStyle::Vu
                | AudioStyle::Oscilloscope
                | AudioStyle::Spectrogram
                | AudioStyle::Stereo
                | AudioStyle::MidSide
                | AudioStyle::Eq24
                | AudioStyle::PanNeedle
                | AudioStyle::Collision
                | AudioStyle::Snake
                | AudioStyle::Gravcenter
                | AudioStyle::Melt
                | AudioStyle::Wavelength => {
                    if fine {
                        local
                    } else {
                        hard_band(&four_e, i, n)
                    }
                }
                AudioStyle::Wave => {
                    let travel = 0.5 + 0.5 * ((*wave_phase) - t * 2.0 * PI).sin();
                    (local * 0.28 + travel * drive.max(peak) * 0.78).clamp(0.0, 1.0)
                }
                AudioStyle::Bloom => {
                    if fine {
                        local
                    } else {
                        lamps::sample_bands(&four_e, t)
                    }
                }
                AudioStyle::Fire => {
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
                AudioStyle::Lissajous => {
                    let loud = energies.first().copied().unwrap_or(0.0).max(peak);
                    let dist = (t - 0.5).abs() * 2.0;
                    let sigma = (0.11 + liss_width * 0.86).clamp(0.10, 0.98);
                    let geo = ((1.0 - dist / sigma).max(0.0).powf(1.28)).clamp(0.0, 1.0);
                    (geo * (0.08 + 0.92 * loud)).clamp(0.0, 1.0)
                }
                AudioStyle::Mirror => {
                    let half = (n / 2).max(2);
                    let src = i.min(n - 1 - i);
                    let t_src = lamps::pos(src, half) * 0.5;
                    lamps::sample_bands(&four_e, t_src)
                }
                AudioStyle::Sparkle => (local * 0.28 + sparkle_lamps[i]).clamp(0.0, 1.0),
                AudioStyle::Bubbles => {
                    let loud = energies.first().copied().unwrap_or(0.0).max(peak);
                    (sparkle_lamps.get(i).copied().unwrap_or(0.0) * (0.2 + 0.8 * loud)).clamp(0.0, 1.0)
                }
                AudioStyle::Chase => {
                    let dist = (i as f32 - *chase).abs().min(n as f32 - (i as f32 - *chase).abs());
                    let sigma = if n > 4 { 1.15f32 } else { 0.55f32 };
                    let spot = (-dist * dist / (2.0 * sigma * sigma)).exp();
                    (spot * (0.2 + drive.max(peak) * 0.9)).clamp(0.0, 1.0)
                }
                AudioStyle::Gradient => {
                    let g = (four[0] * 0.45 + four[1] * 0.3 + four[2] * 0.15 + four[3] * 0.1).clamp(0.0, 1.0);
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
            let beat_amt = if matches!(
                params.style,
                AudioStyle::BeatGates | AudioStyle::Strobe | AudioStyle::Stereo | AudioStyle::MidSide | AudioStyle::Bubbles | AudioStyle::Collision
            ) {
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
            if matches!(params.style, AudioStyle::Pitch | AudioStyle::KeyColor) {
                let (pr, pg, pb) = audio_color::hsv_to_rgb(pitch_hue.rem_euclid(360.0), 0.88, 1.0);
                rgb = [pr, pg, pb];
            }
            if matches!(params.style, AudioStyle::Wavelength) {
                let hue = audio_flow::wavelength_hue(t, column.wave_scroll, column.centroid);
                let (pr, pg, pb) = audio_color::hsv_to_rgb(hue, 0.92, 1.0);
                rgb = [pr, pg, pb];
            }
            if matches!(params.style, AudioStyle::Fire) {
                rgb = lamps::mix_rgb(fire_color(t), rgb, 0.32);
            }
            let (mut r, mut g, mut b) = (rgb[0], rgb[1], rgb[2]);
            if matches!(params.style, AudioStyle::Spectrogram) {
                let centroid = spec_hist_h.get(lamp_col(i, n)).copied().unwrap_or(0.5);
                let hue_add = centroid * 70.0 + params.hue_shift * centroid * 40.0;
                let (h, s, v) = audio_color::rgb_to_hsv(r, g, b);
                (r, g, b) = audio_color::hsv_to_rgb((h + hue_add).rem_euclid(360.0), s.max(0.45), v);
            } else if params.hue_shift > 0.01
                && audio_color::hue_shift_ok(params.color_mode)
                && !matches!(
                    params.style,
                    AudioStyle::Fire | AudioStyle::Pitch | AudioStyle::KeyColor | AudioStyle::Wavelength
                )
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
    let n = levels.len();
    if n <= 4 {
        [
            levels.first().copied().unwrap_or(0.0),
            levels.get(1).copied().unwrap_or(0.0),
            levels.get(2).copied().unwrap_or(0.0),
            levels.get(3).copied().unwrap_or(0.0),
        ]
    } else {
        let rows = (n / 4).max(1);
        [
            levels.first().copied().unwrap_or(0.0),
            levels.get(rows).copied().unwrap_or(0.0),
            levels.get(rows * 2).copied().unwrap_or(0.0),
            levels.get(rows * 3).copied().unwrap_or(0.0),
        ]
    }
}

fn strip_count(n: usize) -> usize {
    n.max(1)
}

fn lamp_col(i: usize, n: usize) -> usize {
    i.min(n.saturating_sub(1))
}

fn meter_cell(energy: f32, row: usize, rows: usize) -> f32 {
    let e = energy.clamp(0.0, 1.0);
    if rows <= 1 {
        return e;
    }
    let y0 = row as f32 / rows as f32;
    let y1 = (row + 1) as f32 / rows as f32;
    if e >= y1 {
        1.0
    } else if e <= y0 {
        0.0
    } else {
        ((e - y0) / (y1 - y0).max(1e-4)).clamp(0.0, 1.0)
    }
}

fn expand_strips(four: &[f32], n: usize) -> Vec<f32> {
    (0..n)
        .map(|i| four.get((i * 4 / n.max(1)).min(3)).copied().unwrap_or(0.0))
        .collect()
}

fn expand_meters(four: &[f32], n: usize) -> Vec<f32> {
    if n <= 4 {
        return expand_strips(four, n);
    }
    let width = (n / 4).max(1);
    let mut out = vec![0.0; n];
    for z in 0..4 {
        let e = four.get(z).copied().unwrap_or(0.0);
        for k in 0..width {
            let i = z * width + k;
            if i < n {
                out[i] = meter_cell(e, k, width);
            }
        }
    }
    out
}

fn hard_band(levels: &[f32; 4], i: usize, n: usize) -> f32 {
    levels[(i * 4 / n.max(1)).min(3)]
}

fn spectral_centroid(levels: &[f32]) -> f32 {
    let mut wsum = 0.0;
    let mut sum = 0.0;
    for (i, v) in levels.iter().enumerate() {
        let e = v.max(0.0);
        wsum += e * (i as f32 + 0.5);
        sum += e;
    }
    if sum < 1e-6 || levels.is_empty() {
        0.5
    } else {
        (wsum / sum) / levels.len() as f32
    }
}

fn layout_stereo(out: &mut [f32], left: &[f32], right: &[f32], fallback: &[f32]) {
    let n = out.len();
    if n == 0 {
        return;
    }
    let l = if left.len() >= 2 { left } else { fallback };
    let r = if right.len() >= 2 { right } else { fallback };
    let lf = to_four(l);
    let rf = to_four(r);
    // Left half / right half. Do not put bass on both outer edges — a centered
    // kick would light far-left and far-right together and look un-split.
    let cols = [
        lf[0].max(lf[1] * 0.65),
        lf[2].max(lf[3]),
        rf[0].max(rf[1] * 0.65),
        rf[2].max(rf[3]),
    ];
    for i in 0..n {
        let z = (i * 4 / n.max(1)).min(3);
        out[i] = cols[z];
    }
}

fn stereo_gains(left: &[f32], right: &[f32], contrast: f32) -> (f32, f32) {
    let n = FAST_WIN.min(left.len()).min(right.len());
    if n == 0 {
        return (0.0, 0.0);
    }
    let sl = left.len() - n;
    let sr = right.len() - n;
    let width = (1.25 + (contrast - 1.0) * 1.4).clamp(1.25, 3.4);
    let mut l_acc = 0.0f32;
    let mut r_acc = 0.0f32;
    for i in 0..n {
        let l = left[sl + i];
        let r = right[sr + i];
        let mid = (l + r) * 0.5;
        let side = (l - r) * 0.5;
        let lw = mid + side * width;
        let rw = mid - side * width;
        l_acc += lw * lw;
        r_acc += rw * rw;
    }
    let l_rms = (l_acc / n as f32).sqrt();
    let r_rms = (r_acc / n as f32).sqrt();
    let peak = l_rms.max(r_rms).max(1e-5);
    let expand = (1.55 + contrast * 0.7).clamp(1.8, 3.8);
    (
        (l_rms / peak).powf(expand).clamp(0.0, 1.0),
        (r_rms / peak).powf(expand).clamp(0.0, 1.0),
    )
}

fn follow_stereo(
    left: &[f32],
    right: &[f32],
    _n: usize,
    dt: f32,
    attack_tau: f32,
    release_tau: f32,
    contrast: f32,
    stereo_l: &mut Vec<f32>,
    stereo_r: &mut Vec<f32>,
) {
    let (l_raw, _) = analyze_bands(left, SAMPLE_RATE);
    let (r_raw, _) = analyze_bands(right, SAMPLE_RATE);
    let spec_peak = l_raw
        .iter()
        .copied()
        .chain(r_raw.iter().copied())
        .fold(0.0f32, f32::max)
        .max(1e-5);
    let (l_gain, r_gain) = stereo_gains(left, right, contrast);

    if stereo_l.len() != 4 {
        stereo_l.resize(4, 0.0);
    }
    if stereo_r.len() != 4 {
        stereo_r.resize(4, 0.0);
    }
    for i in 0..4 {
        let l_t = (l_raw[i] / spec_peak).clamp(0.0, 1.0) * l_gain;
        let r_t = (r_raw[i] / spec_peak).clamp(0.0, 1.0) * r_gain;
        let l_tau = if l_t > stereo_l[i] { attack_tau } else { release_tau };
        let r_tau = if r_t > stereo_r[i] { attack_tau } else { release_tau };
        stereo_l[i] = ema_toward(stereo_l[i], l_t, dt, l_tau);
        stereo_r[i] = ema_toward(stereo_r[i], r_t, dt, r_tau);
    }
}

fn lerp_hue(a: f32, b: f32, t: f32) -> f32 {
    let mut d = (b - a).rem_euclid(360.0);
    if d > 180.0 {
        d -= 360.0;
    }
    (a + d * t.clamp(0.0, 1.0)).rem_euclid(360.0)
}

fn follow_pitch(samples: &[f32], dt: f32, pitch_hz: &mut f32, pitch_hue: &mut f32) {
    if let Some((hz, conf)) = detect_pitch(samples) {
        if conf > 2.2 {
            let alpha = if *pitch_hz < 1.0 { 0.45 } else { (1.0 - (-dt / 0.08).exp()).clamp(0.08, 0.5) };
            *pitch_hz = if *pitch_hz < 1.0 {
                hz
            } else {
                *pitch_hz * (1.0 - alpha) + hz * alpha
            };
            let chroma = (*pitch_hz / 16.35159783).log2().rem_euclid(1.0);
            *pitch_hue = lerp_hue(*pitch_hue, chroma * 360.0, 0.28);
        }
    }
}

fn detect_pitch(samples: &[f32]) -> Option<(f32, f32)> {
    let (re, im) = compute_fft(samples)?;
    let nyquist = FFT_SIZE / 2;
    let bin_hz = SAMPLE_RATE as f32 / FFT_SIZE as f32;
    let mut mag = [0.0f32; FFT_SIZE / 2];
    let mut mag_sum = 0.0f32;
    for i in 1..nyquist {
        mag[i] = (re[i] * re[i] + im[i] * im[i]).sqrt();
        mag_sum += mag[i];
    }
    let mag_mean = mag_sum / (nyquist.saturating_sub(1) as f32).max(1.0);
    let i_lo = (70.0 / bin_hz).round().max(1.0) as usize;
    let i_hi = (1000.0 / bin_hz).round().min((nyquist / 3 - 1) as f32) as usize;
    if i_lo >= i_hi {
        return None;
    }
    let mut best_i = i_lo;
    let mut best_p = 0.0f32;
    let mut hps_sum = 0.0f32;
    let mut hps_n = 0.0f32;
    for i in i_lo..=i_hi {
        let i2 = i * 2;
        let i3 = i * 3;
        if i3 >= nyquist {
            break;
        }
        let p = mag[i] * mag[i2] * mag[i3];
        hps_sum += p;
        hps_n += 1.0;
        if p > best_p {
            best_p = p;
            best_i = i;
        }
    }
    let hps_mean = (hps_sum / hps_n.max(1.0)).max(1e-12);
    let conf = best_p / hps_mean;
    if conf < 2.2 || mag[best_i] < mag_mean * 1.65 {
        return None;
    }
    let y1 = mag[best_i];
    let y0 = mag[best_i.saturating_sub(1)];
    let y2 = mag[(best_i + 1).min(nyquist - 1)];
    let denom = y0 - 2.0 * y1 + y2;
    let delta = if denom.abs() > 1e-9 {
        (0.5 * (y0 - y2) / denom).clamp(-0.5, 0.5)
    } else {
        0.0
    };
    Some(((best_i as f32 + delta) * bin_hz, conf))
}

fn stereo_width(left: &[f32], right: &[f32]) -> f32 {
    let n = FAST_WIN.min(left.len()).min(right.len());
    if n == 0 {
        return 0.0;
    }
    let sl = left.len() - n;
    let sr = right.len() - n;
    let mut lr = 0.0f32;
    let mut l2 = 0.0f32;
    let mut r2 = 0.0f32;
    for i in 0..n {
        let l = left[sl + i];
        let r = right[sr + i];
        lr += l * r;
        l2 += l * l;
        r2 += r * r;
    }
    let rms_l = (l2 / n as f32).sqrt();
    let rms_r = (r2 / n as f32).sqrt();
    if rms_l + rms_r < 1.5e-4 {
        return 0.0;
    }
    let corr = (lr / (l2 * r2).sqrt().max(1e-9)).clamp(-1.0, 1.0);
    let imbalance = (rms_l - rms_r).abs() / (rms_l + rms_r + 1e-8);
    ((1.0 - corr.max(0.0)) * 0.72 + imbalance * 0.55).clamp(0.0, 1.0)
}

struct AudioBubble {
    x: f32,
    y: f32,
    age: f32,
    life: f32,
    amp: f32,
}

fn tick_audio_bubbles(
    bubbles: &mut Vec<AudioBubble>,
    lamps: &mut [f32],
    n: usize,
    dt: f32,
    flux: f32,
    loud: f32,
    bass: f32,
    punch: f32,
    motion: f32,
    rng: &mut impl rand::Rng,
) {
    lamps.fill(0.0);
    let hit = flux > (0.034 - punch * 0.01).max(0.014) && (bass > 0.12 || loud > 0.18);
    if hit && bubbles.len() < 14 && rng.random::<f32>() < (0.22 + loud * 0.55 + punch * 0.12) {
        bubbles.push(AudioBubble {
            x: rng.random::<f32>(),
            y: 0.0,
            age: 0.0,
            life: 0.45 + 0.55 * (1.0 - punch.min(1.0)),
            amp: (0.55 + loud * 0.45).clamp(0.4, 1.0),
        });
    }
    let rise = (0.55 + motion * 1.1) * dt;
    for bubble in bubbles.iter_mut() {
        bubble.age += dt;
        bubble.x = (bubble.x + rise * 0.85).min(1.2);
        bubble.y = bubble.x;
    }
    bubbles.retain(|b| b.age < b.life && b.x < 1.15);

    let cols = n.min(lamps.len());
    for bubble in bubbles.iter() {
        let fade = (1.0 - bubble.age / bubble.life).clamp(0.0, 1.0);
        let amt = bubble.amp * fade;
        // x can travel past 1.0 before retain; never index with that raw column.
        let z = (bubble.x * cols as f32).floor() as i32;
        for (delta, falloff) in [(0i32, 1.0f32), (-1, 0.45), (1, 0.28)] {
            let i = z + delta;
            if i >= 0 && (i as usize) < cols {
                lamps[i as usize] = lamps[i as usize].max(amt * falloff);
            }
        }
    }
}

fn layout_midside(out: &mut [f32], bands: &[f32]) {
    let n = out.len();
    if n == 0 {
        return;
    }
    let cols = to_four(bands);
    for i in 0..n {
        let z = (i * 4 / n.max(1)).min(3);
        out[i] = cols[z];
    }
}

fn follow_midside(
    left: &[f32],
    right: &[f32],
    dt: f32,
    attack_tau: f32,
    release_tau: f32,
    out: &mut Vec<f32>,
) {
    let n = FAST_WIN.min(left.len()).min(right.len());
    if out.len() != 4 {
        out.resize(4, 0.0);
    }
    if n == 0 {
        return;
    }
    let sl = left.len() - n;
    let sr = right.len() - n;
    let mut mid_acc = 0.0f32;
    let mut side_l_acc = 0.0f32;
    let mut side_r_acc = 0.0f32;
    for i in 0..n {
        let l = left[sl + i];
        let r = right[sr + i];
        let mid = (l + r) * 0.5;
        let side = (l - r) * 0.5;
        mid_acc += mid * mid;
        side_l_acc += side.max(0.0) * side.max(0.0);
        side_r_acc += (-side).max(0.0) * (-side).max(0.0);
    }
    let inv = 1.0 / n as f32;
    let mid = (mid_acc * inv).sqrt();
    let side_l = (side_l_acc * inv).sqrt();
    let side_r = (side_r_acc * inv).sqrt();
    let peak = mid.max(side_l).max(side_r).max(1e-5);
    let target = [
        (side_l / peak).clamp(0.0, 1.0),
        (mid / peak).clamp(0.0, 1.0),
        (mid / peak).clamp(0.0, 1.0),
        (side_r / peak).clamp(0.0, 1.0),
    ];
    for i in 0..4 {
        let tau = if target[i] > out[i] { attack_tau } else { release_tau };
        out[i] = ema_toward(out[i], target[i], dt, tau);
    }
}

const KS_MAJOR: [f32; 12] = [6.35, 2.23, 3.48, 2.33, 4.38, 4.09, 2.52, 5.19, 2.39, 3.66, 2.29, 2.88];
const KS_MINOR: [f32; 12] = [6.33, 2.68, 3.52, 5.38, 2.60, 3.53, 2.54, 4.75, 3.98, 2.69, 3.34, 3.17];

fn pearson_shifted(chroma: &[f32; 12], profile: &[f32; 12], root: usize) -> f32 {
    let n = 12.0f32;
    let c_mean = chroma.iter().sum::<f32>() / n;
    let p_mean = profile.iter().sum::<f32>() / n;
    let mut cov = 0.0f32;
    let mut vc = 0.0f32;
    let mut vp = 0.0f32;
    for i in 0..12 {
        let c = chroma[(i + root) % 12] - c_mean;
        let p = profile[i] - p_mean;
        cov += c * p;
        vc += c * c;
        vp += p * p;
    }
    let denom = (vc * vp).sqrt();
    if denom < 1e-9 {
        0.0
    } else {
        cov / denom
    }
}

fn detect_key(samples: &[f32]) -> Option<(usize, bool, f32)> {
    let (re, im) = compute_fft(samples)?;
    let nyquist = FFT_SIZE / 2;
    let bin_hz = SAMPLE_RATE as f32 / FFT_SIZE as f32;
    let mut chroma = [0.0f32; 12];
    let mut total = 0.0f32;
    for i in 1..nyquist {
        let freq = i as f32 * bin_hz;
        if !(55.0..=2000.0).contains(&freq) {
            continue;
        }
        let mag = (re[i] * re[i] + im[i] * im[i]).sqrt();
        let midi = 69.0 + 12.0 * (freq / 440.0).log2();
        let pc = midi.round().rem_euclid(12.0) as usize;
        chroma[pc.min(11)] += mag;
        total += mag;
    }
    if total < 1e-5 {
        return None;
    }
    let mut best = 0usize;
    let mut best_minor = false;
    let mut best_c = f32::NEG_INFINITY;
    for root in 0..12 {
        let maj = pearson_shifted(&chroma, &KS_MAJOR, root);
        if maj > best_c {
            best_c = maj;
            best = root;
            best_minor = false;
        }
        let min = pearson_shifted(&chroma, &KS_MINOR, root);
        if min > best_c {
            best_c = min;
            best = root;
            best_minor = true;
        }
    }
    if best_c < 0.28 {
        return None;
    }
    Some((best, best_minor, best_c))
}

fn follow_key(samples: &[f32], dt: f32, key_hue: &mut f32) {
    if let Some((root, minor, conf)) = detect_key(samples) {
        if conf > 0.34 {
            let target = root as f32 * 30.0 + if minor { 12.0 } else { 0.0 };
            let alpha = (1.0 - (-dt / 0.22).exp()).clamp(0.04, 0.35);
            *key_hue = lerp_hue(*key_hue, target, alpha);
        }
    }
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
        let max_i = (strip_count(n) as f32 - 1.0).max(1.0);
        let speed = if shock { tune.travel() * 1.65 } else { tune.travel() };
        let strip_speed = (speed * max_i).max(2.0);
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
    let max = (strip_count(n) as i32 - 1).max(0);
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
    let max_i = (strip_count(n) as i32 - 1).max(1);
    let col = lamp_col(i, n) as i32;
    for ripple in ripples {
        let origin = origin_strip(ripple.origin, n);
        let shock = ripple.shock;
        let travel = if shock { tune.travel() * 1.65 } else { tune.travel() };
        let strip_speed = (travel * max_i as f32).max(2.0);
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
        let dist = (col - origin).abs() as f32;
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

const EQ_BANDS: usize = 24;

fn analyze_eq24(samples: &[f32], sample_rate: u32) -> [f32; EQ_BANDS] {
    let Some((re, im)) = compute_fft(samples) else {
        return [0.0; EQ_BANDS];
    };
    let lo_hz = 30.0f32;
    let hi_hz = 12_000.0f32;
    let ratio = hi_hz / lo_hz;
    let n = EQ_BANDS as f32;
    let mut bands = [0.0f32; EQ_BANDS];
    for i in 0..EQ_BANDS {
        let t0 = i as f32 / n;
        let t1 = (i + 1) as f32 / n;
        let mid_t = (t0 + t1) * 0.5;
        let mid = lo_hz * ratio.powf(mid_t);
        let lo = (lo_hz * ratio.powf(t0) * 0.72).max(20.0);
        let hi = (lo_hz * ratio.powf(t1) * 1.38).min(16_000.0);
        bands[i] = triangle_band(&re, &im, sample_rate, lo, mid, hi);
        let tilt = (mid / 400.0).powf(-0.18).clamp(0.55, 1.85);
        bands[i] *= tilt;
    }
    let peak = bands.iter().copied().fold(0.0f32, f32::max).max(1e-6);
    for band in &mut bands {
        *band = (*band / peak).clamp(0.0, 1.0);
    }
    bands
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

struct ColumnFx {
    eq: Vec<f32>,
    pan: f32,
    clash_on: bool,
    clash_l: f32,
    clash_r: f32,
    clash_flash: f32,
    clash_x: f32,
    snake_head: f32,
    grav_h: f32,
    melt: Vec<f32>,
    wave_scroll: f32,
    centroid: f32,
}

impl ColumnFx {
    fn new() -> Self {
        Self {
            eq: Vec::new(),
            pan: 0.5,
            clash_on: false,
            clash_l: 0.0,
            clash_r: 1.0,
            clash_flash: 0.0,
            clash_x: 0.5,
            snake_head: 0.0,
            grav_h: 0.0,
            melt: Vec::new(),
            wave_scroll: 0.0,
            centroid: 0.5,
        }
    }

    fn reset(&mut self) {
        self.eq.fill(0.0);
        self.pan = 0.5;
        self.clash_on = false;
        self.clash_l = 0.0;
        self.clash_r = 1.0;
        self.clash_flash = 0.0;
        self.clash_x = 0.5;
        self.snake_head = 0.0;
        self.grav_h = 0.0;
        self.melt.fill(0.0);
        self.wave_scroll = 0.0;
        self.centroid = 0.5;
    }
}

fn stamp_gauss(out: &mut [f32], x: f32, sigma: f32) {
    let n = out.len();
    if n == 0 {
        return;
    }
    let center = x.clamp(0.0, 1.0) * (n.saturating_sub(1) as f32);
    let w = sigma.max(0.22);
    for (i, slot) in out.iter_mut().enumerate() {
        let d = i as f32 - center;
        *slot = (*slot + (-d * d / (2.0 * w * w)).exp()).min(1.0);
    }
}

fn tick_collision(
    column: &mut ColumnFx,
    n: usize,
    dt: f32,
    flux: f32,
    peak: f32,
    drive: f32,
    motion: f32,
    sparkle_lamps: &mut Vec<f32>,
) {
    if sparkle_lamps.len() != n {
        sparkle_lamps.resize(n, 0.0);
    }
    for spark in sparkle_lamps.iter_mut() {
        *spark *= (1.0 - dt * 8.0).max(0.0);
    }
    column.clash_flash = (column.clash_flash - dt * 3.4).max(0.0);
    let hit = flux > 0.032 && (drive > 0.12 || peak > 0.16);
    if hit && !column.clash_on && column.clash_flash < 0.08 {
        column.clash_on = true;
        column.clash_l = 0.0;
        column.clash_r = 1.0;
    }
    if column.clash_on {
        let speed = (1.45 + motion * 1.35 + flux * 1.1).clamp(1.1, 3.4);
        column.clash_l = (column.clash_l + speed * dt).min(1.0);
        column.clash_r = (column.clash_r - speed * dt).max(0.0);
        if column.clash_l + 0.04 >= column.clash_r {
            column.clash_x = ((column.clash_l + column.clash_r) * 0.5).clamp(0.0, 1.0);
            column.clash_flash = (0.72 + peak * 0.28).clamp(0.7, 1.0);
            column.clash_on = false;
            let center = (column.clash_x * n.saturating_sub(1) as f32).round() as i32;
            for d in -2i32..=2 {
                let i = center + d;
                if i >= 0 && (i as usize) < n {
                    let fall = 1.0 - d.unsigned_abs() as f32 * 0.28;
                    sparkle_lamps[i as usize] = sparkle_lamps[i as usize].max(column.clash_flash * fall);
                }
            }
        }
    }
}

fn tick_snake(
    column: &mut ColumnFx,
    energies: &mut [f32],
    n: usize,
    dt: f32,
    flux: f32,
    peak: f32,
    drive: f32,
    motion: f32,
    ring_hit: bool,
) {
    if n == 0 {
        return;
    }
    let crawl = (0.55 + motion * 2.4 + flux * 5.0) * n as f32;
    let nudge = if ring_hit || flux > 0.045 { n as f32 * 0.35 } else { 0.0 };
    column.snake_head = (column.snake_head + (crawl + nudge) * dt).rem_euclid(n as f32);
    let loud = peak.max(drive);
    let max_len = n.max(2);
    let len = (1.0 + loud * (max_len as f32 - 1.0)).round().clamp(1.0, max_len as f32) as usize;
    let head = column.snake_head.floor() as i32;
    for k in 0..len {
        let i = (head - k as i32).rem_euclid(n as i32) as usize;
        let fade = (1.0 - k as f32 / len as f32).powf(1.25);
        energies[i] = energies[i].max((0.18 + 0.82 * loud) * fade);
    }
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
#[derive(Clone)]
struct CapBuf {
    mono: Vec<f32>,
    left: Vec<f32>,
    right: Vec<f32>,
    wrote: u64,
}

#[cfg(target_os = "windows")]
fn capture_loopback(
    samples: Arc<Mutex<Vec<f32>>>,
    stereo: Arc<Mutex<[Vec<f32>; 2]>>,
    seq: Arc<AtomicU64>,
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

    let per_device: Arc<Mutex<HashMap<String, CapBuf>>> = Arc::new(Mutex::new(HashMap::new()));
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
            let stereo = stereo.clone();
            let seq = seq.clone();
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
                        stereo,
                        seq,
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
    per_device: &Mutex<std::collections::HashMap<String, CapBuf>>,
    samples: &Mutex<Vec<f32>>,
    stereo: &Mutex<[Vec<f32>; 2]>,
    seq: &AtomicU64,
    preferred: &Mutex<std::collections::HashSet<String>>,
) {
    let (buf, wrote) = {
        let Ok(mut map) = per_device.lock() else {
            return;
        };
        if map.is_empty() {
            (
                CapBuf {
                    mono: vec![0.0f32; RING],
                    left: vec![0.0f32; RING],
                    right: vec![0.0f32; RING],
                    wrote: 0,
                },
                0u64,
            )
        } else {
            let pref = preferred.lock().ok();
            let is_pref = |id: &str| pref.as_ref().map(|p| p.contains(id)).unwrap_or(false);
            let mut best_pref: Option<(f32, String)> = None;
            let mut best_any: Option<(f32, String)> = None;
            for (id, buf) in map.iter() {
                let rms = recent_rms(&buf.mono);
                if best_any.as_ref().map(|(e, _)| rms > *e).unwrap_or(true) {
                    best_any = Some((rms, id.clone()));
                }
                if is_pref(id) && best_pref.as_ref().map(|(e, _)| rms > *e).unwrap_or(true) {
                    best_pref = Some((rms, id.clone()));
                }
            }
            let picked_id = match (best_pref, best_any) {
                (Some((pref_rms, pref_id)), Some((any_rms, any_id)))
                    if any_rms > pref_rms * 2.8 && any_rms > 0.012 =>
                {
                    any_id
                }
                (Some((_, pref_id)), _) => pref_id,
                (_, Some((_, any_id))) => any_id,
                _ => return,
            };
            if let Some(picked) = map.get_mut(&picked_id) {
                let wrote = picked.wrote;
                picked.wrote = 0;
                (picked.clone(), wrote)
            } else {
                return;
            }
        }
    };
    if let Ok(mut guard) = samples.lock() {
        *guard = buf.mono;
        if wrote > 0 {
            seq.fetch_add(wrote, Ordering::Relaxed);
        }
    }
    if let Ok(mut guard) = stereo.lock() {
        *guard = [buf.left, buf.right];
    }
}

#[cfg(target_os = "windows")]
fn capture_one_output(
    id: &str,
    name: &str,
    samples: Arc<Mutex<Vec<f32>>>,
    stereo: Arc<Mutex<[Vec<f32>; 2]>>,
    seq: Arc<AtomicU64>,
    per_device: Arc<Mutex<std::collections::HashMap<String, CapBuf>>>,
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
    let mix_ch = audio_client
        .get_mixformat()
        .map(|fmt| fmt.get_nchannels() as usize)
        .unwrap_or(2)
        .max(1);
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
    legion_rgb_driver::debug_log(&format!(
        "AUDIO: WASAPI loopback started on {name} mix_ch={mix_ch} capture_ch=2"
    ));

    let mut queue: VecDeque<u8> = VecDeque::new();
    let mut ring = vec![0.0f32; RING];
    let mut ring_l = vec![0.0f32; RING];
    let mut ring_r = vec![0.0f32; RING];
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
        let mut wrote = 0u64;
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
            ring_l[write_at] = l;
            ring_r[write_at] = r;
            ring[write_at] = (l + r) * 0.5;
            write_at = (write_at + 1) % RING;
            wrote += 1;
        }
        let unwrap_ring = |src: &[f32]| {
            let mut window = vec![0.0f32; RING];
            let (tail, head) = src.split_at(write_at);
            let split = RING - write_at;
            window[split..].copy_from_slice(tail);
            window[..split].copy_from_slice(head);
            window
        };
        if let Ok(mut map) = per_device.lock() {
            if let Some(old) = map.get(id) {
                wrote = wrote.saturating_add(old.wrote);
            }
            map.insert(
                id.to_string(),
                CapBuf {
                    mono: unwrap_ring(&ring),
                    left: unwrap_ring(&ring_l),
                    right: unwrap_ring(&ring_r),
                    wrote,
                },
            );
        }
        publish_mix(&per_device, &samples, &stereo, &seq, &preferred);
    }

    let _ = audio_client.stop_stream();
    if let Ok(mut map) = per_device.lock() {
        map.remove(id);
    }
    publish_mix(&per_device, &samples, &stereo, &seq, &preferred);
    live.fetch_sub(1, Ordering::Relaxed);
    capture_ok.store(live.load(Ordering::Relaxed) > 0, Ordering::Relaxed);
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;
    use rand::SeedableRng;
    use strum::IntoEnumIterator;

    fn lum(c: [u8; 3]) -> i32 {
        c[0] as i32 + c[1] as i32 + c[2] as i32
    }

    fn all_same(rgb: &[[u8; 3]]) -> bool {
        rgb.windows(2).all(|w| w[0] == w[1])
    }

    fn test_params(style: AudioStyle, ripple: bool) -> AudioReactParams {
        let mut p = AudioReactParams::default();
        p.style = style;
        p.min_brightness = 0;
        p.idle_brightness = 0;
        p.color_ramp = 0.0;
        p.spread = 0.0;
        p.ripple_color = ripple;
        p.color_mode = AudioColorMode::Custom;
        p.custom_rgb = [255; 12];
        p
    }

    fn render_style(style: AudioStyle, n: usize, levels: [f32; 4], ripple: bool) -> Vec<[u8; 3]> {
        let p = test_params(style, ripple);
        let mut wave_phase = 0.4f32;
        let mut tint_state = vec![[0.0; 3]; n];
        let mut sparkle = [0.0f32; 4];
        let mut sparkle_lamps = vec![0.0f32; n];
        let mut chase = 1.2f32;
        let mut strobe = 0.0f32;
        let mut vu_hold = 0.7f32;
        let mut vu_peak = 0.85f32;
        let mut tempo_phase = 0.0f32;
        let wave: Vec<f32> = (0..256).map(|i| ((i as f32) * 0.2).sin() * 0.4).collect();
        let mut scope_hold = Vec::new();
        let mut scope_peak = 0.002f32;
        let mut spec_hist_e = Vec::new();
        let mut spec_hist_h = Vec::new();
        let mut spec_hist_bands = [[0.0f32; 4]; 4];
        let left = levels.to_vec();
        let right = [levels[3], levels[2], levels[1], levels[0]].to_vec();
        let mut color_ripples = Vec::new();
        let mut prev_gate = [0.0f32; 4];
        let mut rng = rand::rngs::StdRng::seed_from_u64(7);
        let mut column = ColumnFx::new();
        if matches!(style, AudioStyle::Bubbles) {
            sparkle_lamps[0] = 0.8;
            if n > 6 {
                sparkle_lamps[6] = 0.5;
            }
        }
        if matches!(style, AudioStyle::Eq24 | AudioStyle::Wavelength | AudioStyle::Melt) {
            column.eq = (0..n).map(|i| i as f32 / (n.saturating_sub(1) as f32).max(1.0)).collect();
        }
        if matches!(style, AudioStyle::PanNeedle) {
            column.pan = 0.2;
        }
        render_audio(
            p,
            &levels,
            0.0,
            0.12,
            &mut wave_phase,
            0.3,
            &mut tint_state,
            &mut sparkle,
            &mut sparkle_lamps,
            &mut chase,
            &mut strobe,
            &mut vu_hold,
            &mut vu_peak,
            &mut tempo_phase,
            &wave,
            &mut scope_hold,
            &mut scope_peak,
            &mut spec_hist_e,
            &mut spec_hist_h,
            &mut spec_hist_bands,
            &left,
            &right,
            &mut color_ripples,
            &mut prev_gate,
            0.016,
            n,
            ripple,
            false,
            0.5,
            false,
            210.0,
            0.4,
            &mut rng,
            &mut column,
        )
    }

    fn render_spec_24(levels: [f32; 4]) -> Vec<[u8; 3]> {
        let p = test_params(AudioStyle::Spectrogram, false);
        let n = 24;
        let mut wave_phase = 0.0f32;
        let mut tint_state = vec![[0.0; 3]; n];
        let mut sparkle = [0.0f32; 4];
        let mut sparkle_lamps = vec![0.0f32; n];
        let mut chase = 0.0f32;
        let mut strobe = 0.0f32;
        let mut vu_hold = 0.0f32;
        let mut vu_peak = 0.0f32;
        let mut tempo_phase = 0.0f32;
        let wave = vec![0.0f32; 64];
        let mut scope_hold = Vec::new();
        let mut scope_peak = 0.002f32;
        let mut spec_hist_e = Vec::new();
        let mut spec_hist_h = Vec::new();
        let mut spec_hist_bands = [[0.0f32; 4]; 4];
        let left = levels.to_vec();
        let mut color_ripples = Vec::new();
        let mut prev_gate = [0.0f32; 4];
        let mut rng = rand::rngs::StdRng::seed_from_u64(1);
        let mut column = ColumnFx::new();
        let mut rgb = Vec::new();
        for _ in 0..4 {
            rgb = render_audio(
                p,
                &levels,
                0.0,
                0.05,
                &mut wave_phase,
                0.0,
                &mut tint_state,
                &mut sparkle,
                &mut sparkle_lamps,
                &mut chase,
                &mut strobe,
                &mut vu_hold,
                &mut vu_peak,
                &mut tempo_phase,
                &wave,
                &mut scope_hold,
                &mut scope_peak,
                &mut spec_hist_e,
                &mut spec_hist_h,
                &mut spec_hist_bands,
                &left,
                &left,
                &mut color_ripples,
                &mut prev_gate,
                0.016,
                n,
                false,
                false,
                0.5,
                false,
                210.0,
                0.2,
                &mut rng,
                &mut column,
            );
        }
        rgb
    }

    #[test]
    fn every_style_renders_zone_and_grid() {
        let levels = [0.2, 0.5, 0.8, 1.0];
        for style in AudioStyle::iter().filter(|s| !matches!(s, AudioStyle::Ripple)) {
            for n in [4usize, 24] {
                let rgb = render_style(style, n, levels, false);
                assert_eq!(rgb.len(), n, "{style:?} n={n} length");
            }
        }
        let rgb = render_style(AudioStyle::Levels, 24, levels, true);
        assert_eq!(rgb.len(), 24);
    }

    #[test]
    fn washes_stay_flat() {
        let levels = [0.25, 0.55, 0.8, 1.0];
        for style in [
            AudioStyle::Pulse,
            AudioStyle::Strobe,
            AudioStyle::TempoPulse,
            AudioStyle::Pitch,
            AudioStyle::KeyColor,
        ] {
            for n in [4usize, 24] {
                let rgb = render_style(style, n, levels, false);
                assert!(all_same(&rgb), "{style:?} n={n} should be a flat wash");
            }
        }
    }

    #[test]
    fn levels_zone_is_not_a_meter() {
        let rgb = render_style(AudioStyle::Levels, 4, [0.15, 0.45, 0.75, 1.0], false);
        assert!(lum(rgb[0]) < lum(rgb[1]));
        assert!(lum(rgb[1]) < lum(rgb[2]));
        assert!(lum(rgb[2]) < lum(rgb[3]));
    }

    #[test]
    fn levels_grid_fills_left_of_each_band() {
        let rgb = render_style(AudioStyle::Levels, 24, [0.5, 0.0, 0.0, 0.0], false);
        let col0: Vec<i32> = (0..6).map(|i| lum(rgb[i])).collect();
        assert!(col0[0] > col0[5], "band meter should fill from the left: {col0:?}");
        for w in col0.windows(2) {
            assert!(w[0] >= w[1], "band meter should be monotonic: {col0:?}");
        }
        assert!(col0.iter().any(|&a| col0.iter().any(|&b| a != b)), "24-lamp levels must vary inside a band");
    }

    #[test]
    fn spectrogram_scrolls_left_to_right() {
        let rgb = render_spec_24([1.0, 0.05, 0.05, 0.05]);
        assert!(lum(rgb[0]) > lum(rgb[23]), "newest energy should be on the left");
    }

    #[test]
    fn wave_travels_across_columns() {
        let rgb = render_style(AudioStyle::Wave, 24, [0.7, 0.6, 0.5, 0.4], false);
        let lumas: Vec<i32> = rgb.iter().copied().map(lum).collect();
        assert!(
            lumas.iter().any(|&a| lumas.iter().any(|&b| (a - b).abs() > 8)),
            "wave should vary across columns: {lumas:?}"
        );
    }

    #[test]
    fn beat_gates_light_whole_strip() {
        let rgb = render_style(AudioStyle::BeatGates, 24, [1.0, 0.0, 1.0, 0.0], false);
        for row in 1..6 {
            assert_eq!(rgb[row], rgb[0]);
            assert_eq!(rgb[12 + row], rgb[12]);
        }
        assert_ne!(lum(rgb[0]), lum(rgb[6]));
    }

    #[test]
    fn vu_fills_left_to_right() {
        let rgb = render_style(AudioStyle::Vu, 24, [0.7, 0.7, 0.7, 0.7], false);
        assert!(lum(rgb[0]) >= lum(rgb[8]));
        assert!(lum(rgb[8]) >= lum(rgb[20]) || lum(rgb[8]) + lum(rgb[0]) > lum(rgb[20]));
    }

    #[test]
    fn spectrogram_and_meter_helpers() {
        let meters = expand_meters(&[0.5, 0.0, 0.0, 0.0], 24);
        assert!((meters[0] - 1.0).abs() < 1e-5);
        assert!(meters[5] < 0.01);
        let strips = expand_strips(&[0.2, 0.4, 0.6, 0.8], 24);
        assert!((strips[0] - 0.2).abs() < 1e-5);
        assert!((strips[23] - 0.8).abs() < 1e-5);
        let zone = expand_meters(&[0.2, 0.4, 0.6, 0.8], 4);
        assert!((zone[0] - 0.2).abs() < 1e-5);
        assert!((zone[3] - 0.8).abs() < 1e-5);
    }

    #[test]
    fn eq24_grid_varies_across_columns() {
        let rgb = render_style(AudioStyle::Eq24, 24, [0.2, 0.5, 0.8, 1.0], false);
        let lumas: Vec<i32> = rgb.iter().copied().map(lum).collect();
        assert!(
            lumas.iter().any(|&a| lumas.iter().any(|&b| (a - b).abs() > 8)),
            "24-band EQ should vary across columns: {lumas:?}"
        );
        assert!(lum(rgb[0]) < lum(rgb[23]), "eq ramp should be brighter on the right: {lumas:?}");
    }

    #[test]
    fn gravcenter_is_brighter_in_the_middle() {
        let rgb = render_style(AudioStyle::Gravcenter, 24, [0.85, 0.85, 0.85, 0.85], false);
        assert!(
            lum(rgb[11]) + lum(rgb[12]) > lum(rgb[0]) + lum(rgb[23]),
            "gravcenter should fill from the middle"
        );
        let zone = render_style(AudioStyle::Gravcenter, 4, [0.85, 0.85, 0.85, 0.85], false);
        assert!(lum(zone[1]) + lum(zone[2]) > lum(zone[0]) + lum(zone[3]));
    }

    #[test]
    fn wavelength_varies_across_columns() {
        let rgb = render_style(AudioStyle::Wavelength, 24, [0.2, 0.5, 0.8, 1.0], false);
        let lumas: Vec<i32> = rgb.iter().copied().map(lum).collect();
        assert!(
            lumas.iter().any(|&a| lumas.iter().any(|&b| (a - b).abs() > 8)),
            "wavelength should vary: {lumas:?}"
        );
    }

    #[test]
    fn column_styles_render_zone_and_grid() {
        for style in [
            AudioStyle::Eq24,
            AudioStyle::PanNeedle,
            AudioStyle::Collision,
            AudioStyle::Snake,
            AudioStyle::Gravcenter,
            AudioStyle::Melt,
            AudioStyle::Wavelength,
        ] {
            for n in [4usize, 24] {
                let rgb = render_style(style, n, [0.3, 0.5, 0.7, 0.9], false);
                assert_eq!(rgb.len(), n, "{style:?} n={n}");
            }
        }
    }

    #[test]
    fn hud_reports_locked_bpm_and_drop() {
        let mut p = AudioReactParams::default();
        p.bpm = 120;
        p.bpm_locked = true;
        p.drop = true;
        let hud = p.hud();
        assert_eq!(hud.bpm, 120);
        assert!(hud.locked);
        assert!(hud.drop);
        p.bpm_locked = false;
        assert_eq!(p.hud().bpm, 0);
        assert!(!p.hud().locked);
        assert_eq!(audio_flow::bpm_from_ibi(0.5, true), 120);
        assert_eq!(audio_flow::bpm_from_ibi(0.5, false), 0);
    }

    #[test]
    fn audio_bubbles_do_not_panic_past_the_end() {
        let mut rng = rand::rngs::StdRng::seed_from_u64(1);
        for n in [4usize, 24] {
            let mut lamps = vec![0.0f32; n];
            let mut bubbles = vec![
                AudioBubble {
                    x: 1.0,
                    y: 1.0,
                    age: 0.0,
                    life: 1.0,
                    amp: 1.0,
                },
                AudioBubble {
                    x: 1.14,
                    y: 1.14,
                    age: 0.02,
                    life: 1.0,
                    amp: 0.8,
                },
            ];
            tick_audio_bubbles(&mut bubbles, &mut lamps, n, 0.016, 0.0, 0.0, 0.0, 0.0, 0.7, &mut rng);
            assert_eq!(lamps.len(), n);
            assert!(lamps.iter().all(|v| v.is_finite()));
        }
    }
}

