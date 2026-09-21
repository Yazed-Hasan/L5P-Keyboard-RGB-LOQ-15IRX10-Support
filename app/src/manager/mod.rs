use crate::enums::{Direction, Effects, Message, SwipeMode};

use crossbeam_channel::{Receiver, Sender};
use effects::{ambient, audio, aurora, battery, christmas, disco, fade, lightning, rain, ripple, scanner, stars, swipe, temperature};
use error_stack::{Result, ResultExt};
use legion_rgb_driver::{BaseEffects, Keyboard, SPEED_RANGE};
use profile::Profile;
use rand::{rng, rngs::ThreadRng, Rng};
use single_instance::SingleInstance;
use std::{
    sync::atomic::{AtomicBool, AtomicU8, Ordering},
    sync::Mutex,
    thread,
    time::Duration,
};
use std::{sync::Arc, thread::JoinHandle};
use thiserror::Error;

use self::custom_effect::{CustomEffect, EffectType};

pub mod custom_effect;
mod effects;
pub mod profile;

pub use effects::audio::AudioReactParams;
pub use effects::show_effect_ui;

#[derive(Debug, Error, PartialEq)]
#[error("Could not create keyboard manager")]
pub enum ManagerCreationError {
    #[error("There was an error getting a valid keyboard")]
    AcquireKeyboard,
    #[error("An instance of the program is already running")]
    InstanceAlreadyRunning,
}

/// Manager wrapper
pub struct EffectManager {
    pub tx: Sender<Message>,
    pub is_dynamic_lighting: bool,
    pub lamp_count: u16,
    pub window_active: Arc<AtomicBool>,
    effect_speed: Arc<AtomicU8>,
    audio_params: Arc<Mutex<audio::AudioReactParams>>,
    scene_params: Arc<Mutex<effects::scene::SceneLive>>,
    fine_lamps: Arc<AtomicBool>,
    inner_handle: Option<JoinHandle<()>>,
    stop_signals: StopSignals,
}

/// Controls the keyboard lighting logic
struct Inner {
    keyboard: Keyboard,
    rx: Receiver<Message>,
    stop_signals: StopSignals,
    last_profile: Profile,
    is_dynamic_lighting: bool,
    effect_speed: Arc<AtomicU8>,
    audio_params: Arc<Mutex<audio::AudioReactParams>>,
    scene_params: Arc<Mutex<effects::scene::SceneLive>>,
    fine_lamps: Arc<AtomicBool>,
    // Can't drop this else it stops "reserving" whatever underlying implementation identifier it uses
    #[allow(dead_code)]
    single_instance: SingleInstance,
}

#[derive(Clone, Copy)]
pub enum OperationMode {
    Cli,
    Gui,
}

impl EffectManager {
    pub fn new(operation_mode: OperationMode) -> Result<Self, ManagerCreationError> {
        let stop_signals = StopSignals {
            manager_stop_signal: Arc::new(AtomicBool::new(false)),
            keyboard_stop_signal: Arc::new(AtomicBool::new(false)),
        };

        // Use the crate's name as the identifier, should be unique enough
        let single_instance = SingleInstance::new(env!("CARGO_PKG_NAME")).unwrap();

        if !single_instance.is_single() {
            return Err(ManagerCreationError::InstanceAlreadyRunning.into());
        }

        let keyboard = legion_rgb_driver::get_keyboard(stop_signals.keyboard_stop_signal.clone())
            .change_context(ManagerCreationError::AcquireKeyboard)
            .attach_printable("Ensure that you have a supported model and that the application has access to it.")
            .attach_printable("On Linux, you may need to configure additional permissions")
            .attach_printable("https://github.com/4JX/L5P-Keyboard-RGB#usage")?;

        let is_dynamic_lighting = keyboard.is_dynamic_lighting();
        let lamp_count = keyboard.lamp_count();
        let window_active = keyboard.window_active_handle();
        let effect_speed = Arc::new(AtomicU8::new(1));
        let audio_params = Arc::new(Mutex::new(audio::AudioReactParams::default()));
        let scene_params = Arc::new(Mutex::new(effects::scene::SceneLive::default()));
        let fine_lamps = Arc::new(AtomicBool::new(false));

        let (tx, rx) = crossbeam_channel::unbounded::<Message>();

        let mut inner = Inner {
            keyboard,
            rx,
            stop_signals: stop_signals.clone(),
            last_profile: Profile::default(),
            is_dynamic_lighting,
            effect_speed: effect_speed.clone(),
            audio_params: audio_params.clone(),
            scene_params: scene_params.clone(),
            fine_lamps: fine_lamps.clone(),
            single_instance,
        };

        macro_rules! effect_thread_loop {
            ($e: expr) => {
                thread::spawn(move || loop {
                    match $e {
                        Some(message) => match message {
                            Message::Profile { profile } => {
                                inner.set_profile(profile);
                            }
                            Message::CustomEffect { effect } => {
                                inner.custom_effect(&effect);
                            }
                            Message::Exit => break,
                        },
                        None => {
                            // Keep-alive for legacy devices (and those without a threaded keep-alive)
                            // This ensures that when the app is minimized (or no UI changes occur),
                            // Lenovo Vantage doesn't overwrite the static RGB state.
                            let _ = inner.keyboard.refresh();
                            thread::sleep(Duration::from_millis(50));
                        }
                    }
                })
            };
        }

        let inner_handle = match operation_mode {
            OperationMode::Cli => effect_thread_loop!(inner.rx.try_recv().ok()),
            OperationMode::Gui => effect_thread_loop!(inner.rx.try_iter().last()),
        };

        let manager = Self {
            tx,
            is_dynamic_lighting,
            lamp_count,
            window_active,
            effect_speed,
            audio_params,
            scene_params,
            fine_lamps,
            inner_handle: Some(inner_handle),
            stop_signals,
        };

        Ok(manager)
    }

    pub fn set_profile(&mut self, profile: Profile) {
        self.effect_speed.store(profile.speed.max(1), Ordering::Relaxed);
        self.stop_signals.store_true();
        self.tx.try_send(Message::Profile { profile }).unwrap();
    }

    pub fn set_live_audio(&self, params: audio::AudioReactParams) {
        if let Ok(mut guard) = self.audio_params.lock() {
            *guard = params.normalized();
        }
    }

    pub fn set_live_scene(&self, effect: Effects, rgb: [u8; 12]) {
        if let Ok(mut guard) = self.scene_params.lock() {
            *guard = effects::scene::SceneLive { effect, rgb };
        }
    }

    pub fn set_live_speed(&self, speed: u8) {
        let speed = speed.clamp(1, 10);
        self.effect_speed.store(speed, Ordering::Relaxed);
        legion_rgb_driver::debug_log(&format!(
            "SPEED: live={} delay={}ms",
            speed,
            wdl_delay_ms(speed)
        ));
    }

    pub fn set_fine_lamps(&self, enabled: bool) {
        self.fine_lamps.store(enabled, Ordering::Relaxed);
        legion_rgb_driver::debug_log(&format!("LAMPS: fine_24={}", enabled));
    }

    pub fn fine_lamps(&self) -> bool {
        self.fine_lamps.load(Ordering::Relaxed)
    }

    pub fn custom_effect(&self, effect: CustomEffect) {
        self.stop_signals.store_true();
        self.tx.send(Message::CustomEffect { effect }).unwrap();
    }

    pub fn shutdown(mut self) {
        self.stop_signals.store_true();
        self.tx.send(Message::Exit).unwrap();
        if let Some(handle) = self.inner_handle.take() {
            handle.join().unwrap();
        };
    }
}

impl Inner {
    fn set_profile(&mut self, mut profile: Profile) {
        self.last_profile = profile.clone();
        self.stop_signals.store_false();
        self.effect_speed.store(profile.speed.max(1), Ordering::Relaxed);
        if let Ok(mut guard) = self.audio_params.lock() {
            *guard = audio::AudioReactParams::from_effect(profile.effect).with_rgb(profile.rgb_array());
        }
        if let Ok(mut guard) = self.scene_params.lock() {
            *guard = effects::scene::SceneLive {
                effect: profile.effect,
                rgb: profile.rgb_array(),
            };
        }
        legion_rgb_driver::debug_log(&format!(
            "EFFECT: switch to {:?} speed={}",
            profile.effect, profile.speed
        ));
        let mut rng = rng();

        if profile.effect.is_built_in() && !self.is_dynamic_lighting {
            let clamped_speed = self.clamp_speed(profile.speed);
            self.keyboard.set_speed(clamped_speed).unwrap();
        } else {
            // All custom effects rely on rapidly switching a static color
            self.keyboard.set_effect(BaseEffects::Static).unwrap();
        }

        if self.is_dynamic_lighting {
            self.keyboard.set_brightness_percent(profile.brightness_level).unwrap();
        } else {
            self.keyboard.set_brightness(profile.brightness as u8 + 1).unwrap();
        }

        self.apply_effect(&mut profile, &mut rng);
        self.stop_signals.store_false();
    }

    fn clamp_speed(&self, speed: u8) -> u8 {
        speed.clamp(SPEED_RANGE.min().unwrap(), SPEED_RANGE.max().unwrap())
    }

    fn live_speed(&self) -> u8 {
        self.effect_speed.load(Ordering::Relaxed).clamp(1, 10)
    }

    pub(crate) fn lamp_n(&self) -> usize {
        if self.fine_lamps.load(Ordering::Relaxed) {
            effects::lamps::count(self.keyboard.lamp_count())
        } else {
            4
        }
    }

    pub(crate) fn paint_lamps(&mut self, lamps: &[[u8; 3]]) {
        if self.fine_lamps.load(Ordering::Relaxed) && self.keyboard.lamp_count() >= 8 {
            let _ = self.keyboard.set_lamp_colors(lamps);
        } else {
            let rgb = legion_rgb_driver::zone_colors_from_lamps(lamps);
            let _ = self.keyboard.set_colors_to(&rgb);
        }
    }

    fn apply_profile_colors(&mut self, profile: &Profile) {
        if self.fine_lamps.load(Ordering::Relaxed) && self.keyboard.lamp_count() >= 8 && profile.has_lamp_colors() {
            let n = self.lamp_n();
            let _ = self.keyboard.set_lamp_colors(&profile.lamp_colors(n));
        } else {
            let _ = self.keyboard.set_colors_to(&profile.rgb_array());
        }
    }

    fn apply_effect(&mut self, profile: &mut Profile, rng: &mut ThreadRng) {
        match profile.effect {
            Effects::Static => {
                self.apply_profile_colors(profile);
                self.keyboard.set_effect(BaseEffects::Static).unwrap();
                if self.is_dynamic_lighting {
                    self.wdl_maintain();
                }
            }
            Effects::Breath => {
                self.apply_profile_colors(profile);
                if self.is_dynamic_lighting {
                    self.play_breath_wdl(profile);
                } else {
                    self.keyboard.set_effect(BaseEffects::Breath).unwrap();
                }
            }
            Effects::Smooth => {
                if self.is_dynamic_lighting {
                    self.play_smooth_wdl();
                } else {
                    self.keyboard.set_effect(BaseEffects::Smooth).unwrap();
                }
            }
            Effects::Wave => {
                if self.is_dynamic_lighting {
                    self.play_wave_wdl(profile.direction);
                } else {
                    let effect = match profile.direction {
                        Direction::Left => BaseEffects::LeftWave,
                        Direction::Right => BaseEffects::RightWave,
                    };
                    self.keyboard.set_effect(effect).unwrap();
                }
            }
            Effects::Lightning => lightning::play(self, profile, rng),
            Effects::AmbientLight { mut fps, mut saturation_boost } => {
                fps = fps.clamp(1, 60);
                saturation_boost = saturation_boost.clamp(0.0, 1.0);
                ambient::play(self, fps, saturation_boost);
            }
            Effects::SmoothWave { mode, clean_with_black } => {
                profile.rgb_zones = profile::arr_to_zones([255, 0, 0, 0, 255, 0, 0, 0, 255, 255, 0, 255]);
                if self.is_dynamic_lighting {
                    self.play_swipe_wdl(profile, mode, clean_with_black);
                } else {
                    swipe::play(self, profile, mode, clean_with_black);
                }
            }
            Effects::Swipe { mode, clean_with_black } => {
                if self.is_dynamic_lighting {
                    self.play_swipe_wdl(profile, mode, clean_with_black);
                } else {
                    swipe::play(self, profile, mode, clean_with_black);
                }
            }
            Effects::Disco => {
                if self.is_dynamic_lighting {
                    self.play_disco_wdl(profile, rng);
                } else {
                    disco::play(self, profile, rng);
                }
            }
            Effects::Christmas => christmas::play(self, rng),
            Effects::Fade => fade::play(self, profile),
            Effects::Temperature => temperature::play(self),
            Effects::Ripple => ripple::play(self, profile),
            Effects::AudioReact { .. } => audio::play(self, profile),
            Effects::Stars { params } => stars::play(self, params, profile.rgb_array()),
            Effects::Rain { params } => rain::play(self, params, profile.rgb_array()),
            Effects::Aurora { params } => aurora::play(self, params, profile.rgb_array()),
            Effects::Scanner { params } => scanner::play(self, params, profile.rgb_array()),
            Effects::Battery { params } => battery::play(self, params, profile.rgb_array()),
        }
    }

    fn custom_effect(&mut self, custom_effect: &CustomEffect) {
        self.stop_signals.store_false();

        loop {
            for step in &custom_effect.effect_steps {
                if self.is_dynamic_lighting {
                    // Map legacy 1-2 to percent for WDL
                    let percent = match step.brightness {
                        1 => 50u8,
                        2 => 100,
                        v => v.clamp(1, 100),
                    };
                    self.keyboard.set_brightness_percent(percent).unwrap();
                } else {
                    self.keyboard.set_brightness(step.brightness).unwrap();
                }
                match step.step_type {
                    EffectType::Set => {
                        self.keyboard.set_colors_to(&step.rgb_array).unwrap();
                    }
                    _ => {
                        self.keyboard.transition_colors_to(&step.rgb_array, step.steps, step.delay_between_steps).unwrap();
                    }
                }
                if self.stop_signals.manager_stop_signal.load(Ordering::SeqCst) {
                    return;
                }
                thread::sleep(Duration::from_millis(step.sleep));
            }
            if !custom_effect.should_loop {
                break;
            }
        }
    }

    /// Maintenance loop for WDL Static — keeps pushing colors so they persist through focus changes.
    fn wdl_maintain(&mut self) {
        while !self.stop_signals.manager_stop_signal.load(Ordering::SeqCst) {
            self.keyboard.refresh().ok();
            thread::sleep(Duration::from_millis(100));
        }
    }

    /// Software Breath effect for Windows Dynamic Lighting devices.
    fn play_breath_wdl(&mut self, profile: &Profile) {
        let n = self.lamp_n();
        let base = profile.lamp_colors(n);
        let mut phase: f64 = 0.0;
        while !self.stop_signals.manager_stop_signal.load(Ordering::SeqCst) {
            let factor = (phase.sin() * 0.5 + 0.5) as f32;
            let lamps: Vec<[u8; 3]> = base.iter().copied().map(|c| effects::lamps::scale_rgb(c, factor)).collect();
            self.paint_lamps(&lamps);
            phase += self.live_speed() as f64 * 0.08;
            if phase > std::f64::consts::TAU {
                phase -= std::f64::consts::TAU;
            }
            thread::sleep(Duration::from_millis(50));
        }
    }

    /// Software Smooth (rainbow cycle) effect for Windows Dynamic Lighting devices.
    fn play_smooth_wdl(&mut self) {
        let n = self.lamp_n();
        let mut hue: f64 = 0.0;
        legion_rgb_driver::debug_log(&format!("EFFECT: smooth-wave lamps={n}"));
        while !self.stop_signals.manager_stop_signal.load(Ordering::SeqCst) {
            if n <= 4 {
                let (r, g, b) = hsv_to_rgb(hue, 1.0, 1.0);
                let rgb = [r, g, b, r, g, b, r, g, b, r, g, b];
                let _ = self.keyboard.set_colors_to(&rgb);
            } else {
                let lamps: Vec<[u8; 3]> = (0..n)
                    .map(|i| {
                        let t = if n <= 1 { 0.0 } else { i as f64 / (n - 1) as f64 };
                        let (r, g, b) = hsv_to_rgb((hue + t * 90.0).rem_euclid(360.0), 1.0, 1.0);
                        [r, g, b]
                    })
                    .collect();
                self.paint_lamps(&lamps);
            }
            hue = (hue + self.live_speed() as f64 * 3.0) % 360.0;
            thread::sleep(Duration::from_millis(50));
        }
    }

    /// Software Wave effect for Windows Dynamic Lighting devices.
    fn play_wave_wdl(&mut self, direction: Direction) {
        let n = self.lamp_n();
        let mut hue: f64 = 0.0;
        let dir_mul: f64 = match direction {
            Direction::Left => 1.0,
            Direction::Right => -1.0,
        };
        legion_rgb_driver::debug_log(&format!("EFFECT: wave lamps={n} dir={direction:?}"));
        while !self.stop_signals.manager_stop_signal.load(Ordering::SeqCst) {
            if n <= 4 {
                let mut rgb = [0u8; 12];
                for z in 0..4 {
                    let h = (hue + z as f64 * 90.0 * dir_mul).rem_euclid(360.0);
                    let (r, g, b) = hsv_to_rgb(h, 1.0, 1.0);
                    rgb[z * 3] = r;
                    rgb[z * 3 + 1] = g;
                    rgb[z * 3 + 2] = b;
                }
                let _ = self.keyboard.set_colors_to(&rgb);
            } else {
                let lamps: Vec<[u8; 3]> = (0..n)
                    .map(|i| {
                        let t = if n <= 1 { 0.0 } else { i as f64 / n as f64 };
                        let h = (hue + t * 360.0 * dir_mul).rem_euclid(360.0);
                        let (r, g, b) = hsv_to_rgb(h, 1.0, 1.0);
                        [r, g, b]
                    })
                    .collect();
                self.paint_lamps(&lamps);
            }
            hue = (hue + self.live_speed() as f64 * 4.0) % 360.0;
            thread::sleep(Duration::from_millis(if n > 4 { 33 } else { 50 }));
        }
    }

    /// Fast swipe for WDL. 4-zone uses the original 4-step path; 24-lamp steps
    /// no faster than the 30fps painter so strips are not skipped.
    fn play_swipe_wdl(&mut self, profile: &Profile, mode: SwipeMode, clean_with_black: bool) {
        let n = self.lamp_n();
        if n <= 4 {
            self.play_swipe_wdl_zones(profile, mode, clean_with_black);
            return;
        }
        let stops = distinct_swipe_colors(profile.rgb_array());
        let palette = effects::lamps::lamps_from_zones(&stops, n);
        legion_rgb_driver::debug_log(&format!(
            "EFFECT: swipe WDL mode={mode:?} lamps={n} step_ms={}",
            swipe_step_ms(self.live_speed(), n)
        ));
        let mut shift = 0usize;
        while !self.stop_signals.manager_stop_signal.load(Ordering::SeqCst) {
            match mode {
                SwipeMode::Change => {
                    let lamps: Vec<[u8; 3]> = (0..n).map(|i| palette[(i + shift) % n]).collect();
                    self.paint_lamps(&lamps);
                    let stride = swipe_stride(self.live_speed(), n);
                    shift = match profile.direction {
                        Direction::Left => (shift + n - stride) % n,
                        Direction::Right => (shift + stride) % n,
                    };
                    thread::sleep(Duration::from_millis(swipe_step_ms(self.live_speed(), n)));
                }
                SwipeMode::Fill => {
                    let order: Vec<usize> = match profile.direction {
                        Direction::Left => (0..n).collect(),
                        Direction::Right => (0..n).rev().collect(),
                    };
                    let colors = [
                        effects::lamps::zone_rgb(&stops, 0),
                        effects::lamps::zone_rgb(&stops, 1),
                        effects::lamps::zone_rgb(&stops, 2),
                        effects::lamps::zone_rgb(&stops, 3),
                    ];
                    for color in colors {
                        let mut lamps = vec![[0u8; 3]; n];
                        let mut i = 0usize;
                        while i < order.len() {
                            let end = (i + swipe_stride(self.live_speed(), n)).min(order.len());
                            for &idx in &order[i..end] {
                                lamps[idx] = color;
                            }
                            self.paint_lamps(&lamps);
                            thread::sleep(Duration::from_millis(swipe_fill_ms(self.live_speed(), n)));
                            if self.stop_signals.manager_stop_signal.load(Ordering::SeqCst) {
                                return;
                            }
                            i = end;
                        }
                        if clean_with_black {
                            i = 0;
                            while i < order.len() {
                                let end = (i + swipe_stride(self.live_speed(), n)).min(order.len());
                                for &idx in &order[i..end] {
                                    lamps[idx] = [0; 3];
                                }
                                self.paint_lamps(&lamps);
                                thread::sleep(Duration::from_millis(swipe_fill_ms(self.live_speed(), n)));
                                if self.stop_signals.manager_stop_signal.load(Ordering::SeqCst) {
                                    return;
                                }
                                i = end;
                            }
                        }
                    }
                }
            }
        }
    }

    fn play_swipe_wdl_zones(&mut self, profile: &Profile, mode: SwipeMode, clean_with_black: bool) {
        let mut colors = distinct_swipe_colors(profile.rgb_array());
        legion_rgb_driver::debug_log(&format!("EFFECT: swipe WDL mode={mode:?} lamps=4"));
        while !self.stop_signals.manager_stop_signal.load(Ordering::SeqCst) {
            match mode {
                SwipeMode::Change => {
                    match profile.direction {
                        Direction::Left => colors.rotate_right(3),
                        Direction::Right => colors.rotate_left(3),
                    }
                    let _ = self.keyboard.set_colors_to(&colors);
                    thread::sleep(Duration::from_millis(wdl_delay_ms(self.live_speed())));
                }
                SwipeMode::Fill => {
                    let order: [usize; 4] = match profile.direction {
                        Direction::Left => [0, 1, 2, 3],
                        Direction::Right => [3, 2, 1, 0],
                    };
                    for src in order {
                        let mut frame = if clean_with_black { [0u8; 12] } else { colors };
                        for dst in order {
                            frame[dst * 3] = colors[src * 3];
                            frame[dst * 3 + 1] = colors[src * 3 + 1];
                            frame[dst * 3 + 2] = colors[src * 3 + 2];
                            let _ = self.keyboard.set_colors_to(&frame);
                            thread::sleep(Duration::from_millis(wdl_delay_ms(self.live_speed())));
                            if self.stop_signals.manager_stop_signal.load(Ordering::SeqCst) {
                                return;
                            }
                        }
                    }
                }
            }
        }
    }

    fn play_disco_wdl(&mut self, _profile: &Profile, rng: &mut ThreadRng) {
        let n = self.lamp_n();
        let palette = [
            [255, 0, 0],
            [255, 255, 0],
            [0, 255, 0],
            [0, 255, 255],
            [0, 0, 255],
            [255, 0, 255],
        ];
        let mut lamps = vec![[0u8; 3]; n];
        legion_rgb_driver::debug_log("EFFECT: disco WDL");
        self.paint_lamps(&lamps);
        while !self.stop_signals.manager_stop_signal.load(Ordering::SeqCst) {
            for lamp in lamps.iter_mut() {
                *lamp = effects::lamps::scale_rgb(*lamp, 0.72);
            }
            let color = palette[rng.random_range(0..palette.len())];
            lamps[rng.random_range(0..n)] = color;
            if n > 4 {
                lamps[rng.random_range(0..n)] = palette[rng.random_range(0..palette.len())];
            }
            self.paint_lamps(&lamps);
            thread::sleep(Duration::from_millis(wdl_delay_ms(self.live_speed())));
        }
    }
}

fn wdl_delay_ms(speed: u8) -> u64 {
    let speed = speed.clamp(1, 10) as u64;
    (400 / speed).clamp(25, 400)
}

fn swipe_step_ms(_speed: u8, _n: usize) -> u64 {
    // Match the 30fps painter. Faster sleeps just skip strips (the "sudden drop").
    33
}

fn swipe_fill_ms(speed: u8, n: usize) -> u64 {
    swipe_step_ms(speed, n)
}

fn swipe_stride(speed: u8, n: usize) -> usize {
    if n <= 4 {
        return 1;
    }
    match speed.clamp(1, 10) {
        1..=3 => 1,
        4..=7 => 2,
        _ => 3,
    }
}

fn distinct_swipe_colors(rgb: [u8; 12]) -> [u8; 12] {
    let z0 = (rgb[0], rgb[1], rgb[2]);
    let uniform = (0..4).all(|z| (rgb[z * 3], rgb[z * 3 + 1], rgb[z * 3 + 2]) == z0);
    if !uniform && rgb.iter().any(|v| *v >= 24) {
        return rgb;
    }
    [255, 0, 0, 0, 255, 0, 0, 0, 255, 255, 255, 0]
}

/// HSV to RGB conversion. Hue: 0-360, Saturation/Value: 0.0-1.0.
fn hsv_to_rgb(h: f64, s: f64, v: f64) -> (u8, u8, u8) {
    let c = v * s;
    let x = c * (1.0 - ((h / 60.0) % 2.0 - 1.0).abs());
    let m = v - c;
    let (r, g, b) = match h as u32 {
        0..=59 => (c, x, 0.0),
        60..=119 => (x, c, 0.0),
        120..=179 => (0.0, c, x),
        180..=239 => (0.0, x, c),
        240..=299 => (x, 0.0, c),
        _ => (c, 0.0, x),
    };
    (
        ((r + m) * 255.0) as u8,
        ((g + m) * 255.0) as u8,
        ((b + m) * 255.0) as u8,
    )
}

impl Drop for EffectManager {
    fn drop(&mut self) {
        let _ = self.tx.send(Message::Exit);
    }
}

#[derive(Clone)]
pub struct StopSignals {
    pub manager_stop_signal: Arc<AtomicBool>,
    pub keyboard_stop_signal: Arc<AtomicBool>,
}

impl StopSignals {
    pub fn store_true(&self) {
        self.keyboard_stop_signal.store(true, Ordering::SeqCst);
        self.manager_stop_signal.store(true, Ordering::SeqCst);
    }
    pub fn store_false(&self) {
        self.keyboard_stop_signal.store(false, Ordering::SeqCst);
        self.manager_stop_signal.store(false, Ordering::SeqCst);
    }
}
