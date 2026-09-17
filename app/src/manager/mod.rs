use crate::enums::{Direction, Effects, Message, SwipeMode};

use crossbeam_channel::{Receiver, Sender};
use effects::{ambient, audio, christmas, disco, fade, lightning, ripple, swipe, temperature};
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
    pub window_active: Arc<AtomicBool>,
    effect_speed: Arc<AtomicU8>,
    audio_params: Arc<Mutex<audio::AudioReactParams>>,
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
        let window_active = keyboard.window_active_handle();
        let effect_speed = Arc::new(AtomicU8::new(1));
        let audio_params = Arc::new(Mutex::new(audio::AudioReactParams::default()));

        let (tx, rx) = crossbeam_channel::unbounded::<Message>();

        let mut inner = Inner {
            keyboard,
            rx,
            stop_signals: stop_signals.clone(),
            last_profile: Profile::default(),
            is_dynamic_lighting,
            effect_speed: effect_speed.clone(),
            audio_params: audio_params.clone(),
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
            window_active,
            effect_speed,
            audio_params,
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

    pub fn set_live_speed(&self, speed: u8) {
        let speed = speed.clamp(1, 10);
        self.effect_speed.store(speed, Ordering::Relaxed);
        legion_rgb_driver::debug_log(&format!(
            "SPEED: live={} delay={}ms",
            speed,
            wdl_delay_ms(speed)
        ));
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
            *guard = audio::AudioReactParams::from_effect(profile.effect);
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

    fn apply_effect(&mut self, profile: &mut Profile, rng: &mut ThreadRng) {
        match profile.effect {
            Effects::Static => {
                self.keyboard.set_colors_to(&profile.rgb_array()).unwrap();
                self.keyboard.set_effect(BaseEffects::Static).unwrap();
                if self.is_dynamic_lighting {
                    self.wdl_maintain();
                }
            }
            Effects::Breath => {
                self.keyboard.set_colors_to(&profile.rgb_array()).unwrap();
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
        let base_colors = profile.rgb_array();
        let mut phase: f64 = 0.0;
        while !self.stop_signals.manager_stop_signal.load(Ordering::SeqCst) {
            let factor = phase.sin() * 0.5 + 0.5; // oscillates 0.0..1.0
            let mut rgb = [0u8; 12];
            for i in 0..12 {
                rgb[i] = (base_colors[i] as f64 * factor) as u8;
            }
            self.keyboard.set_colors_to(&rgb).unwrap();
            phase += self.live_speed() as f64 * 0.08;
            if phase > std::f64::consts::TAU {
                phase -= std::f64::consts::TAU;
            }
            thread::sleep(Duration::from_millis(50));
        }
    }

    /// Software Smooth (rainbow cycle) effect for Windows Dynamic Lighting devices.
    fn play_smooth_wdl(&mut self) {
        let mut hue: f64 = 0.0;
        while !self.stop_signals.manager_stop_signal.load(Ordering::SeqCst) {
            let (r, g, b) = hsv_to_rgb(hue, 1.0, 1.0);
            let rgb = [r, g, b, r, g, b, r, g, b, r, g, b];
            self.keyboard.set_colors_to(&rgb).unwrap();
            hue = (hue + self.live_speed() as f64 * 3.0) % 360.0;
            thread::sleep(Duration::from_millis(50));
        }
    }

    /// Software Wave effect for Windows Dynamic Lighting devices.
    fn play_wave_wdl(&mut self, direction: Direction) {
        let mut hue: f64 = 0.0;
        let dir_mul: f64 = match direction {
            Direction::Left => 1.0,
            Direction::Right => -1.0,
        };
        while !self.stop_signals.manager_stop_signal.load(Ordering::SeqCst) {
            let mut rgb = [0u8; 12];
            for z in 0..4 {
                let h = (hue + z as f64 * 90.0 * dir_mul).rem_euclid(360.0);
                let (r, g, b) = hsv_to_rgb(h, 1.0, 1.0);
                rgb[z * 3] = r;
                rgb[z * 3 + 1] = g;
                rgb[z * 3 + 2] = b;
            }
            self.keyboard.set_colors_to(&rgb).unwrap();
            hue = (hue + self.live_speed() as f64 * 4.0) % 360.0;
            thread::sleep(Duration::from_millis(50));
        }
    }

    /// Fast swipe for WDL. The HID path uses 150-step fades; that looks frozen here,
    /// and rotating four identical zone colors is invisible.
    fn play_swipe_wdl(&mut self, profile: &Profile, mode: SwipeMode, clean_with_black: bool) {
        let mut colors = distinct_swipe_colors(profile.rgb_array());
        legion_rgb_driver::debug_log(&format!("EFFECT: swipe WDL mode={mode:?}"));
        while !self.stop_signals.manager_stop_signal.load(Ordering::SeqCst) {
            match mode {
                SwipeMode::Change => {
                    match profile.direction {
                        Direction::Left => colors.rotate_right(3),
                        Direction::Right => colors.rotate_left(3),
                    }
                    self.keyboard.set_colors_to(&colors).unwrap();
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
                            self.keyboard.set_colors_to(&frame).unwrap();
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

    fn play_disco_wdl(&mut self, profile: &Profile, rng: &mut ThreadRng) {
        let palette = [
            [255, 0, 0],
            [255, 255, 0],
            [0, 255, 0],
            [0, 255, 255],
            [0, 0, 255],
            [255, 0, 255],
        ];
        let mut rgb = profile.rgb_array();
        legion_rgb_driver::debug_log("EFFECT: disco WDL");
        self.keyboard.set_colors_to(&rgb).unwrap();
        while !self.stop_signals.manager_stop_signal.load(Ordering::SeqCst) {
            let color = palette[rng.random_range(0..palette.len())];
            let zone = rng.random_range(0..4usize);
            rgb[zone * 3] = color[0];
            rgb[zone * 3 + 1] = color[1];
            rgb[zone * 3 + 2] = color[2];
            self.keyboard.set_colors_to(&rgb).unwrap();
            thread::sleep(Duration::from_millis(wdl_delay_ms(self.live_speed())));
        }
    }
}

fn wdl_delay_ms(speed: u8) -> u64 {
    let speed = speed.clamp(1, 10) as u64;
    // Slider 1 is slow (~400ms/step), 10 is fast (~25ms/step).
    (400 / speed).clamp(25, 400)
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
