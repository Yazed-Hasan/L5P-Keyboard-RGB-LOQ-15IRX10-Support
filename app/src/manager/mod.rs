use crate::enums::{Direction, Effects, Message};

use crossbeam_channel::{Receiver, Sender};
use effects::{ambient, christmas, disco, fade, lightning, ripple, swipe, temperature};
use error_stack::{Result, ResultExt};
use legion_rgb_driver::{BaseEffects, Keyboard, SPEED_RANGE};
use profile::Profile;
use rand::{rng, rngs::ThreadRng};
use single_instance::SingleInstance;
use std::{
    sync::atomic::{AtomicBool, Ordering},
    thread,
    time::Duration,
};
use std::{sync::Arc, thread::JoinHandle};
use thiserror::Error;

use self::custom_effect::{CustomEffect, EffectType};

pub mod custom_effect;
mod effects;
pub mod profile;

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

        let (tx, rx) = crossbeam_channel::unbounded::<Message>();

        let mut inner = Inner {
            keyboard,
            rx,
            stop_signals: stop_signals.clone(),
            last_profile: Profile::default(),
            is_dynamic_lighting,
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
                            thread::sleep(Duration::from_millis(20));
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
            inner_handle: Some(inner_handle),
            stop_signals,
        };

        Ok(manager)
    }

    pub fn set_profile(&mut self, profile: Profile) {
        self.stop_signals.store_true();
        self.tx.try_send(Message::Profile { profile }).unwrap();
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
                    self.play_smooth_wdl(profile.speed);
                } else {
                    self.keyboard.set_effect(BaseEffects::Smooth).unwrap();
                }
            }
            Effects::Wave => {
                if self.is_dynamic_lighting {
                    self.play_wave_wdl(profile.direction, profile.speed);
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
                swipe::play(self, profile, mode, clean_with_black);
            }
            Effects::Swipe { mode, clean_with_black } => swipe::play(self, profile, mode, clean_with_black),
            Effects::Disco => disco::play(self, profile, rng),
            Effects::Christmas => christmas::play(self, rng),
            Effects::Fade => fade::play(self, profile),
            Effects::Temperature => temperature::play(self),
            Effects::Ripple => ripple::play(self, profile),
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
        let speed_factor = profile.speed as f64 * 0.05;
        while !self.stop_signals.manager_stop_signal.load(Ordering::SeqCst) {
            let factor = phase.sin() * 0.5 + 0.5; // oscillates 0.0..1.0
            let mut rgb = [0u8; 12];
            for i in 0..12 {
                rgb[i] = (base_colors[i] as f64 * factor) as u8;
            }
            self.keyboard.set_colors_to(&rgb).unwrap();
            phase += speed_factor;
            if phase > std::f64::consts::TAU {
                phase -= std::f64::consts::TAU;
            }
            thread::sleep(Duration::from_millis(50));
        }
    }

    /// Software Smooth (rainbow cycle) effect for Windows Dynamic Lighting devices.
    fn play_smooth_wdl(&mut self, speed: u8) {
        let mut hue: f64 = 0.0;
        let speed_factor = speed as f64 * 2.0;
        while !self.stop_signals.manager_stop_signal.load(Ordering::SeqCst) {
            let (r, g, b) = hsv_to_rgb(hue, 1.0, 1.0);
            let rgb = [r, g, b, r, g, b, r, g, b, r, g, b];
            self.keyboard.set_colors_to(&rgb).unwrap();
            hue = (hue + speed_factor) % 360.0;
            thread::sleep(Duration::from_millis(50));
        }
    }

    /// Software Wave effect for Windows Dynamic Lighting devices.
    fn play_wave_wdl(&mut self, direction: Direction, speed: u8) {
        let mut hue: f64 = 0.0;
        let speed_factor = speed as f64 * 3.0;
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
            hue = (hue + speed_factor) % 360.0;
            thread::sleep(Duration::from_millis(50));
        }
    }
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
