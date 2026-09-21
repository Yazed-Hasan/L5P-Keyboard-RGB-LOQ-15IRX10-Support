use std::{
    collections::HashSet,
    sync::{
        atomic::{AtomicBool, Ordering},
        Arc,
    },
    thread,
    time::{Duration, Instant},
};

use device_query::{DeviceEvents, DeviceEventsHandler, Keycode};

use crate::manager::{
    effects::lamps,
    profile::Profile,
    {effects::zones::KEY_ZONES, Inner},
};

struct Pulse {
    origin: f32,
    age: f32,
}

pub fn play(manager: &mut Inner, p: &Profile) {
    legion_rgb_driver::debug_log("EFFECT: ripple (smooth wave)");
    let kill_thread = Arc::new(AtomicBool::new(false));
    let exit_thread = kill_thread.clone();

    enum Event {
        KeyPress(Keycode),
        KeyRelease(Keycode),
    }

    let (tx, rx) = crossbeam_channel::unbounded::<Event>();

    thread::spawn(move || {
        let event_handler = DeviceEventsHandler::new(Duration::from_millis(8)).unwrap_or(DeviceEventsHandler {});
        let tx_clone = tx.clone();

        let press_guard = event_handler.on_key_down(move |key| {
            let _ = tx_clone.send(Event::KeyPress(*key));
        });

        let release_guard = event_handler.on_key_up(move |key| {
            let _ = tx.send(Event::KeyRelease(*key));
        });

        loop {
            if exit_thread.load(Ordering::SeqCst) {
                drop(press_guard);
                drop(release_guard);
                break;
            }
            thread::sleep(Duration::from_millis(5));
        }
    });

    let n = manager.lamp_n();
    let mut zone_pressed: [HashSet<Keycode>; 4] = [HashSet::new(), HashSet::new(), HashSet::new(), HashSet::new()];
    let mut was_pressed = [false; 4];
    let mut brightness = vec![0.0f32; n];
    let mut pulses: Vec<Pulse> = Vec::new();
    let mut last_tick = Instant::now();

    while !manager.stop_signals.manager_stop_signal.load(Ordering::SeqCst) {
        for event in rx.try_iter() {
            match event {
                Event::KeyPress(key) => {
                    for (i, zone) in KEY_ZONES.iter().enumerate() {
                        if zone.contains(&key) {
                            zone_pressed[i].insert(key);
                        }
                    }
                }
                Event::KeyRelease(key) => {
                    for (i, zone) in KEY_ZONES.iter().enumerate() {
                        if zone.contains(&key) {
                            zone_pressed[i].remove(&key);
                        }
                    }
                }
            }
        }

        let now = Instant::now();
        let dt = now.saturating_duration_since(last_tick).as_secs_f32().clamp(0.008, 0.05);
        last_tick = now;

        let speed = manager.effect_speed.load(Ordering::Relaxed).clamp(1, 10) as f32;
        let spread = (2.4 + speed * 0.85) * n as f32 / 4.0;
        let lifetime = (1.15 - speed * 0.055).clamp(0.55, 1.15);
        let width = 0.85 * n as f32 / 4.0;

        for i in 0..4 {
            let pressed = !zone_pressed[i].is_empty();
            if pressed && !was_pressed[i] {
                pulses.push(Pulse {
                    origin: (i as f32 + 0.5) * n as f32 / 4.0,
                    age: 0.0,
                });
                if pulses.len() > 12 {
                    pulses.remove(0);
                }
            }
            was_pressed[i] = pressed;
        }

        for pulse in pulses.iter_mut() {
            pulse.age += dt;
        }
        pulses.retain(|pulse| pulse.age < lifetime * 1.35);

        let mut target = vec![0.0f32; n];
        for pulse in &pulses {
            let radius = pulse.age * spread;
            let fade = (1.0 - pulse.age / lifetime).clamp(0.0, 1.0).powf(1.15);
            let origin_glow = (-pulse.age / 0.16).exp();
            for z in 0..n {
                let dist = (z as f32 - pulse.origin).abs();
                let ring = (-((dist - radius) * (dist - radius)) / (2.0 * width * width)).exp();
                let center = (-dist * dist / (0.55 * n as f32 / 4.0)).exp() * origin_glow;
                target[z] = target[z].max((ring * fade + center * 0.7).clamp(0.0, 1.0));
            }
        }
        for i in 0..4 {
            if was_pressed[i] {
                let start = (i * n) / 4;
                let end = ((i + 1) * n) / 4;
                for z in start..end {
                    target[z] = target[z].max(0.62);
                }
            }
        }
        for i in 0..n {
            let tau = if target[i] > brightness[i] { 0.04 } else { 0.14 };
            let alpha = 1.0 - (-dt / tau).exp();
            brightness[i] += (target[i] - brightness[i]) * alpha;
            if brightness[i] < 0.008 {
                brightness[i] = 0.0;
            }
        }

        let mut rgb_array = p.rgb_array();
        if rgb_array.iter().all(|&c| c == 0) {
            rgb_array = [255, 40, 80, 255, 160, 40, 40, 220, 120, 80, 120, 255];
        }
        let lamps: Vec<[u8; 3]> = (0..n)
            .map(|i| {
                let amount = brightness[i].clamp(0.0, 1.0).powf(1.12);
                lamps::scale_rgb(lamps::sample_zones(&rgb_array, lamps::pos(i, n)), amount)
            })
            .collect();
        manager.paint_lamps(&lamps);
        thread::sleep(Duration::from_millis(if n > 4 { 33 } else { 16 }));
    }

    kill_thread.store(true, Ordering::SeqCst);
}
