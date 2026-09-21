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

use crate::{
    enums::{Effects, NexusPalette, NexusParams},
    manager::{
        effects::{lamps, zones::KEY_ZONES},
        Inner,
    },
};

use super::scene;

enum Event {
    KeyPress(Keycode),
    KeyRelease(Keycode),
}

struct Pulse {
    zone: usize,
    age: f32,
}

pub fn play(manager: &mut Inner, start: NexusParams, rgb0: [u8; 12]) {
    legion_rgb_driver::debug_log("EFFECT: nexus");
    let kill_thread = Arc::new(AtomicBool::new(false));
    let exit_thread = kill_thread.clone();
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

    let mut params = start.normalized();
    let mut rgb = rgb0;
    let mut seen: [HashSet<Keycode>; 4] = [HashSet::new(), HashSet::new(), HashSet::new(), HashSet::new()];
    let mut pulses: Vec<Pulse> = Vec::new();
    let mut last = Instant::now();

    while !manager.stop_signals.manager_stop_signal.load(Ordering::SeqCst) {
        if let Ok(g) = manager.scene_params.lock() {
            if let Effects::Nexus { params: live } = g.effect {
                params = live.normalized();
                rgb = g.rgb;
            }
        }

        for event in rx.try_iter() {
            match event {
                Event::KeyPress(key) => {
                    for (i, zone) in KEY_ZONES.iter().enumerate() {
                        if zone.contains(&key) && seen[i].insert(key) {
                            pulses.push(Pulse {
                                zone: i,
                                age: 0.0,
                            });
                            if pulses.len() > 16 {
                                pulses.remove(0);
                            }
                        }
                    }
                }
                Event::KeyRelease(key) => {
                    for (i, zone) in KEY_ZONES.iter().enumerate() {
                        if zone.contains(&key) {
                            seen[i].remove(&key);
                        }
                    }
                }
            }
        }

        let dt = last.elapsed().as_secs_f32().clamp(0.008, 0.08);
        last = Instant::now();
        let life = (0.28 + params.fade * 0.55).clamp(0.22, 1.4);
        for pulse in pulses.iter_mut() {
            pulse.age += dt;
        }
        pulses.retain(|p| p.age < life);

        let n = manager.lamp_n().max(1);
        let width = if n > 4 { (n / 4).max(1) } else { 1 };
        let mut energy = vec![params.background; n];
        for pulse in &pulses {
            let fade = (1.0 - pulse.age / life).clamp(0.0, 1.0).powf(1.1);
            let amt = (params.pulse * fade).clamp(0.0, 1.0);
            if n <= 4 {
                energy[pulse.zone.min(n - 1)] = energy[pulse.zone.min(n - 1)].max(amt);
                if pulse.zone > 0 {
                    energy[pulse.zone - 1] = energy[pulse.zone - 1].max(amt * params.cross * 0.7);
                }
                if pulse.zone + 1 < n {
                    energy[pulse.zone + 1] = energy[pulse.zone + 1].max(amt * params.cross * 0.7);
                }
                continue;
            }
            let start = pulse.zone.min(3) * width;
            for k in 0..width {
                let i = start + k;
                if i < n {
                    energy[i] = energy[i].max(amt);
                }
            }
            if start > 0 {
                energy[start - 1] = energy[start - 1].max(amt * params.cross);
            }
            let end = start + width;
            if end < n {
                energy[end] = energy[end].max(amt * params.cross);
            }
        }

        let lamps_now: Vec<[u8; 3]> = (0..n)
            .map(|i| {
                let t = lamps::pos(i, n);
                let e = energy[i].clamp(0.0, 1.0);
                lamps::scale_rgb(nexus_color(params.palette, rgb, t, e), e)
            })
            .collect();
        manager.paint_lamps(&lamps_now);
        thread::sleep(Duration::from_millis(if n > 4 { 33 } else { 16 }));
    }

    kill_thread.store(true, Ordering::SeqCst);
}

fn nexus_color(palette: NexusPalette, rgb: [u8; 12], t: f32, energy: f32) -> [u8; 3] {
    match palette {
        NexusPalette::Cyan => scene::hsv(188.0, 0.85 - energy * 0.25, 0.2 + energy * 0.8),
        NexusPalette::Heat => scene::hsv(18.0 + energy * 22.0, 0.92, 0.25 + energy * 0.75),
        NexusPalette::Ice => scene::hsv(200.0, 0.45, 0.22 + energy * 0.78),
        NexusPalette::Custom => scene::custom(&rgb, t),
    }
}
