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
    enums::{Effects, TypeHeatPalette, TypeHeatParams},
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

pub fn play(manager: &mut Inner, start: TypeHeatParams, rgb0: [u8; 12]) {
    legion_rgb_driver::debug_log("EFFECT: type heat");
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
    let mut zone_pressed: [HashSet<Keycode>; 4] = [HashSet::new(), HashSet::new(), HashSet::new(), HashSet::new()];
    let mut heat = [0.0f32; 4];
    let mut last = Instant::now();

    while !manager.stop_signals.manager_stop_signal.load(Ordering::SeqCst) {
        if let Ok(g) = manager.scene_params.lock() {
            if let Effects::TypeHeat { params: live } = g.effect {
                params = live.normalized();
                rgb = g.rgb;
            }
        }

        for event in rx.try_iter() {
            match event {
                Event::KeyPress(key) => {
                    for (i, zone) in KEY_ZONES.iter().enumerate() {
                        if zone.contains(&key) && zone_pressed[i].insert(key) {
                            heat[i] = (heat[i] + params.heat).min(1.0);
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

        let dt = last.elapsed().as_secs_f32().clamp(0.008, 0.08);
        last = Instant::now();
        let decay = (-params.cool * dt).exp();
        for i in 0..4 {
            heat[i] *= decay;
            if !zone_pressed[i].is_empty() {
                heat[i] = (heat[i] + params.hold_boost * dt).min(1.0);
            }
            if heat[i] < 0.002 {
                heat[i] = 0.0;
            }
        }

        let n = manager.lamp_n();
        let lamps_now: Vec<[u8; 3]> = (0..n)
            .map(|i| {
                let z = (i * 4 / n.max(1)).min(3);
                let shown = (params.background + (1.0 - params.background) * heat[z]).clamp(0.0, 1.0);
                let t = lamps::pos(i, n);
                lamps::scale_rgb(heat_color(params.palette, &rgb, t, shown), shown)
            })
            .collect();
        manager.paint_lamps(&lamps_now);
        thread::sleep(Duration::from_millis(if n > 4 { 33 } else { 16 }));
    }

    kill_thread.store(true, Ordering::SeqCst);
}

fn heat_color(palette: TypeHeatPalette, rgb: &[u8; 12], t: f32, heat: f32) -> [u8; 3] {
    let heat = heat.clamp(0.0, 1.0);
    match palette {
        TypeHeatPalette::Heat => {
            let h = 232.0 - heat * 232.0;
            let s = (0.88 - heat * 0.62).clamp(0.12, 1.0);
            let v = 0.16 + heat * 0.84;
            scene::hsv(h, s, v)
        }
        TypeHeatPalette::Ice => {
            let h = 198.0 - heat * 22.0;
            let s = (0.58 - heat * 0.48).clamp(0.08, 1.0);
            let v = 0.12 + heat * 0.88;
            scene::hsv(h, s, v)
        }
        TypeHeatPalette::Custom => {
            let hot = scene::custom(rgb, t);
            let cool = lamps::scale_rgb(hot, 0.14);
            let white_mix = ((heat - 0.72) / 0.28).clamp(0.0, 1.0);
            let hot = lamps::mix_rgb(hot, [255, 248, 236], white_mix);
            lamps::mix_rgb(cool, hot, heat)
        }
    }
}
