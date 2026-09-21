use std::{
    sync::atomic::Ordering,
    thread,
    time::{Duration, Instant},
};

use crate::{
    enums::{Effects, PacificaPalette, PacificaParams},
    manager::{effects::lamps, Inner},
};

use super::scene;

pub fn play(manager: &mut Inner, start: PacificaParams, rgb0: [u8; 12]) {
    legion_rgb_driver::debug_log("EFFECT: pacifica");
    let mut params = start.normalized();
    let mut rgb = rgb0;
    let mut last = Instant::now();
    let mut time = 0.0f32;

    while !manager.stop_signals.manager_stop_signal.load(Ordering::SeqCst) {
        if let Ok(g) = manager.scene_params.lock() {
            if let Effects::Pacifica { params: live } = g.effect {
                params = live.normalized();
                rgb = g.rgb;
            }
        }

        let dt = last.elapsed().as_secs_f32().clamp(0.008, 0.08);
        last = Instant::now();
        time += dt * params.speed;
        let n = manager.lamp_n();
        let layers = (2.0 + params.depth * 2.0).round() as usize;
        let mut lamps_now = vec![[0u8; 3]; n];

        for i in 0..n {
            let t = lamps::pos(i, n);
            let mut e = 0.0f32;
            for layer in 0..layers {
                let speed = 0.35 + layer as f32 * 0.28;
                let scale = 1.8 + layer as f32 * 1.15;
                let phase = layer as f32 * 1.7;
                let s = ((t * scale + time * speed + phase).sin() * 0.5 + 0.5).powf(1.45);
                e += s / layers as f32;
            }
            let shown = (params.background + (1.0 - params.background) * e * params.intensity).clamp(0.0, 1.0);
            lamps_now[i] = lamps::scale_rgb(pacifica_color(params.palette, rgb, t, e, time), shown);
        }
        manager.paint_lamps(&lamps_now);
        thread::sleep(Duration::from_millis(33));
    }
}

fn pacifica_color(palette: PacificaPalette, rgb: [u8; 12], t: f32, energy: f32, time: f32) -> [u8; 3] {
    match palette {
        PacificaPalette::Ocean => {
            let h = 168.0 + energy * 38.0 + t * 18.0 + time.sin() * 6.0;
            let s = (0.72 - energy * 0.28).clamp(0.28, 1.0);
            scene::hsv(h, s, 0.35 + energy * 0.65)
        }
        PacificaPalette::Ice => {
            let h = 188.0 + energy * 22.0 - t * 8.0;
            let s = (0.48 - energy * 0.32).clamp(0.12, 1.0);
            scene::hsv(h, s, 0.28 + energy * 0.72)
        }
        PacificaPalette::Custom => {
            let base = scene::custom(&rgb, t);
            let foam = lamps::mix_rgb(base, [210, 240, 255], energy * 0.45);
            lamps::mix_rgb(lamps::scale_rgb(base, 0.35), foam, energy)
        }
    }
}
